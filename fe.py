import sys
from PyQt5.QtWidgets import QSizePolicy, QGridLayout, QDialog, QLineEdit, QApplication, QMainWindow, QLabel, QPushButton, QVBoxLayout, QWidget, QMessageBox, QHBoxLayout
from PyQt5.QtCore import Qt, QTimer, QDateTime, QThread, pyqtSignal, QDate
from PyQt5.QtGui import QFont, QPixmap
from ctypes import *
import os
import time
import sqlite3
import logging
import pendulum
from playsound import playsound
import threading
import keyboard

# Определяем путь к базе данных и библиотеке
DB_PATH = os.path.join("C:\\TimeTrackingSystem", "db", "employee_management.db")
DLL_PATH = os.path.join("C:\\TimeTrackingSystem", "libs", "ftrapi.dll")

FTR_RETCODE_OK                  = 0
#FTR_PARAM_IMAGE_WIDTH           = c_ulong(1)
#FTR_PARAM_IMAGE_HEIGHT          = c_ulong(2)
#FTR_PARAM_IMAGE_SIZE            = c_ulong(3)
FTR_PARAM_CB_FRAME_SOURCE       = c_ulong(4)
FTR_PARAM_CB_CONTROL            = c_ulong(5)
FTR_PARAM_MAX_TEMPLATE_SIZE     = c_ulong(6)
#FTR_PARAM_MAX_FAR_REQUESTED     = c_ulong(7)
#FTR_PARAM_SYS_ERROR_CODE        = c_ulong(8)
FTR_PARAM_FAKE_DETECT           = c_ulong(9)
#FTR_PARAM_MAX_MODELS            = c_ulong(10)
FTR_PARAM_FFD_CONTROL           = c_ulong(11)
FTR_PARAM_MIOT_CONTROL          = c_ulong(12)
FTR_PARAM_MAX_FARN_REQUESTED    = c_ulong(13)
FTR_PARAM_VERSION               = c_ulong(14)
FSD_FUTRONIC_USB    = c_void_p(1)
FTR_CB_RESP_CANCEL      = c_ulong(1)
FTR_CB_RESP_CONTINUE    = c_ulong(2)
FTR_PURPOSE_IDENTIFY = c_ulong(2)
#FTR_PURPOSE_ENROLL   = c_ulong(3)
#FTR_PURPOSE_COMPATIBILITY = c_ulong(4)

#FTR_STATE_FRAME_PROVIDED    = 0x01
FTR_STATE_SIGNAL_PROVIDED   = 0x02
FTR_VERSION_CURRENT = c_ulong(3)
# Определение типов
FTR_USER_CTX = c_void_p  # Аналог указателя на произвольные данные
FTR_STATE = c_uint32     # Маска состояния
FTR_SIGNAL = c_uint32    # Тип сигнала
FTR_BITMAP_PTR = c_void_p  # Указатель на bitmap-структуру
FTR_RESPONSE = c_ulong   # Аналог typedef UDGT32 FTR_RESPONSE

class FtrData(Structure):
    _pack_ = 1
    _fields_ = [('dwsize', c_ulong), ('pdata', POINTER(c_void_p))]

class FtrEnrollData(Structure):
        _pack_ = 1
        _fields_ = [('dwsize', c_ulong),
                    ('dwquality', c_ulong)]
        
class FtrIdentifyRecord(Structure):
    _pack_ = 1
    _fields_ = [('keyvalue', c_char * 16), ('pdata', POINTER(FtrData))]

class FtrIdentifyArray(Structure):
    _pack_ = 1
    _fields_ = [('TotalNumber', c_ulong), ('pmembers', POINTER(FtrIdentifyRecord))]

class FtrMatchedXRecord(Structure):
    _pack_ = 1
    _fields_ = [('keyvalue', c_char * 16)]

class FtrMatchedXArray(Structure):
    _pack_ = 1
    _fields_ = [('TotalNumber', c_ulong), ('pmembers', POINTER(FtrMatchedXRecord))]

# Логгер для service_log.txt
service_logger = logging.getLogger("service_logger")
service_logger.setLevel(logging.INFO)
service_handler = logging.FileHandler(os.path.join("C:\\TimeTrackingSystem", "logs", "service_log.txt"))
service_handler.setFormatter(logging.Formatter('%(asctime)s - %(message)s'))
service_logger.addHandler(service_handler)
# Глобальный флаг для управления отменой
cancel_response_flag = False

class EmployeeManager:
    def __init__(self):
        self.db_connection = self.connect_to_database()
    
    def connect_to_database(self):
        try:
            connection = sqlite3.connect(DB_PATH)
            return connection
        except sqlite3.Error as err:
            service_logger.error(f"Ошибка подключения к базе данных: {err}")
            raise

    # Универсальная функция для выполнения запросов
    def execute_query(self, query, params=(), fetch_one=False, fetch_all=False):
        try:
            with self.connect_to_database() as connection:
                connection.row_factory = sqlite3.Row  # Настройка для возврата словаря
                cursor = connection.cursor()
                cursor.execute(query, params)
                if fetch_one:
                    return cursor.fetchone()
                if fetch_all:
                    return cursor.fetchall()
                connection.commit()
        except sqlite3.Error as err:
            service_logger.error(f"[БД] Ошибка выполнения запроса: {err}")
            raise

    def get_all_employees(self):
        try:
            cursor = self.db_connection.cursor()
            cursor.execute("SELECT id, lastname, firstname, patronymic, dob, status, department, wages FROM employees ORDER BY CASE WHEN department = 'Руководство' THEN 1 WHEN department = 'Смена 1' THEN 2 WHEN department = 'Смена 2' THEN 3 WHEN department = 'Прокат' THEN 4 WHEN department = 'Резка' THEN 5 WHEN department = 'Склад' THEN 6 ELSE 7 END, lastname, firstname;")
            result = cursor.fetchall()
            return [{'id': emp_id, 'lastname': lastname, 'firstname': firstname, 'patronymic': patronymic, 'dob': dob,
                'status': status, 'department': department, 'wages' : wages} 
                    for emp_id, lastname, firstname, patronymic, dob, status, department, wages in result]
        except sqlite3.Error as err:
            service_logger.error(f"Ошибка при получении списка сотрудников: {err}")
            return []
        finally:
            cursor.close()

    def get_employee_by_id(self, emp_id):
        try:
            cursor = self.db_connection.cursor()
            cursor.execute(
                "SELECT id, lastname, firstname, patronymic, dob, hire_date, position, photo, status, department, wages FROM employees WHERE id = ?",
                (emp_id,)
            )
            result = cursor.fetchone()
            if result:
                return {
                    'id': result[0],
                    'lastname': result[1],
                    'firstname': result[2],
                    'patronymic': result[3],
                    'dob': result[4],
                    'hire_date': result[5],
                    'position': result[6],
                    'photo': result[7],
                    'status': result[8],
                    'department': result[9],
                    'wages': result[10]
                }
            return None
        except sqlite3.Error as err:
            service_logger.error(f"Ошибка получения данных о сотруднике: {err}")
            return []
        finally:
            cursor.close()

    def get_employee_photo(self, emp_id):
        try:
            cursor = self.db_connection.cursor()
            cursor.execute("SELECT photo FROM employees WHERE id = ?", (emp_id,))
            result = cursor.fetchone()
            return result[0] if result else None
        except sqlite3.Error as err:
            service_logger.error(f"Ошибка при получении фото сотрудника: {err}")
            return []
        finally:
            cursor.close()

    def get_employee_name(self, employee_id):
        """
        Получить имя и фамилию сотрудника по ID.
        """
        id = employee_id
        query = "SELECT lastname, firstname FROM employees WHERE id = ?"
        result = self.execute_query(query, (id,), fetch_one=True)
        if result:
            return f"{result['firstname']} {result['lastname']}"
        else:
            return "Неизвестный сотрудник"

    # Получить запись табеля
    def get_timesheet_entry(self, employee_id, date):
        query = """
            SELECT arrival_time, departure_time
            FROM timesheet
            WHERE employee_id = ? AND date = ?
        """
        return self.execute_query(query, (employee_id, date), fetch_one=True)

    # Обновить запись табеля
    def update_timesheet(self, employee_id, date, arrival_time=None, departure_time=None):
        query = """
            INSERT INTO timesheet (employee_id, date, arrival_time, departure_time)
            VALUES (?, ?, ?, ?)
            ON CONFLICT(employee_id, date) DO UPDATE SET
                arrival_time = COALESCE(EXCLUDED.arrival_time, arrival_time),
                departure_time = COALESCE(EXCLUDED.departure_time, departure_time)
        """
        self.execute_query(query, (employee_id, date, arrival_time, departure_time))
        service_logger.info(f"[БД] Обновлена или добавлена запись для сотрудника {employee_id} на дату {date}.")

    def get_leaves_for_employee(self, employee_id, year):
        try:
            cursor = self.db_connection.cursor()
            cursor.execute("""
                SELECT type, start_date, end_date
                FROM leaves
                WHERE employee_id = ? AND strftime('%Y', start_date) = ?
            """, (employee_id, str(year)))
            return cursor.fetchall()
        except sqlite3.Error as err:
            service_logger.error(f"Ошибка при получении данных отпусков: {err}")
            return []
        finally:
            cursor.close()

    # Получить шаблоны отпечатков пальцев из базы данных
    def get_templates_from_database(self):
        query = "SELECT employee_id, fingerprint_template FROM fingerprints"
        return self.execute_query(query, fetch_all=True)

class NumericKeyboard(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Экранная клавиатура")
        self.setWindowFlag(Qt.WindowStaysOnTopHint)
        # Поле для отображения пароля
        self.password_field = QLineEdit(self)
        self.password_field.setEchoMode(QLineEdit.Password)
        # Сетка для кнопок
        grid_layout = QGridLayout()
        # Создание цифровых кнопок
        positions = [(i, j) for i in range(3) for j in range(3)]
        numbers = [str(i + 1) for i in range(9)]
        for position, number in zip(positions, numbers):
            button = QPushButton(number)
            button.clicked.connect(lambda checked, num=number: self.add_digit(num))
            grid_layout.addWidget(button, *position)
        # Кнопка для нуля
        zero_button = QPushButton("0")
        zero_button.clicked.connect(lambda: self.add_digit("0"))
        grid_layout.addWidget(zero_button, 3, 1)
        # Кнопка подтверждения
        enter_button = QPushButton("Ввод")
        enter_button.clicked.connect(self.accept_password)
        grid_layout.addWidget(enter_button, 3, 2)
        # Основной макет
        layout = QVBoxLayout()
        layout.addWidget(self.password_field)
        layout.addLayout(grid_layout)
        self.setLayout(layout)

    def add_digit(self, digit):
        """Добавить цифру в поле ввода."""
        self.password_field.setText(self.password_field.text() + digit)

    def accept_password(self):
        """Подтверждение ввода пароля."""
        self.accept()

    def get_password(self):
        """Возвращает введённый пароль."""
        return self.password_field.text()

class ServiceInterface(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Мониторинг рабочего времени")
        self.setGeometry(0, 0, QApplication.primaryScreen().size().width(), QApplication.primaryScreen().size().height())
        self.setWindowFlags(Qt.Window)
        self.setWindowFlag(Qt.WindowStaysOnTopHint, True)
        self.setWindowFlag(Qt.FramelessWindowHint, True)
        self.message_queue = []  # Очередь для сообщений
        self.is_displaying = False  # Флаг для отслеживания, отображается ли сейчас сообщение
        # Основной виджет и макет
        screen_width = QApplication.primaryScreen().size().width()
        screen_height = QApplication.primaryScreen().size().height()
        main_layout = QVBoxLayout()
        # Верхняя панель с кнопками и текущей датой/временем
        top_layout = QHBoxLayout()
        self.employee_manager = EmployeeManager()


        # Кнопка с текущей датой и временем
        self.datetime_button = QPushButton(self)
        self.datetime_button.setMinimumHeight(int(screen_height * 0.1))
        self.datetime_button.setFont(QFont("Arial", int(screen_height * 0.05), QFont.Bold))
        self.datetime_button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.datetime_button.clicked.connect(self.request_password)
        
        # Кнопка "Заблокировать"
        self.lock_button = QPushButton("Заблокировать", self)
        self.lock_button.setMinimumHeight(int(screen_height * 0.1))
        self.lock_button.setFont(QFont("Arial", int(screen_height * 0.03)))
        self.lock_button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.lock_button.setVisible(False)
        self.lock_button.clicked.connect(self.lock_desktop)
        top_layout.addWidget(self.lock_button)
        top_layout.addStretch()  # Добавляет гибкое пространство
        top_layout.addWidget(self.datetime_button)
        top_layout.addStretch()
        top_layout.addWidget(self.lock_button)
        # Установить одинаковую ширину для боковых кнопок

        main_layout.addLayout(top_layout)
        # Новый виджет для отображения списка сотрудников
        self.employee_list_widget = QWidget(self)
        self.employee_list_layout = QGridLayout(self.employee_list_widget)
        self.employee_list_layout.setSpacing(10)
        main_layout.addWidget(self.employee_list_widget)
        # Текстовое поле для вывода сообщений заменено на QLabel
        self.output_text = QLabel(self)
        self.output_text.setAlignment(Qt.AlignCenter)
        self.output_text.setStyleSheet("font-size: 52px; font-weight: bold;")
        self.output_text.setWordWrap(True)
        main_layout.addWidget(self.output_text, stretch=1)
        container = QWidget()
        container.setLayout(main_layout)
        self.setCentralWidget(container)
        # Обновление времени и даты
        self.update_datetime()
        timer = QTimer(self)
        timer.timeout.connect(self.update_datetime)
        timer.start(1000)
        self.message_delay = 2  # Задержка между сообщениями в секундах
        self.last_message_time = 0  # Время последнего вывода сообщения
        # Таймер для сообщений
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.display_next_message)
        self.timer.start(500)
        # Поток для работы
        self.scanning_thread = ScanningThread()
        self.scanning_thread.log_signal.connect(self.enqueue_message)
        self.scanning_thread.start()
        self.update_timer = QTimer(self)
        self.update_timer.timeout.connect(self.load_employee_list)
        self.update_timer.start(1000)
        # Заполняем список сотрудников
        self.load_employee_list()
        self.block_keyboard()

    def load_employee_list(self):
        try:
            # Очищаем текущий список
            for i in reversed(range(self.employee_list_layout.count())):
                widget = self.employee_list_layout.itemAt(i).widget()
                if widget is not None:
                    widget.deleteLater()
            screen_width = QApplication.primaryScreen().size().width()
            screen_height = QApplication.primaryScreen().size().height()
            font_size = max(14, int(screen_height * 0.01))
            # Получаем список всех сотрудников, кроме уволенных
            active_employees = [emp for emp in self.employee_manager.get_all_employees() if emp['status'] != 'уволен' and emp['department'] != 'Руководство']
            # Добавляем сотрудников в сетку
            row = 0
            column = 0
            max_per_column = 6  # Максимум 10 сотрудников на колонку
            current_date = QDate.currentDate().toString("yyyy-MM-dd")
            # Перебираем всех сотрудников и добавляем в сетку
            for i, emp in enumerate(active_employees):
                if row >= max_per_column:
                    row = 0
                    column += 1  # Переход на новую колонку
                item_widget = QWidget()
                layout = QHBoxLayout()
                # Фото сотрудника
                photo_data = self.employee_manager.get_employee_by_id(emp['id']).get('photo')
                photo_label = QLabel()
                photo_label.setFixedSize(int(screen_width * 0.01), int(screen_width * 0.01))
                if photo_data:
                    pixmap = QPixmap()
                    if pixmap.loadFromData(photo_data):
                        photo_size = int(QApplication.primaryScreen().size().width() * 0.03)  # 5% от ширины экрана
                        photo_label.setPixmap(pixmap.scaled(photo_size, photo_size, Qt.KeepAspectRatio, Qt.SmoothTransformation))
                        photo_label.setFixedSize(photo_size, photo_size)
                        photo_label.setScaledContents(True)
                # Фамилия и имя сотрудника
                employee_label = QLabel(f"{emp['lastname']} {emp['firstname']}")
                employee_label.setFont(QFont("Arial", font_size, QFont.Bold))
                employee_label.setMinimumWidth(int(screen_width * 0.1))
                employee_label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
                # Проверка записей о приходе и уходе
                timesheet_entry = self.employee_manager.get_timesheet_entry(emp['id'], current_date)
                job_status =  QLabel()
                job_status.setFont(QFont("Arial", font_size, QFont.Bold))
                job_status.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
                leaves = self.employee_manager.get_leaves_for_employee(emp['id'], QDate.currentDate().year())
                is_on_leave = False
                for leave in leaves:
                    start_date = QDate.fromString(leave[1], "yyyy-MM-dd")
                    end_date = QDate.fromString(leave[2], "yyyy-MM-dd")
                    if start_date <= QDate.currentDate() <= end_date:
                        is_on_leave = True
                        if leave[0] == "Отпуск":
                            job_status.setText("В отпуске")
                            job_status.setStyleSheet("color: blue; font-weight: bold;")
                        elif leave[0] == "Больничный":
                            job_status.setText("На больничном")
                            job_status.setStyleSheet("color: orange; font-weight: bold;")
                        break
                # Если сотрудник не на больничном и не в отпуске, проверяем состояние по табелю
                if not is_on_leave:
                    if timesheet_entry:
                        if timesheet_entry['arrival_time'] and not timesheet_entry['departure_time']:
                            job_status.setText("На работе")
                            job_status.setStyleSheet("color: green; font-weight: bold;")
                        elif timesheet_entry['arrival_time'] and timesheet_entry['departure_time']:
                            job_status.setText("Отсутствует")
                            job_status.setStyleSheet("color: red; font-weight: bold;")
                    else:
                        job_status.setText("Отсутствует")
                        job_status.setStyleSheet("color: red; font-weight: bold;")
                
                # Добавляем фото и имя в layout
                layout.addWidget(photo_label)
                layout.addWidget(employee_label)
                layout.addWidget(job_status)
                item_widget.setLayout(layout)
                # Добавляем элемент в сетку
                self.employee_list_layout.addWidget(item_widget, row, column)
                row += 1  # Переход к следующей строке
        except Exception as e:
            service_logger.error(f"Ошибка: {e}")

    def update_datetime(self):
        current_datetime = QDateTime.currentDateTime().toString("dd.MM.yyyy HH:mm:ss")
        self.datetime_button.setText(current_datetime)

    def request_password(self):
        keyboard = NumericKeyboard(self)
        if keyboard.exec_() == QDialog.Accepted:
            password = keyboard.get_password()
            if self.verify_password(password):
                self.setWindowFlag(Qt.WindowStaysOnTopHint, False)
                self.setWindowFlag(Qt.FramelessWindowHint, False)
                self.showNormal()  # Свернуть окно с возможностью управления
                # Показать кнопки после разблокировки
                self.unblock_keyboard()
                self.lock_button.show()
            else:
                QMessageBox.warning(self, "Ошибка", "Неверный пароль!")
                service_logger.info("Попытка ввода пароля.")

    def verify_password(self, password):
        # Простой пример проверки пароля
        # Замените "admin123" на ваш реальный пароль
        return password == "5427720"

    def enqueue_message(self, message):
        """Добавление нового сообщения в очередь."""
        self.message_queue.append(message)

    def display_next_message(self):
        """Отображение следующего сообщения из очереди, если не отображается текущее."""
        if self.message_queue and not self.is_displaying:
            # Проверяем, сколько времени прошло с последнего сообщения
            current_time = time.time()
            if current_time - self.last_message_time < self.message_delay:
                return  # Если прошло меньше времени, чем задержка, не выводим сообщение
            message = self.message_queue.pop(0)  # Извлекаем первое сообщение из очереди
            self.is_displaying = True
            # Проверка на наличие "Ошибка:" в квадратных скобках
            if "[Ошибка:]" in message:
                formatted_message = f'<font color="red">{message}</font>'  # Окрашиваем все сообщение в красный
            # Проверка на наличие "Успех!" в квадратных скобках
            elif "[Успех!]" in message:
                formatted_message = f'<font color="green">{message}</font>'  # Окрашиваем все сообщение в зеленый
            else:
                # Разделяем сообщение на части по ключевым словам
                words = message.split(" ")
                formatted_message = ""
                for word in words:
                    if "Ошибка:" in word:
                        # Если слово содержит "Ошибка", окрашиваем его в красный
                        formatted_message += f'<font color="red">{word}</font> '
                    elif "Успех!" in word:
                        # Если слово содержит "Успех!", окрашиваем его в зеленый
                        formatted_message += f'<font color="green">{word}</font> '
                    else:
                        # Остальной текст остается без изменений
                        formatted_message += word + " "
                # Убираем лишний пробел в конце строки
                formatted_message = formatted_message.strip()
            self.output_text.setText(formatted_message)  # Отображение нового сообщения
            self.last_message_time = current_time  # Обновляем время последнего сообщения
            self.is_displaying = False

    def lock_desktop(self):
        """Перевод окна в заблокированное состояние."""
        self.setWindowFlag(Qt.WindowStaysOnTopHint, True)
        self.setWindowFlag(Qt.FramelessWindowHint, True)
        self.showFullScreen()
        self.update()  # Обновить состояние окна
        # Скрыть кнопки при блокировке
        self.lock_button.hide()
        self.block_keyboard()

    def block_keyboard(self):
        try:
            """Блокировка всех клавиш клавиатуры"""
            keys = [
                'a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n', 'o', 'p', 'q', 'r', 's', 't', 
                'u', 'v', 'w', 'x', 'y', 'z', '1', '2', '3', '4', '5', '6', '7', '8', '9', '0', 'enter', 'space', 'shift', 
                'ctrl', 'alt', 'tab', 'esc', 'backspace', 'delete', 'up', 'down', 'left', 'right', 'home', 'end', 'pageup', 'pagedown',
                'f1', 'f2', 'f3', 'f4', 'f5', 'f6', 'f7', 'f8', 'f9', 'f10', 'f11', 'f12',  # Блокируем функциональные клавиши
                'win', 'windows',  # Блокируем клавишу Windows
            ]
            for key in keys:
                keyboard.block_key(key)  # Блокируем каждую клавишу
            service_logger.info("Клавиатура заблокирована")
        except Exception as e:
            service_logger.error(f"Ошибка: {e}")
    def unblock_keyboard(self):
        """Разблокировка клавиатуры"""
        keys = [
                'a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n', 'o', 'p', 'q', 'r', 's', 't', 
                'u', 'v', 'w', 'x', 'y', 'z', '1', '2', '3', '4', '5', '6', '7', '8', '9', '0', 'enter', 'space', 'shift', 
                'ctrl', 'alt', 'tab', 'esc', 'backspace', 'delete', 'up', 'down', 'left', 'right', 'home', 'end', 'pageup', 'pagedown',
                'f1', 'f2', 'f3', 'f4', 'f5', 'f6', 'f7', 'f8', 'f9', 'f10', 'f11', 'f12',  # Блокируем функциональные клавиши
                'win', 'windows',  # Блокируем клавишу Windows
            ]
        for key in keys:
            keyboard.unblock_key(key)  # Разблокируем каждую клавишу
        service_logger.info("Клавиатура разблокирована")

    def closeEvent(self, event):
        """Обновление cancel_response_flag при закрытии окна и проверка сообщений."""
        global cancel_response_flag
        # Устанавливаем флаг отмены, чтобы процесс сканирования был прерван
        cancel_response_flag = True
        self.scanning_thread.quit()  # Завершаем поток корректно
        self.scanning_thread.wait()  # Ждем завершения потока
        event.accept()

class ScanningThread(QThread):
    log_signal = pyqtSignal(str)
    def __init__(self):
        super().__init__()
        self.employee_manager = EmployeeManager()
    def run(self):
        service_logger.info("Запуск фоновой задачи.")
        try:
            ftrdll = CDLL(DLL_PATH, mode=RTLD_GLOBAL)
            baseSample = FtrData()
            enrolSample = FtrData()

            def play_sound_async(sound_file, event):
                """Функция для воспроизведения звука в отдельном потоке."""
                time.sleep(0.5)  # Задержка в 100 миллисекунд перед воспроизведением
                playsound(sound_file)
                event.set()  # Сигнализируем, что воспроизведение завершено

            def play_sound_and_wait(sound_file):
                """Запуск воспроизведения звука в потоке и ожидание его завершения."""
                event = threading.Event()
                thread = threading.Thread(target=play_sound_async, args=(sound_file, event))
                thread.start()
                event.wait()  # Ждём, пока событие не будет установлено (пока не завершится звук)
                thread.join()  # Дожидаемся завершения потока

            def check_result(result, operation):
                """Проверка результата выполнения функции и логирование ошибок"""
                if result != FTR_RETCODE_OK:
                    error_message = get_error_message(result)
                    self.log_signal.emit(f"[Ошибка:] {error_message}")
                    service_logger.error(f"[Ошибка] {operation}: {error_message} (Код ошибки: {result})")
                    return False
                return True

            def get_error_message(error_code):
                """Возвращает сообщение об ошибке по коду"""
                error_messages = {
                    1: "Недостаточно памяти",
                    2: "Некорректный аргумент",
                    3: "Ресурс уже используется",
                    4: "Некорректное назначение",
                    5: "Внутренняя ошибка",
                    6: "Не удалось захватить данные",
                    7: "Отменено пользователем",
                    8: "Нет больше попыток",
                    10: "Несогласованное сканирование",
                    11: "Истёк пробный период",
                    201: "Источник кадра не установлен",
                    202: "Устройство не подключено",
                    203: "Сбой устройства",
                    204: "Пустой кадр",
                    205: "Фальшивый источник",
                    206: "Несовместимое оборудование",
                    207: "Несовместимая прошивка",
                    208: "Источник кадра изменён",
                    209: "Несовместимое ПО",
                }
                return error_messages.get(error_code, "Неизвестная ошибка")

            callback = CFUNCTYPE(None, 
                        FTR_USER_CTX,  # Ввод: контекст пользователя
                        FTR_STATE,     # Ввод: маска состояния
                        POINTER(FTR_RESPONSE),  # Вывод: управление функцией API
                        FTR_SIGNAL,    # Ввод: сигнал взаимодействия
                        FTR_BITMAP_PTR # Ввод: указатель на bitmap
                        )

            def control(context, state_mask, p_response, signal, p_bitmap):
                global cancel_response_flag
                # Проверяем флаг для отмены
                if cancel_response_flag:
                    p_response.contents.value = FTR_CB_RESP_CANCEL.value
                else:
                    p_response.contents.value = FTR_CB_RESP_CONTINUE.value
                    def signal_undefined():
                        self.log_signal.emit("Ошибка: Неопределенный сигнал...")
                    def signal_touch():
                        self.log_signal.emit("Приложите палец к сканеру.")
                    def signal_takeoff():
                        print("Уберите палец.") 
                    def signal_fakesource():
                        self.log_signal.emit("Ошибка: Обнаружен ложный источник сигнала.")
                    switch = {
                        0: signal_undefined,
                        1: signal_touch,
                        2: signal_takeoff,
                        3: signal_fakesource
                    }
                    if state_mask & FTR_STATE_SIGNAL_PROVIDED:
                        switch[signal]()
            callback_function = callback(control)

            # Инициализация устройства
            def initialize_device():
                try:
                        ftrdll.FTRInitialize()

                        # Если инициализация успешна, идём дальше
                        if not check_result(ftrdll.FTRSetParam(FTR_PARAM_CB_FRAME_SOURCE, FSD_FUTRONIC_USB), "Установка источника кадра"):
                            ftrdll.FTRTerminate()
                            time.sleep(2)
                            return False

                        if not check_result(ftrdll.FTRSetParam(FTR_PARAM_CB_CONTROL, callback_function), "Установка callback-функции"):
                            ftrdll.FTRTerminate()
                            time.sleep(2)
                            return False

                        if not check_result(ftrdll.FTRSetParam(FTR_PARAM_MAX_FARN_REQUESTED, 200), "Установка параметра FARN"):
                            ftrdll.FTRTerminate()
                            time.sleep(2)
                            return False

                        if not check_result(ftrdll.FTRSetParam(FTR_PARAM_FAKE_DETECT, False), "Установка детектора фальшивых отпечатков"):
                            ftrdll.FTRTerminate()
                            time.sleep(2)
                            return False

                        if not check_result(ftrdll.FTRSetParam(FTR_PARAM_FFD_CONTROL, False), "Отключение контроля FFD"):
                            ftrdll.FTRTerminate()
                            time.sleep(2)
                            return False

                        if not check_result(ftrdll.FTRSetParam(FTR_PARAM_MIOT_CONTROL, True), "Отключение контроля MIOT"):
                            ftrdll.FTRTerminate()
                            time.sleep(2)
                            return False

                        if not check_result(ftrdll.FTRSetParam(FTR_PARAM_VERSION, FTR_VERSION_CURRENT), "Установка версии"):
                            ftrdll.FTRTerminate()
                            time.sleep(2)
                            return False

                        if not check_result(ftrdll.FTRGetParam(FTR_PARAM_MAX_TEMPLATE_SIZE, byref(enrolSample, FtrData.dwsize.offset)), "Получение размера шаблона"):
                            ftrdll.FTRTerminate()
                            time.sleep(2)
                            return False
                        return True
                except Exception as e:
                    service_logger.error(f"[Устройство] Ошибка инициализации: {e}")

            # Получить статус прихода или ухода и время
            def get_status_and_times(current_time):
                if 4 <= current_time.hour < 14:
                    return "приход", current_time.strftime("%H:%M"), None
                elif 14 <= current_time.hour < 23:
                    return "уход", None, current_time.strftime("%H:%M")
                else:
                    return None, None, None

            # Обработка результата идентификации
            def handle_match(employee_id):
                employee_name = self.employee_manager.get_employee_name(employee_id)
                # Получаем текущее время в UTC
                current_time_utc = pendulum.now("UTC")
                # Добавляем 3 часа (или 180 минут) к времени в UTC
                current_time = current_time_utc.add(hours=3)
                rounded_time = current_time.start_of('minute').add(minutes=-current_time.minute % 10)
                current_date = current_time.date().to_date_string()
                status, arrival_time, departure_time = get_status_and_times(rounded_time)
                if not status:
                    self.log_signal.emit(f"Ошибка: Сотрудник {employee_name}: вне рабочего времени ({current_date}).")
                    play_sound_and_wait('media/sounds/registration attempt out of time.wav')
                    return
                timesheet_entry = self.employee_manager.get_timesheet_entry(employee_id, current_date)
                if timesheet_entry:
                    if status == "приход" and not timesheet_entry['arrival_time']:
                        # Записываем первый приход
                        self.employee_manager.update_timesheet(employee_id, current_date, arrival_time, departure_time)
                        self.log_signal.emit(f"Успех! Сотрудник {employee_name}: приход в {arrival_time}.")
                        play_sound_and_wait('media/sounds/arrival time registered.wav')
                        return
                    if status == "приход" and timesheet_entry['arrival_time']:
                        last_arrival = pendulum.parse(timesheet_entry['arrival_time'])
                        time_difference = rounded_time.diff(last_arrival).in_minutes()
                        service_logger.info(f"[Идентификация] время: {rounded_time} время регистрации: {last_arrival} разница времени:{time_difference}")
                        if time_difference < 60:
                            self.log_signal.emit(f"Ошибка: Сотрудник {employee_name}: время прихода уже записано.")
                            play_sound_and_wait('media/sounds/arrival has already been registered.wav')
                            return
                        elif time_difference >= 60:
                            service_logger.info(f"[Идентификация] Сотрудник {employee_name}: сканирование на приход засчитано как уход (прошло больше часа).")
                            status = "уход"
                            departure_time = arrival_time
                            arrival_time = None
                    if status == "уход":
                        # Проверяем наличие записи о приходе
                        if not timesheet_entry['arrival_time']:
                            service_logger.error(f"[Идентификация] Сотрудник {employee_name}: нет записи о приходе. Уход засчитываем как приход.")
                            status = "приход"
                            arrival_time = departure_time
                            departure_time = None
                        elif timesheet_entry['departure_time']:
                            last_departure = pendulum.parse(timesheet_entry['departure_time'])
                            time_difference = rounded_time.diff(last_departure).in_minutes()

                            if time_difference < 30:
                                self.log_signal.emit(f"Ошибка: Сотрудник {employee_name}: повторное сканирование.")
                                play_sound_and_wait('media/sounds/registration attempt again.wav')
                                return
                            else:
                                self.log_signal.emit(f"Успех! Сотрудник {employee_name}: обновление времени ухода на {departure_time}.")
                                play_sound_and_wait('media/sounds/departure time updated.wav')
                                self.employee_manager.update_timesheet(employee_id, current_date, arrival_time, departure_time)
                                return
                self.employee_manager.update_timesheet(employee_id, current_date, arrival_time, departure_time)
                self.log_signal.emit(f"Успех! Сотрудник {employee_name}: {status} в {arrival_time or departure_time}.")
                service_logger.info(f"Успех! Сотрудник {employee_name}: {status} в {arrival_time or departure_time}.")
                if status == "приход":
                    play_sound_and_wait('media/sounds/arrival time registered.wav')
                else:
                    play_sound_and_wait('media/sounds/departure time registered.wav')

            # Функция идентификации
            def identification():
                try:
                    templates = self.employee_manager.get_templates_from_database()
                    records = []
                    for employee_id, template_data in templates:
                        enrol_sample = FtrData()
                        enrol_sample.dwsize = len(template_data)
                        data_carray = (c_ubyte * len(template_data)).from_buffer_copy(template_data)
                        enrol_sample.pdata = cast(pointer(data_carray), POINTER(c_void_p))

                        record = FtrIdentifyRecord()
                        record.keyvalue = str(employee_id).encode('utf-8')
                        record.pdata = pointer(enrol_sample)
                        records.append(record)
                    rec_array = FtrIdentifyArray()
                    rec_array.TotalNumber = len(records)
                    rec_array.pmembers = cast((FtrIdentifyRecord * len(records))(*records), POINTER(FtrIdentifyRecord))
                    baseSample.pdata = (c_void_p * enrolSample.dwsize)()
                    result = ftrdll.FTREnroll(None, FTR_PURPOSE_IDENTIFY, byref(baseSample))
                    if result == 8:
                        self.log_signal.emit("[Ошибка:] Процесс сканирования отменён.")
                    elif result == FTR_RETCODE_OK:
                        ftrdll.FTRSetBaseTemplate(byref(baseSample))
                        match_array = FtrMatchedXArray()
                        match_array.TotalNumber = len(records)
                        match_records = (FtrMatchedXRecord * len(records))()
                        match_array.pmembers = cast(match_records, POINTER(FtrMatchedXRecord))
                        res_num = c_ulong(0)
                        ftrdll.FTRIdentifyN(byref(rec_array), byref(res_num), byref(match_array))
                        for i in range(res_num.value):
                            employee_id = int(match_array.pmembers[i].keyvalue.decode('utf-8'))
                            service_logger.info(f"[Идентификация] {self.employee_manager.get_employee_name(employee_id)}")
                            handle_match(employee_id)
                        if res_num.value == 0:
                            self.log_signal.emit("Ошибка: Совпадений отпечатка не найдено.")
                            play_sound_and_wait('media/sounds/fingerprint not recognized.wav')
                    else:
                        error_message = get_error_message(result)
                        self.log_signal.emit(f"[Ошибка:] {error_message}")
                    ftrdll.FTRTerminate()
                    #time.sleep(1)
                except Exception as e:
                    service_logger.error(f"[Идентификация] Ошибка: {e}")

            # Запуск процесса идентификации
            def start_identification():
                    try:
                        if initialize_device():
                            identification()
                    except Exception as e:
                        service_logger.error(f"[Идентификация] Ошибка: {e}")
                        #time.sleep(1)
            while True:
                if not cancel_response_flag:
                    start_identification()
                    time.sleep(1)
                else:
                    break
        except Exception as e:
            self.log_signal.emit(f"[Ошибка:] {str(e)}")

if __name__ == '__main__':
    try:
        app = QApplication(sys.argv)
        window = ServiceInterface()
        window.show()  # Или window.show() для обычного окна
        sys.exit(app.exec_())
    except Exception as e:
        service_logger.error(f"Ошибка: {e}")
