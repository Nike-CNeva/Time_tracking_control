import win32serviceutil
import win32service
import win32event
import servicemanager
import os
import time
import sqlite3
from ctypes import *
import logging
import pendulum
from playsound import playsound
from threading import Thread

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE_DIR)

# Логгер для service_start_log.txt
start_logger = logging.getLogger("start_logger")
start_logger.setLevel(logging.DEBUG)
start_handler = logging.FileHandler(os.path.join(BASE_DIR, "logs", "service_start_log.txt"))
start_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
start_logger.addHandler(start_handler)

# Логгер для service_log.txt
service_logger = logging.getLogger("service_logger")
service_logger.setLevel(logging.INFO)
service_handler = logging.FileHandler(os.path.join(BASE_DIR, "logs", "service_log.txt"))
service_handler.setFormatter(logging.Formatter('%(asctime)s - %(message)s'))
service_logger.addHandler(service_handler)

# Глобальный флаг для управления отменой
cancel_response_flag = False

class AppServerSvc(win32serviceutil.ServiceFramework):
    _svc_name_ = "BioTimeControl"
    _svc_display_name_ = "Bio Time Control"

    def __init__(self, args):
        start_logger.debug("Инициализация службы началась.")
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        self.ftrdll = None
        start_logger.debug("Конструктор службы завершен.")

    def SvcStop(self):
        global cancel_response_flag
        cancel_response_flag = True
        time.sleep(1)
        start_logger.debug("Служба остановлена.")
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.hWaitStop)

    def SvcDoRun(self):
        global cancel_response_flag
        start_logger.debug("Служба начала выполнение.")
        cancel_response_flag = False
        servicemanager.LogMsg(servicemanager.EVENTLOG_INFORMATION_TYPE,
                              servicemanager.PYS_SERVICE_STARTED,
                              (self._svc_name_, ''))
        thread = Thread(target=self.main)
        thread.daemon = True  # Чтобы поток завершался при завершении службы
        thread.start()
        win32event.WaitForSingleObject(self.hWaitStop, win32event.INFINITE)
        start_logger.debug("Служба завершает выполнение.")

    def main(self):
        service_logger.info("Запуск основной логики службы.")

        # Определяем путь к базе данных и библиотеке
        DB_PATH = os.path.join(BASE_DIR, "db", "employee_management.db")
        DLL_PATH = os.path.join(BASE_DIR, "libs", "ftrapi.dll")
        # Константы для работы с библиотекой ctypes
        FTR_PARAM_CB_FRAME_SOURCE       = c_ulong(4)
        FTR_PARAM_CB_CONTROL            = c_ulong(5)
        FTR_PARAM_MAX_TEMPLATE_SIZE     = c_ulong(6)
        FTR_PARAM_FAKE_DETECT           = c_ulong(9)
        FTR_PARAM_FFD_CONTROL           = c_ulong(11)
        FTR_PARAM_MIOT_CONTROL          = c_ulong(12)
        FTR_PARAM_MAX_FARN_REQUESTED    = c_ulong(13)
        FTR_PARAM_VERSION               = c_ulong(14)
        FSD_FUTRONIC_USB    = c_void_p(1)
        FTR_PURPOSE_IDENTIFY = c_ulong(2)
        FTR_VERSION_CURRENT = c_ulong(3)
        FTR_CB_RESP_CANCEL      = c_ulong(1)
        FTR_CB_RESP_CONTINUE    = c_ulong(2)

        # Определение типов
        FTR_USER_CTX = c_void_p  # Аналог указателя на произвольные данные
        FTR_STATE = c_uint32     # Маска состояния
        FTR_SIGNAL = c_uint32    # Тип сигнала
        FTR_BITMAP_PTR = c_void_p  # Указатель на bitmap-структуру
        FTR_RESPONSE = c_ulong   # Аналог typedef UDGT32 FTR_RESPONSE

        class FtrData(Structure):
            _pack_ = 1
            _fields_ = [('dwsize', c_ulong), ('pdata', POINTER(c_void_p))]

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

        ftrdll = CDLL(DLL_PATH, mode=RTLD_GLOBAL)
        baseSample = FtrData()
        enrolSample = FtrData()
        # Функция для подключения к базе данных
        def connect_to_database():
            try:
                connection = sqlite3.connect(DB_PATH)
                connection.row_factory = sqlite3.Row
                return connection
            except sqlite3.Error as err:
                service_logger.error(f"[БД] Ошибка подключения: {err}")
                raise

        # Универсальная функция для выполнения запросов
        def execute_query(query, params=(), fetch_one=False, fetch_all=False):
            try:
                with connect_to_database() as connection:
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

        # Получить запись табеля
        def get_timesheet_entry(employee_id, date):
            query = """
                SELECT arrival_time, departure_time
                FROM timesheet
                WHERE employee_id = ? AND date = ?
            """
            return execute_query(query, (employee_id, date), fetch_one=True)

        # Обновить запись табеля
        def update_timesheet(employee_id, date, arrival_time=None, departure_time=None):
            query = """
                INSERT INTO timesheet (employee_id, date, arrival_time, departure_time)
                VALUES (?, ?, ?, ?)
                ON CONFLICT(employee_id, date) DO UPDATE SET
                    arrival_time = COALESCE(EXCLUDED.arrival_time, arrival_time),
                    departure_time = COALESCE(EXCLUDED.departure_time, departure_time)
            """
            execute_query(query, (employee_id, date, arrival_time, departure_time))
            service_logger.info(f"[БД] Обновлена или добавлена запись для сотрудника {employee_id} на дату {date}.")

        # Получить шаблоны отпечатков пальцев из базы данных
        def get_templates_from_database():
            query = "SELECT employee_id, fingerprint_template FROM fingerprints"
            return execute_query(query, fetch_all=True)

        def get_employee_name(employee_id):
            """
            Получить имя и фамилию сотрудника по ID.
            """
            id = employee_id
            query = "SELECT lastname, firstname FROM employees WHERE id = ?"
            result = execute_query(query, (id,), fetch_one=True)
            if result:
                return f"{result['firstname']} {result['lastname']}"
            else:
                return "Неизвестный сотрудник"

        def check_result(result, operation):
            """Проверка результата выполнения функции и логирование ошибок"""
            if result != 0:
                error_message = get_error_message(result)
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

        callback_function = callback(control)

        # Инициализация устройства
        def initialize_device():
            while True:
                try:
                    # Проверка каждой установки параметра
                    if not check_result(ftrdll.FTRInitialize(), "Инициализация устройства"):
                        service_logger.info("[Устройство] Ошибка при инициализации устройства, перезапуск через 1 секунду...")
                        ftrdll.FTRTerminate()
                        time.sleep(5)
                        continue
                    if not check_result(ftrdll.FTRSetParam(FTR_PARAM_CB_FRAME_SOURCE, FSD_FUTRONIC_USB), "Установка источника кадра"):
                        service_logger.info("[Устройство] Ошибка при установке источника кадра, перезапуск через 1 секунду...")
                        ftrdll.FTRTerminate()
                        time.sleep(1)
                        continue
                    if not check_result(ftrdll.FTRSetParam(FTR_PARAM_CB_CONTROL, callback_function), "Установка callback-функции"):
                        service_logger.info("[Устройство] Ошибка при установке callback-функции, перезапуск через 1 секунду...")
                        ftrdll.FTRTerminate()
                        time.sleep(1)
                        continue
                    if not check_result(ftrdll.FTRSetParam(FTR_PARAM_MAX_FARN_REQUESTED, 245), "Установка параметра FARN"):
                        service_logger.info("[Устройство] Ошибка при установке параметра FARN, перезапуск через 1 секунду...")
                        ftrdll.FTRTerminate()
                        time.sleep(1)
                        continue
                    if not check_result(ftrdll.FTRSetParam(FTR_PARAM_FAKE_DETECT, True), "Установка детектора фальшивых отпечатков"):
                        service_logger.info("[Устройство] Ошибка при установке детектора фальшивых отпечатков, перезапуск через 1 секунду...")
                        ftrdll.FTRTerminate()
                        time.sleep(1)
                        continue
                    if not check_result(ftrdll.FTRSetParam(FTR_PARAM_FFD_CONTROL, False), "Отключение контроля FFD"):
                        service_logger.info("[Устройство] Ошибка при отключении контроля FFD, перезапуск через 1 секунду...")
                        ftrdll.FTRTerminate()
                        time.sleep(1)
                        continue
                    if not check_result(ftrdll.FTRSetParam(FTR_PARAM_MIOT_CONTROL, False), "Отключение контроля MIOT"):
                        service_logger.info("[Устройство] Ошибка при отключении контроля MIOT, перезапуск через 1 секунду...")
                        ftrdll.FTRTerminate()
                        time.sleep(1)
                        continue
                    if not check_result(ftrdll.FTRSetParam(FTR_PARAM_VERSION, FTR_VERSION_CURRENT), "Установка версии"):
                        service_logger.info("[Устройство] Ошибка при установке версии, перезапуск через 1 секунду...")
                        ftrdll.FTRTerminate()
                        time.sleep(1)
                        continue
                    if not check_result(ftrdll.FTRGetParam(FTR_PARAM_MAX_TEMPLATE_SIZE, byref(enrolSample, FtrData.dwsize.offset)), "Получение размера шаблона"):
                        service_logger.info("[Устройство] Ошибка при получении размера шаблона, перезапуск через 1 секунду...")
                        ftrdll.FTRTerminate()
                        time.sleep(1)
                        continue

                    service_logger.info("[Устройство] Инициализация устройства завершена успешно.")
                    break  # Завершаем работу метода при успешной инициализации

                except Exception as e:
                    service_logger.error(f"[Устройство] Ошибка инициализации: {e}")
                    time.sleep(1)  # Задержка в случае ошибки, чтобы перезапустить попытку
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
            employee_name = get_employee_name(employee_id)

            # Получаем текущее время в UTC
            current_time_utc = pendulum.now("UTC")
            # Добавляем 3 часа (или 180 минут) к времени в UTC
            current_time = current_time_utc.add(hours=3)
            rounded_time = current_time.start_of('minute').add(minutes=-current_time.minute % 10)
            current_date = current_time.date().to_date_string()

            status, arrival_time, departure_time = get_status_and_times(rounded_time)
            if not status:
                service_logger.info(f"[Идентификация] Сотрудник {employee_name}: вне рабочего времени ({current_date}).")
                playsound('media/sounds/registration attempt out of time.wav')
                return

            timesheet_entry = get_timesheet_entry(employee_id, current_date)

            if timesheet_entry:
                if status == "приход" and not timesheet_entry['arrival_time']:
                    # Записываем первый приход
                    update_timesheet(employee_id, current_date, arrival_time, departure_time)
                    service_logger.info(f"[Идентификация] Сотрудник {employee_name}: приход в {arrival_time}.")
                    playsound('media/sounds/arrival time registered.wav')
                    return
                
                if status == "приход" and timesheet_entry['arrival_time']:
                    last_arrival = pendulum.parse(timesheet_entry['arrival_time'])
                    time_difference = rounded_time.diff(last_arrival).in_minutes()
                    service_logger.info(f"[Идентификация] время: {rounded_time} время регистрации: {last_arrival} разница времени:{time_difference}")
                    if time_difference < 60:
                        service_logger.info(f"[Идентификация] Сотрудник {employee_name}: время прихода уже записано.")
                        playsound('media/sounds/arrival has already been registered.wav')
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
                            service_logger.info(f"[Идентификация] Сотрудник {employee_name}: повторное сканирование.")
                            playsound('media/sounds/registration attempt again.wav')
                            return
                        else:
                            service_logger.info(f"[Идентификация] Сотрудник {employee_name}: обновление времени ухода.")
                            playsound('media/sounds/departure time updated.wav')
                            update_timesheet(employee_id, current_date, arrival_time, departure_time)
                            return

            update_timesheet(employee_id, current_date, arrival_time, departure_time)
            service_logger.info(f"[Идентификация] Сотрудник {employee_name}: {status} в {arrival_time or departure_time}.")
            if status == "приход":
                playsound('media/sounds/arrival time registered.wav')
            else:
                playsound('media/sounds/departure time registered.wav')

        # Функция идентификации
        def identification():
            try:
                templates = get_templates_from_database()
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
                    service_logger.info("[Идентификация] Процесс сканирования отменён.")

                elif result == 0:
                    ftrdll.FTRSetBaseTemplate(byref(baseSample))
                    match_array = FtrMatchedXArray()
                    match_array.TotalNumber = len(records)
                    match_records = (FtrMatchedXRecord * len(records))()
                    match_array.pmembers = cast(match_records, POINTER(FtrMatchedXRecord))

                    res_num = c_ulong(0)
                    ftrdll.FTRIdentifyN(byref(rec_array), byref(res_num), byref(match_array))
                    for i in range(res_num.value):
                        employee_id = int(match_array.pmembers[i].keyvalue.decode('utf-8'))
                        service_logger.info(f"[Идентификация] {get_employee_name(employee_id)}")
                        handle_match(employee_id)

                    if res_num.value == 0:
                        service_logger.info("[Идентификация] Совпадений не найдено.")
                        playsound('media/sounds/fingerprint not recognized.wav')

                else:
                    error_message = get_error_message(result)
                    service_logger.error(f"[Идентификация] Ошибка: {error_message}")

                ftrdll.FTRTerminate()
                time.sleep(1)

            except Exception as e:
                service_logger.error(f"[Идентификация] Ошибка: {e}")

        # Запуск процесса идентификации
        def start_identification():
                try:
                    initialize_device()
                    identification()
                except Exception as e:
                    service_logger.error(f"[Идентификация] Ошибка: {e}")
                    time.sleep(1)
        while True:
            start_identification()
            time.sleep(1)

if __name__ == '__main__':
    start_logger.debug("Запуск службы через win32serviceutil.")
    win32serviceutil.HandleCommandLine(AppServerSvc)