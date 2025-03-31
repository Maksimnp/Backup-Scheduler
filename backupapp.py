import tkinter as tk
from tkinter import messagebox, filedialog
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.widgets import DateEntry
from ttkbootstrap.tooltip import ToolTip
import subprocess
import datetime
import os
import xml.etree.ElementTree as ET
import re
import logging
import ctypes
import threading
import queue

# Настройка логирования
logging.basicConfig(filename='backup.log', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Глобальные переменные для полей GUI
entry_source = None
entry_dest = None
var_overwrite = None
entry_time = None
var_frequency = None
task_selector = None
root = None
status_label = None
date_entry = None
time_selector = None

# Очередь для передачи результатов из потока
result_queue = queue.Queue()

# Функция для проверки прав администратора
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

# Функция для очистки имени папки от временных штампов
def clean_destination_folder(destination):
    destination = destination.replace('/', '\\')
    return destination

# Функция для создания команды robocopy
def create_robocopy_command(source, destination, overwrite):
    source = source.replace('/', '\\')
    destination = destination.replace('/', '\\')
    destination = clean_destination_folder(destination)
    
    # Базовые параметры: копировать подкаталоги, включая пустые, исключать старые файлы, копировать данные и атрибуты
    base_params = "/E /XO /COPY:DAT /R:5 /W:5 /MT:8"
    
    final_destination = destination
    if not overwrite:  # Если "Создавать с новым именем"
        # Создаём новую папку с временной меткой ВНУТРИ указанного пути
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        folder_name = f"backup_{timestamp}"
        final_destination = os.path.join(destination, folder_name).replace('/', '\\')
    
    command = f'robocopy "{source}" "{final_destination}" {base_params}'
    logging.info(f"Сформирована команда robocopy: {command}")
    return command, final_destination

# Функция для форматирования времени в HH:MM
def format_time(time_str):
    try:
        time_obj = datetime.datetime.strptime(time_str, "%H:%M")
        return time_obj.strftime("%H:%M")
    except ValueError:
        raise ValueError("Неверный формат времени (например, 02:00)!")

# Функция для преобразования даты из YYYY-MM-DD в DD/MM/YYYY
def convert_date_to_schtasks_format(date_str):
    try:
        date_obj = datetime.datetime.strptime(date_str, "%Y-%m-%d")
        return date_obj.strftime("%d/%m/%Y")
    except ValueError:
        raise ValueError("Неверный формат даты (ожидается YYYY-MM-DD)!")

# Функция для обновления даты и времени
def update_time_display():
    selected_date = date_entry.entry.get()  # Получаем дату в формате MM/DD/YY
    try:
        # Преобразуем дату в формат YYYY-MM-DD
        date_obj = datetime.datetime.strptime(selected_date, "%m/%d/%y")
        formatted_date = date_obj.strftime("%Y-%m-%d")
    except ValueError:
        formatted_date = datetime.datetime.now().strftime("%Y-%m-%d")
    time_str = time_selector.get()
    entry_time.delete(0, tk.END)
    entry_time.insert(0, f"{formatted_date} {time_str}")

# Асинхронная функция для выполнения команды
def run_command_async(cmd, success_message, error_message, callback):
    logging.info("Запуск команды (асинхронно): %s" % cmd)
    def execute_command():
        try:
            process = subprocess.Popen(cmd, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, encoding='cp866', errors='ignore')
            stdout, stderr = process.communicate(timeout=10)
            returncode = process.returncode
            if returncode == 0:
                logging.info("Команда выполнена успешно: %s\nВывод: %s" % (cmd, stdout))
                result_queue.put((True, success_message))
            else:
                logging.error("Ошибка команды: %s\nВывод: %s\nКод ошибки: %d" % (cmd, stderr, returncode))
                result_queue.put((False, "%s: %s\nКод ошибки: %d" % (error_message, stderr, returncode)))
        except subprocess.TimeoutExpired:
            logging.error("Тайм-аут команды: %s" % cmd)
            process.kill()
            result_queue.put((False, "%s: Команда не выполнена в течение 10 секунд" % error_message))
        except Exception as e:
            logging.error("Неизвестная ошибка при выполнении команды: %s\nОшибка: %s" % (cmd, str(e)))
            result_queue.put((False, "%s: Неизвестная ошибка: %s" % (error_message, str(e))))
    thread = threading.Thread(target=execute_command)
    thread.start()
    def check_result():
        try:
            success, message = result_queue.get_nowait()
            callback(success, message)
        except queue.Empty:
            root.after(10, check_result)
    root.after(10, check_result)

# Функция для проверки, создана ли задача этим ПО
def is_task_created_by_app(task_name):
    try:
        task_name_cleaned = task_name.strip("\\")
        result = subprocess.run('schtasks /query /tn "%s" /xml' % task_name_cleaned, shell=True, capture_output=True, text=True, encoding='cp866', errors='ignore', timeout=10)
        xml_data = result.stdout
        if not xml_data.strip():
            return False
        root = ET.fromstring(xml_data)
        command = root.find(".//{http://schemas.microsoft.com/windows/2004/02/mit/task}Command").text
        arguments = root.find(".//{http://schemas.microsoft.com/windows/2004/02/mit/task}Arguments").text
        full_command = "%s %s" % (command, arguments) if arguments else command
        return full_command.lower().startswith("cmd.exe /c robocopy")
    except Exception as e:
        logging.error("Ошибка проверки задачи '%s': %s" % (task_name, str(e)))
        return False

# Функция для получения списка существующих задач
def get_existing_tasks():
    logging.info("Получение списка задач")
    try:
        result = subprocess.run('schtasks /query /fo csv', shell=True, capture_output=True, text=True, encoding='cp866', errors='ignore', timeout=10)
        tasks = []
        lines = result.stdout.splitlines()
        for line in lines[1:]:
            if line.strip():
                task_name = line.split(',')[0].strip('"').lstrip("\\")
                if is_task_created_by_app(task_name):
                    tasks.append(task_name)
        logging.info("Список задач (созданных этим ПО): %s" % tasks)
        return tasks
    except subprocess.CalledProcessError as e:
        logging.error("Не удалось получить список задач: %s" % e.stderr)
        messagebox.showerror("Ошибка", "Не удалось получить список задач: %s" % e.stderr)
        return []
    except subprocess.TimeoutExpired:
        logging.error("Не удалось получить список задач: превышено время ожидания")
        messagebox.showerror("Ошибка", "Не удалось получить список задач: превышено время ожидания")
        return []

# Функция для извлечения информации о задаче
def get_task_info(task_name):
    logging.info("Извлечение информации о задаче: %s" % task_name)
    try:
        task_name_cleaned = task_name.strip("\\")
        result = subprocess.run('schtasks /query /tn "%s" /xml' % task_name_cleaned, shell=True, capture_output=True, text=True, encoding='cp866', errors='ignore', timeout=10)
        xml_data = result.stdout
        if not xml_data.strip():
            logging.warning("XML данные пусты для задачи: %s" % task_name)
            return None
        root = ET.fromstring(xml_data)
        logging.info("Полный XML задачи '%s':\n%s" % (task_name, xml_data))
        command = root.find(".//{http://schemas.microsoft.com/windows/2004/02/mit/task}Command").text
        arguments = root.find(".//{http://schemas.microsoft.com/windows/2004/02/mit/task}Arguments").text
        full_command = "%s %s" % (command, arguments) if arguments else command
        robocopy_match = re.search(r'robocopy\s+"([^"]+)"\s+"([^"]+)"\s+(.+)', full_command)
        if not robocopy_match:
            logging.warning("Не удалось извлечь параметры robocopy из команды: %s" % full_command)
            return None
        source = robocopy_match.group(1)
        destination = robocopy_match.group(2)
        # Проверяем, есть ли в пути назначения временная метка
        overwrite = "Перезаписывать"
        if re.search(r'\\backup_\d{8}_\d{6}$', destination):
            overwrite = "Создавать"
            # Убираем временную метку из пути назначения
            destination = os.path.dirname(destination)
        start_time = root.find(".//{http://schemas.microsoft.com/windows/2004/02/mit/task}StartBoundary").text
        date_time = start_time.split("T")
        date_part = date_time[0]
        time_part = date_time[1][:5]
        schedule_by_day = root.find(".//{http://schemas.microsoft.com/windows/2004/02/mit/task}ScheduleByDay")
        schedule_by_week = root.find(".//{http://schemas.microsoft.com/windows/2004/02/mit/task}ScheduleByWeek")
        if schedule_by_day is not None:
            frequency = "Ежедневно"
        elif schedule_by_week is not None:
            frequency = "Еженедельно"
        else:
            frequency = "Ежедневно"  # Дефолтное значение
            logging.warning("Не удалось определить тип расписания для задачи '%s'" % task_name)
        task_info = {
            "source": source,
            "destination": destination,
            "overwrite": overwrite,
            "time": f"{date_part} {time_part}",
            "frequency": frequency
        }
        logging.info("Извлечена информация о задаче '%s': %s" % (task_name, task_info))
        return task_info
    except ET.ParseError:
        logging.error("Ошибка парсинга XML для задачи '%s'" % task_name)
        return None
    except subprocess.CalledProcessError as e:
        logging.error("Не удалось получить информацию о задаче: %s" % e.stderr)
        return None
    except subprocess.TimeoutExpired:
        logging.error("Не удалось получить информацию о задаче: превышено время ожидания")
        return None

# Функция для создания задачи через schtasks
def create_task(task_name, command, start_datetime, frequency, callback):
    logging.info("Создание задачи: %s" % task_name)
    freq = "DAILY" if frequency == "Ежедневно" else "WEEKLY"
    date_str, time_str = start_datetime.split(" ")
    formatted_time = format_time(time_str)
    # Преобразуем дату в формат DD/MM/YYYY для schtasks
    formatted_date = convert_date_to_schtasks_format(date_str)
    escaped_command = command.replace('"', '\\"')
    task_name_cleaned = task_name.strip("\\")
    schtasks_cmd = f'schtasks /create /tn "{task_name_cleaned}" /tr "cmd.exe /c {escaped_command}" /sc {freq} /sd "{formatted_date}" /st {formatted_time} /f'
    logging.info("Формируемая команда: %s" % schtasks_cmd)
    run_command_async(schtasks_cmd, "Задача '%s' успешно создана!" % task_name, "Не удалось создать задачу", callback)

# Функция для удаления задачи
def delete_task(task_name, callback):
    logging.info("Удаление задачи: %s" % task_name)
    task_name_cleaned = task_name.strip("\\")
    schtasks_cmd = 'schtasks /delete /tn "%s" /f' % task_name_cleaned
    run_command_async(schtasks_cmd, "Задача '%s' удалена для пересоздания." % task_name, "Ошибка при удалении задачи '%s'" % task_name, callback)

# Функция для изменения задачи
def modify_task(task_name, command, start_datetime, frequency, callback):
    logging.info("Изменение задачи: %s" % task_name)
    delete_task(task_name, lambda success, message: 
        create_task(task_name, command, start_datetime, frequency, callback) if success else callback(False, message))

# Функция для выбора пути источника
def select_source_path():
    path = filedialog.askdirectory(title="Выберите папку источника")
    if path:
        entry_source.delete(0, tk.END)
        entry_source.insert(0, path)

# Функция для выбора пути назначения
def select_dest_path():
    path = filedialog.askdirectory(title="Выберите папку назначения")
    if path:
        entry_dest.delete(0, tk.END)
        entry_dest.insert(0, path)

# Функция обработки выбора задачи
def on_task_select(event):
    task_name = task_selector.get()
    if not task_name:
        return
    task_info = get_task_info(task_name)
    if task_info:
        logging.info("Загружена информация о задаче '%s': %s" % (task_name, task_info))
        entry_source.delete(0, tk.END)
        entry_source.insert(0, task_info["source"])
        entry_dest.delete(0, tk.END)
        entry_dest.insert(0, task_info["destination"])
        var_overwrite.set(task_info["overwrite"])
        entry_time.delete(0, tk.END)
        entry_time.insert(0, task_info["time"])
        time_selector.set(task_info["time"].split(" ")[1])
        var_frequency.set(task_info["frequency"])
        # Обновляем DateEntry
        date_obj = datetime.datetime.strptime(task_info["time"].split(" ")[0], "%Y-%m-%d")
        date_entry.entry.delete(0, tk.END)
        date_entry.entry.insert(0, date_obj.strftime("%m/%d/%y"))
    else:
        logging.warning("Не удалось загрузить информацию о задаче '%s'" % task_name)
        messagebox.showwarning("Предупреждение", "Не удалось загрузить информацию о задаче. Проверьте её параметры вручную.")

# Функция обработки кнопки "Создать задачу"
def create_backup_task():
    logging.info("Нажата кнопка 'Создать задачу'")
    if not is_admin():
        logging.warning("Программа запущена без прав администратора")
        messagebox.showerror("Ошибка", "Запустите программу с правами администратора!")
        return
    task_name = entry_name.get()
    source = entry_source.get()
    destination = entry_dest.get()
    start_datetime = entry_time.get()
    frequency = var_frequency.get()
    overwrite = var_overwrite.get() == "Перезаписывать"
    if not all([task_name, source, destination, start_datetime]):
        logging.warning("Не все поля заполнены")
        messagebox.showerror("Ошибка", "Заполните все поля!")
        return
    if not os.path.exists(source):
        logging.warning("Папка источника не существует: %s" % source)
        messagebox.showerror("Ошибка", "Папка источника '%s' не существует!" % source)
        return
    if not os.path.exists(destination):
        try:
            os.makedirs(destination)
        except Exception as e:
            logging.error("Не удалось создать папку назначения: %s" % e)
            messagebox.showerror("Ошибка", "Не удалось создать папку назначения: %s" % e)
            return
    try:
        date_str, time_str = start_datetime.split(" ")
        format_time(time_str)
    except ValueError as e:
        logging.warning("Неверный формат времени: %s" % start_datetime)
        messagebox.showerror("Ошибка", str(e))
        return
    robocopy_cmd, final_destination = create_robocopy_command(source, destination, overwrite)
    status_label.config(text="Создание задачи...")
    def callback(success, message):
        logging.info("Callback вызван: success=%s, message=%s" % (success, message))
        status_label.config(text="")
        if success:
            messagebox.showinfo("Успех", "%s\nПапка назначения: %s" % (message, final_destination))
            threading.Thread(target=update_task_list, daemon=True).start()
        else:
            messagebox.showerror("Ошибка", message)
    create_task(task_name, robocopy_cmd, start_datetime, frequency, callback)

# Функция обработки кнопки "Изменить задачу"
def modify_backup_task():
    logging.info("Нажата кнопка 'Изменить задачу'")
    if not is_admin():
        logging.warning("Программа запущена без прав администратора")
        messagebox.showerror("Ошибка", "Запустите программу с правами администратора!")
        return
    task_name = task_selector.get()
    source = entry_source.get()
    destination = entry_dest.get()
    start_datetime = entry_time.get()
    frequency = var_frequency.get()
    overwrite = var_overwrite.get() == "Перезаписывать"
    if not task_name:
        logging.warning("Задача для изменения не выбрана")
        messagebox.showerror("Ошибка", "Выберите задачу для изменения!")
        return
    try:
        task_name_cleaned = task_name.strip("\\")
        subprocess.run('schtasks /query /tn "%s"' % task_name_cleaned, shell=True, check=True, capture_output=True, text=True, encoding='cp866', errors='ignore')
    except subprocess.CalledProcessError as e:
        logging.error("Задача '%s' не найдена: %s" % (task_name, e.stderr))
        messagebox.showerror("Ошибка", "Задача '%s' не найдена. Создайте её заново." % task_name)
        return
    if not all([source, destination, start_datetime]):
        logging.warning("Не все поля заполнены")
        messagebox.showerror("Ошибка", "Заполните все поля!")
        return
    if not os.path.exists(source):
        logging.warning("Папка источника не существует: %s" % source)
        messagebox.showerror("Ошибка", "Папка источника '%s' не существует!" % source)
        return
    if not os.path.exists(destination):
        try:
            os.makedirs(destination)
        except Exception as e:
            logging.error("Не удалось создать папку назначения: %s" % e)
            messagebox.showerror("Ошибка", "Не удалось создать папку назначения: %s" % e)
            return
    try:
        date_str, time_str = start_datetime.split(" ")
        format_time(time_str)
    except ValueError as e:
        logging.warning("Неверный формат времени: %s" % start_datetime)
        messagebox.showerror("Ошибка", str(e))
        return
    robocopy_cmd, final_destination = create_robocopy_command(source, destination, overwrite)
    status_label.config(text="Изменение задачи...")
    def callback(success, message):
        logging.info("Callback вызван: success=%s, message=%s" % (success, message))
        status_label.config(text="")
        if success:
            messagebox.showinfo("Успех", "Задача '%s' успешно изменена!\nПапка назначения: %s" % (task_name, final_destination))
            threading.Thread(target=update_task_list, daemon=True).start()
        else:
            messagebox.showerror("Ошибка", message)
    modify_task(task_name, robocopy_cmd, start_datetime, frequency, callback)

# Функция для обновления списка задач
def update_task_list():
    logging.info("Обновление списка задач")
    tasks = get_existing_tasks()
    task_selector['values'] = tasks
    if tasks:
        if not task_selector.get():
            task_selector.set(tasks[0])
            on_task_select(None)
    else:
        task_selector.set("")
        entry_source.delete(0, tk.END)
        entry_dest.delete(0, tk.END)
        var_overwrite.set("Перезаписывать")
        entry_time.delete(0, tk.END)
        time_selector.set("00:00")
        date_entry.entry.delete(0, tk.END)
        date_entry.entry.insert(0, datetime.datetime.now().strftime("%m/%d/%y"))
        var_frequency.set("Ежедневно")

# Создание GUI
root = ttk.Window(themename="flatly")
root.title("Настройка резервного копирования")
root.geometry("450x750")

# Основной контейнер
main_frame = ttk.Frame(root, padding="20")
main_frame.pack(fill=tk.BOTH, expand=True)

# Заголовок
title_label = ttk.Label(main_frame, text="Резервное копирование", font=("Helvetica", 16, "bold"))
title_label.pack(pady=(0, 15))

# Выбор задачи
ttk.Label(main_frame, text="Выберите задачу:").pack(anchor="w", pady=(0, 5))
task_selector = ttk.Combobox(main_frame, bootstyle="primary")
task_selector.pack(fill=tk.X, pady=(0, 10))
task_selector.bind("<<ComboboxSelected>>", on_task_select)
ToolTip(task_selector, text="Выберите существующую задачу для просмотра или редактирования")

# Имя задачи
ttk.Label(main_frame, text="Имя задачи (для новой):").pack(anchor="w", pady=(0, 5))
entry_name = ttk.Entry(main_frame, bootstyle="primary")
entry_name.pack(fill=tk.X, pady=(0, 10))
ToolTip(entry_name, text="Введите имя для новой задачи резервного копирования")

# Источник
ttk.Label(main_frame, text="Источник:").pack(anchor="w", pady=(0, 5))
frame_source = ttk.Frame(main_frame)
frame_source.pack(fill=tk.X, pady=(0, 10))
entry_source = ttk.Entry(frame_source, bootstyle="primary")
entry_source.pack(side=tk.LEFT, expand=True, fill=tk.X)
source_button = ttk.Button(frame_source, text="Обзор", command=select_source_path, bootstyle="primary")
source_button.pack(side=tk.RIGHT, padx=5)
ToolTip(entry_source, text="Укажите папку, из которой будут копироваться файлы")
ToolTip(source_button, text="Открыть диалог для выбора исходной папки")

# Назначение
ttk.Label(main_frame, text="Назначение:").pack(anchor="w", pady=(0, 5))
frame_dest = ttk.Frame(main_frame)
frame_dest.pack(fill=tk.X, pady=(0, 10))
entry_dest = ttk.Entry(frame_dest, bootstyle="primary")
entry_dest.pack(side=tk.LEFT, expand=True, fill=tk.X)
dest_button = ttk.Button(frame_dest, text="Обзор", command=select_dest_path, bootstyle="primary")
dest_button.pack(side=tk.RIGHT, padx=5)
ToolTip(entry_dest, text="Укажите папку, куда будут копироваться файлы")
ToolTip(dest_button, text="Открыть диалог для выбора папки назначения")

# Действие при существующем файле
ttk.Label(main_frame, text="Действие при существующем файле:").pack(anchor="w", pady=(0, 5))
var_overwrite = tk.StringVar(value="Перезаписывать")
overwrite_radio = ttk.Radiobutton(main_frame, text="Перезаписывать", variable=var_overwrite, value="Перезаписывать", bootstyle="primary")
overwrite_radio.pack(anchor="w")
create_new_radio = ttk.Radiobutton(main_frame, text="Создавать с новым именем", variable=var_overwrite, value="Создавать", bootstyle="primary")
create_new_radio.pack(anchor="w", pady=(0, 10))
ToolTip(overwrite_radio, text="Перезаписывать существующие файлы с совпадающими именами")
ToolTip(create_new_radio, text="Создавать новую папку с временной меткой внутри указанной папки назначения")

# Дата и время
ttk.Label(main_frame, text="Дата начала:").pack(anchor="w", pady=(0, 5))
date_entry = DateEntry(main_frame, bootstyle="primary", firstweekday=0, dateformat="%m/%d/%y")
date_entry.pack(fill=tk.X, pady=(0, 10))
date_entry.bind("<<DateEntrySelected>>", lambda e: update_time_display())
ToolTip(date_entry, text="Выберите дату начала выполнения задачи (формат: MM/DD/YY)")

# Выбор времени
ttk.Label(main_frame, text="Время начала (HH:MM):").pack(anchor="w", pady=(0, 5))
time_selector = ttk.Combobox(main_frame, bootstyle="primary", values=[f"{h:02d}:{m:02d}" for h in range(24) for m in range(0, 60, 15)])
time_selector.set("00:00")
time_selector.pack(fill=tk.X, pady=(0, 10))
time_selector.bind("<<ComboboxSelected>>", lambda e: update_time_display())
ToolTip(time_selector, text="Выберите время начала задачи с интервалом 15 минут")

# Поле для отображения полной даты и времени
ttk.Label(main_frame, text="Полная дата и время:").pack(anchor="w", pady=(0, 5))
entry_time = ttk.Entry(main_frame, bootstyle="primary")
entry_time.pack(fill=tk.X, pady=(0, 10))
entry_time.insert(0, f"{datetime.datetime.now().strftime('%Y-%m-%d')} 00:00")
ToolTip(entry_time, text="Отображает полную дату и время начала задачи (YYYY-MM-DD HH:MM)")

# Периодичность
ttk.Label(main_frame, text="Периодичность:").pack(anchor="w", pady=(0, 5))
var_frequency = tk.StringVar(value="Ежедневно")
daily_radio = ttk.Radiobutton(main_frame, text="Ежедневно", variable=var_frequency, value="Ежедневно", bootstyle="primary")
daily_radio.pack(anchor="w")
weekly_radio = ttk.Radiobutton(main_frame, text="Еженедельно", variable=var_frequency, value="Еженедельно", bootstyle="primary")
weekly_radio.pack(anchor="w", pady=(0, 10))
ToolTip(daily_radio, text="Задача будет выполняться каждый день")
ToolTip(weekly_radio, text="Задача будет выполняться раз в неделю")

# Индикатор выполнения
status_label = ttk.Label(main_frame, text="", font=("Helvetica", 12))
status_label.pack(pady=5)
ToolTip(status_label, text="Показывает статус выполнения текущей операции")

# Кнопки
button_frame = ttk.Frame(main_frame)
button_frame.pack(pady=15)
create_button = ttk.Button(button_frame, text="Создать задачу", command=create_backup_task, bootstyle="primary")
create_button.pack(side=tk.LEFT, padx=5)
modify_button = ttk.Button(button_frame, text="Изменить задачу", command=modify_backup_task, bootstyle="primary")
modify_button.pack(side=tk.LEFT, padx=5)
ToolTip(create_button, text="Создать новую задачу резервного копирования")
ToolTip(modify_button, text="Изменить параметры выбранной задачи")

# Инициализация списка задач
update_task_list()

root.mainloop()
