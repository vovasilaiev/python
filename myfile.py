import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from datetime import datetime, timedelta
import requests
from bs4 import BeautifulSoup
from xls2xlsx import XLS2XLSX
import os
import configparser

base_url = "https://www.muiv.ru"

# страница с расписанием
url = base_url + "/studentu/fakultet-it/raspisaniya/"

# Для скачивания файла
def download_file(download_url, file_name):
    response = requests.get(download_url)
    if response.status_code == 200:
        with open(file_name, 'wb') as f:
            f.write(response.content)
        print(f"Downloaded {file_name}")
        return True
    else:
        print(f"Failed to download {file_name}")
        return False

def save_config(teacher_name, file_path, last_update):
    config = configparser.ConfigParser()
    config['DEFAULT'] = {'TeacherName': teacher_name, 'FilePath': file_path, 'LastUpdate': last_update}
    with open('config.ini', 'w') as configfile:
        config.write(configfile)

def load_config():
    config = configparser.ConfigParser()
    config.read('config.ini')
    teacher_name = config['DEFAULT'].get('TeacherName', '')
    file_path = config['DEFAULT'].get('FilePath', '')
    last_update = config['DEFAULT'].get('LastUpdate', '')
    return teacher_name, file_path, last_update

def check_update_needed(last_update):
    if last_update:
        last_update_time = datetime.strptime(last_update, '%Y-%m-%d %H:%M:%S')
        if datetime.now() - last_update_time < timedelta(hours=4):
            return False
    return True

def update_files():
    response = requests.get(url)
    if response.status_code == 200:
        soup = BeautifulSoup(response.text, 'html.parser')
        
        sections = soup.find_all('h2')
        for section in sections:
            if "Бакалавриат" in section.text:
                download_items = section.find_next_sibling('div', class_='download').find_all('div', class_='download__item')
                for item in download_items:
                    link = item.find('a', class_='download__src')
                    if link and 'href' in link.attrs and (link['href'].endswith('.xlsx') or link['href'].endswith('.xls')):
                        file_url = base_url + link['href']
                        file_name = file_url.split('/')[-1]
                        if download_file(file_url, file_name):
                            if file_name.endswith('.xls'):
                                new_file_name = file_name[:-3] + 'xlsx'
                                x2x = XLS2XLSX(file_name)
                                x2x.to_xlsx(new_file_name)
                                os.remove(file_name)
                            save_config(default_teacher_name, default_file_path, datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
    else:
        print("Failed to retrieve the page")


# Функция поиска расписания
def find_schedule_by_teacher_name(teacher_name: str, file_path: str, sheet_names: list[str], today: str):
    """Поиск расписания для преподавателя по дате.

    Параметры
    ---------
    teacher_name: Имя преподавателя
    file_path: Путь к файлу
    sheet_names: Список листов для чтения
    today: Дата

    Возвращает
    ----------
    scheduleDay: Расписание на день
    """
    scheduleDay = {}
    for sheet in sheet_names:
        try:
            data = pd.read_excel(file_path, sheet_name=sheet, header=None)

        # Остальная часть обработки листа
        except ValueError:
            print(f"Лист '{sheet}' не найден в файле.")
            continue

        for index, row in data.iterrows():
            found = False

            for col_index, cell in enumerate(row):
                if teacher_name.lower() in str(cell).lower():
                    tName = cell
                    tName_col_index = col_index
                    found = True
                    break
            if not found:
                continue

            tName = data.iloc[index][tName_col_index]
            fIndex = 1
            while pd.isnull(data.iloc[index - fIndex][tName_col_index]):
                fIndex += 1
            subject = data.iloc[index - fIndex][tName_col_index]
            subject_indx = index - fIndex
                
            fIndex = 0
            while pd.isnull(data.iloc[index - fIndex][0]):
                fIndex += 1

            day = data.iloc[index - fIndex][0]
            date = data.iloc[index - fIndex][1]
            time = []
            for time_index in range(subject_indx, index):
                time_entry = data.iloc[time_index][2]
                if pd.notnull(time_entry):
                    time.append(time_entry)

            if pd.notnull(date) and date != 'nan':
                date = date.strftime('%Y-%m-%d')

            if date == today:
                schedule_entry = f"{day}, {date}: {subject} в {time}, {tName}"
                if sheet in scheduleDay:
                    scheduleDay[sheet].append(schedule_entry)
                else:
                    scheduleDay[sheet] = [schedule_entry]               
    print(scheduleDay)
    return scheduleDay

# Функция для вызова диалога выбора файла
def open_file_dialog():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    file_path_entry.delete(0, tk.END)
    file_path_entry.insert(0, file_path)

# Функция для обработки данных и вывода результатов
def process_and_show_schedule():
    teacher_name = teacher_name_entry.get()
    file_path = file_path_entry.get()
    if not file_path:
        messagebox.showerror("Ошибка", "Пожалуйста, выберите файл Excel")
        return
    if not teacher_name:
        messagebox.showerror("Ошибка", "Пожалуйста, введите имя преподавателя")
        return

    save_config(teacher_name, file_path, last_update)
    sheet_names = ['1 курс', '1 курс ', '2 курс', '2 курс ', '3 курс', '3 курс ']
    schedule = find_schedule_by_teacher_name(teacher_name, file_path, sheet_names, datetime.today().strftime('%Y-%m-%d'))
    output_text.delete('1.0', tk.END)    

    if schedule:
        for course, schedule_list in schedule.items():
            output_text.insert(tk.END, f"Расписание для {course}:\n")
            for schedule_item in schedule_list:
                output_text.insert(tk.END, schedule_item + "\n")
            output_text.insert(tk.END, "\n")
    else:
        output_text.insert(tk.END, "На сегодня пар нет\n")

default_teacher_name, default_file_path, last_update = load_config()

def get_start_of_week(today):
    # weekday() возвращает день недели в формате, где понедельник это 0, воскресенье - 6
    weekday = today.weekday()

    # Если сегодня воскресенье, начинаем с понедельника следующей недели
    if weekday == 6:
        start_of_week = today + timedelta(days=1)
    else:
        # Иначе, откатываемся назад к последнему понедельнику
        start_of_week = today - timedelta(days=weekday)

    return start_of_week

def show_weekly_schedule():
    teacher_name = teacher_name_entry.get()
    file_path = file_path_entry.get()
    if not file_path or not teacher_name:
        messagebox.showerror("Ошибка", "Пожалуйста, проверьте введенные данные")
        return

    # Получаем расписание на неделю начиная с сегодняшнего дня
    schedule_week = {}
    today = datetime.today()
    start_of_week = get_start_of_week(today)
    sheet_names = ['1 курс', '1 курс ', '2 курс', '2 курс ', '3 курс', '3 курс ']
    for i in range(7):  # для каждого дня в неделе
        day = start_of_week + timedelta(days=i)
        schedule_day = find_schedule_by_teacher_name(teacher_name, file_path, sheet_names, day.strftime('%Y-%m-%d'))
        if schedule_day:
            schedule_week[day.strftime('%Y-%m-%d')] = schedule_day
    
    # Отображение расписания
    output_text.delete('1.0', tk.END)
    if schedule_week:
        for date, schedules in schedule_week.items():
            output_text.insert(tk.END, f"Расписание на {date}:\n")
            for course, schedule_list in schedules.items():
                output_text.insert(tk.END, f"Курс {course}:\n")
                for schedule_item in schedule_list:
                    output_text.insert(tk.END, schedule_item + "\n")
                output_text.insert(tk.END, "\n")
    else:
        output_text.insert(tk.END, "На эту неделю пар нет\n")


if check_update_needed(last_update):
    update_files()
    config = configparser.ConfigParser()
    config.read('config.ini')
    last_update = config['DEFAULT'].get('LastUpdate', '')

root = tk.Tk()
root.title("Расписание преподавателей")

label = tk.Label(root, text="Введите фамилию преподавателя, формат Иванов И.И.:")
label.pack()

teacher_name_entry = tk.Entry(root)
teacher_name_entry.insert(0, default_teacher_name)
teacher_name_entry.pack()

file_path_button = tk.Button(root, text="Выберите файл Excel", command=open_file_dialog)
file_path_button.pack()

file_path_entry = tk.Entry(root)
file_path_entry.insert(0, default_file_path)
file_path_entry.pack()

process_button = tk.Button(root, text="Показать расписание на сегодня", command=process_and_show_schedule)
process_button.pack()
week_schedule_button = tk.Button(root, text="Показать расписание на неделю", command=show_weekly_schedule)
week_schedule_button.pack()

output_text = tk.Text(root, height=10, width=50)
output_text.pack()

root.mainloop()