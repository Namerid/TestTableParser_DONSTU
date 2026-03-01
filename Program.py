import openpyxl
import pandas as pd
import pathlib
import shutil
import sys
from pprint import pprint
import re
import json


WORK_FOLDER = "Work_folder"
DEPARTMENTS_FOLDER = f"{WORK_FOLDER}\\Departments"
POINTS_FOLDER = f"{WORK_FOLDER}\\Points"
OUTPUT_FOLDER = f"{WORK_FOLDER}\\Results"
AUTUMN_POINTS_NAME = "autumn_points.xlsx"
SPRING_POINTS_NAME = "spring_points.xlsx"
AUTUMN_POINTS_PATH = f"{POINTS_FOLDER}\\{AUTUMN_POINTS_NAME}"
SPRING_POINTS_PATH = f"{POINTS_FOLDER}\\{SPRING_POINTS_NAME}"

DEPARTMENT_COLUMNS = {
    "Номер": "A",
    "Unnamed: 1": "B",
    "Учебный план": "C",
    "Факультет группы": "D",
    "Блок": "E",
    "Дисциплина, вид учебной работы": "F",
    "Закреплённая кафедра": "G",
    "Курс/Семестр или Курс/Сессия": "H",
    "Группа": "I",
    "Количество студентов": "J",
    "Недель": "K",
    "Вид занятий": "L",
    "Часов (на поток, группу, студента)": "M",
    "Виды контроля": "N",
    "КСР": "O",
    "Индивидуальные занятия": "P",
    "Контрольные": "Q",
    "Оценка по рейтингу": "R",
    "Рефераты": "S",
    "Эссе": "T",
    "РГР": "U",
    "Контрольных работ (заоч)": "V",
    "Консультации (СПО)": "W",
    "Нагрузка, час": "X",
    "Unnamed: 24": "Y",
    "Unnamed: 25": "Z",
    "Unnamed: 26": "AA",
    "Преподаватель": "AB",
    "Unnamed: 28": "AC",
    "Unnamed: 29": "AD",
    "Unnamed: 30": "AE",
    "Номер потока": "AF",
    "Индикатор первой группы потока": "AG",
    "Ауд. рекомендуемая каф.": "AH",
    "Дополнительно часов": "AI",
    "Unnamed: 35": "AJ",
    "Фактически выполнено": "AK",
    "Unnamed: 37": "AL",
    "Время проведения занятий по графику": "AM",
    "Unnamed: 39": "AN",
    "Распределение нагрузки, час": "AO",
    "Unnamed: 41": "AP",
    "Unnamed: 42": "AQ",
    "Доля внебюджет": "AR",
    "Доля иностр." : "AS",
    "Всего/Договор/Иностр": "AT",
    "Уровень образования": "AU",
    "Форма обучения": "AV",
    "Часов на экзамены": "AW",
    "Сам. работа": "AX",
    "Электронные часы": "AY",
    "Нормирующий коэффициент": "AZ",
    "Примечание": "BA",
    "ЗЕТ": "BB",
    "Число вопросов": "BC",
    "Часов в неделю": "BD"
}
DEPARTMENT_GROUP_HEADER_NAME = "Группа"
DEPARTMENT_DISCIPLINE_HEADER_NAME = "Дисциплина, вид учебной работы"
DEPARTMENT_NUMBER_OF_STUDENTS_HEADER_NAME = "Количество студентов"
DEPARTMENT_TYPE_OF_ACTIVITY_HEADER_NAME = "Вид занятий"
DEPARTMENT_TEACHER_HEADER_NAME = "Преподаватель"
DEPARTMENT_COURSE_HEADER_NAME = "Курс/Семестр или Курс/Сессия"
DEPARTMENT_COLUMNS_READ_LIST = ["Учебный план", "Факультет группы", "Дисциплина, вид учебной работы", "Курс/Семестр или Курс/Сессия", "Группа", "Количество студентов", "Вид занятий", "Преподаватель"]

DEPARTMENT_HEADER_ROWS_SIZE = 5

POINTS_COLUMNS = {
    "группа": "A",
    "ФИО": "B",
    "дисциплина": "C",
    "балл": "D",
    "начало попытки": "E",
    "конец попытки": "F"
}
POINTS_GROUP_HEADER_NAME = "группа"
POINTS_DISCIPLINE_HEADER_NAME = "дисциплина"
POINTS_POINT_HEADER_NAME = "балл"
POINTS_SKIP_VALUES = [None, "", "безоценочно"]

POINTS_HEADER_ROWS_SIZE = 1

OUTPUT_COLUMNS = {
    "Учебный план" : "A",
    "Факультет группы" : "B",
    "Дисциплина, вид учебной работы" : "C",
    "Курс/Семестр или Курс/Сессия" : "D",
    "Группа" : "E",
    "Количество студентов" : "F",
    "Вид занятий" : "G",
    "Преподаватель" : "H",
    "Прошли тест" : "I",
    "Средний балл" : "J"
}

OUTPUT_PASSED_HEADER_NAME = "Прошли тест"
OUTPUT_AVERAGE_HEADER_NAME = "Средний балл"
OUTPUT_COLUMNS_WRITE_LIST = ["Учебный план", "Факультет группы", "Дисциплина, вид учебной работы", "Курс/Семестр или Курс/Сессия", "Группа", "Количество студентов", "Вид занятий", "Преподаватель"]
OUTPUT_HEADER_ROWS_SIZE = 1

PATTERN = r"(?:,\s*п/г\s*\d+)?(?:,\s*часть\s*\d+)?$"


def save_to_json(data, file_path: str | pathlib.Path, indent: int = 2):
    file_path = pathlib.Path(file_path)
    file_path.parent.mkdir(parents=True, exist_ok=True)
    
    with open(file_path, 'w', encoding='utf-8') as f:
        json.dump(
            data,
            f,
            ensure_ascii=False,   # Кириллица
            indent=indent,        # Красиво
            sort_keys=True        # По алфавиту
        )
    print(f"JSON сохранён: {file_path}")

def to_float(value : str | float | int):
    if isinstance(value, str):
        value = value.replace(",", ".", 1)
        if value.replace(".", "").isdigit():
            return float(value)
    elif isinstance(value, float):
        return value
    elif isinstance(value, int):
        return float(value)   

def to_int(value : str | int):
    if isinstance(value, str) and value.isdigit():
        return int(value)
    elif isinstance(value, int):
        return value  

def recreate_work_folder(output_folder_name):
    if pathlib.Path(output_folder_name).exists():
        shutil.rmtree(output_folder_name)

    output_folder = pathlib.Path(output_folder_name)
    output_folder.mkdir(exist_ok=True, parents=True)

    return output_folder

# def start_check(value, check_list):
#     len_value = 0
#     return_value = ""

#     for check_item in check_list:
#         if value.startswith(check_item) and len_value < len(check_item):
#             len_value = len(check_item)
#             return_value = check_item

#     if len_value > 0:
#         return return_value
    
#     return value

def remove_suffixes(text: str, pattern: str) -> str:
    if not text or not isinstance(text, str):
        return text

    # Регулярное выражение:
    #   (?:,\s*п/г\s*\d+)? — опционально: ", п/г 1"
    #   (?:,\s*часть\s*\d+)? — опционально: ", часть 2"
    #   $ — конец строки
    
    # Удаляем совпадение
    cleaned = re.sub(pattern, "", text)
    
    # Убираем лишние запятые и пробелы с конца
    cleaned = re.sub(r"[,\s]+$", "", cleaned)
    
    return cleaned.strip()

def end_check(text: str, pattern: str) -> bool:
    if not text or not isinstance(text, str):
        return False

    # Регулярное выражение:
    #   [,\s]*      — ноль или больше запятых/пробелов
    #   п/г         — буквально "п/г" (регистронезависимо)
    #   \s*         — пробелы после
    #   \d+         — одна или более цифр
    #   $           — конец строки
    
    return bool(re.search(pattern, text, re.IGNORECASE))    

def preparation_of_departments(departments_name, work_folder_name=WORK_FOLDER, output_folder_name=DEPARTMENTS_FOLDER):
    departments_path = pathlib.Path(departments_name.strip("& ").strip("'\""))
    
    if not departments_path.exists():
        print ("Ошибка: Папки/файла не существует")
        shutil.rmtree(work_folder_name)
        return 1
    

    if departments_path.is_dir():
        output_folder = recreate_work_folder(output_folder_name)

        xls_files = list(departments_path .glob("*.xls"))
        xlsx_files = list(departments_path .glob("*.xlsx"))

        if not xls_files and not xlsx_files:
            print("Ошибка: Файлы не найдены.")
            shutil.rmtree(work_folder_name)
            return 2

        if (len(xls_files) > 0):
            print(f"Найдено {len(xls_files)} файл(ов) .xls. Начинаем конвертацию...")
            for xls_file in xls_files:
                file = pd.read_excel(xls_file, engine='xlrd')
                print("Конвертация файла:", xls_file.name)
                try:
                    file.to_excel(f'{str(output_folder)}\\{xls_file.stem}.xlsx', index=False, engine='openpyxl')
                except Exception as e:
                    print("Ошибка при конвертации файла:", xls_file.name)
                    print("Подробности:", e)

        if (len(xlsx_files) > 0):
            print(f"\nНайдено {len(xlsx_files)} файл(ов) .xlsx. Начинаем копирование...")
            for xlsx_file in xlsx_files:
                print ("Копирование файла:", xlsx_file.name)
                try:
                    shutil.copy(xlsx_file, f'{str(output_folder)}\\{xlsx_file.name}')
                except Exception as e:
                    print("Ошибка при копировании файла:", xls_file.name)
                    print("Подробности:", e)
                  
    elif departments_path.is_file():
        output_folder = recreate_work_folder(output_folder_name)

        if departments_path.suffix.lower() == ".xls":
            file = pd.read_excel(departments_path, engine='xlrd')
            print("Конвертация файла: ", departments_path.name)

            try:
                file.to_excel(f'{str(output_folder)}\\{departments_path.stem}.xlsx', index=False, engine='openpyxl')
            except Exception as e:
                print("Ошибка при конвертации файла:", departments_path.name)
                print("Подробности:", e)
                shutil.rmtree(work_folder_name)
                return 5
        elif departments_path.suffix.lower() == ".xlsx":
            print ("Копирование файла:", departments_path.name)

            try:
                shutil.copy(departments_path, f'{str(output_folder)}\\{departments_path.name}')
            except Exception as e:
                print("Ошибка при копировании файла:", departments_path.name)
                print("Подробности:", e)
                shutil.rmtree(work_folder_name)
                return 5    
        else:
            print ("Ошибка: Неверный тип файла.")
            shutil.rmtree(work_folder_name)
            return 3
    
    return 0

def preparation_of_points(autumn_points_name, spring_points_name, work_folder_name=WORK_FOLDER, output_folder_name=POINTS_FOLDER):
    autumn_points_path = pathlib.Path(autumn_points_name.strip("& ").strip("'\""))
    spring_points_path = pathlib.Path(spring_points_name.strip("& ").strip("'\""))

    output_folder_path = pathlib.Path(output_folder_name)

    if not output_folder_path.parent.exists():
        print ("Oшибка: Рабочая папка не существует")
        shutil.rmtree(work_folder_name)
        return 4
        
    output_folder_path.mkdir(exist_ok=True)

    if not autumn_points_path.exists() or not spring_points_path.exists():
        print ("Ошибка: Файла не существует")
        shutil.rmtree(work_folder_name)
        return 1
    

    if autumn_points_path.suffix.lower() == ".xlsx" and spring_points_path.suffix.lower() == ".xlsx":
        print ("Копирование файла:", autumn_points_path.name)

        try:
            shutil.copy(autumn_points_path, f'{str(output_folder_path)}\\{AUTUMN_POINTS_NAME}')
        except Exception as e:
            print("Ошибка при копировании файла:", autumn_points_path.name)
            print("Подробности:", e)
            shutil.rmtree(work_folder_name)
            return 5  
        
        print ("Копирование файла:", spring_points_path.name)
        try:
            shutil.copy(spring_points_path, f'{str(output_folder_path)}\\{SPRING_POINTS_NAME}')
        except Exception as e:
            print("Ошибка при копировании файла:", autumn_points_path.name)
            print("Подробности:", e)
            shutil.rmtree(work_folder_name)
            return 5  

        
    else:
        print ("Ошибка: Неверный тип файла.")
        shutil.rmtree(work_folder_name)
        return 3
    
    return 0

def read_department_file(file_path, department_header_rows_size = DEPARTMENT_HEADER_ROWS_SIZE, department_columns = DEPARTMENT_COLUMNS, department_teacher_header_name = DEPARTMENT_TEACHER_HEADER_NAME, department_number_of_students_header_name = DEPARTMENT_NUMBER_OF_STUDENTS_HEADER_NAME, department_group_header_name = DEPARTMENT_GROUP_HEADER_NAME, department_discipline_header_name = DEPARTMENT_DISCIPLINE_HEADER_NAME, department_course_header_name = DEPARTMENT_COURSE_HEADER_NAME, department_type_of_activity_header_name = DEPARTMENT_TYPE_OF_ACTIVITY_HEADER_NAME, department_columns_read_list = DEPARTMENT_COLUMNS_READ_LIST, pattern = PATTERN):
    
    departments_dict = {}

    try:
        wb = openpyxl.load_workbook(file_path)
    except Exception as e:
        print("Ошибка при открытии файла:", file_path.name)
        print("Подробности:", e)
        return
    
    sheet = wb.active
    n = sheet.max_row

    for i in range(department_header_rows_size + 1, n):
        if sheet[department_columns[department_group_header_name]+str(i)].value in [None, ""] or sheet[department_columns[department_discipline_header_name]+str(i)].value in [None, ""]:
            continue
        group = str(sheet[department_columns[department_group_header_name]+str(i)].value)
        discipline = remove_suffixes(sheet[department_columns[department_discipline_header_name]+str(i)].value, pattern)
        #discipline = sheet[department_columns[department_discipline_header_name]+str(i)].value
        discipline_full_name = str(sheet[department_columns[department_discipline_header_name]+str(i)].value)
        teacher = str(sheet[department_columns[department_teacher_header_name]+str(i)].value)
        if teacher not in ["None", ""]:
            teacher = teacher[1:]
        add_check = False

        if  group  not in departments_dict:
            departments_dict[group] = []
        
        add_dict = {}
        for column in department_columns_read_list:
            if column == department_discipline_header_name:
                add_dict[column] = discipline
                add_dict["Полное название дисциплины"] = discipline_full_name
            elif column == department_teacher_header_name:
                add_dict[column] = teacher
            elif column == department_type_of_activity_header_name and sheet[department_columns[column]+str(i)].value == None:
                add_dict[column] = ""
            else:
                add_dict[column] = str(sheet[department_columns[column]+str(i)].value)
        
        for record in departments_dict[group]:
            if end_check(discipline_full_name, r"[,\s]*п/г\s*\d+$") == end_check(record["Полное название дисциплины"], r"[,\s]*п/г\s*\d+$") and end_check(discipline_full_name, r"[,\s]*п/г\s*\d+$") == True and add_dict[department_course_header_name] == record[department_course_header_name] and group == record[department_group_header_name] and add_dict[department_type_of_activity_header_name] in record[department_type_of_activity_header_name]:
                number_of_students = to_int(record[department_number_of_students_header_name])
                number_of_students += to_int(add_dict[department_number_of_students_header_name])
                record[department_number_of_students_header_name] = str(number_of_students)
                if teacher not in record[department_teacher_header_name]:
                    record[department_teacher_header_name] += f", {teacher}"
                add_check = True
                break
            elif end_check(discipline_full_name, r"часть(?:_к)?\s*\d+$") == end_check(record["Полное название дисциплины"], r"часть(?:_к)?\s*\d+$") and end_check(discipline_full_name, r"часть(?:_к)?\s*\d+$") == True and add_dict[department_course_header_name] == record[department_course_header_name] and group == record[department_group_header_name] and add_dict[department_type_of_activity_header_name] in record[department_type_of_activity_header_name]:
                if teacher not in record[department_teacher_header_name]:
                    record[department_teacher_header_name] += f", {teacher}"
                add_check = True
                break
            
        if not add_check:
            departments_dict[group].append(add_dict)

    for group in departments_dict:
        i = 0
        
        for i in range(len(departments_dict[group])):
            if i >= len(departments_dict[group]):
                break
            for j in range(len(departments_dict[group])):
                if j >= len(departments_dict[group]) or i >= len(departments_dict[group]):
                    break
                # if j >= len(departments_dict[group]):
                #     break
                if j != i:
                    if (departments_dict[group][i][department_discipline_header_name] == departments_dict[group][j][department_discipline_header_name] and 
                        departments_dict[group][i][department_course_header_name] == departments_dict[group][j][department_course_header_name] and 
                        departments_dict[group][i][department_group_header_name] == departments_dict[group][j][department_group_header_name]):
                        
                        if departments_dict[group][j][department_teacher_header_name] not in departments_dict[group][i][department_teacher_header_name]:
                            departments_dict[group][i][department_teacher_header_name] += f", {departments_dict[group][j][department_teacher_header_name]}"

                        if departments_dict[group][j][department_type_of_activity_header_name] not in departments_dict[group][i][department_type_of_activity_header_name]:
                            departments_dict[group][i][department_type_of_activity_header_name] += f", {departments_dict[group][j][department_type_of_activity_header_name]}"
                        
                        i_number_of_students = to_int(departments_dict[group][i][department_number_of_students_header_name])
                        j_number_of_students = to_int(departments_dict[group][j][department_number_of_students_header_name])

                        if i_number_of_students < j_number_of_students:
                            departments_dict[group][i][department_number_of_students_header_name] = str(j_number_of_students)


                        departments_dict[group].pop(j)
            

    return departments_dict

def read_point_file(file_path, point_header_rows_size = POINTS_HEADER_ROWS_SIZE, points_columns = POINTS_COLUMNS, points_group_header_name = POINTS_GROUP_HEADER_NAME, points_discipline_header_name = POINTS_DISCIPLINE_HEADER_NAME, points_point_header_name = POINTS_POINT_HEADER_NAME, points_skip_values = POINTS_SKIP_VALUES):
    points_dict = {}

    try:
        wb = openpyxl.load_workbook(file_path)
    except Exception as e:
        print("Ошибка при открытии файла:", file_path.name)
        print("Подробности:", e)
        return
    
    sheet = wb.active
    n = sheet.max_row

    for i in range(point_header_rows_size + 1, n):
        try:
            group = sheet[points_columns[points_group_header_name]+str(i)].value
            discipline = sheet[points_columns[points_discipline_header_name]+str(i)].value
            if (type(discipline)== str):
                discipline.strip()
            points = sheet[points_columns[points_point_header_name]+str(i)].value

            if not(group in points_skip_values or discipline in points_skip_values or points in points_skip_values):
                if  group not in points_dict :
                    points_dict[group] = {}
                
                if discipline not in points_dict[group] :
                    points_dict[group][discipline] = []

                points_dict[group][discipline].append(to_float(points))
        
        except Exception as e:
            print("Ошибка при чтении баллов из файла:", file_path.name)
            print("Подробности:", e)
            continue
        
    return points_dict

def processing (departments_folder = DEPARTMENTS_FOLDER, autumn_points_path = AUTUMN_POINTS_PATH, sprint_points_path = SPRING_POINTS_PATH, points_folder = POINTS_FOLDER, output_folder = OUTPUT_FOLDER, output_columns = OUTPUT_COLUMNS, output_header_rows_size = OUTPUT_HEADER_ROWS_SIZE, output_columns_write_list = OUTPUT_COLUMNS_WRITE_LIST, output_passed_header_name = OUTPUT_PASSED_HEADER_NAME, output_average_header_name = OUTPUT_AVERAGE_HEADER_NAME, department_course_header_name = DEPARTMENT_COURSE_HEADER_NAME, department_discipline_header_name = DEPARTMENT_DISCIPLINE_HEADER_NAME):
    print("\nОсновная обработка данных начата...")
    
    autumn_points_path = pathlib.Path(autumn_points_path)
    spring_points_path = pathlib.Path(sprint_points_path)
    departments_path = pathlib.Path(departments_folder)
    output_folder_path = pathlib.Path(output_folder)

    autumn_points = read_point_file(autumn_points_path)
    print("Чтение осенних баллов")
    spring_points = read_point_file(spring_points_path)   
    print("Чтение весенних баллов\n") 

    try:    
        if pathlib.Path(output_folder_path).exists():
            shutil.rmtree(output_folder_path)
        output_folder_path.mkdir(exist_ok=True)
    except Exception as e:
        print("Ошибка при создании выходной папки:", output_folder)
        print("Подробности:", e)
        return


    for department in departments_path.glob("*.xlsx"):
        print("Обработка файла кафедры:", department.name)

        row_index = output_header_rows_size + 1

        try:
            departments_dict = read_department_file(department)

            wb = openpyxl.Workbook()
            sheet = wb.active

            for header_name in output_columns:
                sheet[output_columns[header_name] + "1"] = header_name

            row_index = output_header_rows_size + 1
            
            discipline = ""
            for group in departments_dict:
                spring_points_group = spring_points.get(group)
                autumn_points_group = autumn_points.get(group)
                for record in departments_dict[group]:
                    for column in output_columns_write_list:
                        if column == department_discipline_header_name:
                            discipline = record[column]
                        sheet[output_columns[column] + str(row_index)] = record[column]
                    course = record[department_course_header_name]
                    
                    if len(course) > 0 and discipline != "":
                        if course[-1].isdigit() and course[-1] != 0:
                            course = int(course[-1])
                            if course % 2 == 0:
                                if spring_points_group != None:
                                    points = spring_points_group.get(discipline)
                                    if points != None:
                                        sheet[output_columns[output_passed_header_name] + str(row_index)] = str(len (points))
                                        sheet[output_columns[output_average_header_name] + str(row_index)] = f"{(sum (points) / len(points)):.2f}"
                            else:
                                if autumn_points_group != None:
                                    points = autumn_points_group.get(discipline)
                                    if points != None:
                                        sheet[output_columns[output_passed_header_name] + str(row_index)] = str(len (points))
                                        sheet[output_columns[output_average_header_name] + str(row_index)] = f"{(sum (points) / len(points)):.2f}"
                    row_index += 1
  
            bold_font = openpyxl.styles.Font(bold=True)
            thin = openpyxl.styles.Side(border_style="thin", color="000000")
            border = openpyxl.styles.Border(left=thin, right=thin, top=thin, bottom=thin) 

            for cell in sheet[1]:  
                cell.font = bold_font

            for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row,
                min_col=1, max_col=sheet.max_column):
                for cell in row:
                    cell.border = border   

            wb.save(f"{output_folder}\\{department.name}")    
        except Exception as e:
            print("Ошибка при обработке файла кафедры:", department.name)
            print("Подробности:", e)
            continue

    # save_to_json(departments_dict, f"{output_folder}\\json1.json") 
    
    return 0

def main():
    try:
     
        if preparation_of_departments(input("Введите путь к папке/файлу кафедр: ")) != 0:
            print ("\nОшибка при подготовке файлов кафедр. Завершение работы")
            sys.exit()
        if preparation_of_points(input("\nВведите путь к файлу осенних баллов: "), input("Введите путь к файлу весенних баллов: ")) != 0:
            print ("\nОшибка при подготовке файлов баллов. Завершение работы")
            sys.exit()
        if processing() != 0:
            input ("\nОшибка при обработке данных. Завершение работы. Нажмите Enter для выхода...")
            sys.exit()
        else:
            input("\nОбработка завершена. Нажмите Enter для выхода...")

    except KeyboardInterrupt:
        input("\nРабота прервана пользователем. Завершение работы. Нажмите Enter для выхода...")
        sys.exit()
        print()

if __name__ == "__main__":
    main()


