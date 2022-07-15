import json
import sys
import os
import openpyxl
import requests
import time

from requests.auth import HTTPBasicAuth

priorities = [{'Id': 10, 'Name': 'Высокий'},
              {'Id': 9, 'Name': 'Средний'},
              {'Id': 11, 'Name': 'Обычный'},
              {'Id': 14, 'Name': 'Критический'},
              {'Id': 17, 'Name': '(Аптеки)Средний'},
              {'Id': 18, 'Name': '(Аптеки)Низкий'},
              {'Id': 16, 'Name': '(Аптеки)Высокий'},
              {'Id': 19, 'Name': '(Аптеки)Особый'}
              ]

services = [{'Id': 1206, 'Name': 'Техническая поддержка АПТЕК / ОФИСА'},
            {'Id': 395, 'Name': 'Программа лояльности'},
            {'Id': 195, 'Name': 'Горздрав и 36,6'},
            {'Id': 1155, 'Name': 'Не приходит звонок авторизации (или СМС) клиенту'},
            {'Id': 844, 'Name': 'Активация пластиковой карты'},
            {'Id': 845, 'Name': 'Выпуск виртуальной карты'},
            {'Id': 846, 'Name': 'Начисление, Списание бонусов'},
            {'Id': 881, 'Name': 'Подключение новой аптеки, создание сертификата'},
            {'Id': 193, 'Name': 'Тройка-Город'},
            {'Id': 329, 'Name': 'Аэрофлот -бонус'},
            {'Id': 899, 'Name': 'Мобильное приложение'},
            {'Id': 436, 'Name': 'Интрасервис'},
            {'Id': 25, 'Name': 'Интернет-аптека и внешние сайты'},
            {'Id': 132, 'Name': 'Интернет-аптека и бронирование'},
            {'Id': 250, 'Name': 'Консультации по работе интернет-аптеки'},
            {'Id': 255, 'Name': 'Единое рабочее место первостольника (Сайт https://newpharma.366.ru/)'},
            {'Id': 254, 'Name': 'Сайт https://366.ru/'},
            {'Id': 252, 'Name': 'Сайт https://gorzdrav.org'},
            {'Id': 1134, 'Name': 'Сайты Apteka.ru / ZdravCity.ru'},
            {'Id': 345, 'Name': 'Открытие, закрытие, изменение аптеки'},
            {'Id': 156, 'Name': 'Мониторинг'},
            {'Id': 1124, 'Name': 'Мониторинг систем'},
            {'Id': 323, 'Name': 'Техническая поддержка ОФИСА/СКЛАДА (архивный)'},
            {'Id': 21, 'Name': 'Приложения, сайты, банк-клиенты'},
            {'Id': 27, 'Name': 'ServiceDesk'},
            {'Id': 330, 'Name': 'Обработка обращений покупателей'},
            {'Id': 346, 'Name': 'Обслуживание в аптеке'},
            {'Id': 349, 'Name': 'Вопрос по бонусной программе'},
            {'Id': 507, 'Name': 'Техническая проблема на сайте'},
            {'Id': 1097, 'Name': 'Работа мобильного приложения'},
            {'Id': 508, 'Name': 'Выразить благодарность '},
            {'Id': 509, 'Name': 'Другая тема'},
            {'Id': 1181, 'Name': 'Вопрос по бронированию'},
            {'Id': 1182, 'Name': 'Вопрос по доставке'},
            {'Id': 1201, 'Name': 'Тайный покупатель'},
            {'Id': 1208, 'Name': 'Спасибо от Сбербанка'},
            {'Id': 190, 'Name': 'Рекламные акции, дисконтные карты, купоны, скидки'},
            {'Id': 189, 'Name': 'Выдача заказов на кассе'}]


# def get_services_list(ses, url):
#     cred = '366.dutyadmin@fil-it.ru'
#     response = ses.get(url + 'service?pagesize=200&fields=Id,Name&for=filtertasks',
#                        auth=HTTPBasicAuth(cred, cred))
#     return json.loads(response.content).get('Services')


# print(json.dumps(data, indent=4, sort_keys=False, ensure_ascii=False))

def get_editors(ses, url, task_id):
    editors = []
    response = ses.get(url + f'tasklifetime?taskid={task_id}')
    life_time = json.loads(response.content).get('TaskLifetimes')
    for element in life_time:
        editors.append(element['Editor'])
    return editors


def get_element_name(id, element_list):
    for element in element_list:
        if id == element['Id']:
            return element['Name']


def get_tasks_field(ses, url):
    cred = '366.dutyadmin@fil-it.ru'
    arr = []
    isExecutor = False
    fields = []
    inputs = input(
        'Исполнитель: Дежурный администратор, Казьмина Анастасия Сергеевна, Гусев Дмитрий Александрович: ')
    executors = [executor.strip() for executor in inputs.split(',')]

    start_date = input('Начало периода: ')
    end_date = input('Конец периода: ')
    print("Загрузка данных...")
    # services = get_services_list(ses, url)
    service_list_full = '1206,475,134,192,191,194,395,195,1155,844,845,846,847,881,193,329,190,189,848,230,882,' \
                        '1179,1198,136,327,311,661,316,320,319,1153,313,312,163,672,324,168,314,315,321,322,318,' \
                        '317,178,328,325,1151,1152,326,242,412,500,899,436,25,132,250,254,252,255,1134,127,215,' \
                        '211,212,214,345,406,368,398,414,374,903,156,1124,472,237,323,21,27,330,346,349,507,1097,' \
                        '508,509,1181,1182,1201'
    services_list = '1206,395,195,1155,844,845,846,847,881,193,329,846,847,881,193,329,190,899,436,25,132,250,254,' \
                    '252,255,1134,127,215,211,212,214,345,406,368,398,414,374,156,1124,472,237,323,21,27,330,346,' \
                    '349,507,1097,508,509,1181,1182,1201'
    filtered_services_list = '1206,395,195,1155,844,845,846,881,193,329,899,436,25,132,250,254,252,255,1134,345,156,' \
                             '1124,323,21,27,330,346,349,507,1097,508,509,1181,1182,1201,190,1208,189'

    services_list_full2 = '1206,475,134,192,191,194,395,195,1155,844,845,846,847,881,193,329,1208,190,189,848,230,882,' \
                          '1179,1198,136,327,311,661,316,320,319,1153,313,312,163,672,324,168,314,315,321,322,318,317,' \
                          '178,328,325,1151,1152,326,242,412,500,25,436,899,1209,132,255,254,252,1134,250,127,215,211,' \
                          '212,214,1207,156,1124,1210,472,237,345,406,368,398,414,374,903,330,346,349,507,1097,508,509,' \
                          '1181,1182,1201,323,21,27'
    row = 2
    ws, wb, path = config_excel_file(start_date, end_date)

    response = ses.get(
        url + f'task?CreatedMoreThan={start_date}&CreatedLessThan={end_date}&pagesize=200&ServiceIds='
        + filtered_services_list,
        auth=HTTPBasicAuth(cred, cred))
    paginator = json.loads(response.content).get('Paginator')
    count = 1
    total = paginator['Count']
    for page_number in range(1, paginator['PageCount'] + 1):
        response = ses.get(url +
                           f'task?CreatedMoreThan={start_date}&CreatedLessThan={end_date}'
                           f'&page={page_number}&pagesize=200&ServiceIds=' + filtered_services_list)
        tasks = json.loads(response.content).get('Tasks')
        for task in tasks:
            sys.stdout.write(f"\rВсего / Обработано..........................................{total} / {count}")
            time.sleep(0.4)
            editors = get_editors(ses, url, task['Id'])
            for executor in executors:
                for editor in editors:
                    if executor == editor:
                        isExecutor = True
                        break
                if isExecutor == True or executor == task['Creator'] or executor == task['Executors']:
                    service = get_element_name(task['ServiceId'], services)
                    priority = get_element_name(task['PriorityId'], priorities)
                    category = task['Categories']
                    reporting(service, category, arr)
                    fields.append(task['Id'])
                    fields.append(task['Created'])
                    fields.append(priority)
                    fields.append(task['Name'])
                    fields.append(service)
                    fields.append(category)
                    fields.append(executor)
                    write_field(row, fields, ws, wb, path)
                    fields.clear()
                    row += 1
                isExecutor = False
                wb.close()
            count += 1
    return start_date, end_date, arr


def config_excel_file(start, end):
    path = f'C:\\Users\\Alex\\Downloads\\Статистика_{start}-{end}.xlsx'
    #path = f'C:\\Users\\mfedorov\\Desktop\\Статистика_{start}-{end}.xlsx'
    #path = f'C:\\Users\\kgorshkov\\Desktop\\Статистика\\Общая статистика_{start}-{end}.xlsx'


    try:
        if os.path.exists(path):
            os.remove(path)
        wb = openpyxl.load_workbook(filename=path)
    except FileNotFoundError:
        wb = openpyxl.Workbook()

    ws = wb['Sheet']
    ws.column_dimensions['A'].width = 12
    ws.column_dimensions['B'].width = 25
    ws.column_dimensions['C'].width = 18
    ws.column_dimensions['D'].width = 45
    ws.column_dimensions['E'].width = 32
    ws.column_dimensions['F'].width = 38
    ws.column_dimensions['G'].width = 35
    ws.cell(row=1, column=1, value='№')
    ws.cell(row=1, column=2, value='Создана')
    ws.cell(row=1, column=3, value='Приоритет')
    ws.cell(row=1, column=4, value='Наименование заявки')
    ws.cell(row=1, column=5, value='Сервис')
    ws.cell(row=1, column=6, value='Категория')
    ws.cell(row=1, column=7, value='Исполнитель')

    wb.save(path)
    return ws, wb, path


def create_result_file(first, second, arr):
    row = 2
    column = 1
    total = 0
    path = f'C:\\Users\\Alex\\Downloads\\Результат_{first}-{second}.xlsx'
    #path = f'C:\\Users\\mfedorov\\Desktop\\Итог_{first}-{second}.xlsx'
    #path = f'C:\\Users\\kgorshkov\\Desktop\\Статистика\\Итог_{first}-{second}.xlsx'

    try:
        if os.path.exists(path):
            os.remove(path)
        wb = openpyxl.load_workbook(filename=path)
    except FileNotFoundError:
        wb = openpyxl.Workbook()

    ws = wb['Sheet']
    ws.column_dimensions['A'].width = 65
    ws.column_dimensions['B'].width = 50
    ws.column_dimensions['C'].width = 20
    ws.cell(row=1, column=1, value='Сервис')
    ws.cell(row=1, column=2, value='Категория')
    ws.cell(row=1, column=3, value='Количество')

    # for all ids
    for field in arr:
        total += field[2]
        for item in field:
            ws.cell(row=row, column=column, value=item)
            column += 1
        column = 1
        row += 1
    ws.cell(row=row, column=2, value="Всего")
    ws.cell(row=row, column=3, value=total)
    wb.save(filename=path)


def reporting(service, category, arr):
    if category == '':
        category = None
    if len(arr) == 0:
        arr.append([service, category, 1])
        return
    for element in range(0, len(arr)):
        if arr[element][0] == service and arr[element][1] == category:
            arr[element][2] += 1
            return
        else:
            if element + 1 == len(arr):
                arr.append([service, category, 1])


def write_field(row, fields, ws, wb, path):
    column = 1
    for field in fields:
        ws.cell(row=row, column=column, value=field)
        column += 1
    wb.save(filename=path)


if __name__ == '__main__':
   # cred = '366.dutyadmin@fil-it.ru'
    api_url = 'https://api-sd.366.ru/api/'
    session = requests.Session()
    # response = session.get(api_url + f'tasklifetime?taskid={1729443}',
    #                                  auth=HTTPBasicAuth(cred, cred))
    # life_time = json.loads(response.content).get('TaskLifetimes')
    # print(json.dumps(life_time, indent=4, sort_keys=False, ensure_ascii=False))
    start, end, arr = get_tasks_field(session, api_url)
    create_result_file(start, end, arr)
