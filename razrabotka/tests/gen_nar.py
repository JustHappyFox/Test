# Вымышленная выгрузка нарядов той же структуры, что у пользователя (лист TEMP, 22 столбца); лицевые - из gen_byt.py
# python3 gen_nar.py out.xlsx [нарядов]
import sys, datetime, openpyxl
out = sys.argv[1]; n = int(sys.argv[2]) if len(sys.argv) > 2 else 300
H = ['NAB','ФИО','NP','УЛИЦА','ДОМ','КОР','КВ','Н_З','ДАТА_ОФ','ФИО_ПР','ТИП_СЧ','НОМЕР_СЧ','РАЗ','ГОД_П','ДАТА_ПР','ПОК_ПР','ОПЛ_ПОК','РАБОТА','ВИД_АСКУЭ','ДАТА_ВЫП','ФИО_ВЫП','ПРИМЕЧАНИЕ']
works = ['Замена сч.', 'Проверка сч.', 'Опломбировка', 'Отключение']
wb = openpyxl.Workbook(); ws = wb.active; ws.title = 'TEMP'; ws.append(H)
for i in range(n):
    nab = 10019999 if i % 50 == 0 else int('0%04d%03d%d' % (1000 + i // 5, i % 1000, i % 10))
    year = 2026 if i % 3 else 2025
    done = i % 2 == 0
    ws.append([nab, 'Тестов Иван Иванович', 'Тестовка д.', 'Лесная ул.', i % 60 + 1, None, None, year * 100000 + 10000 + i,
               datetime.datetime(year, 1 + i % 12, 1 + i % 27), 'Приёмов Пётр Петрович', 'СЕ-318BY', str(5000 + i), 6, 2020,
               datetime.datetime(year, 1 + i % 12, 2), 0, 1000 + i, works[i % 4], 'Радио модуль',
               datetime.datetime(year, 1 + i % 12, 5) if done else None, 'Монтёров Олег Олегович' if done else None, 'тестовый наряд'])
wb.save(out); print(out, n)
