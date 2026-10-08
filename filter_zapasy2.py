# -*- coding: utf-8 -*-
"""
Перепроверка «Запасов» с учётом упаковочных листов.
Строка попадает в итоговый файл, если:
  1) № чертежа точно совпадает со списком «Виды оборудования», ИЛИ
  2) Номер упаковочного листа есть среди упаковочников листа «Оборудование» файла БМНЛЗ, ИЛИ
  3) Номер упаковочного листа совпадает с упаковочником любой строки, совпавшей по чертежу
     (перенос внутри самих «Запасов», по двум складам).
Ничего, кроме строк с найденным совпадением, в файл не вставляется.
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

SRC_TYPES = 'Виды оборудования.xlsx'
SRC_STOCK = 'Запасы 30.09.2026.XLSX'
SRC_BMNZ = 'Накопительная оборудование (БМНЛЗ).xlsx'
OUT = 'Запасы_отфильтрованные.xlsx'


def norm(v):
    if v is None:
        return ''
    s = str(v).strip()
    return s[:-2] if s.endswith('.0') else s


# 1. Виды оборудования — номера чертежей
wb = openpyxl.load_workbook(SRC_TYPES, read_only=True, data_only=True)
types_set = {norm(r[0]) for r in wb['Лист1'].iter_rows(min_row=2, values_only=True) if r[0] is not None}
wb.close()
types_set.discard('')

# 2. БМНЛЗ «Оборудование» — упаковочники (колонка 15), привязанные к чертежам
wb = openpyxl.load_workbook(SRC_BMNZ, read_only=True, data_only=True)
packs_bmnlz = set()
for row in wb['Оборудование'].iter_rows(min_row=4, values_only=True):
    p = norm(row[15])
    if p:
        packs_bmnlz.add(p)
wb.close()
print(f'Видов оборудования: {len(types_set)} | Упаковочников БМНЛЗ: {len(packs_bmnlz)}')

# 3. Читаем Запасы целиком (оба склада)
wb = openpyxl.load_workbook(SRC_STOCK, read_only=True, data_only=True)
sheets = {}
for sheet_name in ['центр.склад', 'Склад подрядчика']:
    ws = wb[sheet_name]
    it = ws.iter_rows(values_only=True)
    header = list(next(it))
    rows = [list(r) for r in it]
    sheets[sheet_name] = (header, rows)
wb.close()

# 4. Упаковочники строк, совпавших по чертежу (перенос внутри Запасов)
prop_packs = set()
for sheet_name, (header, rows) in sheets.items():
    for r in rows:
        if norm(r[6]) in types_set:
            p = norm(r[8])
            if p:
                prop_packs.add(p)
print(f'Упаковочников со строками, совпавшими по чертежу: {len(prop_packs)}')

# 5. Итоговая фильтрация
wb_out = openpyxl.Workbook()
ws_out = wb_out.active
ws_out.title = 'Запасы (отфильтрованные)'
header_font = Font(bold=True, color='FFFFFF')
header_fill = PatternFill('solid', fgColor='4472C4')

header_written = False
stats = {}
for sheet_name in ['центр.склад', 'Склад подрядчика']:
    header, rows = sheets[sheet_name]
    if not header_written:
        ws_out.append(header)
        for cell in ws_out[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        header_written = True
    total = 0
    kept = 0
    c_draw = c_pack = c_prop = c_only_draw = 0
    for r in rows:
        total += 1
        d = norm(r[6])
        p = norm(r[8])
        m_draw = d in types_set
        m_pack = p in packs_bmnlz
        m_prop = bool(p) and p in prop_packs
        if m_draw or m_pack or m_prop:
            ws_out.append(r)
            kept += 1
            if m_draw:
                c_draw += 1
                if not m_pack and not m_prop:
                    c_only_draw += 1
            if m_pack:
                c_pack += 1
            if m_prop and not m_draw and not m_pack:
                c_prop += 1
    stats[sheet_name] = (total, kept, c_draw, c_pack, c_prop, c_only_draw)
    print(f'{sheet_name}: всего {total}, отобрано {kept} (по чертежу {c_draw}, из них только чертёж {c_only_draw}; '
          f'упаковочник БМНЛЗ {c_pack}; только перенос {c_prop})')

ws_out.freeze_panes = 'A2'
widths = [10, 22, 10, 14, 18, 50, 12, 12, 14, 9, 9, 9, 11, 13, 14]
for i, w in enumerate(widths, start=1):
    ws_out.column_dimensions[get_column_letter(i)].width = w

wb_out.save(OUT)
print('\nСохранено:', OUT)
print('Итого строк:', sum(s[1] for s in stats.values()))
