"""
Write a formatted RTL Hebrew Excel file from Michpal records.
"""

from collections import defaultdict
from typing import List, Dict, Optional

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

from .michpal_parser import MichpalRecord, BUILTIN_CODE_NAMES

CLR_HEADER_BG = '1F4E79'
CLR_HEADER_FG = 'FFFFFF'
CLR_ROW_ALT   = 'DDEEFF'
CLR_ROW_WHITE = 'FFFFFF'
CLR_SUBHEADER = '2E75B6'
CLR_BORDER    = 'AAAAAA'


def _border(style='thin'):
    s = Side(style=style, color=CLR_BORDER)
    return Border(left=s, right=s, top=s, bottom=s)


def _header_cell(ws, row, col, value, wrap=True):
    cell = ws.cell(row=row, column=col, value=value)
    cell.font = Font(name='Arial', bold=True, size=10, color=CLR_HEADER_FG)
    cell.fill = PatternFill(start_color=CLR_HEADER_BG, end_color=CLR_HEADER_BG, fill_type='solid')
    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=wrap)
    cell.border = _border()
    return cell


def _data_cell(ws, row, col, value, is_alt=False, number_format=None, bold=False):
    cell = ws.cell(row=row, column=col, value=value)
    cell.font = Font(name='Arial', size=10, bold=bold)
    bg = CLR_ROW_ALT if is_alt else CLR_ROW_WHITE
    cell.fill = PatternFill(start_color=bg, end_color=bg, fill_type='solid')
    cell.alignment = Alignment(horizontal='center', vertical='center')
    cell.border = _border()
    if number_format:
        cell.number_format = number_format
    return cell


def write_excel_from_michpal(
    records: List[MichpalRecord],
    output_path: str,
    component_names: Optional[Dict[str, str]] = None,
) -> None:
    if not records:
        raise ValueError('אין רשומות לכתיבה')

    names: Dict[str, str] = {**BUILTIN_CODE_NAMES, **(component_names or {})}

    employees: Dict[str, Dict] = defaultdict(lambda: {
        'id_num': '',
        'codes': {},
        'company': '',
        'yymm': '',
    })

    all_codes = sorted(set(r.report_code for r in records))

    for rec in records:
        if rec.record_code == '9':
            continue
        emp = rec.employee_num
        employees[emp]['id_num'] = rec.id_num
        employees[emp]['company'] = rec.company
        employees[emp]['yymm'] = rec.yymm
        if rec.report_code not in employees[emp]['codes']:
            employees[emp]['codes'][rec.report_code] = {'qty': 0.0, 'price': 0.0}
        employees[emp]['codes'][rec.report_code]['qty'] += rec.quantity
        if rec.price != 0:
            employees[emp]['codes'][rec.report_code]['price'] = rec.price

    codes_with_price: set = set()
    for emp_data in employees.values():
        for code, vals in emp_data['codes'].items():
            if vals['price'] != 0:
                codes_with_price.add(code)

    col_defs = []
    for code in all_codes:
        code_name = names.get(code, f'רכיב {code}')
        col_defs.append((code, 'qty', f'{code_name}\n(כמות)'))
        if code in codes_with_price:
            col_defs.append((code, 'price', f'{code_name}\n(מחיר)'))

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'נתוני מיכפל'
    ws.sheet_view.rightToLeft = True

    first = records[0]
    period_label = f'נתוני שכר {first.month:02d}/{first.year}  |  חברה {first.company}'
    total_cols = 2 + len(col_defs)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=total_cols)
    cell = ws.cell(row=1, column=1, value=period_label)
    cell.font = Font(name='Arial', bold=True, size=13, color=CLR_HEADER_FG)
    cell.fill = PatternFill(start_color=CLR_SUBHEADER, end_color=CLR_SUBHEADER, fill_type='solid')
    cell.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 28

    ws.row_dimensions[2].height = 50
    _header_cell(ws, 2, 1, 'מספר עובד')
    _header_cell(ws, 2, 2, 'מספר זהות')
    for c_idx, (code, field, label) in enumerate(col_defs, 3):
        _header_cell(ws, 2, c_idx, label)

    sorted_employees = sorted(employees.items(),
                              key=lambda kv: int(kv[0]) if kv[0].isdigit() else 0)
    for row_offset, (emp_num, emp_data) in enumerate(sorted_employees):
        row = 3 + row_offset
        is_alt = row_offset % 2 == 0
        _data_cell(ws, row, 1, emp_num, is_alt)
        _data_cell(ws, row, 2, emp_data['id_num'], is_alt)
        for c_idx, (code, field, _label) in enumerate(col_defs, 3):
            vals = emp_data['codes'].get(code, {'qty': 0.0, 'price': 0.0})
            val = vals[field]
            display_val = round(val, 2) if val != 0 else None
            _data_cell(ws, row, c_idx, display_val, is_alt, number_format='#,##0.00')

    last_data_row = 2 + len(sorted_employees)
    ws.auto_filter.ref = f'A2:{get_column_letter(total_cols)}{last_data_row}'
    ws.freeze_panes = 'A3'
    ws.column_dimensions['A'].width = 13
    ws.column_dimensions['B'].width = 13
    for c_idx in range(3, total_cols + 1):
        ws.column_dimensions[get_column_letter(c_idx)].width = 16

    ws2 = wb.create_sheet('סיכום')
    ws2.sheet_view.rightToLeft = True
    for i, (k, v) in enumerate([
        ('מספר חברה', first.company),
        ('שנה', first.year),
        ('חודש', first.month),
        ('מספר עובדים', len(employees)),
        ('מספר רשומות', len(records)),
        ('סמלי דיווח', ', '.join(all_codes)),
    ], 1):
        ws2.cell(row=i, column=1, value=k).font = Font(bold=True, name='Arial')
        ws2.cell(row=i, column=2, value=v).font = Font(name='Arial')
    ws2.column_dimensions['A'].width = 20
    ws2.column_dimensions['B'].width = 40

    wb.save(output_path)
