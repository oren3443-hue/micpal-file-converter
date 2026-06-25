"""
Michpal QDIV file writer.
"""

from typing import List, Dict


def _format_value(value: float) -> tuple:
    """Return (10-digit zero-padded string, sign char) for a numeric value."""
    sign = '+' if value >= 0 else '-'
    int_val = round(abs(value) * 100)
    return f"{int_val:010d}", sign


def write_michpal_file(
    employee_records: List[Dict],
    output_path: str,
    company: str,
    year: int,
    month: int,
) -> None:
    """
    Write a QDIV file from a list of employee record dicts.

    Each dict in employee_records:
      {
        'employee_num': str,   # 1-9999
        'id_num': str,         # 9-digit ID (optional, use '000000000' if unknown)
        'bruto_neto': str,     # 'ב'|'נ'|'ק'|'פ'|'ג'|' '
        'report_code': str,    # 3-digit code
        'quantity': float,
        'price': float,
        'has_customer': str,   # '0'|'1'
        'customer_num': str,   # '000' if no customer
        'record_code': str,    # '1' (normal) or '9' (informational)
      }
    """
    yy = year % 100
    yymm = f"{yy:02d}{month:02d}"
    company_str = company.zfill(3)[:3]

    lines = []
    for rec in employee_records:
        emp_raw = str(rec.get('employee_num', '0')).zfill(9)[:9]
        id_raw = str(rec.get('id_num', '000000000')).zfill(9)[:9]
        bn_raw = rec.get('bruto_neto', ' ') or ' '
        # Keep only ASCII printable; Hebrew letters not needed in practice (space = default)
        bn = bn_raw[0] if bn_raw[0].isascii() else ' '
        has_cust = rec.get('has_customer', '0')
        code = str(rec.get('report_code', '001')).zfill(3)[:3]

        qty_str, qty_sign = _format_value(float(rec.get('quantity', 0)))
        price_str, price_sign = _format_value(float(rec.get('price', 0)))

        cust_num = str(rec.get('customer_num', '   ')).ljust(3)[:3]
        record_code = rec.get('record_code', '1')

        line = (
            f"{company_str}"    # 0-2
            f"{yymm}"           # 3-6
            f"{emp_raw}"        # 7-15
            f"{id_raw}"         # 16-24
            f"{bn}"             # 25
            f"{has_cust}"       # 26
            f"{code}"           # 27-29
            f"{qty_str}"        # 30-39
            f"{qty_sign}"       # 40
            f"{price_str}"      # 41-50
            f"{price_sign}"     # 51
            f"{cust_num}"       # 52-54
            f"      "           # 55-60 reserved
            f"{record_code}"    # 61
        )
        assert len(line) == 62, f"Record length {len(line)} != 62"
        lines.append(line)

    with open(output_path, 'wb') as f:
        for line in lines:
            f.write(line.encode('ascii') + b'\r\n')
