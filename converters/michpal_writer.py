"""
Michpal QDIV file writer.
"""

from typing import List, Dict


def _format_value(value: float) -> tuple:
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
    yy = year % 100
    yymm = f"{yy:02d}{month:02d}"
    company_str = company.zfill(3)[:3]

    lines = []
    for rec in employee_records:
        emp_raw = str(rec.get('employee_num', '0')).zfill(9)[:9]
        id_raw = str(rec.get('id_num', '000000000')).zfill(9)[:9]
        bn_raw = rec.get('bruto_neto', ' ') or ' '
        bn = bn_raw[0] if bn_raw[0].isascii() else ' '
        has_cust = rec.get('has_customer', '0')
        code = str(rec.get('report_code', '001')).zfill(3)[:3]

        qty_str, qty_sign = _format_value(float(rec.get('quantity', 0)))
        price_str, price_sign = _format_value(float(rec.get('price', 0)))

        cust_num = str(rec.get('customer_num', '   ')).ljust(3)[:3]
        record_code = rec.get('record_code', '1')

        line = (
            f"{company_str}"
            f"{yymm}"
            f"{emp_raw}"
            f"{id_raw}"
            f"{bn}"
            f"{has_cust}"
            f"{code}"
            f"{qty_str}"
            f"{qty_sign}"
            f"{price_str}"
            f"{price_sign}"
            f"{cust_num}"
            f"      "
            f"{record_code}"
        )
        assert len(line) == 62, f"Record length {len(line)} != 62"
        lines.append(line)

    with open(output_path, 'wb') as f:
        for line in lines:
            f.write(line.encode('ascii') + b'\r\n')
