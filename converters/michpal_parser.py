"""
Michpal QDIV file parser.
Record layout (62 chars + CRLF = 64 total):
 0- 2  company (3 digits)
 3- 6  yymm (4 chars)
 7-15  employee number (9 digits)
16-24  ID number (9 digits)
25     bruto/neto code (1 char)
26     customer existence (1 char)
27-29  report code (3 digits)
30-39  quantity (10 digits, 2 implied decimal places)
40     quantity sign (+ or -)
41-50  price (10 digits, 2 implied decimal places)
51     price sign (+ or -)
52-54  customer number (3 digits)
55-60  reserved (6 spaces)
61     record code (1 char)
62-63  CR+LF
"""

from dataclasses import dataclass
from typing import List


BUILTIN_CODE_NAMES = {
    '100': 'ימי עבודה (משולמים)',
    '120': 'ימי עבודה + שעות/יום',
    '101': 'ניצול חופשה',
    '102': 'ניצול מחלה',
    '103': 'ניצול הבראה',
    '104': 'ניצול מילואים',
    '111': 'צבירת חופשה',
    '112': 'צבירת מחלה',
    '113': 'צבירת הבראה',
    '121': 'כמות בסיס',
    '122': 'שעות עבודה (משולמות)',
    '123': 'שעות היעדרות',
    '124': 'ימי עבודה בפועל',
    '125': 'שעות עבודה בפועל',
    '126': 'ימי עבודה (תקן)',
    '127': 'שעות עבודה (תקן)',
    '131': 'מחלקה',
    '901': 'קוד הפסקה (1)',
    '902': 'איפוס רכיבי שכר',
    '910': 'קוד הפסקה (5)',
}


@dataclass
class MichpalRecord:
    company: str
    yymm: str
    employee_num_raw: str
    id_num: str
    bruto_neto: str
    has_customer: str
    report_code: str
    quantity_raw: str
    qty_sign: str
    price_raw: str
    price_sign: str
    customer_num: str
    reserved: str
    record_code: str

    @property
    def employee_num(self) -> str:
        return self.employee_num_raw.lstrip('0') or '0'

    @property
    def quantity(self) -> float:
        val = int(self.quantity_raw) / 100
        return -val if self.qty_sign == '-' else val

    @property
    def price(self) -> float:
        val = int(self.price_raw) / 100
        return -val if self.price_sign == '-' else val

    @property
    def year(self) -> int:
        return 2000 + int(self.yymm[:2])

    @property
    def month(self) -> int:
        return int(self.yymm[2:4])

    def get_component_number(self) -> int:
        code = int(self.report_code)
        if 400 <= code <= 552:
            return code - 300
        return code


def parse_michpal_file(filepath: str) -> List[MichpalRecord]:
    with open(filepath, 'rb') as f:
        data = f.read()

    records = []
    for line in data.split(b'\r\n'):
        if not line or len(line) < 62:
            continue
        try:
            s = line.decode('cp862', errors='replace')
        except Exception:
            s = line.decode('latin-1', errors='replace')

        if len(s) < 62:
            continue

        record = MichpalRecord(
            company=s[0:3],
            yymm=s[3:7],
            employee_num_raw=s[7:16],
            id_num=s[16:25],
            bruto_neto=s[25],
            has_customer=s[26],
            report_code=s[27:30],
            quantity_raw=s[30:40],
            qty_sign=s[40],
            price_raw=s[41:51],
            price_sign=s[51],
            customer_num=s[52:55],
            reserved=s[55:61],
            record_code=s[61],
        )
        records.append(record)

    return records


def get_michpal_meta(records: List[MichpalRecord]) -> dict:
    if not records:
        return {}
    r = records[0]
    unique_employees = len(set(rec.employee_num for rec in records))
    return {
        'company': r.company,
        'year': r.year,
        'month': r.month,
        'yymm': r.yymm,
        'total_records': len(records),
        'unique_employees': unique_employees,
    }
