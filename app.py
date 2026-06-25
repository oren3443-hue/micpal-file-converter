import os
import re
import tempfile
import traceback
from datetime import datetime
from pathlib import Path

from flask import (
    Flask, request, send_file, render_template,
    jsonify, after_this_request,
)

from converters.michpal_parser import parse_michpal_file, get_michpal_meta
from converters.michpal_writer import write_michpal_file
from converters.excel_parser import parse_excel_file, DEFAULT_COLUMN_MAPPING
from converters.excel_writer import write_excel_from_michpal
from converters.pdf_parser import extract_component_names, extract_component_names_from_excel

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50 MB

UPLOAD_DIR = Path(tempfile.mkdtemp(prefix='michpal_'))


def _tmp(suffix: str) -> str:
    fd, path = tempfile.mkstemp(suffix=suffix, dir=UPLOAD_DIR)
    os.close(fd)
    return path


def _parse_michpal_filename(filename: str):
    """
    Try to extract (company, month, year) from a QDIV filename.
    Accepted patterns:
      - QDIV0626.010  →  month=06, year=2026, company=010
      - 010.QDIV0626  →  same
    """
    name = filename.upper()
    m = re.search(r'QDIV(\d{2})(\d{2})\.(\d+)', name)
    if m:
        mm, yy, company = int(m.group(1)), int(m.group(2)), m.group(3).zfill(3)
        return company, mm, 2000 + yy
    m = re.search(r'(\d+)\.QDIV(\d{2})(\d{2})', name)
    if m:
        company, mm, yy = m.group(1).zfill(3), int(m.group(2)), int(m.group(3))
        return company, mm, 2000 + yy
    return None, None, None


# ── Routes ────────────────────────────────────────────────────────────────────

@app.route('/')
def index():
    return render_template('index.html')


@app.route('/api/michpal-to-excel', methods=['POST'])
def michpal_to_excel():
    try:
        if 'michpal_file' not in request.files:
            return jsonify(error='נא להעלות קובץ מיכפל'), 400

        mf = request.files['michpal_file']
        mpath = _tmp('.010')
        mf.save(mpath)

        # Optional PDF or Excel for component names
        component_names = {}
        mapping_file = request.files.get('pdf_file')
        if mapping_file and mapping_file.filename:
            fname = mapping_file.filename.lower()
            if fname.endswith(('.xlsx', '.xls')):
                mfile_path = _tmp('.xlsx')
                mapping_file.save(mfile_path)
                component_names = extract_component_names_from_excel(mfile_path)
            else:
                mfile_path = _tmp('.pdf')
                mapping_file.save(mfile_path)
                component_names = extract_component_names(mfile_path)

        records = parse_michpal_file(mpath)
        if not records:
            return jsonify(error='הקובץ ריק או בפורמט שגוי'), 400

        meta = get_michpal_meta(records)
        out = _tmp('.xlsx')
        write_excel_from_michpal(records, out, component_names)

        dl_name = f"michpal_{meta['year']}{meta['month']:02d}.xlsx"

        @after_this_request
        def _cleanup(response):
            try:
                os.unlink(out)
                os.unlink(mpath)
            except Exception:
                pass
            return response

        return send_file(out, as_attachment=True, download_name=dl_name,
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    except Exception as exc:
        return jsonify(error=str(exc), trace=traceback.format_exc()), 500


@app.route('/api/excel-to-michpal', methods=['POST'])
def excel_to_michpal():
    try:
        if 'excel_file' not in request.files:
            return jsonify(error='נא להעלות קובץ אקסל'), 400

        ef = request.files['excel_file']
        epath = _tmp('.xlsx')
        ef.save(epath)

        company  = request.form.get('company', '001').strip().zfill(3)
        year_str = request.form.get('year', str(datetime.now().year))
        month_str = request.form.get('month', str(datetime.now().month))
        year  = int(year_str)
        month = int(month_str)

        # Override with custom column mapping if provided
        col_mapping_json = request.form.get('col_mapping', '')
        custom_mapping = None
        if col_mapping_json:
            import json
            try:
                raw = json.loads(col_mapping_json)
                # raw is {header: {report_code, field}}
                custom_mapping = {h: (v['report_code'], v['field']) for h, v in raw.items()}
            except Exception:
                pass

        parsed = parse_excel_file(epath, custom_mapping)
        employees = parsed['employees']
        if not employees:
            return jsonify(error='לא נמצאו עובדים עם נתוני שכר'), 400

        # Flatten to list of dicts for writer
        rows = []
        for emp in employees:
            for comp in emp['components']:
                rows.append({
                    'employee_num': emp['employee_num'],
                    'id_num': emp['id_num'],
                    'bruto_neto': 'ב',
                    'has_customer': '0',
                    'customer_num': '   ',
                    'record_code': '1',
                    **comp,
                })

        out = _tmp('.010')
        write_michpal_file(rows, out, company, year, month)

        mm = month
        yy = year % 100
        dl_name = f"QDIV{mm:02d}{yy:02d}.{company}"

        @after_this_request
        def _cleanup(response):
            try:
                os.unlink(out)
                os.unlink(epath)
            except Exception:
                pass
            return response

        return send_file(out, as_attachment=True, download_name=dl_name,
                         mimetype='application/octet-stream')

    except Exception as exc:
        return jsonify(error=str(exc), trace=traceback.format_exc()), 500


@app.route('/api/inspect-excel', methods=['POST'])
def inspect_excel():
    """Return column info from an Excel file so the UI can show a mapping editor."""
    try:
        if 'excel_file' not in request.files:
            return jsonify(error='חסר קובץ'), 400

        ef = request.files['excel_file']
        epath = _tmp('.xlsx')
        ef.save(epath)

        import openpyxl
        wb = openpyxl.load_workbook(epath, data_only=True, read_only=True)
        ws = wb.active

        # Find header row
        headers = []
        for row in ws.iter_rows(min_row=1, max_row=10, values_only=True):
            non_empty = [c for c in row if c is not None]
            if len(non_empty) >= 3:
                headers = [str(c).strip() if c is not None else '' for c in row]
                break

        wb.close()

        # Suggest mappings
        suggestions = []
        for h in headers:
            if not h:
                continue
            mapped = DEFAULT_COLUMN_MAPPING.get(h)
            suggestions.append({
                'header': h,
                'report_code': mapped[0] if mapped else '',
                'field': mapped[1] if mapped else 'qty',
            })

        return jsonify(columns=suggestions)

    except Exception as exc:
        return jsonify(error=str(exc)), 500


@app.route('/api/inspect-michpal', methods=['POST'])
def inspect_michpal():
    """Return summary of a QDIV file."""
    try:
        if 'michpal_file' not in request.files:
            return jsonify(error='חסר קובץ'), 400

        mf = request.files['michpal_file']
        mpath = _tmp('.010')
        mf.save(mpath)

        records = parse_michpal_file(mpath)
        meta = get_michpal_meta(records)

        from converters.excel_writer import _component_code, _name_for_code
        codes = sorted(set(r.report_code for r in records))
        code_info = [{'code': _component_code(c), 'name': _name_for_code(c, {})} for c in codes]

        return jsonify({**meta, 'codes': code_info})

    except Exception as exc:
        return jsonify(error=str(exc)), 500


@app.route('/api/load-component-mapping', methods=['POST'])
def load_component_mapping():
    """Return {code: name} dict from an uploaded Excel or PDF mapping file."""
    try:
        if 'mapping_file' not in request.files:
            return jsonify(error='חסר קובץ'), 400
        f = request.files['mapping_file']
        fname = f.filename.lower()
        if fname.endswith(('.xlsx', '.xls')):
            fpath = _tmp('.xlsx')
            f.save(fpath)
            names = extract_component_names_from_excel(fpath)
        else:
            fpath = _tmp('.pdf')
            f.save(fpath)
            names = extract_component_names(fpath)
        return jsonify(mapping=names)
    except Exception as exc:
        return jsonify(error=str(exc)), 500


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
