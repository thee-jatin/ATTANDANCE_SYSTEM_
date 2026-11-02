import io
import os
import re
from contextlib import closing
from datetime import datetime

import pandas as pd
import psycopg2
import psycopg2.extras
import xlsxwriter
from flask import Flask, request, jsonify, send_file, send_from_directory
from flask_cors import CORS
from xlsxwriter.utility import xl_col_to_name

app = Flask(__name__)
CORS(app, expose_headers=['Content-Disposition'])

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Password code mein nahi likhna. Environment variable DATABASE_URL set karo.
DATABASE_URL = os.environ.get("DATABASE_URL")
if not DATABASE_URL:
    raise RuntimeError("DATABASE_URL environment variable set nahi hai (Supabase connection string).")

EXCEL_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
REQUIRED_COLUMNS = ['Roll No.', 'Name', 'Course', 'Sem', 'Stream']


def get_db_connection():
    return psycopg2.connect(DATABASE_URL)


def init_db():
    try:
        with closing(get_db_connection()) as conn, conn:
            cur = conn.cursor()
            cur.execute('''
                CREATE TABLE IF NOT EXISTS students (
                    id SERIAL PRIMARY KEY,
                    rollno VARCHAR(50) NOT NULL,
                    name VARCHAR(100) NOT NULL,
                    fathername VARCHAR(100),
                    mobile VARCHAR(15),
                    class_name VARCHAR(50) NOT NULL,
                    sem VARCHAR(10) NOT NULL,
                    stream VARCHAR(100) NOT NULL
                )
            ''')
            cur.execute('''
                CREATE TABLE IF NOT EXISTS attendance (
                    id SERIAL PRIMARY KEY,
                    date DATE NOT NULL,
                    lecture VARCHAR(50) NOT NULL,
                    class_name VARCHAR(50) NOT NULL,
                    sem VARCHAR(10) NOT NULL,
                    student_id INT REFERENCES students(id) ON DELETE CASCADE,
                    status VARCHAR(5) NOT NULL
                )
            ''')
        print("Database tables checked/created successfully!")
    except Exception as e:
        print("Database connection error:", e)


init_db()


# ---------- helpers ----------

def clean(value):
    """Excel cell ko saaf text banata hai: nan -> '', 101.0 -> '101'."""
    if value is None:
        return ''
    s = str(value).strip()
    if s.lower() in ('nan', 'none', 'nat'):
        return ''
    if re.fullmatch(r'\d+\.0+', s):
        s = s.split('.')[0]
    return s


def natural_key(text):
    """Roll no. ko 2 < 10 ke hisab se sort karne ke liye."""
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r'(\d+)', text or '')]


def error(msg, code=500):
    return jsonify({'error': msg}), code


# ---------- frontend serve (taaki mobile/laptop dono ek hi link se khule) ----------

@app.route('/')
def index():
    return send_from_directory(BASE_DIR, 'frontend.html')


@app.route('/logo.jpeg')
def logo():
    return send_from_directory(BASE_DIR, 'logo.jpeg')


# ---------- students ----------

@app.route('/api/students', methods=['GET'])
def get_students():
    try:
        with closing(get_db_connection()) as conn:
            cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
            cur.execute('SELECT * FROM students ORDER BY id ASC')
            return jsonify(cur.fetchall())
    except Exception as e:
        return error(str(e))


@app.route('/api/students', methods=['POST'])
def add_student():
    data = request.get_json(silent=True) or {}
    rollno = clean(data.get('rollno'))
    name = clean(data.get('name'))
    class_name = clean(data.get('className'))
    sem = clean(data.get('sem'))
    stream = clean(data.get('stream'))
    if not (rollno and name and class_name and sem and stream):
        return error('Roll No., Name, Class, Sem aur Stream zaroori hain.', 400)

    try:
        with closing(get_db_connection()) as conn, conn:
            cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
            cur.execute(
                '''INSERT INTO students (rollno, name, fathername, mobile, class_name, sem, stream)
                   SELECT %s, %s, %s, %s, %s, %s, %s
                   WHERE NOT EXISTS (
                       SELECT 1 FROM students WHERE rollno = %s AND class_name = %s AND sem = %s
                   )
                   RETURNING *''',
                (rollno, name[:100], clean(data.get('fathername'))[:100], clean(data.get('mobile'))[:15],
                 class_name[:50], sem[:10], stream[:100],
                 rollno, class_name, sem)
            )
            new_student = cur.fetchone()
        if new_student is None:
            return error(f'Roll No. {rollno} is class/sem mein pehle se hai.', 409)
        return jsonify(new_student)
    except Exception as e:
        return error(str(e))


@app.route('/api/students/bulk', methods=['POST'])
def bulk_upload_students():
    file = request.files.get('file')
    if not file or not file.filename:
        return error('Excel file select karo.', 400)

    try:
        if file.filename.lower().endswith('.csv'):
            df = pd.read_csv(file, dtype=str)
        else:
            df = pd.read_excel(file, dtype=str)
    except Exception:
        return error('File padhi nahi ja saki. Diye gaye format ki .xlsx ya .csv file upload karo.', 400)

    df.columns = [str(c).strip() for c in df.columns]
    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        return error('File mein ye columns nahi mile: ' + ', '.join(missing) +
                     '. Diye gaye Excel format ka use karo.', 400)

    added = 0
    skipped = []
    classes = {}

    try:
        with closing(get_db_connection()) as conn, conn:
            cur = conn.cursor()
            for idx, row in df.iterrows():
                excel_row = idx + 2  # header row 1 hai
                roll = clean(row.get('Roll No.'))
                name = clean(row.get('Name'))
                father = clean(row.get('Father Name'))
                mobile = clean(row.get('Mobile'))
                course = clean(row.get('Course'))
                sem = clean(row.get('Sem'))
                stream = clean(row.get('Stream'))

                if not any([roll, name, father, mobile, course, sem, stream]):
                    continue  # poori khaali row, chup-chaap chhod do

                if not (roll and name and course and sem and stream):
                    skipped.append(f'Row {excel_row}: Roll No., Name, Course, Sem ya Stream khaali hai')
                    continue

                cur.execute(
                    '''INSERT INTO students (rollno, name, fathername, mobile, class_name, sem, stream)
                       SELECT %s, %s, %s, %s, %s, %s, %s
                       WHERE NOT EXISTS (
                           SELECT 1 FROM students WHERE rollno = %s AND class_name = %s AND sem = %s
                       )''',
                    (roll[:50], name[:100], father[:100], mobile[:15], course[:50], sem[:10], stream[:100],
                     roll, course, sem)
                )
                if cur.rowcount:
                    added += 1
                    classes[(course, sem)] = classes.get((course, sem), 0) + 1
                else:
                    skipped.append(f'Row {excel_row}: Roll No. {roll} pehle se hai ({course} Sem {sem})')
    except Exception as e:
        return error(str(e))

    if added == 0 and not skipped:
        return error('File mein koi student data nahi mila.', 400)

    return jsonify({
        'added': added,
        'classes': [f'{c} Sem {s}: {n} students' for (c, s), n in classes.items()],
        'skipped_count': len(skipped),
        'skipped': skipped[:15],
    })


@app.route('/api/students/<int:student_id>', methods=['DELETE'])
def delete_student(student_id):
    try:
        with closing(get_db_connection()) as conn, conn:
            conn.cursor().execute('DELETE FROM students WHERE id = %s', (student_id,))
        return jsonify({'message': 'Student removed successfully'})
    except Exception as e:
        return error(str(e))


# ---------- attendance ----------

@app.route('/api/attendance', methods=['POST'])
def save_attendance():
    data = request.get_json(silent=True) or {}
    date = data.get('date')
    lecture = data.get('lecture')
    class_name = data.get('className')
    sem = data.get('sem')
    records = data.get('records') or {}  # {student_id: 'P'/'A'}

    if not (date and lecture and class_name and sem) or not records:
        return error('Date, Lecture, Class, Sem aur students zaroori hain.', 400)

    rows = [(date, lecture, class_name, sem, int(sid), 'P' if st == 'P' else 'A')
            for sid, st in records.items()]
    try:
        with closing(get_db_connection()) as conn, conn:
            cur = conn.cursor()
            # Same date + lecture + class + sem dobara save ho to purana replace hota hai
            cur.execute(
                'DELETE FROM attendance WHERE date = %s AND lecture = %s AND class_name = %s AND sem = %s',
                (date, lecture, class_name, sem)
            )
            psycopg2.extras.execute_values(
                cur,
                'INSERT INTO attendance (date, lecture, class_name, sem, student_id, status) VALUES %s',
                rows
            )
        return jsonify({'message': 'Attendance saved successfully'})
    except Exception as e:
        return error(str(e))


@app.route('/api/attendance', methods=['GET'])
def get_all_attendance():
    try:
        with closing(get_db_connection()) as conn:
            cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
            cur.execute(
                "SELECT id, to_char(date, 'YYYY-MM-DD') AS date, lecture, class_name, sem, student_id, status "
                "FROM attendance"
            )
            return jsonify(cur.fetchall())
    except Exception as e:
        return error(str(e))


# ---------- statement + Excel ----------

def fetch_statement(class_name, sem, from_date, to_date):
    with closing(get_db_connection()) as conn:
        cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        cur.execute(
            'SELECT id, rollno, name, fathername FROM students WHERE class_name = %s AND sem = %s',
            (class_name, sem)
        )
        students = sorted(cur.fetchall(), key=lambda s: natural_key(s['rollno']))

        query = ("SELECT student_id, to_char(date, 'YYYY-MM-DD') AS date, lecture, status "
                 "FROM attendance WHERE class_name = %s AND sem = %s")
        params = [class_name, sem]
        if from_date and to_date:
            query += ' AND date BETWEEN %s AND %s'
            params += [from_date, to_date]
        cur.execute(query, params)
        records = cur.fetchall()
    return students, records


def build_grid(students, records):
    """(date, lecture) columns aur har student ke P/A ka hisaab banata hai."""
    status = {(r['student_id'], r['date'], r['lecture']): r['status'] for r in records}
    columns = sorted({(r['date'], r['lecture']) for r in records},
                     key=lambda c: (c[0], natural_key(c[1])))
    summary = []
    for s in students:
        present = sum(1 for d, l in columns if status.get((s['id'], d, l)) == 'P')
        total = sum(1 for d, l in columns if (s['id'], d, l) in status)
        summary.append({
            'rollno': s['rollno'], 'name': s['name'], 'fathername': s['fathername'] or '',
            'present': present, 'total': total,
            'percent': round(present * 100 / total, 1) if total else 0,
        })
    return columns, status, summary


def build_excel(students, columns, status, summary):
    output = io.BytesIO()
    wb = xlsxwriter.Workbook(output, {'in_memory': True})
    ws = wb.add_worksheet('Attendance')

    font = {'font_name': 'Arial', 'font_size': 10}
    head = wb.add_format({**font, 'bold': True, 'bg_color': '#DCFCE7', 'border': 1,
                          'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
    left = wb.add_format({**font, 'border': 1})
    center = wb.add_format({**font, 'border': 1, 'align': 'center'})
    percent_fmt = wb.add_format({**font, 'border': 1, 'align': 'center', 'num_format': '0.0%'})
    green = wb.add_format({'font_color': '#16A34A', 'bold': True})
    red = wb.add_format({'font_color': '#DC2626', 'bold': True})

    first = 3                      # D column se dates shuru
    n = len(columns)
    c_present, c_total, c_pct = first + n, first + n + 1, first + n + 2

    for c, title in enumerate(['Roll No.', 'Name', 'Father Name']):
        ws.write(0, c, title, head)
    for i, (d, l) in enumerate(columns):
        label = datetime.strptime(d, '%Y-%m-%d').strftime('%d-%b-%y') + '\n' + l.replace('Lecture ', 'L')
        ws.write(0, first + i, label, head)
    ws.write(0, c_present, 'Total Present', head)
    ws.write(0, c_total, 'Total Lectures', head)
    ws.write(0, c_pct, 'Attendance %', head)
    ws.set_row(0, 32)

    for r, (s, sm) in enumerate(zip(students, summary), start=1):
        ws.write_string(r, 0, s['rollno'], left)
        ws.write_string(r, 1, s['name'], left)
        ws.write_string(r, 2, s['fathername'] or '', left)
        for i, (d, l) in enumerate(columns):
            val = status.get((s['id'], d, l))
            ws.write_string(r, first + i, val, center) if val else ws.write_blank(r, first + i, None, center)

        rng = f'{xl_col_to_name(first)}{r + 1}:{xl_col_to_name(first + n - 1)}{r + 1}'
        pres_cell = f'{xl_col_to_name(c_present)}{r + 1}'
        tot_cell = f'{xl_col_to_name(c_total)}{r + 1}'
        ws.write_formula(r, c_present, f'=COUNTIF({rng},"P")', center, sm['present'])
        ws.write_formula(r, c_total, f'=COUNTIF({rng},"P")+COUNTIF({rng},"A")', center, sm['total'])
        ws.write_formula(r, c_pct, f'=IF({tot_cell}=0,0,{pres_cell}/{tot_cell})', percent_fmt,
                         sm['present'] / sm['total'] if sm['total'] else 0)

    if students and n:
        grid = f'{xl_col_to_name(first)}2:{xl_col_to_name(first + n - 1)}{len(students) + 1}'
        ws.conditional_format(grid, {'type': 'cell', 'criteria': '==', 'value': '"P"', 'format': green})
        ws.conditional_format(grid, {'type': 'cell', 'criteria': '==', 'value': '"A"', 'format': red})

    ws.set_column(0, 0, 12)
    ws.set_column(1, 2, 24)
    ws.set_column(first, first + n - 1, 10)
    ws.set_column(c_present, c_pct, 14)
    ws.freeze_panes(1, 3)
    wb.close()
    output.seek(0)
    return output


def excel_filename(class_name, sem, from_date, to_date, columns):
    def safe(s):
        return re.sub(r'[^A-Za-z0-9]+', '_', str(s)).strip('_') or 'Class'

    def short(d):
        return datetime.strptime(d, '%Y-%m-%d').strftime('%d%b')

    start, end = (from_date, to_date) if (from_date and to_date) else (columns[0][0], columns[-1][0])
    return f'{safe(class_name)}_Sem{safe(sem)}_{short(start)}-{short(end)}.xlsx'


@app.route('/api/attendance/statement', methods=['GET'])
def get_statement():
    class_name = request.args.get('className')
    sem = request.args.get('sem')
    from_date = request.args.get('fromDate')
    to_date = request.args.get('toDate')
    export_excel = request.args.get('export') == 'true'

    if not class_name or not sem:
        return error('SELECT CLASS AND SEMESTER.', 400)

    try:
        students, records = fetch_statement(class_name, sem, from_date, to_date)
        if not records:
            if export_excel:
                return error('NO ENTRIES FOUND.', 404)
            return jsonify([])

        columns, status, summary = build_grid(students, records)
        if not export_excel:
            return jsonify(summary)

        output = build_excel(students, columns, status, summary)
        return send_file(output, as_attachment=True, mimetype=EXCEL_MIME,
                         download_name=excel_filename(class_name, sem, from_date, to_date, columns))
    except Exception as e:
        return error(str(e))


if __name__ == '__main__':
    # Local testing ke liye. Deploy par: gunicorn app:app
    app.run(port=5000, debug=False)