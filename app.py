import sys
import os
import threading
import webbrowser
import configparser
from flask import Flask, render_template, request, jsonify, redirect, url_for, g
import sqlite3
from datetime import datetime

# PyInstaller support
if getattr(sys, 'frozen', False):
    BASE_DIR = sys._MEIPASS
    DATA_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    DATA_DIR = BASE_DIR

DB_PATH       = os.path.join(DATA_DIR, 'vk.db')
SETTINGS_PATH = os.path.join(DATA_DIR, 'settings.ini')
TEMPLATE_DIR  = os.path.join(BASE_DIR, 'templates')
STATIC_DIR    = os.path.join(BASE_DIR, 'static')

app = Flask(__name__, template_folder=TEMPLATE_DIR, static_folder=STATIC_DIR)

# ─── FIX #2: SECRET_KEY через файл у DATA_DIR ────────────────
def _load_or_create_secret_key():
    key_path = os.path.join(DATA_DIR, 'secret.key')
    if os.path.exists(key_path):
        with open(key_path, 'rb') as f:
            return f.read()
    key = os.urandom(32)
    with open(key_path, 'wb') as f:
        f.write(key)
    return key

app.secret_key = _load_or_create_secret_key()

# ─── FIX #3: Обмеження розміру запиту ────────────────────────
_max_mb = int(os.environ.get('MAX_UPLOAD_MB', 2))
app.config['MAX_CONTENT_LENGTH'] = _max_mb * 1024 * 1024

# ─── SETTINGS (INI) ──────────────────────────────────────────

def _migrate_settings_from_db(cfg):
    """Переносить старі налаштування з таблиці settings БД в ini (одноразово при першому запуску)."""
    try:
        conn = sqlite3.connect(DB_PATH, timeout=5)
        conn.row_factory = sqlite3.Row
        rows = conn.execute("SELECT key, value FROM settings").fetchall()
        conn.close()
        for row in rows:
            k, v = row['key'], row['value'] or ''
            if k in ('nazva_pidpryemstva', 'nazva_pidpryemstva_short', 'kod_edrpou'):
                cfg['company'][k] = v
    except Exception:
        pass  # БД ще не існує або таблиці немає — це нормально при першому запуску

def read_settings():
    """Читає settings.ini. Якщо файл не існує — створює з дефолтними значеннями."""
    cfg = configparser.ConfigParser()

    # Дефолтні значення
    cfg['company'] = {
        'nazva_pidpryemstva':       '',
        'nazva_pidpryemstva_short': '',
        'kod_edrpou':               '',
    }
    cfg['network'] = {
        'host': '127.0.0.1',
        'port': '5000',
    }

    if os.path.exists(SETTINGS_PATH):
        cfg.read(SETTINGS_PATH, encoding='utf-8')
    else:
        # Перший запуск — мігруємо з БД якщо є, потім зберігаємо ini
        _migrate_settings_from_db(cfg)
        _write_ini(cfg)

    return cfg

def _write_ini(cfg):
    """Записує configparser об'єкт у файл."""
    with open(SETTINGS_PATH, 'w', encoding='utf-8') as f:
        cfg.write(f)

def get_all_settings():
    """Повертає всі налаштування як плаский dict для шаблонів і API."""
    cfg = read_settings()
    result = {}
    for section in cfg.sections():
        for key, value in cfg[section].items():
            result[key] = value
    return result

# ─── DATABASE ────────────────────────────────────────────────
def get_db():
    if 'db' not in g:
        g.db = sqlite3.connect(DB_PATH, timeout=10)
        g.db.row_factory = sqlite3.Row
        g.db.execute("PRAGMA foreign_keys = ON")
        g.db.execute("PRAGMA journal_mode=WAL")
    return g.db

@app.teardown_appcontext
def close_db(error):
    db = g.pop('db', None)
    if db is not None:
        db.close()

# ─── FIX #1: Allowlist для безпечного ALTER TABLE ────────────
_ALLOWED_COL_TYPES = {'TEXT', 'INTEGER', 'REAL', 'BLOB', 'NUMERIC'}

def _safe_add_column(cursor, table, col_name, col_type):
    """Додає колонку тільки якщо назва і тип є в allowlist."""
    if not col_name.replace('_', '').isalnum():
        raise ValueError(f'Недопустима назва колонки: {col_name!r}')
    if col_type.upper() not in _ALLOWED_COL_TYPES:
        raise ValueError(f'Недопустимий тип колонки: {col_type!r}')
    cursor.execute(f'ALTER TABLE {table} ADD COLUMN {col_name} {col_type}')

# ─── FIX #5: init_db через app.app_context() ─────────────────
def init_db():
    with app.app_context():
        conn = get_db()
        c = conn.cursor()

        c.execute('''CREATE TABLE IF NOT EXISTS employees (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            -- Службові
            tabelny_nomer TEXT,
            data_zapovnennia TEXT DEFAULT (date('now')),
            nazva_pidpryemstva TEXT,
            kod_edrpou TEXT,
            vyd_roboty TEXT,
            -- Персональні
            prizvyshche TEXT NOT NULL,
            imia TEXT NOT NULL,
            po_batkovi TEXT,
            data_narodzhennia TEXT,
            gender TEXT,
            hromadianstvo TEXT DEFAULT 'Українець/ка',
            ipn TEXT,
            -- Паспорт
            pasport TEXT,
            pasport_vydanyi TEXT,
            data_vydachi_pasportu TEXT,
            id_karta TEXT,
            id_karta_diisna_do TEXT,
            -- Адреса
            adresa_faktychna TEXT,
            adresa_reiestratsiya TEXT,
            -- Поточна робота
            nazva_pidrozdilu TEXT,
            data_pryyomu TEXT,
            data_zvilnennia TEXT,
            prychyna_zvilnennia TEXT,
            -- Сім'я
            rodinny_stan TEXT,
            -- Пенсія
            pensiia TEXT,
            -- FIX #9: виправлено typo grupa_oblikyy -> grupa_obliku
            grupa_obliku TEXT,
            katehoriia_obliku TEXT,
            sklad TEXT,
            viiskove_zvannia TEXT,
            viiskova_spetsialnist TEXT,
            prydatnist TEXT,
            nazva_viiskkomatu_reiestr TEXT,
            nazva_viiskkomatu_faktych TEXT,
            spec_oblik TEXT,
            -- Додатково
            data_pryyomu_naek TEXT,
            dodatkovo TEXT,
            created_at TEXT DEFAULT (datetime('now')),
            updated_at TEXT DEFAULT (datetime('now'))
        )''')

        # Міграція: додаємо колонки яких може не бути в старій базі
        # FIX #1: використовуємо _safe_add_column замість f-string напряму
        new_columns = [
            ('nazva_pidpryemstva', 'TEXT'),
            ('kod_edrpou', 'TEXT'),
            ('vyd_roboty', 'TEXT'),
            ('gender', 'TEXT'),
            ('ipn', 'TEXT'),
            ('pasport', 'TEXT'),
            ('pasport_vydanyi', 'TEXT'),
            ('data_vydachi_pasportu', 'TEXT'),
            ('id_karta', 'TEXT'),
            ('id_karta_diisna_do', 'TEXT'),
            ('adresa_faktychna', 'TEXT'),
            ('adresa_reiestratsiya', 'TEXT'),
            ('nazva_pidrozdilu', 'TEXT'),
            ('data_pryyomu', 'TEXT'),
            ('data_zvilnennia', 'TEXT'),
            ('prychyna_zvilnennia', 'TEXT'),
            ('rodinny_stan', 'TEXT'),
            ('pensiia', 'TEXT'),
            # FIX #9: виправлено typo
            ('grupa_obliku', 'TEXT'),
            ('katehoriia_obliku', 'TEXT'),
            ('sklad', 'TEXT'),
            ('viiskove_zvannia', 'TEXT'),
            ('viiskova_spetsialnist', 'TEXT'),
            ('prydatnist', 'TEXT'),
            ('nazva_viiskkomatu_reiestr', 'TEXT'),
            ('nazva_viiskkomatu_faktych', 'TEXT'),
            ('spec_oblik', 'TEXT'),
            ('data_pryyomu_naek', 'TEXT'),
            ('dodatkovo', 'TEXT'),
        ]

        existing = [row[1] for row in c.execute("PRAGMA table_info(employees)").fetchall()]
        for col_name, col_type in new_columns:
            if col_name not in existing:
                _safe_add_column(c, 'employees', col_name, col_type)

        c.execute('''CREATE TABLE IF NOT EXISTS education (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            employee_id INTEGER NOT NULL REFERENCES employees(id) ON DELETE CASCADE,
            zaklad_nazva TEXT,
            dyplom_seriya TEXT,
            dyplom_nomer TEXT,
            rik_zakinch TEXT,
            spetsialnist TEXT,
            kvalifikatsiia TEXT,
            forma_navch TEXT,
            -- Після дипломна
            pislyadyplomna_typ TEXT,
            pislyadyplomna_zakl TEXT,
            pislyadyplomna_dypl TEXT,
            pislyadyplomna_rik TEXT,
            naukovyi_stupin TEXT
        )''')

        # Міграція education
        edu_columns = [
            ('pislyadyplomna_typ', 'TEXT'),
            ('pislyadyplomna_zakl', 'TEXT'),
            ('pislyadyplomna_dypl', 'TEXT'),
            ('pislyadyplomna_rik', 'TEXT'),
            ('naukovyi_stupin', 'TEXT'),
        ]
        existing_edu = [row[1] for row in c.execute("PRAGMA table_info(education)").fetchall()]
        for col_name, col_type in edu_columns:
            if col_name not in existing_edu:
                _safe_add_column(c, 'education', col_name, col_type)

        c.execute('''CREATE TABLE IF NOT EXISTS family (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            employee_id INTEGER NOT NULL REFERENCES employees(id) ON DELETE CASCADE,
            relationship TEXT,
            full_name TEXT,
            birth_year TEXT
        )''')

        c.execute('''CREATE TABLE IF NOT EXISTS work_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            employee_id INTEGER NOT NULL REFERENCES employees(id) ON DELETE CASCADE,
            start_date TEXT,
            end_date TEXT,
            nazva_pidrozdilu TEXT,
            posada TEXT,
            nakaz TEXT,
            prychyna_zvilnennia TEXT
        )''')

        # Розділ ІІІ — Призначення і переведення
        c.execute('''CREATE TABLE IF NOT EXISTS appointments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            employee_id INTEGER NOT NULL REFERENCES employees(id) ON DELETE CASCADE,
            data TEXT,
            nazva_pidrozdilu TEXT,
            profesiya_posada TEXT,
            nakaz_nomer TEXT
        )''')

        # Розділ IV — Відпустки
        c.execute('''CREATE TABLE IF NOT EXISTS vacations (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            employee_id INTEGER NOT NULL REFERENCES employees(id) ON DELETE CASCADE,
            typ TEXT,
            za_yakyi_period TEXT,
            start_date TEXT,
            end_date TEXT,
            calendar_days INTEGER,
            days INTEGER,
            nakaz TEXT
        )''')

        # Міграція vacations
        vac_columns = [('za_yakyi_period', 'TEXT'), ('calendar_days', 'INTEGER')]
        existing_vac = [row[1] for row in c.execute("PRAGMA table_info(vacations)").fetchall()]
        for col_name, col_type in vac_columns:
            if col_name not in existing_vac:
                _safe_add_column(c, 'vacations', col_name, col_type)

        # Стара таблиця settings залишається для зворотної сумісності (не видаляємо)
        # Нові налаштування зберігаються в settings.ini
        c.execute('''CREATE TABLE IF NOT EXISTS settings (
            key TEXT PRIMARY KEY,
            value TEXT
        )''')

        # Індекси для швидкого пошуку
        c.execute('CREATE INDEX IF NOT EXISTS idx_prizvyshche ON employees(prizvyshche)')
        c.execute('CREATE INDEX IF NOT EXISTS idx_tabelny ON employees(tabelny_nomer)')

        # FIX #12: індекси на employee_id в усіх дочірніх таблицях
        c.execute('CREATE INDEX IF NOT EXISTS idx_education_emp ON education(employee_id)')
        c.execute('CREATE INDEX IF NOT EXISTS idx_family_emp ON family(employee_id)')
        c.execute('CREATE INDEX IF NOT EXISTS idx_work_history_emp ON work_history(employee_id)')
        c.execute('CREATE INDEX IF NOT EXISTS idx_appointments_emp ON appointments(employee_id)')
        c.execute('CREATE INDEX IF NOT EXISTS idx_vacations_emp ON vacations(employee_id)')

        conn.commit()

# ─── HELPERS ─────────────────────────────────────────────────
def row_to_dict(row):
    return dict(row) if row else None

def rows_to_list(rows):
    return [dict(r) for r in rows]

def validate_employee(data):
    """Перевіряє вхідні дані картки. Повертає список помилок або порожній список."""
    if not isinstance(data, dict):
        return ['Невалідний формат даних (очікується JSON-об\'єкт)']
    errors = []
    prizvyshche = (data.get('prizvyshche') or '').strip()
    imia        = (data.get('imia') or '').strip()
    if not prizvyshche:
        errors.append('Прізвище є обов\'язковим полем')
    if not imia:
        errors.append('Ім\'я є обов\'язковим полем')
    return errors

# ─── ROUTES: PAGES ───────────────────────────────────────────
@app.route('/')
def index():
    return redirect(url_for('employees_list'))

@app.route('/employees')
def employees_list():
    return render_template('employees.html')

@app.route('/settings')
def settings_page():
    return render_template('settings.html', settings=get_all_settings())

@app.route('/employees/new')
def employee_new():
    return render_template('employee_form.html', employee=None, mode='new', settings=get_all_settings())

@app.route('/employees/<int:emp_id>')
def employee_view(emp_id):
    conn = get_db()
    emp = row_to_dict(conn.execute('SELECT * FROM employees WHERE id=?', (emp_id,)).fetchone())
    if not emp:
        return redirect(url_for('employees_list'))
    return render_template('employee_form.html', employee=emp, mode='view', settings=get_all_settings())

@app.route('/employees/<int:emp_id>/edit')
def employee_edit(emp_id):
    conn = get_db()
    emp = row_to_dict(conn.execute('SELECT * FROM employees WHERE id=?', (emp_id,)).fetchone())
    if not emp:
        return redirect(url_for('employees_list'))
    return render_template('employee_form.html', employee=emp, mode='edit', settings=get_all_settings())

# ─── ROUTES: API ─────────────────────────────────────────────
@app.route('/api/employees')
def api_employees():
    # FIX #4: обрізаємо пошуковий рядок до 100 символів
    q = request.args.get('q', '').strip()[:100]
    conn = get_db()

    base_select = '''
        SELECT e.id, e.tabelny_nomer, e.prizvyshche, e.imia, e.po_batkovi,
               a.data AS data_pryyomu,
               a.nazva_pidrozdilu AS nazva_pidrozdilu,
               a.profesiya_posada AS posada
        FROM employees e
        LEFT JOIN (
            SELECT employee_id,
                   data,
                   nazva_pidrozdilu,
                   profesiya_posada
            FROM appointments
            WHERE id IN (
                SELECT MAX(id) FROM appointments GROUP BY employee_id
            )
        ) a ON a.employee_id = e.id
    '''

    if q:
        like = f'%{q}%'
        rows = conn.execute(base_select + '''
            WHERE e.prizvyshche LIKE ? OR e.imia LIKE ? OR e.po_batkovi LIKE ?
               OR e.tabelny_nomer LIKE ? OR e.ipn LIKE ?
               OR a.nazva_pidrozdilu LIKE ? OR a.profesiya_posada LIKE ?
            ORDER BY e.prizvyshche, e.imia
        ''', (like, like, like, like, like, like, like)).fetchall()
    else:
        rows = conn.execute(base_select + '''
            ORDER BY e.prizvyshche, e.imia
        ''').fetchall()

    return jsonify(rows_to_list(rows))

@app.route('/api/employees', methods=['POST'])
def api_create_employee():
    data = request.json
    errors = validate_employee(data)
    if errors:
        return jsonify({'error': 'validation', 'messages': errors}), 400

    now  = datetime.now().isoformat()
    conn = get_db()
    c    = conn.cursor()

    # Перевірка унікальності табельного номера
    tabelny = data.get('tabelny_nomer', '').strip() if data.get('tabelny_nomer') else None
    if tabelny:
        existing = conn.execute(
            'SELECT id, prizvyshche, imia, po_batkovi FROM employees WHERE TRIM(tabelny_nomer)=?',
            (tabelny,)
        ).fetchone()
        if existing:
            full_name = f"{existing['prizvyshche']} {existing['imia']} {existing['po_batkovi'] or ''}".strip()
            return jsonify({'error': 'duplicate_tabelny', 'message': f'Табельний номер {tabelny} вже існує у працівника: {full_name}', 'employee_id': existing['id']}), 409

    # FIX #9: виправлено назву поля grupa_obliku
    c.execute('''INSERT INTO employees (
        tabelny_nomer, data_zapovnennia, nazva_pidpryemstva, kod_edrpou, vyd_roboty,
        prizvyshche, imia, po_batkovi, data_narodzhennia, gender, hromadianstvo, ipn,
        pasport, pasport_vydanyi, data_vydachi_pasportu, id_karta, id_karta_diisna_do,
        adresa_faktychna, adresa_reiestratsiya,
        nazva_pidrozdilu, data_pryyomu,
        data_zvilnennia, prychyna_zvilnennia,
        rodinny_stan, pensiia,
        grupa_obliku, katehoriia_obliku, sklad, viiskove_zvannia,
        viiskova_spetsialnist, prydatnist,
        nazva_viiskkomatu_reiestr, nazva_viiskkomatu_faktych, spec_oblik,
        data_pryyomu_naek, dodatkovo, updated_at
    ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''', (
        data.get('tabelny_nomer'), data.get('data_zapovnennia'), data.get('nazva_pidpryemstva'),
        data.get('kod_edrpou'), data.get('vyd_roboty'),
        data['prizvyshche'], data['imia'], data.get('po_batkovi'),
        data.get('data_narodzhennia'), data.get('gender'),
        data.get('hromadianstvo', 'Українець/ка'), data.get('ipn'),
        data.get('pasport'), data.get('pasport_vydanyi'), data.get('data_vydachi_pasportu'),
        data.get('id_karta'), data.get('id_karta_diisna_do'),
        data.get('adresa_faktychna'), data.get('adresa_reiestratsiya'),
        data.get('nazva_pidrozdilu'), data.get('data_pryyomu'),
        data.get('data_zvilnennia'), data.get('prychyna_zvilnennia'),
        data.get('rodinny_stan'), data.get('pensiia'),
        data.get('grupa_obliku'), data.get('katehoriia_obliku'), data.get('sklad'),
        data.get('viiskove_zvannia'), data.get('viiskova_spetsialnist'), data.get('prydatnist'),
        data.get('nazva_viiskkomatu_reiestr'), data.get('nazva_viiskkomatu_faktych'),
        data.get('spec_oblik'), data.get('data_pryyomu_naek'), data.get('dodatkovo'), now
    ))

    emp_id = c.lastrowid
    _save_education(c, emp_id, data.get('education', []))
    _save_family(c, emp_id, data.get('family', []))
    _save_work_history(c, emp_id, data.get('work_history', []))
    _save_appointments(c, emp_id, data.get('appointments', []))
    _save_vacations(c, emp_id, data.get('vacations', []))
    conn.commit()

    return jsonify({'id': emp_id, 'status': 'created'}), 201

@app.route('/api/employees/<int:emp_id>', methods=['GET'])
def api_get_employee(emp_id):
    conn = get_db()
    # TODO: замінити SELECT * на явні колонки коли додамо фото (issue #6)
    emp = row_to_dict(conn.execute('SELECT * FROM employees WHERE id=?', (emp_id,)).fetchone())
    if not emp:
        return jsonify({'error': 'Not found'}), 404
    emp['education']    = rows_to_list(conn.execute('SELECT * FROM education WHERE employee_id=?', (emp_id,)).fetchall())
    emp['family']       = rows_to_list(conn.execute('SELECT * FROM family WHERE employee_id=?', (emp_id,)).fetchall())
    emp['work_history'] = rows_to_list(conn.execute('SELECT * FROM work_history WHERE employee_id=?', (emp_id,)).fetchall())
    emp['appointments'] = rows_to_list(conn.execute('SELECT * FROM appointments WHERE employee_id=? ORDER BY data', (emp_id,)).fetchall())
    emp['vacations']    = rows_to_list(conn.execute('SELECT * FROM vacations WHERE employee_id=?', (emp_id,)).fetchall())
    return jsonify(emp)

@app.route('/api/employees/<int:emp_id>', methods=['PUT'])
def api_update_employee(emp_id):
    data = request.json
    errors = validate_employee(data)
    if errors:
        return jsonify({'error': 'validation', 'messages': errors}), 400

    now  = datetime.now().isoformat()
    conn = get_db()
    c    = conn.cursor()

    # Перевірка унікальності табельного номера (ігноруємо поточного працівника)
    tabelny = data.get('tabelny_nomer', '').strip() if data.get('tabelny_nomer') else None
    if tabelny:
        existing = conn.execute(
            'SELECT id, prizvyshche, imia, po_batkovi FROM employees WHERE TRIM(tabelny_nomer)=? AND id!=?',
            (tabelny, emp_id)
        ).fetchone()
        if existing:
            full_name = f"{existing['prizvyshche']} {existing['imia']} {existing['po_batkovi'] or ''}".strip()
            return jsonify({'error': 'duplicate_tabelny', 'message': f'Табельний номер {tabelny} вже існує у працівника: {full_name}', 'employee_id': existing['id']}), 409

    # FIX #6: транзакційний захист через with conn — при помилці автоматичний rollback
    try:
        with conn:
            # FIX #9: виправлено назву поля grupa_obliku
            c.execute('''UPDATE employees SET
                tabelny_nomer=?, data_zapovnennia=?, nazva_pidpryemstva=?, kod_edrpou=?, vyd_roboty=?,
                prizvyshche=?, imia=?, po_batkovi=?, data_narodzhennia=?, gender=?, hromadianstvo=?, ipn=?,
                pasport=?, pasport_vydanyi=?, data_vydachi_pasportu=?, id_karta=?, id_karta_diisna_do=?,
                adresa_faktychna=?, adresa_reiestratsiya=?,
                nazva_pidrozdilu=?, data_pryyomu=?,
                data_zvilnennia=?, prychyna_zvilnennia=?,
                rodinny_stan=?, pensiia=?,
                grupa_obliku=?, katehoriia_obliku=?, sklad=?, viiskove_zvannia=?,
                viiskova_spetsialnist=?, prydatnist=?,
                nazva_viiskkomatu_reiestr=?, nazva_viiskkomatu_faktych=?, spec_oblik=?,
                data_pryyomu_naek=?, dodatkovo=?, updated_at=?
                WHERE id=?''', (
                data.get('tabelny_nomer'), data.get('data_zapovnennia'), data.get('nazva_pidpryemstva'),
                data.get('kod_edrpou'), data.get('vyd_roboty'),
                data['prizvyshche'], data['imia'], data.get('po_batkovi'),
                data.get('data_narodzhennia'), data.get('gender'),
                data.get('hromadianstvo', 'Українець/ка'), data.get('ipn'),
                data.get('pasport'), data.get('pasport_vydanyi'), data.get('data_vydachi_pasportu'),
                data.get('id_karta'), data.get('id_karta_diisna_do'),
                data.get('adresa_faktychna'), data.get('adresa_reiestratsiya'),
                data.get('nazva_pidrozdilu'), data.get('data_pryyomu'),
                data.get('data_zvilnennia'), data.get('prychyna_zvilnennia'),
                data.get('rodinny_stan'), data.get('pensiia'),
                data.get('grupa_obliku'), data.get('katehoriia_obliku'), data.get('sklad'),
                data.get('viiskove_zvannia'), data.get('viiskova_spetsialnist'), data.get('prydatnist'),
                data.get('nazva_viiskkomatu_reiestr'), data.get('nazva_viiskkomatu_faktych'),
                data.get('spec_oblik'), data.get('data_pryyomu_naek'), data.get('dodatkovo'), now,
                emp_id
            ))

            c.execute('DELETE FROM education WHERE employee_id=?',    (emp_id,))
            c.execute('DELETE FROM family WHERE employee_id=?',       (emp_id,))
            c.execute('DELETE FROM work_history WHERE employee_id=?', (emp_id,))
            c.execute('DELETE FROM appointments WHERE employee_id=?', (emp_id,))
            c.execute('DELETE FROM vacations WHERE employee_id=?',    (emp_id,))

            _save_education(c, emp_id, data.get('education', []))
            _save_family(c, emp_id, data.get('family', []))
            _save_work_history(c, emp_id, data.get('work_history', []))
            _save_appointments(c, emp_id, data.get('appointments', []))
            _save_vacations(c, emp_id, data.get('vacations', []))

    except Exception as e:
        app.logger.error(f'Помилка при оновленні працівника {emp_id}: {e}', exc_info=True)
        return jsonify({'error': 'Помилка збереження даних', 'detail': str(e)}), 500

    return jsonify({'status': 'updated'})

@app.route('/api/employees/<int:emp_id>', methods=['DELETE'])
def api_delete_employee(emp_id):
    conn   = get_db()
    result = conn.execute('DELETE FROM employees WHERE id=?', (emp_id,))
    conn.commit()
    # FIX #10: перевіряємо чи запис існував через rowcount
    if result.rowcount == 0:
        return jsonify({'error': 'Not found'}), 404
    return jsonify({'status': 'deleted'})

@app.route('/api/employees/<int:emp_id>/prev')
def api_prev_employee(emp_id):
    # FIX #8: залишаємо навігацію за id (прийнято рішення)
    conn = get_db()
    row  = conn.execute('SELECT id FROM employees WHERE id < ? ORDER BY id DESC LIMIT 1', (emp_id,)).fetchone()
    return jsonify({'id': row['id'] if row else None})

@app.route('/api/employees/<int:emp_id>/next')
def api_next_employee(emp_id):
    conn = get_db()
    row  = conn.execute('SELECT id FROM employees WHERE id > ? ORDER BY id ASC LIMIT 1', (emp_id,)).fetchone()
    return jsonify({'id': row['id'] if row else None})

@app.route('/api/stats')
def api_stats():
    conn  = get_db()
    total = conn.execute('SELECT COUNT(*) FROM employees').fetchone()[0]
    return jsonify({'total': total})

@app.route('/api/settings', methods=['GET'])
def api_get_settings():
    return jsonify(get_all_settings())

@app.route('/api/settings', methods=['POST'])
def api_save_settings():
    data = request.json
    if not isinstance(data, dict):
        return jsonify({'error': 'invalid data'}), 400

    # Маппінг дозволених полів → секції ini
    allowed = {
        'nazva_pidpryemstva':       'company',
        'nazva_pidpryemstva_short': 'company',
        'kod_edrpou':               'company',
        'host':                     'network',
        'port':                     'network',
    }

    cfg     = read_settings()
    changed = False

    for key, section in allowed.items():
        if key not in data:
            continue
        val = (data[key] or '').strip()

        # Валідація порту
        if key == 'port':
            try:
                port_int = int(val)
                if not (1024 <= port_int <= 65535):
                    return jsonify({'error': 'Порт має бути від 1024 до 65535'}), 400
            except ValueError:
                return jsonify({'error': 'Порт має бути числом'}), 400

        # Валідація host
        if key == 'host' and val not in ('127.0.0.1', '0.0.0.0'):
            return jsonify({'error': 'Невірне значення host'}), 400

        cfg[section][key] = val
        changed = True

    if changed:
        _write_ini(cfg)

    return jsonify({'status': 'ok'})

# ─── HELPERS: RELATED DATA ───────────────────────────────────
def _save_education(c, emp_id, items):
    for e in items:
        c.execute('''INSERT INTO education (
            employee_id, zaklad_nazva, dyplom_seriya, dyplom_nomer, rik_zakinch,
            spetsialnist, kvalifikatsiia, forma_navch,
            pislyadyplomna_typ, pislyadyplomna_zakl, pislyadyplomna_dypl,
            pislyadyplomna_rik, naukovyi_stupin)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)''', (
            emp_id,
            e.get('zaklad_nazva'), e.get('dyplom_seriya'), e.get('dyplom_nomer'), e.get('rik_zakinch'),
            e.get('spetsialnist'), e.get('kvalifikatsiia'), e.get('forma_navch'),
            e.get('pislyadyplomna_typ'), e.get('pislyadyplomna_zakl'), e.get('pislyadyplomna_dypl'),
            e.get('pislyadyplomna_rik'), e.get('naukovyi_stupin')
        ))

def _save_family(c, emp_id, items):
    for f in items:
        c.execute('INSERT INTO family (employee_id, relationship, full_name, birth_year) VALUES (?,?,?,?)',
                  (emp_id, f.get('relationship'), f.get('full_name'), f.get('birth_year')))

def _save_work_history(c, emp_id, items):
    for w in items:
        c.execute('''INSERT INTO work_history (employee_id, start_date, end_date,
            nazva_pidrozdilu, posada, nakaz, prychyna_zvilnennia)
            VALUES (?,?,?,?,?,?,?)''', (
            emp_id, w.get('start_date'), w.get('end_date'),
            w.get('nazva_pidrozdilu'), w.get('posada'),
            w.get('nakaz'), w.get('prychyna_zvilnennia')
        ))

def _save_appointments(c, emp_id, items):
    for a in items:
        c.execute('''INSERT INTO appointments (employee_id, data, nazva_pidrozdilu, profesiya_posada, nakaz_nomer)
            VALUES (?,?,?,?,?)''', (
            emp_id, a.get('data'), a.get('nazva_pidrozdilu'),
            a.get('profesiya_posada'), a.get('nakaz_nomer')
        ))

def _save_vacations(c, emp_id, items):
    for v in items:
        c.execute('''INSERT INTO vacations (employee_id, typ, za_yakyi_period, start_date, end_date, calendar_days, days, nakaz)
            VALUES (?,?,?,?,?,?,?,?)''', (
            emp_id, v.get('typ'), v.get('za_yakyi_period'),
            v.get('start_date'), v.get('end_date'), v.get('calendar_days'), v.get('days'), v.get('nakaz')
        ))

# ─── ERROR HANDLING ──────────────────────────────────────────
@app.errorhandler(Exception)
def handle_exception(e):
    app.logger.error(f'Unhandled exception: {e}', exc_info=True)
    return jsonify({'error': 'Внутрішня помилка сервера', 'detail': str(e)}), 500

# ─── STARTUP ─────────────────────────────────────────────────
def setup_logging():
    import logging
    log_path  = os.path.join(DATA_DIR, 'vk_errors.log')
    handler   = logging.FileHandler(log_path, encoding='utf-8')
    handler.setLevel(logging.ERROR)
    formatter = logging.Formatter('%(asctime)s [%(levelname)s] %(message)s')
    handler.setFormatter(formatter)
    app.logger.addHandler(handler)
    app.logger.setLevel(logging.ERROR)

def open_browser(port):
    import time
    time.sleep(1)
    webbrowser.open(f'http://127.0.0.1:{port}')

if __name__ == '__main__':
    setup_logging()
    init_db()

    # Читаємо мережеві налаштування з settings.ini
    # Змінна середовища HOST має пріоритет над ini (для зворотної сумісності з v1.1.0)
    cfg = read_settings()

    env_host = os.environ.get('HOST')
    if env_host in ('127.0.0.1', 'localhost', '0.0.0.0'):
        host = env_host
    else:
        host = cfg.get('network', 'host', fallback='127.0.0.1')
        if host not in ('127.0.0.1', '0.0.0.0'):
            host = '127.0.0.1'

    try:
        port = int(cfg.get('network', 'port', fallback='5000'))
        if not (1024 <= port <= 65535):
            port = 5000
    except ValueError:
        port = 5000

    # FIX #14: браузер відкривається тільки при локальному запуску
    if host in ('127.0.0.1', 'localhost'):
        threading.Thread(target=open_browser, args=(port,), daemon=True).start()

    app.run(debug=False, host=host, port=port, use_reloader=False)
