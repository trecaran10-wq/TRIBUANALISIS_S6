from flask import Flask, render_template, request, redirect, url_for, flash, send_file, abort, session, g
import sqlite3, os, io, csv, json, re, unicodedata
from werkzeug.security import generate_password_hash, check_password_hash
from datetime import datetime
from openpyxl import Workbook, load_workbook
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.environ.get('DATABASE_PATH', os.path.join(BASE_DIR, "tribuanalisis.db"))
app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev-only-cambiar-esta-clave")
app.config.update(SESSION_COOKIE_HTTPONLY=True, SESSION_COOKIE_SAMESITE='Lax',
                  SESSION_COOKIE_SECURE=os.environ.get('SESSION_COOKIE_SECURE', '').lower() in {'1', 'true', 'yes'})

CATEGORIES = {
    "cuarta": "Renta de cuarta categoría",
    "quinta": "Renta de quinta categoría",
    "creador": "Creador de contenido",
    "plataforma": "Economía de plataformas",
}

def db():
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.execute("PRAGMA busy_timeout=30000")
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    with db() as conn:
        conn.execute("""CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            full_name TEXT NOT NULL,
            email TEXT NOT NULL UNIQUE COLLATE NOCASE,
            password_hash TEXT NOT NULL,
            created_at TEXT NOT NULL
        )""")
        conn.execute("""CREATE TABLE IF NOT EXISTS cases (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER,
            code TEXT UNIQUE NOT NULL,
            person TEXT NOT NULL,
            category TEXT NOT NULL,
            platform TEXT DEFAULT '',
            income REAL NOT NULL DEFAULT 0,
            expenses REAL NOT NULL DEFAULT 0,
            withholding REAL NOT NULL DEFAULT 0,
            country TEXT DEFAULT 'Perú',
            notes TEXT DEFAULT '',
            tax_base REAL NOT NULL DEFAULT 0,
            estimated_tax REAL NOT NULL DEFAULT 0,
            created_at TEXT NOT NULL,
            tax_year INTEGER NOT NULL DEFAULT 2026,
            payments REAL NOT NULL DEFAULT 0,
            additional_deductions REAL NOT NULL DEFAULT 0,
            itf_deduction REAL NOT NULL DEFAULT 0,
            donations REAL NOT NULL DEFAULT 0,
            director_role INTEGER NOT NULL DEFAULT 0,
            other_fourth REAL NOT NULL DEFAULT 0,
            other_fifth REAL NOT NULL DEFAULT 0,
            income_frequency TEXT NOT NULL DEFAULT 'annual',
            payments_frequency TEXT NOT NULL DEFAULT 'annual',
            calculation_json TEXT NOT NULL DEFAULT '{}'
        )""")
        existing = {row[1] for row in conn.execute("PRAGMA table_info(cases)").fetchall()}
        if 'user_id' not in existing:
            conn.execute("ALTER TABLE cases ADD COLUMN user_id INTEGER")
        additions = {
            'tax_year': "INTEGER NOT NULL DEFAULT 2026", 'payments': "REAL NOT NULL DEFAULT 0",
            'additional_deductions': "REAL NOT NULL DEFAULT 0", 'itf_deduction': "REAL NOT NULL DEFAULT 0", 'donations': "REAL NOT NULL DEFAULT 0", 'director_role': "INTEGER NOT NULL DEFAULT 0", 'other_fourth': "REAL NOT NULL DEFAULT 0",
            'other_fifth': "REAL NOT NULL DEFAULT 0", 'income_frequency': "TEXT NOT NULL DEFAULT 'annual'",
            'payments_frequency': "TEXT NOT NULL DEFAULT 'annual'", 'calculation_json': "TEXT NOT NULL DEFAULT '{}'"
        }
        for col, definition in additions.items():
            if col not in existing:
                conn.execute(f"ALTER TABLE cases ADD COLUMN {col} {definition}")

@app.before_request
def load_logged_in_user():
    endpoint = request.endpoint or ''
    if endpoint in {'login', 'register', 'static'} or endpoint.startswith('static'):
        g.user = None
        return
    user_id = session.get('user_id')
    if not user_id:
        return redirect(url_for('login', next=request.path))
    with db() as conn:
        g.user = conn.execute("SELECT id, full_name, email FROM users WHERE id=?", (user_id,)).fetchone()
    if g.user is None:
        session.clear()
        return redirect(url_for('login'))

@app.context_processor
def inject_current_user():
    return {'current_user': getattr(g, 'user', None)}

@app.route('/ingresar', methods=['GET', 'POST'])
@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        email = request.form.get('email', '').strip().lower()
        password = request.form.get('password', '')
        with db() as conn:
            user = conn.execute("SELECT * FROM users WHERE email=?", (email,)).fetchone()
        if user and check_password_hash(user['password_hash'], password):
            session.clear()
            session['user_id'] = user['id']
            flash('Sesión iniciada. Tus casos pertenecen a tu cuenta.', 'success')
            return redirect(url_for('index'))
        flash('Correo o contraseña incorrectos.', 'error')
    return render_template('auth.html', mode='login')

@app.route('/crear-cuenta', methods=['GET', 'POST'])
@app.route('/registro', methods=['GET', 'POST'])
def register():
    if request.method == 'POST':
        full_name = request.form.get('full_name', '').strip()
        email = request.form.get('email', '').strip().lower()
        password = request.form.get('password', '')
        if not full_name or not re.fullmatch(r'[^@\s]+@[^@\s]+\.[^@\s]+', email):
            flash('Completa tu nombre y un correo electrónico válido.', 'error')
        elif len(password) < 8:
            flash('La contraseña debe tener al menos 8 caracteres.', 'error')
        else:
            try:
                with db() as conn:
                    cur = conn.execute("INSERT INTO users(full_name,email,password_hash,created_at) VALUES(?,?,?,?)",
                        (full_name, email, generate_password_hash(password), datetime.now().strftime('%d/%m/%Y %H:%M')))
                    new_id = cur.lastrowid
                session.clear(); session['user_id'] = new_id
                flash('Cuenta creada correctamente. Ya puedes guardar tus análisis.', 'success')
                return redirect(url_for('index'))
            except sqlite3.IntegrityError:
                flash('Ese correo ya tiene una cuenta. Inicia sesión o usa otro correo.', 'error')
    return render_template('auth.html', mode='register')

@app.route('/cerrar-sesion', methods=['GET', 'POST'])
@app.route('/logout', methods=['GET', 'POST'])
def logout():
    session.clear()
    flash('Has cerrado sesión correctamente.', 'success')
    return redirect(url_for('login'))

def money(value):
    return f"S/ {float(value or 0):,.2f}"

def progressive_tax(taxable, uit):
    """Escala progresiva acumulativa del artículo 53 de la LIR."""
    remaining = max(float(taxable or 0), 0)
    bands = [(5 * uit, .08), (15 * uit, .14), (15 * uit, .17), (10 * uit, .20), (float("inf"), .30)]
    tax, lower = 0.0, 0.0
    for width, rate in bands:
        upper = lower + width if width != float("inf") else float("inf")
        portion = max(min(remaining, upper) - lower, 0)
        tax += portion * rate
        if remaining <= upper:
            break
        lower = upper
    return round(tax, 2)

def calculate(category, income, expenses=0, withholding=0, *, year=2026, payments=0,
              additional_deductions=0, itf_deduction=0, donations=0, other_fourth=0, other_fifth=0, income_frequency='annual',
              payments_frequency='annual', director=False):
    """Annual employment-income model; digital creator/platform cases are not forced into a flat rate."""
    income = max(float(income or 0), 0)
    payments = max(float(payments or 0), 0)
    withholding = max(float(withholding or 0), 0)
    additional_deductions = max(float(additional_deductions or 0), 0)
    itf_deduction = max(float(itf_deduction or 0), 0)
    donations = max(float(donations or 0), 0)
    other_fourth, other_fifth = max(float(other_fourth or 0), 0), max(float(other_fifth or 0), 0)
    income_annual = income * 12 if income_frequency == 'monthly' else income
    payments_annual = payments * 12 if payments_frequency == 'monthly' else payments
    uit = 5500 if int(year) == 2026 else (5350 if int(year) == 2025 else 5500)
    if category in ('creador', 'plataforma'):
        return {
            'annual_income': round(income_annual, 2), 'base': 0, 'tax': 0, 'tax_due': 0,
            'payments_annual': round(payments_annual, 2), 'withholding': round(withholding, 2),
            'uit': uit, 'status': 'requires_review', 'taxable_base': 0,
            'explanation': 'No se aplica automáticamente el motor de rentas de trabajo. SUNAT considera rentas de tercera categoría determinados ingresos de influencers por publicidad, monetización y otras actividades empresariales; clasifica cada flujo y revisa el Informe N.º 044-2022-SUNAT/7T0000.',
            'steps': ['Separar cada fuente de ingreso y su contraprestación.', 'Revisar contratos, liquidaciones, pagador, residencia y fuente.', 'Determinar si corresponde tercera categoría y el régimen tributario antes de calcular el impuesto.'],
            'credits': round(payments_annual + withholding, 2), 'balance': None, 'deduction_fourth': 0,
            'gross_work_income': round(income_annual, 2), 'additional_deductions_applied': 0, 'itf_applied': 0, 'donations_applied': 0,
            'tax_year': int(year), 'legal_basis': ['Informe N.º 044-2022-SUNAT/7T0000', 'Artículos 28 y 33 de la Ley del Impuesto a la Renta, según corresponda']
        }
    fourth = income_annual if category == 'cuarta' else other_fourth
    fifth = income_annual if category == 'quinta' else other_fifth
    fourth_deduction = 0 if director else min(fourth * .20, 24 * uit)
    if fourth <= 0:
        fourth_reason = 'No se registró ingreso de cuarta categoría; no se aplica deducción del 20% en este caso.'
    elif director:
        fourth_reason = 'No se aplicó la deducción del 20% porque se indicó una función excluida por el artículo 45 de la LIR; validar la naturaleza exacta de la función.'
    elif fourth * .20 > 24 * uit:
        fourth_reason = f'El 20% de la renta bruta supera el tope de 24 UIT (S/ {24*uit:,.2f}); por eso la deducción se limita a S/ {24*uit:,.2f}.'
    else:
        fourth_reason = f'Se aplicó el 20% de la renta bruta (S/ {fourth_deduction:,.2f}); no se supera el límite de 24 UIT (S/ {24*uit:,.2f}).'
    net_fourth = max(fourth - fourth_deduction, 0)
    work_net_before_7uit = net_fourth + fifth
    base_after_7uit = max(work_net_before_7uit - 7 * uit, 0)
    # Additional deductions are accepted only when user has verified that they meet the statutory conditions.
    additional_cap = 3 * uit
    additional_applied = min(additional_deductions, additional_cap, base_after_7uit)
    base_after_personal = max(base_after_7uit - additional_applied, 0)
    itf_applied = min(itf_deduction, base_after_personal)
    base_before_donation = max(base_after_personal - itf_applied, 0)
    donations_applied = min(donations, base_before_donation * .10)
    taxable = max(base_before_donation - donations_applied, 0)
    tax = progressive_tax(taxable, uit)
    credits = payments_annual + withholding
    balance = round(tax - credits, 2)
    fourth_step = (f'Renta bruta anual de cuarta: S/ {fourth:,.2f}. Deducción legal del 20%: S/ {fourth_deduction:,.2f}. {fourth_reason}' if fourth > 0 else 'No se registró renta de cuarta categoría; por ello no se aplica la deducción del 20% en esta ficha.')
    tax_bands = []
    band_defs = [(5 * uit, .08, 'Hasta 5 UIT'), (15 * uit, .14, 'Por el exceso de 5 UIT y hasta 20 UIT'), (15 * uit, .17, 'Por el exceso de 20 UIT y hasta 35 UIT'), (10 * uit, .20, 'Por el exceso de 35 UIT y hasta 45 UIT'), (float('inf'), .30, 'Por el exceso de 45 UIT')]
    lower = 0.0
    remaining_taxable = taxable
    for width, rate, label in band_defs:
        upper = lower + width if width != float('inf') else float('inf')
        portion = max(min(taxable, upper) - lower, 0)
        tax_bands.append({'label': label, 'rate': rate, 'base': round(portion, 2), 'tax': round(portion * rate, 2)})
        if taxable <= upper:
            # Keep subsequent bands visible with zero amounts for transparency.
            lower = upper
            continue
        lower = upper
    schedule = [
        {'label': 'Renta bruta de cuarta categoría', 'amount': round(fourth, 2), 'kind': 'normal'},
        {'label': 'Deducción del 20% de cuarta categoría (límite 24 UIT)', 'amount': -round(fourth_deduction, 2), 'kind': 'deduction'},
        {'label': 'Renta neta de cuarta categoría', 'amount': round(net_fourth, 2), 'kind': 'subtotal'},
        {'label': 'Renta bruta de quinta categoría', 'amount': round(fifth, 2), 'kind': 'normal'},
        {'label': 'Total de rentas netas de trabajo antes de 7 UIT', 'amount': round(work_net_before_7uit, 2), 'kind': 'subtotal'},
        {'label': 'Deducción fija de 7 UIT', 'amount': -round(min(7 * uit, work_net_before_7uit), 2), 'kind': 'deduction'},
        {'label': 'Deducciones adicionales sustentadas (máximo conjunto 3 UIT)', 'amount': -round(additional_applied, 2), 'kind': 'deduction'},
        {'label': 'ITF deducible ingresado y admitido', 'amount': -round(itf_applied, 2), 'kind': 'deduction'},
        {'label': 'Donaciones deducibles admitidas', 'amount': -round(donations_applied, 2), 'kind': 'deduction'},
        {'label': 'Renta neta imponible de trabajo', 'amount': round(taxable, 2), 'kind': 'total'},
        {'label': 'Impuesto anual calculado', 'amount': round(tax, 2), 'kind': 'total'},
        {'label': 'Pagos a cuenta anualizados', 'amount': round(payments_annual, 2), 'kind': 'credit'},
        {'label': 'Retenciones del ejercicio', 'amount': round(withholding, 2), 'kind': 'credit'},
        {'label': 'Saldo por regularizar (+) / crédito por revisar (-)', 'amount': round(balance, 2), 'kind': 'balance'},
    ]
    steps = [
        fourth_step,
        f'Renta neta de cuarta más renta bruta de quinta: S/ {work_net_before_7uit:,.2f}.',
        f'Deducción fija de 7 UIT (S/ {7*uit:,.2f}); base previa a deducción adicional: S/ {base_after_7uit:,.2f}.',
        f'Deducción adicional sustentada aplicada: S/ {additional_applied:,.2f} (máximo conjunto: 3 UIT y sujeta a requisitos).',
        f'Deducción ITF aplicada: S/ {itf_applied:,.2f}; donaciones deducibles aplicadas: S/ {donations_applied:,.2f} (sujetas a sustento y límites legales).',
        f'Renta imponible: S/ {taxable:,.2f}; impuesto anual por escala progresiva: S/ {tax:,.2f}.',
        f'Pagos a cuenta anualizados: S/ {payments_annual:,.2f}; retenciones ingresadas: S/ {withholding:,.2f}.',
        ('Saldo por regularizar' if balance > 0 else 'Saldo a favor / crédito por revisar') + f': S/ {abs(balance):,.2f}.'
    ]
    return {
        'annual_income': round(income_annual, 2), 'base': round(taxable, 2), 'tax': round(tax, 2),
        'tax_due': balance, 'payments_annual': round(payments_annual, 2), 'withholding': round(withholding, 2),
        'uit': uit, 'status': 'calculated', 'taxable_base': round(taxable, 2),
        'explanation': 'Cálculo anual de rentas de trabajo, combinando cuarta y quinta cuando se informaron ambas. Las deducciones adicionales solo deben ingresarse si están sustentadas y cumplen los requisitos legales.',
        'steps': steps, 'credits': round(credits, 2), 'balance': balance, 'schedule': schedule, 'tax_bands': tax_bands, 'uit_deduction': round(min(7 * uit, work_net_before_7uit), 2), 'additional_cap': round(3 * uit, 2), 'fourth_gross': round(fourth, 2), 'fifth_gross': round(fifth, 2), 'net_fourth': round(net_fourth, 2),
        'deduction_fourth': round(fourth_deduction, 2), 'gross_work_income': round(fourth + fifth, 2),
        'additional_deductions_applied': round(additional_applied, 2), 'itf_applied': round(itf_applied, 2), 'donations_applied': round(donations_applied, 2), 'tax_year': int(year),
        'legal_basis': ['Artículo 33, inciso a), de la Ley del Impuesto a la Renta: rentas por ejercicio individual de profesión, arte, ciencia u oficio.', 'Artículo 45, primer párrafo, de la Ley del Impuesto a la Renta: deducción del 20% de la renta bruta de cuarta, con límite de 24 UIT; revisar exclusiones del propio artículo para directores, síndicos, mandatarios, albaceas y funciones similares.', 'Artículo 46, primer párrafo, de la Ley del Impuesto a la Renta: deducción fija de 7 UIT sobre rentas netas de trabajo; los párrafos siguientes regulan deducciones adicionales y requisitos.', 'Artículo 46 de la Ley del Impuesto a la Renta: deducciones adicionales de hasta 3 UIT por conceptos legalmente admitidos y debidamente sustentados; el ITF y donaciones requieren revisión de sus reglas específicas.', 'Artículo 53 de la Ley del Impuesto a la Renta: tasas progresivas acumulativas de 8%, 14%, 17%, 20% y 30%.', ('D.S. N.º 301-2025-EF: UIT 2026 de S/ 5,500.' if int(year) == 2026 else 'D.S. N.º 260-2024-EF: UIT 2025 de S/ 5,350.'), 'Artículo 49 de la Ley del Impuesto a la Renta: revisar requisitos y límites de las donaciones deducibles.']
    }

def get_case(case_id):
    with db() as conn:
        return conn.execute("SELECT * FROM cases WHERE id=? AND user_id=?", (case_id, session.get('user_id'))).fetchone()

@app.route("/")
def index():
    with db() as conn:
        total = conn.execute("SELECT COUNT(*) FROM cases WHERE user_id=?", (session.get('user_id'),)).fetchone()[0]
        recent = conn.execute("SELECT * FROM cases WHERE user_id=? ORDER BY id DESC LIMIT 4", (session.get('user_id'),)).fetchall()
        counts = {row['category']: row['n'] for row in conn.execute("SELECT category, COUNT(*) AS n FROM cases WHERE user_id=? GROUP BY category", (session.get('user_id'),)).fetchall()}
    return render_template("page.html", page="inicio", total=total, recent=recent, categories=CATEGORIES, counts=counts)

@app.route("/nuevo-analisis")
def new_analysis():
    return render_template("page.html", page="nuevo", categories=CATEGORIES)

@app.route("/nuevo-analisis/<category>")
def analysis_form(category):
    if category not in CATEGORIES:
        abort(404)
    return render_template("page.html", page="formulario", category=category, categories=CATEGORIES)

@app.route("/analizar", methods=["POST"])
def analyze():
    try:
        person = request.form.get("person", "").strip() or "Caso sin nombre"
        category = request.form.get("category", "cuarta")
        if category not in CATEGORIES: category = "cuarta"
        platform = request.form.get("platform", "").strip()
        income = max(float(request.form.get("income") or 0), 0)
        expenses = max(float(request.form.get("expenses") or 0), 0)
        withholding = max(float(request.form.get("withholding") or 0), 0)
        payments = max(float(request.form.get("payments") or 0), 0)
        additional_deductions = max(float(request.form.get("additional_deductions") or 0), 0)
        itf_deduction = max(float(request.form.get("itf_deduction") or 0), 0)
        donations = max(float(request.form.get("donations") or 0), 0)
        other_fourth = max(float(request.form.get("other_fourth") or 0), 0)
        other_fifth = max(float(request.form.get("other_fifth") or 0), 0)
        year = int(request.form.get("tax_year") or 2026)
        income_frequency = request.form.get("income_frequency", "annual")
        payments_frequency = request.form.get("payments_frequency", "annual")
        country = request.form.get("country", "Perú")
        notes = request.form.get("notes", "").strip()
        director = request.form.get("director") == "yes"
        calc = calculate(category, income, expenses, withholding, year=year, payments=payments,
                         additional_deductions=additional_deductions, itf_deduction=itf_deduction, donations=donations, other_fourth=other_fourth,
                         other_fifth=other_fifth, income_frequency=income_frequency,
                         payments_frequency=payments_frequency, director=director)
        code = "TR-" + datetime.now().strftime("%y%m%d%H%M%S%f")[-16:]
        income_annual = calc['annual_income']
        payments_annual = calc['payments_annual']
        base, tax = calc['base'], calc['tax']
        with db() as conn:
            cur = conn.execute("""INSERT INTO cases
                (user_id, code, person, category, platform, income, expenses, withholding, country, notes, tax_base, estimated_tax, created_at,
                 tax_year, payments, additional_deductions, itf_deduction, donations, director_role, other_fourth, other_fifth, income_frequency, payments_frequency, calculation_json)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                (session.get('user_id'), code, person, category, platform, income_annual, expenses, withholding, country, notes, base, tax,
                 datetime.now().strftime("%d/%m/%Y %H:%M"), year, payments_annual, additional_deductions, itf_deduction, donations, int(director),
                 other_fourth, other_fifth, income_frequency, payments_frequency, json.dumps(calc, ensure_ascii=False)))
            case_id = cur.lastrowid
        flash("Ficha guardada. Revisa la explicación normativa y la clasificación del ingreso.", "success")
        return redirect(url_for("result", case_id=case_id))
    except (ValueError, TypeError) as exc:
        flash("Verifica que los importes y el ejercicio fiscal sean válidos.", "error")
        return redirect(url_for("new_analysis"))

@app.route("/resultado/<int:case_id>")
def result(case_id):
    case = get_case(case_id)
    if not case: abort(404)
    try: calc = json.loads(case['calculation_json'] or '{}')
    except Exception: calc = {}
    if not calc:
        calc = {'status': 'requires_review', 'explanation': 'Este registro proviene de una versión anterior o fue importado sin trazabilidad del cálculo. Confirma si el ingreso es mensual o anual, completa el ejercicio y vuelve a registrar/editar el caso para evitar usar un resultado anterior no validado.', 'steps': [], 'legal_basis': ['Validar Ley del Impuesto a la Renta y orientación SUNAT aplicables al ejercicio.']}
    return render_template("page.html", page="resultado", case=case, calc=calc, categories=CATEGORIES)

@app.route("/historial")
def history():
    q = request.args.get("q", "").strip()
    category = request.args.get("category", "")
    sql = "SELECT * FROM cases WHERE user_id=? AND (code LIKE ? OR person LIKE ? OR platform LIKE ?)"
    args = [session.get('user_id'), f"%{q}%", f"%{q}%", f"%{q}%"]
    if category in CATEGORIES:
        sql += " AND category=?"; args.append(category)
    sql += " ORDER BY id DESC"
    with db() as conn:
        cases = conn.execute(sql, args).fetchall()
    return render_template("page.html", page="historial", cases=cases, q=q, selected_category=category, categories=CATEGORIES)

@app.route("/mis-casos")
def my_cases():
    return redirect(url_for('history'))

@app.route("/caso/<int:case_id>/eliminar", methods=["POST"])
def delete_case(case_id):
    with db() as conn:
        conn.execute("DELETE FROM cases WHERE id=? AND user_id=?", (case_id, session.get('user_id')))
    flash("Caso eliminado.", "success")
    return redirect(request.referrer or url_for("my_cases"))

@app.route('/historial/eliminar-todos', methods=['POST'])
def clear_history():
    confirm = request.form.get('confirm_clear') == 'yes'
    if not confirm:
        flash('Marca la confirmación para eliminar todo tu historial.', 'error')
        return redirect(url_for('history'))
    with db() as conn:
        conn.execute('DELETE FROM cases WHERE user_id=?', (session.get('user_id'),))
    flash('Se eliminó todo el historial de tu cuenta. Las cuentas de otras personas no fueron modificadas.', 'success')
    return redirect(url_for('history'))

@app.route("/caso/<int:case_id>/editar")
def edit_case(case_id):
    case = get_case(case_id)
    if not case: abort(404)
    return render_template("page.html", page="editar", case=case, categories=CATEGORIES)

@app.route("/caso/<int:case_id>/actualizar", methods=["POST"])
def update_case(case_id):
    case = get_case(case_id)
    if not case: abort(404)
    person = request.form.get("person", "").strip() or "Caso sin nombre"
    category = request.form.get("category", "cuarta")
    if category not in CATEGORIES: category = "cuarta"
    platform = request.form.get("platform", "").strip()
    country = request.form.get("country", case['country'] or 'Perú')
    income = max(float(request.form.get("income") or 0), 0)
    payments = max(float(request.form.get("payments") or 0), 0)
    withholding = max(float(request.form.get("withholding") or 0), 0)
    notes = request.form.get("notes", "").strip()
    year = int(request.form.get("tax_year") or case['tax_year'] or 2026)
    other_fourth = max(float(request.form.get("other_fourth") or 0), 0)
    other_fifth = max(float(request.form.get("other_fifth") or 0), 0)
    additional = max(float(request.form.get("additional_deductions") or 0), 0)
    itf_deduction = max(float(request.form.get("itf_deduction") or 0), 0)
    donations = max(float(request.form.get("donations") or 0), 0)
    calc = calculate(category, income, 0, withholding, year=year, payments=payments,
                     additional_deductions=additional, itf_deduction=itf_deduction, donations=donations, other_fourth=other_fourth, other_fifth=other_fifth,
                     income_frequency='annual', payments_frequency='annual', director=request.form.get('director')=='yes')
    with db() as conn:
        conn.execute("""UPDATE cases SET person=?, category=?, platform=?, income=?, payments=?, withholding=?, notes=?, country=?, tax_year=?,
        other_fourth=?, other_fifth=?, additional_deductions=?, itf_deduction=?, donations=?, director_role=?, tax_base=?, estimated_tax=?, calculation_json=? WHERE id=? AND user_id=?""",
        (person, category, platform, max(income,0), payments, withholding, notes, country, year, other_fourth, other_fifth,
         additional, itf_deduction, donations, int(request.form.get('director')=='yes'), calc['base'], calc['tax'], json.dumps(calc, ensure_ascii=False), case_id, session.get('user_id')))
    flash("Caso actualizado y recalculado.", "success")
    return redirect(url_for('result', case_id=case_id))

@app.route("/plantilla-excel")
def excel_page():
    return render_template("page.html", page="excel", categories=CATEGORIES)

@app.route("/plantilla-excel/descargar")
def download_template():
    wb = Workbook()
    ws = wb.active
    ws.title = "Casos"
    headers = ["persona", "categoria", "plataforma", "ingresos_soles", "frecuencia_ingresos", "gastos_soles",
               "pagos_cuenta_soles", "frecuencia_pagos", "retenciones_soles", "otras_rentas_cuarta_soles",
               "rentas_quinta_soles", "deducciones_adicionales_soles", "itf_deducible_soles", "donaciones_soles",
               "funcion_director", "ejercicio", "pais", "observaciones"]
    ws.append(headers)
    ws.append(["Ejemplo demostrativo (editar o borrar)", "cuarta", "", 18000, "anual", 0, 500, "anual", 1200,
               0, 0, 0, 0, 0, "no", 2026, "Perú", "Importes ilustrativos; validar antes de usar"])
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:R{ws.max_row}"
    ws.sheet_view.showGridLines = False
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.worksheet.datavalidation import DataValidation
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="7A1E3C")
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[1].height = 36
    for col in ws.columns:
        letter = col[0].column_letter
        ws.column_dimensions[letter].width = min(max(max(len(str(c.value or "")) for c in col[:2]) + 3, 15), 30)
    ws.column_dimensions["A"].width = 32; ws.column_dimensions["R"].width = 48
    cat = DataValidation(type="list", formula1='"cuarta,quinta,creador,plataforma"', allow_blank=True)
    freq = DataValidation(type="list", formula1='"anual,mensual"', allow_blank=True)
    yesno = DataValidation(type="list", formula1='"no,si"', allow_blank=True)
    for dv in (cat, freq, yesno): ws.add_data_validation(dv)
    cat.add("B2:B501"); freq.add("E2:E501"); freq.add("H2:H501"); yesno.add("O2:O501")
    helpws = wb.create_sheet("Instrucciones")
    helpws.append(["PLANTILLA DE CARGA · TRIBUANÁLISIS · SEMANA 6"])
    helpws.append(["Cómo completar la plantilla"])
    instructions = [
        ["1. Una fila representa un caso tributario. No cambies los nombres de las columnas de la hoja Casos."],
        ["2. persona: nombre o identificador del caso. No incluyas datos personales reales innecesarios."],
        ["3. categoria: usa exactamente cuarta, quinta, creador o plataforma."],
        ["4. ingresos_soles: ingreso principal. frecuencia_ingresos admite anual o mensual; mensual se anualiza x12."],
        ["5. Si corresponde a rentas combinadas, completa otras_rentas_cuarta_soles y/o rentas_quinta_soles."],
        ["6. pagos_cuenta_soles y retenciones_soles: importes del ejercicio; frecuencia_pagos aplica a pagos a cuenta."],
        ["7. deducciones_adicionales_soles, itf_deducible_soles y donaciones_soles: ingresa solo importes revisados y sustentables."],
        ["8. funcion_director: escribe si o no. Para cuarta categoría, valida si corresponde una función excluida de la deducción del 20%."],
        ["9. ejercicio: año tributario (por ejemplo, 2026). pais y observaciones son campos descriptivos."],
        ["10. gastos_soles se guarda como dato informativo; no se resta automáticamente para rentas de cuarta o quinta."],
        ["11. Creador/plataforma: la carga crea una ficha para clasificación. No determina automáticamente la categoría ni el impuesto."],
        ["12. La importación calcula un estimado académico. Verifica la normativa vigente y los datos antes de utilizar resultados."],
        ["Valores admitidos en categoria: cuarta | quinta | creador | plataforma"],
        ["Valores admitidos en frecuencia_ingresos y frecuencia_pagos: anual | mensual"],
    ]
    for row in instructions: helpws.append(row)
    helpws.merge_cells("A1:F1"); helpws["A1"].font = Font(bold=True, size=15, color="FFFFFF"); helpws["A1"].fill = PatternFill("solid", fgColor="7A1E3C")
    helpws["A1"].alignment = Alignment(vertical="center"); helpws.row_dimensions[1].height = 30
    helpws.column_dimensions["A"].width = 115
    for r in range(2, helpws.max_row+1):
        helpws.cell(r,1).alignment = Alignment(wrap_text=True, vertical="top")
        helpws.row_dimensions[r].height = 32
        if r == 2:
            helpws.cell(r,1).font = Font(bold=True, color="0E3A5A", size=12)
    bio = io.BytesIO()
    wb.save(bio); bio.seek(0)
    return send_file(bio, as_attachment=True, download_name="Plantilla_Tribuanalisis_Semana6.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.route("/plantilla-excel/cargar", methods=["POST"])
def upload_excel():
    f = request.files.get("file")
    if not f or not f.filename.lower().endswith(".xlsx"):
        flash("Selecciona un archivo .xlsx válido.", "error")
        return redirect(url_for("excel_page"))
    try:
        wb = load_workbook(f, data_only=True)
        ws = wb["Casos"] if "Casos" in wb.sheetnames else wb.active
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            raise ValueError("El archivo está vacío.")
        def norm_key(value):
            text = str(value or "").strip().lower()
            text = unicodedata.normalize("NFKD", text)
            return "".join(ch for ch in text if not unicodedata.combining(ch)).replace(" ", "_")

        # Acepta tanto la plantilla ampliada como la plantilla simple que ya usa la web.
        # Ejemplo de encabezados antiguos: persona, categoria, Empresa, ingresos_soles,
        # gastos_soles, retencion, pais, observaciones.
        header_aliases = {
            "empresa": "plataforma",
            "empresa_pagadora": "plataforma",
            "empleador": "plataforma",
            "pagador": "plataforma",
            "retencion": "retenciones_soles",
            "retenciones": "retenciones_soles",
            "retencion_soles": "retenciones_soles",
            "ingreso": "ingresos_soles",
            "ingreso_soles": "ingresos_soles",
            "ingresos": "ingresos_soles",
            "gasto_soles": "gastos_soles",
            "gastos": "gastos_soles",
            "pais_origen": "pais",
            "notas": "observaciones",
        }
        headers = []
        for value in rows[0]:
            key = norm_key(value)
            headers.append(header_aliases.get(key, key))

        inserted = 0
        for row in rows[1:]:
            data = dict(zip(headers, row))
            person = str(data.get("persona") or "").strip()
            if not person or person.lower().startswith("ejemplo demostrativo"):
                continue
            cat_raw = str(data.get("categoria") or "cuarta").strip().lower()
            cat_key = norm_key(cat_raw).replace("_", " ")
            cat_key = cat_key.replace("renta de ", "").replace(" categoría", "").replace(" categoria", "").strip()
            category_aliases = {
                "cuarta": "cuarta", "4ta": "cuarta", "cuarta categoria": "cuarta",
                "quinta": "quinta", "5ta": "quinta", "quinta categoria": "quinta",
                "creador": "creador", "creador de contenido": "creador", "influencer": "creador",
                "plataforma": "plataforma", "economia de plataformas": "plataforma",
            }
            cat = category_aliases.get(cat_key, cat_key)
            if cat not in CATEGORIES:
                raise ValueError(f"Categoría no válida en el caso '{person}'. Usa cuarta, quinta, creador o plataforma.")
            def amount(key):
                value = data.get(key)
                return max(float(value or 0), 0)
            income = amount("ingresos_soles")
            expenses = amount("gastos_soles")
            payments = amount("pagos_cuenta_soles")
            withholding = amount("retenciones_soles")
            other_fourth = amount("otras_rentas_cuarta_soles")
            other_fifth = amount("rentas_quinta_soles")
            additional = amount("deducciones_adicionales_soles")
            itf = amount("itf_deducible_soles")
            donations = amount("donaciones_soles")
            year = int(data.get("ejercicio") or 2026)
            income_frequency = str(data.get("frecuencia_ingresos") or "anual").strip().lower()
            payments_frequency = str(data.get("frecuencia_pagos") or "anual").strip().lower()
            if income_frequency not in ("anual", "mensual"): income_frequency = "anual"
            if payments_frequency not in ("anual", "mensual"): payments_frequency = "anual"
            director = str(data.get("funcion_director") or "no").strip().lower() in ("si", "sí", "yes", "true", "1")
            calc = calculate(cat, income, expenses, withholding, year=year, payments=payments,
                             additional_deductions=additional, itf_deduction=itf, donations=donations,
                             other_fourth=other_fourth, other_fifth=other_fifth,
                             income_frequency="monthly" if income_frequency == "mensual" else "annual",
                             payments_frequency="monthly" if payments_frequency == "mensual" else "annual", director=director)
            code = "TR-" + datetime.now().strftime("%y%m%d%H%M%S%f")[-16:] + f"-{inserted+1}"
            with db() as conn:
                conn.execute("""INSERT INTO cases
                (user_id, code, person, category, platform, income, expenses, withholding, country, notes, tax_base, estimated_tax, created_at,
                 tax_year, payments, additional_deductions, itf_deduction, donations, director_role, other_fourth, other_fifth,
                 income_frequency, payments_frequency, calculation_json)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                (session.get('user_id'), code, person, cat, str(data.get("plataforma") or ""), income, expenses, withholding,
                 str(data.get("pais") or "Perú"), str(data.get("observaciones") or ""), calc.get('base', 0), calc.get('tax', 0),
                 datetime.now().strftime("%d/%m/%Y %H:%M"), year, payments, additional, itf, donations, int(director),
                 other_fourth, other_fifth, "monthly" if income_frequency == "mensual" else "annual",
                 "monthly" if payments_frequency == "mensual" else "annual", json.dumps(calc, ensure_ascii=False)))
            inserted += 1
        flash(f"Importación terminada: {inserted} caso(s) agregado(s). Revisa los datos y la clasificación antes de usar el cálculo.", "success")
    except Exception as exc:
        flash(f"No se pudo leer el Excel: {exc}", "error")
    return redirect(url_for("history"))

@app.route("/formulario-709/<int:case_id>")
def formulario_709(case_id):
    case = get_case(case_id)
    if not case: abort(404)
    if case['category'] not in ('cuarta', 'quinta'):
        flash('Primero clasifica los ingresos digitales antes de preparar una cédula de rentas de trabajo.', 'warning')
        return redirect(url_for('reports'))
    try: calc = json.loads(case['calculation_json'] or '{}')
    except Exception: calc = {}
    if not calc or calc.get('status') != 'calculated':
        flash('Este caso no tiene un cálculo anual válido para preparar el documento.', 'warning')
        return redirect(url_for('reports'))
    return render_template('page.html', page='709_form', case=case, calc=calc, categories=CATEGORIES,
                           metadata={}, today=datetime.now().strftime('%Y-%m-%d'))

@app.route("/formulario-709/<int:case_id>/generar", methods=["POST"])
def generate_709(case_id):
    case = get_case(case_id)
    if not case: abort(404)
    if case['category'] not in ('cuarta', 'quinta'):
        abort(400)
    try: calc = json.loads(case['calculation_json'] or '{}')
    except Exception: calc = {}
    if not calc or calc.get('status') != 'calculated':
        flash('El caso no tiene cálculo anual válido.', 'warning')
        return redirect(url_for('reports'))

    ruc = ''.join(ch for ch in request.form.get('ruc', '') if ch.isdigit())
    full_name = request.form.get('full_name', '').strip()
    period = ''.join(ch for ch in request.form.get('period', '') if ch.isdigit())
    if len(ruc) != 11 or len(period) != 6 or not full_name:
        flash('Revisa el RUC (11 dígitos), el periodo (6 dígitos) y los nombres completos.', 'warning')
        return redirect(url_for('formulario_709', case_id=case_id))
    declaration_type = request.form.get('declaration_type', 'ORIGINAL').strip().upper()
    address = request.form.get('address', '').strip()
    prepared_date = request.form.get('prepared_date', '').strip() or datetime.now().strftime('%Y-%m-%d')
    payments_fourth = max(float(request.form.get('payments_fourth') or 0), 0)
    payments_fifth = max(float(request.form.get('payments_fifth') or 0), 0)
    withholding_fourth = max(float(request.form.get('withholding_fourth') or 0), 0)
    withholding_fifth = max(float(request.form.get('withholding_fifth') or 0), 0)
    prior_credit = max(float(request.form.get('prior_credit') or 0), 0)
    foreign_income = max(float(request.form.get('foreign_income') or 0), 0)

    # Map the analysis into a reference cédula for the work-income section only.
    uit = float(calc.get('uit') or (5500 if int(case['tax_year']) == 2026 else 5350))
    gross_fourth = float(calc.get('fourth_gross') or 0)
    gross_fifth = float(calc.get('fifth_gross') or 0)
    deduction_fourth = float(calc.get('deduction_fourth') or 0)
    net_fourth = float(calc.get('net_fourth') or max(gross_fourth - deduction_fourth, 0))
    total_work = float(calc.get('gross_work_income') or (net_fourth + gross_fifth))
    deduction_7 = float(calc.get('uit_deduction') if calc.get('uit_deduction') is not None else min(7 * uit, net_fourth + gross_fifth))
    deduction_3 = float(calc.get('additional_deductions_applied') or 0)
    itf = float(calc.get('itf_applied') or 0)
    donations = float(calc.get('donations_applied') or 0)
    net_after_deductions = max(total_work - deduction_7 - deduction_3, 0)
    net_after_itf_donations = max(net_after_deductions - itf - donations, 0)
    taxable_base = net_after_itf_donations + foreign_income
    annual_tax = progressive_tax(taxable_base, uit)
    payments = payments_fourth + payments_fifth
    withholding = withholding_fourth + withholding_fifth
    credits = payments + withholding + prior_credit
    balance = round(annual_tax - credits, 2)
    refund = abs(balance) if balance < 0 else 0
    tax_due = balance if balance > 0 else 0

    from xml.sax.saxutils import escape
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, rightMargin=27, leftMargin=27, topMargin=25, bottomMargin=32,
                            title=f"Cédula referencial FV709 {period}", author="Tribuanálisis")
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name='709Title', parent=styles['Title'], fontName='Helvetica-Bold', fontSize=13, leading=16, alignment=1, textColor=colors.HexColor('#222222'), spaceAfter=2))
    styles.add(ParagraphStyle(name='709Small', parent=styles['Normal'], fontName='Helvetica', fontSize=7.5, leading=9, textColor=colors.HexColor('#333333')))
    styles.add(ParagraphStyle(name='709Tiny', parent=styles['Normal'], fontName='Helvetica', fontSize=6.8, leading=8, textColor=colors.HexColor('#333333')))
    styles.add(ParagraphStyle(name='709Section', parent=styles['Heading2'], fontName='Helvetica-Bold', fontSize=8.5, leading=10, textColor=colors.white, backColor=colors.HexColor('#333333'), borderPadding=4, spaceBefore=7, spaceAfter=0))
    def P709(v, style='709Tiny'):
        return Paragraph(escape(str(v if v not in (None, '') else '—')).replace('\n','<br/>'), styles[style])
    def sol(v, negative=False):
        n = float(v or 0)
        return f"({abs(n):,.2f})" if negative and n else f"{n:,.2f}"
    story=[]
    header = [[P709('SUNAT','709Title'), P709('RENTA ANUAL · RENTAS DE TRABAJO Y/O FUENTE EXTRANJERA','709Title')],
              [P709('MODELO ACADÉMICO','709Small'), P709('Cédula referencial para traslado y verificación · No presentado ante SUNAT','709Small')]]
    ht=Table(header, colWidths=[105, 430])
    ht.setStyle(TableStyle([('GRID',(0,0),(-1,-1),.5,colors.HexColor('#777777')),('VALIGN',(0,0),(-1,-1),'MIDDLE'),('ALIGN',(0,0),(0,-1),'CENTER'),('SPAN',(0,0),(0,1)),('LEFTPADDING',(0,0),(-1,-1),6),('RIGHTPADDING',(0,0),(-1,-1),6),('TOPPADDING',(0,0),(-1,-1),6),('BOTTOMPADDING',(0,0),(-1,-1),6)]))
    story += [ht, Spacer(1,5)]
    ident = [
      [P709('RUC'), P709(ruc), P709('Periodo'), P709(period)],
      [P709('Nombres completos / razón social'), P709(full_name), P709('Ejercicio'), P709(case['tax_year'])],
      [P709('Tipo de declaración'), P709(declaration_type), P709('Moneda'), P709('SOLES')],
      [P709('Domicilio fiscal'), P709(address or 'No consignado'), P709('Fecha de elaboración'), P709(prepared_date)],
      [P709('Número de orden SUNAT'), P709('NO ASIGNADO · GENERACIÓN LOCAL'), P709('Caso'), P709(case['code'])],
    ]
    it=Table(ident, colWidths=[100,205,95,135])
    it.setStyle(TableStyle([('GRID',(0,0),(-1,-1),.4,colors.HexColor('#888888')),('BACKGROUND',(0,0),(0,-1),colors.HexColor('#F0F0F0')),('BACKGROUND',(2,0),(2,-1),colors.HexColor('#F0F0F0')),('VALIGN',(0,0),(-1,-1),'MIDDLE'),('LEFTPADDING',(0,0),(-1,-1),5),('RIGHTPADDING',(0,0),(-1,-1),5),('TOPPADDING',(0,0),(-1,-1),4),('BOTTOMPADDING',(0,0),(-1,-1),4)]))
    story += [it, Spacer(1,5), P709('RENTAS DE TRABAJO Y/O FUENTE EXTRANJERA Y DETERMINACIÓN DE LA DEUDA','709Section')]
    rows = [[P709('Concepto'), P709('Casilla'), P709('Importe (S/)')]]
    def add(label, casilla, amount, neg=False): rows.append([P709(label), P709(casilla), P709(sol(amount, neg))])
    add('Renta bruta obtenida por el ejercicio individual (profesión, arte, ciencia u oficio)', '107', gross_fourth)
    add('Deducción del 20% de la casilla 107, con límite de 24 UIT', '507', deduction_fourth, True)
    add('Renta neta obtenida por ejercicio individual', '508', net_fourth)
    add('Otras rentas de cuarta categoría no separadas en el análisis', '108', 0)
    add('TOTAL RENTAS DE CUARTA CATEGORÍA', '509', net_fourth)
    add('TOTAL RENTA DE QUINTA CATEGORÍA', '111', gross_fifth)
    add('TOTAL RENTAS DE CUARTA Y QUINTA CATEGORÍA', '510', total_work)
    add('Deducción de 7 UIT (limitada al importe de la casilla 510)', '511', deduction_7, True)
    add('Deducciones adicionales admitidas, hasta 3 UIT y sujetas a requisitos', '514', deduction_3, True)
    add('TOTAL RENTA NETA DE CUARTA Y QUINTA CATEGORÍA', '512', net_after_deductions)
    add('Deducción por Impuesto a las Transacciones Financieras (ITF)', '522', itf, True)
    add('Deducción por donaciones admitidas', '519', donations, True)
    add('Subtotal después de ITF y donaciones', '513', net_after_itf_donations)
    add('Renta neta de fuente extranjera computable ingresada en este formulario', '116', foreign_income)
    add('TOTAL RENTA IMPONIBLE DE TRABAJO Y FUENTE EXTRANJERA', '517', taxable_base)
    add('DETERMINACIÓN DE LA DEUDA', '', '')
    add('Impuesto a la Renta anual estimado por escala progresiva', '120', annual_tax)
    add('Crédito por impuesto a la renta de fuente extranjera', '122', 0, True)
    add('Subtotal', '158', annual_tax)
    add('Saldo a favor de ejercicios anteriores', '125', prior_credit, True)
    add('Pago directo / pagos a cuenta del Impuesto a la Renta de cuarta categoría', '127', payments_fourth, True)
    add('Pago directo del Impuesto a la Renta de quinta categoría', '128', payments_fifth, True)
    add('Impuesto retenido sobre rentas de cuarta categoría', '130', withholding_fourth, True)
    add('Impuesto retenido sobre rentas de quinta categoría', '131', withholding_fifth, True)
    add('A FAVOR DEL CONTRIBUYENTE (estimado)', '141', refund)
    add('Saldo por regularizar — 1 Devolución / 2 Aplicación', '140', 0 if not refund else 1)
    add('A FAVOR DEL FISCO (estimado)', '142', tax_due)
    add('Pagos contra el saldo a favor del fisco previos a la presentación', '144', 0, True)
    add('Interés moratorio (no calculado en este documento)', '145', 0)
    add('Saldo de la deuda tributaria estimado', '146', tax_due)
    add('IMPORTE A PAGAR estimado', '168', tax_due)
    tbl=Table(rows, colWidths=[390,55,90], repeatRows=1)
    ts=[('GRID',(0,0),(-1,-1),.35,colors.HexColor('#888888')),('VALIGN',(0,0),(-1,-1),'MIDDLE'),('ALIGN',(1,1),(-1,-1),'RIGHT'),('LEFTPADDING',(0,0),(-1,-1),5),('RIGHTPADDING',(0,0),(-1,-1),5),('TOPPADDING',(0,0),(-1,-1),3),('BOTTOMPADDING',(0,0),(-1,-1),3),('BACKGROUND',(0,0),(-1,0),colors.HexColor('#E8E8E8')),('FONTNAME',(0,0),(-1,0),'Helvetica-Bold')]
    for idx,row in enumerate(rows[1:],start=1):
        label = row[0].getPlainText() if hasattr(row[0],'getPlainText') else ''
        if label == 'DETERMINACIÓN DE LA DEUDA':
            ts += [('SPAN',(0,idx),(-1,idx)),('BACKGROUND',(0,idx),(-1,idx),colors.HexColor('#D9D9D9')),('FONTNAME',(0,idx),(-1,idx),'Helvetica-Bold')]
        elif idx in (7,10,15,17,21,25,29,31):
            ts.append(('BACKGROUND',(0,idx),(-1,idx),colors.HexColor('#F1F1F1')))
    tbl.setStyle(TableStyle(ts))
    story += [tbl, Spacer(1,6), P709('NOTA: este documento reconstruye solo la sección de rentas de trabajo/fuente extranjera a partir del caso de Tribuanálisis. Las casillas y reglas deben verificarse contra el Formulario Virtual 709 vigente del ejercicio correspondiente. El número de orden, la fecha de presentación y la aceptación solo los asigna SUNAT.','709Small'), Spacer(1,3), P709(f'Caso de origen: {case["code"]} · Cálculo generado para ejercicio {case["tax_year"]} · Documento no presentado ni validado por SUNAT.','709Small')]
    def footer709(canvas, document):
        canvas.saveState(); canvas.setFont('Helvetica-Bold',7); canvas.setFillColor(colors.HexColor('#7A1E3C'))
        canvas.drawString(27,17,'MODELO ACADÉMICO · NO ES CONSTANCIA OFICIAL SUNAT')
        canvas.setFont('Helvetica',7); canvas.setFillColor(colors.HexColor('#333333')); canvas.drawRightString(A4[0]-27,17,f'Página {document.page}')
        canvas.restoreState()
    doc.build(story, onFirstPage=footer709, onLaterPages=footer709)
    buf.seek(0)
    safe_ruc = ruc
    return send_file(buf, as_attachment=True, download_name=f"Modelo_FV709_{period}_{safe_ruc}.pdf", mimetype='application/pdf')

@app.route("/reportes")
def reports():
    with db() as conn:
        cases = conn.execute("SELECT * FROM cases WHERE user_id=? ORDER BY id DESC LIMIT 30", (session.get('user_id'),)).fetchall()
    return render_template("page.html", page="reportes", cases=cases, categories=CATEGORIES)

@app.route("/reporte/<int:case_id>.pdf")
def pdf_report(case_id):
    case = get_case(case_id)
    if not case: abort(404)
    try: calc = json.loads(case['calculation_json'] or '{}')
    except Exception: calc = {}
    if not calc:
        calc = {'status': 'requires_review', 'explanation': 'Registro anterior sin trazabilidad de cálculo; requiere validación antes de utilizar sus importes.'}
    from xml.sax.saxutils import escape
    from reportlab.platypus import KeepTogether, HRFlowable
    from reportlab.graphics.shapes import Drawing, String, Rect
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, rightMargin=38, leftMargin=38, topMargin=38, bottomMargin=40,
                            title=f"Informe técnico {case['code']}", author="Tribuanálisis")
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="Brand2", parent=styles['Title'], textColor=colors.HexColor('#7A1E3C'), fontSize=24, leading=28, alignment=TA_LEFT, spaceAfter=4))
    styles.add(ParagraphStyle(name="SubBrand", parent=styles['Normal'], textColor=colors.HexColor('#0E3A5A'), fontSize=10, leading=14, spaceAfter=12))
    styles.add(ParagraphStyle(name="SectionBrand", parent=styles['Heading2'], textColor=colors.HexColor('#7A1E3C'), fontSize=13, leading=16, spaceBefore=12, spaceAfter=7))
    styles.add(ParagraphStyle(name="BodySmall2", parent=styles['BodyText'], fontSize=8.5, leading=12, spaceAfter=4))
    styles.add(ParagraphStyle(name="LegalSmall", parent=styles['BodyText'], fontSize=8, leading=11, textColor=colors.HexColor('#0E3A5A')))
    def P(value, style='BodySmall2'):
        return Paragraph(escape(str(value or '—')).replace('\n','<br/>'), styles[style])
    from reportlab.platypus import ListFlowable, ListItem
    def ListFlowableCompat(items):
        return ListFlowable([ListItem(item, leftIndent=12) for item in items], bulletType='bullet', leftIndent=14, bulletFontSize=7, spaceAfter=8)
    story = [Paragraph('TRIBUANÁLISIS', styles['Brand2']), Paragraph('INFORME TÉCNICO DE DIAGNÓSTICO TRIBUTARIO · SEMANA 6', styles['SubBrand']),
             HRFlowable(width='100%', thickness=1.2, color=colors.HexColor('#B85B72'), spaceAfter=12)]
    summary = [
        [P('CÓDIGO'), P(case['code']), P('EJERCICIO'), P(case['tax_year'])],
        [P('CONTRIBUYENTE / CASO'), P(case['person']), P('CATEGORÍA'), P(CATEGORIES.get(case['category'], case['category']))],
        [P('ACTIVIDAD / PLATAFORMA'), P(case['platform'] or 'No indicada'), P('PAÍS / PAGADOR'), P(case['country'])],
    ]
    t = Table(summary, colWidths=[90, 170, 75, 160])
    t.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,-1),colors.HexColor('#F4F1EA')),('GRID',(0,0),(-1,-1),.45,colors.HexColor('#D7E4EF')),('TEXTCOLOR',(0,0),(0,-1),colors.HexColor('#7A1E3C')),('TEXTCOLOR',(2,0),(2,-1),colors.HexColor('#7A1E3C')),('VALIGN',(0,0),(-1,-1),'TOP'),('LEFTPADDING',(0,0),(-1,-1),7),('RIGHTPADDING',(0,0),(-1,-1),7),('TOPPADDING',(0,0),(-1,-1),6),('BOTTOMPADDING',(0,0),(-1,-1),6)]))
    story += [t, Spacer(1,12), Paragraph('01 · Resumen ejecutivo', styles['SectionBrand'])]
    if calc.get('status') == 'requires_review':
        story.append(P('No se emitió una liquidación de impuesto de cuarta/quinta categoría. Primero debe clasificarse la naturaleza de los ingresos digitales y determinar si corresponden a rentas empresariales u otra categoría.'))
        story.append(P(calc.get('explanation',''), 'LegalSmall'))
    else:
        metrics = [[P('RENTA BRUTA DE TRABAJO'), P('BASE IMPONIBLE'), P('IMPUESTO ANUAL'), P('SALDO / CRÉDITO')],
                   [P(money(calc.get('gross_work_income', case['income']))), P(money(calc.get('taxable_base', case['tax_base']))), P(money(calc.get('tax', case['estimated_tax']))), P(money(abs(calc.get('balance') or 0)))]]
        mt = Table(metrics, colWidths=[123,123,123,126])
        mt.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#0E3A5A')),('TEXTCOLOR',(0,0),(-1,0),colors.white),('BACKGROUND',(0,1),(-1,1),colors.HexColor('#F4F1EA')),('GRID',(0,0),(-1,-1),.4,colors.HexColor('#D7E4EF')),('ALIGN',(0,0),(-1,-1),'CENTER'),('VALIGN',(0,0),(-1,-1),'MIDDLE'),('LEFTPADDING',(0,0),(-1,-1),5),('RIGHTPADDING',(0,0),(-1,-1),5),('TOPPADDING',(0,0),(-1,-1),7),('BOTTOMPADDING',(0,0),(-1,-1),7)]))
        story += [mt, Spacer(1,8), P('Resultado: ' + ('saldo por regularizar' if (calc.get('balance') or 0)>0 else 'saldo a favor o crédito por revisar') + '. Los créditos deben coincidir con pagos efectivamente efectuados y retenciones del ejercicio.'), Spacer(1,5)]
        # Simple visual comparison chart
        d = Drawing(470, 92)
        values = [('Renta imponible', float(calc.get('taxable_base') or 0), '#4F8CA8'), ('Impuesto anual', float(calc.get('tax') or 0), '#7A1E3C'), ('Créditos', float(calc.get('credits') or 0), '#B85B72')]
        maxv = max([v for _,v,_ in values]+[1])
        for i,(label,value,color) in enumerate(values):
            y = 64 - i*27
            d.add(String(0,y+4,label, fontSize=8, fillColor=colors.HexColor('#0E3A5A')))
            d.add(Rect(115,y,300*value/maxv,12,fillColor=colors.HexColor(color),strokeColor=None))
            d.add(String(min(425,120+300*value/maxv),y+2,money(value),fontSize=7,fillColor=colors.HexColor('#0E3A5A')))
        story += [d, Paragraph('02 · Cédula de determinación del impuesto anual', styles['SectionBrand'])]
        if calc.get('schedule'):
            schedule_rows = [[P('Descripción'), P('Importe (S/)')]]
            for item in calc.get('schedule', []):
                amount = float(item.get('amount') or 0)
                display_amount = f'(S/ {abs(amount):,.2f})' if amount < 0 else f'S/ {amount:,.2f}'
                schedule_rows.append([P(item.get('label')), P(display_amount)])
            schedule_table = Table(schedule_rows, colWidths=[350, 120], repeatRows=1)
            schedule_styles = [
                ('BACKGROUND',(0,0),(-1,0),colors.HexColor('#D7E4EF')),
                ('TEXTCOLOR',(0,0),(-1,0),colors.HexColor('#0E3A5A')),
                ('GRID',(0,0),(-1,-1),.35,colors.HexColor('#B9CBD8')),
                ('ALIGN',(1,0),(1,-1),'RIGHT'),('VALIGN',(0,0),(-1,-1),'TOP'),
                ('LEFTPADDING',(0,0),(-1,-1),7),('RIGHTPADDING',(0,0),(-1,-1),7),
                ('TOPPADDING',(0,0),(-1,-1),5),('BOTTOMPADDING',(0,0),(-1,-1),5),
            ]
            for idx, item in enumerate(calc.get('schedule', []), start=1):
                if item.get('kind') == 'subtotal': schedule_styles.append(('BACKGROUND',(0,idx),(-1,idx),colors.HexColor('#F0F5F8')))
                elif item.get('kind') == 'deduction': schedule_styles.append(('BACKGROUND',(0,idx),(-1,idx),colors.HexColor('#FFF9F0')))
                elif item.get('kind') == 'total': schedule_styles.append(('BACKGROUND',(0,idx),(-1,idx),colors.HexColor('#7A1E3C'))); schedule_styles.append(('TEXTCOLOR',(0,idx),(-1,idx),colors.white))
                elif item.get('kind') == 'credit': schedule_styles.append(('BACKGROUND',(0,idx),(-1,idx),colors.HexColor('#F8E8EC')))
                elif item.get('kind') == 'balance': schedule_styles.append(('BACKGROUND',(0,idx),(-1,idx),colors.HexColor('#F4F1EA'))); schedule_styles.append(('FONTNAME',(0,idx),(-1,idx),'Helvetica-Bold'))
            schedule_table.setStyle(TableStyle(schedule_styles))
            story += [schedule_table, Spacer(1,8)]
        if calc.get('tax_bands'):
            story.append(P('Escala progresiva acumulativa por tramos'))
            band_rows = [[P('Tramo'), P('Tasa'), P('Base (S/)'), P('Impuesto (S/)')]]
            for band in calc.get('tax_bands', []):
                band_rows.append([P(band['label']), P(f"{band['rate']*100:.0f}%"), P(money(band['base'])), P(money(band['tax']))])
            band_table = Table(band_rows, colWidths=[230, 45, 95, 100], repeatRows=1)
            band_table.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#0E3A5A')),('TEXTCOLOR',(0,0),(-1,0),colors.white),('GRID',(0,0),(-1,-1),.35,colors.HexColor('#D7E4EF')),('ALIGN',(1,1),(-1,-1),'RIGHT'),('VALIGN',(0,0),(-1,-1),'TOP'),('LEFTPADDING',(0,0),(-1,-1),5),('RIGHTPADDING',(0,0),(-1,-1),5),('TOPPADDING',(0,0),(-1,-1),5),('BOTTOMPADDING',(0,0),(-1,-1),5)]))
            story += [band_table, Spacer(1,8)]
        story.append(Paragraph('03 · Explicación de los pasos del cálculo', styles['SectionBrand']))
        steps = calc.get('steps', [])
        if steps:
            story.append(ListFlowableCompat([P(f'{i+1}. {step}') for i,step in enumerate(steps)]))
        detail = [['Concepto','Importe'], ['Ingreso principal anualizado', money(case['income'])], ['Otros ingresos de cuarta', money(case['other_fourth'])], ['Otros ingresos de quinta', money(case['other_fifth'])], ['Pagos a cuenta anualizados', money(case['payments'])], ['Retenciones ingresadas', money(case['withholding'])], ['Deducciones adicionales declaradas', money(case['additional_deductions'])], ['ITF deducible ingresado', money(case['itf_deduction'])], ['Donaciones ingresadas', money(case['donations'])]]
        dt=Table([[P(x) for x in row] for row in detail], colWidths=[300,170])
        dt.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#0E3A5A')),('TEXTCOLOR',(0,0),(-1,0),colors.white),('GRID',(0,0),(-1,-1),.35,colors.HexColor('#D7E4EF')),('BACKGROUND',(0,1),(0,-1),colors.HexColor('#F4F1EA')),('VALIGN',(0,0),(-1,-1),'TOP'),('LEFTPADDING',(0,0),(-1,-1),7),('TOPPADDING',(0,0),(-1,-1),5),('BOTTOMPADDING',(0,0),(-1,-1),5)]))
        story += [dt, Paragraph('04 · Cédula de apoyo para la declaración anual', styles['SectionBrand'])]
        predecl = [['Dato de trabajo para revisar en el formulario oficial', 'Importe'],
                   ['Renta neta de trabajo antes de 7 UIT', money(max(float(calc.get('gross_work_income',0))-float(calc.get('deduction_fourth',0)),0))],
                   ['Renta imponible de trabajo calculada', money(calc.get('taxable_base',0))],
                   ['Impuesto anual determinado por el motor', money(calc.get('tax',0))],
                   ['Pagos a cuenta registrados', money(case['payments'])],
                   ['Retenciones registradas', money(case['withholding'])],
                   ['Saldo por regularizar / crédito preliminar', money(abs(calc.get('balance') or 0))]]
        pd=Table([[P(x) for x in row] for row in predecl], colWidths=[330,140])
        pd.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#7A1E3C')),('TEXTCOLOR',(0,0),(-1,0),colors.white),('GRID',(0,0),(-1,-1),.35,colors.HexColor('#D7E4EF')),('BACKGROUND',(0,1),(0,-1),colors.HexColor('#F4F1EA')),('VALIGN',(0,0),(-1,-1),'TOP'),('LEFTPADDING',(0,0),(-1,-1),7),('TOPPADDING',(0,0),(-1,-1),5),('BOTTOMPADDING',(0,0),(-1,-1),5)]))
        story.append(pd)
    story += [Paragraph('05 · Sustento normativo y criterios', styles['SectionBrand'])]
    for law in calc.get('legal_basis', []): story.append(P('• ' + law, 'LegalSmall'))
    story += [Paragraph('06 · Observaciones y evidencia', styles['SectionBrand']), P(case['notes'] or 'No se registraron observaciones adicionales.'),
              Spacer(1,8), Paragraph('Control profesional: antes de presentar una declaración jurada, validar la clasificación de cada ingreso, ejercicio fiscal, UIT, deducciones con requisitos y límites, pagos a cuenta, retenciones, créditos, reglas de fuente y normativa vigente.', styles['LegalSmall']),
              Spacer(1,8), Paragraph('Documento generado por Tribuanálisis como apoyo al análisis. No constituye una declaración jurada presentada ante SUNAT ni reemplaza la revisión profesional.', styles['LegalSmall'])]
    def footer(canvas, document):
        canvas.saveState(); canvas.setStrokeColor(colors.HexColor('#D7E4EF')); canvas.line(38,28,A4[0]-38,28)
        canvas.setFont('Helvetica',7); canvas.setFillColor(colors.HexColor('#0E3A5A'))
        canvas.drawString(38,16,'Tribuanálisis · Informe técnico de apoyo')
        canvas.drawRightString(A4[0]-38,16,f"Página {document.page}")
        canvas.restoreState()
    doc.build(story, onFirstPage=footer, onLaterPages=footer); buf.seek(0)
    return send_file(buf, as_attachment=True, download_name=f"Informe_Tecnico_{case['code']}.pdf", mimetype="application/pdf")

@app.route("/guia")
def guide():
    return render_template("page.html", page="guia", categories=CATEGORIES)

@app.route("/normativa")
def laws():
    return render_template("page.html", page="normativa", categories=CATEGORIES)

@app.route("/perfil")
def profile():
    return render_template("page.html", page="perfil", categories=CATEGORIES)

@app.errorhandler(404)
def not_found(error):
    return render_template("page.html", page="404", categories=CATEGORIES), 404

# Create/migrate the local database both for `python app.py` and WSGI deployments.
init_db()

if __name__ == "__main__":
    app.run(debug=True)
