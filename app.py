import os
import openpyxl
from flask import Flask, render_template, request, redirect, url_for, session, flash, send_file

app = Flask(__name__)

# Clé secrète pour gérer les sessions de manière sécurisée
app.secret_key = os.environ.get('SECRET_KEY', 'admin')

# Mot de passe requis pour accéder au panneau d'administration
ADMIN_PASSWORD = os.environ.get('ADMIN_PASSWORD', 'admin_1234')

# Route d'accueil (formulaire utilisateur)
@app.route('/')
def index():
    return render_template('index.html')

# Route de connexion Admin
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

# Route de déconnexion Admin
@app.route('/admin/logout')
def admin_logout():
    session.pop('is_admin', None)
    flash('Vous avez été déconnecté.', 'info')
    return redirect(url_for('admin_login'))

# Route Dashboard Admin (protégée)
@app.route('/admin')
def admin_dashboard():
    if not session.get('is_admin'):
        return redirect(url_for('admin_login'))
    
    # Place ici la logique de ton tableau de bord admin existant
    return render_template('admin.html')

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)