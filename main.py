from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from datetime import datetime
import pyodbc
import webbrowser

# Connexion à la base de données
conn = pyodbc.connect('DRIVER={SQL Server}; SERVER=DESKTOP-KIO15OT;DATABASE=BH_BANK; Trusted_Connection=yes;')
cursor = conn.cursor()


def generatePDF(num_ope, acheteur):
    # Importation des données 
    cursor.execute("SELECT DATVAL, MNETAC, DEVACH, TAUXACH, MARACH, DATECH FROM hope WHERE NUMOPE = ?", num_ope)
    datval, mnetac, devach, tauxach, marach, datech = cursor.fetchone()

    # Formatage des dates
    date_op = datetime.strptime(datval, '%Y-%m-%d').strftime('%d/%m/%Y')
    date_paie = datetime.strptime(datech, '%Y-%m-%d').strftime('%d/%m/%Y')

    # Calcul des montants
    montant = f"{float(mnetac):,.2f} {devach}"
    marge_revenant = f"{float(marach):,.2f} {devach}"
    total_principal_profit = f"{float(mnetac) + float(marach):,.2f} {devach}"

    # Créer un PDF
    c = canvas.Canvas(f"rapportf{num_ope}.pdf", pagesize=A4)
    c.setFont("Helvetica", 11)

    # En-tête du rapport
    c.setFont("Helvetica-Bold", 11)
    c.drawRightString(500, 750, f"Tunis, le {date_op}")
    c.drawCentredString(300, 700, "RAPPORT D’EXECUTION")
    c.drawCentredString(300, 680, "D’OPPERATION D’INVESTISSEMENT")
    c.setFont("Helvetica", 11)
    c.drawString(50, 640, "En application de la convention cadre nous avons l’honneur de vous communiquer les détails de")
    c.drawString(50, 620, "l’opération d’investissement réalisée le :")

    # Détails de l'opération
    y = 600 
    line_height = 20
    center_page = A4[0] / 2  # A4[0] représente la largeur de la page

    # En-tête de la section détail
    c.setFont("Helvetica-Bold", 11)

    def drawDetailItem(title, detail, y):
        c.drawRightString(center_page - 20, y, title)
        c.drawString(center_page + 20, y, f": {detail}")

    drawDetailItem("Date de l’opération", date_op, y)
    y -= line_height
    drawDetailItem("Montant", montant, y)
    y -= line_height  
    drawDetailItem("Acheteur", acheteur.upper(), y)
    y -= line_height
    drawDetailItem("Fournisseur", "DIVERS", y)
    y -= line_height
    objet_fin = "FINANCEMENT EN DINAR" if devach == "TND" else "FINANCEMENT EN DEVISE"
    drawDetailItem("Objet du financement", objet_fin, y)
    y -= line_height
    drawDetailItem("Taux de marge prévue", f"{tauxach}%", y)
    y -= line_height
    drawDetailItem("Marge vous revenant", marge_revenant, y)
    y -= line_height
    drawDetailItem("Date de paiement", date_paie, y)

    # Texte statique 
    c.setFont("Helvetica", 11)
    c.drawString(50, 420, "En outre, nous vous confirmons que :")
    c.drawString(50, 400, "- Notre client (acheteur) indiqué ne présente aucun impayé dans le secteur et présente des")
    c.drawString(50, 380, "perspectives futures rassurantes.")
    c.drawString(50, 360, "- Tout incident de paiement futur constaté sur notre client (acheteur) vous sera signalé dans les")
    c.drawString(50, 340, "plus brefs délais.")
    c.drawString(50, 320, "Par ailleurs, nous vous créditeront chez BCT du montant en principal et profit réalisé,")

    text_width = c.stringWidth("Soit ", "Helvetica", 11)
    c.drawString(50, 300, "Soit ")

    c.setFont("Helvetica-Bold", 11)
    c.drawString(50 + text_width, 300, f"{total_principal_profit}")
    text_width += c.stringWidth(f"{total_principal_profit}", "Helvetica-Bold", 11)

    c.setFont("Helvetica", 11)
    c.drawString(50 + text_width, 300, " valeur ")
    text_width += c.stringWidth(" valeur ", "Helvetica", 11)

    c.setFont("Helvetica-Bold", 11)
    c.drawString(50 + text_width, 300, f"{date_paie}")
    text_width += c.stringWidth(f"{date_paie}", "Helvetica-Bold", 11)

    c.setFont("Helvetica", 11)
    c.drawString(50 + text_width, 300, " et ce selon les termes et conditions prévus dans")
    c.drawString(50, 280, "la convention cadre signée entre nos deux institutions.")
    c.drawString(50, 260, "Salutations.")

    c.save()
    webbrowser.open_new(f"rapportf{num_ope}.pdf")