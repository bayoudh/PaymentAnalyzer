from flask import Flask, render_template, request, send_file, url_for
import pandas as pd
import io
import re
import uuid
import os
import sys
import threading
import webbrowser


# ============================================================
# CONFIGURATION FOR NORMAL PYTHON + PYINSTALLER EXE
# ============================================================

if getattr(sys, "frozen", False):
    # When running as a PyInstaller EXE:
    # BASE_DIR = folder containing PaymentAnalyzer.exe
    BASE_DIR = os.path.dirname(sys.executable)

    # PyInstaller extracts --add-data files into _MEIPASS
    RESOURCE_DIR = sys._MEIPASS
else:
    # When running app.py normally
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    RESOURCE_DIR = BASE_DIR


TEMPLATES_FOLDER = os.path.join(
    RESOURCE_DIR,
    "templates"
)

STATIC_FOLDER = os.path.join(
    RESOURCE_DIR,
    "static"
)

# Generated Excel files must stay OUTSIDE the PyInstaller
# temporary extraction directory, beside the EXE.
GENERATED_FOLDER = os.path.join(
    BASE_DIR,
    "generated"
)

os.makedirs(
    GENERATED_FOLDER,
    exist_ok=True
)


app = Flask(
    __name__,
    template_folder=TEMPLATES_FOLDER,
    static_folder=STATIC_FOLDER
)

# Maximum upload size: 50 MB per request
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024



# ============================================================
# NETTOYER NUMERO CARTE
# ============================================================

def clean_card(value):
    if pd.isna(value):
        return ""

    value = str(value).strip()

    # Exemple : ="6665767565100"
    if value.startswith('="') and value.endswith('"'):
        value = value[2:-1]

    # Supprimer les espaces
    value = value.replace(" ", "")

    # Exemple scientifique :
    # 6,6657675651E+12
    # ou 6.6657675651E+12
    if "E+" in value.upper():
        try:
            value = value.replace(",", ".")

            number = float(value)

            # Convertir en entier
            value = str(int(number))

        except (ValueError, OverflowError):
            pass

    # Exemple : 6665767565100.0
    if re.fullmatch(r"\d+\.0", value):
        value = value[:-2]

    return value


# ============================================================
# NETTOYER MONTANT
# ============================================================

def clean_amount(value):
    if pd.isna(value):
        return 0.0

    value = str(value).strip()

    # Supprimer espaces
    value = value.replace(" ", "")

    # Format européen :
    # 866,7
    # 1.266,70
    if "," in value:
        value = value.replace(".", "")
        value = value.replace(",", ".")

    try:
        return float(value)

    except ValueError:
        return 0.0


# ============================================================
# LIRE FICHIER
# ============================================================

def read_file(file):
    filename = file.filename.lower()

    if filename.endswith(".xlsx"):
        return pd.read_excel(file, dtype=str)

    if filename.endswith(".xls"):
        return pd.read_excel(file, dtype=str)

    if filename.endswith(".csv"):
        return pd.read_csv(
            file,
            dtype=str,
            sep=None,
            engine="python"
        )

    raise ValueError(
        "Format non supporté. Utilisez Excel (.xlsx/.xls) ou CSV."
    )


# ============================================================
# TROUVER COLONNE
# ============================================================

def find_column(df, keywords):

    for column in df.columns:

        name = str(column).strip().lower()

        for keyword in keywords:

            if keyword.lower() in name:
                return column

    return None


# ============================================================
# TRAITEMENT PRINCIPAL
# ============================================================

def process_files(transaction_file, truck_file):

    # ========================================================
    # FICHIER 1 : TRANSACTIONS
    # ========================================================

    transactions = read_file(transaction_file)

    transactions.columns = [
        str(c).strip()
        for c in transactions.columns
    ]

    # Colonne carte
    card_column = find_column(
        transactions,
        [
            "numéro du moyen de paiement",
            "numero du moyen de paiement",
            "numéro du moyen de",
            "numero du moyen de",
            "moyen de paiement",
            "carte",
            "card",
            "payment"
        ]
    )

    # Colonne montant
    amount_column = find_column(
        transactions,
        [
            "montant",
            "amount"
        ]
    )

    if card_column is None:
        raise ValueError(
            "Colonne 'Numéro du moyen de paiement' "
            "introuvable dans le fichier 1."
        )

    if amount_column is None:
        raise ValueError(
            "Colonne 'Montant' "
            "introuvable dans le fichier 1."
        )

    # Nettoyage carte
    transactions["Carte"] = (
        transactions[card_column]
        .apply(clean_card)
    )

    # Nettoyage montant
    transactions["Montant"] = (
        transactions[amount_column]
        .apply(clean_amount)
    )

    # Supprimer cartes vides
    transactions = transactions[
        transactions["Carte"] != ""
    ].copy()

    if transactions.empty:
        raise ValueError(
            "Aucune transaction valide trouvée."
        )

    # ========================================================
    # CALCUL DES TRANSACTIONS
    # ========================================================

    summary = (
        transactions
        .groupby("Carte", as_index=False)
        .agg(
            Nb_transactions=("Carte", "size"),
            Montant_total=("Montant", "sum")
        )
    )

    # ========================================================
    # FICHIER 2 : NB TRUCK
    # ========================================================

    trucks = read_file(truck_file)

    trucks.columns = [
        str(c).strip()
        for c in trucks.columns
    ]

    # Colonne carte
    truck_card_column = find_column(
        trucks,
        [
            "numéro du moyen de paiement",
            "numero du moyen de paiement",
            "numéro du moyen de",
            "numero du moyen de",
            "moyen de paiement",
            "carte",
            "card",
            "payment"
        ]
    )

    # Colonne nb truck
    truck_value_column = find_column(
        trucks,
        [
            "nb truck",
            "nb track",
            "nb trak",
            "truck",
            "track"
        ]
    )

    if truck_card_column is None:
        raise ValueError(
            "Colonne 'Numéro du moyen de paiement' "
            "introuvable dans le fichier 2."
        )

    if truck_value_column is None:
        raise ValueError(
            "Colonne 'nb truck' "
            "introuvable dans le fichier 2."
        )

    # Nettoyage carte
    trucks["Carte"] = (
        trucks[truck_card_column]
        .apply(clean_card)
    )

    # Nettoyage nb truck
    trucks["nb truck"] = (
        trucks[truck_value_column]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    # Garder uniquement les colonnes nécessaires
    trucks = trucks[
        ["Carte", "nb truck"]
    ]

    # Supprimer cartes vides
    trucks = trucks[
        trucks["Carte"] != ""
    ]

    # Une seule ligne par carte
    trucks = trucks.drop_duplicates(
        subset=["Carte"],
        keep="first"
    )

    # ========================================================
    # MERGE
    # ========================================================

    result = pd.merge(
        summary,
        trucks,
        on="Carte",
        how="left"
    )

    result["nb truck"] = (
        result["nb truck"]
        .fillna("")
    )

    # Réorganiser les colonnes
    result = result[
        [
            "Carte",
            "nb truck",
            "Nb_transactions",
            "Montant_total"
        ]
    ]

    # Trier par carte
    result = result.sort_values(
        by="Carte"
    ).reset_index(drop=True)

    # ========================================================
    # TOTAUX
    # ========================================================

    total_transactions = int(
        result["Nb_transactions"].sum()
    )

    total_amount = float(
        result["Montant_total"].sum()
    )

    return (
        result,
        total_transactions,
        total_amount
    )


# ============================================================
# FORMAT MONTANT
# ============================================================

def format_money(value):

    return (
        f"{float(value):,.2f}"
        .replace(",", " ")
        .replace(".", ",")
    )


# ============================================================
# PAGE PRINCIPALE
# ============================================================

@app.route("/", methods=["GET", "POST"])
def index():

    result = None

    total_transactions = 0

    total_amount = 0

    error = None

    download_url = None

    if request.method == "POST":

        transaction_file = request.files.get(
            "transaction_file"
        )

        truck_file = request.files.get(
            "truck_file"
        )

        # ====================================================
        # VERIFICATION FICHIERS
        # ====================================================

        if not transaction_file:

            error = (
                "Veuillez sélectionner le fichier transactions."
            )

        elif not truck_file:

            error = (
                "Veuillez sélectionner le fichier nb truck."
            )

        elif transaction_file.filename == "":

            error = (
                "Le fichier transactions est vide."
            )

        elif truck_file.filename == "":

            error = (
                "Le fichier nb truck est vide."
            )

        else:

            try:

                # =================================================
                # TRAITEMENT
                # =================================================

                (
                    result,
                    total_transactions,
                    total_amount
                ) = process_files(
                    transaction_file,
                    truck_file
                )

                # =================================================
                # CREATION EXCEL
                # =================================================

                output = io.BytesIO()

                export_data = result.copy()

                # Ajouter ligne TOTAL
                total_row = pd.DataFrame({
                    "Carte": ["TOTAL"],

                    "nb truck": [""],

                    "Nb_transactions": [
                        total_transactions
                    ],

                    "Montant_total": [
                        total_amount
                    ]
                })

                export_data = pd.concat(
                    [
                        export_data,
                        total_row
                    ],
                    ignore_index=True
                )

                # Créer Excel en mémoire
                with pd.ExcelWriter(
                    output,
                    engine="openpyxl"
                ) as writer:

                    export_data.to_excel(
                        writer,
                        index=False,
                        sheet_name="Résultat"
                    )

                    # Ajuster largeur colonnes
                    worksheet = writer.sheets["Résultat"]

                    worksheet.column_dimensions["A"].width = 22
                    worksheet.column_dimensions["B"].width = 20
                    worksheet.column_dimensions["C"].width = 20
                    worksheet.column_dimensions["D"].width = 20

                    # Format montant
                    for cell in worksheet["D"][1:]:
                        cell.number_format = '#,##0.00'

                # =================================================
                # SAUVEGARDER LE FICHIER SUR LE SERVEUR
                # =================================================

                output.seek(0)

                file_id = str(uuid.uuid4())

                file_path = os.path.join(
                    GENERATED_FOLDER,
                    f"{file_id}.xlsx"
                )

                with open(
                    file_path,
                    "wb"
                ) as f:

                    f.write(
                        output.getvalue()
                    )

                # =================================================
                # URL DE TELECHARGEMENT
                # =================================================

                download_url = url_for(
                    "export",
                    file_id=file_id
                )

            except Exception as e:

                error = str(e)

    return render_template(
        "index.html",

        result=result,

        total_transactions=total_transactions,

        total_amount=total_amount,

        format_money=format_money,

        error=error,

        download_url=download_url
    )


# ============================================================
# TELECHARGEMENT EXCEL
# ============================================================

@app.route("/export/<file_id>")
def export(file_id):

    # The generated folder is absolute and works with the EXE.
    file_path = os.path.join(
        GENERATED_FOLDER,
        f"{file_id}.xlsx"
    )

    # Security: only allow the generated UUID-style filename.
    if not re.fullmatch(
        r"[0-9a-fA-F-]{36}",
        file_id
    ):
        return "Identifiant de fichier invalide.", 400

    if not os.path.isfile(file_path):
        return (
            "Fichier Excel introuvable.",
            404
        )

    try:
        return send_file(
            file_path,
            as_attachment=True,
            download_name="resultat_transactions.xlsx",
            mimetype=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            )
        )

    except Exception as e:
        return (
            f"Erreur lors du téléchargement : {str(e)}",
            500
        )


# ============================================================
# START
# ============================================================

if __name__ == "__main__":

    # Open the browser automatically only after the Flask server
    # has had a moment to start.
    def open_browser():
        webbrowser.open("http://127.0.0.1:5000")

    threading.Timer(
        1.5,
        open_browser
    ).start()

    app.run(
        host="127.0.0.1",
        port=5000,
        debug=False,
        use_reloader=False
    )
