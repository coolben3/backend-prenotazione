from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
from openpyxl import load_workbook
import csv
import os

app = Flask(__name__)
CORS(app)  # 🔓 Consente richieste da domini esterni (GitHub Pages)

EXCEL_PATH = "template.xlsx"
CSV_PATH = "storico.csv"

# Mappatura campi → celle Excel
CELL_MAP = {
    "nome_cognome": "B6",
    "tratta": "E13",
    "giorno": "E16",
    "mese": "E17",
    "adulti": "M6",
    "bambini": "M7",
    "neonati": "M8",
    "moto": "M9",
    "auto": "M10",
    "orario": "D15",
    "email": "B22"
}

@app.route("/invia", methods=["POST"])
def ricevi_dati():
    data = request.get_json()

    if not data:
        return jsonify({"error": "Nessun dato ricevuto"}), 400

    try:
        # ✅ Scrive nei campi specifici del file Excel
        wb = load_workbook(EXCEL_PATH)
        ws = wb.active

        for campo, cella in CELL_MAP.items():
            valore = data.get(campo)
            if valore is not None:
                ws[cella] = valore

        wb.save(EXCEL_PATH)

        # ✅ Salva i dati anche nello storico CSV
        aggiungi_a_csv(data)

        return jsonify({"message": "Dati salvati con successo ✅"}), 200

    except Exception as e:
        return jsonify({"error": "Errore durante la scrittura", "details": str(e)}), 500

def aggiungi_a_csv(data):
    headers = list(data.keys())
    file_esiste = os.path.isfile(CSV_PATH)

    with open(CSV_PATH, mode='a', newline='', encoding='utf-8') as file:
        writer = csv.DictWriter(file, fieldnames=headers)
        if not file_esiste:
            writer.writeheader()
        writer.writerow(data)

@app.route("/download", methods=["GET"])
def scarica_excel():
    try:
        return send_file(
            EXCEL_PATH,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name='template.xlsx'
        )
    except Exception as e:
        return jsonify({"error": "Errore durante il download", "details": str(e)}), 500

@app.route("/storico", methods=["GET"])
def scarica_csv():
    try:
        return send_file(
            CSV_PATH,
            mimetype='text/csv',
            as_attachment=True,
            download_name='storico.csv'
        )
    except Exception as e:
        return jsonify({"error": "Errore durante il download del CSV", "details": str(e)}), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=3000)
