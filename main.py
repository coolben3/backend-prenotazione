from flask import Flask, request, jsonify
from openpyxl import load_workbook

app = Flask(__name__)

EXCEL_PATH = "template.xlsx"

# Mappatura campi → celle
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
        wb = load_workbook(EXCEL_PATH)
        ws = wb.active

        for campo, cella in CELL_MAP.items():
            valore = data.get(campo)
            if valore is not None:
                ws[cella] = valore

        wb.save(EXCEL_PATH)

        return jsonify({"message": "Dati salvati con successo ✅"}), 200

    except Exception as e:
        return jsonify({"error": "Errore durante la scrittura su Excel", "details": str(e)}), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=3000)
