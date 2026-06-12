from flask import Flask, render_template, request, jsonify, session, redirect, url_for, send_file
from functools import wraps
import json, os, io

app = Flask(__name__)
app.secret_key = "wms_bata_2026_secretkey"

USERS = {
    "jhuamani": {"password": "Bata2026", "role": "admin"}
}

DATA_FILE = "data.json"

def load_data():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"prescrito": {}, "personal": [], "semana_actual": "SEM 23"}

def save_data(data):
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if "user" not in session:
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return decorated

def admin_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if session.get("role") != "admin":
            return jsonify({"error": "Sin permisos de edicion"}), 403
        return f(*args, **kwargs)
    return decorated

@app.route("/")
@login_required
def index():
    return render_template("index.html", user=session.get("user"), role=session.get("role"))

@app.route("/login", methods=["GET", "POST"])
def login():
    error = None
    if request.method == "POST":
        usuario = request.form.get("usuario", "").strip()
        password = request.form.get("password", "").strip()
        if usuario in USERS and USERS[usuario]["password"] == password:
            session["user"] = usuario
            session["role"] = USERS[usuario]["role"]
            return redirect(url_for("index"))
        error = "Usuario o contrasena incorrectos"
    return render_template("login.html", error=error)

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))

@app.route("/api/prescrito", methods=["GET"])
@login_required
def get_prescrito():
    data = load_data()
    return jsonify(data.get("prescrito", {}))

@app.route("/api/prescrito", methods=["POST"])
@login_required
@admin_required
def save_prescrito():
    data = load_data()
    data["prescrito"] = request.json
    save_data(data)
    return jsonify({"ok": True})

@app.route("/api/prescrito/excel")
@login_required
def download_prescrito_excel():
    import openpyxl
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    data = load_data()
    presData = data.get("prescrito", {})
    DIAS = ["LUNES", "MARTES", "MIERCOLES", "JUEVES", "VIERNES"]
    COLS = ["accesorios", "calzado", "insumos", "muestras", "materiales", "doble"]
    LABELS = ["ACCESORIOS", "CALZADO", "INSUMOS", "MUESTRAS", "MATERIALES", "DOBLE TRAMO"]

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Prescrito"

    # Colores
    rojo_oscuro = "991B1B"
    rojo_medio  = "7F1D1D"
    rojo_claro  = "FEE2E2"
    blanco      = "FFFFFF"
    gris_fila   = "FEF9F9"
    negro       = "000000"

    thin = Side(style="thin", color="DDDDDD")
    border = Border(top=thin, bottom=thin, left=thin, right=thin)

    def cell_style(cell, bg, fg="FFFFFF", bold=False, size=11, align="center"):
        cell.fill = PatternFill("solid", fgColor=bg)
        cell.font = Font(bold=bold, color=fg, size=size, name="Calibri")
        cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=True)
        cell.border = border

    # Fila 1: PRESCRITO (merge A1:I1)
    ws.merge_cells("A1:I1")
    ws["A1"] = "PRESCRITO"
    cell_style(ws["A1"], rojo_oscuro, blanco, bold=True, size=13)

    # Fila 2: cabeceras
    headers = ["ATENCION"] + LABELS + ["TOTAL GENERAL"]
    for i, h in enumerate(headers, 1):
        c = ws.cell(row=2, column=i, value=h)
        cell_style(c, rojo_medio, blanco, bold=True, size=10)

    # Filas de datos
    for row_idx, dia in enumerate(DIAS, 3):
        vals = [int(presData.get(dia, {}).get(col, 0) or 0) for col in COLS]
        total = sum(vals)
        bg = gris_fila if row_idx % 2 == 0 else blanco
        c = ws.cell(row=row_idx, column=1, value=dia)
        cell_style(c, bg, negro, bold=True, align="left")
        for ci, v in enumerate(vals, 2):
            cc = ws.cell(row=row_idx, column=ci, value=v)
            cell_style(cc, bg, negro)
        ct = ws.cell(row=row_idx, column=9, value=total)
        cell_style(ct, rojo_claro, rojo_oscuro, bold=True)

    # Fila TOTAL
    total_row = 3 + len(DIAS)
    totals = [sum(int(presData.get(d, {}).get(c, 0) or 0) for d in DIAS) for c in COLS]
    grand = sum(totals)
    ct = ws.cell(row=total_row, column=1, value="TOTAL")
    cell_style(ct, rojo_oscuro, blanco, bold=True, align="left")
    for ci, v in enumerate(totals, 2):
        cc = ws.cell(row=total_row, column=ci, value=v)
        cell_style(cc, rojo_oscuro, blanco, bold=True)
    cg = ws.cell(row=total_row, column=9, value=grand)
    cell_style(cg, rojo_oscuro, blanco, bold=True)

    # Anchos de columna
    anchos = [16, 13, 13, 13, 13, 13, 14, 13, 16]
    for i, w in enumerate(anchos, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    for r in range(1, total_row + 1):
        ws.row_dimensions[r].height = 22

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output, download_name="prescrito_wms_bata.xlsx",
                     as_attachment=True, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)
