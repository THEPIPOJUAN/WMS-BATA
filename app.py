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

    rojo_oscuro = "991B1B"
    rojo_medio  = "7F1D1D"
    rojo_claro  = "FEE2E2"
    amarillo    = "854D0E"
    verde       = "166534"
    blanco      = "FFFFFF"
    gris_fila   = "FEF9F9"
    negro       = "1F2937"

    thin = Side(style="thin", color="E5E7EB")
    border = Border(top=thin, bottom=thin, left=thin, right=thin)

    def cs(cell, bg, fg="FFFFFF", bold=False, size=11, align="center"):
        cell.fill = PatternFill("solid", fgColor=bg)
        cell.font = Font(bold=bold, color=fg, size=size, name="Calibri")
        cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=False)
        cell.border = border

    # Calcular totales para KPIs
    totals = {c: sum(int(presData.get(d, {}).get(c, 0) or 0) for d in DIAS) for c in COLS}
    grand = sum(totals.values())

    # ── FILA 1: titulo KPIs ──
    kpi_titles = ["TOTAL GENERAL", "CALZADO", "INSUMOS", "ACCESORIOS", "DOBLE TRAMO"]
    kpi_cols   = [None, "calzado", "insumos", "accesorios", "doble"]
    kpi_vals   = [grand, totals["calzado"], totals["insumos"], totals["accesorios"], totals["doble"]]
    kpi_colors = ["991B1B", "991B1B", "854D0E", "1F2937", "166534"]
    kpi_starts = [1, 3, 5, 7, 9]  # columnas donde empieza cada KPI (2 cols cada uno)

    for i, (title, val, color, col_start) in enumerate(zip(kpi_titles, kpi_vals, kpi_colors, kpi_starts)):
        ws.merge_cells(start_row=1, start_column=col_start, end_row=1, end_column=col_start+1)
        ws.merge_cells(start_row=2, start_column=col_start, end_row=2, end_column=col_start+1)
        c1 = ws.cell(row=1, column=col_start, value=title)
       cs(c1, "000000", "FFFFFF", bold=False, size=9)
        c2 = ws.cell(row=2, column=col_start, value=val)
        cs(c2, "000000", color, bold=True, size=14)

    ws.row_dimensions[1].height = 18
    ws.row_dimensions[2].height = 28

    # ── FILA 3: espacio ──
    ws.row_dimensions[3].height = 8

    # ── FILA 4: titulo PRESCRITO ──
    ws.merge_cells("A4:H4")
    c = ws["A4"]
    c.value = "PRESCRITO"
    cs(c, rojo_oscuro, blanco, bold=True, size=13)
    ws.row_dimensions[4].height = 24

    # ── FILA 5: cabeceras ──
    headers = ["ATENCION"] + LABELS + ["TOTAL GENERAL"]
    for i, h in enumerate(headers, 1):
        c = ws.cell(row=5, column=i, value=h)
        cs(c, rojo_medio, blanco, bold=True, size=10)
    ws.row_dimensions[5].height = 20

    # ── FILAS DE DATOS ──
    for row_idx, dia in enumerate(DIAS, 6):
        vals = [int(presData.get(dia, {}).get(col, 0) or 0) for col in COLS]
        total = sum(vals)
        bg = gris_fila if row_idx % 2 == 0 else blanco
        c = ws.cell(row=row_idx, column=1, value=dia)
        cs(c, bg, negro, bold=True, size=11, align="left")
        for ci, v in enumerate(vals, 2):
            cc = ws.cell(row=row_idx, column=ci, value=v)
            cs(cc, bg, negro, size=11)
        ct = ws.cell(row=row_idx, column=8, value=total)
        cs(ct, rojo_claro, rojo_oscuro, bold=True, size=11)
        ws.row_dimensions[row_idx].height = 20

    # ── FILA TOTAL ──
    total_row = 6 + len(DIAS)
    col_totals = [sum(int(presData.get(d, {}).get(c, 0) or 0) for d in DIAS) for c in COLS]
    grand_total = sum(col_totals)
    ct = ws.cell(row=total_row, column=1, value="TOTAL")
    cs(ct, rojo_oscuro, blanco, bold=True, size=11, align="left")
    for ci, v in enumerate(col_totals, 2):
        cc = ws.cell(row=total_row, column=ci, value=v)
        cs(cc, rojo_oscuro, blanco, bold=True, size=11)
    cg = ws.cell(row=total_row, column=8, value=grand_total)
    cs(cg, rojo_oscuro, blanco, bold=True, size=12)
    ws.row_dimensions[total_row].height = 22

    # ── ANCHOS ──
    anchos = [16, 13, 13, 13, 13, 13, 14, 16]
    for i, w in enumerate(anchos, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output, download_name="prescrito_wms_bata.xlsx",
                     as_attachment=True, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)
