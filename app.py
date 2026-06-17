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

    totals = {c: sum(int(presData.get(d, {}).get(c, 0) or 0) for d in DIAS) for c in COLS}
    grand = sum(totals.values())

    kpi_titles = ["TOTAL GENERAL", "CALZADO", "INSUMOS", "ACCESORIOS", "DOBLE TRAMO"]
    kpi_vals   = [grand, totals["calzado"], totals["insumos"], totals["accesorios"], totals["doble"]]
    kpi_colors = ["991B1B", "991B1B", "854D0E", "1F2937", "166534"]
    kpi_starts = [1, 3, 5, 7, 9]

    for i, (title, val, color, col_start) in enumerate(zip(kpi_titles, kpi_vals, kpi_colors, kpi_starts)):
        ws.merge_cells(start_row=1, start_column=col_start, end_row=1, end_column=col_start+1)
        ws.merge_cells(start_row=2, start_column=col_start, end_row=2, end_column=col_start+1)
        c1 = ws.cell(row=1, column=col_start, value=title)
        cs(c1, "000000", "FFFFFF", bold=False, size=9)
        c2 = ws.cell(row=2, column=col_start, value=val)
        cs(c2, "000000", color, bold=True, size=14)

    ws.row_dimensions[1].height = 18
    ws.row_dimensions[2].height = 28
    ws.row_dimensions[3].height = 8

    ws.merge_cells("A4:H4")
    c = ws["A4"]
    c.value = "PRESCRITO"
    cs(c, rojo_oscuro, blanco, bold=True, size=13)
    ws.row_dimensions[4].height = 24

    headers = ["ATENCION"] + LABELS + ["TOTAL GENERAL"]
    for i, h in enumerate(headers, 1):
        c = ws.cell(row=5, column=i, value=h)
        cs(c, rojo_medio, blanco, bold=True, size=10)
    ws.row_dimensions[5].height = 20

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

    anchos = [16, 13, 13, 13, 13, 13, 14, 16]
    for i, w in enumerate(anchos, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output, download_name="prescrito_wms_bata.xlsx",
                     as_attachment=True, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/api/personal", methods=["GET"])
@login_required
def get_personal():
    data = load_data()
    return jsonify(data.get("personal", []))


@app.route("/api/personal", methods=["POST"])
@login_required
@admin_required
def save_personal():
    data = load_data()
    data["personal"] = request.json
    save_data(data)
    return jsonify({"ok": True})


@app.route("/api/personal/excel")
@login_required
def download_personal_excel():
    import openpyxl
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    data = load_data()
    personal = data.get("personal", [])
    semana = request.args.get("semana", "")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Personal"

    rojo = "991B1B"
    rojo_m = "7F1D1D"
    blanco = "FFFFFF"
    negro = "111827"
    gris = "F9FAFB"

    thin = Side(style="thin", color="E5E7EB")
    border = Border(top=thin, bottom=thin, left=thin, right=thin)

    def cs(cell, bg, fg="FFFFFF", bold=False, size=10, align="center"):
        cell.fill = PatternFill("solid", fgColor=bg)
        cell.font = Font(bold=bold, color=fg, size=size, name="Calibri")
        cell.alignment = Alignment(horizontal=align, vertical="center")
        cell.border = border

    ws.merge_cells("A1:M1")
    c = ws["A1"]
    c.value = f"CONTROL DE PERSONAL — {semana}"
    cs(c, rojo, blanco, bold=True, size=13)
    ws.row_dimensions[1].height = 24

    headers = ["N°", "DNI", "APELLIDOS Y NOMBRE", "AREA", "PUESTO", "ENCARGADO",
               "LUN", "MAR", "MIE", "JUE", "VIE", "SAB", "AREA"]
    for i, h in enumerate(headers, 1):
        c = ws.cell(row=2, column=i, value=h)
        cs(c, rojo_m, blanco, bold=True, size=9)
    ws.row_dimensions[2].height = 18

    COLORES_VAL = {
        "SI":         ("D1FAE5", "065F46"),
        "NO":         ("FEE2E2", "991B1B"),
        "VACACIONES": ("FEF3C7", "92400E"),
        "SUSPENDIDO": ("EDE9FE", "5B21B6"),
        "D.M":        ("DBEAFE", "1E40AF"),
        "C.M":        ("DBEAFE", "1E40AF"),
        "CUMPLEAÑOS": ("DBEAFE", "1E40AF"),
    }

    for row_i, p in enumerate(personal, 3):
        bg = gris if row_i % 2 == 0 else blanco
        vals_fijos = [row_i-2, p.get("dni",""), p.get("nombre",""),
                      p.get("area",""), p.get("puesto",""), p.get("encargado","")]
        aligns = ["center","center","left","left","left","left"]
        for ci, (v, al) in enumerate(zip(vals_fijos, aligns), 1):
            cc = ws.cell(row=row_i, column=ci, value=v)
            cs(cc, bg, negro, size=10, align=al)

        dias_keys = [f"{semana}-{i}" for i in range(6)]
        for ci, dk in enumerate(dias_keys, 7):
            val = (p.get("dias") or {}).get(dk, "")
            col_bg, col_fg = COLORES_VAL.get(val.upper() if val else "", (bg, negro))
            cc = ws.cell(row=row_i, column=ci, value=val)
            cs(cc, col_bg, col_fg, bold=bool(val), size=10)

        cc = ws.cell(row=row_i, column=13, value=p.get("area_abrev",""))
        cs(cc, bg, negro, size=10)
        ws.row_dimensions[row_i].height = 18

    anchos = [5, 12, 28, 20, 20, 22, 10, 10, 10, 10, 10, 10, 12]
    for i, w in enumerate(anchos, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output, download_name="personal_wms_bata.xlsx",
                     as_attachment=True, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/api/personal/importar", methods=["POST"])
@login_required
@admin_required
def importar_personal():
    import openpyxl
    from datetime import datetime, timedelta

    if 'archivo' not in request.files:
        return jsonify({"error": "No se envio archivo"}), 400

    archivo = request.files['archivo']
    wb = openpyxl.load_workbook(archivo, data_only=True)
    ws = wb.active

    personal = []
    filas = list(ws.iter_rows(values_only=True))

    header_row = None
    fecha_cols = []
    for i, fila in enumerate(filas):
        fila_str = [str(c).upper() if c else '' for c in fila]
        if 'DNI' in fila_str:
            header_row = i
            for j, val in enumerate(fila):
                if isinstance(val, (int, float)) and 40000 < val < 60000:
                    fecha_cols.append((j, val))
            break

    if header_row is None:
        return jsonify({"error": "No se encontro cabecera DNI"}), 400

    def excel_date(serial):
        try:
            base = datetime(1899, 12, 30)
            return base + timedelta(days=int(serial))
        except:
            return None

    OMITIR = ['RENUNCIO','LI','DESPACHO','NO RETAIL','NOCHE',
              'PASO A LI','PASO A NO RETAIL','WEB','']

    contador = 0
    for fila in filas[header_row+1:]:
        if not fila or not fila[1]:
            continue
        dni      = str(fila[1]).strip() if fila[1] else ''
        nombre   = str(fila[2]).strip() if fila[2] else ''
        area     = str(fila[3]).strip() if fila[3] else ''
        puesto   = str(fila[4]).strip() if fila[4] else ''
        encargado = str(fila[5]).strip() if fila[5] else ''
        area_abrev = str(fila[-1]).strip() if fila[-1] else ''

        if not dni or not nombre:
            continue

        dias = {}
        for col_idx, serial in fecha_cols:
            fecha = excel_date(serial)
            if not fecha:
                continue
            year = fecha.year
            week = fecha.isocalendar()[1]
            dow  = fecha.weekday()
            dia_key = f"{year}-W{str(week).zfill(2)}-{dow}"
            val = str(fila[col_idx]).strip() if col_idx < len(fila) and fila[col_idx] else ''
            if val and val.upper() not in OMITIR:
                dias[dia_key] = val.upper()

        personal.append({
            "id": contador + 1,
            "dni": dni, "nombre": nombre, "area": area,
            "puesto": puesto, "encargado": encargado,
            "area_abrev": area_abrev, "dias": dias
        })
        contador += 1

    data = load_data()
    data["personal"] = personal
    save_data(data)
    return jsonify({"ok": True, "total": contador})


@app.route("/api/wms/procesar", methods=["POST"])
@login_required
@admin_required
def procesar_wms_endpoint():
    from wms_logic import procesar_wms, resumen_kpis
    import datetime

    if 'archivo' not in request.files:
        return jsonify({"error": "No se envio archivo"}), 400

    archivo = request.files['archivo']
    if not archivo.filename.lower().endswith('.xlsx'):
        return jsonify({"error": "El archivo debe ser .xlsx"}), 400

    try:
        df_full, stk_valid = procesar_wms(archivo)
    except ValueError as e:
        return jsonify({"error": str(e)}), 400
    except Exception as e:
        return jsonify({"error": f"Error procesando el archivo: {e}"}), 500

    kpis = resumen_kpis(df_full)

    nuevo_resultado = {
        "fecha": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "archivo": archivo.filename,
        "kpis": kpis,
    }

    data = load_data()
    historial = data.get("wms_historial", [])
    historial.insert(0, nuevo_resultado)
    data["wms_historial"] = historial[:3]
    save_data(data)

    return jsonify({"ok": True, "kpis": kpis})


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)
