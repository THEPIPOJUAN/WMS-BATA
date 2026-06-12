from flask import Flask, render_template, request, jsonify, session, redirect, url_for
from functools import wraps
import json, os

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
            return jsonify({"error": "Sin permisos de edición"}), 403
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
        error = "Usuario o contraseña incorrectos"
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

if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)
