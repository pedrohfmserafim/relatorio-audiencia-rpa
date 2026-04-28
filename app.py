import json
import os
import queue
import threading
import traceback
import unicodedata
import uuid
from functools import wraps
from pathlib import Path

from flask import (Flask, jsonify, redirect, render_template,
                   request, Response, send_file, session, url_for)
from openpyxl import load_workbook

import rpa as rpa_module

# ── Mapeamento de colunas ─────────────────────────────────────────────────────
_COL_MAP = {
    "numero do processo":   "numero_processo",
    "número do processo":   "numero_processo",
    "data da audiencia":    "data_audiencia",
    "data da audiência":    "data_audiencia",
    "data audiencia":       "data_audiencia",
    "data audiência":       "data_audiencia",
    # snake_case aliases
    "numero_processo":      "numero_processo",
    "data_audiencia":       "data_audiencia",
}

def _norm(s: str) -> str:
    s = str(s).strip().lower()
    s = unicodedata.normalize("NFD", s)
    return "".join(c for c in s if unicodedata.category(c) != "Mn")


# ── App setup ─────────────────────────────────────────────────────────────────
app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev-secret-troque-em-producao")

UPLOAD_DIR = Path(os.environ.get("UPLOAD_DIR", "uploads"))
OUTPUT_DIR = Path(os.environ.get("OUTPUT_DIR", "output"))
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

APP_USER = os.environ.get("APP_USER", "admin")
APP_PASS = os.environ.get("APP_PASS", "admin")

_log_queue: queue.Queue = queue.Queue()
_state = {
    "running":      False,
    "done":         False,
    "report_ready": False,
    "session_id":   None,
    "paused":       False,
    "results":      [],
}


# ── Auth ──────────────────────────────────────────────────────────────────────
def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get("logged_in"):
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return decorated


@app.route("/login", methods=["GET", "POST"])
def login():
    error = None
    if request.method == "POST":
        if (request.form.get("user") == APP_USER and
                request.form.get("pass") == APP_PASS):
            session["logged_in"] = True
            return redirect(url_for("index"))
        error = "Usuário ou senha incorretos."
    return render_template("login.html", error=error)


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


# ── Routes ────────────────────────────────────────────────────────────────────
@app.route("/")
@login_required
def index():
    return render_template("index.html")


@app.route("/upload", methods=["POST"])
@login_required
def upload():
    file = request.files.get("file")
    if not file or not file.filename.endswith(".xlsx"):
        return jsonify({"erro": "Envie um arquivo .xlsx"}), 400

    session_id = uuid.uuid4().hex
    upload_dir = UPLOAD_DIR / session_id
    upload_dir.mkdir(parents=True, exist_ok=True)

    dest = upload_dir / "planilha.xlsx"
    file.save(dest)

    try:
        rows = _parse_spreadsheet(dest)
    except ValueError as e:
        return jsonify({"erro": str(e)}), 400

    session["upload_session"] = session_id
    preview = [
        {"processo": r["numero_processo"], "data": str(r.get("data_audiencia") or "")}
        for r in rows[:5]
    ]
    return jsonify({"ok": True, "total": len(rows), "preview": preview})


@app.route("/execute", methods=["POST"])
@login_required
def execute():
    if _state["running"]:
        return jsonify({"erro": "Já há uma execução em andamento."}), 400

    sid = session.get("upload_session")
    if not sid:
        return jsonify({"erro": "Faça upload da planilha primeiro."}), 400

    dest = UPLOAD_DIR / sid / "planilha.xlsx"
    if not dest.exists():
        return jsonify({"erro": "Arquivo não encontrado. Faça upload novamente."}), 400

    try:
        rows = _parse_spreadsheet(dest)
    except ValueError as e:
        return jsonify({"erro": str(e)}), 400

    out_dir = OUTPUT_DIR / sid
    out_dir.mkdir(parents=True, exist_ok=True)
    report_path = out_dir / "relatorio_audiencias.xlsx"

    while not _log_queue.empty():
        _log_queue.get_nowait()
    _state.update({"running": True, "done": False, "report_ready": False,
                   "session_id": sid, "paused": False, "results": []})

    thread = threading.Thread(target=_run, args=(rows, report_path), daemon=True)
    thread.start()
    return jsonify({"ok": True, "total": len(rows)})


@app.route("/logs")
@login_required
def logs():
    def generate():
        while True:
            try:
                msg = _log_queue.get(timeout=8)
                yield f"data: {msg}\n\n"
                if json.loads(msg).get("done"):
                    break
            except queue.Empty:
                yield 'data: {"heartbeat":true}\n\n'

    return Response(
        generate(),
        mimetype="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )


@app.route("/status")
@login_required
def status():
    return jsonify({k: v for k, v in _state.items() if k != "results"})


@app.route("/pause", methods=["POST"])
@login_required
def pause():
    _state["paused"] = True
    return jsonify({"ok": True})


@app.route("/resume", methods=["POST"])
@login_required
def resume():
    _state["paused"] = False
    return jsonify({"ok": True})


@app.route("/partial-report")
@login_required
def partial_report():
    results = _state.get("results", [])
    if not results:
        return jsonify({"erro": "Nenhum resultado disponível ainda"}), 404
    path = Path("/tmp/relatorio_audiencias_parcial.xlsx")
    rpa_module._build_report(results, path)
    return send_file(path, as_attachment=True,
                     download_name="relatorio_audiencias_parcial.xlsx")


@app.route("/download")
@login_required
def download():
    sid = _state.get("session_id") or session.get("upload_session")
    if not sid:
        return jsonify({"erro": "Relatório não disponível"}), 404
    path = OUTPUT_DIR / sid / "relatorio_audiencias.xlsx"
    if not path.exists():
        return jsonify({"erro": "Relatório ainda não gerado"}), 404
    return send_file(path, as_attachment=True,
                     download_name="relatorio_audiencias.xlsx")


@app.route("/debug-screenshot")
@login_required
def debug_screenshot():
    path = Path("/tmp/debug_login.png")
    if not path.exists():
        return jsonify({"erro": "Nenhum screenshot disponível ainda"}), 404
    return send_file(path, mimetype="image/png")


# ── Error handlers ────────────────────────────────────────────────────────────
@app.errorhandler(404)
def not_found(e):
    return jsonify({"erro": "Rota não encontrada"}), 404


@app.errorhandler(500)
def internal_error(e):
    return jsonify({"erro": "Erro interno do servidor", "detalhe": str(e)}), 500


# ── Internals ─────────────────────────────────────────────────────────────────
def _parse_spreadsheet(path: Path) -> list[dict]:
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    rows_iter = ws.iter_rows(values_only=True)
    raw_headers = next(rows_iter, None)
    if raw_headers is None:
        raise ValueError("Planilha vazia")

    mapped_headers = [_COL_MAP.get(_norm(h or "")) for h in raw_headers]

    present  = {k for k in mapped_headers if k}
    required = {"numero_processo", "data_audiencia"}
    missing  = required - present
    if missing:
        raise ValueError(
            f"Coluna(s) ausente(s): {', '.join(sorted(missing))}. "
            "Esperado: 'Número do Processo' e 'Data da Audiência'."
        )

    records = []
    vazias  = 0
    for row in rows_iter:
        data = {k: v for k, v in zip(mapped_headers, row) if k}
        if not any(data.values()):
            vazias += 1
            if vazias >= 3:
                break
            continue
        vazias = 0
        if not data.get("numero_processo"):
            continue
        # Normaliza data para string DD/MM/YYYY
        raw_date = data.get("data_audiencia")
        if hasattr(raw_date, "strftime"):
            data["data_audiencia"] = raw_date.strftime("%d/%m/%Y")
        else:
            data["data_audiencia"] = str(raw_date or "").strip()
        records.append(data)

    wb.close()
    if not records:
        raise ValueError("Nenhum dado válido encontrado na planilha")
    return records


def _log(msg: str, status: str | None = None):
    _log_queue.put(json.dumps({"msg": msg, "status": status}))


def _run(rows: list[dict], report_path: Path):
    try:
        rpa_module.run_automation(rows, _log, report_path, _state)
        _state["report_ready"] = True
    except Exception:
        _log(f"Erro crítico:\n{traceback.format_exc()}", "error")
    finally:
        _state["running"] = False
        _state["done"]    = True
        _log_queue.put(json.dumps({"done": True}))


# ── Entry point ───────────────────────────────────────────────────────────────
if __name__ == "__main__":
    print("Relatório Audiência RPA → http://localhost:5002")
    app.run(host="127.0.0.1", port=5002, debug=False, threaded=True)
