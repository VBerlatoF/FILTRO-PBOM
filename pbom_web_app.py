from io import BytesIO
import os
from uuid import uuid4

import pandas as pd
from flask import Flask, redirect, render_template_string, request, send_file, session, url_for

from pbom_core import (
    DEFAULT_FILE,
    build_result,
    build_type_groups,
    excel_column_to_index,
    to_excel_bytes,
)


SHEET_NAME = "COMPLETO"
APP_PASSWORD = os.environ.get("APP_PASSWORD", "")

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev-secret-key-change-on-render")
app.config["MAX_CONTENT_LENGTH"] = int(os.environ.get("MAX_UPLOAD_MB", "200")) * 1024 * 1024

USER_STATES = {}


def new_state():
    return {
        "file_path": DEFAULT_FILE,
        "file_bytes": None,
        "file_label": "",
        "df": None,
        "component_column": "",
        "type_columns": [],
        "type_groups": {},
        "type_keys": [],
        "selected_types": [],
        "result": pd.DataFrame(),
        "message": "Carregue um arquivo .xlsb para iniciar. A aba usada sera sempre COMPLETO.",
        "error": "",
    }


def current_state():
    session_id = session.get("session_id")
    if not session_id:
        session_id = uuid4().hex
        session["session_id"] = session_id
    USER_STATES.setdefault(session_id, new_state())
    return USER_STATES[session_id]


HTML = """
<!doctype html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Filtro de PNs Exclusivos</title>
  <style>
    :root {
      color: #172033;
      background: #edf2f7;
      font-family: Inter, "Segoe UI", Arial, sans-serif;
    }
    * { box-sizing: border-box; }
    body { margin: 0; }
    main {
      width: min(1480px, 100%);
      margin: 0 auto;
      padding: 24px;
    }
    .hero {
      border-radius: 8px;
      background: linear-gradient(135deg, #0f766e, #155e75);
      color: white;
      padding: 24px;
      margin-bottom: 16px;
    }
    h1 {
      margin: 0;
      font-size: 30px;
      line-height: 1.2;
      letter-spacing: 0;
    }
    .subtitle {
      max-width: 760px;
      margin: 8px 0 0;
      color: #dbeafe;
      line-height: 1.5;
    }
    .panel {
      border: 1px solid #d8e0eb;
      border-radius: 8px;
      background: #ffffff;
      padding: 16px;
      margin-bottom: 14px;
      box-shadow: 0 10px 28px rgba(23, 32, 51, 0.06);
    }
    label {
      display: block;
      margin-bottom: 6px;
      font-size: 13px;
      font-weight: 700;
      color: #334155;
    }
    input {
      width: 100%;
      min-height: 40px;
      border: 1px solid #cbd5e1;
      border-radius: 6px;
      padding: 8px 10px;
      font: inherit;
      background: #fff;
    }
    input[type="file"] {
      padding: 7px;
    }
    .file-grid {
      display: grid;
      grid-template-columns: minmax(260px, 1fr) minmax(260px, 1fr) 170px;
      gap: 12px;
      align-items: end;
    }
    .body {
      display: grid;
      grid-template-columns: 340px minmax(0, 1fr);
      gap: 14px;
      align-items: start;
    }
    .types-header {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 10px;
      margin-bottom: 10px;
    }
    .types-header strong {
      color: #172033;
    }
    .types {
      display: grid;
      gap: 8px;
      max-height: 500px;
      overflow: auto;
      padding-right: 4px;
    }
    .type-item {
      display: flex;
      align-items: center;
      gap: 9px;
      min-height: 36px;
      padding: 7px 8px;
      border: 1px solid #e2e8f0;
      border-radius: 6px;
      background: #f8fafc;
      color: #1e293b;
      cursor: pointer;
    }
    .type-item:hover {
      border-color: #38bdf8;
      background: #f0f9ff;
    }
    .type-item input {
      width: auto;
      min-height: auto;
      accent-color: #0f766e;
    }
    .actions {
      display: flex;
      gap: 10px;
      flex-wrap: wrap;
      margin-top: 14px;
    }
    button, .button {
      display: inline-flex;
      min-height: 40px;
      align-items: center;
      justify-content: center;
      border: 1px solid #0f766e;
      border-radius: 6px;
      padding: 0 15px;
      background: #0f766e;
      color: white;
      font: inherit;
      font-weight: 700;
      text-decoration: none;
      cursor: pointer;
    }
    button.secondary, .button.secondary {
      border-color: #cbd5e1;
      background: #fff;
      color: #172033;
    }
    .message, .error {
      border-radius: 8px;
      padding: 11px 13px;
      margin-bottom: 14px;
      border: 1px solid transparent;
    }
    .message {
      background: #ecfeff;
      border-color: #bae6fd;
      color: #155e75;
    }
    .error {
      background: #fef2f2;
      border-color: #fecaca;
      color: #991b1b;
    }
    .metrics {
      display: grid;
      grid-template-columns: repeat(3, minmax(120px, 1fr));
      gap: 10px;
      margin-bottom: 12px;
    }
    .metric {
      border: 1px solid #e2e8f0;
      border-radius: 8px;
      padding: 12px;
      background: #f8fafc;
    }
    .metric span {
      display: block;
      color: #64748b;
      font-size: 12px;
      font-weight: 700;
      text-transform: uppercase;
    }
    .metric strong {
      display: block;
      margin-top: 4px;
      color: #0f172a;
      font-size: 22px;
    }
    .table-wrap {
      overflow: auto;
      max-height: 620px;
      border: 1px solid #d8e0eb;
      border-radius: 8px;
      background: white;
    }
    table {
      width: 100%;
      border-collapse: collapse;
      font-size: 12px;
    }
    th, td {
      border-bottom: 1px solid #e5eaf1;
      padding: 8px 10px;
      text-align: left;
      white-space: nowrap;
      vertical-align: top;
    }
    th {
      position: sticky;
      top: 0;
      background: #e6f4f1;
      color: #173f3b;
      z-index: 1;
    }
    .empty {
      color: #64748b;
      line-height: 1.5;
      margin: 0;
    }
    @media (max-width: 920px) {
      main { padding: 16px; }
      .file-grid, .body, .metrics {
        grid-template-columns: 1fr;
      }
    }
  </style>
</head>
<body>
  <main>
    <section class="hero">
      <h1>Filtro de PNs Exclusivos</h1>
      <p class="subtitle">
        Carregue a PBOM, selecione um ou mais tipos tecnicos completos e gere a lista
        de PNs que possuem X exclusivamente nesses tipos. A aba considerada e sempre COMPLETO.
      </p>
    </section>

    {% if state.error %}
      <div class="error">{{ state.error }}</div>
    {% endif %}
    {% if state.message %}
      <div class="message">{{ state.message }}</div>
    {% endif %}

    <section class="panel">
      <form action="{{ url_for('load_file') }}" method="post" enctype="multipart/form-data" class="file-grid">
        <div>
          <label for="pbom_file">Enviar arquivo .xlsb</label>
          <input id="pbom_file" name="pbom_file" type="file" accept=".xlsb">
        </div>
        <div>
          <label for="file_path">Ou usar caminho de rede/local</label>
          <input id="file_path" name="file_path" value="{{ state.file_path }}">
        </div>
        <button type="submit">Carregar COMPLETO</button>
      </form>
    </section>

    <div class="body">
      <section class="panel">
        <form action="{{ url_for('filter_components') }}" method="post">
          <div class="types-header">
            <strong>Tipos tecnicos</strong>
            <span>{{ state.type_keys|length }} encontrados</span>
          </div>
          <div class="types">
            {% for key in state.type_keys %}
              <label class="type-item">
                <input
                  type="checkbox"
                  name="selected_types"
                  value="{{ key }}"
                  {% if key in state.selected_types %}checked{% endif %}
                >
                <span>{{ key }}</span>
              </label>
            {% endfor %}
          </div>
          <div class="actions">
            <button type="submit">Filtrar PNs</button>
            <a class="button secondary" href="{{ url_for('clear_selection') }}">Limpar</a>
            {% if not state.result.empty %}
              <a class="button" href="{{ url_for('download') }}">Baixar Excel</a>
            {% endif %}
          </div>
        </form>
      </section>

      <section class="panel">
        <div class="metrics">
          <div class="metric">
            <span>PNs exclusivos</span>
            <strong>{{ state.result|length }}</strong>
          </div>
          <div class="metric">
            <span>Tipos selecionados</span>
            <strong>{{ state.selected_types|length }}</strong>
          </div>
          <div class="metric">
            <span>Aba</span>
            <strong>{{ sheet_name }}</strong>
          </div>
        </div>
        {% if not state.result.empty %}
          <div class="table-wrap">
            {{ table_html|safe }}
          </div>
        {% else %}
          <p class="empty">
            Depois de carregar a aba COMPLETO, selecione os tipos tecnicos e clique em Filtrar PNs.
            O resultado trara as informacoes originais do BOM, como Source, Denom. e demais colunas.
          </p>
        {% endif %}
      </section>
    </div>
  </main>
</body>
</html>
"""


LOGIN_HTML = """
<!doctype html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Filtro de PNs Exclusivos - Login</title>
  <style>
    :root {
      color: #172033;
      background: #edf2f7;
      font-family: Inter, "Segoe UI", Arial, sans-serif;
    }
    * { box-sizing: border-box; }
    body {
      display: grid;
      min-height: 100vh;
      place-items: center;
      margin: 0;
      padding: 20px;
    }
    form {
      width: min(420px, 100%);
      border: 1px solid #d8e0eb;
      border-radius: 8px;
      background: #fff;
      padding: 24px;
      box-shadow: 0 10px 28px rgba(23, 32, 51, 0.08);
    }
    h1 {
      margin: 0 0 8px;
      font-size: 24px;
      line-height: 1.2;
    }
    p {
      margin: 0 0 18px;
      color: #64748b;
      line-height: 1.5;
    }
    label {
      display: block;
      margin-bottom: 6px;
      font-size: 13px;
      font-weight: 700;
      color: #334155;
    }
    input {
      width: 100%;
      min-height: 42px;
      border: 1px solid #cbd5e1;
      border-radius: 6px;
      padding: 8px 10px;
      font: inherit;
    }
    button {
      width: 100%;
      min-height: 42px;
      margin-top: 14px;
      border: 1px solid #0f766e;
      border-radius: 6px;
      background: #0f766e;
      color: white;
      font: inherit;
      font-weight: 700;
      cursor: pointer;
    }
    .error {
      border-radius: 6px;
      padding: 10px 12px;
      margin-bottom: 14px;
      background: #fef2f2;
      color: #991b1b;
    }
  </style>
</head>
<body>
  <form method="post" action="{{ url_for('login') }}">
    <h1>Filtro de PNs Exclusivos</h1>
    <p>Informe a senha para acessar a aplicacao.</p>
    {% if error %}
      <div class="error">{{ error }}</div>
    {% endif %}
    <label for="password">Senha</label>
    <input id="password" name="password" type="password" autofocus>
    <button type="submit">Entrar</button>
  </form>
</body>
</html>
"""


@app.before_request
def require_login():
    if not APP_PASSWORD:
        return None
    if request.endpoint == "login":
        return None
    if session.get("authenticated"):
        return None
    return redirect(url_for("login"))


def spreadsheet_source(state):
    if state["file_bytes"] is not None:
        return BytesIO(state["file_bytes"])
    return state["file_path"]


def reset_loaded_data(state):
    state["df"] = None
    state["component_column"] = ""
    state["type_columns"] = []
    state["type_groups"] = {}
    state["type_keys"] = []
    state["selected_types"] = []
    state["result"] = pd.DataFrame()


def set_error(state, message):
    state["error"] = str(message)
    state["message"] = ""


def set_message(state, message):
    state["message"] = str(message)
    state["error"] = ""


def load_complete_sheet(state):
    try:
        excel_file = pd.ExcelFile(spreadsheet_source(state), engine="pyxlsb")
    except Exception as exc:
        raise RuntimeError(f"Nao consegui abrir o arquivo: {exc}") from exc

    if SHEET_NAME not in excel_file.sheet_names:
        sheets = ", ".join(excel_file.sheet_names)
        raise RuntimeError(f"A aba {SHEET_NAME} nao foi encontrada. Abas disponiveis: {sheets}")

    try:
        df = pd.read_excel(spreadsheet_source(state), sheet_name=SHEET_NAME, engine="pyxlsb")
    except Exception as exc:
        raise RuntimeError(f"Nao consegui carregar a aba {SHEET_NAME}: {exc}") from exc

    candidates = [column for column in df.columns if str(column).strip().lower() == "componente"]
    if not candidates:
        raise RuntimeError("Nao encontrei a coluna chamada Componente.")

    start_index = excel_column_to_index("AA")
    end_index = excel_column_to_index("BG")
    type_columns = list(df.columns[start_index : end_index + 1])
    type_groups = build_type_groups(type_columns)

    state["df"] = df
    state["component_column"] = candidates[0]
    state["type_columns"] = type_columns
    state["type_groups"] = type_groups
    state["type_keys"] = list(type_groups.keys())
    state["selected_types"] = []
    state["result"] = pd.DataFrame()


def render_page():
    state = current_state()
    table_html = ""
    if not state["result"].empty:
        table_html = state["result"].to_html(index=False, classes="result-table", border=0)
    return render_template_string(HTML, state=state, table_html=table_html, sheet_name=SHEET_NAME)


@app.get("/")
def index():
    return render_page()


@app.route("/login", methods=["GET", "POST"])
def login():
    if not APP_PASSWORD:
        return redirect(url_for("index"))

    error = ""
    if request.method == "POST":
        password = request.form.get("password", "")
        if password == APP_PASSWORD:
            session["authenticated"] = True
            return redirect(url_for("index"))
        error = "Senha invalida."

    return render_template_string(LOGIN_HTML, error=error)


@app.post("/load-file")
def load_file():
    state = current_state()
    uploaded_file = request.files.get("pbom_file")
    path = request.form.get("file_path", "").strip()

    reset_loaded_data(state)

    if uploaded_file and uploaded_file.filename:
        state["file_bytes"] = uploaded_file.read()
        state["file_label"] = uploaded_file.filename
    else:
        if not path:
            set_error(state, "Envie um arquivo .xlsb ou informe um caminho de rede/local.")
            return redirect(url_for("index"))
        state["file_bytes"] = None
        state["file_path"] = path
        state["file_label"] = path

    try:
        load_complete_sheet(state)
    except RuntimeError as exc:
        set_error(state, exc)
        return redirect(url_for("index"))

    source_name = state["file_label"] or state["file_path"]
    set_message(
        state,
        f"Arquivo carregado: {source_name}. Aba {SHEET_NAME} pronta com "
        f"{len(state['type_keys'])} tipos tecnicos completos."
    )
    return redirect(url_for("index"))


@app.post("/filter")
def filter_components():
    state = current_state()
    selected_types = request.form.getlist("selected_types")
    if state["df"] is None:
        set_error(state, "Carregue o arquivo antes de filtrar.")
        return redirect(url_for("index"))
    if not selected_types:
        set_error(state, "Selecione pelo menos um tipo tecnico.")
        return redirect(url_for("index"))

    try:
        result = build_result(
            state["df"],
            state["component_column"],
            state["type_columns"],
            state["type_groups"],
            selected_types,
        )
    except Exception as exc:
        set_error(state, f"Erro ao filtrar: {exc}")
        return redirect(url_for("index"))

    state["selected_types"] = selected_types
    state["result"] = result
    set_message(state, "Filtro aplicado. O resultado contem os PNs exclusivos e as informacoes originais do BOM.")
    return redirect(url_for("index"))


@app.get("/clear")
def clear_selection():
    state = current_state()
    state["selected_types"] = []
    state["result"] = pd.DataFrame()
    set_message(state, "Selecao limpa.")
    return redirect(url_for("index"))


@app.get("/download")
def download():
    state = current_state()
    if state["result"].empty:
        set_error(state, "Filtre os PNs antes de exportar.")
        return redirect(url_for("index"))

    output = BytesIO(to_excel_bytes(state["result"]))
    return send_file(
        output,
        as_attachment=True,
        download_name="pns_exclusivos_filtrados.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "14006"))
    app.run(host="0.0.0.0", port=port, debug=False)
