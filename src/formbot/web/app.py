from __future__ import annotations

import json
import logging
import tempfile
from pathlib import Path
from typing import Any, Final
from uuid import uuid4

from fastapi import FastAPI, File, Form, UploadFile
from fastapi.responses import HTMLResponse, JSONResponse, Response

from formbot.domain.exceptions import FormBotError

LOGGER = logging.getLogger(__name__)

# Raíz del proyecto: src/formbot/web/app.py → parents[3] = raíz del repo
_PROJECT_ROOT: Final[Path] = Path(__file__).resolve().parents[3]
_PROFILE_PATH: Final[Path] = _PROJECT_ROOT / "config" / "data" / "asteco_master_profile.json"

app = FastAPI(title="FormBot", version="2.0.0")

# ---------------------------------------------------------------------------
# Constantes de formato
# ---------------------------------------------------------------------------
SUPPORTED_EXTENSIONS: dict[str, str] = {
    ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    ".xlsm": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    ".pdf":  "application/pdf",
    ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
}

# ---------------------------------------------------------------------------
# Hints de perfil: fragmento de etiqueta → clave en el perfil maestro
# ---------------------------------------------------------------------------
_PROFILE_HINTS: dict[str, str] = {
    "razon social": "razon_social",
    "nombre comercial": "nombre_comercial",
    "numero de identificacion": "numero_identificacion_nit",
    "numero identificacion": "numero_identificacion_nit",
    "identificacion tributaria": "numero_identificacion_nit",
    "nit": "numero_identificacion_nit",
    "digito de verificacion": "digito_verificacion",
    "digito verificacion": "digito_verificacion",
    "verificacion": "digito_verificacion",
    "representante legal": "representante_legal_nombre",
    "nombre representante": "representante_legal_nombre",
    "nombres y apellidos": "representante_legal_nombre",
    "documento representante": "representante_legal_documento",
    "identificacion representante": "representante_legal_documento",
    "direccion": "direccion_principal",
    "ciudad": "ciudad_municipio",
    "municipio": "ciudad_municipio",
    "departamento": "departamento",
    "pais": "pais",
    "telefono fijo": "telefono_fijo",
    "telefono": "telefono_fijo",
    "telefax": "telefono_fijo",
    "celular": "celular",
    "movil": "celular",
    "correo electronico": "correo_electronico",
    "correo": "correo_electronico",
    "email": "correo_electronico",
    "contacto": "contacto_nombre",
    "nombre del contacto": "contacto_nombre",
    "banco": "banco_nombre",
    "entidad bancaria": "banco_nombre",
    "tipo de cuenta": "tipo_cuenta",
    "tipo cuenta": "tipo_cuenta",
    "numero de cuenta": "numero_cuenta",
    "numero cuenta": "numero_cuenta",
    "titular": "titular_cuenta",
    "titular cuenta": "titular_cuenta",
}


# ---------------------------------------------------------------------------
# Carga del perfil maestro
# ---------------------------------------------------------------------------

def _load_master_profile() -> dict[str, Any]:
    """Carga el perfil maestro de datos de la empresa. Retorna {} si no existe."""
    if not _PROFILE_PATH.exists():
        return {}
    try:
        return json.loads(_PROFILE_PATH.read_text(encoding="utf-8"))
    except Exception:
        return {}


def _suggest_from_profile(label: str, profile: dict[str, Any]) -> str | None:
    """Intenta encontrar un valor sugerido en el perfil para la etiqueta dada."""
    if not profile:
        return None

    from formbot.shared.utils import normalize_text
    normalized = normalize_text(label)

    # 1. Match directo: slug de la etiqueta coincide con una clave del perfil
    slug = "_".join(normalized.split())
    if slug in profile:
        return str(profile[slug])

    # 2. Hints manuales: fragmento conocido dentro de la etiqueta normalizada
    for hint, key in _PROFILE_HINTS.items():
        if hint in normalized and key in profile:
            return str(profile[key])

    # 3. Substring inverso: clave normalizada del perfil contenida en la etiqueta
    for key, value in profile.items():
        key_norm = normalize_text(key.replace("_", " "))
        if key_norm and (key_norm in normalized or normalized in key_norm):
            return str(value)

    return None


# ---------------------------------------------------------------------------
# HTML de la interfaz (2 pasos)
# ---------------------------------------------------------------------------

INDEX_HTML: Final[str] = """<!doctype html>
<html lang="es">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>FormBot | Diligenciamiento Automatico</title>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@400;500;700&family=IBM+Plex+Mono:wght@400;600&display=swap" rel="stylesheet">
  <style>
    :root {
      --ink: #13233f;
      --paper: #f4f8ff;
      --accent: #ff6b35;
      --accent-2: #00bcd4;
      --good: #0f8b5f;
      --bad: #b63035;
      --card: #ffffff;
      --line: #dce6f6;
      --muted: #7a92b8;
      --excel: #1d6f42;
      --pdf: #b63035;
      --word: #185abd;
      --suggested: #e8f7f0;
      --suggested-border: #b8ecd9;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      font-family: "Space Grotesk", sans-serif;
      color: var(--ink);
      background:
        radial-gradient(1200px 600px at 10% -20%, #ffd6c7 0%, transparent 60%),
        radial-gradient(1000px 500px at 95% 0%, #c6f7ff 0%, transparent 55%),
        var(--paper);
      min-height: 100vh;
    }
    .wrap { max-width: 900px; margin: 0 auto; padding: 36px 20px 64px; }

    /* Hero */
    .hero { display: grid; gap: 8px; margin-bottom: 24px; animation: rise 480ms ease-out; }
    .badge {
      display: inline-flex; align-items: center; gap: 6px;
      width: fit-content;
      font-family: "IBM Plex Mono", monospace; font-size: 11px;
      letter-spacing: .08em; text-transform: uppercase;
      background: #13233f; color: #fff;
      padding: 5px 12px; border-radius: 999px;
    }
    h1 { margin: 0; font-size: clamp(24px, 4vw, 38px); line-height: 1.1; }
    .sub { margin: 0; max-width: 60ch; color: #2f4469; font-size: 14px; line-height: 1.55; }

    /* Formatos */
    .formats { display: flex; flex-wrap: wrap; gap: 8px; margin-bottom: 20px; }
    .fmt-chip {
      display: inline-flex; align-items: center; gap: 6px;
      padding: 4px 11px; border-radius: 999px;
      font-size: 11px; font-weight: 700; letter-spacing: .03em;
      border: 1.5px solid currentColor; opacity: .8;
    }
    .fmt-chip.excel { color: var(--excel); background: #edfaf3; }
    .fmt-chip.pdf   { color: var(--pdf);   background: #fff4f5; }
    .fmt-chip.word  { color: var(--word);  background: #eef4ff; }
    .fmt-dot { width: 7px; height: 7px; border-radius: 50%; background: currentColor; }

    /* Card */
    .card {
      background: var(--card); border: 1px solid var(--line);
      border-radius: 22px; padding: 28px 28px 24px;
      box-shadow: 0 14px 48px rgba(25,54,106,.09);
      animation: rise 560ms ease-out;
    }

    /* Upload zone */
    .upload-zone {
      display: flex; flex-direction: column; align-items: center;
      justify-content: center; gap: 14px;
      border: 2px dashed #bfd1ee; border-radius: 16px;
      padding: 40px 20px; cursor: pointer;
      background: #f9fbff;
      transition: border-color .2s, background .2s;
      text-align: center;
    }
    .upload-zone.dragover {
      border-color: var(--accent); background: #fff6f3;
    }
    .upload-zone.has-file {
      border-color: #6fa8dc; background: #f0f7ff; border-style: solid;
    }
    .upload-icon { font-size: 40px; line-height: 1; }
    .upload-hint { font-size: 13px; color: var(--muted); font-family: "IBM Plex Mono", monospace; }
    .file-name {
      font-family: "IBM Plex Mono", monospace; font-size: 12px;
      color: var(--ink); font-weight: 600; word-break: break-all;
    }
    #file-input { display: none; }

    /* Format pill */
    .fmt-pill {
      display: inline-block;
      font-size: 10px; font-weight: 700; letter-spacing: .06em;
      text-transform: uppercase; padding: 3px 9px; border-radius: 999px;
    }
    .fmt-pill.excel { background: #d4f3e3; color: var(--excel); }
    .fmt-pill.pdf   { background: #fde8e8; color: var(--pdf); }
    .fmt-pill.word  { background: #dce9ff; color: var(--word); }

    /* Buttons */
    .actions { margin-top: 20px; display: flex; flex-wrap: wrap; gap: 10px; align-items: center; }
    button {
      border: none; border-radius: 12px; padding: 11px 22px;
      font-size: 14px; font-weight: 700; cursor: pointer;
      font-family: "Space Grotesk", sans-serif;
      transition: transform .14s ease, filter .18s ease, opacity .15s;
    }
    .btn-primary { background: linear-gradient(135deg, var(--accent), #ff934f); color: #1f1309; }
    .btn-secondary { background: linear-gradient(135deg, var(--accent-2), #8ce7f2); color: #05343b; }
    .btn-ghost {
      background: transparent; border: 1.5px solid var(--line);
      color: var(--ink); font-size: 13px;
    }
    button:hover { transform: translateY(-1px); filter: brightness(1.04); }
    button:active { transform: translateY(0); filter: brightness(.97); }
    button:disabled { opacity: .5; cursor: not-allowed; transform: none; filter: none; }

    /* Status / progress */
    .status {
      margin-top: 14px; min-height: 40px;
      padding: 10px 14px; border-radius: 11px;
      font-family: "IBM Plex Mono", monospace; font-size: 12px;
      background: #edf3ff; border: 1px solid #c7d8f5;
      white-space: pre-wrap; word-break: break-word; line-height: 1.5;
    }
    .status.ok  { color: var(--good); border-color: #b8ecd9; background: #f0fbf5; }
    .status.err { color: var(--bad);  border-color: #f0bcc0; background: #fff3f4; }
    .status.run { color: #4a6ea8; border-color: #a8c0e8; background: #edf3ff; }
    .progress-bar {
      height: 3px; border-radius: 99px; margin-top: 8px; display: none;
      background: linear-gradient(90deg, var(--accent), var(--accent-2));
      background-size: 200% 100%;
      animation: indeterminate 1.4s linear infinite;
    }
    @keyframes indeterminate {
      0%   { background-position: 200% 0; }
      100% { background-position: -200% 0; }
    }

    /* ── Step 2: Review fields ── */
    #step2 { display: none; }
    .step2-header {
      display: flex; align-items: center; justify-content: space-between;
      flex-wrap: wrap; gap: 12px; margin-bottom: 20px;
    }
    .step2-title { font-size: 17px; font-weight: 700; }
    .step2-meta {
      font-family: "IBM Plex Mono", monospace; font-size: 11px;
      color: var(--muted);
    }

    /* Fields grid */
    .fields-grid {
      display: grid;
      grid-template-columns: repeat(auto-fill, minmax(280px, 1fr));
      gap: 12px;
      max-height: 65vh; overflow-y: auto;
      padding-right: 4px; margin-bottom: 16px;
    }
    .field-card {
      background: #f9fbff; border: 1.5px solid var(--line);
      border-radius: 14px; padding: 14px;
      display: flex; flex-direction: column; gap: 6px;
      transition: border-color .2s;
    }
    .field-card.has-value {
      border-color: var(--suggested-border); background: var(--suggested);
    }
    .field-card-label {
      font-size: 12px; font-weight: 700; letter-spacing: .02em;
      color: #1e3359; word-break: break-word;
    }
    .field-card-key {
      font-family: "IBM Plex Mono", monospace; font-size: 10px;
      color: var(--muted); letter-spacing: .03em;
    }
    .field-card input[type="text"] {
      width: 100%; padding: 8px 10px;
      font-size: 13px; font-family: "Space Grotesk", sans-serif;
      border: 1.5px solid #d0dff5; border-radius: 8px;
      background: #fff; color: var(--ink);
      transition: border-color .18s;
      outline: none;
    }
    .field-card input[type="text"]:focus { border-color: var(--accent-2); }
    .field-card input[type="text"].from-profile {
      border-color: #a8d8c0; background: #f5fdf9;
    }

    /* Empty state */
    .no-fields {
      text-align: center; padding: 40px 20px;
      color: var(--muted); font-size: 14px;
    }

    /* Divider */
    .divider {
      border: none; border-top: 1px solid var(--line); margin: 20px 0;
    }

    @media (max-width: 600px) {
      .card { padding: 20px 16px; }
      .fields-grid { grid-template-columns: 1fr; }
    }
    @keyframes rise {
      from { opacity: 0; transform: translateY(10px); }
      to   { opacity: 1; transform: translateY(0); }
    }
  </style>
</head>
<body>
<div class="wrap">

  <section class="hero">
    <span class="badge">&#9679; FormBot</span>
    <h1>Diligencia formularios<br>automaticamente</h1>
    <p class="sub">Carga el documento, revisa los campos detectados y descarga el resultado.</p>
  </section>

  <div class="formats">
    <span class="fmt-chip excel"><span class="fmt-dot"></span>Excel .xlsx / .xlsm</span>
    <span class="fmt-chip pdf"><span class="fmt-dot"></span>PDF AcroForm .pdf</span>
    <span class="fmt-chip word"><span class="fmt-dot"></span>Word .docx</span>
  </div>

  <!-- ── PASO 1: Subir plantilla ── -->
  <section class="card" id="step1">
    <div class="upload-zone" id="drop-zone">
      <div class="upload-icon">&#128196;</div>
      <div>
        <strong>Arrastra el formulario aquí</strong><br>
        <span class="upload-hint">o haz clic para seleccionar</span>
      </div>
      <span class="upload-hint">.xlsx &middot; .xlsm &middot; .pdf &middot; .docx</span>
      <input id="file-input" type="file" accept=".xlsx,.xlsm,.pdf,.docx" />
    </div>
    <div id="file-info" style="display:none; margin-top:12px; display:none; align-items:center; gap:10px;">
      <span id="fmt-pill" class="fmt-pill"></span>
      <span id="file-name-text" class="file-name"></span>
    </div>

    <div class="actions">
      <button class="btn-primary" id="analyze-btn" onclick="analyzeTemplate()" disabled>
        &#128269; Analizar documento
      </button>
    </div>

    <div class="progress-bar" id="progress1"></div>
    <div class="status" id="status1">Carga un formulario para comenzar.</div>
  </section>

  <!-- ── PASO 2: Revisar y diligenciar ── -->
  <section class="card" id="step2" style="margin-top:18px;">
    <div class="step2-header">
      <div>
        <div class="step2-title" id="step2-title">Campos detectados</div>
        <div class="step2-meta" id="step2-meta"></div>
      </div>
      <button class="btn-ghost" onclick="resetToStep1()">&#8592; Cargar otro</button>
    </div>

    <div class="fields-grid" id="fields-grid">
      <!-- renderizado por JS -->
    </div>

    <hr class="divider">

    <div class="actions">
      <button class="btn-primary" id="fill-btn" onclick="fillDocument()">
        &#9654; Diligenciar y Descargar
      </button>
      <button class="btn-ghost" onclick="clearValues()">&#10005; Limpiar valores</button>
    </div>

    <div class="progress-bar" id="progress2"></div>
    <div class="status" id="status2" style="display:none"></div>
  </section>

</div>

<script>
  /* ── Estado ── */
  let currentFile = null;
  let detectedFields = [];

  const dropZone   = document.getElementById("drop-zone");
  const fileInput  = document.getElementById("file-input");
  const analyzeBtn = document.getElementById("analyze-btn");
  const fmtPill    = document.getElementById("fmt-pill");
  const fileNameEl = document.getElementById("file-name-text");
  const fileInfo   = document.getElementById("file-info");
  const step1      = document.getElementById("step1");
  const step2      = document.getElementById("step2");
  const fieldsGrid = document.getElementById("fields-grid");
  const fillBtn    = document.getElementById("fill-btn");
  const prog1      = document.getElementById("progress1");
  const prog2      = document.getElementById("progress2");
  const status1    = document.getElementById("status1");
  const status2    = document.getElementById("status2");

  const FORMAT_MAP = {
    ".xlsx": { label: "Excel",      cls: "excel" },
    ".xlsm": { label: "Excel Macro",cls: "excel" },
    ".pdf":  { label: "PDF",        cls: "pdf"   },
    ".docx": { label: "Word",       cls: "word"  },
  };

  /* ── Drag & drop ── */
  dropZone.addEventListener("click", () => fileInput.click());
  dropZone.addEventListener("dragover",  e => { e.preventDefault(); dropZone.classList.add("dragover"); });
  dropZone.addEventListener("dragleave", ()=> dropZone.classList.remove("dragover"));
  dropZone.addEventListener("drop", e => {
    e.preventDefault(); dropZone.classList.remove("dragover");
    const file = e.dataTransfer.files[0];
    if (file) setFile(file);
  });
  fileInput.addEventListener("change", () => {
    if (fileInput.files[0]) setFile(fileInput.files[0]);
  });

  function setFile(file) {
    currentFile = file;
    const ext = file.name.slice(file.name.lastIndexOf(".")).toLowerCase();
    const info = FORMAT_MAP[ext];
    dropZone.classList.add("has-file");
    fileInfo.style.display = "flex";
    fileNameEl.textContent = file.name;
    if (info) {
      fmtPill.textContent = info.label;
      fmtPill.className   = "fmt-pill " + info.cls;
      fmtPill.style.display = "inline-block";
    } else {
      fmtPill.style.display = "none";
    }
    analyzeBtn.disabled = false;
    setStatus1("Listo para analizar: " + file.name);
  }

  /* ── Status helpers ── */
  function setStatus1(msg, tone = "") {
    status1.textContent = msg;
    status1.className   = "status" + (tone ? " " + tone : "");
  }
  function setStatus2(msg, tone = "") {
    status2.style.display = "block";
    status2.textContent   = msg;
    status2.className     = "status" + (tone ? " " + tone : "");
  }

  /* ── Paso 1 → Analizar ── */
  async function analyzeTemplate() {
    if (!currentFile) return;
    analyzeBtn.disabled = true;
    prog1.style.display = "block";
    setStatus1("Analizando documento...", "run");

    const body = new FormData();
    body.append("template", currentFile);

    try {
      const resp = await fetch("/api/analyze", { method: "POST", body });
      const data = await resp.json();
      if (!resp.ok) throw new Error(data.detail || "Error " + resp.status);

      detectedFields = data.fields;
      renderStep2(data);
      step2.style.display = "block";
      step2.scrollIntoView({ behavior: "smooth", block: "start" });
      setStatus1(
        detectedFields.length + " campo(s) detectados. Revisa y completa los valores abajo.", "ok"
      );
    } catch (err) {
      setStatus1("Error: " + err.message, "err");
    } finally {
      analyzeBtn.disabled = false;
      prog1.style.display = "none";
    }
  }

  /* ── Renderizar paso 2 ── */
  function renderStep2(data) {
    document.getElementById("step2-title").textContent =
      data.fields.length + " campo(s) detectados";
    document.getElementById("step2-meta").textContent =
      currentFile.name + (data.sheet_count ? "  ·  " + data.sheet_count + " hoja(s)" : "");

    fieldsGrid.innerHTML = "";
    if (data.fields.length === 0) {
      fieldsGrid.innerHTML =
        '<div class="no-fields">No se detectaron campos rellenables en este documento.</div>';
      return;
    }

    data.fields.forEach(field => {
      const card = document.createElement("div");
      const hasSuggestion = Boolean(field.suggested_value);
      card.className = "field-card" + (hasSuggestion ? " has-value" : "");
      card.innerHTML =
        '<div class="field-card-label">' + escHtml(field.label) + "</div>" +
        '<div class="field-card-key">' + escHtml(field.field_key) + "</div>" +
        '<input type="text"' +
          ' name="' + escAttr(field.field_key) + '"' +
          ' data-label="' + escAttr(field.label) + '"' +
          ' value="' + escAttr(field.suggested_value || "") + '"' +
          ' class="' + (hasSuggestion ? "from-profile" : "") + '"' +
          ' placeholder="Dejar vacío para omitir"' +
        " />";
      fieldsGrid.appendChild(card);
    });
  }

  /* ── Paso 2 → Diligenciar y descargar ── */
  async function fillDocument() {
    const inputs = fieldsGrid.querySelectorAll("input[type='text']");
    const fields = [];
    inputs.forEach(inp => {
      fields.push({
        field_key: inp.name,
        label:     inp.dataset.label,
        value:     inp.value,
      });
    });

    fillBtn.disabled = true;
    prog2.style.display = "block";
    setStatus2("Diligenciando documento...", "run");

    const body = new FormData();
    body.append("template", currentFile);
    body.append("fields",   JSON.stringify(fields));

    try {
      const resp = await fetch("/api/fill-smart", { method: "POST", body });
      if (!resp.ok) {
        const json = await resp.json().catch(() => ({}));
        throw new Error(json.detail || "Error " + resp.status);
      }

      const blob     = await resp.blob();
      const filename = extractFilename(resp.headers.get("content-disposition"));
      const url      = URL.createObjectURL(blob);
      const a        = document.createElement("a");
      a.href = url; a.download = filename;
      document.body.appendChild(a); a.click(); a.remove();
      URL.revokeObjectURL(url);

      setStatus2("Listo. Descargando: " + filename, "ok");
    } catch (err) {
      setStatus2("Error: " + err.message, "err");
    } finally {
      fillBtn.disabled = false;
      prog2.style.display = "none";
    }
  }

  /* ── Limpiar valores ── */
  function clearValues() {
    fieldsGrid.querySelectorAll("input[type='text']").forEach(inp => { inp.value = ""; });
  }

  /* ── Volver al paso 1 ── */
  function resetToStep1() {
    step2.style.display = "none";
    status2.style.display = "none";
    detectedFields = [];
    fieldsGrid.innerHTML = "";
    setStatus1("Carga un formulario para comenzar.");
  }

  /* ── Helpers ── */
  function extractFilename(header) {
    if (!header) return "formulario_diligenciado";
    const utf = header.match(/filename\\*=UTF-8''([^;]+)/i);
    if (utf && utf[1]) return decodeURIComponent(utf[1]);
    const ascii = header.match(/filename="?([^";]+)"?/i);
    return (ascii && ascii[1]) ? ascii[1] : "formulario_diligenciado";
  }
  function escHtml(s) {
    return String(s)
      .replace(/&/g, "&amp;").replace(/</g, "&lt;")
      .replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  }
  function escAttr(s) { return escHtml(s); }
</script>
</body>
</html>
"""


# ---------------------------------------------------------------------------
# Rutas
# ---------------------------------------------------------------------------

@app.get("/", response_class=HTMLResponse)
def index() -> HTMLResponse:
    return HTMLResponse(INDEX_HTML)


@app.post("/api/analyze")
async def analyze_template(
    template: UploadFile = File(...),
) -> JSONResponse:
    """Analiza un documento y retorna los campos rellenables detectados."""
    from formbot.infrastructure.document_scanners.field_scanner import scan_document

    with tempfile.TemporaryDirectory(prefix="formbot-analyze-") as tmpdir:
        tmp = Path(tmpdir)
        template_path = tmp / _safe_filename(template.filename, "template.xlsx")
        template_path.write_bytes(await template.read())

        suffix = template_path.suffix.lower()
        if suffix not in SUPPORTED_EXTENSIONS:
            return JSONResponse(
                status_code=400,
                content={"detail": f"Formato '{suffix}' no soportado."},
            )

        try:
            detected = scan_document(template_path)
        except Exception as exc:
            LOGGER.exception("Error escaneando documento")
            return JSONResponse(
                status_code=500,
                content={"detail": f"Error al analizar el documento: {exc}"},
            )

        profile = _load_master_profile()
        fields_payload = []
        for field in detected:
            suggested = _suggest_from_profile(field.label, profile)
            fields_payload.append({
                "field_key":       field.field_key,
                "label":           field.label,
                "sheet":           field.sheet,
                "suggested_value": suggested or "",
            })

        # Contar hojas únicas (para Excel)
        sheet_names = {f.sheet for f in detected if f.sheet}
        return JSONResponse({
            "format":      suffix.lstrip("."),
            "sheet_count": len(sheet_names),
            "fields":      fields_payload,
        })


@app.post("/api/fill-smart")
async def fill_smart(
    template: UploadFile = File(...),
    fields: str = Form(...),
) -> Response:
    """Diligencia el documento usando el mapeo auto-detectado.

    fields: JSON list de objetos {field_key, label, value}.
    Los campos con value vacío se omiten automáticamente.
    """
    with tempfile.TemporaryDirectory(prefix="formbot-smart-") as tmpdir:
        tmp = Path(tmpdir)
        template_path = tmp / _safe_filename(template.filename, "template.xlsx")
        template_path.write_bytes(await template.read())

        suffix = template_path.suffix.lower()
        mime_type = SUPPORTED_EXTENSIONS.get(suffix, "application/octet-stream")
        output_filename = (
            f"{template_path.stem}_diligenciado_{uuid4().hex[:8]}{suffix}"
        )
        output_path = tmp / output_filename

        try:
            field_list: list[dict] = json.loads(fields)
        except Exception:
            return JSONResponse(status_code=400, content={"detail": "JSON de campos inválido."})

        # Solo procesar campos con valor no vacío
        active = [f for f in field_list if str(f.get("value", "")).strip()]
        if not active:
            return JSONResponse(
                status_code=400,
                content={"detail": "No se proporcionaron valores para diligenciar."},
            )

        try:
            if suffix in {".xlsx", ".xlsm"}:
                _fill_excel_smart(template_path, active, output_path)
            elif suffix == ".pdf":
                _fill_via_adapter_smart(template_path, active, output_path, col_offset=0, row_offset=0)
            elif suffix == ".docx":
                _fill_via_adapter_smart(template_path, active, output_path, col_offset=1, row_offset=0)
            else:
                return JSONResponse(
                    status_code=400,
                    content={"detail": f"Formato '{suffix}' no soportado."},
                )
        except FormBotError as exc:
            return JSONResponse(status_code=400, content={"detail": str(exc)})
        except Exception as exc:
            LOGGER.exception("Error en fill-smart")
            return JSONResponse(
                status_code=500,
                content={"detail": f"Error al diligenciar: {type(exc).__name__}: {exc}"},
            )

        payload = output_path.read_bytes()
        headers = {"Content-Disposition": f'attachment; filename="{output_filename}"'}
        return Response(content=payload, media_type=mime_type, headers=headers)


@app.post("/api/fill")
async def fill_form(
    template: UploadFile = File(...),
    mapping: UploadFile = File(...),
    data: UploadFile = File(...),
) -> Response:
    """Endpoint legado: requiere template + mapping YAML + payload JSON."""
    from formbot.app.bootstrap import bootstrap_pipeline

    with tempfile.TemporaryDirectory(prefix="formbot-web-") as tmpdir:
        tmp = Path(tmpdir)
        template_path = tmp / _safe_filename(template.filename, "template.xlsx")
        mapping_path  = tmp / _safe_filename(mapping.filename, "mapping.yaml")
        data_path     = tmp / _safe_filename(data.filename, "data.json")

        template_path.write_bytes(await template.read())
        mapping_path.write_bytes(await mapping.read())
        data_path.write_bytes(await data.read())

        output_suffix   = template_path.suffix.lower() or ".xlsx"
        output_filename = f"{template_path.stem}_filled_{uuid4().hex[:8]}{output_suffix}"
        output_path     = tmp / output_filename

        context = None
        try:
            context = bootstrap_pipeline(
                template_path=template_path,
                mapping_path=mapping_path,
                data_path=data_path,
            )
            context.use_case.execute(
                data=context.data,
                mapping_rules=context.mapping_rules,
                output_path=output_path,
            )
        except FormBotError as exc:
            return JSONResponse(status_code=400, content={"detail": f"{type(exc).__name__}: {exc}"})
        except Exception as exc:
            return JSONResponse(
                status_code=500,
                content={"detail": f"Error no controlado: {type(exc).__name__}: {exc}"},
            )
        finally:
            if context is not None:
                context.use_case.close()
            await template.close()
            await mapping.close()
            await data.close()

        payload = output_path.read_bytes()
        headers = {"Content-Disposition": f'attachment; filename="{output_filename}"'}
        return Response(
            content=payload,
            media_type=context.mime_type,  # type: ignore[union-attr]
            headers=headers,
        )


# ---------------------------------------------------------------------------
# Helpers de diligenciamiento inteligente
# ---------------------------------------------------------------------------

def _fill_excel_smart(template_path: Path, fields: list[dict], output_path: Path) -> None:
    """Excel: usa PrecisionFillUseCase con inferencia automática de celda destino."""
    from formbot.application.precision_fill import PrecisionFillUseCase
    from formbot.domain.models import MappingRule

    rules = []
    data: dict[str, Any] = {}
    seen_keys: set[str] = set()

    for f in fields:
        key   = f["field_key"]
        label = f["label"]
        value = f["value"]

        # Garantizar unicidad de field_name
        if key in seen_keys:
            key = key + "_" + uuid4().hex[:4]
        seen_keys.add(key)

        rules.append(MappingRule(
            field_name=key,
            label=label,
            row_offset=0,
            column_offset=0,
            required=False,
            target_strategy="offset_or_infer",
        ))
        data[key] = value

    use_case = PrecisionFillUseCase(
        template_path=template_path,
        strict_mode=False,
        min_confidence=0.45,
        allow_overwrite_existing=False,
    )
    try:
        use_case.execute(data=data, mapping_rules=rules, output_path=output_path)
    finally:
        use_case.close()


def _fill_via_adapter_smart(
    template_path: Path,
    fields: list[dict],
    output_path: Path,
    *,
    col_offset: int,
    row_offset: int,
) -> None:
    """PDF / Word: usa FillFormUseCase con el adaptador correspondiente."""
    from formbot.app.bootstrap import create_document_adapter
    from formbot.application.fill_form import FillFormUseCase
    from formbot.domain.models import MappingRule
    from formbot.infrastructure.mappers.label_offset_mapper import LabelOffsetMapper

    adapter  = create_document_adapter(template_path)
    mapper   = LabelOffsetMapper()
    use_case = FillFormUseCase(document_adapter=adapter, field_mapper=mapper)

    rules = []
    data: dict[str, Any] = {}
    seen_keys: set[str] = set()

    for f in fields:
        key   = f["field_key"]
        label = f["label"]
        value = f["value"]

        if key in seen_keys:
            key = key + "_" + uuid4().hex[:4]
        seen_keys.add(key)

        rules.append(MappingRule(
            field_name=key,
            label=label,
            row_offset=row_offset,
            column_offset=col_offset,
            required=False,
            target_strategy="offset",
        ))
        data[key] = value

    try:
        use_case.execute(data=data, mapping_rules=rules, output_path=output_path)
    finally:
        use_case.close()


# ---------------------------------------------------------------------------
# Utilidades
# ---------------------------------------------------------------------------

def _safe_filename(candidate: str | None, fallback: str) -> str:
    if not candidate:
        return fallback
    name = Path(candidate).name.strip()
    return name if name else fallback


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("formbot.web.app:app", host="127.0.0.1", port=8000, reload=False)
