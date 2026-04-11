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
# ---------------------------------------------------------------------------
# Hints: fragmento de etiqueta (normalizado) → clave en el perfil maestro.
# Orden importa: los más específicos primero para evitar solapamientos.
# ---------------------------------------------------------------------------
_PROFILE_HINTS: dict[str, str] = {

    # ── Identidad principal ──────────────────────────────────────────────────
    "razon social": "razon_social",
    "nombre o razon": "razon_social",
    "nombre social": "razon_social",
    "denominacion social": "razon_social",
    "nombre empresa": "razon_social",
    "nombre o denominacion": "razon_social",
    "nombre completo empresa": "razon_social",
    "nombre comercial": "nombre_comercial",
    "nombre o sigla": "nombre_comercial",
    "sigla": "sigla",

    # ── NIT / Identificación tributaria ─────────────────────────────────────
    "nit con dv": "nit_completo",
    "nit con digito": "nit_completo",
    "nit / cc": "nit_completo",
    "cc / nit": "nit_completo",
    "identificacion fiscal": "nit_completo",
    "numero de nit": "nit_completo",
    "numero tributario": "nit_completo",
    "numero de rut": "nit_completo",
    "rut": "nit_completo",
    "nif": "nit_completo",
    "numero de identificacion tributaria": "nit_completo",
    "numero de identificacion fiscal": "nit_completo",
    "numero de identificacion": "numero_identificacion_nit",
    "numero identificacion": "numero_identificacion_nit",
    "identificacion tributaria": "numero_identificacion_nit",
    "numero de documento": "numero_identificacion_nit",
    "nit": "numero_identificacion_nit",
    "digito de verificacion": "digito_verificacion",
    "digito verificacion": "digito_verificacion",
    "d.v": "digito_verificacion",
    "d v": "digito_verificacion",
    "dv": "digito_verificacion",
    "verificacion": "digito_verificacion",

    # ── Tipo de persona / empresa / identificación ──────────────────────────
    "tipo de identificacion": "tipo_identificacion",
    "tipo identificacion": "tipo_identificacion",
    # "tipo de documento" y "tipo documento" van DESPUÉS de los hints del representante
    # para no interceptar "tipo de documento representante".
    "clase de identificacion": "tipo_identificacion",
    "tipo de persona": "tipo_persona",
    "tipo persona": "tipo_persona",
    "clase de persona": "tipo_persona",
    "persona natural o juridica": "tipo_persona",
    "tipo de empresa": "tipo_empresa",
    "tipo empresa": "tipo_empresa",
    "naturaleza juridica": "tipo_empresa",
    "tipo de sociedad": "tipo_empresa",

    # ── Matricula / constitución ─────────────────────────────────────────────
    "matricula mercantil": "matricula_mercantil",
    "numero de matricula": "matricula_mercantil",
    "camara de comercio": "matricula_mercantil",
    "fecha de constitucion": "fecha_constitucion",
    "fecha constitucion": "fecha_constitucion",
    "fecha de creacion": "fecha_constitucion",
    "fecha de fundacion": "fecha_constitucion",
    "fecha matricula": "fecha_matricula",

    # ── Dirección / ubicación ────────────────────────────────────────────────
    "direccion domicilio": "direccion_principal",
    "domicilio principal": "direccion_principal",
    "direccion principal": "direccion_principal",
    "direccion comercial": "direccion_principal",
    "direccion fiscal": "direccion_principal",
    "direccion de la empresa": "direccion_principal",
    "domicilio": "direccion_principal",
    "direccion": "direccion_principal",
    "barrio": "barrio",
    "localidad": "barrio",
    "ciudad / municipio": "ciudad_municipio",
    "ciudad o municipio": "ciudad_municipio",
    "ciudad municipio": "ciudad_municipio",
    "municipio": "ciudad_municipio",
    "ciudad": "ciudad_municipio",
    "departamento": "departamento",
    "estado / provincia": "departamento",
    "pais": "pais",
    "country": "pais",

    # ── Contacto / teléfonos ─────────────────────────────────────────────────
    "otros telefonos": "celular",
    "otros telefono": "celular",
    "telefono movil": "celular",
    "telefono celular": "celular",
    "numero celular": "celular",
    "celular": "celular",
    "movil": "contacto_celular",   # "Móvil:" aparece en la sección de contacto comercial
    "whatsapp": "celular",
    "telefono fijo": "telefono_fijo",
    "telefono principal": "telefono_fijo",
    "telefono de contacto": "telefono_fijo",
    "numero de telefono": "telefono_fijo",
    "telefax": "telefono_fijo",
    "fax": "telefono_fijo",
    "telefono": "telefono_fijo",
    "telefono alternativo": "telefono_alternativo",
    "segundo telefono": "telefono_alternativo",

    # ── Correo electrónico ───────────────────────────────────────────────────
    "correo electronico": "correo_electronico",
    "correo electronico empresa": "correo_electronico",
    "email empresa": "correo_electronico",
    "email corporativo": "correo_electronico",
    "e-mail": "correo_electronico",
    "e mail": "correo_electronico",
    "correo": "correo_electronico",
    "email": "correo_electronico",
    "pagina web": "pagina_web",
    "sitio web": "pagina_web",
    "website": "pagina_web",
    "url": "pagina_web",

    # ── Actividad económica / CIIU ───────────────────────────────────────────
    "actividad economica": "actividad_economica",
    "actividad principal": "actividad_economica",
    "objeto social": "actividad_economica",
    "descripcion de la actividad": "descripcion_actividad",
    "descripcion actividad": "descripcion_actividad",
    "actividad o negocio": "descripcion_actividad",
    "giro del negocio": "descripcion_actividad",
    "bien o servicio": "bien_servicio",
    "producto o servicio": "bien_servicio",
    "servicio prestado": "bien_servicio",
    "codigo ciiu": "codigo_ciiu",
    "ciiu": "codigo_ciiu",
    "sector economico": "sector",
    "sector": "sector",

    # ── Representante legal ──────────────────────────────────────────────────
    # Suplente ANTES del titular para que "suplente representante legal" no devuelva al titular
    "suplente representante legal": "representante_legal_suplente_nombre",
    "representante legal suplente": "representante_legal_suplente_nombre",
    "representante suplente": "representante_legal_suplente_nombre",
    "suplente del representante": "representante_legal_suplente_nombre",
    "nombre representante legal": "representante_legal_nombre",
    "nombre del representante": "representante_legal_nombre",
    "representante legal": "representante_legal_nombre",
    "nombre representante": "representante_legal_nombre",
    "nombres y apellidos del representante": "representante_legal_nombre",
    "datos representante": "representante_legal_nombre",
    "apoderado": "representante_legal_nombre",
    "gerente": "representante_legal_nombre",
    # NOTA: "nombres y apellidos" (sin "representante") se omite intencionalmente:
    # aparece también en secciones de contacto y PEP → demasiado genérico.
    "tipo de documento representante": "representante_legal_tipo_doc",
    "tipo doc representante": "representante_legal_tipo_doc",
    "tipo identificacion representante": "representante_legal_tipo_doc",
    # Hints genéricos de tipo de documento DESPUÉS de los específicos del representante
    "tipo de documento": "tipo_identificacion",
    "tipo documento": "tipo_identificacion",
    "cc representante": "representante_legal_documento",
    "cedula representante": "representante_legal_documento",
    "documento representante": "representante_legal_documento",
    "identificacion representante": "representante_legal_documento",
    "numero documento representante": "representante_legal_documento",
    "celular representante": "representante_legal_celular",
    "telefono representante": "representante_legal_celular",
    "correo representante": "representante_legal_correo",
    "email representante": "representante_legal_correo",
    "direccion representante": "representante_legal_direccion",
    "fecha expedicion representante": "representante_legal_fecha_expedicion_doc",
    "fecha de expedicion representante": "representante_legal_fecha_expedicion_doc",
    "lugar expedicion representante": "representante_legal_lugar_expedicion_doc",
    "lugar de expedicion": "representante_legal_lugar_expedicion_doc",
    "nombres y apellidos": "representante_legal_nombre",

    # ── Contador ─────────────────────────────────────────────────────────────
    # Hints específicos de correo/teléfono del contador ANTES de los genéricos
    # para que "correo del contador" no devuelva el correo de la empresa.
    "correo del contador": "contador_correo",
    "correo electronico del contador": "contador_correo",
    "correo contador": "contador_correo",
    "email del contador": "contador_correo",
    "email contador": "contador_correo",
    "correo firma contable": "contador_correo",
    "correo de la firma contable": "contador_correo",
    "telefono del contador": "contador_telefono",
    "telefono de la firma contable": "contador_telefono",
    "telefono contador": "contador_telefono",
    "tel firma contable": "contador_telefono",
    "nit del contador": "contador_empresa_nit",
    "nit contador": "contador_empresa_nit",
    "firma contadora": "contador_empresa",
    "empresa contadora": "contador_empresa",
    "nit empresa contadora": "contador_empresa_nit",
    "tarjeta profesional contador": "contador_tarjeta_profesional",
    "tp contador": "contador_tarjeta_profesional",
    "tarjeta profesional": "contador_tarjeta_profesional",
    "nombre contador": "contador_nombre",
    "revisor fiscal": "contador_nombre",
    "tipo doc contador": "contador_tipo_doc",
    "cc contador": "contador_documento",
    "cedula contador": "contador_documento",
    "documento contador": "contador_documento",
    "nombre del contador": "contador_nombre",
    "contador": "contador_nombre",

    # ── Contacto administrativo ──────────────────────────────────────────────
    "persona de contacto": "contacto_nombre",
    "nombre de contacto": "contacto_nombre",
    "nombre del contacto": "contacto_nombre",
    "contacto administrativo": "contacto_nombre",
    "contacto comercial": "contacto_nombre",
    "responsable": "contacto_nombre",
    "contacto": "contacto_nombre",
    "cargo contacto": "contacto_cargo",
    "cargo del contacto": "contacto_cargo",
    "celular contacto": "contacto_celular",
    "correo contacto": "contacto_correo",

    # ── PEP ──────────────────────────────────────────────────────────────────
    "maneja recursos publicos": "pep_maneja_recursos_publicos",
    "ejerce poder publico": "pep_ejerce_poder_publico",
    "reconocimiento publico": "pep_reconocimiento_publico",
    "vinculo con pep": "pep_vinculo_pep",
    "persona expuesta politicamente": "pep_maneja_recursos_publicos",
    "pep": "pep_maneja_recursos_publicos",

    # ── Operaciones internacionales / cumplimiento ───────────────────────────
    "operaciones internacionales": "op_internacionales",
    "realiza operaciones internacionales": "op_internacionales",
    "tipo de operacion internacional": "tipo_operacion_internacional",
    "importacion exportacion": "tipo_operacion_internacional",
    "cuentas en el exterior": "cuentas_exterior",
    "cuentas exterior": "cuentas_exterior",
    "activos virtuales": "activos_virtuales",
    "criptomonedas": "activos_virtuales",
    "intercambio activos virtuales": "intercambio_activos_virtuales",
    "transferencias activos virtuales": "transferencias_activos_virtuales",
    "controles laft": "controles_laft",
    "sistema de prevencion": "controles_laft",
    "siplaft": "controles_laft",
    "sagrilaft": "controles_laft",
    "prevencion lavado": "controles_laft",
    "lavado de activos": "controles_laft",

    # ── Referencias comerciales ──────────────────────────────────────────────
    "referencia 1 empresa": "referencia_1_empresa",
    "referencia comercial 1": "referencia_1_empresa",
    "empresa referencia 1": "referencia_1_empresa",
    "referencia 2 empresa": "referencia_2_empresa",
    "referencia comercial 2": "referencia_2_empresa",
    "empresa referencia 2": "referencia_2_empresa",
    "nit referencia 1": "referencia_1_nit",
    "nit referencia 2": "referencia_2_nit",
    "telefono referencia 1": "referencia_1_telefono",
    "telefono referencia 2": "referencia_2_telefono",
    "ciudad referencia 1": "referencia_1_ciudad",
    "ciudad referencia 2": "referencia_2_ciudad",
    "contacto referencia 1": "referencia_1_contacto",
    "contacto referencia 2": "referencia_2_contacto",
    "cupo de credito 1": "referencia_1_cupo",
    "cupo de credito 2": "referencia_2_cupo",
    "cupo de credito": "referencia_1_cupo",
    "cupo credito": "referencia_1_cupo",
    "cupo aprobado": "referencia_1_cupo",

    # ── Beneficiarios finales ────────────────────────────────────────────────
    # Hints específicos (más palabras) ANTES de los genéricos para que
    # "fecha expedicion CC beneficiario 2" no devuelva el nombre del beneficiario.
    "fecha expedicion cc beneficiario 2": "beneficiario_2_fecha_expedicion",
    "fecha expedicion beneficiario 2": "beneficiario_2_fecha_expedicion",
    "fecha expedicion cc beneficiario 1": "beneficiario_1_fecha_expedicion",
    "fecha expedicion beneficiario 1": "beneficiario_1_fecha_expedicion",
    "porcentaje participacion beneficiario 2": "beneficiario_2_participacion",
    "participacion accionaria beneficiario 2": "beneficiario_2_participacion",
    "porcentaje beneficiario 2": "beneficiario_2_participacion",
    "participacion beneficiario 2": "beneficiario_2_participacion",
    "beneficiario 2 porcentaje": "beneficiario_2_participacion",
    "beneficiario 2 participacion": "beneficiario_2_participacion",
    "porcentaje beneficiario 1": "beneficiario_1_participacion",
    "participacion beneficiario 1": "beneficiario_1_participacion",
    "beneficiario 1 porcentaje": "beneficiario_1_participacion",
    "beneficiario 1 participacion": "beneficiario_1_participacion",
    "documento beneficiario 2": "beneficiario_2_documento",
    "cc beneficiario 2": "beneficiario_2_documento",
    "tipo doc beneficiario 2": "beneficiario_2_tipo_doc",
    "documento beneficiario 1": "beneficiario_1_documento",
    "cc beneficiario 1": "beneficiario_1_documento",
    "beneficiario 1": "beneficiario_1_nombre",
    "beneficiario final 1": "beneficiario_1_nombre",
    "nombre beneficiario 1": "beneficiario_1_nombre",
    "beneficiario 2": "beneficiario_2_nombre",
    "beneficiario final 2": "beneficiario_2_nombre",
    "nombre beneficiario 2": "beneficiario_2_nombre",
    "porcentaje participacion": "beneficiario_1_participacion",
    "participacion accionaria": "beneficiario_1_participacion",

    # ── Información bancaria ─────────────────────────────────────────────────
    "entidad bancaria": "banco_nombre",
    "nombre del banco": "banco_nombre",
    "banco": "banco_nombre",
    "tipo de cuenta": "tipo_cuenta",
    "tipo cuenta": "tipo_cuenta",
    "clase de cuenta": "tipo_cuenta",
    "producto bancario": "tipo_cuenta",
    "numero de cuenta bancaria": "numero_cuenta",
    "numero de cuenta": "numero_cuenta",
    "numero cuenta": "numero_cuenta",
    "no cuenta": "numero_cuenta",           # "No. Cuenta:" → normaliza a "no cuenta"
    "cuenta bancaria": "numero_cuenta",
    "titular de la cuenta": "titular_cuenta",
    "titular cuenta": "titular_cuenta",
    "titular": "titular_cuenta",
    "sucursal": "banco_sucursal",
    "fecha apertura cuenta": "fecha_apertura_cuenta",

    # ── Estados financieros ──────────────────────────────────────────────────
    "ingresos reportados dian": "ingresos_ordinarios_reportados_dian",
    "ingresos dian": "ingresos_ordinarios_reportados_dian",
    "ingresos ordinarios reportados": "ingresos_ordinarios_reportados_dian",
    "total activos": "fin_activos",
    "activos totales": "fin_activos",
    "activos": "fin_activos",
    "total pasivos": "fin_pasivos",
    "pasivos totales": "fin_pasivos",
    "pasivos": "fin_pasivos",
    "patrimonio neto": "fin_patrimonio",
    "total patrimonio": "fin_patrimonio",
    "patrimonio": "fin_patrimonio",
    "ingresos no operacionales": "fin_otros_ingresos",
    "otros ingresos": "fin_otros_ingresos",
    "ingresos operacionales": "fin_ingresos",
    "ingresos": "fin_ingresos",
    "ventas": "fin_ingresos",
    "costos y gastos": "fin_egresos",
    "egresos": "fin_egresos",
    "gastos": "fin_egresos",
    "periodo financiero": "fin_ano",
    "ano fiscal": "fin_ano",

    # ── Firma / diligenciamiento ─────────────────────────────────────────────
    "firma representante": "firma_representante_nombre",
    "firma del representante": "firma_representante_nombre",
    "nombre para firma": "firma_representante_nombre",
    "firma de quien diligencio": "firma_diligencio_nombre",
    "firma diligenciador": "firma_diligencio_nombre",
    "quien diligencio": "firma_diligencio_nombre",
    "quien diligencia": "firma_diligencio_nombre",
    "nombre de quien diligencia": "firma_diligencio_nombre",
    "nombre quien diligencia": "firma_diligencio_nombre",
    "elaborado por": "firma_diligencio_nombre",
    "diligenciado por": "firma_diligencio_nombre",
    "fecha de diligenciamiento": "fecha_diligenciamiento_hoy",
    "fecha diligenciamiento": "fecha_diligenciamiento_hoy",
    "fecha de elaboracion": "fecha_diligenciamiento_hoy",
    "fecha de inscripcion": "fecha_diligenciamiento_hoy",
    "fecha": "fecha_diligenciamiento_hoy",
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


def _today_str() -> str:
    """Retorna la fecha actual en formato DD/MM/YYYY."""
    from datetime import date
    return date.today().strftime("%d/%m/%Y")


def _suggest_from_profile(label: str, profile: dict[str, Any]) -> str | None:
    """Intenta encontrar un valor sugerido en el perfil para la etiqueta dada.

    Estrategia de resolución (en orden de precedencia):
      1. Match exacto de slug (etiqueta normalizada == clave del perfil).
      2. Hint manual más ESPECÍFICO que coincida (más palabras = más contexto).
         Si el hint más específico tiene ≥ 2 palabras y su valor en el perfil
         es nulo/booleano → devuelve None (no cae en hints genéricos).
      3. Clave del perfil contenida textualmente en la etiqueta.
    """
    if not profile:
        return None

    from formbot.shared.utils import normalize_text
    normalized = normalize_text(label)

    # ── 1. Match directo: slug exacto ────────────────────────────────────────
    slug = "_".join(normalized.split())
    if slug in profile and profile[slug] is not None and not isinstance(profile[slug], bool):
        return str(profile[slug])

    # ── 2. Hints manuales: elegir el MÁS ESPECÍFICO (mayor número de palabras) ──
    # Esto evita que un hint genérico de 1 palabra ("correo", "telefono", "activos")
    # devuelva datos de la empresa cuando la etiqueta pide datos de otro contexto
    # ("correo del contador", "activos virtuales", etc.).
    _fecha_historica: frozenset[str] = frozenset({
        "vencimiento", "expedicion", "constitucion", "apertura",
        "nacimiento", "emision", "fundacion", "matricula", "creacion", "vigencia",
    })

    best_hint: str | None = None
    best_key:  str | None = None
    best_words: int = 0

    for hint, key in _PROFILE_HINTS.items():
        if hint not in normalized:
            continue
        if hint == "rut" and "fecha" in normalized:
            continue
        n_words = len(hint.split())
        if n_words > best_words:
            best_words = n_words
            best_hint  = hint
            best_key   = key

    if best_key is not None:
        # Caso especial: fecha de hoy
        if best_key == "fecha_diligenciamiento_hoy":
            if any(tok in normalized for tok in _fecha_historica):
                return None
            return _today_str()

        profile_value = profile.get(best_key)

        # REGLA CLAVE: si el hint más específico tiene ≥ 2 palabras y el valor
        # en el perfil está ausente (None) o es booleano, devolver None.
        # Esto impide que "Correo del contador" use el correo de la empresa,
        # o que "Activos virtuales" use el total de activos financieros.
        if best_words >= 2 and (profile_value is None or isinstance(profile_value, bool)):
            return None

        if profile_value is not None and not isinstance(profile_value, bool):
            return str(profile_value)
        # hint de 1 sola palabra con valor nulo → continuar a paso 3

    # ── 3. Clave del perfil contenida en la etiqueta (dirección única) ───────
    # Solo "key_norm in normalized" — evita que etiquetas cortas ("nombre", "cargo")
    # coincidan con cualquier clave que las contenga.
    for key, value in profile.items():
        if value is None or isinstance(value, bool):
            continue
        key_norm = normalize_text(key.replace("_", " "))
        if key_norm and key_norm in normalized:
            return str(value)

    return None


# ---------------------------------------------------------------------------
# Smart Mapping — Fases 1-4 (confianza + validación + interacción)
# ---------------------------------------------------------------------------

_CONF_HIGH: float = 0.80    # >= alta  → asignar automáticamente
_CONF_MEDIUM: float = 0.50  # >= media → solicitar confirmación al usuario

_KNOWN_DATE_HISTORICAL: frozenset[str] = frozenset({
    "vencimiento", "expedicion", "constitucion", "apertura",
    "nacimiento", "emision", "fundacion", "matricula", "creacion", "vigencia",
})

# Tipos esperados por clave de perfil — usados en validación cruzada
_KEY_TYPE: dict[str, str] = {
    "nit_completo":                "nit",
    "numero_identificacion_nit":   "nit",
    "digito_verificacion":         "digit",
    "contador_empresa_nit":        "nit",
    "referencia_1_nit":            "nit",
    "referencia_2_nit":            "nit",
    "correo_electronico":          "email",
    "representante_legal_correo":  "email",
    "contacto_correo":             "email",
    "telefono_fijo":               "phone",
    "celular":                     "phone",
    "contacto_celular":            "phone",
    "representante_legal_celular": "phone",
    "telefono_alternativo":        "phone",
    "numero_cuenta":               "numeric",
}

_KEY_READABLE: dict[str, str] = {
    "nit_completo":                "NIT completo (con DV)",
    "numero_identificacion_nit":   "Número de NIT",
    "digito_verificacion":         "Dígito de verificación",
    "razon_social":                "Razón social",
    "nombre_comercial":            "Nombre comercial",
    "correo_electronico":          "Correo electrónico",
    "telefono_fijo":               "Teléfono fijo",
    "celular":                     "Celular",
    "tipo_identificacion":         "Tipo de identificación",
    "tipo_persona":                "Tipo de persona",
    "tipo_empresa":                "Tipo de empresa",
    "representante_legal_nombre":  "Nombre del representante legal",
    "representante_legal_documento": "Documento del representante",
    "ciudad_municipio":            "Ciudad / municipio",
    "departamento":                "Departamento",
    "pais":                        "País",
    "direccion_principal":         "Dirección principal",
    "banco_nombre":                "Banco",
    "numero_cuenta":               "Número de cuenta",
    "tipo_cuenta":                 "Tipo de cuenta",
    "contacto_nombre":             "Nombre de contacto",
    "actividad_economica":         "Actividad económica",
    "matricula_mercantil":         "Matrícula mercantil",
}


def _type_mismatch_penalty(profile_key: str, value: str) -> float:
    """FASE 2 — Validación de tipo: retorna penalización si el valor no concuerda."""
    expected = _KEY_TYPE.get(profile_key)
    if not expected:
        return 0.0
    digits = re.sub(r"\D", "", value)
    if expected == "email":
        parts = value.split("@")
        return 0.0 if (len(parts) == 2 and "." in parts[1]) else -0.50
    if expected in {"nit", "numeric"}:
        if "@" in value:
            return -0.50   # email en campo NIT
        return 0.0 if digits else -0.40
    if expected == "phone":
        if "@" in value:
            return -0.50
        return 0.0 if len(digits) >= 7 else -0.30
    if expected == "digit":
        s = value.strip()
        return 0.0 if (s.isdigit() and len(s) <= 2) else -0.30
    return 0.0


def _readable_key(key: str) -> str:
    return _KEY_READABLE.get(key, key.replace("_", " ").title())


def _low_conf_result(
    key: str | None, value: str | None, possible_keys: list[str]
) -> dict[str, Any]:
    return {
        "confidence": 0.0, "confidence_level": "low",
        "profile_key": key, "value": value,
        "question": None, "possible_keys": possible_keys,
    }


def _smart_map_field(field_key: str, label: str, profile: dict[str, Any]) -> dict[str, Any]:
    """
    Fases 1-4: mapea un campo con semántica, valida el tipo y calcula confianza.

    Retorna dict:
        confidence        float 0-1
        confidence_level  "high" | "medium" | "low"
        profile_key       str | None
        value             str | None
        question          str | None  (solo para medium)
        possible_keys     list[str]
    """
    from formbot.shared.utils import normalize_text
    normalized = normalize_text(label)

    # FASE 1 — Recopilar hints que aplican
    matches: list[tuple[str, str, int]] = []  # (hint, key, n_words)
    for hint, key in _PROFILE_HINTS.items():
        if hint not in normalized:
            continue
        if hint == "rut" and "fecha" in normalized:
            continue
        matches.append((hint, key, len(hint.split())))

    if not matches:
        return _low_conf_result(None, None, [])

    # Ordenar: más palabras primero (más específico)
    matches.sort(key=lambda x: x[2], reverse=True)
    best_hint, best_key, best_words = matches[0]

    # Caso especial: fecha de hoy
    if best_key == "fecha_diligenciamiento_hoy":
        if any(tok in normalized for tok in _KNOWN_DATE_HISTORICAL):
            return _low_conf_result(None, None, [])
        return {
            "confidence": 0.92, "confidence_level": "high",
            "profile_key": "fecha_diligenciamiento_hoy",
            "value": _today_str(),
            "question": None,
            "possible_keys": ["fecha_diligenciamiento_hoy"],
        }

    # Valor del perfil
    profile_value: str | None = None
    if best_key in profile and profile[best_key] not in (None, False, True):
        profile_value = str(profile[best_key])

    if profile_value is None:
        return _low_conf_result(best_key, None, [best_key])

    # FASE 1 — Confianza base según especificidad del hint (n° palabras)
    base = (
        0.93 if best_words >= 4
        else 0.87 if best_words == 3
        else 0.75 if best_words == 2
        else 0.55
    )

    # Penalización por múltiples keys distintos (ambigüedad)
    unique_keys: list[str] = list(dict.fromkeys(m[1] for m in matches))
    if len(unique_keys) > 1:
        base -= 0.10 * (len(unique_keys) - 1)

    # FASE 2 — Validación de tipo
    base += _type_mismatch_penalty(best_key, profile_value)
    confidence = round(max(0.0, min(1.0, base)), 4)

    # FASE 3 — Decisión inteligente
    if confidence >= _CONF_HIGH:
        return {
            "confidence": confidence, "confidence_level": "high",
            "profile_key": best_key, "value": profile_value,
            "question": None, "possible_keys": unique_keys,
        }

    if confidence >= _CONF_MEDIUM:
        # FASE 4 — Generar pregunta clara
        if len(unique_keys) > 1:
            opts = " o ".join(f'"{_readable_key(k)}"' for k in unique_keys[:3])
            question = f"\u00bfEl campo \u00ab{label}\u00bb corresponde a {opts}?"
        else:
            question = (
                f"\u00bf\u00ab{label}\u00bb equivale a "
                f"\u00ab{_readable_key(best_key)}\u00bb "
                f"(valor sugerido: \u00ab{profile_value}\u00bb)?"
            )
        return {
            "confidence": confidence, "confidence_level": "medium",
            "profile_key": best_key, "value": profile_value,
            "question": question, "possible_keys": unique_keys,
        }

    return _low_conf_result(best_key, profile_value, unique_keys)


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
    .btn-auto { background: linear-gradient(135deg, #1d6f42, #2ea863); color: #fff; }
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

    /* ── Smart Analysis UI ── */
    .btn-smart { background: linear-gradient(135deg, #7c3aed, #a855f7); color: #fff; }

    .bucket { margin-bottom: 18px; }
    .bucket-header {
      display: flex; align-items: center; gap: 8px;
      font-size: 13px; font-weight: 700;
      padding: 8px 14px; border-radius: 10px;
      margin-bottom: 10px;
    }
    .bucket-auto-hdr    { background: #edfaf3; color: #0f8b5f; }
    .bucket-confirm-hdr { background: #fff8e8; color: #b07c00; }
    .bucket-reject-hdr  { background: #f3f3f6; color: #5a5a72; }

    .smart-field {
      border: 1.5px solid var(--line); border-radius: 12px;
      padding: 12px 14px; margin-bottom: 8px;
    }
    .smart-field.s-auto   { border-color: #b8ecd9; background: #f0fbf5; }
    .smart-field.s-confirm{ border-color: #ffe5a0; background: #fffbf0; }
    .smart-field.s-reject { border-color: #e0e0e8; background: #f8f8fb; }

    .sf-row { display: flex; justify-content: space-between; align-items: center; gap: 8px; margin-bottom: 4px; }
    .sf-label { font-size: 12px; font-weight: 700; color: #1e3359; word-break: break-word; }
    .sf-value { font-family: "IBM Plex Mono", monospace; font-size: 11px; color: #1d6f42; margin-bottom: 4px; }
    .sf-question { font-size: 12px; color: #7a6000; margin: 5px 0; }
    .sf-reason { font-size: 11px; color: #888; margin-bottom: 5px; }

    .conf-badge {
      white-space: nowrap; font-size: 10px; font-weight: 700;
      padding: 2px 9px; border-radius: 999px; flex-shrink: 0;
    }
    .cb-high   { background: #d4f3e3; color: #0f8b5f; }
    .cb-medium { background: #fff0c0; color: #856300; }
    .cb-low    { background: #ebebf0; color: #666; }

    .smart-input {
      width: 100%; padding: 8px 10px;
      font-size: 13px; font-family: "Space Grotesk", sans-serif;
      border: 1.5px solid #d0dff5; border-radius: 8px;
      background: #fff; color: var(--ink); outline: none;
      transition: border-color .18s;
    }
    .smart-input:focus       { border-color: var(--accent-2); }
    .smart-input.s-confirm-i { border-color: #ffe0a0; }
    .smart-input.s-reject-i  { border-color: #ddd; }
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
      <button class="btn-smart" id="smart-btn" onclick="smartAnalyze()" disabled>
        &#129504; Smart Analysis
      </button>
      <button class="btn-auto" id="auto-btn" onclick="autoFill()" disabled>
        &#9889; Registro autom&#225;tico
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


  <!-- ── PASO 3: Smart Analysis ── -->
  <section class="card" id="step3" style="display:none; margin-top:18px;">
    <div class="step2-header">
      <div>
        <div class="step2-title" id="step3-title">Smart Analysis</div>
        <div class="step2-meta" id="step3-meta"></div>
      </div>
      <button class="btn-ghost" onclick="resetToStep1()">&#8592; Cargar otro</button>
    </div>

    <!-- Bucket: Auto-asignados -->
    <div class="bucket" id="bucket-auto"></div>

    <!-- Bucket: Requieren confirmaci&#243;n -->
    <div class="bucket" id="bucket-confirm"></div>

    <!-- Bucket: Rechazados / manuales -->
    <div class="bucket" id="bucket-reject"></div>

    <hr class="divider">
    <div class="actions">
      <button class="btn-primary" id="smart-fill-btn" onclick="smartFill()">
        &#9654; Completar y Descargar
      </button>
      <button class="btn-ghost" onclick="resetToStep1()">&#10005; Cancelar</button>
    </div>
    <div class="progress-bar" id="progress3"></div>
    <div class="status" id="status3" style="display:none"></div>
  </section>

</div>

<script>
  /* ── Estado ── */
  let currentFile = null;
  let detectedFields = [];
  let smartData = null;

  const dropZone   = document.getElementById("drop-zone");
  const fileInput  = document.getElementById("file-input");
  const analyzeBtn = document.getElementById("analyze-btn");
  const autoBtn    = document.getElementById("auto-btn");
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
    // FIX: B1 — Reset COMPLETO del estado anterior antes de procesar el nuevo archivo.
    // Ningún resultado de la sesión previa debe sobrevivir a la carga de un nuevo documento.
    detectedFields = [];
    smartData = null;
    fieldsGrid.innerHTML = "";
    step2.style.display = "none";
    status2.style.display = "none";
    const _s3 = document.getElementById("step3");
    if (_s3) _s3.style.display = "none";
    const _bAuto = document.getElementById("bucket-auto");
    if (_bAuto) _bAuto.innerHTML = "";
    const _bConf = document.getElementById("bucket-confirm");
    if (_bConf) _bConf.innerHTML = "";
    const _bRej = document.getElementById("bucket-reject");
    if (_bRej) _bRej.innerHTML = "";
    const _st3 = document.getElementById("status3");
    if (_st3) _st3.style.display = "none";
    // FIX: B1 — fin reset

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
    autoBtn.disabled    = false;
    document.getElementById("smart-btn").disabled = false;
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
    const s3 = document.getElementById("step3");
    if (s3) s3.style.display = "none";
    smartData = null;
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

  /* ── Registro automático (un clic) ── */
  async function autoFill() {
    if (!currentFile) return;
    autoBtn.disabled    = true;
    analyzeBtn.disabled = true;
    prog1.style.display = "block";
    setStatus1("Diligenciando autom\u00e1ticamente con perfil ASTECO\u2026", "run");

    const body = new FormData();
    body.append("template", currentFile);

    try {
      const resp = await fetch("/api/fill-auto", { method: "POST", body });
      if (!resp.ok) {
        const json = await resp.json().catch(() => ({}));
        throw new Error(json.detail || "Error " + resp.status);
      }

      const filled   = parseInt(resp.headers.get("X-Fields-Filled") || "0", 10);
      const blob     = await resp.blob();
      const filename = extractFilename(resp.headers.get("content-disposition"));
      const url      = URL.createObjectURL(blob);
      const a        = document.createElement("a");
      a.href = url; a.download = filename;
      document.body.appendChild(a); a.click(); a.remove();
      URL.revokeObjectURL(url);

      setStatus1(
        "\u2705 Registro completado: " + filled + " campo(s) diligenciados. Descargando: " + filename,
        "ok"
      );
    } catch (err) {
      setStatus1("Error: " + err.message, "err");
    } finally {
      autoBtn.disabled    = false;
      analyzeBtn.disabled = false;
      prog1.style.display = "none";
    }
  }

  /* ── Smart Analysis ── */
  async function smartAnalyze() {
    if (!currentFile) return;
    const smartBtn = document.getElementById("smart-btn");
    smartBtn.disabled   = true;
    analyzeBtn.disabled = true;
    autoBtn.disabled    = true;
    prog1.style.display = "block";
    setStatus1("Analizando con sistema de confianza\u2026", "run");

    const body = new FormData();
    body.append("template", currentFile);

    try {
      const resp = await fetch("/api/smart-analyze", { method: "POST", body });
      const data = await resp.json();
      if (!resp.ok) throw new Error(data.detail || "Error " + resp.status);

      smartData = data;
      renderSmartResults(data);
      const s3 = document.getElementById("step3");
      s3.style.display = "block";
      s3.scrollIntoView({ behavior: "smooth", block: "start" });
      const sm = data.summary;
      setStatus1(
        "\u2705 Smart Analysis: " + sm.auto_mapped + " auto \u00b7 " +
        sm.needs_confirmation + " confirmar \u00b7 " + sm.rejected + " sin asignar.",
        "ok"
      );
    } catch (err) {
      setStatus1("Error: " + err.message, "err");
    } finally {
      smartBtn.disabled   = false;
      analyzeBtn.disabled = false;
      autoBtn.disabled    = false;
      prog1.style.display = "none";
    }
  }

  function renderSmartResults(data) {
    const sm = data.summary;
    document.getElementById("step3-title").textContent =
      "Smart Analysis \u2014 " + sm.total + " campos";
    document.getElementById("step3-meta").textContent =
      currentFile.name + " \u00b7 " +
      sm.auto_mapped + " auto \u00b7 " +
      sm.needs_confirmation + " confirmar \u00b7 " +
      sm.rejected + " sin asignar";

    /* ── Bucket: Auto-asignados ── */
    const autoEl = document.getElementById("bucket-auto");
    autoEl.innerHTML = "";
    if (data.auto_mapped.length > 0) {
      autoEl.innerHTML =
        '<div class="bucket-header bucket-auto-hdr">\u2705 Auto-asignados (' + data.auto_mapped.length + ')</div>';
      data.auto_mapped.forEach(f => {
        autoEl.innerHTML +=
          '<div class="smart-field s-auto">' +
            '<div class="sf-row">' +
              '<span class="sf-label">' + escHtml(f.label) + '</span>' +
              '<span class="conf-badge cb-high">Alta ' + Math.round(f.confidence_score * 100) + '%</span>' +
            '</div>' +
            '<div class="sf-value">' + escHtml(f.value) + '</div>' +
            '<input type="hidden" data-field="' + escAttr(f.field) + '" data-label="' + escAttr(f.label) + '" value="' + escAttr(f.value) + '">' +
          '</div>';
      });
    }

    /* ── Bucket: Requieren confirmaci\u00f3n ── */
    const confirmEl = document.getElementById("bucket-confirm");
    confirmEl.innerHTML = "";
    if (data.needs_confirmation.length > 0) {
      confirmEl.innerHTML =
        '<div class="bucket-header bucket-confirm-hdr">\u26a0\ufe0f Requieren confirmaci\u00f3n (' + data.needs_confirmation.length + ')</div>';
      data.needs_confirmation.forEach(f => {
        confirmEl.innerHTML +=
          '<div class="smart-field s-confirm">' +
            '<div class="sf-row">' +
              '<span class="sf-label">' + escHtml(f.label) + '</span>' +
              '<span class="conf-badge cb-medium">Media ' + Math.round(f.confidence_score * 100) + '%</span>' +
            '</div>' +
            '<div class="sf-question">' + escHtml(f.question || "") + '</div>' +
            '<input type="text" class="smart-input s-confirm-i"' +
              ' data-field="' + escAttr(f.field) + '"' +
              ' data-label="' + escAttr(f.label) + '"' +
              ' value="' + escAttr(f.suggested_value || "") + '"' +
              ' placeholder="Confirmar valor o dejar vac\u00edo para omitir">' +
          '</div>';
      });
    }

    /* ── Bucket: Sin asignar / manuales ── */
    const rejectEl = document.getElementById("bucket-reject");
    rejectEl.innerHTML = "";
    if (data.rejected.length > 0) {
      rejectEl.innerHTML =
        '<div class="bucket-header bucket-reject-hdr">\u274c Sin asignar (' + data.rejected.length + ')</div>';
      data.rejected.forEach(f => {
        rejectEl.innerHTML +=
          '<div class="smart-field s-reject">' +
            '<div class="sf-row">' +
              '<span class="sf-label">' + escHtml(f.label) + '</span>' +
              '<span class="conf-badge cb-low">Baja</span>' +
            '</div>' +
            '<div class="sf-reason">' + escHtml(f.reason) + '</div>' +
            '<input type="text" class="smart-input s-reject-i"' +
              ' data-field="' + escAttr(f.field) + '"' +
              ' data-label="' + escAttr(f.label) + '"' +
              ' value=""' +
              ' placeholder="Ingresar valor manualmente (opcional)">' +
          '</div>';
      });
    }
  }

  async function smartFill() {
    if (!currentFile || !smartData) return;
    const sfBtn = document.getElementById("smart-fill-btn");
    sfBtn.disabled = true;
    document.getElementById("progress3").style.display = "block";
    const st3 = document.getElementById("status3");
    st3.style.display = "block";
    st3.textContent = "Diligenciando\u2026";
    st3.className = "status run";

    const fields = [];

    /* Auto-asignados (hidden inputs) */
    document.querySelectorAll('#bucket-auto input[type="hidden"]').forEach(inp => {
      if (inp.value.trim()) fields.push({
        field_key: inp.dataset.field,
        label:     inp.dataset.label,
        value:     inp.value.trim(),
      });
    });

    /* Confirmados por el usuario */
    document.querySelectorAll('#bucket-confirm input[type="text"]').forEach(inp => {
      if (inp.value.trim()) fields.push({
        field_key: inp.dataset.field,
        label:     inp.dataset.label,
        value:     inp.value.trim(),
      });
    });

    /* Ingresados manualmente en rechazados */
    document.querySelectorAll('#bucket-reject input[type="text"]').forEach(inp => {
      if (inp.value.trim()) fields.push({
        field_key: inp.dataset.field,
        label:     inp.dataset.label,
        value:     inp.value.trim(),
      });
    });

    if (fields.length === 0) {
      st3.textContent = "No hay valores para diligenciar. Confirma o ingresa al menos un campo.";
      st3.className = "status err";
      sfBtn.disabled = false;
      document.getElementById("progress3").style.display = "none";
      return;
    }

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
      st3.textContent = "\u2705 Listo. Descargando: " + filename;
      st3.className = "status ok";
    } catch (err) {
      st3.textContent = "Error: " + err.message;
      st3.className = "status err";
    } finally {
      sfBtn.disabled = false;
      document.getElementById("progress3").style.display = "none";
    }
  }
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


@app.post("/api/fill-auto")
async def fill_auto(
    template: UploadFile = File(...),
) -> Response:
    """Diligencia el documento en un solo paso: escanea, sugiere desde el perfil y descarga.

    No requiere intervención del usuario. Solo llena campos con sugerencia del perfil maestro.
    Retorna el documento con el encabezado X-Fields-Filled indicando cuántos campos se llenaron.
    """
    from formbot.infrastructure.document_scanners.field_scanner import scan_document

    with tempfile.TemporaryDirectory(prefix="formbot-auto-") as tmpdir:
        tmp = Path(tmpdir)
        template_path = tmp / _safe_filename(template.filename, "template.xlsx")
        template_path.write_bytes(await template.read())

        suffix = template_path.suffix.lower()
        if suffix not in SUPPORTED_EXTENSIONS:
            return JSONResponse(
                status_code=400,
                content={"detail": f"Formato '{suffix}' no soportado."},
            )
        mime_type = SUPPORTED_EXTENSIONS[suffix]
        output_filename = f"{template_path.stem}_auto_{uuid4().hex[:8]}{suffix}"
        output_path = tmp / output_filename

        try:
            detected = scan_document(template_path)
        except Exception as exc:
            LOGGER.exception("Error escaneando documento en fill-auto")
            return JSONResponse(
                status_code=500,
                content={"detail": f"Error al analizar el documento: {exc}"},
            )

        profile = _load_master_profile()
        active: list[dict] = []
        for field in detected:
            suggested = _suggest_from_profile(field.label, profile)
            if suggested and suggested.strip():
                active.append({
                    "field_key": field.field_key,
                    "label":     field.label,
                    "value":     suggested,
                })

        if not active:
            return JSONResponse(
                status_code=422,
                content={"detail": "No se encontró ningún campo con sugerencia del perfil."},
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
            LOGGER.exception("Error en fill-auto")
            return JSONResponse(
                status_code=500,
                content={"detail": f"Error al diligenciar: {type(exc).__name__}: {exc}"},
            )

        payload = output_path.read_bytes()
        headers = {
            "Content-Disposition": f'attachment; filename="{output_filename}"',
            "X-Fields-Filled":     str(len(active)),
        }
        return Response(content=payload, media_type=mime_type, headers=headers)


@app.post("/api/smart-analyze")
async def smart_analyze(
    template: UploadFile = File(...),
) -> JSONResponse:
    """Fases 1-4: escanea el documento y clasifica campos por nivel de confianza.

    Retorna:
        auto_mapped        → confianza alta, asignación directa
        needs_confirmation → confianza media, requiere aprobación del usuario
        rejected           → confianza baja, NUNCA se asignan automáticamente
    """
    from formbot.infrastructure.document_scanners.field_scanner import scan_document

    with tempfile.TemporaryDirectory(prefix="formbot-smart-") as tmpdir:
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
            LOGGER.exception("smart-analyze: error escaneando documento")
            return JSONResponse(
                status_code=500,
                content={"detail": f"Error al analizar el documento: {exc}"},
            )

        profile = _load_master_profile()
        auto_mapped: list[dict] = []
        needs_confirmation: list[dict] = []
        rejected: list[dict] = []

        for fld in detected:
            result = _smart_map_field(fld.field_key, fld.label, profile)
            level = result["confidence_level"]

            if level == "high" and result["value"] is not None:
                # FASE 3 → confianza alta: asignar automáticamente
                auto_mapped.append({
                    "label":            fld.label,
                    "field":            fld.field_key,
                    "value":            result["value"],
                    "confidence":       "high",
                    "confidence_score": result["confidence"],
                })
            elif level == "medium":
                # FASE 4 → confianza media: preguntar al usuario
                needs_confirmation.append({
                    "label":            fld.label,
                    "field":            fld.field_key,
                    "possible_fields":  result["possible_keys"],
                    "suggested_value":  result["value"],
                    "question":         result["question"],
                    "confidence_score": result["confidence"],
                })
            else:
                # FASE 3 → confianza baja: NUNCA asignar
                if result["profile_key"] and result["value"] is None:
                    reason = f"Campo '{result['profile_key']}' no está en el perfil"
                elif result["value"] is not None:
                    reason = f"Confianza insuficiente ({result['confidence']:.0%})"
                else:
                    reason = "Sin coincidencia en el perfil"
                rejected.append({
                    "label":  fld.label,
                    "field":  fld.field_key,
                    "reason": reason,
                })

        LOGGER.info(
            "smart-analyze: total=%d auto=%d confirmar=%d rechazados=%d",
            len(detected), len(auto_mapped), len(needs_confirmation), len(rejected),
        )
        return JSONResponse({
            "format":             suffix.lstrip("."),
            "auto_mapped":        auto_mapped,
            "needs_confirmation": needs_confirmation,
            "rejected":           rejected,
            "summary": {
                "total":              len(detected),
                "auto_mapped":        len(auto_mapped),
                "needs_confirmation": len(needs_confirmation),
                "rejected":           len(rejected),
            },
        })


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
