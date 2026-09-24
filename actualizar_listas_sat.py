#!/usr/bin/env python3
"""
Descarga las listas del SAT (artículos 69-B y 69 del CFF) y las carga en
Supabase, para que la app pueda validar a los proveedores contra ellas.

El SAT no ofrece una API de consulta por RFC: publica las listas completas como
CSV en datos abiertos. Este proceso:

  1. Descarga cada archivo.
  2. Compara su hash con el de la última carga. Si no cambió, no hace nada:
     las listas se actualizan pocas veces al mes y recargarlas a diario sería
     reescribir cientos de miles de filas para quedar igual.
  3. Si cambió, carga las filas como un LOTE nuevo.
  4. Solo cuando el lote entró completo lo marca como vigente, y entonces
     borra el anterior. La app lee la vista sat_listas_vigente, así que nunca
     ve una lista a medio cargar ni un hueco mientras se reemplaza.

Si una carga falla a medias, se borra lo que alcanzó a entrar del lote nuevo y
el anterior sigue vigente.

Configuración: la misma de mover_adjuntos_a_drive.py, en
%USERPROFILE%\\.smi\\config.env (SUPABASE_URL y SUPABASE_SERVICE_KEY).

Requiere 57-listas-sat.sql
"""

import csv
import hashlib
import io
import os
import re
import sys
import uuid
from datetime import datetime, timezone

import requests

# --------------------------------------------------------------------------
# Configuración — mismo archivo y mismas reglas que el script de Drive
# --------------------------------------------------------------------------
CARPETA_CONF = os.environ.get("SMI_CONF_DIR") or os.path.join(os.path.expanduser("~"), ".smi")
ARCHIVO_CONF = os.path.join(CARPETA_CONF, "config.env")


def _cargar_conf():
    if not os.path.exists(ARCHIVO_CONF):
        return
    with open(ARCHIVO_CONF, encoding="utf8") as fh:
        for linea in fh:
            linea = linea.strip()
            if not linea or linea.startswith("#") or "=" not in linea:
                continue
            clave, _, valor = linea.partition("=")
            clave, valor = clave.strip(), valor.strip().strip('"').strip("'")
            if clave and clave not in os.environ:
                os.environ[clave] = valor


_cargar_conf()

SUPABASE_URL = os.environ.get("SUPABASE_URL", "").rstrip("/")
SERVICE_KEY = os.environ.get("SUPABASE_SERVICE_KEY", "")

if not SUPABASE_URL or not SERVICE_KEY:
    sys.exit(f"Faltan SUPABASE_URL o SUPABASE_SERVICE_KEY. Ponlos en {ARCHIVO_CONF} o en el entorno.")

CAB = {
    "apikey": SERVICE_KEY,
    "Authorization": f"Bearer {SERVICE_KEY}",
    "Content-Type": "application/json",
}

BASE = "https://wu1agsprosta001.blob.core.windows.net/agsc-publicaciones/Datos_abiertos"

# Los archivos que importan para validar a un proveedor. Del 69 se dejan
# fuera condonaciones y reducciones de multas: no dicen nada del riesgo de
# contratar a alguien. Si el SAT vuelve a mover los archivos, basta con
# corregir la URL aquí.
FUENTES = [
    {"fuente": "69B",             "lista": "69-B", "supuesto": None,
     "url": f"{BASE}/Documents_AGAFF/Listado_completo_69-B.csv"},
    {"fuente": "69_no_localizados", "lista": "69", "supuesto": "No localizados",
     "url": f"{BASE}/Documents_AGR/No_localizados.csv"},
    {"fuente": "69_firmes",       "lista": "69", "supuesto": "Créditos firmes",
     "url": f"{BASE}/Documents_AGR/Firmes.csv"},
    {"fuente": "69_exigibles",    "lista": "69", "supuesto": "Créditos exigibles",
     "url": f"{BASE}/Documents_AGR/Exigibles.csv"},
    {"fuente": "69_cancelados",   "lista": "69", "supuesto": "Créditos cancelados",
     "url": f"{BASE}/Documents_AGR/Cancelados.csv"},
    {"fuente": "69_sentencias",   "lista": "69", "supuesto": "Sentencia condenatoria",
     "url": f"{BASE}/Documents_AGR/Sentencias.csv"},
    {"fuente": "69_csd_sin_efectos", "lista": "69", "supuesto": "CSD sin efectos",
     "url": f"{BASE}/Documents_AGR/CSDsinefectos.csv"},
]

LOTE_INSERT = 1000

# `--forzar` recarga aunque el archivo no haya cambiado: sirve cuando lo que
# cambió es este script (por ejemplo, cómo se valida un RFC).
FORZAR = "--forzar" in sys.argv

# Los renglones descartados se guardan aquí para poder revisarlos.
ARCHIVO_DESCARTES = os.path.join(CARPETA_CONF, "sat_descartados.txt")
DESCARTES = []


# --------------------------------------------------------------------------
# Supabase
# --------------------------------------------------------------------------
def rest(metodo, ruta, **kw):
    r = requests.request(metodo, f"{SUPABASE_URL}/rest/v1/{ruta}", headers={**CAB, **kw.pop("headers", {})},
                         timeout=120, **kw)
    if not r.ok:
        raise requests.HTTPError(f"{r.status_code} en {ruta}: {r.text[:300]}", response=r)
    return r.json() if r.content and r.headers.get("content-type", "").startswith("application/json") else None


# --------------------------------------------------------------------------
# Lectura de los CSV del SAT
# --------------------------------------------------------------------------
def _norm(s):
    """Encabezado comparable: sin acentos, mayúsculas, sin espacios dobles."""
    import unicodedata
    s = unicodedata.normalize("NFD", str(s or ""))
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return " ".join(s.upper().split())


def decodificar(contenido):
    """Los archivos del SAT han venido en UTF-8 y en Windows-1252 según el año;
       se prueba el primero y, si no, el segundo, que acepta cualquier byte."""
    try:
        return contenido.decode("utf-8-sig")
    except UnicodeDecodeError:
        return contenido.decode("cp1252", errors="replace")


# Forma oficial: 3 letras (moral) o 4 (física), fecha AAMMDD y homoclave.
# El & y la Ñ son válidos al inicio: "A&C..." es un RFC real, y la versión
# anterior lo descartaba porque exigía solo letras.
RFC_RX = re.compile(r"^[A-ZÑ&]{3,4}\d{6}[A-Z0-9]{3}$")


def rfc_valido(r):
    return bool(RFC_RX.match(r))


def leer_csv(texto, fuente):
    """Devuelve (filas, descartadas). No depende del orden de las columnas:
       las busca por nombre, porque el SAT antepone renglones de título y ha
       cambiado el acomodo entre publicaciones."""
    renglones = list(csv.reader(io.StringIO(texto)))

    # El encabezado es el primer renglón con una celda que diga RFC.
    ih = next((i for i, r in enumerate(renglones) if any(_norm(c) == "RFC" for c in r)), None)
    if ih is None:
        raise ValueError("no se encontró la columna RFC; ¿cambió el formato del archivo?")
    enc = [_norm(c) for c in renglones[ih]]

    def col(*palabras):
        # Por orden de preferencia de la palabra, no de la columna: si no,
        # "Situación del contribuyente" se tomaría por el nombre.
        for p in palabras:
            for i, h in enumerate(enc):
                if p in h and "SITUACION" not in h:
                    return i
        return None

    c_rfc = enc.index("RFC")
    c_nombre = col("NOMBRE", "RAZON SOCIAL", "CONTRIBUYENTE")
    c_sit = next((i for i, h in enumerate(enc) if "SITUACION" in h), None) if fuente["lista"] == "69-B" else None

    filas, descartadas = {}, 0
    for r in renglones[ih + 1:]:
        if c_rfc >= len(r):
            continue
        rfc = "".join(r[c_rfc].split()).replace("-", "").upper()
        if not rfc_valido(rfc):
            descartadas += 1
            if any(c.strip() for c in r):
                DESCARTES.append(f"{fuente['fuente']}\t{r[c_rfc]!r}\t{' | '.join(c.strip() for c in r)[:200]}")
            continue
        fila = {
            "rfc": rfc,
            "nombre": (r[c_nombre].strip()[:300] if c_nombre is not None and c_nombre < len(r) else None) or None,
            "situacion": fuente["supuesto"],
            "detalle": None,
        }
        if fuente["lista"] == "69-B":
            fila["situacion"] = (r[c_sit].strip() if c_sit is not None and c_sit < len(r) else None) or None
            # Oficios y fechas de publicación: es lo que se cita si hay que
            # justificar por qué se dejó de trabajar con alguien.
            fila["detalle"] = {renglones[ih][i].strip(): v.strip()
                               for i, v in enumerate(r)
                               if i not in (c_rfc, c_nombre, c_sit) and i < len(renglones[ih])
                               and enc[i] not in ("NO", "NUM", "#")
                               and renglones[ih][i].strip() and v.strip()} or None
        # Un RFC puede repetirse en el 69 (un renglón por crédito). Para
        # validar basta con saber que está; se conserva uno.
        filas[rfc] = fila
    return list(filas.values()), descartadas


# --------------------------------------------------------------------------
# Carga
# --------------------------------------------------------------------------
def cargar_fuente(f):
    ahora = datetime.now(timezone.utc).isoformat()
    print(f"  {f['fuente']}")

    r = requests.get(f["url"], timeout=300)
    r.raise_for_status()
    contenido = r.content
    huella = hashlib.sha256(contenido).hexdigest()

    previa = (rest("GET", "sat_listas_fuentes",
                   params={"select": "hash,lote_vigente", "fuente": f"eq.{f['fuente']}"}) or [None])[0]
    if previa and previa.get("hash") == huella and previa.get("lote_vigente") and not FORZAR:
        rest("PATCH", "sat_listas_fuentes", params={"fuente": f"eq.{f['fuente']}"},
             json={"revisado_en": ahora}, headers={"Prefer": "return=minimal"})
        print("    sin cambios")
        return "sin_cambios", 0

    filas, descartadas = leer_csv(decodificar(contenido), f)
    if not filas:
        # Un archivo vacío casi seguro es una falla de publicación, no que el
        # SAT haya limpiado la lista. Se conserva lo que había.
        raise ValueError("el archivo no trajo ningún RFC válido; se conserva la carga anterior")

    # La fuente tiene que existir antes que sus filas (llave foránea). Si es
    # nueva, se registra sin lote vigente todavía.
    if not previa:
        rest("POST", "sat_listas_fuentes", headers={"Prefer": "return=minimal"},
             json={"fuente": f["fuente"], "lista": f["lista"], "supuesto": f["supuesto"], "url": f["url"]})

    lote = str(uuid.uuid4())
    try:
        for i in range(0, len(filas), LOTE_INSERT):
            bloque = [{**x, "fuente": f["fuente"], "lote": lote, "lista": f["lista"]}
                      for x in filas[i:i + LOTE_INSERT]]
            rest("POST", "sat_listas", json=bloque, headers={"Prefer": "return=minimal"})
    except Exception:
        # Se retira lo que alcanzó a entrar; el lote anterior sigue vigente.
        try:
            rest("DELETE", "sat_listas", params={"lote": f"eq.{lote}"})
        except Exception as e:
            print(f"    Aviso: no se pudo limpiar el lote incompleto ({e})")
        raise

    # Recién ahora el lote nuevo pasa a ser el que la app ve.
    rest("PATCH", "sat_listas_fuentes", params={"fuente": f"eq.{f['fuente']}"},
         headers={"Prefer": "return=minimal"},
         json={"lote_vigente": lote, "hash": huella, "filas": len(filas), "url": f["url"],
               "supuesto": f["supuesto"], "actualizado_en": ahora, "revisado_en": ahora})

    # Y se borra el anterior. Si esto falla no pasa nada visible: la vista ya
    # no lo muestra, y se limpia en la siguiente carga de esta fuente.
    try:
        rest("DELETE", "sat_listas", params={"fuente": f"eq.{f['fuente']}", "lote": f"neq.{lote}"})
    except Exception as e:
        print(f"    Aviso: quedó el lote anterior sin borrar ({e})")

    extra = f", {descartadas} renglón(es) sin RFC válido" if descartadas else ""
    print(f"    cargadas {len(filas)} RFC{extra}")
    return "actualizada", len(filas)


def main():
    print(f"--- {datetime.now():%Y-%m-%d %H:%M:%S} ---")
    actualizadas = sin_cambios = fallidas = 0
    for f in FUENTES:
        try:
            estado, _ = cargar_fuente(f)
            if estado == "actualizada":
                actualizadas += 1
            else:
                sin_cambios += 1
        except Exception as e:
            # Una fuente que falla no detiene a las demás.
            print(f"    ERROR: {e}")
            fallidas += 1
    if DESCARTES:
        with open(ARCHIVO_DESCARTES, "w", encoding="utf8") as fh:
            fh.write("\n".join(DESCARTES))
        print(f"Renglones descartados guardados en {ARCHIVO_DESCARTES}")
    print(f"Actualizadas {actualizadas}, sin cambios {sin_cambios}, con error {fallidas}.")
    sys.exit(1 if fallidas else 0)


if __name__ == "__main__":
    main()