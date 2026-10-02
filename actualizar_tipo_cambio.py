#!/usr/bin/env python3
"""
Descarga el tipo de cambio oficial pesos por dólar de la API del SIE de Banxico
y lo guarda en la tabla tipos_cambio de Supabase.

Series:
  SF43718  FIX, por fecha de determinación: el día que Banxico lo calcula
           (a partir de las 12:00). Se publica en el DOF el día hábil siguiente.
  SF60653  "Para pagos": el tipo de cambio con que se convierten ese día las
           obligaciones en dólares — el publicado en el DOF el día hábil
           bancario anterior (= FIX determinado dos días hábiles antes). Trae
           valor también en fin de semana. ES EL QUE USA LA APP.

Cada corrida pide los últimos 15 días, no solo el de hoy: si la tarea dejó de
correr unos días (vacaciones, la máquina apagada), la siguiente corrida rellena
los huecos sola. Lo que ya existe se sobrescribe con el mismo valor, así que
repetir una corrida no duplica nada.

Primera carga del histórico:
    python actualizar_tipo_cambio.py --desde 2026-01-01

Configuración: el mismo config.env de los otros scripts
(%USERPROFILE%\\.smi\\config.env), con una línea más:
    BANXICO_TOKEN=<tu token del SIE>

Requiere 61-tipos-cambio.sql
"""

import os
import sys
from datetime import date, datetime, timedelta, timezone

import requests

# --------------------------------------------------------------------------
# Configuración — mismo archivo y mismas reglas que los otros scripts
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
BANXICO_TOKEN = os.environ.get("BANXICO_TOKEN", "")

faltan = [n for n, v in (("SUPABASE_URL", SUPABASE_URL), ("SUPABASE_SERVICE_KEY", SERVICE_KEY),
                         ("BANXICO_TOKEN", BANXICO_TOKEN)) if not v]
if faltan:
    sys.exit(f"Faltan {', '.join(faltan)}. Ponlos en {ARCHIVO_CONF}.")

SERIES = {"SF43718": "FIX (fecha de determinación)", "SF60653": "Para pagos"}
API = "https://www.banxico.org.mx/SieAPIRest/service/v1/series/{series}/datos/{ini}/{fin}"
DIAS_ATRAS = 15


def descargar(desde, hasta):
    """Filas {serie, fecha, valor} de Banxico para el rango. Los días sin dato
       ("N/E") se omiten: Banxico no publica en días inhábiles."""
    r = requests.get(API.format(series=",".join(SERIES), ini=desde.isoformat(), fin=hasta.isoformat()),
                     headers={"Bmx-Token": BANXICO_TOKEN, "Accept": "application/json"}, timeout=60)
    if r.status_code == 401:
        raise RuntimeError("Banxico rechazó el token (401). Revisa BANXICO_TOKEN en config.env.")
    r.raise_for_status()
    filas = []
    for s in r.json().get("bmx", {}).get("series", []):
        for d in s.get("datos") or []:
            valor = str(d.get("dato", "")).replace(",", "").strip()
            try:
                v = float(valor)
            except ValueError:
                continue  # "N/E": sin dato ese día
            f = datetime.strptime(d["fecha"], "%d/%m/%Y").date()
            filas.append({"serie": s["idSerie"], "fecha": f.isoformat(), "valor": v})
    return filas


def guardar(filas):
    """Inserta o actualiza por (serie, fecha), en bloques."""
    cab = {"apikey": SERVICE_KEY, "Authorization": f"Bearer {SERVICE_KEY}",
           "Content-Type": "application/json",
           "Prefer": "resolution=merge-duplicates,return=minimal"}
    ahora = datetime.now(timezone.utc).isoformat()
    for i in range(0, len(filas), 500):
        bloque = [{**x, "actualizado_en": ahora} for x in filas[i:i + 500]]
        r = requests.post(f"{SUPABASE_URL}/rest/v1/tipos_cambio?on_conflict=serie,fecha",
                          headers=cab, json=bloque, timeout=60)
        if not r.ok:
            raise RuntimeError(f"Supabase {r.status_code}: {r.text[:300]}")


def main():
    print(f"--- {datetime.now():%Y-%m-%d %H:%M:%S} ---")
    hoy = date.today()
    desde = hoy - timedelta(days=DIAS_ATRAS)
    if "--desde" in sys.argv:
        try:
            desde = date.fromisoformat(sys.argv[sys.argv.index("--desde") + 1])
        except (IndexError, ValueError):
            sys.exit("Uso: --desde AAAA-MM-DD")

    try:
        filas = descargar(desde, hoy)
        if not filas:
            print(f"  Banxico no devolvió datos entre {desde} y {hoy}.")
            sys.exit(0)
        guardar(filas)
    except Exception as e:
        print(f"  ERROR: {e}")
        sys.exit(1)

    for serie, nombre in SERIES.items():
        de_serie = sorted((x for x in filas if x["serie"] == serie), key=lambda x: x["fecha"])
        if de_serie:
            u = de_serie[-1]
            print(f"  {serie} {nombre}: {len(de_serie)} día(s); último {u['fecha']} = {u['valor']:.4f}")
    print(f"Guardados {len(filas)} valores del {desde} al {hoy}.")


if __name__ == "__main__":
    main()
