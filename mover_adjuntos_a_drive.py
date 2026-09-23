#!/usr/bin/env python3
"""
Mueve los adjuntos de solicitudes del búfer de Supabase a Google Drive.

Es la pieza que reemplaza lo que Zoho hacía por detrás: el navegador entrega
el archivo a Supabase, que sí acepta carga anónima, y este proceso lo lleva a
Drive con credenciales que nunca tocan el navegador.

Qué hace, por cada adjunto en tránsito:
  1. Busca o crea Compras/<COMPAÑÍA>/<AÑO>/<FOLIO> en Drive. La raíz que se
     configura en los parámetros es la de la compañía; el nivel del año lo
     crea este proceso, para no tener que cambiar el parámetro cada enero.
  2. Descarga el archivo del búfer.
  3. Lo sube a Drive.
  4. Marca el adjunto como en_drive con su fileId y su enlace.
  5. Recién entonces borra el objeto del búfer.

El orden de los pasos 4 y 5 no es casual. Si se borrara el búfer antes de
grabar el fileId y algo fallara en medio, el archivo se perdería sin rastro.
Al revés, lo peor que pasa es que quede un objeto huérfano en el búfer, que se
recoge después sin haber perdido nada.

Configuración, en %USERPROFILE%\\.smi\\config.env o en el entorno:
  SUPABASE_URL          https://tu-proyecto.supabase.co
  SUPABASE_SERVICE_KEY  la llave service_role
  GOOGLE_CREDENCIALES   ruta al client_secret.json (por omisión, junto al config)
  GOOGLE_TOKEN          dónde guardar el token   (por omisión, junto al config)
  ALERTA_CORREO         a quién avisar cuando un archivo se rinde (opcional)
  REMITENTE             desde qué cuenta salen los correos (opcional)

La llave service_role omite las políticas de seguridad a propósito: este
proceso no actúa como ningún usuario. NUNCA debe llegar al navegador ni
quedar en el repositorio.
"""

import os
import sys
import io as _io
import mimetypes
import base64
from email.message import EmailMessage
from datetime import datetime, timezone

import requests
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

# --------------------------------------------------------------------------
# Configuración
#
# Se lee de un archivo fuera del repositorio, y las variables de entorno lo
# sobreescriben si están puestas. Así una tarea programada no necesita que
# alguien exporte nada antes de correr, y los secretos no viven donde Git
# pueda verlos.
#
# Por omisión: %USERPROFILE%\.smi\config.env  (o ~/.smi/config.env)
# --------------------------------------------------------------------------
CARPETA_CONF = os.environ.get("SMI_CONF_DIR") or os.path.join(os.path.expanduser("~"), ".smi")
ARCHIVO_CONF = os.path.join(CARPETA_CONF, "config.env")


def _cargar_conf():
    """Lee CLAVE=valor del archivo de configuración. Ignora comentarios y
       líneas vacías, y no pisa lo que ya venga del entorno."""
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
SERVICE_KEY  = os.environ.get("SUPABASE_SERVICE_KEY", "")
CREDENCIALES = os.environ.get("GOOGLE_CREDENCIALES") or os.path.join(CARPETA_CONF, "credenciales.json")
TOKEN        = os.environ.get("GOOGLE_TOKEN") or os.path.join(CARPETA_CONF, "token.json")
BUCKET       = "solicitudes-adjuntos"

# A quién se le avisa cuando un archivo se rinde. Vacío = no se avisa.
ALERTA_CORREO = os.environ.get("ALERTA_CORREO", "")
# Desde qué cuenta salen los correos. Vacío = la que autorizó el token.
REMITENTE     = os.environ.get("REMITENTE", "me")

# Dos scopes, y la combinación importa:
#
#   drive.file             crear y escribir SOLO lo que esta aplicación crea.
#   drive.metadata.readonly ver nombres y jerarquía del resto del Drive, sin
#                          poder leer el contenido de ningún archivo.
#
# Con drive.file a secas, una carpeta creada por otro —Zoho, o tú a mano— es
# invisible para el script: la busca, no la encuentra y crea una duplicada.
# Fue exactamente lo que pasó con la carpeta 2026 que ya existía.
#
# metadata.readonly resuelve la búsqueda sin abrir el contenido de nada, y la
# escritura sigue limitada a lo que este script produce.
SCOPES = [
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/drive.metadata.readonly",
    # Solo enviar. No permite leer ni borrar correo.
    "https://www.googleapis.com/auth/gmail.send",
]

if not SUPABASE_URL or not SERVICE_KEY:
    sys.exit(f"Faltan SUPABASE_URL o SUPABASE_SERVICE_KEY.\n"
             f"Ponlos en {ARCHIVO_CONF} o en el entorno.")

CAB = {
    "apikey": SERVICE_KEY,
    "Authorization": f"Bearer {SERVICE_KEY}",
    "Content-Type": "application/json",
}


# --------------------------------------------------------------------------
# Supabase
# --------------------------------------------------------------------------
def rest(metodo, ruta, **kw):
    r = requests.request(metodo, f"{SUPABASE_URL}/rest/v1/{ruta}", headers=CAB, timeout=60, **kw)
    r.raise_for_status()
    return r.json() if r.content and r.headers.get("content-type", "").startswith("application/json") else None


MAX_INTENTOS = 5


def pendientes():
    """Adjuntos que ya están en el búfer y todavía no en Drive.

       Incluye los que fallaron antes. Guardar el error y no volver a
       intentarlos dejaría el archivo atrapado en el búfer para siempre, que
       es justo lo contrario de lo que el manejo de errores pretende.

       El tope de intentos evita el otro extremo: un adjunto roto reintentado
       en cada corrida, para siempre, llenando el registro de errores."""
    return rest("GET", "adjuntos",
                params={"select": "*",
                        "estado": "in.(en_transito,error)",
                        "intentos": f"lt.{MAX_INTENTOS}",
                        "storage_path": "not.is.null",
                        "order": "subido_en.asc"})


def expediente_de(adj):
    """A qué expediente pertenece el adjunto.

       Devuelve un diccionario con la tabla, el id, el folio que nombra la
       carpeta, la compañía, el año y la carpeta ya resuelta si la tiene.

       Tres orígenes:
         * solicitud            — cotizaciones, soporte, el PDF de la SMI
         * solicitud_concepto   — las especificaciones de una línea
         * transaccion          — la póliza, comprobantes, facturas

       La regla del expediente es una: se llama como el folio de lo que lo
       originó. Una transacción que viene de una SMI comparte la carpeta de
       esa SMI —el comprobante queda junto a la cotización que lo originó— y
       una que no viene de ninguna usa su propio folio."""
    ent, eid = adj["entidad"], adj["entidad_id"]

    if ent == "solicitud_concepto":
        conc = rest("GET", "solicitud_conceptos",
                    params={"select": "solicitud_id", "id": f"eq.{eid}"})
        if not conc:
            return None
        ent, eid = "solicitud", conc[0]["solicitud_id"]

    if ent == "solicitud":
        f = rest("GET", "solicitudes",
                 params={"select": "id,folio,unidad,anio,drive_folder_id", "id": f"eq.{eid}"})
        if not f:
            return None
        s = f[0]
        return {"tabla": "solicitudes", "id": s["id"], "folio": s["folio"],
                "unidad": s["unidad"], "anio": s.get("anio"),
                "carpeta": s.get("drive_folder_id")}

    if ent == "transaccion":
        f = rest("GET", "transacciones",
                 params={"select": "id,folio_transaccion,unidad_detectada,dia,"
                                   "registrada_en,drive_folder_id,solicitud_id",
                         "id": f"eq.{eid}"})
        if not f:
            return None
        t = f[0]

        # Sin registrar no hay folio, y sin folio no hay nombre de carpeta.
        # Se deja en el búfer: no es un error, es que todavía no toca.
        if not t.get("registrada_en") or not t.get("folio_transaccion"):
            return "pendiente"

        # Si vino de una SMI, va a la carpeta de esa SMI.
        if t.get("solicitud_id"):
            s = rest("GET", "solicitudes",
                     params={"select": "id,folio,unidad,anio,drive_folder_id",
                             "id": f"eq.{t['solicitud_id']}"})
            if s:
                return {"tabla": "solicitudes", "id": s[0]["id"], "folio": s[0]["folio"],
                        "unidad": s[0]["unidad"], "anio": s[0].get("anio"),
                        "carpeta": s[0].get("drive_folder_id")}

        anio = int(str(t["dia"])[:4]) if t.get("dia") else datetime.now().year
        return {"tabla": "transacciones", "id": t["id"], "folio": t["folio_transaccion"],
                "unidad": t.get("unidad_detectada"), "anio": anio,
                "carpeta": t.get("drive_folder_id")}

    return None


def carpeta_raiz(unidad):
    filas = rest("GET", "solicitudes_parametros",
                 params={"select": "drive_carpeta_raiz", "unidad": f"eq.{unidad}"})
    return (filas[0]["drive_carpeta_raiz"] if filas else None) or None


def descargar_del_bufer(ruta):
    url = f"{SUPABASE_URL}/storage/v1/object/{BUCKET}/{ruta}"
    r = requests.get(url, headers={"apikey": SERVICE_KEY, "Authorization": f"Bearer {SERVICE_KEY}"}, timeout=120)
    r.raise_for_status()
    return r.content


def borrar_del_bufer(ruta):
    url = f"{SUPABASE_URL}/storage/v1/object/{BUCKET}/{ruta}"
    r = requests.delete(url, headers={"apikey": SERVICE_KEY, "Authorization": f"Bearer {SERVICE_KEY}"}, timeout=60)
    r.raise_for_status()


def marcar(adj_id, campos):
    requests.patch(f"{SUPABASE_URL}/rest/v1/adjuntos",
                   headers={**CAB, "Prefer": "return=minimal"},
                   params={"id": f"eq.{adj_id}"}, json=campos, timeout=60).raise_for_status()


def fallo(adj, mensaje):
    """Se guarda el error y se cuenta el intento. El archivo NO se borra del
       búfer: si el problema fue de Drive, se vuelve a intentar en la
       siguiente corrida sin haber perdido nada."""
    marcar(adj["id"], {"estado": "error",
                       "ultimo_error": str(mensaje)[:500],
                       "intentos": (adj.get("intentos") or 0) + 1})
    print(f"    ERROR: {mensaje}")


# --------------------------------------------------------------------------
# Drive
# --------------------------------------------------------------------------
def credenciales_google():
    """Un solo juego de credenciales para Drive y Gmail: dos tokens separados
       significarían dos autorizaciones que caducan por su cuenta."""
    cred = None
    if os.path.exists(TOKEN):
        cred = Credentials.from_authorized_user_file(TOKEN, SCOPES)
        # Un token emitido con menos scopes de los que ahora se piden no
        # sirve: Google no los amplía solos. Se descarta y se vuelve a
        # autorizar, que es lo que hay que hacer al agregar metadata.readonly.
        if cred and set(SCOPES) - set(cred.scopes or []):
            print("  El token no cubre los permisos actuales. Hay que autorizar otra vez.")
            cred = None
    if not cred or not cred.valid:
        if cred and cred.expired and cred.refresh_token:
            cred.refresh(Request())
        else:
            if not os.path.exists(CREDENCIALES):
                sys.exit(f"No encuentro {CREDENCIALES}. Descarga el client_secret de una app de escritorio.")
            cred = InstalledAppFlow.from_client_secrets_file(CREDENCIALES, SCOPES).run_local_server(port=0)
        with open(TOKEN, "w", encoding="utf8") as fh:
            fh.write(cred.to_json())
    return cred


def subcarpeta(drive, padre, nombre):
    """Busca una carpeta por nombre bajo `padre`; si no está, la crea.

       Se busca antes de crear porque el proceso puede correr varias veces
       sobre la misma solicitud —un adjunto nuevo, un reintento— y Drive
       permite dos carpetas con el mismo nombre sin protestar: acabarías con
       dos 'OSB-264-26' y los archivos repartidos entre ambas."""
    seguro = str(nombre).replace("\\", "\\\\").replace("'", "\\'")
    q = (f"name = '{seguro}' and mimeType = 'application/vnd.google-apps.folder' "
         f"and '{padre}' in parents and trashed = false")
    res = drive.files().list(q=q, fields="files(id)", pageSize=1,
                             supportsAllDrives=True, includeItemsFromAllDrives=True).execute()
    if res.get("files"):
        return res["files"][0]["id"]
    creada = drive.files().create(
        body={"name": str(nombre), "mimeType": "application/vnd.google-apps.folder", "parents": [padre]},
        fields="id", supportsAllDrives=True).execute()
    return creada["id"]


def carpeta_de_solicitud(drive, raiz, anio, folio):
    """Compras/<COMPAÑÍA>/<AÑO>/<FOLIO>.

       La raíz configurada en los parámetros es la de la compañía; el nivel
       del año lo crea el proceso. Si el año estuviera en la raíz habría que
       cambiar el parámetro a mano cada enero, y el primero que capturara en
       enero dejaría su solicitud en la carpeta del año anterior."""
    return subcarpeta(drive, subcarpeta(drive, raiz, anio), folio)


def subir_a_drive(drive, carpeta, nombre, contenido, mime):
    medio = MediaIoBaseUpload(_io.BytesIO(contenido),
                              mimetype=mime or mimetypes.guess_type(nombre)[0] or "application/octet-stream",
                              resumable=len(contenido) > 5 * 1024 * 1024)
    f = drive.files().create(body={"name": nombre, "parents": [carpeta]},
                             media_body=medio, fields="id,webViewLink",
                             supportsAllDrives=True).execute()
    return f["id"], f.get("webViewLink")


# --------------------------------------------------------------------------
# Correo
# --------------------------------------------------------------------------
def enviar_correo(gmail, para, asunto, cuerpo):
    """Texto plano a propósito: un acuse no necesita formato, y el texto llega
       igual de bien a un cliente de correo viejo que a un teléfono."""
    msg = EmailMessage()
    msg["To"] = para
    msg["Subject"] = asunto
    if REMITENTE and REMITENTE != "me":
        msg["From"] = REMITENTE
    msg.set_content(cuerpo)
    gmail.users().messages().send(
        userId="me",
        body={"raw": base64.urlsafe_b64encode(msg.as_bytes()).decode()}).execute()


def acuses(gmail):
    """Confirma la recepción a quien levantó una solicitud.

       Solo las que ya tienen todos sus archivos en Drive: avisar antes
       significaría que si alguien abre el expediente en ese momento, lo
       encuentra a medias."""
    pend = rest("GET", "solicitudes",
                params={"select": "id,folio,unidad,correo_solicitante,nombre_solicitante,"
                                  "descripcion_general,total,divisa,creado_en",
                        "acuse_enviado_en": "is.null",
                        "estado": "neq.cancelada",
                        "order": "creado_en.asc", "limit": "50"}) or []
    enviados = 0
    for s in pend:
        sin_mover = rest("GET", "adjuntos",
                         params={"select": "id", "entidad": "eq.solicitud",
                                 "entidad_id": f"eq.{s['id']}",
                                 "estado": "in.(en_transito,error)", "limit": "1"}) or []
        if sin_mover:
            continue

        cuerpo = (
            f"Hola {s.get('nombre_solicitante') or ''},\n\n"
            f"Recibimos tu solicitud y quedó registrada con el folio {s['folio']}.\n\n"
            f"  Qué pediste: {s.get('descripcion_general') or ''}\n"
            f"  Importe estimado: ${float(s.get('total') or 0):,.2f} "
            f"{'USD' if s.get('divisa') == 'USD' else 'MXN'}\n\n"
            f"Administración la va a revisar. Si hace falta algo más, te buscamos.\n\n"
            f"Cita el folio {s['folio']} para cualquier seguimiento.\n")
        try:
            enviar_correo(gmail, s["correo_solicitante"],
                          f"Solicitud {s['folio']} recibida", cuerpo)
            requests.patch(f"{SUPABASE_URL}/rest/v1/solicitudes",
                           headers={**CAB, "Prefer": "return=minimal"},
                           params={"id": f"eq.{s['id']}"},
                           json={"acuse_enviado_en": datetime.now(timezone.utc).isoformat()},
                           timeout=60).raise_for_status()
            enviados += 1
            print(f"  acuse -> {s['correo_solicitante']} ({s['folio']})")
        except Exception as e:
            print(f"  No se pudo avisar de {s['folio']}: {e}")
    return enviados


def alertar_atorados(gmail):
    """Avisa de los archivos que se rindieron. Corriendo desatendido, sin esto
       un adjunto atorado pasa inadvertido hasta que alguien lo reclama."""
    if not ALERTA_CORREO:
        return 0
    rotos = rest("GET", "adjuntos",
                 params={"select": "id,nombre,entidad,entidad_id,ultimo_error",
                         "estado": "eq.error",
                         "intentos": f"gte.{MAX_INTENTOS}",
                         "alertado_en": "is.null"}) or []
    if not rotos:
        return 0
    lineas = [f"  {a['nombre']}: {a.get('ultimo_error') or 'sin detalle'}" for a in rotos]
    cuerpo = (f"{len(rotos)} archivo(s) no se pudieron mover a Drive tras "
              f"{MAX_INTENTOS} intentos:\n\n" + "\n".join(lineas) +
              "\n\nSiguen en el búfer de Supabase. Para reintentarlos, pon "
              "intentos = 0 en esas filas de la tabla adjuntos.\n")
    try:
        enviar_correo(gmail, ALERTA_CORREO, f"{len(rotos)} adjunto(s) atorados", cuerpo)
        for a in rotos:
            marcar(a["id"], {"alertado_en": datetime.now(timezone.utc).isoformat()})
        print(f"  alerta -> {ALERTA_CORREO} ({len(rotos)} atorados)")
        return len(rotos)
    except Exception as e:
        print(f"  No se pudo alertar: {e}")
        return 0


# --------------------------------------------------------------------------
def main():
    print(f"--- {datetime.now():%Y-%m-%d %H:%M:%S} ---")
    lista = pendientes() or []

    # El servicio se construye siempre, aunque no haya nada que mover: los
    # acuses y las alertas también salen por aquí, y en la mayoría de las
    # corridas eso es lo único que hay que hacer.
    cred = credenciales_google()
    drive = build("drive", "v3", credentials=cred, cache_discovery=False)
    gmail = build("gmail", "v1", credentials=cred, cache_discovery=False)

    movidos = fallidos = pospuestos = 0
    if lista:
        print(f"{len(lista)} adjunto(s) por mover.")
    else:
        print("Sin archivos por mover.")

    for adj in lista:
        print(f"  {adj['nombre']}")
        try:
            exp = expediente_de(adj)

            if exp == "pendiente":
                # Una transacción sin registrar todavía no tiene folio. Se
                # queda en el búfer sin contarse como error: no falló nada,
                # simplemente aún no le toca.
                print("    en espera: la transacción no está registrada")
                pospuestos += 1
                continue

            if not exp:
                fallo(adj, f"El adjunto no pertenece a ninguna {adj['entidad']} existente.")
                fallidos += 1
                continue

            # La carpeta se resuelve una sola vez por solicitud y se guarda:
            # a partir de ahí, los adjuntos siguientes no vuelven a buscarla.
            # La carpeta se resuelve una sola vez por expediente y se guarda:
            # a partir de ahí, los adjuntos siguientes no vuelven a buscarla.
            carpeta = exp.get("carpeta")
            if not carpeta:
                if not exp.get("unidad"):
                    fallo(adj, "El expediente no tiene compañía: no se sabe bajo qué carpeta raíz va.")
                    fallidos += 1
                    continue
                raiz = carpeta_raiz(exp["unidad"])
                if not raiz:
                    fallo(adj, f"{exp['unidad']} no tiene carpeta raíz de Drive en sus parámetros.")
                    fallidos += 1
                    continue
                anio = exp.get("anio") or datetime.now().year
                carpeta = carpeta_de_solicitud(drive, raiz, anio, exp["folio"])
                requests.patch(f"{SUPABASE_URL}/rest/v1/{exp['tabla']}",
                               headers={**CAB, "Prefer": "return=minimal"},
                               params={"id": f"eq.{exp['id']}"},
                               json={"drive_folder_id": carpeta}, timeout=60).raise_for_status()

            contenido = descargar_del_bufer(adj["storage_path"])
            file_id, enlace = subir_a_drive(drive, carpeta, adj["nombre"], contenido, adj.get("mime"))

            # Primero se graba el destino; el búfer se vacía después.
            # La fecha se calcula aquí: PostgREST manda el JSON tal cual y
            # Postgres no puede convertir la cadena "now()" a timestamp.
            marcar(adj["id"], {"drive_file_id": file_id, "drive_link": enlace,
                               "estado": "en_drive", "ultimo_error": None,
                               "verificado_en": datetime.now(timezone.utc).isoformat()})
            try:
                borrar_del_bufer(adj["storage_path"])
                marcar(adj["id"], {"storage_path": None})
            except Exception as e:
                # El archivo ya está a salvo en Drive. El objeto huérfano se
                # recoge en otra corrida.
                print(f"    Aviso: quedó en el búfer ({e})")

            print(f"    -> {file_id}")
            movidos += 1
        except Exception as e:
            fallo(adj, e)
            fallidos += 1

    enviados = acuses(gmail)
    alertados = alertar_atorados(gmail)

    print(f"Movidos {movidos}, en espera {pospuestos}, con error {fallidos}, "
          f"acuses {enviados}, alertas {alertados}.")


if __name__ == "__main__":
    main()