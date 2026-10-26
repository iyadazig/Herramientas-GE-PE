# -*- coding: utf-8 -*-
"""
ALMACEN COMPARTIDO (SQLite)
===========================
Un unico fichero `revisor_ssaa.db` en la carpeta de datos (config.CARPETA_DATOS) con:

  usuarios              usuario, nombre, contrasena (PBKDF2-SHA256 con sal), admin, supervisor
                        (ve el panel del equipo), activo
  contratos             ficha vigente de cada CUPS y quien/cuando la guardo
  contratos_historial   todas las versiones de cada ficha (quien, cuando, que)
  revisiones            cada revision: usuario, fecha, factura, importes, veredicto e informe
  solicitudes_ofertas   cada solicitud de ofertas guardada: quien, cuando, cliente, CUPS,
                        periodo de recepcion, seguimiento (estado, comercializadoras), Excel,
                        texto del correo y los datos del formulario para volver a cargarla
  estudios_potencia     cada estudio de potencia guardado: quien, cuando, titular, CUPS,
                        ahorro e inversion de la propuesta optima, datos del formulario, los
                        suministros con sus curvas ya leidas (comprimidos) y el PDF
  actividad             lo que no se guarda en otra tabla: accesos, paginas abiertas, consultas
                        a Gemweb y descargas (quien, cuando, modulo, accion, detalle breve)

SQLite en modo WAL admite varios compañeros a la vez desde la misma app (servidor).
La primera vez se importan las fichas de contratos_ssaa.json si existe.
"""

import datetime as dt
import hashlib
import hmac
import json
import os
import re
import secrets
import sqlite3
import threading

from contextlib import contextmanager

import config

ITERACIONES = 240_000
_bloqueo = threading.Lock()
_iniciado = set()

ESQUEMA = """
CREATE TABLE IF NOT EXISTS usuarios (
    usuario TEXT PRIMARY KEY, nombre TEXT NOT NULL, sal TEXT NOT NULL, clave TEXT NOT NULL,
    admin INTEGER NOT NULL DEFAULT 0, activo INTEGER NOT NULL DEFAULT 1,
    cambiar_clave INTEGER NOT NULL DEFAULT 0, creado TEXT NOT NULL, ultimo_acceso TEXT,
    intentos_fallidos INTEGER NOT NULL DEFAULT 0, bloqueado_hasta TEXT);
CREATE TABLE IF NOT EXISTS contratos (
    cups TEXT PRIMARY KEY, datos TEXT NOT NULL, actualizado TEXT NOT NULL, usuario TEXT);
CREATE TABLE IF NOT EXISTS contratos_historial (
    id INTEGER PRIMARY KEY AUTOINCREMENT, cups TEXT NOT NULL, datos TEXT NOT NULL,
    fecha TEXT NOT NULL, usuario TEXT, accion TEXT NOT NULL);
CREATE TABLE IF NOT EXISTS revisiones (
    id INTEGER PRIMARY KEY AUTOINCREMENT, fecha TEXT NOT NULL, usuario TEXT,
    cups TEXT, comercializadora TEXT, factura TEXT, emision TEXT, periodo TEXT,
    clausula TEXT, lineas INTEGER, facturado REAL, recalculado REAL, diferencia REAL,
    veredicto TEXT, detalle TEXT, informe BLOB);
CREATE INDEX IF NOT EXISTS ix_rev_cups ON revisiones(cups);
CREATE TABLE IF NOT EXISTS solicitudes_ofertas (
    id INTEGER PRIMARY KEY AUTOINCREMENT, fecha TEXT NOT NULL, usuario TEXT,
    ref TEXT, cliente TEXT, cups TEXT, suministros INTEGER, consumo REAL,
    recepcion_desde TEXT, recepcion_hasta TEXT, estado TEXT NOT NULL,
    comercializadoras TEXT, actualizado TEXT, actualizado_por TEXT,
    datos TEXT NOT NULL, texto TEXT, excel BLOB);
CREATE TABLE IF NOT EXISTS estudios_potencia (
    id INTEGER PRIMARY KEY AUTOINCREMENT, fecha TEXT NOT NULL, usuario TEXT,
    tipo TEXT NOT NULL, titular TEXT, cups TEXT, suministros INTEGER, tarifa TEXT,
    ahorro REAL, inversion REAL, datos TEXT NOT NULL, curvas BLOB, pdf BLOB);
CREATE INDEX IF NOT EXISTS ix_hist_cups ON contratos_historial(cups);
CREATE TABLE IF NOT EXISTS actividad (
    id INTEGER PRIMARY KEY AUTOINCREMENT, fecha TEXT NOT NULL, usuario TEXT,
    modulo TEXT, accion TEXT NOT NULL, detalle TEXT);
CREATE INDEX IF NOT EXISTS ix_act_fecha ON actividad(fecha);
"""


def ruta_bd():
    return config.CARPETA_DATOS / "revisor_ssaa.db"


def _ahora():
    return dt.datetime.now().isoformat(timespec="seconds")


def conectar():
    ruta = ruta_bd()
    ruta.parent.mkdir(parents=True, exist_ok=True)
    # isolation_level=None: las transacciones se abren a mano en _con()
    con = sqlite3.connect(ruta, timeout=30, check_same_thread=False, isolation_level=None)
    con.row_factory = sqlite3.Row
    con.execute("PRAGMA busy_timeout=30000")
    clave = str(ruta)
    if clave not in _iniciado:
        with _bloqueo:
            if clave not in _iniciado:
                con.execute("PRAGMA journal_mode=WAL")     # queda guardado en el fichero
                con.executescript(ESQUEMA)
                _migrar(con)
                _importar_json(con)
                _iniciado.add(clave)
    return con


def _migrar(con):
    """Columnas anadidas despues de crear la base de datos (las bases antiguas no las tienen)."""
    columnas = {r["name"] for r in con.execute("PRAGMA table_info(usuarios)")}
    if "email" not in columnas:
        con.execute("ALTER TABLE usuarios ADD COLUMN email TEXT")
    if "supervisor" not in columnas:
        con.execute("ALTER TABLE usuarios ADD COLUMN supervisor INTEGER NOT NULL DEFAULT 0")


@contextmanager
def _con(escribir=False):
    """Conexion en una transaccion que se confirma al salir (o se deshace si hay error) y
    se cierra siempre. Para escribir se reserva la base de datos al empezar (BEGIN
    IMMEDIATE): si otro compañero esta guardando, se espera en vez de fallar."""
    con = conectar()
    try:
        con.execute("BEGIN IMMEDIATE" if escribir else "BEGIN")
        try:
            yield con
        except BaseException:
            con.execute("ROLLBACK")
            raise
        con.execute("COMMIT")
    finally:
        con.close()


def _importar_json(con):
    """Primera vez: pasa las fichas de contratos_ssaa.json a la base de datos."""
    if con.execute("SELECT COUNT(*) FROM contratos").fetchone()[0]:
        return
    if not config.FICHERO_CONTRATOS.exists():
        return
    with open(config.FICHERO_CONTRATOS, encoding="utf-8") as f:
        fichas = json.load(f)
    ahora = _ahora()
    con.execute("BEGIN IMMEDIATE")
    try:
        for cups, datos in fichas.items():
            txt = json.dumps(datos, ensure_ascii=False, sort_keys=True)
            con.execute("INSERT OR IGNORE INTO contratos VALUES (?,?,?,?)",
                        (cups, txt, ahora, "importado de contratos_ssaa.json"))
            con.execute("INSERT INTO contratos_historial (cups, datos, fecha, usuario, accion) "
                        "VALUES (?,?,?,?,?)", (cups, txt, ahora, None, "importada"))
        con.execute("COMMIT")
    except BaseException:
        con.execute("ROLLBACK")
        raise


# ------------------------------------------------------------------- usuarios
def _hash(clave, sal):
    return hashlib.pbkdf2_hmac("sha256", clave.encode("utf-8"), bytes.fromhex(sal),
                               ITERACIONES).hex()


def hay_usuarios():
    with _con() as con:
        return con.execute("SELECT COUNT(*) FROM usuarios").fetchone()[0] > 0


def validar_clave(clave):
    """Mensaje de error o None. Minimo 10 caracteres con letras y numeros."""
    if len(clave) < 10:
        return "La contraseña debe tener al menos 10 caracteres."
    if not any(c.isdigit() for c in clave) or not any(c.isalpha() for c in clave):
        return "La contraseña debe tener letras y números."
    return None


def crear_usuario(usuario, nombre, clave, admin=False, cambiar_clave=True, email=None,
                  supervisor=False):
    usuario = usuario.strip().lower()
    if not usuario or not nombre.strip():
        raise ValueError("Usuario y nombre son obligatorios.")
    error = validar_clave(clave)
    if error:
        raise ValueError(error)
    sal = secrets.token_hex(16)
    with _con(escribir=True) as con:
        if con.execute("SELECT 1 FROM usuarios WHERE usuario=?", (usuario,)).fetchone():
            raise ValueError("Ya existe el usuario %s." % usuario)
        con.execute("INSERT INTO usuarios (usuario, nombre, sal, clave, admin, activo, "
                    "cambiar_clave, creado, email, supervisor) VALUES (?,?,?,?,?,1,?,?,?,?)",
                    (usuario, nombre.strip(), sal, _hash(clave, sal), int(admin),
                     int(cambiar_clave), _ahora(), _email(email), int(supervisor)))


def comprobar(usuario, clave):
    """dict del usuario si la contrasena es correcta; ValueError con el motivo si no.
    Tras 5 intentos fallidos el usuario queda bloqueado 15 minutos."""
    usuario = usuario.strip().lower()
    error = None
    with _con(escribir=True) as con:
        u = con.execute("SELECT * FROM usuarios WHERE usuario=?", (usuario,)).fetchone()
        if u is None or not u["activo"]:
            error = "Usuario o contraseña incorrectos."
        elif u["bloqueado_hasta"] and u["bloqueado_hasta"] > _ahora():
            error = ("Demasiados intentos fallidos. Prueba de nuevo a las %s."
                     % u["bloqueado_hasta"][11:16])
        elif not hmac.compare_digest(_hash(clave, u["sal"]), u["clave"]):
            fallos = u["intentos_fallidos"] + 1
            bloqueo = (dt.datetime.now() + dt.timedelta(minutes=15)).isoformat(
                timespec="seconds") if fallos >= 5 else None
            con.execute("UPDATE usuarios SET intentos_fallidos=?, bloqueado_hasta=? "
                        "WHERE usuario=?", (0 if bloqueo else fallos, bloqueo, usuario))
            error = "Usuario o contraseña incorrectos."
        else:
            con.execute("UPDATE usuarios SET intentos_fallidos=0, bloqueado_hasta=NULL, "
                        "ultimo_acceso=? WHERE usuario=?", (_ahora(), usuario))
            datos = {k: u[k] for k in ("usuario", "nombre", "admin", "supervisor",
                                       "cambiar_clave", "email")}
    if error:                    # fuera del with: el contador de fallos ya esta guardado
        registrar_actividad(usuario, "Acceso", "Acceso fallido", error)
        raise ValueError(error)
    registrar_actividad(usuario, "Acceso", "Inicio de sesión")
    return datos


def cambiar_clave(usuario, nueva, obligar_cambio=False):
    error = validar_clave(nueva)
    if error:
        raise ValueError(error)
    sal = secrets.token_hex(16)
    with _con(escribir=True) as con:
        con.execute("UPDATE usuarios SET sal=?, clave=?, cambiar_clave=?, intentos_fallidos=0, "
                    "bloqueado_hasta=NULL WHERE usuario=?",
                    (sal, _hash(nueva, sal), int(obligar_cambio), usuario))


def listar_usuarios():
    with _con() as con:
        return [dict(r) for r in con.execute(
            "SELECT usuario, nombre, email, admin, supervisor, activo, creado, ultimo_acceso "
            "FROM usuarios "
            "ORDER BY nombre")]


def _email(email):
    """Email limpio, None si esta vacio; ValueError si no tiene forma de email."""
    email = (email or "").strip().lower()
    if not email:
        return None
    if not re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", email):
        raise ValueError("El email «%s» no es válido." % email)
    return email


def actualizar_usuario(usuario, admin=None, activo=None, nombre=None, email=None,
                       supervisor=None):
    """Si el cambio dejara sin administradores activos, se deshace y da error. email="" lo
    borra; None lo deja como esta."""
    if nombre is not None and not nombre.strip():
        raise ValueError("El nombre no puede quedar vacío.")
    if email is not None:
        email = _email(email) or ""
    with _con(escribir=True) as con:
        if email is not None:
            con.execute("UPDATE usuarios SET email=? WHERE usuario=?", (email or None, usuario))
        if nombre is not None:
            con.execute("UPDATE usuarios SET nombre=? WHERE usuario=?", (nombre.strip(), usuario))
        if admin is not None:
            con.execute("UPDATE usuarios SET admin=? WHERE usuario=?", (int(admin), usuario))
        if supervisor is not None:
            con.execute("UPDATE usuarios SET supervisor=? WHERE usuario=?",
                        (int(supervisor), usuario))
        if activo is not None:
            con.execute("UPDATE usuarios SET activo=? WHERE usuario=?", (int(activo), usuario))
        if (admin is not None or activo is not None) and not con.execute(
                "SELECT COUNT(*) FROM usuarios WHERE admin=1 AND activo=1").fetchone()[0]:
            raise ValueError("Tiene que quedar al menos un administrador activo.")


def puede_supervisar(u):
    """Los administradores y los supervisores ven el panel del equipo."""
    return bool(u and (u.get("admin") or u.get("supervisor")))


# sin caracteres que se confunden al dictarla o copiarla (l, I, 1, O, 0)
_LETRAS_CLAVE = "abcdefghjkmnpqrstuvwxyz"
_CIFRAS_CLAVE = "23456789"


def clave_temporal():
    """Contrasena provisional para altas y reseteos (se obliga a cambiarla al entrar):
    'Geype-' + dos bloques de 4, p. ej. Geype-k7mp-3xha."""
    while True:
        cuerpo = "".join(secrets.choice(_LETRAS_CLAVE + _CIFRAS_CLAVE) for _ in range(8))
        if any(c in _CIFRAS_CLAVE for c in cuerpo):
            return "Geype-%s-%s" % (cuerpo[:4], cuerpo[4:])


# ------------------------------------------------------------------- contratos
def cargar_contratos():
    with _con() as con:
        return {r["cups"]: json.loads(r["datos"]) for r in con.execute("SELECT * FROM contratos")}


def info_contratos():
    """{cups: (fecha, usuario)} de la ultima modificacion de cada ficha."""
    with _con() as con:
        return {r["cups"]: (r["actualizado"], r["usuario"])
                for r in con.execute("SELECT cups, actualizado, usuario FROM contratos")}


def guardar_contrato(cups, datos, usuario):
    txt = json.dumps(datos, ensure_ascii=False, sort_keys=True)
    ahora = _ahora()
    with _con(escribir=True) as con:
        previa = con.execute("SELECT datos FROM contratos WHERE cups=?", (cups,)).fetchone()
        if previa and previa["datos"] == txt:
            return False                              # sin cambios
        con.execute("INSERT INTO contratos VALUES (?,?,?,?) ON CONFLICT(cups) DO UPDATE SET "
                    "datos=excluded.datos, actualizado=excluded.actualizado, "
                    "usuario=excluded.usuario", (cups, txt, ahora, usuario))
        con.execute("INSERT INTO contratos_historial (cups, datos, fecha, usuario, accion) "
                    "VALUES (?,?,?,?,?)", (cups, txt, ahora, usuario,
                                           "modificada" if previa else "creada"))
    return True


def listar_historial_contratos(desde=None, hasta=None):
    """Altas y cambios de fichas (sin los datos) entre dos fechas AAAA-MM-DD, incluidas."""
    with _con() as con:
        return [dict(r) for r in con.execute(
            "SELECT id, cups, fecha, usuario, accion FROM contratos_historial "
            "WHERE substr(fecha, 1, 10) BETWEEN ? AND ? ORDER BY id DESC",
            (desde or "0000", hasta or "9999"))]


def historial_contrato(cups):
    with _con() as con:
        return [dict(r) for r in con.execute(
            "SELECT fecha, usuario, accion, datos FROM contratos_historial WHERE cups=? "
            "ORDER BY id DESC", (cups,))]


# ------------------------------------------------------------------- revisiones
def guardar_revision(usuario, cabecera, lineas, informe_xlsx):
    """cabecera: dict con cups, comercializadora, factura, emision, periodo, clausula.
    lineas: filas de resumen (dicts). Devuelve el id."""
    fact = sum(l.get("Facturado €") or 0 for l in lineas)
    rec = sum(l.get("Recalculado €") or 0 for l in lineas)
    veredictos = {l.get("Veredicto") for l in lineas}
    veredicto = veredictos.pop() if len(veredictos) == 1 else "VARIOS"
    with _con(escribir=True) as con:
        cur = con.execute(
            "INSERT INTO revisiones (fecha, usuario, cups, comercializadora, factura, emision, "
            "periodo, clausula, lineas, facturado, recalculado, diferencia, veredicto, detalle, "
            "informe) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
            (_ahora(), usuario, cabecera.get("cups"), cabecera.get("comercializadora"),
             cabecera.get("factura"), cabecera.get("emision"), cabecera.get("periodo"),
             cabecera.get("clausula"), len(lineas), round(fact, 2), round(rec, 2),
             round(fact - rec, 2), veredicto,
             json.dumps(lineas, ensure_ascii=False, default=str), informe_xlsx))
        return cur.lastrowid


def listar_revisiones(limite=500):
    with _con() as con:
        return [dict(r) for r in con.execute(
            "SELECT id, fecha, usuario, cups, comercializadora, factura, emision, periodo, "
            "clausula, lineas, facturado, recalculado, diferencia, veredicto FROM revisiones "
            "ORDER BY id DESC LIMIT ?", (limite,))]


def informe_revision(id_revision):
    with _con() as con:
        r = con.execute("SELECT informe, cups, factura FROM revisiones WHERE id=?",
                        (id_revision,)).fetchone()
        return (r["informe"], r["cups"], r["factura"]) if r else (None, None, None)


# ------------------------------------------------------------------- solicitudes de ofertas
def guardar_solicitud(usuario, resumen, datos, texto, excel, estado="Preparada"):
    """resumen: ref, cliente, cups (lista), consumo, recepcion_desde, recepcion_hasta.
    datos: lo necesario para volver a cargar el formulario (JSON). Devuelve el id."""
    with _con(escribir=True) as con:
        cur = con.execute(
            "INSERT INTO solicitudes_ofertas (fecha, usuario, ref, cliente, cups, suministros, "
            "consumo, recepcion_desde, recepcion_hasta, estado, datos, texto, excel) "
            "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)",
            (_ahora(), usuario, resumen.get("ref"), resumen.get("cliente"),
             ", ".join(resumen.get("cups") or []), len(resumen.get("cups") or []),
             resumen.get("consumo"), resumen.get("recepcion_desde"),
             resumen.get("recepcion_hasta"), estado,
             json.dumps(datos, ensure_ascii=False, default=str), texto, excel))
        return cur.lastrowid


def listar_solicitudes(limite=1000):
    """Sin datos, Excel ni texto. producto: 'gas' o 'electricidad' (None en las guardadas
    antes de haber solicitudes de gas, que son de electricidad)."""
    with _con() as con:
        return [dict(r) for r in con.execute(
            "SELECT id, fecha, usuario, ref, cliente, cups, suministros, consumo, "
            "recepcion_desde, recepcion_hasta, estado, comercializadoras, actualizado, "
            "actualizado_por, json_extract(datos, '$.cabecera.producto_tipo') AS producto "
            "FROM solicitudes_ofertas ORDER BY id DESC LIMIT ?", (limite,))]


def solicitud(id_solicitud):
    """Solicitud completa (con datos ya leidos del JSON, Excel y texto) o None."""
    with _con() as con:
        r = con.execute("SELECT * FROM solicitudes_ofertas WHERE id=?",
                        (id_solicitud,)).fetchone()
    if not r:
        return None
    r = dict(r)
    r["datos"] = json.loads(r["datos"])
    return r


def actualizar_seguimiento(id_solicitud, estado, comercializadoras, usuario):
    with _con(escribir=True) as con:
        con.execute("UPDATE solicitudes_ofertas SET estado=?, comercializadoras=?, "
                    "actualizado=?, actualizado_por=? WHERE id=?",
                    (estado, (comercializadoras or "").strip(), _ahora(), usuario,
                     id_solicitud))


# ------------------------------------------------------------------- estudios de potencia
def guardar_estudio_potencia(usuario, resumen, datos, curvas, pdf):
    """resumen: tipo, titular, cups (lista), tarifa, ahorro, inversion. datos: lo necesario para
    volver a cargar el formulario (JSON); curvas: los suministros empaquetados
    (estudio_potencia.servicio.empaquetar). Devuelve el id."""
    with _con(escribir=True) as con:
        cur = con.execute(
            "INSERT INTO estudios_potencia (fecha, usuario, tipo, titular, cups, suministros, "
            "tarifa, ahorro, inversion, datos, curvas, pdf) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)",
            (_ahora(), usuario, resumen.get("tipo"), resumen.get("titular"),
             ", ".join(resumen.get("cups") or []), len(resumen.get("cups") or []),
             resumen.get("tarifa"), resumen.get("ahorro"), resumen.get("inversion"),
             json.dumps(datos, ensure_ascii=False, default=str), curvas, pdf))
        return cur.lastrowid


def listar_estudios_potencia(limite=1000):
    """Sin datos, curvas ni PDF."""
    with _con() as con:
        return [dict(r) for r in con.execute(
            "SELECT id, fecha, usuario, tipo, titular, cups, suministros, tarifa, ahorro, "
            "inversion FROM estudios_potencia ORDER BY id DESC LIMIT ?", (limite,))]


def estudio_potencia(id_estudio):
    """Estudio completo (datos ya leidos del JSON, curvas y PDF) o None."""
    with _con() as con:
        r = con.execute("SELECT * FROM estudios_potencia WHERE id=?", (id_estudio,)).fetchone()
    if not r:
        return None
    r = dict(r)
    r["datos"] = json.loads(r["datos"])
    return r


# ------------------------------------------------------------------- actividad
def registrar_actividad(usuario, modulo, accion, detalle=None):
    """Anota un evento (acceso, pagina, consulta, descarga). Nunca rompe la pagina: si la base
    de datos falla, el evento se pierde sin mas."""
    try:
        with _con(escribir=True) as con:
            con.execute("INSERT INTO actividad (fecha, usuario, modulo, accion, detalle) "
                        "VALUES (?,?,?,?,?)",
                        (_ahora(), usuario, modulo, accion,
                         None if detalle is None else str(detalle)[:300]))
    except Exception:                                     # noqa: BLE001
        pass


def listar_actividad(desde=None, hasta=None, usuario=None, limite=20000):
    """Eventos entre dos fechas AAAA-MM-DD (incluidas), del mas reciente al mas antiguo."""
    sql = ("SELECT id, fecha, usuario, modulo, accion, detalle FROM actividad "
           "WHERE substr(fecha, 1, 10) BETWEEN ? AND ?")
    args = [desde or "0000", hasta or "9999"]
    if usuario:
        sql += " AND usuario = ?"
        args.append(usuario)
    with _con() as con:
        return [dict(r) for r in con.execute(sql + " ORDER BY id DESC LIMIT ?",
                                             args + [limite])]


# ------------------------------------------------------------------- copias
def copia_seguridad(carpeta_destino, conservar=30):
    """Copia consistente de la base de datos (aunque haya gente usandola) y borra las mas
    antiguas, dejando las `conservar` ultimas. Devuelve la ruta de la copia."""
    from pathlib import Path
    destino = Path(carpeta_destino)
    destino.mkdir(parents=True, exist_ok=True)
    ruta = destino / ("revisor_ssaa_%s.db" % dt.datetime.now().strftime("%Y%m%d_%H%M%S"))
    origen = conectar()
    copia = sqlite3.connect(ruta)
    try:
        origen.backup(copia)
    finally:
        copia.close()
        origen.close()
    for vieja in sorted(destino.glob("revisor_ssaa_*.db"))[:-conservar]:
        vieja.unlink()
    return ruta


if __name__ == "__main__":
    import argparse
    p = argparse.ArgumentParser(description="Utilidades de la base de datos del revisor de SSAA")
    p.add_argument("--copia", metavar="CARPETA", help="hace una copia de seguridad en CARPETA")
    p.add_argument("--conservar", type=int, default=30, help="copias que se conservan (30)")
    a = p.parse_args()
    if a.copia:
        print("Copia guardada en %s" % copia_seguridad(a.copia, a.conservar))
    else:
        p.print_help()
