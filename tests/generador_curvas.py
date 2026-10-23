"""
Genera una curva sintética (datos ficticios, sin información de clientes) y la
escribe en distintos formatos habituales de distribuidoras / comercializadoras.
"""

import numpy as np
import pandas as pd

TZ = "Europe/Madrid"
CUPS = "ES0000000000000000XX0F"


def curva_fisica(desde="2025-06-01", hasta="2026-06-01", semilla=1):
    """Cuartos de hora físicos (UTC) con potencia media en kW."""
    ini = pd.Timestamp(desde, tz=TZ).tz_convert("UTC")
    fin = pd.Timestamp(hasta, tz=TZ).tz_convert("UTC")
    t = pd.date_range(ini, fin, freq="15min", inclusive="left")
    loc = t.tz_convert(TZ)
    rng = np.random.default_rng(semilla)
    horas = loc.hour + loc.minute / 60
    base = 120 + 80 * np.clip(np.sin((horas - 7) / 14 * np.pi), 0, None)
    laborable = loc.weekday < 5
    kw = base * np.where(laborable, 1.0, 0.55) + rng.normal(0, 12, len(t))
    kw += np.where(rng.random(len(t)) < 0.002, rng.uniform(80, 200, len(t)), 0)   # picos
    kw = np.round(np.clip(kw, 5, None), 2)
    return pd.DataFrame({"utc": t, "kw": kw})


def esperado(fis, horaria=False):
    """Curva normalizada esperada (96 cuartos/día, criterio de cambio de hora de GE&PE)."""
    df = fis.copy()
    if horaria:
        df["hora"] = df["utc"].dt.floor("h")
        df = df.groupby("hora", as_index=False)["kw"].mean().rename(columns={"hora": "utc"})
        paso = 60
    else:
        paso = 15
    df["inicio"] = df["utc"].dt.tz_convert(TZ).dt.tz_localize(None)
    s = df.groupby("inicio")["kw"].mean()
    rejilla = pd.date_range(s.index.min().normalize(), s.index.max().normalize() + pd.Timedelta(days=1),
                            freq=f"{paso}min", inclusive="left")
    s = s.reindex(rejilla)
    n = 60 // paso
    s = s.where(s.notna(), (s.shift(n) + s.shift(-n)) / 2)
    if horaria:
        s = pd.Series(np.repeat(s.to_numpy(), 4),
                      index=pd.date_range(s.index[0], periods=len(s) * 4, freq="15min"))
    return s


def _etiquetas_fin(fis, paso=15):
    fin = fis["utc"] + pd.Timedelta(minutes=paso)
    loc = fin.dt.tz_convert(TZ)
    return loc


def escribir_p5d(fis, ruta):
    """REE P5D: sin cabecera, 'CUPS;AAAA/MM/DD HH:MM;bandera;AE;AS', etiqueta fin, kWh."""
    loc = _etiquetas_fin(fis)
    verano = np.array([bool(x.dst()) for x in loc]).astype(int)
    lineas = [f"{CUPS};{l:%Y/%m/%d %H:%M};{v};{k / 4:.4f};0"
              for l, v, k in zip(loc, verano, fis["kw"])]
    ruta.write_text("\n".join(lineas), encoding="utf-8")


def escribir_datadis_cuartos(fis, ruta):
    """Datadis: cabecera, fecha dd/mm/aaaa, hora HH:MM de fin (00:15..24:00), kWh con coma decimal."""
    loc = _etiquetas_fin(fis)
    filas = ["CUPS;Fecha;Hora;Consumo_kWh;Metodo_obtencion"]
    for l, k in zip(loc, fis["kw"]):
        if l.hour == 0 and l.minute == 0:
            fecha, hora = (l.tz_localize(None) - pd.Timedelta(days=1)).strftime("%d/%m/%Y"), "24:00"
        else:
            fecha, hora = l.strftime("%d/%m/%Y"), l.strftime("%H:%M")
        filas.append(f"{CUPS};{fecha};{hora};{str(round(k / 4, 4)).replace('.', ',')};Real")
    ruta.write_text("\n".join(filas), encoding="cp1252")


def escribir_distribuidora_horaria(fis, ruta):
    """Horaria con hora entera ordinal 1..23/24/25 por día, energía en Wh."""
    df = fis.copy()
    df["hora_utc"] = df["utc"].dt.floor("h")
    h = df.groupby("hora_utc")["kw"].mean().reset_index()
    loc = h["hora_utc"].dt.tz_convert(TZ)
    h["dia"] = loc.dt.strftime("%Y-%m-%d")
    h["n"] = h.groupby("dia").cumcount() + 1
    filas = ["Fecha,Hora,Energia activa (Wh)"]
    filas += [f"{d},{n},{round(k * 1000)}" for d, n, k in zip(h["dia"], h["n"], h["kw"])]
    ruta.write_text("\n".join(filas), encoding="utf-8")


def escribir_excel_geype(esp, ruta):
    """Como la hoja Curva_potencia de GE&PE: fecha, hora 0-23, cuarto 1-4, kW (ya normalizada)."""
    idx = esp.index
    df = pd.DataFrame({"Dia": idx.normalize(), "Horas": idx.hour, "Cuarto": idx.minute // 15 + 1,
                       "P. Demandada (kW)": esp.to_numpy()})
    with pd.ExcelWriter(ruta) as w:
        pd.DataFrame({"x": ["Informe de prueba"]}).to_excel(w, sheet_name="Portada", index=False)
        df.to_excel(w, sheet_name="Curva_potencia", index=False, startrow=3)


def escribir_iso_offset(fis, ruta):
    """ISO 8601 con desfase horario, etiqueta fin, potencia en kW, separador coma."""
    loc = _etiquetas_fin(fis)
    filas = ["timestamp,potencia_kw"]
    filas += [f"{l.isoformat()},{k}" for l, k in zip(loc, fis["kw"])]
    ruta.write_text("\n".join(filas), encoding="utf-8")


def escribir_ancho_horario(fis, ruta):
    """Una fila por día y columnas H1..H25 con kWh (formato 'ancho')."""
    df = fis.copy()
    df["hora_utc"] = df["utc"].dt.floor("h")
    h = df.groupby("hora_utc")["kw"].mean().reset_index()
    loc = h["hora_utc"].dt.tz_convert(TZ)
    h["dia"] = loc.dt.tz_localize(None).dt.normalize()
    h["n"] = h.groupby("dia").cumcount() + 1
    ancho = h.pivot(index="dia", columns="n", values="kw")
    ancho.columns = [f"H{c}" for c in ancho.columns]
    ancho = ancho.reset_index().rename(columns={"dia": "Fecha"})
    ancho.to_excel(ruta, index=False)


def escribir_p5d_multipunto(curvas, ruta):
    """Varios CUPS en un único fichero P5D: curvas = {cups: curva_fisica}."""
    lineas = []
    for cups, fis in curvas.items():
        loc = _etiquetas_fin(fis)
        verano = np.array([bool(x.dst()) for x in loc]).astype(int)
        lineas += [f"{cups};{l:%Y/%m/%d %H:%M};{v};{k / 4:.4f};0" for l, v, k in zip(loc, verano, fis["kw"])]
    ruta.write_text("\n".join(lineas), encoding="utf-8")
