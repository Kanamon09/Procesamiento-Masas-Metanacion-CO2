import io
import json
import re
import unicodedata
from datetime import datetime
from difflib import SequenceMatcher
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from matplotlib.lines import Line2D
from matplotlib.ticker import AutoMinorLocator, MultipleLocator
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

RUTA = "."

PREGUNTAR_INICIO = True
DESFASE_S = 0
UMBRAL_DETECCION = 5.0
UMBRAL_INICIO = 0.3
MIN_ANTES = 0
T_FIN = None
TRAMO_T = "calentamiento"

ISOTERMA = True
ISO_BANDA = 15.0
ISO_CAIDA = 60.0
ISO_MIN_DUR = 5.0
ISO_ANCHO = 0.32

MASAS_CONSOLA = [15]
VENTANA_BASE_MIN = 15
N_SIGMA_ONSET = 5
MIN_SOSTENIDO = 3

PALABRAS_BLANCO = ("blanco", "blank", "sincat", "sin_cat", "sin cat", "bypass", "vacio")
BLANCO_VENTANA = None
MESETA_MIN = 5.0
MESETA_DERIVA_MAX = 2.0
MESETA_TOL_MAX = 1.5
MASA_CO2 = 44
MASA_CH4 = 15
MASA_H2 = 2
MASA_INERTE = 40
USAR_INERTE = "auto"
BLANCO_POR_ARCHIVO = {}
F = "datos"
TABLA_FRAG = {
    "CO2": {"masa": 44, "FF": 0.86, "S": 1.42},
    "CH4": {"masa": 15, "FF": 0.39, "S": 1.40},
}
VENTANA_SIN_REACCION_T = (120, 220)
ESTABLE_MIN = 10.0
DERIVA_ESTABLE_MAX = 3.0
VENTANA_PICO_MIN = 2.0
REPORTE_T = "Tmax"
MOSTRAR_ESTABLE = False
BARRAS_EN = "Tmax"
R_H2_CO2 = 4.0
ERROR_METODO = "bloques"
REF_X_VALIDA = 5.0
REF_H2_VALIDA = 10.0
F_ATIPICO = 0.10
X_MIN_CONV = 0.05

ETIQUETAS_RESUMEN = {
    "OMNi1.6": "OM1", "OMNi1.6Fe0.3": "OM2", "OMNi1.6Cu0.3": "OM3", "OMNi1.6Cu0.3Fe0.3": "OM4",
    "OMNi1.6Cu0.15Fe0.15": "OM5", "OMNi1.6Cu0.15Fe0.3": "OM6", "OMNi1.6Cu0.3Fe0.15": "OM7",
}

PREGUNTAR_ELIMINAR = True
MOSTRAR_VALOR_FINAL = True

VISTAS = [
    {"sufijo": "",            "escala": "log",    "ylim": None},
    {"sufijo": "_lineal",     "escala": "linear", "ylim": None},
    {"sufijo": None,          "escala": "linear", "ylim": (0, 6)},
]

FIG_W, FIG_H = 5.616, 3.1
MARGEN_IZQ_IN, MARGEN_DER_IN = 0.65, 0.91
MARGEN_IZQ, MARGEN_INF, MARGEN_SUP = MARGEN_IZQ_IN / FIG_W, 0.14, 0.03
MARGEN_DER = MARGEN_DER_IN / FIG_W
ANCHO_EXTRA_VALORES = 0.45

FONT_FAMILY = ["Arial", "Liberation Sans", "DejaVu Sans"]
FS_LABEL = 8.5
FS_TICK = 7.5
FS_CURVA = 6.5
LW_CURVE = 0.9
LW_AXES = 0.6
TICK_MAJOR, TICK_MINOR = 2.0, 1.0
DPI = 300

XLABEL_T_MIN = "Tiempo (min)"
XLABEL_TEMP = "Temperatura (°C)"
XLABEL_ISO = "Isoterma (min)"
XLABEL_ISO_T = "T (°C)"
YLABEL = "Presión (mbar)"
YLABEL_TEMP = "Temperatura del reactor (°C)"
SUAVIZADO = 15
SUAVIZADO_T_S = 5
CARPETA_SALIDA = "Resultados_RGA"
EXCEL_TODAS_LAS_MASAS = False
COLOR_T = "#c0392b"

MASAS = {
    2:  (r"H$_2$",       "#e6850e"),
    3:  (r"HD$^+$",      "#999999"),
    12: ("C",            "#808000"),
    14: ("N",            "#7a7a7a"),
    15: (r"CH$_3$",      "#d95fa8"),
    16: (r"CH$_4$ / O",  "#2a9d5c"),
    17: ("OH",           "#5ba3d0"),
    18: (r"H$_2$O",      "#004f84"),
    28: (r"CO · N$_2$",  "#1a1a1a"),
    32: (r"O$_2$",       "#d15e57"),
    36: (r"$^{36}$Ar",   "#a98bd0"),
    40: ("Ar",           "#5b2c83"),
    44: (r"CO$_2$",      "#8c510a"),
}

plt.rcParams.update({
    "svg.fonttype": "none",
    "pdf.fonttype": 42,
    "font.family": "sans-serif",
    "font.sans-serif": FONT_FAMILY,
    "mathtext.fontset": "custom",
    "mathtext.default": "regular",
    "mathtext.cal": "sans",
    "font.size": 11,
    "axes.labelsize": FS_LABEL,
    "xtick.labelsize": FS_TICK,
    "ytick.labelsize": FS_TICK,
    "axes.linewidth": LW_AXES,
    "xtick.major.width": LW_AXES, "ytick.major.width": LW_AXES,
    "xtick.minor.width": LW_AXES * 0.7, "ytick.minor.width": LW_AXES * 0.7,
    "xtick.major.size": TICK_MAJOR, "ytick.major.size": TICK_MAJOR,
    "xtick.minor.size": TICK_MINOR, "ytick.minor.size": TICK_MINOR,
    "xtick.direction": "in", "ytick.direction": "in",
    "xtick.top": False, "ytick.right": False,
    "lines.linewidth": LW_CURVE,
    "figure.dpi": DPI,
    "axes.labelweight": "bold",
})


def leer_lineas(ruta: Path):
    for enc in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            return ruta.read_text(encoding=enc).splitlines()
        except UnicodeDecodeError:
            continue
    raise ValueError("No se pudo decodificar el archivo")


def es_txt_masas(ruta: Path) -> bool:
    if ruta.suffix.lower() != ".txt":
        return False
    try:
        with open(ruta, "r", encoding="latin-1") as f:
            cabeza = f.read(4000)
    except OSError:
        return False
    return "[Scan Data" in cabeza or "Spectra International" in cabeza


def leer_scan(ruta_txt: Path) -> pd.DataFrame:
    lineas = leer_lineas(ruta_txt)
    ini = next(i for i, l in enumerate(lineas) if l.lstrip().startswith('"Time"'))
    bloque = [lineas[ini]]
    for l in lineas[ini + 1:]:
        if not l.strip() or l.lstrip().startswith('"['):
            break
        bloque.append(l)
    df = pd.read_csv(io.StringIO("\n".join(bloque)), sep="\t", quotechar='"', index_col=False)
    df = df.loc[:, ~df.columns.str.startswith("Unnamed")]
    try:
        df["Time"] = pd.to_datetime(df["Time"], format="%d/%m/%Y %H:%M:%S.%f")
    except ValueError:
        df["Time"] = pd.to_datetime(df["Time"], dayfirst=True)
    return df


def normalizar(texto: str) -> str:
    texto = unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode().lower()
    return texto


def es_reactor(ruta: Path) -> bool:
    return (ruta.suffix.lower() in (".xls", ".xlsx", ".csv", ".txt")
            and "prueba" in normalizar(ruta.stem) and not es_txt_masas(ruta))


def leer_reactor(ruta: Path) -> pd.DataFrame:
    r = None
    for enc in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            r = pd.read_csv(ruta, sep="\t", decimal=",", encoding=enc)
            if "Time" in r.columns and "Value" in r.columns:
                break
            r = None
        except (UnicodeDecodeError, pd.errors.ParserError):
            r = None
    if r is None:
        r = pd.read_excel(ruta)
    r.columns = [str(c).strip() for c in r.columns]
    if "T1/T2" in r.columns and (r["T1/T2"].astype(str).str.strip() == "T1").any():
        r = r[r["T1/T2"].astype(str).str.strip() == "T1"]
    r["Value"] = pd.to_numeric(r["Value"].astype(str).str.replace(",", "."), errors="coerce")
    tiempo = r["Time"].astype(str).str.strip()
    try:
        r["Time"] = pd.to_datetime(tiempo, format="%Y/%m/%d_%H:%M:%S")
    except ValueError:
        r["Time"] = pd.to_datetime(tiempo.str.replace("_", " "), errors="coerce")
    r = r.dropna(subset=["Time", "Value"])
    r = r.groupby("Time", as_index=False)["Value"].mean().rename(columns={"Value": "T"})
    if len(r) < 30:
        raise ValueError(f"El archivo del reactor {ruta.name} casi no tiene datos")
    return r.sort_values("Time").reset_index(drop=True)


def emparejar(txt: Path, df: pd.DataFrame, reactores: dict):
    t0, t1 = df["Time"].iloc[0], df["Time"].iloc[-1]
    mejor, puntaje = None, (-1.0, -1.0)
    for ruta, r in reactores.items():
        solape = (min(t1, r["Time"].iloc[-1]) - max(t0, r["Time"].iloc[0])).total_seconds()
        parecido = SequenceMatcher(None, re.sub(r"[^a-z0-9.]", "", normalizar(txt.stem)),
                                   re.sub(r"[^a-z0-9.]", "", normalizar(ruta.stem).replace("prueba", ""))).ratio()
        p = (max(solape, 0.0), parecido)
        if p > puntaje:
            mejor, puntaje = ruta, p
    if mejor is None or puntaje[0] <= 0:
        return None
    return mejor


def detectar_inicio(r: pd.DataFrame):
    T = r["T"].rolling(15, center=True, min_periods=1).median().to_numpy()
    seg = (r["Time"] - r["Time"].iloc[0]).dt.total_seconds().to_numpy()
    base0 = np.median(T[seg <= 120])
    arriba = np.where(T > base0 + UMBRAL_DETECCION)[0]
    if not len(arriba):
        raise ValueError(f"La temperatura del reactor nunca sube más de {UMBRAL_DETECCION} °C sobre la inicial")
    i1 = arriba[0]
    rel = seg - seg[i1]
    sel = (rel >= -420) & (rel <= -120)
    base = np.median(T[sel]) if sel.sum() >= 10 else base0
    ruido = np.std(r["T"].to_numpy()[sel] - T[sel]) if sel.sum() >= 10 else 0.0
    umbral = max(UMBRAL_INICIO, 4 * ruido)
    j = i1
    while j > 0 and T[j] > base + umbral:
        j -= 1
    return j, base


def resolver_hora(valor, referencia: pd.Timestamp):
    if valor is None:
        return None
    if isinstance(valor, (int, float)):
        return referencia + pd.Timedelta(minutes=float(valor))
    texto = str(valor).strip()
    for fmt in ("%H:%M:%S", "%H:%M"):
        try:
            hora = datetime.strptime(texto, fmt).time()
            t = pd.Timestamp(datetime.combine(referencia.date(), hora))
            if t < referencia - pd.Timedelta(hours=12):
                t += pd.Timedelta(days=1)
            return t
        except ValueError:
            continue
    raise ValueError(f"No entiendo el tiempo '{valor}'. Usa 'HH:MM:SS' o minutos (número).")


def preguntar(texto: str) -> str:
    try:
        return input(texto).strip()
    except EOFError:
        return ""


def elegir_inicio(r: pd.DataFrame):
    j, base = detectar_inicio(r)
    t_det = r["Time"].iloc[j]
    print(f"    Reactor: {r['Time'].iloc[0]:%H:%M:%S} a {r['Time'].iloc[-1]:%H:%M:%S}  "
          f"(T {r['T'].min():.1f} a {r['T'].max():.1f} °C)")
    print(f"    Inicio de calentamiento detectado: {t_det:%H:%M:%S}  (T = {r['T'].iloc[j]:.1f} °C, "
          f"línea base {base:.1f} °C)")
    if not PREGUNTAR_INICIO:
        return t_det, base, False
    while True:
        resp = preguntar(f"    t = 0 [Enter = {t_det:%H:%M:%S}] (o escribe HH:MM:SS): ")
        if resp == "":
            return t_det, base, False
        try:
            t = resolver_hora(resp.replace(",", "."), r["Time"].iloc[0])
            if not (r["Time"].iloc[0] <= t <= r["Time"].iloc[-1]):
                raise ValueError("esa hora está fuera del archivo del reactor")
            return t, base, True
        except ValueError as e:
            print(f"    Aviso: {e}")


def sincronizar(df: pd.DataFrame, r: pd.DataFrame, t_inicio: pd.Timestamp):
    desfase = pd.Timedelta(seconds=DESFASE_S)
    t_cero = t_inicio + desfase
    t_min = (df["Time"] - t_cero).dt.total_seconds().to_numpy() / 60.0

    T_suave = r["T"].rolling(max(1, int(SUAVIZADO_T_S)), center=True, min_periods=1).mean().to_numpy()
    x_r = (r["Time"] - t_inicio).dt.total_seconds().to_numpy()
    x_m = (df["Time"] - desfase - t_inicio).dt.total_seconds().to_numpy()
    T_rga = np.interp(x_m, x_r, T_suave, left=np.nan, right=np.nan)

    mascara = (t_min >= -MIN_ANTES) & ~np.isnan(T_rga)
    t_fin = resolver_hora(T_FIN, t_cero)
    if t_fin is not None:
        mascara &= (df["Time"] <= t_fin).to_numpy()
    if mascara.sum() < 10:
        raise ValueError(f"El masas casi no tiene datos después del inicio del calentamiento "
                         f"(masas: {df['Time'].iloc[0]:%H:%M:%S}-{df['Time'].iloc[-1]:%H:%M:%S}, "
                         f"t = 0 en {t_cero:%H:%M:%S}).")

    idx = np.where(mascara)[0]
    i_max = idx[np.nanargmax(T_rga[idx])]
    n = np.arange(len(df))
    tramo = np.full(len(df), "", dtype=object)
    tramo[mascara & (t_min < 0)] = "Antes del inicio"
    iso = detectar_isoterma(r, t_inicio)
    mascara_iso = None
    if iso is not None:
        en_iso = mascara & (t_min >= iso["t_ini"]) & (t_min <= iso["t_fin"])
        if en_iso.sum() < 5:
            iso = None
        else:
            tramo[mascara & (t_min >= 0) & (t_min < iso["t_ini"])] = "Rampa"
            tramo[en_iso] = "Isoterma"
            tramo[mascara & (t_min > iso["t_fin"])] = "Después de la isoterma"
            mascara_T = mascara & (t_min >= 0) & (t_min < iso["t_ini"])
            mascara_iso = en_iso
    if iso is None:
        tramo[mascara & (t_min >= 0) & (n <= i_max)] = "Rampa"
        tramo[mascara & (n > i_max)] = "Después de Tmax"
        mascara_T = mascara & (t_min >= 0)
        if TRAMO_T == "calentamiento":
            mascara_T &= n <= i_max

    info = {"t_cero": t_cero, "i_max": i_max, "tramo": tramo, "iso": iso, "mascara_iso": mascara_iso}
    return t_min, T_rga, mascara, mascara_T, info


def detectar_isoterma(r: pd.DataFrame, t_inicio: pd.Timestamp):
    if not ISOTERMA:
        return None
    t = (r["Time"] - t_inicio).dt.total_seconds().to_numpy() / 60.0
    T = r["T"].rolling(31, center=True, min_periods=1).mean().to_numpy()
    validos = np.where(t >= 0)[0]
    if len(validos) < 30:
        return None
    Tmax = T[validos].max()
    i_s = validos[np.argmax(T[validos] >= Tmax - ISO_BANDA)]
    caida = np.where((np.arange(len(T)) > i_s) & (T < Tmax - ISO_CAIDA))[0]
    i_e = caida[0] - 1 if len(caida) else len(T) - 1
    if t[i_e] - t[i_s] < ISO_MIN_DUR:
        return None
    s = slice(i_s, i_e + 1)
    Tr = r["T"].to_numpy()[s]
    return {"t_ini": float(t[i_s]), "t_fin": float(t[i_e]), "T_ini": float(T[i_s]), "T_fin": float(T[i_e]),
            "T_media": float(Tr.mean()), "T_min": float(Tr.min()), "T_max": float(Tr.max()),
            "deriva": float(np.polyfit(t[s], T[s], 1)[0]), "inicio": r["Time"].iloc[i_s]}


def calcular_rampa(r: pd.DataFrame, t_inicio: pd.Timestamp):
    t = (r["Time"] - t_inicio).dt.total_seconds().to_numpy() / 60.0
    T = r["T"].to_numpy()
    i0 = int(np.argmin(np.abs(t)))
    imax = int(np.argmax(T))
    T_s = pd.Series(T).rolling(61, center=True, min_periods=1).mean().to_numpy()
    dTdt = np.gradient(T_s, t)
    calent = (np.arange(len(T)) >= i0) & (np.arange(len(T)) <= imax)
    tramos = []
    a = 50.0 * np.ceil(T[i0] / 50.0)
    while a < T[imax]:
        b = min(a + 100.0, T[imax])
        s = calent & (T >= a) & (T <= b)
        if s.sum() > 20 and b - a >= 10:
            p = np.polyfit(t[s], T[s], 1)
            tramos.append((a, b, float(p[0]), float(t[s][0]), float(t[s][-1])))
        a += 100.0
    k = int(np.argmax(np.where(calent, dTdt, -np.inf)))
    return {
        "T0": float(T[i0]), "Tmax": float(T[imax]), "t_Tmax": float(t[imax]),
        "promedio": float((T[imax] - T[i0]) / t[imax]) if t[imax] > 0 else np.nan,
        "max_inst": float(dTdt[k]), "T_max_inst": float(T[k]), "t_max_inst": float(t[k]),
        "tramos": tramos, "dTdt": dTdt,
    }


def analizar_masa(t, T, crudo, suave, tramo):
    res = dict.fromkeys(["base", "sigma", "t_on", "T_on", "P_on", "t_tan", "T_tan",
                         "t_50", "T_50", "t_max", "T_max", "P_max"])
    calent = (tramo == "Rampa") | (tramo == "Isoterma")
    b = (t >= 0) & (t <= VENTANA_BASE_MIN)
    if b.sum() < 10 or not calent.any():
        return res
    base, sigma = float(np.median(crudo[b])), float(np.std(suave[b]))
    res.update(base=base, sigma=sigma)

    ic = np.where(calent)[0]
    imax = ic[int(np.nanargmax(suave[ic]))]
    res.update(t_max=float(t[imax]), T_max=_f(T[imax]), P_max=float(suave[imax]))

    umbral = base + N_SIGMA_ONSET * sigma
    for i in ic:
        if t[i] <= VENTANA_BASE_MIN or i > imax:
            continue
        if suave[i] > umbral:
            w = (t >= t[i]) & (t <= t[i] + MIN_SOSTENIDO)
            if (suave[w] > umbral).all():
                res.update(t_on=float(t[i]), T_on=_f(T[i]), P_on=float(suave[i]))
                i_on = i
                break
    if res["t_on"] is None:
        return res

    subida = suave[imax] - base
    i50 = next(i for i in ic if i >= i_on and suave[i] >= base + 0.5 * subida)
    res.update(t_50=float(t[i50]), T_50=_f(T[i50]))
    s = calent & (np.arange(len(t)) <= imax) & (suave >= base + 0.1 * subida) & (suave <= base + 0.5 * subida)
    if s.sum() >= 5:
        pT = np.polyfit(T[s], suave[s], 1)
        pt = np.polyfit(t[s], suave[s], 1)
        t_tan = (base - pt[1]) / pt[0] if pt[0] > 0 else -1
        if pT[0] > 0 and t_tan >= VENTANA_BASE_MIN:
            res.update(T_tan=float((base - pT[1]) / pT[0]), t_tan=float(t_tan))
    return res


def repartir_etiquetas(posiciones, separacion=0.052, minimo=0.02, maximo=0.98):
    n = len(posiciones)
    grupos = [[i] for i in sorted(range(n), key=lambda i: posiciones[i])]

    def centro(g):
        return sum(posiciones[i] for i in g) / len(g)

    cambio = True
    while cambio:
        cambio = False
        for k in range(len(grupos) - 1):
            a, b = grupos[k], grupos[k + 1]
            tope_a = centro(a) + separacion * (len(a) - 1) / 2
            base_b = centro(b) - separacion * (len(b) - 1) / 2
            if base_b - tope_a < separacion:
                grupos[k:k + 2] = [a + b]
                cambio = True
                break

    nuevas = [0.0] * n
    for g in grupos:
        c = centro(g)
        for j, i in enumerate(g):
            nuevas[i] = c + separacion * (j - (len(g) - 1) / 2)

    orden = [i for g in grupos for i in g]
    nuevas[orden[0]] = max(nuevas[orden[0]], minimo)
    for a, b in zip(orden, orden[1:]):
        nuevas[b] = max(nuevas[b], nuevas[a] + separacion)
    nuevas[orden[-1]] = min(nuevas[orden[-1]], maximo)
    for a, b in zip(reversed(orden[1:]), reversed(orden[:-1])):
        nuevas[b] = min(nuevas[b], nuevas[a] - separacion)
    return nuevas


def nombre_plano(etiqueta: str) -> str:
    for a, b in (("$", ""), ("_", ""), ("^", ""), ("{", ""), ("}", ""), (" · ", "-")):
        etiqueta = etiqueta.replace(a, b)
    return etiqueta


def suavizar(df, masa):
    return df[f"Mass {masa}"].astype(float).rolling(SUAVIZADO, center=True, min_periods=1).mean()


def paso_eje_x(span, eje):
    if eje == "temp":
        return 25 if span <= 150 else 50 if span <= 350 else 100
    return 1 if span <= 8 else 2 if span <= 20 else 5 if span <= 40 else 10 if span <= 100 else 20


def _estilo_ejes(ax, escala):
    ax.tick_params(axis="both", labelsize=FS_TICK, length=TICK_MAJOR, width=LW_AXES, direction="in", pad=3)
    ax.tick_params(axis="both", which="minor", length=TICK_MINOR, width=LW_AXES * 0.7, direction="in")
    ax.xaxis.set_minor_locator(AutoMinorLocator(5))
    if escala != "log":
        ax.yaxis.set_minor_locator(AutoMinorLocator(2))
    for sp in ax.spines.values():
        sp.set_linewidth(LW_AXES)


def graficar(df, x_total, mascara, salida_base: Path, escala="log", ylim=None, excluir=(), eje="tiempo", iso=None):
    paneles = [(x_total, mascara, eje)]
    if iso is not None:
        paneles.append((iso["x"], iso["mask"], "iso"))
    union = mascara.copy()
    for _, m_, _ in paneles[1:]:
        union |= m_

    extra = ANCHO_EXTRA_VALORES if (escala == "log" and MOSTRAR_VALOR_FINAL) else 0.0
    fig_w = FIG_W + extra
    fig = plt.figure(figsize=(fig_w, FIG_H))
    izq = MARGEN_IZQ * FIG_W / fig_w
    der = 1 - (MARGEN_DER * FIG_W + extra) / fig_w
    sup = 1 - (0.30 / FIG_H if iso is not None else MARGEN_SUP)
    alto = sup - MARGEN_INF
    if iso is None:
        axs = [fig.add_axes([izq, MARGEN_INF, der - izq, alto])]
    else:
        hueco = 0.12 / fig_w
        w = der - izq - hueco
        w1, w2 = w * (1 - ISO_ANCHO), w * ISO_ANCHO
        ax1 = fig.add_axes([izq, MARGEN_INF, w1, alto])
        axs = [ax1, fig.add_axes([izq + w1 + hueco, MARGEN_INF, w2, alto], sharey=ax1)]

    x_ult, m_ult, _ = paneles[-1]
    pos_ult = np.where(m_ult)[0]
    i_ancla = pos_ult[int(np.nanargmax(x_ult[m_ult]))]
    x_fin_ult = float(x_ult[i_ancla])

    series, omitidas, fuera = [], [], []
    for masa, (etiqueta, color) in MASAS.items():
        col = f"Mass {masa}"
        if col not in df.columns or masa in excluir:
            continue
        crudo_t = df[col].astype(float).to_numpy()
        suave_t = suavizar(df, masa).to_numpy()
        crudo_u, suave_u = pd.Series(crudo_t[union]), pd.Series(suave_t[union])
        if escala == "log" and ((crudo_u <= 0).mean() > 0.25 or suave_u.median() <= 0):
            omitidas.append(masa)
            continue
        if escala == "linear" and ylim and not (suave_u.max() > ylim[0] and suave_u.min() < ylim[1]):
            fuera.append(masa)
            continue
        for ax, (x_p, m_p, _) in zip(axs, paneles):
            y = suave_t[m_p]
            if escala == "log":
                y = np.where(y > 0, y, np.nan)
            ax.plot(x_p[m_p], y, color=color, lw=LW_CURVE, solid_capstyle="round")
        ventana = crudo_t[max(0, i_ancla - SUAVIZADO + 1):i_ancla + 1]
        ultimo = float(np.mean(ventana))
        if escala == "log" and not ultimo > 0:
            ultimo = float(suave_t[i_ancla])
        y_u = suave_u.where(suave_u > 0) if escala == "log" else suave_u
        series.append((masa, etiqueta, color, ultimo, y_u))

    if not series:
        plt.close(fig)
        return omitidas, fuera, []
    for ax in axs:
        ax.set_yscale(escala)
    todo = np.concatenate([s[4].dropna().to_numpy() for s in series])
    if ylim:
        axs[0].set_ylim(*ylim)
    elif escala == "log":
        axs[0].set_ylim(10 ** np.floor(np.log10(np.percentile(todo, 2))), 10 ** np.ceil(np.log10(todo.max())))
    else:
        axs[0].set_ylim(0, todo.max() * 1.05)

    for ax, (x_p, m_p, tipo) in zip(axs, paneles):
        xs = x_p[m_p]
        x_ini, x_fin = float(np.nanmin(xs)), float(np.nanmax(xs))
        if tipo == "tiempo":
            x_ini = -float(MIN_ANTES)
        if tipo == "iso":
            x_ini = 0.0
        ax.set_xlim(x_ini, x_fin)
        ax.xaxis.set_major_locator(MultipleLocator(paso_eje_x(x_fin - x_ini, "temp" if tipo == "temp" else "tiempo")))
        _estilo_ejes(ax, escala)
        ax.set_xlabel({"temp": XLABEL_TEMP, "tiempo": XLABEL_T_MIN, "iso": XLABEL_ISO}[tipo],
                      fontsize=FS_LABEL, fontweight="bold", labelpad=4)
    axs[0].set_ylabel(YLABEL, fontsize=FS_LABEL, fontweight="bold", labelpad=4)

    if iso is not None:
        ax1, ax2 = axs
        ax1.spines["right"].set_visible(False)
        ax2.spines["left"].set_visible(False)
        ax2.tick_params(axis="y", which="both", left=False, labelleft=False)
        d = 0.012
        kw = dict(color="k", lw=LW_AXES, clip_on=False)
        r1 = (FIG_W * (1 - ISO_ANCHO)) / (FIG_W * ISO_ANCHO)
        for yy in (0, 1):
            ax1.plot((1 - d, 1 + d), (yy - d * 2, yy + d * 2), transform=ax1.transAxes, **kw)
            ax2.plot((-d * r1, d * r1), (yy - d * 2, yy + d * 2), transform=ax2.transAxes, **kw)
        xi, Ti = iso["x"][iso["mask"]], iso["T"][iso["mask"]]
        orden_i = np.argsort(xi)
        arriba = ax2.twiny()
        arriba.set_xlim(ax2.get_xlim())
        lo, hi = ax2.get_xlim()
        marcas = [m for m in ax2.get_xticks() if lo - 1e-9 <= m <= hi + 1e-9]
        if len(marcas) > 5:
            marcas = marcas[::2]
        arriba.set_xticks(marcas)
        arriba.set_xticklabels([f"{np.interp(m, xi[orden_i], Ti[orden_i]):.0f}" for m in marcas])
        arriba.tick_params(axis="x", labelsize=FS_TICK - 0.5, length=TICK_MAJOR, width=LW_AXES, direction="in",
                           pad=2, colors=COLOR_T)
        arriba.set_xlabel(XLABEL_ISO_T, fontsize=FS_TICK, color=COLOR_T, labelpad=3)
        for nombre_sp, sp in arriba.spines.items():
            sp.set_visible(nombre_sp == "top")
            sp.set_linewidth(LW_AXES)

    ax = axs[-1]
    ancho_ejes = ax.get_position().width * fig_w
    dx = lambda pulgadas: pulgadas / ancho_ejes
    if escala == "log":
        fig.canvas.draw()
        inv = ax.transAxes.inverted()
        y_ax = [inv.transform(ax.transData.transform((x_fin_ult, s[3])))[1] for s in series]
        for (masa, etiqueta, color, ultimo, _), y1 in zip(series, repartir_etiquetas(y_ax)):
            ax.annotate(
                "", xy=(x_fin_ult, ultimo), xycoords="data",
                xytext=(1 + dx(0.07), y1), textcoords="axes fraction",
                arrowprops=dict(arrowstyle="-", color=color, lw=0.4, shrinkA=0, shrinkB=0),
                annotation_clip=False,
            )
            texto = f"{masa}  {etiqueta}" + (f" ({ultimo:.3g})" if MOSTRAR_VALOR_FINAL else "")
            ax.text(1 + dx(0.10), y1, texto, transform=ax.transAxes, ha="left", va="center",
                    fontsize=FS_CURVA, color=color, clip_on=False)
    else:
        orden = sorted(series, key=lambda s: -s[3])
        handles = [Line2D([0], [0], color=s[2], lw=LW_CURVE + 0.4) for s in orden]
        textos = [f"{s[0]}  {s[1]}" for s in orden]
        ax.legend(handles, textos, loc="center left", bbox_to_anchor=(1 + dx(0.13), 0.5), frameon=False,
                  fontsize=FS_CURVA, handlelength=1.7, handletextpad=0.5, labelspacing=0.5, borderaxespad=0)

    for ext in ("png", "svg"):
        fig.savefig(Path(f"{salida_base}.{ext}"), dpi=DPI, facecolor="white")
    plt.close(fig)
    return omitidas, fuera, [s[0] for s in series]


def graficar_perfil_T(r, t_inicio, base, info, df, salida_base: Path):
    t = (r["Time"] - t_inicio).dt.total_seconds().to_numpy() / 60.0
    fig, ax = plt.subplots(figsize=(FIG_W, FIG_H))
    fig.subplots_adjust(left=MARGEN_IZQ, right=1 - 0.25 / FIG_W, bottom=MARGEN_INF, top=1 - MARGEN_SUP)
    ax.plot(t, r["T"], color=COLOR_T, lw=LW_CURVE)
    ax.axvline(0, color="#555555", lw=0.5, ls="--")
    ax.axhline(base, color="#999999", lw=0.4, ls=":")
    i_max = int(r["T"].idxmax())
    ax.plot([0], [np.interp(0, t, r["T"])], "o", ms=3, color="k")
    ax.annotate(f"t = 0  ({t_inicio:%H:%M:%S})", xy=(0, base), xytext=(4, 14), textcoords="offset points",
                fontsize=FS_CURVA, ha="left")
    ax.annotate(f"T$_{{máx}}$ = {r['T'].iloc[i_max]:.1f} °C", xy=(t[i_max], r["T"].iloc[i_max]),
                xytext=(0, -12), textcoords="offset points", fontsize=FS_CURVA, ha="center")
    m = info["mascara"]
    tm = info["t_min"][m]
    ax.axvspan(tm.min(), tm.max(), color="#dddddd", alpha=0.35, lw=0)
    iso = info.get("iso")
    if iso is not None:
        ax.axvspan(iso["t_ini"], iso["t_fin"], color="#f5b041", alpha=0.25, lw=0)
        ax.text((iso["t_ini"] + iso["t_fin"]) / 2, iso["T_min"] - 0.10 * (r["T"].max() - base),
                f"Isoterma\n{iso['T_min']:.0f}–{iso['T_max']:.0f} °C", ha="center", va="top", fontsize=FS_CURVA)
    ax.set_xlim(t[0], t[-1])
    ax.set_xlabel(XLABEL_T_MIN, fontsize=FS_LABEL, fontweight="bold", labelpad=4)
    ax.set_ylabel(YLABEL_TEMP, fontsize=FS_LABEL, fontweight="bold", labelpad=4)
    ax.tick_params(axis="both", labelsize=FS_TICK, length=TICK_MAJOR, width=LW_AXES, direction="in", pad=3)
    ax.tick_params(axis="both", which="minor", length=TICK_MINOR, width=LW_AXES * 0.7, direction="in")
    ax.xaxis.set_major_locator(MultipleLocator(paso_eje_x(t[-1] - t[0], "tiempo")))
    ax.xaxis.set_minor_locator(AutoMinorLocator(5))
    ax.yaxis.set_minor_locator(AutoMinorLocator(2))
    for ext in ("png", "svg"):
        fig.savefig(Path(f"{salida_base}.{ext}"), dpi=DPI, facecolor="white")
    plt.close(fig)


def sufijo_vista(v) -> str:
    if v.get("sufijo") is not None:
        return v["sufijo"]
    if v["escala"] == "log":
        return ""
    if v.get("ylim"):
        return f"_lineal_{v['ylim'][0]:g}-{v['ylim'][1]:g}"
    return "_lineal"


def generar_vistas(df, t_min, T_rga, mascara, mascara_T, carpeta, nombre, info, excluir=()):
    marca = ("_sin" + "-".join(str(m) for m in sorted(excluir))) if excluir else ""
    iso = None
    if info.get("iso") is not None:
        iso = {"x": t_min - info["iso"]["t_ini"], "mask": info["mascara_iso"], "T": T_rga}
    archivos, dibujadas = [], []
    for eje, x, mask, sufijo_eje in (("tiempo", t_min, mascara, "_vs_tiempo"), ("temp", T_rga, mascara_T, "_vs_temp")):
        for v in VISTAS:
            base = carpeta / f"{nombre}{sufijo_eje}{marca}{sufijo_vista(v)}"
            omitidas, fuera, usadas = graficar(df, x, mask, base, v["escala"], v.get("ylim"), excluir, eje,
                                               iso if eje == "temp" else None)
            if not usadas:
                print(f"    [--] {base.name}: sin curvas para esta vista")
                continue
            notas = []
            if omitidas:
                notas.append(f"omitidas por ruido: {omitidas}")
            if fuera:
                notas.append(f"fuera de escala: {fuera}")
            print(f"    [OK] {base.name}.png / .svg" + (f"  ({'; '.join(notas)})" if notas else ""))
            archivos += [Path(f"{base}.png"), Path(f"{base}.svg")]
            dibujadas += [m for m in usadas if m not in dibujadas]
    return archivos, sorted(dibujadas)


def _a_numero(texto: str):
    try:
        return int(texto)
    except ValueError:
        try:
            return float(texto)
        except ValueError:
            return texto


def _negrita_fila1(ws):
    for c in ws[1]:
        c.font = Font(bold=True)


def _anchos(ws, cabeceras, minimo=12, maximo=26):
    for j, c in enumerate(cabeceras, start=1):
        ws.column_dimensions[get_column_letter(j)].width = max(minimo, min(maximo, len(str(c)) + 3))


def _f(v):
    return None if v is None or (isinstance(v, float) and np.isnan(v)) else float(v)


def exportar_excel(ruta_txt, ruta_reactor, df, r, t_min, T_rga, mascara, mascara_T, info,
                   t_inicio, base, manual, rampa, analisis, ruta_xlsx: Path, conv=None):
    idx = np.where(mascara)[0]
    d = df.iloc[idx].reset_index(drop=True)
    t = t_min[idx]
    T = T_rga[idx]
    tramo = info["tramo"][idx]
    if EXCEL_TODAS_LAS_MASAS:
        masas = [int(c.split()[1]) for c in d.columns if c.startswith("Mass ")]
    else:
        masas = [m for m in MASAS if f"Mass {m}" in d.columns]
    cab = {m: f"Masa {m} ({nombre_plano(MASAS[m][0])})" if m in MASAS else f"Masa {m}" for m in masas}
    masas_graf = [m for m in MASAS if f"Mass {m}" in df.columns]
    suaves = {m: suavizar(df, m).to_numpy()[idx] for m in masas_graf}

    wb = Workbook()

    ws = wb.active
    ws.title = "Datos"
    cols = ["Hora (masas)", "Tiempo (min)", "T reactor (°C)", "Tramo", "Scan"] + [cab[m] for m in masas]
    if "Total Pressure" in d.columns:
        cols.append("Presión total (mbar)")
    ws.append(cols)
    for i in range(len(d)):
        fila = [d["Time"][i].to_pydatetime(), float(t[i]), _f(T[i]), tramo[i], int(d["Scan"][i])]
        fila += [float(d[f"Mass {m}"][i]) for m in masas]
        if "Total Pressure" in d.columns:
            fila.append(float(d["Total Pressure"][i]))
        ws.append(fila)
    _negrita_fila1(ws)
    for row in ws.iter_rows(min_row=2, max_col=3):
        row[0].number_format = "dd/mm/yyyy hh:mm:ss.000"
        row[1].number_format = "0.0000"
        row[2].number_format = "0.0"
    _anchos(ws, cols)
    ws.column_dimensions["A"].width = 24
    ws.freeze_panes = "F2"

    ws2 = wb.create_sheet("Gráficos")
    iso = info.get("iso")
    cols2 = ["Tiempo (min)", "T reactor (°C)", "Tramo", "En gráfico vs T", "Tiempo en la isoterma (min)"] + \
            [f"Masa {m} ({nombre_plano(MASAS[m][0])}) suavizada" for m in masas_graf]
    ws2.append(cols2)
    en_T = mascara_T[idx]
    en_iso = info["mascara_iso"][idx] if iso is not None else np.zeros(len(idx), bool)
    for i in range(len(d)):
        panel = "Panel T" if en_T[i] else "Panel isoterma" if en_iso[i] else "No"
        t_iso = float(t[i] - iso["t_ini"]) if en_iso[i] else None
        ws2.append([float(t[i]), _f(T[i]), tramo[i], panel, t_iso] +
                   [_f(suaves[m][i]) for m in masas_graf])
    _negrita_fila1(ws2)
    for row in ws2.iter_rows(min_row=2, max_col=5):
        row[0].number_format = "0.0000"
        row[1].number_format = "0.0"
        row[4].number_format = "0.0000"
    _anchos(ws2, cols2, maximo=30)
    ws2.freeze_panes = "F2"

    ws3 = wb.create_sheet("Reactor")
    ws3.append(["Hora (reactor)", "Tiempo (min)", "T (°C)", "dT/dt (K/min, media 1 min)"])
    tr = (r["Time"] - t_inicio).dt.total_seconds().to_numpy() / 60.0
    for i in range(len(r)):
        ws3.append([r["Time"].iloc[i].to_pydatetime(), float(tr[i]), float(r["T"].iloc[i]),
                    float(rampa["dTdt"][i])])
    _negrita_fila1(ws3)
    for row in ws3.iter_rows(min_row=2, max_col=4):
        row[0].number_format = "dd/mm/yyyy hh:mm:ss"
        row[1].number_format = "0.0000"
        row[3].number_format = "0.00"
    ws3.column_dimensions["A"].width = 22
    ws3.column_dimensions["B"].width = 12
    ws3.column_dimensions["C"].width = 10
    ws3.column_dimensions["D"].width = 26
    ws3.freeze_panes = "A2"

    ws_r = wb.create_sheet("Rampa")
    ws_r.append(["Parámetro", "Valor", "Unidad"])
    for fila in [
        ("T inicial (t = 0)", round(rampa["T0"], 1), "°C"),
        ("T máxima", round(rampa["Tmax"], 1), "°C"),
        ("Tiempo hasta Tmax", round(rampa["t_Tmax"], 2), "min"),
        ("Rampa promedio (T0 -> Tmax)", round(rampa["promedio"], 2), "K/min"),
        ("Rampa máxima instantánea", round(rampa["max_inst"], 2), "K/min"),
        ("  ... a T", round(rampa["T_max_inst"], 1), "°C"),
        ("  ... a t", round(rampa["t_max_inst"], 2), "min"),
    ]:
        ws_r.append(list(fila))
    if iso is not None:
        for fila in [
            ("Rampa promedio (T0 -> inicio isoterma)", round((iso["T_ini"] - rampa["T0"]) / iso["t_ini"], 2), "K/min"),
            ("Isoterma: inicio (hora reactor)", iso["inicio"].strftime("%H:%M:%S"), ""),
            ("Isoterma: inicio / fin", f"{iso['t_ini']:.2f} / {iso['t_fin']:.2f}", "min desde t = 0"),
            ("Isoterma: duración", round(iso["t_fin"] - iso["t_ini"], 2), "min"),
            ("Isoterma: T media", round(iso["T_media"], 1), "°C"),
            ("Isoterma: T mín / máx", f"{iso['T_min']:.1f} / {iso['T_max']:.1f}", "°C"),
            ("Isoterma: T al inicio / al final", f"{iso['T_ini']:.1f} / {iso['T_fin']:.1f}", "°C"),
            ("Isoterma: deriva (pendiente)", round(iso["deriva"], 2), "K/min"),
        ]:
            ws_r.append(list(fila))
    else:
        ws_r.append(["Isoterma", "No detectada", ""])
    ws_r.append([])
    ws_r.append(["Tramo desde (°C)", "Tramo hasta (°C)", "Rampa (K/min, ajuste lineal)", "t inicio (min)", "t fin (min)"])
    fila_tramos = ws_r.max_row
    for a, b, p, ta, tb in rampa["tramos"]:
        ws_r.append([round(a), round(b, 1), round(p, 2), round(ta, 2), round(tb, 2)])
    for c in list(ws_r[1]) + list(ws_r[fila_tramos]):
        c.font = Font(bold=True)
    for j, w in enumerate((30, 18, 28, 16, 14), start=1):
        ws_r.column_dimensions[get_column_letter(j)].width = w

    ws_o = wb.create_sheet("Inicio_max")
    cols_o = ["Masa", "Especie", "Base (mbar)", "σ base (suavizado)",
              "Inicio T (°C)", "Inicio t (min)", "Inicio P (mbar)",
              "Inicio por tangente T (°C)", "Inicio por tangente t (min)",
              "50 % subida T (°C)", "50 % subida t (min)",
              "Máximo T (°C)", "Máximo t (min)", "Máximo P (mbar, suavizado)"]
    ws_o.append(cols_o)
    for m, a in analisis.items():
        ws_o.append([m, nombre_plano(MASAS[m][0]), a["base"], a["sigma"], a["T_on"], a["t_on"], a["P_on"],
                     a["T_tan"], a["t_tan"], a["T_50"], a["t_50"], a["T_max"], a["t_max"], a["P_max"]])
    _negrita_fila1(ws_o)
    for row in ws_o.iter_rows(min_row=2):
        for c in row[4:]:
            if isinstance(c.value, float):
                c.number_format = "0.00" if c.column_letter in ("F", "I", "K", "M") else "0.0"
        for c in (row[2], row[3], row[6], row[13]):
            c.number_format = "0.000E+00"
    for j in range(1, len(cols_o) + 1):
        ws_o.column_dimensions[get_column_letter(j)].width = 16
    ws_o.freeze_panes = "C2"
    nota = ws_o.max_row + 2
    ws_o.cell(row=nota, column=1, value="Cómo se calcula").font = Font(bold=True)
    for k, texto in enumerate([
        f"Base y σ: mediana (crudo) y desviación estándar (suavizado) de los primeros {VENTANA_BASE_MIN} min desde t = 0.",
        f"Inicio: primer punto (después de esos {VENTANA_BASE_MIN} min) donde la masa suavizada supera base + "
        f"{N_SIGMA_ONSET}σ y se mantiene por encima al menos {MIN_SOSTENIDO} min.",
        "Inicio por tangente: recta ajustada entre el 10 % y el 50 % de la subida, extrapolada hasta la línea base.",
        "Máximo: máximo de la masa suavizada durante el calentamiento (hasta Tmax).",
        "Vacío = la masa no sube durante el calentamiento (o baja, como el H2 consumido).",
        f"Las T no corrigen el retardo del gas reactor -> masas (DESFASE_S = {DESFASE_S} s).",
    ], start=1):
        ws_o.cell(row=nota + k, column=1, value=texto)

    ws4 = wb.create_sheet("Resumen_masas")
    cols4 = ["Masa", "Especie", "Media", "Desv. estándar", "Mín", "Máx", "Inicio (media de los 20 primeros scans)",
             "Final (media de los 20 últimos scans)", "En Tmax (media ±SUAVIZADO/2)",
             "Ruido (σ punto a punto)", "% valores ≤ 0"]
    ws4.append(cols4)
    i_max_local = int(np.where(idx == info["i_max"])[0][0])
    h = SUAVIZADO // 2
    for m in masas:
        xm = d[f"Mass {m}"].astype(float)
        ws4.append([m, nombre_plano(MASAS[m][0]) if m in MASAS else "", xm.mean(), xm.std(), xm.min(), xm.max(),
                    xm.iloc[:20].mean(), xm.iloc[-20:].mean(),
                    xm.iloc[max(0, i_max_local - h):i_max_local + h + 1].mean(),
                    xm.diff().std() / np.sqrt(2), 100 * (xm <= 0).mean()])
    _negrita_fila1(ws4)
    for j in range(1, len(cols4) + 1):
        ws4.column_dimensions[get_column_letter(j)].width = 22
    ws4.freeze_panes = "A2"

    ws5 = wb.create_sheet("Info")
    dt = d["Time"].diff().dt.total_seconds().median()
    i_rmax = int(r["T"].idxmax())
    info_filas = [
        ("Archivo masas", ruta_txt.name),
        ("Archivo reactor", ruta_reactor.name),
        ("Inicio de calentamiento (hora reactor)", t_inicio.strftime("%d/%m/%Y %H:%M:%S")),
        ("Cómo se eligió", "Manual" if manual else
         f"Automático: último punto a < {UMBRAL_INICIO} °C (o 4σ del ruido) sobre la línea base, "
         f"antes de superar línea base + {UMBRAL_DETECCION} °C"),
        ("Línea base del reactor (°C)", round(float(base), 2)),
        ("Desfase reactor -> masas (s)", DESFASE_S),
        ("t = 0 en el masas (hora)", info["t_cero"].strftime("%d/%m/%Y %H:%M:%S")),
        ("Tmax reactor (°C)", float(r["T"].iloc[i_rmax])),
        ("Hora de Tmax", r["Time"].iloc[i_rmax].strftime("%d/%m/%Y %H:%M:%S")),
        ("Tiempo de Tmax (min desde t = 0)", round(float((r["Time"].iloc[i_rmax] - t_inicio).total_seconds() / 60), 3)),
        ("Último scan usado (hora)", d["Time"].iloc[-1].strftime("%d/%m/%Y %H:%M:%S")),
        ("Duración (min)", round(float(t[-1]), 3)),
        ("Scans en la ventana", len(d)),
        ("Intervalo entre scans (s)", round(float(dt), 3)),
        ("T del reactor en cada scan", f"Interpolada en la hora de cada scan (media móvil de {SUAVIZADO_T_S} s)"),
        ("Gráficos vs T", (f"Rampa vs T + panel de la isoterma vs tiempo ({iso['T_min']:.0f}–{iso['T_max']:.0f} °C, "
                           f"T real arriba)") if iso is not None else
         ("Solo calentamiento (hasta Tmax)" if TRAMO_T == "calentamiento" else "Todo")),
        ("Suavizado de masas en gráficos", f"Media móvil de {SUAVIZADO} scans (hoja 'Graficos')"),
        ("Unidad de presión", "mbar (según encabezado '[Scan Data (Pressures in mbar)]' del txt)"),
        ("Hoja 'Datos'", "Crudos (sin suavizar) + T del reactor"),
    ]
    for k, v in info_filas:
        ws5.append([k, v])
    for row in ws5.iter_rows(min_col=1, max_col=1):
        row[0].font = Font(bold=True)
    ws5.append([])
    ws5.append(["Encabezado original del txt"])
    ws5.cell(row=ws5.max_row, column=1).font = Font(bold=True)
    for l in leer_lineas(ruta_txt):
        if l.lstrip().startswith('"[Scan Data'):
            break
        if not l.strip():
            continue
        campos = [c.strip().strip('"') for c in l.split("\t")]
        if len(campos) == 1:
            ws5.append([campos[0]])
            if campos[0].startswith("["):
                ws5.cell(row=ws5.max_row, column=1).font = Font(bold=True)
        else:
            ws5.append([campos[0], _a_numero(campos[1])])
    ws5.column_dimensions["A"].width = 38
    ws5.column_dimensions["B"].width = 80

    if conv is not None:
        hojas_conversion(wb, df, t_min, T_rga, mascara, mascara_T, info, conv)
    wb.save(ruta_xlsx)


CARPETA_BLANCOS = "Blancos"


def es_blanco(ruta: Path) -> bool:
    n = normalizar(ruta.stem).replace("-", "_")
    return any(normalizar(p) in n or normalizar(p).replace(" ", "") in n.replace(" ", "") for p in PALABRAS_BLANCO)


def _cv(x):
    x = np.asarray(x, float)
    m = np.nanmean(x)
    return float(100 * np.nanstd(x) / m) if m > 0 else np.inf


def _col(masa):
    return f"Mass {masa}"


def decidir_inerte(df, mask) -> bool:
    if USAR_INERTE in (True, False):
        return bool(USAR_INERTE) and _col(MASA_INERTE) in df.columns
    if _col(MASA_INERTE) not in df.columns:
        return False
    ar = df[_col(MASA_INERTE)].astype(float).to_numpy()[mask]
    co2 = df[_col(MASA_CO2)].astype(float).to_numpy()[mask]
    return bool(np.nanmedian(ar) / np.nanmedian(co2) > 0.02 and _cv(ar) < 5.0)


def razon(df, masa, usar_inerte, suave=True):
    if suave:
        y = suavizar(df, masa).to_numpy()
        a = suavizar(df, MASA_INERTE).to_numpy() if usar_inerte else 1.0
    else:
        y = df[_col(masa)].astype(float).to_numpy()
        a = df[_col(MASA_INERTE)].astype(float).to_numpy() if usar_inerte else 1.0
    with np.errstate(divide="ignore", invalid="ignore"):
        return y / a


def _deriva(t, y):
    n = len(y)
    m = np.mean(y)
    (p, _), cov = np.polyfit(t, y, 1, cov=True) if n > 3 else ((0.0, m), np.zeros((2, 2)))
    span = t[-1] - t[0]
    return 100 * p * span / m, 100 * np.sqrt(cov[0, 0]) * span / m


def detectar_meseta(t, crudo, suave):
    n = len(t)
    fin = np.searchsorted(t, t + MESETA_MIN, side="right")
    validas = [i for i in range(n) if fin[i] - i >= 10 and t[fin[i] - 1] - t[i] >= 0.9 * MESETA_MIN]
    if not validas:
        raise ValueError(f"la medida del blanco dura menos de {MESETA_MIN} min")
    medias = {i: float(np.mean(crudo[i:fin[i]])) for i in validas}
    tope = max(medias.values())
    ok = np.zeros(n, bool)
    for i in validas:
        d, ed = _deriva(t[i:fin[i]], crudo[i:fin[i]])
        alto = 100 * (tope - medias[i]) / tope <= MESETA_TOL_MAX
        ok[i] = alto and (abs(d) <= MESETA_DERIVA_MAX or abs(d) < 2 * ed)
    mask = np.zeros(n, bool)
    if ok.any():
        ultimo = np.where(ok)[0][-1]
        primero = ultimo
        while primero > 0 and ok[primero - 1]:
            primero -= 1
        mask[(t >= t[primero]) & (t <= t[fin[ultimo] - 1])] = True
        return mask, "automática"
    i = max(validas, key=medias.get)
    mask[i:fin[i]] = True
    return mask, "AVISO: ninguna ventana cumple los criterios; se usa la de media más alta (revisa el gráfico)"


def analizar_blanco(ruta_txt: Path, carpeta_base: Path):
    df = leer_scan(ruta_txt)
    print(f"\n>>> BLANCO: {ruta_txt.name}")
    print(f"    Masas: {df['Time'].iloc[0]:%d/%m/%Y %H:%M:%S} a {df['Time'].iloc[-1]:%H:%M:%S}  ({len(df)} scans)")
    falta = [m for m in (MASA_CO2,) if _col(m) not in df.columns]
    if falta:
        raise ValueError(f"el blanco no tiene la masa {falta}")
    t = (df["Time"] - df["Time"].iloc[0]).dt.total_seconds().to_numpy() / 60.0

    usa_ar = decidir_inerte(df, t >= t[-1] / 2)
    crudo = razon(df, MASA_CO2, usa_ar, suave=False)
    suave = razon(df, MASA_CO2, usa_ar, suave=True)

    if BLANCO_VENTANA is not None:
        a, b = BLANCO_VENTANA
        if isinstance(a, str) or isinstance(b, str):
            ref_t = df["Time"].iloc[0]
            mask = ((df["Time"] >= resolver_hora(a, ref_t)) & (df["Time"] <= resolver_hora(b, ref_t))).to_numpy()
        else:
            mask = (t >= a) & (t <= b)
        motivo = "manual (BLANCO_VENTANA)"
        if mask.sum() < 10:
            raise ValueError("BLANCO_VENTANA casi no tiene scans")
    else:
        mask, motivo = detectar_meseta(t, crudo, suave)

    w = crudo[mask]
    n = int(mask.sum())
    ref = {
        "archivo": ruta_txt.name, "nombre": ruta_txt.stem,
        "fecha": df["Time"].iloc[0].strftime("%Y-%m-%d"),
        "inicio_meseta": df["Time"][mask].iloc[0].strftime("%Y-%m-%d %H:%M:%S"),
        "fin_meseta": df["Time"][mask].iloc[-1].strftime("%Y-%m-%d %H:%M:%S"),
        "t_meseta_min": [float(t[mask][0]), float(t[mask][-1])], "scans": n, "metodo": motivo,
        "usa_inerte": usa_ar, "r44_0": float(np.mean(w)), "r44_0_std": float(np.std(w)),
        "r44_0_cv": _cv(w), "r44_0_err": media_err(w)[1],
        "deriva_pct": _deriva(t[mask], w)[0], "deriva_err": _deriva(t[mask], w)[1],
        "ruido_scan_pct": float(100 * np.std(np.diff(w)) / np.sqrt(2) / np.mean(w)),
        "r44_max_suave": float(np.nanmax(suave)),
        "I44_abs": float(df[_col(MASA_CO2)].astype(float)[mask].mean()),
        "I40_abs": float(df[_col(MASA_INERTE)].astype(float)[mask].mean()) if _col(MASA_INERTE) in df else None,
    }
    for masa, clave in ((MASA_CH4, "r15_0"), (MASA_H2, "r2_0")):
        if _col(masa) in df.columns:
            v = razon(df, masa, usa_ar, suave=False)[mask]
            ref[clave], ref[clave + "_cv"] = float(np.mean(v)), _cv(v)
            ref[clave + "_deriva"] = _deriva(t[mask], v)[0]
        else:
            ref[clave], ref[clave + "_cv"] = None, None
    if "Total Pressure" in df.columns:
        ref["P_total"] = float(df["Total Pressure"].astype(float)[mask].mean())
    if ref.get("r2_0"):
        ref["H2_CO2_senal"] = ref["r2_0"] / ref["r44_0"]

    unidad = f"I{MASA_CO2}/I{MASA_INERTE}" if usa_ar else f"I{MASA_CO2} (mbar)"
    print(f"    Normalización: {'al Ar (m/z 40)' if usa_ar else 'SIN Ar (presión parcial directa)'}")
    print(f"    Meseta ({motivo}): {ref['inicio_meseta'][11:]} a {ref['fin_meseta'][11:]}  "
          f"(t = {ref['t_meseta_min'][0]:.1f}–{ref['t_meseta_min'][1]:.1f} min, {n} scans)")
    print(f"    r44,0 = {ref['r44_0']:.5g} {unidad}  ± {ref['r44_0_err']:.2g} "
          f"({100 * ref['r44_0_err'] / ref['r44_0']:.2f} %, error de la media)  ruido scan a scan "
          f"{ref['ruido_scan_pct']:.1f} %  deriva {ref['deriva_pct']:+.2f} ± {ref['deriva_err']:.2f} %")
    if ref.get("r2_0_deriva") is not None and abs(ref["r2_0_deriva"]) > 3:
        print(f"    AVISO: el H2 (r2) todavía deriva {ref['r2_0_deriva']:+.1f} % dentro de la meseta del CO2 "
              "-> X_H2 menos confiable")
    if ref["r15_0"] is not None:
        print(f"    Fondos: r15,0 = {ref['r15_0']:.4g}" + (f"   r2,0 = {ref['r2_0']:.4g}  "
              f"(H2/CO2 en señal = {ref['H2_CO2_senal']:.2f})" if ref.get("r2_0") else ""))

    carpeta = carpeta_base / CARPETA_BLANCOS / ruta_txt.stem
    carpeta.mkdir(parents=True, exist_ok=True)
    graficar_blanco(t, crudo, suave, mask, ref, carpeta / f"{ruta_txt.stem}_meseta_CO2")
    print(f"    [OK] {ruta_txt.stem}_meseta_CO2.png / .svg")
    todo = np.ones(len(df), bool)
    for v in VISTAS[:2]:
        graficar(df, t, todo, carpeta / f"{ruta_txt.stem}_vs_tiempo{sufijo_vista(v)}", v["escala"], v.get("ylim"))
    exportar_blanco(df, t, crudo, suave, mask, ref, carpeta / f"{ruta_txt.stem}_blanco.xlsx")
    print(f"    [OK] {ruta_txt.stem}_blanco.xlsx")
    with open(carpeta_base / CARPETA_BLANCOS / f"referencia_{ruta_txt.stem}.json", "w", encoding="utf-8") as fh:
        json.dump(ref, fh, indent=2, ensure_ascii=False)
    print(f"    [OK] referencia guardada: {CARPETA_BLANCOS}/referencia_{ruta_txt.stem}.json")
    return ref


def graficar_blanco(t, crudo, suave, mask, ref, salida_base: Path):
    fig, ax = plt.subplots(figsize=(FIG_W, FIG_H))
    fig.subplots_adjust(left=MARGEN_IZQ + 0.03, right=1 - 0.25 / FIG_W, bottom=MARGEN_INF, top=1 - MARGEN_SUP)
    color = MASAS[MASA_CO2][1]
    ax.plot(t, crudo, color=color, lw=0.4, alpha=0.30)
    ax.plot(t, suave, color=color, lw=LW_CURVE)
    ax.axvspan(t[mask][0], t[mask][-1], color="#2a9d5c", alpha=0.15, lw=0)
    ax.hlines(ref["r44_0"], t[mask][0], t[mask][-1], color="k", lw=0.8)
    ax.text(0.98, 0.06, f"r$_{{44,0}}$ = {ref['r44_0']:.4g} ± {100 * ref['r44_0_err'] / ref['r44_0']:.1f} %  "
                        f"({ref['scans']} scans)",
            transform=ax.transAxes, ha="right", va="bottom", fontsize=FS_CURVA)
    ax.set_xlim(t[0], t[-1])
    lo, hi = np.nanpercentile(suave, 1), np.nanpercentile(crudo, 99.5)
    ax.set_ylim(max(0, lo - 0.05 * (hi - lo)), hi + 0.08 * (hi - lo))
    ax.set_xlabel(XLABEL_T_MIN, fontsize=FS_LABEL, fontweight="bold", labelpad=4)
    ax.set_ylabel(r"I$_{44}$ / I$_{40}$" if ref["usa_inerte"] else r"CO$_2$, m/z 44 (mbar)",
                  fontsize=FS_LABEL, fontweight="bold", labelpad=4)
    ax.xaxis.set_major_locator(MultipleLocator(paso_eje_x(t[-1] - t[0], "tiempo")))
    _estilo_ejes(ax, "linear")
    for ext in ("png", "svg"):
        fig.savefig(Path(f"{salida_base}.{ext}"), dpi=DPI, facecolor="white")
    plt.close(fig)


def exportar_blanco(df, t, crudo, suave, mask, ref, ruta_xlsx: Path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Referencia"
    ws.append(["Parámetro", "Valor"])
    unidad = f"I{MASA_CO2}/I{MASA_INERTE}" if ref["usa_inerte"] else f"I{MASA_CO2} (mbar)"
    filas = [
        ("Archivo", ref["archivo"]), ("Normalizado al Ar", "Sí" if ref["usa_inerte"] else "No"),
        ("Meseta: cómo se eligió", ref["metodo"]),
        ("Meseta: inicio / fin (hora)", f"{ref['inicio_meseta']} / {ref['fin_meseta'][11:]}"),
        ("Meseta: t inicio / fin (min desde el 1er scan)", f"{ref['t_meseta_min'][0]:.2f} / {ref['t_meseta_min'][1]:.2f}"),
        ("Scans en la meseta", ref["scans"]),
        (f"r44,0 ({unidad})", ref["r44_0"]), ("σ (scan a scan)", ref["r44_0_std"]),
        ("CV (%)", round(ref["r44_0_cv"], 3)), ("Error de la media", ref["r44_0_err"]),
        ("Error de la media (%)", round(100 * ref["r44_0_err"] / ref["r44_0"], 3)),
        ("Ruido scan a scan de r44 (%)", round(ref["ruido_scan_pct"], 2)),
        ("Deriva de r44 dentro de la meseta (%) ± error", f"{ref['deriva_pct']:+.2f} ± {ref['deriva_err']:.2f}"),
        ("Deriva de r2 (H2) dentro de la meseta (%)", _f(ref.get("r2_0_deriva"))),
        ("Máximo de r44 suavizado", ref["r44_max_suave"]),
        ("r15,0 (fondo CH3+, misma normalización)", ref["r15_0"]), ("CV r15,0 (%)", ref["r15_0_cv"]),
        ("r2,0 (H2, misma normalización)", ref["r2_0"]), ("CV r2,0 (%)", ref["r2_0_cv"]),
        ("H2/CO2 en señal (r2,0 / r44,0)", ref.get("H2_CO2_senal")),
        ("I44 absoluta en la meseta (mbar)", ref["I44_abs"]), ("I40 absoluta en la meseta (mbar)", ref["I40_abs"]),
        ("Presión total en la meseta (mbar)", ref.get("P_total")),
        ("Criterios automáticos", f"ventanas de {MESETA_MIN} min con media a ≤ {MESETA_TOL_MAX} % de la media más "
                                  f"alta y deriva ≤ {MESETA_DERIVA_MAX} % (o no significativa); se toma el último "
                                  "bloque continuo"),
    ]
    for f_ in filas:
        ws.append(list(f_))
    _negrita_fila1(ws)
    ws.column_dimensions["A"].width = 46
    ws.column_dimensions["B"].width = 44

    ws2 = wb.create_sheet("Meseta_masas")
    cols = ["Masa", "Especie", "Media (mbar)", "σ", "CV (%)", "Media / I40"]
    ws2.append(cols)
    ar = df[_col(MASA_INERTE)].astype(float)[mask].mean() if _col(MASA_INERTE) in df else np.nan
    for m in [c for c in df.columns if c.startswith("Mass ")]:
        masa = int(m.split()[1])
        x = df[m].astype(float)[mask]
        ws2.append([masa, nombre_plano(MASAS[masa][0]) if masa in MASAS else "", float(x.mean()), float(x.std()),
                    _f(_cv(x)) if np.isfinite(_cv(x)) else None, _f(x.mean() / ar)])
    _negrita_fila1(ws2)
    _anchos(ws2, cols, minimo=14)
    ws2.freeze_panes = "A2"

    ws3 = wb.create_sheet("Datos")
    masas = [m for m in MASAS if _col(m) in df.columns]
    cols3 = ["Hora (masas)", "Tiempo (min)", "r44 crudo", "r44 suavizado", "En meseta"] + \
            [f"Masa {m} ({nombre_plano(MASAS[m][0])})" for m in masas]
    ws3.append(cols3)
    for i in range(len(df)):
        ws3.append([df["Time"].iloc[i].to_pydatetime(), float(t[i]), _f(crudo[i]), _f(suave[i]),
                    "Sí" if mask[i] else ""] + [float(df[_col(m)].iloc[i]) for m in masas])
    _negrita_fila1(ws3)
    for row in ws3.iter_rows(min_row=2, max_col=2):
        row[0].number_format = "dd/mm/yyyy hh:mm:ss.000"
        row[1].number_format = "0.000"
    _anchos(ws3, cols3)
    ws3.column_dimensions["A"].width = 24
    ws3.freeze_panes = "C2"
    wb.save(ruta_xlsx)


def cargar_referencias(carpeta_base: Path) -> dict:
    refs = {}
    for p in sorted((carpeta_base / CARPETA_BLANCOS).glob("referencia_*.json")):
        try:
            with open(p, encoding="utf-8") as fh:
                ref = json.load(fh)
            refs[ref["nombre"]] = ref
        except (OSError, ValueError, KeyError) as e:
            print(f"[AVISO] no se pudo leer {p.name}: {e}")
    return refs


def elegir_blanco(nombre_run: str, inicio_run: pd.Timestamp, refs: dict):
    if not refs:
        return None, None
    if nombre_run in BLANCO_POR_ARCHIVO:
        nb = Path(BLANCO_POR_ARCHIVO[nombre_run]).stem
        if nb not in refs:
            raise ValueError(f"BLANCO_POR_ARCHIVO pide el blanco '{nb}', pero no está procesado")
        return refs[nb], "forzado en BLANCO_POR_ARCHIVO"
    dist = {k: abs((pd.Timestamp(v["inicio_meseta"]) - inicio_run).total_seconds()) for k, v in refs.items()}
    mismo_dia = [k for k, v in refs.items() if v["fecha"] == inicio_run.strftime("%Y-%m-%d")]
    if mismo_dia:
        k = min(mismo_dia, key=dist.get)
        return refs[k], "mismo día"
    k = min(dist, key=dist.get)
    return refs[k], (f"AVISO: no hay blanco del mismo día; se usa el de {refs[k]['fecha']} "
                     f"(la respuesta del masas cambia de un día a otro)")


def media_err(x):
    x = np.asarray(x, float)
    x = x[np.isfinite(x)]
    if len(x) < 2:
        return (float(x.mean()) if len(x) else float("nan")), 0.0
    m = float(np.mean(x))
    if ERROR_METODO == "bloques" and SUAVIZADO > 1:
        nb = len(x) // SUAVIZADO
        if nb >= 3:
            bm = x[:nb * SUAVIZADO].reshape(nb, SUAVIZADO).mean(axis=1)
            return m, float(np.std(bm, ddof=1) / np.sqrt(nb))
    return m, float(np.std(x, ddof=1) / np.sqrt(len(x)))


def calcular_conversion(df, t_min, T, mascara, info, ref):
    if _col(MASA_CO2) not in df.columns:
        return None
    ua = ref["usa_inerte"]
    if ua and _col(MASA_INERTE) not in df.columns:
        raise ValueError("el blanco se normalizó al Ar pero esta corrida no tiene m/z 40")
    r44 = razon(df, MASA_CO2, ua)
    X = 1 - r44 / ref["r44_0"]
    out = {"ref": ref, "r44": r44, "X": X, "XH": None, "r15": None, "A": None, "S": None, "Y": None}
    tramo = info["tramo"]
    if ua:
        ar_s = suavizar(df, MASA_INERTE).to_numpy()
        nivel = np.nanmedian(ar_s[mascara & (t_min >= np.nanmax(t_min[mascara]) / 2)])
        out["gas"] = ar_s >= 0.8 * nivel
    else:
        out["gas"] = np.ones(len(df), bool)
    sin_reac = mascara & (tramo == "Rampa") & (T >= VENTANA_SIN_REACCION_T[0]) & (T <= VENTANA_SIN_REACCION_T[1])
    hay_ventana = sin_reac.sum() >= 5
    out["origen_fondos"] = (f"propia corrida ({VENTANA_SIN_REACCION_T[0]}–{VENTANA_SIN_REACCION_T[1]} °C)"
                            if hay_ventana else "blanco")
    if _col(MASA_H2) in df.columns:
        out["r2"] = razon(df, MASA_H2, ua)
        out["r2_0"] = float(np.nanmedian(out["r2"][sin_reac])) if hay_ventana else ref.get("r2_0")
        if out["r2_0"]:
            out["XH"] = 1 - out["r2"] / out["r2_0"]
    if _col(MASA_CH4) in df.columns:
        out["r15"] = razon(df, MASA_CH4, ua)
        out["r15_0"] = float(np.nanmedian(out["r15"][sin_reac])) if hay_ventana else ref.get("r15_0")
        if out["r15_0"] is not None:
            out["A"] = out["r15"] - out["r15_0"]

    out["X_sin_reac"] = float(np.nanmedian(X[sin_reac])) if hay_ventana else None
    out["Ar_rel"] = None
    if hay_ventana and ua and ref.get("I40_abs"):
        out["Ar_rel"] = float(np.nanmedian(df[_col(MASA_INERTE)].astype(float).to_numpy()[sin_reac]) / ref["I40_abs"])

    out["est"] = None
    m_iso = info.get("mascara_iso")
    if m_iso is not None and m_iso.any():
        t_fin = float(np.nanmax(t_min[m_iso]))
        est = m_iso & (t_min >= t_fin - ESTABLE_MIN)
        if ua:
            ar = suavizar(df, MASA_INERTE).to_numpy()
            est &= np.abs(ar / np.nanmedian(ar[m_iso]) - 1) <= 0.10
        if est.sum() >= 10:
            n = int(est.sum())
            c44 = razon(df, MASA_CO2, ua, suave=False)[est]
            e = {"mask": est, "n": n, "t": (float(t_min[est][0]), float(t_min[est][-1])),
                 "T": float(np.nanmean(T[est])), "X": float(1 - np.nanmean(c44) / ref["r44_0"])}
            e["X_err"] = media_err(c44)[1] / ref["r44_0"]
            e["X_err_blanco"] = float((1 - e["X"]) * ref["r44_0_err"] / ref["r44_0"])
            e["X_err_total"] = float(np.hypot(e["X_err"], e["X_err_blanco"]))
            e["XH"] = e["XH_err"] = e["S_H2"] = e["S_H2_err"] = None
            if out["XH"] is not None and _col(MASA_H2) in df.columns:
                c2 = razon(df, MASA_H2, ua, suave=False)
                if hay_ventana:
                    r2_0, s_r2_0 = media_err(c2[sin_reac])
                else:
                    r2_0, s_r2_0 = out["r2_0"], 0.0
                ce = c2[est]
                e["XH"] = float(1 - np.nanmean(ce) / r2_0)
                e["XH_err"] = float(np.hypot(media_err(ce)[1] / r2_0, (1 - e["XH"]) * s_r2_0 / r2_0))
                e["r2_0_corrida"] = r2_0
                if e["X"] > X_MIN_CONV:
                    e["S_H2"] = (R_H2_CO2 * e["XH"] / e["X"] - 1) / 3
                    e["S_H2_err"] = (R_H2_CO2 / 3) * float(np.hypot(e["XH_err"] / e["X"],
                                                                    e["XH"] * e["X_err_total"] / e["X"] ** 2))
            e["A"] = float(np.nanmean(out["A"][est])) if out["A"] is not None else None
            e["F_C"] = ((ref["r44_0"] - float(np.nanmean(r44[est]))) / e["A"]) if e["A"] and e["A"] > 0 else None
            e["r44_0_min"] = (float(np.nanmean(r44[est])) / (1 - e["XH"])) if e["XH"] is not None else None
            e["S_alt"], e["S_sist"] = {}, None
            if e.get("S_H2") is not None:
                c44m = float(np.nanmean(c44))
                r2b = ref.get("r2_0")
                if r2b and e.get("r2_0_corrida") and abs(e["r2_0_corrida"] / r2b - 1) * 100 <= REF_H2_VALIDA:
                    xh_b = 1 - float(np.nanmean(c2[est])) / r2b
                    e["S_alt"]["blanco/blanco"] = (R_H2_CO2 * xh_b / e["X"] - 1) / 3
                xs = out.get("X_sin_reac")
                if hay_ventana and xs is not None and abs(xs) * 100 <= REF_X_VALIDA:
                    r0c = float(np.nanmean(razon(df, MASA_CO2, ua, suave=False)[sin_reac]))
                    x_c = 1 - c44m / r0c
                    e["S_alt"]["corrida/corrida"] = (R_H2_CO2 * e["XH"] / x_c - 1) / 3
                if e["S_alt"]:
                    e["S_sist"] = float(max(abs(v - e["S_H2"]) for v in e["S_alt"].values()))
            out["est"] = e
    out["F_C"] = out["est"]["F_C"] if out["est"] else None

    out["tmax"] = None
    out["hasta_tmax"] = None
    Tm = np.where(mascara & out["gas"], T, np.nan)
    if np.isfinite(Tm).any():
        k = int(np.nanargmax(Tm))
        out["hasta_tmax"] = mascara & (t_min >= 0) & (np.arange(len(T)) <= k)
    t_pico = info.get("t_max_ch4")
    if t_pico is not None:
        sel = mascara & out["gas"] & (np.abs(t_min - t_pico) <= VENTANA_PICO_MIN)
        if ua and sel.any():
            ar = suavizar(df, MASA_INERTE).to_numpy()
            sel &= np.abs(ar / np.nanmedian(ar[sel]) - 1) <= 0.10
        if sel.sum() >= 10:
            n = int(sel.sum())
            c44 = razon(df, MASA_CO2, ua, suave=False)[sel]
            tm = {"mask": sel, "n": n, "t": (float(t_min[sel][0]), float(t_min[sel][-1])),
                  "t_pico": float(t_pico), "Tmax": float(info["T_max_ch4"]), "T": float(np.nanmean(T[sel])),
                  "T_rango": (float(np.nanmin(T[sel])), float(np.nanmax(T[sel]))),
                  "X": float(1 - np.nanmean(c44) / ref["r44_0"])}
            tm["X_err"] = media_err(c44)[1] / ref["r44_0"]
            tm["X_err_blanco"] = float((1 - tm["X"]) * ref["r44_0_err"] / ref["r44_0"])
            tm["X_err_total"] = float(np.hypot(tm["X_err"], tm["X_err_blanco"]))
            tm["XH"] = tm["XH_err"] = None
            if out["XH"] is not None and _col(MASA_H2) in df.columns:
                c2 = razon(df, MASA_H2, ua, suave=False)
                if hay_ventana:
                    r2_0, s_r2_0 = media_err(c2[sin_reac])
                else:
                    r2_0, s_r2_0 = out["r2_0"], 0.0
                ce = c2[sel]
                tm["XH"] = float(1 - np.nanmean(ce) / r2_0)
                tm["XH_err"] = float(np.hypot(media_err(ce)[1] / r2_0, (1 - tm["XH"]) * s_r2_0 / r2_0))
            out["tmax"] = tm

    tf = TABLA_FRAG
    out["F_tabla"] = (tf["CO2"]["FF"] * tf["CO2"]["S"]) / (tf["CH4"]["FF"] * tf["CH4"]["S"])
    out["F"] = out["F_err"] = out["F_origen"] = None
    r0, s_r0 = ref["r44_0"], ref["r44_0_err"]
    c15 = razon(df, MASA_CH4, ua, suave=False) if _col(MASA_CH4) in df.columns else None
    c44_all = razon(df, MASA_CO2, ua, suave=False)
    r15_0_raw = s_r15_0 = None
    if c15 is not None and hay_ventana:
        r15_0_raw, s_r15_0 = media_err(c15[sin_reac])

    def _area(sel):
        m15, s15 = media_err(c15[sel])
        return float(m15 - r15_0_raw), float(np.hypot(s15, s_r15_0))

    e = out["est"]
    if e is not None and r15_0_raw is not None:
        e["A"], e["A_err"] = _area(e["mask"])
        c44 = c44_all[e["mask"]]
        d44 = r0 - float(np.nanmean(c44))
        s_d44 = float(np.hypot(s_r0, media_err(c44)[1]))
        if e["A"] > 0 and d44 > 0:
            e["F_C"] = d44 / e["A"]
            e["F_C_err"] = e["F_C"] * float(np.hypot(s_d44 / d44, e["A_err"] / e["A"]))
            out["F_C"] = e["F_C"]
    if F == "datos":
        if e is not None and e.get("S_H2") is not None and e.get("F_C_err") is not None:
            out["F"] = e["S_H2"] * e["F_C"]
            out["F_err"] = out["F"] * float(np.hypot(e["S_H2_err"] / e["S_H2"], e["F_C_err"] / e["F_C"]))
            out["F_origen"] = "calibrado en esta corrida: S_CH4(balance de H2) × F_C (final de la isoterma)"
    elif isinstance(F, (int, float)):
        out["F"], out["F_err"], out["F_origen"] = float(F), 0.0, "fijo (F en la configuración)"

    if out["F"] is not None and out["A"] is not None:
        Fv, sF = out["F"], out["F_err"]
        out["Y"] = Fv * out["A"] / r0
        if e is not None and e.get("A") is not None:
            if F == "datos":
                e["S"], e["S_err"] = e["S_H2"], e["S_H2_err"]
                e["Y"] = e["X"] * e["S"]
                e["Y_err"] = float(np.hypot(R_H2_CO2 * e["XH_err"], e["X_err_total"]) / 3)
            else:
                e["Y"] = Fv * e["A"] / r0
                e["Y_err"] = e["Y"] * float(np.hypot(e["A_err"] / e["A"], s_r0 / r0))
                e["S"] = e["Y"] / e["X"]
                e["S_err"] = e["S"] * float(np.hypot(e["Y_err"] / e["Y"], e["X_err_total"] / e["X"]))
        tm = out["tmax"]
        if tm is not None and r15_0_raw is not None:
            tm["A"], tm["A_err"] = _area(tm["mask"])
            tm["Y"] = Fv * tm["A"] / r0
            tm["Y_err"] = abs(tm["Y"]) * float(np.hypot(np.hypot(sF / Fv, tm["A_err"] / tm["A"]), s_r0 / r0))

    if e is not None:
        tt = t_min[e["mask"]]
        dur = float(tt[-1] - tt[0])
        cambio = lambda y: float(100 * np.polyfit(tt, y[e["mask"]], 1)[0] * dur) if y is not None else None
        e["deriva"] = {"X": cambio(X), "XH": cambio(out["XH"]), "Y": cambio(out["Y"]),
                       "T": float(np.polyfit(tt, T[e["mask"]], 1)[0] * dur), "dur": dur,
                       "iso_dur": float(info["iso"]["t_fin"] - info["iso"]["t_ini"]) if info.get("iso") else None}
        e["estacionario"] = all(abs(v) <= DERIVA_ESTABLE_MAX for k, v in e["deriva"].items()
                                if k in ("X", "XH", "Y") and v is not None)
    return out


def resumir_conversion_T(t_min, T, mascara_T, conv, paso=25.0):
    filas = []
    if not mascara_T.any():
        return filas
    Tm = T[mascara_T]
    a = paso * np.floor(np.nanmin(Tm) / paso)
    while a < np.nanmax(Tm):
        s = mascara_T & (T >= a) & (T < a + paso)
        if s.sum() >= 3:
            fila = [a, a + paso, int(s.sum()), float(np.nanmean(T[s])), float(np.nanmean(conv["X"][s]) * 100),
                    float(np.nanstd(conv["X"][s]) * 100)]
            fila.append(float(np.nanmean(conv["XH"][s]) * 100) if conv["XH"] is not None else None)
            fila.append(float(np.nanmean(conv["Y"][s]) * 100) if conv["Y"] is not None else None)
            fila.append(float(np.nanstd(conv["Y"][s]) * 100) if conv["Y"] is not None else None)
            fila += [None, None]
            filas.append(fila)
        a += paso
    return filas


def graficar_conversion(t_min, T, mascara, mascara_T, conv, info, salida_base: Path):
    curvas = [("X", r"X$_{CO_2}$", MASAS[MASA_CO2][1])]
    if conv["XH"] is not None:
        curvas.append(("XH", r"X$_{H_2}$", MASAS[MASA_H2][1]))
    if conv["Y"] is not None:
        curvas.append(("Y", r"Y$_{CH_4}$", MASAS[MASA_CH4][1]))
    archivos = []
    mascara_T = conv["hasta_tmax"] if conv.get("hasta_tmax") is not None else mascara_T
    for eje, mask, sufijo in (("tiempo", mascara, "_conversion_vs_tiempo"), ("temp", mascara_T, "_conversion_vs_temp")):
        mask = mask & conv["gas"]
        if mask.sum() < 5:
            continue
        x = t_min if eje == "tiempo" else T
        fig, ax = plt.subplots(figsize=(FIG_W, FIG_H))
        fig.subplots_adjust(left=MARGEN_IZQ, right=1 - (0.65 if eje == "tiempo" else 0.25) / FIG_W,
                            bottom=MARGEN_INF, top=1 - MARGEN_SUP)
        orden = np.argsort(x[mask]) if eje == "temp" else slice(None)
        for clave, etiqueta, color in curvas:
            ax.plot(x[mask][orden], 100 * conv[clave][mask][orden], color=color, lw=LW_CURVE, label=etiqueta)
        ax.axhline(0, color="#999999", lw=0.4, ls=":")
        iso = info.get("iso")
        if eje == "tiempo":
            ax.set_xlim(-float(MIN_ANTES), float(np.nanmax(t_min[mask])))
            if iso is not None:
                ax.axvspan(iso["t_ini"], iso["t_fin"], color="#f5b041", alpha=0.15, lw=0)
            if MOSTRAR_ESTABLE and conv.get("est") is not None:
                ax.axvspan(*conv["est"]["t"], color="#2a9d5c", alpha=0.18, lw=0)
            if conv.get("tmax") is not None:
                ax.axvspan(*conv["tmax"]["t"], color="#c0392b", alpha=0.12, lw=0)
            ax2 = ax.twinx()
            ax2.plot(t_min[mask], T[mask], color=COLOR_T, lw=0.6, ls="--")
            ax2.set_ylabel(YLABEL_TEMP, fontsize=FS_LABEL, fontweight="bold", color=COLOR_T, labelpad=4)
            ax2.tick_params(axis="y", labelsize=FS_TICK, length=TICK_MAJOR, width=LW_AXES, direction="in",
                            colors=COLOR_T)
            for sp in ax2.spines.values():
                sp.set_linewidth(LW_AXES)
        else:
            t_lo, t_hi = float(np.nanmin(T[mask])), float(np.nanmax(T[mask]))
            ax.set_xlim(t_lo, t_hi + 0.02 * (t_hi - t_lo))
            lineas = []
            tm, es = conv.get("tmax"), conv.get("est")
            if tm is not None:
                ax.errorbar([tm["Tmax"]], [100 * tm["X"]], yerr=[100 * tm["X_err_total"]], fmt="o", ms=5,
                            color=MASAS[MASA_CO2][1], mec="black", mew=0.6, capsize=2, zorder=5)
                txt = (f"Máximo de CH$_4$ ({tm['Tmax']:.1f} °C): X$_{{CO_2}}$ {100 * tm['X']:.1f} ± "
                       f"{100 * tm['X_err_total']:.1f} %")
                if tm.get("Y") is not None:
                    ax.errorbar([tm["Tmax"]], [100 * tm["Y"]], yerr=[100 * tm["Y_err"]], fmt="s", ms=4.5,
                                color=MASAS[MASA_CH4][1], mec="black", mew=0.6, capsize=2, zorder=5)
                    txt += f", Y$_{{CH_4}}$ {100 * tm['Y']:.1f} ± {100 * tm['Y_err']:.1f} %"
                lineas.append(txt)
            if MOSTRAR_ESTABLE and es is not None:
                txt = f"Final de la isoterma ({es['T']:.1f} °C): X$_{{CO_2}}$ {100 * es['X']:.1f} ± {100 * es['X_err_total']:.1f} %"
                if es.get("S") is not None:
                    txt += (f", S$_{{CH_4}}$ {100 * es['S']:.1f} ± {100 * es['S_err']:.1f} %, "
                            f"Y$_{{CH_4}}$ {100 * es['Y']:.1f} ± {100 * es['Y_err']:.1f} %")
                lineas.append(txt)
            if REPORTE_T != "Tmax":
                lineas.reverse()
            for j, txt in enumerate(lineas):
                ax.text(0.02, 0.22 + 0.065 * (len(lineas) - 1 - j), txt, transform=ax.transAxes, ha="left",
                        va="bottom", fontsize=FS_CURVA - 0.5, zorder=6,
                        bbox=dict(boxstyle="square,pad=0.1", fc="white", ec="none", alpha=0.85))
            if conv.get("X_sin_reac") is not None and conv["X_sin_reac"] > 0.05:
                ax.text(0.98, 0.04, "Rampa: la X$_{CO_2}$ aparente incluye CO$_2$ retenido por el catalizador",
                        transform=ax.transAxes, ha="right", va="bottom", fontsize=FS_CURVA - 1, color="#666666",
                        bbox=dict(boxstyle="square,pad=0.15", fc="white", ec="none"), zorder=6)
        lo, hi = ax.get_ylim()
        ax.set_ylim(max(lo, -20), min(hi, 105))
        span = ax.get_xlim()[1] - ax.get_xlim()[0]
        ax.xaxis.set_major_locator(MultipleLocator(paso_eje_x(span, eje)))
        ax.set_xlabel(XLABEL_T_MIN if eje == "tiempo" else XLABEL_TEMP, fontsize=FS_LABEL, fontweight="bold",
                      labelpad=4)
        ax.set_ylabel("Conversión / rendimiento (%)" if conv["Y"] is not None else "Conversión (%)",
                      fontsize=FS_LABEL, fontweight="bold", labelpad=4)
        _estilo_ejes(ax, "linear")
        ax.legend(loc="center left" if eje == "temp" else "upper left", frameon=False, fontsize=FS_CURVA,
                  handlelength=1.7)
        for ext in ("png", "svg"):
            fig.savefig(Path(f"{salida_base}{sufijo}.{ext}"), dpi=DPI, facecolor="white")
        plt.close(fig)
        archivos.append(f"{salida_base.name}{sufijo}")
    return archivos


def hojas_conversion(wb, df, t_min, T_rga, mascara, mascara_T, info, conv):
    ref = conv["ref"]
    idx = np.where(mascara)[0]
    ws = wb.create_sheet("Conversión", 1)
    cols = ["Tiempo (min)", "T reactor (°C)", "Tramo", "r44 (suav.)", "X_CO2 (%)"]
    if conv["XH"] is not None:
        cols += ["r2 (suav.)", "X_H2 (%)"]
    if conv["A"] is not None:
        cols += ["r15 (suav.)", "A_CH4 = r15 - r15,0", "(r44,0 - r44) / A_CH4"]
    if conv["Y"] is not None:
        cols += ["Y_CH4 (%)"]
    ws.append(cols)
    for i in idx:
        fila = [float(t_min[i]), _f(T_rga[i]), info["tramo"][i], _f(conv["r44"][i]), _f(100 * conv["X"][i])]
        if conv["XH"] is not None:
            fila += [_f(conv["r2"][i]), _f(100 * conv["XH"][i])]
        if conv["A"] is not None:
            a = conv["A"][i]
            fila += [_f(conv["r15"][i]), _f(a),
                     _f((ref["r44_0"] - conv["r44"][i]) / a) if a > 0 and conv["X"][i] > X_MIN_CONV else None]
        if conv["Y"] is not None:
            fila += [_f(100 * conv["Y"][i])]
        ws.append(fila)
    _negrita_fila1(ws)
    for row in ws.iter_rows(min_row=2):
        row[0].number_format = "0.000"
        row[1].number_format = "0.0"
        for c in row[3:]:
            if isinstance(c.value, float):
                c.number_format = "0.00" if "(%)" in str(ws.cell(row=1, column=c.column).value) else "0.0000E+00"
    _anchos(ws, cols, maximo=24)
    ws.freeze_panes = "D2"

    ws2 = wb.create_sheet("Conversión_por_T", 2)
    ws2.append(["T desde (°C)", "T hasta (°C)", "Scans", "T media (°C)", "X_CO2 (%)", "σ X_CO2 (%)",
                "X_H2 (%)", "Y_CH4 (%)", "σ Y_CH4 (%)", "S_CH4 (%)", "σ S_CH4 (%)"])
    m_T = conv["hasta_tmax"] if conv.get("hasta_tmax") is not None else mascara_T
    for fila in resumir_conversion_T(None, T_rga, m_T & conv["gas"], conv):
        ws2.append(fila)
    ws2.append([])
    if conv.get("tmax") is not None:
        e = conv["tmax"]
        ws2.append([f"A T máx del pico de CH4 ({e['Tmax']} °C; ± {VENTANA_PICO_MIN:g} min del máximo de m/z "
                    f"{MASA_CH4})", None, e["n"],
                    (e["T"]), (100 * e["X"]), (100 * e["X_err_total"]),
                    (100 * e["XH"]) if e["XH"] is not None else None,
                    (100 * e["Y"]) if e.get("Y") is not None else None,
                    (100 * e["Y_err"]) if e.get("Y") is not None else None, None, None])
        ws2.append(["(a T máx el catalizador todavía convierte en CH4 el CO2 que guardó: Y_CH4 puede superar a "
                    "X_CO2 y S no se da; X_H2 tampoco es representativo. S_CH4 se da al final de la isoterma)"])
    if conv.get("est") is not None:
        e = conv["est"]
        ws2.append([f"Final de la isoterma (últimos {ESTABLE_MIN:g} min de la isoterma)", None, e["n"], (e["T"]),
                    (100 * e["X"]), (100 * e["X_err_total"]),
                    (100 * e["XH"]) if e["XH"] is not None else None,
                    (100 * e["Y"]) if e.get("Y") is not None else None,
                    (100 * e["Y_err"]) if e.get("Y") is not None else None,
                    (100 * e["S"]) if e.get("S") is not None else
                    ((100 * e["S_H2"]) if e["S_H2"] is not None else None),
                    (100 * e["S_err"]) if e.get("S") is not None else
                    ((100 * e["S_H2_err"]) if e["S_H2"] is not None else None)])
        ws2.append(["(final de la isoterma: σ = ruido y blanco en cuadratura; S_CH4 por balance de H2; Y_CH4 = X_CO2·S_CH4)"])
    if conv.get("Y") is not None:
        ws2.append([f"Y_CH4 = F·A_CH4/r44,0 con F = {conv['F']:.3f} ± {conv['F_err']:.3f} ({conv['F_origen']}). "
                    f"En la rampa S_CH4 no se da: X_CO2 incluye CO2 que el catalizador guarda y el CH4 el que se "
                    f"forma con ese CO2 guardado. σ Y_CH4 en la rampa = dispersión dentro del tramo."])
    if conv.get("X_sin_reac") is not None and conv["X_sin_reac"] > 0.05:
        ws2.append([f"AVISO: en la rampa a {VENTANA_SIN_REACCION_T[0]}–{VENTANA_SIN_REACCION_T[1]} °C (sin reacción) "
                    f"X_CO2 aparente = {100 * conv['X_sin_reac']} %: el catalizador guarda CO2. "
                    "Los valores de la rampa NO son conversión."])
    _negrita_fila1(ws2)
    for row in ws2.iter_rows(min_row=2):
        for c in row[3:]:
            if isinstance(c.value, float):
                c.number_format = "0.0" if c.column == 4 else "0.00"
    for j in range(1, 10):
        ws2.column_dimensions[get_column_letter(j)].width = 14
    ws2.column_dimensions["A"].width = 24

    ws3 = wb.create_sheet("Referencia_blanco", 3)
    ws3.append(["Parámetro", "Valor"])
    for k, v in [
        ("Blanco usado", ref["archivo"]), ("Por qué este blanco", conv.get("motivo_blanco")),
        ("Meseta del blanco", f"{ref['inicio_meseta']} – {ref['fin_meseta'][11:]} ({ref['scans']} scans, {ref['metodo']})"),
        ("Normalizado al Ar", "Sí" if ref["usa_inerte"] else "No"),
        ("r44,0", ref["r44_0"]), ("CV r44,0 (%)", (ref["r44_0_cv"])),
        ("Error de la media de r44,0", ref["r44_0_err"]),
        ("r2,0 y r15,0 tomados de", conv.get("origen_fondos")),
        ("r15,0", conv.get("r15_0")), ("r2,0", conv.get("r2_0")),
        (f"X_CO2 aparente sin reacción ({VENTANA_SIN_REACCION_T[0]}–{VENTANA_SIN_REACCION_T[1]} °C, rampa) (%)",
         _f(100 * conv["X_sin_reac"]) if conv.get("X_sin_reac") is not None else None),
        ("I40 corrida / I40 blanco (misma ventana)", conv.get("Ar_rel")),
        (f"T máx del pico de CH4 (m/z {MASA_CH4}) (°C)", (conv["tmax"]["Tmax"]) if conv.get("tmax") else None),
        (f"A T máx: tramo (min), ± {VENTANA_PICO_MIN:g} min del máximo; T del tramo (°C)",
         f"{conv['tmax']['t'][0]}–{conv['tmax']['t'][1]}; {conv['tmax']['T_rango'][0]}–"
         f"{conv['tmax']['T_rango'][1]}" if conv.get("tmax") else None),
        ("A T máx: X_CO2 (%) ± σ (ruido y blanco en cuadratura)",
         f"{100 * conv['tmax']['X']} ± {100 * conv['tmax']['X_err_total']}" if conv.get("tmax") else None),
        ("Final de la isoterma: tramo (min)", f"{conv['est']['t'][0]}–{conv['est']['t'][1]}" if conv.get("est") else None),
        ("Final de la isoterma: cambio en la ventana (X_CO2 / X_H2 / Y_CH4 en puntos; T en °C)",
         (f"{conv['est']['deriva']['X']:+.1f} / {conv['est']['deriva']['XH']:+.1f} / "
          f"{conv['est']['deriva']['Y']:+.1f}; T {conv['est']['deriva']['T']:+.1f}"
          + ("" if conv["est"]["estacionario"] else "  -> NO estacionario: valores al final de la isoterma"))
         if conv.get("est") and conv["est"].get("deriva") and conv["est"]["deriva"]["XH"] is not None
         and conv["est"]["deriva"]["Y"] is not None else None),
        ("Final de la isoterma: T media (°C)", (conv["est"]["T"]) if conv.get("est") else None),
        ("Final de la isoterma: X_CO2 (%)", (100 * conv["est"]["X"]) if conv.get("est") else None),
        ("  σ ruido / σ blanco / σ total en cuadratura (puntos %)",
         f"{100 * conv['est']['X_err']} / {100 * conv['est']['X_err_blanco']} / "
         f"{100 * conv['est']['X_err_total']}" if conv.get("est") else None),
        ("Final de la isoterma: X_H2 (%) ± σ",
         f"{100 * conv['est']['XH']} ± {100 * conv['est']['XH_err']}"
         if conv.get("est") and conv["est"]["XH"] is not None else None),
        (f"Final de la isoterma: S_CH4 por balance de H2 (R = {R_H2_CO2:g}) (%) ± σ",
         f"{100 * conv['est']['S_H2']} ± {100 * conv['est']['S_H2_err']}"
         if conv.get("est") and conv["est"]["S_H2"] is not None else None),
        ("Errores", "σ = desviación estándar de los scans crudos / √n (error de la media); "
                    "σ_X(blanco) = (1 - X)·σ(r44,0)/r44,0; se combinan en cuadratura. No incluye el error "
                    "sistemático por usar un blanco de otro día."),
        ("Final de la isoterma: r44,0 mínimo compatible con el H2 (S = 100 %)",
         conv["est"]["r44_0_min"] if conv.get("est") else None),
        ("Final de la isoterma: F_C = (r44,0 - r44)/A_CH4 (cota superior de F si hay CO)", conv.get("F_C")),
        ("F usado para Y y S", f"{conv['F']:.3f} ± {conv['F_err']:.3f}  ({conv['F_origen']})" if conv.get("F") is not None
         else "ninguno"),
        ("F de la tabla de fragmentación (70 eV), solo comparación",
         f"{conv['F_tabla']:.3f} = FF44·S_CO2/(FF15·S_CH4) = {TABLA_FRAG['CO2']['FF']}·{TABLA_FRAG['CO2']['S']}/"
         f"({TABLA_FRAG['CH4']['FF']}·{TABLA_FRAG['CH4']['S']})"),
        ("Final de la isoterma: Y_CH4 (%) ± σ", f"{100 * conv['est']['Y']} ± {100 * conv['est']['Y_err']}"
         if conv.get("est") and conv["est"].get("Y") is not None else None),
        ("A T máx: Y_CH4 (%) ± σ", f"{100 * conv['tmax']['Y']} ± {100 * conv['tmax']['Y_err']}"
         if conv.get("tmax") and conv["tmax"].get("Y") is not None else None),
        ("Fórmulas", "X_CO2 = 1 - r44/r44,0 (r44,0 del blanco) ; X_H2 = 1 - r2/r2,0 ; A_CH4 = r15 - r15,0 ; "
                     "S_CH4(H2) = (R·X_H2/X_CO2 - 1)/3 ; F_C = (r44,0 - r44)/A_CH4 ; S = F·A_CH4/(r44,0 - r44) ; "
                     "Y = F·A_CH4/r44,0. "
                     "r = señal suavizada / I40 suavizada (si se normaliza al Ar)."),
    ]:
        ws3.append([k, v])
    _negrita_fila1(ws3)
    ws3.column_dimensions["A"].width = 52
    ws3.column_dimensions["B"].width = 70
    un_decimal(ws, ws2, ws3)


DECIMALES = 1


def _redondear_texto(txt):
    return re.sub(r"-?\d+\.\d{2,}", lambda m: f"{float(m.group()):.{DECIMALES}f}", txt)


def un_decimal(ws, ws2, ws3):
    cab = [c.value or "" for c in ws[1]]
    cols = [j for j, h in enumerate(cab) if "(%)" in h or "(°C)" in h or h == "Tiempo (min)"]
    for row in ws.iter_rows(min_row=2):
        for j in cols:
            if isinstance(row[j].value, float):
                row[j].value = round(row[j].value, DECIMALES)
                row[j].number_format = "0." + "0" * DECIMALES
    for row in ws2.iter_rows(min_row=2):
        for j, c in enumerate(row):
            if isinstance(c.value, float) and j != 2:
                c.value = round(c.value, DECIMALES)
                c.number_format = "0." + "0" * DECIMALES
            elif isinstance(c.value, str) and "F =" not in c.value:
                c.value = _redondear_texto(c.value)
    for row in ws3.iter_rows(min_row=2):
        clave = str(row[0].value or "")
        if any(k in clave for k in ("X_CO2", "X_H2", "S_CH4", "Y_CH4", "T máx", "T media", "tramo", "σ ruido")):
            v = row[1].value
            if isinstance(v, float):
                row[1].value = round(v, DECIMALES)
            elif isinstance(v, str):
                row[1].value = _redondear_texto(v)


def pedir_curvas_a_eliminar(disponibles):
    while True:
        r = preguntar("    ¿Cuál(es) eliminar? Escribe la(s) masa(s), ej: 16, 15, 18  [Enter = cancelar]: ")
        if r == "":
            return set()
        elegidas = [int(x) for x in re.findall(r"\d+", r)]
        invalidas = [m for m in elegidas if m not in disponibles]
        if not elegidas or invalidas:
            print(f"    Aviso: {invalidas or r!r} no está en la lista. Opciones: {', '.join(map(str, disponibles))}")
            continue
        if len(set(elegidas)) >= len(disponibles):
            print("    Aviso: tiene que quedar al menos una curva.")
            continue
        return set(elegidas)


def procesar(ruta_txt: Path, reactores: dict, carpeta_base: Path, refs=None):
    df = leer_scan(ruta_txt)
    print(f"\n>>> {ruta_txt.name}")
    print(f"    Masas: {df['Time'].iloc[0]:%d/%m/%Y %H:%M:%S} a {df['Time'].iloc[-1]:%H:%M:%S}  ({len(df)} scans)")
    ruta_reactor = emparejar(ruta_txt, df, reactores)
    if ruta_reactor is None:
        raise ValueError("ningún archivo 'Prueba' del reactor coincide en hora con este archivo del masas")
    r = reactores[ruta_reactor]
    print(f"    Reactor: {ruta_reactor.name}")

    t_inicio, base, manual = elegir_inicio(r)
    t_min, T_rga, mascara, mascara_T, info = sincronizar(df, r, t_inicio)
    info.update(mascara=mascara, t_min=t_min)
    print(f"    t = 0 en {info['t_cero']:%H:%M:%S}  ->  hasta {df['Time'][mascara].iloc[-1]:%H:%M:%S}  "
          f"({t_min[mascara][-1]:.1f} min, {mascara.sum()} scans, "
          f"T {np.nanmin(T_rga[mascara]):.0f}-{np.nanmax(T_rga[mascara]):.0f} °C)")

    nombre = ruta_txt.stem
    if manual:
        nombre += f"_t0{t_inicio:%H%M%S}"
    if T_FIN is not None:
        fin = f"{T_FIN:g}".replace(".", "_") + "min" if isinstance(T_FIN, (int, float)) else re.sub(r"\D", "", str(T_FIN))
        nombre += f"_hasta{fin}"
    carpeta = carpeta_base / nombre
    carpeta.mkdir(parents=True, exist_ok=True)
    print(f"    Carpeta: {carpeta}")

    rampa = calcular_rampa(r, t_inicio)
    print(f"    Rampa: promedio {rampa['promedio']:.2f} K/min ({rampa['T0']:.0f} -> {rampa['Tmax']:.1f} °C en "
          f"{rampa['t_Tmax']:.1f} min); por tramos: " +
          ", ".join(f"{a:.0f}-{b:.0f} °C {p:.1f}" for a, b, p, _, _ in rampa["tramos"]) + " K/min")
    iso = info["iso"]
    if iso is not None:
        print(f"    Isoterma: desde t = {iso['t_ini']:.1f} min hasta {iso['t_fin']:.1f} min "
              f"({iso['t_fin'] - iso['t_ini']:.1f} min), T real {iso['T_min']:.1f}–{iso['T_max']:.1f} °C "
              f"(media {iso['T_media']:.1f} °C, deriva {iso['deriva']:.2f} K/min)")
    else:
        print("    Isoterma: no detectada (los gráficos vs T llegan hasta Tmax)")
    idx = np.where(mascara)[0]
    analisis = {}
    for m in MASAS:
        if f"Mass {m}" in df.columns:
            analisis[m] = analizar_masa(t_min[idx], T_rga[idx], df[f"Mass {m}"].astype(float).to_numpy()[idx],
                                        suavizar(df, m).to_numpy()[idx], info["tramo"][idx])
    for m in MASAS_CONSOLA:
        a = analisis.get(m)
        if a is None:
            continue
        txt = f"    Masa {m}: "
        txt += (f"inicio {a['T_on']:.1f} °C (t = {a['t_on']:.2f} min)" if a["t_on"] is not None
                else "no sube durante el calentamiento")
        if a["T_tan"] is not None:
            txt += f" | inicio por tangente {a['T_tan']:.1f} °C"
        if a["t_max"] is not None:
            txt += f" | máximo {a['T_max']:.1f} °C (t = {a['t_max']:.2f} min, {a['P_max']:.3g} mbar)"
        print(txt)

    conv = None
    ref, motivo = elegir_blanco(ruta_txt.stem, df["Time"].iloc[0], refs or {})
    if ref is None:
        print("    Conversión: no hay blanco procesado (sin X_CO2). Pon el .txt del blanco en la carpeta "
              f"con {PALABRAS_BLANCO[0]!r} en el nombre.")
    else:
        a_ch4 = analisis.get(MASA_CH4) or {}
        info["t_max_ch4"], info["T_max_ch4"] = a_ch4.get("t_max"), a_ch4.get("T_max")
        conv = calcular_conversion(df, t_min, T_rga, mascara, info, ref)
        conv["motivo_blanco"] = motivo
        print(f"    Blanco: {ref['archivo']} ({motivo}); r44,0 = {ref['r44_0']:.5g} "
              f"(± {100 * ref['r44_0_err'] / ref['r44_0']:.2f} %)")
        if conv["X_sin_reac"] is not None:
            xs = conv["X_sin_reac"]
            print(f"    Control sin reacción ({VENTANA_SIN_REACCION_T[0]}–{VENTANA_SIN_REACCION_T[1]} °C): "
                  f"X_CO2 aparente {100 * xs:+.1f} %" + (
                      "  <- el catalizador está guardando CO2 (o el blanco no corresponde): la X de la RAMPA no es "
                      "conversión; usa el estado isoterma" if xs > 0.05 else "  (ok)" if xs > -0.05 else
                      "  <- negativa: r44,0 del blanco es BAJO para esta corrida"))
        tm = conv["tmax"]
        if tm is not None:
            print(f"    A T máx del pico de CH4 ({tm['Tmax']:.1f} °C, t = {tm['t'][0]:.1f}–{tm['t'][1]:.1f} min, "
                  f"T {tm['T_rango'][0]:.0f}–{tm['T_rango'][1]:.0f} °C, {tm['n']} scans): "
                  f"X_CO2 = {100 * tm['X']:.1f} ± {100 * tm['X_err_total']:.1f} %")
        e = conv["est"]
        if e is not None:
            print(f"    Final de la isoterma (t = {e['t'][0]:.1f}–{e['t'][1]:.1f} min, {e['T']:.0f} °C, {e['n']} scans): "
                  f"X_CO2 = {100 * e['X']:.1f} ± {100 * e['X_err_total']:.1f} %  "
                  f"(ruido {100 * e['X_err']:.1f}, blanco {100 * e['X_err_blanco']:.1f})")
            if e["XH"] is not None:
                print(f"      X_H2 = {100 * e['XH']:.1f} ± {100 * e['XH_err']:.1f} %   S_CH4 por balance de H2 "
                      f"(R = {R_H2_CO2:g}) = {100 * e['S_H2']:.1f} ± {100 * e['S_H2_err']:.1f} %   "
                      f"r44,0 mínimo compatible = {e['r44_0_min']:.4f} "
                      f"({'ok' if e['r44_0_min'] <= ref['r44_0'] * 1.02 else 'MAYOR que el del blanco -> revisa'})")
            if e["F_C"] is not None:
                print(f"      F_C = {e['F_C']:.3f}  (cota superior de F: si se formó CO, el F real es menor)")
            if conv.get("F") is not None and e.get("Y") is not None:
                print(f"      F = {conv['F']:.3f} ± {conv['F_err']:.3f} ({'calibrado en esta corrida' if F == 'datos' else 'fijo'}"
                      f"; tabla de fragmentación: {conv['F_tabla']:.2f})   Y_CH4 = {100 * e['Y']:.1f} ± "
                      f"{100 * e['Y_err']:.1f} %   S_CH4 = {100 * e['S']:.1f} ± {100 * e['S_err']:.1f} %")
        if e is not None and e.get("deriva"):
            dv = e["deriva"]
            f1 = lambda v: "—" if v is None else f"{v:+.1f}"
            print(f"      Cambio dentro de la ventana ({dv['dur']:.1f} min; isoterma de {dv['iso_dur']:.1f} min): "
                  f"X_CO2 {f1(dv['X'])}, X_H2 {f1(dv['XH'])}, Y_CH4 {f1(dv['Y'])} puntos; T {dv['T']:+.1f} °C"
                  + ("" if e["estacionario"] else "  <- AVISO: NO es estado estacionario (la reacción todavía "
                     "evoluciona): son valores al final de la isoterma"))
        tmx = conv.get("tmax")
        if tmx is not None and tmx.get("Y") is not None:
            print(f"    A T máx del pico de CH4: Y_CH4 = {100 * tmx['Y']:.1f} ± {100 * tmx['Y_err']:.1f} %")
        else:
            print("    Final de la isoterma: no hay isoterma -> sin X_CO2 confiable si el catalizador guarda CO2")
        for nombre_g in graficar_conversion(t_min, T_rga, mascara, mascara_T, conv, info, carpeta / nombre):
            print(f"    [OK] {nombre_g}.png / .svg")

    graficar_perfil_T(r, t_inicio, base, info, df, carpeta / f"{nombre}_perfil_T_reactor")
    print(f"    [OK] {nombre}_perfil_T_reactor.png / .svg")
    exportar_excel(ruta_txt, ruta_reactor, df, r, t_min, T_rga, mascara, mascara_T, info,
                   t_inicio, base, manual, rampa, analisis, carpeta / f"{nombre}_datos_graficos.xlsx", conv)
    print(f"    [OK] {nombre}_datos_graficos.xlsx")
    _, dibujadas = generar_vistas(df, t_min, T_rga, mascara, mascara_T, carpeta, nombre, info)

    excluir, previos = set(), []
    while PREGUNTAR_ELIMINAR:
        print("\n    Curvas graficadas:  " + "  |  ".join(f"{m} {nombre_plano(MASAS[m][0])}" for m in dibujadas))
        if preguntar("    ¿Quieres eliminar alguna curva? (s/n) [n]: ").lower() not in ("s", "si", "sí", "y", "yes"):
            break
        nuevas = pedir_curvas_a_eliminar(dibujadas)
        if not nuevas:
            continue
        excluir |= nuevas
        archivos, dibujadas = generar_vistas(df, t_min, T_rga, mascara, mascara_T, carpeta, nombre, info, excluir)
        for f in previos:
            try:
                f.unlink()
            except OSError:
                pass
        previos = archivos
        print(f"    Curvas eliminadas: {', '.join(map(str, sorted(excluir)))}")
    return fila_resumen(ruta_txt, analisis.get(MASA_CH4) or {}, conv)


def etiqueta_corrida(nombre):
    m = re.search(r"OMNi[\d.]+(?:Cu[\d.]+)?(?:Fe[\d.]+)?", nombre)
    if m:
        comp = m.group().rstrip(".")
        return ETIQUETAS_RESUMEN.get(comp, comp), comp
    return nombre, ""


def fila_resumen(ruta_txt, a15, conv):
    et, comp = etiqueta_corrida(ruta_txt.stem)
    r = {"Etiqueta": et, "Composición": comp, "Archivo": ruta_txt.name,
         "T inicio CH4 (°C)": a15.get("T_on"), "T máx CH4 (°C)": a15.get("T_max")}
    if conv is None:
        return r
    v = lambda d, k: (100 * d[k]) if (d is not None and d.get(k) is not None) else None
    tm, es = conv.get("tmax"), conv.get("est")
    r.update({
        "X_CO2 a T máx (%)": v(tm, "X"), "σ X_CO2 a T máx": v(tm, "X_err_total"),
        "X_H2 a T máx (%)": v(tm, "XH"), "σ X_H2 a T máx": v(tm, "XH_err"),
        "Y_CH4 a T máx (%)": v(tm, "Y"), "σ Y_CH4 a T máx": v(tm, "Y_err"),
        "T isoterma (°C)": es["T"] if es else None,
        "X_CO2 isoterma (%)": v(es, "X"), "σ X_CO2 isoterma": v(es, "X_err_total"),
        "X_H2 isoterma (%)": v(es, "XH"), "σ X_H2 isoterma": v(es, "XH_err"),
        "S_CH4 isoterma (%)": v(es, "S") if (es and es.get("S") is not None) else v(es, "S_H2"),
        "σ S_CH4 isoterma": v(es, "S_err") if (es and es.get("S") is not None) else v(es, "S_H2_err"),
        "S_CH4 sistemático (±)": v(es, "S_sist"),
        "Y_CH4 isoterma (%)": v(es, "Y"), "σ Y_CH4 isoterma": v(es, "Y_err"),
        "F usado": conv.get("F"), "σ F": conv.get("F_err"), "F_C": conv.get("F_C"), "F tabla": conv.get("F_tabla"),
        "X_CO2 aparente sin reacción (%)": (100 * conv["X_sin_reac"]) if conv.get("X_sin_reac") is not None else None,
        "Blanco": conv["ref"]["archivo"], "Por qué este blanco": conv.get("motivo_blanco"),
    })
    return r


def exportar_resumen(filas, salida: Path):
    filas = sorted(filas, key=lambda f: (not str(f["Etiqueta"]).startswith("OM"), str(f["Etiqueta"])))
    sist = [f["S_CH4 sistemático (±)"] for f in filas if f.get("S_CH4 sistemático (±)") is not None]
    for f in filas:
        if f.get("S_CH4 isoterma (%)") is not None and f.get("S_CH4 sistemático (±)") is None and sist:
            f["S_CH4 sistemático (±)"] = max(sist)
            f["S sistemático: origen"] = "no calculable en esta corrida: se usa el mayor de las demás (conservador)"
    Fs = {i: f["F usado"] for i, f in enumerate(filas) if f.get("F usado") is not None}
    if len(Fs) >= 3:
        med = float(np.median(list(Fs.values())))
        validos = {i: v for i, v in Fs.items() if abs(v / med - 1) <= F_ATIPICO}
        for i, f in enumerate(filas):
            f["F atípico (excluido)"] = "sí" if (i in Fs and i not in validos) else ("no" if i in Fs else None)
            otros = [v for j, v in validos.items() if j != i]
            if f.get("F_C") and len(otros) >= 2:
                f["F promedio otras corridas"] = float(np.mean(otros))
                f["S_CH4 por CH4 medido (%)"] = 100 * float(np.mean(otros)) / f["F_C"]
                if f.get("S_CH4 isoterma (%)") is not None:
                    f["Diferencia S (CH4 medido - balance H2)"] = f["S_CH4 por CH4 medido (%)"] - f["S_CH4 isoterma (%)"]
    cols = list(dict.fromkeys(k for f in filas for k in f))
    wb = Workbook()
    ws = wb.active
    ws.title = "Resumen"
    ws.append(cols)
    for f in filas:
        ws.append([round(f.get(c), DECIMALES) if isinstance(f.get(c), float) and not c.startswith(("F", "σ F"))
                   else (round(f[c], 3) if isinstance(f.get(c), float) else f.get(c)) for c in cols])
    _negrita_fila1(ws)
    _anchos(ws, cols, minimo=10, maximo=22)
    ws.freeze_panes = "B2"
    wi = wb.create_sheet("Notas")
    for k, t in [
        ("T inicio / T máx", f"Inicio (base + {N_SIGMA_ONSET}σ sostenido {MIN_SOSTENIDO} min) y máximo de m/z {MASA_CH4}"),
        ("A T máx", f"Promedio de ± {VENTANA_PICO_MIN:g} min alrededor del máximo de CH4. Si el máximo cae en la rampa, "
                    "X_CO2 puede incluir CO2 que el catalizador guarda o libera: compárala con el estado isoterma"),
        ("Final isoterma", f"Últimos {ESTABLE_MIN:g} min de la isoterma (sin perturbaciones del Ar > 10 %)"),
        ("Barras", "Resumen_conversion.png en T máx (máximo de CH4): X_CO2, X_H2 y Y_CH4; S_CH4 no se grafica allí "
                   "porque supera 100 % (el catalizador convierte también carbono acumulado en la rampa)"
                   if BARRAS_EN == "Tmax" else "Resumen_conversion.png al final de la isoterma"),
        ("X_CO2", "1 - r44/r44,0, r44,0 del blanco sin catalizador; σ = ruido y blanco en cuadratura"),
        ("Errores", "Medias con todos los scans; error estándar por bloques de " + str(SUAVIZADO) + " scans "
                    "(GUM 4.2.7, series correlacionadas)" if ERROR_METODO == "bloques" else "σ/√n de los scans crudos"),
        ("σ Y isoterma", "Y = X·S = (R·X_H2 - X_CO2)/3  ->  σ_Y = √(R²σ²_XH2 + σ²_XCO2)/3"),
        ("S sistemático", "Mayor diferencia de S al usar ambas conversiones con la referencia del blanco o ambas con "
                          "la de la propia corrida (solo referencias válidas); se reporta aparte del error estadístico"),
        ("S por CH4 medido", "Validación: S_C = F/F_C con F promedio de las otras corridas (sin F atípicos, > "
                             f"{100 * F_ATIPICO:g} % de la mediana); debe coincidir con S por balance de H2"),
        ("S_CH4", f"Balance de H2 con R = {R_H2_CO2:g}: S = (R·X_H2/X_CO2 - 1)/3"),
        ("Y_CH4", "F·A_CH4/r44,0 (CH4 que sale por CO2 que entra); al final de la isoterma = X_CO2·S_CH4"),
        ("F", "Calibrado en cada corrida = S_CH4 · F_C (F = 'datos') o fijo; F tabla = FF44·S_CO2/(FF15·S_CH4), "
              "solo para comparar (tabla de 70 eV)"),
        ("Sistemático", "Si el blanco es de otro día, X_CO2 puede moverse ~(1 - X)·Δ, con Δ el cambio relativo de la "
                        "respuesta 44/40 del equipo (±5 % -> ±2.5 puntos con X ≈ 50 %)"),
    ]:
        wi.append([k, t])
    wi.column_dimensions["A"].width = 18
    wi.column_dimensions["B"].width = 120
    wb.save(salida / "Resumen_conversion.xlsx")

    if BARRAS_EN == "Tmax":
        con = [f for f in filas if f.get("X_CO2 a T máx (%)") is not None]
        series = [("X_CO2 a T máx (%)", "σ X_CO2 a T máx", r"X$_{CO_2}$", MASAS[MASA_CO2][1]),
                  ("X_H2 a T máx (%)", "σ X_H2 a T máx", r"X$_{H_2}$", MASAS[MASA_H2][1]),
                  ("Y_CH4 a T máx (%)", "σ Y_CH4 a T máx", r"Y$_{CH_4}$", MASAS[MASA_CH4][1])]
        etiquetas = [f"{f['Etiqueta']}\n{f['T máx CH4 (°C)']:.0f} °C" if f.get("T máx CH4 (°C)") else f["Etiqueta"]
                     for f in con]
        ylabel = "Conversión / rendimiento (%)"
    else:
        con = [f for f in filas if f.get("X_CO2 isoterma (%)") is not None]
        series = [("X_CO2 isoterma (%)", "σ X_CO2 isoterma", r"X$_{CO_2}$", MASAS[MASA_CO2][1]),
                  ("X_H2 isoterma (%)", "σ X_H2 isoterma", r"X$_{H_2}$", MASAS[MASA_H2][1]),
                  ("Y_CH4 isoterma (%)", "σ Y_CH4 isoterma", r"Y$_{CH_4}$", MASAS[MASA_CH4][1]),
                  ("S_CH4 isoterma (%)", "σ S_CH4 isoterma", r"S$_{CH_4}$", "#4d4d4d")]
        etiquetas = [f["Etiqueta"] for f in con]
        ylabel = "Conversión / rendimiento / selectividad (%)"
    if not con:
        return
    series = [x for x in series if any(f.get(x[0]) is not None for f in con)]
    fig, ax = plt.subplots(figsize=(max(FIG_W, 1.0 + 0.9 * len(con)), FIG_H))
    fig.subplots_adjust(left=MARGEN_IZQ, right=0.98, bottom=MARGEN_INF, top=1 - MARGEN_SUP)
    x = np.arange(len(con))
    ancho = 0.8 / len(series)
    for j, (k, ke, lab, col) in enumerate(series):
        y = [f.get(k) if f.get(k) is not None else np.nan for f in con]
        ye = [f.get(ke) if f.get(ke) is not None else 0 for f in con]
        ax.bar(x - 0.4 + ancho * (j + 0.5), y, ancho * 0.92, yerr=ye, color=col, label=lab, capsize=1.5,
               error_kw={"elinewidth": 0.6, "capthick": 0.6})
    ax.set_xticks(x)
    ax.set_xticklabels(etiquetas, fontsize=FS_TICK)
    ax.set_ylabel(ylabel, fontsize=FS_LABEL, fontweight="bold", labelpad=4)
    ax.set_ylim(0, 105)
    ax.yaxis.set_major_locator(MultipleLocator(20))
    ax.tick_params(axis="both", labelsize=FS_TICK, direction="in", width=LW_AXES)
    for sp in ax.spines.values():
        sp.set_linewidth(LW_AXES)
    ax.legend(ncol=len(series), loc="upper center", frameon=False, fontsize=FS_CURVA, bbox_to_anchor=(0.5, 1.0))
    for ext in ("png", "svg"):
        fig.savefig(salida / f"Resumen_conversion.{ext}", dpi=DPI, facecolor="white")
    plt.close(fig)


def main():
    entrada = Path(RUTA)
    if entrada.is_file():
        carpeta, txts = entrada.parent, [entrada]
    elif entrada.is_dir():
        carpeta = entrada
        txts = None
    else:
        print(f"La ruta no existe: {entrada}")
        return
    archivos = sorted(p for p in carpeta.iterdir() if p.is_file())
    todos_txt = [p for p in archivos if es_txt_masas(p)]
    blancos = [p for p in todos_txt if es_blanco(p)]
    if txts is None:
        txts = todos_txt
    corridas = [p for p in txts if not es_blanco(p)]
    reactores = {}
    for p in archivos:
        if es_reactor(p):
            try:
                reactores[p] = leer_reactor(p)
            except Exception as e:
                print(f"[AVISO] No se pudo leer el reactor {p.name}: {e}")
    if not todos_txt:
        print("No se encontraron .txt del masas.")
        return

    salida = carpeta / CARPETA_SALIDA if CARPETA_SALIDA else carpeta
    salida.mkdir(exist_ok=True)

    refs = cargar_referencias(salida)
    for txt in blancos:
        try:
            ref = analizar_blanco(txt, salida)
            refs[ref["nombre"]] = ref
        except PermissionError:
            print(f"[ERROR] {txt.name}: cierra el Excel / PNG / SVG si lo tienes abierto")
        except Exception as e:
            print(f"[ERROR] blanco {txt.name}: {e}")
    if refs:
        print("\nBlancos disponibles: " + "  |  ".join(f"{k} ({v['fecha']}, r44,0 = {v['r44_0']:.4g})"
                                                     for k, v in refs.items()))

    if not corridas:
        print(f"\nListo (solo blancos). Resultados en: {salida / CARPETA_BLANCOS}")
        return
    if not reactores:
        print("No se encontraron archivos del reactor (nombre con 'Prueba').")
        return
    print("Archivos del reactor: " + "  |  ".join(p.name for p in reactores))
    ok = 0
    resumen = []
    for txt in corridas:
        try:
            fila = procesar(txt, reactores, salida, refs)
            if fila:
                resumen.append(fila)
            ok += 1
        except PermissionError:
            print(f"[ERROR] {txt.name}: cierra el Excel / PNG / SVG si lo tienes abierto")
        except Exception as e:
            print(f"[ERROR] {txt.name}: {e}")
    if resumen:
        try:
            exportar_resumen(resumen, salida)
            print(f"\n[OK] Resumen_conversion.xlsx y Resumen_conversion.png / .svg (todas las corridas)")
        except PermissionError:
            print("[ERROR] Resumen_conversion: cierra el Excel / PNG si lo tienes abierto")
    print(f"\nListo: {ok}/{len(corridas)} corridas procesadas. Resultados en: {salida}")


if __name__ == "__main__":
    main()
