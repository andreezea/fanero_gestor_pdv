"""
Mi Cartera - Gestores Fanero
================================================
Esqueleto simplificado del dashboard de habilitadores: cada GESTOR (DNI +
Nombre) tiene uno o más PUNTOS DE VENTA (PDV) bajo su cargo, cada uno con su
propia Cuota y Avance por Producto.

IMPORTANTE - esto es un ESQUELETO intencionalmente simplificado:
    Se retiraron Coordinadores, Habilitadores y la distinción
    Activadores/Desarrolladores que existían en la versión anterior. Esos
    conceptos se reincorporarán después en una pestaña aparte, sin romper
    esta base. Todo Gestor tiene PDV debajo (ya no es opcional / condicional
    a un tipo de habilitador).

Estructura del archivo:
    1. Configuración y constantes
    2. Funciones de datos (carga / generación / cálculo / publicación)
    3. Funciones de presentación (KPIs, formato semáforo, tablas dinámicas)
    4. Panel de administrador (acceso restringido)
    5. Plantilla Excel descargable
    6. Edición de avances por gestor (acceso restringido)
    7. Interfaz principal (main): Mi Cartera (Gestor) + Vista Gerencial

Accesos ocultos (se mantienen igual que en la versión anterior):
    ?admin=1   → panel de administrador: publica el archivo Excel completo
                 (con día de corte, mes y año).
    ?editar=1  → pestaña "Editar Avances": cada gestor actualiza el avance
                 de sus propios PDV, sin tocar los de otros gestores.

Lógica de proyección (igual que la versión anterior; Cuota y Avance son
unidades, no montos en dinero):
    Proy Unidades = Avance * (días del mes / día de corte)
    Proy %        = Proy Unidades / Cuota
    Días restantes = días del mes - día de corte
    Cuota diaria necesaria = (Cuota - Avance) / Días restantes

Carga de datos del administrador:
    El administrador sube un Excel con Cuota y Avance (acumulado del mes
    hasta la fecha de corte declarada), indica Mes/Año/Día de corte, y
    publica. Cada publicación REEMPLAZA por completo lo que se veía antes.

Listo para desplegar en Streamlit Cloud: `streamlit run app.py`
"""

import calendar
import json
import os
from datetime import datetime
from io import BytesIO

import numpy as np
import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.worksheet.datavalidation import DataValidation

# =============================================================================
# 1. CONFIGURACIÓN Y CONSTANTES
# =============================================================================

st.set_page_config(
    page_title="Mi Cartera - Gestores Fanero",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

DATA_DIR = "data"
DATA_FILE = os.path.join(DATA_DIR, "ultima_carga.xlsx")
DATA_META = os.path.join(DATA_DIR, "meta.json")
LOG_EDICION = os.path.join(DATA_DIR, "ultima_edicion.json")

# Geografía de referencia: Departamento → Provincia → Distritos. Se conserva
# tal cual de la versión anterior (misma fuente para plantilla y ejemplo).
GEOGRAFIA = {
    "Amazonas": {
        "Chachapoyas": ["Chachapoyas", "Huancas", "Levanto"],
        "Bagua": ["Bagua", "Aramango", "Copallín"],
        "Condorcanqui": ["Nieva", "Río Santiago", "El Cenepa"],
    },
    "Cajamarca": {
        "Cajamarca": ["Cajamarca", "Baños del Inca", "Jesús"],
        "Jaén": ["Jaén", "Bellavista", "Chontalí"],
        "Cutervo": ["Cutervo", "Callayuc", "Choros"],
    },
    "Huancavelica": {
        "Huancavelica": ["Huancavelica", "Acoria", "Yauli"],
        "Tayacaja": ["Pampas", "Acraquia", "Colcabamba"],
        "Acobamba": ["Acobamba", "Andabamba", "Anta"],
    },
    "Huánuco": {
        "Huánuco": ["Huánuco", "Amarilis", "Pillco Marca"],
        "Leoncio Prado": ["Rupa-Rupa", "Daniel Alomía Robles", "Hermilio Valdizán"],
        "Ambo": ["Ambo", "Cayna", "Colpas"],
    },
    "Junín": {
        "Huancayo": ["Huancayo", "Chilca", "El Tambo"],
        "Tarma": ["Tarma", "Acobamba", "Huaricolca"],
        "Chanchamayo": ["Chanchamayo", "Perené", "Pichanaqui"],
    },
    "Loreto": {
        "Maynas": ["Iquitos", "Belén", "Punchana"],
        "Alto Amazonas": ["Yurimaguas", "Balsapuerto", "Jeberos"],
        "Requena": ["Requena", "Maquía", "Tapiche"],
    },
    "Pasco": {
        "Pasco": ["Chaupimarca", "Yanacancha", "Simón Bolívar"],
        "Oxapampa": ["Oxapampa", "Pozuzo", "Villa Rica"],
        "Daniel Alcides Carrión": ["Yanahuanca", "Chacayán", "Goyllarisquizga"],
    },
    "San Martín": {
        "Moyobamba": ["Moyobamba", "Calzada", "Habana"],
        "San Martín": ["Tarapoto", "Morales", "La Banda de Shilcayo"],
        "Rioja": ["Rioja", "Elías Soplín Vargas", "Nueva Cajamarca"],
    },
    "Ucayali": {
        "Coronel Portillo": ["Callería", "Yarinacocha", "Manantay"],
        "Padre Abad": ["Padre Abad", "Irazola", "Curimaná"],
        "Atalaya": ["Raymondi", "Sepahua", "Tahuanía"],
    },
}

DEPARTAMENTOS = list(GEOGRAFIA.keys())
TODAS_LAS_PROVINCIAS = sorted({prov for deps in GEOGRAFIA.values() for prov in deps.keys()})

# Productos analizados (el orden aquí define el orden de las columnas)
PRODUCTOS = ["Prepago", "Porta Flex", "Postpago", "OSS"]

# Nombres de referencia para el dataset sintético
NOMBRES_EJEMPLO = [
    "Carlos Ramírez", "María Torres", "Jorge Quispe", "Ana Flores",
    "Luis Mamani", "Rosa Huamán", "Pedro Vargas", "Karen Chávez",
    "Miguel Salazar", "Diana Rojas", "José Cárdenas", "Lucía Paredes",
    "Fernando Ríos", "Patricia Gómez", "Andrés Castillo", "Silvia Cruz",
    "Raúl Medina", "Carmen Delgado", "Víctor Herrera", "Elena Campos",
]
NOMBRES_PDV_EJEMPLO = [
    "Ana Ruiz", "Luis Gómez", "Marco Díaz", "Katia Del Águila",
    "Elmer Torres", "Milagros Chávez", "Sandro Pinedo", "Rosa Vela",
]

# Columnas obligatorias en el Excel. A diferencia de la versión anterior,
# PDV y Nombre PDV YA NO son opcionales: todo Gestor tiene al menos un PDV.
COLUMNAS_REQUERIDAS = {
    "DNI", "Nombre", "Departamento", "Provincia", "Distrito",
    "PDV", "Nombre PDV", "Producto", "Cuota", "Avance",
}


# =============================================================================
# 2. FUNCIONES DE DATOS
# =============================================================================

def _normalizar_texto(df: pd.DataFrame, columna: str, relleno: str = "") -> pd.DataFrame:
    """Garantiza que `columna` exista y quede como texto limpio. Sirve de red
    de seguridad para archivos publicados antes de algún cambio de esquema."""
    df = df.copy()
    if columna not in df.columns:
        df[columna] = relleno
    df[columna] = df[columna].fillna(relleno).astype(str).str.strip()
    return df


def _normalizar_identidad(df: pd.DataFrame) -> pd.DataFrame:
    """Aplica la normalización de texto a todas las columnas de identidad."""
    df = df.copy()
    for columna in ["DNI", "Nombre", "PDV", "Nombre PDV", "Departamento", "Provincia", "Distrito"]:
        relleno = "Sin dato" if columna in ("Departamento", "Provincia", "Distrito") else ""
        df = _normalizar_texto(df, columna, relleno)
    return df


@st.cache_data
def generar_datos_ejemplo(n_gestores: int = 40, seed: int = 42) -> pd.DataFrame:
    """Genera un dataset sintético: cada Gestor tiene entre 1 y 3 PDV, y cada
    PDV tiene una fila por Producto con su propia Cuota/Avance."""
    rng = np.random.default_rng(seed)
    registros = []

    for i in range(n_gestores):
        dni_gestor = str(40000000 + int(rng.integers(0, 9_999_999)))
        nombre_gestor = rng.choice(NOMBRES_EJEMPLO)
        departamento = rng.choice(DEPARTAMENTOS)
        provincia = rng.choice(list(GEOGRAFIA[departamento].keys()))
        distrito = rng.choice(GEOGRAFIA[departamento][provincia])

        n_pdv = int(rng.integers(1, 4))
        for _ in range(n_pdv):
            dni_pdv = str(70000000 + int(rng.integers(0, 9_999_999)))
            nombre_pdv = rng.choice(NOMBRES_PDV_EJEMPLO)

            for producto in PRODUCTOS:
                cuota_pdv = int(rng.integers(30, 200))
                factor_avance = rng.uniform(0.3, 1.3)
                avance_pdv = int(round(cuota_pdv * factor_avance))
                registros.append({
                    "DNI": dni_gestor,
                    "Nombre": nombre_gestor,
                    "Departamento": departamento,
                    "Provincia": provincia,
                    "Distrito": distrito,
                    "PDV": dni_pdv,
                    "Nombre PDV": nombre_pdv,
                    "Producto": producto,
                    "Cuota": cuota_pdv,
                    "Avance": avance_pdv,
                })

    return pd.DataFrame(registros)


def cargar_datos_excel(archivo) -> pd.DataFrame | None:
    """Lee y valida un archivo Excel cargado por el administrador.
    Retorna None (y muestra un error en la UI) si faltan columnas requeridas."""
    try:
        df = pd.read_excel(archivo)
    except Exception as exc:  # noqa: BLE001 - se informa al usuario cualquier error de lectura
        st.error(f"No se pudo leer el archivo Excel: {exc}")
        return None

    faltantes = COLUMNAS_REQUERIDAS - set(df.columns)
    if faltantes:
        st.error("El archivo no contiene las columnas requeridas: " + ", ".join(sorted(faltantes)))
        return None

    return _normalizar_identidad(df)


@st.cache_data
def _leer_excel_publicado(path: str, mtime: float) -> pd.DataFrame:
    """Lee el Excel publicado. `mtime` forma parte de la clave de cache: si
    cambia el archivo, el cache se invalida automáticamente."""
    df = pd.read_excel(path)
    return _normalizar_identidad(df)


def obtener_datos_publicados() -> tuple[pd.DataFrame, int, int, int]:
    """Devuelve (datos, día de corte, mes, año) — lo último publicado por el
    administrador, o el dataset de ejemplo si aún no se publicó nada."""
    ahora = datetime.now()

    if os.path.exists(DATA_FILE):
        df = _leer_excel_publicado(DATA_FILE, os.path.getmtime(DATA_FILE))
        try:
            with open(DATA_META, "r", encoding="utf-8") as f:
                meta = json.load(f)
            dia_corte = int(meta.get("dia_corte", max(ahora.day - 1, 1)))
            mes = int(meta.get("mes", ahora.month))
            anio = int(meta.get("anio", ahora.year))
        except Exception:  # noqa: BLE001 - metadatos no disponibles o corruptos
            dia_corte, mes, anio = max(ahora.day - 1, 1), ahora.month, ahora.year
        return df, dia_corte, mes, anio

    dia_corte = max(ahora.day - 1, 1)
    df_ejemplo = _normalizar_identidad(generar_datos_ejemplo())
    return df_ejemplo, dia_corte, ahora.month, ahora.year


def publicar_datos(df: pd.DataFrame, dia_corte: int, mes: int, anio: int) -> None:
    """Guarda el archivo validado y sus metadatos como la fuente de datos
    oficial del dashboard. Reemplaza por completo lo publicado anteriormente."""
    os.makedirs(DATA_DIR, exist_ok=True)
    df = _normalizar_identidad(df)
    df.to_excel(DATA_FILE, index=False)
    with open(DATA_META, "w", encoding="utf-8") as f:
        json.dump({"dia_corte": dia_corte, "mes": mes, "anio": anio}, f)
    _leer_excel_publicado.clear()  # invalida el cache de lectura


def calcular_metricas(df: pd.DataFrame, dias_en_mes: int, dia_corte: int) -> pd.DataFrame:
    """Calcula las columnas derivadas del análisis (Cuota/Avance son
    unidades, no montos en dinero). Se calculan a nivel de fila (PDV × Producto):

    - Cumplimiento % = Avance / Cuota
    - Proy Unidades  = Avance * (días del mes / día de corte)
    - Proy %         = Proy Unidades / Cuota
    """
    df = df.copy()
    dia_corte = max(dia_corte, 1)  # evita división entre cero

    df["Cumplimiento %"] = np.where(df["Cuota"] > 0, df["Avance"] / df["Cuota"], 0.0)

    factor_proyeccion = dias_en_mes / dia_corte
    df["Proy Unidades"] = df["Avance"] * factor_proyeccion
    df["Proy %"] = np.where(df["Cuota"] > 0, df["Proy Unidades"] / df["Cuota"], 0.0)

    return df


# =============================================================================
# 3. FUNCIONES DE PRESENTACIÓN (FORMATO SEMÁFORO Y TABLAS DINÁMICAS)
# =============================================================================

def color_semaforo(valor: float) -> str:
    """Devuelve el estilo CSS de fondo según el cumplimiento (semáforo)."""
    if pd.isna(valor):
        return ""
    if valor < 0.80:
        color = "#f8d7da"  # rojo
    elif valor < 1.00:
        color = "#fff3cd"  # amarillo
    else:
        color = "#d4edda"  # verde
    return f"background-color: {color}; color: #1a1a1a"


def _aplicar_semaforo(styler, columnas: list):
    """Aplica color_semaforo a las columnas indicadas (nombres simples o
    tuplas, para tablas con columnas MultiIndex)."""
    if hasattr(styler, "map"):
        for col in columnas:
            styler = styler.map(color_semaforo, subset=[col])
    else:  # pragma: no cover - fallback para pandas antiguo
        for col in columnas:
            styler = styler.applymap(color_semaforo, subset=[col])
    return styler


def resumen_por_producto(df_filtrado: pd.DataFrame, departamentos_sel: list,
                          productos_sel: list, dias_en_mes: int, dia_corte: int) -> pd.DataFrame:
    """Resumen por Departamento con los productos como columnas agrupadas
    (Cuota, Avance, Cumplimiento %, Proy Unidades, Proy % debajo de cada uno).
    Se usa en Vista Gerencial. Agrega TODOS los gestores y PDV de cada
    departamento."""
    largo = (
        df_filtrado.groupby(["Departamento", "Producto"], as_index=False)
        .agg(Cuota=("Cuota", "sum"), Avance=("Avance", "sum"))
    )

    orden_dep = [d for d in DEPARTAMENTOS if d in departamentos_sel]
    orden_prod = [p for p in PRODUCTOS if p in productos_sel]
    combinaciones = pd.MultiIndex.from_product([orden_dep, orden_prod], names=["Departamento", "Producto"])
    largo = largo.set_index(["Departamento", "Producto"]).reindex(combinaciones).reset_index()
    largo[["Cuota", "Avance"]] = largo[["Cuota", "Avance"]].fillna(0)

    largo["Cumplimiento %"] = np.where(largo["Cuota"] > 0, largo["Avance"] / largo["Cuota"], 0.0)
    factor_proyeccion = dias_en_mes / max(dia_corte, 1)
    largo["Proy Unidades"] = largo["Avance"] * factor_proyeccion
    largo["Proy %"] = np.where(largo["Cuota"] > 0, largo["Proy Unidades"] / largo["Cuota"], 0.0)

    metricas = ["Cuota", "Avance", "Cumplimiento %", "Proy Unidades", "Proy %"]
    ancho = largo.pivot_table(index="Departamento", columns="Producto", values=metricas, aggfunc="first")
    ancho = ancho.swaplevel(axis=1)
    columnas_orden = pd.MultiIndex.from_product([orden_prod, metricas])
    ancho = ancho.reindex(index=orden_dep, columns=columnas_orden)

    return ancho


def aplicar_estilo_resumen_producto(tabla: pd.DataFrame, orden_prod: list):
    fmt = {}
    for p in orden_prod:
        fmt[(p, "Cuota")] = "{:,.0f}"
        fmt[(p, "Avance")] = "{:,.0f}"
        fmt[(p, "Cumplimiento %")] = "{:.1%}"
        fmt[(p, "Proy Unidades")] = "{:,.0f}"
        fmt[(p, "Proy %")] = "{:.1%}"

    styler = tabla.style.format(fmt, na_rep="-")
    subset = [(p, "Cumplimiento %") for p in orden_prod] + [(p, "Proy %") for p in orden_prod]
    return _aplicar_semaforo(styler, subset)


def ranking_gestores(df_filtrado: pd.DataFrame) -> pd.DataFrame:
    """Ranking de gestores: suma Cuota/Avance de todos sus PDV y productos,
    con cantidad de PDV y Cumplimiento % agregado. Se usa en Vista Gerencial."""
    ranking = (
        df_filtrado.groupby(["DNI", "Nombre"], as_index=False)
        .agg(
            PDV=("PDV", "nunique"),
            Cuota=("Cuota", "sum"),
            Avance=("Avance", "sum"),
        )
    )
    ranking["Cumplimiento %"] = np.where(ranking["Cuota"] > 0, ranking["Avance"] / ranking["Cuota"], 0.0)
    return ranking.sort_values("Avance", ascending=False).reset_index(drop=True)


def aplicar_estilo_ranking(tabla: pd.DataFrame):
    fmt = {"Cuota": "{:,.0f}", "Avance": "{:,.0f}", "Cumplimiento %": "{:.1%}"}
    styler = tabla.style.format(fmt)
    return _aplicar_semaforo(styler, ["Cumplimiento %"])


def detalle_pdv_gestor(df_filtrado: pd.DataFrame, dias_en_mes: int, dia_corte: int) -> pd.DataFrame:
    """Detalle a nivel PDV (una fila por Producto + PDV) con sus propias
    métricas, sin agregar. Es la base para armar la tabla de un gestor."""
    detalle = df_filtrado[
        ["DNI", "Nombre", "Departamento", "Provincia", "Distrito", "Producto", "PDV", "Nombre PDV", "Cuota", "Avance"]
    ].copy()
    detalle["Cumplimiento %"] = np.where(detalle["Cuota"] > 0, detalle["Avance"] / detalle["Cuota"], 0.0)
    factor_proyeccion = dias_en_mes / max(dia_corte, 1)
    detalle["Proy Unidades"] = detalle["Avance"] * factor_proyeccion
    return detalle


def tabla_detalle_gestor(detalle_g: pd.DataFrame, productos_sel: list) -> pd.DataFrame:
    """A partir del detalle de PDV de UN gestor (ya filtrado a su DNI), arma
    una sola tabla: primera fila "Total" (suma de todos sus PDV), luego una
    fila por cada PDV. Los productos van como columnas agrupadas, con Cuota,
    Avance y Cumplimiento % debajo de cada uno."""
    identidad = ["DNI", "Departamento", "Provincia", "Distrito"]

    total = (
        detalle_g.groupby(identidad + ["Producto"], as_index=False)
        .agg(Cuota=("Cuota", "sum"), Avance=("Avance", "sum"))
    )
    total["PDV"] = "Total"
    total["Nombre PDV"] = ""
    total["Cumplimiento %"] = np.where(total["Cuota"] > 0, total["Avance"] / total["Cuota"], 0.0)

    columnas_combinar = [
        "DNI", "PDV", "Nombre PDV", "Departamento", "Provincia", "Distrito",
        "Producto", "Cuota", "Avance", "Cumplimiento %",
    ]
    combinado = pd.concat([total[columnas_combinar], detalle_g[columnas_combinar]], ignore_index=True)

    orden_prod = [p for p in PRODUCTOS if p in productos_sel]
    metricas = ["Cuota", "Avance", "Cumplimiento %"]

    ancho = combinado.pivot_table(
        index=["DNI", "PDV", "Nombre PDV", "Departamento", "Provincia", "Distrito"],
        columns="Producto", values=metricas, aggfunc="first",
    )
    ancho = ancho.swaplevel(axis=1)
    columnas_orden = pd.MultiIndex.from_product([orden_prod, metricas])
    ancho = ancho.reindex(columns=columnas_orden)

    fila_total = ancho[ancho.index.get_level_values("PDV") == "Total"]
    filas_pdv = ancho[ancho.index.get_level_values("PDV") != "Total"].sort_index(level="PDV")
    return pd.concat([fila_total, filas_pdv])


def aplicar_estilo_detalle_pdv(tabla: pd.DataFrame, orden_prod: list):
    fmt = {}
    for p in orden_prod:
        fmt[(p, "Cuota")] = "{:,.0f}"
        fmt[(p, "Avance")] = "{:,.0f}"
        fmt[(p, "Cumplimiento %")] = "{:.1%}"

    styler = tabla.style.format(fmt, na_rep="-")
    subset = [(p, "Cumplimiento %") for p in orden_prod]
    return _aplicar_semaforo(styler, subset)


def ritmo_pdv_gestor(detalle_g: pd.DataFrame, productos_sel: list, dias_restantes: int) -> pd.DataFrame:
    """Igual que tabla_detalle_gestor pero con métricas de ritmo diario: fila
    "Total" (suma de sus PDV) y una fila por cada PDV, con Cuota Diaria,
    Corte y Cump % debajo de cada producto."""
    identidad = ["DNI", "Departamento", "Provincia", "Distrito"]

    total = (
        detalle_g.groupby(identidad + ["Producto"], as_index=False)
        .agg(Cuota=("Cuota", "sum"), Avance=("Avance", "sum"))
    )
    total["PDV"] = "Total"
    total["Nombre PDV"] = ""

    columnas_base = identidad + ["Producto", "PDV", "Nombre PDV", "Cuota", "Avance"]
    filas_pdv = detalle_g[columnas_base].copy()
    combinado = pd.concat([total[columnas_base], filas_pdv], ignore_index=True)

    combinado["Cump %"] = np.where(combinado["Cuota"] > 0, combinado["Avance"] / combinado["Cuota"], 0.0)
    combinado["Corte"] = combinado["Avance"]
    if dias_restantes > 0:
        combinado["Cuota Diaria"] = (combinado["Cuota"] - combinado["Avance"]) / dias_restantes
    else:
        combinado["Cuota Diaria"] = np.nan

    orden_prod = [p for p in PRODUCTOS if p in productos_sel]
    metricas = ["Cuota Diaria", "Corte", "Cump %"]

    ancho = combinado.pivot_table(
        index=["DNI", "PDV", "Nombre PDV", "Departamento", "Provincia", "Distrito"],
        columns="Producto", values=metricas, aggfunc="first",
    )
    ancho = ancho.swaplevel(axis=1)
    columnas_orden = pd.MultiIndex.from_product([orden_prod, metricas])
    ancho = ancho.reindex(columns=columnas_orden)

    fila_total = ancho[ancho.index.get_level_values("PDV") == "Total"]
    filas_pdv_orden = ancho[ancho.index.get_level_values("PDV") != "Total"].sort_index(level="PDV")
    return pd.concat([fila_total, filas_pdv_orden])


def aplicar_estilo_ritmo_gestor(tabla: pd.DataFrame, orden_prod: list):
    fmt = {}
    for p in orden_prod:
        fmt[(p, "Cuota Diaria")] = "{:,.1f}"
        fmt[(p, "Corte")] = "{:,.0f}"
        fmt[(p, "Cump %")] = "{:.1%}"

    styler = tabla.style.format(fmt, na_rep="-")
    subset = [(p, "Cump %") for p in orden_prod]
    return _aplicar_semaforo(styler, subset)


def ritmo_diario_detalle_pdv(df_filtrado: pd.DataFrame, dias_restantes: int) -> pd.DataFrame:
    """Detalle a nivel PDV (una fila por Producto + PDV) con Cuota Diaria,
    Corte y Cump %, sin agregar. Se usa para la descarga CSV."""
    detalle = df_filtrado[
        ["DNI", "Nombre", "Departamento", "Provincia", "Distrito", "Producto", "PDV", "Nombre PDV", "Cuota", "Avance"]
    ].copy()
    detalle["Corte"] = detalle["Avance"]
    detalle["Cump %"] = np.where(detalle["Cuota"] > 0, detalle["Avance"] / detalle["Cuota"], 0.0)
    if dias_restantes > 0:
        detalle["Cuota Diaria"] = (detalle["Cuota"] - detalle["Avance"]) / dias_restantes
    else:
        detalle["Cuota Diaria"] = np.nan
    return detalle


# =============================================================================
# 4. PANEL DE ADMINISTRADOR (ACCESO RESTRINGIDO)
# =============================================================================
#
# Configuración de credenciales (recomendado, no se sube al repositorio):
# En Streamlit Cloud → Settings → Secrets, agregar:
#
#   [admin]
#   usuario = "admin"
#   password = "coloca_aqui_una_clave_segura"

def _credenciales_admin() -> tuple[str, str]:
    try:
        return st.secrets["admin"]["usuario"], st.secrets["admin"]["password"]
    except Exception:  # noqa: BLE001 - no hay secrets configurados aún
        return "admin", "admin2025"


def panel_admin() -> None:
    """Renderiza el control de acceso y la carga de datos. Solo se llama
    cuando la URL incluye ?admin=1."""
    st.header("🔒 Panel administrador")

    if not st.session_state.get("es_admin", False):
        with st.form("form_login_admin"):
            usuario = st.text_input("Usuario")
            clave = st.text_input("Contraseña", type="password")
            enviar = st.form_submit_button("Ingresar")

        if enviar:
            usuario_ok, clave_ok = _credenciales_admin()
            if usuario == usuario_ok and clave == clave_ok:
                st.session_state["es_admin"] = True
                st.rerun()
            else:
                st.error("Usuario o contraseña incorrectos.")
        return

    st.success("Sesión de administrador activa.")

    st.download_button(
        "📥 Plantilla de carga",
        data=generar_plantilla_excel(),
        file_name="plantilla_carga_gestores.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    st.caption(
        "Cuota y Avance ACUMULADO de cada PDV. Cada vez que subas un archivo, "
        "reemplaza por completo lo publicado -no se suma nada-, así que Avance "
        "debe ser siempre el total acumulado real hasta la fecha de corte que declares."
    )
    st.caption(
        "Elimina las filas de ejemplo antes de subir tu archivo real. Cada PDV "
        "debe tener una fila por producto (Prepago, Porta Flex, Postpago, OSS). "
        "PDV = DNI del líder de ese punto de venta; Nombre PDV = su nombre. "
        "DNI y Nombre (columnas iniciales) identifican al Gestor dueño del PDV."
    )

    with st.expander("Columnas del Excel"):
        st.write("Todas obligatorias:", sorted(COLUMNAS_REQUERIDAS))

    ahora = datetime.now()
    col_mes, col_anio = st.columns(2)
    with col_mes:
        mes_sel = st.number_input("Mes de la cuota", min_value=1, max_value=12, value=ahora.month)
    with col_anio:
        anio_sel = st.number_input("Año", min_value=2020, max_value=2100, value=ahora.year, step=1)

    dias_en_mes_sel = calendar.monthrange(int(anio_sel), int(mes_sel))[1]
    dia_corte_defecto = min(max(ahora.day - 1, 1), dias_en_mes_sel)
    dia_corte_sel = st.number_input(
        "Día de corte (último día del mes con ventas cargadas)",
        min_value=1, max_value=dias_en_mes_sel, value=dia_corte_defecto,
    )

    archivo = st.file_uploader("Cargar archivo Excel (.xlsx)", type=["xlsx"])

    if archivo is not None and st.button("Publicar datos"):
        df_validado = cargar_datos_excel(archivo)
        if df_validado is not None:
            publicar_datos(df_validado, int(dia_corte_sel), int(mes_sel), int(anio_sel))
            st.success("Datos publicados. Todos los usuarios verán la actualización al recargar.")

    if os.path.exists(DATA_FILE):
        ultima_actualizacion = datetime.fromtimestamp(os.path.getmtime(DATA_FILE))
        st.caption(f"Última publicación: {ultima_actualizacion:%d/%m/%Y %H:%M}")

    if st.button("Cerrar sesión"):
        st.session_state["es_admin"] = False
        st.rerun()


# =============================================================================
# 5. PLANTILLA EXCEL DESCARGABLE
# =============================================================================

def generar_plantilla_excel() -> bytes:
    """Genera en memoria la plantilla de carga (para el admin) con los
    encabezados requeridos, filas de ejemplo y listas desplegables
    (validación de datos) para Departamento y Provincia."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Datos"

    headers = [
        "DNI", "Nombre", "Departamento", "Provincia", "Distrito",
        "PDV", "Nombre PDV", "Producto", "Cuota", "Avance",
    ]
    ws.append(headers)

    ejemplos = [
        # Mismo gestor (DNI 87654321), 2 PDV distintos, mismo producto
        ["87654321", "Rosa Huamán", "San Martín", "San Martín", "Tarapoto", "71234567", "Ana Ruiz", "Prepago", 180, 150],
        ["87654321", "Rosa Huamán", "San Martín", "San Martín", "Tarapoto", "76543210", "Luis Gómez", "Prepago", 120, 95],
        ["87654321", "Rosa Huamán", "San Martín", "San Martín", "Tarapoto", "71234567", "Ana Ruiz", "Postpago", 60, 30],
    ]
    for fila in ejemplos:
        ws.append(fila)

    ultima_fila = 1000
    dv_departamento = DataValidation(type="list", formula1='"' + ",".join(DEPARTAMENTOS) + '"', allow_blank=True)
    dv_provincia = DataValidation(type="list", formula1='"' + ",".join(TODAS_LAS_PROVINCIAS) + '"', allow_blank=True)
    dv_producto = DataValidation(type="list", formula1='"' + ",".join(PRODUCTOS) + '"', allow_blank=True)

    ws.add_data_validation(dv_departamento)
    ws.add_data_validation(dv_provincia)
    ws.add_data_validation(dv_producto)

    dv_departamento.add(f"C2:C{ultima_fila}")
    dv_provincia.add(f"D2:D{ultima_fila}")
    dv_producto.add(f"H2:H{ultima_fila}")

    anchos = {"A": 12, "B": 22, "C": 16, "D": 18, "E": 18, "F": 12, "G": 18, "H": 12, "I": 10, "J": 10}
    for col, ancho in anchos.items():
        ws.column_dimensions[col].width = ancho

    buffer = BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


def dataframe_a_excel_bytes(df: pd.DataFrame, hoja: str = "Avances") -> bytes:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=hoja)
    return buffer.getvalue()


# =============================================================================
# 6. EDICIÓN DE AVANCES POR GESTOR (ACCESO RESTRINGIDO)
# =============================================================================
#
# Pestaña oculta: solo aparece con ?editar=1 en la URL. Cada gestor elige su
# nombre/DNI en una lista desplegable y solo puede editar el Avance de sus
# PROPIOS PDV. La selección no lleva contraseña: es un control de confianza
# para uso interno, no una autenticación real.

def registrar_ultima_edicion(nombre: str) -> None:
    os.makedirs(DATA_DIR, exist_ok=True)
    with open(LOG_EDICION, "w", encoding="utf-8") as f:
        json.dump({"nombre": nombre, "timestamp": datetime.now().isoformat()}, f)


def obtener_ultima_edicion() -> dict | None:
    if os.path.exists(LOG_EDICION):
        try:
            with open(LOG_EDICION, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:  # noqa: BLE001 - archivo corrupto o incompleto
            return None
    return None


def actualizar_avances(cambios: pd.DataFrame, gestor: str) -> None:
    """Aplica los nuevos valores de Avance (columnas DNI, PDV, Producto,
    Avance) sobre el dataset publicado completo y vuelve a publicar,
    preservando el día de corte, mes y año vigentes."""
    df_actual, dia_corte, mes, anio = obtener_datos_publicados()
    df_actual = _normalizar_identidad(df_actual)

    cambios = cambios.copy()
    cambios = _normalizar_identidad(cambios)

    df_actual = df_actual.set_index(["DNI", "PDV", "Producto"])
    cambios_idx = cambios.set_index(["DNI", "PDV", "Producto"])["Avance"]
    indices_validos = cambios_idx.index.intersection(df_actual.index)
    df_actual.loc[indices_validos, "Avance"] = cambios_idx.loc[indices_validos]
    df_actual = df_actual.reset_index()

    publicar_datos(df_actual, dia_corte, mes, anio)
    registrar_ultima_edicion(gestor)


def panel_editar_avances(df_raw: pd.DataFrame) -> None:
    """Pestaña de edición de avances. `df_raw` es el dataset publicado
    completo (todos los gestores), para que cada gestor pueda encontrarse en
    la lista y editar solo sus propios PDV."""
    ultima = obtener_ultima_edicion()
    if ultima:
        ts = datetime.fromisoformat(ultima["timestamp"])
        st.caption(f"Última conexión: {ultima['nombre']} · {ts:%d/%m/%Y %H:%M}")
    else:
        st.caption("Todavía no se registró ninguna edición de avances.")

    gestores = (
        df_raw[["DNI", "Nombre"]].drop_duplicates().sort_values("Nombre")
    )
    opciones = [f"{fila['Nombre']} · DNI {fila['DNI']}" for _, fila in gestores.iterrows()]
    mapa_opcion_dni = dict(zip(opciones, gestores["DNI"]))

    if not opciones:
        st.info("No hay gestores disponibles en los datos publicados.")
        return

    seleccion = st.selectbox("Gestor que actualizará su avance", opciones, key="gestor_editor")
    dni_sel = mapa_opcion_dni[seleccion]
    nombre_sel = gestores[gestores["DNI"] == dni_sel]["Nombre"].iloc[0]

    df_editable = (
        df_raw[df_raw["DNI"] == dni_sel]
        [["DNI", "Nombre", "Departamento", "Provincia", "Distrito", "PDV", "Nombre PDV", "Producto", "Cuota", "Avance"]]
        .sort_values(["Producto", "PDV"])
        .reset_index(drop=True)
    )
    df_editable = _normalizar_identidad(df_editable)

    if df_editable.empty:
        st.info("No hay registros para este gestor.")
        return

    claves_permitidas = set(zip(df_editable["DNI"], df_editable["PDV"], df_editable["Producto"]))

    st.markdown("#### Opción 1 · Editar directamente en la tabla")
    columnas_bloqueadas = [c for c in df_editable.columns if c != "Avance"]
    editado = st.data_editor(
        df_editable,
        disabled=columnas_bloqueadas,
        hide_index=True,
        width="stretch",
        height=350,
        key=f"editor_avances_{dni_sel}",
    )

    if st.button("Guardar cambios de la tabla", key="guardar_avances_tabla"):
        actualizar_avances(editado[["DNI", "PDV", "Producto", "Avance"]], nombre_sel)
        st.success(f"Avances actualizados por {nombre_sel}.")
        st.rerun()

    st.markdown("---")
    st.markdown("#### Opción 2 · Descargar plantilla, editar en Excel y volver a subir")

    st.download_button(
        "📥 Descargar mis datos actuales (Excel)",
        data=dataframe_a_excel_bytes(df_editable),
        file_name=f"avances_{nombre_sel.replace(' ', '_').lower()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="descargar_plantilla_avances",
    )
    st.caption("Solo modifica la columna Avance. No cambies DNI, PDV, Nombre PDV ni Producto.")

    archivo_avances = st.file_uploader(
        "Subir Excel con avances actualizados", type=["xlsx"], key=f"subir_avances_{dni_sel}"
    )

    if archivo_avances is not None and st.button("Guardar cambios del archivo", key="guardar_avances_archivo"):
        try:
            df_subido = pd.read_excel(archivo_avances)
        except Exception as exc:  # noqa: BLE001 - se informa al usuario cualquier error de lectura
            st.error(f"No se pudo leer el archivo: {exc}")
            df_subido = None

        if df_subido is not None:
            faltantes = {"DNI", "PDV", "Producto", "Avance"} - set(df_subido.columns)
            if faltantes:
                st.error("Al archivo le faltan columnas: " + ", ".join(sorted(faltantes)))
            else:
                df_subido = _normalizar_identidad(df_subido)
                mascara_permitida = df_subido.apply(
                    lambda fila: (fila["DNI"], fila["PDV"], fila["Producto"]) in claves_permitidas, axis=1
                )
                df_filtrado_subida = df_subido[mascara_permitida]
                ignoradas = len(df_subido) - len(df_filtrado_subida)

                if df_filtrado_subida.empty:
                    st.error(f"Ninguna fila del archivo corresponde a los PDV de {nombre_sel}.")
                else:
                    actualizar_avances(df_filtrado_subida[["DNI", "PDV", "Producto", "Avance"]], nombre_sel)
                    if ignoradas > 0:
                        st.warning(f"Se ignoraron {ignoradas} fila(s) que no pertenecen a este gestor.")
                    st.success(f"Avances actualizados por {nombre_sel} desde archivo.")
                    st.rerun()


# =============================================================================
# 7. INTERFAZ PRINCIPAL
# =============================================================================

def vista_gestor(df: pd.DataFrame, dias_en_mes: int, dia_corte: int, dias_restantes: int) -> None:
    """Pestaña 'Mi Cartera': el Gestor se busca a sí mismo y ve solo sus PDV."""
    gestores = df[["DNI", "Nombre"]].drop_duplicates().sort_values("Nombre")
    opciones = [f"{fila['Nombre']} · DNI {fila['DNI']}" for _, fila in gestores.iterrows()]
    mapa_opcion_dni = dict(zip(opciones, gestores["DNI"]))

    if not opciones:
        st.info("No hay gestores disponibles en los datos publicados.")
        return

    seleccion = st.selectbox("Selecciona tu nombre (Gestor)", opciones, key="gestor_seleccionado")
    dni_sel = mapa_opcion_dni[seleccion]

    productos_sel = st.multiselect("Producto", options=PRODUCTOS, default=PRODUCTOS, key="productos_gestor")
    if not productos_sel:
        productos_sel = PRODUCTOS
    orden_prod_sel = [p for p in PRODUCTOS if p in productos_sel]

    df_gestor = df[(df["DNI"] == dni_sel) & (df["Producto"].isin(productos_sel))]
    if df_gestor.empty:
        st.info("No hay datos para los productos seleccionados.")
        return

    # --- KPIs del gestor ---
    cuota_total = df_gestor["Cuota"].sum()
    avance_total = df_gestor["Avance"].sum()
    cumplimiento = (avance_total / cuota_total) if cuota_total > 0 else 0.0
    proy_total = avance_total * (dias_en_mes / max(dia_corte, 1))
    proy_pct = (proy_total / cuota_total) if cuota_total > 0 else 0.0
    n_pdv = df_gestor["PDV"].nunique()

    col1, col2, col3, col4, col5 = st.columns(5)
    col1.metric("PDV a cargo", f"{n_pdv}")
    col2.metric("Cuota total", f"{cuota_total:,.0f}")
    col3.metric("Avance", f"{avance_total:,.0f}", f"{cumplimiento:.1%}")
    col4.metric("Proyección", f"{proy_total:,.0f}", f"{proy_pct:.1%}")
    col5.metric("Días restantes", f"{dias_restantes}")

    st.markdown("---")
    st.markdown("#### Detalle por PDV")

    detalle_pdv = detalle_pdv_gestor(df_gestor, dias_en_mes, dia_corte)
    tabla_g = tabla_detalle_gestor(detalle_pdv, productos_sel)

    es_total = tabla_g.index.get_level_values("PDV") == "Total"
    st.markdown("**Total de mi cartera**")
    st.dataframe(aplicar_estilo_detalle_pdv(tabla_g[es_total], orden_prod_sel), width="stretch")
    st.markdown(f"**Mis {n_pdv} PDV**")
    st.dataframe(aplicar_estilo_detalle_pdv(tabla_g[~es_total], orden_prod_sel), width="stretch")

    st.markdown("---")
    st.markdown("#### Ritmo diario necesario")
    st.caption(f"Quedan {dias_restantes} día(s) para el cierre del mes (día {dia_corte} → día {dias_en_mes}).")
    if dias_restantes == 0:
        st.warning("El mes ya cerró; no quedan días para calcular el ritmo diario.")

    tabla_ritmo = ritmo_pdv_gestor(detalle_pdv, productos_sel, dias_restantes)
    es_total_r = tabla_ritmo.index.get_level_values("PDV") == "Total"
    st.dataframe(aplicar_estilo_ritmo_gestor(tabla_ritmo[es_total_r], orden_prod_sel), width="stretch")
    with st.expander(f"➕ Ver ritmo diario por PDV ({n_pdv} punto(s) de venta)"):
        st.dataframe(aplicar_estilo_ritmo_gestor(tabla_ritmo[~es_total_r], orden_prod_sel), width="stretch")

    st.markdown("---")
    columnas_csv = [
        "DNI", "Nombre", "Departamento", "Provincia", "Distrito", "Producto", "PDV", "Nombre PDV",
        "Cuota", "Avance", "Cumplimiento %", "Proy Unidades",
    ]
    st.download_button(
        "⬇️ Descargar mi cartera (CSV)",
        data=detalle_pdv[columnas_csv].sort_values(["Producto", "PDV"]).to_csv(index=False).encode("utf-8"),
        file_name=f"mi_cartera_{dni_sel}.csv",
        mime="text/csv",
    )
    st.caption("🟥 <80% · 🟨 80%–99% · 🟩 ≥100%")


def vista_gerencial(df: pd.DataFrame, dias_en_mes: int, dia_corte: int) -> None:
    """Pestaña 'Vista Gerencial': todos los gestores, sin filtro individual."""
    col_dep, col_prod = st.columns(2)
    with col_dep:
        departamentos_sel = st.multiselect(
            "Departamento", options=sorted(df["Departamento"].unique()),
            default=sorted(df["Departamento"].unique()), key="depto_gerencial",
        )
    with col_prod:
        productos_sel = st.multiselect(
            "Producto", options=PRODUCTOS, default=PRODUCTOS, key="producto_gerencial",
        )

    if not departamentos_sel:
        departamentos_sel = sorted(df["Departamento"].unique())
    if not productos_sel:
        productos_sel = PRODUCTOS

    df_filtrado = df[df["Departamento"].isin(departamentos_sel) & df["Producto"].isin(productos_sel)]

    if df_filtrado.empty:
        st.info("No hay datos para los filtros seleccionados.")
        return

    # --- KPIs generales ---
    cuota_total = df_filtrado["Cuota"].sum()
    avance_total = df_filtrado["Avance"].sum()
    cumplimiento = (avance_total / cuota_total) if cuota_total > 0 else 0.0
    proy_total = avance_total * (dias_en_mes / max(dia_corte, 1))
    proy_pct = (proy_total / cuota_total) if cuota_total > 0 else 0.0

    col1, col2, col3, col4, col5 = st.columns(5)
    col1.metric("Gestores", f"{df_filtrado['DNI'].nunique():,}")
    col2.metric("PDV", f"{df_filtrado['PDV'].nunique():,}")
    col3.metric("Cuota total", f"{cuota_total:,.0f}")
    col4.metric("Avance", f"{avance_total:,.0f}", f"{cumplimiento:.1%}")
    col5.metric("Proyección", f"{proy_total:,.0f}", f"{proy_pct:.1%}")

    st.markdown("---")
    st.markdown("#### Resumen por Producto (por Departamento)")
    tabla_resumen = resumen_por_producto(df_filtrado, departamentos_sel, productos_sel, dias_en_mes, dia_corte)
    orden_prod_sel = [p for p in PRODUCTOS if p in productos_sel]
    st.dataframe(aplicar_estilo_resumen_producto(tabla_resumen, orden_prod_sel), width="stretch")
    st.caption("🟥 <80% · 🟨 80%–99% · 🟩 ≥100% (aplica a Cumplimiento % y Proy %)")

    st.markdown("---")
    st.markdown("#### 🏆 Ranking de gestores")
    ranking = ranking_gestores(df_filtrado)
    st.dataframe(aplicar_estilo_ranking(ranking), width="stretch", height=420)

    st.download_button(
        "⬇️ Descargar ranking (CSV)",
        data=ranking.to_csv(index=False).encode("utf-8"),
        file_name="ranking_gestores.csv",
        mime="text/csv",
    )


def main():
    st.title("📊 Mi Cartera - Gestores Fanero")

    # Los paneles ocultos solo se renderizan con sus parámetros en la URL.
    if st.query_params.get("admin") == "1":
        with st.sidebar:
            panel_admin()
    mostrar_editor = st.query_params.get("editar") == "1"

    df_raw, dia_corte, mes, anio = obtener_datos_publicados()
    dias_en_mes = calendar.monthrange(anio, mes)[1]
    dias_restantes = max(dias_en_mes - dia_corte, 0)

    st.caption(
        "Seguimiento de PDV por gestor · "
        f"Datos al día {dia_corte} de {dias_en_mes} ({mes:02d}/{anio})"
    )

    df_raw = df_raw[df_raw["Departamento"].isin(DEPARTAMENTOS)]
    df_raw = df_raw[df_raw["Producto"].isin(PRODUCTOS)]

    if df_raw.empty:
        st.warning("No hay datos disponibles para los departamentos/productos configurados.")
        return

    df = calcular_metricas(df_raw, dias_en_mes, dia_corte)

    tab_names = ["👤 Mi Cartera (Gestor)", "🏢 Vista Gerencial"]
    if mostrar_editor:
        tab_names.append("✏️ Editar Avances")
    tabs = st.tabs(tab_names)

    with tabs[0]:
        vista_gestor(df, dias_en_mes, dia_corte, dias_restantes)

    with tabs[1]:
        vista_gerencial(df, dias_en_mes, dia_corte)

    if mostrar_editor:
        with tabs[2]:
            st.subheader("Editar Avances")
            panel_editar_avances(df_raw)


if __name__ == "__main__":
    main()
