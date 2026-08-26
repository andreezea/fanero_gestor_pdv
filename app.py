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

Acceso (un solo link, sin parámetros en la URL):
    La app pide un único login (usuario + contraseña) configurado en
    Secrets. Según qué credencial coincida, la sesión queda como:
    - [admin]        → además del dashboard, ve el panel para publicar datos
                       (barra lateral).
    - [visualizacion]→ solo ve el dashboard (Mi Cartera, Vista Gerencial,
                       Editar Avances).
    Ver `_credenciales_admin` / `_credenciales_visualizacion` para configurar
    las claves reales en Streamlit Cloud → Settings → Secrets.

Lógica de proyección (igual que la versión anterior; Cuota y Avance son
unidades, no montos en dinero):
    Proy Unidades = Avance * (días del mes / día de corte)
    Proy %        = Proy Unidades / Cuota
    Días restantes = días del mes - día de corte
    Cuota diaria necesaria = (Cuota - Avance) / Días restantes

Carga de datos del administrador:
    El administrador sube un Excel con el Avance del día/periodo más
    reciente; la app lo SUMA automáticamente al acumulado del mes (ver
    `publicar_datos_incremental`). Al cambiar de Mes/Año, el acumulado se
    reinicia solo y el mes saliente se archiva para poder comparar
    M0 vs M-1 en Vista Gerencial.

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
    # "auto": Streamlit la expande sola si detecta contenido en la sidebar
    # (el panel admin, una vez logueado). Antes estaba "collapsed" a
    # propósito para ocultar el acceso admin vía URL secreta — ya no aplica,
    # ahora el acceso admin depende del login, no de que la sidebar esté oculta.
    initial_sidebar_state="auto",
)

DATA_DIR = "data"
DATA_FILE = os.path.join(DATA_DIR, "ultima_carga.xlsx")
DATA_META = os.path.join(DATA_DIR, "meta.json")
LOG_EDICION = os.path.join(DATA_DIR, "ultima_edicion.json")
HISTORICO_DIR = os.path.join(DATA_DIR, "historico")

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
PRODUCTOS = ["Prepago", "Porta Prepago", "Postpago", "OSS"]

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

# ---------------------------------------------------------------------------
# REGLAS DE COMISIÓN (editar aquí cuando Fanero defina las reglas reales)
# ---------------------------------------------------------------------------
# Comisión estimada según el % de Cumplimiento (o Proyección) del gestor.
# CumplimientoMaximo = None significa "sin límite superior" (el nivel más alto).
# Estos valores son un placeholder razonable — AJUSTAR con las reglas reales.
REGLAS_COMISION = [
    {"CumplimientoMinimo": 0.00, "CumplimientoMaximo": 0.70, "ComisionSoles": 0.0},
    {"CumplimientoMinimo": 0.70, "CumplimientoMaximo": 0.90, "ComisionSoles": 300.0},
    {"CumplimientoMinimo": 0.90, "CumplimientoMaximo": 1.00, "ComisionSoles": 600.0},
    {"CumplimientoMinimo": 1.00, "CumplimientoMaximo": 1.20, "ComisionSoles": 1000.0},
    {"CumplimientoMinimo": 1.20, "CumplimientoMaximo": None, "ComisionSoles": 1500.0},
]


def calcular_comision_estimada(cumplimiento_pct: float) -> float:
    """Devuelve la comisión estimada (S/) según el % de cumplimiento/proyección,
    usando la tabla de niveles REGLAS_COMISION."""
    comision = 0.0
    for nivel in REGLAS_COMISION:
        minimo = nivel["CumplimientoMinimo"]
        maximo = nivel["CumplimientoMaximo"]
        dentro = cumplimiento_pct >= minimo and (maximo is None or cumplimiento_pct < maximo)
        if dentro or cumplimiento_pct >= minimo:
            comision = nivel["ComisionSoles"]
    return comision


# ---------------------------------------------------------------------------
# RANGOS DE CUMPLIMIENTO (filtro "Rango de Cumplimiento" en Vista Gerencial)
# ---------------------------------------------------------------------------
# El rango se calcula sobre el % de PROYECCIÓN (Proy Unidades / Cuota) de UN
# producto específico — no sobre el total de todos los productos:
#   - Si están seleccionados TODOS los productos, o solo "Prepago" → se usa
#     Prepago, y aparece el rango especial "PDV no activo (0 ventas)".
#   - Si se elige un único producto distinto de Prepago (Porta Prepago,
#     Postpago, OSS) → se usa ese producto, y el rango más bajo es
#     "0.00% – 80%" (sin el rango especial de "no activo").
RANGOS_CON_NO_ACTIVO = [
    "PDV no activo (0 ventas)",
    "0.01% – 80%",
    "80.01% – 90%",
    "90.01% – 95%",
    "95.01% – 100%",
    "Más de 100%",
]
RANGOS_SIN_NO_ACTIVO = [
    "0.00% – 80%",
    "80.01% – 90%",
    "90.01% – 95%",
    "95.01% – 100%",
    "Más de 100%",
]


def producto_base_para_rango(productos_sel: list) -> str:
    """Determina sobre qué Producto se calcula el filtro de Rango de
    Cumplimiento, según lo que esté seleccionado en el filtro Producto."""
    if "Prepago" in productos_sel:
        return "Prepago"
    if len(productos_sel) == 1:
        return productos_sel[0]
    for p in PRODUCTOS:
        if p in productos_sel:
            return p
    return "Prepago"


def clasificar_rango_proyeccion(proy_pct: float, incluir_no_activo: bool) -> str:
    """Clasifica un % de Proyección (fracción, ej 0.85) en un rango.
    Si incluir_no_activo=True (producto base = Prepago), un valor <= 0 cae
    en 'PDV no activo (0 ventas)'; si no, el rango más bajo es '0.00% – 80%'
    e incluye también los valores en 0."""
    if incluir_no_activo and proy_pct <= 0:
        return "PDV no activo (0 ventas)"
    if proy_pct <= 0.80:
        return "0.01% – 80%" if incluir_no_activo else "0.00% – 80%"
    if proy_pct <= 0.90:
        return "80.01% – 90%"
    if proy_pct <= 0.95:
        return "90.01% – 95%"
    if proy_pct <= 1.00:
        return "95.01% – 100%"
    return "Más de 100%"


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
    oficial del dashboard. Reemplaza por completo lo publicado anteriormente.
    Es la función de bajo nivel — para la carga normal del admin usar
    `publicar_datos_incremental`, que sí acumula día a día."""
    os.makedirs(DATA_DIR, exist_ok=True)
    df = _normalizar_identidad(df)
    df.to_excel(DATA_FILE, index=False)
    with open(DATA_META, "w", encoding="utf-8") as f:
        json.dump({"dia_corte": dia_corte, "mes": mes, "anio": anio}, f)
    _leer_excel_publicado.clear()  # invalida el cache de lectura


def _archivar_mes(df: pd.DataFrame, mes: int, anio: int) -> None:
    """Guarda una copia final del mes saliente en data/historico/AAAA_MM.xlsx,
    para poder comparar M0 (mes actual) vs M-1 (mes anterior) más adelante."""
    os.makedirs(HISTORICO_DIR, exist_ok=True)
    ruta = os.path.join(HISTORICO_DIR, f"{anio}_{mes:02d}.xlsx")
    _normalizar_identidad(df).to_excel(ruta, index=False)


def obtener_historico_mes(mes: int, anio: int) -> pd.DataFrame | None:
    """Devuelve los datos archivados de un Mes/Año específico, o None si no
    existe (por ejemplo, el primer mes usando la app, antes de tener historial)."""
    ruta = os.path.join(HISTORICO_DIR, f"{anio}_{mes:02d}.xlsx")
    if os.path.exists(ruta):
        return _normalizar_identidad(pd.read_excel(ruta))
    return None


MESES_ES = {
    "enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6,
    "julio": 7, "agosto": 8, "setiembre": 9, "septiembre": 9, "octubre": 10,
    "noviembre": 11, "diciembre": 12,
}


def _nombre_producto_normalizado(texto: str) -> str | None:
    """Empareja un nombre de columna (ej. 'Porta prepago', 'oss') con el
    Producto oficial correspondiente de PRODUCTOS, sin importar mayúsculas."""
    limpio = str(texto).strip().lower()
    for p in PRODUCTOS:
        if p.lower() == limpio:
            return p
    return None


def procesar_carga_historico_ancho(df_ancho: pd.DataFrame, anio: int, df_referencia: pd.DataFrame) -> dict[int, pd.DataFrame]:
    """Convierte una tabla ANCHA de ventas de meses anteriores (una columna
    por producto) al formato interno que usa la app, agrupando por mes si el
    archivo trae más de uno. Formato esperado:

        DNI          | Mes    | Prepago | Porta Prepago | Postpago | OSS
        11111111     | julio  | 29      | 20            | 10       | 26

    Devuelve {mes_numero: DataFrame_listo_para_archivar}. Nombre/Departamento/
    Provincia/Distrito se completan buscando el DNI en `df_referencia` (los
    datos ya publicados) — si un DNI no aparece ahí, queda con esos campos
    vacíos, pero su venta igual se cuenta en los totales generales.
    """
    if "DNI" not in df_ancho.columns or "Mes" not in df_ancho.columns:
        raise ValueError("El archivo debe tener al menos las columnas 'DNI' y 'Mes'.")

    df_ancho = df_ancho.copy()
    df_ancho["DNI"] = df_ancho["DNI"].astype(str).str.strip()
    df_ancho["_MesTexto"] = df_ancho["Mes"].astype(str).str.strip().str.lower()
    df_ancho["_MesNumero"] = df_ancho["_MesTexto"].map(MESES_ES)

    mapa_col_producto = {}
    for col in df_ancho.columns:
        if col in ("DNI", "Mes", "_MesTexto", "_MesNumero"):
            continue
        producto = _nombre_producto_normalizado(col)
        if producto:
            mapa_col_producto[col] = producto

    if not mapa_col_producto:
        raise ValueError(
            "No se reconoció ninguna columna de producto. Deben llamarse "
            "exactamente: " + ", ".join(PRODUCTOS)
        )

    df_ancho = df_ancho.dropna(subset=["_MesNumero"])
    if df_ancho.empty:
        raise ValueError("No se reconoció ningún mes válido en la columna 'Mes'.")

    ref_cols = [c for c in ["DNI", "Nombre", "Departamento", "Provincia", "Distrito"] if c in df_referencia.columns]
    ref = df_referencia[ref_cols].drop_duplicates(subset=["DNI"]) if "DNI" in ref_cols else pd.DataFrame(columns=["DNI"])

    resultados: dict[int, pd.DataFrame] = {}
    for mes_num, grupo_mes in df_ancho.groupby("_MesNumero"):
        filas = []
        for _, fila in grupo_mes.iterrows():
            for col_original, producto in mapa_col_producto.items():
                valor = fila[col_original]
                if pd.isna(valor):
                    continue
                filas.append({"DNI": fila["DNI"], "Producto": producto, "Avance": float(valor)})
        df_largo = pd.DataFrame(filas)
        if df_largo.empty:
            continue
        df_largo = df_largo.merge(ref, on="DNI", how="left")
        resultados[int(mes_num)] = _normalizar_identidad(df_largo)

    return resultados


def publicar_datos_incremental(df_nuevo: pd.DataFrame, dia_corte: int, mes: int, anio: int) -> None:
    """Publica una carga SUMANDO el Avance nuevo al acumulado ya publicado
    del mismo Mes/Año (en vez de reemplazarlo). Así el administrador puede
    subir solo lo vendido en el día/periodo más reciente, y la app se
    encarga de mantener el acumulado del mes.

    Reglas:
    - DNI + PDV + Producto ya existentes este mes → Avance se SUMA;
      Cuota se actualiza con el valor del archivo nuevo (no se suma, es la
      meta del mes, no una venta).
    - DNI + PDV + Producto nuevos (PDV que se agrega a mitad de mes) → se
      agregan con su Avance tal como vienen en el archivo.
    - Si el Mes/Año de la carga es distinto al que había publicado (mes
      nuevo) → el mes saliente se ARCHIVA automáticamente (para M-1) y se
      empieza de cero, igual que `publicar_datos`.
    """
    df_nuevo = _normalizar_identidad(df_nuevo)

    hay_publicacion_previa = os.path.exists(DATA_FILE)
    mismo_periodo = False
    if hay_publicacion_previa:
        df_actual, _, mes_actual, anio_actual = obtener_datos_publicados()
        mismo_periodo = (int(mes_actual) == int(mes)) and (int(anio_actual) == int(anio))
        if not mismo_periodo:
            _archivar_mes(df_actual, mes_actual, anio_actual)

    if not mismo_periodo:
        # Primera carga del mes (o primera carga de todas): no hay nada que sumar.
        publicar_datos(df_nuevo, dia_corte, mes, anio)
        return

    claves = ["DNI", "PDV", "Producto"]
    df_actual = _normalizar_identidad(df_actual)
    base = df_actual.set_index(claves)
    nuevo = df_nuevo.set_index(claves)

    claves_comunes = base.index.intersection(nuevo.index)
    base.loc[claves_comunes, "Avance"] = (
        base.loc[claves_comunes, "Avance"] + nuevo.loc[claves_comunes, "Avance"]
    )
    if "Cuota" in nuevo.columns:
        base.loc[claves_comunes, "Cuota"] = nuevo.loc[claves_comunes, "Cuota"]

    claves_nuevas = nuevo.index.difference(base.index)
    filas_nuevas = nuevo.loc[claves_nuevas]

    combinado = pd.concat([base, filas_nuevas]).reset_index()
    publicar_datos(combinado, dia_corte, mes, anio)


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


def construir_tabla_producto(
    df_filtrado: pd.DataFrame,
    agrupar_por: str,
    desagrupar: bool,
    productos_sel: list,
    dias_en_mes: int,
    dia_corte: int,
) -> pd.DataFrame:
    """Construye la tabla 'Resumen por Producto' (Producto como columnas
    agrupadas: Cuota, Avance, Proy Unidades, Proy %), con dos variantes
    intercambiables:

    - agrupar_por="Departamento" → una fila por Departamento (o por
      Departamento + PDV si desagrupar=True).
    - agrupar_por="Gestor" → una fila por Gestor (o por Gestor + PDV si
      desagrupar=True).

    Cuando NO se desagrupa, se agrega una fila "Fanero (Total)" al inicio
    con la suma de TODO lo que esté en df_filtrado (todos los departamentos
    o todos los gestores, según corresponda) — sus % se recalculan sobre
    los totales, no se promedian filas.
    """
    df_filtrado = df_filtrado.copy()
    df_filtrado["_Gestor"] = df_filtrado["Nombre"] + " · DNI " + df_filtrado["DNI"]
    df_filtrado["_PDV"] = np.where(
        df_filtrado["Nombre PDV"] != "",
        df_filtrado["PDV"] + " · " + df_filtrado["Nombre PDV"],
        df_filtrado["PDV"],
    )

    nivel_col = "Departamento" if agrupar_por == "Departamento" else "_Gestor"
    index_cols = [nivel_col] + (["_PDV"] if desagrupar else [])

    largo = (
        df_filtrado.groupby(index_cols + ["Producto"], as_index=False)
        .agg(Cuota=("Cuota", "sum"), Avance=("Avance", "sum"))
    )

    orden_prod = [p for p in PRODUCTOS if p in productos_sel]
    factor_proyeccion = dias_en_mes / max(dia_corte, 1)

    if not desagrupar:
        if agrupar_por == "Departamento":
            orden_nivel = [d for d in DEPARTAMENTOS if d in df_filtrado[nivel_col].unique()]
        else:
            orden_nivel = sorted(df_filtrado[nivel_col].unique())
        combinaciones = pd.MultiIndex.from_product([orden_nivel, orden_prod], names=[nivel_col, "Producto"])
        largo = largo.set_index([nivel_col, "Producto"]).reindex(combinaciones).reset_index()
        largo[["Cuota", "Avance"]] = largo[["Cuota", "Avance"]].fillna(0)

    largo["Proy Unidades"] = largo["Avance"] * factor_proyeccion
    largo["Proy %"] = np.where(largo["Cuota"] > 0, largo["Proy Unidades"] / largo["Cuota"], 0.0)

    metricas = ["Cuota", "Avance", "Proy Unidades", "Proy %"]
    ancho = largo.pivot_table(index=index_cols, columns="Producto", values=metricas, aggfunc="first")
    ancho = ancho.swaplevel(axis=1)
    columnas_orden = pd.MultiIndex.from_product([orden_prod, metricas])

    if not desagrupar:
        ancho = ancho.reindex(index=orden_nivel, columns=columnas_orden)

        # --- Fila "Fanero (Total)": suma real de Cuota/Avance/Proy Unidades
        # de TODAS las filas, con Proy % recalculado sobre esos totales
        # (no es el promedio de los % de cada fila). ---
        fila_total = {}
        for p in orden_prod:
            cuota_p = ancho[(p, "Cuota")].sum() if (p, "Cuota") in ancho.columns else 0.0
            avance_p = ancho[(p, "Avance")].sum() if (p, "Avance") in ancho.columns else 0.0
            proy_p = ancho[(p, "Proy Unidades")].sum() if (p, "Proy Unidades") in ancho.columns else 0.0
            fila_total[(p, "Cuota")] = cuota_p
            fila_total[(p, "Avance")] = avance_p
            fila_total[(p, "Proy Unidades")] = proy_p
            fila_total[(p, "Proy %")] = (proy_p / cuota_p) if cuota_p > 0 else 0.0
        df_total = pd.DataFrame([fila_total], index=["Fanero (Total)"], columns=columnas_orden)
        ancho = pd.concat([df_total, ancho])
    else:
        ancho = ancho.reindex(columns=columnas_orden).sort_index(level=0)

    nombre_nivel = "Departamento" if agrupar_por == "Departamento" else "Gestor"
    if desagrupar:
        ancho.index = ancho.index.set_names([nombre_nivel, "PDV"])
    else:
        ancho.index = ancho.index.set_names([nombre_nivel])

    return ancho


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
        fmt[(p, "Proy Unidades")] = "{:,.0f}"
        fmt[(p, "Proy %")] = "{:.1%}"

    tiene_comparativo = ("M0 vs M-1", "M0") in tabla.columns
    if tiene_comparativo:
        fmt[("M0 vs M-1", "M0")] = "{:,.0f}"
        fmt[("M0 vs M-1", "M-1")] = "{:,.0f}"
        fmt[("M0 vs M-1", "%Var")] = "{:+.1%}"

    styler = tabla.style.format(fmt, na_rep="-")
    subset = [(p, "Proy %") for p in orden_prod]
    styler = _aplicar_semaforo(styler, subset)

    if tiene_comparativo:
        def _color_var(v):
            if pd.isna(v):
                return ""
            return "color: #3E9B4F; font-weight: 600" if v >= 0 else "color: #D64545; font-weight: 600"
        styler = styler.map(_color_var, subset=[("M0 vs M-1", "%Var")])

    # Centra todo el contenido de la tabla (números y encabezados).
    styler = styler.set_properties(**{"text-align": "center"})
    styler = styler.set_table_styles([{"selector": "th", "props": [("text-align", "center")]}], overwrite=False)

    # Resalta la fila "Fanero (Total)" en negrita, si está presente.
    if "Fanero (Total)" in tabla.index:
        def _negrita_total(fila):
            return ["font-weight: 700" if fila.name == "Fanero (Total)" else "" for _ in fila]
        styler = styler.apply(_negrita_total, axis=1)

    return styler


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
    ranking = ranking.rename(columns={"DNI": "DNI Gestor", "Nombre": "Nombre Gestor"})
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


def resumen_producto_gestor(df_gestor: pd.DataFrame, productos_sel: list,
                             dias_en_mes: int, dia_corte: int) -> pd.DataFrame:
    """Vista PRINCIPAL de un gestor: una fila por Producto, agregando TODOS
    sus PDV (Cuota, Avance, Cumplimiento %, Proy Unidades, Proy %). Es el
    resumen que se ve primero; el detalle por PDV queda como desagregado
    opcional (ver tabla_detalle_gestor / ritmo_pdv_gestor)."""
    agregado = (
        df_gestor.groupby("Producto", as_index=False)
        .agg(Cuota=("Cuota", "sum"), Avance=("Avance", "sum"))
    )
    orden_prod = [p for p in PRODUCTOS if p in productos_sel]
    agregado = agregado.set_index("Producto").reindex(orden_prod).reset_index()
    agregado[["Cuota", "Avance"]] = agregado[["Cuota", "Avance"]].fillna(0)

    agregado["Cumplimiento %"] = np.where(agregado["Cuota"] > 0, agregado["Avance"] / agregado["Cuota"], 0.0)
    factor_proyeccion = dias_en_mes / max(dia_corte, 1)
    agregado["Proy Unidades"] = agregado["Avance"] * factor_proyeccion
    agregado["Proy %"] = np.where(agregado["Cuota"] > 0, agregado["Proy Unidades"] / agregado["Cuota"], 0.0)

    return agregado.set_index("Producto")


def aplicar_estilo_resumen_gestor(tabla: pd.DataFrame):
    fmt = {"Cuota": "{:,.0f}", "Avance": "{:,.0f}", "Cumplimiento %": "{:.1%}",
           "Proy Unidades": "{:,.0f}", "Proy %": "{:.1%}"}
    styler = tabla.style.format(fmt, na_rep="-")
    return _aplicar_semaforo(styler, ["Cumplimiento %", "Proy %"])


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


def ritmo_producto_gestor(df_gestor: pd.DataFrame, productos_sel: list, dias_restantes: int) -> pd.DataFrame:
    """Ritmo diario necesario agregado por Producto (todos los PDV juntos) —
    vista principal; el desagregado por PDV va en ritmo_pdv_gestor."""
    agregado = (
        df_gestor.groupby("Producto", as_index=False)
        .agg(Cuota=("Cuota", "sum"), Avance=("Avance", "sum"))
    )
    orden_prod = [p for p in PRODUCTOS if p in productos_sel]
    agregado = agregado.set_index("Producto").reindex(orden_prod).reset_index()
    agregado[["Cuota", "Avance"]] = agregado[["Cuota", "Avance"]].fillna(0)

    agregado["Cump %"] = np.where(agregado["Cuota"] > 0, agregado["Avance"] / agregado["Cuota"], 0.0)
    agregado["Corte"] = agregado["Avance"]
    if dias_restantes > 0:
        agregado["Cuota Diaria"] = (agregado["Cuota"] - agregado["Avance"]) / dias_restantes
    else:
        agregado["Cuota Diaria"] = np.nan

    return agregado.set_index("Producto")[["Cuota Diaria", "Corte", "Cump %"]]


def aplicar_estilo_ritmo_resumen(tabla: pd.DataFrame):
    fmt = {"Cuota Diaria": "{:,.1f}", "Corte": "{:,.0f}", "Cump %": "{:.1%}"}
    styler = tabla.style.format(fmt, na_rep="-")
    return _aplicar_semaforo(styler, ["Cump %"])


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
    """Contenido del panel administrador (subir/publicar datos). Se llama
    SOLO cuando ya se inició sesión como admin desde el login general — no
    tiene su propio formulario de acceso ni botón de cerrar sesión."""
    st.header("🔒 Panel administrador")
    st.success("Sesión de administrador activa.")

    st.download_button(
        "📥 Plantilla de carga",
        data=generar_plantilla_excel(),
        file_name="plantilla_carga_gestores.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    st.info(
        "📌 **Cómo funciona la carga:** el Avance que subas se **SUMA automáticamente** "
        "al acumulado que ya está publicado este mes (no lo reemplaza). Sube solo lo "
        "vendido en el día o periodo más reciente — la app se encarga de acumularlo. "
        "Cuota se actualiza con el valor del archivo (no se suma, es la meta del mes). "
        "Al cambiar de Mes/Año, la app empieza el acumulado de cero automáticamente."
    )
    st.caption(
        "Elimina las filas de ejemplo antes de subir tu archivo real. Cada PDV "
        "debe tener una fila por producto (Prepago, Porta Prepago, Postpago, OSS). "
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
        "Día de corte (hasta qué día del mes llega el Avance de esta carga)",
        min_value=1, max_value=dias_en_mes_sel, value=dia_corte_defecto,
    )

    archivo = st.file_uploader("Cargar archivo Excel (.xlsx)", type=["xlsx"])

    forzar_reemplazo = st.checkbox(
        "⚠️ Esta carga REEMPLAZA todo lo publicado, en vez de sumarlo "
        "(úsalo solo para corregir un error de carga)."
    )

    if archivo is not None and st.button("Publicar datos"):
        df_validado = cargar_datos_excel(archivo)
        if df_validado is not None:
            if forzar_reemplazo:
                publicar_datos(df_validado, int(dia_corte_sel), int(mes_sel), int(anio_sel))
                st.success("Datos publicados: se REEMPLAZÓ todo lo anterior con este archivo.")
            else:
                publicar_datos_incremental(df_validado, int(dia_corte_sel), int(mes_sel), int(anio_sel))
                st.success(
                    "Datos publicados: el Avance se sumó al acumulado del mes. "
                    "Todos los usuarios verán la actualización al recargar."
                )

    if os.path.exists(DATA_FILE):
        ultima_actualizacion = datetime.fromtimestamp(os.path.getmtime(DATA_FILE))
        st.caption(f"Última publicación: {ultima_actualizacion:%d/%m/%Y %H:%M}")

    st.markdown("---")
    st.markdown("#### 📂 Cargar ventas de un mes anterior (para M-1)")
    st.caption(
        "Sirve para tener M-1 disponible de inmediato, sin esperar a que la app archive "
        "un mes completo por sí sola. Sube una tabla con una fila por Gestor (DNI), el "
        "Mes en texto, y una columna por producto con el total vendido ese mes:"
    )
    st.code("DNI | Mes | Prepago | Porta Prepago | Postpago | OSS", language=None)
    st.caption(
        "Los nombres de columna deben coincidir exactamente con los productos "
        f"({', '.join(PRODUCTOS)}). Si el archivo trae varios meses en la columna "
        "'Mes', se guardan todos por separado."
    )

    anio_historico = st.number_input(
        "Año de este archivo histórico", min_value=2020, max_value=2100,
        value=datetime.now().year, step=1, key="anio_historico_carga",
    )
    archivo_historico = st.file_uploader(
        "Excel de ventas de mes(es) anterior(es) (formato ancho)", type=["xlsx"], key="uploader_historico_ancho",
    )

    if archivo_historico is not None and st.button("Guardar como histórico"):
        try:
            df_ancho = pd.read_excel(archivo_historico)
        except Exception as exc:  # noqa: BLE001 - se informa al usuario cualquier error de lectura
            st.error(f"No se pudo leer el archivo: {exc}")
            df_ancho = None

        if df_ancho is not None:
            df_referencia, _, _, _ = obtener_datos_publicados()
            try:
                resultados = procesar_carga_historico_ancho(df_ancho, int(anio_historico), df_referencia)
            except ValueError as exc:
                st.error(str(exc))
                resultados = {}

            if resultados:
                for mes_num, df_mes in resultados.items():
                    _archivar_mes(df_mes, mes_num, int(anio_historico))
                meses_nombres = {v: k for k, v in MESES_ES.items() if k not in ("setiembre",)}
                meses_guardados = ", ".join(meses_nombres.get(m, str(m)) for m in sorted(resultados.keys()))
                st.success(f"Histórico guardado para: {meses_guardados} de {int(anio_historico)}.")


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
    """Pestaña 'Mi Cartera': primero se elige Departamento, luego el Gestor
    dentro de ese departamento. Se busca a sí mismo y ve solo sus PDV."""
    col_depto, col_gestor = st.columns(2)

    with col_depto:
        departamentos_disp = sorted(df["Departamento"].unique())
        depto_sel = st.selectbox("Departamento", departamentos_disp, key="depto_gestor")

    df_depto = df[df["Departamento"] == depto_sel]
    gestores = df_depto[["DNI", "Nombre"]].drop_duplicates().sort_values("Nombre")

    with col_gestor:
        opciones = [f"{fila['Nombre']} · DNI {fila['DNI']}" for _, fila in gestores.iterrows()]
        if not opciones:
            st.info("No hay gestores en este departamento.")
            return
        seleccion = st.selectbox("Gestor", opciones, key="gestor_seleccionado")

    mapa_opcion_dni = dict(zip(opciones, gestores["DNI"]))
    dni_sel = mapa_opcion_dni[seleccion]
    nombre_sel = gestores[gestores["DNI"] == dni_sel]["Nombre"].iloc[0]

    st.markdown(f"**DNI Gestor:** {dni_sel}  ·  **Nombre Gestor:** {nombre_sel}  ·  **Departamento:** {depto_sel}")

    productos_sel = st.multiselect("Producto", options=PRODUCTOS, default=PRODUCTOS, key="productos_gestor")
    if not productos_sel:
        productos_sel = PRODUCTOS
    orden_prod_sel = [p for p in PRODUCTOS if p in productos_sel]

    # Nota: se filtra por Departamento + DNI (no solo DNI), porque en teoría un
    # mismo DNI no debería repetirse en 2 departamentos, pero así queda blindado.
    df_gestor = df[(df["DNI"] == dni_sel) & (df["Departamento"] == depto_sel) & (df["Producto"].isin(productos_sel))]
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
    comision_estimada = calcular_comision_estimada(proy_pct)

    col1, col2, col3, col4, col5, col6 = st.columns(6)
    col1.metric("PDV a cargo", f"{n_pdv}")
    col2.metric("Cuota total", f"{cuota_total:,.0f}")
    col3.metric("Avance", f"{avance_total:,.0f}", f"{cumplimiento:.1%}")
    col4.metric("Proyección", f"{proy_total:,.0f}", f"{proy_pct:.1%}")
    col5.metric("Días restantes", f"{dias_restantes}")
    col6.metric("💰 Comisión estimada", f"S/ {comision_estimada:,.0f}")
    st.caption(
        "La comisión estimada se calcula sobre el % de Proyección de cierre y una tabla de "
        "niveles placeholder (REGLAS_COMISION en el código) — ajustar cuando definan las reglas reales."
    )

    st.markdown("---")
    st.markdown("#### Resumen por Producto (todos mis PDV)")

    resumen_prod = resumen_producto_gestor(df_gestor, productos_sel, dias_en_mes, dia_corte)
    st.dataframe(aplicar_estilo_resumen_gestor(resumen_prod), width="stretch")

    detalle_pdv = detalle_pdv_gestor(df_gestor, dias_en_mes, dia_corte)
    tabla_g = tabla_detalle_gestor(detalle_pdv, productos_sel)
    es_total = tabla_g.index.get_level_values("PDV") == "Total"

    with st.expander(f"➕ Desagrupar: ver detalle por PDV ({n_pdv} punto(s) de venta)"):
        st.dataframe(aplicar_estilo_detalle_pdv(tabla_g[~es_total], orden_prod_sel), width="stretch")

    st.markdown("---")
    st.markdown("#### Ritmo diario necesario")
    st.caption(f"Quedan {dias_restantes} día(s) para el cierre del mes (día {dia_corte} → día {dias_en_mes}).")
    if dias_restantes == 0:
        st.warning("El mes ya cerró; no quedan días para calcular el ritmo diario.")

    ritmo_prod = ritmo_producto_gestor(df_gestor, productos_sel, dias_restantes)
    st.dataframe(aplicar_estilo_ritmo_resumen(ritmo_prod), width="stretch")

    tabla_ritmo = ritmo_pdv_gestor(detalle_pdv, productos_sel, dias_restantes)
    es_total_r = tabla_ritmo.index.get_level_values("PDV") == "Total"
    with st.expander(f"➕ Desagrupar: ver ritmo diario por PDV ({n_pdv} punto(s) de venta)"):
        st.dataframe(aplicar_estilo_ritmo_gestor(tabla_ritmo[~es_total_r], orden_prod_sel), width="stretch")

    st.markdown("---")
    columnas_csv = [
        "DNI", "Nombre", "Departamento", "Provincia", "Distrito", "Producto", "PDV", "Nombre PDV",
        "Cuota", "Avance", "Cumplimiento %", "Proy Unidades",
    ]
    tabla_csv = detalle_pdv[columnas_csv].rename(columns={"DNI": "DNI Gestor", "Nombre": "Nombre Gestor"})
    st.download_button(
        "⬇️ Descargar mi cartera (CSV)",
        data=tabla_csv.sort_values(["Producto", "PDV"]).to_csv(index=False).encode("utf-8"),
        file_name=f"mi_cartera_{dni_sel}.csv",
        mime="text/csv",
    )
    st.caption("🟥 <80% · 🟨 80%–99% · 🟩 ≥100%")


def aplicar_estilo_detalle_plano(tabla: pd.DataFrame):
    """Formato numérico + semáforo para una tabla plana (no pivote) con
    columnas Cuota, Avance, Cumplimiento %, Proy Unidades. Se usa para el
    detalle de PDV desagregado a nivel Departamento en Vista Gerencial."""
    fmt = {}
    for col, patron in (("Cuota", "{:,.0f}"), ("Avance", "{:,.0f}"),
                        ("Cumplimiento %", "{:.1%}"), ("Proy Unidades", "{:,.0f}")):
        if col in tabla.columns:
            fmt[col] = patron
    styler = tabla.style.format(fmt)
    columnas_semaforo = [c for c in ["Cumplimiento %"] if c in tabla.columns]
    return _aplicar_semaforo(styler, columnas_semaforo)


def tabla_comparativo_mensual(df_filtrado: pd.DataFrame, mes: int, anio: int, agrupar_por: str) -> tuple[pd.DataFrame, bool]:
    """Construye la tabla M0 (venta Prepago del mes actual) vs M-1 (venta
    Prepago del mismo Departamento/Gestor el mes anterior, según el
    histórico archivado) y %Var. Devuelve (tabla, hay_historico) — si no hay
    histórico del mes anterior (por ejemplo, el primer mes usando la app),
    hay_historico=False y M-1/%Var quedan vacíos en vez de mostrar '0%'
    engañoso."""
    anio_ant, mes_ant = (anio - 1, 12) if mes == 1 else (anio, mes - 1)
    df_hist = obtener_historico_mes(mes_ant, anio_ant)

    df_prepago = df_filtrado[df_filtrado["Producto"] == "Prepago"].copy()
    if agrupar_por == "Departamento":
        nivel_col = "Departamento"
        df_prepago["_Nivel"] = df_prepago["Departamento"]
    else:
        df_prepago["_Nivel"] = df_prepago["Nombre"] + " · DNI " + df_prepago["DNI"]

    m0 = df_prepago.groupby("_Nivel")["Avance"].sum().rename("M0")

    if df_hist is None:
        comparativo = m0.to_frame()
        comparativo["M-1"] = np.nan
        comparativo["%Var"] = np.nan
        comparativo.index.name = "Departamento" if agrupar_por == "Departamento" else "Gestor"
        return comparativo.sort_values("M0", ascending=False), False

    df_prepago_hist = df_hist[df_hist["Producto"] == "Prepago"].copy()
    if agrupar_por == "Departamento":
        df_prepago_hist["_Nivel"] = df_prepago_hist["Departamento"]
    else:
        df_prepago_hist["_Nivel"] = df_prepago_hist["Nombre"] + " · DNI " + df_prepago_hist["DNI"]
    m1 = df_prepago_hist.groupby("_Nivel")["Avance"].sum().rename("M-1")

    comparativo = pd.concat([m0, m1], axis=1)
    comparativo["M0"] = comparativo["M0"].fillna(0)
    comparativo["M-1"] = comparativo["M-1"].fillna(0)
    comparativo["%Var"] = np.where(
        comparativo["M-1"] > 0,
        (comparativo["M0"] - comparativo["M-1"]) / comparativo["M-1"],
        np.where(comparativo["M0"] > 0, 1.0, 0.0),
    )
    comparativo.index.name = "Departamento" if agrupar_por == "Departamento" else "Gestor"
    return comparativo.sort_values("M0", ascending=False), True


def aplicar_estilo_comparativo(tabla: pd.DataFrame):
    def _color_var(v):
        if pd.isna(v):
            return ""
        return "color: #3E9B4F; font-weight: 600" if v >= 0 else "color: #D64545; font-weight: 600"

    styler = tabla.style.format({"M0": "{:,.0f}", "M-1": "{:,.0f}", "%Var": "{:+.1%}"}, na_rep="—")
    return styler.map(_color_var, subset=["%Var"])


def vista_gerencial(df: pd.DataFrame, dias_en_mes: int, dia_corte: int, mes: int, anio: int) -> None:
    """Pestaña 'Vista Gerencial': una sola tabla tipo 'Resumen por Producto'
    (igual formato que el reporte actual), con dos controles:

    - "Agrupar por": Departamento o Gestor → cambia qué va en las filas.
    - "Desagrupar (ver PDV)": agrega un segundo nivel de fila con el detalle
      de cada Punto de Venta dentro de su Departamento o Gestor.

    Así, con solo esos 2 controles se navegan los 4 niveles de detalle:
    Departamento | Departamento+PDV | Gestor | Gestor+PDV.
    """
    col_a, col_b, col_c, col_d = st.columns([1.2, 1.2, 1.6, 1.6])

    with col_a:
        agrupar_por = st.radio("Agrupar por", ["Departamento", "Gestor"], key="agrupar_por_gerencial", horizontal=True)
    with col_b:
        desagrupar = st.checkbox("➕ Desagrupar (ver PDV)", key="desagrupar_gerencial")
    with col_c:
        departamentos_filtro = st.multiselect(
            "Filtrar Departamento (opcional)", options=sorted(df["Departamento"].unique()),
            default=[], key="depto_filtro_gerencial",
        )
    with col_d:
        productos_sel = st.multiselect(
            "Producto", options=PRODUCTOS, default=PRODUCTOS, key="producto_gerencial",
        )

    if not productos_sel:
        productos_sel = PRODUCTOS
    departamentos_activos = departamentos_filtro if departamentos_filtro else sorted(df["Departamento"].unique())

    # M0 / M-1 son SIEMPRE sobre Prepago, independientemente del filtro de
    # Producto. Se calculan aquí y se agregan como columnas adicionales al
    # final de la tabla principal (después de OSS), no como tarjeta aparte.
    anio_ant, mes_ant = (anio - 1, 12) if mes == 1 else (anio, mes - 1)
    df_hist_mes_ant = obtener_historico_mes(mes_ant, anio_ant)

    m0_total = df[(df["Departamento"].isin(departamentos_activos)) & (df["Producto"] == "Prepago")]["Avance"].sum()
    if df_hist_mes_ant is not None:
        m1_total = df_hist_mes_ant[
            (df_hist_mes_ant["Departamento"].isin(departamentos_activos)) & (df_hist_mes_ant["Producto"] == "Prepago")
        ]["Avance"].sum()
        var_total_pct = ((m0_total - m1_total) / m1_total) if m1_total > 0 else (1.0 if m0_total > 0 else 0.0)
    else:
        m1_total = np.nan
        var_total_pct = np.nan

    df_filtrado = df[df["Departamento"].isin(departamentos_activos) & df["Producto"].isin(productos_sel)]

    if df_filtrado.empty:
        cols_vacio = st.columns(6)
        cols_vacio[0].metric("Gestores", "0")
        cols_vacio[1].metric("PDV", "0")
        cols_vacio[2].metric("Cuota total", "0")
        cols_vacio[3].metric("Avance", "0")
        cols_vacio[4].metric("Proyección", "0")
        cols_vacio[5].metric("💰 Comisión total estimada", "S/ 0")
        st.info("No hay datos para el filtro de Producto seleccionado.")
        return

    # --- KPIs generales sobre lo filtrado ---
    cuota_total = df_filtrado["Cuota"].sum()
    avance_total = df_filtrado["Avance"].sum()
    cumplimiento = (avance_total / cuota_total) if cuota_total > 0 else 0.0
    proy_total = avance_total * (dias_en_mes / max(dia_corte, 1))
    proy_pct = (proy_total / cuota_total) if cuota_total > 0 else 0.0
    comision_total_estimada = df_filtrado.groupby("DNI").apply(
        lambda g: calcular_comision_estimada(
            (g["Avance"].sum() * (dias_en_mes / max(dia_corte, 1))) / g["Cuota"].sum()
            if g["Cuota"].sum() > 0 else 0.0
        ),
        include_groups=False,
    ).sum()

    col1, col2, col3, col4, col5, col6 = st.columns(6)
    col1.metric("Gestores", f"{df_filtrado['DNI'].nunique():,}")
    col2.metric("PDV", f"{df_filtrado['PDV'].nunique():,}")
    col3.metric("Cuota total", f"{cuota_total:,.0f}")
    col4.metric("Avance", f"{avance_total:,.0f}", f"{cumplimiento:.1%}")
    col5.metric("Proyección", f"{proy_total:,.0f}", f"{proy_pct:.1%}")
    col6.metric("💰 Comisión total estimada", f"S/ {comision_total_estimada:,.0f}")

    st.markdown("---")
    titulo = f"Resumen por Producto · agrupado por {agrupar_por}"
    if desagrupar:
        titulo += " (desagrupado por PDV)"
    st.markdown(f"#### {titulo}")

    tabla = construir_tabla_producto(df_filtrado, agrupar_por, desagrupar, productos_sel, dias_en_mes, dia_corte)
    orden_prod_sel = [p for p in PRODUCTOS if p in productos_sel]

    # --- Columnas M0 / M-1 / %Var, agregadas AL FINAL de la tabla (después
    # de OSS), como un grupo más "M0 vs M-1" — solo cuando no se desagrupa
    # por PDV (a nivel PDV individual no se calcula, por ahora). ---
    if not desagrupar:
        df_scope_depto = df[df["Departamento"].isin(departamentos_activos)]
        comparativo, hay_historico = tabla_comparativo_mensual(df_scope_depto, mes, anio, agrupar_por)
        fila_total_comp = pd.DataFrame(
            [{"M0": m0_total, "M-1": m1_total, "%Var": var_total_pct}], index=["Fanero (Total)"]
        )
        comparativo = pd.concat([fila_total_comp, comparativo])
        comparativo.columns = pd.MultiIndex.from_tuples([("M0 vs M-1", c) for c in comparativo.columns])
        tabla = tabla.join(comparativo, how="left")
    else:
        hay_historico = None

    # --- Filtro por Rango de Cumplimiento: se basa en el % de PROYECCIÓN
    # (unidades) de un solo producto — Prepago por defecto (o si es el único
    # elegido), o el producto puntual si se eligió solo uno distinto. ---
    producto_base = producto_base_para_rango(productos_sel)
    incluir_no_activo = producto_base == "Prepago"
    opciones_rango = RANGOS_CON_NO_ACTIVO if incluir_no_activo else RANGOS_SIN_NO_ACTIVO

    # Si el usuario ya tenía seleccionado un rango que ya no es válido para
    # el nuevo producto_base (ej. "PDV no activo" al cambiar a OSS), se
    # limpia ANTES de crear el widget para no romper el multiselect.
    valor_previo = st.session_state.get("rango_cumplimiento_gerencial", [])
    interseccion = [v for v in valor_previo if v in opciones_rango]
    if interseccion != valor_previo:
        st.session_state["rango_cumplimiento_gerencial"] = interseccion

    rango_sel = st.multiselect(
        "Filtrar por Rango de Cumplimiento (opcional)",
        options=opciones_rango, key="rango_cumplimiento_gerencial",
    )
    st.caption(f"Este filtro se calcula sobre el % de Proyección (unidades) de **{producto_base}**.")

    col_proy_base = (producto_base, "Proy Unidades")
    col_cuota_base = (producto_base, "Cuota")
    if rango_sel and col_proy_base in tabla.columns and col_cuota_base in tabla.columns:
        proy_fila = tabla[col_proy_base]
        cuota_fila = tabla[col_cuota_base]
        proy_pct_fila = np.where(cuota_fila > 0, proy_fila / cuota_fila, 0.0)
        rango_fila = pd.Series(proy_pct_fila, index=tabla.index).map(
            lambda v: clasificar_rango_proyeccion(v, incluir_no_activo)
        )
        # La fila "Fanero (Total)" nunca se oculta por este filtro — sirve
        # de contexto siempre visible, sin importar qué rango se elija.
        mascara = rango_fila.isin(rango_sel)
        if "Fanero (Total)" in tabla.index:
            mascara.loc["Fanero (Total)"] = True
        tabla = tabla[mascara]

    altura = 480 if desagrupar else "content"
    if tabla.empty:
        st.info("No hay filas para el Rango de Cumplimiento seleccionado.")
    else:
        st.dataframe(aplicar_estilo_resumen_producto(tabla, orden_prod_sel), width="stretch", height=altura)
    leyenda = "🟥 <80% · 🟨 80%–99% · 🟩 ≥100% (aplica a Proy %)"
    if not desagrupar:
        leyenda += " · M0 = venta Prepago este mes · M-1 = venta Prepago mes anterior"
    st.caption(leyenda)
    if not desagrupar and not hay_historico:
        st.info(
            "M-1 y %Var todavía no tienen datos: no hay histórico del mes anterior guardado "
            "(normal en el primer mes usando la app, o puedes cargarlo manualmente en el panel admin)."
        )

    tabla_csv = tabla.copy()
    tabla_csv.columns = [f"{p} - {m}" for p, m in tabla_csv.columns]
    nombre_archivo = f"resumen_{agrupar_por.lower()}{'_pdv' if desagrupar else ''}.csv"
    st.download_button(
        "⬇️ Descargar esta tabla (CSV)",
        data=tabla_csv.reset_index().to_csv(index=False).encode("utf-8"),
        file_name=nombre_archivo,
        mime="text/csv",
    )


def render_acceso_admin_sidebar() -> None:
    """Acceso admin discreto en la barra lateral — el dashboard es público
    y no requiere login para verse. Solo quien conoce las credenciales admin
    puede iniciar sesión ahí para publicar datos."""
    if st.session_state.get("es_admin", False):
        panel_admin()
        st.markdown("---")
        if st.button("Cerrar sesión"):
            st.session_state["es_admin"] = False
            st.rerun()
        return

    with st.expander("🔒 Administrador"):
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


def main():
    # El dashboard es público (no requiere login para verse). El acceso
    # admin vive discretamente en la barra lateral, para quien necesite
    # publicar datos.
    st.title("📊 Mi Cartera - Gestores Fanero")

    with st.sidebar:
        render_acceso_admin_sidebar()

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

    tabs = st.tabs(["👤 Mi Cartera (Gestor)", "🏢 Vista Gerencial", "✏️ Editar Avances"])

    with tabs[0]:
        vista_gestor(df, dias_en_mes, dia_corte, dias_restantes)

    with tabs[1]:
        vista_gerencial(df, dias_en_mes, dia_corte, mes, anio)

    with tabs[2]:
        st.subheader("Editar Avances")
        panel_editar_avances(df_raw)


if __name__ == "__main__":
    main()
