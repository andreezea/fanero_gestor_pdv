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
import unicodedata
from datetime import datetime
from io import BytesIO

import numpy as np
import pandas as pd
import streamlit as st
import altair as alt
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
HISTORIAL_DIARIO_FILE = os.path.join(DATA_DIR, "historial_diario.xlsx")

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

# ---------------------------------------------------------------------------
# CÁLCULO DE COMISIONES DE GESTORES (pestaña dedicada)
# ---------------------------------------------------------------------------
# Solo estos 2 productos entran en el cálculo de comisión (ajustar aquí si
# cambia la regla de negocio).
PRODUCTOS_COMISION = ["Prepago", "Postpago"]

# Monto en soles que se paga por cada 100% de Proy% alcanzado, por producto.
# Ej: si Proy% Prepago = 95%, Comisión Prepago = 0.95 * 644.
MONTO_COMISION_PRODUCTO = {"Prepago": 644.0, "Postpago": 526.0}

# El Proy% nunca puede superar este tope para efectos de comisión (aunque la
# proyección real sea mayor, se paga como máximo hasta aquí).
TOPE_PROY_PCT_COMISION = 1.10

# Archivo donde se guarda lo que el BO ingresa (PDV totales y visitas
# promedio por gestor) — persiste entre sesiones, igual que los demás datos.
VISITAS_FILE = os.path.join(DATA_DIR, "visitas_bo.xlsx")

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


def _normalizar_texto_simple(texto: str) -> str:
    """minúsculas, sin tildes — para comparar 'LORETO' con 'Loreto' sin fallar."""
    if texto is None:
        return ""
    texto = str(texto).strip().lower()
    return unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode("ascii")


_MAPA_DEPARTAMENTO_CANONICO = {_normalizar_texto_simple(d): d for d in DEPARTAMENTOS}
_MAPA_PRODUCTO_CANONICO = {_normalizar_texto_simple(p): p for p in PRODUCTOS}


# Columnas que deben leerse SIEMPRE como texto: si un DNI o código de PDV
# tiene ceros a la izquierda (ej. "05336082"), pandas los interpreta como
# número al leer el Excel y los pierde ("5336082") — dtype=str lo evita.
_COLUMNAS_FORZAR_TEXTO = {"DNI": str, "PDV": str}


def leer_excel_seguro(ruta_o_archivo, **kwargs) -> pd.DataFrame:
    """Envoltorio de pd.read_excel que fuerza DNI/PDV como texto (evita
    perder ceros a la izquierda). Si el archivo no tiene esas columnas,
    pandas simplemente las ignora, sin error."""
    return pd.read_excel(ruta_o_archivo, dtype=_COLUMNAS_FORZAR_TEXTO, **kwargs)


def _normalizar_identidad(df: pd.DataFrame) -> pd.DataFrame:
    """Aplica la normalización de texto a todas las columnas de identidad.
    Además, homogeneiza Departamento y Producto a su forma canónica sin
    importar cómo vengan escritos en el Excel (mayúsculas, tildes, espacios):
    "LORETO", "loreto ", "Loreto" → siempre "Loreto". Sin esto, una fila con
    "LORETO" no calzaría con la lista interna de departamentos y quedaría
    invisible en toda la app (tabla vacía, "sin datos")."""
    df = df.copy()
    for columna in ["DNI", "Nombre", "PDV", "Nombre PDV", "Departamento", "Provincia", "Distrito"]:
        relleno = "Sin dato" if columna in ("Departamento", "Provincia", "Distrito") else ""
        df = _normalizar_texto(df, columna, relleno)

    if "Departamento" in df.columns:
        df["Departamento"] = df["Departamento"].map(
            lambda v: _MAPA_DEPARTAMENTO_CANONICO.get(_normalizar_texto_simple(v), v)
        )
    if "Producto" in df.columns:
        df["Producto"] = df["Producto"].map(
            lambda v: _MAPA_PRODUCTO_CANONICO.get(_normalizar_texto_simple(v), v)
        )
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
        if hasattr(archivo, "seek"):
            archivo.seek(0)  # por si ya se leyó antes (ej. para la vista previa)
        df = leer_excel_seguro(archivo)
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
    df = leer_excel_seguro(path)
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
        return _normalizar_identidad(leer_excel_seguro(ruta))
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
    """Convierte una tabla ANCHA de ventas de meses anteriores (una fila por
    PDV, una columna por producto) al formato interno que usa la app,
    agrupando por mes si el archivo trae más de uno. Columnas esperadas:

        DNI PDV | Nombre PDV | Departamento | Provincia | Distrito | DNI Gestor | Mes | Prepago | Porta Prepago | Postpago | OSS

    Obligatorias: "DNI PDV", "DNI Gestor", "Mes", y al menos un producto.
    "Nombre PDV", "Departamento", "Provincia", "Distrito" son recomendadas
    pero opcionales — si vienen en el archivo, se usan directamente (lo más
    confiable, porque cada PDV ya trae su propio departamento, sin
    depender de nada más). Si faltan, se completan buscando el DNI Gestor
    en `df_referencia` (lo ya publicado este mes) como respaldo — pero si
    ese gestor todavía no está publicado, quedan vacías.

    Devuelve {mes_numero: DataFrame_listo_para_archivar}.
    """
    columnas_obligatorias = {"Mes"}
    faltantes_obligatorias = columnas_obligatorias - set(df_ancho.columns)
    if faltantes_obligatorias:
        raise ValueError("Faltan columnas obligatorias: " + ", ".join(sorted(faltantes_obligatorias)))

    col_pdv_origen = "DNI PDV" if "DNI PDV" in df_ancho.columns else ("PDV" if "PDV" in df_ancho.columns else None)
    col_gestor_origen = "DNI Gestor" if "DNI Gestor" in df_ancho.columns else ("DNI" if "DNI" in df_ancho.columns else None)
    faltantes_id = []
    if col_pdv_origen is None:
        faltantes_id.append("DNI PDV (o PDV)")
    if col_gestor_origen is None:
        faltantes_id.append("DNI Gestor (o DNI)")
    if faltantes_id:
        raise ValueError("Faltan columnas obligatorias: " + ", ".join(faltantes_id))

    df_ancho = df_ancho.copy()
    df_ancho["PDV"] = df_ancho[col_pdv_origen].astype(str).str.strip()
    df_ancho["DNI"] = df_ancho[col_gestor_origen].astype(str).str.strip()
    # Si no se conoce el DNI Gestor de un PDV (celda vacía), se agrupa bajo
    # un "gestor" placeholder claro — así el Departamento sigue sumando bien
    # (no depende del Gestor), y en la Vista Gerencial por Gestor aparece
    # una fila legible "Sin gestor asignado" en vez de una fila en blanco.
    sin_gestor_conocido = df_ancho["DNI"].isna() | (df_ancho["DNI"].isin(["", "nan", "None"]))
    df_ancho.loc[sin_gestor_conocido, "DNI"] = "SIN_GESTOR"
    df_ancho["_MesTexto"] = df_ancho["Mes"].astype(str).str.strip().str.lower()
    df_ancho["_MesNumero"] = df_ancho["_MesTexto"].map(MESES_ES)

    columnas_geo_presentes = [c for c in ["Nombre PDV", "Departamento", "Provincia", "Distrito"] if c in df_ancho.columns]
    nombre_gestor_presente = "Nombre" in df_ancho.columns
    columnas_excluir = {"DNI PDV", "DNI Gestor", "Mes", "PDV", "DNI", "Nombre", "_MesTexto", "_MesNumero"} | set(columnas_geo_presentes)

    mapa_col_producto = {}
    for col in df_ancho.columns:
        if col in columnas_excluir:
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

    # Respaldo por DNI Gestor: para lo que NO venga en el archivo
    # (Nombre PDV/Departamento/Provincia/Distrito), se busca en lo ya
    # publicado este mes. "Nombre" (del gestor) se maneja aparte abajo.
    columnas_geo_faltantes = [c for c in ["Nombre PDV", "Departamento", "Provincia", "Distrito"] if c not in columnas_geo_presentes]
    columnas_respaldo = columnas_geo_faltantes + ([] if nombre_gestor_presente else ["Nombre"])
    ref_cols = [c for c in (["DNI"] + columnas_respaldo) if c in df_referencia.columns]
    ref = df_referencia[ref_cols].drop_duplicates(subset=["DNI"]) if "DNI" in ref_cols and len(ref_cols) > 1 else pd.DataFrame(columns=["DNI"])

    resultados: dict[int, pd.DataFrame] = {}
    for mes_num, grupo_mes in df_ancho.groupby("_MesNumero"):
        filas = []
        for _, fila in grupo_mes.iterrows():
            for col_original, producto in mapa_col_producto.items():
                valor = fila[col_original]
                if pd.isna(valor):
                    continue
                fila_nueva = {"DNI": fila["DNI"], "PDV": fila["PDV"], "Producto": producto, "Avance": float(valor)}
                for col_geo in columnas_geo_presentes:
                    fila_nueva[col_geo] = fila[col_geo]
                if nombre_gestor_presente:
                    fila_nueva["Nombre"] = fila["Nombre"]
                filas.append(fila_nueva)
        df_largo = pd.DataFrame(filas)
        if df_largo.empty:
            continue
        if not ref.empty:
            df_largo = df_largo.merge(ref, on="DNI", how="left")

        # El Nombre del gestor no viene en esta plantilla: si tampoco se
        # encontró en lo ya publicado, se usa un nombre genérico legible en
        # vez de dejarlo vacío (evita filas "sin nombre" en la tabla).
        # "SIN_GESTOR" (DNI Gestor vacío en el archivo) tiene su propio
        # nombre claro, distinto del genérico "Gestor <DNI>".
        if "Nombre" not in df_largo.columns:
            df_largo["Nombre"] = ""
        df_largo["Nombre"] = df_largo["Nombre"].fillna("")
        es_sin_gestor = df_largo["DNI"] == "SIN_GESTOR"
        df_largo.loc[es_sin_gestor & (df_largo["Nombre"] == ""), "Nombre"] = "Sin gestor asignado"
        sin_nombre = df_largo["Nombre"] == ""
        df_largo.loc[sin_nombre, "Nombre"] = "Gestor " + df_largo.loc[sin_nombre, "DNI"]

        resultados[int(mes_num)] = _normalizar_identidad(df_largo)

    return resultados

def registrar_incrementos_diarios_lote(entradas: list) -> None:
    """Como `registrar_incremento_diario`, pero para VARIOS días a la vez —
    hace UNA sola lectura y UNA sola escritura del archivo (en vez de una
    por cada día), mucho más rápido para cargas de muchos días/PDV.
    `entradas` es una lista de tuplas (df_incremento, fecha)."""
    columnas = ["DNI", "Nombre", "Departamento", "PDV", "Producto"]
    agregados = []
    for df_incremento, fecha in entradas:
        columnas_presentes = [c for c in columnas if c in df_incremento.columns]
        agregado = df_incremento.groupby(columnas_presentes, as_index=False)["Avance"].sum()
        agregado["Fecha"] = pd.Timestamp(fecha)
        agregados.append(agregado)

    if not agregados:
        return

    nuevo = pd.concat(agregados, ignore_index=True)
    columnas_presentes = [c for c in columnas if c in nuevo.columns]
    claves = ["Fecha", "DNI", "Producto"] + (["PDV"] if "PDV" in columnas_presentes else [])

    if os.path.exists(HISTORIAL_DIARIO_FILE):
        historial = leer_excel_seguro(HISTORIAL_DIARIO_FILE, parse_dates=["Fecha"])
        historial["DNI"] = historial["DNI"].astype(str).str.strip()  # evita que Excel lo lea como número
        nuevo["DNI"] = nuevo["DNI"].astype(str).str.strip()
        if "PDV" in claves and "PDV" not in historial.columns:
            historial["PDV"] = ""  # archivo viejo sin columna PDV: se completa vacía
        historial = historial.set_index(claves)
        nuevo_idx = nuevo.set_index(claves)
        historial = historial.drop(index=historial.index.intersection(nuevo_idx.index), errors="ignore")
        historial = pd.concat([historial.reset_index(), nuevo_idx.reset_index()], ignore_index=True)
    else:
        historial = nuevo

    os.makedirs(DATA_DIR, exist_ok=True)
    historial.to_excel(HISTORIAL_DIARIO_FILE, index=False)


def registrar_incremento_diario(df_incremento: pd.DataFrame, fecha: "pd.Timestamp") -> None:
    """Guarda, con fecha, lo que se sumó en ESTA publicación (por Gestor,
    PDV y Producto) — es la base para el gráfico de 'Ventas diarias' y para
    la comparación M0 vs M-1 'mismo día contra mismo día'. Si ya existía un
    registro para esa misma Fecha+DNI+PDV+Producto (se volvió a publicar el
    mismo día), se reemplaza en vez de duplicar.

    Para registrar VARIOS días de una sola vez (mucho más rápido — una sola
    lectura/escritura de archivo), usar `registrar_incrementos_diarios_lote`."""
    registrar_incrementos_diarios_lote([(df_incremento, fecha)])


@st.cache_data
def _leer_historial_diario_cacheado(path: str, mtime: float) -> pd.DataFrame:
    """Lee el historial diario. `mtime` forma parte de la clave de cache: si
    el archivo no cambió desde la última carga, no se vuelve a leer del
    disco — esto es clave porque este archivo crece con cada publicación y
    se consulta en cada re-render de Vista Gerencial (para el gráfico de
    Ventas diarias y las comparaciones M0/M-1), sin cache se vuelve lento a
    medida que crecen los datos reales."""
    df = leer_excel_seguro(path, parse_dates=["Fecha"])
    df["DNI"] = df["DNI"].astype(str).str.strip()
    return df


def obtener_historial_diario() -> pd.DataFrame:
    """Devuelve el historial completo de incrementos diarios (Fecha, DNI,
    Nombre, Departamento, Producto, Avance), o una tabla vacía si aún no
    hay ninguna publicación registrada."""
    if os.path.exists(HISTORIAL_DIARIO_FILE):
        try:
            mtime = os.path.getmtime(HISTORIAL_DIARIO_FILE)
            return _leer_historial_diario_cacheado(HISTORIAL_DIARIO_FILE, mtime)
        except Exception:  # noqa: BLE001 - archivo corrupto
            pass
    return pd.DataFrame(columns=["Fecha", "DNI", "Nombre", "Departamento", "Producto", "Avance"])


def hay_detalle_diario_del_mes(mes: int, anio: int) -> bool:
    """True si el historial diario tiene AL MENOS un registro de ese
    Mes/Año — permite decidir si se puede hacer la comparación 'mismo día
    contra mismo día', o si hay que usar el total del mes completo."""
    historial = obtener_historial_diario()
    if historial.empty:
        return False
    return not historial[(historial["Fecha"].dt.year == anio) & (historial["Fecha"].dt.month == mes)].empty


def avance_acumulado_hasta_dia(
    mes: int, anio: int, dia_corte: int, departamentos_activos: list, producto: str,
) -> float:
    """Suma el historial diario de un Mes/Año, para un Producto y un
    conjunto de Departamentos, SOLO hasta el día `dia_corte` (inclusive) —
    es la base para comparar M0 vs M-1 'mismo día contra mismo día' en vez
    de mes completo contra mes completo. Devuelve 0 si no hay registros."""
    historial = obtener_historial_diario()
    if historial.empty:
        return 0.0
    filtro = (
        (historial["Fecha"].dt.year == anio) & (historial["Fecha"].dt.month == mes)
        & (historial["Fecha"].dt.day <= dia_corte)
        & (historial["Departamento"].isin(departamentos_activos))
        & (historial["Producto"] == producto)
    )
    return float(historial[filtro]["Avance"].sum())


def avance_acumulado_hasta_dia_por_pdv(mes: int, anio: int, dia_corte: int, producto: str) -> pd.Series:
    """Como `avance_acumulado_hasta_dia`, pero devuelve una Serie indexada
    por código de PDV (para la comparación M0 vs M-1 desagrupada por PDV).
    Los registros sin PDV (archivos históricos de antes de este cambio) se
    ignoran — no se pueden comparar a nivel PDV, solo a nivel Departamento/Gestor."""
    historial = obtener_historial_diario()
    if historial.empty or "PDV" not in historial.columns:
        return pd.Series(dtype=float)
    filtro = (
        (historial["Fecha"].dt.year == anio) & (historial["Fecha"].dt.month == mes)
        & (historial["Fecha"].dt.day <= dia_corte) & (historial["Producto"] == producto)
        & (historial["PDV"].astype(str).str.strip() != "")
    )
    subset = historial[filtro]
    if subset.empty:
        return pd.Series(dtype=float)
    return subset.groupby("PDV")["Avance"].sum()


def publicar_datos_incremental(
    df_nuevo: pd.DataFrame, dia_corte: int, mes: int, anio: int, registrar_historial: bool = True,
) -> None:
    """Publica una carga SUMANDO el Avance nuevo al acumulado ya publicado
    del mismo Mes/Año (en vez de reemplazarlo). Así el administrador puede
    subir solo lo vendido en el día/periodo más reciente, y la app se
    encarga de mantener el acumulado del mes.

    `registrar_historial=False` evita registrar este incremento en
    `historial_diario` — se usa cuando el llamador ya lo va a registrar por
    su cuenta (ej. `procesar_carga_horizontal_diaria`, que registra todos
    los días de una sola vez en lote, mucho más rápido que uno por uno).

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
        if registrar_historial:
            ultimo_dia_mes = calendar.monthrange(anio, mes)[1]
            fecha_publicacion = pd.Timestamp(year=anio, month=mes, day=min(dia_corte, ultimo_dia_mes))
            registrar_incremento_diario(df_nuevo, fecha_publicacion)
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

    # Registra el incremento de ESTA publicación con su fecha, para el
    # gráfico de "Ventas diarias" en Vista Gerencial.
    if registrar_historial:
        ultimo_dia_mes = calendar.monthrange(anio, mes)[1]
        fecha_publicacion = pd.Timestamp(year=anio, month=mes, day=min(dia_corte, ultimo_dia_mes))
        registrar_incremento_diario(df_nuevo, fecha_publicacion)


def publicar_datos_reemplazo_parcial(
    df_nuevo: pd.DataFrame, dia_corte: int, mes: int, anio: int, registrar_historial: bool = True,
) -> None:
    """Como `publicar_datos_incremental`, pero en vez de SUMAR el Avance
    nuevo al ya publicado, lo REEMPLAZA (sobreescribe) para los DNI+PDV+
    Producto que traiga el archivo — sin tocar nada más de lo que ya está
    publicado (otros productos, otros PDV). Útil para volver a subir un
    rango de días ya cargado antes (corrige un error) sin que se duplique.

    Si el Mes/Año es distinto al publicado, se comporta igual que
    `publicar_datos_incremental` (archiva el mes saliente y empieza de cero).
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
        publicar_datos(df_nuevo, dia_corte, mes, anio)
    else:
        claves = ["DNI", "PDV", "Producto"]
        df_actual = _normalizar_identidad(df_actual)
        base = df_actual.set_index(claves)
        nuevo = df_nuevo.set_index(claves)

        # Reemplaza (no suma) Avance y Cuota para las filas que trae el
        # archivo nuevo; todo lo demás en `base` queda intacto.
        claves_comunes = base.index.intersection(nuevo.index)
        base.loc[claves_comunes, "Avance"] = nuevo.loc[claves_comunes, "Avance"]
        if "Cuota" in nuevo.columns:
            base.loc[claves_comunes, "Cuota"] = nuevo.loc[claves_comunes, "Cuota"]

        claves_nuevas = nuevo.index.difference(base.index)
        filas_nuevas = nuevo.loc[claves_nuevas]

        combinado = pd.concat([base, filas_nuevas]).reset_index()
        publicar_datos(combinado, dia_corte, mes, anio)

    if registrar_historial:
        ultimo_dia_mes = calendar.monthrange(anio, mes)[1]
        fecha_publicacion = pd.Timestamp(year=anio, month=mes, day=min(dia_corte, ultimo_dia_mes))
        registrar_incremento_diario(df_nuevo, fecha_publicacion)


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


def cargar_visitas_bo() -> pd.DataFrame:
    """Carga lo último guardado por el BO: DNI, Nombre, PDV Totales, Visitas
    PDV. Si no existe el archivo todavía, devuelve una tabla vacía. Si el
    archivo es de una versión anterior (columna "Visitas Promedio"), se
    renombra sola para no perder los datos ya guardados."""
    if os.path.exists(VISITAS_FILE):
        try:
            df = leer_excel_seguro(VISITAS_FILE)
            df["DNI"] = df["DNI"].astype(str).str.strip()
            if "Visitas Promedio" in df.columns and "Visitas PDV" not in df.columns:
                df = df.rename(columns={"Visitas Promedio": "Visitas PDV"})
            return df
        except Exception:  # noqa: BLE001 - archivo corrupto o con formato viejo
            pass
    return pd.DataFrame(columns=["DNI", "Nombre", "PDV Totales", "Visitas PDV"])


def guardar_visitas_bo(df: pd.DataFrame) -> None:
    os.makedirs(DATA_DIR, exist_ok=True)
    df.to_excel(VISITAS_FILE, index=False)


def construir_tabla_comisiones(
    df_filtrado: pd.DataFrame, visitas_df: pd.DataFrame, dias_en_mes: int, dia_corte: int,
) -> pd.DataFrame:
    """Tabla ÚNICA de cálculo de comisiones: una fila por Gestor (+ fila
    "Fanero" con el total), con las columnas en este orden exacto:

        Gestor | PREPAGO (Cuota, Avance, Proy Unidades, Proy %)
               | VISITAS GESTOR (Visitas PDV, PDV Totales, Visitas %)
               | POSTPAGO (Cuota, Avance, Proy Unidades, Proy %)
               | COMISIÓN (Prepago, Postpago, Total)

    `visitas_df` trae DNI, "Visitas PDV", "PDV Totales" (lo que ingresa el BO).

    Orden de cálculo (importante, no cambiar el orden):
    1. Avance = venta real acumulada, TAL CUAL viene del archivo (no se toca).
    2. Proy Unidades = Avance * (días del mes / día de corte) — proyección
       normal, sin ningún ajuste. Esto es lo que se MUESTRA en la tabla.
    3. Proy % = Proy Unidades / Cuota, con tope de TOPE_PROY_PCT_COMISION
       (110% por defecto) — aplica a Prepago y Postpago por igual.
    4. SOLO para Prepago: el Proy% (ya con el tope aplicado) se multiplica
       por Visitas % = Visitas PDV / PDV Totales. Este ajuste es al
       PORCENTAJE, nunca a las Unidades — así la Proy Unidades que se ve en
       pantalla siempre refleja la proyección real.
    5. Comisión_producto = Proy% (ya ajustado) * MONTO_COMISION_PRODUCTO[producto].
    """
    mapa_visitas_pdv = dict(zip(visitas_df["DNI"], visitas_df["Visitas PDV"]))
    mapa_pdv_totales = dict(zip(visitas_df["DNI"], visitas_df["PDV Totales"]))
    mapa_pct_visitas = {
        dni: (mapa_visitas_pdv.get(dni, 0) / mapa_pdv_totales[dni]) if mapa_pdv_totales.get(dni, 0) > 0 else 1.0
        for dni in mapa_pdv_totales
    }

    df_prod = df_filtrado[df_filtrado["Producto"].isin(PRODUCTOS_COMISION)].copy()
    factor_proyeccion = dias_en_mes / max(dia_corte, 1)

    agregado = (
        df_prod.groupby(["DNI", "Nombre", "Producto"], as_index=False)
        .agg(Cuota=("Cuota", "sum"), Avance=("Avance", "sum"))
    )
    # Proyección normal, SIN ajustar — esto es lo que se muestra en "Proy Unidades".
    agregado["Proy Unidades"] = agregado["Avance"] * factor_proyeccion
    agregado["Proy %"] = np.where(agregado["Cuota"] > 0, agregado["Proy Unidades"] / agregado["Cuota"], 0.0)
    agregado["Proy %"] = agregado["Proy %"].clip(upper=TOPE_PROY_PCT_COMISION)

    # El ajuste por %Visitas se aplica al % (no a las unidades), y SOLO a Prepago.
    agregado["Proy % Comisión"] = agregado["Proy %"]
    es_prepago = agregado["Producto"] == "Prepago"
    pct_visitas_fila = agregado["DNI"].map(mapa_pct_visitas).fillna(1.0)
    agregado.loc[es_prepago, "Proy % Comisión"] = agregado.loc[es_prepago, "Proy %"] * pct_visitas_fila[es_prepago]
    agregado["Comision"] = agregado["Producto"].map(MONTO_COMISION_PRODUCTO) * agregado["Proy % Comisión"]

    metricas = ["Cuota", "Avance", "Proy Unidades", "Proy %"]
    pivot = agregado.pivot_table(index=["DNI", "Nombre"], columns="Producto", values=metricas, aggfunc="first")
    pivot = pivot.swaplevel(axis=1)

    prepago_cols = pd.MultiIndex.from_product([["Prepago"], metricas])
    postpago_cols = pd.MultiIndex.from_product([["Postpago"], metricas])
    bloque_prepago = pivot.reindex(columns=prepago_cols)
    bloque_postpago = pivot.reindex(columns=postpago_cols)

    # --- Bloque "Visitas Gestor" (entre Prepago y Postpago) ---
    dnis_presentes = pivot.index.get_level_values("DNI")
    visitas_pdv_serie = pd.Series([mapa_visitas_pdv.get(dni, 0) for dni in dnis_presentes], index=pivot.index)
    pdv_totales_serie = pd.Series([mapa_pdv_totales.get(dni, 0) for dni in dnis_presentes], index=pivot.index)
    visitas_pct_serie = pd.Series([mapa_pct_visitas.get(dni, 1.0) for dni in dnis_presentes], index=pivot.index)
    bloque_visitas = pd.DataFrame({
        ("Visitas Gestor", "Visitas PDV"): visitas_pdv_serie,
        ("Visitas Gestor", "PDV Totales"): pdv_totales_serie,
        ("Visitas Gestor", "Visitas %"): visitas_pct_serie,
    })
    bloque_visitas.columns = pd.MultiIndex.from_tuples(bloque_visitas.columns)

    # --- Bloque "Comisión" ---
    comision_pivot = agregado.pivot_table(index=["DNI", "Nombre"], columns="Producto", values="Comision", aggfunc="first")
    comision_pivot = comision_pivot.reindex(columns=PRODUCTOS_COMISION).fillna(0.0)
    comision_pivot["Total"] = comision_pivot.sum(axis=1)
    comision_pivot.columns = pd.MultiIndex.from_tuples([("Comisión", c) for c in comision_pivot.columns])

    # Orden final EXACTO: Prepago | Visitas Gestor | Postpago | Comisión
    tabla = pd.concat([bloque_prepago, bloque_visitas, bloque_postpago, comision_pivot], axis=1)
    tabla = tabla.droplevel("DNI")  # el DNI ya cumplió su función (agrupar bien); solo se muestra el Nombre
    tabla.index.name = "Gestor"
    tabla = tabla.sort_index()

    # --- Fila "Fanero": Prepago/Postpago = totales reales recalculados (sin
    # ajuste de visitas, para que el % sea comparable). Visitas Gestor =
    # suma de Visitas PDV y PDV Totales, con Visitas% recalculado sobre esos
    # totales. Comisión = SUMA de las comisiones individuales (ya ajustadas
    # por visitas) — es el número financieramente correcto a pagar. ---
    fila_total = {}
    for p in PRODUCTOS_COMISION:
        cuota_p = tabla[(p, "Cuota")].sum()
        avance_p = tabla[(p, "Avance")].sum()
        proy_p = tabla[(p, "Proy Unidades")].sum()
        proy_pct_p = min((proy_p / cuota_p) if cuota_p > 0 else 0.0, TOPE_PROY_PCT_COMISION)
        fila_total[(p, "Cuota")] = cuota_p
        fila_total[(p, "Avance")] = avance_p
        fila_total[(p, "Proy Unidades")] = proy_p
        fila_total[(p, "Proy %")] = proy_pct_p
        fila_total[("Comisión", p)] = tabla[("Comisión", p)].sum()
    visitas_pdv_total = tabla[("Visitas Gestor", "Visitas PDV")].sum()
    pdv_totales_total = tabla[("Visitas Gestor", "PDV Totales")].sum()
    fila_total[("Visitas Gestor", "Visitas PDV")] = visitas_pdv_total
    fila_total[("Visitas Gestor", "PDV Totales")] = pdv_totales_total
    fila_total[("Visitas Gestor", "Visitas %")] = (visitas_pdv_total / pdv_totales_total) if pdv_totales_total > 0 else 0.0
    fila_total[("Comisión", "Total")] = sum(fila_total[("Comisión", p)] for p in PRODUCTOS_COMISION)

    df_total = pd.DataFrame([fila_total], index=["Fanero"], columns=tabla.columns)
    tabla = pd.concat([tabla, df_total])
    tabla.index.name = "Gestor"  # pd.concat no siempre conserva el nombre del índice

    return tabla


def aplicar_estilo_comisiones(tabla: pd.DataFrame):
    """Formato + estilo para la tabla de comisiones (reset_index primero,
    igual que aplicar_estilo_resumen_producto, por si hay gestores homónimos)."""
    nombre_indice = tabla.index.name or "Gestor"
    tabla = tabla.reset_index()
    col_gestor = (nombre_indice, "") if isinstance(tabla.columns, pd.MultiIndex) else nombre_indice
    if col_gestor not in tabla.columns:
        # reset_index() puede nombrar la columna "index" si el índice no tenía nombre
        col_gestor = tabla.columns[0]

    fmt = {}
    for p in PRODUCTOS_COMISION:
        fmt[(p, "Cuota")] = "{:,.0f}"
        fmt[(p, "Avance")] = "{:,.0f}"
        fmt[(p, "Proy Unidades")] = "{:,.0f}"
        fmt[(p, "Proy %")] = "{:.1%}"
    fmt[("Visitas Gestor", "Visitas PDV")] = "{:,.0f}"
    fmt[("Visitas Gestor", "PDV Totales")] = "{:,.0f}"
    fmt[("Visitas Gestor", "Visitas %")] = "{:.1%}"
    fmt[("Comisión", "Prepago")] = "S/ {:,.0f}"
    fmt[("Comisión", "Postpago")] = "S/ {:,.0f}"
    fmt[("Comisión", "Total")] = "S/ {:,.0f}"

    styler = tabla.style.format(fmt, na_rep="-").hide(axis="index")
    subset = [(p, "Proy %") for p in PRODUCTOS_COMISION]
    styler = _aplicar_semaforo(styler, subset)

    if (tabla[col_gestor] == "Fanero").any():
        def _estilo_total(fila):
            if fila[col_gestor] == "Fanero":
                return ["background-color: #00405E; color: #FFFFFF; font-weight: 700;" for _ in fila]
            return ["" for _ in fila]
        styler = styler.apply(_estilo_total, axis=1)

    styler = styler.set_table_styles(
        [
            {"selector": "table", "props": [("border-collapse", "collapse"), ("width", "100%"), ("font-size", "0.85rem")]},
            {"selector": "th, td", "props": [("border", "1px solid #E2E4F0"), ("padding", "6px 10px"), ("text-align", "center"), ("white-space", "nowrap")]},
            {"selector": "th", "props": [("background-color", "#F5F6FB"), ("font-weight", "600")]},
        ],
        overwrite=True,
    )
    styler = styler.set_properties(**{"text-align": "center"})
    return styler


def vista_comisiones_gestores(df: pd.DataFrame, dias_en_mes: int, dia_corte: int) -> None:
    """Pestaña 'Cálculo Comisión Gestor': una sola tabla con Prepago |
    Visitas Gestor | Postpago | Comisión. La edición de PDV Totales/Visitas
    PDV por el BO queda en un expander aparte, para que lo único que
    resalte en pantalla sea el cálculo."""
    departamentos_disponibles = sorted(df["Departamento"].unique())
    departamentos_sel_comision = st.multiselect(
        "Departamento (opcional — el BO puede filtrar solo los suyos)",
        options=departamentos_disponibles, default=[], key="depto_filtro_comision",
    )
    if departamentos_sel_comision:
        df = df[df["Departamento"].isin(departamentos_sel_comision)]
        if df.empty:
            st.info("No hay gestores para el/los departamento(s) seleccionado(s).")
            return

    conteo_pdv_actual = df.groupby("DNI")["PDV"].nunique().rename("PDV Totales (actual)")
    gestores_actuales = df[["DNI", "Nombre"]].drop_duplicates()

    visitas_guardadas = cargar_visitas_bo()
    tabla_visitas = gestores_actuales.merge(
        visitas_guardadas[["DNI", "PDV Totales", "Visitas PDV"]], on="DNI", how="left"
    )
    tabla_visitas = tabla_visitas.merge(conteo_pdv_actual, on="DNI", how="left")
    tabla_visitas["PDV Totales"] = tabla_visitas["PDV Totales"].fillna(tabla_visitas["PDV Totales (actual)"]).fillna(0)
    tabla_visitas["Visitas PDV"] = tabla_visitas["Visitas PDV"].fillna(0)
    tabla_visitas = tabla_visitas.drop(columns=["PDV Totales (actual)"]).sort_values("Nombre")

    with st.expander("✏️ Actualizar PDV Totales y Visitas (Back Office)"):
        st.caption(
            "El BO ingresa la cantidad de PDV que tiene cada gestor y cuántos visitó. "
            "El sistema calcula Visitas % = Visitas PDV / PDV Totales, y ese % ajusta la comisión de Prepago."
        )
        editado = st.data_editor(
            tabla_visitas[["DNI", "Nombre", "PDV Totales", "Visitas PDV"]],
            disabled=["DNI", "Nombre"],
            hide_index=True,
            width="stretch",
            key="editor_visitas_bo",
        )
        if st.button("💾 Guardar datos de visitas"):
            guardar_visitas_bo(editado[["DNI", "Nombre", "PDV Totales", "Visitas PDV"]])
            st.success("Datos de visitas guardados.")
            st.rerun()

    st.markdown("#### 💰 Cálculo de comisiones (Prepago + Postpago)")
    st.caption(
        f"Comisión = Proy% (tope {TOPE_PROY_PCT_COMISION:.0%}) × Monto del producto "
        f"(Prepago S/ {MONTO_COMISION_PRODUCTO['Prepago']:,.0f} · Postpago S/ {MONTO_COMISION_PRODUCTO['Postpago']:,.0f}). "
        "Prepago se ajusta además por el % de visitas de cada gestor."
    )

    tabla = construir_tabla_comisiones(df, editado, dias_en_mes, dia_corte)

    n_gestores = df["DNI"].nunique()
    comision_total = tabla.loc["Fanero", ("Comisión", "Total")]
    col1, col2 = st.columns(2)
    col1.metric("Gestores", f"{n_gestores:,}")
    col2.metric("💰 Comisión total estimada", f"S/ {comision_total:,.0f}")

    renderizar_tabla_centrada(aplicar_estilo_comisiones(tabla), "content")
    st.caption("🟥 <80% · 🟨 80%–99% · 🟩 ≥100% (aplica a Proy %, con tope de 110%)")

    tabla_csv = tabla.copy()
    tabla_csv.columns = [f"{p} - {m}" for p, m in tabla_csv.columns]
    st.download_button(
        "⬇️ Descargar cálculo de comisiones (CSV)",
        data=tabla_csv.reset_index().to_csv(index=False).encode("utf-8"),
        file_name="calculo_comisiones_gestores.csv",
        mime="text/csv",
    )


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
      desagrupar=True). Internamente se agrupa por DNI (para no mezclar dos
      gestores distintos que compartan nombre), pero la fila se muestra
      SOLO con el Nombre, sin el DNI.

    Cuando NO se desagrupa, se agrega una fila "Fanero" al final
    con la suma de TODO lo que esté en df_filtrado (todos los departamentos
    o todos los gestores, según corresponda) — sus % se recalculan sobre
    los totales, no se promedian filas.
    """
    df_filtrado = df_filtrado.copy()
    df_filtrado["_PDV"] = np.where(
        df_filtrado["Nombre PDV"] != "",
        df_filtrado["PDV"] + " · " + df_filtrado["Nombre PDV"],
        df_filtrado["PDV"],
    )

    # Para Gestor, se agrupa por DNI (clave única) — el nombre se aplica
    # recién al final, solo para mostrar.
    nivel_col = "Departamento" if agrupar_por == "Departamento" else "DNI"

    # Al desagrupar por PDV, se agregan Departamento/Provincia/Distrito como
    # columnas descriptivas justo al lado de PDV (sin duplicar Departamento
    # si ya es el nivel principal de agrupación).
    columnas_geo = ["Provincia", "Distrito"] if agrupar_por == "Departamento" else ["Departamento", "Provincia", "Distrito"]
    index_cols = [nivel_col] + (["_PDV"] + columnas_geo if desagrupar else [])

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

        # --- Fila "Fanero": suma real de Cuota/Avance/Proy Unidades de
        # TODAS las filas, con Proy % recalculado sobre esos totales (no es
        # el promedio de los % de cada fila). Va al FINAL de la tabla. ---
        fila_total = {}
        for p in orden_prod:
            cuota_p = ancho[(p, "Cuota")].sum() if (p, "Cuota") in ancho.columns else 0.0
            avance_p = ancho[(p, "Avance")].sum() if (p, "Avance") in ancho.columns else 0.0
            proy_p = ancho[(p, "Proy Unidades")].sum() if (p, "Proy Unidades") in ancho.columns else 0.0
            fila_total[(p, "Cuota")] = cuota_p
            fila_total[(p, "Avance")] = avance_p
            fila_total[(p, "Proy Unidades")] = proy_p
            fila_total[(p, "Proy %")] = (proy_p / cuota_p) if cuota_p > 0 else 0.0
        df_total = pd.DataFrame([fila_total], index=["Fanero"], columns=columnas_orden)
        ancho = pd.concat([ancho, df_total])
    else:
        ancho = ancho.reindex(columns=columnas_orden).sort_index(level=0)

    nombre_nivel = "Departamento" if agrupar_por == "Departamento" else "Gestor"

    if agrupar_por == "Gestor":
        # Reemplaza el DNI (usado solo para agrupar bien) por el Nombre, que
        # es lo único que se muestra. "Fanero" (la fila total) se respeta tal cual.
        dni_a_nombre = df_filtrado.drop_duplicates("DNI").set_index("DNI")["Nombre"].to_dict()
        if desagrupar:
            nuevas_tuplas = [(dni_a_nombre.get(tup[0], tup[0]),) + tup[1:] for tup in ancho.index]
            nuevo_index = pd.MultiIndex.from_tuples(
                nuevas_tuplas, names=[nombre_nivel, "PDV"] + columnas_geo,
            )
            ancho.index = nuevo_index
            ancho = ancho.sort_index(level=0)
        else:
            es_total = ancho.index == "Fanero"
            nuevos_labels = [
                "Fanero" if es_total[i] else dni_a_nombre.get(idx, idx)
                for i, idx in enumerate(ancho.index)
            ]
            ancho.index = pd.Index(nuevos_labels, name=nombre_nivel)
            # Reordena alfabéticamente por Nombre, dejando "Fanero" al final.
            sin_total = ancho[ancho.index != "Fanero"].sort_index()
            con_total = ancho[ancho.index == "Fanero"]
            ancho = pd.concat([sin_total, con_total])
    else:
        if desagrupar:
            ancho.index = ancho.index.set_names([nombre_nivel, "PDV"] + columnas_geo)
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
    """Da formato/color a la tabla y la aplana (reset_index) antes de estilar:
    pandas Styler.apply/.map no funcionan con índices con valores repetidos
    (por ejemplo, dos gestores distintos que se llaman igual), así que el
    nombre del nivel (Departamento/Gestor, y PDV si aplica) pasa a ser una
    columna normal en vez de índice — el resultado visual es el mismo."""
    nombres_indice = list(tabla.index.names)  # ["Departamento"] o ["Gestor"] o [.., "PDV"]
    tabla = tabla.reset_index()
    columnas_indice = [(n, "") for n in nombres_indice]

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

    styler = tabla.style.format(fmt, na_rep="-").hide(axis="index")
    subset = [(p, "Proy %") for p in orden_prod]
    styler = _aplicar_semaforo(styler, subset)

    if tiene_comparativo:
        def _color_var(v):
            if pd.isna(v):
                return ""
            return "color: #3E9B4F; font-weight: 600" if v >= 0 else "color: #D64545; font-weight: 600"
        styler = styler.map(_color_var, subset=[("M0 vs M-1", "%Var")])

    # Fila "Fanero" (total general): sombreado ejecutivo distinto — fondo
    # azul marino oscuro con texto blanco, para que se note como el cierre
    # de la tabla y no se confunda con el semáforo de las demás filas.
    col_nivel_principal = columnas_indice[0]
    if (tabla[col_nivel_principal] == "Fanero").any():
        def _estilo_total(fila):
            if fila[col_nivel_principal] == "Fanero":
                return ["background-color: #00405E; color: #FFFFFF; font-weight: 700;" for _ in fila]
            return ["" for _ in fila]
        styler = styler.apply(_estilo_total, axis=1)

    # Bordes + centrado. Esto solo se ve completo cuando la tabla se
    # renderiza como HTML (ver `renderizar_tabla_centrada`) — el widget
    # interactivo st.dataframe NO respeta text-align de un Styler, por eso
    # esta tabla en particular se muestra con st.markdown(unsafe_allow_html).
    styler = styler.set_table_styles(
        [
            {"selector": "table", "props": [("border-collapse", "collapse"), ("width", "100%"), ("font-size", "0.85rem")]},
            {"selector": "th, td", "props": [("border", "1px solid #E2E4F0"), ("padding", "6px 10px"), ("text-align", "center"), ("white-space", "nowrap")]},
            {"selector": "th", "props": [("background-color", "#F5F6FB"), ("font-weight", "600")]},
        ],
        overwrite=True,
    )
    styler = styler.set_properties(**{"text-align": "center"})

    return styler


def renderizar_tabla_centrada(styler, altura: str | int = "content") -> None:
    """Renderiza un Styler como HTML puro (st.markdown unsafe_allow_html),
    en vez de st.dataframe — necesario para que el centrado de texto (y
    cualquier otro CSS del Styler) se vea de verdad, ya que el widget
    interactivo st.dataframe ignora text-align. Se pierde el buscador nativo
    y el ordenamiento por columna del widget interactivo; se mantiene la
    descarga a CSV por separado."""
    html = styler.to_html()
    if isinstance(altura, int):
        contenedor = f'<div style="overflow:auto; max-height:{altura}px; border:1px solid #E2E4F0; border-radius:6px;">{html}</div>'
    else:
        contenedor = f'<div style="overflow-x:auto;">{html}</div>'
    st.markdown(contenedor, unsafe_allow_html=True)


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


def procesar_carga_horizontal_diaria(
    df_horizontal: pd.DataFrame, producto: str, mes: int, anio: int, reemplazar: bool = False,
) -> tuple[int, list[int]]:
    """Convierte una plantilla HORIZONTAL (una columna por día del mes, con
    el nombre de columna = número de día: 1, 2, 3... 31) en publicaciones
    diarias reales, reutilizando exactamente la misma lógica de
    `publicar_datos_incremental` (suma al acumulado, archiva historial
    diario) — así el resultado es idéntico a haber publicado día por día,
    pero en un solo archivo.

    `reemplazar=True`: en vez de SUMAR el total de este archivo a lo ya
    publicado, lo REEMPLAZA para los DNI+PDV+Producto que traiga (sin
    tocar otros productos/PDV) — usar esto si vas a volver a subir el
    MISMO rango de días que ya habías cargado antes (para no duplicar).

    Columnas esperadas: DNI, Nombre, Departamento, Provincia, Distrito, PDV,
    Nombre PDV, y luego una columna por cada día con la venta de ESE día
    (no acumulada). La Cuota no viene en este archivo — se conserva la que
    ya esté publicada para ese DNI+PDV+Producto (si no existe ninguna
    todavía, queda en 0 y se puede corregir luego con la plantilla normal).

    Devuelve (cantidad_de_dias_procesados, lista_de_dias_omitidos_por_vacios).
    """
    columnas_identidad = ["DNI", "Nombre", "Departamento", "Provincia", "Distrito", "PDV", "Nombre PDV"]
    faltantes = [c for c in columnas_identidad if c not in df_horizontal.columns]
    if faltantes:
        raise ValueError("Faltan columnas de identidad: " + ", ".join(faltantes))

    columnas_dia = []
    for col in df_horizontal.columns:
        if col in columnas_identidad:
            continue
        try:
            numero_dia = int(col)
            if 1 <= numero_dia <= 31:
                columnas_dia.append((numero_dia, col))
        except (ValueError, TypeError):
            continue

    if not columnas_dia:
        raise ValueError(
            "No se encontró ninguna columna de día (deben llamarse 1, 2, 3... 31, "
            "una por cada día del mes)."
        )
    columnas_dia.sort(key=lambda t: t[0])

    df_horizontal = _normalizar_identidad(df_horizontal)

    # Cuota: se toma de lo ya publicado para ese DNI+PDV+Producto (si existe).
    cuota_existente = {}
    if os.path.exists(DATA_FILE):
        df_actual, _, _, _ = obtener_datos_publicados()
        df_actual_prod = df_actual[df_actual["Producto"] == producto]
        cuota_existente = {
            (fila["DNI"], fila["PDV"]): fila["Cuota"] for _, fila in df_actual_prod.iterrows()
        }

    dias_procesados = 0
    dias_omitidos = []
    entradas_lote = []
    filas_incremento_total = []

    for numero_dia, nombre_columna in columnas_dia:
        columnas_a_tomar = columnas_identidad + [nombre_columna]
        df_dia = df_horizontal[columnas_a_tomar].rename(columns={nombre_columna: "Avance"})
        df_dia["Avance"] = pd.to_numeric(df_dia["Avance"], errors="coerce")
        df_dia = df_dia.dropna(subset=["Avance"])

        if df_dia.empty:
            dias_omitidos.append(numero_dia)
            continue

        df_dia["Producto"] = producto
        df_dia["Cuota"] = df_dia.apply(
            lambda fila: cuota_existente.get((fila["DNI"], fila["PDV"]), 0.0), axis=1
        )

        ultimo_dia_mes = calendar.monthrange(anio, mes)[1]
        dia_corte_valido = min(numero_dia, ultimo_dia_mes)
        fecha_dia = pd.Timestamp(year=anio, month=mes, day=dia_corte_valido)
        entradas_lote.append((df_dia, fecha_dia))
        filas_incremento_total.append(df_dia)
        dias_procesados += 1

    if filas_incremento_total:
        # UNA sola publicación con la SUMA de todos los días (en vez de una
        # publicación por día — mucho más rápido, mismo resultado final ya
        # que sumar es asociativo). El detalle día por día para el gráfico
        # se registra aparte, también en un solo lote.
        df_incremento_total = pd.concat(filas_incremento_total, ignore_index=True)
        df_incremento_total = (
            df_incremento_total.groupby(
                ["DNI", "Nombre", "Departamento", "Provincia", "Distrito", "PDV", "Nombre PDV", "Producto"],
                as_index=False,
            ).agg(Avance=("Avance", "sum"), Cuota=("Cuota", "max"))
        )
        ultimo_dia_con_datos = max(dia for dia, _ in columnas_dia if dia not in dias_omitidos)
        funcion_publicar = publicar_datos_reemplazo_parcial if reemplazar else publicar_datos_incremental
        funcion_publicar(
            df_incremento_total, dia_corte=ultimo_dia_con_datos, mes=mes, anio=anio, registrar_historial=False,
        )
        registrar_incrementos_diarios_lote(entradas_lote)

    return dias_procesados, dias_omitidos


def procesar_carga_horizontal_historica(
    df_horizontal: pd.DataFrame, producto: str, mes: int, anio: int, df_referencia: pd.DataFrame,
) -> tuple[int, list[int]]:
    """Igual que `procesar_carga_horizontal_diaria`, pero para un MES
    ANTERIOR (histórico) — nunca toca `ultima_carga.xlsx` (los datos del
    mes en curso). Solo alimenta:
    1. `historial_diario` (día por día), para poder comparar M0 vs M-1
       "mismo día contra mismo día".
    2. El archivo mensual archivado (`data/historico/AAAA_MM.xlsx`), con
       el total del mes completo — para que la comparación de mes
       completo siga funcionando igual que antes.

    Columnas esperadas (acepta CUALQUIERA de estos 2 esquemas, con columnas
    de día en vez de una sola columna "Mes"):

        DNI PDV | Nombre PDV | Departamento | Provincia | Distrito | DNI Gestor | 1 | 2 | 3 | ... | 31
        (o, igual que la plantilla de ventas diarias del mes actual:)
        DNI | Nombre | Departamento | Provincia | Distrito | PDV | Nombre PDV | 1 | 2 | 3 | ... | 31

    Obligatorias: el identificador de PDV ("DNI PDV" o "PDV") y el
    identificador de Gestor ("DNI Gestor" o "DNI"), y al menos una columna
    de día. Departamento/Provincia/Distrito/Nombre PDV son opcionales pero
    muy recomendadas (mismo criterio que `procesar_carga_historico_ancho`).
    """
    # Acepta ambos esquemas de nombres de columna (evita que el usuario
    # tenga que recordar cuál plantilla usa qué nombres).
    col_pdv_origen = "DNI PDV" if "DNI PDV" in df_horizontal.columns else ("PDV" if "PDV" in df_horizontal.columns else None)
    col_gestor_origen = "DNI Gestor" if "DNI Gestor" in df_horizontal.columns else ("DNI" if "DNI" in df_horizontal.columns else None)

    faltantes_obligatorias = []
    if col_pdv_origen is None:
        faltantes_obligatorias.append("DNI PDV (o PDV)")
    if col_gestor_origen is None:
        faltantes_obligatorias.append("DNI Gestor (o DNI)")
    if faltantes_obligatorias:
        raise ValueError("Faltan columnas obligatorias: " + ", ".join(faltantes_obligatorias))

    df_horizontal = df_horizontal.copy()
    df_horizontal["PDV"] = df_horizontal[col_pdv_origen].astype(str).str.strip()
    df_horizontal["DNI"] = df_horizontal[col_gestor_origen].astype(str).str.strip()
    sin_gestor_conocido = df_horizontal["DNI"].isna() | (df_horizontal["DNI"].isin(["", "nan", "None"]))
    df_horizontal.loc[sin_gestor_conocido, "DNI"] = "SIN_GESTOR"

    columnas_geo_presentes = [c for c in ["Nombre PDV", "Departamento", "Provincia", "Distrito"] if c in df_horizontal.columns]
    nombre_gestor_presente = "Nombre" in df_horizontal.columns
    columnas_excluir = {"DNI PDV", "DNI Gestor", "PDV", "DNI", "Nombre"} | set(columnas_geo_presentes)

    columnas_dia = []
    for col in df_horizontal.columns:
        if col in columnas_excluir:
            continue
        try:
            numero_dia = int(col)
            ultimo_dia_mes = calendar.monthrange(anio, mes)[1]
            if 1 <= numero_dia <= ultimo_dia_mes:
                columnas_dia.append((numero_dia, col))
        except (ValueError, TypeError):
            continue

    if not columnas_dia:
        raise ValueError(
            "No se encontró ninguna columna de día válida para ese mes (deben llamarse "
            "1, 2, 3... según corresponda)."
        )
    columnas_dia.sort(key=lambda t: t[0])

    # Respaldo por DNI Gestor para geo faltante (mismo criterio que la carga
    # ancha de M-1) y Nombre del gestor con placeholder si no se encuentra.
    columnas_geo_faltantes = [c for c in ["Nombre PDV", "Departamento", "Provincia", "Distrito"] if c not in columnas_geo_presentes]
    columnas_respaldo = columnas_geo_faltantes + ([] if nombre_gestor_presente else ["Nombre"])
    ref_cols = [c for c in (["DNI"] + columnas_respaldo) if c in df_referencia.columns]
    ref = df_referencia[ref_cols].drop_duplicates(subset=["DNI"]) if "DNI" in ref_cols and len(ref_cols) > 1 else pd.DataFrame(columns=["DNI"])

    dias_procesados = 0
    dias_omitidos = []
    entradas_lote = []  # (df_dia, fecha) de cada día — se registran TODOS juntos al final

    for numero_dia, nombre_columna in columnas_dia:
        columnas_base = ["DNI", "PDV"] + columnas_geo_presentes + (["Nombre"] if nombre_gestor_presente else [])
        columnas_a_tomar = columnas_base + [nombre_columna]
        df_dia = df_horizontal[columnas_a_tomar].rename(columns={nombre_columna: "Avance"})
        df_dia["Avance"] = pd.to_numeric(df_dia["Avance"], errors="coerce")
        df_dia = df_dia.dropna(subset=["Avance"])

        if df_dia.empty:
            dias_omitidos.append(numero_dia)
            continue

        df_dia["Producto"] = producto
        if not ref.empty:
            df_dia = df_dia.merge(ref, on="DNI", how="left")

        if "Nombre" not in df_dia.columns:
            df_dia["Nombre"] = ""
        df_dia["Nombre"] = df_dia["Nombre"].fillna("")
        es_sin_gestor = df_dia["DNI"] == "SIN_GESTOR"
        df_dia.loc[es_sin_gestor & (df_dia["Nombre"] == ""), "Nombre"] = "Sin gestor asignado"
        sin_nombre = df_dia["Nombre"] == ""
        df_dia.loc[sin_nombre, "Nombre"] = "Gestor " + df_dia.loc[sin_nombre, "DNI"]

        df_dia = _normalizar_identidad(df_dia)

        fecha_dia = pd.Timestamp(year=anio, month=mes, day=numero_dia)
        entradas_lote.append((df_dia, fecha_dia))
        dias_procesados += 1

    if entradas_lote:
        # UNA sola lectura + UNA sola escritura del archivo para TODOS los
        # días juntos (en vez de una por día) — mucho más rápido con
        # muchos días/PDV.
        registrar_incrementos_diarios_lote(entradas_lote)

    if dias_procesados > 0:
        # El total archivado se recalcula del HISTORIAL DIARIO COMPLETO de
        # este mes — de TODOS los productos ya cargados, no solo el de esta
        # llamada — para no perder lo ya archivado si el mes se sube en
        # varias partes (por día y/o por producto, en cargas separadas).
        historial_completo = obtener_historial_diario()
        historial_mes = historial_completo[
            (historial_completo["Fecha"].dt.year == anio) & (historial_completo["Fecha"].dt.month == mes)
        ]
        if not historial_mes.empty:
            df_mes_completo = historial_mes.groupby(
                [c for c in ["DNI", "Nombre", "Departamento", "PDV", "Producto"] if c in historial_mes.columns],
                as_index=False,
            )["Avance"].sum()
            _archivar_mes(df_mes_completo, mes, anio)

    return dias_procesados, dias_omitidos


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
        "(úsalo solo para corregir un error de carga, o si tu archivo trae el acumulado total del mes)."
    )

    if archivo is not None:
        df_preview = cargar_datos_excel(archivo)
        if df_preview is not None and not forzar_reemplazo:
            avance_nuevo = df_preview["Avance"].sum()
            avance_actual = 0.0
            if os.path.exists(DATA_FILE):
                df_actual_preview, _, mes_actual_preview, anio_actual_preview = obtener_datos_publicados()
                if int(mes_actual_preview) == int(mes_sel) and int(anio_actual_preview) == int(anio_sel):
                    avance_actual = df_actual_preview["Avance"].sum()
            st.info(
                f"📊 **Vista previa:** este archivo suma **{avance_nuevo:,.0f}** unidades de Avance. "
                f"En modo SUMAR (el de arriba, sin marcar), el acumulado del mes pasaría de "
                f"**{avance_actual:,.0f}** a **{avance_actual + avance_nuevo:,.0f}**.\n\n"
                "⚠️ Si tu archivo trae el **acumulado total hasta hoy** (no solo lo vendido en el día/periodo "
                "más reciente), esto va a duplicar tus ventas — en ese caso, marca la casilla de REEMPLAZAR "
                "arriba, o pon Avance = 0 si solo quieres actualizar la Cuota."
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
        "un mes completo por sí sola. Sube una tabla con una fila por PDV, el "
        "Mes en texto, y una columna por producto con el total vendido ese mes:"
    )
    st.code(
        "DNI PDV | Nombre PDV | Departamento | Provincia | Distrito | DNI Gestor | Mes | "
        "Prepago | Porta Prepago | Postpago | OSS",
        language=None,
    )
    st.caption(
        "Obligatorias: **DNI PDV**, **DNI Gestor**, **Mes**, y al menos un producto. "
        "Nombre PDV/Departamento/Provincia/Distrito son opcionales pero muy recomendadas: "
        "si las incluyes, cada PDV queda ubicado correctamente sin depender de nada más. "
        "Si las omites, la app intenta completarlas buscando el DNI Gestor en lo que ya "
        "esté publicado este mes — y si ese gestor todavía no está publicado, esas filas "
        "quedan sin departamento (no se van a poder comparar por departamento, aunque sí "
        "cuentan igual si agrupas por Gestor)."
    )
    st.caption(
        "Los nombres de columna de producto deben coincidir exactamente con "
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
            df_ancho = leer_excel_seguro(archivo_historico)
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

    st.markdown("---")
    st.markdown("#### 📅 Cargar ventas diarias (formato horizontal)")
    st.caption(
        "Alternativa a subir un archivo por día: una sola tabla con una columna por cada día del "
        "mes (nombradas 1, 2, 3... 31), con la venta de ESE día en cada celda (no acumulada). "
        "La app publica automáticamente día por día, en orden, igual que si hubieras subido un "
        "archivo distinto cada día."
    )
    st.code("DNI | Nombre | Departamento | Provincia | Distrito | PDV | Nombre PDV | 1 | 2 | 3 | ... | 31", language=None)
    st.caption("La Cuota no va en este archivo: se conserva la que ya esté publicada para cada PDV+Producto.")

    col_prod_h, col_mes_h, col_anio_h = st.columns(3)
    with col_prod_h:
        producto_horizontal = st.selectbox("Producto de este archivo", PRODUCTOS, key="producto_horizontal")
    with col_mes_h:
        mes_horizontal = st.number_input("Mes", min_value=1, max_value=12, value=datetime.now().month, key="mes_horizontal")
    with col_anio_h:
        anio_horizontal = st.number_input(
            "Año", min_value=2020, max_value=2100, value=datetime.now().year, step=1, key="anio_horizontal",
        )

    archivo_horizontal = st.file_uploader(
        "Excel de ventas diarias (formato horizontal)", type=["xlsx"], key="uploader_horizontal_diario",
    )
    reemplazar_horizontal = st.checkbox(
        "⚠️ Reemplazar en vez de sumar (usa esto si vas a volver a subir el MISMO rango de días "
        "que ya habías cargado antes, para no duplicar)",
        key="reemplazar_horizontal",
    )

    if archivo_horizontal is not None and st.button("Procesar y publicar ventas diarias"):
        try:
            df_horizontal = leer_excel_seguro(archivo_horizontal)
        except Exception as exc:  # noqa: BLE001 - se informa al usuario cualquier error de lectura
            st.error(f"No se pudo leer el archivo: {exc}")
            df_horizontal = None

        if df_horizontal is not None:
            try:
                dias_ok, dias_omitidos = procesar_carga_horizontal_diaria(
                    df_horizontal, producto_horizontal, int(mes_horizontal), int(anio_horizontal),
                    reemplazar=reemplazar_horizontal,
                )
            except ValueError as exc:
                st.error(str(exc))
                dias_ok, dias_omitidos = 0, []

            if dias_ok > 0:
                verbo = "reemplazaron" if reemplazar_horizontal else "sumaron"
                mensaje = f"Se {verbo} {dias_ok} día(s) de {producto_horizontal} para {int(mes_horizontal)}/{int(anio_horizontal)}."
                if dias_omitidos:
                    mensaje += f" Días sin datos (omitidos): {', '.join(map(str, dias_omitidos))}."
                st.success(mensaje)

    st.markdown("---")
    st.markdown("#### 📅➕ Cargar ventas diarias de un MES ANTERIOR (histórico)")
    st.caption(
        "Para poder comparar M0 vs M-1 'mismo día contra mismo día' en Vista Gerencial, en vez del "
        "total del mes completo. Esta carga NO toca los datos del mes en curso — solo alimenta el "
        "histórico. Acepta CUALQUIERA de estos 2 formatos de columnas (uses el que ya tengas armado):"
    )
    st.code(
        "DNI | Nombre | Departamento | Provincia | Distrito | PDV | Nombre PDV | 1 | 2 | 3 | ... | 31\n"
        "DNI PDV | Nombre PDV | Departamento | Provincia | Distrito | DNI Gestor | 1 | 2 | 3 | ... | 31",
        language=None,
    )
    st.caption(
        "Obligatorias: el identificador de PDV (**PDV** o **DNI PDV**) y el de Gestor "
        "(**DNI** o **DNI Gestor**), y al menos una columna de día. "
        "Nombre/Nombre PDV/Departamento/Provincia/Distrito son opcionales pero muy recomendadas — si "
        "las omites, se completan buscando el DNI del gestor en lo ya publicado este mes (y si "
        "tampoco se encuentra ahí, quedan vacías)."
    )

    col_prod_hh, col_mes_hh, col_anio_hh = st.columns(3)
    with col_prod_hh:
        producto_hist_horizontal = st.selectbox("Producto de este archivo", PRODUCTOS, key="producto_hist_horizontal")
    with col_mes_hh:
        mes_hist_horizontal = st.number_input(
            "Mes (el mes anterior, no el actual)", min_value=1, max_value=12, value=datetime.now().month,
            key="mes_hist_horizontal",
        )
    with col_anio_hh:
        anio_hist_horizontal = st.number_input(
            "Año", min_value=2020, max_value=2100, value=datetime.now().year, step=1, key="anio_hist_horizontal",
        )

    archivo_hist_horizontal = st.file_uploader(
        "Excel de ventas diarias históricas (formato horizontal)", type=["xlsx"], key="uploader_hist_horizontal",
    )

    if archivo_hist_horizontal is not None and st.button("Procesar histórico diario"):
        df_referencia_hist, _, mes_actual_check, anio_actual_check = obtener_datos_publicados()
        if int(mes_hist_horizontal) == int(mes_actual_check) and int(anio_hist_horizontal) == int(anio_actual_check):
            st.error(
                f"⚠️ Seleccionaste {int(mes_hist_horizontal)}/{int(anio_hist_horizontal)}, pero ese es el MES "
                "ACTUAL (el que ya está publicado en Vista Gerencial), no un mes anterior. Si subes aquí el mes "
                "actual, la comparación M0 vs M-1 se distorsiona. Cambia el Mes/Año arriba al mes que sí quieres "
                "guardar como histórico (ej. si agosto es el actual, aquí va julio)."
            )
        else:
            try:
                df_hist_horizontal = leer_excel_seguro(archivo_hist_horizontal)
            except Exception as exc:  # noqa: BLE001 - se informa al usuario cualquier error de lectura
                st.error(f"No se pudo leer el archivo: {exc}")
                df_hist_horizontal = None

            if df_hist_horizontal is not None:
                try:
                    dias_ok_h, dias_omitidos_h = procesar_carga_horizontal_historica(
                        df_hist_horizontal, producto_hist_horizontal,
                        int(mes_hist_horizontal), int(anio_hist_horizontal), df_referencia_hist,
                    )
                except ValueError as exc:
                    st.error(str(exc))
                    dias_ok_h, dias_omitidos_h = 0, []

                if dias_ok_h > 0:
                    mensaje_h = (
                        f"Se guardó el histórico diario de {producto_hist_horizontal} para "
                        f"{int(mes_hist_horizontal)}/{int(anio_hist_horizontal)} ({dias_ok_h} día(s)). "
                        "Los datos del mes en curso NO se modificaron."
                    )
                    if dias_omitidos_h:
                        mensaje_h += f" Días sin datos (omitidos): {', '.join(map(str, dias_omitidos_h))}."
                    st.success(mensaje_h)

    st.markdown("---")
    with st.expander("🗑️ Borrar TODOS los datos (reiniciar la app desde cero)"):
        st.warning(
            "Esto borra TODO lo publicado: el mes actual, el histórico de meses anteriores, "
            "el detalle diario y los datos de visitas del BO. No se puede deshacer. Úsalo solo "
            "si quieres empezar de cero (por ejemplo, para hacer pruebas limpias)."
        )
        confirmar_borrado = st.checkbox("Sí, entiendo que esto borra todo y no se puede deshacer", key="confirmar_borrado_total")
        if st.button("🗑️ Borrar todo y reiniciar", disabled=not confirmar_borrado):
            archivos_a_borrar = [DATA_FILE, DATA_META, VISITAS_FILE, HISTORIAL_DIARIO_FILE]
            borrados = []
            for ruta in archivos_a_borrar:
                if os.path.exists(ruta):
                    os.remove(ruta)
                    borrados.append(os.path.basename(ruta))
            if os.path.isdir(HISTORICO_DIR):
                for nombre_archivo in os.listdir(HISTORICO_DIR):
                    os.remove(os.path.join(HISTORICO_DIR, nombre_archivo))
                borrados.append("historico/*")
            _leer_excel_publicado.clear()
            _leer_historial_diario_cacheado.clear()
            generar_datos_ejemplo.clear()
            st.success(f"Listo, se borró todo: {', '.join(borrados) if borrados else '(no había nada que borrar)'}. Recarga la página.")


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
            df_subido = leer_excel_seguro(archivo_avances)
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


def tabla_comparativo_mensual(
    df_filtrado: pd.DataFrame, mes: int, anio: int, agrupar_por: str, dia_corte: int | None = None,
) -> tuple[pd.DataFrame, bool]:
    """Construye la tabla M0 (venta Prepago del mes actual) vs M-1 (venta
    Prepago del mismo Departamento/Gestor el mes anterior) y %Var.

    Si hay detalle diario del mes anterior (se cargó con la plantilla
    horizontal histórica) y se pasa `dia_corte`, M-1 se calcula "mismo día
    contra mismo día" (acumulado del mes anterior hasta ese mismo número de
    día) — la comparación más justa. Si no hay detalle diario, cae de
    vuelta al total del mes anterior completo (el comportamiento anterior).

    Devuelve (tabla, hay_historico) — hay_historico=False solo si no hay
    NINGÚN dato del mes anterior (ni detalle diario ni total archivado).

    Para Gestor, se agrupa internamente por DNI (clave única, evita mezclar
    dos gestores con el mismo nombre) y el resultado se muestra con el
    Nombre — igual criterio que `construir_tabla_producto`, para que ambas
    tablas se puedan unir (join) por el mismo índice."""
    anio_ant, mes_ant = (anio - 1, 12) if mes == 1 else (anio, mes - 1)

    clave_col = "Departamento" if agrupar_por == "Departamento" else "DNI"
    nombre_indice = "Departamento" if agrupar_por == "Departamento" else "Gestor"

    df_prepago = df_filtrado[df_filtrado["Producto"] == "Prepago"].copy()
    m0 = df_prepago.groupby(clave_col)["Avance"].sum().rename("M0")

    usar_mismo_dia = dia_corte is not None and hay_detalle_diario_del_mes(mes_ant, anio_ant)

    if usar_mismo_dia:
        historial = obtener_historial_diario()
        filtro = (
            (historial["Fecha"].dt.year == anio_ant) & (historial["Fecha"].dt.month == mes_ant)
            & (historial["Fecha"].dt.day <= dia_corte) & (historial["Producto"] == "Prepago")
        )
        m1 = historial[filtro].groupby(clave_col)["Avance"].sum().rename("M-1")
        df_hist = obtener_historico_mes(mes_ant, anio_ant)  # solo para completar Nombre de Gestor si falta
    else:
        df_hist = obtener_historico_mes(mes_ant, anio_ant)
        if df_hist is not None:
            df_prepago_hist = df_hist[df_hist["Producto"] == "Prepago"].copy()
            m1 = df_prepago_hist.groupby(clave_col)["Avance"].sum().rename("M-1")
        else:
            m1 = None

    if m1 is None:
        comparativo = m0.to_frame()
        comparativo["M-1"] = np.nan
        comparativo["%Var"] = np.nan
    else:
        comparativo = pd.concat([m0, m1], axis=1)
        comparativo["M0"] = comparativo["M0"].fillna(0)
        comparativo["M-1"] = comparativo["M-1"].fillna(0)
        comparativo["%Var"] = np.where(
            comparativo["M-1"] > 0,
            (comparativo["M0"] - comparativo["M-1"]) / comparativo["M-1"],
            np.where(comparativo["M0"] > 0, 1.0, 0.0),
        )

    if agrupar_por == "Gestor":
        dni_a_nombre = df_filtrado.drop_duplicates("DNI").set_index("DNI")["Nombre"].to_dict()
        if df_hist is not None:
            faltantes = [i for i in comparativo.index if i not in dni_a_nombre]
            if faltantes:
                mapa_extra = df_hist.drop_duplicates("DNI").set_index("DNI")["Nombre"].to_dict()
                for dni_faltante in faltantes:
                    if dni_faltante in mapa_extra:
                        dni_a_nombre[dni_faltante] = mapa_extra[dni_faltante]
        comparativo.index = [dni_a_nombre.get(i, i) for i in comparativo.index]

    comparativo.index.name = nombre_indice
    return comparativo.sort_values("M0", ascending=False), (m1 is not None)


def tabla_comparativo_mensual_pdv(
    df_filtrado: pd.DataFrame, mes: int, anio: int, agrupar_por: str, dia_corte: int | None = None,
) -> tuple[pd.DataFrame, bool]:
    """Versión a nivel PDV de `tabla_comparativo_mensual`: M0 vs M-1 de
    Prepago, pero una fila por (nivel, PDV) en vez de una fila por nivel.

    Se cruza por el código de PDV (columna "_PDV" con formato "código ·
    nombre líder", igual que en `construir_tabla_producto`), NO por
    Departamento ni por Gestor — así:
    - Un PDV que sigue existiendo este mes muestra su M-1 real, aunque
      haya cambiado de Gestor o Departamento entre un mes y otro.
    - Un PDV NUEVO este mes (no existía en el histórico) queda con M-1 = 0
      (no vacío — sí es sabido que no vendió nada el mes pasado, porque no
      existía).
    Los PDV que existían el mes pasado pero ya no siguen este mes
    simplemente no aparecen aquí (esta tabla se arma sobre los PDV
    ACTIVOS de este mes) — sí se cuentan igual en el M-1 departamental,
    ver `tabla_comparativo_mensual`.

    Si hay detalle diario del mes anterior (plantilla horizontal
    histórica) y se pasa `dia_corte`, M-1 se calcula "mismo día contra
    mismo día" por PDV — si no, cae de vuelta al total del mes anterior.
    """
    anio_ant, mes_ant = (anio - 1, 12) if mes == 1 else (anio, mes - 1)

    nivel_col = "Departamento" if agrupar_por == "Departamento" else "DNI"
    nombre_nivel = "Departamento" if agrupar_por == "Departamento" else "Gestor"

    def _con_clave_pdv(df_prod):
        df_prod = df_prod.copy()
        df_prod["_PDV"] = np.where(
            df_prod["Nombre PDV"] != "",
            df_prod["PDV"] + " · " + df_prod["Nombre PDV"],
            df_prod["PDV"],
        )
        return df_prod

    df_prepago = _con_clave_pdv(df_filtrado[df_filtrado["Producto"] == "Prepago"])
    m0 = df_prepago.groupby([nivel_col, "_PDV"])["Avance"].sum().rename("M0")

    usar_mismo_dia = dia_corte is not None and hay_detalle_diario_del_mes(mes_ant, anio_ant)

    if usar_mismo_dia:
        m1_por_pdv_codigo = avance_acumulado_hasta_dia_por_pdv(mes_ant, anio_ant, dia_corte, "Prepago")
        comparativo = m0.to_frame()
        comparativo["M-1"] = [
            m1_por_pdv_codigo.get(pdv.split(" · ")[0], 0.0) for _, pdv in comparativo.index
        ]
        comparativo["%Var"] = np.where(
            comparativo["M-1"] > 0,
            (comparativo["M0"] - comparativo["M-1"]) / comparativo["M-1"],
            np.where(comparativo["M0"] > 0, 1.0, 0.0),
        )
        hay_historico = True
        df_hist = obtener_historico_mes(mes_ant, anio_ant)  # solo para completar Nombre de Gestor
    else:
        df_hist = obtener_historico_mes(mes_ant, anio_ant)
        if df_hist is None:
            comparativo = m0.to_frame()
            comparativo["M-1"] = np.nan
            comparativo["%Var"] = np.nan
            hay_historico = False
        else:
            df_prepago_hist = _con_clave_pdv(df_hist[df_hist["Producto"] == "Prepago"])
            m1_por_pdv = df_prepago_hist.groupby("_PDV")["Avance"].sum()

            # El M-1 se busca SOLO por código de PDV (no por nivel): así, si un
            # PDV cambió de Gestor/Departamento entre un mes y otro, igual se
            # encuentra su venta real del mes pasado.
            comparativo = m0.to_frame()
            comparativo["M-1"] = [m1_por_pdv.get(pdv, 0.0) for _, pdv in comparativo.index]
            comparativo["%Var"] = np.where(
                comparativo["M-1"] > 0,
                (comparativo["M0"] - comparativo["M-1"]) / comparativo["M-1"],
                np.where(comparativo["M0"] > 0, 1.0, 0.0),
            )
            hay_historico = True

    if agrupar_por == "Gestor":
        dni_a_nombre = df_filtrado.drop_duplicates("DNI").set_index("DNI")["Nombre"].to_dict()
        comparativo.index = pd.MultiIndex.from_tuples(
            [(dni_a_nombre.get(dni, dni), pdv) for dni, pdv in comparativo.index],
            names=[nombre_nivel, "PDV"],
        )
    else:
        comparativo.index = comparativo.index.set_names([nombre_nivel, "PDV"])

    return comparativo, hay_historico


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

    # Fila de 3 filtros adicionales: DNI del líder (PDV), Provincia, Distrito.
    col_e, col_f, col_g = st.columns(3)
    with col_e:
        opciones_pdv = sorted(
            (df["PDV"] + " · " + df["Nombre PDV"]).unique(),
        )
        pdv_filtro_sel = st.multiselect(
            "DNI del líder (PDV)", options=opciones_pdv, default=[], key="pdv_filtro_gerencial",
        )
    with col_f:
        provincia_filtro_sel = st.multiselect(
            "Provincia", options=sorted(df["Provincia"].unique()), default=[], key="provincia_filtro_gerencial",
        )
    with col_g:
        distrito_filtro_sel = st.multiselect(
            "Distrito", options=sorted(df["Distrito"].unique()), default=[], key="distrito_filtro_gerencial",
        )

    if not productos_sel:
        productos_sel = PRODUCTOS
    departamentos_activos = departamentos_filtro if departamentos_filtro else sorted(df["Departamento"].unique())

    # Aplica los 3 filtros adicionales a TODO lo que sigue (KPIs, tabla,
    # M0/M-1 y gráfico) — igual que el resto de filtros de esta vista.
    if pdv_filtro_sel:
        codigos_pdv_filtro = {v.split(" · ")[0] for v in pdv_filtro_sel}
        df = df[df["PDV"].isin(codigos_pdv_filtro)]
    if provincia_filtro_sel:
        df = df[df["Provincia"].isin(provincia_filtro_sel)]
    if distrito_filtro_sel:
        df = df[df["Distrito"].isin(distrito_filtro_sel)]

    if df.empty:
        st.info("No hay datos para la combinación de filtros seleccionada (PDV/Provincia/Distrito).")
        return

    # M0 / M-1 son SIEMPRE sobre Prepago, independientemente del filtro de
    # Producto. Se calculan aquí y se agregan como columnas adicionales al
    # final de la tabla principal (después de OSS), no como tarjeta aparte.
    anio_ant, mes_ant = (anio - 1, 12) if mes == 1 else (anio, mes - 1)

    m0_total = df[(df["Departamento"].isin(departamentos_activos)) & (df["Producto"] == "Prepago")]["Avance"].sum()

    if hay_detalle_diario_del_mes(mes_ant, anio_ant):
        # "Mismo día contra mismo día": la comparación más justa, disponible
        # cuando el mes anterior se cargó con la plantilla horizontal histórica.
        m1_total = avance_acumulado_hasta_dia(mes_ant, anio_ant, dia_corte, departamentos_activos, "Prepago")
        var_total_pct = ((m0_total - m1_total) / m1_total) if m1_total > 0 else (1.0 if m0_total > 0 else 0.0)
    else:
        df_hist_mes_ant = obtener_historico_mes(mes_ant, anio_ant)
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
    # de OSS), como un grupo más "M0 vs M-1". Sin desagrupar: una fila por
    # Departamento/Gestor (+ "Fanero"). Desagrupado: una fila por PDV,
    # cruzando por código de PDV (no por nivel) — un PDV nuevo este mes
    # queda con M-1 = 0; uno que ya no sigue simplemente no aparece aquí. ---
    if not desagrupar:
        df_scope_depto = df[df["Departamento"].isin(departamentos_activos)]
        comparativo, hay_historico = tabla_comparativo_mensual(df_scope_depto, mes, anio, agrupar_por, dia_corte)
        fila_total_comp = pd.DataFrame(
            [{"M0": m0_total, "M-1": m1_total, "%Var": var_total_pct}], index=["Fanero"]
        )
        comparativo = pd.concat([fila_total_comp, comparativo])
        comparativo.columns = pd.MultiIndex.from_tuples([("M0 vs M-1", c) for c in comparativo.columns])
        tabla = tabla.join(comparativo, how="left")
    else:
        df_scope_depto = df[df["Departamento"].isin(departamentos_activos)]
        comparativo, hay_historico = tabla_comparativo_mensual_pdv(df_scope_depto, mes, anio, agrupar_por, dia_corte)
        comparativo.columns = pd.MultiIndex.from_tuples([("M0 vs M-1", c) for c in comparativo.columns])
        tabla = tabla.join(comparativo, how="left")
        # Un PDV nuevo este mes (sin fila en el histórico) debe verse como
        # M-1 = 0, no vacío — solo queda vacío si NO hay histórico del todo.
        if hay_historico:
            tabla[("M0 vs M-1", "M-1")] = tabla[("M0 vs M-1", "M-1")].fillna(0.0)
            tabla[("M0 vs M-1", "%Var")] = tabla[("M0 vs M-1", "%Var")].fillna(1.0)

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
        # La fila "Fanero" nunca se oculta por este filtro — sirve
        # de contexto siempre visible, sin importar qué rango se elija.
        mascara = rango_fila.isin(rango_sel)
        if "Fanero" in tabla.index:
            mascara.loc["Fanero"] = True
        tabla = tabla[mascara]

    if tabla.empty:
        st.info("No hay filas para el Rango de Cumplimiento seleccionado.")
    else:
        altura_render = 480 if desagrupar else "content"
        renderizar_tabla_centrada(aplicar_estilo_resumen_producto(tabla, orden_prod_sel), altura_render)
    leyenda = "🟥 <80% · 🟨 80%–99% · 🟩 ≥100% (aplica a Proy %)"
    anio_ant_leyenda, mes_ant_leyenda = (anio - 1, 12) if mes == 1 else (anio, mes - 1)
    if hay_detalle_diario_del_mes(mes_ant_leyenda, anio_ant_leyenda):
        leyenda += f" · M0 = Prepago hasta el día {dia_corte} este mes · M-1 = Prepago hasta el MISMO día ({dia_corte}) el mes anterior"
    else:
        leyenda += " · M0 = venta Prepago este mes · M-1 = venta Prepago TOTAL del mes anterior (sin detalle diario, no es 'mismo día')"
    if desagrupar:
        leyenda += " (por PDV: nuevo este mes = M-1 en 0)"
    st.caption(leyenda)
    if not hay_historico:
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

    st.markdown("---")
    st.markdown("#### 📈 Ventas diarias")

    producto_base_grafico = producto_base_para_rango(productos_sel)
    anio_ant_grafico, mes_ant_grafico = (anio - 1, 12) if mes == 1 else (anio, mes - 1)
    NOMBRES_MES_REV = {7: "Julio", 8: "Agosto", 9: "Setiembre", 1: "Enero", 2: "Febrero", 3: "Marzo",
                        4: "Abril", 5: "Mayo", 6: "Junio", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"}
    nombre_mes_actual = NOMBRES_MES_REV.get(mes, str(mes))
    nombre_mes_anterior = NOMBRES_MES_REV.get(mes_ant_grafico, str(mes_ant_grafico))
    periodo_actual = f"{nombre_mes_actual} {anio}"
    periodo_anterior = f"{nombre_mes_anterior} {anio_ant_grafico}"

    st.caption(
        f"Muestra **{producto_base_grafico}** (si tienes los 4 productos elegidos, se grafica Prepago por "
        "defecto — elige un solo producto distinto arriba para ver ese en su lugar), comparando **día del mes** "
        f"contra el mismo día del mes anterior ({periodo_actual} vs {periodo_anterior}, línea punteada) — no "
        "fecha calendario, para poder superponer ambos meses. Se arma a partir de lo que se suma en cada "
        "publicación del admin: necesita al menos 2 publicaciones en días distintos para verse."
    )

    historial = obtener_historial_diario()
    col_grupo = "Departamento" if agrupar_por == "Departamento" else "Nombre"

    def _filtrar_mes(hist, m, a):
        if hist.empty:
            return hist
        return hist[
            (hist["Fecha"].dt.year == a) & (hist["Fecha"].dt.month == m)
            & (hist["Departamento"].isin(departamentos_activos))
            & (hist["Producto"] == producto_base_grafico)
        ]

    historial_actual = _filtrar_mes(historial, mes, anio)
    historial_anterior = _filtrar_mes(historial, mes_ant_grafico, anio_ant_grafico)

    if historial_actual.empty and historial_anterior.empty:
        st.info("Todavía no hay suficientes publicaciones registradas para graficar ventas diarias.")
    else:
        piezas = []
        if not historial_actual.empty:
            datos_actual = historial_actual.groupby(["Fecha", col_grupo], as_index=False)["Avance"].sum()
            datos_actual["Día"] = datos_actual["Fecha"].dt.day
            datos_actual["Periodo"] = periodo_actual
            piezas.append(datos_actual)
        if not historial_anterior.empty:
            datos_anterior = historial_anterior.groupby(["Fecha", col_grupo], as_index=False)["Avance"].sum()
            datos_anterior["Día"] = datos_anterior["Fecha"].dt.day
            datos_anterior["Periodo"] = periodo_anterior
            piezas.append(datos_anterior)

        datos_grafico = pd.concat(piezas, ignore_index=True)
        datos_grafico = datos_grafico.rename(columns={col_grupo: agrupar_por})

        grafico = (
            alt.Chart(datos_grafico)
            .mark_line(strokeWidth=3.5, point=alt.OverlayMarkDef(size=40))
            .encode(
                x=alt.X("Día:O", title="Día del mes"),
                y=alt.Y("Avance:Q", title=f"Venta diaria ({producto_base_grafico})", scale=alt.Scale(zero=True)),
                color=alt.Color(f"{agrupar_por}:N", title=agrupar_por),
                strokeDash=alt.StrokeDash("Periodo:N", title="Periodo", sort=[periodo_actual, periodo_anterior]),
                tooltip=[
                    alt.Tooltip("Periodo:N", title="Periodo"),
                    alt.Tooltip("Día:O", title="Día"),
                    alt.Tooltip(f"{agrupar_por}:N", title=agrupar_por),
                    alt.Tooltip("Avance:Q", title="Venta", format=",.0f"),
                ],
            )
            .properties(height=420)
        )
        st.altair_chart(grafico, width="stretch")
        if historial_anterior.empty:
            st.caption(
                f"Solo se ve {periodo_actual} — todavía no hay detalle diario de {periodo_anterior} "
                "(cárgalo en el panel admin, sección 'Cargar ventas diarias de un MES ANTERIOR')."
            )
        st.caption(f"Ventas registradas por publicación, agrupadas por {agrupar_por.lower()} — no es un conteo transaccional día por día, sino lo sumado en cada carga del admin.")


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

    # Logo arriba de todo, ANTES de las pestañas — así queda visible sin
    # importar en cuál pestaña esté parado el usuario.
    LOGO_PATH = os.path.join(os.path.dirname(__file__), "assets", "logo_fanero.jpg")
    col_logo, col_titulo = st.columns([1, 5])
    with col_logo:
        if os.path.exists(LOGO_PATH):
            st.image(LOGO_PATH, width=140)
        else:
            # Avisa en vez de fallar en silencio — suele pasar si al subir
            # el proyecto se copió app.py pero no la carpeta assets/.
            st.caption("⚠️ Logo no encontrado (falta la carpeta `assets/` en el repositorio).")
    with col_titulo:
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

    tabs = st.tabs(["🏢 Vista Gerencial", "💰 Cálculo Comisión Gestor", "✏️ Editar Avances"])

    with tabs[0]:
        vista_gerencial(df, dias_en_mes, dia_corte, mes, anio)

    with tabs[1]:
        vista_comisiones_gestores(df, dias_en_mes, dia_corte)

    with tabs[2]:
        st.subheader("Editar Avances")
        panel_editar_avances(df_raw)


if __name__ == "__main__":
    main()
