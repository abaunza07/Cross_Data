"""
cruce_fuid_servidor.py

Cruza un Excel de FUID (por ej. "ACTA1 ENTREGABLE PROYECTOS DE LEY 1.xlsx")
contra los archivos reales encontrados en una carpeta del servidor
(recursivo), usando como llave la columna "NOMBRE DE ARCHIVO" del Excel
contra el nombre real del archivo (con extensión) en el servidor.

El CSV de salida tiene las 165 columnas del esquema de clasificación de
metadatos (Clasificacion_Metadatos_AGN.xlsx). Cada una se llena así:

  1. FUID   -> si existe una columna equivalente en el Excel de la acta,
               se copia su valor tal cual.
  2. TECH   -> se calcula escaneando el archivo real en el servidor
               (Volumen, Ruta, Tamaño, MIME/Extensión real por firma
               binaria, etc.)
  3. EXISTENTE -> campos que ya se diligencian por fuera (Hash, TipoHash,
               Linea, Fase, Dependencia). SE COPIAN TAL CUAL SI EXISTEN
               EN EL EXCEL. NUNCA SE RECALCULAN NI SE GENERAN DESDE EL
               ARCHIVO (por ejemplo, el Hash jamás se calcula con
               hashlib sobre el archivo real).
  4. CALC   -> una derivación simple y segura (ej: AÑO desde FECHA
               INICIAL, o cuál casilla de Soporte marcar según el valor
               de SOPORTE).
  5. Si no hay ninguna fuente de datos para ese campo -> "N/A".

Cualquier columna del Excel FUID que no se haya usado en el cruce anterior
se agrega al final del CSV con el prefijo "EXTRA_FUID__" para que no se
pierda información (ej: FASE, DEPENDENCIA, AG, AC, DP ELIMINAR, etc.)

Uso:
    cruce_fuid_servidor.exe --excel "ACTA1 ENTREGABLE PROYECTOS DE LEY 1.xlsx" ^
                             --carpeta "\\\\192.168.1.10\\Gestion_Documental3\\difusion\\dt\\IMAGENES_PL_1" ^
                             --salida "cruce_ACTA1.csv"

Si no se pasan los argumentos, el programa los pide de forma interactiva
(para poder ejecutarse haciendo doble clic en el .exe).
"""

import argparse
import csv
import os
import re
import sys
import time
import unicodedata
import zipfile
from datetime import datetime, date

try:
    import openpyxl
except ImportError:
    print("ERROR: falta la librería openpyxl.")
    print("Si estás corriendo el .py directamente: pip install openpyxl")
    sys.exit(1)


# ---------------------------------------------------------------------------
# 1) ESQUEMA DE 165 CAMPOS (orden exacto del clasificador AGN)
# ---------------------------------------------------------------------------
CAMPOS_165 = [
    "No DE ORDEN", "DEPE", "SER", "SUBS", "SERIE", "SUBSERIE", "ASUNTO",
    "LEGISLATURA", "CAMARA N°", "PAPEL", "ELECTRÓNICO", "OTROS", "INICIAL",
    "FINAL", "ID CÓDIGO DE BARRAS", "CONSECUTIVO UNIDAD CAMARA",
    "CONSECUTIVO UNIDAD", "ID CÓDIGO DE BARRAS", "CONSECUTIVO UNIDAD CAMARA",
    "CONSECUTIVO UNIDAD", "FOLIOS", "TOMOS", "CAJA", "LIBRO", "CARPETA",
    "OTRO", "NOTAS", "Version TRD/TVD", "Entidad remitente", "Fondo",
    "Unidad Administrativa", "Oficina Productora", "Objeto", "VIGENCIA",
    "CONTRATO", "LINEA", "ACTA", "ENTREGA", "LOTE", "Nombre objeto",
    "NOMBRE NORMALIZADO", "AÑO", "RUTA NORMALIZADA", "Volumen",
    "Ruta Original", "Tamaño Bytes", "Tamaño MB", "Extension Declarada",
    "MIME Real", "Extension Real", "Descripcion Tipo", "Transcripción",
    "TIPOLOGIA DOCUMENTAL", "PAGINAS", "UNIDAD DE RESOLUCIÓN",
    "FECHA DE CREACION", "FECHA DE INGESTA", "HASH ORIGEN", "HAHS DESTINO",
    "PERIODO", "INSTRUMENTO (TVD/TRD)", "CODIGO DE REFERENCIA",
    "CODIGO REFERENCIA PADRE", "PUBLICAR LIBSAFE", "PUBLICAR OPENACCESS",
    "REVISADO", "Nivel de descripción", "Título", "Fecha Inicial",
    "Fecha final", "Soporte", "Volumen",
    "Nombre del productor / Unidad Administrativa / Oficina Productora",
    "Historia institucional / Reseña biográfica", "Historia archivística",
    "Forma de ingreso", "Lugar(es) de emisión", "Alcance y contenido",
    "Información sobre valoración, selección, eliminación.",
    "Nuevos ingresos", "Sistema de organización", "Condiciones de acceso",
    "Condiciones de reproducción", "Lengua / escritura de la documentación",
    "Características físicas y requisitos técnicos",
    "Instrumentos de descripción", "Existencia y localización de los originales",
    "Existencia y localización de copias", "Unidades de descripción relacionadas",
    "Nota de publicaciones", "Notas", "Nota del archivero", "Reglas o normas",
    "Fecha (s) de la (s) descripción (es)", "Nombre", "Cargo", "Títulos",
    "Profesión u oficio", "Grado militar", "Condición étnica", "Lugar",
    "Tipo documental", "Entidad o Institución", "Temático", "Etnia",
    "CodEntRemitente", "EntRemitente", "CodEntProductora", "EntProductora",
    "CodUnidAdmin", "UnidadAdmin", "CodOfcProductora", "OfcProductora",
    "NumOrden", "CodSerie", "Serie", "CodSubserie", "Subserie", "CodTipoDoc",
    "TipoDoc", "CodigoIdentificacionOriginal", "Asunto", "Unidad Conservación",
    "Titulo", "Autor", "Director", "Productor", "Alcance y contenido",
    "Idioma", "PersonaNatural", "PersonaJurídica", "Toponímicos", "Temáticos",
    "Cronológicos", "Fecha Inicial", "Fecha Final", "NombreObjetoDigital",
    "Soporte", "Formato", "Duración", "Velocidad", "Color", "CalidadSonido",
    "MarcasAdicionales", "Notas", "Lote", "IdUnico", "TipoDoc", "Hash",
    "TipoHash", "MetadataSeguridad", "RutaFisica", "Link", "URLRecurso",
    "RutaImagenDifusión", "IDImagenMaestra", "ElaboradoPor", "Cargo",
    "FechaElabPor", "Entregado", "Cargo", "FechaEntregado", "Recibido",
    "Cargo", "FechaRecibido",
]
assert len(CAMPOS_165) == 165, f"Se esperaban 165 campos, hay {len(CAMPOS_165)}"

# ---------------------------------------------------------------------------
# 2) MAPEO campo -> fuente, por POSICIÓN (0-based) en CAMPOS_165.
#    Tipos de fuente:
#      ("FUID", "<encabezado del Excel>")   copia directa
#      ("TECH", "<clave técnica>")          resultado del escaneo del servidor
#      ("EXISTENTE", ["<alias 1>", ...])    ya lo tienen; se copia si existe,
#                                            NUNCA se recalcula/genera
#      ("CALC", "<nombre funcion>")         derivación simple
#      None                                  sin fuente -> "N/A"
# ---------------------------------------------------------------------------
MAPEO = {
    0:  ("FUID", "#"),
    1:  ("FUID", "CÓDIGO DEPENDENCIA (UNIDAD ADMINISTRATIVA)"),
    2:  ("FUID", "CÓDIGO DE SERIE"),
    3:  ("FUID", "CÓDIGO DE SUBSERIE"),
    4:  ("FUID", "NOMBRE DE SERIE"),
    5:  ("FUID", "NOMBRE DE LA SUBSERIE"),
    6:  None,   # ASUNTO
    7:  None,   # LEGISLATURA
    8:  None,   # CAMARA N°
    9:  ("CALC", "soporte_papel"),
    10: ("CALC", "soporte_electronico"),
    11: ("CALC", "soporte_otros"),
    12: ("FUID", "FECHA INICIAL"),
    13: ("FUID", "FECHA FINAL"),
    14: None,   # ID CÓDIGO DE BARRAS
    15: None,   # CONSECUTIVO UNIDAD CAMARA
    16: None,   # CONSECUTIVO UNIDAD
    17: None,   # ID CÓDIGO DE BARRAS (repetido)
    18: None,   # CONSECUTIVO UNIDAD CAMARA (repetido)
    19: None,   # CONSECUTIVO UNIDAD (repetido)
    20: None,   # FOLIOS
    21: None,   # TOMOS
    22: None,   # CAJA
    23: None,   # LIBRO
    24: None,   # CARPETA
    25: None,   # OTRO
    26: None,   # NOTAS
    27: None,   # Version TRD/TVD
    28: None,   # Entidad remitente
    29: None,   # Fondo
    30: ("FUID", "NOMBRE DEPENDENCIA (NOMBRE DE LA UNIDAD ADMINISTRATIVA)"),
    31: ("FUID", "NOMBRE OFICINA (OFICINA PRODUCTORA)"),
    32: None,   # Objeto
    33: None,   # VIGENCIA
    34: None,   # CONTRATO
    35: ("EXISTENTE", ["LINEA", "LÍNEA"]),
    36: None,   # ACTA
    37: None,   # ENTREGA
    38: None,   # LOTE
    39: ("FUID", "NOMBRE DE ARCHIVO"),
    40: None,   # NOMBRE NORMALIZADO
    41: ("CALC", "anio_desde_fecha_inicial"),
    42: None,   # RUTA NORMALIZADA
    43: ("TECH", "Volumen"),
    44: ("TECH", "Ruta Original"),
    45: ("TECH", "Tamaño Bytes"),
    46: ("TECH", "Tamaño MB"),
    47: ("TECH", "Extension Declarada"),
    48: ("TECH", "MIME Real"),
    49: ("TECH", "Extension Real"),
    50: ("TECH", "Descripcion Tipo"),
    51: None,   # Transcripción
    52: ("FUID", "NOMBRE DEL TIPO DOCUMENTAL"),
    53: None,   # PAGINAS
    54: None,   # UNIDAD DE RESOLUCIÓN
    55: None,   # FECHA DE CREACION
    56: None,   # FECHA DE INGESTA
    57: ("EXISTENTE", ["HASH ORIGEN", "HASH", "HASH_ORIGEN"]),
    58: None,   # HAHS DESTINO (se genera solo tras migración; no disponible aún)
    59: ("FUID", "INSTRUMENTO -  PERIODO DOCUMENTAL"),
    60: ("FUID", "INSTRUMENTO -  PERIODO DOCUMENTAL"),
    61: None,   # CODIGO DE REFERENCIA
    62: None,   # CODIGO REFERENCIA PADRE
    63: None,   # PUBLICAR LIBSAFE
    64: None,   # PUBLICAR OPENACCESS
    65: None,   # REVISADO
    66: None,   # Nivel de descripción
    67: None,   # Título
    68: ("FUID", "FECHA INICIAL"),
    69: ("FUID", "FECHA FINAL"),
    70: ("FUID", "SOPORTE"),
    71: ("TECH", "Volumen"),
    72: ("FUID", "NOMBRE OFICINA (OFICINA PRODUCTORA)"),
    73: None,   # Historia institucional
    74: None,   # Historia archivística
    75: None,   # Forma de ingreso
    76: None,   # Lugar(es) de emisión
    77: None,   # Alcance y contenido
    78: None,   # Información valoración/selección/eliminación
    79: None,   # Nuevos ingresos
    80: None,   # Sistema de organización
    81: ("CALC", "confidencialidad"),
    82: None,   # Condiciones de reproducción
    83: None,   # Lengua / escritura
    84: None,   # Características físicas y requisitos técnicos
    85: None,   # Instrumentos de descripción
    86: None,   # Existencia y localización de originales
    87: None,   # Existencia y localización de copias
    88: None,   # Unidades de descripción relacionadas
    89: None,   # Nota de publicaciones
    90: None,   # Notas
    91: None,   # Nota del archivero
    92: None,   # Reglas o normas
    93: None,   # Fecha(s) de la(s) descripción(es)
    94: None,   # Nombre
    95: None,   # Cargo
    96: None,   # Títulos
    97: None,   # Profesión u oficio
    98: None,   # Grado militar
    99: None,   # Condición étnica
    100: None,  # Lugar
    101: ("FUID", "NOMBRE DEL TIPO DOCUMENTAL"),
    102: None,  # Entidad o Institución
    103: None,  # Temático
    104: None,  # Etnia
    105: None,  # CodEntRemitente
    106: None,  # EntRemitente
    107: None,  # CodEntProductora
    108: None,  # EntProductora
    109: ("FUID", "CÓDIGO DEPENDENCIA (UNIDAD ADMINISTRATIVA)"),
    110: ("FUID", "NOMBRE DEPENDENCIA (NOMBRE DE LA UNIDAD ADMINISTRATIVA)"),
    111: ("FUID", "CÓDIGO OFICINA (OFICINA PRODUCTORA PRODUCTORA)"),
    112: ("FUID", "NOMBRE OFICINA (OFICINA PRODUCTORA)"),
    113: ("FUID", "#"),
    114: ("FUID", "CÓDIGO DE SERIE"),
    115: ("FUID", "NOMBRE DE SERIE"),
    116: ("FUID", "CÓDIGO DE SUBSERIE"),
    117: ("FUID", "NOMBRE DE LA SUBSERIE"),
    118: ("FUID", "CODIGO DEL TIPO DOCUMENTAL"),
    119: ("FUID", "NOMBRE DEL TIPO DOCUMENTAL"),
    120: ("FUID", "CODIGO OBJETO"),
    121: None,  # Asunto
    122: None,  # Unidad Conservación
    123: None,  # Titulo
    124: None,  # Autor
    125: None,  # Director
    126: None,  # Productor
    127: None,  # Alcance y contenido
    128: None,  # Idioma
    129: None,  # PersonaNatural
    130: None,  # PersonaJurídica
    131: None,  # Toponímicos
    132: None,  # Temáticos
    133: None,  # Cronológicos
    134: ("FUID", "FECHA INICIAL"),
    135: ("FUID", "FECHA FINAL"),
    136: ("FUID", "NOMBRE DE ARCHIVO"),
    137: ("FUID", "SOPORTE"),
    138: ("TECH", "Extension Real"),
    139: None,  # Duración
    140: None,  # Velocidad
    141: None,  # Color
    142: None,  # CalidadSonido
    143: None,  # MarcasAdicionales
    144: None,  # Notas
    145: None,  # Lote
    146: ("FUID", "ID"),
    147: ("FUID", "NOMBRE DEL TIPO DOCUMENTAL"),
    148: ("EXISTENTE", ["HASH", "HASH ORIGEN", "HASH_ORIGEN"]),
    149: ("EXISTENTE", ["TIPOHASH", "TIPO HASH", "ALGORITMO HASH"]),
    150: ("CALC", "confidencialidad"),
    151: ("TECH", "Ruta Original"),
    152: ("FUID", "HIPERVINCULO"),
    153: ("FUID", "HIPERVINCULO"),
    154: None,  # RutaImagenDifusión
    155: None,  # IDImagenMaestra
    156: None,  # ElaboradoPor
    157: None,  # Cargo
    158: None,  # FechaElabPor
    159: None,  # Entregado
    160: None,  # Cargo
    161: None,  # FechaEntregado
    162: None,  # Recibido
    163: None,  # Cargo
    164: None,  # FechaRecibido
}
assert len(MAPEO) == 165


# ---------------------------------------------------------------------------
# Detección de tipo real de archivo por firma binaria
# ---------------------------------------------------------------------------
FIRMAS = [
    (b"\x25\x50\x44\x46", "application/pdf", ".pdf", "PDF"),
    (b"\xff\xd8\xff", "image/jpeg", ".jpg", "JPEG"),
    (b"\x89\x50\x4e\x47", "image/png", ".png", "PNG"),
    (b"\x47\x49\x46\x38", "image/gif", ".gif", "GIF"),
    (b"\x49\x49\x2a\x00", "image/tiff", ".tif", "TIFF (little endian)"),
    (b"\x4d\x4d\x00\x2a", "image/tiff", ".tif", "TIFF (big endian)"),
    (b"\xd0\xcf\x11\xe0", "application/x-ole", ".doc", "OLE/COM (Office 97-2003)"),
    (b"\x1f\x8b\x08", "application/gzip", ".gz", "GZIP"),
    (b"\x42\x5a\x68", "application/x-bzip2", ".bz2", "BZIP2"),
    (b"\x37\x7a\xbc\xaf", "application/x-7z-compressed", ".7z", "7-Zip"),
    (b"\x52\x61\x72\x21", "application/x-rar-compressed", ".rar", "RAR"),
    (b"\x25\x21\x50\x53", "application/postscript", ".ps", "PostScript"),
]

ZIP_OOXML = {
    "word/": (".docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "Word (OOXML)"),
    "xl/": (".xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Excel (OOXML)"),
    "ppt/": (".pptx", "application/vnd.openxmlformats-officedocument.presentationml.presentation", "PowerPoint (OOXML)"),
}


def a_ruta_larga(ruta):
    if os.name == "nt" and not ruta.startswith("\\\\?\\"):
        ruta_abs = os.path.abspath(ruta)
        if ruta_abs.startswith("\\\\"):
            return "\\\\?\\UNC\\" + ruta_abs[2:]
        return "\\\\?\\" + ruta_abs
    return ruta


def detectar_tipo_real(ruta_completa):
    try:
        with open(a_ruta_larga(ruta_completa), "rb") as f:
            cabecera = f.read(8)
    except Exception:
        return ("N/A", "N/A", "No se pudo leer el archivo")

    for firma, mime, ext, desc in FIRMAS:
        if cabecera.startswith(firma):
            return (mime, ext, desc)

    if cabecera[:4] == b"PK\x03\x04":
        try:
            with zipfile.ZipFile(a_ruta_larga(ruta_completa)) as z:
                nombres = z.namelist()
                for prefijo, (ext, mime, desc) in ZIP_OOXML.items():
                    if any(n.startswith(prefijo) for n in nombres):
                        return (mime, ext, desc)
                return ("application/zip", ".zip", "ZIP / formato comprimido genérico")
        except Exception:
            return ("application/zip", ".zip", "ZIP (no se pudo inspeccionar contenido)")

    return ("application/octet-stream", "N/A", "Tipo no identificado por firma binaria")


def indexar_archivos(carpeta_raiz):
    indice, indice_lower = {}, {}
    duplicados, total_archivos = [], 0
    print(f"Escaneando (recursivo): {carpeta_raiz}")
    for raiz, _dirs, archivos in os.walk(a_ruta_larga(carpeta_raiz)):
        for nombre in archivos:
            total_archivos += 1
            ruta_completa = os.path.join(raiz, nombre)
            if nombre in indice:
                duplicados.append(nombre)
            else:
                indice[nombre] = ruta_completa
                indice_lower[nombre.lower()] = ruta_completa
    print(f"  -> {total_archivos} archivos encontrados en el servidor.")
    if duplicados:
        print(f"  -> AVISO: {len(duplicados)} nombres duplicados en subcarpetas distintas (se usó el primero).")
    return indice, indice_lower


def obtener_volumen(ruta_completa):
    partes = [p for p in ruta_completa.replace("/", "\\").split("\\") if p]
    if ruta_completa.startswith("\\\\") and len(partes) >= 2:
        return "\\\\" + partes[0] + "\\" + partes[1]
    drive, _ = os.path.splitdrive(ruta_completa)
    return drive if drive else "N/A"


# ---------------------------------------------------------------------------
# Normalización de encabezados para hacer el match tolerante a espacios extra
# ---------------------------------------------------------------------------
def normalizar(texto):
    if texto is None:
        return ""
    texto = str(texto).strip()
    texto = re.sub(r"\s+", " ", texto)
    texto = unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode("ascii")
    return texto.upper()


def construir_indice_encabezados(encabezados):
    """dict: encabezado normalizado -> encabezado original tal cual está en el Excel"""
    return {normalizar(h): h for h in encabezados}


def buscar_valor_fuid(fila_dict, indice_encabezados, nombre_buscado):
    clave = indice_encabezados.get(normalizar(nombre_buscado))
    if clave is None:
        return None, None
    return fila_dict.get(clave), clave


def buscar_valor_existente(fila_dict, indice_encabezados, alias_list):
    for alias in alias_list:
        clave = indice_encabezados.get(normalizar(alias))
        if clave is not None and fila_dict.get(clave) not in (None, ""):
            return fila_dict.get(clave), clave
    return None, None


# ---------------------------------------------------------------------------
# Funciones "CALC" (derivaciones simples, NO recalculan hash ni datos técnicos
# que ya vengan dados)
# ---------------------------------------------------------------------------
def calc_soporte_papel(fila_dict, indice_encabezados, **kw):
    v, _ = buscar_valor_fuid(fila_dict, indice_encabezados, "SOPORTE")
    return "X" if v and "PAPEL" in normalizar(v) else None


def calc_soporte_electronico(fila_dict, indice_encabezados, **kw):
    v, _ = buscar_valor_fuid(fila_dict, indice_encabezados, "SOPORTE")
    return "X" if v and "ELECTR" in normalizar(v) else None


def calc_soporte_otros(fila_dict, indice_encabezados, **kw):
    v, _ = buscar_valor_fuid(fila_dict, indice_encabezados, "SOPORTE")
    if v and ("PAPEL" in normalizar(v) or "ELECTR" in normalizar(v)):
        return None
    return "X" if v else None


def calc_anio_desde_fecha_inicial(fila_dict, indice_encabezados, **kw):
    v, _ = buscar_valor_fuid(fila_dict, indice_encabezados, "FECHA INICIAL")
    if isinstance(v, (datetime, date)):
        return v.year
    if isinstance(v, str):
        m = re.search(r"(19|20)\d{2}", v)
        if m:
            return m.group(0)
    return None


def calc_confidencialidad(fila_dict, indice_encabezados, **kw):
    campos_subserie = ["CONFIDENCIALIDAD PÚBLICA (DE LA SUBSERIE)",
                        "CONFIDENIALIDAD CLASIFICADA (DE LA SUBSERIE)",
                        "CONFIDENCIALIDAD RESERVADA (DE LA SUBSERIE)"]
    campos_serie = ["CONFIDENCIALIDAD PÚBLICA (DE LA SERIE)",
                     "CONFIDENCIALIDAD CLASIFICADA  (DE LA SERIE)",
                     "CONFIDENCIALIDAD RESERVADA  (DE LA SERIE)"]
    etiquetas = ["Pública", "Clasificada", "Reservada"]
    for grupo in (campos_subserie, campos_serie):
        for campo, etiqueta in zip(grupo, etiquetas):
            v, _ = buscar_valor_fuid(fila_dict, indice_encabezados, campo)
            if v and str(v).strip().upper() == "X":
                return etiqueta
    return None


FUNCIONES_CALC = {
    "soporte_papel": calc_soporte_papel,
    "soporte_electronico": calc_soporte_electronico,
    "soporte_otros": calc_soporte_otros,
    "anio_desde_fecha_inicial": calc_anio_desde_fecha_inicial,
    "confidencialidad": calc_confidencialidad,
}


# ---------------------------------------------------------------------------
# Cruce principal de una fila
# ---------------------------------------------------------------------------
def procesar_fila(fila_dict, indice_encabezados, indice_archivos, indice_archivos_lower):
    t0 = time.perf_counter()

    nombre_archivo_val, _ = buscar_valor_fuid(fila_dict, indice_encabezados, "NOMBRE DE ARCHIVO")
    nombre_archivo = str(nombre_archivo_val or "").strip()

    ruta_real = indice_archivos.get(nombre_archivo)
    if ruta_real is None and nombre_archivo:
        ruta_real = indice_archivos_lower.get(nombre_archivo.lower())

    tech = {}
    if ruta_real:
        try:
            tam_bytes = os.path.getsize(a_ruta_larga(ruta_real))
        except Exception:
            tam_bytes = None
        mime_real, ext_real, desc_tipo = detectar_tipo_real(ruta_real)
        _, ext_declarada = os.path.splitext(nombre_archivo)
        tech = {
            "Volumen": obtener_volumen(ruta_real),
            "Ruta Original": ruta_real,
            "Tamaño Bytes": tam_bytes if tam_bytes is not None else None,
            "Tamaño MB": round(tam_bytes / (1024 * 1024), 4) if tam_bytes is not None else None,
            "Extension Declarada": ext_declarada if ext_declarada else None,
            "MIME Real": mime_real,
            "Extension Real": ext_real,
            "Descripcion Tipo": desc_tipo,
        }
        coincidencia = "Si"
    else:
        coincidencia = "No"

    encabezados_consumidos = set()
    fila_salida = []

    for i, campo in enumerate(CAMPOS_165):
        fuente = MAPEO.get(i)
        valor = None

        if fuente is None:
            valor = None
        elif fuente[0] == "FUID":
            valor, clave_original = buscar_valor_fuid(fila_dict, indice_encabezados, fuente[1])
            if clave_original:
                encabezados_consumidos.add(clave_original)
        elif fuente[0] == "TECH":
            valor = tech.get(fuente[1])
        elif fuente[0] == "EXISTENTE":
            valor, clave_original = buscar_valor_existente(fila_dict, indice_encabezados, fuente[1])
            if clave_original:
                encabezados_consumidos.add(clave_original)
        elif fuente[0] == "CALC":
            valor = FUNCIONES_CALC[fuente[1]](fila_dict, indice_encabezados)

        if valor is None or (isinstance(valor, str) and valor.strip() == ""):
            valor = "N/A"

        fila_salida.append(valor)

    _, clave_nombre = buscar_valor_fuid(fila_dict, indice_encabezados, "NOMBRE DE ARCHIVO")
    if clave_nombre:
        encabezados_consumidos.add(clave_nombre)

    tiempo = round(time.perf_counter() - t0, 4)
    return fila_salida, coincidencia, tiempo, encabezados_consumidos


def leer_excel(ruta_excel):
    wb = openpyxl.load_workbook(ruta_excel, data_only=True)
    ws = wb.active
    filas = list(ws.iter_rows(values_only=True))
    if not filas:
        return [], []
    encabezados = [str(c).strip() if c is not None else "" for c in filas[0]]
    registros = []
    for fila in filas[1:]:
        if all(c is None for c in fila):
            continue
        registros.append({encabezados[i]: fila[i] if i < len(fila) else None
                           for i in range(len(encabezados))})
    return encabezados, registros


def nombres_columnas_unicos(nombres):
    """Evita encabezados de CSV duplicados (ej: 'Cargo' aparece 4 veces)."""
    vistos = {}
    resultado = []
    for n in nombres:
        if n not in vistos:
            vistos[n] = 1
            resultado.append(n)
        else:
            vistos[n] += 1
            resultado.append(f"{n}_{vistos[n]}")
    return resultado


def main():
    parser = argparse.ArgumentParser(description="Cruce FUID vs archivos del servidor (esquema de 165 campos)")
    parser.add_argument("--excel", help="Ruta del Excel FUID (ej: ACTA1 ENTREGABLE PROYECTOS DE LEY 1.xlsx)")
    parser.add_argument("--carpeta", help=r"Ruta de red de la carpeta a escanear (ej: \\IP\Gestion_Documental3\difusion\dt\IMAGENES_PL_1)")
    parser.add_argument("--salida", help="Ruta del CSV de salida", default=None)
    args = parser.parse_args()

    ruta_excel = args.excel or input("Ruta del Excel FUID: ").strip('"')
    ruta_carpeta = args.carpeta or input("Ruta de la carpeta en el servidor: ").strip('"')
    ruta_salida = args.salida or f"cruce_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"

    if not os.path.isfile(ruta_excel):
        print(f"ERROR: no se encontró el Excel: {ruta_excel}")
        input("Presiona Enter para salir...") if sys.stdin.isatty() else None
        sys.exit(1)
    if not os.path.isdir(ruta_carpeta):
        print(f"ERROR: no se encontró la carpeta: {ruta_carpeta}")
        input("Presiona Enter para salir...") if sys.stdin.isatty() else None
        sys.exit(1)

    t_inicio_total = time.perf_counter()

    print(f"Leyendo Excel: {ruta_excel}")
    encabezados, registros = leer_excel(ruta_excel)
    indice_encabezados = construir_indice_encabezados(encabezados)
    if normalizar("NOMBRE DE ARCHIVO") not in indice_encabezados:
        print("ERROR: el Excel no tiene una columna 'NOMBRE DE ARCHIVO'.")
        input("Presiona Enter para salir...") if sys.stdin.isatty() else None
        sys.exit(1)
    print(f"  -> {len(registros)} filas (actas) leídas.")

    indice_archivos, indice_archivos_lower = indexar_archivos(ruta_carpeta)

    print("Cruzando filas contra el esquema de 165 campos...")
    filas_165 = []
    coincidencias_lista = []
    tiempos = []
    todos_consumidos = set()

    for i, fila in enumerate(registros, start=1):
        fila_165, coincidencia, tiempo, consumidos = procesar_fila(
            fila, indice_encabezados, indice_archivos, indice_archivos_lower)
        filas_165.append(fila_165)
        coincidencias_lista.append(coincidencia)
        tiempos.append(tiempo)
        todos_consumidos |= consumidos
        nombre_mostrar, _ = buscar_valor_fuid(fila, indice_encabezados, "NOMBRE DE ARCHIVO")
        print(f"  [{i}/{len(registros)}] {nombre_mostrar} -> {coincidencia} ({tiempo}s)")

    tiempo_total = round(time.perf_counter() - t_inicio_total, 4)

    extras_fuid = [h for h in encabezados if h not in todos_consumidos and h != ""]

    encabezados_csv = nombres_columnas_unicos(CAMPOS_165)
    encabezados_csv += ["COINCIDENCIA_ARCHIVO", "TIEMPO_PROCESAMIENTO_SEG", "TIEMPO_TOTAL_LOTE_SEG"]
    encabezados_csv += [f"EXTRA_FUID__{h}" for h in extras_fuid]

    with open(ruta_salida, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f, delimiter=",")
        writer.writerow(encabezados_csv)
        for idx, fila_165 in enumerate(filas_165):
            fila_original = registros[idx]
            extra_vals = []
            for h in extras_fuid:
                v = fila_original.get(h)
                extra_vals.append("N/A" if v is None or (isinstance(v, str) and v.strip() == "") else v)
            fila_csv = fila_165 + [coincidencias_lista[idx], tiempos[idx], tiempo_total] + extra_vals
            writer.writerow(fila_csv)

    coincidencias_ok = sum(1 for c in coincidencias_lista if c == "Si")
    print("\n---- RESUMEN ----")
    print(f"Actas procesadas: {len(filas_165)}")
    print(f"Coincidencias encontradas en servidor: {coincidencias_ok}")
    print(f"Sin coincidencia: {len(filas_165) - coincidencias_ok}")
    print(f"Columnas del Excel FUID no mapeadas a las 165 (agregadas como EXTRA_FUID__): {len(extras_fuid)}")
    print(f"Tiempo total del lote: {tiempo_total} segundos")
    print(f"Tiempo promedio por archivo: {round(sum(tiempos)/len(tiempos), 4) if tiempos else 0} segundos")
    print(f"Archivo generado: {os.path.abspath(ruta_salida)}")
    if sys.stdin.isatty():
        try:
            input("\nProceso terminado. Presiona Enter para salir...")
        except EOFError:
            pass


if __name__ == "__main__":
    main()
