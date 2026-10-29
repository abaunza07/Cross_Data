"""
cruce_fuid_v3.py

Procesa DOS tipos de fuente distintos, seleccionables al arrancar:

  [1] FUID (Estructura 1) - el Excel con encabezados tipo "INSTRUMENTO -
      PERIODO DOCUMENTAL", "CÓDIGO DE SERIE", etc. Se cruza contra la
      carpeta del servidor de forma JERÁRQUICA (una o varias columnas,
      cada una un nivel de carpeta).

  [2] MALLA DOCUMENTAL (Estructura 2) - el Excel/CSV que sale de escanear
      archivos directamente (File name, File path, File size, Author...).
      Aquí NO se cruza contra el servidor porque el propio archivo ya
      trae su ruta (File path).

Ambos casos generan un registro con el mismo espíritu: los 56 campos base
del FUID (rellenos donde aplique) + un bloque de campos que hoy en día no
tienen cabida en el diccionario oficial de tu organización y por eso se
guardan bajo tipo_inventario = "MALLA DOCUMENTAL" (para la fuente 2 eso es
siempre así; para la fuente 1, cada FILA se marca "TIPO DOCUMENTAL" salvo
sus columnas sin hogar, que también quedan documentadas con ese tag para
que se sepa de dónde salieron).

REGLAS DE NORMALIZACIÓN (confirmadas contigo):
  - papel / electrónico / otros -> se FUNDEN en una sola columna
    "tipo_soporte" (si hay más de un "X", se listan separadas por coma).
  - unidad_conservacion_caja/libro/carpeta/otro -> NO se funden, se
    copian tal cual (X, texto o número).
  - Confidencialidad (pública/clasificada/reservada, de serie y de
    subserie) -> NO se funden. Cada columna se mantiene, y si tiene "X"
    el valor de salida es el propio nombre normalizado de la columna
    (ej. confidencialidad_publica_serie -> "confidencialidad_publica_serie").
  - Disposición final (DP eliminar/conservación total/selección/
    digitalizar/microfilmar) -> mismo criterio: no se funden, y cada una
    puede estar marcada (no son excluyentes entre sí).
  - AG / AC -> son números (años de retención), se copian tal cual, NUNCA
    se normalizan como si fueran "X".
  - ES PUNTO DE RADICACIÓN? -> es una bandera única (no un grupo), se
    normaliza a "Si"/"No" en vez del nombre de columna.

Uso: todo interactivo, no requiere parámetros. Al arrancar pregunta qué
tipo de fuente vas a procesar.
"""

import csv
import json
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
    print("ERROR: falta la libreria openpyxl. Instala con: pip install openpyxl")
    sys.exit(1)


# ---------------------------------------------------------------------------
# 0) ESQUEMA FINAL DE SALIDA: 184 campos, en el ORDEN EXACTO del archivo
#    maestro "Campos_de_inventario_FUID_CR_leo.xlsx" (columna DESCRIPCIÓN
#    CAMPO). Este es el orden real del CSV final. Tiene 4 nombres
#    repetidos (nivel_descripcion, volumen, alcance_contenido, lote) 
#    porque asi estan en el archivo maestro (aparecen en mas de una
#    categoria) — se numeran automaticamente al escribir el CSV para que
#    el encabezado sea valido (nombres_columnas_unicos).
# ---------------------------------------------------------------------------
CAMPOS_FINALES = [
    'entidad_remitente', 'fondo', 'unidad_administrativa', 'oficina_productora',
    'objeto', 'registro_entrada_fecha', 'registro_entrada_nt', 'numero_orden',
    'cod_oficina_productora', 'cod_serie', 'cod_subserie', 'serie', 'subserie',
    'asunto', 'descripcion', 'legislatura', 'tipo_soporte', 'fecha_inicial',
    'fecha_final', 'id_codigo_barras_carpeta', 'consecutivo_carpeta',
    'id_codigo_barras_caja', 'consecutivo_caja', 'id_codigo_barras_carpeta_almarchivos',
    'consecutivo_carpeta_almarchivos', 'id_codigo_barras_caja_almarchivos',
    'consecutivo_caja_almarchivos', 'id_codigo_barras_carpeta_camara',
    'consecutivo_carpeta_camara', 'id_codigo_barras_caja_camara',
    'consecutivo_caja_camara', 'numero_folios', 'foliacion_aproximada', 'tomos',
    'unidad_conservacion', 'unidad_conservacion_caja', 'unidad_conservacion_libro',
    'unidad_conservacion_carpeta', 'unidad_conservacion_otro', 'notas_fuid',
    'cod_digitalizacion', 'nivel_descripcion', 'codigo_referencia',
    'titulo_isadg_formal', 'titulo_isadg_atribuido', 'tipologia', 'numero_tomo',
    'numero_alterno', 'numero_carpeta', 'desc_fecha_inicial', 'desc_fecha_final',
    'desc_soporte', 'volumen', 'nombre_productor', 'historia_institucional',
    'historia_archivistica', 'lugar', 'alcance_contenido', 'palabras_clave',
    'funciones_congreso', 'condiciones_acceso', 'condiciones_reproduccion',
    'lengua_documentos', 'caracteristicas_requisitos_tecnicos',
    'instrumentos_descripcion', 'existencia_localizacion_originales',
    'unidades_descripcion_relacionadas', 'desc_notas', 'nombre_digitador',
    'reglas_normas', 'fecha_descripcion', 'cod_referencia_padre', 'publicar_libsafe',
    'publicar_openaccess', 'revisado', 'nivel_descripcion', 'forma_ingreso',
    'lugar_emision', 'alcance_contenido', 'informacion_valoracion_seleccion_eliminacion',
    'nuevos_ingresos', 'sistema_organizacion', 'existencia_localizacion_copias',
    'nota_publicaciones', 'nota_archivero', 'nombre', 'cargo', 'titulos',
    'profesion_oficio', 'grado_militar', 'condicion_etnica', 'tipo_documental',
    'entidad_ institucion', 'tematico', 'etnia', 'cod_ent_remitente',
    'cod_ent_productora', 'ent_productora', 'cod_unid_admin', 'nombre_tipo_doc',
    'codigo_identificacion_original', 'av_titulo', 'autor', 'director', 'productor',
    'av_alcance_contenido', 'idioma', 'tipo_persona', 'toponimicos', 'tematicos',
    'cronologicos', 'valoracion_candidato_eliminacion', 'av_fecha_inicial',
    'av_fecha_final', 'nombre_objeto_digital', 'av_soporte', 'formato', 'duracion',
    'velocidad', 'color', 'calidad_sonido', 'calidad_video', 'resolucion',
    'marcas_adicionales', 'lote', 'id_unico', 'tipo_doc', 'hash', 'tipo_hash',
    'metadata_seguridad', 'link', 'url_recurso', 'ruta_imagen_difusion',
    'id_imagen_maestra', 'elaborado_por', 'cargo_elaborado', 'fecha_elab_por',
    'entregado', 'cargo_entregado', 'fecha_entregado', 'recibe', 'cargo_recibido',
    'fecha_recibido', 'id', 'vigencia', 'contrato', 'linea', 'acta', 'entrega',
    'lote', 'nombre_objeto', 'nombre_normalizado', 'anio', 'ruta_normalizada',
    'volumen', 'ruta_original', 'peso_bytes', 'peso_mb', 'extension',
    'descripcion_tipo', 'transcripcion', 'dependencia', 'tipologia_documental',
    'paginas_pdf', 'unidad_resolucion', 'fecha_ingesta', 'hash_origen',
    'hash_destino', 'tipo_inventario', 'tipo_identificacion', 'numero_identificacion',
    'tipo_vinculacion', 'tipo_vinculacion_planta', 'tipo_vinculacion_utl',
    'registra_copias', 'periodo', 'comprendido_rango_fechas', 'nombre_funcionario',
    'vigencia_pago', 'numero_proyecto_camara', 'numero_proyecto_senado', 'ubicacion',
    'numero_documentos_electronicos', 'tamano_documentos_cajacarpeta_electronicos',
]
assert len(CAMPOS_FINALES) == 184, f"Se esperaban 184 campos, hay {len(CAMPOS_FINALES)}"


# ---------------------------------------------------------------------------
# 1) ESQUEMA BASE: 56 campos FUID (INVENTARIO DOCUMENTAL) — uso INTERNO
#    para el mapeo/calculo de valores; el orden de SALIDA real es
#    CAMPOS_FINALES de arriba.
# ---------------------------------------------------------------------------
CAMPOS_56 = [
    "inventario", "tipo_inventario", "entidad_remitente", "fondo",
    "unidad_administrativa", "oficina_productora", "objeto",
    "registro_entrada_fecha", "registro_entrada_nt", "numero_orden",
    "cod_oficina_productora", "cod_serie", "cod_subserie", "serie",
    "subserie", "asunto", "descripcion", "papel", "electrónico", "otros",
    "fecha_inicial", "fecha_final", "id_codigo_barras_carpeta",
    "consecutivo_carpeta", "consecutivo_caja", "id_codigo_barras_caja",
    "id_codigo_barras_carpeta_almarchivos", "consecutivo_carpeta_almarchivos",
    "id_codigo_barras_caja_almarchivos", "consecutivo_caja_almarchivos",
    "id_codigo_barras_carpeta_camara", "consecutivo_carpeta_camara",
    "id_codigo_barras_caja_camara", "consecutivo_caja_camara",
    "numero_folios", "foliacion_aproximada", "tomos",
    "unidad_conservacion_caja", "unidad_conservacion_libro",
    "unidad_conservacion_carpeta", "unidad_conservacion_otro",
    "notas_fuid", "cod_digitalizacion", "tipo_identificacion",
    "numero_identificacion", "tipo_vinculacion_planta", "tipo_vinculacion_utl",
    "registra_copias", "legislatura", "periodo", "comprendido_rango_fechas",
    "nombre_funcionario", "vigencia_pago", "ubicacion",
    "numero_documentos_electronicos", "tamano_documentos_cajacarpeta_electronicos",
]
assert len(CAMPOS_56) == 56, f"Se esperaban 56 campos, hay {len(CAMPOS_56)}"

CAMPOS_SOPORTE = ["papel", "electrónico", "otros"]

# Grupos que se mantienen SEPARADOS pero con normalizacion X -> nombre propio
GRUPO_CONFIDENCIALIDAD_SERIE = [
    ("CONFIDENCIALIDAD PÚBLICA (DE LA SERIE)", "confidencialidad_publica_serie"),
    ("CONFIDENCIALIDAD CLASIFICADA  (DE LA SERIE)", "confidencialidad_clasificada_serie"),
    ("CONFIDENCIALIDAD RESERVADA  (DE LA SERIE)", "confidencialidad_reservada_serie"),
]
GRUPO_CONFIDENCIALIDAD_SUBSERIE = [
    ("CONFIDENCIALIDAD PÚBLICA (DE LA SUBSERIE)", "confidencialidad_publica_subserie"),
    ("CONFIDENIALIDAD CLASIFICADA (DE LA SUBSERIE)", "confidencialidad_clasificada_subserie"),
    ("CONFIDENCIALIDAD RESERVADA (DE LA SUBSERIE)", "confidencialidad_reservada_subserie"),
]
GRUPO_DISPOSICION_FINAL = [
    ("DP ELIMINAR", "dp_eliminar"),
    ("DP CONSERVACIÓN TOTAL", "dp_conservacion_total"),
    ("DP SELECCIÓN", "dp_seleccion"),
    ("DP DIGITALIZAR", "dp_digitalizar"),
    ("DP MICROFILMAR", "dp_microfilmar"),
]

# Campos de Estructura 1 sin equivalente en el diccionario oficial -> MALLA DOCUMENTAL
CAMPOS_MALLA_DESDE_FUID = [
    ("CÓDIGO DEPENDENCIA (UNIDAD ADMINISTRATIVA)", "cod_dependencia"),
    ("ABREVIATURA OFICINA  (OFICINA PRODUCTORA)", "abreviatura_oficina"),
    ("AG", "ag"),
    ("AC", "ac"),
    ("PROCEDIMIENTO", "procedimiento"),
    ("INSTRUMENTO -  PERIODO DOCUMENTAL", "instrumento_periodo_documental"),
    ("IDENTIFICADOR", "identificador"),
    ("CODIGO OBJETO", "codigo_objeto"),
    ("NOMBRE DEL TIPO DOCUMENTAL", "nombre_tipo_documental_fuid"),
    ("CODIGO DEL TIPO DOCUMENTAL", "codigo_tipo_documental_fuid"),
]
# ES PUNTO DE RADICACIÓN? va aparte porque se normaliza a Si/No, no a su propio nombre
CAMPO_PUNTO_RADICACION = ("ES PUNTO DE RADICACIÓN?", "es_punto_radicacion")

# Todos los campos de "malla" que puede llegar a tener una fila de FUID
CAMPOS_MALLA_TODOS = (
    [c[1] for c in CAMPOS_MALLA_DESDE_FUID]
    + [CAMPO_PUNTO_RADICACION[1]]
    + [c[1] for c in GRUPO_CONFIDENCIALIDAD_SERIE]
    + [c[1] for c in GRUPO_CONFIDENCIALIDAD_SUBSERIE]
    + [c[1] for c in GRUPO_DISPOSICION_FINAL]
    # Nuevos de Malla Documental (Estructura 2) sin equivalente
    + ["software_creacion", "propietario", "fecha_modificacion", "fecha_creacion",
       "paginas_impares", "ancho_primera_pagina", "alto_primera_pagina",
       "archivo_difusion", "ciclo_vida"]
)

# Campos que YA existen en otras categorías de tu diccionario (L8 INVENTARIO TECNICO,
# ISAD-G, AUDIOVISUAL) y que la Malla Documental puede llegar a poblar. No son parte
# de los 56 ni son "nuevos sin hogar" -> van en su propia lista para no perderlos.
CAMPOS_OTRAS_CATEGORIAS = [
    "ruta_original", "peso_bytes", "paginas_pdf", "nombre_objeto",
    "descripcion_tipo", "titulo_isadg_formal", "autor", "palabras_clave",
]

# ---------------------------------------------------------------------------
# 2) MAPEO Estructura 1 (FUID) -> esquema de 56 + malla
#    ("FUID_HEADER", campo_normalizado, "56"|"MALLA")
# ---------------------------------------------------------------------------
MAPEO_FUID_DIRECTO = [
    ("INSTRUMENTO -  PERIODO DOCUMENTAL", "instrumento_periodo_documental", "MALLA"),
    ("CÓDIGO DEPENDENCIA (UNIDAD ADMINISTRATIVA)", "cod_dependencia", "MALLA"),
    ("NOMBRE DEPENDENCIA (NOMBRE DE LA UNIDAD ADMINISTRATIVA)", "unidad_administrativa", "56"),
    ("CÓDIGO OFICINA (OFICINA PRODUCTORA PRODUCTORA)", "cod_oficina_productora", "56"),
    ("NOMBRE OFICINA (OFICINA PRODUCTORA)", "oficina_productora", "56"),
    ("ABREVIATURA OFICINA  (OFICINA PRODUCTORA)", "abreviatura_oficina", "MALLA"),
    ("CÓDIGO DE SERIE", "cod_serie", "56"),
    ("NOMBRE DE SERIE", "serie", "56"),
    ("CÓDIGO DE SUBSERIE", "cod_subserie", "56"),
    ("NOMBRE DE LA SUBSERIE", "subserie", "56"),
    ("FECHA INICIAL", "fecha_inicial", "56"),
    ("FECHA FINAL", "fecha_final", "56"),
    ("CODIGO OBJETO", "codigo_objeto", "MALLA"),
    ("NOMBRE DEL TIPO DOCUMENTAL", "nombre_tipo_documental_fuid", "MALLA"),
    ("CODIGO DEL TIPO DOCUMENTAL", "codigo_tipo_documental_fuid", "MALLA"),
    ("IDENTIFICADOR", "identificador", "MALLA"),
    ("HIPERVINCULO", "link", "MALLA"),
    ("DESCRIPCION DE ARCHIVO", "descripcion", "56"),
    ("RUTA", "ubicacion", "56"),
    ("archivo_difusion", "archivo_difusion", "MALLA"),
    ("ciclo_vida", "ciclo_vida", "MALLA"),
]


# ---------------------------------------------------------------------------
# 3) MAPEO Estructura 2 (Malla Documental) -> esquema de 56 + malla
#
# Estos son los nombres de columna EN INGLES que trae la herramienta de
# escaneo mas comun (tipo ExifTool). Si una Malla en particular usa otros
# nombres de columna, esas columnas simplemente NO se reconocen aqui y
# caen en el mecanismo general de "campos_sin_mapear.json" (igual que
# cualquier columna desconocida de FUID): se registran ahi, vos les pones
# el campo_normalizado que corresponda de tus 165, y en la siguiente
# corrida se agregan como columnas extra al final del CSV. Todo queda en
# UN SOLO archivo (campos_sin_mapear.json), no hay un segundo archivo de
# configuracion aparte.
# ---------------------------------------------------------------------------
MAPEO_MALLA_DOCUMENTAL = [
    # (nombre_columna_esperado, campo_normalizado, origen)
    # NOTA: "File path" y "File name" NO se copian tal cual — se resuelven
    # cruzando "nombre_objeto" contra el servidor real (ver procesar_malla),
    # porque la ruta que trae la Malla puede estar desactualizada.
    ("File size", "peso_bytes", "56/malla ya existente"),
    ("Page count", "paginas_pdf", "56/malla ya existente"),
    ("File type", "descripcion_tipo", "56/malla ya existente"),
    ("Title", "titulo_isadg_formal", "56/malla ya existente"),
    ("Subject", "asunto", "56"),
    ("Author", "autor", "56/malla ya existente"),
    ("Keywords", "palabras_clave", "56/malla ya existente"),
    ("Creator", "software_creacion", "MALLA (nuevo)"),
    ("Owner", "propietario", "MALLA (nuevo)"),
    ("Modify date", "fecha_modificacion", "MALLA (nuevo)"),
    ("Create date", "fecha_creacion", "MALLA (nuevo)"),
    ("Odd pages", "paginas_impares", "MALLA (nuevo)"),
    ("First page width", "ancho_primera_pagina", "MALLA (nuevo)"),
    ("First page height", "alto_primera_pagina", "MALLA (nuevo)"),
]


# ---------------------------------------------------------------------------
# REGISTRO GENERAL DE COLUMNAS NO RECONOCIDAS
#
# Cualquier columna que aparezca en un Excel (FUID o Malla) y que el
# programa NO sepa a que campo de los ~165 corresponde, se anota aca en
# vez de perderse o hacer fallar el proceso. El archivo se va acumulando
# entre corridas (nunca se borra lo ya anotado). Vos despues editas este
# JSON dandole a cada columna su "campo_normalizado" y "categoria"
# correctos, y esa info se puede volver a usar en el script mas adelante.
# ---------------------------------------------------------------------------
NOMBRE_ARCHIVO_PENDIENTES = "campos_sin_mapear.json"


def _ruta_pendientes():
    return os.path.join(os.path.dirname(os.path.abspath(sys.argv[0])), NOMBRE_ARCHIVO_PENDIENTES)


def _cargar_pendientes():
    ruta = _ruta_pendientes()
    if os.path.isfile(ruta):
        try:
            with open(ruta, "r", encoding="utf-8") as f:
                return json.load(f), ruta
        except Exception:
            return {}, ruta
    return {}, ruta


def registrar_columnas_no_reconocidas(encabezados, encabezados_reconocidos_norm, tipo_fuente, ruta_archivo_origen):
    """
    encabezados: lista de columnas reales del Excel que se esta procesando.
    encabezados_reconocidos_norm: set de encabezados normalizados que el
        programa SI sabe mapear (para no marcarlos como pendientes).
    tipo_fuente: "FUID" o "MALLA DOCUMENTAL" (para dar contexto).
    Devuelve cuantas columnas NUEVAS se agregaron.
    """
    pendientes, ruta = _cargar_pendientes()
    nuevas = 0

    # Para Malla: lista de los campos que normalmente se esperan (los que
    # ya sabe llenar MAPEO_MALLA_DOCUMENTAL + la llave de cruce + los
    # nuevos sin categoria oficial), para sugerir como pista dentro del
    # propio JSON — asi no hay que adivinar ni buscar en otro lado.
    opciones_malla = sorted(set(
        [campo for _h, campo, _o in MAPEO_MALLA_DOCUMENTAL] + ["nombre_objeto", "archivo_difusion"]
    ))

    for h in encabezados:
        if not h or not h.strip():
            continue
        hn = normalizar_encabezado(h)
        if hn in encabezados_reconocidos_norm:
            continue
        if h not in pendientes:
            pendientes[h] = {
                "campo_normalizado": None,
                "categoria": None,
                "tipo_fuente_donde_aparecio": [tipo_fuente],
                "ejemplo_archivo": ruta_archivo_origen,
                "estado": "SIN_REVISAR",
            }
            if tipo_fuente == "MALLA DOCUMENTAL":
                pendientes[h]["opciones_sugeridas_malla"] = (
                    "Copia uno de estos en campo_normalizado SI corresponde (o escribe uno nuevo "
                    "si es un dato realmente distinto): " + ", ".join(opciones_malla)
                )
            nuevas += 1
        else:
            if tipo_fuente not in pendientes[h].get("tipo_fuente_donde_aparecio", []):
                pendientes[h]["tipo_fuente_donde_aparecio"].append(tipo_fuente)
            if tipo_fuente == "MALLA DOCUMENTAL" and "opciones_sugeridas_malla" not in pendientes[h]:
                pendientes[h]["opciones_sugeridas_malla"] = (
                    "Copia uno de estos en campo_normalizado SI corresponde (o escribe uno nuevo "
                    "si es un dato realmente distinto): " + ", ".join(opciones_malla)
                )

    with open(ruta, "w", encoding="utf-8") as f:
        json.dump(pendientes, f, ensure_ascii=False, indent=2)

    if nuevas > 0:
        print(f"\n⚠ {nuevas} columna(s) NUEVA(S) sin mapeo conocido se registraron en: {ruta}")
        print("  Edita ese archivo para asignarles su campo_normalizado y categoria correctos.")
    return nuevas


def cargar_mapeos_resueltos():
    """
    Lee campos_sin_mapear.json y devuelve SOLO las entradas que el usuario
    ya completo (campo_normalizado no vacio), como:
        {columna_original_del_excel: campo_normalizado_elegido}
    Estas se tratan como columnas YA CONOCIDAS y sus valores se agregan
    como columnas extra al final del CSV de salida.
    """
    pendientes, _ruta = _cargar_pendientes()
    resueltos = {}
    for columna_original, info in pendientes.items():
        campo = (info or {}).get("campo_normalizado")
        if campo and str(campo).strip():
            resueltos[columna_original] = str(campo).strip()
    return resueltos


def encabezados_reconocidos_fuid(columnas_cruce_usadas):
    """Todos los encabezados que el modo FUID SI sabe interpretar."""
    reconocidos = set()
    for header, _campo, _destino in MAPEO_FUID_DIRECTO:
        reconocidos.add(normalizar_encabezado(header))
    for grupo in (GRUPO_CONFIDENCIALIDAD_SERIE, GRUPO_CONFIDENCIALIDAD_SUBSERIE, GRUPO_DISPOSICION_FINAL):
        for header, _campo in grupo:
            reconocidos.add(normalizar_encabezado(header))
    reconocidos.add(normalizar_encabezado("SOPORTE"))
    reconocidos.add(normalizar_encabezado("AG"))
    reconocidos.add(normalizar_encabezado("AC"))
    reconocidos.add(normalizar_encabezado(CAMPO_PUNTO_RADICACION[0]))
    for campo in CAMPOS_SOPORTE:
        reconocidos.add(normalizar_encabezado(campo))
    for campo in CAMPOS_56:
        reconocidos.add(normalizar_encabezado(campo))
    for c in columnas_cruce_usadas:
        reconocidos.add(normalizar_encabezado(c))
    # Columnas que el usuario ya resolvio manualmente en campos_sin_mapear.json
    for columna_original in cargar_mapeos_resueltos():
        reconocidos.add(normalizar_encabezado(columna_original))
    return reconocidos


# Campos que el programa CALCULA/DERIVA (no son simples copias del Excel):
# resultado del cruce jerarquico contra el servidor, inspeccion real del
# archivo, o consolidaciones (tipo_soporte). Si el usuario, por error, le
# pone a una columna nueva uno de estos nombres, NO se deja sobreescribir
# (para no pisar en silencio un valor calculado con uno crudo del Excel).
# El resto de campos (autor, titulo_isadg_formal, peso_bytes, asunto, etc.)
# son simples copias directas -> SI se pueden alias sin riesgo, es
# justamente el caso de "otra herramienta le puso otro nombre a lo mismo".
CAMPOS_PROTEGIDOS = {
    "ruta_original", "nombre_objeto", "nombre_objeto_digital", "tipo_soporte", "inventario",
    "volumen", "peso_bytes", "peso_mb", "extension", "descripcion_tipo",
    "COINCIDENCIA_ARCHIVO", "mime_real",
    "TIEMPO_PROCESAMIENTO_SEG", "TIEMPO_TOTAL_LOTE_SEG",
}


def rellenar_por_nombre_directo(fila_completa, fila_dict, encabezados_norm):
    """
    RESPALDO GENERAL: si el Excel de origen ya trae los encabezados
    normalizados directamente (ej. una columna literal 'serie' en vez de
    'NOMBRE DE SERIE'), esta funcion los usa sin necesidad de que esten
    en MAPEO_FUID_DIRECTO/MAPEO_MALLA_DOCUMENTAL.

    Solo llena campos que sigan en "N/A" (no pisa nada que ya se haya
    calculado o mapeado correctamente), y nunca toca los campos
    protegidos (los que son calculados por el programa: ruta_original,
    inventario, tipo soporte, etc.).
    """
    for campo in CAMPOS_FINALES:
        if campo in CAMPOS_PROTEGIDOS:
            continue
        valor_actual = fila_completa.get(campo)
        if valor_actual not in (None, "N/A"):
            continue  # ya tiene un valor valido, no lo tocamos
        valor_directo = obtener_valor(fila_dict, encabezados_norm, campo)
        if valor_directo is not None and not (isinstance(valor_directo, str) and valor_directo.strip() == ""):
            fila_completa[campo] = valor_directo
    return fila_completa


def extraer_campos_resueltos(fila_dict, encabezados_norm, resueltos):
    """Devuelve {campo_normalizado: valor} para las columnas que el usuario
    ya mapeo manualmente y que existen en ESTE Excel en particular.
    Si el campo_normalizado elegido choca con uno YA gestionado por el
    programa (ej. ruta_original, peso_bytes...), se ignora y se avisa,
    para no pisar en silencio un valor correcto con uno crudo del Excel."""
    fila_extra = {}
    for columna_original, campo_normalizado in resueltos.items():
        if campo_normalizado in CAMPOS_PROTEGIDOS:
            print(f"  ADVERTENCIA: en campos_sin_mapear.json, la columna '{columna_original}' "
                  f"quiere usar el campo '{campo_normalizado}', pero ese campo ya lo calcula "
                  f"el programa internamente. Se IGNORA para no pisar el valor correcto. "
                  f"Usa otro nombre en campo_normalizado.")
            continue
        clave = encabezados_norm.get(normalizar_encabezado(columna_original))
        if clave is not None:
            fila_extra[campo_normalizado] = a_na(fila_dict.get(clave))
    return fila_extra


# ---------------------------------------------------------------------------
# Normalización de texto (espacios raros, invisibles)
# ---------------------------------------------------------------------------
def limpiar_invisibles(texto):
    if texto is None:
        return ""
    texto = str(texto)
    for ch in ["\u00a0", "\u200b", "\u200c", "\u200d", "\ufeff"]:
        texto = texto.replace(ch, " ")
    texto = texto.strip()
    texto = re.sub(r"\s+", " ", texto)
    return texto


def normalizar_encabezado(texto):
    texto = limpiar_invisibles(texto)
    texto = unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode("ascii")
    return texto.upper()


def comparar_nombres(valor_excel, nombre_real):
    a = limpiar_invisibles(valor_excel)
    b = limpiar_invisibles(nombre_real)
    if a == "":
        return (None, None)
    if a == b:
        return ("exacto", None)
    if a.lower() == b.lower():
        return ("aproximado", "difieren en mayusculas/minusculas")
    return (None, None)


def es_x(valor):
    return valor is not None and limpiar_invisibles(str(valor)).upper() == "X"


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
                return ("application/zip", ".zip", "ZIP / formato comprimido generico")
        except Exception:
            return ("application/zip", ".zip", "ZIP (no se pudo inspeccionar contenido)")
    return ("application/octet-stream", "N/A", "Tipo no identificado por firma binaria")


def obtener_volumen(ruta_completa):
    partes = [p for p in ruta_completa.replace("/", "\\").split("\\") if p]
    if ruta_completa.startswith("\\\\") and len(partes) >= 2:
        return "\\\\" + partes[0] + "\\" + partes[1]
    drive, _ = os.path.splitdrive(ruta_completa)
    return drive if drive else "N/A"


# ---------------------------------------------------------------------------
# Índice de directorio con caché + cruce jerárquico (para FUID)
# ---------------------------------------------------------------------------
def construir_indice_global(carpeta_raiz):
    """
    Indexa TODO el arbol (carpetas Y archivos, recursivo) para poder
    encontrar el NIVEL 1 del cruce (ej. la caja) sin importar en que
    subcarpeta este anidada, y sin tener que re-escanear todo el servidor
    fila por fila.

    Cada ARCHIVO se indexa dos veces: con su nombre completo (con
    extension) y con su nombre SIN extension — porque muchos FUID
    guardan el codigo (ej. "830-DA-PE-1-1") sin la extension real del
    archivo ("830-DA-PE-1-1.pdf").

    Devuelve: {nombre_limpio: [rutas]}, {nombre_limpio_lower: [rutas]}
    """
    indice, indice_lower = {}, {}
    raiz_larga = a_ruta_larga(carpeta_raiz)
    for raiz, dirs, archivos in os.walk(raiz_larga):
        for nombre in dirs:
            clave = limpiar_invisibles(nombre)
            ruta_completa = os.path.join(raiz, nombre)
            indice.setdefault(clave, []).append(ruta_completa)
            indice_lower.setdefault(clave.lower(), []).append(ruta_completa)
        for nombre in archivos:
            ruta_completa = os.path.join(raiz, nombre)
            base, _ext = os.path.splitext(nombre)
            for nombre_candidato in {nombre, base}:
                clave = limpiar_invisibles(nombre_candidato)
                indice.setdefault(clave, []).append(ruta_completa)
                indice_lower.setdefault(clave.lower(), []).append(ruta_completa)
    return indice, indice_lower


def candidatos_nivel1(indice_global, indice_global_lower, valor):
    """Todos los lugares del arbol donde el nombre de nivel 1 (ej. caja)
    aparece — puede estar anidado en cualquier subcarpeta."""
    if not valor:
        return []
    clave = limpiar_invisibles(str(valor))
    if clave in indice_global:
        return [(r, "exacto") for r in indice_global[clave]]
    if clave.lower() in indice_global_lower:
        return [(r, "aproximado") for r in indice_global_lower[clave.lower()]]
    return []


def candidatos_dentro_de(carpeta_base, valor):
    """Busca RECURSIVAMENTE (en toda la subcarpeta, no solo el nivel
    inmediato) todas las coincidencias de 'valor' dentro de carpeta_base.
    Se usa para el nivel 2+ del cruce (ej. la carpeta dentro de la caja
    ya encontrada) — puede estar anidada mas de un nivel adentro."""
    if not valor:
        return []
    exactos, aproximados = [], []
    for raiz, dirs, archivos in os.walk(a_ruta_larga(carpeta_base)):
        for nombre in dirs + archivos:
            base, _ext = os.path.splitext(nombre)
            ruta_completa = os.path.join(raiz, nombre)
            for candidato_nombre in {nombre, base}:
                tipo_match, _ = comparar_nombres(valor, candidato_nombre)
                if tipo_match == "exacto":
                    exactos.append((ruta_completa, "exacto"))
                elif tipo_match == "aproximado":
                    aproximados.append((ruta_completa, "aproximado"))
    vistos, resultado = set(), []
    for ruta, tipo in exactos + aproximados:
        if ruta not in vistos:
            vistos.add(ruta)
            resultado.append((ruta, tipo))
    return resultado


def cruzar_jerarquico(carpeta_raiz, valores_columnas_cruce, indice_global, indice_global_lower):
    """
    Cruce jerarquico con BUSQUEDA RECURSIVA en cada nivel (no solo mira
    la carpeta inmediata): el nivel 1 (ej. caja) se busca en TODO el
    arbol del servidor (via indice global). El nivel 2+ (ej. carpeta) se
    busca RECURSIVAMENTE dentro de la carpeta encontrada en el nivel
    anterior (no solo sus hijos directos).

    Si el primer candidato del nivel 1 no tiene el nivel 2 adentro,
    se prueba con el SIGUIENTE candidato del nivel 1 (por si el mismo
    nombre de caja se repite en mas de un lugar del servidor), y asi
    sucesivamente (retroceso simple).

    CASO ESPECIAL: si se usa UNA SOLA columna de cruce (porque el valor
    ya deberia ser especifico/unico, ej. un codigo de digitalizacion) y
    ese valor aparece en MAS DE UN lugar del servidor, NO se elige
    ninguno automaticamente — se marca como "Ambiguo" para que se
    revise a mano, en vez de arriesgarse a traer el archivo equivocado.
    """
    solo_una_columna = len(valores_columnas_cruce) == 1

    def _buscar(idx, base_actual):
        if idx >= len(valores_columnas_cruce):
            return base_actual, []

        valor = valores_columnas_cruce[idx]
        if idx == 0:
            candidatos = candidatos_nivel1(indice_global, indice_global_lower, valor)
        else:
            candidatos = candidatos_dentro_de(base_actual, valor)

        if not candidatos:
            return None, [{"nivel": idx + 1, "valor_buscado": valor,
                            "resultado": "No", "tipo_match": None}]

        if solo_una_columna and idx == 0 and len(candidatos) > 1:
            rutas_candidatas = "; ".join(r for r, _t in candidatos)
            return None, [{"nivel": idx + 1, "valor_buscado": valor,
                            "resultado": f"Ambiguo ({len(candidatos)} lugares distintos, "
                                         f"no se eligio ninguno automaticamente)",
                            "tipo_match": None, "candidatos_ambiguos": rutas_candidatas}]

        for ruta_candidata, tipo_match in candidatos:
            if idx == len(valores_columnas_cruce) - 1:
                # Es el ultimo nivel: este candidato ya es la respuesta.
                return ruta_candidata, [{"nivel": idx + 1, "valor_buscado": valor,
                                          "resultado": "Si", "tipo_match": tipo_match}]
            if not os.path.isdir(a_ruta_larga(ruta_candidata)):
                continue  # no se puede seguir buscando adentro de un archivo
            resultado_final, niveles_resto = _buscar(idx + 1, ruta_candidata)
            if resultado_final is not None:
                nivel_info = {"nivel": idx + 1, "valor_buscado": valor,
                               "resultado": "Si", "tipo_match": tipo_match}
                return resultado_final, [nivel_info] + niveles_resto

        # Ningun candidato del nivel 1 (o superior) tuvo el nivel siguiente adentro
        return None, [{"nivel": idx + 1, "valor_buscado": valor,
                        "resultado": f"No (se encontraron {len(candidatos)} candidato(s), "
                                     f"pero ninguno tenia el siguiente nivel adentro)",
                        "tipo_match": None}]

    ruta_final, niveles_resultado = _buscar(0, carpeta_raiz)

    # Completar niveles no alcanzados (por si el cruce fallo antes de llegar a todos)
    while len(niveles_resultado) < len(valores_columnas_cruce):
        idx_faltante = len(niveles_resultado)
        niveles_resultado.append({
            "nivel": idx_faltante + 1, "valor_buscado": valores_columnas_cruce[idx_faltante],
            "resultado": "N/A (nivel anterior no encontrado)", "tipo_match": None})

    return ruta_final, niveles_resultado


def indexar_archivos_servidor(carpeta_raiz):
    """Recorre TODA la carpeta (recursivo) y arma un índice {nombre: ruta}."""
    indice, indice_lower = {}, {}
    total = 0
    for raiz, _dirs, archivos in os.walk(a_ruta_larga(carpeta_raiz)):
        for nombre in archivos:
            total += 1
            ruta_completa = os.path.join(raiz, nombre)
            if nombre not in indice:
                indice[nombre] = ruta_completa
                indice_lower[nombre.lower()] = ruta_completa
    return indice, indice_lower, total


def buscar_por_nombre(indice, indice_lower, nombre_buscado):
    """Busca nombre_buscado en el índice, tolerando espacios/invisibles y
    mayúsculas/minúsculas. Devuelve (ruta_o_None, tipo_match)."""
    if not nombre_buscado:
        return None, None
    nombre_limpio = limpiar_invisibles(str(nombre_buscado))
    if nombre_limpio in indice:
        return indice[nombre_limpio], "exacto"
    if nombre_limpio.lower() in indice_lower:
        return indice_lower[nombre_limpio.lower()], "aproximado"
    return None, None


def resolver_archivo_final(ruta_final):
    if ruta_final is None:
        return None, "sin_ruta"
    ruta_larga = a_ruta_larga(ruta_final)
    if os.path.isfile(ruta_larga):
        return ruta_final, "archivo_directo"
    if os.path.isdir(ruta_larga):
        try:
            entradas = os.listdir(ruta_larga)
        except Exception:
            entradas = []
        contenido = [f for f in entradas
                     if os.path.isfile(a_ruta_larga(os.path.join(ruta_final, f)))]
        if len(contenido) == 1:
            return os.path.join(ruta_final, contenido[0]), "unico_archivo_en_carpeta"
        if len(contenido) == 0:
            return None, "carpeta_vacia"
        return None, f"carpeta_con_{len(contenido)}_archivos_ambiguo"
    return None, "ruta_invalida"


# ---------------------------------------------------------------------------
# Utilidades de lectura de Excel
# ---------------------------------------------------------------------------
def leer_excel(ruta_excel, nombre_hoja=None):
    wb = openpyxl.load_workbook(ruta_excel, data_only=True)
    ws = wb[nombre_hoja] if nombre_hoja else wb.active
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


def leer_todas_las_hojas(ruta_excel):
    """
    Lee TODAS las hojas de un Excel (no solo la activa). Devuelve una
    lista de (nombre_hoja, encabezados, registros), omitiendo hojas
    vacias o sin encabezados. Si el archivo tiene una sola hoja util,
    la lista tiene un solo elemento (comportamiento identico a antes).
    """
    wb = openpyxl.load_workbook(ruta_excel, data_only=True)
    resultado = []
    for nombre_hoja in wb.sheetnames:
        ws = wb[nombre_hoja]
        filas = list(ws.iter_rows(values_only=True))
        if not filas:
            continue
        encabezados = [str(c).strip() if c is not None else "" for c in filas[0]]
        if not any(h for h in encabezados):
            continue  # hoja sin encabezados reales, se omite
        registros = []
        for fila in filas[1:]:
            if all(c is None for c in fila):
                continue
            registros.append({encabezados[i]: fila[i] if i < len(fila) else None
                               for i in range(len(encabezados))})
        if registros:
            resultado.append((nombre_hoja, encabezados, registros))
    return resultado


def construir_indice_encabezados(encabezados):
    return {normalizar_encabezado(h): h for h in encabezados}


def obtener_valor(fila_dict, encabezados_norm, nombre_campo):
    clave = encabezados_norm.get(normalizar_encabezado(nombre_campo))
    if clave is None:
        return None
    return fila_dict.get(clave)


def a_na(valor):
    if valor is None or (isinstance(valor, str) and valor.strip() == ""):
        return "N/A"
    return valor


def nombres_columnas_unicos(nombres):
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


# ---------------------------------------------------------------------------
# tipo_soporte (unico grupo que SI se funde)
# ---------------------------------------------------------------------------
def calcular_tipo_soporte(fila_dict, encabezados_norm):
    """
    Soporta dos formatos de origen:
      A) 3 columnas separadas (papel/electrónico/otros) marcadas con X
         -> se funden con coma si hay mas de una marcada.
      B) 1 sola columna de texto "SOPORTE" (ej. "Electrónico", "Papel")
         -> se normaliza el texto tal cual, sin necesidad de X.
    """
    etiquetas_valor = {"papel": "papel", "electrónico": "electronico", "otros": "otros"}
    marcados = []
    hay_columnas_separadas = False
    for campo in CAMPOS_SOPORTE:
        valor = obtener_valor(fila_dict, encabezados_norm, campo)
        if valor is not None:
            hay_columnas_separadas = True
        if es_x(valor):
            marcados.append(etiquetas_valor[campo])
    if hay_columnas_separadas:
        return ",".join(marcados) if marcados else "N/A"

    # Formato B: una sola columna de texto "SOPORTE"
    valor_unico = obtener_valor(fila_dict, encabezados_norm, "SOPORTE")
    if valor_unico is None or (isinstance(valor_unico, str) and valor_unico.strip() == ""):
        return "N/A"
    texto = limpiar_invisibles(str(valor_unico))
    texto_norm = unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode("ascii").lower()
    if "papel" in texto_norm:
        return "papel"
    if "electr" in texto_norm:
        return "electronico"
    return texto.lower()


# ---------------------------------------------------------------------------
# MODO 1: FUID (Estructura 1)
# ---------------------------------------------------------------------------
def listar_excels_en_carpeta(carpeta):
    """Lista los archivos .xlsx/.xls dentro de una carpeta (nivel superior,
    no recursivo), ordenados por nombre."""
    extensiones = (".xlsx", ".xls")
    archivos = []
    for nombre in sorted(os.listdir(carpeta)):
        ruta = os.path.join(carpeta, nombre)
        if os.path.isfile(ruta) and nombre.lower().endswith(extensiones) and not nombre.startswith("~$"):
            archivos.append(ruta)
    return archivos


def pedir_columnas_cruce_fuid(encabezados):
    print("\nColumnas disponibles en el Excel FUID:")
    for h in encabezados:
        print(f"  - {h}")
    while True:
        try:
            n = int(input("\nCon cuantas columnas quieres hacer el cruce jerarquico "
                           "(ej. 1 = solo caja, 2 = caja+carpeta): ").strip())
            if n >= 1:
                break
        except ValueError:
            pass
        print("Por favor ingresa un numero entero mayor o igual a 1.")

    columnas = []
    for i in range(1, n + 1):
        while True:
            nombre = input(f"Nombre EXACTO de la columna para el nivel {i}: ").strip()
            if nombre in encabezados:
                columnas.append(nombre)
                break
            print(f"  '{nombre}' no esta en el Excel. Copia el nombre tal cual aparece arriba.")
    return columnas


def procesar_fuid(fila_dict, encabezados_norm, columnas_cruce, carpeta_raiz, indice_global, indice_global_lower):
    t0 = time.perf_counter()

    valores_cruce = [obtener_valor(fila_dict, encabezados_norm, col) for col in columnas_cruce]
    ruta_final_bruta, niveles = cruzar_jerarquico(carpeta_raiz, valores_cruce, indice_global, indice_global_lower)
    ruta_archivo, motivo_resolucion = resolver_archivo_final(ruta_final_bruta)

    if ruta_archivo is None and niveles and "candidatos_ambiguos" in niveles[0]:
        motivo_resolucion = "ambiguo_multiples_candidatos"

    fila_56 = {campo: "N/A" for campo in CAMPOS_56}
    fila_malla = {campo: "N/A" for campo in CAMPOS_MALLA_TODOS}

    # tipo_inventario fijo para esta fuente
    fila_56["tipo_inventario"] = "TIPO DOCUMENTAL"

    # Mapeos directos
    for header_fuid, campo_norm, destino in MAPEO_FUID_DIRECTO:
        valor = a_na(obtener_valor(fila_dict, encabezados_norm, header_fuid))
        if destino == "56":
            fila_56[campo_norm] = valor
        else:
            fila_malla[campo_norm] = valor

    # tipo_soporte (unico grupo fundido)
    fila_56["tipo_soporte"] = calcular_tipo_soporte(fila_dict, encabezados_norm)

    # unidad_conservacion_* : copiar tal cual, NO fundir
    for campo in ("unidad_conservacion_caja", "unidad_conservacion_libro",
                  "unidad_conservacion_carpeta", "unidad_conservacion_otro"):
        etiqueta = {"unidad_conservacion_caja": "CAJA", "unidad_conservacion_libro": "LIBRO",
                    "unidad_conservacion_carpeta": "CARPETA", "unidad_conservacion_otro": "OTRO"}[campo]
        fila_56[campo] = a_na(obtener_valor(fila_dict, encabezados_norm, etiqueta))

    # AG / AC: numeros, se copian tal cual
    fila_malla["ag"] = a_na(obtener_valor(fila_dict, encabezados_norm, "AG"))
    fila_malla["ac"] = a_na(obtener_valor(fila_dict, encabezados_norm, "AC"))

    # ES PUNTO DE RADICACIÓN?: bandera unica -> Si/No
    valor_radicacion = obtener_valor(fila_dict, encabezados_norm, CAMPO_PUNTO_RADICACION[0])
    fila_malla[CAMPO_PUNTO_RADICACION[1]] = "Si" if es_x(valor_radicacion) else ("N/A" if valor_radicacion is None else "No")

    # Grupos que se mantienen separados, X -> nombre propio de la columna
    for grupo in (GRUPO_CONFIDENCIALIDAD_SERIE, GRUPO_CONFIDENCIALIDAD_SUBSERIE, GRUPO_DISPOSICION_FINAL):
        for header_original, campo_norm in grupo:
            valor = obtener_valor(fila_dict, encabezados_norm, header_original)
            fila_malla[campo_norm] = campo_norm if es_x(valor) else "N/A"

    # Campos tecnicos del escaneo del servidor -> se escriben DIRECTAMENTE
    # en los nombres oficiales de la lista de 184 (volumen, ruta_original,
    # peso_bytes, peso_mb, extension, descripcion_tipo). "mime_real" no
    # tiene casilla oficial todavia, se guarda como extra.
    if ruta_archivo:
        try:
            tam_bytes = os.path.getsize(a_ruta_larga(ruta_archivo))
        except Exception:
            tam_bytes = None
        mime_real, ext_real, desc_tipo = detectar_tipo_real(ruta_archivo)
        tecnicos = {
            "COINCIDENCIA_ARCHIVO": "Si",
            "volumen": obtener_volumen(ruta_archivo),
            "ruta_original": ruta_archivo,
            "peso_bytes": tam_bytes if tam_bytes is not None else "N/A",
            "peso_mb": round(tam_bytes / (1024 * 1024), 4) if tam_bytes is not None else "N/A",
            "extension": ext_real,
            "descripcion_tipo": desc_tipo,
            "mime_real": mime_real,
        }
    else:
        tecnicos = {"COINCIDENCIA_ARCHIVO": "No", "volumen": "N/A", "ruta_original": "N/A",
                    "peso_bytes": "N/A", "peso_mb": "N/A", "extension": "N/A",
                    "descripcion_tipo": "N/A", "mime_real": "N/A"}

    tiempo = round(time.perf_counter() - t0, 4)
    return fila_56, fila_malla, tecnicos, niveles, motivo_resolucion, tiempo, ruta_archivo


def modo_fuid():
    ruta_entrada = input("Ruta del Excel FUID (o de una CARPETA con varios Excels FUID): ").strip('"')

    if os.path.isdir(ruta_entrada):
        archivos_excel = listar_excels_en_carpeta(ruta_entrada)
        if not archivos_excel:
            print(f"ERROR: no se encontraron archivos .xlsx/.xls dentro de: {ruta_entrada}")
            return
        print(f"Se encontro una CARPETA. Detectados {len(archivos_excel)} archivo(s) Excel:")
        for a in archivos_excel:
            print(f"  - {os.path.basename(a)}")
    elif os.path.isfile(ruta_entrada):
        archivos_excel = [ruta_entrada]
    else:
        print(f"ERROR: no se encontro el archivo ni la carpeta: {ruta_entrada}")
        return

    ruta_carpeta = input("Ruta de la carpeta en el servidor: ").strip('"')
    if not os.path.isdir(ruta_carpeta):
        print(f"ERROR: no se encontro la carpeta: {ruta_carpeta}")
        return

    t_inicio = time.perf_counter()

    # Leer TODAS LAS HOJAS de TODOS los archivos de antemano (para poder
    # registrar columnas sin mapear de todo antes de decidir si se sigue
    # o no, y para armar la pregunta de columnas de cruce con encabezados
    # reales). Cada hoja util de cada archivo se trata como una unidad
    # de trabajo independiente (misma logica que "un archivo mas").
    datos_por_hoja = []  # [(ruta_excel, nombre_hoja, encabezados, encabezados_norm, registros), ...]
    for ruta_excel_actual in archivos_excel:
        try:
            hojas = leer_todas_las_hojas(ruta_excel_actual)
        except Exception as e:
            print(f"AVISO: no se pudo leer '{os.path.basename(ruta_excel_actual)}' ({e}). Se omite.")
            continue
        if not hojas:
            print(f"AVISO: '{os.path.basename(ruta_excel_actual)}' no tiene hojas con datos. Se omite.")
            continue
        for nombre_hoja, encabezados, registros in hojas:
            encabezados_norm = construir_indice_encabezados(encabezados)
            datos_por_hoja.append((ruta_excel_actual, nombre_hoja, encabezados, encabezados_norm, registros))
            etiqueta_hoja = f" [hoja: {nombre_hoja}]" if len(hojas) > 1 else ""
            print(f"  {os.path.basename(ruta_excel_actual)}{etiqueta_hoja}: {len(registros)} filas leidas.")

    if not datos_por_hoja:
        print("ERROR: no se pudo leer ninguna hoja valida de ningun archivo Excel.")
        return

    total_filas_excel = sum(len(r) for _e, _h, _hd, _hn, r in datos_por_hoja)
    print(f"\nTotal de filas a procesar (todos los archivos/hojas): {total_filas_excel}")

    # 1) Registrar columnas sin mapear de TODAS las hojas antes de
    #    decidir si se continua (evita interrumpir a mitad de un lote).
    reconocidos_fuid = encabezados_reconocidos_fuid([])
    nuevas_total = 0
    for ruta_excel_actual, nombre_hoja, encabezados, _hn, _r in datos_por_hoja:
        nuevas_total += registrar_columnas_no_reconocidas(
            encabezados, reconocidos_fuid, "FUID", f"{ruta_excel_actual} [hoja: {nombre_hoja}]")

    resueltos = cargar_mapeos_resueltos()
    campos_extra_dinamicos = sorted(set(resueltos.values()))
    if campos_extra_dinamicos:
        print(f"Aplicando {len(campos_extra_dinamicos)} campo(s) resuelto(s) manualmente: {campos_extra_dinamicos}")

    if nuevas_total > 0:
        print(f"\nSe encontraron {nuevas_total} columna(s) nueva(s) sin mapeo conocido (en total, "
              f"contando todos los archivos/hojas).")
        respuesta = input("¿Continuar con el cruce de todas formas? (S/N, N para salir y "
                           "completar campos_sin_mapear.json primero): ").strip().upper()
        if respuesta != "S":
            print("Saliendo sin procesar. Completa campos_sin_mapear.json y vuelve a correr.")
            return

    # 2) Preguntar por las columnas de cruce UNA sola vez (se usan las
    #    mismas para todas las hojas/archivos del lote), mostrando los
    #    encabezados de la primera hoja como referencia.
    _primer_ruta, _primer_hoja, primer_encabezados, _hn0, _r0 = datos_por_hoja[0]
    columnas_cruce = pedir_columnas_cruce_fuid(primer_encabezados)
    print(f"\nCruzando por (en orden): {columnas_cruce}")

    print(f"\nIndexando (recursivo) todo el servidor: {ruta_carpeta}")
    indice_global, indice_global_lower = construir_indice_global(ruta_carpeta)
    print(f"  -> {len(indice_global)} nombres unicos indexados (carpetas + archivos).")

    todos_los_archivos = set()
    for raiz, _dirs, archivos in os.walk(a_ruta_larga(ruta_carpeta)):
        for nombre in archivos:
            todos_los_archivos.add(os.path.join(raiz, nombre))

    filas_salida, filas_control, archivos_usados, tiempos = [], [], set(), []
    conteo_ok = {i: 0 for i in range(1, len(columnas_cruce) + 1)}
    conteo_fail = {i: 0 for i in range(1, len(columnas_cruce) + 1)}
    idx_global = 0

    for ruta_excel_actual, nombre_hoja, encabezados, encabezados_norm, registros in datos_por_hoja:
        # Si esta hoja en particular no tiene las columnas de cruce
        # elegidas, se omite completa (con aviso), no se detiene el lote.
        faltantes = [c for c in columnas_cruce if c not in encabezados]
        if faltantes:
            print(f"\nAVISO: '{os.path.basename(ruta_excel_actual)}' [hoja: {nombre_hoja}] no tiene la(s) "
                  f"columna(s) de cruce {faltantes}. Se omite ({len(registros)} filas).")
            continue

        print(f"\n--- Procesando: {os.path.basename(ruta_excel_actual)} [hoja: {nombre_hoja}] ---")
        for fila in registros:
            idx_global += 1
            fila_56, fila_malla, tecnicos, niveles, motivo, tiempo, ruta_archivo = procesar_fuid(
                fila, encabezados_norm, columnas_cruce, ruta_carpeta, indice_global, indice_global_lower)
            fila_extra = extraer_campos_resueltos(fila, encabezados_norm, resueltos)
            tiempos.append(tiempo)
            if ruta_archivo:
                archivos_usados.add(ruta_archivo)
            for nivel_info in niveles:
                n = nivel_info["nivel"]
                (conteo_ok if nivel_info["resultado"] == "Si" else conteo_fail)[n] += 1

            fila_completa = dict(fila_56)
            fila_completa.update(fila_malla)
            fila_completa.update(tecnicos)
            fila_completa.update(fila_extra)
            fila_completa = rellenar_por_nombre_directo(fila_completa, fila, encabezados_norm)
            fila_completa["inventario"] = os.path.basename(ruta_excel_actual)
            fila_completa["nombre_objeto_digital"] = os.path.basename(ruta_archivo) if ruta_archivo else "N/A"
            fila_completa["TIEMPO_PROCESAMIENTO_SEG"] = tiempo
            filas_salida.append(fila_completa)

            fila_ctrl = {"archivo_excel": os.path.basename(ruta_excel_actual), "hoja_excel": nombre_hoja,
                         "fila_excel_num": idx_global, "resultado_final": tecnicos["COINCIDENCIA_ARCHIVO"],
                         "motivo": motivo, "ruta_final": tecnicos["ruta_original"]}
            for nivel_info in niveles:
                n = nivel_info["nivel"]
                fila_ctrl[f"nivel_{n}_columna"] = columnas_cruce[n - 1]
                fila_ctrl[f"nivel_{n}_valor_buscado"] = nivel_info["valor_buscado"]
                fila_ctrl[f"nivel_{n}_resultado"] = nivel_info["resultado"]
                fila_ctrl[f"nivel_{n}_tipo_match"] = nivel_info.get("tipo_match") or "N/A"
                if "candidatos_ambiguos" in nivel_info:
                    fila_ctrl[f"nivel_{n}_candidatos_ambiguos"] = nivel_info["candidatos_ambiguos"]
            filas_control.append(fila_ctrl)

            print(f"  [{idx_global}/{total_filas_excel}] {tecnicos['COINCIDENCIA_ARCHIVO']} ({motivo}, {tiempo}s)")

    huerfanos = todos_los_archivos - archivos_usados
    for ruta_huerfano in sorted(huerfanos):
        filas_control.append({"archivo_excel": "N/A", "hoja_excel": "N/A", "fila_excel_num": "N/A", "resultado_final": "HUERFANO",
                               "motivo": "archivo_en_servidor_sin_fila_excel", "ruta_final": ruta_huerfano})

    tiempo_total = round(time.perf_counter() - t_inicio, 4)
    for fila in filas_salida:
        fila["TIEMPO_TOTAL_LOTE_SEG"] = tiempo_total

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    ruta_datos = f"cruce_FUID_{ts}.csv"
    ruta_control = f"control_FUID_{ts}.csv"

    # Campos de CAMPOS_MALLA_TODOS/CAMPOS_OTRAS_CATEGORIAS que ya tienen
    # hogar en CAMPOS_FINALES no se repiten al final.
    extras_sin_hogar = [c for c in (CAMPOS_MALLA_TODOS + CAMPOS_OTRAS_CATEGORIAS)
                        if c not in CAMPOS_FINALES]

    columnas_datos = (["inventario"] + CAMPOS_FINALES + extras_sin_hogar +
                       ["COINCIDENCIA_ARCHIVO", "mime_real",
                        "TIEMPO_PROCESAMIENTO_SEG", "TIEMPO_TOTAL_LOTE_SEG"] + campos_extra_dinamicos)
    columnas_unicas = nombres_columnas_unicos(columnas_datos)
    with open(ruta_datos, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(columnas_unicas)
        for fila in filas_salida:
            writer.writerow([fila.get(c, "N/A") for c in columnas_datos])

    orden_control = ["archivo_excel", "hoja_excel", "fila_excel_num", "resultado_final", "motivo", "ruta_final"]
    for i in range(1, len(columnas_cruce) + 1):
        orden_control += [f"nivel_{i}_columna", f"nivel_{i}_valor_buscado",
                           f"nivel_{i}_resultado", f"nivel_{i}_tipo_match",
                           f"nivel_{i}_candidatos_ambiguos"]
    with open(ruta_control, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=orden_control)
        writer.writeheader()
        for fila in filas_control:
            writer.writerow({k: fila.get(k, "N/A") for k in orden_control})

    ok = sum(1 for f in filas_salida if f["COINCIDENCIA_ARCHIVO"] == "Si")
    ambiguos = sum(1 for f in filas_control if f.get("motivo") == "ambiguo_multiples_candidatos")
    sin_coincidencia = len(filas_salida) - ok - ambiguos
    print("\n" + "=" * 60)
    print(" RESUMEN FUID")
    print("=" * 60)
    print(f"Archivos Excel encontrados: {len(archivos_excel)} | Hojas con datos procesadas: {len(datos_por_hoja)}")
    print(f"Filas totales: {len(filas_salida)}")
    print(f"Archivos reales en la carpeta (recursivo): {len(todos_los_archivos)}")
    print(f"Coincidencias completas: {ok}")
    print(f"Ambiguos (varios candidatos, sin elegir automaticamente): {ambiguos}")
    print(f"Sin coincidencia: {sin_coincidencia}")
    print(f"Huerfanos (en servidor, sin fila en Excel): {len(huerfanos)}")
    for i in range(1, len(columnas_cruce) + 1):
        print(f"Nivel {i} ({columnas_cruce[i-1]}): {conteo_ok[i]} coincidieron, {conteo_fail[i]} no")
    print(f"Tiempo total: {tiempo_total}s | Promedio por fila: "
          f"{round(sum(tiempos)/len(tiempos), 4) if tiempos else 0}s")
    print(f"\nArchivo de datos: {os.path.abspath(ruta_datos)}")
    print(f"Archivo de control: {os.path.abspath(ruta_control)}")


# ---------------------------------------------------------------------------
# MODO 2: MALLA DOCUMENTAL (Estructura 2)
# ---------------------------------------------------------------------------
def pedir_columna_nombre_archivo(encabezados):
    """Pregunta cual columna del Excel de Malla corresponde al nombre del
    archivo (la llave de cruce). Sugiere 'File name' si existe, pero deja
    escribir cualquier otro nombre si esa Malla en particular usa otro."""
    indice_encabezados = construir_indice_encabezados(encabezados)
    sugerencia = indice_encabezados.get(normalizar_encabezado("File name"))

    if sugerencia:
        return sugerencia

    print("\nNo se encontro una columna llamada 'File name'.")
    print("Columnas disponibles en este Excel:")
    for h in encabezados:
        print(f"  - {h}")
    while True:
        respuesta = input("\n¿Cual columna tiene el NOMBRE del archivo (para cruzar contra "
                           "el servidor)?: ").strip()
        if respuesta in encabezados:
            return respuesta
        print(f"  '{respuesta}' no esta en el Excel. Copia el nombre tal cual aparece arriba.")


def procesar_malla(fila_dict, encabezados_norm, indice, indice_lower, columna_nombre_archivo):
    t0 = time.perf_counter()

    fila_56 = {campo: "N/A" for campo in CAMPOS_56}
    fila_malla = {campo: "N/A" for campo in CAMPOS_MALLA_TODOS}
    fila_otras = {campo: "N/A" for campo in CAMPOS_OTRAS_CATEGORIAS}
    fila_56["tipo_inventario"] = "MALLA DOCUMENTAL"

    for header_malla, campo_norm, _origen in MAPEO_MALLA_DOCUMENTAL:
        valor = a_na(obtener_valor(fila_dict, encabezados_norm, header_malla))
        if campo_norm in fila_56:
            fila_56[campo_norm] = valor
        elif campo_norm in fila_otras:
            fila_otras[campo_norm] = valor
        else:
            fila_malla[campo_norm] = valor

    # nombre_archivo -> nombre_objeto, es la LLAVE de cruce contra el
    # servidor (no se confia en la ruta que trae la Malla, puede estar
    # desactualizada; el archivo puede haberse movido de lugar)
    nombre_objeto = obtener_valor(fila_dict, encabezados_norm, columna_nombre_archivo)
    fila_otras["nombre_objeto"] = a_na(nombre_objeto)

    ruta_encontrada, tipo_match = buscar_por_nombre(indice, indice_lower, nombre_objeto)
    if ruta_encontrada:
        fila_otras["ruta_original"] = ruta_encontrada
        coincidencia = "Si"
    else:
        fila_otras["ruta_original"] = "N/A"
        coincidencia = "No"

    tiempo = round(time.perf_counter() - t0, 4)
    return fila_56, fila_malla, fila_otras, coincidencia, tipo_match, ruta_encontrada, tiempo


def modo_malla():
    ruta_excel = input("Ruta del Excel/CSV de Malla Documental: ").strip('"')
    ruta_carpeta = input("Ruta de la carpeta en el servidor (para cruzar por nombre de archivo): ").strip('"')

    if not os.path.isfile(ruta_excel):
        print(f"ERROR: no se encontro el archivo: {ruta_excel}")
        return
    if not os.path.isdir(ruta_carpeta):
        print(f"ERROR: no se encontro la carpeta: {ruta_carpeta}")
        return

    t_inicio = time.perf_counter()
    encabezados, registros = leer_excel(ruta_excel)
    encabezados_norm = construir_indice_encabezados(encabezados)
    print(f"\nFilas leidas: {len(registros)}")

    # 1) Primero: detectar y registrar columnas sin mapear conocido, y
    #    aplicar las que el usuario ya haya resuelto manualmente en
    #    campos_sin_mapear.json (unico archivo de configuracion).
    reconocidos_malla = {normalizar_encabezado(h) for h, _c, _o in MAPEO_MALLA_DOCUMENTAL}
    reconocidos_malla.add(normalizar_encabezado("File name"))
    reconocidos_malla.add(normalizar_encabezado("File path"))
    nuevas = registrar_columnas_no_reconocidas(encabezados, reconocidos_malla, "MALLA DOCUMENTAL", ruta_excel)

    resueltos = cargar_mapeos_resueltos()
    campos_extra_dinamicos = sorted(set(resueltos.values()))
    if campos_extra_dinamicos:
        print(f"Aplicando {len(campos_extra_dinamicos)} campo(s) resuelto(s) manualmente: {campos_extra_dinamicos}")

    if nuevas > 0:
        print(f"\nSe encontraron {nuevas} columna(s) nueva(s) sin mapeo conocido.")
        respuesta = input("¿Continuar con el cruce de todas formas? (S/N, N para salir y "
                           "completar campos_sin_mapear.json primero): ").strip().upper()
        if respuesta != "S":
            print("Saliendo sin procesar. Completa campos_sin_mapear.json y vuelve a correr.")
            return

    # 2) Ahora si, preguntar por cual columna se va a cruzar contra el servidor.
    columna_nombre_archivo = pedir_columna_nombre_archivo(encabezados)
    print(f"\nCruzando usando la columna: '{columna_nombre_archivo}'")

    print(f"\nEscaneando (recursivo) el servidor: {ruta_carpeta}")
    indice, indice_lower, total_archivos = indexar_archivos_servidor(ruta_carpeta)
    print(f"  -> {total_archivos} archivos encontrados en el servidor.")

    filas_salida, filas_control, tiempos = [], [], []
    coinciden, no_coinciden = 0, 0

    for idx, fila in enumerate(registros, start=1):
        fila_56, fila_malla, fila_otras, coincidencia, tipo_match, ruta_encontrada, tiempo = procesar_malla(
            fila, encabezados_norm, indice, indice_lower, columna_nombre_archivo)
        fila_extra = extraer_campos_resueltos(fila, encabezados_norm, resueltos)
        tiempos.append(tiempo)
        if coincidencia == "Si":
            coinciden += 1
        else:
            no_coinciden += 1

        fila_completa = dict(fila_56)
        fila_completa.update(fila_malla)
        fila_completa.update(fila_otras)
        fila_completa.update(fila_extra)
        fila_completa = rellenar_por_nombre_directo(fila_completa, fila, encabezados_norm)
        fila_completa["inventario"] = os.path.basename(ruta_excel)
        fila_completa["nombre_objeto_digital"] = os.path.basename(ruta_encontrada) if ruta_encontrada else "N/A"
        fila_completa["COINCIDENCIA_ARCHIVO"] = coincidencia
        fila_completa["TIEMPO_PROCESAMIENTO_SEG"] = tiempo
        filas_salida.append(fila_completa)

        filas_control.append({
            "fila_excel_num": idx,
            "nombre_objeto_buscado": fila_otras.get("nombre_objeto", "N/A"),
            "resultado": coincidencia,
            "tipo_match": tipo_match or "N/A",
            "ruta_encontrada": ruta_encontrada or "N/A",
        })

        print(f"  [{idx}/{len(registros)}] {fila_otras.get('nombre_objeto')} -> {coincidencia} "
              f"({tipo_match or 'sin match'}, {tiempo}s)")

    tiempo_total = round(time.perf_counter() - t_inicio, 4)
    for fila in filas_salida:
        fila["TIEMPO_TOTAL_LOTE_SEG"] = tiempo_total

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    ruta_datos = f"cruce_MALLA_{ts}.csv"
    ruta_control = f"control_MALLA_{ts}.csv"

    extras_sin_hogar = [c for c in (CAMPOS_MALLA_TODOS + CAMPOS_OTRAS_CATEGORIAS)
                        if c not in CAMPOS_FINALES]

    columnas_datos = (["inventario"] + CAMPOS_FINALES + extras_sin_hogar +
                       ["COINCIDENCIA_ARCHIVO", "TIEMPO_PROCESAMIENTO_SEG", "TIEMPO_TOTAL_LOTE_SEG"] +
                       campos_extra_dinamicos)
    columnas_unicas = nombres_columnas_unicos(columnas_datos)
    with open(ruta_datos, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(columnas_unicas)
        for fila in filas_salida:
            writer.writerow([fila.get(c, "N/A") for c in columnas_datos])

    with open(ruta_control, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=["fila_excel_num", "nombre_objeto_buscado",
                                                 "resultado", "tipo_match", "ruta_encontrada"])
        writer.writeheader()
        for fila in filas_control:
            writer.writerow(fila)

    print("\n" + "=" * 60)
    print(" RESUMEN MALLA DOCUMENTAL")
    print("=" * 60)
    print(f"Filas procesadas: {len(filas_salida)}")
    print(f"Archivos encontrados en el servidor (coincidencia por nombre_objeto): {coinciden}")
    print(f"Sin coincidencia (posible archivo movido/eliminado/renombrado): {no_coinciden}")
    print(f"Tiempo total: {tiempo_total}s | Promedio por fila: "
          f"{round(sum(tiempos)/len(tiempos), 4) if tiempos else 0}s")
    print(f"\nArchivo de datos: {os.path.abspath(ruta_datos)}")
    print(f"Archivo de control: {os.path.abspath(ruta_control)}")


# ---------------------------------------------------------------------------
def main():
    print("=" * 60)
    print(" CRUCE DE METADATOS - v3")
    print("=" * 60)
    print("¿Qué tipo de fuente vas a procesar?")
    print("  1) FUID (Estructura 1 - se cruza contra el servidor)")
    print("  2) MALLA DOCUMENTAL (Estructura 2 - se cruza por nombre de archivo, "
          "la ruta que trae puede estar desactualizada)")
    while True:
        opcion = input("Elige 1 o 2: ").strip()
        if opcion in ("1", "2"):
            break
        print("Por favor escribe 1 o 2.")

    try:
        if opcion == "1":
            modo_fuid()
        else:
            modo_malla()
    except Exception as e:
        print(f"\nERROR INESPERADO: {e}")
        raise
    finally:
        if sys.stdin.isatty():
            try:
                input("\nProceso terminado. Presiona Enter para salir...")
            except EOFError:
                pass


if __name__ == "__main__":
    main()
