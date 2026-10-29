# Cruce FUID — Instrucciones

El programa es `cruce_fuid.py`. El nombre del script no cambia.
Cada compilación genera un ejecutable con versión (`cruce_fuid_5.1.exe`,
`cruce_fuid_6.0.exe`, …). Las columnas del CSV son las mismas de siempre:
lo nuevo es que, si un campo sigue en `N/A` y el archivo cruzado es un PDF,
ese campo se completa con los metadatos del PDF.

## 1. Compilar el .exe

En un equipo con Python en el PATH, en esta misma carpeta:

1. Doble clic en `compilar_exe.bat`.
2. El `.bat` muestra la versión actual y pregunta:
   - `1` cambio menor: de `5.0` pasa a `5.1`
   - `2` cambio mayor: de `5.0` pasa a `6.0`
3. Instala `pyinstaller`, `openpyxl` y `pypdf`, y empaqueta `cruce_fuid.py`.
4. El ejecutable queda en `dist\`. Ejemplo: `dist\cruce_fuid_5.1.exe`.

La versión actual sale de `cruce_fuid.version` y de los `.exe` que ya estén
en `dist`. `cruce_fuid_v5.exe` cuenta como `5.0`. Si el `.exe` de esa
versión ya existe, el `.bat` se detiene y no lo vuelve a crear. Después de
una compilación correcta, el número queda guardado en `cruce_fuid.version`.

Copia solo el `.exe` a los equipos donde se va a usar. Esos equipos no
necesitan Python.

## 2. Usarlo

Doble clic en el `.exe`. Todo es por pantalla. Las rutas se pueden pegar
con o sin comillas.

```text
1) FUID (Estructura 1 - se cruza contra el servidor)
2) MALLA DOCUMENTAL (Estructura 2 - se cruza por nombre de archivo)
```

### Modo 1 — FUID

Pide el Excel (o una carpeta con varios Excel) y la carpeta del servidor.
Después pide cuántas columnas arman la ruta (caja, carpeta, …) y el nombre
exacto de cada una. El cruce es jerárquico: busca esa ruta dentro del
servidor.

Genera, en la carpeta desde donde se ejecutó:

- `cruce_FUID_AAAAMMDD_HHMMSS.csv` — la sábana
- `control_FUID_AAAAMMDD_HHMMSS.csv` — coincidencias, niveles y huérfanos

### Modo 2 — Malla documental

Pide el Excel o CSV de la malla y la carpeta del servidor. Cruza por el
nombre del archivo. La ruta que traiga la malla no se usa como ubicación
real.

Genera:

- `cruce_MALLA_AAAAMMDD_HHMMSS.csv`
- `control_MALLA_AAAAMMDD_HHMMSS.csv`

## 3. Columnas del CSV

El orden no cambió.

1. `inventario`
2. Los 184 campos de `CAMPOS_FINALES`, en el orden del esquema. Si un nombre
   se repite, el encabezado se numera (`volumen`, `volumen_2`, …).
3. Los campos de malla que no tienen casilla en esos 184, entre ellos
   `fecha_creacion` y `fecha_modificacion`.
4. `COINCIDENCIA_ARCHIVO`, `mime_real` (solo en FUID), los tiempos y las
   columnas resueltas en `campos_sin_mapear.json`.

Lo que ya trae el inventario se respeta. El PDF solo escribe donde el valor
sigue en `N/A`. `peso_bytes` y `peso_mb` salen del tamaño real del archivo
en disco. `hash` y `hash_origen` no se calculan.

## 4. Metadatos que se completan desde el PDF

Solo cuando el archivo cruzado es PDF y el campo está en `N/A`.

| Campo | De dónde sale |
|---|---|
| `autor` | `Author` del PDF. Si no está, `dc:creator` o `Autor:` dentro de Keywords |
| `formato` | `Formato:` dentro de Keywords |
| `tipo_documental` | `Tipo de documento:` dentro de Keywords |
| `paginas_pdf` | Número de páginas del PDF |
| `idioma` y `lengua_documentos` | `Idioma:` dentro de Keywords |
| `palabras_clave` | El texto completo de Keywords, sin partir |
| `fecha_creacion` | `xmp:CreateDate`, en hora local del equipo |
| `fecha_modificacion` | `xmp:ModifyDate`. Si el PDF no la trae, la fecha de modificación del archivo |
| `id_unico` | `xapMM:DocumentID` (Document ID del XMP) |
| `caracteristicas_requisitos_tecnicos` | Una sola frase: versión PDF, linealizado, PDF/A, herramienta, programa que escribió el archivo y toolkit XMP |

`Tamano:` dentro de Keywords no se usa: no coincide con el peso real.
`Language` del PDF (a menudo `en`) no se copia a `idioma`. `Producer` no se
copia a `productor`: el productor del inventario sigue siendo el de la
entidad. `fecha_inicial`, `fecha_final` y `numero_folios` tampoco se toman
del PDF.

Ejemplo de Keywords completo, que es lo que queda en `palabras_clave`:

```text
Autor:Camara de Representantes, Formato:pdf, Tipo de documento: Gaceta, Tamano:49MB, Cantidad de paginas:39, Idioma:es, Fecha del documento: 2025-05-05 21:40:17, Fecha del pdf:2025-05-05 21:40:17
```

## 5. Notas

- Si el PDF no se puede leer, esa fila sigue en `N/A` y el lote continúa.
- El mismo PDF cruzado en varias filas se abre una sola vez.
- La firma del archivo reconoce PDF, JPEG, PNG, TIFF, GIF, Office antiguo,
  DOCX/XLSX/PPTX, ZIP, RAR, 7Z, GZIP y BZIP2.
- Si la ruta del servidor no responde, el programa avisa antes de procesar.
- Una columna del Excel que el programa no reconoce se anota en
  `campos_sin_mapear.json`, junto al ejecutable. Al llenar ahí
  `campo_normalizado`, la siguiente corrida la agrega al final del CSV.
