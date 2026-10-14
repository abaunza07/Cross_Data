# Cruce FUID vs Servidor — Instrucciones

## 1. Compilar el .exe (solo una vez, en un equipo con Python)

1. Copia estos 2 archivos a una misma carpeta en el equipo donde tengas Python:
   - `cruce_fuid_servidor.py`
   - `compilar_exe.bat`
2. Haz doble clic en `compilar_exe.bat`.
3. Espera a que termine. Va a aparecer una carpeta `dist` con el archivo
   `cruce_fuid_servidor.exe` adentro.
4. Copia SOLO ese `.exe` a los equipos donde lo vas a usar. No necesitan
   tener Python instalado.

## 2. Usarlo

### Opción A — Doble clic (modo interactivo)
Haz doble clic en `cruce_fuid_servidor.exe`. Te va a pedir:
```
Ruta del Excel FUID:
Ruta de la carpeta en el servidor:
```
Puedes pegar las rutas directamente (con o sin comillas).

### Opción B — Desde línea de comandos (más rápido para repetir con las 8 actas)
```
cruce_fuid_servidor.exe --excel "ACTA1 ENTREGABLE PROYECTOS DE LEY 1.xlsx" --carpeta "\\IP\Gestion_Documental3\difusion\dt\IMAGENES_PL_1" --salida "cruce_ACTA1.csv"
```

Repite cambiando `--excel`, `--carpeta` (IMAGENES_PL_2, PL_3... según cada acta)
y `--salida` para las otras 7 actas.

## 3. Qué genera

Un solo CSV por corrida, con **una fila por acta** y estas columnas, en este orden:

1. **Las 165 columnas del esquema de clasificación de metadatos**
   (`Clasificacion_Metadatos_AGN.xlsx`), en el mismo orden. Como algunos
   nombres se repiten en ese esquema (ej. `Cargo` aparece 4 veces, `Notas`
   3 veces), el CSV los numera automáticamente: `Cargo`, `Cargo_2`,
   `Cargo_3`, `Cargo_4`.
2. `COINCIDENCIA_ARCHIVO` (`Si`/`No`), `TIEMPO_PROCESAMIENTO_SEG` (por fila),
   `TIEMPO_TOTAL_LOTE_SEG` (total de la corrida, igual en todas las filas).
3. Cualquier columna del Excel FUID que NO tenga equivalente entre las 165
   (por ejemplo `FASE`, `DEPENDENCIA`, `AG`, `AC`, `DP ELIMINAR`...) se
   agrega al final con el prefijo `EXTRA_FUID__` para que no se pierda nada.

## 4. Cómo se llena cada una de las 165 columnas

- **Del FUID**: si existe una columna equivalente en el Excel de la acta
  (ej. `CÓDIGO DE SERIE` -> `SER`, `CodSerie`, `Serie`...), se copia el
  valor tal cual.
- **Del escaneo del servidor**: los campos técnicos (`Volumen`,
  `Ruta Original`, `Tamaño Bytes`, `Tamaño MB`, `Extension Declarada`,
  `MIME Real`, `Extension Real`, `Descripcion Tipo`) se calculan leyendo el
  archivo real en el servidor — detectando el tipo por firma binaria, no
  solo por la extensión declarada.
- **Ya existentes — NUNCA se recalculan**: `Hash`, `HASH ORIGEN`, `TipoHash`
  y `LINEA` se copian del Excel **tal cual vienen**, si existen. El programa
  jamás calcula un hash sobre el archivo real ni genera estos valores por su
  cuenta. Si no están en el Excel, quedan en `N/A`.
- **Calculados (derivaciones simples)**: `AÑO` (se extrae de `FECHA
  INICIAL`), `PAPEL`/`ELECTRÓNICO`/`OTROS` (se marca `X` según el valor de
  `SOPORTE`), `MetadataSeguridad` / `Condiciones de acceso` (se deriva de
  las columnas de confidencialidad de serie/subserie del FUID).
- **Sin fuente de datos**: `N/A`.

## 5. Notas importantes

- **Nombres duplicados en el servidor**: si dos archivos con el mismo
  nombre existen en subcarpetas distintas, el programa avisa en pantalla y
  usa el primero que encuentra. Si esto pasa seguido, dímelo y ajustamos la
  lógica (por ejemplo, priorizar por fecha de modificación o por ruta).
- **Detección de tipo real**: cubre los formatos más comunes en gestión
  documental (PDF, JPEG, PNG, TIFF, GIF, Office antiguo, DOCX/XLSX/PPTX,
  ZIP, RAR, 7Z, GZIP, BZIP2). Si tus actas incluyen formatos raros que no
  reconoce (dirá "Tipo no identificado por firma binaria"), dímelo y agrego
  esa firma.
- **Rutas de red caídas**: si el servidor no responde o la ruta está mal
  escrita, el programa avisa con un mensaje claro antes de procesar nada.
- **Mapeo revisable**: el mapeo campo→fuente vive en el diccionario `MAPEO`
  al inicio de `cruce_fuid_servidor.py`, con un comentario en cada línea. Si
  al ver el CSV real de las 8 actas notas que algún campo debería salir de
  otra columna del FUID (o que alguno que hoy es `N/A` sí tiene dato en
  otro lado), dime cuál y lo ajusto ahí.
