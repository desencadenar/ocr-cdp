# OCR CDP

Herramienta local en Python pensada para contabilidad y auditoría tributaria en Perú. Recorre carpetas y archivos ZIP con representaciones visuales de comprobantes de pago **sin XML** (PDF digitales, híbridos, escaneados e imágenes `JPG`, `JPEG`, `PNG`, `TIF`/`TIFF`, `BMP` y `WebP`), extrae sus campos clave sin inventar datos dudosos y genera:

1. Un libro `CDP.xlsx` con seis pestañas orientadas al registro, revisión y *vouching* contable.
2. Opcionalmente, una carpeta `expediente_pdf/` ordenada por proveedor con anexos vinculados y alertas de peso para Mesa de Partes Virtual (MPV) de SUNAT.

Todo se procesa íntegramente en tu equipo, sin subir documentos a internet ni depender de servicios externos.

---

## Empezar

Necesitas **Python 3.11, 3.12 o 3.13**. Abre una terminal en la carpeta del proyecto y ejecuta:

```text
python iniciar.py
```

Al ejecutarse, `iniciar.py` crea una carpeta local `.venv`, instala las dependencias fijas de `requirements.txt` (solo requiere internet la primera vez) y abre la ventana principal del programa. En adelante puedes volver a iniciarlo con el mismo comando.

Para leer fotos y PDF escaneados, instala [Tesseract OCR](https://tesseract-ocr.github.io/tessdoc/Installation.html) activando los idiomas español e inglés (`spa+eng`). Si Tesseract no está instalado, el sistema sigue leyendo normalmente los PDF que ya traen capa de texto digital y deriva a `revision` las páginas que requieren OCR.

En la ventana gráfica puedes elegir:
- **Carpeta de entrada** con tus comprobantes o ZIPs.
- **Carpeta de salida** (vacía y fuera de la de entrada) donde se guardará `CDP.xlsx`.
- **Padrón RUC (SQLite, opcional)** para contrastar razones sociales en local.
- **RUC objetivo (receptor)** del contribuyente que estás auditando.
- Si deseas inspeccionar contenedores **ZIP** y si deseas generar la carpeta ordenada **`expediente_pdf/`**.

## Abrir el proyecto en VS Code desde cero

1. Instala [Visual Studio Code](https://code.visualstudio.com/), Python 3.11, 3.12 o 3.13 y Tesseract (`spa+eng`) si vas a procesar escaneos.
2. En VS Code ve a **Archivo → Abrir carpeta…** y selecciona la carpeta raíz del proyecto (la que contiene `iniciar.py`).
3. VS Code sugerirá las extensiones recomendadas en [extensions.json](.vscode/extensions.json): **Python** (`ms-python.python`) y **Markdown Preview Mermaid Support** (`bierner.markdown-mermaid`). Conviene aceptar ambas para ver el diagrama de este `README.md` (`Ctrl+Shift+V`).
4. Abre una terminal integrada y ejecuta `python iniciar.py` para inicializar `.venv`. Luego selecciona el intérprete de `.venv` en VS Code (`Ctrl+Shift+P` → *Python: Select Interpreter*).
5. Con `F5` puedes ejecutar las configuraciones predefinidas en [launch.json](.vscode/launch.json): abrir la ventana gráfica o lanzar un análisis directo por carpeta.

En GitHub los diagramas Mermaid se dibujan automáticamente sin instalar extensiones.

---

## Herramientas y librerías que utiliza

| Componente | Para qué se usa en el proyecto | Dónde está definido |
| --- | --- | --- |
| `PyMuPDF` (`1.28.2`) | Lectura de texto nativo, detección de columnas y glifos dañados, renderizado de páginas para OCR, recorte/unión de PDFs en `expediente_pdf/`. | [requirements.txt](requirements.txt) |
| `Pillow` (`12.3.0`) | Carga de imágenes y TIFF multipágina, corrección de orientación EXIF y recompresión adaptativa de imágenes hacia PDF. | [requirements.txt](requirements.txt) |
| `openpyxl` (`3.1.5`) y `et-xmlfile` (`2.0.0`) | Construcción de `CDP.xlsx` (6 pestañas, formatos de texto, moneda y fechas reales) y plantillas de calibración. | [requirements.txt](requirements.txt) |
| `Tesseract OCR` *(programa externo)* | Reconocimiento óptico de caracteres por subproceso con salida TSV palabra por palabra y detección automática de rotación (`--psm 0`). | Configurado en [config.json](config.json) |
| Módulos estándar de Python | `sqlite3` (estado de corrida en `.ocr_fiscal/state.sqlite3` y Padrón RUC local), `zipfile` (apertura segura en memoria de ZIPs anidados), `tkinter` (ventana gráfica) y `unittest` (pruebas). | Biblioteca estándar de Python |
| Extras opcionales del entorno | Si instalas por separado `zxingcpp` en tu Python, decodifica códigos QR SUNAT (incluso vectoriales); si instalas `pillow-heif`, admite imágenes `.heic`. | Opcionales (no incluidos en [requirements.txt](requirements.txt)) |

---

## Mapa de lo que hace

```mermaid
flowchart TD
    A["Carpeta de entrada y ajustes en config.json"] --> B["Recorre archivos y ZIPs anidados con topes de seguridad"]
    B --> C{"¿Qué formato tiene el archivo?"}
    C -->|"XML"| D["Se excluye por formato y queda anotado en inventario"]
    C -->|"PDF o imagen"| E{"¿El PDF tiene texto digital utilizable?"}
    E -->|"Sí"| F["Extrae líneas, columnas y diseño con PyMuPDF"]
    E -->|"No o página mixta"| G["Renderiza a imagen y ejecuta Tesseract OCR"]
    F --> H["Lee códigos QR SUNAT si zxingcpp está disponible"]
    G --> H
    H --> I["Segmenta bloques y separa documentos de soporte no tributarios"]
    I --> J{"¿Cómo clasifica el bloque?"}
    J -->|"Soporte no CDP"| K["Registra soporte: guía, OC, HES, estado bancario o constancia SPOT"]
    J -->|"Indeterminado o error de lectura"| L["Deriva a revision con página y fragmento para cotejo manual"]
    J -->|"CDP detectado (17 tipos SUNAT)"| M["Extrae identidad, RUC/DNI, glosa, ítems y datos de OC, Guía, Placa y SPOT"]
    M --> N["Cuadra importes con tolerancia y evalúa tasa IGV (18%, Rég. esp. restaurantes/hoteles 10/10.5/15%, IVAP 4%, 0%)"]
    N --> O["Reconcilia con Padrón RUC local, coteja QR y deduplica copias por clave fiscal"]
    O --> P["Genera CDP.xlsx con 6 pestañas y trazabilidad por id_documento"]
    O --> Q["Opcional: construye expediente_pdf/ ordenado por proveedor y topes 10 MB MPV"]
    K --> P
    K --> Q
    L --> P
    L --> Q
```

---

## Qué entrega en `CDP.xlsx`

El libro generado contiene **seis pestañas**, con encabezados en minúsculas separados por guiones bajos y una columna común `id_documento` para cruzar cualquier fila entre hojas:

| Pestaña | Qué contiene y para qué sirve |
| --- | --- |
| `cdp_detectados` | Una fila por representación principal de cada comprobante detectado (salvo notas `07` y `08`). Incluye RUC y razón social observada, tipo/serie/correlativo, emisión, vencimiento, condición de pago, cuotas, documento del adquirente (`RUC` o `DNI`), bases e importes, cuadre aritmético, tasa efectiva y control de IGV, `apto_credito_fiscal_preliminar`, `alerta_bancarizacion`, referencias citadas (`orden_de_compra_citada`, `guia_de_remision_citada`, `placa_vehicular`), datos SPOT (`%`, código y cuenta BN), cotejo con Padrón RUC al corte, peso del PDF en MB, alerta `>10 MB MPV` y anexos vinculados del mismo PDF. |
| `notas` | Notas de crédito (`07`) y débito (`08`) con el código y motivo SUNAT leído, uno o varios comprobantes afectados citados, si dichos comprobantes aparecen dentro del lote del mismo emisor, e importes con signo negativo en notas de crédito. |
| `revision` | Todo hallazgo que exige verificación visual humana con ubicación exacta y fragmento: lecturas con baja confianza OCR, identidades incompletas o en conflicto, diferencias aritméticas o de suma de ítems, RUC adquirente distinto del objetivo, posible venta propia mezclada, RUC ausente en padrón local, diferencias contra QR SUNAT, tasas IGV atípicas, RHE mayores a S/ 1,500 sin retención ni mención de suspensión (Form. 1609) y PDFs que superan 10 MB. |
| `inventario` | Inventario completo de archivos físicos y miembros de ZIP revisados, formatos excluidos (como XML), incidencias de contenedores (cifrados, corruptos o límites) y bloques descartados como soportes no CDP (`no_cdp`) indicando si quedaron archivados o vinculados a un comprobante. |
| `detalle_items` | Desglose fila por fila de productos o servicios leídos dentro de la tabla del comprobante (`id_documento`, clave fiscal, cantidad numérica preservando `0.0`, unidad, descripción, valor unitario, descuento e importe de línea), útil para clasificar cuentas del Plan Contable o centros de costo. |
| `resumen` | Totales de la corrida, contadores de revisión/tercero/`>10 MB`, canal de verificación y advertencias explícitas de alcance legal y tributario. Aparece siempre al final del libro. |

---

## Cobertura tributaria, tasas de IGV y cuadre de importes

### Comprobantes SUNAT reconocidos (17 códigos)
- `01` Factura
- `02` Recibo por honorarios
- `03` Boleta de venta *(incluye captura de `DNI` de 8 dígitos del cliente cuando no hay RUC)*
- `04` Liquidación de compra
- `05` Boleto de transporte aéreo *(incluye «Comprobante de compra» aéreo con contexto IATA/ticket)*
- `07` Nota de crédito
- `08` Nota de débito
- `10` Recibo por arrendamiento (Formulario Virtual 1683)
- `12` Ticket de máquina registradora
- `13` Documento autorizado SBS *(diferenciado de estados de cuenta bancarios cuando incluye RUC + IGV)*
- `14` Recibo por servicios públicos *(protegido contra falsos descartes cuando menciona saldo o estado de cuenta anterior)*
- `16` Boleto de viaje terrestre
- `20` Comprobante de retención
- `40` Comprobante de percepción
- `50` Declaración Aduanera de Mercancías (DUA/DAM)
- `52` Despacho simplificado aduanero
- `54` Liquidación de cobranza aduanera

En tipos extendidos (`05`, `10`, `12`, `13`, `14`, `16`, `50`, `52`, `54`) el correlativo admite hasta 20 dígitos ASCII sin truncarse; en tipos estándar (`01`, `02`, `03`, `04`, `07`, `08`, `20`, `40`) exige entre 1 y 8 dígitos y los normaliza a 8 posiciones.

### Control de tasa efectiva de IGV (informativo)
Esta validación vive y se configura directamente en la función `evaluate_igv_rate` de [amounts.py](src/ocr_fiscal/amounts.py) (sin requerir parámetros extra en [config.json](config.json)). Además del cuadre aritmético exacto, contrasta la tasa conjunta facturada (**IGV + IPM**) frente a las tasas vigentes en Perú:
- **18 %**: tasa estándar del Régimen General (aun cuando la Ley N.° 32387 redistribuya gradualmente puntos entre IGV e IPM, el total facturado sigue siendo `18 %`).
- **Tasa especial del IGV para restaurantes, hoteles y alojamientos turísticos (Ley N.° 31556 y modificatoria Ley N.° 32219, junto al IPM de la Ley N.° 32387):** **10 %** hasta 2025 (`8 % IGV + 2 % IPM`), **10.5 %** en 2026 (`8 % IGV + 2.5 % IPM`) y **15 %** en 2027 (`12 % IGV + 3 % IPM`), retornando a `18 %` desde 2028 salvo prórroga (en notas `07` y `08` también se toleran tasas de años previos). **Importante:** este régimen **no aplica a cualquier MYPE**; solo alcanza a micro y pequeñas empresas (ingresos hasta `1,700 UIT` y sin vinculación con grandes grupos) cuyo giro principal —mínimo `70 %` de sus ingresos— sea restaurantes, hoteles o alojamientos turísticos, excluyendo servicios de *catering* y concesionarios de alimentos. Como la herramienta no consulta la ficha RUC/CIIU de actividad del emisor, acepta numéricamente estas tasas cuando coinciden con el año de emisión y deja la verificación del giro del proveedor al auditor contable.
- **4 %**: tasa especial del IVAP.
- **0 %**: operaciones exoneradas, inafectas o de exportación (y reconocimiento específico de IGV referencial en transferencias gratuitas sin derecho a crédito fiscal).

Si una tasa no corresponde a ninguno de los supuestos esperados, no se altera el cuadre aritmético, pero se deja constancia en `control_tasa_igv` y en `revision`.

### Reglas y casos especiales en el cuadre (`amounts.py`)
- Tolerancia de redondeo configurable mediante `tolerancia_redondeo` (`0.02` por defecto), marcando el resultado como `cuadra_con_redondeo` cuando aplica.
- Soporte explícito para operaciones 100 % gratuitas, percepción sumada o detracción restada sobre `total_a_pagar`, RHE con retención del 8 % (inclusive entre paréntesis), ISC incluido en la base imponible, FISE contenido en operaciones inafectas, anticipos deducidos al pie, signos mixtos en Notas de Crédito y Nota de Crédito Tipo 13 en cero.
- Si `currency` resulta contradictorio dentro del bloque (`S/` y `USD` mezclados) o el OCR pega símbolos ambiguos de moneda al primer dígito (por ejemplo `$755.00`), el dato se deja en blanco para revisión en lugar de suponerse.

---

## Organización opcional en `expediente_pdf/`

Cuando activas `--exportar-pdf` en consola, la casilla correspondiente en la ventana o `"exportar_pdf_ordenado": true` en [config.json](config.json), se genera la subcarpeta `expediente_pdf/` con esta estructura:

- `01_cdp_por_proveedor/<RUC>-<PROVEEDOR>/<RUC>-<TIPO>-<SERIE>-<CORRELATIVO>.pdf`: comprobantes con identidad completa y sin alertas bloqueantes. Por defecto incluye también sus notas `07` y `08` en la misma carpeta del proveedor (`unificar_notas_en_proveedor: true`). Si un mismo PDF contenía anexos contiguos clasificados como soporte (guía, OC, constancia SPOT), los extrae en `<clave>_SOPORTES_pX-Y.pdf` dentro de la misma carpeta del proveedor en lugar de separarlos.
- `02_notas_por_proveedor/`: se usa para las notas `07`/`08` solo cuando configuras `"unificar_notas_en_proveedor": false`.
- `03_revision_pendiente/`: comprobantes incompletos, indeterminados, conflictos de clave o colisiones (`COLISION_`), más derivaciones automáticas por riesgo tributario directo con prefijo visible: `TERCERO_` (RUC receptor distinto al objetivo), `VENTA_PROPIA_` (RUC emisor igual al RUC objetivo en un lote de compras) y `QR_DIFIERE_` (diferencia entre el QR SUNAT y el texto extraído).
- `04_soportes_no_cdp/`: documentos de soporte sueltos (guías, OC, HES, estados de cuenta bancarios, constancias SPOT) que llegaron en archivos separados sin un CDP apto en el mismo PDF.
- `05_duplicados_secundarios/`: copias secundarias concordantes cuando el mismo comprobante vino repetido en carpeta y en ZIP.

Las imágenes convertidas a PDF se recomprimen adaptativamente si superan **10 MB** (límite usual de Mesa de Partes Virtual SUNAT) y todo PDF exportado que supere ese peso queda marcado en `cdp_detectados` y `revision`.

---

## Alcance y qué no hace el programa

- **No valida emisión oficial en SUNAT:** `estado_oficial` figura siempre como `pendiente` y «CDP confirmados oficialmente» en `resumen` es `0`. El programa no se conecta a Clave SOL ni a API CPE.
- **El Padrón RUC local es estático:** si configuras una base SQLite creada con `preparar-padron`, solo sirve para comprobar existencia del RUC en esa instantánea, comparar razón social de emisor y receptor (tolerando ruido leve de OCR, abreviaturas y orden invertido en RUC `10`, además de distinguir `posible_nombre_comercial` frente a `discrepancia`), y mostrar estado/condición referencial **a la fecha de corte del padrón**. No acredita `HABIDO` ni `ACTIVO` vigente en la fecha real de emisión.
- **No reemplaza el control tributario ni la fehaciencia:** columnas como `apto_credito_fiscal_preliminar` y `alerta_bancarizacion` (umbral Ley 28194: `>= S/ 2,000` o `>= US$ 500`) son alertas preventivas para facilitar el *vouching*; no comprueban pago bancarizado real, depósito de detracción SPOT ni causalidad del gasto.
- **Privacidad total sin IA ni red externa:** todo corre localmente en tu equipo. El sistema no integra modelos LLM ni RAG, no realiza llamadas de red ni envía documentos o fragmentos fuera de tu computadora.

---

## Ajustar la lectura (`config.json`)

Abre [config.json](config.json) con cualquier editor de texto y cambia el valor a la derecha de la llave. Un formato JSON inválido indica la línea y columna a corregir.

| Ajuste | Valor inicial | Qué controla |
| --- | --- | --- |
| `ruta_tesseract` | `"auto"` | Ruta al ejecutable `tesseract`. Con `"auto"` lo busca en `TESSERACT_CMD`, `PATH` o carpetas estándar de Windows. |
| `idiomas_ocr` | `"spa+eng"` | Idiomas de Tesseract separados por `+` (deben estar instalados en el equipo). |
| `segundos_por_pagina` | `90` | Tiempo máximo (entero ≥ 1) que Tesseract puede tardar por página antes de mandar la página a `revision`. |
| `umbral_confianza_ocr` | `70` | Puntaje mínimo (`0` a `100`). Las lecturas clave por debajo del umbral conservan lo leído y generan hallazgo en `revision`. |
| `resolucion_pdf` | `220` | DPI (`72` a `600`) con los que se renderiza cada página PDF que necesite OCR. Los PDF con texto digital suficiente no pasan por OCR. |
| `maximo_pixeles` | `40000000` | Límite de píxeles permitido por imagen. Las imágenes mayores no se leen y van a `revision`. |
| `releer_recuadro_importes` | `false` | Con `true`, intenta una segunda lectura a mayor resolución sobre el recuadro de importes en facturas escaneadas de una sola página cuyo cuadre falló. |
| `revisar_archivos_zip` | `true` | Abre contenedores `.zip` (y ZIPs anidados) aplicando límites de profundidad, miembros, tamaño total y razón de compresión. |
| `exportar_pdf_ordenado` | `false` | Con `true`, genera automáticamente la carpeta `expediente_pdf/` al finalizar el análisis. |
| `padron_ruc` | `null` | Ruta a la base SQLite local del Padrón RUC (`null` para procesar sin padrón). Las rutas relativas parten de la carpeta del JSON. |
| `padron_dias_maximos` | `null` | Antigüedad máxima permitida del padrón en días (entero ≥ 1, o `null` sin límite). Si se excede, marca los RUC como `obsoleto`. |
| `unificar_notas_en_proveedor` | `true` | Con `true`, agrupa las notas `07`/`08` dentro de `01_cdp_por_proveedor/`; con `false`, las separa en `02_notas_por_proveedor/`. |
| `tolerancia_redondeo` | `0.02` | Diferencia máxima aceptada (`0` a `0.5`) entre el total calculado y el leído para considerar que cuadra con redondeo. |

---

## Uso desde la terminal

Reemplaza `RUC_RECEPTOR` por el RUC de 11 dígitos que deseas contrastar y usa siempre una carpeta de salida vacía y fuera de la carpeta de entrada:

```text
python iniciar.py scan --input entrada --output salida --recipient-ruc RUC_RECEPTOR --exportar-pdf
python iniciar.py scan --input entrada --output salida --recipient-ruc RUC_RECEPTOR --padron-db padron.sqlite3 --exportar-pdf
python iniciar.py export --output salida
```

- `ui`: abre la ventana gráfica (es el comando por defecto si ejecutas `python iniciar.py` sin argumentos).
- `scan`: procesa la carpeta de entrada y genera `CDP.xlsx` (con `--config`, `--padron-db`, `--no-inspect-archives` y `--exportar-pdf` opcionales).
- `export`: regenera `CDP.xlsx` desde `.ocr_fiscal/state.sqlite3` sin volver a leer los PDF ni ejecutar OCR.
- `preparar-padron`: convierte el archivo oficial `.txt` o `.zip` del Padrón Reducido de SUNAT a una base SQLite local lista para usar:

```text
python iniciar.py preparar-padron --input padron_reducido_ruc.zip --output padron.sqlite3
```

- `plantilla-calibracion` y `calibrar`: permiten crear una referencia etiquetada manualmente y comparar distintos perfiles de `config.json` campo por campo:

```text
python iniciar.py plantilla-calibracion --output referencia.xlsx
python iniciar.py calibrar --input muestra --reference referencia.xlsx --output calibracion.xlsx --recipient-ruc RUC_RECEPTOR --config config.json --config ajustes-b.json
```

---

## Pruebas automáticas

El repositorio incluye pruebas sintéticas locales (sin usar comprobantes reales ni exigir Tesseract ni conexión a red) en [test_ocr_fiscal.py](tests/test_ocr_fiscal.py). Para ejecutarlas:

```text
.venv\Scripts\python.exe -m unittest discover -s tests -v
```

*(En Linux o macOS usa `.venv/bin/python -m unittest discover -s tests -v`).*

---

## Leer el código

Los archivos de [src/ocr_fiscal](src/ocr_fiscal) están comentados en español, paso a paso, explicando el criterio técnico y tributario detrás de cada bloque. El orden recomendado para recorrer el código es:

1. [iniciar.py](iniciar.py) y [cli.py](src/ocr_fiscal/cli.py): preparación del entorno `.venv` y subcomandos CLI.
2. [ui.py](src/ocr_fiscal/ui.py) y [config.py](src/ocr_fiscal/config.py): ventana Tkinter y carga/validación de `config.json`.
3. [models.py](src/ocr_fiscal/models.py) y [tax.py](src/ocr_fiscal/tax.py): estructuras de datos, catálogo de 17 tipos SUNAT, validación de RUC/DNI y normalización de correlativos.
4. [archives.py](src/ocr_fiscal/archives.py) y [ocr.py](src/ocr_fiscal/ocr.py): recorrido seguro de carpetas/ZIPs e interfaz de subproceso con Tesseract.
5. [documents.py](src/ocr_fiscal/documents.py) y [amount_panel.py](src/ocr_fiscal/amount_panel.py): decisión entre texto nativo PyMuPDF y OCR, lectura de QR y relectura opcional de recuadro de importes.
6. [amounts.py](src/ocr_fiscal/amounts.py), [visible.py](src/ocr_fiscal/visible.py) y [extraction.py](src/ocr_fiscal/extraction.py): extracción de importes, tasas variables de IGV, leyendas/datos SPOT, clasificación de comprobantes vs soportes y detalle de ítems.
7. [padron.py](src/ocr_fiscal/padron.py) y [storage.py](src/ocr_fiscal/storage.py): consulta SQLite local del Padrón RUC y persistencia de la corrida en `.ocr_fiscal/state.sqlite3`.
8. [processor.py](src/ocr_fiscal/processor.py), [pdf_organizer.py](src/ocr_fiscal/pdf_organizer.py), [exports.py](src/ocr_fiscal/exports.py) y [benchmark.py](src/ocr_fiscal/benchmark.py): orquestación, cotejo QR, deduplicación, armado de `expediente_pdf/`, generación de `CDP.xlsx` y calibración.

---

## Contribuir y licencia

Este proyecto es de código abierto bajo la licencia **GNU Affero General Public License v3.0** (ver [LICENSE](LICENSE)). Las dependencias y herramientas externas conservan sus propias licencias (`PyMuPDF` bajo AGPL-3.0, `Tesseract` bajo Apache-2.0).

Si deseas reportar un caso o proponer mejoras:
- **Nunca adjuntes comprobantes reales** ni archivos con datos de terceros; describe el caso sustituyendo razones sociales, RUC e importes por datos ficticios.
- Indica tu versión de Python, de Tesseract y el contenido relevante de `config.json`.
- Todo ajuste en reglas de extracción debe respetar la regla base del proyecto: si un dato no puede comprobarse en el documento, queda en blanco y pasa a `revision`, sin suposiciones.
