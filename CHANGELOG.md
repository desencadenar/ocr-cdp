# Cambios

## 1.1 · Detección y control más confiables

Lo que no se pudo comprobar ya no pasa por verificado. Lo dudoso llega a `revision` con un motivo claro y el programa lee mejor los comprobantes reales.

- **Detección:** no descarta comprobantes con estructura completa. La serie ya no se confunde con leyendas legales, teléfonos ni cuentas. El cliente se reconoce con rótulos variables o en columnas. Se leen la serie física numérica y las facturas en inglés.
- **Importes:** los formatos ambiguos quedan en blanco, se corrigen las lecturas del `S/` y hay aviso por baja confianza también en los importes. Las descripciones e ítems se leen mejor.
- **Controles:**
  - El cuadre distingue `cuadra` de `no_contradice` y el control de tasa es más estricto.
  - La tasa especial de restaurantes y hoteles se decide por el giro del emisor.
  - El QR solo coincide si de verdad comparó.
  - Se corrigen el aviso de retención en honorarios y el cotejo con el padrón.
- **Hojas:** la revisión es coherente entre hojas. Las ventas propias se reconocen y las etiquetas de crédito fiscal son prudentes.
- **Uso:** nuevo comando `redetectar`. Los ZIP ya no se retienen en memoria. Hay aviso de privacidad de `.ocr_fiscal/` y se limpió el código sin uso.

Detalle de estados y etiquetas en el [README](README.md#estados-y-etiquetas-que-cambiaron-en-la-versión-11).

## 1.0 · Primera versión pública
