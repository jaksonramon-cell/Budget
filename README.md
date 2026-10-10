# Presupuestador CIREC

Herramienta para planear presupuestos de proyectos sin olvidar rubros. Uso personal, corre en el propio computador; los
datos son archivos JSON en `datos/`.

## Cómo abrirlo (Windows)

1. Instalar Python 3.11 o superior (python.org, marcar "Add Python to PATH").
2. Doble clic en `iniciar.bat`. La primera vez instala lo necesario; después abre el navegador en http://localhost:8080.

Otros sistemas: `pip install -r requirements.txt` y `python app.py`.

## La idea

Un evento no se presupuesta línea por línea: se describe **quién va, cómo llega y cuántos días**. Con eso la herramienta
deduce los tiquetes, el transporte, el hospedaje y la alimentación de cada grupo, más el salón y los kits.

Cada rubro que la logística exige termina en uno de tres estados:

| | |
|---|---|
| ✓ costeado | tiene cantidad y tarifa |
| — no necesario | tiene una razón escrita (ej. "la tarifa de hotel incluye transporte") |
| ✗ falta | ni costo ni razón: queda en rojo hasta resolverlo |

Ejemplos de lo que se marca en rojo: un grupo que duerme fuera pero no come, alguien que vuela sin transporte local, un
viaje de varios días sin hospedaje, participantes sin salón o sin nadie del equipo que facilite.

## Pestañas

- **Proyecto**: donante, moneda, TRM, indirectos y las actividades del marco lógico con sus metas.
- **Dónde y cuántos**: matriz lugares × actividades; cada celda dice cuántas veces ocurre. Clic en una celda para ver
  quién va y qué se costea, y ajustar solo lo que ese lugar tenga distinto.
- **Tipos de actividad**: la plantilla de cada tipo de evento (brigada, ENTRÉNATE, curso…). Se define una vez.
- **Personal y otros costos**: cargos × meses × dedicación, y costos que no dependen de un evento.
- **Revisión**: totales por categoría del donante, cuadrícula de logística por evento, lista de chequeo del proyecto
  (seguros, auditoría, visibilidad…) y exportación a Excel.
- **Banco**: tarifas, lugares, cargos y mapeo a categorías de CBPF, AECID y Fondo Multidonante.

## Ejemplos incluidos

- `PRESUP1_CBPF_aprobado`: el presupuesto CBPF aprobado reconstruido. Cuadra al centavo salvo E11 (+USD 193,13: el
  aprobado tiene 21 noches en Miraflores pero solo 14 días de alimentación) y muestra 10 preguntas abiertas.
- `FMD_brechas_ejemplo`: actividades 1.6, 3.3, 4.2 y 4.3 del presupuesto FMD tal como se presupuestaron, con las
  brechas que la revisión detecta.

## Para desarrollo

```
pip install -r requirements.txt
python -m pytest              # pruebas del motor y de los ejemplos
python -m presupuesto.semilla # regenera datos/ (SOBRESCRIBE banco y ejemplos)
```

- `presupuesto/modelo.py`: estructura de datos. `calculo.py`: motor y reglas. `exportar.py`: Excel. `app.py`: interfaz
  (NiceGUI). `semilla.py`: datos iniciales.
- `legacy/constructor_v4.3/`: el constructor en Excel anterior, como referencia.
- El mapeo de bloques a categorías de AECID y Fondo Multidonante es una propuesta a partir de los presupuestos
  Humanicemos y FMD: confirmarlo con cada donante.
