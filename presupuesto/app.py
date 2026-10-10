"""Interfaz (NiceGUI). Un solo usuario, datos en archivos JSON locales.

Pestañas:
  Proyecto · Dónde y cuántos · Tipos de actividad · Personal y otros · Revisión · Banco
"""
from __future__ import annotations

import json
import re
import traceback
from pathlib import Path
from typing import Callable, Optional

from nicegui import events, ui
from pydantic import ValidationError

from .almacen import cargar_banco, cargar_proyecto, guardar
from .calculo import calcular, fmt, grupos_efectivos
from .exportar import exportar
from .modelo import (NOMBRE_RUBRO, RUBROS_GRUPO, Actividad, Banco, Cargo, Evento, Grupo, Kit, LineaPersonal, Lugar,
                     OtroCosto, Proyecto, Tarifa, TipoActividad)

RAIZ = Path(__file__).resolve().parent.parent
DATOS = RAIZ / "datos"
PROYECTOS = DATOS / "proyectos"
SALIDAS = RAIZ / "salidas"

LLEGADA = {"sede": "Ya está allí", "local": "Local (transporte local)", "terrestre": "Por tierra desde otro municipio",
           "aereo": "En avión"}
ICONO = {"ok": "✓", "no_necesario": "—", "falta": "✗", "sin_tarifa": "✗", "na": ""}
COLOR = {"ok": "text-green-700", "no_necesario": "text-gray-500", "falta": "text-red-700 font-bold",
         "sin_tarifa": "text-red-700 font-bold", "na": "text-gray-300"}
LOCALE = {"noRowsToShow": "Sin filas todavía", "loadingOoo": "Cargando…"}


# ============================================================================ estado
class Estado:
    def __init__(self):
        self.banco: Banco = cargar_banco(DATOS / "banco.json")
        self.ruta: Path = self._inicial()
        self.p: Proyecto = cargar_proyecto(self.ruta)
        self.res = calcular(self.banco, self.p)
        self.sel: Optional[tuple[str, str]] = None   # celda de la matriz: (actividad, lugar)
        self.lugares_extra: list[str] = []
        self.al_cambiar: list[Callable] = []

    @staticmethod
    def _inicial() -> Path:
        files = sorted(PROYECTOS.glob("*.json"), key=lambda f: f.stat().st_mtime, reverse=True)
        return files[0] if files else PROYECTOS / "Proyecto_nuevo.json"

    def abrir(self, ruta: Path):
        self.ruta, self.p = ruta, cargar_proyecto(ruta)
        self.sel, self.lugares_extra = None, []
        self.cambio(guardar_p=False)

    def cambio(self, guardar_p=True, guardar_b=False):
        if guardar_p:
            guardar(self.p, self.ruta)
        if guardar_b:
            guardar(self.banco, DATOS / "banco.json")
        self.res = calcular(self.banco, self.p)
        for f in self.al_cambiar:
            f()

    # búsquedas que respetan las sustituciones del proyecto
    def tipos(self) -> dict[str, TipoActividad]:
        d = {t.nombre: t for t in self.banco.tipos}
        d.update({t.nombre: t for t in self.p.tipos})
        return d

    def lugares(self) -> dict[str, Lugar]:
        d = {l.nombre: l for l in self.banco.lugares}
        d.update({l.nombre: l for l in self.p.lugares})
        return d

    def tipo_evento(self, ev: Evento) -> Optional[TipoActividad]:
        act = next((a for a in self.p.actividades if a.codigo == ev.actividad), None)
        return self.tipos().get(ev.tipo or (act.tipo if act else None) or "")

    def ids_tarifa(self) -> list[str]:
        return sorted({t.id for t in self.banco.tarifas} | {t.id for t in self.p.tarifas})

    def m(self, cop: float) -> str:
        dec = 0 if self.p.parametros.moneda == "COP" else 2
        return f"{fmt(cop / self.res.fx, dec)} {self.p.parametros.moneda}"


E: Estado = None  # se crea en construir()


# ============================================================================ utilidades
def orden_codigo(c: str):
    return [int(x) if x.isdigit() else 0 for x in c.split(".")]


def limpiar(d: dict) -> dict:
    """Celdas vacías -> None y quita campos internos de la grilla."""
    return {k: (None if v == "" else v) for k, v in d.items() if not k.startswith("_")}


def a_modelos(cls, filas: list[dict]):
    out = []
    for f in filas:
        f = {k: v for k, v in limpiar(f).items() if v is not None}
        if not any(v not in (None, "", 0) for v in f.values()):
            continue  # fila vacía
        out.append(cls.model_validate(f))
    return out


def numero(texto: str):
    t = texto.strip().replace("$", "").replace(" ", "")
    if re.fullmatch(r"-?[\d.]+,\d+", t):         # 1.234,56
        t = t.replace(".", "").replace(",", ".")
    elif re.fullmatch(r"-?\d{1,3}(\.\d{3})+", t):  # 1.234.567
        t = t.replace(".", "")
    else:
        t = t.replace(",", "")
    try:
        f = float(t)
        return int(f) if f.is_integer() else f
    except ValueError:
        return texto.strip()


def col(field, nombre, ancho=110, editable=True, lista=None, num=False, **kw):
    c = {"field": field, "headerName": nombre, "width": ancho, "editable": editable}
    if lista is not None:
        c.update(cellEditor="agSelectCellEditor", cellEditorParams={"values": lista})
    if num:
        c.update(cellEditor="agNumberCellEditor", type="rightAligned")
        c[":valueFormatter"] = "p => (p.value === null || p.value === undefined || p.value === '') ? '' : Number(p.value).toLocaleString('es-CO')"
    if not editable:
        c["cellClass"] = "bg-gray-50"
    c.update(kw)
    return c


def grilla(filas: list[dict], columnas: list[dict], guardar_filas: Callable[[list[dict]], None], *, alto=420,
           nueva: Optional[Callable[[], dict]] = None, pegar=True):
    """Tabla editable. Cada cambio llama guardar_filas(filas). Devuelve el elemento."""
    datos = [dict(f, _i=i) for i, f in enumerate(filas)]

    def emitir():
        try:
            guardar_filas([dict(d) for d in datos])
        except ValidationError as ex:
            ui.notify(f"Valor no válido: {ex.errors()[0]['loc']} – {ex.errors()[0]['msg']}", type="negative")
        except Exception as ex:  # no perder el error en silencio
            traceback.print_exc()
            ui.notify(f"No se pudo guardar el cambio: {ex}", type="negative")

    def refrescar():
        for i, d in enumerate(datos):
            d["_i"] = i
        g.options["rowData"] = datos
        g.update()

    def cambio(e: events.GenericEventArguments):
        d = e.args.get("data") or {}
        i = d.get("_i")
        if i is None or i >= len(datos):
            return
        datos[i].update(d)
        emitir()

    with ui.row().classes("gap-2 items-center"):
        if nueva:
            ui.button("Agregar fila", icon="add", on_click=lambda: (datos.append(dict(nueva())), refrescar(), emitir())).props("flat dense")

        async def borrar():
            fila = await g.get_selected_row()
            if fila is None:
                ui.notify("Seleccione una fila (clic en ella) y vuelva a intentar.")
                return
            datos.pop(fila["_i"])
            refrescar()
            emitir()
        ui.button("Borrar fila seleccionada", icon="delete", on_click=borrar).props("flat dense color=negative")
        if pegar:
            editables = [c["field"] for c in columnas if c.get("editable")]

            def abrir_pegar():
                with ui.dialog() as dlg, ui.card().classes("w-[900px] max-w-full"):
                    ui.label("Pegar filas desde Excel").classes("text-lg font-bold")
                    ui.label("Copie en Excel las columnas en este orden y péguelas aquí (Ctrl+V):").classes("text-sm")
                    ui.label(" · ".join(c["headerName"] for c in columnas if c.get("editable"))).classes("text-xs text-gray-600")
                    ta = ui.textarea().props("outlined autogrow").classes("w-full font-mono text-xs")

                    def aplicar():
                        n = 0
                        for linea in (ta.value or "").splitlines():
                            if not linea.strip():
                                continue
                            celdas = linea.split("\t")
                            fila = dict(nueva()) if nueva else {}
                            for campo, v in zip(editables, celdas):
                                if v.strip() != "":
                                    fila[campo] = numero(v)
                            datos.append(fila)
                            n += 1
                        refrescar()
                        emitir()
                        dlg.close()
                        ui.notify(f"{n} filas agregadas")
                    with ui.row():
                        ui.button("Agregar filas", on_click=aplicar)
                        ui.button("Cancelar", on_click=dlg.close).props("flat")
                dlg.open()
            ui.button("Pegar desde Excel", icon="content_paste", on_click=abrir_pegar).props("flat dense")

    g = ui.aggrid({
        "columnDefs": columnas,
        "rowData": datos,
        "defaultColDef": {"resizable": True, "sortable": False, "wrapHeaderText": True, "autoHeaderHeight": True},
        "rowSelection": {"mode": "singleRow", "checkboxes": False, "enableClickSelection": True},
        "suppressFieldDotNotation": True,
        "localeText": LOCALE,
        "stopEditingWhenCellsLoseFocus": True,
        "singleClickEdit": True,
    }, auto_size_columns=False).style(f"height:{alto}px").classes("w-full")
    g.on("cellValueChanged", cambio)
    return g


def boton_razon(titulo: str, al_guardar: Callable[[str], None], texto="No necesario…", previo=""):
    def abrir():
        with ui.dialog() as dlg, ui.card().classes("w-[560px] max-w-full"):
            ui.label(titulo).classes("font-bold")
            ui.label("Escriba por qué no se necesita. La razón queda en la revisión y en el Excel.").classes("text-sm text-gray-600")
            inp = ui.input("Razón", value=previo).classes("w-full").props("autofocus")

            def ok():
                if not (inp.value or "").strip():
                    ui.notify("Escriba una razón.", type="warning")
                    return
                al_guardar(inp.value.strip())
                dlg.close()
            inp.on("keydown.enter", ok)
            with ui.row():
                ui.button("Guardar", on_click=ok)
                ui.button("Cancelar", on_click=dlg.close).props("flat")
        dlg.open()
    return ui.button(texto, on_click=abrir).props("flat dense size=sm color=grey-8")


# ============================================================================ encabezado
@ui.refreshable
def resumen_encabezado():
    r = E.res
    with ui.row().classes("items-center gap-6"):
        ui.label(f"Directos {E.m(r.directos_cop)}").classes("text-sm")
        ui.label(f"Total {E.m(r.total_cop)}").classes("text-base font-bold")
        rojo = r.rojos()
        ui.badge(f"{rojo} por resolver" if rojo else "Sin faltantes", color="red" if rojo else "green").classes("text-sm")
        if r.amarillos():
            ui.badge(f"{r.amarillos()} por confirmar", color="orange").classes("text-sm")


def selector_proyecto():
    def opciones():
        return {str(f): f.stem.replace("_", " ") for f in sorted(PROYECTOS.glob("*.json"))}
    sel = ui.select(opciones(), value=str(E.ruta), on_change=lambda e: (E.abrir(Path(e.value)), ui.navigate.reload())) \
        .props("dense outlined dark options-dense").classes("w-80")

    def nuevo(copiar: bool):
        with ui.dialog() as dlg, ui.card():
            ui.label("Duplicar este proyecto" if copiar else "Proyecto nuevo").classes("font-bold")
            nom = ui.input("Nombre del archivo").props("autofocus")

            def crear():
                n = re.sub(r"[^\w\- ]", "", nom.value or "").strip().replace(" ", "_")
                if not n:
                    return
                ruta = PROYECTOS / f"{n}.json"
                if ruta.exists():
                    ui.notify("Ya existe un proyecto con ese nombre.", type="warning")
                    return
                p = E.p.model_copy(deep=True) if copiar else Proyecto()
                p.parametros.nombre = nom.value.strip()
                guardar(p, ruta)
                E.abrir(ruta)
                ui.navigate.reload()
            ui.button("Crear", on_click=crear)
        dlg.open()
    with ui.button(icon="more_vert").props("flat round dense color=white"):
        with ui.menu():
            ui.menu_item("Proyecto nuevo", on_click=lambda: nuevo(False))
            ui.menu_item("Duplicar este proyecto", on_click=lambda: nuevo(True))
    return sel


# ============================================================================ 1. Proyecto
def pestaña_proyecto():
    p = E.p
    par = p.parametros
    donantes = list(E.banco.categorias) or ["CBPF"]

    def set_(campo, conv=lambda x: x):
        def f(e):
            setattr(par, campo, conv(e.value) if e.value not in (None, "") else None)
            E.cambio()
        return f
    ui.label("Datos del proyecto").classes("text-lg font-bold text-teal-900")
    with ui.grid(columns=4).classes("w-full gap-3"):
        ui.input("Nombre", value=par.nombre, on_change=set_("nombre")).classes("col-span-2")
        ui.select(donantes, label="Donante (define las categorías)", value=par.donante, on_change=set_("donante"))
        ui.select(["COP", "USD", "EUR"], label="Moneda de presentación", value=par.moneda, on_change=set_("moneda"))
        ui.number("TRM (COP por USD)", value=par.trm_usd, format="%.0f", on_change=set_("trm_usd", float))
        ui.number("TRM (COP por EUR)", value=par.trm_eur, format="%.0f", on_change=set_("trm_eur", float))
        ui.number("% costos indirectos", value=round(par.pct_indirectos * 100, 2), suffix="%",
                  on_change=lambda e: (setattr(par, "pct_indirectos", (e.value or 0) / 100), E.cambio()))
        ui.number("Techo del donante (moneda de presentación)", value=par.techo, on_change=set_("techo", float))

    ui.separator().classes("my-4")
    ui.label("Actividades (marco lógico)").classes("text-lg font-bold text-teal-900")
    ui.label("Una fila por resultado, producto o actividad. A las actividades con eventos asígneles un tipo; "
             "las metas son lo que dice la propuesta y se comparan con lo presupuestado.").classes("text-sm text-gray-600")
    tipos = [""] + sorted(E.tipos())

    def guardar_acts(filas):
        p.actividades = sorted(a_modelos(Actividad, filas), key=lambda a: orden_codigo(a.codigo))
        E.cambio()
    grilla([a.model_dump() for a in p.actividades],
           [col("codigo", "Código", 80), col("nombre", "Nombre", 420), col("tipo", "Tipo de actividad", 240, lista=tipos),
            col("eventos_meta", "Eventos (propuesta)", 100, num=True), col("participantes_meta", "Participantes (propuesta)", 110, num=True),
            col("notas", "Notas", 300)],
           guardar_acts, nueva=lambda: {"codigo": "", "nombre": ""}, alto=380)


# ============================================================================ 2. Dónde y cuántos
@ui.refreshable
def pestaña_matriz():
    p = E.p
    acts = [a for a in sorted(p.actividades, key=lambda a: orden_codigo(a.codigo)) if a.tipo]
    if not acts:
        ui.label("Primero cree actividades con un tipo en la pestaña Proyecto.").classes("text-gray-600")
        return
    lugares = E.lugares()
    nombres = []
    for ev in p.eventos:
        if ev.lugar not in nombres:
            nombres.append(ev.lugar)
    nombres += [l for l in E.lugares_extra if l not in nombres]
    filas = []
    for lug in nombres:
        f = {"lugar": lug, "dep": lugares[lug].departamento if lug in lugares else "¿?"}
        for a in acts:
            n = sum(e.repeticiones for e in p.eventos if e.actividad == a.codigo and e.lugar == lug)
            f[a.codigo] = n or None
        filas.append(f)
    estado_celda = {}
    for c in E.res.celdas:
        if c.estado in ("falta", "sin_tarifa"):
            ev = p.eventos[c.evento]
            estado_celda[(ev.actividad, ev.lugar)] = True

    ui.label("Dónde y cuántas veces").classes("text-lg font-bold text-teal-900")
    ui.label("Filas = lugares, columnas = actividades. Escriba cuántas veces ocurre la actividad en ese lugar "
             "(ej. 3 fases de brigada). Haga clic en una celda para ver y ajustar quién va y qué se costea.").classes("text-sm text-gray-600")
    ui.html('<span style="background:#fde2e2;color:#b91c1c;font-weight:bold;padding:0 6px">rojo</span> = a ese evento le falta algo '
            '(vuelos, transporte, hospedaje, alimentación, salón o quién facilita).').classes("text-sm")
    columnas = [col("lugar", "Lugar", 210, editable=False, pinned="left"), col("dep", "Departamento", 135, editable=False, pinned="left")]
    for a in acts:
        n = sum(e.repeticiones for e in p.eventos if e.actividad == a.codigo)
        meta = f"{n} de {a.eventos_meta}" if a.eventos_meta is not None else f"{n} eventos"
        columnas.append(col(a.codigo, f"{a.codigo} {a.tipo} ({meta})", 150, num=True, headerTooltip=f"{a.nombre} ({a.tipo})",
                            **{":cellStyle": "params => " + json.dumps([l for (c, l) in estado_celda if c == a.codigo], ensure_ascii=False)
                               + ".includes(params.data.lugar) ? {backgroundColor: '#fde2e2', color: '#b91c1c', fontWeight: 'bold'}"
                               + " : (params.value > 0 ? {backgroundColor: '#eef7f6'} : null)"}))

    def editar(e: events.GenericEventArguments):
        cod, lug, v = e.args["colId"], e.args["data"]["lugar"], e.args.get("newValue")
        v = int(v or 0)
        evs = [x for x in p.eventos if x.actividad == cod and x.lugar == lug]
        if len(evs) > 1:
            ui.notify("Hay varios eventos en esta celda: ajústelos abajo, uno por uno.", type="warning")
            pestaña_matriz.refresh()
            return
        if not evs and v > 0:
            p.eventos.append(Evento(actividad=cod, lugar=lug, repeticiones=v))
        elif evs and v > 0:
            evs[0].repeticiones = v
        elif evs and v <= 0:
            p.eventos.remove(evs[0])
        E.sel = (cod, lug)
        E.cambio()

    def clic(e: events.GenericEventArguments):
        cod = e.args.get("colId")
        if cod in ("lugar", "dep"):
            return
        E.sel = (cod, e.args["data"]["lugar"])
        detalle_celda.refresh()

    g = ui.aggrid({"columnDefs": columnas, "rowData": filas, "stopEditingWhenCellsLoseFocus": True,
                   "suppressFieldDotNotation": True, "singleClickEdit": False,
                   "defaultColDef": {"resizable": True, "wrapHeaderText": True, "autoHeaderHeight": True},
                   "tooltipShowDelay": 200}, auto_size_columns=False).style(f"height:{min(120 + 42 * len(filas), 520)}px").classes("w-full")
    g.on("cellValueChanged", editar)
    g.on("cellClicked", clic)

    with ui.row().classes("items-center gap-2 mt-2"):
        nuevo = ui.select(sorted(lugares), label="Agregar lugar", with_input=True).props("dense outlined").classes("w-72")

        def agregar():
            if nuevo.value and nuevo.value not in nombres:
                E.lugares_extra.append(nuevo.value)
                pestaña_matriz.refresh()
        ui.button("Agregar", icon="add_location", on_click=agregar).props("flat")
        ui.label("¿No está el municipio? Agréguelo en Banco › Lugares.").classes("text-xs text-gray-500")
    ui.separator().classes("my-3")
    detalle_celda()


@ui.refreshable
def detalle_celda():
    if not E.sel:
        ui.label("Haga clic en una celda de la tabla para ver el detalle del evento.").classes("text-gray-500")
        return
    cod, lug = E.sel
    p = E.p
    act = next((a for a in p.actividades if a.codigo == cod), None)
    idx = [i for i, e in enumerate(p.eventos) if e.actividad == cod and e.lugar == lug]
    ui.label(f"{cod} · {act.nombre if act else ''} — {lug}").classes("text-lg font-bold text-teal-900")
    if not idx:
        ui.label("No hay eventos aquí. Escriba un número en la celda para crearlo.").classes("text-gray-600")
        return
    for i in idx:
        editor_evento(i)
    ui.button("Agregar otro evento en esta celda (ej. otra descripción o grupo distinto)", icon="add",
              on_click=lambda: (p.eventos.append(Evento(actividad=cod, lugar=lug, descripcion="Nuevo")), E.cambio())).props("flat")


def editor_evento(i: int):
    p = E.p
    ev = p.eventos[i]
    tipo = E.tipo_evento(ev)
    tipos = E.tipos()
    with ui.card().classes("w-full"):
        with ui.row().classes("items-end gap-3 w-full"):
            ui.input("Descripción", value=ev.descripcion,
                     on_change=lambda e: (setattr(ev, "descripcion", e.value or ""), E.cambio())).classes("w-64")
            ui.number("Repeticiones", value=ev.repeticiones, min=1, format="%d",
                      on_change=lambda e: (setattr(ev, "repeticiones", int(e.value or 1)), E.cambio())).classes("w-28")
            ui.select({"": f"(de la actividad: {tipo.nombre if tipo and not ev.tipo else '—'})", **{t: t for t in sorted(tipos)}},
                      label="Tipo", value=ev.tipo or "",
                      on_change=lambda e: (setattr(ev, "tipo", e.value or None), E.cambio(), detalle_celda.refresh())).classes("w-72")
            ui.select({"": "(la del lugar)", **{z: z for z in ["Bogotá", "Capital departamental", "Municipio", "Rural remoto"]}},
                      label="Zona de precios", value=ev.zona or "",
                      on_change=lambda e: (setattr(ev, "zona", e.value or None), E.cambio())).classes("w-52")
            ui.number("Días de salón", value=ev.dias_salon, placeholder=str(tipo.dias_salon if tipo else 0),
                      on_change=lambda e: (setattr(ev, "dias_salon", e.value), E.cambio())).classes("w-28")
            ui.space()

            def borrar():
                p.eventos.pop(i)
                E.cambio()
                detalle_celda.refresh()
            ui.button(icon="delete", on_click=borrar).props("flat round color=negative").tooltip("Borrar este evento")
        ui.input("Notas", value=ev.notas, on_change=lambda e: (setattr(ev, "notas", e.value or ""), E.cambio())).classes("w-full")

        # ---- grupos
        ui.label("¿Quién va?").classes("font-bold mt-2")
        ui.label("Los grupos vienen de la plantilla. Cambie una celda solo si este lugar es distinto (queda en naranja). "
                 "Agregue un grupo si alguien llega de otra forma (ej. 2 personas que vuelan desde Miraflores).").classes("text-xs text-gray-600")
        base = {g.nombre: g for g in (tipo.grupos if tipo else [])}
        filas = []
        for g in grupos_efectivos(ev, tipo) + [g for g in (tipo.grupos if tipo else []) if ev.ajustes.get(g.nombre, {}).get("personas") == 0]:
            d = g.model_dump()
            d["_extra"] = g.nombre not in base
            d["_aj"] = list(ev.ajustes.get(g.nombre, {}))
            filas.append(d)

        def guardar_grupos(rows):
            extras = []
            for r in rows:
                r = limpiar(r)
                nombre = r.get("nombre")
                if not nombre:
                    continue
                if nombre in base:
                    b = base[nombre].model_dump()
                    # una celda de texto vacía llega como None: equivale al texto vacío de la plantilla
                    r = {k: ("" if v is None and isinstance(b.get(k), str) else v) for k, v in r.items()}
                    aj = {k: v for k, v in r.items() if k in b and v != b[k]}
                    if aj:
                        ev.ajustes[nombre] = Grupo.model_validate({**b, **aj}).model_dump(include=set(aj))
                    else:
                        ev.ajustes.pop(nombre, None)
                else:
                    extras.append(Grupo.model_validate({k: v for k, v in r.items() if v is not None}))
            ev.grupos_extra = extras
            E.cambio()
        aj_js = "data._aj && data._aj.includes(colDef.field)"
        estilo = {"bg-orange-100": aj_js}
        cols = [
            col("nombre", "Grupo", 190, editable=False, **{":editable": "params => params.data._extra"}),
            col("rol", "Rol", 110, lista=["equipo", "participantes"], cellClassRules=estilo),
            col("personas", "Personas", 85, num=True, cellClassRules=estilo),
            col("llegada", "Cómo llega", 120, lista=list(LLEGADA), cellClassRules=estilo,
                headerTooltip="sede: ya está allí · local: transporte local · terrestre: viene de otro municipio · aereo: en avión"),
            col("dias", "Días", 70, num=True, cellClassRules=estilo, headerTooltip="Días en el lugar, contando los de viaje"),
            col("noches", "Noches", 75, num=True, cellClassRules=estilo),
            col("comidas", "Comidas", 105, lista=["completa", "refrigerio", "ninguna"], cellClassRules=estilo),
            col("por_habitacion", "Por habitación", 90, num=True, cellClassRules=estilo),
            col("dias_transporte", "Días de transporte", 95, num=True, cellClassRules=estilo,
                headerTooltip="Solo equipo: días con transporte local. Vacío = los días en el lugar"),
            col("beneficiarios", "¿Beneficiarios?", 95, lista=[True, False], cellClassRules=estilo),
            col("zona_tarifas", "Zona de sus tarifas", 130, lista=["", "Bogotá", "Capital departamental", "Municipio", "Rural remoto"],
                cellClassRules=estilo),
            col("notas", "Notas", 220),
        ]
        grilla(filas, cols, guardar_grupos, alto=90 + 42 * max(len(filas), 1),
               nueva=lambda: {"nombre": f"Grupo {len(ev.grupos_extra) + 1}", "rol": "participantes", "personas": 1,
                              "llegada": "local", "dias": 1, "noches": 0, "comidas": "completa", "por_habitacion": 1,
                              "beneficiarios": True, "_extra": True}, pegar=False)
        lista_rubros(i)


def lista_rubros(i: int):
    """El chequeo de logística del evento: cada rubro costeado, no necesario (con razón) o faltante."""
    ev = E.p.eventos[i]
    tipo = E.tipo_evento(ev)
    celdas = [c for c in E.res.celdas if c.evento == i and c.estado != "na"]
    ui.label("¿Qué se costea?").classes("font-bold mt-2")
    with ui.column().classes("gap-0 w-full"):
        for c in celdas:
            clave = f"{c.grupo}:{c.rubro}" if c.grupo else c.rubro
            with ui.row().classes("items-center gap-2 w-full border-b border-gray-100 py-0.5"):
                ui.label(ICONO[c.estado]).classes(f"w-5 {COLOR[c.estado]}")
                ui.label(f"{c.grupo or 'Evento'} · {NOMBRE_RUBRO[c.rubro]}").classes("w-80 text-sm")
                if c.estado == "ok":
                    ui.label(E.m(c.total_cop) if c.total_cop else "").classes("text-sm text-gray-700 w-40")
                    if c.rubro not in ("facilitacion", "materiales"):
                        boton_razon(f"{c.grupo or 'Evento'} · {NOMBRE_RUBRO[c.rubro]}",
                                    lambda r, k=clave: (ev.no_necesario.__setitem__(k, r), E.cambio(), detalle_celda.refresh()),
                                    texto="No se necesita…")
                elif c.estado == "no_necesario":
                    ui.label(c.razon).classes("text-sm text-gray-500 italic grow")
                    if clave in ev.no_necesario:
                        ui.button("Quitar", on_click=lambda k=clave: (ev.no_necesario.pop(k), E.cambio(), detalle_celda.refresh())) \
                            .props("flat dense size=sm")
                    else:
                        ui.label("(definido en la plantilla)").classes("text-xs text-gray-400")
                else:
                    ui.label("FALTA" if c.estado == "falta" else "SIN TARIFA").classes("text-sm text-red-700 font-bold w-40")
                    if c.estado == "falta":
                        boton_razon(f"{c.grupo or 'Evento'} · {NOMBRE_RUBRO[c.rubro]}",
                                    lambda r, k=clave: (ev.no_necesario.__setitem__(k, r), E.cambio(), detalle_celda.refresh()))
                        if c.rubro == "facilitacion":
                            def agregar_equipo():
                                dias = max([g.dias for g in grupos_efectivos(ev, tipo)] + [1]) + 1
                                ev.grupos_extra.append(Grupo(nombre="Facilitadores CIREC", rol="equipo", personas=2, llegada="aereo",
                                                             dias=dias, noches=dias - 1))
                                E.cambio()
                                detalle_celda.refresh()
                            ui.button("Agregar quién facilita", icon="person_add", on_click=agregar_equipo).props("flat dense size=sm")
                            ui.label("(2 personas en avión; ajústelo en la tabla)").classes("text-xs text-gray-500")
                        else:
                            ui.label("o ajuste el grupo arriba (noches, comidas, días de transporte)").classes("text-xs text-gray-500")
    avisos = [a for a in E.res.avisos if a.evento == i and not a.clave]
    for a in avisos:
        ui.label(("⚠ " if a.nivel == "amarillo" else "✗ ") + a.texto).classes("text-sm " + ("text-orange-700" if a.nivel == "amarillo" else "text-red-700"))


# ============================================================================ 3. Tipos de actividad
@ui.refreshable
def pestaña_tipos():
    tipos = E.tipos()
    ui.label("Tipos de actividad (plantillas)").classes("text-lg font-bold text-teal-900")
    ui.label("El perfil típico de cada tipo de evento: quién va, cómo llega, cuántos días. Se define una vez y lo heredan "
             "todos los eventos. Los tipos del banco sirven para todos los proyectos; si edita uno aquí, el cambio queda "
             "solo en este proyecto.").classes("text-sm text-gray-600")
    if not tipos:
        return
    if not getattr(pestaña_tipos, "actual", None) or pestaña_tipos.actual not in tipos:
        pestaña_tipos.actual = sorted(tipos)[0]

    def elegir(v):
        pestaña_tipos.actual = v
        pestaña_tipos.refresh()
    with ui.row().classes("items-end gap-3"):
        ui.select(sorted(tipos), value=pestaña_tipos.actual, label="Tipo", on_change=lambda e: elegir(e.value)).classes("w-96")

        def nuevo():
            n = f"Tipo nuevo {len(E.p.tipos) + 1}"
            E.p.tipos.append(TipoActividad(nombre=n))
            E.cambio()
            elegir(n)
        ui.button("Nuevo tipo", icon="add", on_click=nuevo).props("flat")
    t = tipos[pestaña_tipos.actual]
    propio = any(x.nombre == t.nombre for x in E.p.tipos)
    ui.label("Tipo propio de este proyecto" if propio else "Tipo del banco: al editarlo se crea una copia en este proyecto") \
        .classes("text-xs " + ("text-teal-700" if propio else "text-orange-700"))

    def tipo_editable() -> TipoActividad:
        for x in E.p.tipos:
            if x.nombre == t.nombre:
                return x
        copia = t.model_copy(deep=True)
        E.p.tipos.append(copia)
        return copia

    def set_(campo, conv=lambda x: x):
        def f(e):
            setattr(tipo_editable(), campo, conv(e.value))
            E.cambio()
        return f
    with ui.row().classes("gap-3 w-full items-end"):
        ui.input("Descripción", value=t.descripcion, on_change=set_("descripcion")).classes("w-[480px]")
        ui.number("Días de salón", value=t.dias_salon, on_change=set_("dias_salon", lambda v: v or 0)).classes("w-32")
        ui.checkbox("Repeticiones con las mismas personas", value=t.mismos_participantes, on_change=set_("mismos_participantes"))

    ui.label("Grupos").classes("font-bold mt-2")

    def guardar_grupos(filas):
        tipo_editable().grupos = a_modelos(Grupo, filas)
        E.cambio()
    grilla([g.model_dump() for g in t.grupos],
           [col("nombre", "Grupo", 180), col("rol", "Rol", 110, lista=["equipo", "participantes"]),
            col("personas", "Personas", 85, num=True), col("llegada", "Cómo llega", 110, lista=list(LLEGADA)),
            col("dias", "Días", 70, num=True), col("noches", "Noches", 75, num=True),
            col("comidas", "Comidas", 105, lista=["completa", "refrigerio", "ninguna"]),
            col("por_habitacion", "Por habitación", 90, num=True), col("dias_transporte", "Días de transporte", 95, num=True),
            col("beneficiarios", "¿Beneficiarios?", 95, lista=[True, False]),
            col("tarifa_transporte", "Tarifa transporte", 130, lista=[""] + E.ids_tarifa()),
            col("tarifa_hospedaje", "Tarifa hospedaje", 120, lista=[""] + E.ids_tarifa()),
            col("notas", "Notas", 220)],
           guardar_grupos, alto=260,
           nueva=lambda: {"nombre": "Nuevo grupo", "rol": "participantes", "personas": 1, "llegada": "local", "dias": 1,
                          "noches": 0, "comidas": "completa", "por_habitacion": 1, "beneficiarios": True})

    ui.label("Kits y materiales").classes("font-bold mt-2")

    def guardar_kits(filas):
        tipo_editable().kits = a_modelos(Kit, filas)
        E.cambio()
    grilla([k.model_dump() for k in t.kits],
           [col("tarifa", "Tarifa (kit)", 160, lista=E.ids_tarifa()), col("por", "Por", 120, lista=["participante", "evento"]),
            col("cantidad", "Cantidad", 90, num=True)],
           guardar_kits, alto=160, nueva=lambda: {"tarifa": "", "por": "participante", "cantidad": 1}, pegar=False)

    ui.label("Rubros que este tipo NO necesita (con su razón)").classes("font-bold mt-2")
    ui.label("Clave: «Grupo:rubro» (rubros: vuelo, transporte, hospedaje, alimentacion) o «salon» / «facilitacion».").classes("text-xs text-gray-600")

    def guardar_nn(filas):
        tipo_editable().no_necesario = {f["clave"]: f["razon"] for f in filas if f.get("clave") and f.get("razon")}
        E.cambio()
    grilla([{"clave": k, "razon": v} for k, v in t.no_necesario.items()],
           [col("clave", "Clave", 260), col("razon", "Razón", 520)], guardar_nn, alto=160,
           nueva=lambda: {"clave": "", "razon": ""}, pegar=False)


# ============================================================================ 4. Personal y otros costos
def pestaña_personal():
    p = E.p
    codigos = [a.codigo for a in sorted(p.actividades, key=lambda a: orden_codigo(a.codigo))]
    cargos = sorted({c.cargo for c in E.banco.cargos} | {c.cargo for c in p.cargos})
    ui.label("Personal").classes("text-lg font-bold text-teal-900")
    ui.label("Cargo × personas × meses × dedicación. El costo mensual sale del banco de cargos.").classes("text-sm text-gray-600")

    def g_personal(filas):
        p.personal = a_modelos(LineaPersonal, filas)
        E.cambio()
    grilla([x.model_dump() for x in p.personal],
           [col("actividad", "Actividad", 90, lista=codigos), col("cargo", "Cargo", 300, lista=cargos),
            col("personas", "Personas", 85, num=True), col("meses", "Meses", 80, num=True),
            col("dedicacion", "Dedicación (0–1)", 110, num=True, headerTooltip="0,15 = 15%"), col("notas", "Rol y justificación", 420)],
           g_personal, nueva=lambda: {"actividad": codigos[0] if codigos else "", "cargo": "", "personas": 1, "meses": 1, "dedicacion": 1},
           alto=330)
    ui.separator().classes("my-4")
    ui.label("Otros costos").classes("text-lg font-bold text-teal-900")
    ui.label("Lo que no depende de un evento: dispositivos, kits masivos, seguros, equipos, operación. Cantidad × factor × tarifa.").classes("text-sm text-gray-600")

    def g_otros(filas):
        p.otros = a_modelos(OtroCosto, filas)
        E.cambio()
    grilla([x.model_dump() for x in p.otros],
           [col("actividad", "Actividad", 90, lista=codigos), col("tarifa", "Tarifa", 140, lista=E.ids_tarifa()),
            col("cantidad", "Cantidad", 90, num=True), col("factor", "Factor", 80, num=True),
            col("unidad_factor", "Unidad del factor", 110), col("zona", "Zona", 130, lista=["", "Bogotá", "Capital departamental", "Municipio", "Rural remoto"]),
            col("notas", "Justificación", 420)],
           g_otros, nueva=lambda: {"actividad": codigos[0] if codigos else "", "tarifa": "", "cantidad": 1, "factor": 1, "unidad_factor": "vez"},
           alto=330)


# ============================================================================ 5. Revisión
@ui.refreshable
def pestaña_revision():
    r, p = E.res, E.p
    par = p.parametros
    with ui.row().classes("w-full gap-6 items-start"):
        with ui.card().classes("w-[520px]"):
            ui.label(f"Por categoría de {par.donante}").classes("font-bold text-teal-900")
            orden = E.banco.categorias.get(par.donante, [])
            for c in sorted(r.por_categoria, key=lambda c: (orden.index(c) if c in orden else 99, c)):
                with ui.row().classes("w-full justify-between"):
                    ui.label(c).classes("text-sm" + (" text-red-700 font-bold" if c == "SIN MAPEO" else ""))
                    ui.label(E.m(r.por_categoria[c])).classes("text-sm")
            ui.separator()
            for lab, v, b in [("Costos directos", r.directos_cop, False), (f"Indirectos ({par.pct_indirectos:.1%})", r.indirectos_cop, False),
                              ("TOTAL", r.total_cop, True)]:
                with ui.row().classes("w-full justify-between"):
                    ui.label(lab).classes("text-sm" + (" font-bold" if b else ""))
                    ui.label(E.m(v)).classes("text-sm" + (" font-bold" if b else ""))

            def descargar():
                SALIDAS.mkdir(exist_ok=True)
                ruta = exportar(E.banco, p, r, SALIDAS / f"{E.ruta.stem}.xlsx")
                ui.download(ruta)
            ui.button("Exportar a Excel", icon="download", on_click=descargar).classes("mt-2")
        with ui.card().classes("grow"):
            ui.label("Lista de chequeo del proyecto").classes("font-bold text-teal-900")
            ui.label("Lo que se suele olvidar y no depende de un evento.").classes("text-xs text-gray-600")
            for it, estado, razon in r.chequeo:
                with ui.row().classes("items-center gap-2 w-full"):
                    ui.label(ICONO["falta" if estado == "falta" else estado]).classes(
                        "w-5 " + ("text-orange-600 font-bold" if estado == "falta" else COLOR[estado]))
                    ui.label(it.nombre).classes("text-sm w-96")
                    if estado == "no_necesario":
                        ui.label(razon).classes("text-xs text-gray-500 italic grow")
                        ui.button("Quitar", on_click=lambda k=it.id: (p.chequeo.pop(k), E.cambio())).props("flat dense size=sm")
                    elif estado == "falta":
                        boton_razon(it.nombre, lambda t, k=it.id: (p.chequeo.__setitem__(k, t), E.cambio()), texto="No aplica…")

    # cuadrícula de completitud
    ui.label("Logística por evento").classes("text-lg font-bold text-teal-900 mt-4")
    ui.label("✓ costeado · — no necesario (con razón) · ✗ FALTA. Pase el ratón sobre una celda para ver el detalle; "
             "haga clic para ir al evento.").classes("text-sm text-gray-600")
    rubros = RUBROS_GRUPO + ["salon", "facilitacion", "materiales"]
    filas = []
    for i, ev in enumerate(p.eventos):
        f = {"_i": i, "act": ev.actividad, "lugar": ev.lugar, "desc": ev.descripcion, "rep": ev.repeticiones}
        tot = 0.0
        for rb in rubros:
            cs = [c for c in r.celdas if c.evento == i and c.rubro == rb and c.estado != "na"]
            tot += sum(c.total_cop for c in cs)
            if not cs:
                f[rb] = ""
                continue
            peor = "falta" if any(c.estado in ("falta", "sin_tarifa") for c in cs) else (
                "ok" if any(c.estado == "ok" for c in cs) else "no_necesario")
            f[rb] = ICONO[peor]
            f[f"_t_{rb}"] = "\n".join(f"{c.grupo or 'Evento'}: " + {"ok": E.m(c.total_cop) if c.total_cop else "sí",
                                                                   "no_necesario": f"no necesario – {c.razon}",
                                                                   "falta": "FALTA", "sin_tarifa": "SIN TARIFA"}[c.estado] for c in cs)
        f["total"] = E.m(tot)
        filas.append(f)
    cols = [col("act", "Act.", 70, editable=False), col("lugar", "Lugar", 190, editable=False),
            col("desc", "Descripción", 190, editable=False), col("rep", "Rep.", 60, editable=False)]
    for rb in rubros:
        cols.append(col(rb, NOMBRE_RUBRO[rb], 105, editable=False, tooltipField=f"_t_{rb}",
                        cellClassRules={"text-red-700 font-bold bg-red-50": "x == '✗'", "text-green-700": "x == '✓'",
                                        "text-gray-400": "x == '—'"}, cellStyle={"textAlign": "center", "fontSize": "16px"}))
    cols.append(col("total", "Total", 150, editable=False, type="rightAligned"))

    def ir(e: events.GenericEventArguments):
        ev = p.eventos[e.args["data"]["_i"]]
        E.sel = (ev.actividad, ev.lugar)
        tabs.set_value("Dónde y cuántos")
        pestaña_matriz.refresh()
    g = ui.aggrid({"columnDefs": cols, "rowData": filas, "tooltipShowDelay": 100,
                   "defaultColDef": {"resizable": True, "sortable": True, "wrapHeaderText": True, "autoHeaderHeight": True}},
                  auto_size_columns=False).style(f"height:{min(110 + 42 * len(filas), 600)}px").classes("w-full")
    g.on("cellClicked", ir)

    ui.label("Por resolver").classes("text-lg font-bold text-teal-900 mt-4")
    if not r.avisos:
        ui.label("Nada pendiente.").classes("text-green-700")
    for a in sorted(r.avisos, key=lambda a: a.nivel != "rojo"):
        if a.donde == "Lista de chequeo":
            continue
        with ui.row().classes("items-center gap-2 w-full border-b border-gray-100 py-1"):
            ui.label("✗" if a.nivel == "rojo" else "⚠").classes("w-5 " + ("text-red-700 font-bold" if a.nivel == "rojo" else "text-orange-600"))
            ui.label(a.donde).classes("text-sm w-96 text-gray-700")
            ui.label(a.texto).classes("text-sm grow")
            if a.evento is not None and a.clave:
                ev = p.eventos[a.evento]
                boton_razon(f"{a.donde} — {a.clave}", lambda t, ev=ev, k=a.clave: (ev.no_necesario.__setitem__(k, t), E.cambio()))
            if a.evento is not None:
                ev = p.eventos[a.evento]

                def ir_a(ev=ev):
                    E.sel = (ev.actividad, ev.lugar)
                    tabs.set_value("Dónde y cuántos")
                    pestaña_matriz.refresh()
                ui.button("Ir al evento", on_click=ir_a).props("flat dense size=sm")


# ============================================================================ 6. Banco
def pestaña_banco():
    b = E.banco
    ui.label("Banco de CIREC: sirve para todos los proyectos").classes("text-lg font-bold text-teal-900")
    ui.label("Cada cambio aquí afecta a todos los proyectos. Para un precio distinto solo en este proyecto, "
             "use 'Tarifas propias del proyecto'.").classes("text-sm text-gray-600")
    zonas = ["Todas", "Bogotá", "Capital departamental", "Municipio", "Rural remoto"]
    bloques = sorted(b.mapeo)
    tcols = [col("id", "ID", 120), col("concepto", "Concepto", 300), col("unidad", "Unidad", 110), col("zona", "Zona", 150, lista=zonas),
             col("valor", "Valor", 110, num=True), col("moneda", "Moneda", 80, lista=["COP", "USD", "EUR"]),
             col("bloque", "Bloque", 200, lista=[""] + bloques), col("fuente", "Fuente", 240), col("fecha", "Fecha", 100),
             col("estado", "Estado", 120, lista=["Confirmado", "Estimado", "Por confirmar"]), col("notas", "Notas", 260)]
    nueva_t = lambda: {"id": "", "zona": "Todas", "valor": 0, "moneda": "COP", "estado": "Por confirmar"}
    with ui.tabs() as t2:
        for n in ["Tarifas del banco", "Tarifas propias del proyecto", "Lugares", "Cargos", "Mapeo a donantes"]:
            ui.tab(n)
    with ui.tab_panels(t2, value="Tarifas del banco").classes("w-full"):
        with ui.tab_panel("Tarifas del banco"):
            def g_t(f):
                b.tarifas = a_modelos(Tarifa, f)
                E.cambio(guardar_p=False, guardar_b=True)
            grilla([x.model_dump() for x in b.tarifas], tcols, g_t, nueva=nueva_t, alto=520)
        with ui.tab_panel("Tarifas propias del proyecto"):
            ui.label("Reemplazan la tarifa del banco con el mismo ID y zona, solo en este proyecto (ej. precios del donante).").classes("text-sm text-gray-600")

            def g_tp(f):
                E.p.tarifas = a_modelos(Tarifa, f)
                E.cambio()
            grilla([x.model_dump() for x in E.p.tarifas], tcols, g_tp, nueva=nueva_t, alto=480)
        with ui.tab_panel("Lugares"):
            def g_l(f):
                b.lugares = a_modelos(Lugar, f)
                E.cambio(guardar_p=False, guardar_b=True)
            grilla([x.model_dump() for x in b.lugares],
                   [col("nombre", "Lugar", 240), col("departamento", "Departamento", 180), col("zona", "Zona de precios", 180, lista=zonas[1:]),
                    col("notas", "Notas", 360)], g_l, nueva=lambda: {"nombre": "", "departamento": "", "zona": "Municipio"}, alto=520)
        with ui.tab_panel("Cargos"):
            def g_c(f):
                b.cargos = a_modelos(Cargo, f)
                E.cambio(guardar_p=False, guardar_b=True)
            grilla([x.model_dump() for x in b.cargos],
                   [col("cargo", "Cargo", 300), col("costo_mensual_cop", "Costo mensual COP", 150, num=True),
                    col("fuente", "Fuente", 260), col("estado", "Estado", 120, lista=["Confirmado", "Estimado", "Por confirmar"]),
                    col("notas", "Notas", 360)], g_c, nueva=lambda: {"cargo": "", "costo_mensual_cop": 0}, alto=520)
        with ui.tab_panel("Mapeo a donantes"):
            don = list(b.categorias)

            def g_m(f):
                b.mapeo = {x["bloque"]: {d: x.get(d) for d in don if x.get(d)} for x in f if x.get("bloque")}
                E.cambio(guardar_p=False, guardar_b=True)
            grilla([{"bloque": k, **v} for k, v in b.mapeo.items()],
                   [col("bloque", "Bloque interno", 240)] + [col(d, d, 300, lista=b.categorias[d]) for d in don], g_m,
                   nueva=lambda: {"bloque": ""}, alto=560, pegar=False)


# ============================================================================ página
tabs: ui.tabs = None  # se crea en la página


def construir():
    global E
    E = Estado()

    @ui.page("/")
    def pagina():
        global tabs
        ui.colors(primary="#1F4E5A")
        ui.query("body").style("background:#f7f8f8")
        with ui.header().classes("items-center gap-4 px-4 py-2 bg-[#1F4E5A]"):
            ui.label("Presupuestador CIREC").classes("text-lg font-bold")
            selector_proyecto()
            ui.space()
            resumen_encabezado()
        with ui.tabs().classes("w-full bg-white text-teal-900") as tabs:
            for n in ["Proyecto", "Dónde y cuántos", "Tipos de actividad", "Personal y otros costos", "Revisión", "Banco"]:
                ui.tab(n)
        with ui.tab_panels(tabs, value="Dónde y cuántos" if E.p.eventos else "Proyecto").classes("w-full bg-transparent"):
            with ui.tab_panel("Proyecto"):
                pestaña_proyecto()
            with ui.tab_panel("Dónde y cuántos"):
                pestaña_matriz()
            with ui.tab_panel("Tipos de actividad"):
                pestaña_tipos()
            with ui.tab_panel("Personal y otros costos"):
                pestaña_personal()
            with ui.tab_panel("Revisión"):
                pestaña_revision()
            with ui.tab_panel("Banco"):
                pestaña_banco()

        def al_cambiar():
            resumen_encabezado.refresh()
            pestaña_revision.refresh()
        E.al_cambiar[:] = [al_cambiar]
        tabs.on_value_change(lambda e: (pestaña_matriz.refresh() if e.value == "Dónde y cuántos" else None,
                                        pestaña_tipos.refresh() if e.value == "Tipos de actividad" else None))


def main(puerto: int = 8080, abrir: bool = True):
    construir()
    ui.run(title="Presupuestador CIREC", port=puerto, reload=False, show=abrir, language="es", favicon="📋")
