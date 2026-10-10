"""Motor de cálculo: proyecto + banco -> líneas de costo, estado de cada rubro y avisos.

Regla de oro: un rubro que la logística exige (por ejemplo, alimentación para
quien pasa días en el lugar) queda en uno de tres estados:
  ok           costeado (cantidad > 0 y con tarifa)
  no_necesario cantidad 0 con una razón escrita
  falta        cantidad 0 sin razón  -> se marca en rojo en la revisión
Nada puede faltar en silencio.
"""
from __future__ import annotations

import math
from dataclasses import dataclass, field
from typing import Optional

from .modelo import (NOMBRE_RUBRO, Banco, Evento, Grupo, Proyecto, Tarifa,
                     TipoActividad)

BLOQUE_RUBRO = {
    ("equipo", "vuelo"): "Tiquetes aéreos staff",
    ("equipo", "transporte"): "Transporte terrestre staff",
    ("equipo", "hospedaje"): "Hospedaje staff",
    ("equipo", "alimentacion"): "Alimentación staff",
    ("participantes", "vuelo"): "Tiquetes aéreos participantes",
    ("participantes", "transporte"): "Transporte participantes",
    ("participantes", "hospedaje"): "Hospedaje participantes",
    ("participantes", "alimentacion"): "Alimentación participantes",
}
UNIDAD_RUBRO = {"vuelo": "tiquete ida y vuelta", "hospedaje": "noche", "alimentacion": "persona-día",
                "salon": "día de salón"}


@dataclass
class Linea:
    origen: str                 # evento · personal · otro
    actividad: str
    lugar: str
    detalle: str                # grupo, cargo o concepto
    rubro: str
    bloque: str
    cantidad: float
    unidad: str
    tarifa: Optional[str]
    unitario_cop: float
    total_cop: float
    calculo: str
    evento: Optional[int] = None


@dataclass
class Celda:
    """Estado de un rubro dentro de un evento (para la cuadrícula de revisión)."""
    evento: int
    grupo: Optional[str]        # None = rubro del evento (salón, facilitación, materiales)
    rubro: str
    estado: str                 # ok · no_necesario · falta · sin_tarifa · na
    razon: str = ""
    total_cop: float = 0


@dataclass
class Aviso:
    nivel: str                  # rojo · amarillo
    donde: str
    texto: str
    evento: Optional[int] = None
    clave: Optional[str] = None  # clave de no_necesario que lo resolvería


@dataclass
class Resultado:
    lineas: list[Linea] = field(default_factory=list)
    celdas: list[Celda] = field(default_factory=list)
    avisos: list[Aviso] = field(default_factory=list)
    chequeo: list[tuple] = field(default_factory=list)   # (item, estado, razón)
    por_bloque: dict[str, float] = field(default_factory=dict)
    por_categoria: dict[str, float] = field(default_factory=dict)
    directos_cop: float = 0
    indirectos_cop: float = 0
    total_cop: float = 0
    fx: float = 1               # COP por unidad de la moneda de presentación

    def rojos(self) -> int:
        return sum(1 for a in self.avisos if a.nivel == "rojo")

    def amarillos(self) -> int:
        return sum(1 for a in self.avisos if a.nivel == "amarillo")


def fmt(x: float, dec: int = 0) -> str:
    """Número con formato colombiano: 1.234.567,89"""
    s = f"{x:,.{dec}f}"
    return s.replace(",", "_").replace(".", ",").replace("_", ".")


def _n(x: float) -> str:
    return fmt(x, 0) if float(x).is_integer() else fmt(x, 2)


class Contexto:
    """Banco + proyecto ya fusionados, con búsquedas."""

    def __init__(self, banco: Banco, proyecto: Proyecto):
        self.banco, self.p = banco, proyecto
        par = proyecto.parametros
        self.trm = {"COP": 1.0, "USD": par.trm_usd, "EUR": par.trm_eur}
        self.tarifas: dict[tuple[str, str], Tarifa] = {(t.id, t.zona): t for t in banco.tarifas}
        base = {t.id: t for t in banco.tarifas}
        for t in proyecto.tarifas:
            b = base.get(t.id)
            if b:  # hereda concepto, unidad y bloque del banco si el proyecto no los da
                t = t.model_copy(update={k: getattr(b, k) for k in ("concepto", "unidad", "bloque") if not getattr(t, k)})
            self.tarifas[(t.id, t.zona)] = t
        self.tipos = {t.nombre: t for t in banco.tipos}
        self.tipos.update({t.nombre: t for t in proyecto.tipos})
        self.lugares = {l.nombre: l for l in banco.lugares}
        self.lugares.update({l.nombre: l for l in proyecto.lugares})
        self.cargos = {c.cargo: c for c in banco.cargos}
        self.cargos.update({c.cargo: c for c in proyecto.cargos})
        self.mapeo = {**banco.mapeo, **proyecto.mapeo}
        self.actividades = {a.codigo: a for a in proyecto.actividades}

    def tarifa(self, tid: Optional[str], zona: Optional[str]) -> Optional[Tarifa]:
        if not tid:
            return None
        return self.tarifas.get((tid, zona or "Todas")) or self.tarifas.get((tid, "Todas"))

    def cop(self, t: Tarifa) -> float:
        return t.valor * self.trm[t.moneda]

    def concepto(self, tid: str) -> str:
        for (i, _z), t in self.tarifas.items():
            if i == tid:
                return t.concepto or tid
        return tid

    def tipo_de(self, ev: Evento) -> Optional[TipoActividad]:
        nombre = ev.tipo or (self.actividades[ev.actividad].tipo if ev.actividad in self.actividades else None)
        return self.tipos.get(nombre) if nombre else None


# ----------------------------------------------------------------------------- grupos
def grupos_efectivos(ev: Evento, tipo: Optional[TipoActividad]) -> list[Grupo]:
    """Grupos de la plantilla con los ajustes del evento, más los grupos extra."""
    out = []
    for g in (tipo.grupos if tipo else []):
        aj = ev.ajustes.get(g.nombre)
        out.append(g.model_copy(update=aj) if aj else g)
    out += ev.grupos_extra
    return [g for g in out if g.personas > 0]


def unidad_transporte(g: Grupo) -> str:
    return g.unidad_transporte or ("dia" if g.rol == "equipo" else "trayecto")


def dias_transporte(g: Grupo) -> float:
    return g.dias if g.dias_transporte is None else g.dias_transporte


def tarifa_defecto(g: Grupo, rubro: str) -> str:
    if rubro == "vuelo":
        return g.tarifa_vuelo or "TIQ"
    if rubro == "transporte":
        if g.tarifa_transporte:
            return g.tarifa_transporte
        if g.rol == "equipo":
            return "TTE-STAFF"
        return "TTE-PART-MUNIC" if g.llegada == "terrestre" else "TTE-PART-LOCAL"
    if rubro == "hospedaje":
        return g.tarifa_hospedaje or ("HOSP-STAFF" if g.rol == "equipo" else "HOSP-PART")
    if rubro == "alimentacion":
        return g.tarifa_alimentacion or ("REFRIG" if g.comidas == "refrigerio" else "ALIM")
    raise ValueError(rubro)


def cantidad_y_regla(g: Grupo, rubro: str, r: int) -> tuple[float, bool, str, str]:
    """(cantidad, ¿lo exige la logística?, unidad, explicación de la cantidad)."""
    n = g.personas
    if rubro == "vuelo":
        q = n * r if g.llegada == "aereo" else 0
        return q, g.llegada == "aereo", UNIDAD_RUBRO["vuelo"], f"{_n(n)} personas × {r} viaje(s)"
    if rubro == "transporte":
        req = g.llegada != "sede"
        if unidad_transporte(g) == "dia":
            d = dias_transporte(g)
            return (n * d * r if req else 0), req, "persona-día", f"{_n(n)} personas × {_n(d)} días × {r}"
        return (n * r if req else 0), req, "trayecto ida y vuelta", f"{_n(n)} personas × {r} viaje(s)"
    if rubro == "hospedaje":
        req = g.llegada in ("terrestre", "aereo") and g.dias >= 2
        hab = math.ceil(n / max(g.por_habitacion, 1))
        unidad = "habitación-noche" if g.por_habitacion > 1 else "noche"
        txt = (f"{hab} habitaciones ({_n(n)} personas, {g.por_habitacion} por habitación)" if g.por_habitacion > 1
               else f"{_n(n)} personas") + f" × {_n(g.noches)} noches × {r}"
        return hab * g.noches * r, req, unidad, txt
    if rubro == "alimentacion":
        q = n * g.dias * r if g.comidas != "ninguna" else 0
        unidad = "refrigerio" if g.comidas == "refrigerio" else "persona-día"
        return q, g.dias >= 1, unidad, f"{_n(n)} personas × {_n(g.dias)} días × {r}"
    raise ValueError(rubro)


def participantes_unicos(ev: Evento, tipo: Optional[TipoActividad]) -> float:
    gs = [g for g in grupos_efectivos(ev, tipo) if g.rol == "participantes" and g.beneficiarios]
    mult = 1 if (tipo is None or tipo.mismos_participantes) else ev.repeticiones
    return sum(g.personas for g in gs) * mult


# ----------------------------------------------------------------------------- cálculo
def calcular(banco: Banco, proyecto: Proyecto) -> Resultado:
    cx = Contexto(banco, proyecto)
    res = Resultado()
    par = proyecto.parametros
    res.fx = cx.trm[par.moneda]

    for i, ev in enumerate(proyecto.eventos):
        _evento(cx, i, ev, res)
    _personal(cx, res)
    _otros(cx, res)
    _actividades(cx, res)
    _chequeo(cx, res)

    for l in res.lineas:
        res.por_bloque[l.bloque] = res.por_bloque.get(l.bloque, 0) + l.total_cop
    for b, v in res.por_bloque.items():
        cat = cx.mapeo.get(b, {}).get(par.donante) or "SIN MAPEO"
        res.por_categoria[cat] = res.por_categoria.get(cat, 0) + v
        if cat == "SIN MAPEO" and v:
            res.avisos.append(Aviso("rojo", "Mapeo", f"El bloque «{b}» no tiene categoría para {par.donante} (Banco › Mapeo)."))
    res.directos_cop = sum(res.por_bloque.values())
    res.indirectos_cop = res.directos_cop * par.pct_indirectos
    res.total_cop = res.directos_cop + res.indirectos_cop
    if par.techo and res.total_cop / res.fx > par.techo:
        res.avisos.append(Aviso("rojo", "Proyecto", f"El total supera el techo del donante en {fmt(res.total_cop / res.fx - par.techo, 2)} {par.moneda}."))
    return res


def _razon(ev: Evento, tipo: Optional[TipoActividad], clave: str) -> str:
    """Razón de 'no necesario' para una clave: primero el evento, luego la plantilla."""
    if clave in ev.no_necesario:
        return ev.no_necesario[clave]
    if tipo and clave in tipo.no_necesario:
        return tipo.no_necesario[clave]
    rubro = clave.split(":")[-1]
    if tipo and ":" in clave and f"*:{rubro}" in tipo.no_necesario:
        return tipo.no_necesario[f"*:{rubro}"]
    return ""


def _evento(cx: Contexto, i: int, ev: Evento, res: Resultado):
    act = cx.actividades.get(ev.actividad)
    donde = f"{ev.actividad} · {ev.lugar}" + (f" · {ev.descripcion}" if ev.descripcion else "")
    if act is None:
        res.avisos.append(Aviso("rojo", donde, f"La actividad {ev.actividad} no existe.", i))
    lugar = cx.lugares.get(ev.lugar)
    if lugar is None:
        res.avisos.append(Aviso("rojo", donde, f"El lugar «{ev.lugar}» no está en el banco de lugares.", i))
    tipo = cx.tipo_de(ev)
    if tipo is None:
        res.avisos.append(Aviso("rojo", donde, "El evento no tiene tipo de actividad (plantilla).", i))
    zona_ev = ev.zona or (lugar.zona if lugar else "Todas")
    r = ev.repeticiones
    if r < 1:
        res.avisos.append(Aviso("rojo", donde, "Repeticiones debe ser 1 o más.", i))
    grupos = grupos_efectivos(ev, tipo)

    def linea(detalle, rubro, bloque, q, unidad, tid, zona, txt):
        t = cx.tarifa(tid, zona)
        if t is None:
            res.avisos.append(Aviso("rojo", donde, f"{detalle}: no hay tarifa «{tid}» para la zona {zona}.", i))
            return 0.0, False
        u = cx.cop(t)
        tot = q * u
        res.lineas.append(Linea("evento", ev.actividad, ev.lugar, detalle, NOMBRE_RUBRO.get(rubro, rubro), bloque, q, unidad,
                                t.id, u, tot, f"{txt} = {_n(q)} {unidad}", i))
        return tot, True

    for g in grupos:
        zona = g.zona_tarifas or zona_ev
        for rubro in ("vuelo", "transporte", "hospedaje", "alimentacion"):
            q, req, unidad, txt = cantidad_y_regla(g, rubro, r)
            clave = f"{g.nombre}:{rubro}"
            razon = _razon(ev, tipo, clave)
            if razon and (q > 0 or req):  # marcado como no necesario: no se costea
                res.celdas.append(Celda(i, g.nombre, rubro, "no_necesario", razon))
            elif q > 0:
                tot, ok = linea(g.nombre, rubro, BLOQUE_RUBRO[(g.rol, rubro)], q, unidad, tarifa_defecto(g, rubro), zona, txt)
                res.celdas.append(Celda(i, g.nombre, rubro, "ok" if ok else "sin_tarifa", total_cop=tot))
            elif req:
                res.celdas.append(Celda(i, g.nombre, rubro, "falta"))
                res.avisos.append(Aviso("rojo", donde, f"{g.nombre} ({_n(g.personas)}): falta {NOMBRE_RUBRO[rubro].lower()}. "
                                        + _por_que(g, rubro), i, clave))
            else:
                res.celdas.append(Celda(i, g.nombre, rubro, "na"))
        # coherencia
        if g.noches > g.dias:
            res.avisos.append(Aviso("amarillo", donde, f"{g.nombre}: {_n(g.noches)} noches pero {_n(g.dias)} días. Convención: días cuentan los de viaje.", i))
        if unidad_transporte(g) == "dia" and g.llegada != "sede" and dias_transporte(g) > g.dias:
            res.avisos.append(Aviso("amarillo", donde, f"{g.nombre}: más días de transporte ({_n(dias_transporte(g))}) que días en el lugar ({_n(g.dias)}).", i))
        if g.llegada in ("sede", "local") and g.noches > 0:
            res.avisos.append(Aviso("amarillo", donde, f"{g.nombre}: viven en el lugar pero tienen {_n(g.noches)} noches de hospedaje. ¿Es correcto?", i))

    hay_part = any(g.rol == "participantes" for g in grupos)
    hay_equipo = any(g.rol == "equipo" for g in grupos)

    # salón
    ds = ev.dias_salon if ev.dias_salon is not None else (tipo.dias_salon if tipo else 0)
    razon_salon = _razon(ev, tipo, "salon")
    if razon_salon:
        res.celdas.append(Celda(i, None, "salon", "no_necesario", razon_salon))
    elif ds > 0:
        tot, ok = linea("Evento", "salon", "Alquiler de salones", ds * r, UNIDAD_RUBRO["salon"], "SALON", zona_ev, f"{_n(ds)} días × {r}")
        res.celdas.append(Celda(i, None, "salon", "ok" if ok else "sin_tarifa", total_cop=tot))
    elif hay_part:
        res.celdas.append(Celda(i, None, "salon", "falta"))
        res.avisos.append(Aviso("rojo", donde, "Hay participantes pero no hay salón o espacio. ¿Dónde se reúnen?", i, "salon"))
    else:
        res.celdas.append(Celda(i, None, "salon", "na"))

    # facilitación
    if hay_part and not hay_equipo:
        razon = _razon(ev, tipo, "facilitacion")
        res.celdas.append(Celda(i, None, "facilitacion", "no_necesario" if razon else "falta", razon))
        if not razon:
            res.avisos.append(Aviso("rojo", donde, "Hay participantes pero nadie del equipo va al evento. ¿Quién lo facilita y cómo llega?", i, "facilitacion"))
    else:
        res.celdas.append(Celda(i, None, "facilitacion", "ok" if hay_equipo else "na"))

    # materiales y kits
    kits = tipo.kits if tipo else []
    tot_k = 0.0
    for k in kits:
        if k.por == "participante":
            q = participantes_unicos(ev, tipo) * k.cantidad
            txt = f"{_n(participantes_unicos(ev, tipo))} participantes × {_n(k.cantidad)}"
        else:
            q = r * k.cantidad
            txt = f"{r} evento(s) × {_n(k.cantidad)}"
        if q > 0:
            t = cx.tarifa(k.tarifa, zona_ev)
            bloque = (t.bloque if t and t.bloque else "Materiales y kits")
            tot, _ = linea("Evento", "materiales", bloque, q, t.unidad if t and t.unidad else "unidad", k.tarifa, zona_ev,
                           f"{cx.concepto(k.tarifa)}: {txt}")
            tot_k += tot
    res.celdas.append(Celda(i, None, "materiales", "ok" if tot_k else "na", total_cop=tot_k))


def _por_que(g: Grupo, rubro: str) -> str:
    if rubro == "vuelo":
        return "Llegan en avión."
    if rubro == "transporte":
        return {"aereo": "Llegan en avión: necesitan traslado al aeropuerto y transporte local.",
                "terrestre": "Vienen de otro municipio por tierra.",
                "local": "Viven en el municipio pero deben desplazarse al evento."}[g.llegada]
    if rubro == "hospedaje":
        return f"Vienen de fuera y están {_n(g.dias)} días."
    if rubro == "alimentacion":
        return f"Están {_n(g.dias)} día(s) en el evento."
    return ""


def _personal(cx: Contexto, res: Resultado):
    for p in cx.p.personal:
        c = cx.cargos.get(p.cargo)
        donde = f"{p.actividad} · Personal · {p.cargo}"
        if c is None:
            res.avisos.append(Aviso("rojo", donde, f"El cargo «{p.cargo}» no está en el banco de cargos."))
            continue
        if p.dedicacion > 1:
            res.avisos.append(Aviso("rojo", donde, "Dedicación mayor a 100%."))
        if p.actividad not in cx.actividades:
            res.avisos.append(Aviso("rojo", donde, f"La actividad {p.actividad} no existe."))
        tot = p.personas * p.meses * p.dedicacion * c.costo_mensual_cop
        res.lineas.append(Linea("personal", p.actividad, "", p.cargo, "Personal", "Personal", p.personas * p.meses * p.dedicacion,
                                "mes-persona", None, c.costo_mensual_cop, tot,
                                f"{_n(p.personas)} × {_n(p.meses)} meses × {fmt(p.dedicacion * 100, 0)}% dedicación"))


def _otros(cx: Contexto, res: Resultado):
    for o in cx.p.otros:
        donde = f"{o.actividad} · Otros costos · {o.tarifa}"
        t = cx.tarifa(o.tarifa, o.zona)
        if t is None:
            res.avisos.append(Aviso("rojo", donde, f"No existe la tarifa «{o.tarifa}»."))
            continue
        if o.actividad not in cx.actividades:
            res.avisos.append(Aviso("rojo", donde, f"La actividad {o.actividad} no existe."))
        bloque = o.bloque or t.bloque
        if not bloque:
            res.avisos.append(Aviso("rojo", donde, f"La tarifa {t.id} no tiene bloque de costo (Banco › Tarifas)."))
            bloque = "Sin bloque"
        u = cx.cop(t)
        q = o.cantidad * o.factor
        res.lineas.append(Linea("otro", o.actividad, o.zona or "", t.concepto or t.id, t.concepto or t.id, bloque, q,
                                t.unidad or "unidad", t.id, u, q * u,
                                f"{_n(o.cantidad)} × {_n(o.factor)} {o.unidad_factor}" + (f" · {o.notas}" if o.notas else "")))


def _actividades(cx: Contexto, res: Resultado):
    for a in cx.p.actividades:
        evs = [e for e in cx.p.eventos if e.actividad == a.codigo]
        n_ev = sum(e.repeticiones for e in evs)
        n_part = sum(participantes_unicos(e, cx.tipo_de(e)) for e in evs)
        donde = f"{a.codigo} · {a.nombre}"
        if a.tipo and a.tipo not in cx.tipos:
            res.avisos.append(Aviso("rojo", donde, f"El tipo de actividad «{a.tipo}» no existe."))
        if a.tipo and not evs:
            res.avisos.append(Aviso("amarillo", donde, "Tiene tipo de actividad pero ningún lugar asignado (0 eventos)."))
        if a.eventos_meta is not None and a.eventos_meta != n_ev:
            res.avisos.append(Aviso("rojo", donde, f"La propuesta dice {a.eventos_meta} eventos; presupuestados: {n_ev}."))
        if a.participantes_meta is not None and a.participantes_meta != round(n_part):
            res.avisos.append(Aviso("rojo", donde, f"La propuesta dice {a.participantes_meta} participantes; presupuestados: {_n(n_part)}."))
        tiene_costo = evs or any(p.actividad == a.codigo for p in cx.p.personal) or any(o.actividad == a.codigo for o in cx.p.otros)
        if not tiene_costo and "." in a.codigo and a.codigo.count(".") >= 2:
            res.avisos.append(Aviso("amarillo", donde, "Actividad sin ningún costo."))


def _chequeo(cx: Contexto, res: Resultado):
    tarifas_usadas = {o.tarifa for o in cx.p.otros}
    for e in cx.p.eventos:
        t = cx.tipo_de(e)
        for k in (t.kits if t else []):
            tarifas_usadas.add(k.tarifa)
    cargos_usados = {p.cargo for p in cx.p.personal}
    tipos_usados = {t.nombre for e in cx.p.eventos if (t := cx.tipo_de(e))}
    for it in cx.banco.chequeo:
        hecho = (tarifas_usadas & set(it.tarifas)) or (cargos_usados & set(it.cargos)) or (tipos_usados & set(it.tipos_actividad))
        razon = cx.p.chequeo.get(it.id, "")
        estado = "ok" if hecho else ("no_necesario" if razon else "falta")
        res.chequeo.append((it, estado, razon))
        if estado == "falta":
            res.avisos.append(Aviso("amarillo", "Lista de chequeo", f"{it.nombre}: no está presupuestado. {it.pista}", clave=f"chequeo:{it.id}"))
