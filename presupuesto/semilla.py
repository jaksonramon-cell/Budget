"""Genera los datos iniciales: datos/banco.json y los proyectos de ejemplo.

    python -m presupuesto.semilla

Fuentes: legacy/constructor_v4.3/v3data.json (banco de tarifas y cargos de la
v3) y legacy/constructor_v4.3/proj_PRESUP1.json (presupuesto CBPF aprobado).
Una vez generado, el banco se mantiene desde la aplicación; volver a correr
este script lo SOBRESCRIBE.
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

from .modelo import (Actividad, Banco, Cargo, Evento, Grupo, ItemChequeo, Kit,
                     LineaPersonal, Lugar, OtroCosto, Parametros, Proyecto,
                     Tarifa, TipoActividad)
from .almacen import guardar

RAIZ = Path(__file__).resolve().parent.parent
LEGACY = RAIZ / "legacy" / "constructor_v4.3"
DATOS = RAIZ / "datos"

DONANTES = ["CBPF", "AECID", "Fondo Multidonante"]
CATEGORIAS = {
    "CBPF": ["A. Staff and Other Personnel Costs", "B. Supplies, Commodities, Materials", "C. Equipment",
             "D. Contractual Services", "E. Travel", "F. Transfers and Grants to Counterparts",
             "G. General Operating and Other Direct Costs"],
    "AECID": ["A.1 Personal operativo", "A.2 Personal administrativo", "A.3.1 Viajes, alojamientos y dietas",
              "A.3.2 Equipos, materiales y suministros", "A.3.3 Auditoría", "A.3.4 Servicios técnicos y profesionales",
              "A.3.5 Gastos financieros", "A.3.6 Alquiler de terrenos e inmuebles"],
    "Fondo Multidonante": ["1. Costos de personal", "2. Suministros, comodidades, materiales",
                           "3. Equipo, vehículo y muebles", "4. Servicios contractuales", "5. Viajes",
                           "6. Transferencias y grants a contrapartes",
                           "7. Costos generales de operación y otros costos directos"],
}
C, A, F = CATEGORIAS["CBPF"], CATEGORIAS["AECID"], CATEGORIAS["Fondo Multidonante"]
# bloque -> (CBPF, AECID, Fondo Multidonante). AECID y FMD: propuesta según los
# presupuestos Humanicemos y CIREC FMD; confirmar con cada donante.
MAPEO = {
    "Personal": (C[0], A[0], F[0]),
    "Personal administrativo": (C[0], A[1], F[0]),
    "Tiquetes aéreos staff": (C[4], A[2], F[4]),
    "Transporte terrestre staff": (C[4], A[2], F[4]),
    "Hospedaje staff": (C[4], A[2], F[4]),
    "Alimentación staff": (C[4], A[2], F[4]),
    "Tiquetes aéreos participantes": (C[1], A[2], F[4]),
    "Transporte participantes": (C[1], A[2], F[4]),
    "Hospedaje participantes": (C[1], A[2], F[4]),
    "Alimentación participantes": (C[1], A[2], F[4]),
    "Alquiler de salones": (C[1], A[7], F[3]),
    "Materiales y kits": (C[1], A[3], F[1]),
    "Paquetes por evento": (C[3], A[5], F[1]),
    "Dispositivos de movilidad": (C[1], A[3], F[1]),
    "Transferencias monetarias": (C[1], A[3], F[5]),
    "Seguros": (C[1], A[3], F[3]),
    "Equipos": (C[2], A[3], F[2]),
    "Gastos operativos": (C[6], A[3], F[6]),
    "Comunicación y visibilidad": (C[1], A[5], F[1]),
    "Servicios profesionales": (C[3], A[5], F[3]),
    "Auditoría": (C[6], A[4], F[6]),
    "Gastos financieros": (C[6], A[6], F[6]),
}
BLOQUE_POR_ID = {
    "HOSP-STAFF": "Hospedaje staff", "HOSP-PART": "Hospedaje participantes", "ALIM": "Alimentación participantes",
    "REFRIG": "Alimentación participantes", "TIQ": "Tiquetes aéreos participantes", "TTE-STAFF": "Transporte terrestre staff",
    "TTE-PART-LOCAL": "Transporte participantes", "TTE-PART-MUNIC": "Transporte participantes", "SALON": "Alquiler de salones",
    "KIT-PED": "Materiales y kits", "KIT-ENT": "Materiales y kits", "KIT-ERAE": "Materiales y kits", "MAT-DIF": "Materiales y kits",
    "KIT-SENAL": "Materiales y kits", "DOTAC": "Materiales y kits", "DISP-MOV": "Dispositivos de movilidad",
    "CASH-EMERG": "Transferencias monetarias", "SEG-VIDA": "Seguros", "CEL-MES": "Gastos operativos", "LAPTOP": "Equipos",
}
# Tarifas nuevas tomadas de los presupuestos AECID (Humanicemos) y FMD (CIREC).
NUEVAS = [
    Tarifa(id="ARL", concepto="ARL personal no vinculado por nómina (comparecientes, voluntarios)", unidad="persona-año",
           valor=120000, bloque="Seguros", fuente="Presupuesto FMD CIREC (ARL comparecientes)"),
    Tarifa(id="EXAM-MED", concepto="Exámenes médicos de ingreso y egreso", unidad="examen", valor=55000, bloque="Seguros",
           fuente="Presupuesto FMD CIREC"),
    Tarifa(id="AUDIT", concepto="Auditoría externa del proyecto", unidad="auditoría", valor=7500000, bloque="Auditoría",
           fuente="Presupuesto AECID Humanicemos (A.3.3.1)"),
    Tarifa(id="GASTOS-FIN", concepto="Gastos financieros y bancarios", unidad="mes", valor=438000, bloque="Gastos financieros",
           fuente="Presupuesto AECID Humanicemos (A.3.5.1)"),
    Tarifa(id="PAPELERIA", concepto="Papelería, impresión y comunicaciones", unidad="mes", valor=700000, bloque="Gastos operativos",
           fuente="Presupuesto AECID Humanicemos (A.3.2.16)"),
    Tarifa(id="EVENTO-CIERRE", concepto="Evento de cierre y socialización de resultados", unidad="evento", valor=10000000,
           bloque="Servicios profesionales", fuente="Presupuesto AECID Humanicemos (A.3.2.7)"),
    Tarifa(id="CUNA-RADIAL", concepto="Cuñas radiales", unidad="mes", valor=1000000, bloque="Comunicación y visibilidad",
           fuente="Presupuesto AECID Humanicemos (A.3.4.5)"),
    Tarifa(id="PIEZAS-VISIB", concepto="Piezas de visibilidad (brochures, agendas, rotafolios, pendones)", unidad="global",
           valor=15000000, bloque="Comunicación y visibilidad", fuente="Presupuesto AECID Humanicemos (A.3.4.4)"),
    Tarifa(id="LINEA-BASE", concepto="Materiales para línea de base / evaluación final", unidad="global", valor=200000,
           bloque="Materiales y kits", fuente="Presupuesto FMD CIREC (ACT 4.1)"),
    Tarifa(id="TALLER-ERAE", concepto="Taller comunitario ERAE / MCMT (metodología CIREC, todo incluido)", unidad="taller",
           valor=500000, bloque="Paquetes por evento", fuente="Presupuesto FMD CIREC (ACT 1.3)"),
]

CAPITALES = [("Leticia", "Amazonas"), ("Medellín", "Antioquia"), ("Arauca", "Arauca"), ("Barranquilla", "Atlántico"),
             ("Cartagena", "Bolívar"), ("Tunja", "Boyacá"), ("Manizales", "Caldas"), ("Florencia", "Caquetá"),
             ("Yopal", "Casanare"), ("Popayán", "Cauca"), ("Valledupar", "Cesar"), ("Quibdó", "Chocó"),
             ("Montería", "Córdoba"), ("Inírida", "Guainía"), ("San José del Guaviare", "Guaviare"), ("Neiva", "Huila"),
             ("Riohacha", "La Guajira"), ("Santa Marta", "Magdalena"), ("Villavicencio", "Meta"), ("Pasto", "Nariño"),
             ("Cúcuta", "Norte de Santander"), ("Mocoa", "Putumayo"), ("Armenia", "Quindío"), ("Pereira", "Risaralda"),
             ("San Andrés", "San Andrés"), ("Bucaramanga", "Santander"), ("Sincelejo", "Sucre"), ("Ibagué", "Tolima"),
             ("Cali", "Valle del Cauca"), ("Mitú", "Vaupés"), ("Puerto Carreño", "Vichada")]
# Municipios de la Planeación general interna y de los presupuestos FMD.
MUNICIPIOS = [("San Vicente del Caguán", "Caquetá"), ("La Montañita", "Caquetá"), ("Cáceres", "Antioquia"),
              ("Tarazá", "Antioquia"), ("Segovia", "Antioquia"), ("Remedios", "Antioquia"), ("Anorí", "Antioquia"),
              ("Valdivia", "Antioquia"), ("Sucre (Cauca)", "Cauca"), ("La Vega", "Cauca"), ("Inzá", "Cauca"),
              ("Morales", "Cauca"), ("El Tambo", "Cauca"), ("Guapi", "Cauca")]


def banco() -> Banco:
    v3 = json.loads((LEGACY / "v3data.json").read_text(encoding="utf-8"))
    tarifas = []
    for t in v3["tblTarifas"]["rows"]:
        bloque = BLOQUE_POR_ID.get(t["ID tarifa"], "Dispositivos de movilidad" if t.get("Categoría") == "Prótesis y órtesis" else None)
        tarifas.append(Tarifa(id=t["ID tarifa"], concepto=t.get("Concepto") or "", unidad=t.get("Unidad") or "",
                              zona=t.get("Zona") or "Todas", valor=t["Valor"], moneda=t.get("Moneda") or "COP", bloque=bloque,
                              fuente=t.get("Fuente") or "", fecha=t.get("Fecha"), estado=t.get("Estado") or "Por confirmar",
                              notas=t.get("Notas") or ""))
    tarifas += NUEVAS
    cargos = [Cargo(cargo=c["Cargo"], costo_mensual_cop=c["Costo mensual COP"], fuente=c.get("Fuente") or "",
                    estado=c.get("Estado") or "Por confirmar", notas=c.get("Notas") or "") for c in v3["tblCargos"]["rows"]]

    lugares = [Lugar(nombre="Bogotá", departamento="Bogotá D.C.", zona="Bogotá", notas="Sede CIREC")]
    lugares += [Lugar(nombre=m, departamento=d, zona="Capital departamental") for m, d in CAPITALES]
    lugares += [Lugar(nombre=m, departamento=d, zona="Municipio") for m, d in MUNICIPIOS]
    lugares += [Lugar(nombre="Miraflores (Guaviare)", departamento="Guaviare", zona="Rural remoto",
                      notas="Acceso aéreo desde San José del Guaviare")]
    lugares += [Lugar(nombre=f"Por definir ({d})", departamento=d, zona="Municipio", notas="Comodín mientras se define el municipio")
                for d in ["Antioquia", "Arauca", "Caquetá", "Cauca", "Chocó", "Guaviare", "Meta", "Nariño", "Norte de Santander", "Putumayo"]]

    tipos = [
        TipoActividad(nombre="Brigada de rehabilitación (fase)",
                      descripcion="Una fase de brigada en un lugar. Repeticiones = número de fases (normalmente 3).",
                      grupos=[Grupo(nombre="Equipo de brigada", rol="equipo", personas=7, llegada="aereo", dias=4, noches=3, dias_transporte=3),
                              Grupo(nombre="Beneficiarios", personas=12, llegada="terrestre", dias=2, noches=1, por_habitacion=2,
                                    tarifa_transporte="TTE-PART-LOCAL", notas="Vienen de zona rural del municipio"),
                              Grupo(nombre="Acompañantes", personas=12, llegada="terrestre", dias=2, noches=1, por_habitacion=2,
                                    tarifa_transporte="TTE-PART-LOCAL", beneficiarios=False)],
                      notas="Equipo: fisiatra, líder de brigada, ortoprotesista, trabajador social, fisioterapeuta, TO, audiovisual. "
                            "Dispositivos de movilidad van en Otros costos."),
        TipoActividad(nombre="Taller ENTRÉNATE",
                      descripcion="Taller de 3 días con cuidadores y líderes en cabecera municipal.",
                      grupos=[Grupo(nombre="Facilitadores", rol="equipo", personas=2, llegada="aereo", dias=4, noches=3),
                              Grupo(nombre="Asistentes", personas=18, llegada="terrestre", dias=3, noches=2,
                                    tarifa_transporte="TTE-PART-LOCAL")],
                      dias_salon=2, kits=[Kit(tarifa="KIT-ENT")]),
        TipoActividad(nombre="Taller de arranque (equipo)", descripcion="Inducción del equipo implementador en Bogotá.",
                      grupos=[Grupo(nombre="Equipo que viaja", rol="equipo", personas=12, llegada="aereo", dias=5, noches=4, dias_transporte=4)],
                      dias_salon=4),
        TipoActividad(nombre="Curso ERAE Básico", descripcion="Curso de 4 días para líderes comunitarios.",
                      grupos=[Grupo(nombre="Facilitadores locales", rol="equipo", personas=2, llegada="sede", dias=4),
                              Grupo(nombre="Participantes", personas=10, llegada="terrestre", dias=4, noches=3)],
                      dias_salon=4, kits=[Kit(tarifa="KIT-PED")]),
        TipoActividad(nombre="Visita de seguimiento", descripcion="1 profesional viaja a territorio (coordinación, monitoreo, género).",
                      grupos=[Grupo(nombre="Profesional", rol="equipo", personas=1, llegada="aereo", dias=3, noches=2)]),
        TipoActividad(nombre="Taller comunitario ERAE / MCMT",
                      descripcion="Taller de 1 día en la comunidad. Facilitadora pagada en Personal. Repeticiones = número de talleres.",
                      grupos=[Grupo(nombre="Facilitadora", rol="equipo", personas=1, llegada="local", dias=1, comidas="refrigerio",
                                    tarifa_transporte="TTE-PART-LOCAL", unidad_transporte="trayecto"),
                              Grupo(nombre="Asistentes", personas=10, llegada="local", dias=1, comidas="refrigerio")],
                      mismos_participantes=False,
                      no_necesario={"salon": "Espacio comunitario sin costo (escuela, salón comunal) – confirmar",
                                    "Asistentes:transporte": "Asistentes de la misma vereda o barrio – confirmar"}),
        TipoActividad(nombre="Atención en sede Bogotá (fondo agotable)",
                      descripcion="Beneficiarios que viajan a la sede CIREC en Bogotá. Repeticiones = fases.",
                      grupos=[Grupo(nombre="Beneficiarios", personas=15, llegada="aereo", dias=3, noches=3)],
                      no_necesario={"salon": "Se atiende en la sede CIREC",
                                    "facilitacion": "Personal de la sede CIREC, pagado en Personal"}),
        TipoActividad(nombre="Marcación de escuelas",
                      descripcion="Señalización preventiva de una escuela. Repeticiones = número de escuelas.",
                      grupos=[Grupo(nombre="Equipo de marcación", rol="equipo", personas=2, llegada="terrestre", dias=1,
                                    unidad_transporte="trayecto")],
                      mismos_participantes=False, kits=[Kit(tarifa="KIT-SENAL", por="evento")]),
    ]
    chequeo = [
        ItemChequeo(id="seguros", nombre="Seguros de vida del personal de campo", tarifas=["SEG-VIDA"],
                    pista="Todo el que va a terreno necesita póliza."),
        ItemChequeo(id="arl", nombre="ARL y exámenes médicos (personal no vinculado por nómina)", tarifas=["ARL", "EXAM-MED"],
                    pista="Comparecientes, voluntarios o contratistas que participan en terreno."),
        ItemChequeo(id="dotacion", nombre="Dotación del equipo de terreno (chaleco, gorra, carné)", tarifas=["DOTAC"]),
        ItemChequeo(id="celulares", nombre="Telefonía celular del equipo de terreno", tarifas=["CEL-MES"]),
        ItemChequeo(id="equipos", nombre="Equipos (computadores, cámaras)", tarifas=["LAPTOP"]),
        ItemChequeo(id="coordinacion", nombre="Dirección y coordinación (gerencia, jefaturas, % de tiempo)",
                    cargos=["Gerente de cooperación e impacto social", "Jefe de cooperación internacional", "Jefe ERAE y AIV",
                            "Coordinador(a) departamental"]),
        ItemChequeo(id="administrativo", nombre="Personal administrativo y financiero",
                    cargos=["Analista financiero", "Auxiliar administrativo"]),
        ItemChequeo(id="monitoreo", nombre="Monitoreo y seguimiento (profesional y visitas a terreno)",
                    cargos=["Profesional de monitoreo y seguimiento"], tipos_actividad=["Visita de seguimiento"]),
        ItemChequeo(id="genero", nombre="Enfoque de género (profesional o consultora)", cargos=["Jefe de género y políticas públicas"]),
        ItemChequeo(id="visibilidad", nombre="Visibilidad y comunicaciones", tarifas=["CUNA-RADIAL", "PIEZAS-VISIB", "MAT-DIF"],
                    cargos=["Coordinador(a) de comunicaciones", "Profesional II audiovisual y multimedia", "Profesional III audiovisual"]),
        ItemChequeo(id="linea_base", nombre="Línea de base y evaluación final", tarifas=["LINEA-BASE"]),
        ItemChequeo(id="cierre", nombre="Evento de cierre y socialización", tarifas=["EVENTO-CIERRE"]),
        ItemChequeo(id="auditoria", nombre="Auditoría externa", tarifas=["AUDIT"], pista="Muchos donantes la exigen."),
        ItemChequeo(id="financieros", nombre="Gastos financieros y bancarios", tarifas=["GASTOS-FIN"]),
        ItemChequeo(id="papeleria", nombre="Papelería e impresión", tarifas=["PAPELERIA"]),
    ]
    return Banco(tarifas=tarifas, lugares=lugares, cargos=cargos, tipos=tipos,
                 mapeo={b: dict(zip(DONANTES, v)) for b, v in MAPEO.items()}, categorias=CATEGORIAS, chequeo=chequeo)


# ============================================================================ PRESUP1 (CBPF aprobado)
def presup1() -> Proyecto:
    """Presupuesto CBPF aprobado (USD 919.803,40) reconstruido en el modelo de grupos."""
    leg = json.loads((LEGACY / "proj_PRESUP1.json").read_text(encoding="utf-8"))
    SRC = "Presupuesto aprobado CBPF (PRESUP1)"
    tarifas = [Tarifa(id=t["ID tarifa"], zona=t["Zona"], valor=t["Valor"], moneda="USD", fuente=SRC, estado="Confirmado",
                      notas=t.get("Notas") or "") for t in leg["tarifas"]]
    cargos = [Cargo(cargo=c["Cargo"], costo_mensual_cop=c["Costo mensual COP"], fuente=SRC, estado="Confirmado", notas=c["Notas"])
              for c in leg["cargos"]]
    personal = [LineaPersonal(actividad=p["Código WBS"], cargo=p["Cargo"], personas=p["Número de personas"], meses=p["Meses"],
                              dedicacion=p["Dedicación"], notas=p["Rol y justificación"]) for p in leg["personal"]]
    # Los tiquetes de beneficiarios (B6, B7, B15) ahora son grupos dentro de sus eventos.
    otros = [OtroCosto(actividad=o["Código WBS"], tarifa=o["ID tarifa"], cantidad=o["Cantidad"], factor=o["Factor"],
                       unidad_factor=o["Unidad del factor"], zona=o["Zona"], notas=o["Justificación"])
             for o in leg["otros"] if not o["Justificación"].startswith(("[B6]", "[B7]", "[B15]"))]
    ver = "Presupuesto aprobado: solo tiquetes; incluidos en el grupo principal – verificar"
    tipos = [
        TipoActividad(nombre="Brigada de rehabilitación (fase)", descripcion="Como en el presupuesto aprobado (E1–E4, B3–B5).",
                      grupos=[Grupo(nombre="Equipo de brigada", rol="equipo", personas=7, llegada="aereo", dias=3, noches=3),
                              Grupo(nombre="Beneficiarios", personas=12, llegada="terrestre", dias=2, noches=1, por_habitacion=2,
                                    tarifa_transporte="TTE-PART-LOCAL"),
                              Grupo(nombre="Acompañantes", personas=12, llegada="terrestre", dias=2, noches=1, por_habitacion=2,
                                    tarifa_transporte="TTE-PART-LOCAL", beneficiarios=False)]),
        TipoActividad(nombre="Atención en sede Bogotá (fondo agotable)", descripcion="B7–B9.",
                      grupos=[Grupo(nombre="Beneficiarios", personas=15, llegada="aereo", dias=3, noches=3)],
                      no_necesario={"salon": "Se atiende en la sede CIREC",
                                    "facilitacion": "Personal de la sede CIREC, pagado en Personal",
                                    "Beneficiarios:transporte": "La tarifa de hospedaje en Bogotá (USD 66) incluye transporte aeropuerto e interno"}),
        TipoActividad(nombre="Taller ENTRÉNATE", descripcion="B10–B15. El aprobado no incluye personal que viaje a facilitar.",
                      grupos=[Grupo(nombre="Asistentes", personas=18, llegada="terrestre", dias=3, noches=2, tarifa_transporte="TTE-PART-LOCAL")],
                      dias_salon=2, kits=[Kit(tarifa="KIT-ENT")]),
        TipoActividad(nombre="Taller de arranque (equipo)", descripcion="E5–E8, B23.",
                      grupos=[Grupo(nombre="Coordinadores y facilitadoras", rol="equipo", personas=12, llegada="aereo", dias=5, noches=5,
                                    dias_transporte=4)],
                      dias_salon=4),
        TipoActividad(nombre="Curso ERAE Básico", descripcion="B24–B28.",
                      grupos=[Grupo(nombre="Facilitadores locales", rol="equipo", personas=2, llegada="sede", dias=4),
                              Grupo(nombre="Participantes", personas=10, llegada="terrestre", dias=4, noches=4)],
                      dias_salon=4, kits=[Kit(tarifa="KIT-PED")]),
        TipoActividad(nombre="Visita de seguimiento", descripcion="E12–E23.",
                      grupos=[Grupo(nombre="Profesional", rol="equipo", personas=1, llegada="aereo", dias=3, noches=3)]),
        TipoActividad(nombre="Visita facilitadora a Miraflores", descripcion="E9–E11. Facilitadora AICMA de San José a Miraflores.",
                      grupos=[Grupo(nombre="Facilitadora AICMA", rol="equipo", personas=1, llegada="aereo", dias=7, noches=7, dias_transporte=0)]),
    ]
    W = {w["Código WBS"]: w for w in leg["wbs"]}
    def act(c, tipo=None, ev=None, part=None):
        w = W[c]
        return Actividad(codigo=c, nombre=w["Nombre"], tipo=tipo, eventos_meta=ev, participantes_meta=part, notas=w.get("Notas") or "")
    actividades = [act("0"), act("0.1"), act("0.2", "Visita de seguimiento", 22), act("0.3"), act("1"), act("1.2"),
                   act("1.2.1", "Brigada de rehabilitación (fase)", 15, 63), act("1.2.2", "Taller ENTRÉNATE", 4, 72), act("1.2.3"),
                   act("1.3"), act("1.3.1", "Taller de arranque (equipo)", 4), act("1.3.2", "Curso ERAE Básico", 1, 10)]
    CAP = "Capital departamental"
    mira = lambda rep: Grupo(nombre="Beneficiarios de Miraflores", personas=2, llegada="aereo", dias=2, zona_tarifas="Rural remoto",
                             beneficiarios=False, notas="Tiquete Miraflores ↔ San José")
    nn_mira = {f"Beneficiarios de Miraflores:{r}": ver for r in ("transporte", "hospedaje", "alimentacion")}
    eventos = [
        Evento(actividad="1.2.1", lugar="San José del Guaviare", repeticiones=3, descripcion="Brigada Guaviare",
               grupos_extra=[mira(3)], no_necesario=nn_mira, notas="B6: 2 beneficiarios de Miraflores vuelan a la brigada."),
        Evento(actividad="1.2.1", lugar="Por definir (Caquetá)", repeticiones=3, descripcion="Brigada Caquetá", zona=CAP),
        Evento(actividad="1.2.1", lugar="Por definir (Cauca)", repeticiones=3, descripcion="Brigada Cauca", zona=CAP),
        Evento(actividad="1.2.1", lugar="Por definir (Antioquia)", repeticiones=3, descripcion="Brigada Antioquia", zona=CAP),
        Evento(actividad="1.2.1", lugar="Bogotá", repeticiones=3, descripcion="Fondo agotable", tipo="Atención en sede Bogotá (fondo agotable)"),
    ]
    for dep, n in [("Guaviare", 12), ("Caquetá", 15), ("Cauca", 22), ("Antioquia", 23)]:
        e = Evento(actividad="1.2.2", lugar=f"Por definir ({dep})", descripcion=f"ENTRÉNATE {dep}", ajustes={"Asistentes": {"personas": n}})
        if dep == "Guaviare":
            e.grupos_extra = [Grupo(nombre="Beneficiarios de Miraflores", personas=2, llegada="aereo", dias=3, zona_tarifas="Rural remoto",
                                    beneficiarios=False)]
            e.no_necesario = dict(nn_mira)
            e.notas = "B15: 2 asistentes vuelan desde Miraflores."
        eventos.append(e)
    eventos += [
        Evento(actividad="1.3.1", lugar="Bogotá", descripcion="Taller de arranque",
               grupos_extra=[Grupo(nombre="Jefe ERAE y AIV", rol="equipo", personas=1, llegada="aereo", dias=5, noches=5, dias_transporte=0)],
               notas="El aprobado no da transporte local en Bogotá al jefe ERAE (E6 = 12 personas)."),
        Evento(actividad="1.3.1", lugar="Miraflores (Guaviare)", repeticiones=3, descripcion="Visitas a Miraflores",
               tipo="Visita facilitadora a Miraflores",
               notas="Aprobado: 21 noches pero 14 días de alimentación. El modelo da 21 días → +USD 193,13 (E11)."),
        Evento(actividad="1.3.2", lugar="Popayán", descripcion="Curso ERAE Básico"),
    ]
    for dep, n in [("Guaviare", 2), ("Antioquia", 2), ("Caquetá", 3), ("Cauca", 3)]:
        eventos.append(Evento(actividad="0.2", lugar=f"Por definir ({dep})", repeticiones=n, descripcion=f"Jefe AIV – {dep}",
                              ajustes={"Profesional": {"dias": 4, "dias_transporte": 3}}))
    for dep in ["Guaviare", "Antioquia", "Caquetá", "Cauca"]:
        eventos.append(Evento(actividad="0.2", lugar=f"Por definir ({dep})", repeticiones=2, descripcion=f"Monitoreo – {dep}"))
    for dep in ["Guaviare", "Antioquia", "Caquetá", "Cauca"]:
        eventos.append(Evento(actividad="0.2", lugar=f"Por definir ({dep})", descripcion=f"Género – {dep}"))
    return Proyecto(
        parametros=Parametros(nombre="CBPF – ERAE, AIV, MCMT y ENTRÉNATE (presupuesto APROBADO, reconstruido)", donante="CBPF",
                              moneda="USD", trm_usd=2900, pct_indirectos=0.07),
        actividades=actividades, eventos=eventos, personal=personal, otros=otros, tarifas=tarifas, tipos=tipos, cargos=cargos,
        mapeo={"Seguros": {"CBPF": "A. Staff and Other Personnel Costs"}},
        chequeo={"arl": "Todo el personal está vinculado por nómina", "linea_base": "Incluida en el trabajo del profesional de monitoreo",
                 "cierre": "No incluido en el aprobado", "auditoria": "CBPF la cubre con sus propios auditores",
                 "financieros": "No elegible / no incluido", "papeleria": "Cubierto con los costos indirectos (PSC)"},
        notas="Reconstrucción del presupuesto CBPF aprobado. Total aprobado USD 919.803,40. Diferencia conocida: E11 (+USD 193,13 directos).")


# ============================================================================ FMD: brechas reales
def fmd_brechas() -> Proyecto:
    """Tres actividades del presupuesto FMD de CIREC tal como se presupuestaron, para ver qué detecta la revisión."""
    F = "Presupuesto FMD CIREC (CIREC_~2.XLS)"
    T = lambda i, v, z="Todas", n="": Tarifa(id=i, zona=z, valor=v, fuente=F, estado="Confirmado", notas=n)
    tarifas = [T("TIQ", 1000000), T("HOSP-STAFF", 180000), T("HOSP-PART", 150000), T("ALIM", 80000), T("REFRIG", 10000),
               T("TTE-STAFF", 100000), T("TTE-PART-LOCAL", 150000), T("TTE-PART-MUNIC", 200000), T("SALON", 750000),
               T("KIT-ENT", 1400000), T("KIT-SENAL", 525000)]
    tipos = [
        TipoActividad(nombre="Taller ENTRÉNATE (como se presupuestó)",
                      descripcion="ACT 3.3: taller ENTRÉNATE de 3 días en Popayán, 12 asistentes.",
                      grupos=[Grupo(nombre="Asistentes de Popayán", personas=6, llegada="local", dias=3, comidas="completa"),
                              Grupo(nombre="Asistentes de Caquetá", personas=4, llegada="aereo", dias=3, noches=3),
                              Grupo(nombre="Asistentes de Guapi", personas=2, llegada="aereo", dias=3, noches=0,
                                    tarifa_transporte="TTE-PART-MUNIC", notas="Transporte desde su residencia al aeropuerto")],
                      dias_salon=2, kits=[Kit(tarifa="KIT-ENT")],
                      notas="Fisioterapeutas y profesional social están pagados por días en Personal, pero sin viaje a Popayán."),
        TipoActividad(nombre="Visita de acompañamiento (como se presupuestó)",
                      descripcion="ACT 4.2: jefe ERAE / consultora de género acompañan réplicas. Solo transporte y hospedaje.",
                      grupos=[Grupo(nombre="Acompañante CIREC", rol="equipo", personas=1, llegada="terrestre", dias=4, noches=4,
                                    comidas="ninguna")]),
        TipoActividad(nombre="Visita de seguimiento (como se presupuestó)",
                      descripcion="ACT 4.3: 9 tiquetes aéreos de seguimiento, sin nada más.",
                      grupos=[Grupo(nombre="Profesional de seguimiento", rol="equipo", personas=1, llegada="aereo", dias=3, noches=0,
                                    comidas="ninguna", dias_transporte=0)]),
        TipoActividad(nombre="Marcación de escuelas",
                      descripcion="ACT 1.6: 1 compareciente + 1 coordinador por escuela, viaje de un día.",
                      grupos=[Grupo(nombre="Equipo de marcación", rol="equipo", personas=2, llegada="terrestre", dias=1,
                                    unidad_transporte="trayecto")],
                      mismos_participantes=False, kits=[Kit(tarifa="KIT-SENAL", por="evento")]),
    ]
    actividades = [
        Actividad(codigo="1.6", nombre="Marcación de escuelas con señalización preventiva en 5 municipios", tipo="Marcación de escuelas",
                  eventos_meta=5),
        Actividad(codigo="3.3", nombre="Programa ENTRÉNATE de rehabilitación física comunitaria", tipo="Taller ENTRÉNATE (como se presupuestó)",
                  eventos_meta=1, participantes_meta=12),
        Actividad(codigo="4.2", nombre="Réplicas ERAE lideradas por mujeres", tipo="Visita de acompañamiento (como se presupuestó)"),
        Actividad(codigo="4.3", nombre="Seguimiento y evaluación de desempeño a facilitadores", tipo="Visita de seguimiento (como se presupuestó)"),
    ]
    eventos = [
        Evento(actividad="1.6", lugar="Por definir (Cauca)", repeticiones=5, descripcion="5 escuelas"),
        Evento(actividad="3.3", lugar="Popayán", descripcion="Taller ENTRÉNATE Popayán",
               notas="Presupuestado: 12 noches (= 4 de Caquetá × 3), 36 días de alimentación, 6 trayectos locales, 4+2 tiquetes."),
        Evento(actividad="4.2", lugar="Por definir (Cauca)", repeticiones=6, descripcion="Acompañamiento a réplicas",
               notas="Presupuestado: 24 noches y 25 desplazamientos; sin alimentación."),
        Evento(actividad="4.3", lugar="Por definir (Caquetá)", repeticiones=9, descripcion="Seguimiento a facilitadores",
               notas="Presupuestado: 9 tiquetes. Los 3 días por visita son un supuesto: el presupuesto no lo dice."),
    ]
    personal = [LineaPersonal(actividad="3.3", cargo="Fisioterapeuta", personas=2, meses=3, dedicacion=1,
                              notas="En el FMD: 2 fisioterapeutas × 90 días")]
    return Proyecto(parametros=Parametros(nombre="FMD CIREC – brechas detectadas (ejemplo)", donante="Fondo Multidonante",
                                          moneda="COP", pct_indirectos=0.07),
                    actividades=actividades, eventos=eventos, personal=personal, tarifas=tarifas, tipos=tipos,
                    chequeo={i.id: "Fuera del alcance de este ejemplo parcial" for i in banco().chequeo},
                    notas="Ejemplo: actividades 1.6, 3.3, 4.2 y 4.3 del presupuesto FMD tal como se presupuestaron. "
                          "Muestra qué marca la revisión. No es un presupuesto completo.")


def main():
    DATOS.mkdir(exist_ok=True)
    (DATOS / "proyectos").mkdir(exist_ok=True)
    guardar(banco(), DATOS / "banco.json")
    guardar(presup1(), DATOS / "proyectos" / "PRESUP1_CBPF_aprobado.json")
    guardar(fmd_brechas(), DATOS / "proyectos" / "FMD_brechas_ejemplo.json")
    guardar(Proyecto(), DATOS / "proyectos" / "Proyecto_nuevo.json")
    print("Datos generados en", DATOS, file=sys.stderr)


if __name__ == "__main__":
    main()
