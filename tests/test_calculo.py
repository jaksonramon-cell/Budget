"""Reglas del motor: qué se costea, qué se exige y cómo se resuelve una falta."""
import pytest

from presupuesto.calculo import calcular
from presupuesto.modelo import (Actividad, Banco, Evento, Grupo, ItemChequeo, Kit, Lugar, Parametros, Proyecto,
                                Tarifa, TipoActividad)


def banco(**kw):
    T = lambda i, v, z="Todas", m="COP": Tarifa(id=i, valor=v, zona=z, moneda=m, fuente="test")
    b = Banco(
        tarifas=[T("TIQ", 1000), T("TIQ", 1500, "Rural remoto"), T("TTE-STAFF", 100), T("TTE-PART-LOCAL", 20),
                 T("TTE-PART-MUNIC", 40), T("HOSP-STAFF", 300), T("HOSP-PART", 200), T("ALIM", 50), T("REFRIG", 10),
                 T("SALON", 500), Tarifa(id="KIT", valor=7, bloque="Materiales y kits"), T("DOLAR", 2, m="USD")],
        lugares=[Lugar(nombre="Pueblo", departamento="Cauca", zona="Municipio"),
                 Lugar(nombre="Lejos", departamento="Guaviare", zona="Rural remoto")],
        mapeo={}, chequeo=[],
    )
    for k, v in kw.items():
        setattr(b, k, v)
    return b


def proyecto(grupos, **ev):
    tipo = TipoActividad(nombre="T", grupos=grupos, dias_salon=ev.pop("dias_salon", 0), kits=ev.pop("kits", []),
                         mismos_participantes=ev.pop("mismos", True), no_necesario=ev.pop("nn_tipo", {}))
    return Proyecto(parametros=Parametros(moneda="COP"), actividades=[Actividad(codigo="1.1.1", nombre="A", tipo="T")],
                    eventos=[Evento(actividad="1.1.1", lugar=ev.pop("lugar", "Pueblo"), **ev)], tipos=[tipo])


def lineas(r, rubro):
    return [l for l in r.lineas if l.rubro == rubro]


def faltas(r):
    return sorted(a.clave for a in r.avisos if a.nivel == "rojo" and a.clave)


EQUIPO = Grupo(nombre="Equipo", rol="equipo", personas=2, llegada="aereo", dias=3, noches=2)


def test_equipo_aereo_genera_toda_la_cadena():
    r = calcular(banco(), proyecto([EQUIPO], repeticiones=2))
    assert lineas(r, "Tiquetes aéreos")[0].total_cop == 2 * 2 * 1000
    assert lineas(r, "Transporte terrestre")[0].cantidad == 2 * 3 * 2          # por día, días = días en el lugar
    assert lineas(r, "Hospedaje")[0].total_cop == 2 * 2 * 2 * 300
    assert lineas(r, "Alimentación")[0].total_cop == 2 * 3 * 2 * 50
    assert faltas(r) == []


def test_alimentacion_olvidada_es_roja():
    g = EQUIPO.model_copy(update={"comidas": "ninguna"})
    r = calcular(banco(), proyecto([g]))
    assert faltas(r) == ["Equipo:alimentacion"]


def test_razon_resuelve_la_falta_y_no_costea():
    g = EQUIPO.model_copy(update={"comidas": "ninguna"})
    r = calcular(banco(), proyecto([g], no_necesario={"Equipo:alimentacion": "Viáticos pagados por el aliado"}))
    assert faltas(r) == []
    celda = next(c for c in r.celdas if c.rubro == "alimentacion")
    assert (celda.estado, celda.razon) == ("no_necesario", "Viáticos pagados por el aliado")


def test_razon_apaga_un_rubro_con_cantidad():
    r = calcular(banco(), proyecto([EQUIPO], no_necesario={"Equipo:transporte": "Incluido en la tarifa de hotel"}))
    assert not lineas(r, "Transporte terrestre")
    assert faltas(r) == []


def test_razon_en_la_plantilla_aplica_a_todos_los_eventos():
    g = EQUIPO.model_copy(update={"dias_transporte": 0})
    r = calcular(banco(), proyecto([g], nn_tipo={"Equipo:transporte": "Van en vehículo propio"}))
    assert faltas(r) == []


def test_vuelo_sin_transporte_local_es_rojo():
    g = EQUIPO.model_copy(update={"dias_transporte": 0})
    assert faltas(calcular(banco(), proyecto([g]))) == ["Equipo:transporte"]


def test_viaje_de_varios_dias_sin_noches_exige_hospedaje():
    g = EQUIPO.model_copy(update={"noches": 0})
    assert faltas(calcular(banco(), proyecto([g]))) == ["Equipo:hospedaje"]


def test_viaje_de_un_dia_no_exige_hospedaje():
    g = Grupo(nombre="Equipo", rol="equipo", personas=2, llegada="terrestre", dias=1, unidad_transporte="trayecto")
    r = calcular(banco(), proyecto([g]))
    assert faltas(r) == []
    assert lineas(r, "Transporte terrestre")[0].cantidad == 2


def test_participantes_sin_equipo_ni_salon():
    p = Grupo(nombre="Asistentes", personas=10, llegada="local", dias=1)
    assert faltas(calcular(banco(), proyecto([p]))) == ["facilitacion", "salon"]
    assert faltas(calcular(banco(), proyecto([p, EQUIPO], dias_salon=1))) == []


def test_transporte_participantes_por_trayecto_y_tarifa_por_llegada():
    loc = Grupo(nombre="Locales", personas=10, llegada="local", dias=2)
    mun = Grupo(nombre="De fuera", personas=5, llegada="terrestre", dias=2, noches=1, por_habitacion=2)
    r = calcular(banco(), proyecto([loc, mun, EQUIPO], dias_salon=2, repeticiones=3))
    tte = {l.detalle: l for l in lineas(r, "Transporte terrestre")}
    assert tte["Locales"].total_cop == 10 * 3 * 20
    assert tte["De fuera"].total_cop == 5 * 3 * 40
    hosp = next(l for l in lineas(r, "Hospedaje") if l.detalle == "De fuera")
    assert hosp.cantidad == 3 * 1 * 3                    # 5 personas en dobles = 3 habitaciones
    assert lineas(r, "Salón / espacio")[0].total_cop == 2 * 3 * 500


def test_zona_con_tarifa_propia_y_respaldo_todas():
    r = calcular(banco(), proyecto([EQUIPO], lugar="Lejos"))
    assert lineas(r, "Tiquetes aéreos")[0].unitario_cop == 1500   # Rural remoto
    assert lineas(r, "Hospedaje")[0].unitario_cop == 300          # sin tarifa de zona: Todas


def test_tarifa_en_dolares_se_convierte():
    b = banco()
    p = proyecto([EQUIPO])
    p.parametros.trm_usd = 4000
    p.otros = [__import__("presupuesto.modelo", fromlist=["OtroCosto"]).OtroCosto(actividad="1.1.1", tarifa="DOLAR", cantidad=3,
                                                                                  bloque="Equipos")]
    r = calcular(b, p)
    assert next(l for l in r.lineas if l.origen == "otro").total_cop == 3 * 2 * 4000


def test_kits_y_participantes_unicos():
    p = Grupo(nombre="Asistentes", personas=10, llegada="local", dias=1)
    a = Grupo(nombre="Acompañantes", personas=10, llegada="local", dias=1, beneficiarios=False)
    kits = [Kit(tarifa="KIT")]
    r1 = calcular(banco(), proyecto([p, a, EQUIPO], dias_salon=1, repeticiones=3, kits=kits))
    assert lineas(r1, "Materiales y kits")[0].cantidad == 10              # mismas personas: kit una vez
    r2 = calcular(banco(), proyecto([p, a, EQUIPO], dias_salon=1, repeticiones=3, kits=kits, mismos=False))
    assert lineas(r2, "Materiales y kits")[0].cantidad == 30


def test_metas_declaradas():
    p = proyecto([EQUIPO], repeticiones=2)
    p.actividades[0].eventos_meta = 3
    r = calcular(banco(), p)
    assert any("dice 3 eventos; presupuestados: 2" in a.texto for a in r.avisos)


def test_ajuste_por_evento_y_grupo_extra():
    extra = Grupo(nombre="Desde lejos", personas=1, llegada="aereo", dias=2, noches=1, zona_tarifas="Rural remoto")
    r = calcular(banco(), proyecto([EQUIPO], ajustes={"Equipo": {"personas": 5}}, grupos_extra=[extra]))
    vuelos = {l.detalle: l for l in lineas(r, "Tiquetes aéreos")}
    assert vuelos["Equipo"].cantidad == 5
    assert vuelos["Desde lejos"].unitario_cop == 1500


def test_lista_de_chequeo():
    b = banco(chequeo=[ItemChequeo(id="seg", nombre="Seguros", tarifas=["KIT"])])
    r = calcular(b, proyecto([EQUIPO]))
    assert [(i.id, e) for i, e, _ in r.chequeo] == [("seg", "falta")]
    p = proyecto([EQUIPO]); p.chequeo = {"seg": "Lo paga CIREC"}
    assert [e for _, e, _ in calcular(b, p).chequeo] == ["no_necesario"]


def test_tarifa_inexistente_es_roja():
    g = EQUIPO.model_copy(update={"tarifa_vuelo": "NO-EXISTE"})
    r = calcular(banco(), proyecto([g]))
    assert any("no hay tarifa «NO-EXISTE»" in a.texto for a in r.avisos)


def test_indirectos_y_mapeo():
    b = banco(mapeo={"Tiquetes aéreos staff": {"CBPF": "E. Travel"}})
    g = Grupo(nombre="E", rol="equipo", personas=1, llegada="aereo", dias=1, comidas="ninguna",
              unidad_transporte="trayecto", tarifa_transporte="TTE-STAFF")
    p = proyecto([g], no_necesario={"E:transporte": "x", "E:alimentacion": "x"})
    p.parametros.pct_indirectos = 0.1
    r = calcular(b, p)
    assert r.por_categoria == {"E. Travel": 1000}
    assert r.total_cop == pytest.approx(1100)
