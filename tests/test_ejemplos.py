"""Los ejemplos reales: el CBPF aprobado cuadra y las brechas del FMD se detectan."""
from pathlib import Path

import pytest

from presupuesto.almacen import cargar_banco, cargar_proyecto
from presupuesto.calculo import calcular
from presupuesto import semilla

DATOS = Path(__file__).resolve().parent.parent / "datos"


@pytest.fixture(scope="module")
def banco():
    return semilla.banco()


def test_datos_versionados_iguales_a_la_semilla(banco):
    """datos/*.json debe ser lo que genera la semilla (si se cambia uno, regenerar)."""
    assert cargar_banco(DATOS / "banco.json") == banco
    assert cargar_proyecto(DATOS / "proyectos" / "PRESUP1_CBPF_aprobado.json") == semilla.presup1()


def test_presup1_cuadra_con_el_aprobado(banco):
    r = calcular(banco, semilla.presup1())
    usd = lambda x: round(x / r.fx, 2)
    # Aprobado: directos 859.629,35. Única diferencia conocida: E11 Miraflores (+193,13),
    # 21 noches de hotel pero solo 14 días de alimentación en el aprobado.
    assert usd(r.directos_cop) == pytest.approx(859629.35 + 193.13, abs=0.01)
    assert usd(r.total_cop) == pytest.approx(920010.05, abs=0.01)
    cat = {k: usd(v) for k, v in r.por_categoria.items()}
    assert cat["A. Staff and Other Personnel Costs"] == pytest.approx(285628.73, abs=0.01)
    assert cat["C. Equipment"] == pytest.approx(939.73, abs=0.01)
    assert cat["G. General Operating and Other Direct Costs"] == pytest.approx(1000.12, abs=0.01)
    # B26: la alimentación de los 2 facilitadores del curso pasa de B a E (220,72).
    assert cat["B. Supplies, Commodities, Materials"] == pytest.approx(481791.40 - 220.72, abs=0.01)
    assert cat["E. Travel"] == pytest.approx(90269.37 + 193.13 + 220.72, abs=0.01)


def test_presup1_preguntas_abiertas(banco):
    r = calcular(banco, semilla.presup1())
    rojos = sorted((a.donde.split(" · ")[0], a.clave) for a in r.avisos if a.nivel == "rojo")
    assert rojos == sorted(
        [("1.2.1", "salon")] * 4                       # brigadas sin salón
        + [("1.2.2", "facilitacion")] * 4              # ENTRÉNATE sin personal que viaje
        + [("1.3.1", "Jefe ERAE y AIV:transporte"), ("1.3.1", "Facilitadora AICMA:transporte")])
    assert r.amarillos() == 0


def test_fmd_brechas_detectadas(banco):
    r = calcular(banco, semilla.fmd_brechas())
    rojos = sorted((a.donde.split(" · ")[0], a.clave) for a in r.avisos if a.nivel == "rojo")
    assert rojos == [
        ("3.3", "Asistentes de Guapi:hospedaje"),
        ("3.3", "facilitacion"),
        ("4.2", "Acompañante CIREC:alimentacion"),
        ("4.3", "Profesional de seguimiento:alimentacion"),
        ("4.3", "Profesional de seguimiento:hospedaje"),
        ("4.3", "Profesional de seguimiento:transporte"),
    ]
    # 1.6 (marcación de escuelas, viaje de un día) está completa.
    assert not [a for a in r.avisos if a.donde.startswith("1.6")]
