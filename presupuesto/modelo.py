"""Modelo de datos del presupuestador.

Dos archivos JSON:
  * Banco (datos/banco.json): tarifas, lugares, cargos, tipos de actividad,
    mapeo a donantes y la lista de chequeo del proyecto. Se llena una vez.
  * Proyecto (datos/proyectos/*.json): actividades, eventos (dónde y cuántas
    veces), personal y otros costos. Uno por propuesta.

La idea central: un evento se describe con GRUPOS DE PERSONAS (quién va,
desde dónde, cuántos días). La logística (vuelos, transporte, hospedaje,
alimentación) se deduce de los grupos; nada se escribe a mano.
"""
from __future__ import annotations

from typing import Literal, Optional

from pydantic import BaseModel, Field

Rol = Literal["equipo", "participantes"]
# sede: ya está en el lugar (no se desplaza) · local: vive en el municipio y
# necesita transporte local · terrestre: viene de otro municipio por tierra ·
# aereo: llega en avión (más traslado al aeropuerto).
Llegada = Literal["sede", "local", "terrestre", "aereo"]
Comidas = Literal["completa", "refrigerio", "ninguna"]
Unidad = Literal["dia", "trayecto"]

# Rubros logísticos de un grupo y del evento.
RUBROS_GRUPO = ["vuelo", "transporte", "hospedaje", "alimentacion"]
RUBROS_EVENTO = ["salon", "facilitacion", "materiales"]
NOMBRE_RUBRO = {
    "vuelo": "Tiquetes aéreos",
    "transporte": "Transporte terrestre",
    "hospedaje": "Hospedaje",
    "alimentacion": "Alimentación",
    "salon": "Salón / espacio",
    "facilitacion": "Quién facilita",
    "materiales": "Materiales y kits",
}


class Tarifa(BaseModel):
    id: str
    concepto: str = ""
    unidad: str = ""
    zona: str = "Todas"
    valor: float
    moneda: Literal["COP", "USD", "EUR"] = "COP"
    bloque: Optional[str] = None  # bloque de costo cuando se usa en Otros costos
    fuente: str = ""
    fecha: Optional[str] = None
    estado: str = "Por confirmar"
    notas: str = ""


class Lugar(BaseModel):
    nombre: str
    departamento: str
    zona: str  # Bogotá · Capital departamental · Municipio · Rural remoto
    notas: str = ""


class Cargo(BaseModel):
    cargo: str
    costo_mensual_cop: float
    fuente: str = ""
    estado: str = "Por confirmar"
    notas: str = ""


class Grupo(BaseModel):
    """Personas que viajan de la misma forma."""
    nombre: str
    rol: Rol = "participantes"
    personas: float = 0
    llegada: Llegada = "local"
    dias: float = 0           # días en el lugar, contando los de viaje
    noches: float = 0
    comidas: Comidas = "completa"
    por_habitacion: int = 1
    # Transporte: para el equipo es por día (aeropuerto + internos); para
    # participantes, por trayecto ida y vuelta. None = valor por defecto.
    unidad_transporte: Optional[Unidad] = None
    dias_transporte: Optional[float] = None
    # Tarifas: None = la de defecto según rol y llegada.
    tarifa_vuelo: Optional[str] = None
    tarifa_transporte: Optional[str] = None
    tarifa_hospedaje: Optional[str] = None
    tarifa_alimentacion: Optional[str] = None
    zona_tarifas: Optional[str] = None  # ej. Rural remoto para un vuelo regional
    beneficiarios: bool = True  # False = no cuentan como beneficiarios (acompañantes, ya contados)
    notas: str = ""


class Kit(BaseModel):
    tarifa: str
    por: Literal["participante", "evento"] = "participante"
    cantidad: float = 1


class TipoActividad(BaseModel):
    """Plantilla: el perfil típico de un tipo de evento. Se define una vez."""
    nombre: str
    descripcion: str = ""
    grupos: list[Grupo] = Field(default_factory=list)
    dias_salon: float = 0
    kits: list[Kit] = Field(default_factory=list)
    mismos_participantes: bool = True  # en repeticiones: las mismas personas
    no_necesario: dict[str, str] = Field(default_factory=dict)  # rubro -> razón
    notas: str = ""


class ItemChequeo(BaseModel):
    """Cosas del proyecto que se olvidan (no dependen de un evento)."""
    id: str
    nombre: str
    pista: str = ""
    tarifas: list[str] = Field(default_factory=list)       # se da por costeado si se usa alguna
    cargos: list[str] = Field(default_factory=list)
    tipos_actividad: list[str] = Field(default_factory=list)


class Banco(BaseModel):
    tarifas: list[Tarifa] = Field(default_factory=list)
    lugares: list[Lugar] = Field(default_factory=list)
    cargos: list[Cargo] = Field(default_factory=list)
    tipos: list[TipoActividad] = Field(default_factory=list)
    # bloque interno -> {donante: categoría}
    mapeo: dict[str, dict[str, str]] = Field(default_factory=dict)
    categorias: dict[str, list[str]] = Field(default_factory=dict)  # donante -> categorías en orden
    chequeo: list[ItemChequeo] = Field(default_factory=list)


class Actividad(BaseModel):
    codigo: str
    nombre: str
    tipo: Optional[str] = None
    eventos_meta: Optional[int] = None       # lo que dice la propuesta
    participantes_meta: Optional[int] = None
    notas: str = ""


class Evento(BaseModel):
    actividad: str            # código
    lugar: str
    repeticiones: int = 1
    descripcion: str = ""
    tipo: Optional[str] = None  # None = el tipo de la actividad
    ajustes: dict[str, dict] = Field(default_factory=dict)   # grupo -> {campo: valor}
    grupos_extra: list[Grupo] = Field(default_factory=list)
    zona: Optional[str] = None  # None = la del lugar
    dias_salon: Optional[float] = None
    no_necesario: dict[str, str] = Field(default_factory=dict)  # "grupo:rubro" o rubro de evento -> razón
    notas: str = ""


class LineaPersonal(BaseModel):
    actividad: str
    cargo: str
    personas: float = 1
    meses: float = 1
    dedicacion: float = 1
    notas: str = ""


class OtroCosto(BaseModel):
    actividad: str
    tarifa: str
    cantidad: float = 1
    factor: float = 1
    unidad_factor: str = "vez"
    zona: Optional[str] = None
    bloque: Optional[str] = None  # None = el de la tarifa
    notas: str = ""


class Parametros(BaseModel):
    nombre: str = "Proyecto nuevo"
    donante: str = "CBPF"
    moneda: Literal["COP", "USD", "EUR"] = "COP"
    trm_usd: float = 4000       # COP por USD
    trm_eur: float = 4400       # COP por EUR
    pct_indirectos: float = 0.07
    techo: Optional[float] = None  # en moneda de presentación


class Proyecto(BaseModel):
    parametros: Parametros = Field(default_factory=Parametros)
    actividades: list[Actividad] = Field(default_factory=list)
    eventos: list[Evento] = Field(default_factory=list)
    personal: list[LineaPersonal] = Field(default_factory=list)
    otros: list[OtroCosto] = Field(default_factory=list)
    tarifas: list[Tarifa] = Field(default_factory=list)  # reemplazan las del banco (mismo id y zona)
    tipos: list[TipoActividad] = Field(default_factory=list)  # reemplazan los del banco (mismo nombre)
    lugares: list[Lugar] = Field(default_factory=list)  # se suman a los del banco
    cargos: list[Cargo] = Field(default_factory=list)  # reemplazan los del banco (mismo cargo)
    mapeo: dict[str, dict[str, str]] = Field(default_factory=dict)  # reemplaza el mapeo del banco por bloque
    chequeo: dict[str, str] = Field(default_factory=dict)  # id -> razón de "no aplica"
    notas: str = ""
