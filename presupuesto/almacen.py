"""Leer y guardar banco y proyectos como JSON legible."""
from __future__ import annotations

import json
from pathlib import Path
from typing import TypeVar

from pydantic import BaseModel

from .modelo import Banco, Proyecto

M = TypeVar("M", bound=BaseModel)


def guardar(obj: BaseModel, ruta: Path) -> None:
    ruta = Path(ruta)
    ruta.parent.mkdir(parents=True, exist_ok=True)
    data = obj.model_dump(mode="json", exclude_defaults=True)
    tmp = ruta.with_suffix(ruta.suffix + ".tmp")
    tmp.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
    tmp.replace(ruta)  # escritura atómica: nunca queda un archivo a medias


def cargar_banco(ruta: Path) -> Banco:
    return Banco.model_validate_json(Path(ruta).read_text(encoding="utf-8"))


def cargar_proyecto(ruta: Path) -> Proyecto:
    return Proyecto.model_validate_json(Path(ruta).read_text(encoding="utf-8"))
