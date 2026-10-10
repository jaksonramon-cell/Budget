"""Abre el Presupuestador CIREC en el navegador.  python app.py"""
import sys

from presupuesto.app import main

if __name__ in {"__main__", "__mp_main__"}:
    main(abrir="--sin-navegador" not in sys.argv)
