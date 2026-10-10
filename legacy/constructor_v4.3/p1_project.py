import json, re, sys
import pandas as pd

SRC, OUT = sys.argv[1], sys.argv[2]
v = pd.read_excel(SRC, header=None)
LINES = {}
for _, row in v.iterrows():
    lab = row[0]
    if isinstance(lab, str) and re.match(r"^[A-G]\d+\.", lab.strip()):
        key = lab.strip().split(".")[0]
        LINES[key] = {"titulo": lab.strip(), "narr": str(row[1]), "total": float(row[7])}


def rol(key):
    m = re.search(r"Rol:\s*(.*?)(\n|Cálculo)", LINES[key]["narr"], re.S)
    return m.group(1).strip() if m else None


TRM = 2900
SRCN = "Presupuesto aprobado CBPF (PRESUP1)"

# ---------------------------------------------------------------- tarifas (USD, approved unit costs)
def T(i, zona, usd, nota=None):
    return {"ID tarifa": i, "Zona": zona, "Valor": usd, "Moneda": "USD", "Fuente": SRCN, "Estado": "Confirmado", "Notas": nota}
tarifas = [
    T("HOSP-STAFF", "Todas", 62.07, "Visitas de seguimiento (E14, E18, E22)"),
    T("HOSP-STAFF", "Bogotá", 51.72, "Taller de arranque (E7)"),
    T("HOSP-STAFF", "Capital departamental", 51.72, "Brigadas (E3)"),
    T("HOSP-STAFF", "Rural remoto", 31.03, "Miraflores (E10)"),
    T("HOSP-PART", "Todas", 51.72, "ENTRÉNATE (B13)"),
    T("HOSP-PART", "Capital departamental", 55.17, "Brigadas (B4), curso Popayán (B27)"),
    T("HOSP-PART", "Bogotá", 66.00, "Fondo agotable en Bogotá (B8): incluye transporte aeropuerto e interno"),
    T("ALIM", "Todas", 27.59), T("REFRIG", "Todas", 3.45),
    T("TIQ", "Todas", 344.83), T("TIQ", "Rural remoto", 517.24, "Miraflores ↔ San José"),
    T("TTE-STAFF", "Todas", 17.24, "Brigadas: E2 = USD 51,72 por persona por fase de 3 días → 17,24/día"),
    T("TTE-STAFF", "Bogotá", 86.21, "Taller de arranque (E6)"),
    T("TTE-STAFF", "Municipio", 103.45, "Visitas de seguimiento (E13, E17, E21)"),
    T("TTE-STAFF", "Rural remoto", 103.45),
    T("TTE-PART-LOCAL", "Todas", 20.60), T("TTE-PART-MUNIC", "Todas", 41.30),
    T("SALON", "Todas", 224.14), T("SALON", "Capital departamental", 448.28), T("SALON", "Bogotá", 293.10),
    T("KIT-PED", "Todas", 17.24), T("KIT-ENT", "Todas", 586.00, "Banco v3: 1.699.915 COP (Planeación general)"),
    T("KIT-ERAE", "Todas", 448.00), T("MAT-DIF", "Todas", 2.93, "Promedio ponderado: folletos 0,69 · cuadernillos 6,85 · manillas 1,25"),
    T("KIT-SENAL", "Todas", 86.20), T("DOTAC", "Todas", 51.72), T("DISP-MOV", "Todas", 4482.76),
    T("CASH-EMERG", "Todas", 517.24), T("SEG-VIDA", "Todas", 36.20), T("CEL-MES", "Todas", 22.73, "Banco v3: 60.000 COP"),
    T("LAPTOP", "Todas", 939.73, "Incluye seguro del equipo"),
]

# ---------------------------------------------------------------- cargos (approved monthly cost in USD → COP)
BANK = {"Médico fisiatra": 18174876, "Fisioterapeuta": 5571357, "Técnico ortoprotesista": 5820452, "Trabajador(a) social": 3909818,
        "Terapeuta ocupacional": 5787990, "Líder de brigada": 4589705, "Gerente de cooperación e impacto social": 14473800,
        "Jefe ERAE y AIV": 10449900, "Profesional audiovisual": 7344000}
cg = [("Oficial de cumplimiento", 4990.97), ("Gerente de cooperación e impacto social", 4990.97), ("Jefe ERAE y AIV", 3587.59),
      ("Jefe de cooperación y enlace institucional", 3429.33), ("Jefe de género y políticas públicas", 2241.38),
      ("Profesional de monitoreo y seguimiento", 2241.38), ("Médico fisiatra", 6331.03), ("Líder de brigada", 1582.66),
      ("Fisioterapeuta", 1995.86), ("Trabajador(a) social", 1367.07), ("Técnico ortoprotesista", 1940.73),
      ("Jefe de comunicaciones", 2532.41), ("Profesional audiovisual", 2532.41), ("Terapeuta ocupacional", 1899.30),
      ("Analista financiero", 1034.48), ("Auxiliar administrativo", 758.62), ("Coordinador(a) departamental", 2086.97),
      ("Facilitador(a) AICMA", 1241.38)]
cargos = []
for n, usd in cg:
    note = f"Aprobado: USD {usd:,.2f}/mes."
    if n in BANK:
        note += f" Banco de salarios 2026: {BANK[n]:,.0f} COP (= USD {BANK[n]/TRM:,.2f})."
    cargos.append({"Cargo": n, "Costo mensual COP": round(usd * TRM, 2), "Fuente": SRCN, "Estado": "Confirmado", "Notas": note})

# ---------------------------------------------------------------- personal
pers = [("0.1", "Oficial de cumplimiento", 1, 10, .20, "A1"), ("0.1", "Gerente de cooperación e impacto social", 1, 6, .08, "A2"),
        ("1.3.1", "Jefe ERAE y AIV", 1, 12, .25, "A3"), ("0.1", "Jefe de cooperación y enlace institucional", 1, 7, .05, "A4"),
        ("0.1", "Jefe de género y políticas públicas", 1, 6, 1, "A5"), ("0.2", "Profesional de monitoreo y seguimiento", 1, 11.5, 1, "A6"),
        ("1.2.1", "Médico fisiatra", 1, 8, .15, "A7"), ("1.2.1", "Líder de brigada", 1, 8, .15, "A8"), ("1.2.1", "Fisioterapeuta", 1, 8, .15, "A9"),
        ("1.2.1", "Trabajador(a) social", 1, 10, .15, "A10"), ("1.2.1", "Técnico ortoprotesista", 1, 8, .15, "A11"),
        ("0.1", "Jefe de comunicaciones", 1, 7, .05, "A12"), ("0.1", "Profesional audiovisual", 1, 9, .05, "A13"),
        ("1.2.1", "Terapeuta ocupacional", 1, 6, .16, "A14"), ("0.1", "Analista financiero", 1, 11, 1, "A15"),
        ("0.1", "Auxiliar administrativo", 1, 11, 1, "A16"), ("1.3.1", "Coordinador(a) departamental", 4, 11, 1, "A17"),
        ("1.3.1", "Facilitador(a) AICMA", 8, 9, 1, "A18")]
personal = [{"Código WBS": c, "Cargo": n, "Número de personas": k, "Meses": m, "Dedicación": d,
             "Rol y justificación": f"[{ln}] " + (rol(ln) or ""), "Estado": "Confirmado"} for c, n, k, m, d, ln in pers]

# ---------------------------------------------------------------- otros costos
def O(cod, tid, zona, q, f, unidad, just, bloque=None):
    return {"Código WBS": cod, "ID tarifa": tid, "Zona": zona, "Cantidad": q, "Factor": f, "Unidad del factor": unidad,
            "Bloque (si difiere)": bloque, "Justificación": just, "Estado": "Confirmado"}
otros = [
    O("1.2.1", "DISP-MOV", None, 48, 1, "vez", "[B1] Dispositivos formulados en las 4 brigadas (12 por brigada)"),
    O("1.2.1", "DISP-MOV", None, 15, 1, "vez", "[B2] Fondo agotable de emergencia – atención en Bogotá"),
    O("1.2.1", "TIQ", "Rural remoto", 2, 3, "fases", "[B6] 2 beneficiarios de Miraflores a la brigada de San José"),
    O("1.2.1", "TIQ", None, 15, 3, "fases", "[B7] Fondo agotable: traslado de 15 beneficiarios a Bogotá en 3 fases"),
    O("1.2.2", "TIQ", "Rural remoto", 2, 1, "vez", "[B15] 2 asistentes ENTRÉNATE desde Miraflores"),
    O("1.2.3", "CASH-EMERG", None, 20, 1, "vez", "[B16] Víctimas de nuevos hechos con artefactos explosivos"),
    O("1.3.1", "KIT-ERAE", None, 22, 1, "vez", "[B17] 4 coordinadores + 8 facilitadoras + 10 curso ERAE Básico"),
    O("1.3.1", "MAT-DIF", None, 7600, 1, "vez", "[B18] 6.750 ERAE + 850 MCMT"),
    O("1.3.1", "REFRIG", None, 760, 10, "meses", "[B19] Talleres comunitarios ERAE/MCMT: 760 asistentes por mes"),
    O("1.3.1", "KIT-SENAL", None, 18, 1, "vez", "[B20] 18 escuelas"),
    O("1.3.1", "DOTAC", None, 17, 1, "vez", "[B21] Equipo de territorio (chaleco + gorra)"),
    O("1.3.1", "KIT-PED", None, 14, 1, "vez", "[B22] Kits del taller de arranque"),
    O("0.3", "SEG-VIDA", None, 28, 1, "año", "[A19] 28 colaboradores del proyecto"),
    O("0.3", "LAPTOP", None, 1, 1, "vez", "[C1] Monitoreo + jefe ERAE (compartido)"),
    O("0.3", "CEL-MES", None, 4, 11, "meses", "[G1] 4 coordinadores departamentales"),
]

# ---------------------------------------------------------------- plantillas
def P(name, desc, **k):
    base = {"Staff viaja": 0, "Staff local": 0, "Staff en avión": 0, "Días staff en sitio": 0, "Noches staff": 0, "Días transporte staff": 0,
            "Participantes": 0, "Acompañantes": 0, "Días participantes": 0, "Noches participantes": 0, "Porcentaje con hospedaje": 0,
            "Personas por habitación": 1, "Porcentaje con transporte": 0, "Días de salón": 0, "Alimentación participantes": "Ninguna",
            "Tarifa transporte participantes": None, "Kit por participante": None, "Mismos participantes en repeticiones": "Sí",
            "Estado": "Confirmado"}
    base.update(k); base["Plantilla"] = name; base["Descripción"] = desc
    return base
BRI, FON, ENT, ARR, CUR, VIS, MIR = ("Brigada de rehabilitación (fase)", "Atención en sede Bogotá (fondo agotable)", "Taller ENTRÉNATE",
                                     "Taller de arranque (equipo)", "Curso ERAE Básico", "Visita de seguimiento", "Visita facilitadora a Miraflores")
plantillas = [
    P(BRI, "Una fase de brigada en un lugar. Repeticiones = 3 fases.", **{"Staff viaja": 7, "Staff en avión": 7, "Días staff en sitio": 3, "Noches staff": 3,
      "Días transporte staff": 3, "Participantes": 12, "Acompañantes": 12, "Días participantes": 2, "Noches participantes": 1, "Porcentaje con hospedaje": 1,
      "Personas por habitación": 2, "Porcentaje con transporte": 1, "Alimentación participantes": "Completa", "Tarifa transporte participantes": "TTE-PART-LOCAL",
      "Notas": "E1–E4 y B3–B5. Equipo: fisiatra, líder de brigada, ortoprotesista, trabajador social, fisioterapeuta, TO, audiovisual."}),
    P(FON, "Beneficiarios que viajan a la sede CIREC en Bogotá, 3 fases. Personal en sede (sin viáticos).", **{"Participantes": 15, "Días participantes": 3,
      "Noches participantes": 3, "Porcentaje con hospedaje": 1, "Alimentación participantes": "Completa", "Notas": "B8–B9. Tiquetes de los beneficiarios en OtrosCostos (B7)."}),
    P(ENT, "Taller de 3 días con cuidadores y líderes en cabecera municipal.", **{"Participantes": 18, "Días participantes": 3, "Noches participantes": 2,
      "Porcentaje con hospedaje": 1, "Porcentaje con transporte": 1, "Días de salón": 2, "Alimentación participantes": "Completa",
      "Tarifa transporte participantes": "TTE-PART-LOCAL", "Kit por participante": "KIT-ENT", "Notas": "B10–B14. El presupuesto aprobado no incluye personal que viaje a facilitar."}),
    P(ARR, "Inducción del equipo implementador en Bogotá.", **{"Staff viaja": 12, "Staff en avión": 12, "Días staff en sitio": 5, "Noches staff": 5,
      "Días transporte staff": 4, "Días de salón": 4, "Notas": "E5–E8, B23. 4 coordinadores + 8 facilitadoras. El jefe ERAE va como fila complementaria."}),
    P(CUR, "Curso de 4 días para líderes comunitarios.", **{"Staff local": 2, "Días staff en sitio": 4, "Participantes": 10, "Días participantes": 4,
      "Noches participantes": 4, "Porcentaje con hospedaje": 1, "Porcentaje con transporte": 1, "Días de salón": 4, "Alimentación participantes": "Completa",
      "Tarifa transporte participantes": "TTE-PART-MUNIC", "Kit por participante": "KIT-PED", "Notas": "B24–B28. 2 facilitadores locales."}),
    P(VIS, "1 profesional viaja a territorio (coordinación, monitoreo, género).", **{"Staff viaja": 1, "Staff en avión": 1, "Días staff en sitio": 3,
      "Noches staff": 3, "Días transporte staff": 3, "Notas": "E16–E23. El jefe AIV ajusta a 4 días de alimentación (E15)."}),
    P(MIR, "Facilitadora AICMA de San José a Miraflores para talleres y señalización.", **{"Staff viaja": 1, "Staff en avión": 1, "Días staff en sitio": 7,
      "Noches staff": 7, "Días transporte staff": 0, "Notas": "E9–E11. Aprobado: 21 noches pero solo 14 días de alimentación (inconsistente)."}),
]

# ---------------------------------------------------------------- lugares extra
lugares_extra = []

# ---------------------------------------------------------------- WBS
NOTA_DECL = "Declarado según la narrativa del presupuesto aprobado; verificar contra el marco lógico."
wbs = [
    {"Código WBS": "0", "Nivel": "Resultado", "Nombre": "Gestión, coordinación y transversales"},
    {"Código WBS": "0.1", "Nivel": "Actividad", "Nombre": "Personal de gestión, administración y comunicaciones"},
    {"Código WBS": "0.2", "Nivel": "Actividad", "Nombre": "Coordinación, monitoreo y enfoque de género (visitas a territorio)", "Plantilla": VIS,
     "Eventos declarados": 22, "Notas": "10 viajes jefe AIV + 8 monitoreo + 4 género. " + NOTA_DECL},
    {"Código WBS": "0.3", "Nivel": "Actividad", "Nombre": "Seguros, equipos y comunicaciones del equipo"},
    {"Código WBS": "1", "Nivel": "Resultado", "Nombre": "Resultado 1 (nombre según marco lógico)"},
    {"Código WBS": "1.2", "Nivel": "Producto", "Nombre": "Rehabilitación física e inclusión de sobrevivientes (nombre según marco lógico)"},
    {"Código WBS": "1.2.1", "Nivel": "Actividad", "Nombre": "Brigadas de rehabilitación física (4 deptos × 3 fases) + fondo agotable en Bogotá", "Plantilla": BRI,
     "Eventos declarados": 15, "Participantes declarados": 63, "Perfil de beneficiarios": "Brigadas de rehabilitación", "Notas": "12 fases de brigada + 3 fases en Bogotá; 48 + 15 beneficiarios. " + NOTA_DECL},
    {"Código WBS": "1.2.2", "Nivel": "Actividad", "Nombre": "Talleres ENTRÉNATE (1 por departamento)", "Plantilla": ENT,
     "Eventos declarados": 4, "Participantes declarados": 72, "Perfil de beneficiarios": "ENTRÉNATE", "Notas": "12 Guaviare, 15 Caquetá, 22 Cauca, 23 Antioquia. " + NOTA_DECL},
    {"Código WBS": "1.2.3", "Nivel": "Actividad", "Nombre": "Transferencias monetarias de emergencia", "Beneficiarios (si difiere)": 20, "Notas": "20 víctimas de nuevos hechos (B16). Sin perfil asignado: elegir o crear uno."},
    {"Código WBS": "1.3", "Nivel": "Producto", "Nombre": "ERAE y Mi Cuerpo Mi Territorio (nombre según marco lógico)"},
    {"Código WBS": "1.3.1", "Nivel": "Actividad", "Nombre": "ERAE y MCMT en territorio: taller de arranque, talleres comunitarios, señalización, Miraflores", "Plantilla": ARR,
     "Eventos declarados": 4, "Beneficiarios (si difiere)": 7600, "Perfil de beneficiarios": "Mi Cuerpo Mi Territorio (MCMT)", "Notas": "Beneficiarios: 7.600 asistentes (6.750 ERAE + 850 MCMT, B18). Se usó el perfil MCMT para todos: confirmar el perfil de ERAE. 1 taller de arranque + 3 visitas a Miraflores. Los talleres comunitarios (7.600 asistentes) se costean por unidad en OtrosCostos. " + NOTA_DECL},
    {"Código WBS": "1.3.2", "Nivel": "Actividad", "Nombre": "Curso ERAE Básico (Popayán)", "Plantilla": CUR,
     "Eventos declarados": 1, "Participantes declarados": 10, "Notas": "Sin perfil asignado (líderes comunitarios): elegir o crear uno. " + NOTA_DECL},
]

# ---------------------------------------------------------------- eventos
def EV(cod, lugar, desc, **k):
    d = {"Código WBS": cod, "Lugar": lugar, "Descripción": desc, "Estado": "Confirmado"}; d.update(k); return d
CAPZ = {"Ajuste zona": "Capital departamental"}
eventos = [
    EV("1.2.1", "San José del Guaviare", "Brigada Guaviare (3 fases)", Repeticiones=3, Notas="B5: el título dice solo Guaviare rural, pero la cantidad (288) cubre los 4 departamentos."),
    EV("1.2.1", "Por definir (Caquetá)", "Brigada Caquetá (3 fases)", Repeticiones=3, **CAPZ, Notas="Municipio no definido; se usa la tarifa de capital, como en el aprobado."),
    EV("1.2.1", "Por definir (Cauca)", "Brigada Cauca (3 fases)", Repeticiones=3, **CAPZ),
    EV("1.2.1", "Por definir (Antioquia)", "Brigada Antioquia (3 fases)", Repeticiones=3, **CAPZ),
    EV("1.2.1", "Bogotá", "Fondo agotable: atención en sede CIREC (3 fases)", Repeticiones=3, **{"Ajuste plantilla": FON}),
    EV("1.2.2", "Por definir (Guaviare)", "ENTRÉNATE Guaviare", **{"Participantes (si difiere)": 12}),
    EV("1.2.2", "Por definir (Caquetá)", "ENTRÉNATE Caquetá", **{"Participantes (si difiere)": 15}),
    EV("1.2.2", "Por definir (Cauca)", "ENTRÉNATE Cauca", **{"Participantes (si difiere)": 22}),
    EV("1.2.2", "Por definir (Antioquia)", "ENTRÉNATE Antioquia", **{"Participantes (si difiere)": 23}),
    EV("1.3.1", "Bogotá", "Taller de arranque – coordinadores y facilitadoras", Mes=1),
    EV("1.3.1", "Bogotá", "Taller de arranque – jefe ERAE y AIV (sin transporte local)", Mes=1, Complementaria="Sí",
       **{"Ajuste staff viaja": 1, "Ajuste staff en avión": 1, "Ajuste días transporte staff": 0, "Ajuste días de salón": 0},
       Notas="Fila complementaria: misma actividad y lugar; no cuenta como evento nuevo."),
    EV("1.3.1", "Miraflores (Guaviare)", "Visitas facilitadora AICMA a Miraflores", Repeticiones=3, **{"Ajuste plantilla": MIR},
       Notas="Aprobado: 21 noches pero 14 días de alimentación. El modelo usa 7 días por visita (21 días) → +USD 193,13."),
    EV("1.3.2", "Popayán", "Curso ERAE Básico"),
]
for dep, n in [("Guaviare", 2), ("Antioquia", 2), ("Caquetá", 3), ("Cauca", 3)]:
    eventos.append(EV("0.2", f"Por definir ({dep})", f"Jefe AIV – coordinación {dep}", Repeticiones=n, **{"Ajuste días staff en sitio": 4},
                      Notas="E12–E15: 4 días de alimentación, 3 noches, 3 días de transporte por viaje." if dep == "Guaviare" else None))
for dep in ["Guaviare", "Antioquia", "Caquetá", "Cauca"]:
    eventos.append(EV("0.2", f"Por definir ({dep})", f"Monitoreo – visita {dep}", Repeticiones=2,
                      Notas="E20–E23: 2 viajes por departamento." if dep == "Guaviare" else None))
for dep in ["Guaviare", "Antioquia", "Caquetá", "Cauca"]:
    eventos.append(EV("0.2", f"Por definir ({dep})", f"Género – visita {dep}", Notas="E16–E19: 1 viaje por departamento." if dep == "Guaviare" else None))

# ---------------------------------------------------------------- conciliación
def ev(block, **crit):
    c = "".join(f',tblEventos[{k}],"{v}"' for k, v in crit.items())
    return f'SUMIFS(tblEventos[COP {block}]{c})'
def evs(blocks, **crit):
    return "(" + "+".join(ev(b, **crit) for b in blocks) + ")/TRM"
def per(cargo):
    return f'SUMIFS(tblPersonal[Total COP],tblPersonal[Cargo],"{cargo}")/TRM'
def otr(i):
    return f'SUMIFS(tblOtros[Total COP],tblOtros[ID],"{i}")/TRM'
PU = "Plantilla usada"
L_ = []
def add(k, donde, expr, nota=None):
    L_.append({"linea": k, "concepto": LINES[k]["titulo"].split(". ", 1)[-1][:120], "aprobado": round(LINES[k]["total"], 2), "donde": donde, "expr": expr, "nota": nota})
L_.append({"seccion": "A. Personal"})
for c, n, k, m, d, ln in pers:
    add(ln, f"Personal › {n}", per(n))
add("A19", "OtrosCostos › O013 (SEG-VIDA)", otr("O013"), "Bloque Seguros → mapeado a la categoría A de CBPF, como en el aprobado.")
L_.append({"seccion": "B. Suministros y materiales"})
add("B1", "OtrosCostos › O001", otr("O001")); add("B2", "OtrosCostos › O002", otr("O002"))
add("B3", "Eventos 1.2.1 brigadas › Alimentación participantes", evs(["Alimentación participantes"], **{PU: BRI}))
add("B4", "Eventos 1.2.1 brigadas › Hospedaje participantes", evs(["Hospedaje participantes"], **{PU: BRI}))
add("B5", "Eventos 1.2.1 brigadas › Transporte participantes", evs(["Transporte participantes"], **{PU: BRI}))
add("B6", "OtrosCostos › O003", otr("O003")); add("B7", "OtrosCostos › O004", otr("O004"))
add("B8", "Evento Bogotá fondo agotable › Hospedaje participantes", evs(["Hospedaje participantes"], **{PU: FON}))
add("B9", "Evento Bogotá fondo agotable › Alimentación participantes", evs(["Alimentación participantes"], **{PU: FON}))
add("B10", "Eventos 1.2.2 › Materiales y kits", evs(["Materiales y kits"], **{"Código WBS": "1.2.2"}))
add("B11", "Eventos 1.2.2 › Alquiler de salones", evs(["Alquiler de salones"], **{"Código WBS": "1.2.2"}))
add("B12", "Eventos 1.2.2 › Alimentación participantes", evs(["Alimentación participantes"], **{"Código WBS": "1.2.2"}))
add("B13", "Eventos 1.2.2 › Hospedaje participantes", evs(["Hospedaje participantes"], **{"Código WBS": "1.2.2"}))
add("B14", "Eventos 1.2.2 › Transporte participantes", evs(["Transporte participantes"], **{"Código WBS": "1.2.2"}))
add("B15", "OtrosCostos › O005", otr("O005")); add("B16", "OtrosCostos › O006", otr("O006"))
for k, o in [("B17", "O007"), ("B18", "O008"), ("B19", "O009"), ("B20", "O010"), ("B21", "O011"), ("B22", "O012")]:
    add(k, f"OtrosCostos › {o}", otr(o))
add("B23", "Evento taller de arranque › Alquiler de salones", evs(["Alquiler de salones"], **{PU: ARR}))
add("B24", "Evento curso 1.3.2 › Alquiler de salones", evs(["Alquiler de salones"], **{"Código WBS": "1.3.2"}))
add("B25", "Evento curso 1.3.2 › Materiales y kits", evs(["Materiales y kits"], **{"Código WBS": "1.3.2"}))
add("B26", "Evento curso 1.3.2 › Alimentación participantes + staff", evs(["Alimentación participantes", "Alimentación staff"], **{"Código WBS": "1.3.2"}),
    "El modelo pone la alimentación de los 2 facilitadores (USD 220,72) en el bloque de staff → categoría E (Viajes). El aprobado la puso toda en B.")
add("B27", "Evento curso 1.3.2 › Hospedaje participantes", evs(["Hospedaje participantes"], **{"Código WBS": "1.3.2"}))
add("B28", "Evento curso 1.3.2 › Transporte participantes", evs(["Transporte participantes"], **{"Código WBS": "1.3.2"}))
L_.append({"seccion": "C. Equipos"})
add("C1", "OtrosCostos › O014", otr("O014"))
L_.append({"seccion": "E. Viajes"})
for k, b in [("E1", "Tiquetes aéreos staff"), ("E2", "Transporte terrestre staff"), ("E3", "Hospedaje staff"), ("E4", "Alimentación staff")]:
    add(k, f"Eventos 1.2.1 brigadas › {b}", evs([b], **{PU: BRI}))
for k, b in [("E5", "Tiquetes aéreos staff"), ("E6", "Transporte terrestre staff"), ("E7", "Hospedaje staff"), ("E8", "Alimentación staff")]:
    add(k, f"Eventos taller de arranque (2 filas) › {b}", evs([b], **{PU: ARR}))
for k, b in [("E9", "Tiquetes aéreos staff"), ("E10", "Hospedaje staff"), ("E11", "Alimentación staff")]:
    add(k, f"Evento Miraflores › {b}", evs([b], **{PU: MIR}),
        "Aprobado: 1 persona × 14 días. Con 3 visitas de 7 noches, el modelo calcula 21 días de alimentación." if k == "E11" else None)
for pref, ks in [("Jefe AIV*", ["E12", "E13", "E14", "E15"]), ("Género*", ["E16", "E17", "E18", "E19"]), ("Monitoreo*", ["E20", "E21", "E22", "E23"])]:
    for k, b in zip(ks, ["Tiquetes aéreos staff", "Transporte terrestre staff", "Hospedaje staff", "Alimentación staff"]):
        add(k, f"Eventos 0.2 '{pref[:-1]}' › {b}", evs([b], **{"Descripción": pref}))
L_.append({"seccion": "G. Gastos operativos"})
add("G1", "OtrosCostos › O015", otr("O015"))

conc = {
    "titulo": "Conciliación: presupuesto aprobado CBPF (PRESUP1) vs. este modelo",
    "subtitulo": "Cada línea del presupuesto aprobado y dónde quedó en el modelo. Verde = cuadra al centavo. Rojo = diferencia (ver comentario).",
    "lineas": L_, "directo": 859629.35, "psc": 60174.05, "total": 919803.40,
    "categorias": [("A. Staff and Other Personnel Costs", 285628.73), ("B. Supplies, Commodities, Materials", 481791.40), ("C. Equipment", 939.73),
                   ("D. Contractual Services", 0), ("E. Travel", 90269.37), ("F. Transfers and Grants to Counterparts", 0),
                   ("G. General Operating and Other Direct Costs", 1000.12)],
    "hallazgos": [
        "1. E11 (Miraflores): 3 visitas × 7 noches = 21 noches de hotel, pero solo 14 días de alimentación. El modelo exige coherencia: o son 21 días de comida (+USD 193,13) o menos noches.",
        "2. B26 (curso ERAE Básico): la alimentación de los 2 facilitadores está en B (suministros); por su naturaleza es viaje de personal (E). El total no cambia; cambia la categoría.",
        "3. B5: el título dice 'zona rural de San José del Guaviare', pero la cantidad (288 trayectos = 96 personas × 3 fases) cubre a los beneficiarios de los 4 departamentos.",
        "4. E2: el transporte terrestre del equipo de brigada (USD 51,72 por persona por fase) equivale a USD 17,24/día, muy por debajo de los USD 103,45/día de las visitas de seguimiento. Vale confirmar si es suficiente.",
        "5. ENTRÉNATE (1.2.2) y el fondo agotable no tienen personal que viaje: la auditoría lo marca en amarillo para confirmar quién facilita / atiende.",
        "6. Convenciones distintas en el mismo presupuesto: el jefe AIV tiene 4 días de comida por 3 noches (justificado en E15), pero monitoreo y género tienen 3 y 3. Conviene una regla única.",
        "7. Salarios: varios costos mensuales aprobados difieren del banco de salarios 2026 (ver Notas en la hoja Cargos).",
    ],
}

proj = {
    "params": {"Proyecto": "CBPF – ERAE, AIV, MCMT y ENTRÉNATE (Guaviare, Caquetá, Cauca, Antioquia) – presupuesto APROBADO",
               "Donante activo (define las categorías del Resumen)": "CBPF", "Moneda de presentación": "USD", "TRM (COP por USD)": 2900,
               "% costos indirectos": 0.07, "Base de los indirectos": "Costos directos", "Máximo % de indirectos permitido por el donante": 0.07,
               "TotalDonante": 919803.40},
    "tarifas": tarifas, "cargos": cargos, "personal": personal, "otros": otros, "plantillas": plantillas,
    "lugares_extra": lugares_extra, "wbs": wbs, "eventos": eventos,
    "mapeo": {"Seguros": ("A. Staff and Other Personnel Costs", "Como en el presupuesto aprobado (A19).")},
    "conciliacion": conc,
}
json.dump(proj, open(OUT, "w"), ensure_ascii=False, indent=1)
print("lines", len(LINES), "conc", sum(1 for x in L_ if "linea" in x))
