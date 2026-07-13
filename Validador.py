import streamlit as st
import pandas as pd
import os
import shutil
import requests
import base64
import re


RUTA_TEMP = "temp.xlsx"

RUTA_OK_BASE = r"C:\Users\william.boconzaca\OneDrive - AGENCIA DE REGULACION Y CONTROL DE ELECTRICIDAD\Validador formularios\Archivos válidos"

# ================= VALIDACIONES CAMPOS SQL =================

DISTRIBUIDORAS = [
    'EE Santiago','CNEL EP Eficiencia Energética','CNEL EP Oficina Central','CNEL EP UN Bolívar',
    'CNEL EP UN El Oro','CNEL EP UN Esmeraldas','CNEL EP UN Guayaquil',
    'CNEL EP UN Guayas Los Ríos','CNEL EP UN Los Ríos','CNEL EP UN Manabí',
    'CNEL EP UN Milagro','CNEL EP UN Santa Elena','CNEL EP UN Santo Domingo',
    'CNEL EP UN Sucumbíos','EE Ambato','EE Azogues','EE Centro Sur',
    'EE Cotopaxi','EE Galápagos','EE Norte','EE Quito','EE Riobamba','EE Sur'
]

ETAPA_FUNCIONAL = [
    'Subtransmisión','Distribución','Administración',
    'Comercialización','Instalaciones de Servicio al Cliente'
]

TIPO_GASTO = [
    'Gastos administrativos','Gastos de venta','Gastos financieros'
]

GRUPO_GASTO = [
    'Materiales','Mano de Obra','Otros gastos','Servicios'
]

SN = ['Sí','No']
ESTADO_PROYECTO = ['Iniciado','No iniciado','Paralizado','Finalizado']
ETAPA_EJECUCION = [
    "No iniciado 0%",
    "Elaboración de pliegos 5%",
    "Disponibilidad de pliegos 10%",
    "Publicación del proceso 15%",
    "Resolución de adjudicación 30%",
    "Adjudicación y firma de contrato 40%",
    "Acta de entrega - recepción 95%",
    "En ejecución 41% - 90%",
    "Registro contable 100%"]
PERM_AMB = ['Certificado Ambiental','Licencia Ambiental','Registro Ambiental','No aplica']
FUENTES = [
    "PLANREP Crédito",
    "PMD  Crédito",
    "PMD BID I",
    "PMD BID II",
    "PMD BID III",
    "PMD CAF",
    "PMD AFD",
    "PMD BID V",
    "FERUM",
    "FERUM BID II",
    "FERUM BID III",
    "FERUM BID V",
    "PRIZA",
    "PMD Fiscales",
    "FERUM Fiscales",
    "BID I",
    "BID II",
    "BIDIII",
    "BID IV",
    "BID V",
    "BID VI",
    "Otras Inversiones y Convenios",
    "Aportes de accionistas",
    "Contribuciones de usuarios"
]

SERVICIOS = [
    "SPEE",
    "SAPG",
    "SCVE"
]


CLAVES_DISTRIBUIDORAS = {

    "EE Santiago": "clave001",
    "CNEL EP Oficina Central": "clave002",
    "CNEL EP UN Bolívar": "clave003",
    "CNEL EP UN El Oro": "clave004",
    "CNEL EP UN Esmeraldas": "clave005",
    "CNEL EP UN Guayaquil": "clave006",
    "CNEL EP UN Guayas Los Ríos": "clave007",
    "CNEL EP UN Los Ríos": "clave008",
    "CNEL EP UN Manabí": "clave009",
    "CNEL EP UN Milagro": "clave010",
    "CNEL EP UN Santa Elena": "clave011",
    "CNEL EP UN Santo Domingo": "clave012",
    "CNEL EP UN Sucumbíos": "clave013",
    "EE Ambato": "clave014",
    "EE Azogues": "clave015",
    "EE Centro Sur": "clave016",
    "EE Cotopaxi": "clave017",
    "EE Galápagos": "clave018",
    "EE Norte": "clave019",
    "EE Quito": "clave020",
    "EE Riobamba": "clave021",
    "EE Sur": "clave022"
}


# ================= LIMPIAR COLUMNAS =================

def validar_duplicados(df, nombre_form):

    errores = []

    df_aux = (
        df.fillna("")
          .astype(str)
          .apply(lambda col: col.str.strip())
    )

    duplicados = df_aux.duplicated(keep=False)

    for i in df_aux[duplicados].index:

        errores.append({
            **df.loc[i].to_dict(),
            "Formulario": nombre_form,
            "Fila": i + 2,
            "Error": "duplicado"
        })

    return errores


def limpiar_columnas(cols):
    return [
        str(c).strip().lower().replace("\\_", "_").replace(" ", "_")
        for c in cols
    ]

def validar_columnas(df, columnas_correctas):
    return limpiar_columnas(df.columns) == limpiar_columnas(columnas_correctas)

# ================= VALIDACIONES =================

def es_decimal(valor):
    try:
        float(valor)
        return True
    except:
        return False
    
def validar_catalogo(df, columna, catalogo, nombre_form):

    errores = []

    for i, row in df.iterrows():

        valor = str(row[columna]).strip()

        if valor and valor not in catalogo:

            
         errores.append(
             {
                 **row.to_dict(),
                 "Formulario": nombre_form,
                 "Fila": i + 2,
                 "Error": f"valor inválido '{valor}' en {columna}"
                 }
                 )


    return errores


def validar_numericos(df, columnas, nombre_form):

    errores = []

    for i, row in df.iterrows():

        for col in columnas:

            if col not in df.columns:
                continue

            valor = row[col]

            if pd.isna(valor):
                continue

            try:
                float(valor)

            except:
                errores.append({
                    **row.to_dict(),
                    "Formulario": nombre_form,
                    "Fila": i + 2,
                    "Error": f"{col} debe ser numérico"
                    }
                    )


    return errores

def validar_fechas(df, columnas, nombre_form):

    errores = []

    for i, row in df.iterrows():

        for col in columnas:

            if col not in df.columns:
                continue

            valor = str(row[col]).strip()

            if valor == "":
                continue

            try:
                pd.to_datetime(
                    valor,
                    dayfirst=True
                )

            except:

                errores.append({
                    **row.to_dict(),
                    "Formulario": nombre_form,
                    "Fila": i + 2,
                    "Error": f"{col} no es una fecha válida"
                })

    return errores



def validar_form2(df):
    errores = []
    for i, row in df.iterrows():

        fila = i + 2
        error = ""

        # SQL
        if row["distribuidora"] not in DISTRIBUIDORAS:
            error += "distribuidora inválida; "
        if row["etapa_funcional"] not in ETAPA_FUNCIONAL:
            error += "etapa_funcional inválido; "
        if row["tipo_gasto"] not in TIPO_GASTO:
            error += "tipo_gasto inválido; "
        if row["grupo_gasto"] not in GRUPO_GASTO:
            error += "grupo_gasto inválido; "

        # ORIGINAL
        if not str(row["nro_partida"]).isdigit():
            error += "nro_partida inválido; "

        for c in [
            "asignacion_inicial","reformas","presupuesto_codificado",
            "pre_compromiso","compromiso","devengado","pagado"
        ]:
            if not es_decimal(row[c]):
                error += f"{c} inválido; "

        if error:
            errores.append({**row,"Formulario":"FORM2","Fila":fila,"Error":error})

    return errores

def validar_form6(df):

    errores = []

    for i, row in df.iterrows():

        fila = i + 2
        error = ""

        if row["distribuidora"] not in DISTRIBUIDORAS:
            error += "distribuidora inválida; "

        if row["tipo_gasto"] not in TIPO_GASTO:
            error += "tipo_gasto inválido; "

        if row["grupo_gasto"] not in GRUPO_GASTO:
            error += "grupo_gasto inválido; "

        if not str(row["nro_partida"]).isdigit():
            error += "nro_partida inválido; "

        for c in [
            "asignacion_inicial",
            "reformas",
            "presupuesto_codificado",
            "pre_compromiso",
            "compromiso",
            "devengado",
            "pagado"
        ]:

            if not es_decimal(row[c]):
                error += f"{c} inválido; "

        if error:

            errores.append({
                **row.to_dict(),
                "Formulario": "FORM6",
                "Fila": fila,
                "Error": error
            })

    return errores

def validar_presupuesto(df, nombre):

    errores = []

    for i, row in df.iterrows():

        try:

            presupuesto = float(row["presupuesto_codificado"])
            compromiso = float(row["compromiso"])
            devengado = float(row["devengado"])
            pagado = float(row["pagado"])

            if compromiso > presupuesto:
                errores.append({
                    **row.to_dict(),
                    "Formulario": nombre,
                    "Fila": i + 2,
                    "Error": "compromiso mayor al presupuesto codificado"
                    }
                    )



            if devengado > compromiso:
                errores.append({
                    **row.to_dict(),
                    "Formulario": nombre,
                    "Fila": i + 2,
                    "Error": "devengado mayor al compromiso"
                    }
                    )

            if pagado > devengado:

                errores.append({
                    **row.to_dict(),
                    "Formulario": nombre,
                    "Fila": i + 2,
                    "Error": "pagado mayor al devengado"
                    }
                    )

        except:
            pass

    return errores

def validar_form3(df):

    errores = []

    errores.extend(
        validar_catalogo(
            df,
            "distribuidora",
            DISTRIBUIDORAS,
            "FORM3"
        )
    )

    errores.extend(
        validar_catalogo(
            df,
            "estado_proyecto",
            ESTADO_PROYECTO,
            "FORM3"
        )
    )

    errores.extend(
        validar_catalogo(
            df,
            "perm_amb",
            PERM_AMB,
            "FORM3"
        )
    )

    errores.extend(
    validar_fechas(
        df,
        [
            "fecha_inicio_proyecto",
            "fecha_pro_fin_proyecto",
            "fecha_fin_proyecto",
            "fecha_perm_amb_planif",
            "fecha_perm_amb_ejec"
        ],
        "FORM3"
    )
)
    

    columnas_numericas = [

        "avance_ejecucion_fisica",
        "avance_ejecucion_total",
        "monto_calificado",
        "presupuesto_codificado_arrastre",
        "devengado_arrastre",
        "pagado_arrastre",
        "asignacion_inicial",
        "reformas",
        "presupuesto_codificado",
        "pre_compromiso",
        "compromiso",
        "devengado",
        "pagado",
        "anticipo_no_amortizado",
        "empleos_generados",
        "nro_personal_fem",
        "bd_planificado",
        "bd_ejecutado",
        "vcs_planificado",
        "vcs_ejecutado",
        "vss_planificado",
        "vss_ejecutado",
        "tv_planificado",
        "tv_ejecutado",
        "ln_planificado",
        "ln_ejecutado",
        "at_planificado",
        "at_ejecutado",
        "mt_planificado",
        "mt_ejecutado",
        "bt_planificado",
        "bt_ejecutado",
        "am_planificado",
        "am_ejecutado",
        "m_planificado",
        "m_ejecutado",
        "td_planificado",
        "td_ejecutado",
        "pitd_planificado",
        "pitd_ejecutado",
        "sdn_planificado",
        "sdn_ejecutado",
        "pisdn_planificado",
        "pisdn_ejecutado"
    ]

    errores.extend(
        validar_numericos(
            df,
            columnas_numericas,
            "FORM3"
        )
    )

    errores.extend(
        validar_presupuesto(
            df,
            "FORM3"
        )
    )

    return errores

def validar_decimal_generico(df, columna, nombre):
    errores = []
    for i, row in df.iterrows():

        fila = i + 2
        error = ""

        if "distribuidora" in df.columns and row["distribuidora"] not in DISTRIBUIDORAS:
            error += "distribuidora inválida; "

        if "proyecto_arrastre" in df.columns and row["proyecto_arrastre"] not in SN:
            error += "proyecto_arrastre inválido; "

        if "estado_proyecto" in df.columns and row["estado_proyecto"] not in ESTADO_PROYECTO:
            error += "estado_proyecto inválido; "

        if "perm_amb" in df.columns and row["perm_amb"] not in PERM_AMB:
            error += "perm_amb inválido; "

        # ORIGINAL
        if not es_decimal(row[columna]):
            error += f"{columna} inválido; "

        if error:
            errores.append({**row,"Formulario":nombre,"Fila":fila,"Error":error})

    return errores

def validar_form4(df):

    errores = []
    for i, row in df.iterrows():
        if row["proyecto_arrastre"] not in SN:
            errores.append({
            **row.to_dict(),
            "Formulario": "FORM4",
            "Fila": i + 2,
            "Error": "proyecto_arrastre inválido"
        })

    errores.extend(
        validar_catalogo(
            df,
            "distribuidora",
            DISTRIBUIDORAS,
            "FORM4"
        )
    )

    errores.extend(
        validar_catalogo(
            df,
            "estado_proyecto",
            ESTADO_PROYECTO,
            "FORM4"
        )
    )

    errores.extend(
    validar_catalogo(
        df,
        "etapa_funcional",
        ETAPA_FUNCIONAL,
        "FORM4"
    )
    )

    errores.extend(
        validar_catalogo(
            df,
            "perm_amb",
            PERM_AMB,
            "FORM4"
        )
    )

    errores.extend(
    validar_fechas(
        df,
        [
            "fecha_inicio_proyecto",
            "fecha_pro_fin_proyecto",
            "fecha_fin_proyecto",
            "fecha_permiso_planif",
            "fecha_permiso_ejec"
        ],
        "FORM4"
    )
)


    columnas_numericas = [

        "avance_ejecucion_fisica",
        "avance_ejecucion_total",

        "presupuesto_codificado_arrastre",
        "devengado_arrastre",
        "pagado_arrastre",

        "asignacion_inicial",
        "reformas",
        "presupuesto_codificado",

        "pre_compromiso",
        "compromiso",
        "devengado",
        "pagado",

        "anticipo_no_amortizado",

        "bd_planificado",
        "bd_ejecutado",

        "vcs_planificado",
        "vcs_ejecutado",

        "vss_planificado",
        "vss_ejecutado",

        "tv_planificado",
        "tv_ejecutado",

        "ln_planificado",
        "ln_ejecutado",

        "at_planificado",
        "at_ejecutado",

        "mt_planificado",
        "mt_ejecutado",

        "bt_planificado",
        "bt_ejecutado",

        "am_planificado",
        "am_ejecutado",

        "m_planificado",
        "m_ejecutado",

        "td_planificado",
        "td_ejecutado",

        "pitd_planificado",
        "pitd_ejecutado",

        "sdn_planificado",
        "sdn_ejecutado",

        "pisdn_planificado",
        "pisdn_ejecutado",

        "empleos_generados",
        "nro_personal_fem"
    ]

    errores.extend(
        validar_numericos(
            df,
            columnas_numericas,
            "FORM4"
        )
    )

    errores.extend(
        validar_presupuesto(
            df,
            "FORM4"
        )
    )

    return errores

def validar_form7_sql(df):

    errores = []

    errores.extend(
        validar_catalogo(
            df,
            "distribuidora",
            DISTRIBUIDORAS,
            "FORM7"
        )
    )

    errores.extend(
    validar_catalogo(
        df,
        "proyecto_calificado_ecostos",
        SN,
        "FORM7"
    )
)

    errores.extend(
        validar_catalogo(
            df,
            "etapa_ejecucion_proyecto",
            ETAPA_EJECUCION,
            "FORM7"
        )
    )

    errores.extend(
        validar_catalogo(
            df,
            "estado_proyecto",
            ESTADO_PROYECTO,
            "FORM7"
        )
    )

    errores.extend(
    validar_fechas(
        df,
        [
            "fecha_inicio_proyecto",
            "fecha_pro_fin_proyecto",
            "fecha_fin_proyecto"
        ],
        "FORM7"  # cambiar según corresponda
    )
)
    

    for i, row in df.iterrows():

        if row["proyecto_arrastre"] not in SN:

            errores.append({
                **row.to_dict(),
                "Formulario": "FORM7",
                "Fila": i + 2,
                "Error": "proyecto_arrastre inválido"
            })

        if (
            row["proyecto_arrastre"] == "Sí"
            and row["rubro_arrastre"] not in [
                "Calidad",
                "Responsabilidad Ambiental",
                "Otros",
                "SIGDE",
                "Confiabilidad"
            ]
        ):

            errores.append({
                **row.to_dict(),
                "Formulario": "FORM7",
                "Fila": i + 2,
                "Error": "rubro_arrastre inválido"
            })

    # VALIDACIÓN DE CAMPOS NUMÉRICOS
    columnas_numericas = [

        "avance_ejecucion_fisica",
        "avance_ejecucion_total",
        "monto_calificado",
        "presupuesto_codificado_arrastre",
        "devengado_arrastre",
        "pagado_arrastre",
        "asignacion_inicial",
        "reformas",
        "presupuesto_codificado",
        "pre_compromiso",
        "compromiso",
        "devengado",
        "pagado",
        "anticipo_no_amortizado",
        "ln_inst_potencia",
        "ln_inst_cantidad",
        "ln_reemp_potencia",
        "ln_reemp_cantidad",
        "lumin_inst_total",
        "postes_nuev_número",
        "postes_reemp_número",
        "benef_dir_número",
        "empleos_generados",
        "nro_personal_fem"
    ]

    errores.extend(
        validar_numericos(
            df,
            columnas_numericas,
            "FORM7"
        )
    )

    errores.extend(
        validar_presupuesto(
            df,
            "FORM7"
        )
    )

    return errores


def validar_form8(df):

    errores = []

    errores.extend(
        validar_catalogo(
            df,
            "distribuidora",
            DISTRIBUIDORAS,
            "FORM8"
        )
    )

    errores.extend(
        validar_catalogo(
            df,
            "estado_proyecto",
            ESTADO_PROYECTO,
            "FORM8"
        )
    )

    errores.extend(
        validar_catalogo(
            df,
            "etapa_ejecucion_proyecto",
            ETAPA_EJECUCION,
            "FORM8"
        )
    )


    errores.extend(
        validar_fechas(
            df,
            [
                "fecha_inicio_proyecto",
                "fecha_pro_fin_proyecto",
                "fecha_fin_proyecto"
            ],
            "FORM8"
        )
    )


    for i, row in df.iterrows():

        if row["proyecto_arrastre"] not in SN:

            errores.append({
                **row.to_dict(),
                "Formulario": "FORM8",
                "Fila": i + 2,
                "Error": "proyecto_arrastre inválido"
            })

    columnas_numericas = [

        "avance_ejecucion_fisica",
        "avance_ejecucion_total",

        "presupuesto_codificado_arrastre",
        "devengado_arrastre",
        "pagado_arrastre",

        "asignacion_inicial",
        "reformas",
        "presupuesto_codificado",

        "pre_compromiso",
        "compromiso",
        "devengado",
        "pagado",

        "anticipo_no_amortizado",

        "ln_inst_potencia",
        "ln_inst_cantidad",

        "ln_reemp_potencia",
        "ln_reemp_cantidad",

        "lumin_inst_total",

        "postes_nuev_número",
        "postes_reemp_número",

        "benef_dir_número",

        "empleos_generados",
        "nro_personal_fem"
    ]

    errores.extend(
        validar_numericos(
            df,
            columnas_numericas,
            "FORM8"
        )
    )

    errores.extend(
        validar_presupuesto(
            df,
            "FORM8"
        )
    )

    return errores



def validar_form9(df):

    errores = []

    errores.extend(
        validar_catalogo(
            df,
            "distribuidora",
            DISTRIBUIDORAS,
            "FORM9"
        )
    )

    errores.extend(
        validar_catalogo(
            df,
            "fuente",
            FUENTES,
            "FORM9"
        )
    )

    errores.extend(
        validar_catalogo(
            df,
            "servicio",
            SERVICIOS,
            "FORM9"
        )
    )

    errores.extend(
        validar_catalogo(
            df,
            "estado_proyecto",
            ESTADO_PROYECTO,
            "FORM9"
        )
    )

    errores.extend(
    validar_catalogo(
        df,
        "etapa_ejecucion_proyecto",
        ETAPA_EJECUCION,
        "FORM9"
    )
)
    
    errores.extend(
        validar_fechas(
            df,
            [
                "fecha_inicio_proyecto",
                "fecha_pro_fin_proyecto",
                "fecha_fin_proyecto"
            ],
            "FORM9"
        )
    )


    for i, row in df.iterrows():

        if row["proyecto_arrastre"] not in SN:

            errores.append({
                **row.to_dict(),
                "Formulario": "FORM9",
                "Fila": i + 2,
                "Error": "proyecto_arrastre inválido"
            })

    columnas_numericas = [

        "avance_ejecucion_fisica",
        "avance_ejecucion_total",

        "presupuesto_codificado_arrastre",
        "devengado_arrastre",
        "pagado_arrastre",

        "asignacion_inicial",
        "reformas",
        "presupuesto_codificado",

        "pre_compromiso",
        "compromiso",
        "devengado",
        "pagado",

        "anticipo_no_amortizado"
    ]

    errores.extend(
        validar_numericos(
            df,
            columnas_numericas,
            "FORM9"
        )
    )

    errores.extend(
        validar_presupuesto(
            df,
            "FORM9"
        )
    )

    return errores

# ================= VALIDADORES =================

VALIDADORES = {

    "FORM 2 GASTO AO&M-C SPEE": validar_form2,
    "FORM 3 ANUALIDAD ACTIVO SPEE": validar_form3,
    "FORM 9 OTROS RECURSOS": validar_form9,
    "FORM 4 EXPANSION SPEE": validar_form4,
    "FORM 6 GASTO AO&M SAPG": validar_form6,
    "FORM 8 EXPANSION SAPG": validar_form8
}

VALIDADORES["FORM 7 ANUALIDAD ACTIVO SAPG"] = validar_form7_sql

# ================= VALIDACIONES NOMBRES COLUMNAS =================

FORMULARIOS = {

    "FORM 2 GASTO AO&M-C SPEE": [
        "distribuidora","nro_partida","descripcion_partida","etapa_funcional",
        "tipo_gasto","grupo_gasto","asignacion_inicial","reformas",
        "presupuesto_codificado","pre_compromiso","compromiso",
        "devengado","pagado"
    ],

    "FORM 3 ANUALIDAD ACTIVO SPEE": [
        "distribuidora","proyecto_arrastre","proyecto_calificado_ecostos",
        "anio_calificacion","codigo_ecostos","rubro_arrastre",
        "codigo_proyecto_eed","nombre_proyecto","objeto_proyecto",
        "provincia","canton","parroquia","aso_transmisión",
        "etapa_funcional","tipo_de_proyecto","estado_proyecto",
        "etapa_ejecucion_proyecto","avance_ejecucion_fisica",
        "avance_ejecucion_total","fecha_inicio_proyecto",
        "fecha_pro_fin_proyecto","fecha_fin_proyecto","monto_calificado",
        "presupuesto_codificado_arrastre","devengado_arrastre",
        "pagado_arrastre","asignacion_inicial","reformas",
        "presupuesto_codificado","pre_compromiso","compromiso",
        "devengado","pagado","anticipo_no_amortizado",
        "bd_planificado","bd_ejecutado","vcs_planificado","vcs_ejecutado",
        "vss_planificado","vss_ejecutado","tv_planificado","tv_ejecutado",
        "ln_planificado","ln_ejecutado","at_planificado","at_ejecutado",
        "mt_planificado","mt_ejecutado","bt_planificado","bt_ejecutado",
        "am_planificado","am_ejecutado","m_planificado","m_ejecutado",
        "td_planificado","td_ejecutado","pitd_planificado","pitd_ejecutado",
        "sdn_planificado","sdn_ejecutado","pisdn_planificado","pisdn_ejecutado",
        "perm_amb","fecha_perm_amb_planif","fecha_perm_amb_ejec",
        "empleos_generados","nro_personal_fem","observaciones"
    ],

    "FORM 4 EXPANSION SPEE": [
        "distribuidora","proyecto_arrastre","anio_calificacion","codigo_proyecto_eed",
        "nombre_proyecto","objeto_proyecto","provincia","canton","parroquia",
        "aso_transmisión","etapa_funcional","estado_proyecto",
        "etapa_ejecucion_proyecto","avance_ejecucion_fisica",
        "avance_ejecucion_total","fecha_inicio_proyecto",
        "fecha_pro_fin_proyecto","fecha_fin_proyecto",
        "presupuesto_codificado_arrastre","devengado_arrastre",
        "pagado_arrastre","asignacion_inicial","reformas",
        "presupuesto_codificado","pre_compromiso","compromiso",
        "devengado","pagado","anticipo_no_amortizado",
        "bd_planificado","bd_ejecutado","vcs_planificado","vcs_ejecutado",
        "vss_planificado","vss_ejecutado","tv_planificado","tv_ejecutado",
        "ln_planificado","ln_ejecutado","at_planificado","at_ejecutado",
        "mt_planificado","mt_ejecutado","bt_planificado","bt_ejecutado",
        "am_planificado","am_ejecutado","m_planificado","m_ejecutado",
        "td_planificado","td_ejecutado","pitd_planificado","pitd_ejecutado",
        "sdn_planificado","sdn_ejecutado","pisdn_planificado","pisdn_ejecutado",
        "perm_amb","fecha_permiso_planif","fecha_permiso_ejec",
        "empleos_generados","nro_personal_fem","observaciones"],
    "FORM 6 GASTO AO&M SAPG": ["distribuidora","nro_partida","descripcion_partida","tipo_gasto",
        "grupo_gasto","asignacion_inicial","reformas",
        "presupuesto_codificado","pre_compromiso","compromiso",
        "devengado","pagado"],
    "FORM 7 ANUALIDAD ACTIVO SAPG": [
        "distribuidora","proyecto_arrastre","proyecto_calificado_ecostos",
        "anio_calificacion","codigo_ecostos","rubro_arrastre",
        "codigo_proyecto_eed","nombre_proyecto","objeto_proyecto",
        "provincia","canton","parroquia","estado_proyecto",
        "etapa_ejecucion_proyecto","avance_ejecucion_fisica",
        "avance_ejecucion_total","fecha_inicio_proyecto",
        "fecha_pro_fin_proyecto","fecha_fin_proyecto","monto_calificado",
        "presupuesto_codificado_arrastre","devengado_arrastre",
        "pagado_arrastre","asignacion_inicial","reformas",
        "presupuesto_codificado","pre_compromiso","compromiso",
        "devengado","pagado","anticipo_no_amortizado",
        "ln_inst_tipo","ln_inst_potencia","ln_inst_cantidad",
        "ln_reemp_tipo","ln_reemp_potencia","ln_reemp_cantidad",
        "lumin_inst_total","postes_nuev_número","postes_reemp_número",
        "benef_dir_número","empleos_generados","nro_personal_fem","observaciones"
    ],
    "FORM 8 EXPANSION SAPG":  [
        "distribuidora","proyecto_arrastre","anio_calificacion",
        "codigo_proyecto_eed","nombre_proyecto","objeto_proyecto",
        "provincia","canton","parroquia","estado_proyecto",
        "etapa_ejecucion_proyecto","avance_ejecucion_fisica",
        "avance_ejecucion_total","fecha_inicio_proyecto",
        "fecha_pro_fin_proyecto","fecha_fin_proyecto",
        "presupuesto_codificado_arrastre","devengado_arrastre",
        "pagado_arrastre","asignacion_inicial","reformas",
        "presupuesto_codificado","pre_compromiso","compromiso",
        "devengado","pagado","anticipo_no_amortizado",
        "ln_inst_tipo","ln_inst_potencia","ln_inst_cantidad",
        "ln_reemp_tipo","ln_reemp_potencia","ln_reemp_cantidad",
        "lumin_inst_total","postes_nuev_número","postes_reemp_número",
        "benef_dir_número","empleos_generados","nro_personal_fem","observaciones"
    ],
"FORM 9 OTROS RECURSOS":[
    "distribuidora","fuente","proyecto_arrastre","servicio","codigo_proyecto_eed","nombre_proyecto",
    "objeto_proyecto","provincia","canton","parroquia","estado_proyecto","etapa_ejecucion_proyecto",
    "avance_ejecucion_fisica","avance_ejecucion_total","fecha_inicio_proyecto","fecha_pro_fin_proyecto",
    "fecha_fin_proyecto","presupuesto_codificado_arrastre","devengado_arrastre","pagado_arrastre","asignacion_inicial",
    "reformas","presupuesto_codificado","pre_compromiso","compromiso","devengado","pagado","anticipo_no_amortizado",
    "observaciones"
]   
}


# *************************************************************
# ***************** INTERFAZ DE LA APLICACIÓN *****************
# *************************************************************


st.markdown("""
<style>
.stDownloadButton button {
    width: 100%;
    font-size: 11px !important;
    padding: 0.25rem 0.4rem !important;
}
</style>
""", unsafe_allow_html=True)

col1, col2 = st.columns([7, 5])

with col1:
    st.markdown(
        "### 📈 Sistema de validaciones para los formularios de Ejecución de Recursos"
    )

with col2:

    # FILA SUPERIOR
    fila1_col1, fila1_col2 = st.columns(2)

    with fila1_col1:
        with open("Instructivo_Formularios_Junio_2026.xlsx", "rb") as f:
            st.download_button(
                "📖 Instructivo para el llenado de formularios",
                f,
                file_name="Instructivo_Formularios_Junio_2026.xlsx"
            )

    with fila1_col2:
        with open("Instructivo_Entrega de información.docx", "rb") as f:
            st.download_button(
                "📖 Instructivo para la carga de archivos",
                f,
                file_name="Instructivo_Entrega de información.docx"
            )

    # FILA INFERIOR
    fila2_col1, fila2_col2 = st.columns(2)

    with fila2_col1:
        with open(
            "Formularios de Control Ejecución Presupuestos_Proyectos.xlsx",
            "rb"
        ) as f:
            st.download_button(
                "📊 Plantillas de formularios de control",
                f,
                file_name="Formularios de Control Ejecución Presupuestos_Proyectos.xlsx"
            )

    with fila2_col2:
        with open(
            "Nombres distribuidoras válidos.xlsx",
            "rb"
        ) as f:
            st.download_button(
                "📋 Catálogo de nombres de carga",
                f,
                file_name="Nombres distribuidoras válidos.xlsx"
            )



st.markdown("""
##### Antes de cargar un archivo tome en cuenta los siguientes lineamientos:

📋✔️ Formato permitido: **.xlsx**\n
📋✔️ El nombre del archivo debe seguir la estructura:\n

`Distribuidora mes año`

📋✔️ El archivo debe contener los formularios oficiales vigentes, especificados en la plantilla.\n
📋✔️ Los nombres de las columnas no deben ser modificados para que su archivo sea validado correctamente.\n
📋✔️ En caso de que la información de los formularios no se ajuste a las directrices y/o validaciones especificadas en el instructivo, se le generará un archivo de descarga para su respectiva corrección.\n
📋✔️ Los archivos válidos serán almacenados automáticamente en el repositorio institucional.\n
""")


archivo = st.file_uploader("**Por favor cargue su Formulario.xlsx**", type=["xlsx"])

if archivo:

    distribuidora_seleccionada = st.selectbox(
        "🏢 Seleccione la distribuidora",
        DISTRIBUIDORAS
    )

    clave_ingresada = st.text_input(
        "🔑 Contraseña",
        type="password"
    )

if archivo:

    EMPRESAS = [
        "CNEL EP Oficina Central",
        "EE Santiago",
        "CNEL EP UN Bolívar",
        "CNEL EP UN El Oro",
        "CNEL EP UN Esmeraldas",
        "CNEL EP UN Guayaquil",
        "CNEL EP UN Guayas Los Ríos",
        "CNEL EP UN Los Ríos",
        "CNEL EP UN Manabí",
        "CNEL EP UN Milagro",
        "CNEL EP UN Santa Elena",
        "CNEL EP UN Santo Domingo",
        "CNEL EP UN Sucumbíos",
        "EE Ambato",
        "EE Azogues",
        "EE Centro Sur",
        "EE Cotopaxi",
        "EE Galápagos",
        "EE Norte",
        "EE Quito",
        "EE Riobamba",
        "EE Sur"
    ]

    nombre = archivo.name.rsplit(".", 1)[0].strip()

    valido = False

    for empresa in EMPRESAS:

        patron = (
            "^"
            + re.escape(empresa)
            + r"\s+"
            + r"(enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|octubre|noviembre|diciembre)"
            + r"\s+\d{4}$"
        )

        if re.match(
            patron,
            nombre
        ):
            valido = True
            break

    if not valido:

        st.error(
            "❌ De acuerdo con lo socializado, el nombre del archivo debe mantener las siguientes directrices.\n\n"
            "• Para la Corporación Nacional de Electricidad:\n"
            "CNEL EP UN El Oro mayo 2027\n\n"
            "• Para las Empresas Eléctricas:\n"
            "EE Quito enero 2027"
        )

        st.stop()


if archivo:

    with open(RUTA_TEMP, "wb") as f:
        f.write(archivo.getbuffer())

    if st.button("🔍 Validar"):

        # Validar contraseña vacía
        if not clave_ingresada:
            st.error("❌ Debe ingresar la contraseña.")
            st.stop()

        # Validar contraseña
        if clave_ingresada != CLAVES_DISTRIBUIDORAS.get(
            distribuidora_seleccionada
        ):
            st.error("❌ Contraseña incorrecta.")
            st.stop()

        # Validar distribuidora vs nombre del archivo
        if not nombre.upper().startswith(
            distribuidora_seleccionada.upper()
        ):
            st.error(
                "❌ La distribuidora seleccionada no coincide con el nombre del archivo."
            )
            st.stop()

        try:

            with pd.ExcelFile(RUTA_TEMP) as xls:

                faltantes = [h for h in FORMULARIOS if h not in xls.sheet_names]
                if faltantes:
                    st.error(f"❌ Faltan formularios: {', '.join(faltantes)}")
                    st.stop()

                errores_totales = []

                for hoja, columnas in FORMULARIOS.items():

                    df = pd.read_excel(xls, hoja, dtype=str).fillna("")

                    # ✅ VALIDACIÓN DE COLUMNAS:
                    if not validar_columnas(df, columnas):
                        st.error(f"❌ Error en {hoja}: columnas incorrectas")
                        st.stop()

                    if hoja in VALIDADORES:
                        errores_validacion = VALIDADORES[hoja](df)

                        if hoja not in [
                            "FORM 2 GASTO AO&M-C SPEE",
                            "FORM 6 GASTO AO&M SAPG"
                        ]:
                         errores_duplicados = validar_duplicados(df, hoja)
                        else:
                            errores_duplicados=[]


                        st.write(
                            hoja,
                            "errores validación:",
                            len(errores_validacion)
                            )
                        
                        st.write(
                            hoja,
                            "errores duplicados:",
                            len(errores_duplicados)
                            )
                        
                        errores_totales.extend(errores_validacion)
                        errores_totales.extend(errores_duplicados)


                st.write("TOTAL ERRORES:", len(errores_totales))

                if errores_totales:

                    nombre = f"errores_{archivo.name}"

                    errores_por_formulario = {}

                    for err in errores_totales:
                        formulario = err.get("Formulario", "SIN_FORMULARIO")

                        if formulario not in errores_por_formulario:
                            errores_por_formulario[formulario] = []

                        errores_por_formulario[formulario].append(err)

                    with pd.ExcelWriter(nombre, engine="openpyxl") as writer:

                        for formulario, lista_errores in errores_por_formulario.items():

                            df_formulario = pd.DataFrame(lista_errores)

                            nombre_hoja = str(formulario)[:31]

                            df_formulario.to_excel(
                                writer,
                                sheet_name=nombre_hoja,
                                index=False
                            )

                    with open(nombre, "rb") as f:
                        st.download_button(
                            "📥 Descargue el archivo de errores",
                            f,
                            file_name=nombre
                        )

                    st.error("❌ Archivo con errores")

                else:
                    with open(RUTA_TEMP, "rb") as f:
                        contenido = base64.b64encode(
                            f.read()
                            ).decode()
                        payload = {
                            "nombre": archivo.name,
                            "contenido": contenido
                            }
                        
                        r = requests.post(
                            "https://default7590cbcbf5e34d29a400a282c2a20f.f4.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/25cbc7bf8f27409bbe175b6331af7fcd/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=_nGDSBVGt8krIgP60t6txEGDiOoh9ePkPWJlp4ji2pI",
                            json=payload
                            )
                    if r.ok:
                        st.success("✅ Formularios validados correctamente. En breve recibirá un correo de confirmación.")
                    else:
                        st.error(f"❌ Error OneDrive: {r.status_code}")


        except Exception as e:
            st.error(str(e))

        finally:
            try:
                if os.path.exists(RUTA_TEMP):
                    os.remove(RUTA_TEMP)
            except:
                pass