# config.py - Configuraciones globales de la app

CRM_COLUMNS = [
    "DOC_DNI_RUC",       "NOM_CLI",           "CARTERA",           "COD_CREDITO",
    "NOM_AGENCIA",       "DEUDA_CAP",         "DEUDA_TOTAL",       "TLF_CELULAR_CLIENTE",
    "DIR_CASA",          "DISTRITO",          "RANGO_DIAS_MORA",   "FEC_ULT_PAGO_ACTUAL",
    "NOM_CONYUGE",       "NOM_AVAL",          "TLF_CELULAR_AVAL",  "NOM_CONYUGE_AVAL",
    "DIR_CASA_AVAL",     "DISTRITO_AVAL",     "EXPEDIENTE",        "JUZGADO",
    "CONDICION",         "REFERENCIA",        "PROCESO_JUDICIAL",  "FEC_DEMANDA",
    "MONTO_DEMANDA",     "FEC_INGRESO_JUDICIAL",
]

NUMERIC_COLS = {"DEUDA_CAP", "DEUDA_TOTAL", "MONTO_DEMANDA"}
NONE_LABEL = "— sin mapeo —"
PROFILES_FILE = "mapping_profiles.json"
