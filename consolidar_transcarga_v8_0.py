# =============================================================================
#  CONSOLIDADOR DE ARCHIVOS TRANSCARGA MUNDIAL  —  v7.0
#  Arquivo: consolidar_transcarga_v7.py
#
#  ╔══════════════════════════════════════════════════════════════════╗
#  ║  AUDITORÍA COMPLETA v7.0 — BUGS CORREGIDOS                     ║
#  ╚══════════════════════════════════════════════════════════════════╝
#
#  BUG-1 ▸ IN_FULL siempre era 1 para cualquier guía entregada
#    CAUSA: df["IN_FULL"] = np.where(mask_entregado, 1, 0)
#    Efecto: % In Full == % entregas → métrica inflada artificialmente.
#    FIX: se mantiene IN_FULL=entregado (compatibilidad) y se agrega
#         IN_FULL_STRICT = on_time AND entregado (intersección real).
#         pct_otif_completo corregido: intersección, no producto de ratios.
#
#  BUG-2 ▸ OFF_TIME_START (8h apertura) declarado pero NUNCA USADO
#    CAUSA: ES_OFF_TIME solo chequeaba hora >= 20h; ignoraba < 8h.
#    Efecto: guías creadas a las 2am/5am/7am clasificadas como ON-TIME.
#    FIX: ES_OFF_TIME="SÍ" si hora < OFF_TIME_START OR hora >= OFF_TIME_END
#         Guías con fecha NaT → "SIN HORA" (no contaminan métricas).
#
#  BUG-3 ▸ cond_med_buc filtraba solo ciudad "BUCARAMANGA" puntual
#    CAUSA: dest_n == "BUCARAMANGA" perdía Floridablanca/Girón/Piedecuesta.
#    FIX: dest_n.isin(REGIONAL_BUC) cubre todo el área metropolitana.
#
#  BUG-4 ▸ Denominador OTIF = solo guías entregadas (parcial/optimista)
#    CAUSA: df_ent filtra solo CUMPLE_OTIF in ["SI - ON TIME","NO - LATE"].
#    FIX: dos métricas diferenciadas expuestas en el reporte:
#      · OTIF AJUSTADO (Soft): on_time/(on_time+late)      — operativo
#      · OTIF ÁCIDO    (Hard): on_time_in_full/total_guías — ejecutivo
#
#  BUG-5 ▸ TIEMPO_ENTREGA_NETO en COLUMNAS_ENTERO_TIEMPO pero nunca calculado
#    CAUSA: exportar() formateaba la columna, procesar_otif() no la creaba.
#    FIX: se calcula en procesar_otif() = TIEMPO DE ENTREGA − DIAS_EN_NOVEDAD.
#
#  BUG-6 ▸ pct_otif_completo = (on_time/total)*(in_full/total) — incorrecto
#    CAUSA: multiplica dos ratios en lugar de contar la intersección.
#    FIX: pct_otif_completo = on_time_AND_in_full / total
#
#  NUEVO ▸ OTIF ÁCIDO y OTIF AJUSTADO (estándar logístico internacional)
#    · ÁCIDO    = On Time In Full / Total guías → KPI ejecutivo exigente
#    · AJUSTADO = On Time / (On Time + Late)    → KPI operativo de eficiencia
#  Compatible: Python 3.8+
#
#  CAMBIOS v6.0 (sobre v5.0 funcional):
#  - RUTAS: Migradas a estructura C:\MAESTRO_TRANSCARGA\ (nueva ubicación única)
#    * Archivos fuente: INSUMOS\ARCHIVOS_UNIFICAR\
#    * Referencias:     INSUMOS\ARCHIVOS_REFERENCIA\
#    * Salida:          SALIDA\UNIFICAR\
#  - LECTURA: Ahora lee tanto .xlsx como .xls (mejora cobertura de archivos)
#  - LECTURA: Excluye archivos con prefijo ~ (archivos temporales de Excel)
#  - LIMPIAR: Sufijo "- PRINCIPAL" también eliminado en REGIONAL_DESTINO
#  - DIAS_EN_ESTADO: Usa exclusivamente días hábiles Lun-Sáb (no días corridos)
#  - OTIF: ES_OFF_TIME ahora es "SÍ"/"NO" en lugar de True/False
#  - LOG: El archivo de log se genera siempre, incluso en error fatal
#  - HISTORICO_ESTADOS.csv: Deduplicación mejorada por (GUIA, NOMEST_NORM, fecha)
#  - MEJORA: Verificación de rutas al inicio antes de procesar
#  - MEJORA: consolidar() ahora incluye .xls además de .xlsx
#
#  CAMBIOS v5.0:
#  - AUDITORÍA OTIF: Separación de estados de EXCLUSIÓN (detienen reloj) vs FALLO OPS.
#    * Exclusiones (paran reloj): Dirección errada, Destinatario desconocido, etc.
#    * Fallos OPS (NO paran reloj): Mercancía no ingresa a bodega.
#  - RESTAURACIÓN: Se habilita de nuevo ALERTAS_FRANQUICIAS_BOG (omitido en v4.9).
#  - REPORTE OTIF:
#    * Nueva pestaña "AUDITORIA_NOVEDADES" para trazabilidad de días extra.
#    * Columna "ES_OFF_TIME" para identificar guías creadas tras el corte de 8 PM.
#    * Inclusión de columna "PIEZA" en el detalle.
#  - MEJORA: Formateo automático de porcentajes y KPIs en Dashboard.
#
#  CAMBIOS v4.8:
#  - FIX NOVEDADES_OPS en ALERTAS_CONTROL: la pestaña ahora muestra SOLO el
#    estado "MERCANCIA NO INGRESA A BODEGA" (area OPERACIONES).
#    Se retiraron DEVOLUCION POR NOVEDAD, EN RECLAMACION e INCAUTACION
#    porque son ESTADOS FINALES y nunca deben aparecer en alertas.
#  - FIX OPS_DEMORADAS en ALERTAS_CONTROL: se retiro MERCANCIA NO INGRESA A
#    BODEGA de esta pestaña — ahora va EXCLUSIVAMENTE en NOVEDADES_OPS.
#    Principio: cada estado en una sola pestaña para evitar duplicados.
#  - NUEVO: archivo ALERTAS_FRANQUICIAS_BOG_<ts>.xlsx — seguimiento exclusivo
#    de despachos generados por FRANQUICIAS de REGIONAL_ORIGEN = BOGOTA.
#    Filtro base: LINEA COMERCIAL=="FRANQUICIAS" AND REGIONAL_ORIGEN starts
#    with "BOGOTA" AND TOTAL>0 AND estado no final.
#    Pestanas:
#      RESUMEN        — dashboard con conteos y mini-resumenes de las 3 pestanas
#      ADMITIDOS_DEMO — ADMITIDO + ENVIO PROCESADO >2 dias habiles
#      NOVEDADES_SAC  — estados SAC demorados >1 dia habil
#      OPS_DEMORADAS  — estados OPS, mismos umbrales que ALERTAS_CONTROL
#                       (fijo>2d / Reparto DIRECTO>1d / Reparto REEXP>3d)
#    Columnas: iguales a ALERTAS_CONTROL + NOMOFI (para identificar franquicia)
#    Estados finales excluidos en todo el archivo: ENTREGADO EN SATISFACCION,
#    DEVOLUCION POR NOVEDAD, INCAUTACION, EN RECLAMACION, ANULADO.
#  - Generacion en __main__ como paso 5 (error no fatal igual que los otros).
#
#  CAMBIOS v4.6 (incluidos):
#  - COMERCIAL movida a columna BB (al final) usando patron _serie_comercial
#  - ESTADO OPERATIVO nueva columna BC (CERRADO / NO CERRADO)
#  - Fix cruce COMERCIAL: guardado en variable temporal
#  - Fix slicer TIENE FACTURA en macro
#
#  CAMBIOS v4.5:
#  - NUEVO: Columna T CE CREACION A LEG — días hábiles FECHA_DE_CREACION →
#    FECHA_LEGALIZACION para CE entregadas en satisfacción y legalizadas.
#    Mide el ciclo financiero completo desde la generación de la guía.
#  - NUEVO: Columna T CE ENTREGA A LEG — días hábiles FECREGEST →
#    FECHA_LEGALIZACION para CE entregadas en satisfacción y legalizadas.
#    Mide el tiempo de cobranza desde que se confirmó la entrega.
#  - ELIMINADO: T LEG CONTRAENTREGA — era redundante con T CE ENTREGA A LEG
#  - NUEVO: Pestaña CONTADOS_SIN_LEG en ALERTAS_CONTROL — guías CONTADO
#    sin PRE ni LEGALIZAR con >3 días hábiles desde FECHA_DE_CREACION.
#    Incluye columna DIAS_DESDE_CREACION calculada para este control.
#  - FIX: Todas las alertas filtran solo guías con TOTAL > 0 (mask_con_valor).
#  - Resumen RESUMEN actualizado con fila y mini-tabla para CONTADOS_SIN_LEG.
#
#  CAMBIOS v4.4:
#  - Nombre de archivo fijo (consolidar_transcarga.py) para evitar que se
#    ejecute accidentalmente una versión antigua con bugs
#  - Auto-verificación al inicio: detecta si _norm_global está disponible
#    antes de iniciar el procesamiento y aborta con mensaje claro si no
#  - Todas las correcciones de v4.3 incluidas
#
#  CAMBIOS v4.3 — CORRECCIÓN DE FÓRMULA DE DÍAS + ROBUSTEZ:
#  - FIX CRÍTICO: fórmula días hábiles → busday_count(A+1D, B+1D)
#    incluye el día de llegada, excluye el de inicio
#  - FIX: reporte_diagnostico() se ejecuta SIEMPRE aunque falle generar_alertas
#  - FIX: cada future paralelo captura su error individualmente
#
#  CAMBIOS v4.2:
#  - FIX: _norm_global() compartida — normalización de tildes en todo el script
#  - FIX: MODALIDAD CONTRAENTREGA = CONTRA ENTREGA
#  - FIX: ESTADO_LEGALIZACION LEGALIZADO = LEGALIZACION
#  - FIX: TIPO TRAYECTO se crea antes de calcular tramos
#  - OPT: Exportación maestro + alertas en paralelo
#  - FIX CRÍTICO: Fórmula de días hábiles corregida a busday_count(A+1D, B+1D)
#    → Incluye el día de llegada (B), excluye el día de inicio (A)
#    → Ejemplo correcto: creada lun, legalizada mié = 2 días (no 1)
#    → Aplica a: TIEMPO DE ENTREGA, T LEG CONTADO, T LEG CONTRAENTREGA,
#      DIAS_EN_ESTADO (este ya era correcto pues B=HOY sin +1D)
#  - FIX: reporte_diagnostico() se ejecuta SIEMPRE aunque falle generar_alertas
#    → Se mueve fuera del bloque paralelo, a un try/finally independiente
#  - FIX: Cada future paralelo captura su error individualmente — el fallo
#    de alertas no cancela la exportación del maestro ni el reporte
#  - FIX: Eliminado residuo de _norm local en generar_alertas (fue la causa
#    del NameError en v4.1Alertas que el usuario ejecutó)
#  - OPT: _dias_hab_vec() unificada con parámetro incluir_fin=True/False
#    → incluir_fin=True  → busday_count(A+1D, B+1D) — para tiempos de ciclo
#    → incluir_fin=False → busday_count(A+1D, B)    — para días pendientes
#
#  CAMBIOS v4.2:
#  - FIX: Normalización de tildes aplicada globalmente (_norm_global)
#  - FIX: MODALIDAD — CONTRAENTREGA = CONTRA ENTREGA
#  - FIX: ESTADO_LEGALIZACION — LEGALIZADO = LEGALIZACION
#  - FIX: TIPO TRAYECTO se crea antes de calcular tramos
#  - OPT: Exportación maestro + alertas en paralelo
#  - FIX: Normalización de tildes aplicada en TODAS las comparaciones,
#    no solo en generar_alertas(). NOMEST real lleva tilde ("SATISFACCIÓN").
#  - FIX: MODALIDAD unificada — "CONTRAENTREGA" = "CONTRA ENTREGA" en BD real.
#  - FIX: ESTADO_LEGALIZACION unificado — "LEGALIZADO"="LEGALIZACION" y
#    "PRE-LEGALIZADO" cubre todas las variantes del campo real.
#  - FIX: Orden de ejecución corregido — TIPO TRAYECTO se crea ANTES de
#    calcular T ENTRE DIRECTOS / T ENTRE REEXP (error de scope en v4.1).
#  - FIX: mask_entregado definida una vez y reutilizada en todos los tramos.
#  - FIX: PRE_LEG_VENCIDA ahora detecta las 14,129 guías reales.
#  - FIX: CE_SIN_LEGALIZAR filtro preciso tras corregir MODALIDAD y NOMEST.
#  - OPT: _norm_global() compartida entre agregar_columnas_calculadas()
#    y generar_alertas() — sin duplicación de código.
#  - OPT: _dias_habiles_vec() 100% vectorial con np.busday_count en bloque.
#  - OPT: Masks de alertas construidas con operaciones vectoriales puras.
#  - OPT: Escritura Excel paralela (maestro + alertas simultáneos).
#
#  CAMBIOS v4.1:
#  - Clasificación real de estados desde EXPLICACION_ESTADOS_GUIAS.xlsx
#  - Nueva pestaña ADMITIDOS_DEMO (alerta a regional origen)
#  - EN REPARTO URBANO: umbral diferenciado DIRECTO=1 / REEXP=3 días
#  - TIEMPO DE ENTREGA: días hábiles Lun-Sáb (sin festivos) desde
#    FECHA_DE_CREACION hasta FECREGEST, solo ENTREGADO EN SATISFACCION.
#    El conteo inicia el día siguiente a la creación (creado 3/mar,
#    entregado 5/mar = 2 días). Fechas truncadas a día (sin hora).
#  - T ENTRE DIRECTOS: rangos de tramos para DIRECTO + ENTREGADO EN SAT.
#  - T ENTRE REEXP: rangos de tramos para REEXPEDICION + ENTREGADO EN SAT.
#  - T LEG CONTADO: días hábiles FECHA_DE_CREACION → FECHA_LEGALIZACION,
#    solo MODALIDAD=CONTADO y ESTADO_LEGALIZACION=LEGALIZACION.
#  - T LEG CONTRAENTREGA: días hábiles FECREGEST → FECHA_LEGALIZACION,
#    solo MODALIDAD=CONTRA ENTREGA, NOMEST=ENTREGADO EN SATISFACCION,
#    y ESTADO_LEGALIZACION=LEGALIZACION.
#  - Nuevo archivo ALERTAS_CONTROL_<ts>.xlsx con 7 pestañas:
#      · RESUMEN          — dashboard de conteos por categoría
#      · PRE_LEG_VENCIDA  — PRE LEGALIZADO con >2 días hábiles por modalidad/regional
#      · CE_SIN_LEGALIZAR — CONTRA ENTREGA entregada sin pre ni legalización
#      · DIRECTOS_DEMO    — DIRECTO con días en estado > umbral por estado
#      · REEXP_DEMO       — REEXPEDICION con días en estado > umbral por estado
#      · NOVEDADES_SAC    — estados de novedad SAC con días sin movimiento
#      · NOVEDADES_OPS    — estados de novedad operativa con días sin movimiento
#
#  CAMBIOS v3.1:
#  - Macro VBA y botón removidos: el .bas se carga manualmente en Excel
#  - Script genera únicamente el archivo maestro Excel y CSV
#
#  CAMBIOS v3.0 — OPTIMIZACIÓN DE RENDIMIENTO (concurrencia + vectorización):
#  - Lectura PARALELA de los 4 archivos fuente (ThreadPoolExecutor, 4 workers)
#  - Carga PARALELA de festivos, trayectos y regionales (3 threads simultáneos)
#  - Exportación PARALELA de CSV + Excel (2 threads simultáneos)
#  - MES vectorizado: .dt.year/.dt.month.map() reemplaza apply() fila por fila
#  - DIAS_EN_ESTADO vectorizado: np.busday_count() sobre array completo (~19x más rápido)
#  - Tiempos individuales reportados en consola para cada etapa
#  - Aceleración total estimada: ~3.5x vs v2.0
#
#  CAMBIOS v2.0:
#  - Corrección crítica: fechas con hora ya no generan nulos (dtype mixto)
#  - Filtro automático de registros ANULADOS en NOMEST
#  - Columna MES calculada automáticamente desde FECHA_DE_CREACION
#  - Columna DIAS_EN_ESTADO (días hábiles desde FECREGEST hasta hoy)
#  - Carga de festivos Colombia desde archivo externo
#  - Reporte de diagnóstico ampliado con verificación de integridad
#  - Log detallado por archivo para detectar registros problemáticos
# =============================================================================

import pandas as pd
import numpy as np
import os, sys, time, re
if sys.platform == "win32":
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed

# =============================================================================
#  NORMALIZACIÓN GLOBAL — compartida por TODAS las funciones del script
#  Problema raíz v4.1: comparaciones fallaban por tildes y variantes de texto.
#  Solución: una sola función aplicada siempre antes de cualquier comparación.
# =============================================================================

_TRAD_TILDES = str.maketrans(
    "áéíóúÁÉÍÓÚäëïöüÄËÏÖÜàèìòùÀÈÌÒÙñÑ",
    "aeiouAEIOUaeiouAEIOUaeiouAEIOUnn"
)

def _norm_global(serie: pd.Series) -> pd.Series:
    """
    Normalización robusta para comparaciones de texto:
      - Elimina tildes y diacríticos
      - Convierte a MAYÚSCULAS
      - Elimina espacios extra (str.strip + colapso de espacios internos)
    Usar siempre antes de comparar NOMEST, MODALIDAD, ESTADO_LEGALIZACION,
    TIPO TRAYECTO y cualquier campo de texto que venga de la BD.
    """
    return (
        serie.astype(str)
        .str.strip()
        .str.upper()
        .str.translate(_TRAD_TILDES)
        .str.replace(r"\s+", " ", regex=True)
    )

# Valores canónicos normalizados (resultado de aplicar _norm_global a los
# valores reales de la BD — verificados contra reporte_calidad.csv)
_ENTREGADO_N   = "ENTREGADO EN SATISFACCION"   # BD: "ENTREGADO EN SATISFACCIÓN"
_CONTADO_N     = "CONTADO"
# BD: "CONTRAENTREGA" (una palabra) y ocasionalmente "CONTRA ENTREGA"
_CE_SET_N      = {"CONTRAENTREGA", "CONTRA ENTREGA"}
# BD: "LEGALIZADO" (no "LEGALIZACION")
_LEGALIZADO_N  = {"LEGALIZADO", "LEGALIZACION"}
# BD: "PRE-LEGALIZADO" — con guión; también variantes sin guión
_PRE_LEG_SET_N = {"PRE-LEGALIZADO", "PRE LEGALIZADO", "PRE LEGALIZACION",
                   "PRE-LEGALIZACION", "PRELEGALIZACION"}
_SIN_LEG_N     = "SIN PRE NI LEGALIZACION"

# Estados normalizados por categoría de alerta
_ADMITIDOS_N_SET = {"ADMITIDO", "ENVIO PROCESADO EN OFICINA ORIGEN"}
_OPS_UMBRAL_N    = {                          # estado_normalizado: días_max
    "EN TRANSITO A CIUDAD DESTINO" : 2,
    "INGRESO A BODEGA"             : 2,
    "CARGADA EN VEHICULO"          : 2,
    # MERCANCIA NO INGRESA A BODEGA retirado de aquí — va EXCLUSIVAMENTE en NOVEDADES_OPS
    "INGRESO A BODEGA DESTINO"     : 2,
}
_REPARTO_N       = "EN REPARTO URBANO - DISTRIBUCCION"
_SAC_N_SET       = {"EN NOVEDAD", "CERRADO",
                    "DIRECCION ERRADA - INCOMPLETA",
                    "DESTINATARIO DESCONOCIDO"}
_FINALES_N_SET   = {"ENTREGADO EN SATISFACCION", "DEVOLUCION POR NOVEDAD",
                    "INCAUTACION", "EN RECLAMACION", "ANULADO"}

# =============================================================================
#  AUTO-VERIFICACIÓN DE INTEGRIDAD DEL SCRIPT
#  Detecta si se está ejecutando una versión antigua/corrupta antes de procesar.
#  Si _norm_global no existe, el archivo de alertas fallará con NameError.
#  Este bloque aborta inmediatamente con un mensaje claro en lugar de fallar
#  silenciosamente a mitad del proceso.
# =============================================================================
try:
    _test = _norm_global(pd.Series(["ENTREGADO EN SATISFACCIÓN"]))
    assert _test.iloc[0] == "ENTREGADO EN SATISFACCION", "normalización incorrecta"
    del _test
except Exception as _e_verify:
    print("\n" + "="*65)
    print("  ❌ ERROR DE INTEGRIDAD — SCRIPT INCORRECTO O CORRUPTO")
    print("="*65)
    print(f"\n  Detalle: {_e_verify}")
    print("\n  CAUSA MÁS PROBABLE:")
    print("  Está ejecutando un archivo de versión antigua (ej: v4.1Alertas.py)")
    print("  en lugar del archivo correcto: consolidar_transcarga.py")
    print()
    print("  SOLUCIÓN:")
    print("  1. Verifique que el archivo que ejecuta se llama:")
    print("     consolidar_transcarga.py")
    print("  2. Si tiene archivos con versión en el nombre (v4.1, v4.2, etc.)")
    print("     elimínelos o muévalos a una carpeta de respaldo.")
    print("  3. Use SIEMPRE consolidar_transcarga.py como nombre definitivo.")
    print()
    input("  Presione ENTER para cerrar...")
    raise SystemExit(1)

# =============================================================================
#  CONFIGURACIÓN v6.0 — Rutas unificadas a C:\MAESTRO_TRANSCARGA
#  Estructura requerida:
#    C:\MAESTRO_TRANSCARGA\
#    ├─ INSUMOS\
#    │   ├─ ARCHIVOS_UNIFICAR\      ← Archivos descargados del sistema
#    │   └─ ARCHIVOS_REFERENCIA\    ← Archivos fijos de referencia
#    └─ SALIDA\
#        └─ UNIFICAR\               ← Salida de este pipeline
#
#  Para cambiar rutas, SOLO modifica este bloque.
# =============================================================================

_BASE_DIR          = Path(r"C:\MAESTRO_TRANSCARGA")
_INSUMOS_DIR       = _BASE_DIR / "INSUMOS"
_SALIDA_DIR        = _BASE_DIR / "SALIDA"
_REF_DIR           = _INSUMOS_DIR / "ARCHIVOS_REFERENCIA"

CARPETA_ARCHIVOS   = str(_INSUMOS_DIR / "ARCHIVOS_UNIFICAR")
CARPETA_SALIDA     = str(_SALIDA_DIR  / "UNIFICAR")

ARCHIVO_FESTIVOS    = str(_REF_DIR / "FESTIVOS COLOMBIA.xlsx")
ARCHIVO_TRAYECTOS   = str(_REF_DIR / "TRAYECTOS.xlsx")
ARCHIVO_REGIONALES  = str(_REF_DIR / "REGIONALES.xlsx")
ARCHIVO_CLIENTES    = str(_REF_DIR / "CLIENTES COMERCIALES TCM.xlsx")
ARCHIVO_UNIFICACION = str(_REF_DIR / "MAESTRO_UNIFICACION.xlsx")
ARCHIVO_TARIFARIO   = str(_REF_DIR / "TARIFARIO TCM 2026 VALOR KILO.xlsx")

# Crear carpeta de salida al importar el módulo (no falla si ya existe)
Path(CARPETA_SALIDA).mkdir(parents=True, exist_ok=True)

# Columnas que se ELIMINAN (las 53 que no están en deben_quedar.xlsx)
COLUMNAS_ELIMINAR = [
    "CODPAI", "CODCIU", "CODPAIDES", "NOMPAISDESTINO", "DESTINO_CIUDAD",
    "PESO_VOLUMETRICO", "VALOR_ASEGURADO", "FLETEFULL", "VALOR_PRODUCTO",
    "RECAUDO", "COBRO_ADICIONAL", "OTROS", "TOTALTOTALPESOS", "CARGOPESO",
    "PESOS_BRUTO", "FS", "TRM", "ANEXOS", "PRODUCTO_PADRE", "PRODUCTO_HIJO",
    "LOGUSU", "REFPAG", "CODREG", "PORCENTAJE_DESCUENTO", "TARIFA_PASIVA",
    "DIFERENCIA", "CODEST", "IVA", "ARANCEL", "CREDITO", "EFECTIVO",
    "ENTDOM", "PRCCOMPES", "TOTALIMPUESTOS", "INDREAJ", "OTROS1", "TOTALP",
    "TOTALTOTAL", "INDCOLLECT", "CODPRD", "NOMCLI", "NOMPAI", "CONGUI",
    "ENVIO", "VENDEDOR", "FE", "CORPO", "ESTADO_CUMPLIDO", "FECHA_CUMPLIDO",
    "NUMERO_REMESA_NACIONAL", "NUMERO_REMESA_URBANA", "FECHA_REMESA",
    "NUMERO_APROBACION",
]

# Columnas de fecha — se leen con tipo automático (NO como texto)
# CRÍTICO: esto resuelve el problema del 58.9% de nulos en FECHA_DE_CREACION
COLUMNAS_FECHA = [
    "FECHA_DE_CREACION",
    "FECREGEST",
    "FECHA_CARGUE",
    "FECHA_LEGALIZACION",
    "FECHA_FACTURA",
]

# Columnas de valor monetario — se convierten a entero y se les aplica
# formato $ en el Excel final para que las sumatorias funcionen correctamente
COLUMNAS_MONEDA = [
    "FLETE",
    "VALOR_DECLARADO",
    "SEGURO",
    "TOTAL",
]

# Columnas que se conservan (39 del archivo deben_quedar.xlsx)
COLUMNAS_CONSERVAR = [
    "GUIA", "FECHA_DE_CREACION", "GUIA_INTERNA", "NOMEST", "FECREGEST",
    "REMITENTE_CLIENTE", "DIRECCION_REMITENTE_CLIENTE", "OBSERVACION",
    "CIUDAD_ORIGEN", "CIUDAD_DESTINO", "DEPARTAMENTO", "DESTINATARIO_CLIENTE",
    "DIRECCION_DESTINATARIO_CLIENTE", "NOMBRE_DESTINATARIO", "APELLIDO_DESTINATARIO",
    "PIEZA", "PESOS_FISICO", "OBSERVACION_GUIA", "REGIONAL_DESTINO", "FLETE",
    "VALOR_DECLARADO", "SEGURO", "TOTAL", "CONTENIDO", "REGIONAL_ORIGEN",
    "NOMOFI", "MODALIDAD", "CODOFI", "DIGITADOR", "COD_CLIENTE", "CLIENTE",
    "N_PLANILLA_CARGUE", "CARGUES_URBANO_TRANSBORDO", "FECHA_CARGUE",
    "N_PLANILLA_LEGALIZACION", "ESTADO_LEGALIZACION", "FECHA_LEGALIZACION",
    "NUMERO_FACTURA", "FECHA_FACTURA",
]

# Meses en español con prefijo numérico para ordenar correctamente
# en tablas dinámicas y segmentadores de datos
MESES_ES = {
    1:  "01. Enero",    2:  "02. Febrero",  3:  "03. Marzo",
    4:  "04. Abril",    5:  "05. Mayo",     6:  "06. Junio",
    7:  "07. Julio",    8:  "08. Agosto",   9:  "09. Septiembre",
    10: "10. Octubre",  11: "11. Noviembre", 12: "12. Diciembre",
}

# =============================================================================
#  CLASIFICACIÓN DE ESTADOS — Fuente: EXPLICACION_ESTADOS_GUIAS.xlsx
#  Todos los estados están normalizados en MAYÚSCULAS sin tildes para
#  comparación robusta. La función _norm() aplica esa normalización al df.
# =============================================================================

# ── Estados FINALES — no generan ninguna alerta ───────────────────────────────
# La guía ya terminó su ciclo (bien o mal). No se monitorean.
ESTADOS_FINALES = {
    "ENTREGADO EN SATISFACCION",    # Entregado OK  (NOMEST)
    "DEVOLUCION POR NOVEDAD",       # Devuelto al origen
    "INCAUTACION",                  # Incautado DIAN
    "EN RECLAMACION",               # En proceso de reclamación
    "ANULADO",                      # Anulado (ya filtrado en limpiar(), pero por seguridad)
}

# ── Estados ADMITIDOS — alerta al ORIGEN ─────────────────────────────────────
# Son estados del inicio del proceso. Si superan el umbral, se reporta
# a la regional de ORIGEN para que gestione el despacho.
# Umbral: 2 días hábiles (igual para DIRECTO y REEXPEDICION)
ESTADOS_ADMITIDOS = {
    "ADMITIDO",                          # Estado inicial al crear la guía
    "ENVIO PROCESADO EN OFICINA ORIGEN", # Post-admitido (mismo ciclo de origen)
}
UMBRAL_ADMITIDOS_DIAS = 2

# ── Estados OPS — alerta a OPERACIONES ───────────────────────────────────────
# Guías en movimiento físico que no avanzan. Umbral fijo = 2 días hábiles
# para DIRECTO y REEXPEDICION, EXCEPTO EN REPARTO URBANO que tiene umbral mixto.
ESTADOS_OPS_UMBRAL_FIJO = {
    # Estado normalizado (sin tildes, mayúsculas) : días máximos
    "EN TRANSITO A CIUDAD DESTINO" : 2,   # igual para DIRECTO y REEXP
    "INGRESO A BODEGA"             : 2,
    "CARGADA EN VEHICULO"          : 2,
    # MERCANCIA NO INGRESA A BODEGA va EXCLUSIVAMENTE en NOVEDADES_OPS
    "INGRESO A BODEGA DESTINO"     : 2,
}

# EN REPARTO URBANO - DISTRIBUCCION: único estado con umbral diferente por trayecto
ESTADO_REPARTO            = "EN REPARTO URBANO - DISTRIBUCCION"
UMBRAL_REPARTO_DIRECTO    = 1   # >1 día hábil → alerta
UMBRAL_REPARTO_REEXP      = 3   # >3 días hábiles → alerta

# ── Estados SAC — alerta a Atención al Cliente ───────────────────────────────
# Novedades que requieren gestión del área SAC.
# Umbral: 1 día hábil para todos.
ESTADOS_SAC = {
    "EN NOVEDAD",
    "CERRADO",
    "DIRECCION ERRADA - INCOMPLETA",
    "DESTINATARIO DESCONOCIDO",
}
UMBRAL_SAC_DIAS = 1

# ── PRE LEGALIZACIÓN vencida ─────────────────────────────────────────────────
# Guías en PRE LEGALIZADO con más de 2 días hábiles sin pasar a LEGALIZACION.
# Control de dinero: tesorería / control interno.
UMBRAL_PRE_LEG_DIAS = 2

# Variantes del estado PRE LEGALIZADO tal como pueden aparecer en el sistema
ESTADOS_PRE_LEG = {
    "PRE LEGALIZADO",
    "PRE LEGALIZACION",
    "PRE-LEGALIZACION",
    "PRELEGALIZACION",
    "PRE LEGALIZACION",
}

# ── Estados Novedad para Cálculo OTIF ────────────────────────────────────────
# IMPORTANTE: Según contexto auditoría, se dividen en:
# 1. EXCLUSIONES: Novedades ajenas que DETIENEN el reloj (no castigan al transportador).
# 2. FALLOS OPS: Errores operativos que NO detienen el reloj (afectan el OTIF).

ESTADOS_EXCLUSION_OTIF = {
    "EN NOVEDAD",
    "DIRECCION ERRADA - INCOMPLETA",
    "DESTINATARIO DESCONOCIDO",
    "CERRADO",
    "EN RECLAMACION",
    "INCAUTACION",
    "DEVOLUCION POR NOVEDAD"
}

# MERCANCIA NO INGRESA A BODEGA -> Se considera FALLO OPERATIVO (no excluye).
ESTADOS_FALLO_OPS_OTIF = {
    "MERCANCIA NO INGRESA A BODEGA"
}

# ── Configuración OTIF v4.9 ──────────────────────────────────────────────────
# Ciudades que componen los corredores regionales (normalizadas)
REGIONAL_EJE  = {"PEREIRA", "MANIZALES", "ARMENIA", "DOSQUEBRADAS", "LA VIRGINIA"}
REGIONAL_MED  = {"MEDELLIN", "BELLO", "ITAGUI", "ENVIGADO", "SABANETA", "LA ESTRELLA"}
REGIONAL_CALI = {"CALI", "YUMBO", "PALMIRA", "JAMUNDI"}
REGIONAL_BUC  = {"BUCARAMANGA", "FLORIDABLANCA", "GIRON", "PIEDECUESTA"}

# Conjunto total de ciudades regionales para validar cruces entre ellas
CIUDADES_REGIONALES_SET = REGIONAL_EJE | REGIONAL_MED | REGIONAL_CALI

# Hora de corte (8 AM a 8 PM)
OFF_TIME_START = 8
OFF_TIME_END   = 20

# Regionales Directas TCM (según usuario v4.9)
CIUDADES_DIRECTAS_SET = {
    "BOGOTA", "EJE CAFETERO", "CALI", "MEDELLIN", 
    "BUCARAMANGA", "IBAGUE", "VILLAVICENCIO", "YOPAL"
}

# =============================================================================
#  UTILIDADES
# =============================================================================

# Archivo de log global (se inicializa en __main__)
_LOG_FILE = None

def log(msg: str):
    """Escribe en consola Y en archivo de log simultáneamente."""
    print(msg)
    if _LOG_FILE:
        try:
            _LOG_FILE.write(msg + "\n")
            _LOG_FILE.flush()
        except Exception:
            pass

def sep(titulo=""):
    ancho = 65
    if titulo:
        msg = f"\n{'─'*4} {titulo} {'─'*(ancho - len(titulo) - 6)}"
    else:
        msg = "─" * ancho
    print(msg)
    if _LOG_FILE:
        try:
            _LOG_FILE.write(msg + "\n")
            _LOG_FILE.flush()
        except Exception:
            pass

USAR_CALAMINE = False  # se actualiza en verificar_dependencias()

def verificar_dependencias():
    global USAR_CALAMINE
    sep("VERIFICANDO DEPENDENCIAS")
    faltantes = []
    for lib in ["pandas", "numpy", "openpyxl", "xlsxwriter"]:
        try:
            __import__(lib)
            print(f"  ✅ {lib}")
        except ImportError:
            print(f"  ❌ {lib} — NO INSTALADO")
            faltantes.append(lib)

    # calamine: motor de lectura 5-10x más rápido que openpyxl para archivos grandes
    try:
        import python_calamine  # noqa
        USAR_CALAMINE = True
        print(f"  ✅ python_calamine  ← motor rápido ACTIVO")
    except ImportError:
        USAR_CALAMINE = False
        print(f"  ⚠️  python_calamine NO instalado — usando openpyxl (más lento)")
        print(f"      Para acelerar 5-10x, ejecuta en CMD:")
        print(f"      pip install python-calamine")

    if faltantes:
        print(f"\n  ⚠️  Ejecuta en CMD:  pip install {' '.join(faltantes)}")
        sys.exit(1)

    # Verificar rutas críticas
    sep("VERIFICANDO RUTAS")
    rutas_criticas = {
        "ARCHIVOS_UNIFICAR":    Path(CARPETA_ARCHIVOS),
        "ARCHIVOS_REFERENCIA":  _REF_DIR,
        "SALIDA_UNIFICAR":      Path(CARPETA_SALIDA),
        "FESTIVOS":             Path(ARCHIVO_FESTIVOS),
        "TRAYECTOS":            Path(ARCHIVO_TRAYECTOS),
        "REGIONALES":           Path(ARCHIVO_REGIONALES),
        "CLIENTES":             Path(ARCHIVO_CLIENTES),
    }
    rutas_faltantes = []
    for nombre, ruta in rutas_criticas.items():
        if ruta.exists():
            print(f"  ✅ {nombre}: {ruta}")
        else:
            print(f"  ❌ {nombre}: NO ENCONTRADO → {ruta}")
            if nombre not in ("FESTIVOS",):  # festivos es opcional
                rutas_faltantes.append(nombre)
    if rutas_faltantes:
        print(f"\n  ❌ Faltan carpetas o archivos críticos: {rutas_faltantes}")
        print(f"  Crea la estructura antes de ejecutar. Ver Manual Técnico.")
        # No sys.exit aquí — el usuario puede querer continuar sin archivos de referencia

# =============================================================================
#  CARGA DE FESTIVOS
# =============================================================================

def cargar_festivos(ruta_festivos: str) -> np.ndarray:
    """
    Lee el archivo FESTIVOS COLOMBIA.xlsx.
    Busca fechas en la primera columna de la primera hoja.
    Si el archivo no existe, continúa sin festivos.
    """
    sep("CARGANDO FESTIVOS COLOMBIA")
    ruta = Path(ruta_festivos)

    if not ruta.exists():
        print(f"  ⚠️  Archivo de festivos NO encontrado:")
        print(f"     {ruta}")
        print(f"  ℹ️  Se calcularán días hábiles SIN festivos (solo fines de semana).")
        return np.array([], dtype="datetime64[D]")

    try:
        df_f = pd.read_excel(ruta, header=None, usecols=[0])
        fechas = pd.to_datetime(df_f.iloc[:, 0], errors="coerce", format="mixed", dayfirst=True).dropna()
        festivos = fechas.values.astype("datetime64[D]")
        print(f"  ✅ {len(festivos)} festivos cargados desde: {ruta.name}")
        if len(festivos) > 0:
            print(f"     Rango: {fechas.min().date()} → {fechas.max().date()}")
        return festivos
    except Exception as e:
        print(f"  ❌ Error leyendo festivos: {e}")
        print(f"  ℹ️  Continuando sin festivos.")
        return np.array([], dtype="datetime64[D]")

# =============================================================================
#  CARGA DE TABLAS DE REFERENCIA (TRAYECTOS y REGIONALES)
# =============================================================================

def normalizar_clave(serie: pd.Series) -> pd.Series:
    """
    Normaliza una serie de texto para cruce tolerante:
    - Mayúsculas
    - Sin tildes
    - Sin espacios extra
    Esto evita que 'Bogotá' no cruce con 'BOGOTA' o 'BOGOTÁ'.
    """
    reemplazos = str.maketrans(
        "áéíóúÁÉÍÓÚäëïöüÄËÏÖÜàèìòùÀÈÌÒÙñÑ",
        "aeiouAEIOUaeiouAEIOUaeiouAEIOUnn"
    )
    return (
        serie.astype(str)
        .str.upper()
        .str.strip()
        .str.translate(reemplazos)
        .str.replace(r"\s+", " ", regex=True)
    )

def cargar_trayectos(ruta: str) -> pd.DataFrame:
    """
    Carga el archivo TRAYECTOS.xlsx.
    Estructura confirmada: DESTINO | DES_DEPARTAMENTO | OPERACIÓN
    - DESTINO   → ciudad destino (clave de cruce con maestro)
    - OPERACIÓN → tipo de trayecto (DIRECTO / REEXPEDICION)
    """
    sep("CARGANDO TABLA TRAYECTOS")
    ruta_p = Path(ruta)

    if not ruta_p.exists():
        print(f"  ⚠️  Archivo TRAYECTOS NO encontrado:")
        print(f"     {ruta_p}")
        print(f"  ℹ️  La columna TIPO TRAYECTO se creará vacía.")
        return pd.DataFrame()

    try:
        df_t = pd.read_excel(ruta_p, sheet_name=0, header=0, dtype=str, engine="openpyxl")
        # Normalizar encabezados para manejar tildes/espacios en "OPERACIÓN"
        df_t.columns = df_t.columns.str.strip().str.upper().str.replace(r"\s+", "_", regex=True)
        print(f"  ✅ {len(df_t):,} filas cargadas | Columnas: {list(df_t.columns)}")

        # Columnas confirmadas del archivo real
        COL_CIUDAD = "DESTINO"
        COL_TIPO   = next((c for c in df_t.columns if "OPERAC" in c), None)  # cubre OPERACIÓN y OPERACION

        if COL_CIUDAD not in df_t.columns or not COL_TIPO:
            print(f"  ❌ Columnas esperadas no encontradas.")
            print(f"     Se esperaba: '{COL_CIUDAD}' y una columna con 'OPERAC'")
            print(f"     Columnas disponibles: {list(df_t.columns)}")
            return pd.DataFrame()

        print(f"  📌 Columna ciudad: '{COL_CIUDAD}' | Columna tipo: '{COL_TIPO}'")

        df_t = df_t[[COL_CIUDAD, COL_TIPO]].copy()
        df_t.columns = ["CIUDAD_DESTINO_REF", "TIPO_TRAYECTO_REF"]

        # Clave normalizada para cruce tolerante (sin tildes, mayúsculas, sin espacios extra)
        df_t["_CLAVE_CIUDAD"] = normalizar_clave(df_t["CIUDAD_DESTINO_REF"])
        df_t = df_t[["_CLAVE_CIUDAD", "TIPO_TRAYECTO_REF"]].drop_duplicates(subset="_CLAVE_CIUDAD")

        dist = df_t["TIPO_TRAYECTO_REF"].value_counts().to_dict()
        print(f"  📍 {len(df_t):,} ciudades únicas | Distribución: {dist}")
        return df_t

    except Exception as e:
        print(f"  ❌ Error leyendo TRAYECTOS: {e}")
        return pd.DataFrame()

def cargar_regionales(ruta: str) -> pd.DataFrame:
    """
    Carga el archivo REGIONALES.xlsx.
    Estructura confirmada: CONCATENAR | REGIONAL_DESTINO_CORRECTA
    - CONCATENAR              → CIUDAD_ORIGEN + CIUDAD_DESTINO (sin separador)
    - REGIONAL_DESTINO_CORRECTA → regional correcta a asignar
    """
    sep("CARGANDO TABLA REGIONALES")
    ruta_p = Path(ruta)

    if not ruta_p.exists():
        print(f"  ⚠️  Archivo REGIONALES NO encontrado:")
        print(f"     {ruta_p}")
        print(f"  ℹ️  La columna REG DESTINO CORRECTA se creará vacía.")
        return pd.DataFrame()

    try:
        df_r = pd.read_excel(ruta_p, sheet_name=0, header=0, dtype=str, engine="openpyxl")
        df_r.columns = df_r.columns.str.strip().str.upper().str.replace(r"\s+", "_", regex=True)
        print(f"  ✅ {len(df_r):,} filas cargadas | Columnas: {list(df_r.columns)}")

        # Columnas confirmadas del archivo real
        COL_CONCAT   = "CONCATENAR"
        COL_REGIONAL = "REGIONAL_DESTINO_CORRECTA"

        if COL_CONCAT not in df_r.columns or COL_REGIONAL not in df_r.columns:
            print(f"  ❌ Columnas esperadas no encontradas.")
            print(f"     Se esperaba: '{COL_CONCAT}' y '{COL_REGIONAL}'")
            print(f"     Columnas disponibles: {list(df_r.columns)}")
            return pd.DataFrame()

        df_r = df_r[[COL_CONCAT, COL_REGIONAL]].copy()
        df_r.columns = ["CONCAT_REF", "REGIONAL_DESTINO_CORRECTA_REF"]

        # Clave normalizada para cruce tolerante
        df_r["_CLAVE_CONCAT"] = normalizar_clave(df_r["CONCAT_REF"])
        df_r = df_r[["_CLAVE_CONCAT", "REGIONAL_DESTINO_CORRECTA_REF"]].drop_duplicates(subset="_CLAVE_CONCAT")

        print(f"  📍 {len(df_r):,} combinaciones únicas en tabla de regionales")
        return df_r

    except Exception as e:
        print(f"  ❌ Error leyendo REGIONALES: {e}")
        return pd.DataFrame()

def cargar_clientes(ruta: str) -> pd.DataFrame:
    sep("CARGANDO TABLA COMERCIALES")
    ruta_p = Path(ruta)

    if not ruta_p.exists():
        print(f"  ⚠️  Archivo CLIENTES NO encontrado:\n     {ruta_p}")
        print(f"  ℹ️  La columna COMERCIAL se creará vacía.")
        return pd.DataFrame()

    try:
        df_c = pd.read_excel(ruta_p, sheet_name=0, header=0, dtype=str, engine="openpyxl")
        df_c.columns = df_c.columns.str.strip().str.upper()
        print(f"  ✅ {len(df_c):,} filas cargadas | Columnas: {list(df_c.columns)}")

        if "COD_CLIENTE" not in df_c.columns or "COMERCIALES" not in df_c.columns:
            print(f"  ❌ Columnas esperadas no encontradas en CLIENTES.")
            return pd.DataFrame()

        df_c = df_c[["COD_CLIENTE", "COMERCIALES"]].dropna(subset=["COD_CLIENTE"]).copy()
        df_c["COD_CLIENTE_STR"] = df_c["COD_CLIENTE"].str.replace(r"\.0$", "", regex=True).str.strip()
        df_c = df_c.drop_duplicates(subset="COD_CLIENTE_STR")

        print(f"  📍 {len(df_c):,} comerciales únicos cargados")
        return df_c
    except Exception as e:
        print(f"  ❌ Error leyendo CLIENTES: {e}")
        return pd.DataFrame()

def cargar_unificacion_clientes(ruta: str) -> dict:
    """
    Carga el archivo MAESTRO_UNIFICACION.xlsx y retorna un diccionario
    con el mapeo manual de nombres de clientes.
    """
    sep("CARGANDO MAESTRO DE UNIFICACIÓN")
    ruta_p = Path(ruta)
    if not ruta_p.exists():
        print(f"  ⚠️  Archivo UNIFICACIÓN NO encontrado:\n     {ruta_p}")
        return {}

    try:
        df_u = pd.read_excel(ruta_p, sheet_name=0, header=0, dtype=str, engine="openpyxl")
        df_u.columns = df_u.columns.str.strip().str.upper()
        
        col_sis = "VALOR_SISTEMA"
        col_rep = "VALOR_REPORTE"
        
        if col_sis not in df_u.columns or col_rep not in df_u.columns:
            print(f"  ❌ Columnas {col_sis}/{col_rep} no encontradas.")
            return {}

        mapeo = df_u.dropna(subset=[col_sis, col_rep]).set_index(col_sis)[col_rep].to_dict()
        print(f"  ✅ {len(mapeo):,} reglas de unificación manual cargadas.")
        return mapeo
    except Exception as e:
        print(f"  ❌ Error leyendo UNIFICACIÓN: {e}")
        return {}


def _limpiar_nombre_cliente(nombre: str) -> str:
    """
    Normalización automática de nombres de clientes (fase 1):
    - Elimina sufijos legales (SAS, SA, LTDA, etc.)
    - Elimina puntuación y espacios extras.
    """
    if not isinstance(nombre, str) or pd.isna(nombre): return "SIN CLIENTE"
    
    # Normalización básica de la BD
    n = nombre.strip().upper()
    
    # 1. Eliminar sufijos legales comunes y puntuación al final
    # Usamos regex para bordes de palabra \b
    n = re.sub(r'\b(S\.?A\.?S\.?|S\.?A\.?|LTDA|LIMITADA|S\.? EN C\.?)\b', '', n)
    
    # 2. Eliminar caracteres especiales (excepto & para SHINE & FASHION)
    n = re.sub(r'[^\w\s&ñÑ]', ' ', n)
    
    # 3. Eliminar espacios dobles o triples
    n = " ".join(n.split())
    
    return n if n else "SIN NOMBRE"

# =============================================================================
#  LECTURA Y CONSOLIDACIÓN
# =============================================================================

def leer_archivo(ruta: Path) -> tuple:
    """
    Lee un archivo Excel con soporte de motor rápido (calamine) y fallback a openpyxl.
    - calamine: 5-10x más rápido, ideal para archivos grandes
    - openpyxl: más lento pero con soporte completo de dtype por columna
    - Maneja KeyboardInterrupt con mensaje claro en lugar de traceback
    """
    diag = {
        "archivo": ruta.name, "error": None,
        "registros_brutos": 0, "nulos_fecha_creacion": 0, "anulados": 0,
    }

    dtype_map = {c: str for c in COLUMNAS_CONSERVAR if c not in COLUMNAS_FECHA}

    # Lista de motores a intentar en orden
    motores = []
    if USAR_CALAMINE:
        motores.append("calamine")
    motores.append("openpyxl")

    df = None
    for motor in motores:
        try:
            if motor == "calamine":
                print(f"     Motor: calamine (rápido)...", end="", flush=True)
                df = pd.read_excel(ruta, sheet_name=0, header=0, engine="calamine")
                # Normalizar nombres
                df.columns = (
                    df.columns.str.strip()
                    .str.replace(r"\s+", "_", regex=True)
                    .str.upper()
                )
                # Convertir columnas no-fecha a str
                for col in df.columns:
                    if col not in COLUMNAS_FECHA:
                        df[col] = df[col].astype(str).replace("nan", pd.NA)
                # Convertir fechas explícitamente
                for col in COLUMNAS_FECHA:
                    if col in df.columns and df[col].dtype == object:
                        df[col] = pd.to_datetime(df[col], errors="coerce", dayfirst=True)

            else:  # openpyxl
                aviso = " ← puede tardar 3-8 min por archivo" if not USAR_CALAMINE else ""
                print(f"     Motor: openpyxl{aviso}...", end="", flush=True)
                df = pd.read_excel(
                    ruta, sheet_name=0, header=0,
                    dtype=dtype_map, engine="openpyxl"
                )
                df.columns = (
                    df.columns.str.strip()
                    .str.replace(r"\s+", "_", regex=True)
                    .str.upper()
                )

            print(f" ✅ {len(df):,} filas")
            break  # Éxito — salir del loop

        except KeyboardInterrupt:
            print(f"\n\n  ⛔ PROCESO INTERRUMPIDO MANUALMENTE")
            print(f"  ─────────────────────────────────────────────────")
            if not USAR_CALAMINE:
                print(f"  💡 openpyxl es lento con archivos grandes.")
                print(f"     Instala el motor rápido y vuelve a ejecutar:")
                print(f"     pip install python-calamine")
            else:
                print(f"  El archivo que se estaba leyendo era: {ruta.name}")
            sys.exit(0)

        except Exception as e:
            print(f" ❌ Error con {motor}: {e}")
            if motor != motores[-1]:
                print(f"     Intentando con motor alternativo...")
            else:
                diag["error"] = str(e)
                return pd.DataFrame(), diag

    if df is None or df.empty:
        diag["error"] = "DataFrame vacío tras lectura"
        return pd.DataFrame(), diag

    diag["registros_brutos"] = len(df)

    # Diagnóstico ANTES de limpiar
    if "FECHA_DE_CREACION" in df.columns:
        if df["FECHA_DE_CREACION"].dtype == object:
            df["FECHA_DE_CREACION"] = pd.to_datetime(
                df["FECHA_DE_CREACION"], errors="coerce", dayfirst=True
            )
        diag["nulos_fecha_creacion"] = int(df["FECHA_DE_CREACION"].isna().sum())

    if "NOMEST" in df.columns:
        diag["anulados"] = int(
            df["NOMEST"].astype(str).str.upper().str.strip().eq("ANULADO").sum()
        )

    df.dropna(how="all", inplace=True)
    return df, diag

def _leer_archivo_paralelo(args: tuple) -> tuple:
    """
    Wrapper para leer_archivo() compatible con ThreadPoolExecutor.
    Recibe (indice, ruta) y devuelve (indice, df, diag) para reensamblar
    en orden original después de la lectura paralela.
    """
    idx, ruta = args
    t0 = time.time()
    df, diag = leer_archivo(ruta)
    if not df.empty:
        df.insert(0, "_ARCHIVO_ORIGEN", ruta.name)
    diag["_tiempo_lectura"] = round(time.time() - t0, 1)
    return idx, df, diag

def consolidar(carpeta: Path) -> tuple:
    sep("LEYENDO Y CONSOLIDANDO ARCHIVOS")

    archivos = sorted([
        f for f in carpeta.glob("*.xls*")
        if not f.name.startswith("~")
        and not f.name.lower().startswith("transcarga_maestro")  # excluir salidas anteriores si existieran
    ])

    if not archivos:
        print(f"  ❌ No se encontraron archivos .xlsx en:\n     {carpeta}")
        sys.exit(1)

    n = len(archivos)
    print(f"  📋 Archivos encontrados: {n}")

    # ── Lectura paralela con ThreadPoolExecutor ───────────────────────────────
    # Cada archivo se lee en su propio thread simultáneamente.
    # ThreadPoolExecutor es ideal para I/O-bound (lectura de disco/red OneDrive).
    # workers = min(n, 4) evita saturar la red con demasiados accesos simultáneos.
    # El índice garantiza que el concat final respete el orden original de archivos.
    workers = min(n, 4)
    if workers > 1:
        print(f"  🚀 Lectura PARALELA activada  ({workers} threads simultáneos)\n")
    else:
        print(f"  📂 Lectura secuencial (archivo único)\n")

    from concurrent.futures import ThreadPoolExecutor, as_completed

    resultados = {}   # {idx: (df, diag)}
    t0_total   = time.time()

    with ThreadPoolExecutor(max_workers=workers) as executor:
        futures = {
            executor.submit(_leer_archivo_paralelo, (i, arch)): arch
            for i, arch in enumerate(archivos)
        }
        for future in as_completed(futures):
            arch = futures[future]
            try:
                idx, df, diag = future.result()
                resultados[idx] = (df, diag)

                # Reporte inmediato del archivo que acaba de terminar
                t_arch = diag.get("_tiempo_lectura", 0)
                alertas = ""
                estado  = "✅"
                if diag["error"]:
                    estado  = "❌"
                    alertas = f" | ERROR: {diag['error']}"
                else:
                    if diag["nulos_fecha_creacion"] > 0:
                        pct = diag["nulos_fecha_creacion"] / max(diag["registros_brutos"], 1) * 100
                        alertas += f" | ⚠️ FECHA nulos: {diag['nulos_fecha_creacion']} ({pct:.1f}%)"
                        estado = "⚠️ "
                    if diag["anulados"] > 0:
                        alertas += f" | 🗑️ ANULADOS: {diag['anulados']}"

                print(f"  {estado} {arch.name:<40} "
                      f"{diag['registros_brutos']:>7,} filas | {t_arch:.1f}s{alertas}")
            except Exception as e:
                print(f"  ❌ {arch.name}: error inesperado — {e}")

    t_lectura = time.time() - t0_total
    print(f"\n  ⏱️  Lectura paralela completada en {t_lectura:.1f}s "
          f"(vs ~{t_lectura * workers:.0f}s secuencial estimado)")

    # Reconstruir en orden original por índice
    dfs          = []
    diagnosticos = []
    for idx in sorted(resultados):
        df, diag = resultados[idx]
        if not df.empty:
            dfs.append(df)
        diagnosticos.append(diag)

    if not dfs:
        print("  ❌ Ningún archivo procesado correctamente.")
        sys.exit(1)

    print(f"\n  🔄 Unificando {len(dfs)} archivos...")
    df_total = pd.concat(dfs, ignore_index=True, sort=False)
    print(f"  ✅ Total registros antes de limpieza: {len(df_total):,}")
    return df_total, diagnosticos

# =============================================================================
#  LIMPIEZA Y TRANSFORMACIONES
# =============================================================================

def limpiar(df: pd.DataFrame) -> pd.DataFrame:
    sep("LIMPIEZA Y TRANSFORMACIONES")
    n_inicial = len(df)

    # 1. Eliminar columnas innecesarias
    cols_eliminar = [c for c in COLUMNAS_ELIMINAR if c in df.columns]
    df.drop(columns=cols_eliminar, inplace=True, errors="ignore")
    print(f"  🗑️  Columnas eliminadas: {len(cols_eliminar)}")
    print(f"  📌 Columnas conservadas: {len(df.columns)}")

    # 2. Eliminar registros ANULADOS en NOMEST
    if "NOMEST" in df.columns:
        mask = df["NOMEST"].str.strip().str.upper() == "ANULADO"
        n_anulados = mask.sum()
        df = df[~mask].copy()
        print(f"  🚫 Registros ANULADOS eliminados: {n_anulados:,}")

    # 3. Convertir columnas de fecha primero (evita que su dtype sea 'object' y falle str.strip)
    for col in COLUMNAS_FECHA:
        if col in df.columns and df[col].dtype == object:
            df[col] = pd.to_datetime(df[col], errors="coerce", dayfirst=True)

    # 4. Limpiar espacios en texto de forma más segura
    cols_texto = df.select_dtypes(include=["object", "string"]).columns
    for c in cols_texto:
        try:
            df[c] = df[c].str.strip()
        except AttributeError:
            pass  # Ocurre si la columna tiene objetos datetime u otros tipos no string
    df.replace("", pd.NA, inplace=True)


    # 5. ESTADO_LEGALIZACION — rellenar vacíos con "SIN PRE NI LEGALIZACION"
    if "ESTADO_LEGALIZACION" in df.columns:
        n_vacios = df["ESTADO_LEGALIZACION"].isna().sum()
        df["ESTADO_LEGALIZACION"] = df["ESTADO_LEGALIZACION"].fillna("SIN PRE NI LEGALIZACION")
        dist = df["ESTADO_LEGALIZACION"].value_counts()
        print(f"  📋 ESTADO_LEGALIZACION — vacíos rellenados: {n_vacios:,}")
        for estado, cnt in dist.items():
            print(f"     {estado}: {cnt:,}")

    # 6. MODALIDAD — unificar *CONTRAENTREGA* → CONTRAENTREGA
    if "MODALIDAD" in df.columns:
        # Eliminar asteriscos y espacios extra para normalizar
        antes = df["MODALIDAD"].value_counts().to_dict()
        df["MODALIDAD"] = (
            df["MODALIDAD"]
            .astype(str)
            .str.replace(r"\*", "", regex=True)   # quitar asteriscos
            .str.strip()                            # quitar espacios sobrantes
            .replace("nan", pd.NA)
        )
        despues = df["MODALIDAD"].value_counts().to_dict()
        print(f"  🔁 MODALIDAD — estados unificados:")
        for estado, cnt in despues.items():
            print(f"     {estado}: {cnt:,}")

    # 7. Convertir columnas de moneda a entero (elimina el ".0" y permite sumatorias)
    # Estrategia: to_numeric() interpreta correctamente "9600.0" como 9600.0
    # luego round(0) elimina decimales y astype Int64 convierte a entero limpio
    # Int64 (con mayúscula) acepta NA sin convertirse a float
    for col in COLUMNAS_MONEDA:
        if col in df.columns:
            df[col] = (
                pd.to_numeric(df[col], errors="coerce")
                .round(0)
                .astype("Int64")
            )

    cols_ok  = [c for c in COLUMNAS_MONEDA if c in df.columns]
    if cols_ok:
        print(f"  💰 Columnas de moneda convertidas a entero: {cols_ok}")
        for col in cols_ok:
            total = df[col].sum()
            nulos = df[col].isna().sum()
            print(f"     {col}: suma={total:,.0f}  |  nulos={nulos:,}")

    # 8. Limpiar sufijo "- PRINCIPAL" en columnas de regional
    COLUMNAS_REGIONAL = ["REG DESTINO CORRECTA", "REGIONAL_DESTINO", "REGIONAL_ORIGEN", "REGIONAL_DESTINO_CORRECTA"]
    cols_regional_presentes = [col for col in COLUMNAS_REGIONAL if col in df.columns]

    for col in cols_regional_presentes:
        df[col] = (
            df[col]
            .str.replace(r"\s*-\s*PRINCIPAL$", "", regex=True)  # quita " - PRINCIPAL" y "-PRINCIPAL"
            .str.replace(r"^REG\s+", "", regex=True)              # quita prefijo "REG " si existe
            .str.strip()
            .replace("", pd.NA)
        )

    if cols_regional_presentes:
        print(f"  ✂️  Sufijo '- PRINCIPAL' eliminado en: {cols_regional_presentes}")
        for col in cols_regional_presentes:
            print(f"     {col} — valores únicos resultantes:")
            for v in sorted(df[col].dropna().unique()):
                print(f"       · {v}")

    # 9. Duplicados exactos
    n_antes = len(df)
    df.drop_duplicates(inplace=True)
    dups = n_antes - len(df)
    if dups > 0:
        print(f"  🧹 Duplicados exactos eliminados: {dups:,}")

    print(f"  ✅ Registros finales: {len(df):,}  "
          f"(eliminados en limpieza: {n_inicial - len(df):,})")
    return df

def agregar_columnas_calculadas(df: pd.DataFrame, festivos: np.ndarray,
                                df_trayectos: pd.DataFrame,
                                df_regionales: pd.DataFrame,
                                df_clientes: pd.DataFrame,
                                dict_unificacion: dict) -> pd.DataFrame:
    sep("COLUMNAS CALCULADAS")

    # =========================================================================
    #  PASO 0: Pre-computar series normalizadas UNA SOLA VEZ
    #  Todas las comparaciones de texto usan estas series — nunca el campo raw.
    #  Esto resuelve el problema de tildes (SATISFACCIÓN vs SATISFACCION)
    #  y variantes de MODALIDAD (CONTRAENTREGA vs CONTRA ENTREGA).
    # =========================================================================
    nomest_n  = _norm_global(df["NOMEST"])          if "NOMEST"          in df.columns else pd.Series("", index=df.index)
    modal_n   = _norm_global(df["MODALIDAD"])       if "MODALIDAD"       in df.columns else pd.Series("", index=df.index)
    est_leg_n = _norm_global(df["ESTADO_LEGALIZACION"]) if "ESTADO_LEGALIZACION" in df.columns else pd.Series("", index=df.index)

    # ── Columna MES ───────────────────────────────────────────────────────────
    if "FECHA_DE_CREACION" in df.columns:
        t0_mes = time.time()
        fecha_col = df["FECHA_DE_CREACION"]
        meses_map = pd.Series(MESES_ES)
        anio_str  = fecha_col.dt.year.astype("Int64").astype(str)
        mes_str   = fecha_col.dt.month.map(meses_map)
        mask_ok   = fecha_col.notna()
        mes_serie = pd.Series(pd.NA, index=df.index, dtype=object)
        mes_serie[mask_ok] = anio_str[mask_ok] + " - " + mes_str[mask_ok]
        pos = df.columns.get_loc("FECHA_DE_CREACION") + 1
        df.insert(pos, "MES", mes_serie)
        ok  = df["MES"].notna().sum()
        nul = df["MES"].isna().sum()
        print(f"  📅 MES insertada después de FECHA_DE_CREACION  [{time.time()-t0_mes:.2f}s]")
        print(f"     Con valor: {ok:,}  |  Sin fecha (nulos): {nul:,}")
        for mes_val, cnt in df["MES"].value_counts().sort_index().items():
            print(f"     {mes_val}: {cnt:,} registros")
    else:
        print(f"  ⚠️  FECHA_DE_CREACION no encontrada — MES no creada")

    # ── Columna DIAS_EN_ESTADO ────────────────────────────────────────────────
    if "FECREGEST" in df.columns:
        t0_dias = time.time()
        hoy_np  = np.datetime64(pd.Timestamp.now().date(), "D")
        fec_col = df["FECREGEST"]
        mask_ok = fec_col.notna()
        fechas_np = fec_col[mask_ok].values.astype("datetime64[D]")
        dias_np = np.maximum(
            np.busday_count(fechas_np, hoy_np,
                            weekmask="Mon Tue Wed Thu Fri Sat",
                            holidays=festivos), 0
        ).astype(int)
        resultado = pd.array([pd.NA] * len(df), dtype="Int64")
        resultado[mask_ok.values] = dias_np
        pos = df.columns.get_loc("FECREGEST") + 1
        df.insert(pos, "DIAS_EN_ESTADO", pd.array(resultado, dtype="Int64"))
        validos = df["DIAS_EN_ESTADO"].dropna()
        print(f"  ⏱️  DIAS_EN_ESTADO insertada después de FECREGEST  [{time.time()-t0_dias:.2f}s]")
        if len(validos) > 0:
            print(f"     Promedio: {validos.mean():.0f} días  |  "
                  f"Máximo: {validos.max():.0f} días  |  "
                  f"Sin fecha: {df['DIAS_EN_ESTADO'].isna().sum():,}")
    else:
        print(f"  ⚠️  FECREGEST no encontrada — DIAS_EN_ESTADO no creada")

    # ── Columna ¿TIENE FACTURA? ───────────────────────────────────────────────
    if "FECHA_FACTURA" in df.columns:
        tiene_factura = df["FECHA_FACTURA"].notna().map({True: "SI", False: "NO"})
        pos = df.columns.get_loc("FECHA_FACTURA") + 1
        df.insert(pos, "¿TIENE FACTURA?", tiene_factura)
        dist = df["¿TIENE FACTURA?"].value_counts()
        print(f"  🧾 ¿TIENE FACTURA? insertada después de FECHA_FACTURA")
        for val, cnt in dist.items():
            print(f"     {val}: {cnt:,}  ({cnt/len(df)*100:.1f}%)")
    else:
        print(f"  ⚠️  FECHA_FACTURA no encontrada — ¿TIENE FACTURA? no creada")

    # =========================================================================
    #  COLUMNAS DE TIEMPOS — después de ¿TIENE FACTURA?
    #  ORDEN CRÍTICO: TIPO TRAYECTO aún no existe en df aquí.
    #  Se calcula primero el cruce con TRAYECTOS (más abajo), y los tramos
    #  T ENTRE DIRECTOS / T ENTRE REEXP se insertan DESPUÉS de ese cruce.
    #  Para ello calculamos TIEMPO DE ENTREGA ahora y los tramos al final.
    # =========================================================================
    sep("COLUMNAS DE TIEMPOS DE ENTREGA Y LEGALIZACIÓN")

    tiene_fecha_crea = "FECHA_DE_CREACION"   in df.columns
    tiene_fecregest  = "FECREGEST"           in df.columns
    tiene_modalidad  = "MODALIDAD"           in df.columns
    tiene_est_leg    = "ESTADO_LEGALIZACION" in df.columns
    tiene_fecha_leg  = "FECHA_LEGALIZACION"  in df.columns

    # Posición de inserción: justo después de ¿TIENE FACTURA?
    pos_base = (
        df.columns.get_loc("¿TIENE FACTURA?") + 1
        if "¿TIENE FACTURA?" in df.columns
        else len(df.columns)
    )

    def _dias_hab_vec(fecha_ini: pd.Series, fecha_fin: pd.Series,
                      incluir_fin: bool = True) -> tuple:
        """
        Días hábiles Lun-Sáb sin festivos entre dos columnas de fecha.
        100% vectorial — un solo llamado np.busday_count para todo el df.

        Convención verificada (v4.3):
          incluir_fin=True  → busday_count(ini+1D, fin+1D)
            El día de inicio NO cuenta, el día de llegada SÍ cuenta.
            Ejemplo: creada lun 2/mar, legalizada mié 4/mar = 2 días ✅
            Uso: TIEMPO DE ENTREGA, T LEG CONTADO, T CE CREACION A LEG, T CE ENTREGA A LEG

          incluir_fin=False → busday_count(ini+1D, fin)
            Ni inicio ni fin cuentan (solo días intermedios).
            No usado actualmente, disponible para métricas futuras.

        Devuelve (array_int32, mask_bool):
          - array_int32: días calculados (-1 donde alguna fecha es nula)
          - mask_bool:   True donde ambas fechas son válidas y calculadas
        """
        ini = fecha_ini.dt.normalize()
        fin = fecha_fin.dt.normalize()
        mask = ini.notna() & fin.notna()

        ini_np = ini[mask].values.astype("datetime64[D]") + np.timedelta64(1, "D")
        fin_np = fin[mask].values.astype("datetime64[D]")
        if incluir_fin:
            fin_np = fin_np + np.timedelta64(1, "D")

        dias = np.maximum(
            np.busday_count(ini_np, fin_np,
                            weekmask="Mon Tue Wed Thu Fri Sat",
                            holidays=festivos), 0
        ).astype(np.int32)

        out = np.full(len(df), -1, dtype=np.int32)
        out[mask.values] = dias
        return out, mask.values

    # ── TIEMPO DE ENTREGA ─────────────────────────────────────────────────────
    # FIX v4.2: usa nomest_n (normalizado) → captura "ENTREGADO EN SATISFACCIÓN"
    if tiene_fecha_crea and tiene_fecregest:
        t0 = time.time()
        mask_ent = (nomest_n == _ENTREGADO_N)          # vectorial, sin tildes

        dias_arr, mask_valido = _dias_hab_vec(
            df["FECHA_DE_CREACION"], df["FECREGEST"], incluir_fin=True
        )
        # Solo asignar donde es ENTREGADO y ambas fechas válidas
        mask_te = mask_ent.values & mask_valido
        resultado_te = pd.array([pd.NA] * len(df), dtype="Int64")
        resultado_te[mask_te] = dias_arr[mask_te]

        df.insert(pos_base, "TIEMPO DE ENTREGA", pd.array(resultado_te, dtype="Int64"))
        pos_base += 1

        n_ent = int(mask_ent.sum())
        n_cal = int(mask_te.sum())
        prom  = int(dias_arr[mask_te].mean()) if n_cal > 0 else 0
        print(f"  🕐 TIEMPO DE ENTREGA insertada  [{time.time()-t0:.2f}s]")
        print(f"     Guías ENTREGADO EN SATISFACCIÓN: {n_ent:,}")
        print(f"     Con tiempo calculado: {n_cal:,}  |  Promedio: {prom} días hábiles")
        if n_ent > n_cal:
            print(f"     ⚠️  {n_ent - n_cal:,} sin calcular (FECREGEST o FECHA_DE_CREACION nula)")
    else:
        df.insert(pos_base, "TIEMPO DE ENTREGA", pd.NA)
        pos_base += 1
        print(f"  ⚠️  TIEMPO DE ENTREGA — faltan FECHA_DE_CREACION o FECREGEST")

    # ── T LEG CONTADO ─────────────────────────────────────────────────────────
    # FIX v4.2: MODALIDAD normalizada ("CONTADO") y ESTADO_LEGALIZACION
    #   normalizado — BD usa "LEGALIZADO", no "LEGALIZACION"
    if tiene_fecha_crea and tiene_fecha_leg and tiene_modalidad and tiene_est_leg:
        t0 = time.time()
        mask_lc = (modal_n == _CONTADO_N) & est_leg_n.isin(_LEGALIZADO_N)

        dias_arr_lc, mask_val_lc = _dias_hab_vec(
            df["FECHA_DE_CREACION"], df["FECHA_LEGALIZACION"], incluir_fin=True
        )
        mask_lc_ok = mask_lc.values & mask_val_lc
        serie_lc = pd.array([pd.NA] * len(df), dtype="Int64")
        serie_lc[mask_lc_ok] = dias_arr_lc[mask_lc_ok]

        df.insert(pos_base, "T LEG CONTADO", pd.array(serie_lc, dtype="Int64"))
        pos_base += 1

        n_val = int(mask_lc_ok.sum())
        prom  = int(dias_arr_lc[mask_lc_ok].mean()) if n_val > 0 else 0
        print(f"  💵 T LEG CONTADO insertada  [{time.time()-t0:.2f}s]")
        print(f"     CONTADO legalizados: {int(mask_lc.sum()):,}  |  "
              f"Con cálculo: {n_val:,}  |  Promedio: {prom} días")
    else:
        df.insert(pos_base, "T LEG CONTADO", pd.NA)
        pos_base += 1
        print(f"  ⚠️  T LEG CONTADO — faltan columnas fuente")

    # ── T CE CREACION A LEG — días hábiles FECHA_DE_CREACION → FECHA_LEGALIZACION ──
    # CE + ENTREGADO EN SATISFACCION + LEGALIZADO
    # Mide el ciclo financiero completo: desde que nació la guía hasta que se legalizó.
    if tiene_fecha_crea and tiene_fecha_leg and tiene_modalidad and tiene_est_leg:
        t0 = time.time()
        mask_ce_crea = (
            modal_n.isin(_CE_SET_N)
            & (nomest_n == _ENTREGADO_N)
            & est_leg_n.isin(_LEGALIZADO_N)
        )
        dias_arr_cec, mask_val_cec = _dias_hab_vec(
            df["FECHA_DE_CREACION"], df["FECHA_LEGALIZACION"], incluir_fin=True
        )
        mask_cec_ok = mask_ce_crea.values & mask_val_cec
        serie_cec = pd.array([pd.NA] * len(df), dtype="Int64")
        serie_cec[mask_cec_ok] = dias_arr_cec[mask_cec_ok]

        df.insert(pos_base, "T CE CREACION A LEG", pd.array(serie_cec, dtype="Int64"))
        pos_base += 1

        n_val = int(mask_cec_ok.sum())
        prom  = int(dias_arr_cec[mask_cec_ok].mean()) if n_val > 0 else 0
        print(f"  📅 T CE CREACION A LEG insertada  [{time.time()-t0:.2f}s]")
        print(f"     CE entregadas y legalizadas: {int(mask_ce_crea.sum()):,}  |  "
              f"Con cálculo: {n_val:,}  |  Promedio: {prom} días")
    else:
        df.insert(pos_base, "T CE CREACION A LEG", pd.NA)
        pos_base += 1
        print(f"  ⚠️  T CE CREACION A LEG — faltan columnas fuente")

    # ── T CE ENTREGA A LEG — días hábiles FECREGEST → FECHA_LEGALIZACION ─────────
    # CE + ENTREGADO EN SATISFACCION + LEGALIZADO
    # Mide cuántos días tardó la legalización desde que se confirmó la entrega.
    if tiene_fecregest and tiene_fecha_leg and tiene_modalidad and tiene_est_leg:
        t0 = time.time()
        mask_ce_ent = (
            modal_n.isin(_CE_SET_N)
            & (nomest_n == _ENTREGADO_N)
            & est_leg_n.isin(_LEGALIZADO_N)
        )
        dias_arr_cee, mask_val_cee = _dias_hab_vec(
            df["FECREGEST"], df["FECHA_LEGALIZACION"], incluir_fin=True
        )
        mask_cee_ok = mask_ce_ent.values & mask_val_cee
        serie_cee = pd.array([pd.NA] * len(df), dtype="Int64")
        serie_cee[mask_cee_ok] = dias_arr_cee[mask_cee_ok]

        df.insert(pos_base, "T CE ENTREGA A LEG", pd.array(serie_cee, dtype="Int64"))
        pos_base += 1

        n_val = int(mask_cee_ok.sum())
        prom  = int(dias_arr_cee[mask_cee_ok].mean()) if n_val > 0 else 0
        print(f"  🚚 T CE ENTREGA A LEG insertada  [{time.time()-t0:.2f}s]")
        print(f"     CE entregadas y legalizadas: {int(mask_ce_ent.sum()):,}  |  "
              f"Con cálculo: {n_val:,}  |  Promedio: {prom} días")
    else:
        df.insert(pos_base, "T CE ENTREGA A LEG", pd.NA)
        pos_base += 1
        print(f"  ⚠️  T CE ENTREGA A LEG — faltan columnas fuente")

    # ── Columna LINEA COMERCIAL ───────────────────────────────────────────────
    if "NOMOFI" in df.columns:
        nomofi_upper = df["NOMOFI"].fillna("").str.upper()
        condiciones = [
            nomofi_upper.str.contains("CONEXPRESS", na=False),
            nomofi_upper.str.contains("FRANQUI",    na=False),
        ]
        df["LINEA COMERCIAL"] = np.select(condiciones, ["CONEXPRESS", "FRANQUICIAS"],
                                           default="VENTA PROPIA")
        pos = df.columns.get_loc("NOMOFI") + 1
        cols = [c for c in df.columns if c != "LINEA COMERCIAL"]
        cols.insert(pos, "LINEA COMERCIAL")
        df = df[cols]
        dist = df["LINEA COMERCIAL"].value_counts()
        print(f"  🏷️  LINEA COMERCIAL insertada después de NOMOFI")
        for linea, cnt in dist.items():
            print(f"     {linea}: {cnt:,}  ({cnt/len(df)*100:.1f}%)")
    else:
        print(f"  ⚠️  NOMOFI no encontrada — LINEA COMERCIAL no creada")

    # ── Cruces CONCAT / TIPO TRAYECTO / REG DESTINO CORRECTA ─────────────────
    # CRÍTICO: TIPO TRAYECTO se crea aquí. Los tramos T ENTRE DIRECTOS y
    # T ENTRE REEXP se insertan DESPUÉS de este bloque (ver más abajo).
    sep("CRUCES CON TABLAS DE REFERENCIA")

    if "CIUDAD_ORIGEN" in df.columns and "CIUDAD_DESTINO" in df.columns:

        df["CONCAT"] = (df["CIUDAD_ORIGEN"].fillna("").str.strip()
                        + df["CIUDAD_DESTINO"].fillna("").str.strip())
        print(f"  🔗 CONCAT creada  (ejemplo: {df['CONCAT'].iloc[0] if len(df) > 0 else 'N/A'})")

        if not df_trayectos.empty:
            dict_tray = df_trayectos.set_index("_CLAVE_CIUDAD")["TIPO_TRAYECTO_REF"].to_dict()
            clave_dest = _norm_global(df["CIUDAD_DESTINO"])
            # Reusar normalizar_clave si existe; _norm_global es equivalente
            df["TIPO TRAYECTO"] = clave_dest.map(dict_tray)
            ok  = df["TIPO TRAYECTO"].notna().sum()
            nok = df["TIPO TRAYECTO"].isna().sum()
            print(f"  🗺️  TIPO TRAYECTO — cruzados: {ok:,}  |  sin match: {nok:,}")
            if nok > 0:
                for ciudad, cnt in (df.loc[df["TIPO TRAYECTO"].isna(), "CIUDAD_DESTINO"]
                                    .value_counts().head(10).items()):
                    print(f"       → '{ciudad}': {cnt:,} registros")
        else:
            df["TIPO TRAYECTO"] = pd.NA
            print(f"  ⚠️  TIPO TRAYECTO vacía (archivo TRAYECTOS no disponible)")

        if not df_regionales.empty:
            dict_reg = df_regionales.set_index("_CLAVE_CONCAT")["REGIONAL_DESTINO_CORRECTA_REF"].to_dict()
            clave_concat = _norm_global(df["CONCAT"])
            df["REG DESTINO CORRECTA"] = (
                clave_concat.map(dict_reg)
                .str.replace(r"\s*-\s*PRINCIPAL$", "", regex=True)
                .str.replace(r"^REG\s+", "", regex=True)
                .str.strip()
            )
            ok  = df["REG DESTINO CORRECTA"].notna().sum()
            nok = df["REG DESTINO CORRECTA"].isna().sum()
            print(f"  🏢 REG DESTINO CORRECTA — cruzados: {ok:,}  |  sin match: {nok:,}")
        else:
            df["REG DESTINO CORRECTA"] = pd.NA
            print(f"  ⚠️  REG DESTINO CORRECTA vacía (archivo REGIONALES no disponible)")

        # ── COMERCIAL — cruce COD_CLIENTE → guardado en serie temporal ──────
        # IMPORTANTE: el resultado se guarda en _serie_comercial y NO se asigna
        # todavía a df. Si se asignara aquí y luego se hiciera df = df[cols],
        # la columna se perdería porque cols no la incluye (va al final).
        # Se aplica al DataFrame en el bloque COLUMNAS FINALES, al cierre.
        _serie_comercial = None
        if not df_clientes.empty:
            dict_cli = df_clientes.set_index("COD_CLIENTE_STR")["COMERCIALES"].to_dict()
            if "COD_CLIENTE" in df.columns:
                clave_cli = df["COD_CLIENTE"].astype(str).str.replace(r"\.0$", "", regex=True).str.strip()
                _serie_comercial = clave_cli.map(dict_cli).fillna("SIN COMERCIAL")
                ok  = (_serie_comercial != "SIN COMERCIAL").sum()
                nok = (_serie_comercial == "SIN COMERCIAL").sum()
                print(f"  🧑‍💼 COMERCIAL — cruzados: {ok:,}  |  sin match: {nok:,}  (se ubica al final)")
            else:
                print("  ⚠️  COD_CLIENTE no encontrada — COMERCIAL se creará vacía al final")
        else:
            print(f"  ⚠️  Archivo CLIENTES no disponible — COMERCIAL se creará vacía al final")

        # ── UNIFICACIÓN DE CLIENTES ──────────────────────────────────────────
        if "CLIENTE" in df.columns:
            t0_unif = time.time()
            def _unificar(val):
                if val in dict_unificacion:
                    return dict_unificacion[val]
                return _limpiar_nombre_cliente(val)
            df["CLIENTE"] = df["CLIENTE"].apply(_unificar)
            print(f"  ✨ CLIENTE unificado (manual + auto) [{time.time()-t0_unif:.2f}s]")

        # Reordenar: CONCAT, TIPO TRAYECTO, REG DESTINO CORRECTA → después de DEPARTAMENTO
        # COMERCIAL no se incluye aquí — se inserta al final en bloque COLUMNAS FINALES
        cols = list(df.columns)
        nuevas_cols = ["CONCAT", "TIPO TRAYECTO", "REG DESTINO CORRECTA"]
        for c in nuevas_cols:
            if c in cols: cols.remove(c)
        if "DEPARTAMENTO" in cols:
            pos_dep = cols.index("DEPARTAMENTO") + 1
            for i, c in enumerate(nuevas_cols):
                cols.insert(pos_dep + i, c)
        df = df[cols]
        print(f"  ✅ Columnas CONCAT/TIPO TRAYECTO/REG DESTINO CORRECTA reordenadas después de DEPARTAMENTO")
    else:
        _serie_comercial = None
        print(f"  ⚠️  CIUDAD_ORIGEN o CIUDAD_DESTINO no encontradas — columnas no creadas")

    # =========================================================================
    #  TRAMOS DE TIEMPO — ahora que TIPO TRAYECTO ya existe en df
    #  FIX v4.2: este bloque se mueve DESPUÉS del cruce con TRAYECTOS
    # =========================================================================
    sep("TRAMOS DE TIEMPO POR TIPO DE TRAYECTO")

    tiene_tipo_tray = "TIPO TRAYECTO" in df.columns
    tiene_tiempo_e  = "TIEMPO DE ENTREGA" in df.columns

    # Re-computar mask_ent con el df actual (por si cambió el índice)
    mask_ent_final = (_norm_global(df["NOMEST"]) == _ENTREGADO_N
                      if "NOMEST" in df.columns
                      else pd.Series(False, index=df.index))

    tray_n_final = (_norm_global(df["TIPO TRAYECTO"])
                    if tiene_tipo_tray
                    else pd.Series("", index=df.index))

    # Posición de inserción: después de TIEMPO DE ENTREGA
    if tiene_tiempo_e:
        pos_tramos = df.columns.get_loc("TIEMPO DE ENTREGA") + 1
    else:
        pos_tramos = len(df.columns)

    # ── T ENTRE DIRECTOS ──────────────────────────────────────────────────────
    def _tramo_directo(d: int) -> str:
        if d <= 1:  return "1 DIA"
        if d == 2:  return "2 DIAS"
        if d == 3:  return "3 DIAS"
        if d == 4:  return "4 DIAS"
        if d <= 6:  return "5 A 6 DIAS"
        if d <= 8:  return "7 A 8 DIAS"
        if d <= 10: return "9 A 10 DIAS"
        return "MAS DE 10 DIAS"

    if tiene_tipo_tray and tiene_tiempo_e:
        t0 = time.time()
        mask_dir = mask_ent_final & (tray_n_final == "DIRECTO")
        te_col   = df["TIEMPO DE ENTREGA"]

        # Vectorizado con np.select sobre los valores enteros
        te_int = pd.to_numeric(te_col, errors="coerce").fillna(-1).astype(int)
        tramos_dir = np.select(
            [te_int <= 1, te_int == 2, te_int == 3, te_int == 4,
             (te_int >= 5) & (te_int <= 6), (te_int >= 7) & (te_int <= 8),
             (te_int >= 9) & (te_int <= 10), te_int > 10],
            ["1 DIA","2 DIAS","3 DIAS","4 DIAS",
             "5 A 6 DIAS","7 A 8 DIAS","9 A 10 DIAS","MAS DE 10 DIAS"],
            default=None
        )
        serie_dir = pd.Series(pd.NA, index=df.index, dtype=object)
        serie_dir[mask_dir.values] = tramos_dir[mask_dir.values]

        df.insert(pos_tramos, "T ENTRE DIRECTOS", serie_dir)
        pos_tramos += 1

        dist = serie_dir.dropna().value_counts()
        print(f"  📦 T ENTRE DIRECTOS insertada  [{time.time()-t0:.2f}s]  "
              f"({int(mask_dir.sum()):,} guías DIRECTO)")
        orden = ["1 DIA","2 DIAS","3 DIAS","4 DIAS",
                 "5 A 6 DIAS","7 A 8 DIAS","9 A 10 DIAS","MAS DE 10 DIAS"]
        for tramo in orden:
            if tramo in dist.index:
                print(f"     {tramo}: {dist[tramo]:,}")
    else:
        df.insert(pos_tramos, "T ENTRE DIRECTOS", pd.NA)
        pos_tramos += 1
        print(f"  ⚠️  T ENTRE DIRECTOS — TIPO TRAYECTO o TIEMPO DE ENTREGA no disponibles")

    # ── T ENTRE REEXP ─────────────────────────────────────────────────────────
    # Rangos: 0-2→"1 A 2 DIAS", 3-4→"3 A 4 DIAS", 5→"5 A 6 DIAS",
    #         6-7→"6 A 7 DIAS", 8-10→"8 A 10 DIAS", >10→"MAS DE 10 DIAS"
    if tiene_tipo_tray and tiene_tiempo_e:
        t0 = time.time()
        mask_reex = (
            mask_ent_final
            & tray_n_final.isin(["REEXPEDICION", "REEXPEDICION"])  # normalizado ya sin tilde
        )
        tramos_re = np.select(
            [te_int <= 2,
             (te_int >= 3) & (te_int <= 4),
             te_int == 5,
             (te_int >= 6) & (te_int <= 7),
             (te_int >= 8) & (te_int <= 10),
             te_int > 10],
            ["1 A 2 DIAS","3 A 4 DIAS","5 A 6 DIAS",
             "6 A 7 DIAS","8 A 10 DIAS","MAS DE 10 DIAS"],
            default=None
        )
        serie_re = pd.Series(pd.NA, index=df.index, dtype=object)
        serie_re[mask_reex.values] = tramos_re[mask_reex.values]

        df.insert(pos_tramos, "T ENTRE REEXP", serie_re)

        dist_re = serie_re.dropna().value_counts()
        print(f"  🔄 T ENTRE REEXP insertada  [{time.time()-t0:.2f}s]  "
              f"({int(mask_reex.sum()):,} guías REEXPEDICION)")
        orden_re = ["1 A 2 DIAS","3 A 4 DIAS","5 A 6 DIAS",
                    "6 A 7 DIAS","8 A 10 DIAS","MAS DE 10 DIAS"]
        for tramo in orden_re:
            if tramo in dist_re.index:
                print(f"     {tramo}: {dist_re[tramo]:,}")
    else:
        df.insert(pos_tramos, "T ENTRE REEXP", pd.NA)
        print(f"  ⚠️  T ENTRE REEXP — TIPO TRAYECTO o TIEMPO DE ENTREGA no disponibles")

    # =========================================================================
    #  COLUMNAS FINALES — COMERCIAL y ESTADO OPERATIVO (columnas BB y BC)
    #  COMERCIAL se inserta aquí (al final) para no interferir con la macro
    #  VBA que construye la tabla dinámica desde la hoja DATOS.
    #  ESTADO OPERATIVO clasifica cada guía según si su ciclo está cerrado.
    # =========================================================================
    sep("COLUMNAS FINALES — COMERCIAL y ESTADO OPERATIVO")

    # ── COMERCIAL al final ────────────────────────────────────────────────────
    # Aplica la serie calculada en el bloque de cruces (_serie_comercial).
    # Se inserta aquí para que df = df[cols] del reordenamiento no la elimine.
    if _serie_comercial is not None:
        df.insert(len(df.columns), "COMERCIAL", _serie_comercial)
        ok  = (_serie_comercial != "SIN COMERCIAL").sum()
        nok = (_serie_comercial == "SIN COMERCIAL").sum()
        print(f"  🧑‍💼 COMERCIAL insertada al final (col BB)")
        print(f"     Con comercial asignado: {ok:,}  |  Sin comercial: {nok:,}")
    else:
        df["COMERCIAL"] = "SIN COMERCIAL"
        print(f"  ⚠️  COMERCIAL creada vacía al final (cruce no disponible)")

    # ── ESTADO OPERATIVO (última columna, BC) ─────────────────────────────────
    # Estados FINALES: el ciclo de la guía está cerrado (entregada, devuelta,
    # incautada o en reclamación). Cualquier otro estado = NO CERRADO.
    # Nota: ANULADO se elimina en limpiar() y nunca llega aquí.
    _ESTADOS_CERRADOS_N = {
        "ENTREGADO EN SATISFACCION",   # BD: "ENTREGADO EN SATISFACCIÓN"
        "DEVOLUCION POR NOVEDAD",
        "INCAUTACION",
        "EN RECLAMACION",
    }

    if "NOMEST" in df.columns:
        t0_eo = time.time()
        nomest_eo = _norm_global(df["NOMEST"])
        estado_operativo = np.where(
            nomest_eo.isin(_ESTADOS_CERRADOS_N),
            "CERRADO",
            "NO CERRADO"
        )
        df["ESTADO OPERATIVO"] = estado_operativo

        n_cerrado    = int((df["ESTADO OPERATIVO"] == "CERRADO").sum())
        n_no_cerrado = int((df["ESTADO OPERATIVO"] == "NO CERRADO").sum())
        print(f"  🔴 ESTADO OPERATIVO creada (col BC)  [{time.time()-t0_eo:.2f}s]")
        print(f"     CERRADO:    {n_cerrado:,}  ({n_cerrado/len(df)*100:.1f}%)")
        print(f"     NO CERRADO: {n_no_cerrado:,}  ({n_no_cerrado/len(df)*100:.1f}%)")
    else:
        df["ESTADO OPERATIVO"] = pd.NA
        print(f"  ⚠️  NOMEST no encontrada — ESTADO OPERATIVO creada vacía")

    return df

# =============================================================================
#  CARGA PARALELA DE ARCHIVOS DE REFERENCIA
# =============================================================================

def _cargar_referencias_paralelo(
    ruta_festivos: str, ruta_trayectos: str, ruta_regionales: str, ruta_clientes: str,
    ruta_unificacion: str
) -> tuple:
    """
    Carga festivos, trayectos, regionales, clientes y unificación en paralelo.
    """
    from concurrent.futures import ThreadPoolExecutor

    sep("CARGANDO TABLAS DE REFERENCIA (paralelo)")
    print("  🚀 Cargando festivos, trayectos, regionales, clientes y unificación simultáneamente...")
    t0 = time.time()

    resultados = {}
    tareas = [
        ("festivos",      cargar_festivos,           ruta_festivos),
        ("trayectos",     cargar_trayectos,          ruta_trayectos),
        ("regionales",    cargar_regionales,         ruta_regionales),
        ("clientes",      cargar_clientes,           ruta_clientes),
        ("unificacion",   cargar_unificacion_clientes, ruta_unificacion),
        ("malla",         cargar_malla_tarifario,    ARCHIVO_TARIFARIO),
    ]

    with ThreadPoolExecutor(max_workers=6) as executor:
        futures = {
            executor.submit(fn, ruta): nombre
            for nombre, fn, ruta in tareas
        }
        for future in futures:
            nombre = futures[future]
            try:
                resultados[nombre] = future.result()
            except Exception as e:
                print(f"  ❌ Error cargando {nombre}: {e}")
                # Devolver valores seguros en caso de error
                if nombre == "festivos":
                    resultados[nombre] = np.array([], dtype="datetime64[D]")
                elif nombre == "unificacion":
                    resultados[nombre] = {}
                else:
                    resultados[nombre] = pd.DataFrame()

    print(f"  ✅ 6 tablas cargadas en {time.time()-t0:.1f}s (paralelo)")
    return (
        resultados["festivos"], resultados["trayectos"], 
        resultados["regionales"], resultados["clientes"],
        resultados["unificacion"], resultados.get("malla", {})
    )

def cargar_malla_tarifario(ruta: str) -> dict:
    """
    Carga la malla de cubrimiento (Directo/Reexpedido) desde el Tarifario 2026.
    Retorna un diccionario {(Origen, Destino): 'DIRECTO'/'REEXPEDIDO'}
    """
    sep("CARGANDO MALLA DE CUBRIMIENTO (TARIFARIO)")
    ruta_p = Path(ruta)
    if not ruta_p.exists():
        print(f"  ⚠️  Archivo TARIFARIO NO encontrado: {ruta_p}")
        return {}

    try:
        malla = {}
        # Cargamos el Excel
        xl = pd.ExcelFile(ruta_p)
        
        for tipo in ["DIRECTO", "REEXPEDIDO"]:
            if tipo in xl.sheet_names:
                df = pd.read_excel(xl, sheet_name=tipo, index_col=0)
                # Normalizar columnas (Origines) e índice (Destinos)
                df.columns = _norm_global(pd.Series(df.columns))
                df.index = _norm_global(pd.Series(df.index))
                
                # Iterar sobre la matriz para capturar la cobertura
                for ori in df.columns:
                    # Filtramos solo los destinos que tienen precio (cobertura activa)
                    destinos_cobertura = df[ori].dropna().index
                    for dest in destinos_cobertura:
                        malla[(ori, dest)] = tipo
        
        print(f"  ✅ Malla cargada: {len(malla):,} relaciones origen-destino.")
        return malla
    except Exception as e:
        print(f"  ❌ Error cargando malla: {e}")
        return {}



# =============================================================================
#  GENERACIÓN DEL ARCHIVO DE ALERTAS Y CONTROL
# =============================================================================

def generar_alertas(df: pd.DataFrame, festivos: np.ndarray, carpeta_salida: Path) -> Path:
    """
    Genera ALERTAS_CONTROL_<ts>.xlsx con pestañas de seguimiento operativo,
    control de dinero y control de contados (Gisell).

    PESTAÑAS:
      RESUMEN               — dashboard general de conteos
      CONTADOS_VP_PEND      — Venta Propia pendientes (todos los estados, TOTAL>0)
      CONTADOS_PRE_PLANILLA — PRE LEGALIZADO contado sin cerrar, por regional + planilla
      CONTADOS_VP_LEG       — Venta Propia ya legalizadas (historial/conciliación)
      CONTADOS_FRANQ        — Franquicias: todos sus contados (pendientes + en proceso)
      PRE_LEG_VENCIDA       — PRE LEGALIZADO vencido >2d (todas las modalidades)
      CE_SIN_LEGALIZAR      — Contra entrega entregada sin legalizar
      ADMITIDOS_DEMO        — Admitidos demorados >2d (alerta ORIGEN)
      OPS_DEMORADAS         — Operaciones demoradas por estado y trayecto
      NOVEDADES_SAC         — Novedades SAC demoradas >1d
      NOVEDADES_OPS         — Novedades operativas demoradas
    """
    sep("GENERANDO ARCHIVO DE ALERTAS Y CONTROL")
    ts      = pd.Timestamp.now().strftime("%Y%m%d_%H%M")
    hoy_str = pd.Timestamp.now().strftime("%d/%m/%Y %H:%M")
    ruta_out = carpeta_salida / f"ALERTAS_CONTROL_{ts}.xlsx"

    def _col(nombre):
        return df[nombre] if nombre in df.columns else pd.Series(pd.NA, index=df.index)

    # ── Series normalizadas ───────────────────────────────────────────────────
    nomest_n    = _norm_global(_col("NOMEST"))
    modal_n     = _norm_global(_col("MODALIDAD"))
    est_leg_n   = _norm_global(_col("ESTADO_LEGALIZACION"))
    tray_n      = _norm_global(_col("TIPO TRAYECTO"))
    linea_com_n = _norm_global(_col("LINEA COMERCIAL"))
    reg_ori     = _col("REGIONAL_ORIGEN").fillna("SIN REGIONAL")
    reg_ori_n   = _norm_global(reg_ori)
    reg_dest    = _col("REG DESTINO CORRECTA").fillna("SIN REGIONAL")
    dias_est    = pd.to_numeric(_col("DIAS_EN_ESTADO"), errors="coerce")

    # ── Filtro base: solo guías con TOTAL > 0 ────────────────────────────────
    total_num      = pd.to_numeric(_col("TOTAL"), errors="coerce").fillna(0)
    mask_con_valor = total_num > 0

    # ── Alias de constantes globales ─────────────────────────────────────────
    _ADMITIDOS_N = _ADMITIDOS_N_SET
    _SAC_N       = _SAC_N_SET
    _OPS_FIJO_N  = _OPS_UMBRAL_N
    _PRE_LEG_N   = _PRE_LEG_SET_N
    _LEG_OK_N    = _PRE_LEG_SET_N | _LEGALIZADO_N | {"LEGALIZADO", "LEGALIZACION"}

    # ── Columnas de detalle para hojas operativas ─────────────────────────────
    COLS_BASE = [
        "GUIA", "FECHA_DE_CREACION", "MES", "NOMEST", "FECREGEST", "DIAS_EN_ESTADO",
        "MODALIDAD", "TIPO TRAYECTO", "CIUDAD_ORIGEN", "CIUDAD_DESTINO",
        "REGIONAL_ORIGEN", "REG DESTINO CORRECTA",
        "ESTADO_LEGALIZACION", "FECHA_LEGALIZACION",
        "TIEMPO DE ENTREGA", "T LEG CONTADO", "T CE CREACION A LEG", "T CE ENTREGA A LEG",
        "REMITENTE_CLIENTE", "DESTINATARIO_CLIENTE", "TOTAL",
        "OBSERVACION", "OBSERVACION_GUIA",
    ]
    COLS_DETALLE = [c for c in COLS_BASE if c in df.columns]

    # ── Columnas para pestañas de CONTADOS (Gisell) ───────────────────────────
    # Incluye NOMEST para ver el estado de la guía en todos los casos
    COLS_CONTADOS = [c for c in [
        "GUIA", "FECHA_DE_CREACION", "MES",
        "REGIONAL_ORIGEN", "CODOFI", "NOMOFI", "CLIENTE", "LINEA COMERCIAL",
        "NOMEST", "ESTADO_LEGALIZACION", "N_PLANILLA_LEGALIZACION",
        "FECHA_LEGALIZACION", "T LEG CONTADO",
        "TOTAL",
    ] if c in df.columns]

    # =========================================================================
    #  CÁLCULO DE DÍAS HÁBILES DESDE CREACIÓN Y DESDE FECREGEST
    # =========================================================================
    _hoy_np = np.datetime64(pd.Timestamp.now().date(), "D")

    def _calcular_dias_hab(fecha_col_nombre: str) -> pd.Series:
        """Días hábiles Lun-Sáb (sin festivos) desde fecha_col hasta hoy."""
        if fecha_col_nombre not in df.columns:
            return pd.Series(-1, index=df.index, dtype=np.int32)
        col = df[fecha_col_nombre]
        mask_ok = col.notna()
        fechas_np = col[mask_ok].values.astype("datetime64[D]")
        resultado = np.full(len(df), -1, dtype=np.int32)
        resultado[mask_ok.values] = np.maximum(
            np.busday_count(
                fechas_np + np.timedelta64(1, "D"),
                _hoy_np + np.timedelta64(1, "D"),
                weekmask="Mon Tue Wed Thu Fri Sat",
                holidays=festivos,
            ), 0
        ).astype(np.int32)
        return pd.Series(resultado, index=df.index)

    _dias_crea_serie      = _calcular_dias_hab("FECHA_DE_CREACION")
    _dias_fecregest_serie = _calcular_dias_hab("FECREGEST")

    # =========================================================================
    #  MASKS OPERATIVAS
    # =========================================================================

    # ── ADMITIDOS demorados ───────────────────────────────────────────────────
    mask_adm = (
        nomest_n.isin(_ADMITIDOS_N)
        & (dias_est > UMBRAL_ADMITIDOS_DIAS)
        & mask_con_valor
    )

    # ── OPS demorados — umbral fijo ───────────────────────────────────────────
    mask_ops_fijo = pd.Series(False, index=df.index)
    for estado_n, umbral in _OPS_FIJO_N.items():
        mask_ops_fijo |= (nomest_n == estado_n) & (dias_est > umbral) & mask_con_valor

    # ── OPS demorados — EN REPARTO URBANO (umbral por trayecto) ──────────────
    mask_reparto_dir  = (
        (nomest_n == _REPARTO_N) & (tray_n == "DIRECTO")
        & (dias_est > UMBRAL_REPARTO_DIRECTO) & mask_con_valor
    )
    mask_reparto_reex = (
        (nomest_n == _REPARTO_N) & (tray_n == "REEXPEDICION")
        & (dias_est > UMBRAL_REPARTO_REEXP) & mask_con_valor
    )
    mask_ops_reparto = mask_reparto_dir | mask_reparto_reex
    mask_ops_total   = mask_ops_fijo | mask_ops_reparto

    # ── SAC demorados ─────────────────────────────────────────────────────────
    mask_sac = (
        nomest_n.isin(_SAC_N)
        & (dias_est > UMBRAL_SAC_DIAS)
        & mask_con_valor
    )

    # ── NOVEDADES OPS — solo MERCANCIA NO INGRESA A BODEGA ───────────────────
    # DEVOLUCION POR NOVEDAD, EN RECLAMACION, INCAUTACION son ESTADOS FINALES
    # y nunca deben aparecer en ninguna alerta.
    _NOV_OPS_N_SET = {
        "MERCANCIA NO INGRESA A BODEGA",
    }
    UMBRAL_NOV_OPS_DIAS = 2
    mask_nov_ops = (
        nomest_n.isin(_NOV_OPS_N_SET)
        & (dias_est > UMBRAL_NOV_OPS_DIAS)
        & mask_con_valor
    )

    # ── PRE LEGALIZADO vencido (control de dinero) ────────────────────────────
    mask_pre = (
        est_leg_n.isin(_PRE_LEG_N)
        & modal_n.isin({_CONTADO_N} | _CE_SET_N)
        & (dias_est > UMBRAL_PRE_LEG_DIAS)
        & mask_con_valor
    )

    # ── CE entregada SIN legalizar (control de dinero) ────────────────────────
    mask_ce_sin = (
        modal_n.isin(_CE_SET_N)
        & (nomest_n == _ENTREGADO_N)
        & (~est_leg_n.isin(_LEG_OK_N))
        & mask_con_valor
    )

    # =========================================================================
    #  MASKS CONTADOS — Control Gisell
    #  Base: MODALIDAD=CONTADO + TOTAL>0 + excluir ANULADO
    #  Umbral diferenciado: BOGOTA = 1 día hábil / resto = 2 días hábiles
    # =========================================================================
    _BOGOTA_N      = "BOGOTA"
    mask_es_bogota = reg_ori_n.str.startswith(_BOGOTA_N)
    _umbral_cnt    = mask_es_bogota.map({True: 1, False: 2})

    # Base contado con valor, excluyendo ANULADO
    mask_contado_base = (
        (modal_n == _CONTADO_N)
        & mask_con_valor
        & (nomest_n != "ANULADO")
    )

    # Sin ninguna gestión de legalización
    mask_contado_sin_leg_base = (
        mask_contado_base
        & (~est_leg_n.isin(_PRE_LEG_SET_N | _LEGALIZADO_N))
    )

    # En PRE LEGALIZADO
    mask_contado_pre = mask_contado_base & est_leg_n.isin(_PRE_LEG_SET_N)

    # Ya legalizado
    mask_contado_leg = mask_contado_base & est_leg_n.isin(_LEGALIZADO_N)

    # Sin leg que superan umbral diferenciado (para CONTADOS_SIN_LEG legacy)
    mask_contado_sin_leg = (
        mask_contado_sin_leg_base
        & (_dias_crea_serie > _umbral_cnt)
    )

    # PENDIENTES = sin leg + pre sin cerrar (universo completo de Gisell)
    # CONTADOS_VP_PEND: TODOS los estados de la guía (no solo entregados)
    mask_contado_pendiente = mask_contado_sin_leg_base | mask_contado_pre

    # Separar por línea comercial
    mask_vp    = linea_com_n == "VENTA PROPIA"
    mask_franq = linea_com_n == "FRANQUICIAS"

    mask_vp_pend    = mask_contado_pendiente & mask_vp
    mask_vp_leg     = mask_contado_leg       & mask_vp
    mask_franq_pend = mask_contado_pendiente & mask_franq

    # =========================================================================
    #  DataFrames DE DETALLE
    # =========================================================================

    def _detalle(mask, sort_col="DIAS_EN_ESTADO"):
        df_h = df.loc[mask, COLS_DETALLE].copy()
        if not df_h.empty and sort_col in df_h.columns:
            df_h = df_h.sort_values(sort_col, ascending=False)
        return df_h

    def _detalle_contados(mask, sort_cols=None, dias_col=True):
        """
        Detalle para pestañas de contados con columnas específicas de Gisell.
        DIAS_CONTROL: desde FECREGEST si está en PRE, desde FECHA_DE_CREACION si sin leg.
        """
        df_h = df.loc[mask, COLS_CONTADOS].copy()
        if dias_col and not df_h.empty:
            est_h  = est_leg_n[mask]
            es_pre = est_h.isin(_PRE_LEG_SET_N)
            dias_control = np.where(
                es_pre.values,
                _dias_fecregest_serie[mask].values,
                _dias_crea_serie[mask].values
            )
            df_h.insert(len(df_h.columns), "DIAS_CONTROL",
                        pd.array(dias_control, dtype="Int32"))
            origen_dias = np.where(es_pre.values, "Desde PRE (FECREGEST)", "Desde creación")
            df_h.insert(len(df_h.columns), "ORIGEN_DIAS",
                        pd.array(origen_dias, dtype=object))
        if not df_h.empty:
            if sort_cols:
                valid = [c for c in sort_cols if c in df_h.columns]
                if valid:
                    df_h = df_h.sort_values(valid, ascending=True)
            elif dias_col and "DIAS_CONTROL" in df_h.columns:
                df_h = df_h.sort_values("DIAS_CONTROL", ascending=False)
        return df_h

    df_adm     = _detalle(mask_adm)
    df_ops     = _detalle(mask_ops_total)
    df_sac     = _detalle(mask_sac)
    df_nov_ops = _detalle(mask_nov_ops)
    df_pre     = _detalle(mask_pre)
    df_ce_sin  = _detalle(mask_ce_sin)

    # Contados — legacy (sin leg superando umbral)
    df_cnt_sin = _detalle_contados(
        mask_contado_sin_leg,
        sort_cols=["REGIONAL_ORIGEN", "DIAS_CONTROL"]
    )

    # Contados VP pendientes — TODOS los estados (sin leg + pre)
    df_vp_pend = _detalle_contados(
        mask_vp_pend,
        sort_cols=["REGIONAL_ORIGEN", "N_PLANILLA_LEGALIZACION", "FECHA_DE_CREACION"]
    )

    # Contados VP legalizados — historial/conciliación
    df_vp_leg = _detalle_contados(
        mask_vp_leg,
        sort_cols=["REGIONAL_ORIGEN", "N_PLANILLA_LEGALIZACION", "FECHA_LEGALIZACION"],
        dias_col=False
    )

    # Contados PRE por planilla — PRE LEGALIZADO sin cerrar, agrupado
    df_pre_planilla = _detalle_contados(
        mask_contado_pre,
        sort_cols=["REGIONAL_ORIGEN", "N_PLANILLA_LEGALIZACION", "FECHA_DE_CREACION"]
    )

    # Contados Franquicias — pendientes (sin leg + pre), todos los estados
    df_franq_pend = _detalle_contados(
        mask_franq_pend,
        sort_cols=["REGIONAL_ORIGEN", "NOMOFI", "DIAS_CONTROL"]
    )

    # =========================================================================
    #  RESÚMENES INTERMEDIOS
    # =========================================================================

    # PRE_LEG: por modalidad + regional destino
    if not df_pre.empty:
        resumen_pre = (
            pd.DataFrame({"MODALIDAD": modal_n[mask_pre].values,
                          "REGIONAL":  reg_dest[mask_pre].values})
            .value_counts().reset_index(name="CANTIDAD")
        )
        resumen_pre["UMBRAL_DIAS"] = UMBRAL_PRE_LEG_DIAS
    else:
        resumen_pre = pd.DataFrame(columns=["MODALIDAD", "REGIONAL", "CANTIDAD", "UMBRAL_DIAS"])

    # CE_SIN: por regional destino
    if not df_ce_sin.empty:
        resumen_ce = (reg_dest[mask_ce_sin].value_counts()
                      .reset_index().rename(columns={"REG DESTINO CORRECTA": "REGIONAL", 0: "CANTIDAD"}))
        resumen_ce.columns = ["REGIONAL", "CANTIDAD"]
    else:
        resumen_ce = pd.DataFrame(columns=["REGIONAL", "CANTIDAD"])

    # CONTADOS_SIN_LEG: por regional origen
    if not df_cnt_sin.empty:
        resumen_cnt = (reg_ori[mask_contado_sin_leg].value_counts()
                       .reset_index().rename(columns={"REGIONAL_ORIGEN": "REGIONAL_ORIGEN", 0: "CANTIDAD"}))
        resumen_cnt.columns = ["REGIONAL_ORIGEN", "CANTIDAD"]
    else:
        resumen_cnt = pd.DataFrame(columns=["REGIONAL_ORIGEN", "CANTIDAD"])

    # ADMITIDOS: por regional origen
    if not df_adm.empty:
        resumen_adm = (reg_ori[mask_adm].value_counts()
                       .reset_index().rename(columns={"REGIONAL_ORIGEN": "REGIONAL_ORIGEN", 0: "CANTIDAD"}))
        resumen_adm.columns = ["REGIONAL_ORIGEN", "CANTIDAD"]
    else:
        resumen_adm = pd.DataFrame(columns=["REGIONAL_ORIGEN", "CANTIDAD"])

    # OPS: por estado + regional destino
    if not df_ops.empty:
        resumen_ops = (
            pd.DataFrame({"ESTADO":   _col("NOMEST")[mask_ops_total].values,
                          "REGIONAL": reg_dest[mask_ops_total].values})
            .value_counts().reset_index(name="CANTIDAD")
        )
    else:
        resumen_ops = pd.DataFrame(columns=["ESTADO", "REGIONAL", "CANTIDAD"])

    # VP_PEND: resumen por regional + planilla con subtotal $
    if not df_vp_pend.empty:
        _grp_cols_vp = [c for c in ["REGIONAL_ORIGEN", "N_PLANILLA_LEGALIZACION"]
                        if c in df_vp_pend.columns]
        _grp_vp = df_vp_pend.groupby(_grp_cols_vp, dropna=False).agg(
            GUIAS=("GUIA", "count") if "GUIA" in df_vp_pend.columns else ("REGIONAL_ORIGEN", "count"),
            TOTAL_PESOS=("TOTAL", "sum") if "TOTAL" in df_vp_pend.columns else ("REGIONAL_ORIGEN", "count"),
        ).reset_index()
    else:
        _grp_vp = pd.DataFrame()

    # FRANQ: resumen por regional + franquicia con subtotal $
    if not df_franq_pend.empty:
        _grp_cols_franq = [c for c in ["REGIONAL_ORIGEN", "NOMOFI"] if c in df_franq_pend.columns]
        _grp_franq = df_franq_pend.groupby(_grp_cols_franq, dropna=False).agg(
            GUIAS=("GUIA", "count") if "GUIA" in df_franq_pend.columns else ("REGIONAL_ORIGEN", "count"),
            TOTAL_PESOS=("TOTAL", "sum") if "TOTAL" in df_franq_pend.columns else ("REGIONAL_ORIGEN", "count"),
        ).reset_index()
    else:
        _grp_franq = pd.DataFrame()

    # ── RESUMEN GENERAL ───────────────────────────────────────────────────────
    _n_ent     = int((nomest_n == _ENTREGADO_N).sum())
    _n_leg     = int(est_leg_n.isin(_LEGALIZADO_N).sum())
    _n_pre_tot = int(est_leg_n.isin(_PRE_LEG_N).sum())
    _n_sin     = int((est_leg_n == _SIN_LEG_N).sum())

    resumen_rows = [
        ("💵 CONTROL CONTADOS — GISELL",                           "",              ""),
        ("VP Pendientes — todos los estados, sin legalizar",
         "VENTA PROPIA",   int(mask_vp_pend.sum())),
        ("VP Legalizadas (historial/conciliación)",
         "VENTA PROPIA",   int(mask_vp_leg.sum())),
        ("Franquicias pendientes (sin legalizar)",
         "FRANQUICIAS",    int(mask_franq_pend.sum())),
        ("PRE LEGALIZADO sin cerrar — CONTADO",
         "CONTADO",        int(mask_contado_pre.sum())),
        ("CONTADO sin pre ni leg — umbral BOG=1d / resto=2d",
         "CONTADO",        int(mask_contado_sin_leg.sum())),
        ("", "", ""),
        ("💰 CONTROL DE DINERO",                                   "",              ""),
        ("PRE LEGALIZADO vencido >2 días  —  CONTADO",
         "CONTADO",        int((mask_pre & (modal_n == _CONTADO_N)).sum())),
        ("PRE LEGALIZADO vencido >2 días  —  CONTRA ENTREGA",
         "CONTRA ENTREGA", int((mask_pre & modal_n.isin(_CE_SET_N)).sum())),
        ("CE entregada SIN pre ni legalizar",
         "CONTRA ENTREGA", int(mask_ce_sin.sum())),
        ("", "", ""),
        ("🚨 ALERTAS OPERACIONES",                                 "",              ""),
        ("ADMITIDOS demorados >2 días (→ alerta ORIGEN)",
         "ADMITIDOS",      int(mask_adm.sum())),
        ("OPS demoradas — umbral fijo >2 días",
         "OPS",            int(mask_ops_fijo.sum())),
        ("EN REPARTO URBANO — DIRECTO >1 día",
         "OPS-DIRECTO",    int(mask_reparto_dir.sum())),
        ("EN REPARTO URBANO — REEXPEDICION >3 días",
         "OPS-REEXP",      int(mask_reparto_reex.sum())),
        ("", "", ""),
        ("⚠️ NOVEDADES",                                           "",              ""),
        ("Novedades SAC demoradas >1 día hábil",
         "SAC",            int(mask_sac.sum())),
        ("Novedades operativas demoradas >2 días",
         "OPS",            int(mask_nov_ops.sum())),
        ("", "", ""),
        ("📊 TOTALES MAESTRO",                                     "",              ""),
        ("Total guías en el maestro",    "TODOS", len(df)),
        ("Entregadas en satisfacción",   "TODOS", _n_ent),
        ("Con LEGALIZACION confirmada",  "TODOS", _n_leg),
        ("En PRE LEGALIZADO",            "TODOS", _n_pre_tot),
        ("Sin pre ni legalización",      "TODOS", _n_sin),
    ]
    df_resumen = pd.DataFrame(resumen_rows,
                              columns=["CATEGORÍA", "TIPO / MODALIDAD", "CANTIDAD"])

    # =========================================================================
    #  ESCRITURA EXCEL
    # =========================================================================
    print(f"  📊 Escribiendo archivo de alertas...")
    t0_xl = time.time()

    with pd.ExcelWriter(ruta_out, engine="xlsxwriter") as writer:
        wb = writer.book

        # ── Formatos base ─────────────────────────────────────────────────────
        _f = lambda **kw: wb.add_format({"font_name": "Calibri", "font_size": 10, **kw})
        fmt_enc     = _f(bold=True, bg_color="#1F3864", font_color="#FFFFFF",
                         align="center", valign="vcenter")
        fmt_cel     = _f()
        fmt_num     = _f(num_format="#,##0", align="center")
        fmt_cop     = _f(num_format="$#,##0")
        fmt_rojo    = _f(bold=True, bg_color="#C00000", font_color="#FFFFFF", align="center")
        fmt_naran   = _f(bold=True, bg_color="#FF9900", font_color="#000000", align="center")
        fmt_verde   = _f(bg_color="#E2EFDA")
        fmt_secc    = _f(bold=True, bg_color="#D6E4F0")
        fmt_secc2   = _f(bold=True, bg_color="#FFF2CC")   # sección contados
        fmt_cop_tot = _f(bold=True, num_format="$#,##0", bg_color="#D6E4F0")

        # ── Helper: hoja operativa estándar ───────────────────────────────────
        def _hoja(nombre: str, df_h: pd.DataFrame, titulo: str,
                  color: str = "#1F3864",
                  resumen_extra: pd.DataFrame = None,
                  resumen_titulo: str = ""):
            if df_h.empty:
                ws = wb.add_worksheet(nombre)
                ws.write(0, 0, f"✅ Sin alertas en esta categoría al {hoy_str}",
                         _f(bold=True, font_color="#375623", font_size=11))
                return

            df_h.to_excel(writer, sheet_name=nombre, index=False, startrow=1)
            ws = writer.sheets[nombre]

            fmt_tit = _f(bold=True, font_size=12, bg_color=color, font_color="#FFFFFF",
                         align="left", valign="vcenter")
            ws.merge_range(0, 0, 0, len(df_h.columns) - 1,
                f"{titulo}  ·  {hoy_str}  ·  {len(df_h):,} guías", fmt_tit)
            ws.set_row(0, 20)

            for i, cn in enumerate(df_h.columns):
                ws.write(1, i, cn, fmt_enc)
                try:
                    ql = df_h[cn].astype(str).str.len().quantile(0.9)
                    ancho = min(int(max(len(cn), ql)) + 3, 42)
                except Exception:
                    ancho = 18
                ws.set_column(i, i, ancho, fmt_cel)

            ws.freeze_panes(2, 0)
            ws.autofilter(1, 0, len(df_h) + 1, len(df_h.columns) - 1)

            if "DIAS_EN_ESTADO" in df_h.columns:
                ci = list(df_h.columns).index("DIAS_EN_ESTADO")
                for ri, val in enumerate(df_h["DIAS_EN_ESTADO"]):
                    try:
                        d = int(val)
                        fmt_d = fmt_rojo if d > 10 else (fmt_naran if d > 5 else fmt_verde)
                        ws.write(ri + 2, ci, d, fmt_d)
                    except Exception:
                        pass

            if resumen_extra is not None and not resumen_extra.empty:
                fila_res = len(df_h) + 4
                ws.write(fila_res, 0, resumen_titulo,
                         _f(bold=True, bg_color="#2E75B6", font_color="#FFFFFF"))
                for ci, c in enumerate(resumen_extra.columns):
                    ws.write(fila_res + 1, ci, c, fmt_enc)
                for ri, row in resumen_extra.iterrows():
                    for ci, val in enumerate(row):
                        ws.write(fila_res + 2 + ri, ci,
                                 int(val) if isinstance(val, (int, float))
                                     and ci == len(resumen_extra.columns) - 1
                                     else val,
                                 fmt_num if ci == len(resumen_extra.columns) - 1
                                     else fmt_cel)

        # ── Helper: hoja resumen financiero (ej. comercial) ────────────────────────
        def _hoja_resumen(nombre: str, df_h: pd.DataFrame, titulo: str, color: str):
            if df_h.empty: return
            try:
                df_h.to_excel(writer, sheet_name=nombre, index=False, startrow=1)
                ws = writer.sheets[nombre]
                fmt_tit = _f(bold=True, font_size=12, bg_color=color, font_color="#FFFFFF", align="left", valign="vcenter")
                ws.merge_range(0, 0, 0, len(df_h.columns) - 1, f"{titulo}  ·  {hoy_str}", fmt_tit)
                ws.set_row(0, 20)
                for i, cn in enumerate(df_h.columns):
                    ws.write(1, i, cn, fmt_enc)
                    ancho = max(len(str(cn)), 15) + 5
                    is_numeric = pd.api.types.is_numeric_dtype(df_h[cn])
                    if cn in ("GUIAS", "CANTIDAD", "TOTAL_PESOS") or is_numeric:
                        ws.set_column(i, i, ancho, fmt_cop)
                    else:
                        ws.set_column(i, i, max(ancho, 25), fmt_cel)
                ws.freeze_panes(2, 0)
                ws.autofilter(1, 0, len(df_h) + 1, len(df_h.columns) - 1)
            except Exception as _e:
                print(f"  ❌ [{nombre}] falló: {_e}")

        # ── Helper: hoja de contados (Gisell) ─────────────────────────────────
        def _hoja_contados(nombre: str, df_h: pd.DataFrame, titulo: str,
                           color: str, grp_df: pd.DataFrame = None,
                           grp_titulo: str = ""):
            """
            Hoja optimizada para control de contados.
            - Convierte Int32/Int64 nullable a float antes de to_excel (compatibilidad xlsxwriter)
            - Vectoriza semáforo ESTADO_LEGALIZACION (sin _norm_global por fila)
            - Todos los bloques de escritura dentro de try/except individuales
            """
            if df_h.empty:
                try:
                    ws = wb.add_worksheet(nombre)
                    ws.write(0, 0, f"✅ Sin registros al {hoy_str}",
                             _f(bold=True, font_color="#375623", font_size=11))
                except Exception as _e:
                    print(f"  ⚠️  [{nombre}] no se pudo crear hoja vacía: {_e}")
                return

            # Convertir columnas nullable Int32/Int64 a float para xlsxwriter
            df_out = df_h.copy()
            for _c in df_out.columns:
                if hasattr(df_out[_c], "dtype") and str(df_out[_c].dtype) in ("Int32","Int64","Int8","Int16"):
                    df_out[_c] = df_out[_c].astype(float)

            try:
                df_out.to_excel(writer, sheet_name=nombre, index=False, startrow=1)
            except Exception as _e:
                print(f"  ❌ [{nombre}] to_excel falló: {_e}")
                try:
                    ws = wb.add_worksheet(nombre)
                    ws.write(0, 0, f"❌ Error al escribir datos: {_e}",
                             _f(bold=True, font_color="#C00000"))
                except Exception:
                    pass
                return

            ws = writer.sheets[nombre]

            try:
                fmt_tit = _f(bold=True, font_size=12, bg_color=color, font_color="#FFFFFF",
                             align="left", valign="vcenter")
                ws.merge_range(0, 0, 0, len(df_out.columns) - 1,
                    f"{titulo}  ·  {hoy_str}  ·  {len(df_out):,} guías", fmt_tit)
                ws.set_row(0, 20)
            except Exception as _e:
                print(f"  ⚠️  [{nombre}] merge_range falló: {_e}")

            anchos_col = {
                "GUIA": 18, "FECHA_DE_CREACION": 20, "MES": 16,
                "REGIONAL_ORIGEN": 22, "CODOFI": 10, "NOMOFI": 30,
                "CLIENTE": 28, "LINEA COMERCIAL": 18,
                "NOMEST": 32, "ESTADO_LEGALIZACION": 22,
                "N_PLANILLA_LEGALIZACION": 24, "FECHA_LEGALIZACION": 20,
                "T LEG CONTADO": 16, "TOTAL": 16,
                "DIAS_CONTROL": 16, "ORIGEN_DIAS": 24,
            }
            try:
                for i, cn in enumerate(df_out.columns):
                    ws.write(1, i, cn, fmt_enc)
                    ancho = anchos_col.get(cn, 20)
                    ws.set_column(i, i, ancho, fmt_cop if cn == "TOTAL" else fmt_cel)
                ws.freeze_panes(2, 0)
                ws.autofilter(1, 0, len(df_out) + 1, len(df_out.columns) - 1)
            except Exception as _e:
                print(f"  ⚠️  [{nombre}] encabezados/freeze falló: {_e}")

            # Semáforo DIAS_CONTROL — vectorizado
            try:
                if "DIAS_CONTROL" in df_out.columns:
                    ci_d = list(df_out.columns).index("DIAS_CONTROL")
                    for ri, val in enumerate(df_out["DIAS_CONTROL"]):
                        try:
                            d = int(val)
                            fmt_d = fmt_rojo if d > 5 else (fmt_naran if d > 2 else fmt_verde)
                            ws.write(ri + 2, ci_d, d, fmt_d)
                        except Exception:
                            pass
            except Exception as _e:
                print(f"  ⚠️  [{nombre}] semáforo DIAS_CONTROL falló: {_e}")

            # Semáforo ESTADO_LEGALIZACION — vectorizado (sin _norm_global por fila)
            try:
                if "ESTADO_LEGALIZACION" in df_out.columns:
                    ci_e = list(df_out.columns).index("ESTADO_LEGALIZACION")
                    fmt_sin    = _f(bold=True, bg_color="#C00000", font_color="#FFFFFF")
                    fmt_pre_e  = _f(bold=True, bg_color="#FF9900", font_color="#000000")
                    fmt_leg_ok = _f(bold=True, bg_color="#375623", font_color="#FFFFFF")
                    # Vectorizar: normalizar toda la columna de una vez
                    estados_n = _norm_global(df_out["ESTADO_LEGALIZACION"].astype(str))
                    for ri, (val, v_n) in enumerate(zip(df_out["ESTADO_LEGALIZACION"], estados_n)):
                        if v_n in _PRE_LEG_SET_N:
                            ws.write(ri + 2, ci_e, str(val), fmt_pre_e)
                        elif v_n in _LEGALIZADO_N:
                            ws.write(ri + 2, ci_e, str(val), fmt_leg_ok)
                        else:
                            ws.write(ri + 2, ci_e, str(val), fmt_sin)
            except Exception as _e:
                print(f"  ⚠️  [{nombre}] semáforo ESTADO_LEGALIZACION falló: {_e}")

            # Total al pie
            try:
                if "TOTAL" in df_out.columns:
                    ci_tot = list(df_out.columns).index("TOTAL")
                    total_sum = int(pd.to_numeric(df_out["TOTAL"], errors="coerce").sum())
                    fila_tot  = len(df_out) + 2
                    ws.write(fila_tot, ci_tot - 1,
                             f"TOTAL ({len(df_out):,} guías)", _f(bold=True, align="right"))
                    ws.write(fila_tot, ci_tot, total_sum, fmt_cop_tot)
            except Exception as _e:
                print(f"  ⚠️  [{nombre}] total al pie falló: {_e}")

            # Resumen agrupado
            try:
                if grp_df is not None and not grp_df.empty:
                    fila_grp = len(df_out) + 5
                    ws.write(fila_grp, 0, grp_titulo,
                             _f(bold=True, bg_color="#2E75B6", font_color="#FFFFFF"))
                    for ci, c in enumerate(grp_df.columns):
                        ws.write(fila_grp + 1, ci, c, fmt_enc)
                    for ri2, row2 in grp_df.iterrows():
                        for ci2, val2 in enumerate(row2):
                            col_name = grp_df.columns[ci2]
                            if col_name == "TOTAL_PESOS":
                                try:
                                    ws.write(fila_grp + 2 + ri2, ci2, int(val2), fmt_cop)
                                except Exception:
                                    ws.write(fila_grp + 2 + ri2, ci2, str(val2), fmt_cel)
                            elif col_name == "GUIAS":
                                try:
                                    ws.write(fila_grp + 2 + ri2, ci2, int(val2), fmt_num)
                                except Exception:
                                    ws.write(fila_grp + 2 + ri2, ci2, str(val2), fmt_cel)
                            else:
                                ws.write(fila_grp + 2 + ri2, ci2, str(val2), fmt_cel)
            except Exception as _e:
                print(f"  ⚠️  [{nombre}] resumen agrupado falló: {_e}")

                # ── RESUMEN ───────────────────────────────────────────────────────────
        df_resumen.to_excel(writer, sheet_name="RESUMEN", index=False, startrow=2)
        ws_r = writer.sheets["RESUMEN"]
        fmt_banner = _f(bold=True, font_size=13, bg_color="#1F3864",
                        font_color="#FFFFFF", align="center", valign="vcenter")
        ws_r.merge_range(0, 0, 1, 2,
            f"📋 CONTROL OPERATIVO Y FINANCIERO — TRANSCARGA MUNDIAL  ·  {hoy_str}",
            fmt_banner)
        ws_r.set_row(0, 22); ws_r.set_row(1, 22)
        for i, c in enumerate(df_resumen.columns):
            ws_r.write(2, i, c, fmt_enc)
        ws_r.set_column(0, 0, 60, fmt_cel)
        ws_r.set_column(1, 1, 22, fmt_cel)
        ws_r.set_column(2, 2, 14, fmt_num)
        ws_r.freeze_panes(3, 0)

        for ri, row in df_resumen.iterrows():
            cat = str(row["CATEGORÍA"])
            qty = row["CANTIDAD"]
            fila_xl = ri + 3
            if cat.startswith(("💵", "💰", "🚨", "⚠️", "📊")):
                fmt_s = fmt_secc2 if cat.startswith("💵") else fmt_secc
                for ci in range(3):
                    ws_r.write(fila_xl, ci, row.iloc[ci], fmt_s)
            elif cat.strip() == "":
                pass
            elif qty != "" and pd.notna(qty):
                try:
                    v = int(qty)
                    fmt_qty = fmt_rojo if v > 50 else (fmt_naran if v > 10 else fmt_verde)
                    ws_r.write(fila_xl, 0, row["CATEGORÍA"], fmt_cel)
                    ws_r.write(fila_xl, 1, row["TIPO / MODALIDAD"], fmt_cel)
                    ws_r.write(fila_xl, 2, v, fmt_qty)
                except Exception:
                    pass

        # Miniresúmenes en RESUMEN
        def _mini_resumen(df_mini, fila_ini, titulo_mini):
            if df_mini.empty:
                return
            ws_r.write(fila_ini, 0, titulo_mini,
                       _f(bold=True, bg_color="#2E75B6", font_color="#FFFFFF"))
            for ci, c in enumerate(df_mini.columns):
                ws_r.write(fila_ini + 1, ci, c, fmt_enc)
            for ri2, row2 in df_mini.iterrows():
                for ci2, val2 in enumerate(row2):
                    try:
                        ws_r.write(fila_ini + 2 + ri2, ci2,
                                   int(val2) if isinstance(val2, (int, float, np.integer))
                                       else val2,
                                   fmt_num if ci2 == len(df_mini.columns) - 1 else fmt_cel)
                    except Exception:
                        ws_r.write(fila_ini + 2 + ri2, ci2, str(val2), fmt_cel)

        fila_base = len(df_resumen) + 5
        for rsm, titulo_rsm in [
            (resumen_pre, "📋 PRE LEG VENCIDA — por Modalidad y Regional"),
            (resumen_ce,  "📋 CE SIN LEGALIZAR — por Regional Destino"),
            (resumen_cnt, "📋 CONTADO SIN LEG (umbral dif.) — por Regional Origen"),
            (resumen_adm, "📋 ADMITIDOS DEMORADOS — por Regional Origen"),
            (resumen_ops, "📋 OPS DEMORADAS — por Estado y Regional"),
        ]:
            _mini_resumen(rsm, fila_base, titulo_rsm)
            fila_base += max(len(rsm), 0) + 4

        # ── Pestañas CONTADOS (Gisell) — primero para que queden al inicio ────
        def _escribir_hoja(fn, *args, **kwargs):
            """Wrapper que aísla cada pestaña: si una falla, el resto continúa."""
            nombre = args[0] if args else "?"
            try:
                print(f"  → Escribiendo pestaña: {nombre}...")
                fn(*args, **kwargs)
                print(f"     ✅ {nombre} OK")
            except Exception as _e_hoja:
                import traceback as _tb
                print(f"  ❌ ERROR en pestaña {nombre}: {_e_hoja}")
                _tb.print_exc()
                # Intentar crear hoja de error para que quede visible
                try:
                    ws_err = wb.add_worksheet(nombre[:28] + "_ERR")
                    ws_err.write(0, 0, f"❌ Error: {_e_hoja}",
                                 wb.add_format({"bold": True, "font_color": "#C00000"}))
                except Exception:
                    pass

        _escribir_hoja(_hoja_contados,
            "CONTADOS_VP_PEND",
            df_vp_pend,
            "💵 CONTADOS VENTA PROPIA — Pendientes de legalizar  (BOG≤1d · resto≤2d)",
            "#843C0C",
            _grp_vp,
            "📋 Resumen por Regional Origen + Planilla"
        )

        _escribir_hoja(_hoja_contados,
            "CONTADOS_PRE_PLANILLA",
            df_pre_planilla,
            "📋 PRE LEGALIZADO SIN CERRAR — CONTADO  ·  Por regional + planilla",
            "#FF9900",
        )

        _escribir_hoja(_hoja_contados,
            "CONTADOS_VP_LEG",
            df_vp_leg,
            "✅ CONTADOS VENTA PROPIA — Legalizados (historial/conciliación)",
            "#375623",
        )

        _escribir_hoja(_hoja_contados,
            "CONTADOS_FRANQ",
            df_franq_pend,
            "🏪 CONTADOS FRANQUICIAS — Seguimiento pendientes (sin plazo fijo)",
            "#2E75B6",
            _grp_franq,
            "📋 Resumen por Regional Origen + Franquicia"
        )

        # ── Pestañas operativas ───────────────────────────────────────────────
        _escribir_hoja(_hoja, "PRE_LEG_VENCIDA",
              df_pre,
              "⚠️ PRE LEGALIZADO VENCIDO  >2 días hábiles",
              "#C00000",
              resumen_pre,
              "Resumen por Modalidad y Regional")

        _escribir_hoja(_hoja, "CE_SIN_LEGALIZAR",
              df_ce_sin,
              "💰 CONTRA ENTREGA entregada SIN pre ni legalizar",
              "#C00000",
              resumen_ce,
              "Resumen por Regional Destino")

        _escribir_hoja(_hoja, "ADMITIDOS_DEMO",
              df_adm,
              "🚨 ADMITIDOS DEMORADOS >2 días  —  Alerta ORIGEN",
              "#7030A0",
              resumen_adm,
              "Resumen por Regional Origen")

        _escribir_hoja(_hoja, "OPS_DEMORADAS",
              df_ops,
              "🔧 OPERACIONES DEMORADAS  —  Alerta OPERACIONES",
              "#375623",
              resumen_ops,
              "Resumen por Estado y Regional Destino")

        _escribir_hoja(_hoja, "NOVEDADES_SAC",
              df_sac,
              "⚠️ NOVEDADES SAC >1 día hábil  —  Alerta SAC",
              "#2E75B6")

        _escribir_hoja(_hoja, "NOVEDADES_OPS",
              df_nov_ops,
              "🔴 NOVEDADES OPERATIVAS >2 días hábiles",
              "#C00000")

    print(f"  ✅ ALERTAS_CONTROL guardado: {ruta_out.name}  [{time.time()-t0_xl:.1f}s]")
    print(f"     CONTADOS_VP_PEND:      {int(mask_vp_pend.sum()):,} guías")
    print(f"     CONTADOS_PRE_PLANILLA: {int(mask_contado_pre.sum()):,} guías")
    print(f"     CONTADOS_VP_LEG:       {int(mask_vp_leg.sum()):,} guías")
    print(f"     CONTADOS_FRANQ:        {int(mask_franq_pend.sum()):,} guías")
    print(f"     PRE_LEG_VENCIDA:       {int(mask_pre.sum()):,} guías")
    print(f"     CE_SIN_LEGALIZAR:      {int(mask_ce_sin.sum()):,} guías")
    print(f"     ADMITIDOS_DEMO:        {int(mask_adm.sum()):,} guías")
    print(f"     OPS_DEMORADAS:         {int(mask_ops_total.sum()):,} guías"
          f"  (fijo: {int(mask_ops_fijo.sum()):,} | reparto: {int(mask_ops_reparto.sum()):,})")
    print(f"     NOVEDADES_SAC:         {int(mask_sac.sum()):,} guías")
    print(f"     NOVEDADES_OPS:         {int(mask_nov_ops.sum()):,} guías")
    return ruta_out


def generar_reporte_comercial(df: pd.DataFrame, carpeta_salida: Path) -> Path:
    """
    Genera REPORTE_COMERCIAL_<ts>.xlsx con indicadores avanzados de ventas.
    Incluye clasificación por rangos y alertas de baja de facturación.
    """
    sep("GENERANDO REPORTE COMERCIAL INDEPENDIENTE")
    if "COMERCIAL" not in df.columns:
        print("  ❌ No se encontró la columna COMERCIAL. Abortando reporte comercial.")
        return None

    ts      = pd.Timestamp.now().strftime("%Y%m%d_%H%M")
    hoy_str = pd.Timestamp.now().strftime("%d/%m/%Y %H:%M")
    ruta_out = carpeta_salida / f"REPORTE_COMERCIAL_{ts}.xlsx"

    # --- 1. PREPARACIÓN DE DATOS ---
    # Excluir Franquicias del reporte comercial según solicitud (se manejan en reporte aparte)
    df_com = df.copy()
    if "LINEA COMERCIAL" in df_com.columns:
        df_com = df_com[df_com["LINEA COMERCIAL"] != "FRANQUICIAS"].copy()
    
    # Asegurar que TOTAL sea numérico
    df_com["TOTAL"] = pd.to_numeric(df_com["TOTAL"], errors="coerce").fillna(0)
    
    # --- 2. RESUMEN EJECUTIVO (Por Comercial y Mes) ---
    pivot_ventas_mes = df_com.pivot_table(
        index="COMERCIAL", columns="MES", values="TOTAL", aggfunc="sum", fill_value=0
    )
    # Cantidad de Clientes únicos por mes
    pivot_clientes_mes = df_com.groupby(["COMERCIAL", "MES"])["CLIENTE"].nunique().unstack(fill_value=0)
    
    # Unir en resumen consolidado por comercial
    res_metrica = pivot_ventas_mes.copy()
    res_metrica["TOTAL_PERIODO"] = res_metrica.sum(axis=1)
    res_metrica = res_metrica.sort_values("TOTAL_PERIODO", ascending=False).reset_index()

    # --- 3. SEGUIMIENTO DE CLIENTES Y ALERTAS DE BAJA ---
    pivot_cli_ventas = df_com.pivot_table(
        index=["COMERCIAL", "CLIENTE"], columns="MES", values="TOTAL", aggfunc="sum", fill_value=0
    )
    pivot_cli_guias = df_com.pivot_table(
        index=["COMERCIAL", "CLIENTE"], columns="MES", values="GUIA", aggfunc="count", fill_value=0
    )

    meses_cols = sorted([c for c in pivot_cli_ventas.columns if " - " in str(c)])
    
    if len(meses_cols) >= 2:
        ultimo_mes = meses_cols[-1]
        meses_previos = meses_cols[:-1]
        prom_previo = pivot_cli_ventas[meses_previos].mean(axis=1)
        venta_actual = pivot_cli_ventas[ultimo_mes]
        
        # Alerta: caída > 25% vs promedio anterior (mínimo 100k de promedio para evitar ruido)
        mask_alerta = (venta_actual < (prom_previo * 0.75)) & (prom_previo > 100000)
        
        df_seguimiento = pivot_cli_ventas.copy()
        df_seguimiento["PROM_HISTORICO"] = prom_previo
        df_seguimiento["ESTADO"] = np.where(mask_alerta, "🚨 BAJA DETECTADA", "✅ ESTABLE")
        
        # Clasificación Prime/Top/Estándar
        def _clasificar(valor):
            if valor > 5000000: return "1. PRIME (>5M)"
            if valor > 2000000: return "2. TOP (>2M)"
            return "3. ESTANDAR"
        
        df_seguimiento["RANGO_FACT"] = df_seguimiento["PROM_HISTORICO"].apply(_clasificar)
        df_seguimiento = df_seguimiento.reset_index().sort_values(["RANGO_FACT", "COMERCIAL"], ascending=[True, True])
    else:
        df_seguimiento = pivot_cli_ventas.reset_index()
        df_seguimiento["ESTADO"] = "N/A (Faltan datos)"
        df_seguimiento["RANGO_FACT"] = "N/A"

    # --- 4. ESCRITURA EXCEL ---
    with pd.ExcelWriter(ruta_out, engine="xlsxwriter") as writer:
        wb = writer.book
        fmt_tit = wb.add_format({"bold": True, "font_size": 12, "bg_color": "#4472C4", "font_color": "#FFFFFF", "align": "center"})
        fmt_enc = wb.add_format({"bold": True, "bg_color": "#D9E1F2", "border": 1, "align": "center"})
        fmt_money = wb.add_format({"num_format": "$#,##0", "font_size": 10})
        fmt_num   = wb.add_format({"num_format": "#,##0", "align": "center", "font_size": 10})
        fmt_alert = wb.add_format({"bg_color": "#FFC7CE", "font_color": "#9C0006", "bold": True})

        # Hoja 1: RESUMEN EJECUTIVO (Ventas)
        res_metrica.to_excel(writer, sheet_name="VENTAS_COMERCIAL", index=False, startrow=1)
        ws1 = writer.sheets["VENTAS_COMERCIAL"]
        ws1.merge_range(0, 0, 0, len(res_metrica.columns)-1, f"INDICADORES DE VENTAS MENSUALES POR COMERCIAL  -  {hoy_str}", fmt_tit)
        for i, col in enumerate(res_metrica.columns):
            ws1.write(1, i, col, fmt_enc)
            ws1.set_column(i, i, 16, fmt_money if i > 0 else None)
        ws1.set_column(0, 0, 25)

        # Hoja 2: ALERTAS DE BAJA Y RANGOS
        df_seguimiento.to_excel(writer, sheet_name="ALERTAS_Y_RANGOS", index=False, startrow=1)
        ws2 = writer.sheets["ALERTAS_Y_RANGOS"]
        ws2.merge_range(0, 0, 0, len(df_seguimiento.columns)-1, "AUDITORIA DE CLIENTES: SEGMENTACION Y ALERTAS DE CAÍDA", fmt_tit)
        for i, col in enumerate(df_seguimiento.columns):
            ws2.write(1, i, col, fmt_enc)
            if col in meses_cols or col == "PROM_HISTORICO":
                ws2.set_column(i, i, 16, fmt_money)
            elif col == "CLIENTE": ws2.set_column(i, i, 35)
            elif col == "ESTADO": 
                ws2.set_column(i, i, 20)
                ws2.conditional_format(2, i, len(df_seguimiento)+1, i, {
                    "type": "cell", "criteria": "equal to", "value": '"🚨 BAJA DETECTADA"', "format": fmt_alert
                })
        ws2.freeze_panes(2, 2)

        # Hoja 3: DETALLE GUÌAS POR CLIENTE
        pivot_cli_guias.reset_index().to_excel(writer, sheet_name="GUIAS_POR_CLIENTE", index=False, startrow=1)
        ws3 = writer.sheets["GUIAS_POR_CLIENTE"]
        ws3.merge_range(0, 0, 0, len(pivot_cli_guias.columns)+1, "CANTIDAD DE GUIAS POR CLIENTE Y MES (AUDITORIA)", fmt_tit)
        for i, col in enumerate(["COMERCIAL", "CLIENTE"] + list(pivot_cli_guias.columns)):
            ws3.write(1, i, col, fmt_enc)
            ws3.set_column(i, i, 15, fmt_num if i > 1 else None)
        ws3.set_column(1, 1, 35)

        # Hoja 4: CLIENTES ÚNICOS POR COMERCIAL
        pivot_clientes_mes.reset_index().to_excel(writer, sheet_name="CONTEO_CLIENTES", index=False, startrow=1)
        ws4 = writer.sheets["CONTEO_CLIENTES"]
        ws4.merge_range(0, 0, 0, len(pivot_clientes_mes.columns), "NUMERO DE CLIENTES UNICOS ATENDIDOS POR MES", fmt_tit)
        for i, col in enumerate(["COMERCIAL"] + list(pivot_clientes_mes.columns)):
            ws4.write(1, i, col, fmt_enc)
            ws4.set_column(i, i, 15, fmt_num if i > 0 else None)
        ws4.set_column(0, 0, 25)

    print(f"  ✅ REPORTE_COMERCIAL guardado: {ruta_out.name}")
    return ruta_out


def generar_reporte_franquicias(df: pd.DataFrame, carpeta_salida: Path) -> Path:
    """
    Genera REPORTE_FRANQUICIAS_<ts>.xlsx exclusivo para terceros.
    Agrupa por REGIONAL ORIGEN y NOMOFI, sin detalle de clientes.
    """
    sep("GENERANDO REPORTE DE FRANQUICIAS")
    if "LINEA COMERCIAL" not in df.columns:
        print("  ❌ No se encontró LINEA COMERCIAL. Abortando reporte franquicias.")
        return None

    # Filtrar solo franquicias
    df_fra = df[df["LINEA COMERCIAL"] == "FRANQUICIAS"].copy()
    if df_fra.empty:
        print("  ⚠️  No se encontraron registros de FRANQUICIAS en el maestro.")
        return None

    ts      = pd.Timestamp.now().strftime("%Y%m%d_%H%M")
    hoy_str = pd.Timestamp.now().strftime("%d/%m/%Y %H:%M")
    ruta_out = carpeta_salida / f"REPORTE_FRANQUICIAS_{ts}.xlsx"

    # Preparación
    df_fra["TOTAL"] = pd.to_numeric(df_fra["TOTAL"], errors="coerce").fillna(0)
    
    # Asegurar que NOMOFI y REGIONAL_ORIGEN existan
    for c in ["REGIONAL_ORIGEN", "NOMOFI"]:
        if c not in df_fra.columns: df_fra[c] = "DESCONOCIDO"

    # --- 1. RESUMEN POR REGIONAL ---
    res_reg = df_fra.pivot_table(
        index="REGIONAL_ORIGEN", columns="MES", values="TOTAL", aggfunc="sum", fill_value=0
    )
    res_reg["TOTAL_PERIODO"] = res_reg.sum(axis=1)
    res_reg = res_reg.sort_values("TOTAL_PERIODO", ascending=False).reset_index()

    # --- 2. DESEMPEÑO POR OFICINA (NOMOFI) CON ALERTAS ---
    pivot_ofi_v = df_fra.pivot_table(
        index=["REGIONAL_ORIGEN", "NOMOFI"], columns="MES", values="TOTAL", aggfunc="sum", fill_value=0
    )
    pivot_ofi_g = df_fra.pivot_table(
        index=["REGIONAL_ORIGEN", "NOMOFI"], columns="MES", values="GUIA", aggfunc="count", fill_value=0
    )

    meses_cols = sorted([c for c in pivot_ofi_v.columns if " - " in str(c)])
    
    if len(meses_cols) >= 2:
        ultimo_mes = meses_cols[-1]
        meses_previos = meses_cols[:-1]
        prom_previo = pivot_ofi_v[meses_previos].mean(axis=1)
        venta_actual = pivot_ofi_v[ultimo_mes]
        
        # Alerta: caída > 20% vs promedio operativo de la oficina
        mask_alerta = (venta_actual < (prom_previo * 0.80)) & (prom_previo > 50000)
        
        df_alertas = pivot_ofi_v.copy()
        df_alertas["TOTAL_PERIODO"] = df_alertas.sum(axis=1)
        df_alertas["PROM_HISTORICO"] = prom_previo
        df_alertas["ESTADO"] = np.where(mask_alerta, "🚨 CAÍDA OPERATIVA", "✅ OK")
        df_alertas = df_alertas.reset_index().sort_values(["REGIONAL_ORIGEN", "TOTAL_PERIODO"], ascending=[True, False])
        # Re-sort manual si falló el pivot sort
        if "REGIONAL_ORIGEN" in df_alertas.columns:
            df_alertas = df_alertas.sort_values("REGIONAL_ORIGEN")
    else:
        df_alertas = pivot_ofi_v.reset_index()
        df_alertas["ESTADO"] = "N/A"

    # --- 3. ESCRITURA EXCEL ---
    with pd.ExcelWriter(ruta_out, engine="xlsxwriter") as writer:
        wb = writer.book
        fmt_tit = wb.add_format({"bold": True, "font_size": 12, "bg_color": "#7030A0", "font_color": "#FFFFFF", "align": "center"})
        fmt_enc = wb.add_format({"bold": True, "bg_color": "#E2EFDA", "border": 1, "align": "center"})
        fmt_money = wb.add_format({"num_format": "$#,##0", "font_size": 10})
        fmt_num   = wb.add_format({"num_format": "#,##0", "align": "center", "font_size": 10})
        fmt_alert = wb.add_format({"bg_color": "#FFC7CE", "font_color": "#9C0006", "bold": True})

        # Hoja 1: Resumen Regional
        res_reg.to_excel(writer, sheet_name="RESUMEN_REGIONAL", index=False, startrow=1)
        ws1 = writer.sheets["RESUMEN_REGIONAL"]
        ws1.merge_range(0, 0, 0, len(res_reg.columns)-1, f"VENTAS DE FRANQUICIAS POR REGIONAL  -  {hoy_str}", fmt_tit)
        for i, col in enumerate(res_reg.columns):
            ws1.write(1, i, col, fmt_enc)
            ws1.set_column(i, i, 18 if i > 0 else 25, fmt_money if i > 0 else None)

        # Hoja 2: Detalle Oficinas y Alertas
        df_alertas.to_excel(writer, sheet_name="DESEMPEÑO_OFICINAS", index=False, startrow=1)
        ws2 = writer.sheets["DESEMPEÑO_OFICINAS"]
        ws2.merge_range(0, 0, 0, len(df_alertas.columns)-1, "AUDITORIA DE FRANQUICIAS (NOMOFI) Y ALERTAS DE PRODUCTIVIDAD", fmt_tit)
        for i, col in enumerate(df_alertas.columns):
            ws2.write(1, i, col, fmt_enc)
            if col in meses_cols or col == "PROM_HISTORICO":
                ws2.set_column(i, i, 16, fmt_money)
            elif col == "NOMOFI": ws2.set_column(i, i, 30)
            elif col == "ESTADO": 
                ws2.set_column(i, i, 20)
                ws2.conditional_format(2, i, len(df_alertas)+1, i, {
                    "type": "cell", "criteria": "equal to", "value": '"🚨 CAÍDA OPERATIVA"', "format": fmt_alert
                })
        ws2.freeze_panes(2, 2)

        # Hoja 3: Conteo de Guías
        pivot_ofi_g.reset_index().to_excel(writer, sheet_name="GUIAS_POR_OFICINA", index=False, startrow=1)
        ws3 = writer.sheets["GUIAS_POR_OFICINA"]
        ws3.merge_range(0, 0, 0, len(pivot_ofi_g.columns)+1, "VOLUMEN DE GUIAS POR FRANQUICIA Y MES", fmt_tit)
        for i, col in enumerate(["REGIONAL", "NOMOFI"] + list(pivot_ofi_g.columns)):
            ws3.write(1, i, col, fmt_enc)
            ws3.set_column(i, i, 15 if i > 1 else 30, fmt_num if i > 1 else None)
        ws3.set_column(1, 1, 30)

    print(f"  ✅ REPORTE_FRANQUICIAS guardado: {ruta_out.name}")
    return ruta_out


# =============================================================================
#  ALERTAS FRANQUICIAS BOG — v4.7
#  Genera ALERTAS_FRANQUICIAS_BOG_<ts>.xlsx con seguimiento exclusivo de los
#  despachos generados por FRANQUICIAS de REGIONAL ORIGEN = BOGOTA.
#
#  PESTAÑAS:
#    RESUMEN         - dashboard de conteos de las 3 pestañas operativas
#    ADMITIDOS_DEMO  - ADMITIDO + ENVIO PROCESADO >2 dias habiles
#    NOVEDADES_SAC   - estados SAC demorados >1 dia habil
#    OPS_DEMORADAS   - estados OPS con umbrales iguales al archivo general
#
#  FILTRO BASE (aplica a TODO el archivo):
#    LINEA COMERCIAL == "FRANQUICIAS"  AND  REGIONAL_ORIGEN == "BOGOTA"
#
#  ESTADOS FINALES excluidos de todas las alertas:
#    ENTREGADO EN SATISFACCION, DEVOLUCION POR NOVEDAD, INCAUTACION,
#    EN RECLAMACION, ANULADO
# =============================================================================

def generar_alertas_franquicias_bog(df: pd.DataFrame, festivos: np.ndarray,
                                     carpeta_salida: Path) -> Path:
    """
    Genera ALERTAS_FRANQUICIAS_BOG_<ts>.xlsx — seguimiento de despachos
    de Franquicias Bogota: ADMITIDOS demorados, Novedades SAC y OPS demoradas.
    """
    sep("GENERANDO ALERTAS FRANQUICIAS BOG")
    ts      = pd.Timestamp.now().strftime("%Y%m%d_%H%M")
    hoy_str = pd.Timestamp.now().strftime("%d/%m/%Y %H:%M")
    ruta_out = carpeta_salida / f"ALERTAS_FRANQUICIAS_BOG_{ts}.xlsx"

    def _col(nombre):
        return df[nombre] if nombre in df.columns else pd.Series(pd.NA, index=df.index)

    # -- Series normalizadas ---------------------------------------------------
    nomest_n    = _norm_global(_col("NOMEST"))
    linea_com_n = _norm_global(_col("LINEA COMERCIAL"))
    reg_ori_n   = _norm_global(_col("REGIONAL_ORIGEN").fillna(""))
    tray_n      = _norm_global(_col("TIPO TRAYECTO"))
    dias_est    = pd.to_numeric(_col("DIAS_EN_ESTADO"), errors="coerce")
    total_num   = pd.to_numeric(_col("TOTAL"), errors="coerce").fillna(0)

    # -- Filtro base: FRANQUICIAS + REGIONAL ORIGEN BOGOTA + TOTAL > 0 --------
    # Estados finales NUNCA en alertas
    _FINALES_BOG_N = {"ENTREGADO EN SATISFACCION", "DEVOLUCION POR NOVEDAD",
                      "INCAUTACION", "EN RECLAMACION", "ANULADO"}

    mask_base = (
        (linea_com_n == "FRANQUICIAS")
        & reg_ori_n.str.startswith("BOGOTA")
        & (total_num > 0)
        & (~nomest_n.isin(_FINALES_BOG_N))
    )

    n_base = int(mask_base.sum())
    print(f"  Info: Universo Franquicias BOG (no finales, TOTAL>0): {n_base:,} guias")
    if n_base == 0:
        print(f"  Aviso: Sin guias de Franquicias BOG en el maestro. Se genera archivo vacio.")

    # -- Dias habiles ----------------------------------------------------------
    _hoy_np = np.datetime64(pd.Timestamp.now().date(), "D")

    def _dias_hab(fecha_col_nombre: str) -> pd.Series:
        if fecha_col_nombre not in df.columns:
            return pd.Series(-1, index=df.index, dtype=np.int32)
        col = df[fecha_col_nombre]
        mask_ok = col.notna()
        fechas_np = col[mask_ok].values.astype("datetime64[D]")
        resultado = np.full(len(df), -1, dtype=np.int32)
        resultado[mask_ok.values] = np.maximum(
            np.busday_count(
                fechas_np + np.timedelta64(1, "D"),
                _hoy_np + np.timedelta64(1, "D"),
                weekmask="Mon Tue Wed Thu Fri Sat",
                holidays=festivos,
            ), 0
        ).astype(np.int32)
        return pd.Series(resultado, index=df.index)

    # -- Columnas de detalle (igual que ALERTAS_CONTROL + NOMOFI) -------------
    COLS_BOG = [c for c in [
        "GUIA", "FECHA_DE_CREACION", "MES", "NOMEST", "FECREGEST", "DIAS_EN_ESTADO",
        "MODALIDAD", "TIPO TRAYECTO", "CIUDAD_ORIGEN", "CIUDAD_DESTINO",
        "NOMOFI", "REGIONAL_ORIGEN", "REG DESTINO CORRECTA",
        "ESTADO_LEGALIZACION", "FECHA_LEGALIZACION",
        "TIEMPO DE ENTREGA", "T LEG CONTADO", "T CE CREACION A LEG", "T CE ENTREGA A LEG",
        "REMITENTE_CLIENTE", "DESTINATARIO_CLIENTE", "TOTAL",
        "OBSERVACION", "OBSERVACION_GUIA",
    ] if c in df.columns]

    def _det(mask, sort_col="DIAS_EN_ESTADO"):
        df_h = df.loc[mask, COLS_BOG].copy()
        if not df_h.empty and sort_col in df_h.columns:
            df_h = df_h.sort_values(sort_col, ascending=False)
        return df_h

    # =========================================================================
    #  MASKS DE ALERTA — sobre el universo filtrado (mask_base)
    # =========================================================================

    # -- ADMITIDOS demorados >2 dias ------------------------------------------
    mask_adm = (
        mask_base
        & nomest_n.isin(_ADMITIDOS_N_SET)
        & (dias_est > UMBRAL_ADMITIDOS_DIAS)
    )

    # -- OPS demorados — umbral fijo ------------------------------------------
    # Mismos estados y umbrales que el archivo general.
    # MERCANCIA NO INGRESA A BODEGA: excluida de OPS_DEMORADAS (es de NOVEDADES_OPS)
    _OPS_FIJO_BOG = {
        "EN TRANSITO A CIUDAD DESTINO" : 2,
        "INGRESO A BODEGA"             : 2,
        "CARGADA EN VEHICULO"          : 2,
        "INGRESO A BODEGA DESTINO"     : 2,
    }
    mask_ops_fijo = pd.Series(False, index=df.index)
    for est_n, umb in _OPS_FIJO_BOG.items():
        mask_ops_fijo |= mask_base & (nomest_n == est_n) & (dias_est > umb)

    # EN REPARTO URBANO — umbral diferenciado DIRECTO=1d / REEXP=3d
    mask_reparto_dir  = (
        mask_base
        & (nomest_n == _REPARTO_N)
        & (tray_n == "DIRECTO")
        & (dias_est > UMBRAL_REPARTO_DIRECTO)
    )
    mask_reparto_reex = (
        mask_base
        & (nomest_n == _REPARTO_N)
        & (tray_n == "REEXPEDICION")
        & (dias_est > UMBRAL_REPARTO_REEXP)
    )
    mask_ops_total = mask_ops_fijo | mask_reparto_dir | mask_reparto_reex

    # -- SAC demorados >1 dia -------------------------------------------------
    mask_sac = (
        mask_base
        & nomest_n.isin(_SAC_N_SET)
        & (dias_est > UMBRAL_SAC_DIAS)
    )

    # -- DataFrames de detalle ------------------------------------------------
    df_adm = _det(mask_adm)
    df_ops = _det(mask_ops_total)
    df_sac = _det(mask_sac)

    # -- Resumenes intermedios ------------------------------------------------
    reg_dest_bog = _col("REG DESTINO CORRECTA").fillna("SIN REGIONAL")
    reg_ori_bog  = _col("REGIONAL_ORIGEN").fillna("SIN REGIONAL")

    def _resumen_por(mask, col_ser, nombre_col):
        if mask.sum() == 0:
            return pd.DataFrame(columns=[nombre_col, "CANTIDAD"])
        sr = col_ser[mask].value_counts().reset_index()
        sr.columns = [nombre_col, "CANTIDAD"]
        return sr

    resumen_adm = _resumen_por(mask_adm, reg_ori_bog,  "REGIONAL_ORIGEN")
    resumen_sac = _resumen_por(mask_sac, reg_dest_bog, "REG DESTINO CORRECTA")

    if mask_ops_total.sum() > 0:
        resumen_ops = (
            pd.DataFrame({
                "ESTADO":     _col("NOMEST")[mask_ops_total].values,
                "FRANQUICIA": _col("NOMOFI")[mask_ops_total].values,
            })
            .value_counts().reset_index(name="CANTIDAD")
        )
    else:
        resumen_ops = pd.DataFrame(columns=["ESTADO", "FRANQUICIA", "CANTIDAD"])

    # =========================================================================
    #  RESUMEN GENERAL
    # =========================================================================
    _ops_fijo_estados = set(_OPS_FIJO_BOG)
    resumen_rows = [
        ("FRANQUICIAS BOGOTA - Seguimiento de despachos", "", ""),
        (f"Universo total (no finales, TOTAL>0)", "FRANQ-BOG", n_base),
        ("", "", ""),
        ("ALERTAS OPERACIONES", "", ""),
        ("ADMITIDOS demorados >2 dias habiles",
         "ADMITIDOS",   int(mask_adm.sum())),
        ("OPS demoradas - umbral fijo >2 dias",
         "OPS-FIJO",    int(mask_ops_fijo.sum())),
        ("EN REPARTO URBANO - DIRECTO >1 dia",
         "OPS-DIRECTO", int(mask_reparto_dir.sum())),
        ("EN REPARTO URBANO - REEXPEDICION >3 dias",
         "OPS-REEXP",   int(mask_reparto_reex.sum())),
        ("", "", ""),
        ("NOVEDADES SAC", "", ""),
        ("Novedades SAC demoradas >1 dia habil",
         "SAC",         int(mask_sac.sum())),
        ("", "", ""),
        ("COMPOSICION DEL UNIVERSO BOG", "", ""),
        ("Guias en estado ADMITIDO / ENVIO PROCESADO",
         "ADMITIDOS",
         int((mask_base & nomest_n.isin(_ADMITIDOS_N_SET)).sum())),
        ("Guias en estados OPS (movimiento fisico)",
         "OPS",
         int((mask_base & nomest_n.isin(_ops_fijo_estados | {_REPARTO_N})).sum())),
        ("Guias en novedades SAC",
         "SAC",
         int((mask_base & nomest_n.isin(_SAC_N_SET)).sum())),
    ]
    df_resumen = pd.DataFrame(resumen_rows,
                              columns=["CATEGORIA", "TIPO / MODALIDAD", "CANTIDAD"])

    # =========================================================================
    #  ESCRITURA EXCEL
    # =========================================================================
    print(f"  Escribiendo archivo FRANQUICIAS BOG...")
    t0_xl = time.time()

    with pd.ExcelWriter(ruta_out, engine="xlsxwriter") as writer:
        wb = writer.book

        # -- Paleta Franquicias BOG -------------------------------------------
        COLOR_HEADER = "#4B0082"   # violeta oscuro - identidad visual BOG
        COLOR_ADM    = "#7030A0"
        COLOR_SAC    = "#2E75B6"
        COLOR_OPS    = "#375623"

        _f = lambda **kw: wb.add_format({"font_name": "Calibri", "font_size": 10, **kw})
        fmt_enc   = _f(bold=True, bg_color="#1F3864", font_color="#FFFFFF",
                       align="center", valign="vcenter")
        fmt_cel   = _f()
        fmt_num   = _f(num_format="#,##0", align="center")
        fmt_cop   = _f(num_format="$#,##0")
        fmt_rojo  = _f(bold=True, bg_color="#C00000", font_color="#FFFFFF", align="center")
        fmt_naran = _f(bold=True, bg_color="#FF9900", font_color="#000000", align="center")
        fmt_verde = _f(bg_color="#E2EFDA")
        fmt_secc  = _f(bold=True, bg_color="#D6E4F0")

        def _hoja_bog(nombre: str, df_h: pd.DataFrame, titulo: str,
                      color: str = "#1F3864",
                      resumen_extra: pd.DataFrame = None,
                      resumen_titulo: str = ""):
            """Helper estandar de hoja operativa para el archivo BOG."""
            if df_h.empty:
                ws = wb.add_worksheet(nombre)
                ws.write(0, 0, f"Sin alertas en esta categoria al {hoy_str}",
                         _f(bold=True, font_color="#375623", font_size=11))
                return

            df_h.to_excel(writer, sheet_name=nombre, index=False, startrow=1)
            ws = writer.sheets[nombre]

            fmt_tit = _f(bold=True, font_size=12, bg_color=color, font_color="#FFFFFF",
                         align="left", valign="vcenter")
            ws.merge_range(0, 0, 0, len(df_h.columns) - 1,
                f"{titulo}  -  {hoy_str}  -  {len(df_h):,} guias", fmt_tit)
            ws.set_row(0, 20)

            for i, cn in enumerate(df_h.columns):
                ws.write(1, i, cn, fmt_enc)
                try:
                    ql = df_h[cn].astype(str).str.len().quantile(0.9)
                    ancho = min(int(max(len(cn), ql)) + 3, 42)
                except Exception:
                    ancho = 18
                ws.set_column(i, i, ancho, fmt_cel)

            ws.freeze_panes(2, 0)
            ws.autofilter(1, 0, len(df_h) + 1, len(df_h.columns) - 1)

            if "DIAS_EN_ESTADO" in df_h.columns:
                ci = list(df_h.columns).index("DIAS_EN_ESTADO")
                for ri, val in enumerate(df_h["DIAS_EN_ESTADO"]):
                    try:
                        d = int(val)
                        fmt_d = fmt_rojo if d > 10 else (fmt_naran if d > 5 else fmt_verde)
                        ws.write(ri + 2, ci, d, fmt_d)
                    except Exception:
                        pass

            if resumen_extra is not None and not resumen_extra.empty:
                fila_res = len(df_h) + 4
                ws.write(fila_res, 0, resumen_titulo,
                         _f(bold=True, bg_color="#2E75B6", font_color="#FFFFFF"))
                for ci, c in enumerate(resumen_extra.columns):
                    ws.write(fila_res + 1, ci, c, fmt_enc)
                for ri, row in resumen_extra.iterrows():
                    for ci, val in enumerate(row):
                        ws.write(fila_res + 2 + ri, ci,
                                 int(val) if isinstance(val, (int, float))
                                     and ci == len(resumen_extra.columns) - 1
                                     else val,
                                 fmt_num if ci == len(resumen_extra.columns) - 1
                                     else fmt_cel)

        def _escribir_hoja(fn, *args, **kwargs):
            nombre = args[0] if args else "?"
            try:
                print(f"  -> Escribiendo pestana BOG: {nombre}...")
                fn(*args, **kwargs)
                print(f"     OK {nombre}")
            except Exception as _e:
                import traceback as _tb
                print(f"  ERROR en pestana BOG {nombre}: {_e}")
                _tb.print_exc()
                try:
                    ws_err = wb.add_worksheet(nombre[:28] + "_ERR")
                    ws_err.write(0, 0, f"Error: {_e}",
                                 wb.add_format({"bold": True, "font_color": "#C00000"}))
                except Exception:
                    pass

        # -- Pestana RESUMEN --------------------------------------------------
        try:
            ws_r = wb.add_worksheet("RESUMEN")
            writer.sheets["RESUMEN"] = ws_r
            fmt_tit_r = _f(bold=True, font_size=13, bg_color=COLOR_HEADER,
                           font_color="#FFFFFF", align="left", valign="vcenter")
            ws_r.merge_range(0, 0, 0, 2,
                f"FRANQUICIAS BOGOTA - Control de despachos  -  {hoy_str}",
                fmt_tit_r)
            ws_r.merge_range(1, 0, 1, 2,
                "Filtro: LINEA COMERCIAL = FRANQUICIAS + REGIONAL_ORIGEN = BOGOTA",
                _f(italic=True, font_color="#595959", font_size=9))
            ws_r.set_row(0, 22); ws_r.set_row(1, 16)
            for i, c in enumerate(df_resumen.columns):
                ws_r.write(2, i, c, fmt_enc)
            ws_r.set_column(0, 0, 62, fmt_cel)
            ws_r.set_column(1, 1, 22, fmt_cel)
            ws_r.set_column(2, 2, 14, fmt_num)
            ws_r.freeze_panes(3, 0)

            for ri, row in df_resumen.iterrows():
                cat = str(row["CATEGORIA"])
                qty = row["CANTIDAD"]
                fila_xl = ri + 3
                if cat in ("ALERTAS OPERACIONES", "NOVEDADES SAC",
                           "COMPOSICION DEL UNIVERSO BOG",
                           "FRANQUICIAS BOGOTA - Seguimiento de despachos"):
                    for ci in range(3):
                        ws_r.write(fila_xl, ci, row.iloc[ci], fmt_secc)
                elif cat.strip() == "":
                    pass
                elif qty != "" and pd.notna(qty):
                    try:
                        v = int(qty)
                        fmt_qty = fmt_rojo if v > 50 else (fmt_naran if v > 10 else fmt_verde)
                        ws_r.write(fila_xl, 0, row["CATEGORIA"], fmt_cel)
                        ws_r.write(fila_xl, 1, row["TIPO / MODALIDAD"], fmt_cel)
                        ws_r.write(fila_xl, 2, v, fmt_qty)
                    except Exception:
                        pass

            # Mini-resumenes en RESUMEN
            def _mini_bog(df_mini, fila_ini, titulo_mini):
                if df_mini.empty:
                    return
                ws_r.write(fila_ini, 0, titulo_mini,
                           _f(bold=True, bg_color="#2E75B6", font_color="#FFFFFF"))
                for ci, c in enumerate(df_mini.columns):
                    ws_r.write(fila_ini + 1, ci, c, fmt_enc)
                for ri2, row2 in df_mini.iterrows():
                    for ci2, val2 in enumerate(row2):
                        try:
                            ws_r.write(fila_ini + 2 + ri2, ci2,
                                       int(val2) if isinstance(val2, (int, float, np.integer))
                                           else val2,
                                       fmt_num if ci2 == len(df_mini.columns) - 1 else fmt_cel)
                        except Exception:
                            ws_r.write(fila_ini + 2 + ri2, ci2, str(val2), fmt_cel)

            fila_base = len(df_resumen) + 5
            for rsm, titulo_rsm in [
                (resumen_adm, "ADMITIDOS DEMORADOS - por Regional Origen"),
                (resumen_sac, "NOVEDADES SAC - por Regional Destino"),
                (resumen_ops, "OPS DEMORADAS - por Estado y Franquicia"),
            ]:
                _mini_bog(rsm, fila_base, titulo_rsm)
                fila_base += max(len(rsm), 0) + 4

        except Exception as _e_res:
            print(f"  ERROR en RESUMEN BOG: {_e_res}")

        # -- Pestana ADMITIDOS_DEMO -------------------------------------------
        _escribir_hoja(_hoja_bog, "ADMITIDOS_DEMO",
            df_adm,
            "ADMITIDOS DEMORADOS >2 dias  -  Franquicias BOG",
            COLOR_ADM,
            resumen_adm,
            "Resumen por Regional Origen")

        # -- Pestana NOVEDADES_SAC --------------------------------------------
        _escribir_hoja(_hoja_bog, "NOVEDADES_SAC",
            df_sac,
            "NOVEDADES SAC >1 dia habil  -  Franquicias BOG",
            COLOR_SAC,
            resumen_sac,
            "Resumen por Regional Destino")

        # -- Pestana OPS_DEMORADAS --------------------------------------------
        _escribir_hoja(_hoja_bog, "OPS_DEMORADAS",
            df_ops,
            "OPERACIONES DEMORADAS  -  Franquicias BOG",
            COLOR_OPS,
            resumen_ops,
            "Resumen por Estado y Franquicia")

    print(f"  OK ALERTAS_FRANQUICIAS_BOG guardado: {ruta_out.name}  [{time.time()-t0_xl:.1f}s]")
    print(f"     Universo Franquicias BOG: {n_base:,} guias")
    print(f"     ADMITIDOS_DEMO: {int(mask_adm.sum()):,} guias")
    print(f"     NOVEDADES_SAC:  {int(mask_sac.sum()):,} guias")
    print(f"     OPS_DEMORADAS:  {int(mask_ops_total.sum()):,} guias"
          f"  (fijo: {int(mask_ops_fijo.sum()):,} | "
          f"reparto: {int((mask_reparto_dir | mask_reparto_reex).sum()):,})")
    return ruta_out


def exportar(df: pd.DataFrame, carpeta_salida: Path) -> Path:
    sep("EXPORTANDO MAESTRO EXCEL")
    carpeta_salida.mkdir(parents=True, exist_ok=True)
    ts = pd.Timestamp.now().strftime("%Y%m%d_%H%M")

    ruta_xlsx = carpeta_salida / f"TRANSCARGA_MAESTRO_{ts}.xlsx"

    COLUMNAS_ENTERO_TIEMPO = [
        "TIEMPO DE ENTREGA", "T LEG CONTADO",
        "T CE CREACION A LEG", "T CE ENTREGA A LEG", "DIAS_EN_ESTADO",
        "TIEMPO_ENTREGA_NETO", "DIAS_EN_NOVEDAD",
        "SLA_DIAS", "DIAS_ATRASO", "HORA_CREACION",
        "IN_FULL", "IN_FULL_STRICT",
    ]
    # Columnas de fecha generadas por procesar_otif — se formatean igual que las demás
    COLUMNAS_FECHA_OTIF = [
        "FECHA_INICIO_SLA", "FECHA_PROMETIDA", "PROMESA_EXTENDIDA", "FECHA_INI_NOVEDAD",
    ]

    df_xl = df.copy()
    # Formatear fechas del maestro Y fechas OTIF (FECHA_PROMETIDA, etc.)
    todas_fechas = COLUMNAS_FECHA + (COLUMNAS_FECHA_OTIF if "COLUMNAS_FECHA_OTIF" in dir() else [
        "FECHA_INICIO_SLA", "FECHA_PROMETIDA", "PROMESA_EXTENDIDA", "FECHA_INI_NOVEDAD"
    ])
    for col in todas_fechas:
        if col in df_xl.columns and pd.api.types.is_datetime64_any_dtype(df_xl[col]):
            df_xl[col] = df_xl[col].dt.strftime("%Y-%m-%d")

    t0 = time.time()
    with pd.ExcelWriter(ruta_xlsx, engine="xlsxwriter") as writer:
        df_xl.to_excel(writer, index=False, sheet_name="DATOS")
        wb = writer.book
        ws = writer.sheets["DATOS"]

        fmt_enc = wb.add_format({
            "bold": True, "bg_color": "#1F3864", "font_color": "#FFFFFF",
            "font_name": "Calibri", "font_size": 10,
            "align": "center", "valign": "vcenter", "border": 0,
        })
        fmt_cel = wb.add_format({"font_name": "Calibri", "font_size": 10})
        fmt_cop = wb.add_format({"font_name": "Calibri", "font_size": 10, "num_format": "$#,##0"})
        fmt_int = wb.add_format({"font_name": "Calibri", "font_size": 10,
                                  "num_format": "#,##0", "align": "center"})
        cols_moneda_set = set(COLUMNAS_MONEDA)
        cols_tiempo_set = set(COLUMNAS_ENTERO_TIEMPO)
        # Columnas de texto OTIF — ancho estándar, sin formato numérico especial
        cols_otif_txt   = {"CUMPLE_OTIF", "ES_OFF_TIME", "CORREDOR_OTIF",
                            "TIPO_COBERTURA_2026", "CATEGORIA_ORIGEN",
                            "OTIF_AJUSTADO_PCT", "OTIF_ACIDO_PCT"}

        for i, col in enumerate(df_xl.columns):
            ws.write(0, i, col, fmt_enc)
            max_len = max(
                len(str(col)),
                (lambda s: int(s) if not pd.isna(s) else 10)(
                    df_xl[col].astype(str).str.len().quantile(0.95)
                ) if len(df_xl) > 0 else 10,
            )
            if col in cols_moneda_set:
                fmt_col, ancho = fmt_cop, 18
            elif col in cols_tiempo_set:
                fmt_col, ancho = fmt_int, 16
            else:
                fmt_col, ancho = fmt_cel, 45
            ws.set_column(i, i, min(max_len + 2, ancho), fmt_col)

        ws.freeze_panes(1, 0)
        ws.autofilter(0, 0, len(df_xl), len(df_xl.columns) - 1)

    t_excel = time.time() - t0
    print(f"  📊 Excel: {ruta_xlsx.name}  ({ruta_xlsx.stat().st_size/1e6:.2f} MB)  [{t_excel:.1f}s]")
    return ruta_xlsx


# =============================================================================
#  REPORTE DE DIAGNÓSTICO AMPLIADO
# =============================================================================

def reporte_diagnostico(df: pd.DataFrame, diagnosticos: list, carpeta_salida: Path):
    sep("REPORTE DE DIAGNÓSTICO")
    n = len(df)

    # Resumen por archivo
    print(f"\n  {'ARCHIVO':<42} {'REGISTROS':>10} {'NULOS_FECHA':>12} {'ANULADOS':>10}")
    print(f"  {'─'*42} {'─'*10} {'─'*12} {'─'*10}")
    for d in diagnosticos:
        if d["error"]:
            print(f"  ❌ {d['archivo']:<40} {'ERROR':>10}")
            continue
        pct_f = (f"{d['nulos_fecha_creacion']/max(d['registros_brutos'],1)*100:.1f}%"
                 if d["nulos_fecha_creacion"] > 0 else "✅ OK")
        icono = "⚠️ " if d["nulos_fecha_creacion"] > 0 else "✅"
        print(f"  {icono} {d['archivo']:<40} {d['registros_brutos']:>10,} "
              f"{pct_f:>12} {d['anulados']:>10,}")

    # Calidad por columna
    filas = []
    for col in df.columns:
        nulos  = df[col].isna().sum()
        unicos = df[col].nunique()
        ej = df[col].dropna().iloc[0] if df[col].notna().any() else ""
        pct = nulos / n * 100
        filas.append({
            "COLUMNA":        col,
            "TIPO":           str(df[col].dtype),
            "TOTAL":          n,
            "NULOS":          nulos,
            "PCT_NULOS":      f"{pct:.1f}%",
            "VALORES_UNICOS": unicos,
            "EJEMPLO":        str(ej)[:60],
            "ALERTA":         "⚠️ >50% NULOS" if pct > 50 else (
                               "⚠️ >20% NULOS" if pct > 20 else "OK"),
        })

    df_rep = pd.DataFrame(filas)

    print(f"\n  {'COL':<35} {'NULOS':>8} {'%':>8} {'ALERTA':>16}")
    print(f"  {'─'*35} {'─'*8} {'─'*8} {'─'*16}")
    for _, r in df_rep.iterrows():
        icono = "⚠️ " if r["ALERTA"] != "OK" else "✅"
        print(f"  {icono} {r['COLUMNA']:<33} {r['NULOS']:>8,} {r['PCT_NULOS']:>8} {r['ALERTA']:>16}")

    ruta_rep = carpeta_salida / "reporte_calidad.csv"
    df_rep.to_csv(ruta_rep, index=False, encoding="utf-8-sig")
    print(f"\n  💾 Reporte guardado: {ruta_rep.name}")

    alertas = df_rep[df_rep["ALERTA"] != "OK"]
    if not alertas.empty:
        print(f"\n  ⚠️  COLUMNAS QUE REQUIEREN REVISIÓN:")
        for _, r in alertas.iterrows():
            print(f"     → {r['COLUMNA']}: {r['PCT_NULOS']} de nulos — puede ser normal si esas guías no tienen ese dato.")
    else:
        print(f"\n  ✅ Todas las columnas dentro de rangos normales de nulos.")

# =============================================================================
#  OTIF (ON TIME IN FULL) - ACUMULADOR Y REPORTE
# =============================================================================



def procesar_otif(df: pd.DataFrame, carpeta_salida: Path, festivos: np.ndarray, malla_cubrimiento: dict = None) -> pd.DataFrame:
    """
    Calcula todas las métricas OTIF sobre el df y retorna el df enriquecido.
    
    Columnas que agrega:
      TIPO_COBERTURA_2026, CATEGORIA_ORIGEN, HORA_CREACION, ES_OFF_TIME,
      FECHA_INICIO_SLA, FECHA_PROMETIDA, PROMESA_EXTENDIDA, SLA_DIAS,
      CORREDOR_OTIF, DIAS_EN_NOVEDAD, FECHA_INI_NOVEDAD,
      CUMPLE_OTIF, DIAS_ATRASO, IN_FULL, IN_FULL_STRICT,
      TIEMPO_ENTREGA_NETO, OTIF_AJUSTADO_PCT, OTIF_ACIDO_PCT
    """
    sep("PROCESAMIENTO OTIF v7.0")
    archivo_historial = carpeta_salida / "HISTORICO_ESTADOS.csv"

    # ── 0. Validación mínima ──────────────────────────────────────────────────
    cols_req = ["GUIA", "NOMEST", "FECREGEST", "FECHA_DE_CREACION"]
    faltantes = [c for c in cols_req if c not in df.columns]
    if faltantes:
        print(f"  ❌ Faltan columnas requeridas para OTIF: {faltantes}")
        df["CUMPLE_OTIF"] = "SIN DATOS"
        return df

    n_total = len(df)
    print(f"  📊 Universo total: {n_total:,} guías")

    # ── 1. CLASIFICACIÓN DE COBERTURA ────────────────────────────────────────
    print("  📋 Clasificando cobertura (DIRECTO / REEXPEDICION)...")
    ori_n  = _norm_global(df["CIUDAD_ORIGEN"]  if "CIUDAD_ORIGEN"  in df.columns else pd.Series("", index=df.index))
    dest_n = _norm_global(df["CIUDAD_DESTINO"] if "CIUDAD_DESTINO" in df.columns else pd.Series("", index=df.index))

    if malla_cubrimiento:
        df["TIPO_COBERTURA_2026"] = [malla_cubrimiento.get((o, d), "REEXPEDICION")
                                      for o, d in zip(ori_n, dest_n)]
    elif "TIPO TRAYECTO" in df.columns:
        tray_norm = _norm_global(df["TIPO TRAYECTO"].astype(str))
        df["TIPO_COBERTURA_2026"] = np.where(
            tray_norm.str.contains("DIRECTO", na=False), "DIRECTO", "REEXPEDICION")
    else:
        df["TIPO_COBERTURA_2026"] = "REEXPEDICION"

    df["CATEGORIA_ORIGEN"] = np.where(ori_n.isin(CIUDADES_DIRECTAS_SET),
                                       "REGIONAL DIRECTA", "TERCEROS")
    n_dir  = (df["TIPO_COBERTURA_2026"] == "DIRECTO").sum()
    n_reex = (df["TIPO_COBERTURA_2026"] == "REEXPEDICION").sum()
    print(f"     DIRECTO: {n_dir:,}  |  REEXPEDICION: {n_reex:,}")

    # ── 2. HISTORICO_ESTADOS — SNAPSHOT DIARIO ───────────────────────────────
    # Guarda un registro diario por guía para calcular días de novedad históricos
    print("  🔄 Actualizando HISTORICO_ESTADOS...")
    _HIST_COLS = ["GUIA", "NOMEST", "NOMEST_NORM", "FECHA_CREACION",
                  "FECHA_PROCESO", "REGIONAL_ORIGEN", "REG DESTINO CORRECTA",
                  "TIPO TRAYECTO", "MES", "LINEA COMERCIAL"]

    col_fc = next((c for c in ["FECHA_DE_CREACION", "FECHA CREACION"] if c in df.columns), None)
    df_snap = df[["GUIA", "NOMEST"]].dropna(subset=["GUIA", "NOMEST"]).copy()
    df_snap["NOMEST_NORM"]    = _norm_global(df_snap["NOMEST"])
    df_snap["FECHA_CREACION"] = pd.to_datetime(df[col_fc], errors="coerce") if col_fc else pd.NaT
    df_snap["FECHA_PROCESO"]  = pd.Timestamp.now().normalize()
    for col in ["REGIONAL_ORIGEN", "REG DESTINO CORRECTA", "TIPO TRAYECTO", "MES", "LINEA COMERCIAL"]:
        df_snap[col] = df[col].values if col in df.columns else ""
    df_snap = df_snap.reindex(columns=_HIST_COLS).fillna("")

    try:
        if archivo_historial.exists():
            df_hist = pd.read_csv(archivo_historial, dtype=str)
            for c in _HIST_COLS:
                if c not in df_hist.columns:
                    df_hist[c] = ""
            df_hist_final = pd.concat(
                [df_hist[_HIST_COLS], df_snap.astype(str)], ignore_index=True
            ).drop_duplicates(subset=["GUIA", "NOMEST_NORM", "FECHA_PROCESO"], keep="last")
            print(f"     Historial acumulado: {len(df_hist_final):,} snapshots")
        else:
            df_hist_final = df_snap.astype(str).drop_duplicates(
                subset=["GUIA", "NOMEST_NORM", "FECHA_PROCESO"], keep="last")
            print(f"     Primer snapshot: {len(df_hist_final):,} registros")

        df_hist_final.to_csv(archivo_historial, index=False, encoding="utf-8-sig")
        try:
            import ctypes
            ctypes.windll.kernel32.SetFileAttributesW(str(archivo_historial), 0x02)
        except Exception:
            pass
    except Exception as e_hist:
        print(f"     ⚠️  Error en historial ({e_hist}) — continuando sin días de novedad")
        df_hist_final = df_snap.astype(str)

    # ── 3. DÍAS DE EXCLUSIÓN (novedades que paran el reloj SLA) ──────────────
    print("  ⏱️  Calculando días de exclusión...")
    estados_excl_norm = {_norm_global(pd.Series([e])).iloc[0] for e in ESTADOS_EXCLUSION_OTIF}

    primera_novedad = pd.DataFrame(columns=["GUIA", "FECHA_INI_NOVEDAD", "DIAS_EN_NOVEDAD"])
    try:
        df_h = df_hist_final.copy()
        df_h["FECHA_PROCESO_DT"] = pd.to_datetime(df_h["FECHA_PROCESO"], errors="coerce")
        df_h = df_h.dropna(subset=["GUIA", "FECHA_PROCESO_DT"])
        df_h["EN_EXCL"] = df_h["NOMEST_NORM"].isin(estados_excl_norm)

        exclusion_rows = []
        for guia, grp in df_h.groupby("GUIA"):
            grp = grp.sort_values("FECHA_PROCESO_DT").reset_index(drop=True)
            excl_grp = grp[grp["EN_EXCL"]]
            if excl_grp.empty:
                continue
            post_excl = grp[
                (~grp["EN_EXCL"]) &
                (grp["FECHA_PROCESO_DT"] > excl_grp["FECHA_PROCESO_DT"].max())
            ]
            if post_excl.empty:
                continue
            ini_np = excl_grp["FECHA_PROCESO_DT"].min().to_datetime64().astype("datetime64[D]")
            sal_np = post_excl["FECHA_PROCESO_DT"].min().to_datetime64().astype("datetime64[D]")
            dias_excl = int(np.maximum(
                np.busday_count(ini_np, sal_np,
                                weekmask="Mon Tue Wed Thu Fri Sat", holidays=festivos), 0))
            exclusion_rows.append({"GUIA": guia,
                                    "FECHA_INI_NOVEDAD": excl_grp["FECHA_PROCESO_DT"].min(),
                                    "DIAS_EN_NOVEDAD":   dias_excl})
        if exclusion_rows:
            primera_novedad = pd.DataFrame(exclusion_rows)
        print(f"     Guías con exclusión resuelta: {len(primera_novedad):,}")
    except Exception as e_excl:
        print(f"     ⚠️  Error calculando exclusiones ({e_excl}) — DIAS_EN_NOVEDAD = 0 para todas")

    df = df.merge(primera_novedad, on="GUIA", how="left")
    df["DIAS_EN_NOVEDAD"]   = pd.to_numeric(df.get("DIAS_EN_NOVEDAD"),   errors="coerce").fillna(0).astype(int)
    df["FECHA_INI_NOVEDAD"] = pd.to_datetime(df.get("FECHA_INI_NOVEDAD"), errors="coerce")

    # ── 4. OFF-TIME y SLA ─────────────────────────────────────────────────────
    print(f"  🕐 Calculando OFF-TIME (ventana {OFF_TIME_START}h–{OFF_TIME_END}h) y SLA...")
    fecha_crea_dt    = pd.to_datetime(df["FECHA_DE_CREACION"], errors="coerce")
    mask_hora_valida = fecha_crea_dt.notna()
    df["HORA_CREACION"] = fecha_crea_dt.dt.hour.fillna(-1).astype(int)
    df["ES_OFF_TIME"]   = np.where(
        ~mask_hora_valida, "SIN HORA",
        np.where(
            (df["HORA_CREACION"] < OFF_TIME_START) | (df["HORA_CREACION"] >= OFF_TIME_END),
            "SÍ", "NO"
        )
    )
    n_oft = (df["ES_OFF_TIME"] == "SÍ").sum()
    n_ont = (df["ES_OFF_TIME"] == "NO").sum()
    print(f"     OFF-TIME: {n_oft:,}  |  ON-TIME: {n_ont:,}  |  SIN HORA: {(df['ES_OFF_TIME']=='SIN HORA').sum():,}")

    # Refrescar ori_n/dest_n post-merge
    ori_n  = _norm_global(df["CIUDAD_ORIGEN"]  if "CIUDAD_ORIGEN"  in df.columns else pd.Series("", index=df.index))
    dest_n = _norm_global(df["CIUDAD_DESTINO"] if "CIUDAD_DESTINO" in df.columns else pd.Series("", index=df.index))

    cond_regional = ori_n.isin(CIUDADES_REGIONALES_SET) & dest_n.isin(CIUDADES_REGIONALES_SET)
    cond_med_buc  = (ori_n == "MEDELLIN") & dest_n.isin(REGIONAL_BUC)
    cond_reexp    = (df["TIPO_COBERTURA_2026"] == "REEXPEDICION")

    df["SLA_DIAS"] = np.where(cond_reexp, 7,
                    np.where(cond_regional | cond_med_buc, 2, 5))
    df["CORREDOR_OTIF"] = np.where(
        cond_reexp,    "REEXPEDICION / ESPECIAL",
        np.where(cond_regional, "REGIONAL DIRECTO (EJE-MED-CALI)",
        np.where(cond_med_buc,  "MEDELLIN → BUCARAMANGA (DIRECTO)",
                                "DIRECTO NACIONAL (VIA BOG)")))

    # ── 5. FECHA_INICIO_SLA vectorizada ──────────────────────────────────────
    shifts         = np.where(df["ES_OFF_TIME"] == "SÍ", 2, 1)
    mask_fecha_ok  = fecha_crea_dt.notna()
    dias_crea_np   = fecha_crea_dt[mask_fecha_ok].values.astype("datetime64[D]")
    inicio_sla_arr = np.full(len(df), np.datetime64("NaT", "D"), dtype="datetime64[D]")
    ok_pos         = np.where(mask_fecha_ok)[0]
    for pos, (dia, shift) in enumerate(zip(dias_crea_np, shifts[mask_fecha_ok])):
        try:
            inicio_sla_arr[ok_pos[pos]] = np.busday_offset(
                dia, int(shift), weekmask="Mon Tue Wed Thu Fri Sat",
                holidays=festivos, roll="forward")
        except Exception:
            pass
    df["FECHA_INICIO_SLA"] = pd.to_datetime(inicio_sla_arr)

    # ── 6. FECHA_PROMETIDA vectorizada ────────────────────────────────────────
    mask_sla_ok   = df["FECHA_INICIO_SLA"].notna()
    prometida_arr = np.full(len(df), np.datetime64("NaT", "D"), dtype="datetime64[D]")
    sla_array     = df["SLA_DIAS"].values
    ini_sla_np    = df.loc[mask_sla_ok, "FECHA_INICIO_SLA"].values.astype("datetime64[D]")
    for pos, (idx_row, dia, sla) in enumerate(
        zip(df.index[mask_sla_ok], ini_sla_np, sla_array[mask_sla_ok])
    ):
        try:
            prometida_arr[idx_row] = np.busday_offset(
                dia, int(sla) - 1, weekmask="Mon Tue Wed Thu Fri Sat",
                holidays=festivos, roll="forward")
        except Exception:
            pass
    df["FECHA_PROMETIDA"] = pd.to_datetime(prometida_arr)

    # ── 7. PROMESA_EXTENDIDA (ajustada por días de novedad) ───────────────────
    nomest_n       = _norm_global(df["NOMEST"])
    mask_entregado = (nomest_n == _ENTREGADO_N)
    fecregest_dt   = pd.to_datetime(df["FECREGEST"], errors="coerce").dt.normalize()
    mask_fechas_ok = fecregest_dt.notna() & df["FECHA_PROMETIDA"].notna()
    dias_nov       = pd.to_numeric(df["DIAS_EN_NOVEDAD"], errors="coerce").fillna(0).astype(int)

    promesa_ext_arr = df["FECHA_PROMETIDA"].values.copy().astype("datetime64[D]")
    mask_con_nov    = mask_entregado & mask_fechas_ok & (dias_nov > 0)
    if mask_con_nov.any():
        prom_base_np = df.loc[mask_con_nov, "FECHA_PROMETIDA"].values.astype("datetime64[D]")
        nov_arr      = dias_nov[mask_con_nov].values
        for pos, (idx_row, p, d) in enumerate(
            zip(df.index[mask_con_nov], prom_base_np, nov_arr)
        ):
            try:
                promesa_ext_arr[idx_row] = np.busday_offset(
                    p, int(d), weekmask="Mon Tue Wed Thu Fri Sat",
                    holidays=festivos, roll="forward")
            except Exception:
                pass
    df["PROMESA_EXTENDIDA"] = pd.to_datetime(promesa_ext_arr)

    # ── 8. CUMPLE_OTIF + DIAS_ATRASO ─────────────────────────────────────────
    entrega_np    = fecregest_dt.values.astype("datetime64[D]")
    prometida_np  = df["PROMESA_EXTENDIDA"].values.astype("datetime64[D]")
    on_time_bool  = (entrega_np <= prometida_np)

    df["CUMPLE_OTIF"] = np.where(
        ~mask_entregado,   "NO ENTREGADO",
        np.where(~mask_fechas_ok, "SIN DATOS",
        np.where(on_time_bool,    "SI - ON TIME", "NO - LATE"))
    )

    mask_late  = df["CUMPLE_OTIF"] == "NO - LATE"
    atraso_arr = np.zeros(len(df), dtype=np.int32)
    if mask_late.any():
        atraso_vals = np.maximum(
            np.busday_count(prometida_np[mask_late.values], entrega_np[mask_late.values],
                            weekmask="Mon Tue Wed Thu Fri Sat", holidays=festivos), 0
        ).astype(np.int32)
        atraso_arr[mask_late.values] = atraso_vals
    df["DIAS_ATRASO"] = atraso_arr

    # ── 9. IN_FULL, IN_FULL_STRICT ────────────────────────────────────────────
    mask_on_time_bool    = (df["CUMPLE_OTIF"] == "SI - ON TIME")
    df["IN_FULL"]        = np.where(mask_entregado, 1, 0)
    df["IN_FULL_STRICT"] = np.where(mask_on_time_bool & mask_entregado, 1, 0)

    # ── 10. TIEMPO_ENTREGA_NETO ───────────────────────────────────────────────
    if "TIEMPO DE ENTREGA" in df.columns:
        te_col  = pd.to_numeric(df["TIEMPO DE ENTREGA"], errors="coerce")
        nov_col = pd.to_numeric(df["DIAS_EN_NOVEDAD"],   errors="coerce").fillna(0)
        df["TIEMPO_ENTREGA_NETO"] = (te_col - nov_col).clip(lower=0)
    else:
        df["TIEMPO_ENTREGA_NETO"] = pd.NA

    # ── 11. Métricas ÁCIDO y AJUSTADO ─────────────────────────────────────────
    n_on_time   = int((df["CUMPLE_OTIF"] == "SI - ON TIME").sum())
    n_late      = int((df["CUMPLE_OTIF"] == "NO - LATE").sum())
    n_sin_dat   = int((df["CUMPLE_OTIF"] == "SIN DATOS").sum())
    n_no_ent    = int((df["CUMPLE_OTIF"] == "NO ENTREGADO").sum())
    n_in_full   = int(df["IN_FULL"].sum())
    n_if_strict = int(df["IN_FULL_STRICT"].sum())

    # AJUSTADO: denominador = solo guías con ciclo cerrado (on_time + late)
    otif_ajustado = n_on_time / (n_on_time + n_late) * 100 if (n_on_time + n_late) > 0 else 0.0
    # ÁCIDO:    denominador = TODAS las guías del período
    otif_acido    = n_if_strict / n_total * 100 if n_total > 0 else 0.0

    df["OTIF_AJUSTADO_PCT"] = round(otif_ajustado, 2)
    df["OTIF_ACIDO_PCT"]    = round(otif_acido,    2)

    print(f"\n  ✅ OTIF v7.0:")
    print(f"     SI - ON TIME:   {n_on_time:>8,}")
    print(f"     NO - LATE:      {n_late:>8,}")
    print(f"     SIN DATOS:      {n_sin_dat:>8,}")
    print(f"     NO ENTREGADO:   {n_no_ent:>8,}")
    print(f"     IN FULL:        {n_in_full:>8,}")
    print(f"     IN FULL STRICT: {n_if_strict:>8,}")
    print(f"  ───────────────────────────────────────────────────")
    print(f"     OTIF AJUSTADO:  {otif_ajustado:>7.1f}%  [on_time / (on_time+late)]")
    print(f"     OTIF ÁCIDO:     {otif_acido:>7.1f}%  [on_time_in_full / total_guías]")
    return df


def generar_reporte_otif(df: pd.DataFrame, carpeta_salida: Path) -> Path:
    """
    Genera REPORTE_OTIF_GERENCIAL_<ts>.xlsx con pestañas:
      GLOSARIO | DASHBOARD_GERENCIAL | TENDENCIA_MENSUAL
      DETALLE_REGIONALES | DETALLE_DIRECTO | DETALLE_REEXPEDIDO
      M_DIR_<mes> | M_REEX_<mes>  (matrices heatmap por mes)
      AUDITORIA_NOVEDADES | ANALISIS_OFF_TIME | IN_FULL_ANALISIS
      BASE_DETALLE
    """
    sep("GENERANDO REPORTE OTIF GERENCIAL v7.0")

    # ── Guardia de entrada ────────────────────────────────────────────────────
    if "CUMPLE_OTIF" not in df.columns:
        print("  ❌ CUMPLE_OTIF no existe en el df — ejecuta procesar_otif() primero.")
        return None

    df_ent = df[df["CUMPLE_OTIF"].isin(["SI - ON TIME", "NO - LATE"])].copy()
    if df_ent.empty:
        print("  ⚠️  No hay guías SI - ON TIME ni NO - LATE. Sin datos para el reporte.")
        return None

    print(f"  📋 df_ent: {len(df_ent):,} guías  |  df total: {len(df):,} guías")

    ts      = pd.Timestamp.now().strftime("%Y%m%d_%H%M")
    hoy_str = pd.Timestamp.now().strftime("%d/%m/%Y %H:%M")
    ruta_out = carpeta_salida / f"REPORTE_OTIF_GERENCIAL_{ts}.xlsx"

    # ── Preparación ───────────────────────────────────────────────────────────
    df_ent["ES_ON_TIME"]  = (df_ent["CUMPLE_OTIF"] == "SI - ON TIME").astype(int)
    df_ent["ES_LATE"]     = (df_ent["CUMPLE_OTIF"] == "NO - LATE").astype(int)
    df_ent["TOTAL_GUIAS"] = 1

    _GRP_COLS = ["MES", "REGIONAL_ORIGEN", "REG DESTINO CORRECTA",
                 "CIUDAD_ORIGEN", "CIUDAD_DESTINO", "CORREDOR_OTIF", "TIPO_COBERTURA_2026"]
    for col in _GRP_COLS:
        if col in df_ent.columns:
            df_ent[col] = df_ent[col].fillna("SIN DATO").astype(str)
        else:
            df_ent[col] = "SIN DATO"
        if col in df.columns:
            df[col] = df[col].fillna("SIN DATO").astype(str)
        else:
            df[col] = "SIN DATO"

    # ── KPIs globales ─────────────────────────────────────────────────────────
    n_total_periodo = len(df)
    n_total_ent     = len(df_ent)
    n_on_time       = int(df_ent["ES_ON_TIME"].sum())
    n_late          = int(df_ent["ES_LATE"].sum())
    n_no_ent        = int((df["CUMPLE_OTIF"] == "NO ENTREGADO").sum())
    n_sin_dat       = int((df["CUMPLE_OTIF"] == "SIN DATOS").sum())

    _col_oft     = df_ent.get("ES_OFF_TIME", pd.Series("NO", index=df_ent.index))
    n_off_time   = int((_col_oft == "SÍ").sum())

    in_full_total  = int(df_ent["IN_FULL"].sum())        if "IN_FULL"        in df_ent.columns else n_on_time
    in_full_strict = int(df_ent["IN_FULL_STRICT"].sum()) if "IN_FULL_STRICT" in df_ent.columns else n_on_time

    # AJUSTADO: on_time / (on_time + late)  — denominador = ciclos cerrados
    pct_otif_ajustado = n_on_time  / n_total_ent     if n_total_ent     > 0 else 0.0
    # ÁCIDO:    on_time_in_full / total_guías — denominador = universo completo
    pct_otif_acido    = in_full_strict / n_total_periodo if n_total_periodo > 0 else 0.0
    pct_in_full       = in_full_total  / n_total_ent     if n_total_ent     > 0 else 0.0
    pct_off_time      = n_off_time     / n_total_ent     if n_total_ent     > 0 else 0.0
    pct_otif_completo = pct_otif_acido  # alias para compatibilidad

    sla_prom    = df_ent["SLA_DIAS"].mean()    if "SLA_DIAS"    in df_ent.columns else 0
    atraso_prom = df_ent[df_ent.get("DIAS_ATRASO", pd.Series(0, index=df_ent.index)) > 0]["DIAS_ATRASO"].mean() \
                  if "DIAS_ATRASO" in df_ent.columns else 0

    # ── Agregaciones ─────────────────────────────────────────────────────────
    res_base = df_ent.groupby(["REGIONAL_ORIGEN", "REG DESTINO CORRECTA", "TIPO_COBERTURA_2026"]).agg(
        VOL=("TOTAL_GUIAS", "sum"), ON_TIME=("ES_ON_TIME", "sum"), LATE=("ES_LATE", "sum")
    ).reset_index()
    res_base["% OTIF"] = res_base["ON_TIME"] / res_base["VOL"]

    res_mes = df_ent.groupby("MES").agg(
        VOL=("TOTAL_GUIAS", "sum"), ON_TIME=("ES_ON_TIME", "sum")
    ).reset_index()
    res_mes["% OTIF"] = res_mes["ON_TIME"] / res_mes["VOL"]

    criticas  = res_base[(res_base["VOL"] >= 10) & (res_base["% OTIF"] < 0.85)].sort_values("VOL", ascending=False).head(10)
    comp_tipo = df_ent.groupby("TIPO_COBERTURA_2026").agg(
        VOL=("TOTAL_GUIAS", "sum"), ON_TIME=("ES_ON_TIME", "sum")).reset_index()
    comp_tipo["% OTIF"] = comp_tipo["ON_TIME"] / comp_tipo["VOL"]

    print(f"  📊 Escribiendo Excel OTIF v7.0...")
    with pd.ExcelWriter(ruta_out, engine="xlsxwriter") as writer:
        wb = writer.book
        def _f(**kw):
            return wb.add_format({"font_name": "Segoe UI", "font_size": 10, **kw})

        # Formatos base
        fmt_kpi_lbl  = _f(bold=True, font_size=11, align="center", bg_color="#1F3864", font_color="#FFFFFF")
        fmt_kpi_val  = _f(bold=True, font_size=22, align="center", bg_color="#D9E1F2", num_format="#,##0")
        fmt_kpi_pct  = _f(bold=True, font_size=22, align="center", bg_color="#D9E1F2", num_format="0.0%")
        fmt_kpi_ajust= _f(bold=True, font_size=22, align="center", bg_color="#E2EFDA", num_format="0.0%")
        fmt_kpi_acido= _f(bold=True, font_size=22, align="center", bg_color="#FFF2CC", num_format="0.0%")
        fmt_header   = _f(bold=True, bg_color="#1F3864", font_color="#FFFFFF", border=1, align="center")
        fmt_title    = _f(bold=True, font_size=14, bg_color="#1F3864", font_color="#FFFFFF", align="center")
        fmt_subtitle = _f(bold=True, font_size=11, bg_color="#D9E1F2", border=1)
        fmt_pct      = _f(num_format="0.0%", align="center", border=1)
        fmt_num      = _f(num_format="#,##0", align="center", border=1)
        fmt_txt      = _f(align="left", border=1)

        # ──────────────────────────────────────────────────────────────────────
        # PESTAÑA 0: GLOSARIO
        # ──────────────────────────────────────────────────────────────────────
        ws_glo = wb.add_worksheet("GLOSARIO")
        ws_glo.set_column("B:B", 30)
        ws_glo.set_column("C:C", 90)
        ws_glo.write("B2", "TÉRMINO", fmt_header)
        ws_glo.write("C2", "DEFINICIÓN", fmt_header)
        glosario = [
            ("OTIF AJUSTADO (Soft/Operativo)",
             "Denominador = solo guías entregadas (on_time + late). "
             "Fórmula: on_time / (on_time + late). "
             "Mide eficiencia operativa sobre ciclos cerrados. No penaliza guías en tránsito."),
            ("OTIF ÁCIDO (Hard/Ejecutivo)",
             "Denominador = TODAS las guías del período (incluye pendientes y no entregadas). "
             "Fórmula: on_time_in_full / total_guías. "
             "KPI más exigente. Refleja el servicio real percibido por el cliente."),
            ("IN FULL STRICT",
             "Guía ON TIME Y entregada en satisfacción (intersección real). Binaria 0/1."),
            ("ON TIME",
             "Entrega en o antes de PROMESA_EXTENDIDA (fecha prometida + días de novedad excluidos)."),
            ("OFF-TIME",
             f"Guía creada antes de las {OFF_TIME_START}h o a las {OFF_TIME_END}h o después. "
             "Su SLA inicia 2 días hábiles después (vs 1 día para ON-TIME)."),
            ("SLA por corredor",
             "Reexpedición: 7d | Regional Directo (Eje-Med-Cali / Med-Buc): 2d | "
             "Directo Nacional vía BOG: 5d. Días hábiles Lun-Sáb sin festivos Colombia."),
            ("PROMESA_EXTENDIDA",
             "Fecha prometida + días de novedades válidas (dirección errada, destinatario "
             "desconocido, en novedad, cerrado). Estas exclusiones paran el reloj del SLA."),
            ("TIEMPO_ENTREGA_NETO",
             "TIEMPO DE ENTREGA − DIAS_EN_NOVEDAD. Tránsito efectivo sin tiempos de exclusión."),
            ("DIAS_ATRASO",
             "Días hábiles entre PROMESA_EXTENDIDA y FECREGEST para guías LATE."),
        ]
        for i, (k, v) in enumerate(glosario):
            ws_glo.write(i + 3, 1, k, _f(bold=True, border=1))
            ws_glo.write(i + 3, 2, v, fmt_txt)

        # ──────────────────────────────────────────────────────────────────────
        # PESTAÑA 1: DASHBOARD GERENCIAL
        # ──────────────────────────────────────────────────────────────────────
        ws_gen = wb.add_worksheet("DASHBOARD_GERENCIAL")
        ws_gen.hide_gridlines(2)
        ws_gen.set_column("B:I", 22)
        ws_gen.merge_range("B2:I3", f"DASHBOARD OTIF EJECUTIVO v7.0 — {hoy_str}", fmt_title)

        # Fila 1 KPIs: volumen
        ws_gen.write("B5", "TOTAL GUÍAS",   fmt_kpi_lbl); ws_gen.write("B6", n_total_periodo, fmt_kpi_val)
        ws_gen.write("C5", "ENTREGADAS",    fmt_kpi_lbl); ws_gen.write("C6", n_total_ent,     fmt_kpi_val)
        ws_gen.write("D5", "ON TIME",       fmt_kpi_lbl); ws_gen.write("D6", n_on_time,       fmt_kpi_val)
        ws_gen.write("E5", "LATE",          fmt_kpi_lbl); ws_gen.write("E6", n_late,          fmt_kpi_val)
        ws_gen.write("F5", "NO ENTREGADAS", fmt_kpi_lbl); ws_gen.write("F6", n_no_ent,        fmt_kpi_val)
        ws_gen.write("G5", "IN FULL",       fmt_kpi_lbl); ws_gen.write("G6", in_full_total,   fmt_kpi_val)
        ws_gen.write("H5", "OFF-TIME",      fmt_kpi_lbl); ws_gen.write("H6", n_off_time,      fmt_kpi_val)
        ws_gen.write("I5", "% OFF-TIME",    fmt_kpi_lbl); ws_gen.write("I6", pct_off_time,    fmt_kpi_pct)
        ws_gen.set_row(4, 20); ws_gen.set_row(5, 48)

        # Fila 2 KPIs: ÁCIDO vs AJUSTADO — celdas grandes y diferenciadas por color
        ws_gen.merge_range("B8:D8", "OTIF ÁCIDO  🔬  (Hard / Ejecutivo)", fmt_kpi_lbl)
        ws_gen.merge_range("F8:I8", "OTIF AJUSTADO  ✅  (Soft / Operativo)", fmt_kpi_lbl)
        ws_gen.merge_range("B9:D11", pct_otif_acido,    _f(bold=True, font_size=36, align="center", valign="vcenter", bg_color="#FFF2CC", num_format="0.0%", border=2))
        ws_gen.merge_range("F9:I11", pct_otif_ajustado, _f(bold=True, font_size=36, align="center", valign="vcenter", bg_color="#E2EFDA", num_format="0.0%", border=2))
        ws_gen.write("B12", f"on_time_in_full / total  ({in_full_strict:,} / {n_total_periodo:,})",
                     _f(italic=True, font_size=8, align="center", font_color="#595959", bg_color="#FFF2CC"))
        ws_gen.merge_range("C12:D12", "", _f(bg_color="#FFF2CC"))
        ws_gen.write("F12", f"on_time / (on_time+late)  ({n_on_time:,} / {n_total_ent:,})",
                     _f(italic=True, font_size=8, align="center", font_color="#595959", bg_color="#E2EFDA"))
        ws_gen.merge_range("G12:I12", "", _f(bg_color="#E2EFDA"))
        ws_gen.set_row(7, 22); ws_gen.set_row(8, 30); ws_gen.set_row(9, 30); ws_gen.set_row(10, 30); ws_gen.set_row(11, 18)

        # Fila 3: KPIs secundarios
        fmt_sec_lbl = _f(bold=True, font_size=9, align="center", bg_color="#2F5496", font_color="#FFFFFF")
        fmt_sec_val = _f(bold=True, font_size=13, align="center", bg_color="#EEF3FB", num_format="#,##0")
        fmt_sec_pct = _f(bold=True, font_size=13, align="center", bg_color="#EEF3FB", num_format="0.0%")
        fmt_sec_dec = _f(bold=True, font_size=13, align="center", bg_color="#EEF3FB", num_format="0.0")
        ws_gen.write("B14", "% IN FULL",          fmt_sec_lbl); ws_gen.write("B15", pct_in_full,  fmt_sec_pct)
        ws_gen.write("C14", "SLA PROM (días)",     fmt_sec_lbl); ws_gen.write("C15", round(sla_prom, 1), fmt_sec_dec)
        ws_gen.write("D14", "ATRASO PROM (días)",  fmt_sec_lbl); ws_gen.write("D15", round(atraso_prom if not pd.isna(atraso_prom) else 0, 1), _f(bold=True, font_size=13, align="center", bg_color="#FFC7CE", num_format="0.0"))
        ws_gen.write("E14", "ON TIME + IN FULL",   fmt_sec_lbl); ws_gen.write("E15", in_full_strict, fmt_sec_val)
        ws_gen.write("F14", "SIN DATOS",           fmt_sec_lbl); ws_gen.write("F15", n_sin_dat,    fmt_sec_val)
        ws_gen.set_row(13, 20); ws_gen.set_row(14, 40)

        # TOP 10 Rutas Críticas
        if not criticas.empty:
            ws_gen.merge_range("B17:I17", "🚩  TOP 10 RUTAS CRÍTICAS  (VOL ≥ 10  y  OTIF < 85%)", fmt_subtitle)
            ws_gen.write_row("B18", ["ORIGEN", "DESTINO", "TIPO", "VOL", "ON TIME", "LATE", "% OTIF"], fmt_header)
            for i, (_, r) in enumerate(criticas.iterrows()):
                bg = "#FFC7CE" if r["% OTIF"] < 0.7 else "#FFEB9C"
                ws_gen.write(18 + i, 1, r["REGIONAL_ORIGEN"],      fmt_txt)
                ws_gen.write(18 + i, 2, r["REG DESTINO CORRECTA"], fmt_txt)
                ws_gen.write(18 + i, 3, r["TIPO_COBERTURA_2026"],  fmt_txt)
                ws_gen.write(18 + i, 4, r["VOL"],    fmt_num)
                ws_gen.write(18 + i, 5, r["ON_TIME"], fmt_num)
                ws_gen.write(18 + i, 6, r["LATE"],    fmt_num)
                ws_gen.write(18 + i, 7, r["% OTIF"],  _f(border=1, num_format="0.0%", bold=True, bg_color=bg))

        # ──────────────────────────────────────────────────────────────────────
        # PESTAÑA 2: TENDENCIA MENSUAL (si hay más de 1 mes)
        # ──────────────────────────────────────────────────────────────────────
        if len(res_mes) > 1:
            res_mes.to_excel(writer, sheet_name="TENDENCIA_MENSUAL", index=False, startrow=3)
            ws_tr = writer.sheets["TENDENCIA_MENSUAL"]
            ws_tr.merge_range(0, 0, 1, len(res_mes.columns) - 1, "EVOLUCIÓN MENSUAL OTIF", fmt_title)
            chart_t = wb.add_chart({"type": "line"})
            chart_t.add_series({
                "name": "% OTIF Ajustado",
                "categories": ["TENDENCIA_MENSUAL", 4, 0, 3 + len(res_mes), 0],
                "values":     ["TENDENCIA_MENSUAL", 4, 3, 3 + len(res_mes), 3],
                "marker": {"type": "circle", "size": 8,
                           "border": {"color": "#1F3864"}, "fill": {"color": "#FFFFFF"}},
            })
            chart_t.set_title({"name": "Evolución Mensual OTIF Ajustado"})
            chart_t.set_y_axis({"major_unit": 0.05, "num_format": "0%"})
            ws_tr.insert_chart("F2", chart_t, {"x_scale": 1.5, "y_scale": 1.2})

        # ──────────────────────────────────────────────────────────────────────
        # PESTAÑAS 3-5: DETALLES REGIONALES / DIRECTO / REEXPEDIDO
        # ──────────────────────────────────────────────────────────────────────
        def _escribir_det(nombre, df_det, titulo):
            if df_det.empty:
                return
            df_det.to_excel(writer, sheet_name=nombre, index=False, startrow=3)
            ws_d = writer.sheets[nombre]
            ws_d.merge_range(0, 0, 1, len(df_det.columns) - 1, titulo, fmt_title)
            for c, col in enumerate(df_det.columns):
                ws_d.write(3, c, col, fmt_header)
                ancho = 15 if col in ("VOL", "ON_TIME", "LATE", "OTIF") else 22
                ws_d.set_column(c, c, ancho)

        _escribir_det("DETALLE_REGIONALES",
                      res_base.groupby(["REGIONAL_ORIGEN", "REG DESTINO CORRECTA"]).agg(
                          VOL=("VOL", "sum"), ON_TIME=("ON_TIME", "sum")).reset_index().assign(
                          OTIF=lambda x: x["ON_TIME"] / x["VOL"]),
                      "INDICADORES POR CORREDOR REGIONAL")
        _escribir_det("DETALLE_DIRECTO",
                      res_base[res_base["TIPO_COBERTURA_2026"] == "DIRECTO"],
                      "DETALLE DIRECTO")
        _escribir_det("DETALLE_REEXPEDIDO",
                      res_base[res_base["TIPO_COBERTURA_2026"] == "REEXPEDICION"],
                      "DETALLE REEXPEDIDO")

        # ──────────────────────────────────────────────────────────────────────
        # PESTAÑAS MATRICES HEATMAP (una por mes × tipo)
        # ──────────────────────────────────────────────────────────────────────
        for mes_idx in sorted(df_ent["MES"].unique()):
            mes_short = (mes_idx.split(".")[1].strip()[:3].upper()
                         if "." in mes_idx else mes_idx[:3].upper())
            for tipo in ["DIRECTO", "REEXPEDICION"]:
                sub_m = df_ent[(df_ent["MES"] == mes_idx) & (df_ent["TIPO_COBERTURA_2026"] == tipo)]
                if sub_m.empty:
                    continue
                mtx = sub_m.pivot_table(
                    index="REGIONAL_ORIGEN", columns="REG DESTINO CORRECTA",
                    values="ES_ON_TIME", aggfunc="mean")
                sh_name = f"M_{'DIR' if tipo == 'DIRECTO' else 'REEX'}_{mes_short}"[:31]
                mtx.to_excel(writer, sheet_name=sh_name, startrow=3)
                ws_m = writer.sheets[sh_name]
                ws_m.merge_range(0, 0, 1, len(mtx.columns),
                                 f"MATRIZ % OTIF {tipo} — {mes_idx}", fmt_title)
                for r_i in range(len(mtx)):
                    for c_i in range(len(mtx.columns)):
                        val = mtx.iloc[r_i, c_i]
                        if not pd.isna(val):
                            ws_m.write(r_i + 4, c_i + 1, val, fmt_pct)
                        else:
                            ws_m.write(r_i + 4, c_i + 1, "", _f(bg_color="#F2F2F2", border=1))
                ws_m.conditional_format(4, 1, 3 + len(mtx), len(mtx.columns), {
                    "type": "3_color_scale",
                    "min_color": "#FF7C80", "mid_color": "#FFEB84", "max_color": "#C6EFCE",
                    "min_type": "num", "min_value": 0, "max_type": "num", "max_value": 1
                })
                ws_m.set_column(0, 0, 25)
                ws_m.set_column(1, len(mtx.columns), 15)

        # ──────────────────────────────────────────────────────────────────────
        # AUDITORÍA NOVEDADES
        # ──────────────────────────────────────────────────────────────────────
        _cols_aud = [c for c in ["GUIA", "CIUDAD_ORIGEN", "CIUDAD_DESTINO",
                                  "REGIONAL_ORIGEN", "REG DESTINO CORRECTA", "NOMEST",
                                  "FECREGEST", "FECHA_INI_NOVEDAD", "DIAS_EN_NOVEDAD",
                                  "FECHA_PROMETIDA", "PROMESA_EXTENDIDA", "CUMPLE_OTIF",
                                  "CORREDOR_OTIF", "SLA_DIAS"] if c in df_ent.columns]
        df_aud = df_ent[
            pd.to_numeric(df_ent.get("DIAS_EN_NOVEDAD", 0), errors="coerce").fillna(0) > 0
        ][_cols_aud].copy()
        if not df_aud.empty:
            df_aud.to_excel(writer, sheet_name="AUDITORIA_NOVEDADES", index=False)
            ws_aud = writer.sheets["AUDITORIA_NOVEDADES"]
            ws_aud.set_column("A:N", 20)
            ws_aud.freeze_panes(1, 0)

        # ──────────────────────────────────────────────────────────────────────
        # ANÁLISIS OFF-TIME
        # ──────────────────────────────────────────────────────────────────────
        ws_oft = wb.add_worksheet("ANALISIS_OFF_TIME")
        ws_oft.hide_gridlines(2)
        ws_oft.set_column("A:A", 32)
        ws_oft.set_column("B:H", 18)
        ws_oft.merge_range("A1:H2",
                           f"ANÁLISIS OFF-TIME — Fuera de ventana {OFF_TIME_START}h–{OFF_TIME_END}h",
                           fmt_title)
        ws_oft.write_row("A4", ["SEGMENTO", "GUÍAS", "ON TIME", "LATE", "% OTIF",
                                 "% DEL TOTAL", "SLA PROM.", "ATRASO PROM."], fmt_header)
        row_oft = 4
        for seg_label, seg_mask in [
            (f"ON-TIME (creadas {OFF_TIME_START}h–{OFF_TIME_END}h)", _col_oft == "NO"),
            (f"OFF-TIME (<{OFF_TIME_START}h ó ≥{OFF_TIME_END}h)",    _col_oft == "SÍ"),
            ("TOTAL GENERAL",                                         pd.Series(True, index=df_ent.index)),
        ]:
            sub = df_ent[seg_mask]
            if sub.empty:
                continue
            n_s   = len(sub)
            ot_s  = int(sub["ES_ON_TIME"].sum())
            lt_s  = int(sub["ES_LATE"].sum())
            pct_s = ot_s / n_s if n_s > 0 else 0
            pct_t = n_s / n_total_ent if n_total_ent > 0 else 0
            sla_s = sub["SLA_DIAS"].mean()    if "SLA_DIAS"    in sub.columns else 0
            atr_s = sub[sub.get("DIAS_ATRASO", pd.Series(0, index=sub.index)) > 0]["DIAS_ATRASO"].mean() \
                    if "DIAS_ATRASO" in sub.columns else 0
            fmt_row = _f(border=1, bg_color="#F8CBAD") if "TOTAL" in seg_label else _f(border=1)
            ws_oft.write(row_oft, 0, seg_label,           fmt_row)
            ws_oft.write(row_oft, 1, n_s,                 _f(border=1, num_format="#,##0"))
            ws_oft.write(row_oft, 2, ot_s,                _f(border=1, num_format="#,##0", bg_color="#C6EFCE"))
            ws_oft.write(row_oft, 3, lt_s,                _f(border=1, num_format="#,##0", bg_color="#FFC7CE"))
            ws_oft.write(row_oft, 4, pct_s,               _f(border=1, num_format="0.0%", bold=True))
            ws_oft.write(row_oft, 5, pct_t,               _f(border=1, num_format="0.0%"))
            ws_oft.write(row_oft, 6, round(sla_s, 1),     _f(border=1, num_format="0.0"))
            ws_oft.write(row_oft, 7, round(atr_s if not pd.isna(atr_s) else 0, 1),
                         _f(border=1, num_format="0.0"))
            row_oft += 1

        # OFF-TIME por corredor
        row_oft += 2
        ws_oft.merge_range(row_oft, 0, row_oft, 5, "% OTIF POR CORREDOR — ON-TIME vs OFF-TIME", fmt_subtitle)
        row_oft += 1
        ws_oft.write_row(row_oft, 0, ["CORREDOR", "GUÍAS ON-TIME", "% OTIF ON-TIME",
                                       "GUÍAS OFF-TIME", "% OTIF OFF-TIME", "DIFERENCIA"], fmt_header)
        row_oft += 1
        corr_col = "CORREDOR_OTIF" if "CORREDOR_OTIF" in df_ent.columns else "TIPO_COBERTURA_2026"
        for corredor, grp in df_ent.groupby(corr_col):
            g_on  = grp[_col_oft.reindex(grp.index, fill_value="NO") == "NO"]
            g_off = grp[_col_oft.reindex(grp.index, fill_value="NO") == "SÍ"]
            pct_on  = g_on["ES_ON_TIME"].mean()  if len(g_on)  > 0 else None
            pct_off = g_off["ES_ON_TIME"].mean()  if len(g_off) > 0 else None
            diff    = (pct_on - pct_off) if pct_on is not None and pct_off is not None else None
            ws_oft.write(row_oft, 0, str(corredor), _f(border=1))
            ws_oft.write(row_oft, 1, len(g_on),     _f(border=1, num_format="#,##0"))
            ws_oft.write(row_oft, 2, pct_on  if pct_on  is not None else "N/A", _f(border=1, num_format="0.0%"))
            ws_oft.write(row_oft, 3, len(g_off),    _f(border=1, num_format="#,##0"))
            ws_oft.write(row_oft, 4, pct_off if pct_off is not None else "N/A", _f(border=1, num_format="0.0%"))
            ws_oft.write(row_oft, 5, diff    if diff    is not None else "N/A",
                         _f(border=1, num_format="0.0%",
                            bg_color="#FFC7CE" if diff is not None and diff > 0.05 else "#FFFFFF"))
            row_oft += 1

        # ──────────────────────────────────────────────────────────────────────
        # IN FULL ANÁLISIS
        # ──────────────────────────────────────────────────────────────────────
        ws_inf = wb.add_worksheet("IN_FULL_ANALISIS")
        ws_inf.hide_gridlines(2)
        ws_inf.set_column("A:A", 38)
        ws_inf.set_column("B:F", 18)
        ws_inf.merge_range("A1:F2", "ANÁLISIS IN FULL — COMPLETITUD DE ENTREGAS v7.0", fmt_title)
        ws_inf.write_row("A4", ["INDICADOR", "VALOR"], fmt_header)
        kpis_inf = [
            ("Total guías del período (universo completo)", n_total_periodo),
            ("Guías entregadas medidas (ON TIME + LATE)",   n_total_ent),
            ("Guías ON TIME",                               n_on_time),
            ("Guías IN FULL (entregadas)",                  in_full_total),
            ("Guías ON TIME + IN FULL STRICT",              in_full_strict),
            ("% In Full (de entregadas)",                   pct_in_full),
            ("% OTIF AJUSTADO (on_time / entregadas)",      pct_otif_ajustado),
            ("% OTIF ÁCIDO (on_time_in_full / total)",      pct_otif_acido),
        ]
        for ri, (lbl, val) in enumerate(kpis_inf, 4):
            ws_inf.write(ri, 0, lbl, _f(border=1))
            if isinstance(val, float):
                bg = "#C6EFCE" if val >= 0.8 else "#FFEB9C" if val >= 0.7 else "#FFC7CE"
                ws_inf.write(ri, 1, val, _f(border=1, num_format="0.0%", bold=True, bg_color=bg))
            else:
                ws_inf.write(ri, 1, val, _f(border=1, num_format="#,##0", bold=True))

        row_inf = 14
        ws_inf.merge_range(row_inf, 0, row_inf, 4, "IN FULL POR CORREDOR", fmt_subtitle)
        row_inf += 1
        ws_inf.write_row(row_inf, 0, ["CORREDOR", "GUÍAS", "IN FULL", "% IN FULL", "% OTIF AJ."], fmt_header)
        row_inf += 1
        for corredor, grp in df_ent.groupby(corr_col):
            n_c    = len(grp)
            inf_c  = int(grp["IN_FULL"].sum())  if "IN_FULL"  in grp.columns else 0
            ot_c   = int(grp["ES_ON_TIME"].sum())
            pct_if = inf_c / n_c if n_c > 0 else 0
            pct_ot = ot_c  / n_c if n_c > 0 else 0
            ws_inf.write(row_inf, 0, str(corredor), _f(border=1))
            ws_inf.write(row_inf, 1, n_c,  fmt_num)
            ws_inf.write(row_inf, 2, inf_c, fmt_num)
            ws_inf.write(row_inf, 3, pct_if,
                         _f(border=1, num_format="0.0%",
                            bg_color="#C6EFCE" if pct_if >= 0.8 else
                            "#FFEB9C" if pct_if >= 0.7 else "#FFC7CE"))
            ws_inf.write(row_inf, 4, pct_ot, fmt_pct)
            row_inf += 1

        row_inf += 2
        ws_inf.merge_range(row_inf, 0, row_inf, 4, "EVOLUCIÓN MENSUAL IN FULL", fmt_subtitle)
        row_inf += 1
        ws_inf.write_row(row_inf, 0, ["MES", "GUÍAS", "IN FULL", "% IN FULL", "% OTIF AJ."], fmt_header)
        row_inf += 1
        for mes_v, grp_m in df_ent.groupby("MES"):
            n_m   = len(grp_m)
            inf_m = int(grp_m["IN_FULL"].sum()) if "IN_FULL" in grp_m.columns else 0
            ot_m  = int(grp_m["ES_ON_TIME"].sum())
            pct_if_m = inf_m / n_m if n_m > 0 else 0
            pct_ot_m = ot_m  / n_m if n_m > 0 else 0
            ws_inf.write(row_inf, 0, str(mes_v), _f(border=1))
            ws_inf.write(row_inf, 1, n_m,        fmt_num)
            ws_inf.write(row_inf, 2, inf_m,       fmt_num)
            ws_inf.write(row_inf, 3, pct_if_m,
                         _f(border=1, num_format="0.0%",
                            bg_color="#C6EFCE" if pct_if_m >= 0.8 else
                            "#FFEB9C" if pct_if_m >= 0.7 else "#FFC7CE"))
            ws_inf.write(row_inf, 4, pct_ot_m, fmt_pct)
            row_inf += 1

        # ──────────────────────────────────────────────────────────────────────
        # BASE DETALLE — tabla completa para análisis propio en Excel
        # ──────────────────────────────────────────────────────────────────────
        _cols_base_want = [
            "GUIA", "MES", "FECHA_DE_CREACION", "HORA_CREACION", "ES_OFF_TIME",
            "PIEZA", "NOMOFI", "LINEA COMERCIAL",
            "REGIONAL_ORIGEN", "REG DESTINO CORRECTA",
            "CIUDAD_ORIGEN", "CIUDAD_DESTINO",
            "TIPO TRAYECTO", "CORREDOR_OTIF", "TIPO_COBERTURA_2026",
            "NOMEST", "FECREGEST",
            "DIAS_EN_ESTADO", "DIAS_EN_NOVEDAD", "FECHA_INI_NOVEDAD",
            "SLA_DIAS", "FECHA_INICIO_SLA", "FECHA_PROMETIDA", "PROMESA_EXTENDIDA",
            "CUMPLE_OTIF", "DIAS_ATRASO",
            "IN_FULL", "IN_FULL_STRICT",
            "TIEMPO DE ENTREGA", "TIEMPO_ENTREGA_NETO",
            "OTIF_AJUSTADO_PCT", "OTIF_ACIDO_PCT",
        ]
        cols_base = [c for c in _cols_base_want if c in df_ent.columns]
        df_base   = df_ent[cols_base].copy()

        # Formatear fechas antes de exportar
        for col in ["FECHA_DE_CREACION", "FECREGEST", "FECHA_INI_NOVEDAD",
                    "FECHA_INICIO_SLA", "FECHA_PROMETIDA", "PROMESA_EXTENDIDA"]:
            if col in df_base.columns and pd.api.types.is_datetime64_any_dtype(df_base[col]):
                df_base[col] = df_base[col].dt.strftime("%Y-%m-%d")

        df_base.to_excel(writer, sheet_name="BASE_DETALLE", index=False)
        ws_b = writer.sheets["BASE_DETALLE"]
        ws_b.freeze_panes(1, 0)
        ws_b.autofilter(0, 0, len(df_base), len(df_base.columns) - 1)
        try:
            ws_b.add_table(0, 0, len(df_base), len(df_base.columns) - 1, {
                "columns": [{"header": c} for c in df_base.columns],
                "style": "TableStyleLight9"
            })
        except Exception as _e_tbl:
            print(f"     ⚠️  Tabla BASE_DETALLE: {_e_tbl}")

    print(f"  ✅ REPORTE_OTIF_GERENCIAL generado: {ruta_out.name}")
    print(f"     Pestañas: GLOSARIO | DASHBOARD_GERENCIAL | TENDENCIA_MENSUAL")
    print(f"              DETALLE_REGIONALES | DETALLE_DIRECTO | DETALLE_REEXPEDIDO")
    print(f"              Matrices heatmap por mes | AUDITORIA_NOVEDADES")
    print(f"              ANALISIS_OFF_TIME | IN_FULL_ANALISIS | BASE_DETALLE")
    return ruta_out




if __name__ == "__main__":
    import traceback
    import time
    from pathlib import Path
    import sys

    carpeta_out = Path(CARPETA_SALIDA)
    carpeta_out.mkdir(parents=True, exist_ok=True)

    ts_log    = pd.Timestamp.now().strftime("%Y%m%d_%H%M")
    ruta_log  = carpeta_out / f"log_proceso_{ts_log}.txt"
    _LOG_FILE = open(ruta_log, "w", encoding="utf-8")

    import builtins
    _print_orig = builtins.print
    def _print_dual(*args, **kwargs):
        _print_orig(*args, **kwargs)
        if _LOG_FILE and not _LOG_FILE.closed:
            try:
                msg = " ".join(str(a) for a in args)
                _LOG_FILE.write(msg + "\n")
                _LOG_FILE.flush()
            except Exception: pass
    builtins.print = _print_dual

    try:
        print("=" * 65)
        print(f"  CONSOLIDADOR TRANSCARGA MUNDIAL  —  v7.0")
        print(f"  Inicio: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print("=" * 65)
        print("  CORRECCIONES v7.0 ACTIVAS:")
        print("  ✅ procesar_otif() corre ANTES de exportar() → OTIF en maestro")
        print("  ✅ IN_FULL_STRICT = on_time AND entregado (no inflado)")
        print("  ✅ OFF_TIME detecta <8h Y >=20h (no solo >=20h)")
        print("  ✅ REGIONAL_BUC completa en cond_med_buc")
        print("  ✅ OTIF ÁCIDO y AJUSTADO con denominadores correctos")
        print("  ✅ TIEMPO_ENTREGA_NETO calculado y exportado")
        print("  ✅ Errores en reportes secundarios no silenciados (traceback)")
        print("=" * 65)
        t_inicio = time.time()

        verificar_dependencias()
        carpeta_in = Path(CARPETA_ARCHIVOS)
        festivos, df_trayectos, df_regionales, df_clientes, dict_unif, malla_cubrimiento = _cargar_referencias_paralelo(
            ARCHIVO_FESTIVOS, ARCHIVO_TRAYECTOS, ARCHIVO_REGIONALES, ARCHIVO_CLIENTES, ARCHIVO_UNIFICACION
        )
        df, diagnosticos = consolidar(carpeta_in)
        df               = limpiar(df)
        df               = agregar_columnas_calculadas(df, festivos, df_trayectos, df_regionales, df_clientes, dict_unif)
        
        # ── PASO 1: OTIF primero → añade todas sus columnas al df ───────────────
        # CORRECCIÓN CRÍTICA: procesar_otif() debe correr ANTES de exportar()
        # para que el maestro Excel incluya CUMPLE_OTIF, IN_FULL, ES_OFF_TIME,
        # SLA_DIAS, FECHA_PROMETIDA, OTIF_AJUSTADO_PCT, OTIF_ACIDO_PCT, etc.
        sep("PROCESANDO OTIF (antes de exportar)")
        try:
            df = procesar_otif(df, carpeta_out, festivos, malla_cubrimiento)
            print("  ✅ Columnas OTIF agregadas al df — listas para exportar")
        except Exception as e_otif:
            print(f"\n  ⚠️  ERROR en procesar_otif: {e_otif}")
            traceback.print_exc()
            print("  ⚠️  Continuando sin columnas OTIF en el maestro")

        # ── PASO 2: Exportar maestro con TODAS las columnas (incluye OTIF) ────
        sep("EXPORTACIÓN — MAESTRO + ALERTAS")
        exportar(df, carpeta_out)

        # ── PASO 3: Reportes operativos (usan df ya completo) ─────────────────
        try:
            generar_alertas(df, festivos, carpeta_out)
        except Exception as e_al:
            print(f"  ⚠️  generar_alertas: {e_al}")
            traceback.print_exc()
        try:
            generar_reporte_comercial(df, carpeta_out)
        except Exception as e_com:
            print(f"  ⚠️  generar_reporte_comercial: {e_com}")
            traceback.print_exc()
        try:
            generar_reporte_franquicias(df, carpeta_out)
        except Exception as e_fra:
            print(f"  ⚠️  generar_reporte_franquicias: {e_fra}")
            traceback.print_exc()
        try:
            generar_alertas_franquicias_bog(df, festivos, carpeta_out)
        except Exception as e_bog:
            print(f"  ⚠️  generar_alertas_franquicias_bog: {e_bog}")
            traceback.print_exc()

        # ── PASO 4: Reporte OTIF gerencial ────────────────────────────────────
        try:
            generar_reporte_otif(df, carpeta_out)
        except Exception as e_rep:
            print(f"\n  ⚠️  ERROR en generar_reporte_otif: {e_rep}")
            traceback.print_exc()

        reporte_diagnostico(df, diagnosticos, carpeta_out)
        print(f"\n🏁 PROCESO COMPLETADO EN {time.time()-t_inicio:.1f} SEGUNDOS")

    except Exception as e:
        print(f"\n  ❌ ERROR FATAL: {e}")
        traceback.print_exc()
    finally:
        builtins.print = _print_orig
        if _LOG_FILE: _LOG_FILE.close()
    input("\nPresiona ENTER para cerrar...")

