#!/usr/bin/env python3
"""
convert.py - Convierte INFORME FIBRA FINAL.xlsx → data.js
Se ejecuta automáticamente via GitHub Actions al subir el Excel.
"""

import os
import json
import glob
import re
import pandas as pd
import numpy as np
from datetime import datetime

# ── Buscar el archivo Excel ──────────────────────────────────────────────────
xlsx_files = glob.glob("*.xlsx") + glob.glob("*.xls")
if not xlsx_files:
    print("❌ No se encontró ningún archivo .xlsx en el repositorio")
    exit(1)

EXCEL_PATH = xlsx_files[0]
print(f"📂 Procesando: {EXCEL_PATH}")

def safe_val(v):
    """Convierte valores de pandas a tipos seguros para JSON."""
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return None
    if isinstance(v, (np.integer,)):
        return int(v)
    if isinstance(v, (np.floating,)):
        return round(float(v), 4)
    if isinstance(v, pd.Timestamp):
        return v.strftime("%Y-%m-%d")
    if hasattr(v, 'item'):
        return v.item()
    return v

def df_to_records(df):
    """Convierte un DataFrame a lista de dicts con valores seguros."""
    df = df.dropna(how='all').reset_index(drop=True)
    records = []
    for _, row in df.iterrows():
        rec = {str(k): safe_val(v) for k, v in row.items()}
        records.append(rec)
    return records

def clean_col(col):
    """Limpia nombre de columna."""
    return str(col).strip().upper().replace(" ", "_").replace("/", "_")

# ── Leer todas las hojas ─────────────────────────────────────────────────────
xl = pd.ExcelFile(EXCEL_PATH)
sheet_names = xl.sheet_names
print(f"📋 Hojas encontradas: {sheet_names}")

output = {
    "meta": {
        "archivo": EXCEL_PATH,
        "hojas": sheet_names,
        "actualizado": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "timestamp": int(datetime.now().timestamp())
    },
    "hojas": {}
}

# Mapeo de nombres de hoja a claves del dashboard
HOJA_MAP = {
    # Ajusta estos nombres según tus hojas reales
    "RESUMEN": "resumen",
    "LIDERAZGO": "liderazgo",
    "SUPERVISORES": "liderazgo",
    "TEMPORAL": "temporal",
    "DIARIO": "temporal",
    "ASESORES": "asesores",
    "AGENTES": "asesores",
    "QUIEBRES": "quiebres",
    "CANCELACIONES": "quiebres",
    "LEGALIZACIONES": "legalizaciones",
    "LEGALIZ": "legalizaciones",
    "AGENDA": "agenda",
    "AGENDAS": "agenda",
    "GEO": "geografia",
    "GEOGRAFIA": "geografia",
    "DEPARTAMENTOS": "geografia",
    "MIX": "mix",
    "PRODUCTOS": "mix",
    "EMPAQUETADOS": "mix",
    "INSIGHTS": "insights",
    "PROYECCION": "insights",
}

for sheet in sheet_names:
    try:
        df = xl.parse(sheet, header=0)
        df.columns = [clean_col(c) for c in df.columns]
        records = df_to_records(df)
        
        # Determinar clave normalizada
        clave = None
        sheet_upper = sheet.upper().strip()
        for k, v in HOJA_MAP.items():
            if k in sheet_upper:
                clave = v
                break
        if not clave:
            clave = re.sub(r'[^A-Z0-9]', '_', sheet_upper).lower()
        
        # Si ya existe la clave, fusionar
        if clave in output["hojas"] and isinstance(output["hojas"][clave], list):
            output["hojas"][clave].extend(records)
        else:
            output["hojas"][clave] = records
        
        print(f"  ✅ {sheet} → [{clave}] {len(records)} filas")
    except Exception as e:
        print(f"  ⚠️  Error en hoja '{sheet}': {e}")

# ── Calcular métricas globales desde los datos ───────────────────────────────
def calc_metricas(hojas):
    metricas = {}
    
    # Intentar extraer métricas del resumen si existe
    resumen = hojas.get("resumen", [])
    if resumen:
        # Buscar columnas numéricas típicas
        for row in resumen:
            for k, v in row.items():
                if v is not None and isinstance(v, (int, float)):
                    key_lower = k.lower()
                    if "alta" in key_lower or "terminad" in key_lower:
                        metricas["total_altas"] = metricas.get("total_altas", 0) + v
                    if "anulad" in key_lower or "cancel" in key_lower:
                        metricas["total_anuladas"] = metricas.get("total_anuladas", 0) + v
                    if "pendient" in key_lower:
                        metricas["total_pendientes"] = metricas.get("total_pendientes", 0) + v
    
    # Calcular efectividad global
    altas = metricas.get("total_altas", 0)
    anuladas = metricas.get("total_anuladas", 0)
    if (altas + anuladas) > 0:
        metricas["efectividad_global"] = round(altas / (altas + anuladas) * 100, 1)
    
    return metricas

output["metricas"] = calc_metricas(output["hojas"])

# ── Guardar como data.js ─────────────────────────────────────────────────────
json_str = json.dumps(output, ensure_ascii=False, indent=2)
js_content = f"""// ⚡ AUTO-GENERADO por GitHub Actions — NO EDITAR MANUALMENTE
// Última actualización: {output['meta']['actualizado']}
// Archivo fuente: {EXCEL_PATH}

const INFORME_DATA = {json_str};

// Emitir evento para que el dashboard se actualice
if (typeof window !== 'undefined') {{
  window.dispatchEvent(new CustomEvent('datosActualizados', {{ detail: INFORME_DATA }}));
}}
"""

with open("data.js", "w", encoding="utf-8") as f:
    f.write(js_content)

print(f"\n✅ data.js generado exitosamente")
print(f"   Hojas procesadas: {list(output['hojas'].keys())}")
print(f"   Tamaño: {len(js_content):,} bytes")
