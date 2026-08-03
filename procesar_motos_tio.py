#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
procesar_motos_tio.py
=====================
Lee "control_motos_tio.xlsx" (hoja "Registro Diario" + hoja "Clientes")
y genera el dashboard HTML.

Uso:
    python procesar_motos_tio.py

Con fecha manual (debug):
    python procesar_motos_tio.py --hoy 2026-07-28
"""

import argparse
import csv
import io
import json
import os
import sys
import urllib.request
from collections import Counter, defaultdict
from datetime import date, datetime, timedelta
from pathlib import Path

import openpyxl

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

# ---------------------------------------------------------------------------
# CONFIGURACIÓN
# ---------------------------------------------------------------------------

NOMBRE_EXCEL   = "control_motos_tio.xlsx"
HOJA_REGISTRO  = "Registro Diario"
HOJA_CLIENTES  = "Clientes"
NOMBRE_NEGOCIO = "SEVERA MOTOS"    # ← cambia si quieres otro nombre en el dashboard
UMBRAL_ATRASO_LEVE = 2

# ---------------------------------------------------------------------------
# GOOGLE SHEETS — datos del cobrador en línea
# ---------------------------------------------------------------------------
# Cuando el cobrador llena el formulario (cobro_Code.gs), los pagos van a un
# Google Sheet.  Para que este script los lea automáticamente:
#   1. Abre el Google Sheet → Archivo → Compartir → Publicar en la web
#   2. Elige la hoja "Registro Diario" → formato CSV → Publicar
#   3. Pega el ID del Sheet en GOOGLE_SHEET_ID  (está en la URL del sheet)
#   4. Vuelve a correr: python procesar_motos_tio.py
# Si GOOGLE_SHEET_ID está vacío, el script solo usa el Excel local.
GOOGLE_SHEET_ID  = "1b_d0lDAYCnPfbcCKRxzPuWN4NOS75Yw0VorFXUBOSPU"
GOOGLE_SHEET_GID = "0"  # ← GID de la pestaña "Registro Diario" (normalmente 0)

MESES_ES = ["", "enero","febrero","marzo","abril","mayo","junio",
            "julio","agosto","septiembre","octubre","noviembre","diciembre"]
DIAS_ES       = ["lunes","martes","miércoles","jueves","viernes","sábado","domingo"]
DIAS_ES_ABREV = ["lun","mar","mié","jue","vie","sáb","dom"]

# ---------------------------------------------------------------------------
# UTILIDADES
# ---------------------------------------------------------------------------

def formatear_pesos(valor):
    valor = round(valor or 0)
    signo = "-" if valor < 0 else ""
    return f"{signo}${abs(valor):,.0f}".replace(",",".")

def fecha_legible(d):
    return f"{DIAS_ES[d.weekday()].capitalize()} {d.day} de {MESES_ES[d.month]} de {d.year}"

def resolver_ruta_excel():
    base = Path(__file__).resolve().parent
    ruta = base / NOMBRE_EXCEL
    if ruta.exists():
        return ruta
    print(f"No encontré '{NOMBRE_EXCEL}' junto al script ({base}).")
    print("Ejecuta 'python crear_datos_tio.py' primero para generarlo.")
    sys.exit(1)

# ---------------------------------------------------------------------------
# CARGA DE DATOS
# ---------------------------------------------------------------------------

def cargar_clientes(wb):
    info = {}
    if HOJA_CLIENTES not in wb.sheetnames:
        return info
    ws = wb[HOJA_CLIENTES]
    for row in ws.iter_rows(min_row=2, values_only=True):
        row = (list(row) + [None]*6)[:6]
        nombre, moto, placa, inicio, obs, total_cuotas_raw = row
        if not nombre:
            continue
        clave = str(nombre).strip().lower()
        info[clave] = {
            "nombre": str(nombre).strip().title(),
            "moto": moto,
            "placa": placa,
            "inicio": inicio.date() if isinstance(inicio, datetime) else None,
            "observaciones": obs,
            "total_cuotas": int(total_cuotas_raw) if isinstance(total_cuotas_raw, (int, float)) else 64,
        }
    return info

def cargar_registro(wb):
    ws = wb[HOJA_REGISTRO]
    filas = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        fecha   = row[0] if len(row) > 0 else None
        cliente = row[1] if len(row) > 1 else None
        if not isinstance(fecha, datetime) or not cliente:
            continue
        pago_diario_raw = row[7] if len(row) > 7 else None
        pago_diario = float(pago_diario_raw) if isinstance(pago_diario_raw, (int,float)) else 0.0
        pago_recibido = row[8] if len(row) > 8 else None
        medio_pago    = row[9] if len(row) > 9 else None
        saldo_raw     = row[10] if len(row) > 10 else None
        observaciones = row[15] if len(row) > 15 else None
        filas.append({
            "fecha":          fecha.date(),
            "cliente_clave":  str(cliente).strip().lower(),
            "cliente_raw":    str(cliente).strip(),
            "moto":           row[2] if len(row) > 2 else None,
            "placa":          row[3] if len(row) > 3 else None,
            "pago_diario":    pago_diario,
            "pago_recibido":  float(pago_recibido) if isinstance(pago_recibido,(int,float)) else None,
            "medio_pago":     medio_pago if isinstance(medio_pago, str) else None,
            "saldo_cache":    float(saldo_raw) if isinstance(saldo_raw,(int,float)) else None,
            "observaciones":  str(observaciones).strip() if observaciones else None,
        })
    return filas

# ---------------------------------------------------------------------------
# GOOGLE SHEETS — lectura de pagos en línea
# ---------------------------------------------------------------------------

def cargar_registro_sheets(sheet_id, gid="0"):
    """Descarga el Registro Diario desde Google Sheets (hoja publicada como CSV)."""
    url = (
        f"https://docs.google.com/spreadsheets/d/{sheet_id}"
        f"/export?format=csv&gid={gid}"
    )
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=10) as resp:
            text = resp.read().decode("utf-8-sig")
    except Exception as e:
        print(f"⚠️  No se pudo leer Google Sheets: {e}")
        return []

    filas = []
    reader = csv.reader(io.StringIO(text))
    for i, row in enumerate(reader):
        if i == 0:          # saltar encabezado
            continue
        if len(row) < 10:
            continue
        # Mapeo de columnas: igual que el Excel (A=0 … P=15)
        fecha_raw  = row[0].strip()
        cliente    = row[1].strip()
        if not fecha_raw or not cliente:
            continue
        # Parsear fecha (Google Sheets exporta como M/D/YYYY o YYYY-MM-DD)
        fecha = None
        for fmt in ("%m/%d/%Y", "%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y %H:%M:%S"):
            try:
                fecha = datetime.strptime(fecha_raw.split()[0], fmt.split()[0]).date()
                break
            except ValueError:
                continue
        if fecha is None:
            continue

        def _num(v):
            try:
                return float(str(v).replace(",", "").strip())
            except (ValueError, TypeError):
                return None

        pago_diario   = _num(row[7]) if len(row) > 7 else None
        pago_recibido = _num(row[8]) if len(row) > 8 else None
        medio_pago    = row[9].strip()  if len(row) > 9 else None
        observaciones = row[15].strip() if len(row) > 15 else None

        filas.append({
            "fecha":          fecha,
            "cliente_clave":  cliente.strip().lower(),
            "cliente_raw":    cliente,
            "moto":           row[2].strip() if len(row) > 2 else None,
            "placa":          row[3].strip() if len(row) > 3 else None,
            "pago_diario":    pago_diario or 0.0,
            "pago_recibido":  pago_recibido,
            "medio_pago":     medio_pago if medio_pago else None,
            "saldo_cache":    None,
            "observaciones":  observaciones if observaciones else None,
        })

    print(f"📡 Google Sheets: {len(filas)} pagos cargados online.")
    return filas

# ---------------------------------------------------------------------------
# PROCESAMIENTO  (idéntico a la versión de Jeffer)
# ---------------------------------------------------------------------------

def saldo_en_o_antes(historial_saldo, fecha_corte):
    vigente = 0.0
    for f, s in historial_saldo:
        if f > fecha_corte:
            break
        vigente = s
    return vigente

def calcular_estado(saldo, tarifa_ref):
    if saldo <= 0:
        return "al_dia"
    tarifas = saldo / tarifa_ref if tarifa_ref else 0
    return "atraso_leve" if tarifas <= UMBRAL_ATRASO_LEVE else "atraso_alto"

def frase_resumen_semana(nombre, dias_atrasados, monto):
    if not dias_atrasados:
        return f"Esta semana {nombre} está al día."
    nombres_dias = [DIAS_ES[d.weekday()] for d in sorted(dias_atrasados)]
    texto_dias = (", ".join(nombres_dias[:-1]) + " y " + nombres_dias[-1]) if len(nombres_dias) > 1 else nombres_dias[0]
    plural = "día" if len(nombres_dias) == 1 else "días"
    return f"Esta semana {nombre} se quedó {len(nombres_dias)} {plural} ({texto_dias}) por {formatear_pesos(monto)}."

def procesar(filas, clientes_meta, hoy):
    por_cliente = defaultdict(list)
    for f in filas:
        por_cliente[f["cliente_clave"]].append(f)

    resultado_clientes = []
    tendencia = defaultdict(lambda: {"esperado": 0.0, "recibido": 0.0})
    tabla_detalle = []
    total_global = defaultdict(float)
    ultimo_dia_mes_anterior = hoy.replace(day=1) - timedelta(days=1)

    for clave, registros in por_cliente.items():
        registros.sort(key=lambda r: r["fecha"])
        meta = clientes_meta.get(clave, {})
        nombre = meta.get("nombre") or registros[0]["cliente_raw"].title()
        moto   = meta.get("moto")   or registros[0]["moto"]
        placa  = meta.get("placa")  or registros[0]["placa"]
        inicio = meta.get("inicio") or registros[0]["fecha"]

        tarifas_validas = [r["pago_diario"] for r in registros if r["pago_diario"] > 0]
        tarifa_ref = Counter(tarifas_validas).most_common(1)[0][0] if tarifas_validas else 240_000.0

        saldo_efectivo = 0.0
        historial_saldo = []
        historial = {}
        recibido_total = 0.0
        recibido_mes   = 0.0
        ultimo_pago_fecha = None

        for r in registros:
            if r["fecha"] > hoy:
                continue
            recibido = r["pago_recibido"] or 0.0
            if r["saldo_cache"] is not None:
                saldo_efectivo = r["saldo_cache"]
            else:
                saldo_efectivo += r["pago_diario"] - recibido
            historial_saldo.append((r["fecha"], saldo_efectivo))
            historial[r["fecha"]] = r
            recibido_total += recibido
            if r["fecha"].year == hoy.year and r["fecha"].month == hoy.month:
                recibido_mes += recibido
            if recibido > 0:
                ultimo_pago_fecha = r["fecha"]
            tabla_detalle.append({
                "fecha":      r["fecha"].isoformat(),
                "cliente":    nombre,
                "placa":      placa,
                "esperado":   r["pago_diario"],
                "recibido":   recibido,
                "diferencia": r["pago_diario"] - recibido,
                "saldo":      saldo_efectivo,
                "medio_pago": r["medio_pago"],
                "nota":       r["observaciones"],
            })

        saldo_atrasado = historial_saldo[-1][1] if historial_saldo else 0.0
        fecha_fin_contrato = max((r["fecha"] for r in registros), default=hoy)
        restante_programado = sum(max(0.0, r["pago_diario"]) for r in registros if r["fecha"] > hoy)
        deuda_total_con_programado = saldo_atrasado + restante_programado
        atrasado_mes = saldo_atrasado - saldo_en_o_antes(historial_saldo, ultimo_dia_mes_anterior)

        esperado_total = recibido_total + saldo_atrasado
        esperado_mes   = recibido_mes + atrasado_mes

        # --- métricas semanales / contrato ---
        total_cuotas    = meta.get("total_cuotas", 64)
        pagos_historicos = sorted([r for r in registros if r["fecha"] <= hoy], key=lambda r: r["fecha"])
        pagos_realizados = len(pagos_historicos)
        pagos_restantes  = max(0, total_cuotas - pagos_realizados)
        porc_completado  = round(pagos_realizados / total_cuotas * 100, 1) if total_cuotas > 0 else 0

        # fecha fin de contrato = primer pago + total_cuotas semanas
        fecha_primer_pago = pagos_historicos[0]["fecha"] if pagos_historicos else inicio
        fecha_fin_prevista = fecha_primer_pago + timedelta(weeks=total_cuotas)

        # contrato completado?
        contrato_completo = pagos_realizados >= total_cuotas

        # días sin pagar: hoy vs (último pago + 7 días)
        # no aplica si el contrato ya está finalizado
        dias_sin_pagar = 0
        proximo_pago_esperado = None
        if not contrato_completo and ultimo_pago_fecha:
            proximo_pago_esperado = ultimo_pago_fecha + timedelta(days=7)
            if proximo_pago_esperado < hoy:
                dias_sin_pagar = (hoy - proximo_pago_esperado).days

        # últimos 8 pagos semanales para mostrar en la tarjeta
        ultimos_8 = pagos_historicos[-8:]
        base_num  = pagos_realizados - len(ultimos_8) + 1
        ultimos_pagos_data = []
        for i_p, r in enumerate(ultimos_8):
            rec = r["pago_recibido"] or 0
            esp = r["pago_diario"]
            if rec >= esp and esp > 0:
                est = "pagado"
            elif rec > 0:
                est = "parcial"
            else:
                est = "no_pago"
            ultimos_pagos_data.append({
                "fecha":     r["fecha"].isoformat(),
                "mes_dia":   f"{r['fecha'].day}/{r['fecha'].month}",
                "dia_nombre": DIAS_ES_ABREV[r["fecha"].weekday()],
                "recibido":  rec,
                "esperado":  esp,
                "estado":    est,
                "num_cuota": base_num + i_p,
            })

        meses_con_actividad = sorted({(r["fecha"].year, r["fecha"].month) for r in registros if r["fecha"] <= hoy})
        for anio_m, mes_m in meses_con_actividad:
            inicio_mes = date(anio_m, mes_m, 1)
            fin_mes = date(anio_m, mes_m+1, 1) - timedelta(days=1) if mes_m < 12 else date(anio_m+1, 1, 1) - timedelta(days=1)
            corte_fin = min(fin_mes, hoy)
            recibido_del_mes = sum(
                (r["pago_recibido"] or 0.0)
                for r in registros
                if r["fecha"] <= hoy and r["fecha"].year == anio_m and r["fecha"].month == mes_m
            )
            saldo_fin_mes = saldo_en_o_antes(historial_saldo, corte_fin)
            saldo_fin_mes_anterior = saldo_en_o_antes(historial_saldo, inicio_mes - timedelta(days=1))
            clave_mes = (anio_m, mes_m)
            tendencia[clave_mes]["recibido"] += recibido_del_mes
            tendencia[clave_mes]["esperado"] += recibido_del_mes + (saldo_fin_mes - saldo_fin_mes_anterior)

        semana = []
        dias_atrasados_semana = []
        monto_atrasado_semana = 0.0
        for i in range(6, -1, -1):
            d = hoy - timedelta(days=i)
            if d < inicio:
                continue
            r = historial.get(d)
            if r is None:
                estado_dia = "sin_dato"
                esperado_d = recibido_d = nota_d = None
            else:
                esperado_d = r["pago_diario"]
                recibido_d = r["pago_recibido"]
                nota_d     = r["observaciones"]
                if esperado_d == 0 and recibido_d:
                    estado_dia = "pagado"
                elif esperado_d == 0:
                    estado_dia = "sin_cobro"
                elif d == hoy and recibido_d is None:
                    estado_dia = "pendiente"
                elif recibido_d is None or recibido_d == 0:
                    estado_dia = "no_pago"
                elif recibido_d < esperado_d:
                    estado_dia = "parcial"
                else:
                    estado_dia = "pagado"
                if estado_dia in ("no_pago", "parcial"):
                    dias_atrasados_semana.append(d)
                    monto_atrasado_semana += esperado_d - (recibido_d or 0)
            semana.append({
                "fecha": d.isoformat(), "dia_nombre": DIAS_ES_ABREV[d.weekday()],
                "dia_numero": d.day, "esperado": esperado_d, "recibido": recibido_d,
                "estado": estado_dia, "nota": nota_d, "es_hoy": d == hoy,
            })

        cumplimiento  = (recibido_total / esperado_total * 100) if esperado_total > 0 else 100.0
        if contrato_completo:
            estado_general = "finalizado"
        elif dias_sin_pagar > 14:
            estado_general = "atraso_alto"
        elif dias_sin_pagar > 0:
            estado_general = "atraso_leve"
        else:
            estado_general = calcular_estado(saldo_atrasado, tarifa_ref)

        resultado_clientes.append({
            "clave": clave, "nombre": nombre, "moto": moto, "placa": placa,
            "inicio": inicio.isoformat(), "inicio_legible": fecha_legible(inicio),
            "tarifa_ref": tarifa_ref,
            "saldo_atrasado": saldo_atrasado, "saldo_en_tarifas": round(saldo_atrasado/tarifa_ref,1) if tarifa_ref else 0,
            "atrasado_mes": atrasado_mes, "recibido_total": recibido_total,
            "recibido_mes": recibido_mes, "esperado_total": esperado_total,
            "esperado_mes": esperado_mes, "restante_programado": restante_programado,
            "fecha_fin_contrato": fecha_fin_contrato.isoformat(),
            "fecha_fin_contrato_legible": fecha_legible(fecha_fin_contrato),
            "deuda_total_con_programado": deuda_total_con_programado,
            "cumplimiento": round(cumplimiento, 1),
            "ultimo_pago_fecha": ultimo_pago_fecha.isoformat() if ultimo_pago_fecha else None,
            "estado_general": estado_general, "semana": semana,
            "resumen_semana": frase_resumen_semana(nombre, dias_atrasados_semana, monto_atrasado_semana),
            "dias_atrasados_semana": len(dias_atrasados_semana),
            "monto_atrasado_semana": monto_atrasado_semana,
            # campos semanales / contrato
            "total_cuotas":    total_cuotas,
            "pagos_realizados": pagos_realizados,
            "pagos_restantes":  pagos_restantes,
            "porc_completado":  porc_completado,
            "fecha_fin_prevista": fecha_fin_prevista.isoformat(),
            "fecha_fin_prevista_legible": fecha_legible(fecha_fin_prevista),
            "contrato_completo": contrato_completo,
            "dias_sin_pagar":   dias_sin_pagar,
            "proximo_pago_esperado": proximo_pago_esperado.isoformat() if proximo_pago_esperado else None,
            "proximo_pago_esperado_legible": fecha_legible(proximo_pago_esperado) if proximo_pago_esperado else "—",
            "ultimos_pagos":    ultimos_pagos_data,
        })

        total_global["recibido_historico"]  += recibido_total
        total_global["recibido_mes"]        += recibido_mes
        total_global["esperado_historico"]  += esperado_total
        total_global["esperado_mes"]        += esperado_mes
        total_global["atrasado_historico"]  += saldo_atrasado
        total_global["atrasado_mes"]        += atrasado_mes
        total_global["restante_programado_total"] += restante_programado

    # contratos activos primero (por saldo), finalizados al final
    resultado_clientes.sort(key=lambda c: (1 if c["contrato_completo"] else 0, -c["saldo_atrasado"]))

    dias_en_mes_actual = (date(hoy.year, hoy.month+1, 1) - timedelta(days=1)).day if hoy.month < 12 else 31
    total_global["cantidad_motos"] = len({c["moto"] for c in resultado_clientes if c["moto"]}) or len(resultado_clientes)
    total_global["tarifa_diaria_combinada"] = sum(c["tarifa_ref"] for c in resultado_clientes)
    total_global["esperado_mes_completo"]   = total_global["tarifa_diaria_combinada"] * dias_en_mes_actual
    total_global["deuda_total_con_programado"] = total_global["atrasado_historico"] + total_global["restante_programado_total"]
    total_global["cumplimiento"] = round(
        (total_global["recibido_historico"] / total_global["esperado_historico"] * 100)
        if total_global["esperado_historico"] > 0 else 100.0, 1)

    tendencia_lista = [
        {"anio": a, "mes": m, "mes_nombre": MESES_ES[m][:3], "esperado": v["esperado"], "recibido": v["recibido"]}
        for (a, m), v in sorted(tendencia.items())
    ]
    tabla_detalle.sort(key=lambda x: x["fecha"], reverse=True)

    return {"clientes": resultado_clientes, "global": dict(total_global),
            "tendencia": tendencia_lista, "detalle": tabla_detalle}

# ---------------------------------------------------------------------------
# SALIDA
# ---------------------------------------------------------------------------

def imprimir_resumen_terminal(datos, hoy):
    banderas = {"al_dia": "[OK]", "atraso_leve": "[!] ", "atraso_alto": "[!!]", "finalizado": "[✓] "}
    print()
    print("=" * 64)
    print(f" {NOMBRE_NEGOCIO} · resumen al {hoy.strftime('%d/%m/%Y')}")
    print("=" * 64)
    for c in datos["clientes"]:
        b = banderas[c["estado_general"]]
        print(f" {b} {c['nombre']:<22} atrasado: {formatear_pesos(c['saldo_atrasado']):>13}")
        if c["dias_atrasados_semana"]:
            print(f"        -> {c['resumen_semana']}")
    print("-" * 64)
    g = datos["global"]
    print(f" Recibido histórico : {formatear_pesos(g['recibido_historico'])}")
    print(f" Atrasado total     : {formatear_pesos(g['atrasado_historico'])}")
    print(f" Cumplimiento       : {g['cumplimiento']}%")
    print("=" * 64)

def generar_html(datos, ruta_plantilla, ruta_salida, hoy):
    plantilla = ruta_plantilla.read_text(encoding="utf-8")
    payload = {
        "generado_en": datetime.now().strftime("%d/%m/%Y %I:%M %p"),
        "hoy": hoy.isoformat(), "hoy_legible": fecha_legible(hoy),
        **datos,
    }
    html = plantilla.replace("__DATOS_JSON_AQUI__", json.dumps(payload, ensure_ascii=False))
    html = html.replace("JEFFER MOTOS", NOMBRE_NEGOCIO)
    ruta_salida.write_text(html, encoding="utf-8")

# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--hoy", help="AAAA-MM-DD (debug)")
    parser.add_argument("--salida", help="Ruta del HTML de salida")
    args = parser.parse_args()

    ruta_excel = resolver_ruta_excel()
    hoy = datetime.strptime(args.hoy, "%Y-%m-%d").date() if args.hoy else date.today()

    print(f"Leyendo: {ruta_excel}")
    wb = openpyxl.load_workbook(ruta_excel, data_only=True)
    clientes_meta = cargar_clientes(wb)
    filas = cargar_registro(wb)
    print(f"{len(filas)} filas del Excel ({len(clientes_meta)} clientes)")

    # Combinar con Google Sheets si está configurado
    if GOOGLE_SHEET_ID:
        filas_online = cargar_registro_sheets(GOOGLE_SHEET_ID, GOOGLE_SHEET_GID)
        if filas_online:
            # Deduplicar por (fecha, cliente_clave) — el Excel tiene precedencia
            claves_excel = {(f["fecha"], f["cliente_clave"]) for f in filas}
            nuevos = [f for f in filas_online
                      if (f["fecha"], f["cliente_clave"]) not in claves_excel]
            filas.extend(nuevos)
            print(f"   + {len(nuevos)} pagos nuevos desde Google Sheets")
    print(f"Total: {len(filas)} registros combinados")

    datos = procesar(filas, clientes_meta, hoy)

    base_dir = Path(__file__).resolve().parent
    # Usa la plantilla tío (semanal); si no existe, cae al original
    ruta_plantilla = base_dir / "plantilla_dashboard_tio.html"
    if not ruta_plantilla.exists():
        ruta_plantilla = base_dir.parent / "control-motos" / "plantilla_dashboard.html"
    if not ruta_plantilla.exists():
        print("No encontré plantilla_dashboard.html.")
        sys.exit(1)

    ruta_salida = Path(args.salida) if args.salida else base_dir / "dashboard_motos_tio.html"
    generar_html(datos, ruta_plantilla, ruta_salida, hoy)
    imprimir_resumen_terminal(datos, hoy)
    print(f"\nDashboard generado en: {ruta_salida}")

if __name__ == "__main__":
    main()
