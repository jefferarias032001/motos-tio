#!/usr/bin/env python3
"""
generar_facturas.py
-------------------
Genera todos los recibos PDF de todos los clientes automáticamente.

Estructura de salida:
  facturas/
    Darwin Yaimis Ysabel/
      Recibo_Darwin-Yaimis-Ysabel_2026-06-30_Cuota46.pdf
      Recibo_Darwin-Yaimis-Ysabel_2026-07-06_Cuota47.pdf
    Gustavo Primo/
      Recibo_Gustavo-Primo_2026-01-15_Cuota1.pdf
    ...

Requisitos (solo la primera vez):
  pip install playwright
  playwright install chromium
"""

import sys, os, re, json, tempfile
from pathlib import Path
from datetime import date, datetime
from collections import defaultdict

# ── Re-usar procesamiento existente ─────────────────────────────────────────
BASE = Path(__file__).resolve().parent
sys.path.insert(0, str(BASE))
import openpyxl
import procesar_motos_tio as pmt

SALIDA = BASE / "facturas"

MESES = ["","Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio",
         "Agosto","Septiembre","Octubre","Noviembre","Diciembre"]

# ── Helpers ──────────────────────────────────────────────────────────────────
def _fmt_cop(n):
    return "${:,.0f}".format(int(n)).replace(",", ".")

def _fecha_larga(iso):
    if not iso: return "Fecha no disponible"
    try:
        y, m, d = str(iso)[:10].split("-")
        return f"{int(d)} de {MESES[int(m)]} de {y}"
    except Exception:
        return str(iso)

def _monto_letras(n):
    n = int(n)
    UNI = ["","un","dos","tres","cuatro","cinco","seis","siete","ocho","nueve",
           "diez","once","doce","trece","catorce","quince","dieciséis","diecisiete",
           "dieciocho","diecinueve"]
    DEC = ["","diez","veinte","treinta","cuarenta","cincuenta","sesenta","setenta","ochenta","noventa"]
    CENT= ["","cien","doscientos","trescientos","cuatrocientos","quinientos",
           "seiscientos","setecientos","ochocientos","novecientos"]

    def s100(x):
        if x < 20: return UNI[x]
        d, u = divmod(x, 10)
        return DEC[d] if u == 0 else DEC[d] + " y " + UNI[u]

    def s1000(x):
        if x == 0: return ""
        c, r = divmod(x, 100)
        a = ("ciento" if c == 1 else CENT[c]) if c else ""
        b = s100(r)
        return (a + " " + b).strip() if a and b else a or b

    def s_miles(x):
        # x puede ser > 99 (ej: 260 miles en 260_000)
        c, r = divmod(x, 100)
        a = ("ciento" if c == 1 else CENT[c]) if c else ""
        b = s100(r)
        return (a + " " + b).strip() if a and b else a or b

    miles, resto = divmod(n, 1000)
    if miles == 0:
        base = s1000(n)
    elif miles == 1:
        base = "mil" + (" " + s1000(resto) if resto else "")
    else:
        base = s_miles(miles) + " mil" + (" " + s1000(resto) if resto else "")
    return base.upper() + " PESOS M/CTE"

def _nombre_archivo(nombre):
    """Darwin Yaimis Ysabel → Darwin-Yaimis-Ysabel"""
    return re.sub(r'\s+', '-', nombre.strip())

def _nombre_dir(nombre):
    """Limpia caracteres inválidos para nombre de carpeta."""
    return re.sub(r'[<>:"/\\|?*]', '_', nombre).strip()

# ── Agrupar pagos en cuotas ──────────────────────────────────────────────────
def _agrupar_cuotas(registros, tarifa):
    """
    registros: lista de filas ordenadas por fecha.
    Acumula pago_recibido hasta completar una cuota (= tarifa).
    Cuotas del cobro_app ya vienen agrupadas: se expanden si > 1 tarifa.
    """
    cuotas = []
    acum   = 0.0
    for r in sorted(registros, key=lambda x: str(x.get("fecha", ""))):
        rec = float(r.get("pago_recibido") or 0)
        if rec <= 0:
            continue
        fecha = str(r.get("fecha", ""))[:10]

        if r.get("observaciones") == "cobro_app":
            # Puede cubrir varias cuotas si rec > tarifa
            n = max(1, round(rec / tarifa)) if tarifa > 0 else 1
            for _ in range(n):
                cuotas.append({"num": len(cuotas) + 1, "fecha": fecha, "total": int(tarifa)})
            continue

        acum += rec
        while acum >= tarifa - 0.5:
            cuotas.append({"num": len(cuotas) + 1, "fecha": fecha, "total": int(tarifa)})
            acum -= tarifa

    return cuotas

# ── HTML de una factura ──────────────────────────────────────────────────────
LOGO_SVG = (
    '<svg viewBox="0 0 56 56" fill="none" xmlns="http://www.w3.org/2000/svg" '
    'style="width:56px;height:56px">'
    '<circle cx="14" cy="38" r="9" stroke="#1A1714" stroke-width="3" fill="none"/>'
    '<circle cx="42" cy="38" r="9" stroke="#1A1714" stroke-width="3" fill="none"/>'
    '<path d="M14 38 L22 22 L36 22 L42 38" stroke="#1A1714" stroke-width="2.5" fill="none" stroke-linejoin="round"/>'
    '<path d="M22 22 L26 14 L34 14 L36 22" stroke="#1A1714" stroke-width="2" fill="none" stroke-linejoin="round"/>'
    '<path d="M28 22 L28 32" stroke="#1A1714" stroke-width="2"/>'
    '<circle cx="14" cy="38" r="3.5" fill="#F4C34B"/>'
    '<circle cx="42" cy="38" r="3.5" fill="#F4C34B"/>'
    '</svg>'
)
def _cargar_firma():
    import base64
    firma_path = BASE / "Firma" / "Severa motos - firma.png"
    if firma_path.exists():
        b64 = base64.b64encode(firma_path.read_bytes()).decode()
        return f'<img src="data:image/png;base64,{b64}" style="display:block;margin:0 auto 6px;max-height:60px;max-width:180px;object-fit:contain;">'
    # fallback si no existe el archivo
    return (
        '<svg width="160" height="50" viewBox="0 0 160 50" style="display:block;margin:0 auto 6px">'
        '<path d="M20 40 Q40 10 60 35 Q80 55 100 25 Q120 5 140 30" '
        'stroke="#1A1714" stroke-width="2" fill="none" stroke-linecap="round"/></svg>'
    )

FIRMA_SVG = _cargar_firma()
CSS = """
    @page{margin:12mm 14mm;size:A4 portrait;}
    *{-webkit-print-color-adjust:exact!important;print-color-adjust:exact!important;box-sizing:border-box;}
    body{margin:0;padding:40px;background:#fff;font-family:'Inter',Arial,sans-serif;color:#111;}
    .fw{max-width:520px;margin:0 auto;}
    .fh{display:flex;align-items:center;justify-content:space-between;border-bottom:3px solid #1A1714;padding-bottom:20px;margin-bottom:24px;}
    .fl{display:flex;align-items:center;gap:14px;}
    .ft{font-size:1.6rem;font-weight:900;letter-spacing:.04em;line-height:1.1;text-transform:uppercase;}
    .fsub{font-size:.65rem;letter-spacing:.18em;text-transform:uppercase;color:#555;margin-top:2px;}
    .fn-l{font-size:.65rem;letter-spacing:.15em;text-transform:uppercase;color:#888;text-align:right;}
    .fn-v{font-size:2rem;font-weight:900;font-family:'Courier New',monospace;color:#1A1714;line-height:1;}
    .tit{text-align:center;font-size:.7rem;letter-spacing:.25em;text-transform:uppercase;color:#666;margin-bottom:28px;font-weight:600;}
    .datos{display:grid;grid-template-columns:1fr 1fr;gap:14px 24px;margin-bottom:28px;}
    .campo label{font-size:.6rem;letter-spacing:.14em;text-transform:uppercase;color:#888;display:block;margin-bottom:3px;}
    .campo span{font-size:.95rem;font-weight:700;color:#111;display:block;}
    .mbox{background:#1A1714!important;color:#F4EFE2!important;border-radius:10px;padding:20px 24px;margin-bottom:28px;display:flex;align-items:center;justify-content:space-between;}
    .ml{font-size:.65rem;letter-spacing:.2em;text-transform:uppercase;color:#C8BEA0;}
    .mv{font-size:2rem;font-weight:900;font-family:'Courier New',monospace;color:#F4C34B!important;}
    .mlet{font-size:.75rem;color:#C8BEA0;margin-top:3px;}
    .sello{display:flex;align-items:center;gap:10px;background:#f0fdf4!important;border:2px solid #16a34a;border-radius:8px;padding:10px 16px;margin-bottom:28px;}
    .sello-txt{font-weight:800;color:#16a34a;font-size:.95rem;letter-spacing:.06em;text-transform:uppercase;}
    .sello-sub{font-size:.7rem;color:#166534;margin-top:1px;}
    .firma-area{display:flex;justify-content:flex-end;}
    .firma-bl{text-align:center;min-width:220px;}
    .firma-lin{border-top:1.5px solid #444;margin-bottom:5px;}
    .firma-nom{font-size:.7rem;letter-spacing:.1em;text-transform:uppercase;color:#333;font-weight:700;}
    .firma-cgo{font-size:.62rem;color:#888;letter-spacing:.08em;text-transform:uppercase;}
    .footer{margin-top:24px;border-top:1px solid #e5e5e5;padding-top:14px;text-align:center;font-size:.62rem;color:#aaa;letter-spacing:.06em;}
"""

def _html(nombre_display, moto, placa, cuota_num, fecha, total):
    ns = str(cuota_num).zfill(4)
    return f"""<!DOCTYPE html>
<html lang="es"><head><meta charset="utf-8">
<title>Recibo {ns} · {nombre_display}</title>
<style>{CSS}</style>
</head><body><div class="fw">
  <div class="fh">
    <div class="fl">{LOGO_SVG}
      <div><div class="ft">Severa<br>Motos</div><div class="fsub">Financiamiento de motos</div></div>
    </div>
    <div><div class="fn-l">Recibo N°</div><div class="fn-v">{ns}</div></div>
  </div>
  <div class="tit">Recibo de pago · Cuota de financiamiento</div>
  <div class="datos">
    <div class="campo"><label>Cliente</label><span>{nombre_display}</span></div>
    <div class="campo"><label>Fecha de pago</label><span>{_fecha_larga(fecha)}</span></div>
    <div class="campo"><label>Motocicleta</label><span>{moto or '—'}</span></div>
    <div class="campo"><label>Placa</label><span>{placa or '—'}</span></div>
  </div>
  <div class="mbox">
    <div>
      <div class="ml">Valor recibido</div>
      <div class="mv">{_fmt_cop(total)}</div>
      <div class="mlet">{_monto_letras(total)}</div>
    </div>
    <div style="font-size:2.5rem;opacity:.3">🏍️</div>
  </div>
  <div class="sello">
    <span style="font-size:1.4rem">✅</span>
    <div>
      <div class="sello-txt">Pago recibido</div>
      <div class="sello-sub">Cuota N° {cuota_num} cancelada · {_fecha_larga(fecha)}</div>
    </div>
  </div>
  <div class="firma-area"><div class="firma-bl">
    {FIRMA_SVG}
    <div class="firma-lin"></div>
    <div class="firma-nom">Severa Motos</div>
    <div class="firma-cgo">Administración · Firma autorizada</div>
  </div></div>
  <div class="footer">Severa Motos · Constancia de pago · Cuota {cuota_num} de financiamiento</div>
</div></body></html>"""

# ── Generación principal ─────────────────────────────────────────────────────
def generar_facturas(filas_todas, datos):
    """
    filas_todas : lista completa de registros (Excel + Sheets mezclados).
    datos       : resultado de procesar() con clientes, global, detalle.
    """
    try:
        from playwright.sync_api import sync_playwright
        tiene_playwright = True
    except ImportError:
        tiene_playwright = False
        print("\n⚠  Playwright no está instalado. Generando archivos HTML en su lugar.")
        print("   Para generar PDF instala con:")
        print("     pip install playwright")
        print("     playwright install chromium\n")

    # Índice de registros por clave de cliente
    por_cliente = defaultdict(list)
    for r in filas_todas:
        por_cliente[r["cliente_clave"]].append(r)

    # Info de clientes procesados
    info_cliente = {pmt.normalizar_nombre(c["nombre"].lower()): c
                    for c in datos.get("clientes", [])}

    total_ok  = 0
    total_skip = 0
    SALIDA.mkdir(exist_ok=True)

    def _generar(browser):
        nonlocal total_ok, total_skip
        for clave, registros in sorted(por_cliente.items()):
            info  = info_cliente.get(clave, {})
            nombre = info.get("nombre") or registros[0].get("cliente_raw", clave).title()
            moto   = info.get("moto") or registros[0].get("moto") or ""
            placa  = info.get("placa") or registros[0].get("placa") or ""
            tarifa = float(info.get("tarifa_ref") or 0)

            if tarifa <= 0:
                cfg = pmt.get_config(clave)
                tarifa = float(cfg["meta"]) if cfg and cfg.get("meta") else 240_000

            nombre_display = pmt.NOMBRE_DISPLAY.get(nombre.lower().strip(), nombre)
            cuotas = _agrupar_cuotas(registros, tarifa)
            if not cuotas:
                total_skip += 1
                continue

            carpeta = SALIDA / _nombre_dir(nombre)
            carpeta.mkdir(parents=True, exist_ok=True)
            ext = "pdf" if tiene_playwright else "html"
            nombre_arc = _nombre_archivo(nombre)

            nuevas = 0
            for q in cuotas:
                fname = carpeta / f"Recibo_{nombre_arc}_{q['fecha']}_Cuota{q['num']}.{ext}"
                if fname.exists():
                    continue  # ya fue generada/enviada antes
                html_content = _html(nombre_display, moto, placa, q["num"], q["fecha"], q["total"])

                if tiene_playwright:
                    with tempfile.NamedTemporaryFile(suffix=".html", delete=False,
                                                     mode="w", encoding="utf-8") as tmp:
                        tmp.write(html_content)
                        tmp_path = tmp.name
                    try:
                        page = browser.new_page()
                        page.goto(f"file:///{tmp_path.replace(os.sep, '/')}")
                        page.pdf(
                            path=str(fname),
                            format="A4",
                            margin={"top": "12mm", "bottom": "12mm",
                                    "left": "14mm", "right": "14mm"},
                            print_background=True,
                        )
                        page.close()
                    finally:
                        try: os.unlink(tmp_path)
                        except Exception: pass
                else:
                    fname.write_text(html_content, encoding="utf-8")

                total_ok += 1
                nuevas += 1

            if nuevas:
                print(f"  ✓ {nombre:<35} {nuevas:>3} nuevas")

    if tiene_playwright:
        with sync_playwright() as pw:
            browser = pw.chromium.launch()
            _generar(browser)
            browser.close()
    else:
        _generar(None)

    return total_ok, total_skip

# ── Punto de entrada ─────────────────────────────────────────────────────────
def main():
    hoy = date.today()
    print(f"{'='*60}")
    print(f"  Generador de facturas — Severa Motos")
    print(f"  Fecha: {hoy}  |  Salida: {SALIDA}")
    print(f"{'='*60}")

    # Cargar datos (igual que procesar_motos_tio.py main)
    ruta_excel = pmt.resolver_ruta_excel()
    print(f"\nLeyendo Excel: {ruta_excel}")
    wb = openpyxl.load_workbook(ruta_excel, data_only=True)
    clientes_meta, tel_por_placa = pmt.cargar_clientes(wb)
    filas = pmt.cargar_registro(wb)
    print(f"  {len(filas)} filas / {len(clientes_meta)} clientes")

    # Sheets (opcional)
    import urllib.request as _ur, json as _js
    ultimos_sheets = {}
    if pmt.GOOGLE_SHEET_ID:
        try:
            url = f"https://docs.google.com/spreadsheets/d/{pmt.GOOGLE_SHEET_ID}/export?format=csv&gid={pmt.GOOGLE_SHEET_GID}"
            filas_online = pmt.cargar_registro_sheets(pmt.GOOGLE_SHEET_ID, pmt.GOOGLE_SHEET_GID)
            if filas_online:
                for f in filas_online:
                    if (f.get("pago_recibido") or 0) > 0:
                        ck = f["cliente_clave"]
                        if ck not in ultimos_sheets or f["fecha"] > ultimos_sheets[ck]:
                            ultimos_sheets[ck] = f["fecha"]
                cuotas_online = pmt.agrupar_en_cuotas(filas_online)
                ultimo_excel = {}
                for f in filas:
                    ck = f["cliente_clave"]
                    if ck not in ultimo_excel or f["fecha"] > ultimo_excel[ck]:
                        ultimo_excel[ck] = f["fecha"]
                nuevos = [f for f in cuotas_online
                          if f["fecha"] > ultimo_excel.get(f["cliente_clave"], date.min)]
                filas.extend(nuevos)
                print(f"  + {len(nuevos)} cuotas desde Google Sheets")
        except Exception as e:
            print(f"  [!] Sheets no disponible: {e}")

    print(f"\nGenerando facturas...")
    total_ok_all   = 0
    total_skip_all = 0

    for empresa in ('severa', 'luna', 'jomar'):
        filas_emp    = [f for f in filas if pmt._empresa_de(f["cliente_clave"], f.get("placa")) == empresa]
        clientes_emp = {k: v for k, v in clientes_meta.items()
                        if pmt._empresa_de(k, v.get("placa")) == empresa}
        if not filas_emp and not clientes_emp:
            continue
        datos_emp = pmt.procesar(filas_emp, clientes_emp, hoy, tel_por_placa, ultimos_sheets)
        ok, skip  = generar_facturas(filas_emp, datos_emp)
        total_ok_all   += ok
        total_skip_all += skip

    # Contar total de PDFs existentes en la carpeta
    total_existentes = sum(1 for p in SALIDA.rglob("*.pdf"))
    print(f"\n{'='*60}")
    if total_ok_all:
        print(f"  Nuevas generadas   : {total_ok_all}")
    else:
        print(f"  Sin facturas nuevas — todo al día ✓")
    print(f"  Total acumulado    : {total_existentes} facturas")
    print(f"  Carpeta            : {SALIDA}")
    print(f"{'='*60}\n")

if __name__ == "__main__":
    main()
