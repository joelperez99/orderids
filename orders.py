import os
import re
import time
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st

# =========================
# CONFIG
# =========================
ML_URL = "https://www.mercadolibre.com.mx/ventas/omni/listado"
PROFILE_DIR = Path(".ml_profile")
PROFILE_DIR.mkdir(exist_ok=True)

# =========================
# UI BOOT (para evitar "pantalla blanca")
# =========================
st.set_page_config(page_title="ML Orders → Excel", layout="wide")
st.title("Mercado Libre: Order IDs → cargos → Excel")
st.caption("Si ves esto, la app cargó bien. Si se queda en blanco, es crash antes de renderizar.")

def is_cloud_env() -> bool:
    # Heurística simple
    return (
        os.getenv("STREAMLIT_SERVER_HEADLESS") == "true"
        or os.getenv("HOSTNAME", "").startswith("streamlit")
        or os.getenv("HOME", "").startswith("/home/adminuser")
    )

CLOUD = is_cloud_env()
if CLOUD:
    st.warning(
        "Detecté entorno tipo Streamlit Cloud. Ahí Playwright puede fallar o colgarse. "
        "Para login/2FA, lo más estable es correrlo LOCAL en tu PC."
    )

# =========================
# HELPERS
# =========================
def parse_mxn(value: str) -> float | None:
    if value is None:
        return None
    s = value.strip().replace("MXN", "").replace("$", "").replace(" ", "")
    s = s.replace(",", "")
    s = re.sub(r"[^0-9\.\-]", "", s)
    if s in ("", "-", ".", "-."):
        return None
    try:
        return float(s)
    except Exception:
        return None

def clean_order_ids(raw: str) -> list[str]:
    if not raw:
        return []
    ids = re.findall(r"\d{8,}", raw)
    seen, out = set(), []
    for x in ids:
        if x not in seen:
            seen.add(x)
            out.append(x)
    return out

def label_value(page, label_text: str, timeout_ms: int = 8000) -> str | None:
    # Busca el label y extrae el valor monetario más probable del contenedor
    try:
        loc = page.locator(f"text={label_text}").first
        loc.wait_for(timeout=timeout_ms)
        container = loc.locator("xpath=ancestor::*[self::div or self::li][1]")
        text = container.inner_text().strip()
        m = re.findall(r"-?\$?\s?\d[\d,]*\.\d{2}", text)
        if m:
            return m[-1]
        m2 = re.findall(r"-?\d[\d,]*\.\d{2}", text)
        return m2[-1] if m2 else None
    except Exception:
        return None

def launch_ctx(p):
    # IMPORTANTE: args no-sandbox para Cloud
    args = ["--no-sandbox", "--disable-dev-shm-usage"] if CLOUD else None
    headless = True if CLOUD else False  # local visible, cloud headless
    return p.chromium.launch_persistent_context(
        user_data_dir=str(PROFILE_DIR),
        headless=headless,
        viewport={"width": 1400, "height": 900},
        timeout=60000,
        args=args,
    )

# =========================
# PLAYWRIGHT ACTIONS (imports dentro para evitar blanco)
# =========================
def open_browser_and_login():
    # Import dentro (clave para evitar pantalla blanca si playwright truena al importar)
    from playwright.sync_api import sync_playwright

    with sync_playwright() as p:
        ctx = launch_ctx(p)
        page = ctx.new_page()
        page.goto(ML_URL, wait_until="domcontentloaded")
        # En local: verás el browser y haces login/2FA.
        # En cloud: headless (no interactivo).
        ctx.close()

def fetch_orders(order_ids: list[str]) -> pd.DataFrame:
    from playwright.sync_api import sync_playwright

    rows = []
    with sync_playwright() as p:
        ctx = launch_ctx(p)
        page = ctx.new_page()
        page.goto(ML_URL, wait_until="domcontentloaded")

        # Espera buscador
        search = None
        try:
            search = page.get_by_placeholder(re.compile("Buscar", re.I)).first
            search.wait_for(timeout=20000)
        except Exception:
            search = page.locator("input[type='text']").first
            search.wait_for(timeout=20000)

        for oid in order_ids:
            # Buscar Order ID
            search.click()
            search.press("Control+A")
            search.type(oid, delay=25)
            time.sleep(0.9)

            # Abrir "Ver detalle"
            opened = False
            try:
                btn = page.get_by_role("button", name=re.compile("Ver detalle", re.I)).first
                btn.wait_for(timeout=12000)
                btn.click()
                opened = True
            except Exception:
                try:
                    page.get_by_text(re.compile("Ver detalle", re.I)).first.click(timeout=12000)
                    opened = True
                except Exception:
                    opened = False

            if not opened:
                rows.append({
                    "Order ID": oid,
                    "Monto": None,
                    "Comisión por venta": None,
                    "Cargo por envío": None,
                    "ISR 2.5%": None,
                })
                continue

            # Espera panel derecho
            try:
                page.locator("text=Precio del producto").first.wait_for(timeout=20000)
            except Exception:
                pass

            monto = parse_mxn(label_value(page, "Precio del producto"))
            comision = parse_mxn(label_value(page, "Cargos por venta"))
            envio = parse_mxn(label_value(page, "Envíos"))
            isr = parse_mxn(label_value(page, "Impuestos"))  # ISR 2.5% = Impuestos

            rows.append({
                "Order ID": oid,
                "Monto": monto,
                "Comisión por venta": comision,
                "Cargo por envío": envio,
                "ISR 2.5%": isr,
            })

            # Volver
            try:
                page.go_back(wait_until="domcontentloaded")
            except Exception:
                page.goto(ML_URL, wait_until="domcontentloaded")

            # Re-espera buscador
            try:
                search.wait_for(timeout=20000)
            except Exception:
                try:
                    search = page.get_by_placeholder(re.compile("Buscar", re.I)).first
                    search.wait_for(timeout=20000)
                except Exception:
                    search = page.locator("input[type='text']").first
                    search.wait_for(timeout=20000)

        ctx.close()

    df = pd.DataFrame(rows)
    df = df[["Order ID", "Monto", "Comisión por venta", "Cargo por envío", "ISR 2.5%"]]
    return df

def df_to_excel_bytes(df: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Ordenes")
        ws = writer.sheets["Ordenes"]

        # ancho + formato moneda
        for col in ws.columns:
            max_len = 0
            col_letter = col[0].column_letter
            for cell in col:
                v = "" if cell.value is None else str(cell.value)
                max_len = max(max_len, len(v))
            ws.column_dimensions[col_letter].width = min(max_len + 2, 28)

        # moneda en B-E
        for row in ws.iter_rows(min_row=2, min_col=2, max_col=5):
            for cell in row:
                cell.number_format = '"$"#,##0.00;-"$"#,##0.00'

    return output.getvalue()

# =========================
# UI CONTROLS
# =========================
st.markdown(
    """
**Columnas del Excel:**
- **Monto** = Precio del producto  
- **Comisión por venta** = Cargos por venta  
- **Cargo por envío** = Envíos  
- **ISR 2.5%** = Impuestos  
"""
)

col1, col2 = st.columns([1, 1])

with col1:
    if st.button("1) Abrir Mercado Libre (login manual)", use_container_width=True):
        try:
            open_browser_and_login()
            if CLOUD:
                st.info("Se ejecutó en headless (Cloud). Para login/2FA necesitas correrlo LOCAL.")
            else:
                st.success("Se abrió el navegador. Haz login + 2FA y regresa aquí.")
        except Exception as e:
            st.error("Falló al abrir navegador.")
            st.exception(e)

with col2:
    st.caption("Sesión se guarda en ./.ml_profile (cookies persistentes).")

raw_ids = st.text_area(
    "2) Pega Order IDs (uno por línea o separados por coma):",
    height=160,
    placeholder="2000011446863697\n2000014245438812\n..."
)
order_ids = clean_order_ids(raw_ids)
st.write(f"Order IDs detectados: **{len(order_ids)}**")

if st.button("3) Capturar información y generar Excel", type="primary", disabled=(len(order_ids) == 0)):
    try:
        with st.spinner("Capturando datos..."):
            df = fetch_orders(order_ids)

        st.success("Listo.")
        st.dataframe(df, use_container_width=True)

        xlsx = df_to_excel_bytes(df)
        st.download_button(
            "Descargar Excel",
            data=xlsx,
            file_name="ordenes_mercadolibre.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    except Exception as e:
        st.error("Falló la captura con Playwright.")
        st.exception(e)
