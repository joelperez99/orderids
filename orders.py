import re
import time
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeoutError

ML_URL = "https://www.mercadolibre.com.mx/ventas/omni/listado"
PROFILE_DIR = Path(".ml_profile")
PROFILE_DIR.mkdir(exist_ok=True)

def parse_mxn(value: str) -> float | None:
    if value is None:
        return None
    s = value.strip()
    s = s.replace("MXN", "").replace("$", "").replace(" ", "")
    s = s.replace(",", "")
    s = re.sub(r"[^0-9\.\-]", "", s)
    if s in ("", "-", ".", "-."):
        return None
    try:
        return float(s)
    except:
        return None

def clean_order_ids(raw: str) -> list[str]:
    if not raw:
        return []
    ids = re.findall(r"\d{8,}", raw)
    seen = set()
    out = []
    for x in ids:
        if x not in seen:
            seen.add(x)
            out.append(x)
    return out

def label_value(page, label_text: str, timeout_ms: int = 6000) -> str | None:
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
    except PWTimeoutError:
        return None
    except Exception:
        return None

def open_browser_and_login():
    with sync_playwright() as p:
        ctx = p.chromium.launch_persistent_context(
            user_data_dir=str(PROFILE_DIR),
            headless=False,
            viewport={"width": 1400, "height": 900},
        )
        page = ctx.new_page()
        page.goto(ML_URL, wait_until="domcontentloaded")
        st.info(
            "Se abrió el navegador. Haz login y completa 2FA MANUALMENTE. "
            "Cuando ya veas la pantalla de Ventas (con el buscador), regresa a Streamlit."
        )
        ctx.close()

def fetch_orders(order_ids: list[str]) -> pd.DataFrame:
    rows = []
    with sync_playwright() as p:
        ctx = p.chromium.launch_persistent_context(
            user_data_dir=str(PROFILE_DIR),
            headless=False,
            viewport={"width": 1400, "height": 900},
        )
        page = ctx.new_page()
        page.goto(ML_URL, wait_until="domcontentloaded")

        # buscador
        try:
            search = page.get_by_placeholder(re.compile("Buscar", re.I)).first
            search.wait_for(timeout=15000)
        except Exception:
            search = page.locator("input[type='text']").first
            search.wait_for(timeout=15000)

        for oid in order_ids:
            search.click()
            search.press("Control+A")
            search.type(oid, delay=25)
            time.sleep(0.8)

            # abre detalle
            opened = False
            try:
                page.get_by_role("button", name=re.compile("Ver detalle", re.I)).first.wait_for(timeout=10000)
                page.get_by_role("button", name=re.compile("Ver detalle", re.I)).first.click()
                opened = True
            except Exception:
                try:
                    page.get_by_text(re.compile("Ver detalle", re.I)).first.click(timeout=10000)
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
                    "Estatus": "No encontrado / No se pudo abrir detalle",
                })
                continue

            # espera panel derecho
            try:
                page.locator("text=Precio del producto").first.wait_for(timeout=15000)
            except Exception:
                pass

            # mapeo EXACTO a tu formato
            monto = parse_mxn(label_value(page, "Precio del producto"))
            comision = parse_mxn(label_value(page, "Cargos por venta"))
            envio = parse_mxn(label_value(page, "Envíos"))
            isr = parse_mxn(label_value(page, "Impuestos"))  # <-- ISR 2.5% = Impuestos

            # estatus (opcional)
            estatus = None
            try:
                for candidate in ["Entregado", "En camino", "Cancelada", "Devuelta", "Reclamo"]:
                    if page.locator(f"text={candidate}").count() > 0:
                        estatus = candidate
                        break
            except Exception:
                estatus = None

            rows.append({
                "Order ID": oid,
                "Monto": monto,
                "Comisión por venta": comision,
                "Cargo por envío": envio,
                "ISR 2.5%": isr,
                "Estatus": estatus,
            })

            page.go_back(wait_until="domcontentloaded")
            try:
                search.wait_for(timeout=15000)
            except Exception:
                pass

        ctx.close()

    df = pd.DataFrame(rows)

    # Si quieres EXACTAMENTE como tu imagen (sin Estatus), descomenta:
    # df = df[["Order ID", "Monto", "Comisión por venta", "Cargo por envío", "ISR 2.5%"]]

    return df

def df_to_excel_bytes(df: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Ordenes")
        ws = writer.sheets["Ordenes"]
        for col in ws.columns:
            max_len = 0
            col_letter = col[0].column_letter
            for cell in col:
                v = "" if cell.value is None else str(cell.value)
                max_len = max(max_len, len(v))
            ws.column_dimensions[col_letter].width = min(max_len + 2, 35)
    return output.getvalue()

st.set_page_config(page_title="ML Orders → Excel", layout="wide")
st.title("Mercado Libre: cargos por Order ID → Excel")

st.markdown(
    """
**Columnas finales:**
- **Monto** = Precio del producto  
- **Comisión por venta** = Cargos por venta  
- **Cargo por envío** = Envíos  
- **ISR 2.5%** = **Impuestos** (tal cual tu captura)
"""
)

col1, col2 = st.columns([1, 1])
with col1:
    if st.button("1) Abrir navegador para iniciar sesión (manual)", use_container_width=True):
        open_browser_and_login()
with col2:
    st.caption("La sesión se guarda en ./.ml_profile para reutilizar cookies.")

raw_ids = st.text_area("2) Pega Order IDs:", height=150, placeholder="2000011446863697\n2000014245438812\n...")
order_ids = clean_order_ids(raw_ids)
st.write(f"Order IDs detectados: **{len(order_ids)}**")

if st.button("3) Capturar información", type="primary", disabled=(len(order_ids) == 0)):
    with st.spinner("Abriendo órdenes y capturando datos..."):
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
