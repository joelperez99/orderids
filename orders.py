import json
import os
import re
import subprocess
import sys
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
SESSION_FILE = Path("ml_session.json")

# =========================
# UI BOOT
# =========================
st.set_page_config(page_title="ML Orders → Excel", layout="wide")
st.title("Mercado Libre: Order IDs → cargos → Excel")


# =========================
# INSTALAR BROWSERS UNA SOLA VEZ
# =========================
@st.cache_resource(show_spinner="Instalando navegador Chromium (solo la primera vez)...")
def install_playwright_browsers():
    try:
        result = subprocess.run(
            [sys.executable, "-m", "playwright", "install", "chromium"],
            capture_output=True,
            text=True,
            timeout=180,
        )
        if result.returncode != 0:
            return False, result.stderr
        return True, "OK"
    except Exception as e:
        return False, str(e)


ok, msg = install_playwright_browsers()
if not ok:
    st.error(f"No se pudo instalar Chromium:\n{msg}")
    st.stop()


# =========================
# DETECTAR ENTORNO
# =========================
def is_cloud_env() -> bool:
    return (
        os.getenv("STREAMLIT_SERVER_HEADLESS") == "true"
        or os.getenv("HOSTNAME", "").startswith("streamlit")
        or os.getenv("HOME", "").startswith("/home/adminuser")
    )


CLOUD = is_cloud_env()

if CLOUD:
    st.info(
        "☁️ **Entorno Cloud detectado.** El navegador corre en modo headless (sin ventana visible).  \n"
        "Para hacer login con 2FA: corre la app **localmente**, haz login, exporta la sesión "
        "con el botón de abajo y sube el archivo `ml_session.json` a tu repositorio."
    )
else:
    st.success("💻 Entorno **local** detectado. El navegador se abrirá de forma visible para que hagas login.")


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


def build_launch_kwargs() -> dict:
    """Parámetros de lanzamiento según entorno."""
    kwargs = dict(
        user_data_dir=str(PROFILE_DIR),
        headless=CLOUD,
        viewport={"width": 1400, "height": 900},
        timeout=60000,
        args=["--no-sandbox", "--disable-dev-shm-usage"],
    )
    # Si existe archivo de sesión exportado, úsalo
    if SESSION_FILE.exists():
        kwargs["storage_state"] = str(SESSION_FILE)
    return kwargs


# =========================
# PLAYWRIGHT: LOGIN
# =========================
def open_browser_and_login():
    from playwright.sync_api import sync_playwright

    with sync_playwright() as p:
        kwargs = build_launch_kwargs()
        ctx = p.chromium.launch_persistent_context(**kwargs)
        page = ctx.new_page()
        page.goto(ML_URL, wait_until="domcontentloaded")

        if not CLOUD:
            # Espera hasta 3 minutos para que el usuario haga login manual
            st.info("🌐 Navegador abierto. Haz login y completa el 2FA. Tienes 3 minutos.")
            page.wait_for_timeout(180_000)

        # Exportar sesión al terminar (útil para subir a Cloud después)
        storage = ctx.storage_state()
        with open(SESSION_FILE, "w") as f:
            json.dump(storage, f)

        ctx.close()

    return storage


# =========================
# PLAYWRIGHT: FETCH ORDERS
# =========================
def fetch_orders(order_ids: list[str]) -> pd.DataFrame:
    from playwright.sync_api import sync_playwright

    rows = []
    with sync_playwright() as p:
        kwargs = build_launch_kwargs()
        ctx = p.chromium.launch_persistent_context(**kwargs)
        page = ctx.new_page()
        page.goto(ML_URL, wait_until="domcontentloaded")

        # Localizar buscador
        search = None
        for attempt in range(2):
            try:
                search = page.get_by_placeholder(re.compile("Buscar", re.I)).first
                search.wait_for(timeout=20000)
                break
            except Exception:
                try:
                    search = page.locator("input[type='text']").first
                    search.wait_for(timeout=20000)
                    break
                except Exception:
                    if attempt == 1:
                        raise RuntimeError(
                            "No se encontró el buscador. Probablemente la sesión expiró. "
                            "Vuelve a hacer login."
                        )

        for oid in order_ids:
            try:
                # Buscar Order ID
                search.click()
                search.press("Control+A")
                search.type(oid, delay=25)
                time.sleep(0.9)

                # Abrir detalle
                opened = False
                for selector in [
                    lambda: page.get_by_role("button", name=re.compile("Ver detalle", re.I)).first,
                    lambda: page.get_by_text(re.compile("Ver detalle", re.I)).first,
                ]:
                    try:
                        el = selector()
                        el.wait_for(timeout=12000)
                        el.click()
                        opened = True
                        break
                    except Exception:
                        continue

                if not opened:
                    rows.append(_empty_row(oid, "No se encontró botón 'Ver detalle'"))
                    continue

                # Esperar panel de detalle
                try:
                    page.locator("text=Precio del producto").first.wait_for(timeout=20000)
                except Exception:
                    pass

                monto     = parse_mxn(label_value(page, "Precio del producto"))
                comision  = parse_mxn(label_value(page, "Cargos por venta"))
                envio     = parse_mxn(label_value(page, "Envíos"))
                isr       = parse_mxn(label_value(page, "Impuestos"))

                rows.append({
                    "Order ID": oid,
                    "Monto": monto,
                    "Comisión por venta": comision,
                    "Cargo por envío": envio,
                    "ISR 2.5%": isr,
                    "Error": None,
                })

            except Exception as e:
                rows.append(_empty_row(oid, str(e)))

            # Volver a la lista
            try:
                page.go_back(wait_until="domcontentloaded")
            except Exception:
                page.goto(ML_URL, wait_until="domcontentloaded")

            # Re-localizar buscador
            try:
                search.wait_for(timeout=20000)
            except Exception:
                try:
                    search = page.get_by_placeholder(re.compile("Buscar", re.I)).first
                    search.wait_for(timeout=15000)
                except Exception:
                    search = page.locator("input[type='text']").first

        ctx.close()

    df = pd.DataFrame(rows)
    return df[["Order ID", "Monto", "Comisión por venta", "Cargo por envío", "ISR 2.5%", "Error"]]


def _empty_row(oid: str, error: str) -> dict:
    return {
        "Order ID": oid,
        "Monto": None,
        "Comisión por venta": None,
        "Cargo por envío": None,
        "ISR 2.5%": None,
        "Error": error,
    }


# =========================
# EXCEL EXPORT
# =========================
def df_to_excel_bytes(df: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Ordenes")
        ws = writer.sheets["Ordenes"]

        for col in ws.columns:
            max_len = max((len(str(c.value or "")) for c in col), default=10)
            ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 30)

        # Formato moneda columnas B-E (Monto, Comisión, Envío, ISR)
        for row in ws.iter_rows(min_row=2, min_col=2, max_col=5):
            for cell in row:
                cell.number_format = '"$"#,##0.00;-"$"#,##0.00'

    return output.getvalue()


# =========================
# UI
# =========================
st.markdown("""
**Columnas del Excel generado:**
| Columna | Fuente en ML |
|---|---|
| Monto | Precio del producto |
| Comisión por venta | Cargos por venta |
| Cargo por envío | Envíos |
| ISR 2.5% | Impuestos |
""")

st.divider()

# --- PASO 1: LOGIN ---
st.subheader("Paso 1 · Sesión de Mercado Libre")

col1, col2 = st.columns(2)

with col1:
    session_exists = SESSION_FILE.exists()
    status_label = "✅ Sesión guardada (`ml_session.json` encontrado)" if session_exists else "⚠️ Sin sesión guardada"
    st.caption(status_label)

    if st.button("Abrir navegador y hacer login", use_container_width=True, disabled=CLOUD):
        with st.spinner("Abriendo navegador..."):
            try:
                storage = open_browser_and_login()
                st.success("Login completado y sesión guardada en `ml_session.json`.")
                st.download_button(
                    "⬇️ Descargar ml_session.json (para subir a Cloud)",
                    data=json.dumps(storage, indent=2),
                    file_name="ml_session.json",
                    mime="application/json",
                )
            except Exception as e:
                st.error("Falló al abrir navegador.")
                st.exception(e)

with col2:
    st.caption("Sube aquí tu `ml_session.json` si lo tienes de una sesión local previa:")
    uploaded = st.file_uploader("Subir ml_session.json", type="json", label_visibility="collapsed")
    if uploaded:
        SESSION_FILE.write_bytes(uploaded.read())
        st.success("Sesión cargada correctamente.")
        st.rerun()

st.divider()

# --- PASO 2: ORDER IDs ---
st.subheader("Paso 2 · Ingresar Order IDs")

raw_ids = st.text_area(
    "Pega los Order IDs (uno por línea, o separados por comas/espacios):",
    height=160,
    placeholder="2000011446863697\n2000014245438812\n...",
)
order_ids = clean_order_ids(raw_ids)
st.caption(f"Order IDs detectados: **{len(order_ids)}**")

st.divider()

# --- PASO 3: CAPTURAR ---
st.subheader("Paso 3 · Capturar y exportar")

if not SESSION_FILE.exists() and CLOUD:
    st.warning("⚠️ En Cloud necesitas subir un `ml_session.json` válido antes de capturar.")

if st.button(
    "▶️ Capturar información y generar Excel",
    type="primary",
    disabled=(len(order_ids) == 0),
    use_container_width=True,
):
    try:
        progress = st.progress(0, text="Iniciando...")
        status_box = st.empty()

        with st.spinner(f"Capturando {len(order_ids)} órdenes..."):
            df = fetch_orders(order_ids)

        progress.progress(100, text="¡Listo!")
        st.success(f"✅ {len(df)} órdenes procesadas.")

        # Mostrar errores si hubo
        errores = df[df["Error"].notna()]
        if not errores.empty:
            st.warning(f"⚠️ {len(errores)} órdenes con error:")
            st.dataframe(errores[["Order ID", "Error"]], use_container_width=True)

        st.dataframe(df.drop(columns=["Error"]), use_container_width=True)

        xlsx = df_to_excel_bytes(df.drop(columns=["Error"]))
        st.download_button(
            "⬇️ Descargar Excel",
            data=xlsx,
            file_name="ordenes_mercadolibre.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    except Exception as e:
        st.error("Falló la captura con Playwright.")
        st.exception(e)
