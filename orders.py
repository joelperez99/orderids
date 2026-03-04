import re
import subprocess
import sys
import time
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st

ML_URL = "https://www.mercadolibre.com.mx/ventas/omni/listado"
PROFILE_DIR = Path(".ml_profile")
PROFILE_DIR.mkdir(exist_ok=True)

st.set_page_config(page_title="ML Orders → Excel", layout="wide")
st.title("Mercado Libre: Order IDs → Excel")


# ── Instalar navegador una sola vez ──────────────────────────────────────────
@st.cache_resource(show_spinner="Instalando navegador...")
def install_playwright():
    r = subprocess.run(
        [sys.executable, "-m", "playwright", "install", "chromium"],
        capture_output=True, text=True, timeout=180,
    )
    return r.returncode == 0, r.stderr

ok, err = install_playwright()
if not ok:
    st.error(f"No se pudo instalar el navegador:\n{err}")
    st.stop()


# ── Helpers ───────────────────────────────────────────────────────────────────
def parse_mxn(value):
    if not value:
        return None
    s = re.sub(r"[^0-9.\-]", "", value.replace(",", ""))
    try:
        return float(s) if s not in ("", "-", ".", "-.") else None
    except Exception:
        return None

def clean_order_ids(raw):
    ids = re.findall(r"\d{8,}", raw or "")
    seen, out = set(), []
    for x in ids:
        if x not in seen:
            seen.add(x); out.append(x)
    return out

def label_value(page, label_text, timeout_ms=8000):
    try:
        loc = page.locator(f"text={label_text}").first
        loc.wait_for(timeout=timeout_ms)
        container = loc.locator("xpath=ancestor::*[self::div or self::li][1]")
        text = container.inner_text().strip()
        m = re.findall(r"-?\$?\s?\d[\d,]*\.\d{2}", text)
        if m: return m[-1]
        m2 = re.findall(r"-?\d[\d,]*\.\d{2}", text)
        return m2[-1] if m2 else None
    except Exception:
        return None

def empty_row(oid, error):
    return {"Order ID": oid, "Monto": None, "Comisión por venta": None,
            "Cargo por envío": None, "ISR 2.5%": None, "Error": error}


# ── Abrir navegador y esperar login manual ────────────────────────────────────
def abrir_y_login():
    from playwright.sync_api import sync_playwright
    with sync_playwright() as p:
        ctx = p.chromium.launch_persistent_context(
            user_data_dir=str(PROFILE_DIR),
            headless=False,                          # ← ventana VISIBLE
            args=["--no-sandbox", "--disable-dev-shm-usage"],
            viewport={"width": 1280, "height": 800},
        )
        page = ctx.new_page()
        page.goto("https://www.mercadolibre.com.mx", wait_until="domcontentloaded")

        # Espera hasta que el usuario llegue a la página de ventas (máx 5 min)
        st.info("🌐 Navegador abierto. Inicia sesión y cuando estés dentro de ML haz clic en **'Listo, continuar'** abajo.")
        page.wait_for_url("**/mercadolibre.com.mx/**", timeout=300_000)

        # Guardar cookies para reutilizar
        ctx.storage_state(path=str(PROFILE_DIR / "session.json"))
        ctx.close()


# ── Capturar órdenes ──────────────────────────────────────────────────────────
def fetch_orders(order_ids):
    from playwright.sync_api import sync_playwright

    session_file = PROFILE_DIR / "session.json"
    rows = []

    with sync_playwright() as p:
        launch_kwargs = dict(
            user_data_dir=str(PROFILE_DIR),
            headless=False,                          # ← visible para que veas qué pasa
            args=["--no-sandbox", "--disable-dev-shm-usage"],
            viewport={"width": 1280, "height": 800},
        )
        if session_file.exists():
            launch_kwargs["storage_state"] = str(session_file)

        ctx = p.chromium.launch_persistent_context(**launch_kwargs)
        page = ctx.new_page()
        page.goto(ML_URL, wait_until="domcontentloaded")

        # Si pide login, esperar que el usuario lo haga
        if "login" in page.url.lower() or "identificate" in page.url.lower():
            st.warning("⚠️ Se requiere login. Inicia sesión en el navegador que se abrió y luego la app continuará automáticamente.")
            page.wait_for_url("**/mercadolibre.com.mx/**", timeout=300_000)
            page.goto(ML_URL, wait_until="domcontentloaded")

        # Localizar buscador
        try:
            search = page.get_by_placeholder(re.compile("Buscar", re.I)).first
            search.wait_for(timeout=20_000)
        except Exception:
            search = page.locator("input[type='text']").first
            search.wait_for(timeout=20_000)

        total = len(order_ids)
        bar = st.progress(0, text="Iniciando...")

        for i, oid in enumerate(order_ids, 1):
            bar.progress(i / total, text=f"Procesando {i}/{total}: {oid}")
            try:
                search.click()
                search.press("Control+A")
                search.type(oid, delay=30)
                time.sleep(0.9)

                opened = False
                for fn in [
                    lambda: page.get_by_role("button", name=re.compile("Ver detalle", re.I)).first,
                    lambda: page.get_by_text(re.compile("Ver detalle", re.I)).first,
                ]:
                    try:
                        el = fn(); el.wait_for(timeout=12_000); el.click()
                        opened = True; break
                    except Exception:
                        continue

                if not opened:
                    rows.append(empty_row(oid, "No se encontró 'Ver detalle'")); continue

                try:
                    page.locator("text=Precio del producto").first.wait_for(timeout=20_000)
                except Exception:
                    pass

                rows.append({
                    "Order ID":           oid,
                    "Monto":              parse_mxn(label_value(page, "Precio del producto")),
                    "Comisión por venta": parse_mxn(label_value(page, "Cargos por venta")),
                    "Cargo por envío":    parse_mxn(label_value(page, "Envíos")),
                    "ISR 2.5%":           parse_mxn(label_value(page, "Impuestos")),
                    "Error":              None,
                })

            except Exception as e:
                rows.append(empty_row(oid, str(e)))

            try:
                page.go_back(wait_until="domcontentloaded")
            except Exception:
                page.goto(ML_URL, wait_until="domcontentloaded")

            try:
                search.wait_for(timeout=15_000)
            except Exception:
                try:
                    search = page.get_by_placeholder(re.compile("Buscar", re.I)).first
                    search.wait_for(timeout=15_000)
                except Exception:
                    search = page.locator("input[type='text']").first

        ctx.close()
        bar.progress(1.0, text="¡Listo!")

    return pd.DataFrame(rows)[["Order ID", "Monto", "Comisión por venta", "Cargo por envío", "ISR 2.5%", "Error"]]


# ── Excel ─────────────────────────────────────────────────────────────────────
def to_excel(df):
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name="Ordenes")
        ws = w.sheets["Ordenes"]
        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = min(
                max(len(str(c.value or "")) for c in col) + 2, 30)
        for row in ws.iter_rows(min_row=2, min_col=2, max_col=5):
            for cell in row:
                cell.number_format = '"$"#,##0.00;-"$"#,##0.00'
    return out.getvalue()


# ── UI ────────────────────────────────────────────────────────────────────────
st.markdown("""
| Columna | Fuente en ML |
|---|---|
| Monto | Precio del producto |
| Comisión por venta | Cargos por venta |
| Cargo por envío | Envíos |
| ISR 2.5% | Impuestos |
""")

st.divider()

# Paso 1: login
st.subheader("Paso 1 · Login (solo la primera vez)")
col1, col2 = st.columns(2)
with col1:
    if st.button("🌐 Abrir navegador para iniciar sesión", use_container_width=True):
        try:
            abrir_y_login()
            st.success("✅ Sesión guardada. Ya puedes capturar órdenes.")
        except Exception as e:
            st.error("Algo falló al abrir el navegador.")
            st.exception(e)
with col2:
    st.caption(
        "Se abrirá Chrome visible. Escribe tu correo, contraseña y 2FA normalmente. "
        "La sesión se guarda automáticamente para la próxima vez."
    )

st.divider()

# Paso 2: capturar
st.subheader("Paso 2 · Capturar órdenes")

raw_ids = st.text_area(
    "Pega los Order IDs (uno por línea o separados por comas):",
    height=160,
    placeholder="2000011446863697\n2000014245438812\n...",
)
order_ids = clean_order_ids(raw_ids)
st.caption(f"Order IDs detectados: **{len(order_ids)}**")

if st.button("▶️ Capturar y generar Excel", type="primary",
             disabled=len(order_ids) == 0, use_container_width=True):
    try:
        df = fetch_orders(order_ids)
        st.success(f"✅ {len(df)} órdenes procesadas.")

        errores = df[df["Error"].notna()]
        if not errores.empty:
            st.warning(f"⚠️ {len(errores)} con error:")
            st.dataframe(errores[["Order ID", "Error"]], use_container_width=True)

        df_show = df.drop(columns=["Error"])
        st.dataframe(df_show, use_container_width=True)

        st.download_button(
            "⬇️ Descargar Excel",
            data=to_excel(df_show),
            file_name="ordenes_mercadolibre.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    except Exception as e:
        st.error("Falló la captura.")
        st.exception(e)
