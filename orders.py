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

# Ruta al perfil de Chrome donde ya tienes sesión de ML iniciada.
# Ajusta según tu sistema operativo:
#   Windows : Path.home() / "AppData/Local/Google/Chrome/User Data"
#   macOS   : Path.home() / "Library/Application Support/Google/Chrome"
#   Linux   : Path.home() / ".config/google-chrome"
CHROME_PROFILE = Path.home() / ".config" / "google-chrome"
CHROME_PROFILE_NAME = "Default"   # cambia si usas "Profile 1", etc.

# =========================
# UI
# =========================
st.set_page_config(page_title="ML Orders → Excel", layout="wide")
st.title("Mercado Libre: Order IDs → Excel")


# =========================
# INSTALAR CHROMIUM UNA SOLA VEZ
# =========================
@st.cache_resource(show_spinner="Preparando navegador...")
def install_playwright():
    result = subprocess.run(
        [sys.executable, "-m", "playwright", "install", "chromium"],
        capture_output=True, text=True, timeout=180,
    )
    if result.returncode != 0:
        return False, result.stderr
    return True, "OK"


ok, msg = install_playwright()
if not ok:
    st.error(f"No se pudo instalar Chromium:\n{msg}")
    st.stop()


# =========================
# HELPERS
# =========================
def parse_mxn(value: str) -> float | None:
    if value is None:
        return None
    s = re.sub(r"[^0-9.\-]", "", value.replace(",", ""))
    if s in ("", "-", ".", "-."):
        return None
    try:
        return float(s)
    except Exception:
        return None


def clean_order_ids(raw: str) -> list[str]:
    ids = re.findall(r"\d{8,}", raw or "")
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


def _empty_row(oid: str, error: str) -> dict:
    return {
        "Order ID": oid, "Monto": None, "Comisión por venta": None,
        "Cargo por envío": None, "ISR 2.5%": None, "Error": error,
    }


# =========================
# FETCH ORDERS — usa perfil existente, sin login
# =========================
def fetch_orders(order_ids: list[str], profile_path: Path, profile_name: str) -> pd.DataFrame:
    from playwright.sync_api import sync_playwright

    rows = []
    with sync_playwright() as p:
        ctx = p.chromium.launch_persistent_context(
            user_data_dir=str(profile_path),
            channel="chrome",                        # usa el Chrome instalado en el sistema
            headless=True,                           # sin ventana, no interfiere con Chrome abierto
            args=[
                "--no-sandbox",
                "--disable-dev-shm-usage",
                f"--profile-directory={profile_name}",
            ],
            viewport={"width": 1400, "height": 900},
            timeout=60_000,
        )

        page = ctx.new_page()
        page.goto(ML_URL, wait_until="domcontentloaded")

        # Verificar sesión activa
        if "login" in page.url.lower() or "identificate" in page.url.lower():
            ctx.close()
            raise RuntimeError(
                "La sesión expiró. Abre Chrome, entra a ML manualmente y vuelve a intentarlo."
            )

        # Localizar buscador
        try:
            search = page.get_by_placeholder(re.compile("Buscar", re.I)).first
            search.wait_for(timeout=20_000)
        except Exception:
            search = page.locator("input[type='text']").first
            search.wait_for(timeout=20_000)

        total = len(order_ids)
        progress_bar = st.progress(0, text="Iniciando...")

        for i, oid in enumerate(order_ids, 1):
            progress_bar.progress(i / total, text=f"Procesando {i}/{total}:  {oid}")

            try:
                search.click()
                search.press("Control+A")
                search.type(oid, delay=25)
                time.sleep(0.8)

                # Abrir detalle
                opened = False
                for fn in [
                    lambda: page.get_by_role("button", name=re.compile("Ver detalle", re.I)).first,
                    lambda: page.get_by_text(re.compile("Ver detalle", re.I)).first,
                ]:
                    try:
                        el = fn()
                        el.wait_for(timeout=12_000)
                        el.click()
                        opened = True
                        break
                    except Exception:
                        continue

                if not opened:
                    rows.append(_empty_row(oid, "No se encontró 'Ver detalle'"))
                    continue

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
                rows.append(_empty_row(oid, str(e)))

            # Regresar a la lista
            try:
                page.go_back(wait_until="domcontentloaded")
            except Exception:
                page.goto(ML_URL, wait_until="domcontentloaded")

            # Re-localizar buscador tras volver
            try:
                search.wait_for(timeout=15_000)
            except Exception:
                try:
                    search = page.get_by_placeholder(re.compile("Buscar", re.I)).first
                    search.wait_for(timeout=15_000)
                except Exception:
                    search = page.locator("input[type='text']").first

        ctx.close()
        progress_bar.progress(1.0, text="¡Listo!")

    df = pd.DataFrame(rows)
    return df[["Order ID", "Monto", "Comisión por venta", "Cargo por envío", "ISR 2.5%", "Error"]]


# =========================
# EXCEL
# =========================
def df_to_excel_bytes(df: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Ordenes")
        ws = writer.sheets["Ordenes"]
        for col in ws.columns:
            max_len = max((len(str(c.value or "")) for c in col), default=10)
            ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 30)
        for row in ws.iter_rows(min_row=2, min_col=2, max_col=5):
            for cell in row:
                cell.number_format = '"$"#,##0.00;-"$"#,##0.00'
    return output.getvalue()


# =========================
# UI
# =========================
st.markdown("""
| Columna Excel | Fuente en Mercado Libre |
|---|---|
| Monto | Precio del producto |
| Comisión por venta | Cargos por venta |
| Cargo por envío | Envíos |
| ISR 2.5% | Impuestos |
""")

st.divider()

# Configuración del perfil
with st.expander("⚙️ Configuración de perfil de Chrome", expanded=False):
    profile_path_str = st.text_input(
        "Ruta al directorio de perfiles de Chrome:",
        value=str(CHROME_PROFILE),
        help=(
            "Windows: C:/Users/USUARIO/AppData/Local/Google/Chrome/User Data  |  "
            "macOS: /Users/USUARIO/Library/Application Support/Google/Chrome  |  "
            "Linux: /home/USUARIO/.config/google-chrome"
        ),
    )
    profile_name_str = st.text_input(
        "Nombre del perfil:",
        value=CHROME_PROFILE_NAME,
        help="Normalmente 'Default'. Si usas múltiples perfiles puede ser 'Profile 1', etc.",
    )

profile_path = Path(profile_path_str)

if not profile_path.exists():
    st.warning(
        f"⚠️ Directorio no encontrado: `{profile_path}`  \n"
        "Edita la ruta en ⚙️ **Configuración de perfil de Chrome** arriba."
    )

st.subheader("Order IDs")
raw_ids = st.text_area(
    "Pega los Order IDs (uno por línea o separados por comas):",
    height=160,
    placeholder="2000011446863697\n2000014245438812\n...",
)
order_ids = clean_order_ids(raw_ids)
st.caption(f"Order IDs detectados: **{len(order_ids)}**")

st.divider()

if st.button(
    "▶️ Capturar y generar Excel",
    type="primary",
    disabled=(len(order_ids) == 0 or not profile_path.exists()),
    use_container_width=True,
):
    try:
        df = fetch_orders(order_ids, profile_path, profile_name_str)

        st.success(f"✅ {len(df)} órdenes procesadas.")

        errores = df[df["Error"].notna()]
        if not errores.empty:
            st.warning(f"⚠️ {len(errores)} órdenes con error:")
            st.dataframe(errores[["Order ID", "Error"]], use_container_width=True)

        df_show = df.drop(columns=["Error"])
        st.dataframe(df_show, use_container_width=True)

        xlsx = df_to_excel_bytes(df_show)
        st.download_button(
            "⬇️ Descargar Excel",
            data=xlsx,
            file_name="ordenes_mercadolibre.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    except RuntimeError as e:
        st.error(str(e))
    except Exception as e:
        st.error("Falló la captura.")
        st.exception(e)
