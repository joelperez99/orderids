import os
from playwright.sync_api import sync_playwright

def is_streamlit_cloud() -> bool:
    # Señales comunes en Streamlit Cloud
    return os.getenv("STREAMLIT_SERVER_HEADLESS") == "true" or os.getenv("HOSTNAME", "").startswith("streamlit")

def launch_ctx(p):
    cloud = is_streamlit_cloud()

    return p.chromium.launch_persistent_context(
        user_data_dir=str(PROFILE_DIR),
        headless=True if cloud else False,  # <- Cloud headless, Local visible
        viewport={"width": 1400, "height": 900},
        args=[
            "--no-sandbox",
            "--disable-dev-shm-usage",
        ] if cloud else None,
    )

def open_browser_and_login():
    with sync_playwright() as p:
        ctx = launch_ctx(p)
        page = ctx.new_page()
        page.goto(ML_URL, wait_until="domcontentloaded")
        # En cloud no podrás hacer login interactivo, pero al menos no truena
        ctx.close()
