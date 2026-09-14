"""
VALIDAR_OC_6PM.py
Ejecuta la validacion de OC vs Formularios a las 6 PM,
toma screenshot del dashboard y lo envia por WhatsApp a Carlos.
"""
import subprocess, sys, os, time, pathlib
from datetime import datetime, timedelta

PYTHON = r"C:\Users\trujillo.c\AppData\Local\Python\pythoncore-3.14-64\python.exe"
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
VALIDAR = os.path.join(SCRIPT_DIR, "VALIDAR_OC_FORMULARIOS.py")
HTML_OUTPUT = os.path.join(SCRIPT_DIR, "VALIDACION_OC_FORMULARIOS.html")
SCREENSHOT = os.path.join(SCRIPT_DIR, "VALIDACION_OC_FORMULARIOS.png")
ENVIAR_WA = r"C:\Users\trujillo.c\OneDrive - Me Llega, S.A\Escritorio\_CASTRO_AGENTE\enviar_imagen_wa.py"
DESTINO_CARLOS = "50237639994"

LOG_FILE = os.path.join(SCRIPT_DIR, "VALIDAR_OC_6PM.log")

def log(msg):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] {msg}"
    print(line)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(line + "\n")

def tomar_screenshot():
    from selenium import webdriver
    options = webdriver.EdgeOptions()
    options.add_argument("--headless=new")
    options.add_argument("--window-size=1400,900")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")

    driver = webdriver.Edge(options=options)
    file_url = pathlib.Path(HTML_OUTPUT).as_uri()
    driver.get(file_url)
    time.sleep(3)

    height = driver.execute_script("return document.body.scrollHeight")
    driver.set_window_size(1400, min(height + 120, 10000))
    time.sleep(1)
    driver.save_screenshot(SCREENSHOT)
    driver.quit()
    return os.path.exists(SCREENSHOT)


def main():
    manana = (datetime.now() + timedelta(days=1)).strftime("%d/%m/%Y")
    log(f"=== Inicio validacion OC para entregas {manana} ===")

    log("Paso 1: Ejecutando VALIDAR_OC_FORMULARIOS.py ...")
    result = subprocess.run(
        [PYTHON, VALIDAR],
        capture_output=True, text=True, timeout=300,
        cwd=SCRIPT_DIR,
    )
    output = result.stdout or ""
    last_lines = "\n".join(output.strip().splitlines()[-10:])
    log(f"Salida (ultimas lineas):\n{last_lines}")
    if result.returncode != 0:
        log(f"ERROR: script retorno codigo {result.returncode}")
        log(result.stderr[:500] if result.stderr else "(sin stderr)")

    if not os.path.exists(HTML_OUTPUT):
        log("ERROR: No se genero VALIDACION_OC_FORMULARIOS.html")
        return False

    log("Paso 2: Tomando screenshot del dashboard ...")
    try:
        if not tomar_screenshot():
            log("ERROR: screenshot no se guardo")
            return False
        size_kb = os.path.getsize(SCREENSHOT) / 1024
        log(f"Screenshot OK ({size_kb:.0f} KB): {SCREENSHOT}")
    except Exception as e:
        log(f"ERROR en screenshot: {e}")
        return False

    log("Paso 3: Enviando por WhatsApp ...")
    texto = f"Validacion OC Formularios - Entregas {manana}"
    try:
        wa = subprocess.run(
            [PYTHON, ENVIAR_WA, DESTINO_CARLOS, texto, SCREENSHOT],
            capture_output=True, text=True, timeout=300,
            cwd=os.path.dirname(ENVIAR_WA),
        )
        wa_out = (wa.stdout or "").strip()
        if wa_out:
            log(f"WA salida: {wa_out[-300:]}")
        if wa.returncode == 0:
            log("ENVIO EXITOSO")
            return True
        else:
            log(f"ERROR envio WA (exit code {wa.returncode})")
            if wa.stderr:
                log(f"WA stderr: {wa.stderr[:300]}")
            return False
    except subprocess.TimeoutExpired:
        log("ERROR: envio WA excedio timeout de 300s")
        return False
    except Exception as e:
        log(f"ERROR inesperado en envio WA: {e}")
        return False


if __name__ == "__main__":
    ok = False
    try:
        ok = main()
    except Exception as e:
        log(f"ERROR FATAL: {e}")
    log(f"=== Fin (exito={ok}) ===\n")
    sys.exit(0 if ok else 1)
