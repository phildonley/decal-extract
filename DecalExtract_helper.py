import os
import json
import getpass
import socket
import requests
import time

KEY_FILE = os.path.expanduser("~/.decal_api_key.json")
API_ENDPOINT = "https://hal4ecrr1k.execute-api.us-east-1.amazonaws.com/prod/get_current_drawing"
API_KEY = None

def get_valid_api_key() -> str:
    """
    Prompt the user once for X-API-KEY, store it in ~/.decal_api_key.json,
    and return it.  On subsequent runs, re-use the saved key.
    This also sets the module-global API_KEY so fetch_pdf_via_api() can see it.
    """
    global API_KEY

    # 1) Try to load existing key
    if os.path.exists(KEY_FILE):
        try:
            with open(KEY_FILE, "r") as f:
                data = json.load(f)
            key = data.get("x_api_key", "").strip()
            if key:
                API_KEY = key
                return API_KEY
        except Exception as e:
            print(f"[WARN] Failed to read API key file: {e}")

    # 2) Ask the user to paste in their API key
    api_key = getpass.getpass("Please paste your X-API-KEY for the signed-URL service: ").strip()

    # 3) Save it for next time
    try:
        with open(KEY_FILE, "w") as f:
            json.dump({"x_api_key": api_key}, f)
        os.chmod(KEY_FILE, 0o600)
        print("[OK] API key saved to disk.")
    except Exception as e:
        print(f"[WARN] Could not save API key to {KEY_FILE}: {e}")

    # 4) Store into module-global and return
    API_KEY = api_key
    return API_KEY

def fetch_pdf_via_api(part_number: str, pdf_dir: str) -> str | None:
    global API_KEY
    if API_KEY is None:
        raise RuntimeError("API_KEY has not been initialized!")

    # ensure output folder exists
    os.makedirs(pdf_dir, exist_ok=True)

    # DNS debug (optional)
    host = "hal4ecrr1k.execute-api.us-east-1.amazonaws.com"
    try:
        addr = socket.getaddrinfo(host, 443)
        print(f"[DEBUG] DNS lookup succeeded for {host} → {addr[0][4][0]}")
    except Exception as dns_err:
        print(f"[ERROR] DNS resolution failed for {host}: {dns_err}")
        return None

    headers = {
        "Content-Type": "application/json",
        "x-api-key":    API_KEY,
    }
    body = {"part_number": part_number}

    # 1) call the API
    try:
        resp = requests.post(API_ENDPOINT, headers=headers, json=body, timeout=30)
        resp.raise_for_status()
    except Exception as e:
        print(f"[ERROR] API call failed for '{part_number}': {e}")
        return None

    # 2) extract signed URL (JSON or raw text)
    url = None
    try:
        payload = resp.json()
        url = payload.get("url")
    except ValueError:
        txt = resp.text.strip()
        if txt.startswith("http"):
            url = txt

    if not url:
        print(f"[ERROR] No PDF URL in API response for '{part_number}'")
        return None

    # 3) download the PDF
    pdf_name = f"{part_number}_{int(time.time())}.pdf"
    pdf_path = os.path.join(pdf_dir, pdf_name)
    r = requests.get(url, stream=True, timeout=30)
    try:
        r.raise_for_status()
        with open(pdf_path, "wb") as f:
            for chunk in r.iter_content(chunk_size=8192):
                f.write(chunk)
    except Exception as download_err:
        print(f"[ERROR] Failed to download PDF: {download_err}")
        return None
    finally:
        r.close()

    print(f"[OK] Downloaded PDF → {pdf_path}")
    return pdf_path
