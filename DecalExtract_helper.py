import os
import json
import getpass
import socket
import requests

KEY_FILE = os.path.expanduser("~/.decal_api_key.json")
API_ENDPOINT = "https://hal4ecrr1k.execute-api.us-east-1.amazonaws.com/prod/get_current_drawing"
API_KEY = None

def get_valid_api_key():
    if os.path.exists(KEY_FILE):
        try:
            with open(KEY_FILE, "r") as f:
                return json.load(f)["x_api_key"]
        except Exception as e:
            print(f"[WARN] Failed to read API key file: {e}")

    api_key = getpass.getpass("Please paste your X-API-KEY for the signed-URL service: ")
    with open(KEY_FILE, "w") as f:
        json.dump({"x_api_key": api_key.strip()}, f)  # <-- use "x_api_key"
    print("[OK] API key saved to disk.")
    return api_key.strip()

def fetch_pdf_via_api(part_number: str, pdf_dir: str) -> str | None:
    global API_KEY

    if API_KEY is None:
        raise RuntimeError("API_KEY has not been initialized!")

    host = "hal4ecrr1k.execute-api.us-east-1.amazonaws.com"
    try:
        addr = socket.getaddrinfo(host, 443)
        print(f"[DEBUG] DNS lookup succeeded for {host} → {addr[0][4][0]}")
    except Exception as dns_err:
        print(f"[ERROR] DNS resolution failed for {host}: {dns_err}")
        return None

    headers = {
        "Content-Type": "application/json",
        "x-api-key": API_KEY
    }
    body = {"part_number": part_number}

    try:
        response = requests.post(API_ENDPOINT, headers=headers, json=body, timeout=30)
        response.raise_for_status()
    except Exception as e:
        print(f"[ERROR] API call failed for '{part_number}': {e}")
        return None

    try:
        url = response.json().get("url")
        if not url:
            print(f"[ERROR] No 'url' field in API response for '{part_number}'")
            return None
    except Exception as parse_err:
        print(f"[ERROR] Failed to parse JSON from API response: {parse_err}")
        return None

    try:
        pdf_name = f"{part_number}.pdf"
        pdf_path = os.path.join(pdf_dir, pdf_name)
        with requests.get(url, stream=True, timeout=30) as r:
            r.raise_for_status()
            with open(pdf_path, "wb") as f:
                for chunk in r.iter_content(chunk_size=8192):
                    f.write(chunk)
        print(f"[OK] Downloaded PDF → {pdf_path}")
        return pdf_path
    except Exception as download_err:
        print(f"[ERROR] Failed to download PDF from signed URL: {download_err}")
        return None
