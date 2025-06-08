import os
import json
import getpass
import socket
import requests

KEY_FILE = os.path.expanduser("~/.decal_api_key.json")

def get_valid_api_key():
    """
    Returns a cached API key or prompts the user to paste it the first time.
    Saves to ~/.decal_api_key.json for reuse.
    """
    if os.path.exists(KEY_FILE):
        try:
            with open(KEY_FILE, "r") as f:
                return json.load(f)["api_key"]
        except Exception as e:
            print(f"[WARN] Failed to read API key file: {e}")

    api_key = getpass.getpass("Please paste your X-API-KEY for the signed-URL service: ")
    with open(KEY_FILE, "w") as f:
        json.dump({"api_key": api_key.strip()}, f)
    print("[OK] API key saved to disk.")
    return api_key.strip()

def fetch_pdf_via_api(part_number: str, pdf_dir: str) -> str | None:
    """
    Fetches a PDF for a part number using the API and saves it locally.
    Returns the file path if successful, otherwise None.
    """
    from DecalExtract import API_ENDPOINT, API_KEY  # Assumes these are defined in the main script

    # (a) Ensure API key is initialized
    if API_KEY is None:
        raise RuntimeError("API_KEY has not been initialized!")

    # (b) Extract hostname for DNS check
    host = API_ENDPOINT.split("/")[2]

    # (c) DNS resolution check
    try:
        addr = socket.getaddrinfo(host, 443)
        print(f"[DEBUG] DNS lookup succeeded for {host} → {addr[0][4][0]}")
    except Exception as dns_err:
        print(f"[ERROR] DNS resolution failed for {host}: {dns_err}")
        print("         Are you connected to the work network or VPN?")
        return None

    # (d) Build API request
    headers = {
        "Content-Type": "application/json",
        "x-api-key": API_KEY
    }
    body = {
        "part_number": part_number
    }

    try:
        response = requests.post(API_ENDPOINT, headers=headers, json=body, timeout=30)
    except Exception as e:
        print(f"[ERROR] API call failed for '{part_number}': {e}")
        return None

    if response.status_code != 200:
        print(f"[ERROR] API returned HTTP {response.status_code} for '{part_number}'")
        print("        Raw response body:\n" + response.text[:200] + ("..." if len(response.text) > 200 else ""))
        return None

    # (e) Parse JSON and extract signed URL
    try:
        json_data = response.json()
        url = json_data.get("url")
        if not url:
            print(f"[ERROR] No 'url' field in API response for '{part_number}'")
            return None
    except Exception as parse_err:
        print(f"[ERROR] Failed to parse JSON from API response: {parse_err}")
        return None

    # (f) Download the actual PDF
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
