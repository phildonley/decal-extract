import os
import json
import getpass

KEY_FILE = os.path.expanduser("~/.decal_api_key.json")

def get_valid_api_key() -> str:
    """
    Prompt the user once for X-API-KEY, store it in ~/.decal_api_key.json,
    and return it.  On subsequent runs, re-use the saved key.
    """
    # 1) Try to load existing key
    if os.path.exists(KEY_FILE):
        try:
            with open(KEY_FILE, "r") as f:
                data = json.load(f)
                key = data.get("x_api_key")
                if key:
                    return key
        except Exception:
            pass

    # 2) Ask the user to paste in their API key
    print("Please paste your X-API-KEY for the signed-URL service:")
    key = getpass.getpass(prompt="X-API-KEY: ")

    # 3) Save it for next time
    try:
        with open(KEY_FILE, "w") as f:
            json.dump({"x_api_key": key}, f)
        os.chmod(KEY_FILE, 0o600)
    except Exception as e:
        print(f"Warning: could not save key to {KEY_FILE}: {e}")

    return key
