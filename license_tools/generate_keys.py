# license_tools/generate_keys.py
# Generates valid LF-XXXX-XXXX-XXXX-CC keys for this MVP.
# IMPORTANT: Must match LICENSE_SECRET in app.py.

import hmac
import base64
import random

LICENSE_SECRET = b"LeadForgeFI_MVP_secret_change_this_before_big_launch"

def sig2(payload: str) -> str:
    mac = hmac.new(LICENSE_SECRET, payload.encode("utf-8"), digestmod="sha256").digest()
    return base64.b32encode(mac)[:2].decode("ascii")

def block(n=4):
    alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
    return "".join(random.choice(alphabet) for _ in range(n))

def make_key():
    payload = f"LF-{block()}-{block()}-{block()}"
    cc = sig2(payload)
    return f"{payload}-{cc}"

if __name__ == "__main__":
    for _ in range(20):
        print(make_key())
