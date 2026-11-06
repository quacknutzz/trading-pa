"""
setup_fyers.py -- Turnkey 1-Click FYERS API v3 Authenticator & Token Generator.

Guides the user through connecting their Fyers Demat account to Quacknutzz Terminal.
Saves verified access token into `fyers_tokens.json`.
"""
import json
import os
import sys
import time
import urllib.parse
import webbrowser
from http.server import HTTPServer, BaseHTTPRequestHandler

ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
TOKEN_FILE = os.path.join(ROOT_DIR, "fyers_tokens.json")
CONFIG_FILE = os.path.join(ROOT_DIR, "fyers_config.json")
PORT = 8765
REDIRECT_URI = f"http://127.0.0.1:{PORT}/api/fyers/callback"


class FyersAuthCallbackHandler(BaseHTTPRequestHandler):
    app_id = ""
    secret_key = ""
    token_saved = False
    error_msg = ""

    def log_message(self, format, *args):
        pass  # Quiet HTTP logging

    def do_GET(self):
        url_parts = urllib.parse.urlparse(self.path)
        if url_parts.path == "/api/fyers/callback":
            params = urllib.parse.parse_qs(url_parts.query)
            auth_code = params.get("auth_code", [""])[0]

            if not auth_code:
                self.send_response(400)
                self.send_header("Content-type", "text/html")
                self.end_headers()
                self.wfile.write(b"<h2>Missing auth_code from Fyers. Please retry.</h2>")
                return

            try:
                from fyers_apiv3 import fyersModel
                session = fyersModel.SessionModel(
                    client_id=self.app_id,
                    secret_key=self.secret_key,
                    redirect_uri=REDIRECT_URI,
                    response_type="code",
                    grant_type="authorization_code"
                )
                session.set_token(auth_code)
                resp = session.generate_token()

                if resp.get("s") == "ok" or "access_token" in resp:
                    access_token = resp.get("access_token")
                    with open(TOKEN_FILE, "w", encoding="utf-8") as f:
                        json.dump({"app_id": self.app_id, "access_token": access_token}, f, indent=2)
                    FyersAuthCallbackHandler.token_saved = True

                    # Verify profile
                    name = "Verified Account"
                    fy_id = ""
                    try:
                        client = fyersModel.FyersModel(client_id=self.app_id, token=access_token, log_path=None)
                        prof = client.get_profile()
                        if prof.get("s") == "ok":
                            name = prof.get("data", {}).get("name", "Verified Account")
                            fy_id = prof.get("data", {}).get("fy_id", "")
                    except Exception:
                        pass

                    self.send_response(200)
                    self.send_header("Content-type", "text/html")
                    self.end_headers()
                    success_html = f"""
                    <!DOCTYPE html>
                    <html>
                    <head>
                      <meta charset="utf-8">
                      <title>FYERS Connected</title>
                      <style>
                        body {{ background: #0a0d15; color: #f8fafc; font-family: sans-serif; display: flex; align-items: center; justify-content: center; height: 100vh; margin: 0; }}
                        .card {{ background: #111622; border: 1px solid #10b981; border-radius: 12px; padding: 40px; text-align: center; max-width: 460px; }}
                        .icon {{ font-size: 44px; color: #10b981; margin-bottom: 12px; }}
                        h2 {{ margin: 0 0 8px; color: #ffffff; }}
                        p {{ color: #94a3b8; font-size: 13px; }}
                        .btn {{ display: inline-block; background: #10b981; color: #07090e; font-weight: bold; padding: 10px 24px; border-radius: 6px; text-decoration: none; margin-top: 16px; }}
                      </style>
                    </head>
                    <body>
                      <div class="card">
                        <div class="icon">&#10003;</div>
                        <h2>FYERS Successfully Connected!</h2>
                        <p>Account: <b style="color:#22c55e;">{name} ({fy_id or self.app_id})</b></p>
                        <p style="color:#64748b; font-size: 11px;">You may now close this tab and return to your terminal.</p>
                      </div>
                    </body>
                    </html>
                    """
                    self.wfile.write(success_html.encode("utf-8"))
                else:
                    err = resp.get("message", "Token validation failed")
                    FyersAuthCallbackHandler.error_msg = err
                    self.send_response(400)
                    self.send_header("Content-type", "text/html")
                    self.end_headers()
                    self.wfile.write(f"<h2>Token Exchange Failed: {err}</h2>".encode("utf-8"))
            except Exception as e:
                FyersAuthCallbackHandler.error_msg = str(e)
                self.send_response(500)
                self.send_header("Content-type", "text/html")
                self.end_headers()
                self.wfile.write(f"<h2>Exception: {str(e)}</h2>".encode("utf-8"))
        else:
            self.send_response(404)
            self.end_headers()


def main():
    print("=" * 72)
    print("   QUACKNUTZZ TERMINAL // FYERS API v3 TURNKEY AUTHENTICATOR")
    print("=" * 72)
    print()
    print("This script connects your FYERS Demat account for live market execution.")
    print("SEBI requires generating a fresh broker session token every 24 hours.")
    print()
    print("--- 3 QUICK STEPS TO GET YOUR CREDENTIALS (2 MINUTES) ---")
    print("1. Log in to your Fyers Developer Dashboard: https://myapi.fyers.in/")
    print("2. Click 'Create App' and enter:")
    print("   - App Name:     Quacknutzz")
    print("   - Redirect URL: http://127.0.0.1:8765/api/fyers/callback")
    print("   - Permissions:  Check All (Orders, Positions, Funds, Quotes)")
    print("3. Copy your 'App ID' (e.g. 9ABC123XYZ-100) and 'Secret Key'")
    print("-" * 72)
    print()

    # Pre-load existing config if available
    saved_app_id = ""
    saved_secret = ""
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                cfg = json.load(f)
                saved_app_id = cfg.get("app_id", "")
                saved_secret = cfg.get("secret_key", "")
        except Exception:
            pass

    prompt_app = f"Enter Fyers App ID [{saved_app_id}]: " if saved_app_id else "Enter Fyers App ID (e.g. 9ABC123XYZ-100): "
    app_id = input(prompt_app).strip() or saved_app_id

    prompt_sec = f"Enter Fyers Secret Key [stored]: " if saved_secret else "Enter Fyers Secret Key: "
    secret_key = input(prompt_sec).strip() or saved_secret

    if not app_id or not secret_key:
        print("[ERROR] Both App ID and Secret Key are required to authenticate.")
        input("\nPress Enter to exit...")
        sys.exit(1)

    # Save config for future convenience
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump({"app_id": app_id, "secret_key": secret_key}, f, indent=2)

    try:
        from fyers_apiv3 import fyersModel
    except ImportError:
        print("[ERROR] fyers-apiv3 is not installed. Run: pip install fyers-apiv3 aws_lambda_powertools")
        input("\nPress Enter to exit...")
        sys.exit(1)

    # Generate Auth URL
    session = fyersModel.SessionModel(
        client_id=app_id,
        secret_key=secret_key,
        redirect_uri=REDIRECT_URI,
        response_type="code",
        grant_type="authorization_code"
    )
    auth_url = session.generate_authcode()

    print()
    print("[1/2] Opening Fyers Login Page in your default browser...")
    print(f"      Auth URL: {auth_url}")
    print()
    webbrowser.open(auth_url)

    # Check if port 8765 is already occupied by the server or free
    import socket
    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    port_in_use = sock.connect_ex(('127.0.0.1', PORT)) == 0
    sock.close()

    if port_in_use:
        print("[2/2] Cockpit backend is currently running on port 8765.")
        print("      Please complete your login and OTP in the browser.")
        print("      Waiting for authentication...")
        for _ in range(60):
            time.sleep(2)
            if os.path.exists(TOKEN_FILE):
                try:
                    with open(TOKEN_FILE, "r", encoding="utf-8") as f:
                        td = json.load(f)
                    if td.get("access_token"):
                        print()
                        print("=" * 72)
                        print("✓ FYERS CONNECTED SUCCESSFULLY!")
                        print(f"  Credentials verified and saved in: {TOKEN_FILE}")
                        print("  Your Cockpit terminal is now LIVE.")
                        print("=" * 72)
                        input("\nPress Enter to exit...")
                        return
                except Exception:
                    pass
        print("[WARN] Timed out waiting for browser callback. You can retry anytime.")
    else:
        print("[2/2] Starting temporary authentication listener on port 8765...")
        FyersAuthCallbackHandler.app_id = app_id
        FyersAuthCallbackHandler.secret_key = secret_key

        httpd = HTTPServer(('127.0.0.1', PORT), FyersAuthCallbackHandler)
        print("      Waiting for you to enter your OTP on the Fyers login page...")
        while not FyersAuthCallbackHandler.token_saved and not FyersAuthCallbackHandler.error_msg:
            httpd.handle_request()

        if FyersAuthCallbackHandler.token_saved:
            print()
            print("=" * 72)
            print("✓ FYERS CONNECTED SUCCESSFULLY!")
            print(f"  Credentials verified and saved in: {TOKEN_FILE}")
            print("  Launch Cockpit anytime with: LAUNCH_COCKPIT.bat")
            print("=" * 72)
        else:
            print(f"\n[ERROR] Authentication failed: {FyersAuthCallbackHandler.error_msg}")

    input("\nPress Enter to finish...")


if __name__ == "__main__":
    main()
