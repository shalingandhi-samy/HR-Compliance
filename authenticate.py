"""One-time authentication helper.

Run this ONCE in a terminal to authenticate with Microsoft.
The token is saved to .token_cache.json and reused by the server automatically.

Usage:
    uv run python authenticate.py
"""
import sys
import msal
from pathlib import Path

TENANT_ID = "wmsc.wal-mart.com"
CLIENT_ID = "d3590ed6-52b3-4102-aeff-aad2292ab01c"  # MS Office public client
SCOPES = ["Files.Read", "Files.Read.All"]
TOKEN_CACHE_PATH = Path(__file__).parent / ".token_cache.json"


def main():
    token_cache = msal.SerializableTokenCache()
    if TOKEN_CACHE_PATH.exists():
        token_cache.deserialize(TOKEN_CACHE_PATH.read_text(encoding="utf-8"))

    app = msal.PublicClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        token_cache=token_cache,
    )

    # Check if already authenticated
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
        if result and "access_token" in result:
            print("\nAlready authenticated! Token is still valid.")
            print(f"Signed in as: {accounts[0].get('username', 'unknown')}")
            print("\nYou can start the server now:")
            print("  uv run uvicorn main:app --host 0.0.0.0 --port 8501")
            return

    # Device code flow - no browser redirect needed!
    print("\nStarting authentication...")
    print("=" * 60)
    flow = app.initiate_device_flow(scopes=SCOPES)

    if "user_code" not in flow:
        print(f"ERROR: Could not start auth flow: {flow}")
        sys.exit(1)

    # This prints: "Go to https://microsoft.com/devicelogin and enter code XXXXXXX"
    print(flow["message"])
    print("=" * 60)
    print("\nWaiting for you to sign in...")

    result = app.acquire_token_by_device_flow(flow)  # Blocks until user signs in

    if "access_token" not in result:
        print(f"\nERROR: Auth failed - {result.get('error_description', result)}")
        sys.exit(1)

    # Save token cache
    if token_cache.has_state_changed:
        TOKEN_CACHE_PATH.write_text(token_cache.serialize(), encoding="utf-8")

    print("\nAuthentication successful! Token saved.")
    print("\nYou can now start the server:")
    print("  uv run uvicorn main:app --host 0.0.0.0 --port 8501")


if __name__ == "__main__":
    main()
