import os
import sys
import urllib.parse
from pathlib import Path
import requests


def _env(name: str, required: bool = True, default: str | None = None) -> str:
    v = os.getenv(name, default)
    if required and not v:
        print(f"Missing environment variable: {name}", file=sys.stderr)
        sys.exit(2)
    return v or ""


def token(tenant: str, client_id: str, client_secret: str) -> str:
    resp = requests.post(
        f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token",
        data={
            "client_id": client_id,
            "client_secret": client_secret,
            "grant_type": "client_credentials",
            "scope": "https://graph.microsoft.com/.default",
        },
        timeout=30,
    )
    if resp.status_code != 200:
        print(resp.status_code, resp.text, file=sys.stderr)
        sys.exit(3)
    return resp.json()["access_token"]


def fetch_to(token: str, upn: str, drive_path: str, out: Path) -> None:
    if not drive_path.startswith('/'):
        drive_path = '/' + drive_path
    url = (
        "https://graph.microsoft.com/v1.0/users/"
        + urllib.parse.quote(upn)
        + "/drive/root:"
        + urllib.parse.quote(drive_path)
        + ":/content"
    )
    with requests.get(url, headers={"Authorization": f"Bearer {token}"}, stream=True, timeout=120) as r:
        if r.status_code != 200:
            print(r.status_code, r.text[:500], file=sys.stderr)
            sys.exit(4)
        out.parent.mkdir(parents=True, exist_ok=True)
        with open(out, 'wb') as f:
            for chunk in r.iter_content(1 << 20):
                if chunk:
                    f.write(chunk)


def main() -> None:
    tenant = _env("TENANT_ID")
    client_id = _env("CLIENT_ID")
    client_secret = _env("CLIENT_SECRET")
    upn = _env("GRAPH_USER_UPN")
    drive_path = _env("DRIVE_PATH")
    out = Path(os.getenv("OUTPUT_PATH", "data/source.xlsx"))

    tok = token(tenant, client_id, client_secret)
    fetch_to(tok, upn, drive_path, out)
    print(f"Downloaded to {out}")


if __name__ == "__main__":
    main()
