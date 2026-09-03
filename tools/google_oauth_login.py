"""Одноразовая авторизация Google OAuth (только для dev-машины).

Открывает браузер, сохраняет token.json рядом. Дальше токен вшивается в exe:
    python tools/gen_embedded_oauth.py
Конечным пользователям браузер и файлы НЕ нужны.

Нужен OAuth-клиент типа Desktop (GCP → APIs & Services → Credentials → Create Credentials
→ OAuth client ID → Desktop app → скачать JSON) — положите его как client_secret.json
рядом с этим скриптом или укажите путь аргументом.

ВАЖНО: на экране OAuth consent переведите приложение в Production,
иначе refresh_token протухнет через 7 дней (Testing).
"""

import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SCOPES = ["https://www.googleapis.com/auth/drive"]
TOKEN_OUT = ROOT / "token.json"


def main() -> int:
    try:
        from google_auth_oauthlib.flow import InstalledAppFlow
    except ImportError:
        print("ERROR: pip install google-auth-oauthlib", file=sys.stderr)
        return 1

    secret_arg = Path(sys.argv[1]) if len(sys.argv) > 1 else None
    cands = [p for p in (secret_arg, ROOT / "client_secret.json", Path.cwd() / "client_secret.json") if p]
    secret = next((p for p in cands if p and p.is_file()), None)
    if not secret:
        print("ERROR: нет client_secret.json (OAuth-клиент Desktop из GCP).", file=sys.stderr)
        print("Скачайте JSON и положите рядом или передайте путь аргументом.", file=sys.stderr)
        return 1

    flow = InstalledAppFlow.from_client_secrets_file(str(secret), SCOPES)
    creds = flow.run_local_server(port=0)
    TOKEN_OUT.write_text(creds.to_json(), encoding="utf-8")
    info = json.loads(creds.to_json())
    print(f"OK: {TOKEN_OUT} (refresh_token: {'есть' if info.get('refresh_token') else 'НЕТ — удалите доступ и пройдите заново с access_type=offline'})")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
