"""
Test POST Supabase CRM — eseguire standalone per verificare connessione e schema.
Uso: python test_supabase_post.py

Inserisce un cliente e un documento di test, poi li cancella.
Richiede SUPABASE_URL e SUPABASE_KEY in local.env o variabili d'ambiente.
"""
import json
import os
import sys
import urllib.request
import urllib.error
from datetime import datetime

# ── carica .env / local.env ──────────────────────────────────────────────────
def _load_dotenv(path: str) -> None:
    if not os.path.isfile(path):
        return
    with open(path, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            k, v = line.split("=", 1)
            k = k.strip()
            v = v.strip().strip('"').strip("'")
            if k:
                os.environ.setdefault(k, v)

_here = os.path.dirname(os.path.abspath(__file__))
_load_dotenv(os.path.join(_here, ".env"))
_load_dotenv(os.path.join(_here, "local.env"))

SUPABASE_URL = os.environ.get("SUPABASE_URL", "").strip().rstrip("/")
SUPABASE_KEY = os.environ.get("SUPABASE_KEY", "").strip()

# ────────────────────────────────────────────────────────────────────────────

def _req(method: str, path: str, payload: dict | None = None, prefer: str = "") -> tuple[int, dict | list]:
    url = f"{SUPABASE_URL}/rest/v1/{path}"
    data = json.dumps(payload).encode() if payload else None
    headers = {
        "apikey": SUPABASE_KEY,
        "Authorization": f"Bearer {SUPABASE_KEY}",
        "Content-Type": "application/json",
    }
    if prefer:
        headers["Prefer"] = prefer
    req = urllib.request.Request(url, data=data, headers=headers, method=method)
    try:
        with urllib.request.urlopen(req, timeout=15) as resp:
            body = resp.read().decode()
            return resp.status, json.loads(body) if body.strip() else {}
    except urllib.error.HTTPError as e:
        body = e.read().decode()
        return e.code, {"error": body}


def ok(msg: str) -> None:
    print(f"  ✅ {msg}")

def fail(msg: str) -> None:
    print(f"  ❌ {msg}")
    sys.exit(1)


def main() -> None:
    print("\nTest connessione Supabase CRM")
    print("=" * 40)

    # 0. Verifica config
    if not SUPABASE_URL or not SUPABASE_KEY:
        fail("SUPABASE_URL o SUPABASE_KEY mancanti in local.env")

    print(f"  URL: {SUPABASE_URL}")
    print(f"  KEY: {SUPABASE_KEY[:20]}…")
    print()

    TEST_PIVA = "TEST00000001"
    doc_id_creato = None

    # 1. Upsert cliente di test
    print("[1] Upsert clienti…")
    status, body = _req("POST", "clienti", {
        "piva":            TEST_PIVA,
        "ragione_sociale": "CLIENTE TEST (cancellare)",
        "cap":             "00000",
        "comune":          "Test City",
        "sigla_provincia": "TS",
    }, prefer="resolution=merge-duplicates,return=representation")
    if status in (200, 201):
        ok(f"clienti upsert → HTTP {status}")
    else:
        fail(f"clienti upsert → HTTP {status}: {body}")

    # 2. Insert documento di test
    print("[2] Insert documenti…")
    status, body = _req("POST", "documenti", {
        "numero_documento": "TEST-001",
        "tipo":             "PC",
        "data_documento":   datetime.now().strftime("%Y-%m-%d"),
        "piva_cliente":     TEST_PIVA,
        "percorso_pdf_locale": "C:\\test\\test.pdf",
    }, prefer="return=representation")
    if status in (200, 201):
        result = body if isinstance(body, list) else [body]
        doc_id_creato = result[0].get("id") if result else None
        ok(f"documenti insert → HTTP {status}  id={doc_id_creato}")
    else:
        fail(f"documenti insert → HTTP {status}: {body}")

    # 3. Cleanup — elimina documento e cliente di test
    print("[3] Cleanup dati di test…")
    if doc_id_creato:
        status, _ = _req("DELETE", f"documenti?id=eq.{doc_id_creato}")
        ok(f"documento cancellato → HTTP {status}") if status in (200, 204) else fail(f"delete doc → {status}")

    status, _ = _req("DELETE", f"clienti?piva=eq.{TEST_PIVA}")
    ok(f"cliente cancellato → HTTP {status}") if status in (200, 204) else fail(f"delete cliente → {status}")

    print()
    print("=" * 40)
    print("  Tutto OK — Supabase CRM raggiungibile e schema corretto.")
    print()


if __name__ == "__main__":
    main()
