"""
turso_db.py — drop-in replacement for sqlite3 usando a HTTP API do Turso.
Suporta os padrões usados no app.py:
  - connect() → TursoConnection
  - conn.execute(sql, params) → TursoCursor
  - conn.executescript(sql)
  - cursor.fetchall() / fetchone()
  - row["coluna"] e row[index]
  - with conn: ...
"""

import json
import os
import urllib.request
import urllib.error
from typing import Any, List, Optional


# ── Configuração ─────────────────────────────────────────────────────────────

def _base_url() -> str:
    url = os.getenv("TURSO_DATABASE_URL", "")
    if url.startswith("libsql://"):
        url = "https://" + url[9:]
    return url.rstrip("/")


def _token() -> str:
    return os.getenv("TURSO_AUTH_TOKEN", "")


# ── Tipos ─────────────────────────────────────────────────────────────────────

def _py_to_turso(v: Any) -> dict:
    """Converte valor Python para o formato de argumento da Turso HTTP API."""
    if v is None:
        return {"type": "null", "value": None}
    if isinstance(v, bool):
        return {"type": "integer", "value": str(int(v))}
    if isinstance(v, int):
        return {"type": "integer", "value": str(v)}
    if isinstance(v, float):
        return {"type": "float", "value": v}
    return {"type": "text", "value": str(v)}


def _turso_to_py(cell: dict) -> Any:
    """Converte célula da resposta Turso para valor Python."""
    t = cell.get("type", "null")
    v = cell.get("value")
    if t == "null" or v is None:
        return None
    if t == "integer":
        return int(v)
    if t == "float":
        return float(v)
    return v  # text / blob


class Row(dict):
    """Row que suporta acesso por chave (row["col"]) e por índice (row[0])."""

    def __init__(self, data: dict):
        super().__init__(data)
        self._keys_list = list(data.keys())

    def __getitem__(self, key):
        if isinstance(key, int):
            return super().__getitem__(self._keys_list[key])
        return super().__getitem__(key)

    def keys(self):
        return self._keys_list

    def get(self, key, default=None):
        try:
            return self[key]
        except (KeyError, IndexError):
            return default


# Alias para compatibilidade com `conn.row_factory = sqlite3.Row`
sqlite3_Row = Row


# ── HTTP Pipeline ─────────────────────────────────────────────────────────────

def _pipeline(statements: List[dict]) -> List[dict]:
    """
    Executa lista de statements via Turso HTTP pipeline.
    Cada statement: {"sql": "...", "args": [val, ...]}
    Retorna lista de resultados (um por statement).
    """
    requests = []
    for s in statements:
        requests.append({
            "type": "execute",
            "stmt": {
                "sql":  s["sql"],
                "args": [_py_to_turso(a) for a in s.get("args", [])],
            },
        })
    requests.append({"type": "close"})

    body = json.dumps({"requests": requests}).encode("utf-8")
    url  = f"{_base_url()}/v2/pipeline"
    req  = urllib.request.Request(
        url,
        data=body,
        headers={
            "Authorization": f"Bearer {_token()}",
            "Content-Type":  "application/json",
        },
    )

    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            data = json.loads(resp.read())
    except urllib.error.HTTPError as e:
        body_err = e.read().decode("utf-8", errors="replace")
        raise Exception(f"Turso HTTP {e.code}: {body_err[:300]}")

    results = []
    for r in data.get("results", []):
        if r.get("type") == "error":
            msg = (r.get("error") or {}).get("message", "Turso error")
            raise Exception(msg)
        if r.get("type") == "ok":
            inner = r.get("response", {})
            if inner.get("type") == "execute":
                rr   = inner.get("result", {})
                cols = [c["name"] for c in rr.get("cols", [])]
                rows = []
                for row_data in rr.get("rows", []):
                    d = {}
                    for i, col in enumerate(cols):
                        cell = row_data[i] if i < len(row_data) else {"type": "null", "value": None}
                        d[col] = _turso_to_py(cell)
                    rows.append(Row(d))
                results.append({
                    "rows":             rows,
                    "last_insert_rowid": rr.get("last_insert_rowid"),
                    "affected_rows":    rr.get("affected_row_count", 0),
                })
    return results


# ── Cursor ───────────────────────────────────────────────────────────────────

class TursoCursor:
    def __init__(self, rows: list = None, lastrowid: int = 0):
        self._rows    = rows or []
        self.lastrowid = lastrowid

    def fetchall(self) -> list:
        return self._rows

    def fetchone(self) -> Optional[Row]:
        return self._rows[0] if self._rows else None

    def __iter__(self):
        return iter(self._rows)

    def __getitem__(self, key):
        return self._rows[key]


# ── Connection ────────────────────────────────────────────────────────────────

class TursoConnection:
    """
    Conexão compatível com sqlite3.Connection.
    Aceita row_factory (ignorado — sempre usa Row).
    """

    def __init__(self):
        self.row_factory = Row  # aceita atribuição, mas é ignorado internamente

    def execute(self, sql: str, params=()) -> TursoCursor:
        results = _pipeline([{"sql": sql, "args": list(params)}])
        if results:
            r      = results[0]
            last   = r.get("last_insert_rowid")
            return TursoCursor(r["rows"], int(last) if last else 0)
        return TursoCursor([], 0)

    def executescript(self, script: str) -> None:
        """Divide o script em statements e executa em batch."""
        stmts = []
        for raw in script.split(";"):
            s = raw.strip()
            if s:
                stmts.append({"sql": s, "args": []})
        if stmts:
            _pipeline(stmts)

    def commit(self) -> None:
        pass  # Turso faz auto-commit

    def close(self) -> None:
        pass

    def __enter__(self):
        return self

    def __exit__(self, *args):
        pass


# ── API pública ───────────────────────────────────────────────────────────────

def connect(*args, **kwargs) -> TursoConnection:
    """Drop-in para sqlite3.connect() — ignora o caminho do arquivo."""
    return TursoConnection()


# Compatibilidade: sqlite3.Row
Row = Row
