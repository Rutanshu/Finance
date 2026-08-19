"""
PaperTrail :: server

A local-only web server. Three deliberate constraints:

  1. It binds to 127.0.0.1 only. Not 0.0.0.0. Nothing on your network can reach it,
     let alone anything on the internet.
  2. It makes zero outbound requests. There is no HTTP client imported anywhere in
     this package. No CDNs, no fonts, no analytics, no update check.
  3. The only files it writes are papertrail.annotations.json (your notes) and any
     export you explicitly ask for.

Your workbooks are opened read-only and never modified.
"""

from __future__ import annotations

import argparse
import csv
import io
import json
import socket
import threading
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse, parse_qs

from . import graph as G
from .scanner import scan_folder

WEB_DIR = Path(__file__).parent / "web"

STATE: dict = {"root": None, "scan": None, "graph": None}
LOCK = threading.Lock()


def rebuild() -> dict:
    ann = G.load_annotations(STATE["root"])
    STATE["graph"] = G.build_graph(STATE["scan"], ann)
    return STATE["graph"]


class Handler(BaseHTTPRequestHandler):
    server_version = "PaperTrail"

    # keep the console readable
    def log_message(self, fmt, *args):
        if "/api/" in (args[0] if args else ""):
            return
        return

    # ---------------------------------------------------------------- helpers
    def _send(self, code: int, body: bytes, ctype: str = "application/json") -> None:
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        # belt and braces: forbid the page from talking to anything but itself
        self.send_header("Content-Security-Policy",
                         "default-src 'self' 'unsafe-inline' data:; connect-src 'self'")
        self.end_headers()
        self.wfile.write(body)

    def _json(self, obj, code: int = 200) -> None:
        self._send(code, json.dumps(obj).encode("utf-8"))

    def _body(self) -> dict:
        n = int(self.headers.get("Content-Length") or 0)
        if not n:
            return {}
        try:
            return json.loads(self.rfile.read(n).decode("utf-8"))
        except json.JSONDecodeError:
            return {}

    # ------------------------------------------------------------------- GET
    def do_GET(self):
        u = urlparse(self.path)
        q = parse_qs(u.query)

        if u.path in ("/", "/index.html"):
            return self._send(200, (WEB_DIR / "index.html").read_bytes(), "text/html; charset=utf-8")

        if u.path == "/api/state":
            return self._json({
                "root": STATE["root"],
                "loaded": STATE["graph"] is not None,
                "graph": STATE["graph"],
            })

        if u.path == "/api/trace":
            g = STATE["graph"]
            if not g:
                return self._json({"error": "nothing scanned yet"}, 400)
            node = q.get("node", [""])[0]
            up = sorted(G.walk(g["edges"], node, "up"))
            down = sorted(G.walk(g["edges"], node, "down"))
            chains = G.paths_to(g["edges"], node)
            return self._json({"node": node, "upstream": up, "downstream": down,
                               "chains": chains})

        if u.path == "/api/export.csv":
            g = STATE["graph"]
            if not g:
                return self._json({"error": "nothing scanned yet"}, 400)
            names = {n["id"]: n["name"] for n in g["nodes"]}
            buf = io.StringIO()
            w = csv.writer(buf)
            w.writerow(["from", "to", "mechanism", "risk", "discovered", "owner", "note", "tab_and_column_lineage"])
            for e in sorted(g["edges"], key=lambda e: G.RISK_ORDER[e["risk"]]):
                cols = "; ".join(
                    f"{c['src_sheet']}!{c['src_col']} → {c['dest_sheet']}!{c['dest_col']}"
                    + (f" ({c['dest_header']})" if c["dest_header"] else "") + f" x{c['cells']}"
                    for c in e.get("columns", [])
                )
                w.writerow([names.get(e["from"], e["from"]), names.get(e["to"], e["to"]),
                            G.MECHANISMS[e["mechanism"]]["label"], e["risk"],
                            "auto" if e["discovered"] else "declared", e["owner"], e["note"], cols])
            body = buf.getvalue().encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/csv")
            self.send_header("Content-Disposition", 'attachment; filename="papertrail-handoffs.csv"')
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            return self.wfile.write(body)

        return self._json({"error": "not found"}, 404)

    # ------------------------------------------------------------------ POST
    def do_POST(self):
        u = urlparse(self.path)
        data = self._body()

        if u.path == "/api/scan":
            root = (data.get("root") or "").strip()
            if not root:
                return self._json({"error": "no folder given"}, 400)
            try:
                with LOCK:
                    STATE["root"] = str(Path(root).expanduser().resolve())
                    STATE["scan"] = scan_folder(STATE["root"])
                    g = rebuild()
                return self._json({"root": STATE["root"], "graph": g})
            except Exception as exc:
                return self._json({"error": f"{type(exc).__name__}: {exc}"}, 400)

        if u.path == "/api/annotate":
            if not STATE["root"]:
                return self._json({"error": "nothing scanned yet"}, 400)
            with LOCK:
                ann = G.load_annotations(STATE["root"])
                kind = data.get("kind")
                if kind == "node":
                    ann["nodes"].setdefault(data["id"], {}).update(data.get("value", {}))
                elif kind == "edge":
                    ann["edges"].setdefault(data["id"], {}).update(data.get("value", {}))
                elif kind == "manual_edge":
                    v = data.get("value", {})
                    ann["manual_edges"] = [m for m in ann["manual_edges"]
                                           if not (m["from"] == v.get("from") and m["to"] == v.get("to"))]
                    if not v.get("_delete"):
                        ann["manual_edges"].append(v)
                else:
                    return self._json({"error": "unknown annotation kind"}, 400)
                G.save_annotations(STATE["root"], ann)
                g = rebuild()
            return self._json({"graph": g})

        if u.path == "/api/rescan":
            if not STATE["root"]:
                return self._json({"error": "nothing scanned yet"}, 400)
            with LOCK:
                STATE["scan"] = scan_folder(STATE["root"])
                g = rebuild()
            return self._json({"graph": g})

        return self._json({"error": "not found"}, 404)


def free_port(preferred: int) -> int:
    for p in [preferred, 0]:
        try:
            s = socket.socket()
            s.bind(("127.0.0.1", p))
            port = s.getsockname()[1]
            s.close()
            return port
        except OSError:
            continue
    return preferred


def main() -> None:
    ap = argparse.ArgumentParser(
        prog="papertrail",
        description="Map how finance data flows through your Excel workbooks. Runs entirely on this machine.")
    ap.add_argument("folder", nargs="?", help="folder of workbooks to scan on startup")
    ap.add_argument("--port", type=int, default=8737)
    ap.add_argument("--no-browser", action="store_true")
    args = ap.parse_args()

    if args.folder:
        STATE["root"] = str(Path(args.folder).expanduser().resolve())
        print(f"  scanning {STATE['root']} …")
        STATE["scan"] = scan_folder(STATE["root"])
        g = rebuild()
        print(f"  found {g['stats']['workbooks']} workbooks, "
              f"{g['stats']['handoffs']} handoffs, "
              f"{g['stats']['fragile']} fragile")

    port = free_port(args.port)
    url = f"http://127.0.0.1:{port}"
    httpd = ThreadingHTTPServer(("127.0.0.1", port), Handler)

    print("\n  PaperTrail")
    print("  " + "-" * 46)
    print(f"  open   {url}")
    print("  bound  127.0.0.1 only — not reachable from your network")
    print("  data   never leaves this machine")
    print("  stop   Ctrl+C\n")

    if not args.no_browser:
        threading.Timer(0.6, lambda: webbrowser.open(url)).start()
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\n  stopped.\n")


if __name__ == "__main__":
    main()
