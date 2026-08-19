"""
PaperTrail :: graph
Turns a scan result into a lineage graph: nodes (workbooks) + typed edges (handoffs).

The core idea: risk is not a judgement, it is read off the MECHANISM of the handoff.
  automated query / refresh -> low      (auditable, repeatable)
  formula link across files -> medium   (breaks silently when a file moves)
  human copy-paste          -> high     (no audit trail at all)
  manual keying             -> high

Formula links are discovered automatically. Human handoffs cannot be detected by any
scanner on earth, so they are ANNOTATED by the user in the UI and stored locally.
"""

from __future__ import annotations

import hashlib
import json
from collections import defaultdict
from pathlib import Path

ANNOTATION_FILE = "papertrail.annotations.json"

MECHANISMS = {
    "query":       {"label": "Automated refresh",       "risk": "low"},
    "formula":     {"label": "Cross-file formula",       "risk": "med"},
    "paste":       {"label": "Copy-paste by hand",       "risk": "high"},
    "manual":      {"label": "Typed in by hand",         "risk": "high"},
    "email":       {"label": "Emailed between people",   "risk": "high"},
    # external / outside-the-folder sources — see graph.EXTERNAL_KINDS below for how these are detected
    "web_query":   {"label": "Web query / API",          "risk": "med"},
    "database":    {"label": "Database connection",      "risk": "med"},
    "sharepoint":  {"label": "SharePoint / cloud file",  "risk": "low"},
    "exchange":    {"label": "Exchange / email data",    "risk": "med"},
    "local_folder":{"label": "Local folder (Power Query)", "risk": "low"},
    "web_workbook":{"label": "Linked workbook via URL",  "risk": "high"},
    "web_formula": {"label": "Live web formula",         "risk": "high"},
}

RISK_ORDER = {"high": 0, "med": 1, "low": 2}

# how each external-source "kind" (from scanner._external_sources / web_links) is drawn: icon + friendly label
EXTERNAL_KINDS = {
    "web_query":    {"icon": "🌐", "label": "Website / web API"},
    "database":     {"icon": "🗄️", "label": "Database"},
    "sharepoint":   {"icon": "☁️", "label": "SharePoint / cloud"},
    "exchange":     {"icon": "📧", "label": "Exchange / email"},
    "local_folder": {"icon": "📁", "label": "Local folder"},
    "web_workbook": {"icon": "🔗", "label": "Linked workbook (URL)"},
    "web_formula":  {"icon": "📡", "label": "Live web formula"},
}


# --------------------------------------------------------------------------- #
# annotations: the human layer, stored beside the workbooks, never transmitted
# --------------------------------------------------------------------------- #
def annotation_path(root: str | Path) -> Path:
    return Path(root) / ANNOTATION_FILE


def load_annotations(root: str | Path) -> dict:
    p = annotation_path(root)
    if not p.exists():
        return {"edges": {}, "nodes": {}, "manual_edges": []}
    try:
        data = json.loads(p.read_text("utf-8"))
    except (json.JSONDecodeError, OSError):
        return {"edges": {}, "nodes": {}, "manual_edges": []}
    data.setdefault("edges", {})
    data.setdefault("nodes", {})
    data.setdefault("manual_edges", [])
    return data


def save_annotations(root: str | Path, data: dict) -> None:
    annotation_path(root).write_text(json.dumps(data, indent=2), "utf-8")


# --------------------------------------------------------------------------- #
# graph construction
# --------------------------------------------------------------------------- #
MAX_COLUMN_ROWS = 60


def _group_columns(links: list[dict], key_fn, seed_fn) -> list[dict]:
    """
    Shared roll-up: cell-level links -> one row per (source, destination) column pair, so the UI
    can show "Data!B feeds Summary!B (12 cells)" instead of a wall of individual cells. Used for
    both cross-file formula links and web-formula links — they differ only in grouping key / seed.
    """
    groups: dict[tuple, dict] = {}
    for link in links:
        g = groups.setdefault(key_fn(link), seed_fn(link))
        g["cells"] += 1
        if len(g["sample_cells"]) < 6:
            g["sample_cells"].append(link["dest_cell"])
        g["dest_header"] = g["dest_header"] or link["dest_header"]
    out = sorted(groups.values(), key=lambda g: (-g["cells"], g["dest_sheet"], g["dest_col"]))
    return out[:MAX_COLUMN_ROWS]


def _column_lineage(dest_book: dict, src_ref_name: str) -> list[dict]:
    links = [l for l in dest_book.get("links", []) if l["src_book"].lower() == src_ref_name.lower()]
    return _group_columns(
        links,
        key_fn=lambda l: (l["src_sheet"], l["src_col"], l["dest_sheet"], l["dest_col"]),
        seed_fn=lambda l: {"src_sheet": l["src_sheet"] or "?", "src_col": l["src_col"] or "?",
                            "dest_sheet": l["dest_sheet"], "dest_col": l["dest_col"],
                            "dest_header": l["dest_header"], "cells": 0, "sample_cells": [],
                            "sample_formula": l["formula"]})


def _web_column_lineage(book: dict) -> list[dict]:
    """Same idea as _column_lineage, but the 'source' is the internet, not a sheet/column."""
    return _group_columns(
        book.get("web_links", []),
        key_fn=lambda l: (l["dest_sheet"], l["dest_col"]),
        seed_fn=lambda l: {"src_sheet": "web", "src_col": "", "dest_sheet": l["dest_sheet"],
                            "dest_col": l["dest_col"], "dest_header": l["dest_header"],
                            "cells": 0, "sample_cells": [], "sample_formula": l["formula"]})


def _ext_id(kind: str, detail: str) -> str:
    return f"ext_{kind}_{hashlib.sha1(f'{kind}|{detail}'.encode()).hexdigest()[:10]}"


def _sheet_sentinel_columns(sheet: str) -> list[dict]:
    """
    A connections.xml-based source (database / legacy web query) has no cell-level formula to point at,
    but scanner._connection_sheet_map may still know which sheet it loads into. Represent that as one
    "whole sheet" row so the sheet-level drill-down can show it — same shape as a real column row, just
    without a specific column, since a query table loads a whole range, not one formula's worth of cells.
    """
    if not sheet:
        return []
    return [{"src_sheet": "", "src_col": "", "dest_sheet": sheet, "dest_col": "(query result)",
             "dest_header": "", "cells": 1, "sample_cells": [], "sample_formula": ""}]


def _stage_of(book: dict, indeg: int, outdeg: int) -> int:
    """
    Infer a rough pipeline stage so the map has left-to-right meaning.
      0 = source (nothing feeds it)   2 = staging   4 = report (feeds nothing)
    """
    if indeg == 0 and outdeg > 0:
        return 0
    if outdeg == 0 and indeg > 0:
        return 4
    if indeg == 0 and outdeg == 0:
        return 2
    return 2 if indeg <= outdeg else 3


def build_graph(scan: dict, annotations: dict | None = None) -> dict:
    annotations = annotations or {"edges": {}, "nodes": {}, "manual_edges": []}
    books = scan["books"]
    by_name = {b["name"].lower(): b for b in books}

    edges: list[dict] = []
    seen: set[tuple[str, str]] = set()

    # a human's description of a hop always beats the scanner's guess: if someone
    # declares that a link is really a copy-paste, that is the truth of record.
    declared = {(m["from"], m["to"]): m for m in annotations["manual_edges"]}

    # 1. discovered edges: file A contains a formula pointing at file B  => B feeds A
    for b in books:
        for ref in b["external_refs"]:
            src = by_name.get(ref.lower())
            if not src or src["id"] == b["id"]:
                continue
            key = (src["id"], b["id"])
            if key in seen:
                continue
            seen.add(key)
            ann = annotations["edges"].get(f"{src['id']}->{b['id']}", {})
            override = declared.get(key, {})
            mech = override.get("mechanism") or ann.get("mechanism", "formula")
            edges.append({
                "id": f"{src['id']}->{b['id']}",
                "from": src["id"],
                "to": b["id"],
                "mechanism": mech,
                "risk": MECHANISMS.get(mech, MECHANISMS["formula"])["risk"],
                "discovered": not override,
                "note": override.get("note") or ann.get("note", ""),
                "owner": override.get("owner") or ann.get("owner", ""),
                "columns": _column_lineage(b, ref),
            })

    # 2. annotated edges: handoffs a human declared (paste, email, re-keying)
    for me in annotations["manual_edges"]:
        key = (me["from"], me["to"])
        if key in seen:
            continue
        seen.add(key)
        mech = me.get("mechanism", "paste")
        edges.append({
            "id": f"{me['from']}->{me['to']}",
            "from": me["from"],
            "to": me["to"],
            "mechanism": mech,
            "risk": MECHANISMS.get(mech, MECHANISMS["paste"])["risk"],
            "discovered": False,
            "note": me.get("note", ""),
            "owner": me.get("owner", ""),
            "columns": [],
        })

    # stage (pipeline lane) is decided from book<->book edges only, so a workbook that also happens
    # to pull from a website doesn't get visually shoved rightward — that would misrepresent its
    # place in the file-to-file pipeline. External sources are drawn in their own lane instead.
    book_indeg: dict[str, int] = defaultdict(int)
    book_outdeg: dict[str, int] = defaultdict(int)
    for e in edges:
        book_outdeg[e["from"]] += 1
        book_indeg[e["to"]] += 1

    # 3. external sources: connections.xml / Power Query sources, plus live WEBSERVICE-style formulas.
    # Each distinct (kind, detail) becomes ONE node, even if several workbooks pull from it.
    def link_external(eid: str, node: dict, book_id: str, mechanism: str, note: str, columns: list[dict]) -> None:
        ext_nodes.setdefault(eid, node)
        key = (eid, book_id)
        if key in seen:
            return
        seen.add(key)
        edges.append({"id": f"{eid}->{book_id}", "from": eid, "to": book_id,
                      "mechanism": mechanism, "risk": MECHANISMS.get(mechanism, {"risk": "med"})["risk"],
                      "discovered": True, "note": note, "owner": "", "columns": columns})

    ext_nodes: dict[str, dict] = {}
    for b in books:
        for src in b.get("external_sources", []):
            kind = src["kind"]
            eid = _ext_id(kind, src["detail"])
            info = EXTERNAL_KINDS.get(kind, {"icon": "🔌", "label": kind})
            node = {"id": eid, "kind": "external", "ext_kind": kind,
                    "name": f"{info['icon']} {src['detail']}"[:90],
                    "label": info["label"], "detail": src["detail"], "via": src.get("via", "")}
            link_external(eid, node, b["id"], kind, src.get("via", ""), _sheet_sentinel_columns(src.get("sheet", "")))

        if b.get("web_links"):
            eid = _ext_id("web_formula", b["id"])
            urls = sorted({wl["url"] for wl in b["web_links"] if wl.get("url")})
            detail = urls[0] if len(urls) == 1 else (f"{len(urls)} endpoints" if urls else "live web formula")
            node = {"id": eid, "kind": "external", "ext_kind": "web_formula",
                    "name": f"📡 {detail}"[:90], "label": EXTERNAL_KINDS["web_formula"]["label"],
                    "detail": detail, "via": "WEBSERVICE / FILTERXML / RTD formula"}
            link_external(eid, node, b["id"], "web_formula",
                          "cell formula reaches the internet at calculation time", _web_column_lineage(b))

    indeg: dict[str, int] = defaultdict(int)
    outdeg: dict[str, int] = defaultdict(int)
    for e in edges:
        outdeg[e["from"]] += 1
        indeg[e["to"]] += 1

    nodes = []
    for b in books:
        ann = annotations["nodes"].get(b["id"], {})
        formulas = sum(s["formula_cells"] for s in b["sheets"])
        nodes.append({
            "id": b["id"],
            "kind": "workbook",
            "name": b["name"],
            "rel_path": b["rel_path"],
            "path": b["path"],
            "size_kb": b["size_kb"],
            "modified": b["modified"],
            "sheets": b["sheets"],
            "sheet_count": len(b["sheets"]),
            "formula_count": formulas,
            "volatile": sum(s["volatile_cells"] for s in b["sheets"]),
            "lookups": sum(s["lookup_cells"] for s in b["sheets"]),
            "hardcoded": sum(s["hardcoded_in_formula"] for s in b["sheets"]),
            "web_formula_cells": sum(s.get("web_formula_cells", 0) for s in b["sheets"]),
            "external_sources": b.get("external_sources", []),
            "unresolved_refs": b["unresolved_refs"],
            "error": b["error"],
            "in_degree": indeg[b["id"]],
            "out_degree": outdeg[b["id"]],
            "stage": _stage_of(b, book_indeg[b["id"]], book_outdeg[b["id"]]) + 1,
            "owner": ann.get("owner", ""),
            "note": ann.get("note", ""),
            "critical": ann.get("critical", False),
        })

    for en in ext_nodes.values():
        nodes.append({**en, "stage": 0, "in_degree": indeg[en["id"]], "out_degree": outdeg[en["id"]]})

    return {"root": scan["root"], "nodes": nodes, "edges": edges,
            "findings": findings(nodes, edges), "stats": stats(nodes, edges)}


# --------------------------------------------------------------------------- #
# traversal
# --------------------------------------------------------------------------- #
def walk(edges: list[dict], start: str, direction: str) -> set[str]:
    """Depth-first walk. direction 'up' = ancestors, 'down' = descendants."""
    seen: set[str] = set()
    stack = [start]
    while stack:
        cur = stack.pop()
        for e in edges:
            if direction == "up" and e["to"] == cur and e["from"] not in seen:
                seen.add(e["from"])
                stack.append(e["from"])
            elif direction == "down" and e["from"] == cur and e["to"] not in seen:
                seen.add(e["to"])
                stack.append(e["to"])
    return seen


def paths_to(edges: list[dict], start: str, max_paths: int = 40) -> list[list[dict]]:
    """All upstream chains feeding `start`, as ordered lists of edges."""
    out: list[list[dict]] = []

    def rec(node: str, chain: list[dict], visited: set[str]) -> None:
        if len(out) >= max_paths:
            return
        parents = [e for e in edges if e["to"] == node and e["from"] not in visited]
        if not parents:
            if chain:
                out.append(list(reversed(chain)))
            return
        for e in parents:
            rec(e["from"], chain + [e], visited | {e["from"]})

    rec(start, [], {start})
    return out


# --------------------------------------------------------------------------- #
# analysis
# --------------------------------------------------------------------------- #
def findings(nodes: list[dict], edges: list[dict]) -> list[dict]:
    """Concrete, actionable problems. Ordered worst first."""
    out: list[dict] = []
    by_id = {n["id"]: n for n in nodes}

    for e in edges:
        src_node = by_id[e["from"]]
        if src_node["kind"] == "external":
            sev = {"high": "high", "med": "med", "low": "low"}[e["risk"]]
            out.append({
                "severity": sev,
                "kind": "External data source",
                "where": by_id[e["to"]]["name"],
                "what": f"{MECHANISMS[e['mechanism']]['label']} — pulls from {src_node['detail']} "
                        f"({src_node.get('via','')}). Nothing in the folder controls this data.",
                "fix": "Confirm who owns this connection/credential and what happens if the source changes "
                       "or goes away. Document it like any other pipeline dependency.",
            })
        elif e["risk"] == "high":
            out.append({
                "severity": "high",
                "kind": "Fragile handoff",
                "where": f"{by_id[e['from']]['name']} → {by_id[e['to']]['name']}",
                "what": f"{MECHANISMS[e['mechanism']]['label']} — no audit trail between these files.",
                "fix": "Add a control total (row count or gross sum) that must match on both sides, "
                       "and turn it red on mismatch.",
            })

    for n in nodes:
        if n["kind"] != "workbook":
            continue
        if n["error"]:
            out.append({"severity": "high", "kind": "Unreadable file", "where": n["rel_path"],
                        "what": n["error"],
                        "fix": "File is corrupt, password-protected, or not a real workbook. "
                               "It cannot be audited in its current form."})
        for ref in n["unresolved_refs"]:
            out.append({"severity": "high", "kind": "Broken link", "where": n["name"],
                        "what": f"Points at '{ref}', which is not in the scanned folder.",
                        "fix": "Either the file moved (link will show #REF!) or it lives outside "
                               "the folder — a shadow dependency. Find it and bring it in."})
        if n["volatile"]:
            out.append({"severity": "med", "kind": "Volatile formulas", "where": n["name"],
                        "what": f"{n['volatile']} cells use INDIRECT/OFFSET/TODAY-style functions.",
                        "fix": "INDIRECT and OFFSET are invisible to dependency tracing — no tool "
                               "can follow them. Replace with direct references where possible."})
        if n["hardcoded"] > 5:
            out.append({"severity": "med", "kind": "Hardcoded numbers", "where": n["name"],
                        "what": f"{n['hardcoded']} formulas contain a literal number of 3+ digits.",
                        "fix": "Move constants (rates, thresholds, prices) into a labelled "
                               "assumptions cell so they can be reviewed and changed in one place."})
        if n["in_degree"] == 0 and n["out_degree"] == 0 and not n["error"]:
            out.append({"severity": "low", "kind": "Orphan workbook", "where": n["name"],
                        "what": "Nothing feeds it and it feeds nothing that the scanner can see.",
                        "fix": "Either it is genuinely standalone, or a person moves data in and "
                               "out of it by hand. If so, record that handoff — it is invisible risk."})

    out.sort(key=lambda f: RISK_ORDER.get(f["severity"], 3))
    return out


def stats(nodes: list[dict], edges: list[dict]) -> dict:
    books = [n for n in nodes if n["kind"] == "workbook"]
    ext = [n for n in nodes if n["kind"] == "external"]
    high = [e for e in edges if e["risk"] == "high"]
    return {
        "workbooks": len(books),
        "sheets": sum(n["sheet_count"] for n in books),
        "formulas": sum(n["formula_count"] for n in books),
        "handoffs": len(edges),
        "fragile": len(high),
        "manual_pct": round(len(high) / len(edges) * 100) if edges else 0,
        "broken_links": sum(len(n["unresolved_refs"]) for n in books),
        "unreadable": sum(1 for n in books if n["error"]),
        "traced_columns": sum(len(e.get("columns", [])) for e in edges),
        "external_sources": len(ext),
    }
