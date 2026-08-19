# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

PaperTrail: a local-only tool that scans a folder of Excel workbooks, works out which file/sheet/column
feeds which via cross-file formulas and connection metadata, lets a user annotate the human handoffs
(copy-paste, email, re-keying) no scanner can detect, and serves an interactive lineage map over HTTP on
`127.0.0.1`. No network calls, no telemetry — see README.md "Local-only" section for the invariants
(binds to loopback only, no HTTP client imported anywhere in `papertrail/`, CSP header on every response).

## Commands

```bash
pip install -r requirements.txt          # only dependency: openpyxl
python make_sample.py                    # (re)builds sample_books/ — a realistic demo folder
python -m papertrail sample_books        # scan + serve; opens the browser
python -m papertrail ~/Finance --port 9000 --no-browser
./run.sh   # or run.bat on Windows        # convenience launcher, installs deps if missing
```

There is no test suite and no lint config in this repo. Verify changes by running a real scan against
`sample_books/` and checking the JSON: `python3 -c "from papertrail.scanner import scan_folder; from papertrail import graph as G; g = G.build_graph(scan_folder('sample_books')); print(g['stats'])"`.
For frontend changes, the fastest correctness check is a syntax check on the extracted `<script>` block
(`node --check`) plus an actual click-through — the UI has no build step, `papertrail/web/index.html` is
served byte-for-byte from disk on every request, so editing it and refreshing the browser is enough.

Packaging (only needed when cutting a release): `pyinstaller --onefile --name PaperTrail --add-data
"papertrail/web:papertrail/web" --collect-submodules openpyxl entrypoint.py` builds a standalone binary;
`dist/` holds the built packages and is not meant to be committed.

## Architecture

Three Python modules, one HTML file, no framework, no build step:

```
papertrail/scanner.py   opens .xlsx/.xlsm read-only, returns BookInfo/SheetInfo dataclasses
papertrail/graph.py     scan result -> {nodes, edges, findings, stats} — the entire API response shape
papertrail/server.py    stdlib http.server, JSON API + CSV export, in-memory STATE dict
papertrail/web/index.html   single-file UI: inline CSS + one <script> block, no external assets
```

**Data flow is one-directional and stateless per request**: `scan_folder()` walks the folder and produces
plain dicts (via `dataclasses.asdict`); `build_graph()` takes that plus the on-disk annotations file and
produces the graph the frontend renders. Nothing is cached between requests except `server.STATE`
(root/scan/graph), rebuilt wholesale on every `/api/scan`, `/api/rescan`, and `/api/annotate` call — there
is no incremental update path, re-scanning is always a full re-walk of the folder.

**Cross-file lineage has two granularities that both flow from the same raw data.** `scanner.py` regex-matches
every formula cell for external references (`RE_EXTCELL`) and records one entry per match in
`BookInfo.links` (source book/sheet/cell -> dest sheet/cell/column/header). `graph.py`'s `_column_lineage()`
rolls those up per edge into `edge["columns"]` (grouped by source-column -> dest-column pair, capped at
`MAX_COLUMN_ROWS`). The file-level edge (`formula` mechanism, one row per book pair) and the column-level
detail are therefore always in sync by construction — never hand-maintain one without the other.

**External/website/database source detection is a second, parallel discovery mechanism** alongside formula
scanning, because Power Query and legacy web queries don't leave formula text — they live in
`xl/connections.xml` and `xl/queries/*.xml` inside the zip container. `scanner._external_sources()` regex-scans
those XML parts (see `RE_M_SOURCE` for the Power Query "M" source functions it recognizes: `Web.Contents`,
`Sql.Database`, `SharePoint.Contents`, etc.) and `_redact()` strips credential-shaped tokens
(`Pwd=`/`UID=`/...) before anything is ever returned — do not remove that redaction. `graph.build_graph()`
turns each distinct `(kind, detail)` pair into one synthetic node (`kind: "external"`, id hashed via
`_ext_id`), even when several workbooks share the same source. `WEBSERVICE()`/`FILTERXML()`/`RTD()` cell
formulas are a third path into the same external-node system (`web_formula` mechanism), scanned inline
alongside cross-file formulas rather than from the zip's XML parts.

A connections.xml source has no formula to point at, but `scanner._connection_sheet_map()` can still
usually resolve *which sheet* it loads into, by walking the same relationship chain Excel itself uses
(`xl/workbook.xml` sheet name -> r:id -> `xl/_rels/workbook.xml.rels` -> `worksheets/sheetN.xml` ->
`xl/worksheets/_rels/sheetN.xml.rels` -> a queryTable part -> its `connectionId`). Attribute order inside
an OOXML tag is not guaranteed, so every tag is matched whole and its attributes pulled out independently
via `_attr()` — do not add a new regex here that assumes one attribute precedes another. When a sheet is
resolved, `graph._sheet_sentinel_columns()` turns it into a one-row synthetic `columns` entry (real column
letter unknown, dest_col is literally `"(query result)"`) so the existing per-sheet lineage machinery below
picks it up for free.

**The frontend has two drill-down levels, both built from the same `columns` data at different
granularities.** Selecting a workbook first shows a *file-level* radial diagram (`radialSVG()`): center =
the file, ring 1 = every one of its sheets (even ones with no cross-file connections — `tabSummary()`
collapses `bySheet()`'s per-column rows up to one row per (peer file/source, sheet) pair so a sheet's
several peers still land on one ring-1 spoke), ring 2 = the peer files/external sources fanned out from
each sheet's own spoke. Click a ring-1 node (or a row in the plain sheet list below) to open the
*sheet-level* column diagram (`sheetRowHTML`, drawn by the separate `flowSVG()` two-lane primitive —
columns don't fan out radially as cleanly as files-per-sheet do, so this level stays linear). Click a
ring-2 node to jump to that peer via `openPeer()` (dispatches to `select()` or `selectExternal()` by node
kind). If a sheet has many peers, `radialSVG` widens that sheet's angular span and staggers label
radius/offset by index parity so labels don't collide — don't reintroduce a fixed span/offset, a "collector"
sheet fed by ten files needs much more room than one fed by two.

**A third view, `drawBigFlow()`, takes over the main canvas** (the same `#svg` element `drawGraph()` draws
the pipeline map into) when `flowView` (a workbook id, or `null`) is set via `openBigFlow()` — reachable by
clicking a radial diagram's center hub, or the "⤢ open full flow" link next to it. It lays out one
workbook's SOURCES → TABS → COLUMNS → DESTINATIONS left to right using the *same* box-and-curve drawing
conventions as `drawGraph()` (empty tiers are dropped, e.g. a pure source file has no SOURCES column), just
with real screen width instead of the 342px side panel. `render()` checks `flowView` and calls
`drawBigFlow()` instead of `drawGraph()` whenever it's set, so anything that mutates state and re-renders
(saving an owner, rescanning) keeps the user in the big-flow view rather than snapping back to the map.
Column boxes wire straight into `selectEdge()`, so clicking one opens the exact formula-level detail in the
side panel without leaving the big view — the three drill-down levels (map / big-flow / edge-or-sheet
detail) are meant to compose, not replace each other.

**The side panel is ordered by how often each thing is actually used, not by data type** — deliberately, after
an early version buried the radial "Tabs & connections" diagram below a wall of raw stat rows, requiring two
full scrolls to reach it. Title → the diagram (with isolate/full-flow controls right next to its heading) →
a one-line `.statstrip` of the same numbers that used to be seven stacked `.kv` rows → owner/fed-by/feeds →
the plain sheet list last, since the diagram is now the primary entry point into it. Keep new high-value,
frequently-clicked content near the top; raw metadata belongs at the bottom.

**`updateScrollHints()` (the "→ more" / "← more" pills on the map canvas) has one real gotcha**: calling it
while `#map`'s section is still `.hidden` (`display:none`, e.g. during the initial page-load render before
`goTab("map")` reveals it) reads `clientWidth`/`scrollWidth` as 0, so it silently computes no overflow. That's
why the `.tab` click handler calls `updateScrollHints()` again *after* un-hiding the section, on top of the
calls at the end of `drawGraph()`/`drawBigFlow()`. If a similar overflow-detecting affordance gets added
elsewhere, it needs the same "re-check after becoming visible" call, not just a call at draw time.

**Node "stage" (map column) is deliberately computed from book<->book edges only** (`book_indeg`/`book_outdeg`
in `build_graph`), then workbook stages are shifted +1 and external nodes pinned to stage 0. This is so a
workbook that happens to also pull from a website isn't visually shoved out of its place in the file-to-file
pipeline — see the comment at that point in `graph.py` before changing it.

**The frontend is one big script with a handful of module-level state variables** (`G` = current graph
payload, `selected`, `traceMode`, `riskMode`, `isolateMode`, `chainState`, `expandedSheet`) that every
render function reads. `select(id, keepSheet)` is the central re-render for the workbook side panel;
`keepSheet` controls whether `expandedSheet` survives the re-render — pass `true` from any handler that's
just refreshing state the user is already looking at (save owner, toggle isolate, step a flow), and leave it
default (`false`) for a fresh click on a different node. `selectEdge()` and `selectExternal()` are the other
two side-panel entry points, for clicking an arrow or an external-source node respectively.

**One diagram primitive, `flowSVG()`, draws every box-and-curve lineage diagram in the app** — the edge-detail
view and the per-sheet drill-down both call it, differing only in the `leftKey`/`rightKey`/`colorOf`
functions passed in. `bySheet()` pre-groups a node's incoming/outgoing edge columns by sheet name once per
`select()` call so `sheetRowHTML()` never re-scans edges itself. If you're adding another lineage view,
extend `flowSVG`'s options rather than writing a new SVG-string builder.

**Map node height is dynamic**, not a constant — `nodeH(n)` computes it from sheet-tab-row count (workbook
nodes list their sheet names inline, click one to jump straight to that sheet's column drill-down) and
`layout()` stacks nodes in each stage column using per-node height, not a fixed row height. Edge routing in
`drawGraph()` reads `pos[id].h` for the vertical midpoint — if you change node content, keep `nodeH()` and
the actual rendered content in sync or edges will attach at the wrong point.

Annotations (owner, notes, declared handoffs) are the only mutable, persisted state, written to
`<scanned-folder>/papertrail.annotations.json` — never sent anywhere, loaded fresh on every graph rebuild
via `graph.load_annotations()`.
