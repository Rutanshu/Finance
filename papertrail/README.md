# PaperTrail

**See how finance data really moves through your spreadsheets.**
Runs entirely on your own machine. Nothing is uploaded, ever.

---

## The problem

Excel is not the problem. **Undocumented handoffs between spreadsheets** are the problem.

A formula error is visible and usually caught. What is invisible is that the revenue block in your
trial balance is *pasted in by one person every month*, from a file that is itself pasted together
from two exports whose date filters are chosen by hand. Nothing in the workbook records that. It
lives in one person's memory — and it walks out of the door when they do.

PaperTrail reads your workbooks, works out which file feeds which — down to the exact **tab, column
and formula** — lets you record the human handoffs no scanner can detect, and gives you one clickable
map of the whole pipeline.

---

## Install and run

Requires Python 3.9+.

```bash
pip install -r requirements.txt

python make_sample.py                 # optional: builds a demo folder to try it on
python -m papertrail sample_books     # scan the demo and open the UI
```

To point it at your own files:

```bash
python -m papertrail                          # then paste your folder path into the UI
python -m papertrail "C:\Finance\Month-end"   # or scan on startup
python -m papertrail ~/Finance --port 9000 --no-browser
```

Then open <http://127.0.0.1:8737>.

Convenience launchers: `./run.sh` (macOS/Linux) or `run.bat` (Windows).

---

## How to use it

**1. Scan.** Point it at your month-end folder. Every cross-file formula reference —
`='[Revenue_Register.xlsx]Data'!$B$4` — plus the external-link table hidden inside each file becomes
an arrow on the map. Click any arrow (or a "N columns traced" chip on a workbook) and it opens into the
exact tab and column on each side, how many cells are involved, and a sample of the formula text itself.

**2. Declare the invisible handoffs.** The honest limit of every scanner ever written: *a copy-paste
leaves no trace in the file.* Neither does an email or someone re-keying figures off a PDF — and those
are usually your highest-risk hops. Click **+ Declare a handoff**, pick source and destination, say how
the data actually moves. Ten minutes with whoever runs the close captures them all.

**3. Trace.** Turn on **Trace** and click a workbook. Everything unrelated fades out:

- **Upstream** — "Revenue looks off." Follow it back. You usually land on a dashed line.
- **Downstream** — "The bank changed its CSV format." See instantly which reports are now unsafe,
  *before* anyone presents them.

**4. Work the findings.** Ordered worst-first. Fragile handoffs, broken links to files outside your
folder, volatile formulas no tool can trace through, hardcoded constants, orphan workbooks.

**5. Re-run monthly.** Annotations persist in `papertrail.annotations.json` inside your own folder, so
you describe each handoff once. Export the Handoffs tab to CSV as audit evidence.

### The one fix worth doing first

At every fragile handoff, carry **one number that must match on both sides** — a row count, a gross
total, or a hash of the key column. Put it in a visible cell with conditional formatting that turns
red on mismatch. Most spreadsheet finance errors surface within seconds once that exists.
PaperTrail tells you where to put them.

---

## What you get

| | |
|---|---|
| **Close faster** | Most "the numbers don't tie" hours go on finding *where a figure came from*, not fixing it. That becomes a click. |
| **Survive an audit** | Auditors want the end-to-end flow of financial information. Export the Handoffs tab — accurate, because it was read out of the files. |
| **Reduce key-person risk** | When the person who runs the close leaves, the map stays. |
| **Make the risk arguable** | "Our process is a bit manual" gets no budget. "Sixty percent of the pipeline moves by copy-paste, here are the eleven hops, ranked" gets a decision. |
| **Aim automation where it pays** | The fragile handoffs are your migration order. Fix eleven hops instead of rebuilding everything. |

---

## Local-only — and how to verify it

Finance data is the most sensitive data a company holds. This is not a cloud service with a privacy
policy; it is a program on your machine with no way to phone home.

- The server binds to `127.0.0.1` only — **not** `0.0.0.0`. Nothing else on your network can reach it.
- **No HTTP client is imported anywhere in the package.** Check it yourself:
  ```bash
  grep -rn "requests\|urllib.request\|httpx\|socket.connect\|urlopen" papertrail/
  ```
  Returns nothing.
- No CDN, no web fonts, no analytics, no update check. The page is one HTML file served from disk,
  with a `connect-src 'self'` content-security policy on top.
- Workbooks are opened **read-only** and never modified. The only file written is
  `papertrail.annotations.json`, in your own folder.
- Want certainty? **Pull the network cable and run it.** Everything works.

---

## How it works

Two data structures and one traversal.

```
NODES  = one per workbook   (sheets, formula counts, owner, stage)
EDGES  = one per handoff    (from, to, mechanism, risk, columns[])
```

Each edge's `columns` list is the cross-file links rolled up by source-tab/column →
destination-tab/column, so you get "Data!B feeds Summary!C — 42 cells" instead of a wall of individual
cell references, with a sample formula attached to each row.

**Risk is not a judgement — it is read off the mechanism.** Who moves the data?

| Mechanism | Risk | Why |
|---|---|---|
| Automated refresh / query | low | Auditable, repeatable |
| Cross-file formula | medium | Breaks silently when a file moves or is renamed |
| Copy-paste by hand | **high** | No audit trail at all |
| Typed in by hand | **high** | No audit trail, plus transcription error |
| Emailed between people | **high** | Version ambiguity on top of everything else |

Tracing is a depth-first walk over `EDGES` — `"up"` for ancestors, `"down"` for descendants. Same
function, one flag. Everything visual is downstream of that; the map, the CSV export and the findings
list are three renderings of the same result.

```
papertrail/
  scanner.py   opens workbooks read-only, extracts formulas + external links
  graph.py     builds nodes/edges, scores risk, DFS traversal, findings
  server.py    localhost-only HTTP server, JSON API, CSV export
  web/         the single-file glassmorphic UI
make_sample.py builds a realistic demo folder
```

---

## Limits — stated plainly

- For cross-file formulas it resolves the exact source/destination **tab and column**, but it does not
  recompute the number for you. It tells you *where to look*, not what the number should be.
- A formula that points at a named range or a whole column (`A:A`) is still shown as a file-level
  handoff — the column breakdown only fills in when the formula names a specific cell or range.
- `INDIRECT` and `OFFSET` build references at runtime, so no static tool can follow them. They are
  flagged, not traced.
- Files outside the scanned folder appear as broken links — which is itself the point. Shadow
  workbooks on someone's desktop are exactly what you want to discover.
- Password-protected or corrupt files are reported as unreadable rather than skipped silently.
- Lineage *within* one workbook (sheet A feeds sheet B in the same file) is still summarised at the
  sheet level, not drawn cell-by-cell. If you need that, [pycel](https://github.com/dgorissen/pycel)
  is the tool for it.

---

## Where it sits next to existing tools

Cell-level tools ([pycel](https://github.com/dgorissen/pycel),
[ExceLint](https://github.com/ExceLint/ExceLint-addin)) map *A1 feeds B7* inside one workbook.
Enterprise lineage platforms ([OpenLineage](https://github.com/OpenLineage/OpenLineage),
[Marquez](https://github.com/MarquezProject/marquez)) map database pipelines and assume instrumented
jobs. Neither can see a human carrying a file between two workbooks — which, in a spreadsheet-run
finance function, is most of the pipeline. That gap is what PaperTrail fills.

---

MIT-style: do what you like with it.
