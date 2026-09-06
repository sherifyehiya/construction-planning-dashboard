# construction-planning-dashboard
KPI Planning Dashboard

Static, client-side dashboard: `index.html` + `app-data.js` + `data/manifest.json` + `lib/`.
Parses Primavera P6 (`.xer`/`.xml`) and Excel (`.xlsx`, via SheetJS) exports entirely in the
browser — no backend, no build step. Data date, source formats and "LIVE" status in the
footer are all computed from whatever's actually loaded, not hardcoded.

## KPI registry & rollup engine

`app-data.js` exposes two things every KPI in the dashboard is built from:

- **`rollup(level, scopeKey, model)`** — aggregates the raw parsed model (activities,
  assignments, resources, relationships) into one node of measures `{ev, pv, cpi, spi, eac,
  weightedComplete, count, criticalCount, negFloatCount, unlinkedCount, avgUtilization, ...}`
  for one scope: `portfolio` (everything), `project` (`{project: ProjectKey}`), `wbs`
  (`{project, wbs: WBSKey}`), or `activity` (`{activity: ActivityKey}`).
- **`KPI_REGISTRY`** — one array, every KPI defined once:
  ```js
  {
    id,          // 'spi' — unique
    aspect,      // 'schedule' | 'cost' | 'progress' | 'integrity' | 'resources' | 'risk'
    name,        // 'Schedule Performance Index'
    shortName,   // 'SPI' — shown on cards
    definition,  // one-line plain-English meaning, shown nowhere yet but ready for a tooltip
    formula,     // (node) => number|null — reads a rollup() node, no math outside this
    unit,        // 'ratio' | 'pct' | 'egp' | 'days' | 'count'
    decimals,
    target,      // e.g. 1.0 — null if there's no defensible target (shown as "informational")
    direction,   // 'higher-better' | 'lower-better'
    rag,         // {green, amber} thresholds — everything past amber is red; null = no RAG
    levels,      // which of portfolio/project/wbs/activity this KPI is valid at
    isHeadline   // true = shown in the aspect's Verdict slot
  }
  ```

**To add a new KPI**: add one object to `KPI_REGISTRY` in `app-data.js`. It appears in its
aspect tab automatically — Verdict (if headline), Breakdown-by-next-level (if it's the
aspect's first headline KPI), and it's available to any future slot that reads the registry.
No `index.html` edit needed unless the underlying `rollup()` node doesn't yet carry the raw
measure your formula needs (in which case, add that one field to the node it returns).

## Aspect anatomy (in `index.html`)

Six aspect tabs (Schedule / Cost / Progress / Integrity / Resources / Risk) all render off
the **same** `renderAspect(aspectId, container)` function — four slots, always in this order:

1. **Verdict** — headline KPI cards (first one as a gauge) with value, target, RAG color.
2. **Breakdown** — the aspect's primary KPI split by the next level down (portfolio → by
   project, project → by WBS), sorted worst-first.
3. **Exceptions** — breakdown rows currently in the red for that KPI ("what's failing now").
4. **Detail** — the activities behind the current scope.

A breadcrumb-style scope selector (Project → WBS dropdowns, shown only on aspect tabs)
re-scopes all six aspects at once via `rollup()` — that's the only navigation concept beyond
picking a tab. Portfolio Overview and Activity Explorer are intentionally *not* aspects; they
stay as their own tabs (portfolio-wide landing view and a flat searchable activity table,
respectively — neither maps cleanly onto the aspect/drill-level model).

## Data source panel

The ⚙ icon next to the theme toggle opens a small popup: currently-loaded files (from
`data/manifest.json`) plus links to GitHub's own upload/edit UI for `data/`. It's behind a
password prompt, but **that's a soft, casual gate, not real security** — this is a public
static page, so anything shipped in the page's JS (including the password check) is readable
via view-source and bypassable in devtools. What actually protects the data is that the
upload links require a real GitHub login with write access to this repo. To change the
password, see the comment above `DATA_GATE_HASH` in `index.html`.
