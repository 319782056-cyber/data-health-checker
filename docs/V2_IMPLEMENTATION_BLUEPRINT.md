# Data Workbench 2.0 — Implementation Blueprint

> Status: approved product direction / implementation preparation  
> Working branch: `feature/data-workbench-v2`  
> Baseline: existing V1.0 on `main` remains the production fallback.

## 1. Product decisions (approved)

- Product model: AI + professional operation. AI proposes structured plans; deterministic data operations execute them.
- Audience: professional analysts and everyday office users, served by one application with adaptive Simple / Professional modes.
- Privacy: local-first processing by default. Cloud AI is optional and requires explicit, informed user authorization for the specific data/context being sent.
- Delivery: unified V2.0 release after integrated implementation and end-to-end acceptance. Engineering work may proceed in internal milestones, but no partial V2.0 feature release is implied.
- Existing repository and public Pages URL are retained. No V1.0 business-specific content is to be introduced.

## 2. Product principles

1. Original source files are immutable; transformations produce derived datasets/versions.
2. Every transformation is a typed, validated, deterministic step with parameters and execution metadata.
3. AI never directly mutates data or executes arbitrary code/SQL. It returns a constrained operation plan.
4. Every AI plan is inspectable and editable; execution requires user confirmation.
5. The UI must expose real navigation and working actions. Disabled/not-yet-implemented actions must be visibly disabled rather than decorative.
6. Simple and Professional modes share the same datasets, workflow, and execution engine.
7. Exported data must match the selected result version and clearly state which workbook features are not preserved.

## 3. Information architecture and routes

| Navigation | Route | Primary responsibility |
|---|---|---|
| Data overview | `/overview` | Projects, recent datasets, quality alerts, recent runs |
| Data browser | `/data` | Dataset/workbook navigation, grid, filters, sorting, column profile |
| Smart processing | `/transform` | Cleaning, type conversion, split/merge columns, calculated columns, step editor |
| Merge tables | `/merge` | Append and joins, key selection, match diagnostics |
| Data analysis | `/analysis` | Grouping, pivot, aggregation, calculated measures, charts |
| Workflows | `/workflows` | Workflow authoring, versioning, rerun, schema drift handling |
| Activity | `/activity` | Execution history, operation details, errors, undo/restore where supported |
| Settings | `/settings` | User mode, privacy preferences, AI consent controls, local data management |

AI assistant is a contextual side panel available from all major workspaces, not a separate chat-only product.

## 4. Page layout and interaction contract

- App shell: collapsible left navigation, top breadcrumb/command area, central workspace, optional right inspector/AI panel.
- Responsive behavior: on narrow screens, navigation and inspectors become drawers; core import, preview, and operation confirmation remain usable.
- Data browser: virtualized grid, server/engine-side paging where applicable, column menu, typed filter builder, sort, hide/reorder/freeze columns, column profile.
- Transform: field list + result preview + ordered applied-steps inspector. Selecting a step shows its parameters and before/after impact.
- Merge: source selection, join kind, key mapping, duplicate-key warnings, matched/unmatched counts, row-count impact preview.
- Analysis: drag/drop field wells with accessible alternatives, aggregation selector, result preview, export.
- Workflow: ordered graph/list representation first; visual node canvas is optional only after reliable execution semantics exist.
- Every primary button must have an implemented action, a clear disabled state, or an explicit not-yet-available explanation.

## 5. Technical architecture

### Frontend
- React + TypeScript + Vite (retain current foundation).
- React Router for explicit route ownership and working navigation.
- TanStack Table (or equivalent) for typed table state; add row/column virtualization for large previews.
- A shared design system and feature-level page components.
- Web Workers for CPU-heavy parsing/transformation so the UI remains responsive.

### Data plane
- DuckDB-Wasm as the analytical execution engine, subject to browser compatibility and realistic memory benchmarks.
- Apache Arrow as an efficient interchange format where useful.
- SheetJS remains the initial XLSX/XLS/CSV import/export adapter; test formula/value semantics and document unsupported workbook fidelity.
- IndexedDB for local project metadata, workflow definitions, preferences, and optional cached derived data. Avoid silently persisting original files without clear user controls.
- Source files are read-only. Derived datasets are addressed by stable IDs and execution/version metadata.

### AI plane
- Provider-neutral AI adapter; model provider can be changed without changing transformation semantics.
- AI input is minimized: schema and aggregate/profile summaries by default.
- Sending sample rows or selected source content to a cloud model requires a just-in-time consent screen that identifies what will be sent, the destination/provider, and the purpose.
- No API keys or provider secrets in browser code or public repository. Cloud AI calls go through a secured server-side endpoint with authentication, rate limits, validation, and secret storage.
- AI output must conform to a versioned JSON operation-plan schema and pass schema, field-reference, type, capability, and safety validation before preview.
- Deterministic engine executes only allow-listed operations. No arbitrary JavaScript or unrestricted SQL from model output.

### Hosting implications
- GitHub Pages can continue serving the static frontend.
- Optional cloud AI cannot safely be implemented by placing provider credentials in the static app. It requires a separately hosted server-side API (for example, a Vercel Function or another secured backend) and explicit environment configuration.
- Local-only workflows must remain usable when the AI backend is absent or unavailable.

## 6. Core domain model

Initial logical entities:

- Project: workspace container and user preferences.
- DataSource: source file metadata, import settings, content fingerprint; source bytes remain local unless explicitly authorized.
- Dataset: named table with schema and a reference to its source/derived lineage.
- ColumnSchema: stable column ID, display name, logical type, nullable flag, profile statistics.
- TransformStep: step ID, operation type, parameters, enabled state, input/output schema expectations.
- Workflow: ordered steps, source bindings, version, parameterization.
- Execution: status, timestamps, row counts, warnings, errors, engine version.
- DatasetVersion: immutable logical result reference and lineage to parent version/step.

Suggested operation families: select/rename/drop columns; filter; sort; cast type; trim/normalize; replace/fill null; deduplicate; split/merge columns; calculated column; group/aggregate; append; join; pivot/unpivot; export.

## 7. AI operation-plan contract

A plan should include:
- plan schema version;
- user intent summary;
- referenced dataset IDs and column IDs;
- ordered allow-listed operations and typed parameters;
- assumptions and unresolved ambiguities;
- expected impact and warnings;
- preview requirements.

Validation rules:
- Unknown operations, missing columns, incompatible types, invalid parameters, and unsupported functions are rejected.
- Ambiguous field mappings require user selection.
- Destructive-looking operations (drop columns/rows, overwrite/export destination) require explicit confirmation.
- The model cannot mark a plan as user-approved.

## 8. Initial repository structure target

```text
src/
  app/
    App.tsx
    router.tsx
    providers/
  components/
    layout/
    data-grid/
    common/
  features/
    overview/
    data-browser/
    transform/
    merge/
    analysis/
    workflows/
    activity/
    settings/
    ai-assistant/
  engine/
    contracts/
    operations/
    execution/
    adapters/
    workers/
  domain/
    models/
    validation/
  services/
    persistence/
    ai/
  styles/
  test/
docs/
  V2_IMPLEMENTATION_BLUEPRINT.md
```

This is a target structure, not a claim that these modules have already been implemented.

## 9. Unified V2.0 delivery gates

Internal engineering milestones are allowed, but the public V2.0 release is one integrated release.

1. Foundation: route shell, adaptive modes, design tokens, project/data-source state.
2. Data core: import adapters, DuckDB-Wasm worker, schema/profile, virtualized preview.
3. Transform and merge: validated operations, preview, lineage, undo/restore semantics.
4. Analysis and workflow: pivot/grouping, workflow save/replay, schema drift checks.
5. AI and privacy: constrained plan generation, consent UX, secured backend adapter, local fallback.
6. Hardening: accessibility, responsive behavior, performance, error handling, security and privacy review.
7. Release candidate: complete end-to-end tests, documentation, Pages/backend deployment verification, rollback plan.

No public V2.0 announcement until all release gates pass.

## 10. Acceptance checklist

- Every navigation item routes to the correct independent page and preserves relevant workspace state.
- Import multiple workbooks/sheets with Chinese headers, mixed date formats, blanks, duplicate headers, and malformed files.
- Filters, sorting, transformations, joins, pivots, and exports are checked against deterministic fixtures.
- Join tests include one-to-many and many-to-many row multiplication warnings.
- Workflow replay detects missing/renamed/type-changed columns and pauses safely.
- AI plan tests cover valid plans, unknown operations, invalid columns, ambiguous fields, malformed JSON, provider failure, and prompt injection in cell contents.
- Cloud AI consent is explicit, scoped, revocable, and never inferred from general terms acceptance.
- Local-only use works without network access to the AI service.
- No secrets are present in built frontend assets or repository history.
- Test realistic file sizes on representative desktop and mobile-class devices; publish measured limits rather than unsupported row-count promises.
- Build, automated tests, manual end-to-end checks, and deployment checks all pass before release.

## 11. Known V1.0 baseline and migration notes

V1.0 is a single React application in `src/App.tsx`, with state and transformations coupled to the page. Sidebar items are static non-interactive elements. The current app imports files into in-memory row arrays and provides basic text filtering, sorting, cleaning, pivot summaries, export, and session-only undo.

Keep `main` as the stable V1.0 production baseline while V2.0 is built on `feature/data-workbench-v2`. Reuse proven import/export behavior only behind adapters and tests; do not preserve the monolithic component architecture. Do not mix unrelated legacy project content into this repository.
