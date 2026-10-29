# Customer Allocation v2.0 Implementation Plan

> Execution: User explicitly requested unattended end-to-end implementation; no further approval gates. Use test-first development and independent file ownership.

**Goal:** Deliver the complete approved local customer allocation query module, filtered XLSX export, and verified Windows v2.0 package.

**Spec:** ../specs/2026-10-09-customer-allocation-v2-design.md

**Architecture:** Keep the existing local Python server and reviewed quantity data. Isolate customer parsing/matching/query planning, streaming XLSX output, session-version coordination, and UI. Do not change XLS templates or quantity parsing.

## Tasks

- [x] Customer engine (`customer_allocation.py`, `test_customer_allocation.py`): header discovery, required fields per template, code-preserving deduplication, strict dimensions, grouped totals, lazy detail/issue rows, validated filters, pagination. Tests cover all four templates, zero vs unmatched, duplicates, conflicting products, manual quantities, and units.
- [x] XLSX writer (`query_workbook.py`, `test_query_workbook.py`): stream typed safe cells to XLSX, readable widths/styles, repeated filter metadata and headers, automatic sheet splitting, row/sheet progress, all-or-nothing destination. Verify ZIP/XML and read-back including >1,048,576-row boundary.
- [x] Session coordination (`customer_service.py`, server integration/tests): shared confirmed product choices, revision-bound customer queries, invalidation on strategy/customer/choice changes, immutable export jobs and progress, protected save/open targets, no stale export.
- [x] Query UI (`customer_ui.py`, CSS, server/review integration): tab navigation, customer upload, confirmation independent of XLS save, multi-select filters, summary/detail/issues, pagination, incomplete totals indicators, stale/empty/error/busy states, current-results-only export and save/open controls.
- [x] Verification: backend tests, real customer workbook read-only inspection, 20,000-customer/200-product synthetic query and multi-sheet export, browser end-to-end lifecycle and matching UI/export assertions, desktop/mobile screenshots, existing v1.9 regression tests.
- [x] Delivery: update README, v2.0 version, build dated Chinese Windows package, verify executable archive and unchanged templates. Do not replace or delete desktop versions or push business data to Git.

## Contracts

`parse_customers(content, template_key, book_factory)` returns a dataset object/dict suitable for `AllocationQuery(dataset, review, choices, template_key)`. Query engine exposes `filter_options()` and `plan_query(filters, view)`; each plan exposes `columns` (key/label dictionaries), `totals`, `row_count`, `page(page, page_size)` and lazy `iter_rows()` (row dictionaries).

Filters: `productRows` (source-row integer list or null), `counties`, `markets` (string lists or null), `minTier`, `maxTier`. Null means all, empty list means none. Views: `summary`, `detail`, `issues`.

`write_query_workbook(path, columns, rows, metadata, sheet_row_limit=1048576, progress=None)` returns row/sheet counts. `columns` have key/label; rows are dicts. No formulas are emitted for input text. Metadata is a list of human-readable strings. Output is XLSX only.

Server session owns confirmed choices and a monotonic generation; query IDs refer to immutable engine plans with their generation. Applied strategy mutations and replaced customers invalidate existing plans. Export starts only against a current query, pins plan and workbook target while writing, and blocks mutations until complete. Long-running job adds a lifetime work lease and always releases it.

## Review Focus

Do not silently coerce unknown dimensions to valid zero; do not double-count repeated customer codes; do not export only the visible page; preserve numeric codes as strings; ensure pinned export never reads a newer live filter state; include no customer files or identifiers in Git.

## Verification Evidence (2026-10-09)

- Final Python suite: 62 tests passed, no skips, with `RUN_CUSTOMER_SCALE=1`, `RUN_XLSX_BOUNDARY=1`, `TEST_NATIVE_DIALOG=1`, and the built executable in `PACKAGED_TOOL_EXE`.
- Real engine: 20,000 synthetic customers x 200 products; 4,000,000 exported rows; four sheets; every code and quantity checked with bounded-memory XML reading. 19,820,000 units = 79,280.000 boxes. Write 59.840 seconds; verification 110.662 seconds; test-process peak working set 178,438,144 bytes. Temporary output removed.
- Excel physical row boundary separately tested with 1,048,576 data rows, repeated metadata/headers and correct second-sheet continuation.
- Eight browser scripts passed: customer query, county multiselect, product export list, search, UI states, zero export, 181-product workbench, native-save workflow. Desktop/mobile screenshots visually inspected; no page overflow.
- Customer query read-back verifies all filtered rows beyond the first page, text codes and totals; applied strategy changes invalidate earlier queries/exports. Export failure preserves the existing destination and releases the exit lock.
- Independent review found and prompted regression fixes for conflicting-customer issue rows leaking outside filters and invalid-customer cross-product memory. Revised plans retain shared groups and per-product counts; regression memory checks cover invalid and unknown-business-area workloads.
- Real customer workbook inspected read-only: explicit table-footer metadata is excluded, no customer identifiers copied into source/tests. Additional real-dimension query exercised 200 synthetic products and 3,882,400 potential details.
- Four original XLS formats retain sheet names, merged cells, hidden type identifiers and selected period. All nine packaged private template files match original bytes.
- Release: `dist/v2.0-20261009/卷烟投放策略转excel工具.zip` at the workspace root, 11,102,866 bytes. SHA-256: `ca355de2d18eedc361701be76abce3f284281743fa651fcfcd903eeb33923ffb`.
- Packaged EXE starts, reports v2.0, loads templates, parses synthetic strategy/customer uploads, and returns correct customer query results. Compiled archive contains all new modules. ZIP integrity checked; no customer workbook or generated report bundled.
- No desktop replacement/deletion or Git push performed. macOS paths retained but not tested on a Mac.
