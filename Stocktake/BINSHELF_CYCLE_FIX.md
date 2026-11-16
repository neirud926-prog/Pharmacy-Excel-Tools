# Stocktake – BinShelf "1" / "T" never counted (investigation & fix, 16-Nov-2026)

Reported by: Chloe (NLTH, Osaka). The Daily Stock Take never reaches the BinShelf "1*" and "T*" bins.
Evidence: `Stocktake/Record.txt` (Access `[Record]` export, IDs 6601–8266, 04-May-2026 – 16-Nov-2026) and the `BinShelfLocator` sheet in `Stocktake.xlsm` (705 items).

## Finding

The report is confirmed, and the scope is wider than "1" and "T". **216 of the 705 BinShelfLocator items have had no Daily count since 04-May-2026**:

| Bin group | Items never Daily-counted |
|-----------|---------------------------|
| 1*        | 99 (every 1* bin except 1A2A, which had one off-sequence count) |
| R*        | 30 (all) |
| T*        | 29 (all) |
| P*        | 25 (all) |
| MF2A and later M* | 20 |
| 2A1B–2A5A | 11 |
| S*, 6B9B  | 1 each |

The full list is in `Daily_Not_Counted_Since_2026-05-04.csv`.

## Root cause

`GetNextBinShelf` (DBCRUD) returned **the bin after the latest saved Daily record** (`ORDER BY RecDate DESC, ID DESC`). "Print 5 items" then printed the next 5 rows of BinShelfLocator. Whatever bin was saved last became the new position, including spot checks of other locations entered as Daily. This is the suspected cause in the email: current bin 6A6A, someone enters 2B3A, and the sequence moves to 2B3A.

What the record shows:

| Date | Saved by / Ref | What happened | Effect |
|------|----------------|---------------|--------|
| 03-Sep-2026 | wwh136 / ST2026090321 | Counted ME2B–ME5A. The next bin should have been **MF2A**, followed by P, R, S, T and then wrap-around to 1A2A. | – |
| **04-Sep-2026** | sym656 / ST2026090416 | Off-sequence Daily batch: 3B1A, 3A4A, 2A8A, 3B7A, 6A2A, **2A8A** (last row). | Position jumped to **2A8B**. MF*+, P*, R*, S*, T* and all 1* were skipped, and the cycle restarted in 2A. |
| 05-Jun-2026 | cpc490 / ST2026060516 | Off-sequence batch 2A8A, 2B2B, 2C3B, 2B6B, **2C9C**. | Position jumped back to 3A1A, and 3A–3B were counted a second time (5–15 Jun). |
| 16-Jun-2026 | wfw325 | Restarted at 4A8D. | 4A8D was counted twice. |
| 10-Nov-2026 | wwh136 / ST2026111021 | Counted 5B7A–5B8A, which lkt773 had already counted on 07-Nov. The sheet was probably printed before the earlier batch was saved. | Duplicate work. Then 5B8B–5C1A were also repeated on 12-Nov. |

The cycle that started on 04-May began at 2A6A, so 2A1B–2A5A were never reached. Because the rotation was reset on 04-Sep, it never reached the end of the list again.

## Fix (agreed in the grilling session)

The rule is recorded in ADR `docs/adr/0001-stocktake-daily-cycle-by-coverage.md`, and the terms are defined in `GLOSSARY.md`.

* **Daily Cycle**: replay every Daily record in (RecDate, ID) order. The cycle completes only when **every** BinShelfLocator item has been counted, and the next cycle starts after that.
* **Next Items**: the items not yet counted in the current cycle, in bin order. The latest record no longer drives the sequence.
* An **off-sequence entry** counts that item as done, but it cannot move or skip the sequence.
* **Print 5 items** prints the next 5 uncounted items. Items already counted in this cycle are never reprinted, and if fewer than 5 remain it prints what is left.
* When a Daily save finishes the cycle, a message reports it, and the next print starts a new cycle from the first bin.
* Annual records are ignored. No database change is needed.

Expected after the fix: the next print is **1A2B, 1A3A, 1A3B, 1A4A, 1A4B**. The current cycle (started 04-May-2026) has 489 of 705 items counted and 216 to go, about 43 sessions at 5 items each.

### Changed files (re-import into `Stocktake.xlsm`)

| File | Change |
|------|--------|
| `VBA/Modules/BinCycle.bas` | **New.** Cycle replay, `GetNextCycleItems`, `GetCyclesCompleted`, `GetCycleSummary`, `NotifyIfCycleCompleted`. |
| `VBA/Modules/DBCRUD.bas` | `GetNextBinShelf` delegates to BinCycle. `SaveListViewToAccess` shows a message when a Daily save completes a cycle. |
| `VBA/Modules/WSPrinting.bas` | `Print5itemsWorkSheet` prints the next 5 uncounted items and clears unused rows. `test333` shows the cycle progress. |

To install: in the VBA editor, import `BinCycle.bas`, then remove and re-import `DBCRUD.bas` and `WSPrinting.bas`, and save. Run `test333` to check the result: it should show "Next bin: 1A2B" and "489 of 705".

### Not verified here

The logic was checked against `Record.txt` with a Python model of the same rule. The VBA itself has not been compiled or run, because no Excel or Access is available in this environment. Run Debug ▸ Compile and `test333` on a pharmacy PC before use.
