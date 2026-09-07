# Short Expiry Log Book – date integrity redesign

## 1. What was going wrong

`Use Before Date` is typed into a free-text box on `frmEntry`. The old code
accepted anything `IsDate()` liked and wrote the **text** into the cell:

```vb
Target.Offset(0, 5).Value = Me.ListToAdd.List(i, 3)   ' a String, e.g. "11/1/2026"
```

When VBA assigns a date-looking string to a cell, Excel coerces it by trying
the **en-US m/d/y** reading first and the PC's regional format second. So the
same keystrokes give different results on different PCs:

| Typed      | Meant (HK habit) | Stored on a d/m/y PC | Stored on an m/d/y PC |
|------------|------------------|----------------------|-----------------------|
| `8/1/2021` | 1 Aug 2021       | 8 Jan 2021 (US wins) | 8 Jan 2021            |
| `20/6/2021`| 20 Jun 2021      | 20 Jun 2021          | **text** `20/6/2021`  |
| `01062021` | 1 Jun 2021       | 01/Jun/2021 ✔        | 01/Jun/2021 ✔         |

The 8-digit shortcut in `cbUBDate_Change` was the only safe path, and only
because it spelled the month out.

The same bug hit `ChangeStatusToEnd`, which stored `Now` in a **String**
before writing it to the Reference column.

### Evidence in the current workbook (3 797 rows)

| Column            | Real dates | Text values | Contradictions |
|-------------------|-----------:|------------:|----------------|
| Date of Checking  | 3 386      | 411         | –              |
| Use Before Date   | 3 699      | 98          | 88 rows expire *before* their checking date (day/month swapped) |
| Reference         | 108        | 1 437       | –              |

Text dates are invisible to the month filters (`UBDateRangeFilter`,
`UBDateStatusFilter` compare numeric serials), so those rows never appear on
the printed report and never get closed. 2 974 of the real dates have a day
≤ 12, i.e. they *look* fine but cannot be proven right from the cell alone.

## 2. Design rules

1. **Refuse ambiguity at the door.** `DateSafe.TryParseExpiry` is the only
   function allowed to interpret typed dates. It accepts
   `01112026`, `01-Nov-2026`, `2026-11-01`, `Nov-2026` / `11/2026` (first day
   of month) and a numeric `d/m/y` **only when one reading is impossible**
   (`20/6/2021`). `11/1/2026` is rejected with a message showing both
   readings and how to type it.
2. **Build dates from parts.** Every date is `DateSerial(y, m, d)` from
   explicitly identified parts, then checked for roll-over (`31/02` is
   refused). `CDate`, `IsDate`, `Format` are never used to *interpret* input.
3. **Confirm before commit.** The box is rewritten as `dd-Mmm-yyyy` and a
   preview label shows `Sun 01-Nov-2026 (424 days)`. Past dates and dates
   more than 10 years out need a Yes/No confirmation.
4. **Carry the number, not the text.** The staging list keeps the serial in a
   hidden column; `btnSubmit` writes that serial.
5. **Write serials, never strings.** `WriteDateCell` sets the number format
   and assigns `Value2 = CDbl(date)`. No coercion is possible.
6. **One display format everywhere.** `[$-409]dd-mmm-yyyy` on Date of
   Checking and Use Before Date, `[$-409]dd-mmm-yyyy hh:mm` on Reference.
   The `[$-409]` locale tag pins the month name to English (`Jan`, `Feb`),
   so every PC shows the same thing whatever its regional settings or Excel
   display language.
7. **All-or-nothing save.** Every staged row is validated before the first
   write; any error rolls back the rows added in that batch.
8. **Repair is audited, not silent.** Legacy cells are listed on a `Date
   Audit` sheet with a proposed value and a reason; applying them backs up
   the sheet and writes a note into Remarks.

## 3. Decisions

Rows marked **confirmed** were answered by the pharmacy; the rest are my
assumptions. Change any of these and the code needs a one-line change in the
place named.

| # | Question | Assumed answer | Where |
|---|----------|----------------|-------|
| 1 | Does the site read numeric dates as **day/month/year**? | **Confirmed.** Used **only** by the repair routines to read legacy text like `20/6/2021`. The entry form never assumes it. Display is always an English month name, never a number. | `TryParseExpiry(..., assumeDMY:=True)` in `DateRepair` |
| 2 | Should `11/1/2026` be rejected outright, or resolved by a per-user setting? | **Confirmed:** rejected; `20/6/2027` (only one reading possible) is accepted. A per-user preference is exactly what caused the bug. | `TryParseExpiry` |
| 3 | Is a **month-only** expiry (`Nov-2026`, `11/2026`) valid, and which day does it mean? | **Confirmed:** valid, stored as the **first** day of that month (conservative reading). | `FirstDayOfMonth` in `DateSafe` |
| 4 | Are 2-digit years allowed (`31/10/26`)? | **Confirmed:** yes, as 20xx. Legacy data already contains them. | `ExpandYear` |
| 5 | May an already-expired item be logged? | **Confirmed:** allowed after a Yes/No warning (stock found expired on the shelf). | `ExpiryPolicyWarning` + `btnAdd_Click` |
| 6 | What is "implausibly far away" for short-expiry stock? | Warn above 10 years on entry; audit flags > 3 years after the checking date. | `EXPIRY_MAX_YEARS_AHEAD`, `PLAUSIBLE_MONTHS_AHEAD` |
| 7 | Can historic rows be auto-corrected? | **Confirmed:** only when the cell is *provably* wrong: text → date, or expiry before the checking date whose day/month swap lands within 3 years. Everything else is listed with `Apply = N` for a human. | `AuditRecordDates` |
| 8 | Same item + lot recorded with two dates that are each other's swap – which is right? | Cannot be decided by code; listed with `Apply = N` and both row numbers. (45 such rows today.) | `RowsWithDate` |
| 9 | Is the checking date trustworthy? | Yes when it is a real date (`Date` was assigned as a Date). 411 text values were typed by hand and are converted like column F. | audit, column A branch |
| 10 | Should the repair be reversible? | Yes: backup sheet `Record_bak_yyyymmdd_hhnn` + a note in Remarks per changed cell. | `ApplyDateAudit` |
| 11 | Should the form get separate Day / Month / Year controls or a calendar picker? | Not needed: strict parsing + preview gives the same guarantee with no `.frx` changes. A design-time label named `lbDatePreview` is used if you add one; otherwise the code creates it at run time to the right of the date box. | `SetupPreviewLabel` |
| 13 | Item code not in the ItemLocation list? | **Confirmed:** warn, save on Yes with a blank Item Name. | `btnAdd_Click` |
| 12 | Quantity: keep writing it as typed (text), as before? | Yes, unchanged. `Qty on hand` already mixes numbers and text such as `459x12's`; out of scope here. | `btnSubmit_Click` |

## 4. Is Excel the right place to hold the data?

**Decision (confirmed): Excel now, Access as a separate follow-up.**

Excel is not a database; a cell has no type, so anyone can type `N/A` or
`5/DEC/206` into the expiry column and the sheet cannot stop them. The
redesign above closes the *form* path; the *sheet* path is still open.

The other three tools in this repository already keep their records in
**Access (`.accdb`) on the NAS share** via ADO with parameterised inserts
(`Stocktake/VBA/Modules/DBCRUD.bas`, `Return Drug/VBA/Modules/modDatabase.bas`).
That is the natural next step for this log book:

* A `Date/Time` column in Access stores a serial, not text. Passing the date
  as an `adDate` parameter (as `DBCRUD.bas` does) means no string is ever
  interpreted.
* Concurrent users, a primary key, and a change log come for free.
* The workbook becomes a front end only; the printed report is a query.

XML or JSON files would still be a text format you parse yourself, with no
concurrency and no typing; they solve nothing here. IndexedDB is a browser
store and does not apply to VBA. If the log book later moves to a web app,
store ISO-8601 (`2026-11-01`) strings or epoch numbers and never a locale
string.

`DateSafe` is written so it can be reused unchanged when the sheet write in
`btnSubmit_Click` is replaced by an ADO insert.

## 5. Files

| File | Change |
|------|--------|
| `vba/DateSafe.bas` | **New.** Parsing, canonical text, preview, policy check, `WriteDateCell` / `WriteStampCell`. |
| `vba/DateRepair.bas` | **New.** `AuditRecordDates` and `ApplyDateAudit`. |
| `vba/frmEntry.frm` | Rewritten code-behind (form layout untouched). |
| `vba/General.bas` | `ChangeStatusToEnd` writes a real timestamp; `ClearRecordFilters`; `EnsureRecordDateFormats`; filters and sort qualified to the table. |
| `vba/frmReport.frm`, `vba/CustomLookup.bas`, `*.frx` | Unchanged. |

## 6. Importing into the workbook

1. Open the `.xlsm`, `Alt+F11`.
2. **Remove** the existing `General` module and the code of `frmEntry`
   (right-click → Remove; answer *No* to export).
3. `File → Import File…` for `DateSafe.bas`, `DateRepair.bas`, `General.bas`
   and `frmEntry.frm` (the `.frx` next to it is picked up automatically).
4. `Debug → Compile VBAProject`. It must compile clean; `Option Explicit`
   is on in every changed module.
5. Save.

Optional: in the form designer add a Label named `lbDatePreview` to the
right of `cbUBDate` if you want to control where the preview sits.

## 7. Repairing the existing data (once, by one person)

1. `Alt+F8` → `AuditRecordDates`. Nothing is changed; a `Date Audit` sheet
   is created. Expected on today's data:

   | Apply | Count | What |
   |-------|------:|------|
   | Y | 401 | Date of Checking text → date |
   | Y | 96  | Use Before Date text → date |
   | Y | 84  | Use Before Date before checking date, day/month swapped |
   | Y | 1 437 | Reference text timestamp → date-time |
   | N | 7   | Expiry before checking date but no plausible swap |
   | N | 45  | Same item + lot recorded with swapped dates elsewhere |
   | N | 12  | Unreadable text (`N/A`, `1/1/20220`, `5/112024`, `5/DEC/206`) |

2. Review the `N` rows. Type a date into **Proposed** and set **Apply** to
   `Y` where you know the answer; leave the rest.
3. `Alt+F8` → `ApplyDateAudit`. Confirms, backs up the sheet, writes the
   serials, notes each change in Remarks, and applies the display formats.
4. Check the backup against the table, then delete the backup sheet.

## 8. Verification checklist after import

* Type `11/1/2026` → rejected, message names both readings.
* Type `01112026` → box becomes `01-Nov-2026`, preview green.
* Type `20/6/2027` → accepted as 20-Jun-2027 (only one reading possible).
* Type `Nov-2026` → `01-Nov-2026`.
* Type `01012020` → warning "already in the past", *No* keeps the row out.
* Add two rows, Save → both rows show `01-Nov-2026` style with an English
  month even on a Chinese-language Excel, cells are numeric
  (`=ISNUMBER(F…)` is TRUE) on **both** a d/m/y and an m/d/y PC.
* Print report → Reference column shows `dd-mmm-yyyy hh:mm`, `=ISNUMBER` TRUE.
