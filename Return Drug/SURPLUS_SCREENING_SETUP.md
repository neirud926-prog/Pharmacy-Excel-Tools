# Return Drug – Surplus Screening setup

The VBA code is done. The new on-screen controls live in the form's binary
(`frmApp.frx`), so they must be added in the VBA editor's form designer. Add the
controls below with the **exact (Name)** values, then import the updated code
modules. The project will compile once all names exist.

## 1. New panel menu button (place right after Home, before Data Entry)
On the left panel (`FrPannel`), copy the Home button group (Frame + Label +
Image) and rename the copies:

| Control | (Name)         | Notes                                   |
|---------|----------------|-----------------------------------------|
| Frame   | `btnFrSurplus` | the clickable panel tile                |
| Label   | `btnlbSurplus` | Caption e.g. "Surplus"                  |
| Image   | `IconSurplus`  | optional icon                           |

Drag the group so it sits **2nd, between Home and the Data Entry button**.

## 2. New MultiPage page
Right-click the `MutiPage` tabs → **New Page**. Then:

- **(Name) = `pgSurplus`**, Caption = `Surplus Screening`.
- Leave it as the **last** page – do **not** reorder the existing pages. The
  code finds this page by name, while the existing menu buttons still use their
  fixed page numbers.

On `pgSurplus` add:

| Control  | (Name)               | Notes                                            |
|----------|----------------------|--------------------------------------------------|
| TextBox  | `Surplus_tbScan`     | where staff scan (set the scanner to send Enter) |
| Frame    | `Surplus_frResult`   | the colored result box                           |
| Label    | `Surplus_lbStatus`   | big font – shows SHORTAGE / SURPLUS / NO DATA    |
| Label    | `Surplus_lbVariance` | shows the signed variance, e.g. `+25` / `-80`    |
| Label    | `Surplus_lbItem`     | shows `ITEMCODE - Description`                    |

Put the three labels **inside** `Surplus_frResult` and set each label's
`BackStyle` to **Transparent** so the frame's colour shows through.

## 3. Home page user guide
On the Home page (page 0) add:

| Control | (Name)        | Notes                                       |
|---------|---------------|---------------------------------------------|
| Label   | `Home_lbGuide`| `WordWrap = True`, large; text set by code  |

## Behaviour (already coded)
- Scan an OP code `ABCD01` or IP code `XXXXABCD01XXX` → the 6-char item code is
  extracted, looked up in `Stocktake.accdb` (latest `Record` row, any type).
- `variance = TotalSum − ERP`. **variance ≥ 20 → SURPLUS (green); variance < 20
  → SHORTAGE (red).** No record → grey "NO DATA".
- Result frame colour + the status word, the signed variance and the item
  code/description are shown in `Surplus_frResult`.

## Colours used (from `ConstColor`)
`green` = surplus, `red` = shortage, `grey` = no data.
