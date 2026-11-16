# Pharmacy Excel Tools

Shared language for the pharmacy's Excel/VBA tools. The Stocktake terms were settled while investigating the BinShelf rotation issue (Nov 2026).

## Stocktake

**BinShelf**:
The physical shelf location code of an item, e.g. `5C4D` or `TD9A`. The bin order in BinShelfLocator is the walking order of the Daily Stock Take.
_Avoid_: Bin No, location, shelf

**BinShelfLocator**:
The master list of items covered by the Daily Stock Take (manual-dispensing items, excluding C*, D* and R5* bins, which are not stock-takable), sorted by BinShelf.
_Avoid_: Locator list, bin list

**Daily Stock Take**:
The rolling count of a few BinShelfLocator items each day (records with Type `Daily`).
_Avoid_: Daily check, cycle count

**Annual Stock Take**:
The once-a-year full count (records with Type `Annual`). It has its own Annual Cycle and never counts towards the Daily Cycle.

**Annual Year**:
The window 1 June to 31 May of the next year. Every Annual record dated in it belongs to that year's Annual Stock Take.
_Avoid_: Financial year, calendar year

**Annual Cycle**:
The Annual Stock Take of one Annual Year. It is complete when every BinShelfLocator item has an Annual count dated in that Annual Year. Daily counts never count towards it.

**Daily Cycle**:
One complete pass of the Daily Stock Take in which every BinShelfLocator item has at least one Daily count. The next cycle starts only after the current one is complete.
_Avoid_: Round, rotation (in writing)

**Next Items**:
The items not yet counted in the current Daily Cycle, taken in BinShelfLocator order. "Print 5 items" prints the first five of them.
_Avoid_: Next bin (when it means "the bin after the last record")

**Off-sequence entry**:
A Daily count of an item that is not one of the Next Items, e.g. a spot check. It counts the item as done in the cycle but never moves the sequence.
_Avoid_: Ad-hoc record, random entry
