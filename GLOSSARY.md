# Pharmacy Excel Tools

Shared language for the pharmacy's Excel/VBA tools. The Stocktake terms were settled while investigating the BinShelf rotation issue (Nov 2026).

## Stocktake

**BinShelf**:
The physical shelf location code of an item, e.g. `5C4D` or `TD9A`. The bin order in BinShelfLocator is the walking order of the Daily Stock Take.
_Avoid_: Bin No, location, shelf

**BinShelfLocator**:
The master list of items covered by the Daily Stock Take (manual-dispensing items, excluding C* and D* bins), sorted by BinShelf.
_Avoid_: Locator list, bin list

**Daily Stock Take**:
The rolling count of a few BinShelfLocator items each day (records with Type `Daily`).
_Avoid_: Daily check, cycle count

**Annual Stock Take**:
The once-a-year full count (records with Type `Annual`). It does not count towards the Daily Cycle.

**Daily Cycle**:
One complete pass of the Daily Stock Take in which every BinShelfLocator item has at least one Daily count. The next cycle starts only after the current one is complete.
_Avoid_: Round, rotation (in writing)

**Next Items**:
The items not yet counted in the current Daily Cycle, taken in BinShelfLocator order. "Print 5 items" prints the first five of them.
_Avoid_: Next bin (when it means "the bin after the last record")

**Off-sequence entry**:
A Daily count of an item that is not one of the Next Items, e.g. a spot check. It counts the item as done in the cycle but never moves the sequence.
_Avoid_: Ad-hoc record, random entry
