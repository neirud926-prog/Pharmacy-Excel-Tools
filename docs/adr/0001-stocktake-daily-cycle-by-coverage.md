# Stocktake: the next Daily bin comes from cycle coverage, not the latest record

The Daily Stock Take used to pick the bin after the latest saved Daily record. One off-sequence entry could therefore move the rotation. On 04-Sep-2026 a batch that ended at 2A8A sent it back to 2A8B, so the 1*, P*, R*, S*, T* and MF*+ bins were never counted. We now replay every Daily record in (RecDate, ID) order against BinShelfLocator. A Daily Cycle completes only when every item has been counted, and the Next Items are the uncounted items of the current cycle, in bin order. The state is derived, so the Access database needs no new table or column, and no stored pointer can drift.

## Considered Options

- **Stored pointer, advanced on print**: rejected. It needs a schema change, and bins on a printed sheet that is never counted would be skipped silently.
- **Least-recently-counted first**: rejected. Pharmacy staff want an explicit rule that a full cycle is completed and verified before the next one starts.

## Consequences

- Annual records do not count towards the Daily Cycle.
- An item added to BinShelfLocator mid-cycle has to be counted before that cycle can complete. Items removed from BinShelfLocator are ignored.
