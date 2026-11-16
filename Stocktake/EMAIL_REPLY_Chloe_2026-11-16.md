Subject: RE: BinShelf "1" / "T" never stock taken – investigation result and fix

Dear Chloe,

Thank you for sending the record. We checked it against the bin list, and you are right. The full result is below.

**What we found**
- Since 4 May 2026, 216 of the 705 items in the Daily Stock Take list have not had a Daily count. These are all the 1, P, R, S and T bins, the M bins from MF2A onwards, and a few in 2A (2A1B–2A5A, plus 6B9B).
- Cause: the "next bin" function took the bin **after the last saved record**. On 4 Sep 2026 a Daily entry for other locations was saved, and its last row was 2A8A. This happened just as the rotation reached the end of the M bins, so the sequence jumped back to 2A8B and never reached the P, R, S, T and 1 bins. A similar jump on 5 Jun 2026 made the 3A–3B bins get counted twice, and 5B7A–5C1A were also repeated on 10 and 12 Nov.

**What we changed**
- The next bins now come from the **current cycle**: the system lists the items not yet counted in this cycle, in bin order. Entries for other locations still count those items as done, but they can no longer change or skip the sequence.
- A new cycle starts only after **every bin** has been counted. You will get a message when a cycle is complete.
- "Print 5 items" now prints the next 5 items not yet counted, so items already counted in this cycle are not printed again.

**What to expect**
- The next sheet will start at 1A2B (1A2B, 1A3A, 1A3B, 1A4A, 1A4B). The missed 1, P, R, S and T bins will be covered over about the next 43 sessions, and then the cycle will finish.
- The list of missed items is attached (Daily_Not_Counted_Since_2026-05-04.csv) for your reference.

We will install the updated tool on the pharmacy PC and check it there before you use it. We will let you know once it is ready.

Kind regards,
Daniel
Tech Support
