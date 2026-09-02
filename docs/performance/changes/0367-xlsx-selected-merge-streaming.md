# Change 0367: XLSX selected merge streaming

Change 0367 extends the selected-worksheet scanner to support valid direct
active `mergeCells` content globally through verified worksheet EOF. Merge
validation is exact: the declared count, nonempty `ref`, reference grid,
singleton rejection, placement and direct-child rules, and overlap rules are
all enforced before publication. Eligible scans build the canonical transient
`merge::Index`.

After all worksheet, dependency, ZIP, source, and execution fences complete,
a selected single-cell non-anchor returns `Covered`; anchors retain the
existing `Stored`/`Missing` outcomes. Range `cells` and `visit` remain sparse
physical records, including merge followers, with no synthetic covered cells,
and `stored_extent` is unchanged.

The eligible range path retains at most 16,384 merge ranges using
`try_reserve`. A 16,385th range drains through verified EOF and then takes the
mandatory eager fallback. Unknown merge attributes, children, or payload also
fall back after the required drain; malformed merge structure is a hard typed
error. Eligible cold paths publish no `Store`, `PartData`, or semantic caches.
The transient index's internal `BTreeMap` and heap allocations are bounded by
the cap but are not individually fallible, so this change makes no
fixed-memory, RSS, or OOM claim.

Focused validation passed `14/14`, full `litchi-xlsx` library validation passed
`906/906`, and scoped Clippy passed with only the unrelated
`clippy::useless-asref` issue allowed. `performance_claim: none`; no latency
claim follows.
