# Existing-document append data path

The selector uses the ordinary owned ODP editing API. It deliberately includes
work that the older `odp_semantic_one_edit_save` interval excluded: opening
owned bytes as an editing snapshot.

`Snapshot::from_bytes` admits a package under the existing 128 MiB package
limit, retains shared archive bytes, opens a Presentation, and constructs the
complete slide projection. The snapshot checks slide count and aggregate draft
resources. `transaction()` retains the source authority and an isolated draft;
`add` stages one title/body slide. `commit` serializes the changed presentation
and validates the candidate before publishing a snapshot and reversible patch.
The runner then borrows the published snapshot bytes for one sequential
`write_all` to a hashing discard sink. The sink's hash update is timed; hash
finalization is not.

The source archive, expected output, and fixture oracles are prepared once
outside the measured loop. Each iteration clones the source Vec and constructs
append strings and the sink before the clock and allocator region. Opening,
transaction creation, append, commit, and output share one interval. The source
and committed snapshot remain live through the allocator endpoint. Their
retention makes nonzero live-byte deltas expected; it is neither a leak finding
nor a bounded-memory commit claim. Release occurs afterward.

The independent Rust gates reopen source and candidate and compare every
source slide against deterministic title/body expectations, preserve the whole
ordered source sequence, and require exactly one expected tail slide. Six ZIP
members, compression methods, exact manifest bindings and untouched decoded
bytes are checked. Opaque compressed data-span equality is observed separately;
it does not prove raw passthrough, unchanged ZIP framing, or avoided
recompression. The Python report oracle independently regenerates semantic,
order, text-projection, and opaque expectations. Reports do not embed complete
archives; Python does not independently reopen absent archive bytes.

Generic sink `input_bytes` is the full output text projection and
`authored_part_bytes` is the full committed content.xml size. They do not count
physical archive reads or just the one appended title/body argument. Source
read counts, decompression/recompression bytes, physical copies, cancellation,
remote ranges, cold-cache behavior, and native rendering are not measured by
this selector. Whole-process profiles also include corpus construction,
untimed oracles, and report serialization, so their symbols identify candidates
for later attribution rather than operation-only causes.
