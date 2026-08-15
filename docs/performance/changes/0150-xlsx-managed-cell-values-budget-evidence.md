# Change 0150: managed XLSX cell-value budget-boundary evidence

Date: 2026-08-16

The scalar-cell CRUD harness now records the managed source-backed budget
dimensions around publication without adding selectors. This evidence-only
harness delta is separate from the production editor adoption recorded in
[Change 0151](0151-xlsx-managed-source-editors.md); it does not itself alter
the scalar-cell production mechanism. The existing four managed controls
remain opt-in:

- one existing scalar cell;
- the deterministic `ceil(1%)` existing-cell set;
- the exact 256-cell batch; and
- one existing cell on each of two worksheets.

Each managed iteration retains the pre-publication `SourceCacheDiagnostics`
values for cumulative `InputBytes`, `OutputBytes`, declared `Work`, and
retained `Objects`, including their limits and catalog/cache object
reservations. The managed context now gives `OutputBytes` a finite limit based
on the successful XLSX output ceiling already used by the bounded sink. An
immediate post-publication shared-budget snapshot records the cumulative
output charge and the other resource dimensions after the consuming publisher
has released its package catalog/cache reservations. Those catalog/cache
reservation counters are no longer observable after package consumption and
are serialized as `null`, rather than being reported as observed zero. Live
commit/snapshot handles, if any, remain represented in `Objects` usage. A
second sample after all retained handles drop must report both `Memory` and
`Objects` usage as zero. Declared `Work` is reported as its own budget
dimension; it is not presented as decompressed or source-read bytes.

For each managed selector, one untimed replay first discovers the initial
bounded sequential publication request, then repeats the same source-backed
transaction with a one-byte-under `OutputBytes` limit for that request. The
replay must return the typed `Resource::OutputBytes` refusal before accepting
any sink byte, preserve the immutable source version, and leave cumulative
`OutputBytes` at zero. This boundary replay is evidence of refusal and source
identity only; it is outside every elapsed sample.

The existing correctness gates remain unchanged: deterministic generated
corpora, semantic cell readback, exact output/source hashes, untouched-member
identity, bounded sink writes, and release-to-zero checks for both `Memory` and
`Objects`. Unmanaged controls explicitly serialize zero/false values for every
observable budget dimension, limit, usage, pre-publication reservation, and
output-refusal field. Post-publication catalog/cache reservation fields are
`null` for both modes because the consuming publisher has released the package
and its cache diagnostics are no longer observable; shared-Budget object usage
is sampled directly instead. The default matrix remains 36 cases and 198
records, and the selectable case count remains 291.
No release ABBA comparison or performance result is claimed. The evidence
does not claim allocation counts, RSS/peak memory, hardware/CPU pinning, cold
I/O, decompression, copied bytes, or real-producer breadth. The managed Budget
still excludes parsed cell stores, relationship/graph metadata, staging,
rewritten candidates, and output-buffer allocation.

Validation performed for this change:

- focused deterministic/bounded harness coverage for the matched scalar-cell
  controls, including the managed output refusal;
- strict `cargo clippy --all-targets -- -D warnings`, formatting, and
  `RUSTFLAGS='-D deprecated' cargo check --all-targets` checks; and
- the existing tiny XLSX smoke coverage, with no selector-count changes.
