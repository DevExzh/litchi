# Next measured implementation

Implement a private, fixed-field cell-attribute view in the eager worksheet
parser (`crates/litchi-xlsx/src/raw/worksheet/codec.rs:667`). The FP profile
places the shared attribute lookup at 15.38% of exact commit-ancestor weight,
and Heaptrack identifies repeated quick-xml duplicate-check allocations.
This is a measured candidate, not a promised speedup.

For prioritization only, eliminating the entire observed attribute-lookup
subtree would bound the sampled commit-context CPU speedup near
`1 / (1 - 0.153809) = 1.182`. A real one-scan implementation retains decoding
and validation, so it cannot realize that complete-elimination model. This
is an Amdahl model of the sampled method context, not an elapsed-latency or
end-to-end speedup forecast. The larger parsing/rewrite closure remains the
next residual target after a useful narrow change is verified.

Use one default checked `element.attributes()` scan. Recognize only unqualified
`r`, `s`, `cm`, `vm`, and `t`. Decode `r` immediately during scanning, then
finish the complete scan before parsing its coordinate, matching the current
first helper call. Retain the other `Attribute` values without decoding, and
consume them in the existing semantic order: `r`, `s`, `cm`, `vm`, `t`.
Decoding all fields eagerly would change error precedence. Preserve raw-name
duplicate checks for ignored/unknown and qualified attributes as well.

Keep the common public helper and the selected/source-backed parser unchanged.
Do not add a heap-backed name map, alter Store retention, skip output parsing,
or relax unknown/namespace handling. Fixed fields should remove four repeated
checked scans per cell without growing retention with document size.

Focused tests must cover each duplicate field; duplicate ignored/prefixed
attributes; prefixed `x:r` alongside unqualified `r`; unknown attributes;
normalized/entity-encoded `r` and `t`; coordinate/style/metadata bounds; and
error precedence such as invalid `r` with malformed/invalid `s`, and invalid
`s` with invalid `t`. Retain parser namespace/entity, malformed-input, commit,
patch, no-op, preservation and output-failure gates.

Establish clean source/binary-bound control and candidate builds, then measure
normal ABBA one-cell and one-percent commit/save on tiny, medium and dense-wide,
with operation allocation evidence or explicitly scoped Heaptrack attribution.
Keep the full current default matrix as a regression guard where practical.
Apply the approximately 5% latency/RSS review triggers individually; retain
the change only with useful measured benefit and preserved semantics. The
0466 source snapshots and normal source epoch supply the control provenance,
but its descriptive same-source captures do not substitute for this matched
optimization experiment.

The broader goal remains open: 22 representative mappings remain correctness-only,
native producer coverage is incomplete, and physical cold input, caller-supplied
nonzero-latency range sources, bounded streaming and worker scaling still require
scenario-specific evidence. This candidate does not redefine those requirements.
