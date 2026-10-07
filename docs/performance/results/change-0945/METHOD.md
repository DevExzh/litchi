# 0945 XLSX source payload sharing protocol

This packet measures the frozen source backed XLSX publication candidate at
authoritative source revision `e037befa6ce52b3988c8ae2032ae652196f2a09d`.
The candidate source manifest contains exactly two production changes:
`crates/litchi-opc/src/source_backed.rs` and
`crates/litchi-xlsx/src/cell_values/snapshot.rs`.  The packet has no DOCX or
iWork workload.

The four fixed rows are medium and dense-sparse one-percent cell edits,
unmanaged one-edit medium, and managed one-edit medium. Qualification has 16
children: native and allocator lanes for both roles and all four rows, with
zero warmups and two samples. Formal native capture has six alternating blocks,
ten warmups, and 100 samples per child. Formal allocator capture has two
alternating repeats, two warmups, and three samples per child; the odd repeat
reverses role and case order. Every child is
serialized on CPU 12 under `PATH=/usr/bin:/bin`, `LANG=C`, `LC_ALL=C`,
`TZ=UTC`, and `RUST_BACKTRACE=0`; `/usr/bin/time` records maximum RSS in KiB.

Before timing reduction, each report must pass the source, corpus, sink,
output, reopen, untouched-member, and phase-alignment checks. Managed controls
must prove budget limits, pre/post usage, reservation release, typed output
refusal, and source identity. Allocator reports must prove nonnegative
counters, zero failed calls, call/reallocation ordering, live-byte
conservation, and region/high-water bounds. Native latency safety covers whole
workflow and publication p50 for all four rows: every paired median is at most
1.05 and the deterministic 10,000-draw two-sided 95% bootstrap upper endpoint
is below 1.05 (seed `945009`, endpoints 2.5 and 97.5 percentiles).

The allocation gate is deliberately stricter for the two primary rows:
candidate publication `allocated_bytes` must be at most 0.95 times baseline in
each of the two allocator pairs. Allocation calls, region peak live bytes,
live bytes after publication, and whole-process RSS use a 1.05 paired-median
safety threshold on every row; individual pair regressions remain review
flags. A measured zero is valid. Equal zero pairs are recorded as
`degenerate_equal_zero`; a nonzero candidate over a zero baseline is
non-comparable and fails closed.

`tools/capture_0945.py` records terminal and process-group receipts before
parsing child output and never retries a child. `tools/reduce_0945.py` is an
offline reducer; `tools/freeze_0945.py` writes the plan only after source and
binary/build descriptors exist. Large binaries, reports, and logs remain
outside the repository.

The baseline native binary is reused only through the versioned
`receipts/baseline-native-reuse-v2.json` receipt. Its packet-local references
contain the copied 0944 start, terminal, source manifest, closure selection,
and export receipts; `tools/reuse_v2.py` compares all 10,788 closure entries
by path, bytes, SHA-256, Git object, and mode before accepting the binary.
The final candidate binding defaults to `candidate-source-v2`,
`candidate-source-manifest-v2.json`, and `receipts/candidate-freeze-v2.json`.
Build custody is role-bound: baseline native reuse and baseline allocator use
the v1 build namespace (`baseline1`), while both candidate v2 wrappers use the
v2 namespace (`candidate2`) for targets, temporary directories, logs, and
terminal schemas.  The frozen plan repeats these revision and binding IDs so
the shared validator cannot silently accept a candidate v1 receipt.
Artifact descriptors stored in the packet are relative and fail closed on
absolute or symlink paths; historical absolute command and working-directory
strings remain in terminal custody records for replay verification.
