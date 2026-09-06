# Validation and evidence limits

The production candidate passes the full OPC release suite: 458 tests, zero
failed, one existing ignored. The standalone harness passes 373 tests, zero
failed, one existing ignored. Tests run serially, Cargo jobs are capped at four,
Rust is pinned to 1.98.1 and incremental compilation is disabled. No new test
merely duplicates the one-line ownership transfer. Existing content-type/URI,
XML normalization, resource limits, malformed input, duplicate/equivalent names,
signatures, ZIP64, source freshness, short-read and partial-sink gates remain.

Owner Clippy and owner/harness rustdoc pass with warnings denied. Workspace
no-iWork and standalone all-feature/all-target checks pass, as do explicit-file
rustfmt and crate boundaries. The unchanged standalone harness has inherited
strict Clippy debt; the retained comparison must admit zero new diagnostics.
Its baseline receipt/log/driver are copied from sealed 0445, with an exact source
manifest match to the before build. This is not a clean harness strict-lint claim.

Six deterministic ZIP fixtures and their metadata are copied unchanged from
0445. The independent oracle checks payload formulas, content types, relationships,
raw local/central records, offset allowances, ordering and comments. It checks
reports against those artifacts for both builds using the plain-source role.
The initial six baseline pilots and 29 corruption probes pass. Producer refusal
gates remain Rust checks; Python is not claimed to independently replay those
failures. Plain source counters remain unavailable, not measured zeros.

The first freeze attempt failed because data-path.md had not been written.
The attempted build descriptor and A1 capture then failed on missing protocol.
That capture failed at preflight: no production edit or workload ran and no
retained sample existed. The failure receipt/log remain. After writing the
reviewed data path, the actual timestamped protocol was frozen, the descriptor
was bound, and the retained A1 block used the unchanged baseline source. No
threshold or capture-driver change followed the freeze.

The production delta is limited to one string ownership handoff. There is no
new public API, dependency, unsafe code, executor or ambient production I/O.
Accepted ADRs and user-owned GOAL.md remain unchanged. The registry stays at
438 selectors/36 defaults; the semantic representative index remains 15
categories/33 mappings/10 measured/23 correctness-only. Native breadth,
semantic owner creation, repackaging, bounded existing append, cold/range and
scaling obligations remain open. The full non-iWork goal remains active.

A later, stronger per-sample byte-saving assertion rejected its own proposed
`26*(3*N+1)` formula. Existing Part names have 26 bytes, while the newly added
name has 25: the exact saving is `78*N+25`. The failed verifier draft/receipt
remain retained. Raw observations and the frozen allocation-call gate did not
change. The corrected assertion checks every sample against the derived vectors.

Both builds' six pilots pass (12 total), as do all 24 retained reports and four
profile reports. Strict harness comparison finds 29 rendered diagnostics before
and after, zero new. Both perf stack exports contain no unparsed lines; retained
record logs contain no lost-sample or addr2line warning. The allocation gate
passes; the normal latency gate fails. No paired latency/RSS regression or
repeat drift crosses 5%. Every absolute paired trigger is an allocation-call
improvement; peaks remain unchanged and endpoint live deltas remain zero.

Portable sealed export verification passes before and after cleanup. Five
pre-cleanup and six post-cleanup tamper probes reject, for 40 corruption probes
including the 29 archive/report mutations. Cleanup removes only
`/tmp/litchi-goal-0446-binaries` (1,833,577,768 regular-file bytes), preserving both
Cargo target directory identities and GOAL.md's pinned hash. The portable
verifier replays derivations and independent report oracles without copied
executables, Git, Cargo, perf or workload execution.
