# ODS byte-position text performance profile

This profile compares the frozen byte-function candidate with baseline commit
`3844f235bac545ff0ae1580b97612883c1fd9f89` using the retained gate lock
`58b4be6cf88d7f7c5c2b16bd069a589e261e2a68e45a808a5cf3f12e1340a3e3` and the
reviewed contract hash
`88dafc6f0d111672e724b4238289afc0a17d879b856059ac3886ca60bbc99db5`.

The candidate matrix has 56 byte and computed-reducer workloads and 28 matched
controls. Every named case runs in `evaluate` and `parse-evaluate` phases. The
byte workloads cover all seven §6.7 functions across tiny scalar values, long
Unicode and ASCII values, borrowed 16x4 references, matrix broadcasting,
typed ReferenceList refusal, sticky cancellation, zero-cell resource refusal,
near-match searches, and bounded REPLACEB output growth. Three additional
projected lanes pin `AVERAGE(LENB(reference))`, `SUM(LENB(reference))`, and
`AVERAGE(LEN(reference))`; their expected coordinates and exact read bounds are
independent harness oracles.

The harness uses a borrowing resolver. It records elapsed time, allocator calls
and bytes, peak live bytes, execution work, retained budget, resolver reads,
input bytes, output bytes, and external peak RSS. Each selected case is
preflighted against an independent UTF-8 boundary model before timing. The
timed consumer validates result shape/type, consumes checksums and drops every
result so allocation and live-byte observations include the owned output path.

The final capture is deliberately source-frozen: candidate hashes are checked
against `gates/freeze.json`, the selected profile inputs are hashed before and
after both captures, and the source closure is checked for mutation. Three
warmups and fifteen fresh child processes per case and phase are the intended
final settings. Cancellation lanes use four repeats and must observe exactly
one successful cell read in every child. Failed or preliminary captures are
moved under a timestamped `results/diagnostic-*` directory before a retry.

The source-freeze handoff is complete. The final capture may now run against
the frozen candidate and `gates/freeze.json`; timing claims belong only in a
separate results report after `verify.py` checks both receipts.
