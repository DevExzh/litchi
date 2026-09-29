# 0839 independent reader review

`readers.py` is the offline evidence boundary for the cached-Part scheduling
trial. The capture driver creates child processes and retains their report,
receipt, and acknowledged `/proc` snapshots; a later analyzer owns pairing,
candidate admission, and any decision. The reader does not import or execute
the probe and does not infer a candidate result from a report's self-authored
verification fields.

The payload oracle reproduces the probe's 32 members from the fixed
`litchi-0786-member-NN-` lane formula. It independently checks member index,
size, names, SHA-256 values, the ordered 32-member sequence hash, and logical
byte total for all three shapes. ZIP and CFB container hashes are retained as
identities and validated as hashes, while compressed byte totals remain
observable evidence for the paired analyzer rather than historical constants.

`validate_report(path, case, samples, warmup, feature=False)` checks the report
schema, frozen case/config, exact sample count, timing vectors, corpus identity,
resource limits, CPU-task lifecycle, and source-metric availability. For the
scheduled CFB and source-backed Parts routes, a feature build must report 64
logical source reads for a fresh sample and zero for a primed sample. Requested
and returned byte totals are checked for nonnegative, bounded relationships;
they are not pinned to a previous compressed ZIP layout. Every sample must
quiesce active reads, stay within worker width, and release worker and I/O
permits. The public operation's 32 CPU-task charge is checked as
`0 -> 32 -> 32` for fresh and `32 -> 64 -> 64` for primed; the borrowed
`opc.from_bytes` control is explicitly allowed its zero member-read increment.

The same interface accepts `allocation=True` for the isolated allocator
observer schema. It requires measured operation-scoped System allocator
records, zero failed allocations, exact live-byte conservation
(`after = before + allocated - deallocated`), monotone process high-water
marks, and a region peak bounded by entry/exit live bytes and the final
high-water mark. It returns the raw allocation rows plus vectors for allocation
calls, allocated bytes, region peak above entry, and retained live-byte delta.
No timing value is fabricated for this lane.

`nearest_rank`, `bootstrap`, `bootstrap_ratio`, and `paired_metric` preserve
the 0839 comparison contract: seeded median bootstrap, seed `839083`, 10,000
resamples, and endpoint indexes 249 and 9749. The raw vectors remain in the
returned summaries so the analyzer can make paired comparisons without losing
distribution evidence.

`parse_snapshot_rss` reads only raw `/proc/smaps`, `/proc/smaps_rollup`, and
`/proc/status` strings captured at an acknowledged phase boundary. It sums all
`smaps` `Rss` entries and separately parses the rollup `Rss` and status `VmRSS`.
It returns all three values, pairwise equality flags, and byte conversions so
the analyzer can account for known kernel/reporting discrepancies. The result
is a single whole-process resident snapshot and is explicitly labeled as not a
physical peak; no peak is inferred from it. `parse_snapshot_series` additionally
rejects a changing PID, parent PID, or process start time.

The module intentionally leaves cross-leg corpus equality, compressed-size
comparisons, native exit RSS, scheduling-candidate policy, and final adoption
to the root-owned capture/analyzer workflow.
