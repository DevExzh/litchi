# 0821 execution notes

The previous goal turn made progress: commit `8312aaa29b` repairs the invalid
process-wide allocator test assertion and retains six passing quality gates.
Root rechecked history, the three unrelated dirty paths, all 35 previously read
normative input hashes, and the complete 0820 seal before creating this packet.

This batch resumes the previously unexecuted 24-case durability matrix on the
repaired base. It changes no production or runtime harness source. Quality is
reused only through exact source/input and committed evidence verification;
release builds, artifact admission, qualification, and measurements are fresh.
Root owns all Cargo, exporter, workload, and profiler execution and retains
process handles to terminal. Agents prepare and statically review the protocol
and readers; heavy offline replay waits until all captures are terminal.

Prior packets are immutable. iWork and unrelated workspace changes are excluded.

## Admission ordering failure retained

Root invoked artifact admission before the standalone ZIP-preservation witness
producer. Session 8294 exited 1 with `FileNotFoundError: zip-preservation.json`
after the independent artifact auditor had exited zero. `admission-0` retains
that successful child audit and receipt; the parent invocation failed and did
not publish admission. Root then runs preservation and a new admission attempt
on the same fresh exports. No driver, output, or frozen input is changed, and
no timing capture started before acceptance. See `execution-errors.json`.

## Terminal execution

All three fresh release builds and the artifact exporter exited zero.
ZIP preservation and artifact admission pass; qualification admission passes
for all 24 cases. All capture children completed successfully: 24 qualification
reports/24 samples, 144 native reports/4,320 samples, and 48 observer reports/144
samples. Root retained each handle to terminal and only then authorized offline
analysis. No source or frozen input changed during execution.

## Offline replay and independent audit

The first reader attempt rejected the prior seal's numeric file count versus
its file map. The reader was corrected to compare map length; attempt 1 and
replay pass for 216 reports and 4,488 samples. Aggregate validation initially
looked for an extra `admission` wrapper under `admissions`; root corrected the
lookup to the actual flattened schema. Both failed outputs are retained.
Aggregate validation now passes. The independent root audit matches all 144
native reports, quantiles/spreads/tails, and 24 absolute/paired bootstrap
intervals. No frozen driver, receipt, output, or raw timing sample changed.

## Final cleanup and replay

Independent results review passes. Root verified the three release binary
hashes, removed target and scratch (1,808 files / 3,767,474,281 logical bytes),
and ran `validate.py --final` to exit 0 with cleanup checked. Derived analysis
bytes remain identical after deletion. Final sealing and exact commit checks
follow; only the three pre-existing unrelated workspace paths remain excluded.
