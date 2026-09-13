# 0548 OLE2 checkpoint validation experiment

The previous goal turn made progress: sealed 0547 profiles attribute 20–22% of
large-constructor instruction work to per-step visited-map validation. This
batch tests the reviewed bounded checkpoint/terminal-proof candidate, retaining
initial reservations and zero-fill plus authoritative error replay. OLE2 and
OOXML remain the priority; ODF is deferred until completion and iWork excluded.

The committed baseline is frozen with one common public CFB guard example.
The production candidate is isolated until baseline captures terminate and
source review admits its exact patch. Root runs all builds, native captures,
allocator runs, profiles, hardware counters, guards and quality checks serially.
No source changes occur while any such process is live. Capture inputs and gates
are frozen before the first build. Analysis-only changes are explicitly recorded
and never alter source, executable, raw samples, thresholds or capture order.

Main native measurements use CPU 2, 20 warmups and 1,000 samples, two repeats
in baseline/candidate/candidate/baseline order. All four primary XLS p50 cases
must improve by at least 3% in both repeats. The baseline repeat-2 control uses
its retained executable while candidate source is active; receipts record both
identities. Matched allocations use 3 warmups and 30 samples; constructor Ir
and collector self Ir are independently scoped to positive incoming owner
ancestry over five timed dumps per job. CFB setup and termination dumps remain
separate. Whole-child hardware/RSS are diagnostics, not operation-local claims.

The public malformed guard builds before the large baseline executable, then
captures eight cases at 128 and 16,384 sectors, 20 warmups and 200 samples per
child. Two-repeat guard order is also A/B/B/A. Timing stops immediately after
OleFile::open; oracle comparisons and returned-value destruction are outside
the clock. Rejection cleanup inside open is naturally included. Inputs and
exact errors must match across stages/repeats. Candidate-invalid p50 and mean
must each remain within 4x same-invalid baseline and 2x baseline-valid at the
same size/repeat. Every >5% adverse metric and absolute >5% repeat variation is
reviewed; none is erased by these admission envelopes. The guard has no
allocator instrumentation; ordinary allocation gates remain independent.

Retain only when native, allocation, profile, guard and correctness/quality
gates pass. Otherwise restore the exact baseline production file and keep the
measured enabler and evidence. Fifteen final quality commands are declared in
checks.py. Existing parser fuzz setup is inspected: cargo-fuzz and nightly are
unavailable, so no sanitizer campaign is claimed; exhaustive Rust differential
and public malformed guards remain required. Prior 0546's 1,306-test suite is
the XLSX owner suite, not a whole-workspace test count. This batch reports the
actual OLE owner test commands and counts separately from workspace checking.

Before cleanup, verify all retained executable/source hashes, replay reports,
check every adverse row and inspect measured assembly. Remove the sole owned
target only after all processes terminate and accessible process references
are absent. Seal complete raw evidence, stage exact reviewed bytes, commit,
then repeat strict verification with a clean worktree. The broad goal remains
open regardless of this candidate's disposition.
