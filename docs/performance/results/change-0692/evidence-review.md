# Independent evidence review

Read-only review of the completed 0692 packet. I inspected the frozen
constraints, build and probe bindings; the native and allocation manifests and
raw TSV metadata; semantic bindings and comparisons; control, profile, trace,
integration, evidence, cleanup, and final-validation receipts; the probe
oracle; and the report tables and review triggers. No Cargo, perf, probe, or
production measurement was run for this review.

The evidence bindings are coherent. The packet retains 78 successful native
process receipts covering 13 cases and six legs, with 100 contiguous samples
and five warmups per process (7,800 samples total). Output and stderr hashes,
source archive hashes, target coordinates, workflow metadata, and correctness
metadata are stable across each case's legs. The 26 allocation receipts cover
the same cases with three samples and two warmups per phase. Baseline and
candidate binaries, probe inputs, production source hashes, profile files, and
the 103-member/43-replacement control are recorded and cross-bound; owned
scratch binaries and generated control input were removed under the retained
cleanup receipt.

The correctness oracles support the stated scope. One-edit runs require a
revision and semantic change plus both in-process and reopened target-marker
checks. Two-edit runs require distinct slides, both reopened markers, and
unchanged counts. No-op runs require `commit_is_changed=false`, unchanged
revision and semantic digest, and byte-identical before/after serialization.
The untouched digest covers sorted part identity, content type, relationships,
non-part metadata, and exact payload hashes while excluding only edited slide
payloads. The packet explicitly records that metadata ordering and edited-part
physical ZIP details are outside this oracle, and that no-op serialization is
compared before versus after rather than against the original archive.

The diagnostic trace is internally consistent: candidate capture/apply
intervals contain 31 successful calls with no errors and no skipped trace
runs; the real capture falls from the retained 0691 count of 44 to 31, with
819,319 to 550,141 input bytes and 944,179 to 634,121 owned output bytes.
Temporary instrumentation restoration is exact. Allocation comparisons retain
the unchanged real one-edit peak-above-start (463,159 bytes) and net-live
change (195,730 bytes), while recording lower call/request totals. Native
profiles are bound to the real source and the matching binary; the report
keeps the instruction/cycle reduction, page-fault increase, one-run RSS
observations, +24-byte scratch entry, and all 21 positive phase-tail triggers
visible.

The retained integration gate has seven successful entries, including format,
locked checks, Clippy, default and all-feature PPTX tests, the narrow PPTX
facade run, and rustdoc. The quality receipt records 908 default-test passes,
922 all-feature passes, 45 facade passes, and two existing ignored tests. The
six repository evidence gates, post-cleanup audit, structural/report/coverage
checks, and non-iWork check all pass. The facade warning-denial limitation and
the three archived broader-feature attempts are disclosed in the report; no
facade source was changed.

There is no evidence blocker for the packet's bounded conclusion. Results are
capture/edit phase observations on a shared warm Linux host pinned to CPU 12;
file read, initial open, target discovery, save, and reopen are outside the
timers. The packet makes no cold-cache, concurrent, cross-platform, native
Office, complete save workflow, general PPTX, iWork, or registered performance
claim. The allocation and trace artifacts remain diagnostics, and the 21
phase-tail triggers and profile limitations remain disclosed.
