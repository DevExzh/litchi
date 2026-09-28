# 0822 execution notes

The previous goal turn made progress: commits e540859d7a and 353aa00a7d
retain the admitted real-file save durability comparison and corrected
post-commit replay. Root rechecked HEAD, the three unrelated workspace paths,
all 35 previously read normative input hashes, and all 774 sealed 0821 paths
before starting this batch. No source change is planned.

This batch profiles the real-file PPTX public edit transaction rather than
changing synchronization guarantees. Opening the path, serialization, digest
checks, semantic readback, and owner drop remain outside the measured edit.
The input and expected output are pinned to the admitted 0821 real-002-pptx
row. The standalone probe is diagnostic evidence, not the ordinary-save harness
or a production latency improvement.

Root owns every Cargo, binary, native capture, profiler, and decoder invocation.
Builds and workload lanes run serially, with process handles retained to terminal.
Agents prepare source, drivers, and offline readers; heavy offline replay waits
for all workload captures and decoding to finish. Prior packets are immutable.
The three unrelated paths and iWork remain outside this batch.

After the workspace-update note, root rechecked the recent commits and
working tree: HEAD remained 353aa00a7d, the 35 normative hashes and three
unrelated hashes remained exact, and only the 0822 draft packet was new.
Static pre-build review identified probe argument/path/bounds and quality
receipt-schema issues; these are corrected before freeze, not workload
failures. No measurements have been executed at this point.

The read-only production-quality reuse preflight passed (641 passed, zero
failed, one ignored across 28 summaries). A subsequent read-only custody
preflight stopped before any persistent freeze or Cargo invocation because
corpus metadata used relative paths while its checker expected absolute
paths. The driver owner was asked to normalize the metadata/check boundary.

The corrected read-only custody preflight passed: 9,197 production files,
87 tool files, 35 normative inputs, and 773 prior payloads plus their seal.
Root finalized the decoder archive cleanup and explicit frame-pointer/call
assembly gate. All five execution drivers are frozen by the fresh quality
receipt; offline readers and execution notes remain mutable until final seal.

The fresh five-gate probe quality run completed successfully: formatting,
offline locked release check, all three focused tests, warning-denied Clippy,
and warning-denied rustdoc. Its root process reached exit zero before release
builds started. Production quality is explicitly reused, not rerun.

All build, qualification, native, perf, and decoder processes reached exit
zero before offline readers started. The independent native audit passed.
The first independent frame audit rejected a legitimate empty sampled stack.
Its exact script and failure log are retained; the corrected audit counts
empty stacks explicitly in the whole-process/outside-owner denominator.
No capture, raw profile, frozen input, or measured sample changed.

The analysis agent completed --write, --check, and nonfinal validation before
its handoff reached root. Root then attempted --write and hit the intended
no-overwrite guard; reader-attempt-0.log retains that duplicate invocation.
The successful derived outputs were left intact, and root switched to replay.
Root strengthened the terminal validator to compare the entire independent
frame leaf partition and to rerun both raw audits in check mode.

Root's strengthened validator passed before cleanup. The independent results
review accepted the measurements and requested wording clarifications, which
were applied to the main report. Root then verified both exact binary hashes
and removed only `/home/zhuhe/code/litchi-target-0822`: 3,684 files and
2,054,019,575 logical bytes. Final validation passed after removal, including
both independent raw audits and stack-depth diagnostics. The main report's
pending marker was removed only after that terminal gate. The three unrelated
workspace paths remain unchanged; iWork is excluded. The non-iWork program
continues with a bounded source review, not an adopted optimization.

Before commit, root explicitly audited the analysis agent's earlier failures.
The agent reported six failed --write attempts during reader-schema alignment:
input contract, previous-seal witness shape, plan hash shape, missing label
argument, symbol descriptor path, and compressed original-identity shape.
The agent did not retain their original console logs or intermediate reader
sources. `reader-failure-attestation.json` records the reported error strings
and order as a retrospective attestation, without inventing timestamps or
exit codes. This is a logging limitation. The subsequent successful output
and root's complete raw-data replay are independently retained; captures and
frozen drivers were unchanged. Root's duplicate write refusal remains a
separate actual log. All final audits and validation pass.
