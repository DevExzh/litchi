# 0509 final evidence review

This is a bounded read-only review of the retained 0509 evidence. I ran
`python3 -B verify.py` in this evidence directory against the final reports
and profiles; it exited 0. I did not rebuild, rerun native or instrumented
captures, or edit Rust. `initial-summary.json` remains retained as the original
capture; the longer tail follow-up supplements it.

## Disposition

The evidence supports the scoped ODT paragraph-buffer reuse admission. I found
no concrete blocker for that bounded claim. The 4,096-byte value is a retained
spare-capacity bound, not a text-size or whole-process memory limit. The batch
does not establish a general latency, tail-latency, peak-memory, hardware
counter, or allocation-total improvement.

## Evidence consistency

The final verifier and `summary.json` agree on:

- 6,000 initial native samples plus 8,000 large-only follow-up samples, or
  **14,000 native samples** in total;
- 20 large-export Heaptrack samples per side and five large-export Callgrind
  samples per side, or **50 instrumented samples** in total;
- 200,000 to 20 `append_sink_precharged` allocation calls across 20 exports,
  one retained allocation per export after reuse;
- a 4,096-byte actual `String::capacity()` retention cap; and
- six matched native rows, paired corpus/output/sink identities, drift checks,
  and rejection of the deliberately short sample vector.

The profile summary reports 304,455,496 versus 298,891,534 inclusive Callgrind
references (-1.83%) and 9.91M rounded whole-child peak heap on both sides.
Those are respectively simulated parser-scoped references and rounded
whole-child profiler output. The denied hardware probe leaves no hardware
counter or cache claim available.

## Tail disclosure

The initial large R1 pair has a **+9.1402% p99** adverse latency flag; the
initial large R2 p99 change is +1.3751%. The post-hoc large-only follow-up
retains the original capture and reports p99 changes of **+2.1788%** and
**+2.7775%** for tail-R1 and tail-R2, respectively. Neither follow-up pair
crosses the 5% review threshold. Candidate follow-up median drift reaches
**4.37%**, below the recorded 5% drift ceiling but larger than the observed
median gain. The follow-up is therefore supporting evidence, not proof that
the initial tail flag was noise; the original flag remains disclosed.

Whole-child RSS changes are +0.34% and +0.77% in the initial repeats, with
-0.41% and 0.00% in the follow-up. These children include setup, opening and
verification and do not support an operation-local RSS or peak-memory claim.

## Source and semantic scope

The source manifest and change record identify exactly two production/test
source files: `crates/litchi-odt/src/elements/text.rs` and
`crates/litchi-odt/tests/sequential_text.rs`; the performance harness is
unchanged. Reuse remains operation-local, follows a successful ordered writer
publication, and drops values whose actual capacity exceeds 4,096 bytes.
Writer failures return before recycling, preserving frontier order, accepted
bytes, completed-object counts, typed errors and source-staleness behavior.

The unchanged ODT timing contract opens and validates before the export clock
and uses the bounded hashing-discard sink with digest updates inside timing and
verification outside timing. The profile and native comparison therefore use
matched sink boundaries, while the instrumented profile remains excluded from
native latency claims.

The source review identifies only modest fixture gaps: no dedicated multi-block
64 MiB decoded-text-ceiling case and no standalone exactly-4,096-byte fixture.
Static invariants and adjacent tests cover the relevant behavior; these gaps
are not blockers for the admitted bounded reuse claim.

## Validation and remaining scope

The recorded gates show **1,491 passing Rust tests/doctests plus one ignored**
existing opt-in test: 1,008 focused ODT tests/doctests and 483 unchanged
harness tests. Formatting, Clippy, rustdoc, boundaries and strict claim
classification passed in the retained evidence.

This admission remains limited to ODT semantic text export and the filtered
`append_sink_precharged` allocation stack. It does not close ODS joined-output,
provider scheduling, native-producer, or broader non-iWork coverage. Root still
owns final cleanup and commit sealing after this review.
