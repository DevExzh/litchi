# 0844 matched missing-ancestor staging experiment

Base commit `1b6072fc32`. The candidate changes only
`crates/litchi-opc/src/members.rs`; `applied.patch` contains its formatted code
and three focused tests. The retained before/after source file copies and full
source hash inventories bind both legs. See `selection.md` for frozen guard
thresholds and `source-review.md` for the prefix-closure and atomicity review.

The independent probe has four path shapes, four Part counts and two timing
modes. Both binaries use the same frozen probe, root-aligned dependency versions,
compiler settings and owner symbol. Output artifacts and source archives must
match across both legs and modes for each shape/count. The reader independently
reconstructs payloads/CRCs, untouched local spans, and changed-target local and
central stable metadata. Thirty-four malformed-input mutations are rejected.

The final `measurement-plan.json` corrects leg order before any formal timing:
each case has three before-first and three after-first pairs. The original
build-bound `plan.json` remains immutable; `schedule-repair.json` records this
order-only correction, with unchanged cases, counts, and thresholds.

The formal schedule contains 64 qualifications, 384 native processes with
30 samples after 3 warmups, and 64 one-sample Heaptrack processes. Total:
512 reports / 11,648 samples. Two tiny Heaptrack compatibility preflights are outside
that total. Native and instrumented timings are never pooled. Heaptrack counts
allocation events / requested bytes whose full stack contains profile_publish;
lifetime frees need not occur inside the timed operation. Process RSS includes
fixture/oracle construction and verification, and is not operation peak live
memory. Every native case and both allocation repeats are retained.

Replay retained evidence after cleanup:

```sh
python3 -B docs/performance/results/change-0844/reader_tests.py
python3 -B docs/performance/results/change-0844/heap_reader_tests.py
python3 -B docs/performance/results/change-0844/close.py verify
```

Build/capture commands and environment are retained under `commands/`.
Reproduction requires a fresh isolated checkout/packet at the recorded base,
owned roots initialized as recorded by `origin.json` and marker witnesses,
the frozen probe/lock, a baseline `build.py before`, candidate application,
`build.py after`, all qualification/quality gates, and `admit.py` before native
and Heaptrack capture. Existing evidence paths are exclusive and must not be
overwritten. Build binaries are removed only after closure validates their
recorded identities; cleanup witnesses preserve those identities.

The draft `candidate-initial.patch` was never applied or built. It had invalid
handwritten patch context and per-key Vec growth. The final generated candidate
uses one exact reservation for the first missing ancestor and its suffix.
Intermediate unsafe map-reservation proposals were rejected in review; the
applied candidate retains the original final whole-batch map reservation.
The production baseline reuses the unchanged prior owner/workspace gate
inventory, with a fresh replay and baseline OPC tests. Candidate quality is
fresh. Full-workspace iWork example/rustdoc limitations remain explicitly
excluded; owner docs/tests/all-target checks and workspace library gates remain
required. The broad GOAL remains active regardless of this batch's decision.
