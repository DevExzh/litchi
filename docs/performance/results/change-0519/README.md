# 0519: reuse the OPC publication XML proof

Base revision: `45d71cb6f0cd5d003c544b61c748552a513b0245`.
The active goal remains OLE2/OOXML optimization. ODF is deferred until that
goal is complete; iWork is excluded. This batch does not close the program's
broader coverage or historical 0499/0500 limitations.

## Source contract

`SourceXmlPart::check_for_publication` reuses the complete XML validation
performed at capture or splice finish when the full destination `ReadLimits`
equal the proof's limits. Content type, lineage, source/context, destination
PartBytes, replacement identity/original bytes, and final source fences remain.
Each publication still consumes one payload-length Work charge. Different
limits retain the complete destination validator. Equal limits no longer need
the unused parser-memory reservation, allowing success with less transient
memory. DOCX candidate semantic reparse/readback and package security, topology,
ZIP preservation, and output checks are unchanged.

`source-review.md` inventories every proof constructor and payload mutation;
`adr-review.md` maps the accepted contracts. The three-file patch in
`candidate/source.patch` replays exactly against the base revision.

## Reproduction

Recreate only the two owned temporary directories in `start.json`. Build the
base with `run.py`, then serially run `capture.py` lanes `preflight`, `r1`, `r2`,
`profile-preflight`, `profile-r1`, `profile-r2`, and `hardware`. The plan fixes
24 native cases, 30 samples, three warmups and two internal repeats, plus six
publication-profile cases with one sample, no warmups and one repeat. Second
campaigns reverse the case order. The example source is unchanged.

Build and capture the standalone allocation baseline with `allocations.py
build baseline`, followed by `allocations.py capture baseline --repeat 1` and
`--repeat 2`. After applying the source patch, build with `run.py --variant
candidate`, run the corresponding `after-r1`, `after-r2`, `profile-after-r1`,
`profile-after-r2`, and `hardware-after` lanes, then repeat the allocator steps
with `candidate`. Capture files are created exclusively; use a fresh evidence
directory for a new experiment.

Native publication timing includes destruction of the returned Snapshot.
Callgrind and allocator regions end at method return, before that destruction.
Hardware counters and RSS cover the whole child, including setup and oracles;
they are diagnostic and cannot attribute costs to the publication method.
Allocator comparisons use the same separately instrumented standalone build,
whose dependency resolution and default release profile differ from the native
workspace. Absolute allocator counts are not native production-binary counts.

`tail_guard.py` defines a separate interleaved follow-up for the two original
end-to-end p99 flags. Its results supplement the original campaigns; they never
replace or delete those flags. Baseline captures in that follow-up execute the
retained, hash-bound baseline binary while separately recording the unchanged
current candidate source tree.

## Audit

`checks.py` records nine source-bound correctness and policy gates. Eight
focused proof tests passed in `focused-check/`; the full all-features format
suite and independent ZIP64 case are separately retained under `checks/`.
`replay_reports.py` recomputes five JSON/Markdown report pairs exactly.
`verify.py --require-complete` validates custody, oracles, scope, counters,
reports, quality gates, and the `SHA256SUMS` leaf inventory. `verify_test.py`
checks that tampered evidence is rejected. Final verification and cleanup
receipts record removal of the owned scratch and target directories.

The per-change report records the retention decision, every remaining flag,
and measurement limitations. No cold-cache, native Office producer, broad
range-provider, concurrent-scaling, or full-goal completion claim is made.
