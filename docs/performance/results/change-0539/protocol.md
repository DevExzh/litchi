# 0539 XLSX transient attribute ownership experiment

This batch tests the unapplied 0537 draft against the 0538 planning-allocation
harness. The production difference is confined to the raw worksheet codec:
decoded references and numeric metadata can remain borrowed for immediate
consumption; the cell type becomes owned before retention across XML events.
Checked scanning, normalization, validation, resource limits and error order
remain unchanged. The 0531 MCE search and 0522 lossless layout shortcut are not
part of this candidate.

The strengthened attribute test is installed before baseline freezing. It
keeps the plain numeric attribute case, adds entity-encoded numeric attributes,
and checks exact style, metadata and shared-string values for both cells. This
same test source remains in baseline and candidate builds. Baseline/candidate
manifests must differ only at `raw/worksheet/codec.rs`.

The serial protocol is baseline native/allocation repeat 1, candidate repeat 1,
candidate repeat 2, then retained baseline repeat 2. The last baseline runs
under the candidate checkout with both executable and execution-source custody
recorded. Builds and tests do not overlap timed children. The owned target
contains test temporary files and retained executables, avoiding the default
`/tmp` quota encountered in 0538.

Primary native rows are unmanaged one-percent edits for medium and dense-sparse
corpora, 20 warmups and 200 samples, two repeats per build. Guards cover one-cell
edits, managed one-percent edits, vendor extensions and noncompact attributes,
10 warmups and 30 samples. The 36 native children contain 2,440 measured samples.
Allocation captures independently use three warmups and 30 samples, yielding
240 observations per phase across eight children. Instrumented elapsed values
are not native performance evidence.

The frozen primary gates require at least 1% reductions in workflow p50 and
mean, and at least 2% in planning p50, for every primary shape/repeat. Planning
allocation calls must fall at least 5% in every pair; the supplementary
allocation policy requires non-increasing planning bytes. Reallocations and
incremental region peak are reported separately for planning, commit and
publication. Every adverse change or same-build drift over 5% is retained for
individual review. No mixed geometric mean can replace a failing primary row.

If the native/allocation pilot passes, admission also requires planning Ir
reduction of at least 1% in every profile, eager XLSX read guards, and applicable
quality gates. If it fails, reject the codec change and retain the evidence and
behavioral guard. A lower allocation count alone does not admit the candidate.
No cold-cache, physical-range, native-producer, scaling, or program-level speedup
claim follows from this experiment. OLE2/OOXML remain first; ODF is deferred and
iWork is excluded.
