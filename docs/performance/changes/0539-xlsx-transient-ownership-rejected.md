# 0539: reject transient XLSX attribute ownership after matched gates

The candidate avoided forced ownership of event-local references and numeric
attributes in the raw worksheet parser, retaining owned cell types across XML
events. The same strengthened entity-decoding test was present in both builds.
The codec change is rejected and restored; only the behavioral guard remains.

Only one of four primary shape/repeat rows satisfies all frozen native gates.
Positive values below mean improvement; negative values mean regression.

| Shape/repeat | Workflow p50 | Workflow mean | Planning p50 | All native gates |
| --- | ---: | ---: | ---: | --- |
| medium/r1 | 1.3376% | 1.0910% | 2.5353% | pass |
| medium/r2 | -0.2076% | 0.0942% | 2.9517% | fail |
| dense-sparse/r1 | 0.2615% | 0.6959% | 2.0413% | fail |
| dense-sparse/r2 | -1.1195% | -1.0189% | 1.6821% | fail |

The mechanism worked locally: planning allocation calls fell from 67,957 to
58,741 on medium and 129,541 to 111,749 on dense-sparse in both repeats,
reductions of 13.5615% and 13.7346%. These were short-lived strings: planning
allocated bytes fell only 0.2639% and 0.4617%, while all 12 planning/commit/
publication pairs retained identical reallocation and incremental-peak vectors.
Commit lost one small allocation per updated cell; publication vectors were
unchanged. Allocation savings do not override the failed workflow gates.

The [evidence bundle](../results/change-0539/README.md) retains 2,440 native
samples, 240 allocator samples per phase, 28 matched adverse flags and 69
same-build drift flags. The conditional planning profile, hardware and eager
read lanes were not run after native rejection, as frozen before capture.
No instruction, cold-cache, range-source, native-producer or scaling claim is
made. Baseline, candidate and restored-source XLSX tests each passed 1,292
executions; final applicable checks are recorded in the bundle.

The test now checks plain and entity-encoded style/cell-metadata/value-metadata
attributes and their exact stored values. Production remains at the baseline.
OLE2/OOXML stay first, ODF remains deferred, and iWork is excluded.
