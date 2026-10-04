# 0900 methodology review

This is a post-acquisition methodology record for the 0900 XLSX phase
diagnostic. Review correspondence began before qualification, but the written
review overlapped native capture and makes no pre-capture approval claim. It
does not make a performance or adoption claim; the canonical native replay and
its immutable-artifact checks are the packet's evidence.

## Scope and provenance

The selector is
`xlsx_source_backed_cell_values_one_percent_edit_save`, using the `medium` and
`dense-sparse` shapes. The packet is diagnostic only: it does not compare
timings with an older binary and does not justify a production change.

The executable is the retained 0894 native binary. Its source revision is
`506aa2edae6666d9c76ae6827c2f5941a8e47896`, bound by the 0894 manifest's
10,771 file hashes, the native build receipt, and the binary digest. The
repository `HEAD` recorded in `plan.json` is the acquisition context; it is
not evidence that the detached retained executable can report that Git
revision at runtime.

## Acquisition and validation

Qualification used 40 measured samples and two warmups per shape. Each child
was semantically and resource validated before native capture. Native uses six
fresh child blocks per shape, 100 samples and ten warmups per child, alternating
shape order on CPU 12 with no retry or replacement. Qualification is kept out
of the native summary.

The frozen 0705 validator is used only through its individual schema,
corpus, sink, statistics, source-vector, and invariant checks. Its byte digest
is bound in `plan.json`; its historical timing output is not imported.

The first medium qualification workload exited successfully, but the original
helper rejected it because the retained binary was built with a compile-time
manifest path that has since been removed, so its runtime Git revision and
dirty-state fields are null. The child report and failure receipt remain
unchanged. `metadata-addendum.json` records the cause and custody; `run_v2.py`
accepts the unavailable runtime metadata only after the external source,
build, and binary checks pass, then replays that child without rerunning the
workload. This correction changes validation metadata only.

## Timing boundary

The resolved 0900 boundary is recorded in `scope-clarification.json` and
`source-scope.md`: total sums open, planning, staged sets plus commit, and
publication. Publication includes consuming the editor and dropping the
returned publication snapshot before elapsed time is sampled. Sink
preparation, remaining commit/sink/handle drops, source/oracle diagnostics and
output reopen are excluded; reopen remains a separate diagnostic phase. This
exact boundary must accompany any 0900 result.

The 0705 packet also included the returned publication snapshot's drop, so the
earlier scope concern is resolved. Current production source differs from
0705; this packet remains a diagnostic refresh and does not become a
historical before/after comparison. The current packages have seventeen
archive members, including eight media members, and twelve untouched members.

## Preliminary signal and next capture

The qualification children initially pointed to publication as the largest
timed phase: 43.97% of the medium total (16.683 ms p50) and 46.51% of the
dense-sparse total (32.936 ms p50). The completed canonical native analysis
retains that ordering across twelve children (six per shape) and 1,200 samples: publication
shares range 43.88–44.33% for medium and 46.26–46.81% for dense-sparse;
planning ranges 33.89–34.40% and 32.53–32.79%, respectively. Qualification
remains separate from the native summary. This stable diagnostic ranking
supports a scoped publication follow-up without implying a speedup or adoption
eligibility.

With canonical replay confirming this ordering, the next bounded capture
should scope the publication owner path for both shapes under the resolved
timer boundary. A publication allocation or Callgrind profile can then identify
an implementation target. Planning should be the second follow-up because its
prior 0705 allocation evidence was large (67,845 medium and 129,411
dense-sparse calls), while 0831/0832 only measured small lifecycle/allocation
changes and did not attribute planning versus publication time. No allocator
result is implied by this phase-only packet.

## Evidence

- `plan.json` — source/binary custody, schedule, scope, and environment.
- `qualification/b00-medium/` — retained successful workload plus original
  validation failure and corrected validation receipt.
- `qualification/b00-dense-sparse/` — corrected qualification receipt.
- `metadata-addendum.json` — immutable metadata correction and no-rerun
  record.
- `scope-clarification.json` and `source-scope.md` — resolved timer and source
  boundaries.
- `run.py` and `run_v2.py` — original and metadata-corrected helpers.
- `../change-0705/README.md`, `../change-0705/plan.json`, and
  `../change-0705/analyze.py` — prior phase boundary and frozen validator.
- `../../0831-xlsx-empty-column-actions.md` and
  `../../0832-xlsx-inline-column-map.md` — prior small allocation/lifecycle
  attempts without phase attribution.
