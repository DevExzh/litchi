# 0908 PPTX multi-slide instruction profile review

This is a finite, independent review of the four retained 0908 Callgrind
profiles. I did not build, run a workload, run a profiler, or rerun replay. The
review binds to the [canonical reducer receipt](/home/zhuhe/code/litchi/docs/performance/results/change-0908/analysis.json)
and the parent [0908 instruction-attribution report](/home/zhuhe/code/litchi/docs/performance/0908-pptx-multi-slide-instruction-profile.md).
The raw evidence passes its qualification checks with scope caveats.
The post-cleanup [analyzer checks](/home/zhuhe/code/litchi/docs/performance/results/change-0908/analyzer-checks.json)
also pass: fresh equality and four fail-closed negative mutations were
rejected.

The profiles reuse the 0907 release binary, SHA-256
`e0a7707dcd849c92e01c8dc9115d61414fba10584bb8415e442f4c225828af84`, from
source commit `3d3650ac03be2cd8ba923a5a6e940e84b14f5eab`. They run on CPU 12,
one worker, zero warmups and one sample per child in serial eager/source/source/eager
order. The Callgrind owner is the single `nm -C` symbol
`litchi_perf_baseline::publish_pptx_multi_slide_batch_edit`, with collection
initially disabled and the same full symbol used for `--toggle-collect`,
`--zero-before` and `--dump-after`.

Each leg retains exactly one numbered dump, `callgrind.out.1`, and one
program-termination dump, `callgrind.out`; no `.2` dump is present. The numbered
dump has only `Ir`, equal `summary` and `totals`, and the exact dump-after
trigger. Its only positive incoming owner edge is:

```text
run_pptx_multi_slide_batch_edit_save -> publish_pptx_multi_slide_batch_edit
source position 51174, calls 1
```

The expected-output position `51143` has zero positive collected owner edges in
all four raw dumps. That is the observed emitted-symbol scope. The receipts do
not diagnose why that edge is absent, and the review does not infer inlining,
monomorphization or another compiler explanation. No expected work is
subtracted or reconstructed. Every terminal dump has `Ir` summary and totals
of zero.

The owner partition is conserved in each raw dump. Eager legs have 22 positive
direct-child rows; source legs have 16. In every leg, owner self plus direct
children equals the summary, and the sum of all function self costs equals the
same summary:

| Leg | Owner Ir | Owner self | Direct children | Exclusive self sum |
| --- | ---: | ---: | ---: | ---: |
| eager-a | 4,577,995,784 | 6,428 | 4,577,989,356 | 4,577,995,784 |
| source-a | 108,165,330 | 2,408 | 108,162,922 | 108,165,330 |
| source-b | 108,181,418 | 2,408 | 108,179,010 | 108,181,418 |
| eager-b | 4,578,507,829 | 6,428 | 4,578,501,401 | 4,578,507,829 |

The first raw profile uses the original `capture.py` and `plan.json`. The three
remaining profiles use `capture_v2.py` and `plan-v2.json`. The v2 change is a
semantic-validation amendment: it reduces constant native 200-sample vectors
to the one captured sample and checks the remaining metadata and values. It
does not change the binary, workload, command or Callgrind options. The first
eager workload exited zero, but its original wrapper recorded
`AssertionError('operation_metrics')` because it compared the 200-sample
native envelope with a one-sample profile. The corrected semantic receipt
passes, the original failure is retained, and the workload was not recaptured.

Every leg matches its 0907 reference for corpus, sink, output and source
counters under that documented sample-axis reduction. The corpus is the
17,017,139-byte `litchi-pptx-source-edit-media-v1` archive with 445 physical
members, 229 OPC entries and 17,568,429 uncompressed payload bytes; its archive
SHA-256 is
`61b2b99083ca27ebd37955db600955e3f41289b93dba71951983164239eff757`. Every
output is 17,017,145 accepted bytes with SHA-256
`23d6d7b8dd433ff453307ee56485efadc34d6b18b455a6858d9f565c4b1b6cd9`.

| Route | Materializations | Ordinary reads / bytes | All reads / bytes | Sink calls / largest |
| --- | ---: | ---: | ---: | ---: |
| eager-a and eager-b | 229 | 741 / 16,907,750 | 1,864 / 17,030,811 | 260 / 65,536 |
| source-a and source-b | 9 | 486 / 16,909,038 | 3,372 / 17,229,633 | 2,053 / 65,536 |

The corrected semantic receipts also match the generic operation-metrics
envelope after the same constant-vector reduction. The native report and the
one-sample report are therefore not literal JSON matches; elapsed vectors and
sample axes have different lengths. The parity claim covers the retained
semantic, source, sink, corpus and output fields only.

Each stderr file records Valgrind's `brk segment overflow` warning. It is
retained as a limitation and does not add an instruction count. Each file also
prints the same collected `Ir` and `I refs` summary twice, once from the
forked runtime probe and once from the parent. Those are duplicate
presentations of one raw summary and are never summed. The raw Callgrind files,
rather than stderr, are the count authority.

The most useful direct-owner targets are:

| Route | Direct owner edge | First / second Ir |
| --- | --- | ---: |
| eager | `Package::to_bytes` | 3,189,476,327 / 3,189,730,670 |
| eager | `Package::opened_presentation_with_limits` | 994,945,242 / 995,093,332 |
| eager | `SourceBackedPackage::into_opc_package_inner` | 98,544,032 / 98,553,345 |
| source | `publish_slide_batch_commit_to_stream` | 55,654,082 / 55,654,546 |
| source | `SourceBackedPresentationEditor::from_read_at` | 38,232,988 / 38,246,730 |
| source | `SourceBackedSlideBatchEdit::set_shape_texts` | 13,962,259 / 13,963,877 |

These rows are inclusive immediate-owner attribution. Nested rows must not be
added to them. The opened eager edge mixes initial snapshot/fingerprint work
with replay-owned capture in this helper, so it needs explicit phase boundaries
before it can be assigned to production. Source-path classification is
evidence for attribution and does not prove semantic ownership. The raw
profile is `Ir` guest-instruction attribution, not native elapsed time, cycles,
allocation, RSS, physical I/O or a speedup measurement.

The next candidate priority is an `owned_from_vec` control. The eager ReadAt
materialization creates an owned package without the physical source provenance
that enables preservation copying, while ordinary owned routes already
retain/use the preservation writer. Qualify that control with the same corpus
and edit, separate production and replay boundaries, and copy/regenerate
evidence. It must preserve source lineage and freshness, resource and sink
limits, complete changed dependency closure, opaque and ordering records,
refusal identities, semantic and untouched-member checks, durable inverse
checks, and sink-failure behavior. This is a route qualification handoff and
does not authorize a production change or skipped snapshot validation.

The retained raw Callgrind files are byte-exact evidence and include expected
trailing spaces in `desc`, cache and calls records. Whitespace checks are
scoped to authored review/reducer files; no raw capture was rewritten.
