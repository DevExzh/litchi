# 0819 — current real-file ordinary-save baseline

This batch completes the three-file timing matrix queued by 0817 and admitted
after the DOCX preservation repair in 0818. Base is `0c784f7aec`. Production,
the runtime harness, tests, and dependencies remain unchanged. Prior packets
remain immutable; prior timings and development-profile outputs are not pooled
into this baseline.

The plan covers lifecycle, semantic edit, default full-durability atomic save,
and counting publication for the named DOCX, XLSX, and PPTX inputs. Fresh
release artifacts and independent XML/ZIP preservation admission precede
qualification and every timing lane. PPTX counting publication materializes
bytes; independent phase measurements are not an additive decomposition.

Root owns all Cargo, exporter, benchmark, and profiler execution and retains
each process handle until terminal. Preparation and static review are delegated;
heavy offline readers wait until captures have terminated. Drivers freeze input
identities before execution, retain failures, and refuse output replacement.
The current harness lock differs from the root lock; exact differences are
retained without changing dependencies or claiming historical comparability.

The completed matrix has 12 qualification reports/samples, 72 native reports with
2,160 samples, and 24 observer reports with 72 samples. Edit samples publish no
file; report/sample counts must not be relabeled as saved-file counts. Native
and observer results remain separate, and procfs empty controls are retained
without subtraction. The broader OLE2/OOXML goal remains active; iWork is excluded.

All six fresh quality gates pass, including 641 tests with one ignored and zero
failures. Three release builds, six-case artifact admission, qualification,
native, and observer captures pass. Native lifecycle p50 medians are
5.220218/5.419553/7.422314 ms for the DOCX/XLSX/PPTX inputs. Eight cases retain
fourteen spread flags and four tail flags; no p50 block spread exceeds 1.05.
This is a baseline, not an optimization or historical speedup claim.

See [the main report](../../0819-real-file-ordinary-save-baseline.md),
[analysis.md](analysis.md), [resource-summary.json](resource-summary.json), and
[results-review.md](results-review.md). Reader-only schema failures are retained
in numbered `reader-attempt-*.log`; no workload was recaptured to repair a
reader. Final cleanup and exact staged/committed custody are bound by
`cleanup.json` and `seal.json`.
