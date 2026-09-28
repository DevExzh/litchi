# 0817 — real-file ordinary-save admission failure

This batch advances the queued real-file corpus work from 0816 on unchanged
production and runtime `tools/perf-baseline` code. One integration test is
corrected to expect the feature-dependent instrumentation identity. The rejected plan would measure
the documented open/edit/save lifecycle, semantic edit, full-durability atomic
publication, and counting publication for three named DOCX/XLSX/PPTX inputs.
The source files have tracked provenance; this does not infer a Microsoft
producer or establish a new independent Office round trip.

Artifact export and an independent ZIP/XML audit precede timing admission.
Any refusal or unexplained preservation difference remains evidence and cannot
be relabeled a successful edit merely because the output is deterministic.
Native timing uses a separate binary from allocator/procfs observations.
No production optimization or historical before/after speedup is claimed.

The existing tracked harness lockfile differs from the root workspace lock.
Both are retained with exact external-package differences in `lock-parity.json`;
this baseline uses the harness lock without updating dependencies. All quality
gates use that same locked graph; earlier root-build timings are not pooled.

Root owns all Cargo, workload and profiling execution. Offline readers may run
only when the relevant capture handles have terminated. Drivers retain failed
attempts and refuse overwriting evidence. Owned build and scratch directories
will be removed after verification; unrelated workspace files remain intact.

The broader OLE2/OOXML goal remains active. ODF is deferred and iWork excluded.

The actual outcome is failed artifact admission: six corpora / thirty outputs
were exported, with zero qualification/native/observer reports. The frozen
audit and its 22 errors are retained under `admission-0/`; the diagnosis must
not equate every audit error with a production defect. See the
[main report](../../0817-real-file-ordinary-save.md) for exact distinctions.

Offline replay: `python3 -B docs/performance/results/change-0817/validate.py --final`.
Root verified all three executable hashes and removed both owned target and
scratch directories (9,063 files / 8,591,960,251 logical bytes). Post-cleanup replay passes.
The seal binds all packet files, six performance documents, and the single
allowlisted integration test; unrelated workspace changes remain outside it.
