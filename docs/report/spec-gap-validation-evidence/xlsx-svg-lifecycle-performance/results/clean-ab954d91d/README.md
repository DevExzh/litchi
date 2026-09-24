# XLSX SVG lifecycle sealed baseline

Review status: root and independent review approved the capture integrity and
descriptive measurements. This is an allocator-instrumented absolute baseline,
not a before/after improvement, causal speedup, regression, scaling guarantee,
native changed-save acceptance, or rendering claim.

The measured clean source is
`ab954d91dd6d6ca3e0df18b59ee826cef99bcae8`, with lifecycle production pinned to
`ac288a303264f9ea0bb4081baa44031bee5b79a7`. All 19 production paths passed exact
Git-blob checks. The complete source manifests cover 4,871 inputs,
including the committed fixture, harness, lockfile, transitive Cargo sources,
runner, verifier, and uncertainty tool. No ignored third-party symlink or
omitted build input supplied this run.

The 69 acceptance lanes each ran in three fresh processes with two warmups
and 20 measured samples: 207 JSON receipts and 4,140 samples. Seven lanes
record expected typed refusals (420 samples); these are refusal evidence,
not successful edit measurements. All 207 stderr files are empty, and all
GNU time sidecars report successful process exits. The seven exploratory
lanes are outside this capture.

`verification.json` and the independently replayed `root-verification.json`
agree. Both recompute the report and check source identity, semantic/readback
outcomes, allocator equations, typed refusals, process shape, input hashes,
and stable executable identity. Before/after source manifests have SHA-256
`5094e3169512176d7f107ecad233568d2287a13892b837ea8bd6e5ecc96adc1d`.
The measured executable's before/after SHA-256 is
`76c82487bb04cf89df28adf185e049e61bf86ce8488be2b1bf537388edf64abd`.
The disposable Cargo target was removed after capture; its retained digest
receipts bind the measured executable rather than requiring it to stay live.

In-run `build-provenance.txt` records AMD EPYC 9R45, 32 cores, 129,447,068 kB
memory, Linux 7.0.0-1012-aws x86_64, and Rust 1.95.0. This capture includes
the runner's host-metadata quoting correction. It does not use the earlier
capture's post-run host supplement.

`uncertainty.json` and `uncertainty.md` regenerate exactly from these raw
receipts. They report each process median and sample range, the range of the
three process medians, and process RSS. With n=3 these ranges are descriptive,
not confidence intervals. Requested allocation bytes and peak-live deltas
cover each instrumented sample; RSS covers the entire fresh process. Fixture
setup, timed scope, validation, and limits follow the requirements at the
measured commit. No comparison to the earlier provisional run is implied.

The original 635 capture files are retained byte-for-byte, alongside the
matching root verification receipt and this README. Raw reports, timings,
provenance, and JSON formatting were not rewritten.

To replay, use `verify.py` from the measured commit with `--root` pointing to
a clean checkout of that commit, `--results` pointing to this directory,
`--report` pointing to `report.md`, and `--output` pointing to a new review
receipt. After successful verification, run the same commit's
`uncertainty.py --results <this-directory> --output <new-markdown-path>
--json-output <new-json-path>` and compare both derived files. Verification
does not rebuild or replace raw measurements.
