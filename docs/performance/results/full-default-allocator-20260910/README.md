# Full default allocator resource capture

The retained report contains all 201 default case/corpus rows, with 15 samples
and 11 measured allocation vectors per row. Sample identities remain aligned
with sorted elapsed observations. These instrumented elapsed values are not
latency evidence.

The measured source is `fccbe6595f3a29561a8bfc6c192d8fa5e874ad05` plus the frozen
instrumentation patch retained in the bundle. Publication integration was
validated at `bfdc911c0577c6a3304b4a70b8b70e2661721355`; later corpus and XLDM
changes are not retroactively attributed to the capture.

Extract and verify the portable bundle:

```sh
tar -xzf publication-bundle.tar.gz
cd litchi-full-default-allocator-publication-bfdc-20260910
python3 verify-publication.py
```

The bundle contains the raw compressed report/catalog, exact commands, locks,
source and binary identities, host/run provenance, checker, adversarial tests,
and integration patch. The measured binary is represented by its hash and
build recipe. The report has 31 corpora and 201 exact case/corpus bindings;
all allocation vectors are measured and preserve their sample order.

Root validation passed 360 harness library tests with one ignore, strict
all-target Clippy, scoped Rust 1.95.0 formatting, and the portable verifier.
Independent reviews cleared the capture and publication. Commands, source
hashes, and compressed root logs are retained here.
