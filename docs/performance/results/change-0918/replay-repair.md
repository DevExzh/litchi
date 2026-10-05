# 0918 postcapture relocation repair — 2026-10-05

The frozen `comparison-replay.py` succeeds in the original checkout. The
independent final review's literal reconstructed-copy replay fails on its
first child with `fixture path mismatch`: the raw probe report embeds the
original absolute input path, while the unchanged 0917 validator accepts the
copied repository path. The pre-capture relocation smoke checked path
resolution and copied inputs, not a complete captured report, and therefore
did not detect this failure. No performance child is rerun or excluded.

`comparison-analysis.py` is a separate replay-only adapter. It binds the
frozen runner, protocol, manifest, replay and validator hashes and verifies
the copied fixture bytes against the frozen manifest. It then supplies the
validator's existing `path_abs` field with the original input path derived
only from the frozen manifest's `repository_root` and relative fixture path.
This is a validation alias: the adapter must not read the original fixture
path, change the raw report, change its canonical hash, or alter any metric,
correctness gate, reduction, confidence interval or admission threshold.

The frozen runner, original replay, protocol, manifest, raw reports, child
receipts and reduced result remain byte-for-byte unchanged. Their original
replay failure remains part of the independent review evidence.

For a copied packet, reconstruct these paths below a fresh external root:

```text
docs/performance/results/change-0918/       complete packet, including adapter
docs/performance/results/change-0917/capture.py
docs/performance/results/change-0917/native-protocol.json
docs/performance/results/change-0917/probe/ complete unchanged eight-file tree
test-data/ole/doc/NoHeadFoot.doc
test-data/ole/doc/FloatingPictures.doc
```

Run:

```sh
python3 -B docs/performance/results/change-0918/comparison-analysis.py
```

The optional `--packet` argument selects a packet in that same reconstructed
layout. Qualification and profiling have separate verifiers and retain their
documented external custody requirements. The final independent review must
confirm full copied-packet replay and rejection of altered fixture paths,
fixture bytes and captured/reduced evidence before this repair is accepted.

Before final acceptance, the helper's initial default output collided with its
already retained report and returned exit 1 after validation. This invocation
is preserved in `analysis-initial-failure.json`. The final helper defaults to
stdout without writing a report; `--output` is optional. Earlier helper reports
remain historical evidence, and final review binds the corrected helper hash.
The final adapter SHA-256 is
`1c72d1d84ba7a789fced7ec24d4d7e275ef9a6a8a1b9fbd92ef24f7f9bd60d87`.
