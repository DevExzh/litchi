# 0890 matched MCE candidate packet

This directory contains the reviewed source freeze, retained correctness and
release-build custody, managed executable preparation, and matched capture
evidence for per-start expanded-name reuse. The candidate is rejected: primary
paired p50 ratio 0.99381 with 95% interval 0.98861–1.01112 fails the frozen
adoption gate. See `decision.json` and the [0890 report](../../0890-mce-validated-name-reuse.md).
Production is restored; the tested source remains in `candidate-v2.patch`.

Freeze the candidate source and its patch before any release build. The
create-only `candidate-freeze-v2.json` lists every changed file under
`crates/litchi-ooxml-common/src/mce/`, including a new regression test file if
one is added, and `candidate-v2.patch` must have exactly that path set. The
release builder consumes those files without rewriting them, reconstructs the
retained 0885 source archive, overlays that inventory, and checks the complete selected source tree. The four
`tools/perf-baseline/src/bin/xml_attribute_bounds*` files are the only current
source differences explicitly excluded because they belong to the 0889
diagnostic binary. Root build manifests, Cargo configuration, and the release
harness remain covered by the selected-tree check.

The following records the original capture sequence. Its outputs are
create-only; use a separate packet location for a new experiment rather than
overwriting retained evidence. Candidate build consumes the frozen overlay,
which must be restored before repeating live source checks:

```sh
python3 -B docs/performance/results/change-0890/build-executed.py
python3 -B docs/performance/results/change-0890/prepare_exec.py \
  --before-native docs/performance/results/change-0885/binaries/native \
  --before-allocator docs/performance/results/change-0885/binaries/allocator \
  --after-native docs/performance/results/change-0890/binaries/native \
  --after-allocator docs/performance/results/change-0890/binaries/allocator
python3 -B docs/performance/results/change-0890/freeze.py
python3 -B docs/performance/results/change-0890/experiment.py \
  --plan docs/performance/results/change-0890/plan.json \
  --output-root docs/performance/results/change-0890/capture
python3 -B docs/performance/results/change-0890/reduce.py \
  --plan docs/performance/results/change-0890/plan.json \
  --experiment-root docs/performance/results/change-0890/capture \
  --output docs/performance/results/change-0890/reduction.json
python3 -B docs/performance/results/change-0890/verify_packet.py \
  --plan docs/performance/results/change-0890/plan.json \
  --experiment-root docs/performance/results/change-0890/capture \
  --reduction docs/performance/results/change-0890/reduction.json \
  --output docs/performance/results/change-0890/verification.json
```

To verify the retained packet without building or timing, rerun `verify_packet.py`
with a new `--output` path, and replay the
retained plan with `python3 -B docs/performance/results/change-0890/freeze.py --verify --retained-only`; this checks the retained descriptors without launching a workload.

The experiment reuses the reviewed 0887 timing and semantic driver, with
schemas and paths adapted to 0890. It schedules 64 native children over the
four 0887 cases and eight allocator children for the primary case, with three
warmups, 50 native samples, 15 allocator samples, fresh processes, and
counterbalanced AB/BA order. Each role/mode has one content-identical managed
copy, created by `prepare_exec.py` and reused by its scheduled children. The
copy control narrows executable-path and inode history differences; it does
not prove that RSS variation has been eliminated. RSS remains a review flag.

The primary gate remains a paired native p50 median ratio at most `0.98` with
bootstrap upper endpoint below `1.0`. Any control or primary metric pair over
`1.05`, and any native p50 spread over `1.20`, remains visible for review.
Samples are never pooled and allocator counters are not an adoption gate.

Retain the 0885 before binaries and the 0890 candidate binaries and build
receipts for replay. After the aggregate and review receipts pass, remove only
the marked 0890 source/target and managed-copy directories. The packet
verifier can replay command and reduction evidence from retained managed-copy
descriptors after those copies are removed; it never rebuilds or launches a
workload.

Large executable binaries and the 0885 source archive are intentionally ignored
by Git and retained locally at the receipt paths. Replay requires these
artifacts; a fresh checkout alone does not contain them. Their hashes and
build/source receipts are committed, including the correctness binary custody
JSON files.
