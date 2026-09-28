# 0822 — bounded PPTX edit CPU profile

This packet measures a small diagnostic boundary around the public PPTX edit
transaction on the unchanged `353aa00a7da2795e6b4c28708138a103a143917e`
source. It does not change production code, ordinary-save semantics, or the
0819/0821 latency baseline. iWork remains outside this packet.

The probe opens the canonical checked-in
`test-data/ooxml/pptx/shapes.pptx`, edits shape `(0, 0)` through the public
transaction API, commits and applies that edit, and serializes the result
outside the timed edit region. The input is 68,822 bytes with SHA-256
`19fde9b87e33dd1a95fdbba0cf6abc2278bf03874f4665c7f8b88b6afe4a2571`. The
expected output is the admitted 0821
`docs/performance/results/change-0821/artifacts/real-002-pptx/default.pptx`,
68,284 bytes with SHA-256
`38c7fd3cc5037316e0a3a9c8cb6e6b013a95dd03b39bbd7fb357fc3cc3b9e5cf`.
Both identities are checked before and after every child process.

The standalone packet probe is `pptx-edit-profile-0822`; its exact sampled
owner is `pptx_edit_profile_0822::edit_region_0822`. The `control` arm calls
the edit directly, `wrapped` calls the same operation through the named
non-inlined wrapper, and `fp` uses that wrapper in a release build with
`-C force-frame-pointers=yes`. Only two executables are built (`ordinary` and
`fp`); the three arms are CLI configurations of those binaries.

Run the root-owned lanes serially after the independent `source-review.md` is
present and frozen:

```sh
python3 -B docs/performance/results/change-0822/quality.py
python3 -B docs/performance/results/change-0822/build.py
python3 -B docs/performance/results/change-0822/decode.py --symbols
python3 -B docs/performance/results/change-0822/capture.py qualification
python3 -B docs/performance/results/change-0822/capture.py native
python3 -B docs/performance/results/change-0822/capture.py perf
python3 -B docs/performance/results/change-0822/decode.py
```

Qualification has one counterbalanced order, three arms, three samples, and no
warmup. Native capture uses six deterministic counterbalanced orders, three
warmups, and thirty samples per arm: 18 reports and 540 timed samples. All
workloads are pinned to CPU 12. Perf uses two whole-process repeats of 2,000
wrapped samples at `cycles:u`, 997 Hz, and frame-pointer call graphs. Its
decoded stacks use exactly `perf script --no-inline --ns`; `nm -S
--defined-only` and `objdump --disassemble=<matched-symbol>` bind the owner to
the live FP binary before cleanup.

If perf is unavailable or permission-denied, the packet retains a typed
unavailable receipt and its reason. It never substitutes a fabricated profile.
Raw perf data and decoded frames are retained in deterministic gzip archives;
redundant uncompressed copies are removed only after byte/hash verification. Symbol output,
assembly, source hashes, and build hashes remain available to the final
reader. Timings and sampled stacks are descriptive perturbation evidence only;
they do not support a source speedup, ordinary-save phase fraction, or adoption
decision.


The commands above record the original capture order. Existing evidence is
immutable, so capture drivers refuse to overwrite this packet. New
measurements require a fresh packet and target; offline replay uses the
retained artifacts without rebuilding:

```sh
python3 -B docs/performance/results/change-0822/analysis.py --check
python3 -B docs/performance/results/change-0822/validate.py --final
python3 -B docs/performance/results/change-0822/seal.py --check-head
```

The validator also replays `root_audit.py`, `frame_audit.py`, and
`stack_diagnostics.py` in check mode. Final validation retains 23 reports /
4,549 samples. The two binary hashes survive in build and cleanup receipts;
the target directory was removed after exact verification, releasing 3,684
files / 2,054,019,575 logical bytes. There is no scratch directory. The frame
audit's original empty-stack rejection and the duplicate analysis-write
refusal remain documented in `execution-notes.md` and their retained logs.
