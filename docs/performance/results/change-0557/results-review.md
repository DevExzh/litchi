# 0557 results review

The baseline noise pilot is valid as retained evidence and must stop the
candidate gate. It has eight primary case-shape rows, two repeats, and 1,000
timed samples per report. The largest p50 repeat change is 9.787197844389572%,
which is above the frozen 5.0 percentage-point limit. No candidate result or
optimization claim follows from this pilot.

The reviewed input and evidence hashes are:

```text
plan.json                    d8111317e6d6d44219c844aaf740ea50fc149c361601bb6483a47d33843f88d8
run.py                       c9c0db8ac91da2d8604c3a7d70b77f5927f0b6ac9c6fb8066e4b7c350b8b8999
analyze.py                   010f44e811d852ade5f5a62fcd3092961815ef0a3e81217697bc170b5e92bbb2
expected-primary.json        84f11d08100d03bf00cdec0484a18bb981dbaa35f3486be67b07d49b2ec6a97e
frozen-inputs.json            d42c414d25caff240f400b6fddae04f4c768e43e9ad21b391f408fd5bdcf53c0
baseline/source-manifest.json 68fdc09a68bf779e17c5852e5f7a0c0a8b15253e35ed6762e42018a37c5011c5
baseline/binary-normal.json   e433f0710f54f803865543825b0a802b2a669effb42e51d3ec00b294ed551641
retained baseline binary       fa69fc4675d67a1f3c73d056f77ccccf7d795befcfa6458b2a014fa86d0778c4
noise-analysis.json            949e1210163396a54fe4e7ea98f5a4c2324a2ef47b31c534b11ad7fe6d87bb3e
review_noise.py                5fe3d7c12dc6c370cafaf91f117af6b0f39b97064272fccca355a79c9c75f9be
noise-review.json              746537da97303d232f7fa3894281804e3f629465d1c4c269568ffbf3f8bc8f92
noise-review.md                4cc7cbaff2da4b43cdf47a0a438a6d459a7b97ad37f536512820c6f9da7001a5
workspace-lock.json            e3a8151a81efac7ba3eba708467ea4d420c5a218830c829274ba591bc3fd2000
workspace-Cargo.lock           9111221ee9d100daf90328a544613cb3f70287611dcc55a37d3b1b7a5d99c91a
```

The frozen input digests still match `frozen-inputs.json`. The baseline
manifest contains 8,595 source paths and independently matches the live source
path set and hashes. The workspace lock and retained lock copy also match the
bound lock digest.

The expected noise matrix exactly matches the plan: 16 noise receipts, 16
reports, 16 catalogs, 16 RSS sidecars, and 16 host records. Every noise
receipt has `success: true`, `child_started: true`, and exit code zero. All
receipts use baseline output and baseline execution stages. Their guard scopes
are 32 `manifest-only` output observations and 32 `live-source` execution
observations, all bound to manifest
`68fdc09a68bf779e17c5852e5f7a0c0a8b15253e35ed6762e42018a37c5011c5`. Every
noise child used the same retained normal binary hash
`fa69fc4675d67a1f3c73d056f77ccccf7d795befcfa6458b2a014fa86d0778c4`.

All 16 reports have the two source-backed cases, four declared shapes, 20
warmups, and 1,000 samples. Each report has 1,000 values for `open_ns`,
`plan_ns`, `commit_ns`, `publication_ns`, and `reopen_ns`. The five allocation
phase vectors are present as 80,000 explicit unavailable samples with the
operation allocator scope; they are not synthesized zeros. The baseline
quality and allocator integration receipts passed separately, but no
allocator performance binary or allocator capture matrix exists.

The p50 calculation was independently reproduced from the retained reports:

| Case | Shape | R1 p50 (ns) | R2 p50 (ns) | Change |
| --- | --- | ---: | ---: | ---: |
| one edit | medium | 5,361,466 | 5,348,315 | -0.245287% |
| one edit | dense-sparse | 35,068,957 | 35,340,346 | +0.773872% |
| one edit | noncompact | 6,524,370 | 5,885,817 | -9.787198% |
| one edit | vendor-extension | 5,354,064 | 5,390,660 | +0.683518% |
| one percent | medium | 20,229,965 | 20,205,582 | -0.120529% |
| one percent | dense-sparse | 39,110,999 | 39,090,805 | -0.051633% |
| one percent | noncompact | 22,275,983 | 22,302,837 | +0.120551% |
| one percent | vendor-extension | 20,296,597 | 20,278,280 | -0.090247% |

For the determining row, the retained values give
`100 * abs(5,885,817 - 6,524,370) / 6,524,370 = 9.787197844389573...%`;
the JSON float is `9.787197844389572`. The stored analysis correctly reports
`too_unstable: true` and `candidate_gate_preregistered: false`.

The supplemental repeat review independently reproduces 464 comparisons and
135 absolute changes above 5%: 20 elapsed-statistic flags, 111 source-phase
distribution flags, and 4 residual process-sidecar flags. Only one elapsed
p50 flag exists, and it is the determining noncompact one-edit row above. The
other elapsed flags concern tails, extrema, dispersion, or confidence
endpoints. Phase flags compare statistics over the five retained timing
vectors for the same baseline binary and unchanged phase definitions. They
show repeat distribution variability; they do not identify a cause, establish
a candidate effect, or relax the p50 stop. The review is distribution-level
and does not provide paired per-sample causal attribution.

There is no `candidate/` or `final/` stage directory, candidate source
manifest, candidate binary, candidate report, candidate receipt, or candidate
RSS sidecar. Baseline contains no `alloc-*.receipt.json`, no
`native-*.receipt.json`, and no `binary-alloc.json`; the allocator test receipt
is a quality/integration check rather than a performance capture. Therefore
the 50,000 ns floor is frozen by the plan but has no matched native baseline
denominator to evaluate, and no allocator phase speedup or OLE2/OOXML
optimization claim can be made.

The analyzer self-test passed with `python3 -B`. This review performed only
read-only JSON/hash/cardinality/statistical checks and did not run Rust jobs,
builds, measurements, or captures. `run.py` records the noise custody and
ABBA identity fields; the coordinator remains responsible for honoring the
unstable-noise stop before any future candidate invocation.

