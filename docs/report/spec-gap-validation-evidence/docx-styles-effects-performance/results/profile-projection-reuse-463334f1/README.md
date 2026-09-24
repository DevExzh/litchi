# DOCX `stylesWithEffects` projection-reuse profile

This directory retains the bounded profile captured from the clean isolated
source checkout at harness commit
`463334f1a32de687ffb2b3357e55816dcfd3102e`. Its production source pin is
`d000d977b99e03f8542c7dae74acf767a91b1feb`, and its approved current
correctness smoke is retained by commit
`c4ac516353591f88a9b797349002d05737614e78`. The profile runner and
`verification.json` both passed.

The matrix has 31 lane/scale rows, three fresh processes per row, two
warmups per process, and twenty measured samples per process: 186 warmups and
1,860 measured samples. It covers native resources plus deterministic 64 KiB
and 1 MiB generated resources. The correctness smoke retains the refusal,
malformed-input, signed, absent-owner, inverse, and independence coverage;
those smoke-only cases are not timing rows. No 8 MiB or near-limit timing was
run.

The raw capture set is 306 files totaling 12,391,147 bytes: 93 profile JSON
receipts, 93 `/usr/bin/time -v` sidecars, 93 empty stderr logs, and the
runner's build, source, fixture, metadata, command, host, and verification
receipts. `raw-manifest.sha256` lists every raw file's SHA-256, relative path,
and byte count. Its digest is
`e8b5f5669ee2bac663c1d908152de94b99f0a8eea1c0eb391f4cfa5085aedbaf`.

The captured binary SHA-256 is
`5d9d02b5236964aad9c3441fe52ebeebafd802effea5a9151c6636db01862de6`; the
binary and disposable Cargo target were removed only after the successful
verification sentinel. The locked offline build used `CARGO_INCREMENTAL=0`,
`LC_ALL=C`, the default linker, and Rust `1.95.0` (`rustc 1.95.0
(59807616e 2026-04-14)`, Cargo `1.95.0 (f2d3ce0bd 2026-03-21)`). The source
bundle is `captured-source.bundle`; `git bundle verify` passes, it requires the
known prerequisite commit `d000d977b99e03f8542c7dae74acf767a91b1feb`, and its
SHA-256 is
`1bbef1c673cd4947e794167d52f5a1f4282c398ac8b56b96b3abfb9463c3e994`.

The host was Ubuntu 26.04.1 on AMD EPYC 9R45 with 32 logical CPUs and
129,447,068 KiB total memory, with affinity `0-31`. Host one-minute load was
4.73 before and 5.45 after. An external `cargo-fmt`/`rustfmt` workload was
observed immediately before launch; the host was shared and unisolated.
Physical-core count, mount identity, and a continuous competing-process
census were not captured. Therefore these are bounded absolute observations
under observed contention, with no idle-host, uncontended, speedup, or
all-workflow claim.

All 1,860 samples passed the semantic, opaque-member, exact-inverse, actual
success, and allocator-balance checks. Allocator counters are the harness's
measured allocation window and are distinct from process RSS. Each
`/usr/bin/time -v` sidecar records one fresh process's maximum RSS, including
its warmups and samples; neither is a managed memory-cap result.

Selected after-profile medians (milliseconds unless noted) are descriptive
only:

| row | elapsed | apply / publish / serialize | reopen | allocation calls | requested bytes | peak-live delta |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| replace main, 1 MiB | 141.637 | 75.949 / 79.677 / 3.721 | 29.346 | 1,101,390 | 116,701,991 | 5,767,866 |
| replace glossary, 1 MiB | 144.718 | 77.586 / 81.320 / 3.720 | 30.027 | 1,115,855 | 118,082,796 | 5,790,767 |
| replace main, native | 22.294 | 9.450 / 9.627 / 0.180 | 5.902 | 138,320 | 23,296,261 | 562,415 |
| replace glossary, native | 21.478 | 9.978 / 10.132 / 0.152 | 5.268 | 132,962 | 25,106,122 | 721,925 |

The 1 MiB replacement rows contain the largest repeated apply and allocation
windows in this capture. The phase values are nested according to the
measurement contract; `apply_ns` is contained by `publish_ns`, so they must
not be added. These receipts identify a bounded investigation target but do
not establish a causal bottleneck or an optimization result.

## Replay layout

`generated-fixtures.json` contains absolute paths under the original external
results directory. Replaying therefore requires restoring the raw files at
that recorded external path and using the captured source checkout; the
retained repository subtree cannot be passed directly as `--results`.
Restore all paths listed in `raw-manifest.sha256` without rewriting bytes or
the embedded fixture paths. Then use the bundle to reconstruct the source:

```sh
git bundle verify /path/to/profile-projection-reuse-463334f1/captured-source.bundle
git fetch /path/to/profile-projection-reuse-463334f1/captured-source.bundle \
  HEAD:refs/heads/docx-styles-effects-profile-projection-reuse-463334f1
git worktree add --detach \
  /var/tmp/litchi-docx-styles-effects-profile-projection-reuse-463334f1-replay \
  463334f1a32de687ffb2b3357e55816dcfd3102e
```

With the raw files restored at
`/var/tmp/litchi-docx-styles-effects-profile-results-463334f1a-20260912-a`,
run the captured verifier from that checkout and write its replay receipt in
the same external results directory:

```sh
python3 -B \
  /var/tmp/litchi-docx-styles-effects-profile-projection-reuse-463334f1-replay/docs/report/spec-gap-validation-evidence/docx-styles-effects-performance/verify_profile.py \
  --root /var/tmp/litchi-docx-styles-effects-profile-projection-reuse-463334f1-replay \
  --evidence /var/tmp/litchi-docx-styles-effects-profile-projection-reuse-463334f1-replay/docs/report/spec-gap-validation-evidence/docx-styles-effects-performance \
  --results /var/tmp/litchi-docx-styles-effects-profile-results-463334f1a-20260912-a \
  --output /var/tmp/litchi-docx-styles-effects-profile-results-463334f1a-20260912-a/replay-verification.json
```

The verifier's output-containment rule rejects a direct replay against this
retained repository directory. The original raw receipts must remain
unchanged; `replay-verification.json` is a separate receipt.
