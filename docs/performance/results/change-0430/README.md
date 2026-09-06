# 0430: recover PPTX publication CPU attribution

Two fresh processes profile the unchanged 0429 release executable with frame
pointers, resolving the publication callers that the earlier 16 KiB DWARF
capture often omitted. This is an attribution batch. It changes no production
or benchmark Rust code and establishes no performance improvement.

Each process uses the synthetic media-rich PPTX cross-copy corpus, CPU 2,
`cycles:u` at 499 Hz, three warmups and 100 retained lifecycle samples. Providers
are bytes and a recently staged, warm file. Corpus construction and gates are
inside the full-process capture; API clocks retain their original narrower
scope. Frame-pointer stacks cannot distinguish warmup iterations from retained
iterations. Raw perf data, decoded stacks, reports, command receipts, source
hashes, and machine records are retained with lossless compression.

The executable SHA-256 is
`9b3fca778443b6c9bd49a8522f56a251f1df27fc775a76af6124186398f78969`, built from
`56ec912e5b387e8f4f55f7044a73bb4cdd11e324` using Rust 1.98.1, release debug
information and `-Cforce-frame-pointers=yes`. The copied build receipt and log
retain that historical build evidence. No rebuild was performed for this batch.

The source manifest covers Rust, TOML and lockfiles. It omits compile-time
non-Rust assets, including embedded RTF/XLS fixtures and resource templates,
so this is not a hermetic rebuild bundle. Recapture requires the preserved
workspace and asset tree as well as the declared toolchain, and the resulting
executable must match the pinned hash. The copied 0429 build record also names
historical helpers and a validation amendment that are not all copied here;
those references remain metadata pins. This replay verifies the relevant copied
build receipt/log, source manifest, binary identity and new captures. Replaying
the whole 0429 validation history requires its original bundle.

See [derived counts](profile-summary.json), the
[source boundary audit](source-audit.md), [transfer design](transfer-design.md),
[ADR review](adr-review.md), [verification](verification.md), and
[remaining goal scope](goal-scope.md).
The result prioritizes avoiding Deflate of unchanged copied media. It does not
authorize bypassing decoded validation or transferring unchecked archive bytes.

Replay from an exported copy of this directory:

```sh
python3 -B verify.py --portable-check
```

The replay needs Python and the retained bundle, not the executable, source
checkout, original perf installation, or task temporary directory. It verifies
capture custody, decoded reports, stack-derived counts, and rejection of
mutated evidence. Stack replay uses the retained symbolized text; independently
resymbolizing the raw recording would also need the original executable and
its referenced debug information. `SHA256SUMS` inventories stored artifacts; `compression.json`
binds each compressed capture/log to its original bytes.

The original capture command, requiring the hash-bound executable and the
external workspace/asset inputs described above, was:

```sh
python3 -B docs/performance/results/change-0430/check.py --tag frame-pointer-profiles -- python3 -B docs/performance/results/change-0430/capture-profiles.py
```

The capture driver uses exclusive output paths. Preserve the retained bundle
and use a fresh destination for a new experiment; do not rerun this command
against the retained directory. Its exact commands and
start/finish times are recorded in each profile receipt. The temporary capture
executable is removed only after portable replay; the original build binary,
both existing target directories, and the user's `docs/GOAL.md` are preserved.

No before/after latency, RSS, allocation, copied-byte, cold-I/O, native
cross-copy, or scaling claim follows from these two profiles. Deflate includes
changed XML/relationships as well as copied media, so its complete sampled
share is not an estimate of work a media-transfer change will eliminate.
Normal matched captures and preservation/adversarial tests are required before
keeping that change. Existing strict lint debts and the broader goal remain
open. Rust checks from 0429 are prior evidence; this unchanged-source batch
runs evidence verification rather than repeating those builds and tests.

The [source review](checks/source-review.md) records the rebuild and machine
provenance limits. Current machine metadata includes platform, uname and
affinity, with a hash of the prior hardware/toolchain record; it is not a new
complete CPU/perf-version census. Decoded sample totals are checked against
`perf record`; a zero-lost-sample report field is not an automatically checked
invariant. The two failed replay-checker attempts and their original scripts
are retained in [validation corrections](validation-corrections/README.md).
