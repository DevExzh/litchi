# 0945 retained evidence

See the [report](../../0945-xlsx-shared-replacement-payloads.md). The candidate
is adopted after all frozen gates pass. The compact `summary-v1.json` preserves
paired p50/allocation ratios and confidence bounds; the independent raw audit is retained
under `comparison-tooling-review/`. Source changes in the checkout match the
frozen candidate exactly.

The complete packet is `/home/zhuhe/litchi-performance-0945`. Large source
manifests, raw reports, the full 5,439,340-byte `analysis.json`, source trees,
and executables remain outside Git. Every committed evidence file is smaller
than 1 MiB. Binaries are identified by source, build command, size, and SHA-256
in the plan and build receipts; no compiled executable is committed.

The verified archive is
`/home/zhuhe/litchi-performance-0945-archive/litchi-performance-0945-v1.tar.zst`
(**782,627,603 bytes**, SHA-256
`a7bd7f729acf633483e6b5cfd639deb7f8a46db24f91ada6484fdbf382ef07ea`).
It contains 76,730 files totaling 3,226,301,466 bytes. `archive-summary.json`
records the 13,944,297-byte external inventory's hash, exact extraction checks,
cleanup, and the independent moved-replay result. The latter is 1,792,992 bytes
and remains external, as does the 4,037,192-byte cleanup path listing; their
digests and reduced results are committed. Archive and replay supplements were
produced after sealing and are identified separately.

## Offline replay

Extract the archive into an empty directory while preserving recorded file
modes (for example, set `umask 000` in the extraction subshell). Change into the
extracted `litchi-performance-0945` directory, then run:

```sh
python3 -B tools/summarize_0945_v1.py --output /tmp/0945-summary-replay.json
python3 -B comparison-tooling-review/independent-post-capture-v1.py \
  --packet . --plan plan-v2.json --reducer-output analysis.json \
  --output /tmp/0945-independent-replay.json
```

Use fresh output names: helpers refuse replacement. These commands read retained
evidence and do not run benchmarks. The compact summary must match
`summary-v1.json` byte for byte; the independent auditor recomputes raw statistics,
semantics, and custody from the extracted packet. Original and moved audit
projections matched exactly.

The frozen main reducer is **original-root-only**. It rebuilds historical argv
using its current packet root and writes absolute folder labels, so it fails
invocation custody after relocation. Both failed moved invocations are retained
in `moved-main-reducer-failure-v1.json`; no helper, plan, or workload was changed.
At the original packet location, reproduce the full reduction from the packet
working directory with `python3 -B tools/reduce_0945_v2.py all --output
/path/to/new-analysis.json`. The working directory also resolves its semantic
helper import. Use the independent auditor above for supported relocated replay.

This directory holds bounded review copies, not the complete runnable packet.
External `tools/` and receipt layout are preserved where copied. The original
`plan.json`, v1 helpers, failed empty namespace receipt, and superseding
`plan-v2.json` remain distinct. `METHOD.md` describes the historical v1 helpers;
use the v2 commands below. No source, threshold, or schedule changed in that
prelaunch correction.

## Regeneration

The selected Git source export, source manifests, candidate freeze, patch, and
build receipts are retained in the external archive. Export base commit
`e037befa6ce52b3988c8ae2032ae652196f2a09d` using the selection recorded in
`source-export.json`; the candidate patch is `candidate-v2-portable.patch`.
Use a fresh external packet, the recorded Rust 1.95.0 toolchain, standalone
tracked lockfile, environment, and build commands. The baseline-native reuse
receipt proves exact selected source equivalence to the 0944 build; its copied
provenance is in `references/0944-baseline-native/` in the full packet.

The exact capture order, after preparing and freezing a fresh packet, is:

```sh
python3 -B tools/capture_0945_v2.py check
python3 -B tools/capture_0945_v2.py qualification
python3 -B tools/reduce_0945_v2.py qualification
python3 -B tools/capture_0945_v2.py native
python3 -B tools/capture_0945_v2.py allocator
python3 -B tools/reduce_0945_v2.py all
```

Run from the packet root. Existing capture namespaces cannot be reused.
Qualification must pass before formal capture. `plan-v2.json` fixes every
case/order/count, CPU, environment, source/build binding, and admission gate.
`tools/summarize_0945_v1.py` regenerates the bounded summary from full reduction
and custody records; its help lists input/output arguments.

For the intact original packet and a fresh archive destination, regenerate the
archive with the retained script and commands:

```sh
LITCHI_PERF_0945_ARCHIVE_GO=1 python3 -B archive-0945-v2.py --go --archive
LITCHI_PERF_0945_ARCHIVE_GO=1 python3 -B archive-0945-v2.py --go --verify
LITCHI_PERF_0945_ARCHIVE_GO=1 python3 -B archive-0945-v2.py --go --cleanup
```

The archive script binds the external five-binary descriptor, checks all frozen
source and quality inputs, and verifies exact extracted file sets and hashes.
Cleanup removes only the seven named build/quality cache roots while preserving
the three bound release binaries inside them, then removes the verification
extraction. It does not filter arbitrary nested `target-*` fixture names.
The exact tar regeneration command and artifact descriptors are retained in
`archive-summary.json`.

## Retained limitations

The report records the first quality matrix's failures, three omitted fixtures,
two test-only candidate corrections, unchanged OPC formatting drift, and the
ignored external ZIP64 source test. Quality summary recovery and the reducer's
working-directory failure are offline corrections; no benchmark was rerun.
Original preparation/review shortcomings remain in versioned receipts.

The measured allocation saving coexists with a 1.21–1.46% workflow p50 increase.
Gross requested bytes, live bytes, peak memory, and process RSS are distinct.
The strict historical claim registry does not validate this batch's numbers;
the raw independent audit and offline replay provide that validation.
