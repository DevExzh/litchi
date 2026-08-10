# ODT consuming full-text block ownership

Date: 2026-08-11

Production base: `f12786397d3e677981ed3c441950f94fe6155d48`

Scope: ODT full-text extraction only. OLE2, OOXML and RTF production code are
unchanged, and iWork/IWA crates were explicitly excluded.

## Hypothesis

`parse_text_blocks` already created one validated `String` for each paragraph
or heading, but stored it through `Element::set_text(&str)`, which cloned the
complete string. `Elements::extract_text` then borrowed every parsed block and
called its public text accessor, creating another complete string before
copying that text into the final document string. On the 10,000-block corpus,
Heaptrack attributed two approximately 100,000-call clone buckets across ten
full-text extractions to those two handoffs.

The full-text caller owns the parsed block vector and does not need either
intermediate string after extraction. Moving both values should remove exactly
two string allocations per block without weakening the parser or changing the
structured paragraph/list API.

## Change

The shared parser now has a private ownership mode. Only
`Elements::extract_text` selects it: the parser moves its already bounded,
validated string into the block element, then the consuming block iterator
moves the first block string into the final output and appends subsequent
strings in order. The ordinary `parse_text_blocks` path retains its former
borrowing-and-cloning behavior, so public paragraph and heading ownership,
list queries and one-paragraph queries are unchanged.

The same namespace-aware event scan, nesting and visibility rules, text-byte
limit, element validation, block order and newline insertion remain. This
changes no public API, dependency, document identity, transaction, patch,
publication, limit, security policy, runtime, lock, cache or unsafe-code
boundary.

## Iteration and rejection record

The first implementation moved strings for every parser caller. Although large
full-text extraction improved, it regressed the large list query by 5.71% p50
and 5.38% mean and the large one-paragraph query by 5.30% p50. That version was
fully removed. Its raw evidence remains in
`abba-odt-full-text-move-{primary,query}-*.json`.

A narrower const-generic parser variant avoided the query regression but moved
the unchanged open guard by +5.65% p95, consistent with code-layout sensitivity
in a warm process. It was replaced by one shared runtime-mode parser. No code
from either rejected variant remains.

## Matched latency measurement

The before release executable SHA-256 is
`5ab508644d0c7cb01c47361ad076bd2d89d9d25592520dd3fd5fa30d321e0c91`;
the final after executable SHA-256 is
`290e4a34d10c6aa2e438666c82afde5ea8a5c984723e8f85496b609aa1395fa8`.
The final after executable was rebuilt after documentation-only lint cleanup;
its `.text` section matches the earlier measured-after executable and has
SHA-256
`c11bedc758da5fff5494fdedec18e76b09e312ea297a291e2b65c4279fcea1d7`
in both binaries.

Environment: release Rust 1.95.0, Linux 6.8.0-101-generic, x86_64 AMD EPYC
9575F VM, Rust system allocator and CPU 2 pinned with `taskset`. The
deterministic large ODT has 10,000 blocks, 490,000 logical text bytes, a 28,420
byte archive, and archive SHA-256
`9d724c649cb5e4b4adce30c4ede2059ff9efc26109c1b84ac8460df00ecf89a9`.
Every timed extraction is checked outside timing against the complete expected
text.

Four short ABBA cycles each used 50 warmups and 250 samples per leg. Pooling
2,000 samples per state gives:

| Large ODT full text | Before | After | Delta |
|---|---:|---:|---:|
| p50 | 4.127 ms | 3.993 ms | **-3.25%** |
| p95 | 4.486 ms | 4.232 ms | **-5.67%** |
| p99 | 4.848 ms | 4.444 ms | **-8.32%** |
| mean | 4.165 ms | 3.965 ms | **-4.81%** |

The approximate independent-sample 95% interval for the mean delta is
`[-5.14%, -4.49%]` of the before mean. Pooled leg p50s are 4.145/4.103 ms
before and 3.973/4.009 ms after, so every same-state A/B spread is about 1.1%
or less. Cycle 3 contains one fast after-A leg; the pooled result retains it
rather than selecting favorable legs.

Raw primary files are
`abba-odt-full-text-single-repeat-{1,2,3,4}-{before-a,after-a,after-b,before-b}.json`.
Their SHA-256 digests are:

| Cycle | Before A | After A | After B | Before B |
|---:|---|---|---|---|
| 1 | `ea97030c5853003d0adb45c40824f8355d8dd995a6a28470eab2d72bbbdaded6` | `9a8e2af3ab4f57008099ffa6c1ee57471029c57623f890753043698b2a1d70e0` | `d53ea6221fd2848f478ec1f214d4badf9b79e33287632d0f65e389123865bb2c` | `cd45515cc1af33e9fa02cdd817edec89481fc2100c1e588f58e6c7a23651e635` |
| 2 | `950e463bcffe873d34b949278571a6bc6bd5604dc8a4c2ddf99432931285d54f` | `df5ca590c1813dfec8215149754348fd60cc94d9b49d7d7cde4e3e768bed4bec` | `3bf76c26a47508b29091225015b3c4e7a4c98b1a5bbc57f0b01b4cfc99c90c21` | `681103e964dd5e7787543785f645199bc90fe6ce2056287f1cddb05e7209a998` |
| 3 | `8e1b889fb2c4648dde8b8a35f7ef67a15c4396217ecb75367bd583e251e82ce2` | `6f286335cb0546aa18efe23e9237a996f0f0948c2612c46d72d050f10da75518` | `9f6dcd23c82042ac0902021ada1259144457e34cba3b1687d46f89532721b4fb` | `6c7a1396deaf717d0df2e35810fa079721b7c43ffceaf3d3bf8ee0c3dfe96da7` |
| 4 | `3e8d8f2bd7f8942beedc18cd7497721bf258b4ca39a950f939dfa90f9f86921e` | `81084049b44c1ead9a57242dd331ef80afea43fb596be59fa7a8dcbd65ca8668` | `e0de20fda152055f4f00114d3273f597d2dbf79e4fe767b7ca65680f3cc60d62` | `18000b0229714a7c742bc98baf59cf782e62662f39c6f06aa6703f9dadbd9160` |

## Query, size, open and edit guardrails

Independent final four-leg runs keep the unchanged structured-query paths near
neutral and show benefits at smaller full-text sizes:

| Guardrail | p50 delta | p95 delta | p99 delta | Mean delta |
|---|---:|---:|---:|---:|
| List 10,000 blocks | -0.24% | +2.22% | +9.05% | +0.41% |
| One paragraph, 10,000 blocks | +0.02% | +0.83% | +3.32% | +0.18% |
| Full text, tiny | -1.58% | -1.69% | -7.56% | -1.20% |
| Full text, medium | -1.15% | -2.15% | -2.73% | -2.27% |
| Open, 10,000 blocks | +3.94% | +4.93% | +10.95% | +4.17% |
| Exact no-op edit/save, 10,000 blocks | -4.17% | -9.67% | -8.11% | -2.57% |
| One edit/save, 10,000 blocks | +1.30% | +1.54% | -3.75% | +1.26% |

The open case does not execute the changed full-text path. Its p50, mean and
p95 remain inside the 5% review threshold, while p99 exceeds it and is
explicitly disclosed. The warm-process result is treated as allocator or code
layout sensitivity, not as an open improvement or a reason to weaken open
validation. The one-edit mean interval is `[+0.76%, +1.76%]`; the tiny no-op
timer is only about 230 ns.

Raw reports are the `abba-odt-full-text-single-{query,open,size,noop,edit}-*.json`
files. An earlier single primary run is retained under
`abba-odt-full-text-single-primary-*.json`, but it is not the headline because
its same-state before legs drifted by slightly more than 5%.

## Allocation, counters and RSS

Matched Heaptrack processes used ten large full-text samples. Whole-process
allocation calls fall from 2,713,601 to 2,293,582 (-420,019, **-15.48%**) and
temporary allocations from 460,646 to 250,942 (-45.52%). The two full-text
clone buckets disappear; the remaining approximately 100,000 structured-text
calls come from the harness's out-of-timer semantic verification. Peak heap is
18.05 MiB in both processes. Profiler RSS moves from 32.83 to 32.94 MiB
(+0.34%), and the 1.78 KiB process-exit leak report is unchanged.

GNU Time ABBA processes used 20 warmups and 500 samples per leg. Maximum RSS is
30,848 KiB in all four legs. Matched `perf stat` ABBA processes at the same
sample count give:

| Process-wide counter | Before A+B | After A+B | Delta |
|---|---:|---:|---:|
| task clock | 14,792.23 ms | 14,438.16 ms | -2.39% |
| cycles | 72,874,069,496 | 71,364,190,295 | -2.07% |
| instructions | 265,796,477,038 | 259,120,767,647 | -2.51% |
| branches | 56,267,759,070 | 54,669,202,592 | -2.84% |
| branch misses | 39,790,753 | 41,714,843 | +4.84% |
| cache references | 4,832,330,223 | 4,538,187,966 | -6.09% |
| cache misses | 238,035,892 | 206,975,168 | -13.05% |
| page faults | 1,421,402 | 1,335,668 | -6.03% |

The sampled profile reduces exclusive `Element::get_text_recursive` share from
3.78% to 0.84%; the remainder belongs to the unchanged semantic verification.
Raw evidence is in `odt-full-text-single-*-heaptrack.txt`,
`odt-full-text-single-*-perf-report.txt`,
`odt-full-text-single-perf-stat-*.csv`, and
`odt-full-text-single-time-*.txt`.

## Correctness verification

- all-feature `litchi-odt` passed 524 unit tests, every integration suite and
  55 doctests;
- warning-denied all-target/all-feature Clippy and warning-denied crate rustdoc
  pass, including small pre-existing lint corrections in the touched crate;
- the ODF fuzz target and its production dependency graph compile offline;
- the unchanged benchmark harness passes 23 tests and warning-denied Clippy;
- focused allocation-identity and nested parent/child ordering tests pass; and
- formatting, JSON parsing and `git diff --check` are final commit gates.

Existing malformed XML, depth/size limits, namespace visibility, real-producer
packages, media and metadata preservation, signed/encrypted refusal, exact
no-op identity, patch/inverse behavior and complete changed-output reopen stay
covered. A workspace-wide gate was not rerun because iWork was explicitly
excluded while its crates are being modified independently.

## Next non-iWork audits

1. OLE2: add a dedicated PPT root slide-order benchmark before considering a
   private already-open CFB handoff; keep final-render reuse separate.
2. OOXML: benchmark `SourceBackedPackage` materialization before changing its
   private Arc-to-Vec boundary; current production callers do not exercise it.
3. RTF: expand the corpus to editable byte-1252, read-only LZFu and genuine
   formatted/media-rich producer files before another specialization.
4. ODF: attribute source-backed reads, repeated semantic scans and unchanged
   member publication independently; do not revive either rejected package
   adoption candidate without resolving its guard regression.

iWork remains deferred while the `iwa-*` crates are modified independently.
