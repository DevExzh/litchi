# Change 0250: ZIP ordering/index ABBA evidence

Date: 2026-08-21

Status: rejected on latency evidence; do not land the production
monotonic-offset sort fast path; retain the candidate only

## Evidence-package identity

The independently audited package is retained outside the repository at
`/home/zhuhe/CodeProjects/litchi-perf-evidence/0250-zip-ordering-20260821/`.
Its manifest is `0250-zip-ordering-abba-manifest.json` with SHA-256
`84b67141fa08a1ea1af7f398501fcf1ee58874bc121b6e34826858f153b23bde`.
The `summary.json` file SHA-256 is
`aa511946b050e92461645c9d8e57ed2ca6cca8c827b2d5776006efbc6f25cf1c`; its
canonical SHA-256 is
`7772f973341f4f8b941a7e7381bbce8e73df8438b5f0f44174f981573ee42732`.
The package is schema version 1, has change ID `0250-zip-ordering-abba`, and
contains eight result rows from `litchi-perf-abba-summary` 0.1.0.

The retained package members have these identities:

| Member | File SHA-256 | Canonical SHA-256 |
|---|---|---|
| `0250-zip-ordering-abba-manifest.json` | `84b67141fa08a1ea1af7f398501fcf1ee58874bc121b6e34826858f153b23bde` | — |
| `summary.json` | `aa511946b050e92461645c9d8e57ed2ca6cca8c827b2d5776006efbc6f25cf1c` | `7772f973341f4f8b941a7e7381bbce8e73df8438b5f0f44174f981573ee42732` |
| `a1.json.zst` | `fe2bfbed510c3d53b09ba9965e1dd00813bd9ebecd647bce27860a7d62dd3a9a` | `3646d55ea5c6eeafe00e7d2015f8a6c15b09a57f4343cee222a6fbff70fa3b35` |
| `a2.json.zst` | `65c1130aa6a22a89ff50626effa4c13dd9be695abdc9dcdbe1da80b0d58c64d0` | `9b604ac9d71afee8cb3dfda38b5a2a6ae01657dc2c7a0c6e12673dfa2bbcfc9b` |
| `b1.json.zst` | `361df1f51a98a3eb90106c69745dd3776a19abb6f0f90b6efe7bbedf76def686` | `be87544d53441bf08f779016f5cafb6f4ac0ab38dff88379e967cfa67c321cc9` |
| `b2.json.zst` | `da934f1f66366f709b4a5c0f7a3b0846623804eb3a4a78dfd81ff1635fe49e7e` | `34d57fb2ec8e182329070bff9639820bad62578f5133c45cfe46bffed2937774` |

The control revision is `fd814d9e7cfed0451455de016bb33487acba89a3`
(A1/A2); the candidate revision is
`1a52188c6d6aa58afd19baea722c0a1ac59727a1` (B1/B2). The frozen control and
candidate binaries are identified by SHA-256 as
`2418519d870ef6026b2dd579566dafabbc94ccbc97df7b54cb7e3c630d6d7705` and
`b5a3abb3f514fff02ae583dec1ef8c744ee2961f001cac37a370ad82bbe39bdc`,
respectively. Source, corpus, configuration, environment, and statistic
recomputation identity checks pass; output, sink, and operation metrics are
consistently absent because this is index-only evidence. The recorded legs
have clean worktrees, one visible logical CPU, Rust 1.95.0, and the AMD EPYC
host.

## Protocol and selector scope

The order is A1 control, B1 candidate, B2 candidate, A2 control on CPU 2,
with 30 warmups and 500 retained samples per row. The single selector is
`zip_index`, covering the default four shapes × two payload kinds: `tiny`,
`many-small`, `few-large`, and `wide-root`, each with compressible and
incompressible payloads. The corpus generator is
`litchi-opc-synthetic-v2` and the package format is OPC/ZIP.

This is a very short in-memory selector. Across the eight rows, retained leg
p50 values range from 501 ns to 39,522 ns, and retained leg means range from
504.444 ns to 38,487.018 ns. Drift ceilings are 5%/5%/10%/15% for
p50/mean/p95/p99. Positive readings would mean lower candidate elapsed time;
they are not treated as a claim unless the paired-direction and drift gates
pass.

## Exact latency classification

No statistic is accepted in any row. The complete counts are:

| Statistic | Accepted rows | Adverse-both rows |
|---|---:|---:|
| p50 | 0 | 5 |
| mean | 0 | 3 |
| p95 | 0 | 3 |
| p99 | 0 | 3 |

The complete paired readings below are A1→B1 / A2→B2, in percent lower, in
the order p50 / mean / p95 / p99. The exact source rows remain in the
retained `summary.json` file.

| Corpus row | A1→B1 (p50 / mean / p95 / p99) | A2→B2 (p50 / mean / p95 / p99) | Adverse-both |
|---|---|---|---|
| `tiny` / compressible | −9.803921568627452 / −32.140001508636914 / −68.39186691312385 / −65.40447504302927 | −0.19607843137254902 / −2.4970961708829424 / −1.8484288354898337 / −3.5650623885918007 | all four |
| `tiny` / incompressible | −9.780439121756487 / −25.12224893381062 / −23.25581395348837 / −18.963337547408344 | +1.7647058823529411 / +1.0047805574962243 / 0.0 / +1.607142857142857 | none |
| `many-small` / compressible | −13.935058559506277 / −21.892554553608903 / −27.112799374983158 / −26.867659522944663 | −0.6006744414781509 / −0.5015633685848142 / −4.542300507072325 / −0.32893958397425277 | all four |
| `many-small` / incompressible | −39.029795616843145 / −33.39193698505037 / −46.61173287418522 / −52.742108079186735 | +0.3854509776438433 / +2.012069236646559 / +10.404428082817681 / +9.650088028169014 | none |
| `few-large` / compressible | −3.502626970227671 / −2.417072107299481 / −0.16129032258064516 / 0.0 | 0.0 / +5.052190839495223 / +1.5847860538827259 / +32.597266035751844 | none |
| `few-large` / incompressible | −7.130124777183601 / −5.587999098312055 / −4.9916805324459235 / −1.5847860538827259 | −1.7513134851138354 / +0.22776243046756978 / +1.5847860538827259 / +30.401737242128124 | p50 |
| `wide-root` / compressible | −4.5808277325787055 / −5.139887445559329 / −7.752751054368854 / −12.70832906883917 | −2.7682540291450684 / −3.15942208981756 / −5.581405844522171 / −7.83136945489978 | all four |
| `wide-root` / incompressible | −5.131238729498421 / −4.65437010759153 / −4.122773247471389 / −2.909540539891638 | −1.0006786642016334 / +1.6622540795821175 / +30.51790683062562 / +12.515844645120664 | p50 |

The broadly adverse rows are `tiny`/compressible,
`wide-root`/compressible, and `many-small`/compressible, each adverse-both
at all four statistics. The remaining rows are rejected by paired-direction
disagreement and/or drift gates; none supplies an accepted latency cell.

## Oracle and decision boundary

The independent ZIP-index count oracle was shared by the control and candidate
measurements. It landed separately as
`5eb8c1959490d7bc3f596b8765bb855c3f28dd08`, with the test-only bound fix in
`197bd3645d398c055c64c1b7122883b848be655a`. Those commits retain the actual
observed member count and reject manifest-count drift; they are correctness
and evidence-harness changes, not approval of the production fast path.

Do not land the production monotonic-offset sort fast path on this evidence.
Retain the candidate only. The rejected result does not support a broad ZIP,
OPC, or archive-ordering performance claim.

## Claim boundary

This is narrow in-memory `zip_index` selector evidence only. It makes no claim
about physical I/O, device traffic, filesystem or cold-cache behavior,
decompression, allocations, RSS, source-backed reads, output identity, sink
identity, operation metrics, or a general ZIP/OPC performance improvement.
The short selector timings and shared count/digest oracle do not turn this
rejected latency run into production approval.

Verification for this documentation-only update is Python package/hash/count
and link checking plus `git diff --check`; no Cargo command or benchmark is
part of this record.
