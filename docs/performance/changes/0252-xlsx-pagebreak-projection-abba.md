# Change 0252: XLSX page-break projection ABBA evidence

Date: 2026-08-21

Status: latency-supported on the listed cells; production landing remains
pending an independent code/oracle audit and no broad XLSX claim is made

## Decision

Retain the current-schema ABBA package as narrow evidence for the two fixed
media-rich XLSX page-break edit selectors. The eager selector accepts p50,
mean, p95, and p99. The source-backed selector accepts p50, mean, and p95;
its p99 is excluded because the A2-to-B2 direction is a `-0.30069801395129325%`
regression. All accepted cells pass the declared drift ceilings, and no
adverse-both cell is present.

This result does not land the page-break projection production candidate.
An independent audit must still review the production cache/projection
correctness and the benchmark oracle/expected-output construction. Latency
evidence alone is not approval for landing.

## Evidence-package identity

The strict current-schema package is retained outside the repository at
`/home/zhuhe/CodeProjects/litchi-perf-evidence/0252-xlsx-pagebreak-projection-abba/`.
The earlier schema-incompatible attempt remains preserved separately at
`/home/zhuhe/CodeProjects/litchi-perf-evidence/0252-xlsx-pagebreak-projection-abba-legacy-schema/`.

The package has change ID `0252-xlsx-pagebreak-projection-abba`, schema
version 1, and was summarized by `litchi-perf-abba-summary` 0.1.0. The
summary and manifest identities are:

| File | SHA-256 |
|---|---|
| `summary.json` | `27dd2f2b05213cbca75891c59fdeb31a6382aa9b40568a427abcb441decdde9b` |
| summary canonical JSON | `cd771a639cac959be197bd992e162300aafcfd03a2ee4c7ea7d88edeafd65697` |
| `0252-xlsx-pagebreak-projection-abba-manifest.json` | `00c86ebf08cef60aad3bdb3157e4e8e3127fa98663fe1611cc64ad99935889a0` |

The frozen release binaries were built sequentially into one explicit
external target with `CARGO_INCREMENTAL=0`, jobs 2, and `--release --locked`:

| Role | Revision | Binary SHA-256 |
|---|---|---|
| Control | `526ea52fdb39c90d9fc2ea07fd2c837ba84aee41` | `53d97cfffa308de816957367d00ef49be7d7a6e8529a15270025dbe2a1b16d49` |
| Candidate | `e619debe0a7b61ea24d76f03341ced6110245888` | `a0dbf8a00a90e605e0517bfcb2890882f0bcffe0694424db32a02c9ec7983523` |

The clean matched worktrees were based on `172be4c966224fd0b36dd39e24207ff6c7d97579`.
The control range is `172be4c..526ea52f`; the candidate range is
`172be4c..e619debe`. The current-schema tool identity in every leg is:

```text
name=litchi-perf-baseline
version=0.1.0
binary=litchi-perf-baseline
profile=release
target_os=linux
target_arch=x86_64
instrumentation=none
git_worktree_dirty=false
```

The package's report identities, raw JSON hashes, and compressed-member hashes
are retained below. The raw reports were not edited before compression.

| Role | Canonical report SHA-256 | Raw JSON SHA-256 | zstd member SHA-256 |
|---|---|---|---|
| `a1` | `39b13467929403144112ea2b48a2a10a4aee41e587941c496938cda975283def` | `f7e03c843d9331c65f7fca9ce13bd6a39a468049f3c72ab98054208d9de29fd6` | `a74969d7628afe48ed3156a5ec8ca312c6f8fc3a154df84cb95b317a025ded55` |
| `b1` | `422b68aaf8c5b1772fde8c762b3eaf331ac4162df595280f708ad802084920b8` | `5be5d22e80ebc7ee6755a4e94ff4514412b2429de168259337560d03197506af` | `91a10aaebca5476df24ab7c92f63ea9744c0163aa04cfa67675f5a67044d5783` |
| `b2` | `280d037dd7dd18c7afacdd28e00bf2995483dc4412697d1c51db464979ea6114` | `7c2a23f3fc430f80ba90aa549b4a5f0b06d3a1d9a69c2e33762bf48c55fdcd29` | `85a977665cee9a543c37ae7be06b84819ef7e37a4b2be5ed07105e8c13cd2de8` |
| `a2` | `65c29382a7bf999195c59e1263126cd2dd085ca0189e96f82d69f676b97be394` | `66902d92ac42db437dfec39aac60c3b5e2a38f5e9f3211289217475374f17770` | `cf9c33fd97e7a6632ced78b9094671c45ab7eebbbf594908dcf267df7139025b` |

The package uses zstd 1.5.7, level 3, one compression thread, and executable
SHA-256 `fe50fd600cad89b775dc5b4b10bc8ff95e2d33ddb4f7632966464b1fbdab8598`.
All four compressed members passed integrity testing.

## Corpus and current-schema identity

Both selectors use the exact fixed corpus identity below; no CLI shape is
substituted for the fixed page-break corpus:

| Field | Value |
|---|---|
| Corpus name | `xlsx-page-break-media` |
| Generator | `litchi-xlsx-page-break-source-edit-media-v1` |
| Shape | `media-rich` |
| Package format | `XLSX/OPC/ZIP` |
| Archive bytes / members | `16,786,830 / 17` |
| Archive SHA-256 | `c11a9424accfc6ce56e4deb6ecb18a2142d2f0076395018ef00ba93897049f7c` |
| Uncompressed payload bytes | `16,782,412` |
| Target | `worksheet:Sheet1:rowBreaks` |
| Target payload bytes | `207` |
| Target payload SHA-256 | `c78af5ae31d6f622b0a9544bd9476c3f6f370258727219d9f62fb27b4e66d91f` |
| Output SHA-256 | `1e3b7a9f763feaed4ad4888aa8aa0cd3773cdb9fd9f12e16f3c05b7fd0cd95b3` |

The current-schema summary verified equal configuration, stable environment,
tool identity, corpus identity, output identity, source/sink identity, and
statistics recomputed from all retained samples. The fixed output hash is the
same for both selectors and all four legs.

## Protocol and selector scope

The order is A1 control, B1 candidate, B2 candidate, A2 control. Each leg was
pinned with `taskset -c 3`, using 30 untimed warmups and 500 retained samples
per result row. Setup, corpus construction, expected-output preparation, and
semantic/preservation checks are outside the timed operation. The exact
opt-in selectors are:

| Selector | Corpus |
|---|---|
| `xlsx_eager_page_break_edit_save` | the fixed `xlsx-page-break-media` archive |
| `xlsx_source_backed_page_break_edit_save` | the fixed `xlsx-page-break-media` archive |

The declared shape configuration was the current harness default set:
`--shape tiny,many-small,few-large,wide-root`,
`--payload compressible,incompressible`,
`--writer-shape tiny,large,payload-heavy`,
`--xlsx-shape tiny,medium,dense-wide`,
`--xlsx-cell-crud-shape medium,dense-sparse`,
`--xlsx-row-visibility-shape medium,large`,
`--semantic-shape tiny,medium,large`, `--rtf-variant plain`, and
`--workers 1`. These declarations do not replace the fixed page-break corpus.

Drift ceilings are 5%/5%/10%/15% for p50/mean/p95/p99. Positive readings
below mean lower candidate elapsed time; negative readings mean the candidate
was slower.

## Exact paired latency readings

Values are candidate reduction percentages, A1→B1 / A2→B2.

### `xlsx_eager_page_break_edit_save`

| Statistic | A1→B1 | A2→B2 | Verdict |
|---|---:|---:|---|
| p50 | +2.024529665205901% | +1.999411243992957% | accepted |
| mean | +1.9210802948498922% | +1.9691631253229647% | accepted |
| p95 | +1.5684237457025247% | +1.7101234030945538% | accepted |
| p99 | +1.204790864679875% | +1.255593150997107% | accepted |

All four eager statistics are accepted; their reductions are approximately
1.2%-2.0% in both paired directions.

### `xlsx_source_backed_page_break_edit_save`

| Statistic | A1→B1 | A2→B2 | Verdict |
|---|---:|---:|---|
| p50 | +0.5918311791543713% | +2.570460237130754% | accepted |
| mean | +0.9941718286289549% | +2.258279975364421% | accepted |
| p95 | +0.7850929439146723% | +1.3562361375900667% | accepted |
| p99 | +7.220226598577385% | −0.30069801395129325% | excluded: direction-disagree |

The source-backed p99 is excluded rather than averaged into a claim. No
adverse-both cell is present, and all candidate/control drift checks pass.

## Claim boundary and verification

This note covers only the two listed fixed-corpus page-break edit selectors.
It makes no claim about other XLSX shapes, worksheets, page-break layouts,
physical I/O, decompression, filesystem or cold-cache behavior, allocations,
RSS, real-world producers, durable output breadth, or general XLSX CRUD.

The measured production candidate remains pending an independent code/oracle
audit and is not landed by this evidence-only commit. In particular, the
latency package does not by itself establish that the production projection
cache is correct for every snapshot/relationship transition or that the
expected-byte oracle is sufficiently independent for a production decision.

Verification for this documentation-only record is package/hash inspection,
compressed-member integrity testing, current-schema summary validation,
Python summarizer tests, and `git diff --check`; no Cargo command or benchmark
is part of this commit.
