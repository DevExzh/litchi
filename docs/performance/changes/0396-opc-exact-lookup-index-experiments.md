# Change 0396: OPC exact-lookup index experiments

Date: 2026-09-04

Status: benchmark coverage accepted; all production experiments rejected. No
0396 production implementation or performance claim is retained.

`performance_claim: none`

`claim_authorized: false`

## Benchmark coverage

The existing OPC case-fold coverage expands from four to seven opt-in
selectors: this change adds three class-isolated source selectors for exact,
ASCII-case-alias, and genuine-miss lookup:

```text
opc_casefold_eager_open
opc_casefold_source_open
opc_casefold_eager_lookup
opc_casefold_source_lookup
opc_casefold_source_exact_lookup
opc_casefold_source_case_alias_lookup
opc_casefold_source_genuine_miss_lookup
```

The corpus coverage adds 2,047 Parts to the prior 256-, 2,048-, and 16,384-
Part deterministic stored OPC corpora. The 2,047-Part corpus is the boundary
control immediately below the 2,048 source case-fold threshold. Each ordinary
member has a 32-byte payload; the fixed structural members are
`[Content_Types].xml` and `_rels/.rels`.

The combined lookup selector opens once outside timing and repeats a fixed
nine-query vector 16 times (144 lookups): exact first/middle/last, case-only
aliases at those positions, and genuine first/middle/last misses. The three
source-backed class selectors independently repeat their exact, alias, or
miss positions 48 times, also producing exactly 144 lookups each. Query
classes, canonical positions, found/missing outcomes, output digests, and
the malformed equivalent-name gate are fixed correctness oracles. An
independent source replay reports source/version/payload counters and performs
no ordinary payload reads.

The additions raise the selectable registry from 415 to **418**. The default
matrix remains **36 cases / 198 rows**. This is benchmark coverage, not CRUD
semantic coverage and not a recommendation to change the production index.

## Timing and allocator boundary

Every latency delta in the experiment table is derived from normal,
non-allocator release-binary p50 evidence. Source-open measurements
time normal unmanaged `SourceBackedPackage::from_read_at`; lookup measurements
time fixed pre-open unmanaged packages. The paired values in the experiment
table are in the order **2,048 Parts / 16,384 Parts**; raw ABBA reports retain
the complete per-leg samples. Allocator-enabled elapsed time is observational
only: that target exists to measure allocation vectors and cannot authorize a
latency claim. Validation-constructor coverage is correctness-only. The 256-
and 2,047-Part cases establish small-catalog and threshold boundary coverage,
but no timing claim is made for them.

The retained bundle spans two compiler identities: the eight earliest
mapless reports identify `rustc 1.95.0`, while the other 40 identify
`rustc 1.98.1`. The final Arc gate and final harness QA selected
Rust/Cargo/Rustdoc 1.98.1 explicitly because the repository-pinned 1.95
installation does not include Cargo. Every report binds its exact compiler,
binary hash, revision, and CPU identity.

The control is revision `c0ca6cb5f22ddc68d827b743018855f6b9dc89bd`.
The final pooled `Arc<str>` experiment is revision
`8f7714ee011b170d938f2532fdd385fb2b61cd32`. The [0396 evidence bundle](../results/change-0396/)
retains the candidate reports, catalogs, allocator observations, and
adjudication for all experiments.

## Rejected production experiments

Every candidate below was evaluated against the existing exact-name and
case-fold lookup behavior. The normal-binary p50 results are reported even
though they are rejected; this prevents the benchmark record from turning a
failed optimization search into a success claim.

| Candidate | Exact lookup p50 delta — normal, non-allocator release-binary p50 (2,048 / 16,384) | Other measured evidence | Decision |
| --- | ---: | --- | --- |
| Mapless exact lookup | approximately `+2,750% / +3,500%` | Removing the exact-name map routes exact queries through the case-fold ordering. | Rejected: severe exact-lookup regression. |
| Preliminary scalar-`Vec` exact-position probe | `+14.85% / +20.96%` | Five-sample probe before the final inlining/layout iteration. | Rejected: exact-lookup regression. |
| Inlined linear-probe `Vec`, full ABBA | `+20.22% / +12.42%` | Saves `N` source-open allocator allocation calls, with `N` equal to the ordinary Part count. | Rejected: allocation saving does not offset exact-lookup latency. |
| `std` prehashed exact index | `+13.42% / +15.88%` | Prehashing does not recover the exact-name fast path under the measured protocol. | Rejected: exact-lookup regression. |
| Direct `HashTable` exact index | `+14.66% / +13.36%` | A higher-sample follow-up remained approximately `+14.7%`–`+15.6%`. | Rejected: repeated exact-lookup regression. |
| Final pooled `Arc<str>` PackURI storage | `+6.09% / +6.96%` | Source-open p50 `+3.38% / +4.30%`; mixed exact/alias/miss lookup `-0.59% / -0.50%`. Allocator source-open vectors add 3 allocation calls and approximately `N` deallocation calls; net-live bytes fall by 65,536 / 524,288 at 2,048 / 16,384 Parts. | Rejected: lifecycle and latency regressions outweigh the retained-live-byte observation. |

The `Arc<str>` result is exact allocator and net-live footprint
evidence: its net-live observation is not an RSS, total-memory, or
system-footprint claim. Its additional deallocation work is part of the
lifecycle regression. No candidate was promoted into the production API or
default behavior.

## Correctness and reproducibility

The exact, alias, miss, iteration, canonical-name, source-counter, and
equivalent-name oracles remain independent of elapsed time. The corpus and
selector additions were checked with the existing harness validation, and
the normal and allocator report families are linked in the evidence bundle.
The control and final pooled candidate use the same fixed vectors and corpus
shapes, so the rejected results are comparable benchmark evidence even though
they are not production claims.

No claim follows for eager or managed packages, mutable `OpcPackage`, source
validation-constructor latency, RSS, total memory, physical I/O,
decompression, cold-cache behavior, throughput, scaling, or general OPC/
OOXML behavior. Allocator-enabled latency is observational only, and all
normal-binary latency results above remain rejected experiment observations.
