# 0359rb — RTF empty underline refusal fast path

- Date: 2026-10-22.
- Status: measured implementation candidate; native qualification remains open.
- Scope: the ordinary RTF `set_text_underline` producer after the shared
  `Refusal` migration in 0359b.

## Change

`Edit::set_text_underline` now separates the empty-operation case from the
general staged-operation case. With no staged operation, the body-compatibility
and paragraph-layout scans are vacuous. The method still checks the operation
limit first. An empty span at byte position zero is always a valid UTF-8
boundary; an empty span at the retained semantic text length is also a valid
boundary. Both necessarily fail the existing non-empty-text rule, so they
return the same `UnsupportedContent` refusal before materializing the lazy
`Package::text` projection.

All other spans go through the existing body and UTF-8 validation. Interior
empty spans still distinguish a character boundary from a non-boundary, and an
out-of-range span still reports its range error. When operations are already
staged, the existing compatibility, operation-limit, paragraph-layout and
conflict precedence is unchanged. The shared `Refusal` representation,
operation/reason strings, error classification, allocation contract and
redaction behavior are unchanged.

The focused test covers the first and last boundaries, an interior multibyte
boundary, an interior non-boundary, an out-of-range span, a zero operation
limit, a successful non-empty ASCII edit, and a staged paragraph-layout
operation that must retain its `UnsupportedOperation` refusal.

## Deterministic instruction evidence

The comparison uses the same current workspace source with and without this
RTF-only change. Each side was built in a separate external target with one
cargo job using:

```sh
cargo build --offline --locked --release -p litchi-rtf --example refusal_contract_bench
```

Callgrind 3.26.0 ran the retained 0359b probe once on CPU 7 with
`taskset -c 7`, `--tool=callgrind --dump-instr=yes --collect-jumps=yes` and
`Ir` collection. This is deterministic instruction evidence, not native timing
qualification. The exact retained probe workload and dispositions are the
eight 0359b scenarios: open, no-op, text commit, text commit plus save,
replay, tail commit, unsupported content, and protected commit.

| Build | Binary size | Whole-probe `Ir` | `set_text_underline` direct `Ir` | Calls | Direct `Ir`/call |
| --- | ---: | ---: | ---: | ---: | ---: |
| Current source | 4,145,672 B | 627,438,818 | 1,496,000 | 34,000 | 44 |
| Empty-boundary fast path | 4,145,800 B | 626,418,435 | 476,000 | 34,000 | 14 |
| Delta | +128 B | -1,020,383 | -1,020,000 | 0 | -30 |

The profile's other RTF edit/publication blocks are identical between these
two builds: `Edit::replace_text` 569,279 `Ir` over 3,469 calls,
`Edit::commit` 268,167 over 6,869, `Edit::commit_staged` 3,553,967 over
6,869, tail-append commit 14,892 over 34, `Package::to_bytes` 2,652 over
34, and `Package::apply` 1,159,400 over 3,400. The reduction is therefore
localized to the refusal producer in this matched profile.

The producer's static symbol grows from 0x199 (409 bytes) to 0x1ec (492
bytes), a measured 83-byte code-size increase. The new entry branch is the
cost of making the empty refusal route cheap; the matched normal producer
route was checked separately before accepting this tradeoff.

## Successful-path controls

A private no-inline probe used the same source and a non-empty ASCII span. Its
source is retained outside the checkout at
`(deleted-coord-artifact)`.
The cold case parsed 34 fresh packages and then made one successful underline
commit per package. The warm case precomputed `Package::text()` and made
34,000 successful underline commits. The no-inline `success` wrapper contains
the edit and commit route:

| Case | Current wrapper direct `Ir`/call | Fast wrapper direct `Ir`/call | Whole-process `Ir` current → fast |
| --- | ---: | ---: | ---: |
| Cold, lazy text per package | 200 | 200 | 2,639,005 → 2,658,053 |
| Warm, text already materialized | 194 | 194 | 1,581,012,112 → 1,576,402,637 |

The wrapper cost is unchanged in both states. The different whole-process
totals include parsing, driver and code-layout effects and are not used as a
native performance claim. The exact 0359b normal edit/publication controls
also remain instruction-identical above.

## Correctness and qualification

`cargo fmt --all -- --check` passes. The focused edge test and the complete
`cargo test --offline --locked -p litchi-rtf` run pass, including 380 unit
tests, all RTF integration suites, and the RTF doctests. No native A-B-B-A or
heaptrack run was made on this candidate while the host was busy.

Before owner acceptance, rebuild the before/after release binaries in isolated
external targets, run the retained three-trial CPU-7 A-B-B-A driver with the
0359b sampling policy, and repeat paired heaptrack captures. Retain the native
JSONL outputs, process monitor, binary hashes, heap summaries/captures and the
complete RTF correctness, corpus/fidelity and glossary logs. The candidate is
not a native timing waiver.

The deterministic profile and private probe are outside the repository:

| Artifact | Size | SHA-256 |
| --- | ---: | --- |
| Current-source release binary | 4,145,672 B | `b9fccd840476aac1beaa35cb5e14d5410f6272493dab33e361d83fb4d9aed8aa` |
| Fast-path release binary | 4,145,800 B | `126db1632fe150eae900782a98906933dec9aa3945542a0e1a99d22ae4758157` |
| Current refusal Callgrind profile | 516,802 B | `e08d62defe8159505062639519ec22eac75aa8d497bcc4673ac352e3b53ef005` |
| Fast-path refusal Callgrind profile | 516,817 B | `fa11171db686a8873b9c660be9f823606ebdec0d94487b83c95ca8f02676ed4a` |
| Private successful-route probe | 1,177 B | `c2f7ec5c3f121cabf3b143bd459cbed9dbfebe08a17e0ef7eca1520331a3a618` |
| Cold current successful-route profile | 442,114 B | `e186f698e4efe1e79fe95ef1fc97baa793d8f393a173d143ea90b1169f77643f` |
| Warm current successful-route profile | 506,591 B | `3ba21cd59824e2b2b7c3c81f4a530687f95700ec8a1fe7d344be028e703799e0` |
| Cold fast successful-route profile | 442,712 B | `1f546ceee65c1f904ee37c157227ae591ced23fa5b55d6986580bcd4e7018af9` |
| Warm fast successful-route profile | 507,304 B | `6de10246ec5a324739658036b9c226c37854a57e2d438893bf6a634a2b403081` |
