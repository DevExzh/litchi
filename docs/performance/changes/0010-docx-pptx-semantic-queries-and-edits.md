# Change 0010: DOCX/PPTX semantic queries and repeated edits

Date: 2026-08-10

## Decision

Accept three narrow changes behind new end-to-end semantic evidence:

- DOCX one-paragraph lookup scans the complete bounded XML for validation but
  constructs only the requested shared range instead of materializing the
  complete paragraph collection. The source-backed paragraph count now scans
  without constructing paragraph values, and the source-backed facade gains
  the same one-paragraph selector.
- PPTX opened transactions reuse the already validated selected `Scene` when
  mapping a shape back to its raw XML span. This removes one complete scene
  parse per `set_shape_text` operation without changing source fingerprints,
  MCE selection, raw-span validation, staged readback, patch construction, or
  candidate recapture.
- DOCX plaintext package output now accepts a forward-only `Write` sink. The
  private implementation already forwarded to OPC's sequential writer and did
  not seek. Atomic path save remains separate and unchanged.

The selector and shape-span changes remove temporary semantic work. They do
not introduce a persistent cache or change immutable snapshot identity.

## Benchmark coverage

The standalone harness adds 16 opt-in public-API cases over deterministic
tiny, medium and large DOCX/PPTX corpora. The cases cover owned open, semantic
listing, one object, full text, small creation, exact no-op edit/save, one
edit/save and approximately 1% edit/save. Every mutation is reopened and its
complete semantic state is checked. The default matrix remains 36 cases and
198 records; the harness now has 60 selectable cases in total.

The decisive generated inputs contain 10,000 DOCX paragraphs and 100 PPTX
slides with 100 text boxes each. The 1% PPTX case changes 100 evenly spaced
text boxes in one transaction. These are synthetic scale corpora, not claims
about producer-specific compatibility or media/dependency preservation.

## Matched latency result

The release binaries were frozen independently of the dirty worktree:

- before: `perf-baseline-before-d98458d62`, benchmark commit `d98458d62` over
  production base `aa1adf3fb`, SHA-256
  `1f065106cc5d91f1e0c6eaae350d00992bce38d13f94105d2f24fdc5b270dc74`
- after: `perf-baseline-after-dense`, SHA-256
  `d232d55c250e4aac1980a93a81adaa4c8e7059b4a79845b54d0c814215b70f42`

Both states ran on CPU 2 in fixed before-A, after-A, after-B, before-B order.
The broad capture used three warmups and 15 measured samples per leg. The
single-edit guardrail was repeated separately with five warmups and 30 samples
per leg because the first broad pass showed order-sensitive drift. Tables pool
both matching legs. Times are milliseconds; mean intervals are two-sided
Student's-t 95% intervals over the pooled samples.

| Case | Before p50 / p95 / p99 | After p50 / p95 / p99 | p50 delta | Before mean (95% CI) | After mean (95% CI) | Mean delta |
|---|---:|---:|---:|---:|---:|---:|
| DOCX one paragraph, 10,000 paragraphs | 2.945 / 3.147 / 3.207 | 2.805 / 2.976 / 3.102 | **-4.72%** | 2.981 (2.951-3.011) | 2.832 (2.802-2.863) | **-4.99%** |
| PPTX 1% edit/save, 10,000 text boxes | 399.320 / 419.774 / 421.042 | 361.915 / 379.003 / 386.557 | **-9.37%** | 401.730 (398.676-404.785) | 364.076 (361.201-366.950) | **-9.37%** |
| PPTX one edit/save guardrail | 126.607 / 132.514 / 141.408 | 126.956 / 133.179 / 137.591 | +0.28% | 127.097 (126.196-127.999) | 127.410 (126.656-128.164) | +0.25% |

The one-edit result is treated as latency-neutral: its mean intervals overlap
and the aggregate transaction is dominated by complete presentation capture,
fingerprinting, candidate recapture and publication. The repeated-edit result
is the acceptance cell for scene reuse.

Raw broad samples:
[`before A`](../results/abba-semantic-before-a.json),
[`after A`](../results/abba-semantic-after-a.json),
[`after B`](../results/abba-semantic-after-b.json), and
[`before B`](../results/abba-semantic-before-b.json). Dedicated one-edit
guardrail:
[`before A`](../results/abba-semantic-one-before-a.json),
[`after A`](../results/abba-semantic-one-after-a.json),
[`after B`](../results/abba-semantic-one-after-b.json), and
[`before B`](../results/abba-semantic-one-before-b.json).

## Allocation and memory result

Heaptrack used the identical fixed-CPU commands for each binary. For 33 DOCX
selector invocations (three warmups and 30 samples), total process allocation
calls fell from 5,894,798 to 5,894,468: the old selector's 330 collection-growth
allocations, ten per invocation, disappeared. Peak heap remained 35.55 MB
because corpus construction and complete post-operation verification dominate
the process peak. The full bounded XML scan is deliberately retained, so the
selector remains linear in document XML size.

For the PPTX 1% case (one warmup and three samples), full-process allocation
calls fell from 33,944,220 to 29,983,027 (**-11.67%**). Temporary allocation
count was effectively flat (-0.03%), peak heap stayed 17.88 MB, and profiler
RSS moved from 29.63 to 30.01 MB (+1.28%). This is accepted as work and
allocation removal, not as a peak-memory improvement.

## Correctness and contract gates

- DOCX selector tests cover transitional/strict namespaces, nested table
  paragraphs, empty paragraphs, out-of-range selection, shared text semantics,
  and malformed trailing XML. The complete DOCX library suite passes: 831
  tests; the three-test positional source-backed integration suite also passes.
- A write-only sink with no `Seek` implementation serializes and reopens a
  complete DOCX package. Existing failed/panicking sink rollback tests retain
  retryable semantic and property state.
- The complete PPTX library suite passes: 463 tests, including opened patch,
  inverse/history, complex chart, notes, picture/dependency closure, MCE,
  unknown markup and full reopen cases.
- Warning-denied all-feature library Clippy passes for DOCX and PPTX. The
  all-target variant remains blocked by existing unrelated DOCX test-only
  lints and is not reported as passing.
- The performance tool passes 19 tests, warning-denied release Clippy,
  formatting, and a 16-record semantic release smoke. Corpus hashes, operation
  counts, complete semantic readback, and DOCX sink summaries are recorded.
- CI now checks the actual 36-case smoke/198-record default contract, runs all
  16 semantic cases on the tiny corpus for pushes and pull requests, and
  publishes scheduled tiny/large semantic release JSON alongside the default
  matrix.

No archive type, physical identifier, executor, unsafe code, network access,
ambient filesystem input, or new production dependency is introduced.

## Remaining limitations

This tranche does not complete the CRUD matrix. DOCX 1% replacement still
rebuilds and reparses the complete document once per changed paragraph; a
coalesced same-structure replacement plan remains the higher-return edit
candidate. PPTX capture/apply still fingerprints and recaptures the complete
presentation graph. Forward-only DOCX final serialization is not a
memory-bounded authoring stream or an existing-document tail append.

Cold filesystem, real remote range, conversion/export, durable patch timing,
dependency-copy timing, malformed/security/protection, unknown-extension
semantic edits, ODF and iWork matrices remain separate required work.
