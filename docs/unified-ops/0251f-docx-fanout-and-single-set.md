# 0251f — DOCX 10,000-paragraph fanout and the single-set overhead

This record covers plan step 7a items 7 and 8 of
[0248](0248-coordinator-restart-plan.md) §6 Phase B. Both are follow-ups of
[0089](0089-docx-paragraph-fanout-baseline.md).

## Item 7: the 10k-batch refusal

In 0089, a 10,000-paragraph text fanout was a typed refusal. Two bounds
caused it:

- the native composite replacement stopped at 4,096
  (`CompositeParagraphReplacements`);
- the edit's operation ledger stopped at 4,096 (`Operations`).

The second bound alone would have refused a chunked write.

**Change.** The leaf's per-edit operation bound `MAX_OPERATIONS` rises from
4,096 to 65,536. The new value equals the unified layer's expanded-target
bound (`SemanticLimits::max_expanded_targets`).

- **Bounds that follow it.** The composite paragraph and hyperlink
  replacement limits, the managed paragraph operation counts and durable
  patch decoding all derive from `MAX_OPERATIONS`.
- **Bounds that do not follow it.** Two unrelated per-element scans had
  reused the constant, and each now has its own 4,096 bound:
  - field-result run scans (`MAX_FIELD_RUNS`), whose cost is quadratic in
    the runs of one paragraph;
  - nested element path steps (`MAX_PATH_STEPS`).

  Their adversarial worst cases are therefore unchanged.
- **Other limits.** Replacement text stays bounded at 16 MiB per edit, and
  managed owners still charge every operation to their execution budget.

`set /body > p text=…` over 10,000 paragraphs is now one write on both
owners. The leaf `document_limits` test moves to the new bound.

The frozen 0089 harness (`tools/perf-baseline/unified-docx-text-fanout`)
still asserts the old refusal above 4,096 targets. It reproduces 0089 at
that record's commit, not at this one.

## Item 8: the single-set overhead miss

0089 measured `set /body/p[1] text=…` against the leaf
`replace_paragraph_text` at 1.43 on eager owners. Profiling the unified leg
(perf, frame pointers, CPU 5) showed the cost breakdown. Of the unified-only
work:

- path resolution read the body index twice: once to resolve `/body/p[1]`,
  and again to prove the write position;
- each read built a `Document`, which resolves the main part by
  normalizing the `officeDocument` relationship target, about 2% per call.

**Change.** The write-position proof of a direct body paragraph reuses the
body index from which the paragraph was just resolved. It comes from the
same owner generation, so it is the same index. The ratio fell from about
1.26 (this branch before the change) to 1.18.

**Still a miss.** 1.18 is above the 10% budget.

- **Where the rest goes.** The remaining unified-only cost is resolving the
  path against the body index, about 0.9 µs of 10.1 µs. The leaf leg gets
  its `Position` from the caller and pays nothing for it.
- **Root cause.** It is mostly `Package::document()` →
  `OpcPackage::main_document_part()` → `PackURI::from_rel_ref`. The leaf
  pays the same cost four more times per commit.
- **Candidate fixes**, neither taken here:
  - memoize the main part name in `litchi-opc`. That is outside this
    branch's crates, and it would speed both legs;
  - keep a body index per immutable `Docx` handle, which helps repeated
    writes from one handle but not a first write.

## Measurements

The harness is `crates/litchi-ops/examples/docx_fanout_bench.rs`, run in
release on AMD EPYC 9R45, pinned to CPU 5, with ABBA rounds (9 rounds,
2 warmup).

- **Fixture.** Eager owners of N one-run paragraphs.
- **Input.** The text contains a tab and a line feed, as in 0089.
- **Legs.** N = 1 times the scalar paths (`replace_paragraph_text` against
  `set /body/p[1]`). N > 1 times the composite `replace_body_paragraph_texts`
  against `set /body > p`.

| Paragraphs | Leaf µs | Unified µs | Ratio, trial 1 | Ratio, trial 2 |
|---:|---:|---:|---:|---:|
| 1 | 8.6 | 10.1 | 1.177 | 1.183 |
| 100 | 347 | 369 | 1.064 | 1.051 |
| 1,000 | 3,348 | 3,537 | 1.056 | 1.045 |
| 10,000 | 33,552 | 35,443 | 1.056 | 1.046 |

The microsecond columns are from trial 1.

- Scaling is linear: 1,000 paragraphs take 3.35 ms and 10,000 take 33.6 ms
  on the leaf.
- Fanouts are inside the 10% budget at every size, including the 10,000
  case that 0089 could not time.
- 0089's ratios at 100 and 1,000 paragraphs were 1.29 and 1.28 on its
  corpus.
- The load average was about 6 on 32 CPUs.

Evidence: `results/0251f-docx-fanout/trial-{1,2}.json`. Regenerate with:

```
taskset -c 5 target/release/examples/docx_fanout_bench
```

Use `DOCX_FANOUT_SIZES` to choose sizes, and `DOCX_FANOUT_PROFILE=unified`
to profile one leg.

## Verification

| Command | Exit |
|---|---|
| `cargo test -p litchi-docx --all-features` | 0 (`document_limits` at 65,536) |
| `cargo test -p litchi-ops --all-features` | 0 (`docx_unified_fanout`: 10,000 paragraphs on both owners, reopened; `docx_batches` capacity tests moved to 65,536; the body-load count of paragraph admission is now 0) |
| `cargo test -p litchi --all-features` | 0 (see below) |

The facade's `markdown_docx_semantics` golden had recorded a picture title
as `Picture &quot;one&quot;`. That was the escaped `docPr/@name`, which
0251d's `InlineImage` reader fix now decodes. The golden now expects
`Picture \"one\"`, as the hyperlink tooltip beside it already did.

## Not claimed

- The single-set miss (1.18) is reduced, not closed.
- Source-backed owners were not timed here; their correctness is tested at
  10,000 paragraphs.
- No allocation or RSS measurement.
- The 0089 command-batch requirement (per-item results and errors) is not
  addressed by a one-query fanout.
