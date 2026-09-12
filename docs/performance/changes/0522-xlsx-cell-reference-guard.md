# 0522: reject the common-cell scanner candidate; retain XLSX guards

The scanner candidate failed the frozen requirement for useful, repeatable
primary total improvement and was reverted. Dense-sparse improved in both
repeats, but medium regressed 1.29% in the first and improved 1.30% in the
second. Lower commit instruction and allocation counts do not override that
native admission rule. This batch retains three baseline-compatible codec
tests, an opt-in noncompact benchmark shape, and the complete rejected-candidate
evidence. It makes no adopted production speedup claim.

## Candidate and retained guards

The proposed private scanner helper borrowed the decoded cell reference and
used its checked attribute pass to prove that an unprefixed `c` with no
attributes or only `r` had the existing compact `tag = None` representation.
Only that case skipped a second `cell_tag` scan. All other tags kept the old
path, and attribute error order, coordinate checks, semantic validation,
readback, publication and resource controls remained. The
[source review](../results/change-0522/final-source-review.md) found no static
blocker; it reviews the measured candidate, not an adopted change. The final
production scanner is byte-identical to the baseline.

The retained codec tests cover ordinary and prefixed cells, inferred and
escaped references, extra attributes, tag/attribute order, exact typed/display
error precedence, invalid UTF-8 and genuine truncation. Initial pre-capture
failures exposed a truncation-fixture mistake and a missing harness counter
type; both were corrected before freezing, and the original records remain.

The opt-in `noncompact` shape preserves medium geometry (four 48×48 sheets)
and scalar values while alternating unprefixed cells with `r` plus `t="n"`
and prefixed `x:c` cells with `r`. Every sheet has both forms. Its test checks
deterministic bytes, namespace/tag counts, public readback, source-backed
edit/save and untouched members. It is excluded from the default shape list
and explicitly selected for normal and managed guards. It is a generated
numeric fixture, not a native Office-producer compatibility corpus.

## Frozen native comparison

The native block order was baseline–candidate–candidate–retained baseline.
The final baseline block used the unchanged baseline executable under the
candidate checkout; receipts separately bind compiled and working source
manifests and check binary hashes before and after each child. Only the
production scanner differs between stage manifests. Both stages use identical
tests and harness inputs. CPU affinity is 2; builds use Rust 1.95.0, two jobs,
no incremental compilation and ordinary standalone-harness release defaults.

| Primary shape / repeat | Baseline p50 ms | Candidate p50 ms | Candidate change |
| --- | ---: | ---: | ---: |
| Medium / 1 | 25.476 | 25.804 | +1.29% |
| Medium / 2 | 25.609 | 25.275 | −1.30% |
| Dense-sparse / 1 | 50.311 | 49.276 | −2.06% |
| Dense-sparse / 2 | 49.774 | 49.119 | −1.32% |

Medium repeat 1 also regresses p95, p99 and mean by 1.01%, 1.34% and 1.25%.
Primary commit p50 improves 4.36–5.71%, but total improvement is not repeatable
across the primary matrix. Noncompact total p50 ranges from −0.49% to +1.14%;
all guard results remain in the
[comparison](../results/change-0522/comparison.json).

The matrix retains 1,640 native samples: two fresh children per primary
shape with 20 warmups and 100 samples, plus two children per guard case/shape
with 10 warmups and 30 samples. Total time sums open, planning, staged
sets/commit and publication, including the returned snapshot's publication
drop. Setup, other destruction, reopen and oracles are excluded. The source
is instrumented in-memory `ReadAt`, with a fresh editor/cache each iteration.
Printed filesystem/range defaults do not activate those providers.

All 43 adverse phase metrics above 5% remain in the
[flag review](../results/change-0522/flag-review.json): 20 open, 17 publication
and six excluded-reopen metrics. All 67 same-build repeat variations remain.
Their causes are unproven; neither variation nor the absence of a >5% total
flag overrides the admission rule. Within-child bootstrap intervals are
descriptive, and two children do not establish cross-host confidence. Serial
campaign scheduling does not assert exclusive use of the shared host.

## Rejected-candidate mechanism measurements

Separate canonical allocator captures retain 40 samples, ten per shape per
stage, covering staged sets plus commit. Every sample per shape/stage agrees.

| Shape | Allocation calls before → candidate | Allocated bytes before → candidate | Incremental region peak |
| --- | ---: | ---: | ---: |
| Medium | 118,744 → 100,312 | 20,344,427 → 19,724,459 | 2,984,983 unchanged |
| Dense-sparse | 225,771 → 190,187 | 27,331,029 → 26,122,091 | 7,335,225 unchanged |

Allocation calls fall 15.52–15.76% and allocated bytes 3.05–4.42% in the
rejected candidate. Reallocations and exact live-byte balances remain
consistent. Incremental peak is region peak minus entry live bytes; it is
distinct from whole-child RSS. Allocator-build timings are excluded from
native comparisons, and no RSS reduction is claimed.

Eight separate Callgrind children isolate one timed commit each. Raw summary,
one positive incoming runner edge, and self plus direct-callee costs reconcile.
Commit Ir falls 3.76–4.17%. The scanner-to-`cell_tag` positive edge is present
in every baseline primary profile and absent in every candidate primary
profile. The older validator-to-owned-event edge remains absent in both
stages and is not credited to this candidate. All lifecycle/final raw dumps,
annotations and warnings are retained. Inner Callgrind call metadata is not
used for timed event or allocation counts. See
[profile analysis](../results/change-0522/profile-analysis.json).

## Correctness, custody and remaining work

Corpus, output, semantic, source I/O/cache and managed budget identities match
across stages and instrumentation lanes after planned iteration normalization.
Oracles cover deterministic output, untouched members, no-op, clear/remove,
foreign/stale refusal and semantic inverse restoration. The vendor-extension
guard also checks partial-sink failure. Primary one-percent edits touch all
four sheets; zero unselected reads does not establish selective-subset access.
Unknown package members are distinct from unknown worksheet grammar.

The quality summary retains twelve candidate gates, including 1,277
all-feature XLSX tests and five harness tests (1,282 total), formatting,
all-feature workspace check, XLSX/harness clippy, XLSX rustdoc, crate boundaries
and strict existing-claim checks. Candidate checks bind the measured candidate; the final 19-test codec
check binds the restored baseline scanner and retained tests. The complete
bundle verifies source patches through private Git indexes, all serial receipt
intervals, exact report and annotation replay, four negative evidence vectors,
final source disposition, cleanup and a recursive hash inventory. See
[reproduction and scope](../results/change-0522/README.md).

No API, dependency, unsafe-code, budget policy or preservation-contract change
is retained. The 30 previously read ADR/index hashes remain unchanged.
No fuzz, native Office-producer, physical FileSource/range, cold-cache,
hardware-counter, scaling or coverage-catalog completion is claimed. Required
semantic parsing/readback remains; 0514/0516 fusion remains rejected. Further
work needs a larger measured end-to-end opportunity and an independent
work-elimination proof, not a relaxed gate or repeated attempts to erase
these mixed results. The [next-priority review](../results/change-0522/next-priority-review.md)
selects fresh CFB/OLE2 source-open attribution, followed by a bounded
work-elimination candidate only if the profile and allocation evidence justify
it. OOXML semantic-reference ownership and reconstruction remain alternatives.

OLE2/OOXML remains the active optimization priority. ODF, including the ODG
regression investigation, is deferred until that goal completes; iWork is
excluded. The broader goal remains open.
