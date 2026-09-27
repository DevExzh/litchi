# 0775 — MCE stream namespace sharing

Status: integration validated with scoped paired evidence and explicit
regression flags below.

The clean 0771 candidate applies to `92dfbe0bd5` as three commits ending at
`88cb5bf0fa`. The separate dirty 0770 worktree is outside this batch.
[Evidence packet](results/change-0775/README.md).

The stream shares immutable namespace text through `NamespaceUri` and uses
scope identities for duplicate-name and compatibility-directive checks.
XLSX stream consumers retain shared names. Local names remain owned strings;
namespace declarations still hash and compare URI text when entered into scope.
This eliminates repeated URI copies in expanded event names and repeated URI
hashing in those internal checks. It does not make arbitrary public hashing,
observer work, or matching registered-extension checks independent of URI size.

## Breaking changes

Raw and semantic MCE stream events expose `ExpandedName` in place of `Name`.
Its `namespace` field is `NamespaceUri`, supporting `as_str`, `AsRef<str>`,
`Borrow<str>`, dereference and text equality/hash semantics. Use `.into()` when
constructing an owned namespace or converting it to `String`. `Name` remains
the owned policy type accepted by `Capabilities`. Downstream code explicitly
typing stream names or constructing fields from strings must migrate.

## Architecture and limits

| Constraint | Disposition |
| --- | --- |
| ADR 0001/0006 | Existing name/resource checks and typed errors remain; alias changes require explicit semantic triage. |
| ADR 0002/0024 | Production changes stay in common OOXML and XLSX owners, with no new runtime dependency. |
| ADR 0005 | Shared immutable strings replace repeated event-name copies. Scope identities expire with bindings; no global cache or hidden parallel execution is added. |
| ADR 0032 | This is scope-owned text sharing, not a new snapshot memo or cross-operation cache. |

URI string reservation remains fallible; the constant-sized `Arc<String>`
allocation uses ordinary `Arc::new`. This is not a claim of complete OOM
recovery. Retained event names can keep their URI allocation alive after scope
exit. Public text hashing and equality between independent allocations may
still read the entire URI.

Independent review found that matching an element against registered extension
names still constructs a full owned `Name` in both processors. That remaining
cost is measured separately and excluded from the optimization claim. Many
aliases also retain work proportional to the text in their declarations.

## Verification and evidence

Nine serial gates passed on `88cb5bf0fa`: owner format and all-target/all-feature
check, common OOXML/XLSX and DOCX/PPTX tests, warning-denied library Clippy and
rustdoc, the OOXML-enabled facade check, standalone harness all-target check,
and crate-boundary validation. Test logs contain 5,669 passing tests in 237
suites and 35 ignored tests. This is not a full facade/harness test run.
The only subsequent Rust changes at `fbb579914c` narrow comments; their exact
patch is archived and separate format/rustdoc checks pass.

Two initial differential attempts are preserved. The first malformed controls
missed the tree processor's MCE path; the second exposed a pre-existing gap:
the tree rejects lexical duplicate attributes but accepts the tested duplicate
expanded name through aliased prefixes. The stream rejects both. The final
probe records that gap explicitly and gates the stream refusals and lexical
tree refusal. It does not declare the malformed input valid. No production
change was made in response, and no timing observation came from either
failed attempt. See `probe-corrections.md`.

The final differential covers 168 admitted ZIP packages and 3,870 XML members,
4,096 ordinary generated documents, and 4,096 alias-focused generated documents,
under the recorded capability sets. It produces 64,541 comparisons. Real-fixture
and ordinary-generated outcomes agree across both legs. There are 1,037 changed
outcomes: 1,017 alias-generated and 20 reserved-XML compatibility pair results.
The bounded pair matrix repeats six distinct templates across 64 entries and
three capability sets; those repetitions are not 64 distinct test designs.
Normal/MCE aliases, default namespace reset and nested rebinding pairs pass.
Reserved XML/XMLNS cases are parser-compatibility observations, not standards
conformance proof. The pair oracle compares each processor across spellings;
it does not require a tree transcript to equal a stream transcript.

A separate serial run at `bdea6e8b86`, immediately before namespace identities,
uses the identical final probe and lock. All 64,541 candidate results exactly
match that text-based reference, including every changed baseline result.
All 1,017 generated changes are tree-processor keys; stream transcripts remain
unchanged even in the alias corpus. The 20 pair changes are repetitions of the
reserved XML alias case under bare/understood capabilities. This restores
pre-identity text semantics rather than introducing unexplained new behavior.
`alias-triage.json` and the complete reference report preserve the comparison.
The prefix-normalized codec oracle reparses codec output with the stream under
empty capabilities; it is not independent raw serialization/report comparison
and does not cover opaque extension output.

## Current paired observations

The AMD EPYC 9R45 Linux host exposes 32 CPUs. Rust 1.95.0 release builds use
LTO, panic abort, identical standalone source/dependency locks and CPU 12.
Three processes per case and leg run in before/after/after/before/before/after
order. Long-URI cases use three samples after one warmup; others nine after two.
Input generation is outside timing. Parsing and observers are inside; the
extension policy is initialized during warmup. Process RSS includes inputs,
policies, output and observer state. Fixture/binary/source hashes remain fixed.
All 108 timed process runs have equal input identities and outcomes across
legs. The two real-part digest observers and their count-only controls are
reported separately because observer work is not parser work.


The non-iWork program remains active. This batch cannot establish full CRUD
coverage, cold-cache/range-source behavior, parallel scaling, native Office
interoperability, or completion of the full goal audit.

| Case | Median process p50, before → after (ms) | Change | Median peak RSS, before → after (KiB) |
| --- | ---: | ---: | ---: |
| `mce_stream_long_uri` | 2253.992192 → 1.738848 | -99.92% | 2,057,076 → 11,256 |
| `mce_stream_review_long_uri` | 3416.031787 → 1.879548 | -99.94% | 2,040,340 → 10,896 |
| `mce_stream_long_uri_elements` | 209.195042 → 2.234689 | -98.93% | 11,140 → 11,188 |
| `mce_stream_long_uri_skipped` | 2140.642877 → 2.151960 | -99.90% | 13,040 → 11,108 |
| `mce_stream_long_uri_tokens` | 2669.767862 → 5.338622 | -99.80% | 12,208 → 12,332 |
| `mce_stream_short_uri` | 0.999564 → 0.878553 | -12.11% | 3,572 → 3,344 |
| `mce_stream_benign_worksheet` | 73.657814 → 71.633066 | -2.75% | 4,540 → 4,552 |
| `mce_stream_benign_document` | 9.552011 → 9.262670 | -3.03% | 3,556 → 3,580 |
| `mce_stream_count_worksheet` | 40.259022 → 40.530403 | +0.67% | 4,416 → 4,552 |
| `mce_stream_count_document` | 5.328743 → 5.128912 | -3.75% | 3,288 → 3,236 |
| `mce_extension_long_uri` | 13.802719 → 13.817679 | +0.11% | 10,884 → 10,948 |
| `mce_stream_extension_long_uri` | 14.700023 → 12.568634 | -14.50% | 12,028 → 12,020 |
| `mce_extension_short_uri` | 0.009950 → 0.010650 | +7.04% | 2,800 → 2,624 |
| `mce_stream_extension_short_uri` | 0.026070 → 0.027900 | +7.02% | 2,796 → 2,740 |
| `mce_aliases_low` | 0.029410 → 0.029740 | +1.12% | 2,516 → 2,780 |
| `mce_aliases_high` | 0.618163 → 0.624062 | +0.95% | 3,224 → 3,228 |
| `mce_stream_aliases_low` | 0.063300 → 0.060390 | -4.60% | 2,664 → 2,772 |
| `mce_stream_aliases_high` | 1.136745 → 1.111195 | -2.25% | 3,812 → 3,824 |

The long-URI attribute case uses four elements with 1,000 attributes each under
a roughly 1 MiB URI. Its median peak process RSS drops from about 1.96 GiB to
11 MiB as event names share text. The long-element/skipped/directive cases
exercise 4,000 elements. These are named adversarial in-memory parser scenarios,
not end-to-end Office editing speedups. Matching-extension controls use only
64 elements; their remaining URI-size dependence is visible. Alias controls
use 64/1,000 prefixes for one short URI and do not eliminate declaration work.

Two short-extension microcases cross the 5% latency flag: tree +7.04% (0.700 µs)
and stream +7.02% (1.830 µs). The low-alias tree case crosses the RSS flag,
+10.49% (+264 KiB). These are accepted as explicit bounded tradeoffs alongside
the correction to namespace identity and large hostile-input improvement;
they are not dismissed as proven noise. The real-document cases stay below
5% latency/RSS regression. Process-p50 spread reaches 11.68% on the now-short
long-attribute candidate and 5.28% on the review candidate; all raw ranges and
flags are archived. Three processes and three/nine samples do not establish
stable p95/p99 tails. No allocation-stack or hardware-counter attribution is
claimed by this packet; work-counter unit tests and RSS support the mechanism.

Historical 0771 performance numbers are not adopted. Remaining extension-name
copies, the pre-existing tree expanded-duplicate gap, and broader alias/directive
matrices remain follow-up work. The candidate must not be described as removing
all URI-dependent work or proving general XML namespace conformance.

## Completion

The final packet retains both failed probe attempts, exact source/lock and
binary identities, all raw process measurements, differential reports and
reference capture. Offline validation checks these receipts, recomputes the
statistics and compares the entire reference result map. Owned temporary
build targets are removed only after binary identity verification. Original
0770/0771/reference worktrees and unrelated main files remain untouched.

After integration, the owned worktree, copied root lock and reference symlinks
were removed. Final production and sealed packet bytes were verified in main;
the three unrelated main files remain byte-for-byte unchanged.
