# 0776 — refuse duplicate expanded attributes in tree MCE processing

Status: validated and accepted at `6ed76a881b`, with sealed evidence and
owned build targets removed. The broader non-iWork goal remains active.

[0775](0775-mce-stream-integration.md) proved an existing mismatch: the tree
processor accepted `a:k` and `b:k` on one element when both prefixes named the
same URI, while the stream processor rejected them. The quick-xml check only
compares lexical names. This change closes that tree-side gap before MCE
selection, attribute removal, or the opaque-child early return.

## Implementation

After local namespace declarations are installed and the element scope is
available, the common MCE codec filters prefixed, non-`xmlns` attributes. It
compares borrowed local-name text first and resolves namespace identities only
for possible local-name collisions. Unprefixed attributes have no namespace and
are already covered by the lexical duplicate check; declarations cannot bind a
prefix to the empty URI.

The first eight candidates use fixed stack storage. For a larger list, the
codec fallibly reserves at most the source attribute count, copies the eight
borrowed tuples into that vector, sorts by local name, resolves only colliding
groups, and sorts each group by namespace identity. A tuple is 40 bytes on the
captured target. Namespace URI text is neither copied nor hashed per name, and
the scratch is dropped before children are processed. The existing quick-xml
lexical check is unchanged; this is a bounded duplicate check, not an overall
parser complexity claim.

The DOCX tail-append requested-storage envelope charges the possible vector
capacity, eight fixed tuple slots, and vector owner conservatively even when
the common path does not spill. No new dependency, public API type, hidden
cache, or parallel work is introduced.

## Breaking changes and scope

MCE-marked inputs with duplicate expanded attributes now return
`Error::NonConformant("duplicate attribute")`, including duplicates in ignored,
skipped, unselected, or opaque branches that were previously accepted. A DOCX
operation near its memory ceiling may refuse earlier because its scratch
allowance is larger. The no-MCE borrowed passthrough and empty-offset shortcut
retain their existing scope.

This check does not promote the preprocessing API to a general XML validator.
The helper leaves existing single-attribute qualified-name handling in place;
full malformed-name parity for every skipped or opaque path is a separate
audit. The seven new tests cover duplicate detection, branch behavior, and
overflow ordering without claiming exhaustive XML validation.

## Verification

ADR 0006 requires malformed known payloads to fail closed. ADR 0005 requires
bounded fallible scratch storage and explicit accounting by its caller. ADR
0002 and ADR 0024 keep that requested-storage envelope with the existing DOCX
tail-append owner. This change stays within those owners and adds no ordinary
public API type.

The final `quality-2` run is bound to source `6ed76a881b` and passed all nine
serial gates: formatting, affected-owner all-target/all-feature checks and
tests, warning-denied Clippy and rustdoc, the OOXML-enabled facade check, the
harness all-target check, and crate-boundary checks. It records 5,676 passing
tests, 35 ignored tests, and 237 suites. The earlier `quality-1` run passed the
same gates for the intermediate `2cdd75caad` candidate. The archived
`quality-0` attempt stopped at its second gate because a test used an
unnecessarily qualified `super::stream` path under `-D unused-qualifications`;
the correction is retained in the later source and logs.

The focused tests cover ordinary/XML/MCE aliases and late declarations,
opaque children, skipped and unselected branches, eight-name and nine-name
lists, overflow-only duplicates, legal distinct names, active-offset refusal,
and unchanged borrowed passthrough. Large negative cases put the duplicate
late enough to exercise sorting.

## Performance and differential evidence

The captures use fresh paired release builds at base `22f62ee328` and candidate
`6ed76a881b`, 18 cases, three processes per leg, and the alternating
before/after/after/before/before/after order on CPU 12. Long-URI cases use three
samples and one warmup; all other cases use nine samples and two warmups.
Each adversarial input is built before the `Instant` timing interval; the
interval covers the selected parser and its observer/digest work, while writing
the JSON report occurs afterward. RSS is a whole-process peak and therefore
includes process startup and all surrounding allocations.

The first complete capture (`measure-0`) used `2cdd75caad`, before the revised
common-path algorithm. Its rows with a positive change above five percent in
either latency or RSS were:

| case | median p50 change | median RSS change | observation |
| --- | ---: | ---: | --- |
| `mce_benign_worksheet` | +0.30% | +20.49% | ordinary fixture |
| `mce_benign_document` | +0.46% | +5.34% | ordinary fixture |
| `mce_aliases_low` | +5.59% | +7.88% | small alias case |
| `mce_aliases_high` | +1.33% | +6.08% | larger alias case |
| `docx_styles_benign` | +11.89% | −0.56% | known refusal in both legs |
| `mce_prefixed_1` | +2.04% | +6.73% | synthetic prefixed control |
| `mce_prefixed_2` | +24.41% | +6.46% | synthetic prefixed control |
| `mce_prefixed_8` | +30.06% | +2.66% | synthetic prefixed control |
| `mce_prefixed_9` | +32.25% | +8.54% | synthetic prefixed control |
| `mce_prefixed_32` | +37.57% | +3.29% | synthetic prefixed control |

`docx_styles_benign` is an inherited refusal control, not an accepted benign
DOCX case: both legs return
`err:invalid DOCX format: style numPr is missing numId`. The other initial rows
were below five percent in both columns. The synthetic 24–38% latency results
showed that the first design resolved too much namespace information on the
ordinary path, which led to the revised local-name collision filter captured in
`6ed76a881b`.

The final `measure-1` rows are reported below. A positive latency change above
five percent remains in the worksheet, low-alias, style-refusal, two-name,
eight-name, nine-name, and 32-name rows. There is no positive final RSS change
above five percent; the `mce_aliases_high` RSS change is a −5.93% decrease.

| case | p50 ms before → after | Δp50 | RSS KiB before → after | ΔRSS | outcome |
| --- | ---: | ---: | ---: | ---: | --- |
| `mce_benign_worksheet` | 12.917 → 13.662 | +5.77% | 7068 → 7088 | +0.28% | accepted fixture |
| `mce_benign_document` | 1.762 → 1.793 | +1.76% | 3628 → 3616 | −0.33% | accepted fixture |
| `mce_review_long_uri` | 18.200 → 18.735 | +2.94% | 31676 → 31576 | −0.32% | accepted control |
| `mce_long_uri_plain` | 6.111 → 6.095 | −0.28% | 27460 → 27332 | −0.47% | accepted control |
| `mce_long_uri_ignorable` | 16.872 → 17.559 | +4.07% | 27576 → 27292 | −1.03% | accepted control |
| `mce_long_uri_preserved` | 16.943 → 17.540 | +3.52% | 27564 → 27300 | −0.96% | accepted control |
| `mce_extension_short_uri` | 0.011 → 0.011 | +1.79% | 2752 → 2688 | −2.33% | accepted control |
| `mce_extension_long_uri` | 13.822 → 13.798 | −0.18% | 10916 → 10944 | +0.26% | accepted control |
| `mce_aliases_low` | 0.029 → 0.031 | +5.54% | 2708 → 2648 | −2.22% | accepted control |
| `mce_aliases_high` | 0.626 → 0.644 | +2.97% | 3308 → 3112 | −5.93% | accepted control |
| `mce_stream_count_worksheet` | 40.391 → 39.420 | −2.40% | 4600 → 4644 | +0.96% | accepted control |
| `mce_stream_count_document` | 5.253 → 5.161 | −1.76% | 3368 → 3392 | +0.71% | accepted control |
| `docx_styles_benign` | 1.629 → 1.721 | +5.62% | 3604 → 3572 | −0.89% | known refusal in both legs |
| `mce_prefixed_1` | 0.767 → 0.796 | +3.69% | 2996 → 3004 | +0.27% | accepted synthetic control |
| `mce_prefixed_2` | 1.165 → 1.236 | +6.11% | 3268 → 3380 | +3.43% | accepted synthetic control |
| `mce_prefixed_8` | 3.926 → 4.362 | +11.11% | 3792 → 3908 | +3.06% | accepted synthetic control |
| `mce_prefixed_9` | 4.588 → 5.216 | +13.70% | 3824 → 3696 | −3.35% | accepted synthetic control |
| `mce_prefixed_32` | 15.831 → 18.950 | +19.70% | 6420 → 6276 | −2.24% | accepted synthetic control |

The worksheet p50 ranges were 12.705836–12.952276 ms before and
13.607050–13.799011 ms after, so the reported +5.77% change is retained as an
observed result rather than dismissed as sampling noise. These three-process
captures do not establish stable tail quantiles, cold-cache behavior, range
source behavior, or CRUD coverage.

The final differential compares 64,541 candidate outcomes: 64,292 unchanged
and 249 changed. The malformed duplicate control changed from tree accepted /
stream rejected to both rejected. The other 248 changes are generated aliasing
inputs across 86 documents; 155 changed from an earlier successful tree result
to the duplicate-attribute refusal, and 93 changed the first error returned.
Independent deterministic triage found an expanded-name duplicate witness in
each changed generated document. There were no valid alias-pair failures and no
changes in the real fixtures, ordinary generated inputs, or stream outcomes.

The `allocations-0/` lane used one whole-process `heaptrack` sample for each
leg of the worksheet, document, and 32-name cases. Instrumented elapsed times
are excluded from the native timing table; these measurements include startup,
input generation, parsing, observation, and report serialization.

| case | allocation calls before → after | allocated bytes before → after |
| --- | ---: | ---: |
| `mce_benign_worksheet` | 129,512 → 129,512 | 12,455,824 → 12,455,822 |
| `mce_benign_document` | 16,056 → 16,056 | 2,539,495 → 2,539,493 |
| `mce_prefixed_32` | 36,093 → 40,093 | 24,239,691 → 29,359,689 |

The two-byte differences in the ordinary rows are consistent with the two-byte
path-string length difference between legs; they are not claimed as heap savings. The 32-name synthetic
case shows the bounded spill cost directly: 4,000 additional calls and
5,119,998 additional whole-process bytes. The evidence supports retaining the
correctness-first refusal while documenting the remaining costs; it does not
support a universal speedup claim.

## Differential scope and retained evidence

The packet retains source hashes for base and candidate, both quality attempts
and the corrected final quality run, paired release builds, raw process logs,
the deterministic probe, final analysis, independent triage witnesses, and the
allocation captures. The six-template alias-pair oracle is a compatibility
check rather than an exhaustive serialized-output oracle and does not cover
opaque extension output.

The original dirty 0770 worktree remains outside this batch. Follow-up work
must preserve this packet and the unrelated workspace changes while continuing
the broader cold, range, scaling, CRUD, and format-completeness audits.
