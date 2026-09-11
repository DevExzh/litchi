# 0506: discover ODG shape attribute spans in one checked pass

Shape parsing previously ran the same checked XML attribute iterator once for
each of 16 editable source spans. The fresh baseline attributes 211,618,426
inclusive Callgrind instruction references (18.51% of the metadata-large child)
to `attribute_source_span`; raw value-span scanning contributes only 0.94%.
This batch replaces the shape-only discovery calls with one checked pass and
16 fixed stack slots. Value parsing, page discovery and all other callers keep
their existing paths. No persistent cache or public API is added.

The slots retain the selected raw qualified names. The same lexical scanner
produces ranges in the original field order. Any raw iterator, duplicate-name
or value-span error invokes the original ordered helper calls, retaining exact
error precedence. Successful batching does not skip malformed trailing
attributes, namespace matching, or duplicate expanded-name checks. The wrapper
publishes no partial result. Existing document, shape, depth and source limits
remain unchanged; transient storage is fixed independently of attribute count.

## Matched measurement

[Evidence](../results/change-0506/README.md) retains the unchanged locked 0502
release probe, clean baseline `56e396d05`, the candidate patch, source/binary
identities, tool receipts and raw reports. The frozen baseline executable is
byte-identical to the accepted 0505 candidate. A separate 400-sample pilot
preceded candidate acceptance. Formal measurements run serial A1/B1/B2/A2 blocks
on CPU 2, with 25 warmups and 200 samples per corpus per child: 16 children and
3,200 measured samples. Fixture generation and hashing are outside the timer;
owned-byte open, traversal, checksum assertion and drop are inside it.

| Corpus | Before p50 ms, r1 / r2 | After p50 ms, r1 / r2 | p50 change, r1 / r2 | RSS change, r1 / r2 |
| --- | --- | --- | --- | --- |
| plain-small | 1.933 / 1.940 | 1.409 / 1.420 | -27.11% / -26.81% | -0.76% / +0.16% |
| plain-large | 59.906 / 60.206 | 43.035 / 43.172 | -28.16% / -28.29% | **+10.15% / +9.86%** |
| metadata-small | 1.598 / 1.596 | 1.325 / 1.324 | -17.03% / -17.00% | +4.10% / +3.76% |
| metadata-large | 22.509 / 22.398 | 18.503 / 18.353 | -17.80% / -18.06% | -1.69% / +1.98% |

The summary retains every p95/p99 and throughput pair, process RSS receipt and
10,000-resample median-ratio interval. The fixed 0502 bootstrap seed measures
within-run uncertainty, not machine-to-machine variation. There are two adverse
>5% flags, both plain-large RSS; no paired timing or throughput regression
exceeds 5%. The original traversal checksum and input hashes match, but that
checksum does not cover every semantic field. Focused differential and edit
preservation tests supply additional correctness evidence.

## Memory review and acceptance

The uninstrumented plain-large process grows from 20,460/20,480 KiB peak RSS
to 22,536/22,500 KiB: 2,076/2,020 KiB extra. This is an accepted, explicitly
retained tradeoff for the 28% plain-large median reduction and substantial
allocation-work elimination. It is not a claim of improved resident memory.
No resource limit is relaxed and the helper adds no persistent state. Further
allocator/working-set attribution remains open; the cause of this RSS change
has not been established.

Separate whole-child heaptrack runs reduce allocation calls by 26.81% on
metadata-large (507,836 to 371,676) and 36.48% on plain-large (1,571,975 to
998,535). Rounded peak live heap remains 6.98M and 12.57M respectively. Equal
rounded displays do not prove exact byte equality. Plain-large profiler RSS
moves 21.04M to 20.98M, which does not negate the uninstrumented regression:
heaptrack changes process behavior and these are distinct measurements.
Whole-child counts include setup and two opens (preflight plus one sample),
not isolated operation allocation totals. No allocator-specific workaround is
introduced to improve one RSS measurement.

## Attribution and verification

Metadata-large whole-child instruction references fall 1,143,461,037 to
961,714,297 (15.89%). `parse_content` inclusive references fall 695,800,871 to
514,007,645. The new shape-span wrapper accounts for 30,609,023 references;
other old-helper calls remain. Repeated value lookups still account for about
355M inclusive references and are a remaining measured opportunity. Hardware
counters remain unavailable (`perf_event_paranoid=4`); Callgrind does not
provide hardware cycles, IPC or cache results here. Untimed setup and hashing
are included in its whole-child scope.

Four integration tests exercise all 16 source-span mappings across rectangle,
line, polygon, path and control shapes. They assert exact XML after multi-field
edits, aliases and same-local foreign fields, quote/ordering/unknown-byte
preservation, durable patch application and inverse restoration, plus malformed
and duplicate trailing attribute refusal with valid controls. Private tests
compare the batched wrapper against the original ordered helper, including
error precedence and unusual raw spans. The all-target suite passes 108 unique tests (including eight private
differential tests). Commands are recorded in `gates.json`; the four-test
focused rerun is not counted again. An initial test compile failed on a
redundant import; its log is retained, and removing that test-only import
resolved the failure without changing the timed production candidate.

[The ADR review](../results/change-0506/adr-review.md) and its mechanically
hashed manifest cover the applicable architecture contracts. Scoped ODG,
ODF-umbrella and boundary checks are recorded with their exact logs. This batch
does not claim a full-workspace build, live Office run, sanitizer or fuzz
campaign, or real scaling/remote-provider coverage. Existing fixture tests are
included in the owner suite. The strict claim registry stays at ten entries.
These owned-byte synthetic open/traversal results do not close the full CRUD
program or establish equivalence with the older, less expressive 0502 parser.
Batch-owned scratch is removed after evidence retention and final checks.
