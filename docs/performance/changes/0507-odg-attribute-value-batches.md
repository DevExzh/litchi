# 0507: batch contiguous ODG attribute-value requests

The fresh 0506-baseline profile attributes 355,298,309 inclusive instruction
references to `attribute` (36.92% of the metadata-large child). Calls directly
from `Scanner::observe` contribute 241,328,793 references (25.08%). Repeated
checked attribute walks remain a substantial bottleneck after span batching.

This change combines two contiguous groups of eight string requests. Geometry
and line geometry stay after z-index conversion and before lexical validation.
The remaining eight shape properties stay after that validation. Name, frame,
3D, z-index, and viewBox/points validation lookups retain their positions.
Sixteen scalar walks become two checked walks on successful input.

The private const-size helper selects matching expanded names and uses the
same normalized-value decoder. It moves strings into fixed output slots. On
raw syntax, value-decode or semantic-duplicate detection, it drops partial
values and replays the original scalar helper in request order. This retains
which error is returned, including decoding an invalid second alias value
before reporting a semantic duplicate. The fast path does not format errors
that replay would discard. No persistent cache, public API, source ownership,
configured budget, dependency, thread, or unsafe policy changes.

## Matched comparison

[Evidence](../results/change-0507/README.md) retains clean base `4c6c63718`, the
exact candidate patch, source and executable hashes, tool output and reports.
The baseline executable is byte-identical to the accepted 0506 candidate;
the final rebuild is byte-identical to the timed after executable. Both use the
unchanged locked 0502 probe and Rust/Cargo 1.95.0. A 400-sample pilot preceded
acceptance; it is not pooled with the formal capture.

The serial A1/B1/B2/A2 capture pins each child to CPU 2 and runs 25 warmups and
200 samples per corpus per child: 16 children and 3,200 measured samples.
Fixture generation and hashing are outside the timer. Owned-byte copy/open,
semantic traversal, checksum assertion and drop are inside it. Input hashes
and traversal checksums agree, but that checksum does not cover all metadata;
source review and focused tests provide additional semantic evidence.

| Corpus | Before p50 ms, r1 / r2 | After p50 ms, r1 / r2 | p50 change, r1 / r2 | RSS change, r1 / r2 |
| --- | --- | --- | --- | --- |
| plain-small | 1.399 / 1.408 | 0.938 / 0.947 | -32.94% / -32.74% | +0.49% / -0.33% |
| plain-large | 42.474 / 42.401 | 28.239 / 28.214 | -33.51% / -33.46% | -9.08% / -9.08% |
| metadata-small | 1.312 / 1.312 | 1.088 / 1.075 | -17.06% / -18.10% | -2.84% / -3.46% |
| metadata-large | 18.180 / 18.185 | 14.614 / 14.613 | -19.61% / -19.64% | +3.45% / -1.72% |

No paired measured latency, throughput or RSS regression exceeds 5%. Every
p95/p99, throughput and RSS pair is retained in the summary, together with
10,000-resample median-ratio intervals using the fixed 0502 seed. Intervals
measure within-run uncertainty, not variation across hosts. Two reversed
repeat pairs do not promote a claim into the separate strict registry.

Plain-large RSS falls 22,544/22,560 KiB to 20,496/20,512 KiB, a 2,048 KiB
reduction in each pair. This fresh comparison improves on the prior batch's
RSS cost; it does not establish the allocator mechanism behind either change
or replace the historical 0506 review evidence.

## Allocation and instruction attribution

| Whole-child heaptrack corpus | Allocation calls before → after | Change | Rounded peak heap before / after |
| --- | --- | --- | --- |
| metadata-large | 371,676 → 262,812 | -29.29% | 6.98M / 6.98M |
| plain-large | 998,535 → 539,783 | -45.94% | 12.57M / 12.57M |

Temporary-allocation counts fall 247,651 to 146,467 and 713,414 to 287,430
respectively. Counts include setup and two opens (preflight plus one sample),
with zero warmups. They are not isolated per-open allocation totals. Equal
rounded peak displays do not prove exact byte equality; profiler RSS is
separate from the formal uninstrumented process measurement.

Metadata-large Callgrind references fall 962,330,213 to 797,288,072 (17.15%).
`parse_content` falls to 349,438,258 inclusive references. The original
`attribute` helper now accounts for 153,191,355 references, and the batched
helper for 37,881,764. Other value lookups remain; this is not elimination of
all repeated parsing. Whole-child profiling includes untimed setup and hashing.
A fresh hardware-counter capability probe fails under `perf_event_paranoid=4`;
these are instruction references, not hardware cycles, IPC or cache metrics.

## Correctness, architecture and remaining scope

The owner suite passes 124 unique tests: the prior 108 plus nine private
differential tests and seven public integration tests. Coverage includes all
selected property groups, namespace aliases/rebinding, foreign/default/unbound
names, duplicate request slots, empty request arrays, irrelevant attributes,
malformed raw input, and exact error-message parity against ordered scalar
lookup. Public cases check frame/3D/z-index/geometry validation precedence,
normalized values, exact semantic no-op bytes, edit/reopen and inverse
restoration. Numeric character references retain their referenced characters;
literal attribute whitespace normalizes under the existing XML decoder.

An initial private-test compile failed because a helper's returned reader
needed an explicit lifetime tied to its XML input. The corrected test signature
passes; the failed log is retained. No production code changed after timing.
Formatting, warning-denying Clippy, doctests, warning-denying rustdoc, the ODF
umbrella's ODG-only build, and crate boundaries are recorded in `gates.json`.

[The ADR matrix](../results/change-0507/adr-review.md) records the additional
value/error obligations against all unchanged accepted-ADR inputs. Successful
batches check the complete raw iterator; error replay drops temporary output
before the original failing path. Existing source/shape/depth limits and atomic
publication remain active. The helper introduces no concurrency or lazy state.

These synthetic owned-byte open/traversal scenarios do not establish full CRUD,
remote/provider, cold-cache, scaling, or live native-application performance.
No full-workspace build, sanitizer or fuzz campaign was run; the owner suite
includes existing fixtures. The strict registry remains at ten entries. Broader
scenario coverage and the comparison to the older, less expressive 0502 parser
remain separate work. Batch scratch is removed after retaining and hashing the
measured evidence and final checks.
