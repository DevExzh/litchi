# 0811 opportunity review: remaining PPTX opened-capture hot paths

This is a read-only ranking of the current source after the retained 0810
direct-event change. It proposes no production candidate, makes no adoption
decision, and imports no historical timing pool. It uses only the two sealed
0810 **after** Callgrind publications to identify work that a fresh 0811 native
diagnosis should either confirm or reject.

The current source is commit `2894fbd628` and the only production source
changed by 0810 is
[`crates/litchi-pptx/src/notes/codec.rs`](../../../../crates/litchi-pptx/src/notes/codec.rs).
The current codec SHA-256 is
`466e504588843a0b2538fc25ce3378b6fb9bf13aa9b0d74efebcb73ca0ebab5e`.
The 0810 after profiles are the exact owner-scoped publications in
[`profile-analysis.json`](../change-0810/profile-analysis.json), with raw
numbered inputs
[`0-after.callgrind.1`](../change-0810/profiles/0-after.callgrind.1) and
[`1-after.callgrind.1`](../change-0810/profiles/1-after.callgrind.1).

## Evidence boundary

Both publications collect only `Ir` while
`namespace_uri_probe::capture_region_0793` is active. Their summaries are
`537,148,235` and `537,183,147` guest instructions. The profile reader passes
the exact owner-call, numbered-publication, empty-termination, source, output,
semantic, and self-sum checks. The average after summary used for percentages
below is `537,165,691` Ir.

The ranking uses each function's `self_ir`. Those rows are a disjoint
partition of the owner publication; nested inclusive edge rows are not added
to them. The two repeats have identical self costs for the ranked functions,
apart from small allocator/libc differences lower in the table. An edge table
is included only to show where repeated work enters the graph. Edge inclusive
costs overlap their callees and must not be summed as another cost estimate.

The profile is an attribution diagnostic. It does not establish native cycle
share, elapsed-time share, RSS, allocation change, phase fraction, or a
production speedup. The 0811 native/perf protocol is therefore a diagnosis
lane for the current source, not a benefit gate.

Callgrind runs SHA-256 through the software `sha2::sha256::compress256` path:
Valgrind masks the SHA CPUID feature, while established native capture evidence
names `sha2::sha256::x86_sha::compress` ([0810-era native capture](../../0807-pptx-current-capture-profile.md);
[the established measurement caveat](../../0645-pptx-memoized-revision-proof-design.md)).
The SHA row below can therefore rank guest instruction burden only; it cannot
rank native cycle ROI or by itself justify removing digest work.

`docs/GOAL.md` is unchanged at SHA-256
`bed4058bb76330daab8ce9d4bceff639ab3fbd7ea06634158bef41b133c4d1f1`.
All 35 goal, taxonomy, accepted-ADR, and architecture input hashes recorded
by 0810 still match. In particular, this review remains bound to the accepted
measurement, ownership, preservation, bounded-resource, and snapshot/edit
constraints rather than treating a hot symbol as permission to change them.

## Disjoint after self-cost ranking

| Rank | Function | Mean self Ir | Share of owner | Measured repetition or context |
| ---: | --- | ---: | ---: | --- |
| 1 | `sha2::sha256::compress256` | 207,605,811 | 38.648% | 142 calls from the large payload `feed` edge dominate this symbol. |
| 2 | `quick_xml::events::attributes::IterState::next` | 35,114,190 | 6.537% | 550,297 calls arrive from namespace declaration processing; the same iterator is also used by checked-attribute paths. |
| 3 | `quick_xml::reader::Reader<R>::read_event_impl` | 34,766,287 | 6.472% | The notes scanner makes 282,612 calls in the retained direct-reader path. |
| 4 | `litchi_pptx::notes::codec::inspect_element` | 28,983,633 | 5.396% | 181,678 element inspections; the function owns the attribute, name, resolution, and unescape work below it. |
| 5 | `core::str::converts::from_utf8` | 23,550,533 | 4.384% | 364,226 calls, mostly from element and attribute name/value checks. |
| 6 | `litchi_pptx::notes::codec::scan_processed_xml` | 22,295,758 | 4.151% | One scanner loop over the processed slide XML for each captured slide. |
| 7 | `quick_xml::name::NamespaceResolver::resolve_prefix` | 19,119,253 | 3.559% | 272,952 calls from `inspect_element` in the direct path. |
| 8 | `quick_xml::reader::state::ReaderState::emit_start` | 17,427,424 | 3.244% | 363,409 start-event emissions in `Reader`. |
| 9 | `memchr::arch::x86_64::memchr::memchr3_raw::find_avx2` | 13,750,320 | 2.560% | 912,300 parser scans. |
| 10 | `quick_xml::name::NamespaceResolver::push` | 13,019,166 | 2.424% | 181,678 direct-scanner scope pushes; the resolver's own namespace and declaration checks remain active. |
| 11 | `memchr::arch::x86_64::avx2::memchr::One::find_raw` | 12,244,626 | 2.279% | Parser delimiter search. |
| 12 | `litchi_pptx::notes::resolved` | 9,753,543 | 1.816% | 272,952 namespace-result conversions/checks. |
| 13 | `__memcmp_avx2_movbe` | 9,577,860 | 1.783% | Mostly namespace/prefix and attribute comparisons. |
| 14 | `memchr::arch::x86_64::memchr::memchr2_raw::find_avx2` | 8,750,664 | 1.629% | Parser delimiter search. |
| 15 | `memchr::arch::x86_64::avx2::memchr::Three::find_raw_avx2` | 8,739,068 | 1.627% | Parser delimiter search. |
| 16 | `<litchi_opc::xml_attributes::CheckedAttributes as Iterator>::next` | 8,660,171 | 1.612% | 152,410 checked-attribute iterations in `inspect_element` and nearby callers. |
| 17 | `quick_xml::events::attributes::IterState::check_for_duplicates` | 7,691,636 | 1.432% | 556,335 duplicate-check calls under the attribute iterator. |
| 18 | `quick_xml::reader::state::ReaderState::emit_end` | 7,530,341 | 1.402% | 181,231 end-event emissions. |
| 19 | `memchr::arch::x86_64::memchr::memchr_raw::find_avx2` | 7,508,272 | 1.398% | Parser and attribute delimiter search. |
| 20 | `memchr::arch::x86_64::avx2::memchr::Two::find_raw` | 7,018,439 | 1.307% | Parser delimiter search. |

The top three disjoint functions account for 51.657% of the owner self
partition; the top ten account for 77.375%, and the top twenty account for
93.660%. Those sums are valid only because they use function self cost. The
large `memchr`, resolver, and inspector rows are descendants of parser or
scanner calls and are not additional costs when their inclusive edges are
examined.

## Hot edges and what they mean

The following edges are from the first after publication; the second repeat
has the same mechanism shape. Their inclusive values are reported to retain
the call-site evidence, not to create a second additive ranking.

| Caller edge | Calls | Inclusive Ir | Reading |
| --- | ---: | ---: | --- |
| `capture_internal → package_fingerprint_with_memo` | 1 | 207,969,787 | Complete-package revision construction is a 38.717% owner boundary in this publication. |
| `package_fingerprint_with_memo → feed` (payload feed) | 121 | 203,850,803 | The large payload update is the dominant fingerprint edge. |
| `feed → sha2::sha256::compress256` (large payload) | 142 | 203,503,587 | This explains nearly all of the SHA-256 self row. |
| `scan_processed_xml → Reader::read_event_impl` | 282,612 | 104,433,533 | Direct event transport still performs the same parser event count after 0810. |
| `scan_processed_xml → inspect_element` | 181,678 | 158,428,003 | Element inspection is the central XML validation/resolution fan-out. |
| `inspect_element → CheckedAttributes::next` | 152,410 | 37,176,938 | Attribute validation traverses the event attributes after namespace processing. |
| `inspect_element → NamespaceResolver::resolve_prefix` | 272,952 | 23,088,053 | Element and attribute names repeatedly search the current resolver bindings. |
| `inspect_element → core::str::converts::from_utf8` | 364,226 | 23,507,883 | Name and value validity remains checked on the success path. |
| `scan_processed_xml → NamespaceResolver::push` | 181,678 | 34,159,863 | Namespace scope/declaration work remains required before inspection. |
| `NamespaceResolver::push → IterState::next` | 550,297 | 21,471,516 | Namespace declaration processing scans attribute iterators of every pushed event. |

The direct reader change removed the scanner's `NsReader::process_event` edge,
but it did not remove the required namespace and attribute work. The edge
counts show a concrete repeated-work hypothesis: `NamespaceResolver::push`
walks the event's attributes to discover namespace declarations, and
`inspect_element` walks checked attributes again to apply the notes policy.
The data does not prove that every byte or every duplicate check is repeated
in an equivalent way, so a candidate must establish that relationship with a
source-level oracle and a fresh native measurement.

## Ranked opportunities

### 1. Complete-package revision and payload digesting

This is the highest absolute self-cost in the after owner’s guest profile:
207.606M Ir in `sha2::sha256::compress256`. The source path is
`capture_internal → package_fingerprint_with_memo`; the current cold route is
`Revision::Cold` in `crates/litchi-pptx/src/opened/model.rs`, where every
payload miss is hashed before the snapshot publishes its complete-package
revision and its per-part digest memo. The existing parent-digest route already
reuses hashes for shared payload allocations on later captures.

The bounded opportunity is narrower: a future candidate may remove an actually
repeated digest at a safe source/version seam, or establish a separately
justified early digest/provenance path whose total end-to-end cost is lower.
`Revision::Cold` has no parent proof and cannot reuse a digest merely because
the bytes appear unchanged; capturing a digest before cold capture is itself
work. It is not safe to replace the semantic revision with
`exact_source_sha256`, a ZIP CRC, a file timestamp, an address, or a caller
assertion. The semantic revision covers the package root relationships,
non-part members, each part's name and content type, payload bytes, and part
relationships. Physical ZIP layout and metadata are not interchangeable with
that graph proof, and an in-memory package can change without retaining an
exact source archive.

Any future implementation must retain all of the existing revision and patch
semantics:

* a hit must be tied to an allocation or explicit immutable source identity
  that proves the exact bytes being hashed;
* the memo must retain only package-owned bytes, project across a rebind, stay
  bounded by the operation's finite policy, and use fallible reservations;
* a miss must recompute the ordinary SHA-256 result and no memo state may
  change a value, refusal, limit, or published byte;
* mutations, source changes, relationship changes, unknown members, and
  cancellation must invalidate or bypass any reusable result; and
* the digest algorithm and complete graph coverage remain the authorization
  proof for snapshots, commits, patches, removal plans, and cross-document
  operations.

The current 0811 diagnosis should first determine whether native cycles follow
this guest-Ir ranking. Native evidence must identify the hardware
`sha2::sha256::x86_sha::compress` cycles and separate cold-open from
repeated/memo-hit cases; the guest software-SHA row cannot stand in for that
measurement. The next candidate qualification, if the native result supports
it, needs a large media-heavy package, many-small-part packages, and a package
with relationship changes but shared payload allocations. It must compare
revision bytes, stale-source refusals, exact no-op behavior,
allocation/resource counters, and all relevant commit/patch paths. No
candidate or speedup claim follows from this profile alone.

### 2. Attribute traversal and namespace resolution

This is the strongest XML-local repeated-work hypothesis. The profile records
35.114M self Ir in `IterState::next`, 8.660M in checked-attribute iteration,
7.692M in duplicate checking, 19.119M in `resolve_prefix`, and 23.551M in
UTF-8 validation. These rows are not a valid additive savings estimate, but
their call graph is specific: 181,678 element events enter `inspect_element`,
the scanner makes 181,678 namespace pushes, and the inspector performs
152,410 checked-attribute iterations plus 272,952 prefix resolutions.

A future candidate could investigate a bounded event-local representation that
lets namespace declaration discovery and notes-attribute policy share one
validated attribute walk, or a small per-event prefix-resolution cache. Such a
candidate must remain within the existing format/codec ownership boundary and
must not hand-roll an unbounded namespace map. It also cannot assume that
non-declaration attributes are irrelevant: the current policy counts them,
checks duplicate syntax, validates UTF-8, unescapes values, and may refuse an
unknown prefixed name.

The resolver's public `push`/`pop` state transition is part of the current
semantic contract. Skipping a push for an apparently declaration-free element
would be unsafe unless an equivalent bounded scope mechanism proves the same
nested lookup behavior. Reserved `xml`/`xmlns` checks, the per-element
declaration cap, empty-element pop timing, end-element timing, unknown-prefix
errors, duplicate attributes, and error ordering must remain identical. The
existing direct-vs-buffered oracle and the focused nested/rebinding tests are
necessary but not sufficient evidence for a new shared walk; add adversarial
attribute and namespace cases before measuring.

Fresh evidence should include attribute-free, four-attribute, vendor, and
unicode-vendor shapes at tiny, medium, and large sizes, with capture, commit,
and lifecycle paths. Record native p50/p95/p99, exact output and refusal
identity, allocation calls/bytes/net-live/peak-above-entry, and owner-scoped
native cycle stacks. A lower `IterState` count with a slower end-to-end result
does not pass the project measurement contract.

### 3. Byte/name conversion on the validated success path

`from_utf8`, `resolved`, and `resolve_prefix` together point to repeated name
and namespace conversion. The exact-known namespace fast path already exists
in `crates/litchi-pptx/src/notes/mod.rs`: `known_namespace` checks length and
bytes for P, PS, A, AS, R, and RS before falling back to `from_utf8`. That is
current behavior, not a missing opportunity. The remaining measured
`from_utf8` work in `inspect_element` is the local element name and the
attribute local-name/value conversions; investigate only a measured reduction
there, with a checked fallback for all other bytes. The fallback is mandatory:
malformed UTF-8, undeclared prefixes, reserved namespace bindings, and invalid
attribute values remain typed refusals with their current priority.

The candidate must also preserve the current attribute-byte accounting. Counting
raw bytes in place of unescaped value bytes, or deferring validation until after
another limit check, could change both observed limits and refusal ordering.
The candidate owner is the PPTX notes codec; changing quick-XML internals or
adding unsafe SIMD is outside this bounded review unless a separate owner and
accepted evidence permit it.

### 4. Reader and delimiter search

`Reader::read_event_impl` and its `emit_start`, `emit_end`, and `memchr`
descendants remain substantial after the 0810 wrapper removal. The current
source already uses the public direct `Reader` path, so this is evidence for a
future parser-level investigation rather than a reason to reintroduce an
internal dependency or write a second XML parser in `codec.rs`.

Any change here would need a quick-XML ownership decision or a separately
bounded local owner, scalar safe fallback, malformed-input parity, and native
cycle evidence. Guest instruction reductions alone are insufficient. The
current 0811 frame-pointer/perf lane should establish whether this parser work
is actually a native bottleneck before any implementation is considered.

## Constraints that gate the next hypothesis

The next candidate must satisfy the goal and accepted ADRs already bound by the
0810 architecture receipt:

* preserve lossless bytes, namespace choices, unknown markup and relationships;
* keep typed refusal, validation order, XML/attribute/depth/node limits,
  cancellation, and security boundaries unchanged;
* keep immutable `Send + Sync` snapshots, source-checked edits/commits/patches,
  deterministic conflicts, and exact no-op behavior;
* avoid public exposure of archive implementation types, raw locks, runtimes,
  ambient I/O, hidden thread pools, or unbounded caches;
* keep changes in the crate that owns the grammar and preserve dependency
  direction; and
* provide fresh representative profile, native timing, allocation, output,
  semantic, refusal, and resource evidence before making a performance claim.

The 0810 Callgrind profile can rank hypotheses. It cannot authorize changing
the semantic fingerprint, weakening namespace validation, replacing SHA-256,
or declaring the current source faster. The immediate 0811 deliverable is
therefore fresh native attribution and a bounded decision about which of the
two leading hypotheses—complete-package digesting or XML attribute/resolution
work—deserves a separately measured candidate.
