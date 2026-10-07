# 0964 raw reader and namespace transport source review

Status: independent, read-only static review of the external candidate. No
Rust source in the production checkout, build, test, benchmark, or workload
was run by this review. Runtime quality and performance gates remain pending.

## Source scope

The reviewed trees are `/home/zhuhe/litchi-performance-0964/source-baseline`
and `/home/zhuhe/litchi-performance-0964/source-candidate`, based on the
current repository source at `93d32552d9`. The candidate production diff is
confined to:

* `crates/litchi-xlsx/src/cell_values/validation.rs`, where
  `validate_xml` uses `Reader` plus a local `NamespaceResolver`; and
* `crates/litchi-xlsx/src/raw/worksheet/codec.rs`, where
  `parse_source_with_observer` uses the same transport.

The candidate also adds focused checks in
`crates/litchi-xlsx/src/cell_values/shared_traversal_tests.rs` and
`crates/litchi-xlsx/src/cell_values/validation_borrow_tests.rs`. The final
custody record must list both test paths individually, as required by the
0964 protocol; the test allowlist cannot remain an implicit wildcard.

The reviewed production file hashes were:

| file | baseline | candidate |
| --- | --- | --- |
| `cell_values/validation.rs` | `183ab8ebb334a3d15e5d0312350915a756467fb578826833da17ac0565d6d53d` | `f588cd2b1983c96df3d58a07c21ba5b4c14a57e776e35d5ac152159d788cb39e` |
| `raw/worksheet/codec.rs` | `94b5c115d0e44b51e66bc934a874eb51283567759ccf649a69422f3e096ea2bb` | `0ed7754f2b6ebb9eb16fcdb9896d0383c471709af2579770138936147e6bebaa` |

## Transport equivalence

The workspace is pinned to quick-xml 0.41.0. Its `NsReader::read_event_impl`
first consumes a deferred pop, reads one event, and then processes that event.
Processing pushes declarations for `Start` and `Empty`, and marks `Empty` and
`End` for the next pop. The candidate reproduces this order in both loops:

1. pop immediately before the underlying `Reader::read_event`;
2. push a `Start` or `Empty` scope before post-event positions, event counting,
   namespace resolution, or callbacks;
3. keep an `End` in its active scope while resolving and consuming it;
4. defer `Empty` and `End` pops until the next read; and
5. resolve only element events with `resolve_prefix(..., true)`, leaving all
   other events `Unbound`.

The raw worksheet loop retains the baseline order after the read: it computes
the BOM-adjusted end offset, updates the event counter, resolves the element,
runs `MceRewriteEquivalence`, calls the observer, and then transitions the
parser. Resolver failures return `ReaderFailed` before any observer or parser
consumer, matching the old `NsReader` failure boundary. The validation loop
converts `NamespaceError` through `quick_xml::Error::from` before formatting
the existing `Invalid` diagnostic, preserving the baseline error namespace and
priority over validator policy.

`NamespaceResolver::default()` retains quick-xml's reserved-prefix rules and
256-declaration per-element bound. No resolver setting, reader configuration,
parser limit, event-count cap, or callback order changed. The existing
`Reader::check_end_names = true` setting remains explicit in the raw loop and
is the default used by the validation loop, as it was through `NsReader`.

## Offsets, MCE, and fallback boundaries

`parse_source_with_observer` still constructs `ReaderOrigin::of(content)` and
uses it for both event offsets. Since quick-xml removes a leading UTF-8 BOM
from its reader position, this preserves the source-relative start and end
spans without double-counting the three skipped bytes. The new frozen source
test compares the candidate's borrowed event observations with an independent
`NsReader` oracle on a BOM-prefixed, nested namespace-rebinding fixture.

`MceRewriteEquivalence` still observes the raw event only after the current
namespace scope has been installed and resolved, and before the validator
observer and raw parser. Any proof refusal, event-cap refusal, parser refusal,
or reader/namespace failure still drops the provisional path and invokes the
existing authoritative validation and refusing parser. The source admission
and MCE fallback policy, output publication, cancellation checks, and resource
cleanup are outside the changed transport and remain unchanged in the source
diff.

The retained fallback and MCE tests cover declaration-only markers, valid
Ignorable attributes, rewrite-only refusals, declaration limits, real
worksheet admission, and exact authoritative error/value comparison. The new
source transport tests cover BOM spans, scope restoration, reserved namespace
failures, and the declaration ceiling. The validation differential tests add
reserved `xml`/`xmlns` bindings and the 256/257 declaration boundary while
continuing to compare the manual loop with the frozen `NsReader` reference.

## Static disposition

**PASS for static transport equivalence.** I found no typed-reader,
namespace-error, BOM-offset, cap-ordering, MCE-observer, fallback, borrowed
event, or frozen-oracle blocker in the two production loops. This is a source
review only: the candidate still requires the root-owned quality suite,
including the full differential and boundary tests, source-freeze and
allowlist verification, and the prescribed native and allocator screens.

Evidence and precedent references: quick-xml 0.41.0 source under
`/home/zhuhe/.cargo/registry/src/index.crates.io-1949cf8c6b5b557f/quick-xml-0.41.0`,
`docs/performance/results/change-0810/source-review.md`,
`docs/performance/results/change-0823/candidate-review.md`,
`docs/performance/0810-pptx-direct-event-workflow.md`, and
`docs/performance/0823-pptx-scene-reader-workflow.md`. The 0964 protocol is
`/home/zhuhe/litchi-performance-0964/protocol/README-prospective-v1.md` and
`plan-prospective-v2.json`.
