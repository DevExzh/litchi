# 0540 worksheet validation local-name ownership review

This is a read-only source and evidence review. It covers the current
`crates/litchi-xlsx/src/cell_values/validation.rs`, the pinned `quick-xml`
0.41.0 source, and the prior XLSX/DOCX ownership records. No Rust source was
edited, and this review did not build or run a candidate.

## Current ownership boundary

The current validator has three different local-name lifetimes:

| Event path | Current code | Lifetime required by the path |
| --- | --- | --- |
| `Start` | `validation.rs:45-65` copies `element.name().local_name()` into a boxed slice and pushes it into `elements` | The name must survive until the matching `End`, and while all nested events use it as their parent. |
| `Empty` | `validation.rs:67-79` copies the local name into a temporary `Vec` | It is consumed by `validate_element` in the same event arm and is never put on the open-element stack. |
| `End` | `validation.rs:80-93` pops the owned expected name and compares the event's local name | The event name is only needed for this comparison and the fixed error result. The previously opened name must already survive in `elements`. |

The persistent allocation is `elements: Vec<Box<[u8]>>` at
`validation.rs:34`. A `Start` name is copied at lines 47-52, used for the
namespace/grammar/parent/attribute checks at lines 53-60, and then retained at
line 65. The `Empty` copy at line 69 and the `End` view at line 87 are
event-local. Thus, a blanket replacement of every `to_vec()` is not the
correct boundary: `Start` needs a persistent representation, while `Empty`
and `End` can already use a borrowed name during their event arms.

The parent stack is semantically active in more places than
`validate_parent`. `text_allowed` and `text_context_allowed` at
`validation.rs:156-162` use the top name to permit text only beneath `f`, `v`,
or `t`; the `Text`, `CData`, and `GeneralRef` arms pass that stack view at
lines 101-127. A replacement must preserve those calls and their event order.

## Why a direct borrowed slice does not fit the existing stack

The input is a slice-backed reader. In the pinned `quick-xml` source,
`reader/slice_reader.rs:244-246` documents that `&[u8]` events borrow from the
input, and `reader/ns_reader.rs:455-520` makes `NsReader<&[u8]>::read_event`
return `Event<'i>`. That establishes that the bytes themselves remain
available for the whole validation call.

There is a more restrictive lifetime at the public event-name API, however.
`BytesStart::name` and `BytesStart::local_name` return views from `&self` at
`events/mod.rs:215-227`; the corresponding `BytesEnd` methods do the same at
`events/mod.rs:442-455`. `QName::local_name` and `LocalName::into_inner` in
`name.rs:125-127` and `name.rs:219-225` preserve the lifetime of the name view
they receive; they do not turn a borrow of an event object into an independent
`'i` borrow of the original input.

The proposed consuming `BytesStart::into_inner() -> Cow<'input, [u8]>`
escape hatch is not present in the pinned `quick-xml` 0.41.0 API. `BytesStart`
stores its `Cow` buffer and `name_len` as `pub(crate)` fields
(`events/mod.rs:93-100`), while its public ownership/name methods are only
`into_owned`, `borrow`, `name`, and `local_name` (`events/mod.rs:144-227`).
Those methods cannot expose the borrowed buffer to a caller after consuming the
event. `BytesEnd` likewise exposes `into_owned`, `borrow`,
`name`, and `local_name`, with no `into_inner` (`events/mod.rs:406-455`); the
`into_inner` methods elsewhere in this module belong to text/CDATA/PI/reference
events. `BytesStart`'s public `Deref` returns the complete start-tag content
(`events/mod.rs:346-351`), and `attributes_raw` returns the suffix after the
name (`events/mod.rs:296-301`), so a caller can infer the name boundary by
length while the event is borrowed. Neither view, however, yields the
underlying `Cow` after consuming the event; the resulting local slice still
has the event borrow. Therefore a safe `Cow`-based stack is unavailable
through this pinned public API. This is a version/API boundary: if a future
quick-xml release adds a consuming buffer accessor (or an equivalent
event-to-input-slice API), this conclusion should be re-audited.

Consequently, the tempting forms below do not provide a valid persistent stack
entry through this API:

```text
Vec<&'input [u8]>
Vec<LocalName<'input>>
```

The local name obtained from `element.local_name()` is borrowed through the
`BytesStart` value bound by `Event::Start(element)`. That value is dropped when
the match arm ends. Binding the `LocalName` wrapper first is necessary for an
event-local use, but does not extend its borrow past that event. The prior 0481
source review records the related temporary-wrapper trap: assigning
`start.local_name().as_ref()` directly can produce `E0716`; the wrapper must be
bound before using it (`docs/performance/results/change-0481/source-review.md:20-37`).
Here, even the wrapper-first form cannot be pushed into a stack that outlives
`element`.

This is a type/lifetime issue rather than a claim that the slice reader copies
the name bytes. Relying on the current internal `Cow::Borrowed` representation,
casting the view to a longer lifetime, or using `unsafe` to bypass the event
borrow would violate the safe API boundary and the repository's safe-Rust
policy. `LocalName::into_inner()` is useful for a borrow that ends in the same
event arm; it does not make the persistent `elements` stack safe.

The event-local cases are different. A wrapper bound in the `Empty` or `End`
arm can be passed to `validate_element` or compared before the event is
dropped. This is the same ownership shape used by the accepted 0481 DOCX
scanner, where no name enters a returned layout or a cross-event stack.

## Resolver and error-order constraints

The namespace result has its own independent lifetime. `NamespaceResolver::resolve_event`
resolves only `Start`, `Empty`, and `End`, returns the event unchanged, and
returns a bound namespace borrowing the resolver (`quick-xml` `name.rs:881-942`).
The resolver's delayed scope pop is implemented at `reader/ns_reader.rs:64-98`.
The current loop resolves and consumes the result before the next
`read_event`, so that scope behavior must remain. The first dialect is copied
into an owned box by `bind_dialect` at `validation.rs:137-153` because the
resolver's internal namespace buffer can change on later events; a local-name
candidate must not accidentally replace that owned state with a resolver
borrow.

The following order is observable and must remain unchanged:

* `Start` and `Empty` call `bind_dialect` before `validate_element`
  (`validation.rs:45-46` and `67-68`).
* `validate_element` checks the resolved namespace, allowed local name, parent,
  attributes, and root depth in that order (`validation.rs:164-226`). Local
  bytes are also used in the exact refusal messages at lines 174-176 and
  212-215, and attribute checks use the same local view at lines 266-289.
* A successful `Start` performs checked depth increment before pushing the
  persistent stack entry (`validation.rs:61-65`).
* `End` performs checked depth decrement before popping, then compares the raw
  closing local name and the current dialect namespace (`validation.rs:80-93`).
  The `Some(expected)` in the `matches!` expression is the dialect box; it is
  not the popped opening name. A refactor must preserve this exact current
  namespace/error behavior.
* `Text`, `CData`, and `GeneralRef` consult the current parent after the event
  is resolved (`validation.rs:101-127`), and the complete-root check remains at
  lines 131-134.

The larger source-backed load path also requires validation to precede the raw
worksheet parser. The 0539 next-priority review identifies
`worksheet_xml -> validate_xml` and `raw::worksheet::parse` as sibling stages
and explicitly retains validation-first error precedence
(`docs/performance/results/change-0539/next-priority.md:17-20,30-34`). A
local-name representation change must not fuse those stages, defer the
validator, or use the parser's names as a substitute for this validator's
grammar and refusal checks.

## Safe candidate shapes

The direct `Vec<&[u8]>`/`Vec<LocalName>` idea should be rejected for the
existing stack. Two safe alternatives are technically available, with
different costs.

The preferred shape, if fresh attribution shows that these short-name
allocations are material, is a private validated-name representation: an
exhaustive enum or another compact discriminant for the finite names admitted
by `validate_element` (the workbook and worksheet tables are at
`validation.rs:179-209`). The `Start` arm can borrow a wrapper for the existing
checks, and only after `validate_element` returns `Ok` map the exact bytes to a
discriminant and push it. Parent and text helpers can project that
discriminant to static byte literals; `End` can compare its event-local name
against the same projection. `Empty` and `End` need no copy.

The mapping must happen after the complete `validate_element` call. Unknown,
foreign, wrong-parent, wrong-root, and attribute errors must still format the
original borrowed bytes before any mapping, and no `unreachable` branch should
be used to bypass those checks. The push must stay after the checked depth
increment. With those conditions, every value entering the stack is one of the
same byte literals that the existing allowed-name table accepted, so parent
classification, text admission, closing-name comparison, and success/refusal
semantics can be identical. This is a candidate for a future measured patch,
not an adopted change in this attribution batch.

A second possible design is to retain the full borrowed `BytesStart<'input>`
event in the stack and call `local_name()` on the retained event when needed.
Moving the event itself can preserve its input lifetime without storing a
borrow into the event, and the slice reader's events are borrowed. This is
more intrusive and retains the complete raw start-tag wrapper (including
attribute span metadata and decoder state) for every open element, while the
current stack retains only a short local name. It can therefore increase
per-depth live storage and working-set pressure, especially on deeply nested
or adversarial input. It should be considered only with explicit memory and
depth evidence; it is not the same as borrowing a local-name slice.

Deriving source offsets by rescanning the input, depending on private
`quick-xml` fields, or extending the lifetime with `unsafe` is not a suitable
candidate. Those approaches add parser work or violate the ownership/safe-Rust
boundary without proving a benefit.

## Prior evaluated mechanisms

| Record | Disposition | What it establishes for this review |
| --- | --- | --- |
| 0521, `docs/performance/changes/0521-xlsx-borrow-validation-events.md` and `docs/performance/results/change-0521/borrow-design-review.md` | Accepted event/resolver borrowing | 0521 removed `Event::into_owned()` and the per-event resolver clone, but deliberately kept stack/name ownership. Its design review says the resolver borrow is live only until the next reader mutation, and its final README explicitly says remaining names and stacks are intentional (`docs/performance/results/change-0521/README.md:97-101`). It is the closest prior XLSX validator result and does not authorize borrowing a cross-event name stack. |
| 0481, `docs/performance/changes/0481-docx-borrowed-scanner-names.md` and `docs/performance/results/change-0481/source-review.md` | Accepted event-local name borrowing | The scanner consumes each `LocalName` before its event buffer is cleared and stores no local name in its layout, scope stack, patch, or diagnostics (`docs/performance/results/change-0481/source-review.md:12-18,39-65`). Its successful pattern applies to the current `Empty`/`End` uses, not to the current `Start` stack. |
| 0469, `docs/performance/changes/0469-xlsx-borrowed-compaction-events.md` | Accepted event borrowing in XLSX compaction | The compactor consumes each event before the next read and retains only boolean state (`docs/performance/changes/0469-xlsx-borrowed-compaction-events.md:7-10`; the event-local writer path is described at lines 85-88). It has no persistent local-name stack and is therefore not a direct precedent for `elements`. |
| 0526 scanner attribution, `docs/performance/results/change-0526/profile-review.md` and `docs/performance/results/change-0526/root-review.md` | Diagnostic, with related scanner work kept separate | `QName::local_name` appears as a 3.6292% direct scanner child in a different raw worksheet scanner (`docs/performance/results/change-0526/profile-review.md:15-25`). The root review explicitly distinguishes that scanner's unused closing namespace from this value-only validator, which uses the resolved `End` namespace and cannot take that shortcut (`docs/performance/results/change-0526/root-review.md:38-48`). It is not evidence that validator names or namespace checks can be skipped. |
| 0522 cell-reference fusion, `docs/performance/changes/0522-xlsx-cell-reference-guard.md` | Rejected native candidate | It lowered instruction/allocation counts but failed repeatable native admission; the record keeps full validation, semantic readback, and error order (`docs/performance/changes/0522-xlsx-cell-reference-guard.md:3-9,86-99,122-132`). It is a guard against treating a local allocation reduction as an end-to-end result. |
| 0539 transient attribute ownership, `docs/performance/changes/0539-xlsx-transient-ownership-rejected.md` | Rejected native candidate | Its planning allocation calls fell 13.6–13.7%, but native workflow gates failed in three of four rows; the record explicitly says allocation savings do not override workflow gates (`docs/performance/changes/0539-xlsx-transient-ownership-rejected.md:18-24`). A local-name candidate needs its own source, native, allocation, and semantic evidence and must not be revived or combined with 0539. |

The prior records therefore separate two cases: event-local views have been
successfully adopted, while the current cross-event stack has intentionally
remained owned. No prior accepted or rejected change demonstrates that a
`Vec<&[u8]>` local-name stack is valid through the public `quick-xml` API.

## Recommendation

Do not implement the direct borrowed-slice replacement in the existing
`Vec<Box<[u8]>>` stack through the pinned `quick-xml` 0.41.0 public API. It
cannot safely outlive the `BytesStart` event through the available name
methods, even though the underlying input bytes are immutable and slice-backed;
the proposed consuming `Cow` route is unavailable because `BytesStart` does
not expose `into_inner` (and its internal buffer remains private). Keep the current owned stack
while the 0540 attribution records the actual validation/parser boundary. A
future quick-xml API that exposes those pieces would warrant a fresh lifetime
and semantics review rather than this categorical conclusion being reused.

If a future profile isolates enough local-name ownership to justify a
candidate, use the post-validation private discriminant design and measure it
as a standalone source change. Required differential coverage includes the
ordinary, Transitional, Strict, prefixed/default-alias and namespace-rebinding
cases; unknown/qualified/duplicate/malformed attributes; DTD/PI, UTF-8,
truncation, mismatched-end, text and general-reference errors; formulas and
shared formulas; and the source/resource/error-order checks listed in the 0539
next-priority review (`next-priority.md:45-55`). Candidate quality and fresh
native/allocation gates remain mandatory; attribution Ir or an expected count
reduction alone cannot admit it.

The 0540 plan is attribution-only, so this review makes no latency,
allocation, RSS, or production-speedup claim.
