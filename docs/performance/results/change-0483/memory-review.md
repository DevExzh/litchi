# DOCX tail-append memory review

Status: source-derived review of the requested-storage envelopes in the
source-backed DOCX tail route. The review covers the pinned `quick-xml
0.41.0` owners and the DOCX/OPC limit handoffs. It is not an allocator,
process-RSS, package-wide, or peak-resident-memory measurement. Caller-owned
`BufRead` storage, allocator metadata, package/cache retention, ZIP
decompression, publication buffers, and error-string storage remain outside
this document.

The applicable accepted constraints are ADR 0005's finite hierarchical
budgets, explicit scratch ownership, and fallible reservation rules, and ADR
0006's deterministic fail-closed validation rule. The independent XML audit
is recorded in
`docs/performance/results/change-0482/xml-audit-review.md`; the MCE arithmetic
has a separate review in `mce-memory-review.md`.

## Locked `quick-xml` scanner owners

The scanner envelope is implemented by
`scanner_workspace_requirement` in
`crates/litchi-docx/src/source_backed/tail_append.rs`. It is checked `usize`
arithmetic converted to the execution-budget type only after the complete sum
succeeds. Its terms correspond to owners that exist before semantic closure
decides whether an event is admissible.

* `GuardedBufRead` reserves exact captured and exposed token windows. The
  caller also reserves the event buffer. The envelope therefore charges four
  `T` byte windows, where `T = max_token + 1`; the fourth covers an unknown
  qualified-name value returned by `NsReader`.
* The locked reader retains open element names and their start indexes in
  `ReaderState::opened_buffer` and `ReaderState::opened_starts`. A start event
  is recorded before the DOCX depth check, so the parser admission floor is
  `L = max_depth + 1`.
* `NsReader` resolves namespaces before returning the event. Its resolver
  retains a byte buffer and `NamespaceBinding` vector; popping a scope
  truncates lengths without shrinking capacities. The envelope includes the
  73-byte initial resolver bindings, all `L` levels, and up to the configured
  256 declarations per level (also limited by `T`).
* QuickXML's attribute iterator retains duplicate-check ranges and a hash
  table. The DOCX opaque-attribute validator also reserves its
  `Vec<(&[u8], &[u8])>` before iterating. Opaque attributes are admitted by
  the token window, so `A = T` is the conservative event attribute count.
* The settings guard owns a namespace-scope vector of `(usize, usize)` pairs.
  The main story scanner owns its `Vec<FrameKind>` scope stack. Both are
  charged separately; the fixed three-byte BOM probe is included.

The private dependency sources backing these owners are
`quick-xml-0.41.0/src/reader/state.rs`,
`quick-xml-0.41.0/src/reader/ns_reader.rs`,
`quick-xml-0.41.0/src/name.rs`, and
`quick-xml-0.41.0/src/events/attributes.rs`.

## Scanner envelope

The current helper is equivalent to the following checked formula. `max`
denotes the floor used by the implementation; every multiplication and
addition is checked.

```text
T  = max_token + 1
L  = max_depth + 1
M  = min(256, T)
A  = T
U  = size_of::<usize>()
B  = 4 * U                         // practical NamespaceBinding size
R  = size_of::<Range<usize>>()
F  = size_of::<FrameKind>()
H  = size_of::<u64>() + size_of::<u8>()
P  = size_of::<(&[u8], &[u8])>()   // opaque seen-pair slot

windows               = 4 * T
opened_names          = max(8, 2 * L * T)
opened_indexes        = max(4 * U, 2 * L * U)
namespace_bytes       = max(128, 2 * (73 + L * T))
namespace_bindings    = max(8 * B, 2 * (2 + L * M) * B)
attribute_ranges      = max(4 * R, 2 * (A + 1) * R)
attribute_hash        = max(8 * H, 4 * (A + 1) * H)
attribute_seen        = max(8 * P, 2 * (A + 1) * P)
scope_stack           = max(8 * F, 2 * L * F)
namespace_scope_stack = 2 * L * size_of::<(usize, usize)>()
fixed                 = 3

scanner_workspace = checked_sum(
    windows,
    opened_names,
    opened_indexes,
    namespace_bytes,
    namespace_bindings,
    attribute_ranges,
    attribute_hash,
    attribute_seen,
    scope_stack,
    namespace_scope_stack,
    fixed,
)
```

The factors of two are the pinned requested-capacity convention used by the
XML audit helpers. They are a practical allocation envelope, not a statement
about allocator metadata or physical RSS. `namespace_bindings` includes the
two implicit resolver bindings and the declaration vector's geometric
capacity. `attribute_seen` accounts for the opaque duplicate-check vector in
addition to QuickXML's own range/hash scratch.

On a 64-bit target, the default `max_token = 65,536` and `max_depth = 256`
give `T = 65,537`, `L = 257`, and a scanner requirement of **80,509,015
bytes (76.779 MiB)**. The terms are:

```text
windows               262,148
opened_names       33,686,018
opened_indexes          4,112
namespace_bytes    33,686,164
namespace_bindings  4,210,816
attribute_ranges    2,097,216
attribute_hash      2,359,368
attribute_seen      4,194,432
scope_stack               514
namespace_scope_stack  8,224
fixed                       3
```

For the finite harness profile `max_token = 4,096` and `max_depth = 16`, the
same scanner helper requires **1,115,575 bytes (1.064 MiB)**. The historical
formula that omitted resolver storage, opened indexes, duplicate scratch,
opaque seen-pairs, and depth-plus-one admission is superseded. The historical
64 MiB default is also superseded: the current DOCX default is 128 MiB. A
caller that selects a 2 MiB workspace must use an explicit narrower token and
depth profile; the larger defaults must not be silently accepted with a
smaller reservation.

The root-dialect probe calls the same helper with depth zero and keeps its
reservation nested and separate from the settings guard. It source-mins the
token window and caps resolver declarations at `bytes.len().clamp(1, 256)`.
The settings guard source-mins both token bytes and depth (`bytes / 3`),
checks the child depth before admitting a `Start` or `Empty` event, and
reserves its own `max_depth + 1` namespace-scope capacity. These facts are
part of the current admission contract; the old `2 * T` root reservation is
historical.

One small owner detail remains for the main story route: `scan_reader`
currently calls `try_reserve_exact(max_depth)` for its `Vec<FrameKind>`, while
the parser and the envelope admit one over-limit start before the semantic
depth refusal. The envelope's `2 * L * size_of::<FrameKind>()` term covers the
usual geometric growth, but exact no-growth ownership would reserve
`max_depth + 1` there as well. This is a source-owner cleanup, not a reason to
reduce the checked envelope.

## Settings admission topology

The settings path now has distinct source-derived phases. This replaces the
historical single-lease description and avoids charging the root probe twice.

1. The package verifies the declared settings size and obtains the existing
   `PartData` bytes. The root-dialect probe uses its own nested scanner lease.
2. `guard_settings_xml` scans the original bytes with the complete event/name,
   declaration, depth, token, event-count, namespace, attribute, and decoded
   MCE-directive checks. It returns `SettingsGuardFacts`; its scanner lease is
   released before the MCE lease is acquired.
3. `bounded_settings_mce_limits` derives depth, namespace, and directive
   ceilings from the authenticated facts and source length. MCE input remains
   bounded by the source length. MCE output has the explicit
   `max_settings_xml_bytes` ceiling, because namespace reinjection can make
   valid processed XML larger than its source.
4. `mce_workspace_requirement` calls the allocation-free
   `tail_append::mce_workspace` helper. The result includes the output owner.
   The caller splits it into an output reservation and MCE scratch
   reservation, runs exactly one MCE pass, checks the phase boundary, drops
   scratch, and retains the output reservation only when the result is owned.
   A borrowed MCE result releases that output lease.
5. The processed bytes are guarded again. This is required because MCE can
   reinject namespace declarations and change token, depth, resolver, node,
   attribute, or directive facts. The processed guard receives the remaining
   workspace after any retained MCE output owner, then releases its scanner
   lease.
6. `settings_workspace::model_memory_requirement` computes the complete model
   envelope from the processed facts and processed byte length. The caller
   checks `retained_output + model_workspace` before reserving the model
   lease, then runs the complete borrowed settings/mail-merge/extension model
   validation against the processed XML and the original relationship map.

The package helper `process_bytes_with_mce_limits` is the one MCE boundary;
`extract_from_processed_xml_with_relationships` deliberately uses
`extract_from_processed_xml`, so the complete model path does not preprocess
MCE a second time. Relationship validation consumes the original map through
borrowing helpers; it does not clone every relationship string into a staged
settings blob. The `PartData`/package cache owner remains an OPC owner and is
outside this XML workspace lease.

## MCE and model owner terms

`mce_workspace.rs` uses checked versions of the pinned requested-capacity
functions:

```text
Str(bytes, count) = 2 * bytes + 8 * count
V<T>(count)       = max(8 * size_of::<T>(), 2 * count * size_of::<T>())
H<T>(count)       = max(8 * (size_of::<T>() + 1),
                         4 * (count + 1) * (size_of::<T>() + 1))
```

Its sum covers the caller event buffer, QuickXML open names/indexes, MCE
frames and names, raw decoded attributes, QuickXML attribute ranges/hash,
Arc-shared namespace and directive layers, directive vectors/sets/patterns,
transient expanded names, capabilities, fixed reader/report state, and the
separately retained output vector. The namespace layers and `Ctx` use Arc
handles; the audit does not multiply their contents as if each frame deeply
cloned an inherited context. Conversely, local declaration strings and
accepted directive patterns are charged as owned data.

`settings_workspace.rs` separately charges the complete model passes:

* the private mail-merge `Node` tree, child and stack vectors, attribute
  headers/slots, semantic owned strings, event storage, QuickXML opened names,
  and resolver bytes/bindings;
* the mail-merge model and field-map vectors while the tree is live, including
  its bounded relationship-ID strings;
* `Extensions` opaque ranges, the source-range/output overlap while an
  unknown child is made self-contained, extension values, active-binding and
  declaration scratch, and a resolver for generated namespace declarations;
* direct `DocumentSettings` vectors, decoded semantic strings, relationship
  and attached-template target limits, and their reader/resolver owners.

Unknown extension children are charged with the conservative
`nodes * max_namespace_buffer * 6` plus declaration syntax/header term. This
specifically covers unused long namespace URIs copied into every retained
unknown child; counting only resolved element names would undercharge that
owner. The helper also applies the codec's finite 128-extension and 4 MiB
opaque-child ceilings. `namespace_copy_bytes` is collected by the guard but
the current helper uses the conservative fallback rather than that fact; it
is a possible tightening/cleanup item, not an undercharge.

All three owning model passes are precharged as a bounded phase using three
times the processed event count. The preflight checks cancellation per event;
the MCE and model paths check before and after their bounded owning phases.
This is an explicit cancellation granularity for the complete model call and
does not claim an internal per-event cancellation point inside every codec.

The full parser remains authoritative for protection values, duplicate/order
rules, mail-merge schema and relationship checks, extension semantics, and MCE
fail-closed behavior. The preflight supplies admission facts and workspace
planning; it is not a security-only replacement for the model parser.

## Policy and OPC audit consistency

The current DOCX defaults are `max_token_bytes = 64 KiB`, `max_depth = 256`,
and `max_workspace_bytes = 128 MiB`. The 128 MiB setting leaves room above
the 76.78 MiB scanner envelope; the source-derived MCE/model terms still
perform an aggregate workspace check and can refuse a profile whose complete
settings owners do not fit. The 2 MiB harness profile is coherent for its
4 KiB/16-depth scanner and small source-derived fixtures.

`make_splice_limits` constructs `xml_minifier::audit::Limits` with
`Limits::new`, so DOCX's explicit profile is applied rather than being
silently narrowed by `Limits::narrow`. It passes the DOCX depth, event, byte,
text, and token values. Its aggregate `max_attributes` is a finite
candidate-byte-derived count, bounded by the audit crate's immutable
attribute ceiling; it is not `max_events` and it is not the token window.
That count admits valid documents containing many small attributes. The
locked XML helper still computes per-event duplicate scratch with
`min(max_attributes, max_token + 1)`, so raising the aggregate policy does
not silently turn the per-event scanner envelope into an aggregate-sized
allocation.

The aggregate attribute count and the token-sized scratch are therefore two
separate policy terms. Any future increase to the token limit must update the
scanner and XML-helper envelopes; any future increase to the aggregate count
must remain within the immutable audit ceiling and its finite candidate-byte
proof.

The main-story `FrameKind` vector now reserves `max_depth + 1`, matching the
parser's depth-plus-one admission and the scanner envelope's scope term. This
closes the earlier exact-reservation gap. The settings path's `Cow` ownership
test also does not move the value when its pattern binds no owned fields;
subsequent uses of `processed` remain valid.

## Remaining review items

The only concrete accounting cleanup visible in this checkpoint is the model
helper's conservative extension namespace fallback, which is wider than the
collected `namespace_copy_bytes` fact. A later tightening can use the fact
only after proving it includes every active-binding/self-contained-copy owner;
removing the fact or wiring it into the helper should be done as one
owner-contract change.

The earlier scanner-stack and `Cow` concerns are resolved and are retained
above only to prevent the historical review text from being mistaken for a
current blocker.

These are requested-storage and admission observations. No statement here is
an RSS bound, allocator guarantee, throughput result, or claim that package
cache and publication owners overlap at a particular peak.
