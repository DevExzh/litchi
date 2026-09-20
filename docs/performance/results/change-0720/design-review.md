# Independent design review: DOCX body scan fusion

This is a read-only qualification of the candidate described by
`docs/performance/results/change-0719/next-experiment.md`. It does not approve
a production change or make a performance claim. The current source is the
authority; the review was made against the unchanged `0719` source tree.

## Owner and current sequence

The actual writer owner is `DocumentBody::from_xml` in
`crates/litchi-docx/src/writer/doc/package.rs`, beginning at line 85. It
returns `ParsedDocumentBody`; that return type is not the owner and there is
no `ParsedDocumentBody::from_xml` method. The public wrapper is
`MutableDocument::from_xml` in `writer/doc/model.rs:288`, which calls
`DocumentBody::from_xml` and then validates section placement.

After stripping one leading UTF-8 BOM, `DocumentBody::from_xml` currently has
this order:

1. `scan(bytes)` in `alt/codec.rs` walks the complete source for every
   recognized `altChunk`, validates its relationship and child grammar, and
   calls `active(bytes, alt_offsets)` once.
2. `active_block_ranges(bytes)` first runs
   `scan_word_element_ranges` for `p`, `tbl`, and `altChunk`, then calls
   `active(bytes, block_offsets)` a second time.
3. The writer's own `NsReader` walk captures the body prefix, every direct
   child, and the suffix. This third walk is outside the proposed fusion and
   must remain after both selections.

The candidate combines the first two lexical walks and removes one duplicate
walk. It cannot replace the final body capture, and it cannot turn the two MCE
selections into one selection.

## Limits and behavior that differ

The two lexical consumers do not have one common parser contract.

| Concern | `alt::scan` and its first `active` call | `scan_word_element_ranges` and `active_block_ranges` |
| --- | --- | --- |
| Raw source bytes | `MAX_XML_BYTES` = 32 MiB is checked before the alt walk | No byte check before its structural walk; `active` checks the 32 MiB source limit only after that walk |
| Element depth | `MAX_XML_DEPTH` = 256; `next_depth` is evaluated for both `Start` and `Empty` events, while only `Start` changes the persistent depth | `MAX_SCAN_DEPTH` = 128; persistent total depth and captured-element depth are checked for `Start`, while `Empty` does not increment either counter |
| Node count | No general node ceiling | `MAX_SCAN_NODES` = 1,000,000 for every `Start` or `Empty`; emitted range storage also uses the 1,000,000 semantic-value ceiling |
| Alt anchors | `MAX_CHUNKS` = 4,096 entries, including anchors nested inside a paragraph/table and anchors in inactive MCE branches | Captures only the outermost matching target; any `Start` while a target is captured suppresses nested `p`, `tbl`, and `altChunk` ranges |
| Namespace match | Only a bound transitional or Strict WordprocessingML namespace is an alt anchor namespace | Uses `is_fragment_word_namespace`: bound Word namespaces match, and an unbound/unknown first fragment prefix can make later same-prefix or unbound names match |
| XML grammar | Validates alt relationship attributes, duplicate valid relationship IDs, `matchSrc` values, alt child order, non-whitespace text, and foreign opaque subtrees; it does not add a general root/name-matching validator | Tracks byte spans and namespace-qualified local names; it does not validate alt metadata or body grammar, and it intentionally does not require a closed root at EOF when no target capture remains |
| Offset storage | Offsets are the sorted `BTreeMap` keys for every parsed anchor; inactive and nested anchors are included before MCE filtering | Ranges are emitted in scanner order for outer targets only; the separate `starts` vector is allocated after the walk |
| MCE marked input | `MAX_VISIBILITY_OFFSETS` = 1,000,000 and `MAX_MARKED_XML_BYTES` = 128 MiB; processing input/output are each 128 MiB, depth 256, namespace bindings 4,096, directive tokens 4,096, and choices per alternate 1,024 | The same `active` policy, but applied to a different offset vector and only after the second lexical walk |

The two `active` calls also have meaningful empty-input behavior. `scan` calls
`active` even when there are no anchors; `active` validates the source and then
returns before MCE namespace processing when its offset vector is empty.
`active_block_ranges` still supplies paragraph/table/anchor offsets, so its
second call can expose an MCE refusal in a source for which the first call was
an empty fast path.

## Counterexamples that constrain fusion

The following are source-level counterexamples, several of which are covered
by the 0720 diagnostic matrix.

* **No anchors does not make the second MCE call redundant.** A document with
  malformed `mc:AlternateContent` lacking `Choice`, or with an unsupported
  `mc:MustUnderstand` namespace, can make `scan` succeed and its empty
  `active` call succeed. The paragraph/table offset vector is nonempty, so the
  second `active` call returns the MCE refusal. A fused implementation must
  issue both calls, including the first empty call, in the same order.

* **Nested anchors make the vectors non-interchangeable.** In
  `<w:p><w:altChunk .../></w:p><w:altChunk .../>`, `alt::scan` parses both
  anchors. The range scanner emits only the outer `w:p` and the direct
  `w:altChunk`; its captured `w:p` suppresses the nested anchor range. The
  same distinction occurs for an anchor inside a table. A union, or using the
  range vector for both calls, changes both marker placement and the resulting
  body metadata.

* **Namespace rules differ on fragments.** For
  `<document><body><p/><altChunk/></body></document>` with no namespace
  declarations, the range scanner's fragment heuristic treats the unbound
  names as Word targets, while `alt::scan` recognizes no bound Word anchor.
  The current writer reaches its typed
  `active altChunk range lacks parsed anchor metadata` refusal. A combined
  reader must preserve this apparently inconsistent result rather than
  normalizing both predicates to one namespace test.

* **The range depth limit is earlier but lower priority.** With 127 nested
  wrappers inside the document/body, the range scanner reaches its depth-128
  refusal while the alt scanner's depth-256 policy can still finish. If a
  missing `r:id` anchor follows those wrappers, the current sequence reports
  the alt relationship error before the deferred range-depth error, because
  the complete alt scan is first. A one-reader implementation must remember
  the range error and continue the alt state machine.

* **Empty elements use a different depth rule.** At the boundary near depth
  256, `alt::scan` calls `next_depth` for an `Empty` event even though it does
  not retain that increment. The range scanner does not count the empty event.
  A shared depth counter changes whether the source is accepted and which
  error is returned.

* **Anchor count and structural node count are independent.** An alt anchor
  can exceed `MAX_CHUNKS` even while the range scanner would suppress its
  range inside an outer target. Conversely, a million non-target elements can
  trip `MAX_SCAN_NODES` even though `alt::scan` has no node limit. The first
  alt error must win if it occurs later in the source than a remembered range
  error.

* **Inactive branches are still parsed by both lexical passes.** The current
  alt walk validates anchor metadata in an unsupported `mc:Choice` before the
  first MCE selection, and the range walk records target offsets in both
  branches before its own selection. Filtering during the lexical walk would
  move malformed-input behavior and marker accounting.

* **Malformed XML must not gain a new scanner-owned validation contract.** The
  default `quick_xml` reader already checks matching start/end names; for
  example, the malformed-tail control reaches its reader error
  `ill-formed document: expected </w:p>, but </w:body> was found`. That check
  is shared input behavior, not a separate stack owned by either scanner.
  Neither scanner adds a general single-root or body-shape validator in all
  cases, and the later body-capture walk has its own body checks. A fused
  implementation must retain the reader's existing name checks without adding
  another grammar or single-root policy that changes typed refusals and error
  precedence. The single reader should consume events with the same
  `read_event` plus resolver behavior as `alt::scan`, then drive two
  independent state machines.

The BOM is part of this boundary: `DocumentBody::from_xml` removes one leading
BOM before either current scanner runs. All retained source offsets and range
lengths therefore address the BOM-adjusted slice, and the writer later
re-attaches the BOM in the preserved prefix.

## Viable architecture

One reader is viable only as a private, writer-bound dual-state lexical walk.
It is not viable as a generic replacement for `alt::scan`, as a change to the
shared namespace scanner's contract, or as an offset-union optimization.

The fused helper should retain two independent state tuples:

* The alt tuple must retain the alt depth, pending anchor, opaque foreign
  subtree, properties depth, relationship and `matchSrc` state, and a
  `BTreeMap<u32, Chunk>`. It must apply `MAX_CHUNKS` and every current alt
  grammar check exactly where `alt::scan` does.
* The range tuple must retain its separate total depth, captured outer target,
  capture depth, first fragment prefix, node count, range vector and first
  deferred range error. It must use the exact target suppression and
  `is_fragment_word_namespace` rules of `scan_word_element_ranges`.

For each event, execute the alt semantics first. An alt relationship, child,
text, depth, XML, or anchor-limit error must return immediately, because the
alt scan precedes the range walk today. Run the range semantics independently;
if its depth, node, offset, capture, or fallible-reservation operation fails,
retain the first error and disable only the range collection that can no
longer be trusted. Continue consuming events through the alt state machine so
that a later alt error can still take precedence.

At EOF, preserve the current stages rather than returning the collected values
directly:

1. Finish the alt state and call `active(bytes, alt_offsets)` even when the
   vector is empty. If this call fails, return that error before any range
   error or second call.
2. If a range error was remembered, return it now. This is the point where the
   existing `active_block_ranges` structural walk would have returned, before
   its MCE call.
3. Allocate the block `starts` vector with the existing exact reservation,
   call `active(bytes, starts)` as the second and separate MCE operation, and
   filter the collected ranges by its selected starts in their original order.
4. Return the chunk map and block ranges to the unchanged body-capture walk.

This preserves the observable precedence `alt lexical error -> alt MCE error ->
range lexical/error -> block MCE error`, while preserving two independent MCE
inputs. It also preserves the case where the range walk reports an error but a
later alt error is the result.

The implementation must use the exact per-call vectors:

* `alt_offsets`: every anchor start accepted by the alt parser, in the
  `BTreeMap` key order, including inactive and nested anchors before the first
  MCE filter.
* `block_offsets`: every outer `p`, `tbl`, or `altChunk` range start emitted by
  the range scanner, in event/source order, including inactive branches and
  fragment-namespace matches.

They must not be unioned, deduplicated together, sorted by a common marker
index, or replaced by the final selected offsets. `active_offsets` preserves
caller order and duplicates, and its marked-byte calculation depends on the
individual input count and decimal marker width. The second call must still
occur when the first call has zero offsets.

## Resource and allocation qualification

The diagnostic and any later differential test must report the two input
vectors and selected vectors, not only final block/chunk counts. It must also
exercise the exact limits at and beyond their boundaries: 32 MiB source,
128 MiB marked input, 1,000,000 offsets/nodes, 4,096 anchors, depths 128 and
256, MCE namespace/directive/choice ceilings, and u32 range conversion.

The raw byte preflight must remain before the fused walk. The current alt path
rejects an oversized source before doing either lexical pass; allowing the
range scanner to walk it first would change work and refusal order. The fused
helper must retain separate range and alt counters even when a source stays
below the raw cap.

Fallible reservations need special care. `active_block_ranges` reserves one
range at a time and then performs `starts.try_reserve_exact` after the walk.
The fused helper must preserve those resource labels and the stage of the
exact `starts` reservation. Interleaving chunk-map and range-vector growth
necessarily changes the timing of host allocator exhaustion; this is not a
proof of byte-for-byte allocation-failure equivalence. If the project requires
that failure schedule as part of parity, a one-reader implementation is not
qualified without an explicit shared allocation policy or retained two-pass
boundary. Semantic/resource limit parity and normal successful-path output
parity are achievable; global allocator-failure schedule parity is not
established by a lexical trace.

## Decision

Proceed with the current 0720 source-bound trace and differential controls.
It is useful for measuring the real two-pass denominator and checking that the
proposed boundary is reached in `DocumentBody::from_xml`. The trace alone
cannot authorize production fusion.

The candidate is qualified for a later A/B pilot only if an implementation
passes a differential matrix containing ordinary, Strict, BOM, MCE active and
inactive branches, nested anchors in paragraphs and tables, fragment and
foreign namespaces, marker-like text, malformed tails, depth boundaries,
node/anchor/offset limits, typed relationship errors, and allocation/resource
refusals. The pilot must show two MCE calls with the exact baseline vectors and
order, unchanged body-capture output and source preservation, and fresh native
edit/lifecycle/allocator evidence. Until then, no production fusion or speedup
claim is justified.

## Trace instrumentation audit

The exact-source recipe and retained trace were audited after restoration. The
`DocumentBody::from_xml` scope is installed after the BOM split and before
`scan(bytes)`, so it covers the alt scan, the range pass, both `active` calls,
and the existing body-capture walk. Standalone oracle calls to `alt::active`
have no document scope and are not mixed into the body records. Within that
scope the source call order is preserved: the first active record belongs to
`alt::scan`, and the second (when reached) belongs to `active_block_ranges`.
The wrappers around validation and the final MCE call return the original
`Result` unchanged; the event and metadata hooks only observe state. Baseline,
trace, and repeat public JSON are byte-identical, and the two trace stderr
files are byte-identical.

The alt event hook runs after a successful `read_event` and before alt
semantic handling, so a reader failure is not counted while a semantic error
on an already-read event is counted. The range hook runs after
`read_resolved_event`, node and total-depth checks, fragment-prefix discovery,
and event classification, but before capture-depth, range-emission, and
callback-reservation handling. Consequently a node or total-depth refusal
does not count its rejecting event; a later capture or emission refusal can
include its event. EOF is counted on a complete successful pass. The report
should keep that distinction when describing refusal counts.

The records preserve exact active input vectors and debug representations of
results, but the analyzer independently cross-checks only the block input
vector against raw ranges; the alt raw/selected maps remain debug strings that
are hashed rather than structurally compared with the first active record.
That is adequate for this source-bound diagnostic when the retained trace is
reviewed, but it is not a standalone machine proof of the alt vector/map
relationship. No instrumentation defect changes the observed logic or public
results; OOM behavior of the diagnostic's own allocations remains outside
this qualification.
