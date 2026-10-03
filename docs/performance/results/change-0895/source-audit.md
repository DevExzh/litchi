# 0895 MCE and selected-route allocation source audit

This is a read-only source attribution for the 0894 allocator observation and
the 0895 heaptrack preflight. It records bounded hypotheses for a later
diagnostic; it does not establish causality or make a current production
performance claim. No implementation source, workload, or build was changed
for this audit.

## Identity and binding

The restored source under review is HEAD
`af9f2346a3fd65d762d82ea94dbb49a082c4bc82`. The relevant current-source
hashes are:

| source | SHA-256 |
| --- | --- |
| `crates/litchi-ooxml-common/src/mce/stream.rs` | `1e194cae9ec2f2b559543656011c198fb67c84f52dc47894dba44f18e4f9710f` |
| `crates/litchi-ooxml-common/src/mce/scope.rs` | `579513e80008bd4dddea4f8e0f603a963e0646fd6e24db4f558459c5b48cbfc2` |
| `crates/litchi-ooxml-common/src/mce/model.rs` | `44169877f7c689b9266583610f6dfb84138b2a1039c7edc22b5d82ce75eed070` |
| `crates/litchi-ooxml-common/src/xml_name/model.rs` | `c877419718808eb78e434424448be174a51c79d71c21fedbd2d83d234469576c` |
| `crates/litchi-ooxml-common/src/xml_name/codec.rs` | `ec6ea83c47bc9bdbacf7bbc1f682aadaa1fb72e97e8173a49ac39f1320459a6a` |
| `crates/litchi-xlsx/src/raw/worksheet/x14ac.rs` | `9f766274df5cc0a6a80291644ecbc4f9d8afc6dac556af86bd21fbde66d148` |
| `crates/litchi-xlsx/src/raw/worksheet/selected.rs` | `f83cefd73e2cd5ab8f741fbd3e9eab9c2ee3793a211f5e2be63cfd7e39af1eb2` |

The 0894 reduction reports allocator calls changing from `25,933` to
`126,171` (ratio `4.8652681911`), with allocated bytes changing from
`37,618,900` to `47,632,498` (ratio `1.2661851572`). Those are paired global
allocator counters, not per-symbol counts. The 0895 heaptrack sample is tied
to the frozen 0894 candidate binary
(`candidate-native`, SHA-256
`1907d95f2248c1b548731cede3b0692e6daf357d26a6d305583532d2a33cb548`), not
to the restored production HEAD. Its report, stack export, and compressed
trace hashes are respectively
`98999120eb4c3404b2a6bd67c454fb340753319a3687547cf7152ac0d59219d1`,
`61e98a7c4c9d13a8026a516498cbc33d326b502824587e448c09e3eb66ce34c7`, and
`0bd62ba79f56bb2f4383565984f150b44c33cadc9df15f179d5e171c361fce34`.
The line references below are therefore current-source bindings for resolved
symbols, while the allocation evidence remains a candidate-binary diagnostic.
The 0894 candidate quality diff is selected-route-only
(`docs/performance/results/change-0894/attempts/quality-v9/git-diff-after.txt:1`);
the MCE stream, scope, model, and XML-name files listed above are byte-identical
for the candidate and restored source bindings. The selected-route source is
therefore the only current-versus-candidate binding that needs separate
qualification here.

## Resolved allocation groups

The preflight stack export contains overlapping call-path groups. Counts in
this table must not be added together or treated as a causal decomposition.

| heaptrack nearest symbol | current source site | ownership interpretation |
| --- | --- | --- |
| `clone_bounded_bytes` / `clone_bounded_name_part` from `Processor::parse_element` / `expand` | `stream.rs:2710-2811`, especially `2758-2779`, `2790-2809`, and `3322-3381` | Per-element and per-attribute owned lexical/name pieces. `QualifiedName::from_parts` allocates its boxed lexical name at `xml_name/model.rs:56-62`, through `xml_name/codec.rs:8-25`. |
| `clone_bounded_string` from `raw_start_commit` and `close_start` | `stream.rs:2508-2557` and `2198-2245` (name copies at `2544` and `2233-2243`) | Raw and semantic frame names retained until the corresponding frame is closed or recovery clears the stack. These are parser state, though usually short-lived on shallow input. |
| `clone_shared_name` from `SemanticEvent::from_element`, raw snapshots, and `borrow_raw_attribute` | `stream.rs:2876-2898`, `2927-2956`, `3516-3526` | The namespace URI is shared, but the local name is copied. The semantic and raw callback views can therefore account for separate local-name allocations. |
| attribute `clone_bounded_bytes` / `clone_bounded_text` | `stream.rs:2742-2779` | Raw key, raw value, and normalized/decoded value are separate bounded owned fields in the parsed snapshot. They are temporary unless a directive, frame, or callback-owned event retains them. |
| `try_reserve_set` from `validate_duplicate_attributes` | `stream.rs:2859-2873`, reserve helper near `3576` | A per-element `HashSet<(UriId, &str)>`; the stack report marks the observed groups temporary. This is a validation guard, not selected output. |
| `reserve_vec` from `semantic_attrs` | `stream.rs:2959-3012`, reserve helper near `3551` | A semantic attribute result vector built for an event. It is temporary when the observer consumes the event and is cleared with the event buffer. |
| `QualifiedName::from_parts` / `xml_name::codec::parse` | `stream.rs:3358`, `xml_name/codec.rs:8-25`, `xml_name/model.rs:56-62` | The QName parser allocates even though `expand` already has validated prefix/local slices. This is the clearest allocation-free parsing hypothesis, subject to exact error and grammar preservation. |
| `event_buffer.try_reserve_exact` | `stream.rs:976-987` | A retained stream buffer sized by `max_event_bytes`; heaptrack attributes about `8.39 MiB` of peak allocation over four calls here. This is a large retained allocation, but not an explanation for the 126,171 call count by itself. |

## Selected and x14ac ownership

The x14ac observer is a guard and metadata capture path, not evidence that all
MCE allocations become selected records. `capture_stream_with_stoppable_active`
installs the raw and active callbacks at `x14ac.rs:195-227`. The raw callback
only constructs a `PendingCandidate` when it sees the exact `x14ac:dyDescent`
marker (`x14ac.rs:346-391`); its `Box<[u8]>`, expanded name, and `Box<str>`
copies are made by `copy_bytes`, `copy_name`, and `copy_str`
(`x14ac.rs:311-329, 494-537`). `active` clears or consumes that pending state
and pops its stack (`x14ac.rs:394-423`). Candidate matching checks sequence,
kind, qualified name, and expanded name (`x14ac.rs:494-508`), so those copies
are guards for an exact cross-callback match. `Values.rows` is retained only
for `RowMode::Retain` (`x14ac.rs:42-46`); the selected scanner passes
`RowMode::ValidateOnly` (`selected.rs:282-309`), so the preflight route does
not justify attributing a retained row map to selected output.

The selected scanner's persistent ownership is separate: its `selected` and
`merges` vectors, dependency state, and stack are declared at
`selected.rs:1179-1200`; final selected records are published by `finish` at
`1227-1280` and `retain_selected` at `1954-1963`. Cell text/formula scratch
is built at `1655-1773`, parsed into selected values at `2117-2149`, and
recycled for unselected numeric values at `1896-1905`. Those records, formulas,
values, and required merge/dependency results are intentional output or
qualification state. They must not be counted as removable per-event
temporaries. The scanner stops on its first ineligible event
(`selected.rs:1282-1329`); the current source route then falls back through
`SourceWorksheet::stream_selection` (`source.rs:1034-1068`). The rejected 0894
metadata overlay is not the current production route.

## Bounded future hypotheses

These are source hypotheses for a later isolated change, not implementation
instructions for this diagnostic.

1. **Borrow-only QName expansion (highest signal).** Keep
   `is_qualified_name` validation allocation-free (`xml_name/validation.rs:25-32`)
   and resolve prefix/local slices directly in `stream.rs:3358-3381`, avoiding
   `QualifiedName::try_from` and `from_parts`. Preserve invalid-QName errors,
   prefix/default-namespace behavior, local-name limits, and every downstream
   comparison. Exhaustive QName and MCE differential coverage is required.
2. **Bounded duplicate-name validation.** Replace the ordinary-element
   per-event `HashSet` only with a small fixed/inlined set plus a spill path.
   Preserve the exact `(UriId, local_name)` key, namespace declaration rules,
   `max_attributes`, duplicate typed errors, and reserve/error behavior. A
   linear scan is safe only under the existing finite attribute bound and must
   be measured separately from name parsing.
3. **Borrow ordinary event fields until the callback boundary.** Avoid copying
   raw key/value/decoded value and expanded local names when no directive,
   frame, or retained result needs ownership. Preserve XML normalization and
   decode semantics, raw-before-active ordering, callback-visible lifetime, and
   `event_buffer` clearing after callbacks (`stream.rs:1159-1169`). The current
   `RawElement` and `SemanticElement` structs carry owning vectors and name
   copies while their `Cow` fields remain lifetime-bound; the public API lets a
   consumer retain or clone those owned pieces. An internal observer/view path
   is therefore likely needed rather than changing the public lifetime
   contract.
4. **Internal selected/x14ac view path.** Let the selected scanner and x14ac
   callbacks inspect bounded borrowed element/attribute views instead of
   materializing ordinary `RawElement`/`SemanticElement` vectors. Keep the
   public stream API unchanged, and preserve hidden MCE branch selection,
   callback error precedence, recovery, exact x14ac sequence matching, and the
   scanner's first-refusal/no-callback contract. Do not remove selected output
   or dependency retention.
5. **Share frame-name storage only if stack independence remains exact.** The
   duplicate semantic/raw frame name copies at `close_start` and
   `raw_start_commit` could share an immutable handle, but end-name checks,
   recovery clearing, namespace context, and frame lifetimes must remain
   independent. This is lower confidence than eliminating QName parser boxes.
6. **Event-buffer growth policy.** Investigate lazy or capped reserve around
   `stream.rs:976-987` while retaining `max_event_bytes`, typed allocation
   errors, reader behavior, callback timing, and peak/RSS limits. The peak
   allocation does not establish that this is the source of the allocator-call
   ratio.
7. **x14ac copy specialization.** Revisit pending-candidate copies only after
   a sequence-sensitive differential proves that the raw marker and active
   element have identical identity and ordering. Preserve marker aliases,
   duplicate detection, invalid-descent handling, and `ValidateOnly` row-mode
   semantics. The current evidence does not show these guard copies dominate.

Namespace declaration copies in `Namespaces::with_local` (`stream.rs:1605-1652`)
and `Scope::push` (`scope.rs:362-438`) are a lower-priority follow-up. Any
sharing must preserve URI identity, prefix ordering, context-byte accounting,
and release/recovery behavior. No hypothesis permits weakening stream limits,
turning a typed stream error into `NotEligible`, swallowing callback errors,
or replacing authoritative fallback with a partially validated selected result.

The formal 0895 four-child profile should isolate these hypotheses with
symbol-filtered captures or targeted instrumentation and compare the same
candidate/fallback route. The current one-sample heaptrack trace and the 0894
global allocator pair are sufficient to prioritize source sites, not to accept
an optimization.
