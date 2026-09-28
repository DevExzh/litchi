# 0810 source review: direct Reader and namespace resolver transport

This is a read-only review of the 0810 candidate at base commit
`3677e31be5c9d5582a1f6d531ebb4d54db5a0acc`. The base source is archived as
`candidate/before/codec.rs` with SHA-256
`8485d43c99f19bda9b3510c8323aa6bc2b117fb7df262372ce82868a5397239c`. The
current worktree source is the candidate after leg, whose archived SHA-256 is
`466e504588843a0b2538fc25ce3378b6fb9bf13aa9b0d74efebcb73ca0ebab5e`. The
candidate manifest and the packet's design and protocol reviews identify only
`crates/litchi-pptx/src/notes/codec.rs` as the production path.

No Cargo, rustfmt, test, workload, or profiling command was run for this
review. This document makes no performance or adoption claim.

## Transport and state-machine equivalence

The workspace resolves this crate to quick-xml 0.41.0. Its local
`NsReader::read_event_impl` performs three steps in order: consume a pending
namespace pop, read an event from its underlying `Reader`, and process the
event. Its `process_event` pushes `NamespaceResolver` for `Start` and `Empty`,
marks `Empty` and `End` for a pop on the next read, and passes parser errors
and other events through.

The candidate spells out that same sequence in `scan_processed_xml`:

* `pending_pop` is consumed before every `Reader::read_event()` call, including
  the call that returns `Eof` or a parser error;
* `Start` and `Empty` call `NamespaceResolver::push` before node/depth
  accounting, root classification, or `inspect_element`;
* `Empty` and `End` set `pending_pop` before their existing validation branches;
* an `End` scope therefore remains active while the returned end event is
  handled and is removed before the following event; and
* namespace failures are converted through `quick_xml::Error::from` and then
  the existing `xml_error` adapter.

This matches the local quick-xml implementation, including the case where an
end event is followed by a parser error. A failed `push` returns before the
candidate's node, depth, or attribute checks, just as it did when the push was
inside `NsReader`. `NamespaceResolver::default()` retains quick-xml's
per-element 256-declaration bound and reserved-prefix checks, so declaration
errors retain their priority over the scanner's later depth/node checks.

The inspector now receives `&NamespaceResolver` directly and calls the same
`resolve_element` and `resolve_attribute` methods. Root namespace and local
name checks, checked-attribute parsing, relationship collection, unescaping,
attribute counters, raw and processed byte limits, depth and node limits, and
the existing refusal branches remain in their previous order. The test-only
counter helper still uses `NsReader` and passes its resolver to the inspector;
it does not exercise a second production transport.

## Borrowed data and ownership

Both the old `NsReader<&[u8]>::read_event()` path and the candidate's
`Reader<&[u8]>::read_event()` path return events borrowing the already
processed input slice. The candidate does not introduce a scratch buffer or
extend an event lifetime. `NamespaceResolver` owns its bounded namespace
binding storage, as it does inside `NsReader`; the inspector uses resolved
names immediately. The scanner continues to own only the relationship and ID
values it reports with `to_owned()`. No public signature, dependency, unsafe
code, or ambient state changes in this candidate.

## Oracle preservation and test coverage

The retained `buffered_scan_oracle` remains the independent `NsReader` plus
`read_event_into` implementation. The retained `inspect_element_oracle` also
remains unchanged. Read-only extraction of both functions from the archived
before source and the current after source produced identical bytes and
identical SHA-256 values (`397863f9fef32a35e20ac2c1baf2874dbf4d6891c0de4c2b9f39bde0ceb15b38`
for `buffered_scan_oracle` and
`a563325f899a53544b344b78d5946383ba5e871d56d2475cef4b11054cd1c7af` for
`inspect_element_oracle`).

The existing differential helper compares direct and buffered outcomes across
all supported root names and conformance modes, independent raw and processed
ceilings, and both accepted and refused inputs. Its handcrafted and mutation
inputs retain default-root resolution, prefixed-root resolution, Empty and
Start elements, malformed and duplicate attributes, unknown prefixes,
invalid UTF-8, DTD/PI/CDATA refusal, missing roots, mismatched ends, and depth
and node ceilings.

The three focused tests add useful ordering cases: nested Start/End and Empty
scope transitions with prefixed and default rebindings; a reserved namespace
error at the depth boundary; and reserved-prefix errors plus the
declaration-cap boundary. Each compares the direct result and refusal text to
the retained oracle.

One non-blocking coverage note is actionable: the nested-scope test declares
default namespace rebindings, but its explicit successful result depends on a
prefixed `r:id` after scope restoration. The current scanner does not report a
non-root element's default namespace, so this test does not independently
observe a default binding on a child. Existing default-root cases still cover
root resolution and the oracle comparison covers the current contract. If
future scanner behavior depends on child default namespaces, add a direct
assertion or fixture whose result depends on that binding.

## Review result

**PASS for static source equivalence; no namespace, error-ordering, limit-ordering,
borrowed-semantics, or oracle-integrity blocker found.** The focused coverage
note above is a test-strength improvement, not a source regression. Runtime
build/test gates, matched captures, independent replay, and the final adoption
decision remain root-owned.
