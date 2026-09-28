# 0812 scanner source review

This is a bounded, read-only source review of the retained 0811 production
path. It does not change production code and no Cargo, workload, or profiler
command was run for this review. The review is attached to the current
`crates/litchi-pptx/src/notes/codec.rs` scanner after 0810:
`scan_processed_xml` reads borrowed events from `Reader<&[u8]>`, maintains a
public `NamespaceResolver`, and passes the borrowed `BytesStart` to the
existing inspector. The current worktree's three unrelated files remain
outside this review.

## What the source proves

The loop at `scan_processed_xml` has one event transport boundary:

```rust
match reader.read_event().map_err(xml_error)? {
    Event::Start(element) => { ... }
    Event::Empty(element) => { ... }
    Event::End(_) => { ... }
    Event::DocType(_) | Event::PI(_) => { ... }
    Event::CData(_) => { ... }
    Event::Eof => break,
    _ => {},
}
```

`Reader::read_event` returns `Result<Event<'_>, quick_xml::Error>`. The
`map_err` changes only the error type; on the success path it must carry the
whole `Event` through `Result` and the `?` residual boundary before the match
can inspect its discriminant. The `Event` enum includes payload-bearing
variants, and `BytesStart` itself contains a borrowed `Cow` payload, a name
length, and a decoder. A release compiler can elide some or all moves, but the
source-level construction is a plausible place for aggregate event payload
moves or stack/register shuffling. It is the smallest source seam that is
specific to the direct scanner and does not repeat the rejected 0806 checked
attribute iterator idea.

The explicit equivalent worth measuring is a direct `Result` match that
handles the error arm beside the existing event arms:

```rust
match reader.read_event() {
    Ok(Event::Start(element)) => { /* existing Start arm */ }
    Ok(Event::Empty(element)) => { /* existing Empty arm */ }
    Ok(Event::End(_)) => { /* existing End arm */ }
    Ok(Event::DocType(_)) | Ok(Event::PI(_)) => { /* existing refusal */ }
    Ok(Event::CData(_)) => { /* existing refusal */ }
    Ok(Event::Eof) => break,
    Ok(_) => {},
    Err(error) => return Err(xml_error(error)),
}
```

If the compiler needs a local to keep the existing arms readable, the next
smallest spelling is equivalent but a weaker movement probe:

```rust
let event = match reader.read_event() {
    Ok(event) => event,
    Err(error) => return Err(xml_error(error)),
};
match event {
    // the existing arms, unchanged
}
```

This is only a code-generation experiment. `Result::map_err` is inline in
`core`, so the compiler may already produce the same machine code. The
experiment is useful only if exact release disassembly or native profiles show
a difference; a source rewrite by itself is not evidence of fewer moves or
faster capture. The explicit branch must not be retained if it changes
published output-byte identity, semantic results, or refusal identity/ordering
without a separately justified reason.

## Existing movement evidence and its limit

The archived 0807 `event-assembly.txt` shows the old
`NsReader::process_event` wrapper copying the returned `Event` payload between
stack/return locations with `movups` instructions. Its retained sampled
offsets concentrate on those copies, including the return-side payload stores.
That evidence explains why 0810 removed the scanner's direct
`scan_processed_xml -> NsReader::process_event` edge. It does not prove that
0811's inlined loop still emits the same copies: the 0810 profile preserves
`Reader::read_event_impl` calls and the current 0811 native profile reports
`scan_processed_xml` as the self leaf, but a self-leaf name is a symbol bucket,
not a line-level cost attribution.

The current source also makes two other classes of work visible inside that
same bucket:

* pending resolver pops happen at the top of every iteration;
* `Start` and `Empty` arms perform resolver pushes, limit checks, and the
  complete `inspect_element` call, while `End` performs scope timing and depth
  updates.

Therefore the scanner self-leaf cannot be interpreted as event movement until
the exact 0811 frame-pointer binary maps the sampled offset to an instruction
sequence. `Reader::read_event_impl`, `ReaderState::emit_start`/`emit_end`,
`NamespaceResolver::push`, checked attributes, duplicate checking, prefix
resolution, UTF-8 validation, and unescape remain separate sampled symbols in
the retained native evidence. Their presence is consistent with the scanner
being a loop/control bucket; it is not proof that any one child dominates its
self samples.

## Smallest bounded experiment

If the exact disassembly places the sampled scanner offset at a success-path
`Result`/`Event` move around `map_err(...)?`, qualify one candidate that changes
only that expression to the explicit `match` above. Keep all of the following
identical:

* `Reader<&[u8]>`, parser configuration, and every event arm;
* resolver push/pop timing, including the pending pop before the next read and
  the `Empty`/`End` scope behavior;
* node, depth, attribute, and attribute-byte limits and their order;
* `inspect_element`, checked attributes, duplicate detection, UTF-8 checks,
  unescape, namespace resolution, relationship collection, and ownership of
  reported values;
* parser error conversion and the existing generic XML/invalid/limit error
  precedence; and
* the buffered scanner oracle, all focused namespace/refusal cases, output
  bytes, semantic readback, cancellation, and resource measurements.

The qualification must first compare the old and candidate scanners directly
on the full malformed/refusal corpus. Only after byte-identical values and
exact error identity/ordering pass should it use the existing release quality
gates and fresh counterbalanced native capture, with allocation/resource
checks and owner-scoped stacks. A lower count of `mov` instructions or a
smaller symbol is a mechanism observation; retention still requires an
end-to-end workflow benefit under the established gate. The candidate should
be abandoned if the disassembly shows no event-result movement or if the
explicit branch compiles identically.

## Boundaries

This review does not propose a custom event type, a second XML parser, unsafe
code, a quick-XML fork, or a resolver/attribute traversal rewrite. It does not
revive 0806's combined `xml_attributes`/iterator candidate: that candidate
already failed the public workflow timing gate despite allocation reductions.
It also does not infer ordinary-build cost from the 0811 frame-pointer lane;
the large `fp/profile` perturbation is material. The event-result hypothesis
must remain conditional on the exact current binary disassembly and a fresh
end-to-end experiment.
