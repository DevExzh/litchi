# 0815 source review: borrowed event-arm bindings

This review covers the archive-only candidate for commit
`55bb2ead3498043dd22b53555507b24402486248`. It does not apply the patch or
authorize production retention. The candidate is limited to
`crates/litchi-pptx/src/notes/codec.rs` and is intended to remove the two
remaining `BytesStart` arm-local copy sequences identified in the 0814 native
attribution packet.

The archived identities are:

| Artifact | Bytes | SHA-256 |
| --- | ---: | --- |
| `candidate/before/codec.rs` | 52,443 | `007c2429b44027c37bb6f648f8a2841bfce7ef74b164aa70ed08fd06248b4f0d` |
| `candidate/after/codec.rs` | 52,447 | `ab3f12b1831a13358e9a8d5c31c40abe26a1eafc693c3cabc0fa731653afdf83` |
| `candidate/candidate.patch` | 1,462 | `b7cf56cc674a92b76b2e40fdeeebcde06feaf746d22bc8385280acc5ba2519ce` |

The before snapshot matches the current production source archive. The
candidate changes exactly six source spellings in `scan_processed_xml`:

```rust
Ok(Event::Start(element)) => {
    // ... resolver.push(&element) ... inspect_element(..., &element, ...)
}
```

becomes the equivalent borrowed binding and calls:

```rust
Ok(Event::Start(ref element)) => {
    // ... resolver.push(element) ... inspect_element(..., element, ...)
}
```

The same two substitutions occur in the `Empty` arm. There is one `ref`
binding change per arm, two `resolver.push` argument changes, and two
`inspect_element` argument changes. The candidate keeps the existing direct
`Reader::read_event()` result match from the retained 0813 source; it does not
rewrite event dispatch, add a parser, or alter an error conversion.

## Lifetime and drop review

`Reader::from_reader(processed)` is a `Reader<&[u8]>`, and
`Reader::read_event()` returns `Result<Event<'a>>` whose `Start` and `Empty`
payloads borrow the processed slice. With the original binding, the
`BytesStart` value is moved into the arm and then borrowed for each call. With
`ref`, the match binds a reference to the payload already held by the
temporary `Event`; both calls still receive `&BytesStart` and the payload is
destroyed when the arm finishes. No reference escapes the arm, and no event is
stored in `resolver`, `scan`, or any returned value.

The quick-xml 0.41 `NamespaceResolver::push(&BytesStart)` implementation
iterates the element's attributes and copies namespace prefixes and URIs into
its own buffer. It does not retain the `BytesStart` reference. The
`inspect_element` signature remains `element: &BytesStart<'_>`; it resolves
names and attributes during the call and owns only relationship values that it
appends to `XmlScan`. Thus the borrowed binding does not change the resolver's
or inspector's lifetime requirements.

`pending_pop` is set in the same place in the `Empty` arm and is consumed at
the top of the next loop iteration. The event payload still drops before that
next iteration, exactly as in the owned binding. `Start` and `End` depth,
namespace, and root state transitions are unchanged.

## Error and ordering review

The candidate leaves the following order and typed paths unchanged:

1. `resolver.push` runs before node/depth checks in both success arms.
2. `NamespaceError` is converted through `quick_xml::Error::from` and then
   `xml_error` exactly as before.
3. Node, depth, root, checked-attribute, UTF-8, unescape, and attribute-byte
   limits remain in their existing locations.
4. `Event::End`, DTD/PI, CDATA, EOF, ignored events, and the `Err(error)` arm
   are byte-identical.

In particular, the candidate does not move `pending_pop`, change when a
namespace scope is removed, or replace the parser's error with a synthesized
error. It has no public API, dependency, unsafe-code, ambient-state, or
resource-budget change.

## Oracle and coverage review

The complete `#[cfg(test)]` suffix is byte-identical in both snapshots:
33,693 bytes with SHA-256
`78e16e72fbb50e24dc682f4420a593cce5f1d613bd8dbcbf8650ddd75579bb40`.
Therefore the buffered scanner oracle, `inspect_element_oracle`, handcrafted
boundary/refusal cases, mutated XML differential cases, and PPTX corpus
differential cases are unchanged. The candidate adds no test-only behavior;
the existing tests exercise both `Start` and `Empty` paths, valid attributes,
namespace declarations, duplicate attributes, malformed values, UTF-8 and
resource-limit refusals.

This is a static source decision only. Root must still run the frozen format,
production, probe, qualification, ordinary/profile code-generation, public
workflow, allocation, profile, independent-reader, and final custody gates.
The code-generation gate must show that the arm-local copy pattern is reduced
in the candidate assembly before any fresh timing lane is allowed to proceed.
No timing, allocation, instruction-count, RSS, or adoption claim follows from
this review.

**Review disposition: source candidate accepted for bounded execution.**
