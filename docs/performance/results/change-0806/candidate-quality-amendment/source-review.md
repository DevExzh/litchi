# 0806 candidate quality amendment source review

This is a static source review of the five-file quality amendment archived beside
this document. It addresses the warnings-denied quality failure retained in
`../quality-0/01.log`: the exact 0805 candidate left the private
`BytesStartExt::unchecked_attributes` method unused in `litchi-sign`. No Cargo,
build, test, capture, or profiling command was run for this review.

## Scope and custody

The parent candidate is bound to the original candidate manifest, candidate
patch, application record, and baseline source census by the artifact hashes in
`manifest.json`. The amendment has five production paths:

| Archive key | Production path |
| --- | --- |
| `litchi-ole-common-xml_attributes.rs` | `crates/litchi-ole-common/src/xml_attributes.rs` |
| `litchi-opc-xml_attributes.rs` | `crates/litchi-opc/src/xml_attributes.rs` |
| `litchi-sign-xml_attributes.rs` | `crates/litchi-sign/src/xml_attributes.rs` |
| `litchi-xldm-xml_attributes.rs` | `crates/litchi-xldm/src/xml_attributes.rs` |
| `xml-minifier-xml_attributes.rs` | `crates/xml-minifier/src/xml_attributes.rs` |

The `before/` bytes are the exact original 0806 candidate `after/` bytes for
these five paths. The amendment does not include or alter
`crates/litchi-opc/src/xml_attributes/tests.rs`; its original candidate bytes
are bound in the manifest by SHA-256
`e26c97c83e0de48ddf9e66f8b33dec8f62996fd70cddf357c802c0e7ca3ba2bf`.

## Equivalence review

Each original `CheckedAttributes::new` constructor performed the following
two operations:

```rust
let mut attributes = tag.attributes();
attributes.with_checks(false);
```

The amendment replaces those operations with:

```rust
let attributes = tag.unchecked_attributes();
```

`BytesStartExt::unchecked_attributes` is the existing local helper in every
one of the five files. Its body still performs `self.attributes()` followed by
`with_checks(false)` and returns that same `Attributes` value. The call is made
on the same `&'a BytesStart<'a>` reference, so the iterator's item lifetime,
raw tag bytes, unchecked duplicate policy, lexical-error behavior, and
iteration order remain the same. The `Self` initializer retains the same
`tag`, `attributes`, and `Phase::First` fields, and no state-machine branch,
comparison, allocation, error mapping, or public/private visibility changes.

The constructor-local `#[allow(clippy::disallowed_methods)]` is removed in all
five amended files because the constructor no longer directly invokes the
quick-xml method. The helper's existing allow remains exactly where the one
direct `with_checks(false)` call is intentionally encapsulated. This keeps the
workspace lint policy explicit without suppressing an unused constructor call.

The OPC helper remains public and the other four helper visibility levels are
unchanged. All other lines in the five files are byte-identical to their
parent-candidate `before/` files outside this constructor replacement. The
production-path patch records exactly those five replacements.

## Review conclusion and required gates

The static review finds the amendment semantically equivalent to the parent
candidate at the XML-attribute iterator boundary and finds no source-level
correctness or API blocker. This conclusion is limited to source equivalence;
it is not a performance result and does not waive execution gates.

The root coordinator must obtain an independent review of this archive, run a
fresh protected native micro-preflight against production-before, apply the
amendment only after that preflight passes, and rerun the production quality
gates and the remaining 0806 workflow evidence. Any adoption decision must be
bound to the amended source census and those fresh results.
