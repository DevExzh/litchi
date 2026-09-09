# litchi-ole-common

`litchi-ole-common` is the host-neutral layer above [`litchi-cfb`](../litchi-cfb)
for bounded legacy Office OLE data. It provides shared CFB object ownership,
OLEDS stream codecs, OLE1 compatibility handling, and OLE Property Set models
used by the DOC, PPT, and XLS format crates.

The public module map is in [`src/lib.rs`](src/lib.rs). The APIs below keep
payloads inert: callers can inspect and edit declared metadata, but the crate
does not resolve links, open native documents, render presentations, activate
classes, or execute embedded content.

## Read, write, and limits

| Surface | Read | Write | Limits and publication boundary |
| --- | --- | --- | --- |
| Target-selected CFB objects ([`object`](src/object/mod.rs)) | [`Editor::open`](src/object/editor.rs), [`Snapshot`](src/object/snapshot.rs), and [`discover`](src/object/discovery.rs) capture host-selected storages, directory metadata, and opaque descendant streams. | [`Editor`](src/object/editor.rs) supports whole-object replacement, opaque stream add/replace/remove, storage add/remove, and typed link, presentation, and native-stream edits. Commits expose a [`Snapshot`](src/object/snapshot.rs) and exact-source [`Patch`](src/object/patch.rs). | [`object::Limits`](src/object/model.rs) bounds target count, storage depth, stream counts, individual stream/object size, and aggregate captured bytes. Candidates are rendered, reopened, and rediscovered before publication; a true no-op keeps the original package bytes. |
| OLEDS OLE2 streams ([`ole_streams`](src/ole_streams.rs)) | Parses `\x01Ole`, `\x02OlePres###`, TOC entries, and `\x01Ole10Native` into checked stream values and shared snapshots. Presentation and native payload bytes remain opaque. | [`PresentationTransaction`](src/ole_streams.rs) and [`NativeTransaction`](src/ole_streams.rs) edit checked metadata and opaque data, then return source-checked reversible patches. | [`ole_streams::Limits`](src/ole_streams.rs) bounds complete stream bytes, opaque data bytes, and TOC entries. Presentation names and indices are checked independently of TOC size. |
| OLE1 objects and presentations ([`ole1`](src/ole1.rs)) | [`ObjectRef`](src/ole1.rs) provides borrowed reads; [`ObjectSnapshot`](src/ole1.rs) owns a checked source allocation. `Compatibility::Strict` is the default. | Canonical embedded and linked objects can be created with the five normative presentation forms. [`ole1::Edit`](src/ole1.rs) changes supported fields and returns an exact-source [`Patch`](src/ole1.rs) with an inverse. | [`ole1::Limits`](src/ole1.rs) bounds the complete object, native and presentation payloads, strings, and registered format names. `EmptyPresentationHeader` and `MissingPresentation` are explicit read/preserve-only compatibility profiles; they do not enable noncanonical authoring. |
| Property Sets ([`property_set`](src/property_set/mod.rs)) | [`Stream::parse`](src/property_set/codec/binary/mod.rs), [`PropertySetReader`](src/property_set/codec/package.rs), and [`SharedPropertySetReader`](src/property_set/codec/package.rs) expose checked sections, typed `Value` variants, arrays/vectors, dictionaries, unknown values, and typed Summary/Document Summary projections. [`NonSimpleSnapshot`](src/property_set/non_simple/mod.rs) lazily inspects a source-backed non-simple Property Set CFB storage and resolves its typed indirect stream/storage values on demand. | [`Stream::to_bytes`](src/property_set/codec/binary/mod.rs), [`property_set::Editor`](src/property_set/codec/editor.rs), the typed Summary Information, Document Summary Information, and user-defined hyperlink editors, and [`NonSimpleEditor`](src/property_set/non_simple/mod.rs) support validated section and bounded physical-element changes. | The generic codec applies section-local property, array/vector, text, and allocation ceilings; generic parsing leaves non-simple indirect values opaque. [`NonSimpleLimits`](src/property_set/non_simple/model.rs) bounds CFB metadata, streams, contents, indirect references, and rendered output. Changed non-simple candidates are reopened and closure-checked; exact no-ops and inverse patches retain source bytes. |
| Binding names ([source](src/property_set/binding/mod.rs)) | [`Binding::from_name`](src/property_set/binding/model.rs) and `Binding::from_format_identifier` recognize standard and GUID-derived OLEPS bindings. | `Binding::name` emits the canonical CFB name; `BindingName` retains its validated fixed-size representation. | Names obey the standard named or 27-byte GUID-derived grammar and bit order. `PROPERTY_BAG_FMTID` is an identifier for the generic binding codec; it does not provide a non-simple PropertyBag storage editor. |
| Alternate-stream metadata ([source](src/property_set/alternate_stream/mod.rs)) | [`AlternateStreamControl`](src/property_set/alternate_stream/model.rs) parses the exact 8-byte or 24-byte control packet. [`NonSimpleAlternateStreamName`](src/property_set/alternate_stream/model.rs) parses the `Docf_` filesystem selector and its binding suffix. | Fresh control packets and canonical selectors can be created; parsed control packets retain the ignored source reserved word through typed state changes. | The control packet validates its fixed lengths and required-zero field. `Docf_` values are selectors only, not CFB paths: this module performs no filesystem alternate-stream lookup or non-simple storage publication. |

Whole-object replacement starts with a target resolved by the host-format owner:

```rust
use litchi_cfb::OleError;
use litchi_ole_common::object::{Editor, Limits, Target, Targets};

fn replace_selected_object(
    source: Vec<u8>,
    target: Target,
    replacement: Vec<u8>,
) -> Result<Vec<u8>, OleError> {
    let key = target.key().to_owned();
    let mut editor = Editor::open(source, Targets::one(target), Limits::default())?;
    editor.replace(&key, replacement)?;
    editor.finish()
}
```

## Source preservation and execution boundaries

Snapshots share captured source allocations where the API permits it. OLE1
and OLEDS typed transactions validate a candidate, reparse it, and expose a
patch that checks its exact source before applying or inverting it. The common
CFB editor keeps unrelated streams and directory metadata in its captured
package and validates the complete candidate before replacing editor state.
Property Set sections retain supported ordering and unknown values in the
semantic model; an unchanged Property Set editor returns its original CFB
bytes.

The crate does not make host-format admission decisions. In particular, a
linked path is retained as inert text, native and presentation data are not
opened or decoded as documents or images, and no class, macro, or embedded
payload is activated. `NonSimpleAlternateStreamName` describes a binding
selector only; [`NonSimpleSnapshot`](src/property_set/non_simple/mod.rs) owns
the separate source-backed non-simple Property Set storage boundary. Its
indirect `propN` references are checked against physical CFB children before
changed publication, while unsupported directory kinds and root-name rewrites
refuse rather than dropping opaque data. Host-format reference checks belong
to the format layer. The common object editor exposes
`prepare_replacement` and `replace_prepared` so a host can validate a whole-object
replacement before cloning its outer editor. Opening consumes the original `Vec`,
and a uniquely owned unchanged finish returns that allocation.

## Validation evidence

The checked-in evidence batches bind their source files, commands, and logs:

- [OLE1 streams](../../docs/report/spec-gap-validation-evidence/ole1-streams-v1/README.md), including the read/preserve-only compatibility profiles and reversible edits ([receipt](../../docs/report/spec-gap-validation-evidence/ole1-streams-v1/receipt.json)).
- [OLE2 streams](../../docs/report/spec-gap-validation-evidence/ole2-streams-v1/README.md), covering presentation, TOC, native-stream, and target-selected object integration ([receipt](../../docs/report/spec-gap-validation-evidence/ole2-streams-v1/receipt.json)).
- [Common CFB and DOC source admission](../../docs/report/spec-gap-validation-evidence/doc-source-admission-v1/README.md), including original allocation retention and prepared replacement validation.
- [OLEPS binding names](../../docs/report/spec-gap-validation-evidence/oleps-binding-v1/README.md) and [alternate-stream metadata](../../docs/report/spec-gap-validation-evidence/oleps-alternate-stream-v1/README.md).
- In-repository integration coverage for [property-set editing](tests/property_set_editor.rs), [object integration](tests/object_streams.rs), [OLEDS streams](tests/ole_streams.rs), and [OLE1 objects](tests/ole1.rs).

These batches validate the bounded APIs and their refusal paths. They do not
claim universal native-producer compatibility, rendering, execution, or a
performance profile.

## License

Licensed under the Apache License, Version 2.0. Part of the [Litchi](https://github.com/DevExzh/litchi) workspace.
