# litchi-opc

Open Packaging Conventions (OPC) implementation: the ZIP-based container layer used by all OOXML formats.

## Overview

OPC defines how `.docx`, `.xlsx`, `.pptx`, and other Office Open XML documents
are packaged: parts, content types, and relationships inside a ZIP archive.
This crate provides the package model, `PackURI` resolution, and reader/writer
plumbing on top of `soapberry-zip` and `quick-xml`. It is consumed directly by
`litchi-docx`, `litchi-pptx`, `litchi-xlsx`, and `litchi-xlsb`; shared OOXML
vocabulary and graph services live in `litchi-ooxml-common`.

## Usage

```toml
[dependencies]
litchi-opc = "0.0.1"
```

```rust
use litchi_opc::{OpcPackage, Part};

let pkg = OpcPackage::open("example.docx")?;
for part in pkg.iter_parts() {
    println!("{} ({})", part.partname(), part.content_type());
}
# Ok::<(), litchi_opc::OpcError>(())
```

## Bounded Ingestion

All ordinary package constructors use bounded `ReadLimits::default()` values.
For untrusted or multi-tenant input, build a checked profile from those defaults
and use an `*_with_limits` constructor:

```rust
use litchi_opc::{OpcPackage, ReadLimits};

let limits = ReadLimits::builder()
    .max_input_bytes(32 * 1024 * 1024)?
    .max_archive_members(10_000)?
    .build()?;
let package = OpcPackage::open_with_limits("untrusted.docx", limits)?;
# let _ = package;
# Ok::<(), litchi_opc::OpcError>(())
```

The standard profile caps input at 512 MiB; ZIP members and materialized OPC
parts at 100,000 each; one ZIP member and one materialized part at 512 MiB;
and aggregate declared ZIP bytes at 2 GiB. It also bounds ZIP names, central
directory metadata, compressed bytes, aggregate materialized parts,
`[Content_Types].xml` and its mappings, relationship parts and XML, individual
and aggregate relationships, graph nodes, XML events and depth, and
relationship attribute and target lengths. `ReadResource` identifies the
specific rejected resource.

These ceilings are Litchi safety policy rather than ECMA-376 size maxima. They
implement a defensive consumer boundary for the physical package and
relationships described by ECMA-376 Part 2 sections 7.3.6 and 10, and the
corresponding MS-OI29500 sections 2.1.1749-1752. DOCX and PPTX expose
`Package::*_with_limits`; XLSX exposes both `Package::*_with_limits` and
`Workbook::*_with_limits`; XLSB exposes `Workbook::new_with_limits`. Each takes
the same `ReadLimits` profile so callers can use one contextual policy across
OOXML formats.

OPC readers treat macros, VBA, ActiveX, controls, OLE objects, and embedded
code as inert blobs only when they are retained or exposed. They are never
executed or activated.

## Features

- Streaming OPC package reader with content-type and relationship resolution
- `PackURI` parsing and normalisation per ISO/IEC 29500-2
- Zero-copy XML parsing via `quick-xml`, SIMD integer parsing via `atoi_simd`
- `PackageWriter` for authoring new OPC packages

## License

Licensed under the Apache License, Version 2.0. Part of the [Litchi](https://github.com/DevExzh/litchi) workspace.

## Owned XML publication

`OpcPackage::source_xml_part` issues an immutable `OwnedXmlPart` token after
validating bounded XML. Noncompact XML must match retained source bytes; new XML
must satisfy the authored compactness audit. The token shares its source buffer.

`source_content_types_with_limits` and `source_relationships_with_limits` capture
package metadata under a caller-selected `ReadLimits` profile. They check XML
bytes, declaration or relationship counts, and XML attributes for both retained
and regenerated XML. Their existing convenience methods use default limits.
Relationship capture compares the current graph with borrowed source bindings,
avoiding metadata clones just to decide whether source bytes can be retained.

`OwnedXmlPart::replace_attributes` accepts ordered, complete attribute-value
ranges and XML-escaped replacements. It checks the source spans and assembled
XML, retains all other source bytes, and caps source/output/replacement XML at
32 MiB and edits at 65,536. `try_replace_owned_xml_part` checks exact expected
bytes and content type before publishing; `try_add_owned_xml_part` supports exact
restoration or transfer. Both obey signature policy. Raw `set_blob` cannot grant
source provenance. Format owners still validate attribute semantics and their
reference graph; these primitives only establish XML and publication safety.

`insert_unqualified_attribute` verifies an exact opening tag and refuses duplicate
attributes and namespace declarations. `append_element`, `replace_element`, and
`remove_element` verify complete source element boundaries. New element fragments
must independently pass XML validation and the compact authoring audit; exact
replacement with the original subtree is a shared-buffer no-op. Empty parents
can be expanded while retaining their existing attributes and namespace context.
Every assembled result is bounded and validated before publication.

`update_attributes` accepts ordered `OwnedAttributeUpdate` values to replace,
insert, or remove unqualified attributes across multiple source opening tags.
It checks exact tag spans, duplicate names, escaped values, and the caller's
output limit before allocating the result. Attribute discovery uses one source
scan, reuses its name lookup table, and retains source order without sorting.
The assembled XML is validated before a new source token is returned. Update
debug output reports value lengths rather than values.

`update_elements` accepts ordered `OwnedElementUpdate` records for insertion
before an element, child append, replacement, and removal. A single source scan
locates complete element spans; overlapping destructive edits are rejected.
Multiple appends to an empty parent share one expansion. New fragments pass
standalone XML validation and the compact authoring audit, and the exact assembled
size is checked before allocating the result. Replacing a subtree with its exact
source bytes remains a shared-buffer no-op.

`replace_child_sequence` accepts ordered source child tags and a sequence of
`OwnedChildElement` values. Retained elements refer to positions in that selected
source list. They can move or repeat only under their original parent, preserving
inherited namespace and XML context. Authored elements are independently validated
and audited. Unselected source spans keep their slots, and additional elements
follow the last selected slot. The operation bounds its exact output size before
allocation and validates the assembled XML. An unchanged sequence shares the
source buffer.

The Custom Data owner in `litchi-xlsx` uses this path for storage and connection
UID edits, including exact inverse publication. Data Model load-version edits
and descriptor graph insertion/removal also retain source publication provenance
through saves and inverses. Survey scalar edits on existing elements use the
batch path, preserving comments, namespace context, and lexical spelling outside
the edited attributes. Survey property insertion, removal, and replacement with
fresh typed properties use structural batches in schema order while retaining
neighboring source XML. Question sequence edits move and copy opaque subtrees
within the original parent; column-name selectors resolve against the owning
table during ordinary edits. Survey snapshots retain source tokens for exact
inverse publication and restoration of removed parts. The positional, budget-managed
`SourceXmlPart` API remains available for source-backed package editing.
