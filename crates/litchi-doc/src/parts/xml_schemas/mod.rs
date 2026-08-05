//! Word 2003 XML schema reference tables (`Hplxsdr`) and the custom XML
//! save transform (`fcCustomXForm`).
//!
//! The `Hplxsdr` (MS-DOC 2.9.117) records the XML schema definitions a
//! document references: each `XSDR` (MS-DOC 2.9.352) carries the schema URI,
//! an optional expansion-pack manifest location, and the element and
//! attribute name string tables that namespace the `TIQ` name references
//! (MS-DOC 2.9.325) of structured document tags. `fcCustomXForm` points at a
//! UTF-16 path of the XML stylesheet Word applies when saving the document
//! in XML format.
//!
//! All structures are parsed as inert metadata: no schema or stylesheet is
//! fetched, resolved, or applied, and no document content is modified.

mod codec;
mod model;

#[cfg(test)]
mod tests;

pub use codec::parse_custom_xml_transform;
pub use model::{Collection, Reference};
