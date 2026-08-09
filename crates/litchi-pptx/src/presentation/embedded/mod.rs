//! Bounded embedded-resource services for a `PresentationML` graph.
//!
//! Each child module keeps its semantic model, bounded XML codec, and OPC
//! graph operations together. Executable payloads remain inert: controls and
//! OLE objects expose metadata only, VBA exposes an opaque package part, and
//! `InkML` exposes structural counts rather than handwriting data, and general
//! `PresentationML` content parts expose bounded opaque payloads without
//! interpreting or executing their contents.

pub mod content_parts;
pub mod controls;
pub mod excel;
pub mod ink;
pub mod ole;
#[cfg(feature = "vba-inspection")]
pub mod vba;

pub(crate) const PML: &[u8] = b"http://schemas.openxmlformats.org/presentationml/2006/main";
pub(crate) const STRICT_PML: &[u8] = b"http://purl.oclc.org/ooxml/presentationml/main";
pub(crate) const DML: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/main";
pub(crate) const REL: &[u8] =
    b"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
pub(crate) const STRICT_REL: &[u8] = b"http://purl.oclc.org/ooxml/officeDocument/relationships";
pub(crate) const MAX_XML_BYTES: usize = 32 * 1024 * 1024;
pub(crate) const MAX_XML_NODES: usize = 250_000;
pub(crate) const MAX_XML_DEPTH: usize = 128;
pub(crate) const MAX_XML_ATTRIBUTES: usize = 64;
pub(crate) const MAX_ATTRIBUTE_BYTES: usize = 4 * 1024;

pub(crate) fn invalid(message: impl Into<String>) -> crate::Error {
    crate::Error::Invalid(message.into())
}

pub(crate) fn limit(resource: &'static str, maximum: usize) -> crate::Error {
    crate::Error::Limit {
        resource,
        limit: maximum,
    }
}

pub(crate) fn is_presentationml_name(
    namespace: &quick_xml::name::ResolveResult<'_>,
    name: quick_xml::name::QName<'_>,
    local: &[u8],
) -> bool {
    if name.local_name().as_ref() != local {
        return false;
    }
    match namespace {
        quick_xml::name::ResolveResult::Bound(quick_xml::name::Namespace(value)) => {
            *value == PML || *value == STRICT_PML
        },
        quick_xml::name::ResolveResult::Unknown(prefix) => prefix.as_slice() == b"p",
        quick_xml::name::ResolveResult::Unbound => false,
    }
}

pub(crate) fn relationship_value(
    element: &quick_xml::events::BytesStart<'_>,
    name: &[u8],
    decoder: quick_xml::encoding::Decoder,
    resolver: &quick_xml::name::NamespaceResolver,
) -> crate::Result<Option<String>> {
    Ok(litchi_ooxml_common::relationships::attribute_value(
        element, name, decoder, resolver,
    )?)
}

pub(crate) fn bounded(value: &str, resource: &'static str) -> crate::Result<()> {
    if value.len() > MAX_ATTRIBUTE_BYTES {
        return Err(limit(resource, MAX_ATTRIBUTE_BYTES));
    }
    Ok(())
}

pub(crate) fn increment_nodes(nodes: &mut usize) -> crate::Result<()> {
    *nodes = nodes
        .checked_add(1)
        .ok_or_else(|| limit("embedded XML nodes", MAX_XML_NODES))?;
    if *nodes > MAX_XML_NODES {
        return Err(limit("embedded XML nodes", MAX_XML_NODES));
    }
    Ok(())
}

pub(crate) fn validate_root(
    namespace: &quick_xml::name::ResolveResult<'_>,
    name: quick_xml::name::QName<'_>,
    seen: bool,
) -> crate::Result<()> {
    if seen || !is_presentationml_name(namespace, name, b"sld") {
        return Err(invalid(
            "embedded-resource XML must have one PresentationML sld root",
        ));
    }
    Ok(())
}
