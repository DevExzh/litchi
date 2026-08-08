//! Bounded, inert `PresentationML` notes-slide and notes-master package graphs.

mod codec;
mod model;
mod package;
mod transaction;
mod validation;

pub use codec::{master_xml, write_text, write_text_with};
pub use model::{Conformance, Graph, Link, Master, Slide, Theme};
pub(crate) use package::{
    apply_commit, apply_patch, clear_checked, load, load_snapshot, remove_checked,
};
#[cfg(test)]
pub(crate) use package::{clear, put, remove};
pub use transaction::{Commit, Patch, Revision, Snapshot, Transaction};

pub(crate) const P: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
pub(crate) const PS: &str = "http://purl.oclc.org/ooxml/presentationml/main";
pub(crate) const A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
pub(crate) const AS: &str = "http://purl.oclc.org/ooxml/drawingml/main";
pub(crate) const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
pub(crate) const RS: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
pub(crate) const SLIDE_CT: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
pub(crate) const THEME_CT: &str = "application/vnd.openxmlformats-officedocument.theme+xml";
pub(crate) const MAX_PRESENTATION_XML: usize = 32 * 1024 * 1024;
pub(crate) const MAX_SLIDE_XML: usize = 16 * 1024 * 1024;
pub(crate) const MAX_NOTES_XML: usize = 8 * 1024 * 1024;
pub(crate) const MAX_MASTER_XML: usize = 16 * 1024 * 1024;
pub(crate) const MAX_THEME_XML: usize = 16 * 1024 * 1024;
pub(crate) const MAX_TOTAL_BYTES: usize = 64 * 1024 * 1024;
pub(crate) const MAX_NOTES_SLIDES: usize = 4096;
pub(crate) const MAX_OWNED_PARTS: usize = 4_096;
pub(crate) const MAX_NODES: usize = 100_000;
pub(crate) const MAX_DEPTH: usize = 128;
pub(crate) const MAX_ATTRIBUTES: usize = 500_000;
pub(crate) const MAX_ATTRIBUTE_BYTES: usize = 8 * 1024 * 1024;

pub(crate) fn checked_add(left: usize, right: usize, label: &str) -> crate::Result<usize> {
    left.checked_add(right)
        .ok_or_else(|| invalid(format!("PPTX notes {label} overflow")))
}

pub(crate) fn own_blob(blob: &[u8], resource: &'static str) -> crate::Result<Vec<u8>> {
    let mut owned = Vec::new();
    owned
        .try_reserve_exact(blob.len())
        .map_err(|source| allocation(resource, source))?;
    owned.extend_from_slice(blob);
    Ok(owned)
}

pub(crate) fn xml_error(error: impl std::fmt::Display) -> crate::Error {
    crate::Error::Xml(error.to_string())
}

pub(crate) fn invalid(message: impl Into<String>) -> crate::Error {
    crate::Error::Invalid(message.into())
}

pub(crate) fn allocation(
    resource: &'static str,
    source: std::collections::TryReserveError,
) -> crate::Error {
    crate::Error::Allocation { resource, source }
}

pub(crate) fn limit(resource: &'static str, limit: usize) -> crate::Error {
    crate::Error::Limit { resource, limit }
}

pub(crate) fn resolved(value: quick_xml::name::ResolveResult<'_>) -> crate::Result<String> {
    match value {
        quick_xml::name::ResolveResult::Bound(quick_xml::name::Namespace(value)) => {
            Ok(std::str::from_utf8(value).map_err(xml_error)?.to_owned())
        },
        quick_xml::name::ResolveResult::Unbound => Ok(String::new()),
        quick_xml::name::ResolveResult::Unknown(prefix) => Err(invalid(format!(
            "unbound XML prefix '{}'",
            String::from_utf8_lossy(prefix.as_ref())
        ))),
    }
}

#[cfg(test)]
mod tests;
