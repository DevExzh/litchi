//! OOXML custom-property schema vocabulary and resource limits.

use crate::{Error, Result};

pub(super) const PART_NAME: &str = "/docProps/custom.xml";
pub(super) const PART_TARGET: &str = "docProps/custom.xml";
pub(super) const FORMAT_ID: &str = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}";
pub(super) const SUMMARY_FORMAT_ID: &str = "{F29F85E0-4FF9-1068-AB91-08002B27B3D9}";
pub(super) const DOCUMENT_SUMMARY_FORMAT_ID: &str = "{D5CDD502-2E9C-101B-9397-08002B2CF9AE}";
pub(super) const CUSTOM_NS: &[u8] =
    b"http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";
pub(super) const VT_NS: &[u8] =
    b"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";

pub(super) const MAX_XML_BYTES: usize = 8 * 1024 * 1024;
pub(super) const MAX_XML_DEPTH: usize = 3;
pub(super) const MAX_XML_NODES: usize = 4_096;
pub(super) const MAX_ATTRIBUTES: usize = 32;
pub(super) const MAX_ATTRIBUTE_BYTES: usize = 4_096;
pub(super) const MAX_PROPERTIES: usize = 1_024;
pub(super) const MAX_NAME_CHARS: usize = 255;
pub(super) const MAX_NAME_BYTES: usize = 1_024;
pub(super) const MAX_TOTAL_NAME_BYTES: usize = 256 * 1024;
pub(super) const MAX_TEXT_BYTES: usize = 1024 * 1024;
pub(super) const MAX_TOTAL_TEXT_BYTES: usize = 1024 * 1024;

pub(super) fn checked_depth(depth: usize) -> Result<usize> {
    let actual = checked_increment(depth, "custom-properties XML depth")?;
    if actual > MAX_XML_DEPTH {
        return Err(limit("custom-properties XML depth", MAX_XML_DEPTH, actual));
    }
    Ok(actual)
}

pub(super) fn count_node(nodes: &mut usize) -> Result<()> {
    let actual = checked_increment(*nodes, "custom-properties XML nodes")?;
    if actual > MAX_XML_NODES {
        return Err(limit("custom-properties XML nodes", MAX_XML_NODES, actual));
    }
    *nodes = actual;
    Ok(())
}

pub(super) fn checked_increment(value: usize, resource: &'static str) -> Result<usize> {
    value
        .checked_add(1)
        .ok_or_else(|| limit(resource, usize::MAX, usize::MAX))
}

pub(super) fn checked_total(
    current: usize,
    added: usize,
    max: usize,
    resource: &'static str,
) -> Result<usize> {
    let actual = current
        .checked_add(added)
        .ok_or_else(|| limit(resource, max, usize::MAX))?;
    if actual > max {
        return Err(limit(resource, max, actual));
    }
    Ok(actual)
}

pub(super) fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

pub(super) fn limit(resource: &'static str, max: usize, actual: usize) -> Error {
    Error::Limit {
        resource,
        max,
        actual,
    }
}
