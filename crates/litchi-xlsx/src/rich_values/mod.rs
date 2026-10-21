//! Layered, inert XLSX rich-value and feature-property-bag ownership.
//!
//! The semantic model is deliberately independent from cell recalculation,
//! rendering, control activation, and service refresh. XML codecs retain
//! bounded extension/unknown subtrees, while [`package`] snapshots the OPC
//! parts and every relationship edge without following external targets.

mod model;
mod validation;

pub mod codec;
pub mod package;
pub mod refresh_intervals;

#[cfg(test)]
mod tests;

use crate::error::{Error, Result};

pub(crate) const RICH_DATA: &str =
    "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata";
pub(crate) const RICH_DATA_2: &str =
    "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2";
pub(crate) const FEATURE_BAG: &str =
    "http://schemas.microsoft.com/office/spreadsheetml/2022/featurepropertybag";
pub(crate) const RICH_VALUE_REL: &str =
    "http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel";
pub(crate) const RELATIONSHIPS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
pub(crate) const STRICT_RELATIONSHIPS: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships";
pub(crate) const SPREADSHEETML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

pub(crate) const MAX_XML_BYTES: usize = 16 * 1024 * 1024;
pub(crate) const MAX_OUTPUT_BYTES: usize = 32 * 1024 * 1024;
pub(crate) const MAX_OPAQUE_BYTES: usize = 8 * 1024 * 1024;
pub(crate) const MAX_STRING_BYTES: usize = 1024 * 1024;
pub(crate) const MAX_NODES: usize = 1_000_000;
pub(crate) const MAX_DEPTH: usize = 256;
pub(crate) const MAX_ITEMS: usize = 1_000_000;
pub(crate) const MAX_BAGS: usize = 65_536;
pub(crate) const MAX_RELATIONSHIPS: usize = 65_536;
pub(crate) const OFFICE_MAX_COUNT: u32 = 2_147_483_647;
pub(crate) const MAX_ARRAY_ROWS: u32 = 1_048_576;
pub(crate) const MAX_ARRAY_COLUMNS: u32 = 16_384;

crate::record::limit_kinds!("rich-values", RichValues {
    XmlBytes => "rich-values XML bytes",
    OutputBytes => "rich-values output bytes",
    OpaqueBytes => "rich-values opaque XML bytes",
    StringBytes => "rich-values string bytes",
    Nodes => "rich-values XML nodes",
    Depth => "rich-values XML depth",
    ValueItems => "rich-value items",
    StructureItems => "rich-value structure items",
    StructureKeys => "rich-value structure keys",
    ArrayItems => "rich-value array items",
    BagItems => "feature property bags",
    RelationshipItems => "rich-value relationships",
    ExtensionItems => "rich-value extension items",
    PropertyItems => "feature property items",
    PropertyArrayItems => "feature property array items",
    ArrayValues => "rich-array values",
    RelationshipEdges => "rich-value relationship edges",
});

pub(crate) fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

pub(crate) fn limit(kind: LimitKind, observed: usize, maximum: usize) -> Error {
    Error::limit(kind, observed, maximum)
}

pub(crate) fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::Xml(litchi_ooxml_common::XmlError::Malformed(error.to_string()))
}

pub(crate) fn bounded(value: &str, kind: LimitKind) -> Result<()> {
    if value.len() <= MAX_STRING_BYTES {
        Ok(())
    } else {
        Err(limit(kind, value.len(), MAX_STRING_BYTES))
    }
}

pub(crate) fn bounded_nonempty(value: &str, field: &'static str, kind: LimitKind) -> Result<()> {
    if value.is_empty() {
        return Err(invalid(format!("{field} cannot be empty")));
    }
    bounded(value, kind)
}

pub use model::{
    Array, ArrayData, ArrayValue, ArrayValueType, Bag, BagType, Bags, Checkbox, CheckboxState,
    DxfComplement, Fallback, FallbackType, Key, Link, Mode, Opaque, Property, PropertyValue,
    RichValue, RichValueData, RichValueRels, Structure, Structures, ValueType, XfComplement,
};
pub use validation::{
    validate_arrays, validate_bags, validate_data, validate_dxf_complement,
    validate_rich_value_rels, validate_structures, validate_xf_complement,
};
