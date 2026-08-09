//! Typed `SpreadsheetML` named-sheet-view metadata.
//!
//! The owner is layered by responsibility: semantic declarations in
//! model, bounded XML/MCE conversion in codec, and OPC relationship
//! operations in package.

mod codec;
mod model;
mod package;

pub use codec::{parse_named_sheet_views, write_named_sheet_views};
pub use model::{
    ColumnFilter, DifferentialFormat, Extension, Filter, Guid, IconSet, Markup, Range,
    SortCondition, SortConditionKind, SortRule, SortRules, View, Views,
};
pub use package::{
    discover_named_sheet_views, load_worksheet_named_sheet_views,
    remove_worksheet_named_sheet_views, store_worksheet_named_sheet_views,
};

pub(crate) const NSV: &[u8] =
    b"http://schemas.microsoft.com/office/spreadsheetml/2019/namedsheetviews";
pub(crate) const CORE: &[u8] = b"http://schemas.openxmlformats.org/spreadsheetml/2006/main";
pub(crate) const STRICT: &[u8] = b"http://purl.oclc.org/ooxml/spreadsheetml/main";
pub(crate) const X14: &[u8] = b"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
pub(crate) const RICH: &[u8] = b"http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2";
pub(crate) const RELATIONSHIP: &str =
    "http://schemas.microsoft.com/office/2019/04/relationships/namedSheetView";
pub(crate) const CONTENT_TYPE: &str = "application/vnd.ms-excel.namedsheetviews+xml";
pub(crate) const MAX_PART_BYTES: usize = 32 * 1024 * 1024;
pub(crate) const MAX_VIEWS: usize = 1024;
pub(crate) const MAX_FILTERS: usize = 65_536;
pub(crate) const MAX_COLUMNS: usize = 16_384;
pub(crate) const MAX_EXTENSIONS: usize = 65_536;
pub(crate) const MAX_RETAINED_BYTES: usize = 8 * 1024 * 1024;
pub(crate) const MAX_MARKUP_BYTES: usize = 1024 * 1024;
pub(crate) const MAX_NAMESPACE_DECLARATIONS: usize = 256;
pub(crate) const MAX_FRAGMENT_DEPTH: usize = 256;
pub(crate) const MAX_FRAGMENT_NODES: usize = 100_000;
pub(crate) const FRAGMENT_VIEW_ID: &str = "{01234567-89AB-CDEF-0123-456789ABCDEF}";
pub(crate) const FRAGMENT_FILTER_ID: &str = "{11111111-2222-3333-4444-555555555555}";

use crate::error::Error;
use litchi_ooxml_common::Error as CommonError;

pub(crate) fn invalid(value: impl Into<String>) -> Error {
    Error::Invalid(value.into())
}

pub(crate) fn uri_error(value: impl std::fmt::Display) -> Error {
    Error::Common(CommonError::Uri(value.to_string()))
}

pub(crate) fn content_type_mismatch(expected: &str, actual: &str) -> Error {
    Error::Common(CommonError::ContentType {
        expected: expected.into(),
        actual: actual.into(),
    })
}

pub(crate) fn xml_error(value: impl std::fmt::Display) -> Error {
    Error::Xml(litchi_ooxml_common::XmlError::Malformed(value.to_string()))
}
