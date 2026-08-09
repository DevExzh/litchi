//! ODF data-pilot (pivot-table) declarations.
//!
//! The public contextual owner lives in [`crate::data_pilot`].  This module
//! retains the format vocabulary and standalone XML grammar that it reuses.

mod codec;
mod model;
mod range;
mod validation;

pub use model::{
    DisplayInfo, DisplayMemberMode, Field, FieldReference, GrandTotal, GrandTotalElement,
    GrandTotalOrientation, Group, GroupBoundary, GroupBy, Groups, LayoutInfo, LayoutMode, Level,
    Member, Orientation, ReferenceMemberType, ReferenceType, SortInfo, SortMode, SortOrder, Source,
    Table,
};

pub(crate) fn parse_data_pilot_tables(xml: &str) -> litchi_core::Result<Vec<Table>> {
    codec::parse_data_pilot_tables(xml)
}

pub(crate) fn write_data_pilot_tables(
    output: &mut String,
    tables: &[Table],
) -> litchi_core::Result<()> {
    codec::write_data_pilot_tables(output, tables)
}

/// # Errors
///
/// Returns an error when the value cannot be serialized.
pub fn write_data_pilot_table_fragment(table: &Table) -> litchi_core::Result<String> {
    codec::write_data_pilot_table_fragment(table)
}

pub(crate) fn parse_data_pilot_range(value: &str) -> litchi_core::Result<range::ParsedRange> {
    range::parse_data_pilot_range(value)
}

pub(crate) fn validate_data_pilot_tables(tables: &[Table]) -> litchi_core::Result<()> {
    validation::validate_data_pilot_tables(tables)
}

const TABLE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const OFFICE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TABLE_EXT_NAMESPACE: &[u8] =
    b"urn:org:documentfoundation:names:experimental:office:xmlns:table:1.0";
const CALC_EXT_NAMESPACE: &[u8] =
    b"urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0";
const MAX_DATA_PILOT_TABLES: usize = 65_536;
const MAX_DATA_PILOT_FIELDS: usize = 65_536;
const MAX_DATA_PILOT_ITEMS: usize = 1_000_000;
const MAX_DATA_PILOT_STRING: usize = 1024 * 1024;

pub(super) fn invalid(kind: &str, value: &str) -> litchi_core::Error {
    invalid_message(&format!("invalid {kind} '{value}'"))
}

pub(super) fn invalid_message(message: &str) -> litchi_core::Error {
    litchi_core::Error::InvalidFormat(message.to_string())
}
