//! Structured data extraction from iWork documents.
//!
//! Archive traversal is kept in focused private adapters. The public result
//! remains composed entirely of the semantic values owned by the leaf crates:
//! [`litchi_numbers::Table`], [`litchi_keynote::Slide`], and
//! [`litchi_pages::Section`].

mod keynote;
mod numbers;
mod pages;

use crate::Result;
use crate::bundle::Bundle;
use crate::object_index::ObjectIndex;
use litchi_keynote::Slide;
use litchi_numbers::Table;
use litchi_pages::Section;

pub use litchi_iwa_structured::StructuredData;

/// Extract Numbers tables as canonical leaf-owned semantic values.
pub(crate) fn extract_tables(bundle: &Bundle, object_index: &ObjectIndex) -> Result<Vec<Table>> {
    numbers::extract(bundle, object_index)
}

/// Extract Keynote slides as canonical leaf-owned semantic values.
pub(crate) fn extract_slides(bundle: &Bundle, object_index: &ObjectIndex) -> Result<Vec<Slide>> {
    keynote::extract(bundle, object_index)
}

/// Extract the main Pages body section as a canonical leaf-owned value.
pub(crate) fn extract_sections(
    bundle: &Bundle,
    object_index: &ObjectIndex,
) -> Result<Vec<Section>> {
    pages::extract(bundle, object_index)
}

/// Extract all supported structured data from one document snapshot.
pub(crate) fn extract_all(bundle: &Bundle, object_index: &ObjectIndex) -> Result<StructuredData> {
    Ok(StructuredData {
        tables: extract_tables(bundle, object_index)?,
        slides: extract_slides(bundle, object_index)?,
        sections: extract_sections(bundle, object_index)?,
    })
}

#[cfg(test)]
mod tests;
