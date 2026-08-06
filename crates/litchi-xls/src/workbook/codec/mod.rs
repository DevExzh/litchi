//! Layered BIFF workbook codec facade.
//
// The facade keeps workbook/package integration stable while placing
// record-wire helpers, semantic collectors, validation, and focused tests
// in their respective owners.

mod semantic;
mod validation;
mod wire;

#[cfg(test)]
mod tests;

use crate::defined_names::DefinedNameSlot;
use crate::leniency::ToleranceLog;
use crate::records::{BoundSheetRecord, SharedStringProperties};

pub(super) use wire::pivot_cache_stream_paths;

/// Inputs owned by the workbook package and consumed by the globals collector.
pub(crate) struct WorkbookGlobalsSink<'a> {
    /// BoundSheet8 entries in stream order.
    pub(super) bound_sheets: &'a mut Vec<BoundSheetRecord>,
    /// Shared-string table contents.
    pub(super) strings: &'a mut Vec<String>,
    /// Rich-text and phonetic properties parallel to strings.
    pub(super) string_properties: &'a mut Vec<Option<Box<SharedStringProperties>>>,
    /// Lbl records and their trailing optional records.
    pub(super) defined_name_slots: &'a mut Vec<DefinedNameSlot>,
    /// Formatting-defect policy and the repairs it recorded.
    pub(super) tolerance: &'a mut ToleranceLog,
}
