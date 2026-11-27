//! Raw `SpreadsheetML` values converted into semantic workbook types.
//!
//! The XML grammar itself remains in [`crate::raw`]. This boundary only maps
//! [MS-XLSX] wire values into the typed facade and deliberately preserves
//! unknown worksheet states.

use litchi_opc::constants::content_type as ct;

use super::model::{Flavor, Visibility};
use crate::raw;

pub(super) fn flavor(value: &str) -> Option<Flavor> {
    match value {
        ct::SML_SHEET_MAIN => Some(Flavor::Workbook),
        ct::SML_TEMPLATE_MAIN => Some(Flavor::Template),
        ct::SML_SHEET_MACRO_MAIN => Some(Flavor::MacroWorkbook),
        ct::SML_TEMPLATE_MACRO_MAIN => Some(Flavor::MacroTemplate),
        _ => None,
    }
}

pub(super) fn visibility(value: raw::Visibility) -> Visibility {
    match value {
        raw::Visibility::Visible => Visibility::Visible,
        raw::Visibility::Hidden => Visibility::Hidden,
        raw::Visibility::VeryHidden => Visibility::VeryHidden,
        raw::Visibility::Unknown(value) => Visibility::Unknown(value),
    }
}
