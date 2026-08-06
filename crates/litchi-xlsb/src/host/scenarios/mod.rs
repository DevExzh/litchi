//! Layered XLSB Scenario Manager and what-if analysis support.
//!
//! The semantic model, BIFF12 payload codec, structural validation, worksheet
//! package splice, and focused regressions live in separate contextual files.
//! Scenario values are inert metadata and are never applied or recalculated.

mod codec;
mod model;
pub mod package;
mod validation;

#[cfg(test)]
mod tests;

pub use model::{
    CellRange, ChangedCell, MAX_CHANGED_CELLS, MAX_RESULT_RANGES, MAX_SCENARIO_TEXT, MAX_SCENARIOS,
    MAX_UNKNOWN_PAYLOAD, MAX_UNKNOWN_RECORDS, MAX_USER_NAME, Manager, Scenario, UnknownRecord,
};
pub use package::{parse_worksheet, replace_worksheet};
