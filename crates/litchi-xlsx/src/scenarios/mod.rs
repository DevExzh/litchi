//! Worksheet what-if scenarios.
//!
//! The facade keeps worksheet context at the module boundary while splitting
//! typed values, XML conversion, and regression coverage by responsibility.

mod codec;
mod model;
#[cfg(test)]
mod tests;

pub use codec::{parse_worksheet_scenarios, write_worksheet_scenarios};
pub use model::{
    CellReference, Collection, Conformance, InputCell, RangeReference, Scenario, UnknownAttribute,
    UnknownElement,
};
