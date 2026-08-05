//! Layered BIFF8 worksheet data-validation owner.
//!
//! The semantic model, record codecs, and regression coverage live in their
//! own layers. The legacy crate-level names are retained here only so the
//! existing crate facade and worksheet reader can continue to resolve the
//! canonical owner without caller edits.

mod codec;
mod model;

#[cfg(test)]
mod tests;

pub(crate) use codec::{parse_dv, parse_dval};
pub(crate) use model::{DV_RECORD_TYPE, DVAL_RECORD_TYPE};

pub use model::{
    ErrorStyle as DataValidationErrorStyle, Formula as DataValidationFormula,
    ImeMode as DataValidationImeMode, Kind as DataValidationKind,
    Operator as DataValidationOperator, Range as DataValidationRange, Rule as DataValidationRule,
    Settings as DataValidationSettings,
};
