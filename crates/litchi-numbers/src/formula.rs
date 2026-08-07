//! Shared iWork formula vocabulary exposed through the Numbers crate.
//!
//! The neutral owner is [`litchi_iwa_common::formula`]. Numbers keeps this
//! focused module as its direct semantic entry point while archive-bound
//! compilation remains private to the Numbers package adapter.

#![allow(
    clippy::module_name_repetitions,
    reason = "Formula-prefixed names keep the shared public vocabulary explicit at call sites"
)]

pub use litchi_iwa_common::formula::{
    FormulaAxisReference, FormulaBinaryOperator, FormulaCachedValue, FormulaCellReference,
    FormulaExpression, FormulaPivotCategoryReference, FormulaUuid,
};
