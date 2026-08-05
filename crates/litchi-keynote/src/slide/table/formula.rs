//! Archive-free formula values for a Keynote slide table.
//!
//! Numbers owns the canonical formula model because the same values are used
//! by multiple iWork formats. The Keynote leaf gives those values a focused
//! table context without introducing a wrapper type or an allocation-bearing
//! duplicate.

#![allow(
    clippy::module_name_repetitions,
    reason = "ADR 4 keeps Formula-prefixed names explicit in shared formula facades"
)]

/// A typed whole-row or whole-column reference.
pub use litchi_numbers::formula::FormulaAxisReference;
/// A typed binary operator.
pub use litchi_numbers::formula::FormulaBinaryOperator;
/// A formula result displayed before the next native recalculation.
pub use litchi_numbers::formula::FormulaCachedValue;
/// A typed cell address.
pub use litchi_numbers::formula::FormulaCellReference;
/// A typed formula expression.
pub use litchi_numbers::formula::FormulaExpression;
/// An absolute category aggregate in a pivot table.
pub use litchi_numbers::formula::FormulaPivotCategoryReference;
/// The compact two-word identifier used by the pivot formula model.
pub use litchi_numbers::formula::FormulaUuid;
