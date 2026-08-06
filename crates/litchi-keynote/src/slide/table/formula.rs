//! Archive-free formula values for a Keynote slide table.
//!
//! The neutral common leaf owns the canonical formula model because the same
//! values are used by multiple iWork formats. This focused Keynote module
//! provides the table-context import without exposing a concrete peer crate or
//! introducing a wrapper allocation.

#![allow(
    clippy::module_name_repetitions,
    reason = "ADR 4 keeps Formula-prefixed names explicit in shared formula facades"
)]

/// A typed whole-row or whole-column reference.
pub use litchi_iwa_common::formula::FormulaAxisReference;
/// A typed binary operator.
pub use litchi_iwa_common::formula::FormulaBinaryOperator;
/// A formula result displayed before the next native recalculation.
pub use litchi_iwa_common::formula::FormulaCachedValue;
/// A typed cell address.
pub use litchi_iwa_common::formula::FormulaCellReference;
/// A typed formula expression.
pub use litchi_iwa_common::formula::FormulaExpression;
/// An absolute category aggregate in a pivot table.
pub use litchi_iwa_common::formula::FormulaPivotCategoryReference;
/// The compact two-word identifier used by the pivot formula model.
pub use litchi_iwa_common::formula::FormulaUuid;
