//! Numbers cell vocabulary.
//!
//! Native Binary Numbers Cell (BNC) storage is an implementation detail. Use
//! the semantic [`Value`](crate::cell::Value) and
//! [`data_format`](crate::cell::data_format) APIs instead of depending on its
//! byte layout.

/// Archive-free direct replies attached to a cell comment.
pub mod comment;

/// Checked, archive-free cell display formats.
pub mod data_format;
/// Unified archive-free interactive cell controls.
///
/// The value types are defined by the focused data-format modules and are
/// re-exported here as a selector-friendly entry point. Native control-cell
/// records and package transactions remain private to the package adapter.
pub use data_format::control::CellControl;

/// Archive-free semantic cell values shared by the concrete iWork owners.
pub use litchi_iwa_common::table::cell::value::{
    APPLE_EPOCH_UNIX_OFFSET_SECONDS, FiniteF64, FiniteF64Error, Type, Update, Value,
};
/// Native BNC adapters stay crate-private.  Their wire/common finite scalar
/// is converted to the shared semantic [`FiniteF64`] at this boundary.
pub(crate) mod wire {
    use core::ops::{Deref, DerefMut};

    use litchi_numbers_wire as native;

    use super::FiniteF64;

    #[cfg(test)]
    pub(crate) use native::decimal128_le;
    pub(crate) use native::{ClearValue, Error, RewritePlan, StoredValue};

    /// A finite scalar accepted by the private BNC rewrite adapter.
    #[derive(Debug, Clone, Copy, PartialEq)]
    pub(crate) enum ScalarValue {
        String(u32),
        RichText(u32),
        Number(FiniteF64),
        Boolean(bool),
        Date(FiniteF64),
        Duration(FiniteF64),
    }

    impl ScalarValue {
        fn into_native(self) -> native::ScalarValue {
            match self {
                Self::String(identifier) => native::ScalarValue::String(identifier),
                Self::RichText(identifier) => native::ScalarValue::RichText(identifier),
                Self::Number(value) => native::ScalarValue::Number(to_native(value)),
                Self::Boolean(value) => native::ScalarValue::Boolean(value),
                Self::Date(value) => native::ScalarValue::Date(to_native(value)),
                Self::Duration(value) => native::ScalarValue::Duration(to_native(value)),
            }
        }
    }

    /// A decoded BNC scalar represented in the Numbers semantic vocabulary.
    #[derive(Debug, Clone, Copy, PartialEq)]
    pub(crate) enum CachedScalar {
        Number(FiniteF64),
        Boolean(bool),
        Date(FiniteF64),
        Duration(FiniteF64),
        Unsupported(u8),
    }

    impl CachedScalar {
        fn from_native(value: native::CachedScalar) -> Self {
            match value {
                native::CachedScalar::Number(value) => Self::Number(from_native(value)),
                native::CachedScalar::Boolean(value) => Self::Boolean(value),
                native::CachedScalar::Date(value) => Self::Date(from_native(value)),
                native::CachedScalar::Duration(value) => Self::Duration(from_native(value)),
                native::CachedScalar::Unsupported(value) => Self::Unsupported(value),
            }
        }

        fn into_native(self) -> native::CachedScalar {
            match self {
                Self::Number(value) => native::CachedScalar::Number(to_native(value)),
                Self::Boolean(value) => native::CachedScalar::Boolean(value),
                Self::Date(value) => native::CachedScalar::Date(to_native(value)),
                Self::Duration(value) => native::CachedScalar::Duration(to_native(value)),
                Self::Unsupported(value) => native::CachedScalar::Unsupported(value),
            }
        }
    }

    /// Owned BNC cell adapter with semantic scalar conversion at the seam.
    pub(crate) struct BncCell(native::BncCell);

    impl Deref for BncCell {
        type Target = native::BncCell;

        fn deref(&self) -> &Self::Target {
            &self.0
        }
    }

    impl DerefMut for BncCell {
        fn deref_mut(&mut self) -> &mut Self::Target {
            &mut self.0
        }
    }

    impl BncCell {
        pub(crate) fn parse(data: &[u8]) -> Result<Self, Error> {
            native::BncCell::parse(data).map(Self)
        }

        #[cfg(test)]
        #[must_use]
        pub(crate) fn minimal() -> Self {
            Self(native::BncCell::minimal())
        }

        #[cfg(test)]
        pub(crate) fn cached_scalar(&self) -> Result<Option<CachedScalar>, Error> {
            self.0
                .cached_scalar()
                .map(|value| value.map(CachedScalar::from_native))
        }
    }

    /// Borrowed BNC view adapter with semantic scalar conversion at the seam.
    pub(crate) struct BncCellView<'a>(native::BncCellView<'a>);

    impl<'a> BncCellView<'a> {
        pub(crate) fn parse(data: &'a [u8]) -> Result<Self, Error> {
            native::BncCellView::parse(data).map(Self)
        }

        #[must_use]
        pub(crate) fn stored_value(&self) -> StoredValue {
            self.0.stored_value()
        }

        #[must_use]
        pub(crate) fn cached_scalar(&self) -> Option<CachedScalar> {
            self.0.cached_scalar().map(CachedScalar::from_native)
        }

        #[must_use]
        pub(crate) fn formula_text_key(&self) -> Option<u32> {
            self.0.formula_text_key()
        }

        #[must_use]
        pub(crate) fn scalar_equals(&self, value: ScalarValue) -> bool {
            self.0.scalar_equals(value.into_native())
        }

        pub(crate) fn plan_scalar_rewrite(&self, value: ScalarValue) -> Result<RewritePlan, Error> {
            self.0.plan_scalar_rewrite(value.into_native())
        }

        pub(crate) fn plan_formula_rewrite(
            &self,
            identifier: u32,
            cache: Option<ScalarValue>,
        ) -> Result<RewritePlan, Error> {
            self.0
                .plan_formula_rewrite(identifier, cache.map(ScalarValue::into_native))
        }

        pub(crate) fn plan_formula_cache_rewrite(
            &self,
            cache: CachedScalar,
        ) -> Result<RewritePlan, Error> {
            self.0.plan_formula_cache_rewrite(cache.into_native())
        }

        pub(crate) fn plan_clear_value(&self, retain_empty: bool) -> Result<RewritePlan, Error> {
            self.0.plan_clear_value(retain_empty)
        }

        pub(crate) fn rewrite_scalar_with_limit(
            &self,
            value: ScalarValue,
            max_output_bytes: usize,
        ) -> Result<Vec<u8>, Error> {
            self.0
                .rewrite_scalar_with_limit(value.into_native(), max_output_bytes)
        }

        pub(crate) fn clear_value_with_limit(
            &self,
            max_output_bytes: usize,
        ) -> Result<ClearValue, Error> {
            self.0.clear_value_with_limit(max_output_bytes)
        }

        pub(crate) fn clear_comment_with_limit(
            &self,
            expected_identifier: u32,
            max_output_bytes: usize,
        ) -> Result<Vec<u8>, Error> {
            self.0
                .clear_comment_with_limit(expected_identifier, max_output_bytes)
        }

        pub(crate) fn formula_cache_equals(&self, value: CachedScalar) -> bool {
            self.0.formula_cache_equals(value.into_native())
        }

        pub(crate) fn formula_value_equals(
            &self,
            identifier: u32,
            value: ScalarValue,
        ) -> Result<bool, Error> {
            self.0.formula_value_equals(identifier, value.into_native())
        }

        pub(crate) fn rewrite_formula_with_limit(
            &self,
            identifier: u32,
            cache: ScalarValue,
            max_output_bytes: usize,
        ) -> Result<Vec<u8>, Error> {
            self.0
                .rewrite_formula_with_limit(identifier, cache.into_native(), max_output_bytes)
        }

        pub(crate) fn rewrite_formula_without_cache_with_limit(
            &self,
            identifier: u32,
            max_output_bytes: usize,
        ) -> Result<Vec<u8>, Error> {
            self.0
                .rewrite_formula_without_cache_with_limit(identifier, max_output_bytes)
        }

        pub(crate) fn rewrite_formula_cache_with_limit(
            &self,
            cache: CachedScalar,
            max_output_bytes: usize,
        ) -> Result<Vec<u8>, Error> {
            self.0
                .rewrite_formula_cache_with_limit(cache.into_native(), max_output_bytes)
        }

        #[must_use]
        pub(crate) fn formula_error_identifier(&self) -> Option<u32> {
            self.0.formula_error_identifier()
        }

        #[must_use]
        pub(crate) fn comment_identifier(&self) -> Option<u32> {
            self.0.comment_identifier()
        }
    }

    fn from_native(value: litchi_iwa_common::formula::FiniteF64) -> FiniteF64 {
        FiniteF64::new(value.get()).expect("wire parser guarantees finite scalar")
    }

    fn to_native(value: FiniteF64) -> litchi_iwa_common::formula::FiniteF64 {
        litchi_iwa_common::formula::FiniteF64::new(value.get())
            .expect("Numbers scalar invariant guarantees finite value")
    }
}
