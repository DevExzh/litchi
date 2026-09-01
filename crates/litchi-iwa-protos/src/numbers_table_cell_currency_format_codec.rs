//! Neutral strict native-currency format seam for Numbers table-cell adapters.
//!
//! Currency uses the shared source-preserving `FormatStructArchive` wire core
//! while keeping a nominal family boundary around its code, decimal, and
//! accounting fields. Generated Buffa values never cross this module.

#![allow(clippy::module_name_repetitions)]

use crate::numbers_table_cell_pop_up_menu_codec as core;

pub use core::{
    DecodeError, DecodeLimit, DecodeOptions, DecodeReport, NATIVE_AUTOMATIC_DECIMAL_PLACES,
    RewriteExecutionLimits, RewriteExecutionRequirements, RewriteOutput,
};

/// Native Numbers display-format discriminator for a currency cell.
pub use core::NATIVE_CURRENCY_FORMAT_TYPE;

/// Largest fixed decimal-place value accepted by Numbers' native currency
/// format.
pub const MAX_CURRENCY_DECIMAL_PLACES: u32 = core::MAX_NUMBER_DECIMAL_PLACES;

/// Borrowed scalar values for one strict native Currency format payload.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct CurrencyFormatSnapshot<'source>(core::CurrencyFormatSnapshot<'source>);

impl<'source> CurrencyFormatSnapshot<'source> {
    /// Borrow the original wire payload.
    #[must_use]
    pub const fn raw(self) -> &'source [u8] {
        self.0.raw()
    }

    /// Return the native Currency discriminator (`257`).
    #[must_use]
    pub const fn format_type(self) -> u32 {
        self.0.format_type()
    }

    /// Return the explicitly encoded decimal-place value, if present.
    #[must_use]
    pub const fn decimal_places(self) -> Option<u32> {
        self.0.decimal_places()
    }

    /// Return the explicitly encoded three-letter uppercase currency code.
    #[must_use]
    pub const fn currency_code(self) -> Option<&'source str> {
        self.0.currency_code()
    }

    /// Return the explicitly encoded native negative-number style.
    #[must_use]
    pub const fn negative_style(self) -> Option<u32> {
        self.0.negative_style()
    }

    /// Return the explicitly encoded thousands-separator setting.
    #[must_use]
    pub const fn show_thousands_separator(self) -> Option<bool> {
        self.0.show_thousands_separator()
    }

    /// Return the explicitly encoded accounting-style setting.
    #[must_use]
    pub const fn use_accounting_style(self) -> Option<bool> {
        self.0.use_accounting_style()
    }

    /// Resolve omitted decimal places using Numbers' native automatic value.
    #[must_use]
    pub const fn decimal_places_or_default(self) -> u32 {
        match self.decimal_places() {
            Some(value) => value,
            None => NATIVE_AUTOMATIC_DECIMAL_PLACES,
        }
    }

    /// Resolve an omitted currency code using Numbers' native default.
    #[must_use]
    pub const fn currency_code_or_default(self) -> &'source str {
        match self.currency_code() {
            Some(value) => value,
            None => "USD",
        }
    }

    /// Resolve an omitted accounting setting using Numbers' native default.
    #[must_use]
    pub const fn use_accounting_style_or_default(self) -> bool {
        match self.use_accounting_style() {
            Some(value) => value,
            None => false,
        }
    }
}

/// Owned scalar values accepted by the native Currency format writer.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct CurrencyFormatWrite<'source>(core::CurrencyFormatWrite<'source>);

impl<'source> CurrencyFormatWrite<'source> {
    /// Construct a complete native Currency format update.
    #[must_use]
    pub const fn new(
        currency_code: &'source str,
        decimal_places: u32,
        negative_style: u32,
        show_thousands_separator: bool,
        use_accounting_style: bool,
    ) -> Self {
        Self(core::CurrencyFormatWrite::new(
            currency_code,
            decimal_places,
            negative_style,
            show_thousands_separator,
            use_accounting_style,
        ))
    }

    /// Copy the semantic values and field presence from a decoded snapshot.
    #[must_use]
    pub const fn from_snapshot(snapshot: CurrencyFormatSnapshot<'source>) -> Self {
        Self(core::CurrencyFormatWrite::from_snapshot(snapshot.0))
    }

    /// Return the explicitly requested decimal-place value, if any.
    #[must_use]
    pub const fn decimal_places(self) -> Option<u32> {
        self.0.decimal_places()
    }

    /// Return the explicitly requested currency code, if any.
    #[must_use]
    pub const fn currency_code(self) -> Option<&'source str> {
        self.0.currency_code()
    }

    /// Return the explicitly requested negative-number style, if any.
    #[must_use]
    pub const fn negative_style(self) -> Option<u32> {
        self.0.negative_style()
    }

    /// Return the explicitly requested thousands-separator setting, if any.
    #[must_use]
    pub const fn show_thousands_separator(self) -> Option<bool> {
        self.0.show_thousands_separator()
    }

    /// Return the explicitly requested accounting-style setting, if any.
    #[must_use]
    pub const fn use_accounting_style(self) -> Option<bool> {
        self.0.use_accounting_style()
    }

    /// Set the decimal-place value.
    #[must_use]
    pub const fn with_decimal_places(self, value: u32) -> Self {
        Self(self.0.with_decimal_places(value))
    }

    /// Set the three-letter uppercase currency code.
    #[must_use]
    pub const fn with_currency_code(self, value: &'source str) -> Self {
        Self(self.0.with_currency_code(value))
    }

    /// Set the native negative-number style.
    #[must_use]
    pub const fn with_negative_style(self, value: u32) -> Self {
        Self(self.0.with_negative_style(value))
    }

    /// Set the thousands-separator setting.
    #[must_use]
    pub const fn with_show_thousands_separator(self, value: bool) -> Self {
        Self(self.0.with_show_thousands_separator(value))
    }

    /// Set the accounting-style setting.
    #[must_use]
    pub const fn with_use_accounting_style(self, value: bool) -> Self {
        Self(self.0.with_use_accounting_style(value))
    }
}

/// Prepared source-preserving native Currency rewrite.
#[derive(Debug, Clone, Copy)]
pub struct PreparedCurrencyFormatRewrite<'source>(core::PreparedCurrencyFormatRewrite<'source>);

impl PreparedCurrencyFormatRewrite<'_> {
    /// Return the measured requirements that execution must be allowed.
    #[must_use]
    pub const fn execution_requirements(self) -> RewriteExecutionRequirements {
        self.0.execution_requirements()
    }

    /// Return a report-shaped view of the prepared operation.
    #[must_use]
    pub fn prepare_report(self) -> DecodeReport {
        self.0.prepare_report()
    }

    /// Emit, strictly read back, and publish the source-preserving candidate.
    pub fn execute(self, limits: RewriteExecutionLimits) -> Result<RewriteOutput, DecodeError> {
        self.0.execute(limits)
    }
}

/// Prepared canonical native Currency append.
#[derive(Debug, Clone, Copy)]
pub struct PreparedCurrencyFormatWrite<'source>(core::PreparedCurrencyFormatWrite<'source>);

impl PreparedCurrencyFormatWrite<'_> {
    /// Return the finite measured requirements for execution.
    #[must_use]
    pub const fn execution_requirements(self) -> RewriteExecutionRequirements {
        self.0.execution_requirements()
    }

    /// Return a report-shaped view of the prepared operation.
    #[must_use]
    pub fn prepare_report(self) -> DecodeReport {
        self.0.prepare_report()
    }

    /// Emit and strictly read back a canonical Currency payload.
    pub fn execute(self, limits: RewriteExecutionLimits) -> Result<RewriteOutput, DecodeError> {
        self.0.execute(limits)
    }
}

/// Strictly decode one native Currency `FormatStructArchive`.
pub fn decode_currency_format(
    source: &[u8],
    options: DecodeOptions,
) -> Result<CurrencyFormatSnapshot<'_>, DecodeError> {
    core::decode_currency_format(source, options).map(CurrencyFormatSnapshot)
}

/// Strictly decode one native Currency format and return measured wire use.
pub fn decode_currency_format_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(CurrencyFormatSnapshot<'_>, DecodeReport), DecodeError> {
    core::decode_currency_format_with_report(source, options)
        .map(|(snapshot, report)| (CurrencyFormatSnapshot(snapshot), report))
}

/// Prepare a source-preserving native Currency format rewrite.
pub fn prepare_currency_format_rewrite<'source>(
    source: &'source [u8],
    write: CurrencyFormatWrite<'source>,
    options: DecodeOptions,
) -> Result<PreparedCurrencyFormatRewrite<'source>, DecodeError> {
    core::prepare_currency_format_rewrite(source, write.0, options)
        .map(PreparedCurrencyFormatRewrite)
}

/// Rewrite one native Currency format while preserving unknown source fields
/// byte-for-byte.
pub fn rewrite_currency_format(
    source: &[u8],
    write: CurrencyFormatWrite<'_>,
    options: DecodeOptions,
) -> Result<RewriteOutput, DecodeError> {
    core::rewrite_currency_format(source, write.0, options)
}

/// Compatibility spelling for the table-cell Currency route.
pub use rewrite_currency_format as rewrite_table_cell_currency_format;

/// Prepare a canonical native Currency format payload for a new list entry.
pub fn prepare_currency_format_write<'source>(
    write: CurrencyFormatWrite<'source>,
    options: DecodeOptions,
) -> Result<PreparedCurrencyFormatWrite<'source>, DecodeError> {
    core::prepare_currency_format_write(write.0, options).map(PreparedCurrencyFormatWrite)
}

/// Prepare a canonical Currency append under the explicit append spelling.
pub use prepare_currency_format_write as prepare_currency_format_append;

/// Encode a canonical native Currency format payload for a new list entry.
pub fn canonical_currency_format(
    write: CurrencyFormatWrite<'_>,
    options: DecodeOptions,
) -> Result<RewriteOutput, DecodeError> {
    core::canonical_currency_format(write.0, options)
}
