//! Neutral strict plain-percentage format seam for Numbers table-cell adapters.
//!
//! Percentage payloads use the same four scalar fields and source-preserving
//! wire machinery as plain Number payloads.  The only family discriminator is
//! the native `FormatStructArchive.format_type` value (`258`), so this module
//! exposes typed wrappers around the single audited decimal-format core.
//! Generated Buffa values never cross the crate boundary.

#![allow(clippy::module_name_repetitions)]

use crate::numbers_table_cell_pop_up_menu_codec as core;

pub use core::{
    DecodeError, DecodeLimit, DecodeOptions, DecodeReport, NATIVE_AUTOMATIC_DECIMAL_PLACES,
    RewriteExecutionLimits, RewriteExecutionRequirements, RewriteOutput,
};

/// Native Numbers display-format discriminator for a plain percentage cell.
pub use core::NATIVE_PERCENTAGE_FORMAT_TYPE;

/// Largest fixed decimal-place value accepted by Numbers' native percentage
/// format.
pub const MAX_PERCENTAGE_DECIMAL_PLACES: u32 = core::MAX_NUMBER_DECIMAL_PLACES;

/// Borrowed scalar values for one strict native Percentage format payload.
///
/// This nominal wrapper prevents a decoded Number snapshot from crossing the
/// Percentage writer boundary even though both families share one private
/// decimal wire implementation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PercentageFormatSnapshot<'source>(core::NumberFormatSnapshot<'source>);

impl<'source> PercentageFormatSnapshot<'source> {
    /// Borrow the original wire payload.
    #[must_use]
    pub const fn raw(self) -> &'source [u8] {
        self.0.raw()
    }

    /// Return the native Percentage discriminator (`258`).
    #[must_use]
    pub const fn format_type(self) -> u32 {
        self.0.format_type()
    }

    /// Return `253` for automatic places, otherwise a fixed value in `0..=30`.
    #[must_use]
    pub const fn decimal_places(self) -> u32 {
        self.0.decimal_places()
    }

    /// Return the native negative-number style (`0..=3`).
    #[must_use]
    pub const fn negative_style(self) -> u32 {
        self.0.negative_style()
    }

    /// Whether the thousands separator is displayed.
    #[must_use]
    pub const fn show_thousands_separator(self) -> bool {
        self.0.show_thousands_separator()
    }
}

/// Owned scalar values accepted by the native Percentage format writer.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PercentageFormatWrite(core::NumberFormatWrite);

impl PercentageFormatWrite {
    /// Construct a native Percentage format update.
    #[must_use]
    pub const fn new(
        decimal_places: u32,
        negative_style: u32,
        show_thousands_separator: bool,
    ) -> Self {
        Self(core::NumberFormatWrite::new(
            decimal_places,
            negative_style,
            show_thousands_separator,
        ))
    }

    /// Copy the semantic values from a decoded Percentage snapshot.
    #[must_use]
    pub const fn from_snapshot(snapshot: PercentageFormatSnapshot<'_>) -> Self {
        Self(core::NumberFormatWrite::from_snapshot(snapshot.0))
    }

    /// Return the native decimal-place discriminator/value.
    #[must_use]
    pub const fn decimal_places(self) -> u32 {
        self.0.decimal_places()
    }

    /// Return the native negative-number style.
    #[must_use]
    pub const fn negative_style(self) -> u32 {
        self.0.negative_style()
    }

    /// Return the thousands-separator setting.
    #[must_use]
    pub const fn show_thousands_separator(self) -> bool {
        self.0.show_thousands_separator()
    }

    /// Replace the decimal-place discriminator/value.
    #[must_use]
    pub const fn with_decimal_places(self, value: u32) -> Self {
        Self(self.0.with_decimal_places(value))
    }

    /// Replace the native negative-number style.
    #[must_use]
    pub const fn with_negative_style(self, value: u32) -> Self {
        Self(self.0.with_negative_style(value))
    }

    /// Replace the thousands-separator setting.
    #[must_use]
    pub const fn with_show_thousands_separator(self, value: bool) -> Self {
        Self(self.0.with_show_thousands_separator(value))
    }
}

/// Prepared source-preserving native Percentage rewrite.
#[derive(Debug, Clone, Copy)]
pub struct PreparedPercentageFormatRewrite<'source>(core::PreparedNumberFormatRewrite<'source>);

impl PreparedPercentageFormatRewrite<'_> {
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

/// Prepared canonical native Percentage append.
#[derive(Debug, Clone, Copy)]
pub struct PreparedPercentageFormatWrite(core::PreparedNumberFormatWrite);

impl PreparedPercentageFormatWrite {
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

    /// Emit and strictly read back a canonical Percentage payload.
    pub fn execute(self, limits: RewriteExecutionLimits) -> Result<RewriteOutput, DecodeError> {
        self.0.execute(limits)
    }
}

/// Strictly decode one native Percentage `FormatStructArchive`.
pub fn decode_percentage_format(
    source: &[u8],
    options: DecodeOptions,
) -> Result<PercentageFormatSnapshot<'_>, DecodeError> {
    core::decode_decimal_format(source, NATIVE_PERCENTAGE_FORMAT_TYPE, options)
        .map(PercentageFormatSnapshot)
}

/// Strictly decode one native Percentage format and return measured wire use.
pub fn decode_percentage_format_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(PercentageFormatSnapshot<'_>, DecodeReport), DecodeError> {
    core::decode_decimal_format_with_report(source, NATIVE_PERCENTAGE_FORMAT_TYPE, options)
        .map(|(snapshot, report)| (PercentageFormatSnapshot(snapshot), report))
}

/// Prepare a source-preserving native Percentage format rewrite.
pub fn prepare_percentage_format_rewrite<'source>(
    source: &'source [u8],
    write: PercentageFormatWrite,
    options: DecodeOptions,
) -> Result<PreparedPercentageFormatRewrite<'source>, DecodeError> {
    core::prepare_decimal_format_rewrite(source, write.0, NATIVE_PERCENTAGE_FORMAT_TYPE, options)
        .map(PreparedPercentageFormatRewrite)
}

/// Rewrite one native Percentage format while preserving unknown source
/// fields byte-for-byte.
pub fn rewrite_percentage_format(
    source: &[u8],
    write: PercentageFormatWrite,
    options: DecodeOptions,
) -> Result<RewriteOutput, DecodeError> {
    core::rewrite_decimal_format(source, write.0, NATIVE_PERCENTAGE_FORMAT_TYPE, options)
}

/// Compatibility spelling for the table-cell Percentage route.
pub use rewrite_percentage_format as rewrite_table_cell_percentage_format;

/// Prepare a canonical native Percentage format payload for a new list entry.
pub fn prepare_percentage_format_write(
    write: PercentageFormatWrite,
    options: DecodeOptions,
) -> Result<PreparedPercentageFormatWrite, DecodeError> {
    core::prepare_decimal_format_write(write.0, NATIVE_PERCENTAGE_FORMAT_TYPE, options)
        .map(PreparedPercentageFormatWrite)
}

/// Prepare a canonical Percentage format append under the explicit append
/// spelling.
pub use prepare_percentage_format_write as prepare_percentage_format_append;

/// Encode a canonical native Percentage format payload for a new list entry.
pub fn canonical_percentage_format(
    write: PercentageFormatWrite,
    options: DecodeOptions,
) -> Result<RewriteOutput, DecodeError> {
    core::canonical_decimal_format(write.0, NATIVE_PERCENTAGE_FORMAT_TYPE, options)
}
