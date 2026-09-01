//! Neutral strict scientific-format seam for Numbers table-cell adapters.
//!
//! Scientific payloads use the same four scalar fields and source-preserving
//! wire machinery as plain Number and Percentage payloads. The native family
//! discriminator is `259`, and Scientific additionally fixes the negative
//! style to the native minus-sign value and hides the thousands separator.
//! Generated Buffa values never cross the crate boundary.

#![allow(clippy::module_name_repetitions)]

use crate::numbers_table_cell_pop_up_menu_codec as core;

pub use core::{
    DecodeError, DecodeLimit, DecodeOptions, DecodeReport, RewriteExecutionLimits,
    RewriteExecutionRequirements, RewriteOutput,
};

/// Native Numbers display-format discriminator for a scientific cell.
pub use core::NATIVE_SCIENTIFIC_FORMAT_TYPE;

/// Native discriminator used for automatic decimal places by other decimal
/// format families. Scientific rejects this value, but re-exporting it keeps
/// family callers on one Numbers format vocabulary.
pub use core::NATIVE_AUTOMATIC_DECIMAL_PLACES;

/// Largest fixed decimal-place value accepted by Numbers' native scientific
/// format.
pub const MAX_SCIENTIFIC_DECIMAL_PLACES: u32 = core::MAX_NUMBER_DECIMAL_PLACES;

/// Native negative-number style used by scientific formats.
pub use core::NATIVE_SCIENTIFIC_NEGATIVE_STYLE;

/// Scientific formats never display a thousands separator.
pub use core::NATIVE_SCIENTIFIC_SHOW_THOUSANDS_SEPARATOR;

/// Borrowed scalar values for one strict native Scientific format payload.
///
/// The complete source payload remains available through [`Self::raw`].
/// Unknown extension fields and groups are never decoded into owned storage
/// and are copied byte-for-byte by [`rewrite_scientific_format`].
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ScientificFormatSnapshot<'source>(core::NumberFormatSnapshot<'source>);

impl<'source> ScientificFormatSnapshot<'source> {
    /// Borrow the original wire payload.
    #[must_use]
    pub const fn raw(self) -> &'source [u8] {
        self.0.raw()
    }

    /// Return the native Scientific discriminator (`259`).
    #[must_use]
    pub const fn format_type(self) -> u32 {
        self.0.format_type()
    }

    /// Return the fixed decimal-place count in `0..=30`.
    #[must_use]
    pub const fn decimal_places(self) -> u32 {
        self.0.decimal_places()
    }

    /// Return the native minus-sign style (`0`).
    #[must_use]
    pub const fn negative_style(self) -> u32 {
        self.0.negative_style()
    }

    /// Return the native thousands-separator setting (`false`).
    #[must_use]
    pub const fn show_thousands_separator(self) -> bool {
        self.0.show_thousands_separator()
    }
}

/// Owned scalar values accepted by the native Scientific format writer.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ScientificFormatWrite(core::NumberFormatWrite);

impl ScientificFormatWrite {
    /// Construct a native Scientific format update with fixed precision.
    #[must_use]
    pub const fn new(decimal_places: u32) -> Self {
        Self(core::NumberFormatWrite::new(
            decimal_places,
            NATIVE_SCIENTIFIC_NEGATIVE_STYLE,
            NATIVE_SCIENTIFIC_SHOW_THOUSANDS_SEPARATOR,
        ))
    }

    /// Copy the semantic values from a decoded Scientific snapshot.
    #[must_use]
    pub const fn from_snapshot(snapshot: ScientificFormatSnapshot<'_>) -> Self {
        Self(core::NumberFormatWrite::from_snapshot(snapshot.0))
    }

    /// Return the fixed decimal-place count.
    #[must_use]
    pub const fn decimal_places(self) -> u32 {
        self.0.decimal_places()
    }

    /// Replace the fixed decimal-place count.
    #[must_use]
    pub const fn with_decimal_places(self, value: u32) -> Self {
        Self(self.0.with_decimal_places(value))
    }
}

/// Prepared source-preserving native Scientific rewrite.
#[derive(Debug, Clone, Copy)]
pub struct PreparedScientificFormatRewrite<'source>(core::PreparedNumberFormatRewrite<'source>);

impl PreparedScientificFormatRewrite<'_> {
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

/// Prepared canonical native Scientific append.
#[derive(Debug, Clone, Copy)]
pub struct PreparedScientificFormatWrite(core::PreparedNumberFormatWrite);

impl PreparedScientificFormatWrite {
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

    /// Emit and strictly read back a canonical Scientific payload.
    pub fn execute(self, limits: RewriteExecutionLimits) -> Result<RewriteOutput, DecodeError> {
        self.0.execute(limits)
    }
}

/// Strictly decode one native Scientific `FormatStructArchive`.
pub fn decode_scientific_format(
    source: &[u8],
    options: DecodeOptions,
) -> Result<ScientificFormatSnapshot<'_>, DecodeError> {
    core::decode_decimal_format(source, NATIVE_SCIENTIFIC_FORMAT_TYPE, options)
        .and_then(wrap_snapshot)
}

/// Strictly decode one native Scientific format and return measured wire use.
pub fn decode_scientific_format_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(ScientificFormatSnapshot<'_>, DecodeReport), DecodeError> {
    core::decode_decimal_format_with_report(source, NATIVE_SCIENTIFIC_FORMAT_TYPE, options)
        .and_then(|(snapshot, report)| wrap_snapshot(snapshot).map(|snapshot| (snapshot, report)))
}

/// Prepare a source-preserving native Scientific format rewrite.
pub fn prepare_scientific_format_rewrite<'source>(
    source: &'source [u8],
    write: ScientificFormatWrite,
    options: DecodeOptions,
) -> Result<PreparedScientificFormatRewrite<'source>, DecodeError> {
    core::prepare_decimal_format_rewrite(source, write.0, NATIVE_SCIENTIFIC_FORMAT_TYPE, options)
        .map(PreparedScientificFormatRewrite)
}

/// Rewrite one native Scientific format while preserving unknown source
/// fields byte-for-byte.
pub fn rewrite_scientific_format(
    source: &[u8],
    write: ScientificFormatWrite,
    options: DecodeOptions,
) -> Result<RewriteOutput, DecodeError> {
    core::rewrite_decimal_format(source, write.0, NATIVE_SCIENTIFIC_FORMAT_TYPE, options)
}

/// Compatibility spelling for the table-cell Scientific route.
pub use rewrite_scientific_format as rewrite_table_cell_scientific_format;

/// Prepare a canonical native Scientific format payload for a new list entry.
pub fn prepare_scientific_format_write(
    write: ScientificFormatWrite,
    options: DecodeOptions,
) -> Result<PreparedScientificFormatWrite, DecodeError> {
    core::prepare_decimal_format_write(write.0, NATIVE_SCIENTIFIC_FORMAT_TYPE, options)
        .map(PreparedScientificFormatWrite)
}

/// Prepare a canonical Scientific format append under the explicit append
/// spelling.
pub use prepare_scientific_format_write as prepare_scientific_format_append;

/// Encode a canonical native Scientific format payload for a new list entry.
pub fn canonical_scientific_format(
    write: ScientificFormatWrite,
    options: DecodeOptions,
) -> Result<RewriteOutput, DecodeError> {
    core::canonical_decimal_format(write.0, NATIVE_SCIENTIFIC_FORMAT_TYPE, options)
}

fn wrap_snapshot<'source>(
    snapshot: core::NumberFormatSnapshot<'source>,
) -> Result<ScientificFormatSnapshot<'source>, DecodeError> {
    if snapshot.format_type() != NATIVE_SCIENTIFIC_FORMAT_TYPE
        || snapshot.decimal_places() == NATIVE_AUTOMATIC_DECIMAL_PLACES
        || snapshot.decimal_places() > MAX_SCIENTIFIC_DECIMAL_PLACES
        || snapshot.negative_style() != NATIVE_SCIENTIFIC_NEGATIVE_STYLE
        || snapshot.show_thousands_separator() != NATIVE_SCIENTIFIC_SHOW_THOUSANDS_SEPARATOR
    {
        return Err(DecodeError::invalid());
    }
    Ok(ScientificFormatSnapshot(snapshot))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn options(_source: &[u8]) -> DecodeOptions {
        DecodeOptions::new(16 * 1024, 16 * 1024, 16 * 1024, 64 * 1024, 64, 16, 64, 4096)
    }

    fn native_scientific_format(decimal_places: u32, negative_style: u32, show: bool) -> Vec<u8> {
        let mut source = Vec::new();
        source.extend_from_slice(&[0x08, 0x83, 0x02]);
        source.extend_from_slice(&[0x10]);
        encode_varint(&mut source, u64::from(decimal_places));
        source.extend_from_slice(&[0x20]);
        encode_varint(&mut source, u64::from(negative_style));
        source.extend_from_slice(&[0x28]);
        encode_varint(&mut source, u64::from(show));
        source
    }

    fn encode_varint(output: &mut Vec<u8>, mut value: u64) {
        while value >= 0x80 {
            output.push((value as u8 & 0x7f) | 0x80);
            value >>= 7;
        }
        output.push(value as u8);
    }

    #[test]
    fn scientific_format_requires_canonical_scalar_domain() {
        for decimal_places in [0, MAX_SCIENTIFIC_DECIMAL_PLACES] {
            let source = native_scientific_format(decimal_places, 0, false);
            let snapshot = decode_scientific_format(&source, options(&source)).expect("valid");
            assert_eq!(snapshot.format_type(), NATIVE_SCIENTIFIC_FORMAT_TYPE);
            assert_eq!(snapshot.decimal_places(), decimal_places);
            assert_eq!(snapshot.negative_style(), NATIVE_SCIENTIFIC_NEGATIVE_STYLE);
            assert_eq!(
                snapshot.show_thousands_separator(),
                NATIVE_SCIENTIFIC_SHOW_THOUSANDS_SEPARATOR
            );
        }
        for decimal_places in [NATIVE_AUTOMATIC_DECIMAL_PLACES, 31, 254] {
            let source = native_scientific_format(decimal_places, 0, false);
            assert!(decode_scientific_format(&source, options(&source)).is_err());
        }
        let source = native_scientific_format(2, 1, false);
        assert!(decode_scientific_format(&source, options(&source)).is_err());
        let source = native_scientific_format(2, 0, true);
        assert!(decode_scientific_format(&source, options(&source)).is_err());
    }

    #[test]
    fn scientific_format_requires_exact_fields_and_known_wire() {
        let fields = [
            &[0x08, 0x83, 0x02][..],
            &[0x10, 0x02][..],
            &[0x20, 0x00][..],
            &[0x28, 0x00][..],
        ];
        for omitted in 0..fields.len() {
            let mut source = Vec::new();
            for (index, field) in fields.iter().enumerate() {
                if index != omitted {
                    source.extend_from_slice(field);
                }
            }
            assert!(decode_scientific_format(&source, options(&source)).is_err());
        }

        let mut duplicate = native_scientific_format(2, 0, false);
        duplicate.extend_from_slice(&[0x08, 0x83, 0x02]);
        assert!(decode_scientific_format(&duplicate, options(&duplicate)).is_err());

        let wrong_family = [0x08, 0x80, 0x02, 0x10, 0x02, 0x20, 0x00, 0x28, 0x00];
        assert!(decode_scientific_format(&wrong_family, options(&wrong_family)).is_err());

        let wrong_wire = [0x0a, 0x01, 0x00, 0x10, 0x02, 0x20, 0x00, 0x28, 0x00];
        assert!(decode_scientific_format(&wrong_wire, options(&wrong_wire)).is_err());

        let noncanonical = [0x08, 0x83, 0x02, 0x10, 0x82, 0x00, 0x20, 0x00, 0x28, 0x00];
        assert!(decode_scientific_format(&noncanonical, options(&noncanonical)).is_err());

        for incompatible in [[0x1a, 0x00], [0x30, 0x01], [0x72, 0x00]] {
            let mut source = native_scientific_format(2, 0, false);
            source.extend_from_slice(&incompatible);
            assert!(decode_scientific_format(&source, options(&source)).is_err());
        }
    }

    #[test]
    fn scientific_format_rewrite_preserves_unknown_source_and_reports_bounds() {
        let mut source = native_scientific_format(2, 0, false);
        let unknown = [
            0xa0, 0x06, 0x81, 0x00, // unknown scalar 100, overlong value
            0xa3, 0x06, 0xa8, 0x06, 0x01, 0xa4, 0x06, // unknown balanced group
        ];
        source.extend_from_slice(&unknown);
        let write = ScientificFormatWrite::new(MAX_SCIENTIFIC_DECIMAL_PLACES);
        let (snapshot, report) =
            decode_scientific_format_with_report(&source, options(&source)).expect("decode");
        assert_eq!(snapshot.raw(), source.as_slice());
        assert_eq!(report.input_bytes(), source.len());
        assert_eq!(report.work_bytes(), source.len() * 2);
        assert_eq!(report.allocations(), 0);

        let prepared = prepare_scientific_format_rewrite(&source, write, options(&source))
            .expect("prepare rewrite");
        let requirements = prepared.execution_requirements();
        let output = prepared
            .execute(RewriteExecutionLimits::exact(requirements))
            .expect("execute rewrite");
        assert_eq!(output.bytes().len(), requirements.output_bytes());
        assert_eq!(output.report().fields(), requirements.fields());
        assert_eq!(output.report().work_bytes(), requirements.work_bytes());
        assert!(output.bytes().ends_with(&unknown));
        let rewritten = decode_scientific_format(output.bytes(), options(output.bytes()))
            .expect("rewritten scientific");
        assert_eq!(ScientificFormatWrite::from_snapshot(rewritten), write);

        let no_op = rewrite_scientific_format(
            output.bytes(),
            ScientificFormatWrite::from_snapshot(rewritten),
            options(output.bytes()),
        )
        .expect("no-op rewrite");
        assert_eq!(no_op.bytes(), output.bytes());
        assert!(matches!(
            prepare_scientific_format_rewrite(&source, write, options(&source))
                .expect("prepare limits")
                .execute(
                    RewriteExecutionLimits::exact(requirements)
                        .with_output_bytes(requirements.output_bytes() - 1),
                )
                .expect_err("output ceiling")
                .resource_limit(),
            Some(DecodeLimit::OutputBytes { .. })
        ));
        assert!(matches!(
            prepare_scientific_format_rewrite(&source, write, options(&source))
                .expect("prepare limits")
                .execute(
                    RewriteExecutionLimits::exact(requirements)
                        .with_work_bytes(requirements.work_bytes() - 1),
                )
                .expect_err("work ceiling")
                .resource_limit(),
            Some(DecodeLimit::Work { .. })
        ));
    }

    #[test]
    fn scientific_format_canonical_append_has_exact_wire() {
        let write = ScientificFormatWrite::new(5);
        let prepared = prepare_scientific_format_append(write, options(&[])).expect("prepare");
        let requirements = prepared.execution_requirements();
        let output = prepared
            .execute(RewriteExecutionLimits::exact(requirements))
            .expect("append");
        assert_eq!(
            output.bytes(),
            &[0x08, 0x83, 0x02, 0x10, 0x05, 0x20, 0x00, 0x28, 0x00]
        );
        assert_eq!(requirements.output_bytes(), output.bytes().len());
        assert_eq!(requirements.fields(), 4);
        assert_eq!(requirements.work_bytes(), output.bytes().len() * 3);
        let snapshot = decode_scientific_format(output.bytes(), options(output.bytes()))
            .expect("decode append");
        assert_eq!(ScientificFormatWrite::from_snapshot(snapshot), write);
        assert_eq!(
            canonical_scientific_format(write, options(&[]))
                .expect("one-shot")
                .bytes(),
            output.bytes()
        );
    }
}
