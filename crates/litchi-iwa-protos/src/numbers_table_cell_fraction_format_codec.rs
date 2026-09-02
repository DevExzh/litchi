//! Neutral strict native-fraction format seam for Numbers table-cell adapters.
//!
//! Fraction payloads use the shared source-preserving `FormatStructArchive`
//! wire core while keeping their nine native denominator strategies behind a
//! nominal family boundary. Ordinary Fraction owns fields 1 and 11; a legacy
//! field-20 (`requires_fraction_replacement`) marker is accepted only when it
//! is canonical `false`, preserved byte-for-byte, and never synthesized.
//! `true` is rejected because this seam cannot safely implement replacement
//! semantics.
//! Generated Buffa values never cross this module.

#![allow(clippy::module_name_repetitions)]

use crate::numbers_table_cell_pop_up_menu_codec as core;

pub use core::{
    DecodeError, DecodeLimit, DecodeOptions, DecodeReport, RewriteExecutionLimits,
    RewriteExecutionRequirements, RewriteOutput,
};

/// Native Numbers display-format discriminator for a fraction cell.
pub use core::NATIVE_FRACTION_FORMAT_TYPE;

/// Native fraction accuracy for a denominator with at most one digit.
pub use core::NATIVE_FRACTION_UP_TO_ONE_DIGIT;

/// Native fraction accuracy for a denominator with at most two digits.
pub use core::NATIVE_FRACTION_UP_TO_TWO_DIGITS;

/// Native fraction accuracy for a denominator with at most three digits.
pub use core::NATIVE_FRACTION_UP_TO_THREE_DIGITS;

/// Native fraction accuracy for halves.
pub use core::NATIVE_FRACTION_HALVES;

/// Native fraction accuracy for quarters.
pub use core::NATIVE_FRACTION_QUARTERS;

/// Native fraction accuracy for eighths.
pub use core::NATIVE_FRACTION_EIGHTHS;

/// Native fraction accuracy for sixteenths.
pub use core::NATIVE_FRACTION_SIXTEENTHS;

/// Native fraction accuracy for tenths.
pub use core::NATIVE_FRACTION_TENTHS;

/// Native fraction accuracy for hundredths.
pub use core::NATIVE_FRACTION_HUNDREDTHS;

/// The nine denominator strategies accepted by native Numbers Fraction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u32)]
pub enum FractionAccuracy {
    /// Use a denominator with at most one digit.
    UpToOneDigit = NATIVE_FRACTION_UP_TO_ONE_DIGIT,
    /// Use a denominator with at most two digits.
    UpToTwoDigits = NATIVE_FRACTION_UP_TO_TWO_DIGITS,
    /// Use a denominator with at most three digits.
    UpToThreeDigits = NATIVE_FRACTION_UP_TO_THREE_DIGITS,
    /// Always use halves.
    Halves = NATIVE_FRACTION_HALVES,
    /// Always use quarters.
    Quarters = NATIVE_FRACTION_QUARTERS,
    /// Always use eighths.
    Eighths = NATIVE_FRACTION_EIGHTHS,
    /// Always use sixteenths.
    Sixteenths = NATIVE_FRACTION_SIXTEENTHS,
    /// Always use tenths.
    Tenths = NATIVE_FRACTION_TENTHS,
    /// Always use hundredths.
    Hundredths = NATIVE_FRACTION_HUNDREDTHS,
}

impl FractionAccuracy {
    /// Convert a native wire value to one of the nine supported strategies.
    #[must_use]
    pub const fn from_native(value: u32) -> Option<Self> {
        match value {
            NATIVE_FRACTION_UP_TO_ONE_DIGIT => Some(Self::UpToOneDigit),
            NATIVE_FRACTION_UP_TO_TWO_DIGITS => Some(Self::UpToTwoDigits),
            NATIVE_FRACTION_UP_TO_THREE_DIGITS => Some(Self::UpToThreeDigits),
            NATIVE_FRACTION_HALVES => Some(Self::Halves),
            NATIVE_FRACTION_QUARTERS => Some(Self::Quarters),
            NATIVE_FRACTION_EIGHTHS => Some(Self::Eighths),
            NATIVE_FRACTION_SIXTEENTHS => Some(Self::Sixteenths),
            NATIVE_FRACTION_TENTHS => Some(Self::Tenths),
            NATIVE_FRACTION_HUNDREDTHS => Some(Self::Hundredths),
            _ => None,
        }
    }

    /// Return the exact unsigned value stored in field 11.
    #[must_use]
    pub const fn native_value(self) -> u32 {
        self as u32
    }
}

/// Borrowed scalar values for one strict native Fraction format payload.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct FractionFormatSnapshot<'source>(core::FractionFormatSnapshot<'source>);

impl<'source> FractionFormatSnapshot<'source> {
    /// Borrow the original wire payload.
    #[must_use]
    pub const fn raw(self) -> &'source [u8] {
        self.0.raw()
    }

    /// Return the native Fraction discriminator (`262`).
    #[must_use]
    pub const fn format_type(self) -> u32 {
        self.0.format_type()
    }

    /// Return the exact native field-11 denominator strategy value.
    #[must_use]
    pub const fn fraction_accuracy(self) -> u32 {
        self.0.fraction_accuracy()
    }

    /// Return the typed denominator strategy.
    #[must_use]
    pub const fn accuracy(self) -> Option<FractionAccuracy> {
        FractionAccuracy::from_native(self.fraction_accuracy())
    }

    /// Return the optional legacy field-20 marker. `Some(false)` is accepted
    /// and preserved from source; `Some(true)` is rejected during decode.
    #[must_use]
    pub const fn requires_fraction_replacement(self) -> Option<bool> {
        self.0.requires_fraction_replacement()
    }
}

/// Scalar values accepted by the strict native Fraction writer.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct FractionFormatWrite(core::FractionFormatWrite);

impl FractionFormatWrite {
    /// Construct a native Fraction format update from its field-11 value.
    #[must_use]
    pub const fn new(fraction_accuracy: u32) -> Self {
        Self(core::FractionFormatWrite::new(fraction_accuracy))
    }

    /// Construct a native Fraction format update from a typed strategy.
    #[must_use]
    pub const fn from_accuracy(accuracy: FractionAccuracy) -> Self {
        Self::new(accuracy.native_value())
    }

    /// Copy the semantic values from a decoded Fraction snapshot.
    #[must_use]
    pub const fn from_snapshot(snapshot: FractionFormatSnapshot<'_>) -> Self {
        Self(core::FractionFormatWrite::from_snapshot(snapshot.0))
    }

    /// Return the exact native field-11 denominator strategy value.
    #[must_use]
    pub const fn fraction_accuracy(self) -> u32 {
        self.0.fraction_accuracy()
    }

    /// Return the typed denominator strategy.
    #[must_use]
    pub const fn accuracy(self) -> Option<FractionAccuracy> {
        FractionAccuracy::from_native(self.fraction_accuracy())
    }

    /// Replace the native field-11 denominator strategy value.
    #[must_use]
    pub const fn with_fraction_accuracy(self, value: u32) -> Self {
        Self(self.0.with_fraction_accuracy(value))
    }

    /// Replace the denominator strategy with a typed value.
    #[must_use]
    pub const fn with_accuracy(self, value: FractionAccuracy) -> Self {
        self.with_fraction_accuracy(value.native_value())
    }
}

/// Prepared source-preserving native Fraction rewrite.
#[derive(Debug, Clone, Copy)]
pub struct PreparedFractionFormatRewrite<'source>(core::PreparedFractionFormatRewrite<'source>);

impl PreparedFractionFormatRewrite<'_> {
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

/// Prepared canonical native Fraction append.
#[derive(Debug, Clone, Copy)]
pub struct PreparedFractionFormatWrite(core::PreparedFractionFormatWrite);

impl PreparedFractionFormatWrite {
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

    /// Emit and strictly read back a canonical Fraction payload.
    pub fn execute(self, limits: RewriteExecutionLimits) -> Result<RewriteOutput, DecodeError> {
        self.0.execute(limits)
    }
}

/// Strictly decode one native Fraction `FormatStructArchive`.
pub fn decode_fraction_format(
    source: &[u8],
    options: DecodeOptions,
) -> Result<FractionFormatSnapshot<'_>, DecodeError> {
    core::decode_fraction_format(source, options).map(FractionFormatSnapshot)
}

/// Strictly decode one native Fraction format and return measured wire use.
pub fn decode_fraction_format_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(FractionFormatSnapshot<'_>, DecodeReport), DecodeError> {
    core::decode_fraction_format_with_report(source, options)
        .map(|(snapshot, report)| (FractionFormatSnapshot(snapshot), report))
}

/// Prepare a source-preserving native Fraction format rewrite.
pub fn prepare_fraction_format_rewrite<'source>(
    source: &'source [u8],
    write: FractionFormatWrite,
    options: DecodeOptions,
) -> Result<PreparedFractionFormatRewrite<'source>, DecodeError> {
    core::prepare_fraction_format_rewrite(source, write.0, options)
        .map(PreparedFractionFormatRewrite)
}

/// Rewrite one native Fraction format while preserving unknown source fields.
pub fn rewrite_fraction_format(
    source: &[u8],
    write: FractionFormatWrite,
    options: DecodeOptions,
) -> Result<RewriteOutput, DecodeError> {
    core::rewrite_fraction_format(source, write.0, options)
}

/// Compatibility spelling for the table-cell Fraction route.
pub use rewrite_fraction_format as rewrite_table_cell_fraction_format;

/// Prepare a canonical native Fraction format payload for a new list entry.
pub fn prepare_fraction_format_write(
    write: FractionFormatWrite,
    options: DecodeOptions,
) -> Result<PreparedFractionFormatWrite, DecodeError> {
    core::prepare_fraction_format_write(write.0, options).map(PreparedFractionFormatWrite)
}

/// Prepare a canonical Fraction append under the explicit append spelling.
pub use prepare_fraction_format_write as prepare_fraction_format_append;

/// Encode a canonical native Fraction format payload for a new list entry.
pub fn canonical_fraction_format(
    write: FractionFormatWrite,
    options: DecodeOptions,
) -> Result<RewriteOutput, DecodeError> {
    core::canonical_fraction_format(write.0, options)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn options() -> DecodeOptions {
        DecodeOptions::new(16 * 1024, 16 * 1024, 64, 64 * 1024, 64, 0, 0, 0)
    }

    fn field(number: u32, value: u32) -> Vec<u8> {
        fn push_varint(output: &mut Vec<u8>, mut value: u64) {
            while value >= 0x80 {
                output.push((value as u8 & 0x7f) | 0x80);
                value >>= 7;
            }
            output.push(value as u8);
        }

        let mut output = Vec::new();
        push_varint(&mut output, u64::from(number) << 3);
        push_varint(&mut output, u64::from(value));
        output
    }

    fn native_fraction(accuracy: u32) -> Vec<u8> {
        let mut output = field(1, NATIVE_FRACTION_FORMAT_TYPE);
        output.extend_from_slice(&field(11, accuracy));
        output
    }

    #[test]
    fn fraction_format_accepts_all_nine_native_accuracies() {
        let accuracies = [
            (
                FractionAccuracy::UpToOneDigit,
                NATIVE_FRACTION_UP_TO_ONE_DIGIT,
            ),
            (
                FractionAccuracy::UpToTwoDigits,
                NATIVE_FRACTION_UP_TO_TWO_DIGITS,
            ),
            (
                FractionAccuracy::UpToThreeDigits,
                NATIVE_FRACTION_UP_TO_THREE_DIGITS,
            ),
            (FractionAccuracy::Halves, NATIVE_FRACTION_HALVES),
            (FractionAccuracy::Quarters, NATIVE_FRACTION_QUARTERS),
            (FractionAccuracy::Eighths, NATIVE_FRACTION_EIGHTHS),
            (FractionAccuracy::Sixteenths, NATIVE_FRACTION_SIXTEENTHS),
            (FractionAccuracy::Tenths, NATIVE_FRACTION_TENTHS),
            (FractionAccuracy::Hundredths, NATIVE_FRACTION_HUNDREDTHS),
        ];
        for (accuracy, native) in accuracies {
            let source = native_fraction(native);
            let snapshot = decode_fraction_format(&source, options()).expect("fraction");
            assert_eq!(snapshot.format_type(), NATIVE_FRACTION_FORMAT_TYPE);
            assert_eq!(snapshot.fraction_accuracy(), native);
            assert_eq!(snapshot.accuracy(), Some(accuracy));
            assert_eq!(snapshot.requires_fraction_replacement(), None);
        }
    }

    #[test]
    fn fraction_format_rejects_incompatible_fields_and_domains() {
        let valid = native_fraction(NATIVE_FRACTION_UP_TO_THREE_DIGITS);
        for extra in [
            field(2, 2),
            field(4, 0),
            field(5, 0),
            field(6, 0),
            field(7, 0),
            field(8, 10),
            field(9, 2),
            field(10, 1),
            field(20, 1),
            field(21, 0),
            field(45, 0),
        ] {
            let mut source = valid.clone();
            source.extend_from_slice(&extra);
            assert!(decode_fraction_format(&source, options()).is_err());
        }
        assert!(decode_fraction_format(&native_fraction(0), options()).is_err());
        assert!(decode_fraction_format(&native_fraction(3), options()).is_err());
        let mut duplicate = valid.clone();
        duplicate.extend_from_slice(&field(11, NATIVE_FRACTION_HALVES));
        assert!(decode_fraction_format(&duplicate, options()).is_err());
        let mut legacy_false = valid.clone();
        legacy_false.extend_from_slice(&field(20, 0));
        let snapshot = decode_fraction_format(&legacy_false, options()).expect("false marker");
        assert_eq!(snapshot.requires_fraction_replacement(), Some(false));
        let wrong_type = field(1, 261);
        let mut wrong_family = wrong_type;
        wrong_family.extend_from_slice(&field(11, NATIVE_FRACTION_HALVES));
        assert!(decode_fraction_format(&wrong_family, options()).is_err());
    }

    #[test]
    fn fraction_format_rewrite_preserves_unknown_spans_and_prepared_accounting() {
        let mut source = native_fraction(NATIVE_FRACTION_UP_TO_THREE_DIGITS);
        let unknown = [
            0xa0, 0x06, 0x81, 0x00, // unknown scalar with overlong value spelling
            0xa3, 0x06, 0xa8, 0x06, 0x01, 0xa4, 0x06, // balanced unknown group
        ];
        source.splice(3..3, unknown);
        source.extend_from_slice(&field(20, 0));
        let write = FractionFormatWrite::new(NATIVE_FRACTION_HUNDREDTHS);
        let prepared = prepare_fraction_format_rewrite(&source, write, options()).expect("prepare");
        let requirements = prepared.execution_requirements();
        let output = prepared
            .execute(RewriteExecutionLimits::exact(requirements))
            .expect("execute");
        assert!(
            output
                .bytes()
                .windows(unknown.len())
                .any(|window| window == unknown)
        );
        assert!(output.bytes().ends_with(&field(20, 0)));
        assert_eq!(output.report().fields(), requirements.fields());
        assert_eq!(output.report().work_bytes(), requirements.work_bytes());
        let snapshot = decode_fraction_format(output.bytes(), options()).expect("readback");
        assert_eq!(FractionFormatWrite::from_snapshot(snapshot), write);
        let (_, source_report) =
            decode_fraction_format_with_report(&source, options()).expect("source report");
        let (_, candidate_report) = decode_fraction_format_with_report(output.bytes(), options())
            .expect("candidate report");
        assert_eq!(
            requirements.work_bytes(),
            source_report.work_bytes() + output.bytes().len() + candidate_report.work_bytes()
        );
    }

    #[test]
    fn fraction_format_canonical_writer_is_exact_and_rejects_bad_writes() {
        let write = FractionFormatWrite::from_accuracy(FractionAccuracy::Halves);
        let prepared = prepare_fraction_format_write(write, options()).expect("prepare");
        let requirements = prepared.execution_requirements();
        let output = prepared
            .execute(RewriteExecutionLimits::exact(requirements))
            .expect("execute");
        assert_eq!(output.bytes(), native_fraction(NATIVE_FRACTION_HALVES));
        assert_eq!(requirements.fields(), 2);
        assert_eq!(requirements.work_bytes(), output.bytes().len() * 3);
        assert!(prepare_fraction_format_write(FractionFormatWrite::new(3), options()).is_err());
        assert!(canonical_fraction_format(write, options()).is_ok());
    }
}
