//! Neutral strict native date-and-time format seam for Numbers table-cell
//! adapters.
//!
//! DateTime uses the shared source-preserving `FormatStructArchive` wire core
//! while keeping its pattern behind a nominal family boundary. The complete
//! source payload remains authoritative; generated Buffa values never cross
//! this module.

#![allow(clippy::module_name_repetitions)]

use crate::numbers_table_cell_pop_up_menu_codec as core;

pub use core::{
    DecodeError, DecodeLimit, DecodeOptions, DecodeReport, MAX_DATE_TIME_PATTERN_BYTES,
    RewriteExecutionLimits, RewriteExecutionRequirements, RewriteOutput,
};

/// Native Numbers display-format discriminator for a date-and-time cell.
pub use core::NATIVE_DATE_TIME_FORMAT_TYPE;

/// Borrowed scalar values for one strict native DateTime format payload.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DateTimeFormatSnapshot<'source>(core::DateTimeFormatSnapshot<'source>);

impl<'source> DateTimeFormatSnapshot<'source> {
    /// Borrow the original wire payload.
    #[must_use]
    pub const fn raw(self) -> &'source [u8] {
        self.0.raw()
    }

    /// Return the native DateTime discriminator (`261`).
    #[must_use]
    pub const fn format_type(self) -> u32 {
        self.0.format_type()
    }

    /// Return the exact UTF-8 date-and-time pattern borrowed from the source.
    #[must_use]
    pub const fn date_time_format(self) -> &'source str {
        self.0.date_time_format()
    }

    /// Return the exact pattern borrowed from the source.
    #[must_use]
    pub const fn pattern(self) -> &'source str {
        self.date_time_format()
    }
}

/// Scalar values accepted by the strict native DateTime writer.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DateTimeFormatWrite<'source>(core::DateTimeFormatWrite<'source>);

impl<'source> DateTimeFormatWrite<'source> {
    /// Construct a native DateTime format update from a pattern.
    #[must_use]
    pub const fn new(pattern: &'source str) -> Self {
        Self(core::DateTimeFormatWrite::new(pattern))
    }

    /// Copy the semantic pattern from a decoded DateTime snapshot.
    #[must_use]
    pub const fn from_snapshot(snapshot: DateTimeFormatSnapshot<'source>) -> Self {
        Self(core::DateTimeFormatWrite::from_snapshot(snapshot.0))
    }

    /// Return the requested date-and-time pattern.
    #[must_use]
    pub const fn date_time_format(self) -> &'source str {
        self.0.date_time_format()
    }

    /// Return the requested pattern.
    #[must_use]
    pub const fn pattern(self) -> &'source str {
        self.date_time_format()
    }

    /// Replace the requested pattern.
    #[must_use]
    pub const fn with_pattern(self, pattern: &'source str) -> Self {
        Self::new(pattern)
    }

    /// Replace the requested date-and-time pattern.
    #[must_use]
    pub const fn with_date_time_format(self, pattern: &'source str) -> Self {
        self.with_pattern(pattern)
    }
}

/// Prepared source-preserving native DateTime rewrite.
#[derive(Debug, Clone, Copy)]
pub struct PreparedDateTimeFormatRewrite<'source>(core::PreparedDateTimeFormatRewrite<'source>);

impl PreparedDateTimeFormatRewrite<'_> {
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

/// Prepared canonical native DateTime append.
#[derive(Debug, Clone, Copy)]
pub struct PreparedDateTimeFormatWrite<'source>(core::PreparedDateTimeFormatWrite<'source>);

impl PreparedDateTimeFormatWrite<'_> {
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

    /// Emit and strictly read back a canonical DateTime payload.
    pub fn execute(self, limits: RewriteExecutionLimits) -> Result<RewriteOutput, DecodeError> {
        self.0.execute(limits)
    }
}

/// Strictly decode one native DateTime `FormatStructArchive`.
pub fn decode_date_time_format(
    source: &[u8],
    options: DecodeOptions,
) -> Result<DateTimeFormatSnapshot<'_>, DecodeError> {
    core::decode_date_time_format(source, options).map(DateTimeFormatSnapshot)
}

/// Strictly decode one native DateTime format and return measured wire use.
pub fn decode_date_time_format_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(DateTimeFormatSnapshot<'_>, DecodeReport), DecodeError> {
    core::decode_date_time_format_with_report(source, options)
        .map(|(snapshot, report)| (DateTimeFormatSnapshot(snapshot), report))
}

/// Prepare a source-preserving native DateTime format rewrite.
pub fn prepare_date_time_format_rewrite<'source>(
    source: &'source [u8],
    write: DateTimeFormatWrite<'source>,
    options: DecodeOptions,
) -> Result<PreparedDateTimeFormatRewrite<'source>, DecodeError> {
    core::prepare_date_time_format_rewrite(source, write.0, options)
        .map(PreparedDateTimeFormatRewrite)
}

/// Rewrite one native DateTime format while preserving unknown source fields
/// byte-for-byte.
pub fn rewrite_date_time_format(
    source: &[u8],
    write: DateTimeFormatWrite<'_>,
    options: DecodeOptions,
) -> Result<RewriteOutput, DecodeError> {
    core::rewrite_date_time_format(source, write.0, options)
}

/// Compatibility spelling for the table-cell DateTime route.
pub use rewrite_date_time_format as rewrite_table_cell_date_time_format;

/// Prepare a canonical native DateTime format payload for a new list entry.
pub fn prepare_date_time_format_write<'source>(
    write: DateTimeFormatWrite<'source>,
    options: DecodeOptions,
) -> Result<PreparedDateTimeFormatWrite<'source>, DecodeError> {
    core::prepare_date_time_format_write(write.0, options).map(PreparedDateTimeFormatWrite)
}

/// Prepare a canonical DateTime append under the explicit append spelling.
pub use prepare_date_time_format_write as prepare_date_time_format_append;

/// Encode a canonical native DateTime format payload for a new list entry.
pub fn canonical_date_time_format(
    write: DateTimeFormatWrite<'_>,
    options: DecodeOptions,
) -> Result<RewriteOutput, DecodeError> {
    core::canonical_date_time_format(write.0, options)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn options() -> DecodeOptions {
        DecodeOptions::new(16 * 1024, 16 * 1024, 16 * 1024, 64 * 1024, 64, 0, 0, 4096)
    }

    fn push_varint(output: &mut Vec<u8>, mut value: u64) {
        while value >= 0x80 {
            output.push((value as u8 & 0x7f) | 0x80);
            value >>= 7;
        }
        output.push(value as u8);
    }

    fn varint_field(number: u32, value: u64) -> Vec<u8> {
        let mut output = Vec::new();
        push_varint(&mut output, u64::from(number) << 3);
        push_varint(&mut output, value);
        output
    }

    fn length_field(number: u32, value: &[u8]) -> Vec<u8> {
        let mut output = Vec::new();
        push_varint(&mut output, (u64::from(number) << 3) | 2);
        push_varint(
            &mut output,
            u64::try_from(value.len()).expect("test pattern length fits"),
        );
        output.extend_from_slice(value);
        output
    }

    fn group_field(number: u32, body: &[u8]) -> Vec<u8> {
        let mut output = Vec::new();
        push_varint(&mut output, (u64::from(number) << 3) | 3);
        output.extend_from_slice(body);
        push_varint(&mut output, (u64::from(number) << 3) | 4);
        output
    }

    fn native(pattern: &str) -> Vec<u8> {
        let mut output = varint_field(1, u64::from(NATIVE_DATE_TIME_FORMAT_TYPE));
        output.extend_from_slice(&length_field(14, pattern.as_bytes()));
        output
    }

    #[test]
    fn date_time_format_accepts_native_discriminator_and_borrows_pattern() {
        let source = native("yyyy-MM-dd H:mm:ss");
        let snapshot = decode_date_time_format(&source, options()).expect("date time");
        assert_eq!(snapshot.raw(), source.as_slice());
        assert_eq!(snapshot.format_type(), NATIVE_DATE_TIME_FORMAT_TYPE);
        assert_eq!(snapshot.pattern(), "yyyy-MM-dd H:mm:ss");
        assert_eq!(
            DateTimeFormatWrite::from_snapshot(snapshot).pattern(),
            snapshot.pattern()
        );
        assert!(
            snapshot
                .raw()
                .as_ptr_range()
                .contains(&snapshot.pattern().as_ptr())
        );
    }

    #[test]
    fn date_time_format_rejects_known_cross_family_fields_and_bad_patterns() {
        let valid = native("yyyy-MM-dd");
        for extra in [
            varint_field(2, 0),
            varint_field(4, 0),
            length_field(3, b"USD"),
            varint_field(20, 0),
            varint_field(45, 0),
        ] {
            let mut source = valid.clone();
            source.extend_from_slice(&extra);
            assert!(decode_date_time_format(&source, options()).is_err());
        }
        for pattern in ["", "   ", "yyyy\0MM", "yyyy\nMM"] {
            assert!(decode_date_time_format(&native(pattern), options()).is_err());
        }
        assert!(decode_date_time_format(&varint_field(1, 260), options()).is_err());
        assert!(decode_date_time_format(&length_field(14, b"yyyy-MM-dd"), options()).is_err());
        let mut duplicate = valid.clone();
        duplicate.extend_from_slice(&length_field(14, b"MM/dd/yyyy"));
        assert!(decode_date_time_format(&duplicate, options()).is_err());
    }

    #[test]
    fn date_time_format_rewrite_preserves_unknown_shapes_and_exact_accounting() {
        let mut source = native("yyyy-MM-dd");
        let unknown = [
            0xa0, 0x06, 0x81, 0x00, // unknown scalar with overlong value
            0xa5, 0x06, 0x01, 0x23, 0x45, 0x67, // fixed32
            0xa1, 0x06, 0x89, 0xab, 0xcd, 0xef, 0x01, 0x23, 0x45, 0x67, // fixed64
            0xaa, 0x06, 0x03, 0xde, 0xad, 0xbe, // length-delimited
        ];
        source.splice(3..3, unknown);
        source.extend_from_slice(&group_field(50, &varint_field(51, 9)));
        let write = DateTimeFormatWrite::new("MM/dd/yyyy HH:mm:ss");
        let prepared =
            prepare_date_time_format_rewrite(&source, write, options()).expect("prepare");
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
        assert!(
            output
                .bytes()
                .windows(group_field(50, &varint_field(51, 9)).len())
                .any(|window| { window == group_field(50, &varint_field(51, 9)).as_slice() })
        );
        let snapshot = decode_date_time_format(output.bytes(), options()).expect("readback");
        assert_eq!(snapshot.pattern(), "MM/dd/yyyy HH:mm:ss");
        assert_eq!(output.report().fields(), requirements.fields());
        assert_eq!(output.report().work_bytes(), requirements.work_bytes());
        let (_, source_report) =
            decode_date_time_format_with_report(&source, options()).expect("source report");
        let (_, candidate_report) = decode_date_time_format_with_report(output.bytes(), options())
            .expect("candidate report");
        assert_eq!(
            requirements.work_bytes(),
            source_report.work_bytes() + output.bytes().len() + candidate_report.work_bytes()
        );
    }

    #[test]
    fn date_time_format_canonical_write_is_exact_and_bounded() {
        let write = DateTimeFormatWrite::new("yyyy-MM-dd");
        let prepared = prepare_date_time_format_write(write, options()).expect("prepare");
        let requirements = prepared.execution_requirements();
        let output = prepared
            .execute(RewriteExecutionLimits::exact(requirements))
            .expect("execute");
        assert_eq!(output.bytes(), native("yyyy-MM-dd"));
        assert_eq!(requirements.fields(), 2);
        assert_eq!(requirements.work_bytes(), output.bytes().len() * 3);
        assert_eq!(
            canonical_date_time_format(write, options())
                .expect("one shot")
                .bytes(),
            output.bytes()
        );
        assert!(matches!(
            prepared
                .execute(
                    RewriteExecutionLimits::exact(requirements)
                        .with_output_bytes(requirements.output_bytes() - 1)
                )
                .expect_err("output ceiling")
                .resource_limit(),
            Some(DecodeLimit::OutputBytes { .. })
        ));
    }

    #[test]
    fn date_time_format_pattern_budget_is_enforced_before_output_allocation() {
        let oversized = "x".repeat(MAX_DATE_TIME_PATTERN_BYTES + 1);
        assert!(decode_date_time_format(&native(&oversized), options()).is_err());
        assert!(
            prepare_date_time_format_write(DateTimeFormatWrite::new(&oversized), options())
                .is_err()
        );
        let mut limited = options();
        limited = limited.with_max_text_bytes(4);
        assert!(matches!(
            prepare_date_time_format_write(DateTimeFormatWrite::new("yyyy-MM-dd"), limited)
                .expect_err("text ceiling")
                .resource_limit(),
            Some(DecodeLimit::Text { .. })
        ));
    }
}
