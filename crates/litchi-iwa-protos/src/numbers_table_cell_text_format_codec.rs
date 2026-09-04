//! Neutral strict native-Text format seam for Numbers table-cell adapters.
//!
//! Text `FormatStructArchive` payloads contain only the native discriminator
//! (`260`). The source-preserving core keeps all unknown extension records in
//! the caller's bytes while a private Buffa lazy view validates the selected
//! scalar boundary. The BNC explicit metadata route is deliberately separate:
//! both native markers (`0x80` and `0x81`) are accepted only with the Text
//! cell-format kind, and are never treated as a `FormatStructArchive` field.
//! Generated Buffa values never cross this module.

#![allow(clippy::module_name_repetitions)]

use crate::numbers_table_cell_pop_up_menu_codec as core;

pub use core::{
    DecodeError, DecodeLimit, DecodeOptions, DecodeReport, RewriteExecutionLimits,
    RewriteExecutionRequirements, RewriteOutput,
};

/// Native Numbers display-format discriminator for a Text cell.
pub use core::NATIVE_TEXT_FORMAT_TYPE;

/// BNC explicit marker for a plain Text cell.
pub const EXPLICIT_TEXT_FORMAT: u16 = 0x0080;

/// BNC explicit marker for a Text value converted from numeric with retained
/// numeric provenance.
pub const EXPLICIT_CONVERTED_TEXT_FORMAT: u16 = 0x0081;

/// BNC cell-format kind used by both explicit Text markers.
pub const TEXT_CELL_FORMAT_KIND: u32 = 5;

/// The two explicit BNC encodings that identify a Text display route.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u16)]
pub enum TextFormatEncoding {
    /// A cell whose native value is already Text.
    Plain = EXPLICIT_TEXT_FORMAT,
    /// A Text value converted from numeric with retained numeric provenance.
    Converted = EXPLICIT_CONVERTED_TEXT_FORMAT,
}

impl TextFormatEncoding {
    /// Convert an explicit BNC marker to its typed Text encoding.
    #[must_use]
    pub const fn from_native(value: u16) -> Option<Self> {
        match value {
            EXPLICIT_TEXT_FORMAT => Some(Self::Plain),
            EXPLICIT_CONVERTED_TEXT_FORMAT => Some(Self::Converted),
            _ => None,
        }
    }

    /// Return the exact explicit BNC marker represented by this encoding.
    #[must_use]
    pub const fn native_value(self) -> u16 {
        self as u16
    }

    /// Whether this marker denotes numeric-to-Text conversion.
    #[must_use]
    pub const fn is_converted(self) -> bool {
        matches!(self, Self::Converted)
    }
}

/// Validate the explicit BNC marker/kind pair used by a Text cell.
///
/// This checks only the metadata discriminator. Cell-body decoding and value
/// semantics remain owned by the Numbers package's storage codec.
pub fn decode_text_format_encoding(
    explicit_marker: u16,
    cell_format_kind: u32,
) -> Result<TextFormatEncoding, DecodeError> {
    if cell_format_kind != TEXT_CELL_FORMAT_KIND {
        return Err(DecodeError::invalid());
    }
    TextFormatEncoding::from_native(explicit_marker).ok_or_else(DecodeError::invalid)
}

/// Borrowed scalar values for one strict native Text format payload.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct TextFormatSnapshot<'source>(core::TextFormatSnapshot<'source>);

impl<'source> TextFormatSnapshot<'source> {
    /// Borrow the original wire payload.
    #[must_use]
    pub const fn raw(self) -> &'source [u8] {
        self.0.raw()
    }

    /// Return the native Text discriminator (`260`).
    #[must_use]
    pub const fn format_type(self) -> u32 {
        self.0.format_type()
    }
}

/// Scalar values accepted by the strict native Text writer.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct TextFormatWrite(core::TextFormatWrite);

impl TextFormatWrite {
    /// Construct a native Text format update.
    #[must_use]
    pub const fn new() -> Self {
        Self(core::TextFormatWrite::new())
    }

    /// Copy the semantic values from a decoded Text snapshot.
    #[must_use]
    pub const fn from_snapshot(snapshot: TextFormatSnapshot<'_>) -> Self {
        Self(core::TextFormatWrite::from_snapshot(snapshot.0))
    }

    /// Return the native Text discriminator (`260`).
    #[must_use]
    pub const fn format_type(self) -> u32 {
        self.0.format_type()
    }
}

impl Default for TextFormatWrite {
    fn default() -> Self {
        Self::new()
    }
}

/// Prepared source-preserving native Text rewrite.
#[derive(Debug, Clone, Copy)]
pub struct PreparedTextFormatRewrite<'source>(core::PreparedTextFormatRewrite<'source>);

impl PreparedTextFormatRewrite<'_> {
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

    /// Emit, strictly read back, and publish the source-preserving rewrite.
    pub fn execute(self, limits: RewriteExecutionLimits) -> Result<RewriteOutput, DecodeError> {
        self.0.execute(limits)
    }
}

/// Prepared canonical native Text append.
#[derive(Debug, Clone, Copy)]
pub struct PreparedTextFormatWrite(core::PreparedTextFormatWrite);

impl PreparedTextFormatWrite {
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

    /// Emit and strictly read back a canonical Text payload.
    pub fn execute(self, limits: RewriteExecutionLimits) -> Result<RewriteOutput, DecodeError> {
        self.0.execute(limits)
    }
}

/// Strictly decode one native Text `FormatStructArchive`.
pub fn decode_text_format(
    source: &[u8],
    options: DecodeOptions,
) -> Result<TextFormatSnapshot<'_>, DecodeError> {
    core::decode_text_format(source, options).map(TextFormatSnapshot)
}

/// Strictly decode one native Text format and return measured wire use.
pub fn decode_text_format_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(TextFormatSnapshot<'_>, DecodeReport), DecodeError> {
    core::decode_text_format_with_report(source, options)
        .map(|(snapshot, report)| (TextFormatSnapshot(snapshot), report))
}

/// Prepare a source-preserving native Text format rewrite.
pub fn prepare_text_format_rewrite<'source>(
    source: &'source [u8],
    write: TextFormatWrite,
    options: DecodeOptions,
) -> Result<PreparedTextFormatRewrite<'source>, DecodeError> {
    core::prepare_text_format_rewrite(source, write.0, options).map(PreparedTextFormatRewrite)
}

/// Rewrite one native Text format while preserving unknown source fields.
pub fn rewrite_text_format(
    source: &[u8],
    write: TextFormatWrite,
    options: DecodeOptions,
) -> Result<RewriteOutput, DecodeError> {
    core::rewrite_text_format(source, write.0, options)
}

/// Compatibility spelling for the table-cell Text route.
pub use rewrite_text_format as rewrite_table_cell_text_format;

/// Prepare a canonical native Text format payload for a new list entry.
pub fn prepare_text_format_write(
    write: TextFormatWrite,
    options: DecodeOptions,
) -> Result<PreparedTextFormatWrite, DecodeError> {
    core::prepare_text_format_write(write.0, options).map(PreparedTextFormatWrite)
}

/// Prepare a canonical Text append under the explicit append spelling.
pub use prepare_text_format_write as prepare_text_format_append;

/// Encode a canonical native Text format payload for a new list entry.
pub fn canonical_text_format(
    write: TextFormatWrite,
    options: DecodeOptions,
) -> Result<RewriteOutput, DecodeError> {
    core::canonical_text_format(write.0, options)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn options() -> DecodeOptions {
        DecodeOptions::new(16 * 1024, 16 * 1024, 64, 64 * 1024, 64, 0, 0, 0)
    }

    fn push_varint(output: &mut Vec<u8>, mut value: u64) {
        while value >= 0x80 {
            output.push((value as u8 & 0x7f) | 0x80);
            value >>= 7;
        }
        output.push(value as u8);
    }

    fn key(number: u32, wire: u8) -> Vec<u8> {
        let mut output = Vec::new();
        push_varint(&mut output, (u64::from(number) << 3) | u64::from(wire));
        output
    }

    fn field(number: u32, value: u32) -> Vec<u8> {
        varint_field(number, u64::from(value))
    }

    fn varint_field(number: u32, value: u64) -> Vec<u8> {
        let mut output = key(number, 0);
        push_varint(&mut output, value);
        output
    }

    fn fixed32_field(number: u32, value: [u8; 4]) -> Vec<u8> {
        let mut output = key(number, 5);
        output.extend_from_slice(&value);
        output
    }

    fn fixed64_field(number: u32, value: [u8; 8]) -> Vec<u8> {
        let mut output = key(number, 1);
        output.extend_from_slice(&value);
        output
    }

    fn length_field(number: u32, value: &[u8]) -> Vec<u8> {
        let mut output = key(number, 2);
        push_varint(
            &mut output,
            u64::try_from(value.len()).expect("test payload length fits in u64"),
        );
        output.extend_from_slice(value);
        output
    }

    fn group_field(number: u32, body: &[u8]) -> Vec<u8> {
        let mut output = key(number, 3);
        output.extend_from_slice(body);
        output.extend_from_slice(&key(number, 4));
        output
    }

    fn native_text() -> Vec<u8> {
        field(1, NATIVE_TEXT_FORMAT_TYPE)
    }

    #[test]
    fn text_format_accepts_native_discriminator_and_borrows_source() {
        let source = native_text();
        let snapshot = decode_text_format(&source, options()).expect("text format");
        assert_eq!(snapshot.raw(), source.as_slice());
        assert_eq!(snapshot.format_type(), NATIVE_TEXT_FORMAT_TYPE);
        assert_eq!(
            TextFormatWrite::from_snapshot(snapshot),
            TextFormatWrite::new()
        );
    }

    #[test]
    fn text_format_rejects_other_known_fields_and_malformed_shapes() {
        let valid = native_text();
        for extra in [field(2, 0), field(4, 0), field(20, 0), field(45, 0)] {
            let mut source = valid.clone();
            source.extend_from_slice(&extra);
            assert!(decode_text_format(&source, options()).is_err());
        }

        let mut duplicate = valid.clone();
        duplicate.extend_from_slice(&valid);
        assert!(decode_text_format(&duplicate, options()).is_err());
        assert!(decode_text_format(&field(1, 259), options()).is_err());
        assert!(decode_text_format(&[0x08, 0x84, 0x82, 0x00], options()).is_err());
        assert!(decode_text_format(&[], options()).is_err());
    }

    #[test]
    fn text_format_rewrite_preserves_unknown_source_spans() {
        let mut source = native_text();
        let unknown = field(46, 7);
        source.extend_from_slice(&unknown);
        let prepared = prepare_text_format_rewrite(&source, TextFormatWrite::new(), options())
            .expect("prepare");
        let requirements = prepared.execution_requirements();
        let output = prepared
            .execute(RewriteExecutionLimits::exact(requirements))
            .expect("execute");
        assert_eq!(output.bytes(), source.as_slice());
        assert!(
            output
                .bytes()
                .windows(unknown.len())
                .any(|window| window == unknown)
        );
        assert_eq!(output.report().fields(), requirements.fields());
        assert_eq!(output.report().work_bytes(), requirements.work_bytes());
        let (_, source_report) =
            decode_text_format_with_report(&source, options()).expect("source report");
        let (_, candidate_report) =
            decode_text_format_with_report(output.bytes(), options()).expect("candidate report");
        assert_eq!(
            requirements.work_bytes(),
            source_report.work_bytes() + output.bytes().len() + candidate_report.work_bytes()
        );
    }

    #[test]
    fn text_format_rewrite_preserves_every_unknown_wire_shape_and_exact_accounting() {
        let unknown_varint = field(46, 7);
        let unknown_overlong_varint = vec![0xf0, 0x02, 0x87, 0x00];
        let unknown_fixed32 = fixed32_field(47, [0x01, 0x23, 0x45, 0x67]);
        let unknown_fixed64 = fixed64_field(48, [0x89, 0xab, 0xcd, 0xef, 0x10, 0x32, 0x54, 0x76]);
        let unknown_length = length_field(49, b"opaque");
        let unknown_group = group_field(50, &field(51, 9));
        let mut source = native_text();
        for record in [
            &unknown_varint,
            &unknown_overlong_varint,
            &unknown_fixed32,
            &unknown_fixed64,
            &unknown_length,
            &unknown_group,
        ] {
            source.extend_from_slice(record);
        }

        let (_, source_report) =
            decode_text_format_with_report(&source, options()).expect("source report");
        assert!(source_report.work_bytes() > source.len() * 2);
        let prepared = prepare_text_format_rewrite(&source, TextFormatWrite::new(), options())
            .expect("prepare");
        let requirements = prepared.execution_requirements();
        assert_eq!(requirements.output_bytes(), source.len());
        assert_eq!(requirements.fields(), source_report.fields() * 2);
        assert_eq!(requirements.max_depth(), source_report.max_depth());
        assert_eq!(
            requirements.work_bytes(),
            source_report.work_bytes() + source.len() + source_report.work_bytes()
        );
        assert_eq!(requirements.retained_bytes(), source.len() * 2);

        let prepared_report = prepared.prepare_report();
        let output = prepared
            .execute(RewriteExecutionLimits::exact(requirements))
            .expect("execute at exact requirements");
        assert_eq!(prepared_report, output.report());
        assert_eq!(output.bytes(), source.as_slice());
        for record in [
            &unknown_varint,
            &unknown_overlong_varint,
            &unknown_fixed32,
            &unknown_fixed64,
            &unknown_length,
            &unknown_group,
        ] {
            assert!(
                output
                    .bytes()
                    .windows(record.len())
                    .any(|window| window == record.as_slice())
            );
        }
        let (_, candidate_report) =
            decode_text_format_with_report(output.bytes(), options()).expect("candidate report");
        assert_eq!(candidate_report, source_report);
        assert_eq!(
            requirements.work_bytes(),
            source_report.work_bytes()
                + requirements.output_bytes()
                + candidate_report.work_bytes()
        );
    }

    #[test]
    fn text_format_rejects_malformed_groups_and_known_sibling_fields() {
        let valid = native_text();
        for number in 2..=45 {
            let mut source = valid.clone();
            source.extend_from_slice(&field(number, 0));
            assert!(
                decode_text_format(&source, options()).is_err(),
                "known sibling field {number} was accepted"
            );
        }

        let mut unterminated = valid.clone();
        unterminated.extend_from_slice(&key(46, 3));
        unterminated.extend_from_slice(&field(47, 1));
        assert!(decode_text_format(&unterminated, options()).is_err());

        let mut mismatched = valid.clone();
        mismatched.extend_from_slice(&key(46, 3));
        mismatched.extend_from_slice(&field(47, 1));
        mismatched.extend_from_slice(&key(48, 4));
        assert!(decode_text_format(&mismatched, options()).is_err());

        let mut stray_end = valid;
        stray_end.extend_from_slice(&key(46, 4));
        assert!(decode_text_format(&stray_end, options()).is_err());
    }

    #[test]
    fn text_format_rejects_noncanonical_or_wrong_wire_known_field_one() {
        let malformed = [
            vec![0x88, 0x00, 0x84, 0x02],          // overlong field key
            vec![0x08, 0x84, 0x82, 0x00],          // overlong field value
            vec![0x09, 0, 0, 0, 0, 0, 0, 0, 0, 0], // fixed64
            vec![0x0a, 0x01, 0x00],                // length-delimited
            vec![0x0d, 0, 0, 0, 0],                // fixed32
            group_field(1, &[]),                   // group
        ];
        for source in malformed {
            assert!(decode_text_format(&source, options()).is_err());
        }
    }

    #[test]
    fn text_format_accepts_every_unknown_wire_kind_and_high_field_numbers() {
        const MAX_FIELD_NUMBER: u32 = 0x1fff_ffff;

        let unknown_varint = varint_field(46, 7);
        let unknown_fixed64 = fixed64_field(47, [0x89, 0xab, 0xcd, 0xef, 0x10, 0x32, 0x54, 0x76]);
        let unknown_length = length_field(48, &[0x00, 0xff, 0x80]);
        let unknown_group = group_field(49, &varint_field(50, 9));
        let high_unknown = varint_field(MAX_FIELD_NUMBER, 1);

        // Unknown records are valid only when their framing is structurally
        // complete. Their order, wire kind, and high field number must not
        // affect the selected Text discriminator.
        let mut source = unknown_varint.clone();
        source.extend_from_slice(&unknown_fixed64);
        source.extend_from_slice(&unknown_length);
        source.extend_from_slice(&native_text());
        source.extend_from_slice(&unknown_group);
        source.extend_from_slice(&high_unknown);

        let (snapshot, report) =
            decode_text_format_with_report(&source, options()).expect("unknown wire kinds");
        assert_eq!(snapshot.raw(), source.as_slice());
        assert_eq!(snapshot.format_type(), NATIVE_TEXT_FORMAT_TYPE);
        assert_eq!(report.input_bytes(), source.len());
        assert_eq!(report.fields(), 8);
        assert_eq!(report.max_depth(), 1);
        assert!(report.work_bytes() >= source.len() * 2);

        let rewritten = rewrite_text_format(&source, TextFormatWrite::new(), options())
            .expect("source-preserving unknown wire rewrite");
        assert_eq!(rewritten.bytes(), source.as_slice());
        for record in [
            unknown_varint,
            unknown_fixed64,
            unknown_length,
            unknown_group,
            high_unknown,
        ] {
            assert!(
                rewritten
                    .bytes()
                    .windows(record.len())
                    .any(|window| window == record.as_slice()),
                "unknown record was not retained"
            );
        }
    }

    #[test]
    fn text_format_accepts_overlong_unknown_scalar_values_and_preserves_them() {
        // The key is canonical, but the scalar value `1` uses two bytes. This
        // is the deliberately opaque compatibility policy: unknown scalar
        // values are accepted and source-authoritative, unlike selected field
        // values, which must use canonical varints.
        let unknown_root = [0xa0, 0x06, 0x81, 0x00];
        let unknown_nested = [0xa3, 0x06, 0xa8, 0x06, 0x81, 0x00, 0xa4, 0x06];
        let mut source = unknown_root.to_vec();
        source.extend_from_slice(&native_text());
        source.extend_from_slice(&unknown_nested);

        let snapshot = decode_text_format(&source, options()).expect("overlong unknown values");
        assert_eq!(snapshot.raw(), source.as_slice());
        let rewritten =
            rewrite_text_format(&source, TextFormatWrite::from_snapshot(snapshot), options())
                .expect("preserve overlong unknown values");
        assert_eq!(rewritten.bytes(), source.as_slice());
        assert!(
            rewritten
                .bytes()
                .windows(unknown_root.len())
                .any(|window| window == unknown_root)
        );
        assert!(
            rewritten
                .bytes()
                .windows(unknown_nested.len())
                .any(|window| window == unknown_nested)
        );
    }

    #[test]
    fn text_format_rejects_noncanonical_unknown_keys_and_lengths() {
        let malformed = [
            // Field 100 key (`800`) with a redundant zero continuation byte.
            vec![0xa0, 0x86, 0x00, 0x01],
            // Field 102 length-delimited value with an overlong zero length.
            vec![0xb2, 0x06, 0x80, 0x00],
            // The same noncanonical length with one opaque payload byte.
            vec![0xb2, 0x06, 0x81, 0x00, 0x7f],
            // Noncanonical nested key inside an otherwise balanced group.
            vec![0xa3, 0x06, 0xa8, 0x86, 0x00, 0x01, 0xa4, 0x06],
            // Noncanonical nested length inside a balanced group.
            vec![0xa3, 0x06, 0xb2, 0x06, 0x80, 0x00, 0xa4, 0x06],
        ];
        for unknown in malformed {
            let mut source = unknown;
            source.extend_from_slice(&native_text());
            assert!(
                decode_text_format(&source, options()).is_err(),
                "noncanonical unknown framing was accepted: {source:02x?}"
            );
        }
    }

    #[test]
    fn text_format_rejects_all_known_sibling_numbers_and_wire_kinds() {
        for number in 2..=45 {
            for extra in [
                varint_field(number, 0),
                fixed64_field(number, [0; 8]),
                length_field(number, &[0]),
                fixed32_field(number, [0; 4]),
                group_field(number, &[]),
            ] {
                let mut source = native_text();
                source.extend_from_slice(&extra);
                assert!(
                    decode_text_format(&source, options()).is_err(),
                    "known sibling field {number} with wire {} was accepted",
                    extra.first().copied().unwrap_or_default() & 7
                );
            }
        }
    }

    #[test]
    fn text_format_rejects_every_invalid_wire_and_malformed_group_shape() {
        let mut malformed = Vec::new();
        // Wire types 6 and 7 are reserved/invalid, while a top-level end-group
        // can never close a group owned by this message.
        malformed.push(key(46, 4));
        malformed.push(key(46, 6));
        malformed.push(key(46, 7));

        // Truncated fixed-width and length-delimited values.
        malformed.push([key(46, 1), vec![0; 7]].concat());
        malformed.push([key(46, 5), vec![0; 3]].concat());
        malformed.push([key(46, 2), vec![2, 0]].concat());

        // Missing, mismatched, and nested-mismatched group ends.
        malformed.push([key(46, 3), varint_field(47, 1)].concat());
        malformed.push([key(46, 3), varint_field(47, 1), key(48, 4)].concat());
        malformed.push([key(46, 3), key(47, 3), key(46, 4), key(47, 4)].concat());
        malformed.push([key(46, 3), key(47, 3), key(48, 4), key(47, 4)].concat());

        for extra in malformed {
            let mut source = native_text();
            source.extend_from_slice(&extra);
            assert!(
                decode_text_format(&source, options()).is_err(),
                "malformed group/wire shape was accepted: {source:02x?}"
            );
        }
    }

    #[test]
    fn text_format_decode_and_rewrite_budgets_are_inclusive_and_tight() {
        let unknown_group = group_field(49, &varint_field(50, 9));
        let mut source = vec![0xa0, 0x06, 0x81, 0x00];
        source.extend_from_slice(&native_text());
        source.extend_from_slice(&unknown_group);

        let (_, source_report) =
            decode_text_format_with_report(&source, options()).expect("source report");
        let source_options = DecodeOptions::new(
            source.len(),
            source.len(),
            source_report.fields(),
            source_report.work_bytes(),
            source_report.max_depth().max(1),
            0,
            0,
            0,
        );
        // Every decode ceiling is inclusive at the measured boundary.
        decode_text_format(&source, source_options).expect("exact decode limits");
        assert!(matches!(
            decode_text_format(
                &source,
                DecodeOptions::new(
                    source.len() - 1,
                    source.len(),
                    source_report.fields(),
                    source_report.work_bytes(),
                    source_report.max_depth().max(1),
                    0,
                    0,
                    0,
                )
            )
            .expect_err("input minus one")
            .resource_limit(),
            Some(DecodeLimit::InputBytes { .. })
        ));
        assert!(matches!(
            decode_text_format(
                &source,
                DecodeOptions::new(
                    source.len(),
                    source.len(),
                    source_report.fields() - 1,
                    source_report.work_bytes(),
                    source_report.max_depth().max(1),
                    0,
                    0,
                    0,
                )
            )
            .expect_err("fields minus one")
            .resource_limit(),
            Some(DecodeLimit::Fields { .. })
        ));
        assert!(matches!(
            decode_text_format(
                &source,
                DecodeOptions::new(
                    source.len(),
                    source.len(),
                    source_report.fields(),
                    source_report.work_bytes() - 1,
                    source_report.max_depth().max(1),
                    0,
                    0,
                    0,
                )
            )
            .expect_err("work minus one")
            .resource_limit(),
            Some(DecodeLimit::Work { .. })
        ));
        assert!(matches!(
            decode_text_format(
                &source,
                DecodeOptions::new(
                    source.len(),
                    source.len(),
                    source_report.fields(),
                    source_report.work_bytes(),
                    source_report.max_depth() - 1,
                    0,
                    0,
                    0,
                )
            )
            .expect_err("depth minus one")
            .resource_limit(),
            Some(DecodeLimit::Nesting { .. })
        ));

        let broad = DecodeOptions::new(
            source.len(),
            source.len(),
            source_report.fields().checked_mul(2).expect("fields"),
            source_report
                .work_bytes()
                .checked_mul(2)
                .and_then(|work| work.checked_add(source.len()))
                .expect("work"),
            source_report.max_depth().max(1),
            0,
            0,
            0,
        );
        let prepared = prepare_text_format_rewrite(&source, TextFormatWrite::new(), broad)
            .expect("prepare exact rewrite limits");
        let requirements = prepared.execution_requirements();
        prepared
            .execute(RewriteExecutionLimits::exact(requirements))
            .expect("exact rewrite limits");

        for limits in [
            RewriteExecutionLimits::exact(requirements)
                .with_output_bytes(requirements.output_bytes() - 1),
            RewriteExecutionLimits::exact(requirements).with_fields(requirements.fields() - 1),
            RewriteExecutionLimits::exact(requirements)
                .with_work_bytes(requirements.work_bytes() - 1),
            RewriteExecutionLimits::exact(requirements)
                .with_retained_bytes(requirements.retained_bytes() - 1),
            RewriteExecutionLimits::exact(requirements).with_allocations(0),
        ] {
            assert!(
                prepared.execute(limits).is_err(),
                "rewrite accepted one-below exact limits"
            );
        }
    }

    #[test]
    fn text_format_canonical_writer_is_exact() {
        let prepared = prepare_text_format_write(TextFormatWrite::new(), options())
            .expect("prepare canonical");
        let requirements = prepared.execution_requirements();
        let prepared_report = prepared.prepare_report();
        let output = prepared
            .execute(RewriteExecutionLimits::exact(requirements))
            .expect("canonical");
        assert_eq!(output.bytes(), native_text().as_slice());
        assert_eq!(output.report().fields(), 1);
        assert_eq!(output.report().work_bytes(), output.bytes().len() * 3);
        assert_eq!(prepared_report, output.report());
        assert_eq!(
            canonical_text_format(TextFormatWrite::new(), options())
                .expect("canonical helper")
                .bytes(),
            output.bytes()
        );
    }

    #[test]
    fn text_format_encoding_accepts_plain_and_converted_markers() {
        assert_eq!(
            decode_text_format_encoding(EXPLICIT_TEXT_FORMAT, TEXT_CELL_FORMAT_KIND),
            Ok(TextFormatEncoding::Plain)
        );
        assert_eq!(
            decode_text_format_encoding(EXPLICIT_CONVERTED_TEXT_FORMAT, TEXT_CELL_FORMAT_KIND),
            Ok(TextFormatEncoding::Converted)
        );
        assert!(!TextFormatEncoding::Plain.is_converted());
        assert!(TextFormatEncoding::Converted.is_converted());
        assert_eq!(TextFormatEncoding::Plain.native_value(), 0x80);
        assert_eq!(TextFormatEncoding::Converted.native_value(), 0x81);
        for (marker, kind) in [
            (0x00, TEXT_CELL_FORMAT_KIND),
            (0x82, TEXT_CELL_FORMAT_KIND),
            (EXPLICIT_TEXT_FORMAT, 4),
        ] {
            assert!(decode_text_format_encoding(marker, kind).is_err());
        }
    }
}
