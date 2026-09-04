//! Neutral strict native Duration format seam for Numbers table-cell
//! adapters.
//!
//! Duration uses the shared source-preserving `FormatStructArchive` wire core
//! while keeping its style and unit vocabulary behind a nominal family
//! boundary. The complete source payload remains authoritative; generated
//! Buffa values never cross this module.

#![allow(clippy::module_name_repetitions)]

use std::fmt;

use crate::numbers_table_cell_pop_up_menu_codec as core;

pub use core::{
    DecodeError, DecodeLimit, DecodeOptions, DecodeReport, RewriteExecutionLimits,
    RewriteExecutionRequirements, RewriteOutput,
};

/// Native Numbers display-format discriminator for a duration cell.
pub use core::NATIVE_DURATION_FORMAT_TYPE;

/// Native duration style for colon-separated fields.
pub use core::NATIVE_DURATION_STYLE_COLON;

/// Native duration style for abbreviated unit symbols.
pub use core::NATIVE_DURATION_STYLE_ABBREVIATED;

/// Native duration style for complete unit names.
pub use core::NATIVE_DURATION_STYLE_FULL_NAMES;

/// Native duration unit bit for weeks.
pub use core::NATIVE_DURATION_UNIT_WEEKS;

/// Native duration unit bit for days.
pub use core::NATIVE_DURATION_UNIT_DAYS;

/// Native duration unit bit for hours.
pub use core::NATIVE_DURATION_UNIT_HOURS;

/// Native duration unit bit for minutes.
pub use core::NATIVE_DURATION_UNIT_MINUTES;

/// Native duration unit bit for seconds.
pub use core::NATIVE_DURATION_UNIT_SECONDS;

/// Native duration unit bit for milliseconds.
pub use core::NATIVE_DURATION_UNIT_MILLISECONDS;

/// The presentation styles accepted by native Numbers Duration.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u32)]
pub enum DurationStyle {
    /// Display selected units as colon-separated fields.
    Colon = NATIVE_DURATION_STYLE_COLON,
    /// Display compact unit symbols.
    Abbreviated = NATIVE_DURATION_STYLE_ABBREVIATED,
    /// Display complete unit names.
    FullNames = NATIVE_DURATION_STYLE_FULL_NAMES,
}

impl DurationStyle {
    /// Convert a native wire value to a supported presentation style.
    #[must_use]
    pub const fn from_native(value: u32) -> Option<Self> {
        match value {
            NATIVE_DURATION_STYLE_COLON => Some(Self::Colon),
            NATIVE_DURATION_STYLE_ABBREVIATED => Some(Self::Abbreviated),
            NATIVE_DURATION_STYLE_FULL_NAMES => Some(Self::FullNames),
            _ => None,
        }
    }

    /// Return the exact unsigned value stored in field 7.
    #[must_use]
    pub const fn native_value(self) -> u32 {
        self as u32
    }
}

/// The unit bit values accepted by native Numbers Duration.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
#[repr(u32)]
pub enum DurationUnit {
    /// Weeks.
    Weeks = NATIVE_DURATION_UNIT_WEEKS,
    /// Days.
    Days = NATIVE_DURATION_UNIT_DAYS,
    /// Hours.
    Hours = NATIVE_DURATION_UNIT_HOURS,
    /// Minutes.
    Minutes = NATIVE_DURATION_UNIT_MINUTES,
    /// Seconds.
    Seconds = NATIVE_DURATION_UNIT_SECONDS,
    /// Milliseconds.
    Milliseconds = NATIVE_DURATION_UNIT_MILLISECONDS,
}

impl DurationUnit {
    /// Convert a native wire value to a supported unit.
    #[must_use]
    pub const fn from_native(value: u32) -> Option<Self> {
        match value {
            NATIVE_DURATION_UNIT_WEEKS => Some(Self::Weeks),
            NATIVE_DURATION_UNIT_DAYS => Some(Self::Days),
            NATIVE_DURATION_UNIT_HOURS => Some(Self::Hours),
            NATIVE_DURATION_UNIT_MINUTES => Some(Self::Minutes),
            NATIVE_DURATION_UNIT_SECONDS => Some(Self::Seconds),
            NATIVE_DURATION_UNIT_MILLISECONDS => Some(Self::Milliseconds),
            _ => None,
        }
    }

    /// Return the exact unsigned value stored in fields 15 or 16.
    #[must_use]
    pub const fn native_value(self) -> u32 {
        self as u32
    }
}

/// Borrowed scalar values for one strict native Duration format payload.
#[derive(Clone, Copy, PartialEq, Eq)]
pub struct DurationFormatSnapshot<'source>(core::DurationFormatSnapshot<'source>);

impl fmt::Debug for DurationFormatSnapshot<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("DurationFormatSnapshot")
            .field("format_type", &self.format_type())
            .field("duration_style", &self.duration_style())
            .field("style", &self.style())
            .field("duration_unit_largest", &self.duration_unit_largest())
            .field("largest_unit", &self.largest_unit())
            .field("duration_unit_smallest", &self.duration_unit_smallest())
            .field("smallest_unit", &self.smallest_unit())
            .field(
                "use_automatic_duration_units",
                &self.use_automatic_duration_units(),
            )
            .field("source_bytes", &self.raw().len())
            .finish()
    }
}

impl<'source> DurationFormatSnapshot<'source> {
    /// Borrow the original wire payload.
    #[must_use]
    pub const fn raw(self) -> &'source [u8] {
        self.0.raw()
    }

    /// Return the native Duration discriminator (`268`).
    #[must_use]
    pub const fn format_type(self) -> u32 {
        self.0.format_type()
    }

    /// Return the native field-7 presentation style.
    #[must_use]
    pub const fn duration_style(self) -> u32 {
        self.0.duration_style()
    }

    /// Return the typed field-7 presentation style.
    #[must_use]
    pub const fn style(self) -> Option<DurationStyle> {
        DurationStyle::from_native(self.duration_style())
    }

    /// Return the native field-15 largest displayed unit.
    #[must_use]
    pub const fn duration_unit_largest(self) -> u32 {
        self.0.duration_unit_largest()
    }

    /// Return the typed field-15 largest displayed unit.
    #[must_use]
    pub const fn largest_unit(self) -> Option<DurationUnit> {
        DurationUnit::from_native(self.duration_unit_largest())
    }

    /// Return the native field-16 smallest displayed unit.
    #[must_use]
    pub const fn duration_unit_smallest(self) -> u32 {
        self.0.duration_unit_smallest()
    }

    /// Return the typed field-16 smallest displayed unit.
    #[must_use]
    pub const fn smallest_unit(self) -> Option<DurationUnit> {
        DurationUnit::from_native(self.duration_unit_smallest())
    }

    /// Return whether Numbers selects visible units automatically.
    #[must_use]
    pub const fn use_automatic_duration_units(self) -> bool {
        self.0.use_automatic_duration_units()
    }

    /// Return whether Numbers selects visible units automatically.
    #[must_use]
    pub const fn is_automatic(self) -> bool {
        self.use_automatic_duration_units()
    }
}

/// Scalar values accepted by the strict native Duration writer.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DurationFormatWrite(core::DurationFormatWrite);

impl DurationFormatWrite {
    /// Construct a native Duration format update from its four wire values.
    #[must_use]
    pub const fn new(
        duration_style: u32,
        duration_unit_largest: u32,
        duration_unit_smallest: u32,
        use_automatic_duration_units: bool,
    ) -> Self {
        Self(core::DurationFormatWrite::new(
            duration_style,
            duration_unit_largest,
            duration_unit_smallest,
            use_automatic_duration_units,
        ))
    }

    /// Construct a native Duration update from typed values.
    #[must_use]
    pub const fn from_parts(
        style: DurationStyle,
        largest_unit: DurationUnit,
        smallest_unit: DurationUnit,
        automatic_units: bool,
    ) -> Self {
        Self::new(
            style.native_value(),
            largest_unit.native_value(),
            smallest_unit.native_value(),
            automatic_units,
        )
    }

    /// Copy the semantic values from a decoded Duration snapshot.
    #[must_use]
    pub const fn from_snapshot(snapshot: DurationFormatSnapshot<'_>) -> Self {
        Self(core::DurationFormatWrite::from_snapshot(snapshot.0))
    }

    /// Return the requested field-7 presentation style.
    #[must_use]
    pub const fn duration_style(self) -> u32 {
        self.0.duration_style()
    }

    /// Return the typed requested presentation style.
    #[must_use]
    pub const fn style(self) -> Option<DurationStyle> {
        DurationStyle::from_native(self.duration_style())
    }

    /// Return the requested field-15 largest displayed unit.
    #[must_use]
    pub const fn duration_unit_largest(self) -> u32 {
        self.0.duration_unit_largest()
    }

    /// Return the typed requested largest displayed unit.
    #[must_use]
    pub const fn largest_unit(self) -> Option<DurationUnit> {
        DurationUnit::from_native(self.duration_unit_largest())
    }

    /// Return the requested field-16 smallest displayed unit.
    #[must_use]
    pub const fn duration_unit_smallest(self) -> u32 {
        self.0.duration_unit_smallest()
    }

    /// Return the typed requested smallest displayed unit.
    #[must_use]
    pub const fn smallest_unit(self) -> Option<DurationUnit> {
        DurationUnit::from_native(self.duration_unit_smallest())
    }

    /// Return whether the requested format uses automatic unit selection.
    #[must_use]
    pub const fn use_automatic_duration_units(self) -> bool {
        self.0.use_automatic_duration_units()
    }

    /// Return whether the requested format uses automatic unit selection.
    #[must_use]
    pub const fn is_automatic(self) -> bool {
        self.use_automatic_duration_units()
    }

    /// Replace the requested presentation style.
    #[must_use]
    pub const fn with_duration_style(self, value: u32) -> Self {
        Self(self.0.with_duration_style(value))
    }

    /// Replace the requested presentation style with a typed value.
    #[must_use]
    pub const fn with_style(self, value: DurationStyle) -> Self {
        self.with_duration_style(value.native_value())
    }

    /// Replace the requested largest displayed unit.
    #[must_use]
    pub const fn with_duration_unit_largest(self, value: u32) -> Self {
        Self(self.0.with_duration_unit_largest(value))
    }

    /// Replace the requested largest displayed unit with a typed value.
    #[must_use]
    pub const fn with_largest_unit(self, value: DurationUnit) -> Self {
        self.with_duration_unit_largest(value.native_value())
    }

    /// Replace the requested smallest displayed unit.
    #[must_use]
    pub const fn with_duration_unit_smallest(self, value: u32) -> Self {
        Self(self.0.with_duration_unit_smallest(value))
    }

    /// Replace the requested smallest displayed unit with a typed value.
    #[must_use]
    pub const fn with_smallest_unit(self, value: DurationUnit) -> Self {
        self.with_duration_unit_smallest(value.native_value())
    }

    /// Replace the requested automatic-unit selection.
    #[must_use]
    pub const fn with_use_automatic_duration_units(self, value: bool) -> Self {
        Self(self.0.with_use_automatic_duration_units(value))
    }

    /// Replace the requested automatic-unit selection.
    #[must_use]
    pub const fn with_automatic(self, value: bool) -> Self {
        self.with_use_automatic_duration_units(value)
    }
}

/// Prepared source-preserving native Duration rewrite.
#[derive(Clone, Copy)]
pub struct PreparedDurationFormatRewrite<'source>(core::PreparedDurationFormatRewrite<'source>);

impl fmt::Debug for PreparedDurationFormatRewrite<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("PreparedDurationFormatRewrite")
            .field("prepared", &self.0)
            .finish()
    }
}

impl PreparedDurationFormatRewrite<'_> {
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

/// Prepared canonical native Duration append.
#[derive(Clone, Copy)]
pub struct PreparedDurationFormatWrite(core::PreparedDurationFormatWrite);

impl fmt::Debug for PreparedDurationFormatWrite {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("PreparedDurationFormatWrite")
            .field("prepared", &self.0)
            .finish()
    }
}

impl PreparedDurationFormatWrite {
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

    /// Emit and strictly read back a canonical Duration payload.
    pub fn execute(self, limits: RewriteExecutionLimits) -> Result<RewriteOutput, DecodeError> {
        self.0.execute(limits)
    }
}

/// Strictly decode one native Duration `FormatStructArchive`.
pub fn decode_duration_format(
    source: &[u8],
    options: DecodeOptions,
) -> Result<DurationFormatSnapshot<'_>, DecodeError> {
    core::decode_duration_format(source, options).map(DurationFormatSnapshot)
}

/// Strictly decode one native Duration format and return measured wire use.
pub fn decode_duration_format_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(DurationFormatSnapshot<'_>, DecodeReport), DecodeError> {
    core::decode_duration_format_with_report(source, options)
        .map(|(snapshot, report)| (DurationFormatSnapshot(snapshot), report))
}

/// Prepare a source-preserving native Duration format rewrite.
pub fn prepare_duration_format_rewrite<'source>(
    source: &'source [u8],
    write: DurationFormatWrite,
    options: DecodeOptions,
) -> Result<PreparedDurationFormatRewrite<'source>, DecodeError> {
    core::prepare_duration_format_rewrite(source, write.0, options)
        .map(PreparedDurationFormatRewrite)
}

/// Rewrite one native Duration format while preserving unknown source fields
/// byte-for-byte.
pub fn rewrite_duration_format(
    source: &[u8],
    write: DurationFormatWrite,
    options: DecodeOptions,
) -> Result<RewriteOutput, DecodeError> {
    core::rewrite_duration_format(source, write.0, options)
}

/// Compatibility spelling for the table-cell Duration route.
pub use rewrite_duration_format as rewrite_table_cell_duration_format;

/// Prepare a canonical native Duration format payload for a new list entry.
pub fn prepare_duration_format_write(
    write: DurationFormatWrite,
    options: DecodeOptions,
) -> Result<PreparedDurationFormatWrite, DecodeError> {
    core::prepare_duration_format_write(write.0, options).map(PreparedDurationFormatWrite)
}

/// Prepare a canonical Duration append under the explicit append spelling.
pub use prepare_duration_format_write as prepare_duration_format_append;

/// Encode a canonical native Duration format payload for a new list entry.
pub fn canonical_duration_format(
    write: DurationFormatWrite,
    options: DecodeOptions,
) -> Result<RewriteOutput, DecodeError> {
    core::canonical_duration_format(write.0, options)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn options() -> DecodeOptions {
        DecodeOptions::new(16 * 1024, 16 * 1024, 16 * 1024, 64 * 1024, 64, 0, 0, 0)
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

    fn native(style: u32, largest: u32, smallest: u32, automatic_units: bool) -> Vec<u8> {
        let mut output = varint_field(1, u64::from(NATIVE_DURATION_FORMAT_TYPE));
        output.extend_from_slice(&varint_field(7, u64::from(style)));
        output.extend_from_slice(&varint_field(15, u64::from(largest)));
        output.extend_from_slice(&varint_field(16, u64::from(smallest)));
        output.extend_from_slice(&varint_field(40, u64::from(automatic_units)));
        output
    }

    fn native_encoded(
        format_type: &[u8],
        style: &[u8],
        largest: &[u8],
        smallest: &[u8],
        automatic_units: &[u8],
    ) -> Vec<u8> {
        let mut output = Vec::new();
        for (number, value) in [
            (1, format_type),
            (7, style),
            (15, largest),
            (16, smallest),
            (40, automatic_units),
        ] {
            output.extend_from_slice(&key(number, 0));
            output.extend_from_slice(value);
        }
        output
    }

    #[test]
    fn duration_format_accepts_native_fields_and_typed_values() {
        let source = native(
            NATIVE_DURATION_STYLE_ABBREVIATED,
            NATIVE_DURATION_UNIT_HOURS,
            NATIVE_DURATION_UNIT_MILLISECONDS,
            true,
        );
        let snapshot = decode_duration_format(&source, options()).expect("duration");
        assert_eq!(snapshot.raw(), source.as_slice());
        assert_eq!(snapshot.format_type(), NATIVE_DURATION_FORMAT_TYPE);
        assert_eq!(snapshot.duration_style(), NATIVE_DURATION_STYLE_ABBREVIATED);
        assert_eq!(snapshot.style(), Some(DurationStyle::Abbreviated));
        assert_eq!(snapshot.duration_unit_largest(), NATIVE_DURATION_UNIT_HOURS);
        assert_eq!(snapshot.largest_unit(), Some(DurationUnit::Hours));
        assert_eq!(snapshot.smallest_unit(), Some(DurationUnit::Milliseconds));
        assert!(snapshot.is_automatic());
        assert_eq!(
            DurationFormatWrite::from_snapshot(snapshot).style(),
            snapshot.style()
        );
    }

    #[test]
    fn duration_format_accepts_every_style_and_ordered_unit_pair() {
        let styles = [
            DurationStyle::Colon,
            DurationStyle::Abbreviated,
            DurationStyle::FullNames,
        ];
        let units = [
            DurationUnit::Weeks,
            DurationUnit::Days,
            DurationUnit::Hours,
            DurationUnit::Minutes,
            DurationUnit::Seconds,
            DurationUnit::Milliseconds,
        ];

        for style in styles {
            for (largest_index, largest_unit) in units.into_iter().enumerate() {
                for smallest_unit in units.into_iter().skip(largest_index) {
                    for automatic_units in [false, true] {
                        let source = native(
                            style.native_value(),
                            largest_unit.native_value(),
                            smallest_unit.native_value(),
                            automatic_units,
                        );
                        let (snapshot, report) =
                            decode_duration_format_with_report(&source, options())
                                .expect("valid style/unit pair");

                        assert_eq!(snapshot.raw(), source.as_slice());
                        assert_eq!(snapshot.style(), Some(style));
                        assert_eq!(snapshot.largest_unit(), Some(largest_unit));
                        assert_eq!(snapshot.smallest_unit(), Some(smallest_unit));
                        assert_eq!(snapshot.is_automatic(), automatic_units);
                        assert_eq!(
                            DurationFormatWrite::from_snapshot(snapshot),
                            DurationFormatWrite::from_parts(
                                style,
                                largest_unit,
                                smallest_unit,
                                automatic_units
                            )
                        );
                        assert_eq!(report.input_bytes(), source.len());
                        assert_eq!(report.output_bytes(), 0);
                        assert_eq!(report.fields(), 5);
                        assert_eq!(report.max_depth(), 0);
                        assert_eq!(report.references(), 0);
                        assert_eq!(report.items(), 0);
                        assert_eq!(report.text_bytes(), 0);
                        assert_eq!(report.allocations(), 0);
                        assert_eq!(report.retained_bytes(), 0);
                        assert_eq!(report.scratch_bytes(), 0);
                    }
                }
            }
        }
    }

    #[test]
    fn duration_format_rejects_bad_domains_missing_fields_and_cross_families() {
        let valid = native(
            NATIVE_DURATION_STYLE_COLON,
            NATIVE_DURATION_UNIT_WEEKS,
            NATIVE_DURATION_UNIT_SECONDS,
            false,
        );
        for source in [
            native(
                3,
                NATIVE_DURATION_UNIT_WEEKS,
                NATIVE_DURATION_UNIT_SECONDS,
                false,
            ),
            native(
                NATIVE_DURATION_STYLE_COLON,
                NATIVE_DURATION_UNIT_SECONDS,
                NATIVE_DURATION_UNIT_HOURS,
                false,
            ),
        ] {
            if source != valid {
                assert!(decode_duration_format(&source, options()).is_err());
            }
        }
        let mut invalid_bool = native(
            NATIVE_DURATION_STYLE_COLON,
            NATIVE_DURATION_UNIT_WEEKS,
            NATIVE_DURATION_UNIT_SECONDS,
            true,
        );
        *invalid_bool
            .last_mut()
            .expect("canonical Duration bool has a value byte") = 2;
        assert!(decode_duration_format(&invalid_bool, options()).is_err());
        for omitted in [1_u32, 7, 15, 16, 40] {
            let mut source = Vec::new();
            for (field, value) in [
                (1, u64::from(NATIVE_DURATION_FORMAT_TYPE)),
                (7, u64::from(NATIVE_DURATION_STYLE_COLON)),
                (15, u64::from(NATIVE_DURATION_UNIT_WEEKS)),
                (16, u64::from(NATIVE_DURATION_UNIT_SECONDS)),
                (40, 0),
            ] {
                if field != omitted {
                    source.extend_from_slice(&varint_field(field, value));
                }
            }
            assert!(decode_duration_format(&source, options()).is_err());
        }
        for extra in [varint_field(2, 0), varint_field(4, 0), varint_field(20, 0)] {
            let mut source = valid.clone();
            source.extend_from_slice(&extra);
            assert!(decode_duration_format(&source, options()).is_err());
        }
        let mut duplicate = valid.clone();
        duplicate.extend_from_slice(&varint_field(7, 0));
        assert!(decode_duration_format(&duplicate, options()).is_err());
        assert!(decode_duration_format(&varint_field(1, 261), options()).is_err());
        for (style, largest, smallest) in [
            (NATIVE_DURATION_STYLE_COLON, 0, NATIVE_DURATION_UNIT_SECONDS),
            (NATIVE_DURATION_STYLE_COLON, 3, NATIVE_DURATION_UNIT_SECONDS),
            (NATIVE_DURATION_STYLE_COLON, NATIVE_DURATION_UNIT_WEEKS, 3),
            (
                NATIVE_DURATION_STYLE_COLON,
                64,
                NATIVE_DURATION_UNIT_SECONDS,
            ),
            (NATIVE_DURATION_STYLE_COLON, NATIVE_DURATION_UNIT_WEEKS, 64),
            (
                u32::MAX,
                NATIVE_DURATION_UNIT_WEEKS,
                NATIVE_DURATION_UNIT_SECONDS,
            ),
        ] {
            assert!(
                decode_duration_format(&native(style, largest, smallest, false), options())
                    .is_err(),
                "invalid duration domain was accepted: style={style}, largest={largest}, smallest={smallest}"
            );
        }
        assert!(
            decode_duration_format(&varint_field(1, u64::from(u32::MAX) + 1), options()).is_err()
        );
    }

    #[test]
    fn duration_format_rejects_every_known_wire_shape_and_noncanonical_value() {
        let valid = native(
            NATIVE_DURATION_STYLE_COLON,
            NATIVE_DURATION_UNIT_WEEKS,
            NATIVE_DURATION_UNIT_SECONDS,
            false,
        );

        // All fields in the owning archive's known range are reserved for a
        // different format family unless they are one of Duration's five
        // selected fields. This remains true regardless of wire kind.
        for number in 2..=45 {
            for extra in [
                varint_field(number, 0),
                fixed64_field(number, [0; 8]),
                length_field(number, &[0]),
                fixed32_field(number, [0; 4]),
                group_field(number, &[]),
            ] {
                let mut source = valid.clone();
                source.extend_from_slice(&extra);
                assert!(
                    decode_duration_format(&source, options()).is_err(),
                    "known sibling field {number} was accepted"
                );
            }
        }

        // Selected fields are scalar varints, and their key/value varints
        // must use their shortest encoding. Appending a wrong wire kind is
        // intentionally tested separately from duplicate detection.
        for number in [1_u32, 7, 15, 16, 40] {
            for extra in [
                fixed64_field(number, [0; 8]),
                length_field(number, &[0]),
                group_field(number, &[]),
                fixed32_field(number, [0; 4]),
            ] {
                let mut source = valid.clone();
                source.extend_from_slice(&extra);
                assert!(
                    decode_duration_format(&source, options()).is_err(),
                    "selected field {number} accepted a non-varint wire kind"
                );
            }
        }

        let canonical_style = [NATIVE_DURATION_STYLE_COLON as u8];
        let canonical_largest = [NATIVE_DURATION_UNIT_WEEKS as u8];
        let canonical_smallest = [NATIVE_DURATION_UNIT_SECONDS as u8];
        let canonical_automatic = [0_u8];
        let malformed_values = [
            // Type 268 encoded with a redundant continuation group.
            native_encoded(
                &[0x8c, 0x82, 0x00],
                &canonical_style,
                &canonical_largest,
                &canonical_smallest,
                &canonical_automatic,
            ),
            // Each selected scalar below is zero/one but overlong.
            native_encoded(
                &[0x8c, 0x02],
                &[0x80, 0x00],
                &canonical_largest,
                &canonical_smallest,
                &canonical_automatic,
            ),
            native_encoded(
                &[0x8c, 0x02],
                &canonical_style,
                &[0x81, 0x00],
                &canonical_smallest,
                &canonical_automatic,
            ),
            native_encoded(
                &[0x8c, 0x02],
                &canonical_style,
                &canonical_largest,
                &[0x90, 0x00],
                &canonical_automatic,
            ),
            native_encoded(
                &[0x8c, 0x02],
                &canonical_style,
                &canonical_largest,
                &canonical_smallest,
                &[0x80, 0x00],
            ),
        ];
        for source in malformed_values {
            assert!(
                decode_duration_format(&source, options()).is_err(),
                "noncanonical selected value was accepted: {source:02x?}"
            );
        }

        // The first selected key is also required to be canonical.
        let mut noncanonical_key = valid.clone();
        noncanonical_key.splice(0..1, [0x88, 0x00]);
        assert!(decode_duration_format(&noncanonical_key, options()).is_err());

        // Unknown fields remain opaque semantically, but their framing is
        // still bounded and canonical so a malformed length/key cannot hide
        // bytes from the scanner.
        for unknown in [
            vec![0xf0, 0x82, 0x00, 0x01],       // field 46 key with redundant zero
            vec![0xfa, 0x02, 0x80, 0x00],       // field 47 zero length overlong
            vec![0xfa, 0x02, 0x81, 0x00, 0xff], // field 47 overlong length
        ] {
            let mut source = unknown;
            source.extend_from_slice(&valid);
            assert!(decode_duration_format(&source, options()).is_err());
        }
    }

    #[test]
    fn duration_format_rejects_duplicate_selected_fields() {
        let valid = native(
            NATIVE_DURATION_STYLE_COLON,
            NATIVE_DURATION_UNIT_WEEKS,
            NATIVE_DURATION_UNIT_SECONDS,
            false,
        );
        for (number, value) in [
            (1, u64::from(NATIVE_DURATION_FORMAT_TYPE)),
            (7, u64::from(NATIVE_DURATION_STYLE_COLON)),
            (15, u64::from(NATIVE_DURATION_UNIT_WEEKS)),
            (16, u64::from(NATIVE_DURATION_UNIT_SECONDS)),
            (40, 0),
        ] {
            let mut source = valid.clone();
            source.extend_from_slice(&varint_field(number, value));
            assert!(
                decode_duration_format(&source, options()).is_err(),
                "selected field {number} duplicate was accepted"
            );
        }
    }

    #[test]
    fn duration_format_rewrite_preserves_unknown_shapes_and_order() {
        let mut source = native(
            NATIVE_DURATION_STYLE_ABBREVIATED,
            NATIVE_DURATION_UNIT_HOURS,
            NATIVE_DURATION_UNIT_MILLISECONDS,
            true,
        );
        let unknown = [
            0xa0, 0x06, 0x81, 0x00, // unknown scalar with overlong value
            0xa5, 0x06, 0x01, 0x23, 0x45, 0x67, // fixed32
            0xa1, 0x06, 0x89, 0xab, 0xcd, 0xef, 0x01, 0x23, 0x45, 0x67, // fixed64
            0xaa, 0x06, 0x03, 0xde, 0xad, 0xbe, // length-delimited
        ];
        source.splice(3..3, unknown);
        let group = group_field(50, &varint_field(51, 9));
        source.extend_from_slice(&group);
        let write = DurationFormatWrite::from_parts(
            DurationStyle::FullNames,
            DurationUnit::Minutes,
            DurationUnit::Seconds,
            false,
        );
        let prepared = prepare_duration_format_rewrite(&source, write, options()).expect("prepare");
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
                .windows(group.len())
                .any(|window| window == group)
        );
        let snapshot = decode_duration_format(output.bytes(), options()).expect("readback");
        assert_eq!(snapshot.style(), Some(DurationStyle::FullNames));
        assert_eq!(snapshot.largest_unit(), Some(DurationUnit::Minutes));
        assert_eq!(snapshot.smallest_unit(), Some(DurationUnit::Seconds));
        assert!(!snapshot.is_automatic());
        assert_eq!(output.report().fields(), requirements.fields());
        assert_eq!(output.report().work_bytes(), requirements.work_bytes());
    }

    #[test]
    fn duration_format_accepts_every_unknown_wire_kind_and_preserves_it() {
        const MAX_FIELD_NUMBER: u32 = 0x1fff_ffff;

        let unknown_varint = vec![0xf0, 0x02, 0x81, 0x00];
        let unknown_fixed32 = fixed32_field(47, [0x01, 0x23, 0x45, 0x67]);
        let unknown_fixed64 = fixed64_field(48, [0x89, 0xab, 0xcd, 0xef, 0x10, 0x32, 0x54, 0x76]);
        let unknown_length = length_field(49, &[0x00, 0xff, 0x80]);
        let unknown_group = group_field(50, &varint_field(51, 9));
        let high_unknown = varint_field(MAX_FIELD_NUMBER, 1);

        // Put unknown records on both sides of selected fields to ensure the
        // rewrite's source-span ordering does not normalize or drop them.
        let mut source = unknown_varint.clone();
        source.extend_from_slice(&unknown_fixed32);
        source.extend_from_slice(&native(
            NATIVE_DURATION_STYLE_ABBREVIATED,
            NATIVE_DURATION_UNIT_HOURS,
            NATIVE_DURATION_UNIT_MILLISECONDS,
            true,
        ));
        source.extend_from_slice(&unknown_fixed64);
        source.extend_from_slice(&unknown_length);
        source.extend_from_slice(&unknown_group);
        source.extend_from_slice(&high_unknown);

        let (snapshot, source_report) =
            decode_duration_format_with_report(&source, options()).expect("unknown wire kinds");
        assert_eq!(snapshot.raw(), source.as_slice());
        assert_eq!(snapshot.style(), Some(DurationStyle::Abbreviated));
        assert_eq!(snapshot.largest_unit(), Some(DurationUnit::Hours));
        assert_eq!(snapshot.smallest_unit(), Some(DurationUnit::Milliseconds));
        assert!(snapshot.is_automatic());
        assert_eq!(source_report.input_bytes(), source.len());
        assert!(source_report.fields() > 5);
        assert_eq!(source_report.max_depth(), 1);
        assert_eq!(source_report.allocations(), 0);
        assert_eq!(source_report.retained_bytes(), 0);
        assert_eq!(source_report.scratch_bytes(), 0);

        let write = DurationFormatWrite::from_parts(
            DurationStyle::FullNames,
            DurationUnit::Minutes,
            DurationUnit::Seconds,
            false,
        );
        let output = rewrite_duration_format(&source, write, options()).expect("rewrite");
        for record in [
            unknown_varint,
            unknown_fixed32,
            unknown_fixed64,
            unknown_length,
            unknown_group,
            high_unknown,
        ] {
            assert!(
                output
                    .bytes()
                    .windows(record.len())
                    .any(|window| window == record.as_slice()),
                "unknown record was not retained: {record:02x?}"
            );
        }
        let candidate = decode_duration_format(output.bytes(), options()).expect("readback");
        assert_eq!(candidate.style(), Some(DurationStyle::FullNames));
        assert_eq!(candidate.largest_unit(), Some(DurationUnit::Minutes));
        assert_eq!(candidate.smallest_unit(), Some(DurationUnit::Seconds));
        assert!(!candidate.is_automatic());
        assert_eq!(output.report().fields(), source_report.fields());
        assert_eq!(output.report().max_depth(), source_report.max_depth());
        assert_eq!(
            output.report().retained_bytes(),
            output.bytes().len() + source.len()
        );
        assert_eq!(output.report().allocations(), 1);
        assert_eq!(output.report().scratch_bytes(), 0);
    }

    #[test]
    fn duration_format_limits_are_inclusive_and_writes_are_fallible() {
        let mut source = vec![0xf0, 0x02, 0x81, 0x00];
        source.extend_from_slice(&native(
            NATIVE_DURATION_STYLE_ABBREVIATED,
            NATIVE_DURATION_UNIT_HOURS,
            NATIVE_DURATION_UNIT_MILLISECONDS,
            true,
        ));
        source.extend_from_slice(&group_field(50, &varint_field(51, 9)));

        let (_, source_report) =
            decode_duration_format_with_report(&source, options()).expect("source report");
        let exact_decode_options = DecodeOptions::new(
            source.len(),
            source.len(),
            source_report.fields(),
            source_report.work_bytes(),
            source_report.max_depth().max(1),
            0,
            0,
            0,
        );
        decode_duration_format(&source, exact_decode_options).expect("inclusive decode limits");
        assert!(matches!(
            decode_duration_format(
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
            decode_duration_format(
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
            decode_duration_format(
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
            decode_duration_format(
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
        // Rewrite preparation performs the same finite preflight, and
        // execution rejects every non-zero dimension one unit below its
        // measured requirement before reserving the output Vec.
        let write = DurationFormatWrite::from_parts(
            DurationStyle::Abbreviated,
            DurationUnit::Hours,
            DurationUnit::Milliseconds,
            true,
        );
        let broad_options = DecodeOptions::new(
            source.len(),
            source.len(),
            source_report.fields().checked_mul(2).expect("fields"),
            source_report.work_bytes().checked_mul(4).expect("work"),
            source_report.max_depth().max(1),
            0,
            0,
            0,
        );
        let prepared = prepare_duration_format_rewrite(&source, write, broad_options)
            .expect("prepare rewrite");
        let requirements = prepared.execution_requirements();
        assert_eq!(requirements.references(), 0);
        assert_eq!(requirements.items(), 0);
        assert_eq!(requirements.text_bytes(), 0);
        assert_eq!(requirements.scratch_bytes(), 0);
        let exact_limits = RewriteExecutionLimits::exact(requirements);
        let output = prepared
            .execute(exact_limits)
            .expect("exact rewrite limits");
        assert_eq!(output.report(), prepared.prepare_report());

        let output_error = prepared
            .execute(exact_limits.with_output_bytes(requirements.output_bytes() - 1))
            .expect_err("output minus one");
        assert!(matches!(
            output_error.resource_limit(),
            Some(DecodeLimit::OutputBytes { .. })
        ));
        let fields_error = prepared
            .execute(exact_limits.with_fields(requirements.fields() - 1))
            .expect_err("fields minus one");
        assert!(matches!(
            fields_error.resource_limit(),
            Some(DecodeLimit::Fields { .. })
        ));
        let work_error = prepared
            .execute(exact_limits.with_work_bytes(requirements.work_bytes() - 1))
            .expect_err("work minus one");
        assert!(matches!(
            work_error.resource_limit(),
            Some(DecodeLimit::Work { .. })
        ));
        let depth_error = prepared
            .execute(exact_limits.with_max_depth(requirements.max_depth() - 1))
            .expect_err("depth minus one");
        assert!(matches!(
            depth_error.resource_limit(),
            Some(DecodeLimit::Nesting { .. })
        ));
        let allocations_error = prepared
            .execute(exact_limits.with_allocations(requirements.allocations() - 1))
            .expect_err("allocations minus one");
        assert!(matches!(
            allocations_error.resource_limit(),
            Some(DecodeLimit::Allocation { .. })
        ));
        let retained_error = prepared
            .execute(exact_limits.with_retained_bytes(requirements.retained_bytes() - 1))
            .expect_err("retained bytes minus one");
        assert!(matches!(
            retained_error.resource_limit(),
            Some(DecodeLimit::Retained { .. })
        ));
    }

    #[test]
    fn duration_format_canonical_write_is_exact_and_bounded() {
        let write = DurationFormatWrite::from_parts(
            DurationStyle::Colon,
            DurationUnit::Weeks,
            DurationUnit::Milliseconds,
            true,
        );
        let prepared = prepare_duration_format_write(write, options()).expect("prepare");
        let requirements = prepared.execution_requirements();
        let output = prepared
            .execute(RewriteExecutionLimits::exact(requirements))
            .expect("execute");
        assert_eq!(output.bytes(), native(0, 1, 32, true));
        assert_eq!(requirements.fields(), 5);
        assert_eq!(requirements.work_bytes(), output.bytes().len() * 3);
        assert_eq!(
            canonical_duration_format(write, options())
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

        for limits in [
            RewriteExecutionLimits::exact(requirements).with_fields(requirements.fields() - 1),
            RewriteExecutionLimits::exact(requirements)
                .with_work_bytes(requirements.work_bytes() - 1),
            RewriteExecutionLimits::exact(requirements)
                .with_allocations(requirements.allocations() - 1),
            RewriteExecutionLimits::exact(requirements)
                .with_retained_bytes(requirements.retained_bytes() - 1),
        ] {
            assert!(
                prepared.execute(limits).is_err(),
                "canonical write accepted one-below execution requirements"
            );
        }

        for invalid in [
            DurationFormatWrite::new(
                3,
                NATIVE_DURATION_UNIT_WEEKS,
                NATIVE_DURATION_UNIT_SECONDS,
                false,
            ),
            DurationFormatWrite::new(
                NATIVE_DURATION_STYLE_COLON,
                0,
                NATIVE_DURATION_UNIT_SECONDS,
                false,
            ),
            DurationFormatWrite::new(
                NATIVE_DURATION_STYLE_COLON,
                NATIVE_DURATION_UNIT_WEEKS,
                0,
                false,
            ),
            DurationFormatWrite::new(
                NATIVE_DURATION_STYLE_COLON,
                NATIVE_DURATION_UNIT_SECONDS,
                NATIVE_DURATION_UNIT_HOURS,
                false,
            ),
            DurationFormatWrite::new(
                NATIVE_DURATION_STYLE_COLON,
                3,
                NATIVE_DURATION_UNIT_SECONDS,
                false,
            ),
            DurationFormatWrite::new(
                NATIVE_DURATION_STYLE_COLON,
                NATIVE_DURATION_UNIT_WEEKS,
                3,
                false,
            ),
            DurationFormatWrite::new(
                u32::MAX,
                NATIVE_DURATION_UNIT_WEEKS,
                NATIVE_DURATION_UNIT_SECONDS,
                false,
            ),
        ] {
            assert!(prepare_duration_format_write(invalid, options()).is_err());
            assert!(canonical_duration_format(invalid, options()).is_err());
        }
    }

    #[test]
    fn duration_format_prepare_rejects_unusable_hard_limits() {
        let write = DurationFormatWrite::from_parts(
            DurationStyle::Colon,
            DurationUnit::Weeks,
            DurationUnit::Seconds,
            false,
        );
        let oversized_message =
            DecodeOptions::new(usize::MAX, 16 * 1024, 16 * 1024, 64 * 1024, 64, 0, 0, 0);
        assert!(matches!(
            prepare_duration_format_write(write, oversized_message)
                .expect_err("oversized message must fail during preparation")
                .resource_limit(),
            Some(DecodeLimit::InputBytes { observed, .. }) if observed == usize::MAX
        ));

        for recursion_limit in [0, 65] {
            let invalid_recursion = DecodeOptions::new(
                16 * 1024,
                16 * 1024,
                16 * 1024,
                64 * 1024,
                recursion_limit,
                0,
                0,
                0,
            );
            assert!(matches!(
                prepare_duration_format_write(write, invalid_recursion)
                    .expect_err("invalid recursion must fail during preparation")
                    .resource_limit(),
                Some(DecodeLimit::Nesting { .. })
            ));
        }
    }

    #[test]
    fn duration_debug_redacts_source_bytes_from_snapshots_and_prepared_views() {
        // Use values that would be visible in a byte-slice Debug rendering;
        // semantic fields and lengths may be shown, but this payload must not.
        let sentinel = [241_u8, 199, 233, 211, 157, 179, 251];
        let mut source = native(
            NATIVE_DURATION_STYLE_ABBREVIATED,
            NATIVE_DURATION_UNIT_HOURS,
            NATIVE_DURATION_UNIT_SECONDS,
            true,
        );
        source.extend_from_slice(&length_field(46, &sentinel));
        let snapshot = decode_duration_format(&source, options()).expect("duration");
        let snapshot_debug = format!("{snapshot:?}");
        assert!(snapshot_debug.contains("source_bytes"));
        assert!(!snapshot_debug.contains("241, 199, 233, 211, 157, 179, 251"));

        let write = DurationFormatWrite::from_parts(
            DurationStyle::FullNames,
            DurationUnit::Minutes,
            DurationUnit::Seconds,
            false,
        );
        let prepared =
            prepare_duration_format_rewrite(&source, write, options()).expect("prepare rewrite");
        let prepared_debug = format!("{prepared:?}");
        assert!(prepared_debug.contains("source_bytes"));
        assert!(prepared_debug.contains("requirements"));
        assert!(!prepared_debug.contains("241, 199, 233, 211, 157, 179, 251"));

        let canonical = prepare_duration_format_write(write, options()).expect("prepare write");
        let canonical_debug = format!("{canonical:?}");
        assert!(canonical_debug.contains("requirements"));
        assert!(!canonical_debug.contains("241, 199, 233, 211, 157, 179, 251"));
    }
}
