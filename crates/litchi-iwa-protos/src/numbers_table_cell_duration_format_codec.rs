//! Neutral strict native Duration format seam for Numbers table-cell
//! adapters.
//!
//! Duration uses the shared source-preserving `FormatStructArchive` wire core
//! while keeping its style and unit vocabulary behind a nominal family
//! boundary. The complete source payload remains authoritative; generated
//! Buffa values never cross this module.

#![allow(clippy::module_name_repetitions)]

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
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DurationFormatSnapshot<'source>(core::DurationFormatSnapshot<'source>);

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
#[derive(Debug, Clone, Copy)]
pub struct PreparedDurationFormatRewrite<'source>(core::PreparedDurationFormatRewrite<'source>);

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
#[derive(Debug, Clone, Copy)]
pub struct PreparedDurationFormatWrite(core::PreparedDurationFormatWrite);

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

    fn varint_field(number: u32, value: u64) -> Vec<u8> {
        let mut output = Vec::new();
        push_varint(&mut output, u64::from(number) << 3);
        push_varint(&mut output, value);
        output
    }

    fn group_field(number: u32, body: &[u8]) -> Vec<u8> {
        let mut output = Vec::new();
        push_varint(&mut output, (u64::from(number) << 3) | 3);
        output.extend_from_slice(body);
        push_varint(&mut output, (u64::from(number) << 3) | 4);
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
    fn duration_format_rejects_bad_domains_missing_fields_and_cross_families() {
        let valid = native(
            NATIVE_DURATION_STYLE_COLON,
            NATIVE_DURATION_UNIT_WEEKS,
            NATIVE_DURATION_UNIT_SECONDS,
            false,
        );
        for invalid in [
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
            native(
                NATIVE_DURATION_STYLE_COLON,
                NATIVE_DURATION_UNIT_WEEKS,
                NATIVE_DURATION_UNIT_SECONDS,
                true,
            ),
        ] {
            let mut source = invalid;
            if source.last() == Some(&0x01) {
                // Replace the final canonical bool with an invalid domain.
                source.pop();
                source.extend_from_slice(&varint_field(40, 2));
            }
            if source != valid {
                assert!(decode_duration_format(&source, options()).is_err());
            }
        }
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
    }
}
