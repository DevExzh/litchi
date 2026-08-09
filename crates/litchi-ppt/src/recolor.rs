//! Typed metafile recolor mappings from MS-PPT sections 2.7.9 through 2.7.13.
//!
//! Parsing is limited to bytes already present in a caller-supplied PPT record.
//! This module never opens or renders the referenced metafile.

use super::package::{Error, Result};
use super::records::Record;
use crate::consts::RecordType;

const PREFIX_BYTES: usize = 12;
const ENTRY_BYTES: usize = 44;
const VARIANT_BYTES: usize = 34;
const USED_HEADER_FLAGS: u16 = 0x0017;

#[allow(
    clippy::module_name_repetitions,
    reason = "`RecolorLimits` is the established public API name for the recolor parsing limits; renaming it would break downstream crates"
)]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RecolorLimits {
    pub max_record_bytes: usize,
    pub max_entries: usize,
    pub max_trailing_bytes: usize,
}

impl Default for RecolorLimits {
    fn default() -> Self {
        Self {
            max_record_bytes: 16 * 1024 * 1024,
            max_entries: 65_536,
            max_trailing_bytes: 1024 * 1024,
        }
    }
}

/// Six-byte `WideColorStruct` from MS-PPT section 2.12.3.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct WideColor {
    pub red: u16,
    pub green: u16,
    pub blue: u16,
}

impl WideColor {
    #[must_use]
    pub const fn new(red: u16, green: u16, blue: u16) -> Self {
        Self { red, green, blue }
    }
}

/// MS-WMF section 2.1.1.4 `BrushStyle`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u16)]
pub enum WmfBrushStyle {
    Solid = 0,
    Null = 1,
    Hatched = 2,
    Pattern = 3,
    Indexed = 4,
    DibPattern = 5,
    DibPatternPointer = 6,
    Pattern8x8 = 7,
    DibPattern8x8 = 8,
    MonoPattern = 9,
}

impl TryFrom<u16> for WmfBrushStyle {
    type Error = Error;

    fn try_from(value: u16) -> Result<Self> {
        match value {
            0 => Ok(Self::Solid),
            1 => Ok(Self::Null),
            2 => Ok(Self::Hatched),
            3 => Ok(Self::Pattern),
            4 => Ok(Self::Indexed),
            5 => Ok(Self::DibPattern),
            6 => Ok(Self::DibPatternPointer),
            7 => Ok(Self::Pattern8x8),
            8 => Ok(Self::DibPattern8x8),
            9 => Ok(Self::MonoPattern),
            _ => corrupted("RecolorEntryBrush has an invalid BrushStyle"),
        }
    }
}

/// MS-WMF section 2.1.1.12 `HatchStyle`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u16)]
pub enum WmfHatchStyle {
    Horizontal = 0,
    Vertical = 1,
    ForwardDiagonal = 2,
    BackwardDiagonal = 3,
    Cross = 4,
    DiagonalCross = 5,
}

impl TryFrom<u16> for WmfHatchStyle {
    type Error = Error;

    fn try_from(value: u16) -> Result<Self> {
        match value {
            0 => Ok(Self::Horizontal),
            1 => Ok(Self::Vertical),
            2 => Ok(Self::ForwardDiagonal),
            3 => Ok(Self::BackwardDiagonal),
            4 => Ok(Self::Cross),
            5 => Ok(Self::DiagonalCross),
            _ => corrupted("RecolorEntryBrush has an invalid HatchStyle"),
        }
    }
}

#[allow(
    clippy::module_name_repetitions,
    reason = "`RecolorBitmapType` is the established public API name mirroring the MS-PPT recolor bitmap-type field; renaming it would break downstream crates"
)]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u16)]
pub enum RecolorBitmapType {
    MonochromePattern = 0,
    DibPattern = 1,
    NonMonochromeOrIndirect = 3,
}

impl TryFrom<u16> for RecolorBitmapType {
    type Error = Error;

    fn try_from(value: u16) -> Result<Self> {
        match value {
            0 => Ok(Self::MonochromePattern),
            1 => Ok(Self::DibPattern),
            3 => Ok(Self::NonMonochromeOrIndirect),
            _ => corrupted("RecolorEntryBrush has an invalid bitmapType"),
        }
    }
}

/// Conditional `lbHatch` representation. Non-hatched brushes retain its ignored value.
#[allow(
    clippy::module_name_repetitions,
    reason = "`RecolorHatch` is the established public API name for the recolor brush hatch field; renaming it would break downstream crates"
)]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RecolorHatch {
    Hatched(WmfHatchStyle),
    Ignored(u16),
}

/// Conditional pattern representation. Non-pattern brushes retain ignored bytes.
#[allow(
    clippy::module_name_repetitions,
    reason = "`RecolorPattern` is the established public API name for the recolor brush pattern field; renaming it would break downstream crates"
)]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RecolorPattern {
    Pattern {
        bitmap_type: RecolorBitmapType,
        bytes: [u8; 8],
    },
    Ignored {
        bitmap_type: u16,
        bytes: [u8; 8],
    },
}

#[allow(
    clippy::module_name_repetitions,
    reason = "`RecolorBrush` is the established public API name for the recolor brush structure; renaming it would break downstream crates"
)]
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RecolorBrush {
    pub style: WmfBrushStyle,
    pub color: WideColor,
    pub hatch: RecolorHatch,
    pub foreground_color: WideColor,
    pub background_color: WideColor,
    pub pattern: RecolorPattern,
}

#[allow(
    clippy::module_name_repetitions,
    reason = "`RecolorSource` is the established public API name for the recolor source variant; renaming it would break downstream crates"
)]
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum RecolorSource {
    Color {
        color: WideColor,
        /// Undefined bytes retained without interpretation.
        unused: [u8; 26],
    },
    Brush(RecolorBrush),
}

#[allow(
    clippy::module_name_repetitions,
    reason = "`RecolorEntry` is the established public API name mirroring the MS-PPT `RecolorEntryStruct`; renaming it would break downstream crates"
)]
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RecolorEntry {
    pub do_recolor: bool,
    /// Destination color; ignored by consumers when `destination_index < 8`.
    pub destination_color: WideColor,
    /// Scheme index for values below eight, otherwise `destination_color` is used.
    pub destination_index: u8,
    /// Undefined byte retained without interpretation.
    pub unused: u8,
    pub source: RecolorSource,
}

#[allow(
    clippy::module_name_repetitions,
    reason = "`RecolorInfo` is the established public API name mirroring the MS-PPT `RecolorInfoAtom` record; renaming it would break downstream crates"
)]
#[allow(
    clippy::struct_excessive_bools,
    reason = "each bool maps one-to-one to an independent flag bit of the MS-PPT `RecolorInfoAtom` header bitfield; grouping them into enums would misrepresent the on-disk layout and churn the public API"
)]
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RecolorInfo {
    pub should_recolor: bool,
    pub missing_colors: bool,
    pub missing_fills: bool,
    pub mono_recolor: bool,
    /// Undefined flag bits retained without interpretation.
    pub ignored_flags: u16,
    pub mono_color: WideColor,
    pub entries: Vec<RecolorEntry>,
    /// Undefined bytes following the fixed-size entry array.
    pub trailing_unused: Vec<u8>,
}

impl RecolorInfo {
    /// Parse a complete `RecolorInfoAtom` record within `limits`.
    ///
    /// # Errors
    ///
    /// Returns an error if the record header is invalid, the declared data
    /// length disagrees with the payload, or the payload violates `limits` or
    /// the MS-PPT structure constraints.
    pub fn parse(record: &Record, limits: RecolorLimits) -> Result<Self> {
        if record.record_type != RecordType::RecolorInfoAtom
            || record.record_type_raw != 0x0fe7
            || record.version != 0
            || record.instance != 0
        {
            return corrupted("RecolorInfoAtom has an invalid record header");
        }
        if usize::try_from(record.data_length).ok() != Some(record.data.len()) {
            return corrupted("RecolorInfoAtom data length disagrees with its payload");
        }
        Self::parse_payload(&record.data, limits)
    }

    /// Parse a `RecolorInfoAtom` payload within `limits`.
    ///
    /// # Errors
    ///
    /// Returns an error if the payload is truncated, exceeds `limits`, or
    /// violates the MS-PPT entry structure or count constraints.
    pub fn parse_payload(payload: &[u8], limits: RecolorLimits) -> Result<Self> {
        if payload.len() > limits.max_record_bytes {
            return corrupted("RecolorInfoAtom exceeds the record byte limit");
        }
        if payload.len() < PREFIX_BYTES {
            return corrupted("RecolorInfoAtom is truncated");
        }
        let flags = u16_at(payload, 0)?;
        let color_count = usize::from(u16_at(payload, 2)?);
        let fill_count = usize::from(u16_at(payload, 4)?);
        let entry_count = color_count
            .checked_add(fill_count)
            .ok_or_else(|| corrupt("RecolorInfoAtom entry count overflow"))?;
        if entry_count > limits.max_entries {
            return corrupted("RecolorInfoAtom exceeds the entry count limit");
        }
        let entry_bytes = entry_count
            .checked_mul(ENTRY_BYTES)
            .ok_or_else(|| corrupt("RecolorInfoAtom entry size overflow"))?;
        let required = PREFIX_BYTES
            .checked_add(entry_bytes)
            .ok_or_else(|| corrupt("RecolorInfoAtom payload size overflow"))?;
        if required > payload.len() {
            return corrupted("RecolorInfoAtom entry array is truncated");
        }
        if payload.len() - required > limits.max_trailing_bytes {
            return corrupted("RecolorInfoAtom exceeds the trailing byte limit");
        }
        let mono_color = parse_color(&payload[6..12])?;
        let mut entries = Vec::with_capacity(entry_count);
        let mut parsed_colors = 0usize;
        let mut parsed_fills = 0usize;
        for bytes in payload[PREFIX_BYTES..required].chunks_exact(ENTRY_BYTES) {
            let entry = parse_entry(bytes)?;
            match entry.source {
                RecolorSource::Color { .. } => parsed_colors += 1,
                RecolorSource::Brush(_) => parsed_fills += 1,
            }
            entries.push(entry);
        }
        if parsed_colors != color_count || parsed_fills != fill_count {
            return corrupted("RecolorInfoAtom variant counts disagree with cColors and cFills");
        }
        Ok(Self {
            should_recolor: flags & 0x0001 != 0,
            missing_colors: flags & 0x0002 != 0,
            missing_fills: flags & 0x0004 != 0,
            mono_recolor: flags & 0x0010 != 0,
            ignored_flags: flags & !USED_HEADER_FLAGS,
            mono_color,
            entries,
            trailing_unused: payload[required..].to_vec(),
        })
    }

    /// Serialize the payload within `limits`.
    ///
    /// # Errors
    ///
    /// Returns an error if the entries, counts, or trailing bytes violate
    /// `limits` or the MS-PPT structure constraints.
    pub fn to_payload(&self, limits: RecolorLimits) -> Result<Vec<u8>> {
        validate(self, limits)?;
        let color_count = self
            .entries
            .iter()
            .filter(|entry| matches!(entry.source, RecolorSource::Color { .. }))
            .count();
        let fill_count = self.entries.len() - color_count;
        let size = PREFIX_BYTES
            .checked_add(
                self.entries
                    .len()
                    .checked_mul(ENTRY_BYTES)
                    .ok_or_else(|| corrupt("RecolorInfoAtom entry size overflow"))?,
            )
            .and_then(|value| value.checked_add(self.trailing_unused.len()))
            .ok_or_else(|| corrupt("RecolorInfoAtom payload size overflow"))?;
        let mut output = Vec::with_capacity(size);
        let mut flags = self.ignored_flags;
        flags |= u16::from(self.should_recolor);
        flags |= u16::from(self.missing_colors) << 1;
        flags |= u16::from(self.missing_fills) << 2;
        flags |= u16::from(self.mono_recolor) << 4;
        output.extend_from_slice(&flags.to_le_bytes());
        output.extend_from_slice(
            &u16::try_from(color_count)
                .map_err(|_err| corrupt("RecolorInfoAtom color count exceeds u16"))?
                .to_le_bytes(),
        );
        output.extend_from_slice(
            &u16::try_from(fill_count)
                .map_err(|_err| corrupt("RecolorInfoAtom fill count exceeds u16"))?
                .to_le_bytes(),
        );
        write_color(&mut output, self.mono_color);
        for entry in &self.entries {
            write_entry(&mut output, entry)?;
        }
        output.extend_from_slice(&self.trailing_unused);
        Ok(output)
    }

    /// Serialize as a `Record` within `limits`.
    ///
    /// # Errors
    ///
    /// Returns an error if the value violates `limits` or the serialized
    /// payload exceeds a 32-bit length.
    pub fn to_record(&self, limits: RecolorLimits) -> Result<Record> {
        let data = self.to_payload(limits)?;
        let data_length = u32::try_from(data.len())
            .map_err(|_err| corrupt("RecolorInfoAtom payload exceeds u32"))?;
        Ok(Record {
            record_type: RecordType::RecolorInfoAtom,
            record_type_raw: 0x0fe7,
            version: 0,
            instance: 0,
            data_length,
            data,
            children: Vec::new(),
        })
    }
}

fn parse_entry(bytes: &[u8]) -> Result<RecolorEntry> {
    if bytes.len() != ENTRY_BYTES {
        return corrupted("RecolorEntry has an invalid size");
    }
    let flags = u16_at(bytes, 0)?;
    if flags & !1 != 0 {
        return corrupted("RecolorEntry reserved bits are nonzero");
    }
    let source = parse_source(&bytes[10..])?;
    Ok(RecolorEntry {
        do_recolor: flags & 1 != 0,
        destination_color: parse_color(&bytes[2..8])?,
        destination_index: bytes[8],
        unused: bytes[9],
        source,
    })
}

fn parse_source(bytes: &[u8]) -> Result<RecolorSource> {
    if bytes.len() != VARIANT_BYTES {
        return corrupted("RecolorEntryVariant has an invalid size");
    }
    match u16_at(bytes, 0)? {
        0 => {
            let mut unused = [0u8; 26];
            unused.copy_from_slice(&bytes[8..34]);
            Ok(RecolorSource::Color {
                color: parse_color(&bytes[2..8])?,
                unused,
            })
        },
        1 => {
            let style = WmfBrushStyle::try_from(u16_at(bytes, 2)?)?;
            let raw_hatch = u16_at(bytes, 10)?;
            let hatch = if style == WmfBrushStyle::Hatched {
                RecolorHatch::Hatched(WmfHatchStyle::try_from(raw_hatch)?)
            } else {
                RecolorHatch::Ignored(raw_hatch)
            };
            let raw_bitmap_type = u16_at(bytes, 24)?;
            let mut pattern_bytes = [0u8; 8];
            pattern_bytes.copy_from_slice(&bytes[26..34]);
            let pattern = if style == WmfBrushStyle::Pattern {
                RecolorPattern::Pattern {
                    bitmap_type: RecolorBitmapType::try_from(raw_bitmap_type)?,
                    bytes: pattern_bytes,
                }
            } else {
                RecolorPattern::Ignored {
                    bitmap_type: raw_bitmap_type,
                    bytes: pattern_bytes,
                }
            };
            Ok(RecolorSource::Brush(RecolorBrush {
                style,
                color: parse_color(&bytes[4..10])?,
                hatch,
                foreground_color: parse_color(&bytes[12..18])?,
                background_color: parse_color(&bytes[18..24])?,
                pattern,
            }))
        },
        _ => corrupted("RecolorEntryVariant has an invalid type"),
    }
}

fn write_entry(output: &mut Vec<u8>, entry: &RecolorEntry) -> Result<()> {
    output.extend_from_slice(&u16::from(entry.do_recolor).to_le_bytes());
    write_color(output, entry.destination_color);
    output.push(entry.destination_index);
    output.push(entry.unused);
    match &entry.source {
        RecolorSource::Color { color, unused } => {
            output.extend_from_slice(&0u16.to_le_bytes());
            write_color(output, *color);
            output.extend_from_slice(unused);
        },
        RecolorSource::Brush(brush) => {
            validate_brush(brush)?;
            output.extend_from_slice(&1u16.to_le_bytes());
            output.extend_from_slice(&(brush.style as u16).to_le_bytes());
            write_color(output, brush.color);
            let hatch = match brush.hatch {
                RecolorHatch::Hatched(value) => value as u16,
                RecolorHatch::Ignored(value) => value,
            };
            output.extend_from_slice(&hatch.to_le_bytes());
            write_color(output, brush.foreground_color);
            write_color(output, brush.background_color);
            match brush.pattern {
                RecolorPattern::Pattern { bitmap_type, bytes } => {
                    output.extend_from_slice(&(bitmap_type as u16).to_le_bytes());
                    output.extend_from_slice(&bytes);
                },
                RecolorPattern::Ignored { bitmap_type, bytes } => {
                    output.extend_from_slice(&bitmap_type.to_le_bytes());
                    output.extend_from_slice(&bytes);
                },
            }
        },
    }
    Ok(())
}

fn validate(value: &RecolorInfo, limits: RecolorLimits) -> Result<()> {
    if value.ignored_flags & USED_HEADER_FLAGS != 0 {
        return corrupted("RecolorInfoAtom ignored flags overlap defined flags");
    }
    if value.entries.len() > limits.max_entries {
        return corrupted("RecolorInfoAtom exceeds the entry count limit");
    }
    if value.trailing_unused.len() > limits.max_trailing_bytes {
        return corrupted("RecolorInfoAtom exceeds the trailing byte limit");
    }
    let colors = value
        .entries
        .iter()
        .filter(|entry| matches!(entry.source, RecolorSource::Color { .. }))
        .count();
    let fills = value.entries.len() - colors;
    if colors > usize::from(u16::MAX) || fills > usize::from(u16::MAX) {
        return corrupted("RecolorInfoAtom variant count exceeds u16");
    }
    let size = PREFIX_BYTES
        .checked_add(
            value
                .entries
                .len()
                .checked_mul(ENTRY_BYTES)
                .ok_or_else(|| corrupt("RecolorInfoAtom entry size overflow"))?,
        )
        .and_then(|size| size.checked_add(value.trailing_unused.len()))
        .ok_or_else(|| corrupt("RecolorInfoAtom payload size overflow"))?;
    if size > limits.max_record_bytes || size > u32::MAX as usize {
        return corrupted("RecolorInfoAtom exceeds the record byte limit");
    }
    for entry in &value.entries {
        if let RecolorSource::Brush(brush) = &entry.source {
            validate_brush(brush)?;
        }
    }
    Ok(())
}

fn validate_brush(brush: &RecolorBrush) -> Result<()> {
    match (brush.style, brush.hatch) {
        (WmfBrushStyle::Hatched, RecolorHatch::Hatched(_)) => {},
        (WmfBrushStyle::Hatched, RecolorHatch::Ignored(_)) => {
            return corrupted("hatched RecolorEntryBrush lacks a typed HatchStyle");
        },
        (_, RecolorHatch::Hatched(_)) => {
            return corrupted("non-hatched RecolorEntryBrush contains a typed HatchStyle");
        },
        (_, RecolorHatch::Ignored(_)) => {},
    }
    match (brush.style, brush.pattern) {
        (WmfBrushStyle::Pattern, RecolorPattern::Pattern { .. }) => {},
        (WmfBrushStyle::Pattern, RecolorPattern::Ignored { .. }) => {
            return corrupted("pattern RecolorEntryBrush lacks a typed bitmapType");
        },
        (_, RecolorPattern::Pattern { .. }) => {
            return corrupted("non-pattern RecolorEntryBrush contains a typed bitmapType");
        },
        (_, RecolorPattern::Ignored { .. }) => {},
    }
    Ok(())
}

fn parse_color(bytes: &[u8]) -> Result<WideColor> {
    if bytes.len() != 6 {
        return corrupted("WideColorStruct has an invalid size");
    }
    Ok(WideColor {
        red: u16_at(bytes, 0)?,
        green: u16_at(bytes, 2)?,
        blue: u16_at(bytes, 4)?,
    })
}

fn write_color(output: &mut Vec<u8>, color: WideColor) {
    output.extend_from_slice(&color.red.to_le_bytes());
    output.extend_from_slice(&color.green.to_le_bytes());
    output.extend_from_slice(&color.blue.to_le_bytes());
}

fn u16_at(bytes: &[u8], offset: usize) -> Result<u16> {
    let value = bytes
        .get(offset..offset + 2)
        .ok_or_else(|| corrupt("recolor structure is truncated"))?;
    Ok(u16::from_le_bytes([value[0], value[1]]))
}

fn corrupt(message: &str) -> Error {
    Error::Corrupted(message.to_string())
}

fn corrupted<T>(message: &str) -> Result<T> {
    Err(corrupt(message))
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
mod tests {
    use super::*;

    fn limits() -> RecolorLimits {
        RecolorLimits {
            max_record_bytes: 4096,
            max_entries: 8,
            max_trailing_bytes: 16,
        }
    }

    fn color(red: u16, green: u16, blue: u16) -> WideColor {
        WideColor::new(red, green, blue)
    }

    fn sample() -> RecolorInfo {
        RecolorInfo {
            should_recolor: true,
            missing_colors: false,
            missing_fills: true,
            mono_recolor: true,
            ignored_flags: 0xa008,
            mono_color: color(1, 2, 3),
            entries: vec![
                RecolorEntry {
                    do_recolor: true,
                    destination_color: color(4, 5, 6),
                    destination_index: 8,
                    unused: 0x7f,
                    source: RecolorSource::Color {
                        color: color(7, 8, 9),
                        unused: [0x55; 26],
                    },
                },
                RecolorEntry {
                    do_recolor: false,
                    destination_color: color(10, 11, 12),
                    destination_index: 2,
                    unused: 0x80,
                    source: RecolorSource::Brush(RecolorBrush {
                        style: WmfBrushStyle::Pattern,
                        color: color(13, 14, 15),
                        hatch: RecolorHatch::Ignored(0xbeef),
                        foreground_color: color(16, 17, 18),
                        background_color: color(19, 20, 21),
                        pattern: RecolorPattern::Pattern {
                            bitmap_type: RecolorBitmapType::DibPattern,
                            bytes: [0xaa, 0x55, 0xaa, 0x55, 0xaa, 0x55, 0xaa, 0x55],
                        },
                    }),
                },
            ],
            trailing_unused: vec![0xde, 0xad],
        }
    }

    #[test]
    fn parses_and_writes_complete_recolor_family_byte_exactly() {
        let expected = sample();
        let record = expected.to_record(limits()).unwrap();
        assert_eq!(record.record_type, RecordType::RecolorInfoAtom);
        assert_eq!(record.data.len(), 102);
        assert_eq!(&record.data[0..6], &[0x1d, 0xa0, 1, 0, 1, 0]);
        assert_eq!(&record.data[12..14], &[1, 0]);
        assert_eq!(&record.data[22..24], &[0, 0]);
        assert_eq!(&record.data[56..58], &[0, 0]);
        assert_eq!(&record.data[66..68], &[1, 0]);
        assert_eq!(&record.data[68..70], &[3, 0]);
        assert_eq!(&record.data[100..], &[0xde, 0xad]);
        let parsed = RecolorInfo::parse(&record, limits()).unwrap();
        assert_eq!(parsed, expected);
        assert_eq!(parsed.to_payload(limits()).unwrap(), record.data);
    }

    #[test]
    fn validates_hatched_and_pattern_conditional_payloads() {
        let mut value = sample();
        {
            let RecolorSource::Brush(brush) = &mut value.entries[1].source else {
                unreachable!()
            };
            brush.style = WmfBrushStyle::Hatched;
            brush.hatch = RecolorHatch::Hatched(WmfHatchStyle::DiagonalCross);
            brush.pattern = RecolorPattern::Ignored {
                bitmap_type: 0xffff,
                bytes: [0xcc; 8],
            };
        }
        let payload = value.to_payload(limits()).unwrap();
        assert_eq!(
            RecolorInfo::parse_payload(&payload, limits()).unwrap(),
            value
        );

        let RecolorSource::Brush(brush) = &mut value.entries[1].source else {
            unreachable!()
        };
        brush.hatch = RecolorHatch::Ignored(5);
        assert!(value.to_payload(limits()).is_err());
        let RecolorSource::Brush(solid_brush) = &mut value.entries[1].source else {
            unreachable!()
        };
        solid_brush.style = WmfBrushStyle::Solid;
        solid_brush.hatch = RecolorHatch::Hatched(WmfHatchStyle::Horizontal);
        assert!(value.to_payload(limits()).is_err());
    }

    #[test]
    fn rejects_reserved_discriminant_count_size_and_limit_attacks() {
        let payload = sample().to_payload(limits()).unwrap();
        let mutate = |offset: usize, bytes: &[u8]| {
            let mut value = payload.clone();
            value[offset..offset + bytes.len()].copy_from_slice(bytes);
            value
        };
        assert!(RecolorInfo::parse_payload(&mutate(12, &[2, 0]), limits()).is_err());
        assert!(RecolorInfo::parse_payload(&mutate(22, &[2, 0]), limits()).is_err());
        assert!(RecolorInfo::parse_payload(&mutate(2, &[2, 0]), limits()).is_err());
        assert!(RecolorInfo::parse_payload(&payload[..99], limits()).is_err());
        assert!(
            RecolorInfo::parse_payload(
                &payload,
                RecolorLimits {
                    max_record_bytes: 99,
                    ..limits()
                }
            )
            .is_err()
        );
        assert!(
            RecolorInfo::parse_payload(
                &payload,
                RecolorLimits {
                    max_entries: 1,
                    ..limits()
                }
            )
            .is_err()
        );
        assert!(
            RecolorInfo::parse_payload(
                &payload,
                RecolorLimits {
                    max_trailing_bytes: 1,
                    ..limits()
                }
            )
            .is_err()
        );

        let mut invalid_style = payload.clone();
        invalid_style[68..70].copy_from_slice(&10u16.to_le_bytes());
        assert!(RecolorInfo::parse_payload(&invalid_style, limits()).is_err());
        let mut invalid_bitmap = payload.clone();
        invalid_bitmap[90..92].copy_from_slice(&2u16.to_le_bytes());
        assert!(RecolorInfo::parse_payload(&invalid_bitmap, limits()).is_err());
    }
}
