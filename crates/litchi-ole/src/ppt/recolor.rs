//! Typed metafile recolor mappings from MS-PPT sections 2.7.9 through 2.7.13.
//!
//! Parsing is limited to bytes already present in a caller-supplied PPT record.
//! This module never opens or renders the referenced metafile.

use super::package::{PptError, Result};
use super::records::PptRecord;
use crate::consts::PptRecordType;

const PREFIX_BYTES: usize = 12;
const ENTRY_BYTES: usize = 44;
const VARIANT_BYTES: usize = 34;
const USED_HEADER_FLAGS: u16 = 0x0017;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PowerPointRecolorLimits {
    pub max_record_bytes: usize,
    pub max_entries: usize,
    pub max_trailing_bytes: usize,
}

impl Default for PowerPointRecolorLimits {
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
pub struct PowerPointWideColor {
    pub red: u16,
    pub green: u16,
    pub blue: u16,
}

impl PowerPointWideColor {
    pub const fn new(red: u16, green: u16, blue: u16) -> Self {
        Self { red, green, blue }
    }
}

/// MS-WMF section 2.1.1.4 `BrushStyle`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u16)]
pub enum PowerPointWmfBrushStyle {
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

impl TryFrom<u16> for PowerPointWmfBrushStyle {
    type Error = PptError;

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
pub enum PowerPointWmfHatchStyle {
    Horizontal = 0,
    Vertical = 1,
    ForwardDiagonal = 2,
    BackwardDiagonal = 3,
    Cross = 4,
    DiagonalCross = 5,
}

impl TryFrom<u16> for PowerPointWmfHatchStyle {
    type Error = PptError;

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

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u16)]
pub enum PowerPointRecolorBitmapType {
    MonochromePattern = 0,
    DibPattern = 1,
    NonMonochromeOrIndirect = 3,
}

impl TryFrom<u16> for PowerPointRecolorBitmapType {
    type Error = PptError;

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
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PowerPointRecolorHatch {
    Hatched(PowerPointWmfHatchStyle),
    Ignored(u16),
}

/// Conditional pattern representation. Non-pattern brushes retain ignored bytes.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PowerPointRecolorPattern {
    Pattern {
        bitmap_type: PowerPointRecolorBitmapType,
        bytes: [u8; 8],
    },
    Ignored {
        bitmap_type: u16,
        bytes: [u8; 8],
    },
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PowerPointRecolorBrush {
    pub style: PowerPointWmfBrushStyle,
    pub color: PowerPointWideColor,
    pub hatch: PowerPointRecolorHatch,
    pub foreground_color: PowerPointWideColor,
    pub background_color: PowerPointWideColor,
    pub pattern: PowerPointRecolorPattern,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum PowerPointRecolorSource {
    Color {
        color: PowerPointWideColor,
        /// Undefined bytes retained without interpretation.
        unused: [u8; 26],
    },
    Brush(PowerPointRecolorBrush),
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PowerPointRecolorEntry {
    pub do_recolor: bool,
    /// Destination color; ignored by consumers when `destination_index < 8`.
    pub destination_color: PowerPointWideColor,
    /// Scheme index for values below eight, otherwise `destination_color` is used.
    pub destination_index: u8,
    /// Undefined byte retained without interpretation.
    pub unused: u8,
    pub source: PowerPointRecolorSource,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PowerPointRecolorInfo {
    pub should_recolor: bool,
    pub missing_colors: bool,
    pub missing_fills: bool,
    pub mono_recolor: bool,
    /// Undefined flag bits retained without interpretation.
    pub ignored_flags: u16,
    pub mono_color: PowerPointWideColor,
    pub entries: Vec<PowerPointRecolorEntry>,
    /// Undefined bytes following the fixed-size entry array.
    pub trailing_unused: Vec<u8>,
}

impl PowerPointRecolorInfo {
    pub fn parse(record: &PptRecord, limits: PowerPointRecolorLimits) -> Result<Self> {
        if record.record_type != PptRecordType::RecolorInfoAtom
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

    pub fn parse_payload(payload: &[u8], limits: PowerPointRecolorLimits) -> Result<Self> {
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
                PowerPointRecolorSource::Color { .. } => parsed_colors += 1,
                PowerPointRecolorSource::Brush(_) => parsed_fills += 1,
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

    pub fn to_payload(&self, limits: PowerPointRecolorLimits) -> Result<Vec<u8>> {
        validate(self, limits)?;
        let color_count = self
            .entries
            .iter()
            .filter(|entry| matches!(entry.source, PowerPointRecolorSource::Color { .. }))
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
                .map_err(|_| corrupt("RecolorInfoAtom color count exceeds u16"))?
                .to_le_bytes(),
        );
        output.extend_from_slice(
            &u16::try_from(fill_count)
                .map_err(|_| corrupt("RecolorInfoAtom fill count exceeds u16"))?
                .to_le_bytes(),
        );
        write_color(&mut output, self.mono_color);
        for entry in &self.entries {
            write_entry(&mut output, entry)?;
        }
        output.extend_from_slice(&self.trailing_unused);
        Ok(output)
    }

    pub fn to_record(&self, limits: PowerPointRecolorLimits) -> Result<PptRecord> {
        let data = self.to_payload(limits)?;
        let data_length = u32::try_from(data.len())
            .map_err(|_| corrupt("RecolorInfoAtom payload exceeds u32"))?;
        Ok(PptRecord {
            record_type: PptRecordType::RecolorInfoAtom,
            record_type_raw: 0x0fe7,
            version: 0,
            instance: 0,
            data_length,
            data,
            children: Vec::new(),
        })
    }
}

fn parse_entry(bytes: &[u8]) -> Result<PowerPointRecolorEntry> {
    if bytes.len() != ENTRY_BYTES {
        return corrupted("RecolorEntry has an invalid size");
    }
    let flags = u16_at(bytes, 0)?;
    if flags & !1 != 0 {
        return corrupted("RecolorEntry reserved bits are nonzero");
    }
    let source = parse_source(&bytes[10..])?;
    Ok(PowerPointRecolorEntry {
        do_recolor: flags & 1 != 0,
        destination_color: parse_color(&bytes[2..8])?,
        destination_index: bytes[8],
        unused: bytes[9],
        source,
    })
}

fn parse_source(bytes: &[u8]) -> Result<PowerPointRecolorSource> {
    if bytes.len() != VARIANT_BYTES {
        return corrupted("RecolorEntryVariant has an invalid size");
    }
    match u16_at(bytes, 0)? {
        0 => {
            let mut unused = [0u8; 26];
            unused.copy_from_slice(&bytes[8..34]);
            Ok(PowerPointRecolorSource::Color {
                color: parse_color(&bytes[2..8])?,
                unused,
            })
        },
        1 => {
            let style = PowerPointWmfBrushStyle::try_from(u16_at(bytes, 2)?)?;
            let raw_hatch = u16_at(bytes, 10)?;
            let hatch = if style == PowerPointWmfBrushStyle::Hatched {
                PowerPointRecolorHatch::Hatched(PowerPointWmfHatchStyle::try_from(raw_hatch)?)
            } else {
                PowerPointRecolorHatch::Ignored(raw_hatch)
            };
            let raw_bitmap_type = u16_at(bytes, 24)?;
            let mut pattern_bytes = [0u8; 8];
            pattern_bytes.copy_from_slice(&bytes[26..34]);
            let pattern = if style == PowerPointWmfBrushStyle::Pattern {
                PowerPointRecolorPattern::Pattern {
                    bitmap_type: PowerPointRecolorBitmapType::try_from(raw_bitmap_type)?,
                    bytes: pattern_bytes,
                }
            } else {
                PowerPointRecolorPattern::Ignored {
                    bitmap_type: raw_bitmap_type,
                    bytes: pattern_bytes,
                }
            };
            Ok(PowerPointRecolorSource::Brush(PowerPointRecolorBrush {
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

fn write_entry(output: &mut Vec<u8>, entry: &PowerPointRecolorEntry) -> Result<()> {
    output.extend_from_slice(&u16::from(entry.do_recolor).to_le_bytes());
    write_color(output, entry.destination_color);
    output.push(entry.destination_index);
    output.push(entry.unused);
    match &entry.source {
        PowerPointRecolorSource::Color { color, unused } => {
            output.extend_from_slice(&0u16.to_le_bytes());
            write_color(output, *color);
            output.extend_from_slice(unused);
        },
        PowerPointRecolorSource::Brush(brush) => {
            validate_brush(brush)?;
            output.extend_from_slice(&1u16.to_le_bytes());
            output.extend_from_slice(&(brush.style as u16).to_le_bytes());
            write_color(output, brush.color);
            let hatch = match brush.hatch {
                PowerPointRecolorHatch::Hatched(value) => value as u16,
                PowerPointRecolorHatch::Ignored(value) => value,
            };
            output.extend_from_slice(&hatch.to_le_bytes());
            write_color(output, brush.foreground_color);
            write_color(output, brush.background_color);
            match brush.pattern {
                PowerPointRecolorPattern::Pattern { bitmap_type, bytes } => {
                    output.extend_from_slice(&(bitmap_type as u16).to_le_bytes());
                    output.extend_from_slice(&bytes);
                },
                PowerPointRecolorPattern::Ignored { bitmap_type, bytes } => {
                    output.extend_from_slice(&bitmap_type.to_le_bytes());
                    output.extend_from_slice(&bytes);
                },
            }
        },
    }
    Ok(())
}

fn validate(value: &PowerPointRecolorInfo, limits: PowerPointRecolorLimits) -> Result<()> {
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
        .filter(|entry| matches!(entry.source, PowerPointRecolorSource::Color { .. }))
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
        if let PowerPointRecolorSource::Brush(brush) = &entry.source {
            validate_brush(brush)?;
        }
    }
    Ok(())
}

fn validate_brush(brush: &PowerPointRecolorBrush) -> Result<()> {
    match (brush.style, brush.hatch) {
        (PowerPointWmfBrushStyle::Hatched, PowerPointRecolorHatch::Hatched(_)) => {},
        (PowerPointWmfBrushStyle::Hatched, PowerPointRecolorHatch::Ignored(_)) => {
            return corrupted("hatched RecolorEntryBrush lacks a typed HatchStyle");
        },
        (_, PowerPointRecolorHatch::Hatched(_)) => {
            return corrupted("non-hatched RecolorEntryBrush contains a typed HatchStyle");
        },
        (_, PowerPointRecolorHatch::Ignored(_)) => {},
    }
    match (brush.style, brush.pattern) {
        (PowerPointWmfBrushStyle::Pattern, PowerPointRecolorPattern::Pattern { .. }) => {},
        (PowerPointWmfBrushStyle::Pattern, PowerPointRecolorPattern::Ignored { .. }) => {
            return corrupted("pattern RecolorEntryBrush lacks a typed bitmapType");
        },
        (_, PowerPointRecolorPattern::Pattern { .. }) => {
            return corrupted("non-pattern RecolorEntryBrush contains a typed bitmapType");
        },
        (_, PowerPointRecolorPattern::Ignored { .. }) => {},
    }
    Ok(())
}

fn parse_color(bytes: &[u8]) -> Result<PowerPointWideColor> {
    if bytes.len() != 6 {
        return corrupted("WideColorStruct has an invalid size");
    }
    Ok(PowerPointWideColor {
        red: u16_at(bytes, 0)?,
        green: u16_at(bytes, 2)?,
        blue: u16_at(bytes, 4)?,
    })
}

fn write_color(output: &mut Vec<u8>, color: PowerPointWideColor) {
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

fn corrupt(message: &str) -> PptError {
    PptError::Corrupted(message.to_string())
}

fn corrupted<T>(message: &str) -> Result<T> {
    Err(corrupt(message))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn limits() -> PowerPointRecolorLimits {
        PowerPointRecolorLimits {
            max_record_bytes: 4096,
            max_entries: 8,
            max_trailing_bytes: 16,
        }
    }

    fn color(red: u16, green: u16, blue: u16) -> PowerPointWideColor {
        PowerPointWideColor::new(red, green, blue)
    }

    fn sample() -> PowerPointRecolorInfo {
        PowerPointRecolorInfo {
            should_recolor: true,
            missing_colors: false,
            missing_fills: true,
            mono_recolor: true,
            ignored_flags: 0xa008,
            mono_color: color(1, 2, 3),
            entries: vec![
                PowerPointRecolorEntry {
                    do_recolor: true,
                    destination_color: color(4, 5, 6),
                    destination_index: 8,
                    unused: 0x7f,
                    source: PowerPointRecolorSource::Color {
                        color: color(7, 8, 9),
                        unused: [0x55; 26],
                    },
                },
                PowerPointRecolorEntry {
                    do_recolor: false,
                    destination_color: color(10, 11, 12),
                    destination_index: 2,
                    unused: 0x80,
                    source: PowerPointRecolorSource::Brush(PowerPointRecolorBrush {
                        style: PowerPointWmfBrushStyle::Pattern,
                        color: color(13, 14, 15),
                        hatch: PowerPointRecolorHatch::Ignored(0xbeef),
                        foreground_color: color(16, 17, 18),
                        background_color: color(19, 20, 21),
                        pattern: PowerPointRecolorPattern::Pattern {
                            bitmap_type: PowerPointRecolorBitmapType::DibPattern,
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
        assert_eq!(record.record_type, PptRecordType::RecolorInfoAtom);
        assert_eq!(record.data.len(), 102);
        assert_eq!(&record.data[0..6], &[0x1d, 0xa0, 1, 0, 1, 0]);
        assert_eq!(&record.data[12..14], &[1, 0]);
        assert_eq!(&record.data[22..24], &[0, 0]);
        assert_eq!(&record.data[56..58], &[0, 0]);
        assert_eq!(&record.data[66..68], &[1, 0]);
        assert_eq!(&record.data[68..70], &[3, 0]);
        assert_eq!(&record.data[100..], &[0xde, 0xad]);
        let parsed = PowerPointRecolorInfo::parse(&record, limits()).unwrap();
        assert_eq!(parsed, expected);
        assert_eq!(parsed.to_payload(limits()).unwrap(), record.data);
    }

    #[test]
    fn validates_hatched_and_pattern_conditional_payloads() {
        let mut value = sample();
        {
            let PowerPointRecolorSource::Brush(brush) = &mut value.entries[1].source else {
                unreachable!()
            };
            brush.style = PowerPointWmfBrushStyle::Hatched;
            brush.hatch = PowerPointRecolorHatch::Hatched(PowerPointWmfHatchStyle::DiagonalCross);
            brush.pattern = PowerPointRecolorPattern::Ignored {
                bitmap_type: 0xffff,
                bytes: [0xcc; 8],
            };
        }
        let payload = value.to_payload(limits()).unwrap();
        assert_eq!(
            PowerPointRecolorInfo::parse_payload(&payload, limits()).unwrap(),
            value
        );

        let PowerPointRecolorSource::Brush(brush) = &mut value.entries[1].source else {
            unreachable!()
        };
        brush.hatch = PowerPointRecolorHatch::Ignored(5);
        assert!(value.to_payload(limits()).is_err());
        let PowerPointRecolorSource::Brush(brush) = &mut value.entries[1].source else {
            unreachable!()
        };
        brush.style = PowerPointWmfBrushStyle::Solid;
        brush.hatch = PowerPointRecolorHatch::Hatched(PowerPointWmfHatchStyle::Horizontal);
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
        assert!(PowerPointRecolorInfo::parse_payload(&mutate(12, &[2, 0]), limits()).is_err());
        assert!(PowerPointRecolorInfo::parse_payload(&mutate(22, &[2, 0]), limits()).is_err());
        assert!(PowerPointRecolorInfo::parse_payload(&mutate(2, &[2, 0]), limits()).is_err());
        assert!(PowerPointRecolorInfo::parse_payload(&payload[..99], limits()).is_err());
        assert!(
            PowerPointRecolorInfo::parse_payload(
                &payload,
                PowerPointRecolorLimits {
                    max_record_bytes: 99,
                    ..limits()
                }
            )
            .is_err()
        );
        assert!(
            PowerPointRecolorInfo::parse_payload(
                &payload,
                PowerPointRecolorLimits {
                    max_entries: 1,
                    ..limits()
                }
            )
            .is_err()
        );
        assert!(
            PowerPointRecolorInfo::parse_payload(
                &payload,
                PowerPointRecolorLimits {
                    max_trailing_bytes: 1,
                    ..limits()
                }
            )
            .is_err()
        );

        let mut invalid_style = payload.clone();
        invalid_style[68..70].copy_from_slice(&10u16.to_le_bytes());
        assert!(PowerPointRecolorInfo::parse_payload(&invalid_style, limits()).is_err());
        let mut invalid_bitmap = payload.clone();
        invalid_bitmap[90..92].copy_from_slice(&2u16.to_le_bytes());
        assert!(PowerPointRecolorInfo::parse_payload(&invalid_bitmap, limits()).is_err());
    }
}
