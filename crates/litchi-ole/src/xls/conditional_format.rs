//! Legacy BIFF8 conditional formatting (`CONDFMT` and `CF`).

use crate::xls::error::{XlsError, XlsResult};

pub(crate) const CONDFMT_RECORD_TYPE: u16 = 0x01b0;
pub(crate) const CF_RECORD_TYPE: u16 = 0x01b1;

fn invalid(record_type: u16, message: impl Into<String>) -> XlsError {
    XlsError::InvalidRecord {
        record_type,
        message: message.into(),
    }
}

fn read_u16(data: &[u8], offset: usize) -> u16 {
    u16::from_le_bytes([data[offset], data[offset + 1]])
}

fn read_u32(data: &[u8], offset: usize) -> u32 {
    u32::from_le_bytes(data[offset..offset + 4].try_into().unwrap())
}

/// Inclusive worksheet range affected by conditional formatting.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XlsConditionalFormatRange {
    first_row: u16,
    last_row: u16,
    first_column: u8,
    last_column: u8,
}

impl XlsConditionalFormatRange {
    pub fn first_row(&self) -> u16 { self.first_row }
    pub fn last_row(&self) -> u16 { self.last_row }
    pub fn first_column(&self) -> u8 { self.first_column }
    pub fn last_column(&self) -> u8 { self.last_column }
}

/// Comparison performed by a cell-value conditional formatting rule.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlsConditionalComparison {
    Between,
    NotBetween,
    Equal,
    NotEqual,
    GreaterThan,
    LessThan,
    GreaterThanOrEqual,
    LessThanOrEqual,
}

/// Type of condition used by a legacy rule.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlsConditionalRuleKind {
    CellValue(XlsConditionalComparison),
    Formula,
}

/// Conditional number-format override.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum XlsConditionalNumberFormat {
    Identifier(u8),
    Custom(String),
}

/// Raw BIFF font differential block with common typed properties.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsConditionalFont {
    raw: Vec<u8>,
    name: Option<String>,
}

impl XlsConditionalFont {
    pub fn name(&self) -> Option<&str> { self.name.as_deref() }
    pub fn height_twips(&self) -> Option<u32> {
        let value = read_u32(&self.raw, 64);
        (value != u32::MAX).then_some(value)
    }
    pub fn is_italic(&self) -> bool { read_u32(&self.raw, 68) & 0x0002 != 0 }
    pub fn is_outline(&self) -> bool { read_u32(&self.raw, 68) & 0x0008 != 0 }
    pub fn has_shadow(&self) -> bool { read_u32(&self.raw, 68) & 0x0010 != 0 }
    pub fn is_struck_out(&self) -> bool { read_u32(&self.raw, 68) & 0x0080 != 0 }
    pub fn weight(&self) -> u16 { read_u16(&self.raw, 72) }
    pub fn escapement(&self) -> u16 { read_u16(&self.raw, 74) }
    pub fn underline(&self) -> u8 { self.raw[76] }
    pub fn color_index(&self) -> i32 { read_u32(&self.raw, 80) as i32 }
    pub fn raw_data(&self) -> &[u8] { &self.raw }
}

/// Text alignment differential block.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XlsConditionalAlignment {
    horizontal: u8,
    vertical: u8,
    wrap_text: bool,
    rotation: u8,
    absolute_indent: u8,
    relative_indent: i32,
    shrink_to_fit: bool,
    merge_cell: bool,
    reading_order: u8,
}

impl XlsConditionalAlignment {
    pub fn horizontal(&self) -> u8 { self.horizontal }
    pub fn vertical(&self) -> u8 { self.vertical }
    pub fn wraps_text(&self) -> bool { self.wrap_text }
    pub fn rotation(&self) -> u8 { self.rotation }
    pub fn absolute_indent(&self) -> u8 { self.absolute_indent }
    pub fn relative_indent(&self) -> i32 { self.relative_indent }
    pub fn shrinks_to_fit(&self) -> bool { self.shrink_to_fit }
    pub fn merges_cell(&self) -> bool { self.merge_cell }
    pub fn reading_order(&self) -> u8 { self.reading_order }
}

/// Cell border differential block.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XlsConditionalBorder {
    styles: [u8; 5],
    colors: [u8; 5],
    diagonal_down: bool,
    diagonal_up: bool,
}

impl XlsConditionalBorder {
    /// Left, right, top, bottom, and diagonal styles.
    pub fn styles(&self) -> &[u8; 5] { &self.styles }
    /// Left, right, top, bottom, and diagonal color indexes.
    pub fn color_indexes(&self) -> &[u8; 5] { &self.colors }
    pub fn has_diagonal_down(&self) -> bool { self.diagonal_down }
    pub fn has_diagonal_up(&self) -> bool { self.diagonal_up }
}

/// Fill pattern differential block.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XlsConditionalPattern {
    fill_pattern: u8,
    foreground_color_index: u8,
    background_color_index: u8,
}

impl XlsConditionalPattern {
    pub fn fill_pattern(&self) -> u8 { self.fill_pattern }
    pub fn foreground_color_index(&self) -> u8 { self.foreground_color_index }
    pub fn background_color_index(&self) -> u8 { self.background_color_index }
}

/// Cell protection differential block.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XlsConditionalProtection {
    locked: bool,
    hidden: bool,
}

impl XlsConditionalProtection {
    pub fn is_locked(&self) -> bool { self.locked }
    pub fn is_hidden(&self) -> bool { self.hidden }
}

/// Differential formatting applied when a rule evaluates to true.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsConditionalStyle {
    options: u32,
    new_border: bool,
    number_format: Option<XlsConditionalNumberFormat>,
    font: Option<XlsConditionalFont>,
    alignment: Option<XlsConditionalAlignment>,
    border: Option<XlsConditionalBorder>,
    pattern: Option<XlsConditionalPattern>,
    protection: Option<XlsConditionalProtection>,
}

impl XlsConditionalStyle {
    pub fn number_format(&self) -> Option<&XlsConditionalNumberFormat> { self.number_format.as_ref() }
    pub fn font(&self) -> Option<&XlsConditionalFont> { self.font.as_ref() }
    pub fn alignment(&self) -> Option<&XlsConditionalAlignment> { self.alignment.as_ref() }
    pub fn border(&self) -> Option<&XlsConditionalBorder> { self.border.as_ref() }
    pub fn pattern(&self) -> Option<&XlsConditionalPattern> { self.pattern.as_ref() }
    pub fn protection(&self) -> Option<&XlsConditionalProtection> { self.protection.as_ref() }
    pub fn applies_border_to_range_outline(&self) -> bool { self.new_border }
    pub fn is_pattern_style_modified(&self) -> bool { self.options & 0x0001_0000 == 0 }
    pub fn is_pattern_foreground_modified(&self) -> bool { self.options & 0x0002_0000 == 0 }
    pub fn is_pattern_background_modified(&self) -> bool { self.options & 0x0004_0000 == 0 }
}

/// One legacy conditional formatting rule.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsConditionalRule {
    kind: XlsConditionalRuleKind,
    style: XlsConditionalStyle,
    formula1_tokens: Vec<u8>,
    formula2_tokens: Vec<u8>,
}

impl XlsConditionalRule {
    pub fn kind(&self) -> XlsConditionalRuleKind { self.kind }
    pub fn style(&self) -> &XlsConditionalStyle { &self.style }
    pub fn formula1_tokens(&self) -> &[u8] { &self.formula1_tokens }
    pub fn formula2_tokens(&self) -> &[u8] { &self.formula2_tokens }
}

/// A range set and its one-to-three legacy conditional formatting rules.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsConditionalFormatting {
    identifier: u16,
    tough_recalculation: bool,
    enclosing_range: XlsConditionalFormatRange,
    ranges: Vec<XlsConditionalFormatRange>,
    rules: Vec<XlsConditionalRule>,
}

impl XlsConditionalFormatting {
    pub fn identifier(&self) -> u16 { self.identifier }
    pub fn requires_tough_recalculation(&self) -> bool { self.tough_recalculation }
    pub fn enclosing_range(&self) -> XlsConditionalFormatRange { self.enclosing_range }
    pub fn ranges(&self) -> &[XlsConditionalFormatRange] { &self.ranges }
    pub fn rules(&self) -> &[XlsConditionalRule] { &self.rules }
}

fn parse_range(data: &[u8], record_type: u16) -> XlsResult<XlsConditionalFormatRange> {
    let first_row = read_u16(data, 0);
    let last_row = read_u16(data, 2);
    let first_column = read_u16(data, 4);
    let last_column = read_u16(data, 6);
    if first_row > last_row || first_column > last_column || last_column > 255 {
        return Err(invalid(record_type, "conditional formatting range is invalid"));
    }
    Ok(XlsConditionalFormatRange {
        first_row,
        last_row,
        first_column: first_column as u8,
        last_column: last_column as u8,
    })
}

struct PendingFormatting {
    group: XlsConditionalFormatting,
    declared_rules: usize,
}

fn parse_condfmt(data: &[u8]) -> XlsResult<PendingFormatting> {
    if data.len() < 14 {
        return Err(invalid(CONDFMT_RECORD_TYPE, "CONDFMT payload is shorter than 14 bytes"));
    }
    let declared_rules = usize::from(read_u16(data, 0));
    if !(1..=3).contains(&declared_rules) {
        return Err(invalid(CONDFMT_RECORD_TYPE, "CONDFMT rule count must be between 1 and 3"));
    }
    let flags = read_u16(data, 2);
    let enclosing_range = parse_range(&data[4..12], CONDFMT_RECORD_TYPE)?;
    let range_count = usize::from(read_u16(data, 12));
    if !(1..=1026).contains(&range_count) || data.len() != 14 + range_count * 8 {
        return Err(invalid(CONDFMT_RECORD_TYPE, "CONDFMT range count does not match its payload"));
    }
    let mut ranges = Vec::with_capacity(range_count);
    for chunk in data[14..].chunks_exact(8) {
        let range = parse_range(chunk, CONDFMT_RECORD_TYPE)?;
        if range.first_row < enclosing_range.first_row
            || range.last_row > enclosing_range.last_row
            || range.first_column < enclosing_range.first_column
            || range.last_column > enclosing_range.last_column
        {
            return Err(invalid(CONDFMT_RECORD_TYPE, "CONDFMT enclosing range does not contain every target range"));
        }
        ranges.push(range);
    }
    Ok(PendingFormatting {
        group: XlsConditionalFormatting {
            identifier: flags >> 1,
            tough_recalculation: flags & 1 != 0,
            enclosing_range,
            ranges,
            rules: Vec::with_capacity(declared_rules),
        },
        declared_rules,
    })
}

fn parse_simple_xl_unicode(data: &[u8], record_type: u16) -> XlsResult<String> {
    if data.len() < 3 {
        return Err(invalid(record_type, "truncated differential number-format string"));
    }
    let count = usize::from(read_u16(data, 0));
    let flags = data[2];
    if flags & 0xfe != 0 {
        return Err(invalid(record_type, "differential number-format string has reserved flags"));
    }
    let width = if flags & 1 != 0 { 2 } else { 1 };
    if data.len() != 3 + count * width {
        return Err(invalid(record_type, "differential number-format string length mismatch"));
    }
    if width == 1 {
        Ok(data[3..].iter().map(|&byte| char::from(byte)).collect())
    } else {
        let units = data[3..]
            .chunks_exact(2)
            .map(|bytes| u16::from_le_bytes([bytes[0], bytes[1]]))
            .collect::<Vec<_>>();
        String::from_utf16(&units).map_err(|_| invalid(record_type, "invalid UTF-16 number format"))
    }
}

fn take<'a>(data: &'a [u8], offset: &mut usize, length: usize, name: &str) -> XlsResult<&'a [u8]> {
    let bytes = data
        .get(*offset..*offset + length)
        .ok_or_else(|| invalid(CF_RECORD_TYPE, format!("truncated {name} differential block")))?;
    *offset += length;
    Ok(bytes)
}

fn parse_font(data: &[u8]) -> XlsResult<XlsConditionalFont> {
    let count = usize::from(data[0]);
    let name = if count == 0 {
        None
    } else {
        let flags = data[1];
        if flags & 0xfe != 0 {
            return Err(invalid(CF_RECORD_TYPE, "conditional font name has reserved flags"));
        }
        let width = if flags & 1 != 0 { 2 } else { 1 };
        let byte_count = count * width;
        if 2 + byte_count > 64 || (width == 1 && count > 62) || (width == 2 && count > 31) {
            return Err(invalid(CF_RECORD_TYPE, "conditional font name exceeds its fixed block"));
        }
        let chars = &data[2..2 + byte_count];
        Some(if width == 1 {
            chars.iter().map(|&byte| char::from(byte)).collect()
        } else {
            let units = chars
                .chunks_exact(2)
                .map(|bytes| u16::from_le_bytes([bytes[0], bytes[1]]))
                .collect::<Vec<_>>();
            String::from_utf16(&units).map_err(|_| invalid(CF_RECORD_TYPE, "invalid UTF-16 conditional font name"))?
        })
    };
    Ok(XlsConditionalFont { raw: data.to_vec(), name })
}

fn parse_style(data: &[u8]) -> XlsResult<(XlsConditionalStyle, usize)> {
    if data.len() < 6 {
        return Err(invalid(CF_RECORD_TYPE, "CF differential formatting header is truncated"));
    }
    let options = read_u32(data, 0);
    let secondary = read_u16(data, 4);
    if options & 0x01c0_0000 != 0 || secondary & 0x7ff8 != 0 {
        return Err(invalid(CF_RECORD_TYPE, "CF differential formatting has nonzero reserved bits"));
    }
    let mut offset = 6usize;
    let number_format = if options & 0x0200_0000 != 0 {
        if secondary & 1 != 0 {
            let length_bytes = take(data, &mut offset, 2, "number format")?;
            let length = usize::from(read_u16(length_bytes, 0));
            if length < 2 {
                return Err(invalid(CF_RECORD_TYPE, "custom differential number format is too short"));
            }
            let rest = take(data, &mut offset, length - 2, "number format")?;
            Some(XlsConditionalNumberFormat::Custom(parse_simple_xl_unicode(rest, CF_RECORD_TYPE)?))
        } else {
            let bytes = take(data, &mut offset, 2, "number format")?;
            Some(XlsConditionalNumberFormat::Identifier(bytes[1]))
        }
    } else {
        None
    };
    let font = if options & 0x0400_0000 != 0 {
        Some(parse_font(take(data, &mut offset, 118, "font")?)?)
    } else {
        None
    };
    let alignment = if options & 0x0800_0000 != 0 {
        let bytes = take(data, &mut offset, 8, "alignment")?;
        let relative_indent = read_u32(bytes, 4) as i32;
        if !(-15..=255).contains(&relative_indent) {
            return Err(invalid(CF_RECORD_TYPE, "conditional relative indent is outside -15 through 255"));
        }
        Some(XlsConditionalAlignment {
            horizontal: bytes[0] & 7,
            vertical: (bytes[0] >> 4) & 7,
            wrap_text: bytes[0] & 8 != 0,
            rotation: bytes[1],
            absolute_indent: bytes[2] & 15,
            relative_indent,
            shrink_to_fit: bytes[2] & 0x10 != 0,
            merge_cell: bytes[2] & 0x20 != 0,
            reading_order: bytes[2] >> 6,
        })
    } else {
        None
    };
    let border = if options & 0x1000_0000 != 0 {
        let bytes = take(data, &mut offset, 8, "border")?;
        let first = read_u32(bytes, 0);
        let second = read_u32(bytes, 4);
        Some(XlsConditionalBorder {
            styles: [
                (first & 15) as u8,
                ((first >> 4) & 15) as u8,
                ((first >> 8) & 15) as u8,
                ((first >> 12) & 15) as u8,
                ((second >> 21) & 15) as u8,
            ],
            colors: [
                ((first >> 16) & 0x7f) as u8,
                ((first >> 23) & 0x7f) as u8,
                (second & 0x7f) as u8,
                ((second >> 7) & 0x7f) as u8,
                ((second >> 14) & 0x7f) as u8,
            ],
            diagonal_down: first & 0x4000_0000 != 0,
            diagonal_up: first & 0x8000_0000 != 0,
        })
    } else {
        None
    };
    let pattern = if options & 0x2000_0000 != 0 {
        let bytes = take(data, &mut offset, 4, "pattern")?;
        let style = read_u16(bytes, 0);
        let colors = read_u16(bytes, 2);
        Some(XlsConditionalPattern {
            fill_pattern: (style >> 10) as u8,
            foreground_color_index: (colors & 0x7f) as u8,
            background_color_index: ((colors >> 7) & 0x7f) as u8,
        })
    } else {
        None
    };
    let protection = if options & 0x4000_0000 != 0 {
        let bits = read_u16(take(data, &mut offset, 2, "protection")?, 0);
        if bits & !3 != 0 {
            return Err(invalid(CF_RECORD_TYPE, "conditional protection has nonzero reserved bits"));
        }
        Some(XlsConditionalProtection {
            locked: bits & 1 != 0,
            hidden: bits & 2 != 0,
        })
    } else {
        None
    };
    Ok((XlsConditionalStyle {
        options,
        new_border: secondary & 4 != 0,
        number_format,
        font,
        alignment,
        border,
        pattern,
        protection,
    }, offset))
}

fn parse_cf(data: &[u8]) -> XlsResult<XlsConditionalRule> {
    if data.len() < 12 {
        return Err(invalid(CF_RECORD_TYPE, "CF payload is shorter than 12 bytes"));
    }
    let formula1_len = usize::from(read_u16(data, 2));
    let formula2_len = usize::from(read_u16(data, 4));
    if formula1_len > 16409 || formula2_len > 16409 {
        return Err(invalid(CF_RECORD_TYPE, "CF formula exceeds 16409 bytes"));
    }
    let kind = match (data[0], data[1]) {
        (1, operator @ 1..=8) => XlsConditionalRuleKind::CellValue(match operator {
            1 => XlsConditionalComparison::Between,
            2 => XlsConditionalComparison::NotBetween,
            3 => XlsConditionalComparison::Equal,
            4 => XlsConditionalComparison::NotEqual,
            5 => XlsConditionalComparison::GreaterThan,
            6 => XlsConditionalComparison::LessThan,
            7 => XlsConditionalComparison::GreaterThanOrEqual,
            _ => XlsConditionalComparison::LessThanOrEqual,
        }),
        (2, 0) => XlsConditionalRuleKind::Formula,
        (1, _) => return Err(invalid(CF_RECORD_TYPE, "cell-value CF operator must be between 1 and 8")),
        (2, _) => return Err(invalid(CF_RECORD_TYPE, "formula CF operator must be zero")),
        _ => return Err(invalid(CF_RECORD_TYPE, "legacy CF condition type must be 1 or 2")),
    };
    if matches!(kind, XlsConditionalRuleKind::Formula) && formula2_len != 0 {
        return Err(invalid(CF_RECORD_TYPE, "formula CF rule cannot contain a second formula"));
    }
    if matches!(kind, XlsConditionalRuleKind::CellValue(operator) if !matches!(operator, XlsConditionalComparison::Between | XlsConditionalComparison::NotBetween))
        && formula2_len != 0
    {
        return Err(invalid(CF_RECORD_TYPE, "single-operand CF comparison cannot contain a second formula"));
    }
    let (style, style_len) = parse_style(&data[6..])?;
    let formula_offset = 6 + style_len;
    if data.len() != formula_offset + formula1_len + formula2_len {
        return Err(invalid(CF_RECORD_TYPE, "CF formula lengths do not match the record payload"));
    }
    Ok(XlsConditionalRule {
        kind,
        style,
        formula1_tokens: data[formula_offset..formula_offset + formula1_len].to_vec(),
        formula2_tokens: data[formula_offset + formula1_len..].to_vec(),
    })
}

/// Enforces the `CondFmt 1*3CF` collection grammar.
pub(crate) struct ConditionalFormatCollector {
    groups: Vec<XlsConditionalFormatting>,
    pending: Option<PendingFormatting>,
}

impl ConditionalFormatCollector {
    pub(crate) fn new() -> Self {
        Self { groups: Vec::new(), pending: None }
    }

    pub(crate) fn feed_record(&mut self, record_type: u16, data: &[u8]) -> XlsResult<()> {
        if self.pending.is_some() && record_type != CF_RECORD_TYPE {
            return Err(invalid(record_type, "CONDFMT must be followed immediately by its declared CF records"));
        }
        match record_type {
            CONDFMT_RECORD_TYPE => {
                self.pending = Some(parse_condfmt(data)?);
            }
            CF_RECORD_TYPE => {
                let pending = self.pending.as_mut().ok_or_else(|| invalid(record_type, "orphan CF record without CONDFMT"))?;
                pending.group.rules.push(parse_cf(data)?);
                if pending.group.rules.len() == pending.declared_rules {
                    self.groups.push(self.pending.take().unwrap().group);
                }
            }
            _ => {}
        }
        Ok(())
    }

    pub(crate) fn finish(self) -> XlsResult<Vec<XlsConditionalFormatting>> {
        if self.pending.is_some() {
            Err(invalid(CONDFMT_RECORD_TYPE, "worksheet ended before all declared CF rules were read"))
        } else {
            Ok(self.groups)
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn header(rule_count: u16) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&rule_count.to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(&[0, 0, 7, 0, 0, 0, 0, 0]);
        data.extend_from_slice(&1u16.to_le_bytes());
        data.extend_from_slice(&[0, 0, 7, 0, 0, 0, 0, 0]);
        data
    }

    fn rule(condition: u8, operator: u8, formula1: &[u8], formula2: &[u8]) -> Vec<u8> {
        let mut data = vec![condition, operator];
        data.extend_from_slice(&(formula1.len() as u16).to_le_bytes());
        data.extend_from_slice(&(formula2.len() as u16).to_le_bytes());
        data.extend_from_slice(&0x003f_ffffu32.to_le_bytes());
        data.extend_from_slice(&0x8002u16.to_le_bytes());
        data.extend_from_slice(formula1);
        data.extend_from_slice(formula2);
        data
    }

    #[test]
    fn parses_and_collects_legacy_rules() {
        let mut collector = ConditionalFormatCollector::new();
        collector.feed_record(CONDFMT_RECORD_TYPE, &header(2)).unwrap();
        collector.feed_record(CF_RECORD_TYPE, &rule(1, 1, &[0x1e, 1, 0], &[0x1e, 5, 0])).unwrap();
        collector.feed_record(CF_RECORD_TYPE, &rule(2, 0, &[0x1d], &[])).unwrap();
        let groups = collector.finish().unwrap();
        assert_eq!(groups[0].rules().len(), 2);
        assert_eq!(groups[0].rules()[0].kind(), XlsConditionalRuleKind::CellValue(XlsConditionalComparison::Between));
        assert_eq!(groups[0].rules()[1].formula1_tokens(), &[0x1d]);
    }

    #[test]
    fn rejects_malformed_ranges_formulas_and_sequences() {
        assert!(parse_condfmt(&header(0)).is_err());
        assert!(parse_cf(&rule(2, 1, &[0x1d], &[])).is_err());
        assert!(parse_cf(&rule(1, 5, &[0x1e, 1, 0], &[0x1e, 2, 0])).is_err());

        let mut collector = ConditionalFormatCollector::new();
        assert!(collector.feed_record(CF_RECORD_TYPE, &rule(2, 0, &[0x1d], &[])).is_err());
        let mut collector = ConditionalFormatCollector::new();
        collector.feed_record(CONDFMT_RECORD_TYPE, &header(1)).unwrap();
        assert!(collector.feed_record(0x000a, &[]).is_err());
    }

    #[test]
    fn reads_poi_legacy_conditional_formatting_fixture() {
        use crate::xls::XlsWorkbook;
        use std::fs::File;
        use std::path::Path;

        let fixture = Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../3rdparty/poi/test-data/spreadsheet/WithConditionalFormatting.xls");
        let workbook = XlsWorkbook::new(File::open(fixture).unwrap()).unwrap();
        let groups = workbook.xls_worksheet(0).unwrap().conditional_formattings();
        assert_eq!(groups.len(), 3);
        assert_eq!(groups[0].rules().len(), 2);
        assert_eq!(groups[0].ranges()[0].last_row(), 7);
        assert_eq!(groups[0].rules()[0].kind(), XlsConditionalRuleKind::CellValue(XlsConditionalComparison::GreaterThan));
        assert!(groups[0].rules()[0].style().font().is_some());
        assert_eq!(groups[1].rules()[0].kind(), XlsConditionalRuleKind::Formula);
        assert_eq!(groups[2].rules()[1].kind(), XlsConditionalRuleKind::CellValue(XlsConditionalComparison::Between));
        assert!(!groups[2].rules()[1].formula2_tokens().is_empty());
    }
}
