//! Legacy BIFF8 conditional formatting (`CONDFMT` and `CF`).

use crate::error::{XlsError, XlsResult};
use crate::formula::{FormulaContext, render_formula};
use std::collections::HashSet;

pub(crate) const CONDFMT_RECORD_TYPE: u16 = 0x01b0;
pub(crate) const CF_RECORD_TYPE: u16 = 0x01b1;
pub(crate) const CONDFMT12_RECORD_TYPE: u16 = 0x0879;
pub(crate) const CF12_RECORD_TYPE: u16 = 0x087a;
pub(crate) const CFEX_RECORD_TYPE: u16 = 0x087b;

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
    pub fn first_row(&self) -> u16 {
        self.first_row
    }
    pub fn last_row(&self) -> u16 {
        self.last_row
    }
    pub fn first_column(&self) -> u8 {
        self.first_column
    }
    pub fn last_column(&self) -> u8 {
        self.last_column
    }
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
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }
    pub fn height_twips(&self) -> Option<u32> {
        let value = read_u32(&self.raw, 64);
        (value != u32::MAX).then_some(value)
    }
    pub fn is_italic(&self) -> bool {
        read_u32(&self.raw, 68) & 0x0002 != 0
    }
    pub fn is_outline(&self) -> bool {
        read_u32(&self.raw, 68) & 0x0008 != 0
    }
    pub fn has_shadow(&self) -> bool {
        read_u32(&self.raw, 68) & 0x0010 != 0
    }
    pub fn is_struck_out(&self) -> bool {
        read_u32(&self.raw, 68) & 0x0080 != 0
    }
    pub fn weight(&self) -> u16 {
        read_u16(&self.raw, 72)
    }
    pub fn escapement(&self) -> u16 {
        read_u16(&self.raw, 74)
    }
    pub fn underline(&self) -> u8 {
        self.raw[76]
    }
    pub fn color_index(&self) -> i32 {
        read_u32(&self.raw, 80) as i32
    }
    pub fn raw_data(&self) -> &[u8] {
        &self.raw
    }
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
    pub fn horizontal(&self) -> u8 {
        self.horizontal
    }
    pub fn vertical(&self) -> u8 {
        self.vertical
    }
    pub fn wraps_text(&self) -> bool {
        self.wrap_text
    }
    pub fn rotation(&self) -> u8 {
        self.rotation
    }
    pub fn absolute_indent(&self) -> u8 {
        self.absolute_indent
    }
    pub fn relative_indent(&self) -> i32 {
        self.relative_indent
    }
    pub fn shrinks_to_fit(&self) -> bool {
        self.shrink_to_fit
    }
    pub fn merges_cell(&self) -> bool {
        self.merge_cell
    }
    pub fn reading_order(&self) -> u8 {
        self.reading_order
    }
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
    pub fn styles(&self) -> &[u8; 5] {
        &self.styles
    }
    /// Left, right, top, bottom, and diagonal color indexes.
    pub fn color_indexes(&self) -> &[u8; 5] {
        &self.colors
    }
    pub fn has_diagonal_down(&self) -> bool {
        self.diagonal_down
    }
    pub fn has_diagonal_up(&self) -> bool {
        self.diagonal_up
    }
}

/// Fill pattern differential block.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XlsConditionalPattern {
    fill_pattern: u8,
    foreground_color_index: u8,
    background_color_index: u8,
}

impl XlsConditionalPattern {
    pub fn fill_pattern(&self) -> u8 {
        self.fill_pattern
    }
    pub fn foreground_color_index(&self) -> u8 {
        self.foreground_color_index
    }
    pub fn background_color_index(&self) -> u8 {
        self.background_color_index
    }
}

/// Cell protection differential block.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XlsConditionalProtection {
    locked: bool,
    hidden: bool,
}

impl XlsConditionalProtection {
    pub fn is_locked(&self) -> bool {
        self.locked
    }
    pub fn is_hidden(&self) -> bool {
        self.hidden
    }
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
    pub fn number_format(&self) -> Option<&XlsConditionalNumberFormat> {
        self.number_format.as_ref()
    }
    pub fn font(&self) -> Option<&XlsConditionalFont> {
        self.font.as_ref()
    }
    pub fn alignment(&self) -> Option<&XlsConditionalAlignment> {
        self.alignment.as_ref()
    }
    pub fn border(&self) -> Option<&XlsConditionalBorder> {
        self.border.as_ref()
    }
    pub fn pattern(&self) -> Option<&XlsConditionalPattern> {
        self.pattern.as_ref()
    }
    pub fn protection(&self) -> Option<&XlsConditionalProtection> {
        self.protection.as_ref()
    }
    pub fn applies_border_to_range_outline(&self) -> bool {
        self.new_border
    }
    pub fn is_pattern_style_modified(&self) -> bool {
        self.options & 0x0001_0000 == 0
    }
    pub fn is_pattern_foreground_modified(&self) -> bool {
        self.options & 0x0002_0000 == 0
    }
    pub fn is_pattern_background_modified(&self) -> bool {
        self.options & 0x0004_0000 == 0
    }
}

/// One legacy conditional formatting rule.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsConditionalRule {
    kind: XlsConditionalRuleKind,
    style: XlsConditionalStyle,
    formula1_tokens: Vec<u8>,
    formula2_tokens: Vec<u8>,
    formula1_rendered: Option<String>,
    formula2_rendered: Option<String>,
}

impl XlsConditionalRule {
    pub fn kind(&self) -> XlsConditionalRuleKind {
        self.kind
    }
    pub fn style(&self) -> &XlsConditionalStyle {
        &self.style
    }
    pub fn formula1_tokens(&self) -> &[u8] {
        &self.formula1_tokens
    }
    pub fn formula2_tokens(&self) -> &[u8] {
        &self.formula2_tokens
    }
    pub fn formula1_rendered(&self) -> Option<&str> {
        self.formula1_rendered.as_deref()
    }
    pub fn formula2_rendered(&self) -> Option<&str> {
        self.formula2_rendered.as_deref()
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlsConditionalRule12Kind {
    CellValue(XlsConditionalComparison),
    Formula,
    ColorScale,
    DataBar,
    Filter,
    IconSet,
}

/// Office 2007 future conditional-formatting rule. Visual payloads remain inert bytes.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsConditionalRule12 {
    kind: XlsConditionalRule12Kind,
    priority: u16,
    stop_if_true: bool,
    template: u16,
    differential_format: Vec<u8>,
    formula1_tokens: Vec<u8>,
    formula2_tokens: Vec<u8>,
    active_formula_tokens: Vec<u8>,
    formula1_rendered: Option<String>,
    formula2_rendered: Option<String>,
    active_formula_rendered: Option<String>,
    template_parameters: [u8; 16],
    rule_payload: Vec<u8>,
}
impl XlsConditionalRule12 {
    pub fn kind(&self) -> XlsConditionalRule12Kind {
        self.kind
    }
    pub fn priority(&self) -> u16 {
        self.priority
    }
    pub fn stop_if_true(&self) -> bool {
        self.stop_if_true
    }
    pub fn template(&self) -> u16 {
        self.template
    }
    pub fn differential_format(&self) -> &[u8] {
        &self.differential_format
    }
    pub fn formula1_tokens(&self) -> &[u8] {
        &self.formula1_tokens
    }
    pub fn formula2_tokens(&self) -> &[u8] {
        &self.formula2_tokens
    }
    pub fn active_formula_tokens(&self) -> &[u8] {
        &self.active_formula_tokens
    }
    pub fn formula1_rendered(&self) -> Option<&str> {
        self.formula1_rendered.as_deref()
    }
    pub fn formula2_rendered(&self) -> Option<&str> {
        self.formula2_rendered.as_deref()
    }
    pub fn active_formula_rendered(&self) -> Option<&str> {
        self.active_formula_rendered.as_deref()
    }
    pub fn template_parameters(&self) -> &[u8; 16] {
        &self.template_parameters
    }
    pub fn rule_payload(&self) -> &[u8] {
        &self.rule_payload
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsConditionalFormatting12 {
    identifier: u16,
    tough_recalculation: bool,
    enclosing_range: XlsConditionalFormatRange,
    ranges: Vec<XlsConditionalFormatRange>,
    rules: Vec<XlsConditionalRule12>,
}
impl XlsConditionalFormatting12 {
    pub fn identifier(&self) -> u16 {
        self.identifier
    }
    pub fn requires_tough_recalculation(&self) -> bool {
        self.tough_recalculation
    }
    pub fn enclosing_range(&self) -> XlsConditionalFormatRange {
        self.enclosing_range
    }
    pub fn ranges(&self) -> &[XlsConditionalFormatRange] {
        &self.ranges
    }
    pub fn rules(&self) -> &[XlsConditionalRule12] {
        &self.rules
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsConditionalExtension {
    identifier: u16,
    legacy_rule_index: Option<u16>,
    priority: u16,
    active: bool,
    stop_if_true: bool,
    template: u8,
    differential_format: Vec<u8>,
    template_parameters: [u8; 16],
    future_rule: Option<XlsConditionalRule12>,
}
impl XlsConditionalExtension {
    pub fn identifier(&self) -> u16 {
        self.identifier
    }
    pub fn legacy_rule_index(&self) -> Option<u16> {
        self.legacy_rule_index
    }
    pub fn priority(&self) -> u16 {
        self.priority
    }
    pub fn active(&self) -> bool {
        self.active
    }
    pub fn stop_if_true(&self) -> bool {
        self.stop_if_true
    }
    pub fn template(&self) -> u8 {
        self.template
    }
    pub fn differential_format(&self) -> &[u8] {
        &self.differential_format
    }
    pub fn template_parameters(&self) -> &[u8; 16] {
        &self.template_parameters
    }
    pub fn future_rule(&self) -> Option<&XlsConditionalRule12> {
        self.future_rule.as_ref()
    }
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
    pub fn identifier(&self) -> u16 {
        self.identifier
    }
    pub fn requires_tough_recalculation(&self) -> bool {
        self.tough_recalculation
    }
    pub fn enclosing_range(&self) -> XlsConditionalFormatRange {
        self.enclosing_range
    }
    pub fn ranges(&self) -> &[XlsConditionalFormatRange] {
        &self.ranges
    }
    pub fn rules(&self) -> &[XlsConditionalRule] {
        &self.rules
    }
}

fn parse_range(data: &[u8], record_type: u16) -> XlsResult<XlsConditionalFormatRange> {
    let first_row = read_u16(data, 0);
    let last_row = read_u16(data, 2);
    let first_column = read_u16(data, 4);
    let last_column = read_u16(data, 6);
    if first_row > last_row || first_column > last_column || last_column > 255 {
        return Err(invalid(
            record_type,
            "conditional formatting range is invalid",
        ));
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
        return Err(invalid(
            CONDFMT_RECORD_TYPE,
            "CONDFMT payload is shorter than 14 bytes",
        ));
    }
    let declared_rules = usize::from(read_u16(data, 0));
    if !(1..=3).contains(&declared_rules) {
        return Err(invalid(
            CONDFMT_RECORD_TYPE,
            "CONDFMT rule count must be between 1 and 3",
        ));
    }
    let flags = read_u16(data, 2);
    let enclosing_range = parse_range(&data[4..12], CONDFMT_RECORD_TYPE)?;
    let range_count = usize::from(read_u16(data, 12));
    if !(1..=1026).contains(&range_count) || data.len() != 14 + range_count * 8 {
        return Err(invalid(
            CONDFMT_RECORD_TYPE,
            "CONDFMT range count does not match its payload",
        ));
    }
    let mut ranges = Vec::with_capacity(range_count);
    for chunk in data[14..].chunks_exact(8) {
        let range = parse_range(chunk, CONDFMT_RECORD_TYPE)?;
        if range.first_row < enclosing_range.first_row
            || range.last_row > enclosing_range.last_row
            || range.first_column < enclosing_range.first_column
            || range.last_column > enclosing_range.last_column
        {
            return Err(invalid(
                CONDFMT_RECORD_TYPE,
                "CONDFMT enclosing range does not contain every target range",
            ));
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
        return Err(invalid(
            record_type,
            "truncated differential number-format string",
        ));
    }
    let count = usize::from(read_u16(data, 0));
    let flags = data[2];
    if flags & 0xfe != 0 {
        return Err(invalid(
            record_type,
            "differential number-format string has reserved flags",
        ));
    }
    let width = if flags & 1 != 0 { 2 } else { 1 };
    if data.len() != 3 + count * width {
        return Err(invalid(
            record_type,
            "differential number-format string length mismatch",
        ));
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
    let bytes = data.get(*offset..*offset + length).ok_or_else(|| {
        invalid(
            CF_RECORD_TYPE,
            format!("truncated {name} differential block"),
        )
    })?;
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
            return Err(invalid(
                CF_RECORD_TYPE,
                "conditional font name has reserved flags",
            ));
        }
        let width = if flags & 1 != 0 { 2 } else { 1 };
        let byte_count = count * width;
        if 2 + byte_count > 64 || (width == 1 && count > 62) || (width == 2 && count > 31) {
            return Err(invalid(
                CF_RECORD_TYPE,
                "conditional font name exceeds its fixed block",
            ));
        }
        let chars = &data[2..2 + byte_count];
        Some(if width == 1 {
            chars.iter().map(|&byte| char::from(byte)).collect()
        } else {
            let units = chars
                .chunks_exact(2)
                .map(|bytes| u16::from_le_bytes([bytes[0], bytes[1]]))
                .collect::<Vec<_>>();
            String::from_utf16(&units)
                .map_err(|_| invalid(CF_RECORD_TYPE, "invalid UTF-16 conditional font name"))?
        })
    };
    Ok(XlsConditionalFont {
        raw: data.to_vec(),
        name,
    })
}

fn parse_style(data: &[u8]) -> XlsResult<(XlsConditionalStyle, usize)> {
    if data.len() < 6 {
        return Err(invalid(
            CF_RECORD_TYPE,
            "CF differential formatting header is truncated",
        ));
    }
    let options = read_u32(data, 0);
    let secondary = read_u16(data, 4);
    if options & 0x01c0_0000 != 0 || secondary & 0x7ff8 != 0 {
        return Err(invalid(
            CF_RECORD_TYPE,
            "CF differential formatting has nonzero reserved bits",
        ));
    }
    let mut offset = 6usize;
    let number_format = if options & 0x0200_0000 != 0 {
        if secondary & 1 != 0 {
            let length_bytes = take(data, &mut offset, 2, "number format")?;
            let length = usize::from(read_u16(length_bytes, 0));
            if length < 2 {
                return Err(invalid(
                    CF_RECORD_TYPE,
                    "custom differential number format is too short",
                ));
            }
            let rest = take(data, &mut offset, length - 2, "number format")?;
            Some(XlsConditionalNumberFormat::Custom(parse_simple_xl_unicode(
                rest,
                CF_RECORD_TYPE,
            )?))
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
            return Err(invalid(
                CF_RECORD_TYPE,
                "conditional relative indent is outside -15 through 255",
            ));
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
            return Err(invalid(
                CF_RECORD_TYPE,
                "conditional protection has nonzero reserved bits",
            ));
        }
        Some(XlsConditionalProtection {
            locked: bits & 1 != 0,
            hidden: bits & 2 != 0,
        })
    } else {
        None
    };
    Ok((
        XlsConditionalStyle {
            options,
            new_border: secondary & 4 != 0,
            number_format,
            font,
            alignment,
            border,
            pattern,
            protection,
        },
        offset,
    ))
}

fn parse_cf(data: &[u8], context: Option<&FormulaContext>) -> XlsResult<XlsConditionalRule> {
    if data.len() < 12 {
        return Err(invalid(
            CF_RECORD_TYPE,
            "CF payload is shorter than 12 bytes",
        ));
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
        (1, _) => {
            return Err(invalid(
                CF_RECORD_TYPE,
                "cell-value CF operator must be between 1 and 8",
            ));
        },
        (2, _) => return Err(invalid(CF_RECORD_TYPE, "formula CF operator must be zero")),
        _ => {
            return Err(invalid(
                CF_RECORD_TYPE,
                "legacy CF condition type must be 1 or 2",
            ));
        },
    };
    if matches!(kind, XlsConditionalRuleKind::Formula) && formula2_len != 0 {
        return Err(invalid(
            CF_RECORD_TYPE,
            "formula CF rule cannot contain a second formula",
        ));
    }
    if matches!(kind, XlsConditionalRuleKind::CellValue(operator) if !matches!(operator, XlsConditionalComparison::Between | XlsConditionalComparison::NotBetween))
        && formula2_len != 0
    {
        return Err(invalid(
            CF_RECORD_TYPE,
            "single-operand CF comparison cannot contain a second formula",
        ));
    }
    let (style, style_len) = parse_style(&data[6..])?;
    let formula_offset = 6 + style_len;
    if data.len() != formula_offset + formula1_len + formula2_len {
        return Err(invalid(
            CF_RECORD_TYPE,
            "CF formula lengths do not match the record payload",
        ));
    }
    let formula1_tokens = data[formula_offset..formula_offset + formula1_len].to_vec();
    let formula2_tokens = data[formula_offset + formula1_len..].to_vec();
    Ok(XlsConditionalRule {
        kind,
        style,
        formula1_rendered: render_formula(&formula1_tokens, context),
        formula2_rendered: render_formula(&formula2_tokens, context),
        formula1_tokens,
        formula2_tokens,
    })
}

fn parse_frt_header(
    data: &[u8],
    record_type: u16,
    referenced: bool,
) -> XlsResult<XlsConditionalFormatRange> {
    if data.len() < 12 {
        return Err(invalid(
            record_type,
            "future conditional-format record is shorter than its FRT header",
        ));
    }
    if read_u16(data, 0) != record_type {
        return Err(invalid(
            record_type,
            "FRT header record type does not match its containing record",
        ));
    }
    let flags = read_u16(data, 2);
    if flags != u16::from(referenced) {
        return Err(invalid(record_type, "FRT reference flags are invalid"));
    }
    let range = parse_range(&data[4..12], record_type)?;
    if !referenced && data[4..12].iter().any(|byte| *byte != 0) {
        return Err(invalid(
            record_type,
            "unreferenced FRT header range must be zero",
        ));
    }
    Ok(range)
}
fn dxf12_length(data: &[u8], offset: usize, record_type: u16) -> XlsResult<usize> {
    let cb = usize::try_from(
        *data
            .get(offset..offset + 4)
            .and_then(|bytes| bytes.try_into().ok())
            .map(u32::from_le_bytes)
            .as_ref()
            .ok_or_else(|| invalid(record_type, "truncated DXFN12"))?,
    )
    .map_err(|_| invalid(record_type, "DXFN12 length overflows"))?;
    if cb == 0 {
        if data.get(offset + 4..offset + 6) != Some(&[0, 0]) {
            return Err(invalid(
                record_type,
                "empty DXFN12 reserved field must be zero",
            ));
        }
        Ok(6)
    } else {
        let length = 4usize
            .checked_add(cb)
            .ok_or_else(|| invalid(record_type, "DXFN12 length overflows"))?;
        if data.get(offset..offset + length).is_none() {
            return Err(invalid(record_type, "truncated DXFN12 payload"));
        }
        Ok(length)
    }
}
fn comparison(value: u8, record_type: u16) -> XlsResult<XlsConditionalComparison> {
    Ok(match value {
        1 => XlsConditionalComparison::Between,
        2 => XlsConditionalComparison::NotBetween,
        3 => XlsConditionalComparison::Equal,
        4 => XlsConditionalComparison::NotEqual,
        5 => XlsConditionalComparison::GreaterThan,
        6 => XlsConditionalComparison::LessThan,
        7 => XlsConditionalComparison::GreaterThanOrEqual,
        8 => XlsConditionalComparison::LessThanOrEqual,
        _ => {
            return Err(invalid(
                record_type,
                "conditional comparison must be in 1..=8",
            ));
        },
    })
}
fn valid_template(value: u16) -> bool {
    matches!(value,0..=5|7..=12|15..=27|29|30)
}

fn parse_cf12(
    data: &[u8],
    context: Option<&FormulaContext>,
    priorities: &mut HashSet<u16>,
) -> XlsResult<XlsConditionalRule12> {
    parse_frt_header(data, CF12_RECORD_TYPE, false)?;
    if data.len() < 24 {
        return Err(invalid(CF12_RECORD_TYPE, "CF12 payload is truncated"));
    }
    let ct = data[12];
    let cp = data[13];
    let cce1 = usize::from(read_u16(data, 14));
    let cce2 = usize::from(read_u16(data, 16));
    if cce1 > 16409 || cce2 > 16409 {
        return Err(invalid(
            CF12_RECORD_TYPE,
            "CF12 formula exceeds 16409 bytes",
        ));
    }
    let kind = match ct {
        1 => XlsConditionalRule12Kind::CellValue(comparison(cp, CF12_RECORD_TYPE)?),
        2 if cp == 0 => XlsConditionalRule12Kind::Formula,
        3 if cp == 0 => XlsConditionalRule12Kind::ColorScale,
        4 if cp == 0 => XlsConditionalRule12Kind::DataBar,
        5 if cp == 0 => XlsConditionalRule12Kind::Filter,
        6 if cp == 0 => XlsConditionalRule12Kind::IconSet,
        2..=6 => {
            return Err(invalid(
                CF12_RECORD_TYPE,
                "non-comparison CF12 rule has a nonzero operator",
            ));
        },
        _ => {
            return Err(invalid(
                CF12_RECORD_TYPE,
                "CF12 condition type must be in 1..=6",
            ));
        },
    };
    if !matches!(
        kind,
        XlsConditionalRule12Kind::CellValue(
            XlsConditionalComparison::Between | XlsConditionalComparison::NotBetween
        )
    ) && cce2 != 0
    {
        return Err(invalid(
            CF12_RECORD_TYPE,
            "CF12 second formula is not allowed for this rule",
        ));
    }
    if matches!(
        kind,
        XlsConditionalRule12Kind::ColorScale
            | XlsConditionalRule12Kind::DataBar
            | XlsConditionalRule12Kind::Filter
            | XlsConditionalRule12Kind::IconSet
    ) && cce1 + cce2 != 0
    {
        return Err(invalid(
            CF12_RECORD_TYPE,
            "visual CF12 rule cannot contain comparison formulas",
        ));
    }
    let dxf_len = dxf12_length(data, 18, CF12_RECORD_TYPE)?;
    let differential_format = data[18..18 + dxf_len].to_vec();
    if matches!(
        kind,
        XlsConditionalRule12Kind::ColorScale
            | XlsConditionalRule12Kind::DataBar
            | XlsConditionalRule12Kind::IconSet
    ) && read_u32(data, 18) != 0
    {
        return Err(invalid(
            CF12_RECORD_TYPE,
            "visual CF12 DXFN12 must be empty",
        ));
    }
    let mut offset = 18 + dxf_len;
    let formula1_tokens = data
        .get(offset..offset + cce1)
        .ok_or_else(|| invalid(CF12_RECORD_TYPE, "truncated CF12 first formula"))?
        .to_vec();
    offset += cce1;
    let formula2_tokens = data
        .get(offset..offset + cce2)
        .ok_or_else(|| invalid(CF12_RECORD_TYPE, "truncated CF12 second formula"))?
        .to_vec();
    offset += cce2;
    let active_len = usize::from(
        *data
            .get(offset..offset + 2)
            .and_then(|bytes| bytes.try_into().ok())
            .map(u16::from_le_bytes)
            .as_ref()
            .ok_or_else(|| invalid(CF12_RECORD_TYPE, "truncated CF12 activity formula"))?,
    );
    offset += 2;
    let active_formula_tokens = data
        .get(offset..offset + active_len)
        .ok_or_else(|| invalid(CF12_RECORD_TYPE, "truncated CF12 activity formula"))?
        .to_vec();
    offset += active_len;
    if !matches!(
        kind,
        XlsConditionalRule12Kind::ColorScale
            | XlsConditionalRule12Kind::DataBar
            | XlsConditionalRule12Kind::IconSet
    ) && active_len != 0
    {
        return Err(invalid(
            CF12_RECORD_TYPE,
            "activity formula is only valid for visual CF12 rules",
        ));
    }
    let options = *data
        .get(offset)
        .ok_or_else(|| invalid(CF12_RECORD_TYPE, "truncated CF12 options"))?;
    offset += 1;
    if options & 0xec != 0 {
        return Err(invalid(
            CF12_RECORD_TYPE,
            "CF12 options contain reserved bits",
        ));
    }
    let stop_if_true = options & 2 != 0;
    if stop_if_true
        && matches!(
            kind,
            XlsConditionalRule12Kind::ColorScale
                | XlsConditionalRule12Kind::DataBar
                | XlsConditionalRule12Kind::IconSet
        )
    {
        return Err(invalid(
            CF12_RECORD_TYPE,
            "visual CF12 rule cannot stop-if-true",
        ));
    }
    let priority = read_u16(
        data.get(offset..offset + 2)
            .ok_or_else(|| invalid(CF12_RECORD_TYPE, "truncated CF12 priority"))?,
        0,
    );
    offset += 2;
    if !priorities.insert(priority) {
        return Err(invalid(
            CF12_RECORD_TYPE,
            "conditional-format priority is duplicated",
        ));
    }
    let template = read_u16(
        data.get(offset..offset + 2)
            .ok_or_else(|| invalid(CF12_RECORD_TYPE, "truncated CF12 template"))?,
        0,
    );
    offset += 2;
    if !valid_template(template) {
        return Err(invalid(CF12_RECORD_TYPE, "CF12 template is invalid"));
    }
    if data.get(offset) != Some(&16) {
        return Err(invalid(
            CF12_RECORD_TYPE,
            "CF12 template parameter length must be 16",
        ));
    }
    offset += 1;
    let template_parameters: [u8; 16] = data
        .get(offset..offset + 16)
        .ok_or_else(|| invalid(CF12_RECORD_TYPE, "truncated CF12 template parameters"))?
        .try_into()
        .unwrap();
    offset += 16;
    let rule_payload = data[offset..].to_vec();
    if matches!(
        kind,
        XlsConditionalRule12Kind::CellValue(_) | XlsConditionalRule12Kind::Formula
    ) && !rule_payload.is_empty()
    {
        return Err(invalid(
            CF12_RECORD_TYPE,
            "comparison/formula CF12 has unexpected rule payload",
        ));
    }
    Ok(XlsConditionalRule12 {
        kind,
        priority,
        stop_if_true,
        template,
        differential_format,
        formula1_rendered: render_formula(&formula1_tokens, context),
        formula2_rendered: render_formula(&formula2_tokens, context),
        active_formula_rendered: render_formula(&active_formula_tokens, context),
        formula1_tokens,
        formula2_tokens,
        active_formula_tokens,
        template_parameters,
        rule_payload,
    })
}

fn parse_condfmt12(data: &[u8]) -> XlsResult<PendingFormatting12> {
    let reference = parse_frt_header(data, CONDFMT12_RECORD_TYPE, true)?;
    let pending = parse_condfmt(
        data.get(12..)
            .ok_or_else(|| invalid(CONDFMT12_RECORD_TYPE, "truncated CondFmt12"))?,
    )?;
    if reference != pending.group.enclosing_range {
        return Err(invalid(
            CONDFMT12_RECORD_TYPE,
            "CondFmt12 FRT range does not match its enclosing range",
        ));
    }
    Ok(PendingFormatting12 {
        group: XlsConditionalFormatting12 {
            identifier: pending.group.identifier,
            tough_recalculation: pending.group.tough_recalculation,
            enclosing_range: pending.group.enclosing_range,
            ranges: pending.group.ranges,
            rules: Vec::with_capacity(pending.declared_rules),
        },
        declared_rules: pending.declared_rules,
    })
}
struct PendingFormatting12 {
    group: XlsConditionalFormatting12,
    declared_rules: usize,
}

enum ParsedExtension {
    Legacy {
        extension: Box<XlsConditionalExtension>,
        group_index: usize,
    },
    Future {
        identifier: u16,
        reference: XlsConditionalFormatRange,
    },
}
fn parse_cfex(
    data: &[u8],
    legacy: &[(u16, usize, XlsConditionalFormatRange)],
    priorities: &mut HashSet<u16>,
) -> XlsResult<ParsedExtension> {
    let reference = parse_frt_header(data, CFEX_RECORD_TYPE, true)?;
    if data.len() < 18 {
        return Err(invalid(CFEX_RECORD_TYPE, "CFEx payload is truncated"));
    }
    let future = read_u32(data, 12);
    if future > 1 {
        return Err(invalid(
            CFEX_RECORD_TYPE,
            "CFEx fIsCF12 must be zero or one",
        ));
    }
    let identifier = read_u16(data, 16);
    let group_index = legacy
        .iter()
        .find_map(|(candidate, index, enclosing)| {
            (*candidate == identifier && *enclosing == reference).then_some(*index)
        })
        .ok_or_else(|| {
            invalid(
                CFEX_RECORD_TYPE,
                "CFEx references an unknown legacy CondFmt identifier and range",
            )
        })?;
    if future == 1 {
        if data.len() != 18 {
            return Err(invalid(
                CFEX_RECORD_TYPE,
                "CFEx preceding CF12 must omit extension content",
            ));
        }
        return Ok(ParsedExtension::Future {
            identifier,
            reference,
        });
    }
    if data.len() < 43 {
        return Err(invalid(
            CFEX_RECORD_TYPE,
            "CFExNonCF12 payload is truncated",
        ));
    }
    let rule_index = read_u16(data, 18);
    let cp = data[20];
    if cp > 8 {
        return Err(invalid(CFEX_RECORD_TYPE, "CFEx comparison is invalid"));
    }
    let template = data[21];
    if !valid_template(u16::from(template)) {
        return Err(invalid(CFEX_RECORD_TYPE, "CFEx template is invalid"));
    }
    let priority = read_u16(data, 22);
    if !priorities.insert(priority) {
        return Err(invalid(
            CFEX_RECORD_TYPE,
            "conditional-format priority is duplicated",
        ));
    }
    let flags = data[24];
    if flags & 0xf4 != 0 {
        return Err(invalid(
            CFEX_RECORD_TYPE,
            "CFEx flags contain reserved bits",
        ));
    }
    let has_dxf = data[25];
    if has_dxf > 1 {
        return Err(invalid(
            CFEX_RECORD_TYPE,
            "CFEx fHasDXF must be zero or one",
        ));
    }
    let dxf_len = if has_dxf == 1 {
        dxf12_length(data, 26, CFEX_RECORD_TYPE)?
    } else {
        0
    };
    let mut offset = 26 + dxf_len;
    if data.get(offset) != Some(&16) {
        return Err(invalid(
            CFEX_RECORD_TYPE,
            "CFEx template parameter length must be 16",
        ));
    }
    offset += 1;
    let template_parameters = data
        .get(offset..offset + 16)
        .ok_or_else(|| invalid(CFEX_RECORD_TYPE, "truncated CFEx template parameters"))?
        .try_into()
        .unwrap();
    offset += 16;
    if offset != data.len() {
        return Err(invalid(CFEX_RECORD_TYPE, "CFEx has trailing bytes"));
    }
    Ok(ParsedExtension::Legacy {
        extension: Box::new(XlsConditionalExtension {
            identifier,
            legacy_rule_index: Some(rule_index),
            priority,
            active: flags & 1 != 0,
            stop_if_true: flags & 2 != 0,
            template,
            differential_format: if dxf_len == 0 {
                Vec::new()
            } else {
                data[26..26 + dxf_len].to_vec()
            },
            template_parameters,
            future_rule: None,
        }),
        group_index,
    })
}

/// Enforces the `CondFmt 1*3CF` collection grammar.
pub(crate) struct ConditionalFormatCollector {
    groups: Vec<XlsConditionalFormatting>,
    pending: Option<PendingFormatting>,
    future_groups: Vec<XlsConditionalFormatting12>,
    pending12: Option<PendingFormatting12>,
    extensions: Vec<XlsConditionalExtension>,
    pending_extension: Option<(u16, XlsConditionalFormatRange)>,
    identifiers: Vec<(u16, usize, XlsConditionalFormatRange)>,
    priorities: HashSet<u16>,
    extension_phase: bool,
}

impl ConditionalFormatCollector {
    pub(crate) fn new() -> Self {
        Self {
            groups: Vec::new(),
            pending: None,
            future_groups: Vec::new(),
            pending12: None,
            extensions: Vec::new(),
            pending_extension: None,
            identifiers: Vec::new(),
            priorities: HashSet::new(),
            extension_phase: false,
        }
    }

    pub(crate) fn feed_record(
        &mut self,
        record_type: u16,
        data: &[u8],
        context: Option<&FormulaContext>,
    ) -> XlsResult<()> {
        if self.pending.is_some() && record_type != CF_RECORD_TYPE {
            return Err(invalid(
                record_type,
                "CONDFMT must be followed immediately by its declared CF records",
            ));
        }
        if self.pending12.is_some() && record_type != CF12_RECORD_TYPE {
            return Err(invalid(
                record_type,
                "CondFmt12 must be followed immediately by its declared CF12 records",
            ));
        }
        if self.pending_extension.is_some() && record_type != CF12_RECORD_TYPE {
            return Err(invalid(
                record_type,
                "CFEx with fIsCF12 must be followed immediately by CF12",
            ));
        }
        match record_type {
            CONDFMT_RECORD_TYPE => {
                if self.extension_phase {
                    return Err(invalid(record_type, "CondFmt cannot follow CFEx records"));
                }
                self.pending = Some(parse_condfmt(data)?);
            },
            CF_RECORD_TYPE => {
                let pending = self
                    .pending
                    .as_mut()
                    .ok_or_else(|| invalid(record_type, "orphan CF record without CONDFMT"))?;
                pending.group.rules.push(parse_cf(data, context)?);
                if pending.group.rules.len() == pending.declared_rules {
                    let group = self.pending.take().unwrap().group;
                    if self.identifiers.iter().any(|(identifier, _, range)| {
                        *identifier == group.identifier && *range == group.enclosing_range
                    }) {
                        return Err(invalid(
                            record_type,
                            "conditional-format identifier and range are duplicated",
                        ));
                    }
                    self.identifiers.push((
                        group.identifier,
                        self.groups.len(),
                        group.enclosing_range,
                    ));
                    self.groups.push(group);
                }
            },
            CONDFMT12_RECORD_TYPE => {
                if self.extension_phase {
                    return Err(invalid(record_type, "CondFmt12 cannot follow CFEx records"));
                }
                self.pending12 = Some(parse_condfmt12(data)?);
            },
            CF12_RECORD_TYPE => {
                let rule = parse_cf12(data, context, &mut self.priorities)?;
                if let Some(pending) = self.pending12.as_mut() {
                    pending.group.rules.push(rule);
                    if pending.group.rules.len() == pending.declared_rules {
                        self.future_groups
                            .push(self.pending12.take().unwrap().group);
                    }
                } else if let Some((identifier, _)) = self.pending_extension.take() {
                    self.extensions.push(XlsConditionalExtension {
                        identifier,
                        legacy_rule_index: None,
                        priority: rule.priority,
                        active: true,
                        stop_if_true: rule.stop_if_true,
                        template: rule.template as u8,
                        differential_format: Vec::new(),
                        template_parameters: rule.template_parameters,
                        future_rule: Some(rule),
                    });
                } else {
                    return Err(invalid(record_type, "orphan CF12 record"));
                }
            },
            CFEX_RECORD_TYPE => {
                self.extension_phase = true;
                match parse_cfex(data, &self.identifiers, &mut self.priorities)? {
                    ParsedExtension::Legacy {
                        extension,
                        group_index,
                    } => {
                        if usize::from(extension.legacy_rule_index.unwrap())
                            >= self.groups[group_index].rules.len()
                        {
                            return Err(invalid(
                                record_type,
                                "CFEx legacy rule index is out of range",
                            ));
                        }
                        self.extensions.push(*extension)
                    },
                    ParsedExtension::Future {
                        identifier,
                        reference,
                    } => self.pending_extension = Some((identifier, reference)),
                }
            },
            _ => {},
        }
        Ok(())
    }

    pub(crate) fn finish(
        self,
    ) -> XlsResult<(
        Vec<XlsConditionalFormatting>,
        Vec<XlsConditionalFormatting12>,
        Vec<XlsConditionalExtension>,
    )> {
        if self.pending.is_some() || self.pending12.is_some() || self.pending_extension.is_some() {
            Err(invalid(
                CONDFMT_RECORD_TYPE,
                "worksheet ended before all declared CF rules were read",
            ))
        } else {
            Ok((self.groups, self.future_groups, self.extensions))
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_core::sheet::WorkbookTrait;

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
        collector
            .feed_record(CONDFMT_RECORD_TYPE, &header(2), None)
            .unwrap();
        collector
            .feed_record(
                CF_RECORD_TYPE,
                &rule(1, 1, &[0x1e, 1, 0], &[0x1e, 5, 0]),
                None,
            )
            .unwrap();
        collector
            .feed_record(CF_RECORD_TYPE, &rule(2, 0, &[0x1d], &[]), None)
            .unwrap();
        let groups = collector.finish().unwrap().0;
        assert_eq!(groups[0].rules().len(), 2);
        assert_eq!(
            groups[0].rules()[0].kind(),
            XlsConditionalRuleKind::CellValue(XlsConditionalComparison::Between)
        );
        assert_eq!(groups[0].rules()[1].formula1_tokens(), &[0x1d]);
    }

    #[test]
    fn rejects_malformed_ranges_formulas_and_sequences() {
        assert!(parse_condfmt(&header(0)).is_err());
        assert!(parse_cf(&rule(2, 1, &[0x1d], &[]), None).is_err());
        assert!(parse_cf(&rule(1, 5, &[0x1e, 1, 0], &[0x1e, 2, 0]), None).is_err());

        let mut collector = ConditionalFormatCollector::new();
        assert!(
            collector
                .feed_record(CF_RECORD_TYPE, &rule(2, 0, &[0x1d], &[]), None)
                .is_err()
        );
        let mut collector = ConditionalFormatCollector::new();
        collector
            .feed_record(CONDFMT_RECORD_TYPE, &header(1), None)
            .unwrap();
        assert!(collector.feed_record(0x000a, &[], None).is_err());
    }

    #[test]
    fn reads_poi_legacy_conditional_formatting_fixture() {
        use crate::XlsWorkbook;
        use std::fs::File;
        use std::path::Path;

        let fixture = Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/poi/test-data/spreadsheet/WithConditionalFormatting.xls");
        let workbook = XlsWorkbook::new(File::open(fixture).unwrap()).unwrap();
        let groups = workbook.xls_worksheet(0).unwrap().conditional_formattings();
        assert_eq!(groups.len(), 3);
        assert_eq!(groups[0].rules().len(), 2);
        assert_eq!(groups[0].ranges()[0].last_row(), 7);
        assert_eq!(
            groups[0].rules()[0].kind(),
            XlsConditionalRuleKind::CellValue(XlsConditionalComparison::GreaterThan)
        );
        assert!(groups[0].rules()[0].style().font().is_some());
        assert_eq!(groups[1].rules()[0].kind(), XlsConditionalRuleKind::Formula);
        assert!(
            groups
                .iter()
                .flat_map(|group| group.rules())
                .any(|rule| rule.formula1_rendered().is_some())
        );
        assert_eq!(
            groups[2].rules()[1].kind(),
            XlsConditionalRuleKind::CellValue(XlsConditionalComparison::Between)
        );
        assert!(!groups[2].rules()[1].formula2_tokens().is_empty());
    }

    #[test]
    fn reads_poi_future_conditional_formatting_fixture() {
        use crate::XlsWorkbook;
        use std::fs::File;
        use std::path::Path;
        let fixture = Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/poi/test-data/spreadsheet/NewStyleConditionalFormattings.xls");
        let workbook = XlsWorkbook::new(File::open(fixture).unwrap()).unwrap();
        let mut count = 0usize;
        let mut priorities = HashSet::new();
        for sheet in 0..workbook.worksheet_count() {
            let worksheet = workbook.xls_worksheet(sheet).unwrap();
            for group in worksheet.conditional_formattings12() {
                assert!(!group.ranges().is_empty());
                for rule in group.rules() {
                    assert!(priorities.insert(rule.priority()));
                    assert!(rule.differential_format().len() >= 6);
                    count += 1;
                }
            }
            for extension in worksheet.conditional_format_extensions() {
                assert!(priorities.insert(extension.priority()));
                count += 1;
            }
        }
        assert!(count > 0);
    }
}
