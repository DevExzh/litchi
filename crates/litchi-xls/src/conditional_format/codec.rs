//! BIFF8 conditional-format record codecs and worksheet collection grammar.

use super::model::{
    Alignment, Border, CF_RECORD_TYPE, CF12_RECORD_TYPE, CFEX_RECORD_TYPE, CONDFMT_RECORD_TYPE,
    CONDFMT12_RECORD_TYPE, Comparison, Extension, Font, Formatting, Formatting12, NumberFormat,
    Pattern, Protection, Range, Rule, Rule12, Rule12Kind, RuleKind, Style, read_u16, read_u32,
};
use crate::error::{Error, Result};
use crate::formula::{FormulaContext, render_formula};
use std::collections::HashSet;

fn invalid(record_type: u16, message: impl Into<String>) -> Error {
    Error::InvalidRecord {
        record_type,
        message: message.into(),
    }
}

fn parse_range(data: &[u8], record_type: u16) -> Result<Range> {
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
    Ok(Range {
        first_row,
        last_row,
        first_column: crate::utils::truncate_u16_to_u8(first_column),
        last_column: crate::utils::truncate_u16_to_u8(last_column),
    })
}

pub(crate) struct PendingFormatting {
    group: Formatting,
    declared_rules: usize,
}

pub(crate) fn parse_condfmt(data: &[u8]) -> Result<PendingFormatting> {
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
        group: Formatting {
            identifier: flags >> 1,
            tough_recalculation: flags & 1 != 0,
            enclosing_range,
            ranges,
            rules: Vec::with_capacity(declared_rules),
        },
        declared_rules,
    })
}

fn parse_simple_xl_unicode(data: &[u8], record_type: u16) -> Result<String> {
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
        String::from_utf16(&units)
            .map_err(|_error| invalid(record_type, "invalid UTF-16 number format"))
    }
}

fn take<'a>(data: &'a [u8], offset: &mut usize, length: usize, name: &str) -> Result<&'a [u8]> {
    let bytes = data.get(*offset..*offset + length).ok_or_else(|| {
        invalid(
            CF_RECORD_TYPE,
            format!("truncated {name} differential block"),
        )
    })?;
    *offset += length;
    Ok(bytes)
}

fn parse_font(data: &[u8]) -> Result<Font> {
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
                .map_err(|_error| invalid(CF_RECORD_TYPE, "invalid UTF-16 conditional font name"))?
        })
    };
    Ok(Font {
        raw: data.to_vec(),
        name,
    })
}

fn parse_style(data: &[u8]) -> Result<(Style, usize)> {
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
            Some(NumberFormat::Custom(parse_simple_xl_unicode(
                rest,
                CF_RECORD_TYPE,
            )?))
        } else {
            let bytes = take(data, &mut offset, 2, "number format")?;
            Some(NumberFormat::Identifier(bytes[1]))
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
        let relative_indent = crate::utils::wrap_u32_to_i32(read_u32(bytes, 4));
        if !(-15..=255).contains(&relative_indent) {
            return Err(invalid(
                CF_RECORD_TYPE,
                "conditional relative indent is outside -15 through 255",
            ));
        }
        Some(Alignment {
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
        Some(Border {
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
        Some(Pattern {
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
        Some(Protection {
            locked: bits & 1 != 0,
            hidden: bits & 2 != 0,
        })
    } else {
        None
    };
    Ok((
        Style {
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

pub(crate) fn parse_cf(data: &[u8], context: Option<&FormulaContext>) -> Result<Rule> {
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
        (1, operator @ 1..=8) => RuleKind::CellValue(match operator {
            1 => Comparison::Between,
            2 => Comparison::NotBetween,
            3 => Comparison::Equal,
            4 => Comparison::NotEqual,
            5 => Comparison::GreaterThan,
            6 => Comparison::LessThan,
            7 => Comparison::GreaterThanOrEqual,
            _ => Comparison::LessThanOrEqual,
        }),
        (2, 0) => RuleKind::Formula,
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
    if matches!(kind, RuleKind::Formula) && formula2_len != 0 {
        return Err(invalid(
            CF_RECORD_TYPE,
            "formula CF rule cannot contain a second formula",
        ));
    }
    if matches!(kind, RuleKind::CellValue(operator) if !matches!(operator, Comparison::Between | Comparison::NotBetween))
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
    Ok(Rule {
        kind,
        style,
        formula1_rendered: render_formula(&formula1_tokens, context),
        formula2_rendered: render_formula(&formula2_tokens, context),
        formula1_tokens,
        formula2_tokens,
    })
}

fn parse_frt_header(data: &[u8], record_type: u16, referenced: bool) -> Result<Range> {
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
fn dxf12_length(data: &[u8], offset: usize, record_type: u16) -> Result<usize> {
    let cb = usize::try_from(
        *data
            .get(offset..offset + 4)
            .and_then(|bytes| bytes.try_into().ok())
            .map(u32::from_le_bytes)
            .as_ref()
            .ok_or_else(|| invalid(record_type, "truncated DXFN12"))?,
    )
    .map_err(|_error| invalid(record_type, "DXFN12 length overflows"))?;
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
fn comparison(value: u8, record_type: u16) -> Result<Comparison> {
    Ok(match value {
        1 => Comparison::Between,
        2 => Comparison::NotBetween,
        3 => Comparison::Equal,
        4 => Comparison::NotEqual,
        5 => Comparison::GreaterThan,
        6 => Comparison::LessThan,
        7 => Comparison::GreaterThanOrEqual,
        8 => Comparison::LessThanOrEqual,
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
) -> Result<Rule12> {
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
        1 => Rule12Kind::CellValue(comparison(cp, CF12_RECORD_TYPE)?),
        2 if cp == 0 => Rule12Kind::Formula,
        3 if cp == 0 => Rule12Kind::ColorScale,
        4 if cp == 0 => Rule12Kind::DataBar,
        5 if cp == 0 => Rule12Kind::Filter,
        6 if cp == 0 => Rule12Kind::IconSet,
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
        Rule12Kind::CellValue(Comparison::Between | Comparison::NotBetween)
    ) && cce2 != 0
    {
        return Err(invalid(
            CF12_RECORD_TYPE,
            "CF12 second formula is not allowed for this rule",
        ));
    }
    if matches!(
        kind,
        Rule12Kind::ColorScale | Rule12Kind::DataBar | Rule12Kind::Filter | Rule12Kind::IconSet
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
        Rule12Kind::ColorScale | Rule12Kind::DataBar | Rule12Kind::IconSet
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
        Rule12Kind::ColorScale | Rule12Kind::DataBar | Rule12Kind::IconSet
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
            Rule12Kind::ColorScale | Rule12Kind::DataBar | Rule12Kind::IconSet
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
    if matches!(kind, Rule12Kind::CellValue(_) | Rule12Kind::Formula) && !rule_payload.is_empty() {
        return Err(invalid(
            CF12_RECORD_TYPE,
            "comparison/formula CF12 has unexpected rule payload",
        ));
    }
    Ok(Rule12 {
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

fn parse_condfmt12(data: &[u8]) -> Result<PendingFormatting12> {
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
        group: Formatting12 {
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
    group: Formatting12,
    declared_rules: usize,
}

enum ParsedExtension {
    Legacy {
        extension: Box<Extension>,
        group_index: usize,
    },
    Future {
        identifier: u16,
        reference: Range,
    },
}
fn parse_cfex(
    data: &[u8],
    legacy: &[(u16, usize, Range)],
    priorities: &mut HashSet<u16>,
) -> Result<ParsedExtension> {
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
        extension: Box::new(Extension {
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
pub(crate) struct Collector {
    groups: Vec<Formatting>,
    pending: Option<PendingFormatting>,
    future_groups: Vec<Formatting12>,
    pending12: Option<PendingFormatting12>,
    extensions: Vec<Extension>,
    pending_extension: Option<(u16, Range)>,
    identifiers: Vec<(u16, usize, Range)>,
    priorities: HashSet<u16>,
    extension_phase: bool,
}

impl Collector {
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
    ) -> Result<()> {
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
                    self.extensions.push(Extension {
                        identifier,
                        legacy_rule_index: None,
                        priority: rule.priority,
                        active: true,
                        stop_if_true: rule.stop_if_true,
                        template: crate::utils::truncate_u16_to_u8(rule.template),
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
                        self.extensions.push(*extension);
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

    pub(crate) fn finish(self) -> Result<(Vec<Formatting>, Vec<Formatting12>, Vec<Extension>)> {
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
