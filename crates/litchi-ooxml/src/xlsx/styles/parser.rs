//! Namespace-aware streaming parser for `xl/styles.xml`.

use std::collections::HashMap;

use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, QName, ResolveResult};
use quick_xml::reader::NsReader;

use super::border::{Color, Dir, Line, Rgb, Side, Tint};
use super::{Alignment, Border, CellStyle, Fill, Font, NumberFormat, Styles};
use crate::error::{OoxmlError, Result};
use litchi_ooxml_common::xml::unqualified_attribute_value;

const SPREADSHEETML_NAMESPACE: &[u8] = b"http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT_SPREADSHEETML_NAMESPACE: &[u8] = b"http://purl.oclc.org/ooxml/spreadsheetml/main";

type XmlReader<'a> = NsReader<&'a [u8]>;

fn is_spreadsheetml_name(spreadsheet_namespace: bool, name: QName<'_>, local_name: &[u8]) -> bool {
    spreadsheet_namespace && name.local_name().as_ref() == local_name
}

/// Parse styles from `xl/styles.xml` XML content.
pub fn parse_styles(content: &str) -> Result<Styles> {
    let differential_formats =
        crate::xlsx::conditional_formatting::parse_differential_formats(content.as_bytes())?;
    let processed = litchi_ooxml_common::mce::process_ooxml(content.as_bytes())?;
    let content =
        std::str::from_utf8(processed.as_ref()).map_err(|e| OoxmlError::Xml(e.to_string()))?;
    let mut reader = NsReader::from_reader(content.as_bytes());
    let mut styles = Styles::new();
    let mut seen_root = false;
    let mut closed_root = false;
    let mut depth = 0usize;
    let mut sections = 0u8;

    loop {
        let decoder = reader.decoder();
        let (namespace, event) = resolved_event(&mut reader, "styles")?;
        match event {
            Event::Start(element)
                if is_spreadsheetml_name(namespace, element.name(), b"styleSheet") =>
            {
                if seen_root || depth != 0 {
                    return Err(invalid("duplicate SpreadsheetML styleSheet element"));
                }
                seen_root = true;
                depth = 1;
            },
            Event::Start(element) if seen_root && !closed_root && depth == 1 => {
                match element.local_name().as_ref() {
                    b"numFmts" if is_spreadsheetml_name(namespace, element.name(), b"numFmts") => {
                        mark_section(&mut sections, 1, "numFmts")?;
                        let expected =
                            optional_u32(&element, b"count", decoder, "number-format count")?;
                        parse_number_formats(&mut reader, &mut styles.number_formats, expected)?;
                    },
                    b"fonts" if is_spreadsheetml_name(namespace, element.name(), b"fonts") => {
                        mark_section(&mut sections, 2, "fonts")?;
                        let expected = optional_u32(&element, b"count", decoder, "font count")?;
                        parse_fonts(&mut reader, &mut styles.fonts, expected)?;
                    },
                    b"fills" if is_spreadsheetml_name(namespace, element.name(), b"fills") => {
                        mark_section(&mut sections, 4, "fills")?;
                        let expected = optional_u32(&element, b"count", decoder, "fill count")?;
                        parse_fills(&mut reader, &mut styles.fills, expected)?;
                    },
                    b"borders" if is_spreadsheetml_name(namespace, element.name(), b"borders") => {
                        mark_section(&mut sections, 8, "borders")?;
                        let expected = optional_u32(&element, b"count", decoder, "border count")?;
                        parse_borders(&mut reader, &mut styles.borders, expected)?;
                    },
                    b"cellStyleXfs"
                        if is_spreadsheetml_name(namespace, element.name(), b"cellStyleXfs") =>
                    {
                        mark_section(&mut sections, 16, "cellStyleXfs")?;
                        let expected =
                            optional_u32(&element, b"count", decoder, "cell-style XF count")?;
                        parse_cell_xfs(
                            &mut reader,
                            &mut styles.cell_styles,
                            b"cellStyleXfs",
                            expected,
                        )?;
                    },
                    b"cellXfs" if is_spreadsheetml_name(namespace, element.name(), b"cellXfs") => {
                        mark_section(&mut sections, 32, "cellXfs")?;
                        let expected = optional_u32(&element, b"count", decoder, "cell XF count")?;
                        parse_cell_xfs(&mut reader, &mut styles.cell_xfs, b"cellXfs", expected)?;
                    },
                    _ => {
                        depth = depth
                            .checked_add(1)
                            .ok_or_else(|| invalid("SpreadsheetML style nesting is too deep"))?;
                    },
                }
            },
            Event::Start(_) if seen_root && !closed_root => {
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("SpreadsheetML style nesting is too deep"))?;
            },
            Event::Empty(element) if seen_root && !closed_root && depth == 1 => {
                match element.local_name().as_ref() {
                    b"numFmts" if is_spreadsheetml_name(namespace, element.name(), b"numFmts") => {
                        mark_section(&mut sections, 1, "numFmts")?;
                        validate_count(&element, decoder, 0, "number-format")?;
                    },
                    b"fonts" if is_spreadsheetml_name(namespace, element.name(), b"fonts") => {
                        mark_section(&mut sections, 2, "fonts")?;
                        validate_count(&element, decoder, 0, "font")?;
                    },
                    b"fills" if is_spreadsheetml_name(namespace, element.name(), b"fills") => {
                        mark_section(&mut sections, 4, "fills")?;
                        validate_count(&element, decoder, 0, "fill")?;
                    },
                    b"borders" if is_spreadsheetml_name(namespace, element.name(), b"borders") => {
                        mark_section(&mut sections, 8, "borders")?;
                        validate_count(&element, decoder, 0, "border")?;
                    },
                    b"cellStyleXfs"
                        if is_spreadsheetml_name(namespace, element.name(), b"cellStyleXfs") =>
                    {
                        mark_section(&mut sections, 16, "cellStyleXfs")?;
                        validate_count(&element, decoder, 0, "cell-style XF")?;
                    },
                    b"cellXfs" if is_spreadsheetml_name(namespace, element.name(), b"cellXfs") => {
                        mark_section(&mut sections, 32, "cellXfs")?;
                        validate_count(&element, decoder, 0, "cell XF")?;
                    },
                    _ => {},
                }
            },
            Event::End(element)
                if depth == 1
                    && is_spreadsheetml_name(namespace, element.name(), b"styleSheet") =>
            {
                depth = 0;
                closed_root = true;
            },
            Event::End(_) if seen_root && !closed_root => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("invalid SpreadsheetML style nesting"))?;
            },
            Event::Start(_) | Event::Empty(_) if !seen_root || closed_root => {
                return Err(invalid(
                    "styles XML must contain exactly one SpreadsheetML styleSheet root",
                ));
            },
            Event::Eof if !seen_root || !closed_root || depth != 0 => {
                return Err(invalid(
                    "styles XML has a missing or unterminated SpreadsheetML styleSheet root",
                ));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    styles.differential_formats = differential_formats;
    Ok(styles)
}

fn parse_number_formats(
    reader: &mut XmlReader<'_>,
    formats: &mut HashMap<u32, NumberFormat>,
    expected: Option<u32>,
) -> Result<()> {
    loop {
        let decoder = reader.decoder();
        let (namespace, event) = resolved_event(reader, "numFmts")?;
        match event {
            Event::Start(element) | Event::Empty(element)
                if is_spreadsheetml_name(namespace, element.name(), b"numFmt") =>
            {
                let id = required_u32(&element, b"numFmtId", decoder, "number-format ID")?;
                let code = required_string(&element, b"formatCode", decoder, "format code")?;
                if formats.insert(id, NumberFormat::new(id, code)).is_some() {
                    return Err(invalid(format!("duplicate number-format ID {id}")));
                }
            },
            Event::End(element) if is_spreadsheetml_name(namespace, element.name(), b"numFmts") => {
                check_count(expected, formats.len(), "number-format")?;
                return Ok(());
            },
            Event::Eof => return Err(unterminated("numFmts")),
            _ => {},
        }
    }
}

fn parse_fonts(
    reader: &mut XmlReader<'_>,
    fonts: &mut Vec<Font>,
    expected: Option<u32>,
) -> Result<()> {
    loop {
        let (namespace, event) = resolved_event(reader, "fonts")?;
        match event {
            Event::Start(element) if is_spreadsheetml_name(namespace, element.name(), b"font") => {
                fonts.push(parse_font(reader)?);
            },
            Event::Empty(element) if is_spreadsheetml_name(namespace, element.name(), b"font") => {
                fonts.push(Font::new());
            },
            Event::End(element) if is_spreadsheetml_name(namespace, element.name(), b"fonts") => {
                check_count(expected, fonts.len(), "font")?;
                return Ok(());
            },
            Event::Eof => return Err(unterminated("fonts")),
            _ => {},
        }
    }
}

fn parse_font(reader: &mut XmlReader<'_>) -> Result<Font> {
    let mut font = Font::new();
    let mut seen = 0u16;
    loop {
        let decoder = reader.decoder();
        let (namespace, event) = resolved_event(reader, "font")?;
        match event {
            Event::Start(element) | Event::Empty(element)
                if is_spreadsheetml_name(namespace, element.name(), b"name") =>
            {
                mark_property(&mut seen, 1, "font name")?;
                font.name = Some(required_string(&element, b"val", decoder, "font name")?);
            },
            Event::Start(element) | Event::Empty(element)
                if is_spreadsheetml_name(namespace, element.name(), b"sz") =>
            {
                mark_property(&mut seen, 2, "font size")?;
                let size = required_f64(&element, b"val", decoder, "font size")?;
                if !size.is_finite() || size <= 0.0 {
                    return Err(invalid(format!("invalid font size '{size}'")));
                }
                font.size = Some(size);
            },
            Event::Start(element) | Event::Empty(element)
                if is_spreadsheetml_name(namespace, element.name(), b"b") =>
            {
                mark_property(&mut seen, 4, "bold property")?;
                font.bold = boolean_property(&element, decoder, "bold")?;
            },
            Event::Start(element) | Event::Empty(element)
                if is_spreadsheetml_name(namespace, element.name(), b"i") =>
            {
                mark_property(&mut seen, 8, "italic property")?;
                font.italic = boolean_property(&element, decoder, "italic")?;
            },
            Event::Start(element) | Event::Empty(element)
                if is_spreadsheetml_name(namespace, element.name(), b"strike") =>
            {
                mark_property(&mut seen, 16, "strike property")?;
                font.strike = boolean_property(&element, decoder, "strike")?;
            },
            Event::Start(element) | Event::Empty(element)
                if is_spreadsheetml_name(namespace, element.name(), b"u") =>
            {
                mark_property(&mut seen, 32, "underline property")?;
                let value =
                    optional_string(&element, b"val", decoder)?.unwrap_or_else(|| "single".into());
                if !matches!(
                    value.as_str(),
                    "single" | "double" | "singleAccounting" | "doubleAccounting" | "none"
                ) {
                    return Err(invalid(format!("invalid underline style '{value}'")));
                }
                font.underline = (value != "none").then_some(value);
            },
            Event::Start(element) | Event::Empty(element)
                if is_spreadsheetml_name(namespace, element.name(), b"color") =>
            {
                mark_property(&mut seen, 64, "font color")?;
                font.color = parse_color(&element, decoder)?;
            },
            Event::Start(element) | Event::Empty(element)
                if is_spreadsheetml_name(namespace, element.name(), b"charset") =>
            {
                mark_property(&mut seen, 128, "font charset")?;
                font.charset = Some(required_u32(&element, b"val", decoder, "font charset")?);
            },
            Event::Start(element) | Event::Empty(element)
                if is_spreadsheetml_name(namespace, element.name(), b"family") =>
            {
                mark_property(&mut seen, 256, "font family")?;
                font.family = Some(required_u32(&element, b"val", decoder, "font family")?);
            },
            Event::Start(element) | Event::Empty(element)
                if is_spreadsheetml_name(namespace, element.name(), b"scheme") =>
            {
                mark_property(&mut seen, 512, "font scheme")?;
                let value = required_string(&element, b"val", decoder, "font scheme")?;
                if !matches!(value.as_str(), "major" | "minor" | "none") {
                    return Err(invalid(format!("invalid font scheme '{value}'")));
                }
                font.scheme = Some(value);
            },
            Event::End(element) if is_spreadsheetml_name(namespace, element.name(), b"font") => {
                return Ok(font);
            },
            Event::Eof => return Err(unterminated("font")),
            _ => {},
        }
    }
}

fn parse_fills(
    reader: &mut XmlReader<'_>,
    fills: &mut Vec<Fill>,
    expected: Option<u32>,
) -> Result<()> {
    loop {
        let (namespace, event) = resolved_event(reader, "fills")?;
        match event {
            Event::Start(element) if is_spreadsheetml_name(namespace, element.name(), b"fill") => {
                fills.push(parse_fill(reader)?);
            },
            Event::Empty(element) if is_spreadsheetml_name(namespace, element.name(), b"fill") => {
                fills.push(Fill::None);
            },
            Event::End(element) if is_spreadsheetml_name(namespace, element.name(), b"fills") => {
                check_count(expected, fills.len(), "fill")?;
                return Ok(());
            },
            Event::Eof => return Err(unterminated("fills")),
            _ => {},
        }
    }
}

fn parse_fill(reader: &mut XmlReader<'_>) -> Result<Fill> {
    let mut fill = None;
    loop {
        let decoder = reader.decoder();
        let (namespace, event) = resolved_event(reader, "fill")?;
        match event {
            Event::Start(element)
                if is_spreadsheetml_name(namespace, element.name(), b"patternFill") =>
            {
                set_once(
                    &mut fill,
                    parse_pattern_fill(reader, &element, decoder)?,
                    "fill definition",
                )?;
            },
            Event::Empty(element)
                if is_spreadsheetml_name(namespace, element.name(), b"patternFill") =>
            {
                set_once(
                    &mut fill,
                    empty_pattern_fill(&element, decoder)?,
                    "fill definition",
                )?;
            },
            Event::Start(element)
                if is_spreadsheetml_name(namespace, element.name(), b"gradientFill") =>
            {
                set_once(
                    &mut fill,
                    parse_gradient_fill(reader, &element, decoder)?,
                    "fill definition",
                )?;
            },
            Event::Empty(element)
                if is_spreadsheetml_name(namespace, element.name(), b"gradientFill") =>
            {
                set_once(
                    &mut fill,
                    empty_gradient_fill(&element, decoder)?,
                    "fill definition",
                )?;
            },
            Event::End(element) if is_spreadsheetml_name(namespace, element.name(), b"fill") => {
                return Ok(fill.unwrap_or(Fill::None));
            },
            Event::Eof => return Err(unterminated("fill")),
            _ => {},
        }
    }
}

fn empty_pattern_fill(element: &BytesStart<'_>, decoder: Decoder) -> Result<Fill> {
    let pattern =
        optional_string(element, b"patternType", decoder)?.unwrap_or_else(|| "none".into());
    validate_pattern(&pattern)?;
    if pattern == "none" {
        Ok(Fill::None)
    } else {
        Ok(Fill::Pattern {
            pattern_type: pattern,
            fg_color: None,
            bg_color: None,
        })
    }
}

fn parse_pattern_fill(
    reader: &mut XmlReader<'_>,
    element: &BytesStart<'_>,
    decoder: Decoder,
) -> Result<Fill> {
    let pattern =
        optional_string(element, b"patternType", decoder)?.unwrap_or_else(|| "none".into());
    validate_pattern(&pattern)?;
    let mut foreground = None;
    let mut background = None;
    loop {
        let decoder = reader.decoder();
        let (namespace, event) = resolved_event(reader, "patternFill")?;
        match event {
            Event::Start(element) | Event::Empty(element)
                if is_spreadsheetml_name(namespace, element.name(), b"fgColor") =>
            {
                let color = parse_color(&element, decoder)?;
                set_once(&mut foreground, color, "foreground color")?;
            },
            Event::Start(element) | Event::Empty(element)
                if is_spreadsheetml_name(namespace, element.name(), b"bgColor") =>
            {
                let color = parse_color(&element, decoder)?;
                set_once(&mut background, color, "background color")?;
            },
            Event::End(element)
                if is_spreadsheetml_name(namespace, element.name(), b"patternFill") =>
            {
                return if pattern == "none" {
                    Ok(Fill::None)
                } else {
                    Ok(Fill::Pattern {
                        pattern_type: pattern,
                        fg_color: foreground.flatten(),
                        bg_color: background.flatten(),
                    })
                };
            },
            Event::Eof => return Err(unterminated("patternFill")),
            _ => {},
        }
    }
}

fn empty_gradient_fill(element: &BytesStart<'_>, decoder: Decoder) -> Result<Fill> {
    Ok(Fill::Gradient {
        gradient_type: parse_gradient_type(element, decoder)?,
        stops: Vec::new(),
    })
}

fn parse_gradient_fill(
    reader: &mut XmlReader<'_>,
    element: &BytesStart<'_>,
    decoder: Decoder,
) -> Result<Fill> {
    let gradient_type = parse_gradient_type(element, decoder)?;
    let mut stops = Vec::new();
    loop {
        let decoder = reader.decoder();
        let (namespace, event) = resolved_event(reader, "gradientFill")?;
        match event {
            Event::Start(element) if is_spreadsheetml_name(namespace, element.name(), b"stop") => {
                stops.push(parse_gradient_stop(reader, &element, decoder)?);
            },
            Event::Empty(element) if is_spreadsheetml_name(namespace, element.name(), b"stop") => {
                return Err(invalid("gradient stop is missing its color"));
            },
            Event::End(element)
                if is_spreadsheetml_name(namespace, element.name(), b"gradientFill") =>
            {
                return Ok(Fill::Gradient {
                    gradient_type,
                    stops,
                });
            },
            Event::Eof => return Err(unterminated("gradientFill")),
            _ => {},
        }
    }
}

fn parse_gradient_stop(
    reader: &mut XmlReader<'_>,
    element: &BytesStart<'_>,
    decoder: Decoder,
) -> Result<(f64, String)> {
    let position = required_f64(element, b"position", decoder, "gradient-stop position")?;
    if !position.is_finite() || !(0.0..=1.0).contains(&position) {
        return Err(invalid(format!(
            "invalid gradient-stop position '{position}'"
        )));
    }
    let mut color = None;
    loop {
        let decoder = reader.decoder();
        let (namespace, event) = resolved_event(reader, "gradient stop")?;
        match event {
            Event::Start(element) | Event::Empty(element)
                if is_spreadsheetml_name(namespace, element.name(), b"color") =>
            {
                let parsed = parse_color(&element, decoder)?
                    .ok_or_else(|| invalid("gradient-stop color is empty"))?;
                set_once(&mut color, parsed, "gradient-stop color")?;
            },
            Event::End(element) if is_spreadsheetml_name(namespace, element.name(), b"stop") => {
                return Ok((
                    position,
                    color.ok_or_else(|| invalid("gradient stop is missing its color"))?,
                ));
            },
            Event::Eof => return Err(unterminated("gradient stop")),
            _ => {},
        }
    }
}

fn parse_gradient_type(element: &BytesStart<'_>, decoder: Decoder) -> Result<Option<String>> {
    let value = optional_string(element, b"type", decoder)?;
    if let Some(value) = &value
        && !matches!(value.as_str(), "linear" | "path")
    {
        return Err(invalid(format!("invalid gradient type '{value}'")));
    }
    Ok(value)
}

fn parse_borders(
    reader: &mut XmlReader<'_>,
    borders: &mut Vec<Border>,
    expected: Option<u32>,
) -> Result<()> {
    loop {
        let decoder = reader.decoder();
        let (namespace, event) = resolved_event(reader, "borders")?;
        match event {
            Event::Start(element)
                if is_spreadsheetml_name(namespace, element.name(), b"border") =>
            {
                borders.push(parse_border(reader, &element, decoder)?);
            },
            Event::Empty(element)
                if is_spreadsheetml_name(namespace, element.name(), b"border") =>
            {
                borders.push(parse_empty_border(&element, decoder)?);
            },
            Event::End(element) if is_spreadsheetml_name(namespace, element.name(), b"borders") => {
                check_count(expected, borders.len(), "border")?;
                return Ok(());
            },
            Event::Eof => return Err(unterminated("borders")),
            _ => {},
        }
    }
}

fn parse_empty_border(element: &BytesStart<'_>, decoder: Decoder) -> Result<Border> {
    let mut border = Border::new();
    set_border_attributes(&mut border, element, decoder)?;
    Ok(border)
}

fn parse_border(
    reader: &mut XmlReader<'_>,
    element: &BytesStart<'_>,
    decoder: Decoder,
) -> Result<Border> {
    let mut border = Border::new();
    set_border_attributes(&mut border, element, decoder)?;
    let mut seen = 0u16;
    loop {
        let decoder = reader.decoder();
        let (namespace, event) = resolved_event(reader, "border")?;
        match event {
            Event::Start(element) if namespace => {
                parse_border_child(reader, &element, decoder, false, &mut seen, &mut border)?;
            },
            Event::Empty(element) if namespace => {
                parse_border_child(reader, &element, decoder, true, &mut seen, &mut border)?;
            },
            Event::End(element) if is_spreadsheetml_name(namespace, element.name(), b"border") => {
                return Ok(border);
            },
            Event::Eof => return Err(unterminated("border")),
            _ => {},
        }
    }
}

fn parse_border_child(
    reader: &mut XmlReader<'_>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    empty: bool,
    seen: &mut u16,
    border: &mut Border,
) -> Result<()> {
    match element.local_name().as_ref() {
        b"start" => {
            mark_property(seen, 1, "start border")?;
            border.start = parse_border_side_event(reader, element, decoder, empty)?;
        },
        b"end" => {
            mark_property(seen, 2, "end border")?;
            border.end = parse_border_side_event(reader, element, decoder, empty)?;
        },
        b"left" => {
            mark_property(seen, 4, "left border")?;
            border.left = parse_border_side_event(reader, element, decoder, empty)?;
        },
        b"right" => {
            mark_property(seen, 8, "right border")?;
            border.right = parse_border_side_event(reader, element, decoder, empty)?;
        },
        b"top" => {
            mark_property(seen, 16, "top border")?;
            border.top = parse_border_side_event(reader, element, decoder, empty)?;
        },
        b"bottom" => {
            mark_property(seen, 32, "bottom border")?;
            border.bottom = parse_border_side_event(reader, element, decoder, empty)?;
        },
        b"diagonal" => {
            mark_property(seen, 64, "diagonal border")?;
            let side = parse_border_side_event(reader, element, decoder, empty)?;
            border.set_diagonal_side(side);
        },
        b"vertical" => {
            mark_property(seen, 128, "vertical border")?;
            border.vertical = parse_border_side_event(reader, element, decoder, empty)?;
        },
        b"horizontal" => {
            mark_property(seen, 256, "horizontal border")?;
            border.horizontal = parse_border_side_event(reader, element, decoder, empty)?;
        },
        _ => {},
    }
    Ok(())
}

fn parse_border_side_event(
    reader: &mut XmlReader<'_>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    empty: bool,
) -> Result<Option<Side>> {
    let style = optional_string(element, b"style", decoder)?.unwrap_or_else(|| "none".into());
    let line =
        Line::from_xml(&style).map_err(|_| invalid(format!("invalid border style '{style}'")))?;
    if empty {
        return Ok(line.map(Side::new));
    }
    let side = element.local_name().as_ref().to_vec();
    let mut color = None;
    loop {
        let decoder = reader.decoder();
        let (namespace, event) = resolved_event(reader, "border side")?;
        match event {
            Event::Start(element) | Event::Empty(element)
                if is_spreadsheetml_name(namespace, element.name(), b"color") =>
            {
                let parsed = parse_border_color(&element, decoder)?;
                set_once(&mut color, parsed, "border color")?;
            },
            Event::End(element) if is_spreadsheetml_name(namespace, element.name(), &side) => {
                return Ok(line.map(|line| {
                    let side = Side::new(line);
                    match color.flatten() {
                        Some(color) => side.with_color(color),
                        None => side,
                    }
                }));
            },
            Event::Eof => return Err(unterminated("border side")),
            _ => {},
        }
    }
}

fn set_border_attributes(
    border: &mut Border,
    element: &BytesStart<'_>,
    decoder: Decoder,
) -> Result<()> {
    let up = optional_bool(element, b"diagonalUp", decoder, "diagonalUp")?.unwrap_or(false);
    let down = optional_bool(element, b"diagonalDown", decoder, "diagonalDown")?.unwrap_or(false);
    border.set_diagonal_dir(Dir::from_flags(up, down));
    border.outline = optional_bool(element, b"outline", decoder, "border outline")?;
    Ok(())
}

fn parse_cell_xfs(
    reader: &mut XmlReader<'_>,
    styles: &mut Vec<CellStyle>,
    section: &[u8],
    expected: Option<u32>,
) -> Result<()> {
    loop {
        let decoder = reader.decoder();
        let (namespace, event) = resolved_event(reader, "cell XFs")?;
        match event {
            Event::Start(element) if is_spreadsheetml_name(namespace, element.name(), b"xf") => {
                styles.push(parse_xf(reader, &element, decoder)?);
            },
            Event::Empty(element) if is_spreadsheetml_name(namespace, element.name(), b"xf") => {
                styles.push(parse_xf_attributes(&element, decoder)?);
            },
            Event::End(element) if is_spreadsheetml_name(namespace, element.name(), section) => {
                check_count(expected, styles.len(), "cell XF")?;
                return Ok(());
            },
            Event::Eof => return Err(unterminated("cell XFs")),
            _ => {},
        }
    }
}

fn parse_xf(
    reader: &mut XmlReader<'_>,
    element: &BytesStart<'_>,
    decoder: Decoder,
) -> Result<CellStyle> {
    let mut style = parse_xf_attributes(element, decoder)?;
    loop {
        let decoder = reader.decoder();
        let (namespace, event) = resolved_event(reader, "xf")?;
        match event {
            Event::Start(element) | Event::Empty(element)
                if is_spreadsheetml_name(namespace, element.name(), b"alignment") =>
            {
                if style.alignment.is_some() {
                    return Err(invalid("duplicate cell alignment"));
                }
                style.alignment = Some(parse_alignment(&element, decoder)?);
            },
            Event::End(element) if is_spreadsheetml_name(namespace, element.name(), b"xf") => {
                return Ok(style);
            },
            Event::Eof => return Err(unterminated("xf")),
            _ => {},
        }
    }
}

fn parse_xf_attributes(element: &BytesStart<'_>, decoder: Decoder) -> Result<CellStyle> {
    Ok(CellStyle {
        num_fmt_id: optional_u32(element, b"numFmtId", decoder, "number-format ID")?,
        font_id: optional_u32(element, b"fontId", decoder, "font ID")?,
        fill_id: optional_u32(element, b"fillId", decoder, "fill ID")?,
        border_id: optional_u32(element, b"borderId", decoder, "border ID")?,
        xf_id: optional_u32(element, b"xfId", decoder, "XF ID")?,
        alignment: None,
        apply_number_format: optional_bool(
            element,
            b"applyNumberFormat",
            decoder,
            "applyNumberFormat",
        )?
        .unwrap_or(false),
        apply_font: optional_bool(element, b"applyFont", decoder, "applyFont")?.unwrap_or(false),
        apply_fill: optional_bool(element, b"applyFill", decoder, "applyFill")?.unwrap_or(false),
        apply_border: optional_bool(element, b"applyBorder", decoder, "applyBorder")?
            .unwrap_or(false),
        apply_alignment: optional_bool(element, b"applyAlignment", decoder, "applyAlignment")?
            .unwrap_or(false),
        quote_prefix: optional_bool(element, b"quotePrefix", decoder, "quotePrefix")?
            .unwrap_or(false),
    })
}

fn parse_alignment(element: &BytesStart<'_>, decoder: Decoder) -> Result<Alignment> {
    let horizontal = optional_string(element, b"horizontal", decoder)?;
    if let Some(value) = &horizontal
        && !matches!(
            value.as_str(),
            "general"
                | "left"
                | "center"
                | "right"
                | "fill"
                | "justify"
                | "centerContinuous"
                | "distributed"
        )
    {
        return Err(invalid(format!("invalid horizontal alignment '{value}'")));
    }
    let vertical = optional_string(element, b"vertical", decoder)?;
    if let Some(value) = &vertical
        && !matches!(
            value.as_str(),
            "top" | "center" | "bottom" | "justify" | "distributed"
        )
    {
        return Err(invalid(format!("invalid vertical alignment '{value}'")));
    }
    let text_rotation = optional_u32(element, b"textRotation", decoder, "text rotation")?;
    if text_rotation.is_some_and(|value| value > 180 && value != 255) {
        return Err(invalid("text rotation must be 0..=180 or 255"));
    }
    let reading_order = optional_u32(element, b"readingOrder", decoder, "reading order")?;
    if reading_order.is_some_and(|value| value > 2) {
        return Err(invalid("reading order must be 0, 1, or 2"));
    }
    Ok(Alignment {
        horizontal,
        vertical,
        text_rotation,
        wrap_text: optional_bool(element, b"wrapText", decoder, "wrapText")?.unwrap_or(false),
        indent: optional_u32(element, b"indent", decoder, "alignment indent")?,
        shrink_to_fit: optional_bool(element, b"shrinkToFit", decoder, "shrinkToFit")?
            .unwrap_or(false),
        reading_order,
    })
}

fn parse_color(element: &BytesStart<'_>, decoder: Decoder) -> Result<Option<String>> {
    let rgb = optional_string(element, b"rgb", decoder)?;
    let theme = optional_u32(element, b"theme", decoder, "theme color")?;
    let indexed = optional_u32(element, b"indexed", decoder, "indexed color")?;
    let auto = optional_bool(element, b"auto", decoder, "automatic color")?;
    let specified = usize::from(rgb.is_some())
        + usize::from(theme.is_some())
        + usize::from(indexed.is_some())
        + usize::from(auto.is_some());
    if specified > 1 {
        return Err(invalid("color has multiple mutually exclusive values"));
    }
    if let Some(tint) = optional_f64(element, b"tint", decoder, "color tint")?
        && (!tint.is_finite() || !(-1.0..=1.0).contains(&tint))
    {
        return Err(invalid(format!("invalid color tint '{tint}'")));
    }
    if let Some(rgb) = rgb {
        if !matches!(rgb.len(), 6 | 8) || !rgb.bytes().all(|byte| byte.is_ascii_hexdigit()) {
            return Err(invalid(format!("invalid RGB color '{rgb}'")));
        }
        Ok(Some(format!("#{rgb}")))
    } else if let Some(theme) = theme {
        Ok(Some(format!("theme:{theme}")))
    } else if let Some(indexed) = indexed {
        Ok(Some(format!("indexed:{indexed}")))
    } else if auto == Some(true) {
        Ok(Some("auto".to_string()))
    } else {
        Ok(None)
    }
}

fn parse_border_color(element: &BytesStart<'_>, decoder: Decoder) -> Result<Option<Color>> {
    let rgb = optional_string(element, b"rgb", decoder)?;
    let theme = optional_u32(element, b"theme", decoder, "theme color")?;
    let indexed = optional_u32(element, b"indexed", decoder, "indexed color")?;
    let automatic = optional_bool(element, b"auto", decoder, "automatic color")?;
    let specified = usize::from(rgb.is_some())
        + usize::from(theme.is_some())
        + usize::from(indexed.is_some())
        + usize::from(automatic.is_some());
    if specified > 1 {
        return Err(invalid("border color has multiple base values"));
    }
    let tint = optional_f64(element, b"tint", decoder, "border color tint")?
        .map(Tint::new)
        .transpose()
        .map_err(|_| invalid("border color tint must be finite and between -1 and 1"))?;

    let color = if let Some(value) = rgb {
        Color::from_rgb(
            value
                .parse::<Rgb>()
                .map_err(|_| invalid(format!("invalid RGB color '{value}'")))?,
        )
    } else if let Some(index) = theme {
        Color::theme(index)
    } else if let Some(index) = indexed {
        Color::indexed(index)
    } else if let Some(enabled) = automatic {
        Color::auto_value(enabled)
    } else {
        Color::default_base()
    };

    Ok(Some(match tint {
        Some(tint) => color.with_tint(tint),
        None => color,
    }))
}

fn validate_pattern(value: &str) -> Result<()> {
    if matches!(
        value,
        "none"
            | "solid"
            | "mediumGray"
            | "darkGray"
            | "lightGray"
            | "darkHorizontal"
            | "darkVertical"
            | "darkDown"
            | "darkUp"
            | "darkGrid"
            | "darkTrellis"
            | "lightHorizontal"
            | "lightVertical"
            | "lightDown"
            | "lightUp"
            | "lightGrid"
            | "lightTrellis"
            | "gray125"
            | "gray0625"
    ) {
        Ok(())
    } else {
        Err(invalid(format!("invalid fill pattern '{value}'")))
    }
}

fn boolean_property(element: &BytesStart<'_>, decoder: Decoder, name: &str) -> Result<bool> {
    optional_bool(element, b"val", decoder, name).map(|value| value.unwrap_or(true))
}

fn optional_bool(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> Result<Option<bool>> {
    optional_string(element, name, decoder)?
        .map(|value| match value.as_str() {
            "1" | "true" => Ok(true),
            "0" | "false" => Ok(false),
            _ => Err(invalid(format!("invalid {description} boolean '{value}'"))),
        })
        .transpose()
}

fn required_string(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> Result<String> {
    optional_string(element, name, decoder)?
        .ok_or_else(|| invalid(format!("missing {description} attribute")))
}

fn optional_string(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
) -> Result<Option<String>> {
    Ok(unqualified_attribute_value(element, name, decoder)?)
}

fn required_u32(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> Result<u32> {
    optional_u32(element, name, decoder, description)?
        .ok_or_else(|| invalid(format!("missing {description} attribute")))
}

fn optional_u32(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> Result<Option<u32>> {
    optional_string(element, name, decoder)?
        .map(|value| {
            value
                .parse::<u32>()
                .map_err(|_| invalid(format!("invalid {description} value '{value}'")))
        })
        .transpose()
}

fn required_f64(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> Result<f64> {
    optional_f64(element, name, decoder, description)?
        .ok_or_else(|| invalid(format!("missing {description} attribute")))
}

fn optional_f64(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> Result<Option<f64>> {
    optional_string(element, name, decoder)?
        .map(|value| {
            value
                .parse::<f64>()
                .map_err(|_| invalid(format!("invalid {description} value '{value}'")))
        })
        .transpose()
}

fn resolved_event(reader: &mut XmlReader<'_>, context: &str) -> Result<(bool, Event<'static>)> {
    let event = reader
        .read_event()
        .map_err(|error| xml_error(context, error))?
        .into_owned();
    let resolver = reader.resolver().clone();
    let (namespace, event) = resolver.resolve_event(event);
    let spreadsheet_namespace = matches!(
        namespace,
        ResolveResult::Bound(Namespace(value))
            if value == SPREADSHEETML_NAMESPACE || value == STRICT_SPREADSHEETML_NAMESPACE
    );
    Ok((spreadsheet_namespace, event))
}

fn validate_count(
    element: &BytesStart<'_>,
    decoder: Decoder,
    actual: usize,
    description: &str,
) -> Result<()> {
    check_count(
        optional_u32(element, b"count", decoder, &format!("{description} count"))?,
        actual,
        description,
    )
}

fn check_count(expected: Option<u32>, actual: usize, description: &str) -> Result<()> {
    if let Some(expected) = expected {
        let actual = u32::try_from(actual)
            .map_err(|_| invalid(format!("{description} count exceeds u32")))?;
        if expected != actual {
            return Err(invalid(format!(
                "{description} count declares {expected}, parsed {actual}"
            )));
        }
    }
    Ok(())
}

fn mark_section(seen: &mut u8, bit: u8, name: &str) -> Result<()> {
    mark_property(seen, bit, &format!("{name} section"))
}

fn mark_property<T>(seen: &mut T, bit: T, description: &str) -> Result<()>
where
    T: Copy + std::ops::BitAnd<Output = T> + std::ops::BitOrAssign + PartialEq + From<u8>,
{
    if (*seen & bit) != T::from(0) {
        return Err(invalid(format!("duplicate {description}")));
    }
    *seen |= bit;
    Ok(())
}

fn set_once<T>(target: &mut Option<T>, value: T, description: &str) -> Result<()> {
    if target.is_some() {
        return Err(invalid(format!("duplicate {description}")));
    }
    *target = Some(value);
    Ok(())
}

fn invalid(message: impl Into<String>) -> OoxmlError {
    OoxmlError::InvalidFormat(message.into())
}

fn unterminated(element: &str) -> OoxmlError {
    invalid(format!("unterminated SpreadsheetML {element} element"))
}

fn xml_error(context: &str, error: quick_xml::Error) -> OoxmlError {
    OoxmlError::Xml(format!("XML error in {context}: {error}"))
}

#[cfg(test)]
mod tests {
    use super::*;

    const S: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    #[test]
    fn parses_complete_namespaced_style_sheet() {
        let xml = format!(
            r#"<s:styleSheet xmlns:s="{S}" xmlns:f="urn:foreign">
                <s:numFmts count="1"><s:numFmt numFmtId="164" formatCode="0.00&amp; units"/></s:numFmts>
                <s:fonts count="1"><s:font><s:name val="A &amp; B"/><s:sz val="11.5"/>
                    <s:b val="0"/><s:i/><s:u val="double"/><s:color rgb="FF112233"/>
                    <s:charset val="1"/><s:family val="2"/><s:scheme val="minor"/></s:font></s:fonts>
                <s:fills count="2"><s:fill><s:patternFill patternType="solid"><s:fgColor theme="2"/>
                    <s:bgColor indexed="64"/></s:patternFill></s:fill>
                    <s:fill><s:gradientFill type="linear"><s:stop position="0"><s:color rgb="FF000000"/></s:stop>
                        <s:stop position="1"><s:color rgb="FFFFFFFF"/></s:stop></s:gradientFill></s:fill></s:fills>
                <s:borders count="1"><s:border diagonalUp="1"><s:left style="thin"><s:color auto="1"/></s:left>
                    <s:right/><s:top style="dashed"/><s:bottom style="double"/><s:diagonal style="hair"/></s:border></s:borders>
                <s:cellStyleXfs count="1"><s:xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></s:cellStyleXfs>
                <s:cellXfs count="2"><s:xf numFmtId="164" fontId="0" fillId="1" borderId="0"
                    xfId="0" applyNumberFormat="true" applyFont="0" applyFill="1" applyBorder="1"
                    applyAlignment="1" quotePrefix="false"><s:alignment horizontal="center" vertical="top"
                    textRotation="45" wrapText="1" indent="2" shrinkToFit="0" readingOrder="2"/></s:xf><s:xf/></s:cellXfs>
                <f:fonts><f:font/></f:fonts>
            </s:styleSheet>"#
        );
        let styles = parse_styles(&xml).unwrap();
        assert_eq!(styles.number_formats[&164].code, "0.00& units");
        assert_eq!(styles.fonts.len(), 1);
        assert_eq!(styles.fonts[0].name.as_deref(), Some("A & B"));
        assert!(!styles.fonts[0].bold);
        assert!(styles.fonts[0].italic);
        assert_eq!(styles.fonts[0].underline.as_deref(), Some("double"));
        assert_eq!(styles.fills.len(), 2);
        match &styles.fills[1] {
            Fill::Gradient {
                gradient_type,
                stops,
            } => {
                assert_eq!(gradient_type.as_deref(), Some("linear"));
                assert_eq!(
                    stops,
                    &[(0.0, "#FF000000".into()), (1.0, "#FFFFFFFF".into())]
                );
            },
            _ => panic!("expected gradient fill"),
        }
        assert_eq!(
            styles.borders[0]
                .diagonal
                .as_ref()
                .and_then(super::super::border::Diagonal::dir),
            Some(Dir::Up)
        );
        assert_eq!(styles.borders[0].left.as_ref().unwrap().line, Line::Thin);
        assert_eq!(
            styles.borders[0].left.as_ref().unwrap().color,
            Some(Color::auto())
        );
        assert_eq!(styles.cell_styles.len(), 1);
        assert_eq!(styles.cell_xfs.len(), 2);
        assert!(styles.cell_xfs[0].apply_number_format);
        assert!(!styles.cell_xfs[0].apply_font);
        assert_eq!(
            styles.cell_xfs[0].alignment.as_ref().unwrap().reading_order,
            Some(2)
        );
    }

    #[test]
    fn supports_strict_aliases_and_empty_sections() {
        let xml = r#"<x:styleSheet xmlns:x="http://purl.oclc.org/ooxml/spreadsheetml/main">
            <x:numFmts count="0"/><x:fonts count="1"><x:font/></x:fonts><x:fills count="0"/>
            <x:borders count="1"><x:border outline="0" diagonalDown="1">
                <x:start style="thin"><x:color theme="2" tint="-0.25"/></x:start>
                <x:end style="dotted"/><x:bottom style="medium"><x:color tint="0.5"/></x:bottom>
                <x:diagonal style="dashed"/><x:vertical style="hair"/><x:horizontal style="double"/>
                </x:border></x:borders><x:cellStyleXfs count="0"/>
            <x:cellXfs count="1"><x:xf quotePrefix="1"/></x:cellXfs></x:styleSheet>"#;
        let styles = parse_styles(xml).unwrap();
        assert_eq!(styles.fonts.len(), 1);
        assert_eq!(styles.borders.len(), 1);
        let border = &styles.borders[0];
        assert_eq!(border.outline, Some(false));
        assert_eq!(
            border.start.as_ref().map(|side| side.line),
            Some(Line::Thin)
        );
        assert_eq!(
            border.end.as_ref().map(|side| side.line),
            Some(Line::Dotted)
        );
        assert_eq!(
            border.vertical.as_ref().map(|side| side.line),
            Some(Line::Hair)
        );
        assert_eq!(
            border.horizontal.as_ref().map(|side| side.line),
            Some(Line::Double)
        );
        assert_eq!(
            border
                .diagonal
                .as_ref()
                .and_then(super::super::border::Diagonal::dir),
            Some(Dir::Down)
        );
        assert!(matches!(
            border.start.as_ref().and_then(|side| side.color),
            Some(Color::Theme {
                index: 2,
                tint: Some(value),
            }) if value.get() == -0.25
        ));
        assert!(matches!(
            border.bottom.as_ref().and_then(|side| side.color),
            Some(Color::Default {
                tint: Some(value),
            }) if value.get() == 0.5
        ));
        assert!(styles.cell_xfs[0].quote_prefix);
    }

    #[test]
    fn rejects_malformed_styles() {
        let invalid = [
            "<styleSheet/>",
            &format!(r#"<styleSheet xmlns="{S}"><fonts count="2"><font/></fonts></styleSheet>"#),
            &format!(
                r#"<styleSheet xmlns="{S}"><numFmts><numFmt numFmtId="x" formatCode="0"/></numFmts></styleSheet>"#
            ),
            &format!(
                r#"<styleSheet xmlns="{S}"><fonts><font><b val="maybe"/></font></fonts></styleSheet>"#
            ),
            &format!(
                r#"<styleSheet xmlns="{S}"><fills><fill><gradientFill><stop position="2"><color rgb="FF000000"/></stop></gradientFill></fill></fills></styleSheet>"#
            ),
            &format!(
                r#"<styleSheet xmlns="{S}"><borders><border><left style="invalid"/></border></borders></styleSheet>"#
            ),
            &format!(
                r#"<styleSheet xmlns="{S}"><borders><border><left style="thin"><color rgb="112233"/></left></border></borders></styleSheet>"#
            ),
            &format!(
                r#"<styleSheet xmlns="{S}"><borders><border><left style="thin"><color theme="1" indexed="2"/></left></border></borders></styleSheet>"#
            ),
            &format!(
                r#"<styleSheet xmlns="{S}"><borders><border><left style="thin"><color rgb="FF112233" tint="2"/></left></border></borders></styleSheet>"#
            ),
            &format!(
                r#"<styleSheet xmlns="{S}"><cellXfs><xf><alignment textRotation="200"/></xf></cellXfs></styleSheet>"#
            ),
            &format!(r#"<styleSheet xmlns="{S}"><fonts><font>"#),
        ];
        for xml in invalid {
            assert!(
                parse_styles(xml).is_err(),
                "accepted invalid styles XML: {xml}"
            );
        }
    }
}
