//! XML parser for styles.xml file.
//!
//! This module contains the core parsing logic for Excel styles.xml files.
//! It uses quick-xml for efficient streaming XML parsing.

use std::collections::HashMap;

use quick_xml::Reader;
use quick_xml::events::Event;

use super::{Alignment, Border, BorderStyle, CellStyle, Fill, Font, NumberFormat, Styles};
use crate::ooxml::error::{OoxmlError, Result};

/// Parse styles from xl/styles.xml XML content.
///
/// This is the main entry point for parsing Excel styles. It processes
/// the XML and extracts all formatting information.
pub fn parse_styles(content: &str) -> Result<Styles> {
    let mut reader = Reader::from_str(content);
    reader.config_mut().trim_text(true);

    let mut styles = Styles::new();
    let mut buf = Vec::with_capacity(1024);

    loop {
        buf.clear();
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => match e.local_name().as_ref() {
                b"numFmts" => {
                    parse_number_formats(&mut reader, &mut styles.number_formats)?;
                },
                b"fonts" => {
                    parse_fonts(&mut reader, &mut styles.fonts)?;
                },
                b"fills" => {
                    parse_fills(&mut reader, &mut styles.fills)?;
                },
                b"borders" => {
                    parse_borders(&mut reader, &mut styles.borders)?;
                },
                b"cellStyleXfs" => {
                    parse_cell_xfs(&mut reader, &mut styles.cell_styles)?;
                },
                b"cellXfs" => {
                    parse_cell_xfs(&mut reader, &mut styles.cell_xfs)?;
                },
                _ => {},
            },
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(OoxmlError::Xml(format!("XML parsing error: {}", e)));
            },
            _ => {},
        }
    }

    Ok(styles)
}

/// Parse number formats section.
fn parse_number_formats(
    reader: &mut Reader<&[u8]>,
    number_formats: &mut HashMap<u32, NumberFormat>,
) -> Result<()> {
    let mut buf = Vec::with_capacity(512);

    loop {
        buf.clear();
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) if e.local_name().as_ref() == b"numFmt" => {
                let mut id = None;
                let mut code = None;

                for attr in e.attributes().flatten() {
                    match attr.key.local_name().as_ref() {
                        b"numFmtId" => {
                            if let Ok(value) = attr.decode_and_unescape_value(reader.decoder()) {
                                id = value.parse::<u32>().ok();
                            }
                        },
                        b"formatCode" => {
                            if let Ok(value) = attr.decode_and_unescape_value(reader.decoder()) {
                                code = Some(value.to_string());
                            }
                        },
                        _ => {},
                    }
                }

                if let (Some(id), Some(code)) = (id, code) {
                    number_formats.insert(id, NumberFormat::new(id, code));
                }
            },
            Ok(Event::End(e)) if e.local_name().as_ref() == b"numFmts" => break,
            Ok(Event::Eof) => break,
            Err(e) => return Err(OoxmlError::Xml(format!("XML error in numFmts: {}", e))),
            _ => {},
        }
    }

    Ok(())
}

/// Parse fonts section.
fn parse_fonts(reader: &mut Reader<&[u8]>, fonts: &mut Vec<Font>) -> Result<()> {
    let mut buf = Vec::with_capacity(512);

    loop {
        buf.clear();
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) if e.local_name().as_ref() == b"font" => {
                let font = parse_font(reader)?;
                fonts.push(font);
            },
            Ok(Event::End(e)) if e.local_name().as_ref() == b"fonts" => break,
            Ok(Event::Eof) => break,
            Err(e) => return Err(OoxmlError::Xml(format!("XML error in fonts: {}", e))),
            _ => {},
        }
    }

    Ok(())
}

/// Parse a single font element.
fn parse_font(reader: &mut Reader<&[u8]>) -> Result<Font> {
    let mut font = Font::new();
    let mut buf = Vec::with_capacity(256);

    loop {
        buf.clear();
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                match e.local_name().as_ref() {
                    b"name" => {
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"val"
                                && let Ok(value) = attr.decode_and_unescape_value(reader.decoder())
                            {
                                font.name = Some(value.to_string());
                            }
                        }
                    },
                    b"sz" => {
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"val"
                                && let Ok(value) = attr.decode_and_unescape_value(reader.decoder())
                            {
                                font.size = value.parse::<f64>().ok();
                            }
                        }
                    },
                    b"b" => font.bold = true,
                    b"i" => font.italic = true,
                    b"strike" => font.strike = true,
                    b"u" => {
                        // Underline can have a val attribute, default is "single"
                        let mut underline_type = "single".to_string();
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"val"
                                && let Ok(value) = attr.decode_and_unescape_value(reader.decoder())
                            {
                                underline_type = value.to_string();
                            }
                        }
                        font.underline = Some(underline_type);
                    },
                    b"color" => {
                        font.color = parse_color(reader, &e)?;
                    },
                    b"charset" => {
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"val"
                                && let Ok(value) = attr.decode_and_unescape_value(reader.decoder())
                            {
                                font.charset = value.parse::<u32>().ok();
                            }
                        }
                    },
                    b"family" => {
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"val"
                                && let Ok(value) = attr.decode_and_unescape_value(reader.decoder())
                            {
                                font.family = value.parse::<u32>().ok();
                            }
                        }
                    },
                    b"scheme" => {
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"val"
                                && let Ok(value) = attr.decode_and_unescape_value(reader.decoder())
                            {
                                font.scheme = Some(value.to_string());
                            }
                        }
                    },
                    _ => {},
                }
            },
            Ok(Event::End(e)) if e.local_name().as_ref() == b"font" => break,
            Ok(Event::Eof) => break,
            Err(e) => return Err(OoxmlError::Xml(format!("XML error in font: {}", e))),
            _ => {},
        }
    }

    Ok(font)
}

/// Parse fills section.
fn parse_fills(reader: &mut Reader<&[u8]>, fills: &mut Vec<Fill>) -> Result<()> {
    let mut buf = Vec::with_capacity(512);

    loop {
        buf.clear();
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) if e.local_name().as_ref() == b"fill" => {
                let fill = parse_fill(reader)?;
                fills.push(fill);
            },
            Ok(Event::End(e)) if e.local_name().as_ref() == b"fills" => break,
            Ok(Event::Eof) => break,
            Err(e) => return Err(OoxmlError::Xml(format!("XML error in fills: {}", e))),
            _ => {},
        }
    }

    Ok(())
}

/// Parse a single fill element.
fn parse_fill(reader: &mut Reader<&[u8]>) -> Result<Fill> {
    let mut fill = Fill::None;
    let mut buf = Vec::with_capacity(256);

    loop {
        buf.clear();
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) if e.local_name().as_ref() == b"patternFill" => {
                fill = parse_pattern_fill(reader, &e)?;
            },
            Ok(Event::Start(e)) if e.local_name().as_ref() == b"gradientFill" => {
                fill = parse_gradient_fill(reader, &e)?;
            },
            Ok(Event::End(e)) if e.local_name().as_ref() == b"fill" => break,
            Ok(Event::Eof) => break,
            Err(e) => return Err(OoxmlError::Xml(format!("XML error in fill: {}", e))),
            _ => {},
        }
    }

    Ok(fill)
}

/// Parse a pattern fill.
fn parse_pattern_fill(
    reader: &mut Reader<&[u8]>,
    start: &quick_xml::events::BytesStart,
) -> Result<Fill> {
    let mut pattern_type = String::from("none");
    let mut fg_color = None;
    let mut bg_color = None;

    // Get pattern type from attributes
    for attr in start.attributes().flatten() {
        if attr.key.local_name().as_ref() == b"patternType"
            && let Ok(value) = attr.decode_and_unescape_value(reader.decoder())
        {
            pattern_type = value.to_string();
        }
    }

    let mut buf = Vec::with_capacity(128);
    loop {
        buf.clear();
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => match e.local_name().as_ref() {
                b"fgColor" => {
                    fg_color = parse_color(reader, &e)?;
                },
                b"bgColor" => {
                    bg_color = parse_color(reader, &e)?;
                },
                _ => {},
            },
            Ok(Event::End(e)) if e.local_name().as_ref() == b"patternFill" => break,
            Ok(Event::Eof) => break,
            Err(e) => return Err(OoxmlError::Xml(format!("XML error in patternFill: {}", e))),
            _ => {},
        }
    }

    if pattern_type == "none" {
        Ok(Fill::None)
    } else {
        Ok(Fill::Pattern {
            pattern_type,
            fg_color,
            bg_color,
        })
    }
}

/// Parse a gradient fill.
fn parse_gradient_fill(
    reader: &mut Reader<&[u8]>,
    start: &quick_xml::events::BytesStart,
) -> Result<Fill> {
    let mut gradient_type = None;
    let stops = Vec::new();

    // Get gradient type from attributes
    for attr in start.attributes().flatten() {
        if attr.key.local_name().as_ref() == b"type"
            && let Ok(value) = attr.decode_and_unescape_value(reader.decoder())
        {
            gradient_type = Some(value.to_string());
        }
    }

    // Skip to end of gradientFill (full gradient parsing can be added later)
    let mut buf = Vec::with_capacity(128);
    let mut depth = 1;
    loop {
        buf.clear();
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) if e.local_name().as_ref() == b"gradientFill" => depth += 1,
            Ok(Event::End(e)) if e.local_name().as_ref() == b"gradientFill" => {
                depth -= 1;
                if depth == 0 {
                    break;
                }
            },
            Ok(Event::Eof) => break,
            Err(e) => return Err(OoxmlError::Xml(format!("XML error in gradientFill: {}", e))),
            _ => {},
        }
    }

    Ok(Fill::Gradient {
        gradient_type,
        stops,
    })
}

/// Parse borders section.
fn parse_borders(reader: &mut Reader<&[u8]>, borders: &mut Vec<Border>) -> Result<()> {
    let mut buf = Vec::with_capacity(512);

    loop {
        buf.clear();
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) if e.local_name().as_ref() == b"border" => {
                let border = parse_border(reader, &e)?;
                borders.push(border);
            },
            Ok(Event::End(e)) if e.local_name().as_ref() == b"borders" => break,
            Ok(Event::Eof) => break,
            Err(e) => return Err(OoxmlError::Xml(format!("XML error in borders: {}", e))),
            _ => {},
        }
    }

    Ok(())
}

/// Parse a single border element.
fn parse_border(
    reader: &mut Reader<&[u8]>,
    start: &quick_xml::events::BytesStart,
) -> Result<Border> {
    let mut border = Border::new();

    // Check for diagonal attributes
    for attr in start.attributes().flatten() {
        match attr.key.local_name().as_ref() {
            b"diagonalUp" => {
                if let Ok(value) = attr.decode_and_unescape_value(reader.decoder())
                    && (value == "1" || value == "true")
                {
                    border.diagonal_direction = Some(border.diagonal_direction.unwrap_or(0) | 1);
                }
            },
            b"diagonalDown" => {
                if let Ok(value) = attr.decode_and_unescape_value(reader.decoder())
                    && (value == "1" || value == "true")
                {
                    border.diagonal_direction = Some(border.diagonal_direction.unwrap_or(0) | 2);
                }
            },
            _ => {},
        }
    }

    let mut buf = Vec::with_capacity(256);
    loop {
        buf.clear();
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) if e.local_name().as_ref() == b"left" => {
                border.left = parse_border_side(reader, &e)?;
            },
            Ok(Event::Start(e)) if e.local_name().as_ref() == b"right" => {
                border.right = parse_border_side(reader, &e)?;
            },
            Ok(Event::Start(e)) if e.local_name().as_ref() == b"top" => {
                border.top = parse_border_side(reader, &e)?;
            },
            Ok(Event::Start(e)) if e.local_name().as_ref() == b"bottom" => {
                border.bottom = parse_border_side(reader, &e)?;
            },
            Ok(Event::Start(e)) if e.local_name().as_ref() == b"diagonal" => {
                border.diagonal = parse_border_side(reader, &e)?;
            },
            Ok(Event::End(e)) if e.local_name().as_ref() == b"border" => break,
            Ok(Event::Eof) => break,
            Err(e) => return Err(OoxmlError::Xml(format!("XML error in border: {}", e))),
            _ => {},
        }
    }

    Ok(border)
}

/// Parse a single border side (left, right, top, bottom, diagonal).
fn parse_border_side(
    reader: &mut Reader<&[u8]>,
    start: &quick_xml::events::BytesStart,
) -> Result<Option<BorderStyle>> {
    let mut style = String::from("none");
    let mut color = None;

    // Get style from attributes
    for attr in start.attributes().flatten() {
        if attr.key.local_name().as_ref() == b"style"
            && let Ok(value) = attr.decode_and_unescape_value(reader.decoder())
        {
            style = value.to_string();
        }
    }

    // Parse color
    let mut buf = Vec::with_capacity(128);
    let side_name = start.local_name();
    loop {
        buf.clear();
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) if e.local_name().as_ref() == b"color" => {
                color = parse_color(reader, &e)?;
            },
            Ok(Event::End(e)) if e.local_name() == side_name => break,
            Ok(Event::Eof) => break,
            Err(e) => return Err(OoxmlError::Xml(format!("XML error in border side: {}", e))),
            _ => {},
        }
    }

    if style == "none" {
        Ok(None)
    } else {
        Ok(Some(BorderStyle::new(style, color)))
    }
}

/// Parse cell XFs (cell format records).
fn parse_cell_xfs(reader: &mut Reader<&[u8]>, cell_xfs: &mut Vec<CellStyle>) -> Result<()> {
    let mut buf = Vec::with_capacity(512);

    loop {
        buf.clear();
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) if e.local_name().as_ref() == b"xf" => {
                let style = parse_xf(reader, &e)?;
                cell_xfs.push(style);
            },
            Ok(Event::End(e))
                if e.local_name().as_ref() == b"cellXfs"
                    || e.local_name().as_ref() == b"cellStyleXfs" =>
            {
                break;
            },
            Ok(Event::Eof) => break,
            Err(e) => return Err(OoxmlError::Xml(format!("XML error in cellXfs: {}", e))),
            _ => {},
        }
    }

    Ok(())
}

/// Parse a single xf (format) element.
fn parse_xf(
    reader: &mut Reader<&[u8]>,
    start: &quick_xml::events::BytesStart,
) -> Result<CellStyle> {
    let mut style = CellStyle::new();

    // Parse attributes
    for attr in start.attributes().flatten() {
        match attr.key.local_name().as_ref() {
            b"numFmtId" => {
                if let Ok(value) = attr.decode_and_unescape_value(reader.decoder()) {
                    style.num_fmt_id = value.parse::<u32>().ok();
                }
            },
            b"fontId" => {
                if let Ok(value) = attr.decode_and_unescape_value(reader.decoder()) {
                    style.font_id = value.parse::<u32>().ok();
                }
            },
            b"fillId" => {
                if let Ok(value) = attr.decode_and_unescape_value(reader.decoder()) {
                    style.fill_id = value.parse::<u32>().ok();
                }
            },
            b"borderId" => {
                if let Ok(value) = attr.decode_and_unescape_value(reader.decoder()) {
                    style.border_id = value.parse::<u32>().ok();
                }
            },
            b"xfId" => {
                if let Ok(value) = attr.decode_and_unescape_value(reader.decoder()) {
                    style.xf_id = value.parse::<u32>().ok();
                }
            },
            b"applyNumberFormat" => {
                if let Ok(value) = attr.decode_and_unescape_value(reader.decoder()) {
                    style.apply_number_format = value == "1" || value == "true";
                }
            },
            b"applyFont" => {
                if let Ok(value) = attr.decode_and_unescape_value(reader.decoder()) {
                    style.apply_font = value == "1" || value == "true";
                }
            },
            b"applyFill" => {
                if let Ok(value) = attr.decode_and_unescape_value(reader.decoder()) {
                    style.apply_fill = value == "1" || value == "true";
                }
            },
            b"applyBorder" => {
                if let Ok(value) = attr.decode_and_unescape_value(reader.decoder()) {
                    style.apply_border = value == "1" || value == "true";
                }
            },
            b"applyAlignment" => {
                if let Ok(value) = attr.decode_and_unescape_value(reader.decoder()) {
                    style.apply_alignment = value == "1" || value == "true";
                }
            },
            b"quotePrefix" => {
                if let Ok(value) = attr.decode_and_unescape_value(reader.decoder()) {
                    style.quote_prefix = value == "1" || value == "true";
                }
            },
            _ => {},
        }
    }

    // Parse child elements
    let mut buf = Vec::with_capacity(256);
    loop {
        buf.clear();
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e))
                if e.local_name().as_ref() == b"alignment" =>
            {
                style.alignment = Some(parse_alignment(reader, &e)?);
            },
            Ok(Event::End(e)) if e.local_name().as_ref() == b"xf" => break,
            Ok(Event::Eof) => break,
            Err(e) => return Err(OoxmlError::Xml(format!("XML error in xf: {}", e))),
            _ => {},
        }
    }

    Ok(style)
}

/// Parse alignment element.
fn parse_alignment(
    reader: &mut Reader<&[u8]>,
    start: &quick_xml::events::BytesStart,
) -> Result<Alignment> {
    let mut alignment = Alignment::new();

    for attr in start.attributes().flatten() {
        match attr.key.local_name().as_ref() {
            b"horizontal" => {
                if let Ok(value) = attr.decode_and_unescape_value(reader.decoder()) {
                    alignment.horizontal = Some(value.to_string());
                }
            },
            b"vertical" => {
                if let Ok(value) = attr.decode_and_unescape_value(reader.decoder()) {
                    alignment.vertical = Some(value.to_string());
                }
            },
            b"textRotation" => {
                if let Ok(value) = attr.decode_and_unescape_value(reader.decoder()) {
                    alignment.text_rotation = value.parse::<u32>().ok();
                }
            },
            b"wrapText" => {
                if let Ok(value) = attr.decode_and_unescape_value(reader.decoder()) {
                    alignment.wrap_text = value == "1" || value == "true";
                }
            },
            b"indent" => {
                if let Ok(value) = attr.decode_and_unescape_value(reader.decoder()) {
                    alignment.indent = value.parse::<u32>().ok();
                }
            },
            b"shrinkToFit" => {
                if let Ok(value) = attr.decode_and_unescape_value(reader.decoder()) {
                    alignment.shrink_to_fit = value == "1" || value == "true";
                }
            },
            b"readingOrder" => {
                if let Ok(value) = attr.decode_and_unescape_value(reader.decoder()) {
                    alignment.reading_order = value.parse::<u32>().ok();
                }
            },
            _ => {},
        }
    }

    Ok(alignment)
}

/// Parse color from a color element.
///
/// Colors can be specified as:
/// - RGB hex value (rgb attribute)
/// - Theme color (theme attribute with optional tint)
/// - Indexed color (indexed attribute)
/// - Auto color
fn parse_color(
    reader: &mut Reader<&[u8]>,
    start: &quick_xml::events::BytesStart,
) -> Result<Option<String>> {
    for attr in start.attributes().flatten() {
        match attr.key.local_name().as_ref() {
            b"rgb" => {
                if let Ok(value) = attr.decode_and_unescape_value(reader.decoder()) {
                    return Ok(Some(format!("#{}", value)));
                }
            },
            b"theme" => {
                if let Ok(value) = attr.decode_and_unescape_value(reader.decoder()) {
                    // For now, just store theme reference as-is
                    // A full implementation would resolve theme colors
                    return Ok(Some(format!("theme:{}", value)));
                }
            },
            b"indexed" => {
                if let Ok(value) = attr.decode_and_unescape_value(reader.decoder()) {
                    return Ok(Some(format!("indexed:{}", value)));
                }
            },
            b"auto" => {
                return Ok(Some("auto".to_string()));
            },
            _ => {},
        }
    }

    Ok(None)
}
