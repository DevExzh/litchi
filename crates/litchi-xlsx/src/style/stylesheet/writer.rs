//! SpreadsheetML stylesheet publication.
//!
//! The writer consumes the same semantic stylesheet records as the parser.
//! It intentionally does not introduce a second builder model: callers can
//! assemble or edit [`super::Styles`] and publish that canonical collection.

use std::fmt::{self, Write as _};

use litchi_core::xml::escape_xml;

use crate::conditional_formatting::Differential;
use crate::error::{Result, invalid};

use super::alignment::Alignment;
use super::border::{Border, Color, Conformance, Side};
use super::cell_style::CellStyle;
use super::fill::Fill;
use super::font::{Font, FontColor, FontColorKind};
use super::{NumberFormat, Styles};

const TRANSITIONAL_NAMESPACE: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT_NAMESPACE: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";

/// Write a stylesheet using Transitional SpreadsheetML vocabulary.
pub fn write(styles: &Styles) -> Result<Vec<u8>> {
    write_in(styles, Conformance::Transitional)
}

/// Write a stylesheet using the requested SpreadsheetML conformance.
///
/// Transitional publication uses physical `left`/`right` border edges;
/// Strict publication uses logical `start`/`end` edges. A stylesheet that
/// contains the other vocabulary is rejected rather than silently losing a
/// border while changing conformance.
pub fn write_in(styles: &Styles, conformance: Conformance) -> Result<Vec<u8>> {
    let namespace = match conformance {
        Conformance::Transitional => TRANSITIONAL_NAMESPACE,
        Conformance::Strict => STRICT_NAMESPACE,
    };
    let mut xml = String::with_capacity(4096);
    xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    write_xml(
        &mut xml,
        format_args!(r#"<styleSheet xmlns="{namespace}">"#),
    )?;

    write_number_formats(&mut xml, styles)?;
    write_xml(
        &mut xml,
        format_args!(r#"<fonts count="{}">"#, styles.fonts.len()),
    )?;
    for font in &styles.fonts {
        write_font(&mut xml, font)?;
    }
    xml.push_str("</fonts>");

    write_xml(
        &mut xml,
        format_args!(r#"<fills count="{}">"#, styles.fills.len()),
    )?;
    for fill in &styles.fills {
        write_fill(&mut xml, fill)?;
    }
    xml.push_str("</fills>");

    write_xml(
        &mut xml,
        format_args!(r#"<borders count="{}">"#, styles.borders.len()),
    )?;
    for border in &styles.borders {
        write_border(&mut xml, border, conformance)?;
    }
    xml.push_str("</borders>");

    write_xfs(&mut xml, "cellStyleXfs", &styles.cell_styles, conformance)?;
    write_xfs(&mut xml, "cellXfs", &styles.cell_xfs, conformance)?;

    // The semantic stylesheet model intentionally does not invent names for
    // cellStyleXfs. An empty cellStyles section keeps the package schema
    // explicit while avoiding a duplicate public cell-style-name model.
    xml.push_str(r#"<cellStyles count="0"/>"#);
    write_differential_formats(&mut xml, &styles.differential_formats)?;
    xml.push_str("</styleSheet>");
    Ok(xml.into_bytes())
}

fn write_number_formats(xml: &mut String, styles: &Styles) -> Result<()> {
    if styles.number_formats.is_empty() {
        return Ok(());
    }
    let mut formats: Vec<&NumberFormat> = styles.number_formats.values().collect();
    formats.sort_unstable_by_key(|format| format.id());
    write_xml(xml, format_args!(r#"<numFmts count="{}">"#, formats.len()))?;
    for format in formats {
        let mut element = String::from("<numFmt");
        attr(&mut element, "numFmtId", &format.id().to_string())?;
        attr(&mut element, "formatCode", format.code())?;
        element.push_str("/>");
        xml.push_str(&element);
    }
    xml.push_str("</numFmts>");
    Ok(())
}

fn write_font(xml: &mut String, font: &Font) -> Result<()> {
    xml.push_str("<font>");
    if let Some(name) = &font.name {
        empty_attribute_element(xml, "name", "val", name)?;
    }
    if let Some(size) = font.size {
        if !size.is_finite() || size <= 0.0 {
            return Err(invalid("font size must be finite and positive"));
        }
        empty_attribute_element(xml, "sz", "val", &size.to_string())?;
    }
    if font.bold {
        xml.push_str("<b/>");
    }
    if font.italic {
        xml.push_str("<i/>");
    }
    if font.strike {
        xml.push_str("<strike/>");
    }
    if let Some(underline) = font.underline {
        empty_attribute_element(xml, "u", "val", underline.as_str())?;
    }
    if let Some(color) = &font.color {
        write_font_color(xml, "color", color)?;
    }
    if let Some(charset) = font.charset {
        empty_attribute_element(xml, "charset", "val", &charset.to_string())?;
    }
    if let Some(family) = font.family {
        empty_attribute_element(xml, "family", "val", &family.to_string())?;
    }
    if let Some(scheme) = font.scheme {
        empty_attribute_element(xml, "scheme", "val", scheme.as_str())?;
    }
    if let Some(script) = font.script {
        empty_attribute_element(xml, "vertAlign", "val", script.as_str())?;
    }
    xml.push_str("</font>");
    Ok(())
}

fn write_fill(xml: &mut String, fill: &Fill) -> Result<()> {
    xml.push_str("<fill>");
    match fill {
        Fill::None => xml.push_str(r#"<patternFill patternType="none"/>"#),
        Fill::Pattern {
            pattern_type,
            fg_color,
            bg_color,
        } => {
            validate_pattern(pattern_type)?;
            let mut element = String::from("<patternFill");
            attr(&mut element, "patternType", pattern_type)?;
            if fg_color.is_none() && bg_color.is_none() {
                element.push_str("/>");
                xml.push_str(&element);
            } else {
                element.push('>');
                xml.push_str(&element);
                if let Some(color) = fg_color {
                    write_string_color(xml, "fgColor", color)?;
                }
                if let Some(color) = bg_color {
                    write_string_color(xml, "bgColor", color)?;
                }
                xml.push_str("</patternFill>");
            }
        },
        Fill::Gradient {
            gradient_type,
            stops,
        } => {
            if let Some(kind) = gradient_type {
                if !matches!(kind.as_str(), "linear" | "path") {
                    return Err(invalid(format!("invalid gradient type '{kind}'")));
                }
            }
            let mut element = String::from("<gradientFill");
            if let Some(kind) = gradient_type {
                attr(&mut element, "type", kind)?;
            }
            element.push('>');
            xml.push_str(&element);
            for (position, color) in stops {
                if !position.is_finite() || !(0.0..=1.0).contains(position) {
                    return Err(invalid(
                        "gradient-stop position must be finite and between 0 and 1",
                    ));
                }
                let mut stop = String::from("<stop");
                attr(&mut stop, "position", &position.to_string())?;
                stop.push('>');
                xml.push_str(&stop);
                write_string_color(xml, "color", color)?;
                xml.push_str("</stop>");
            }
            xml.push_str("</gradientFill>");
        },
    }
    xml.push_str("</fill>");
    Ok(())
}

fn write_border(xml: &mut String, border: &Border, conformance: Conformance) -> Result<()> {
    if matches!(conformance, Conformance::Transitional)
        && (border.start.is_some() || border.end.is_some())
    {
        return Err(invalid(
            "strict start/end borders require Strict SpreadsheetML",
        ));
    }
    if matches!(conformance, Conformance::Strict)
        && (border.left.is_some() || border.right.is_some())
    {
        return Err(invalid(
            "physical left/right borders require Transitional SpreadsheetML",
        ));
    }

    let mut element = String::from("<border");
    if let Some(direction) = border.diagonal.as_ref().and_then(|diagonal| diagonal.dir()) {
        if direction.is_up() {
            attr(&mut element, "diagonalUp", "1")?;
        }
        if direction.is_down() {
            attr(&mut element, "diagonalDown", "1")?;
        }
    }
    if let Some(outline) = border.outline {
        attr(&mut element, "outline", if outline { "1" } else { "0" })?;
    }
    element.push('>');
    xml.push_str(&element);

    match conformance {
        Conformance::Transitional => {
            write_side(xml, "left", border.left.as_ref())?;
            write_side(xml, "right", border.right.as_ref())?;
        },
        Conformance::Strict => {
            write_side(xml, "start", border.start.as_ref())?;
            write_side(xml, "end", border.end.as_ref())?;
        },
    }
    write_side(xml, "top", border.top.as_ref())?;
    write_side(xml, "bottom", border.bottom.as_ref())?;
    if let Some(diagonal) = &border.diagonal {
        write_side(xml, "diagonal", diagonal.side())?;
    }
    write_side(xml, "vertical", border.vertical.as_ref())?;
    write_side(xml, "horizontal", border.horizontal.as_ref())?;
    xml.push_str("</border>");
    Ok(())
}

fn write_side(xml: &mut String, name: &str, side: Option<&Side>) -> Result<()> {
    let Some(side) = side else {
        write_xml(xml, format_args!("<{name}/>"))
            .map_err(|_| invalid("styles XML formatting failed"))?;
        return Ok(());
    };
    let mut element = format!("<{name}");
    attr(&mut element, "style", side.line.as_str())?;
    element.push('>');
    xml.push_str(&element);
    if let Some(color) = side.color {
        write_border_color(xml, color)?;
    }
    write_xml(xml, format_args!("</{name}>"))
        .map_err(|_| invalid("styles XML formatting failed"))?;
    Ok(())
}

fn write_border_color(xml: &mut String, color: Color) -> Result<()> {
    let mut element = String::from("<color");
    match color {
        Color::Default { .. } => {},
        Color::Rgb { value, .. } => attr(&mut element, "rgb", &value.to_string())?,
        Color::Theme { index, .. } => attr(&mut element, "theme", &index.to_string())?,
        Color::Indexed { index, .. } => attr(&mut element, "indexed", &index.to_string())?,
        Color::Auto { enabled, .. } => {
            attr(&mut element, "auto", if enabled { "1" } else { "0" })?;
        },
    }
    if let Some(tint) = color.tint() {
        attr(&mut element, "tint", &tint.get().to_string())?;
    }
    element.push_str("/>");
    xml.push_str(&element);
    Ok(())
}

fn write_xfs(
    xml: &mut String,
    section: &str,
    styles: &[CellStyle],
    conformance: Conformance,
) -> Result<()> {
    write_xml(xml, format_args!(r#"<{section} count="{}">"#, styles.len()))?;
    for style in styles {
        let mut element = String::from("<xf");
        if let Some(value) = style.num_fmt_id {
            attr(&mut element, "numFmtId", &value.to_string())?;
        }
        if let Some(value) = style.font_id {
            attr(&mut element, "fontId", &value.to_string())?;
        }
        if let Some(value) = style.fill_id {
            attr(&mut element, "fillId", &value.to_string())?;
        }
        if let Some(value) = style.border_id {
            attr(&mut element, "borderId", &value.to_string())?;
        }
        if let Some(value) = style.xf_id {
            attr(&mut element, "xfId", &value.to_string())?;
        }
        if style.apply_number_format {
            attr(&mut element, "applyNumberFormat", "1")?;
        }
        if style.apply_font {
            attr(&mut element, "applyFont", "1")?;
        }
        if style.apply_fill {
            attr(&mut element, "applyFill", "1")?;
        }
        if style.apply_border {
            attr(&mut element, "applyBorder", "1")?;
        }
        if style.apply_alignment {
            attr(&mut element, "applyAlignment", "1")?;
        }
        if style.quote_prefix {
            attr(&mut element, "quotePrefix", "1")?;
        }
        if let Some(alignment) = style.alignment {
            element.push('>');
            xml.push_str(&element);
            write_alignment(xml, alignment, conformance)?;
            xml.push_str("</xf>");
        } else {
            element.push_str("/>");
            xml.push_str(&element);
        }
    }
    write_xml(xml, format_args!("</{section}>"))
        .map_err(|_| invalid("styles XML formatting failed"))?;
    Ok(())
}

fn write_alignment(xml: &mut String, alignment: Alignment, conformance: Conformance) -> Result<()> {
    if matches!(conformance, Conformance::Strict)
        && alignment
            .text_rotation
            .is_some_and(|rotation| rotation.is_contextual())
    {
        return Err(invalid(
            "context-dependent text rotation 254 requires Transitional SpreadsheetML",
        ));
    }
    let mut element = String::from("<alignment");
    if let Some(value) = alignment.horizontal {
        attr(&mut element, "horizontal", value.as_str())?;
    }
    if let Some(value) = alignment.vertical {
        attr(&mut element, "vertical", value.as_str())?;
    }
    if let Some(value) = alignment.text_rotation {
        attr(&mut element, "textRotation", &value.get().to_string())?;
    }
    if alignment.wrap_text {
        attr(&mut element, "wrapText", "1")?;
    }
    if let Some(value) = alignment.indent {
        attr(&mut element, "indent", &value.get().to_string())?;
    }
    if let Some(value) = alignment.relative_indent {
        attr(&mut element, "relativeIndent", &value.to_string())?;
    }
    if alignment.justify_last_line {
        attr(&mut element, "justifyLastLine", "1")?;
    }
    if alignment.shrink_to_fit {
        attr(&mut element, "shrinkToFit", "1")?;
    }
    if let Some(value) = alignment.reading_order {
        attr(&mut element, "readingOrder", &value.get().to_string())?;
    }
    element.push_str("/>");
    xml.push_str(&element);
    Ok(())
}

fn write_differential_formats(xml: &mut String, formats: &[Differential]) -> Result<()> {
    if formats.is_empty() {
        xml.push_str(r#"<dxfs count="0"/>"#);
        return Ok(());
    }
    write_xml(xml, format_args!(r#"<dxfs count="{}">"#, formats.len()))?;
    for format in formats {
        let raw = format.raw_xml();
        if raw.is_empty() {
            xml.push_str("<dxf/>");
        } else {
            let value = std::str::from_utf8(raw)
                .map_err(|_| invalid("differential-format XML is not UTF-8"))?;
            xml.push_str(value);
        }
    }
    xml.push_str("</dxfs>");
    Ok(())
}

fn write_string_color(xml: &mut String, element: &str, value: &str) -> Result<()> {
    let mut output = format!("<{element}");
    if value.eq_ignore_ascii_case("auto") {
        attr(&mut output, "auto", "1")?;
    } else if let Some(index) = value.strip_prefix("theme:") {
        index
            .parse::<u32>()
            .map_err(|_| invalid(format!("invalid theme color '{value}'")))?;
        attr(&mut output, "theme", index)?;
    } else if let Some(index) = value.strip_prefix("indexed:") {
        index
            .parse::<u32>()
            .map_err(|_| invalid(format!("invalid indexed color '{value}'")))?;
        attr(&mut output, "indexed", index)?;
    } else {
        let rgb = value.strip_prefix('#').unwrap_or(value);
        if !matches!(rgb.len(), 6 | 8) || !rgb.bytes().all(|byte| byte.is_ascii_hexdigit()) {
            return Err(invalid(format!("invalid RGB color '{value}'")));
        }
        attr(&mut output, "rgb", rgb)?;
    }
    output.push_str("/>");
    xml.push_str(&output);
    Ok(())
}

fn write_font_color(xml: &mut String, element: &str, color: &FontColor) -> Result<()> {
    let mut output = format!("<{element}");
    match color.kind() {
        FontColorKind::Default => {},
        FontColorKind::Rgb(value) => attr(&mut output, "rgb", &value.to_string())?,
        FontColorKind::Theme(index) => attr(&mut output, "theme", &index.to_string())?,
        FontColorKind::Indexed(index) => attr(&mut output, "indexed", &index.to_string())?,
        FontColorKind::Auto(enabled) => {
            attr(&mut output, "auto", if enabled { "1" } else { "0" })?;
        },
    }
    if let Some(tint) = color.tint() {
        attr(&mut output, "tint", &tint.get().to_string())?;
    }
    output.push_str("/>");
    xml.push_str(&output);
    Ok(())
}

fn validate_pattern(pattern: &str) -> Result<()> {
    if matches!(
        pattern,
        "none"
            | "solid"
            | "mediumGray"
            | "darkGray"
            | "lightGray"
            | "gray125"
            | "gray0625"
            | "darkHorizontal"
            | "darkVertical"
            | "darkDown"
            | "darkUp"
            | "darkGrid"
            | "darkTrellis"
    ) {
        Ok(())
    } else {
        Err(invalid(format!("invalid fill pattern '{pattern}'")))
    }
}

fn empty_attribute_element(xml: &mut String, element: &str, name: &str, value: &str) -> Result<()> {
    let mut output = format!("<{element}");
    attr(&mut output, name, value)?;
    output.push_str("/>");
    xml.push_str(&output);
    Ok(())
}

fn attr(xml: &mut String, name: &str, value: &str) -> Result<()> {
    write_xml(xml, format_args!(r#" {name}="{}""#, escape_xml(value)))
}

fn write_xml(xml: &mut String, arguments: fmt::Arguments<'_>) -> Result<()> {
    xml.write_fmt(arguments)
        .map_err(|_| invalid("styles XML formatting failed"))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::style::stylesheet::alignment::{Horizontal, Vertical};
    use crate::style::stylesheet::border::{Line, Tint};

    #[test]
    fn publishes_and_parses_canonical_stylesheet_records() {
        let mut styles = Styles::new();
        styles.fonts.push(Font {
            name: Some("A&B".into()),
            size: Some(11.0),
            bold: true,
            color: Some(FontColor::rgb(crate::color::Rgb::argb(
                0xFF, 0x11, 0x22, 0x33,
            ))),
            ..Font::default()
        });
        styles.fills.push(Fill::solid("#FFFF0000".into()));
        styles.borders.push(Border {
            left: Some(
                Side::new(Line::Thin)
                    .with_color(Color::rgb(0x10, 0x20, 0x30).with_tint(Tint::new(0.25).unwrap())),
            ),
            ..Border::default()
        });
        styles.cell_xfs.push(CellStyle {
            font_id: Some(0),
            fill_id: Some(0),
            border_id: Some(0),
            alignment: Some(Alignment::both(Horizontal::Center, Vertical::Bottom)),
            apply_font: true,
            apply_alignment: true,
            ..CellStyle::default()
        });

        let xml = write(&styles).unwrap();
        let xml = String::from_utf8(xml).unwrap();
        let parsed = Styles::parse(&xml).unwrap();
        assert_eq!(parsed.fonts.len(), 1);
        assert_eq!(parsed.fills.len(), 1);
        assert_eq!(parsed.borders.len(), 1);
        assert_eq!(parsed.cell_xfs.len(), 1);
        assert_eq!(parsed.cell_xfs[0].alignment, styles.cell_xfs[0].alignment);
    }

    #[test]
    fn rejects_conformance_mismatches_instead_of_dropping_edges() {
        let styles = Styles {
            borders: vec![Border {
                start: Some(Side::new(Line::Thin)),
                ..Border::default()
            }],
            ..Styles::default()
        };
        assert!(write_in(&styles, Conformance::Transitional).is_err());
    }
}
