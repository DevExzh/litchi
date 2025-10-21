//! Parser for Excel styles.xml files.
//!
//! This module provides parsing functionality for the styles.xml file
//! which contains formatting information for cells, fonts, etc.

use std::collections::HashMap;

use crate::ooxml::xlsx::styles::{Border, CellStyle, Fill, Font, NumberFormat, Styles};
use crate::sheet::Result;

/// Parse styles.xml content to extract style information.
pub fn parse_styles_xml(content: &str) -> Result<Styles> {
    let mut styles = Styles::default();

    // Parse number formats
    if let Some(num_fmts_start) = content.find("<numFmts")
        && let Some(num_fmts_end) = content[num_fmts_start..].find("</numFmts>")
    {
        let num_fmts_content = &content[num_fmts_start..num_fmts_start + num_fmts_end];
        parse_number_formats(num_fmts_content, &mut styles.number_formats)?;
    }

    // Parse fonts
    if let Some(fonts_start) = content.find("<fonts")
        && let Some(fonts_end) = content[fonts_start..].find("</fonts>")
    {
        let fonts_content = &content[fonts_start..fonts_start + fonts_end];
        styles.fonts = parse_fonts(fonts_content)?;
    }

    // Parse fills
    if let Some(fills_start) = content.find("<fills")
        && let Some(fills_end) = content[fills_start..].find("</fills>")
    {
        let fills_content = &content[fills_start..fills_start + fills_end];
        styles.fills = parse_fills(fills_content)?;
    }

    // Parse borders
    if let Some(borders_start) = content.find("<borders")
        && let Some(borders_end) = content[borders_start..].find("</borders>")
    {
        let borders_content = &content[borders_start..borders_start + borders_end];
        styles.borders = parse_borders(borders_content)?;
    }

    // Parse cell style formats
    if let Some(cell_style_xfs_start) = content.find("<cellStyleXfs")
        && let Some(cell_style_xfs_end) = content[cell_style_xfs_start..].find("</cellStyleXfs>")
    {
        let cell_style_xfs_content =
            &content[cell_style_xfs_start..cell_style_xfs_start + cell_style_xfs_end];
        styles.cell_styles = parse_cell_styles(cell_style_xfs_content)?;
    }

    // Parse cell XFs
    if let Some(cell_xfs_start) = content.find("<cellXfs")
        && let Some(cell_xfs_end) = content[cell_xfs_start..].find("</cellXfs>")
    {
        let cell_xfs_content = &content[cell_xfs_start..cell_xfs_start + cell_xfs_end];
        styles.cell_xfs = parse_cell_xfs(cell_xfs_content)?;
    }

    Ok(styles)
}

/// Parse number formats section.
fn parse_number_formats(
    content: &str,
    number_formats: &mut HashMap<u32, NumberFormat>,
) -> Result<()> {
    let mut pos = 0;
    while let Some(num_fmt_start) = content[pos..].find("<numFmt ") {
        let num_fmt_start_pos = pos + num_fmt_start;
        if let Some(num_fmt_end) = content[num_fmt_start_pos..].find("/>") {
            let num_fmt_xml = &content[num_fmt_start_pos..num_fmt_start_pos + num_fmt_end + 2];

            if let Some(fmt) = parse_number_format(num_fmt_xml)? {
                number_formats.insert(fmt.id, fmt);
            }
            pos = num_fmt_start_pos + num_fmt_end + 2;
        } else {
            break;
        }
    }
    Ok(())
}

/// Parse individual number format.
fn parse_number_format(num_fmt_xml: &str) -> Result<Option<NumberFormat>> {
    let id = if let Some(id_start) = num_fmt_xml.find("numFmtId=\"") {
        let id_content = &num_fmt_xml[id_start + 10..];
        if let Some(quote_pos) = id_content.find('"') {
            id_content[..quote_pos].parse::<u32>().ok()
        } else {
            None
        }
    } else {
        None
    };

    let code = if let Some(code_start) = num_fmt_xml.find("formatCode=\"") {
        let code_content = &num_fmt_xml[code_start + 12..];
        code_content
            .find('"')
            .map(|quote_pos| code_content[..quote_pos].to_string())
    } else {
        None
    };

    match (id, code) {
        (Some(id), Some(code)) => Ok(Some(NumberFormat { id, code })),
        _ => Ok(None),
    }
}

/// Parse fonts section.
fn parse_fonts(content: &str) -> Result<Vec<Font>> {
    let mut fonts = Vec::new();
    let mut pos = 0;

    while let Some(font_start) = content[pos..].find("<font>") {
        let font_start_pos = pos + font_start;
        if let Some(font_end) = content[font_start_pos..].find("</font>") {
            let font_xml = &content[font_start_pos..font_start_pos + font_end + 7];
            fonts.push(parse_font(font_xml));
            pos = font_start_pos + font_end + 7;
        } else {
            break;
        }
    }

    Ok(fonts)
}

/// Parse individual font.
fn parse_font(font_xml: &str) -> Font {
    let mut font = Font::default();

    // Parse font name
    if let Some(name_start) = font_xml.find("<name val=\"") {
        let name_content = &font_xml[name_start + 11..];
        if let Some(quote_pos) = name_content.find('"') {
            font.name = Some(name_content[..quote_pos].to_string());
        }
    }

    // Parse font size
    if let Some(sz_start) = font_xml.find("<sz val=\"") {
        let sz_content = &font_xml[sz_start + 9..];
        if let Some(quote_pos) = sz_content.find('"')
            && let Ok(size) = sz_content[..quote_pos].parse::<f64>()
        {
            font.size = Some(size);
        }
    }

    // Parse bold
    font.bold = font_xml.contains("<b/>") || font_xml.contains("<b>");

    // Parse italic
    font.italic = font_xml.contains("<i/>") || font_xml.contains("<i>");

    font
}

/// Parse fills section.
fn parse_fills(content: &str) -> Result<Vec<Fill>> {
    let mut fills = Vec::new();
    let mut pos = 0;

    while let Some(fill_start) = content[pos..].find("<fill>") {
        let fill_start_pos = pos + fill_start;
        if let Some(fill_end) = content[fill_start_pos..].find("</fill>") {
            let fill_xml = &content[fill_start_pos..fill_start_pos + fill_end + 7];
            fills.push(parse_fill(fill_xml));
            pos = fill_start_pos + fill_end + 7;
        } else {
            break;
        }
    }

    Ok(fills)
}

/// Parse individual fill.
fn parse_fill(fill_xml: &str) -> Fill {
    if let Some(pattern_start) = fill_xml.find("<patternFill ") {
        let pattern_content = &fill_xml[pattern_start..];

        let pattern_type = if let Some(pt_start) = pattern_content.find("patternType=\"") {
            let pt_content = &pattern_content[pt_start + 13..];
            if let Some(quote_pos) = pt_content.find('"') {
                pt_content[..quote_pos].to_string()
            } else {
                "solid".to_string()
            }
        } else {
            "solid".to_string()
        };

        let fg_color = parse_color(pattern_content, "fgColor");
        let bg_color = parse_color(pattern_content, "bgColor");

        Fill::Pattern {
            pattern_type,
            fg_color,
            bg_color,
        }
    } else {
        Fill::Pattern {
            pattern_type: "solid".to_string(),
            fg_color: None,
            bg_color: None,
        }
    }
}

/// Parse borders section.
fn parse_borders(content: &str) -> Result<Vec<Border>> {
    let mut borders = Vec::new();
    let mut pos = 0;

    while let Some(border_start) = content[pos..].find("<border>") {
        let border_start_pos = pos + border_start;
        if let Some(border_end) = content[border_start_pos..].find("</border>") {
            let border_xml = &content[border_start_pos..border_start_pos + border_end + 9];
            borders.push(parse_border(border_xml));
            pos = border_start_pos + border_end + 9;
        } else {
            break;
        }
    }

    Ok(borders)
}

/// Parse individual border.
fn parse_border(border_xml: &str) -> Border {
    Border {
        left: parse_border_style(border_xml, "left"),
        right: parse_border_style(border_xml, "right"),
        top: parse_border_style(border_xml, "top"),
        bottom: parse_border_style(border_xml, "bottom"),
    }
}

/// Parse cell styles section.
fn parse_cell_styles(content: &str) -> Result<Vec<CellStyle>> {
    let mut styles = Vec::new();
    let mut pos = 0;

    while let Some(xf_start) = content[pos..].find("<xf ") {
        let xf_start_pos = pos + xf_start;
        if let Some(xf_end) = content[xf_start_pos..].find("/>") {
            let xf_xml = &content[xf_start_pos..xf_start_pos + xf_end + 2];
            if let Some(style) = parse_cell_style(xf_xml)? {
                styles.push(style);
            }
            pos = xf_start_pos + xf_end + 2;
        } else {
            break;
        }
    }

    Ok(styles)
}

/// Parse cell XFs section.
fn parse_cell_xfs(content: &str) -> Result<Vec<CellStyle>> {
    let mut xfs = Vec::new();
    let mut pos = 0;

    while let Some(xf_start) = content[pos..].find("<xf ") {
        let xf_start_pos = pos + xf_start;
        if let Some(xf_end) = content[xf_start_pos..].find("/>") {
            let xf_xml = &content[xf_start_pos..xf_start_pos + xf_end + 2];
            if let Some(xf) = parse_cell_style(xf_xml)? {
                xfs.push(xf);
            }
            pos = xf_start_pos + xf_end + 2;
        } else {
            break;
        }
    }

    Ok(xfs)
}

/// Parse individual cell style/XF.
fn parse_cell_style(xf_xml: &str) -> Result<Option<CellStyle>> {
    let num_fmt_id = if let Some(nf_start) = xf_xml.find("numFmtId=\"") {
        let nf_content = &xf_xml[nf_start + 10..];
        if let Some(quote_pos) = nf_content.find('"') {
            nf_content[..quote_pos].parse::<u32>().ok()
        } else {
            None
        }
    } else {
        None
    };

    let font_id = if let Some(f_start) = xf_xml.find("fontId=\"") {
        let f_content = &xf_xml[f_start + 8..];
        if let Some(quote_pos) = f_content.find('"') {
            f_content[..quote_pos].parse::<u32>().ok()
        } else {
            None
        }
    } else {
        None
    };

    let fill_id = if let Some(fill_start) = xf_xml.find("fillId=\"") {
        let fill_content = &xf_xml[fill_start + 8..];
        if let Some(quote_pos) = fill_content.find('"') {
            fill_content[..quote_pos].parse::<u32>().ok()
        } else {
            None
        }
    } else {
        None
    };

    let border_id = if let Some(b_start) = xf_xml.find("borderId=\"") {
        let b_content = &xf_xml[b_start + 10..];
        if let Some(quote_pos) = b_content.find('"') {
            b_content[..quote_pos].parse::<u32>().ok()
        } else {
            None
        }
    } else {
        None
    };

    Ok(Some(CellStyle {
        num_fmt_id,
        font_id,
        fill_id,
        border_id,
        alignment: None, // TODO: Parse alignment
    }))
}

/// Helper function to parse border style.
fn parse_border_style(
    border_xml: &str,
    side: &str,
) -> Option<crate::ooxml::xlsx::styles::BorderStyle> {
    let side_tag = format!("<{}>", side);
    if let Some(side_start) = border_xml.find(&side_tag) {
        let side_content = &border_xml[side_start..];

        let style = if let Some(style_start) = side_content.find("style=\"") {
            let style_content = &side_content[style_start + 7..];
            style_content
                .find('"')
                .map(|quote_pos| style_content[..quote_pos].to_string())
        } else {
            None
        };

        style.map(|s| crate::ooxml::xlsx::styles::BorderStyle {
            style: s,
            color: parse_color(side_content, "color"),
        })
    } else {
        None
    }
}

/// Helper function to parse color.
fn parse_color(content: &str, color_type: &str) -> Option<String> {
    let color_tag = format!("<{} ", color_type);
    if let Some(color_start) = content.find(&color_tag) {
        let color_content = &content[color_start..];

        // Look for rgb attribute
        if let Some(rgb_start) = color_content.find("rgb=\"") {
            let rgb_content = &color_content[rgb_start + 5..];
            if let Some(quote_pos) = rgb_content.find('"') {
                return Some(rgb_content[..quote_pos].to_string());
            }
        }

        // Look for theme attribute
        if let Some(theme_start) = color_content.find("theme=\"") {
            let theme_content = &color_content[theme_start + 7..];
            if let Some(quote_pos) = theme_content.find('"') {
                return Some(format!("theme:{}", &theme_content[..quote_pos]));
            }
        }
    }

    None
}
