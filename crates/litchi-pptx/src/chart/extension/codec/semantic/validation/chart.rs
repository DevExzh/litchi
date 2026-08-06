//! Chart-space formatting and chart validation concerns for the ChartEx graph.

use super::*;
use std::collections::HashSet;

pub(super) fn parse_chart_space_formatting(root: &MiniNode) -> Result<ChartSpaceFormatting> {
    let shape_properties = one_child(root, CX, "spPr")?
        .map(|node| parse_drawing_payload(node, "chartSpace spPr"))
        .transpose()?;
    let text_properties = one_child(root, CX, "txPr")?
        .map(|node| parse_drawing_payload(node, "chartSpace txPr"))
        .transpose()?;
    let color_mapping_override = one_child(root, CX, "clrMapOvr")?
        .map(|node| parse_drawing_payload(node, "chartSpace clrMapOvr"))
        .transpose()?;
    let format_overrides = one_child(root, CX, "fmtOvrs")?
        .map(parse_format_overrides)
        .transpose()?;
    let print_settings = one_child(root, CX, "printSettings")?
        .map(parse_print_settings)
        .transpose()?;
    Ok(ChartSpaceFormatting {
        shape_properties,
        text_properties,
        color_mapping_override,
        format_overrides,
        print_settings,
        has_extension_list: one_child(root, CX, "extLst")?.is_some(),
    })
}

pub(super) fn parse_format_overrides(node: &MiniNode) -> Result<Vec<FormatOverride>> {
    reject_unknown(&node.attributes, &[], "fmtOvrs")?;
    if !node.text.trim().is_empty() {
        return invalid("unexpected text in  fmtOvrs");
    }
    let mut values = Vec::new();
    let mut indices = HashSet::new();
    for child in &node.children {
        if child.namespace != CX || child.name != "fmtOvr" {
            return invalid("invalid direct child in  fmtOvrs");
        }
        if values.len() >= MAX_FORMAT_OVERRIDES {
            return limit(" format override count");
        }
        let value = parse_format_override(child)?;
        if !indices.insert(value.index) {
            return invalid("duplicate  format override index");
        }
        values.push(value);
    }
    Ok(values)
}

pub(super) fn parse_format_override(node: &MiniNode) -> Result<FormatOverride> {
    reject_unknown(&node.attributes, &[("", "idx")], "fmtOvr")?;
    if !node.text.trim().is_empty() {
        return invalid("unexpected text in  fmtOvr");
    }
    let index = parse_u32(
        required(&node.attributes, "", "idx")?,
        "format override index",
    )?;
    let mut shape_properties = None;
    let mut has_extension_list = false;
    for child in &node.children {
        if child.namespace != CX {
            return invalid("foreign direct child in  fmtOvr");
        }
        match child.name.as_str() {
            "spPr" if shape_properties.is_none() && !has_extension_list => {
                shape_properties = Some(parse_drawing_payload(child, "fmtOvr spPr")?);
            },
            "extLst" if !has_extension_list => has_extension_list = true,
            _ => {
                return invalid(" fmtOvr children are invalid, duplicated, or out of order");
            },
        }
    }
    Ok(FormatOverride {
        index,
        shape_properties,
        has_extension_list,
    })
}

pub(super) fn parse_print_settings(node: &MiniNode) -> Result<PrintSettings> {
    reject_unknown(&node.attributes, &[], "printSettings")?;
    if !node.text.trim().is_empty() {
        return invalid("unexpected text in  printSettings");
    }
    let mut result = PrintSettings::default();
    let mut rank = 0u8;
    let mut seen = HashSet::new();
    for child in &node.children {
        if child.namespace != CX {
            return invalid("foreign direct child in  printSettings");
        }
        let current = match child.name.as_str() {
            "headerFooter" => 0,
            "pageMargins" => 1,
            "pageSetup" => 2,
            _ => return invalid("invalid direct child in  printSettings"),
        };
        if current < rank || !seen.insert(child.name.as_str()) {
            return invalid(" printSettings children are out of order or duplicated");
        }
        rank = current;
        match child.name.as_str() {
            "headerFooter" => result.header_footer = Some(parse_header_footer(child)?),
            "pageMargins" => result.page_margins = Some(parse_page_margins(child)?),
            "pageSetup" => result.page_setup = Some(parse_page_setup(child)?),
            _ => unreachable!(),
        }
    }
    Ok(result)
}

pub(super) fn parse_header_footer(node: &MiniNode) -> Result<HeaderFooter> {
    reject_unknown(
        &node.attributes,
        &[
            ("", "alignWithMargins"),
            ("", "differentOddEven"),
            ("", "differentFirst"),
        ],
        "headerFooter",
    )?;
    if !node.text.trim().is_empty() {
        return invalid("unexpected text in  headerFooter");
    }
    let mut texts: [Option<String>; 6] = Default::default();
    let mut rank = 0usize;
    for child in &node.children {
        if child.namespace != CX {
            return invalid("foreign direct child in  headerFooter");
        }
        let current = match child.name.as_str() {
            "oddHeader" => 0,
            "oddFooter" => 1,
            "evenHeader" => 2,
            "evenFooter" => 3,
            "firstHeader" => 4,
            "firstFooter" => 5,
            _ => return invalid("invalid direct child in  headerFooter"),
        };
        if current < rank || texts[current].is_some() {
            return invalid(" headerFooter children are out of order or duplicated");
        }
        rank = current;
        texts[current] = Some(parse_print_text(child)?);
    }
    Ok(HeaderFooter {
        align_with_margins: optional(&node.attributes, "", "alignWithMargins")
            .map(parse_bool)
            .transpose()?
            .unwrap_or(true),
        different_odd_even: optional(&node.attributes, "", "differentOddEven")
            .map(parse_bool)
            .transpose()?
            .unwrap_or(false),
        different_first: optional(&node.attributes, "", "differentFirst")
            .map(parse_bool)
            .transpose()?
            .unwrap_or(false),
        odd_header: texts[0].take(),
        odd_footer: texts[1].take(),
        even_header: texts[2].take(),
        even_footer: texts[3].take(),
        first_header: texts[4].take(),
        first_footer: texts[5].take(),
    })
}

pub(super) fn parse_print_text(node: &MiniNode) -> Result<String> {
    reject_unknown(&node.attributes, &[], "header/footer text")?;
    if !node.children.is_empty() {
        return invalid(" header/footer text must have simple content");
    }
    if node.text.len() > MAX_PRINT_TEXT_BYTES {
        return limit(" header/footer text bytes");
    }
    Ok(node.text.clone())
}

pub(super) fn parse_page_margins(node: &MiniNode) -> Result<PageMargins> {
    reject_unknown(
        &node.attributes,
        &[
            ("", "l"),
            ("", "r"),
            ("", "t"),
            ("", "b"),
            ("", "header"),
            ("", "footer"),
        ],
        "pageMargins",
    )?;
    require_empty_content(node, "pageMargins")?;
    let value = |name| -> Result<String> {
        let value = required(&node.attributes, "", name)?;
        if !valid_xml_double(value) {
            return invalid("invalid  page margin");
        }
        Ok(value.to_owned())
    };
    Ok(PageMargins {
        left: value("l")?,
        right: value("r")?,
        top: value("t")?,
        bottom: value("b")?,
        header: value("header")?,
        footer: value("footer")?,
    })
}

pub(super) fn parse_page_setup(node: &MiniNode) -> Result<PageSetup> {
    reject_unknown(
        &node.attributes,
        &[
            ("", "paperSize"),
            ("", "firstPageNumber"),
            ("", "orientation"),
            ("", "blackAndWhite"),
            ("", "draft"),
            ("", "useFirstPageNumber"),
            ("", "horizontalDpi"),
            ("", "verticalDpi"),
            ("", "copies"),
        ],
        "pageSetup",
    )?;
    require_empty_content(node, "pageSetup")?;
    Ok(PageSetup {
        paper_size: optional(&node.attributes, "", "paperSize")
            .map(|value| parse_u32(value, "pageSetup paperSize"))
            .transpose()?
            .unwrap_or(1),
        first_page_number: optional(&node.attributes, "", "firstPageNumber")
            .map(|value| parse_u32(value, "pageSetup firstPageNumber"))
            .transpose()?
            .unwrap_or(1),
        orientation: parse_page_orientation(
            optional(&node.attributes, "", "orientation").unwrap_or("default"),
        )?,
        black_and_white: optional(&node.attributes, "", "blackAndWhite")
            .map(parse_bool)
            .transpose()?
            .unwrap_or(false),
        draft: optional(&node.attributes, "", "draft")
            .map(parse_bool)
            .transpose()?
            .unwrap_or(false),
        use_first_page_number: optional(&node.attributes, "", "useFirstPageNumber")
            .map(parse_bool)
            .transpose()?
            .unwrap_or(false),
        horizontal_dpi: optional(&node.attributes, "", "horizontalDpi")
            .map(|value| parse_i32(value, "pageSetup horizontalDpi"))
            .transpose()?
            .unwrap_or(600),
        vertical_dpi: optional(&node.attributes, "", "verticalDpi")
            .map(|value| parse_i32(value, "pageSetup verticalDpi"))
            .transpose()?
            .unwrap_or(600),
        copies: optional(&node.attributes, "", "copies")
            .map(|value| parse_u32(value, "pageSetup copies"))
            .transpose()?
            .unwrap_or(1),
    })
}

pub(super) fn parse_page_orientation(value: &str) -> Result<PageOrientation> {
    match value {
        "default" => Ok(PageOrientation::Default),
        "portrait" => Ok(PageOrientation::Portrait),
        "landscape" => Ok(PageOrientation::Landscape),
        _ => invalid("invalid  page orientation"),
    }
}

pub(super) fn parse_chart(node: &MiniNode, offset_allowed: bool) -> Result<Chart> {
    reject_unknown(&node.attributes, &[], "chart")?;
    if !node.text.trim().is_empty() {
        return invalid("unexpected text in  chart");
    }
    let mut title = None;
    let mut legend = None;
    let mut plot_area_seen = false;
    let mut has_extension_list = false;
    let mut rank = 0u8;
    let mut seen = HashSet::new();
    for child in &node.children {
        if child.namespace != CX {
            return invalid("foreign direct child in  chart");
        }
        let current = match child.name.as_str() {
            "title" => 0,
            "plotArea" => 1,
            "legend" => 2,
            "extLst" => 3,
            _ => return invalid("invalid direct child in  chart"),
        };
        if current < rank || !seen.insert(child.name.as_str()) {
            return invalid(" chart children are out of order or duplicated");
        }
        rank = current;
        match child.name.as_str() {
            "title" => title = Some(parse_chart_title(child, offset_allowed)?),
            "plotArea" => plot_area_seen = true,
            "legend" => legend = Some(parse_legend(child, offset_allowed)?),
            "extLst" => has_extension_list = true,
            _ => unreachable!(),
        }
    }
    if !plot_area_seen {
        return invalid(" chart requires plotArea");
    }
    Ok(Chart {
        title,
        legend,
        has_extension_list,
    })
}

pub(super) fn parse_chart_title(node: &MiniNode, offset_allowed: bool) -> Result<ChartTitle> {
    reject_unknown(
        &node.attributes,
        &[("", "pos"), ("", "align"), ("", "overlay")],
        "chart title",
    )?;
    if !node.text.trim().is_empty() {
        return invalid("unexpected text in  chart title");
    }
    let position = parse_side_position(optional(&node.attributes, "", "pos").unwrap_or("t"))?;
    let alignment =
        parse_position_alignment(optional(&node.attributes, "", "align").unwrap_or("ctr"))?;
    let overlay = optional(&node.attributes, "", "overlay")
        .map(parse_bool)
        .transpose()?
        .unwrap_or(false);
    let mut text = None;
    let mut shape_properties = None;
    let mut text_properties = None;
    let mut offset = None;
    let mut rank = 0u8;
    let mut seen = HashSet::new();
    for child in &node.children {
        if child.namespace != CX {
            return invalid("foreign direct child in  chart title");
        }
        let current = match child.name.as_str() {
            "tx" => 0,
            "spPr" => 1,
            "txPr" => 2,
            "offset" => 3,
            "extLst" => 4,
            _ => return invalid("invalid direct child in  chart title"),
        };
        if current < rank || !seen.insert(child.name.as_str()) {
            return invalid(" chart title children are out of order or duplicated");
        }
        rank = current;
        match child.name.as_str() {
            "tx" => text = Some(parse_shared_text(child, "chart title tx")?),
            "spPr" => shape_properties = Some(parse_drawing_payload(child, "chart title spPr")?),
            "txPr" => text_properties = Some(parse_drawing_payload(child, "chart title txPr")?),
            "offset" => {
                if !offset_allowed {
                    return invalid(" chart title offset requires version 1.0 or feature mp");
                }
                offset = Some(parse_offset(child)?);
            },
            "extLst" => {},
            _ => unreachable!(),
        }
    }
    Ok(ChartTitle {
        position,
        alignment,
        overlay,
        text,
        shape_properties,
        text_properties,
        offset,
    })
}

pub(super) fn parse_legend(node: &MiniNode, offset_allowed: bool) -> Result<Legend> {
    reject_unknown(
        &node.attributes,
        &[("", "pos"), ("", "align"), ("", "overlay")],
        "legend",
    )?;
    if !node.text.trim().is_empty() {
        return invalid("unexpected text in  legend");
    }
    let position = parse_side_position(optional(&node.attributes, "", "pos").unwrap_or("r"))?;
    let alignment =
        parse_position_alignment(optional(&node.attributes, "", "align").unwrap_or("ctr"))?;
    let overlay = optional(&node.attributes, "", "overlay")
        .map(parse_bool)
        .transpose()?
        .unwrap_or(false);
    let mut shape_properties = None;
    let mut text_properties = None;
    let mut offset = None;
    let mut rank = 0u8;
    let mut seen = HashSet::new();
    for child in &node.children {
        if child.namespace != CX {
            return invalid("foreign direct child in  legend");
        }
        let current = match child.name.as_str() {
            "spPr" => 0,
            "txPr" => 1,
            "offset" => 2,
            "extLst" => 3,
            _ => return invalid("invalid direct child in  legend"),
        };
        if current < rank || !seen.insert(child.name.as_str()) {
            return invalid(" legend children are out of order or duplicated");
        }
        rank = current;
        match child.name.as_str() {
            "spPr" => shape_properties = Some(parse_drawing_payload(child, "legend spPr")?),
            "txPr" => text_properties = Some(parse_drawing_payload(child, "legend txPr")?),
            "offset" => {
                if !offset_allowed {
                    return invalid(" legend offset requires version 1.0 or feature mp");
                }
                offset = Some(parse_offset(child)?);
            },
            "extLst" => {},
            _ => unreachable!(),
        }
    }
    Ok(Legend {
        position,
        alignment,
        overlay,
        shape_properties,
        text_properties,
        offset,
    })
}

pub(super) fn parse_offset(node: &MiniNode) -> Result<Offset> {
    reject_unknown(&node.attributes, &[("", "top"), ("", "left")], "offset")?;
    require_empty_content(node, "offset")?;
    let top = required(&node.attributes, "", "top")?;
    let left = required(&node.attributes, "", "left")?;
    if !valid_xml_double(top) || !valid_xml_double(left) {
        return invalid("invalid  offset coordinate");
    }
    Ok(Offset {
        top: top.to_owned(),
        left: left.to_owned(),
    })
}

pub(super) fn parse_side_position(value: &str) -> Result<SidePosition> {
    match value {
        "l" => Ok(SidePosition::Left),
        "r" => Ok(SidePosition::Right),
        "t" => Ok(SidePosition::Top),
        "b" => Ok(SidePosition::Bottom),
        _ => invalid("invalid  side position"),
    }
}

pub(super) fn parse_position_alignment(value: &str) -> Result<PositionAlignment> {
    match value {
        "min" => Ok(PositionAlignment::Minimum),
        "ctr" => Ok(PositionAlignment::Center),
        "max" => Ok(PositionAlignment::Maximum),
        _ => invalid("invalid  position alignment"),
    }
}

pub(super) fn offset_feature_allowed(version: &str, features: &[String]) -> bool {
    version
        .split('.')
        .next()
        .and_then(|value| value.parse::<u32>().ok())
        .is_some_and(|major| major >= 1)
        || features.iter().any(|feature| feature == "mp")
}
