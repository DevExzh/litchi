//! Drawing, labels, and series validation concerns for the ChartEx graph.

use super::*;
use std::collections::HashSet;

pub(super) fn parse_drawing_payload(node: &MiniNode, label: &str) -> Result<DrawingPayload> {
    if !node.text.trim().is_empty() {
        return invalid(format!("unexpected text in  {label}"));
    }
    if node
        .children
        .iter()
        .any(|child| !matches!(child.namespace.as_str(), A | A_STRICT))
    {
        return invalid(format!("foreign direct child in  {label}"));
    }
    Ok(DrawingPayload {
        child_elements: node.children.len(),
        attributes: node.attributes.len(),
    })
}

pub(super) fn parse_shared_text(node: &MiniNode, label: &str) -> Result<Text> {
    reject_unknown(&node.attributes, &[], label)?;
    if !node.text.trim().is_empty() || node.children.len() != 1 {
        return invalid(format!(" {label} requires exactly one text choice"));
    }
    let child = &node.children[0];
    if child.namespace != CX {
        return invalid(format!("foreign  {label} choice"));
    }
    match child.name.as_str() {
        "txData" => parse_text_data(child),
        "rich" => Ok(Text::Rich(parse_drawing_payload(child, "rich text")?)),
        _ => invalid(format!("invalid  {label} choice")),
    }
}

pub(super) fn parse_text_data(node: &MiniNode) -> Result<Text> {
    reject_unknown(&node.attributes, &[], "txData")?;
    if !node.text.trim().is_empty() || node.children.is_empty() || node.children.len() > 2 {
        return invalid("invalid  txData choice");
    }
    let mut formula = None;
    let mut value = None;
    for (index, child) in node.children.iter().enumerate() {
        if child.namespace != CX || !matches!(child.name.as_str(), "f" | "v") {
            return invalid("invalid direct child in  txData");
        }
        match child.name.as_str() {
            "f" if index == 0 && formula.is_none() => formula = Some(parse_formula(child)?),
            "v" if value.is_none() && child.children.is_empty() && child.attributes.is_empty() => {
                if child.text.len() > MAX_LABEL_TEXT_BYTES {
                    return limit(" text value bytes");
                }
                value = Some(child.text.clone());
            },
            _ => return invalid(" txData children are out of order or duplicated"),
        }
    }
    if formula.is_none() && value.is_none() {
        return invalid(" txData is empty");
    }
    Ok(Text::Data { formula, value })
}

pub(super) fn parse_value_colors(node: &MiniNode) -> Result<ValueColors> {
    reject_unknown(&node.attributes, &[], "valueColors")?;
    if !node.text.trim().is_empty() {
        return invalid("unexpected text in  valueColors");
    }
    let mut result = ValueColors::default();
    let mut rank = 0u8;
    let mut seen = HashSet::new();
    for child in &node.children {
        if child.namespace != CX {
            return invalid("foreign direct child in  valueColors");
        }
        let current = match child.name.as_str() {
            "minColor" => 0,
            "midColor" => 1,
            "maxColor" => 2,
            _ => return invalid("invalid direct child in  valueColors"),
        };
        if current < rank || !seen.insert(child.name.as_str()) {
            return invalid(" valueColors children are out of order or duplicated");
        }
        rank = current;
        let color = parse_solid_color(child)?;
        match child.name.as_str() {
            "minColor" => result.minimum = Some(color),
            "midColor" => result.middle = Some(color),
            "maxColor" => result.maximum = Some(color),
            _ => unreachable!(),
        }
    }
    Ok(result)
}

pub(super) fn parse_solid_color(node: &MiniNode) -> Result<SolidColor> {
    reject_unknown(&node.attributes, &[], "solid color")?;
    if !node.text.trim().is_empty() || node.children.len() != 1 {
        return invalid(" solid color requires exactly one DrawingML color choice");
    }
    let color = &node.children[0];
    if !matches!(color.namespace.as_str(), A | A_STRICT) {
        return invalid(" solid color choice has the wrong namespace");
    }
    let kind = match color.name.as_str() {
        "scrgbClr" => ColorKind::ScRgb,
        "srgbClr" => ColorKind::Srgb,
        "hslClr" => ColorKind::Hsl,
        "sysClr" => ColorKind::System,
        "schemeClr" => ColorKind::Scheme,
        "prstClr" => ColorKind::Preset,
        _ => return invalid("invalid  DrawingML color choice"),
    };
    if !color.text.trim().is_empty()
        || color
            .children
            .iter()
            .any(|child| !matches!(child.namespace.as_str(), A | A_STRICT))
    {
        return invalid("invalid direct payload in  DrawingML color");
    }
    let value = optional(&color.attributes, "", "val").map(str::to_owned);
    match kind {
        ColorKind::Srgb => {
            let value = value
                .as_deref()
                .ok_or_else(|| invalid_error("missing  sRGB color value"))?;
            if value.len() != 6 || !value.bytes().all(|byte| byte.is_ascii_hexdigit()) {
                return invalid("invalid  sRGB color value");
            }
        },
        ColorKind::System | ColorKind::Scheme | ColorKind::Preset => {
            if value.as_deref().is_none_or(|value| {
                value.is_empty()
                    || value.len() > 128
                    || value.bytes().any(|byte| byte.is_ascii_whitespace())
            }) {
                return invalid("invalid  DrawingML color token");
            }
        },
        ColorKind::ScRgb | ColorKind::Hsl => {},
    }
    Ok(SolidColor {
        kind,
        value,
        modifier_count: color.children.len(),
    })
}

pub(super) fn parse_value_color_positions(node: &MiniNode) -> Result<ValueColorPositions> {
    reject_unknown(&node.attributes, &[("", "count")], "valueColorPositions")?;
    if !node.text.trim().is_empty() {
        return invalid("unexpected text in  valueColorPositions");
    }
    let count = match optional(&node.attributes, "", "count").unwrap_or("2") {
        "2" => 2,
        "3" => 3,
        _ => return invalid("invalid  valueColorPositions count"),
    };
    let mut minimum = None;
    let mut middle = None;
    let mut maximum = None;
    let mut rank = 0u8;
    let mut seen = HashSet::new();
    for child in &node.children {
        if child.namespace != CX {
            return invalid("foreign direct child in  valueColorPositions");
        }
        let current = match child.name.as_str() {
            "min" => 0,
            "mid" => 1,
            "max" => 2,
            _ => return invalid("invalid direct child in  valueColorPositions"),
        };
        if current < rank || !seen.insert(child.name.as_str()) {
            return invalid(" valueColorPositions children are out of order or duplicated");
        }
        rank = current;
        let value = parse_color_position(child, child.name != "mid")?;
        match child.name.as_str() {
            "min" => minimum = Some(value),
            "mid" => middle = Some(value),
            "max" => maximum = Some(value),
            _ => unreachable!(),
        }
    }
    Ok(ValueColorPositions {
        count,
        minimum,
        middle,
        maximum,
    })
}

pub(super) fn parse_color_position(node: &MiniNode, allow_extreme: bool) -> Result<ColorPosition> {
    reject_unknown(&node.attributes, &[], "color position")?;
    if !node.text.trim().is_empty() || node.children.len() != 1 {
        return invalid(" color position requires exactly one choice");
    }
    let child = &node.children[0];
    if child.namespace != CX {
        return invalid("foreign  color position choice");
    }
    match child.name.as_str() {
        "extreme" if allow_extreme => {
            require_empty_element(child, "extreme color position")?;
            Ok(ColorPosition::Extreme)
        },
        "number" => {
            let value = parse_position_value(child, "number color position")?;
            if !valid_xml_double(&value) {
                return invalid("invalid  number color position");
            }
            Ok(ColorPosition::Number(value))
        },
        "percent" => {
            let value = parse_position_value(child, "percent color position")?;
            let number = value
                .parse::<f64>()
                .map_err(|_| invalid_error("invalid  percent color position"))?;
            if !number.is_finite() || !(0.0..=100.0).contains(&number) {
                return invalid("invalid  percent color position");
            }
            Ok(ColorPosition::Percent(value))
        },
        _ => invalid("invalid  color position choice"),
    }
}

pub(super) fn parse_position_value(node: &MiniNode, label: &str) -> Result<String> {
    reject_unknown(&node.attributes, &[("", "val")], label)?;
    require_empty_content(node, label)?;
    Ok(required(&node.attributes, "", "val")?.to_owned())
}

pub(super) fn parse_data_point(node: &MiniNode) -> Result<DataPoint> {
    reject_unknown(&node.attributes, &[("", "idx")], "dataPt")?;
    if !node.text.trim().is_empty() {
        return invalid("unexpected text in  dataPt");
    }
    let index = parse_u32(required(&node.attributes, "", "idx")?, "data point index")?;
    let mut shape_properties = None;
    let mut ext_seen = false;
    for child in &node.children {
        if child.namespace != CX {
            return invalid("foreign direct child in  dataPt");
        }
        match child.name.as_str() {
            "spPr" if shape_properties.is_none() && !ext_seen => {
                shape_properties = Some(parse_drawing_payload(child, "data point spPr")?)
            },
            "extLst" if !ext_seen => ext_seen = true,
            _ => {
                return invalid(" dataPt children are invalid, duplicated, or out of order");
            },
        }
    }
    Ok(DataPoint {
        index,
        shape_properties,
    })
}

pub(super) fn parse_data_labels(node: &MiniNode) -> Result<DataLabels> {
    reject_unknown(&node.attributes, &[("", "pos")], "dataLabels")?;
    if !node.text.trim().is_empty() {
        return invalid("unexpected text in  dataLabels");
    }
    let position = optional(&node.attributes, "", "pos")
        .map(parse_label_position)
        .transpose()?;
    let mut number_format = None;
    let mut shape_properties = None;
    let mut text_properties = None;
    let mut visibility = None;
    let mut separator = None;
    let mut labels = Vec::new();
    let mut hidden_indices = Vec::new();
    let mut rank = 0u8;
    let mut singleton_seen = HashSet::new();
    let mut label_indices = HashSet::new();
    let mut hidden_set = HashSet::new();
    for child in &node.children {
        if child.namespace != CX {
            return invalid("foreign direct child in  dataLabels");
        }
        let current = match child.name.as_str() {
            "numFmt" => 0,
            "spPr" => 1,
            "txPr" => 2,
            "visibility" => 3,
            "separator" => 4,
            "dataLabel" => 5,
            "dataLabelHidden" => 6,
            "extLst" => 7,
            _ => return invalid("invalid direct child in  dataLabels"),
        };
        if current < rank {
            return invalid(" dataLabels children are out of order");
        }
        rank = current;
        if !matches!(child.name.as_str(), "dataLabel" | "dataLabelHidden")
            && !singleton_seen.insert(child.name.as_str())
        {
            return invalid("duplicate  dataLabels child");
        }
        match child.name.as_str() {
            "numFmt" => number_format = Some(parse_number_format(child)?),
            "spPr" => shape_properties = Some(parse_drawing_payload(child, "dataLabels spPr")?),
            "txPr" => text_properties = Some(parse_drawing_payload(child, "dataLabels txPr")?),
            "visibility" => visibility = Some(parse_label_visibility(child)?),
            "separator" => separator = Some(parse_separator(child)?),
            "dataLabel" => {
                if labels.len() >= MAX_DATA_LABELS {
                    return limit(" data label count");
                }
                let label = parse_data_label(child)?;
                if !label_indices.insert(label.index) || hidden_set.contains(&label.index) {
                    return invalid("duplicate or conflicting  data label index");
                }
                labels.push(label);
            },
            "dataLabelHidden" => {
                if hidden_indices.len() >= MAX_DATA_LABELS {
                    return limit(" hidden data label count");
                }
                let index = parse_hidden_label(child)?;
                if !hidden_set.insert(index) || label_indices.contains(&index) {
                    return invalid("duplicate or conflicting  data label index");
                }
                hidden_indices.push(index);
            },
            "extLst" => {},
            _ => unreachable!(),
        }
    }
    Ok(DataLabels {
        position,
        number_format,
        shape_properties,
        text_properties,
        visibility,
        separator,
        labels,
        hidden_indices,
    })
}

pub(super) fn parse_data_label(node: &MiniNode) -> Result<DataLabel> {
    reject_unknown(&node.attributes, &[("", "idx"), ("", "pos")], "dataLabel")?;
    if !node.text.trim().is_empty() {
        return invalid("unexpected text in  dataLabel");
    }
    let index = parse_u32(required(&node.attributes, "", "idx")?, "data label index")?;
    let position = optional(&node.attributes, "", "pos")
        .map(parse_label_position)
        .transpose()?;
    let mut number_format = None;
    let mut shape_properties = None;
    let mut text_properties = None;
    let mut visibility = None;
    let mut separator = None;
    let mut rank = 0u8;
    let mut seen = HashSet::new();
    for child in &node.children {
        if child.namespace != CX {
            return invalid("foreign direct child in  dataLabel");
        }
        let current = match child.name.as_str() {
            "numFmt" => 0,
            "spPr" => 1,
            "txPr" => 2,
            "visibility" => 3,
            "separator" => 4,
            "extLst" => 5,
            _ => return invalid("invalid direct child in  dataLabel"),
        };
        if current < rank || !seen.insert(child.name.as_str()) {
            return invalid(" dataLabel children are out of order or duplicated");
        }
        rank = current;
        match child.name.as_str() {
            "numFmt" => number_format = Some(parse_number_format(child)?),
            "spPr" => shape_properties = Some(parse_drawing_payload(child, "dataLabel spPr")?),
            "txPr" => text_properties = Some(parse_drawing_payload(child, "dataLabel txPr")?),
            "visibility" => visibility = Some(parse_label_visibility(child)?),
            "separator" => separator = Some(parse_separator(child)?),
            "extLst" => {},
            _ => unreachable!(),
        }
    }
    Ok(DataLabel {
        index,
        position,
        number_format,
        shape_properties,
        text_properties,
        visibility,
        separator,
    })
}

pub(super) fn parse_hidden_label(node: &MiniNode) -> Result<u32> {
    reject_unknown(&node.attributes, &[("", "idx")], "dataLabelHidden")?;
    require_empty_content(node, "dataLabelHidden")?;
    parse_u32(
        required(&node.attributes, "", "idx")?,
        "hidden data label index",
    )
}

pub(super) fn parse_number_format(node: &MiniNode) -> Result<NumberFormat> {
    reject_unknown(
        &node.attributes,
        &[("", "formatCode"), ("", "sourceLinked")],
        "numFmt",
    )?;
    require_empty_content(node, "numFmt")?;
    let format_code = bounded_required(node, "formatCode", 255)?;
    let source_linked = optional(&node.attributes, "", "sourceLinked")
        .map(parse_bool)
        .transpose()?;
    Ok(NumberFormat {
        format_code,
        source_linked,
    })
}

pub(super) fn parse_label_visibility(node: &MiniNode) -> Result<DataLabelVisibility> {
    reject_unknown(
        &node.attributes,
        &[("", "seriesName"), ("", "categoryName"), ("", "value")],
        "data label visibility",
    )?;
    require_empty_content(node, "data label visibility")?;
    Ok(DataLabelVisibility {
        series_name: optional(&node.attributes, "", "seriesName")
            .map(parse_bool)
            .transpose()?,
        category_name: optional(&node.attributes, "", "categoryName")
            .map(parse_bool)
            .transpose()?,
        value: optional(&node.attributes, "", "value")
            .map(parse_bool)
            .transpose()?,
    })
}

pub(super) fn parse_separator(node: &MiniNode) -> Result<String> {
    reject_unknown(&node.attributes, &[], "data label separator")?;
    if !node.children.is_empty() {
        return invalid(" data label separator must have simple content");
    }
    if node.text.len() > MAX_LABEL_TEXT_BYTES {
        return limit(" data label separator bytes");
    }
    Ok(node.text.clone())
}

pub(super) fn parse_label_position(value: &str) -> Result<DataLabelPosition> {
    match value {
        "bestFit" => Ok(DataLabelPosition::BestFit),
        "b" => Ok(DataLabelPosition::Bottom),
        "ctr" => Ok(DataLabelPosition::Center),
        "inBase" => Ok(DataLabelPosition::InsideBase),
        "inEnd" => Ok(DataLabelPosition::InsideEnd),
        "l" => Ok(DataLabelPosition::Left),
        "outEnd" => Ok(DataLabelPosition::OutsideEnd),
        "r" => Ok(DataLabelPosition::Right),
        "t" => Ok(DataLabelPosition::Top),
        _ => invalid("invalid  data label position"),
    }
}

pub(super) fn parse_series(node: &MiniNode) -> Result<SeriesDataReference> {
    reject_unknown(
        &node.attributes,
        &[
            ("", "layoutId"),
            ("", "hidden"),
            ("", "ownerIdx"),
            ("", "uniqueId"),
            ("", "formatIdx"),
        ],
        "series",
    )?;
    let layout = match required(&node.attributes, "", "layoutId")? {
        "boxWhisker" => SeriesLayout::BoxWhisker,
        "clusteredColumn" => SeriesLayout::ClusteredColumn,
        "funnel" => SeriesLayout::Funnel,
        "paretoLine" => SeriesLayout::ParetoLine,
        "regionMap" => SeriesLayout::RegionMap,
        "sunburst" => SeriesLayout::Sunburst,
        "treemap" => SeriesLayout::Treemap,
        "waterfall" => SeriesLayout::Waterfall,
        _ => return invalid("invalid  series layoutId"),
    };
    let mut rank = 0u8;
    let mut text = None;
    let mut shape_properties = None;
    let mut value_colors = None;
    let mut value_color_positions = None;
    let mut data_points = Vec::new();
    let mut data_point_indices = HashSet::new();
    let mut data_labels = None;
    let mut data_id = None;
    let mut layout_properties = None;
    let mut axis_ids = Vec::new();
    let mut singleton_seen = HashSet::new();
    for child in &node.children {
        let current = series_child_rank(child)
            .ok_or_else(|| invalid_error("invalid direct  series child"))?;
        if current < rank {
            return invalid(" series children are out of order");
        }
        rank = current;
        if !matches!(child.name.as_str(), "dataPt" | "axisId")
            && !singleton_seen.insert(child.name.as_str())
        {
            return invalid("duplicate  series child");
        }
        if child.namespace == CX && child.name == "tx" {
            text = Some(parse_shared_text(child, "series tx")?);
        } else if child.namespace == CX && child.name == "spPr" {
            shape_properties = Some(parse_drawing_payload(child, "series spPr")?);
        } else if child.namespace == CX && child.name == "valueColors" {
            value_colors = Some(parse_value_colors(child)?);
        } else if child.namespace == CX && child.name == "valueColorPositions" {
            value_color_positions = Some(parse_value_color_positions(child)?);
        } else if child.namespace == CX && child.name == "dataPt" {
            if data_points.len() >= MAX_SERIES_POINTS {
                return limit(" series data point count");
            }
            let point = parse_data_point(child)?;
            if !data_point_indices.insert(point.index) {
                return invalid("duplicate  series data point index");
            }
            data_points.push(point);
        } else if child.namespace == CX && child.name == "dataLabels" {
            data_labels = Some(parse_data_labels(child)?);
        } else if child.namespace == CX && child.name == "dataId" {
            if data_id.is_some() || !child.children.is_empty() || !child.text.trim().is_empty() {
                return invalid(" series dataId must be a unique leaf");
            }
            reject_unknown(&child.attributes, &[("", "val")], "series dataId")?;
            data_id = Some(parse_u32(
                required(&child.attributes, "", "val")?,
                "series dataId",
            )?);
        } else if child.namespace == CX && child.name == "layoutPr" {
            if layout_properties.is_some() {
                return invalid("duplicate  series layoutPr");
            }
            layout_properties = Some(parse_layout_properties(child)?);
        } else if child.namespace == CX && child.name == "axisId" {
            if axis_ids.len() >= MAX_AXIS_REFS_PER_SERIES {
                return limit(" series axis reference count");
            }
            reject_unknown(&child.attributes, &[], "series axisId")?;
            if !child.children.is_empty() {
                return invalid(" series axisId must have simple content");
            }
            axis_ids.push(parse_u32(child.text.trim(), "series axisId")?);
        }
    }
    let unique_id = bounded_optional(node, "uniqueId", 1024)?;
    Ok(SeriesDataReference {
        layout,
        text,
        shape_properties,
        value_colors,
        value_color_positions,
        data_points,
        data_labels,
        data_id,
        hidden: optional(&node.attributes, "", "hidden")
            .map(parse_bool)
            .transpose()?
            .unwrap_or(false),
        owner_index: optional(&node.attributes, "", "ownerIdx")
            .map(|value| parse_u32(value, "series ownerIdx"))
            .transpose()?,
        unique_id,
        format_index: optional(&node.attributes, "", "formatIdx")
            .map(|value| parse_u32(value, "series formatIdx"))
            .transpose()?,
        layout_properties,
        axis_ids,
    })
}

pub(super) fn series_child_rank(node: &MiniNode) -> Option<u8> {
    if node.namespace == CX {
        match node.name.as_str() {
            "tx" => Some(0),
            "spPr" => Some(1),
            "valueColors" => Some(2),
            "valueColorPositions" => Some(3),
            "dataPt" => Some(4),
            "dataLabels" => Some(5),
            "dataId" => Some(6),
            "layoutPr" => Some(7),
            "axisId" => Some(8),
            "extLst" => Some(9),
            _ => None,
        }
    } else {
        None
    }
}
