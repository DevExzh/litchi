//! Plot-area axes and layout validation concerns for the `ChartEx` graph.

use super::{
    Axis, AxisScaling, AxisTitle, AxisUnit, AxisUnits, AxisUnitsLabel, Binning, BinningChoice, CX,
    ClosedSide, ElementVisibility, Gridlines, LayoutProperties, MiniNode, ParentLabelLayout,
    PlotSurface, RegionLabelLayout, Result, TickLabels, TickMarkType, TickMarks, invalid,
    invalid_error, optional, parse_bool, parse_double_or_auto, parse_drawing_payload,
    parse_geography, parse_nonnegative_or_auto, parse_number_format, parse_offset,
    parse_positive_or_auto, parse_shared_text, parse_statistics, parse_subtotals, parse_u32,
    reject_unknown, require_empty_content, require_empty_element, required, valid_xml_double,
};
use std::collections::HashSet;

pub(super) fn parse_plot_surface(node: &MiniNode) -> Result<PlotSurface> {
    reject_unknown(&node.attributes, &[], "plotSurface")?;
    if !node.text.trim().is_empty() {
        return invalid("unexpected text in  plotSurface");
    }
    let mut rank = 0u8;
    let mut seen = HashSet::new();
    let mut shape_properties = None;
    let mut has_extension_list = false;
    for child in &node.children {
        if child.namespace != CX {
            return invalid("foreign direct child in  plotSurface");
        }
        let current = match child.name.as_str() {
            "spPr" => 0,
            "extLst" => 1,
            _ => return invalid("invalid direct child in  plotSurface"),
        };
        if current < rank || !seen.insert(child.name.as_str()) {
            return invalid(" plotSurface children are out of order or duplicated");
        }
        rank = current;
        match child.name.as_str() {
            "spPr" => shape_properties = Some(parse_drawing_payload(child, "plotSurface spPr")?),
            "extLst" => has_extension_list = true,
            _ => unreachable!(),
        }
    }
    Ok(PlotSurface {
        shape_properties,
        has_extension_list,
    })
}

pub(super) fn parse_axis(node: &MiniNode, offset_allowed: bool) -> Result<Axis> {
    reject_unknown(&node.attributes, &[("", "id"), ("", "hidden")], "axis")?;
    if !node.text.trim().is_empty() {
        return invalid("unexpected text in  axis");
    }
    let id = parse_u32(required(&node.attributes, "", "id")?, "axis id")?;
    let hidden = optional(&node.attributes, "", "hidden")
        .map(parse_bool)
        .transpose()?
        .unwrap_or(false);
    let mut scaling = None;
    let mut title = None;
    let mut units = None;
    let mut major_gridlines = None;
    let mut minor_gridlines = None;
    let mut major_tick_marks = None;
    let mut minor_tick_marks = None;
    let mut tick_labels = None;
    let mut number_format = None;
    let mut shape_properties = None;
    let mut text_properties = None;
    let mut has_extension_list = false;
    let mut rank = 0u8;
    let mut seen = HashSet::new();
    for child in &node.children {
        if child.namespace != CX {
            return invalid("foreign direct child in  axis");
        }
        let current = match child.name.as_str() {
            "catScaling" | "valScaling" => 0,
            "title" => 1,
            "units" => 2,
            "majorGridlines" => 3,
            "minorGridlines" => 4,
            "majorTickMarks" => 5,
            "minorTickMarks" => 6,
            "tickLabels" => 7,
            "numFmt" => 8,
            "spPr" => 9,
            "txPr" => 10,
            "extLst" => 11,
            _ => return invalid("invalid direct child in  axis"),
        };
        if current < rank {
            return invalid(" axis children are out of order");
        }
        rank = current;
        if current == 0 {
            if scaling.is_some() {
                return invalid(" axis requires exactly one scaling choice");
            }
            scaling = Some(if child.name == "catScaling" {
                parse_category_scaling(child)?
            } else {
                parse_value_scaling(child)?
            });
        } else if !seen.insert(child.name.as_str()) {
            return invalid("duplicate  axis child");
        } else {
            match child.name.as_str() {
                "title" => title = Some(parse_axis_title(child, offset_allowed)?),
                "units" => units = Some(parse_axis_units(child)?),
                "majorGridlines" => {
                    major_gridlines = Some(parse_gridlines(child, "majorGridlines")?);
                },
                "minorGridlines" => {
                    minor_gridlines = Some(parse_gridlines(child, "minorGridlines")?);
                },
                "majorTickMarks" => {
                    major_tick_marks = Some(parse_tick_marks(child, "majorTickMarks")?);
                },
                "minorTickMarks" => {
                    minor_tick_marks = Some(parse_tick_marks(child, "minorTickMarks")?);
                },
                "tickLabels" => tick_labels = Some(parse_tick_labels(child)?),
                "numFmt" => number_format = Some(parse_number_format(child)?),
                "spPr" => shape_properties = Some(parse_drawing_payload(child, "axis spPr")?),
                "txPr" => text_properties = Some(parse_drawing_payload(child, "axis txPr")?),
                "extLst" => has_extension_list = true,
                _ => unreachable!(),
            }
        }
    }
    Ok(Axis {
        id,
        hidden,
        scaling: scaling.ok_or_else(|| invalid_error(" axis is missing scaling"))?,
        title,
        units,
        major_gridlines,
        minor_gridlines,
        major_tick_marks,
        minor_tick_marks,
        tick_labels,
        number_format,
        shape_properties,
        text_properties,
        has_extension_list,
    })
}

pub(super) fn parse_axis_title(node: &MiniNode, offset_allowed: bool) -> Result<AxisTitle> {
    reject_unknown(&node.attributes, &[], "axis title")?;
    if !node.text.trim().is_empty() {
        return invalid("unexpected text in  axis title");
    }
    let mut text = None;
    let mut shape_properties = None;
    let mut text_properties = None;
    let mut offset = None;
    let mut has_extension_list = false;
    let mut rank = 0u8;
    let mut seen = HashSet::new();
    for child in &node.children {
        if child.namespace != CX {
            return invalid("foreign direct child in  axis title");
        }
        let current = match child.name.as_str() {
            "tx" => 0,
            "spPr" => 1,
            "txPr" => 2,
            "offset" => 3,
            "extLst" => 4,
            _ => return invalid("invalid direct child in  axis title"),
        };
        if current < rank || !seen.insert(child.name.as_str()) {
            return invalid(" axis title children are out of order or duplicated");
        }
        rank = current;
        match child.name.as_str() {
            "tx" => text = Some(parse_shared_text(child, "axis title tx")?),
            "spPr" => shape_properties = Some(parse_drawing_payload(child, "axis title spPr")?),
            "txPr" => text_properties = Some(parse_drawing_payload(child, "axis title txPr")?),
            "offset" => {
                if !offset_allowed {
                    return invalid(" axis title offset requires version 1.0 or feature mp");
                }
                offset = Some(parse_offset(child)?);
            },
            "extLst" => has_extension_list = true,
            _ => unreachable!(),
        }
    }
    Ok(AxisTitle {
        text,
        shape_properties,
        text_properties,
        offset,
        has_extension_list,
    })
}

pub(super) fn parse_axis_units(node: &MiniNode) -> Result<AxisUnits> {
    reject_unknown(&node.attributes, &[("", "unit")], "axis units")?;
    if !node.text.trim().is_empty() {
        return invalid("unexpected text in  axis units");
    }
    let unit = optional(&node.attributes, "", "unit")
        .map(parse_axis_unit)
        .transpose()?;
    let mut label = None;
    let mut has_extension_list = false;
    let mut rank = 0u8;
    let mut seen = HashSet::new();
    for child in &node.children {
        if child.namespace != CX {
            return invalid("foreign direct child in  axis units");
        }
        let current = match child.name.as_str() {
            "unitsLabel" => 0,
            "extLst" => 1,
            _ => return invalid("invalid direct child in  axis units"),
        };
        if current < rank || !seen.insert(child.name.as_str()) {
            return invalid(" axis units children are out of order or duplicated");
        }
        rank = current;
        match child.name.as_str() {
            "unitsLabel" => label = Some(parse_axis_units_label(child)?),
            "extLst" => has_extension_list = true,
            _ => unreachable!(),
        }
    }
    Ok(AxisUnits {
        unit,
        label,
        has_extension_list,
    })
}

pub(super) fn parse_axis_unit(value: &str) -> Result<AxisUnit> {
    match value {
        "hundreds" => Ok(AxisUnit::Hundreds),
        "thousands" => Ok(AxisUnit::Thousands),
        "tenThousands" => Ok(AxisUnit::TenThousands),
        "hundredThousands" => Ok(AxisUnit::HundredThousands),
        "millions" => Ok(AxisUnit::Millions),
        "tenMillions" => Ok(AxisUnit::TenMillions),
        "hundredMillions" => Ok(AxisUnit::HundredMillions),
        "billions" => Ok(AxisUnit::Billions),
        "trillions" => Ok(AxisUnit::Trillions),
        "percentage" => Ok(AxisUnit::Percentage),
        _ => invalid("invalid  axis display unit"),
    }
}

pub(super) fn parse_axis_units_label(node: &MiniNode) -> Result<AxisUnitsLabel> {
    reject_unknown(&node.attributes, &[], "axis units label")?;
    if !node.text.trim().is_empty() {
        return invalid("unexpected text in  axis units label");
    }
    let mut text = None;
    let mut shape_properties = None;
    let mut text_properties = None;
    let mut has_extension_list = false;
    let mut rank = 0u8;
    let mut seen = HashSet::new();
    for child in &node.children {
        if child.namespace != CX {
            return invalid("foreign direct child in  axis units label");
        }
        let current = match child.name.as_str() {
            "tx" => 0,
            "spPr" => 1,
            "txPr" => 2,
            "extLst" => 3,
            _ => return invalid("invalid direct child in  axis units label"),
        };
        if current < rank || !seen.insert(child.name.as_str()) {
            return invalid(" axis units label children are out of order or duplicated");
        }
        rank = current;
        match child.name.as_str() {
            "tx" => text = Some(parse_shared_text(child, "axis units label tx")?),
            "spPr" => {
                shape_properties = Some(parse_drawing_payload(child, "axis units label spPr")?);
            },
            "txPr" => {
                text_properties = Some(parse_drawing_payload(child, "axis units label txPr")?);
            },
            "extLst" => has_extension_list = true,
            _ => unreachable!(),
        }
    }
    Ok(AxisUnitsLabel {
        text,
        shape_properties,
        text_properties,
        has_extension_list,
    })
}

pub(super) fn parse_gridlines(node: &MiniNode, label: &str) -> Result<Gridlines> {
    reject_unknown(&node.attributes, &[], label)?;
    if !node.text.trim().is_empty() {
        return invalid(format!("unexpected text in  {label}"));
    }
    let mut shape_properties = None;
    let mut has_extension_list = false;
    let mut rank = 0u8;
    let mut seen = HashSet::new();
    for child in &node.children {
        if child.namespace != CX {
            return invalid(format!("foreign direct child in  {label}"));
        }
        let current = match child.name.as_str() {
            "spPr" => 0,
            "extLst" => 1,
            _ => return invalid(format!("invalid direct child in  {label}")),
        };
        if current < rank || !seen.insert(child.name.as_str()) {
            return invalid(format!(" {label} children are out of order or duplicated"));
        }
        rank = current;
        match child.name.as_str() {
            "spPr" => shape_properties = Some(parse_drawing_payload(child, label)?),
            "extLst" => has_extension_list = true,
            _ => unreachable!(),
        }
    }
    Ok(Gridlines {
        shape_properties,
        has_extension_list,
    })
}

pub(super) fn parse_tick_marks(node: &MiniNode, label: &str) -> Result<TickMarks> {
    reject_unknown(&node.attributes, &[("", "type")], label)?;
    if !node.text.trim().is_empty() {
        return invalid(format!("unexpected text in  {label}"));
    }
    let kind = optional(&node.attributes, "", "type")
        .map(parse_tick_mark_type)
        .transpose()?;
    let has_extension_list = parse_extension_only(node, label)?;
    Ok(TickMarks {
        kind,
        has_extension_list,
    })
}

pub(super) fn parse_tick_mark_type(value: &str) -> Result<TickMarkType> {
    match value {
        "in" => Ok(TickMarkType::Inside),
        "out" => Ok(TickMarkType::Outside),
        "cross" => Ok(TickMarkType::Cross),
        "none" => Ok(TickMarkType::None),
        _ => invalid("invalid  tick mark type"),
    }
}

pub(super) fn parse_tick_labels(node: &MiniNode) -> Result<TickLabels> {
    reject_unknown(&node.attributes, &[], "tickLabels")?;
    if !node.text.trim().is_empty() {
        return invalid("unexpected text in  tickLabels");
    }
    Ok(TickLabels {
        has_extension_list: parse_extension_only(node, "tickLabels")?,
    })
}

pub(super) fn parse_extension_only(node: &MiniNode, label: &str) -> Result<bool> {
    let mut has_extension_list = false;
    for child in &node.children {
        if child.namespace != CX || child.name != "extLst" || has_extension_list {
            return invalid(format!("invalid or duplicate direct child in  {label}"));
        }
        has_extension_list = true;
    }
    Ok(has_extension_list)
}

pub(super) fn parse_category_scaling(node: &MiniNode) -> Result<AxisScaling> {
    reject_unknown(&node.attributes, &[("", "gapWidth")], "catScaling")?;
    require_empty_content(node, "catScaling")?;
    let gap_width = optional(&node.attributes, "", "gapWidth")
        .map(|value| parse_nonnegative_or_auto(value, "category gapWidth"))
        .transpose()?;
    Ok(AxisScaling::Category { gap_width })
}

pub(super) fn parse_value_scaling(node: &MiniNode) -> Result<AxisScaling> {
    reject_unknown(
        &node.attributes,
        &[
            ("", "max"),
            ("", "min"),
            ("", "majorUnit"),
            ("", "minorUnit"),
        ],
        "valScaling",
    )?;
    require_empty_content(node, "valScaling")?;
    let maximum = optional(&node.attributes, "", "max")
        .map(|value| parse_double_or_auto(value, "value axis maximum"))
        .transpose()?;
    let minimum = optional(&node.attributes, "", "min")
        .map(|value| parse_double_or_auto(value, "value axis minimum"))
        .transpose()?;
    let major_unit = optional(&node.attributes, "", "majorUnit")
        .map(|value| parse_positive_or_auto(value, "value axis majorUnit"))
        .transpose()?;
    let minor_unit = optional(&node.attributes, "", "minorUnit")
        .map(|value| parse_positive_or_auto(value, "value axis minorUnit"))
        .transpose()?;
    Ok(AxisScaling::Value {
        minimum,
        maximum,
        major_unit,
        minor_unit,
    })
}

pub(super) fn parse_layout_properties(node: &MiniNode) -> Result<LayoutProperties> {
    reject_unknown(&node.attributes, &[], "layoutPr")?;
    if !node.text.trim().is_empty() {
        return invalid("unexpected text in  layoutPr");
    }
    let mut result = LayoutProperties::default();
    let mut rank = 0u8;
    let mut seen = HashSet::new();
    let mut aggregation_choice = false;
    for child in &node.children {
        if child.namespace != CX {
            return invalid("foreign direct child in  layoutPr");
        }
        let current = match child.name.as_str() {
            "parentLabelLayout" => 0,
            "regionLabelLayout" => 1,
            "visibility" => 2,
            "aggregation" | "binning" => 3,
            "geography" => 4,
            "statistics" => 5,
            "subtotals" => 6,
            "extLst" => 7,
            _ => return invalid("invalid direct child in  layoutPr"),
        };
        if current < rank || !seen.insert(child.name.as_str()) {
            return invalid(" layoutPr children are out of order or duplicated");
        }
        rank = current;
        match child.name.as_str() {
            "parentLabelLayout" => result.parent_label = Some(parse_parent_label(child)?),
            "regionLabelLayout" => result.region_label = Some(parse_region_label(child)?),
            "visibility" => result.visibility = Some(parse_visibility(child)?),
            "aggregation" => {
                if aggregation_choice {
                    return invalid(" layoutPr aggregation and binning are mutually exclusive");
                }
                require_empty_element(child, "aggregation")?;
                aggregation_choice = true;
                result.aggregation = true;
            },
            "binning" => {
                if aggregation_choice {
                    return invalid(" layoutPr aggregation and binning are mutually exclusive");
                }
                aggregation_choice = true;
                result.binning = Some(parse_binning(child)?);
            },
            "geography" => result.geography = Some(parse_geography(child)?),
            "statistics" => result.quartile_method = parse_statistics(child)?,
            "subtotals" => result.subtotals = parse_subtotals(child)?,
            "extLst" => {},
            _ => unreachable!(),
        }
    }
    Ok(result)
}

pub(super) fn parse_parent_label(node: &MiniNode) -> Result<ParentLabelLayout> {
    reject_unknown(&node.attributes, &[("", "val")], "parentLabelLayout")?;
    require_empty_content(node, "parentLabelLayout")?;
    match required(&node.attributes, "", "val")? {
        "none" => Ok(ParentLabelLayout::None),
        "banner" => Ok(ParentLabelLayout::Banner),
        "overlapping" => Ok(ParentLabelLayout::Overlapping),
        _ => invalid("invalid  parentLabelLayout value"),
    }
}

pub(super) fn parse_region_label(node: &MiniNode) -> Result<RegionLabelLayout> {
    reject_unknown(&node.attributes, &[("", "val")], "regionLabelLayout")?;
    require_empty_content(node, "regionLabelLayout")?;
    match required(&node.attributes, "", "val")? {
        "none" => Ok(RegionLabelLayout::None),
        "bestFitOnly" => Ok(RegionLabelLayout::BestFitOnly),
        "showAll" => Ok(RegionLabelLayout::ShowAll),
        _ => invalid("invalid  regionLabelLayout value"),
    }
}

pub(super) fn parse_visibility(node: &MiniNode) -> Result<ElementVisibility> {
    let allowed = &[
        ("", "connectorLines"),
        ("", "meanLine"),
        ("", "meanMarker"),
        ("", "nonoutliers"),
        ("", "outliers"),
    ];
    reject_unknown(&node.attributes, allowed, "visibility")?;
    require_empty_content(node, "visibility")?;
    Ok(ElementVisibility {
        connector_lines: optional(&node.attributes, "", "connectorLines")
            .map(parse_bool)
            .transpose()?,
        mean_line: optional(&node.attributes, "", "meanLine")
            .map(parse_bool)
            .transpose()?,
        mean_marker: optional(&node.attributes, "", "meanMarker")
            .map(parse_bool)
            .transpose()?,
        nonoutliers: optional(&node.attributes, "", "nonoutliers")
            .map(parse_bool)
            .transpose()?,
        outliers: optional(&node.attributes, "", "outliers")
            .map(parse_bool)
            .transpose()?,
    })
}

pub(super) fn parse_binning(node: &MiniNode) -> Result<Binning> {
    reject_unknown(
        &node.attributes,
        &[("", "intervalClosed"), ("", "underflow"), ("", "overflow")],
        "binning",
    )?;
    if !node.text.trim().is_empty() {
        return invalid("unexpected text in  binning");
    }
    let interval_closed = optional(&node.attributes, "", "intervalClosed")
        .map(|value| match value {
            "l" => Ok(ClosedSide::Left),
            "r" => Ok(ClosedSide::Right),
            _ => invalid("invalid  binning intervalClosed"),
        })
        .transpose()?;
    let underflow = optional(&node.attributes, "", "underflow")
        .map(|value| parse_double_or_auto(value, "binning underflow"))
        .transpose()?;
    let overflow = optional(&node.attributes, "", "overflow")
        .map(|value| parse_double_or_auto(value, "binning overflow"))
        .transpose()?;
    let mut choice = None;
    for child in &node.children {
        if child.namespace != CX || !matches!(child.name.as_str(), "binSize" | "binCount") {
            return invalid("invalid direct child in  binning");
        }
        if choice.is_some() || !child.attributes.is_empty() || !child.children.is_empty() {
            return invalid(" binning permits at most one simple-content choice");
        }
        let value = child.text.trim();
        choice = Some(if child.name == "binSize" {
            if !valid_xml_double(value) {
                return invalid("invalid  binSize");
            }
            BinningChoice::Size(value.to_owned())
        } else {
            BinningChoice::Count(parse_u32(value, "binCount")?)
        });
    }
    Ok(Binning {
        choice,
        interval_closed,
        underflow,
        overflow,
    })
}
