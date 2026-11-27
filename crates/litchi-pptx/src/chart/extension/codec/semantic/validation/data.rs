//! Chart data and dimensions validation concerns for the `ChartEx` graph.

use super::{
    CX, DataSet, Dimension, Formula, FormulaDirection, MAX_FORMULA_BYTES, MAX_LEVELS_PER_DIMENSION,
    MAX_POINTS_PER_LEVEL, MiniNode, NumericDimensionType, NumericLevel, NumericPoint, Result,
    SeriesDataReference, StringDimensionType, StringLevel, StringPoint, bounded_optional, invalid,
    limit, optional, parse_u32, reject_unknown, required, valid_xml_double,
};
use std::collections::HashSet;

pub(super) fn validate_series_point_references(
    series: &SeriesDataReference,
    data: &DataSet,
) -> Result<()> {
    let bound = data
        .dimensions
        .iter()
        .flat_map(|dimension| match dimension {
            Dimension::String { levels, .. } => levels
                .iter()
                .map(|level| level.point_count)
                .collect::<Vec<_>>(),
            Dimension::Numeric { levels, .. } => levels
                .iter()
                .map(|level| level.point_count)
                .collect::<Vec<_>>(),
        })
        .max();
    let Some(bound) = bound else {
        return Ok(());
    };
    if series.data_points.iter().any(|point| point.index >= bound) {
        return invalid(" dataPt index does not resolve to cached series data");
    }
    if let Some(labels) = &series.data_labels
        && (labels.labels.iter().any(|label| label.index >= bound)
            || labels.hidden_indices.iter().any(|index| *index >= bound))
    {
        return invalid(" data label index does not resolve to cached series data");
    }
    Ok(())
}

pub(super) fn parse_chart_data(node: &MiniNode) -> Result<Vec<DataSet>> {
    let mut rank = 0u8;
    let mut data_sets = Vec::new();
    let mut ids = HashSet::new();
    for child in &node.children {
        if child.namespace != CX {
            return invalid("foreign direct child in  chartData");
        }
        let current = match child.name.as_str() {
            "externalData" => 0,
            "data" => 1,
            "extLst" => 2,
            _ => return invalid("invalid direct child in  chartData"),
        };
        if current < rank || (current != 1 && current == rank && current != 0) {
            return invalid(" chartData children are out of order or duplicated");
        }
        rank = current;
        if child.name == "data" {
            let value = parse_data_set(child)?;
            if !ids.insert(value.id) {
                return invalid("duplicate  data ID");
            }
            data_sets.push(value);
        }
    }
    if data_sets.is_empty() {
        return invalid(" chartData requires data");
    }
    Ok(data_sets)
}

pub(super) fn parse_data_set(node: &MiniNode) -> Result<DataSet> {
    reject_unknown(&node.attributes, &[("", "id")], "data")?;
    let id = parse_u32(required(&node.attributes, "", "id")?, "data id")?;
    let mut dimensions = Vec::new();
    let mut ext_seen = false;
    for child in &node.children {
        if child.namespace != CX {
            return invalid("foreign direct child in  data");
        }
        match child.name.as_str() {
            "strDim" if !ext_seen => dimensions.push(parse_string_dimension(child)?),
            "numDim" if !ext_seen => dimensions.push(parse_numeric_dimension(child)?),
            "extLst" if !ext_seen => ext_seen = true,
            _ => return invalid(" data dimensions are invalid or out of order"),
        }
    }
    if dimensions.is_empty() {
        return invalid(" data requires at least one dimension");
    }
    let string_dimensions = dimensions
        .iter()
        .filter(|value| matches!(value, Dimension::String { .. }))
        .count();
    let numeric_dimensions = dimensions.len() - string_dimensions;
    Ok(DataSet {
        id,
        string_dimensions,
        numeric_dimensions,
        dimensions,
    })
}

pub(super) fn parse_string_dimension(node: &MiniNode) -> Result<Dimension> {
    reject_unknown(&node.attributes, &[("", "type")], "strDim")?;
    let kind = match required(&node.attributes, "", "type")? {
        "cat" => StringDimensionType::Category,
        "colorStr" => StringDimensionType::ColorString,
        "entityId" => StringDimensionType::EntityId,
        _ => return invalid("invalid  string dimension type"),
    };
    let (formula, name_formula, level_nodes) = dimension_children(node)?;
    let levels = level_nodes
        .into_iter()
        .map(parse_string_level)
        .collect::<Result<Vec<_>>>()?;
    Ok(Dimension::String {
        kind,
        formula,
        name_formula,
        levels,
    })
}

pub(super) fn parse_numeric_dimension(node: &MiniNode) -> Result<Dimension> {
    reject_unknown(&node.attributes, &[("", "type")], "numDim")?;
    let kind = match required(&node.attributes, "", "type")? {
        "val" => NumericDimensionType::Value,
        "x" => NumericDimensionType::X,
        "y" => NumericDimensionType::Y,
        "size" => NumericDimensionType::Size,
        "colorVal" => NumericDimensionType::ColorValue,
        _ => return invalid("invalid  numeric dimension type"),
    };
    let (formula, name_formula, level_nodes) = dimension_children(node)?;
    let levels = level_nodes
        .into_iter()
        .map(parse_numeric_level)
        .collect::<Result<Vec<_>>>()?;
    Ok(Dimension::Numeric {
        kind,
        formula,
        name_formula,
        levels,
    })
}

pub(super) fn dimension_children(
    node: &MiniNode,
) -> Result<(Option<Formula>, Option<Formula>, Vec<&MiniNode>)> {
    let mut formula = None;
    let mut name_formula = None;
    let mut levels = Vec::new();
    let mut rank = 0u8;
    for child in &node.children {
        if child.namespace != CX {
            return invalid("foreign direct child in  dimension");
        }
        let current = match child.name.as_str() {
            "f" => 0,
            "nf" => 1,
            "lvl" => 2,
            _ => return invalid("invalid  dimension child"),
        };
        if current < rank {
            return invalid(" dimension children are out of order");
        }
        rank = current;
        match child.name.as_str() {
            "f" if formula.is_none() && levels.is_empty() => formula = Some(parse_formula(child)?),
            "nf" if formula.is_some() && name_formula.is_none() && levels.is_empty() => {
                name_formula = Some(parse_formula(child)?);
            },
            "lvl" => levels.push(child),
            _ => return invalid(" dimension formula/literal choice is invalid"),
        }
    }
    if formula.is_none() && levels.is_empty() {
        return invalid(" dimension requires a formula or literal levels");
    }
    if levels.len() > MAX_LEVELS_PER_DIMENSION {
        return limit(" dimension levels");
    }
    Ok((formula, name_formula, levels))
}

pub(super) fn parse_formula(node: &MiniNode) -> Result<Formula> {
    if !node.children.is_empty() {
        return invalid(" formula must have simple content");
    }
    reject_unknown(&node.attributes, &[("", "dir")], "formula")?;
    if node.text.is_empty() || node.text.len() > MAX_FORMULA_BYTES {
        return invalid(" formula is empty or excessive");
    }
    let direction = match optional(&node.attributes, "", "dir").unwrap_or("col") {
        "col" => FormulaDirection::Column,
        "row" => FormulaDirection::Row,
        _ => return invalid("invalid  formula direction"),
    };
    Ok(Formula {
        expression: node.text.clone(),
        direction,
    })
}

pub(super) fn parse_string_level(node: &MiniNode) -> Result<StringLevel> {
    reject_unknown(
        &node.attributes,
        &[("", "ptCount"), ("", "name")],
        "string level",
    )?;
    let point_count = level_count(node)?;
    let mut indices = HashSet::new();
    let mut points = Vec::new();
    for point in &node.children {
        if point.namespace != CX || point.name != "pt" || !point.children.is_empty() {
            return invalid("invalid  string level point");
        }
        reject_unknown(&point.attributes, &[("", "idx")], "string point")?;
        let index = parse_u32(
            required(&point.attributes, "", "idx")?,
            "string point index",
        )?;
        if index >= point_count || !indices.insert(index) {
            return invalid(" string point index is duplicate or outside ptCount");
        }
        points.push(StringPoint {
            index,
            value: point.text.clone(),
        });
    }
    Ok(StringLevel {
        point_count,
        name: bounded_optional(node, "name", 1024)?,
        points,
    })
}

pub(super) fn parse_numeric_level(node: &MiniNode) -> Result<NumericLevel> {
    reject_unknown(
        &node.attributes,
        &[("", "ptCount"), ("", "formatCode"), ("", "name")],
        "numeric level",
    )?;
    let point_count = level_count(node)?;
    let mut indices = HashSet::new();
    let mut points = Vec::new();
    for point in &node.children {
        if point.namespace != CX || point.name != "pt" || !point.children.is_empty() {
            return invalid("invalid  numeric level point");
        }
        reject_unknown(&point.attributes, &[("", "idx")], "numeric point")?;
        let index = parse_u32(
            required(&point.attributes, "", "idx")?,
            "numeric point index",
        )?;
        let value = point.text.trim();
        if index >= point_count || !indices.insert(index) || !valid_xml_double(value) {
            return invalid("invalid  numeric point");
        }
        points.push(NumericPoint {
            index,
            value: value.to_owned(),
        });
    }
    Ok(NumericLevel {
        point_count,
        name: bounded_optional(node, "name", 1024)?,
        format_code: bounded_optional(node, "formatCode", 255)?,
        points,
    })
}

pub(super) fn level_count(node: &MiniNode) -> Result<u32> {
    let value = parse_u32(required(&node.attributes, "", "ptCount")?, "level ptCount")?;
    if value > MAX_POINTS_PER_LEVEL || node.children.len() > value as usize {
        return limit(" level point count");
    }
    Ok(value)
}
