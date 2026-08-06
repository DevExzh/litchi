//! ChartEx XML graph parsing and semantic validation.

mod axes;
mod chart;
mod data;
mod drawing;
mod geography;
mod primitives;

use super::super::super::model::*;
use super::super::limits::*;
use super::super::xml::*;
use super::model::ParsedDataGraph;
use crate::Result;
use std::collections::HashSet;

use self::{axes::*, chart::*, data::*, drawing::*, geography::*, primitives::*};

pub(in crate::chart::extension::codec) fn parse_data_graph(
    xml: &[u8],
    version: &str,
    features: &[String],
) -> Result<ParsedDataGraph> {
    let root = parse_mini_tree(xml)?;
    let chart_space_formatting = parse_chart_space_formatting(&root)?;
    let chart_data =
        one_child(&root, CX, "chartData")?.ok_or_else(|| invalid_error("missing  chartData"))?;
    let chart = one_child(&root, CX, "chart")?.ok_or_else(|| invalid_error("missing  chart"))?;
    let chart_info = parse_chart(chart, offset_feature_allowed(version, features))?;
    let data_sets = parse_chart_data(chart_data)?;
    let plot_area = one_child(chart, CX, "plotArea")?
        .ok_or_else(|| invalid_error(" chart is missing plotArea"))?;
    reject_unknown(&plot_area.attributes, &[], "plotArea")?;
    if !plot_area.text.trim().is_empty() {
        return invalid("unexpected text in  plotArea");
    }
    let mut plot_rank = 0u8;
    let mut region = None;
    let mut axes = Vec::new();
    let mut plot_shape_properties = None;
    let mut plot_has_extension_list = false;
    let mut singleton_seen = HashSet::new();
    for child in &plot_area.children {
        if child.namespace != CX {
            return invalid("foreign direct child in  plotArea");
        }
        let current = match child.name.as_str() {
            "plotAreaRegion" => 0,
            "axis" => 1,
            "spPr" => 2,
            "extLst" => 3,
            _ => return invalid("invalid direct child in  plotArea"),
        };
        if current < plot_rank {
            return invalid(" plotArea children are out of order");
        }
        plot_rank = current;
        match child.name.as_str() {
            "plotAreaRegion" if region.is_none() => region = Some(child),
            "axis" => {
                if axes.len() >= MAX_AXES {
                    return limit(" axis count");
                }
                axes.push(parse_axis(
                    child,
                    offset_feature_allowed(version, features),
                )?);
            },
            "spPr" if singleton_seen.insert(child.name.as_str()) => {
                plot_shape_properties = Some(parse_drawing_payload(child, "plotArea spPr")?);
            },
            "extLst" if singleton_seen.insert(child.name.as_str()) => {
                plot_has_extension_list = true
            },
            _ => return invalid("duplicate  plotArea child"),
        }
    }
    let region = region.ok_or_else(|| invalid_error(" plotArea is missing plotAreaRegion"))?;
    reject_unknown(&region.attributes, &[], "plotAreaRegion")?;
    if !region.text.trim().is_empty() {
        return invalid("unexpected text in  plotAreaRegion");
    }
    let mut series = Vec::new();
    let mut region_rank = 0u8;
    let mut plot_surface = None;
    let mut region_ext_seen = false;
    for child in &region.children {
        if child.namespace != CX {
            return invalid("foreign direct child in  plotAreaRegion");
        }
        let current = match child.name.as_str() {
            "plotSurface" => 0,
            "series" => 1,
            "extLst" => 2,
            _ => return invalid("invalid direct child in  plotAreaRegion"),
        };
        if current < region_rank {
            return invalid(" plotAreaRegion children are out of order");
        }
        region_rank = current;
        match child.name.as_str() {
            "plotSurface" if plot_surface.is_none() => {
                plot_surface = Some(parse_plot_surface(child)?);
            },
            "series" => {
                if series.len() >= MAX_SERIES {
                    return limit(" series count");
                }
                series.push(parse_series(child)?);
            },
            "extLst" if !region_ext_seen => region_ext_seen = true,
            _ => return invalid("duplicate  plotAreaRegion child"),
        }
    }
    let ids = data_sets
        .iter()
        .map(|value| value.id)
        .collect::<HashSet<_>>();
    let mut axis_ids = HashSet::new();
    for axis in &axes {
        if !axis_ids.insert(axis.id) {
            return invalid("duplicate  axis ID");
        }
    }
    let mut unique_ids = HashSet::new();
    for (index, value) in series.iter().enumerate() {
        if value.data_id.is_some_and(|id| !ids.contains(&id)) {
            return invalid(" series dataId does not resolve to chartData");
        }
        if let Some(data_id) = value.data_id {
            let data = data_sets
                .iter()
                .find(|data| data.id == data_id)
                .ok_or_else(|| invalid_error(" series dataId does not resolve to chartData"))?;
            validate_series_point_references(value, data)?;
        }
        if value
            .owner_index
            .is_some_and(|owner| owner as usize >= series.len() || owner as usize == index)
        {
            return invalid(" series ownerIdx is out of range or self-referential");
        }
        if let Some(id) = &value.unique_id
            && !unique_ids.insert(id.as_str())
        {
            return invalid("duplicate  series uniqueId");
        }
        let mut references = HashSet::new();
        for axis_id in &value.axis_ids {
            if !references.insert(*axis_id) {
                return invalid("duplicate  series axisId");
            }
            if !axis_ids.contains(axis_id) {
                return invalid(" series axisId does not resolve to plotArea axis");
            }
        }
    }
    let has_plot_surface = plot_surface.is_some();
    let plot_area_info = PlotArea {
        region: PlotAreaRegion {
            plot_surface,
            has_extension_list: region_ext_seen,
        },
        shape_properties: plot_shape_properties,
        has_extension_list: plot_has_extension_list,
    };
    Ok((
        data_sets,
        series,
        axes,
        has_plot_surface,
        chart_info,
        plot_area_info,
        chart_space_formatting,
    ))
}
