//! Typed ChartEx graph conversion and semantic validation.

use super::super::super::style::{ColorDocument, Document as StyleDocument};
use super::super::model::*;
use super::limits::*;
use super::xml::*;
use crate::Result;
use std::collections::HashSet;

impl Document {
    pub fn info(&self) -> &Info {
        &self.info
    }

    pub fn external_data_target(&self) -> Option<&ExternalDataTarget> {
        self.external_data_target.as_ref()
    }

    pub fn fallback_image_part_name(&self) -> Option<&str> {
        self.fallback_image_part_name.as_deref()
    }

    pub fn chart_style(&self) -> Option<&StyleDocument> {
        self.chart_style.as_ref()
    }

    pub fn chart_color_style(&self) -> Option<&ColorDocument> {
        self.chart_color_style.as_ref()
    }

    /// Return the validated source XML unchanged.
    pub fn to_xml(&self) -> Vec<u8> {
        self.xml.clone()
    }
}

pub(super) type ParsedDataGraph = (
    Vec<DataSet>,
    Vec<SeriesDataReference>,
    Vec<Axis>,
    bool,
    Chart,
    PlotArea,
    ChartSpaceFormatting,
);

pub(super) fn parse_data_graph(
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

fn parse_chart_space_formatting(root: &MiniNode) -> Result<ChartSpaceFormatting> {
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

fn parse_format_overrides(node: &MiniNode) -> Result<Vec<FormatOverride>> {
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

fn parse_format_override(node: &MiniNode) -> Result<FormatOverride> {
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

fn parse_print_settings(node: &MiniNode) -> Result<PrintSettings> {
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

fn parse_header_footer(node: &MiniNode) -> Result<HeaderFooter> {
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

fn parse_print_text(node: &MiniNode) -> Result<String> {
    reject_unknown(&node.attributes, &[], "header/footer text")?;
    if !node.children.is_empty() {
        return invalid(" header/footer text must have simple content");
    }
    if node.text.len() > MAX_PRINT_TEXT_BYTES {
        return limit(" header/footer text bytes");
    }
    Ok(node.text.clone())
}

fn parse_page_margins(node: &MiniNode) -> Result<PageMargins> {
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

fn parse_page_setup(node: &MiniNode) -> Result<PageSetup> {
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

fn parse_page_orientation(value: &str) -> Result<PageOrientation> {
    match value {
        "default" => Ok(PageOrientation::Default),
        "portrait" => Ok(PageOrientation::Portrait),
        "landscape" => Ok(PageOrientation::Landscape),
        _ => invalid("invalid  page orientation"),
    }
}

fn parse_chart(node: &MiniNode, offset_allowed: bool) -> Result<Chart> {
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

fn parse_chart_title(node: &MiniNode, offset_allowed: bool) -> Result<ChartTitle> {
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

fn parse_legend(node: &MiniNode, offset_allowed: bool) -> Result<Legend> {
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

fn parse_offset(node: &MiniNode) -> Result<Offset> {
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

fn parse_side_position(value: &str) -> Result<SidePosition> {
    match value {
        "l" => Ok(SidePosition::Left),
        "r" => Ok(SidePosition::Right),
        "t" => Ok(SidePosition::Top),
        "b" => Ok(SidePosition::Bottom),
        _ => invalid("invalid  side position"),
    }
}

fn parse_position_alignment(value: &str) -> Result<PositionAlignment> {
    match value {
        "min" => Ok(PositionAlignment::Minimum),
        "ctr" => Ok(PositionAlignment::Center),
        "max" => Ok(PositionAlignment::Maximum),
        _ => invalid("invalid  position alignment"),
    }
}

fn offset_feature_allowed(version: &str, features: &[String]) -> bool {
    version
        .split('.')
        .next()
        .and_then(|value| value.parse::<u32>().ok())
        .is_some_and(|major| major >= 1)
        || features.iter().any(|feature| feature == "mp")
}

fn validate_series_point_references(series: &SeriesDataReference, data: &DataSet) -> Result<()> {
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

fn parse_chart_data(node: &MiniNode) -> Result<Vec<DataSet>> {
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

fn parse_data_set(node: &MiniNode) -> Result<DataSet> {
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

fn parse_string_dimension(node: &MiniNode) -> Result<Dimension> {
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

fn parse_numeric_dimension(node: &MiniNode) -> Result<Dimension> {
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

fn dimension_children(
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
                name_formula = Some(parse_formula(child)?)
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

fn parse_formula(node: &MiniNode) -> Result<Formula> {
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

fn parse_string_level(node: &MiniNode) -> Result<StringLevel> {
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

fn parse_numeric_level(node: &MiniNode) -> Result<NumericLevel> {
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

fn level_count(node: &MiniNode) -> Result<u32> {
    let value = parse_u32(required(&node.attributes, "", "ptCount")?, "level ptCount")?;
    if value > MAX_POINTS_PER_LEVEL || node.children.len() > value as usize {
        return limit(" level point count");
    }
    Ok(value)
}

fn parse_plot_surface(node: &MiniNode) -> Result<PlotSurface> {
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

fn parse_axis(node: &MiniNode, offset_allowed: bool) -> Result<Axis> {
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
                    major_gridlines = Some(parse_gridlines(child, "majorGridlines")?)
                },
                "minorGridlines" => {
                    minor_gridlines = Some(parse_gridlines(child, "minorGridlines")?)
                },
                "majorTickMarks" => {
                    major_tick_marks = Some(parse_tick_marks(child, "majorTickMarks")?)
                },
                "minorTickMarks" => {
                    minor_tick_marks = Some(parse_tick_marks(child, "minorTickMarks")?)
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

fn parse_axis_title(node: &MiniNode, offset_allowed: bool) -> Result<AxisTitle> {
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

fn parse_axis_units(node: &MiniNode) -> Result<AxisUnits> {
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

fn parse_axis_unit(value: &str) -> Result<AxisUnit> {
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

fn parse_axis_units_label(node: &MiniNode) -> Result<AxisUnitsLabel> {
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
                shape_properties = Some(parse_drawing_payload(child, "axis units label spPr")?)
            },
            "txPr" => {
                text_properties = Some(parse_drawing_payload(child, "axis units label txPr")?)
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

fn parse_gridlines(node: &MiniNode, label: &str) -> Result<Gridlines> {
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

fn parse_tick_marks(node: &MiniNode, label: &str) -> Result<TickMarks> {
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

fn parse_tick_mark_type(value: &str) -> Result<TickMarkType> {
    match value {
        "in" => Ok(TickMarkType::Inside),
        "out" => Ok(TickMarkType::Outside),
        "cross" => Ok(TickMarkType::Cross),
        "none" => Ok(TickMarkType::None),
        _ => invalid("invalid  tick mark type"),
    }
}

fn parse_tick_labels(node: &MiniNode) -> Result<TickLabels> {
    reject_unknown(&node.attributes, &[], "tickLabels")?;
    if !node.text.trim().is_empty() {
        return invalid("unexpected text in  tickLabels");
    }
    Ok(TickLabels {
        has_extension_list: parse_extension_only(node, "tickLabels")?,
    })
}

fn parse_extension_only(node: &MiniNode, label: &str) -> Result<bool> {
    let mut has_extension_list = false;
    for child in &node.children {
        if child.namespace != CX || child.name != "extLst" || has_extension_list {
            return invalid(format!("invalid or duplicate direct child in  {label}"));
        }
        has_extension_list = true;
    }
    Ok(has_extension_list)
}

fn parse_category_scaling(node: &MiniNode) -> Result<AxisScaling> {
    reject_unknown(&node.attributes, &[("", "gapWidth")], "catScaling")?;
    require_empty_content(node, "catScaling")?;
    let gap_width = optional(&node.attributes, "", "gapWidth")
        .map(|value| parse_nonnegative_or_auto(value, "category gapWidth"))
        .transpose()?;
    Ok(AxisScaling::Category { gap_width })
}

fn parse_value_scaling(node: &MiniNode) -> Result<AxisScaling> {
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

fn parse_layout_properties(node: &MiniNode) -> Result<LayoutProperties> {
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

fn parse_parent_label(node: &MiniNode) -> Result<ParentLabelLayout> {
    reject_unknown(&node.attributes, &[("", "val")], "parentLabelLayout")?;
    require_empty_content(node, "parentLabelLayout")?;
    match required(&node.attributes, "", "val")? {
        "none" => Ok(ParentLabelLayout::None),
        "banner" => Ok(ParentLabelLayout::Banner),
        "overlapping" => Ok(ParentLabelLayout::Overlapping),
        _ => invalid("invalid  parentLabelLayout value"),
    }
}

fn parse_region_label(node: &MiniNode) -> Result<RegionLabelLayout> {
    reject_unknown(&node.attributes, &[("", "val")], "regionLabelLayout")?;
    require_empty_content(node, "regionLabelLayout")?;
    match required(&node.attributes, "", "val")? {
        "none" => Ok(RegionLabelLayout::None),
        "bestFitOnly" => Ok(RegionLabelLayout::BestFitOnly),
        "showAll" => Ok(RegionLabelLayout::ShowAll),
        _ => invalid("invalid  regionLabelLayout value"),
    }
}

fn parse_visibility(node: &MiniNode) -> Result<ElementVisibility> {
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

fn parse_binning(node: &MiniNode) -> Result<Binning> {
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

fn parse_geography(node: &MiniNode) -> Result<Geography> {
    let allowed = &[
        ("", "projectionType"),
        ("", "viewedRegionType"),
        ("", "cultureLanguage"),
        ("", "cultureRegion"),
        ("", "attribution"),
    ];
    reject_unknown(&node.attributes, allowed, "geography")?;
    if !node.text.trim().is_empty() {
        return invalid("unexpected text in  geography");
    }
    let projection = optional(&node.attributes, "", "projectionType")
        .map(|value| match value {
            "mercator" => Ok(GeoProjection::Mercator),
            "miller" => Ok(GeoProjection::Miller),
            "robinson" => Ok(GeoProjection::Robinson),
            "albers" => Ok(GeoProjection::Albers),
            _ => invalid("invalid  geography projectionType"),
        })
        .transpose()?;
    let viewed_region = optional(&node.attributes, "", "viewedRegionType")
        .map(|value| match value {
            "dataOnly" => Ok(GeoMappingLevel::DataOnly),
            "postalCode" => Ok(GeoMappingLevel::PostalCode),
            "county" => Ok(GeoMappingLevel::County),
            "state" => Ok(GeoMappingLevel::State),
            "countryRegion" => Ok(GeoMappingLevel::CountryRegion),
            "countryRegionList" => Ok(GeoMappingLevel::CountryRegionList),
            "world" => Ok(GeoMappingLevel::World),
            _ => invalid("invalid  geography viewedRegionType"),
        })
        .transpose()?;
    let culture_language = bounded_required(node, "cultureLanguage", MAX_CULTURE_NAME_LEN)?;
    let culture_region = bounded_required(node, "cultureRegion", MAX_CULTURE_NAME_LEN)?;
    let attribution = bounded_required(node, "attribution", MAX_ATTRIBUTION_LEN)?;
    let mut cache = None;
    for child in &node.children {
        if child.namespace != CX || child.name != "geoCache" || cache.is_some() {
            return invalid("invalid or duplicate  geography child");
        }
        cache = Some(parse_geo_cache(child)?);
    }
    let has_cache = cache.is_some();
    Ok(Geography {
        projection,
        viewed_region,
        culture_language,
        culture_region,
        attribution,
        has_cache,
        cache,
    })
}

fn parse_geo_cache(node: &MiniNode) -> Result<GeoCache> {
    reject_unknown(&node.attributes, &[("", "provider")], "geoCache")?;
    if !node.text.trim().is_empty() {
        return invalid("unexpected text in  geoCache");
    }
    let provider = geo_required_string(node, "provider", MAX_GEO_STRING_LEN)?;
    if node.children.is_empty() {
        return invalid(" geoCache requires binary or clear content");
    }
    if node.children.len() > MAX_GEO_CACHE_ENTRIES {
        return limit(" geography cache entries");
    }
    let mut entries = Vec::with_capacity(node.children.len());
    for child in &node.children {
        if child.namespace != CX {
            return invalid("foreign direct child in  geoCache");
        }
        entries.push(match child.name.as_str() {
            "binary" => {
                reject_unknown(&child.attributes, &[], "geography binary")?;
                if !child.children.is_empty() {
                    return invalid(" geography binary contains elements");
                }
                let (encoded_characters, decoded_bytes) = validate_geo_base64(&child.text)?;
                GeoCacheEntry::Binary {
                    encoded_characters,
                    decoded_bytes,
                }
            },
            "clear" => GeoCacheEntry::Clear(parse_geo_clear(child)?),
            _ => return invalid("invalid direct child in  geoCache"),
        });
    }
    Ok(GeoCache { provider, entries })
}

fn parse_geo_clear(node: &MiniNode) -> Result<GeoClear> {
    reject_geo_container(node, "geography clear cache")?;
    let mut result = GeoClear::default();
    let mut rank = 0u8;
    let mut seen = HashSet::new();
    for child in &node.children {
        let current = geo_ordered_child(
            child,
            &[
                "geoLocationQueryResults",
                "geoDataEntityQueryResults",
                "geoDataPointToEntityQueryResults",
                "geoChildEntitiesQueryResults",
                "geoParentEntitiesQueryResults",
            ],
        )?;
        if current < rank || !seen.insert(child.name.as_str()) {
            return invalid(" clear geography cache children are out of order or duplicated");
        }
        rank = current;
        match child.name.as_str() {
            "geoLocationQueryResults" => {
                result.location_query_results = Some(parse_geo_collection(
                    child,
                    "geoLocationQueryResult",
                    parse_geo_location_query_result,
                )?)
            },
            "geoDataEntityQueryResults" => {
                result.data_entity_query_results = Some(parse_geo_collection(
                    child,
                    "geoDataEntityQueryResult",
                    parse_geo_data_entity_query_result,
                )?)
            },
            "geoDataPointToEntityQueryResults" => {
                result.data_point_to_entity_query_results = Some(parse_geo_collection(
                    child,
                    "geoDataPointToEntityQueryResult",
                    parse_geo_data_point_to_entity_query_result,
                )?)
            },
            "geoChildEntitiesQueryResults" => {
                result.child_entities_query_results = Some(parse_geo_collection(
                    child,
                    "geoChildEntitiesQueryResult",
                    parse_geo_child_entities_query_result,
                )?)
            },
            "geoParentEntitiesQueryResults" => {
                result.parent_entities_query_results = Some(parse_geo_collection(
                    child,
                    "geoParentEntitiesQueryResult",
                    parse_geo_parent_entities_query_result,
                )?)
            },
            _ => unreachable!(),
        }
    }
    Ok(result)
}

fn parse_geo_collection<T>(
    node: &MiniNode,
    item_name: &str,
    parser: fn(&MiniNode) -> Result<T>,
) -> Result<Vec<T>> {
    reject_geo_container(node, &node.name)?;
    if node.children.len() > MAX_GEO_RESULTS {
        return limit(" geography query results");
    }
    node.children
        .iter()
        .map(|child| {
            if child.namespace != CX || child.name != item_name {
                return invalid(format!("invalid direct child in  {}", node.name));
            }
            parser(child)
        })
        .collect()
}

fn parse_geo_location_query_result(node: &MiniNode) -> Result<GeoLocationQueryResult> {
    reject_geo_container(node, "geoLocationQueryResult")?;
    let mut result = GeoLocationQueryResult::default();
    for child in geo_unique_ordered(node, &["geoLocationQuery", "geoLocations"])? {
        if child.name == "geoLocationQuery" {
            result.query = Some(parse_geo_location_query(child)?);
        } else {
            reject_geo_container(child, "geoLocations")?;
            if child.children.len() > 1 {
                return invalid("geoLocations permits at most one geoLocation");
            }
            result.location = child
                .children
                .first()
                .map(|value| {
                    if value.namespace != CX || value.name != "geoLocation" {
                        return invalid("invalid direct child in geoLocations");
                    }
                    parse_geo_location(value)
                })
                .transpose()?;
        }
    }
    Ok(result)
}

fn parse_geo_location_query(node: &MiniNode) -> Result<GeoLocationQuery> {
    let allowed = &[
        ("", "countryRegion"),
        ("", "adminDistrict1"),
        ("", "adminDistrict2"),
        ("", "postalCode"),
        ("", "entityType"),
    ];
    reject_unknown(&node.attributes, allowed, "geoLocationQuery")?;
    require_empty_content(node, "geoLocationQuery")?;
    Ok(GeoLocationQuery {
        country_region: geo_optional_string(node, "countryRegion", MAX_GEO_STRING_LEN)?,
        admin_district1: geo_optional_string(node, "adminDistrict1", MAX_GEO_STRING_LEN)?,
        admin_district2: geo_optional_string(node, "adminDistrict2", MAX_GEO_STRING_LEN)?,
        postal_code: geo_optional_string(node, "postalCode", MAX_GEO_STRING_LEN)?,
        entity_type: parse_geo_entity_type(required(&node.attributes, "", "entityType")?)?,
    })
}

fn parse_geo_location(node: &MiniNode) -> Result<GeoLocation> {
    reject_unknown(
        &node.attributes,
        &[
            ("", "latitude"),
            ("", "longitude"),
            ("", "entityName"),
            ("", "entityType"),
        ],
        "geoLocation",
    )?;
    if !node.text.trim().is_empty() {
        return invalid("unexpected text in  geoLocation");
    }
    let address = match node.children.as_slice() {
        [] => None,
        [child] if child.namespace == CX && child.name == "address" => {
            Some(parse_geo_address(child)?)
        },
        _ => return invalid("geoLocation permits at most one ordered address"),
    };
    Ok(GeoLocation {
        latitude: geo_optional_double(node, "latitude")?,
        longitude: geo_optional_double(node, "longitude")?,
        entity_name: geo_required_string(node, "entityName", MAX_GEO_STRING_LEN)?,
        entity_type: parse_geo_entity_type(required(&node.attributes, "", "entityType")?)?,
        address,
    })
}

fn parse_geo_address(node: &MiniNode) -> Result<GeoAddress> {
    let allowed = &[
        ("", "address1"),
        ("", "countryRegion"),
        ("", "adminDistrict1"),
        ("", "adminDistrict2"),
        ("", "postalCode"),
        ("", "locality"),
        ("", "isoCountryCode"),
    ];
    reject_unknown(&node.attributes, allowed, "geography address")?;
    require_empty_content(node, "geography address")?;
    Ok(GeoAddress {
        address1: geo_optional_string(node, "address1", MAX_GEO_STRING_LEN)?,
        country_region: geo_optional_string(node, "countryRegion", MAX_GEO_STRING_LEN)?,
        admin_district1: geo_optional_string(node, "adminDistrict1", MAX_GEO_STRING_LEN)?,
        admin_district2: geo_optional_string(node, "adminDistrict2", MAX_GEO_STRING_LEN)?,
        postal_code: geo_optional_string(node, "postalCode", MAX_GEO_STRING_LEN)?,
        locality: geo_optional_string(node, "locality", MAX_GEO_STRING_LEN)?,
        iso_country_code: geo_optional_string(node, "isoCountryCode", MAX_GEO_STRING_LEN)?,
    })
}

fn parse_geo_data_entity_query_result(node: &MiniNode) -> Result<GeoDataEntityQueryResult> {
    reject_geo_container(node, "geoDataEntityQueryResult")?;
    let mut result = GeoDataEntityQueryResult::default();
    for child in geo_unique_ordered(node, &["geoDataEntityQuery", "geoData"])? {
        if child.name == "geoDataEntityQuery" {
            result.query = Some(parse_geo_data_entity_query(child)?);
        } else {
            result.data = Some(parse_geo_data(child)?);
        }
    }
    Ok(result)
}

fn parse_geo_data_entity_query(node: &MiniNode) -> Result<GeoDataEntityQuery> {
    reject_unknown(
        &node.attributes,
        &[("", "entityType"), ("", "entityId")],
        "geoDataEntityQuery",
    )?;
    require_empty_content(node, "geoDataEntityQuery")?;
    Ok(GeoDataEntityQuery {
        entity_type: parse_geo_entity_type(required(&node.attributes, "", "entityType")?)?,
        entity_id: geo_required_string(node, "entityId", MAX_GEO_STRING_LEN)?,
    })
}

fn parse_geo_data(node: &MiniNode) -> Result<GeoData> {
    reject_unknown(
        &node.attributes,
        &[
            ("", "entityName"),
            ("", "entityId"),
            ("", "east"),
            ("", "west"),
            ("", "north"),
            ("", "south"),
        ],
        "geoData",
    )?;
    if !node.text.trim().is_empty() {
        return invalid("unexpected text in  geoData");
    }
    let mut polygons = None;
    let mut copyrights = None;
    for child in geo_unique_ordered(node, &["geoPolygons", "copyrights"])? {
        if child.name == "geoPolygons" {
            polygons = Some(parse_geo_collection(
                child,
                "geoPolygon",
                parse_geo_polygon,
            )?);
        } else {
            copyrights = Some(parse_geo_copyrights(child)?);
        }
    }
    Ok(GeoData {
        entity_name: geo_required_string(node, "entityName", MAX_GEO_STRING_LEN)?,
        entity_id: geo_required_string(node, "entityId", MAX_GEO_STRING_LEN)?,
        east: geo_required_double(node, "east")?,
        west: geo_required_double(node, "west")?,
        north: geo_required_double(node, "north")?,
        south: geo_required_double(node, "south")?,
        polygons,
        copyrights,
    })
}

fn parse_geo_polygon(node: &MiniNode) -> Result<GeoPolygon> {
    reject_unknown(
        &node.attributes,
        &[("", "polygonId"), ("", "numPoints"), ("", "pcaRings")],
        "geoPolygon",
    )?;
    require_empty_content(node, "geoPolygon")?;
    let num_points = geo_required_string(node, "numPoints", 128)?;
    validate_xsd_integer(&num_points, "geoPolygon numPoints")?;
    Ok(GeoPolygon {
        polygon_id: geo_required_string(node, "polygonId", MAX_GEO_STRING_LEN)?,
        num_points,
        pca_rings: geo_required_string(node, "pcaRings", MAX_GEO_POLYGON_DATA_LEN)?,
    })
}

fn parse_geo_copyrights(node: &MiniNode) -> Result<Vec<String>> {
    reject_geo_container(node, "copyrights")?;
    if node.children.len() > MAX_GEO_RESULTS {
        return limit(" geography copyrights");
    }
    node.children
        .iter()
        .map(|child| {
            if child.namespace != CX
                || child.name != "copyright"
                || !child.attributes.is_empty()
                || !child.children.is_empty()
            {
                return invalid("invalid direct child in  copyrights");
            }
            if child.text.len() > MAX_GEO_STRING_LEN {
                return limit(" geography copyright");
            }
            Ok(child.text.clone())
        })
        .collect()
}

fn parse_geo_data_point_to_entity_query_result(
    node: &MiniNode,
) -> Result<GeoDataPointToEntityQueryResult> {
    reject_geo_container(node, "geoDataPointToEntityQueryResult")?;
    let mut result = GeoDataPointToEntityQueryResult::default();
    for child in geo_unique_ordered(node, &["geoDataPointQuery", "geoDataPointToEntityQuery"])? {
        if child.name == "geoDataPointQuery" {
            result.point_query = Some(parse_geo_data_point_query(child)?);
        } else {
            result.entity_query = Some(parse_geo_data_point_to_entity_query(child)?);
        }
    }
    Ok(result)
}

fn parse_geo_data_point_query(node: &MiniNode) -> Result<GeoDataPointQuery> {
    reject_unknown(
        &node.attributes,
        &[("", "entityType"), ("", "latitude"), ("", "longitude")],
        "geoDataPointQuery",
    )?;
    require_empty_content(node, "geoDataPointQuery")?;
    Ok(GeoDataPointQuery {
        entity_type: parse_geo_entity_type(required(&node.attributes, "", "entityType")?)?,
        latitude: geo_required_double(node, "latitude")?,
        longitude: geo_required_double(node, "longitude")?,
    })
}

fn parse_geo_data_point_to_entity_query(node: &MiniNode) -> Result<GeoDataPointToEntityQuery> {
    reject_unknown(
        &node.attributes,
        &[("", "entityType"), ("", "entityId")],
        "geoDataPointToEntityQuery",
    )?;
    require_empty_content(node, "geoDataPointToEntityQuery")?;
    Ok(GeoDataPointToEntityQuery {
        entity_type: parse_geo_entity_type(required(&node.attributes, "", "entityType")?)?,
        entity_id: geo_required_string(node, "entityId", MAX_GEO_STRING_LEN)?,
    })
}

fn parse_geo_child_entities_query_result(node: &MiniNode) -> Result<GeoChildEntitiesQueryResult> {
    reject_geo_container(node, "geoChildEntitiesQueryResult")?;
    let mut result = GeoChildEntitiesQueryResult::default();
    for child in geo_unique_ordered(node, &["geoChildEntitiesQuery", "geoChildEntities"])? {
        if child.name == "geoChildEntitiesQuery" {
            result.query = Some(parse_geo_child_entities_query(child)?);
        } else {
            result.children = Some(parse_geo_collection(
                child,
                "geoHierarchyEntity",
                parse_geo_hierarchy_entity,
            )?);
        }
    }
    Ok(result)
}

fn parse_geo_child_entities_query(node: &MiniNode) -> Result<GeoChildEntitiesQuery> {
    reject_unknown(
        &node.attributes,
        &[("", "entityId")],
        "geoChildEntitiesQuery",
    )?;
    if !node.text.trim().is_empty() {
        return invalid("unexpected text in geoChildEntitiesQuery");
    }
    let child_types = match node.children.as_slice() {
        [] => None,
        [child] if child.namespace == CX && child.name == "geoChildTypes" => {
            reject_geo_container(child, "geoChildTypes")?;
            if child.children.len() > MAX_GEO_RESULTS {
                return limit(" geography child types");
            }
            Some(
                child
                    .children
                    .iter()
                    .map(|value| {
                        if value.namespace != CX
                            || value.name != "entityType"
                            || !value.attributes.is_empty()
                            || !value.children.is_empty()
                        {
                            return invalid("invalid direct child in geoChildTypes");
                        }
                        parse_geo_entity_type(value.text.trim())
                    })
                    .collect::<Result<Vec<_>>>()?,
            )
        },
        _ => return invalid("geoChildEntitiesQuery permits at most one geoChildTypes"),
    };
    Ok(GeoChildEntitiesQuery {
        entity_id: geo_required_string(node, "entityId", MAX_GEO_STRING_LEN)?,
        child_types,
    })
}

fn parse_geo_hierarchy_entity(node: &MiniNode) -> Result<GeoHierarchyEntity> {
    reject_unknown(
        &node.attributes,
        &[("", "entityName"), ("", "entityId"), ("", "entityType")],
        "geoHierarchyEntity",
    )?;
    require_empty_content(node, "geoHierarchyEntity")?;
    Ok(GeoHierarchyEntity {
        entity_name: geo_required_string(node, "entityName", MAX_GEO_STRING_LEN)?,
        entity_id: geo_required_string(node, "entityId", MAX_GEO_STRING_LEN)?,
        entity_type: parse_geo_entity_type(required(&node.attributes, "", "entityType")?)?,
    })
}

fn parse_geo_parent_entities_query_result(node: &MiniNode) -> Result<GeoParentEntitiesQueryResult> {
    reject_geo_container(node, "geoParentEntitiesQueryResult")?;
    let children = &node.children;
    if children.is_empty()
        || children[0].namespace != CX
        || children[0].name != "geoParentEntitiesQuery"
    {
        return invalid("geoParentEntitiesQueryResult requires geoParentEntitiesQuery first");
    }
    reject_unknown(
        &children[0].attributes,
        &[("", "entityId")],
        "geoParentEntitiesQuery",
    )?;
    require_empty_content(&children[0], "geoParentEntitiesQuery")?;
    let entity_id = geo_required_string(&children[0], "entityId", MAX_GEO_STRING_LEN)?;
    let mut entity = None;
    let mut parent_entity_id = None;
    let mut rank = 0u8;
    for child in children.iter().skip(1) {
        let current = geo_ordered_child(child, &["geoEntity", "geoParentEntity"])?;
        if current < rank {
            return invalid("invalid geoParentEntitiesQueryResult order");
        }
        rank = current;
        if child.name == "geoEntity" {
            if entity.is_some() {
                return invalid("duplicate geoEntity");
            }
            reject_unknown(
                &child.attributes,
                &[("", "entityName"), ("", "entityType")],
                "geoEntity",
            )?;
            require_empty_content(child, "geoEntity")?;
            entity = Some(GeoEntity {
                entity_name: geo_required_string(child, "entityName", MAX_GEO_STRING_LEN)?,
                entity_type: parse_geo_entity_type(required(&child.attributes, "", "entityType")?)?,
            });
        } else {
            if parent_entity_id.is_some() {
                return invalid("duplicate geoParentEntity");
            }
            reject_unknown(&child.attributes, &[("", "entityId")], "geoParentEntity")?;
            require_empty_content(child, "geoParentEntity")?;
            parent_entity_id = Some(geo_required_string(child, "entityId", MAX_GEO_STRING_LEN)?);
        }
    }
    Ok(GeoParentEntitiesQueryResult {
        entity_id,
        entity,
        parent_entity_id,
    })
}

fn reject_geo_container(node: &MiniNode, label: &str) -> Result<()> {
    reject_unknown(&node.attributes, &[], label)?;
    if !node.text.trim().is_empty() {
        return invalid(format!("unexpected text in  {label}"));
    }
    Ok(())
}

fn geo_ordered_child(child: &MiniNode, names: &[&str]) -> Result<u8> {
    if child.namespace != CX {
        return invalid("foreign child in  geography cache");
    }
    names
        .iter()
        .position(|name| *name == child.name)
        .map(|value| value as u8)
        .ok_or_else(|| invalid_error(format!("invalid geography cache child '{}'", child.name)))
}

fn geo_unique_ordered<'a>(node: &'a MiniNode, names: &[&str]) -> Result<Vec<&'a MiniNode>> {
    let mut rank = 0u8;
    let mut seen = HashSet::new();
    for child in &node.children {
        let current = geo_ordered_child(child, names)?;
        if current < rank || !seen.insert(child.name.as_str()) {
            return invalid(format!("invalid {} order or cardinality", node.name));
        }
        rank = current;
    }
    Ok(node.children.iter().collect())
}

fn parse_geo_entity_type(value: &str) -> Result<GeoEntityType> {
    match value {
        "Address" => Ok(GeoEntityType::Address),
        "AdminDistrict" => Ok(GeoEntityType::AdminDistrict),
        "AdminDistrict2" => Ok(GeoEntityType::AdminDistrict2),
        "AdminDistrict3" => Ok(GeoEntityType::AdminDistrict3),
        "Continent" => Ok(GeoEntityType::Continent),
        "CountryRegion" => Ok(GeoEntityType::CountryRegion),
        "Locality" => Ok(GeoEntityType::Locality),
        "Ocean" => Ok(GeoEntityType::Ocean),
        "Planet" => Ok(GeoEntityType::Planet),
        "PostalCode" => Ok(GeoEntityType::PostalCode),
        "Region" => Ok(GeoEntityType::Region),
        "Unsupported" => Ok(GeoEntityType::Unsupported),
        _ => invalid("invalid  geography entity type"),
    }
}

fn geo_required_string(node: &MiniNode, name: &str, maximum: usize) -> Result<String> {
    let value = required(&node.attributes, "", name)?;
    if value.len() > maximum {
        return limit(" geography string");
    }
    Ok(value.to_owned())
}

fn geo_optional_string(node: &MiniNode, name: &str, maximum: usize) -> Result<Option<String>> {
    optional(&node.attributes, "", name)
        .map(|value| {
            if value.len() > maximum {
                return limit(" geography string");
            }
            Ok(value.to_owned())
        })
        .transpose()
}

fn geo_required_double(node: &MiniNode, name: &str) -> Result<String> {
    let value = required(&node.attributes, "", name)?;
    if !valid_xml_double(value) {
        return invalid(format!("invalid  geography {name}"));
    }
    Ok(value.to_owned())
}

fn geo_optional_double(node: &MiniNode, name: &str) -> Result<Option<String>> {
    optional(&node.attributes, "", name)
        .map(|value| {
            if !valid_xml_double(value) {
                return invalid(format!("invalid  geography {name}"));
            }
            Ok(value.to_owned())
        })
        .transpose()
}

fn validate_xsd_integer(value: &str, label: &str) -> Result<()> {
    let digits = value.strip_prefix(['+', '-']).unwrap_or(value);
    if digits.is_empty() || !digits.bytes().all(|byte| byte.is_ascii_digit()) {
        return invalid(format!("invalid  {label}"));
    }
    Ok(())
}

fn validate_geo_base64(value: &str) -> Result<(usize, usize)> {
    let mut encoded = 0usize;
    let mut padding = 0usize;
    let mut saw_padding = false;
    for byte in value.bytes() {
        if matches!(byte, b' ' | b'\t' | b'\r' | b'\n') {
            continue;
        }
        encoded += 1;
        if byte == b'=' {
            saw_padding = true;
            padding += 1;
            if padding > 2 {
                return invalid("invalid  geography base64 padding");
            }
        } else if saw_padding || !(byte.is_ascii_alphanumeric() || matches!(byte, b'+' | b'/')) {
            return invalid("invalid  geography base64 data");
        }
    }
    if !encoded.is_multiple_of(4) {
        return invalid("invalid  geography base64 length");
    }
    let decoded = encoded
        .checked_div(4)
        .and_then(|value| value.checked_mul(3))
        .and_then(|value| value.checked_sub(padding))
        .ok_or_else(|| invalid_error(" geography base64 size overflow"))?;
    if decoded > MAX_GEO_BINARY_BYTES {
        return limit(" geography binary data");
    }
    Ok((encoded, decoded))
}

fn parse_statistics(node: &MiniNode) -> Result<Option<QuartileMethod>> {
    reject_unknown(&node.attributes, &[("", "quartileMethod")], "statistics")?;
    require_empty_content(node, "statistics")?;
    optional(&node.attributes, "", "quartileMethod")
        .map(|value| match value {
            "inclusive" => Ok(QuartileMethod::Inclusive),
            "exclusive" => Ok(QuartileMethod::Exclusive),
            _ => invalid("invalid  statistics quartileMethod"),
        })
        .transpose()
}

fn parse_subtotals(node: &MiniNode) -> Result<Vec<u32>> {
    reject_unknown(&node.attributes, &[], "subtotals")?;
    if !node.text.trim().is_empty() {
        return invalid("unexpected text in  subtotals");
    }
    let mut values = Vec::new();
    let mut unique = HashSet::new();
    for child in &node.children {
        if values.len() >= MAX_SUBTOTALS {
            return limit(" subtotal count");
        }
        if child.namespace != CX
            || child.name != "idx"
            || !child.attributes.is_empty()
            || !child.children.is_empty()
        {
            return invalid("invalid  subtotal index");
        }
        let value = parse_u32(child.text.trim(), "subtotal index")?;
        if !unique.insert(value) {
            return invalid("duplicate  subtotal index");
        }
        values.push(value);
    }
    Ok(values)
}

fn parse_double_or_auto(value: &str, label: &str) -> Result<DoubleOrAutomatic> {
    if value == "auto" {
        return Ok(DoubleOrAutomatic::Automatic);
    }
    if !valid_xml_double(value) {
        return invalid(format!("invalid  {label}"));
    }
    Ok(DoubleOrAutomatic::Number(value.to_owned()))
}

fn parse_nonnegative_or_auto(value: &str, label: &str) -> Result<DoubleOrAutomatic> {
    if value == "auto" {
        return Ok(DoubleOrAutomatic::Automatic);
    }
    let number = value
        .parse::<f64>()
        .map_err(|_| invalid_error(format!("invalid  {label}")))?;
    if number.is_nan() || number < 0.0 {
        return invalid(format!("invalid  {label}"));
    }
    Ok(DoubleOrAutomatic::Number(value.to_owned()))
}

fn parse_positive_or_auto(value: &str, label: &str) -> Result<DoubleOrAutomatic> {
    if value == "auto" {
        return Ok(DoubleOrAutomatic::Automatic);
    }
    let number = value
        .parse::<f64>()
        .map_err(|_| invalid_error(format!("invalid  {label}")))?;
    if number.is_nan() || number <= 0.0 {
        return invalid(format!("invalid  {label}"));
    }
    Ok(DoubleOrAutomatic::Number(value.to_owned()))
}

fn require_empty_element(node: &MiniNode, label: &str) -> Result<()> {
    reject_unknown(&node.attributes, &[], label)?;
    require_empty_content(node, label)
}

fn require_empty_content(node: &MiniNode, label: &str) -> Result<()> {
    if !node.children.is_empty() || !node.text.trim().is_empty() {
        invalid(format!(" {label} must be empty"))
    } else {
        Ok(())
    }
}

fn bounded_required(node: &MiniNode, name: &str, max: usize) -> Result<String> {
    let value = required(&node.attributes, "", name)?;
    if value.len() > max {
        return limit(" attribute string");
    }
    Ok(value.to_owned())
}

fn parse_drawing_payload(node: &MiniNode, label: &str) -> Result<DrawingPayload> {
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

fn parse_shared_text(node: &MiniNode, label: &str) -> Result<Text> {
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

fn parse_text_data(node: &MiniNode) -> Result<Text> {
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

fn parse_value_colors(node: &MiniNode) -> Result<ValueColors> {
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

fn parse_solid_color(node: &MiniNode) -> Result<SolidColor> {
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

fn parse_value_color_positions(node: &MiniNode) -> Result<ValueColorPositions> {
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

fn parse_color_position(node: &MiniNode, allow_extreme: bool) -> Result<ColorPosition> {
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

fn parse_position_value(node: &MiniNode, label: &str) -> Result<String> {
    reject_unknown(&node.attributes, &[("", "val")], label)?;
    require_empty_content(node, label)?;
    Ok(required(&node.attributes, "", "val")?.to_owned())
}

fn parse_data_point(node: &MiniNode) -> Result<DataPoint> {
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

fn parse_data_labels(node: &MiniNode) -> Result<DataLabels> {
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

fn parse_data_label(node: &MiniNode) -> Result<DataLabel> {
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

fn parse_hidden_label(node: &MiniNode) -> Result<u32> {
    reject_unknown(&node.attributes, &[("", "idx")], "dataLabelHidden")?;
    require_empty_content(node, "dataLabelHidden")?;
    parse_u32(
        required(&node.attributes, "", "idx")?,
        "hidden data label index",
    )
}

fn parse_number_format(node: &MiniNode) -> Result<NumberFormat> {
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

fn parse_label_visibility(node: &MiniNode) -> Result<DataLabelVisibility> {
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

fn parse_separator(node: &MiniNode) -> Result<String> {
    reject_unknown(&node.attributes, &[], "data label separator")?;
    if !node.children.is_empty() {
        return invalid(" data label separator must have simple content");
    }
    if node.text.len() > MAX_LABEL_TEXT_BYTES {
        return limit(" data label separator bytes");
    }
    Ok(node.text.clone())
}

fn parse_label_position(value: &str) -> Result<DataLabelPosition> {
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

fn parse_series(node: &MiniNode) -> Result<SeriesDataReference> {
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

fn series_child_rank(node: &MiniNode) -> Option<u8> {
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
