//! Chart content validation.

use litchi_core::{Error, Result};
use litchi_odf_common::chart::{Element, Kind, read};
use std::collections::BTreeSet;

const CHART: &str = "urn:oasis:names:tc:opendocument:xmlns:chart:1.0";
const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TABLE: &str = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";

/// Validate a UTF-8 content part before authoring it into a package.
pub(crate) fn validate(xml: &str) -> Result<()> {
    let chart = read(xml)?;
    let _ = chart.chart_class()?;
    if let Some(plot_area) = chart.plot_area() {
        for axis in plot_area.axes() {
            let _ = axis.dimension()?;
        }
    }
    validate_tree(&chart, crate::Limits::default())?;
    Ok(())
}

pub(crate) fn validate_tree(root: &Element, limits: crate::Limits) -> Result<()> {
    validate_chart_schema(root, limits)?;
    let mut stack = vec![root];
    while let Some(element) = stack.pop() {
        for attribute in element.attributes() {
            if attribute.value().len() > limits.max_scalar_bytes() {
                return Err(Error::InvalidFormat(
                    "ODC attribute exceeds the caller-selected scalar limit".into(),
                ));
            }
            let is_range = (attribute.namespace_uri() == Some(CHART)
                && matches!(
                    attribute.local_name(),
                    "values-cell-range-address" | "label-cell-address"
                ))
                || (attribute.namespace_uri() == Some(TABLE)
                    && matches!(attribute.local_name(), "cell-range-address" | "cell-range"));
            if is_range {
                crate::validation::validate_range_list_with_limits(
                    attribute.value(),
                    limits.max_scalar_bytes(),
                    limits.max_range_items(),
                )?;
            }
            if attribute.namespace_uri() == Some(TABLE) && attribute.local_name() == "formula" {
                crate::validation::validate_formula_with_limit(
                    attribute.value(),
                    limits.max_scalar_bytes(),
                )?;
            }
        }
        if element.children().is_empty() && element.all_text().len() > limits.max_scalar_bytes() {
            return Err(Error::InvalidFormat(
                "ODC text exceeds the caller-selected scalar limit".into(),
            ));
        }
        stack.extend(element.children());
    }
    Ok(())
}

fn validate_chart_schema(root: &Element, limits: crate::Limits) -> Result<()> {
    for (kind, label) in [
        (Kind::Title, "chart:title"),
        (Kind::Subtitle, "chart:subtitle"),
        (Kind::Footer, "chart:footer"),
        (Kind::Legend, "chart:legend"),
    ] {
        ensure_at_most_one(root, kind, label)?;
    }
    if root.children_of_kind(Kind::PlotArea).count() != 1 {
        return invalid("ODC chart requires exactly one chart:plot-area");
    }
    if let Some(legend) = root.legend() {
        let _ = legend.position()?;
    }
    let plot = root
        .plot_area()
        .ok_or_else(|| Error::InvalidFormat("ODC chart plot area is absent".into()))?;
    if plot
        .element()
        .attribute(Some(CHART), "data-source-has-labels")
        .is_some()
    {
        let _ = plot.data_source_labels()?;
    }
    for (kind, label) in [
        (Kind::Wall, "chart:wall"),
        (Kind::Floor, "chart:floor"),
        (Kind::StockGainMarker, "chart:stock-gain-marker"),
        (Kind::StockLossMarker, "chart:stock-loss-marker"),
        (Kind::StockRangeLine, "chart:stock-range-line"),
    ] {
        ensure_at_most_one(plot.element(), kind, label)?;
    }
    let mut axis_names = BTreeSet::new();
    let mut axis_count = 0usize;
    for axis in plot.axes() {
        axis_count = axis_count
            .checked_add(1)
            .ok_or_else(|| Error::InvalidFormat("ODC axis count overflow".into()))?;
        if axis_count > limits.max_axes() {
            return invalid("ODC axis count exceeds the caller-selected limit");
        }
        let _ = axis.dimension()?;
        ensure_at_most_one(axis.element(), Kind::Title, "axis chart:title")?;
        ensure_at_most_one(axis.element(), Kind::Categories, "chart:categories")?;
        if let Some(name) = axis.name()
            && !axis_names.insert(name)
        {
            return invalid("ODC chart contains duplicate axis names");
        }
        for grid in axis.grids() {
            let _ = grid.class()?;
        }
    }
    let mut expanded_points = 0usize;
    let mut series_ids = BTreeSet::new();
    let mut series_count = 0usize;
    for value in plot.series() {
        series_count = series_count
            .checked_add(1)
            .ok_or_else(|| Error::InvalidFormat("ODC series count overflow".into()))?;
        if series_count > limits.max_series() {
            return invalid("ODC series count exceeds the caller-selected limit");
        }
        if let Some(id) = value.xml_id()
            && !series_ids.insert(id)
        {
            return invalid("ODC chart contains duplicate series xml:id values");
        }
        if value.element().attribute(Some(CHART), "class").is_some() {
            let _ = value.element().chart_class()?;
        }
        if let Some(axis) = value.attached_axis()
            && !axis_names.contains(axis)
        {
            return invalid("ODC series references an unknown axis");
        }
        ensure_at_most_one(value.element(), Kind::DataLabel, "series chart:data-label")?;
        ensure_at_most_one(value.element(), Kind::MeanValue, "chart:mean-value")?;
        ensure_at_most_one(
            value.element(),
            Kind::ErrorIndicator,
            "chart:error-indicator",
        )?;
        for point in value.data_points() {
            let repeated = usize::try_from(point.repeated()?).map_err(|error| {
                Error::InvalidFormat(format!(
                    "ODC data-point repeat count exceeds this platform: {error}"
                ))
            })?;
            expanded_points = expanded_points
                .checked_add(repeated)
                .ok_or_else(|| Error::InvalidFormat("ODC data-point count overflow".into()))?;
            if expanded_points > limits.max_data_points() {
                return invalid("ODC data-point count exceeds the caller-selected limit");
            }
        }
    }
    validate_cached_table(root, limits)
}

fn ensure_at_most_one(element: &Element, kind: Kind, label: &str) -> Result<()> {
    if element.children_of_kind(kind).count() > 1 {
        return invalid(format!("ODC chart contains duplicate {label} elements"));
    }
    Ok(())
}

fn validate_cached_table(root: &Element, limits: crate::Limits) -> Result<()> {
    let mut tables = root.children().iter().filter(|element| {
        element.namespace_uri() == Some(TABLE) && element.local_name() == "table"
    });
    let first_table = tables.next();
    if tables.next().is_some() {
        return invalid("ODC chart contains more than one cached table");
    }
    let Some(table) = first_table else {
        return Ok(());
    };
    if table.attribute(Some(TABLE), "name").is_none() {
        return invalid("ODC cached table requires table:name");
    }
    let mut expanded_rows = 0usize;
    let mut expanded_cells = 0usize;
    let mut stack = vec![table];
    while let Some(element) = stack.pop() {
        if element.namespace_uri() == Some(TABLE) && element.local_name() == "table-row" {
            let row_repeat = repeat(element, "number-rows-repeated")?;
            expanded_rows = expanded_rows
                .checked_add(row_repeat)
                .ok_or_else(|| Error::InvalidFormat("ODC cached-row count overflow".into()))?;
            if expanded_rows > limits.max_cached_rows() {
                return invalid("ODC cached-row count exceeds the caller-selected limit");
            }
            let mut row_cells = 0usize;
            for cell in element.children().iter().filter(|child| {
                child.namespace_uri() == Some(TABLE) && child.local_name() == "table-cell"
            }) {
                validate_cached_cell(cell)?;
                row_cells = row_cells
                    .checked_add(repeat(cell, "number-columns-repeated")?)
                    .ok_or_else(|| Error::InvalidFormat("ODC cached-cell count overflow".into()))?;
            }
            expanded_cells = expanded_cells
                .checked_add(row_cells.checked_mul(row_repeat).ok_or_else(|| {
                    Error::InvalidFormat("ODC expanded cached-cell count overflow".into())
                })?)
                .ok_or_else(|| {
                    Error::InvalidFormat("ODC expanded cached-cell count overflow".into())
                })?;
            if expanded_cells > limits.max_cached_cells() {
                return invalid("ODC cached-cell count exceeds the caller-selected limit");
            }
        }
        stack.extend(element.children());
    }
    Ok(())
}

fn validate_cached_cell(cell: &Element) -> Result<()> {
    let value_type = cell.attribute(Some(OFFICE), "value-type");
    match value_type {
        None | Some("string") => {},
        Some("float" | "percentage") => {
            finite_number(cell, "value")?;
        },
        Some("currency") => {
            finite_number(cell, "value")?;
            if cell.attribute(Some(OFFICE), "currency").is_none() {
                return invalid("ODC currency cell requires office:currency");
            }
        },
        Some("boolean") => {
            if !matches!(
                cell.attribute(Some(OFFICE), "boolean-value"),
                Some("true" | "false")
            ) {
                return invalid("ODC boolean cell requires true or false");
            }
        },
        Some("date") => required_value(cell, "date-value")?,
        Some("time") => required_value(cell, "time-value")?,
        Some(_) => return invalid("ODC cached cell uses an unsupported office:value-type"),
    }
    Ok(())
}

fn finite_number(cell: &Element, local: &str) -> Result<()> {
    let lexical = cell
        .attribute(Some(OFFICE), local)
        .ok_or_else(|| Error::InvalidFormat(format!("ODC cached cell requires office:{local}")))?;
    let value = lexical.parse::<f64>().map_err(|error| {
        Error::InvalidFormat(format!("invalid ODC cached numeric value: {error}"))
    })?;
    if !value.is_finite() {
        return invalid("ODC cached numeric value must be finite");
    }
    Ok(())
}

fn required_value(cell: &Element, local: &str) -> Result<()> {
    if cell
        .attribute(Some(OFFICE), local)
        .is_none_or(str::is_empty)
    {
        return invalid(format!("ODC cached cell requires office:{local}"));
    }
    Ok(())
}

fn repeat(element: &Element, local: &str) -> Result<usize> {
    let value = element.attribute(Some(TABLE), local).unwrap_or("1");
    let parsed = value
        .parse::<usize>()
        .map_err(|error| Error::InvalidFormat(format!("invalid ODC repeat count: {error}")))?;
    if parsed == 0 {
        return invalid("ODC repeat count must be nonzero");
    }
    Ok(parsed)
}

fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::InvalidFormat(message.into()))
}

#[cfg(test)]
mod tests {
    use super::validate;

    #[test]
    fn requires_family_body() {
        assert!(validate(
            r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0"><office:body><office:chart><chart:chart/></office:chart></office:body></office:document-content>"#
        )
        .is_err());
        assert!(validate(
            r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0"><office:body><office:chart><chart:chart chart:class="chart:line"><chart:plot-area/></chart:chart></office:chart></office:body></office:document-content>"#
        )
        .is_ok());
        assert!(validate(
            r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0"><office:body><office:chart><chart:chart><chart:plot-area/></chart:chart></office:chart></office:body></office:document-content>"#
        )
        .is_err());
        assert!(validate("<office:text/>").is_err());
    }
}
