//! Bounded chart inventory and opt-in semantic checks.
//!
//! The regular chart parser deliberately accepts a useful, inert projection of
//! a BIFF8 chart substream. This module adds the stricter relationships that
//! callers need when they want to reason about chart data without evaluating a
//! formula or rendering a chart. The checks follow the MS-XLS `SERIESFORMAT`
//! and `SERIESDATA` rules; the surrounding `OfficeArt` drawing graph remains
//! outside this capability.

use std::collections::HashSet;

use super::model::{
    CacheKind, Chart, DataKind, DataLink, Format, GroupKind, Kind, Limits, Role, Source, Value,
};
use super::wire::{BLANK, BRAI, CHART, DATA_FORMAT, LABEL, NUMBER, SCATTER, SERIES, SI_INDEX};
use crate::{Error, Result};

const SERIES_AI_ROLES: [Role; 4] = [Role::Name, Role::Values, Role::Categories, Role::Bubbles];

/// Whether the parsed projection contains enough known data-link structure
/// for the semantic checks to cover every series.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum SemanticCompleteness {
    /// Every series has the four ordered AI links and no opaque chart record is
    /// retained by this XLS projection.
    Complete,
    /// At least one series, formula, or chart record is only partially modeled.
    Partial,
}

/// Read-only summary of one parsed BIFF8 chart.
///
/// This is an inventory, not a rendering description. Formula tokens remain
/// inert; `cell_reference_count` counts only references decoded by the bounded
/// chart formula codec, and `opaque_formula_count` identifies cell links whose
/// token grammar was not decoded into a reference.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ChartInventory {
    /// Concise family derived from the chart groups.
    pub kind: Kind,
    /// Number of `Series` records.
    pub series_count: usize,
    /// Number of `ChartFormat` groups.
    pub group_count: usize,
    /// Number of parsed `Axis` records.
    pub axis_count: usize,
    /// Number of `BRAI` data links retained across all series.
    pub data_link_count: usize,
    /// Number of decoded area references across all data links.
    pub cell_reference_count: usize,
    /// Number of cell links whose formula has no decoded reference.
    pub opaque_formula_count: usize,
    /// Number of cached `Number`, `Label`, or `Blank` values.
    pub cached_value_count: usize,
    /// Number of retained records not interpreted by this XLS projection.
    pub unknown_record_count: usize,
    /// Unique retained record identifiers in source order.
    pub unknown_record_types: Vec<u16>,
    /// Whether the known semantic projection is complete enough for the
    /// stricter [`Chart::validate_semantics`] pass.
    pub semantic_completeness: SemanticCompleteness,
}

impl Chart {
    /// Builds a bounded, read-only inventory of the parsed chart.
    ///
    /// The ordinary structural validation is performed first, so a caller can
    /// use the returned counts as a checked projection of the chart model. The
    /// stricter relationship checks are available through
    /// [`Chart::validate_semantics`].
    ///
    /// # Errors
    ///
    /// Returns an error when the chart fails structural validation, a checked
    /// count overflows, or the bounded inventory allocation cannot be made.
    pub fn inventory(&self, limits: Limits) -> Result<ChartInventory> {
        self.validate(limits)?;

        let data_link_count = self
            .series
            .iter()
            .try_fold(0usize, |count, series| {
                count.checked_add(series.links.len())
            })
            .ok_or_else(|| Error::InvalidData("chart data-link count overflow".into()))?;
        let cell_reference_count = self
            .series
            .iter()
            .flat_map(|series| series.links.iter())
            .try_fold(0usize, |count, link| {
                count.checked_add(link.references.len())
            })
            .ok_or_else(|| Error::InvalidData("chart reference count overflow".into()))?;
        let opaque_formula_count = self
            .series
            .iter()
            .flat_map(|series| series.links.iter())
            .filter(|link| link.source == Source::Cells && link.references.is_empty())
            .count();

        let mut unknown_record_types = Vec::new();
        unknown_record_types
            .try_reserve(self.unknown_records.len())
            .map_err(|_allocation_error| Error::Allocation("chart inventory record types"))?;
        let mut seen = HashSet::new();
        seen.try_reserve(self.unknown_records.len())
            .map_err(|_allocation_error| Error::Allocation("chart inventory record types"))?;
        for record in &self.unknown_records {
            if seen.insert(record.record_type) {
                unknown_record_types.push(record.record_type);
            }
        }

        Ok(ChartInventory {
            kind: self.kind(),
            series_count: self.series.len(),
            group_count: self.groups.len(),
            axis_count: self.axes.len(),
            data_link_count,
            cell_reference_count,
            opaque_formula_count,
            cached_value_count: self.cached_values.len(),
            unknown_record_count: self.unknown_records.len(),
            unknown_record_types,
            semantic_completeness: if self.unknown_records.is_empty()
                && self.validate_semantics(limits).is_ok()
            {
                SemanticCompleteness::Complete
            } else {
                SemanticCompleteness::Partial
            },
        })
    }

    /// Validates the known MS-XLS relationships needed for safe chart-data
    /// inspection.
    ///
    /// This is intentionally opt-in. [`Chart::validate`] remains the parser's
    /// compatibility validation and continues to retain partial or opaque
    /// records. This method rejects a partial `Series` AI sequence, formulas
    /// whose cell references cannot be decoded, and cross-record references
    /// that cannot be proven safe. It never evaluates a formula, follows an
    /// external workbook, or claims to validate an `OfficeArt` anchor.
    ///
    /// # Errors
    ///
    /// Returns an error when the chart fails structural validation or its
    /// known data-link, cache, or cross-record relationships are incomplete or
    /// inconsistent.
    pub fn validate_semantics(&self, limits: Limits) -> Result<()> {
        self.validate(limits)?;
        if self.series.is_empty() {
            return invalid(
                CHART,
                "semantic chart validation requires at least one Series",
            );
        }

        for (index, series) in self.series.iter().enumerate() {
            validate_series_links(series, index)?;
            let group = self
                .groups
                .get(usize::from(series.chart_group))
                .ok_or_else(|| invalid_error(SERIES, "series chart group is missing"))?;
            if matches!(group.kind, GroupKind::Scatter { .. })
                && series.category_kind != DataKind::Numeric
            {
                return invalid(
                    SCATTER,
                    "scatter and bubble series require numeric horizontal values",
                );
            }
        }

        validate_data_formats(self)?;
        validate_cache(self)?;
        Ok(())
    }
}

fn validate_series_links(series: &super::model::Series, index: usize) -> Result<()> {
    if series.links.len() != SERIES_AI_ROLES.len() {
        return invalid(
            SERIES,
            format!("series {index} must contain exactly four ordered AI/BRAI links"),
        );
    }

    for (link, expected_role) in series.links.iter().zip(SERIES_AI_ROLES) {
        if link.role != expected_role {
            return invalid(
                BRAI,
                format!(
                    "series {index} AI links are missing or out of order: expected {expected_role:?}"
                ),
            );
        }
        validate_link_source(link, expected_role, series, index)?;
    }
    Ok(())
}

fn validate_link_source(
    link: &DataLink,
    role: Role,
    series: &super::model::Series,
    index: usize,
) -> Result<()> {
    match link.source {
        Source::Automatic => {
            if !link.formula_tokens.is_empty() || !link.references.is_empty() {
                return invalid(
                    BRAI,
                    "automatic BRAI must not carry a formula or references",
                );
            }
        },
        Source::Literal => {
            if !link.references.is_empty() {
                return invalid(BRAI, "literal BRAI must not reference worksheet cells");
            }
        },
        Source::Cells => {
            if link.formula_tokens.is_empty() || link.references.is_empty() {
                return invalid(
                    BRAI,
                    format!("series {index} has an opaque or empty cell-link formula"),
                );
            }
            if let Some(expected) = expected_count(role, series) {
                let actual = reference_cell_count(&link.references)?;
                if actual != usize::from(expected) {
                    return invalid(
                        BRAI,
                        format!(
                            "series {index} {role:?} reference contains {actual} cells, expected {expected}"
                        ),
                    );
                }
            }
        },
    }
    Ok(())
}

fn expected_count(role: Role, series: &super::model::Series) -> Option<u16> {
    match role {
        Role::Name => None,
        Role::Values => Some(series.value_count),
        Role::Categories => Some(series.category_count),
        Role::Bubbles => Some(series.bubble_count),
    }
}

fn reference_cell_count(references: &[super::model::CellRef]) -> Result<usize> {
    references.iter().try_fold(0usize, |count, reference| {
        let rows = usize::from(
            reference
                .last_row
                .checked_sub(reference.first_row)
                .ok_or_else(|| Error::InvalidData("chart reference row order is invalid".into()))?,
        )
        .checked_add(1)
        .ok_or_else(|| Error::InvalidData("chart reference row count overflow".into()))?;
        let columns = usize::from(
            reference
                .last_column
                .checked_sub(reference.first_column)
                .ok_or_else(|| {
                    Error::InvalidData("chart reference column order is invalid".into())
                })?,
        )
        .checked_add(1)
        .ok_or_else(|| Error::InvalidData("chart reference column count overflow".into()))?;
        let cells = rows
            .checked_mul(columns)
            .ok_or_else(|| Error::InvalidData("chart reference cell count overflow".into()))?;
        count
            .checked_add(cells)
            .ok_or_else(|| Error::InvalidData("chart reference cell count overflow".into()))
    })
}

fn validate_data_formats(chart: &Chart) -> Result<()> {
    for format in &chart.formatting {
        let Format::Data { series, .. } = format else {
            continue;
        };
        if usize::from(*series) >= chart.series.len() {
            return invalid(DATA_FORMAT, "DataFormat refers to a missing Series record");
        }
    }
    Ok(())
}

fn validate_cache(chart: &Chart) -> Result<()> {
    let mut cells = HashSet::new();
    cells
        .try_reserve(chart.cached_values.len())
        .map_err(|_allocation_error| Error::Allocation("chart cache validation"))?;
    for cache in &chart.cached_values {
        let Some(series) = chart.series.get(usize::from(cache.series)) else {
            return invalid(
                SI_INDEX,
                "cached chart column does not identify a Series record",
            );
        };
        let point_count = match cache.kind {
            CacheKind::Values => series.value_count,
            CacheKind::Categories => series.category_count,
            CacheKind::Bubbles => series.bubble_count,
        };
        if cache.point >= point_count {
            let record_type = match cache.value {
                Value::Blank => BLANK,
                Value::Number(_) => NUMBER,
                Value::Text(_) => LABEL,
            };
            return invalid(record_type, "cached chart point exceeds its Series count");
        }
        if !cells.insert((cache.kind, cache.point, cache.series)) {
            return invalid(SI_INDEX, "chart data cache contains a duplicate cell");
        }
        match &cache.value {
            Value::Number(value) if !value.is_finite() => {
                return invalid(NUMBER, "cached chart number must be finite");
            },
            Value::Text(value) if value.encode_utf16().count() > 255 => {
                return invalid(LABEL, "cached chart label exceeds 255 UTF-16 code units");
            },
            Value::Number(_) | Value::Text(_) | Value::Blank => {},
        }
    }
    Ok(())
}

fn invalid_error(kind: u16, message: impl Into<String>) -> Error {
    Error::InvalidRecord {
        record_type: kind,
        message: message.into(),
    }
}

fn invalid<T>(kind: u16, message: impl Into<String>) -> Result<T> {
    Err(invalid_error(kind, message))
}
