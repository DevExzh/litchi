//! Invariants and bounded derived state for the chart aggregate.

use super::super::super::{Kind, cache as chart_cache};
use super::super::cache::Cache;
use super::super::series::RowCol;
use crate::{Error, Result};

pub(in crate::chart) fn cache_dimensions(
    values: &[Cache],
    kind: Kind,
) -> Result<chart_cache::Dims> {
    match kind {
        Kind::Excel => {
            let mut bounds: Option<(u16, u16, u8, u8)> = None;
            for value in values {
                let Cache::Excel { row, col, .. } = value else {
                    return Err(Error::InvalidModel {
                        field: "cache",
                        reason: "Graph cache cell appears in an Excel chart",
                    });
                };
                bounds = Some(match bounds {
                    Some((first_row, last_row, first_col, last_col)) => (
                        first_row.min(*row),
                        last_row.max(*row),
                        first_col.min(*col),
                        last_col.max(*col),
                    ),
                    None => (*row, *row, *col, *col),
                });
            }
            let dimensions = match bounds {
                Some((first_row, last_row, first_col, last_col)) => chart_cache::ExcelDims::new(
                    u32::from(first_row),
                    u32::from(last_row) + 1,
                    u16::from(first_col),
                    u16::from(last_col) + 1,
                )
                .ok_or(Error::InvalidModel {
                    field: "Dimensions",
                    reason: "Excel cache bounds are outside the BIFF8 grid",
                })?,
                None => chart_cache::ExcelDims::default(),
            };
            Ok(chart_cache::Dims::Excel(dimensions))
        },
        Kind::Graph => {
            let mut coordinates = Vec::new();
            coordinates
                .try_reserve_exact(values.len())
                .map_err(|_| Error::Allocation {
                    resource: "Graph cache coordinates",
                })?;
            for value in values {
                let Cache::Graph { row, col, .. } = value else {
                    return Err(Error::InvalidModel {
                        field: "cache",
                        reason: "Excel cache cell appears in a Graph chart",
                    });
                };
                coordinates.push((u32::from(row.get()) << 12) | u32::from(col.get()));
            }
            coordinates.sort_unstable();
            coordinates.dedup();

            let mut current_row = None;
            let mut width = 0u16;
            let mut longest = 0u16;
            let mut rows = 0u16;
            for coordinate in coordinates {
                let row = coordinate >> 12;
                if current_row != Some(row) {
                    longest = longest.max(width);
                    width = 0;
                    rows = rows.checked_add(1).ok_or(Error::SizeOverflow {
                        resource: "Graph cache rows",
                    })?;
                    current_row = Some(row);
                }
                width = width.checked_add(1).ok_or(Error::SizeOverflow {
                    resource: "Graph cache row width",
                })?;
            }
            longest = longest.max(width);
            let rows = u8::try_from(rows).map_err(|_| Error::InvalidModel {
                field: "Dimensions",
                reason: "Graph cache has more than 255 non-empty rows",
            })?;
            let longest = RowCol::new(longest).ok_or(Error::InvalidModel {
                field: "Dimensions",
                reason: "Graph cache row has more than 3,999 cells",
            })?;
            let dimensions =
                chart_cache::GraphDims::new(longest, rows).ok_or(Error::InvalidModel {
                    field: "Dimensions",
                    reason: "Graph cache dimensions are inconsistent",
                })?;
            Ok(chart_cache::Dims::Graph(dimensions))
        },
    }
}

pub(in crate::chart) const fn dimensions_cover(
    declared: chart_cache::Dims,
    derived: chart_cache::Dims,
) -> bool {
    match (declared, derived) {
        (chart_cache::Dims::Excel(declared), chart_cache::Dims::Excel(derived)) => {
            if derived.row_after() == 0 {
                declared.row_after() == 0
            } else {
                declared.row_after() != 0
                    && declared.first_row() <= derived.first_row()
                    && declared.row_after() >= derived.row_after()
                    && declared.first_col() <= derived.first_col()
                    && declared.col_after() >= derived.col_after()
            }
        },
        (chart_cache::Dims::Graph(declared), chart_cache::Dims::Graph(derived)) => {
            declared.longest_row().get() == derived.longest_row().get()
                && declared.rows() == derived.rows()
        },
        _ => false,
    }
}

pub(super) fn reserve_one<T>(values: &mut Vec<T>, resource: &'static str) -> Result<()> {
    values
        .try_reserve(1)
        .map_err(|_| Error::Allocation { resource })
}

pub(super) fn check_add(current: usize, maximum: usize, resource: &'static str) -> Result<()> {
    let observed = current
        .checked_add(1)
        .ok_or(Error::SizeOverflow { resource })?;
    if observed > maximum {
        return Err(Error::LimitExceeded {
            resource,
            observed: crate::limits::as_u64(observed),
            maximum: crate::limits::as_u64(maximum),
        });
    }
    Ok(())
}
