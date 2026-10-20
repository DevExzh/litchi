//! Granular edits of a chart's typed definition.
//!
//! [`Edit`] is the draft a package chart edit hands out
//! ([`crate::Edit::definition_edit`]): its verbs stage axis, series, data
//! point, cached-table and style changes that the package commit publishes as
//! one chart replacement. [`Change`] records each change in the package
//! patch ([`crate::Patch::definition_changes`]).

use crate::{
    AxisSpec, CachedCell, CachedRow, CachedTable, DataPointSpec, Definition, Limits, PlotAreaSpec,
    SeriesSpec,
};
use litchi_core::{Error, Result};
use std::sync::Arc;

/// A style-bearing semantic chart site.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum StyleTarget {
    Chart,
    PlotArea,
    Axis(usize),
    Series(usize),
    DataPoint { series: usize, point: usize },
    Wall,
    Floor,
    StockGainMarker,
    StockLossMarker,
    StockRangeLine,
}

/// One granular semantic operation retained by a definition patch.
#[derive(Clone, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum Change {
    DefinitionUpdated,
    PlotAreaUpdated,
    AxisInserted {
        index: usize,
    },
    AxisUpdated {
        index: usize,
    },
    AxisRemoved {
        index: usize,
    },
    SeriesInserted {
        index: usize,
    },
    SeriesUpdated {
        index: usize,
    },
    SeriesRemoved {
        index: usize,
    },
    DataPointInserted {
        series: usize,
        index: usize,
    },
    DataPointUpdated {
        series: usize,
        index: usize,
    },
    DataPointRemoved {
        series: usize,
        index: usize,
    },
    CachedTableUpdated,
    CachedRowInserted {
        header: bool,
        index: usize,
    },
    CachedRowUpdated {
        header: bool,
        index: usize,
    },
    CachedRowRemoved {
        header: bool,
        index: usize,
    },
    CachedCellInserted {
        header: bool,
        row: usize,
        index: usize,
    },
    CachedCellUpdated {
        header: bool,
        row: usize,
        index: usize,
    },
    CachedCellRemoved {
        header: bool,
        row: usize,
        index: usize,
    },
    StyleUpdated {
        target: StyleTarget,
    },
}

/// An immutable validated chart definition and its retained limits.
#[derive(Clone, Debug)]
pub(crate) struct Snapshot {
    pub(crate) definition: Arc<Definition>,
    pub(crate) limits: Limits,
}

impl Snapshot {
    /// Validate and retain a typed definition with caller-selected limits.
    ///
    /// # Errors
    ///
    /// Returns an error when the definition violates semantic or size limits.
    pub(crate) fn new(definition: Definition, limits: Limits) -> Result<Self> {
        crate::validation::validate_definition(&definition, limits)?;
        Ok(Self {
            definition: Arc::new(definition),
            limits,
        })
    }

    #[must_use]
    pub(crate) fn definition(&self) -> &Definition {
        &self.definition
    }

    #[must_use]
    pub(crate) fn edit(&self) -> Edit {
        Edit {
            source: self.clone(),
            candidate: self.definition().clone(),
            changes: Vec::new(),
        }
    }
}

/// A source-bound granular chart-definition transaction.
pub struct Edit {
    source: Snapshot,
    candidate: Definition,
    changes: Vec<Change>,
}

impl Edit {
    #[must_use]
    pub fn definition(&self) -> &Definition {
        &self.candidate
    }

    /// Replace the plot area. ODF requires one plot area, so deletion is
    /// represented by [`Self::clear_plot_area`].
    pub fn update_plot_area(&mut self, plot: PlotAreaSpec) {
        self.candidate.plot_area = plot;
        self.changes.push(Change::PlotAreaUpdated);
    }

    /// Reset the required plot area to an empty typed value.
    pub fn clear_plot_area(&mut self) {
        self.update_plot_area(PlotAreaSpec::default());
    }

    /// Insert an axis at a zero-based position.
    ///
    /// # Errors
    ///
    /// Returns an error when the insertion selector is out of bounds.
    pub fn insert_axis(&mut self, index: usize, axis: AxisSpec) -> Result<()> {
        if index > self.candidate.plot_area.axes.len() {
            return selector("ODC axis insertion selector is out of bounds");
        }
        self.candidate.plot_area.axes.insert(index, axis);
        self.changes.push(Change::AxisInserted { index });
        Ok(())
    }

    /// Replace an axis at a zero-based position.
    ///
    /// # Errors
    ///
    /// Returns an error when the selector is out of bounds.
    pub fn update_axis(&mut self, index: usize, axis: AxisSpec) -> Result<()> {
        *self
            .candidate
            .plot_area
            .axes
            .get_mut(index)
            .ok_or_else(|| invalid("ODC axis selector is out of bounds"))? = axis;
        self.changes.push(Change::AxisUpdated { index });
        Ok(())
    }

    /// Remove and return an axis at a zero-based position.
    ///
    /// # Errors
    ///
    /// Returns an error when the selector is out of bounds.
    pub fn remove_axis(&mut self, index: usize) -> Result<AxisSpec> {
        if index >= self.candidate.plot_area.axes.len() {
            return selector("ODC axis selector is out of bounds");
        }
        self.changes.push(Change::AxisRemoved { index });
        Ok(self.candidate.plot_area.axes.remove(index))
    }

    /// Insert a series at a zero-based position.
    ///
    /// # Errors
    ///
    /// Returns an error when the insertion selector is out of bounds.
    pub fn insert_series(&mut self, index: usize, series: SeriesSpec) -> Result<()> {
        if index > self.candidate.plot_area.series.len() {
            return selector("ODC series insertion selector is out of bounds");
        }
        self.candidate.plot_area.series.insert(index, series);
        self.changes.push(Change::SeriesInserted { index });
        Ok(())
    }

    /// Replace a series at a zero-based position.
    ///
    /// # Errors
    ///
    /// Returns an error when the selector is out of bounds.
    pub fn update_series(&mut self, index: usize, series: SeriesSpec) -> Result<()> {
        *self
            .candidate
            .plot_area
            .series
            .get_mut(index)
            .ok_or_else(|| invalid("ODC series selector is out of bounds"))? = series;
        self.changes.push(Change::SeriesUpdated { index });
        Ok(())
    }

    /// Remove and return a series at a zero-based position.
    ///
    /// # Errors
    ///
    /// Returns an error when the selector is out of bounds.
    pub fn remove_series(&mut self, index: usize) -> Result<SeriesSpec> {
        if index >= self.candidate.plot_area.series.len() {
            return selector("ODC series selector is out of bounds");
        }
        self.changes.push(Change::SeriesRemoved { index });
        Ok(self.candidate.plot_area.series.remove(index))
    }

    /// Insert a data point into one series.
    ///
    /// # Errors
    ///
    /// Returns an error when either selector is out of bounds.
    pub fn insert_data_point(
        &mut self,
        series: usize,
        index: usize,
        point: DataPointSpec,
    ) -> Result<()> {
        let points = self.points_mut(series)?;
        if index > points.len() {
            return selector("ODC data-point insertion selector is out of bounds");
        }
        points.insert(index, point);
        self.changes
            .push(Change::DataPointInserted { series, index });
        Ok(())
    }

    /// Replace a data point in one series.
    ///
    /// # Errors
    ///
    /// Returns an error when either selector is out of bounds.
    pub fn update_data_point(
        &mut self,
        series: usize,
        index: usize,
        point: DataPointSpec,
    ) -> Result<()> {
        *self
            .points_mut(series)?
            .get_mut(index)
            .ok_or_else(|| invalid("ODC data-point selector is out of bounds"))? = point;
        self.changes
            .push(Change::DataPointUpdated { series, index });
        Ok(())
    }

    /// Remove and return a data point from one series.
    ///
    /// # Errors
    ///
    /// Returns an error when either selector is out of bounds.
    pub fn remove_data_point(&mut self, series: usize, index: usize) -> Result<DataPointSpec> {
        let points = self.points_mut(series)?;
        if index >= points.len() {
            return selector("ODC data-point selector is out of bounds");
        }
        let removed = points.remove(index);
        self.changes
            .push(Change::DataPointRemoved { series, index });
        Ok(removed)
    }

    pub fn set_cached_table(&mut self, table: Option<CachedTable>) {
        self.candidate.cached_table = table;
        self.changes.push(Change::CachedTableUpdated);
    }

    /// Insert a row into the header or body cached-table collection.
    ///
    /// # Errors
    ///
    /// Returns an error when the table is absent or the selector is invalid.
    pub fn insert_cached_row(&mut self, header: bool, index: usize, row: CachedRow) -> Result<()> {
        let rows = self.rows_mut(header)?;
        if index > rows.len() {
            return selector("ODC cached-row insertion selector is out of bounds");
        }
        rows.insert(index, row);
        self.changes
            .push(Change::CachedRowInserted { header, index });
        Ok(())
    }

    /// Replace a row in the header or body cached-table collection.
    ///
    /// # Errors
    ///
    /// Returns an error when the table is absent or the selector is invalid.
    pub fn update_cached_row(&mut self, header: bool, index: usize, row: CachedRow) -> Result<()> {
        *self
            .rows_mut(header)?
            .get_mut(index)
            .ok_or_else(|| invalid("ODC cached-row selector is out of bounds"))? = row;
        self.changes
            .push(Change::CachedRowUpdated { header, index });
        Ok(())
    }

    /// Remove a row from the header or body cached-table collection.
    ///
    /// # Errors
    ///
    /// Returns an error when the table is absent or the selector is invalid.
    pub fn remove_cached_row(&mut self, header: bool, index: usize) -> Result<CachedRow> {
        let rows = self.rows_mut(header)?;
        if index >= rows.len() {
            return selector("ODC cached-row selector is out of bounds");
        }
        let removed = rows.remove(index);
        self.changes
            .push(Change::CachedRowRemoved { header, index });
        Ok(removed)
    }

    /// Insert a cell into one cached row.
    ///
    /// # Errors
    ///
    /// Returns an error when the table, row, or insertion selector is invalid.
    pub fn insert_cached_cell(
        &mut self,
        header: bool,
        row: usize,
        index: usize,
        cell: CachedCell,
    ) -> Result<()> {
        let cells = self.cells_mut(header, row)?;
        if index > cells.len() {
            return selector("ODC cached-cell insertion selector is out of bounds");
        }
        cells.insert(index, cell);
        self.changes
            .push(Change::CachedCellInserted { header, row, index });
        Ok(())
    }

    /// Replace a cell in one cached row.
    ///
    /// # Errors
    ///
    /// Returns an error when the table, row, or cell selector is invalid.
    pub fn update_cached_cell(
        &mut self,
        header: bool,
        row: usize,
        index: usize,
        cell: CachedCell,
    ) -> Result<()> {
        *self
            .cells_mut(header, row)?
            .get_mut(index)
            .ok_or_else(|| invalid("ODC cached-cell selector is out of bounds"))? = cell;
        self.changes
            .push(Change::CachedCellUpdated { header, row, index });
        Ok(())
    }

    /// Remove and return a cell from one cached row.
    ///
    /// # Errors
    ///
    /// Returns an error when the table, row, or cell selector is invalid.
    pub fn remove_cached_cell(
        &mut self,
        header: bool,
        row: usize,
        index: usize,
    ) -> Result<CachedCell> {
        let cells = self.cells_mut(header, row)?;
        if index >= cells.len() {
            return selector("ODC cached-cell selector is out of bounds");
        }
        let removed = cells.remove(index);
        self.changes
            .push(Change::CachedCellRemoved { header, row, index });
        Ok(removed)
    }

    /// Set or remove `chart:style-name` at a typed semantic site.
    ///
    /// # Errors
    ///
    /// Returns an error when the selected style-bearing element is absent.
    pub fn set_style(&mut self, target: StyleTarget, style_name: Option<String>) -> Result<()> {
        let field = match target {
            StyleTarget::Chart => &mut self.candidate.style_name,
            StyleTarget::PlotArea => &mut self.candidate.plot_area.style_name,
            StyleTarget::Axis(index) => {
                &mut self
                    .candidate
                    .plot_area
                    .axes
                    .get_mut(index)
                    .ok_or_else(|| invalid("ODC style axis selector is out of bounds"))?
                    .style_name
            },
            StyleTarget::Series(index) => {
                &mut self
                    .candidate
                    .plot_area
                    .series
                    .get_mut(index)
                    .ok_or_else(|| invalid("ODC style series selector is out of bounds"))?
                    .style_name
            },
            StyleTarget::DataPoint { series, point } => {
                &mut self
                    .candidate
                    .plot_area
                    .series
                    .get_mut(series)
                    .ok_or_else(|| invalid("ODC style series selector is out of bounds"))?
                    .data_points
                    .get_mut(point)
                    .ok_or_else(|| invalid("ODC style data-point selector is out of bounds"))?
                    .style_name
            },
            StyleTarget::Wall => style_element(&mut self.candidate.plot_area.wall, "wall")?,
            StyleTarget::Floor => style_element(&mut self.candidate.plot_area.floor, "floor")?,
            StyleTarget::StockGainMarker => style_element(
                &mut self.candidate.plot_area.stock_gain_marker,
                "stock-gain-marker",
            )?,
            StyleTarget::StockLossMarker => style_element(
                &mut self.candidate.plot_area.stock_loss_marker,
                "stock-loss-marker",
            )?,
            StyleTarget::StockRangeLine => style_element(
                &mut self.candidate.plot_area.stock_range_line,
                "stock-range-line",
            )?,
        };
        *field = style_name;
        self.changes.push(Change::StyleUpdated { target });
        Ok(())
    }

    /// Validate and atomically publish this detached edit.
    ///
    /// # Errors
    ///
    /// Returns an error when the completed definition violates retained limits.
    pub(crate) fn commit(mut self) -> Result<Commit> {
        crate::validation::validate_definition(&self.candidate, self.source.limits)?;
        if self.candidate == *self.source.definition {
            self.changes.clear();
        }
        let target = Snapshot {
            definition: Arc::new(self.candidate),
            limits: self.source.limits,
        };
        Ok(Commit {
            snapshot: target.clone(),
            patch: Patch {
                source: self.source,
                target,
                changes: self.changes,
            },
        })
    }

    fn points_mut(&mut self, series: usize) -> Result<&mut Vec<DataPointSpec>> {
        self.candidate
            .plot_area
            .series
            .get_mut(series)
            .map(|value| &mut value.data_points)
            .ok_or_else(|| invalid("ODC data-point series selector is out of bounds"))
    }

    fn rows_mut(&mut self, header: bool) -> Result<&mut Vec<CachedRow>> {
        let table = self
            .candidate
            .cached_table
            .as_mut()
            .ok_or_else(|| invalid("ODC cached table is absent"))?;
        Ok(if header {
            &mut table.header_rows
        } else {
            &mut table.rows
        })
    }

    fn cells_mut(&mut self, header: bool, row: usize) -> Result<&mut Vec<CachedCell>> {
        self.rows_mut(header)?
            .get_mut(row)
            .map(|value| &mut value.cells)
            .ok_or_else(|| invalid("ODC cached-row selector is out of bounds"))
    }
}

/// A committed definition snapshot and its reversible patch.
pub(crate) struct Commit {
    snapshot: Snapshot,
    patch: Patch,
}

impl Commit {
    #[must_use]
    pub(crate) fn patch(&self) -> &Patch {
        &self.patch
    }

    #[must_use]
    pub(crate) fn into_snapshot(self) -> Snapshot {
        self.snapshot
    }
}

/// A semantic exact-source patch that supports inversion and composition.
#[derive(Clone, Debug)]
pub(crate) struct Patch {
    pub(crate) source: Snapshot,
    pub(crate) target: Snapshot,
    pub(crate) changes: Vec<Change>,
}

impl Patch {
    #[must_use]
    pub(crate) fn changes(&self) -> &[Change] {
        &self.changes
    }

    /// Join two patches that share the same semantic source.
    ///
    /// Inputs remain immutable. Independent paths merge; divergent edits to
    /// one path are reported as deterministic conflicts.
    #[must_use]
    pub(crate) fn join(&self, other: &Self) -> crate::merge::DefinitionMerge {
        crate::merge::join(self, other)
    }

    /// Transfer this patch's semantic dependencies onto another snapshot.
    ///
    /// The destination remains unchanged. Dangling axis references and other
    /// invalid combined dependencies are returned as `chart.dependencies`.
    #[must_use]
    pub(crate) fn transfer_to(&self, destination: &Snapshot) -> crate::merge::DefinitionMerge {
        crate::merge::transfer(self, destination)
    }
}

pub(crate) fn inverse_change(change: &Change) -> Change {
    match change {
        Change::DefinitionUpdated => Change::DefinitionUpdated,
        Change::AxisInserted { index } => Change::AxisRemoved { index: *index },
        Change::AxisRemoved { index } => Change::AxisInserted { index: *index },
        Change::SeriesInserted { index } => Change::SeriesRemoved { index: *index },
        Change::SeriesRemoved { index } => Change::SeriesInserted { index: *index },
        Change::DataPointInserted { series, index } => Change::DataPointRemoved {
            series: *series,
            index: *index,
        },
        Change::DataPointRemoved { series, index } => Change::DataPointInserted {
            series: *series,
            index: *index,
        },
        Change::CachedRowInserted { header, index } => Change::CachedRowRemoved {
            header: *header,
            index: *index,
        },
        Change::CachedRowRemoved { header, index } => Change::CachedRowInserted {
            header: *header,
            index: *index,
        },
        Change::CachedCellInserted { header, row, index } => Change::CachedCellRemoved {
            header: *header,
            row: *row,
            index: *index,
        },
        Change::CachedCellRemoved { header, row, index } => Change::CachedCellInserted {
            header: *header,
            row: *row,
            index: *index,
        },
        Change::PlotAreaUpdated => Change::PlotAreaUpdated,
        Change::AxisUpdated { index } => Change::AxisUpdated { index: *index },
        Change::SeriesUpdated { index } => Change::SeriesUpdated { index: *index },
        Change::DataPointUpdated { series, index } => Change::DataPointUpdated {
            series: *series,
            index: *index,
        },
        Change::CachedTableUpdated => Change::CachedTableUpdated,
        Change::CachedRowUpdated { header, index } => Change::CachedRowUpdated {
            header: *header,
            index: *index,
        },
        Change::CachedCellUpdated { header, row, index } => Change::CachedCellUpdated {
            header: *header,
            row: *row,
            index: *index,
        },
        Change::StyleUpdated { target } => Change::StyleUpdated { target: *target },
    }
}

fn style_element<'a>(
    element: &'a mut Option<crate::StyleElement>,
    name: &str,
) -> Result<&'a mut Option<String>> {
    element
        .as_mut()
        .map(|value| &mut value.style_name)
        .ok_or_else(|| invalid(format!("ODC {name} style target is absent")))
}

fn selector<T>(message: impl Into<String>) -> Result<T> {
    Err(invalid(message))
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}
