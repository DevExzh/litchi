//! Granular transactions over detached typed chart definitions.
#![allow(
    clippy::missing_errors_doc,
    clippy::module_name_repetitions,
    reason = "all selectors share the transaction's documented invalid-format refusal contract"
)]

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
pub enum DefinitionChange {
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
pub struct DefinitionSnapshot {
    definition: Arc<Definition>,
    limits: Limits,
}

impl DefinitionSnapshot {
    /// Validate and retain a typed definition with caller-selected limits.
    pub fn new(definition: Definition, limits: Limits) -> Result<Self> {
        crate::validation::validate_definition(&definition, limits)?;
        Ok(Self {
            definition: Arc::new(definition),
            limits,
        })
    }

    /// Validate with default limits.
    pub fn with_default_limits(definition: Definition) -> Result<Self> {
        Self::new(definition, Limits::default())
    }

    #[must_use]
    pub fn definition(&self) -> &Definition {
        &self.definition
    }

    #[must_use]
    pub const fn limits(&self) -> Limits {
        self.limits
    }

    #[must_use]
    pub fn edit(&self) -> DefinitionEdit {
        DefinitionEdit {
            source: self.clone(),
            candidate: self.definition().clone(),
            changes: Vec::new(),
        }
    }
}

/// A source-bound granular chart-definition transaction.
pub struct DefinitionEdit {
    source: DefinitionSnapshot,
    candidate: Definition,
    changes: Vec<DefinitionChange>,
}

impl DefinitionEdit {
    #[must_use]
    pub fn definition(&self) -> &Definition {
        &self.candidate
    }

    /// Replace the plot area. ODF requires one plot area, so deletion is
    /// represented by [`Self::clear_plot_area`].
    pub fn update_plot_area(&mut self, plot: PlotAreaSpec) {
        self.candidate.plot_area = plot;
        self.changes.push(DefinitionChange::PlotAreaUpdated);
    }

    /// Reset the required plot area to an empty typed value.
    pub fn clear_plot_area(&mut self) {
        self.update_plot_area(PlotAreaSpec::default());
    }

    pub fn insert_axis(&mut self, index: usize, axis: AxisSpec) -> Result<()> {
        if index > self.candidate.plot_area.axes.len() {
            return selector("ODC axis insertion selector is out of bounds");
        }
        self.candidate.plot_area.axes.insert(index, axis);
        self.changes.push(DefinitionChange::AxisInserted { index });
        Ok(())
    }

    pub fn update_axis(&mut self, index: usize, axis: AxisSpec) -> Result<()> {
        *self
            .candidate
            .plot_area
            .axes
            .get_mut(index)
            .ok_or_else(|| invalid("ODC axis selector is out of bounds"))? = axis;
        self.changes.push(DefinitionChange::AxisUpdated { index });
        Ok(())
    }

    pub fn remove_axis(&mut self, index: usize) -> Result<AxisSpec> {
        if index >= self.candidate.plot_area.axes.len() {
            return selector("ODC axis selector is out of bounds");
        }
        self.changes.push(DefinitionChange::AxisRemoved { index });
        Ok(self.candidate.plot_area.axes.remove(index))
    }

    pub fn insert_series(&mut self, index: usize, series: SeriesSpec) -> Result<()> {
        if index > self.candidate.plot_area.series.len() {
            return selector("ODC series insertion selector is out of bounds");
        }
        self.candidate.plot_area.series.insert(index, series);
        self.changes
            .push(DefinitionChange::SeriesInserted { index });
        Ok(())
    }

    pub fn update_series(&mut self, index: usize, series: SeriesSpec) -> Result<()> {
        *self
            .candidate
            .plot_area
            .series
            .get_mut(index)
            .ok_or_else(|| invalid("ODC series selector is out of bounds"))? = series;
        self.changes.push(DefinitionChange::SeriesUpdated { index });
        Ok(())
    }

    pub fn remove_series(&mut self, index: usize) -> Result<SeriesSpec> {
        if index >= self.candidate.plot_area.series.len() {
            return selector("ODC series selector is out of bounds");
        }
        self.changes.push(DefinitionChange::SeriesRemoved { index });
        Ok(self.candidate.plot_area.series.remove(index))
    }

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
            .push(DefinitionChange::DataPointInserted { series, index });
        Ok(())
    }

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
            .push(DefinitionChange::DataPointUpdated { series, index });
        Ok(())
    }

    pub fn remove_data_point(&mut self, series: usize, index: usize) -> Result<DataPointSpec> {
        let points = self.points_mut(series)?;
        if index >= points.len() {
            return selector("ODC data-point selector is out of bounds");
        }
        let removed = points.remove(index);
        self.changes
            .push(DefinitionChange::DataPointRemoved { series, index });
        Ok(removed)
    }

    pub fn set_cached_table(&mut self, table: Option<CachedTable>) {
        self.candidate.cached_table = table;
        self.changes.push(DefinitionChange::CachedTableUpdated);
    }

    pub fn insert_cached_row(&mut self, header: bool, index: usize, row: CachedRow) -> Result<()> {
        let rows = self.rows_mut(header)?;
        if index > rows.len() {
            return selector("ODC cached-row insertion selector is out of bounds");
        }
        rows.insert(index, row);
        self.changes
            .push(DefinitionChange::CachedRowInserted { header, index });
        Ok(())
    }

    pub fn update_cached_row(&mut self, header: bool, index: usize, row: CachedRow) -> Result<()> {
        *self
            .rows_mut(header)?
            .get_mut(index)
            .ok_or_else(|| invalid("ODC cached-row selector is out of bounds"))? = row;
        self.changes
            .push(DefinitionChange::CachedRowUpdated { header, index });
        Ok(())
    }

    pub fn remove_cached_row(&mut self, header: bool, index: usize) -> Result<CachedRow> {
        let rows = self.rows_mut(header)?;
        if index >= rows.len() {
            return selector("ODC cached-row selector is out of bounds");
        }
        let removed = rows.remove(index);
        self.changes
            .push(DefinitionChange::CachedRowRemoved { header, index });
        Ok(removed)
    }

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
            .push(DefinitionChange::CachedCellInserted { header, row, index });
        Ok(())
    }

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
            .push(DefinitionChange::CachedCellUpdated { header, row, index });
        Ok(())
    }

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
            .push(DefinitionChange::CachedCellRemoved { header, row, index });
        Ok(removed)
    }

    /// Set or remove `chart:style-name` at a typed semantic site.
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
        self.changes.push(DefinitionChange::StyleUpdated { target });
        Ok(())
    }

    /// Validate and atomically publish this detached edit.
    pub fn commit(mut self) -> Result<DefinitionCommit> {
        crate::validation::validate_definition(&self.candidate, self.source.limits)?;
        if self.candidate == *self.source.definition {
            self.changes.clear();
        }
        let target = DefinitionSnapshot {
            definition: Arc::new(self.candidate),
            limits: self.source.limits,
        };
        Ok(DefinitionCommit {
            snapshot: target.clone(),
            patch: DefinitionPatch {
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

fn style_element<'a>(
    element: &'a mut Option<crate::StyleElement>,
    name: &str,
) -> Result<&'a mut Option<String>> {
    element
        .as_mut()
        .map(|value| &mut value.style_name)
        .ok_or_else(|| invalid(format!("ODC {name} style target is absent")))
}

/// A committed definition snapshot and its reversible patch.
pub struct DefinitionCommit {
    snapshot: DefinitionSnapshot,
    patch: DefinitionPatch,
}

impl DefinitionCommit {
    #[must_use]
    pub fn changed(&self) -> bool {
        !self.patch.changes.is_empty()
    }

    #[must_use]
    pub fn snapshot(&self) -> &DefinitionSnapshot {
        &self.snapshot
    }

    #[must_use]
    pub fn patch(&self) -> &DefinitionPatch {
        &self.patch
    }

    #[must_use]
    pub fn into_snapshot(self) -> DefinitionSnapshot {
        self.snapshot
    }
}

/// A semantic exact-source patch that supports inversion and composition.
#[derive(Clone, Debug)]
pub struct DefinitionPatch {
    source: DefinitionSnapshot,
    target: DefinitionSnapshot,
    changes: Vec<DefinitionChange>,
}

impl DefinitionPatch {
    #[must_use]
    pub fn changes(&self) -> &[DefinitionChange] {
        &self.changes
    }

    #[must_use]
    pub fn is_applicable_to(&self, source: &DefinitionSnapshot) -> bool {
        self.source.limits == source.limits && *self.source.definition == *source.definition
    }

    pub fn apply(&self, source: &DefinitionSnapshot) -> Result<DefinitionSnapshot> {
        if !self.is_applicable_to(source) {
            return selector("ODC definition patch source does not match");
        }
        Ok(self.target.clone())
    }

    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            source: self.target.clone(),
            target: self.source.clone(),
            changes: self.changes.iter().rev().map(inverse_change).collect(),
        }
    }

    /// Compose two contiguous patches into one exact-source patch.
    pub fn compose(&self, next: &Self) -> Result<Self> {
        if self.target.limits != next.source.limits
            || *self.target.definition != *next.source.definition
        {
            return selector("ODC definition patches are not contiguous");
        }
        let mut changes = self.changes.clone();
        changes.extend_from_slice(&next.changes);
        Ok(Self {
            source: self.source.clone(),
            target: next.target.clone(),
            changes,
        })
    }
}

/// Bounded undo/redo history for validated definition patches.
pub struct DefinitionHistory {
    current: DefinitionSnapshot,
    undo: Vec<DefinitionPatch>,
    redo: Vec<DefinitionPatch>,
}

impl DefinitionHistory {
    #[must_use]
    pub fn new(snapshot: DefinitionSnapshot) -> Self {
        Self {
            current: snapshot,
            undo: Vec::new(),
            redo: Vec::new(),
        }
    }

    #[must_use]
    pub fn current(&self) -> &DefinitionSnapshot {
        &self.current
    }

    /// Record a contiguous committed patch and discard the redo branch.
    pub fn record(&mut self, patch: &DefinitionPatch) -> Result<()> {
        if !patch.is_applicable_to(&self.current) {
            return selector("ODC history patch is not contiguous");
        }
        if patch.changes.is_empty() {
            return Ok(());
        }
        if self.undo.len() >= self.current.limits.max_history() {
            return selector("ODC history exceeds the caller-selected limit");
        }
        self.current = patch.target.clone();
        self.undo.push(patch.clone());
        self.redo.clear();
        Ok(())
    }

    pub fn undo(&mut self) -> Result<bool> {
        let Some(patch) = self.undo.pop() else {
            return Ok(false);
        };
        let inverse = patch.inverse();
        self.current = inverse.apply(&self.current)?;
        self.redo.push(patch);
        Ok(true)
    }

    pub fn redo(&mut self) -> Result<bool> {
        let Some(patch) = self.redo.pop() else {
            return Ok(false);
        };
        self.current = patch.apply(&self.current)?;
        self.undo.push(patch);
        Ok(true)
    }

    #[must_use]
    pub fn undo_len(&self) -> usize {
        self.undo.len()
    }

    #[must_use]
    pub fn redo_len(&self) -> usize {
        self.redo.len()
    }
}

fn inverse_change(change: &DefinitionChange) -> DefinitionChange {
    match change {
        DefinitionChange::AxisInserted { index } => DefinitionChange::AxisRemoved { index: *index },
        DefinitionChange::AxisRemoved { index } => DefinitionChange::AxisInserted { index: *index },
        DefinitionChange::SeriesInserted { index } => {
            DefinitionChange::SeriesRemoved { index: *index }
        },
        DefinitionChange::SeriesRemoved { index } => {
            DefinitionChange::SeriesInserted { index: *index }
        },
        DefinitionChange::DataPointInserted { series, index } => {
            DefinitionChange::DataPointRemoved {
                series: *series,
                index: *index,
            }
        },
        DefinitionChange::DataPointRemoved { series, index } => {
            DefinitionChange::DataPointInserted {
                series: *series,
                index: *index,
            }
        },
        DefinitionChange::CachedRowInserted { header, index } => {
            DefinitionChange::CachedRowRemoved {
                header: *header,
                index: *index,
            }
        },
        DefinitionChange::CachedRowRemoved { header, index } => {
            DefinitionChange::CachedRowInserted {
                header: *header,
                index: *index,
            }
        },
        DefinitionChange::CachedCellInserted { header, row, index } => {
            DefinitionChange::CachedCellRemoved {
                header: *header,
                row: *row,
                index: *index,
            }
        },
        DefinitionChange::CachedCellRemoved { header, row, index } => {
            DefinitionChange::CachedCellInserted {
                header: *header,
                row: *row,
                index: *index,
            }
        },
        DefinitionChange::PlotAreaUpdated => DefinitionChange::PlotAreaUpdated,
        DefinitionChange::AxisUpdated { index } => DefinitionChange::AxisUpdated { index: *index },
        DefinitionChange::SeriesUpdated { index } => {
            DefinitionChange::SeriesUpdated { index: *index }
        },
        DefinitionChange::DataPointUpdated { series, index } => {
            DefinitionChange::DataPointUpdated {
                series: *series,
                index: *index,
            }
        },
        DefinitionChange::CachedTableUpdated => DefinitionChange::CachedTableUpdated,
        DefinitionChange::CachedRowUpdated { header, index } => {
            DefinitionChange::CachedRowUpdated {
                header: *header,
                index: *index,
            }
        },
        DefinitionChange::CachedCellUpdated { header, row, index } => {
            DefinitionChange::CachedCellUpdated {
                header: *header,
                row: *row,
                index: *index,
            }
        },
        DefinitionChange::StyleUpdated { target } => {
            DefinitionChange::StyleUpdated { target: *target }
        },
    }
}

fn selector<T>(message: impl Into<String>) -> Result<T> {
    Err(invalid(message))
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}
