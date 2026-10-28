//! Selector-first Numbers sheet-chart axis-title facade.
//!
//! The category/value axis wire engine is shared with Keynote through
//! [`litchi_iwa_style::chart_text`].  This module keeps only selector and
//! semantic-axis entry points at the Numbers boundary.

pub use super::sheet_chart_text::{Commit, Diagnostics, Edit, Error, LimitKind, Patch, Refusal};
pub use litchi_iwa_common::chart::axis::Axis;

use super::Package;
use crate::selector::{ChartSelector, SheetSelector};

impl Package {
    /// Read the visible title of one selected chart axis.
    pub fn sheet_chart_axis_title<'sheet>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        chart: impl Into<ChartSelector>,
        axis: Axis,
    ) -> Result<Option<String>, Error> {
        super::sheet_chart_text::read(
            self,
            sheet,
            chart,
            super::sheet_chart_text::Kind::Axis(axis),
        )
    }

    /// Begin an immutable exact-source edit of one selected chart axis title.
    pub fn edit_sheet_chart_axis_title<'sheet>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        chart: impl Into<ChartSelector>,
        axis: Axis,
    ) -> Result<Edit<'_>, Error> {
        super::sheet_chart_text::edit(
            self,
            sheet,
            chart,
            super::sheet_chart_text::Kind::Axis(axis),
        )
    }

    /// Apply an exact-source axis-title patch to this package.
    pub fn apply_sheet_chart_axis_title(&self, patch: &Patch) -> Result<Commit, Error> {
        super::sheet_chart_text::apply_patch(self, patch)
    }
}
