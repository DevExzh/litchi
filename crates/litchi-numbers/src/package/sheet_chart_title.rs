//! Selector-first Numbers sheet-chart title facade.
//!
//! Resolution, bounded wire work and exact publication live in
//! [`super::sheet_chart_text`].  This module owns only the semantic title
//! entry points so the public surface does not expose native chart objects or
//! generated extension values.

pub use super::sheet_chart_text::{Commit, Diagnostics, Edit, Error, LimitKind, Patch, Refusal};

use super::Package;
use crate::selector::{ChartSelector, SheetSelector};

impl Package {
    /// Read the visible title of one selected sheet chart.
    pub fn sheet_chart_title<'sheet>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        chart: impl Into<ChartSelector>,
    ) -> Result<Option<String>, Error> {
        super::sheet_chart_text::read(self, sheet, chart, super::sheet_chart_text::Kind::Title)
    }

    /// Begin an immutable exact-source edit of one selected sheet-chart title.
    pub fn edit_sheet_chart_title<'sheet>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        chart: impl Into<ChartSelector>,
    ) -> Result<Edit<'_>, Error> {
        super::sheet_chart_text::edit(self, sheet, chart, super::sheet_chart_text::Kind::Title)
    }

    /// Apply an exact-source title patch to this package.
    pub fn apply_sheet_chart_title(&self, patch: &Patch) -> Result<Commit, Error> {
        super::sheet_chart_text::apply_patch(self, patch)
    }
}
