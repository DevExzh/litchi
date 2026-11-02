//! Implementation of the public [`crate::sheet::chart::legend_style`]
//! transaction: the legend fill, border and shadow of one sheet chart on the
//! shared legend-style variation graph (`litchi_iwa_style::legend_style`,
//! record 0414). A write derives a private legend-style variation when the
//! chart's legend style is shared or not a variation, writes the value and
//! collapses the variation when it equals its parent again; the changes are
//! published through the chart-series workspace path (record 0378a).

use std::fmt;

use litchi_iwa_archive::package::OwnedExactArtifacts;
use litchi_iwa_style::legend_style as engine;

use super::Package;
use super::chart_arrangement::ChartBudget;
use super::sheet_chart_series::{
    Error, Refusal, map_arrangement, prepare_components, publish_changes_exact, target,
};
use super::sheet_drawable_graph as graph;
use super::structure;
use crate::selector::{ChartSelector, SheetSelector};

pub use engine::{
    LegendFill, LegendFont, LegendFontSize, LegendShadow, LegendStroke, Property, PropertyKind,
};

/// Compact publication diagnostics.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Diagnostics {
    changed: bool,
}

impl Diagnostics {
    /// Whether the package changed.
    #[must_use]
    pub const fn changed(&self) -> bool {
        self.changed
    }
}

/// An exact-source reversible legend-style patch.
#[derive(Clone, PartialEq)]
pub struct Patch {
    artifacts: OwnedExactArtifacts,
    sheet: usize,
    chart: ChartSelector,
    before: Property,
    after: Property,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("sheet", &self.sheet)
            .field("chart", &self.chart)
            .field("before", &self.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl Patch {
    /// Whether the patch is an exact no-op.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.artifacts.is_byte_noop() && self.before == self.after
    }

    /// The exact inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            sheet: self.sheet,
            chart: self.chart,
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

    /// The source value.
    #[must_use]
    pub const fn before(&self) -> &Property {
        &self.before
    }

    /// The target value.
    #[must_use]
    pub const fn after(&self) -> &Property {
        &self.after
    }
}

/// One verified legend-style commit.
#[derive(Debug)]
pub struct Commit {
    package: Package,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
    /// Borrow the committed package.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Consume the commit and return its package.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Borrow the exact patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Borrow the publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }

    /// Consume the commit and return its package and patch.
    #[must_use]
    pub fn into_parts(self) -> (Package, Patch) {
        (self.package, self.patch)
    }
}

/// Legend-style changes staged against one chart of an immutable package.
pub struct Edit<'a> {
    source: &'a Package,
    sheet: usize,
    chart: ChartSelector,
    staged: Option<Property>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("sheet", &self.sheet)
            .field("chart", &self.chart)
            .field("staged", &self.staged)
            .finish_non_exhaustive()
    }
}

fn read(
    package: &Package,
    sheet: usize,
    chart: ChartSelector,
    kind: PropertyKind,
    budget: &mut ChartBudget,
) -> Result<Property, Error> {
    let target = target(package, sheet, chart, budget)?;
    let prepared = prepare_components(package, budget)?;
    let workspace = prepared.workspace()?;
    let limits = budget.series_limits().map_err(map_arrangement)?;
    let (value, report) = engine::read(&workspace, target.chart, kind, limits)?;
    budget
        .charge_series_report(report)
        .map_err(map_arrangement)?;
    Ok(value)
}

impl Edit<'_> {
    /// Stage `property` as the legend's value, replacing any staged one.
    ///
    /// # Errors
    ///
    /// None today; the signature matches the other chart families.
    pub fn set(&mut self, property: Property) -> Result<&mut Self, Error> {
        self.staged = Some(property);
        Ok(self)
    }

    /// Stage removing the legend's own value of `kind` (inheriting again).
    ///
    /// # Errors
    ///
    /// None today; the signature matches the other chart families.
    pub fn reset(&mut self, kind: PropertyKind) -> Result<&mut Self, Error> {
        self.staged = Some(kind.inherited());
        Ok(self)
    }

    /// Validate and atomically publish the staged value. A value equal to
    /// the legend's is an exact no-op.
    ///
    /// # Errors
    ///
    /// [`Refusal::NoStagedOperation`], [`Refusal::MissingStyle`],
    /// [`Refusal::ImageFill`], source, limit and verification errors.
    pub fn commit(self) -> Result<Commit, Error> {
        let value = self.staged.ok_or(Refusal::NoStagedOperation)?;
        let source = self.source;
        let mut budget = ChartBudget::for_package(source).map_err(map_arrangement)?;
        let before = read(source, self.sheet, self.chart, value.kind(), &mut budget)?;
        let owner = structure::source_owner(source)?;
        if before == value {
            return Ok(Commit {
                package: source.clone(),
                patch: Patch {
                    artifacts: OwnedExactArtifacts::new(owner.clone(), owner),
                    sheet: self.sheet,
                    chart: self.chart,
                    before: before.clone(),
                    after: before,
                },
                diagnostics: Diagnostics { changed: false },
            });
        }
        let selected = target(source, self.sheet, self.chart, &mut budget)?;
        let prepared = prepare_components(source, &mut budget)?;
        let mut workspace = prepared.workspace()?;
        let limits = budget.series_limits().map_err(map_arrangement)?;
        let (changed, report) = engine::write(&mut workspace, selected.chart, &value, limits)?;
        budget
            .charge_series_report(report)
            .map_err(map_arrangement)?;
        let finish_limits = budget.finish_limits().map_err(map_arrangement)?;
        workspace.finish_requirements().map_err(Error::from)?;
        let (changes, finish_report) = workspace
            .finish_bounded(finish_limits)
            .map_err(Error::from)?;
        budget
            .charge_finish_report(finish_report)
            .map_err(map_arrangement)?;
        if !changed || changes.is_empty() {
            return Err(Error::Verification);
        }
        let published = publish_changes_exact(source, &changes, &mut budget)?;
        let after = read(
            &published.package,
            self.sheet,
            self.chart,
            value.kind(),
            &mut budget,
        )?;
        if after != value {
            return Err(Error::Verification);
        }
        let target_owner = structure::source_owner(&published.package)?;
        Ok(Commit {
            package: published.package,
            patch: Patch {
                artifacts: OwnedExactArtifacts::new(owner, target_owner),
                sheet: self.sheet,
                chart: self.chart,
                before,
                after,
            },
            diagnostics: Diagnostics { changed: true },
        })
    }
}

fn resolve_sheet(package: &Package, selector: SheetSelector<'_>) -> Result<usize, Error> {
    graph::sheet_position(package, selector)?.ok_or_else(|| Refusal::SheetNotFound.into())
}

impl Package {
    /// The legend's own value of one legend-style property of one sheet
    /// chart.
    ///
    /// # Errors
    ///
    /// Selector errors, [`Refusal::MissingStyle`] and source errors.
    pub fn sheet_chart_legend_style<'sheet>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        chart: impl Into<ChartSelector>,
        kind: PropertyKind,
    ) -> Result<Property, Error> {
        let sheet = resolve_sheet(self, sheet.into())?;
        let mut budget = ChartBudget::for_package(self).map_err(map_arrangement)?;
        read(self, sheet, chart.into(), kind, &mut budget)
    }

    /// Start an immutable legend-style edit of one sheet chart.
    ///
    /// # Errors
    ///
    /// Selector errors.
    pub fn edit_sheet_chart_legend_style<'sheet>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        chart: impl Into<ChartSelector>,
    ) -> Result<Edit<'_>, Error> {
        let sheet = resolve_sheet(self, sheet.into())?;
        let chart = chart.into();
        let mut budget = ChartBudget::for_package(self).map_err(map_arrangement)?;
        target(self, sheet, chart, &mut budget)?;
        Ok(Edit {
            source: self,
            sheet,
            chart,
            staged: None,
        })
    }

    /// Apply a legend-style patch to its exact source package.
    ///
    /// # Errors
    ///
    /// [`Error::PatchConflict`] when this package is not the patch's exact
    /// source; verification errors.
    pub fn apply_sheet_chart_legend_style(&self, patch: &Patch) -> Result<Commit, Error> {
        let owner = structure::source_owner(self)?;
        if !patch.artifacts.authorizes_owner(&owner) {
            return Err(Error::PatchConflict);
        }
        if patch.is_noop() {
            return Ok(Commit {
                package: self.clone(),
                patch: patch.clone(),
                diagnostics: Diagnostics { changed: false },
            });
        }
        let mut budget = ChartBudget::for_package(self).map_err(map_arrangement)?;
        let target_bytes = patch.artifacts.target_owner().as_ref().len();
        budget
            .candidate_reopen(target_bytes)
            .map_err(map_arrangement)?;
        let candidate = Package::from_source_owner_with_options(
            patch.artifacts.target_owner(),
            self.state.options,
        )
        .map_err(|_error| Error::Verification)?;
        if read(
            &candidate,
            patch.sheet,
            patch.chart,
            patch.after.kind(),
            &mut budget,
        )? != patch.after
        {
            return Err(Error::Verification);
        }
        Ok(Commit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: Diagnostics { changed: true },
        })
    }
}
