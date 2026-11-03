//! Implementation of the public [`crate::sheet::chart::value_axis`]
//! transaction: the bounds, step counts and scale of one sheet chart's
//! primary value axis, edited as one aggregate on the shared rules and
//! in-place write of `litchi_iwa_style::value_axis` (record 0422). As the
//! host's `set_sheet_chart_value_axis_{bounds,steps,scale}` do, the axis
//! non-style object is rewritten in place and one another chart axis shares
//! is refused; the change is published through the chart-series workspace
//! path (record 0378a).

use std::fmt;

use litchi_iwa_archive::package::OwnedExactArtifacts;
use litchi_iwa_style::value_axis as engine;

use super::Package;
use super::chart_arrangement::ChartBudget;
use super::sheet_chart_series::{
    Error, Refusal, map_arrangement, prepare_components, publish_changes_exact, target,
};
use super::sheet_drawable_graph as graph;
use super::structure;
use crate::selector::{ChartSelector, SheetSelector};

pub use litchi_iwa_common::chart::axis::{
    Bound, Bounds, MajorStepCount, MinorStepCount, Scale, Steps, ValueAxisSettings,
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

/// An exact-source reversible value-axis patch.
#[derive(Clone, PartialEq)]
pub struct Patch {
    artifacts: OwnedExactArtifacts,
    sheet: usize,
    chart: ChartSelector,
    before: ValueAxisSettings,
    after: ValueAxisSettings,
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
        self.artifacts.is_byte_noop() && engine::equal_exact(self.before, self.after)
    }

    /// The exact inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            sheet: self.sheet,
            chart: self.chart,
            before: self.after,
            after: self.before,
        }
    }

    /// The source settings.
    #[must_use]
    pub const fn before(&self) -> ValueAxisSettings {
        self.before
    }

    /// The target settings.
    #[must_use]
    pub const fn after(&self) -> ValueAxisSettings {
        self.after
    }
}

/// One verified value-axis commit.
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

/// Value-axis settings staged against one chart of an immutable package:
/// the settings read when the edit began and the staged settings.
pub struct Edit<'a> {
    source: &'a Package,
    sheet: usize,
    chart: ChartSelector,
    before: ValueAxisSettings,
    after: ValueAxisSettings,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("sheet", &self.sheet)
            .field("chart", &self.chart)
            .field("before", &self.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

fn read(
    package: &Package,
    sheet: usize,
    chart: ChartSelector,
    budget: &mut ChartBudget,
) -> Result<ValueAxisSettings, Error> {
    let target = target(package, sheet, chart, budget)?;
    let prepared = prepare_components(package, budget)?;
    let workspace = prepared.workspace()?;
    let limits = budget.series_limits().map_err(map_arrangement)?;
    let (value, report) = engine::read(&workspace, target.chart, limits)?;
    budget
        .charge_series_report(report)
        .map_err(map_arrangement)?;
    Ok(value)
}

fn validate(settings: ValueAxisSettings) -> Result<ValueAxisSettings, Error> {
    engine::validate(settings).map_err(|_| Refusal::InvalidSettings.into())
}

impl Edit<'_> {
    /// The settings read when the edit began.
    #[must_use]
    pub const fn before(&self) -> ValueAxisSettings {
        self.before
    }

    /// The staged settings.
    #[must_use]
    pub const fn after(&self) -> ValueAxisSettings {
        self.after
    }

    /// Stage all value-axis settings.
    ///
    /// # Errors
    ///
    /// [`Refusal::InvalidSettings`]: a non-finite endpoint, an inverted
    /// range, or a non-positive endpoint on a logarithmic axis.
    pub fn set(&mut self, settings: ValueAxisSettings) -> Result<&mut Self, Error> {
        self.after = validate(settings)?;
        Ok(self)
    }

    /// Stage replacement bounds, keeping the staged steps and scale.
    ///
    /// # Errors
    ///
    /// As [`Self::set`].
    pub fn set_bounds(&mut self, bounds: Bounds) -> Result<&mut Self, Error> {
        self.after = validate(self.after.with_bounds(bounds))?;
        Ok(self)
    }

    /// Stage replacement step counts, keeping the staged bounds and scale.
    ///
    /// # Errors
    ///
    /// As [`Self::set`].
    pub fn set_steps(&mut self, steps: Steps) -> Result<&mut Self, Error> {
        self.after = validate(self.after.with_steps(steps))?;
        Ok(self)
    }

    /// Stage a replacement scale, keeping the staged bounds and steps.
    ///
    /// # Errors
    ///
    /// As [`Self::set`].
    pub fn set_scale(&mut self, scale: Scale) -> Result<&mut Self, Error> {
        self.after = validate(self.after.with_scale(scale))?;
        Ok(self)
    }

    /// Validate and atomically publish the staged settings. Settings equal
    /// to the chart's are an exact no-op.
    ///
    /// # Errors
    ///
    /// [`Refusal::SharedStyle`] (the axis object is shared with another
    /// chart axis); source, limit and verification errors.
    pub fn commit(self) -> Result<Commit, Error> {
        let value = self.after;
        let source = self.source;
        let mut budget = ChartBudget::for_package(source).map_err(map_arrangement)?;
        let before = read(source, self.sheet, self.chart, &mut budget)?;
        if !engine::equal_exact(before, self.before) {
            return Err(Error::Verification);
        }
        let owner = structure::source_owner(source)?;
        if engine::equal_exact(before, value) {
            return Ok(Commit {
                package: source.clone(),
                patch: Patch {
                    artifacts: OwnedExactArtifacts::new(owner.clone(), owner),
                    sheet: self.sheet,
                    chart: self.chart,
                    before,
                    after: before,
                },
                diagnostics: Diagnostics { changed: false },
            });
        }
        let selected = target(source, self.sheet, self.chart, &mut budget)?;
        let prepared = prepare_components(source, &mut budget)?;
        let mut workspace = prepared.workspace()?;
        let limits = budget.series_limits().map_err(map_arrangement)?;
        let (changed, report) = engine::write(&mut workspace, selected.chart, value, limits)?;
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
        let after = read(&published.package, self.sheet, self.chart, &mut budget)?;
        if !engine::equal_exact(after, value) {
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
    /// The bounds, step counts and scale of one sheet chart's primary value
    /// axis.
    ///
    /// # Errors
    ///
    /// Selector errors and source errors (stored settings the native format
    /// could not have written are [`Error::InvalidSource`]).
    pub fn sheet_chart_value_axis_settings<'sheet>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        chart: impl Into<ChartSelector>,
    ) -> Result<ValueAxisSettings, Error> {
        let sheet = resolve_sheet(self, sheet.into())?;
        let mut budget = ChartBudget::for_package(self).map_err(map_arrangement)?;
        read(self, sheet, chart.into(), &mut budget)
    }

    /// Start an immutable value-axis edit of one sheet chart.
    ///
    /// ```no_run
    /// use litchi_numbers::Package;
    /// use litchi_numbers::sheet::chart::value_axis::{Bound, Bounds, Scale};
    ///
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let package = Package::open("input.numbers")?;
    /// let mut edit = package.edit_sheet_chart_value_axis_settings(0, 0)?;
    /// edit.set_bounds(Bounds::fixed(Bound::new(1.0)?, Bound::new(1000.0)?)?)?;
    /// edit.set_scale(Scale::Logarithmic)?;
    /// let commit = edit.commit()?;
    /// let restored = commit
    ///     .package()
    ///     .apply_sheet_chart_value_axis_settings(&commit.patch().inverse())?;
    /// # let _ = restored;
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// # Errors
    ///
    /// Selector errors and source errors.
    pub fn edit_sheet_chart_value_axis_settings<'sheet>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        chart: impl Into<ChartSelector>,
    ) -> Result<Edit<'_>, Error> {
        let sheet = resolve_sheet(self, sheet.into())?;
        let chart = chart.into();
        let mut budget = ChartBudget::for_package(self).map_err(map_arrangement)?;
        let before = read(self, sheet, chart, &mut budget)?;
        Ok(Edit {
            source: self,
            sheet,
            chart,
            before,
            after: before,
        })
    }

    /// Apply a value-axis patch to its exact source package.
    ///
    /// # Errors
    ///
    /// [`Error::PatchConflict`] when this package is not the patch's exact
    /// source; verification errors.
    pub fn apply_sheet_chart_value_axis_settings(&self, patch: &Patch) -> Result<Commit, Error> {
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
        if !engine::equal_exact(
            read(&candidate, patch.sheet, patch.chart, &mut budget)?,
            patch.after,
        ) {
            return Err(Error::Verification);
        }
        Ok(Commit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: Diagnostics { changed: true },
        })
    }
}
