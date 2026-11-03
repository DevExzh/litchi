//! Implementation of the public [`crate::sheet::chart::reference_lines`]
//! transaction: the ordered reference lines of one sheet chart's primary
//! value axis on the shared reference-line graph
//! (`litchi_iwa_style::reference_line`, records 0375a and 0419). A write
//! patches retained lines in place, removes surplus ones and allocates new
//! ones with fresh UUIDs, as the host's `set_sheet_chart_reference_lines`
//! does; the changes are published through the chart-series workspace path
//! (record 0378a).

use std::fmt;

use litchi_iwa_archive::package::OwnedExactArtifacts;
use litchi_iwa_style::reference_line as engine;

use super::Package;
use super::chart_arrangement::ChartBudget;
use super::sheet_chart_series::{
    Error, Refusal, map_arrangement, prepare_components, publish_changes_exact, target,
};
use super::sheet_drawable_graph as graph;
use super::structure;
use crate::selector::{ChartSelector, SheetSelector};

pub use litchi_iwa_common::chart::reference_line::{Kind, Line, Unknown, Value};

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

/// An exact-source reversible reference-line patch.
#[derive(Clone, PartialEq)]
pub struct Patch {
    artifacts: OwnedExactArtifacts,
    sheet: usize,
    chart: ChartSelector,
    before: Vec<Line>,
    after: Vec<Line>,
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

    /// The source lines.
    #[must_use]
    pub fn before(&self) -> &[Line] {
        &self.before
    }

    /// The target lines.
    #[must_use]
    pub fn after(&self) -> &[Line] {
        &self.after
    }
}

/// One verified reference-line commit.
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

/// Reference-line changes staged against one chart of an immutable package.
pub struct Edit<'a> {
    source: &'a Package,
    sheet: usize,
    chart: ChartSelector,
    staged: Option<Vec<Line>>,
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
    budget: &mut ChartBudget,
) -> Result<Vec<Line>, Error> {
    let target = target(package, sheet, chart, budget)?;
    let prepared = prepare_components(package, budget)?;
    let workspace = prepared.workspace()?;
    let limits = budget.series_limits().map_err(map_arrangement)?;
    let (value, report) = engine::read_lines(&workspace, target.chart, limits)?;
    budget
        .charge_series_report(report)
        .map_err(map_arrangement)?;
    Ok(value)
}

impl Edit<'_> {
    /// Stage `lines` as the chart's reference lines, in order (at most
    /// five, one of each kind; an empty list removes them all), replacing
    /// any staged list.
    ///
    /// # Errors
    ///
    /// None today; the signature matches the other chart families.
    pub fn set(&mut self, lines: Vec<Line>) -> Result<&mut Self, Error> {
        self.staged = Some(lines);
        Ok(self)
    }

    /// Validate and atomically publish the staged lines. Lines equal to the
    /// chart's are an exact no-op.
    ///
    /// # Errors
    ///
    /// [`Refusal::NoStagedOperation`]; more than five lines, a repeated
    /// kind or a line another chart lists ([`Error::InvalidSource`]);
    /// source, limit and verification errors.
    pub fn commit(self) -> Result<Commit, Error> {
        let value = self.staged.ok_or(Refusal::NoStagedOperation)?;
        let source = self.source;
        let mut budget = ChartBudget::for_package(source).map_err(map_arrangement)?;
        let before = read(source, self.sheet, self.chart, &mut budget)?;
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
        let (changed, report) =
            engine::write_lines(&mut workspace, selected.chart, &value, limits)?;
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
    /// The ordered reference lines of one sheet chart's primary value axis.
    ///
    /// # Errors
    ///
    /// Selector errors and source errors.
    pub fn sheet_chart_reference_lines<'sheet>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        chart: impl Into<ChartSelector>,
    ) -> Result<Vec<Line>, Error> {
        let sheet = resolve_sheet(self, sheet.into())?;
        let mut budget = ChartBudget::for_package(self).map_err(map_arrangement)?;
        read(self, sheet, chart.into(), &mut budget)
    }

    /// Start an immutable reference-line edit of one sheet chart.
    ///
    /// ```no_run
    /// use litchi_numbers::Package;
    /// use litchi_numbers::sheet::chart::reference_lines::Line;
    ///
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let package = Package::open("input.numbers")?;
    /// let mut edit = package.edit_sheet_chart_reference_lines(0, 0)?;
    /// edit.set(vec![Line::average(), Line::maximum()])?;
    /// let commit = edit.commit()?;
    /// let restored = commit
    ///     .package()
    ///     .apply_sheet_chart_reference_lines(&commit.patch().inverse())?;
    /// # let _ = restored;
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// # Errors
    ///
    /// Selector errors.
    pub fn edit_sheet_chart_reference_lines<'sheet>(
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

    /// Apply a reference-line patch to its exact source package.
    ///
    /// # Errors
    ///
    /// [`Error::PatchConflict`] when this package is not the patch's exact
    /// source; verification errors.
    pub fn apply_sheet_chart_reference_lines(&self, patch: &Patch) -> Result<Commit, Error> {
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
        if read(&candidate, patch.sheet, patch.chart, &mut budget)? != patch.after {
            return Err(Error::Verification);
        }
        Ok(Commit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: Diagnostics { changed: true },
        })
    }
}
