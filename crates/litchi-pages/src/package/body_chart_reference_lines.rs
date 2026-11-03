//! Selector-first Pages body-chart reference lines (record 0419):
//! [`crate::chart::reference_lines`].
//!
//! The shared reference-line graph (`litchi_iwa_style::reference_line`)
//! replaces the ordered lines of the chart's primary value axis as the
//! host's `set_body_chart_reference_lines` does: retained lines are patched
//! in place (a line another chart lists is refused), surplus lines are
//! removed and new ones allocated with fresh UUIDs. Publication, candidate
//! reopen and locality run on the shared Pages chart-property core
//! (`super::body_chart_series`).

use std::fmt;

use litchi_core::Position;
use litchi_iwa_archive::package::OwnedExactArtifacts as ExactArtifacts;
use litchi_iwa_style::reference_line as engine;

use super::Package;
use super::body_chart_series::{
    Diagnostics, Error, Ledger, MutationPlan, Refusal, Target, artifacts, authorizes, inverse_plan,
    noop_artifacts,
};
use crate::selector::BodyChartSelector;

/// Exact-source reference-line types.
pub(crate) mod reference_lines {
    use super::{Diagnostics, ExactArtifacts, MutationPlan, Package, Position, fmt, inverse_plan};

    pub use litchi_iwa_common::chart::reference_line::{Kind, Line, Unknown, Value};

    /// A compact exact-source patch of one chart's reference lines.
    #[derive(Clone, PartialEq)]
    pub struct Patch {
        pub(in super::super) artifacts: ExactArtifacts,
        pub(in super::super) chart: Position,
        pub(in super::super) before: Vec<Line>,
        pub(in super::super) after: Vec<Line>,
        pub(in super::super) plan: MutationPlan,
    }

    impl fmt::Debug for Patch {
        fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
            formatter
                .debug_struct("Patch")
                .field("chart", &self.chart)
                .field("before", &self.before)
                .field("after", &self.after)
                .finish_non_exhaustive()
        }
    }

    impl Patch {
        /// Whether the patch is an exact byte no-op.
        #[must_use]
        pub fn is_noop(&self) -> bool {
            self.artifacts.is_byte_noop()
        }

        /// Return the exact inverse patch.
        #[must_use]
        pub fn inverse(&self) -> Self {
            Self {
                artifacts: self.artifacts.inverse(),
                chart: self.chart,
                before: self.after.clone(),
                after: self.before.clone(),
                plan: inverse_plan(&self.plan),
            }
        }

        /// Return the selected semantic chart position.
        #[must_use]
        pub const fn chart_position(&self) -> Position {
            self.chart
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
        pub(in super::super) package: Package,
        pub(in super::super) patch: Patch,
        pub(in super::super) diagnostics: Diagnostics,
    }

    impl Commit {
        /// Borrow the candidate package.
        #[must_use]
        pub const fn package(&self) -> &Package {
            &self.package
        }

        /// Borrow the exact patch.
        #[must_use]
        pub const fn patch(&self) -> &Patch {
            &self.patch
        }

        /// Borrow compact publication diagnostics.
        #[must_use]
        pub const fn diagnostics(&self) -> &Diagnostics {
            &self.diagnostics
        }

        /// Consume the commit and return its package and exact patch.
        #[must_use]
        pub fn into_parts(self) -> (Package, Patch) {
            (self.package, self.patch)
        }

        /// Consume this commit.
        #[must_use]
        pub fn into_package(self) -> Package {
            self.package
        }
    }
}

use reference_lines::Line;

/// One immutable selector-first reference-line edit.
pub struct ReferenceLinesEdit<'a> {
    source: &'a Package,
    selector: BodyChartSelector,
    staged: Option<Vec<Line>>,
}

impl fmt::Debug for ReferenceLinesEdit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("staged", &self.staged)
            .finish_non_exhaustive()
    }
}

fn read_lines(ledger: &mut Ledger, package: &Package, target: &Target) -> Result<Vec<Line>, Error> {
    ledger.read(package, |workspace, limits| {
        engine::read_lines(workspace, target.chart.drawable_identifier.get(), limits)
    })
}

impl ReferenceLinesEdit<'_> {
    /// Stage `lines` as the chart's reference lines, in order (at most
    /// five, one of each kind; an empty list removes them all), replacing
    /// any staged list.
    ///
    /// # Errors
    ///
    /// None today; the signature matches the other chart families.
    pub fn set(mut self, lines: Vec<Line>) -> Result<Self, Error> {
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
    pub fn commit(self) -> Result<reference_lines::Commit, Error> {
        let value = self.staged.ok_or(Refusal::NoStagedOperation)?;
        let mut ledger = Ledger::new(self.source)?;
        let target = ledger.target(self.source, self.selector)?;
        let before = read_lines(&mut ledger, self.source, &target)?;
        let published = if before == value {
            None
        } else {
            ledger.write(self.source, |workspace, limits| {
                engine::write_lines(
                    workspace,
                    target.chart.drawable_identifier.get(),
                    &value,
                    limits,
                )
            })?
        };
        let Some(published) = published else {
            return Ok(reference_lines::Commit {
                package: self.source.clone(),
                patch: reference_lines::Patch {
                    artifacts: noop_artifacts(self.source),
                    chart: target.chart.position,
                    before: before.clone(),
                    after: before,
                    plan: MutationPlan::default(),
                },
                diagnostics: Diagnostics::unchanged(),
            });
        };
        let candidate = ledger.target(&published.package, self.selector)?;
        if candidate.kind != target.kind
            || read_lines(&mut ledger, &published.package, &candidate)? != value
        {
            return Err(Error::Verification);
        }
        Ok(reference_lines::Commit {
            patch: reference_lines::Patch {
                artifacts: artifacts(self.source, &published.package),
                chart: target.chart.position,
                before,
                after: value,
                plan: published.plan,
            },
            package: published.package,
            diagnostics: Diagnostics::published(published.touched),
        })
    }
}

impl Package {
    /// The ordered reference lines of one body chart's primary value axis.
    ///
    /// # Errors
    ///
    /// Selector errors, source and limit errors.
    pub fn body_chart_reference_lines(
        &self,
        selector: impl Into<BodyChartSelector>,
    ) -> Result<Vec<Line>, Error> {
        let mut ledger = Ledger::new(self)?;
        let target = ledger.target(self, selector.into())?;
        read_lines(&mut ledger, self, &target)
    }

    /// Start an exact-source reference-line edit of one body chart.
    ///
    /// ```no_run
    /// use litchi_pages::Package;
    /// use litchi_pages::chart::reference_lines::Line;
    ///
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let package = Package::open("input.pages")?;
    /// let commit = package
    ///     .edit_body_chart_reference_lines(0)?
    ///     .set(vec![Line::average()])?
    ///     .commit()?;
    /// let restored = commit
    ///     .package()
    ///     .apply_body_chart_reference_lines(&commit.patch().inverse())?;
    /// # let _ = restored;
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// # Errors
    ///
    /// Selector ceiling errors.
    pub fn edit_body_chart_reference_lines(
        &self,
        selector: impl Into<BodyChartSelector>,
    ) -> Result<ReferenceLinesEdit<'_>, Error> {
        let selector = selector.into();
        Ledger::new(self)?.target(self, selector)?;
        Ok(ReferenceLinesEdit {
            source: self,
            selector,
            staged: None,
        })
    }

    /// Apply a reference-line patch only to its exact source package.
    ///
    /// # Errors
    ///
    /// [`Error::PatchConflict`] when this package is not the patch's exact
    /// source; locality and verification errors.
    pub fn apply_body_chart_reference_lines(
        &self,
        patch: &reference_lines::Patch,
    ) -> Result<reference_lines::Commit, Error> {
        let mut ledger = Ledger::new(self)?;
        if !authorizes(self, &patch.artifacts) {
            return Err(Error::PatchConflict);
        }
        if patch.is_noop() {
            return Ok(reference_lines::Commit {
                package: self.clone(),
                patch: patch.clone(),
                diagnostics: Diagnostics::unchanged(),
            });
        }
        let candidate = ledger.reopen(self, &patch.artifacts, &patch.plan)?;
        let target = ledger.target(&candidate, BodyChartSelector::position(patch.chart))?;
        if read_lines(&mut ledger, &candidate, &target)? != patch.after {
            return Err(Error::Verification);
        }
        Ok(reference_lines::Commit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: Diagnostics::published(0),
        })
    }
}
