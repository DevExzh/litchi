//! Selector-first Pages body-chart value-axis settings (record 0422):
//! [`crate::chart::value_axis`].
//!
//! The bounds, step counts and scale of a chart's primary value axis are
//! edited as one aggregate on the shared rules and in-place write of
//! `litchi_iwa_style::value_axis` (the host's
//! `set_body_chart_value_axis_{bounds,steps,scale}`: the axis non-style
//! object is rewritten in place, and one another chart axis shares is
//! refused). Publication, candidate reopen and locality run on the shared
//! Pages chart-property core (`super::body_chart_series`).

use std::fmt;

use litchi_core::Position;
use litchi_iwa_archive::package::OwnedExactArtifacts as ExactArtifacts;
use litchi_iwa_style::value_axis as engine;

use super::Package;
use super::body_chart_series::{
    Diagnostics, Error, Ledger, MutationPlan, Refusal, Target, artifacts, authorizes, inverse_plan,
    noop_artifacts,
};
use crate::selector::BodyChartSelector;

/// Exact-source value-axis types.
pub(crate) mod value_axis {
    use super::{Diagnostics, ExactArtifacts, MutationPlan, Package, Position, fmt, inverse_plan};

    pub use litchi_iwa_common::chart::axis::{
        Bound, Bounds, MajorStepCount, MinorStepCount, Scale, Steps, ValueAxisSettings,
    };

    /// A compact exact-source patch of one chart's value-axis settings.
    #[derive(Clone, PartialEq)]
    pub struct Patch {
        pub(in super::super) artifacts: ExactArtifacts,
        pub(in super::super) chart: Position,
        pub(in super::super) before: ValueAxisSettings,
        pub(in super::super) after: ValueAxisSettings,
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
                before: self.after,
                after: self.before,
                plan: inverse_plan(&self.plan),
            }
        }

        /// Return the selected semantic chart position.
        #[must_use]
        pub const fn chart_position(&self) -> Position {
            self.chart
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

use value_axis::{Bounds, Scale, Steps, ValueAxisSettings};

/// One immutable selector-first value-axis edit: the settings read when the
/// edit began and the staged settings.
pub struct ValueAxisEdit<'a> {
    source: &'a Package,
    selector: BodyChartSelector,
    before: ValueAxisSettings,
    after: ValueAxisSettings,
}

impl fmt::Debug for ValueAxisEdit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("before", &self.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

fn read_settings(
    ledger: &mut Ledger,
    package: &Package,
    target: &Target,
) -> Result<ValueAxisSettings, Error> {
    ledger.read(package, |workspace, limits| {
        engine::read(workspace, target.chart.drawable_identifier.get(), limits)
    })
}

fn validate(settings: ValueAxisSettings) -> Result<ValueAxisSettings, Error> {
    engine::validate(settings).map_err(|_| Refusal::InvalidSettings.into())
}

impl ValueAxisEdit<'_> {
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
    pub fn set(mut self, settings: ValueAxisSettings) -> Result<Self, Error> {
        self.after = validate(settings)?;
        Ok(self)
    }

    /// Stage replacement bounds, keeping the staged steps and scale.
    ///
    /// # Errors
    ///
    /// As [`Self::set`].
    pub fn set_bounds(mut self, bounds: Bounds) -> Result<Self, Error> {
        self.after = validate(self.after.with_bounds(bounds))?;
        Ok(self)
    }

    /// Stage replacement step counts, keeping the staged bounds and scale.
    ///
    /// # Errors
    ///
    /// As [`Self::set`].
    pub fn set_steps(mut self, steps: Steps) -> Result<Self, Error> {
        self.after = validate(self.after.with_steps(steps))?;
        Ok(self)
    }

    /// Stage a replacement scale, keeping the staged bounds and steps.
    ///
    /// # Errors
    ///
    /// As [`Self::set`].
    pub fn set_scale(mut self, scale: Scale) -> Result<Self, Error> {
        self.after = validate(self.after.with_scale(scale))?;
        Ok(self)
    }

    /// Validate and atomically publish the staged settings. Settings equal
    /// to the chart's are an exact no-op.
    ///
    /// # Errors
    ///
    /// [`Refusal::SharedStyle`] (the axis object is shared with another
    /// chart axis), [`Refusal::UnsupportedAxis`] (no primary value axis);
    /// source, limit and verification errors.
    pub fn commit(self) -> Result<value_axis::Commit, Error> {
        let mut ledger = Ledger::new(self.source)?;
        let target = ledger.target(self.source, self.selector)?;
        let before = read_settings(&mut ledger, self.source, &target)?;
        if !engine::equal_exact(before, self.before) {
            return Err(Error::Verification);
        }
        let value = self.after;
        let published = if engine::equal_exact(before, value) {
            None
        } else {
            ledger.write(self.source, |workspace, limits| {
                engine::write(
                    workspace,
                    target.chart.drawable_identifier.get(),
                    value,
                    limits,
                )
            })?
        };
        let Some(published) = published else {
            return Ok(value_axis::Commit {
                package: self.source.clone(),
                patch: value_axis::Patch {
                    artifacts: noop_artifacts(self.source),
                    chart: target.chart.position,
                    before,
                    after: before,
                    plan: MutationPlan::default(),
                },
                diagnostics: Diagnostics::unchanged(),
            });
        };
        let candidate = ledger.target(&published.package, self.selector)?;
        if candidate.kind != target.kind
            || !engine::equal_exact(
                read_settings(&mut ledger, &published.package, &candidate)?,
                value,
            )
        {
            return Err(Error::Verification);
        }
        Ok(value_axis::Commit {
            patch: value_axis::Patch {
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
    /// The bounds, step counts and scale of one body chart's primary value
    /// axis.
    ///
    /// # Errors
    ///
    /// Selector errors, [`Refusal::UnsupportedAxis`], source and limit
    /// errors (stored settings the native format could not have written
    /// are [`Error::InvalidSource`]).
    pub fn body_chart_value_axis_settings(
        &self,
        selector: impl Into<BodyChartSelector>,
    ) -> Result<ValueAxisSettings, Error> {
        let mut ledger = Ledger::new(self)?;
        let target = ledger.target(self, selector.into())?;
        read_settings(&mut ledger, self, &target)
    }

    /// Start an exact-source value-axis edit of one body chart.
    ///
    /// ```no_run
    /// use litchi_pages::Package;
    /// use litchi_pages::chart::value_axis::{Bound, Bounds, Scale};
    ///
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let package = Package::open("input.pages")?;
    /// let bounds = Bounds::fixed(Bound::new(1.0)?, Bound::new(1000.0)?)?;
    /// let commit = package
    ///     .edit_body_chart_value_axis_settings(0)?
    ///     .set_bounds(bounds)?
    ///     .set_scale(Scale::Logarithmic)?
    ///     .commit()?;
    /// let restored = commit
    ///     .package()
    ///     .apply_body_chart_value_axis_settings(&commit.patch().inverse())?;
    /// # let _ = restored;
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// # Errors
    ///
    /// Selector errors, [`Refusal::UnsupportedAxis`], source and limit
    /// errors.
    pub fn edit_body_chart_value_axis_settings(
        &self,
        selector: impl Into<BodyChartSelector>,
    ) -> Result<ValueAxisEdit<'_>, Error> {
        let selector = selector.into();
        let mut ledger = Ledger::new(self)?;
        let target = ledger.target(self, selector)?;
        let before = read_settings(&mut ledger, self, &target)?;
        Ok(ValueAxisEdit {
            source: self,
            selector,
            before,
            after: before,
        })
    }

    /// Apply a value-axis patch only to its exact source package.
    ///
    /// # Errors
    ///
    /// [`Error::PatchConflict`] when this package is not the patch's exact
    /// source; locality and verification errors.
    pub fn apply_body_chart_value_axis_settings(
        &self,
        patch: &value_axis::Patch,
    ) -> Result<value_axis::Commit, Error> {
        let mut ledger = Ledger::new(self)?;
        if !authorizes(self, &patch.artifacts) {
            return Err(Error::PatchConflict);
        }
        if patch.is_noop() {
            return Ok(value_axis::Commit {
                package: self.clone(),
                patch: patch.clone(),
                diagnostics: Diagnostics::unchanged(),
            });
        }
        let candidate = ledger.reopen(self, &patch.artifacts, &patch.plan)?;
        let target = ledger.target(&candidate, BodyChartSelector::position(patch.chart))?;
        if !engine::equal_exact(
            read_settings(&mut ledger, &candidate, &target)?,
            patch.after,
        ) {
            return Err(Error::Verification);
        }
        Ok(value_axis::Commit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: Diagnostics::published(0),
        })
    }
}
