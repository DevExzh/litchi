//! Selector-first Keynote slide-chart chart-level and axis-level style
//! transactions (record 0414): [`crate::chart::style`] and the axis-style
//! half of [`crate::chart::axis`].
//!
//! The shared chart style engine (`litchi_iwa_style::chart_style`) resolves
//! the chart's style object, refuses one another chart shares (the host's
//! `ensure_exclusive`) and rewrites it in place; the legend frame rewrites
//! the chart drawable. Publication, candidate reopen and locality run on
//! the shared Keynote chart-property core (`super::slide_chart_series`).
//! The host editor's `set_slide_chart_{pie_start_angle, donut_inner_radius,
//! radar_start_angle, 3d_series_gap, 3d_rotation, 3d_depth,
//! rounded_corners, legend_frame, border_visible, 3d_bar_shape,
//! radar_grid_shape, gap_spacing}` and `set_slide_chart_axis_*` routes
//! collapse into the two families' `Edit::set`.

use std::fmt;

use litchi_core::Position;
use litchi_iwa_archive::package::ExactArtifacts;
use litchi_iwa_common::chart::axis::Axis;
use litchi_iwa_style::chart_style::{self as engine, axis as axis_engine, chart as chart_engine};
use litchi_iwa_style::legend_style as legend_engine;

use super::Package;
use super::slide_chart_font::MutationPlan;
use super::slide_chart_series::{
    Diagnostics, Error, Ledger, Refusal, Target, artifacts, authorizes, noop_artifacts,
    owned_selectors,
};
use crate::{ChartSelector, SlideSelector};

/// The chart-drawable extension field of the legend style reference.
const LEGEND_STYLE_FIELDS: &[u32] = &[11];

/// The chart-drawable extension field of the legend frame.
const LEGEND_FRAME_FIELDS: &[u32] = &[3];

const fn chart_fields(kind: chart_engine::PropertyKind) -> &'static [u32] {
    match kind.archive() {
        None => LEGEND_FRAME_FIELDS,
        Some(_) => &[],
    }
}

/// Exact-source chart-level style types.
pub(crate) mod style {
    use super::{Diagnostics, ExactArtifacts, MutationPlan, Package, Position, chart_engine, fmt};

    pub use chart_engine::{Property, PropertyKind, RadarGridShape};

    /// A compact exact-source patch of one chart-level style property.
    #[derive(Clone, PartialEq)]
    pub struct Patch {
        pub(in super::super) artifacts: ExactArtifacts,
        pub(in super::super) slide: Position,
        pub(in super::super) chart: Position,
        pub(in super::super) before: Property,
        pub(in super::super) after: Property,
        pub(in super::super) plan: MutationPlan,
    }

    impl fmt::Debug for Patch {
        fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
            formatter
                .debug_struct("Patch")
                .field("slide", &self.slide)
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
                slide: self.slide,
                chart: self.chart,
                before: self.after,
                after: self.before,
                plan: super::super::slide_chart_font::inverse_mutation_plan(&self.plan),
            }
        }

        /// Return the selected semantic slide position.
        #[must_use]
        pub const fn slide_position(&self) -> Position {
            self.slide
        }

        /// Return the selected semantic chart position.
        #[must_use]
        pub const fn chart_position(&self) -> Position {
            self.chart
        }

        /// The source value.
        #[must_use]
        pub const fn before(&self) -> Property {
            self.before
        }

        /// The target value.
        #[must_use]
        pub const fn after(&self) -> Property {
            self.after
        }
    }

    /// One verified chart-level style commit.
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

/// Exact-source legend-style types.
pub(crate) mod legend_style {
    use super::{Diagnostics, ExactArtifacts, MutationPlan, Package, Position, fmt, legend_engine};

    pub use legend_engine::{LegendFill, LegendShadow, LegendStroke, Property, PropertyKind};

    /// A compact exact-source patch of one legend-style property.
    #[derive(Clone, PartialEq)]
    pub struct Patch {
        pub(in super::super) artifacts: ExactArtifacts,
        pub(in super::super) slide: Position,
        pub(in super::super) chart: Position,
        pub(in super::super) before: Property,
        pub(in super::super) after: Property,
        pub(in super::super) plan: MutationPlan,
    }

    impl fmt::Debug for Patch {
        fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
            formatter
                .debug_struct("Patch")
                .field("slide", &self.slide)
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
                slide: self.slide,
                chart: self.chart,
                before: self.after.clone(),
                after: self.before.clone(),
                plan: super::super::slide_chart_font::inverse_mutation_plan(&self.plan),
            }
        }

        /// Return the selected semantic slide position.
        #[must_use]
        pub const fn slide_position(&self) -> Position {
            self.slide
        }

        /// Return the selected semantic chart position.
        #[must_use]
        pub const fn chart_position(&self) -> Position {
            self.chart
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

/// Exact-source axis-style types.
pub(crate) mod axis {
    use super::{
        Axis, Diagnostics, ExactArtifacts, MutationPlan, Package, Position, axis_engine, fmt,
    };

    pub use axis_engine::{GridlineStroke, Property, PropertyKind};

    /// A compact exact-source patch of one axis-style property.
    #[derive(Clone, PartialEq)]
    pub struct Patch {
        pub(in super::super) artifacts: ExactArtifacts,
        pub(in super::super) slide: Position,
        pub(in super::super) chart: Position,
        pub(in super::super) axis: Axis,
        pub(in super::super) before: Property,
        pub(in super::super) after: Property,
        pub(in super::super) plan: MutationPlan,
    }

    impl fmt::Debug for Patch {
        fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
            formatter
                .debug_struct("Patch")
                .field("slide", &self.slide)
                .field("chart", &self.chart)
                .field("axis", &self.axis)
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
                slide: self.slide,
                chart: self.chart,
                axis: self.axis,
                before: self.after.clone(),
                after: self.before.clone(),
                plan: super::super::slide_chart_font::inverse_mutation_plan(&self.plan),
            }
        }

        /// Return the selected semantic slide position.
        #[must_use]
        pub const fn slide_position(&self) -> Position {
            self.slide
        }

        /// Return the selected semantic chart position.
        #[must_use]
        pub const fn chart_position(&self) -> Position {
            self.chart
        }

        /// The selected axis.
        #[must_use]
        pub const fn axis(&self) -> Axis {
            self.axis
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

    /// One verified axis-style commit.
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

/// One immutable selector-first chart-level style edit.
pub struct StyleEdit<'a> {
    source: &'a Package,
    slide: super::edit::OwnedSlideSelector,
    chart: super::edit::OwnedChartSelector,
    staged: Option<style::Property>,
}

impl fmt::Debug for StyleEdit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("staged", &self.staged)
            .finish_non_exhaustive()
    }
}

fn read_style(
    ledger: &mut Ledger,
    package: &Package,
    target: &Target,
    kind: style::PropertyKind,
) -> Result<style::Property, Error> {
    ledger.read(package, |workspace, limits| {
        engine::read_chart(
            workspace,
            target.selection.identifier,
            target.kind,
            kind,
            limits,
        )
    })
}

impl StyleEdit<'_> {
    /// Stage `property` as the chart's value, replacing any staged one.
    ///
    /// # Errors
    ///
    /// None today; the signature matches the other chart families.
    pub fn set(&mut self, property: style::Property) -> Result<&mut Self, Error> {
        self.staged = Some(property);
        Ok(self)
    }

    /// Stage resetting `kind` to its native default.
    ///
    /// # Errors
    ///
    /// None today; the signature matches the other chart families.
    pub fn reset(&mut self, kind: style::PropertyKind) -> Result<&mut Self, Error> {
        self.staged = Some(kind.native_default());
        Ok(self)
    }

    /// Validate and atomically publish the staged value. A value equal to
    /// the chart's is an exact no-op.
    ///
    /// # Errors
    ///
    /// [`Refusal::NoStagedOperation`], [`Refusal::UnsupportedKind`],
    /// [`Refusal::MissingStyle`], [`Refusal::SharedStyle`], source, limit
    /// and verification errors.
    pub fn commit(self) -> Result<style::Commit, Error> {
        let value = self.staged.ok_or(Refusal::NoStagedOperation)?;
        let mut ledger = Ledger::new(self.source)?;
        let target = ledger.target(
            self.source,
            self.slide.as_selector(),
            self.chart.as_selector(),
        )?;
        let before = read_style(&mut ledger, self.source, &target, value.kind())?;
        let published = if before == value {
            None
        } else {
            ledger.write(
                self.source,
                &target,
                chart_fields(value.kind()),
                |workspace, limits| {
                    engine::write_chart(
                        workspace,
                        target.selection.identifier,
                        target.kind,
                        value,
                        limits,
                    )
                },
            )?
        };
        let Some(published) = published else {
            return Ok(style::Commit {
                package: self.source.snapshot(),
                patch: style::Patch {
                    artifacts: noop_artifacts(self.source)?,
                    slide: target.selection.slide,
                    chart: target.selection.chart,
                    before,
                    after: before,
                    plan: MutationPlan::default(),
                },
                diagnostics: Diagnostics::unchanged(),
            });
        };
        let candidate = ledger.target(
            &published.package,
            self.slide.as_selector(),
            self.chart.as_selector(),
        )?;
        if candidate.kind != target.kind
            || read_style(&mut ledger, &published.package, &candidate, value.kind())? != value
        {
            return Err(Error::Verification);
        }
        Ok(style::Commit {
            patch: style::Patch {
                artifacts: artifacts(self.source, &published.package)?,
                slide: target.selection.slide,
                chart: target.selection.chart,
                before,
                after: value,
                plan: published.plan,
            },
            package: published.package,
            diagnostics: Diagnostics::published(published.touched),
        })
    }
}

/// One immutable selector-first legend-style edit.
pub struct LegendStyleEdit<'a> {
    source: &'a Package,
    slide: super::edit::OwnedSlideSelector,
    chart: super::edit::OwnedChartSelector,
    staged: Option<legend_style::Property>,
}

impl fmt::Debug for LegendStyleEdit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("staged", &self.staged)
            .finish_non_exhaustive()
    }
}

fn read_legend_style(
    ledger: &mut Ledger,
    package: &Package,
    target: &Target,
    kind: legend_style::PropertyKind,
) -> Result<legend_style::Property, Error> {
    ledger.read(package, |workspace, limits| {
        legend_engine::read(workspace, target.selection.identifier, kind, limits)
    })
}

impl LegendStyleEdit<'_> {
    /// Stage `property` as the chart's value, replacing any staged one.
    ///
    /// # Errors
    ///
    /// None today; the signature matches the other chart families.
    pub fn set(&mut self, property: legend_style::Property) -> Result<&mut Self, Error> {
        self.staged = Some(property);
        Ok(self)
    }

    /// Stage removing the legend's own value of `kind` (inheriting again).
    ///
    /// # Errors
    ///
    /// None today; the signature matches the other chart families.
    pub fn reset(&mut self, kind: legend_style::PropertyKind) -> Result<&mut Self, Error> {
        self.staged = Some(kind.inherited());
        Ok(self)
    }

    /// Validate and atomically publish the staged value. A value equal to
    /// the chart's is an exact no-op.
    ///
    /// # Errors
    ///
    /// [`Refusal::NoStagedOperation`], [`Refusal::MissingStyle`],
    /// [`Refusal::ImageFill`], source, limit and verification errors.
    pub fn commit(self) -> Result<legend_style::Commit, Error> {
        let value = self.staged.ok_or(Refusal::NoStagedOperation)?;
        let mut ledger = Ledger::new(self.source)?;
        let target = ledger.target(
            self.source,
            self.slide.as_selector(),
            self.chart.as_selector(),
        )?;
        let before = read_legend_style(&mut ledger, self.source, &target, value.kind())?;
        let published = if before == value {
            None
        } else {
            ledger.write(
                self.source,
                &target,
                LEGEND_STYLE_FIELDS,
                |workspace, limits| {
                    legend_engine::write(workspace, target.selection.identifier, &value, limits)
                },
            )?
        };
        let Some(published) = published else {
            return Ok(legend_style::Commit {
                package: self.source.snapshot(),
                patch: legend_style::Patch {
                    artifacts: noop_artifacts(self.source)?,
                    slide: target.selection.slide,
                    chart: target.selection.chart,
                    before: before.clone(),
                    after: before,
                    plan: MutationPlan::default(),
                },
                diagnostics: Diagnostics::unchanged(),
            });
        };
        let candidate = ledger.target(
            &published.package,
            self.slide.as_selector(),
            self.chart.as_selector(),
        )?;
        if candidate.kind != target.kind
            || read_legend_style(&mut ledger, &published.package, &candidate, value.kind())?
                != value
        {
            return Err(Error::Verification);
        }
        Ok(legend_style::Commit {
            patch: legend_style::Patch {
                artifacts: artifacts(self.source, &published.package)?,
                slide: target.selection.slide,
                chart: target.selection.chart,
                before,
                after: value,
                plan: published.plan,
            },
            package: published.package,
            diagnostics: Diagnostics::published(published.touched),
        })
    }
}

/// One immutable selector-first axis-style edit.
pub struct AxisEdit<'a> {
    source: &'a Package,
    slide: super::edit::OwnedSlideSelector,
    chart: super::edit::OwnedChartSelector,
    staged: Option<(Axis, axis::Property)>,
}

impl fmt::Debug for AxisEdit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("staged", &self.staged)
            .finish_non_exhaustive()
    }
}

fn read_axis(
    ledger: &mut Ledger,
    package: &Package,
    target: &Target,
    axis: Axis,
    kind: axis::PropertyKind,
) -> Result<axis::Property, Error> {
    ledger.read(package, |workspace, limits| {
        engine::read_axis(
            workspace,
            target.selection.identifier,
            target.kind,
            axis,
            kind,
            limits,
        )
    })
}

impl AxisEdit<'_> {
    /// Stage `property` on the chart's primary `axis`, replacing any staged
    /// value.
    ///
    /// # Errors
    ///
    /// [`Refusal::UnsupportedAxis`] for a property the axis does not have.
    pub fn set(&mut self, axis: Axis, property: axis::Property) -> Result<&mut Self, Error> {
        if property.kind().field(axis).is_none() {
            return Err(Refusal::UnsupportedAxis.into());
        }
        self.staged = Some((axis, property));
        Ok(self)
    }

    /// Validate and atomically publish the staged value. A value equal to
    /// the axis's is an exact no-op.
    ///
    /// # Errors
    ///
    /// [`Refusal::NoStagedOperation`], [`Refusal::UnsupportedAxis`],
    /// [`Refusal::SharedStyle`], source, limit and verification errors.
    pub fn commit(self) -> Result<axis::Commit, Error> {
        let (axis, value) = self.staged.ok_or(Refusal::NoStagedOperation)?;
        let mut ledger = Ledger::new(self.source)?;
        let target = ledger.target(
            self.source,
            self.slide.as_selector(),
            self.chart.as_selector(),
        )?;
        let before = read_axis(&mut ledger, self.source, &target, axis, value.kind())?;
        let published = if before == value {
            None
        } else {
            ledger.write(self.source, &target, &[], |workspace, limits| {
                engine::write_axis(
                    workspace,
                    target.selection.identifier,
                    target.kind,
                    axis,
                    &value,
                    limits,
                )
            })?
        };
        let Some(published) = published else {
            return Ok(axis::Commit {
                package: self.source.snapshot(),
                patch: axis::Patch {
                    artifacts: noop_artifacts(self.source)?,
                    slide: target.selection.slide,
                    chart: target.selection.chart,
                    axis,
                    before: before.clone(),
                    after: before,
                    plan: MutationPlan::default(),
                },
                diagnostics: Diagnostics::unchanged(),
            });
        };
        let candidate = ledger.target(
            &published.package,
            self.slide.as_selector(),
            self.chart.as_selector(),
        )?;
        if read_axis(
            &mut ledger,
            &published.package,
            &candidate,
            axis,
            value.kind(),
        )? != value
        {
            return Err(Error::Verification);
        }
        Ok(axis::Commit {
            patch: axis::Patch {
                artifacts: artifacts(self.source, &published.package)?,
                slide: target.selection.slide,
                chart: target.selection.chart,
                axis,
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
    /// The stored (or native default) value of one chart-level style
    /// property of one slide chart.
    ///
    /// # Errors
    ///
    /// Selector errors, [`Refusal::UnsupportedKind`],
    /// [`Refusal::MissingStyle`], source and limit errors.
    pub fn slide_chart_style<'slide, 'chart>(
        &self,
        slide: impl Into<SlideSelector<'slide>>,
        chart: impl Into<ChartSelector<'chart>>,
        kind: style::PropertyKind,
    ) -> Result<style::Property, Error> {
        let mut ledger = Ledger::new(self)?;
        let target = ledger.target(self, slide.into(), chart.into())?;
        read_style(&mut ledger, self, &target, kind)
    }

    /// Start an exact-source chart-level style edit of one slide chart.
    ///
    /// ```no_run
    /// use litchi_keynote::Package;
    /// use litchi_keynote::chart::style::PropertyKind;
    ///
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let package = Package::open("input.key")?;
    /// let mut edit = package.edit_slide_chart_style(0, 0)?;
    /// edit.reset(PropertyKind::RoundedCorners)?;
    /// let commit = edit.commit()?;
    /// let restored = commit.package().apply_slide_chart_style(&commit.patch().inverse())?;
    /// # let _ = restored;
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// # Errors
    ///
    /// Selector ceiling errors.
    pub fn edit_slide_chart_style<'slide, 'chart>(
        &self,
        slide: impl Into<SlideSelector<'slide>>,
        chart: impl Into<ChartSelector<'chart>>,
    ) -> Result<StyleEdit<'_>, Error> {
        let (slide, chart) = owned_selectors(self, slide, chart)?;
        Ok(StyleEdit {
            source: self,
            slide,
            chart,
            staged: None,
        })
    }

    /// Apply a chart-level style patch only to its exact source package.
    ///
    /// # Errors
    ///
    /// [`Error::PatchConflict`] when this package is not the patch's exact
    /// source; locality and verification errors.
    pub fn apply_slide_chart_style(&self, patch: &style::Patch) -> Result<style::Commit, Error> {
        let mut ledger = Ledger::new(self)?;
        if !authorizes(self, &patch.artifacts)? {
            return Err(Error::PatchConflict);
        }
        if patch.is_noop() {
            return Ok(style::Commit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: Diagnostics::unchanged(),
            });
        }
        let candidate = ledger.reopen(
            self,
            &patch.artifacts,
            &patch.plan,
            chart_fields(patch.after.kind()),
        )?;
        let target = ledger.target(
            &candidate,
            SlideSelector::position(patch.slide),
            ChartSelector::position(patch.chart),
        )?;
        if read_style(&mut ledger, &candidate, &target, patch.after.kind())? != patch.after {
            return Err(Error::Verification);
        }
        Ok(style::Commit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: Diagnostics::published(0),
        })
    }

    /// The stored (or native default) value of one legend-style
    /// property of one slide chart.
    ///
    /// # Errors
    ///
    /// Selector errors, [`Refusal::UnsupportedKind`],
    /// [`Refusal::MissingStyle`], source and limit errors.
    pub fn slide_chart_legend_style<'slide, 'chart>(
        &self,
        slide: impl Into<SlideSelector<'slide>>,
        chart: impl Into<ChartSelector<'chart>>,
        kind: legend_style::PropertyKind,
    ) -> Result<legend_style::Property, Error> {
        let mut ledger = Ledger::new(self)?;
        let target = ledger.target(self, slide.into(), chart.into())?;
        read_legend_style(&mut ledger, self, &target, kind)
    }

    /// Start an exact-source legend-style edit of one slide chart.
    ///
    /// ```no_run
    /// use litchi_keynote::Package;
    /// use litchi_keynote::chart::legend_style::PropertyKind;
    ///
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let package = Package::open("input.key")?;
    /// let mut edit = package.edit_slide_chart_legend_style(0, 0)?;
    /// edit.reset(PropertyKind::Fill)?;
    /// let commit = edit.commit()?;
    /// let restored = commit.package().apply_slide_chart_legend_style(&commit.patch().inverse())?;
    /// # let _ = restored;
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// # Errors
    ///
    /// Selector ceiling errors.
    pub fn edit_slide_chart_legend_style<'slide, 'chart>(
        &self,
        slide: impl Into<SlideSelector<'slide>>,
        chart: impl Into<ChartSelector<'chart>>,
    ) -> Result<LegendStyleEdit<'_>, Error> {
        let (slide, chart) = owned_selectors(self, slide, chart)?;
        Ok(LegendStyleEdit {
            source: self,
            slide,
            chart,
            staged: None,
        })
    }

    /// Apply a legend-style patch only to its exact source package.
    ///
    /// # Errors
    ///
    /// [`Error::PatchConflict`] when this package is not the patch's exact
    /// source; locality and verification errors.
    pub fn apply_slide_chart_legend_style(
        &self,
        patch: &legend_style::Patch,
    ) -> Result<legend_style::Commit, Error> {
        let mut ledger = Ledger::new(self)?;
        if !authorizes(self, &patch.artifacts)? {
            return Err(Error::PatchConflict);
        }
        if patch.is_noop() {
            return Ok(legend_style::Commit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: Diagnostics::unchanged(),
            });
        }
        let candidate = ledger.reopen(self, &patch.artifacts, &patch.plan, LEGEND_STYLE_FIELDS)?;
        let target = ledger.target(
            &candidate,
            SlideSelector::position(patch.slide),
            ChartSelector::position(patch.chart),
        )?;
        if read_legend_style(&mut ledger, &candidate, &target, patch.after.kind())? != patch.after {
            return Err(Error::Verification);
        }
        Ok(legend_style::Commit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: Diagnostics::published(0),
        })
    }

    /// The stored (or native default) value of one axis-style property of
    /// one slide chart's primary `axis`.
    ///
    /// # Errors
    ///
    /// Selector errors, [`Refusal::UnsupportedAxis`], source and limit
    /// errors.
    pub fn slide_chart_axis<'slide, 'chart>(
        &self,
        slide: impl Into<SlideSelector<'slide>>,
        chart: impl Into<ChartSelector<'chart>>,
        axis: Axis,
        kind: axis::PropertyKind,
    ) -> Result<axis::Property, Error> {
        let mut ledger = Ledger::new(self)?;
        let target = ledger.target(self, slide.into(), chart.into())?;
        read_axis(&mut ledger, self, &target, axis, kind)
    }

    /// Start an exact-source axis-style edit of one slide chart.
    ///
    /// # Errors
    ///
    /// Selector ceiling errors.
    pub fn edit_slide_chart_axis<'slide, 'chart>(
        &self,
        slide: impl Into<SlideSelector<'slide>>,
        chart: impl Into<ChartSelector<'chart>>,
    ) -> Result<AxisEdit<'_>, Error> {
        let (slide, chart) = owned_selectors(self, slide, chart)?;
        Ok(AxisEdit {
            source: self,
            slide,
            chart,
            staged: None,
        })
    }

    /// Apply an axis-style patch only to its exact source package.
    ///
    /// # Errors
    ///
    /// [`Error::PatchConflict`] when this package is not the patch's exact
    /// source; locality and verification errors.
    pub fn apply_slide_chart_axis(&self, patch: &axis::Patch) -> Result<axis::Commit, Error> {
        let mut ledger = Ledger::new(self)?;
        if !authorizes(self, &patch.artifacts)? {
            return Err(Error::PatchConflict);
        }
        if patch.is_noop() {
            return Ok(axis::Commit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: Diagnostics::unchanged(),
            });
        }
        let candidate = ledger.reopen(self, &patch.artifacts, &patch.plan, &[])?;
        let target = ledger.target(
            &candidate,
            SlideSelector::position(patch.slide),
            ChartSelector::position(patch.chart),
        )?;
        if read_axis(
            &mut ledger,
            &candidate,
            &target,
            patch.axis,
            patch.after.kind(),
        )? != patch.after
        {
            return Err(Error::Verification);
        }
        Ok(axis::Commit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: Diagnostics::published(0),
        })
    }
}
