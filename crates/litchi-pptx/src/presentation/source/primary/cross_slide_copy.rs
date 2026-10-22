//! Immutable cross-presentation slide copies into a source-backed
//! [`Package`] (record 0356b).
//!
//! [`Package::plan_cross_slide_copy`] checks one dependency-free copy of a
//! slide of another source-backed presentation into this package and returns
//! a [`Plan`]; [`Package::edit_cross_slide_copy`] starts an [`Edit`] whose
//! [`Edit::stage`] stages a plan and [`Edit::copy`] plans and stages in one
//! call. [`Edit::commit`] returns a [`Commit`] with the new package, the
//! [`Patch`] and the [`Diagnostics`]. The source and destination packages
//! never change.
//!
//! The committed package stays source-backed: writing it
//! ([`Package::write_to`], [`Package::save`]) reruns the bounded planner
//! against both sources, then streams the destination with the copied slide,
//! its media and chart parts added and every untouched member raw-copied.
//! Reading or editing the committed package materializes the published
//! artifact once.
//!
//! The copy is one-way: a [`Patch`] replays onto a package over the same
//! retained destination source, and its inverse applied to the commit's own
//! package restores the package the commit was made on; it is not a durable
//! removal of the copied slide (stage [`super::Edit::remove_slide`] for that).
//!
//! ```no_run
//! use litchi_pptx::SourceBackedPresentation;
//! use litchi_pptx::source::Package;
//!
//! # fn main() -> litchi_pptx::Result<()> {
//! let source = SourceBackedPresentation::open("source.pptx")?;
//! let destination = Package::open("destination.pptx")?;
//! let mut edit = destination.edit_cross_slide_copy()?;
//! // Copy source slide 0 after destination slide 0, reusing its layout.
//! let copied = edit.copy(&source, 0, 0, 1)?;
//! assert_eq!(copied.insertion_position(), 1);
//! edit.commit()?.package().save("copied.pptx")?;
//! # Ok(())
//! # }
//! ```

use std::sync::Arc;

use litchi_core::Position;

use super::Package;
use crate::Result;
use crate::presentation::source::{SourceBackedPresentation, SourceBackedPresentationEditor};
use crate::presentation::source_cross_copy::SourceBackedCrossSlideCopyPlan;

pub use crate::Error;
pub use crate::opened::LimitKind;
pub use crate::presentation::source_cross_copy::CopiedSlide as Copied;
pub use litchi_core::edit::Diagnostics;

/// The result of a cross-slide-copy edit's `commit`: the new source-backed
/// package, the patch and the diagnostics.
pub type Commit = litchi_core::edit::Commit<Package, Patch>;

/// One checked cross-presentation slide copy into one destination package.
///
/// A plan is process-local. Staging it requires the destination it was
/// planned on; publication reruns the planner against both current sources.
#[derive(Clone)]
pub struct Plan {
    plan: Arc<SourceBackedCrossSlideCopyPlan>,
    destination: Arc<SourceBackedPresentationEditor>,
}

impl std::fmt::Debug for Plan {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Plan")
            .field("source_position", &self.source_position())
            .field("insertion_position", &self.insertion_position())
            .finish_non_exhaustive()
    }
}

impl Plan {
    /// Source zero-based slide position captured by this plan.
    #[must_use]
    pub fn source_position(&self) -> usize {
        self.plan.source_position()
    }

    /// Destination slide position whose registered layout is reused.
    #[must_use]
    pub fn destination_slide_position(&self) -> usize {
        self.plan.destination_slide_position()
    }

    /// Destination zero-based insertion position captured by this plan.
    #[must_use]
    pub fn insertion_position(&self) -> usize {
        self.plan.insertion_position()
    }

    /// Destination slide count after the planned insertion.
    #[must_use]
    pub fn destination_slide_count(&self) -> usize {
        self.plan.destination_slide_count()
    }

    fn copied(&self) -> Result<Copied> {
        self.plan.predicted()
    }

    /// Check the plan against `package` and return the committed package
    /// whose output publishes it.
    fn publish(&self, package: &Package) -> Result<Package> {
        let base = super::svg::physical(package)?;
        if !base.same_destination(&self.destination) {
            return Err(Error::PatchConflict);
        }
        let owner = Arc::clone(&self.destination);
        let plan = Arc::clone(&self.plan);
        Ok(base.deferred(move |writer| {
            (*owner)
                .clone()
                .publish_cross_slide_copy_to_stream(writer, &plan)
                .map(drop)
        }))
    }
}

/// The change of a cross-slide-copy edit: the plans it staged, in order.
#[derive(Clone, Default)]
pub struct Patch {
    steps: Vec<Plan>,
    inverse: bool,
    /// The packages the patch was committed between, so its inverse applied
    /// to the commit's own package restores the exact source package.
    restore: Option<(Package, Package)>,
}

impl Patch {
    /// Whether the patch changes nothing.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.steps.is_empty()
    }

    /// The patch that restores the before state: it applies to the commit's
    /// own package only.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            steps: self.steps.clone(),
            inverse: !self.inverse,
            restore: self
                .restore
                .as_ref()
                .map(|(before, after)| (after.clone(), before.clone())),
        }
    }
}

impl std::fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("steps", &self.steps.len())
            .field("inverse", &self.inverse)
            .finish()
    }
}

/// A staged cross-slide-copy edit of one source-backed [`Package`].
/// Dropping it discards it.
pub struct Edit<'a> {
    source: &'a Package,
    current: Package,
    steps: Vec<Plan>,
}

impl std::fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("steps", &self.steps.len())
            .finish_non_exhaustive()
    }
}

impl Package {
    /// Plan one dependency-free copy of the slide `source_slide` of `source`
    /// into this package at `insertion`, reusing the registered layout of
    /// the destination slide `layout_slide`.
    ///
    /// Only a dependency-free slide closure is supported: one internal
    /// slide-layout relationship plus direct embedded images and chart hosts,
    /// no diagram, table, notes, comments, external or shared-owner
    /// relationship. Unnamed slides keep their absent or empty producer
    /// names; nonempty name collisions are refused.
    ///
    /// # Errors
    ///
    /// Returns the typed refusal for an unsupported source closure or a
    /// name collision, an error for a position out of range, and the
    /// source's execution or limit errors.
    pub fn plan_cross_slide_copy(
        &self,
        source: &SourceBackedPresentation,
        source_slide: impl Into<Position>,
        layout_slide: impl Into<Position>,
        insertion: impl Into<Position>,
    ) -> Result<Plan> {
        let base = super::svg::physical(self)?;
        let destination = Arc::clone(&base.owner);
        let plan = destination.plan_cross_slide_copy(
            source,
            source_slide.into().get(),
            layout_slide.into().get(),
            insertion.into().get(),
        )?;
        Ok(Plan {
            plan: Arc::new(plan),
            destination,
        })
    }

    /// Start a cross-slide-copy edit.
    ///
    /// # Errors
    ///
    /// Returns the source's execution or version error.
    pub fn edit_cross_slide_copy(&self) -> Result<Edit<'_>> {
        Ok(Edit {
            source: self,
            current: self.effective()?.clone(),
            steps: Vec::new(),
        })
    }

    /// Apply a cross-slide-copy patch to a new package; this one is
    /// unchanged.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package does not hold the
    /// patch's before state. An independently opened package is accepted when
    /// its published bytes match the recorded state; the replayed output is
    /// then rebased onto that recipient's source policy.
    pub fn apply_cross_slide_copy(&self, patch: &Patch) -> Result<Commit> {
        if let Some((before, after)) = &patch.restore {
            // Keep the original same-owner path. It compares the retained
            // semantic revision without scanning an output and returns the
            // already-published package, so ordinary replay keeps its O(1)
            // identity and allocation behavior.
            if self.same_source_owner(before) {
                if self.same_content(before)? {
                    return Ok(Commit::new(
                        after.clone(),
                        patch.clone(),
                        Diagnostics::new(!patch.is_noop()),
                    ));
                }
            // A different source owner may still carry the exact recorded
            // bytes. Only this independently-owned path pays the bounded
            // content proof and receives a recipient-context deferred result.
            } else if self.same_replay_content(before)? {
                if patch.is_noop() {
                    return Ok(Commit::new(
                        self.clone(),
                        patch.clone(),
                        Diagnostics::new(false),
                    ));
                }
                let package = self.deferred_replay_from(after)?;
                return Ok(Commit::new(
                    package,
                    patch.clone(),
                    Diagnostics::new(!patch.is_noop()),
                ));
            }
        }
        if patch.inverse && !patch.is_noop() {
            return Err(Error::PatchConflict);
        }
        let mut current = self.clone();
        for plan in &patch.steps {
            current = plan.publish(&current)?;
        }
        Ok(Commit::new(
            current,
            patch.clone(),
            Diagnostics::new(!patch.is_noop()),
        ))
    }

    fn same_destination(&self, destination: &Arc<SourceBackedPresentationEditor>) -> bool {
        Arc::ptr_eq(&self.owner, destination)
            || self.owner.package.source_lineage() == destination.package.source_lineage()
    }
}

impl Edit<'_> {
    /// Stage a copy planned on this edit's staged package. Returns what the
    /// copy will publish.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] for a plan made on another package
    /// (including this edit's source after an earlier staged copy: plan
    /// further copies with [`Self::copy`]).
    pub fn stage(&mut self, plan: &Plan) -> Result<Copied> {
        let copied = plan.copied()?;
        self.current = plan.publish(&self.current)?;
        self.steps.push(plan.clone());
        Ok(copied)
    }

    /// Plan one copy against the staged package and stage it; see
    /// [`Package::plan_cross_slide_copy`].
    ///
    /// # Errors
    ///
    /// As [`Package::plan_cross_slide_copy`].
    pub fn copy(
        &mut self,
        source: &SourceBackedPresentation,
        source_slide: impl Into<Position>,
        layout_slide: impl Into<Position>,
        insertion: impl Into<Position>,
    ) -> Result<Copied> {
        let plan =
            self.current
                .plan_cross_slide_copy(source, source_slide, layout_slide, insertion)?;
        self.stage(&plan)
    }

    /// Discard every staged change.
    pub fn reset(&mut self) {
        self.current = self.source.clone();
        self.steps.clear();
    }

    /// Publish the staged changes to a new package.
    ///
    /// # Errors
    ///
    /// Infallible today; the `Result` keeps the family's commit shape.
    pub fn commit(self) -> Result<Commit> {
        let changed = !self.steps.is_empty();
        let package = if changed {
            self.current
        } else {
            self.source.clone()
        };
        let patch = Patch {
            steps: self.steps,
            inverse: false,
            restore: Some((self.source.clone(), package.clone())),
        };
        Ok(Commit::new(package, patch, Diagnostics::new(changed)))
    }
}
