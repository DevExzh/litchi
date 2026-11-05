//! Detached composition of ordinary presentation-domain edits.

mod comments;
mod links;
mod masters;
mod media;
mod modern_comments;
mod notes;
mod picture;
mod placeholders;
mod smartart;
mod table;
mod text;
mod theme;

pub use media::MediaFile;
pub use modern_comments::ModernCommentChange;
pub use smartart::SmartArtParts;

use std::collections::{BTreeSet, HashMap, HashSet, VecDeque};
use std::sync::{Arc, OnceLock, Weak};

use litchi_opc::{BlobPart, OpcPackage, PackURI, Part, TargetMode};
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

pub use litchi_core::edit::Diagnostics;

use super::model::{Slide, SlideNameIndex, Snapshot, invalid};
use super::patch::Patch;
use crate::{Error, Result};
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;

fn drawing_error(error: litchi_drawingml::Error, context: &str) -> Error {
    match error {
        litchi_drawingml::Error::Invalid(_) | litchi_drawingml::Error::Xml(_) => {
            invalid(format!("{context} is invalid"))
        },
        error => Error::Drawing(error),
    }
}

/// One failure-atomic edit rooted in an immutable opened-package snapshot.
///
/// Start with [`Package::edit`](crate::Package::edit) or [`Snapshot::edit`].
/// Dropping this value discards all staged changes without changing its source.
///
/// A signed package is edited only after its signature is removed
/// explicitly with [`Package::unsign`](crate::Package::unsign) (ADR 0006):
/// [`Self::commit`] refuses a changed commit over a signed package, and the
/// verbs that stage through the package-level edit families (comments,
/// modern comments, table styles, charts, masters and layouts) refuse it
/// while staging.
#[derive(Clone)]
pub struct Edit {
    source: Option<Snapshot>,
    lazy: Option<LazySource>,
    working: Working,
    slides: Vec<Slide>,
    slide_name_index: Option<SlideNameIndex>,
    state_ready: bool,
    /// Slide payloads this transaction read back as a complete scene when it
    /// staged them, by part name.
    ///
    /// A record is a `Weak`, so it pins no payload, and it vouches only for
    /// the exact allocation it names: a payload replaced or mutated in place
    /// since then no longer matches (`Arc::make_mut` on a sole owner
    /// disassociates weak references). Commit consults it to avoid reading
    /// the same bytes as a scene a second time; a miss is the ordinary read.
    scene_read: HashMap<PackURI, Weak<Vec<u8>>>,
    /// The published snapshot and generic patch of a staged slide plan that
    /// was this edit's first change. Commit publishes them directly while the
    /// staged graph is still exactly the plan's candidate.
    planned: Option<(Snapshot, Patch)>,
}

/// The edit-only state retained by [`Package::edit`](crate::Package::edit)
/// before the first operation needs a validated opened snapshot. Keeping this
/// separate from [`Snapshot`] is important: a lazy edit may be dropped
/// without parsing the slide catalog, notes, names, or the complete package
/// fingerprint, while explicit read projections remain eager and fallible.
#[derive(Clone)]
struct LazySource {
    package: Arc<OpcPackage>,
    replay_index: Arc<crate::package::replay_index::ReplayIndex>,
    limits: super::Limits,
    physical_source_provenance: bool,
    owner_state: Arc<crate::package::OpenedOwnerState>,
    parent_digests: Option<Arc<super::PartDigests>>,
    digest_sink: Arc<OnceLock<Arc<super::PartDigests>>>,
    release_retained_mce: bool,
}

/// Copy-on-write graph state for an edit. A detached edit keeps the source's
/// `Arc` until an operation actually needs a mutable graph; this covers both
/// `Package::edit` and `Snapshot::edit` without traversing the OPC part map in
/// their constructors.
#[derive(Clone)]
enum Working {
    Deferred(Arc<OpcPackage>),
    Owned(Box<OpcPackage>),
}

impl Working {
    fn new(package: Arc<OpcPackage>) -> Self {
        Self::Deferred(package)
    }

    fn into_owned(self) -> OpcPackage {
        match self {
            Self::Owned(package) => *package,
            Self::Deferred(package) => match Arc::try_unwrap(package) {
                Ok(package) => package,
                Err(package) => package.as_ref().clone(),
            },
        }
    }
}

impl std::ops::Deref for Working {
    type Target = OpcPackage;

    fn deref(&self) -> &Self::Target {
        match self {
            Self::Deferred(package) => package.as_ref(),
            Self::Owned(package) => package,
        }
    }
}

impl std::ops::DerefMut for Working {
    fn deref_mut(&mut self) -> &mut Self::Target {
        loop {
            match self {
                Self::Owned(package) => return package,
                Self::Deferred(package) => {
                    let package = package.as_ref().clone();
                    *self = Self::Owned(Box::new(package));
                },
            }
        }
    }
}

impl std::fmt::Debug for Edit {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("captured", &self.source.is_some())
            .field(
                "slide_count",
                &self.state_ready.then_some(self.slides.len()),
            )
            .finish_non_exhaustive()
    }
}

impl Edit {
    fn semantic_limits_ref(&self) -> Option<crate::open_options::SemanticLimitsRef> {
        self.source
            .as_ref()
            .map(|source| source.owner_state.semantic_limits.clone())
            .or_else(|| {
                self.lazy
                    .as_ref()
                    .map(|lazy| lazy.owner_state.semantic_limits.clone())
            })
            .flatten()
    }

    pub(crate) fn new(source: Snapshot) -> Self {
        Self {
            working: Working::new(Arc::clone(&source.package)),
            slides: Vec::new(),
            slide_name_index: None,
            state_ready: false,
            scene_read: HashMap::new(),
            planned: None,
            source: Some(source),
            lazy: None,
        }
    }

    /// Construct an edit without capturing the presentation. The first
    /// operation that needs semantic state calls [`Self::ensure_ready`].
    pub(crate) fn new_lazy(
        package: Arc<OpcPackage>,
        replay_index: Arc<crate::package::replay_index::ReplayIndex>,
        limits: super::Limits,
        physical_source_provenance: bool,
        owner_state: Arc<crate::package::OpenedOwnerState>,
        digest_sink: Arc<OnceLock<Arc<super::PartDigests>>>,
    ) -> Self {
        let parent_digests = digest_sink.get().cloned();
        Self {
            working: Working::new(Arc::clone(&package)),
            slides: Vec::new(),
            slide_name_index: None,
            state_ready: false,
            scene_read: HashMap::new(),
            planned: None,
            source: None,
            lazy: Some(LazySource {
                package,
                replay_index,
                limits,
                physical_source_provenance,
                owner_state,
                parent_digests,
                digest_sink,
                release_retained_mce: false,
            }),
        }
    }

    /// Materialize the existing validated opened-presentation state exactly
    /// once. An error leaves the lazy source intact, so a caller can retry a
    /// failed operation without observing a partially initialized edit.
    fn ensure_ready(&mut self) -> Result<()> {
        if self.source.is_some() && self.state_ready {
            return Ok(());
        }
        if let Some(source) = self.source.as_ref() {
            self.slides = source.slides.as_ref().clone();
            self.slide_name_index = Some(source.slide_name_index.as_ref().clone());
            self.state_ready = true;
            return Ok(());
        }
        let Some(lazy) = self.lazy.take() else {
            return Err(invalid(
                "opened-presentation edit source state is unavailable",
            ));
        };
        let captured = match lazy.parent_digests.as_deref() {
            Some(parent) => super::model::capture_with_parent_digests_and_semantic_limits(
                &lazy.package,
                lazy.limits,
                lazy.physical_source_provenance,
                parent,
                lazy.owner_state.semantic_limits.clone(),
            ),
            None => super::model::capture_with_provenance_and_semantic_limits(
                &lazy.package,
                lazy.limits,
                lazy.physical_source_provenance,
                lazy.owner_state.semantic_limits.clone(),
            ),
        };
        let mut snapshot = match captured {
            Ok(snapshot) => snapshot,
            Err(error) => {
                self.lazy = Some(lazy);
                return Err(error);
            },
        };
        snapshot.replay_index = Some(Arc::clone(&lazy.replay_index));
        snapshot.owner_state = Arc::clone(&lazy.owner_state);
        if lazy.release_retained_mce {
            snapshot.release_retained_mce();
        }
        if lazy.digest_sink.get().is_none()
            && let Ok(projected) = snapshot.package_memo(lazy.package.as_ref())
        {
            let _kept = lazy.digest_sink.set(Arc::new(projected));
        }
        self.slides = snapshot.slides.as_ref().clone();
        self.slide_name_index = Some(snapshot.slide_name_index.as_ref().clone());
        self.state_ready = true;
        self.source = Some(snapshot);
        Ok(())
    }

    /// Borrow the currently staged graph, or the immutable source graph before
    /// the first mutation has required a private working copy.
    fn working_ref(&self) -> Result<&OpcPackage> {
        Ok(&self.working)
    }

    fn source_ref(&self) -> Result<&Snapshot> {
        self.source
            .as_ref()
            .ok_or_else(|| invalid("opened-presentation edit source state is unavailable"))
    }

    pub(crate) fn semantic_limits(&self) -> crate::open_options::SemanticLimits {
        self.source
            .as_ref()
            .and_then(|source| source.owner_state.semantic_limits.as_deref().copied())
            .or_else(|| {
                self.lazy
                    .as_ref()
                    .and_then(|source| source.owner_state.semantic_limits.as_deref().copied())
            })
            .unwrap_or_default()
    }

    fn notes_limits(&self) -> crate::notes::Limits {
        self.semantic_limits()
            .notes()
            .constrained_by_source(self.working.read_limits())
    }

    fn comments_limits(&self) -> crate::comments::Limits {
        self.semantic_limits()
            .comments()
            .constrained_by_source(self.working.read_limits())
    }

    /// Record that the payload now staged for `part_name` was read back as a
    /// complete scene. A part whose shared allocation does not alias its
    /// visible payload, or a record that cannot be reserved, is simply not
    /// recorded, which only costs the commit a second read.
    fn note_scene_read(&mut self, part_name: &PackURI) {
        let Ok(part) = self.working.get_part(part_name) else {
            return;
        };
        let payload = part.blob_arc();
        if !std::ptr::eq(payload.as_slice(), part.blob()) || self.scene_read.try_reserve(1).is_err()
        {
            return;
        }
        self.scene_read
            .insert(part_name.clone(), Arc::downgrade(&payload));
    }

    /// Immutable source root for this transaction.
    #[must_use = "use the source snapshot or handle its error"]
    pub fn source(&mut self) -> Result<&Snapshot> {
        self.ensure_ready()?;
        self.source_ref()
    }

    /// Slides in the transaction's currently staged order.
    #[must_use = "use the staged slides or handle its error"]
    pub fn slides(&mut self) -> Result<&[Slide]> {
        self.ensure_ready()?;
        Ok(&self.slides)
    }

    /// Aggregate charged bytes held by the source snapshot's optional MCE
    /// projection table. Cloned snapshots may report the same shared charge,
    /// so per-owner values are not additive physical-memory measurements.
    #[must_use = "use the retained byte count or handle its error"]
    pub fn retained_mce_bytes(&mut self) -> Result<usize> {
        self.ensure_ready()?;
        Ok(self.source_ref()?.retained_mce_bytes())
    }

    /// Release this transaction's source reference to retained transformed
    /// slide projections. Its staged package and semantic edits are unchanged;
    /// snapshot clones may keep the shared table alive.
    pub fn release_retained_mce(&mut self) {
        if let Some(source) = self.source.as_mut() {
            source.release_retained_mce();
        } else if let Some(lazy) = self.lazy.as_mut() {
            lazy.release_retained_mce = true;
        }
    }

    /// Insert an empty slide at a checked zero-based position using an exact
    /// registered layout name. The position may equal the current slide count.
    ///
    /// The layout is shared; its placeholder shapes are not copied into the
    /// new slide. Add shapes through this transaction after insertion. Existing
    /// presentation and metadata XML retain their source spelling.
    ///
    /// ```
    /// let source = litchi_pptx::Package::new()?;
    /// let mut edit = source.edit()?;
    /// let slide = edit.insert_slide(litchi_core::Position::new(0), "Blank")?;
    /// edit.add_text_box(slide.name(), "Hello", (0, 0, 914_400, 914_400))?;
    /// let presentation = edit.commit()?.into_package();
    /// # Ok::<(), litchi_pptx::Error>(())
    /// ```
    ///
    /// # Errors
    /// Returns an error for a missing or ambiguous layout, an invalid position,
    /// unsupported source topology, signatures/protection, or exceeded limits.
    /// Failure leaves every previously staged change intact.
    pub fn insert_slide(
        &mut self,
        position: litchi_core::Position,
        layout_name: &str,
    ) -> Result<Slide> {
        self.ensure_ready()?;
        let (limits, physical_source_provenance) = {
            let source = self.source_ref()?;
            (source.limits, source.physical_source_provenance)
        };
        let owner_state = Arc::clone(&self.source_ref()?.owner_state);
        let mut current = super::model::capture_with_provenance_and_semantic_limits(
            &self.working,
            limits,
            physical_source_provenance,
            owner_state.semantic_limits.clone(),
        )?;
        current.owner_state = Arc::clone(&owner_state);
        let patch = super::insert::plan(&current, position.get(), layout_name)?;
        let (working, published) = super::patch::apply_to_new_with_semantic_limits(
            &self.working,
            &patch,
            physical_source_provenance,
            &current.part_digests,
            owner_state.semantic_limits.clone(),
        )?;
        let selected = published
            .slides
            .get(position.get())
            .cloned()
            .ok_or_else(|| invalid("slide insertion lost its checked position"))?;
        self.working = Working::Owned(Box::new(working));
        self.slides = published.slides.as_ref().clone();
        self.slide_name_index = Some(published.slide_name_index.as_ref().clone());
        Ok(selected)
    }

    /// Move a slide selected by exact name or position to a checked slot.
    ///
    /// Both selection and destination refer to the currently staged order.
    /// The destination is the final zero-based position after detachment.
    /// Missing/ambiguous selectors and unsupported slide-list dependencies
    /// fail before changing this edit.
    pub fn move_slide<'s>(
        &mut self,
        from: impl Into<crate::slide::Selector<'s>>,
        to: litchi_core::Position,
    ) -> Result<bool> {
        self.ensure_ready()?;
        let from = self.resolve_slide_position(from.into())?;
        let to = to.get();
        if to >= self.slides.len() {
            return Err(Error::SlideIndexOutOfBounds {
                index: to,
                len: self.slides.len(),
            });
        }
        if from == to {
            return Ok(false);
        }
        let mut order = self.slide_order_ids()?;
        let id = order.remove(from);
        order.insert(to, id);
        self.reorder_slide_ids_moving(&order, Some(to))
    }

    /// Exchange two selected slides in one failure-atomic order edit.
    pub fn swap_slides<'a, 'b>(
        &mut self,
        left: impl Into<crate::slide::Selector<'a>>,
        right: impl Into<crate::slide::Selector<'b>>,
    ) -> Result<bool> {
        self.ensure_ready()?;
        let left = self.resolve_slide_position(left.into())?;
        let right = self.resolve_slide_position(right.into())?;
        if left == right {
            return Ok(false);
        }
        let mut order = self.slide_order_ids()?;
        order.swap(left, right);
        self.reorder_slide_ids(&order)
    }

    /// Replace the complete order by semantic slide selectors.
    ///
    /// All selectors resolve against the same currently staged order. They
    /// must name every slide exactly once; raw slide IDs remain internal.
    pub fn reorder_slides(&mut self, selectors: &[crate::slide::Selector<'_>]) -> Result<bool> {
        self.ensure_ready()?;
        if selectors.len() != self.slides.len() {
            return Err(invalid(
                "slide order must select every staged slide exactly once",
            ));
        }
        let mut order = Vec::new();
        order
            .try_reserve_exact(selectors.len())
            .map_err(|source| Error::Allocation {
                resource: "opened-presentation selected slide order",
                source,
            })?;
        for selector in selectors {
            let selector = match selector {
                crate::slide::Selector::Name(name) => crate::slide::Selector::from(name.as_ref()),
                crate::slide::Selector::Position(position) => {
                    crate::slide::Selector::from(*position)
                },
            };
            let position = self.resolve_slide_position(selector)?;
            order.push(self.slides[position].id);
        }
        self.reorder_slide_ids(&order)
    }

    fn slide_order_ids(&self) -> Result<Vec<u32>> {
        let mut order = Vec::new();
        order
            .try_reserve_exact(self.slides.len())
            .map_err(|source| Error::Allocation {
                resource: "opened-presentation slide order identities",
                source,
            })?;
        order.extend(self.slides.iter().map(Slide::id));
        Ok(order)
    }

    fn reorder_slide_ids(&mut self, ordered_ids: &[u32]) -> Result<bool> {
        self.reorder_slide_ids_moving(ordered_ids, None)
    }

    /// Stage `ordered_ids`. With `moved`, the slide at that position left its
    /// section and joins its new neighbour's, as `PowerPoint` moves a slide;
    /// otherwise every slide keeps its section and each section is re-sorted
    /// into the new order (p14 sections, record 0252).
    fn reorder_slide_ids_moving(
        &mut self,
        ordered_ids: &[u32],
        moved: Option<usize>,
    ) -> Result<bool> {
        self.ensure_ready()?;
        if ordered_ids.len() == self.slides.len()
            && ordered_ids
                .iter()
                .zip(&self.slides)
                .all(|(id, slide)| *id == slide.id)
        {
            return Ok(false);
        }
        let presentation_name = self.source_ref()?.presentation_name.clone();
        let main = self.working.get_part(&presentation_name)?;
        let xml = super::xml::reorder_slides(main.blob(), &self.slides, ordered_ids)?;
        let sectioned = match moved {
            Some(position) => super::sections::place(&xml, ordered_ids, position)?,
            None => super::sections::sort(&xml, ordered_ids)?,
        };
        let xml = sectioned.unwrap_or(xml);
        let mut reordered = Vec::new();
        reordered
            .try_reserve_exact(self.slides.len())
            .map_err(|source| Error::Allocation {
                resource: "opened-presentation reordered identities",
                source,
            })?;
        for id in ordered_ids {
            reordered.push(
                self.slides
                    .iter()
                    .find(|slide| slide.id == *id)
                    .cloned()
                    .ok_or_else(|| invalid("opened-presentation slide order lost an identity"))?,
            );
        }
        let slide_name_index = SlideNameIndex::build(&reordered)?;
        self.working.get_part_mut(&presentation_name)?.set_blob(xml);
        self.slides = reordered;
        self.slide_name_index = Some(slide_name_index);
        Ok(true)
    }

    /// Remove one slide through the same bounded ownership proof used by
    /// [`Snapshot::plan_slide_removal`].
    ///
    /// The Parts the slide owns exclusively leave with it (record 0303).
    /// Signatures, macros, MCE, protection, and incoming secondary owners
    /// are refused before this transaction changes.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing selector, the final remaining slide, or
    /// an unsupported slide-list or relationship topology.
    pub fn remove_slide<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
    ) -> Result<Slide> {
        self.ensure_ready()?;
        let (limits, physical_source_provenance, semantic_limits) = {
            let source = self.source_ref()?;
            (
                source.limits,
                source.physical_source_provenance,
                source.owner_state.semantic_limits.clone(),
            )
        };
        let snapshot = super::model::capture_with_provenance_and_semantic_limits(
            &self.working,
            limits,
            physical_source_provenance,
            semantic_limits.clone(),
        )?;
        let plan = snapshot.plan_slide_removal(slide)?;
        let selected = plan.source().clone();
        let published = super::remove_plan::apply_patch_with_semantic_limits(
            &mut self.working,
            plan.patch(),
            physical_source_provenance,
            semantic_limits,
        )?;
        self.slides = published.slides.as_ref().clone();
        self.slide_name_index = Some(published.slide_name_index.as_ref().clone());
        Ok(selected)
    }

    /// Record a staged plan's publication when it was this edit's first
    /// change, so commit can publish it without capturing it again.
    fn record_plan(&mut self, source_revision: [u8; 32], published: Snapshot, patch: &Patch) {
        self.slides = published.slides.as_ref().clone();
        self.slide_name_index = Some(published.slide_name_index.as_ref().clone());
        self.planned = (self
            .source
            .as_ref()
            .is_some_and(|source| source.revision == source_revision)
            && self.planned.is_none())
        .then(|| (published, patch.clone()));
    }

    /// Stage a source-bound same-package whole-slide copy plan
    /// ([`Snapshot::plan_slide_copy`]).
    ///
    /// Planning validated and materialized the complete candidate. Staging
    /// refuses when any part of the staged graph differs from the planned
    /// source, including resources outside the copy write set, so a plan is
    /// staged before any other change of its edit. The plan's durable patch
    /// replays through [`crate::Package::apply`].
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] for a stale plan, the encryption
    /// policy refusal, or an error for a candidate graph that no longer
    /// applies exactly.
    pub fn stage_slide_copy(&mut self, plan: &crate::slide_copy::Plan) -> Result<()> {
        self.ensure_ready()?;
        let (physical_source_provenance, owner_state) = {
            let source = self.source_ref()?;
            (
                source.physical_source_provenance,
                Arc::clone(&source.owner_state),
            )
        };
        owner_state.ensure_plain_mutation("stage_slide_copy")?;
        if super::package_fingerprint(&self.working)? != plan.source_revision() {
            return Err(Error::PatchConflict);
        }
        let mut snapshot = super::apply_with_semantic_limits(
            &mut self.working,
            plan.patch(),
            physical_source_provenance,
            owner_state.semantic_limits.clone(),
        )?;
        snapshot.owner_state = owner_state;
        self.record_plan(plan.source_revision(), snapshot, plan.patch());
        Ok(())
    }

    /// Stage a source-bound whole-slide removal plan
    /// ([`Snapshot::plan_slide_removal`]).
    ///
    /// Planning captured a complete-package fingerprint; staging refuses when
    /// the staged graph differs from it, so a plan is staged before any other
    /// change of its edit. The plan's durable [`crate::slide_removal::Patch`]
    /// replays through [`crate::Package::apply_slide_removal`].
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] for a stale plan, the encryption
    /// policy refusal, or an error for a candidate graph that no longer
    /// applies exactly.
    pub fn stage_slide_removal(&mut self, plan: &crate::slide_removal::Plan) -> Result<()> {
        self.ensure_ready()?;
        let (physical_source_provenance, owner_state) = {
            let source = self.source_ref()?;
            (
                source.physical_source_provenance,
                Arc::clone(&source.owner_state),
            )
        };
        owner_state.ensure_plain_mutation("stage_slide_removal")?;
        let mut snapshot = super::remove_plan::apply_patch_with_semantic_limits(
            &mut self.working,
            plan.patch(),
            physical_source_provenance,
            owner_state.semantic_limits.clone(),
        )?;
        snapshot.owner_state = owner_state;
        let source_revision = plan.patch().source_revision();
        self.record_plan(source_revision, snapshot, plan.patch().opened_patch());
        Ok(())
    }

    /// Stage a source-bound cross-presentation slide-copy plan
    /// ([`Snapshot::plan_cross_slide_copy`]) from the read-only `source`
    /// package.
    ///
    /// The source must retain the complete graph revision captured during
    /// planning and the staged graph must still be the planned destination,
    /// so a plan is staged before any other change of its edit. Source and
    /// destination must have been opened through source-preserving ingress
    /// (`open` or `from_reader`) so discarded ZIP ordering and extra fields
    /// cannot alias the authorization revision. A plan whose candidate
    /// transfers source-compressed media
    /// ([`crate::cross_slide_copy::Plan::transfers_source_compressed_media`])
    /// is rebuilt from the two packages' bytes; into a destination that is
    /// not an unmodified owned source it publishes the candidate reopened
    /// from its archive, and a destination holding a caller-defined `Part`
    /// implementation refuses it. The plan's durable
    /// [`crate::cross_slide_copy::Patch`] replays through
    /// [`crate::Package::apply_cross_slide_copy`].
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] for stale source or destination
    /// graphs or a foreign source, and an error for non-source-preserving
    /// ingress, the encryption policy refusal, a transferring plan applied
    /// to a destination holding a caller-defined part
    /// ([`crate::SlideCopyRefusal::CallerDefinedPart`]), or a candidate graph
    /// that no longer applies exactly.
    pub fn stage_cross_slide_copy(
        &mut self,
        source: &crate::Package,
        plan: &crate::cross_slide_copy::Plan,
    ) -> Result<()> {
        self.ensure_ready()?;
        let (physical_source_provenance, owner_state, destination_digests) = {
            let destination = self.source_ref()?;
            (
                destination.physical_source_provenance,
                Arc::clone(&destination.owner_state),
                Arc::clone(&destination.part_digests),
            )
        };
        if !source.physical_source_provenance || !physical_source_provenance {
            return Err(Error::slide_copy_refused(
                crate::SlideCopyRefusal::UnknownPhysicalMember,
            ));
        }
        source.ensure_plain_mutation("stage_cross_slide_copy")?;
        owner_state.ensure_plain_mutation("stage_cross_slide_copy")?;
        let mut snapshot = super::cross_copy_plan::apply_plan(
            source.opc()?,
            &mut self.working,
            plan,
            source.physical_source_provenance,
            physical_source_provenance,
            super::cross_copy_plan::LiveDigests {
                source: source.part_digest_memo(),
                destination: Some(destination_digests.as_ref()),
            },
            source.semantic_limits.clone(),
            owner_state.semantic_limits.clone(),
        )?;
        snapshot.owner_state = owner_state;
        let source_revision = plan.destination_revision();
        self.record_plan(source_revision, snapshot, plan.patch().opened_patch());
        Ok(())
    }

    /// Replace all visible text runs in one existing semantic shape.
    ///
    /// The first existing `a:t` run receives the escaped replacement and
    /// later runs become empty. Shape structure, formatting, unknown XML, and
    /// relationships remain exact.
    ///
    /// # Errors
    ///
    /// Returns an error for missing or ambiguous selectors, shapes without a
    /// text body, malformed raw XML, invalid characters, or exceeded bounds.
    pub fn set_shape_text<'s, 'k>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        shape: impl Into<crate::shape::Key<'k>>,
        text: impl AsRef<str>,
    ) -> Result<bool> {
        let selected_slide = self.resolve_slide(slide.into())?;
        let key = shape.into();
        let text = text.as_ref();
        let owner = self.working.get_part(&selected_slide.part_name)?;
        crate::parts::validate_content_type(owner, litchi_opc::constants::content_type::PML_SLIDE)?;
        let shape_limits = self.semantic_limits().shape();
        let scene = crate::shape::Scene::read(owner.blob(), shape_limits)?;
        let selected_shape = scene.shape(key)?;
        if selected_shape.common().text() == Some(text) {
            return Ok(false);
        }
        if selected_shape.common().text().is_none() {
            return Err(invalid(
                "opened-presentation selected shape has no text body",
            ));
        }
        let span =
            crate::tag::shape::selected_raw_span_for_shape(owner.blob(), &scene, selected_shape)?;
        let xml = super::xml::rewrite_shape_text_with_limits(
            owner.blob(),
            span,
            text,
            self.source_ref()?.limits.max_text_bytes(),
            shape_limits,
        )?;
        let staged = crate::shape::Scene::read(&xml, shape_limits)?;
        let staged_shape = staged.shape(key)?;
        if staged_shape.common().text() != Some(text) || staged.len() != scene.len() {
            return Err(invalid(
                "opened-presentation shape text did not round-trip semantically",
            ));
        }
        self.working
            .get_part_mut(&selected_slide.part_name)?
            .set_blob(xml);
        self.note_scene_read(&selected_slide.part_name);
        Ok(true)
    }

    /// Replace the authored `p:cSld@name` while preserving every unrelated
    /// slide byte. An equal semantic name is an exact lexical no-op.
    ///
    /// ```
    /// # fn main() -> litchi_pptx::Result<()> {
    /// let source = litchi_pptx::Package::new()?;
    /// let mut edit = source.edit()?;
    /// edit.insert_slide(litchi_core::Position::new(0), "Blank")?;
    /// edit.set_slide_name(0_usize, "Roadmap")?;
    /// let _package = edit.commit()?.into_package();
    /// # Ok(())
    /// # }
    /// ```
    pub fn set_slide_name<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        name: impl AsRef<str>,
    ) -> Result<bool> {
        self.set_slide_properties(slide, Some(name.as_ref()), None)
    }

    /// Set whether the selected slide is hidden during presentation. The
    /// native `p:sld/@show` default is visible, so setting `hidden` to false
    /// preserves an omitted or semantically visible attribute exactly.
    ///
    /// ```
    /// # fn main() -> litchi_pptx::Result<()> {
    /// let source = litchi_pptx::Package::new()?;
    /// let mut edit = source.edit()?;
    /// edit.insert_slide(litchi_core::Position::new(0), "Blank")?;
    /// edit.set_slide_hidden(0_usize, true)?;
    /// let _package = edit.commit()?.into_package();
    /// # Ok(())
    /// # }
    /// ```
    pub fn set_slide_hidden<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        hidden: bool,
    ) -> Result<bool> {
        self.set_slide_properties(slide, None, Some(hidden))
    }

    /// Atomically update the authored slide name and visibility together.
    /// Omitted fields remain byte-for-byte untouched.
    ///
    /// ```
    /// # fn main() -> litchi_pptx::Result<()> {
    /// let source = litchi_pptx::Package::new()?;
    /// let mut edit = source.edit()?;
    /// edit.insert_slide(litchi_core::Position::new(0), "Blank")?;
    /// edit.set_slide_properties(0_usize, Some("Roadmap"), Some(true))?;
    /// let _package = edit.commit()?.into_package();
    /// # Ok(())
    /// # }
    /// ```
    pub fn set_slide_properties<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        name: Option<&str>,
        hidden: Option<bool>,
    ) -> Result<bool> {
        self.ensure_ready()?;
        if name.is_none() && hidden.is_none() {
            return Ok(false);
        }
        let position = self.resolve_slide_position(slide.into())?;
        let selected_slide = self.slides[position].clone();
        let owner = self.working.get_part(&selected_slide.part_name)?;
        crate::parts::validate_content_type(owner, litchi_opc::constants::content_type::PML_SLIDE)?;
        let Some(xml) = super::stage_slide_properties(
            owner.blob(),
            name,
            hidden,
            self.semantic_limits().shape().output_bytes(),
        )?
        else {
            return Ok(false);
        };
        let renamed = if let Some(name) = name {
            let mut slides = self.slides.clone();
            let mut owned = String::new();
            owned
                .try_reserve_exact(name.len())
                .map_err(|source| Error::Allocation {
                    resource: "opened-presentation slide name",
                    source,
                })?;
            owned.push_str(name);
            slides[position].name = owned;
            let index = SlideNameIndex::build(&slides)?;
            Some((slides, index))
        } else {
            None
        };
        self.working
            .get_part_mut(&selected_slide.part_name)?
            .set_blob(xml);
        if let Some((slides, index)) = renamed {
            self.slides = slides;
            self.slide_name_index = Some(index);
        }
        Ok(true)
    }

    /// Set (`Some`) or clear (`None`) the selected slide's direct
    /// transition (`p:transition`, or its markup-compatibility owner for the
    /// PowerPoint 2010+ effects and Morph).
    ///
    /// The value may use the supported PowerPoint extension vocabulary and a
    /// custom duration; unknown effects are refused. An equal modeled value,
    /// or clearing an absent transition, is an exact byte no-op. Inherited
    /// layout or master transitions are never touched. Returns whether the
    /// slide changed.
    ///
    /// ```
    /// # fn main() -> litchi_pptx::Result<()> {
    /// use litchi_pptx::transition::{Kind, Side, Transition};
    /// let source = litchi_pptx::Package::new()?;
    /// let mut edit = source.edit()?;
    /// edit.insert_slide(litchi_core::Position::new(0), "Blank")?;
    /// let push = Transition::new(Kind::Push(Side::Left));
    /// assert!(edit.set_slide_transition(0_usize, Some(&push))?);
    /// assert!(!edit.set_slide_transition(0_usize, Some(&push))?);
    /// let _package = edit.commit()?.into_package();
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// # Errors
    ///
    /// Returns an error for an out-of-range slide, an unsupported or invalid
    /// transition, a modification-protected presentation, malformed slide
    /// XML, or an exhausted output limit. A failure leaves the edit unchanged.
    pub fn set_slide_transition<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        value: Option<&crate::transition::Transition>,
    ) -> Result<bool> {
        let selected_slide = self.resolve_slide(slide.into())?;
        let owner = self.working.get_part(&selected_slide.part_name)?;
        crate::parts::validate_content_type(owner, litchi_opc::constants::content_type::PML_SLIDE)?;
        let presentation_name = self.source_ref()?.presentation_name.clone();
        let main = self.working.get_part(&presentation_name)?;
        let semantic = self.semantic_limits();
        let (xml, changed) = crate::presentation::transition::stage_with_limits(
            owner.blob(),
            main.blob(),
            value,
            semantic.shape().output_bytes(),
            semantic.transition(),
            if value.is_some() {
                "set_slide_transition"
            } else {
                "clear_slide_transition"
            },
            crate::presentation_properties::metadata::protection::Limits::from_source_limits(
                self.working.read_limits(),
            ),
        )?;
        if let Some(xml) = xml {
            self.working
                .get_part_mut(&selected_slide.part_name)?
                .set_blob(xml);
        }
        Ok(changed)
    }

    /// Atomically replace visible text in a bounded set of shapes on one slide.
    ///
    /// Every selector is resolved against the same immutable staged slide.
    /// Duplicate or overlapping selections are refused, caller order does not
    /// affect output bytes, and no package state changes unless the complete
    /// batch reparses and reads back exactly.
    ///
    /// The return value is the number of shapes whose text changed. At most
    /// 256 selector/value pairs and at most the transaction's aggregate text
    /// byte limit are accepted.
    pub fn set_shape_texts<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        replacements: &[super::ShapeTextReplacement<'_>],
    ) -> Result<usize> {
        let selected_slide = self.resolve_slide(slide.into())?;
        let owner = self.working.get_part(&selected_slide.part_name)?;
        crate::parts::validate_content_type(owner, litchi_opc::constants::content_type::PML_SLIDE)?;
        let (xml, changed) = super::xml::stage_shape_texts_with_limits(
            owner.blob(),
            replacements,
            self.source_ref()?.limits.max_text_bytes(),
            self.semantic_limits().shape().output_bytes(),
            self.semantic_limits().shape(),
        )?;
        if let Some(xml) = xml {
            self.working
                .get_part_mut(&selected_slide.part_name)?
                .set_blob(xml);
            // Staging read the batch's output back as a complete scene.
            self.note_scene_read(&selected_slide.part_name);
        }
        Ok(changed)
    }

    /// Atomically replace text from several already bounded same-slide
    /// batches in one scene read, rewrite, and read-back.
    ///
    /// Each inner slice keeps the ordinary 256-selector and aggregate text
    /// limits of [`Self::set_shape_texts`]. The chunks are intended for a
    /// caller that has already established a globally disjoint target set;
    /// the codec still validates that invariant before publishing anything.
    /// This is the bounded chunked transaction form for adapters that retain
    /// the per-call contract while avoiding an O(chunks * slide-size) walk.
    /// At most 100,000 chunks and selections are accepted, with aggregate
    /// replacement text bounded by the default shape output-byte limit.
    pub fn set_shape_text_chunks<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        chunks: &[&[super::ShapeTextReplacement<'_>]],
    ) -> Result<usize> {
        let selected_slide = self.resolve_slide(slide.into())?;
        let owner = self.working.get_part(&selected_slide.part_name)?;
        crate::parts::validate_content_type(owner, litchi_opc::constants::content_type::PML_SLIDE)?;
        let (xml, changed) = super::stage_shape_text_chunks_with_limits(
            owner.blob(),
            chunks,
            self.source_ref()?.limits.max_text_bytes(),
            self.semantic_limits().shape().output_bytes(),
            self.semantic_limits().shape(),
        )?;
        if let Some(xml) = xml {
            self.working
                .get_part_mut(&selected_slide.part_name)?
                .set_blob(xml);
            self.note_scene_read(&selected_slide.part_name);
        }
        Ok(changed)
    }

    /// Atomically replace visible text and/or complete local transform bounds
    /// in a bounded set of positional shapes on one slide.
    ///
    /// Every target is resolved against the same immutable scene. Duplicate
    /// targets are refused, text and geometry are staged against one local
    /// candidate, and the transaction is changed only after the complete
    /// candidate reads back with the requested values. The return value is
    /// the number of shapes whose supplied properties changed.
    pub fn set_shape_properties<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        replacements: &[super::ShapePropertiesReplacement<'_>],
    ) -> Result<usize> {
        let selected_slide = self.resolve_slide(slide.into())?;
        let owner = self.working.get_part(&selected_slide.part_name)?;
        crate::parts::validate_content_type(owner, litchi_opc::constants::content_type::PML_SLIDE)?;
        let (xml, changed) = super::properties::stage_shape_properties_with_limits(
            owner.blob(),
            replacements,
            self.source_ref()?.limits.max_text_bytes(),
            self.semantic_limits().shape().output_bytes(),
            self.semantic_limits().shape(),
        )?;
        if let Some(xml) = xml {
            self.working
                .get_part_mut(&selected_slide.part_name)?
                .set_blob(xml);
            self.note_scene_read(&selected_slide.part_name);
        }
        Ok(changed)
    }

    /// Apply a bounded batch of rich text edits to one slide, in order.
    ///
    /// Targets are a shape's whole text body, one paragraph, one run, or one
    /// table cell (see [`super::TextTarget`]); properties are run formatting
    /// (bold, italic, size, colour, typefaces), paragraph alignment,
    /// paragraph/run/cell text and cell fill. Untouched markup keeps its
    /// exact bytes. The candidate slide is read back as a complete scene
    /// with unchanged shape identities before the transaction changes. The
    /// return value is the number of edits that changed bytes.
    ///
    /// # Errors
    ///
    /// Returns an error for an out-of-range shape, paragraph, run, row or
    /// cell, a target without a text body, a property that does not apply
    /// to the target, an invalid value, markup-compatibility branch
    /// selection on the slide, or an exhausted limit. A failed batch leaves
    /// the transaction unchanged.
    pub fn set_text_properties<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        edits: &[super::TextAssignment<'_>],
    ) -> Result<usize> {
        let selected_slide = self.resolve_slide(slide.into())?;
        let owner = self.working.get_part(&selected_slide.part_name)?;
        crate::parts::validate_content_type(owner, litchi_opc::constants::content_type::PML_SLIDE)?;
        let (xml, changed) = super::stage_text_properties_with_limits(
            owner.blob(),
            edits,
            self.source_ref()?.limits.max_text_bytes(),
            self.semantic_limits().shape().output_bytes(),
            self.semantic_limits().shape(),
        )?;
        if let Some(xml) = xml {
            self.working
                .get_part_mut(&selected_slide.part_name)?
                .set_blob(xml);
            self.note_scene_read(&selected_slide.part_name);
        }
        Ok(changed)
    }

    /// Apply several already bounded batches of rich text edits to one
    /// slide, in order, in one scene read, rewrite and read-back (record
    /// 0333). Each chunk keeps the limits of [`Self::set_text_properties`]
    /// (256 edits and its text bytes); the total is bounded by 100,000
    /// edits. The return value is the number of edits that changed bytes.
    ///
    /// # Errors
    ///
    /// As [`Self::set_text_properties`], and a limit error for too many
    /// chunks or edits. A failed batch leaves the transaction unchanged.
    pub fn set_text_property_chunks<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        chunks: &[&[super::TextAssignment<'_>]],
    ) -> Result<usize> {
        self.text_property_chunks(slide.into(), chunks, None)
    }

    /// [`Self::set_text_property_chunks`], returning for each edit, in
    /// order, whether it changed bytes (record 0348): the change one
    /// [`Self::set_text_properties`] call per edit, applied one after the
    /// other, would report. A batch planner learns each edit's change here
    /// instead of reading the targets before staging.
    ///
    /// # Errors
    ///
    /// As [`Self::set_text_property_chunks`].
    pub fn set_text_property_chunks_per_edit<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        chunks: &[&[super::TextAssignment<'_>]],
    ) -> Result<Vec<bool>> {
        let mut changes = Vec::new();
        self.text_property_chunks(slide.into(), chunks, Some(&mut changes))?;
        Ok(changes)
    }

    fn text_property_chunks(
        &mut self,
        slide: crate::slide::Selector<'_>,
        chunks: &[&[super::TextAssignment<'_>]],
        changes: Option<&mut Vec<bool>>,
    ) -> Result<usize> {
        let selected_slide = self.resolve_slide(slide)?;
        let owner = self.working.get_part(&selected_slide.part_name)?;
        crate::parts::validate_content_type(owner, litchi_opc::constants::content_type::PML_SLIDE)?;
        let (xml, changed) = super::stage_text_property_chunks_with_limits(
            owner.blob(),
            chunks,
            self.source_ref()?.limits.max_text_bytes(),
            self.semantic_limits().shape().output_bytes(),
            changes,
            self.semantic_limits().shape(),
        )?;
        if let Some(xml) = xml {
            self.working
                .get_part_mut(&selected_slide.part_name)?
                .set_blob(xml);
            self.note_scene_read(&selected_slide.part_name);
        }
        Ok(changed)
    }

    /// Atomically replace text and/or bounds from several already bounded
    /// same-slide property batches in one scene read, rewrite, and read-back.
    /// Each inner slice retains the ordinary per-call property limits.
    /// Chunk count and total selections are limited to 100,000; aggregate
    /// replacement text cannot exceed the default shape output-byte limit.
    pub fn set_shape_property_chunks<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        chunks: &[&[super::ShapePropertiesReplacement<'_>]],
    ) -> Result<usize> {
        let selected_slide = self.resolve_slide(slide.into())?;
        let owner = self.working.get_part(&selected_slide.part_name)?;
        crate::parts::validate_content_type(owner, litchi_opc::constants::content_type::PML_SLIDE)?;
        let (xml, changed) = super::stage_shape_property_chunks_with_limits(
            owner.blob(),
            chunks,
            self.source_ref()?.limits.max_text_bytes(),
            self.semantic_limits().shape().output_bytes(),
            self.semantic_limits().shape(),
        )?;
        if let Some(xml) = xml {
            self.working
                .get_part_mut(&selected_slide.part_name)?
                .set_blob(xml);
            self.note_scene_read(&selected_slide.part_name);
        }
        Ok(changed)
    }

    /// Add a plain `DrawingML` text box without rebuilding the surrounding slide.
    ///
    /// # Errors
    ///
    /// Returns an error for invalid text, geometry, malformed slide XML, or an
    /// exhausted non-visual shape-ID space.
    pub fn add_text_box<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        text: impl AsRef<str>,
        bounds: (i64, i64, i64, i64),
    ) -> Result<u32> {
        let selected = self.resolve_slide(slide.into())?;
        let owner = self.working.get_part(&selected.part_name)?;
        let shape_limits = self.semantic_limits().shape();
        let scene = crate::shape::Scene::read(owner.blob(), shape_limits)?;
        let shape_id = next_shape_id(&scene, "text box")?;
        let conformance = crate::media_parts::document_conformance_with_limits(
            owner.blob(),
            self.semantic_limits()
                .media_parts()
                .constrained_by_source(self.working.read_limits()),
        )?;
        let fragment = text_box_fragment(
            text.as_ref(),
            bounds,
            shape_id,
            conformance,
            self.source_ref()?.limits,
        )?;
        self.append_checked_shape(&selected, shape_id, fragment.as_bytes())?;
        Ok(shape_id)
    }

    /// Add a rectangular preset shape with an optional six-digit sRGB fill.
    ///
    /// # Errors
    ///
    /// Returns an error for invalid geometry/fill, malformed slide XML, or an
    /// exhausted non-visual shape-ID space.
    pub fn add_rectangle<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        bounds: (i64, i64, i64, i64),
        fill_color: Option<&str>,
    ) -> Result<u32> {
        self.add_preset_shape(slide.into(), "rect", "Rectangle", bounds, fill_color)
    }

    /// Add an elliptical preset shape with an optional six-digit sRGB fill.
    ///
    /// # Errors
    ///
    /// Returns an error for invalid geometry/fill, malformed slide XML, or an
    /// exhausted non-visual shape-ID space.
    pub fn add_ellipse<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        bounds: (i64, i64, i64, i64),
        fill_color: Option<&str>,
    ) -> Result<u32> {
        self.add_preset_shape(slide.into(), "ellipse", "Ellipse", bounds, fill_color)
    }

    /// Add one visible picture and its collision-safe inert image part.
    ///
    /// The resource must use an `image/*` content type and a `/ppt/media/`
    /// preferred part name. Its bytes are retained inertly and never decoded.
    ///
    /// # Errors
    ///
    /// Returns an error for invalid resource metadata, empty/oversized data,
    /// invalid geometry, malformed slide XML, or an exhausted shape-ID space.
    pub fn add_picture<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        name: impl AsRef<str>,
        resource: &crate::media_parts::Resource,
        bounds: (i64, i64, i64, i64),
    ) -> Result<u32> {
        let selected = self.resolve_slide(slide.into())?;
        validate_bounds(bounds, "picture")?;
        validate_shape_text(name.as_ref(), self.source_ref()?.limits, "picture name")?;
        if resource.data.is_empty() {
            return Err(invalid(
                "opened-presentation picture resource cannot be empty",
            ));
        }
        if resource.data.len() > self.source_ref()?.limits.max_patch_bytes() {
            return Err(Error::Limit {
                kind: crate::opened::LimitKind::PatchBytes,
                observed: resource.data.len(),
                maximum: self.source_ref()?.limits.max_patch_bytes(),
            });
        }
        if !resource.content_type.starts_with("image/") {
            return Err(invalid(
                "opened-presentation picture resource must use an image content type",
            ));
        }
        let preferred = PackURI::new(&resource.part_name)
            .map_err(|_error| invalid("opened-presentation resource part name is invalid"))?;
        if !preferred.as_str().starts_with("/ppt/media/") {
            return Err(invalid(
                "opened-presentation picture resource must use a /ppt/media/ part name",
            ));
        }
        let owner = self.working.get_part(&selected.part_name)?;
        let shape_limits = self.semantic_limits().shape();
        let scene = crate::shape::Scene::read(owner.blob(), shape_limits)?;
        let shape_id = next_shape_id(&scene, "picture")?;
        let conformance = crate::media_parts::document_conformance_with_limits(
            owner.blob(),
            self.semantic_limits()
                .media_parts()
                .constrained_by_source(self.working.read_limits()),
        )?;
        let reserved: HashSet<_> = self
            .working
            .iter_parts()
            .map(|part| part.partname().clone())
            .collect();
        let part_name = available_transfer_name(&preferred, &reserved)?;
        let mut candidate = self.working.clone();
        candidate.try_add_part(Box::new(BlobPart::new(
            part_name.clone(),
            resource.content_type.clone(),
            resource.data.as_slice().to_vec(),
        )))?;
        let target = part_name.relative_ref(selected.part_name.base_uri());
        let relationship_id = candidate
            .get_part_mut(&selected.part_name)?
            .relate_to(&target, conformance_image_relationship(conformance));
        let fragment = picture_fragment(
            name.as_ref(),
            bounds,
            shape_id,
            &relationship_id,
            conformance,
        );
        let xml = super::xml::append_shape_with_limit(
            candidate.get_part(&selected.part_name)?.blob(),
            fragment.as_bytes(),
            shape_limits.output_bytes(),
        )?;
        ensure_shape_id(&xml, shape_id, "picture", shape_limits)?;
        candidate.get_part_mut(&selected.part_name)?.set_blob(xml);
        self.working = candidate;
        Ok(shape_id)
    }

    /// Remove one top-level shape and any relationship closure made unreachable.
    ///
    /// Relationships still referenced elsewhere on the slide and parts shared
    /// from any package owner are retained.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing/ambiguous selector or malformed shape and
    /// relationship topology.
    pub fn remove_shape<'s, 'k>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        shape: impl Into<crate::shape::Key<'k>>,
    ) -> Result<bool> {
        let selected = self.resolve_slide(slide.into())?;
        let key = shape.into();
        let owner = self.working.get_part(&selected.part_name)?;
        let blob = owner.blob();
        let shape_limits = self.semantic_limits().shape();
        let span = super::xml::shape_removal_span_with_limits(blob, key, shape_limits)?;
        let mut removed_relationships =
            shape_relationship_ids(super::xml::span_bytes(blob, span.clone())?, owner.rels())?;
        // A SmartArt's cached drawing is related by the slide but named only
        // by its data model, so it leaves with the frame (record 0284).
        let drawings: Vec<String> = removed_relationships
            .iter()
            .filter_map(|id| smartart::drawing_relationship(&self.working, &selected.part_name, id))
            .collect();
        removed_relationships.extend(drawings);
        let (xml, _) = super::xml::stage_shape_removal_with_limits(
            blob,
            key,
            self.semantic_limits().shape().output_bytes(),
            shape_limits,
        )?;
        let xml =
            xml.ok_or_else(|| invalid("opened-presentation shape removal produced no candidate"))?;
        let retained_relationships = shape_relationship_ids(&xml, owner.rels())?;
        let mut candidate = self.working.clone();
        let mut dependency_roots = Vec::new();
        {
            let owner = candidate.get_part_mut(&selected.part_name)?;
            owner.set_blob(xml);
            for relationship_id in removed_relationships {
                if retained_relationships.contains(&relationship_id) {
                    continue;
                }
                let Some(relationship) = owner.rels_mut().remove(&relationship_id) else {
                    continue;
                };
                if !relationship.is_external() {
                    dependency_roots.push(relationship.target_partname()?);
                }
            }
        }
        remove_unreferenced_dependencies(&mut candidate, dependency_roots)?;
        self.working = candidate;
        Ok(true)
    }

    /// Copy one common top-level shape and every internal relationship target
    /// reachable from it into a destination slide.
    ///
    /// Part collisions, relationship IDs, every grouped shape identity, and
    /// connector endpoint references are remapped deterministically. Attached
    /// connectors pull in the complete top-level owner of each endpoint,
    /// including grouped children and their dependency relationships. Unknown
    /// shape kinds and unresolved endpoint identities remain explicit refusal
    /// boundaries.
    ///
    /// # Errors
    ///
    /// Returns an error for unsupported shapes, malformed XML/relationships,
    /// missing dependencies, or exceeded transaction bounds.
    pub fn transfer_shape<'ss, 'sk, 'ds>(
        &mut self,
        source: &Snapshot,
        source_slide: impl Into<crate::slide::Selector<'ss>>,
        source_shape: impl Into<crate::shape::Key<'sk>>,
        destination_slide: impl Into<crate::slide::Selector<'ds>>,
    ) -> Result<u32> {
        let source_slide = resolve_snapshot_slide(source, source_slide.into())?;
        let destination = self.resolve_slide(destination_slide.into())?;
        let source_owner = source.package.get_part(&source_slide.part_name)?;
        let source_shape_limits = source
            .owner_state
            .semantic_limits
            .as_deref()
            .copied()
            .unwrap_or_default()
            .shape();
        let source_scene = crate::shape::Scene::read(source_owner.blob(), source_shape_limits)?;
        let key = source_shape.into();
        let shape = source_scene.shape(key)?;
        let source_roots: Vec<_> = source_scene.roots().collect();
        let selected_root = source_roots
            .iter()
            .position(|root| root.common().index() == shape.common().index());
        let Some(selected_root) = selected_root else {
            return Err(shape_transfer_refusal(
                crate::ShapeTransferRefusal::NestedShape,
                "transfer requires the selected shape's top-level owner",
            ));
        };
        let mut root_by_shape_id = HashMap::new();
        for (root_index, root) in source_roots.iter().copied().enumerate() {
            let mut root_ids = BTreeSet::new();
            collect_routable_shape_ids(root, &mut root_ids)?;
            for shape_id in root_ids {
                if root_by_shape_id.insert(shape_id, root_index).is_some() {
                    return Err(invalid(
                        "opened-presentation source slide repeats a shape identity",
                    ));
                }
            }
        }
        let mut pending = VecDeque::from([selected_root]);
        let mut included_roots = BTreeSet::new();
        let mut planned_roots = Vec::new();
        let mut source_shape_ids = BTreeSet::new();
        let mut relationship_ids = BTreeSet::new();
        while let Some(root_index) = pending.pop_front() {
            if !included_roots.insert(root_index) {
                continue;
            }
            if included_roots.len() > self.source_ref()?.limits.max_parts() {
                return Err(Error::Limit {
                    kind: crate::opened::LimitKind::Parts,
                    observed: included_roots.len(),
                    maximum: self.source_ref()?.limits.max_parts(),
                });
            }
            let root = source_roots
                .get(root_index)
                .copied()
                .ok_or_else(|| invalid("opened-presentation transferred shape root disappeared"))?;
            collect_transfer_shape_ids(root, &mut source_shape_ids)?;
            let span = crate::tag::shape::selected_raw_span_with_limits(
                source_owner.blob(),
                crate::shape::Key::Index(root.common().index()),
                source_shape_limits,
            )?;
            let fragment = super::xml::span_bytes(source_owner.blob(), span.clone())?;
            relationship_ids.extend(shape_relationship_ids(fragment, source_owner.rels())?);
            for connected_id in super::xml::connector_connection_ids(fragment)? {
                let endpoint_root =
                    root_by_shape_id
                        .get(&connected_id)
                        .copied()
                        .ok_or_else(|| {
                            shape_transfer_refusal(
                                crate::ShapeTransferRefusal::UnresolvedConnectorEndpoint,
                                "connector endpoint does not identify a transferable source shape",
                            )
                        })?;
                pending.push_back(endpoint_root);
            }
            planned_roots.push((root_index, root, span));
        }
        planned_roots.sort_unstable_by_key(|(root_index, _, _)| *root_index);
        let source_shape_id = shape.id().ok_or_else(|| {
            shape_transfer_refusal(
                crate::ShapeTransferRefusal::MissingIdentity,
                "selected shape has no non-visual identity",
            )
        })?;
        if !source_shape_ids.contains(&source_shape_id) {
            return Err(invalid(
                "opened-presentation transferred root identity is absent from its XML",
            ));
        }
        let mut roots = Vec::new();
        let mut relationships = Vec::new();
        for relationship_id in relationship_ids {
            let relationship = source_owner.rels().get(&relationship_id).ok_or_else(|| {
                invalid("opened-presentation transferred shape relationship is missing")
            })?;
            if !relationship.is_external() {
                roots.push(relationship.target_partname()?);
            }
            relationships.push(relationship.clone());
        }
        let mapping = plan_transfer_roots(
            source.package.as_ref(),
            &self.working,
            roots,
            self.source_ref()?.limits.max_parts(),
        )?;
        let mut candidate = self.working.clone();
        publish_transfer(source.package.as_ref(), &mut candidate, &mapping)?;
        let destination_scene = crate::shape::Scene::read(
            candidate.get_part(&destination.part_name)?.blob(),
            self.semantic_limits().shape(),
        )?;
        let shape_id_mapping = plan_shape_identity_transfer(&destination_scene, &source_shape_ids)?;
        let shape_id = *shape_id_mapping.get(&source_shape_id).ok_or_else(|| {
            invalid("opened-presentation transferred root identity mapping disappeared")
        })?;
        let mut relationship_mapping = HashMap::new();
        for relationship in relationships {
            let target = if relationship.is_external() {
                relationship.target_ref().to_owned()
            } else {
                let source_target = relationship.target_partname()?;
                let copied_target = mapping.get(&source_target).ok_or_else(|| {
                    invalid("opened-presentation transferred shape dependency is missing")
                })?;
                copied_target.relative_ref(destination.part_name.base_uri())
            };
            let new_id = candidate
                .get_part_mut(&destination.part_name)?
                .relate_to(&target, relationship.reltype());
            relationship_mapping.insert(relationship.r_id().to_owned(), new_id);
        }
        let mut xml = candidate.get_part(&destination.part_name)?.blob().to_vec();
        for (_root_index, root, span) in planned_roots {
            let root_source_id = root.id().ok_or_else(|| {
                invalid("opened-presentation transferred root has no non-visual identity")
            })?;
            let root_destination_id = shape_id_mapping.get(&root_source_id).ok_or_else(|| {
                invalid("opened-presentation transferred root identity mapping disappeared")
            })?;
            let name = format!(
                "{} Copy {root_destination_id}",
                root.name().unwrap_or("Shape")
            );
            let fragment = super::xml::remap_shape_fragment(
                super::xml::span_bytes(source_owner.blob(), span)?,
                root_source_id,
                &name,
                &shape_id_mapping,
                &relationship_mapping,
            )?;
            xml = super::xml::append_shape_with_limit(
                &xml,
                &fragment,
                self.semantic_limits().shape().output_bytes(),
            )?;
        }
        ensure_shape_id(
            &xml,
            shape_id,
            "transferred shape",
            self.semantic_limits().shape(),
        )?;
        candidate
            .get_part_mut(&destination.part_name)?
            .set_blob(xml);
        self.working = candidate;
        Ok(shape_id)
    }

    /// Apply a bounded batch of rich text edits to the shapes of one slide's
    /// notes page, in order, as [`Self::set_text_properties`] applies them to
    /// a slide (record 0285): targets address the notes page's own shapes by
    /// their positional scene index. Returns the number of edits whose bytes
    /// changed.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Refused`] for a slide without a notes page, and
    /// the errors of [`Self::set_text_properties`]. A failed batch leaves
    /// the transaction unchanged.
    pub fn set_notes_text_properties<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        edits: &[super::TextAssignment<'_>],
    ) -> Result<usize> {
        let selected_slide = self.resolve_slide(slide.into())?;
        let presentation_name = self.source_ref()?.presentation_name.clone();
        let source = crate::notes::load_snapshot_with_limits(
            &self.working,
            &presentation_name,
            self.notes_limits(),
        )?;
        let Some(note) = source.as_ref().and_then(|source| {
            source
                .slides()
                .iter()
                .find(|note| note.owner() == selected_slide.part_name.as_str())
        }) else {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedTarget,
                "set_notes_text_properties",
                "the slide has no notes page",
            ));
        };
        let part_name = PackURI::new(note.part())
            .map_err(|_error| invalid("opened-presentation notes part name is invalid"))?;
        let (xml, changed) = super::stage_text_properties_with_limits(
            note.xml(),
            edits,
            self.source_ref()?.limits.max_text_bytes(),
            self.semantic_limits().shape().output_bytes(),
            self.semantic_limits().shape(),
        )?;
        if let Some(xml) = xml {
            self.working.get_part_mut(&part_name)?.set_blob(xml);
        }
        Ok(changed)
    }

    /// Apply several already bounded batches of rich text edits to the
    /// notes page of one slide, in order, in one notes-part read, rewrite
    /// and read-back (record 0348), as [`Self::set_text_property_chunks`]
    /// does for a slide. The result equals one
    /// [`Self::set_notes_text_properties`] call per chunk, in order. The
    /// return value is the number of edits that changed bytes.
    ///
    /// # Errors
    ///
    /// As [`Self::set_notes_text_properties`], and a limit error for too
    /// many chunks or edits. A failed batch leaves the transaction
    /// unchanged.
    pub fn set_notes_text_property_chunks<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        chunks: &[&[super::TextAssignment<'_>]],
    ) -> Result<usize> {
        self.notes_text_property_chunks(slide.into(), chunks, None)
    }

    /// [`Self::set_notes_text_property_chunks`], returning each edit's
    /// change in order (record 0348), as
    /// [`Self::set_text_property_chunks_per_edit`] does for a slide.
    ///
    /// # Errors
    ///
    /// As [`Self::set_notes_text_property_chunks`].
    pub fn set_notes_text_property_chunks_per_edit<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        chunks: &[&[super::TextAssignment<'_>]],
    ) -> Result<Vec<bool>> {
        let mut changes = Vec::new();
        self.notes_text_property_chunks(slide.into(), chunks, Some(&mut changes))?;
        Ok(changes)
    }

    fn notes_text_property_chunks(
        &mut self,
        slide: crate::slide::Selector<'_>,
        chunks: &[&[super::TextAssignment<'_>]],
        changes: Option<&mut Vec<bool>>,
    ) -> Result<usize> {
        let selected_slide = self.resolve_slide(slide)?;
        let presentation_name = self.source_ref()?.presentation_name.clone();
        let source = crate::notes::load_snapshot_with_limits(
            &self.working,
            &presentation_name,
            self.notes_limits(),
        )?;
        let Some(note) = source.as_ref().and_then(|source| {
            source
                .slides()
                .iter()
                .find(|note| note.owner() == selected_slide.part_name.as_str())
        }) else {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedTarget,
                "set_notes_text_properties",
                "the slide has no notes page",
            ));
        };
        let part_name = PackURI::new(note.part())
            .map_err(|_error| invalid("opened-presentation notes part name is invalid"))?;
        let (xml, changed) = super::stage_text_property_chunks_with_limits(
            note.xml(),
            chunks,
            self.source_ref()?.limits.max_text_bytes(),
            self.semantic_limits().shape().output_bytes(),
            changes,
            self.semantic_limits().shape(),
        )?;
        if let Some(xml) = xml {
            self.working.get_part_mut(&part_name)?.set_blob(xml);
        }
        Ok(changed)
    }

    /// Replace the notes text owned by one checked slide, creating the
    /// slide's notes page from the presentation's notes master when it has
    /// none (an empty text on a slide without notes is a no-op).
    ///
    /// # Errors
    ///
    /// Returns [`Error::Refused`] when a notes page must be created but
    /// the presentation has no single notes master or is strict, and an
    /// error when the notes graph has an unsupported dependency or the
    /// replacement is malformed or unbounded. A failure leaves the edit
    /// unchanged.
    pub fn set_notes_text<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        text: impl AsRef<str>,
    ) -> Result<bool> {
        let selected_slide = self.resolve_slide(slide.into())?;
        let text = text.as_ref();
        if text.len() > self.source_ref()?.limits.max_text_bytes() {
            return Err(Error::Limit {
                kind: crate::opened::LimitKind::TextBytes,
                observed: text.len(),
                maximum: self.source_ref()?.limits.max_text_bytes(),
            });
        }
        let presentation_name = self.source_ref()?.presentation_name.clone();
        let source = crate::notes::load_snapshot_with_limits(
            &self.working,
            &presentation_name,
            self.notes_limits(),
        )?;
        let note = source.as_ref().and_then(|source| {
            source
                .slides()
                .iter()
                .find(|note| note.owner() == selected_slide.part_name.as_str())
        });
        match note {
            Some(note) => {
                let limits = self.notes_limits();
                if note.text(limits)?.as_deref() == (!text.is_empty()).then_some(text) {
                    return Ok(false);
                }
                let part_name = PackURI::new(note.part())
                    .map_err(|_error| invalid("opened-presentation notes part name is invalid"))?;
                let xml = crate::notes::rewrite_text(note.xml(), text, limits)?;
                self.working.get_part_mut(&part_name)?.set_blob(xml);
            },
            // A slide without notes and empty text is already as requested.
            None if text.is_empty() => return Ok(false),
            None => {
                let slide_part = selected_slide.part_name.clone();
                let before = self.working.clone();
                let created = self
                    .create_notes_slide(&slide_part, text)
                    .and_then(|()| self.verify_notes_text(&slide_part, text));
                if let Err(error) = created {
                    self.working = before;
                    return Err(error);
                }
                return Ok(true);
            },
        }
        self.verify_notes_text(&selected_slide.part_name, text)?;
        Ok(true)
    }

    /// Read the staged notes of `slide` back and require `text`.
    fn verify_notes_text(&self, slide: &PackURI, text: &str) -> Result<()> {
        let presentation_name = self.source_ref()?.presentation_name.clone();
        let staged = crate::notes::load_snapshot_with_limits(
            &self.working,
            &presentation_name,
            self.notes_limits(),
        )?
        .ok_or_else(|| invalid("opened-presentation staged notes graph disappeared"))?;
        let staged_note = staged
            .slides()
            .iter()
            .find(|note| note.owner() == slide.as_str())
            .ok_or_else(|| invalid("opened-presentation staged notes owner disappeared"))?;
        let actual = staged_note.text(self.notes_limits())?;
        if actual.as_deref() != (!text.is_empty()).then_some(text) {
            return Err(invalid(
                "opened-presentation notes text did not round-trip semantically",
            ));
        }
        Ok(())
    }

    /// Create or replace the validated table-style catalog in this transaction.
    ///
    /// Staging a change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006): remove the signature explicitly with
    /// [`Package::unsign`](crate::Package::unsign) before the transaction
    /// starts.
    ///
    /// # Errors
    ///
    /// Returns an error when the catalog or its attachment graph is invalid.
    pub fn put_table_styles(&mut self, styles: crate::table::style::List) -> Result<bool> {
        self.ensure_ready()?;
        crate::table::style::put(&mut self.working, styles)
    }

    /// Remove the optional table-style catalog and its relationship.
    ///
    /// Staging a change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006): remove the signature explicitly with
    /// [`Package::unsign`](crate::Package::unsign) before the transaction
    /// starts.
    ///
    /// # Errors
    ///
    /// Returns an error when the existing attachment graph is unsafe.
    pub fn remove_table_styles(&mut self) -> Result<Option<crate::table::style::List>> {
        self.ensure_ready()?;
        crate::table::style::remove(&mut self.working)
    }

    /// Add a visible rectangular table with uniformly sized rows and columns.
    ///
    /// Cell text is XML-escaped, row widths must be rectangular, and geometry
    /// must use positive extents. The surrounding producer XML is preserved.
    ///
    /// # Errors
    ///
    /// Returns an error for empty/ragged input, invalid XML characters,
    /// exceeded text bounds, malformed slide XML, or exhausted shape IDs.
    pub fn add_table<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        cells: &[Vec<String>],
        bounds: (i64, i64, i64, i64),
    ) -> Result<u32> {
        let selected = self.resolve_slide(slide.into())?;
        let owner = self.working.get_part(&selected.part_name)?;
        let shape_limits = self.semantic_limits().shape();
        let scene = crate::shape::Scene::read(owner.blob(), shape_limits)?;
        let shape_id = next_shape_id(&scene, "table")?;
        let conformance = crate::media_parts::document_conformance_with_limits(
            owner.blob(),
            self.semantic_limits()
                .media_parts()
                .constrained_by_source(self.working.read_limits()),
        )?;
        let fragment = table_fragment(
            cells,
            bounds,
            shape_id,
            conformance,
            self.source_ref()?.limits,
        )?;
        let _table = crate::table::Table::from_graphic_frame(fragment.as_bytes())?;
        let xml = super::xml::append_shape_with_limit(
            owner.blob(),
            fragment.as_bytes(),
            shape_limits.output_bytes(),
        )?;
        let staged = crate::shape::Scene::read(&xml, shape_limits)?;
        if !staged.iter().any(|shape| shape.id() == Some(shape_id)) {
            return Err(invalid(
                "opened-presentation table did not round-trip semantically",
            ));
        }
        self.working
            .get_part_mut(&selected.part_name)?
            .set_blob(xml);
        Ok(shape_id)
    }

    /// Add an authored chart: its chart part, an embedded workbook holding
    /// the chart data, both relationships, and a visible graphic frame.
    ///
    /// The chart refers to the workbook's `Sheet1` in the layout of
    /// [`Spec::sheet_layout`](litchi_drawingml::chart::authoring::Spec::sheet_layout),
    /// whatever binding `chart` carries, so PowerPoint can edit the data.
    /// `bounds` is `(x, y, width, height)` in EMU. The parts are named
    /// `/ppt/charts/chartN.xml` and `/ppt/embeddings/Microsoft_Excel_WorksheetN.xlsx`
    /// with the first free `N`, and both are generated deterministically.
    /// Returns the new frame's shape ID.
    ///
    /// Staging a change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006): remove the signature explicitly with
    /// [`Package::unsign`](crate::Package::unsign) before the transaction
    /// starts.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing slide, an invalid chart spec,
    /// non-positive extents, invalid name text, a malformed shape tree, or
    /// exhausted shape IDs.
    pub fn add_chart<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        chart: &litchi_drawingml::chart::authoring::Spec,
        name: &str,
        bounds: (i64, i64, i64, i64),
    ) -> Result<u32> {
        use litchi_drawingml::chart::authoring::Binding;
        use litchi_drawingml::chart::chartex;
        use litchi_opc::constants::{content_type as ct, relationship_type as rt};

        self.ensure_ready()?;
        if bounds.2 <= 0 || bounds.3 <= 0 {
            return Err(invalid("chart frame extents must be positive"));
        }
        if name.contains('\0') || name.len() > 32_767 {
            return Err(invalid("chart frame name is not valid text"));
        }
        let selected = self.resolve_slide(slide.into())?;
        let spec = chart.clone().with_binding(Binding::sheet("Sheet1"));
        let extended = spec.variant().kind.is_extended();
        let workbook =
            crate::presentation::embedded::excel::Workbook::from_spec(&spec)?.into_bytes();
        let mut candidate = self.working.clone();
        crate::package::signature::refuse_signed_change(&candidate)?;
        let shape_limits = self.semantic_limits().shape();
        let scene = crate::shape::Scene::read(
            candidate.get_part(&selected.part_name)?.blob(),
            shape_limits,
        )?;
        let shape_id = next_shape_id(&scene, "chart")?;
        let (chart_stem, content_type, relationship_type) = if extended {
            (
                "/ppt/charts/chartEx",
                chartex::CONTENT_TYPE,
                chartex::RELATIONSHIP_TYPE,
            )
        } else {
            ("/ppt/charts/chart", ct::DML_CHART, rt::CHART)
        };
        let chart_name = free_part_name(&candidate, chart_stem, ".xml")?;
        let workbook_name = free_part_name(
            &candidate,
            "/ppt/embeddings/Microsoft_Excel_Worksheet",
            ".xlsx",
        )?;
        let mut chart_part = BlobPart::new(chart_name.clone(), content_type.to_owned(), Vec::new());
        let workbook_ref = workbook_name.relative_ref(chart_name.base_uri());
        let workbook_rel = chart_part.relate_to(&workbook_ref, rt::PACKAGE);
        if extended {
            // An extended chart needs its style and color companions.
            for (stem, content_type, relationship_type, bytes) in [
                (
                    "/ppt/charts/style",
                    chartex::STYLE_CONTENT_TYPE,
                    chartex::STYLE_RELATIONSHIP_TYPE,
                    chartex::style_xml(),
                ),
                (
                    "/ppt/charts/colors",
                    chartex::COLORS_CONTENT_TYPE,
                    chartex::COLORS_RELATIONSHIP_TYPE,
                    chartex::colors_xml(),
                ),
            ] {
                let companion = free_part_name(&candidate, stem, ".xml")?;
                let reference = companion.relative_ref(chart_name.base_uri());
                chart_part.relate_to(&reference, relationship_type);
                candidate.add_part(Box::new(BlobPart::new(
                    companion,
                    content_type.to_owned(),
                    bytes,
                )));
            }
        }
        let xml = spec
            .to_xml(Some(&workbook_rel))
            .map_err(|error| invalid(format!("authored chart: {error}")))?;
        chart_part.set_blob(xml);
        candidate.add_part(Box::new(BlobPart::new(
            workbook_name,
            ct::SML_SHEET.to_owned(),
            workbook,
        )));
        candidate.add_part(Box::new(chart_part));
        let chart_ref = chart_name.relative_ref(selected.part_name.base_uri());
        let relationship_id = candidate
            .get_part_mut(&selected.part_name)?
            .relate_to(&chart_ref, relationship_type);
        let fragment = chart_frame(shape_id, &relationship_id, name, bounds, extended);
        let owner = candidate.get_part(&selected.part_name)?;
        let xml = super::xml::append_shape_with_limit(
            owner.blob(),
            fragment.as_bytes(),
            shape_limits.output_bytes(),
        )?;
        let scene = crate::shape::Scene::read(&xml, shape_limits)?;
        if !scene.iter().any(|shape| shape.id() == Some(shape_id)) {
            return Err(invalid(
                "opened-presentation chart frame did not round-trip semantically",
            ));
        }
        candidate.get_part_mut(&selected.part_name)?.set_blob(xml);
        crate::package::signature::settle_unsigned_change(&mut candidate)?;
        self.working = candidate;
        Ok(shape_id)
    }

    /// The chart-space XML of the chart frame at a depth-first shape
    /// position of a slide, with the bytes of its embedded workbook when the
    /// chart has one.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing slide or shape, a shape that is not a
    /// chart frame, or a broken chart relationship.
    pub fn chart<'s, 'k>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        shape: impl Into<crate::shape::Key<'k>>,
    ) -> Result<(Vec<u8>, Option<Vec<u8>>)> {
        let selected = self.resolve_slide(slide.into())?;
        let (chart_name, workbook_name) = chart_targets(
            &self.working,
            &selected.part_name,
            shape.into(),
            self.semantic_limits().shape(),
        )?;
        let xml = self.working.get_part(&chart_name)?.blob().to_vec();
        let workbook = match workbook_name {
            Some(name) => Some(self.working.get_part(&name)?.blob().to_vec()),
            None => None,
        };
        Ok((xml, workbook))
    }

    /// Replace the chart-space XML of the chart frame at a depth-first shape
    /// position of a slide and, when `workbook` is given, the bytes of the
    /// embedded workbook the chart already refers to. The frame, the chart's
    /// relationships and every other part are kept.
    ///
    /// The new XML must read back as a chart. Returns whether anything
    /// changed; identical bytes are an exact no-op.
    ///
    /// Staging a change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006).
    ///
    /// # Errors
    ///
    /// Returns an error for a missing slide or shape, a shape that is not a
    /// chart frame, XML that is not a chart, or a workbook for a chart that
    /// has none.
    pub fn set_chart<'s, 'k>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        shape: impl Into<crate::shape::Key<'k>>,
        xml: Vec<u8>,
        workbook: Option<Vec<u8>>,
    ) -> Result<bool> {
        let selected = self.resolve_slide(slide.into())?;
        let (chart_name, workbook_name) = chart_targets(
            &self.working,
            &selected.part_name,
            shape.into(),
            self.semantic_limits().shape(),
        )?;
        let current = self.working.get_part(&chart_name)?.content_type();
        if current == litchi_drawingml::chart::chartex::CONTENT_TYPE {
            litchi_drawingml::chart::chartex::read(&xml, &litchi_drawingml::chart::Limits::DEFAULT)
                .map_err(|error| drawing_error(error, "replacement chart XML"))?;
        } else {
            litchi_drawingml::chart::reader::read(
                &xml[..],
                &litchi_drawingml::chart::Limits::DEFAULT,
            )
            .map_err(|error| drawing_error(error, "replacement chart XML"))?;
        }
        let chart_changed = self.working.get_part(&chart_name)?.blob() != xml.as_slice();
        let workbook_change = match (workbook, workbook_name) {
            (None, _) => None,
            (Some(_), None) => {
                return Err(invalid("the chart has no embedded workbook to replace"));
            },
            (Some(bytes), Some(name)) => {
                (self.working.get_part(&name)?.blob() != bytes.as_slice()).then_some((name, bytes))
            },
        };
        if !chart_changed && workbook_change.is_none() {
            return Ok(false);
        }
        let mut candidate = self.working.clone();
        crate::package::signature::refuse_signed_change(&candidate)?;
        if chart_changed {
            candidate.get_part_mut(&chart_name)?.set_blob(xml);
        }
        if let Some((name, bytes)) = workbook_change {
            candidate.get_part_mut(&name)?.set_blob(bytes);
        }
        crate::package::signature::settle_unsigned_change(&mut candidate)?;
        self.working = candidate;
        Ok(true)
    }

    /// Store bounded slide media, its visible picture frames, and inert resources.
    ///
    /// # Errors
    ///
    /// Returns an error when the slide already contains media or the typed
    /// media graph is invalid.
    pub fn store_media<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        media: &crate::media_parts::List,
        conformance: crate::media_parts::Conformance,
    ) -> Result<()> {
        let selected = self.resolve_slide(slide.into())?;
        let limits = self
            .semantic_limits()
            .media_parts()
            .constrained_by_source(self.working.read_limits());
        crate::media_parts::package::store_with_limits(
            &mut self.working,
            &selected.part_name,
            media,
            conformance,
            false,
            limits,
        )
    }

    /// Add a slide master and all required theme relationships.
    ///
    /// Staging a change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006): remove the signature explicitly with
    /// [`Package::unsign`](crate::Package::unsign) before the transaction
    /// starts.
    ///
    /// # Errors
    ///
    /// Returns an error when the master/layout graph cannot be extended safely.
    pub fn add_slide_master(&mut self) -> Result<crate::master_layout::AuthoredSlideMaster> {
        self.ensure_ready()?;
        crate::master_layout::validate_master_layout_graph(&self.working)?;
        crate::master_layout::add_slide_master(&mut self.working)
    }

    /// Add a typed slide layout to an existing master.
    ///
    /// Staging a change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006): remove the signature explicitly with
    /// [`Package::unsign`](crate::Package::unsign) before the transaction
    /// starts.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid master, name, placeholder set, or graph.
    pub fn add_slide_layout(
        &mut self,
        master: &PackURI,
        kind: crate::master_layout::SlideLayoutKind,
        name: &str,
        placeholders: &[crate::master_layout::PlaceholderSpec],
    ) -> Result<crate::master_layout::AuthoredSlideLayout> {
        self.ensure_ready()?;
        crate::master_layout::validate_master_layout_graph(&self.working)?;
        crate::master_layout::add_slide_layout(&mut self.working, master, kind, name, placeholders)
    }

    /// Remove an unreferenced slide layout and its exact graph edges.
    ///
    /// Staging a change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006): remove the signature explicitly with
    /// [`Package::unsign`](crate::Package::unsign) before the transaction
    /// starts.
    ///
    /// # Errors
    ///
    /// Returns an error when the layout is still used or its graph is malformed.
    pub fn remove_slide_layout(&mut self, layout: &PackURI) -> Result<()> {
        self.ensure_ready()?;
        crate::master_layout::validate_master_layout_graph(&self.working)?;
        crate::master_layout::remove_slide_layout(&mut self.working, layout)
    }

    /// Add a legacy comment author to the presentation comment graph.
    ///
    /// Staging a change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006): remove the signature explicitly with
    /// [`Package::unsign`](crate::Package::unsign) before the transaction
    /// starts.
    ///
    /// # Errors
    ///
    /// Returns an error when the author or existing comment graph is invalid.
    pub fn add_comment_author(
        &mut self,
        author: crate::comments::Author,
        conformance: crate::comments::Conformance,
    ) -> Result<()> {
        self.ensure_ready()?;
        let limits = self.comments_limits();
        crate::comments::add_presentation_comment_author_with_limits(
            &mut self.working,
            author,
            conformance,
            limits,
        )
    }

    /// Add a legacy comment owned by one selected slide.
    ///
    /// Staging a change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006): remove the signature explicitly with
    /// [`Package::unsign`](crate::Package::unsign) before the transaction
    /// starts.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing author, duplicate identity, or invalid graph.
    pub fn add_comment<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        comment: crate::comments::Comment,
        conformance: crate::comments::Conformance,
    ) -> Result<()> {
        let selected = self.resolve_slide(slide.into())?;
        let limits = self.comments_limits();
        crate::comments::add_presentation_comment_with_limits(
            &mut self.working,
            selected.part_name.as_str(),
            comment,
            conformance,
            limits,
        )
    }

    /// Remove one legacy comment by stable author/index identity.
    ///
    /// Staging a change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006): remove the signature explicitly with
    /// [`Package::unsign`](crate::Package::unsign) before the transaction
    /// starts.
    ///
    /// # Errors
    ///
    /// Returns an error when the comment graph is malformed.
    pub fn remove_comment<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        author_id: u32,
        index: u32,
        conformance: crate::comments::Conformance,
    ) -> Result<bool> {
        let selected = self.resolve_slide(slide.into())?;
        let limits = self.comments_limits();
        crate::comments::remove_presentation_comment_with_limits(
            &mut self.working,
            selected.part_name.as_str(),
            author_id,
            index,
            conformance,
            limits,
        )
    }

    /// Add a typed `PowerPoint` 2018 modern-comment author.
    ///
    /// A missing author part and presentation relationship are allocated
    /// collision-safely by the modern-comment owner.
    ///
    /// Staging a change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006): remove the signature explicitly with
    /// [`Package::unsign`](crate::Package::unsign) before the transaction
    /// starts.
    ///
    /// # Errors
    ///
    /// Returns an error for duplicate/invalid authors or a malformed graph.
    pub fn add_modern_comment_author(
        &mut self,
        author: crate::modern_comments::Author,
    ) -> Result<crate::modern_comments::AuthorPart> {
        self.ensure_ready()?;
        let mut candidate = self.working.clone();
        let result = crate::modern_comments::add_modern_comment_author(&mut candidate, author)?;
        crate::modern_comments::load_modern_comment_graph(&candidate)?;
        self.working = candidate;
        Ok(result)
    }

    /// Replace a modern-comment author without changing its stable GUID.
    ///
    /// Staging a change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006): remove the signature explicitly with
    /// [`Package::unsign`](crate::Package::unsign) before the transaction
    /// starts.
    ///
    /// # Errors
    ///
    /// Returns an error for an identity mismatch, unresolved references, or a
    /// malformed modern-comment graph.
    pub fn replace_modern_comment_author(
        &mut self,
        author_id: &str,
        replacement: crate::modern_comments::Author,
    ) -> Result<bool> {
        self.ensure_ready()?;
        let mut candidate = self.working.clone();
        let result = crate::modern_comments::replace_modern_comment_author(
            &mut candidate,
            author_id,
            replacement,
        )?;
        crate::modern_comments::load_modern_comment_graph(&candidate)?;
        self.working = candidate;
        Ok(result)
    }

    /// Remove one unreferenced modern-comment author.
    ///
    /// Staging a change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006): remove the signature explicitly with
    /// [`Package::unsign`](crate::Package::unsign) before the transaction
    /// starts.
    ///
    /// # Errors
    ///
    /// Returns an error while any modeled comment/reply still references the
    /// author or when the graph is malformed.
    pub fn remove_modern_comment_author(&mut self, author_id: &str) -> Result<bool> {
        self.ensure_ready()?;
        let mut candidate = self.working.clone();
        let result =
            crate::modern_comments::remove_modern_comment_author(&mut candidate, author_id)?;
        crate::modern_comments::load_modern_comment_graph(&candidate)?;
        self.working = candidate;
        Ok(result)
    }

    /// Reorder all modern-comment authors by a complete stable-GUID list.
    ///
    /// Staging a change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006): remove the signature explicitly with
    /// [`Package::unsign`](crate::Package::unsign) before the transaction
    /// starts.
    ///
    /// # Errors
    ///
    /// Returns an error for unknown, duplicate, or incomplete identities.
    pub fn reorder_modern_comment_authors(
        &mut self,
        ordered_author_ids: &[String],
    ) -> Result<Vec<crate::modern_comments::Author>> {
        self.ensure_ready()?;
        let mut candidate = self.working.clone();
        let result = crate::modern_comments::reorder_modern_comment_authors(
            &mut candidate,
            ordered_author_ids,
        )?;
        crate::modern_comments::load_modern_comment_graph(&candidate)?;
        self.working = candidate;
        Ok(result)
    }

    /// Add a typed `PowerPoint` 2018 modern comment to one selected slide.
    ///
    /// Staging a change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006): remove the signature explicitly with
    /// [`Package::unsign`](crate::Package::unsign) before the transaction
    /// starts.
    ///
    /// # Errors
    ///
    /// Returns an error for unresolved authors, duplicate identities, invalid
    /// extension payloads, or a malformed package graph.
    pub fn add_modern_comment<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        comment: crate::modern_comments::Comment,
    ) -> Result<crate::modern_comments::Part> {
        let selected = self.resolve_slide(slide.into())?;
        let mut candidate = self.working.clone();
        let result = crate::modern_comments::add_modern_comment(
            &mut candidate,
            &selected.part_name,
            comment,
        )?;
        crate::modern_comments::load_modern_comment_graph(&candidate)?;
        self.working = candidate;
        Ok(result)
    }

    /// Replace one modern comment without changing its stable GUID.
    ///
    /// Staging a change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006): remove the signature explicitly with
    /// [`Package::unsign`](crate::Package::unsign) before the transaction
    /// starts.
    ///
    /// # Errors
    ///
    /// Returns an error for an identity mismatch, invalid author/extension
    /// references, or a malformed package graph.
    pub fn replace_modern_comment<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        comment_id: &str,
        replacement: crate::modern_comments::Comment,
    ) -> Result<bool> {
        let selected = self.resolve_slide(slide.into())?;
        let mut candidate = self.working.clone();
        let result = crate::modern_comments::replace_modern_comment(
            &mut candidate,
            &selected.part_name,
            comment_id,
            replacement,
        )?;
        crate::modern_comments::load_modern_comment_graph(&candidate)?;
        self.working = candidate;
        Ok(result)
    }

    /// Reorder all modern comments on one slide by a complete stable-GUID list.
    ///
    /// Staging a change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006): remove the signature explicitly with
    /// [`Package::unsign`](crate::Package::unsign) before the transaction
    /// starts.
    ///
    /// # Errors
    ///
    /// Returns an error for unknown, duplicate, or incomplete identities.
    pub fn reorder_modern_comments<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        ordered_comment_ids: &[String],
    ) -> Result<Vec<crate::modern_comments::Comment>> {
        let selected = self.resolve_slide(slide.into())?;
        let mut candidate = self.working.clone();
        let result = crate::modern_comments::reorder_modern_comments(
            &mut candidate,
            &selected.part_name,
            ordered_comment_ids,
        )?;
        crate::modern_comments::load_modern_comment_graph(&candidate)?;
        self.working = candidate;
        Ok(result)
    }

    /// Add a stable-ID reply to one modern-comment thread.
    ///
    /// Staging a change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006): remove the signature explicitly with
    /// [`Package::unsign`](crate::Package::unsign) before the transaction
    /// starts.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing thread/author, duplicate reply identity,
    /// invalid extensions, or a malformed package graph.
    pub fn add_modern_comment_reply<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        comment_id: &str,
        reply: crate::modern_comments::Reply,
    ) -> Result<bool> {
        let selected = self.resolve_slide(slide.into())?;
        let mut candidate = self.working.clone();
        let result = crate::modern_comments::add_modern_comment_reply(
            &mut candidate,
            &selected.part_name,
            comment_id,
            reply,
        )?;
        crate::modern_comments::load_modern_comment_graph(&candidate)?;
        self.working = candidate;
        Ok(result)
    }

    /// Replace one modern-comment reply without changing its stable GUID.
    ///
    /// Staging a change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006): remove the signature explicitly with
    /// [`Package::unsign`](crate::Package::unsign) before the transaction
    /// starts.
    ///
    /// # Errors
    ///
    /// Returns an error for an identity mismatch, unresolved author, invalid
    /// extensions, or a malformed graph.
    pub fn replace_modern_comment_reply<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        comment_id: &str,
        reply_id: &str,
        replacement: crate::modern_comments::Reply,
    ) -> Result<bool> {
        let selected = self.resolve_slide(slide.into())?;
        let mut candidate = self.working.clone();
        let result = crate::modern_comments::replace_modern_comment_reply(
            &mut candidate,
            &selected.part_name,
            comment_id,
            reply_id,
            replacement,
        )?;
        crate::modern_comments::load_modern_comment_graph(&candidate)?;
        self.working = candidate;
        Ok(result)
    }

    /// Update typed task/reaction extensions on one modern comment while
    /// retaining opaque extension entries.
    ///
    /// Staging a change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006): remove the signature explicitly with
    /// [`Package::unsign`](crate::Package::unsign) before the transaction
    /// starts.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing/malformed thread or invalid extension.
    pub fn update_modern_comment_extensions<'s, F>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        comment_id: &str,
        update: F,
    ) -> Result<bool>
    where
        F: FnOnce(&mut crate::modern_comments::semantic::extensions::List),
    {
        let selected = self.resolve_slide(slide.into())?;
        let mut candidate = self.working.clone();
        let result = crate::modern_comments::update_modern_comment_extensions(
            &mut candidate,
            &selected.part_name,
            comment_id,
            update,
        )?;
        crate::modern_comments::load_modern_comment_graph(&candidate)?;
        self.working = candidate;
        Ok(result)
    }

    /// Update typed reaction extensions on one modern-comment reply while
    /// retaining opaque extension entries.
    ///
    /// Staging a change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006): remove the signature explicitly with
    /// [`Package::unsign`](crate::Package::unsign) before the transaction
    /// starts.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing/malformed reply or invalid extension.
    pub fn update_modern_comment_reply_extensions<'s, F>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        comment_id: &str,
        reply_id: &str,
        update: F,
    ) -> Result<bool>
    where
        F: FnOnce(&mut crate::modern_comments::semantic::extensions::List),
    {
        let selected = self.resolve_slide(slide.into())?;
        let mut candidate = self.working.clone();
        let result = crate::modern_comments::update_modern_comment_reply_extensions(
            &mut candidate,
            &selected.part_name,
            comment_id,
            reply_id,
            update,
        )?;
        crate::modern_comments::load_modern_comment_graph(&candidate)?;
        self.working = candidate;
        Ok(result)
    }

    /// Remove one modern comment and its empty per-slide part when unshared.
    ///
    /// Staging a change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006): remove the signature explicitly with
    /// [`Package::unsign`](crate::Package::unsign) before the transaction
    /// starts.
    ///
    /// # Errors
    ///
    /// Returns an error when the modern-comment graph is malformed.
    pub fn remove_modern_comment<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        comment_id: &str,
    ) -> Result<bool> {
        let selected = self.resolve_slide(slide.into())?;
        let mut candidate = self.working.clone();
        let result = crate::modern_comments::remove_modern_comment(
            &mut candidate,
            &selected.part_name,
            comment_id,
        )?;
        crate::modern_comments::load_modern_comment_graph(&candidate)?;
        self.working = candidate;
        Ok(result)
    }

    /// Remove one modern-comment reply by stable identity.
    ///
    /// Staging a change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006): remove the signature explicitly with
    /// [`Package::unsign`](crate::Package::unsign) before the transaction
    /// starts.
    ///
    /// # Errors
    ///
    /// Returns an error when the modern-comment graph is malformed.
    pub fn remove_modern_comment_reply<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        comment_id: &str,
        reply_id: &str,
    ) -> Result<bool> {
        let selected = self.resolve_slide(slide.into())?;
        let mut candidate = self.working.clone();
        let result = crate::modern_comments::remove_modern_comment_reply(
            &mut candidate,
            &selected.part_name,
            comment_id,
            reply_id,
        )?;
        crate::modern_comments::load_modern_comment_graph(&candidate)?;
        self.working = candidate;
        Ok(result)
    }

    /// Copy one internal relationship target and its complete dependency closure.
    ///
    /// Part-name collisions are deterministically remapped, relationship IDs
    /// inside copied parts remain stable, and every internal edge is retargeted
    /// to the copied dependency. The returned relationship ID is attached to
    /// `destination_slide`; callers that transfer a visually referenced object
    /// must also use an object-specific API such as [`Self::add_chart`] to add
    /// its slide XML anchor.
    ///
    /// # Errors
    ///
    /// Returns an error for external roots, missing dependencies, fragments or
    /// queries on internal targets, or exceeded transaction bounds.
    pub fn transfer_relationship_closure<'s>(
        &mut self,
        source: &Snapshot,
        source_owner: &PackURI,
        relationship_id: &str,
        destination_slide: impl Into<crate::slide::Selector<'s>>,
    ) -> Result<String> {
        let destination = self.resolve_slide(destination_slide.into())?;
        let owner = source.package.get_part(source_owner)?;
        let relationship = owner
            .rels()
            .get(relationship_id)
            .ok_or_else(|| invalid("opened-presentation transfer relationship is missing"))?;
        if relationship.is_external() {
            return Err(invalid(
                "opened-presentation transfer root relationship is external",
            ));
        }
        let root = relationship.target_partname()?;
        let mapping = plan_transfer(
            source.package.as_ref(),
            &self.working,
            &root,
            self.source_ref()?.limits.max_parts(),
        )?;
        let mut candidate = self.working.clone();
        publish_transfer(source.package.as_ref(), &mut candidate, &mapping)?;
        let copied_root = mapping
            .get(&root)
            .ok_or_else(|| invalid("opened-presentation transfer root disappeared"))?;
        let target = copied_root.relative_ref(destination.part_name.base_uri());
        let id = candidate
            .get_part_mut(&destination.part_name)?
            .relate_to(&target, relationship.reltype());
        self.working = candidate;
        Ok(id)
    }

    /// Validate and consume all staged edits into one atomic commit.
    ///
    /// A changed commit invalidates every package signature, so a changed
    /// commit over a signed package is refused (ADR 0006): the signature must
    /// be removed explicitly first, with
    /// [`Package::unsign`](crate::Package::unsign) before the transaction
    /// starts. An exact no-op commits the unchanged source, signature
    /// included.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Opc`] with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// for a changed commit over a signed package, and an error if a
    /// dependency changed, the patch exceeds bounds, or the complete staged
    /// package cannot be captured as a coherent snapshot.
    pub fn commit(mut self) -> Result<Commit> {
        self.ensure_ready()?;
        let source = self
            .source
            .take()
            .ok_or_else(|| invalid("opened-presentation edit source state is unavailable"))?;
        let planned = self.planned.take();
        if let Some((snapshot, patch)) = planned
            && super::model::packages_equal(&self.working, snapshot.package.as_ref())
        {
            // The staged graph is exactly the candidate a plan published as
            // this edit's first change: its validated snapshot and patch stand.
            let working = self.working.into_owned();
            let changed = patch.resources().cloned().collect::<Vec<_>>();
            let base = source.replay_index.clone().map_or_else(
                || {
                    crate::package::replay_index::ReplayIndex::from_package(source.package.as_ref())
                        .map(Arc::new)
                },
                Ok,
            )?;
            let mut snapshot = snapshot;
            snapshot.replay_index = Some(Arc::new(base.update(
                source.package.as_ref(),
                &working,
                &changed,
            )?));
            let package = crate::Package::from_opened_candidate(working, &snapshot)?;
            return Ok(Commit::new(package, snapshot, patch));
        }
        let shape_limits = self.semantic_limits().shape();
        let mut working = self.working;
        compact_changed_slides_with_limits(
            &mut working,
            source.package.as_ref(),
            &self.slides,
            &self.scene_read,
            shape_limits,
        )?;
        // The staged package shares every untouched payload allocation with
        // the source snapshot, so the source memo answers every part the
        // transaction did not rewrite; the rewritten ones miss and are hashed.
        let (revision, digests) =
            super::model::package_fingerprint_with_memo(&working, Some(&source.part_digests))?;
        if revision != source.revision {
            refuse_signed_change(source.package.as_ref(), &working)?;
        }
        let patch = Patch::capture(
            source.package.as_ref(),
            &working,
            source.presentation_name.clone(),
            source.limits,
        )?;
        if patch.is_noop() {
            let package =
                crate::Package::from_opened_candidate(source.package.as_ref().clone(), &source)?;
            return Ok(Commit {
                package: Arc::new(package),
                snapshot: source,
                patch,
            });
        }
        // The staged package shares every slide payload allocation the
        // transaction did not rewrite, so the source snapshot's proved
        // notes-root classifications answer for those slides and only the
        // rewritten ones are rescanned; every other validation still runs.
        let mut snapshot =
            super::model::capture_with_revision_and_digests_and_mce_and_semantic_limits(
                &working,
                source.limits,
                source.physical_source_provenance,
                revision,
                digests,
                source.retained_mce.as_deref(),
                Some(source.slide_roots.as_ref()),
                source.owner_state.semantic_limits.clone(),
            )?;
        let changed = patch.resources().cloned().collect::<Vec<_>>();
        let base = source.replay_index.clone().map_or_else(
            || {
                crate::package::replay_index::ReplayIndex::from_package(source.package.as_ref())
                    .map(Arc::new)
            },
            Ok,
        )?;
        snapshot.replay_index = Some(Arc::new(base.update(
            source.package.as_ref(),
            &working,
            &changed,
        )?));
        snapshot.owner_state = Arc::clone(&source.owner_state);
        let package = crate::Package::from_opened_candidate(working.into_owned(), &snapshot)?;
        Ok(Commit {
            package: Arc::new(package),
            snapshot,
            patch,
        })
    }

    fn add_preset_shape(
        &mut self,
        slide: crate::slide::Selector<'_>,
        preset: &str,
        label: &str,
        bounds: (i64, i64, i64, i64),
        fill_color: Option<&str>,
    ) -> Result<u32> {
        let selected = self.resolve_slide(slide)?;
        let owner = self.working.get_part(&selected.part_name)?;
        let scene = crate::shape::Scene::read(owner.blob(), self.semantic_limits().shape())?;
        let shape_id = next_shape_id(&scene, label)?;
        let conformance = crate::media_parts::document_conformance_with_limits(
            owner.blob(),
            self.semantic_limits()
                .media_parts()
                .constrained_by_source(self.working.read_limits()),
        )?;
        let fragment =
            preset_shape_fragment(preset, label, bounds, fill_color, shape_id, conformance)?;
        self.append_checked_shape(&selected, shape_id, fragment.as_bytes())?;
        Ok(shape_id)
    }

    fn append_checked_shape(
        &mut self,
        slide: &Slide,
        shape_id: u32,
        fragment: &[u8],
    ) -> Result<()> {
        let shape_limits = self.semantic_limits().shape();
        let xml = super::xml::append_shape_with_limit(
            self.working.get_part(&slide.part_name)?.blob(),
            fragment,
            shape_limits.output_bytes(),
        )?;
        ensure_shape_id(&xml, shape_id, "shape", shape_limits)?;
        self.working.get_part_mut(&slide.part_name)?.set_blob(xml);
        Ok(())
    }

    fn resolve_slide_position(&mut self, selector: crate::slide::Selector<'_>) -> Result<usize> {
        self.ensure_ready()?;
        let position = match selector {
            crate::slide::Selector::Position(position) => position.get(),
            crate::slide::Selector::Name(name) => self
                .slide_name_index
                .as_ref()
                .ok_or_else(|| invalid("opened-presentation slide names are unavailable"))?
                .position(&name)?,
        };
        if position >= self.slides.len() {
            return Err(Error::SlideIndexOutOfBounds {
                index: position,
                len: self.slides.len(),
            });
        }
        Ok(position)
    }

    fn resolve_slide(&mut self, key: crate::slide::Selector<'_>) -> Result<Slide> {
        self.ensure_ready()?;
        match key {
            crate::slide::Selector::Position(position) => {
                let index = position.get();
                self.slides
                    .get(index)
                    .cloned()
                    .ok_or(Error::SlideIndexOutOfBounds {
                        index,
                        len: self.slides.len(),
                    })
            },
            crate::slide::Selector::Name(name) => self
                .slide_name_index
                .as_ref()
                .ok_or_else(|| invalid("opened-presentation slide names are unavailable"))?
                .resolve(&self.slides, &name),
        }
    }
}

/// Refuse a changed commit over a signed package (ADR 0006).
///
/// The change invalidates every package signature, and removing one is an
/// explicit operation ([`crate::Package::unsign`]), never a side effect of an
/// edit. The source's OPC policy requirement covers a graph whose signature
/// parts were removed by generic part removal, which does not count as an
/// explicit removal; `working` is checked as well, so a staged edit cannot
/// introduce signature infrastructure.
fn refuse_signed_change(source: &OpcPackage, working: &OpcPackage) -> Result<()> {
    if source.is_signed() || source.requires_signature_edit_policy() || working.is_signed() {
        return Err(Error::Opc(
            litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy,
        ));
    }
    Ok(())
}

#[cfg(test)]
fn compact_changed_slides(
    working: &mut OpcPackage,
    source: &OpcPackage,
    slides: &[Slide],
    scene_read: &HashMap<PackURI, Weak<Vec<u8>>>,
) -> Result<()> {
    compact_changed_slides_with_limits(
        working,
        source,
        slides,
        scene_read,
        crate::shape::Limits::DEFAULT,
    )
}

fn compact_changed_slides_with_limits(
    working: &mut OpcPackage,
    source: &OpcPackage,
    slides: &[Slide],
    scene_read: &HashMap<PackURI, Weak<Vec<u8>>>,
    shape_limits: crate::shape::Limits,
) -> Result<()> {
    for slide in slides {
        let Ok(before) = source.get_part(&slide.part_name) else {
            continue;
        };
        let after = working.get_part(&slide.part_name)?;
        let (before_blob, after_blob) = (before.blob(), after.blob());
        // An untouched slide shares its payload allocation with the source,
        // so identity decides equality without comparing every byte.
        if std::ptr::eq(before_blob, after_blob) || before_blob == after_blob {
            continue;
        }
        // A payload the transaction already read back as a complete scene
        // cannot fail that read now, so it is read again only if compaction
        // changes a byte and the two scenes must be compared.
        let staged_scene_read = scene_read
            .get(&slide.part_name)
            .and_then(Weak::upgrade)
            .is_some_and(|read| std::ptr::eq(read.as_slice(), after_blob));
        let expected = if staged_scene_read {
            debug_assert!(
                shape_semantics_with_limits(after.blob(), shape_limits).is_ok(),
                "a staged slide recorded as read back fails the scene read"
            );
            None
        } else {
            Some(compaction_scene_with_limits(after.blob(), shape_limits)?)
        };
        let compacted = super::xml::compact_changed_slide_xml(after.blob())?;
        // Compaction that reproduces the staged bytes exactly needs no read
        // back. For a different payload, an exact root witness can prove the
        // same semantic scene; all other payloads use the original comparison.
        if compacted.xml.as_slice() != after_blob {
            if compacted.root_unchanged {
                // An exact root-byte witness leaves only the already-ignored
                // outer ASCII whitespace changed. Keep a debug differential
                // oracle for both the staged-read and cold paths; direct
                // `shape_semantics` calls are deliberately outside the read
                // counter used to measure this production shortcut.
                #[cfg(debug_assertions)]
                if let Some(expected) = expected.as_ref() {
                    debug_assert!(
                        shape_semantics_with_limits(&compacted.xml, shape_limits)
                            .as_ref()
                            .is_ok_and(|actual| actual == expected),
                        "root-byte proof changed staged shape semantics",
                    );
                } else {
                    debug_assert!(
                        matches!(
                            (
                                shape_semantics_with_limits(&compacted.xml, shape_limits),
                                shape_semantics_with_limits(after.blob(), shape_limits),
                            ),
                            (Ok(compacted), Ok(staged)) if compacted == staged
                        ),
                        "root-byte proof changed read-back shape semantics",
                    );
                }
            } else {
                let expected = match expected {
                    Some(expected) => expected,
                    None => compaction_scene_with_limits(after.blob(), shape_limits)?,
                };
                if compaction_scene_with_limits(&compacted.xml, shape_limits)? != expected {
                    return Err(invalid(
                        "opened-presentation slide compaction changed shape semantics",
                    ));
                }
            }
        }
        working
            .get_part_mut(&slide.part_name)?
            .set_blob(compacted.xml);
    }
    Ok(())
}

#[cfg(test)]
thread_local! {
    /// Scene reads made by commit compaction on this thread, for tests that
    /// must show a read was skipped rather than only that the result held.
    static COMPACTION_SCENE_READS: std::cell::Cell<usize> = const { std::cell::Cell::new(0) };
}

type ShapeSemantics = Vec<(Option<u32>, Option<String>, Option<String>)>;

fn compaction_scene_with_limits(
    xml: &[u8],
    shape_limits: crate::shape::Limits,
) -> Result<ShapeSemantics> {
    #[cfg(test)]
    COMPACTION_SCENE_READS.with(|reads| reads.set(reads.get() + 1));
    shape_semantics_with_limits(xml, shape_limits)
}

#[cfg(test)]
fn shape_semantics(xml: &[u8]) -> Result<ShapeSemantics> {
    shape_semantics_with_limits(xml, crate::shape::Limits::DEFAULT)
}

fn shape_semantics_with_limits(
    xml: &[u8],
    shape_limits: crate::shape::Limits,
) -> Result<ShapeSemantics> {
    let scene = crate::shape::Scene::read(xml, shape_limits)?;
    Ok(scene
        .iter()
        .map(|shape| {
            (
                shape.id(),
                shape.name().map(str::to_owned),
                shape.common().text().map(str::to_owned),
            )
        })
        .collect())
}

fn remove_unreferenced_dependencies(package: &mut OpcPackage, roots: Vec<PackURI>) -> Result<()> {
    let mut queue = VecDeque::from(roots);
    let mut checked = HashSet::new();
    while let Some(name) = queue.pop_front() {
        if !checked.insert(name.clone()) {
            continue;
        }
        match package.get_part(&name) {
            Ok(_) => {},
            Err(litchi_opc::OpcError::PartNotFound(_)) => continue,
            Err(error) => return Err(error.into()),
        }
        if package
            .rels()
            .iter()
            .filter(|relationship| !relationship.is_external())
            .any(|relationship| {
                relationship
                    .target_partname()
                    .is_ok_and(|target| target == name)
            })
            || package.iter_parts().any(|part| {
                part.rels()
                    .iter()
                    .filter(|relationship| !relationship.is_external())
                    .any(|relationship| {
                        relationship
                            .target_partname()
                            .is_ok_and(|target| target == name)
                    })
            })
        {
            continue;
        }
        let dependencies: Vec<_> = package
            .get_part(&name)?
            .rels()
            .iter()
            .filter(|relationship| !relationship.is_external())
            .map(|relationship| relationship.target_partname().map_err(Error::Opc))
            .collect::<Result<_>>()?;
        if package.remove_part(&name) {
            queue.extend(dependencies);
        }
    }
    Ok(())
}

/// The first `{prefix}N{suffix}` part name not in the package.
fn free_part_name(package: &OpcPackage, prefix: &str, suffix: &str) -> Result<PackURI> {
    for index in 1_u32..=1_000_000 {
        let candidate = PackURI::new(format!("{prefix}{index}{suffix}")).map_err(Error::Uri)?;
        if !package.contains_part(&candidate) {
            return Ok(candidate);
        }
    }
    Err(invalid("no free part name is left"))
}

/// The chart part, and its embedded workbook part when it has one, of the
/// chart frame at `shape` on the slide part `slide`.
fn chart_targets(
    package: &OpcPackage,
    slide: &PackURI,
    shape: crate::shape::Key<'_>,
    shape_limits: crate::shape::Limits,
) -> Result<(PackURI, Option<PackURI>)> {
    use litchi_opc::constants::{content_type as ct, relationship_type as rt};

    let owner = package.get_part(slide)?;
    let scene = crate::shape::Scene::read(owner.blob(), shape_limits)?;
    let shape = scene
        .get(shape)
        .map_err(|error| invalid(format!("chart frame lookup: {error}")))?
        .ok_or_else(|| invalid("the chart frame was not found"))?;
    if !matches!(shape, crate::shape::Shape::Chart(_)) {
        return Err(invalid("the shape is not a chart frame"));
    }
    let relationship_id = crate::shape::raw::payload(shape)
        .map_err(|error| invalid(format!("chart frame relationship: {error}")))?
        .ok_or_else(|| invalid("the chart frame names no chart part"))?
        .relationship_id()
        .map_err(|error| invalid(format!("chart frame relationship: {error}")))?
        .into_owned();
    let relationship = owner
        .rels()
        .get(&relationship_id)
        .ok_or_else(|| invalid("the chart frame relationship is missing"))?;
    let extended = relationship.reltype() == litchi_drawingml::chart::chartex::RELATIONSHIP_TYPE;
    if relationship.is_external()
        || !(extended
            || crate::parts::is_relationship_type(relationship.reltype(), rt::CHART, "chart"))
    {
        return Err(invalid(
            "the chart frame relationship is not an internal chart",
        ));
    }
    let chart = relationship.target_partname()?;
    let part = package.get_part(&chart)?;
    let expected = if extended {
        litchi_drawingml::chart::chartex::CONTENT_TYPE
    } else {
        ct::DML_CHART
    };
    if part.content_type() != expected {
        return Err(invalid("the chart frame target is not a chart part"));
    }
    let mut workbook = None;
    for relationship in part.rels().iter() {
        if relationship.is_external()
            || !crate::parts::is_relationship_type(relationship.reltype(), rt::PACKAGE, "package")
        {
            continue;
        }
        if workbook.is_some() {
            return Err(invalid("the chart has more than one embedded workbook"));
        }
        workbook = Some(relationship.target_partname()?);
    }
    Ok((chart, workbook))
}

/// The graphic frame of an authored chart.
fn chart_frame(
    shape_id: u32,
    relationship_id: &str,
    name: &str,
    bounds: (i64, i64, i64, i64),
    extended: bool,
) -> String {
    use litchi_core::xml::escape_xml;
    let (x, y, width, height) = bounds;
    let (uri, element) = if extended {
        (litchi_drawingml::chart::chartex::GRAPHIC_DATA_URI, "cx")
    } else {
        (
            "http://schemas.openxmlformats.org/drawingml/2006/chart",
            "c",
        )
    };
    format!(
        concat!(
            r#"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="{id}" name="{name}"/>"#,
            r#"<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>"#,
            r#"<p:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{width}" cy="{height}"/></p:xfrm>"#,
            r#"<a:graphic><a:graphicData uri="{uri}">"#,
            r#"<{prefix}:chart xmlns:{prefix}="{uri}" r:id="{rel}"/>"#,
            r#"</a:graphicData></a:graphic></p:graphicFrame>"#
        ),
        uri = uri,
        prefix = element,
        id = shape_id,
        name = escape_xml(name),
        x = x,
        y = y,
        width = width,
        height = height,
        rel = escape_xml(relationship_id),
    )
}

pub(crate) fn next_shape_id(scene: &crate::shape::Scene<'_>, label: &str) -> Result<u32> {
    if scene.is_rewritten() {
        return Err(invalid(
            "opened-presentation shape authoring refuses an MCE-rewritten scene",
        ));
    }
    let maximum_shape_id = scene
        .iter()
        .filter_map(crate::shape::Shape::id)
        .max()
        .unwrap_or(1);
    let maximum_reference_id = max_shape_reference_id(scene.xml())?;
    maximum_shape_id
        .max(maximum_reference_id.unwrap_or(1))
        .checked_add(1)
        .ok_or_else(|| invalid(format!("opened-presentation {label} shape ID overflow")))
}

/// Find shape identities that are referenced outside the scene's ordinary
/// `cNvPr/@id` records.  Office uses these values for timing targets and
/// connector endpoints, so allocating only above visible shape IDs can still
/// activate a dangling reference or retarget an existing animation.
fn max_shape_reference_id(xml: &[u8]) -> Result<Option<u32>> {
    const P14: &[u8] = b"http://schemas.microsoft.com/office/powerpoint/2010/main";
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut maximum = None;
    loop {
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let (Event::Start(element) | Event::Empty(element)) = event else {
            if matches!(event, Event::Eof) {
                break;
            }
            continue;
        };
        let local = element.local_name();
        let namespace = match namespace {
            ResolveResult::Bound(Namespace(value)) => value,
            _ => continue,
        };
        let (accepted_namespace, attribute_name) = if (namespace
            == crate::namespace::PRESENTATIONML_NAMESPACE
            || namespace == crate::namespace::STRICT_PRESENTATIONML_NAMESPACE)
            && matches!(
                local.as_ref(),
                b"spTgt" | b"bldP" | b"bldDgm" | b"bldOleChart" | b"bldGraphic" | b"control"
            ) {
            (true, b"spid".as_slice())
        } else if namespace == P14 && local.as_ref() == b"bmkTgt" {
            (true, b"spid".as_slice())
        } else if namespace == P14
            && matches!(
                local.as_ref(),
                b"triggerEvt"
                    | b"playEvt"
                    | b"stopEvt"
                    | b"pauseEvt"
                    | b"resumeEvt"
                    | b"seekEvt"
                    | b"nullEvt"
            )
        {
            (true, b"objId".as_slice())
        } else if (namespace == litchi_ooxml_common::xml::DRAWINGML_NAMESPACE
            || namespace == litchi_ooxml_common::xml::STRICT_DRAWINGML_NAMESPACE)
            && matches!(local.as_ref(), b"stCxn" | b"endCxn")
        {
            (true, b"id".as_slice())
        } else {
            (false, &[][..])
        };
        if !accepted_namespace {
            continue;
        }
        for attribute in element.checked_attributes() {
            let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
            if attribute.key.as_ref() != attribute_name {
                continue;
            }
            let value = attribute
                .decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, reader.decoder())
                .map_err(|error| Error::Xml(error.to_string()))?;
            let value = value
                .trim()
                .parse::<u32>()
                .map_err(|_| invalid("opened-presentation shape reference identity is invalid"))?;
            maximum = Some(maximum.map_or(value, |current: u32| current.max(value)));
        }
    }
    Ok(maximum)
}

/// Bound the serialized text-box payload before allocating its escaped text,
/// fragment, or appended slide buffers. Six bytes per input byte covers every
/// XML character escape emitted by the native writer with a small margin for
/// UTF-8 and formatting; the fixed envelope covers namespaces, geometry, and
/// the surrounding DrawingML elements.
pub(crate) fn text_box_output_bound(
    input_bytes: usize,
    text_bytes: usize,
    max_output_bytes: usize,
) -> Result<usize> {
    const ESCAPED_TEXT_MULTIPLIER: usize = 6;
    const FRAGMENT_OVERHEAD: usize = 4096;
    let fragment = text_bytes
        .checked_mul(ESCAPED_TEXT_MULTIPLIER)
        .and_then(|bytes| bytes.checked_add(FRAGMENT_OVERHEAD))
        .ok_or(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::OutputBytes),
            observed: usize::MAX,
            maximum: max_output_bytes,
        })?;
    let output = input_bytes.checked_add(fragment).ok_or(Error::Limit {
        kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::OutputBytes),
        observed: usize::MAX,
        maximum: max_output_bytes,
    })?;
    if output > max_output_bytes {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::OutputBytes),
            observed: output,
            maximum: max_output_bytes,
        });
    }
    Ok(output)
}

/// Stage one plain text-box fragment against an already selected slide.
///
/// The source-backed primary editor uses this small native path instead of
/// constructing an eager package.  It shares the eager editor's identity
/// allocation, conformance-aware authoring, lexical shape-tree splice, and
/// scene readback checks.
pub(crate) fn stage_text_box(
    xml: &[u8],
    text: &str,
    bounds: (i64, i64, i64, i64),
    max_output_bytes: usize,
    shape_limits: crate::shape::Limits,
) -> Result<(Vec<u8>, u32)> {
    // Reject cheap caller errors before scene indexing or any fragment
    // allocation. The fragment builder repeats these checks at its own
    // boundary so future callers cannot bypass them.
    validate_text_box_input(text, bounds)?;
    // Admit the worst-case escaped fragment before parsing the scene or
    // allocating the exact fragment. The exact post-escape length is checked
    // again below before the lexical append.
    let _output_bound = text_box_output_bound(xml.len(), text.len(), max_output_bytes)?;
    let scene = crate::shape::Scene::read(xml, shape_limits)?;
    let shape_id = next_shape_id(&scene, "text box")?;
    let conformance = crate::media_parts::document_conformance(xml)?;
    let fragment = text_box_fragment(
        text,
        bounds,
        shape_id,
        conformance,
        super::Limits::default(),
    )?;
    let output_len = xml.len().checked_add(fragment.len()).ok_or(Error::Limit {
        kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::OutputBytes),
        observed: usize::MAX,
        maximum: max_output_bytes,
    })?;
    if output_len > max_output_bytes {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::OutputBytes),
            observed: output_len,
            maximum: max_output_bytes,
        });
    }
    let output = super::xml::append_shape_with_limit(xml, fragment.as_bytes(), max_output_bytes)?;
    if output.len() != output_len {
        return Err(Error::Invalid(
            "native text-box append changed the preflighted output size".into(),
        ));
    }
    ensure_shape_id(&output, shape_id, "text box", shape_limits)?;
    Ok((output, shape_id))
}

fn collect_transfer_shape_ids(
    shape: crate::shape::Shape<'_>,
    identities: &mut BTreeSet<u32>,
) -> Result<()> {
    match shape {
        crate::shape::Shape::Content(_) => Err(shape_transfer_refusal(
            crate::ShapeTransferRefusal::ContentPart,
            "p:contentPart has no common non-visual identity; use its typed content owner",
        )),
        crate::shape::Shape::Unknown(_) => Err(shape_transfer_refusal(
            crate::ShapeTransferRefusal::UnknownExtensionShape,
            "unknown source element is retained but cannot be dependency-closed",
        )),
        crate::shape::Shape::Frame(_) => Err(shape_transfer_refusal(
            crate::ShapeTransferRefusal::UnclassifiedGraphicFrame,
            "graphic frame has no recognized table, chart, diagram, or OLE payload",
        )),
        crate::shape::Shape::Group(group) => {
            insert_transfer_shape_id(shape, identities)?;
            for child in group.shapes() {
                collect_transfer_shape_ids(child, identities)?;
            }
            Ok(())
        },
        _ => insert_transfer_shape_id(shape, identities),
    }
}

fn collect_routable_shape_ids(
    shape: crate::shape::Shape<'_>,
    identities: &mut BTreeSet<u32>,
) -> Result<()> {
    if let Some(identity) = shape.id()
        && !identities.insert(identity)
    {
        return Err(invalid(
            "opened-presentation source shape root repeats an identity",
        ));
    }
    if let crate::shape::Shape::Group(group) = shape {
        for child in group.shapes() {
            collect_routable_shape_ids(child, identities)?;
        }
    }
    Ok(())
}

fn insert_transfer_shape_id(
    shape: crate::shape::Shape<'_>,
    identities: &mut BTreeSet<u32>,
) -> Result<()> {
    let identity = shape.id().ok_or_else(|| {
        shape_transfer_refusal(
            crate::ShapeTransferRefusal::MissingIdentity,
            "shape has no non-visual identity",
        )
    })?;
    if identities.insert(identity) {
        Ok(())
    } else {
        Err(invalid(
            "opened-presentation transferred shape repeats an identity",
        ))
    }
}

fn shape_transfer_refusal(kind: crate::ShapeTransferRefusal, _detail: impl Into<String>) -> Error {
    Error::shape_transfer_refused(kind)
}

fn plan_shape_identity_transfer(
    destination: &crate::shape::Scene<'_>,
    source_ids: &BTreeSet<u32>,
) -> Result<HashMap<u32, u32>> {
    if source_ids.is_empty() {
        return Err(invalid(
            "opened-presentation transferred shape has no non-visual identities",
        ));
    }
    let mut next = destination
        .iter()
        .filter_map(crate::shape::Shape::id)
        .max()
        .unwrap_or(1);
    let mut mapping = HashMap::new();
    mapping
        .try_reserve(source_ids.len())
        .map_err(|source| Error::Allocation {
            resource: "opened-presentation transferred shape identities",
            source,
        })?;
    for source_id in source_ids {
        next = next.checked_add(1).ok_or_else(|| {
            invalid("opened-presentation transferred shape identity space is exhausted")
        })?;
        mapping.insert(*source_id, next);
    }
    Ok(mapping)
}

fn resolve_snapshot_slide(source: &Snapshot, key: crate::slide::Selector<'_>) -> Result<Slide> {
    source.resolve_slide(key)
}

fn ensure_shape_id(
    xml: &[u8],
    shape_id: u32,
    label: &str,
    shape_limits: crate::shape::Limits,
) -> Result<()> {
    let scene = crate::shape::Scene::read(xml, shape_limits)?;
    if scene.iter().any(|shape| shape.id() == Some(shape_id)) {
        Ok(())
    } else {
        Err(invalid(format!(
            "opened-presentation {label} did not round-trip semantically"
        )))
    }
}

fn validate_bounds((_, _, width, height): (i64, i64, i64, i64), label: &str) -> Result<()> {
    if width <= 0 || height <= 0 {
        Err(invalid(format!(
            "opened-presentation {label} extents must be positive"
        )))
    } else {
        Ok(())
    }
}

/// Validate the caller-owned text-box inputs without reading the slide scene.
/// Source-backed admission uses this before reserving managed resources so an
/// invalid request fails without consuming parser or output budget.
pub(crate) fn validate_text_box_input(text: &str, bounds: (i64, i64, i64, i64)) -> Result<()> {
    validate_bounds(bounds, "text box")?;
    validate_shape_text(text, super::Limits::default(), "text box")
}

fn validate_shape_text(text: &str, limits: super::Limits, label: &str) -> Result<()> {
    if text.len() > limits.max_text_bytes() {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::TextBytes,
            observed: text.len(),
            maximum: limits.max_text_bytes(),
        });
    }
    if !text.chars().all(is_xml_char) {
        return Err(invalid(format!(
            "opened-presentation {label} contains an invalid XML character"
        )));
    }
    Ok(())
}

fn conformance_namespaces(
    conformance: crate::media_parts::Conformance,
) -> (&'static str, &'static str, &'static str) {
    match conformance {
        crate::media_parts::Conformance::Transitional => (
            "http://schemas.openxmlformats.org/presentationml/2006/main",
            "http://schemas.openxmlformats.org/drawingml/2006/main",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        ),
        crate::media_parts::Conformance::Strict => (
            "http://purl.oclc.org/ooxml/presentationml/main",
            "http://purl.oclc.org/ooxml/drawingml/main",
            "http://purl.oclc.org/ooxml/officeDocument/relationships",
        ),
    }
}

fn conformance_image_relationship(conformance: crate::media_parts::Conformance) -> &'static str {
    match conformance {
        crate::media_parts::Conformance::Transitional => {
            litchi_opc::constants::relationship_type::IMAGE
        },
        crate::media_parts::Conformance::Strict => {
            litchi_opc::constants::relationship_type::STRICT_IMAGE
        },
    }
}

pub(crate) fn text_box_fragment(
    text: &str,
    bounds: (i64, i64, i64, i64),
    shape_id: u32,
    conformance: crate::media_parts::Conformance,
    limits: super::Limits,
) -> Result<String> {
    validate_bounds(bounds, "text box")?;
    validate_shape_text(text, limits, "text box")?;
    let (pml, dml, _) = conformance_namespaces(conformance);
    let (x, y, width, height) = bounds;
    Ok(format!(
        "<p:sp xmlns:p=\"{pml}\" xmlns:a=\"{dml}\"><p:nvSpPr><p:cNvPr id=\"{shape_id}\" name=\"Text Box {shape_id}\"/><p:cNvSpPr txBox=\"1\"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x=\"{x}\" y=\"{y}\"/><a:ext cx=\"{width}\" cy=\"{height}\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody><a:bodyPr wrap=\"square\"/><a:lstStyle/><a:p><a:r><a:t xml:space=\"preserve\">{}</a:t></a:r><a:endParaRPr/></a:p></p:txBody></p:sp>",
        super::xml::escape_character_data(text, limits.max_patch_bytes())?
    ))
}

fn preset_shape_fragment(
    preset: &str,
    label: &str,
    bounds: (i64, i64, i64, i64),
    fill_color: Option<&str>,
    shape_id: u32,
    conformance: crate::media_parts::Conformance,
) -> Result<String> {
    validate_bounds(bounds, label)?;
    if !matches!(preset, "rect" | "ellipse") {
        return Err(invalid("opened-presentation preset shape is unsupported"));
    }
    if fill_color.is_some_and(|color| {
        color.len() != 6 || !color.bytes().all(|byte| byte.is_ascii_hexdigit())
    }) {
        return Err(invalid(
            "opened-presentation preset fill must be six hexadecimal digits",
        ));
    }
    let (pml, dml, _) = conformance_namespaces(conformance);
    let (x, y, width, height) = bounds;
    let fill = fill_color.map_or_else(
        || "<a:noFill/>".to_owned(),
        |color| format!("<a:solidFill><a:srgbClr val=\"{color}\"/></a:solidFill>"),
    );
    Ok(format!(
        "<p:sp xmlns:p=\"{pml}\" xmlns:a=\"{dml}\"><p:nvSpPr><p:cNvPr id=\"{shape_id}\" name=\"{label} {shape_id}\"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x=\"{x}\" y=\"{y}\"/><a:ext cx=\"{width}\" cy=\"{height}\"/></a:xfrm><a:prstGeom prst=\"{preset}\"><a:avLst/></a:prstGeom>{fill}</p:spPr></p:sp>"
    ))
}

fn picture_fragment(
    name: &str,
    bounds: (i64, i64, i64, i64),
    shape_id: u32,
    relationship_id: &str,
    conformance: crate::media_parts::Conformance,
) -> String {
    let (pml, dml, rel) = conformance_namespaces(conformance);
    let (x, y, width, height) = bounds;
    format!(
        "<p:pic xmlns:p=\"{pml}\" xmlns:a=\"{dml}\" xmlns:r=\"{rel}\"><p:nvPicPr><p:cNvPr id=\"{shape_id}\" name=\"{}\"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed=\"{relationship_id}\"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x=\"{x}\" y=\"{y}\"/><a:ext cx=\"{width}\" cy=\"{height}\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></p:spPr></p:pic>",
        quick_xml::escape::escape(name)
    )
}

fn shape_relationship_ids(
    fragment: &[u8],
    owner_relationships: &litchi_opc::Relationships,
) -> Result<BTreeSet<String>> {
    let mut reader = NsReader::from_reader(fragment);
    reader.config_mut().trim_text(false);
    let mut relationships = BTreeSet::new();
    loop {
        match reader
            .read_event()
            .map_err(|_error| Error::Xml("invalid transferred shape XML".into()))?
        {
            Event::Start(element) | Event::Empty(element) => {
                for attribute in element.checked_attributes() {
                    let attribute = attribute
                        .map_err(|_error| Error::Xml("invalid transferred shape XML".into()))?;
                    let key = attribute.key.as_ref();
                    if key == b"xmlns" || key.starts_with(b"xmlns:") {
                        continue;
                    }
                    let value = attribute
                        .decoded_and_normalized_value(
                            quick_xml::XmlVersion::Implicit1_0,
                            reader.decoder(),
                        )
                        .map_err(|_error| Error::Xml("invalid transferred shape XML".into()))?;
                    let (namespace, _) = reader.resolver().resolve_attribute(attribute.key);
                    let is_relationship_namespace = matches!(
                        &namespace,
                        ResolveResult::Bound(Namespace(value))
                            if *value == litchi_ooxml_common::relationships::TRANSITIONAL_NAMESPACE
                                || *value == litchi_ooxml_common::relationships::STRICT_NAMESPACE
                    );
                    let is_inherited_relationship_alias = matches!(
                        &namespace,
                        ResolveResult::Unknown(prefix)
                            if prefix.as_slice() == b"r"
                                || owner_relationships.get(value.as_ref()).is_some()
                    );
                    if is_relationship_namespace || is_inherited_relationship_alias {
                        relationships.insert(value.into_owned());
                    }
                }
            },
            Event::DocType(_) => {
                return Err(invalid(
                    "opened-presentation shape contains an unsupported document type",
                ));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(relationships)
}

fn table_fragment(
    cells: &[Vec<String>],
    (x, y, width, height): (i64, i64, i64, i64),
    shape_id: u32,
    conformance: crate::media_parts::Conformance,
    limits: super::Limits,
) -> Result<String> {
    let columns = cells
        .first()
        .map(Vec::len)
        .filter(|columns| *columns != 0)
        .ok_or_else(|| invalid("opened-presentation table must contain cells"))?;
    if cells.iter().any(|row| row.len() != columns) {
        return Err(invalid("opened-presentation table rows are ragged"));
    }
    if width <= 0 || height <= 0 {
        return Err(invalid(
            "opened-presentation table extents must be positive",
        ));
    }
    let text_bytes = cells
        .iter()
        .flatten()
        .try_fold(0usize, |total, cell| total.checked_add(cell.len()))
        .ok_or_else(|| invalid("opened-presentation table text size overflow"))?;
    if text_bytes > limits.max_text_bytes() {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::TextBytes,
            observed: text_bytes,
            maximum: limits.max_text_bytes(),
        });
    }
    if !cells
        .iter()
        .flatten()
        .flat_map(|cell| cell.chars())
        .all(is_xml_char)
    {
        return Err(invalid(
            "opened-presentation table contains an invalid XML character",
        ));
    }
    let columns_i64 = i64::try_from(columns)
        .map_err(|_err| invalid("opened-presentation table column count exceeds i64"))?;
    let rows_i64 = i64::try_from(cells.len())
        .map_err(|_err| invalid("opened-presentation table row count exceeds i64"))?;
    let column_width = width / columns_i64;
    let row_height = height / rows_i64;
    let (presentation_namespace, drawing_namespace) = match conformance {
        crate::media_parts::Conformance::Transitional => (
            "http://schemas.openxmlformats.org/presentationml/2006/main",
            "http://schemas.openxmlformats.org/drawingml/2006/main",
        ),
        crate::media_parts::Conformance::Strict => (
            "http://purl.oclc.org/ooxml/presentationml/main",
            "http://purl.oclc.org/ooxml/drawingml/main",
        ),
    };
    let mut xml = format!(
        "<p:graphicFrame xmlns:p=\"{presentation_namespace}\" xmlns:a=\"{drawing_namespace}\"><p:nvGraphicFramePr><p:cNvPr id=\"{shape_id}\" name=\"Table {shape_id}\"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x=\"{x}\" y=\"{y}\"/><a:ext cx=\"{width}\" cy=\"{height}\"/></p:xfrm><a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/table\"><a:tbl><a:tblPr firstRow=\"1\" bandRow=\"1\"/><a:tblGrid>"
    );
    for _ in 0..columns {
        xml.push_str(&format!("<a:gridCol w=\"{column_width}\"/>"));
    }
    xml.push_str("</a:tblGrid>");
    for row in cells {
        xml.push_str(&format!("<a:tr h=\"{row_height}\">"));
        for cell in row {
            xml.push_str("<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>");
            xml.push_str(&super::xml::escape_character_data(
                cell,
                limits.max_patch_bytes(),
            )?);
            xml.push_str("</a:t></a:r><a:endParaRPr/></a:p></a:txBody><a:tcPr/></a:tc>");
        }
        xml.push_str("</a:tr>");
    }
    xml.push_str("</a:tbl></a:graphicData></a:graphic></p:graphicFrame>");
    Ok(xml)
}

fn is_xml_char(value: char) -> bool {
    matches!(value, '\u{9}' | '\u{A}' | '\u{D}' | '\u{20}'..='\u{D7FF}' | '\u{E000}'..='\u{FFFD}' | '\u{10000}'..='\u{10FFFF}')
}

fn plan_transfer(
    source: &OpcPackage,
    destination: &OpcPackage,
    root: &PackURI,
    limit: usize,
) -> Result<HashMap<PackURI, PackURI>> {
    plan_transfer_roots(source, destination, vec![root.clone()], limit)
}

fn plan_transfer_roots(
    source: &OpcPackage,
    destination: &OpcPackage,
    roots: Vec<PackURI>,
    limit: usize,
) -> Result<HashMap<PackURI, PackURI>> {
    let mut queue = VecDeque::from(roots);
    let mut seen = HashSet::new();
    while let Some(name) = queue.pop_front() {
        if !seen.insert(name.clone()) {
            continue;
        }
        if seen.len() > limit {
            return Err(Error::Limit {
                kind: crate::opened::LimitKind::Parts,
                observed: seen.len(),
                maximum: limit,
            });
        }
        let part = source.get_part(&name)?;
        for relationship in part.rels().iter().filter(|item| !item.is_external()) {
            if relationship.target_query().is_some() || relationship.target_fragment().is_some() {
                return Err(invalid(
                    "opened-presentation transfer dependency has a query or fragment",
                ));
            }
            queue.push_back(relationship.target_partname()?);
        }
    }
    let mut names: Vec<_> = seen.into_iter().collect();
    names.sort_unstable_by(|left, right| left.as_str().cmp(right.as_str()));
    let mut reserved: HashSet<_> = destination
        .iter_parts()
        .map(|part| part.partname().clone())
        .collect();
    let mut mapping = HashMap::new();
    for source_name in names {
        let target = available_transfer_name(&source_name, &reserved)?;
        reserved.insert(target.clone());
        mapping.insert(source_name, target);
    }
    Ok(mapping)
}

fn available_transfer_name(source: &PackURI, reserved: &HashSet<PackURI>) -> Result<PackURI> {
    if !reserved.contains(source) {
        return Ok(source.clone());
    }
    let value = source.as_str();
    let (stem, extension) = value
        .rfind('.')
        .map_or((value, ""), |position| value.split_at(position));
    for index in 1..=u32::MAX {
        let candidate = PackURI::new(format!("{stem}-transfer{index}{extension}"))
            .map_err(|_error| invalid("opened-presentation transferred part name is invalid"))?;
        if !reserved.contains(&candidate) {
            return Ok(candidate);
        }
    }
    Err(invalid(
        "opened-presentation transfer part-name space is exhausted",
    ))
}

fn publish_transfer(
    source: &OpcPackage,
    destination: &mut OpcPackage,
    mapping: &HashMap<PackURI, PackURI>,
) -> Result<()> {
    let mut names: Vec<_> = mapping.keys().cloned().collect();
    names.sort_unstable_by(|left, right| left.as_str().cmp(right.as_str()));
    let mut staged = Vec::new();
    for source_name in names {
        let source_part = source.get_part(&source_name)?;
        let target_name = mapping
            .get(&source_name)
            .ok_or_else(|| invalid("opened-presentation transfer mapping disappeared"))?;
        let mut part = BlobPart::new_shared(
            target_name.clone(),
            source_part.content_type().to_owned(),
            source_part.blob_arc(),
        );
        for relationship in source_part.rels().iter() {
            let target = if relationship.is_external() {
                relationship.target_ref().to_owned()
            } else {
                let source_target = relationship.target_partname()?;
                let copied_target = mapping.get(&source_target).ok_or_else(|| {
                    invalid("opened-presentation transfer dependency closure is incomplete")
                })?;
                copied_target.relative_ref(target_name.base_uri())
            };
            part.rels_mut().try_add_relationship(
                relationship.reltype().to_owned(),
                target,
                relationship.r_id().to_owned(),
                if relationship.is_external() {
                    TargetMode::External
                } else {
                    TargetMode::Internal
                },
            )?;
        }
        staged.push(part);
    }
    for part in staged {
        destination.try_add_part(Box::new(part))?;
    }
    Ok(())
}

/// Validated immutable package and durable patch from a presentation edit.
///
/// ```
/// use litchi_pptx::Package;
/// # fn main() -> litchi_pptx::Result<()> {
/// let mut builder = litchi_pptx::builder::Builder::new()?;
/// builder.presentation()?.add_slide()?
///     .add_text_box("before", 0, 0, 914_400, 914_400);
/// let bytes = builder.build()?.to_bytes()?;
/// let source = Package::from_reader(std::io::Cursor::new(bytes.clone()))?;
/// let mut edit = source.edit()?;
/// edit.set_shape_text(0_usize, 0_usize, "after")?;
/// let commit = edit.commit()?;
/// assert!(commit.diagnostics().changed());
/// assert!(!commit.patch().is_noop());
/// let (changed, patch) = commit.into_parts();
/// let mut restored = changed.apply(&patch.inverse())?.into_package();
/// assert_eq!(restored.to_bytes()?, bytes);
/// # Ok(())
/// # }
/// ```
#[derive(Clone)]
pub struct Commit {
    package: Arc<crate::Package>,
    snapshot: Snapshot,
    patch: Patch,
}

impl std::fmt::Debug for Commit {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Commit")
            .field("diagnostics", &self.diagnostics())
            .field("resource_count", &self.patch.resource_count())
            .finish_non_exhaustive()
    }
}

impl Commit {
    pub(crate) fn new(package: crate::Package, snapshot: Snapshot, patch: Patch) -> Self {
        Self {
            package: Arc::new(package),
            snapshot,
            patch,
        }
    }

    /// The immutable package produced by this edit.
    #[must_use]
    pub fn package(&self) -> &crate::Package {
        &self.package
    }

    /// Consume the committed package without replaying its patch.
    ///
    /// A uniquely owned commit moves its package. When another commit clone
    /// shares it, the OPC graph is cloned while retained source buffers remain
    /// shared; native owner policy and tolerance evidence are preserved.
    #[must_use]
    pub fn into_package(self) -> crate::Package {
        self.into_parts().0
    }

    /// Content-free commit diagnostics.
    #[must_use]
    pub fn diagnostics(&self) -> Diagnostics {
        Diagnostics::new(!self.patch.is_noop())
    }

    #[cfg(test)]
    pub(crate) const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Aggregate charge for retained transformed slide projections.
    #[must_use]
    pub fn retained_mce_bytes(&self) -> usize {
        self.snapshot.retained_mce_bytes()
    }

    /// Release this commit's retained transformed slide projections.
    pub fn release_retained_mce(&mut self) {
        self.snapshot.release_retained_mce();
    }

    /// Durable exact-write-set patch for this commit.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Consume the result into its package and durable patch.
    #[must_use]
    pub fn into_parts(self) -> (crate::Package, Patch) {
        let package = Arc::try_unwrap(self.package).unwrap_or_else(|package| {
            package
                .derive_with_replay_index((*package.opc).clone(), Arc::clone(&package.replay_index))
        });
        (package, self.patch)
    }

    /// Consume the commit into the durable patch.
    #[must_use]
    pub fn into_patch(self) -> Patch {
        self.patch
    }
}

#[cfg(test)]
mod compaction_tests {
    #![allow(
        clippy::expect_used,
        clippy::unwrap_used,
        reason = "focused compaction tests use panic-on-fixture-failure assertions"
    )]

    use std::sync::Arc;

    use super::{COMPACTION_SCENE_READS, compact_changed_slides};
    use crate::{Package, Result};

    fn take_scene_reads() -> usize {
        COMPACTION_SCENE_READS.with(|reads| reads.replace(0))
    }

    fn three_slide_snapshot() -> Result<super::Snapshot> {
        let mut builder = crate::builder::Builder::new()?;
        for index in 0..3 {
            let slide = builder.presentation()?.add_slide()?;
            slide.add_text_box(&format!("compaction {index}"), 10, 20, 300, 400);
        }
        let package = builder.build()?;
        Package::from_reader(std::io::Cursor::new(package.to_bytes()?))?.opened_presentation()
    }

    fn three_slide_snapshot_with_outer_crlf() -> Result<super::Snapshot> {
        let mut builder = crate::builder::Builder::new()?;
        for index in 0..3 {
            let slide = builder.presentation()?.add_slide()?;
            slide.add_text_box(&format!("compaction {index}"), 10, 20, 300, 400);
        }
        let package = builder.build()?;
        let mut package = Package::from_reader(std::io::Cursor::new(package.to_bytes()?))?;
        let part_name = litchi_opc::PackURI::new("/ppt/slides/slide1.xml").unwrap();
        let original = package.opc.get_part(&part_name)?.blob().to_vec();
        let declaration_end = original
            .windows(2)
            .position(|bytes| bytes == b"?>")
            .expect("writer emits an XML declaration")
            + 2;
        let mut changed = Vec::with_capacity(original.len() + 2);
        changed.extend_from_slice(&original[..declaration_end]);
        changed.extend_from_slice(b"\r\n");
        changed.extend_from_slice(&original[declaration_end..]);
        Arc::make_mut(&mut package.opc)
            .get_part_mut(&part_name)?
            .set_blob(changed);
        Package::from_reader(std::io::Cursor::new(package.to_bytes()?))?.opened_presentation()
    }

    fn blob(package: &litchi_opc::OpcPackage, slide: &super::Slide) -> Arc<Vec<u8>> {
        package.get_part(&slide.part_name).unwrap().blob_arc()
    }

    #[test]
    fn untouched_slides_keep_their_allocation_and_a_reproduced_compaction_keeps_its_bytes()
    -> Result<()> {
        take_scene_reads();
        let snapshot = three_slide_snapshot()?;
        let mut working = snapshot.package.as_ref().clone();
        let edited = &snapshot.slides[1];
        let staged = String::from_utf8(blob(&working, edited).to_vec())
            .unwrap()
            .replace("compaction 1", "compaction one");
        working
            .get_part_mut(&edited.part_name)?
            .set_blob(staged.clone().into_bytes());
        compact_changed_slides(
            &mut working,
            snapshot.package.as_ref(),
            &snapshot.slides,
            &std::collections::HashMap::new(),
        )?;
        assert_eq!(
            take_scene_reads(),
            1,
            "cold byte-identical compaction keeps its initial validation read"
        );
        for untouched in [&snapshot.slides[0], &snapshot.slides[2]] {
            assert!(Arc::ptr_eq(
                &blob(&working, untouched),
                &blob(snapshot.package.as_ref(), untouched)
            ));
        }
        assert_eq!(blob(&working, edited).as_slice(), staged.as_bytes());
        Ok(())
    }

    #[test]
    fn a_compaction_that_changes_bytes_is_read_back_and_compared() -> Result<()> {
        take_scene_reads();
        let snapshot = three_slide_snapshot()?;
        let mut working = snapshot.package.as_ref().clone();
        let edited = &snapshot.slides[0];
        let staged = String::from_utf8(blob(&working, edited).to_vec())
            .unwrap()
            .replace("<p:spTree>", "<p:spTree>\n  ")
            .replace("</p:sp>", "</p:sp>\n  ");
        working
            .get_part_mut(&edited.part_name)?
            .set_blob(staged.clone().into_bytes());
        compact_changed_slides(
            &mut working,
            snapshot.package.as_ref(),
            &snapshot.slides,
            &std::collections::HashMap::new(),
        )?;
        assert_eq!(
            take_scene_reads(),
            2,
            "root-different compaction keeps both reads"
        );
        let compacted = blob(&working, edited);
        assert_ne!(compacted.as_slice(), staged.as_bytes());
        assert_eq!(
            super::shape_semantics(compacted.as_slice())?,
            super::shape_semantics(staged.as_bytes())?
        );
        Ok(())
    }

    #[test]
    fn a_staged_slide_the_scene_reader_refuses_is_still_refused() -> Result<()> {
        let snapshot = three_slide_snapshot()?;
        let mut working = snapshot.package.as_ref().clone();
        let edited = &snapshot.slides[2];
        let original = String::from_utf8(blob(&working, edited).to_vec()).unwrap();
        let shape = original.find("<p:sp>").expect("text box shape");
        let start = shape + original[shape..].find("<p:cNvPr ").expect("shape identity");
        let end = start + original[start..].find("/>").expect("identity ends") + 2;
        let duplicated = format!(
            "{}{}{}",
            &original[..end],
            &original[start..end],
            &original[end..]
        );
        working
            .get_part_mut(&edited.part_name)?
            .set_blob(duplicated.into_bytes());
        let error = compact_changed_slides(
            &mut working,
            snapshot.package.as_ref(),
            &snapshot.slides,
            &std::collections::HashMap::new(),
        )
        .expect_err("a scene the reader refuses must refuse compaction");
        assert!(
            error
                .to_string()
                .contains("more than one direct non-visual property record"),
            "{error}"
        );
        Ok(())
    }

    #[test]
    fn a_staged_payload_read_back_by_its_verb_is_not_read_again_at_commit() -> Result<()> {
        let snapshot = three_slide_snapshot()?;
        let mut edit = snapshot.edit();
        assert!(edit.set_shape_text(1, 0, "read once")?);
        take_scene_reads();
        let commit = edit.commit()?;
        assert_eq!(take_scene_reads(), 0, "the staged scene was already read");
        let text = crate::shape::Scene::read(
            commit
                .snapshot()
                .package
                .get_part(&snapshot.slides[1].part_name)?
                .blob(),
            crate::shape::Limits::DEFAULT,
        )?
        .at(0)
        .map_err(|error| crate::Error::Invalid(error.to_string()))?
        .common()
        .text()
        .map(str::to_owned);
        assert_eq!(text.as_deref(), Some("read once"));
        Ok(())
    }

    #[test]
    fn a_staged_read_with_outer_crlf_uses_the_root_witness() -> Result<()> {
        let snapshot = three_slide_snapshot_with_outer_crlf()?;
        let mut edit = snapshot.edit();
        assert!(edit.set_shape_text(0, 0, "read once")?);
        take_scene_reads();
        let commit = edit.commit()?;
        assert_eq!(
            take_scene_reads(),
            0,
            "the staged read and unchanged root prove compaction"
        );
        let blob = commit
            .snapshot()
            .package
            .get_part(&snapshot.slides[0].part_name)?
            .blob();
        assert!(!blob.windows(2).any(|bytes| bytes == b"\r\n"));
        Ok(())
    }

    #[test]
    fn a_cold_outer_crlf_compaction_keeps_the_initial_read_only() -> Result<()> {
        let snapshot = three_slide_snapshot_with_outer_crlf()?;
        let mut working = snapshot.package.as_ref().clone();
        let edited = &snapshot.slides[0];
        let staged = String::from_utf8(blob(&working, edited).to_vec())
            .unwrap()
            .replace("compaction 0", "compaction zero");
        working
            .get_part_mut(&edited.part_name)?
            .set_blob(staged.into_bytes());
        take_scene_reads();
        compact_changed_slides(
            &mut working,
            snapshot.package.as_ref(),
            &snapshot.slides,
            &std::collections::HashMap::new(),
        )?;
        assert_eq!(
            take_scene_reads(),
            1,
            "the staged validation remains; compacted root needs no second read"
        );
        Ok(())
    }

    #[test]
    fn admitted_real_fixture_root_witness_has_a_scene_differential_oracle() -> Result<()> {
        let bytes = std::fs::read("../../test-data/ooxml/pptx/shapes.pptx").unwrap();
        let package = Package::from_reader(std::io::Cursor::new(bytes))?;
        let snapshot = package.opened_presentation()?;
        let mut witnessed = 0usize;
        for slide in snapshot.slides.as_ref() {
            let xml = snapshot.package.get_part(&slide.part_name)?.blob();
            let compacted = super::super::xml::compact_changed_slide_xml(xml)?;
            if compacted.root_unchanged {
                let original = super::shape_semantics(xml)?;
                let compacted_semantics = super::shape_semantics(&compacted.xml)?;
                assert_eq!(compacted_semantics, original, "{}", slide.part_name);
                witnessed += 1;
            }
        }
        assert!(witnessed > 0, "the admitted fixture had no root witness");
        Ok(())
    }

    #[test]
    fn admitted_real_fixture_crlf_after_declaration_skips_two_readbacks() -> Result<()> {
        let bytes = std::fs::read("../../test-data/ooxml/pptx/shapes.pptx").unwrap();
        let package = Package::from_reader(std::io::Cursor::new(bytes))?;
        let snapshot = package.opened_presentation()?;
        let source = snapshot
            .package
            .get_part(&snapshot.slides[0].part_name)?
            .blob();
        assert!(source.windows(10).any(|bytes| bytes == b"?>\r\n<p:sld"));
        let mut edit = snapshot.edit();
        assert!(edit.set_shape_text(0, 0, "root-witness")?);
        take_scene_reads();
        let commit = edit.commit()?;
        assert_eq!(
            take_scene_reads(),
            0,
            "the real fixture's staged read and exact root skip both readbacks"
        );
        let output = commit
            .snapshot()
            .package
            .get_part(&snapshot.slides[0].part_name)?
            .blob();
        assert!(!output.windows(10).any(|bytes| bytes == b"?>\r\n<p:sld"));
        Ok(())
    }

    #[test]
    fn a_payload_replaced_after_its_scene_read_is_read_again_at_commit() -> Result<()> {
        let mut builder = crate::builder::Builder::new()?;
        let slide = builder.presentation()?.add_slide()?;
        slide.add_text_box("first", 10, 20, 300, 400);
        slide.add_text_box("second", 10, 500, 300, 400);
        let package = builder.build()?;
        let snapshot = Package::from_reader(std::io::Cursor::new(package.to_bytes()?))?
            .opened_presentation()?;
        let mut edit = snapshot.edit();
        assert!(edit.set_shape_text(0, 0, "recorded")?);
        // Removing a shape stages new bytes without recording a scene read.
        assert!(edit.remove_shape(0, 1)?);
        take_scene_reads();
        edit.commit()?;
        assert!(
            take_scene_reads() >= 1,
            "unrecorded bytes are read at commit"
        );
        Ok(())
    }

    #[test]
    fn a_record_for_another_allocation_does_not_vouch_for_the_staged_bytes() -> Result<()> {
        let snapshot = three_slide_snapshot()?;
        let mut working = snapshot.package.as_ref().clone();
        let edited = &snapshot.slides[0];
        let staged = String::from_utf8(blob(&working, edited).to_vec())
            .unwrap()
            .replace("compaction 0", "compaction zero");
        working
            .get_part_mut(&edited.part_name)?
            .set_blob(staged.into_bytes());
        // A record naming the source allocation, not the staged one.
        let stale = blob(snapshot.package.as_ref(), edited);
        let mut records = std::collections::HashMap::new();
        records.insert(edited.part_name.clone(), Arc::downgrade(&stale));
        take_scene_reads();
        compact_changed_slides(
            &mut working,
            snapshot.package.as_ref(),
            &snapshot.slides,
            &records,
        )?;
        assert_eq!(take_scene_reads(), 1, "the staged bytes are read once");
        Ok(())
    }
}
