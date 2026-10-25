//! Immutable deferred-source ownership, slide ordering and shape-property edits.
//!
//! Packages share the admitted source and retain changed root and slide XML.
//! Patches transport exact root/order context and selected slide preimages.
//! Structural revisions additionally retain one prepared OPC topology and a
//! stable-key recipe for lazy readback and reversible publication.

pub mod cross_slide_copy;
mod deferred;
mod durable;
mod insertion;
pub mod svg;
mod topology;
mod topology_admission;
mod topology_digest;
mod topology_wire;

use std::io::Write;
use std::path::Path;
use std::sync::{Arc, OnceLock};

use litchi_core::{Position, ReadAt, Reservation};
use litchi_opc::constants::relationship_type as rt;
use litchi_opc::{PackURI, PartData, TargetMode};
use sha2::{Digest, Sha256};

use super::{
    MAX_SOURCE_BACKED_SLIDE_BATCH, SourceBackedPresentation, SourceBackedPresentationEditor,
    SourceBackedSlidePatch, SourceBackedSlideSnapshot, SourcePayload,
};
use crate::Result;

use crate::presentation::order::{
    StableSlideKey, reorder_stable_bindings, reserve_managed_output, stable_slide_key,
    validate_source_order_publication, validate_source_order_root,
};

#[derive(Clone)]
enum RootPayload {
    Original(PartData),
    Edited(Arc<Vec<u8>>),
    ReservedEdit {
        bytes: Arc<Vec<u8>>,
        _reservation: Arc<Reservation>,
    },
}

impl RootPayload {
    fn as_bytes(&self) -> &[u8] {
        match self {
            Self::Original(data) => data.as_bytes(),
            Self::Edited(bytes) | Self::ReservedEdit { bytes, .. } => bytes.as_slice(),
        }
    }

    fn edited_bytes(&self) -> Option<Arc<Vec<u8>>> {
        match self {
            Self::Original(_) => None,
            Self::Edited(bytes) | Self::ReservedEdit { bytes, .. } => Some(Arc::clone(bytes)),
        }
    }
}

/// Immutable package backed by a retained positional source and slide overlays.
///
/// ```no_run
/// # fn main() -> litchi_pptx::Result<()> {
/// use litchi_pptx::source::{Package, Patch};
/// let source = Package::open("deck.pptx")?;
/// let mut edit = source.edit()?;
/// edit.set_shape_text(0, 0_usize, "Updated title")?;
/// let commit = edit.commit()?;
/// assert!(commit.diagnostics().changed());
/// let (updated, patch) = commit.into_parts();
/// updated.save("updated.pptx")?;
/// let transported = Patch::from_bytes(&patch.to_bytes()?)?;
/// let reopened = Package::open("deck.pptx")?;
/// let replayed = reopened.apply(&transported)?.into_package();
/// let restored = replayed.apply(&transported.inverse())?.into_package();
/// assert_eq!(restored.to_bytes()?, source.to_bytes()?);
/// # Ok(())
/// # }
/// ```
#[derive(Clone)]
pub struct Package {
    owner: Arc<SourceBackedPresentationEditor>,
    // Keep the admitted source owner alongside the effective topology.  A
    // retained physical slide can therefore continue through the native
    // related-part readers even when its logical position was changed by a
    // structural revision.
    source: Arc<super::SourceInner>,
    root: RootPayload,
    order: Arc<[StableSlideKey]>,
    physical: Arc<std::collections::HashMap<StableSlideKey, usize>>,
    overlays: Arc<[SourceBackedSlideSnapshot]>,
    topology: Option<Arc<topology::Revision>>,
    tolerance: Arc<crate::leniency::ToleranceReport>,
    /// The deferred output of a committed non-overlay family edit (record
    /// 0356a); reads and edits go through its materialized artifact.
    deferred: Option<Arc<deferred::Deferred>>,
}

/// Configuration of one source-backed primary edit.
///
/// The retained-patch ceiling is checked by the native editor's exact
/// encoded-length preflight before a commit is published.  It is intentionally
/// an edit option rather than a second package-wide source policy: callers can
/// pass the remaining budget of an outer transaction without reopening the
/// source owner.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[must_use]
pub struct EditOptions {
    max_patch_bytes: usize,
}

impl EditOptions {
    /// The native default retained-patch ceiling.
    pub const fn new() -> Self {
        Self {
            max_patch_bytes: crate::opened::Limits::DEFAULT.max_patch_bytes(),
        }
    }

    /// Set the exact encoded retained-patch ceiling for this edit.
    pub const fn max_patch_bytes(mut self, maximum: usize) -> Self {
        self.max_patch_bytes = maximum;
        self
    }

    /// The selected retained-patch ceiling.
    #[must_use]
    pub const fn max_patch_bytes_value(self) -> usize {
        self.max_patch_bytes
    }
}

impl Default for EditOptions {
    fn default() -> Self {
        Self::new()
    }
}

impl Package {
    /// Open an ordinary PPTX without materializing unselected parts.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    #[cfg(any(unix, windows))]
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        Self::from_presentation(&SourceBackedPresentation::open(path)?)
    }

    /// Open an ordinary PPTX under explicit ingress options (source limits,
    /// payload-cache policy, execution context).
    ///
    /// # Errors
    ///
    /// Returns an error if the input exceeds the policies, cannot be read or
    /// is malformed.
    #[cfg(any(unix, windows))]
    pub fn open_with_options(path: impl AsRef<Path>, options: crate::OpenOptions) -> Result<Self> {
        Self::from_presentation(&SourceBackedPresentation::open_with_options(path, options)?)
    }

    /// Admit an immutable positional source with finite default read limits.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn from_read_at(source: Arc<dyn ReadAt>) -> Result<Self> {
        Self::from_presentation(&SourceBackedPresentation::from_read_at(source)?)
    }

    /// Admit a positional source under explicit ingress options.
    ///
    /// # Errors
    ///
    /// Returns an error if the input exceeds the policies, cannot be read or
    /// is malformed.
    pub fn from_read_at_with_options(
        source: Arc<dyn ReadAt>,
        options: crate::OpenOptions,
    ) -> Result<Self> {
        Self::from_presentation(&SourceBackedPresentation::from_read_at_with_options(
            source, options,
        )?)
    }

    pub fn from_presentation(presentation: &SourceBackedPresentation) -> Result<Self> {
        let seed = presentation.editor_seed()?;
        let inner = &presentation.inner;
        let (root, slides) = seed
            .strict_catalog
            .unwrap_or_else(|| (inner._presentation.clone(), inner.slides.clone()));
        let owner = SourceBackedPresentationEditor {
            package: inner.package.clone(),
            _presentation: root,
            slides,
            package_relationships: seed.package_relationships,
            presentation_binding: seed.presentation_binding,
            limits: inner.limits,
            semantic_limits: inner.semantic_limits.clone(),
        };
        let order = source_order(&owner)?;
        let physical = source_physical_index(&owner)?;
        let root = RootPayload::Original(owner._presentation.source_data().clone());
        let package = Self {
            owner: Arc::new(owner),
            source: Arc::clone(inner),
            root,
            order: order.into(),
            physical: physical.into(),
            overlays: Arc::from([]),
            topology: None,
            tolerance: Arc::new(inner.tolerance.clone()),
            deferred: None,
        };
        package.check()?;
        Ok(package)
    }

    /// The number of the first slide, `p:presentation/@firstSlideNum`
    /// (1 when absent), from the current presentation part (record 0279).
    ///
    /// # Errors
    ///
    /// Returns an XML error for a malformed root or attribute.
    pub fn first_slide_number(&self) -> Result<u32> {
        let this = self.effective()?;
        if !std::ptr::eq(this, self) {
            return this.first_slide_number();
        }
        crate::presentation::first_slide_number(self.root.as_bytes())
    }

    /// Retained reference-compatible read repairs.
    pub fn tolerance_report(&self) -> &crate::leniency::ToleranceReport {
        &self.tolerance
    }

    /// Obtain an owned read-only semantic projection.
    pub fn presentation(&self) -> Result<Presentation> {
        let this = self.effective()?;
        if !std::ptr::eq(this, self) {
            return this.presentation();
        }
        self.check()?;
        Ok(Presentation {
            package: self.clone(),
        })
    }

    /// Begin an isolated primary edit; dropping it leaves this package intact.
    pub fn edit(&self) -> Result<Edit> {
        self.edit_with_options(EditOptions::default())
    }

    /// Begin an isolated primary edit under explicit retained-patch options.
    ///
    /// The source owner and its physical catalog are reused.  The selected
    /// ceiling is enforced by the same exact patch sizing used by `commit`,
    /// before the candidate is returned to the caller.
    pub fn edit_with_options(&self, options: EditOptions) -> Result<Edit> {
        let this = self.effective()?;
        if !std::ptr::eq(this, self) {
            return this.edit_with_options(options);
        }
        self.check()?;
        Ok(Edit {
            source: self.clone(),
            working: self.clone(),
            changes: Vec::new(),
            shape_removals: false,
            max_patch_bytes: options.max_patch_bytes_value(),
        })
    }

    /// Translate a logical presentation position to its retained physical
    /// source-catalog position.
    ///
    /// The mapping is checked against the immutable source catalog on every
    /// call.  An order edit therefore never makes a logical position point at
    /// an unrelated slide merely because two entries happen to share a
    /// numeric position. A Part restored after reopening has no physical
    /// source position, so this method returns an error for that Part. Use
    /// [`Presentation::slide`] for reads from the effective revision.
    pub fn source_position(&self, position: Position) -> Result<Position> {
        let this = self.effective()?;
        if !std::ptr::eq(this, self) {
            return this.source_position(position);
        }
        self.check()?;
        let logical = position.get();
        let key = self
            .order
            .get(logical)
            .ok_or(Error::SlideIndexOutOfBounds {
                index: logical,
                len: self.order.len(),
            })?;
        let physical = self
            .physical
            .get(key)
            .copied()
            .ok_or(Error::PatchConflict)?;
        self.check()?;
        Ok(Position::new(physical))
    }

    /// Whether two packages retain the same immutable physical source owner.
    ///
    /// This deliberately ignores staged root/order/slide revisions; callers
    /// that need revision equality should use [`Self::same_content`].
    #[must_use]
    pub fn same_source_owner(&self, other: &Self) -> bool {
        self.same_deferred(other)
            && self.owner.package.source_lineage() == other.owner.package.source_lineage()
            && self.owner._presentation.source_partname()
                == other.owner._presentation.source_partname()
            && self.owner._presentation.source_data().as_bytes()
                == other.owner._presentation.source_data().as_bytes()
    }

    /// Compare revisions of the same retained source without reading untouched parts.
    /// Independently opened sources are distinct even when their bytes match.
    pub fn same_content(&self, other: &Self) -> Result<bool> {
        if self.deferred.is_some() || other.deferred.is_some() {
            return Ok(self.same_deferred(other) && self.same_source_owner(other));
        }
        self.check()?;
        other.check()?;
        let equal = self.same_source_owner(other)
            && self.root.as_bytes() == other.root.as_bytes()
            && self.order == other.order
            && match (&self.topology, &other.topology) {
                (Some(left), Some(right)) => {
                    left.data.before.topology_digest == right.data.before.topology_digest
                        && left.data.after.topology_digest == right.data.after.topology_digest
                        && left.data.after.root == right.data.after.root
                        && left.data.after.order == right.data.after.order
                        && left.data.before.presentation_relationships
                            == right.data.before.presentation_relationships
                        && left.data.before.content_types == right.data.before.content_types
                        && left.data.after.presentation_relationships
                            == right.data.after.presentation_relationships
                        && left.data.after.content_types == right.data.after.content_types
                        && left.data.dependents == right.data.dependents
                        && left.data.changes.len() == right.data.changes.len()
                        && left.data.changes.iter().all(|change| {
                            right.data.changes.iter().any(|candidate| {
                                candidate.key == change.key
                                    && candidate.content_type == change.content_type
                                    && candidate.relationship_type == change.relationship_type
                                    && candidate.before == change.before
                                    && candidate.after == change.after
                            })
                        })
                },
                (None, None) => true,
                _ => false,
            }
            && self.overlays.len() == other.overlays.len()
            && self.overlays.iter().all(|left| {
                other.overlays.iter().any(|right| {
                    left.position == right.position && left.xml.as_bytes() == right.xml.as_bytes()
                })
            });
        self.check()?;
        other.check()?;
        Ok(equal)
    }

    /// Compare the bytes a replay would publish without requiring the two
    /// packages to share a source lineage.
    ///
    /// This is deliberately private to source-backed replay.  The public
    /// [`Self::same_content`] contract remains a same-owner semantic
    /// comparison, while a replay recipient may be an independent open of
    /// the exact artifact recorded by a patch.  Pristine packages use the
    /// retained source artifact fingerprint directly; an edited package is
    /// streamed through a bounded hash sink, so this proof never constructs
    /// a second full output buffer.
    pub(super) fn same_replay_content(&self, other: &Self) -> Result<bool> {
        Ok(self.replay_content_fingerprint()? == other.replay_content_fingerprint()?)
    }

    fn replay_content_fingerprint(&self) -> Result<(u64, [u8; 32])> {
        let effective = self.effective()?;
        if !std::ptr::eq(effective, self) {
            return effective.replay_content_fingerprint();
        }
        self.check()?;
        if self.is_pristine() {
            let artifact = self.owner.package.source_artifact();
            return Ok((
                artifact.len(),
                artifact.fingerprint().map_err(Error::from)?.into_sha256(),
            ));
        }
        let mut sink = ReplayHashWriter {
            digest: Sha256::new(),
            bytes: 0,
        };
        self.write_to(&mut sink)
            .map_err(map_replay_hash_write_error)?;
        Ok((sink.bytes, sink.digest.finalize().into()))
    }

    /// Rebase a replayed artifact onto this package's source policy.
    ///
    /// The artifact remains deferred and is streamed only when the recipient
    /// is written.  Both the recipient source and the recorded artifact source
    /// are checked before and after publication, while the deferred options
    /// and execution context come from the recipient package.
    pub(super) fn deferred_replay_from(&self, artifact: &Self) -> Result<Self> {
        let recipient = self.effective()?.clone();
        let check_recipient = recipient.clone();
        let recorded = artifact.clone();
        Ok(recipient.deferred(move |writer| {
            check_recipient.check()?;
            let result = recorded.write_to(writer);
            check_recipient.check()?;
            result
        }))
    }

    /// Replay an exact-source patch without changing this package.
    pub fn apply(&self, patch: &Patch) -> Result<Commit> {
        let this = self.effective()?;
        if !std::ptr::eq(this, self) {
            return this.apply(patch);
        }
        self.check()?;
        let (source, before_root, before_order, after_root, after_order, changes) =
            match &patch.data {
                PatchData::Local {
                    source,
                    before_root,
                    before_order,
                    after_root,
                    after_order,
                    changes,
                } => (
                    source,
                    before_root,
                    before_order,
                    after_root,
                    after_order,
                    changes,
                ),
                PatchData::Durable(data) => {
                    return durable::apply(self, data, patch.shape_removals)
                        .map(|package| commit_of(package, patch.clone()));
                },
                PatchData::Topology(data) => {
                    return data
                        .apply(self, patch.shape_removals)
                        .map(|package| commit_of(package, patch.clone()));
                },
            };
        source.check()?;
        if !self.same_source_owner(source)
            || self.root.as_bytes() != before_root.as_bytes()
            || self.order.as_ref() != before_order.as_ref()
        {
            return Err(Error::PatchConflict);
        }
        // Validate the complete write set before constructing a candidate.
        for change in changes.iter() {
            let key = source_key_for_position(source, change.before.position)?;
            let physical = self.physical_position(&key)?;
            change.apply(&self.snapshot(physical)?)?;
        }
        let mut candidate = self.with_order(after_root.clone(), after_order.clone())?;
        for change in changes.iter() {
            let key = source_key_for_position(source, change.before.position)?;
            let physical = candidate.physical_position(&key)?;
            let mut after = change.after.clone();
            after.position = physical;
            candidate = candidate.with_snapshot(after)?;
        }
        candidate.check()?;
        Ok(commit_of(candidate, patch.clone()))
    }

    /// Serialize with exact no-op and raw-copy preservation of untouched members.
    pub fn write_to(&self, mut writer: impl Write) -> Result<()> {
        if let Some(result) = self.write_deferred(&mut writer) {
            return result;
        }
        self.check()?;
        if let Some(revision) = &self.topology {
            revision.check()?;
            revision.prepared.publish_to_stream(writer)?;
            return Ok(());
        }
        let mut replacements = Vec::new();
        replacements
            .try_reserve_exact(
                self.overlays.len() + usize::from(self.root.edited_bytes().is_some()),
            )
            .map_err(allocation)?;
        if let Some(bytes) = self.root.edited_bytes() {
            validate_source_order_publication(
                &self.owner.package,
                self.owner._presentation.source_partname(),
            )?;
            replacements.push((self.owner._presentation.source_partname().clone(), bytes));
        }
        for slide in self.overlays.iter() {
            if let Some(bytes) = slide.xml.edited_bytes() {
                replacements.push((slide.part_uri.clone(), bytes));
            }
        }
        self.owner
            .package
            .clone()
            .write_part_overlays_shared_to_stream(writer, replacements)?;
        Ok(())
    }

    /// Serialize the retained package to memory.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        let mut bytes = Vec::new();
        self.write_to(&mut bytes)?;
        Ok(bytes)
    }

    /// Atomically replace a file through a sibling temporary artifact.
    pub fn save(&self, path: impl AsRef<Path>) -> Result<()> {
        self.save_with_durability(path, litchi_core::Durability::Full)
    }

    /// Atomically save with an explicit crash-durability policy.
    pub fn save_with_durability(
        &self,
        path: impl AsRef<Path>,
        durability: litchi_core::Durability,
    ) -> Result<()> {
        self.check()?;
        let context = self.owner.package.execution_context();
        litchi_opc::atomic::replace_guarded_with_durability(
            path.as_ref(),
            durability,
            context.as_ref(),
            |sink| self.write_to(sink),
            || self.check(),
        )
    }

    fn check(&self) -> Result<()> {
        self.owner.package.check_execution()?;
        self.owner.package.source_version()?;
        if let Some(revision) = &self.topology {
            revision.check()?;
            if self.order.as_ref() != revision.data.after.order.as_ref()
                || self.root.as_bytes() != revision.data.after.root.as_slice()
            {
                return Err(Error::PatchConflict);
            }
        } else if self.order.len() != self.owner.slides.len() {
            return Err(Error::PatchConflict);
        }
        Ok(())
    }

    fn physical_position(&self, key: &StableSlideKey) -> Result<usize> {
        self.physical.get(key).copied().ok_or(Error::PatchConflict)
    }

    fn source_slide_for_key(&self, key: &StableSlideKey) -> Option<super::SourceSlide> {
        // `physical` is the immutable source index.  Looking up the catalog
        // entry by stable identity keeps this independent of the effective
        // order and avoids treating a restored Part as source-backed.
        let position = *self.physical.get(key)?;
        let data = self.owner.slides.get(position)?.clone();
        Some(super::SourceSlide {
            owner: Arc::clone(&self.source),
            data,
        })
    }

    fn read_snapshot(&self, position: usize) -> Result<SourceBackedSlideSnapshot> {
        self.check()?;
        if let Some(revision) = &self.topology {
            let key = self
                .order
                .get(position)
                .ok_or(Error::SlideIndexOutOfBounds {
                    index: position,
                    len: self.order.len(),
                })?;
            let snapshot = revision.snapshot_for_key(self, key, position)?;
            self.check()?;
            return Ok(snapshot);
        }
        let key = self
            .order
            .get(position)
            .ok_or(Error::SlideIndexOutOfBounds {
                index: position,
                len: self.order.len(),
            })?;
        let physical = self.physical_position(key)?;
        let snapshot = match self
            .overlays
            .iter()
            .find(|slide| slide.position == physical)
        {
            Some(slide) => slide.clone(),
            None => self
                .owner
                .publication_raw_slide_snapshot_from_retained_catalog(physical)?,
        };
        crate::source_limits::validate_xml_limits(snapshot.xml.as_bytes(), self.owner.limits)?;
        self.check()?;
        Ok(snapshot)
    }

    fn snapshot(&self, position: usize) -> Result<SourceBackedSlideSnapshot> {
        self.check()?;
        if let Some(revision) = &self.topology {
            let key = self.owner.slides.get(position).map(slide_key).ok_or(
                Error::SlideIndexOutOfBounds {
                    index: position,
                    len: self.owner.slides.len(),
                },
            )?;
            if self.order.iter().any(|candidate| candidate == &key) {
                let logical = self
                    .order
                    .iter()
                    .position(|candidate| candidate == &key)
                    .ok_or(Error::PatchConflict)?;
                return revision.snapshot_for_key(self, &key, logical);
            }
        }
        let snapshot = match self
            .overlays
            .iter()
            .find(|slide| slide.position == position)
        {
            Some(slide) => slide.clone(),
            None => self.owner.slide_snapshot(position)?,
        };
        crate::source_limits::validate_xml_limits(snapshot.xml.as_bytes(), self.owner.limits)?;
        self.check()?;
        Ok(snapshot)
    }

    fn snapshot_logical(&self, position: usize) -> Result<SourceBackedSlideSnapshot> {
        self.check()?;
        let key = self
            .order
            .get(position)
            .ok_or(Error::SlideIndexOutOfBounds {
                index: position,
                len: self.order.len(),
            })?;
        if let Some(revision) = &self.topology {
            return revision.snapshot_for_key(self, key, position);
        }
        let physical = self.physical_position(key)?;
        self.snapshot(physical)
    }

    fn with_snapshot(&self, snapshot: SourceBackedSlideSnapshot) -> Result<Self> {
        self.with_snapshot_with_max_patch_bytes(
            snapshot,
            crate::opened::Limits::default().max_patch_bytes(),
        )
    }

    fn with_snapshot_with_max_patch_bytes(
        &self,
        mut snapshot: SourceBackedSlideSnapshot,
        max_patch_bytes: usize,
    ) -> Result<Self> {
        self.check()?;
        if self.topology.is_some() {
            return topology::replace_snapshot(self, snapshot, max_patch_bytes);
        }
        if snapshot.xml.as_bytes().len() > snapshot.max_output_bytes {
            return Err(limit(
                crate::error::CodecLimitKind::OutputBytes,
                snapshot.xml.as_bytes().len(),
                snapshot.max_output_bytes,
            ));
        }
        crate::source_limits::validate_xml_limits(snapshot.xml.as_bytes(), self.owner.limits)?;
        let original = self.owner.slide_snapshot(snapshot.position)?;
        let changed = !snapshot.same_source(&original);
        if changed && self.owner.package.has_signature_infrastructure() {
            return Err(litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy.into());
        }
        if changed {
            // A replay target may retain an original payload of another
            // logical revision. Every physical-source difference is an overlay.
            if let SourcePayload::Original(_) = snapshot.xml {
                return Err(Error::PatchConflict);
            }
        } else {
            snapshot = original;
        }
        let count = self
            .overlays
            .iter()
            .filter(|slide| slide.position != snapshot.position)
            .count()
            + usize::from(changed);
        if count > MAX_SOURCE_BACKED_SLIDE_BATCH {
            return Err(limit(
                crate::error::CodecLimitKind::Slides,
                count,
                MAX_SOURCE_BACKED_SLIDE_BATCH,
            ));
        }
        let mut overlays = Vec::new();
        overlays.try_reserve_exact(count).map_err(allocation)?;
        overlays.extend(
            self.overlays
                .iter()
                .filter(|slide| slide.position != snapshot.position)
                .cloned(),
        );
        if changed {
            overlays.push(snapshot);
        }
        overlays.sort_unstable_by_key(|slide| slide.position);
        let root_bytes = self.root.edited_bytes();
        check_payload_budget(root_bytes.as_deref().map(Vec::as_slice), overlays.iter())?;
        self.check()?;
        Ok(Self {
            overlays: overlays.into(),
            ..self.clone()
        })
    }

    fn with_order(&self, root: RootPayload, order: Arc<[StableSlideKey]>) -> Result<Self> {
        self.with_order_with_max_patch_bytes(
            root,
            order,
            crate::opened::Limits::default().max_patch_bytes(),
        )
    }

    fn with_order_with_max_patch_bytes(
        &self,
        root: RootPayload,
        order: Arc<[StableSlideKey]>,
        max_patch_bytes: usize,
    ) -> Result<Self> {
        self.check()?;
        if self.topology.is_some() {
            let bytes = root.edited_bytes().ok_or(Error::PatchConflict)?;
            return topology::reorder(self, bytes, order, max_patch_bytes);
        }
        if order.len() != self.owner.slides.len() {
            return Err(Error::PatchConflict);
        }
        let changed = root.as_bytes() != self.root.as_bytes() || order != self.order;
        if !changed {
            return Ok(self.clone());
        }
        if self.owner.package.has_signature_infrastructure() {
            return Err(litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy.into());
        }
        validate_source_order_root(
            &self.owner.package,
            self.owner._presentation.source_partname(),
            root.as_bytes(),
            &order,
            self.owner.package.execution_context().as_ref(),
        )?;
        if root.as_bytes().len() > self.owner.limits.max_part_bytes() as usize {
            return Err(limit(
                crate::error::CodecLimitKind::OutputBytes,
                root.as_bytes().len(),
                self.owner.limits.max_part_bytes() as usize,
            ));
        }
        let original_order = source_order(&self.owner)?;
        if order.as_ref() == original_order.as_ref()
            && root.as_bytes() == self.owner._presentation.source_data().as_bytes()
        {
            return Ok(Self {
                root: RootPayload::Original(self.owner._presentation.source_data().clone()),
                order,
                ..self.clone()
            });
        }
        let root_bytes = root.edited_bytes();
        check_payload_budget(
            root_bytes.as_deref().map(Vec::as_slice),
            self.overlays.iter(),
        )?;
        Ok(Self {
            root,
            order,
            ..self.clone()
        })
    }

    /// Stage a new logical order. `moved` names the position of a slide
    /// that one move placed; it leaves its p14 section and joins its new
    /// neighbour's, as `PowerPoint` moves a slide. Without it every slide
    /// keeps its section and each section is re-sorted (record 0303).
    fn stage_order(
        &self,
        ordered: Vec<StableSlideKey>,
        moved: Option<usize>,
        max_patch_bytes: usize,
        source: &Package,
        changes: &[SourceBackedSlidePatch],
        shape_removals: bool,
    ) -> Result<Self> {
        self.check()?;
        let ordered: Arc<[StableSlideKey]> = ordered.into();
        if ordered == self.order {
            return Ok(self.clone());
        }
        if crate::opened::sections::present(self.root.as_bytes()) {
            return self.stage_sectioned_order(
                ordered,
                moved,
                max_patch_bytes,
                source,
                changes,
                shape_removals,
            );
        }
        if self.topology.is_some() {
            let xml = reorder_stable_bindings(
                self.root.as_bytes(),
                &self.order,
                &ordered,
                self.owner.package.execution_context().as_ref(),
            )?;
            let topology_budget = topology::topology_patch_budget(max_patch_bytes, shape_removals)?;
            return topology::reorder(self, Arc::new(xml), ordered, topology_budget);
        }
        let source_order = source_order(&self.owner)?;
        let source_xml = self.owner._presentation.source_data().as_bytes();
        let mut root = if ordered.as_ref() == source_order.as_ref() {
            RootPayload::Original(self.owner._presentation.source_data().clone())
        } else {
            validate_source_order_root(
                &self.owner.package,
                self.owner._presentation.source_partname(),
                source_xml,
                &source_order,
                self.owner.package.execution_context().as_ref(),
            )?;
            let xml = reorder_stable_bindings(
                source_xml,
                &source_order,
                &ordered,
                self.owner.package.execution_context().as_ref(),
            )?;
            RootPayload::Edited(Arc::new(xml))
        };
        self.check_local_order_patch_budget(
            source,
            &ordered,
            root.as_bytes().len(),
            changes,
            shape_removals,
            max_patch_bytes,
        )?;
        root = match root {
            RootPayload::Edited(bytes) => match reserve_managed_output(
                self.owner.package.execution_context().as_ref(),
                bytes.len(),
            )? {
                Some(reservation) => RootPayload::ReservedEdit {
                    bytes,
                    _reservation: reservation,
                },
                None => RootPayload::Edited(bytes),
            },
            root => root,
        };
        self.with_order_with_max_patch_bytes(root, ordered, max_patch_bytes)
    }

    /// Stage an order on a root with a p14 section list. Section membership
    /// depends on the sequence of moves, so the candidate is derived from
    /// the current root rather than rebuilt from the source.
    fn stage_sectioned_order(
        &self,
        ordered: Arc<[StableSlideKey]>,
        moved: Option<usize>,
        max_patch_bytes: usize,
        source: &Package,
        changes: &[SourceBackedSlidePatch],
        shape_removals: bool,
    ) -> Result<Self> {
        let context = self.owner.package.execution_context();
        validate_source_order_root(
            &self.owner.package,
            self.owner._presentation.source_partname(),
            self.root.as_bytes(),
            &self.order,
            context.as_ref(),
        )?;
        let xml = reorder_stable_bindings(
            self.root.as_bytes(),
            &self.order,
            &ordered,
            context.as_ref(),
        )?;
        let ids: Vec<u32> = ordered.iter().map(|key| key.slide_id).collect();
        let sectioned = match moved {
            Some(position) => crate::opened::sections::place(&xml, &ids, position)?,
            None => crate::opened::sections::sort(&xml, &ids)?,
        };
        let xml = Arc::new(sectioned.unwrap_or(xml));
        if self.topology.is_some() {
            let topology_budget = topology::topology_patch_budget(max_patch_bytes, shape_removals)?;
            return topology::reorder(self, xml, ordered, topology_budget);
        }
        // The exact local durable envelope is admitted before its managed
        // root reservation.  A rejected order therefore does not retain the
        // newly generated root merely to discover that the patch ceiling was
        // exceeded at commit.
        self.check_local_order_patch_budget(
            source,
            &ordered,
            xml.len(),
            changes,
            shape_removals,
            max_patch_bytes,
        )?;
        let root = match reserve_managed_output(context.as_ref(), xml.len())? {
            Some(reservation) => RootPayload::ReservedEdit {
                bytes: xml,
                _reservation: reservation,
            },
            None => RootPayload::Edited(xml),
        };
        self.with_order_with_max_patch_bytes(root, ordered, max_patch_bytes)
    }

    fn check_local_order_patch_budget(
        &self,
        source: &Package,
        order: &[StableSlideKey],
        root_len: usize,
        changes: &[SourceBackedSlidePatch],
        shape_removals: bool,
        max_patch_bytes: usize,
    ) -> Result<()> {
        let mut lengths = Vec::new();
        lengths
            .try_reserve_exact(changes.len())
            .map_err(allocation)?;
        for change in changes {
            let key = source_key_for_position(source, change.before.position)?;
            lengths.push((
                key,
                change.before.xml.as_bytes().len(),
                change.after.xml.as_bytes().len(),
            ));
        }
        let encoded = durable::local_wire_len(
            shape_removals,
            &source.order,
            order,
            source.root.as_bytes().len(),
            root_len,
            &lengths,
        )?;
        if encoded > max_patch_bytes {
            return Err(limit(
                crate::error::CodecLimitKind::PatchBytes,
                encoded,
                max_patch_bytes,
            ));
        }
        Ok(())
    }
}

struct ReplayHashWriter {
    digest: Sha256,
    bytes: u64,
}

fn map_replay_hash_write_error(error: Error) -> Error {
    match error {
        Error::Io(error) if is_replay_hash_error(&error) => Error::CapacityOverflow {
            resource: "replay content fingerprint",
        },
        Error::Opc(litchi_opc::OpcError::IoError(error)) if is_replay_hash_error(&error) => {
            Error::CapacityOverflow {
                resource: "replay content fingerprint",
            }
        },
        error => error,
    }
}

fn is_replay_hash_error(error: &std::io::Error) -> bool {
    error
        .get_ref()
        .and_then(|source| source.downcast_ref::<ReplayHashError>())
        .is_some()
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum ReplayHashError {
    InputLengthExceedsU64,
    LengthOverflow,
}

impl std::fmt::Display for ReplayHashError {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(match self {
            Self::InputLengthExceedsU64 => "replay input length exceeds u64",
            Self::LengthOverflow => "replay content length overflows u64",
        })
    }
}

impl std::error::Error for ReplayHashError {}

impl Write for ReplayHashWriter {
    fn write(&mut self, bytes: &[u8]) -> std::io::Result<usize> {
        let amount = u64::try_from(bytes.len())
            .map_err(|_| std::io::Error::other(ReplayHashError::InputLengthExceedsU64))?;
        let total = self
            .bytes
            .checked_add(amount)
            .ok_or_else(|| std::io::Error::other(ReplayHashError::LengthOverflow))?;
        self.digest.update(bytes);
        self.bytes = total;
        Ok(bytes.len())
    }

    fn flush(&mut self) -> std::io::Result<()> {
        Ok(())
    }
}

#[cfg(test)]
mod replay_hash_tests {
    use super::{ReplayHashError, ReplayHashWriter, map_replay_hash_write_error};
    use crate::Error;
    use sha2::{Digest, Sha256};
    use std::io::Write;

    #[test]
    fn replay_hash_writer_rejects_length_overflow() {
        let mut writer = ReplayHashWriter {
            digest: Sha256::new(),
            bytes: u64::MAX,
        };
        let error = writer
            .write_all(b"x")
            .expect_err("a length past u64::MAX must be refused");
        assert_eq!(error.kind(), std::io::ErrorKind::Other);
        assert!(matches!(
            error
                .get_ref()
                .and_then(|source| source.downcast_ref::<ReplayHashError>()),
            Some(ReplayHashError::LengthOverflow)
        ));
        assert_eq!(writer.bytes, u64::MAX);
        assert_eq!(writer.digest.clone().finalize(), Sha256::digest([]));
    }

    #[test]
    fn replay_hash_boundary_maps_wrapped_capacity_error() {
        let source = std::io::Error::other(ReplayHashError::LengthOverflow);
        let error = map_replay_hash_write_error(Error::Opc(litchi_opc::OpcError::IoError(source)));
        assert!(matches!(
            error,
            Error::CapacityOverflow {
                resource: "replay content fingerprint"
            }
        ));
    }

    #[test]
    fn replay_hash_boundary_maps_direct_capacity_error() {
        let source = std::io::Error::other(ReplayHashError::InputLengthExceedsU64);
        let error = map_replay_hash_write_error(Error::Io(source));
        assert!(matches!(
            error,
            Error::CapacityOverflow {
                resource: "replay content fingerprint"
            }
        ));
    }

    #[test]
    fn replay_hash_boundary_preserves_destination_io() {
        let source = std::io::Error::new(std::io::ErrorKind::BrokenPipe, "sink");
        let error = map_replay_hash_write_error(Error::Opc(litchi_opc::OpcError::IoError(source)));
        assert!(matches!(
            error,
            Error::Opc(litchi_opc::OpcError::IoError(error))
                if error.kind() == std::io::ErrorKind::BrokenPipe
        ));
    }
}

/// Owned read-only presentation over one immutable source package revision.
#[derive(Clone)]
pub struct Presentation {
    package: Package,
}

const NOTES_SLIDE_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide";

/// Inputs accepted by the compatibility read method.  New callers should
/// pass [`Position`]; the integer implementations keep existing native read
/// call sites source-compatible while still checking signed values.
pub trait PresentationPosition {
    fn into_presentation_position(self) -> Result<Position>;
}

impl PresentationPosition for Position {
    fn into_presentation_position(self) -> Result<Position> {
        Ok(self)
    }
}

impl PresentationPosition for usize {
    fn into_presentation_position(self) -> Result<Position> {
        Ok(Position::new(self))
    }
}

impl PresentationPosition for i32 {
    fn into_presentation_position(self) -> Result<Position> {
        usize::try_from(self)
            .map(Position::new)
            .map_err(|_| Error::Invalid("presentation position is negative".into()))
    }
}

impl Presentation {
    /// Number of slides in authored order.
    pub fn slide_count(&self) -> Result<usize> {
        self.package.check()?;
        Ok(self.package.order.len())
    }

    /// Return the stable `p:sldId@id` at a logical position without reading
    /// the selected slide XML or any related Part.
    pub fn slide_id(&self, position: impl PresentationPosition) -> Result<u32> {
        let logical = position.into_presentation_position()?.get();
        self.package.check()?;
        let slide_id = self
            .package
            .order
            .get(logical)
            .ok_or(Error::SlideIndexOutOfBounds {
                index: logical,
                len: self.package.order.len(),
            })?
            .slide_id;
        self.package.check()?;
        Ok(slide_id)
    }

    /// Acquire one slide from this exact revision.
    pub fn slide(&self, position: impl PresentationPosition) -> Result<Slide> {
        let logical = position.into_presentation_position()?.get();
        self.package.check()?;
        let key = self
            .package
            .order
            .get(logical)
            .ok_or(Error::SlideIndexOutOfBounds {
                index: logical,
                len: self.package.order.len(),
            })?
            .clone();
        Ok(Slide {
            snapshot: self.package.read_snapshot(logical)?,
            package: self.package.clone(),
            key,
        })
    }
}

/// Owned read-only slide whose reads retain source freshness and execution fences.
#[derive(Clone)]
pub struct Slide {
    snapshot: SourceBackedSlideSnapshot,
    package: Package,
    key: StableSlideKey,
}
impl Slide {
    /// Flatten text from this revision using the native slide reader.
    pub fn text(&self) -> Result<String> {
        self.text_with_tolerance(&mut crate::leniency::ToleranceReport::default())
    }

    /// Read text and append repairs only after all source fences succeed.
    pub fn text_with_tolerance(
        &self,
        tolerance: &mut crate::leniency::ToleranceReport,
    ) -> Result<String> {
        let mut repairs = crate::leniency::ToleranceReport::default();
        let text = self.with_part(|part| {
            crate::parts::SlidePart::from_part(part)?.text_with_tolerance(&mut repairs)
        })?;
        tolerance.append(&repairs);
        Ok(text)
    }

    /// Read the native flattened text and producer-visible name together.
    pub fn text_and_name(&self) -> Result<(String, String)> {
        self.with_part(|part| crate::parts::SlidePart::from_part(part)?.text_and_name())
    }

    /// Return the authored `p:cSld@name`, without a physical part-name
    /// fallback.
    pub fn authored_name(&self) -> Result<Option<String>> {
        self.with_part(|part| crate::parts::SlidePart::from_part(part)?.authored_name())
    }

    /// Whether the effective slide is authored as hidden.
    pub fn is_hidden(&self) -> Result<bool> {
        self.with_part(|part| crate::parts::SlidePart::from_part(part)?.is_hidden())
    }

    /// The effective slide's direct transition, when it has one (record
    /// 0356a; the read the source-backed slide snapshot offered). A
    /// transition inside a supported markup-compatibility owner is read
    /// through its typed choice and fallback.
    ///
    /// # Errors
    ///
    /// Returns an error for malformed or unsupported transition markup.
    pub fn transition(&self) -> Result<Option<crate::transition::Transition>> {
        self.package.check()?;
        let transition = self.snapshot.transition();
        self.package.check()?;
        transition
    }

    /// Acquire the effective slide's selected layout snapshot.
    ///
    /// This uses the effective relationship binding retained by the snapshot.
    /// A restored slide therefore reads its unchanged source layout Part,
    /// including after durable inverse replay on a source ZIP where the slide
    /// Part is physically absent. The native layout reader preserves related
    /// metadata, XML policy, scene caching, and managed reservations.
    pub fn layout(&self) -> Result<Option<super::related::Layout>> {
        self.package.check()?;
        if let Some(slide) = self.physical_source_slide() {
            let result = slide.layout();
            self.package.check()?;
            return result;
        }
        let result = self.read_layout();
        self.package.check()?;
        result
    }

    /// Acquire the effective slide's optional notes snapshot.
    ///
    /// Dependency-free structural primary revisions reject notes edges during
    /// admission, so a restored slide has no notes snapshot. The return shape
    /// matches the source-backed slide API and keeps the related graph out of
    /// this slide-only structural reader.
    pub fn notes(&self) -> Result<Option<super::related::Notes>> {
        self.package.check()?;
        if let Some(slide) = self.physical_source_slide() {
            let result = slide.notes();
            self.package.check()?;
            return result;
        }
        self.reject_related_edge(
            NOTES_SLIDE_REL,
            "notesSlide",
            "effective primary slide notes are outside the dependency-free topology",
        )?;
        self.package.check()?;
        Ok(None)
    }

    /// The chart-space XML of the ordinary chart this slide relates through
    /// `relationship_id`, or `None` when it has no such chart.
    ///
    /// A restored slide carries no chart relationships (they are dependency
    /// edges that slide removal refuses), so only a physical source slide
    /// can answer `Some`.
    pub fn chart_xml(&self, relationship_id: &str) -> Result<Option<Vec<u8>>> {
        self.package.check()?;
        let result = match self.physical_source_slide() {
            Some(slide) => slide.chart_xml(relationship_id),
            None => Ok(None),
        };
        self.package.check()?;
        result
    }

    /// The Office Math equations of this slide's text (record 0321); see
    /// [`crate::slide::Slide::equations`]. Like charts, only a physical
    /// source slide can answer them; a restored slide answers none.
    pub fn equations(&self) -> Result<Vec<crate::shape::math::Equation>> {
        self.package.check()?;
        let result = match self.physical_source_slide() {
            Some(slide) => slide.equations(),
            None => Ok(Vec::new()),
        };
        self.package.check()?;
        result
    }

    /// Resolve a click action of this slide's shapes or text (record
    /// 0321); see [`crate::slide::Slide::hyperlink`]. A slide target is
    /// reported at its position in this revision. Like charts, only a
    /// physical source slide can answer `Some`.
    pub fn hyperlink(
        &self,
        click: &crate::hyperlinks::Click,
    ) -> Result<Option<crate::hyperlinks::Link>> {
        self.package.check()?;
        let result = match self.physical_source_slide() {
            Some(slide) => slide.hyperlink(click).map(|link| {
                link.and_then(|mut link| {
                    if let crate::hyperlinks::Target::Slide(physical) = link.target {
                        let logical = self.package.order.iter().position(|key| {
                            self.package.physical.get(key).copied() == Some(physical.get())
                        })?;
                        link.target = crate::hyperlinks::Target::Slide(Position::new(logical));
                    }
                    Some(link)
                })
            }),
            None => Ok(None),
        };
        self.package.check()?;
        result
    }

    /// The data-model XML of the SmartArt this slide relates through
    /// `relationship_id`, or `None` when it has no such diagram (record
    /// 0284). Like charts, only a physical source slide can answer `Some`.
    pub fn diagram_xml(&self, relationship_id: &str) -> Result<Option<Vec<u8>>> {
        self.package.check()?;
        let result = match self.physical_source_slide() {
            Some(slide) => slide.diagram_xml(relationship_id),
            None => Ok(None),
        };
        self.package.check()?;
        result
    }

    /// Acquire the effective slide's optional legacy comments snapshot.
    ///
    /// Legacy comments are a dependency of a slide removal and are therefore
    /// absent from every admitted restored slide. Existing physical source
    /// reads continue to use [`super::SourceSlide::comments`].
    pub fn comments(&self) -> Result<Option<super::related::Comments>> {
        self.package.check()?;
        if let Some(slide) = self.physical_source_slide() {
            let result = slide.comments();
            self.package.check()?;
            return result;
        }
        self.reject_related_edge(
            crate::comments::COMMENTS_REL,
            "comments",
            "effective primary slide comments are outside the dependency-free topology",
        )?;
        self.package.check()?;
        Ok(None)
    }

    /// Acquire the effective slide's optional modern-comments snapshot.
    ///
    /// Modern comments are a dependency of a slide removal and are therefore
    /// absent from every admitted restored slide. Existing physical source
    /// reads continue to use [`super::SourceSlide::modern_comments`].
    pub fn modern_comments(&self) -> Result<Option<super::modern_comments::List>> {
        self.package.check()?;
        if let Some(slide) = self.physical_source_slide() {
            let result = slide.modern_comments();
            self.package.check()?;
            return result;
        }
        self.reject_related_edge(
            crate::modern_comments::MODERN_COMMENT_RELATIONSHIP_TYPE,
            "modern comments",
            "effective primary slide modern comments are outside the dependency-free topology",
        )?;
        self.package.check()?;
        Ok(None)
    }

    fn physical_source_slide(&self) -> Option<super::SourceSlide> {
        self.package.source_slide_for_key(&self.key)
    }

    fn read_layout(&self) -> Result<Option<super::related::Layout>> {
        let mut selected = None;
        for relationship in self.snapshot.closure.slide.slide.relationships.iter() {
            if !crate::parts::is_relationship_type(
                &relationship.kind,
                rt::SLIDE_LAYOUT,
                "slideLayout",
            ) {
                continue;
            }
            if selected.is_some() {
                return Err(Error::Relationship(
                    "effective primary slide has multiple layout relationships".into(),
                ));
            }
            selected = Some(relationship);
        }
        let Some(relationship) = selected else {
            return Ok(None);
        };
        if relationship.mode == TargetMode::External {
            return Err(Error::Relationship(
                "effective primary slide layout relationship must be internal".into(),
            ));
        }
        let target = PackURI::from_rel_ref(
            self.snapshot.closure.slide.slide.uri.base_uri(),
            &relationship.target,
        )
        .map_err(|_| {
            Error::Relationship("effective primary slide layout target is invalid".into())
        })?;
        self.read_layout_part(&target, &relationship.id).map(Some)
    }

    fn read_layout_part(
        &self,
        target: &PackURI,
        relationship_id: &str,
    ) -> Result<super::related::Layout> {
        self.package.source.package.check_execution()?;
        let result =
            super::related::layout_for_target(&self.package.source, target, relationship_id);
        self.package.source.package.check_execution()?;
        result
    }

    fn reject_related_edge(
        &self,
        relationship_type: &str,
        local_type: &str,
        detail: &str,
    ) -> Result<()> {
        if self
            .snapshot
            .closure
            .slide
            .slide
            .relationships
            .iter()
            .any(|relationship| {
                crate::parts::is_relationship_type(
                    &relationship.kind,
                    relationship_type,
                    local_type,
                )
            })
        {
            return Err(Error::Relationship(detail.into()));
        }
        Ok(())
    }

    fn with_part<T>(&self, read: impl FnOnce(&dyn litchi_opc::Part) -> Result<T>) -> Result<T> {
        self.package.check()?;
        let result = (|| {
            // read_snapshot already admitted this immutable payload against
            // the package's physical XML limits; no second full scan is needed.
            let binding = &self.snapshot.closure.slide.slide;
            match &self.snapshot.xml {
                SourcePayload::Original(data) => {
                    let part = super::SourcePart::from_binding(binding, data.clone())?;
                    super::validate_source_slide_root(&part)?;
                    read(&part)
                },
                SourcePayload::Edited(bytes) | SourcePayload::ReservedEdit { bytes, .. } => {
                    // Text/name readers do not follow relationships. The snapshot
                    // keeps the payload reservation alive throughout this borrowed read.
                    let part = litchi_opc::XmlPart::new_shared(
                        binding.uri.clone(),
                        binding.content_type.clone(),
                        Arc::clone(bytes),
                    );
                    read(&part)
                },
            }
        })();
        self.package.check()?;
        result
    }

    /// Read an owned shape scene without losing managed source reservations.
    ///
    /// Managed reads reserve `64 * input_bytes + 1024` bytes of parsing and
    /// retention headroom, capped at the native edit's default patch budget.
    /// Callers with a smaller remaining budget use
    /// [`Self::scene_with_max_patch_bytes`] instead.
    pub fn scene(&self) -> Result<crate::shape::owned::Scene> {
        self.scene_with_max_patch_bytes(crate::opened::Limits::default().max_patch_bytes())
    }

    /// Read an owned shape scene under a caller-selected retained-memory
    /// ceiling. Facade edits pass their remaining snapshot budget here so a
    /// source-backed planning read cannot reserve the native default after
    /// the edit has already been bounded more tightly.
    pub fn scene_with_max_patch_bytes(&self, maximum: usize) -> Result<crate::shape::owned::Scene> {
        let maximum = maximum.min(self.snapshot.max_output_bytes);
        self.package.check()?;
        crate::source_limits::validate_xml_limits(
            self.snapshot.xml.as_bytes(),
            self.package.owner.limits,
        )?;
        if let SourcePayload::Original(data) = &self.snapshot.xml {
            let part =
                super::SourcePart::from_binding(&self.snapshot.closure.slide.slide, data.clone())?;
            super::validate_source_slide_root(&part)?;
        }
        let scene_reservation =
            reserve_scene(&self.package, self.snapshot.xml.as_bytes().len(), maximum)?;
        let parsed = crate::shape::Scene::read(
            self.snapshot.xml.as_bytes(),
            crate::open_options::selected(self.snapshot.semantic_limits.as_ref()).shape(),
        )?;
        let scene = match &self.snapshot.xml {
            SourcePayload::Original(data) => {
                crate::shape::owned::Scene::from_source(parsed, data.clone())?
            },
            SourcePayload::Edited(data) => {
                crate::shape::owned::Scene::from_parsed(parsed, Arc::clone(data))?
            },
            SourcePayload::ReservedEdit { bytes, reservation } => {
                crate::shape::owned::Scene::from_reserved(
                    parsed,
                    Arc::clone(bytes),
                    Arc::clone(reservation),
                )?
            },
        };
        let scene = if let Some(reservation) = scene_reservation {
            let retained_bytes = scene.retained_bytes();
            let retained = u64::try_from(retained_bytes).map_err(|_| {
                limit(
                    crate::error::CodecLimitKind::SourceBytes,
                    usize::MAX,
                    usize::MAX,
                )
            })?;
            if retained > reservation.amount() {
                return Err(limit(
                    crate::error::CodecLimitKind::SourceBytes,
                    retained_bytes,
                    usize::try_from(reservation.amount()).unwrap_or(usize::MAX),
                ));
            }
            scene.with_scene_reservation(Arc::new(reservation))?
        } else {
            scene
        };
        self.package.check()?;
        Ok(scene)
    }
}

/// Isolated edit of shape creation, properties and removal, slide order,
/// and dependency-free slide removal.
pub struct Edit {
    source: Package,
    working: Package,
    changes: Vec<SourceBackedSlidePatch>,
    shape_removals: bool,
    max_patch_bytes: usize,
}
impl Edit {
    /// Insert an empty slide at a checked logical position using an exact
    /// registered physical layout name.
    ///
    /// The source-backed transaction stages the new slide as one structural
    /// topology revision. The slide Part and its layout relationship remain
    /// lazy OPC additions, so this operation does not materialize or reopen a
    /// package. The returned handle reads from the staged effective revision;
    /// dropping the edit still discards the candidate atomically.
    ///
    /// The position may equal the current slide count. Layout names use the
    /// same exact, case-sensitive matching rule as the eager editor. A missing
    /// or ambiguous name, an unsupported source boundary, a protected or
    /// signed package, and a failed resource admission leave this edit's
    /// previous staged state unchanged.
    pub fn insert_slide(&mut self, position: Position, layout_name: &str) -> Result<Slide> {
        self.working.check()?;
        let topology_budget =
            topology::topology_patch_budget(self.max_patch_bytes, self.shape_removals)?;
        let candidate = topology::insert(&self.working, position, layout_name, topology_budget)
            .map_err(insertion::map_error)?;
        let slide = candidate.presentation()?.slide(position)?;
        self.working = candidate;
        Ok(slide)
    }

    /// Add one plain DrawingML text box to an existing effective slide.
    ///
    /// The slide payload is staged through the native source-backed shape
    /// authoring path. The operation is failure-atomic, composes with earlier
    /// text/order/topology edits, and retains its managed resource admission
    /// through local and durable publication.
    pub fn add_text_box(
        &mut self,
        slide: impl Into<Position>,
        text: impl AsRef<str>,
        bounds: (i64, i64, i64, i64),
    ) -> Result<u32> {
        let logical = slide.into();
        let text = text.as_ref();
        self.working.check()?;
        let logical_position = logical.get();
        let target_key = self.working.order.get(logical_position).cloned().ok_or(
            Error::SlideIndexOutOfBounds {
                index: logical_position,
                len: self.working.order.len(),
            },
        )?;
        let snapshot = self.working.snapshot_logical(logical_position)?;
        let immutable_before = if self.working.topology.is_none() {
            let physical = self.source.physical_position(&target_key)?;
            Some(self.source.snapshot(physical)?)
        } else {
            None
        };
        let before_len = immutable_before
            .as_ref()
            .map(|before| before.xml.as_bytes().len())
            .unwrap_or_else(|| snapshot.xml.as_bytes().len());
        let topology_budget = if self.working.topology.is_some() {
            Some(topology::topology_patch_budget(
                self.max_patch_bytes,
                self.shape_removals,
            )?)
        } else {
            None
        };
        let output_limit = if self.working.topology.is_some() {
            topology::replacement_after_budget(
                &self.working,
                &target_key,
                topology_budget.ok_or(Error::PatchConflict)?,
            )?
        } else {
            self.local_after_budget(&target_key, before_len, self.shape_removals)?
        };
        let mut edit = snapshot.edit_checked_with_max_patch_bytes(output_limit)?;
        let shape_id = edit.add_text_box(text, bounds)?;
        let after = edit.commit_checked()?.snapshot;
        if after.xml.as_bytes().len() > output_limit {
            return Err(limit(
                crate::error::CodecLimitKind::OutputBytes,
                after.xml.as_bytes().len(),
                output_limit,
            ));
        }
        if self.working.topology.is_none() {
            self.check_local_patch_budget(
                &target_key,
                before_len,
                after.xml.as_bytes().len(),
                self.shape_removals,
            )?;
        }
        let candidate = self.working.with_snapshot_with_max_patch_bytes(
            after.clone(),
            topology_budget.unwrap_or(self.max_patch_bytes),
        )?;
        if self.working.topology.is_some() {
            self.working = candidate;
            return Ok(shape_id);
        }
        let before = immutable_before.ok_or(Error::PatchConflict)?;
        let mut changes = Vec::new();
        changes
            .try_reserve_exact(self.changes.len() + 1)
            .map_err(allocation)?;
        changes.extend(
            self.changes
                .iter()
                .filter(|change| {
                    source_key_for_position(&self.source, change.before.position)
                        .map_or(true, |key| key != target_key)
                })
                .cloned(),
        );
        if !before.same_source(&after) {
            changes.push(SourceBackedSlidePatch { before, after });
        }
        if changes.len() > MAX_SOURCE_BACKED_SLIDE_BATCH {
            return Err(limit(
                crate::error::CodecLimitKind::Slides,
                changes.len(),
                MAX_SOURCE_BACKED_SLIDE_BATCH,
            ));
        }
        let root_bytes = self.working.root.edited_bytes();
        check_payload_budget_with_limit(
            root_bytes.as_deref().map(Vec::as_slice),
            changes
                .iter()
                .flat_map(|change| [&change.before, &change.after]),
            self.max_patch_bytes,
        )?;
        self.working.check()?;
        self.changes = changes;
        self.working = candidate;
        Ok(shape_id)
    }

    /// Remove one slide from the currently staged logical order, with the
    /// Parts it owns exclusively: its notes slide, pictures, diagrams, VML
    /// drawings, comments and embedded objects (record 0303). A Part that
    /// another owner still targets stays. Admission is delegated to the
    /// shared source-native planner; the resulting candidate retains a
    /// prepared effective topology and can therefore be read before it is
    /// published. The final slide, referenced slides, and unmodeled surfaces
    /// are refused. Durable inversion requires canonical owner-relative
    /// relationship targets for the presentation binding and every internal
    /// edge of the slide and its owned Parts, and refuses an edge to a
    /// missing Part, which an inverse could not restore after reopening.
    ///
    /// ```
    /// # fn main() -> litchi_pptx::Result<()> {
    /// use std::sync::Arc;
    /// use litchi_core::{OwnedSource, Position};
    /// use litchi_pptx::{builder::Builder, source::Package};
    /// let mut builder = Builder::new()?;
    /// {
    ///     let presentation = builder.presentation()?;
    ///     presentation.add_slide()?;
    ///     presentation.add_slide()?;
    /// }
    /// let mut authored = builder.build()?;
    /// let source = Package::from_read_at(Arc::new(OwnedSource::new(authored.to_bytes()?)))?;
    /// let mut edit = source.edit()?;
    /// edit.remove_slide(Position::new(0))?;
    /// let (updated, patch) = edit.commit()?.into_parts();
    /// assert_eq!(updated.presentation()?.slide_count()?, 1);
    /// assert_eq!(
    ///     updated.apply(&patch.inverse())?.into_package().to_bytes()?,
    ///     source.to_bytes()?
    /// );
    /// # Ok(())
    /// # }
    /// ```
    pub fn remove_slide(&mut self, position: Position) -> Result<bool> {
        self.working.check()?;
        let logical = position.get();
        if logical >= self.working.order.len() {
            return Err(Error::SlideIndexOutOfBounds {
                index: logical,
                len: self.working.order.len(),
            });
        }
        if self.working.order.len() < 2 {
            return Err(Error::slide_removal_refused(
                crate::SlideRemovalRefusal::FinalSlide,
            ));
        }
        let key = self.working.order[logical].clone();
        let topology_budget =
            topology::topology_patch_budget(self.max_patch_bytes, self.shape_removals)?;
        let candidate = topology::remove(&self.working, &key, topology_budget)?;
        self.working = candidate;
        Ok(true)
    }

    /// Move one logical slide to another logical position in the currently
    /// staged order.
    pub fn move_slide(&mut self, from: Position, to: Position) -> Result<bool> {
        self.working.check()?;
        let from = from.get();
        let to = to.get();
        let length = self.working.order.len();
        if from >= length {
            return Err(Error::SlideIndexOutOfBounds {
                index: from,
                len: length,
            });
        }
        if to >= length {
            return Err(Error::SlideIndexOutOfBounds {
                index: to,
                len: length,
            });
        }
        if from == to {
            return Ok(false);
        }
        let mut order = self.working.order.to_vec();
        let key = order.remove(from);
        order.insert(to, key);
        let candidate = self.working.stage_order(
            order,
            Some(to),
            self.max_patch_bytes,
            &self.source,
            &self.changes,
            self.shape_removals,
        )?;
        let changed = candidate.order != self.working.order;
        self.working = candidate;
        Ok(changed)
    }

    /// Exchange two logical slide positions in the currently staged order.
    pub fn swap_slides(&mut self, left: Position, right: Position) -> Result<bool> {
        self.working.check()?;
        let left = left.get();
        let right = right.get();
        let length = self.working.order.len();
        if left >= length {
            return Err(Error::SlideIndexOutOfBounds {
                index: left,
                len: length,
            });
        }
        if right >= length {
            return Err(Error::SlideIndexOutOfBounds {
                index: right,
                len: length,
            });
        }
        if left == right {
            return Ok(false);
        }
        let mut order = self.working.order.to_vec();
        order.swap(left, right);
        let candidate = self.working.stage_order(
            order,
            None,
            self.max_patch_bytes,
            &self.source,
            &self.changes,
            self.shape_removals,
        )?;
        let changed = candidate.order != self.working.order;
        self.working = candidate;
        Ok(changed)
    }

    /// Replace one existing slide's authored `p:cSld@name` value.
    ///
    /// The edit is staged against the selected slide's exact source bytes,
    /// retains unrelated XML and relationships, and composes with earlier
    /// changes to the same slide. An equal name is an exact no-op.
    ///
    /// ```
    /// # fn main() -> litchi_pptx::Result<()> {
    /// use std::sync::Arc;
    /// use litchi_core::OwnedSource;
    /// use litchi_pptx::{builder::Builder, source::Package};
    /// let mut builder = Builder::new()?;
    /// builder.presentation()?.add_slide()?;
    /// let mut authored = builder.build()?;
    /// let source = Package::from_read_at(Arc::new(OwnedSource::new(authored.to_bytes()?)))?;
    /// let mut edit = source.edit()?;
    /// edit.set_slide_name(0, "Roadmap")?;
    /// let _commit = edit.commit()?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn set_slide_name(
        &mut self,
        slide: impl Into<Position>,
        name: impl AsRef<str>,
    ) -> Result<bool> {
        self.set_slide_properties(slide, Some(name.as_ref()), None)
    }

    /// Set whether one existing slide is hidden during presentation.
    ///
    /// Only the slide-root `p:sld@show` policy is changed. Unrelated root,
    /// shape, relationship, and opaque XML bytes remain source-owned.
    ///
    /// ```
    /// # fn main() -> litchi_pptx::Result<()> {
    /// use std::sync::Arc;
    /// use litchi_core::OwnedSource;
    /// use litchi_pptx::{builder::Builder, source::Package};
    /// let mut builder = Builder::new()?;
    /// builder.presentation()?.add_slide()?;
    /// let mut authored = builder.build()?;
    /// let source = Package::from_read_at(Arc::new(OwnedSource::new(authored.to_bytes()?)))?;
    /// let mut edit = source.edit()?;
    /// edit.set_slide_hidden(0, true)?;
    /// let _commit = edit.commit()?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn set_slide_hidden(&mut self, slide: impl Into<Position>, hidden: bool) -> Result<bool> {
        self.set_slide_properties(slide, None, Some(hidden))
    }

    /// Atomically set a slide's authored name and presentation visibility.
    ///
    /// Each omitted property remains unchanged. Both values share one source
    /// admission and lexical publication; failed staging retains earlier edits.
    ///
    /// ```
    /// # fn main() -> litchi_pptx::Result<()> {
    /// use std::sync::Arc;
    /// use litchi_core::OwnedSource;
    /// use litchi_pptx::{builder::Builder, source::Package};
    /// let mut builder = Builder::new()?;
    /// builder.presentation()?.add_slide()?;
    /// let bytes = builder.build()?.to_bytes()?;
    /// let source = Package::from_read_at(Arc::new(OwnedSource::new(bytes)))?;
    /// let mut edit = source.edit()?;
    /// edit.set_slide_properties(0, Some("Roadmap"), Some(true))?;
    /// let _commit = edit.commit()?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn set_slide_properties(
        &mut self,
        slide: impl Into<Position>,
        name: Option<&str>,
        hidden: Option<bool>,
    ) -> Result<bool> {
        if name.is_none() && hidden.is_none() {
            return Ok(false);
        }
        let changed = self.stage_slide_edit_with_marker(
            slide.into(),
            |_package, _snapshot| Ok(None),
            |edit| edit.set_slide_properties(name, hidden).map(usize::from),
            true,
        )? != 0;
        self.shape_removals |= changed;
        Ok(changed)
    }

    /// Remove one top-level slide shape, composing with earlier changes.
    ///
    /// Removing a group removes its entire subtree. Grouped descendants and
    /// removals leaving connector or animation references are refused. Unused
    /// relationships and parts are retained for exact source replay.
    ///
    /// ```no_run
    /// # fn example() -> litchi_pptx::Result<()> {
    /// let package = litchi_pptx::source::Package::open("deck.pptx")?;
    /// let mut edit = package.edit()?;
    /// edit.remove_shape(0, 0_usize)?;
    /// edit.commit()?.into_package().save("trimmed.pptx")?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn remove_shape<'k>(
        &mut self,
        slide: impl Into<Position>,
        shape: impl Into<crate::shape::Key<'k>>,
    ) -> Result<bool> {
        let shape = shape.into();
        let changed = self.stage_slide_edit_with_marker(
            slide.into(),
            |_package, _snapshot| Ok(None),
            |edit| edit.remove_shape(shape).map(usize::from),
            true,
        )? != 0;
        self.shape_removals |= changed;
        Ok(changed)
    }

    /// Replace text in one shape, composing with earlier staged changes.
    pub fn set_shape_text<'k>(
        &mut self,
        slide: impl Into<Position>,
        shape: impl Into<crate::shape::Key<'k>>,
        text: impl AsRef<str>,
    ) -> Result<bool> {
        let replacement = crate::opened::ShapeTextReplacement::new(shape.into(), text.as_ref());
        Ok(self.set_shape_texts(slide, std::slice::from_ref(&replacement))? != 0)
    }

    /// Atomically replace a bounded batch of shapes on one existing slide.
    pub fn set_shape_texts(
        &mut self,
        slide: impl Into<Position>,
        replacements: &[crate::opened::ShapeTextReplacement<'_>],
    ) -> Result<usize> {
        self.stage_slide_edit(
            slide.into(),
            |package, snapshot| reserve_edit(package, snapshot, replacements),
            move |edit| edit.set_shape_texts(replacements),
        )
    }

    /// Atomically replace text from several already bounded same-slide
    /// batches in one source scene read, rewrite, and read-back.
    ///
    /// Each chunk retains the 256-target and 8 MiB replacement-text limits.
    /// Globally disjoint selections and chunk count are bounded to 100,000;
    /// aggregate replacement text is bounded by the slide output-byte limit.
    pub fn set_shape_text_chunks(
        &mut self,
        slide: impl Into<Position>,
        chunks: &[&[crate::opened::ShapeTextReplacement<'_>]],
    ) -> Result<usize> {
        self.stage_slide_edit(
            slide.into(),
            |_package, _snapshot| Ok(None),
            |edit| edit.set_shape_text_chunks(chunks),
        )
    }

    /// Atomically replace text and bounds together on one effective slide.
    ///
    /// Each shape contributes one combined update. Failed validation leaves
    /// every earlier staged operation intact; successful changes participate
    /// in source-local and transported replay, including topology revisions.
    pub fn set_shape_properties(
        &mut self,
        slide: impl Into<Position>,
        replacements: &[crate::opened::ShapePropertiesReplacement<'_>],
    ) -> Result<usize> {
        self.stage_slide_edit(
            slide.into(),
            |_package, _snapshot| Ok(None),
            |edit| edit.set_shape_properties(replacements),
        )
    }

    /// Atomically replace text and/or bounds from several already bounded
    /// same-slide property batches in one source scene read, rewrite, and
    /// read-back.
    ///
    /// Each chunk retains the 256-target and 8 MiB replacement-text limits.
    /// Globally disjoint selections and chunk count are bounded to 100,000;
    /// aggregate replacement text is bounded by the slide output-byte limit.
    pub fn set_shape_property_chunks(
        &mut self,
        slide: impl Into<Position>,
        chunks: &[&[crate::opened::ShapePropertiesReplacement<'_>]],
    ) -> Result<usize> {
        self.stage_slide_edit(
            slide.into(),
            |_package, _snapshot| Ok(None),
            |edit| edit.set_shape_property_chunks(chunks),
        )
    }

    /// Apply a bounded batch of rich text edits to one effective slide.
    ///
    /// The edits have the semantics of
    /// [`crate::opened::Edit::set_text_properties`]; they compose with
    /// earlier staged changes and participate in source-local replay.
    pub fn set_text_properties(
        &mut self,
        slide: impl Into<Position>,
        edits: &[crate::opened::TextAssignment<'_>],
    ) -> Result<usize> {
        self.stage_slide_edit(
            slide.into(),
            |_package, _snapshot| Ok(None),
            |edit| edit.set_text_properties(edits),
        )
    }

    /// Insert a paragraph into a text body of one effective slide, with the
    /// semantics of [`crate::opened::Edit::insert_paragraph`]; it composes
    /// with earlier staged changes and participates in source-local replay.
    /// Notes-page bodies are refused (record 0333).
    ///
    /// # Errors
    ///
    /// As the eager verb, and [`Error::Refused`] for a notes-page body.
    pub fn insert_paragraph(
        &mut self,
        slide: impl Into<Position>,
        body: crate::opened::TextBody,
        position: Position,
        properties: crate::opened::TextProperties<'_>,
    ) -> Result<()> {
        self.change_text_body(
            slide.into(),
            body,
            &crate::opened::TextChange::InsertParagraph {
                position: position.get(),
                properties,
            },
        )
    }

    /// Remove a paragraph of a text body of one effective slide, with the
    /// semantics of [`crate::opened::Edit::remove_paragraph`].
    ///
    /// # Errors
    ///
    /// As the eager verb, and [`Error::Refused`] for a notes-page body.
    pub fn remove_paragraph(
        &mut self,
        slide: impl Into<Position>,
        body: crate::opened::TextBody,
        paragraph: Position,
    ) -> Result<()> {
        self.change_text_body(
            slide.into(),
            body,
            &crate::opened::TextChange::RemoveParagraph {
                paragraph: paragraph.get(),
            },
        )
    }

    /// Insert a text run into a paragraph of one effective slide, with the
    /// semantics of [`crate::opened::Edit::insert_run`].
    ///
    /// # Errors
    ///
    /// As the eager verb, and [`Error::Refused`] for a notes-page body.
    pub fn insert_run(
        &mut self,
        slide: impl Into<Position>,
        body: crate::opened::TextBody,
        paragraph: Position,
        position: Position,
        properties: crate::opened::TextProperties<'_>,
    ) -> Result<()> {
        self.change_text_body(
            slide.into(),
            body,
            &crate::opened::TextChange::InsertRun {
                paragraph: paragraph.get(),
                position: position.get(),
                properties,
            },
        )
    }

    /// Remove a run child of a paragraph of one effective slide, with the
    /// semantics of [`crate::opened::Edit::remove_run`].
    ///
    /// # Errors
    ///
    /// As the eager verb, and [`Error::Refused`] for a notes-page body.
    pub fn remove_run(
        &mut self,
        slide: impl Into<Position>,
        body: crate::opened::TextBody,
        paragraph: Position,
        run: Position,
    ) -> Result<()> {
        self.change_text_body(
            slide.into(),
            body,
            &crate::opened::TextChange::RemoveRun {
                paragraph: paragraph.get(),
                run: run.get(),
            },
        )
    }

    fn change_text_body(
        &mut self,
        slide: Position,
        body: crate::opened::TextBody,
        change: &crate::opened::TextChange<'_>,
    ) -> Result<()> {
        self.stage_slide_edit(
            slide,
            |_package, _snapshot| Ok(None),
            |edit| edit.change_text_body(body, change),
        )
        .map(|_| ())
    }

    /// Apply several bounded batches of rich text edits to one effective
    /// slide in one pass, with the semantics of
    /// [`crate::opened::Edit::set_text_property_chunks`] (record 0333).
    pub fn set_text_property_chunks(
        &mut self,
        slide: impl Into<Position>,
        chunks: &[&[crate::opened::TextAssignment<'_>]],
    ) -> Result<usize> {
        self.stage_slide_edit(
            slide.into(),
            |_package, _snapshot| Ok(None),
            |edit| edit.set_text_property_chunks(chunks),
        )
    }

    /// [`Self::set_text_property_chunks`], returning each edit's change in
    /// order (record 0348); see
    /// [`crate::opened::Edit::set_text_property_chunks_per_edit`].
    pub fn set_text_property_chunks_per_edit(
        &mut self,
        slide: impl Into<Position>,
        chunks: &[&[crate::opened::TextAssignment<'_>]],
    ) -> Result<Vec<bool>> {
        let mut changes = Vec::new();
        self.stage_slide_edit(
            slide.into(),
            |_package, _snapshot| Ok(None),
            |edit| edit.text_property_chunks(chunks, Some(&mut changes)),
        )?;
        Ok(changes)
    }

    /// Set (`Some`) or clear (`None`) one effective slide's direct
    /// transition: a set replaces the direct `p:transition` (or its
    /// supported markup-compatibility owner) and a clear removes it; it
    /// composes with earlier staged changes and participates in source-local
    /// replay.
    /// Returns whether the slide changed.
    ///
    /// # Errors
    ///
    /// Returns an error for an out-of-range slide, an unsupported or invalid
    /// transition, a modification-protected presentation, or an exhausted
    /// limit. A failure leaves this edit unchanged.
    pub fn set_slide_transition(
        &mut self,
        slide: impl Into<Position>,
        value: Option<&crate::transition::Transition>,
    ) -> Result<bool> {
        self.stage_slide_edit(
            slide.into(),
            |_package, _snapshot| Ok(None),
            |edit| {
                let changed = match value {
                    Some(value) => edit.set_transition(value)?,
                    None => edit.clear_transition()?,
                };
                Ok(usize::from(changed))
            },
        )
        .map(|changed| changed != 0)
    }

    fn stage_slide_edit(
        &mut self,
        logical: Position,
        reserve: impl FnOnce(&Package, &SourceBackedSlideSnapshot) -> Result<Option<Reservation>>,
        stage: impl FnOnce(&mut super::SourceBackedSlideEdit) -> Result<usize>,
    ) -> Result<usize> {
        self.stage_slide_edit_with_marker(logical, reserve, stage, self.shape_removals)
    }

    fn stage_slide_edit_with_marker(
        &mut self,
        logical: Position,
        reserve: impl FnOnce(&Package, &SourceBackedSlideSnapshot) -> Result<Option<Reservation>>,
        stage: impl FnOnce(&mut super::SourceBackedSlideEdit) -> Result<usize>,
        shape_removals: bool,
    ) -> Result<usize> {
        self.working.check()?;
        let logical_position = logical.get();
        let target_key = self.working.order.get(logical_position).cloned().ok_or(
            Error::SlideIndexOutOfBounds {
                index: logical_position,
                len: self.working.order.len(),
            },
        )?;
        let snapshot = self.working.snapshot_logical(logical_position)?;
        let immutable_before = if self.working.topology.is_none() {
            let physical = self.source.physical_position(&target_key)?;
            Some(self.source.snapshot(physical)?)
        } else {
            None
        };
        let before_len = immutable_before
            .as_ref()
            .map(|before| before.xml.as_bytes().len())
            .unwrap_or_else(|| snapshot.xml.as_bytes().len());
        let topology_budget = if self.working.topology.is_some() {
            Some(topology::topology_patch_budget(
                self.max_patch_bytes,
                shape_removals,
            )?)
        } else {
            None
        };
        let output_limit = if self.working.topology.is_some() {
            topology::replacement_after_budget(
                &self.working,
                &target_key,
                topology_budget.ok_or(Error::PatchConflict)?,
            )?
        } else {
            self.local_after_budget(&target_key, before_len, shape_removals)?
        };
        // The retained-patch ceiling is an exact wire bound, not a parser or
        // candidate-allocation budget.  Reserve source edit memory against
        // the snapshot's native/source policy; applying the exact patch
        // ceiling here made a small fitting patch reject when the unchanged
        // slide XML was larger than that ceiling.  The bounded edit below
        // still enforces the candidate output bound, and commit performs the
        // final exact envelope admission.
        let reservation = reserve(&self.working, &snapshot)?;
        let mut edit = snapshot.edit_checked_with_max_patch_bytes(output_limit)?;
        let changed = stage(&mut edit)?;
        if changed == 0 {
            self.working.check()?;
            return Ok(0);
        }
        let mut after = edit.commit_checked()?.snapshot;
        // The bounded edit has already performed the exact native rewrite,
        // but check its resulting payload before handing it to the package
        // revision builder.  Passing a clone into `with_snapshot` first used
        // to allocate the candidate XML before the facade ceiling was
        // observed; a rejected candidate must never escape that preflight.
        if after.xml.as_bytes().len() > output_limit {
            return Err(limit(
                crate::error::CodecLimitKind::OutputBytes,
                after.xml.as_bytes().len(),
                output_limit,
            ));
        }
        if let Some(reservation) = reservation {
            if let SourcePayload::Edited(bytes) = after.xml {
                after.xml = SourcePayload::ReservedEdit {
                    bytes,
                    reservation: Arc::new(reservation),
                };
            }
        }
        if self.working.topology.is_none() {
            self.check_local_patch_budget(
                &target_key,
                before_len,
                after.xml.as_bytes().len(),
                shape_removals,
            )?;
        }
        let candidate = self.working.with_snapshot_with_max_patch_bytes(
            after.clone(),
            topology_budget.unwrap_or(self.max_patch_bytes),
        )?;
        if self.working.topology.is_some() {
            self.working = candidate;
            return Ok(changed);
        }
        let before = immutable_before.ok_or(Error::PatchConflict)?;
        let mut changes = Vec::new();
        changes
            .try_reserve_exact(self.changes.len() + 1)
            .map_err(allocation)?;
        changes.extend(
            self.changes
                .iter()
                .filter(|change| {
                    source_key_for_position(&self.source, change.before.position)
                        .map_or(true, |key| key != target_key)
                })
                .cloned(),
        );
        if !before.same_source(&after) {
            changes.push(SourceBackedSlidePatch { before, after });
        }
        if changes.len() > MAX_SOURCE_BACKED_SLIDE_BATCH {
            return Err(limit(
                crate::error::CodecLimitKind::Slides,
                changes.len(),
                MAX_SOURCE_BACKED_SLIDE_BATCH,
            ));
        }
        let root_bytes = self.working.root.edited_bytes();
        check_payload_budget_with_limit(
            root_bytes.as_deref().map(Vec::as_slice),
            changes
                .iter()
                .flat_map(|change| [&change.before, &change.after]),
            self.max_patch_bytes,
        )?;
        self.working.check()?;
        self.changes = changes;
        self.working = candidate;
        Ok(changed)
    }

    /// Return the output-only remainder after accounting for the complete
    /// local durable envelope except the new slide's after image.  The native
    /// focused editor uses this bound before building candidate XML; the final
    /// exact envelope check below includes the after image itself.
    fn local_after_budget(
        &self,
        key: &StableSlideKey,
        before: usize,
        shape_removals: bool,
    ) -> Result<usize> {
        let length = self.local_patch_wire_len(key, before, 0, shape_removals)?;
        if length > self.max_patch_bytes {
            return Err(limit(
                crate::error::CodecLimitKind::PatchBytes,
                length,
                self.max_patch_bytes,
            ));
        }
        Ok(self.max_patch_bytes - length)
    }

    fn check_local_patch_budget(
        &self,
        key: &StableSlideKey,
        before: usize,
        after: usize,
        shape_removals: bool,
    ) -> Result<()> {
        let length = self.local_patch_wire_len(key, before, after, shape_removals)?;
        if length > self.max_patch_bytes {
            return Err(limit(
                crate::error::CodecLimitKind::PatchBytes,
                length,
                self.max_patch_bytes,
            ));
        }
        Ok(())
    }

    fn local_patch_wire_len(
        &self,
        target: &StableSlideKey,
        target_before: usize,
        target_after: usize,
        shape_removals: bool,
    ) -> Result<usize> {
        let mut changes = Vec::new();
        changes
            .try_reserve_exact(self.changes.len().saturating_add(1))
            .map_err(allocation)?;
        for change in &self.changes {
            let key = source_key_for_position(&self.source, change.before.position)?;
            if key == *target {
                continue;
            }
            changes.push((
                key,
                change.before.xml.as_bytes().len(),
                change.after.xml.as_bytes().len(),
            ));
        }
        changes.push((target.clone(), target_before, target_after));
        durable::local_wire_len(
            shape_removals,
            &self.source.order,
            &self.working.order,
            self.source.root.as_bytes().len(),
            self.working.root.as_bytes().len(),
            &changes,
        )
    }

    /// Publish a new immutable package, exact-source patch, and diagnostics.
    pub fn commit(mut self) -> Result<Commit> {
        self.working.check()?;
        let max_patch_bytes = self.max_patch_bytes;
        // A refused operation leaves the working package untouched.  A
        // no-op commit must still be publishable even when the caller's
        // selected patch ceiling is smaller than the durable envelope's
        // framing: no retained change is being emitted in this case.
        if self.working.topology.is_none()
            && self.changes.is_empty()
            && self.working.root.as_bytes() == self.source.root.as_bytes()
            && self.working.order.as_ref() == self.source.order.as_ref()
        {
            let before_root = self.source.root.clone();
            let before_order = self.source.order.clone();
            let after_root = self.working.root.clone();
            let after_order = self.working.order.clone();
            let Edit {
                source,
                working,
                changes,
                shape_removals,
                ..
            } = self;
            let patch = Patch {
                shape_removals,
                data: PatchData::Local {
                    source,
                    before_root,
                    before_order,
                    after_root,
                    after_order,
                    changes: changes.into(),
                },
                encoded_len_cache: Arc::new(OnceLock::new()),
            };
            return Ok(commit_of(working, patch));
        }
        if self.working.topology.is_some() {
            let source = self.source;
            let topology_budget =
                topology::topology_patch_budget(max_patch_bytes, self.shape_removals)?;
            let topology_patch =
                topology::patch_from_packages(&source, &self.working, topology_budget)?;
            let patch = Patch {
                shape_removals: self.shape_removals,
                data: PatchData::Topology(Arc::new(topology_patch)),
                encoded_len_cache: Arc::new(OnceLock::new()),
            };
            let encoded = patch.encoded_len()?;
            if encoded > max_patch_bytes {
                return Err(limit(
                    crate::error::CodecLimitKind::PatchBytes,
                    encoded,
                    max_patch_bytes,
                ));
            }
            return Ok(commit_of(self.working, patch));
        }
        self.changes
            .sort_unstable_by_key(|change| change.before.position);
        let source = self.source;
        let before_root = source.root.clone();
        let before_order = source.order.clone();
        let after_root = self.working.root.clone();
        let after_order = self.working.order.clone();
        let mut lengths = Vec::new();
        lengths
            .try_reserve_exact(self.changes.len())
            .map_err(allocation)?;
        for change in &self.changes {
            let key = source_key_for_position(&source, change.before.position)?;
            lengths.push((
                key,
                change.before.xml.as_bytes().len(),
                change.after.xml.as_bytes().len(),
            ));
        }
        let encoded = durable::local_wire_len(
            self.shape_removals,
            &before_order,
            &after_order,
            before_root.as_bytes().len(),
            after_root.as_bytes().len(),
            &lengths,
        )?;
        if encoded > max_patch_bytes {
            return Err(limit(
                crate::error::CodecLimitKind::PatchBytes,
                encoded,
                max_patch_bytes,
            ));
        }
        let patch = Patch {
            shape_removals: self.shape_removals,
            data: PatchData::Local {
                source,
                before_root,
                before_order,
                after_root,
                after_order,
                changes: self.changes.into(),
            },
            encoded_len_cache: Arc::new(OnceLock::from(encoded)),
        };
        Ok(commit_of(self.working, patch))
    }
}

/// Reversible exact-source slide and shape patch, with bounded durable encoding.
#[derive(Clone)]
pub struct Patch {
    data: PatchData,
    shape_removals: bool,
    encoded_len_cache: Arc<OnceLock<usize>>,
}
#[derive(Clone)]
enum PatchData {
    Local {
        source: Package,
        before_root: RootPayload,
        before_order: Arc<[StableSlideKey]>,
        after_root: RootPayload,
        after_order: Arc<[StableSlideKey]>,
        changes: Arc<[SourceBackedSlidePatch]>,
    },
    Durable(durable::Data),
    Topology(Arc<topology::Patch>),
}
impl Patch {
    /// Conservative wire-size charge, including before/after XML and framing.
    #[must_use]
    pub fn retained_bytes(&self) -> usize {
        self.encoded_len().unwrap_or(usize::MAX)
    }

    /// Whether the root, order and selected slides are unchanged.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        match &self.data {
            PatchData::Local {
                before_root,
                before_order,
                after_root,
                after_order,
                changes,
                ..
            } => {
                changes.is_empty()
                    && before_root.as_bytes() == after_root.as_bytes()
                    && before_order.as_ref() == after_order.as_ref()
            },
            PatchData::Durable(data) => {
                data.changes.is_empty()
                    && data.before_root.as_slice() == data.after_root.as_slice()
                    && data.before_order.as_ref() == data.after_order.as_ref()
            },
            PatchData::Topology(data) => data.is_noop(),
        }
    }

    fn encoded_len(&self) -> Result<usize> {
        if let PatchData::Local { source, .. } = &self.data {
            source.check()?;
        }
        if let PatchData::Topology(data) = &self.data {
            data.check()?;
        }
        if let Some(length) = self.encoded_len_cache.get() {
            return Ok(*length);
        }
        let length = match &self.data {
            PatchData::Local {
                source,
                before_root,
                before_order,
                after_root,
                after_order,
                changes,
            } => {
                let mut lengths = Vec::new();
                lengths
                    .try_reserve_exact(changes.len())
                    .map_err(allocation)?;
                for change in changes.iter() {
                    let key = source_key_for_position(source, change.before.position)?;
                    lengths.push((
                        key,
                        change.before.xml.as_bytes().len(),
                        change.after.xml.as_bytes().len(),
                    ));
                }
                durable::local_wire_len(
                    self.shape_removals,
                    before_order,
                    after_order,
                    before_root.as_bytes().len(),
                    after_root.as_bytes().len(),
                    &lengths,
                )?
            },
            PatchData::Durable(data) => durable::data_wire_len(data)?
                .checked_add(if self.shape_removals {
                    durable::REMOVAL_MARKER_BYTES
                } else {
                    0
                })
                .ok_or_else(|| {
                    limit(
                        crate::error::CodecLimitKind::PatchBytes,
                        usize::MAX,
                        usize::MAX,
                    )
                })?,
            PatchData::Topology(data) => data
                .encoded_len()?
                .checked_add(if self.shape_removals {
                    durable::REMOVAL_MARKER_BYTES
                } else {
                    0
                })
                .ok_or_else(|| {
                    limit(
                        crate::error::CodecLimitKind::PatchBytes,
                        usize::MAX,
                        usize::MAX,
                    )
                })?,
        };
        let _ = self.encoded_len_cache.set(length);
        Ok(length)
    }

    /// Reverse the root/order and selected slides’ exact before/after states.
    pub fn inverse(&self) -> Self {
        let data = match &self.data {
            PatchData::Local {
                source,
                before_root,
                before_order,
                after_root,
                after_order,
                changes,
            } => PatchData::Local {
                source: source.clone(),
                before_root: after_root.clone(),
                before_order: after_order.clone(),
                after_root: before_root.clone(),
                after_order: before_order.clone(),
                changes: changes
                    .iter()
                    .map(SourceBackedSlidePatch::inverse)
                    .collect(),
            },
            PatchData::Durable(data) => PatchData::Durable(data.inverse()),
            PatchData::Topology(data) => PatchData::Topology(Arc::new(data.inverse())),
        };
        Self {
            data,
            shape_removals: self.shape_removals,
            encoded_len_cache: Arc::clone(&self.encoded_len_cache),
        }
    }
}

pub use crate::Error;
/// The bounded dimension of the source-backed primary family.
pub use crate::opened::LimitKind;
pub use litchi_core::edit::Diagnostics;

/// A validated package, patch, and content-free diagnostics.
pub type Commit = litchi_core::edit::Commit<Package, Patch>;

fn commit_of(package: Package, patch: Patch) -> Commit {
    let changed = !patch.is_noop();
    Commit::new(package, patch, Diagnostics::new(changed))
}

fn slide_key(slide: &Arc<super::SourceSlideData>) -> StableSlideKey {
    stable_slide_key(
        slide.part_uri.clone(),
        slide.slide_id,
        slide.binding.slide_reference_id.clone(),
    )
}

fn source_order(owner: &SourceBackedPresentationEditor) -> Result<Box<[StableSlideKey]>> {
    let mut order = Vec::new();
    order
        .try_reserve_exact(owner.slides.len())
        .map_err(allocation)?;
    let mut seen = std::collections::HashSet::new();
    seen.try_reserve(owner.slides.len()).map_err(allocation)?;
    for slide in owner.slides.iter() {
        let key = slide_key(slide);
        if !seen.insert(key.clone()) {
            return Err(Error::PatchConflict);
        }
        order.push(key);
    }
    Ok(order.into_boxed_slice())
}

fn source_physical_index(
    owner: &SourceBackedPresentationEditor,
) -> Result<std::collections::HashMap<StableSlideKey, usize>> {
    let mut physical = std::collections::HashMap::new();
    physical
        .try_reserve(owner.slides.len())
        .map_err(allocation)?;
    for (position, slide) in owner.slides.iter().enumerate() {
        if physical.insert(slide_key(slide), position).is_some() {
            return Err(Error::PatchConflict);
        }
    }
    Ok(physical)
}

fn source_key_for_position(package: &Package, position: usize) -> Result<StableSlideKey> {
    package
        .owner
        .slides
        .get(position)
        .map(slide_key)
        .ok_or(Error::SlideIndexOutOfBounds {
            index: position,
            len: package.owner.slides.len(),
        })
}

fn check_payload_budget<'a>(
    root: Option<&[u8]>,
    slides: impl Iterator<Item = &'a SourceBackedSlideSnapshot>,
) -> Result<()> {
    check_payload_budget_with_limit(
        root,
        slides,
        crate::opened::Limits::default().max_patch_bytes(),
    )
}

fn check_payload_budget_with_limit<'a>(
    root: Option<&[u8]>,
    mut slides: impl Iterator<Item = &'a SourceBackedSlideSnapshot>,
    maximum: usize,
) -> Result<()> {
    let root_bytes = root.map_or(0, <[u8]>::len);
    if root_bytes > maximum {
        return Err(limit(
            crate::error::CodecLimitKind::RetainedBytes,
            root_bytes,
            maximum,
        ));
    }
    slides.try_fold(root_bytes, |sum, slide| {
        let total = sum.checked_add(slide.xml.as_bytes().len()).ok_or_else(|| {
            limit(
                crate::error::CodecLimitKind::RetainedBytes,
                usize::MAX,
                maximum,
            )
        })?;
        if total > maximum {
            return Err(limit(
                crate::error::CodecLimitKind::RetainedBytes,
                total,
                maximum,
            ));
        }
        Ok(total)
    })?;
    Ok(())
}

fn allocation(source: std::collections::TryReserveError) -> Error {
    Error::Allocation {
        resource: "source-backed primary edit metadata",
        source,
    }
}
fn limit(kind: crate::error::CodecLimitKind, observed: usize, maximum: usize) -> Error {
    Error::Limit {
        kind: LimitKind::Codec(kind),
        observed,
        maximum,
    }
}

macro_rules! redacted_debug {
    ($($ty:ident),+ $(,)?) => { $(impl std::fmt::Debug for $ty {
        fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
            f.debug_struct(stringify!($ty)).finish_non_exhaustive()
        }
    })+ };
}
redacted_debug!(Package, Presentation, Slide, Edit, Patch);

#[cfg(test)]
mod deferred_policy_tests {
    use std::sync::Arc;

    use litchi_opc::SourceCacheLimits;

    use super::Package;

    fn fixture() -> Vec<u8> {
        super::super::tests::svg_picture_pptx()
    }

    #[test]
    fn deferred_svg_materialization_retains_mixed_profiles_and_cache_policy() {
        let shape_defaults = crate::shape::Limits::DEFAULT;
        let shape_limits = crate::shape::Limits::new(
            shape_defaults.input_bytes(),
            shape_defaults.output_bytes(),
            shape_defaults.depth(),
            shape_defaults.nodes(),
            1,
            shape_defaults.retained_text_bytes(),
        )
        .unwrap();
        let notes_defaults = crate::notes::Limits::DEFAULT;
        let notes_limits = crate::notes::Limits {
            max_nodes: notes_defaults.max_nodes - 1,
            ..notes_defaults
        };
        let events_defaults = crate::presentation_properties::metadata::events::Limits::DEFAULT;
        let events_limits = crate::presentation_properties::metadata::events::Limits::new(
            events_defaults.max_slide_xml_bytes(),
            events_defaults.max_total_slide_xml_bytes(),
            events_defaults.max_events() - 1,
            events_defaults.max_xml_nodes(),
            events_defaults.max_xml_depth(),
            events_defaults.max_event_string_bytes(),
            events_defaults.max_output_bytes(),
        )
        .unwrap();
        let cache_limits = SourceCacheLimits::new(64 * 1024, 1).unwrap();
        let options = crate::OpenOptions::default()
            .cache_limits(cache_limits)
            .shape_limits(shape_limits)
            .notes_limits(notes_limits)
            .events_limits(events_limits);
        let selected = options
            .semantic_limits
            .clone()
            .expect("selected semantic policy");
        let source = Package::from_read_at_with_options(
            Arc::new(litchi_core::OwnedSource::new(fixture())),
            options,
        )
        .unwrap();

        let mut edit = source.edit_svg_images().unwrap();
        assert!(
            edit.replace(
                0,
                0,
                crate::source::svg::Replacement::new(
                    br#"<svg xmlns="http://www.w3.org/2000/svg"><path d="M0 0"/></svg>"#.to_vec(),
                    b"\x89PNGchanged".to_vec(),
                ),
            )
            .unwrap()
        );
        let updated = edit.commit().unwrap().into_package();

        // The public write takes the deferred stream route. The effective
        // owner check then forces the same artifact through materialization.
        assert!(!updated.to_bytes().unwrap().is_empty());
        let effective = updated.effective().unwrap();
        assert!(Arc::ptr_eq(
            &selected,
            effective.owner.semantic_limits.as_ref().unwrap(),
        ));
        assert_eq!(effective.owner.package.cache_limits(), cache_limits);

        // Make the package non-pristine, then exercise the SVG physical
        // reopen path rather than only reading the already-materialized owner.
        let mut shape_edit = updated.edit().unwrap();
        assert!(shape_edit.set_slide_name(0, "physical reopen").unwrap());
        let nonpristine = shape_edit.commit().unwrap().into_package();
        let reopened = super::svg::physical(&nonpristine).unwrap();
        assert!(Arc::ptr_eq(
            &selected,
            reopened.owner.semantic_limits.as_ref().unwrap(),
        ));
        assert_eq!(reopened.owner.package.cache_limits(), cache_limits);
        let image = nonpristine.edit_svg_images().unwrap().image(0, 0).unwrap();
        assert_eq!(image.raster_bytes(), b"\x89PNGchanged");

        // The selected shape ceiling rejects a second shape; the default
        // profile admits the same public edit.
        let mut selected_edit = nonpristine.edit().unwrap();
        assert!(matches!(
            selected_edit.add_text_box(0, "selected policy", (0, 0, 914_400, 914_400)),
            Err(crate::Error::Limit {
                kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::Shapes),
                observed: 2,
                maximum: 1,
            })
        ));
        let default =
            Package::from_read_at(Arc::new(litchi_core::OwnedSource::new(fixture()))).unwrap();
        let mut default_edit = default.edit().unwrap();
        assert!(
            default_edit
                .add_text_box(0, "default policy", (0, 0, 914_400, 914_400))
                .is_ok()
        );
    }
}

// Charge conservative input/output and escaping headroom before lowering. The
// token follows the candidate bytes through patches and detached shape handles.
fn reserve_edit(
    package: &Package,
    snapshot: &SourceBackedSlideSnapshot,
    replacements: &[crate::opened::ShapeTextReplacement<'_>],
) -> Result<Option<Reservation>> {
    if replacements.is_empty() {
        return Ok(None);
    }
    let Some(context) = package.owner.package.execution_context() else {
        return Ok(None);
    };
    let maximum = snapshot.max_output_bytes;
    let initial = snapshot
        .xml
        .as_bytes()
        .len()
        .checked_mul(2)
        .ok_or_else(|| {
            limit(
                crate::error::CodecLimitKind::SourceBytes,
                usize::MAX,
                maximum,
            )
        })?;
    if initial > maximum {
        return Err(limit(
            crate::error::CodecLimitKind::SourceBytes,
            initial,
            maximum,
        ));
    }
    let bytes = replacements.iter().try_fold(initial, |sum, replacement| {
        replacement
            .text()
            .len()
            .checked_mul(6)
            .and_then(|size| size.checked_add(256))
            .ok_or_else(|| {
                limit(
                    crate::error::CodecLimitKind::SourceBytes,
                    usize::MAX,
                    maximum,
                )
            })
            .and_then(|size| {
                let total = sum.checked_add(size).ok_or_else(|| {
                    limit(
                        crate::error::CodecLimitKind::SourceBytes,
                        usize::MAX,
                        maximum,
                    )
                })?;
                if total > maximum {
                    Err(limit(
                        crate::error::CodecLimitKind::SourceBytes,
                        total,
                        maximum,
                    ))
                } else {
                    Ok(total)
                }
            })
    })?;
    let bytes = u64::try_from(bytes).map_err(|_| {
        limit(
            crate::error::CodecLimitKind::SourceBytes,
            usize::MAX,
            maximum,
        )
    })?;
    context
        .reserve(litchi_core::Resource::Memory, bytes)
        .map(Some)
        .map_err(|error| {
            Error::Opc(match error {
                litchi_core::ExecutionError::Cancelled => litchi_opc::OpcError::Cancelled,
                error => litchi_opc::OpcError::Execution(error),
            })
        })
}

// Reserve parser/index/MCE retention headroom before processing. Refuse an
// unexpectedly larger retained scene before it can escape the admission charge.
fn reserve_scene(
    package: &Package,
    input_bytes: usize,
    maximum: usize,
) -> Result<Option<Reservation>> {
    let Some(context) = package.owner.package.execution_context() else {
        return Ok(None);
    };
    let bytes = input_bytes
        .checked_mul(64)
        .and_then(|bytes| bytes.checked_add(1024))
        .ok_or_else(|| {
            limit(
                crate::error::CodecLimitKind::SourceBytes,
                usize::MAX,
                maximum,
            )
        })?;
    if bytes > maximum {
        return Err(limit(
            crate::error::CodecLimitKind::SourceBytes,
            bytes,
            maximum,
        ));
    }
    let amount = u64::try_from(bytes).map_err(|_| {
        limit(
            crate::error::CodecLimitKind::SourceBytes,
            usize::MAX,
            maximum,
        )
    })?;
    context
        .reserve(litchi_core::Resource::Memory, amount)
        .map(Some)
        .map_err(|error| {
            Error::Opc(match error {
                litchi_core::ExecutionError::Cancelled => litchi_opc::OpcError::Cancelled,
                error => litchi_opc::OpcError::Execution(error),
            })
        })
}
