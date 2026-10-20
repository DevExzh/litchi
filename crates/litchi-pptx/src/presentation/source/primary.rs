//! Immutable deferred-source ownership, slide ordering and shape-property edits.
//!
//! Packages share the admitted source and retain changed root and slide XML.
//! Patches transport exact root/order context and selected slide preimages.
//! Structural revisions additionally retain one prepared OPC topology and a
//! stable-key recipe for lazy readback and reversible publication.

mod durable;
mod insertion;
mod topology;
mod topology_admission;
mod topology_digest;
mod topology_wire;

use std::io::Write;
use std::path::Path;
use std::sync::Arc;

use litchi_core::{Position, ReadAt, Reservation};
use litchi_opc::constants::relationship_type as rt;
use litchi_opc::{PackURI, PartData, TargetMode};

use super::{
    MAX_SOURCE_BACKED_SLIDE_BATCH, SourceBackedPresentation, SourceBackedPresentationEditor,
    SourceBackedSlidePatch, SourceBackedSlideSnapshot, SourcePayload,
};
use crate::{Error, Result};

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
/// let replayed = reopened.apply(&transported)?;
/// let restored = replayed.apply(&transported.inverse())?;
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
        crate::presentation::first_slide_number(self.root.as_bytes())
    }

    /// Retained reference-compatible read repairs.
    pub fn tolerance_report(&self) -> &crate::leniency::ToleranceReport {
        &self.tolerance
    }

    /// Obtain an owned read-only semantic projection.
    pub fn presentation(&self) -> Result<Presentation> {
        self.check()?;
        Ok(Presentation {
            package: self.clone(),
        })
    }

    /// Begin an isolated primary edit; dropping it leaves this package intact.
    pub fn edit(&self) -> Result<Edit> {
        self.check()?;
        Ok(Edit {
            source: self.clone(),
            working: self.clone(),
            changes: Vec::new(),
            shape_removals: false,
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
        self.owner.package.source_lineage() == other.owner.package.source_lineage()
            && self.owner._presentation.source_partname()
                == other.owner._presentation.source_partname()
            && self.owner._presentation.source_data().as_bytes()
                == other.owner._presentation.source_data().as_bytes()
    }

    /// Compare revisions of the same retained source without reading untouched parts.
    /// Independently opened sources are distinct even when their bytes match.
    pub fn same_content(&self, other: &Self) -> Result<bool> {
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

    /// Replay an exact-source patch without changing this package.
    pub fn apply(&self, patch: &Patch) -> Result<Self> {
        self.check()?;
        let (source, before_root, before_order, after_root, after_order, changes) = match &patch
            .data
        {
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
            PatchData::Durable(data) => return durable::apply(self, data, patch.shape_removals),
            PatchData::Topology(data) => return data.apply(self, patch.shape_removals),
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
        Ok(candidate)
    }

    /// Serialize with exact no-op and raw-copy preservation of untouched members.
    pub fn write_to(&self, writer: impl Write) -> Result<()> {
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

    fn with_snapshot(&self, mut snapshot: SourceBackedSlideSnapshot) -> Result<Self> {
        self.check()?;
        if self.topology.is_some() {
            return topology::replace_snapshot(self, snapshot);
        }
        if snapshot.xml.as_bytes().len() > snapshot.max_output_bytes {
            return Err(limit(
                "source-backed primary slide output bytes",
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
                "source-backed primary changed slides",
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
        self.check()?;
        if self.topology.is_some() {
            let bytes = root.edited_bytes().ok_or(Error::PatchConflict)?;
            return topology::reorder(self, bytes, order);
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
                "source-backed primary presentation output bytes",
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
    fn stage_order(&self, ordered: Vec<StableSlideKey>, moved: Option<usize>) -> Result<Self> {
        self.check()?;
        let ordered: Arc<[StableSlideKey]> = ordered.into();
        if ordered == self.order {
            return Ok(self.clone());
        }
        if crate::opened::sections::present(self.root.as_bytes()) {
            return self.stage_sectioned_order(ordered, moved);
        }
        if self.topology.is_some() {
            let _reservation = reserve_managed_output(
                self.owner.package.execution_context().as_ref(),
                self.root.as_bytes().len(),
            )?;
            let xml = reorder_stable_bindings(
                self.root.as_bytes(),
                &self.order,
                &ordered,
                self.owner.package.execution_context().as_ref(),
            )?;
            return topology::reorder(self, Arc::new(xml), ordered);
        }
        let source_order = source_order(&self.owner)?;
        let source_xml = self.owner._presentation.source_data().as_bytes();
        let root = if ordered.as_ref() == source_order.as_ref() {
            RootPayload::Original(self.owner._presentation.source_data().clone())
        } else {
            validate_source_order_root(
                &self.owner.package,
                self.owner._presentation.source_partname(),
                source_xml,
                &source_order,
                self.owner.package.execution_context().as_ref(),
            )?;
            let reservation = reserve_managed_output(
                self.owner.package.execution_context().as_ref(),
                source_xml.len(),
            )?;
            let xml = reorder_stable_bindings(
                source_xml,
                &source_order,
                &ordered,
                self.owner.package.execution_context().as_ref(),
            )?;
            match reservation {
                Some(reservation) => RootPayload::ReservedEdit {
                    bytes: Arc::new(xml),
                    _reservation: reservation,
                },
                None => RootPayload::Edited(Arc::new(xml)),
            }
        };
        self.with_order(root, ordered)
    }

    /// Stage an order on a root with a p14 section list. Section membership
    /// depends on the sequence of moves, so the candidate is derived from
    /// the current root rather than rebuilt from the source.
    fn stage_sectioned_order(
        &self,
        ordered: Arc<[StableSlideKey]>,
        moved: Option<usize>,
    ) -> Result<Self> {
        let context = self.owner.package.execution_context();
        validate_source_order_root(
            &self.owner.package,
            self.owner._presentation.source_partname(),
            self.root.as_bytes(),
            &self.order,
            context.as_ref(),
        )?;
        let reservation = reserve_managed_output(context.as_ref(), self.root.as_bytes().len())?;
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
            return topology::reorder(self, xml, ordered);
        }
        let root = match reservation {
            Some(reservation) => RootPayload::ReservedEdit {
                bytes: xml,
                _reservation: reservation,
            },
            None => RootPayload::Edited(xml),
        };
        self.with_order(root, ordered)
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
    /// retention headroom, capped at 128 MiB. This conservative policy can
    /// refuse large slides even when their final scene would fit the budget.
    pub fn scene(&self) -> Result<crate::shape::owned::Scene> {
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
        let scene_reservation = reserve_scene(&self.package, self.snapshot.xml.as_bytes().len())?;
        let parsed = crate::shape::Scene::read(self.snapshot.xml.as_bytes())?;
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
            let retained = u64::try_from(scene.retained_bytes())
                .map_err(|_| limit("source-backed primary scene bytes", usize::MAX))?;
            if retained > reservation.amount() {
                return Err(limit(
                    "source-backed primary scene bytes",
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
        let candidate =
            topology::insert(&self.working, position, layout_name).map_err(insertion::map_error)?;
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
        let mut edit = snapshot.edit_checked()?;
        let shape_id = edit.add_text_box(text, bounds)?;
        let after = edit.commit_checked()?.snapshot;
        let candidate = self.working.with_snapshot(after.clone())?;
        if self.working.topology.is_some() {
            self.working = candidate;
            return Ok(shape_id);
        }
        let physical = self.working.physical_position(&target_key)?;
        let before = self.source.snapshot(physical)?;
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
                "source-backed primary patch slides",
                MAX_SOURCE_BACKED_SLIDE_BATCH,
            ));
        }
        let root_bytes = self.working.root.edited_bytes();
        check_payload_budget(
            root_bytes.as_deref().map(Vec::as_slice),
            changes
                .iter()
                .flat_map(|change| [&change.before, &change.after]),
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
    /// assert_eq!(updated.apply(&patch.inverse())?.to_bytes()?, source.to_bytes()?);
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
            return Err(Error::SlideRemovalPlan {
                kind: crate::SlideRemovalRefusal::FinalSlide,
                detail: "a presentation must retain at least one slide".into(),
            });
        }
        let key = self.working.order[logical].clone();
        let candidate = topology::remove(&self.working, &key)?;
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
        let candidate = self.working.stage_order(order, Some(to))?;
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
        let candidate = self.working.stage_order(order, None)?;
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
        let changed = self.stage_slide_edit(
            slide.into(),
            |_package, _snapshot| Ok(None),
            |edit| edit.set_slide_properties(name, hidden).map(usize::from),
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
        let changed = self.stage_slide_edit(
            slide.into(),
            |_package, _snapshot| Ok(None),
            |edit| edit.remove_shape(shape).map(usize::from),
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
        edits: &[crate::opened::TextEdit<'_>],
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
        chunks: &[&[crate::opened::TextEdit<'_>]],
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
        chunks: &[&[crate::opened::TextEdit<'_>]],
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
    /// transition, with the semantics of
    /// [`super::SourceBackedSlideEdit::set_transition`] and
    /// [`super::SourceBackedSlideEdit::clear_transition`]; it composes with
    /// earlier staged changes and participates in source-local replay.
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
        self.working.check()?;
        let logical_position = logical.get();
        let target_key = self.working.order.get(logical_position).cloned().ok_or(
            Error::SlideIndexOutOfBounds {
                index: logical_position,
                len: self.working.order.len(),
            },
        )?;
        let snapshot = self.working.snapshot_logical(logical_position)?;
        let reservation = reserve(&self.working, &snapshot)?;
        let mut edit = snapshot.edit_checked()?;
        let changed = stage(&mut edit)?;
        if changed == 0 {
            self.working.check()?;
            return Ok(0);
        }
        let mut after = edit.commit_checked()?.snapshot;
        if let Some(reservation) = reservation {
            if let SourcePayload::Edited(bytes) = after.xml {
                after.xml = SourcePayload::ReservedEdit {
                    bytes,
                    reservation: Arc::new(reservation),
                };
            }
        }
        let candidate = self.working.with_snapshot(after.clone())?;
        if self.working.topology.is_some() {
            self.working = candidate;
            return Ok(changed);
        }
        let physical = self.working.physical_position(&target_key)?;
        let before = self.source.snapshot(physical)?;
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
                "source-backed primary patch slides",
                MAX_SOURCE_BACKED_SLIDE_BATCH,
            ));
        }
        let root_bytes = self.working.root.edited_bytes();
        check_payload_budget(
            root_bytes.as_deref().map(Vec::as_slice),
            changes
                .iter()
                .flat_map(|change| [&change.before, &change.after]),
        )?;
        self.working.check()?;
        self.changes = changes;
        self.working = candidate;
        Ok(changed)
    }

    /// Publish a new immutable package, exact-source patch, and diagnostics.
    pub fn commit(mut self) -> Result<Commit> {
        self.working.check()?;
        if self.working.topology.is_some() {
            let source = self.source;
            let topology_patch = topology::patch_from_packages(&source, &self.working)?;
            let patch = Patch {
                shape_removals: self.shape_removals,
                data: PatchData::Topology(Arc::new(topology_patch)),
            };
            patch.encoded_len()?;
            return Ok(Commit {
                package: self.working,
                patch,
            });
        }
        self.changes
            .sort_unstable_by_key(|change| change.before.position);
        let source = self.source;
        let before_root = source.root.clone();
        let before_order = source.order.clone();
        let after_root = self.working.root.clone();
        let after_order = self.working.order.clone();
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
        };
        patch.encoded_len()?;
        Ok(Commit {
            package: self.working,
            patch,
        })
    }
}

/// Reversible exact-source slide and shape patch, with bounded durable encoding.
#[derive(Clone)]
pub struct Patch {
    data: PatchData,
    shape_removals: bool,
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
        }
    }
}

/// Content-free primary commit facts.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Diagnostics {
    changed: bool,
}
impl Diagnostics {
    /// Whether the edit changed the package's visible slide XML.
    pub const fn changed(self) -> bool {
        self.changed
    }
}

/// A validated package, patch, and content-free diagnostics.
#[derive(Clone)]
pub struct Commit {
    package: Package,
    patch: Patch,
}
impl Commit {
    /// Borrow the resulting immutable package.
    pub const fn package(&self) -> &Package {
        &self.package
    }
    /// Borrow the exact-source patch.
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }
    /// Read content-free commit facts.
    pub fn diagnostics(&self) -> Diagnostics {
        Diagnostics {
            changed: !self.patch.is_noop(),
        }
    }
    /// Consume this commit into its package without replay or serialization.
    pub fn into_package(self) -> Package {
        self.package
    }
    /// Consume this commit into its package and patch.
    pub fn into_parts(self) -> (Package, Patch) {
        (self.package, self.patch)
    }
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
    mut slides: impl Iterator<Item = &'a SourceBackedSlideSnapshot>,
) -> Result<()> {
    let maximum = crate::opened::Limits::default().max_patch_bytes();
    let root_bytes = root.map_or(0, <[u8]>::len);
    if root_bytes > maximum {
        return Err(limit("source-backed primary retained root bytes", maximum));
    }
    slides.try_fold(root_bytes, |sum, slide| {
        sum.checked_add(slide.xml.as_bytes().len())
            .filter(|total| *total <= maximum)
            .ok_or_else(|| limit("source-backed primary retained slide bytes", maximum))
    })?;
    Ok(())
}

fn allocation(source: std::collections::TryReserveError) -> Error {
    Error::Allocation {
        resource: "source-backed primary edit metadata",
        source,
    }
}
fn limit(resource: &'static str, limit: usize) -> Error {
    Error::Limit { resource, limit }
}

macro_rules! redacted_debug {
    ($($ty:ident),+ $(,)?) => { $(impl std::fmt::Debug for $ty {
        fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
            f.debug_struct(stringify!($ty)).finish_non_exhaustive()
        }
    })+ };
}
redacted_debug!(Package, Presentation, Slide, Edit, Patch, Commit);

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
    let maximum = crate::opened::Limits::default().max_patch_bytes();
    let bytes = replacements.iter().try_fold(
        snapshot
            .xml
            .as_bytes()
            .len()
            .checked_mul(2)
            .ok_or_else(|| limit("source-backed primary edit bytes", maximum))?,
        |sum, replacement| {
            replacement
                .text()
                .len()
                .checked_mul(6)
                .and_then(|size| size.checked_add(256))
                .and_then(|size| sum.checked_add(size))
                .filter(|total| *total <= maximum)
                .ok_or_else(|| limit("source-backed primary edit bytes", maximum))
        },
    )?;
    let bytes =
        u64::try_from(bytes).map_err(|_| limit("source-backed primary edit bytes", maximum))?;
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
fn reserve_scene(package: &Package, input_bytes: usize) -> Result<Option<Reservation>> {
    let Some(context) = package.owner.package.execution_context() else {
        return Ok(None);
    };
    let maximum = crate::opened::Limits::default().max_patch_bytes();
    let bytes = input_bytes
        .checked_mul(64)
        .and_then(|bytes| bytes.checked_add(1024))
        .filter(|bytes| *bytes <= maximum)
        .ok_or_else(|| limit("source-backed primary scene bytes", maximum))?;
    context
        .reserve(litchi_core::Resource::Memory, bytes as u64)
        .map(Some)
        .map_err(|error| {
            Error::Opc(match error {
                litchi_core::ExecutionError::Cancelled => litchi_opc::OpcError::Cancelled,
                error => litchi_opc::OpcError::Execution(error),
            })
        })
}
