//! Typed `PresentationML` package facade and semantic package operations.

use litchi_opc::OpcPackage;
use litchi_opc::packuri::PackURI;

use crate::custom::{Host, Props};
use crate::parts::PresentationPart;
use crate::presentation::Presentation;
use crate::{Error, Result};

/// Main entry point for `PresentationML` package ownership.
pub struct Package {
    pub(crate) opc: std::sync::Arc<OpcPackage>,
    /// Aggregate ordinary-part payload bytes retained from ingress or the
    /// existing publication/capture pass. Replay updates this by the patch's
    /// changed-part delta instead of rescanning unrelated recipient parts.
    pub(crate) part_bytes: u64,
    /// Complete visible-content identity used by exact raw-patch no-ops.
    ///
    /// The slot is populated by opened captures or by the first no-op patch
    /// check. Keeping it separate from the payload-digest memo lets a raw
    /// patch prove its recorded base in constant time after the normal
    /// package-open/capture work, without making a no-op replay rewalk the
    /// recipient graph.
    pub(crate) content_revision: std::sync::Arc<std::sync::OnceLock<[u8; 32]>>,
    /// Whether ingress retained the exact source archive needed for physical
    /// revision authorization of topology-changing cross-package edits.
    pub(crate) physical_source_provenance: bool,
    /// Non-default semantic codec policies selected at ingress.  Keeping the
    /// override bundle optional leaves ordinary package owners at the same
    /// footprint while edits and source projections retain custom policy.
    pub(crate) semantic_limits: Option<crate::open_options::SemanticLimitsRef>,
    /// Every non-structural defect repaired while the package was opened, in
    /// discovery order: the OPC manifest repairs folded from
    /// `litchi_opc` ([`crate::Defect::Package`]), the slide-list repairs of
    /// audit finding 6.7, then the core- and custom-properties repairs folded
    /// from `litchi_ooxml_common` ([`crate::Defect::Properties`]). Repairs a
    /// lazy slide parse performs surface at that parse's own entry point.
    /// See [`crate::leniency`].
    pub(crate) tolerance: crate::leniency::ToleranceReport,
    /// Payload digests of `opc`'s parts, refreshed at every
    /// opened-presentation publication and offered as the parent memo of the
    /// next one.
    ///
    /// This is the accelerator that makes a durable-patch application cheap:
    /// without it every bare-patch replay re-hashes the complete package. It
    /// is keyed on payload allocation identity and retains each payload it
    /// names, so a memo left over from an earlier package state can only miss,
    /// never answer for bytes that no longer exist. Nothing observable depends
    /// on it: a miss is an ordinary hash.
    ///
    /// The slot is also filled, when empty, by a capture of the current graph
    /// (change 0751), which is why it is a cell: a capture takes `&self`. It
    /// is only ever set to a snapshot's memo re-projected onto `opc`'s own
    /// allocations, so every entry names and retains an allocation `opc`
    /// holds, and every mutation that publishes no opened-presentation
    /// snapshot still empties it.
    pub(crate) part_digests:
        std::sync::Arc<std::sync::OnceLock<std::sync::Arc<crate::opened::PartDigests>>>,
    /// Retained source-backed relationship closures for semantic replay.
    /// The index is updated by changed-part deltas, not rescanned on every
    /// package publication.
    pub(crate) replay_index: std::sync::Arc<crate::package::replay_index::ReplayIndex>,
    #[cfg(feature = "encryption")]
    pub(crate) encryption: litchi_ooxml_common::package_encryption::PackageEncryption,
    /// Limits of the encrypted container this package was opened under,
    /// also used for its encrypted output.
    #[cfg(feature = "encryption")]
    pub(crate) encryption_limits: litchi_crypto::ooxml::Limits,
}

impl std::fmt::Debug for Package {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Package")
            .field("part_count", &self.opc.part_count())
            .field("physical_source_retained", &self.physical_source_provenance)
            .finish_non_exhaustive()
    }
}

/// Package policy that must survive snapshot-derived edits without reopening.
#[derive(Clone, Default)]
pub(crate) struct OpenedOwnerState {
    tolerance: crate::leniency::ToleranceReport,
    pub(crate) semantic_limits: Option<crate::open_options::SemanticLimitsRef>,
    #[cfg(feature = "encryption")]
    encryption: litchi_ooxml_common::package_encryption::PackageEncryption,
    #[cfg(feature = "encryption")]
    encryption_limits: litchi_crypto::ooxml::Limits,
}

impl OpenedOwnerState {
    pub(crate) fn with_semantic_limits(
        semantic_limits: Option<crate::open_options::SemanticLimitsRef>,
    ) -> Self {
        Self {
            semantic_limits,
            ..Self::default()
        }
    }

    /// Refuse a typed mutation of a package with retained encryption
    /// provenance, as [`Package`]'s typed families do.
    pub(crate) fn ensure_plain_mutation(&self, operation: &'static str) -> Result<()> {
        let _ = operation;
        #[cfg(feature = "encryption")]
        self.encryption
            .ordinary_output()
            .map_err(|source| Error::EncryptionPolicy { operation, source })?;
        Ok(())
    }
}

impl Package {
    pub(crate) fn part_bytes_for(opc: &OpcPackage) -> Result<u64> {
        if let Some(bytes) = opc.admitted_part_bytes() {
            return Ok(bytes);
        }
        opc.iter_parts().try_fold(0_u64, |total, part| {
            let bytes = u64::try_from(part.payload_len()?).map_err(|_| {
                Error::internal(
                    "part_bytes_for",
                    "ordinary-part payload length exceeds the owner metric range",
                )
            })?;
            total.checked_add(bytes).ok_or_else(|| {
                Error::internal(
                    "part_bytes_for",
                    "ordinary-part payload aggregate exceeds the owner metric range",
                )
            })
        })
    }

    /// Complete visible-content identity for raw patch base matching.
    ///
    /// A package that has not yet entered the opened facade computes this once
    /// and retains the result when a replay needs an independent-content
    /// proof. Subsequent no-op patch replays are O(1), while independent equal
    /// opens converge on the same content identity.
    pub(crate) fn content_revision(&self) -> Result<[u8; 32]> {
        if let Some(revision) = self.content_revision.get() {
            return Ok(*revision);
        }
        let (revision, memo) =
            crate::opened::package_fingerprint_with_memo(&self.opc, self.part_digest_memo())?;
        if self.part_digests.get().is_none() {
            let _ = self.part_digests.set(std::sync::Arc::new(memo));
        }
        let _ = self.content_revision.set(revision);
        Ok(revision)
    }

    /// Return a previously established complete-content identity without
    /// starting a cold fingerprint pass.
    pub(crate) fn cached_content_revision(&self) -> Option<[u8; 32]> {
        self.content_revision.get().copied()
    }

    pub(crate) fn semantic_limits(&self) -> crate::open_options::SemanticLimits {
        crate::open_options::selected(self.semantic_limits.as_ref())
    }

    /// Every non-structural defect the reader repaired while opening this
    /// package, in discovery order.
    ///
    /// An empty report means the package conformed; see
    /// [`crate::ToleranceReport`]. Slide payloads parse lazily at accessor
    /// time, so repairs a lazy parse performs are reported by that parse's
    /// own entry point rather than here; see [`crate::leniency`].
    #[must_use]
    pub fn tolerance_report(&self) -> &crate::leniency::ToleranceReport {
        &self.tolerance
    }

    /// Read the package's inert custom document properties.
    ///
    /// A package without a custom-properties part returns an empty collection.
    /// The shared OOXML owner validates the complete package-level relationship
    /// graph and preserves its bounded value semantics. See MS-OI29500 3.11,
    /// "Reserved Custom File Properties".
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn custom_props(&self) -> Result<Props> {
        Ok(Props::read_for(&self.opc, Host::PowerPoint)?)
    }

    /// Read the complete inert Office Add-in task-pane graph, when present.
    ///
    /// The returned model covers MS-OWEXML task panes, `webextensionref`
    /// relationships, add-in references, bindings, properties, snapshots, and
    /// retained extension lists. It is data only: this crate never contacts
    /// external references or executes add-in content.
    ///
    /// # Errors
    ///
    /// Returns an error when managed authoring state is stale or the bounded
    /// package graph is malformed.
    pub fn task_panes(&self) -> Result<Option<crate::web::Panes>> {
        Ok(litchi_ooxml_common::web::load(
            &self.opc,
            &litchi_ooxml_common::web::Limits::standard(),
        )?)
    }

    /// Borrow the canonical presentation graph when no mutable state is stale.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn presentation(&self) -> Result<Presentation<'_>> {
        Ok(Presentation::new_with_semantic_limits(
            PresentationPart::from_package(&self.opc)?,
            &self.opc,
            self.semantic_limits.clone(),
        ))
    }

    /// Read the typed, source-backed InkAction snapshots owned by the
    /// presentation's slides in semantic slide order.
    ///
    /// # Errors
    ///
    /// Returns an error if the package graph or one typed action closure is
    /// malformed or outside the default bounds.
    pub fn ink_actions(&self) -> Result<Vec<crate::presentation::embedded::ink_actions::Snapshot>> {
        self.presentation()?.ink_actions()
    }

    /// Capture one immutable transaction root for a canonical presentation.
    ///
    /// The root composes existing slide ordering, checked shape text, and
    /// existing notes text without hydrating the legacy mutable writer.
    ///
    /// # Errors
    ///
    /// Returns an error for malformed package topology, unsupported notes
    /// dependencies, or bounds. Fresh and reopened packages use the same root.
    pub fn opened_presentation(&self) -> Result<crate::opened::Snapshot> {
        self.opened_presentation_with_limits(crate::opened::Limits::default())
    }

    /// Capture an opened-presentation transaction root under explicit edit
    /// options.
    ///
    /// # Errors
    ///
    /// Returns an error when the package is not a coherent canonical
    /// presentation under the options' finite bounds.
    pub fn opened_presentation_with_options(
        &self,
        options: crate::opened::EditOptions,
    ) -> Result<crate::opened::Snapshot> {
        self.opened_presentation_with_limits(options.limits_value())
    }

    /// Start one detached [`Edit`](crate::opened::Edit) under explicit edit
    /// options; see [`Self::edit`]. Capture of the opened presentation is
    /// deferred until the edit first needs its source state.
    ///
    /// # Errors
    ///
    /// The constructor itself does not parse the presentation. The first
    /// operation, source getter, or commit reports the same errors as
    /// [`Self::opened_presentation_with_options`].
    pub fn edit_with_options(
        &self,
        options: crate::opened::EditOptions,
    ) -> Result<crate::opened::Edit> {
        Ok(crate::opened::Edit::new_lazy(
            std::sync::Arc::clone(&self.opc),
            std::sync::Arc::clone(&self.replay_index),
            options.limits_value(),
            self.physical_source_provenance,
            std::sync::Arc::new(self.opened_owner_state()),
            std::sync::Arc::clone(&self.part_digests),
        ))
    }

    pub(crate) fn opened_presentation_with_limits(
        &self,
        limits: crate::opened::Limits,
    ) -> Result<crate::opened::Snapshot> {
        let mut snapshot = match self.part_digest_memo() {
            Some(parent) => crate::opened::capture_with_parent_digests_and_semantic_limits(
                &self.opc,
                limits,
                self.physical_source_provenance,
                parent,
                self.semantic_limits.clone(),
            )?,
            None => crate::opened::capture_with_provenance_and_semantic_limits(
                &self.opc,
                limits,
                self.physical_source_provenance,
                self.semantic_limits.clone(),
            )?,
        };
        snapshot.replay_index = Some(std::sync::Arc::clone(&self.replay_index));
        snapshot.owner_state = std::sync::Arc::new(self.opened_owner_state());
        self.offer_part_digests(&snapshot);
        Ok(snapshot)
    }

    /// Start one detached [`Edit`](crate::opened::Edit) over an immutable opened root.
    ///
    /// Dropping the edit discards its staged changes. The source package remains
    /// unchanged until a committed result is explicitly published.
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi_pptx::{Package, opened::Edit};
    /// # fn main() -> litchi_pptx::Result<()> {
    /// let mut builder = litchi_pptx::builder::Builder::new()?;
    /// builder.presentation()?.add_slide()?
    ///     .add_text_box("original", 0, 0, 914_400, 914_400);
    /// let bytes = builder.build()?.to_bytes()?;
    /// let mut package = Package::from_reader(std::io::Cursor::new(bytes.clone()))?;
    /// let mut edit: Edit = package.edit()?;
    /// edit.set_shape_text(0_usize, 0_usize, "staged")?;
    /// drop(edit);
    /// assert_eq!(package.to_bytes()?, bytes);
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// Capture is deferred until the edit first needs its source state. The
    /// first operation, source getter, or commit reports any malformed
    /// presentation or exceeded edit bound.
    pub fn edit(&self) -> Result<crate::opened::Edit> {
        self.edit_with_options(crate::opened::EditOptions::default())
    }

    /// Publish a slide-removal plan's durable patch, or a replayed one, to a
    /// new package.
    pub(crate) fn publish_slide_removal(
        &self,
        patch: &crate::slide_removal::Patch,
    ) -> Result<(Self, crate::opened::Snapshot)> {
        self.ensure_plain_mutation("remove_slide")?;
        let mut opc = (*self.opc).clone();
        let snapshot = crate::opened::apply_removal_patch_with_semantic_limits(
            &mut opc,
            patch,
            self.physical_source_provenance,
            self.semantic_limits.clone(),
        )?;
        let changed = patch
            .opened_patch()
            .resources()
            .cloned()
            .collect::<Vec<_>>();
        self.published(opc, snapshot, &changed)
    }

    fn require_physical_provenance(&self, source: &Self) -> Result<()> {
        if !source.physical_source_provenance || !self.physical_source_provenance {
            return Err(Error::slide_copy_refused(
                crate::SlideCopyRefusal::UnknownPhysicalMember,
            ));
        }
        source.ensure_plain_mutation("copy_slide")?;
        self.ensure_plain_mutation("copy_slide")
    }

    /// Publish a durable cross-presentation slide-copy patch from `source`
    /// to a new package.
    pub(crate) fn publish_cross_slide_copy_patch(
        &self,
        source: &Self,
        patch: &crate::cross_slide_copy::Patch,
    ) -> Result<(Self, crate::opened::Snapshot)> {
        self.require_physical_provenance(source)?;
        let mut opc = (*self.opc).clone();
        let snapshot = crate::opened::cross_copy_plan::apply_patch(
            &source.opc,
            &mut opc,
            patch,
            source.physical_source_provenance,
            self.physical_source_provenance,
            crate::opened::cross_copy_plan::LiveDigests {
                source: source.part_digest_memo(),
                destination: self.part_digest_memo(),
            },
            source.semantic_limits.clone(),
            self.semantic_limits.clone(),
        )?;
        let changed = patch
            .opened_patch()
            .resources()
            .cloned()
            .collect::<Vec<_>>();
        self.published(opc, snapshot, &changed)
    }

    /// Apply a durable complete-revision-bound slide-removal patch (or its
    /// inverse) to a new package; this one is unchanged.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when the complete package graph
    /// differs from the patch source, and an error when candidate validation
    /// fails or the encryption policy refuses typed mutation.
    pub fn apply_slide_removal(
        &self,
        patch: &crate::slide_removal::Patch,
    ) -> Result<crate::opened::Commit> {
        let (package, snapshot) = self.publish_slide_removal(patch)?;
        Ok(crate::opened::Commit::new(
            package,
            snapshot,
            patch.opened_patch().clone(),
        ))
    }

    /// Apply a durable source- and destination-bound cross-presentation
    /// slide-copy patch (or its inverse) to a new package; this one and
    /// `source` are unchanged.
    ///
    /// Both packages must retain source-preserving physical provenance. A
    /// forward patch that transfers source-compressed media
    /// ([`crate::cross_slide_copy::Patch::transfers_source_compressed_media`])
    /// is rebuilt from the two packages' bytes with the encoding it records,
    /// so any destination with the recorded revisions publishes the output of
    /// the first application; a destination holding a caller-defined `Part`
    /// implementation refuses it.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] for stale or foreign packages, and an
    /// error for malformed forged patch descriptors, rejected mutation
    /// policy, a transferring forward patch applied to a destination holding
    /// a caller-defined part ([`crate::SlideCopyRefusal::CallerDefinedPart`]),
    /// or candidate validation failure.
    pub fn apply_cross_slide_copy(
        &self,
        source: &Self,
        patch: &crate::cross_slide_copy::Patch,
    ) -> Result<crate::opened::Commit> {
        let (package, snapshot) = self.publish_cross_slide_copy_patch(source, patch)?;
        Ok(crate::opened::Commit::new(
            package,
            snapshot,
            patch.opened_patch().clone(),
        ))
    }

    /// The package publishing `opc`, with `snapshot`'s memo and this
    /// package's retained policy.
    fn published(
        &self,
        opc: OpcPackage,
        mut snapshot: crate::opened::Snapshot,
        changed: &[PackURI],
    ) -> Result<(Self, crate::opened::Snapshot)> {
        snapshot.replay_index = Some(std::sync::Arc::new(
            self.replay_index.update(&self.opc, &opc, changed)?,
        ));
        snapshot.owner_state = std::sync::Arc::new(self.opened_owner_state());
        let package = Self::from_opened_candidate(opc, &snapshot)?;
        Ok((package, snapshot))
    }

    /// Compare exact visible OPC content using the native revision proof.
    /// Pending mutable authoring state is refused; ZIP record encoding is not compared.
    pub fn same_content(&self, other: &Self) -> Result<bool> {
        let left = self.content_revision()?;
        let right = other.content_revision()?;
        Ok(left == right)
    }

    /// Publish an immutable opened package without flushing mutable authoring state.
    pub fn write_to(&self, writer: impl std::io::Write) -> Result<()> {
        #[cfg(feature = "encryption")]
        self.encryption
            .ordinary_output()
            .map_err(|source| Error::EncryptionPolicy {
                operation: "write_to",
                source,
            })?;
        litchi_opc::PackageWriter::write_to(writer, &self.opc)?;
        Ok(())
    }

    /// Apply a durable presentation patch to a new package, preserving this source.
    ///
    /// Patch replay validates its exact write set. Unrelated resources retain
    /// this package's state; a stale write-set resource refuses the whole apply.
    /// Encryption provenance is retained, so ordinary plaintext output stays
    /// refused until the caller explicitly selects an output policy.
    /// See [`opened::Commit`](crate::opened::Commit) for a replay/inverse example.
    ///
    /// # Errors
    ///
    /// Refuses pending authoring state, signed changes,
    /// stale resources, and invalid or unbounded candidate graphs.
    pub fn apply(&self, patch: &crate::opened::Patch) -> Result<crate::opened::Commit> {
        if !patch.is_noop() {
            super::signature::refuse_signed_change(&self.opc)?;
        }
        let no_memo = crate::opened::PartDigests::default();
        let parent = self
            .part_digests
            .get()
            .map_or(&no_memo, std::sync::Arc::as_ref);
        let (candidate, mut snapshot) = crate::opened::apply_to_new_with_semantic_limits(
            &self.opc,
            patch,
            self.physical_source_provenance,
            parent,
            self.semantic_limits.clone(),
        )?;
        if !patch.is_noop() {
            super::signature::refuse_signed_change(&candidate)?;
            candidate.validate_signature_edit_from(&self.opc)?;
        }
        let changed = patch.resources().cloned().collect::<Vec<_>>();
        snapshot.replay_index = Some(std::sync::Arc::new(
            self.replay_index.update(&self.opc, &candidate, &changed)?,
        ));
        snapshot.owner_state = std::sync::Arc::new(self.opened_owner_state());
        Ok(crate::opened::Commit::new(
            Self::from_opened_candidate(candidate, &snapshot)?,
            snapshot,
            patch.clone(),
        ))
    }

    fn opened_owner_state(&self) -> OpenedOwnerState {
        OpenedOwnerState {
            tolerance: self.tolerance.clone(),
            semantic_limits: self.semantic_limits.clone(),
            #[cfg(feature = "encryption")]
            encryption: self.encryption,
            #[cfg(feature = "encryption")]
            encryption_limits: self.encryption_limits,
        }
    }

    /// An equal package sharing this one's graph, payloads and memo: the
    /// package an exact no-op publishes.
    pub(crate) fn duplicate(&self) -> Self {
        let mut package = self.derive_with_part_bytes(self.opc.clone(), self.part_bytes);
        package.part_digests = self.part_digests.clone();
        package.content_revision = self.content_revision.clone();
        package
    }

    /// Derive a package when the caller has already accounted for the
    /// candidate's ordinary-part payload bytes in the same pass that built it.
    pub(crate) fn derive_with_part_bytes(
        &self,
        opc: impl Into<std::sync::Arc<OpcPackage>>,
        part_bytes: u64,
    ) -> Self {
        self.derive_with_part_bytes_and_replay_index(
            opc,
            part_bytes,
            std::sync::Arc::clone(&self.replay_index),
        )
    }

    pub(crate) fn derive_with_replay_index(
        &self,
        opc: impl Into<std::sync::Arc<OpcPackage>>,
        replay_index: std::sync::Arc<crate::package::replay_index::ReplayIndex>,
    ) -> Self {
        let opc = opc.into();
        let part_bytes = opc.admitted_part_bytes().unwrap_or(self.part_bytes);
        self.derive_with_part_bytes_and_replay_index(opc, part_bytes, replay_index)
    }

    pub(crate) fn derive_with_part_bytes_and_replay_index(
        &self,
        opc: impl Into<std::sync::Arc<OpcPackage>>,
        part_bytes: u64,
        replay_index: std::sync::Arc<crate::package::replay_index::ReplayIndex>,
    ) -> Self {
        Self {
            opc: opc.into(),
            part_bytes,
            content_revision: std::sync::Arc::new(std::sync::OnceLock::new()),
            physical_source_provenance: self.physical_source_provenance,
            semantic_limits: self.semantic_limits.clone(),
            tolerance: self.tolerance.clone(),
            part_digests: std::sync::Arc::new(std::sync::OnceLock::new()),
            replay_index,
            #[cfg(feature = "encryption")]
            encryption: self.encryption,
            #[cfg(feature = "encryption")]
            encryption_limits: self.encryption_limits,
        }
    }

    /// Refresh the retained physical and content metrics after native
    /// authoring has published its graph in place.  Authoring is the one
    /// package construction path that mutates the initial OPC owner rather
    /// than deriving a new facade, so both lazy identity cells must be
    /// replaced together with the byte aggregate.  The identity remains
    /// lazy: the publication pass already owns the graph construction, while
    /// an exact digest is only needed if a later raw no-op asks for it.
    pub(crate) fn refresh_authored_publication(&mut self) -> Result<()> {
        self.part_bytes = Self::part_bytes_for(&self.opc)?;
        self.replay_index = std::sync::Arc::new(
            crate::package::replay_index::ReplayIndex::from_package(&self.opc)?,
        );
        self.content_revision = std::sync::Arc::new(std::sync::OnceLock::new());
        self.part_digests = std::sync::Arc::new(std::sync::OnceLock::new());
        Ok(())
    }

    /// Drop the retained encryption provenance: the package publishes as an
    /// ordinary plaintext package from now on (the raw OPC family's
    /// declassification boundary).
    pub(crate) fn declassify(&mut self) {
        #[cfg(feature = "encryption")]
        {
            self.encryption = litchi_ooxml_common::package_encryption::PackageEncryption::plain();
        }
    }

    /// Publish already-validated native state without parsing or serialization.
    pub(crate) fn from_opened_candidate(
        opc: OpcPackage,
        snapshot: &crate::opened::Snapshot,
    ) -> Result<Self> {
        let state = &snapshot.owner_state;
        let replay_index = snapshot.replay_index.clone().map_or_else(
            || {
                crate::package::replay_index::ReplayIndex::from_package(&opc)
                    .map(std::sync::Arc::new)
            },
            Ok,
        )?;
        let mut package = Self {
            opc: std::sync::Arc::new(opc),
            part_bytes: snapshot.part_bytes,
            content_revision: {
                let slot = std::sync::OnceLock::new();
                let _ = slot.set(snapshot.revision);
                std::sync::Arc::new(slot)
            },
            physical_source_provenance: snapshot.physical_source_provenance,
            semantic_limits: state.semantic_limits.clone(),
            tolerance: state.tolerance.clone(),
            part_digests: std::sync::Arc::new(std::sync::OnceLock::new()),
            replay_index,
            #[cfg(feature = "encryption")]
            encryption: state.encryption,
            #[cfg(feature = "encryption")]
            encryption_limits: state.encryption_limits,
        };
        package.adopt_part_digests(snapshot);
        Ok(package)
    }

    /// Capture one slide's inert `[MS-PPTX]` change-tracking identifiers.
    ///
    /// The snapshot includes the slide creation ID and each shape modification
    /// ID in semantic scene order. It never interprets the identifiers as
    /// history, clocks, authorship, or executable collaboration state.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing, ambiguous, or out-of-range slide,
    /// malformed extension XML or exceeded bounds. Duplicate producer values
    /// remain readable for checked-position repair but cannot be committed.
    pub fn change_tracking<'a>(
        &self,
        slide: impl Into<crate::slide::Selector<'a>>,
    ) -> Result<crate::change_tracking::Snapshot> {
        let owner = self.resolve_slide(slide.into())?;
        crate::change_tracking::load_with_limits(&self.opc, &owner, self.semantic_limits().shape())
    }

    /// Load the complete inert speaker-notes graph, when present.
    ///
    /// Physical part names remain diagnostic details of the returned graph;
    /// ordinary callers do not need them to read or publish notes.
    ///
    /// # Errors
    ///
    /// Returns an error when mutable authoring state is stale or the notes
    /// graph is malformed or exceeds its resource limits.
    pub fn notes(&self) -> Result<Option<crate::notes::Graph>> {
        let presentation = PresentationPart::from_package(&self.opc)?;
        crate::notes::load_with_limits(
            &self.opc,
            presentation.part().partname(),
            self.semantic_limits()
                .notes()
                .constrained_by_source(self.opc.read_limits()),
        )
    }

    /// Capture the complete notes graph as a source-checked snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error when mutable authoring state is stale or the notes
    /// graph is malformed or exceeds its resource limits.
    pub fn notes_snapshot(&self) -> Result<Option<crate::notes::Snapshot>> {
        let presentation = PresentationPart::from_package(&self.opc)?;
        crate::notes::load_snapshot_with_limits(
            &self.opc,
            presentation.part().partname(),
            self.semantic_limits()
                .notes()
                .constrained_by_source(self.opc.read_limits()),
        )
    }

    /// Read the presentation-owned embedded-font collection.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn fonts(&self) -> Result<Option<crate::font::Fonts>> {
        crate::font::load(&self.opc)
    }

    /// Borrow the underlying clear OPC graph when it is current.
    ///
    /// For a package with retained encryption provenance, this is an explicit
    /// declassification boundary: the borrowed graph is plaintext. Merely
    /// borrowing it does not discard the retained output policy.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn opc(&self) -> Result<&OpcPackage> {
        Ok(&self.opc)
    }

    /// The presentation theme: the presentation part's own theme, else the
    /// first slide master's (record 0321). `None` when there is neither.
    ///
    /// ```
    /// let package = litchi_pptx::Package::new()?;
    /// let theme = package.theme()?.expect("a fresh deck has a theme");
    /// assert!(!theme.fonts.minor().latin.is_empty());
    /// # Ok::<(), litchi_pptx::Error>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns an error for a malformed presentation graph or theme part.
    pub fn theme(&self) -> Result<Option<litchi_drawingml::theme::Theme>> {
        let opc = self.opc()?;
        let Some(name) = crate::shape::theme::presentation_theme(opc)? else {
            return Ok(None);
        };
        crate::shape::theme::load(opc, name.as_str()).map(Some)
    }

    /// The number of the first slide, `p:presentation/@firstSlideNum`
    /// (1 when absent), the number slide-number fields count from
    /// (record 0279).
    ///
    /// # Errors
    ///
    /// Returns an error for a missing or malformed presentation part.
    pub fn first_slide_number(&self) -> Result<u32> {
        let part = self.opc()?.main_document_part()?;
        crate::presentation::first_slide_number(part.blob())
    }

    /// Run a read-only operation against the current clear OPC graph.
    ///
    /// For encrypted provenance this explicitly exposes plaintext OPC content;
    /// retained output policy remains active after the closure returns.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn with_opc<T>(&self, operation: impl FnOnce(&OpcPackage) -> Result<T>) -> Result<T> {
        operation(&self.opc)
    }

    /// Read one slide's direct, inert programmable-tag list.
    ///
    /// Names are the ordinary selector and zero-based presentation positions
    /// remain available for ordered repair workflows. The returned list owns
    /// its bounded strings, so the read does not borrow the package graph.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn tags<'a>(
        &self,
        slide: impl Into<crate::slide::Selector<'a>>,
    ) -> Result<Option<crate::tag::List>> {
        let slide_name = self.resolve_slide(slide.into())?;
        Ok(crate::tag::load_with_limits(
            &self.opc,
            &slide_name,
            self.semantic_limits()
                .tag()
                .constrained_by_source(self.opc.read_limits()),
        )?
        .map(crate::tag::Source::into_list))
    }

    /// Read the source-bound Designer tags owned by one stable slide ID.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn slide_designer_tags<'a>(
        &self,
        slide: impl Into<crate::slide::Selector<'a>>,
    ) -> Result<crate::presentation_properties::metadata::designer_tags::Snapshot> {
        let (_, slide_id) = self.resolve_slide_identity(slide.into())?;
        crate::presentation_properties::metadata::designer_tags::load_snapshot(&self.opc, slide_id)
    }

    /// Read the typed document-level math defaults from presentation properties.
    ///
    /// The presentation-properties repairs this read applies were recorded
    /// when the package was opened ([`Self::tolerance_report`]).
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn math_properties(
        &self,
    ) -> Result<Option<crate::presentation_properties::math::Properties>> {
        let limits =
            crate::presentation_properties::Limits::from_source_limits(self.opc.read_limits());
        crate::presentation_properties::load_math_from_package(&self.opc, limits)
    }

    /// Read the inert `p1710:readonlyRecommended` presentation hint.
    ///
    /// The recommendation is document data only. It does not make this
    /// package read-only and never changes host editing or save policy.
    ///
    /// # Errors
    ///
    /// Returns an error if the presentation-properties owner is malformed or
    /// ambiguous.
    pub fn read_only_recommended(&self) -> Result<Option<bool>> {
        crate::presentation_properties::read_only_recommended::load(&self.opc)
    }

    /// Capture the exact presentation-properties source of the hint.
    ///
    /// The snapshot is absent when the package has no presentation-properties
    /// owner. An existing owner still yields a snapshot when the
    /// recommendation itself is absent.
    ///
    /// # Errors
    ///
    /// Returns an error if the owner is malformed or ambiguous.
    pub fn read_only_recommended_snapshot(
        &self,
    ) -> Result<Option<crate::presentation_properties::read_only_recommended::Snapshot>> {
        crate::presentation_properties::read_only_recommended::load_snapshot(&self.opc)
    }

    /// Read the lossless zoom owner of one slide.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn zooms<'a>(
        &self,
        slide: impl Into<crate::slide::Selector<'a>>,
    ) -> Result<crate::shape::zoom::Owner> {
        let slide_name = self.resolve_slide(slide.into())?;
        crate::shape::zoom::load(&self.opc, &slide_name)
    }

    /// Read one semantic shape's optional programmable-tag list.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn shape_tags<'s, 'k>(
        &self,
        slide: impl Into<crate::slide::Selector<'s>>,
        shape: impl Into<crate::shape::Key<'k>>,
    ) -> Result<Option<crate::tag::List>> {
        let slide_name = self.resolve_slide(slide.into())?;
        Ok(crate::tag::shape::load(
            &self.opc,
            &slide_name,
            shape,
            self.semantic_limits()
                .tag()
                .constrained_by_source(self.opc.read_limits()),
            self.semantic_limits().shape(),
        )?
        .map(crate::tag::Source::into_list))
    }

    /// Read the optional classification outcome attached to one semantic
    /// shape. The shape selector uses its exact producer name by default and
    /// retains checked source-order indices for repair workflows.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn shape_classification<'s, 'k>(
        &self,
        slide: impl Into<crate::slide::Selector<'s>>,
        shape: impl Into<crate::shape::Key<'k>>,
    ) -> Result<Option<crate::shape::classification::Snapshot>> {
        let slide_name = self.resolve_slide(slide.into())?;
        crate::shape::classification::load_with_limits(
            &self.opc,
            &slide_name,
            shape,
            self.semantic_limits().shape(),
        )
    }

    /// Read the optional `p15:designElem` value attached to one semantic
    /// shape. The selector is name-first, with checked source-order indices
    /// retained for repair workflows.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn shape_design_element<'s, 'k>(
        &self,
        slide: impl Into<crate::slide::Selector<'s>>,
        shape: impl Into<crate::shape::Key<'k>>,
    ) -> Result<Option<crate::shape::designer::Snapshot>> {
        let slide_name = self.resolve_slide(slide.into())?;
        crate::shape::designer::load_with_limits(
            &self.opc,
            &slide_name,
            shape,
            self.semantic_limits().shape(),
        )
    }

    /// Read one shape's source-bound `PowerPoint` 2020 Designer properties.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn shape_designer_properties<'s, 'k>(
        &self,
        slide: impl Into<crate::slide::Selector<'s>>,
        shape: impl Into<crate::shape::Key<'k>>,
    ) -> Result<crate::shape::designer::properties::Snapshot> {
        let (slide_name, _) = self.resolve_slide_identity(slide.into())?;
        crate::shape::designer::load_properties_with_shape_limits(
            &self.opc,
            &slide_name,
            shape,
            self.semantic_limits().shape(),
        )
    }

    /// Read all contextual 3D-model owners on one slide in source order.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn model3ds<'a>(
        &self,
        slide: impl Into<crate::slide::Selector<'a>>,
    ) -> Result<Vec<crate::model3d::Model>> {
        let slide_name = self.resolve_slide(slide.into())?;
        let limits = crate::model3d::codec::Limits::from_source_limits(self.opc.read_limits());
        crate::model3d::package::load_all(&self.opc, &slide_name, limits)
    }

    /// Read the model3d owner attached to one semantic shape, if present.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn model3d<'s, 'k>(
        &self,
        slide: impl Into<crate::slide::Selector<'s>>,
        shape: impl Into<crate::shape::Key<'k>>,
    ) -> Result<Option<crate::model3d::Model>> {
        let slide_name = self.resolve_slide(slide.into())?;
        let limits = crate::model3d::codec::Limits::from_source_limits(self.opc.read_limits());
        crate::model3d::package::load(&self.opc, &slide_name, shape.into(), limits)
    }

    /// Load the presentation's optional `DrawingML` table-style catalog.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn styles(&self) -> Result<Option<crate::table::style::List>> {
        // Table styles are owned by their OPC part, independently of the
        // slide-authoring model. They remain a safe immutable read while a
        // newly authored slide is still pending publication.
        crate::table::style::load(&self.opc)
    }

    /// Validate every master/layout relationship reachable from the package.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn validate_master_layout_graph(&self) -> Result<()> {
        self.with_opc(crate::master_layout::validate_master_layout_graph)
    }

    /// Load all contextual slide-library synchronization metadata.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn load_slide_sync(
        &self,
    ) -> Result<Vec<crate::presentation_properties::metadata::slide_sync::Part>> {
        let limits = self.semantic_limits().slide_sync();
        self.with_opc(|opc| {
            crate::presentation_properties::metadata::slide_sync::load(
                opc,
                limits.constrained_by_source(opc.read_limits()),
            )
        })
    }

    /// A new package with every signature removed explicitly (ADR 0006);
    /// this package is unchanged.
    ///
    /// Every signature relationship and signature part is removed, and the
    /// resulting unsigned graph is authorized for publication
    /// ([`OpcPackage::unsign`]). A changed edit refuses a signed package,
    /// because the change invalidates every signature: a changed
    /// opened-presentation commit ([`crate::opened::Edit::commit`]),
    /// the package-level edit families of this crate (their methods on this
    /// facade and their functions on an [`OpcPackage`]), and a changed
    /// [`Self::edit_opc`]. Call this first to edit one. A signed package that
    /// is saved unchanged, or edited by an exact no-op, keeps its signature
    /// byte for byte. For a package that carries no signature, and no
    /// signature-policy requirement, the result is an equal package.
    ///
    /// ```
    /// use litchi_pptx::Package;
    ///
    /// let package = Package::from_reader(std::io::Cursor::new(Package::new()?.to_bytes()?))?;
    /// let unsigned = package.unsign()?;
    /// assert!(!unsigned.opc()?.is_signed());
    /// # Ok::<(), litchi_pptx::Error>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns the encryption-policy error for a package opened with a
    /// password, whose typed mutations are refused.
    pub fn unsign(&self) -> Result<Self> {
        if !(self.opc.is_signed() || self.opc.requires_signature_edit_policy()) {
            return Ok(self.duplicate());
        }
        self.ensure_plain_mutation("unsign")?;
        let mut opc = (*self.opc).clone();
        opc.unsign();
        let part_bytes = Self::part_bytes_for(&opc)?;
        Ok(self.derive_with_part_bytes(opc, part_bytes))
    }

    /// Adopt, as this package's memo, `snapshot`'s payload-digest memo
    /// re-projected onto the allocations `self.opc` itself holds.
    ///
    /// The snapshot was captured from the graph now in `self.opc`, but its
    /// memo names the allocations of the snapshot's own package, a clone. A
    /// clone of a built-in part shares its payload allocation; a caller-defined
    /// part may copy its payload when cloned, and adopting the snapshot's memo
    /// as it is would then keep that copy alive after the snapshot is dropped.
    /// The projection keeps an entry only for an allocation `self.opc` holds,
    /// and retains `self.opc`'s own `Arc` for it (change 0751's review), so the
    /// facade retains no payload its own graph does not hold. A projection that
    /// cannot allocate leaves the memo empty, which only costs a later hash.
    fn adopt_part_digests(&mut self, snapshot: &crate::opened::Snapshot) {
        self.part_digests = std::sync::Arc::new(std::sync::OnceLock::new());
        let _ = self.content_revision.set(snapshot.revision);
        if let Some(memo) = Self::memo_for(&self.opc, snapshot) {
            let _kept = self.part_digests.set(memo);
        }
    }

    /// Keep the memo of a capture of the current graph, re-projected onto
    /// `self.opc`'s own allocations, when this package holds none (change
    /// 0751).
    ///
    /// The ADR 0005 memo amendment's adoption clause says when the facade must
    /// refresh its memo (every publication) and when it must release it (every
    /// mutation that publishes nothing); filling an empty slot from a snapshot
    /// of the current, unmutated graph is permitted under the coordinator's
    /// ruling in 0751's review, with the amendment's re-projection clause met
    /// in code: the memo kept is [`Self::memo_for`]'s projection, whatever the
    /// parts are. A filled slot is kept: it describes this same graph, because
    /// every mutation that publishes no snapshot empties it and every
    /// publication replaces it.
    fn offer_part_digests(&self, snapshot: &crate::opened::Snapshot) {
        let _ = self.content_revision.set(snapshot.revision);
        if self.part_digests.get().is_none()
            && let Some(memo) = Self::memo_for(&self.opc, snapshot)
        {
            let _kept = self.part_digests.set(memo);
        }
    }

    /// `snapshot`'s memo re-projected onto `opc`'s own payload allocations:
    /// every entry kept names, and retains, an allocation `opc` holds, and an
    /// entry naming any other allocation is dropped. `None` when the
    /// projection cannot allocate.
    fn memo_for(
        opc: &OpcPackage,
        snapshot: &crate::opened::Snapshot,
    ) -> Option<std::sync::Arc<crate::opened::PartDigests>> {
        snapshot.package_memo(opc).ok().map(std::sync::Arc::new)
    }

    /// The payload-digest memo this package retains for its current graph.
    pub(crate) fn part_digest_memo(&self) -> Option<&crate::opened::PartDigests> {
        self.part_digests.get().map(std::sync::Arc::as_ref)
    }

    pub(crate) fn ensure_plain_mutation(&self, operation: &'static str) -> Result<()> {
        let _ = operation;
        #[cfg(feature = "encryption")]
        self.encryption
            .ordinary_output()
            .map_err(|source| Error::EncryptionPolicy { operation, source })?;
        Ok(())
    }

    fn resolve_slide_identity(&self, key: crate::slide::Selector<'_>) -> Result<(PackURI, u32)> {
        resolve_slide_identity(&self.opc, key)
    }

    fn resolve_slide(&self, key: crate::slide::Selector<'_>) -> Result<PackURI> {
        resolve_slide(&self.opc, key)
    }
}

pub(crate) fn resolve_slide_identity(
    opc: &OpcPackage,
    key: crate::slide::Selector<'_>,
) -> Result<(PackURI, u32)> {
    let presentation = Presentation::new(PresentationPart::from_package(opc)?, opc);
    let references = presentation.slide_references()?;
    match key {
        crate::slide::Selector::Position(position) => {
            let index = position.get();
            let reference = references.get(index).ok_or(Error::SlideIndexOutOfBounds {
                index,
                len: references.len(),
            })?;
            let slide = presentation
                .slide(index)?
                .ok_or(Error::SlideIndexOutOfBounds {
                    index,
                    len: references.len(),
                })?;
            Ok((slide.part().part().partname().clone(), reference.id()))
        },
        crate::slide::Selector::Name(name) => {
            let slides = presentation.slides()?;
            let mut selected = None;
            let mut matches = 0usize;
            for (reference, slide) in references.iter().zip(slides) {
                if slide.name()? == name {
                    matches = matches.saturating_add(1);
                    selected = Some((slide.part().part().partname().clone(), reference.id()));
                }
            }
            match (matches, selected) {
                (0, _) => Err(Error::SlideNameNotFound(name.into_owned())),
                (1, Some(value)) => Ok(value),
                _ => Err(Error::AmbiguousSlideName {
                    name: name.into_owned(),
                    matches,
                }),
            }
        },
    }
}

/// The zero-based position and part of the selected slide (record 0356a).
pub(crate) fn resolve_slide_position(
    opc: &OpcPackage,
    key: crate::slide::Selector<'_>,
) -> Result<(usize, PackURI)> {
    if let crate::slide::Selector::Position(position) = &key {
        let index = position.get();
        return Ok((index, resolve_slide(opc, key)?));
    }
    let name = resolve_slide(opc, key)?;
    let presentation = Presentation::new(PresentationPart::from_package(opc)?, opc);
    for index in 0..presentation.slide_count()? {
        if let Some(slide) = presentation.slide(index)?
            && *slide.part().part().partname() == name
        {
            return Ok((index, name));
        }
    }
    Err(Error::internal(
        "resolve_slide_position",
        "a named slide is not in the presentation order",
    ))
}

pub(crate) fn resolve_slide(opc: &OpcPackage, key: crate::slide::Selector<'_>) -> Result<PackURI> {
    let presentation = Presentation::new(PresentationPart::from_package(opc)?, opc);
    match key {
        crate::slide::Selector::Position(position) => {
            let index = position.get();
            let length = presentation.slide_count()?;
            let slide = presentation
                .slide(index)?
                .ok_or(Error::SlideIndexOutOfBounds { index, len: length })?;
            Ok(slide.part().part().partname().clone())
        },
        crate::slide::Selector::Name(name) => {
            let slide = presentation
                .find_slide(name.as_ref())?
                .ok_or_else(|| Error::SlideNameNotFound(name.into_owned()))?;
            Ok(slide.part().part().partname().clone())
        },
    }
}

#[cfg(test)]
impl Package {
    /// Test fixture seam: run one raw OPC operation through [`Self::edit_opc`]
    /// and replace this package with the committed one.
    pub(crate) fn edit_opc_in_place<T>(
        &mut self,
        operation: impl FnOnce(&mut OpcPackage) -> Result<T>,
    ) -> Result<T> {
        let mut edit = self.edit_opc()?;
        let value = edit.update(operation)?;
        let package = edit.commit()?.into_package();
        *self = package;
        Ok(value)
    }
}

#[cfg(test)]
impl Package {
    /// Test fixture seam: stage and commit a same-package slide-copy plan,
    /// replacing this package.
    pub(crate) fn apply_slide_copy_plan(
        &mut self,
        plan: &crate::slide_copy::Plan,
    ) -> Result<crate::opened::Snapshot> {
        let mut edit = self.edit()?;
        edit.stage_slide_copy(plan)?;
        self.adopt_commit(edit.commit()?)
    }

    /// Test fixture seam: stage and commit a slide-removal plan, replacing
    /// this package.
    pub(crate) fn apply_slide_removal_plan(
        &mut self,
        plan: &crate::slide_removal::Plan,
    ) -> Result<crate::opened::Snapshot> {
        let mut edit = self.edit()?;
        edit.stage_slide_removal(plan)?;
        self.adopt_commit(edit.commit()?)
    }

    /// Test fixture seam: replay a durable slide-removal patch, replacing
    /// this package.
    pub(crate) fn apply_slide_removal_patch(
        &mut self,
        patch: &crate::slide_removal::Patch,
    ) -> Result<crate::opened::Snapshot> {
        let commit = self.apply_slide_removal(patch)?;
        self.adopt_commit(commit)
    }

    /// Test fixture seam: stage and commit a cross-presentation slide-copy
    /// plan, replacing this package.
    pub(crate) fn apply_cross_slide_copy_plan(
        &mut self,
        source: &Self,
        plan: &crate::cross_slide_copy::Plan,
    ) -> Result<crate::opened::Snapshot> {
        let mut edit = self.edit()?;
        edit.stage_cross_slide_copy(source, plan)?;
        self.adopt_commit(edit.commit()?)
    }

    /// Test fixture seam: replay a durable cross-presentation slide-copy
    /// patch, replacing this package.
    pub(crate) fn apply_cross_slide_copy_patch(
        &mut self,
        source: &Self,
        patch: &crate::cross_slide_copy::Patch,
    ) -> Result<crate::opened::Snapshot> {
        let commit = self.apply_cross_slide_copy(source, patch)?;
        self.adopt_commit(commit)
    }

    fn adopt_commit(&mut self, commit: crate::opened::Commit) -> Result<crate::opened::Snapshot> {
        let snapshot = commit.snapshot().clone();
        *self = commit.into_package();
        Ok(snapshot)
    }
}
