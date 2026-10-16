//! Typed `PresentationML` package facade and semantic package operations.

use litchi_opc::OpcPackage;
use litchi_opc::packuri::PackURI;

use crate::custom::{Host, Props};
use crate::parts::PresentationPart;
use crate::presentation::Presentation;
use crate::writer::MutablePresentation;
use crate::{Error, Result};

/// Main entry point for `PresentationML` package ownership.
pub struct Package {
    pub(crate) opc: OpcPackage,
    pub(crate) mutable_pres: Option<MutablePresentation>,
    /// Whether ingress retained the exact source archive needed for physical
    /// revision authorization of topology-changing cross-package edits.
    pub(crate) physical_source_provenance: bool,
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
    pub(crate) part_digests: std::sync::OnceLock<std::sync::Arc<crate::opened::PartDigests>>,
    #[cfg(feature = "encryption")]
    pub(crate) encryption: litchi_ooxml_common::package_encryption::PackageEncryption,
    #[cfg(feature = "automatic-fonts")]
    pub(crate) font_embedding_dirty: bool,
}

impl std::fmt::Debug for Package {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Package")
            .field("part_count", &self.opc.part_count())
            .field("pending_authoring", &self.mutable_pres.is_some())
            .field("physical_source_retained", &self.physical_source_provenance)
            .finish_non_exhaustive()
    }
}

/// Package policy that must survive snapshot-derived edits without reopening.
#[derive(Clone, Default)]
pub(crate) struct OpenedOwnerState {
    tolerance: crate::leniency::ToleranceReport,
    #[cfg(feature = "encryption")]
    encryption: litchi_ooxml_common::package_encryption::PackageEncryption,
    #[cfg(feature = "automatic-fonts")]
    font_embedding_dirty: bool,
}

impl Package {
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

    /// Replace the package's inert custom document properties transactionally.
    ///
    /// Empty properties remove the package-level relationship and target part.
    /// The shared OOXML owner validates and stages the graph before publication.
    /// Properties that serialize to the stored bytes publish nothing, so a
    /// signed package keeps its signature. See MS-OI29500 3.11, "Reserved
    /// Custom File Properties".
    ///
    /// A change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006) and leaves the package unchanged: remove the signature
    /// explicitly first with [`Self::unsign`].
    ///
    /// # Errors
    ///
    /// Returns [`Error::Opc`] for a change to a signed package, and an error
    /// if the operation fails.
    pub fn put_custom_props(&mut self, props: Props) -> Result<()> {
        self.edit_typed(move |opc| {
            props
                .write_for(opc, Host::PowerPoint)
                .map_err(crate::package::signature::from_shared_custom_props)
        })
    }

    /// Remove every inert custom document property transactionally.
    ///
    /// This is idempotent. See MS-OI29500 3.11, "Reserved Custom File
    /// Properties". Removing absent properties keeps a signed package's
    /// signature; removing present ones from a signed package is refused as
    /// in [`Self::put_custom_props`].
    ///
    /// # Errors
    ///
    /// Returns [`Error::Opc`] for a change to a signed package, and an error
    /// if the operation fails.
    pub fn remove_custom_props(&mut self) -> Result<()> {
        self.put_custom_props(Props::new())
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
        self.ensure_graph_current("task_panes")?;
        Ok(litchi_ooxml_common::web::load(&self.opc)?)
    }

    /// Plan an exact, reversible replacement for the inert task-pane graph.
    ///
    /// The patch records the precise source graph and must be published with
    /// [`Self::apply_task_panes_patch`]. An unchanged replacement yields an
    /// empty patch and consequently preserves digital signatures.
    ///
    /// # Errors
    ///
    /// Returns an error when managed authoring state is stale or the graph or
    /// candidate violates bounded MS-OWEXML/OPC invariants.
    pub fn task_panes_patch(
        &self,
        panes: crate::web::Panes,
        conformance: crate::web::Conformance,
    ) -> Result<crate::web::Patch> {
        self.ensure_graph_current("task_panes_patch")?;
        Ok(litchi_ooxml_common::web::plan_put(
            &self.opc,
            panes,
            conformance,
        )?)
    }

    /// Plan removal of the inert Office Add-in task-pane graph.
    ///
    /// An absent graph yields an empty patch. Referenced resources are removed
    /// only when OPC ownership checks prove they are not shared.
    ///
    /// # Errors
    ///
    /// Returns an error when managed authoring state is stale or the bounded
    /// package graph is malformed.
    pub fn clear_task_panes_patch(&self) -> Result<crate::web::Patch> {
        self.ensure_graph_current("clear_task_panes_patch")?;
        Ok(litchi_ooxml_common::web::plan_remove(&self.opc)?)
    }

    /// Apply an exact source-checked Office Add-in task-pane patch atomically.
    ///
    /// A stale patch is rejected before publication. Its [`crate::web::Patch::inverse`]
    /// can restore the original graph when no intervening change occurred.
    ///
    /// # Errors
    ///
    /// Returns an error when managed authoring state is stale, package policy
    /// rejects mutation, or the patch's source no longer matches the package.
    pub fn apply_task_panes_patch(&mut self, patch: &crate::web::Patch) -> Result<bool> {
        self.ensure_graph_current("apply_task_panes_patch")?;
        self.edit_typed(|opc| Ok(patch.apply(opc)?))
    }

    /// Borrow the canonical presentation graph when no mutable state is stale.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn presentation(&self) -> Result<Presentation<'_>> {
        self.ensure_graph_current("presentation")?;
        Ok(Presentation::new(
            PresentationPart::from_package(&self.opc)?,
            &self.opc,
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

    /// Read InkAction snapshots under explicit owner and target bounds.
    ///
    /// # Errors
    ///
    /// Returns an error if one selected slide exceeds the supplied bounds or
    /// has an unsupported or malformed typed action closure.
    pub fn ink_actions_with_limits(
        &self,
        limits: crate::presentation::embedded::ink_actions::Limits,
    ) -> Result<Vec<crate::presentation::embedded::ink_actions::Snapshot>> {
        self.presentation()?.ink_actions_with_limits(limits)
    }

    /// Apply an existing-target InkAction patch with exact source checks.
    ///
    /// # Errors
    ///
    /// Returns an error for stale source, signature policy, invalid closure,
    /// or bounded readback failure. The package remains unchanged on error.
    pub fn apply_ink_actions_patch(
        &mut self,
        patch: &crate::presentation::embedded::ink_actions::Patch,
    ) -> Result<crate::presentation::embedded::ink_actions::Snapshot> {
        self.ensure_graph_current("apply_ink_actions_patch")?;
        self.ensure_plain_mutation("apply_ink_actions_patch")?;
        let changed = !patch.is_noop();
        // The InkAction package owner stages and publishes its own candidate
        // atomically.  Avoid wrapping that operation in `edit_typed`, whose
        // generic rollback snapshot would duplicate the complete OPC graph.
        let snapshot = patch.apply(&mut self.opc)?;
        if changed {
            self.mutable_pres = None;
        }
        Ok(snapshot)
    }

    /// Borrow the mutable presentation model for a newly authored package.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub(crate) fn authoring_presentation(&mut self) -> Result<&mut MutablePresentation> {
        self.ensure_plain_mutation("authoring_presentation")?;
        self.mutable_pres.as_mut().ok_or(Error::UnsafeEdit {
            operation: "authoring_presentation",
            reason: "the lossless facade cannot hydrate an opened package into the mutable writer",
        })
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

    /// Capture an opened-presentation transaction root with explicit bounds.
    ///
    /// # Errors
    ///
    /// Returns an error when the package is not a coherent canonical
    /// presentation under the supplied finite policy.
    pub fn opened_presentation_with_limits(
        &self,
        limits: crate::opened::Limits,
    ) -> Result<crate::opened::Snapshot> {
        self.ensure_graph_current("opened_presentation")?;
        if self.mutable_pres.is_some() {
            return Err(Error::UnsafeEdit {
                operation: "opened_presentation",
                reason: "finish the builder before starting a package transaction",
            });
        }
        let mut snapshot = match self.part_digest_memo() {
            Some(parent) => crate::opened::capture_with_parent_digests(
                &self.opc,
                limits,
                self.physical_source_provenance,
                parent,
            )?,
            None => crate::opened::capture_with_provenance(
                &self.opc,
                limits,
                self.physical_source_provenance,
            )?,
        };
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
    /// # Errors
    ///
    /// Returns an error under the same conditions as [`Self::opened_presentation`].
    pub fn edit(&self) -> Result<crate::opened::Edit> {
        Ok(self.opened_presentation()?.edit())
    }

    /// Atomically publish a source-bound same-package whole-slide copy plan.
    ///
    /// Planning validates and materializes the complete candidate without
    /// mutation. Publication refuses when any part of the package graph has
    /// changed since planning, including resources outside the copy write set.
    /// Unrelated OPC members retain their exact in-memory part state and are
    /// left outside the finite resource patch.
    ///
    /// No standalone inverse is exposed: a generic durable patch does not
    /// encode the plan's complete-package revision binding.
    ///
    /// # Errors
    ///
    /// Returns an error for a stale plan, package mutation-policy refusal, or a
    /// candidate graph that no longer applies exactly.
    pub fn apply_slide_copy_plan(
        &mut self,
        plan: &crate::opened::SlideCopyPlan,
    ) -> Result<crate::opened::Snapshot> {
        self.ensure_graph_current("apply_slide_copy_plan")?;
        if self.mutable_pres.is_some() {
            return Err(Error::UnsafeEdit {
                operation: "apply_slide_copy_plan",
                reason: "slide-copy plans require an already-opened package",
            });
        }
        self.ensure_plain_mutation("apply_slide_copy_plan")?;
        if crate::opened::package_fingerprint(&self.opc)? != plan.source_revision() {
            return Err(Error::UnsafeEdit {
                operation: "apply_slide_copy_plan",
                reason: "the complete package graph changed after slide-copy planning",
            });
        }
        let mut snapshot =
            crate::opened::apply(&mut self.opc, plan.patch(), self.physical_source_provenance)?;
        snapshot.owner_state = std::sync::Arc::new(self.opened_owner_state());
        self.adopt_part_digests(&snapshot);
        self.mutable_pres = None;
        Ok(snapshot)
    }

    /// Atomically publish a source-bound cross-presentation slide-copy plan.
    ///
    /// The source package is read-only and must retain the complete graph
    /// revision captured during planning.  The destination is cloned and
    /// fully validated before one publication assignment; stale or foreign
    /// source packages are rejected before mutation.  The source layout is
    /// never copied: the plan has already proved its inheritance equivalent to
    /// the selected destination layout.
    /// Source and destination must have been opened through source-preserving
    /// ingress (`open` or `from_reader`) so discarded ZIP
    /// ordering and extra fields cannot alias the authorization revision.
    ///
    /// A plan whose candidate transfers source-compressed media
    /// ([`crate::opened::CrossSlideCopyPlan::transfers_source_compressed_media`])
    /// is rebuilt from the two packages' bytes, so it publishes the same
    /// output into any destination with the recorded revisions. Into a
    /// destination that is not an unmodified owned source it publishes the
    /// candidate reopened from its archive, which carries the destination's
    /// save preferences; a destination holding a caller-defined `Part`
    /// implementation refuses it rather than lose that part's behavior.
    ///
    /// # Errors
    ///
    /// Returns an error for stale source/destination graphs, a foreign source,
    /// package mutation policy, a transferring plan applied to a destination
    /// holding a caller-defined part
    /// ([`crate::SlideCopyRefusal::CallerDefinedPart`]), or a candidate graph
    /// that no longer applies exactly.
    pub fn apply_cross_slide_copy_plan(
        &mut self,
        source: &Self,
        plan: &crate::opened::CrossSlideCopyPlan,
    ) -> Result<crate::opened::Snapshot> {
        self.ensure_graph_current("apply_cross_slide_copy_plan")?;
        source.ensure_graph_current("apply_cross_slide_copy_plan")?;
        if self.mutable_pres.is_some() {
            return Err(Error::UnsafeEdit {
                operation: "apply_cross_slide_copy_plan",
                reason: "cross-slide plans require an already-opened destination package",
            });
        }
        if source.mutable_pres.is_some() {
            return Err(Error::UnsafeEdit {
                operation: "apply_cross_slide_copy_plan",
                reason: "cross-slide plans require an already-opened source package",
            });
        }
        if !source.physical_source_provenance || !self.physical_source_provenance {
            return Err(Error::SlideCopyPlan {
                kind: crate::SlideCopyRefusal::UnknownPhysicalMember,
                detail: "cross-slide physical authorization requires source-preserving package ingress (use Package::open, from_reader, or their *_with_options forms)".to_owned(),
            });
        }
        source.ensure_plain_mutation("apply_cross_slide_copy_plan")?;
        self.ensure_plain_mutation("apply_cross_slide_copy_plan")?;
        let mut snapshot = crate::opened::cross_copy_plan::apply_plan(
            &source.opc,
            &mut self.opc,
            plan,
            source.physical_source_provenance,
            self.physical_source_provenance,
            crate::opened::cross_copy_plan::LiveDigests {
                source: source.part_digest_memo(),
                // A field borrow, disjoint from `self.opc`.
                destination: self.part_digests.get().map(std::sync::Arc::as_ref),
            },
        )?;
        snapshot.owner_state = std::sync::Arc::new(self.opened_owner_state());
        self.adopt_part_digests(&snapshot);
        self.mutable_pres = None;
        Ok(snapshot)
    }

    /// Atomically publish a durable source- and destination-bound
    /// cross-presentation slide-copy patch.
    ///
    /// The source package remains immutable.  The patch's inverse can restore
    /// the destination when the same source revision and the corresponding
    /// destination target revision are still present.
    /// Both packages must retain source-preserving physical provenance.
    /// A forward patch that transfers source-compressed media
    /// ([`crate::opened::CrossSlideCopyPatch::transfers_source_compressed_media`])
    /// is rebuilt from the two packages' bytes with the encoding it records,
    /// so a redo after its inverse, or any destination with the recorded
    /// revisions, publishes the output of the first application; a
    /// destination holding a caller-defined `Part` implementation refuses it.
    ///
    /// # Errors
    ///
    /// Returns an error for stale or foreign packages, malformed forged patch
    /// descriptors, rejected mutation policy, a transferring forward patch
    /// applied to a destination holding a caller-defined part
    /// ([`crate::SlideCopyRefusal::CallerDefinedPart`]), or candidate
    /// validation failure.
    pub fn apply_cross_slide_copy_patch(
        &mut self,
        source: &Self,
        patch: &crate::opened::CrossSlideCopyPatch,
    ) -> Result<crate::opened::Snapshot> {
        self.ensure_graph_current("apply_cross_slide_copy_patch")?;
        source.ensure_graph_current("apply_cross_slide_copy_patch")?;
        if self.mutable_pres.is_some() {
            return Err(Error::UnsafeEdit {
                operation: "apply_cross_slide_copy_patch",
                reason: "cross-slide patches require an already-opened destination package",
            });
        }
        if source.mutable_pres.is_some() {
            return Err(Error::UnsafeEdit {
                operation: "apply_cross_slide_copy_patch",
                reason: "cross-slide patches require an already-opened source package",
            });
        }
        if !source.physical_source_provenance || !self.physical_source_provenance {
            return Err(Error::SlideCopyPlan {
                kind: crate::SlideCopyRefusal::UnknownPhysicalMember,
                detail: "cross-slide physical authorization requires source-preserving package ingress (use Package::open, from_reader, or their *_with_options forms)".to_owned(),
            });
        }
        source.ensure_plain_mutation("apply_cross_slide_copy_patch")?;
        self.ensure_plain_mutation("apply_cross_slide_copy_patch")?;
        let mut snapshot = crate::opened::cross_copy_plan::apply_patch(
            &source.opc,
            &mut self.opc,
            patch,
            source.physical_source_provenance,
            self.physical_source_provenance,
            crate::opened::cross_copy_plan::LiveDigests {
                source: source.part_digest_memo(),
                // A field borrow, disjoint from `self.opc`.
                destination: self.part_digests.get().map(std::sync::Arc::as_ref),
            },
        )?;
        snapshot.owner_state = std::sync::Arc::new(self.opened_owner_state());
        self.adopt_part_digests(&snapshot);
        self.mutable_pres = None;
        Ok(snapshot)
    }

    /// Atomically publish a source-bound whole-slide removal.
    ///
    /// Planning captures a complete-package fingerprint and an exact durable
    /// patch over only the presentation owner and selected slide. Publication
    /// refuses if any package resource changed after planning. The plan's
    /// [`crate::opened::SlideRemovalPlan::patch`] can be serialized and inverted
    /// through [`Self::apply_slide_removal_patch`].
    ///
    /// # Errors
    ///
    /// Returns an error for a stale plan, package mutation-policy refusal, or
    /// a candidate graph that no longer applies exactly.
    pub fn apply_slide_removal_plan(
        &mut self,
        plan: &crate::opened::SlideRemovalPlan,
    ) -> Result<crate::opened::Snapshot> {
        self.ensure_graph_current("apply_slide_removal_plan")?;
        if self.mutable_pres.is_some() {
            return Err(Error::UnsafeEdit {
                operation: "apply_slide_removal_plan",
                reason: "slide-removal plans require an already-opened package",
            });
        }
        self.ensure_plain_mutation("apply_slide_removal_plan")?;
        let mut snapshot = crate::opened::apply_removal_patch(
            &mut self.opc,
            plan.patch(),
            "apply_slide_removal_plan",
            self.physical_source_provenance,
        )?;
        // Like every other publication, adopt the published snapshot's memo,
        // so the memo of the graph this removal replaced (which a capture may
        // have offered, change 0751) is not kept.
        snapshot.owner_state = std::sync::Arc::new(self.opened_owner_state());
        self.adopt_part_digests(&snapshot);
        self.mutable_pres = None;
        Ok(snapshot)
    }

    /// Atomically publish a durable complete-revision-bound slide removal patch.
    ///
    /// # Errors
    ///
    /// Returns an error before mutation when the complete package graph differs
    /// from the patch source, or when candidate validation fails.
    pub fn apply_slide_removal_patch(
        &mut self,
        patch: &crate::opened::SlideRemovalPatch,
    ) -> Result<crate::opened::Snapshot> {
        self.ensure_graph_current("apply_slide_removal_patch")?;
        if self.mutable_pres.is_some() {
            return Err(Error::UnsafeEdit {
                operation: "apply_slide_removal_patch",
                reason: "slide-removal patches require an already-opened package",
            });
        }
        self.ensure_plain_mutation("apply_slide_removal_patch")?;
        let mut snapshot = crate::opened::apply_removal_patch(
            &mut self.opc,
            patch,
            "apply_slide_removal_patch",
            self.physical_source_provenance,
        )?;
        snapshot.owner_state = std::sync::Arc::new(self.opened_owner_state());
        self.adopt_part_digests(&snapshot);
        self.mutable_pres = None;
        Ok(snapshot)
    }

    /// Compare exact visible OPC content using the native revision proof.
    /// Pending mutable authoring state is refused; ZIP record encoding is not compared.
    pub fn same_content(&self, other: &Self) -> Result<bool> {
        self.ensure_graph_current("same_content")?;
        other.ensure_graph_current("same_content")?;
        if self.mutable_pres.is_some() || other.mutable_pres.is_some() {
            return Err(Error::UnsafeEdit {
                operation: "same_content",
                reason: "comparison requires already-opened packages",
            });
        }
        let left =
            crate::opened::package_fingerprint_with_memo(&self.opc, self.part_digest_memo())?.0;
        let right =
            crate::opened::package_fingerprint_with_memo(&other.opc, other.part_digest_memo())?.0;
        Ok(left == right)
    }

    /// Publish an immutable opened package without flushing mutable authoring state.
    pub fn write_to(&self, writer: impl std::io::Write) -> Result<()> {
        self.ensure_graph_current("write_to")?;
        if self.mutable_pres.is_some() {
            return Err(Error::UnsafeEdit {
                operation: "write_to",
                reason: "publication requires an already-opened package",
            });
        }
        #[cfg(feature = "encryption")]
        self.encryption
            .ordinary_output()
            .map_err(|source| Error::EncryptionPolicy {
                operation: "write_to",
                source,
            })?;
        litchi_opc::PackageWriter::write_to_stream(writer, &self.opc)?;
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
    pub fn apply(&self, patch: &crate::opened::Patch) -> Result<Self> {
        self.ensure_graph_current("apply")?;
        if self.mutable_pres.is_some() {
            return Err(Error::UnsafeEdit {
                operation: "apply",
                reason: "presentation patches require an already-opened package",
            });
        }
        if !patch.is_noop() {
            super::signature::refuse_signed_change(&self.opc)?;
        }
        let no_memo = crate::opened::PartDigests::default();
        let parent = self
            .part_digests
            .get()
            .map_or(&no_memo, std::sync::Arc::as_ref);
        let (candidate, mut snapshot) =
            crate::opened::apply_to_new(&self.opc, patch, self.physical_source_provenance, parent)?;
        if !patch.is_noop() {
            super::signature::refuse_signed_change(&candidate)?;
            candidate.validate_signature_edit_from(&self.opc)?;
        }
        snapshot.owner_state = std::sync::Arc::new(self.opened_owner_state());
        Ok(Self::from_opened_candidate(candidate, &snapshot))
    }

    fn opened_owner_state(&self) -> OpenedOwnerState {
        OpenedOwnerState {
            tolerance: self.tolerance.clone(),
            #[cfg(feature = "encryption")]
            encryption: self.encryption,
            #[cfg(feature = "automatic-fonts")]
            font_embedding_dirty: self.font_embedding_dirty,
        }
    }

    /// Publish already-validated native state without parsing or serialization.
    pub(crate) fn from_opened_candidate(
        opc: OpcPackage,
        snapshot: &crate::opened::Snapshot,
    ) -> Self {
        let state = &snapshot.owner_state;
        let mut package = Self {
            opc,
            mutable_pres: None,
            physical_source_provenance: snapshot.physical_source_provenance,
            tolerance: state.tolerance.clone(),
            part_digests: std::sync::OnceLock::new(),
            #[cfg(feature = "encryption")]
            encryption: state.encryption,
            #[cfg(feature = "automatic-fonts")]
            font_embedding_dirty: state.font_embedding_dirty,
        };
        package.adopt_part_digests(snapshot);
        package
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
        self.ensure_graph_current("change_tracking")?;
        let owner = self.resolve_slide(slide.into())?;
        crate::change_tracking::load(&self.opc, &owner)
    }

    /// Publish a committed change-tracking edit atomically.
    ///
    /// A change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006) and leaves the package unchanged: remove the signature
    /// explicitly first with [`Self::unsign`]. An exact no-op keeps the
    /// signature.
    ///
    /// # Errors
    ///
    /// Returns an error when the selected slide no longer exactly matches the
    /// commit source or package mutation policy refuses publication.
    pub fn apply_change_tracking_commit(
        &mut self,
        commit: crate::change_tracking::Commit,
    ) -> Result<crate::change_tracking::Snapshot> {
        self.ensure_graph_current("apply_change_tracking_commit")?;
        self.ensure_plain_mutation("apply_change_tracking_commit")?;
        let changed = !commit.patch().is_noop();
        let snapshot = crate::change_tracking::apply_commit(&mut self.opc, commit)?;
        self.release_part_digests();
        if changed {
            self.mutable_pres = None;
        }
        Ok(snapshot)
    }

    /// Apply an exact-source reversible change-tracking patch atomically.
    ///
    /// A change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006) and leaves the package unchanged: remove the signature
    /// explicitly first with [`Self::unsign`]. An exact no-op keeps the
    /// signature.
    ///
    /// # Errors
    ///
    /// Returns an error when the patch is stale, malformed, or rejected by
    /// package mutation policy.
    pub fn apply_change_tracking_patch(
        &mut self,
        patch: &crate::change_tracking::Patch,
    ) -> Result<crate::change_tracking::Snapshot> {
        self.ensure_graph_current("apply_change_tracking_patch")?;
        self.ensure_plain_mutation("apply_change_tracking_patch")?;
        let changed = !patch.is_noop();
        let snapshot = crate::change_tracking::apply_patch(&mut self.opc, patch)?;
        self.release_part_digests();
        if changed {
            self.mutable_pres = None;
        }
        Ok(snapshot)
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
        self.ensure_graph_current("notes")?;
        let presentation = PresentationPart::from_package(&self.opc)?;
        crate::notes::load(&self.opc, presentation.part().partname())
    }

    /// Capture the complete notes graph as a source-checked snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error when mutable authoring state is stale or the notes
    /// graph is malformed or exceeds its resource limits.
    pub fn notes_snapshot(&self) -> Result<Option<crate::notes::Snapshot>> {
        self.ensure_graph_current("notes_snapshot")?;
        let presentation = PresentationPart::from_package(&self.opc)?;
        crate::notes::load_snapshot(&self.opc, presentation.part().partname())
    }

    /// Replace an existing coherent notes graph atomically.
    ///
    /// Exact replacements are byte-stable no-ops. Opaque XML, inert links,
    /// and related payloads retained by the graph are never interpreted or
    /// activated.
    ///
    /// A change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006) and leaves the package unchanged: remove the signature
    /// explicitly first with [`Self::unsign`]. An exact no-op keeps the
    /// signature.
    ///
    /// # Errors
    ///
    /// Returns an error when the source graph is stale, malformed, encrypted
    /// against ordinary mutation, or differs in physical ownership.
    pub fn put_notes(&mut self, graph: crate::notes::Graph) -> Result<()> {
        self.ensure_graph_current("put_notes")?;
        let source = self
            .notes_snapshot()?
            .ok_or_else(|| Error::Invalid("notes graph is absent".into()))?;
        let mut edit = source.edit();
        edit.replace(graph)?;
        let commit = edit.commit()?;
        self.apply_notes_commit(commit).map(|_| ())
    }

    /// Publish a committed source-checked notes edit atomically.
    ///
    /// A change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006) and leaves the package unchanged: remove the signature
    /// explicitly first with [`Self::unsign`]. An exact no-op keeps the
    /// signature.
    ///
    /// # Errors
    ///
    /// Returns an error when the commit source is stale, the graph is
    /// malformed, or package policy rejects mutation.
    pub fn apply_notes_commit(
        &mut self,
        commit: crate::notes::Commit,
    ) -> Result<crate::notes::Snapshot> {
        self.ensure_graph_current("apply_notes_commit")?;
        let changed = commit.diagnostics().changed();
        let snapshot = self.edit_typed(move |opc| crate::notes::apply_commit(opc, commit))?;
        if changed {
            self.mutable_pres = None;
        }
        Ok(snapshot)
    }

    /// Publish a reversible source-checked notes patch atomically.
    ///
    /// A change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006) and leaves the package unchanged: remove the signature
    /// explicitly first with [`Self::unsign`]. An exact no-op keeps the
    /// signature.
    ///
    /// # Errors
    ///
    /// Returns an error when the patch source is stale, the graph is
    /// malformed, or package policy rejects mutation.
    pub fn apply_notes_patch(
        &mut self,
        patch: &crate::notes::Patch,
    ) -> Result<crate::notes::Snapshot> {
        self.ensure_graph_current("apply_notes_patch")?;
        let changed = !patch.is_noop();
        let snapshot = self.edit_typed(|opc| crate::notes::apply_patch(opc, patch))?;
        if changed {
            self.mutable_pres = None;
        }
        Ok(snapshot)
    }

    /// Remove one slide's notes by checked position or exact visible name.
    ///
    /// Slide-owned classic and modern comments are retained. Descendants of
    /// the removed Notes Slide are collected only when package-wide inbound
    /// edge validation proves exclusive ownership.
    ///
    /// A change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006) and leaves the package unchanged: remove the signature
    /// explicitly first with [`Self::unsign`]. An exact no-op keeps the
    /// signature.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing, ambiguous, or out-of-range selector,
    /// malformed ownership, shared descendants, or rejected package policy.
    pub fn remove_notes<'a>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'a>>,
    ) -> Result<bool> {
        self.ensure_graph_current("remove_notes")?;
        let slide_name = self.resolve_slide(slide.into())?;
        let Some(source) = self.notes_snapshot()? else {
            return Ok(false);
        };
        let removed =
            self.edit_typed(move |opc| crate::notes::remove_checked(opc, &source, &slide_name))?;
        if removed {
            self.mutable_pres = None;
        }
        Ok(removed)
    }

    /// Remove notes from every slide while retaining shared notes resources.
    ///
    /// A change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006) and leaves the package unchanged: remove the signature
    /// explicitly first with [`Self::unsign`]. An exact no-op keeps the
    /// signature.
    ///
    /// # Errors
    ///
    /// Returns an error when the graph is malformed, descendants have shared
    /// ownership, or package policy rejects mutation.
    pub fn clear_notes(&mut self) -> Result<usize> {
        self.ensure_graph_current("clear_notes")?;
        let Some(source) = self.notes_snapshot()? else {
            return Ok(0);
        };
        let removed = self.edit_typed(move |opc| crate::notes::clear_checked(opc, &source))?;
        if removed != 0 {
            self.mutable_pres = None;
        }
        Ok(removed)
    }

    /// Read the presentation-owned embedded-font collection.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn fonts(&self) -> Result<Option<crate::font::Fonts>> {
        self.ensure_graph_current("fonts")?;
        crate::font::load(&self.opc)
    }

    /// Replace the complete presentation-owned embedded-font collection.
    ///
    /// This is the explicit, inert font-resource API. It remains independent
    /// from the optional system-font discovery policy used at managed save.
    ///
    /// A change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006) and leaves the package unchanged: remove the signature
    /// explicitly first with [`Self::unsign`]. An exact no-op keeps the
    /// signature.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn put_fonts(&mut self, fonts: crate::font::Fonts) -> Result<bool> {
        self.edit_typed(move |opc| crate::font::put(opc, fonts))
    }

    /// Remove all presentation-owned embedded fonts and orphaned resources.
    ///
    /// A change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006) and leaves the package unchanged: remove the signature
    /// explicitly first with [`Self::unsign`]. An exact no-op keeps the
    /// signature.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn remove_fonts(&mut self) -> Result<Option<crate::font::Fonts>> {
        self.edit_typed(crate::font::remove)
    }

    /// Whether a mutable model is currently pending managed publication.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn is_modified(&self) -> bool {
        self.mutable_pres
            .as_ref()
            .is_some_and(MutablePresentation::is_modified)
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
        self.ensure_graph_current("opc")?;
        Ok(&self.opc)
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
        self.ensure_graph_current("with_opc")?;
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
        self.ensure_graph_current("tags")?;
        let slide_name = self.resolve_slide(slide.into())?;
        Ok(crate::tag::load(&self.opc, &slide_name)?.map(crate::tag::Source::into_list))
    }

    /// Create or replace one slide's direct programmable-tag list.
    ///
    /// The list is moved into the package transaction and the staged owner
    /// relationship and part are published together.
    ///
    /// A change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006) and leaves the package unchanged: remove the signature
    /// explicitly first with [`Self::unsign`]. An exact no-op keeps the
    /// signature.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn put_tags<'a>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'a>>,
        list: crate::tag::List,
    ) -> Result<Option<crate::tag::List>> {
        self.ensure_graph_current("put_tags")?;
        let slide_name = self.resolve_slide(slide.into())?;
        self.edit_typed(move |opc| crate::tag::put(opc, &slide_name, list))
    }

    /// Remove one slide's direct programmable-tag list.
    ///
    /// Removal is idempotent and only collects an orphaned tag part after the
    /// package-wide inbound-edge check succeeds.
    ///
    /// A change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006) and leaves the package unchanged: remove the signature
    /// explicitly first with [`Self::unsign`]. An exact no-op keeps the
    /// signature.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn remove_tags<'a>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'a>>,
    ) -> Result<Option<crate::tag::List>> {
        self.ensure_graph_current("remove_tags")?;
        let slide_name = self.resolve_slide(slide.into())?;
        self.edit_typed(move |opc| crate::tag::remove(opc, &slide_name))
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
        self.slide_designer_tags_with_limits(slide, crate::shape::designer::Limits::default())
    }

    /// Read slide-ID Designer tags under caller-supplied resource bounds.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn slide_designer_tags_with_limits<'a>(
        &self,
        slide: impl Into<crate::slide::Selector<'a>>,
        limits: crate::shape::designer::Limits,
    ) -> Result<crate::presentation_properties::metadata::designer_tags::Snapshot> {
        self.ensure_graph_current("slide_designer_tags")?;
        let (_, slide_id) = self.resolve_slide_identity(slide.into())?;
        crate::presentation_properties::metadata::designer_tags::load_snapshot_with_limits(
            &self.opc, slide_id, limits,
        )
    }

    /// Create or replace Designer tags on one stable slide ID atomically.
    ///
    /// A change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006) and leaves the package unchanged: remove the signature
    /// explicitly first with [`Self::unsign`]. An exact no-op keeps the
    /// signature.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn put_slide_designer_tags<'a>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'a>>,
        tags: crate::shape::designer::Tags,
    ) -> Result<crate::presentation_properties::metadata::designer_tags::Snapshot> {
        self.put_slide_designer_tags_with_limits(
            slide,
            tags,
            crate::shape::designer::Limits::default(),
        )
    }

    /// Create or replace slide-ID Designer tags under explicit bounds.
    ///
    /// A change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006) and leaves the package unchanged: remove the signature
    /// explicitly first with [`Self::unsign`]. An exact no-op keeps the
    /// signature.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn put_slide_designer_tags_with_limits<'a>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'a>>,
        tags: crate::shape::designer::Tags,
        limits: crate::shape::designer::Limits,
    ) -> Result<crate::presentation_properties::metadata::designer_tags::Snapshot> {
        self.ensure_graph_current("put_slide_designer_tags")?;
        let (_, slide_id) = self.resolve_slide_identity(slide.into())?;
        let mut edit =
            crate::presentation_properties::metadata::designer_tags::load_snapshot_with_limits(
                &self.opc, slide_id, limits,
            )?
            .edit()?;
        edit.set(tags)?;
        let commit = edit.commit()?;
        let changed = commit.diagnostics().changed();
        self.publish_designer(changed, move |opc| commit.apply(opc))
    }

    /// Remove Designer tags from one stable slide ID atomically.
    ///
    /// A change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006) and leaves the package unchanged: remove the signature
    /// explicitly first with [`Self::unsign`]. An exact no-op keeps the
    /// signature.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn remove_slide_designer_tags<'a>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'a>>,
    ) -> Result<crate::presentation_properties::metadata::designer_tags::Snapshot> {
        self.remove_slide_designer_tags_with_limits(
            slide,
            crate::shape::designer::Limits::default(),
        )
    }

    /// Remove slide-ID Designer tags under explicit resource bounds.
    ///
    /// A change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006) and leaves the package unchanged: remove the signature
    /// explicitly first with [`Self::unsign`]. An exact no-op keeps the
    /// signature.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn remove_slide_designer_tags_with_limits<'a>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'a>>,
        limits: crate::shape::designer::Limits,
    ) -> Result<crate::presentation_properties::metadata::designer_tags::Snapshot> {
        self.ensure_graph_current("remove_slide_designer_tags")?;
        let (_, slide_id) = self.resolve_slide_identity(slide.into())?;
        let mut edit =
            crate::presentation_properties::metadata::designer_tags::load_snapshot_with_limits(
                &self.opc, slide_id, limits,
            )?
            .edit()?;
        edit.remove();
        let commit = edit.commit()?;
        let changed = commit.diagnostics().changed();
        self.publish_designer(changed, move |opc| commit.apply(opc))
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
        self.ensure_graph_current("math_properties")?;
        crate::presentation_properties::load_math_from_package(&self.opc)
    }

    /// Replace the document-level math defaults transactionally.
    ///
    /// A change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006) and leaves the package unchanged: remove the signature
    /// explicitly first with [`Self::unsign`]. An exact no-op keeps the
    /// signature.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn put_math_properties(
        &mut self,
        value: crate::presentation_properties::math::Properties,
    ) -> Result<Option<crate::presentation_properties::math::Properties>> {
        self.ensure_graph_current("put_math_properties")?;
        self.edit_typed(move |opc| crate::presentation_properties::put_math_to_package(opc, value))
    }

    /// Remove the document-level math defaults transactionally.
    ///
    /// A change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006) and leaves the package unchanged: remove the signature
    /// explicitly first with [`Self::unsign`]. An exact no-op keeps the
    /// signature.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn remove_math_properties(
        &mut self,
    ) -> Result<Option<crate::presentation_properties::math::Properties>> {
        self.ensure_graph_current("remove_math_properties")?;
        self.edit_typed(crate::presentation_properties::remove_math_from_package)
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
    pub fn readonly_recommended(&self) -> Result<Option<bool>> {
        self.ensure_graph_current("readonly_recommended")?;
        crate::presentation_properties::readonly_recommended::load(&self.opc)
    }

    /// Alias using the conventional Rust spelling.
    pub fn read_only_recommended(&self) -> Result<Option<bool>> {
        self.readonly_recommended()
    }

    /// Capture the exact presentation-properties source for a checked edit.
    ///
    /// The snapshot is absent when the package has no presentation-properties
    /// owner. An existing owner still yields a snapshot when the recommendation
    /// itself is absent, allowing an insertion to preserve the owner's source.
    ///
    /// # Errors
    ///
    /// Returns an error if the owner is malformed or ambiguous.
    pub fn readonly_recommended_snapshot(
        &self,
    ) -> Result<Option<crate::presentation_properties::readonly_recommended::Snapshot>> {
        self.ensure_graph_current("readonly_recommended_snapshot")?;
        crate::presentation_properties::readonly_recommended::load_snapshot(&self.opc)
    }

    /// Alias using the conventional Rust spelling.
    pub fn read_only_recommended_snapshot(
        &self,
    ) -> Result<Option<crate::presentation_properties::readonly_recommended::Snapshot>> {
        self.readonly_recommended_snapshot()
    }

    /// Set the inert `p1710:readonlyRecommended` hint transactionally.
    ///
    /// An exact semantic no-op preserves the package bytes and any retained
    /// signature. A changed signed source requires an explicit package
    /// signature policy, as do all other typed package edits.
    ///
    /// # Errors
    ///
    /// Returns an error if the owner is malformed, ambiguous, stale, or the
    /// package signature policy rejects the mutation.
    pub fn put_readonly_recommended(&mut self, value: bool) -> Result<Option<bool>> {
        self.ensure_graph_current("put_readonly_recommended")?;
        let previous = self.edit_typed(move |opc| {
            crate::presentation_properties::readonly_recommended::put(opc, value)
        })?;
        if previous != Some(value) {
            self.mutable_pres = None;
        }
        Ok(previous)
    }

    /// Alias using the conventional Rust spelling.
    pub fn put_read_only_recommended(&mut self, value: bool) -> Result<Option<bool>> {
        self.put_readonly_recommended(value)
    }

    /// Remove the inert recommendation while retaining its owner part.
    ///
    /// # Errors
    ///
    /// Returns an error if the owner is malformed, ambiguous, stale, or the
    /// package signature policy rejects the mutation.
    pub fn remove_readonly_recommended(&mut self) -> Result<Option<bool>> {
        self.ensure_graph_current("remove_readonly_recommended")?;
        let previous =
            self.edit_typed(crate::presentation_properties::readonly_recommended::remove)?;
        if previous.is_some() {
            self.mutable_pres = None;
        }
        Ok(previous)
    }

    /// Alias using the conventional Rust spelling.
    pub fn remove_read_only_recommended(&mut self) -> Result<Option<bool>> {
        self.remove_readonly_recommended()
    }

    /// Publish a committed source-preserving recommendation edit.
    ///
    /// # Errors
    ///
    /// Returns an error if the source is stale, malformed, or the package
    /// signature policy rejects the mutation.
    pub fn apply_readonly_recommended_commit(
        &mut self,
        commit: crate::presentation_properties::readonly_recommended::Commit,
    ) -> Result<crate::presentation_properties::readonly_recommended::Snapshot> {
        self.ensure_graph_current("apply_readonly_recommended_commit")?;
        let changed = commit.diagnostics().changed();
        let snapshot = self.edit_typed(move |opc| {
            crate::presentation_properties::readonly_recommended::apply_commit(opc, commit)
        })?;
        if changed {
            self.mutable_pres = None;
        }
        Ok(snapshot)
    }

    /// Alias using the conventional Rust spelling.
    pub fn apply_read_only_recommended_commit(
        &mut self,
        commit: crate::presentation_properties::readonly_recommended::Commit,
    ) -> Result<crate::presentation_properties::readonly_recommended::Snapshot> {
        self.apply_readonly_recommended_commit(commit)
    }

    /// Publish a reversible source-checked recommendation patch.
    ///
    /// # Errors
    ///
    /// Returns an error if the source is stale, malformed, or the package
    /// signature policy rejects the mutation.
    pub fn apply_readonly_recommended_patch(
        &mut self,
        patch: &crate::presentation_properties::readonly_recommended::Patch,
    ) -> Result<crate::presentation_properties::readonly_recommended::Snapshot> {
        self.ensure_graph_current("apply_readonly_recommended_patch")?;
        let changed = !patch.is_noop();
        let snapshot = self.edit_typed(|opc| {
            crate::presentation_properties::readonly_recommended::apply_patch(opc, patch)
        })?;
        if changed {
            self.mutable_pres = None;
        }
        Ok(snapshot)
    }

    /// Alias using the conventional Rust spelling.
    pub fn apply_read_only_recommended_patch(
        &mut self,
        patch: &crate::presentation_properties::readonly_recommended::Patch,
    ) -> Result<crate::presentation_properties::readonly_recommended::Snapshot> {
        self.apply_readonly_recommended_patch(patch)
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
        self.ensure_graph_current("zooms")?;
        let slide_name = self.resolve_slide(slide.into())?;
        crate::shape::zoom::load(&self.opc, &slide_name)
    }

    /// Replace one slide's zoom owner transactionally.
    ///
    /// A change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006) and leaves the package unchanged: remove the signature
    /// explicitly first with [`Self::unsign`]. An exact no-op keeps the
    /// signature.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn put_zooms<'a>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'a>>,
        owner: crate::shape::zoom::Owner,
    ) -> Result<Option<crate::shape::zoom::Owner>> {
        self.ensure_graph_current("put_zooms")?;
        let slide_name = self.resolve_slide(slide.into())?;
        self.edit_typed(move |opc| crate::shape::zoom::store(opc, &slide_name, owner))
    }

    /// Remove all zoom metadata from one slide while retaining its other XML.
    ///
    /// A change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006) and leaves the package unchanged: remove the signature
    /// explicitly first with [`Self::unsign`]. An exact no-op keeps the
    /// signature.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn remove_zooms<'a>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'a>>,
    ) -> Result<Option<crate::shape::zoom::Owner>> {
        self.ensure_graph_current("remove_zooms")?;
        let slide_name = self.resolve_slide(slide.into())?;
        self.edit_typed(move |opc| crate::shape::zoom::remove(opc, &slide_name))
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
        self.ensure_graph_current("shape_tags")?;
        let slide_name = self.resolve_slide(slide.into())?;
        Ok(crate::tag::shape::load(&self.opc, &slide_name, shape)?
            .map(crate::tag::Source::into_list))
    }

    /// Create or replace one semantic shape's programmable-tag list.
    ///
    /// A change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006) and leaves the package unchanged: remove the signature
    /// explicitly first with [`Self::unsign`]. An exact no-op keeps the
    /// signature.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn put_shape_tags<'s, 'k>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        shape: impl Into<crate::shape::Key<'k>>,
        list: crate::tag::List,
    ) -> Result<Option<crate::tag::List>> {
        self.ensure_graph_current("put_shape_tags")?;
        let slide_name = self.resolve_slide(slide.into())?;
        self.edit_typed(move |opc| crate::tag::shape::put(opc, &slide_name, shape, list))
    }

    /// Remove one semantic shape's programmable-tag list.
    ///
    /// A change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006) and leaves the package unchanged: remove the signature
    /// explicitly first with [`Self::unsign`]. An exact no-op keeps the
    /// signature.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn remove_shape_tags<'s, 'k>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        shape: impl Into<crate::shape::Key<'k>>,
    ) -> Result<Option<crate::tag::List>> {
        self.ensure_graph_current("remove_shape_tags")?;
        let slide_name = self.resolve_slide(slide.into())?;
        self.edit_typed(move |opc| crate::tag::shape::remove(opc, &slide_name, shape))
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
        self.ensure_graph_current("shape_classification")?;
        let slide_name = self.resolve_slide(slide.into())?;
        crate::shape::classification::load(&self.opc, &slide_name, shape)
    }

    /// Create or replace one semantic shape's typed classification outcome
    /// transactionally while retaining unrelated extension markup.
    ///
    /// A change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006) and leaves the package unchanged: remove the signature
    /// explicitly first with [`Self::unsign`]. An exact no-op keeps the
    /// signature.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn put_shape_classification<'s, 'k>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        shape: impl Into<crate::shape::Key<'k>>,
        outcome: crate::shape::classification::Outcome,
    ) -> Result<Option<crate::shape::classification::Snapshot>> {
        self.ensure_graph_current("put_shape_classification")?;
        let slide_name = self.resolve_slide(slide.into())?;
        self.edit_typed(move |opc| {
            crate::shape::classification::put(opc, &slide_name, shape, outcome)
        })
    }

    /// Remove one semantic shape's typed classification element
    /// transactionally. Unknown extension entries remain intact.
    ///
    /// A change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006) and leaves the package unchanged: remove the signature
    /// explicitly first with [`Self::unsign`]. An exact no-op keeps the
    /// signature.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn remove_shape_classification<'s, 'k>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        shape: impl Into<crate::shape::Key<'k>>,
    ) -> Result<Option<crate::shape::classification::Snapshot>> {
        self.ensure_graph_current("remove_shape_classification")?;
        let slide_name = self.resolve_slide(slide.into())?;
        self.edit_typed(move |opc| crate::shape::classification::remove(opc, &slide_name, shape))
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
        self.ensure_graph_current("shape_design_element")?;
        let slide_name = self.resolve_slide(slide.into())?;
        crate::shape::designer::load(&self.opc, &slide_name, shape)
    }

    /// Create or replace one shape's typed `p15:designElem` boolean
    /// transactionally while retaining unrelated extension entries.
    ///
    /// A change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006) and leaves the package unchanged: remove the signature
    /// explicitly first with [`Self::unsign`]. An exact no-op keeps the
    /// signature.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn put_shape_design_element<'s, 'k>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        shape: impl Into<crate::shape::Key<'k>>,
        value: bool,
    ) -> Result<Option<crate::shape::designer::Snapshot>> {
        self.ensure_graph_current("put_shape_design_element")?;
        let slide_name = self.resolve_slide(slide.into())?;
        self.edit_typed(move |opc| crate::shape::designer::put(opc, &slide_name, shape, value))
    }

    /// Remove one shape's typed `p15:designElem` value transactionally.
    ///
    /// A change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006) and leaves the package unchanged: remove the signature
    /// explicitly first with [`Self::unsign`]. An exact no-op keeps the
    /// signature.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn remove_shape_design_element<'s, 'k>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        shape: impl Into<crate::shape::Key<'k>>,
    ) -> Result<Option<crate::shape::designer::Snapshot>> {
        self.ensure_graph_current("remove_shape_design_element")?;
        let slide_name = self.resolve_slide(slide.into())?;
        self.edit_typed(move |opc| crate::shape::designer::remove(opc, &slide_name, shape))
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
    ) -> Result<crate::shape::designer::PropertiesSnapshot> {
        self.shape_designer_properties_with_limits(
            slide,
            shape,
            crate::shape::designer::Limits::default(),
        )
    }

    /// Read shape Designer properties under caller-supplied resource bounds.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn shape_designer_properties_with_limits<'s, 'k>(
        &self,
        slide: impl Into<crate::slide::Selector<'s>>,
        shape: impl Into<crate::shape::Key<'k>>,
        limits: crate::shape::designer::Limits,
    ) -> Result<crate::shape::designer::PropertiesSnapshot> {
        self.ensure_graph_current("shape_designer_properties")?;
        let (slide_name, _) = self.resolve_slide_identity(slide.into())?;
        crate::shape::designer::load_properties_with_limits(&self.opc, &slide_name, shape, limits)
    }

    /// Create or replace one shape's Designer drawing properties atomically.
    ///
    /// A change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006) and leaves the package unchanged: remove the signature
    /// explicitly first with [`Self::unsign`]. An exact no-op keeps the
    /// signature.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn put_shape_designer_properties<'s, 'k>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        shape: impl Into<crate::shape::Key<'k>>,
        properties: crate::shape::designer::DrawingProperties,
    ) -> Result<crate::shape::designer::PropertiesSnapshot> {
        self.ensure_graph_current("put_shape_designer_properties")?;
        let (slide_name, _) = self.resolve_slide_identity(slide.into())?;
        let shape = shape.into();
        let before = crate::shape::designer::load_properties(&self.opc, &slide_name, shape)?;
        let changed = before.properties() != Some(&properties);
        self.publish_designer(changed, move |opc| {
            crate::shape::designer::put_properties(opc, &slide_name, shape, properties)
        })
    }

    /// Create or replace shape Designer properties under explicit bounds.
    ///
    /// A change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006) and leaves the package unchanged: remove the signature
    /// explicitly first with [`Self::unsign`]. An exact no-op keeps the
    /// signature.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn put_shape_designer_properties_with_limits<'s, 'k>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        shape: impl Into<crate::shape::Key<'k>>,
        properties: crate::shape::designer::DrawingProperties,
        limits: crate::shape::designer::Limits,
    ) -> Result<crate::shape::designer::PropertiesSnapshot> {
        self.ensure_graph_current("put_shape_designer_properties")?;
        let (slide_name, _) = self.resolve_slide_identity(slide.into())?;
        let mut edit = crate::shape::designer::load_properties_with_limits(
            &self.opc,
            &slide_name,
            shape,
            limits,
        )?
        .edit();
        edit.set(properties)?;
        let commit = edit.commit()?;
        let changed = !commit.is_noop();
        self.publish_designer(changed, move |opc| {
            crate::shape::designer::apply_properties(opc, commit)
        })
    }

    /// Remove one shape's Designer drawing properties atomically.
    ///
    /// A change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006) and leaves the package unchanged: remove the signature
    /// explicitly first with [`Self::unsign`]. An exact no-op keeps the
    /// signature.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn remove_shape_designer_properties<'s, 'k>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        shape: impl Into<crate::shape::Key<'k>>,
    ) -> Result<crate::shape::designer::PropertiesSnapshot> {
        self.ensure_graph_current("remove_shape_designer_properties")?;
        let (slide_name, _) = self.resolve_slide_identity(slide.into())?;
        let shape = shape.into();
        let before = crate::shape::designer::load_properties(&self.opc, &slide_name, shape)?;
        let changed = before.is_present();
        self.publish_designer(changed, move |opc| {
            crate::shape::designer::remove_properties(opc, &slide_name, shape)
        })
    }

    /// Remove shape Designer properties under explicit resource bounds.
    ///
    /// A change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006) and leaves the package unchanged: remove the signature
    /// explicitly first with [`Self::unsign`]. An exact no-op keeps the
    /// signature.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn remove_shape_designer_properties_with_limits<'s, 'k>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        shape: impl Into<crate::shape::Key<'k>>,
        limits: crate::shape::designer::Limits,
    ) -> Result<crate::shape::designer::PropertiesSnapshot> {
        self.ensure_graph_current("remove_shape_designer_properties")?;
        let (slide_name, _) = self.resolve_slide_identity(slide.into())?;
        let mut edit = crate::shape::designer::load_properties_with_limits(
            &self.opc,
            &slide_name,
            shape,
            limits,
        )?
        .edit();
        edit.remove();
        let commit = edit.commit()?;
        let changed = !commit.is_noop();
        self.publish_designer(changed, move |opc| {
            crate::shape::designer::apply_properties(opc, commit)
        })
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
        self.ensure_graph_current("model3ds")?;
        let slide_name = self.resolve_slide(slide.into())?;
        crate::model3d::package::load_all(&self.opc, &slide_name)
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
        self.ensure_graph_current("model3d")?;
        let slide_name = self.resolve_slide(slide.into())?;
        crate::model3d::package::load(&self.opc, &slide_name, shape.into())
    }

    /// Replace one existing model3d owner transactionally.
    ///
    /// A change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006) and leaves the package unchanged: remove the signature
    /// explicitly first with [`Self::unsign`]. An exact no-op keeps the
    /// signature.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn put_model3d<'s, 'k>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        shape: impl Into<crate::shape::Key<'k>>,
        model: crate::model3d::Model,
    ) -> Result<Option<crate::model3d::Model>> {
        self.ensure_graph_current("put_model3d")?;
        let slide_name = self.resolve_slide(slide.into())?;
        self.edit_typed(move |opc| {
            crate::model3d::package::put(opc, &slide_name, shape.into(), model)
        })
    }

    /// Remove one model3d owner and collect unreachable binary resources.
    ///
    /// A change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006) and leaves the package unchanged: remove the signature
    /// explicitly first with [`Self::unsign`]. An exact no-op keeps the
    /// signature.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn remove_model3d<'s, 'k>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        shape: impl Into<crate::shape::Key<'k>>,
    ) -> Result<Option<crate::model3d::Model>> {
        self.ensure_graph_current("remove_model3d")?;
        let slide_name = self.resolve_slide(slide.into())?;
        self.edit_typed(move |opc| crate::model3d::package::remove(opc, &slide_name, shape.into()))
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

    /// Create or replace the presentation's table-style catalog atomically.
    ///
    /// A change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006) and leaves the package unchanged: remove the signature
    /// explicitly first with [`Self::unsign`]. An exact no-op keeps the
    /// signature.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn put_styles(&mut self, styles: crate::table::style::List) -> Result<bool> {
        self.edit_typed(move |opc| crate::table::style::put(opc, styles))
    }

    /// Remove the presentation's optional table-style catalog atomically.
    ///
    /// A change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006) and leaves the package unchanged: remove the signature
    /// explicitly first with [`Self::unsign`]. An exact no-op keeps the
    /// signature.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn remove_styles(&mut self) -> Result<Option<crate::table::style::List>> {
        self.edit_typed(crate::table::style::remove)
    }

    /// Add a slide master and update the `PresentationML` relationship graph.
    ///
    /// A change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006) and leaves the package unchanged: remove the signature
    /// explicitly first with [`Self::unsign`]. An exact no-op keeps the
    /// signature.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn add_slide_master(&mut self) -> Result<crate::master_layout::AuthoredSlideMaster> {
        self.edit_typed(crate::master_layout::add_slide_master)
    }

    /// Add a layout to an existing master and update both sides of the graph.
    ///
    /// A change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006) and leaves the package unchanged: remove the signature
    /// explicitly first with [`Self::unsign`]. An exact no-op keeps the
    /// signature.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn add_slide_layout(
        &mut self,
        master_part_name: &PackURI,
        kind: crate::master_layout::SlideLayoutKind,
        name: &str,
        placeholders: &[crate::master_layout::PlaceholderSpec],
    ) -> Result<crate::master_layout::AuthoredSlideLayout> {
        self.synchronize_legacy_raw_graph("add_slide_layout")?;
        crate::master_layout::preflight_add_slide_layout(
            &self.opc,
            master_part_name,
            kind,
            name,
            placeholders,
        )?;
        let result = self.edit_typed(|opc| {
            crate::master_layout::add_slide_layout(opc, master_part_name, kind, name, placeholders)
        });
        if result.is_ok() {
            self.mutable_pres = None;
        }
        result
    }

    /// Add or update one slide, master, or layout placeholder shape.
    ///
    /// A change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006) and leaves the package unchanged: remove the signature
    /// explicitly first with [`Self::unsign`]. An exact no-op keeps the
    /// signature.
    ///
    /// # Errors
    ///
    /// Returns an error if the output cannot be encoded or written.
    pub fn store_placeholder_shape(
        &mut self,
        part_name: &PackURI,
        spec: &crate::master_layout::PlaceholderSpec,
    ) -> Result<()> {
        self.synchronize_legacy_raw_graph("store_placeholder_shape")?;
        if crate::master_layout::preflight_store_placeholder_shape(&self.opc, part_name, spec)? {
            return Ok(());
        }
        let result = self
            .edit_typed(|opc| crate::master_layout::store_placeholder_shape(opc, part_name, spec));
        if result.is_ok() {
            self.mutable_pres = None;
        }
        result
    }

    /// Read one source-bound PowerPoint 2023 placeholder type extension.
    ///
    /// The snapshot retains the exact slide, master, or layout XML needed for a
    /// scalar, lossless edit. Markup-compatibility owners that require branch
    /// selection are rejected by the source-preserving transaction.
    ///
    /// # Errors
    ///
    /// Returns an error if the owner, selected shape, or p232 extension is
    /// missing or malformed.
    pub fn placeholder_type_extension_snapshot<'a>(
        &self,
        part_name: &PackURI,
        key: impl Into<crate::shape::Key<'a>>,
    ) -> Result<crate::master_layout::PlaceholderTypeExtensionSnapshot> {
        self.ensure_graph_current("placeholder_type_extension_snapshot")?;
        crate::master_layout::load_placeholder_type_extension_snapshot(&self.opc, part_name, key)
    }

    /// Apply a source-checked reversible p232 placeholder patch.
    ///
    /// A changed signed source is rejected. Use
    /// [`Self::apply_placeholder_type_extension_patch_with_policy`] when the
    /// caller has deliberately chosen signature invalidation.
    ///
    /// # Errors
    ///
    /// Returns an error if the selected owner is stale, malformed, or signed.
    pub fn apply_placeholder_type_extension_patch(
        &mut self,
        patch: &crate::master_layout::PlaceholderTypeExtensionPatch,
    ) -> Result<crate::master_layout::PlaceholderTypeExtensionSnapshot> {
        self.apply_placeholder_type_extension_patch_with_policy(
            patch,
            crate::master_layout::PlaceholderSignaturePolicy::Reject,
        )
    }

    /// Apply a p232 placeholder patch under an explicit signature policy.
    ///
    /// # Errors
    ///
    /// Returns an error if the selected owner is stale or malformed, or if
    /// the chosen policy cannot publish the package.
    pub fn apply_placeholder_type_extension_patch_with_policy(
        &mut self,
        patch: &crate::master_layout::PlaceholderTypeExtensionPatch,
        policy: crate::master_layout::PlaceholderSignaturePolicy,
    ) -> Result<crate::master_layout::PlaceholderTypeExtensionSnapshot> {
        self.ensure_graph_current("apply_placeholder_type_extension_patch")?;
        self.ensure_plain_mutation("apply_placeholder_type_extension_patch")?;
        let changed = !patch.is_noop();
        let snapshot = crate::master_layout::apply_placeholder_type_extension_patch_with_policy(
            &mut self.opc,
            patch,
            policy,
        )?;
        if changed {
            self.mutable_pres = None;
        }
        Ok(snapshot)
    }

    /// Read an optional, source-bound PowerPoint 2023 placeholder extension
    /// from a slide, master, or layout owner.
    pub fn placeholder_type_extension_slot_snapshot<'a>(
        &self,
        part_name: &PackURI,
        key: impl Into<crate::shape::Key<'a>>,
    ) -> Result<crate::master_layout::PlaceholderTypeExtensionSlotSnapshot> {
        self.ensure_graph_current("placeholder_type_extension_slot_snapshot")?;
        crate::master_layout::load_placeholder_type_extension_slot_snapshot(
            &self.opc, part_name, key,
        )
    }

    /// Apply an optional p232 placeholder extension patch with the default
    /// signature policy.
    pub fn apply_placeholder_type_extension_slot_patch(
        &mut self,
        patch: &crate::master_layout::PlaceholderTypeExtensionSlotPatch,
    ) -> Result<crate::master_layout::PlaceholderTypeExtensionSlotSnapshot> {
        self.apply_placeholder_type_extension_slot_patch_with_policy(
            patch,
            crate::master_layout::PlaceholderSignaturePolicy::Reject,
        )
    }

    /// Apply an optional p232 placeholder extension patch with an explicit
    /// signature disposition.
    pub fn apply_placeholder_type_extension_slot_patch_with_policy(
        &mut self,
        patch: &crate::master_layout::PlaceholderTypeExtensionSlotPatch,
        policy: crate::master_layout::PlaceholderSignaturePolicy,
    ) -> Result<crate::master_layout::PlaceholderTypeExtensionSlotSnapshot> {
        self.ensure_graph_current("apply_placeholder_type_extension_slot_patch")?;
        self.ensure_plain_mutation("apply_placeholder_type_extension_slot_patch")?;
        let changed = !patch.is_noop();
        let snapshot =
            crate::master_layout::apply_placeholder_type_extension_slot_patch_with_policy(
                &mut self.opc,
                patch,
                policy,
            )?;
        if changed {
            self.mutable_pres = None;
        }
        Ok(snapshot)
    }

    /// Remove an unreferenced layout and its owning relationship.
    ///
    /// A change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006) and leaves the package unchanged: remove the signature
    /// explicitly first with [`Self::unsign`]. An exact no-op keeps the
    /// signature.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn remove_slide_layout(&mut self, layout_part_name: &PackURI) -> Result<()> {
        self.edit_typed(|opc| crate::master_layout::remove_slide_layout(opc, layout_part_name))
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
        self.with_opc(crate::presentation_properties::metadata::slide_sync::load)
    }

    /// Attach one slide-library synchronization part transactionally.
    ///
    /// # Errors
    ///
    /// Returns an error if the output cannot be encoded or written.
    pub fn store_slide_sync(
        &mut self,
        value: &crate::presentation_properties::metadata::slide_sync::Part,
    ) -> Result<()> {
        self.edit_typed(|opc| {
            crate::presentation_properties::metadata::slide_sync::store(opc, value)
        })
    }

    /// Remove every package signature explicitly (ADR 0006).
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
    /// byte for byte. A package that carries no signature, and no
    /// signature-policy requirement, is left untouched.
    ///
    /// ```
    /// use litchi_pptx::Package;
    ///
    /// let mut package = Package::from_reader(std::io::Cursor::new(Package::new()?.to_bytes()?))?;
    /// package.unsign()?;
    /// assert!(!package.opc()?.is_signed());
    /// # Ok::<(), litchi_pptx::Error>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns the encryption-policy error for a package opened with a
    /// password, whose typed mutations are refused, and an error when pending
    /// authoring state cannot be published first.
    pub fn unsign(&mut self) -> Result<&mut Self> {
        if self.opc.is_signed() || self.opc.requires_signature_edit_policy() {
            self.edit_typed(|opc| {
                opc.unsign();
                Ok(())
            })?;
        }
        Ok(self)
    }

    /// Run one transactional low-level clear-OPC edit.
    ///
    /// The closure receives the current graph only after pending authoring
    /// state has been published. Any error rolls the graph back to its exact
    /// pre-edit snapshot; successful edits commit the candidate in place.
    /// Callers that need a typed, semantic operation should prefer the
    /// contextual methods on this facade. For retained encryption provenance,
    /// choosing this raw escape hatch explicitly declassifies the package. A
    /// successful edit clears provenance; a failed edit preserves it.
    ///
    /// Signature state is preserved (ADR 0006). A closure that changes a
    /// signed package must remove the signature explicitly with
    /// [`OpcPackage::unsign`], or the caller calls [`Self::unsign`] before
    /// the edit; deleting signature parts alone does not authorize the change
    /// of a signature the package was opened with. The package-level edit
    /// families refuse a signed package themselves. A closure that changes
    /// nothing keeps the signature byte for byte.
    ///
    /// # Errors
    ///
    /// Returns the closure's error, and [`Error::Opc`] with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// for a changed signed package that still carries its signature; the
    /// graph is rolled back either way.
    pub fn edit_opc<T>(
        &mut self,
        operation: impl FnOnce(&mut OpcPackage) -> Result<T>,
    ) -> Result<T> {
        let value = self.edit_raw(|opc| {
            let signed_source = opc
                .is_signed()
                .then(|| litchi_opc::PackageWriter::to_bytes(opc))
                .transpose()?;
            let value = operation(opc)?;
            // Unless the closure removed the signature explicitly, it must
            // have changed nothing: a signed candidate that publishes other
            // bytes is refused, and one the writer refuses to publish (a
            // changed signed source) is refused with the writer's error.
            if let Some(source) = signed_source
                && crate::package::signature::requires_disposition(opc)
                && litchi_opc::PackageWriter::to_bytes(opc)? != source
            {
                return Err(Error::Opc(
                    litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy,
                ));
            }
            Ok(value)
        })?;
        // A raw graph edit cannot be reflected into the lossless mutable
        // writer. Retire that facade after publication so later authoring
        // cannot overwrite the committed OPC graph with stale state.
        self.mutable_pres = None;
        #[cfg(feature = "encryption")]
        {
            self.encryption = litchi_ooxml_common::package_encryption::PackageEncryption::plain();
        }
        Ok(value)
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
        self.part_digests = std::sync::OnceLock::new();
        if let Some(memo) = Self::memo_for(&self.opc, snapshot) {
            self.part_digests = std::sync::OnceLock::from(memo);
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
        snapshot
            .part_digests
            .project(opc)
            .ok()
            .map(std::sync::Arc::new)
    }

    /// The payload-digest memo this package retains for its current graph.
    pub(crate) fn part_digest_memo(&self) -> Option<&crate::opened::PartDigests> {
        self.part_digests.get().map(std::sync::Arc::as_ref)
    }

    /// Drop the retained payload-digest memo.
    ///
    /// Called by every mutation that does not produce an opened-presentation
    /// snapshot to adopt, so the facade never retains a payload allocation its
    /// own graph has replaced. The next capture hashes from cold, which costs
    /// time and can never change a value.
    pub(crate) fn release_part_digests(&mut self) {
        self.part_digests = std::sync::OnceLock::new();
    }

    pub(crate) fn ensure_plain_mutation(&self, operation: &'static str) -> Result<()> {
        let _ = operation;
        #[cfg(feature = "encryption")]
        self.encryption
            .ordinary_output()
            .map_err(|source| Error::EncryptionPolicy { operation, source })?;
        Ok(())
    }

    /// Publish pending mutable-writer state before a legacy raw-owner preflight.
    /// The legacy owner operations inspect the canonical OPC graph, while the
    /// new-package writer may still hold newly added slide parts only in its
    /// mutable model. Successful raw-owner mutations retire that model below;
    /// otherwise a later presentation flush could overwrite the raw edit.
    fn synchronize_legacy_raw_graph(&mut self, operation: &'static str) -> Result<()> {
        self.ensure_plain_mutation(operation)?;
        self.flush_presentation()
    }

    pub(super) fn ensure_graph_current(&self, operation: &'static str) -> Result<()> {
        if self.is_modified() {
            return Err(Error::UnsafeEdit {
                operation,
                reason: super::codec::STALE_PRESENTATION_GRAPH_REASON,
            });
        }
        Ok(())
    }

    fn publish_designer<T>(
        &mut self,
        changed: bool,
        operation: impl FnOnce(&mut OpcPackage) -> Result<T>,
    ) -> Result<T> {
        let value = self.edit_typed(operation)?;
        if changed {
            self.mutable_pres = None;
        }
        Ok(value)
    }

    fn resolve_slide_identity(&self, key: crate::slide::Selector<'_>) -> Result<(PackURI, u32)> {
        let presentation = self.presentation()?;
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

    fn resolve_slide(&self, key: crate::slide::Selector<'_>) -> Result<PackURI> {
        let presentation = self.presentation()?;
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
}
