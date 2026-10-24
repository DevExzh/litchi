//! Immutable document transactions and legacy history orchestration.

use super::super::model::{Error, Package};
use super::transfer::{
    apply_graph_change, apply_transfer_graph, relationship_graph_digest,
    relationship_graph_digest_opc,
};
use sha2::{Digest as _, Sha256};
use std::io::{self, Write};

struct ArtifactFingerprintWriter {
    hasher: Sha256,
}

impl Write for ArtifactFingerprintWriter {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        self.hasher.update(bytes);
        Ok(bytes.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

/// The main-document bytes `apply` publishes.
enum MainReplacement {
    /// The candidate's own allocation, with the proof that it passed the
    /// writer's source publication audit.
    Proven(xml_minifier::audit::VerifiedSource),
    /// A copy of the candidate's bytes, audited at publication.
    Copied(Vec<u8>),
}

/// How `apply` obtained the snapshot it publishes.
enum DocumentPatchSource {
    /// The opened package still retains the patch's exact source bytes, so the
    /// patch handed back its own retained target and no snapshot was built.
    Reused(crate::document::Snapshot),
    /// The byte proof did not hold; this is the freshly scanned current
    /// snapshot that `Patch::apply` must decide against.
    Rescanned(crate::document::Snapshot),
}

impl Package {
    fn materialize_edit(
        package: &Self,
        options: crate::document::EditOptions,
    ) -> Result<crate::document::TransactionEdit, crate::document::Error> {
        package
            .document_snapshot_with_options(Some(&options))
            .map(|snapshot| snapshot.edit())
    }

    /// Start the ordinary immutable main-document edit directly from this
    /// opened package. The edit retains this source; its first operation,
    /// projection or `commit()` performs the fallible document admission, and
    /// a successful commit returns a new package without mutating the caller.
    ///
    /// # Errors
    ///
    /// Admission errors are returned by the first operation, projection or
    /// commit, when package state is stale, malformed, or outside the bounds.
    pub fn edit(&self) -> Result<crate::document::Edit, crate::document::Error> {
        self.edit_with_options(crate::document::EditOptions::default())
    }

    /// Start the ordinary immutable main-document edit under one explicit
    /// document-edit profile.
    pub fn edit_with_options(
        &self,
        options: crate::document::EditOptions,
    ) -> Result<crate::document::Edit, crate::document::Error> {
        Ok(crate::document::Edit::deferred(
            self.clone(),
            options,
            Self::materialize_edit,
        ))
    }

    /// Start a bounded composition tied to this package's complete artifact.
    /// Reopened equal-content packages may publish it; a package with a
    /// different untouched part is refused before document staging.
    pub fn compose(
        &self,
        limits: litchi_core::patch::CompositionLimits,
    ) -> Result<crate::document::composition::Composition, crate::document::Error> {
        let source = self.document_snapshot()?;
        let fingerprint = self.artifact_fingerprint()?;
        Ok(source
            .compose(limits)
            .with_artifact_fingerprint(fingerprint))
    }

    fn artifact_fingerprint(&self) -> Result<[u8; 32], crate::document::Error> {
        if let Some(fingerprint) = self.opc.fingerprint() {
            return Ok(fingerprint);
        }
        let mut writer = ArtifactFingerprintWriter {
            hasher: Sha256::new(),
        };
        self.write_to(&mut writer)?;
        let digest = writer.hasher.finalize();
        let mut fingerprint = [0_u8; 32];
        fingerprint.copy_from_slice(&digest);
        // `SharedOpc` keeps this memo inline with the immutable graph state;
        // cloning the package copies the memo while a later mutation clears
        // only that branch.
        self.opc.remember_fingerprint(fingerprint);
        Ok(fingerprint)
    }

    /// Publish one joined document composition through the complete package
    /// owner. All accepted operations are staged in one edit and validated in
    /// one commit; the composition's exact source must still be current.
    pub fn apply_composition(
        &self,
        composition: crate::document::composition::Composition,
    ) -> Result<crate::document::Commit, crate::document::Error> {
        let (source, operations, artifact_fingerprint) = composition.into_parts();
        let current = self.document_snapshot()?;
        if current.xml_bytes() != source.xml_bytes() {
            return Err(crate::document::Error::PatchConflict);
        }
        if let Some(expected) = artifact_fingerprint
            && self.artifact_fingerprint()? != expected
        {
            return Err(crate::document::Error::PatchConflict);
        }
        let mut edit = crate::document::Edit::new(self.clone(), current.edit());
        edit.apply_composition_operations(&operations)?;
        edit.commit()
    }

    /// Capture an immutable, source-preserving main-document snapshot.
    ///
    /// The payload is admitted by the markup-compatibility rule of
    /// [`crate::source_backed::Package::document_snapshot`]: the snapshot
    /// edits the payload's own bytes while readers see the projection, so a
    /// payload whose projection selects, unwraps or hides content is refused
    /// with a typed unsafe-edit error, and a projection error is returned as
    /// is. Payloads that never mention the namespace, bind it without using
    /// it, or carry valid `mc:Ignorable` directives that hide nothing a
    /// reader sees are admitted, and so are payloads whose only selections
    /// are `mc:AlternateContent` wrappers inside runs (a drawing or text box
    /// with a fallback; record 0307), each an opaque inline object to the
    /// snapshot's edits.
    ///
    /// The snapshot shares the main part's payload allocation instead of
    /// copying it. `Part::set_blob` and `set_blob_shared` replace the part's
    /// `Arc` rather than mutating through it, and no route in the workspace
    /// mutates a part payload in place, so a snapshot taken here keeps the
    /// bytes it was built from for as long as it lives even after the package
    /// publishes a different main document. Sharing also lets
    /// [`Self::apply`]'s exact-source proof settle on pointer
    /// identity instead of a whole-part byte comparison.
    ///
    /// # Errors
    ///
    /// Returns a typed transaction error when the package state is stale or
    /// the main document is missing, malformed, or over its resource bounds.
    pub fn document_snapshot(&self) -> Result<crate::document::Snapshot, crate::document::Error> {
        self.document_snapshot_with_options(None)
    }

    fn document_snapshot_with_options(
        &self,
        options: Option<&crate::document::EditOptions>,
    ) -> Result<crate::document::Snapshot, crate::document::Error> {
        self.ensure_story_opc_current("document_snapshot")?;
        let main = self.opc.main_document_part().map_err(Error::from)?;
        let xml = main.blob_arc();
        if let Some(options) = options {
            options.admit_source_bytes(u64::try_from(xml.len()).unwrap_or(u64::MAX))?;
        }
        self.admit_main_document(&xml, "document_snapshot")?;
        crate::document::Snapshot::from_shared_xml(xml)
    }

    /// Prove the main-document payload admissible to a transaction over its
    /// own bytes, as the source-backed owner does
    /// ([`crate::parts::document_part::admit_document_source`]): one rule for
    /// both owners. A payload whose projection selects, unwraps or hides
    /// content is refused, and a projection error is returned as is. The
    /// verdict is remembered for the allocation it was proved on.
    fn admit_main_document(
        &self,
        xml: &std::sync::Arc<Vec<u8>>,
        operation: &'static str,
    ) -> Result<(), crate::document::Error> {
        let proved = |memo: &std::sync::Weak<Vec<u8>>| {
            memo.upgrade()
                .is_some_and(|admitted| std::sync::Arc::ptr_eq(&admitted, xml))
        };
        if proved(
            &self
                .main_admission
                .lock()
                .unwrap_or_else(std::sync::PoisonError::into_inner),
        ) {
            return Ok(());
        }
        crate::parts::document_part::admit_document_source(
            xml.as_slice(),
            operation,
            "document transactions do not support markup-compatibility branch selection",
        )?;
        *self
            .main_admission
            .lock()
            .unwrap_or_else(std::sync::PoisonError::into_inner) = std::sync::Arc::downgrade(xml);
        Ok(())
    }

    /// Answer `Patch::apply`'s exact-source proof from the retained main part
    /// when it can be answered from the bytes alone.
    ///
    /// `document_snapshot` copies the whole main part and rescans its layout;
    /// on the ordinary edit-and-save route that snapshot exists only to feed
    /// the patch's byte comparison and is dropped in the same statement. The
    /// comparison needs the bytes alone, so this proves it against
    /// `main.blob()` directly. Package staleness and main-part resolution are
    /// checked first, in the order `document_snapshot` checks them, so a stale
    /// or malformed package still fails the same way; anything the byte proof
    /// does not settle falls through to the unchanged snapshot route, which
    /// decides every refusal exactly as before.
    fn document_patch_source(
        &self,
        patch: &crate::document::Patch,
    ) -> Result<DocumentPatchSource, crate::document::Error> {
        self.ensure_story_opc_current("document_snapshot")?;
        let main = self.opc.main_document_part().map_err(Error::from)?;
        if let Some(candidate) = patch.target_for_exact_unmanaged_source(main.blob()) {
            // The rescanned route proves the payload in `document_snapshot`;
            // a reused candidate is held to the same admission before it is
            // published, however its source snapshot was built.
            if !patch.is_noop() {
                self.admit_main_document(&main.blob_arc(), "apply")?;
            }
            return Ok(DocumentPatchSource::Reused(candidate));
        }
        Ok(DocumentPatchSource::Rescanned(self.document_snapshot()?))
    }

    /// Apply a main-document patch and publish a new package.
    ///
    /// The source is retained unchanged on success and failure. An exact
    /// no-op shares its physical graph, metadata and retained writer state.
    ///
    /// ```
    /// use litchi_docx::Package;
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let package = Package::new()?;
    /// let commit = package.edit()?.commit()?;
    /// let next = package.apply(commit.patch())?.into_package();
    /// assert_eq!(package.document()?.text()?, next.document()?.text()?);
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// A stale patch leaves the package untouched. An exact no-op preserves
    /// signatures and the main-part payload allocation; a real edit validates
    /// the complete candidate facade state before publication. A real edit is
    /// refused, as on the source-backed publisher, while document or write
    /// protection is enforced or revisions are tracked, and when the current
    /// payload fails the markup-compatibility admission of
    /// [`Self::document_snapshot`].
    ///
    /// # Errors
    ///
    /// Returns a stale-source or package-validation error without publishing
    /// any partial package mutation.
    pub fn apply(
        &self,
        patch: &crate::document::Patch,
    ) -> crate::document::Result<crate::document::Commit> {
        let mut package = self.clone();
        package.publish_patch(patch)?;
        let changed = !patch.is_noop();
        let package = if changed { package } else { self.clone() };
        Ok(crate::document::Commit::from_parts(
            package,
            patch.clone(),
            crate::document::Diagnostics::new(patch.operations().len(), changed),
        ))
    }

    // Internal publication engine for the remaining legacy/history adapters.
    // The public patch entry always works on a share and returns a new owner.
    fn publish_patch(
        &mut self,
        patch: &crate::document::Patch,
    ) -> Result<crate::document::Snapshot, crate::document::Error> {
        let source = self.document_patch_source(patch)?;
        self.validate_transfer_operations(patch.operations())?;
        let graph_transition = transfer_graph_transition(patch.operations())?;
        let graph_changes: Vec<(std::sync::Arc<crate::document::GraphChange>, bool)> = patch
            .operations()
            .iter()
            .filter_map(|operation| match operation {
                crate::document::Operation::ChangeGraph { change, forward } => {
                    Some((std::sync::Arc::clone(change), *forward))
                },
                _ => None,
            })
            .collect();
        let candidate = match source {
            DocumentPatchSource::Reused(candidate) => candidate,
            DocumentPatchSource::Rescanned(current) => patch.apply(&current)?,
        };
        if patch.is_noop() {
            return Ok(candidate);
        }
        // Protection and tracked revisions refuse a changed publication here
        // exactly as they do on the source-backed publisher.
        crate::source_backed::document_policy::validate_changed_owned_document(&self.opc)?;
        // A commit that proved its candidate's exact allocation to the
        // writer's source audit publishes that allocation with the proof
        // (change 0754): the writer then does not audit those bytes again,
        // and the part shares the snapshot's bytes instead of copying them.
        // Any other candidate is copied and audited at publication as before.
        let replacement = match candidate.publication_proof() {
            Some(proof) => MainReplacement::Proven(proof.clone()),
            None => MainReplacement::Copied(candidate.xml_bytes().to_vec()),
        };
        self.edit_semantic_opc("apply", move |opc| {
            for (change, forward) in &graph_changes {
                apply_graph_change(opc, change, *forward)
                    .map_err(|error| Error::InvalidFormat(error.to_string()))?;
            }
            if let Some((graph, insert, expected_digest)) = graph_transition {
                apply_transfer_graph(opc, &graph, insert)
                    .map_err(|error| Error::InvalidFormat(error.to_string()))?;
                let actual = relationship_graph_digest_opc(opc)
                    .map_err(|error| Error::InvalidFormat(error.to_string()))?;
                if actual != expected_digest.as_ref() {
                    return Err(Error::InvalidFormat(
                        "paragraph transfer graph produced an unexpected target".into(),
                    ));
                }
            }
            let main_name = opc.main_document_part()?.partname().clone();
            let part = opc.get_part_mut(&main_name)?;
            match replacement {
                MainReplacement::Proven(proof) => part.set_blob_verified(proof),
                MainReplacement::Copied(bytes) => part.set_blob(bytes),
            }
            Ok(())
        })?;
        Ok(candidate)
    }

    /// Publish a snapshot commit and record it in bounded history as one
    /// coupled action; returns the published package.
    ///
    /// History capacity is preflighted before package publication, and the
    /// exact history head must match the commit source. This package is not
    /// changed.
    ///
    /// # Errors
    ///
    /// Returns a stale-source, history-bound, or package-validation error.
    pub fn publish_document_commit_with_history(
        &self,
        commit: crate::document::Commit,
        history: &mut crate::document::composition::History,
    ) -> Result<crate::document::Commit, crate::document::Error> {
        if history.current().xml_bytes() != commit.patch().source().xml_bytes() {
            return Err(crate::document::Error::PatchConflict);
        }
        let diagnostics = commit.diagnostics();
        let patch = commit.patch().clone();
        let mut next = self.clone();
        if patch.is_noop() {
            next.publish_patch(&patch)?;
            return Ok(crate::document::Commit::from_parts(
                next,
                patch,
                diagnostics,
            ));
        }
        history.ensure_can_record(&commit)?;
        next.publish_patch(&patch)?;
        history.record(commit)?;
        Ok(crate::document::Commit::from_parts(
            next,
            patch,
            diagnostics,
        ))
    }

    /// Publish one bounded undo transition; returns the published package,
    /// which is this package unchanged when there is nothing to undo.
    ///
    /// # Errors
    ///
    /// Returns a stale-source or package-validation error. A publication
    /// failure restores the history cursor before returning.
    pub fn undo_document(
        &self,
        history: &mut crate::document::composition::History,
    ) -> Result<crate::document::Commit, crate::document::Error> {
        self.publish_history_transition(history, false)
    }

    /// Publish one bounded redo transition; returns the published package,
    /// which is this package unchanged when there is nothing to redo.
    ///
    /// # Errors
    ///
    /// Returns a stale-source or package-validation error. A publication
    /// failure restores the history cursor before returning.
    pub fn redo_document(
        &self,
        history: &mut crate::document::composition::History,
    ) -> Result<crate::document::Commit, crate::document::Error> {
        self.publish_history_transition(history, true)
    }

    /// Apply a common durable semantic main-document patch atomically.
    ///
    /// The complete candidate facade is validated before publication. Direct
    /// hyperlink relationships and every untouched OPC part remain unchanged.
    ///
    /// # Errors
    ///
    /// Returns a durable-vocabulary, stale-source, semantic-precondition, or
    /// package-validation error without publishing a partial mutation.
    pub fn apply_durable_document<Mode>(
        &self,
        patch: &litchi_core::patch::Patch<Mode>,
    ) -> Result<crate::document::Commit, crate::document::Error> {
        let current = self.document_snapshot()?;
        let mut next = self.clone();
        let candidate = next.apply_durable_document_in_place(patch, current.clone())?;
        let changed = candidate.xml_bytes() != current.xml_bytes();
        let package = if changed { next } else { self.clone() };
        Ok(crate::document::Commit::from_parts(
            package,
            crate::document::Patch::transition(current, candidate),
            crate::document::Diagnostics::new(patch.operations().len(), changed),
        ))
    }

    fn apply_durable_document_in_place<Mode>(
        &mut self,
        patch: &litchi_core::patch::Patch<Mode>,
        current: crate::document::Snapshot,
    ) -> Result<crate::document::Snapshot, crate::document::Error> {
        let has_transfer = patch.operations().iter().any(|operation| {
            matches!(
                operation.op.as_str(),
                "paragraph.transfer.insert"
                    | "paragraph.transfer.remove"
                    | "document.restore-transfer.insert"
                    | "document.restore-transfer.remove"
            )
        });
        if has_transfer {
            let dependency_digest = relationship_graph_digest(self)?;
            for operation in patch.operations().iter().filter(|operation| {
                matches!(
                    operation.op.as_str(),
                    "paragraph.transfer.insert"
                        | "paragraph.transfer.remove"
                        | "document.restore-transfer.insert"
                        | "document.restore-transfer.remove"
                )
            }) {
                if operation
                    .preconditions
                    .get("dependency_sha256")
                    .and_then(serde_json::Value::as_str)
                    != Some(dependency_digest.as_str())
                {
                    return Err(crate::document::Error::PatchConflict);
                }
            }
        }
        if patch
            .operations()
            .iter()
            .any(|operation| operation.op == "graph.change")
        {
            return Err(crate::document::Error::InvalidDurable(
                "package graph changes publish through Package::apply".into(),
            ));
        }
        let transfer_operations = crate::document::durable_transfer_operations(patch)?;
        let graph_transition = transfer_graph_transition(&transfer_operations)?;
        let candidate = current.apply_durable(patch)?;
        if candidate.xml_bytes() == current.xml_bytes() {
            return Ok(candidate);
        }
        crate::source_backed::document_policy::validate_changed_owned_document(&self.opc)?;
        let replacement = candidate.xml_bytes().to_vec();
        self.edit_semantic_opc("apply_durable_document_patch", move |opc| {
            if let Some((graph, insert, expected_digest)) = graph_transition {
                apply_transfer_graph(opc, &graph, insert)
                    .map_err(|error| Error::InvalidFormat(error.to_string()))?;
                let actual = relationship_graph_digest_opc(opc)
                    .map_err(|error| Error::InvalidFormat(error.to_string()))?;
                if actual != expected_digest.as_ref() {
                    return Err(Error::InvalidFormat(
                        "durable paragraph transfer graph produced an unexpected target".into(),
                    ));
                }
            }
            let main_name = opc.main_document_part()?.partname().clone();
            opc.get_part_mut(&main_name)?.set_blob(replacement);
            Ok(())
        })?;
        Ok(candidate)
    }

    fn publish_history_transition(
        &self,
        history: &mut crate::document::composition::History,
        redo: bool,
    ) -> Result<crate::document::Commit, crate::document::Error> {
        let current = self.document_snapshot()?;
        if current.xml_bytes() != history.current().xml_bytes() {
            return Err(crate::document::Error::PatchConflict);
        }
        let moved = if redo { history.redo() } else { history.undo() };
        if !moved {
            let patch = crate::document::Patch::transition(current.clone(), current);
            return Ok(crate::document::Commit::from_parts(
                self.clone(),
                patch,
                crate::document::Diagnostics::new(0, false),
            ));
        }
        let mut next = self.clone();
        let graph_transition = history.take_graph_transition();
        if let Some((transition, forward)) = &graph_transition {
            let expected_source = if *forward {
                transition.before_digest.as_ref()
            } else {
                transition.after_digest.as_ref()
            };
            if relationship_graph_digest(&next)? != expected_source {
                if redo {
                    let _restored = history.undo();
                } else {
                    let _restored = history.redo();
                }
                let _pending = history.take_graph_transition();
                return Err(crate::document::Error::PatchConflict);
            }
        }
        let target = history.current().clone();
        let replacement = target.xml_bytes().to_vec();
        let result = next.edit_semantic_opc("publish_document_history", move |opc| {
            if let Some((transition, forward)) = graph_transition {
                let insert = if forward {
                    transition.forward_insert
                } else {
                    !transition.forward_insert
                };
                apply_transfer_graph(opc, &transition.graph, insert)
                    .map_err(|error| Error::InvalidFormat(error.to_string()))?;
                let expected_target = if forward {
                    transition.after_digest
                } else {
                    transition.before_digest
                };
                let actual = relationship_graph_digest_opc(opc)
                    .map_err(|error| Error::InvalidFormat(error.to_string()))?;
                if actual != expected_target.as_ref() {
                    return Err(Error::InvalidFormat(
                        "document history graph produced an unexpected target".into(),
                    ));
                }
            }
            let main_name = opc.main_document_part()?.partname().clone();
            opc.get_part_mut(&main_name)?.set_blob(replacement);
            Ok(())
        });
        if let Err(error) = result {
            if redo {
                let _restored = history.undo();
            } else {
                let _restored = history.redo();
            }
            let _pending = history.take_graph_transition();
            return Err(crate::document::Error::from(error));
        }
        let patch = crate::document::Patch::transition(current, target);
        Ok(crate::document::Commit::from_parts(
            next,
            patch,
            crate::document::Diagnostics::new(0, true),
        ))
    }

    fn validate_transfer_operations(
        &self,
        operations: &[crate::document::Operation],
    ) -> Result<(), crate::document::Error> {
        if !operations.iter().any(|operation| {
            matches!(
                operation,
                crate::document::Operation::InsertTransferredParagraph { .. }
                    | crate::document::Operation::RemoveTransferredParagraph { .. }
            )
        }) {
            return Ok(());
        }
        let expected = relationship_graph_digest(self)?;
        if operations.iter().any(|operation| {
            matches!(
                operation,
                crate::document::Operation::InsertTransferredParagraph {
                    dependency_digest,
                    ..
                } | crate::document::Operation::RemoveTransferredParagraph {
                    dependency_digest,
                    ..
                } if dependency_digest.as_ref() != expected.as_str()
            )
        }) {
            return Err(crate::document::Error::PatchConflict);
        }
        Ok(())
    }
}

fn transfer_graph_transition(
    operations: &[crate::document::Operation],
) -> Result<
    Option<(
        std::sync::Arc<crate::document::TransferGraph>,
        bool,
        std::sync::Arc<str>,
    )>,
    crate::document::Error,
> {
    let mut selected = None;
    for operation in operations {
        let candidate = match operation {
            crate::document::Operation::InsertTransferredParagraph {
                graph,
                inverse_dependency_digest,
                ..
            } if !graph.is_empty() => Some((
                std::sync::Arc::clone(graph),
                true,
                std::sync::Arc::clone(inverse_dependency_digest),
            )),
            crate::document::Operation::RemoveTransferredParagraph {
                graph,
                inverse_dependency_digest,
                ..
            } if !graph.is_empty() => Some((
                std::sync::Arc::clone(graph),
                false,
                std::sync::Arc::clone(inverse_dependency_digest),
            )),
            crate::document::Operation::InsertTransferredParagraph { .. }
            | crate::document::Operation::RemoveTransferredParagraph { .. }
            | crate::document::Operation::ReplaceParagraphText { .. }
            | crate::document::Operation::ReplaceHyperlinkText { .. }
            | crate::document::Operation::ReplaceRunText { .. }
            | crate::document::Operation::ReplaceRunTextNode { .. }
            | crate::document::Operation::MarkCachedRunFieldDirty { .. }
            | crate::document::Operation::ReplaceSimpleFieldText { .. }
            | crate::document::Operation::ReplaceComplexFieldText { .. }
            | crate::document::Operation::ReplaceRevisionText { .. }
            | crate::document::Operation::ApplyRevision { .. }
            | crate::document::Operation::ReplaceContentControlText { .. }
            | crate::document::Operation::ReplaceNestedContentControlText { .. }
            | crate::document::Operation::ReplaceNestedContentControlHyperlinkText { .. }
            | crate::document::Operation::ReplaceBlockContentControlParagraphText { .. }
            | crate::document::Operation::ReplaceBlockContentControlParagraphHyperlinkText {
                ..
            }
            | crate::document::Operation::ReplaceCellText { .. }
            | crate::document::Operation::ReplaceCellParagraphText { .. }
            | crate::document::Operation::ReplaceNestedCellParagraphText { .. }
            | crate::document::Operation::ReplaceNestedCellParagraphHyperlinkText { .. }
            | crate::document::Operation::InsertParagraph { .. }
            | crate::document::Operation::RemoveParagraph { .. }
            | crate::document::Operation::ReplaceRunProperties { .. }
            | crate::document::Operation::ReplaceParagraphProperties { .. }
            | crate::document::Operation::InsertParagraphFragment { .. }
            | crate::document::Operation::RemoveParagraphFragment { .. }
            | crate::document::Operation::InsertRun { .. }
            | crate::document::Operation::RemoveRun { .. }
            | crate::document::Operation::ReplaceElement { .. }
            | crate::document::Operation::InsertElement { .. }
            | crate::document::Operation::RemoveElement { .. }
            | crate::document::Operation::ChangeGraph { .. } => None,
        };
        if let Some(selected_transition) = candidate {
            if selected.is_some() {
                return Err(crate::document::Error::InvalidDurable(
                    "one commit cannot publish multiple dependency subgraphs".into(),
                ));
            }
            selected = Some(selected_transition);
        }
    }
    Ok(selected)
}
