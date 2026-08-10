//! Mutable document acquisition for the package facade.

use super::super::model::{Error, MutableDocument, PackURI, Package, Result};
use super::transfer::{
    apply_transfer_graph, relationship_graph_digest, relationship_graph_digest_opc,
};

impl Package {
    /// Start the ordinary immutable main-document edit directly from this
    /// opened package.
    ///
    /// # Errors
    ///
    /// Returns a typed snapshot error when package state is stale, malformed,
    /// or outside the document transaction bounds.
    pub fn edit_document(
        &self,
    ) -> std::result::Result<crate::document::Edit, crate::document::TransactionError> {
        self.document_snapshot().map(|snapshot| snapshot.edit())
    }

    /// Capture an immutable, source-preserving main-document snapshot.
    ///
    /// # Errors
    ///
    /// Returns a typed transaction error when the package state is stale or
    /// the main document is missing, malformed, or over its resource bounds.
    pub fn document_snapshot(
        &self,
    ) -> std::result::Result<crate::document::Snapshot, crate::document::TransactionError> {
        self.ensure_story_opc_current("document_snapshot")?;
        let main = self.opc.main_document_part().map_err(Error::from)?;
        crate::document::Snapshot::from_xml(main.blob().to_vec())
    }

    /// Apply a main-document patch atomically to its exact source package.
    ///
    /// A stale patch leaves the package untouched. An exact no-op preserves
    /// signatures and the main-part payload allocation; a real edit validates
    /// the complete candidate facade state before publication.
    ///
    /// # Errors
    ///
    /// Returns a stale-source or package-validation error without publishing
    /// any partial package mutation.
    pub fn apply_document_patch(
        &mut self,
        patch: &crate::document::Patch,
    ) -> std::result::Result<crate::document::Snapshot, crate::document::TransactionError> {
        let current = self.document_snapshot()?;
        self.validate_transfer_operations(patch.operations())?;
        let graph_transition = transfer_graph_transition(patch.operations())?;
        let candidate = patch.apply(&current)?;
        if !patch.changed() {
            return Ok(candidate);
        }
        let replacement = candidate.xml_bytes().to_vec();
        self.edit_semantic_opc("apply_document_patch", move |opc| {
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
            opc.get_part_mut(&main_name)?.set_blob(replacement);
            Ok(())
        })?;
        Ok(candidate)
    }

    /// Commit and atomically publish an edit created by [`Self::edit_document`].
    ///
    /// # Errors
    ///
    /// Returns a commit, stale-source, compaction, or package-validation error
    /// without publishing a partial edit.
    pub fn publish_document_edit(
        &mut self,
        edit: crate::document::Edit,
    ) -> std::result::Result<crate::document::Commit, crate::document::TransactionError> {
        let commit = edit.commit()?;
        self.publish_document_commit(commit)
    }

    /// Atomically publish a previously committed document edit.
    ///
    /// This is the common package boundary for ordinary edits, deterministic
    /// composition, and resolved three-way plans.
    ///
    /// # Errors
    ///
    /// Returns a stale-source or package-validation error without publishing
    /// any partial mutation.
    pub fn publish_document_commit(
        &mut self,
        commit: crate::document::Commit,
    ) -> std::result::Result<crate::document::Commit, crate::document::TransactionError> {
        self.apply_document_patch(commit.patch())?;
        Ok(commit)
    }

    /// Publish a commit and record it in bounded history as one coupled action.
    ///
    /// History capacity is preflighted before package publication, and the
    /// exact history head must match the commit source.
    ///
    /// # Errors
    ///
    /// Returns a stale-source, history-bound, or package-validation error.
    pub fn publish_document_commit_with_history(
        &mut self,
        commit: crate::document::Commit,
        history: &mut crate::document::History,
    ) -> std::result::Result<Vec<crate::document::Snapshot>, crate::document::TransactionError>
    {
        if history.current().xml_bytes() != commit.patch().source().xml_bytes() {
            return Err(crate::document::TransactionError::StaleSource);
        }
        if !commit.patch().changed() {
            self.apply_document_patch(commit.patch())?;
            return Ok(Vec::new());
        }
        history.ensure_can_record(&commit)?;
        self.apply_document_patch(commit.patch())?;
        history.record(commit)
    }

    /// Publish one bounded undo transition atomically with the package graph.
    ///
    /// # Errors
    ///
    /// Returns a stale-source or package-validation error. A publication
    /// failure restores the history cursor before returning.
    pub fn undo_document(
        &mut self,
        history: &mut crate::document::History,
    ) -> std::result::Result<bool, crate::document::TransactionError> {
        self.publish_history_transition(history, false)
    }

    /// Publish one bounded redo transition atomically with the package graph.
    ///
    /// # Errors
    ///
    /// Returns a stale-source or package-validation error. A publication
    /// failure restores the history cursor before returning.
    pub fn redo_document(
        &mut self,
        history: &mut crate::document::History,
    ) -> std::result::Result<bool, crate::document::TransactionError> {
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
    pub fn apply_durable_document_patch<Mode>(
        &mut self,
        patch: &litchi_core::patch::Patch<Mode>,
    ) -> std::result::Result<crate::document::Snapshot, crate::document::TransactionError> {
        let current = self.document_snapshot()?;
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
                    return Err(crate::document::TransactionError::StaleSource);
                }
            }
        }
        let transfer_operations = crate::document::durable_transfer_operations(patch)?;
        let graph_transition = transfer_graph_transition(&transfer_operations)?;
        let candidate = current.apply_durable(patch)?;
        if candidate.xml_bytes() == current.xml_bytes() {
            return Ok(candidate);
        }
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
        &mut self,
        history: &mut crate::document::History,
        redo: bool,
    ) -> std::result::Result<bool, crate::document::TransactionError> {
        let current = self.document_snapshot()?;
        if current.xml_bytes() != history.current().xml_bytes() {
            return Err(crate::document::TransactionError::StaleSource);
        }
        let moved = if redo { history.redo() } else { history.undo() };
        if !moved {
            return Ok(false);
        }
        let graph_transition = history.take_graph_transition();
        if let Some((transition, forward)) = &graph_transition {
            let expected_source = if *forward {
                transition.before_digest.as_ref()
            } else {
                transition.after_digest.as_ref()
            };
            if relationship_graph_digest(self)? != expected_source {
                if redo {
                    let _restored = history.undo();
                } else {
                    let _restored = history.redo();
                }
                let _pending = history.take_graph_transition();
                return Err(crate::document::TransactionError::StaleSource);
            }
        }
        let replacement = history.current().xml_bytes().to_vec();
        let result = self.edit_semantic_opc("publish_document_history", move |opc| {
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
            return Err(crate::document::TransactionError::from(error));
        }
        Ok(true)
    }

    fn validate_transfer_operations(
        &self,
        operations: &[crate::document::Operation],
    ) -> std::result::Result<(), crate::document::TransactionError> {
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
            return Err(crate::document::TransactionError::StaleSource);
        }
        Ok(())
    }

    /// Get a mutable document for writing and modification.
    ///
    /// This returns a `MutableDocument` that allows you to add and modify
    /// paragraphs, tables, and other document elements.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_docx::Package;
    ///
    /// let mut pkg = Package::new()?;
    /// let mut doc = pkg.document_mut()?;
    ///
    /// // Add content
    /// doc.add_paragraph_with_text("Hello, World!");
    /// let para = doc.add_paragraph();
    /// para.add_run_with_text("Bold text").bold(true);
    ///
    /// // Add a table
    /// let table = doc.add_table(3, 2);
    /// if let Some(cell) = table.cell(0, 0) {
    ///     cell.set_text("Header 1");
    /// }
    ///
    /// pkg.save("output.docx")?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns an error when mutable semantic state cannot be synchronized or
    /// the main document XML is invalid.
    pub fn document_mut(&mut self) -> Result<&mut MutableDocument> {
        if self.raw_edit_committed {
            return Err(Error::UnsafeEdit {
                format: "DOCX",
                operation: "document_mut",
                reason: "a raw OPC edit committed; use edit_opc for further low-level changes",
            });
        }

        // If we don't have a mutable document, try to load it from the package
        if self.mutable_doc.is_none() {
            let doc_uri = PackURI::new("/word/document.xml")
                .map_err(|error| Error::InvalidUri(format!("document URI: {error}")))?;

            // Try to get existing document content
            if let Ok(part) = self.opc.get_part(&doc_uri) {
                let xml = std::str::from_utf8(part.blob())
                    .map_err(|error| Error::InvalidFormat(format!("Invalid UTF-8: {error}")))?;
                self.mutable_doc = Some(MutableDocument::from_xml(xml)?);
            } else {
                // Create a new empty document
                self.mutable_doc = Some(MutableDocument::new());
            }
        }

        self.mutable_doc.as_mut().ok_or_else(|| {
            Error::InvalidFormat("mutable document initialization did not complete".into())
        })
    }
}

fn transfer_graph_transition(
    operations: &[crate::document::Operation],
) -> std::result::Result<
    Option<(
        std::sync::Arc<crate::document::TransferGraph>,
        bool,
        std::sync::Arc<str>,
    )>,
    crate::document::TransactionError,
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
            | crate::document::Operation::ReplaceSimpleFieldText { .. }
            | crate::document::Operation::ReplaceComplexFieldText { .. }
            | crate::document::Operation::ReplaceRevisionText { .. }
            | crate::document::Operation::ReplaceContentControlText { .. }
            | crate::document::Operation::ReplaceNestedContentControlText { .. }
            | crate::document::Operation::ReplaceBlockContentControlParagraphText { .. }
            | crate::document::Operation::ReplaceCellText { .. }
            | crate::document::Operation::ReplaceCellParagraphText { .. }
            | crate::document::Operation::ReplaceNestedCellParagraphText { .. }
            | crate::document::Operation::InsertParagraph { .. }
            | crate::document::Operation::RemoveParagraph { .. } => None,
        };
        if let Some(selected_transition) = candidate {
            if selected.is_some() {
                return Err(crate::document::TransactionError::InvalidDurable(
                    "one commit cannot publish multiple dependency subgraphs".into(),
                ));
            }
            selected = Some(selected_transition);
        }
    }
    Ok(selected)
}
