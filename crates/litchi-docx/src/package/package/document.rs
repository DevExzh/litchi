//! Mutable document acquisition for the package facade.

use super::super::model::{Error, MutableDocument, PackURI, Package, Result};

impl Package {
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
        let candidate = patch.apply(&current)?;
        if !patch.changed() {
            return Ok(candidate);
        }
        let replacement = candidate.xml_bytes().to_vec();
        self.edit_semantic_opc("apply_document_patch", move |opc| {
            let main_name = opc.main_document_part()?.partname().clone();
            opc.get_part_mut(&main_name)?.set_blob(replacement);
            Ok(())
        })?;
        Ok(candidate)
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
