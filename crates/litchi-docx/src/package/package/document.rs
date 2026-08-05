//! Mutable document acquisition for the package facade.

use super::super::model::*;

impl Package {
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
                .map_err(|e| Error::InvalidUri(format!("document URI: {}", e)))?;

            // Try to get existing document content
            if let Ok(part) = self.opc.get_part(&doc_uri) {
                let xml = std::str::from_utf8(part.blob())
                    .map_err(|e| Error::InvalidFormat(format!("Invalid UTF-8: {}", e)))?;
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
