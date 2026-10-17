//! Located style, numbering and story catalogs of the owning package
//! (record 0319, L-DOCX-R5/R6). Each accessor reads the main part's
//! relationships and, where it returns one, the one related part.

use std::sync::Arc;

use super::super::model::{Package, Result};

impl Package {
    /// The styles part, located, or `None` when the document has none.
    ///
    /// ```
    /// let package = litchi_docx::Builder::new()?.build()?;
    /// let catalog = package.style_catalog()?;
    /// assert!(catalog.is_none_or(|catalog| catalog.get("NoSuchStyle").is_none()));
    /// # Ok::<(), litchi_docx::Error>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns an error for an unreadable relationship or a malformed part.
    pub fn style_catalog(&self) -> Result<Option<crate::styles::catalog::Catalog>> {
        match self.related_xml("/styles")? {
            Some(xml) => Ok(Some(crate::styles::catalog::Catalog::from_xml(xml)?)),
            None => Ok(None),
        }
    }

    /// The numbering part, located, or `None` when the document has none.
    ///
    /// # Errors
    ///
    /// Returns an error for an unreadable relationship or a malformed part.
    pub fn numbering_catalog(&self) -> Result<Option<crate::numbering::catalog::Catalog>> {
        match self.related_xml("/numbering")? {
            Some(xml) => Ok(Some(crate::numbering::catalog::Catalog::from_xml(xml)?)),
            None => Ok(None),
        }
    }

    /// The header and footer parts the main document relates, in
    /// relationship order; no part is read.
    ///
    /// # Errors
    ///
    /// Returns an error for unreadable relationships.
    pub fn header_footer_entries(&self) -> Result<Vec<crate::header_footer::catalog::Entry>> {
        let main = self.opc.main_document_part()?;
        crate::header_footer::catalog::entries_of(main.rels())
    }

    /// The header and footer references of every section, in document
    /// order, each with its role; only the main part is read.
    ///
    /// # Errors
    ///
    /// Returns an error for a main part that is not well-formed.
    pub fn header_footer_references(
        &self,
    ) -> Result<Vec<(crate::header_footer::Role, crate::section::Reference)>> {
        let main = self.opc.main_document_part()?;
        crate::header_footer::catalog::references_of(main.blob())
    }

    /// The payload of part `name`, or `None` when the package has no such
    /// part.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid part name.
    pub fn part_xml(&self, name: &str) -> Result<Option<Arc<Vec<u8>>>> {
        let uri = litchi_opc::PackURI::new(name.to_owned()).map_err(crate::Error::Invalid)?;
        Ok(self.opc.get_part(&uri).ok().map(|part| part.blob_arc()))
    }

    /// The payload of the first internal part the main part relates by a
    /// relationship type ending in `suffix`.
    fn related_xml(&self, suffix: &str) -> Result<Option<Arc<Vec<u8>>>> {
        let main = self.opc.main_document_part()?;
        let Some(relationship) = main.rels().iter().find(|relationship| {
            !relationship.is_external() && relationship.reltype().ends_with(suffix)
        }) else {
            return Ok(None);
        };
        let target = relationship.target_partname()?;
        Ok(Some(self.opc.get_part(&target)?.blob_arc()))
    }
}
