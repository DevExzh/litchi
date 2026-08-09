//! Relationship and state validation for inert document package views.

use crate::alt::{Chunk, Part, Target, is_relationship};
use crate::error::{Error, Result};

use super::super::model::Document;

impl Document<'_> {
    /// Resolve an alternative-format anchor to its borrowed opaque OPC payload.
    ///
    /// This validates the relationship type and internal target but never parses,
    /// imports, executes, or fetches the foreign content.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn resolve_alt<'b>(&'b self, chunk: &Chunk) -> Result<Part<'b>> {
        let relationship = self
            .part
            .part()
            .rels()
            .get(chunk.relationship().as_str())
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "altChunk relationship '{}' is missing",
                    chunk.relationship().as_str()
                ))
            })?;
        if relationship.is_external() {
            return Err(Error::InvalidFormat(
                "altChunk relationship must have an internal target".into(),
            ));
        }
        if !is_relationship(relationship.reltype()) {
            return Err(Error::InvalidFormat(format!(
                "altChunk relationship '{}' has invalid type '{}'",
                chunk.relationship().as_str(),
                relationship.reltype()
            )));
        }
        let target = relationship
            .target_partname()
            .map_err(|error| Error::InvalidFormat(format!("invalid altChunk target: {error}")))?;
        let part = self.opc.get_part(&target).map_err(|error| {
            Error::PartNotFound(format!("altChunk target '{}': {error}", target.as_str()))
        })?;
        Ok(Part::new(part))
    }

    /// Resolve an alternative-format target without fetching or interpreting it.
    ///
    /// Internal targets are returned as opaque package bytes. External targets
    /// are returned as their relationship URI and are never accessed.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn alt_target<'b>(&'b self, chunk: &Chunk) -> Result<Target<'b>> {
        let relationship = self
            .part
            .part()
            .rels()
            .get(chunk.relationship().as_str())
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "altChunk relationship {:?} is missing",
                    chunk.relationship().as_str()
                ))
            })?;
        if !is_relationship(relationship.reltype()) {
            return Err(Error::InvalidFormat(format!(
                "relationship {:?} is not an alternative-format import",
                chunk.relationship().as_str()
            )));
        }
        if relationship.is_external() {
            return Ok(Target::Link(relationship.target_ref()));
        }
        self.resolve_alt(chunk).map(Target::Part)
    }
    /// Check if the document is protected.
    ///
    /// This is a convenience method that checks the settings for protection status.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    ///
    /// if doc.is_protected()? {
    ///     println!("This document is protected");
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn is_protected(&self) -> Result<bool> {
        Ok(self.settings()?.is_some_and(|s| s.is_protected()))
    }
}
