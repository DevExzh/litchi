//! Package entry points of the core-properties transaction family.

use super::super::model::{Package, Props};
use crate::metadata::{Edit, Error, Patch, Result, admit};

impl Package {
    /// Borrow the typed core properties, retaining package absence: `None`
    /// means the package has no core-properties part.
    ///
    /// ```
    /// use litchi_docx::Package;
    ///
    /// let package = Package::new()?;
    /// assert!(package.metadata().is_some());
    /// # Ok::<(), litchi_docx::Error>(())
    /// ```
    #[must_use]
    pub fn metadata(&self) -> Option<&Props> {
        self.properties.get()
    }

    /// Stage a core-properties change over this snapshot.
    ///
    /// ```
    /// use litchi_docx::Package;
    ///
    /// let package = Package::new()?;
    /// let mut edit = package.edit_metadata();
    /// edit.update(|props| props.subject = Some("Plans".to_owned()));
    /// let next = edit.commit()?.into_package();
    /// assert_eq!(
    ///     next.metadata().and_then(|props| props.subject.as_deref()),
    ///     Some("Plans")
    /// );
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    #[must_use]
    pub fn edit_metadata(&self) -> Edit {
        Edit::new(self.clone(), self.metadata().cloned())
    }

    /// Replay a metadata patch onto this snapshot and return a new package.
    ///
    /// ```
    /// use litchi_docx::Package;
    ///
    /// let package = Package::new()?;
    /// let mut edit = package.edit_metadata();
    /// edit.update(|props| props.title = Some("Draft".to_owned()));
    /// let (changed, patch) = edit.commit()?.into_parts();
    /// let restored = changed.apply_metadata(&patch.inverse())?;
    /// assert_eq!(restored.metadata(), package.metadata());
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package's core properties
    /// differ from the patch's base, and the [`Edit::commit`] errors
    /// otherwise.
    pub fn apply_metadata(&self, patch: &Patch) -> Result<Self> {
        if self.metadata() != patch.base() {
            return Err(Error::PatchConflict);
        }
        self.publish_metadata(patch.target())
            .map(|(package, _)| package)
    }

    /// Publish `target` as the core properties of a new package. Returns
    /// whether it differs from this package's value.
    pub(crate) fn publish_metadata(&self, target: Option<&Props>) -> Result<(Self, bool)> {
        if self.metadata() == target {
            return Ok((self.clone(), false));
        }
        if let Some(props) = target {
            admit(props)?;
        }
        if self.opc.is_signed() || self.opc.requires_signature_edit_policy() {
            return Err(Error::Signed);
        }
        let mut package = self.clone();
        package.edit_semantic_opc("edit_metadata", |opc| {
            match target {
                Some(props) => litchi_ooxml_common::properties::write(opc, props)?,
                None => litchi_ooxml_common::properties::clear(opc)?,
            };
            Ok(())
        })?;
        Ok((package, true))
    }
}
