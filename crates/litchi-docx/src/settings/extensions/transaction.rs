#![expect(
    clippy::shadow_reuse,
    reason = "parser bindings are intentionally refined after validation"
)]
//! The Word settings-extension family (record 0355): the `w14`/`w15`
//! extensions of the settings part, read as a [`Snapshot`] and changed
//! through an [`Edit`] that publishes a new package.
//!
//! ```
//! use litchi_docx::Package;
//! use litchi_docx::settings::extensions::OnOff;
//!
//! let package = Package::new()?;
//! let mut edit = package.edit_settings_extensions();
//! edit.set_conflict_mode(Some(OnOff::on()))?;
//! let commit = edit.commit()?;
//! assert!(commit.diagnostics().changed());
//! assert_eq!(
//!     commit.package().settings_extensions()?.extensions().conflict_mode(),
//!     Some(OnOff::on())
//! );
//! assert_eq!(package.settings_extensions()?.extensions().conflict_mode(), None);
//! # Ok::<(), litchi_docx::Error>(())
//! ```

use std::sync::Arc;

use super::codec;
use super::model::{Extensions, Guid, OnOff};
use crate::package::family::{Staged, graph_family};

/// An immutable complete `w:settings` snapshot with its typed extension view.
#[derive(Debug, Clone)]
pub struct Snapshot {
    xml: Arc<[u8]>,
    extensions: Extensions,
}

impl Snapshot {
    /// Parse and retain a bounded complete settings XML snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn from_xml(xml: impl Into<Vec<u8>>) -> Result<Self> {
        let xml = xml.into();
        let extensions = Extensions::parse(&xml)?;
        Ok(Self {
            xml: Arc::from(xml.into_boxed_slice()),
            extensions,
        })
    }

    /// Borrow the exact settings XML retained by this snapshot.
    #[inline]
    #[must_use]
    pub fn xml_bytes(&self) -> &[u8] {
        &self.xml
    }

    /// Borrow the contextual typed extension collection in source order.
    #[inline]
    #[must_use]
    pub const fn extensions(&self) -> &Extensions {
        &self.extensions
    }
}

graph_family!("settings-extension", own_limits);

pub use super::super::limit::LimitKind;

/// A staged settings-extension change over one package snapshot.
///
/// The settings part is read on the first staging verb. [`Self::commit`]
/// rewrites only the extension elements that changed and publishes a new
/// package; the source is never changed, and dropping the edit discards it.
#[derive(Clone)]
pub struct Edit {
    staged: Staged,
    extensions: Option<(Snapshot, Extensions)>,
}

impl std::fmt::Debug for Edit {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.debug_struct("Edit").finish_non_exhaustive()
    }
}

impl Edit {
    fn next(&mut self) -> Result<&mut Extensions> {
        if self.extensions.is_none() {
            let snapshot = self.staged.working().settings_extensions()?;
            let extensions = snapshot.extensions().clone();
            self.extensions = Some((snapshot, extensions));
        }
        Ok(&mut self
            .extensions
            .as_mut()
            .expect("settings extensions were loaded above")
            .1)
    }

    /// The staged extension collection.
    ///
    /// # Errors
    ///
    /// Returns an error if the settings part cannot be read.
    pub fn extensions(&mut self) -> Result<&Extensions> {
        Ok(&*self.next()?)
    }

    /// Set or remove `w14:chartTrackingRefBased`.
    ///
    /// # Errors
    ///
    /// Returns an error if the settings part cannot be read or the value is
    /// invalid.
    pub fn set_chart_tracking_ref_based(&mut self, value: Option<OnOff>) -> Result<&mut Self> {
        self.next()?.set_chart_tracking_ref_based(value)?;
        Ok(self)
    }

    /// Set or remove `w14:docId`.
    ///
    /// # Errors
    ///
    /// Returns an error if the settings part cannot be read or the value is
    /// invalid.
    pub fn set_document_id(&mut self, value: Option<u32>) -> Result<&mut Self> {
        self.next()?.set_document_id(value)?;
        Ok(self)
    }

    /// Set or remove `w15:docId`.
    ///
    /// # Errors
    ///
    /// Returns an error if the settings part cannot be read or the value is
    /// invalid.
    pub fn set_source_document_id(&mut self, value: Option<Guid>) -> Result<&mut Self> {
        self.next()?.set_source_document_id(value)?;
        Ok(self)
    }

    /// Set or remove a valueless `w15:docId`.
    ///
    /// # Errors
    ///
    /// Returns an error if the settings part cannot be read.
    pub fn set_source_document_id_without_value(&mut self, present: bool) -> Result<&mut Self> {
        self.next()?.set_source_document_id_without_value(present)?;
        Ok(self)
    }

    /// Set or remove `w14:conflictMode`.
    ///
    /// # Errors
    ///
    /// Returns an error if the settings part cannot be read or the value is
    /// invalid.
    pub fn set_conflict_mode(&mut self, value: Option<OnOff>) -> Result<&mut Self> {
        self.next()?.set_conflict_mode(value)?;
        Ok(self)
    }

    /// Set or remove `w14:discardImageEditingData`.
    ///
    /// # Errors
    ///
    /// Returns an error if the settings part cannot be read or the value is
    /// invalid.
    pub fn set_discard_image_editing_data(&mut self, value: Option<OnOff>) -> Result<&mut Self> {
        self.next()?.set_discard_image_editing_data(value)?;
        Ok(self)
    }

    /// Set or remove `w14:defaultImageDpi`.
    ///
    /// # Errors
    ///
    /// Returns an error if the settings part cannot be read or the value is
    /// invalid.
    pub fn set_default_image_dpi(&mut self, value: Option<i32>) -> Result<&mut Self> {
        self.next()?.set_default_image_dpi(value)?;
        Ok(self)
    }

    /// Discard every staged change, keeping the edit.
    pub fn reset(&mut self) {
        self.staged.reset();
        self.extensions = None;
    }

    /// Validate the staged extensions and publish them as a new package.
    /// Unchanged extensions publish the source package unchanged.
    ///
    /// # Errors
    ///
    /// Returns an error for invalid staged extensions or a settings part the
    /// rewrite cannot preserve.
    pub fn commit(mut self) -> Result<Commit> {
        if let Some((base, next)) = self.extensions.take() {
            next.validate()?;
            if &next != base.extensions() {
                let xml = codec::rewrite(base.xml_bytes(), &next)?;
                Snapshot::from_xml(xml.clone())?;
                let mut working = self.staged.working().clone();
                working.update_settings_part(xml)?;
                *self.staged.working_mut() = working;
            }
        }
        commit_staged(self.staged)
    }
}

impl crate::Package {
    /// The settings part's typed Word extensions; a package without a
    /// settings part reads the default settings.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid settings graph or part.
    pub fn settings_extensions(&self) -> Result<Snapshot> {
        Snapshot::from_xml(self.settings_part_snapshot()?.xml)
    }

    /// Stage a settings-extension change over this package. O(1): the edit
    /// shares the package until a verb reads or changes it.
    #[must_use]
    pub fn edit_settings_extensions(&self) -> Edit {
        Edit {
            staged: Staged::new(self),
            extensions: None,
        }
    }

    /// Replay a settings-extension patch onto this package, which must hold
    /// the patch's before state.
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::PatchConflict`] for a package that does not
    /// hold the before state.
    pub fn apply_settings_extensions(&self, patch: &Patch) -> Result<Commit> {
        apply_patch(self, patch, "apply_settings_extensions")
    }
}
