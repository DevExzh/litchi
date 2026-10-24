//! Package-level font edits (record 0357).
//!
//! [`Package::edit_fonts`] stages edits of the live document's font
//! collections, [`Edit::commit`] publishes a new [`Package`] through one
//! append-only user edit, and [`Package::apply_fonts`] replays a
//! [`Patch`] onto the exact package it was made from. Dropping an [`Edit`]
//! discards it.
//!
//! ```
//! use litchi_ppt::font::{FontEmbeddingFlags, Scope};
//!
//! let package = litchi_ppt::Package::new()?;
//! let mut edit = package.edit_fonts()?;
//! let font = edit.fonts()?.collection(Scope::Base).unwrap().get(0).unwrap().clone();
//! edit.replace_font(Scope::Base, 0, font)?;
//! let commit = edit.commit()?;
//! assert!(!commit.diagnostics().changed());
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```

use litchi_core::edit::{Refusal, RefusalKind};

use super::transaction::Transaction as RecordEdit;
use super::{
    Change, Commit, Error, Facet, Font, FontCollections, FontEmbeddingFlags, PackageLimits, Patch,
    Result, Scope, SharedFontData, Snapshot,
};
use crate::Package;
use crate::scoped::{self, Base, Failure, Lazy};

/// A staged font edit over one immutable [`Package`].
#[derive(Debug)]
pub struct Edit {
    base: Base,
    inner: Lazy<RecordEdit>,
}

impl Edit {
    fn new(package: &Package) -> Result<Self> {
        let base = Base::capture(package)?;
        Ok(Self {
            inner: Lazy::new(),
            base,
        })
    }

    fn initialize_inner(&self) -> std::result::Result<RecordEdit, Failure> {
        let snapshot = Snapshot::from_arc(self.base.shared_bytes()?, PackageLimits::default())
            .map_err(scoped::publication("edit_fonts"))?;
        Ok(RecordEdit::new(snapshot))
    }

    fn inner(&self) -> std::result::Result<&RecordEdit, Failure> {
        self.inner.get_or_try_init(|| self.initialize_inner())
    }

    fn inner_mut(&mut self) -> std::result::Result<&mut RecordEdit, Failure> {
        let base = self.base.clone();
        self.inner.get_or_try_init_mut(|| {
            let snapshot = Snapshot::from_arc(base.shared_bytes()?, PackageLimits::default())
                .map_err(scoped::publication("edit_fonts"))?;
            Ok::<RecordEdit, Failure>(RecordEdit::new(snapshot))
        })
    }

    /// The staged font collections.
    #[must_use = "handle the fallible font projection result"]
    pub fn fonts(&self) -> Result<&FontCollections> {
        Ok(self.inner()?.fonts())
    }

    /// The staged operations, in order.
    #[must_use = "handle the fallible font changes result"]
    pub fn changes(&self) -> Result<&[Change]> {
        Ok(self.inner()?.changes())
    }

    /// Stage a replacement of the font at `index` in the collection for
    /// `scope`. An equal font is an exact no-op.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Refused`] when the collection or ordinal is absent or
    /// the font cannot be stored, leaving the edit unchanged.
    pub fn replace_font(&mut self, scope: Scope, index: u16, font: Font) -> Result<()> {
        Ok(self
            .inner_mut()?
            .replace_font(scope, index, font)
            .map_err(scoped::staging("font.replace_font"))?)
    }

    /// Stage an append of `font` to the collection for `scope`, returning its
    /// new ordinal.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Refused`] when the collection is absent or full, or
    /// the font cannot be stored.
    pub fn append_font(&mut self, scope: Scope, font: Font) -> Result<u16> {
        Ok(self
            .inner_mut()?
            .append_font(scope, font)
            .map_err(scoped::staging("font.append_font"))?)
    }

    /// Stage a validated EOT facet on the font at `index` in the collection
    /// for `scope`.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Refused`] when the collection or ordinal is absent or
    /// `data` fails EOT validation.
    pub fn set_facet(
        &mut self,
        scope: Scope,
        index: u16,
        facet: Facet,
        data: impl Into<SharedFontData>,
    ) -> Result<()> {
        Ok(self
            .inner_mut()?
            .set_facet(scope, index, facet, data)
            .map_err(scoped::staging("font.set_facet"))?)
    }

    /// Prepare and stage one inert EOT facet without loading or executing
    /// it; the facet follows `font.style`.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Refused`] when the collection or ordinal is absent or
    /// the prepared font cannot be encoded under `limits`.
    #[cfg(feature = "fonts")]
    pub fn set_prepared_facet(
        &mut self,
        scope: Scope,
        index: u16,
        font: &mut super::PreparedFont,
        intent: super::EotIntent,
        limits: super::EotLimits,
    ) -> Result<()> {
        Ok(self
            .inner_mut()?
            .set_prepared_facet(scope, index, font, intent, limits)
            .map_err(scoped::staging("font.set_prepared_facet"))?)
    }

    /// Stage removal of `facet` from the font at `index` in the collection
    /// for `scope`, returning whether a facet was removed.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Refused`] when the collection or ordinal is absent.
    pub fn remove_facet(&mut self, scope: Scope, index: u16, facet: Facet) -> Result<bool> {
        Ok(self
            .inner_mut()?
            .remove_facet(scope, index, facet)
            .map_err(scoped::staging("font.remove_facet"))?)
    }

    /// Stage document-wide embedding flags.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Refused`] when the flags cannot be stored.
    pub fn set_embedding_flags(&mut self, flags: Option<FontEmbeddingFlags>) -> Result<()> {
        Ok(self
            .inner_mut()?
            .set_embedding_flags(flags)
            .map_err(scoped::staging("font.set_embedding_flags"))?)
    }

    /// Font removal is refused until every base and `PowerPoint` 10
    /// reference can be proven and remapped.
    ///
    /// # Errors
    ///
    /// Always returns [`Error::Refused`] with
    /// [`RefusalKind::UnsupportedOperation`].
    pub fn remove_font(&mut self, _scope: Scope, _index: u16) -> Result<Font> {
        Err(Error::Refused(Refusal::new(
            RefusalKind::UnsupportedOperation,
            "font.remove_font",
            "font removal requires a complete base and PP10 reference remap",
        )))
    }

    /// Font reordering is refused until every base and `PowerPoint` 10
    /// reference can be proven and remapped.
    ///
    /// # Errors
    ///
    /// Always returns [`Error::Refused`] with
    /// [`RefusalKind::UnsupportedOperation`].
    pub fn reorder_fonts(&mut self, _scope: Scope, _order: &[u16]) -> Result<()> {
        Err(Error::Refused(Refusal::new(
            RefusalKind::UnsupportedOperation,
            "font.reorder_fonts",
            "font reordering requires a complete base and PP10 reference remap",
        )))
    }

    /// Publish the staged fonts as a new package. An edit with no net change
    /// publishes the source package unchanged.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Signed`] for a changed write to a signed package,
    /// [`Error::Refused`] when the package holds storage the publisher cannot
    /// preserve, and [`Error::Package`] when the result cannot be
    /// republished.
    pub fn commit(self) -> Result<Commit> {
        let Self { base, inner } = self;
        inner
            .get_or_try_init(|| {
                let snapshot = Snapshot::from_arc(base.shared_bytes()?, PackageLimits::default())
                    .map_err(scoped::publication("edit_fonts"))?;
                Ok::<RecordEdit, Failure>(RecordEdit::new(snapshot))
            })
            .map_err(Error::from)?;
        let inner = inner.into_inner().ok_or_else(|| {
            Failure::Package(crate::package::Error::Corrupted(
                "PPT font edit state was not initialized".into(),
            ))
        })?;
        let after = match inner.commit() {
            Ok(commit) => commit.into_parts().0.bytes().to_vec(),
            // The publisher refuses the signature storages a changed write
            // would invalidate; report the signature, not the storage.
            Err(_) if base.is_signed()? => return Err(Error::Signed),
            Err(error) => return Err(scoped::publication("font.commit")(error).into()),
        };
        Ok(super::published(base.publish(after)?))
    }
}

impl Package {
    /// Start a font edit over this package.
    ///
    /// # Errors
    ///
    /// An encrypted presentation is refused immediately. A live font owner
    /// that cannot be read or losslessly represented is reported by the first
    /// projection, staging operation, or commit.
    pub fn edit_fonts(&self) -> Result<Edit> {
        Edit::new(self)
    }

    /// Apply a font patch to the exact package it was made from.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// source.
    pub fn apply_fonts(&self, patch: &Patch) -> Result<Commit> {
        super::apply_patch(self, patch)
    }
}
