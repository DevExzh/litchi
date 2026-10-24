//! Immutable timeline edits of an XLSB [`Package`] (record 0360b).
//!
//! [`Package::edit_timelines`] starts an [`Edit`] over the workbook timeline
//! caches and every worksheet's timeline views; its verbs replace, upsert or
//! remove caches and views, each change validated against the timeline
//! dependency closure (every view names an existing cache) and read back.
//! A change that stores timeline content refuses a source whose stored XML
//! is not canonical: its lexical or unmodeled content is preserved only by
//! an exact no-op. Removing every cache or view is always allowed.
//! [`Edit::commit`] returns a [`Commit`] with the new package, the
//! reversible [`Patch`] and the [`Diagnostics`]; [`Package::apply_timelines`]
//! replays a patch onto an equal package. Content equal to the stored
//! content publishes nothing, so an exact no-op keeps every byte, and a
//! changed edit of a signed workbook is refused with [`Error::Signed`].
//!
//! ```
//! use litchi_xlsb::Package;
//! use litchi_xlsb::timeline::{Cache, FilterType, Range, State};
//!
//! # fn main() -> Result<(), litchi_xlsb::package::PackageError> {
//! let package = Package::new()?;
//! let bounds = Range::new("2024-01-01T00:00:00Z", "2024-12-31T23:59:59Z")?;
//! let cache = Cache::new("TimelineCache1", "Date", State::new(bounds, 0, FilterType::DateBetween));
//! let mut edit = package.edit_timelines()?;
//! assert!(edit.set_caches(vec![cache])?);
//! let (after, patch) = edit.commit()?.into_parts();
//! assert_eq!(after.workbook()?.timeline_caches()?.len(), 1);
//! let restored = after.apply_timelines(&patch.inverse())?.into_package();
//! assert_eq!(restored.to_bytes()?, package.to_bytes()?);
//! # Ok(())
//! # }
//! ```

use litchi_opc::PackURI;

use super::transaction::{
    Content, Scope, inspect, store, validate_caches, validate_closure, validate_content,
};
use super::{Cache, View, Views};
use crate::package::Package;
use crate::package::error::Result;
use crate::part_edit::{Staged, family};

family!("timeline");

/// A bounded dimension of this family.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Timeline caches of one workbook.
    Caches,
    /// Timeline views of one worksheet.
    Views,
}

/// A staged timeline edit. Dropping it discards it.
#[derive(Debug, Clone)]
pub struct Edit {
    staged: Staged,
}

impl Package {
    /// Start a timeline edit. Starting one copies nothing and defers workbook
    /// projection validation until the first operation or commit that needs
    /// the projection.
    ///
    /// # Errors
    ///
    /// Returns a cold edit. Workbook validation errors are reported by the
    /// first fallible operation or by [`Edit::commit`] after the projection is
    /// needed.
    pub fn edit_timelines(&self) -> Result<Edit> {
        Ok(Edit {
            staged: Staged::of_package(self)?,
        })
    }

    /// Apply a timeline patch to a new package; this one is unchanged.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state and [`Error::Signed`] for a changed patch of a signed
    /// workbook.
    pub fn apply_timelines(&self, patch: &Patch) -> Result<Commit> {
        Ok(Commit::new((patch.0.apply(self)?, patch.0.clone())))
    }
}

impl Edit {
    fn workbook_uri(&self) -> Result<PackURI> {
        Ok(self.staged.opc().main_document_part()?.partname().clone())
    }

    /// Discard every staged change.
    pub fn reset(&mut self) {
        self.staged.reset();
    }

    /// The staged workbook timeline caches, in relationship order.
    ///
    /// # Errors
    ///
    /// Returns the timeline graph and decoding errors.
    pub fn caches(&self) -> Result<Vec<Cache>> {
        match inspect(
            self.staged.opc(),
            &Scope::Caches(self.workbook_uri()?),
            self.staged.workbook()?.semantic_limits().xml_parts_value(),
        )?
        .0
        {
            Content::Caches(caches) => Ok(caches),
            Content::Views(_) => Ok(Vec::new()),
        }
    }

    /// The staged timeline views of the worksheet at complete-catalog position
    /// `sheet`; empty when it has none.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing tab, a non-worksheet tab and the timeline
    /// graph and decoding errors.
    pub fn views(&self, sheet: usize) -> Result<Views> {
        match inspect(
            self.staged.opc(),
            &Scope::Views(self.staged.worksheet(sheet)?),
            self.staged.workbook()?.semantic_limits().xml_parts_value(),
        )?
        .0
        {
            Content::Views(views) => Ok(views),
            Content::Caches(_) => Ok(Views::new()),
        }
    }

    fn stage(&mut self, scope: Scope, content: Content) -> Result<bool> {
        let xml_limits = self.staged.workbook()?.semantic_limits().xml_parts_value();
        validate_content(&content, xml_limits)?;
        let (stored, canonical) = inspect(self.staged.opc(), &scope, xml_limits)?;
        if stored == content {
            return Ok(false);
        }
        let empty = match &content {
            Content::Caches(caches) => caches.is_empty(),
            Content::Views(views) => views.items.is_empty(),
        };
        if !canonical && !empty {
            return Err(crate::package::PackageError::refused(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                "edit_timelines",
                "the stored timeline XML is not canonical",
            ));
        }
        crate::signature_policy::ensure_unsigned(self.staged.opc())?;
        self.staged.update(|opc| {
            store(opc, &scope, &content, xml_limits)?;
            if inspect(opc, &scope, xml_limits)?.0 != content {
                return Err(crate::part_edit::invalid(
                    "timeline readback did not match the staged content",
                ));
            }
            validate_closure(opc, &scope, &content, xml_limits)
        })?;
        Ok(true)
    }

    /// Replace every workbook timeline cache; an empty list removes them.
    /// Returns whether anything changed.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid or duplicate cache, a cache a view
    /// still names, a shared target, and a changed edit of a signed
    /// workbook.
    pub fn set_caches(&mut self, caches: Vec<Cache>) -> Result<bool> {
        validate_caches(
            &caches,
            self.staged.workbook()?.semantic_limits().xml_parts_value(),
        )?;
        let scope = Scope::Caches(self.workbook_uri()?);
        self.stage(scope, Content::Caches(caches))
    }

    /// Add or replace the cache with `cache`'s case-insensitive name.
    /// Returns whether anything changed.
    ///
    /// # Errors
    ///
    /// The errors of [`Self::set_caches`].
    pub fn upsert_cache(&mut self, cache: Cache) -> Result<bool> {
        let mut caches = self.caches()?;
        match caches
            .iter()
            .position(|value| value.name.eq_ignore_ascii_case(&cache.name))
        {
            Some(index) => caches[index] = cache,
            None => caches.push(cache),
        }
        self.set_caches(caches)
    }

    /// Remove the cache with this case-insensitive name and return it.
    ///
    /// # Errors
    ///
    /// The errors of [`Self::set_caches`].
    pub fn remove_cache(&mut self, name: &str) -> Result<Option<Cache>> {
        let mut caches = self.caches()?;
        let Some(index) = caches
            .iter()
            .position(|value| value.name.eq_ignore_ascii_case(name))
        else {
            return Ok(None);
        };
        let removed = caches.remove(index);
        self.set_caches(caches)?;
        Ok(Some(removed))
    }

    /// Remove every workbook timeline cache. Returns whether any were present.
    ///
    /// # Errors
    ///
    /// The errors of [`Self::set_caches`].
    pub fn remove_caches(&mut self) -> Result<bool> {
        self.set_caches(Vec::new())
    }

    /// Replace the timeline views of the worksheet at complete-catalog position
    /// `sheet`; empty views remove its timeline part. Returns whether anything
    /// changed.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing or non-worksheet tab, invalid views, a
    /// view naming a missing cache, a shared target, and a changed edit of a
    /// signed workbook.
    pub fn set_views(&mut self, sheet: usize, views: Views) -> Result<bool> {
        let scope = Scope::Views(self.staged.worksheet(sheet)?);
        self.stage(scope, Content::Views(views))
    }

    /// Add or replace the view with `view`'s case-insensitive name on a
    /// worksheet. Returns whether anything changed.
    ///
    /// # Errors
    ///
    /// The errors of [`Self::set_views`].
    pub fn upsert_view(&mut self, sheet: usize, view: View) -> Result<bool> {
        let mut views = self.views(sheet)?;
        match views
            .items
            .iter()
            .position(|value| value.name.eq_ignore_ascii_case(&view.name))
        {
            Some(index) => views.items[index] = view,
            None => views.items.push(view),
        }
        self.set_views(sheet, views)
    }

    /// Remove the view with this case-insensitive name from a worksheet and
    /// return it.
    ///
    /// # Errors
    ///
    /// The errors of [`Self::set_views`].
    pub fn remove_view(&mut self, sheet: usize, name: &str) -> Result<Option<View>> {
        let mut views = self.views(sheet)?;
        let Some(index) = views
            .items
            .iter()
            .position(|value| value.name.eq_ignore_ascii_case(name))
        else {
            return Ok(None);
        };
        let removed = views.items.remove(index);
        self.set_views(sheet, views)?;
        Ok(Some(removed))
    }

    /// Remove every timeline view of a worksheet. Returns whether a timeline
    /// part was present.
    ///
    /// # Errors
    ///
    /// The errors of [`Self::set_views`].
    pub fn remove_views(&mut self, sheet: usize) -> Result<bool> {
        let worksheet = self.staged.worksheet(sheet)?;
        let present = super::package::load_views(
            self.staged.opc(),
            &worksheet,
            self.staged.workbook()?.semantic_limits().xml_parts_value(),
        )?
        .is_some();
        self.set_views(sheet, Views::new())?;
        Ok(present)
    }

    /// Validate the staged timeline graph and publish it to a new package. An
    /// unchanged staging returns the source package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Signed`] for a changed staging of a signed workbook
    /// and the workbook validation errors of the staged package.
    pub fn commit(self) -> Result<Commit> {
        Ok(Commit::new(self.staged.commit()?))
    }
}
