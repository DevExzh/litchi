//! Clone-staged `database-ranges` transactions and contextual CRUD editors.

use std::borrow::Cow;

use crate::model::database_range::Range;
use crate::package::Archive;
use litchi_core::{Error, Result};

use super::model::{Catalog, Selector, select};
use super::{codec::Location, package, validation};

/// An isolated mutable draft derived from an immutable `database-ranges` catalog.
pub(crate) struct Transaction<'source> {
    pub(crate) source: &'source Archive,
    pub(crate) source_xml: &'source str,
    pub(crate) location: Location,
    pub(crate) original: Option<Vec<Range>>,
    pub(crate) draft: Option<Vec<Range>>,
}

impl<'source> Transaction<'source> {
    pub(crate) fn from_catalog(catalog: &Catalog<'source>) -> Self {
        let original = catalog.present.then(|| catalog.ranges.clone());
        Self {
            source: catalog.source,
            source_xml: catalog.source_xml,
            location: catalog.location.clone(),
            draft: original.clone(),
            original,
        }
    }

    /// Borrow the current staged declarations.  An absent owner is an empty
    /// semantic catalog, while [`Self::has_owner`] retains physical presence.
    #[must_use]
    pub(crate) fn ranges(&self) -> &[Range] {
        self.draft.as_deref().unwrap_or(&[])
    }

    /// Open a short-lived contextual editor.
    pub(crate) fn editor(&mut self) -> Editor<'_, 'source> {
        Editor { transaction: self }
    }

    /// Replace the complete ordered `database-ranges` catalog.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub(crate) fn replace(&mut self, ranges: Vec<Range>) -> Result<()> {
        let candidate = Some(ranges);
        validation::validate_candidate(&self.location, &self.original, &candidate)?;
        self.draft = candidate;
        Ok(())
    }

    /// Remove the physical `database-ranges` owner.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub(crate) fn remove(&mut self) -> Result<()> {
        let candidate = None;
        validation::validate_candidate(&self.location, &self.original, &candidate)?;
        self.draft = candidate;
        Ok(())
    }

    /// Publish the staged catalog as one package transaction.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub(crate) fn commit(self) -> Result<Commit<'source>> {
        validation::validate_candidate(&self.location, &self.original, &self.draft)?;
        if self.original == self.draft {
            return Ok(Commit {
                bytes: Cow::Borrowed(self.source.package().as_bytes()),
                changed: false,
            });
        }
        let bytes = package::replace(
            self.source,
            self.source_xml,
            &self.location,
            self.draft.as_deref(),
        )?;
        Ok(Commit {
            bytes: Cow::Owned(bytes),
            changed: true,
        })
    }
}

/// A short-lived semantic `database-ranges` editor.
pub struct Editor<'transaction, 'source> {
    pub(crate) transaction: &'transaction mut Transaction<'source>,
}

impl Editor<'_, '_> {
    /// Borrow the current staged declarations.
    #[must_use]
    pub fn ranges(&self) -> &[Range] {
        self.transaction.ranges()
    }

    /// Add one validated declaration at the catalog tail.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn add(&mut self, range: Range) -> Result<()> {
        let mut candidate = self.transaction.draft.clone().unwrap_or_default();
        candidate.push(range);
        self.transaction.replace(candidate)
    }

    /// Replace one declaration selected by exact name or source position.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn replace<'a, S>(&mut self, selector: S, range: Range) -> Result<()>
    where
        S: Into<Selector<'a>>,
    {
        let mut candidate = self.transaction.draft.clone().unwrap_or_default();
        let index = select(&candidate, selector.into())?.ok_or_else(|| {
            Error::InvalidFormat("ODS database-ranges selector did not match".to_string())
        })?;
        candidate[index] = range;
        self.transaction.replace(candidate)
    }

    /// Apply a checked update to one selected declaration.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn update<'a, S, F>(&mut self, selector: S, update: F) -> Result<()>
    where
        S: Into<Selector<'a>>,
        F: FnOnce(&mut Range) -> Result<()>,
    {
        let mut candidate = self.transaction.draft.clone().unwrap_or_default();
        let index = select(&candidate, selector.into())?.ok_or_else(|| {
            Error::InvalidFormat("ODS database-ranges selector did not match".to_string())
        })?;
        update(&mut candidate[index])?;
        self.transaction.replace(candidate)
    }

    /// Remove one selected declaration and return the detached value.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn remove<'a, S>(&mut self, selector: S) -> Result<Range>
    where
        S: Into<Selector<'a>>,
    {
        let mut candidate = self.transaction.draft.clone().unwrap_or_default();
        let index = select(&candidate, selector.into())?.ok_or_else(|| {
            Error::InvalidFormat("ODS database-ranges selector did not match".to_string())
        })?;
        let removed = candidate.remove(index);
        if candidate.is_empty() {
            self.transaction.remove()?;
        } else {
            self.transaction.replace(candidate)?;
        }
        Ok(removed)
    }

    /// Remove all declarations and the physical owner.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn clear(&mut self) -> Result<()> {
        self.transaction.remove()
    }
}

/// The result of publishing a `database-ranges` transaction.
pub(crate) struct Commit<'source> {
    bytes: Cow<'source, [u8]>,
    changed: bool,
}

impl Commit<'_> {
    #[cfg(test)]
    /// Borrow the resulting package bytes.
    #[must_use]
    pub(crate) fn bytes(&self) -> &[u8] {
        &self.bytes
    }

    /// Return whether the physical package changed.
    #[must_use]
    pub(crate) const fn changed(&self) -> bool {
        self.changed
    }
}

impl<'source> Commit<'source> {
    /// Consume the commit into package-owned bytes.
    #[must_use]
    pub(crate) fn into_owned_bytes(self) -> Vec<u8> {
        self.bytes.into_owned()
    }
}
