//! Clone-staged DataPilot transactions and contextual CRUD editors.

use std::borrow::Cow;

use crate::model::data_pilot::Table;
use crate::package::Package;
use litchi_core::{Error, Result};

use super::model::{Catalog, Selector, select};
use super::{codec::Location, package, validation};

/// An isolated mutable draft derived from an immutable DataPilot catalog.
pub struct Transaction<'source> {
    pub(crate) source: &'source Package,
    pub(crate) source_xml: &'source str,
    pub(crate) location: Location,
    pub(crate) original: Option<Vec<Table>>,
    pub(crate) draft: Option<Vec<Table>>,
}

impl<'source> Transaction<'source> {
    pub(crate) fn from_catalog(catalog: &Catalog<'source>) -> Self {
        let original = catalog.present.then(|| catalog.tables.clone());
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
    pub fn tables(&self) -> &[Table] {
        self.draft.as_deref().unwrap_or(&[])
    }

    /// Return whether the staged XML contains a physical owner.
    #[must_use]
    pub fn has_owner(&self) -> bool {
        self.draft.is_some()
    }

    /// Open a short-lived contextual editor.
    pub fn editor(&mut self) -> Editor<'_, 'source> {
        Editor { transaction: self }
    }

    /// Replace the complete ordered DataPilot catalog.
    pub fn replace(&mut self, tables: Vec<Table>) -> Result<()> {
        let candidate = Some(tables);
        validation::validate_candidate(&self.location, &self.original, &candidate)?;
        self.draft = candidate;
        Ok(())
    }

    /// Remove the physical DataPilot owner.
    pub fn remove(&mut self) -> Result<()> {
        let candidate = None;
        validation::validate_candidate(&self.location, &self.original, &candidate)?;
        self.draft = candidate;
        Ok(())
    }

    /// Publish the staged catalog as one package transaction.
    pub fn commit(self) -> Result<Commit<'source>> {
        validation::validate_candidate(&self.location, &self.original, &self.draft)?;
        if self.original == self.draft {
            return Ok(Commit {
                bytes: Cow::Borrowed(self.source.package().as_bytes()),
                tables: self.draft.unwrap_or_default(),
                has_owner: self.original.is_some(),
                changed: false,
            });
        }
        let bytes = package::replace(
            self.source,
            self.source_xml,
            &self.location,
            self.draft.as_deref(),
        )?;
        let has_owner = self.draft.is_some();
        Ok(Commit {
            bytes: Cow::Owned(bytes),
            tables: self.draft.unwrap_or_default(),
            has_owner,
            changed: true,
        })
    }
}

/// A short-lived semantic DataPilot editor.
pub struct Editor<'transaction, 'source> {
    pub(crate) transaction: &'transaction mut Transaction<'source>,
}

impl<'transaction, 'source> Editor<'transaction, 'source> {
    /// Borrow the current staged declarations.
    pub fn tables(&self) -> &[Table] {
        self.transaction.tables()
    }

    /// Add one validated declaration at the catalog tail.
    pub fn add(&mut self, table: Table) -> Result<()> {
        let mut candidate = self.transaction.draft.clone().unwrap_or_default();
        candidate.push(table);
        self.transaction.replace(candidate)
    }

    /// Replace one declaration selected by exact name or source position.
    pub fn replace<'a, S>(&mut self, selector: S, table: Table) -> Result<()>
    where
        S: Into<Selector<'a>>,
    {
        let mut candidate = self.transaction.draft.clone().unwrap_or_default();
        let index = select(&candidate, selector.into())?.ok_or_else(|| {
            Error::InvalidFormat("ODS DataPilot selector did not match".to_string())
        })?;
        candidate[index] = table;
        self.transaction.replace(candidate)
    }

    /// Apply a checked update to one selected declaration.
    pub fn update<'a, S, F>(&mut self, selector: S, update: F) -> Result<()>
    where
        S: Into<Selector<'a>>,
        F: FnOnce(&mut Table) -> Result<()>,
    {
        let mut candidate = self.transaction.draft.clone().unwrap_or_default();
        let index = select(&candidate, selector.into())?.ok_or_else(|| {
            Error::InvalidFormat("ODS DataPilot selector did not match".to_string())
        })?;
        update(&mut candidate[index])?;
        self.transaction.replace(candidate)
    }

    /// Remove one selected declaration and return the detached value.
    pub fn remove<'a, S>(&mut self, selector: S) -> Result<Table>
    where
        S: Into<Selector<'a>>,
    {
        let mut candidate = self.transaction.draft.clone().unwrap_or_default();
        let index = select(&candidate, selector.into())?.ok_or_else(|| {
            Error::InvalidFormat("ODS DataPilot selector did not match".to_string())
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
    pub fn clear(&mut self) -> Result<()> {
        self.transaction.remove()
    }
}

/// The result of publishing a DataPilot transaction.
pub struct Commit<'source> {
    bytes: Cow<'source, [u8]>,
    tables: Vec<Table>,
    has_owner: bool,
    changed: bool,
}

impl Commit<'_> {
    /// Borrow the resulting package bytes.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.bytes
    }

    /// Borrow the declarations represented by the commit.
    #[must_use]
    pub fn tables(&self) -> &[Table] {
        &self.tables
    }

    /// Return whether the physical package changed.
    #[must_use]
    pub const fn changed(&self) -> bool {
        self.changed
    }

    /// Return whether the resulting package has a physical owner.
    #[must_use]
    pub const fn has_owner(&self) -> bool {
        self.has_owner
    }
}

impl<'source> Commit<'source> {
    /// Consume the commit while retaining a borrow for a no-op result.
    pub fn into_bytes(self) -> Cow<'source, [u8]> {
        self.bytes
    }

    /// Consume the commit into package-owned bytes.
    pub fn into_owned_bytes(self) -> Vec<u8> {
        self.bytes.into_owned()
    }
}
