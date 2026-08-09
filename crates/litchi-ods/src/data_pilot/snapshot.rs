//! Owned DataPilot snapshots, edits, commits, and reversible patches.

use super::{Catalog, Selector, validation};
use crate::{model::data_pilot::Table, package::Package};
use litchi_core::{Error, Result};

/// Immutable DataPilot owner bound to one complete ODS package source.
#[derive(Clone, Debug)]
pub struct Snapshot {
    source: Vec<u8>,
    tables: Vec<Table>,
    has_owner: bool,
}

impl Snapshot {
    /// Parse the DataPilot owner from an owned ODS package snapshot.
    pub fn from_bytes(source: Vec<u8>) -> Result<Self> {
        let package = Package::from_bytes(source.clone())?;
        let catalog = Catalog::load(&package)?;
        Ok(Self {
            source,
            tables: catalog.tables.clone(),
            has_owner: catalog.present,
        })
    }

    /// Exact ODS package bytes captured by this snapshot.
    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        &self.source
    }

    /// Typed declarations in source order.
    #[must_use]
    pub fn tables(&self) -> &[Table] {
        &self.tables
    }

    /// Whether `table:data-pilot-tables` is physically present.
    #[must_use]
    pub const fn has_owner(&self) -> bool {
        self.has_owner
    }

    /// Resolve one declaration by exact name or checked source position.
    pub fn get<'a, S>(&self, selector: S) -> Result<Option<&Table>>
    where
        S: Into<Selector<'a>>,
    {
        super::model::select(&self.tables, selector.into())
            .map(|index| index.map(|index| &self.tables[index]))
    }

    /// Start a source-checked, failure-atomic edit.
    #[must_use]
    pub fn edit(&self) -> Edit {
        Edit {
            before: self.clone(),
            draft: self.has_owner.then(|| self.tables.clone()),
        }
    }
}

/// A staged DataPilot edit derived from one immutable [`Snapshot`].
#[derive(Clone, Debug)]
pub struct Edit {
    before: Snapshot,
    draft: Option<Vec<Table>>,
}

impl Edit {
    /// Typed declarations in the current candidate.
    #[must_use]
    pub fn tables(&self) -> &[Table] {
        self.draft.as_deref().unwrap_or(&[])
    }

    /// Whether the current candidate retains the physical owner.
    #[must_use]
    pub const fn has_owner(&self) -> bool {
        self.draft.is_some()
    }

    /// Replace the complete ordered DataPilot catalog.
    pub fn replace(&mut self, tables: Vec<Table>) -> Result<()> {
        self.stage(Some(tables))
    }

    /// Remove the physical DataPilot owner.
    pub fn remove(&mut self) -> Result<()> {
        self.stage(None)
    }

    /// Open the existing short-lived contextual CRUD editor.
    pub fn editor(&mut self) -> OwnedEditor<'_> {
        OwnedEditor { edit: self }
    }

    /// Restore the exact source candidate.
    pub fn rollback(&mut self) {
        self.draft = self.before.has_owner.then(|| self.before.tables.clone());
    }

    /// Validate, rehydrate, and atomically publish one immutable snapshot.
    pub fn commit(self) -> Result<Commit> {
        validation::validate_candidate(
            &self.before_location()?,
            &self.before_owner(),
            &self.draft,
        )?;
        let unchanged = self.before_owner() == self.draft;
        let target = if unchanged {
            self.before.clone()
        } else {
            let package = Package::from_bytes(self.before.source.clone())?;
            let catalog = Catalog::load(&package)?;
            let bytes = super::package::replace(
                &package,
                catalog.source_xml,
                &catalog.location,
                self.draft.as_deref(),
            )?;
            let target = Snapshot::from_bytes(bytes)?;
            if target.has_owner != self.draft.is_some()
                || target.tables != self.draft.unwrap_or_default()
            {
                return Err(Error::InvalidFormat(
                    "ODS DataPilot commit failed typed readback".to_string(),
                ));
            }
            target
        };
        Ok(Commit {
            changed: !unchanged,
            patch: Patch {
                source: self.before.source,
                target: target.source.clone(),
            },
            snapshot: target,
        })
    }

    fn before_owner(&self) -> Option<Vec<Table>> {
        self.before.has_owner.then(|| self.before.tables.clone())
    }

    fn before_location(&self) -> Result<super::codec::Location> {
        let package = Package::from_bytes(self.before.source.clone())?;
        Ok(Catalog::load(&package)?.location)
    }

    fn stage(&mut self, candidate: Option<Vec<Table>>) -> Result<()> {
        validation::validate_candidate(&self.before_location()?, &self.before_owner(), &candidate)?;
        self.draft = candidate;
        Ok(())
    }
}

/// Short-lived semantic verbs over an owned [`Edit`].
pub struct OwnedEditor<'edit> {
    edit: &'edit mut Edit,
}

impl OwnedEditor<'_> {
    /// Add one declaration at the catalog tail.
    pub fn add(&mut self, table: Table) -> Result<()> {
        let mut candidate = self.edit.draft.clone().unwrap_or_default();
        candidate.push(table);
        self.edit.replace(candidate)
    }

    /// Replace one declaration selected by exact name or checked source position.
    pub fn replace<'a, S>(&mut self, selector: S, table: Table) -> Result<()>
    where
        S: Into<Selector<'a>>,
    {
        let mut candidate = self.edit.draft.clone().unwrap_or_default();
        let index = super::model::select(&candidate, selector.into())?.ok_or_else(|| {
            Error::InvalidFormat("ODS DataPilot selector did not match".to_string())
        })?;
        candidate[index] = table;
        self.edit.replace(candidate)
    }

    /// Apply one checked update to a selected declaration.
    pub fn update<'a, S, F>(&mut self, selector: S, update: F) -> Result<()>
    where
        S: Into<Selector<'a>>,
        F: FnOnce(&mut Table) -> Result<()>,
    {
        let mut candidate = self.edit.draft.clone().unwrap_or_default();
        let index = super::model::select(&candidate, selector.into())?.ok_or_else(|| {
            Error::InvalidFormat("ODS DataPilot selector did not match".to_string())
        })?;
        update(&mut candidate[index])?;
        self.edit.replace(candidate)
    }

    /// Remove one selected declaration and return the detached semantic value.
    pub fn remove<'a, S>(&mut self, selector: S) -> Result<Table>
    where
        S: Into<Selector<'a>>,
    {
        let mut candidate = self.edit.draft.clone().unwrap_or_default();
        let index = super::model::select(&candidate, selector.into())?.ok_or_else(|| {
            Error::InvalidFormat("ODS DataPilot selector did not match".to_string())
        })?;
        let removed = candidate.remove(index);
        if candidate.is_empty() {
            self.edit.remove()?;
        } else {
            self.edit.replace(candidate)?;
        }
        Ok(removed)
    }
}

/// A source-checked, reversible package patch.
#[derive(Clone, Debug)]
pub struct Patch {
    source: Vec<u8>,
    target: Vec<u8>,
}

impl Patch {
    /// Whether this patch is physically empty.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.source == self.target
    }

    /// Return the exact-source patch that restores the accepted package.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            source: self.target.clone(),
            target: self.source.clone(),
        }
    }

    /// Apply only to the exact source package snapshot.
    pub fn apply(&self, snapshot: &Snapshot) -> Result<Commit> {
        if snapshot.source != self.source {
            return Err(Error::InvalidFormat(
                "ODS DataPilot patch source snapshot does not match".to_string(),
            ));
        }
        let target = Snapshot::from_bytes(self.target.clone())?;
        Ok(Commit {
            changed: !self.is_empty(),
            patch: self.clone(),
            snapshot: target,
        })
    }
}

/// A fully rehydrated DataPilot publication.
#[derive(Clone, Debug)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
    changed: bool,
}

impl Commit {
    /// Whether package bytes changed.
    #[must_use]
    pub const fn changed(&self) -> bool {
        self.changed
    }

    /// Resulting immutable snapshot.
    #[must_use]
    pub fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Reversible exact-source patch.
    #[must_use]
    pub fn patch(&self) -> &Patch {
        &self.patch
    }
}
