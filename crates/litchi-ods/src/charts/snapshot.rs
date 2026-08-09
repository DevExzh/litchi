//! Owned chart snapshots, transactional edits, and reversible package patches.

use super::{Limits, Part, Selector, inventory, model::Chart, package};
use crate::package::Package;
use litchi_core::{Error, Result};

/// Immutable embedded-chart inventory tied to exact ODS package bytes.
#[derive(Clone, Debug)]
pub struct Snapshot {
    source: Vec<u8>,
    limits: Limits,
    charts: Vec<Chart>,
}

impl Snapshot {
    /// Parse an owned ODS package under the supplied chart resource budget.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn from_bytes(source: Vec<u8>) -> Result<Self> {
        Self::from_bytes_with(source, Limits::default())
    }

    /// Parse an owned ODS package with an explicit resource budget.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn from_bytes_with(source: Vec<u8>, limits: Limits) -> Result<Self> {
        let package = Package::from_bytes(source.clone())?;
        let charts = inventory(&package, limits)?.charts;
        Ok(Self {
            source,
            limits,
            charts,
        })
    }

    /// Exact ODS package bytes captured by this snapshot.
    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        &self.source
    }

    /// Chart resource budget used for this snapshot.
    #[must_use]
    pub const fn limits(&self) -> Limits {
        self.limits
    }

    /// Embedded charts in drawing discovery order.
    #[must_use]
    pub fn charts(&self) -> &[Chart] {
        &self.charts
    }

    /// Select a chart by exact drawing name or checked discovery position.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn get<'a, S>(&self, selector: S) -> Result<Option<&Chart>>
    where
        S: Into<Selector<'a>>,
    {
        select(&self.charts, selector.into()).map(|index| index.map(|index| &self.charts[index]))
    }

    /// Start a source-checked, failure-atomic edit.
    #[must_use]
    pub fn edit(&self) -> Edit {
        Edit {
            before: self.clone(),
            draft: self.charts.clone(),
        }
    }
}

/// Clone-staged embedded-chart edit.
#[derive(Clone, Debug)]
pub struct Edit {
    before: Snapshot,
    draft: Vec<Chart>,
}

impl Edit {
    /// Current candidate chart inventory.
    #[must_use]
    pub fn charts(&self) -> &[Chart] {
        &self.draft
    }

    /// Replace one selected chart part and every occurrence sharing its
    /// package-backed chart owner.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn replace<'a, S>(&mut self, selector: S, part: Part) -> Result<()>
    where
        S: Into<Selector<'a>>,
    {
        if part.xml().len() > self.before.limits.max_part_bytes() {
            return Err(Error::InvalidFormat(
                "ODS replacement chart exceeds the part-byte limit".to_string(),
            ));
        }
        let index = select(&self.draft, selector.into())?.ok_or_else(|| {
            Error::InvalidFormat("ODS embedded chart selector did not match".to_string())
        })?;
        let selected = self.draft[index].clone();
        for chart in &mut self.draft {
            if chart == &selected || chart.shares_storage(&selected) {
                *chart = chart.with_part(part.clone());
            }
        }
        Ok(())
    }

    /// Restore the exact source candidate.
    pub fn rollback(&mut self) {
        self.draft = self.before.charts.clone();
    }

    /// Validate package topology, reparse the candidate, and atomically publish it.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn commit(self) -> Result<Commit> {
        let changed = self.before.charts != self.draft;
        let snapshot = if changed {
            let package_source = Package::from_bytes(self.before.source.clone())?;
            let bytes = package::replace(&package_source, &self.before.charts, &self.draft)?;
            let result = Snapshot::from_bytes_with(bytes, self.before.limits)?;
            if result.charts != self.draft {
                return Err(Error::InvalidFormat(
                    "ODS chart commit failed typed readback".to_string(),
                ));
            }
            result
        } else {
            self.before.clone()
        };
        Ok(Commit {
            changed,
            patch: Patch {
                source: self.before.source,
                target: snapshot.source.clone(),
                limits: snapshot.limits,
            },
            snapshot,
        })
    }
}

/// Reversible patch guarded by exact complete-package source bytes.
#[derive(Clone, Debug)]
pub struct Patch {
    source: Vec<u8>,
    target: Vec<u8>,
    limits: Limits,
}

impl Patch {
    /// Whether this patch is physically empty.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.source == self.target
    }

    /// Return the patch restoring the exact accepted source package.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            source: self.target.clone(),
            target: self.source.clone(),
            limits: self.limits,
        }
    }

    /// Apply this patch only to its exact source snapshot.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn apply(&self, snapshot: &Snapshot) -> Result<Commit> {
        if snapshot.source != self.source {
            return Err(Error::InvalidFormat(
                "ODS chart patch source snapshot does not match".to_string(),
            ));
        }
        let target = Snapshot::from_bytes_with(self.target.clone(), self.limits)?;
        Ok(Commit {
            changed: !self.is_empty(),
            patch: self.clone(),
            snapshot: target,
        })
    }
}

/// A fully rehydrated chart publication.
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

fn select(charts: &[Chart], selector: Selector<'_>) -> Result<Option<usize>> {
    match selector {
        Selector::Index(index) => Ok((index < charts.len()).then_some(index)),
        Selector::Name(name) => {
            let mut selected = None;
            for (index, chart) in charts.iter().enumerate() {
                if chart.name() == Some(name) {
                    if selected.is_some() {
                        return Err(Error::InvalidFormat(format!(
                            "ODS embedded chart name '{name}' is ambiguous"
                        )));
                    }
                    selected = Some(index);
                }
            }
            Ok(selected)
        },
    }
}
