//! Clone-staged ODS chart inventory and atomic replacement transactions.

use super::codec::inventory;
use super::model::{Chart, Limits, Part, Selector};
use super::package;
use crate::package::Package;
use litchi_core::{Error, Result};
use std::borrow::Cow;

/// Immutable chart inventory bound to one ODS package snapshot.
pub struct Inventory<'source> {
    pub(crate) source: &'source Package,
    pub(crate) limits: Limits,
    pub(crate) charts: Vec<Chart>,
}

impl<'source> Inventory<'source> {
    pub(crate) fn load(source: &'source Package, limits: Limits) -> Result<Self> {
        Ok(Self {
            source,
            limits,
            charts: inventory(source, limits)?,
        })
    }

    /// Return the resource budget used for this inventory.
    #[must_use]
    pub const fn limits(&self) -> Limits {
        self.limits
    }

    /// Return the number of content-level embedded charts.
    #[must_use]
    pub fn len(&self) -> usize {
        self.charts.len()
    }

    /// Return whether the inventory is empty.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.charts.is_empty()
    }

    /// Iterate charts in drawing discovery order.
    #[must_use]
    pub fn iter(&self) -> impl ExactSizeIterator<Item = &Chart> {
        self.charts.iter()
    }

    /// Select a chart by checked zero-based discovery order.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn at(&self, index: usize) -> Result<Option<&Chart>> {
        Ok(self.charts.get(index))
    }

    /// Select a chart by exact producer-visible drawing name.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn named(&self, name: &str) -> Result<Option<&Chart>> {
        select(&self.charts, Selector::Name(name))
            .map(|index| index.map(|index| &self.charts[index]))
    }

    /// Select by either exact name or checked zero-based position.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn get<'a, S>(&self, selector: S) -> Result<Option<&Chart>>
    where
        S: Into<Selector<'a>>,
    {
        select(&self.charts, selector.into()).map(|index| index.map(|index| &self.charts[index]))
    }

    /// Start an isolated clone-staged transaction.
    #[must_use]
    pub fn transaction(&self) -> Transaction<'source> {
        Transaction {
            source: self.source,
            limits: self.limits,
            original: self.charts.clone(),
            draft: self.charts.clone(),
        }
    }
}

/// An isolated mutable draft of an immutable chart inventory.
pub struct Transaction<'source> {
    pub(crate) source: &'source Package,
    pub(crate) limits: Limits,
    pub(crate) original: Vec<Chart>,
    pub(crate) draft: Vec<Chart>,
}

impl<'source> Transaction<'source> {
    /// Borrow the current staged chart values.
    #[must_use]
    pub fn charts(&self) -> &[Chart] {
        &self.draft
    }

    /// Open a short-lived contextual editor over this transaction.
    pub fn editor(&mut self) -> Editor<'_, 'source> {
        Editor { transaction: self }
    }

    /// Replace one chart part and stage the operation atomically.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn replace<'a, S>(&mut self, selector: S, part: Part) -> Result<()>
    where
        S: Into<Selector<'a>>,
    {
        self.editor().replace(selector, part)
    }

    /// Publish the staged replacement as one package commit.
    ///
    /// An unchanged transaction borrows the original archive bytes, so a
    /// caller can prove that a no-op did not normalize the ZIP container.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn commit(self) -> Result<Commit<'source>> {
        let changed = self
            .original
            .iter()
            .zip(&self.draft)
            .any(|(before, after)| before.part().xml() != after.part().xml());
        let bytes = if changed {
            Cow::Owned(package::replace(self.source, &self.original, &self.draft)?)
        } else {
            Cow::Borrowed(self.source.package().as_bytes())
        };
        Ok(Commit {
            bytes,
            charts: self.draft,
            changed,
        })
    }
}

/// A short-lived contextual chart editor.
pub struct Editor<'transaction, 'source> {
    transaction: &'transaction mut Transaction<'source>,
}

impl Editor<'_, '_> {
    /// Borrow the current staged chart values.
    #[must_use]
    pub fn charts(&self) -> &[Chart] {
        self.transaction.charts()
    }

    /// Replace one chart part after resolving the semantic selector.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn replace<'a, S>(&mut self, selector: S, part: Part) -> Result<()>
    where
        S: Into<Selector<'a>>,
    {
        if part.xml().len() > self.transaction.limits.max_part_bytes() {
            return Err(Error::InvalidFormat(
                "ODS replacement chart exceeds the part-byte limit".to_string(),
            ));
        }
        let index = select(&self.transaction.draft, selector.into())?.ok_or_else(|| {
            Error::InvalidFormat("ODS embedded chart selector did not match".to_string())
        })?;
        let selected = self.transaction.draft[index].clone();
        for chart in &mut self.transaction.draft {
            if chart == &selected || chart.shares_storage(&selected) {
                *chart = chart.with_part(part.clone());
            }
        }
        Ok(())
    }
}

/// The result of publishing a chart transaction.
pub struct Commit<'source> {
    bytes: Cow<'source, [u8]>,
    charts: Vec<Chart>,
    changed: bool,
}

impl Commit<'_> {
    /// Borrow the resulting package bytes.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.bytes
    }

    /// Borrow the staged chart inventory represented by the commit.
    #[must_use]
    pub fn charts(&self) -> &[Chart] {
        &self.charts
    }

    /// Return whether package bytes had to be rebuilt.
    #[must_use]
    pub const fn changed(&self) -> bool {
        self.changed
    }
}

impl<'source> Commit<'source> {
    /// Consume the commit while retaining a borrow for an unchanged result.
    #[must_use]
    pub fn into_bytes(self) -> Cow<'source, [u8]> {
        self.bytes
    }

    /// Consume the commit into owned package bytes.
    #[must_use]
    pub fn into_owned_bytes(self) -> Vec<u8> {
        self.bytes.into_owned()
    }
}

fn select<'a, S>(charts: &[Chart], selector: S) -> Result<Option<usize>>
where
    S: Into<Selector<'a>>,
{
    match selector.into() {
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
