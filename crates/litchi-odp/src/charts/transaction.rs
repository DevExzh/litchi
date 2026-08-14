//! Clone-staged ODP chart inventories and atomic edits.

use super::codec::{inventory, locate_pages, page_index};
use super::model::{Chart, Limits, Location, Page, Part, Selector, Storage};
use super::package;
use crate::core::OwnedPackage;
use litchi_core::{Error, Result};
use std::borrow::Cow;

/// Immutable chart inventory bound to one ODP package snapshot.
pub struct Inventory<'source> {
    pub(crate) source: &'source OwnedPackage,
    pub(crate) limits: Limits,
    pub(crate) charts: Vec<Chart>,
}

impl<'source> Inventory<'source> {
    pub(crate) fn load(source: &'source OwnedPackage, limits: Limits) -> Result<Self> {
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

    /// Return the number of embedded charts in drawing order.
    #[must_use]
    pub fn len(&self) -> usize {
        self.charts.len()
    }

    /// Return whether no chart occurrence was discovered.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.charts.is_empty()
    }

    /// Iterate chart occurrences in source order.
    #[must_use]
    pub fn iter(&self) -> impl ExactSizeIterator<Item = &Chart> {
        self.charts.iter()
    }

    /// Select by checked zero-based discovery order.
    ///
    /// # Errors
    /// Returns an error when the chart data is malformed or a configured limit is exceeded.
    pub fn at(&self, index: usize) -> Result<Option<&Chart>> {
        Ok(self.charts.get(index))
    }

    /// Select by exact producer-visible frame name.
    ///
    /// # Errors
    /// Returns an error when the chart data is malformed or a configured limit is exceeded.
    pub fn named(&self, name: &str) -> Result<Option<&Chart>> {
        select(&self.charts, Selector::Name(name))
            .map(|index| index.map(|selected| &self.charts[selected]))
    }

    /// Select by exact name or checked zero-based position.
    ///
    /// # Errors
    /// Returns an error when the chart data is malformed or a configured limit is exceeded.
    pub fn get<'a, S>(&self, selector: S) -> Result<Option<&Chart>>
    where
        S: Into<Selector<'a>>,
    {
        select(&self.charts, selector.into())
            .map(|index| index.map(|selected| &self.charts[selected]))
    }

    /// Start an isolated clone-staged transaction.
    #[must_use]
    pub fn transaction(&self) -> Transaction<'source> {
        Transaction {
            source: self.source,
            limits: self.limits,
            original: self.charts.clone(),
            draft: self.charts.clone(),
            next_add_token: 0,
        }
    }
}

/// An isolated mutable draft of an immutable chart inventory.
pub struct Transaction<'source> {
    pub(crate) source: &'source OwnedPackage,
    pub(crate) limits: Limits,
    pub(crate) original: Vec<Chart>,
    pub(crate) draft: Vec<Chart>,
    next_add_token: usize,
}

impl<'source> Transaction<'source> {
    /// Borrow the current staged chart values.
    #[must_use]
    pub fn charts(&self) -> &[Chart] {
        &self.draft
    }

    /// Open a short-lived contextual chart editor.
    pub fn editor(&mut self) -> Editor<'_, 'source> {
        Editor { transaction: self }
    }

    /// Replace one chart part, including every occurrence sharing its package part.
    ///
    /// # Errors
    /// Returns an error when the chart data is malformed or a configured limit is exceeded.
    pub fn replace<'a, S>(&mut self, selector: S, part: Part) -> Result<()>
    where
        S: Into<Selector<'a>>,
    {
        self.editor().replace(selector, part)
    }

    /// Remove one chart occurrence from its containing drawing page.
    ///
    /// # Errors
    /// Returns an error when the chart data is malformed or a configured limit is exceeded.
    pub fn remove<'a, S>(&mut self, selector: S) -> Result<Chart>
    where
        S: Into<Selector<'a>>,
    {
        self.editor().remove(selector)
    }

    /// Add a named chart frame to a selected page.
    ///
    /// # Errors
    /// Returns an error when the chart data is malformed or a configured limit is exceeded.
    pub fn add<'a, P, N>(&mut self, page: P, name: N, storage: Storage, part: Part) -> Result<usize>
    where
        P: Into<Page<'a>>,
        N: Into<String>,
    {
        self.editor().add(page, name, storage, part)
    }

    /// Publish all staged structural and part edits atomically.
    ///
    /// # Errors
    /// Returns an error when the chart data is malformed or a configured limit is exceeded.
    pub fn commit(self) -> Result<Commit<'source>> {
        let changed = self.original != self.draft;
        let package = if changed {
            Some(OwnedPackage::from_bytes(package::apply(
                self.source,
                &self.original,
                &self.draft,
            )?)?)
        } else {
            None
        };
        if let Some(package) = package.as_ref() {
            super::snapshot::validate_compact_package(package)?;
            let reopened =
                super::snapshot::Snapshot::from_owned_package(package.clone(), self.limits)?;
            if !super::snapshot::charts_semantically_equal(reopened.charts(), &self.draft) {
                return invalid("ODP chart transaction failed typed readback");
            }
        }
        let bytes = package
            .map(|package| Cow::Owned(package.into_inner()))
            .unwrap_or_else(|| Cow::Borrowed(self.source.as_bytes()));
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

    /// Replace one chart part after resolving a semantic selector.
    ///
    /// # Errors
    /// Returns an error when the chart data is malformed or a configured limit is exceeded.
    #[allow(
        clippy::needless_pass_by_value,
        reason = "matches the by-value part-taking signature of Transaction::replace; changing it to a borrow would alter the public API"
    )]
    pub fn replace<'a, S>(&mut self, selector: S, part: Part) -> Result<()>
    where
        S: Into<Selector<'a>>,
    {
        self.validate_part(&part)?;
        let index = select(&self.transaction.draft, selector.into())?
            .ok_or_else(|| invalid_error("ODP embedded chart selector did not match"))?;
        let selected = self.transaction.draft[index].clone();
        for chart in &mut self.transaction.draft {
            if chart.same_identity(&selected) || chart.shares_storage(&selected) {
                *chart = chart.with_part(part.clone());
            }
        }
        self.validate_total()?;
        Ok(())
    }

    /// Remove one chart occurrence and return its staged source value.
    ///
    /// # Errors
    /// Returns an error when the chart data is malformed or a configured limit is exceeded.
    pub fn remove<'a, S>(&mut self, selector: S) -> Result<Chart>
    where
        S: Into<Selector<'a>>,
    {
        let index = select(&self.transaction.draft, selector.into())?
            .ok_or_else(|| invalid_error("ODP embedded chart selector did not match"))?;
        Ok(self.transaction.draft.remove(index))
    }

    /// Add a named chart frame to a selected `draw:page`.
    ///
    /// # Errors
    /// Returns an error when the chart data is malformed or a configured limit is exceeded.
    pub fn add<'a, P, N>(&mut self, page: P, name: N, storage: Storage, part: Part) -> Result<usize>
    where
        P: Into<Page<'a>>,
        N: Into<String>,
    {
        self.validate_part(&part)?;
        let frame_name = name.into();
        validate_name(&frame_name)?;
        if self
            .transaction
            .draft
            .iter()
            .any(|chart| chart.name() == Some(frame_name.as_str()))
        {
            return invalid("ODP chart frame names must be unique for selector stability");
        }
        let content = super::codec::content_xml(self.transaction.source)?;
        let target_page = page.into();
        let page_index = page_index(&content, target_page)?;
        let page_name = locate_pages(&content)?
            .into_iter()
            .find(|value| value.index == page_index)
            .and_then(|value| value.name);
        let index = self.transaction.draft.len();
        let token = self.transaction.next_add_token;
        self.transaction.next_add_token = token
            .checked_add(1)
            .ok_or_else(|| invalid_error("ODP chart insertion token overflow"))?;
        self.transaction.draft.push(Chart {
            frame: Some(litchi_odf_common::drawing::Frame {
                name: Some(frame_name),
                page_name,
                ..Default::default()
            }),
            storage,
            part,
            location: Location::Added { page_index, token },
        });
        self.validate_total()?;
        Ok(index)
    }

    fn validate_part(&self, part: &Part) -> Result<()> {
        if part.xml().is_empty() || part.xml().len() > self.transaction.limits.max_part_bytes() {
            return invalid("ODP replacement chart exceeds the part-byte limit");
        }
        Ok(())
    }

    fn validate_total(&self) -> Result<()> {
        let total = self
            .transaction
            .draft
            .iter()
            .try_fold(0usize, |total, chart| {
                total
                    .checked_add(chart.part().xml().len())
                    .ok_or_else(|| invalid_error("ODP chart byte count overflow"))
            })?;
        if total > self.transaction.limits.max_total_bytes() {
            return invalid("ODP staged chart content exceeds the total-byte limit");
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
                        return invalid("ODP embedded chart name is ambiguous");
                    }
                    selected = Some(index);
                }
            }
            Ok(selected)
        },
    }
}

fn validate_name(name: &str) -> Result<()> {
    if name.is_empty() || name.len() > 64 * 1024 || name.chars().any(char::is_control) {
        return invalid("ODP chart frame name is empty or invalid");
    }
    Ok(())
}

fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(invalid_error(message))
}

fn invalid_error(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}
