//! Owned chart snapshots, transactional edits, and reversible package patches.

use super::{Limits, Part, Selector, inventory, model::Chart, package};
use crate::package::Archive;
use litchi_core::{Error, Result};
use litchi_odf_common::XmlStreamLimits;

/// Immutable embedded-chart inventory tied to exact ODS package bytes.
#[derive(Clone, Debug)]
pub struct Snapshot {
    source: Vec<u8>,
    limits: Limits,
    charts: Vec<Chart>,
    xml_limits: XmlStreamLimits,
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
        Self::from_bytes_with_xml_limits(source, limits, XmlStreamLimits::owner_default_value())
    }

    pub(crate) fn from_bytes_with_xml_limits(
        source: Vec<u8>,
        limits: Limits,
        xml_limits: XmlStreamLimits,
    ) -> Result<Self> {
        let package = Archive::from_bytes_with_xml_limits(source.clone(), xml_limits)?;
        let charts = inventory(&package, limits)?.charts;
        Ok(Self {
            source,
            limits,
            charts,
            xml_limits,
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
    pub(crate) fn edit(&self) -> Edit {
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

    /// Discard every staged change in place.
    pub fn reset(&mut self) {
        self.draft = self.before.charts.clone();
    }

    /// Validate package topology, reparse the candidate, and atomically publish it.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub(crate) fn commit(self) -> Result<Commit> {
        let changed = self.before.charts != self.draft;
        let snapshot = if changed {
            let package_source = Archive::from_bytes_with_xml_limits(
                self.before.source.clone(),
                self.before.xml_limits,
            )?;
            let bytes = package::replace(&package_source, &self.before.charts, &self.draft)?;
            let result = Snapshot::from_bytes_with_xml_limits(
                bytes,
                self.before.limits,
                self.before.xml_limits,
            )?;
            if result.charts != self.draft {
                return Err(Error::InvalidFormat(
                    "ODS chart commit failed typed readback".to_string(),
                ));
            }
            result
        } else {
            self.before.clone()
        };
        Ok(Commit { snapshot })
    }
}

/// A fully rehydrated chart publication.
#[derive(Clone, Debug)]
pub(crate) struct Commit {
    snapshot: Snapshot,
}

impl Commit {
    /// Resulting immutable snapshot.
    #[must_use]
    pub(crate) fn snapshot(&self) -> &Snapshot {
        &self.snapshot
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
