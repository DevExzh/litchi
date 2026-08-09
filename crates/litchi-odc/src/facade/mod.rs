//! Concise family entry points.

use litchi_core::{Error, Metadata, Result};
use litchi_odf_common::chart::{ChartClass, Element, Legend, PlotArea, read};
use std::{path::Path, sync::Arc};

pub use crate::authoring::Builder;
use crate::authoring::Definition;

/// Immutable document snapshot.
#[derive(Clone)]
pub struct Chart(Arc<State>);

struct State {
    package: crate::package::Snapshot,
    chart: Element,
}

impl Chart {
    /// Build a chart snapshot from a typed definition.
    ///
    /// # Errors
    ///
    /// Returns an error if the definition fails to build into a valid chart
    /// package.
    pub fn from_definition(definition: Definition) -> Result<Self> {
        Self::from_bytes(Builder::new().with_definition(definition).build()?)
    }

    /// Open a chart package from a file path.
    ///
    /// # Errors
    ///
    /// Returns an error if the package cannot be read or the chart content
    /// cannot be parsed.
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        let package = crate::package::Snapshot::open(path)?;
        Self::from_package(package)
    }

    /// Open a chart package from in-memory bytes.
    ///
    /// # Errors
    ///
    /// Returns an error if the package cannot be read or the chart content
    /// cannot be parsed.
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        let package = crate::package::Snapshot::from_bytes(bytes)?;
        Self::from_package(package)
    }

    fn from_package(package: crate::package::Snapshot) -> Result<Self> {
        let chart = read(package.content_xml())?;
        Ok(Self(Arc::new(State { package, chart })))
    }

    #[must_use]
    pub fn content_xml(&self) -> &str {
        self.0.package.content_xml()
    }

    #[must_use]
    pub fn chart(&self) -> &Element {
        &self.0.chart
    }

    /// Return the typed root `chart:class` without normalizing its `QName`.
    ///
    /// # Errors
    ///
    /// Returns an invalid-format error if the retained chart has no valid
    /// `chart:class` value.
    pub fn class(&self) -> Result<ChartClass> {
        self.0.chart.chart_class()
    }

    #[must_use]
    pub fn plot_area(&self) -> Option<PlotArea<'_>> {
        self.0.chart.plot_area()
    }

    #[must_use]
    pub fn legend(&self) -> Option<Legend<'_>> {
        self.0.chart.legend()
    }

    #[must_use]
    pub fn styles_xml(&self) -> Option<&str> {
        self.0.package.styles_xml()
    }

    #[must_use]
    pub fn metadata(&self) -> Option<&Metadata> {
        self.0.package.metadata()
    }

    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        self.0.package.as_bytes()
    }

    /// List the file names stored in the package.
    ///
    /// # Errors
    ///
    /// Returns an error if the package entries cannot be enumerated.
    pub fn files(&self) -> Result<Vec<String>> {
        self.0.package.files()
    }

    /// Starts a source-bound package axis transaction.
    ///
    /// The transaction edits only `chart:name` attributes on existing axes.
    /// Untouched package members and unmodeled chart XML remain payload-exact.
    #[must_use]
    pub fn edit(&self) -> Edit<'_> {
        Edit {
            source: self,
            transaction: self.0.package.content_snapshot().edit(),
            replacement: None,
        }
    }

    #[must_use]
    pub fn into_bytes(self) -> Vec<u8> {
        match Arc::try_unwrap(self.0) {
            Ok(state) => state.package.into_bytes(),
            Err(state) => state.package.as_bytes().to_vec(),
        }
    }
}

/// A source-bound package chart transaction.
pub struct Edit<'a> {
    source: &'a Chart,
    transaction: crate::FlatChartEdit,
    replacement: Option<Definition>,
}

impl Edit<'_> {
    /// Stages an axis-name update by zero-based plot-area axis position.
    ///
    /// # Errors
    ///
    /// Returns an error when the selector is out of bounds or the update is
    /// outside the bounded, losslessly editable axis-name surface.
    pub fn update_axis(&mut self, index: usize, update: crate::AxisUpdate) -> Result<()> {
        if self.replacement.is_some() {
            return Err(Error::InvalidFormat(
                "an ODC whole-chart replacement is already staged".to_string(),
            ));
        }
        self.transaction.update_axis(index, update)
    }

    /// Stages explicit replacement of the complete chart content definition.
    ///
    /// This deliberately replaces the selected chart subtree, including any
    /// unknown XML within it. Other package member payloads remain preserved.
    /// A replacement supersedes axis updates staged earlier in this edit.
    ///
    /// # Errors
    ///
    /// Returns an error when the detached definition violates the bounded
    /// typed chart invariants.
    pub fn replace_chart(&mut self, definition: &Definition) -> Result<()> {
        definition.validate()?;
        self.replacement = Some(definition.clone());
        Ok(())
    }

    /// Atomically validates, rebuilds, and publishes the package edit.
    ///
    /// # Errors
    ///
    /// Returns an error when the source cannot be losslessly rewritten, when
    /// the package is signed, or when full reopen and typed readback fail.
    pub fn commit(mut self) -> Result<Commit> {
        if let Some(definition) = self.replacement.take() {
            return self.commit_replacement(&definition);
        }
        let content_commit = self.transaction.commit()?;
        let changes = content_commit.patch().changes().to_vec();
        let snapshot = if changes.is_empty() {
            self.source.clone()
        } else {
            let content = std::str::from_utf8(content_commit.snapshot().as_bytes()).map_err(
                |_utf8_error| Error::InvalidFormat("edited ODC content is not UTF-8".to_string()),
            )?;
            let package = self.source.0.package.rebuild_with_content(content)?;
            Chart::from_package(package)?
        };
        for change in &changes {
            let actual = snapshot
                .plot_area()
                .and_then(|plot_area| plot_area.axes().nth(change.index()))
                .ok_or_else(|| {
                    Error::InvalidFormat("edited ODC axis disappeared during readback".to_string())
                })?;
            if actual.name() != change.after() {
                return Err(Error::InvalidFormat(
                    "ODC package edit failed semantic readback".to_string(),
                ));
            }
        }
        Ok(Commit {
            snapshot: snapshot.clone(),
            patch: Patch {
                source: self.source.clone(),
                target: snapshot,
                changes,
                replaces_chart: false,
            },
        })
    }

    fn commit_replacement(self, definition: &Definition) -> Result<Commit> {
        let content = crate::serialize_content(definition)?;
        if content == self.source.content_xml() {
            return Ok(Commit {
                snapshot: self.source.clone(),
                patch: Patch {
                    source: self.source.clone(),
                    target: self.source.clone(),
                    changes: Vec::new(),
                    replaces_chart: false,
                },
            });
        }
        let package = self.source.0.package.rebuild_with_content(&content)?;
        let snapshot = Chart::from_package(package)?;
        if snapshot.content_xml() != content || snapshot.class()? != definition.class {
            return Err(Error::InvalidFormat(
                "ODC chart replacement failed typed readback".to_string(),
            ));
        }
        Ok(Commit {
            snapshot: snapshot.clone(),
            patch: Patch {
                source: self.source.clone(),
                target: snapshot,
                changes: Vec::new(),
                replaces_chart: true,
            },
        })
    }
}

/// A committed immutable chart package and its reversible patch.
pub struct Commit {
    snapshot: Chart,
    patch: Patch,
}

impl Commit {
    /// Returns whether package bytes changed.
    #[must_use]
    pub fn changed(&self) -> bool {
        self.patch.replaces_chart || !self.patch.changes.is_empty()
    }

    /// Returns the committed chart snapshot.
    #[must_use]
    pub fn chart(&self) -> &Chart {
        &self.snapshot
    }

    /// Returns the exact-source reversible patch.
    #[must_use]
    pub fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Consumes the commit into its immutable chart snapshot.
    #[must_use]
    pub fn into_chart(self) -> Chart {
        self.snapshot
    }
}

/// A source-checked reversible package chart patch.
#[derive(Clone)]
pub struct Patch {
    source: Chart,
    target: Chart,
    changes: Vec<crate::AxisChange>,
    replaces_chart: bool,
}

impl Patch {
    /// Returns whether this patch applies to the supplied exact package bytes.
    #[must_use]
    pub fn is_applicable_to(&self, source: &Chart) -> bool {
        self.source.as_bytes() == source.as_bytes()
    }

    /// Applies this patch only to its exact immutable source.
    ///
    /// # Errors
    ///
    /// Returns an error when the supplied package differs byte-for-byte from
    /// the source against which this patch was committed.
    pub fn apply(&self, source: &Chart) -> Result<Chart> {
        if !self.is_applicable_to(source) {
            return Err(Error::InvalidFormat(
                "ODC package patch source does not match its expected snapshot".to_string(),
            ));
        }
        Ok(self.target.clone())
    }

    /// Returns the semantic axis changes in transaction order.
    #[must_use]
    pub fn changes(&self) -> &[crate::AxisChange] {
        &self.changes
    }

    /// Returns whether this patch replaces the complete chart definition.
    #[must_use]
    pub fn replaces_chart(&self) -> bool {
        self.replaces_chart
    }

    /// Returns a patch that restores the exact source package.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            source: self.target.clone(),
            target: self.source.clone(),
            changes: self
                .changes
                .iter()
                .map(crate::AxisChange::new_inverse)
                .collect(),
            replaces_chart: self.replaces_chart,
        }
    }
}

#[cfg(test)]
mod tests {
    #![allow(
        clippy::unwrap_used,
        reason = "tests are expected to panic on unexpected errors"
    )]

    use super::{Builder, Chart};

    #[test]
    fn builder_opens_as_validated_snapshot() {
        let bytes = Builder::new().build().unwrap();
        let document = Chart::from_bytes(bytes).unwrap();
        assert!(document.content_xml().contains("<office:chart"));
        assert!(document.plot_area().is_some());
        assert!(!document.as_bytes().is_empty());
    }
}
