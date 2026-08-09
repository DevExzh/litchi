//! Concise family entry points.

use litchi_core::{Metadata, Result};
use litchi_odf_common::chart::{ChartClass, Element, Legend, PlotArea, read};
use std::path::Path;

pub use crate::authoring::Builder;
use crate::authoring::Definition;

/// Immutable document snapshot.
pub struct Chart {
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
        let chart = read(package.content_xml())?;
        Ok(Self { package, chart })
    }

    /// Open a chart package from in-memory bytes.
    ///
    /// # Errors
    ///
    /// Returns an error if the package cannot be read or the chart content
    /// cannot be parsed.
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        let package = crate::package::Snapshot::from_bytes(bytes)?;
        let chart = read(package.content_xml())?;
        Ok(Self { package, chart })
    }

    #[must_use]
    pub fn content_xml(&self) -> &str {
        self.package.content_xml()
    }

    #[must_use]
    pub fn chart(&self) -> &Element {
        &self.chart
    }

    /// Return the typed root `chart:class` without normalizing its QName.
    ///
    /// # Errors
    ///
    /// Returns an invalid-format error if the retained chart has no valid
    /// `chart:class` value.
    pub fn class(&self) -> Result<ChartClass> {
        self.chart.chart_class()
    }

    #[must_use]
    pub fn plot_area(&self) -> Option<PlotArea<'_>> {
        self.chart.plot_area()
    }

    #[must_use]
    pub fn legend(&self) -> Option<Legend<'_>> {
        self.chart.legend()
    }

    #[must_use]
    pub fn styles_xml(&self) -> Option<&str> {
        self.package.styles_xml()
    }

    #[must_use]
    pub fn metadata(&self) -> Option<&Metadata> {
        self.package.metadata()
    }

    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        self.package.as_bytes()
    }

    /// List the file names stored in the package.
    ///
    /// # Errors
    ///
    /// Returns an error if the package entries cannot be enumerated.
    pub fn files(&self) -> Result<Vec<String>> {
        self.package.files()
    }

    #[must_use]
    pub fn into_bytes(self) -> Vec<u8> {
        self.package.into_bytes()
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
