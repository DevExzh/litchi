//! Concise family entry points.

use litchi_core::{Metadata, Result};
use litchi_odf_common::chart::{Element, Legend, PlotArea, read};
use std::path::Path;

pub use crate::authoring::Builder;

/// Immutable document snapshot.
pub struct Chart {
    package: crate::package::Snapshot,
    chart: Element,
}

impl Chart {
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        let package = crate::package::Snapshot::open(path)?;
        let chart = read(package.content_xml())?;
        Ok(Self { package, chart })
    }
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        let package = crate::package::Snapshot::from_bytes(bytes)?;
        let chart = read(package.content_xml())?;
        Ok(Self { package, chart })
    }
    pub fn content_xml(&self) -> &str {
        self.package.content_xml()
    }
    pub fn chart(&self) -> &Element {
        &self.chart
    }
    pub fn plot_area(&self) -> Option<PlotArea<'_>> {
        self.chart.plot_area()
    }
    pub fn legend(&self) -> Option<Legend<'_>> {
        self.chart.legend()
    }
    pub fn styles_xml(&self) -> Option<&str> {
        self.package.styles_xml()
    }
    pub fn metadata(&self) -> Option<&Metadata> {
        self.package.metadata()
    }
    pub fn as_bytes(&self) -> &[u8] {
        self.package.as_bytes()
    }
    pub fn files(&self) -> Result<Vec<String>> {
        self.package.files()
    }
    pub fn into_bytes(self) -> Vec<u8> {
        self.package.into_bytes()
    }
}

#[cfg(test)]
mod tests {
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
