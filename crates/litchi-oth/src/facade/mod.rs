//! Concise family entry points.

use litchi_core::{Metadata, Result};
use std::path::Path;

pub use crate::authoring::Builder;

/// Immutable document snapshot.
pub struct Template {
    package: crate::package::Package,
}

impl Template {
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        crate::package::Package::open(path).map(|package| Self { package })
    }
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        crate::package::Package::from_bytes(bytes).map(|package| Self { package })
    }
    pub fn content_xml(&self) -> &str {
        self.package.content_xml()
    }
    pub fn styles_xml(&self) -> Option<&str> {
        self.package.styles_xml()
    }
    pub fn metadata(&self) -> Option<&Metadata> {
        self.package.metadata()
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
    use super::{Builder, Template};

    #[test]
    fn builder_opens_as_validated_snapshot() {
        let bytes = Builder::new().build().unwrap();
        let document = Template::from_bytes(bytes).unwrap();
        assert!(document.content_xml().contains("<office:text"));
    }
}
