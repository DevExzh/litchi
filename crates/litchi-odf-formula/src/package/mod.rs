//! Formula-family package ownership and ODF ZIP orchestration.

use litchi_core::{Error, Result};
use litchi_odf_common::constants::{ODF_CONTENT, ODF_FORMULA, ODF_FORMULA_TEMPLATE};
use litchi_odf_common::core::{OwnedPackage, PackageWriter};
use litchi_odf_common::package::rebuild_package;
use std::io::Read;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Flavor {
    Formula,
    Template,
}

impl Flavor {
    const fn mimetype(self) -> &'static str {
        match self {
            Self::Formula => ODF_FORMULA,
            Self::Template => ODF_FORMULA_TEMPLATE,
        }
    }

    fn from_mimetype(mimetype: &str) -> Result<Self> {
        match mimetype {
            ODF_FORMULA => Ok(Self::Formula),
            ODF_FORMULA_TEMPLATE => Ok(Self::Template),
            other => Err(Error::InvalidFormat(format!(
                "not an OpenDocument Formula package: MIME type is '{other}'"
            ))),
        }
    }
}

/// An owned, family-validated Formula ZIP package.
#[derive(Clone)]
pub struct Package {
    owned: OwnedPackage,
    flavor: Flavor,
}

impl Package {
    /// Build a package containing a validated Formula-family `content.xml`.
    pub(crate) fn create(content: &[u8], template: bool) -> Result<Self> {
        let flavor = if template {
            Flavor::Template
        } else {
            Flavor::Formula
        };
        let mut writer = PackageWriter::new();
        writer.set_mimetype(flavor.mimetype())?;
        writer.add_file(ODF_CONTENT, content)?;
        Self::from_bytes(writer.finish_to_bytes()?)
    }

    /// Read and validate a Formula-family package from owned bytes.
    ///
    /// # Errors
    ///
    /// Returns an error when the bytes are not a valid Formula-family package
    /// (bad archive, wrong MIME type, or missing or non-UTF-8 `content.xml`).
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        let owned = OwnedPackage::from_bytes(bytes)?;
        let flavor = Flavor::from_mimetype(&owned.mimetype()?)?;
        if !owned.has_file(ODF_CONTENT)? {
            return Err(Error::InvalidFormat(
                "Formula package is missing content.xml".to_string(),
            ));
        }
        let content = owned.get_file(ODF_CONTENT)?;
        if std::str::from_utf8(&content).is_err() {
            return Err(Error::InvalidFormat(
                "Formula content.xml is not valid UTF-8".to_string(),
            ));
        }
        Ok(Self { owned, flavor })
    }

    /// Read and validate a Formula-family package from a stream.
    ///
    /// # Errors
    ///
    /// Returns an error when reading fails or the bytes are not a valid
    /// Formula-family package.
    pub fn from_reader(mut reader: impl Read) -> Result<Self> {
        let mut bytes = Vec::new();
        reader.read_to_end(&mut bytes)?;
        Self::from_bytes(bytes)
    }

    /// Return the exact package MIME type.
    #[must_use]
    pub fn mimetype(&self) -> &'static str {
        self.flavor.mimetype()
    }

    /// Whether this package uses the Formula template MIME type.
    #[must_use]
    pub const fn is_template(&self) -> bool {
        matches!(self.flavor, Flavor::Template)
    }

    /// Read the package's exact `content.xml` bytes as UTF-8.
    ///
    /// # Errors
    ///
    /// Returns an error when `content.xml` is missing or not valid UTF-8.
    pub fn content_xml(&self) -> Result<String> {
        let bytes = self.owned.get_file(ODF_CONTENT)?;
        match String::from_utf8(bytes) {
            Ok(text) => Ok(text),
            Err(_) => Err(Error::InvalidFormat(
                "Formula content.xml is not valid UTF-8".to_string(),
            )),
        }
    }

    /// List the package members in archive order.
    ///
    /// # Errors
    ///
    /// Returns an error when the archive member list cannot be read.
    pub fn files(&self) -> Result<Vec<String>> {
        self.owned.files()
    }

    /// Return the exact original package bytes.
    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        self.owned.as_bytes()
    }

    /// Consume the package and return its exact bytes.
    #[must_use]
    pub fn into_bytes(self) -> Vec<u8> {
        self.owned.into_inner()
    }

    /// Replace `content.xml` while preserving non-core package members.
    pub(crate) fn replace_content(&self, content: &[u8]) -> Result<Self> {
        let content_xml = std::str::from_utf8(content).map_err(|_encoding_error| {
            Error::InvalidFormat("Formula content.xml is not valid UTF-8".to_string())
        })?;
        Self::from_bytes(rebuild_package(
            &self.owned,
            content_xml,
            Vec::new(),
            Vec::new(),
            [],
            [],
        )?)
    }
}
