//! Formula-family archive validation and ODF ZIP orchestration.

use litchi_core::{Error, Result};
use litchi_odf_common::constants::{ODF_CONTENT, ODF_FORMULA, ODF_FORMULA_TEMPLATE};
use litchi_odf_common::core::{OwnedPackage, PackageWriter, SourceLimits};
use litchi_odf_common::package::rebuild_package;
use litchi_odf_common::{SemanticLimitKind, SemanticLimits};

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

/// An owned, family-validated Formula ZIP archive.
#[derive(Clone)]
pub(crate) struct Archive {
    owned: OwnedPackage,
    flavor: Flavor,
}

impl Archive {
    /// The retained physical package.
    pub(crate) const fn owned(&self) -> &OwnedPackage {
        &self.owned
    }

    /// Build a package containing a validated Formula-family `content.xml`.
    pub(crate) fn create(content: &[u8], template: bool) -> Result<Self> {
        let flavor = if template {
            Flavor::Template
        } else {
            Flavor::Formula
        };
        let mut writer =
            PackageWriter::new(litchi_odf_common::core::StreamingArchiveLimits::default());
        writer.set_mimetype(flavor.mimetype())?;
        writer.add_file(ODF_CONTENT, content)?;
        Self::from_bytes(writer.finish_to_bytes()?)
    }

    /// Read and validate a Formula-family package from owned bytes under the
    /// default source limits.
    pub(crate) fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        Self::from_owned(OwnedPackage::from_bytes_with_options(
            bytes,
            litchi_odf_common::OpenOptions::new().source_limits(SourceLimits::default()),
        )?)
    }

    /// Adopt an already indexed archive after checking the declared
    /// `content.xml` size against the semantic limits before inflating it.
    pub(crate) fn from_owned_checked(
        owned: OwnedPackage,
        limits: SemanticLimits,
    ) -> std::result::Result<Self, litchi_odf_common::Error> {
        let declared = owned.package()?.member_materialized_size(ODF_CONTENT)?;
        if let Some(declared) = declared {
            limits.check(SemanticLimitKind::ContentBytes, declared)?;
        }
        let archive = Self::from_owned(owned)?;
        if declared.is_none() {
            let observed = u64::try_from(archive.content_xml()?.len()).unwrap_or(u64::MAX);
            limits.check(SemanticLimitKind::ContentBytes, observed)?;
        }
        Ok(archive)
    }

    /// Adopt an already indexed archive after the family checks.
    fn from_owned(owned: OwnedPackage) -> Result<Self> {
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

    /// Return the exact package MIME type.
    #[must_use]
    pub(crate) fn mimetype(&self) -> &'static str {
        self.flavor.mimetype()
    }

    /// Whether this package uses the Formula template MIME type.
    #[must_use]
    pub(crate) const fn is_template(&self) -> bool {
        matches!(self.flavor, Flavor::Template)
    }

    /// Read the package's exact `content.xml` bytes as UTF-8.
    ///
    /// # Errors
    ///
    /// Returns an error when `content.xml` is missing or not valid UTF-8.
    pub(crate) fn content_xml(&self) -> Result<String> {
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
    pub(crate) fn files(&self) -> Result<Vec<String>> {
        self.owned.files()
    }

    /// The identity of the retained validated archive index.
    pub(crate) fn prepared_index_identity(&self) -> usize {
        self.owned.prepared_index_identity()
    }

    /// Return the exact original package bytes.
    #[must_use]
    pub(crate) fn as_bytes(&self) -> &[u8] {
        self.owned.as_bytes()
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
