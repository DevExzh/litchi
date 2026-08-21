use litchi_core::{Error, Result};
use litchi_odf_common::calculation::Settings;
use litchi_odf_common::core::{
    OwnedPackage, PackageWriter, PreparedPackage, XmlSplicePublication, family,
};
use litchi_odf_common::package::{
    replace_content_xml as replace_package_content_xml, replace_content_xml_spliced,
};
use std::sync::Arc;
use std::{fmt, path::Path};

use crate::model::names::Definition;

const MIMETYPE: &str = "application/vnd.oasis.opendocument.spreadsheet";
const BODY_MARKER: &str = "<office:spreadsheet";

/// Validated ownership boundary for an ODS package.
///
/// The family package is immutable after construction.  Keeping one private
/// shared handle here lets nested ODS transaction owners pass a validated
/// archive/index handoff without reparsing the ZIP central directory.  The
/// wrapper remains crate-private in ordinary signatures; callers still see
/// semantic snapshots and byte slices.
pub struct Package {
    inner: Arc<family::Package>,
}

impl fmt::Debug for Package {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Package")
            .field("bytes", &self.inner.as_bytes().len())
            .field(
                "prepared_index",
                &self.inner.package().prepared_index_identity(),
            )
            .finish()
    }
}

impl Package {
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        Self::from_bytes(std::fs::read(path.as_ref())?)
    }

    /// Open a password-encrypted ODS package and validate its decrypted semantic owners.
    ///
    /// # Errors
    ///
    /// Returns an error for file I/O, an incorrect password, malformed encryption metadata,
    /// MIME/body mismatch, invalid XML, or worksheet validation failure.
    pub fn open_with_password(path: impl AsRef<Path>, password: impl Into<String>) -> Result<Self> {
        Self::from_bytes_with_password(std::fs::read(path)?, password)
    }

    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        let package = family::Package::from_bytes(bytes, MIMETYPE, BODY_MARKER, "ODS")?;
        crate::authoring::validate_content_xml(package.content_xml())?;
        Ok(Self {
            inner: Arc::new(package),
        })
    }

    /// Adopt the indexed package retained by smart ODF detection.
    ///
    /// The prepared value is consumed only after the ODS MIME contract is
    /// checked, so a mismatched family is reported without exposing or
    /// rebuilding the underlying archive index.
    pub fn from_prepared_package(prepared: PreparedPackage) -> Result<Self> {
        if prepared.format() != litchi_core::detection::FileFormat::Ods {
            return Err(Error::InvalidFormat(
                "prepared ODF package is not an ODS family document".to_string(),
            ));
        }
        let package = family::Package::from_owned_package(
            prepared.into_package(),
            MIMETYPE,
            BODY_MARKER,
            "ODS",
        )?;
        crate::authoring::validate_content_xml(package.content_xml())?;
        Ok(Self {
            inner: Arc::new(package),
        })
    }

    /// Alias for [`Self::from_prepared_package`].
    #[inline]
    pub fn from_prepared(prepared: PreparedPackage) -> Result<Self> {
        Self::from_prepared_package(prepared)
    }

    /// Adopt an already materialized, validated archive without copying its
    /// bytes.  This is the explicit handoff from the positional source owner
    /// to the established mutable package facade.
    pub(crate) fn from_owned_package(archive: OwnedPackage) -> Result<Self> {
        let package = family::Package::from_owned_package(archive, MIMETYPE, BODY_MARKER, "ODS")?;
        crate::authoring::validate_content_xml(package.content_xml())?;
        Ok(Self {
            inner: Arc::new(package),
        })
    }

    /// Adopt shared ODS bytes for an internal source-bound transaction.
    pub(crate) fn from_shared_bytes(bytes: Arc<Vec<u8>>) -> Result<Self> {
        let package = family::Package::from_shared_bytes(bytes, MIMETYPE, BODY_MARKER, "ODS")?;
        crate::authoring::validate_content_xml(package.content_xml())?;
        Ok(Self {
            inner: Arc::new(package),
        })
    }

    /// Decode a password-encrypted ODS package and validate its decrypted semantic owners.
    ///
    /// # Errors
    ///
    /// Returns an error for an incorrect password, malformed encryption metadata, MIME/body
    /// mismatch, invalid XML, or worksheet validation failure.
    pub fn from_bytes_with_password(bytes: Vec<u8>, password: impl Into<String>) -> Result<Self> {
        let package = family::Package::from_bytes_with_password(
            bytes,
            password,
            MIMETYPE,
            BODY_MARKER,
            "ODS",
        )?;
        crate::authoring::validate_content_xml(package.content_xml())?;
        Ok(Self {
            inner: Arc::new(package),
        })
    }

    /// Return a cloneable package owner that never retains a decryption
    /// credential.  The common owned archive clone shares both its bytes and
    /// prepared ZIP index; family XML projections are validated without
    /// reparsing the archive.
    pub(crate) fn clone_without_password(&self) -> Result<Self> {
        let package = family::Package::from_owned_package(
            self.inner.package().clone_without_password(),
            MIMETYPE,
            BODY_MARKER,
            "ODS",
        )?;
        crate::authoring::validate_content_xml(package.content_xml())?;
        Ok(Self {
            inner: Arc::new(package),
        })
    }

    #[must_use]
    pub fn content_xml(&self) -> &str {
        self.inner.content_xml()
    }

    #[must_use]
    pub fn styles_xml(&self) -> Option<&str> {
        self.inner.styles_xml()
    }

    #[must_use]
    pub fn package(&self) -> &OwnedPackage {
        self.inner.package()
    }

    /// Return the identity of the archive index retained by smart detection.
    #[doc(hidden)]
    #[must_use]
    pub fn prepared_index_identity(&self) -> usize {
        self.inner.package().prepared_index_identity()
    }

    /// Decode the complete ODS metadata snapshot, retaining the bounded
    /// source part for unknown-XML-preserving transactions.
    pub(crate) fn metadata_snapshot(&self) -> Result<crate::metadata::Snapshot> {
        let source = if self.package().has_file("meta.xml")? {
            let bytes = self.package().get_file("meta.xml")?;
            Some(String::from_utf8(bytes).map_err(|_error| {
                Error::InvalidFormat("ODS meta.xml is not valid UTF-8".to_string())
            })?)
        } else {
            None
        };
        crate::metadata::Snapshot::from_source(source)
    }

    /// Decode the optional spreadsheet calculation-settings owner.
    pub(crate) fn calculation_settings(&self) -> Result<Option<Settings>> {
        crate::settings::Snapshot::from_content_xml(self.content_xml())
            .map(|snapshot| snapshot.calculation().cloned())
    }

    /// Read the ordered named-definition catalog from `content.xml`.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn definitions(&self) -> Result<Vec<Definition>> {
        crate::codec::names::parse(self.content_xml())
    }

    /// Decode the public typed worksheet graph without expanding repetition
    /// runs into one object per logical row or cell.
    pub(crate) fn sheets(&self) -> Result<Vec<crate::worksheet::Sheet>> {
        crate::worksheet::codec::parse(self.content_xml())
    }

    #[cfg(test)]
    pub(crate) fn shared_bytes(&self) -> Arc<Vec<u8>> {
        self.inner.shared_bytes()
    }

    /// Clone the exact immutable archive byte owner retained by this package.
    pub(crate) fn shared_bytes_owner(&self) -> Arc<Vec<u8>> {
        self.inner.shared_bytes()
    }

    /// Rebuild this package with a replacement `content.xml`.
    pub(crate) fn replace_content_xml(&self, content_xml: &str) -> Result<Self> {
        let bytes = replace_package_content_xml(self.package(), content_xml)?;
        Self::from_bytes(bytes)
    }

    /// Rebuild this package from exact checked `content.xml` source ranges.
    pub(crate) fn replace_spliced_content_xml(
        &self,
        content_xml: &str,
        publication: XmlSplicePublication,
    ) -> Result<Self> {
        let bytes = replace_content_xml_spliced(self.package(), content_xml, publication)?;
        Self::from_bytes(bytes)
    }

    /// Publish an exact-source checked tracked-change commit.
    pub(crate) fn replace_tracked_changes(
        &self,
        commit: &crate::tracked_changes::Commit,
    ) -> Result<Self> {
        if !commit.changed() {
            return Err(Error::InvalidFormat(
                "unchanged tracked-change commit must not rebuild the package".to_string(),
            ));
        }
        if self.content_xml() != commit.patch().source() {
            return Err(Error::InvalidFormat(
                "tracked-change package source snapshot does not match".to_string(),
            ));
        }
        self.replace_content_xml(commit.content_xml())
    }

    /// Replace or remove `meta.xml` while preserving every other package
    /// member and all unknown metadata XML not selected by the common patch.
    pub(crate) fn replace_metadata_xml(&self, metadata_xml: Option<&str>) -> Result<Self> {
        let bytes = self.rebuild(
            self.content_xml(),
            Part::from_option(metadata_xml),
            Part::Preserve,
        )?;
        Self::from_bytes(bytes)
    }

    /// Replace the content-level calculation-settings owner atomically.
    pub(crate) fn replace_calculation_settings(&self, settings: Option<&Settings>) -> Result<Self> {
        let snapshot = crate::settings::Snapshot::from_content_xml(self.content_xml())?;
        let mut transaction = snapshot.transaction();
        match settings {
            Some(settings) => transaction.replace(settings.clone())?,
            None => transaction.remove(),
        }
        let commit = transaction.commit()?;
        if !commit.changed() {
            return Self::from_bytes(self.inner.as_bytes().to_vec());
        }
        let content_xml = commit.into_owned();
        let bytes = self.rebuild(&content_xml, Part::Preserve, Part::Preserve)?;
        Self::from_bytes(bytes)
    }

    fn rebuild(
        &self,
        content_xml: &str,
        metadata: Part<'_>,
        settings: Part<'_>,
    ) -> Result<Vec<u8>> {
        let source = self.package();
        let mut writer = PackageWriter::new();
        writer.set_mimetype(&source.mimetype()?)?;
        writer.add_file("content.xml", content_xml.as_bytes())?;
        if source.has_file("styles.xml")? {
            writer.add_file("styles.xml", &source.get_file("styles.xml")?)?;
        }
        write_part(&mut writer, source, "meta.xml", metadata)?;
        write_part(&mut writer, source, "settings.xml", settings)?;
        writer.copy_auxiliary_files_from_except(
            source,
            &["meta.xml".to_string(), "settings.xml".to_string()],
            &[],
        )?;
        writer.finish_to_bytes()
    }

    /// Rebuild the package after an atomic semantic worksheet edit.
    pub(crate) fn replace_sheets(&self, sheets: &[crate::worksheet::Sheet]) -> Result<Self> {
        let content_xml = crate::worksheet::package::replace_tables(self.content_xml(), sheets)?;
        self.replace_content_xml(&content_xml)
    }

    #[must_use]
    pub fn into_bytes(self) -> Vec<u8> {
        match Arc::try_unwrap(self.inner) {
            Ok(package) => package.into_bytes(),
            Err(package) => package.as_bytes().to_vec(),
        }
    }
}

enum Part<'a> {
    Preserve,
    Set(&'a str),
    Remove,
}

impl<'a> Part<'a> {
    fn from_option(value: Option<&'a str>) -> Self {
        match value {
            Some(value) => Self::Set(value),
            None => Self::Remove,
        }
    }
}

fn write_part(
    writer: &mut PackageWriter,
    source: &OwnedPackage,
    path: &str,
    part: Part<'_>,
) -> Result<()> {
    match part {
        Part::Preserve => {
            if source.has_file(path)? {
                writer.add_file(path, &source.get_file(path)?)?;
            }
        },
        Part::Set(xml) => writer.add_file(path, xml.as_bytes())?,
        Part::Remove => {},
    }
    Ok(())
}
