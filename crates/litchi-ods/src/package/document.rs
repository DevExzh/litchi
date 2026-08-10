use litchi_core::{Error, Result};
use litchi_odf_common::calculation::Settings;
use litchi_odf_common::core::{OwnedPackage, PackageWriter, family};
use std::path::Path;

use crate::model::names::Definition;

const MIMETYPE: &str = "application/vnd.oasis.opendocument.spreadsheet";
const BODY_MARKER: &str = "<office:spreadsheet";

/// Validated ownership boundary for an ODS package.
pub struct Package(family::Package);

impl Package {
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        let package = family::Package::open(path, MIMETYPE, BODY_MARKER, "ODS")?;
        crate::authoring::validate_content_xml(package.content_xml())?;
        Ok(Self(package))
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
        Ok(Self(package))
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
        Ok(Self(package))
    }

    #[must_use]
    pub fn content_xml(&self) -> &str {
        self.0.content_xml()
    }

    #[must_use]
    pub fn styles_xml(&self) -> Option<&str> {
        self.0.styles_xml()
    }

    #[must_use]
    pub fn package(&self) -> &OwnedPackage {
        self.0.package()
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

    /// Rebuild this package with a replacement `content.xml`.
    pub(crate) fn replace_content_xml(&self, content_xml: &str) -> Result<Self> {
        let bytes = self.rebuild(content_xml, Part::Preserve, Part::Preserve)?;
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
            return Self::from_bytes(self.0.as_bytes().to_vec());
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
        self.0.into_bytes()
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
