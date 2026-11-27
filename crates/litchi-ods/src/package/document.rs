use litchi_core::{Error, Result};
use litchi_odf_common::XmlStreamLimits;
use litchi_odf_common::calculation::Settings;
use litchi_odf_common::core::{
    OwnedPackage, PackageWriter, PreparedPackage, XmlSplicePublication, family,
};
use litchi_odf_common::package::{
    replace_content_xml as replace_package_content_xml, replace_content_xml_spliced,
};
use std::borrow::Cow;
use std::fmt;
use std::sync::Arc;

use crate::model::names::Definition;

const MIMETYPE: &str = "application/vnd.oasis.opendocument.spreadsheet";

/// `xml` without the whitespace between its XML declaration and its root
/// element. Prolog whitespace is no part of the document (XML 1.0 §2.8);
/// LibreOffice ends every declaration with a newline, which the compact
/// authored-XML audit would otherwise refuse when a family republishes a
/// producer part it loaded.
pub(crate) fn compact_prolog(xml: &str) -> Cow<'_, str> {
    let Some(declaration_end) = xml
        .strip_prefix("<?xml")
        .and_then(|rest| rest.find("?>"))
        .map(|end| end + "<?xml".len() + 2)
    else {
        return Cow::Borrowed(xml);
    };
    let tail = &xml[declaration_end..];
    let trimmed = tail.trim_start_matches([' ', '\t', '\r', '\n']);
    if trimmed.len() == tail.len() || !trimmed.starts_with('<') {
        return Cow::Borrowed(xml);
    }
    let mut out = String::with_capacity(xml.len() - (tail.len() - trimmed.len()));
    out.push_str(&xml[..declaration_end]);
    out.push_str(trimmed);
    Cow::Owned(out)
}

/// [`compact_prolog`] over a UTF-8 part; other bytes are returned as they
/// are.
pub(crate) fn compact_prolog_bytes(bytes: &[u8]) -> Cow<'_, [u8]> {
    match std::str::from_utf8(bytes).map(compact_prolog) {
        Ok(Cow::Owned(xml)) => Cow::Owned(xml.into_bytes()),
        Ok(Cow::Borrowed(_)) | Err(_) => Cow::Borrowed(bytes),
    }
}
const BODY_MARKER: &str = "<office:spreadsheet";

/// Validated ownership boundary for an ODS package.
///
/// The family package is immutable after construction.  Keeping one private
/// shared handle here lets nested ODS transaction owners pass a validated
/// archive/index handoff without reparsing the ZIP central directory.  The
/// wrapper remains crate-private in ordinary signatures; callers still see
/// semantic snapshots and byte slices.
pub(crate) struct Archive {
    inner: Arc<family::Package>,
    xml_limits: XmlStreamLimits,
    resource_count: std::sync::OnceLock<usize>,
}

impl fmt::Debug for Archive {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Archive")
            .field("bytes", &self.inner.as_bytes().len())
            .field(
                "prepared_index",
                &self.inner.package().prepared_index_identity(),
            )
            .field("resources", &self.resource_count.get())
            .finish()
    }
}

impl Archive {
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub(crate) fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        Self::from_bytes_with_xml_limits(bytes, XmlStreamLimits::owner_default_value())
    }

    pub(crate) fn from_bytes_with_xml_limits(
        bytes: Vec<u8>,
        xml_limits: XmlStreamLimits,
    ) -> Result<Self> {
        let package = family::Package::from_bytes(bytes, MIMETYPE, BODY_MARKER, "ODS")?;
        Self::from_family_package(package, xml_limits)
    }

    /// Adopt the indexed package retained by smart ODF detection.
    ///
    /// The prepared value is consumed only after the ODS MIME contract is
    /// checked, so a mismatched family is reported without exposing or
    /// rebuilding the underlying archive index.
    pub(crate) fn from_prepared_package(prepared: PreparedPackage) -> Result<Self> {
        Self::from_prepared_package_with_xml_limits(
            prepared,
            XmlStreamLimits::owner_default_value(),
        )
    }

    pub(crate) fn from_prepared_package_with_xml_limits(
        prepared: PreparedPackage,
        xml_limits: XmlStreamLimits,
    ) -> Result<Self> {
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
        Self::from_family_package(package, xml_limits)
    }

    /// Adopt an already materialized, validated archive without copying its
    /// bytes.  This is the explicit handoff from the positional source owner
    /// to the established mutable package facade.
    pub(crate) fn from_owned_package(archive: OwnedPackage) -> Result<Self> {
        Self::from_owned_package_with_xml_limits(archive, XmlStreamLimits::owner_default_value())
    }

    pub(crate) fn from_owned_package_with_xml_limits(
        archive: OwnedPackage,
        xml_limits: XmlStreamLimits,
    ) -> Result<Self> {
        let package = family::Package::from_owned_package(archive, MIMETYPE, BODY_MARKER, "ODS")?;
        Self::from_family_package(package, xml_limits)
    }

    /// Adopt shared ODS bytes for an internal source-bound transaction.
    #[cfg(test)]
    pub(crate) fn from_shared_bytes(bytes: Arc<Vec<u8>>) -> Result<Self> {
        Self::from_shared_bytes_with_xml_limits(bytes, XmlStreamLimits::owner_default_value())
    }

    pub(crate) fn from_shared_bytes_with_xml_limits(
        bytes: Arc<Vec<u8>>,
        xml_limits: XmlStreamLimits,
    ) -> Result<Self> {
        let package = family::Package::from_shared_bytes(bytes, MIMETYPE, BODY_MARKER, "ODS")?;
        Self::from_family_package(package, xml_limits)
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
        Self::from_family_package(package, self.xml_limits)
    }

    fn from_family_package(package: family::Package, xml_limits: XmlStreamLimits) -> Result<Self> {
        crate::authoring::validate_content_xml(package.content_xml())?;
        Ok(Self {
            inner: Arc::new(package),
            xml_limits,
            resource_count: std::sync::OnceLock::new(),
        })
    }

    #[must_use]
    pub(crate) const fn xml_limits_value(&self) -> XmlStreamLimits {
        self.xml_limits
    }

    #[must_use]
    pub(crate) fn content_xml(&self) -> &str {
        self.inner.content_xml()
    }

    #[must_use]
    pub(crate) fn styles_xml(&self) -> Option<&str> {
        self.inner.styles_xml()
    }

    #[must_use]
    pub(crate) fn package(&self) -> &OwnedPackage {
        self.inner.package()
    }

    pub(crate) fn resource_count(&self) -> Result<usize> {
        if let Some(count) = self.resource_count.get() {
            return Ok(*count);
        }
        // Resource admission is deliberately deferred until the first
        // operation that has a caller-selected resource limit. The family
        // package has already retained its validated indexed archive at this
        // point; opening a package must not allocate a file-name Vec merely
        // to prepare a later edit. Once admitted, later validations reuse the
        // same count.
        let count = self
            .inner
            .files()?
            .into_iter()
            .filter(|path| !crate::document::is_reserved_path(path) && !path.ends_with('/'))
            .count();
        let _ = self.resource_count.set(count);
        Ok(self.resource_count.get().copied().unwrap_or(count))
    }

    #[cfg(test)]
    pub(crate) fn resource_count_cached(&self) -> Option<usize> {
        self.resource_count.get().copied()
    }

    /// Everything a tolerant read repaired while this package was opened.
    ///
    /// An empty report means the package conformed; see
    /// [`litchi_odf_common::leniency::ToleranceReport`].
    #[inline]
    #[must_use]
    pub(crate) fn tolerance_report(&self) -> &litchi_odf_common::leniency::ToleranceReport {
        self.inner.tolerance_report()
    }

    /// Return the identity of the archive index retained by smart detection.
    #[doc(hidden)]
    #[must_use]
    pub(crate) fn prepared_index_identity(&self) -> usize {
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
        crate::metadata::Snapshot::from_source(source, self.xml_limits)
    }

    /// Read the ordered named-definition catalog from `content.xml`.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub(crate) fn definitions(&self) -> Result<Vec<Definition>> {
        crate::codec::names::parse(self.content_xml())
    }

    /// Decode worksheets under the policy retained by the owning package.
    pub(crate) fn sheets_with_xml_limits(
        &self,
        xml_limits: XmlStreamLimits,
    ) -> Result<Vec<crate::worksheet::Sheet>> {
        crate::worksheet::codec::parse_with_xml_limits(self.content_xml(), xml_limits)
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
        let bytes = match replace_package_content_xml(self.package(), content_xml) {
            Ok(bytes) => bytes,
            // A producer prolog newline fails the compact audit of a whole
            // republished part; it carries no content, so drop it and retry.
            Err(error) => match compact_prolog(content_xml) {
                Cow::Owned(compact) => replace_package_content_xml(self.package(), &compact)?,
                Cow::Borrowed(_) => return Err(error),
            },
        };
        Self::from_bytes_with_xml_limits(bytes, self.xml_limits)
    }

    /// Rebuild this package from exact checked `content.xml` source ranges.
    pub(crate) fn replace_spliced_content_xml(
        &self,
        content_xml: &str,
        publication: XmlSplicePublication,
    ) -> Result<Self> {
        let bytes = replace_content_xml_spliced(self.package(), content_xml, publication)?;
        Self::from_bytes_with_xml_limits(bytes, self.xml_limits)
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
        // Every member but `meta.xml` is copied as the exact source bytes, so
        // a producer's `styles.xml` and `settings.xml` are never republished
        // as authored XML.
        let additions = metadata_xml
            .map(|xml| {
                vec![litchi_odf_common::package::Addition {
                    path: "meta.xml".to_string(),
                    bytes: compact_prolog(xml).into_owned().into_bytes(),
                    media_type: "text/xml".to_string(),
                }]
            })
            .unwrap_or_default();
        let bytes = litchi_odf_common::package::rebuild_package(
            self.package(),
            self.content_xml(),
            additions,
            Vec::new(),
            vec!["meta.xml".to_string()],
            Vec::<String>::new(),
        )?;
        Self::from_bytes_with_xml_limits(bytes, self.xml_limits)
    }

    /// Replace the content-level calculation-settings owner atomically.
    pub(crate) fn replace_calculation_settings(&self, settings: Option<&Settings>) -> Result<Self> {
        let snapshot =
            crate::settings::Snapshot::from_content_xml(self.content_xml(), self.xml_limits)?;
        let mut transaction = snapshot.transaction();
        match settings {
            Some(settings) => transaction.replace(settings.clone())?,
            None => transaction.remove(),
        }
        let commit = transaction.commit()?;
        if !commit.changed() {
            return Self::from_bytes_with_xml_limits(
                self.inner.as_bytes().to_vec(),
                self.xml_limits,
            );
        }
        let content_xml = commit.into_owned();
        let bytes = self.rebuild(&content_xml, Part::Preserve, Part::Preserve, Part::Preserve)?;
        Self::from_bytes_with_xml_limits(bytes, self.xml_limits)
    }

    /// Replace or remove the document styles part while preserving every
    /// unrelated package member.
    pub(crate) fn replace_styles_xml(&self, styles_xml: Option<&str>) -> Result<Self> {
        let bytes = self.rebuild(
            self.content_xml(),
            Part::from_option(styles_xml),
            Part::Preserve,
            Part::Preserve,
        )?;
        Self::from_bytes_with_xml_limits(bytes, self.xml_limits)
    }

    fn rebuild(
        &self,
        content_xml: &str,
        styles: Part<'_>,
        metadata: Part<'_>,
        settings: Part<'_>,
    ) -> Result<Vec<u8>> {
        let source = self.package();
        let mut writer =
            PackageWriter::new(litchi_odf_common::core::StreamingArchiveLimits::default());
        writer.set_mimetype(&source.mimetype()?)?;
        writer.add_file("content.xml", compact_prolog(content_xml).as_bytes())?;
        write_part(&mut writer, source, "styles.xml", styles)?;
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
    pub(crate) fn into_bytes(self) -> Vec<u8> {
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
                writer.add_file(path, &compact_prolog_bytes(&source.get_file(path)?))?;
            }
        },
        Part::Set(xml) => writer.add_file(path, compact_prolog(xml).as_bytes())?,
        Part::Remove => {},
    }
    Ok(())
}
