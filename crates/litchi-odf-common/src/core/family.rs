//! Shared ownership for the simple packaged ODF families.

use super::{
    Content, Meta, OwnedPackage, SourcePackageLimits, Styles, package::zeroizing_password,
};
use litchi_core::{Error, Metadata, Result};
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::{fs, path::Path, sync::Arc};

const MAX_CONTENT_BYTES: usize = 256 * 1024 * 1024;

/// A packaged ODF detection result that retains the validated archive index.
///
/// The result is intentionally family-neutral. A concrete family facade must
/// still validate its own MIME type and XML body before constructing a semantic
/// document. Moving this value into that facade reuses the one bounded ZIP
/// central-directory index built during detection.
#[derive(Clone)]
pub struct PreparedPackage {
    package: OwnedPackage,
    format: litchi_core::detection::FileFormat,
}

impl PreparedPackage {
    pub(crate) fn new(package: OwnedPackage, format: litchi_core::detection::FileFormat) -> Self {
        Self { package, format }
    }

    /// Return the format identified from the package MIME type.
    #[must_use]
    pub fn format(&self) -> litchi_core::detection::FileFormat {
        self.format
    }

    /// Borrow the owned package retained by detection.
    #[must_use]
    pub fn package(&self) -> &OwnedPackage {
        &self.package
    }

    /// Consume the prepared result and transfer its indexed package owner.
    #[must_use]
    pub fn into_package(self) -> OwnedPackage {
        self.package
    }

    /// Return the retained index identity for diagnostics and tests.
    #[doc(hidden)]
    #[must_use]
    pub fn prepared_index_identity(&self) -> usize {
        self.package.prepared_index_identity()
    }
}

/// A validated immutable ODF package with the standard XML parts decoded once.
///
/// Concrete family crates retain a small contextual wrapper around this type
/// so MIME and body validation remain visible at their package boundary while
/// archive, content, style, metadata, and file-list ownership stays shared.
pub struct Package {
    archive: OwnedPackage,
    content: Content,
    styles: Option<Styles>,
    metadata: Option<Metadata>,
}

impl Package {
    /// Open a package after validating its MIME type and content root marker.
    ///
    /// # Errors
    ///
    /// Returns an error when the file cannot be read or the package fails
    /// archive, MIME type, or family-content validation.
    pub fn open(
        path: impl AsRef<Path>,
        mimetype: &str,
        body_marker: &str,
        family_name: &str,
    ) -> Result<Self> {
        Self::open_with_limits(
            path,
            SourcePackageLimits::default(),
            mimetype,
            body_marker,
            family_name,
        )
    }

    /// Open a package from a path under explicit finite archive limits.
    pub fn open_with_limits(
        path: impl AsRef<Path>,
        limits: SourcePackageLimits,
        mimetype: &str,
        body_marker: &str,
        family_name: &str,
    ) -> Result<Self> {
        let file = fs::File::open(path)?;
        Self::from_reader_with_limits(file, limits, mimetype, body_marker, family_name)
    }

    /// Open a password-protected package from a path under the default finite
    /// archive limits.
    pub fn open_with_password(
        path: impl AsRef<Path>,
        password: impl Into<String>,
        mimetype: &str,
        body_marker: &str,
        family_name: &str,
    ) -> Result<Self> {
        let password = zeroizing_password(password);
        Self::open_with_zeroizing_password(
            path,
            SourcePackageLimits::default(),
            password,
            mimetype,
            body_marker,
            family_name,
        )
    }

    /// Open a password-protected package from a path with explicit finite
    /// archive limits.
    pub fn open_with_limits_and_password(
        path: impl AsRef<Path>,
        limits: SourcePackageLimits,
        password: impl Into<String>,
        mimetype: &str,
        body_marker: &str,
        family_name: &str,
    ) -> Result<Self> {
        let password = zeroizing_password(password);
        Self::open_with_zeroizing_password(
            path,
            limits,
            password,
            mimetype,
            body_marker,
            family_name,
        )
    }

    fn open_with_zeroizing_password(
        path: impl AsRef<Path>,
        limits: SourcePackageLimits,
        password: zeroize::Zeroizing<String>,
        mimetype: &str,
        body_marker: &str,
        family_name: &str,
    ) -> Result<Self> {
        let file = fs::File::open(path)?;
        Self::from_reader_with_zeroizing_password(
            file,
            limits,
            password,
            mimetype,
            body_marker,
            family_name,
        )
    }

    /// Read a package from a stream under the default finite archive limits.
    pub fn from_reader(
        reader: impl std::io::Read,
        mimetype: &str,
        body_marker: &str,
        family_name: &str,
    ) -> Result<Self> {
        Self::from_reader_with_limits(
            reader,
            SourcePackageLimits::default(),
            mimetype,
            body_marker,
            family_name,
        )
    }

    /// Read a package from a stream under explicit finite archive limits.
    pub fn from_reader_with_limits(
        reader: impl std::io::Read,
        limits: SourcePackageLimits,
        mimetype: &str,
        body_marker: &str,
        family_name: &str,
    ) -> Result<Self> {
        Self::from_owned_package(
            OwnedPackage::from_reader_with_limits(reader, limits)?,
            mimetype,
            body_marker,
            family_name,
        )
    }

    /// Read a password-protected package from a stream under default finite
    /// archive limits.
    pub fn from_reader_with_password(
        reader: impl std::io::Read,
        password: impl Into<String>,
        mimetype: &str,
        body_marker: &str,
        family_name: &str,
    ) -> Result<Self> {
        let password = zeroizing_password(password);
        Self::from_reader_with_zeroizing_password(
            reader,
            SourcePackageLimits::default(),
            password,
            mimetype,
            body_marker,
            family_name,
        )
    }

    /// Read a password-protected package from a stream with explicit finite
    /// archive limits.
    pub fn from_reader_with_limits_and_password(
        reader: impl std::io::Read,
        limits: SourcePackageLimits,
        password: impl Into<String>,
        mimetype: &str,
        body_marker: &str,
        family_name: &str,
    ) -> Result<Self> {
        let password = zeroizing_password(password);
        Self::from_reader_with_zeroizing_password(
            reader,
            limits,
            password,
            mimetype,
            body_marker,
            family_name,
        )
    }

    fn from_reader_with_zeroizing_password(
        reader: impl std::io::Read,
        limits: SourcePackageLimits,
        password: zeroize::Zeroizing<String>,
        mimetype: &str,
        body_marker: &str,
        family_name: &str,
    ) -> Result<Self> {
        Self::from_owned_package(
            OwnedPackage::from_reader_with_zeroizing_password(reader, limits, password)?,
            mimetype,
            body_marker,
            family_name,
        )
    }

    /// Decode a package after validating its MIME type and content root marker.
    ///
    /// # Errors
    ///
    /// Returns an error when the archive, MIME type, or family content is
    /// invalid.
    pub fn from_bytes(
        bytes: Vec<u8>,
        mimetype: &str,
        body_marker: &str,
        family_name: &str,
    ) -> Result<Self> {
        Self::from_bytes_with_limits(
            bytes,
            SourcePackageLimits::default(),
            mimetype,
            body_marker,
            family_name,
        )
    }

    /// Decode a package from bytes under explicit finite archive limits.
    pub fn from_bytes_with_limits(
        bytes: Vec<u8>,
        limits: SourcePackageLimits,
        mimetype: &str,
        body_marker: &str,
        family_name: &str,
    ) -> Result<Self> {
        Self::from_owned_package(
            OwnedPackage::from_bytes_with_limits(bytes, limits)?,
            mimetype,
            body_marker,
            family_name,
        )
    }

    /// Decode shared package bytes without copying the archive buffer.
    ///
    /// # Errors
    ///
    /// Returns an error when the archive, MIME type, or family content is
    /// invalid.
    pub fn from_shared_bytes(
        bytes: Arc<Vec<u8>>,
        mimetype: &str,
        body_marker: &str,
        family_name: &str,
    ) -> Result<Self> {
        Self::from_shared_bytes_with_limits(
            bytes,
            SourcePackageLimits::default(),
            mimetype,
            body_marker,
            family_name,
        )
    }

    /// Decode shared package bytes under explicit finite archive limits.
    pub fn from_shared_bytes_with_limits(
        bytes: Arc<Vec<u8>>,
        limits: SourcePackageLimits,
        mimetype: &str,
        body_marker: &str,
        family_name: &str,
    ) -> Result<Self> {
        Self::from_owned_package(
            OwnedPackage::from_shared_bytes_with_limits(bytes, limits)?,
            mimetype,
            body_marker,
            family_name,
        )
    }

    /// Decode a password-encrypted package after validating its MIME type and
    /// content root marker.
    ///
    /// # Errors
    ///
    /// Returns an error when the archive, password-protected entries, MIME
    /// type, or family content is invalid.
    pub fn from_bytes_with_password(
        bytes: Vec<u8>,
        password: impl Into<String>,
        mimetype: &str,
        body_marker: &str,
        family_name: &str,
    ) -> Result<Self> {
        let password = zeroizing_password(password);
        Self::from_bytes_with_zeroizing_password(
            bytes,
            SourcePackageLimits::default(),
            password,
            mimetype,
            body_marker,
            family_name,
        )
    }

    /// Decode a password-protected package from bytes with explicit finite
    /// archive limits.
    pub fn from_bytes_with_limits_and_password(
        bytes: Vec<u8>,
        limits: SourcePackageLimits,
        password: impl Into<String>,
        mimetype: &str,
        body_marker: &str,
        family_name: &str,
    ) -> Result<Self> {
        let password = zeroizing_password(password);
        Self::from_bytes_with_zeroizing_password(
            bytes,
            limits,
            password,
            mimetype,
            body_marker,
            family_name,
        )
    }

    fn from_bytes_with_zeroizing_password(
        bytes: Vec<u8>,
        limits: SourcePackageLimits,
        password: zeroize::Zeroizing<String>,
        mimetype: &str,
        body_marker: &str,
        family_name: &str,
    ) -> Result<Self> {
        Self::from_owned_package(
            OwnedPackage::from_bytes_with_zeroizing_password(bytes, limits, password)?,
            mimetype,
            body_marker,
            family_name,
        )
    }

    /// Adopt an already parsed archive without reparsing its ZIP structure.
    ///
    /// # Errors
    ///
    /// Returns an error when the package MIME type, XML parts, or family
    /// content violates the supplied contract.
    pub fn from_owned_package(
        archive: OwnedPackage,
        mimetype: &str,
        body_marker: &str,
        family_name: &str,
    ) -> Result<Self> {
        let found = archive.mimetype()?;
        if found != mimetype {
            return Err(Error::InvalidFormat(format!(
                "expected {family_name} package MIME type '{mimetype}', found '{found}'"
            )));
        }

        let content_bytes = archive.get_file("content.xml")?;
        let content_xml = std::str::from_utf8(&content_bytes).map_err(|error| {
            Error::InvalidFormat(format!("{family_name} content.xml is not UTF-8: {error}"))
        })?;
        validate_content_part(content_xml, body_marker, family_name)?;
        let content = Content::from_bytes(&content_bytes)?;

        let styles = archive
            .has_file("styles.xml")?
            .then(|| archive.get_file("styles.xml"))
            .transpose()?
            .map(|bytes| Styles::from_bytes(&bytes))
            .transpose()?;
        let metadata = archive
            .has_file("meta.xml")?
            .then(|| archive.get_file("meta.xml"))
            .transpose()?
            .map(|bytes| Meta::from_bytes(&bytes))
            .transpose()?
            .map(|meta| meta.try_extract_metadata())
            .transpose()?;

        Ok(Self {
            archive,
            content,
            styles,
            metadata,
        })
    }

    /// Return the decoded content XML.
    #[must_use]
    pub fn content_xml(&self) -> &str {
        self.content.xml_content()
    }

    /// Return the optional decoded styles XML.
    #[must_use]
    pub fn styles_xml(&self) -> Option<&str> {
        self.styles.as_ref().map(Styles::xml_content)
    }

    /// Return the optional common metadata snapshot.
    #[must_use]
    pub fn metadata(&self) -> Option<&Metadata> {
        self.metadata.as_ref()
    }

    /// Borrow the owned package for family-specific package edits.
    #[must_use]
    pub fn package(&self) -> &OwnedPackage {
        &self.archive
    }

    /// Borrow the original archive bytes without allocating.
    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        self.archive.as_bytes()
    }

    /// Clone the shared handle to the exact archive allocation.
    #[must_use]
    pub fn shared_bytes(&self) -> Arc<Vec<u8>> {
        self.archive.shared_bytes()
    }

    /// List all safe package paths.
    ///
    /// # Errors
    ///
    /// Returns an error when the archive file list cannot be read.
    pub fn files(&self) -> Result<Vec<String>> {
        self.archive.files()
    }

    /// Consume the snapshot and return the original package bytes.
    #[must_use]
    pub fn into_bytes(self) -> Vec<u8> {
        self.archive.into_inner()
    }
}

/// Validate the bounded, family-specific `content.xml` contract shared by
/// package facades and detached family builders.
///
/// # Errors
///
/// Returns an error when the XML exceeds the family size limit or does not
/// contain the expected family body marker.
pub fn validate_content_part(xml: &str, body_marker: &str, family_name: &str) -> Result<()> {
    if xml.len() > MAX_CONTENT_BYTES {
        return Err(Error::InvalidFormat(format!(
            "{family_name} content.xml exceeds the family limit"
        )));
    }
    if !xml.contains(body_marker) {
        return Err(Error::InvalidFormat(format!(
            "{family_name} content.xml has no expected body"
        )));
    }
    Ok(())
}

/// Validate a complete ODF `content.xml` root, body, and family element.
pub fn validate_content_document_part(
    xml: &str,
    body_marker: &str,
    family_name: &str,
) -> Result<()> {
    if xml.len() > MAX_CONTENT_BYTES {
        return Err(Error::InvalidFormat(format!(
            "{family_name} content.xml exceeds the family limit"
        )));
    }
    let expected_local = body_marker
        .strip_prefix("<office:")
        .and_then(|marker| {
            marker
                .split(|character: char| !character.is_ascii_alphanumeric() && character != '-')
                .next()
        })
        .filter(|local| !local.is_empty())
        .ok_or_else(|| {
            Error::InvalidFormat(format!("{family_name} content.xml has no expected body"))
        })?;
    const OFFICE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    reader.config_mut().check_comments = true;
    let mut depth = 0usize;
    let mut root_closed = false;
    let mut body_seen = false;
    let mut expected_seen = false;
    let mut in_body = false;
    let mut declaration_seen = false;
    let mut first_event = true;

    loop {
        // Borrowing read: events borrow `xml` directly, avoiding the
        // per-event buffer copy of the `_into` API. Binding push/pop still
        // runs inside `read_event` for every event, so prefix-rebinding
        // semantics and the tokenization error stream are unchanged.
        let event = reader.read_event().map_err(|error| {
            Error::InvalidFormat(format!("invalid {family_name} content.xml: {error}"))
        })?;
        // The resolved namespace is consumed only by the Start/Empty arms at
        // depth <= 2 below (root, `office:body`, `office:forms`/family
        // element); deeper arms use `local_name()` only, and bindings declared
        // at depth >= 3 scope just their own subtree, so no consumed
        // resolution can change. Resolve only where the result is observable.
        let office = match &event {
            Event::Start(element) | Event::Empty(element) if depth <= 2 => {
                let (namespace, _) = reader.resolver().resolve_element(element.name());
                matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == OFFICE_NAMESPACE)
            },
            _ => false,
        };
        match event {
            Event::Start(element) => {
                if root_closed {
                    return Err(Error::InvalidFormat(format!(
                        "{family_name} content.xml has content after its root"
                    )));
                }
                let local = element.local_name();
                match depth {
                    0 if office && local.as_ref() == b"document-content" => depth = 1,
                    0 => {
                        return Err(Error::InvalidFormat(format!(
                            "{family_name} content.xml has the wrong root"
                        )));
                    },
                    1 if office && local.as_ref() == b"body" && !body_seen => {
                        body_seen = true;
                        in_body = true;
                        depth = 2;
                    },
                    1 if office && local.as_ref() == b"body" => {
                        return Err(Error::InvalidFormat(format!(
                            "{family_name} content.xml has duplicate office:body"
                        )));
                    },
                    1 => {
                        depth = depth.checked_add(1).ok_or_else(|| {
                            Error::InvalidFormat(format!(
                                "{family_name} content.xml nesting overflows"
                            ))
                        })?
                    },
                    2 if in_body && office && local.as_ref() == b"forms" && !expected_seen => {
                        depth = 3;
                    },
                    2 if in_body
                        && office
                        && local.as_ref() == expected_local.as_bytes()
                        && !expected_seen =>
                    {
                        expected_seen = true;
                        depth = 3;
                    },
                    2 if in_body => {
                        return Err(Error::InvalidFormat(format!(
                            "{family_name} content.xml has the wrong office body"
                        )));
                    },
                    _ => {
                        depth = depth.checked_add(1).ok_or_else(|| {
                            Error::InvalidFormat(format!(
                                "{family_name} content.xml nesting overflows"
                            ))
                        })?
                    },
                }
            },
            Event::Empty(element) => {
                if root_closed || depth == 0 {
                    return Err(Error::InvalidFormat(format!(
                        "{family_name} content.xml has an invalid empty root"
                    )));
                }
                let local = element.local_name();
                if depth == 1 {
                    if office && local.as_ref() == b"body" && !body_seen {
                        body_seen = true;
                    } else if office && local.as_ref() == b"body" {
                        return Err(Error::InvalidFormat(format!(
                            "{family_name} content.xml has duplicate office:body"
                        )));
                    }
                } else if in_body
                    && depth == 2
                    && office
                    && local.as_ref() == b"forms"
                    && !expected_seen
                {
                    // `office:forms` may precede the family body.
                } else if in_body
                    && depth == 2
                    && office
                    && local.as_ref() == expected_local.as_bytes()
                    && !expected_seen
                {
                    expected_seen = true;
                } else if in_body && depth == 2 {
                    return Err(Error::InvalidFormat(format!(
                        "{family_name} content.xml has the wrong office body"
                    )));
                }
            },
            Event::End(_) => {
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat(format!("{family_name} content.xml has an unexpected end"))
                })?;
                if depth == 0 {
                    root_closed = true;
                } else if in_body && depth == 1 {
                    in_body = false;
                }
            },
            Event::Text(text) => {
                if (depth == 0 || root_closed) && !text.iter().all(u8::is_ascii_whitespace) {
                    return Err(Error::InvalidFormat(format!(
                        "{family_name} content.xml has unexpected text outside its root"
                    )));
                }
            },
            Event::CData(_) | Event::GeneralRef(_) if depth == 0 || root_closed => {
                return Err(Error::InvalidFormat(format!(
                    "{family_name} content.xml has content outside its root"
                )));
            },
            Event::DocType(_) => {
                return Err(Error::InvalidFormat(format!(
                    "{family_name} content.xml must not contain a doctype"
                )));
            },
            Event::GeneralRef(reference) if !crate::validation::valid_xml_reference(&reference) => {
                return Err(Error::InvalidFormat(format!(
                    "{family_name} content.xml has an invalid character or entity reference"
                )));
            },
            Event::Decl(_) if declaration_seen || !first_event || depth != 0 || root_closed => {
                return Err(Error::InvalidFormat(format!(
                    "{family_name} content.xml has an XML declaration outside its prologue"
                )));
            },
            Event::Decl(_) => declaration_seen = true,
            Event::Eof => break,
            Event::Comment(_) | Event::PI(_) | Event::CData(_) | Event::GeneralRef(_) => {},
        }
        first_event = false;
    }

    if !root_closed || depth != 0 || !body_seen || !expected_seen {
        return Err(Error::InvalidFormat(format!(
            "{family_name} content.xml has no complete expected body"
        )));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::{OwnedPackage, Package, validate_content_document_part, validate_content_part};
    use std::io::{Cursor, Write};

    const MIMETYPE: &str = "application/vnd.oasis.opendocument.presentation";
    const CONTENT: &str = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:body><office:presentation/></office:body></office:document-content>"#;

    type TestResult = Result<(), Box<dyn std::error::Error>>;

    fn package() -> Result<Vec<u8>, Box<dyn std::error::Error>> {
        let mut output = Cursor::new(Vec::new());
        let mut zip = zip::ZipWriter::new(&mut output);
        let options = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Stored);
        zip.start_file("mimetype", options)?;
        zip.write_all(MIMETYPE.as_bytes())?;
        zip.start_file("META-INF/manifest.xml", options)?;
        zip.write_all(
            format!(
                r#"<m:manifest xmlns:m="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><m:file-entry m:full-path="/" m:media-type="{MIMETYPE}"/><m:file-entry m:full-path="content.xml" m:media-type="text/xml"/><m:file-entry m:full-path="styles.xml" m:media-type="text/xml"/></m:manifest>"#
            )
            .as_bytes(),
        )
        ?;
        zip.start_file("content.xml", options)?;
        zip.write_all(CONTENT.as_bytes())?;
        zip.start_file("styles.xml", options)?;
        zip.write_all(b"<office:document-styles xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\"/>")
            ?;
        zip.finish()?;
        Ok(output.into_inner())
    }

    #[test]
    fn validates_and_reuses_shared_package_state() -> TestResult {
        let bytes = package()?;
        let value = Package::from_bytes(bytes.clone(), MIMETYPE, "<office:presentation", "ODP")?;
        assert_eq!(value.content_xml(), CONTENT);
        assert!(value.styles_xml().is_some());
        assert_eq!(value.as_bytes(), bytes.as_slice());
        assert!(value.files()?.contains(&"content.xml".to_string()));

        let owned = OwnedPackage::from_bytes(bytes.clone())?;
        let adopted = Package::from_owned_package(owned, MIMETYPE, "<office:presentation", "ODP")?;
        assert_eq!(adopted.as_bytes(), bytes.as_slice());
        Ok(())
    }

    #[test]
    fn rejects_wrong_mime_and_body_marker() -> TestResult {
        let bytes = package()?;
        assert!(
            Package::from_bytes(bytes.clone(), "text/plain", "<office:presentation", "ODP")
                .is_err()
        );
        assert!(Package::from_bytes(bytes, MIMETYPE, "<office:text", "ODP").is_err());
        Ok(())
    }

    #[test]
    fn validates_detached_content_with_the_same_family_contract() -> TestResult {
        assert!(validate_content_part("<office:drawing/>", "<office:drawing", "ODG").is_ok());
        let error = match validate_content_part("<office:text/>", "<office:drawing", "ODG") {
            Err(error) => error.to_string(),
            Ok(()) => {
                return Err(std::io::Error::other("expected family validation failure").into());
            },
        };
        assert!(error.contains("ODG content.xml has no expected body"));
        Ok(())
    }

    /// Pre-0220 oracle: the historical buffered, fully-resolved
    /// implementation of [`validate_content_document_part`], retained
    /// byte-identically to cross-check the borrowing, depth-gated
    /// implementation on both the synthetic edge cases and the fixture
    /// corpus.
    fn validate_content_document_part_oracle(
        xml: &str,
        body_marker: &str,
        family_name: &str,
    ) -> litchi_core::Result<()> {
        use litchi_core::Error;
        use quick_xml::events::Event;
        use quick_xml::name::{Namespace, ResolveResult};
        use quick_xml::reader::NsReader;

        if xml.len() > super::MAX_CONTENT_BYTES {
            return Err(Error::InvalidFormat(format!(
                "{family_name} content.xml exceeds the family limit"
            )));
        }
        let expected_local = body_marker
            .strip_prefix("<office:")
            .and_then(|marker| {
                marker
                    .split(|character: char| !character.is_ascii_alphanumeric() && character != '-')
                    .next()
            })
            .filter(|local| !local.is_empty())
            .ok_or_else(|| {
                Error::InvalidFormat(format!("{family_name} content.xml has no expected body"))
            })?;
        const OFFICE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
        let mut reader = NsReader::from_str(xml);
        reader.config_mut().check_end_names = true;
        reader.config_mut().check_comments = true;
        let mut buffer = Vec::new();
        let mut depth = 0usize;
        let mut root_closed = false;
        let mut body_seen = false;
        let mut expected_seen = false;
        let mut in_body = false;
        let mut declaration_seen = false;
        let mut first_event = true;

        loop {
            let (namespace, event) =
                reader
                    .read_resolved_event_into(&mut buffer)
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid {family_name} content.xml: {error}"))
                    })?;
            let office = matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == OFFICE_NAMESPACE);
            match event {
                Event::Start(element) => {
                    if root_closed {
                        return Err(Error::InvalidFormat(format!(
                            "{family_name} content.xml has content after its root"
                        )));
                    }
                    let local = element.local_name();
                    match depth {
                        0 if office && local.as_ref() == b"document-content" => depth = 1,
                        0 => {
                            return Err(Error::InvalidFormat(format!(
                                "{family_name} content.xml has the wrong root"
                            )));
                        },
                        1 if office && local.as_ref() == b"body" && !body_seen => {
                            body_seen = true;
                            in_body = true;
                            depth = 2;
                        },
                        1 if office && local.as_ref() == b"body" => {
                            return Err(Error::InvalidFormat(format!(
                                "{family_name} content.xml has duplicate office:body"
                            )));
                        },
                        1 => {
                            depth = depth.checked_add(1).ok_or_else(|| {
                                Error::InvalidFormat(format!(
                                    "{family_name} content.xml nesting overflows"
                                ))
                            })?
                        },
                        2 if in_body && office && local.as_ref() == b"forms" && !expected_seen => {
                            depth = 3;
                        },
                        2 if in_body
                            && office
                            && local.as_ref() == expected_local.as_bytes()
                            && !expected_seen =>
                        {
                            expected_seen = true;
                            depth = 3;
                        },
                        2 if in_body => {
                            return Err(Error::InvalidFormat(format!(
                                "{family_name} content.xml has the wrong office body"
                            )));
                        },
                        _ => {
                            depth = depth.checked_add(1).ok_or_else(|| {
                                Error::InvalidFormat(format!(
                                    "{family_name} content.xml nesting overflows"
                                ))
                            })?
                        },
                    }
                },
                Event::Empty(element) => {
                    if root_closed || depth == 0 {
                        return Err(Error::InvalidFormat(format!(
                            "{family_name} content.xml has an invalid empty root"
                        )));
                    }
                    let local = element.local_name();
                    if depth == 1 {
                        if office && local.as_ref() == b"body" && !body_seen {
                            body_seen = true;
                        } else if office && local.as_ref() == b"body" {
                            return Err(Error::InvalidFormat(format!(
                                "{family_name} content.xml has duplicate office:body"
                            )));
                        }
                    } else if in_body
                        && depth == 2
                        && office
                        && local.as_ref() == b"forms"
                        && !expected_seen
                    {
                        // `office:forms` may precede the family body.
                    } else if in_body
                        && depth == 2
                        && office
                        && local.as_ref() == expected_local.as_bytes()
                        && !expected_seen
                    {
                        expected_seen = true;
                    } else if in_body && depth == 2 {
                        return Err(Error::InvalidFormat(format!(
                            "{family_name} content.xml has the wrong office body"
                        )));
                    }
                },
                Event::End(_) => {
                    depth = depth.checked_sub(1).ok_or_else(|| {
                        Error::InvalidFormat(format!(
                            "{family_name} content.xml has an unexpected end"
                        ))
                    })?;
                    if depth == 0 {
                        root_closed = true;
                    } else if in_body && depth == 1 {
                        in_body = false;
                    }
                },
                Event::Text(text) => {
                    if (depth == 0 || root_closed) && !text.iter().all(u8::is_ascii_whitespace) {
                        return Err(Error::InvalidFormat(format!(
                            "{family_name} content.xml has unexpected text outside its root"
                        )));
                    }
                },
                Event::CData(_) | Event::GeneralRef(_) if depth == 0 || root_closed => {
                    return Err(Error::InvalidFormat(format!(
                        "{family_name} content.xml has content outside its root"
                    )));
                },
                Event::DocType(_) => {
                    return Err(Error::InvalidFormat(format!(
                        "{family_name} content.xml must not contain a doctype"
                    )));
                },
                Event::GeneralRef(reference)
                    if !crate::validation::valid_xml_reference(&reference) =>
                {
                    return Err(Error::InvalidFormat(format!(
                        "{family_name} content.xml has an invalid character or entity reference"
                    )));
                },
                Event::Decl(_) if declaration_seen || !first_event || depth != 0 || root_closed => {
                    return Err(Error::InvalidFormat(format!(
                        "{family_name} content.xml has an XML declaration outside its prologue"
                    )));
                },
                Event::Decl(_) => declaration_seen = true,
                Event::Eof => break,
                Event::Comment(_) | Event::PI(_) | Event::CData(_) | Event::GeneralRef(_) => {},
            }
            first_event = false;
            buffer.clear();
        }

        if !root_closed || depth != 0 || !body_seen || !expected_seen {
            return Err(Error::InvalidFormat(format!(
                "{family_name} content.xml has no complete expected body"
            )));
        }
        Ok(())
    }

    const TEXT_MARKER: &str = "<office:text";
    const OFFICE_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";

    fn document(inner: &str) -> String {
        format!(
            r#"<office:document-content xmlns:office="{OFFICE_NS}">{inner}</office:document-content>"#
        )
    }

    fn assert_validator_parity(label: &str, xml: &str, body_marker: &str, family_name: &str) {
        let expected = validate_content_document_part_oracle(xml, body_marker, family_name)
            .map_err(|error| error.to_string());
        let actual = validate_content_document_part(xml, body_marker, family_name)
            .map_err(|error| error.to_string());
        assert_eq!(
            expected, actual,
            "{label}: borrowing validator and buffered oracle disagree"
        );
    }

    #[test]
    fn borrowing_validator_matches_oracle_on_synthetic_edge_cases() {
        let text_body = r#"<office:body><office:text/></office:body>"#;
        let cases: Vec<(&str, String, &str, &str)> = vec![
            (
                "minimal-empty-family",
                document(text_body),
                TEXT_MARKER,
                "ODT",
            ),
            (
                "family-start-end",
                document(r#"<office:body><office:text></office:text></office:body>"#),
                TEXT_MARKER,
                "ODT",
            ),
            (
                "xml-decl-prologue",
                format!(
                    r#"<?xml version="1.0" encoding="UTF-8"?>{}"#,
                    document(text_body)
                ),
                TEXT_MARKER,
                "ODT",
            ),
            (
                "comment-and-pi-prologue",
                format!(r#"<!--c--><?p i?>{}"#, document(text_body)),
                TEXT_MARKER,
                "ODT",
            ),
            (
                "forms-before-family",
                document(r#"<office:body><office:forms/><office:text/></office:body>"#),
                TEXT_MARKER,
                "ODT",
            ),
            (
                "forms-start-end-before-family",
                document(
                    r#"<office:body><office:forms></office:forms><office:text/></office:body>"#,
                ),
                TEXT_MARKER,
                "ODT",
            ),
            (
                "non-office-element-at-depth-one",
                document(r#"<office:meta/><office:body><office:text/></office:body>"#),
                TEXT_MARKER,
                "ODT",
            ),
            (
                "deep-prefix-rebinding-accepted",
                document(
                    r#"<office:body><office:text><text:p xmlns:office="urn:evil"><office:annotation/></text:p></office:text></office:body>"#,
                ),
                TEXT_MARKER,
                "ODT",
            ),
            (
                "unknown-prefix-deep-accepted",
                document(
                    r#"<office:body><office:text><text:p><weird:thing/></text:p></office:text></office:body>"#,
                ),
                TEXT_MARKER,
                "ODT",
            ),
            (
                "default-namespace-binding-accepted",
                format!(
                    r#"<document-content xmlns="{OFFICE_NS}"><body><text/></body></document-content>"#
                ),
                TEXT_MARKER,
                "ODT",
            ),
            (
                "aliased-prefix-binding-accepted",
                format!(
                    r#"<x:document-content xmlns:x="{OFFICE_NS}"><x:body><x:text/></x:body></x:document-content>"#
                ),
                TEXT_MARKER,
                "ODT",
            ),
            (
                "rebinding-on-body-hides-it",
                document(text_body)
                    .replace("<office:body>", r#"<office:body xmlns:office="urn:evil">"#),
                TEXT_MARKER,
                "ODT",
            ),
            (
                "rebinding-on-empty-body-hides-it",
                document(r#"<office:body xmlns:office="urn:evil"/><office:text/>"#),
                TEXT_MARKER,
                "ODT",
            ),
            (
                "rebinding-on-family-element-rejected",
                document(text_body).replace(
                    "<office:text/>",
                    r#"<office:text xmlns:office="urn:evil"/>"#,
                ),
                TEXT_MARKER,
                "ODT",
            ),
            (
                "cdata-and-reference-inside-body",
                document(
                    r#"<office:body><office:text><text:p><![CDATA[x]]>&amp;&#x9;</text:p></office:text></office:body>"#,
                ),
                TEXT_MARKER,
                "ODT",
            ),
            (
                "whitespace-outside-root-accepted",
                format!("  \n{} \t ", document(text_body)),
                TEXT_MARKER,
                "ODT",
            ),
            ("empty-input", String::new(), TEXT_MARKER, "ODT"),
            (
                "wrong-root",
                document(text_body).replace("document-content", "document"),
                TEXT_MARKER,
                "ODT",
            ),
            (
                "unknown-prefix-at-root",
                document(text_body).replace(
                    r#" xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0""#,
                    "",
                ),
                TEXT_MARKER,
                "ODT",
            ),
            (
                "duplicate-body",
                document(r#"<office:body/><office:body><office:text/></office:body>"#),
                TEXT_MARKER,
                "ODT",
            ),
            (
                "duplicate-empty-body",
                document(r#"<office:body><office:text/></office:body><office:body/>"#),
                TEXT_MARKER,
                "ODT",
            ),
            (
                "wrong-office-body",
                document(r#"<office:body><office:presentation/></office:body>"#),
                TEXT_MARKER,
                "ODT",
            ),
            (
                "family-then-forms",
                document(r#"<office:body><office:text/><office:forms/></office:body>"#),
                TEXT_MARKER,
                "ODT",
            ),
            (
                "duplicate-family-element",
                document(r#"<office:body><office:text/><office:text/></office:body>"#),
                TEXT_MARKER,
                "ODT",
            ),
            ("missing-body", document(""), TEXT_MARKER, "ODT"),
            (
                "body-without-family",
                document(r#"<office:body/>"#),
                TEXT_MARKER,
                "ODT",
            ),
            (
                "self-closing-root",
                format!(r#"<office:document-content xmlns:office="{OFFICE_NS}"/>"#),
                TEXT_MARKER,
                "ODT",
            ),
            (
                "text-before-root",
                format!("hello{}", document(text_body)),
                TEXT_MARKER,
                "ODT",
            ),
            (
                "text-after-root",
                format!("{}tail", document(text_body)),
                TEXT_MARKER,
                "ODT",
            ),
            (
                "cdata-outside-root",
                format!("<![CDATA[x]]>{}", document(text_body)),
                TEXT_MARKER,
                "ODT",
            ),
            (
                "reference-outside-root",
                format!("&amp;{}", document(text_body)),
                TEXT_MARKER,
                "ODT",
            ),
            (
                "doctype-rejected",
                format!("<!DOCTYPE office:document-content>{}", document(text_body)),
                TEXT_MARKER,
                "ODT",
            ),
            (
                "invalid-general-reference",
                document(
                    r#"<office:body><office:text><text:p>&nosuch;</text:p></office:text></office:body>"#,
                ),
                TEXT_MARKER,
                "ODT",
            ),
            (
                "invalid-character-reference",
                document(
                    r#"<office:body><office:text><text:p>&#xD800;</text:p></office:text></office:body>"#,
                ),
                TEXT_MARKER,
                "ODT",
            ),
            (
                "content-after-root",
                format!("{}<extra></extra>", document(text_body)),
                TEXT_MARKER,
                "ODT",
            ),
            (
                "unclosed-elements-at-eof",
                format!(r#"<office:document-content xmlns:office="{OFFICE_NS}"><office:body>"#),
                TEXT_MARKER,
                "ODT",
            ),
            (
                "mismatched-end-tag-deep",
                document(
                    r#"<office:body><office:text><text:p></text:q></office:text></office:body>"#,
                ),
                TEXT_MARKER,
                "ODT",
            ),
            (
                "mismatched-end-tag-shallow",
                format!(
                    r#"<office:document-content xmlns:office="{OFFICE_NS}"><office:body></office:document-content>"#
                ),
                TEXT_MARKER,
                "ODT",
            ),
            (
                "truncated-tag-at-eof",
                format!("{}<text:p", document(text_body)),
                TEXT_MARKER,
                "ODT",
            ),
            (
                "double-hyphen-comment",
                document(
                    r#"<office:body><office:text><!-- a -- b --></office:text></office:body>"#,
                ),
                TEXT_MARKER,
                "ODT",
            ),
            (
                "decl-after-whitespace",
                format!(r#" <?xml version="1.0"?>{}"#, document(text_body)),
                TEXT_MARKER,
                "ODT",
            ),
            (
                "duplicate-decl",
                format!(
                    r#"<?xml version="1.0"?><?xml version="1.0"?>{}"#,
                    document(text_body)
                ),
                TEXT_MARKER,
                "ODT",
            ),
            (
                "stray-end-tag",
                "</office:document-content>".to_string(),
                TEXT_MARKER,
                "ODT",
            ),
            (
                "other-family-marker",
                document(r#"<office:body><office:table/></office:body>"#),
                "<office:table",
                "ODS",
            ),
            (
                "other-family-mismatch",
                document(r#"<office:body><office:text/></office:body>"#),
                "<office:drawing",
                "ODG",
            ),
        ];
        for (label, xml, body_marker, family_name) in &cases {
            assert_validator_parity(label, xml, body_marker, family_name);
        }
    }

    #[test]
    fn validator_error_messages_stay_pinned() {
        let text_body = r#"<office:body><office:text/></office:body>"#;
        let prefix = "Invalid format: ";
        let cases: Vec<(&str, String, &str)> = vec![
            (
                "wrong-root",
                document(text_body).replace("document-content", "document"),
                "ODT content.xml has the wrong root",
            ),
            (
                "duplicate-body",
                document(r#"<office:body/><office:body><office:text/></office:body>"#),
                "ODT content.xml has duplicate office:body",
            ),
            (
                "wrong-office-body",
                document(r#"<office:body><office:presentation/></office:body>"#),
                "ODT content.xml has the wrong office body",
            ),
            (
                "self-closing-root",
                format!(r#"<office:document-content xmlns:office="{OFFICE_NS}"/>"#),
                "ODT content.xml has an invalid empty root",
            ),
            (
                "text-before-root",
                format!("hello{}", document(text_body)),
                "ODT content.xml has unexpected text outside its root",
            ),
            (
                "cdata-outside-root",
                format!("<![CDATA[x]]>{}", document(text_body)),
                "ODT content.xml has content outside its root",
            ),
            (
                "doctype",
                format!("<!DOCTYPE office:document-content>{}", document(text_body)),
                "ODT content.xml must not contain a doctype",
            ),
            (
                "invalid-reference",
                document(
                    r#"<office:body><office:text><text:p>&nosuch;</text:p></office:text></office:body>"#,
                ),
                "ODT content.xml has an invalid character or entity reference",
            ),
            (
                "decl-outside-prologue",
                format!(r#" <?xml version="1.0"?>{}"#, document(text_body)),
                "ODT content.xml has an XML declaration outside its prologue",
            ),
            (
                "missing-body",
                document(""),
                "ODT content.xml has no complete expected body",
            ),
            (
                "content-after-root",
                format!("{}<extra></extra>", document(text_body)),
                "ODT content.xml has content after its root",
            ),
            (
                "stray-end-tag",
                "</office:document-content>".to_string(),
                "invalid ODT content.xml: ill-formed document: close tag \
                 `</office:document-content>` does not match any open tag",
            ),
        ];
        for (label, xml, expected) in &cases {
            let actual = match validate_content_document_part(xml, TEXT_MARKER, "ODT") {
                Err(error) => error.to_string(),
                Ok(()) => panic!("{label}: expected validation failure"),
            };
            let expected = format!("{prefix}{expected}");
            assert_eq!(&actual, &expected, "{label}: error message drifted");
        }
        // Tokenizer failures keep the quick_xml mapping prefix.
        let malformed =
            document(r#"<office:body><office:text><text:p></text:q></office:text></office:body>"#);
        match validate_content_document_part(&malformed, TEXT_MARKER, "ODT") {
            Err(error) => assert!(
                error
                    .to_string()
                    .starts_with("Invalid format: invalid ODT content.xml: "),
                "tokenizer error lost its mapping: {error}"
            ),
            Ok(()) => panic!("mismatched end tag accepted"),
        }
    }

    fn odf_corpus_files() -> Vec<std::path::PathBuf> {
        let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../..")
            .join("test-data");
        let mut files = Vec::new();
        collect_odf(&root, &mut files);
        files.sort();
        files
    }

    fn collect_odf(directory: &std::path::Path, files: &mut Vec<std::path::PathBuf>) {
        let Ok(entries) = std::fs::read_dir(directory) else {
            return;
        };
        for entry in entries.flatten() {
            let path = entry.path();
            if path.is_dir() {
                collect_odf(&path, files);
            } else if path
                .extension()
                .and_then(|e| e.to_str())
                .is_some_and(|extension| {
                    matches!(
                        extension,
                        "odt"
                            | "ods"
                            | "odp"
                            | "odg"
                            | "ott"
                            | "ots"
                            | "otp"
                            | "otg"
                            | "fodt"
                            | "fods"
                            | "fodp"
                            | "fodg"
                    )
                })
            {
                files.push(path);
            }
        }
    }

    fn corpus_content_xml(path: &std::path::Path) -> Option<(String, &'static str, &'static str)> {
        let extension = path.extension()?.to_str()?;
        let (flat, body_marker, family_name) = match extension {
            "odt" | "ott" => (false, "<office:text", "ODT"),
            "fodt" => (true, "<office:text", "ODT"),
            "ods" | "ots" => (false, "<office:table", "ODS"),
            "fods" => (true, "<office:table", "ODS"),
            "odp" | "otp" => (false, "<office:presentation", "ODP"),
            "fodp" => (true, "<office:presentation", "ODP"),
            "odg" | "otg" => (false, "<office:drawing", "ODG"),
            "fodg" => (true, "<office:drawing", "ODG"),
            _ => return None,
        };
        let bytes = std::fs::read(path).ok()?;
        if flat {
            return String::from_utf8(bytes)
                .ok()
                .map(|xml| (xml, body_marker, family_name));
        }
        let reader = soapberry_zip::office::ArchiveReader::new(&bytes).ok()?;
        let entry = reader.read("content.xml").ok()?;
        String::from_utf8(entry)
            .ok()
            .map(|xml| (xml, body_marker, family_name))
    }

    #[test]
    fn borrowing_validator_matches_oracle_on_odf_corpus() {
        let files = odf_corpus_files();
        assert!(!files.is_empty(), "no ODF corpus fixtures discovered");
        let mut compared = 0usize;
        for path in &files {
            let Some((xml, body_marker, family_name)) = corpus_content_xml(path) else {
                continue;
            };
            assert_validator_parity(&path.display().to_string(), &xml, body_marker, family_name);
            compared += 1;
        }
        assert!(compared > 0, "no ODF corpus fixtures yielded content.xml");
    }
}
