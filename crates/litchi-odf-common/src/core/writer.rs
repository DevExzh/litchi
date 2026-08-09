//! ODF package writing functionality.
//!
//! This module provides utilities for creating and writing ODF files as ZIP archives,
//! including generating manifests and proper file structure.
//!
//! Uses soapberry-zip for high-performance ZIP writing.

use base64::{Engine as _, engine::general_purpose::STANDARD as BASE64_STANDARD};
use litchi_core::{Error, Result, xml::escape_xml};
use soapberry_zip::office::StreamingArchiveWriter;
use std::collections::HashSet;
use std::io::{self, Write};
use zeroize::Zeroizing;

use super::encryption::{Profile, encrypt_entry};
use super::manifest::{
    ManifestChecksumAlgorithm, ManifestEncryption, ManifestEncryptionAlgorithm,
    ManifestKeyDerivation, ManifestStartKeyGeneration,
};
use super::package::OwnedPackage;
use super::xml_splice::XmlSplicePublication;

/// Builder for creating ODF packages (ZIP archives)
///
/// This struct helps create valid ODF files by managing the ZIP archive structure,
/// manifest, and required files.
///
/// # Examples
///
/// ```ignore
/// # use litchi_odf::core::PackageWriter;
/// # use litchi_core::Result;
/// # fn example() -> Result<()> {
/// let mut writer = PackageWriter::new();
/// writer.set_mimetype("application/vnd.oasis.opendocument.text")?;
/// writer.add_file("content.xml", b"<office:document-content>...</office:document-content>")?;
/// writer.add_file("styles.xml", b"<office:document-styles>...</office:document-styles>")?;
/// writer.add_file("meta.xml", b"<office:document-meta>...</office:document-meta>")?;
///
/// let bytes = writer.finish()?;
/// std::fs::write("document.odt", bytes)?;
/// # Ok(())
/// # }
/// ```
#[allow(
    clippy::module_name_repetitions,
    reason = "`PackageWriter` is the established public ODF package writer name."
)]
pub struct PackageWriter<W: Write = io::Cursor<Vec<u8>>> {
    zip_writer: StreamingArchiveWriter<W>,
    mimetype: Option<String>,
    manifest_entries: Vec<ManifestEntry>,
    wrote_any_entry: bool,
    wrote_mimetype: bool,
    wrote_payload_entry: bool,
    encryption: Option<WriterEncryption>,
    document_signer: Option<crate::signature::DocumentSigner>,
}

struct WriterEncryption {
    profile: Profile,
    password: Zeroizing<String>,
}

#[derive(Clone, Copy)]
enum PayloadOrigin {
    AuthoredOrChanged,
    CheckedSplice,
    ExactSource,
}

/// Entry in the ODF manifest
#[derive(Debug, Clone)]
struct ManifestEntry {
    full_path: String,
    media_type: String,
    size: Option<u64>,
    encryption: Option<ManifestEncryption>,
}

/// Helper to create standard ODF directory structure.
pub struct Structure;

/// A bounded, fallibly growing in-memory package sink.
///
/// Every write checks the configured byte limit before reserving exactly the
/// required capacity. This is intended for package publication paths that must
/// never first materialize an oversized archive.
pub struct BoundedBytes {
    bytes: Vec<u8>,
    limit: usize,
}

impl BoundedBytes {
    fn new(limit: usize) -> Self {
        Self {
            bytes: Vec::new(),
            limit,
        }
    }

    fn into_inner(self) -> Vec<u8> {
        self.bytes
    }
}

impl Write for BoundedBytes {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        let length = self
            .bytes
            .len()
            .checked_add(bytes.len())
            .ok_or_else(|| io::Error::other("ODF bounded package output size overflow"))?;
        if length > self.limit {
            return Err(io::Error::other(
                "ODF bounded package output exceeds its limit",
            ));
        }
        self.bytes.try_reserve_exact(bytes.len()).map_err(|error| {
            io::Error::other(format!(
                "ODF bounded package output allocation failed: {error}"
            ))
        })?;
        self.bytes.extend_from_slice(bytes);
        Ok(bytes.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

impl PackageWriter<io::Cursor<Vec<u8>>> {
    /// Create a new package writer that writes to memory
    #[must_use]
    pub fn new() -> Self {
        Self {
            zip_writer: StreamingArchiveWriter::new(),
            mimetype: None,
            manifest_entries: Vec::new(),
            wrote_any_entry: false,
            wrote_mimetype: false,
            wrote_payload_entry: false,
            encryption: None,
            document_signer: None,
        }
    }

    /// Create a writer whose archive bytes are bounded before materialization.
    #[must_use]
    pub fn new_bounded(limit: usize) -> PackageWriter<BoundedBytes> {
        PackageWriter::with_writer(BoundedBytes::new(limit))
    }
}

impl Default for PackageWriter<io::Cursor<Vec<u8>>> {
    fn default() -> Self {
        Self::new()
    }
}

impl<W: Write> PackageWriter<W> {
    /// Create a package writer over an arbitrary output sink.
    pub fn with_writer(writer: W) -> Self {
        Self {
            zip_writer: StreamingArchiveWriter::with_writer(writer),
            mimetype: None,
            manifest_entries: Vec::new(),
            wrote_any_entry: false,
            wrote_mimetype: false,
            wrote_payload_entry: false,
            encryption: None,
            document_signer: None,
        }
    }

    /// Configure a document signature generated after every other package entry is final.
    ///
    /// # Errors
    ///
    /// Returns an error when payload entries have already been written.
    pub fn set_document_signer(&mut self, signer: crate::signature::DocumentSigner) -> Result<()> {
        if self.wrote_payload_entry {
            return Err(Error::InvalidFormat(
                "ODF signing must be configured before payload entries".to_string(),
            ));
        }
        self.document_signer = Some(signer);
        Ok(())
    }

    /// Clear document signing before any payload entry is written.
    ///
    /// # Errors
    ///
    /// Returns an error when payload entries have already been written.
    pub fn clear_document_signer(&mut self) -> Result<()> {
        if self.wrote_payload_entry {
            return Err(Error::InvalidFormat(
                "ODF signing cannot be changed after payload entries".to_string(),
            ));
        }
        self.document_signer = None;
        Ok(())
    }

    /// Configure encryption for subsequently written payload entries.
    ///
    /// This may be called after `mimetype`, but not after any payload entry was emitted.
    ///
    /// # Errors
    ///
    /// Returns an error when payload entries have already been written.
    pub fn set_encryption(&mut self, password: impl Into<String>, profile: Profile) -> Result<()> {
        if self.wrote_payload_entry {
            return Err(Error::InvalidFormat(
                "ODF encryption must be configured before payload entries".to_string(),
            ));
        }
        // Profiles can only be constructed after validation; evaluate the password before
        // mutating state so a late call remains atomic.
        let secret = Zeroizing::new(password.into());
        self.encryption = Some(WriterEncryption {
            profile,
            password: secret,
        });
        Ok(())
    }

    /// Clear encryption before any payload entry is written.
    ///
    /// # Errors
    ///
    /// Returns an error when payload entries have already been written.
    pub fn clear_encryption(&mut self) -> Result<()> {
        if self.wrote_payload_entry {
            return Err(Error::InvalidFormat(
                "ODF encryption cannot be changed after payload entries".to_string(),
            ));
        }
        self.encryption = None;
        Ok(())
    }

    /// Set the MIME type for the document
    ///
    /// This sets both the mimetype file and the root manifest entry.
    ///
    /// # Arguments
    ///
    /// * `mimetype` - MIME type string (e.g., "application/vnd.oasis.opendocument.text")
    ///
    /// # Errors
    ///
    /// Returns an error when the MIME type has already been written, another
    /// package entry was written first, or the archive write fails.
    pub fn set_mimetype(&mut self, mimetype: &str) -> Result<()> {
        if self.wrote_mimetype {
            return Err(Error::InvalidFormat("MIME type already set".to_string()));
        }
        if self.wrote_any_entry {
            return Err(Error::InvalidFormat(
                "Cannot set MIME type after writing other files".to_string(),
            ));
        }

        self.mimetype = Some(mimetype.to_string());

        self.zip_writer
            .write_stored("mimetype", mimetype.as_bytes())
            .map_err(|e| Error::ZipError(e.to_string()))?;
        self.wrote_any_entry = true;
        self.wrote_mimetype = true;

        self.manifest_entries.push(ManifestEntry {
            full_path: "/".to_string(),
            media_type: mimetype.to_string(),
            size: None,
            encryption: None,
        });

        Ok(())
    }

    /// Add a file to the package
    ///
    /// # Arguments
    ///
    /// * `path` - Path within the ZIP archive (e.g., "content.xml", "Pictures/image1.png")
    /// * `content` - File content as bytes
    ///
    /// # Note
    ///
    /// This method automatically adds the file to the manifest with an appropriate media type.
    ///
    /// # Errors
    ///
    /// Returns an error when no MIME type is configured, the path is reserved,
    /// encryption fails, or the archive write fails.
    pub fn add_file(&mut self, path: &str, content: &[u8]) -> Result<()> {
        if path == "mimetype" {
            return Err(Error::InvalidFormat(
                "mimetype is written via set_mimetype()".to_string(),
            ));
        }
        if !self.wrote_mimetype {
            return Err(Error::InvalidFormat("MIME type not set".to_string()));
        }

        // Determine media type based on file extension
        let media_type = Self::guess_media_type(path);

        self.write_file(path, content, media_type, PayloadOrigin::AuthoredOrChanged)
    }

    /// Add a file to the package with a specific media type
    ///
    /// # Arguments
    ///
    /// * `path` - Path within the ZIP archive
    /// * `content` - File content as bytes
    /// * `media_type` - MIME type for the manifest entry
    ///
    /// # Errors
    ///
    /// Returns an error when no MIME type is configured, the path is reserved,
    /// encryption fails, or the archive write fails.
    pub fn add_file_with_media_type(
        &mut self,
        path: &str,
        content: &[u8],
        media_type: &str,
    ) -> Result<()> {
        if path == "mimetype" {
            return Err(Error::InvalidFormat(
                "mimetype is written via set_mimetype()".to_string(),
            ));
        }
        if !self.wrote_mimetype {
            return Err(Error::InvalidFormat("MIME type not set".to_string()));
        }

        self.write_file(path, content, media_type, PayloadOrigin::AuthoredOrChanged)
    }

    fn write_file(
        &mut self,
        path: &str,
        content: &[u8],
        media_type: &str,
        origin: PayloadOrigin,
    ) -> Result<()> {
        if matches!(origin, PayloadOrigin::AuthoredOrChanged) {
            Self::validate_authored_xml(path, content, media_type)?;
        }
        let encrypt = self
            .encryption
            .as_ref()
            .filter(|_| !path.starts_with("META-INF/"));
        let (size, encryption) = if let Some(settings) = encrypt {
            let (ciphertext, descriptor) =
                encrypt_entry(content, settings.password.as_str(), settings.profile)?;
            self.zip_writer
                .write_stored(path, &ciphertext)
                .map_err(|e| Error::ZipError(e.to_string()))?;
            let plaintext_size = u64::try_from(content.len()).map_err(|error| {
                Error::InvalidFormat(format!("ODF plaintext entry is too large: {error}"))
            })?;
            (Some(plaintext_size), Some(descriptor))
        } else {
            self.zip_writer
                .write_deflated(path, content)
                .map_err(|e| Error::ZipError(e.to_string()))?;
            (None, None)
        };
        self.manifest_entries.push(ManifestEntry {
            full_path: path.to_string(),
            media_type: media_type.to_string(),
            size,
            encryption,
        });
        self.wrote_any_entry = true;
        if !path.starts_with("META-INF/") {
            self.wrote_payload_entry = true;
        }
        Ok(())
    }

    fn write_exact_source_file(
        &mut self,
        path: &str,
        content: &[u8],
        media_type: &str,
    ) -> Result<()> {
        self.write_file(path, content, media_type, PayloadOrigin::ExactSource)
    }

    /// Add a provenance-bearing, individually audited XML splice publication.
    ///
    /// # Errors
    ///
    /// Returns an error when no MIME type is configured, the checked part
    /// cannot be assembled, or the archive write fails.
    pub fn add_spliced_xml(&mut self, publication: XmlSplicePublication) -> Result<()> {
        if !self.wrote_mimetype {
            return Err(Error::InvalidFormat("MIME type not set".to_string()));
        }
        let (path, content, media_type) = publication.assemble()?;
        self.write_file(&path, &content, &media_type, PayloadOrigin::CheckedSplice)
    }

    /// Copy every exact source member except regenerated metadata, signatures,
    /// and the explicitly excluded replacement paths.
    ///
    /// Unlike [`Self::copy_auxiliary_files_from`], this preserves source-loaded
    /// core parts such as `styles.xml` and `meta.xml` during a splice rebuild.
    ///
    /// # Errors
    ///
    /// Returns an error for unsupported source encryption or archive failures.
    pub(crate) fn copy_source_files_from_except(
        &mut self,
        source: &OwnedPackage,
        excluded_paths: &HashSet<String>,
    ) -> Result<()> {
        let package = source.package()?;
        if package.manifest().has_encrypted_entries() && self.encryption.is_none() {
            return Err(Error::InvalidFormat(
                "Rewriting encrypted ODF entries requires writer encryption".to_string(),
            ));
        }
        for (path, entry) in &package.manifest().entries {
            if path.ends_with('/')
                && !matches!(path.as_str(), "/" | "META-INF/")
                && !excluded_paths.contains(path)
            {
                self.add_manifest_entry(path, &entry.media_type)?;
            }
        }
        for path in package.files()? {
            if path.ends_with('/')
                || matches!(path.as_str(), "mimetype" | "META-INF/manifest.xml")
                || (path.starts_with("META-INF/") && path.ends_with("signatures.xml"))
                || excluded_paths.contains(&path)
            {
                continue;
            }
            let bytes = package.get_file(&path)?;
            let media_type = package
                .manifest()
                .get_media_type(&path)
                .unwrap_or_else(|| Self::guess_media_type(&path));
            self.write_exact_source_file(&path, &bytes, media_type)?;
        }
        Ok(())
    }

    fn validate_authored_xml(path: &str, content: &[u8], media_type: &str) -> Result<()> {
        if !xml_minifier::audit::package::is_xml_part(path, media_type) {
            return Ok(());
        }
        xml_minifier::audit::verify_authored(content, xml_minifier::audit::Limits::default())
            .map(|_report| ())
            .map_err(|source| {
                Error::InvalidFormat(format!("XML publication rejected for '{path}': {source}"))
            })
    }

    /// Add an entry to the package manifest without writing a ZIP member.
    ///
    /// ODF uses manifest-only entries for directories such as embedded objects.
    ///
    /// # Errors
    ///
    /// Returns an error when no MIME type is configured or the path is
    /// reserved or empty.
    pub fn add_manifest_entry(&mut self, path: &str, media_type: &str) -> Result<()> {
        if !self.wrote_mimetype {
            return Err(Error::InvalidFormat("MIME type not set".to_string()));
        }
        if path.is_empty() || path == "mimetype" {
            return Err(Error::InvalidFormat(
                "Invalid manifest-only path".to_string(),
            ));
        }

        self.manifest_entries.push(ManifestEntry {
            full_path: path.to_string(),
            media_type: media_type.to_string(),
            size: None,
            encryption: None,
        });
        Ok(())
    }

    /// Copy all non-core parts from an existing ODF package.
    ///
    /// Core XML parts are regenerated by mutable format writers. Digital
    /// signatures are deliberately omitted because changing those parts
    /// invalidates the signatures. Encrypted parts cannot be reconstructed
    /// faithfully with the current manifest writer and are rejected.
    ///
    /// # Errors
    ///
    /// Returns an error when the source package cannot be read, contains
    /// unsupported encrypted entries, or an entry cannot be copied.
    pub fn copy_auxiliary_files_from(&mut self, source: &OwnedPackage) -> Result<()> {
        let package = source.package()?;
        if package.manifest().has_encrypted_entries() && self.encryption.is_none() {
            return Err(Error::InvalidFormat(
                "Rewriting encrypted ODF entries requires writer encryption".to_string(),
            ));
        }

        for (path, entry) in &package.manifest().entries {
            if path.ends_with('/') && !Self::is_regenerated_package_part(path) {
                self.add_manifest_entry(path, &entry.media_type)?;
            }
        }

        for path in package.files()? {
            if path.ends_with('/') || Self::is_regenerated_package_part(&path) {
                continue;
            }
            let bytes = package.get_file(&path)?;
            let media_type = package
                .manifest()
                .get_media_type(&path)
                .unwrap_or_else(|| Self::guess_media_type(&path));
            self.write_exact_source_file(&path, &bytes, media_type)?;
        }

        Ok(())
    }

    /// Add a directory entry to the generated manifest without a ZIP payload.
    ///
    /// # Errors
    ///
    /// Returns an error when the path is not a safe relative directory or the
    /// manifest entry cannot be added.
    pub fn add_manifest_directory(&mut self, path: &str, media_type: &str) -> Result<()> {
        if !path.ends_with('/') || path.starts_with('/') || path.contains("..") {
            return Err(Error::InvalidFormat(
                "invalid embedded-object manifest directory".to_string(),
            ));
        }
        self.add_manifest_entry(path, media_type)
    }

    /// Copy auxiliary entries except selected exact paths and directory trees.
    ///
    /// # Errors
    ///
    /// Returns an error when the source package cannot be read, contains
    /// unsupported encrypted entries, or an entry cannot be copied.
    pub fn copy_auxiliary_files_from_except(
        &mut self,
        source: &OwnedPackage,
        excluded_paths: &[String],
        excluded_prefixes: &[String],
    ) -> Result<()> {
        let package = source.package()?;
        if package.manifest().has_encrypted_entries() && self.encryption.is_none() {
            return Err(Error::InvalidFormat(
                "Rewriting encrypted ODF entries requires writer encryption".to_string(),
            ));
        }
        let excluded = |path: &str| {
            excluded_paths.iter().any(|candidate| candidate == path)
                || excluded_prefixes
                    .iter()
                    .any(|candidate| path.starts_with(candidate))
        };

        for (path, entry) in &package.manifest().entries {
            if path.ends_with('/') && !Self::is_regenerated_package_part(path) && !excluded(path) {
                self.add_manifest_entry(path, &entry.media_type)?;
            }
        }
        for path in package.files()? {
            if path.ends_with('/') || Self::is_regenerated_package_part(&path) || excluded(&path) {
                continue;
            }
            let bytes = package.get_file(&path)?;
            let media_type = package
                .manifest()
                .get_media_type(&path)
                .unwrap_or_else(|| Self::guess_media_type(&path));
            self.write_exact_source_file(&path, &bytes, media_type)?;
        }
        Ok(())
    }

    /// Generate the manifest.xml content
    fn generate_manifest(&self) -> String {
        let mut manifest = String::from(
            r#"<?xml version="1.0" encoding="UTF-8"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3">"#,
        );

        // Add manifest entries
        let mut seen_paths: HashSet<&str> = HashSet::with_capacity(self.manifest_entries.len());
        for entry in &self.manifest_entries {
            if !seen_paths.insert(entry.full_path.as_str()) {
                continue;
            }
            Self::write_manifest_entry(&mut manifest, entry);
        }

        manifest.push_str("</manifest:manifest>");
        manifest
    }

    /// Guess media type from file path
    fn guess_media_type(path: &str) -> &'static str {
        if path.ends_with('/') {
            return "";
        }
        let extension = path.rsplit('.').next().unwrap_or_default();
        if extension.eq_ignore_ascii_case("xml") {
            "text/xml"
        } else if extension.eq_ignore_ascii_case("rdf") {
            "application/rdf+xml"
        } else if extension.eq_ignore_ascii_case("png") {
            "image/png"
        } else if extension.eq_ignore_ascii_case("jpg") || extension.eq_ignore_ascii_case("jpeg") {
            "image/jpeg"
        } else if extension.eq_ignore_ascii_case("gif") {
            "image/gif"
        } else if extension.eq_ignore_ascii_case("svg") {
            "image/svg+xml"
        } else {
            "application/octet-stream"
        }
    }

    /// Finish writing the package and return the bytes.
    ///
    /// This method writes the mimetype file, manifest, and finalizes the ZIP archive.
    ///
    /// # Errors
    ///
    /// Returns an error if:
    /// - No MIME type has been set
    /// - Writing to the ZIP archive fails
    fn finish_into_writer(mut self) -> Result<(W, Option<crate::signature::DocumentSigner>)> {
        if !self.wrote_mimetype {
            return Err(Error::InvalidFormat("MIME type not set".to_string()));
        }

        // Add META-INF directory to manifest
        self.manifest_entries.push(ManifestEntry {
            full_path: "META-INF/".to_string(),
            media_type: String::new(),
            size: None,
            encryption: None,
        });

        self.manifest_entries.push(ManifestEntry {
            full_path: "META-INF/manifest.xml".to_string(),
            media_type: "text/xml".to_string(),
            size: None,
            encryption: None,
        });

        // Generate and write manifest
        let manifest_content = self.generate_manifest();
        Self::validate_authored_xml(
            "META-INF/manifest.xml",
            manifest_content.as_bytes(),
            "text/xml",
        )?;
        self.zip_writer
            .write_deflated("META-INF/manifest.xml", manifest_content.as_bytes())
            .map_err(|e| Error::ZipError(e.to_string()))?;

        // Finish ZIP archive and return bytes
        let signer = self.document_signer.take();
        let writer = self
            .zip_writer
            .finish()
            .map_err(|e| Error::ZipError(e.to_string()))?;
        Ok((writer, signer))
    }
}

impl PackageWriter<io::Cursor<Vec<u8>>> {
    /// Finish writing the package and return the bytes.
    ///
    /// This method writes the mimetype file, manifest, and finalizes the ZIP archive.
    ///
    /// # Errors
    ///
    /// Returns an error if:
    /// - No MIME type has been set
    /// - Writing to the ZIP archive fails
    pub fn finish(self) -> Result<Vec<u8>> {
        let (cursor, document_signer) = self.finish_into_writer()?;
        let bytes = cursor.into_inner();
        if let Some(signer) = &document_signer {
            crate::signature::sign_package(&bytes, signer)
        } else {
            Ok(bytes)
        }
    }

    /// Alias for `finish()` for API compatibility.
    ///
    /// # Errors
    ///
    /// Returns an error under the same conditions as [`Self::finish`].
    pub fn finish_to_bytes(self) -> Result<Vec<u8>> {
        self.finish()
    }
}

impl PackageWriter<BoundedBytes> {
    /// Finish a package into the configured bounded sink.
    ///
    /// Document signing is refused because signing can create a second package
    /// representation outside this sink.
    ///
    /// # Errors
    ///
    /// Returns an error when signing is configured, a MIME type is missing, or
    /// the archive cannot be finalized into the bounded sink.
    pub fn finish_to_bounded_bytes(self) -> Result<Vec<u8>> {
        if self.document_signer.is_some() {
            return Err(Error::InvalidFormat(
                "ODF bounded package output does not support document signing".to_string(),
            ));
        }
        let (sink, signer) = self.finish_into_writer()?;
        debug_assert!(signer.is_none());
        Ok(sink.into_inner())
    }
}

impl<W: Write> PackageWriter<W> {
    fn write_manifest_entry(xml: &mut String, entry: &ManifestEntry) {
        xml.push_str("<manifest:file-entry manifest:full-path=\"");
        xml.push_str(&escape_xml(&entry.full_path));
        xml.push_str("\" manifest:media-type=\"");
        xml.push_str(&escape_xml(&entry.media_type));
        xml.push('"');
        if let Some(size) = entry.size {
            xml.push_str(" manifest:size=\"");
            xml.push_str(&size.to_string());
            xml.push('"');
        }
        let Some(encryption) = &entry.encryption else {
            xml.push_str("/>");
            return;
        };
        xml.push('>');
        xml.push_str("<manifest:encryption-data");
        if let Some(checksum) = &encryption.checksum {
            let algorithm = match checksum.algorithm {
                ManifestChecksumAlgorithm::Sha1First1024 => "SHA1/1K",
                ManifestChecksumAlgorithm::Sha256First1024 => {
                    "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0#sha256-1k"
                },
            };
            xml.push_str(" manifest:checksum-type=\"");
            xml.push_str(algorithm);
            xml.push_str("\" manifest:checksum=\"");
            xml.push_str(&BASE64_STANDARD.encode(&checksum.value));
            xml.push('"');
        }
        xml.push('>');

        let (algorithm_name, iv): (&str, &[u8]) = match &encryption.algorithm {
            ManifestEncryptionAlgorithm::Aes128Cbc { iv } => {
                ("http://www.w3.org/2001/04/xmlenc#aes128-cbc", iv)
            },
            ManifestEncryptionAlgorithm::Aes192Cbc { iv } => {
                ("http://www.w3.org/2001/04/xmlenc#aes192-cbc", iv)
            },
            ManifestEncryptionAlgorithm::Aes256Cbc { iv } => {
                ("http://www.w3.org/2001/04/xmlenc#aes256-cbc", iv)
            },
            ManifestEncryptionAlgorithm::Aes128Gcm { iv } => {
                ("http://www.w3.org/2009/xmlenc11#aes128-gcm", iv)
            },
            ManifestEncryptionAlgorithm::Aes192Gcm { iv } => {
                ("http://www.w3.org/2009/xmlenc11#aes192-gcm", iv)
            },
            ManifestEncryptionAlgorithm::Aes256Gcm { iv } => {
                ("http://www.w3.org/2009/xmlenc11#aes256-gcm", iv)
            },
            ManifestEncryptionAlgorithm::BlowfishCfb8 { iv } => ("Blowfish CFB", iv),
        };
        xml.push_str("<manifest:algorithm manifest:algorithm-name=\"");
        xml.push_str(algorithm_name);
        xml.push_str("\" manifest:initialisation-vector=\"");
        xml.push_str(&BASE64_STANDARD.encode(iv));
        xml.push_str("\"/>");

        let (start_name, start_size) = match encryption.start_key {
            ManifestStartKeyGeneration::Sha1 => ("SHA1", 20),
            ManifestStartKeyGeneration::Sha256 => ("http://www.w3.org/2001/04/xmlenc#sha256", 32),
        };
        xml.push_str("<manifest:start-key-generation manifest:start-key-generation-name=\"");
        xml.push_str(start_name);
        xml.push_str("\" manifest:key-size=\"");
        xml.push_str(&start_size.to_string());
        xml.push_str("\"/>");

        match &encryption.key_derivation {
            ManifestKeyDerivation::Pbkdf2 {
                salt,
                iterations,
                key_size,
            } => {
                xml.push_str(
                "<manifest:key-derivation manifest:key-derivation-name=\"PBKDF2\" manifest:salt=\"",
            );
                xml.push_str(&BASE64_STANDARD.encode(salt));
                xml.push_str("\" manifest:iteration-count=\"");
                xml.push_str(&iterations.get().to_string());
                xml.push_str("\" manifest:key-size=\"");
                xml.push_str(&key_size.to_string());
                xml.push_str("\"/>");
            },
            ManifestKeyDerivation::Argon2id {
                salt,
                iterations,
                memory_kib,
                lanes,
                key_size,
            } => {
                xml.push_str("<manifest:key-derivation manifest:key-derivation-name=\"urn:oasis:names:tc:opendocument:xmlns:manifest:1.5#argon2id\" manifest:salt=\"");
                xml.push_str(&BASE64_STANDARD.encode(salt));
                xml.push_str("\" manifest:argon2-iterations=\"");
                xml.push_str(&iterations.get().to_string());
                xml.push_str("\" manifest:argon2-memory=\"");
                xml.push_str(&memory_kib.get().to_string());
                xml.push_str("\" manifest:argon2-lanes=\"");
                xml.push_str(&lanes.get().to_string());
                xml.push('"');
                if let Some(optional_key_size) = key_size {
                    xml.push_str(" manifest:key-size=\"");
                    xml.push_str(&optional_key_size.to_string());
                    xml.push('"');
                }
                xml.push_str("/>");
            },
        }
        xml.push_str("</manifest:encryption-data></manifest:file-entry>");
    }

    fn is_regenerated_package_part(path: &str) -> bool {
        matches!(
            path,
            "/" | "mimetype"
                | "content.xml"
                | "styles.xml"
                | "meta.xml"
                | "manifest.xml"
                | "META-INF/"
                | "META-INF/manifest.xml"
        ) || (path.starts_with("META-INF/") && path.ends_with("signatures.xml"))
    }
}

impl Structure {
    /// Generate a default content.xml skeleton
    #[must_use]
    pub fn default_content_xml(office_type: &str) -> String {
        format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0" xmlns:math="http://www.w3.org/1998/Math/MathML" xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0" xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0" xmlns:ooo="http://openoffice.org/2004/office" xmlns:ooow="http://openoffice.org/2004/writer" xmlns:oooc="http://openoffice.org/2004/calc" xmlns:dom="http://www.w3.org/2001/xml-events" xmlns:xforms="http://www.w3.org/2002/xforms" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:rpt="http://openoffice.org/2005/report" xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2" xmlns:xhtml="http://www.w3.org/1999/xhtml" xmlns:grddl="http://www.w3.org/2003/g/data-view#" xmlns:tableooo="http://openoffice.org/2009/table" xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0" xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0" xmlns:field="urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0" xmlns:formx="urn:openoffice:names:experimental:ooxml-odf-interop:xmlns:form:1.0" xmlns:css3t="http://www.w3.org/TR/css3-text/" office:version="1.3"><office:scripts/><office:font-face-decls/><office:automatic-styles/><office:body><{office_type}></{office_type}></office:body></office:document-content>"#
        )
    }

    /// Generate a default styles.xml skeleton
    #[must_use]
    pub fn default_styles_xml() -> String {
        r#"<?xml version="1.0" encoding="UTF-8"?><office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0" xmlns:math="http://www.w3.org/1998/Math/MathML" xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0" xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0" xmlns:ooo="http://openoffice.org/2004/office" xmlns:ooow="http://openoffice.org/2004/writer" xmlns:oooc="http://openoffice.org/2004/calc" xmlns:dom="http://www.w3.org/2001/xml-events" xmlns:rpt="http://openoffice.org/2005/report" xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2" xmlns:xhtml="http://www.w3.org/1999/xhtml" xmlns:grddl="http://www.w3.org/2003/g/data-view#" xmlns:tableooo="http://openoffice.org/2009/table" xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0" xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0" xmlns:field="urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0" xmlns:formx="urn:openoffice:names:experimental:ooxml-odf-interop:xmlns:form:1.0" xmlns:css3t="http://www.w3.org/TR/css3-text/" office:version="1.3"><office:font-face-decls/><office:styles/><office:automatic-styles/><office:master-styles/></office:document-styles>"#.to_string()
    }

    /// Generate a default meta.xml skeleton
    #[must_use]
    pub fn default_meta_xml() -> String {
        let now = chrono::Utc::now().to_rfc3339();
        format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:ooo="http://openoffice.org/2004/office" xmlns:grddl="http://www.w3.org/2003/g/data-view#" office:version="1.3"><office:meta><meta:generator>Litchi/0.0.1</meta:generator><meta:creation-date>{now}</meta:creation-date><dc:date>{now}</dc:date></office:meta></office:document-meta>"#
        )
    }

    /// Generate a default settings.xml skeleton
    #[must_use]
    pub fn default_settings_xml() -> String {
        r#"<?xml version="1.0" encoding="UTF-8"?><office:document-settings xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0" xmlns:ooo="http://openoffice.org/2004/office" office:version="1.3"><office:settings><config:config-item-set config:name="ooo:view-settings"><config:config-item config:name="ViewAreaTop" config:type="long">0</config:config-item><config:config-item config:name="ViewAreaLeft" config:type="long">0</config:config-item><config:config-item config:name="ViewAreaWidth" config:type="long">1</config:config-item><config:config-item config:name="ViewAreaHeight" config:type="long">1</config:config-item></config:config-item-set></office:settings></office:document-settings>"#.to_string()
    }
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    reason = "Test fixtures use infallible ZIP setup operations so assertions can focus on writer behavior."
)]
mod tests {
    use super::*;
    use soapberry_zip::office::ArchiveReader;
    use std::io::{Cursor, Write};

    #[test]
    fn test_package_writer_new() {
        let writer = PackageWriter::new();
        assert!(!writer.wrote_mimetype);
        assert!(!writer.wrote_any_entry);
        assert!(writer.mimetype.is_none());
    }

    #[test]
    fn test_package_writer_default() {
        let writer = PackageWriter::default();
        assert!(!writer.wrote_mimetype);
    }

    #[test]
    fn test_package_writer_set_mimetype() {
        let mut writer = PackageWriter::new();
        assert!(
            writer
                .set_mimetype("application/vnd.oasis.opendocument.text")
                .is_ok()
        );
        assert!(writer.wrote_mimetype);
        assert_eq!(
            writer.mimetype,
            Some("application/vnd.oasis.opendocument.text".to_string())
        );
    }

    #[test]
    fn test_package_writer_set_mimetype_twice() {
        let mut writer = PackageWriter::new();
        writer
            .set_mimetype("application/vnd.oasis.opendocument.text")
            .unwrap();
        assert!(
            writer
                .set_mimetype("application/vnd.oasis.opendocument.spreadsheet")
                .is_err()
        );
    }

    #[test]
    fn test_package_writer_add_file_without_mimetype() {
        let mut writer = PackageWriter::new();
        assert!(writer.add_file("content.xml", b"test").is_err());
    }

    #[test]
    fn test_package_writer_add_mimetype_as_file() {
        let mut writer = PackageWriter::new();
        writer
            .set_mimetype("application/vnd.oasis.opendocument.text")
            .unwrap();
        assert!(writer.add_file("mimetype", b"test").is_err());
    }

    #[test]
    fn test_package_writer_finish_without_mimetype() {
        let writer = PackageWriter::new();
        assert!(writer.finish().is_err());
    }

    #[test]
    fn test_package_writer_finish_to_bytes() {
        let mut writer = PackageWriter::new();
        writer
            .set_mimetype("application/vnd.oasis.opendocument.text")
            .unwrap();
        let result = writer.finish_to_bytes();
        assert!(result.is_ok());
        let bytes = result.unwrap();
        assert!(!bytes.is_empty());
    }

    #[test]
    fn test_package_writer_add_file_with_media_type() {
        let mut writer = PackageWriter::new();
        writer
            .set_mimetype("application/vnd.oasis.opendocument.text")
            .unwrap();
        assert!(
            writer
                .add_file_with_media_type("custom.dat", b"data", "application/octet-stream")
                .is_ok()
        );
    }

    #[test]
    fn arbitrary_writer_bytes_are_refused_for_every_xml_classification() {
        for (path, media_type) in [
            ("manifest.rdf", "application/octet-stream"),
            ("custom/metadata", "application/rdf+xml"),
            (
                "META-INF/custom-signature",
                "application/vnd.oasis.opendocument.digital-signature+xml",
            ),
        ] {
            let mut writer = PackageWriter::new();
            writer
                .set_mimetype("application/vnd.oasis.opendocument.text")
                .unwrap();
            let error = writer
                .add_file_with_media_type(path, b"<root> <child/></root>", media_type)
                .unwrap_err();
            assert!(error.to_string().contains("XML publication rejected"));
        }
    }

    #[test]
    fn real_package_enumeration_includes_rdf_and_manifest_declared_xml() {
        let mut writer = PackageWriter::new();
        writer
            .set_mimetype("application/vnd.oasis.opendocument.text")
            .unwrap();
        writer
            .add_file_with_media_type(
                "manifest.rdf",
                b"<rdf:RDF xmlns:rdf=\"urn:test\"><rdf:Description/></rdf:RDF>",
                "application/rdf+xml",
            )
            .unwrap();
        writer
            .add_file_with_media_type(
                "custom/metadata",
                b"<metadata><value>content</value></metadata>",
                "application/vnd.example.metadata+xml",
            )
            .unwrap();
        let bytes = writer.finish().unwrap();
        let archive = ArchiveReader::new(&bytes).unwrap();
        let manifest_xml = archive.read("META-INF/manifest.xml").unwrap();
        let manifest_text = std::str::from_utf8(&manifest_xml).unwrap();
        let manifest = super::super::manifest::Manifest::parse(manifest_text).unwrap();
        let mut audited = Vec::new();
        for name in archive.file_names() {
            let media_type = manifest.get_media_type(name).unwrap_or_default();
            if xml_minifier::audit::package::is_xml_part(name, media_type) {
                let payload = archive.read(name).unwrap();
                let _report = xml_minifier::audit::verify_authored(
                    &payload,
                    xml_minifier::audit::Limits::default(),
                )
                .unwrap();
                audited.push(name.to_string());
            }
        }
        audited.sort();
        assert_eq!(
            audited,
            ["META-INF/manifest.xml", "custom/metadata", "manifest.rdf"]
        );
    }

    #[test]
    fn auxiliary_copy_preserves_noncompact_source_rdf_exactly() {
        let mut source_bytes = Vec::new();
        let source_rdf = b"<rdf:RDF xmlns:rdf=\"urn:test\">\n <rdf:Description/>\n</rdf:RDF>";
        {
            let mut zip = zip::ZipWriter::new(Cursor::new(&mut source_bytes));
            let options = zip::write::SimpleFileOptions::default()
                .compression_method(zip::CompressionMethod::Stored);
            zip.start_file("mimetype", options).unwrap();
            zip.write_all(b"application/vnd.oasis.opendocument.text")
                .unwrap();
            zip.start_file("manifest.rdf", options).unwrap();
            zip.write_all(source_rdf).unwrap();
            zip.start_file("META-INF/manifest.xml", options).unwrap();
            zip.write_all(br#"<?xml version="1.0"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/><manifest:file-entry manifest:full-path="manifest.rdf" manifest:media-type="application/rdf+xml"/></manifest:manifest>"#).unwrap();
            zip.finish().unwrap();
        }
        let source = OwnedPackage::from_bytes(source_bytes).unwrap();
        let mut writer = PackageWriter::new();
        writer
            .set_mimetype("application/vnd.oasis.opendocument.text")
            .unwrap();
        writer.copy_auxiliary_files_from(&source).unwrap();
        let output = writer.finish().unwrap();
        let archive = ArchiveReader::new(&output).unwrap();
        assert_eq!(archive.read("manifest.rdf").unwrap(), source_rdf);
    }

    #[test]
    fn copying_auxiliary_files_rejects_encrypted_manifest_entries() {
        let mut bytes = Vec::new();
        {
            let mut zip = zip::ZipWriter::new(Cursor::new(&mut bytes));
            let options = zip::write::SimpleFileOptions::default()
                .compression_method(zip::CompressionMethod::Stored);
            zip.start_file("mimetype", options).unwrap();
            zip.write_all(b"application/vnd.oasis.opendocument.text")
                .unwrap();
            zip.start_file("content.xml", options).unwrap();
            zip.write_all(b"encrypted payload").unwrap();
            zip.start_file("META-INF/manifest.xml", options).unwrap();
            zip.write_all(br#"<?xml version="1.0"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"><manifest:encryption-data manifest:checksum-type="SHA256" manifest:checksum="checksum"/></manifest:file-entry></manifest:manifest>"#).unwrap();
            zip.finish().unwrap();
        }

        let source = OwnedPackage::from_bytes(bytes).unwrap();
        let mut writer = PackageWriter::new();
        writer
            .set_mimetype("application/vnd.oasis.opendocument.text")
            .unwrap();
        let error = writer.copy_auxiliary_files_from(&source).unwrap_err();
        assert!(error.to_string().contains("encrypted entries"));
    }

    #[test]
    fn test_guess_media_type() {
        type MemoryWriter = PackageWriter<Cursor<Vec<u8>>>;
        assert_eq!(MemoryWriter::guess_media_type("content.xml"), "text/xml");
        assert_eq!(
            MemoryWriter::guess_media_type("manifest.rdf"),
            "application/rdf+xml"
        );
        assert_eq!(MemoryWriter::guess_media_type("image.png"), "image/png");
        assert_eq!(MemoryWriter::guess_media_type("image.jpg"), "image/jpeg");
        assert_eq!(MemoryWriter::guess_media_type("image.jpeg"), "image/jpeg");
        assert_eq!(MemoryWriter::guess_media_type("image.gif"), "image/gif");
        assert_eq!(MemoryWriter::guess_media_type("image.svg"), "image/svg+xml");
        assert_eq!(MemoryWriter::guess_media_type("META-INF/"), "");
        assert_eq!(
            MemoryWriter::guess_media_type("data.bin"),
            "application/octet-stream"
        );
    }

    #[test]
    fn test_generate_manifest() {
        let mut writer = PackageWriter::new();
        writer
            .set_mimetype("application/vnd.oasis.opendocument.text")
            .unwrap();
        writer.add_file("content.xml", b"<content/>").unwrap();

        let manifest = writer.generate_manifest();
        assert!(manifest.contains("manifest:manifest"));
        assert!(manifest.contains("content.xml"));
        assert!(manifest.contains("text/xml"));
    }

    #[test]
    fn test_odf_structure_default_styles_xml() {
        let styles = Structure::default_styles_xml();
        assert!(styles.contains("office:document-styles"));
        assert!(styles.contains("office:styles"));
    }

    #[test]
    fn test_odf_structure_default_meta_xml() {
        let meta = Structure::default_meta_xml();
        assert!(meta.contains("office:document-meta"));
        assert!(meta.contains("Litchi"));
        assert!(meta.contains("meta:creation-date"));
    }

    #[test]
    fn test_odf_structure_default_settings_xml() {
        let settings = Structure::default_settings_xml();
        assert!(settings.contains("office:document-settings"));
        assert!(settings.contains("config:config-item"));
    }

    #[test]
    fn test_odf_structure_default_content_xml() {
        let content = Structure::default_content_xml("office:text");
        assert!(content.contains("office:document-content"));
        assert!(content.contains("office:text"));
        assert!(content.contains("office:body"));
    }

    #[test]
    fn test_manifest_entry_debug() {
        let entry = ManifestEntry {
            full_path: "content.xml".to_string(),
            media_type: "text/xml".to_string(),
            size: None,
            encryption: None,
        };
        let debug_str = format!("{entry:?}");
        assert!(debug_str.contains("content.xml"));
        assert!(debug_str.contains("text/xml"));
    }

    #[test]
    fn test_package_writer_full_package() {
        let mut writer = PackageWriter::new();

        // Set mimetype
        writer
            .set_mimetype("application/vnd.oasis.opendocument.text")
            .unwrap();

        // Add files
        writer
            .add_file("content.xml", b"<office:document-content/>")
            .unwrap();
        writer
            .add_file("styles.xml", b"<office:document-styles/>")
            .unwrap();
        writer
            .add_file("meta.xml", b"<office:document-meta/>")
            .unwrap();

        // Finish
        let bytes = writer.finish().unwrap();
        assert!(!bytes.is_empty());

        // Verify it's a valid ZIP (starts with PK)
        assert_eq!(&bytes[0..2], b"PK");
    }
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    reason = "Test fixtures use infallible encrypted-package setup so assertions can focus on rewrite behavior."
)]
mod encrypted_copy_tests {
    use super::*;

    #[test]
    fn encrypted_auxiliary_entries_require_plaintext_and_are_reencrypted() {
        let mut source_writer = PackageWriter::new();
        source_writer
            .set_mimetype("application/vnd.oasis.opendocument.text")
            .unwrap();
        source_writer
            .set_encryption("source-password", Profile::compatible())
            .unwrap();
        source_writer
            .add_file("Pictures/asset.bin", b"encrypted auxiliary bytes")
            .unwrap();
        source_writer
            .add_file("META-INF/documentsignatures.xml", b"<signatures/>")
            .unwrap();
        let source = OwnedPackage::from_bytes_with_password(
            source_writer.finish().unwrap(),
            "source-password",
        )
        .unwrap();

        let mut unencrypted = PackageWriter::new();
        unencrypted
            .set_mimetype("application/vnd.oasis.opendocument.text")
            .unwrap();
        assert!(unencrypted.copy_auxiliary_files_from(&source).is_err());

        let mut destination = PackageWriter::new();
        destination
            .set_mimetype("application/vnd.oasis.opendocument.text")
            .unwrap();
        destination
            .set_encryption("new-password", Profile::compatible())
            .unwrap();
        destination.copy_auxiliary_files_from(&source).unwrap();
        let rewritten =
            OwnedPackage::from_bytes_with_password(destination.finish().unwrap(), "new-password")
                .unwrap();
        assert_eq!(
            rewritten.get_file("Pictures/asset.bin").unwrap(),
            b"encrypted auxiliary bytes"
        );
        assert!(
            !rewritten
                .has_file("META-INF/documentsignatures.xml")
                .unwrap()
        );
    }
}
