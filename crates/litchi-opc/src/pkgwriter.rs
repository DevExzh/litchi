//! Package writer for OPC packages.
//!
//! This module provides functionality to serialize and write OPC packages to disk,
//! including writing `[Content_Types].xml`, relationships, and all parts.
use crate::constants::content_type as ct;
use crate::content_type::ContentType;
use crate::error::Result;
use crate::package::OpcPackage;
use crate::packuri::{CONTENT_TYPES_URI, PACKAGE_URI, PackURI};
use crate::phys_pkg::PhysPkgWriter;
use litchi_core::xml::escape_xml;
use std::collections::HashMap;
use std::fmt::Write as _;
use std::io::Write;
use std::path::Path;

struct Counted<'a, W> {
    inner: W,
    written: &'a mut u64,
}

impl<W: Write> Write for Counted<'_, W> {
    fn write(&mut self, bytes: &[u8]) -> std::io::Result<usize> {
        let written = self.inner.write(bytes)?;
        *self.written = self.written.saturating_add(written as u64);
        Ok(written)
    }

    fn flush(&mut self) -> std::io::Result<()> {
        self.inner.flush()
    }
}

/// Package writer that serializes an OPC package to a ZIP file.
///
/// This is the main entry point for saving packages. It handles writing:
/// - `[Content_Types].xml`
/// - _rels/.rels (package relationships)
/// - All parts and their relationships
///
/// # Example
///
/// ```no_run
/// use litchi_opc::package::OpcPackage;
/// use litchi_opc::pkgwriter::PackageWriter;
///
/// let mut pkg = OpcPackage::new();
/// // ... add parts to package ...
/// PackageWriter::write("output.docx", &pkg)?;
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
pub struct PackageWriter;

impl PackageWriter {
    /// Atomically write an OPC package to a file.
    ///
    /// # Arguments
    /// * `path` - Path where the package should be written
    /// * `package` - The OPC package to write
    pub fn write<P: AsRef<Path>>(path: P, package: &OpcPackage) -> Result<()> {
        crate::atomic::replace(path.as_ref(), |writer| {
            Self::write_to_stream(writer, package)
        })
    }

    /// Write an OPC package directly to a sequential stream.
    ///
    /// On failure after output begins, [`crate::OpcError::IncompleteOutput`]
    /// reports how many bytes the sink accepted. Seeking is not required.
    ///
    /// # Arguments
    /// * `writer` - A writer that implements Write
    /// * `package` - The OPC package to write
    pub fn write_to_stream<W: Write>(writer: W, package: &OpcPackage) -> Result<()> {
        let mut written = 0_u64;
        let result = Self::write_counted(
            Counted {
                inner: writer,
                written: &mut written,
            },
            package,
        );
        match result {
            Err(source) if written != 0 => Err(crate::OpcError::IncompleteOutput {
                written,
                source: Box::new(source),
            }),
            result => result,
        }
    }

    fn write_counted<W: Write>(writer: W, package: &OpcPackage) -> Result<()> {
        let mut physical = PhysPkgWriter::with_writer(writer);
        Self::write_package(&mut physical, package)?;
        let mut writer = physical.finish_into_inner()?;
        writer.flush()?;
        Ok(())
    }

    /// Serialize an OPC package to bytes.
    ///
    /// # Arguments
    /// * `package` - The OPC package to serialize
    ///
    /// # Returns
    /// The serialized package as a byte vector
    pub fn to_bytes(package: &OpcPackage) -> Result<Vec<u8>> {
        let mut physical = PhysPkgWriter::new();
        Self::write_package(&mut physical, package)?;
        physical.finish()
    }

    fn write_package<W: Write>(
        physical: &mut PhysPkgWriter<W>,
        package: &OpcPackage,
    ) -> Result<()> {
        Self::write_content_types(physical, package)?;
        Self::write_pkg_rels(physical, package)?;
        Self::write_parts(physical, package)
    }

    /// Write the `[Content_Types].xml` part.
    ///
    /// This file maps file extensions and part names to content types.
    fn write_content_types<W: Write>(
        phys_writer: &mut PhysPkgWriter<W>,
        package: &OpcPackage,
    ) -> Result<()> {
        let cti = ContentTypesItem::from_package(package)?;
        let blob = cti.to_xml();

        let content_types_uri =
            PackURI::new(CONTENT_TYPES_URI).map_err(crate::error::OpcError::InvalidPackUri)?;
        phys_writer.write(&content_types_uri, blob.as_bytes())?;

        Ok(())
    }

    /// Write package-level relationships.
    fn write_pkg_rels<W: Write>(
        phys_writer: &mut PhysPkgWriter<W>,
        package: &OpcPackage,
    ) -> Result<()> {
        let package_uri =
            PackURI::new(PACKAGE_URI).map_err(crate::error::OpcError::InvalidPackUri)?;
        let rels_uri = package_uri
            .rels_uri()
            .map_err(crate::error::OpcError::InvalidPackUri)?;
        let rels_xml = package.rels().to_xml();
        phys_writer.write(&rels_uri, rels_xml.as_bytes())?;

        Ok(())
    }

    /// Write all parts and their relationships.
    fn write_parts<W: Write>(
        phys_writer: &mut PhysPkgWriter<W>,
        package: &OpcPackage,
    ) -> Result<()> {
        // `OpcPackage` uses a hash map for O(1) lookup. Never let its randomized
        // iteration order leak into serialized artifacts.
        let mut parts = package.iter_parts().collect::<Vec<_>>();
        parts.sort_unstable_by(|left, right| {
            left.partname().as_str().cmp(right.partname().as_str())
        });
        for part in parts {
            // Write the part itself
            let blob = part.blob();
            phys_writer.write(part.partname(), blob)?;

            // Write the part's relationships if it has any
            if !part.rels().is_empty() {
                let rels_uri = part
                    .partname()
                    .rels_uri()
                    .map_err(crate::error::OpcError::InvalidPackUri)?;
                let rels_xml = part.rels().to_xml();
                phys_writer.write(&rels_uri, rels_xml.as_bytes())?;
            }
        }

        Ok(())
    }
}

/// Helper for building `[Content_Types].xml` content.
///
/// Manages Default and Override elements for content type mapping.
struct ContentTypesItem {
    /// Default content types by extension
    defaults: HashMap<String, ContentType>,

    /// Override content types by partname
    overrides: HashMap<String, ContentType>,
}

impl ContentTypesItem {
    /// Create a new `ContentTypesItem`.
    fn new() -> Result<Self> {
        let mut defaults = HashMap::new();

        // Add standard defaults
        defaults.insert("rels".to_string(), ContentType::new(ct::OPC_RELATIONSHIPS)?);
        defaults.insert("xml".to_string(), ContentType::new(ct::XML)?);

        Ok(Self {
            defaults,
            overrides: HashMap::new(),
        })
    }

    /// Build `ContentTypesItem` from an OPC package.
    fn from_package(package: &OpcPackage) -> Result<Self> {
        let mut cti = Self::new()?;

        for part in package.iter_parts() {
            cti.add_content_type(part.partname(), part.content_type())?;
        }

        Ok(cti)
    }

    /// Add a content type for a part.
    ///
    /// Uses a default mapping if the extension matches a well-known type,
    /// otherwise uses an override for the specific partname.
    fn add_content_type(&mut self, partname: &PackURI, content_type: &str) -> Result<()> {
        let ext = partname.ext().to_ascii_lowercase();
        let content_type = ContentType::new(content_type)?;

        // Check if this is a standard default mapping
        if Self::is_default_content_type(&ext, content_type.as_str()) {
            self.defaults.insert(ext, content_type);
        } else {
            self.overrides.insert(partname.to_string(), content_type);
        }
        Ok(())
    }

    /// Check if an extension/content-type pair is a standard default.
    fn is_default_content_type(ext: &str, content_type: &str) -> bool {
        matches!(
            (ext, content_type),
            ("rels", ct::OPC_RELATIONSHIPS)
                | ("xml", ct::XML)
                | ("bin", ct::XLSB_BIN)
                | ("png", "image/png")
                | ("jpg", "image/jpeg")
                | ("jpeg", "image/jpeg")
                | ("gif", "image/gif")
                | ("emf", "image/x-emf")
                | ("wmf", "image/x-wmf")
                | (
                    "odttf",
                    "application/vnd.openxmlformats-officedocument.obfuscatedFont"
                )
        )
    }

    /// Generate the XML for `[Content_Types].xml`.
    fn to_xml(&self) -> String {
        let mut xml = String::with_capacity(4096);

        xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
        xml.push_str(
            r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">"#,
        );

        // Write Default elements (sorted by extension)
        let mut exts: Vec<_> = self.defaults.keys().collect();
        exts.sort();
        for ext in exts {
            let content_type = &self.defaults[ext];
            let _ignored = write!(
                xml,
                r#"<Default Extension="{}" ContentType="{}"/>"#,
                escape_xml(ext),
                escape_xml(content_type.as_str())
            );
        }

        // Write Override elements (sorted by partname)
        let mut partnames: Vec<_> = self.overrides.keys().collect();
        partnames.sort();
        for partname in partnames {
            let content_type = &self.overrides[partname];
            let _ignored = write!(
                xml,
                r#"<Override PartName="{}" ContentType="{}"/>"#,
                escape_xml(partname),
                escape_xml(content_type.as_str())
            );
        }

        xml.push_str("</Types>");

        xml
    }
}

#[cfg(test)]
mod tests {
    #![allow(
        clippy::unwrap_used,
        clippy::expect_used,
        reason = "test assertions panic on failure by design"
    )]
    use std::io;

    use super::*;

    struct ChunkSink {
        total: usize,
        writes: usize,
        largest: usize,
        limit: usize,
    }

    struct FailAfter {
        written: usize,
        limit: usize,
    }

    impl Write for FailAfter {
        fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
            let available = self.limit.saturating_sub(self.written);
            if available == 0 {
                return Err(io::Error::new(
                    io::ErrorKind::BrokenPipe,
                    "injected sink failure",
                ));
            }
            let accepted = available.min(bytes.len());
            self.written += accepted;
            Ok(accepted)
        }

        fn flush(&mut self) -> io::Result<()> {
            Ok(())
        }
    }

    impl Write for ChunkSink {
        fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
            if bytes.len() > self.limit {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidData,
                    "writer received an archive-sized chunk",
                ));
            }
            self.total = self.total.saturating_add(bytes.len());
            self.writes = self.writes.saturating_add(1);
            self.largest = self.largest.max(bytes.len());
            Ok(bytes.len())
        }

        fn flush(&mut self) -> io::Result<()> {
            Ok(())
        }
    }

    #[test]
    fn test_content_types_xml() {
        let mut cti = ContentTypesItem::new().unwrap();
        cti.defaults
            .insert("png".to_string(), ContentType::new("image/png").unwrap());
        cti.overrides.insert(
            "/word/document.xml".to_string(),
            ContentType::new(ct::WML_DOCUMENT_MAIN).unwrap(),
        );

        let xml = cti.to_xml();

        assert!(xml.contains(r#"<Default Extension="png" ContentType="image/png"/>"#));
        assert!(xml.contains(r#"<Override PartName="/word/document.xml""#));
    }

    #[test]
    fn rejects_invalid_part_content_type_before_writing() {
        let mut package = OpcPackage::new();
        package.add_part(Box::new(crate::BlobPart::new(
            PackURI::new("/custom/data.bin").unwrap(),
            "application/octet-stream (comment)".to_string(),
            Vec::new(),
        )));
        assert!(matches!(
            PackageWriter::to_bytes(&package),
            Err(crate::OpcError::InvalidContentType { .. })
        ));
    }

    #[test]
    fn streams_large_packages_to_a_non_seekable_bounded_chunk_sink() {
        let mut state = 0x9e37_79b9_u32;
        let mut payload = Vec::with_capacity(2 * 1024 * 1024);
        while payload.len() < payload.capacity() {
            state ^= state << 13;
            state ^= state >> 17;
            state ^= state << 5;
            payload.push((state >> 24) as u8);
        }
        let mut package = OpcPackage::new();
        package.add_part(Box::new(crate::BlobPart::new(
            PackURI::new("/custom/random.bin").expect("valid part URI"),
            "application/octet-stream".to_owned(),
            payload,
        )));
        let mut sink = ChunkSink {
            total: 0,
            writes: 0,
            largest: 0,
            limit: 64 * 1024,
        };

        PackageWriter::write_to_stream(&mut sink, &package).expect("stream package");

        assert!(sink.total > 1024 * 1024);
        assert!(sink.writes > 1);
        assert!(sink.largest <= sink.limit);
    }

    #[test]
    fn incomplete_stream_errors_report_accepted_bytes() {
        let package = OpcPackage::new();
        let sink = FailAfter {
            written: 0,
            limit: 128,
        };

        let error = PackageWriter::write_to_stream(sink, &package)
            .expect_err("bounded sink must reject the package");

        assert!(matches!(
            error,
            crate::OpcError::IncompleteOutput {
                written: 128,
                source,
            } if matches!(*source, crate::OpcError::ZipError(_))
        ));
    }

    #[test]
    fn filesystem_write_replaces_only_with_a_finalized_package() {
        let directory = tempfile::tempdir().expect("temporary directory");
        let destination = directory.path().join("package.xlsx");
        std::fs::write(&destination, b"previous artifact").expect("seed destination");
        let mut package = OpcPackage::new();
        let partname = PackURI::new("/custom/data.bin").expect("valid part URI");
        package.add_part(Box::new(crate::BlobPart::new(
            partname.clone(),
            "application/octet-stream".to_owned(),
            b"payload".to_vec(),
        )));

        PackageWriter::write(&destination, &package).expect("atomic package write");

        let reopened = OpcPackage::open(destination).expect("reopen package");
        assert_eq!(
            reopened.get_part(&partname).expect("saved part").blob(),
            b"payload"
        );
    }

    #[test]
    fn invalid_packages_never_replace_an_existing_artifact() {
        let directory = tempfile::tempdir().expect("temporary directory");
        let destination = directory.path().join("package.xlsx");
        std::fs::write(&destination, b"previous artifact").expect("seed destination");
        let mut package = OpcPackage::new();
        package.add_part(Box::new(crate::BlobPart::new(
            PackURI::new("/custom/data.bin").expect("valid part URI"),
            "invalid content type".to_owned(),
            Vec::new(),
        )));

        let result = PackageWriter::write(&destination, &package);

        assert!(matches!(
            result,
            Err(crate::OpcError::InvalidContentType { .. })
        ));
        assert_eq!(
            std::fs::read(destination).expect("read destination"),
            b"previous artifact"
        );
    }
}
