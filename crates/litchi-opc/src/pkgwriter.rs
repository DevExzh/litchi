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
use crate::rel::Relationships;
use litchi_core::xml::escape_xml;
use std::collections::HashMap;
use std::fmt::Write as _;
use std::io::Write;
use std::path::Path;

const EXACT_SOURCE_CHUNK_BYTES: usize = 64 * 1024;

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

/// Fully validated package metadata and the stable order used for publication.
///
/// Building this plan is deliberately separate from emission: every fallible
/// serialization and XML audit completes before a sequential sink sees bytes.
struct PublicationPlan<'package> {
    content_types_uri: PackURI,
    content_types_xml: String,
    package_rels_uri: PackURI,
    package_rels_xml: String,
    parts: Vec<PlannedPart<'package>>,
}

struct PlannedPart<'package> {
    partname: &'package PackURI,
    content_type: &'package str,
    blob: &'package [u8],
    authored_xml: bool,
    rels: &'package Relationships,
    relationships: Option<PlannedRelationships>,
}

struct PlannedRelationships {
    uri: PackURI,
    xml: String,
}

impl<'package> PublicationPlan<'package> {
    fn from_package(package: &'package OpcPackage) -> Result<Self> {
        let mut parts = Vec::new();
        parts
            .try_reserve_exact(package.part_count())
            .map_err(|source| crate::OpcError::Allocation {
                resource: "OPC XML publication part plan",
                source,
            })?;
        for part in package.iter_parts() {
            parts.push(PlannedPart {
                partname: part.partname(),
                content_type: part.content_type(),
                blob: part.blob(),
                authored_xml: xml_minifier::audit::package::is_xml_part(
                    part.partname().as_str(),
                    part.content_type(),
                ) && !package.is_exact_source_xml(part),
                rels: part.rels(),
                relationships: None,
            });
        }
        parts.sort_unstable_by(|left, right| left.partname.as_str().cmp(right.partname.as_str()));

        let content_types_uri =
            PackURI::new(CONTENT_TYPES_URI).map_err(crate::OpcError::InvalidPackUri)?;
        let content_types_xml = ContentTypesItem::from_parts(&parts)?.to_xml();
        PackageWriter::validate_authored_xml("[Content_Types].xml", content_types_xml.as_bytes())?;

        let package_uri = PackURI::new(PACKAGE_URI).map_err(crate::OpcError::InvalidPackUri)?;
        let package_rels_uri = package_uri
            .rels_uri()
            .map_err(crate::OpcError::InvalidPackUri)?;
        let package_rels_xml = package.rels().to_xml();
        PackageWriter::validate_authored_xml("_rels/.rels", package_rels_xml.as_bytes())?;

        for part in &mut parts {
            if part.authored_xml {
                PackageWriter::validate_authored_xml(part.partname.as_str(), part.blob)?;
            }
            if !part.rels.is_empty() {
                let uri = part
                    .partname
                    .rels_uri()
                    .map_err(crate::OpcError::InvalidPackUri)?;
                let xml = part.rels.to_xml();
                PackageWriter::validate_authored_xml(uri.as_str(), xml.as_bytes())?;
                part.relationships = Some(PlannedRelationships { uri, xml });
            }
        }

        Ok(Self {
            content_types_uri,
            content_types_xml,
            package_rels_uri,
            package_rels_xml,
            parts,
        })
    }

    fn write<W: Write>(&self, physical: &mut PhysPkgWriter<W>) -> Result<()> {
        physical.write(&self.content_types_uri, self.content_types_xml.as_bytes())?;
        physical.write(&self.package_rels_uri, self.package_rels_xml.as_bytes())?;
        for part in &self.parts {
            physical.write(part.partname, part.blob)?;
            if let Some(relationships) = &part.relationships {
                physical.write(&relationships.uri, relationships.xml.as_bytes())?;
            }
        }
        Ok(())
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
    ///
    /// # Errors
    /// Returns an error if the package cannot be serialized (for example an
    /// invalid content type or partname) or if writing to the filesystem fails.
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
    ///
    /// # Errors
    /// Returns an error if the package cannot be serialized (for example an
    /// invalid content type or partname) or if the sink rejects a write. When
    /// the sink has already accepted bytes, the error is wrapped in
    /// [`crate::OpcError::IncompleteOutput`] with the accepted byte count.
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
            other => other,
        }
    }

    fn write_counted<W: Write>(writer: W, package: &OpcPackage) -> Result<()> {
        if let Some(source) = package.exact_source() {
            let mut writer = writer;
            for chunk in source.chunks(EXACT_SOURCE_CHUNK_BYTES) {
                writer.write_all(chunk)?;
            }
            writer.flush()?;
            return Ok(());
        }
        let plan = PublicationPlan::from_package(package)?;
        let mut physical = PhysPkgWriter::with_writer(writer);
        plan.write(&mut physical)?;
        let mut finished = physical.finish_into_inner()?;
        finished.flush()?;
        Ok(())
    }

    /// Serialize an OPC package to bytes.
    ///
    /// # Arguments
    /// * `package` - The OPC package to serialize
    ///
    /// # Returns
    /// The serialized package as a byte vector
    ///
    /// # Errors
    /// Returns an error if the package cannot be serialized (for example an
    /// invalid content type or partname) or if the in-memory zip writer fails.
    pub fn to_bytes(package: &OpcPackage) -> Result<Vec<u8>> {
        if let Some(source) = package.exact_source() {
            let mut bytes = Vec::new();
            bytes.try_reserve_exact(source.len()).map_err(|source| {
                crate::OpcError::Allocation {
                    resource: "OPC exact source copy",
                    source,
                }
            })?;
            bytes.extend_from_slice(source);
            return Ok(bytes);
        }
        let plan = PublicationPlan::from_package(package)?;
        let mut physical = PhysPkgWriter::new();
        plan.write(&mut physical)?;
        physical.finish()
    }

    fn validate_authored_xml(name: &str, bytes: &[u8]) -> Result<()> {
        xml_minifier::audit::verify_authored(bytes, xml_minifier::audit::Limits::default())
            .map(|_report| ())
            .map_err(|source| crate::OpcError::XmlPublication {
                part: name.to_string(),
                source,
            })
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

    /// Build `ContentTypesItem` from the sorted publication parts.
    fn from_parts(parts: &[PlannedPart<'_>]) -> Result<Self> {
        let mut cti = Self::new()?;

        for part in parts {
            cti.add_content_type(part.partname, part.content_type)?;
        }

        Ok(cti)
    }

    /// Add a content type for a part.
    ///
    /// Uses a default mapping if the extension matches a well-known type,
    /// otherwise uses an override for the specific partname.
    fn add_content_type(&mut self, partname: &PackURI, content_type: &str) -> Result<()> {
        let ext = partname.ext().to_ascii_lowercase();
        let parsed_content_type = ContentType::new(content_type)?;

        // Check if this is a standard default mapping
        if Self::is_default_content_type(&ext, parsed_content_type.as_str()) {
            self.defaults.insert(ext, parsed_content_type);
        } else {
            self.overrides
                .insert(partname.to_string(), parsed_content_type);
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
                | ("jpg" | "jpeg", "image/jpeg")
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

    fn with_eocd_comment(mut archive: Vec<u8>, comment: &[u8]) -> Vec<u8> {
        let comment_len = u16::try_from(comment.len()).expect("ZIP comment fits in EOCD");
        let eocd = archive.len().checked_sub(22).expect("archive has an EOCD");
        assert_eq!(&archive[eocd..eocd + 4], b"PK\x05\x06");
        archive[eocd + 20..eocd + 22].copy_from_slice(&comment_len.to_le_bytes());
        archive.extend_from_slice(comment);
        archive
    }

    fn exact_empty_archive(comment: &[u8]) -> Vec<u8> {
        with_eocd_comment(
            PackageWriter::to_bytes(&OpcPackage::new()).expect("serialize source package"),
            comment,
        )
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
    fn owned_source_round_trips_exactly_to_bytes_and_stream() {
        let source = exact_empty_archive(b"nonzero EOCD comment");
        let package = OpcPackage::from_vec(source.clone()).expect("open owned source");

        assert_eq!(
            PackageWriter::to_bytes(&package).expect("copy exact source"),
            source
        );

        let mut streamed = Vec::new();
        package
            .to_stream(&mut streamed)
            .expect("stream exact source");
        assert_eq!(streamed, source);
    }

    #[test]
    fn open_and_reader_retain_owned_source_but_borrowed_bytes_do_not() {
        let source = exact_empty_archive(b"owned source");

        let from_reader =
            OpcPackage::from_reader(io::Cursor::new(source.clone())).expect("open reader source");
        assert_eq!(
            PackageWriter::to_bytes(&from_reader).expect("copy reader source"),
            source
        );

        let directory = tempfile::tempdir().expect("temporary directory");
        let path = directory.path().join("source.opc");
        std::fs::write(&path, &source).expect("write source package");
        let opened = OpcPackage::open(path).expect("open file source");
        assert_eq!(
            PackageWriter::to_bytes(&opened).expect("copy file source"),
            source
        );

        let borrowed = OpcPackage::from_bytes(&source).expect("open borrowed source");
        assert!(borrowed.exact_source().is_none());
        assert_ne!(
            PackageWriter::to_bytes(&borrowed).expect("republish borrowed source"),
            source
        );
    }

    #[test]
    fn part_edit_revokes_exact_source_and_republishes_edited_package() {
        let partname = PackURI::new("/custom/metadata.xml").expect("valid part URI");
        let mut source_package = OpcPackage::new();
        source_package.add_part(Box::new(crate::BlobPart::new(
            partname.clone(),
            ct::XML.to_owned(),
            b"<original/>".to_vec(),
        )));
        let source = with_eocd_comment(
            PackageWriter::to_bytes(&source_package).expect("serialize source package"),
            b"exact source",
        );
        let mut package = OpcPackage::from_vec(source.clone()).expect("open owned source");

        package
            .get_part_mut(&partname)
            .expect("editable part")
            .set_blob(b"<edited/>".to_vec());
        let rewritten = PackageWriter::to_bytes(&package).expect("republish edited package");

        assert_ne!(rewritten, source);
        let reopened = OpcPackage::from_bytes(&rewritten).expect("reopen edited package");
        assert_eq!(
            reopened.get_part(&partname).expect("edited part").blob(),
            b"<edited/>"
        );
    }

    #[test]
    fn exact_source_streaming_respects_the_64_kib_chunk_bound() {
        let comment = vec![b'x'; u16::MAX as usize];
        let source = exact_empty_archive(&comment);
        assert!(source.len() > EXACT_SOURCE_CHUNK_BYTES);
        let package = OpcPackage::from_vec(source.clone()).expect("open owned source");
        let mut sink = ChunkSink {
            total: 0,
            writes: 0,
            largest: 0,
            limit: EXACT_SOURCE_CHUNK_BYTES,
        };

        PackageWriter::write_to_stream(&mut sink, &package).expect("stream exact source");

        assert_eq!(sink.total, source.len());
        assert!(sink.writes > 1);
        assert!(sink.largest <= EXACT_SOURCE_CHUNK_BYTES);
    }

    #[test]
    fn incomplete_exact_source_stream_reports_accepted_bytes() {
        let source = exact_empty_archive(b"exact source");
        let package = OpcPackage::from_vec(source).expect("open owned source");
        let sink = FailAfter {
            written: 0,
            limit: 128,
        };

        let error = PackageWriter::write_to_stream(sink, &package)
            .expect_err("bounded sink must reject exact source");

        assert!(matches!(
            error,
            crate::OpcError::IncompleteOutput {
                written: 128,
                source,
            } if matches!(*source, crate::OpcError::IoError(_))
        ));
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
    fn refuses_arbitrary_authored_xml_bytes_before_publication() {
        for (part_name, content_type) in [
            ("/custom/manifest.rdf", "application/octet-stream"),
            ("/custom/metadata", "application/rdf+xml"),
            (
                "/_xmlsignatures/sig1.bin",
                "application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml",
            ),
        ] {
            let mut package = OpcPackage::new();
            package.add_part(Box::new(crate::BlobPart::new(
                PackURI::new(part_name).expect("valid part URI"),
                content_type.to_string(),
                b"<root> <child/></root>".to_vec(),
            )));

            assert!(matches!(
                PackageWriter::to_bytes(&package),
                Err(crate::OpcError::XmlPublication { part, .. }) if part == part_name
            ));
        }
    }

    #[test]
    fn publication_plan_failure_leaves_sequential_sink_untouched() {
        let mut package = OpcPackage::new();
        package.add_part(Box::new(crate::BlobPart::new(
            PackURI::new("/custom/metadata.xml").expect("valid part URI"),
            ct::XML.to_owned(),
            b"<root> <child/></root>".to_vec(),
        )));
        let mut sink = ChunkSink {
            total: 0,
            writes: 0,
            largest: 0,
            limit: usize::MAX,
        };

        let error = PackageWriter::write_to_stream(&mut sink, &package)
            .expect_err("authored XML must fail publication planning");

        assert!(matches!(
            error,
            crate::OpcError::XmlPublication { part, .. }
                if part == "/custom/metadata.xml"
        ));
        assert_eq!(sink.total, 0);
        assert_eq!(sink.writes, 0);
    }

    #[test]
    fn exact_source_xml_bytes_may_remain_opaque() {
        let content_types = br#"<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/custom/manifest.rdf" ContentType="application/rdf+xml"/></Types>"#;
        let relationships = br#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>"#;
        let source_rdf = b"<rdf:RDF xmlns:rdf=\"urn:test\">\n <rdf:Description/>\n</rdf:RDF>";
        let mut physical = PhysPkgWriter::new();
        physical
            .write(
                &PackURI::new("/[Content_Types].xml").expect("content-types URI"),
                content_types,
            )
            .expect("write content types");
        physical
            .write(
                &PackURI::new("/_rels/.rels").expect("relationship URI"),
                relationships,
            )
            .expect("write relationships");
        physical
            .write(
                &PackURI::new("/custom/manifest.rdf").expect("RDF URI"),
                source_rdf,
            )
            .expect("write source RDF");
        let source = physical.finish().expect("finish source package");

        let package = OpcPackage::from_vec(source).expect("open source package");
        let rewritten = PackageWriter::to_bytes(&package).expect("preserve source RDF");
        let rewritten_physical = crate::phys_pkg::OwnedPhysPkgReader::from_bytes(rewritten)
            .expect("open rewritten package");
        assert_eq!(
            rewritten_physical
                .read_member("custom/manifest.rdf")
                .expect("read rewritten RDF"),
            source_rdf
        );
    }

    #[test]
    fn real_package_enumeration_covers_all_xml_bearing_members() {
        let mut package = OpcPackage::new();
        for (part_name, content_type, payload) in [
            (
                "/custom/manifest.rdf",
                "application/rdf+xml",
                b"<rdf:RDF xmlns:rdf=\"urn:test\"/>".as_slice(),
            ),
            (
                "/custom/metadata",
                "application/vnd.example.metadata+xml",
                b"<metadata/>".as_slice(),
            ),
            (
                "/_xmlsignatures/sig1.xml",
                "application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml",
                b"<Signature xmlns=\"http://www.w3.org/2000/09/xmldsig#\"/>".as_slice(),
            ),
        ] {
            package.add_part(Box::new(crate::BlobPart::new(
                PackURI::new(part_name).expect("valid part URI"),
                content_type.to_string(),
                payload.to_vec(),
            )));
        }
        let bytes = PackageWriter::to_bytes(&package).expect("publish package");
        let physical =
            crate::phys_pkg::OwnedPhysPkgReader::from_bytes(bytes).expect("open published package");
        let mut audited = Vec::new();
        for name in physical.member_names().expect("enumerate members") {
            let media_type = match name.as_str() {
                "custom/metadata" => "application/vnd.example.metadata+xml",
                "custom/manifest.rdf" => "application/rdf+xml",
                "_xmlsignatures/sig1.xml" => {
                    "application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml"
                },
                _ => "application/octet-stream",
            };
            if xml_minifier::audit::package::is_xml_part(&name, media_type) {
                let payload = physical.read_member(&name).expect("read XML member");
                let _report = xml_minifier::audit::verify_authored(
                    &payload,
                    xml_minifier::audit::Limits::default(),
                )
                .expect("emitted XML is compact");
                audited.push(name);
            }
        }
        audited.sort();
        assert_eq!(
            audited,
            [
                "[Content_Types].xml",
                "_rels/.rels",
                "_xmlsignatures/sig1.xml",
                "custom/manifest.rdf",
                "custom/metadata",
            ]
        );
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
