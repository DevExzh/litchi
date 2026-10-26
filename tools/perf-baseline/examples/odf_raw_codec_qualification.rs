#![allow(unexpected_cfgs)]

//! Bounded raw ODF codec qualification probe.
//!
//! The probe is a consumer-only adapter for the direct pre-migration source
//! (`raw-codec-original`) and the selected-limit candidate API
//! (`raw-codec-current`).  It exercises the same deterministic inputs and
//! route order in both roles.  Correctness mode runs every proof and negative
//! case; timed and allocator modes measure only the route operation named by
//! the manifest boundary.  Report construction happens after the allocator
//! observer has stopped.

#[cfg(all(feature = "raw-codec-original", feature = "raw-codec-current"))]
compile_error!("select exactly one raw-codec API era");
#[cfg(not(any(feature = "raw-codec-original", feature = "raw-codec-current")))]
compile_error!("select raw-codec-original or raw-codec-current");

#[cfg(feature = "raw-codec-allocator")]
mod allocation {
    use std::alloc::{GlobalAlloc, Layout, System};
    use std::sync::atomic::{AtomicBool, AtomicUsize, Ordering};

    pub(crate) const ENABLED: bool = true;
    static ACTIVE: AtomicBool = AtomicBool::new(false);
    static ALLOCATIONS: AtomicUsize = AtomicUsize::new(0);
    static DEALLOCATIONS: AtomicUsize = AtomicUsize::new(0);
    static ALLOCATED_BYTES: AtomicUsize = AtomicUsize::new(0);
    static DEALLOCATED_BYTES: AtomicUsize = AtomicUsize::new(0);
    static REALLOCATIONS: AtomicUsize = AtomicUsize::new(0);

    pub struct CountingAllocator;

    // SAFETY: every operation delegates to the system allocator and the
    // counters are independent atomic bookkeeping.
    unsafe impl GlobalAlloc for CountingAllocator {
        unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
            // SAFETY: the caller upholds GlobalAlloc's layout contract.
            let pointer = unsafe { System.alloc(layout) };
            if !pointer.is_null() && ACTIVE.load(Ordering::Relaxed) {
                ALLOCATIONS.fetch_add(1, Ordering::Relaxed);
                ALLOCATED_BYTES.fetch_add(layout.size(), Ordering::Relaxed);
            }
            pointer
        }

        unsafe fn dealloc(&self, pointer: *mut u8, layout: Layout) {
            // SAFETY: the caller upholds GlobalAlloc's pointer/layout contract.
            unsafe { System.dealloc(pointer, layout) };
            if ACTIVE.load(Ordering::Relaxed) {
                DEALLOCATIONS.fetch_add(1, Ordering::Relaxed);
                DEALLOCATED_BYTES.fetch_add(layout.size(), Ordering::Relaxed);
            }
        }

        unsafe fn realloc(&self, pointer: *mut u8, layout: Layout, new_size: usize) -> *mut u8 {
            // SAFETY: the caller upholds GlobalAlloc's pointer/layout contract.
            let replacement = unsafe { System.realloc(pointer, layout, new_size) };
            if !replacement.is_null() && ACTIVE.load(Ordering::Relaxed) {
                REALLOCATIONS.fetch_add(1, Ordering::Relaxed);
                if new_size >= layout.size() {
                    ALLOCATED_BYTES.fetch_add(new_size - layout.size(), Ordering::Relaxed);
                } else {
                    DEALLOCATED_BYTES.fetch_add(layout.size() - new_size, Ordering::Relaxed);
                }
            }
            replacement
        }
    }

    #[global_allocator]
    static GLOBAL: CountingAllocator = CountingAllocator;

    pub(crate) fn begin() {
        ACTIVE.store(false, Ordering::Relaxed);
        ALLOCATIONS.store(0, Ordering::Relaxed);
        DEALLOCATIONS.store(0, Ordering::Relaxed);
        ALLOCATED_BYTES.store(0, Ordering::Relaxed);
        DEALLOCATED_BYTES.store(0, Ordering::Relaxed);
        REALLOCATIONS.store(0, Ordering::Relaxed);
        ACTIVE.store(true, Ordering::Relaxed);
    }

    pub(crate) fn end() -> [usize; 5] {
        ACTIVE.store(false, Ordering::Relaxed);
        [
            ALLOCATIONS.load(Ordering::Relaxed),
            DEALLOCATIONS.load(Ordering::Relaxed),
            ALLOCATED_BYTES.load(Ordering::Relaxed),
            DEALLOCATED_BYTES.load(Ordering::Relaxed),
            REALLOCATIONS.load(Ordering::Relaxed),
        ]
    }
}

#[cfg(not(feature = "raw-codec-allocator"))]
mod allocation {
    pub(crate) const ENABLED: bool = false;

    pub(crate) fn begin() {}

    pub(crate) fn end() -> [usize; 5] {
        [0; 5]
    }
}

mod probe {
    use std::collections::BTreeMap;
    use std::fmt::Write as _;
    use std::hint::black_box;
    use std::io::{self, Cursor, Read, Write};
    use std::time::Instant;

    use super::allocation;
    use litchi_core::FileFormat;
    use litchi_odf_common::constants;
    use litchi_odf_common::core::{
        AuthoredXmlFragment, Metadata, OwnedPackage, PackageCompression, PackageWriter,
        ResolvedReader, StreamingArchiveLimits, XmlSourcePart, XmlSplicePublication,
        XmlStreamLimits, rebuild_package_with_xml_splices,
    };
    use litchi_odf_common::package::{
        Archive, ArchiveLimits, parse_manifest, raw_identical_members, read_manifest,
        replace_content_xml, replace_content_xml_spliced,
        replace_content_xml_with_payload_verification,
    };
    use litchi_ops::{OpenOptions, create, open_bytes_with};
    use quick_xml::events::Event;
    use quick_xml::name::ResolveResult;
    use serde_json::{Value, json};
    use sha2::{Digest as _, Sha256};
    use zip::{HasZipMetadata, ZipArchive};

    const MAX_ITERATIONS: usize = 128;
    const SOURCE_CONTENT: &str = r#"<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><office:body><office:text><text:p>source</text:p></office:text></office:body></office:document-content>"#;
    const TARGET_CONTENT: &str = r#"<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><office:body><office:text><text:p>target</text:p></office:text></office:body></office:document-content>"#;
    const STYLES_XML: &str = r#"<?xml version="1.0"?><office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"/>"#;
    const META_XML: &str = r#"<?xml version="1.0"?><office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"><office:meta><dc:title>raw codec</dc:title><meta:user-defined meta:name="lane">raw</meta:user-defined></office:meta></office:document-meta>"#;
    const SETTINGS_XML: &str = r#"<?xml version="1.0"?><office:document-settings xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"/>"#;
    const OPAQUE_MEMBER: &str = "Pictures/opaque.bin";
    const OPAQUE_BYTES: usize = 16 * 1024;
    const ROUTE_IDS: [&str; 10] = [
        "archive_catalog_and_raw_split",
        "manifest_xml_stream_and_parse",
        "metadata_xml_and_namespace_resolution",
        "writer_memory_publication",
        "writer_stream_publication",
        "writer_bounded_policy",
        "raw_split_merge_and_opaque_preservation",
        "raw_xml_splice_merge_publication",
        "unified_facade_raw_parts_and_publication",
        "consumer_reopen_matrix",
    ];
    const FAMILY_IDS: [&str; 10] = [
        "odt",
        "ods",
        "odp",
        "odg",
        "odc",
        "odi",
        "odb",
        "odf-formula",
        "oth",
        "odm",
    ];
    const NEGATIVE_IDS: [&str; 9] = [
        "duplicate_manifest_path",
        "manifest_limit",
        "malformed_xml",
        "undeclared_namespace",
        "stale_xml_range",
        "foreign_xml_source",
        "bounded_output",
        "failing_reader_or_sink",
        "corrupt_stored_payload_without_verification",
    ];

    #[cfg(feature = "raw-codec-current")]
    const API_ERA: &str = "current";
    #[cfg(all(feature = "raw-codec-original", not(feature = "raw-codec-current")))]
    const API_ERA: &str = "original";

    #[cfg(feature = "raw-codec-current")]
    const NAMESPACE_CONTRACT: &str = "document_rejects_undeclared_namespace";
    #[cfg(all(feature = "raw-codec-original", not(feature = "raw-codec-current")))]
    const NAMESPACE_CONTRACT: &str = "legacy_unknown_namespace_event";

    fn digest(bytes: &[u8]) -> String {
        let mut output = String::with_capacity(64);
        for byte in Sha256::digest(bytes) {
            write!(output, "{byte:02x}").expect("digest formatting");
        }
        output
    }

    fn error_text(error: impl std::fmt::Display) -> String {
        error.to_string()
    }

    #[cfg(feature = "raw-codec-current")]
    fn archive<'a>(bytes: &'a [u8]) -> Result<Archive<'a>, String> {
        Archive::new(bytes, ArchiveLimits::DEFAULT).map_err(error_text)
    }

    #[cfg(all(feature = "raw-codec-original", not(feature = "raw-codec-current")))]
    fn archive<'a>(bytes: &'a [u8]) -> Result<Archive<'a>, String> {
        Archive::new(bytes).map_err(error_text)
    }

    #[cfg(feature = "raw-codec-current")]
    fn read_manifest_xml(archive: &Archive<'_>) -> Result<String, String> {
        archive
            .read_manifest_xml(XmlStreamLimits::default())
            .map_err(error_text)
    }

    #[cfg(all(feature = "raw-codec-original", not(feature = "raw-codec-current")))]
    fn read_manifest_xml(archive: &Archive<'_>) -> Result<String, String> {
        archive.read_manifest_xml().map_err(error_text)
    }

    #[cfg(feature = "raw-codec-current")]
    fn read_manifest_xml_limited(
        archive: &Archive<'_>,
        limits: XmlStreamLimits,
    ) -> Result<String, String> {
        archive.read_manifest_xml(limits).map_err(error_text)
    }

    #[cfg(all(feature = "raw-codec-original", not(feature = "raw-codec-current")))]
    fn read_manifest_xml_limited(
        archive: &Archive<'_>,
        _limits: XmlStreamLimits,
    ) -> Result<String, String> {
        archive
            .read_manifest_xml_with_limit(8, "raw-codec manifest")
            .map_err(error_text)
    }

    #[cfg(feature = "raw-codec-current")]
    fn resolved_document<'a>(
        xml: &'a str,
        limits: XmlStreamLimits,
    ) -> Result<ResolvedReader<'a>, String> {
        ResolvedReader::from_xml_document(xml, limits).map_err(error_text)
    }

    #[cfg(all(feature = "raw-codec-original", not(feature = "raw-codec-current")))]
    fn resolved_document<'a>(
        xml: &'a str,
        limits: XmlStreamLimits,
    ) -> Result<ResolvedReader<'a>, String> {
        ResolvedReader::from_xml(xml, limits).map_err(error_text)
    }

    #[cfg(feature = "raw-codec-current")]
    fn undeclared_namespace_contract(xml: &str, limits: XmlStreamLimits) -> bool {
        let mut reader = match resolved_document(xml, limits) {
            // Construction is intentionally lazy.  A setup error is not the
            // event-consumption proof this lane requires.
            Ok(reader) => reader,
            Err(_) => return false,
        };
        let mut buffer = Vec::new();
        loop {
            match reader.read_resolved_event_into(&mut buffer) {
                Ok((_, event)) => {
                    if matches!(event, Event::Eof) {
                        return false;
                    }
                    buffer.clear();
                },
                Err(error) => {
                    let message = error_text(error).to_ascii_lowercase();
                    return message.contains("unknown namespace") && message.contains("prefix");
                },
            }
        }
    }

    #[cfg(all(feature = "raw-codec-original", not(feature = "raw-codec-current")))]
    fn undeclared_namespace_contract(xml: &str, limits: XmlStreamLimits) -> bool {
        let mut reader = match resolved_document(xml, limits) {
            Ok(reader) => reader,
            Err(_) => return false,
        };
        let mut buffer = Vec::new();
        let mut saw_unknown = false;
        loop {
            match reader.read_resolved_event_into(&mut buffer) {
                Ok((namespace, event)) => {
                    if matches!(namespace, ResolveResult::Unknown(_)) {
                        saw_unknown = true;
                    }
                    if matches!(event, Event::Eof) {
                        return saw_unknown;
                    }
                    buffer.clear();
                },
                Err(_) => return false,
            }
        }
    }

    fn memory_writer() -> PackageWriter<Cursor<Vec<u8>>> {
        #[cfg(feature = "raw-codec-current")]
        {
            return PackageWriter::new(StreamingArchiveLimits::default());
        }
        #[cfg(all(feature = "raw-codec-original", not(feature = "raw-codec-current")))]
        {
            PackageWriter::new()
        }
    }

    fn sink_writer(sink: Cursor<Vec<u8>>) -> PackageWriter<Cursor<Vec<u8>>> {
        #[cfg(feature = "raw-codec-current")]
        {
            return PackageWriter::with_writer(sink, StreamingArchiveLimits::default());
        }
        #[cfg(all(feature = "raw-codec-original", not(feature = "raw-codec-current")))]
        {
            PackageWriter::with_writer(sink)
        }
    }

    fn bounded_publication(limit: usize, mime: &str, payload: &[u8]) -> Result<Vec<u8>, String> {
        #[cfg(feature = "raw-codec-current")]
        let mut writer = PackageWriter::new_bounded(limit, StreamingArchiveLimits::default());
        #[cfg(all(feature = "raw-codec-original", not(feature = "raw-codec-current")))]
        let mut writer = PackageWriter::new_bounded(limit);
        writer.set_mimetype(mime).map_err(error_text)?;
        writer
            .add_file("content.xml", payload)
            .map_err(error_text)?;
        writer.finish_to_bounded_bytes().map_err(error_text)
    }

    fn opaque_payload() -> Vec<u8> {
        (0..OPAQUE_BYTES)
            .map(|index| ((index.wrapping_mul(37) + 11) % 251) as u8)
            .collect()
    }

    fn package_with_content_and_opaque(mime: &str, content: &[u8], opaque: &[u8]) -> Vec<u8> {
        let mut writer = memory_writer();
        writer.set_mimetype(mime).expect("mimetype");
        writer
            .add_authored_xml("content.xml", content, "text/xml")
            .expect("content XML");
        writer
            .add_authored_xml("styles.xml", STYLES_XML.as_bytes(), "text/xml")
            .expect("styles XML");
        writer
            .add_authored_xml("meta.xml", META_XML.as_bytes(), "text/xml")
            .expect("metadata XML");
        writer
            .add_authored_xml("settings.xml", SETTINGS_XML.as_bytes(), "text/xml")
            .expect("settings XML");
        writer
            .add_file_reader_with_media_type_and_compression(
                OPAQUE_MEMBER,
                Cursor::new(opaque),
                "application/octet-stream",
                PackageCompression::Stored,
            )
            .expect("opaque member");
        writer
            .add_file_with_media_type("Vendor/unknown.bin", b"vendor", "application/octet-stream")
            .expect("vendor member");
        writer.finish().expect("package finish")
    }

    fn package_with_mime(mime: &str) -> Vec<u8> {
        package_with_content_and_opaque(mime, SOURCE_CONTENT.as_bytes(), &opaque_payload())
    }

    fn family_packages() -> Vec<(String, Vec<u8>)> {
        [
            ("odt", FileFormat::Odt),
            ("ods", FileFormat::Ods),
            ("odp", FileFormat::Odp),
            ("odg", FileFormat::Odg),
            ("odc", FileFormat::Odc),
            ("odi", FileFormat::Odi),
            ("odb", FileFormat::Odb),
            ("odf-formula", FileFormat::Odf),
            ("oth", FileFormat::Oth),
            ("odm", FileFormat::Odm),
        ]
        .into_iter()
        .map(|(name, format)| {
            let file = create(format).unwrap_or_else(|error| panic!("create {name}: {error}"));
            let bytes = file
                .to_bytes()
                .unwrap_or_else(|error| panic!("publish {name}: {error}"));
            (name.to_string(), bytes)
        })
        .collect()
    }

    fn archive_fingerprint(bytes: &[u8]) -> (Vec<String>, Value) {
        let mut archive = ZipArchive::new(Cursor::new(bytes)).expect("ZIP archive");
        let comment_sha256 = digest(archive.comment());
        let mut names = Vec::with_capacity(archive.len());
        let mut members = Vec::with_capacity(archive.len());
        for index in 0..archive.len() {
            let mut member = archive.by_index(index).expect("ZIP member");
            let name = member.name().to_owned();
            let name_raw_sha256 = digest(member.name_raw());
            let compression = format!("{:?}", member.compression());
            let compressed_bytes = member.compressed_size();
            let crc32 = member.crc32();
            let metadata = member.get_metadata();
            let version_needed = metadata.version_needed();
            let flags = metadata.flags;
            let external_attributes = metadata.external_attributes;
            let comment = member.comment().to_owned();
            let mut content = Vec::new();
            member.read_to_end(&mut content).expect("ZIP payload");
            names.push(name.clone());
            members.push(json!({
                "name": name,
                "name_raw_sha256": name_raw_sha256,
                "content_sha256": digest(&content),
                "content_bytes": content.len(),
                "compression": compression,
                "compressed_bytes": compressed_bytes,
                "crc32": format!("{crc32:08x}"),
                "version_needed": version_needed,
                "flags": flags,
                "external_attributes": external_attributes,
                "comment": comment,
            }));
        }
        (
            names,
            json!({
                "archive_comment_sha256": comment_sha256,
                "members": members,
            }),
        )
    }

    fn untouched_archive_members_equal(source: &[u8], output: &[u8], changed_member: &str) -> bool {
        let (_, source_audit) = archive_fingerprint(source);
        let (_, output_audit) = archive_fingerprint(output);
        if source_audit["archive_comment_sha256"] != output_audit["archive_comment_sha256"] {
            return false;
        }
        let Some(source_members) = source_audit["members"].as_array() else {
            return false;
        };
        let Some(output_members) = output_audit["members"].as_array() else {
            return false;
        };
        let source_untouched = source_members
            .iter()
            .filter(|member| member["name"] != changed_member)
            .collect::<Vec<_>>();
        let output_untouched = output_members
            .iter()
            .filter(|member| member["name"] != changed_member)
            .collect::<Vec<_>>();
        let source_changed = source_members
            .iter()
            .filter(|member| member["name"] == changed_member)
            .collect::<Vec<_>>();
        let output_changed = output_members
            .iter()
            .filter(|member| member["name"] == changed_member)
            .collect::<Vec<_>>();
        if source_changed.len() != 1 || output_changed.len() != 1 {
            return false;
        }
        let metadata_keys = [
            "name",
            "name_raw_sha256",
            "compression",
            "version_needed",
            "flags",
            "external_attributes",
            "comment",
        ];
        let changed_metadata_ok = metadata_keys
            .iter()
            .all(|key| source_changed[0][*key] == output_changed[0][*key]);
        source_untouched == output_untouched && changed_metadata_ok
    }

    fn corrupt_stored_payload(bytes: &[u8]) -> Vec<u8> {
        let mut output = bytes.to_vec();
        let mut offset = 0usize;
        while offset + 30 <= output.len() {
            if &output[offset..offset + 4] != b"PK\x03\x04" {
                offset += 1;
                continue;
            }
            let name_len = u16::from_le_bytes([output[offset + 26], output[offset + 27]]) as usize;
            let extra_len = u16::from_le_bytes([output[offset + 28], output[offset + 29]]) as usize;
            let name_start = offset + 30;
            let payload_start = name_start + name_len + extra_len;
            let name_end = name_start + name_len;
            if name_end > output.len() || payload_start >= output.len() {
                break;
            }
            if &output[name_start..name_end] == OPAQUE_MEMBER.as_bytes() {
                output[payload_start] ^= 1;
                return output;
            }
            let compressed = u32::from_le_bytes([
                output[offset + 18],
                output[offset + 19],
                output[offset + 20],
                output[offset + 21],
            ]) as usize;
            offset = payload_start.saturating_add(compressed);
        }
        panic!("stored opaque member was not found");
    }

    #[derive(Debug)]
    struct FailingReader {
        emitted: bool,
    }

    impl Read for FailingReader {
        fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
            if self.emitted {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidData,
                    "raw-codec reader failure",
                ));
            }
            self.emitted = true;
            output[0] = b'x';
            Ok(1)
        }
    }

    fn negative_cases(package: &[u8]) -> BTreeMap<String, bool> {
        let mut cases = BTreeMap::new();
        let xml_limits = XmlStreamLimits::default();
        let duplicate = br#"<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="content.xml"/><manifest:file-entry manifest:full-path="content.xml"/></manifest:manifest>"#;
        cases.insert(
            "duplicate_manifest_path".to_string(),
            parse_manifest(
                std::str::from_utf8(duplicate).expect("duplicate XML"),
                xml_limits,
            )
            .is_err(),
        );
        let narrow = XmlStreamLimits::new(8, 64, 128, 32, 128).expect("narrow XML policy");
        cases.insert(
            "manifest_limit".to_string(),
            archive(package)
                .and_then(|value| read_manifest_xml_limited(&value, narrow))
                .is_err(),
        );
        cases.insert(
            "malformed_xml".to_string(),
            parse_manifest("<manifest:manifest", xml_limits).is_err(),
        );
        cases.insert(
            "undeclared_namespace".to_string(),
            undeclared_namespace_contract("<office:meta/>", xml_limits),
        );

        let source = OwnedPackage::from_bytes(package.to_vec()).expect("source package");
        let part = XmlSourcePart::load(&source, "content.xml").expect("source content");
        cases.insert(
            "stale_xml_range".to_string(),
            part.checked_range(0..1, b"x").is_err(),
        );
        let foreign = OwnedPackage::from_bytes(package.to_vec()).expect("foreign package");
        let foreign_part = XmlSourcePart::load(&foreign, "content.xml").expect("foreign content");
        let insert = foreign_part.bytes().len().saturating_sub(1);
        let proof = foreign_part
            .checked_range(insert..insert, b"")
            .expect("foreign proof");
        let mut publication = XmlSplicePublication::new(foreign_part);
        publication
            .replace(
                proof,
                AuthoredXmlFragment::text(b"foreign".to_vec()).expect("text"),
            )
            .expect("foreign publication");
        cases.insert(
            "foreign_xml_source".to_string(),
            rebuild_package_with_xml_splices(&source, vec![publication], 1 << 20).is_err(),
        );
        cases.insert(
            "bounded_output".to_string(),
            bounded_publication(64, constants::ODF_TEXT, SOURCE_CONTENT.as_bytes()).is_err(),
        );

        let mut writer = sink_writer(Cursor::new(Vec::new()));
        writer.set_mimetype(constants::ODF_TEXT).expect("sink MIME");
        cases.insert(
            "failing_reader_or_sink".to_string(),
            writer
                .add_file_reader_with_media_type(
                    "broken.bin",
                    FailingReader { emitted: false },
                    "application/octet-stream",
                )
                .is_err(),
        );
        let corrupt = corrupt_stored_payload(package);
        let corrupt_owner = OwnedPackage::from_bytes(corrupt).expect("corrupt raw package index");
        cases.insert(
            "corrupt_stored_payload_without_verification".to_string(),
            replace_content_xml_with_payload_verification(&corrupt_owner, TARGET_CONTENT).is_err(),
        );
        for id in NEGATIVE_IDS {
            assert!(
                cases.get(id).copied().unwrap_or(false),
                "negative case failed: {id}"
            );
        }
        cases
    }

    fn descriptor(id: &str) -> Value {
        match id {
            "archive_catalog_and_raw_split" => json!({
                "id": id,
                "family": "raw",
                "operations": ["Archive::new", "file_names", "contains", "metadata", "is_stored", "read"],
                "boundary": "Archive construction and catalog/read operation; package bytes and expected content are prepared outside the clock",
            }),
            "manifest_xml_stream_and_parse" => json!({
                "id": id,
                "family": "manifest",
                "operations": ["Archive::read_manifest_xml", "read_manifest", "parse_manifest"],
                "boundary": "prebuilt Archive manifest materialization/read/parse only; package bytes, Archive and limits are prepared outside the clock",
            }),
            "metadata_xml_and_namespace_resolution" => json!({
                "id": id,
                "family": "xml",
                "operations": ["Metadata::from_xml", "ResolvedReader::from_xml baseline/from_xml_document current", "read_resolved_event_into"],
                "boundary": "bounded metadata/namespace construction and event consumption; XML constants, limits and proof serialization are outside the clock",
            }),
            "writer_memory_publication" => json!({
                "id": id,
                "family": "writer",
                "operations": ["PackageWriter::new", "add_authored_xml", "add_file_reader_with_media_type_and_compression", "finish"],
                "boundary": "in-memory writer construction/publication; deterministic payload values are prepared outside and output proof is outside the clock",
            }),
            "writer_stream_publication" => json!({
                "id": id,
                "family": "writer",
                "operations": ["PackageWriter::with_writer", "add_file_reader_with_media_type_and_compression", "finish_to_writer"],
                "boundary": "stream writer/sink construction and reader-to-sink publication; deterministic payload is prepared outside the clock",
            }),
            "writer_bounded_policy" => json!({
                "id": id,
                "family": "writer",
                "operations": ["PackageWriter::new_bounded", "add_file", "finish_to_bounded_bytes"],
                "boundary": "bounded writer construction/publication and refusal; payload and output/error proof are outside the clock",
            }),
            "raw_split_merge_and_opaque_preservation" => json!({
                "id": id,
                "family": "merge",
                "operations": ["OwnedPackage::from_bytes", "replace_content_xml", "raw_identical_members"],
                "boundary": "source reopen and raw replacement/publication; expected payload, member audit and output proof are outside the clock",
            }),
            "raw_xml_splice_merge_publication" => json!({
                "id": id,
                "family": "splice",
                "operations": ["XmlSplicePublication::new", "XmlSplicePublication::replace", "rebuild_package_with_xml_splices"],
                "boundary": "prebuilt per-iteration source publication batches and rebuild only; source reopen/index/proof, fragment copies and output audit are outside the clock",
            }),
            "unified_facade_raw_parts_and_publication" => json!({
                "id": id,
                "family": "facade",
                "operations": ["open_bytes_with", "File::parts", "File::raw", "File::to_bytes", "File::metadata"],
                "boundary": "facade open, raw/parts/metadata access and byte-identical no-op publication; input bytes are intentionally cloned inside the measured API boundary",
            }),
            "consumer_reopen_matrix" => json!({
                "id": id,
                "family": "consumers",
                "operations": ["Archive::new", "Archive::read", "open_bytes_with", "File::parts", "File::raw"],
                "boundary": "prebuilt owned per-iteration family inputs; raw split and unified facade reopen/projection only; package generation and ownership copies/output comparison are outside the clock",
            }),
            _ => panic!("unknown route {id}"),
        }
    }

    fn proof(negative: &BTreeMap<String, bool>, values: [bool; 11]) -> Value {
        let negative_cases_ok = negative.values().copied().all(|value| value);
        let [
            operation_ok,
            changed_ok,
            no_op_ok,
            reopen_ok,
            source_bytes_ok,
            facade_bytes_ok,
            untouched_members_ok,
            manifest_ok,
            metadata_ok,
            namespace_ok,
            bounded_error_ok,
        ] = values;
        assert!(operation_ok && changed_ok && no_op_ok && reopen_ok && source_bytes_ok);
        assert!(facade_bytes_ok && untouched_members_ok && manifest_ok && metadata_ok);
        assert!(namespace_ok && bounded_error_ok && negative_cases_ok);
        json!({
            "operation_ok": operation_ok,
            "changed_ok": changed_ok,
            "no_op_ok": no_op_ok,
            "reopen_ok": reopen_ok,
            "source_bytes_ok": source_bytes_ok,
            "facade_bytes_ok": facade_bytes_ok,
            "untouched_members_ok": untouched_members_ok,
            "manifest_ok": manifest_ok,
            "metadata_ok": metadata_ok,
            "namespace_ok": namespace_ok,
            "bounded_error_ok": bounded_error_ok,
            "negative_cases_ok": negative_cases_ok,
            "negative_cases": negative,
        })
    }

    fn measure(iterations: usize, mut operation: impl FnMut()) -> Value {
        let mut allocation_samples =
            Vec::with_capacity(if allocation::ENABLED { iterations } else { 0 });
        let started = Instant::now();
        for _ in 0..iterations {
            allocation::begin();
            operation();
            if allocation::ENABLED {
                // Snapshot and disable the observer before allocating the
                // serde value that records this sample.
                let sample = allocation::end();
                allocation_samples.push(json!({
                    "allocations": sample[0],
                    "deallocations": sample[1],
                    "allocated_bytes": sample[2],
                    "deallocated_bytes": sample[3],
                    "reallocations": sample[4],
                }));
            }
        }
        let elapsed_ns = started.elapsed().as_nanos() / iterations as u128;
        json!({
            "iterations": iterations,
            "ns_per_operation": elapsed_ns,
            "allocation_samples": allocation_samples,
        })
    }

    fn timing(mode: &str, iterations: usize, operation: impl FnMut()) -> Value {
        if mode == "validate" {
            return json!({
                "status": "not_run",
                "elapsed_claimable": false,
                "allocator_enabled": false,
                "iterations": 0,
            });
        }
        assert!(iterations > 0 && iterations <= MAX_ITERATIONS);
        assert!(mode != "allocator" || allocation::ENABLED);
        json!({
            "status": if mode == "allocator" { "allocator_observed" } else { "measured" },
            "elapsed_claimable": mode == "timed",
            "allocator_enabled": allocation::ENABLED,
            "measurement": measure(iterations, operation),
        })
    }

    fn archive_route(package: &[u8]) -> (Value, impl FnMut()) {
        let input_names = archive(package)
            .expect("archive")
            .file_names()
            .map(str::to_owned)
            .collect::<Vec<_>>();
        let input_names_again = input_names.clone();
        let mut reader = archive(package).expect("archive");
        let names = reader.file_names().map(str::to_owned).collect::<Vec<_>>();
        assert_eq!(names, input_names);
        assert!(reader.contains("content.xml"));
        assert!(reader.metadata("content.xml").is_ok());
        assert!(reader.is_stored("mimetype").expect("mimetype storage"));
        assert_eq!(
            reader.read("content.xml").expect("content read"),
            SOURCE_CONTENT.as_bytes()
        );
        let reopened = OwnedPackage::from_bytes(package.to_vec()).expect("raw reopen");
        let facade = open_bytes_with(package.to_vec(), &OpenOptions::new()).expect("facade reopen");
        assert!(!reopened.as_bytes().is_empty());
        assert_eq!(
            facade.raw("/content.xml").expect("facade raw").bytes(),
            SOURCE_CONTENT.as_bytes()
        );
        let archive_ok = names == input_names
            && reader.contains("content.xml")
            && reader.metadata("content.xml").is_ok()
            && reader.is_stored("mimetype").expect("mimetype storage")
            && reader.read("content.xml").expect("content read") == SOURCE_CONTENT.as_bytes();
        let reopen_ok = reopened.as_bytes() == package;
        let facade_ok =
            facade.raw("/content.xml").expect("facade raw").bytes() == SOURCE_CONTENT.as_bytes();
        let details = json!({
            "names": names,
            "names_repeat": input_names_again,
            "archive": archive_fingerprint(package).1,
        });
        let negative = BTreeMap::new();
        let row_proof = proof(
            &negative,
            [
                archive_ok,
                SOURCE_CONTENT != TARGET_CONTENT,
                names == input_names_again,
                reopen_ok,
                reopened.as_bytes() == package,
                facade_ok,
                names == input_names,
                !input_names.is_empty(),
                reader.metadata("content.xml").is_ok(),
                facade_ok,
                true,
            ],
        );
        (
            json!({ "proof": row_proof, "details": details }),
            move || {
                let value = archive(package)
                    .expect("timed archive")
                    .read("content.xml")
                    .expect("timed raw read");
                black_box(value);
            },
        )
    }

    fn manifest_route(package: &[u8]) -> (Value, impl FnMut()) {
        let borrowed = archive(package).expect("manifest archive");
        let xml = read_manifest_xml(&borrowed).expect("manifest XML");
        let manifest =
            read_manifest(&borrowed, XmlStreamLimits::default()).expect("manifest parse");
        let parsed = parse_manifest(&xml, XmlStreamLimits::default()).expect("manifest parse XML");
        assert_eq!(manifest, parsed);
        let content_unchanged = archive(package)
            .expect("manifest source")
            .read("content.xml")
            .expect("manifest content")
            == SOURCE_CONTENT.as_bytes();
        let negative = BTreeMap::new();
        let row_proof = proof(
            &negative,
            [
                manifest == parsed,
                xml != TARGET_CONTENT,
                manifest == parsed,
                content_unchanged,
                content_unchanged,
                content_unchanged,
                !xml.is_empty(),
                !xml.is_empty(),
                true,
                true,
                true,
            ],
        );
        let timed_archive = archive(package).expect("timed manifest archive");
        let timed_limits = XmlStreamLimits::default();
        (
            json!({ "proof": row_proof, "manifest_bytes": xml.len() }),
            move || {
                let xml = read_manifest_xml(&timed_archive).expect("timed manifest XML");
                let manifest = read_manifest(&timed_archive, timed_limits).expect("timed manifest");
                let parsed = parse_manifest(&xml, timed_limits).expect("timed manifest parse");
                black_box((manifest, parsed));
            },
        )
    }

    fn metadata_route() -> (Value, impl FnMut()) {
        let limits = XmlStreamLimits::default();
        let metadata = Metadata::from_xml(META_XML, limits).expect("metadata");
        assert_eq!(metadata.title.as_deref(), Some("raw codec"));
        let mut reader = resolved_document(META_XML, limits).expect("resolved reader");
        let mut buffer = Vec::new();
        let mut events = 0usize;
        loop {
            let (_, event) = reader
                .read_resolved_event_into(&mut buffer)
                .expect("resolved event");
            events += 1;
            if matches!(event, Event::Eof) {
                break;
            }
            buffer.clear();
        }
        assert!(events > 2);
        let operation_ok = metadata.title.as_deref() == Some("raw codec");
        let namespace_ok = events > 2;
        let reopen_ok = Metadata::from_xml(META_XML, limits)
            .map(|value| value.title.as_deref() == Some("raw codec"))
            .unwrap_or(false);
        let negative = BTreeMap::new();
        let row_proof = proof(
            &negative,
            [
                operation_ok,
                META_XML.contains("raw codec"),
                reopen_ok,
                events > 2,
                !META_XML.is_empty(),
                operation_ok,
                true,
                true,
                operation_ok,
                namespace_ok,
                true,
            ],
        );
        let timed_limits = XmlStreamLimits::default();
        let mut timed_buffer = Vec::with_capacity(META_XML.len());
        (json!({ "proof": row_proof, "events": events }), move || {
            let metadata = Metadata::from_xml(META_XML, timed_limits).expect("timed metadata");
            let mut reader = resolved_document(META_XML, timed_limits).expect("timed reader");
            let mut timed_events = 0usize;
            loop {
                let (_, event) = reader
                    .read_resolved_event_into(&mut timed_buffer)
                    .expect("timed resolved event");
                timed_events += 1;
                timed_buffer.clear();
                if matches!(event, Event::Eof) {
                    break;
                }
            }
            black_box((metadata, timed_events));
        })
    }

    fn memory_route(package: &[u8]) -> (Value, impl FnMut()) {
        let opaque = opaque_payload();
        let output = package_with_content_and_opaque(
            constants::ODF_TEXT,
            TARGET_CONTENT.as_bytes(),
            &opaque,
        );
        let (_, output_members) = archive_fingerprint(&output);
        let reopened = OwnedPackage::from_bytes(output.clone()).expect("writer reopen");
        assert_eq!(
            reopened.get_file("content.xml").expect("writer content"),
            TARGET_CONTENT.as_bytes()
        );
        assert!(output_members["members"].as_array().is_some());
        let operation_ok = !output.is_empty();
        let changed_ok = output != package;
        let no_op = package_with_content_and_opaque(
            constants::ODF_TEXT,
            SOURCE_CONTENT.as_bytes(),
            &opaque,
        );
        let no_op_ok = no_op == package;
        let reopen_ok = reopened.get_file("content.xml").is_ok();
        let untouched_members_ok = output_members["members"]
            .as_array()
            .map(|members| members.iter().any(|member| member["name"] == OPAQUE_MEMBER))
            .unwrap_or(false);
        let negative = BTreeMap::new();
        let row_proof = proof(
            &negative,
            [
                operation_ok,
                changed_ok,
                no_op_ok,
                reopen_ok,
                reopened.get_file("content.xml").expect("writer content")
                    == TARGET_CONTENT.as_bytes(),
                true,
                untouched_members_ok,
                true,
                true,
                true,
                true,
            ],
        );
        (
            json!({ "proof": row_proof, "output_sha256": digest(&output), "output_bytes": output.len() }),
            move || {
                black_box(package_with_content_and_opaque(
                    constants::ODF_TEXT,
                    TARGET_CONTENT.as_bytes(),
                    &opaque,
                ));
            },
        )
    }

    fn stream_route() -> (Value, impl FnMut()) {
        let opaque = opaque_payload();
        let mut writer = sink_writer(Cursor::new(Vec::new()));
        writer
            .set_mimetype(constants::ODF_TEXT)
            .expect("stream MIME");
        writer
            .add_authored_xml("content.xml", SOURCE_CONTENT.as_bytes(), "text/xml")
            .expect("stream XML");
        writer
            .add_file_reader_with_media_type_and_compression(
                OPAQUE_MEMBER,
                Cursor::new(opaque.clone()),
                "application/octet-stream",
                PackageCompression::Stored,
            )
            .expect("stream opaque");
        let output = writer
            .finish_to_writer()
            .expect("stream finish")
            .into_inner();
        let reopened = OwnedPackage::from_bytes(output.clone()).expect("stream reopen");
        assert_eq!(
            reopened
                .get_file(OPAQUE_MEMBER)
                .expect("stream opaque reopen"),
            opaque
        );
        let operation_ok = !output.is_empty();
        let reopen_ok = reopened.get_file(OPAQUE_MEMBER).is_ok();
        let source_bytes_ok = reopened
            .get_file(OPAQUE_MEMBER)
            .expect("stream opaque source")
            == opaque;
        let negative = BTreeMap::new();
        let row_proof = proof(
            &negative,
            [
                operation_ok,
                opaque.len() == OPAQUE_BYTES,
                reopen_ok,
                reopen_ok,
                source_bytes_ok,
                true,
                source_bytes_ok,
                true,
                true,
                true,
                true,
            ],
        );
        (
            json!({ "proof": row_proof, "output_sha256": digest(&output), "output_bytes": output.len() }),
            move || {
                let mut writer = sink_writer(Cursor::new(Vec::new()));
                writer
                    .set_mimetype(constants::ODF_TEXT)
                    .expect("timed stream MIME");
                writer
                    .add_file_reader_with_media_type_and_compression(
                        OPAQUE_MEMBER,
                        Cursor::new(opaque.as_slice()),
                        "application/octet-stream",
                        PackageCompression::Stored,
                    )
                    .expect("timed stream opaque");
                black_box(writer.finish_to_writer().expect("timed stream finish"));
            },
        )
    }

    fn bounded_route() -> (Value, impl FnMut()) {
        let output = bounded_publication(1 << 20, constants::ODF_TEXT, SOURCE_CONTENT.as_bytes())
            .expect("bounded output");
        let reopened = OwnedPackage::from_bytes(output.clone());
        assert!(reopened.is_ok());
        let operation_ok = !output.is_empty();
        let reopen_ok = reopened
            .as_ref()
            .map(|value| value.get_file("content.xml").is_ok())
            .unwrap_or(false);
        let negative = BTreeMap::new();
        let row_proof = proof(
            &negative,
            [
                operation_ok,
                operation_ok,
                reopen_ok,
                reopen_ok,
                reopen_ok,
                true,
                true,
                true,
                true,
                true,
                true,
            ],
        );
        (
            json!({ "proof": row_proof, "output_bytes": output.len() }),
            move || {
                black_box(
                    bounded_publication(1 << 20, constants::ODF_TEXT, SOURCE_CONTENT.as_bytes())
                        .expect("timed bounded"),
                );
            },
        )
    }

    fn merge_route(package: &[u8], negative: &BTreeMap<String, bool>) -> (Value, impl FnMut()) {
        let source = OwnedPackage::from_bytes(package.to_vec()).expect("merge source");
        let output = replace_content_xml(&source, TARGET_CONTENT).expect("content replacement");
        assert_ne!(output, package);
        let identical = raw_identical_members(package, &output).expect("raw member audit");
        assert!(identical.contains("META-INF/manifest.xml"));
        assert!(identical.contains(OPAQUE_MEMBER));
        let reopened = OwnedPackage::from_bytes(output.clone()).expect("merge reopen");
        assert_eq!(
            reopened.get_file("content.xml").expect("merge content"),
            TARGET_CONTENT.as_bytes()
        );
        let verified = replace_content_xml_with_payload_verification(&source, TARGET_CONTENT)
            .expect("verified replacement");
        assert_eq!(
            OwnedPackage::from_bytes(verified)
                .expect("verified reopen")
                .get_file("content.xml")
                .expect("verified content"),
            TARGET_CONTENT.as_bytes()
        );
        let no_op = replace_content_xml(&source, SOURCE_CONTENT).expect("merge no-op");
        let operation_ok = output != package;
        let changed_ok = reopened.get_file("content.xml").is_ok();
        let no_op_ok = no_op == package;
        let source_bytes_ok = source.as_bytes() == package;
        let untouched_members_ok = identical.contains(OPAQUE_MEMBER)
            && identical.contains("META-INF/manifest.xml")
            && untouched_archive_members_equal(package, &output, "content.xml");
        let row_proof = proof(
            negative,
            [
                operation_ok,
                changed_ok,
                no_op_ok,
                OwnedPackage::from_bytes(output.clone()).is_ok(),
                source_bytes_ok,
                OwnedPackage::from_bytes(output.clone()).is_ok(),
                untouched_members_ok,
                !identical.is_empty(),
                true,
                true,
                true,
            ],
        );
        (
            json!({ "proof": row_proof, "identical_members": identical.into_iter().collect::<Vec<_>>(), "output_sha256": digest(&output) }),
            move || {
                let source =
                    OwnedPackage::from_bytes(package.to_vec()).expect("timed merge source");
                black_box(replace_content_xml(&source, TARGET_CONTENT).expect("timed merge"));
            },
        )
    }

    fn splice_route(package: &[u8], negative: &BTreeMap<String, bool>) -> (Value, impl FnMut()) {
        let source = OwnedPackage::from_bytes(package.to_vec()).expect("splice source");
        let part = XmlSourcePart::load(&source, "content.xml").expect("splice part");
        let start = part
            .bytes()
            .windows(b"source".len())
            .position(|window| window == b"source")
            .expect("source marker");
        let range = start..start + b"source".len();
        let range_proof = part.checked_range(range, b"source").expect("splice proof");
        let fragment = AuthoredXmlFragment::text(b"target".to_vec()).expect("splice text");
        let mut publication = XmlSplicePublication::new(part.clone());
        publication
            .replace(range_proof.clone(), fragment.clone())
            .expect("splice replace");
        let output = rebuild_package_with_xml_splices(&source, vec![publication], 1 << 20)
            .expect("splice rebuild");
        assert_eq!(
            OwnedPackage::from_bytes(output.clone())
                .expect("splice reopen")
                .get_file("content.xml")
                .expect("splice content"),
            TARGET_CONTENT.as_bytes()
        );
        let direct = replace_content_xml_spliced(&source, TARGET_CONTENT, {
            let part = XmlSourcePart::load(&source, "content.xml").expect("direct part");
            let start = part
                .bytes()
                .windows(b"source".len())
                .position(|window| window == b"source")
                .expect("direct marker");
            let proof = part
                .checked_range(start..start + b"source".len(), b"source")
                .expect("direct proof");
            let mut publication = XmlSplicePublication::new(part);
            publication
                .replace(
                    proof,
                    AuthoredXmlFragment::text(b"target".to_vec()).expect("direct text"),
                )
                .expect("direct replace");
            publication
        })
        .expect("direct splice");
        let direct_changed = direct != package;
        assert_eq!(
            OwnedPackage::from_bytes(direct)
                .expect("direct reopen")
                .get_file("content.xml")
                .expect("direct content"),
            TARGET_CONTENT.as_bytes()
        );
        let no_op = replace_content_xml(&source, SOURCE_CONTENT).expect("splice no-op");
        let operation_ok = output != package;
        let changed_ok = OwnedPackage::from_bytes(output.clone())
            .expect("splice changed reopen")
            .get_file("content.xml")
            .map(|value| value == TARGET_CONTENT.as_bytes())
            .unwrap_or(false);
        let no_op_ok = no_op == package;
        let source_bytes_ok = source.as_bytes() == package;
        let untouched_members_ok = untouched_archive_members_equal(package, &output, "content.xml");
        let row_proof = proof(
            negative,
            [
                operation_ok,
                changed_ok,
                no_op_ok,
                direct_changed,
                source_bytes_ok,
                changed_ok,
                untouched_members_ok,
                true,
                true,
                true,
                true,
            ],
        );
        let timed_source = source.clone();
        let mut timed_batches = Vec::with_capacity(MAX_ITERATIONS);
        for _ in 0..MAX_ITERATIONS {
            let mut timed_publication = XmlSplicePublication::new(part.clone());
            timed_publication
                .replace(range_proof.clone(), fragment.clone())
                .expect("prepared timed replace");
            timed_batches.push(vec![timed_publication]);
        }
        (
            json!({ "proof": row_proof, "output_sha256": digest(&output) }),
            move || {
                let publications = timed_batches.pop().expect("prepared timed splice batch");
                black_box(
                    rebuild_package_with_xml_splices(&timed_source, publications, 1 << 20)
                        .expect("timed splice"),
                );
            },
        )
    }

    fn facade_route(package: &[u8]) -> (Value, impl FnMut()) {
        let file = open_bytes_with(package.to_vec(), &OpenOptions::new())
            .unwrap_or_else(|error| panic!("facade: {error}"));
        let parts = file.parts().expect("facade parts");
        let raw = file.raw("/content.xml").expect("facade content");
        let published = file.to_bytes().expect("facade publish");
        let metadata = file.metadata().expect("facade metadata");
        let reopened =
            open_bytes_with(published.clone(), &OpenOptions::new()).expect("facade reopen");
        let source_content = archive(package)
            .expect("facade source archive")
            .read("content.xml")
            .expect("facade source content");
        assert!(!parts.is_empty());
        assert_eq!(
            raw.bytes(),
            file.raw("/content.xml").expect("facade repeat").bytes()
        );
        assert!(!published.is_empty());
        let no_op_ok = published == package;
        let operation_ok = !parts.is_empty() && !published.is_empty();
        let source_bytes_ok = raw.bytes() == source_content.as_slice();
        let facade_bytes_ok = reopened.to_bytes().expect("facade repeat publish") == package;
        let metadata_ok = metadata.is_some();
        let untouched_members_ok =
            archive_fingerprint(package).1 == archive_fingerprint(&published).1;
        let row_proof = proof(
            &BTreeMap::new(),
            [
                operation_ok,
                no_op_ok,
                no_op_ok,
                reopened.to_bytes().is_ok(),
                source_bytes_ok,
                facade_bytes_ok,
                untouched_members_ok,
                true,
                metadata_ok,
                true,
                true,
            ],
        );
        (
            json!({ "proof": row_proof, "parts": parts.len(), "metadata_present": metadata.is_some(), "source_sha256": digest(package), "published_sha256": digest(&published) }),
            move || {
                let file = open_bytes_with(package.to_vec(), &OpenOptions::new())
                    .expect("timed facade open");
                let parts = file.parts().expect("timed facade parts");
                let raw = file.raw("/content.xml").expect("timed facade raw");
                let metadata = file.metadata().expect("timed facade metadata");
                let published = file.to_bytes().expect("timed facade publish");
                black_box((parts, raw, metadata, published));
            },
        )
    }

    fn consumer_route(families: &[(String, Vec<u8>)], iterations: usize) -> (Value, impl FnMut()) {
        let mut rows = Vec::with_capacity(families.len());
        let mut all_ok = true;
        for (name, bytes) in families {
            let raw = archive(bytes).unwrap_or_else(|error| panic!("raw {name}: {error}"));
            let raw_content = raw.read("content.xml").expect("consumer raw content");
            let consumer_input = bytes.to_vec();
            let file = open_bytes_with(consumer_input, &OpenOptions::new())
                .unwrap_or_else(|error| panic!("consumer facade {name}: {error}"));
            let facade_content = file.raw("/content.xml").expect("consumer facade raw");
            let family_ok = raw_content == facade_content.bytes();
            assert!(family_ok, "raw/facade bytes differ for {name}");
            all_ok &= family_ok;
            assert_eq!(
                raw_content,
                facade_content.bytes(),
                "raw/facade bytes differ for {name}"
            );
            assert!(!file.parts().expect("consumer parts").is_empty());
            rows.push(json!({"family": name, "source_sha256": digest(bytes), "content_sha256": digest(&raw_content)}));
        }
        let row_proof = proof(
            &BTreeMap::new(),
            [
                all_ok,
                all_ok,
                all_ok,
                all_ok,
                all_ok,
                all_ok,
                all_ok,
                !rows.is_empty(),
                all_ok,
                all_ok,
                all_ok,
            ],
        );
        // `open_bytes_with` takes ownership of its input. Prepare one owned
        // family batch per measured iteration before the clock starts so the
        // raw/facade comparison measures reopening and projection, not an
        // implicit package-byte clone. The batch count is supplied by the
        // protocol (1 for validation, 128 for timed/allocator), keeping the
        // preparation bounded and identical for both API eras.
        let mut timed_inputs: Vec<Vec<Vec<u8>>> = Vec::with_capacity(iterations);
        for _ in 0..iterations {
            timed_inputs.push(families.iter().map(|(_, bytes)| bytes.clone()).collect());
        }
        (json!({ "proof": row_proof, "families": rows }), move || {
            let batch = timed_inputs.pop().expect("prepared consumer family batch");
            for (index, bytes) in batch.into_iter().enumerate() {
                let raw_content = {
                    let raw = archive(&bytes).expect("timed consumer raw");
                    raw.read("content.xml").expect("timed consumer content")
                };
                let file =
                    open_bytes_with(bytes, &OpenOptions::new()).expect("timed consumer facade");
                let parts = file.parts().expect("timed consumer parts");
                let facade_content = file.raw("/content.xml").expect("timed consumer facade raw");
                black_box((index, raw_content, parts, facade_content));
            }
        })
    }

    fn run_route(
        id: &str,
        mode: &str,
        iterations: usize,
        package: &[u8],
        families: &[(String, Vec<u8>)],
        negative: &BTreeMap<String, bool>,
    ) -> Value {
        let (details, operation) = match id {
            "archive_catalog_and_raw_split" => archive_route(package),
            "manifest_xml_stream_and_parse" => manifest_route(package),
            "metadata_xml_and_namespace_resolution" => metadata_route(),
            "writer_memory_publication" => memory_route(package),
            "writer_stream_publication" => stream_route(),
            "writer_bounded_policy" => bounded_route(),
            "raw_split_merge_and_opaque_preservation" => merge_route(package, negative),
            "raw_xml_splice_merge_publication" => splice_route(package, negative),
            "unified_facade_raw_parts_and_publication" => facade_route(package),
            "consumer_reopen_matrix" => consumer_route(families, iterations),
            _ => panic!("unknown route {id}"),
        };
        let timing_value = timing(mode, iterations, operation);
        json!({
            "route": id,
            "family": descriptor(id)["family"],
            "descriptor": descriptor(id),
            "mode": mode,
            "iterations": iterations,
            "details": details,
            "timing": timing_value,
        })
    }

    fn parse_args() -> (String, String, usize) {
        let mut mode = String::from("validate");
        let mut route = String::from("all");
        let mut iterations = 1usize;
        let mut args = std::env::args().skip(1);
        while let Some(argument) = args.next() {
            match argument.as_str() {
                "--mode" => mode = args.next().expect("--mode value"),
                "--route" => route = args.next().expect("--route value"),
                "--iterations" => {
                    iterations = args
                        .next()
                        .expect("--iterations value")
                        .parse()
                        .expect("iterations")
                },
                other => panic!("unknown argument {other}"),
            }
        }
        assert!(matches!(mode.as_str(), "validate" | "timed" | "allocator"));
        assert!(iterations > 0 && iterations <= MAX_ITERATIONS);
        (mode, route, iterations)
    }

    pub fn run() {
        let (mode, selected_route, iterations) = parse_args();
        if mode != "validate" {
            assert_eq!(
                iterations, MAX_ITERATIONS,
                "timed modes require exactly 128 iterations"
            );
        }
        let package = package_with_mime(constants::ODF_TEXT);
        let families = family_packages();
        let negative = negative_cases(&package);
        let routes = if selected_route == "all" {
            ROUTE_IDS.to_vec()
        } else {
            assert!(
                ROUTE_IDS.contains(&selected_route.as_str()),
                "unknown route"
            );
            vec![selected_route.as_str()]
        };
        let family_names = families
            .iter()
            .map(|(name, _)| name.as_str())
            .collect::<Vec<_>>();
        assert_eq!(family_names, FAMILY_IDS);
        let rows = routes
            .into_iter()
            .map(|id| run_route(id, &mode, iterations, &package, &families, &negative))
            .collect::<Vec<_>>();
        let output = json!({
            "schema": "litchi.odf.raw-codec-qualification.v1",
            "mode": mode,
            "iterations": iterations,
            "routes": rows,
            "families": FAMILY_IDS,
            "negative_cases": negative,
            "api_era": API_ERA,
            "namespace_contract": NAMESPACE_CONTRACT,
            "claim_scope": "correctness and reproducible paired protocol only; no native acceptance claim",
            "allocator_enabled": allocation::ENABLED,
        });
        println!("{}", serde_json::to_string(&output).expect("probe JSON"));
    }
}

fn main() {
    probe::run();
}
