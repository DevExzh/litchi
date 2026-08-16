#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design"
)]

use std::collections::BTreeMap;
use std::io::{self, Write};
use std::num::{NonZeroU64, NonZeroUsize};
use std::sync::Arc;

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits, OwnedSource,
};
use litchi_opc::{
    OpcError, OpcPackage, PackURI, ReadLimits, SourceBackedPackage, SourceTopologyPlan,
};

const CONTENT_TYPES_NS: &str = "http://schemas.openxmlformats.org/package/2006/content-types";
const RELATIONSHIPS_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const OFFICE_DOCUMENT_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const CUSTOM_REL: &str = "http://example.invalid/relationships/custom";

#[derive(Clone, Copy)]
struct Entry<'a> {
    name: &'a [u8],
    data: &'a [u8],
    local_extra: &'a [u8],
    central_extra: &'a [u8],
    comment: &'a [u8],
}

#[derive(Debug, Clone)]
struct RawRecord {
    local: Vec<u8>,
    central: Vec<u8>,
}

fn pack(uri: &str) -> PackURI {
    PackURI::new(uri).unwrap()
}

fn push_u16(output: &mut Vec<u8>, value: u16) {
    output.extend_from_slice(&value.to_le_bytes());
}

fn push_u32(output: &mut Vec<u8>, value: u32) {
    output.extend_from_slice(&value.to_le_bytes());
}

fn read_u16(bytes: &[u8], offset: usize) -> u16 {
    u16::from_le_bytes([bytes[offset], bytes[offset + 1]])
}

fn read_u32(bytes: &[u8], offset: usize) -> u32 {
    u32::from_le_bytes([
        bytes[offset],
        bytes[offset + 1],
        bytes[offset + 2],
        bytes[offset + 3],
    ])
}

fn stored_archive(entries: &[Entry<'_>], archive_comment: &[u8]) -> Vec<u8> {
    let mut output = Vec::new();
    let mut central = Vec::new();
    for entry in entries {
        let local_offset = u32::try_from(output.len()).unwrap();
        let size = u32::try_from(entry.data.len()).unwrap();
        let crc = soapberry_zip::crc32(entry.data);
        push_u32(&mut output, 0x0403_4b50);
        push_u16(&mut output, 20);
        push_u16(&mut output, 0);
        push_u16(&mut output, 0);
        push_u16(&mut output, 0);
        push_u16(&mut output, 0);
        push_u32(&mut output, crc);
        push_u32(&mut output, size);
        push_u32(&mut output, size);
        push_u16(&mut output, u16::try_from(entry.name.len()).unwrap());
        push_u16(&mut output, u16::try_from(entry.local_extra.len()).unwrap());
        output.extend_from_slice(entry.name);
        output.extend_from_slice(entry.local_extra);
        output.extend_from_slice(entry.data);

        push_u32(&mut central, 0x0201_4b50);
        push_u16(&mut central, 20);
        push_u16(&mut central, 20);
        push_u16(&mut central, 0);
        push_u16(&mut central, 0);
        push_u16(&mut central, 0);
        push_u16(&mut central, 0);
        push_u32(&mut central, crc);
        push_u32(&mut central, size);
        push_u32(&mut central, size);
        push_u16(&mut central, u16::try_from(entry.name.len()).unwrap());
        push_u16(
            &mut central,
            u16::try_from(entry.central_extra.len()).unwrap(),
        );
        push_u16(&mut central, u16::try_from(entry.comment.len()).unwrap());
        push_u16(&mut central, 0);
        push_u16(&mut central, 0);
        push_u32(&mut central, 0);
        push_u32(&mut central, local_offset);
        central.extend_from_slice(entry.name);
        central.extend_from_slice(entry.central_extra);
        central.extend_from_slice(entry.comment);
    }
    let central_offset = u32::try_from(output.len()).unwrap();
    let central_size = u32::try_from(central.len()).unwrap();
    output.extend_from_slice(&central);
    push_u32(&mut output, 0x0605_4b50);
    push_u16(&mut output, 0);
    push_u16(&mut output, 0);
    let count = u16::try_from(entries.len()).unwrap();
    push_u16(&mut output, count);
    push_u16(&mut output, count);
    push_u32(&mut output, central_size);
    push_u32(&mut output, central_offset);
    push_u16(&mut output, u16::try_from(archive_comment.len()).unwrap());
    output.extend_from_slice(archive_comment);
    output
}

fn canonical_root_relationships() -> Vec<u8> {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="{RELATIONSHIPS_NS}"><Relationship Id="rId1" Type="{OFFICE_DOCUMENT_REL}" Target="word/document.xml"/></Relationships>"#
    )
    .into_bytes()
}

fn canonical_content_types() -> Vec<u8> {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="{CONTENT_TYPES_NS}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>"#
    )
    .into_bytes()
}

fn noncanonical_content_types() -> Vec<u8> {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="{CONTENT_TYPES_NS}">
  <Default ContentType="application/xml" Extension="xml" />
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" />
</Types>
"#
    )
    .into_bytes()
}

fn source_bytes(
    noncanonical_manifest: bool,
    with_document_rels: bool,
    noncanonical_document_rels: bool,
) -> Vec<u8> {
    let content_types = if noncanonical_manifest {
        noncanonical_content_types()
    } else {
        canonical_content_types()
    };
    let root_relationships = canonical_root_relationships();
    let document_relationships = if noncanonical_document_rels {
        format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?> <Relationships xmlns="{RELATIONSHIPS_NS}"><Relationship Id="rId1" Type="{CUSTOM_REL}" Target="../custom/existing.xml"/></Relationships>"#
        )
    } else {
        format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="{RELATIONSHIPS_NS}"><Relationship Id="rId1" Type="{CUSTOM_REL}" Target="../custom/existing.xml"/></Relationships>"#
        )
    }
    .into_bytes();
    let mut entries = vec![
        Entry {
            name: b"[Content_Types].xml",
            data: &content_types,
            local_extra: b"\x99\x99\x04\x00meta",
            central_extra: b"\x88\x88\x02\x00ce",
            comment: b"manifest-comment",
        },
        Entry {
            name: b"_rels/.rels",
            data: &root_relationships,
            local_extra: b"\x99\x99\x04\x00root",
            central_extra: b"\x88\x88\x02\x00cr",
            comment: b"root-comment",
        },
        Entry {
            name: b"word/document.xml",
            data: b"<before/>",
            local_extra: b"\x99\x99\x04\x00part",
            central_extra: b"\x88\x88\x02\x00cp",
            comment: b"part-comment",
        },
        Entry {
            name: b"custom/existing.xml",
            data: b"<existing/>",
            local_extra: b"\x99\x99\x04\x00blob",
            central_extra: b"\x88\x88\x02\x00cb",
            comment: b"blob-comment",
        },
        Entry {
            name: b"custom/untouched.xml",
            data: b"<untouched/>",
            local_extra: b"\x99\x99\x04\x00keep",
            central_extra: b"\x88\x88\x02\x00ck",
            comment: b"keep-comment",
        },
    ];
    if with_document_rels {
        entries.push(Entry {
            name: b"word/_rels/document.xml.rels",
            data: &document_relationships,
            local_extra: b"\x99\x99\x04\x00rels",
            central_extra: b"\x88\x88\x02\x00cr",
            comment: b"rels-comment",
        });
    }
    stored_archive(&entries, b"archive-comment")
}

fn open(bytes: &[u8]) -> SourceBackedPackage {
    SourceBackedPackage::from_read_at(Arc::new(OwnedSource::new(bytes.to_vec()))).unwrap()
}

fn publish(bytes: &[u8], plan: SourceTopologyPlan) -> litchi_opc::Result<Vec<u8>> {
    let package = open(bytes);
    let mut output = Vec::new();
    package.write_topology_to_stream(&mut output, plan)?;
    Ok(output)
}

fn raw_records(bytes: &[u8]) -> BTreeMap<String, RawRecord> {
    let eocd = bytes
        .windows(4)
        .rposition(|window| window == 0x0605_4b50_u32.to_le_bytes())
        .unwrap();
    let count = read_u16(bytes, eocd + 10) as usize;
    let central_offset = read_u32(bytes, eocd + 16) as usize;
    let mut cursor = central_offset;
    let mut records = BTreeMap::new();
    for _ in 0..count {
        assert_eq!(&bytes[cursor..cursor + 4], &0x0201_4b50_u32.to_le_bytes());
        let name_len = read_u16(bytes, cursor + 28) as usize;
        let extra_len = read_u16(bytes, cursor + 30) as usize;
        let comment_len = read_u16(bytes, cursor + 32) as usize;
        let central_len = 46 + name_len + extra_len + comment_len;
        let name_start = cursor + 46;
        let name = String::from_utf8(bytes[name_start..name_start + name_len].to_vec()).unwrap();
        let local_offset = read_u32(bytes, cursor + 42) as usize;
        assert_eq!(
            &bytes[local_offset..local_offset + 4],
            &0x0403_4b50_u32.to_le_bytes()
        );
        let local_name_len = read_u16(bytes, local_offset + 26) as usize;
        let local_extra_len = read_u16(bytes, local_offset + 28) as usize;
        let compressed_len = read_u32(bytes, local_offset + 18) as usize;
        let local_len = 30 + local_name_len + local_extra_len + compressed_len;
        let mut central = bytes[cursor..cursor + central_len].to_vec();
        central[42..46].fill(0);
        records.insert(
            name,
            RawRecord {
                local: bytes[local_offset..local_offset + local_len].to_vec(),
                central,
            },
        );
        cursor += central_len;
    }
    records
}

fn zip_member(bytes: &[u8], wanted: &str) -> Vec<u8> {
    soapberry_zip::office::ArchiveReader::new(bytes)
        .unwrap()
        .read(wanted)
        .unwrap()
}

fn mark_entry_encrypted(mut bytes: Vec<u8>, wanted: &str) -> Vec<u8> {
    let eocd = bytes
        .windows(4)
        .rposition(|window| window == 0x0605_4b50_u32.to_le_bytes())
        .unwrap();
    let count = read_u16(&bytes, eocd + 10) as usize;
    let central_offset = read_u32(&bytes, eocd + 16) as usize;
    let mut cursor = central_offset;
    for _ in 0..count {
        assert_eq!(&bytes[cursor..cursor + 4], &0x0201_4b50_u32.to_le_bytes());
        let name_len = read_u16(&bytes, cursor + 28) as usize;
        let extra_len = read_u16(&bytes, cursor + 30) as usize;
        let comment_len = read_u16(&bytes, cursor + 32) as usize;
        let name_start = cursor + 46;
        let name = &bytes[name_start..name_start + name_len];
        let local_offset = read_u32(&bytes, cursor + 42) as usize;
        if name == wanted.as_bytes() {
            let central_flags = read_u16(&bytes, cursor + 8) | 1;
            bytes[cursor + 8..cursor + 10].copy_from_slice(&central_flags.to_le_bytes());
            let local_flags = read_u16(&bytes, local_offset + 6) | 1;
            bytes[local_offset + 6..local_offset + 8].copy_from_slice(&local_flags.to_le_bytes());
            return bytes;
        }
        cursor += 46 + name_len + extra_len + comment_len;
    }
    panic!("missing ZIP member {wanted}");
}

fn signed_source_bytes() -> Vec<u8> {
    let content_types = canonical_content_types();
    let root_relationships = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="{RELATIONSHIPS_NS}"><Relationship Id="rId1" Type="{OFFICE_DOCUMENT_REL}" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin" Target="signature/origin.xml"/></Relationships>"#
    )
    .into_bytes();
    let entries = [
        Entry {
            name: b"[Content_Types].xml",
            data: &content_types,
            local_extra: b"",
            central_extra: b"",
            comment: b"",
        },
        Entry {
            name: b"_rels/.rels",
            data: &root_relationships,
            local_extra: b"",
            central_extra: b"",
            comment: b"",
        },
        Entry {
            name: b"word/document.xml",
            data: b"<before/>",
            local_extra: b"",
            central_extra: b"",
            comment: b"",
        },
        Entry {
            name: b"signature/origin.xml",
            data: b"<origin/>",
            local_extra: b"",
            central_extra: b"",
            comment: b"",
        },
    ];
    stored_archive(&entries, b"")
}

fn malformed_xml_source_bytes() -> Vec<u8> {
    let content_types = canonical_content_types();
    let root_relationships = canonical_root_relationships();
    let entries = [
        Entry {
            name: b"[Content_Types].xml",
            data: &content_types,
            local_extra: b"",
            central_extra: b"",
            comment: b"",
        },
        Entry {
            name: b"_rels/.rels",
            data: &root_relationships,
            local_extra: b"",
            central_extra: b"",
            comment: b"",
        },
        Entry {
            name: b"word/document.xml",
            data: b"<broken",
            local_extra: b"",
            central_extra: b"",
            comment: b"",
        },
    ];
    stored_archive(&entries, b"")
}

fn managed_context() -> (CancellationSource, ExecutionContext) {
    let memory = 64 * 1024 * 1024;
    let budget = Budget::root(
        "opc-source-topology-integration",
        Limits::new(memory, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
    );
    let (cancellation_source, cancellation) = CancellationSource::pair();
    let execution_limits = ExecutionLimits::new(
        NonZeroUsize::new(1).unwrap(),
        NonZeroUsize::new(1).unwrap(),
        NonZeroU64::new(memory).unwrap(),
        0,
    )
    .unwrap();
    (
        cancellation_source,
        ExecutionContext::new(budget, cancellation, execution_limits),
    )
}

fn add_basic_topology_plan() -> SourceTopologyPlan {
    let mut plan = SourceTopologyPlan::new();
    plan.try_replace_part(pack("/word/document.xml"), b"<after/>".to_vec())
        .unwrap();
    plan.try_add_part(
        pack("/custom/new.xml"),
        "application/vnd.example.new+xml",
        b"<new/>".to_vec(),
    )
    .unwrap();
    plan.try_add_internal_relationship(
        pack("/"),
        "rId2",
        OFFICE_DOCUMENT_REL,
        pack("/custom/new.xml"),
    )
    .unwrap();
    plan.try_add_internal_relationship(
        pack("/word/document.xml"),
        "rId2",
        CUSTOM_REL,
        pack("/custom/new.xml"),
    )
    .unwrap();
    plan.try_add_internal_relationship(
        pack("/custom/new.xml"),
        "rIdChild",
        CUSTOM_REL,
        pack("/word/document.xml"),
    )
    .unwrap();
    plan
}

#[test]
fn empty_topology_plan_is_an_exact_source_copy() {
    let source = source_bytes(false, false, false);
    let output = publish(&source, SourceTopologyPlan::new()).unwrap();
    assert_eq!(output, source);
}

#[test]
fn topology_publish_replaces_adds_and_links_parts_atomically() {
    let source = source_bytes(false, false, false);
    let output = publish(&source, add_basic_topology_plan()).unwrap();
    let package = OpcPackage::from_bytes(&output).unwrap();

    assert_eq!(
        package
            .get_part(&pack("/word/document.xml"))
            .unwrap()
            .blob(),
        b"<after/>"
    );
    assert_eq!(
        package.get_part(&pack("/custom/new.xml")).unwrap().blob(),
        b"<new/>"
    );
    assert_eq!(
        package
            .rels()
            .get("rId2")
            .unwrap()
            .target_partname()
            .unwrap(),
        pack("/custom/new.xml")
    );
    assert_eq!(
        package
            .get_part(&pack("/word/document.xml"))
            .unwrap()
            .rels()
            .get("rId2")
            .unwrap()
            .target_partname()
            .unwrap(),
        pack("/custom/new.xml")
    );
    assert_eq!(
        package
            .get_part(&pack("/custom/new.xml"))
            .unwrap()
            .rels()
            .get("rIdChild")
            .unwrap()
            .target_partname()
            .unwrap(),
        pack("/word/document.xml")
    );

    let text = String::from_utf8(zip_member(&output, "[Content_Types].xml")).unwrap();
    assert!(text.contains("PartName=\"/custom/new.xml\""));
    assert!(text.contains("ContentType=\"application/vnd.example.new+xml\""));
}

#[test]
fn untouched_zip_records_retain_extras_comments_and_payload_records() {
    let source = source_bytes(false, false, false);
    let output = publish(&source, add_basic_topology_plan()).unwrap();
    let before = raw_records(&source);
    let after = raw_records(&output);
    for name in ["custom/existing.xml", "custom/untouched.xml"] {
        assert_eq!(after[name].local, before[name].local, "local record {name}");
        assert_eq!(
            after[name].central, before[name].central,
            "central record {name}"
        );
    }
    assert!(after.contains_key("custom/new.xml"));
    assert!(after.contains_key("custom/_rels/new.xml.rels"));
}

#[test]
fn noncanonical_content_types_are_preserved_while_new_overrides_are_sorted() {
    let source = source_bytes(true, false, false);
    let mut plan = SourceTopologyPlan::new();
    plan.try_add_part(
        pack("/zeta/new.xml"),
        "application/vnd.example.zeta+xml",
        b"<zeta/>".to_vec(),
    )
    .unwrap();
    plan.try_add_part(
        pack("/alpha/new.xml"),
        "application/vnd.example.alpha+xml",
        b"<alpha/>".to_vec(),
    )
    .unwrap();
    let output = publish(&source, plan).unwrap();
    let text = String::from_utf8(zip_member(&output, "[Content_Types].xml")).unwrap();
    assert!(text.contains("<Default ContentType=\"application/xml\" Extension=\"xml\" />"));
    assert!(text.contains("<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\" />"));
    let alpha = text.find("PartName=\"/alpha/new.xml\"").unwrap();
    let zeta = text.find("PartName=\"/zeta/new.xml\"").unwrap();
    assert!(alpha < zeta);
}

#[test]
fn noncanonical_existing_relationships_are_refused_before_output() {
    let source = source_bytes(false, true, true);
    let mut plan = SourceTopologyPlan::new();
    plan.try_add_part(
        pack("/custom/new.xml"),
        "application/vnd.example.new+xml",
        b"<new/>".to_vec(),
    )
    .unwrap();
    plan.try_add_internal_relationship(
        pack("/word/document.xml"),
        "rId2",
        CUSTOM_REL,
        pack("/custom/new.xml"),
    )
    .unwrap();

    let package = open(&source);
    let mut output = Vec::new();
    let error = package
        .write_topology_to_stream(&mut output, plan)
        .unwrap_err();
    assert!(matches!(
        error,
        OpcError::DuplicatePartName(_) | OpcError::SourceBackedOverlayUnavailable { .. }
    ));
    assert!(output.is_empty());
}

#[test]
fn topology_plan_rejects_duplicate_and_invalid_relationship_requests() {
    let mut plan = SourceTopologyPlan::new();
    plan.try_add_part(
        pack("/custom/new.xml"),
        "application/xml",
        b"<new/>".to_vec(),
    )
    .unwrap();
    assert!(matches!(
        plan.try_add_part(
            pack("/CUSTOM/NEW.XML"),
            "application/octet-stream",
            b"<two/>".to_vec()
        ),
        Err(OpcError::DuplicatePartName(_))
    ));
    assert!(matches!(
        plan.try_add_internal_relationship(
            pack("/"),
            "not an xml id",
            CUSTOM_REL,
            pack("/custom/new.xml")
        ),
        Err(OpcError::InvalidRelationship(_))
    ));
    plan.try_add_internal_relationship(pack("/"), "rId2", CUSTOM_REL, pack("/custom/new.xml"))
        .unwrap();
    assert!(matches!(
        plan.try_add_internal_relationship(pack("/"), "rId2", CUSTOM_REL, pack("/custom/new.xml")),
        Err(OpcError::DuplicateRelationshipId(_))
    ));
}

#[test]
fn topology_publish_refuses_collisions_dangling_targets_and_duplicate_existing_ids() {
    let source = source_bytes(false, false, false);

    let mut collision = SourceTopologyPlan::new();
    collision
        .try_add_part(
            pack("/custom/existing.xml"),
            "application/octet-stream",
            b"x".to_vec(),
        )
        .unwrap();
    let mut output = Vec::new();
    let error = open(&source)
        .write_topology_to_stream(&mut output, collision)
        .unwrap_err();
    assert!(matches!(
        error,
        OpcError::DuplicatePartName(_) | OpcError::SourceBackedOverlayUnavailable { .. }
    ));
    assert!(output.is_empty());

    let mut dangling = SourceTopologyPlan::new();
    dangling
        .try_add_internal_relationship(pack("/"), "rId2", CUSTOM_REL, pack("/missing.xml"))
        .unwrap();
    let mut output = Vec::new();
    let error = open(&source)
        .write_topology_to_stream(&mut output, dangling)
        .unwrap_err();
    assert!(matches!(error, OpcError::PartNotFound(_)));
    assert!(output.is_empty());

    let mut duplicate = SourceTopologyPlan::new();
    duplicate
        .try_add_internal_relationship(pack("/"), "rId1", CUSTOM_REL, pack("/custom/existing.xml"))
        .unwrap();
    let mut output = Vec::new();
    let error = open(&source)
        .write_topology_to_stream(&mut output, duplicate)
        .unwrap_err();
    assert!(matches!(error, OpcError::DuplicateRelationshipId(_)));
    assert!(output.is_empty());
}

#[test]
fn topology_publish_rejects_invalid_xml_and_part_byte_limits_before_output() {
    let source = source_bytes(false, false, false);
    let mut malformed = SourceTopologyPlan::new();
    malformed
        .try_replace_part(pack("/word/document.xml"), b"<broken".to_vec())
        .unwrap();
    let mut output = Vec::new();
    let error = open(&source)
        .write_topology_to_stream(&mut output, malformed)
        .unwrap_err();
    assert!(matches!(error, OpcError::XmlPublication { .. }));
    assert!(output.is_empty());

    let limits = ReadLimits::builder()
        .max_part_bytes(64)
        .unwrap()
        .max_total_part_bytes(32)
        .unwrap()
        .build()
        .unwrap();
    let package =
        SourceBackedPackage::from_read_at_with_limits(Arc::new(OwnedSource::new(source)), limits)
            .unwrap();
    let mut oversized = SourceTopologyPlan::new();
    oversized
        .try_add_part(
            pack("/custom/new.xml"),
            "application/octet-stream",
            b"123456789".to_vec(),
        )
        .unwrap();
    let mut output = Vec::new();
    let error = package
        .write_topology_to_stream(&mut output, oversized)
        .unwrap_err();
    assert!(matches!(error, OpcError::ReadLimit { .. }));
    assert!(output.is_empty());
}

#[test]
fn topology_plan_has_a_finite_exact_and_one_over_part_bound() {
    let mut exact = SourceTopologyPlan::new();
    for index in 0..64 {
        exact
            .try_add_part(
                pack(&format!("/custom/part{index}.bin")),
                "application/octet-stream",
                vec![index as u8],
            )
            .unwrap();
    }
    assert!(matches!(
        exact.try_add_part(
            pack("/custom/over.bin"),
            "application/octet-stream",
            vec![0]
        ),
        Err(OpcError::SourceBackedOverlayUnavailable { .. })
    ));
}

struct FailingWriter;

impl Write for FailingWriter {
    fn write(&mut self, _bytes: &[u8]) -> io::Result<usize> {
        Err(io::Error::new(io::ErrorKind::BrokenPipe, "test sink"))
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[test]
fn topology_publication_reports_sink_failures_with_partial_output_state() {
    let source = source_bytes(false, false, false);
    let mut plan = SourceTopologyPlan::new();
    plan.try_replace_part(pack("/word/document.xml"), b"<after/>".to_vec())
        .unwrap();
    let error = open(&source)
        .write_topology_to_stream(FailingWriter, plan)
        .unwrap_err();
    assert!(matches!(
        error,
        OpcError::IncompleteOutput { .. } | OpcError::IoError(_) | OpcError::ZipError(_)
    ));
}

struct PartialFailingWriter {
    bytes: Vec<u8>,
    remaining: usize,
}

impl Write for PartialFailingWriter {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if self.remaining == 0 {
            return Err(io::Error::new(io::ErrorKind::BrokenPipe, "test sink"));
        }
        let accepted = self.remaining.min(bytes.len());
        self.bytes.extend_from_slice(&bytes[..accepted]);
        self.remaining -= accepted;
        Ok(accepted)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[test]
fn topology_publication_reports_incomplete_output_after_partial_sink_progress() {
    let source = source_bytes(false, false, false);
    let mut plan = SourceTopologyPlan::new();
    plan.try_replace_part(pack("/word/document.xml"), b"<after/>".to_vec())
        .unwrap();
    let mut sink = PartialFailingWriter {
        bytes: Vec::new(),
        remaining: 37,
    };
    let error = open(&source)
        .write_topology_to_stream(&mut sink, plan)
        .unwrap_err();
    match error {
        OpcError::IncompleteOutput { written, .. } => {
            assert!(written > 0);
            assert_eq!(written as usize, sink.bytes.len());
        },
        other => panic!("expected incomplete output, got {other:?}"),
    }
}

#[test]
fn topology_publication_refuses_trailing_data_before_writing() {
    let mut source = source_bytes(false, false, false);
    source.extend_from_slice(b"trailing-data");
    let mut plan = SourceTopologyPlan::new();
    plan.try_add_part(
        pack("/custom/new.bin"),
        "application/octet-stream",
        b"new bytes".to_vec(),
    )
    .unwrap();
    let mut output = Vec::new();
    let error = open(&source)
        .write_topology_to_stream(&mut output, plan)
        .unwrap_err();
    assert!(matches!(
        error,
        OpcError::SourceBackedOverlayUnavailable { .. }
    ));
    assert!(output.is_empty());
}

#[test]
fn topology_publication_refuses_an_encrypted_entry_before_writing() {
    let source = mark_entry_encrypted(source_bytes(false, false, false), "custom/untouched.xml");
    let mut plan = SourceTopologyPlan::new();
    plan.try_add_part(
        pack("/custom/new.bin"),
        "application/octet-stream",
        b"new bytes".to_vec(),
    )
    .unwrap();
    let mut output = Vec::new();
    let error = open(&source)
        .write_topology_to_stream(&mut output, plan)
        .unwrap_err();
    assert!(matches!(
        error,
        OpcError::SourceBackedOverlayUnavailable { .. }
    ));
    assert!(output.is_empty());
}

#[test]
fn signed_exact_noop_replacement_is_byte_exact_but_topology_change_is_refused() {
    let source = signed_source_bytes();
    let mut noop = SourceTopologyPlan::new();
    noop.try_replace_part(pack("/word/document.xml"), b"<before/>".to_vec())
        .unwrap();
    let output = publish(&source, noop).unwrap();
    assert_eq!(output, source);

    let mut change = SourceTopologyPlan::new();
    change
        .try_add_part(
            pack("/custom/new.bin"),
            "application/octet-stream",
            b"new bytes".to_vec(),
        )
        .unwrap();
    let mut output = Vec::new();
    let error = open(&source)
        .write_topology_to_stream(&mut output, change)
        .unwrap_err();
    assert!(matches!(
        error,
        OpcError::SignedSourceRequiresExplicitPolicy
    ));
    assert!(output.is_empty());
}

#[test]
fn topology_output_is_deterministic_and_reopens_after_repeated_publication() {
    let source = source_bytes(false, false, false);
    let first = publish(&source, add_basic_topology_plan()).unwrap();
    let second = publish(&source, add_basic_topology_plan()).unwrap();
    assert_eq!(first, second);
    let reopened = OpcPackage::from_bytes(&first).unwrap();
    assert_eq!(
        reopened.get_part(&pack("/custom/new.xml")).unwrap().blob(),
        b"<new/>"
    );
}

#[test]
fn malformed_exact_replacement_is_a_byte_exact_noop() {
    let source = malformed_xml_source_bytes();
    let mut plan = SourceTopologyPlan::new();
    plan.try_replace_part(pack("/word/document.xml"), b"<broken".to_vec())
        .unwrap();
    let output = publish(&source, plan).unwrap();
    assert_eq!(output, source);
}

#[test]
fn case_equivalent_relationship_owners_group_once_and_refuse_duplicate_ids() {
    let source = source_bytes(false, false, false);
    let mut plan = SourceTopologyPlan::new();
    plan.try_add_part(
        pack("/custom/owner-case.bin"),
        "application/octet-stream",
        b"case target".to_vec(),
    )
    .unwrap();
    plan.try_add_internal_relationship(
        pack("/WORD/DOCUMENT.XML"),
        "rIdCaseA",
        CUSTOM_REL,
        pack("/custom/owner-case.bin"),
    )
    .unwrap();
    plan.try_add_internal_relationship(
        pack("/word/document.xml"),
        "rIdCaseB",
        CUSTOM_REL,
        pack("/custom/owner-case.bin"),
    )
    .unwrap();
    let first = publish(&source, plan).unwrap();

    let mut repeated = SourceTopologyPlan::new();
    repeated
        .try_add_part(
            pack("/custom/owner-case.bin"),
            "application/octet-stream",
            b"case target".to_vec(),
        )
        .unwrap();
    repeated
        .try_add_internal_relationship(
            pack("/word/document.xml"),
            "rIdCaseB",
            CUSTOM_REL,
            pack("/custom/owner-case.bin"),
        )
        .unwrap();
    repeated
        .try_add_internal_relationship(
            pack("/WORD/DOCUMENT.XML"),
            "rIdCaseA",
            CUSTOM_REL,
            pack("/custom/owner-case.bin"),
        )
        .unwrap();
    let second = publish(&source, repeated).unwrap();
    assert_eq!(first, second);
    let package = OpcPackage::from_bytes(&first).unwrap();
    let rels = package
        .get_part(&pack("/word/document.xml"))
        .unwrap()
        .rels();
    assert!(rels.get("rIdCaseA").is_some());
    assert!(rels.get("rIdCaseB").is_some());

    let mut duplicate = SourceTopologyPlan::new();
    duplicate
        .try_add_internal_relationship(
            pack("/WORD/DOCUMENT.XML"),
            "rIdSame",
            CUSTOM_REL,
            pack("/custom/existing.xml"),
        )
        .unwrap();
    assert!(matches!(
        duplicate.try_add_internal_relationship(
            pack("/word/document.xml"),
            "rIdSame",
            CUSTOM_REL,
            pack("/custom/existing.xml"),
        ),
        Err(OpcError::DuplicateRelationshipId(_))
    ));
}

#[test]
fn existing_canonical_relationships_are_regenerated_with_new_relationships() {
    let source = source_bytes(false, true, false);
    let mut plan = SourceTopologyPlan::new();
    plan.try_add_part(
        pack("/custom/rel-target.bin"),
        "application/octet-stream",
        b"target".to_vec(),
    )
    .unwrap();
    plan.try_add_internal_relationship(
        pack("/word/document.xml"),
        "rId2",
        CUSTOM_REL,
        pack("/custom/rel-target.bin"),
    )
    .unwrap();
    let output = publish(&source, plan).unwrap();
    let package = OpcPackage::from_bytes(&output).unwrap();
    let rels = package
        .get_part(&pack("/word/document.xml"))
        .unwrap()
        .rels();
    assert!(rels.get("rId1").is_some());
    assert_eq!(
        rels.get("rId2").unwrap().target_partname().unwrap(),
        pack("/custom/rel-target.bin")
    );
}

#[test]
fn managed_cancellation_before_topology_publication_writes_nothing() {
    let source = source_bytes(false, false, false);
    let (cancellation_source, context) = managed_context();
    let package = SourceBackedPackage::from_read_at_with_execution_context(
        Arc::new(OwnedSource::new(source)),
        ReadLimits::default(),
        context,
    )
    .unwrap();
    cancellation_source.cancel();
    let mut plan = SourceTopologyPlan::new();
    plan.try_add_part(
        pack("/custom/cancelled.bin"),
        "application/octet-stream",
        b"cancelled".to_vec(),
    )
    .unwrap();
    let mut output = Vec::new();
    let error = package
        .write_topology_to_stream(&mut output, plan)
        .unwrap_err();
    assert!(matches!(error, OpcError::Cancelled));
    assert!(output.is_empty());
}
