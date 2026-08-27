#![cfg(any(unix, windows))]
#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "focused filesystem-source assertions"
)]

//! Filesystem-backed source tests for the lazy OPC catalog ingress.

use std::fs;

use litchi_opc::{OpcError, PackURI, ReadLimits, SourceBackedPackage};
use soapberry_zip::office::StreamingArchiveWriter;
use tempfile::NamedTempFile;

const CONTENT_TYPES_NS: &str = "http://schemas.openxmlformats.org/package/2006/content-types";
const RELATIONSHIPS_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const OFFICE_DOCUMENT_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const SELECTED_MEMBER: &str = "word/document.xml";
const SELECTED_URI: &str = "/word/document.xml";
const UNUSED_MEMBER: &str = "custom/unused.bin";
const SELECTED_PAYLOAD: &[u8] = b"<document>selected</document>";
const UNUSED_PAYLOAD: &[u8] = b"unselected ordinary payload";

fn pack(uri: &str) -> PackURI {
    PackURI::new(uri).unwrap()
}

fn archive_bytes(selected: &[u8], unused: &[u8]) -> Vec<u8> {
    let content_types = format!(
        r#"<Types xmlns="{CONTENT_TYPES_NS}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="bin" ContentType="application/octet-stream"/></Types>"#
    );
    let root_relationships = format!(
        r#"<Relationships xmlns="{RELATIONSHIPS_NS}"><Relationship Id="rId1" Type="{OFFICE_DOCUMENT_REL}" Target="{SELECTED_MEMBER}"/></Relationships>"#
    );

    let mut writer = StreamingArchiveWriter::new();
    writer
        .write_stored("[Content_Types].xml", content_types.as_bytes())
        .unwrap();
    writer
        .write_stored("_rels/.rels", root_relationships.as_bytes())
        .unwrap();
    writer
        .write_deflated_sized(SELECTED_MEMBER, selected)
        .unwrap();
    writer.write_deflated_sized(UNUSED_MEMBER, unused).unwrap();
    writer.finish_to_bytes().unwrap()
}

fn write_fixture(bytes: &[u8]) -> NamedTempFile {
    let file = NamedTempFile::new().unwrap();
    fs::write(file.path(), bytes).unwrap();
    file
}

fn catalog(package: &SourceBackedPackage) -> Vec<(String, String, u64)> {
    package
        .iter_parts()
        .map(|part| {
            (
                part.partname().as_str().to_owned(),
                part.content_type().to_owned(),
                part.declared_uncompressed_size().unwrap(),
            )
        })
        .collect()
}

fn selected_payload(package: &SourceBackedPackage) -> Vec<u8> {
    package
        .part(&pack(SELECTED_URI))
        .unwrap()
        .data()
        .unwrap()
        .as_bytes()
        .to_vec()
}

fn read_u16(bytes: &[u8], offset: usize) -> usize {
    usize::from(u16::from_le_bytes(
        bytes[offset..offset + 2].try_into().unwrap(),
    ))
}

fn read_u32(bytes: &[u8], offset: usize) -> usize {
    usize::try_from(u32::from_le_bytes(
        bytes[offset..offset + 4].try_into().unwrap(),
    ))
    .unwrap()
}

fn corrupt_member_crc(mut bytes: Vec<u8>, wanted: &str) -> Vec<u8> {
    let eocd = bytes
        .windows(4)
        .rposition(|window| window == 0x0605_4b50_u32.to_le_bytes())
        .unwrap();
    let count = read_u16(&bytes, eocd + 10);
    let mut cursor = read_u32(&bytes, eocd + 16);

    for _ in 0..count {
        assert_eq!(&bytes[cursor..cursor + 4], &0x0201_4b50_u32.to_le_bytes());
        let name_len = read_u16(&bytes, cursor + 28);
        let extra_len = read_u16(&bytes, cursor + 30);
        let comment_len = read_u16(&bytes, cursor + 32);
        let name_start = cursor + 46;
        let local_offset = read_u32(&bytes, cursor + 42);
        if &bytes[name_start..name_start + name_len] == wanted.as_bytes() {
            let crc = u32::from_le_bytes(bytes[cursor + 16..cursor + 20].try_into().unwrap()) ^ 1;
            bytes[cursor + 16..cursor + 20].copy_from_slice(&crc.to_le_bytes());

            let local_crc = u32::from_le_bytes(
                bytes[local_offset + 14..local_offset + 18]
                    .try_into()
                    .unwrap(),
            ) ^ 1;
            bytes[local_offset + 14..local_offset + 18].copy_from_slice(&local_crc.to_le_bytes());
            return bytes;
        }
        cursor += 46 + name_len + extra_len + comment_len;
    }

    panic!("missing ZIP member {wanted}");
}

#[test]
fn path_default_catalog_and_selected_data_match_from_vec_without_eager_unselected_read() {
    let source = corrupt_member_crc(
        archive_bytes(SELECTED_PAYLOAD, UNUSED_PAYLOAD),
        UNUSED_MEMBER,
    );
    let file = write_fixture(&source);

    let from_vec = SourceBackedPackage::from_vec(source).unwrap();
    let from_path = SourceBackedPackage::from_path(file.path()).unwrap();

    assert_eq!(catalog(&from_path), catalog(&from_vec));
    assert_eq!(
        from_path.physical_member_names().collect::<Vec<_>>(),
        from_vec.physical_member_names().collect::<Vec<_>>()
    );
    assert_eq!(from_path.cache_diagnostics().cold_loads, 0);
    assert_eq!(from_path.cache_diagnostics().successful_loads, 0);
    assert_eq!(selected_payload(&from_path), SELECTED_PAYLOAD);
    assert_eq!(from_path.cache_diagnostics().cold_loads, 1);
    assert_eq!(from_path.cache_diagnostics().successful_loads, 1);

    let error = from_path
        .part(&pack("/custom/unused.bin"))
        .unwrap()
        .data()
        .unwrap_err();
    assert!(matches!(error, OpcError::ZipError(_)));
}

#[test]
fn path_with_limits_matches_from_vec_and_rejects_exact_too_small_input_and_part_limits() {
    let unused = vec![0xA5; 128];
    let source = archive_bytes(SELECTED_PAYLOAD, &unused);
    let file = write_fixture(&source);

    let exact_input = ReadLimits::builder()
        .max_input_bytes(source.len() as u64)
        .unwrap()
        .build()
        .unwrap();
    let from_path = SourceBackedPackage::from_path_with_limits(file.path(), exact_input).unwrap();
    let from_vec = SourceBackedPackage::from_vec_with_limits(
        source.clone(),
        ReadLimits::builder()
            .max_input_bytes(source.len() as u64)
            .unwrap()
            .build()
            .unwrap(),
    )
    .unwrap();
    assert_eq!(catalog(&from_path), catalog(&from_vec));
    assert_eq!(selected_payload(&from_path), SELECTED_PAYLOAD);

    let too_small_input = ReadLimits::builder()
        .max_input_bytes((source.len() - 1) as u64)
        .unwrap()
        .build()
        .unwrap();
    assert!(matches!(
        SourceBackedPackage::from_path_with_limits(file.path(), too_small_input),
        Err(OpcError::ReadLimit { .. })
    ));
    assert!(matches!(
        SourceBackedPackage::from_vec_with_limits(
            source.clone(),
            ReadLimits::builder()
                .max_input_bytes((source.len() - 1) as u64)
                .unwrap()
                .build()
                .unwrap(),
        ),
        Err(OpcError::ReadLimit { .. })
    ));

    let too_small_part = ReadLimits::builder()
        .max_part_bytes((unused.len() - 1) as u64)
        .unwrap()
        .build()
        .unwrap();
    assert!(matches!(
        SourceBackedPackage::from_path_with_limits(file.path(), too_small_part),
        Err(OpcError::ReadLimit { .. })
    ));
    assert!(matches!(
        SourceBackedPackage::from_vec_with_limits(
            source,
            ReadLimits::builder()
                .max_part_bytes((unused.len() - 1) as u64)
                .unwrap()
                .build()
                .unwrap(),
        ),
        Err(OpcError::ReadLimit { .. })
    ));
}

#[test]
fn path_source_reports_external_change_without_retargeting_or_stale_hydration() {
    let original = archive_bytes(SELECTED_PAYLOAD, UNUSED_PAYLOAD);
    let replacement = archive_bytes(b"<document>replacement</document>", UNUSED_PAYLOAD);
    let file = write_fixture(&original);
    let package = SourceBackedPackage::from_path(file.path()).unwrap();
    let selected = package.part(&pack(SELECTED_URI)).unwrap();
    fs::write(file.path(), replacement).unwrap();

    assert!(matches!(
        package.source_version(),
        Err(OpcError::SourceChanged { .. })
    ));
    assert!(matches!(
        selected.data(),
        Err(OpcError::SourceChanged { .. })
    ));

    let reopened = SourceBackedPackage::from_path(file.path()).unwrap();
    assert_eq!(
        selected_payload(&reopened),
        b"<document>replacement</document>"
    );
}
