#![allow(
    clippy::unwrap_used,
    reason = "Fixed in-memory ODF fixtures keep raw preservation assertions concise."
)]

use std::collections::BTreeMap;

use litchi_odf_common::{
    constants,
    core::{AuthoredXmlFragment, OwnedPackage, PackageWriter, XmlSourcePart, XmlSplicePublication},
    package::{raw_identical_members, replace_content_xml, replace_content_xml_spliced},
};
use soapberry_zip::{PreservationIndex, ZipArchive};

const SOURCE_CONTENT: &str = r#"<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><office:body><office:spreadsheet><text:p>source</text:p></office:spreadsheet></office:body></office:document-content>"#;
const TARGET_CONTENT: &str = r#"<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><office:body><office:spreadsheet><text:p>target</text:p></office:spreadsheet></office:body></office:document-content>"#;

#[derive(Debug, PartialEq, Eq)]
struct RawMember {
    local: Vec<u8>,
    central_without_offset: Vec<u8>,
}

fn media_payload() -> Vec<u8> {
    let mut state = 0xd1b5_4a32_d192_ed03_u64;
    (0..4 * 1024 * 1024)
        .map(|_| {
            state ^= state << 13;
            state ^= state >> 7;
            state ^= state << 17;
            (state >> 24) as u8
        })
        .collect()
}

fn source_package(with_signature: bool) -> Vec<u8> {
    let mut writer = PackageWriter::new();
    writer.set_mimetype(constants::ODF_SPREADSHEET).unwrap();
    writer
        .add_file("content.xml", SOURCE_CONTENT.as_bytes())
        .unwrap();
    writer
        .add_file_with_media_type(
            "Pictures/opaque.bin",
            &media_payload(),
            "application/octet-stream",
        )
        .unwrap();
    writer
        .add_file("Vendor/unknown.bin", b"unknown source bytes")
        .unwrap();
    if with_signature {
        writer
            .add_file("META-INF/documentsignatures.xml", b"<document-signatures/>")
            .unwrap();
    }
    writer.finish_to_bytes().unwrap()
}

fn raw_members(bytes: &[u8]) -> BTreeMap<String, RawMember> {
    let archive = ZipArchive::from_slice(bytes).unwrap().into_zip_archive();
    let mut buffer = vec![0_u8; soapberry_zip::RECOMMENDED_BUFFER_SIZE];
    let index = PreservationIndex::new(&archive, &mut buffer).unwrap();
    let mut records = archive.entries(&mut buffer);
    index
        .entries()
        .iter()
        .map(|preserved| {
            let record = records.next_entry().unwrap().unwrap();
            let name = record
                .file_path()
                .try_normalize()
                .unwrap()
                .as_ref()
                .to_owned();
            let local = bytes
                [preserved.local_span().start as usize..preserved.local_span().end as usize]
                .to_vec();
            let central_range = preserved.central_record();
            let mut central_without_offset =
                bytes[central_range.start as usize..central_range.end as usize].to_vec();
            central_without_offset[42..46].fill(0);
            (
                name,
                RawMember {
                    local,
                    central_without_offset,
                },
            )
        })
        .collect()
}

fn text_replacement_publication(source: &OwnedPackage) -> XmlSplicePublication {
    let part = XmlSourcePart::load(source, "content.xml").unwrap();
    let start = part
        .bytes()
        .windows(b"source".len())
        .position(|window| window == b"source")
        .unwrap();
    let range = start..start + b"source".len();
    let proof = part.checked_range(range, b"source").unwrap();
    let mut publication = XmlSplicePublication::new(part);
    publication
        .replace(
            proof,
            AuthoredXmlFragment::text(b"target".to_vec()).unwrap(),
        )
        .unwrap();
    publication
}

#[test]
fn content_replacement_raw_copies_every_untouched_member() {
    let source = source_package(false);
    let source_package = OwnedPackage::from_bytes(source.clone()).unwrap();
    let output = replace_content_xml(&source_package, TARGET_CONTENT).unwrap();
    let source_raw = raw_members(&source);
    let output_raw = raw_members(&output);

    assert_ne!(output, source);
    assert_eq!(
        source_raw.keys().collect::<Vec<_>>(),
        output_raw.keys().collect::<Vec<_>>()
    );
    for (name, source_member) in &source_raw {
        if name == "content.xml" {
            assert_ne!(output_raw[name].local, source_member.local);
        } else {
            assert_eq!(output_raw[name], *source_member, "{name}");
        }
    }
    let raw_identical = raw_identical_members(&source, &output).unwrap();
    assert!(!raw_identical.contains("content.xml"));
    assert!(raw_identical.contains("META-INF/manifest.xml"));
    assert!(raw_identical.contains("Pictures/opaque.bin"));
    assert!(raw_identical.contains("Vendor/unknown.bin"));

    let reopened = OwnedPackage::from_bytes(output).unwrap();
    assert_eq!(
        reopened.get_file("content.xml").unwrap(),
        TARGET_CONTENT.as_bytes()
    );
    assert_eq!(
        reopened.get_file("Pictures/opaque.bin").unwrap(),
        media_payload()
    );
    assert_eq!(
        reopened.get_file("Vendor/unknown.bin").unwrap(),
        b"unknown source bytes"
    );
    assert_eq!(
        replace_content_xml(&source_package, SOURCE_CONTENT).unwrap(),
        source
    );
}

#[test]
fn signed_content_replacement_uses_the_signature_stripping_fallback() {
    let source = OwnedPackage::from_bytes(source_package(true)).unwrap();
    let output = replace_content_xml(&source, TARGET_CONTENT).unwrap();
    let reopened = OwnedPackage::from_bytes(output).unwrap();

    assert_eq!(
        reopened.get_file("content.xml").unwrap(),
        TARGET_CONTENT.as_bytes()
    );
    assert!(
        !reopened
            .has_file("META-INF/documentsignatures.xml")
            .unwrap()
    );
    assert_eq!(
        reopened.get_file("Pictures/opaque.bin").unwrap(),
        media_payload()
    );

    let spliced_output = replace_content_xml_spliced(
        &source,
        TARGET_CONTENT,
        text_replacement_publication(&source),
    )
    .unwrap();
    let spliced = OwnedPackage::from_bytes(spliced_output).unwrap();
    assert_eq!(
        spliced.get_file("content.xml").unwrap(),
        TARGET_CONTENT.as_bytes()
    );
    assert!(!spliced.has_file("META-INF/documentsignatures.xml").unwrap());
    assert_eq!(
        spliced.get_file("Pictures/opaque.bin").unwrap(),
        media_payload()
    );
}

#[test]
fn explicit_content_splice_preserves_raw_members_and_checks_provenance() {
    let source = source_package(false);
    let source_package = OwnedPackage::from_bytes(source.clone()).unwrap();
    let output = replace_content_xml_spliced(
        &source_package,
        TARGET_CONTENT,
        text_replacement_publication(&source_package),
    )
    .unwrap();
    let identical = raw_identical_members(&source, &output).unwrap();
    assert!(identical.contains("META-INF/manifest.xml"));
    assert!(identical.contains("Pictures/opaque.bin"));
    assert!(identical.contains("Vendor/unknown.bin"));
    assert!(!identical.contains("content.xml"));

    let identical_but_foreign = OwnedPackage::from_bytes(source).unwrap();
    assert!(
        replace_content_xml_spliced(
            &identical_but_foreign,
            TARGET_CONTENT,
            text_replacement_publication(&source_package),
        )
        .is_err()
    );
    assert!(
        replace_content_xml_spliced(
            &source_package,
            SOURCE_CONTENT,
            text_replacement_publication(&source_package),
        )
        .is_err()
    );
}
