#![allow(
    clippy::unwrap_used,
    reason = "fixed ZIP fixtures keep repair assertions focused"
)]

use std::io::{self, Cursor, Write};

use litchi_odf_common::{
    MIMETYPE_LOCAL_EXTRA_REPAIR, MIMETYPE_REPAIR_PLAN_SCHEMA, OdfRepairLimits, RepairChangedRegion,
    RepairError, RepairIntentKind, RepairOutputProgress, plan_mimetype_local_extra,
    plan_odf_repair, validate_package,
};
use zip::{CompressionMethod, ZipWriter, write::FullFileOptions};

const MANIFEST_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0";
const OFFICE_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const MIME: &str = "application/vnd.oasis.opendocument.text";

fn package(extra_id: u16, extra_body: &[u8]) -> Vec<u8> {
    let manifest = format!(
        "<manifest:manifest xmlns:manifest=\"{MANIFEST_NS}\"><manifest:file-entry manifest:full-path=\"/\" manifest:media-type=\"{MIME}\"/><manifest:file-entry manifest:full-path=\"content.xml\" manifest:media-type=\"text/xml\"/></manifest:manifest>"
    );
    let content = format!(
        "<office:document-content xmlns:office=\"{OFFICE_NS}\"><office:body><office:text/></office:body></office:document-content>"
    );
    let mut zip = ZipWriter::new(Cursor::new(Vec::new()));
    let mut mimetype = FullFileOptions::default().compression_method(CompressionMethod::Stored);
    mimetype
        .add_extra_data(extra_id, extra_body, false)
        .unwrap();
    zip.start_file("mimetype", mimetype).unwrap();
    zip.write_all(MIME.as_bytes()).unwrap();
    zip.start_file(
        "META-INF/manifest.xml",
        FullFileOptions::default().compression_method(CompressionMethod::Deflated),
    )
    .unwrap();
    zip.write_all(manifest.as_bytes()).unwrap();
    zip.start_file(
        "content.xml",
        FullFileOptions::default().compression_method(CompressionMethod::Deflated),
    )
    .unwrap();
    zip.write_all(content.as_bytes()).unwrap();
    zip.finish().unwrap().into_inner()
}

fn package_with_directory(extra_id: u16, extra_body: &[u8]) -> Vec<u8> {
    let manifest = format!(
        "<manifest:manifest xmlns:manifest=\"{MANIFEST_NS}\"><manifest:file-entry manifest:full-path=\"/\" manifest:media-type=\"{MIME}\"/><manifest:file-entry manifest:full-path=\"content.xml\" manifest:media-type=\"text/xml\"/></manifest:manifest>"
    );
    let content = format!(
        "<office:document-content xmlns:office=\"{OFFICE_NS}\"><office:body><office:text/></office:body></office:document-content>"
    );
    let mut zip = ZipWriter::new(Cursor::new(Vec::new()));
    let mut mimetype = FullFileOptions::default().compression_method(CompressionMethod::Stored);
    mimetype
        .add_extra_data(extra_id, extra_body, false)
        .unwrap();
    zip.start_file("mimetype", mimetype).unwrap();
    zip.write_all(MIME.as_bytes()).unwrap();
    zip.add_directory(
        "META-INF/",
        FullFileOptions::default().compression_method(CompressionMethod::Stored),
    )
    .unwrap();
    zip.start_file(
        "META-INF/manifest.xml",
        FullFileOptions::default().compression_method(CompressionMethod::Deflated),
    )
    .unwrap();
    zip.write_all(manifest.as_bytes()).unwrap();
    zip.start_file(
        "content.xml",
        FullFileOptions::default().compression_method(CompressionMethod::Deflated),
    )
    .unwrap();
    zip.write_all(content.as_bytes()).unwrap();
    zip.finish().unwrap().into_inner()
}

fn package_with_xml_member(path: &str, media_type: &str, xml: &[u8]) -> Vec<u8> {
    let manifest = format!(
        "<manifest:manifest xmlns:manifest=\"{MANIFEST_NS}\"><manifest:file-entry manifest:full-path=\"/\" manifest:media-type=\"{MIME}\"/><manifest:file-entry manifest:full-path=\"content.xml\" manifest:media-type=\"text/xml\"/><manifest:file-entry manifest:full-path=\"{path}\" manifest:media-type=\"{media_type}\"/></manifest:manifest>"
    );
    let content = format!(
        "<office:document-content xmlns:office=\"{OFFICE_NS}\"><office:body><office:text/></office:body></office:document-content>"
    );
    let mut zip = ZipWriter::new(Cursor::new(Vec::new()));
    let mut mimetype = FullFileOptions::default().compression_method(CompressionMethod::Stored);
    mimetype
        .add_extra_data(0x5455, [1, 0, 0, 0, 0], false)
        .unwrap();
    zip.start_file("mimetype", mimetype).unwrap();
    zip.write_all(MIME.as_bytes()).unwrap();
    zip.start_file(
        "META-INF/manifest.xml",
        FullFileOptions::default().compression_method(CompressionMethod::Deflated),
    )
    .unwrap();
    zip.write_all(manifest.as_bytes()).unwrap();
    zip.start_file(
        "content.xml",
        FullFileOptions::default().compression_method(CompressionMethod::Deflated),
    )
    .unwrap();
    zip.write_all(content.as_bytes()).unwrap();
    zip.start_file(
        path,
        FullFileOptions::default().compression_method(CompressionMethod::Deflated),
    )
    .unwrap();
    zip.write_all(xml).unwrap();
    remove_central_extra(zip.finish().unwrap().into_inner())
}

fn package_with_mimetype_last() -> Vec<u8> {
    let manifest = format!(
        "<manifest:manifest xmlns:manifest=\"{MANIFEST_NS}\"><manifest:file-entry manifest:full-path=\"/\" manifest:media-type=\"{MIME}\"/><manifest:file-entry manifest:full-path=\"content.xml\" manifest:media-type=\"text/xml\"/></manifest:manifest>"
    );
    let content = format!(
        "<office:document-content xmlns:office=\"{OFFICE_NS}\"><office:body><office:text/></office:body></office:document-content>"
    );
    let mut zip = ZipWriter::new(Cursor::new(Vec::new()));
    let stored = FullFileOptions::default().compression_method(CompressionMethod::Stored);
    let deflated = FullFileOptions::default().compression_method(CompressionMethod::Deflated);
    zip.start_file("META-INF/manifest.xml", deflated.clone())
        .unwrap();
    zip.write_all(manifest.as_bytes()).unwrap();
    zip.start_file("content.xml", deflated).unwrap();
    zip.write_all(content.as_bytes()).unwrap();
    let mut mimetype = stored;
    mimetype
        .add_extra_data(0x5455, [1, 0, 0, 0, 0], false)
        .unwrap();
    zip.start_file("mimetype", mimetype).unwrap();
    zip.write_all(MIME.as_bytes()).unwrap();
    zip.finish().unwrap().into_inner()
}

fn package_metadata_base(member_count: u64, names: &[&str]) -> u64 {
    let raw_name_bytes = names.iter().map(|name| name.len() as u64).sum::<u64>();
    let central_record_bytes = member_count * 46 + raw_name_bytes;
    member_count * (128 + 64 + 128 + 16) + raw_name_bytes + central_record_bytes
}

fn package_manifest_metadata_bytes() -> u64 {
    1 + MIME.len() as u64 + "content.xml".len() as u64 + "text/xml".len() as u64
}

fn remove_central_extra(mut source: Vec<u8>) -> Vec<u8> {
    let archive = soapberry_zip::ZipArchive::from_slice(&source).unwrap();
    let eocd = archive.eocd_offset() as usize;
    let first = archive.entries().next().unwrap().unwrap();
    let record_start = first.central_directory_offset() as usize;
    let name_len =
        u16::from_le_bytes([source[record_start + 28], source[record_start + 29]]) as usize;
    let extra_len =
        u16::from_le_bytes([source[record_start + 30], source[record_start + 31]]) as usize;
    assert!(extra_len > 0);
    let extra_start = record_start + 46 + name_len;
    let extra_end = extra_start + extra_len;
    let mut output = Vec::with_capacity(source.len() - extra_len);
    output.extend_from_slice(&source[..extra_start]);
    output.extend_from_slice(&source[extra_end..eocd]);
    output.extend_from_slice(&source[eocd..]);
    output[record_start + 30..record_start + 32].copy_from_slice(&[0, 0]);
    let new_eocd = eocd - extra_len;
    let central_size = u32::from_le_bytes(output[new_eocd + 12..new_eocd + 16].try_into().unwrap());
    output[new_eocd + 12..new_eocd + 16]
        .copy_from_slice(&(central_size - extra_len as u32).to_le_bytes());
    source.clear();
    output
}

fn insert_local_padding(mut source: Vec<u8>) -> Vec<u8> {
    let archive = soapberry_zip::ZipArchive::from_slice(&source).unwrap();
    let central_start = archive.directory_offset() as usize;
    let eocd = archive.eocd_offset() as usize;
    let mut records = Vec::new();
    let entries = archive.entries();
    for entry in entries {
        let entry = entry.unwrap();
        records.push((
            entry.central_directory_offset() as usize,
            entry.local_header_offset() as usize,
        ));
    }
    let second_local = records
        .iter()
        .map(|(_, local)| *local)
        .filter(|local| *local > 0)
        .min()
        .unwrap();
    source.insert(second_local, 0xa5);
    for (record_start, local_offset) in records {
        let shifted_record = record_start + 1;
        let shifted_local = if local_offset >= second_local {
            local_offset + 1
        } else {
            local_offset
        };
        source[shifted_record + 42..shifted_record + 46]
            .copy_from_slice(&(shifted_local as u32).to_le_bytes());
    }
    let shifted_eocd = eocd + 1;
    source[shifted_eocd + 16..shifted_eocd + 20]
        .copy_from_slice(&((central_start + 1) as u32).to_le_bytes());
    source
}

fn insert_central_padding(mut source: Vec<u8>) -> Vec<u8> {
    let archive = soapberry_zip::ZipArchive::from_slice(&source).unwrap();
    let central_start = archive.directory_offset() as usize;
    let eocd = archive.eocd_offset() as usize;
    source.insert(central_start, 0xa5);
    let shifted_eocd = eocd + 1;
    let central_size = u32::from_le_bytes(
        source[shifted_eocd + 12..shifted_eocd + 16]
            .try_into()
            .unwrap(),
    );
    source[shifted_eocd + 12..shifted_eocd + 16].copy_from_slice(&(central_size + 1).to_le_bytes());
    source
}

fn add_descriptor_to_second_member(mut source: Vec<u8>) -> Vec<u8> {
    let archive = soapberry_zip::ZipArchive::from_slice(&source).unwrap();
    let central_start = archive.directory_offset() as usize;
    let eocd = archive.eocd_offset() as usize;
    let mut records = Vec::new();
    for entry in archive.entries() {
        let entry = entry.unwrap();
        records.push((
            entry.central_directory_offset() as usize,
            entry.local_header_offset() as usize,
            entry.compressed_size_hint() as usize,
            entry.crc32(),
            entry.uncompressed_size_hint() as usize,
        ));
    }
    let (_, second_local, second_compressed, crc, uncompressed) = records[1];
    let name_len = u16::from_le_bytes(
        source[second_local + 26..second_local + 28]
            .try_into()
            .unwrap(),
    ) as usize;
    let extra_len = u16::from_le_bytes(
        source[second_local + 28..second_local + 30]
            .try_into()
            .unwrap(),
    ) as usize;
    let descriptor_at = second_local + 30 + name_len + extra_len + second_compressed;
    let mut descriptor = Vec::with_capacity(16);
    descriptor.extend_from_slice(&0x0807_4b50_u32.to_le_bytes());
    descriptor.extend_from_slice(&crc.to_le_bytes());
    descriptor.extend_from_slice(&(second_compressed as u32).to_le_bytes());
    descriptor.extend_from_slice(&(uncompressed as u32).to_le_bytes());
    source.splice(descriptor_at..descriptor_at, descriptor);

    for &(record_start, local_offset, _, _, _) in &records {
        let shifted_record = record_start + 16;
        let shifted_local = if local_offset >= descriptor_at {
            local_offset + 16
        } else {
            local_offset
        };
        source[shifted_record + 42..shifted_record + 46]
            .copy_from_slice(&(shifted_local as u32).to_le_bytes());
    }
    let second_record = records[1].0 + 16;
    let central_flags = u16::from_le_bytes(
        source[second_record + 8..second_record + 10]
            .try_into()
            .unwrap(),
    );
    source[second_record + 8..second_record + 10]
        .copy_from_slice(&(central_flags | 0x08).to_le_bytes());
    let second_local_shifted = second_local;
    let local_flags = u16::from_le_bytes(
        source[second_local_shifted + 6..second_local_shifted + 8]
            .try_into()
            .unwrap(),
    );
    source[second_local_shifted + 6..second_local_shifted + 8]
        .copy_from_slice(&(local_flags | 0x08).to_le_bytes());
    source[second_local_shifted + 14..second_local_shifted + 26].fill(0);

    let shifted_eocd = eocd + 16;
    source[shifted_eocd + 16..shifted_eocd + 20]
        .copy_from_slice(&((central_start + 16) as u32).to_le_bytes());
    source
}

#[test]
fn valid_plan_is_bounded_deterministic_and_non_destructive() {
    let source = remove_central_extra(package(0x5455, &[1, 0, 0, 0, 0]));
    let source_before = source.clone();
    let report = validate_package(&source).unwrap();
    let plan = plan_mimetype_local_extra(&source, &report, OdfRepairLimits::default()).unwrap();
    assert_eq!(plan.action().field_id(), 0x5455);
    assert_eq!(plan.action().field_bytes(), 9);
    assert_eq!(plan.output_len(), source.len() as u64 - 9);
    assert_eq!(plan.to_json().unwrap(), plan.to_json().unwrap());
    assert!(
        plan.to_json()
            .unwrap()
            .contains(MIMETYPE_LOCAL_EXTRA_REPAIR)
    );
    assert!(!plan.to_json().unwrap().contains(MIME));

    let mut output = Vec::new();
    let publication = plan.write_to(&mut output).unwrap();
    assert_eq!(publication.bytes(), output.len() as u64);
    assert_eq!(output.len(), plan.output_len() as usize);
    assert_eq!(validate_package(&output).unwrap().issues().len(), 0);
    assert_eq!(source, source_before);
}

#[test]
fn typed_preview_reports_effects_and_apply_inverse_restores_exact_source() {
    let source = remove_central_extra(package(0x5455, &[1, 0, 0, 0, 0]));
    let source_before = source.clone();
    let report = validate_package(&source).unwrap();
    let plan = plan_odf_repair(&source, &report, OdfRepairLimits::default()).unwrap();
    let preview = plan.preview();
    assert_eq!(preview.schema(), MIMETYPE_REPAIR_PLAN_SCHEMA);
    assert_eq!(preview.validation_issue_id(), report.issues()[0].id());
    assert_eq!(preview.intent(), RepairIntentKind::NonDestructive);
    assert_eq!(preview.repair_id(), MIMETYPE_LOCAL_EXTRA_REPAIR);
    assert!(!preview.is_noop());
    assert_eq!(preview.effects().changed_members(), ["mimetype"]);
    assert_eq!(
        preview.effects().changed_regions(),
        [
            RepairChangedRegion::MimetypeLocalHeader,
            RepairChangedRegion::CentralDirectoryOffsets,
            RepairChangedRegion::EndOfCentralDirectory,
        ]
    );
    assert!(preview.effects().member_payloads_preserved());
    assert!(preview.effects().reversible());

    let metadata = plan.to_json().unwrap();
    assert!(metadata.contains("\"intent\":\"non_destructive\""));
    assert!(metadata.contains("\"changed_members\":[\"mimetype\"]"));
    assert!(metadata.contains("\"member_payloads_preserved\":true"));
    assert!(metadata.contains("\"reversible\":true"));

    let patch = plan.apply().unwrap();
    assert_eq!(patch.source_fingerprint(), plan.source_fingerprint());
    assert_eq!(patch.target_fingerprint(), plan.output_fingerprint());
    assert_eq!(patch.target_bytes().len(), plan.output_len() as usize);

    let mut target = Vec::new();
    patch.write_to(&mut target).unwrap();
    assert_eq!(target, patch.target_bytes());
    let canonical_report = validate_package(&target).unwrap();
    assert!(canonical_report.issues().is_empty());
    let canonical_before = target.clone();
    assert!(matches!(
        plan_odf_repair(&target, &canonical_report, OdfRepairLimits::default()),
        Err(RepairError::ReportMismatch)
    ));
    assert_eq!(target, canonical_before);

    let mut restored = Vec::new();
    patch.inverse().write_to(&mut restored).unwrap();
    assert_eq!(restored, source_before);

    let mut stale_sink = Vec::new();
    let mut stale = target.clone();
    stale[0] ^= 1;
    assert!(matches!(
        patch.inverse().apply_to(&stale, &mut stale_sink),
        Err(RepairError::SourceChanged { .. })
    ));
    assert!(stale_sink.is_empty());
    assert_eq!(source, source_before);
}

#[test]
fn stale_or_foreign_reports_are_refused() {
    let source = remove_central_extra(package(0x5455, &[1, 0, 0, 0, 0]));
    let foreign = remove_central_extra(package(0x5455, &[1, 1, 0, 0, 0]));
    let report = validate_package(&foreign).unwrap();
    assert!(matches!(
        plan_mimetype_local_extra(&source, &report, OdfRepairLimits::default()),
        Err(RepairError::ReportMismatch)
    ));
}

#[test]
fn unsupported_unknown_and_malformed_extras_are_refused() {
    let source = remove_central_extra(package(0x1234, &[1, 2]));
    let report = validate_package(&source).unwrap();
    assert!(plan_mimetype_local_extra(&source, &report, OdfRepairLimits::default()).is_err());
    let mut malformed = remove_central_extra(package(0x5455, &[1, 0, 0, 0, 0]));
    malformed[42] = 8;
    let report = validate_package(&malformed).unwrap();
    assert!(plan_mimetype_local_extra(&malformed, &report, OdfRepairLimits::default()).is_err());
}

#[test]
fn central_extra_and_wrong_member_order_are_not_repair_targets() {
    let central_extra = package(0x5455, &[1, 0, 0, 0, 0]);
    let central_report = validate_package(&central_extra).unwrap();
    assert!(
        plan_mimetype_local_extra(&central_extra, &central_report, OdfRepairLimits::default())
            .is_err()
    );

    let wrong_order = package_with_mimetype_last();
    let wrong_order_report = validate_package(&wrong_order).unwrap();
    assert!(
        plan_mimetype_local_extra(
            &wrong_order,
            &wrong_order_report,
            OdfRepairLimits::default()
        )
        .is_err()
    );
}

#[test]
fn local_and_central_padding_are_refused_before_candidate_allocation() {
    let source = remove_central_extra(package(0x5455, &[1, 0, 0, 0, 0]));

    let local_padding = insert_local_padding(source.clone());
    let report = validate_package(&local_padding).unwrap();
    assert!(
        plan_mimetype_local_extra(&local_padding, &report, OdfRepairLimits::default()).is_err()
    );

    let central_padding = insert_central_padding(source);
    let report = validate_package(&central_padding).unwrap();
    assert!(
        plan_mimetype_local_extra(&central_padding, &report, OdfRepairLimits::default()).is_err()
    );
}

#[test]
fn untouched_data_descriptor_span_is_raw_preserved() {
    let source =
        add_descriptor_to_second_member(remove_central_extra(package(0x5455, &[1, 0, 0, 0, 0])));
    let report = validate_package(&source).unwrap();
    let plan = plan_mimetype_local_extra(&source, &report, OdfRepairLimits::default()).unwrap();
    let mut output = Vec::new();
    plan.write_to(&mut output).unwrap();
    assert_eq!(validate_package(&output).unwrap().issues().len(), 0);
}

#[test]
fn directory_entries_are_preserved_while_digest_count_excludes_them() {
    let source = remove_central_extra(package_with_directory(0x5455, &[1, 0, 0, 0, 0]));
    let source_before = source.clone();
    let report = validate_package(&source).unwrap();
    let plan = plan_mimetype_local_extra(&source, &report, OdfRepairLimits::default()).unwrap();
    assert_eq!(plan.member_count(), 4);
    let mut output = Vec::new();
    plan.write_to(&mut output).unwrap();
    assert_eq!(source, source_before);
    assert_eq!(validate_package(&output).unwrap().issues().len(), 0);
    let output_archive = soapberry_zip::ZipArchive::from_slice(&output).unwrap();
    assert_eq!(
        output_archive
            .entries()
            .map(|entry| entry.unwrap())
            .filter(|entry| entry.is_dir())
            .count(),
        1
    );
}

#[test]
fn package_wide_xml_security_scan_blocks_dtd_external_and_plus_xml_members() {
    let dtd = b"<!DOCTYPE root [<!ENTITY secret SYSTEM 'file:///secret'>]><root/>";
    let dtd_source = package_with_xml_member("styles.xml", "text/xml", dtd);
    let dtd_report = validate_package(&dtd_source).unwrap();
    assert_eq!(dtd_report.issues().len(), 1);
    assert!(matches!(
        plan_mimetype_local_extra(&dtd_source, &dtd_report, OdfRepairLimits::default()),
        Err(RepairError::Unsupported {
            reason: "XML member contains a DTD or entity declaration"
        })
    ));

    let external = b"<root href=\"https://example.invalid/out\"/>";
    let external_source = package_with_xml_member("metadata.rdf", "application/rdf+xml", external);
    let external_report = validate_package(&external_source).unwrap();
    assert_eq!(external_report.issues().len(), 1);
    assert!(matches!(
        plan_mimetype_local_extra(
            &external_source,
            &external_report,
            OdfRepairLimits::default()
        ),
        Err(RepairError::Unsupported {
            reason: "XML member contains an external or unsafe link"
        })
    ));

    let malformed = package_with_xml_member("settings.bin", "application/custom+xml", b"<root>");
    let malformed_report = validate_package(&malformed).unwrap();
    assert_eq!(malformed_report.issues().len(), 1);
    assert!(matches!(
        plan_mimetype_local_extra(&malformed, &malformed_report, OdfRepairLimits::default()),
        Err(RepairError::Unsupported {
            reason: "XML member is malformed during package-wide security scan"
        })
    ));
}

#[test]
fn package_wide_xml_security_scan_accepts_safe_references_and_uppercase_xml_type() {
    let source = package_with_xml_member(
        "settings.bin",
        "application/custom+XML",
        br#"<?xml version="1.0"?><root href="Pictures/a.png">safe &amp; &#x41;</root>"#,
    );
    let report = validate_package(&source).unwrap();
    let plan = plan_mimetype_local_extra(&source, &report, OdfRepairLimits::default()).unwrap();
    let mut output = Vec::new();
    plan.write_to(&mut output).unwrap();
    assert!(validate_package(&output).unwrap().issues().is_empty());
}

#[test]
fn package_wide_xml_security_scan_rejects_base_rdf_pi_and_invalid_characters() {
    for xml in [
        br#"<root xml:base="https://example.invalid/"><x href="Pictures/a.png"/></root>"#.as_slice(),
        br#"<rdf:RDF xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#"><rdf:Description rdf:resource="https://example.invalid/"/></rdf:RDF>"#.as_slice(),
        br#"<?xml-stylesheet href="https://example.invalid/style.xsl"?><root/>"#.as_slice(),
        b"<root>\xff</root>".as_slice(),
        b"<root>\x01</root>".as_slice(),
        b"<!--a--b--><root/>".as_slice(),
        br#"<?xml version="1.0" version="1.0"?><root/>"#.as_slice(),
        br#"<?xml version="1.0" foo="bar"?><root/>"#.as_slice(),
        b" \n<?xml version=\"1.0\"?><root/>".as_slice(),
        b"<1root/>".as_slice(),
        b"<x:root/>".as_slice(),
        b"<xmlns:root/>".as_slice(),
        br#"<root xmlns:xml="https://example.invalid/"/>"#.as_slice(),
        br#"<root xmlns:xmlns="urn:bad"/>"#.as_slice(),
        br#"<root xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:schemaLocation="urn:test https://example.invalid/schema.xsd"/>"#.as_slice(),
        br#"<root xmlns:a="urn:u" xmlns:b="urn:u" a:x="1" b:x="2"/>"#.as_slice(),
        br#"<root xmlns:a="urn:u" xmlns:b="urn:&#x75;" a:x="1" b:x="2"/>"#.as_slice(),
        br#"<office:dde-source xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"/>"#.as_slice(),
        br#"<text:dde-connection xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"/>"#.as_slice(),
        br#"<table:dde-link xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"/>"#.as_slice(),
    ] {
        let source = package_with_xml_member("styles.xml", "text/xml", xml);
        let report = validate_package(&source).unwrap();
        assert!(plan_mimetype_local_extra(&source, &report, OdfRepairLimits::default()).is_err());
    }

    let dde = package_with_xml_member(
        "settings.xml",
        "text/xml",
        br#"<root office:dde-topic="external" xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"/>"#,
    );
    let report = validate_package(&dde).unwrap();
    assert!(plan_mimetype_local_extra(&dde, &report, OdfRepairLimits::default()).is_err());
}

struct ShortSink {
    bytes: Vec<u8>,
    max_write: usize,
}

struct ZeroSink;

impl Write for ZeroSink {
    fn write(&mut self, _bytes: &[u8]) -> io::Result<usize> {
        Ok(0)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

struct InterruptedSink {
    interrupted: bool,
    bytes: Vec<u8>,
}

impl Write for InterruptedSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if !self.interrupted {
            self.interrupted = true;
            return Err(io::Error::from(io::ErrorKind::Interrupted));
        }
        self.bytes.extend_from_slice(bytes);
        Ok(bytes.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

struct FlushFailureSink {
    bytes: Vec<u8>,
}

struct OverreportSink;

impl Write for OverreportSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        Ok(bytes.len().saturating_add(1))
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

impl Write for FlushFailureSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        self.bytes.extend_from_slice(bytes);
        Ok(bytes.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Err(io::Error::other("flush failed"))
    }
}

impl Write for ShortSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        let count = bytes.len().min(self.max_write);
        self.bytes.extend_from_slice(&bytes[..count]);
        Ok(count)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[test]
fn short_writes_are_retried_without_losing_progress() {
    let source = remove_central_extra(package(0x5455, &[1, 0, 0, 0, 0]));
    let report = validate_package(&source).unwrap();
    let plan = plan_mimetype_local_extra(&source, &report, OdfRepairLimits::default()).unwrap();
    let mut sink = ShortSink {
        bytes: Vec::new(),
        max_write: 3,
    };
    plan.write_to(&mut sink).unwrap();
    assert_eq!(sink.bytes.len(), plan.output_len() as usize);
}

#[test]
fn zero_sink_reports_typed_uncommitted_output() {
    let source = remove_central_extra(package(0x5455, &[1, 0, 0, 0, 0]));
    let report = validate_package(&source).unwrap();
    let plan = plan_mimetype_local_extra(&source, &report, OdfRepairLimits::default()).unwrap();
    let mut sink = ZeroSink;
    let result = plan.write_to(&mut sink);
    assert!(matches!(
        result,
        Err(RepairError::Io(error)) if error.kind() == io::ErrorKind::WriteZero
    ));
}

#[test]
fn interrupted_writes_are_retried_and_flush_failure_reports_complete_prefix() {
    let source = remove_central_extra(package(0x5455, &[1, 0, 0, 0, 0]));
    let report = validate_package(&source).unwrap();
    let plan = plan_mimetype_local_extra(&source, &report, OdfRepairLimits::default()).unwrap();
    let mut interrupted = InterruptedSink {
        interrupted: false,
        bytes: Vec::new(),
    };
    plan.write_to(&mut interrupted).unwrap();
    assert_eq!(interrupted.bytes.len(), plan.output_len() as usize);

    let mut flush_failure = FlushFailureSink { bytes: Vec::new() };
    assert!(matches!(
        plan.write_to(&mut flush_failure),
        Err(RepairError::IncompleteOutput {
            progress: RepairOutputProgress::CompleteUnflushed { .. },
            ..
        })
    ));
    assert_eq!(flush_failure.bytes.len(), plan.output_len() as usize);

    let mut overreport = OverreportSink;
    assert!(matches!(
        plan.write_to(&mut overreport),
        Err(RepairError::IncompleteOutput {
            progress: RepairOutputProgress::Indeterminate { .. },
            source,
        }) if matches!(*source, RepairError::SinkOverreported)
    ));
}

#[test]
fn plan_limits_are_enforced_before_publication() {
    let source = remove_central_extra(package(0x5455, &[1, 0, 0, 0, 0]));
    let report = validate_package(&source).unwrap();
    assert!(matches!(
        plan_mimetype_local_extra(
            &source,
            &report,
            OdfRepairLimits::default().with_max_preflight_candidate_bytes(0)
        ),
        Err(RepairError::InvalidLimits)
    ));
    let limits = OdfRepairLimits::new(
        source.len() as u64 - 1,
        source.len() as u64,
        100,
        1 << 20,
        1 << 20,
        4096,
        1024,
        64 * 1024,
    )
    .unwrap();
    assert!(matches!(
        plan_mimetype_local_extra(&source, &report, limits),
        Err(RepairError::Limit {
            resource: "input bytes",
            ..
        })
    ));

    let json_limits = OdfRepairLimits::new(
        source.len() as u64,
        source.len() as u64,
        100,
        1 << 20,
        1 << 20,
        1,
        1024,
        64 * 1024,
    )
    .unwrap();
    assert!(matches!(
        plan_mimetype_local_extra(&source, &report, json_limits),
        Err(RepairError::Limit {
            resource: "plan JSON bytes",
            ..
        })
    ));
}

#[test]
fn raw_central_preflight_rejects_metadata_and_record_limits_before_indexing() {
    let source = remove_central_extra(package(0x5455, &[1, 0, 0, 0, 0]));
    let report = validate_package(&source).unwrap();
    let metadata_limits = OdfRepairLimits::default().with_max_metadata_bytes(512);
    assert!(matches!(
        plan_mimetype_local_extra(&source, &report, metadata_limits),
        Err(RepairError::Limit {
            resource: "repair metadata bytes",
            ..
        })
    ));

    let directory_source = remove_central_extra(package_with_directory(0x5455, &[1, 0, 0, 0, 0]));
    let directory_report = validate_package(&directory_source).unwrap();
    let member_limits = OdfRepairLimits::new(
        directory_source.len() as u64,
        directory_source.len() as u64,
        3,
        1 << 20,
        1 << 20,
        4096,
        1024,
        64 * 1024,
    )
    .unwrap();
    assert!(matches!(
        plan_mimetype_local_extra(&directory_source, &directory_report, member_limits),
        Err(RepairError::Limit {
            resource: "member count",
            ..
        })
    ));
}

#[test]
fn aggregate_metadata_budget_accepts_exact_and_rejects_one_over() {
    let source = remove_central_extra(package(0x5455, &[1, 0, 0, 0, 0]));
    let report = validate_package(&source).unwrap();
    let raw_name_bytes = b"mimetype".len() + b"META-INF/manifest.xml".len() + b"content.xml".len();
    // One fixed estimate for each retained catalog/layout entry, plus the
    // copied-name accounting and complete raw central records.
    let central_record_bytes = 3_u64 * 46 + raw_name_bytes as u64;
    let exact = 3_u64 * (128 + 64 + 128 + 16)
        + raw_name_bytes as u64
        + central_record_bytes
        + package_manifest_metadata_bytes();
    let exact_limits = OdfRepairLimits::default().with_max_metadata_bytes(exact);
    assert!(plan_mimetype_local_extra(&source, &report, exact_limits).is_ok());

    let one_over_limits = OdfRepairLimits::default().with_max_metadata_bytes(exact - 1);
    assert!(matches!(
        plan_mimetype_local_extra(&source, &report, one_over_limits),
        Err(RepairError::Limit {
            resource: "repair metadata bytes",
            observed,
            limit,
        }) if observed == exact && limit + 1 == exact
    ));
}

#[test]
fn decoded_manifest_metadata_is_charged_at_exact_and_one_over_boundaries() {
    let source = remove_central_extra(package(0x5455, &[1, 0, 0, 0, 0]));
    let report = validate_package(&source).unwrap();
    let base = package_metadata_base(3, &["mimetype", "META-INF/manifest.xml", "content.xml"]);
    let exact = base + package_manifest_metadata_bytes();
    assert!(
        plan_mimetype_local_extra(
            &source,
            &report,
            OdfRepairLimits::default().with_max_metadata_bytes(exact),
        )
        .is_ok()
    );
    assert!(matches!(
        plan_mimetype_local_extra(
            &source,
            &report,
            OdfRepairLimits::default().with_max_metadata_bytes(exact - 1),
        ),
        Err(RepairError::Limit {
            resource: "repair metadata bytes",
            observed,
            limit,
        }) if observed == exact && limit == exact - 1
    ));
}

#[test]
fn output_progress_type_remains_public_for_sink_failures() {
    let _ = RepairOutputProgress::Untouched;
    let _ = RepairOutputProgress::CompleteUnverified { bytes: 0 };
}
