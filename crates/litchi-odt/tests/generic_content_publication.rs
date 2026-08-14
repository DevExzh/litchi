use std::{collections::BTreeMap, io::Write};

use flate2::{Compression, write::DeflateEncoder};
use litchi_odf_common::constants;
use litchi_odt::core::OwnedPackage as CoreOwnedPackage;
use litchi_odt::odc::{ChartClass, Definition, Document, Text, serialize_content};
use soapberry_zip::{
    CompressionMethod, Header, PreservationIndex, ZipArchive, ZipArchiveWriter,
    extra_fields::ExtraFieldId,
};

#[derive(Debug, PartialEq, Eq)]
struct RawMember {
    local: Vec<u8>,
    central_without_offset: Vec<u8>,
}

fn chart_definition(title: &str) -> Definition {
    let mut definition = Definition::new(ChartClass::bar());
    definition.title = Some(Text::new(title));
    definition
}

fn write_deflated(writer: &mut ZipArchiveWriter<Vec<u8>>, path: &str, bytes: &[u8]) {
    let (mut file, config) = writer
        .new_file(path)
        .compression_method(CompressionMethod::Deflate)
        .start()
        .unwrap();
    let encoder = DeflateEncoder::new(&mut file, Compression::default());
    let mut data_writer = config.wrap(encoder);
    data_writer.write_all(bytes).unwrap();
    let (encoder, descriptor) = data_writer.finish().unwrap();
    encoder.finish().unwrap();
    file.finish(descriptor).unwrap();
}

fn write_deflated_with_metadata(writer: &mut ZipArchiveWriter<Vec<u8>>, path: &str, bytes: &[u8]) {
    let (mut file, config) = writer
        .new_file(path)
        .compression_method(CompressionMethod::Deflate)
        .extra_field(ExtraFieldId::new(0xaaaa), b"local", Header::LOCAL)
        .unwrap()
        .extra_field(ExtraFieldId::new(0xbbbb), b"central", Header::CENTRAL)
        .unwrap()
        .start()
        .unwrap();
    let encoder = DeflateEncoder::new(&mut file, Compression::default());
    let mut data_writer = config.wrap(encoder);
    data_writer.write_all(bytes).unwrap();
    let (encoder, descriptor) = data_writer.finish().unwrap();
    encoder.finish().unwrap();
    file.finish(descriptor).unwrap();
}

fn manifest(
    content_size: bool,
    signature: bool,
    encrypted_path: bool,
    root_mimetype: &str,
    content_alias: Option<&str>,
) -> String {
    let alias = content_alias.map_or_else(String::new, |path| {
        format!(
            r#"<manifest:file-entry manifest:full-path="{path}" manifest:media-type="text/xml" manifest:size="999"/>"#
        )
    });
    let mut value = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3"><manifest:file-entry manifest:full-path="/" manifest:media-type="{mime}"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"{size}/>{alias}<manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="settings.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="Pictures/stored.bin" manifest:media-type="application/octet-stream"/><manifest:file-entry manifest:full-path="Pictures/deflated.bin" manifest:media-type="application/octet-stream"/><manifest:file-entry manifest:full-path="Vendor/unknown.bin" manifest:media-type="application/octet-stream""#,
        mime = root_mimetype,
        size = if content_size {
            r#" manifest:size="1""#
        } else {
            ""
        },
        alias = alias,
    );
    if encrypted_path {
        value.push_str(
            r#" manifest:size="16"><manifest:encryption-data><manifest:algorithm manifest:algorithm-name="http://www.w3.org/2009/xmlenc11#aes256-gcm" manifest:initialisation-vector="AAAAAAAAAAAAAAAA"/><manifest:start-key-generation manifest:start-key-generation-name="SHA1" manifest:key-size="20"/><manifest:key-derivation manifest:key-derivation-name="PBKDF2" manifest:salt="AQ==" manifest:iteration-count="1000" manifest:key-size="32"/></manifest:encryption-data></manifest:file-entry>"#,
        );
    } else {
        value.push_str("/>");
    }
    if signature {
        value.push_str(
            r#"<manifest:file-entry manifest:full-path="META-INF/documentsignatures.xml" manifest:media-type="application/vnd.oasis.opendocument.digital-signature+xml"/>"#,
        );
    }
    value.push_str(
        r#"<manifest:file-entry manifest:full-path="META-INF/manifest.xml" manifest:media-type="text/xml"/></manifest:manifest>"#,
    );
    value
}

fn source_package(
    definition: &Definition,
    content_size: bool,
    signature: bool,
    encrypted_path: bool,
) -> Vec<u8> {
    source_package_with_options(
        definition,
        content_size,
        signature,
        encrypted_path,
        constants::ODF_CHART,
        false,
        None,
    )
}

fn write_stored_with_local_metadata(
    writer: &mut ZipArchiveWriter<Vec<u8>>,
    path: &str,
    bytes: &[u8],
) {
    let (mut file, config) = writer
        .new_file(path)
        .compression_method(CompressionMethod::Store)
        .extra_field(ExtraFieldId::new(0xcccc), b"noncanonical", Header::LOCAL)
        .unwrap()
        .start()
        .unwrap();
    let mut data_writer = config.wrap(&mut file);
    data_writer.write_all(bytes).unwrap();
    let (_, descriptor) = data_writer.finish().unwrap();
    file.finish(descriptor).unwrap();
}

fn source_package_with_options(
    definition: &Definition,
    content_size: bool,
    signature: bool,
    encrypted_path: bool,
    root_mimetype: &str,
    noncanonical_mimetype: bool,
    content_alias: Option<&str>,
) -> Vec<u8> {
    let content = serialize_content(definition).unwrap();
    let mut writer = ZipArchiveWriter::new(Vec::new());
    if noncanonical_mimetype {
        write_stored_with_local_metadata(&mut writer, "mimetype", root_mimetype.as_bytes());
    } else {
        writer
            .write_stored_file("mimetype", constants::ODF_CHART.as_bytes())
            .unwrap();
    }
    write_deflated(&mut writer, "content.xml", content.as_bytes());
    writer
        .write_stored_file("styles.xml", b"<styles-preserved/>")
        .unwrap();
    writer
        .write_stored_file("meta.xml", b"<meta-preserved/>")
        .unwrap();
    writer
        .write_stored_file("settings.xml", b"<settings-preserved/>")
        .unwrap();
    writer
        .write_stored_file("Pictures/stored.bin", b"stored media")
        .unwrap();
    write_deflated_with_metadata(&mut writer, "Pictures/deflated.bin", b"deflated media");
    writer
        .write_stored_file("Vendor/unknown.bin", b"unknown bytes")
        .unwrap();
    if signature {
        writer
            .write_stored_file("META-INF/documentsignatures.xml", b"<signatures/>")
            .unwrap();
    }
    writer
        .write_stored_file(
            "META-INF/manifest.xml",
            manifest(
                content_size,
                signature,
                encrypted_path,
                root_mimetype,
                content_alias,
            )
            .as_bytes(),
        )
        .unwrap();
    writer.finish().unwrap()
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

fn member_order(bytes: &[u8]) -> Vec<String> {
    let archive = ZipArchive::from_slice(bytes).unwrap().into_zip_archive();
    let mut buffer = vec![0_u8; soapberry_zip::RECOMMENDED_BUFFER_SIZE];
    let mut entries = archive.entries(&mut buffer);
    let mut names = Vec::new();
    while let Some(entry) = entries.next_entry().unwrap() {
        names.push(
            entry
                .file_path()
                .try_normalize()
                .unwrap()
                .as_ref()
                .to_owned(),
        );
    }
    names
}

fn local_member_order(bytes: &[u8]) -> Vec<String> {
    let archive = ZipArchive::from_slice(bytes).unwrap().into_zip_archive();
    let mut buffer = vec![0_u8; soapberry_zip::RECOMMENDED_BUFFER_SIZE];
    let index = PreservationIndex::new(&archive, &mut buffer).unwrap();
    let mut records = archive.entries(&mut buffer);
    let mut names = index
        .entries()
        .iter()
        .map(|preserved| {
            let record = records.next_entry().unwrap().unwrap();
            (
                preserved.local_span().start,
                record
                    .file_path()
                    .try_normalize()
                    .unwrap()
                    .as_ref()
                    .to_owned(),
            )
        })
        .collect::<Vec<_>>();
    assert!(records.next_entry().unwrap().is_none());
    names.sort_unstable_by_key(|(local_start, _)| *local_start);
    names.into_iter().map(|(_, name)| name).collect()
}

fn swap_central_records(bytes: &[u8], left_name: &str, right_name: &str) -> Vec<u8> {
    let archive = ZipArchive::from_slice(bytes).unwrap().into_zip_archive();
    let mut buffer = vec![0_u8; soapberry_zip::RECOMMENDED_BUFFER_SIZE];
    let index = PreservationIndex::new(&archive, &mut buffer).unwrap();
    let mut records = archive.entries(&mut buffer);
    let mut central = Vec::new();
    for preserved in index.entries() {
        let record = records.next_entry().unwrap().unwrap();
        central.push((
            record
                .file_path()
                .try_normalize()
                .unwrap()
                .as_ref()
                .to_owned(),
            preserved.central_record(),
        ));
    }
    assert!(records.next_entry().unwrap().is_none());
    let left = central
        .iter()
        .position(|(name, _)| name == left_name)
        .unwrap();
    let right = central
        .iter()
        .position(|(name, _)| name == right_name)
        .unwrap();
    central.swap(left, right);
    let directory_start = archive.directory_offset() as usize;
    let directory_end = archive.eocd_offset() as usize;
    let mut output = bytes.to_vec();
    let mut cursor = directory_start;
    for (_, range) in central {
        let source_start = range.start as usize;
        let source_end = range.end as usize;
        let record = &bytes[source_start..source_end];
        output[cursor..cursor + record.len()].copy_from_slice(record);
        cursor += record.len();
    }
    assert_eq!(cursor, directory_end);
    output
}

fn swap_local_records(bytes: &[u8], left_name: &str, right_name: &str) -> Vec<u8> {
    let archive = ZipArchive::from_slice(bytes).unwrap().into_zip_archive();
    let mut buffer = vec![0_u8; soapberry_zip::RECOMMENDED_BUFFER_SIZE];
    let index = PreservationIndex::new(&archive, &mut buffer).unwrap();
    let mut records = archive.entries(&mut buffer);
    let mut members = Vec::new();
    for preserved in index.entries() {
        let record = records.next_entry().unwrap().unwrap();
        members.push((
            record
                .file_path()
                .try_normalize()
                .unwrap()
                .as_ref()
                .to_owned(),
            preserved.local_span(),
            preserved.central_record(),
        ));
    }
    assert!(records.next_entry().unwrap().is_none());
    let left = members
        .iter()
        .position(|(name, _, _)| name == left_name)
        .unwrap();
    let right = members
        .iter()
        .position(|(name, _, _)| name == right_name)
        .unwrap();
    let mut local_order = (0..members.len()).collect::<Vec<_>>();
    local_order.sort_unstable_by_key(|index| members[*index].1.start);
    let left_position = local_order.iter().position(|index| *index == left).unwrap();
    let right_position = local_order
        .iter()
        .position(|index| *index == right)
        .unwrap();
    local_order.swap(left_position, right_position);

    let directory_start = archive.directory_offset() as usize;
    let directory_end = archive.eocd_offset() as usize;
    let mut local_bytes = Vec::with_capacity(directory_start);
    let mut new_offsets = vec![0_u32; members.len()];
    for member_index in local_order {
        let local = &members[member_index].1;
        let local_start = local_bytes.len();
        local_bytes.extend_from_slice(&bytes[local.start as usize..local.end as usize]);
        new_offsets[member_index] = u32::try_from(local_start).unwrap();
    }
    assert_eq!(local_bytes.len(), directory_start);

    let mut output = local_bytes;
    for (member_index, (_, _, central)) in members.iter().enumerate() {
        let mut central_bytes = bytes[central.start as usize..central.end as usize].to_vec();
        central_bytes[42..46].copy_from_slice(&new_offsets[member_index].to_le_bytes());
        output.extend_from_slice(&central_bytes);
    }
    output.extend_from_slice(&bytes[directory_end..]);
    output
}

fn corrupt_deflated_descriptor(source: &[u8], field_offset: usize) -> Vec<u8> {
    let mut malformed = source.to_vec();
    let archive = ZipArchive::from_slice(source).unwrap().into_zip_archive();
    let mut buffer = vec![0_u8; soapberry_zip::RECOMMENDED_BUFFER_SIZE];
    let index = PreservationIndex::new(&archive, &mut buffer).unwrap();
    let mut records = archive.entries(&mut buffer);
    for preserved in index.entries() {
        let record = records.next_entry().unwrap().unwrap();
        if record.file_path().try_normalize().unwrap().as_ref() == "Pictures/deflated.bin" {
            let descriptor_start = preserved.local_span().end as usize - 16;
            assert_eq!(
                &malformed[descriptor_start..descriptor_start + 4],
                b"PK\x07\x08"
            );
            malformed[descriptor_start + field_offset] ^= 1;
            return malformed;
        }
    }
    panic!("deflated test member was not found");
}

fn corrupt_deflated_payload(source: &[u8]) -> Vec<u8> {
    let mut malformed = source.to_vec();
    let archive = ZipArchive::from_slice(source).unwrap().into_zip_archive();
    let mut buffer = vec![0_u8; soapberry_zip::RECOMMENDED_BUFFER_SIZE];
    let index = PreservationIndex::new(&archive, &mut buffer).unwrap();
    let mut records = archive.entries(&mut buffer);
    for preserved in index.entries() {
        let record = records.next_entry().unwrap().unwrap();
        if record.file_path().try_normalize().unwrap().as_ref() == "Pictures/deflated.bin" {
            let local = preserved.local_span().start as usize;
            let name_len =
                u16::from_le_bytes([malformed[local + 26], malformed[local + 27]]) as usize;
            let extra_len =
                u16::from_le_bytes([malformed[local + 28], malformed[local + 29]]) as usize;
            let payload = local + 30 + name_len + extra_len;
            let compressed_len = usize::try_from(record.compressed_size_hint()).unwrap();
            malformed[payload + compressed_len - 1] ^= 1;
            return malformed;
        }
    }
    panic!("deflated test member was not found");
}

fn assert_raw_equal_except_content(source: &[u8], output: &[u8]) {
    let source_members = raw_members(source);
    let output_members = raw_members(output);
    assert_eq!(member_order(source), member_order(output));
    assert_eq!(local_member_order(source), local_member_order(output));
    assert_eq!(
        source_members.keys().collect::<Vec<_>>(),
        output_members.keys().collect::<Vec<_>>()
    );
    for (name, source_member) in source_members {
        if name == "content.xml" {
            assert_ne!(output_members[&name], source_member);
        } else {
            assert_eq!(output_members[&name], source_member, "{name}");
        }
    }
}

#[test]
fn generic_content_publication_raw_preserves_store_and_deflate_members() {
    let source_definition = chart_definition("before");
    let source_bytes = swap_central_records(
        &source_package(&source_definition, false, false, false),
        "content.xml",
        "Vendor/unknown.bin",
    );
    let mut document = Document::from_bytes(source_bytes.clone()).unwrap();

    assert_ne!(
        member_order(&source_bytes),
        local_member_order(&source_bytes)
    );

    let mut replacement = source_definition.clone();
    replacement.title = Some(Text::new("after"));
    document.set_definition(&replacement).unwrap();
    let output = document.to_bytes();

    assert_raw_equal_except_content(&source_bytes, &output);
    assert_eq!(document.text(), "after");
    assert_eq!(
        Document::from_bytes(output.clone()).unwrap().text(),
        "after"
    );

    let mut no_op = Document::from_bytes(source_bytes.clone()).unwrap();
    no_op.set_definition(&source_definition).unwrap();
    assert_eq!(no_op.to_bytes(), source_bytes);
}

#[test]
fn generic_content_publication_keeps_signature_fallback_and_manifest_size_fallback() {
    let source_definition = chart_definition("before");
    let mut replacement = source_definition.clone();
    replacement.title = Some(Text::new("after"));

    let signed = source_package(&source_definition, false, true, false);
    let signed_original = signed.clone();
    let mut signed_document = Document::from_bytes(signed).unwrap();
    signed_document.set_definition(&replacement).unwrap();
    let signed_package = CoreOwnedPackage::from_bytes(signed_document.to_bytes()).unwrap();
    assert!(
        !signed_package
            .has_file("META-INF/documentsignatures.xml")
            .unwrap()
    );
    assert_eq!(signed_document.text(), "after");
    let mut signed_no_op = Document::from_bytes(signed_original.clone()).unwrap();
    signed_no_op.set_definition(&source_definition).unwrap();
    assert_eq!(signed_no_op.to_bytes(), signed_original);

    let sized = source_package(&source_definition, true, false, false);
    let mut sized_document = Document::from_bytes(sized).unwrap();
    sized_document.set_definition(&replacement).unwrap();
    let sized_package = CoreOwnedPackage::from_bytes(sized_document.to_bytes()).unwrap();
    let manifest =
        String::from_utf8(sized_package.get_file("META-INF/manifest.xml").unwrap()).unwrap();
    assert!(!manifest.contains(
        "manifest:full-path=\"content.xml\" manifest:media-type=\"text/xml\" manifest:size="
    ));
    assert_eq!(sized_document.text(), "after");
}

#[test]
fn generic_content_publication_refuses_encrypted_source_without_password_writer() {
    let source_definition = chart_definition("before");
    let mut replacement = source_definition.clone();
    replacement.title = Some(Text::new("after"));
    let encrypted = source_package(&source_definition, false, false, true);
    let encrypted_original = encrypted.clone();
    let mut no_op = Document::from_bytes(encrypted_original.clone()).unwrap();
    no_op.set_definition(&source_definition).unwrap();
    assert_eq!(no_op.to_bytes(), encrypted_original);
    let mut document = Document::from_bytes(encrypted).unwrap();

    let error = document.set_definition(&replacement).unwrap_err();
    assert!(error.to_string().contains("encrypted ODF entries"));
}

#[test]
fn generic_content_publication_falls_back_for_noncanonical_mimetype_and_root_mime() {
    let source_definition = chart_definition("before");
    let mut replacement = source_definition.clone();
    replacement.title = Some(Text::new("after"));

    let noncanonical = source_package_with_options(
        &source_definition,
        false,
        false,
        false,
        constants::ODF_CHART,
        true,
        None,
    );
    let mut noncanonical_document = Document::from_bytes(noncanonical.clone()).unwrap();
    noncanonical_document.set_definition(&replacement).unwrap();
    let noncanonical_output = noncanonical_document.to_bytes();
    assert_ne!(
        raw_members(&noncanonical)["mimetype"],
        raw_members(&noncanonical_output)["mimetype"]
    );
    assert_eq!(noncanonical_document.text(), "after");

    let wrong_central_order = swap_central_records(
        &source_package(&source_definition, false, false, false),
        "mimetype",
        "content.xml",
    );
    let mut wrong_central_document = Document::from_bytes(wrong_central_order.clone()).unwrap();
    wrong_central_document.set_definition(&replacement).unwrap();
    let wrong_central_output = wrong_central_document.to_bytes();
    assert_ne!(
        member_order(&wrong_central_order),
        member_order(&wrong_central_output)
    );

    let wrong_local_order = swap_local_records(
        &source_package(&source_definition, false, false, false),
        "mimetype",
        "content.xml",
    );
    let mut wrong_local_document = Document::from_bytes(wrong_local_order.clone()).unwrap();
    wrong_local_document.set_definition(&replacement).unwrap();
    let wrong_local_output = wrong_local_document.to_bytes();
    assert_ne!(
        local_member_order(&wrong_local_order),
        local_member_order(&wrong_local_output)
    );

    let mismatched_root = source_package_with_options(
        &source_definition,
        false,
        false,
        false,
        constants::ODF_TEXT,
        false,
        None,
    );
    let mut mismatched_document = Document::from_bytes(mismatched_root).unwrap();
    mismatched_document.set_definition(&replacement).unwrap();
    let output = mismatched_document.to_bytes();
    let reopened = CoreOwnedPackage::from_bytes(output).unwrap();
    assert_eq!(
        reopened.package().unwrap().manifest().mimetype,
        constants::ODF_CHART
    );

    for alias in [
        "/content.xml",
        "./content.xml",
        "../content.xml",
        r"\content.xml",
        "foo/../../content.xml",
    ] {
        let aliased = source_package_with_options(
            &source_definition,
            false,
            false,
            false,
            constants::ODF_CHART,
            false,
            Some(alias),
        );
        let mut aliased_document = Document::from_bytes(aliased).unwrap();
        aliased_document.set_definition(&replacement).unwrap();
        let output = aliased_document.to_bytes();
        let reopened = CoreOwnedPackage::from_bytes(output).unwrap();
        let manifest =
            String::from_utf8(reopened.get_file("META-INF/manifest.xml").unwrap()).unwrap();
        assert!(!manifest.contains(alias));
        assert!(!manifest.contains("manifest:size=\"999\""));
    }
}

#[test]
fn generic_content_publication_rejects_malformed_preserved_descriptor_crc_or_size() {
    let source_definition = chart_definition("before");
    let mut replacement = source_definition.clone();
    replacement.title = Some(Text::new("after"));
    let source = source_package(&source_definition, false, false, false);
    for field_offset in [4, 8, 12] {
        let malformed = corrupt_deflated_descriptor(&source, field_offset);
        let mut document = Document::from_bytes(malformed).unwrap();
        let result = document.set_definition(&replacement);
        if field_offset == 4 {
            let error = result.expect_err("descriptor CRC unexpectedly accepted");
            assert!(error.to_string().contains("deflated.bin"));
        } else {
            let output = result
                .map(|()| document.to_bytes())
                .unwrap_or_else(|error| panic!("descriptor size fallback failed: {error}"));
            assert_ne!(
                raw_members(&source)["Pictures/deflated.bin"],
                raw_members(&output)["Pictures/deflated.bin"],
                "descriptor size mutation incorrectly took the raw path for field {field_offset}"
            );
        }
    }
}

#[test]
fn generic_content_publication_verifies_changed_package_payloads() {
    let source_definition = chart_definition("before");
    let mut replacement = source_definition.clone();
    replacement.title = Some(Text::new("after"));
    let malformed =
        corrupt_deflated_payload(&source_package(&source_definition, false, false, false));
    assert!(
        CoreOwnedPackage::from_bytes(malformed.clone())
            .unwrap()
            .get_file("Pictures/deflated.bin")
            .is_err()
    );
    let mut document = Document::from_bytes(malformed).unwrap();
    let error = document.set_definition(&replacement).unwrap_err();
    assert!(error.to_string().contains("deflated.bin"));
}
