use std::collections::BTreeMap;
use std::fmt::Write as _;
use std::io::Cursor;

use base64::{Engine as _, engine::general_purpose::STANDARD as BASE64};
use litchi_core::patch::{BlobLimits, Patch as CorePatch, PatchLimits, Reversible};
use litchi_docx::Package;
use litchi_docx::ink::{
    AnchorGeometry, BaseProfile, BrushDraft, ContextDraft, ContextKind, Destination, Draft,
    EditLimits, FallbackImage, Geometry, HorizontalAlignment, HorizontalPosition,
    HorizontalRelativeFrom, Placement, Point, Prepared, Style, TraceDraft, VerticalAlignment,
    VerticalPosition, VerticalRelativeFrom,
};
use litchi_opc::{PackageWriter, authored_xml_requires_source_proof};
use serde_json::Value;
use sha2::{Digest as _, Sha256};
use soapberry_zip::office::{ArchiveReader, StreamingArchiveWriter};

use super::*;

const PNG: &[u8] = &[
    137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 0, 1, 0, 0, 0, 1, 8, 6, 0,
    0, 0, 31, 21, 196, 137, 0, 0, 0, 13, 73, 68, 65, 84, 120, 156, 99, 16, 80, 48, 248, 15, 0, 2,
    4, 1, 96, 141, 188, 187, 113, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130,
];

fn prepared() -> Prepared {
    Draft::default().finish().unwrap()
}

fn strokes() -> Prepared {
    let mut draft = Draft::default()
        .context(
            ContextDraft::new(ContextKind::InkDrawing)
                .with_xml_id("ctx")
                .unwrap(),
        )
        .unwrap()
        .brush(BrushDraft::new("brush").unwrap())
        .unwrap();
    for data in ["0 0, 10 20, -10 30", "5 6, 7 8"] {
        draft = draft
            .trace(
                TraceDraft::new(data)
                    .unwrap()
                    .with_context_ref("#ctx")
                    .unwrap()
                    .with_brush_ref("#brush")
                    .unwrap(),
            )
            .unwrap();
    }
    draft.finish().unwrap()
}

fn one_trace(data: &str) -> Prepared {
    Draft::default()
        .context(
            ContextDraft::new(ContextKind::InkDrawing)
                .with_xml_id("ctx")
                .unwrap(),
        )
        .unwrap()
        .brush(BrushDraft::new("brush").unwrap())
        .unwrap()
        .trace(
            TraceDraft::new(data)
                .unwrap()
                .with_context_ref("#ctx")
                .unwrap()
                .with_brush_ref("#brush")
                .unwrap(),
        )
        .unwrap()
        .finish()
        .unwrap()
}

fn image() -> FallbackImage {
    FallbackImage::from_bytes(PNG.to_vec()).unwrap()
}

fn png_with_ancillary_chunk() -> Vec<u8> {
    let mut chunk_data = b"Comment\0".to_vec();
    chunk_data.extend(std::iter::repeat_n(b'x', 4096));
    let mut chunk = Vec::with_capacity(4 + 4 + chunk_data.len() + 4);
    chunk.extend_from_slice(&(chunk_data.len() as u32).to_be_bytes());
    chunk.extend_from_slice(b"tEXt");
    chunk.extend_from_slice(&chunk_data);
    let mut crc_input = Vec::with_capacity(4 + chunk_data.len());
    crc_input.extend_from_slice(b"tEXt");
    crc_input.extend_from_slice(&chunk_data);
    chunk.extend_from_slice(&soapberry_zip::crc32(&crc_input).to_be_bytes());

    let iend = PNG.len() - 12;
    let mut result = Vec::with_capacity(PNG.len() + chunk.len());
    result.extend_from_slice(&PNG[..iend]);
    result.extend_from_slice(&chunk);
    result.extend_from_slice(&PNG[iend..]);
    result
}

fn drawing_source(payload: &Prepared, fallback: &[u8]) -> Vec<u8> {
    let mut opc = source(false, false);
    let main = format!(
        r#"<w:document xmlns:w="{W}" xmlns:r="{R}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:wi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml" mc:Ignorable="wi w14"><w:body><w:p><w:r><mc:AlternateContent><mc:Choice Requires="wi"><w:drawing><wp:inline><wp:extent cx="127000" cy="127000"/><wp:docPr id="1" name="Ink"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingInk"><w14:contentPart r:id="ink"/></a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice><mc:Fallback><w:pict><v:shape id="fallback" style="width:10pt;height:10pt"><v:imagedata r:id="fallbackImage"/></v:shape></w:pict></mc:Fallback></mc:AlternateContent></w:r></w:p></w:body></w:document>"#
    );
    opc.get_part_mut(&PackURI::new("/word/document.xml").unwrap())
        .unwrap()
        .set_blob(main.as_bytes().to_vec());
    opc.get_part_mut(&PackURI::new("/payload/handwriting.xml").unwrap())
        .unwrap()
        .set_blob(payload.as_bytes().to_vec());
    part(&mut opc, "/payload/fallback.png", "image/png", fallback);
    edge(
        &mut opc,
        "/word/document.xml",
        "fallbackImage",
        "../payload/fallback.png",
        rt::IMAGE,
    );
    PackageWriter::to_bytes(&opc).unwrap()
}

fn durable_limits() -> PatchLimits {
    PatchLimits::new(
        BlobLimits::new(16, 64 * 1024 * 1024, 128 * 1024 * 1024),
        32 * 1024 * 1024,
        64,
        32,
        1024 * 1024,
        64 * 1024 * 1024,
    )
}

fn members(bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
    let archive = ArchiveReader::new(bytes).unwrap();
    archive
        .file_names()
        .map(|name| (name.to_owned(), archive.read(name).unwrap()))
        .collect()
}

fn rewrite_member<F>(bytes: &[u8], selected: &str, mut rewrite: F) -> Vec<u8>
where
    F: FnMut(&[u8]) -> Vec<u8>,
{
    let archive = ArchiveReader::new(bytes).unwrap();
    let mut writer = StreamingArchiveWriter::new();
    for name in archive.file_names() {
        let content = archive.read(name).unwrap();
        if name == selected {
            writer.write_stored(name, &rewrite(&content)).unwrap();
        } else {
            writer.write_stored(name, &content).unwrap();
        }
    }
    writer.finish_to_bytes().unwrap()
}

fn add_member(bytes: &[u8], name: &str, content: &[u8]) -> Vec<u8> {
    let archive = ArchiveReader::new(bytes).unwrap();
    let mut writer = StreamingArchiveWriter::new();
    for member in archive.file_names() {
        writer
            .write_stored(member, &archive.read(member).unwrap())
            .unwrap();
    }
    writer.write_stored(name, content).unwrap();
    writer.finish_to_bytes().unwrap()
}

fn remove_member(bytes: &[u8], selected: &str) -> Vec<u8> {
    let archive = ArchiveReader::new(bytes).unwrap();
    let mut writer = StreamingArchiveWriter::new();
    for member in archive.file_names() {
        if member != selected {
            writer
                .write_stored(member, &archive.read(member).unwrap())
                .unwrap();
        }
    }
    writer.finish_to_bytes().unwrap()
}

fn lexical_comment(content: &[u8]) -> Vec<u8> {
    let closing = if content
        .windows(b"</Types>".len())
        .any(|window| window == b"</Types>")
    {
        b"</Types>".as_slice()
    } else {
        b"</Relationships>".as_slice()
    };
    let index = content
        .windows(closing.len())
        .position(|window| window == closing)
        .unwrap();
    let mut result = Vec::with_capacity(content.len() + 32);
    result.extend_from_slice(&content[..index]);
    result.extend_from_slice(b"\n<!-- durable lexical stale -->\n");
    result.extend_from_slice(&content[index..]);
    result
}

fn xml_attr(tag: &str, name: &str) -> String {
    let double = format!("{name}=\"");
    let single = format!(r#"{name}='"#);
    let (start, quote) = if let Some(start) = tag.find(&double) {
        (start + double.len(), b'\"')
    } else if let Some(start) = tag.find(&single) {
        (start + single.len(), b'\'')
    } else {
        panic!("missing {name} in {tag}");
    };
    let end = tag[start..]
        .bytes()
        .position(|byte| byte == quote)
        .map(|offset| start + offset)
        .unwrap();
    tag[start..end].to_owned()
}

fn xml_declaration(xml: &str) -> String {
    let end = xml.find("?>").map(|index| index + 2).unwrap_or(0);
    xml[..end].replace('"', "'")
}

fn lexical_content_types(content: &[u8]) -> Vec<u8> {
    let xml = String::from_utf8(content.to_vec()).unwrap();
    let root_start = xml.find("<Types ").unwrap();
    let root_end = xml[root_start..]
        .find('>')
        .map(|index| root_start + index)
        .unwrap();
    let close = xml.rfind("</Types>").unwrap();
    let mut output = String::new();
    output.push_str(&xml_declaration(&xml));
    output.push_str(
        "\n<ct:Types xmlns:ct='http://schemas.openxmlformats.org/package/2006/content-types'>\n",
    );
    output.push_str("<!-- producer header -->\n");
    let body = &xml[root_end + 1..close];
    let mut cursor = 0;
    while let Some(relative) = body[cursor..]
        .find("<Default ")
        .into_iter()
        .chain(body[cursor..].find("<Override "))
        .min()
    {
        let start = cursor + relative;
        let end = body[start..]
            .find("/>")
            .map(|index| start + index + 2)
            .unwrap();
        let tag = &body[start..end];
        if tag.starts_with("<Default ") {
            let extension = xml_attr(tag, "Extension");
            let content_type = xml_attr(tag, "ContentType");
            writeln!(
                output,
                "<!-- interleaved default -->\n<ct:Default ContentType='{content_type}' Extension='{extension}'/>"
            )
            .unwrap();
        } else {
            let part_name = xml_attr(tag, "PartName");
            let content_type = xml_attr(tag, "ContentType");
            writeln!(
                output,
                "<!-- interleaved override -->\n<ct:Override ContentType='{content_type}' PartName='{part_name}'/>"
            )
            .unwrap();
        }
        cursor = end;
    }
    output.push_str("</ct:Types>");
    output.into_bytes()
}

fn lexical_relationships(content: &[u8]) -> Vec<u8> {
    let xml = String::from_utf8(content.to_vec()).unwrap();
    let root_start = xml.find("<Relationships ").unwrap();
    let root_end = xml[root_start..]
        .find('>')
        .map(|index| root_start + index)
        .unwrap();
    let close = xml.rfind("</Relationships>").unwrap();
    let mut output = String::new();
    output.push_str(&xml_declaration(&xml));
    output.push_str(
        "\n<pr:Relationships xmlns:pr='http://schemas.openxmlformats.org/package/2006/relationships'>\n",
    );
    output.push_str("<!-- relationship prelude -->\n");
    let body = &xml[root_end + 1..close];
    let mut cursor = 0;
    while let Some(relative) = body[cursor..].find("<Relationship ") {
        let start = cursor + relative;
        let end = body[start..]
            .find("/>")
            .map(|index| start + index + 2)
            .unwrap();
        let tag = &body[start..end];
        let id = xml_attr(tag, "Id");
        let rel_type = xml_attr(tag, "Type");
        let target = xml_attr(tag, "Target");
        let mode = tag.find("TargetMode=").map(|_| xml_attr(tag, "TargetMode"));
        write!(output, "<!-- before next owner -->\n<pr:Relationship Target='{target}' Type='{rel_type}' Id='{id}'").unwrap();
        if let Some(mode) = mode {
            write!(output, " TargetMode='{mode}'").unwrap();
        }
        output.push_str("/>\n");
        cursor = end;
    }
    output.push_str("</pr:Relationships>");
    output.into_bytes()
}

fn formatted_story_xml(content: &[u8]) -> Vec<u8> {
    let content = replace_once(
        content,
        b"<u:body>",
        b"<u:body>\n    <!-- retained story source -->\n    ",
    );
    replace_once(&content, b"</u:body>", b"\n  </u:body>")
}

fn noncanonical_shared_source() -> Vec<u8> {
    let original = PackageWriter::to_bytes(&source(false, true)).unwrap();
    let archive = ArchiveReader::new(&original).unwrap();
    let mut writer = StreamingArchiveWriter::new();
    for name in archive.file_names() {
        let content = archive.read(name).unwrap();
        let content = match name {
            "[Content_Types].xml" => lexical_content_types(&content),
            "_rels/.rels" | "word/_rels/document.xml.rels" => lexical_relationships(&content),
            "word/document.xml" => formatted_story_xml(&content),
            _ => content,
        };
        writer.write_stored(name, &content).unwrap();
    }
    writer.finish_to_bytes().unwrap()
}

fn source_with_empty_ink_relationship_member() -> Vec<u8> {
    let original = PackageWriter::to_bytes(&source(false, false)).unwrap();
    add_member(
        &original,
        "payload/_rels/handwriting.xml.rels",
        br#"<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<pr:Relationships xmlns:pr='http://schemas.openxmlformats.org/package/2006/relationships'>
<!-- explicitly empty Ink owner relationships -->
</pr:Relationships>"#,
    )
}

fn replace_once(bytes: &[u8], from: &[u8], to: &[u8]) -> Vec<u8> {
    let index = bytes
        .windows(from.len())
        .position(|window| window == from)
        .unwrap_or_else(|| panic!("missing tamper marker {:?}", from));
    let mut result = Vec::with_capacity(bytes.len() + to.len().saturating_sub(from.len()));
    result.extend_from_slice(&bytes[..index]);
    result.extend_from_slice(to);
    result.extend_from_slice(&bytes[index + from.len()..]);
    result
}

fn replace_last(bytes: &[u8], from: &[u8], to: &[u8]) -> Vec<u8> {
    let index = bytes
        .windows(from.len())
        .rposition(|window| window == from)
        .unwrap_or_else(|| panic!("missing tamper marker {:?}", from));
    let mut result = Vec::with_capacity(bytes.len() + to.len().saturating_sub(from.len()));
    result.extend_from_slice(&bytes[..index]);
    result.extend_from_slice(to);
    result.extend_from_slice(&bytes[index + from.len()..]);
    result
}

fn flip_first_blob_digest(bytes: &[u8]) -> Vec<u8> {
    let marker = b"\"sha256\":\"";
    let start = bytes
        .windows(marker.len())
        .position(|window| window == marker)
        .map(|index| index + marker.len())
        .unwrap();
    let mut result = bytes.to_vec();
    result[start] = if result[start] == b'0' { b'1' } else { b'0' };
    result
}

fn sha256_hex(bytes: &[u8]) -> String {
    let mut output = String::with_capacity(64);
    for byte in Sha256::digest(bytes) {
        write!(output, "{byte:02x}").unwrap();
    }
    output
}

fn wire_blob(wire: &[u8], collection: &str) -> (String, String, Vec<u8>) {
    let root: Value = serde_json::from_slice(wire).unwrap();
    let blob = root
        .get(collection)
        .and_then(Value::as_array)
        .and_then(|blobs| blobs.first())
        .and_then(Value::as_object)
        .unwrap();
    let encoded = blob
        .get("bytes")
        .and_then(Value::as_str)
        .unwrap()
        .to_owned();
    let digest = blob
        .get("sha256")
        .and_then(Value::as_str)
        .unwrap()
        .to_owned();
    let bytes = BASE64.decode(&encoded).unwrap();
    (encoded, digest, bytes)
}

fn wire_target_hash(wire: &[u8]) -> String {
    let root: Value = serde_json::from_slice(wire).unwrap();
    root.get("operations")
        .and_then(Value::as_array)
        .and_then(|operations| operations.first())
        .and_then(Value::as_object)
        .and_then(|operation| operation.get("forward"))
        .and_then(Value::as_object)
        .and_then(|operation| operation.get("preconditions"))
        .and_then(Value::as_object)
        .and_then(|preconditions| preconditions.get("target_sha256"))
        .and_then(Value::as_str)
        .unwrap()
        .to_owned()
}

fn replace_wire_blob(wire: &[u8], collection: &str, precondition: &str, bytes: &[u8]) -> Vec<u8> {
    let (old_encoded, old_digest, _) = wire_blob(wire, collection);
    let new_encoded = BASE64.encode(bytes);
    let new_digest = sha256_hex(bytes);
    let old_entry = format!(r#"{{"bytes":"{old_encoded}","sha256":"{old_digest}"}}"#);
    let new_entry = format!(r#"{{"bytes":"{new_encoded}","sha256":"{new_digest}"}}"#);
    let mut result = replace_once(wire, old_entry.as_bytes(), new_entry.as_bytes());
    let old_precondition = format!(r#""{precondition}":"{old_digest}""#);
    let new_precondition = format!(r#""{precondition}":"{new_digest}""#);
    result = replace_once(
        &result,
        old_precondition.as_bytes(),
        new_precondition.as_bytes(),
    );
    result
}

fn replace_wire_target(wire: &[u8], target: &str) -> Vec<u8> {
    let old_target = wire_target_hash(wire);
    let old_marker = format!(r#""target_sha256":"{old_target}""#);
    let new_marker = format!(r#""target_sha256":"{target}""#);
    replace_once(wire, old_marker.as_bytes(), new_marker.as_bytes())
}

fn mutate_restore_intent(restore: &[u8], from: &[u8], to: &[u8]) -> Vec<u8> {
    assert!(restore.starts_with(b"LIR1"));
    let intent_length = u64::from_le_bytes(restore[4..12].try_into().unwrap()) as usize;
    let intent_start = 12;
    let intent_end = intent_start + intent_length;
    let replacement = replace_once(&restore[intent_start..intent_end], from, to);
    let mut result = restore.to_vec();
    result[intent_start..intent_end].copy_from_slice(&replacement);
    result
}

fn restore_component_lengths(restore: &[u8]) -> (usize, usize) {
    assert!(restore.starts_with(b"LIR1"));
    let intent_length = u64::from_le_bytes(restore[4..12].try_into().unwrap()) as usize;
    let delta_length_offset = 12 + intent_length;
    let delta_length = u64::from_le_bytes(
        restore[delta_length_offset..delta_length_offset + 8]
            .try_into()
            .unwrap(),
    ) as usize;
    assert_eq!(delta_length_offset + 8 + delta_length, restore.len());
    (intent_length, delta_length)
}

fn source_with_sidecar() -> Vec<u8> {
    let mut opc = source(false, false);
    part(
        &mut opc,
        "/payload/sidecar.bin",
        "application/octet-stream",
        b"sidecar",
    );
    edge(
        &mut opc,
        "/word/document.xml",
        "sidecar",
        "../payload/sidecar.bin",
        rt::IMAGE,
    );
    PackageWriter::to_bytes(&opc).unwrap()
}

fn source_with_unrelated_text() -> Vec<u8> {
    let mut opc = source(false, false);
    let main = opc
        .get_part(&PackURI::new("/word/document.xml").unwrap())
        .unwrap()
        .blob()
        .to_vec();
    let main = replace_once(
        &main,
        b"<u:r><u:contentPart",
        b"<u:r><u:t>unrelated</u:t><u:contentPart",
    );
    opc.get_part_mut(&PackURI::new("/word/document.xml").unwrap())
        .unwrap()
        .set_blob(main);
    PackageWriter::to_bytes(&opc).unwrap()
}

fn source_with_unrelated_relationship() -> Vec<u8> {
    let mut opc = source(false, false);
    part(
        &mut opc,
        "/payload/unrelated.xml",
        "application/xml",
        b"<u:root xmlns:u=\"urn:unrelated\"/>",
    );
    part(
        &mut opc,
        "/payload/unrelated-sidecar.bin",
        "application/octet-stream",
        b"unrelated sidecar",
    );
    edge(
        &mut opc,
        "/payload/unrelated.xml",
        "sidecar",
        "unrelated-sidecar.bin",
        "urn:producer:sidecar",
    );
    PackageWriter::to_bytes(&opc).unwrap()
}

fn durable_patch(commit: &litchi_docx::ink::Commit) -> CorePatch<Reversible> {
    let durable = commit.patch().to_durable(durable_limits()).unwrap();
    let wire = durable.to_deterministic_json().unwrap();
    let decoded =
        CorePatch::<Reversible>::from_deterministic_json(&wire, durable_limits()).unwrap();
    assert_eq!(wire, decoded.to_deterministic_json().unwrap());
    decoded
}

fn artifact_hash(bytes: &[u8]) -> String {
    let package = Package::from_reader(Cursor::new(bytes)).unwrap();
    let commit = package.edit_ink().unwrap().commit().unwrap();
    let durable = commit.patch().to_durable(durable_limits()).unwrap();
    wire_target_hash(&durable.to_deterministic_json().unwrap())
}

fn apply_and_inverse(original: &[u8], durable: &CorePatch<Reversible>, annotations: usize) {
    let mut package = Package::from_reader(Cursor::new(original)).unwrap();
    package.apply_durable_ink_patch(durable).unwrap();
    let published = save(&mut package);
    let mut reopened = Package::from_reader(Cursor::new(&published)).unwrap();
    assert_eq!(reopened.ink().unwrap().annotations().len(), annotations);
    reopened
        .apply_durable_ink_patch(&durable.inverse())
        .unwrap();
    assert_eq!(members(&save(&mut reopened)), members(original));
}

fn assert_rejected_unchanged(bytes: &[u8], durable: &CorePatch<Reversible>) {
    let mut package = Package::from_reader(Cursor::new(bytes)).unwrap();
    let before = members(&save(&mut package));
    assert!(package.apply_durable_ink_patch(durable).is_err());
    assert_eq!(members(&save(&mut package)), before);
}

fn assert_rejected_unchanged_named(bytes: &[u8], durable: &CorePatch<Reversible>, label: &str) {
    let mut package = Package::from_reader(Cursor::new(bytes)).unwrap();
    let before = members(&save(&mut package));
    let error = package.apply_durable_ink_patch(durable);
    assert!(error.is_err(), "{label}: durable patch was accepted");
    assert_eq!(
        members(&save(&mut package)),
        before,
        "{label}: package changed"
    );
}

fn assert_rejected_with_message(bytes: &[u8], durable: &CorePatch<Reversible>, expected: &str) {
    let mut package = Package::from_reader(Cursor::new(bytes)).unwrap();
    let before = members(&save(&mut package));
    let error = package
        .apply_durable_ink_patch(durable)
        .expect_err("tampered durable patch should be rejected");
    assert!(
        error.to_string().contains(expected),
        "expected {expected:?} in {error}"
    );
    assert_eq!(members(&save(&mut package)), before);
}

fn assert_limit_unchanged(
    bytes: &[u8],
    durable: &CorePatch<Reversible>,
    limits: EditLimits,
    expected_resource: Option<&str>,
) {
    let mut package = Package::from_reader(Cursor::new(bytes)).unwrap();
    let before = members(&save(&mut package));
    let Some(error) = package
        .apply_durable_ink_patch_with_limits(durable, limits)
        .err()
    else {
        panic!("durable replay should exceed the supplied bound ({expected_resource:?})");
    };
    match (error, expected_resource) {
        (Error::InkLimit { resource, .. }, Some(expected)) => {
            assert_eq!(resource, expected)
        },
        (Error::InkLimit { .. }, None) => {},
        (error, expected) => panic!("expected InkLimit {expected:?}, received {error}"),
    }
    assert_eq!(members(&save(&mut package)), before);
}

#[test]
fn durable_base_profiles_and_dialects_roundtrip_deterministically() {
    for strict in [false, true] {
        for profile in [BaseProfile::WordTextXml, BaseProfile::InkContent] {
            let original = PackageWriter::to_bytes(&source(strict, false)).unwrap();
            let package = Package::from_reader(Cursor::new(&original)).unwrap();
            let mut edit = package.edit_ink().unwrap();
            edit.insert(
                Destination::main(Position::new(0)),
                prepared(),
                Style::Base(profile),
            )
            .unwrap();
            let commit = edit.commit().unwrap();
            let durable = durable_patch(&commit);
            apply_and_inverse(&original, &durable, 2);
        }
    }
}

#[test]
fn durable_replacement_keeps_same_trace_count_and_restores_legacy_inkml_exactly() {
    let original = PackageWriter::to_bytes(&source(false, false)).unwrap();
    let package = Package::from_reader(Cursor::new(&original)).unwrap();
    let mut edit = package.edit_ink().unwrap();
    let replacement = one_trace("9 8, 7 6");
    edit.replace(Position::new(0), replacement, None).unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.changed());
    let durable = durable_patch(&commit);

    let mut published = Package::from_reader(Cursor::new(&original)).unwrap();
    published.apply_durable_ink_patch(&durable).unwrap();
    assert_eq!(published.ink().unwrap().annotations().len(), 1);
    assert_eq!(published.ink().unwrap().annotations()[0].trace_count(), 1);
    let published_members = members(&save(&mut published));
    assert!(published_members.values().any(|bytes| {
        bytes
            .windows(b"9 8, 7 6".len())
            .any(|window| window == b"9 8, 7 6")
    }));

    let mut reopened = Package::from_reader(Cursor::new(&save(&mut published))).unwrap();
    reopened
        .apply_durable_ink_patch(&durable.inverse())
        .unwrap();
    assert_eq!(members(&save(&mut reopened)), members(&original));
}

#[test]
fn durable_replacement_restores_noncanonical_relationship_and_content_type_source_exactly() {
    let original = noncanonical_shared_source();
    let source_members = members(&original);
    assert!(
        source_members["[Content_Types].xml"]
            .windows(b"<ct:Types".len())
            .any(|window| window == b"<ct:Types")
    );
    assert!(
        source_members["[Content_Types].xml"]
            .windows(b"ContentType='".len())
            .any(|window| window == b"ContentType='")
    );
    assert!(
        source_members["[Content_Types].xml"]
            .windows(b"interleaved".len())
            .next()
            .is_some()
    );
    assert!(
        source_members["word/_rels/document.xml.rels"]
            .windows(b"<pr:Relationships".len())
            .any(|window| window == b"<pr:Relationships")
    );
    assert!(
        source_members["word/_rels/document.xml.rels"]
            .windows(b"Target='".len())
            .any(|window| window == b"Target='")
    );
    assert!(
        source_members["word/_rels/document.xml.rels"]
            .windows(b"before next owner".len())
            .next()
            .is_some()
    );
    let document_uri = PackURI::new("/word/document.xml").unwrap();
    assert!(
        authored_xml_requires_source_proof(
            &document_uri,
            ct::WML_DOCUMENT_MAIN,
            &source_members["word/document.xml"],
        )
        .unwrap()
    );
    let baseline = Package::from_reader(Cursor::new(&original)).unwrap();
    let mut edit = baseline.edit_ink().unwrap();
    edit.replace(Position::new(0), one_trace("90 80, 70 60"), None)
        .unwrap();
    let durable = durable_patch(&edit.commit().unwrap());

    let mut published = Package::from_reader(Cursor::new(&original)).unwrap();
    published.apply_durable_ink_patch(&durable).unwrap();
    let published_bytes = save(&mut published);
    let published_members = members(&published_bytes);
    assert!(
        published_members
            .keys()
            .any(|name| name.starts_with("word/ink/"))
    );
    assert!(published_members.contains_key("payload/handwriting.xml"));
    let mut reopened = Package::from_reader(Cursor::new(&published_bytes)).unwrap();
    reopened
        .apply_durable_ink_patch(&durable.inverse())
        .unwrap();
    assert_eq!(members(&save(&mut reopened)), members(&original));
}

#[test]
fn durable_exclusive_ink_resource_restores_explicit_empty_relationship_member() {
    let original = source_with_empty_ink_relationship_member();
    let original_members = members(&original);
    let empty_relationships = original_members
        .get("payload/_rels/handwriting.xml.rels")
        .unwrap();
    assert!(
        empty_relationships
            .windows(b"<pr:Relationships".len())
            .any(|window| { window == b"<pr:Relationships" })
    );
    assert!(
        empty_relationships
            .windows(b"explicitly empty Ink owner relationships".len())
            .next()
            .is_some()
    );

    let baseline = Package::from_reader(Cursor::new(&original)).unwrap();
    let mut edit = baseline.edit_ink().unwrap();
    edit.remove(Position::new(0)).unwrap();
    let commit = edit.commit().unwrap();

    let mut ordinary = Package::from_reader(Cursor::new(&original)).unwrap();
    ordinary.apply_ink_patch(commit.patch()).unwrap();
    let ordinary_published = save(&mut ordinary);
    let ordinary_members = members(&ordinary_published);
    assert!(!ordinary_members.contains_key("payload/handwriting.xml"));
    assert!(!ordinary_members.contains_key("payload/_rels/handwriting.xml.rels"));
    let mut ordinary_reopened = Package::from_reader(Cursor::new(&ordinary_published)).unwrap();
    ordinary_reopened
        .apply_ink_patch(&commit.patch().inverse())
        .unwrap();
    assert_eq!(members(&save(&mut ordinary_reopened)), original_members);

    let durable = durable_patch(&commit);
    let mut published = Package::from_reader(Cursor::new(&original)).unwrap();
    published.apply_durable_ink_patch(&durable).unwrap();
    let published_bytes = save(&mut published);
    let mut reopened = Package::from_reader(Cursor::new(&published_bytes)).unwrap();
    reopened
        .apply_durable_ink_patch(&durable.inverse())
        .unwrap();
    assert_eq!(members(&save(&mut reopened)), original_members);
}

#[test]
fn durable_fallback_only_replacement_changes_image_bytes_and_restores_exactly() {
    let payload = prepared();
    let original = drawing_source(&payload, PNG);
    let replacement_image = png_with_ancillary_chunk();
    let replacement_fallback = FallbackImage::from_bytes(replacement_image.clone()).unwrap();
    let package = Package::from_reader(Cursor::new(&original)).unwrap();
    let mut edit = package.edit_ink().unwrap();
    edit.replace(Position::new(0), payload, Some(replacement_fallback))
        .unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.changed());
    let durable = durable_patch(&commit);

    let mut published = Package::from_reader(Cursor::new(&original)).unwrap();
    published.apply_durable_ink_patch(&durable).unwrap();
    let published_members = members(&save(&mut published));
    assert!(
        published_members
            .values()
            .any(|bytes| bytes == &replacement_image)
    );
    assert!(!published_members.values().any(|bytes| bytes == PNG));

    let mut reopened = Package::from_reader(Cursor::new(&save(&mut published))).unwrap();
    reopened
        .apply_durable_ink_patch(&durable.inverse())
        .unwrap();
    assert_eq!(members(&save(&mut reopened)), members(&original));
}

#[test]
fn durable_drawing_canvas_and_group_inline_and_anchor_restore_exactly() {
    let geometry = Geometry::new(914400, 914400).unwrap();
    let placements = [
        Placement::Inline(geometry),
        Placement::Anchor(AnchorGeometry::new(geometry)),
        Placement::Anchor(
            AnchorGeometry::new(geometry)
                .with_simple_position(Point::new(111, -222).unwrap())
                .with_horizontal(HorizontalPosition::align(
                    HorizontalRelativeFrom::Margin,
                    HorizontalAlignment::Center,
                ))
                .with_vertical(VerticalPosition::offset(
                    VerticalRelativeFrom::Paragraph,
                    -333,
                ))
                .with_relative_height(42)
                .with_behind_document(true)
                .with_locked(true)
                .with_layout_in_cell(false)
                .with_allow_overlap(false),
        ),
        Placement::Anchor(
            AnchorGeometry::new(geometry)
                .with_simple_position(Point::new(-444, 555).unwrap())
                .with_horizontal(HorizontalPosition::offset(
                    HorizontalRelativeFrom::RightMargin,
                    127000,
                ))
                .with_vertical(VerticalPosition::align(
                    VerticalRelativeFrom::BottomMargin,
                    VerticalAlignment::Inside,
                ))
                .with_relative_height(7)
                .with_behind_document(false)
                .with_locked(false)
                .with_layout_in_cell(true)
                .with_allow_overlap(true),
        ),
    ];
    for placement in placements {
        for kind in ["drawing", "canvas", "group"] {
            let style = match kind {
                "drawing" => Style::drawing(placement, image()),
                "canvas" => Style::canvas(placement, image()),
                "group" => Style::group(placement, image()),
                _ => unreachable!(),
            };
            let original = PackageWriter::to_bytes(&source(false, false)).unwrap();
            let payload = strokes();

            let expected_package = Package::from_reader(Cursor::new(&original)).unwrap();
            let mut expected_edit = expected_package.edit_ink().unwrap();
            expected_edit
                .insert(
                    Destination::main(Position::new(0)),
                    payload.clone(),
                    style.clone(),
                )
                .unwrap();
            let expected_commit = expected_edit.commit().unwrap();
            let mut expected_package = expected_package;
            expected_package
                .apply_ink_patch(expected_commit.patch())
                .unwrap();
            let expected_members = members(&save(&mut expected_package));

            let package = Package::from_reader(Cursor::new(&original)).unwrap();
            let mut edit = package.edit_ink().unwrap();
            edit.insert(Destination::main(Position::new(0)), payload, style)
                .unwrap();
            let commit = edit.commit().unwrap();
            let durable = durable_patch(&commit);

            let mut published = Package::from_reader(Cursor::new(&original)).unwrap();
            published.apply_durable_ink_patch(&durable).unwrap();
            assert_eq!(members(&save(&mut published)), expected_members);
            let published_bytes = save(&mut published);
            let mut reopened = Package::from_reader(Cursor::new(&published_bytes)).unwrap();
            assert_eq!(reopened.ink().unwrap().annotations().len(), 2);
            reopened
                .apply_durable_ink_patch(&durable.inverse())
                .unwrap();
            assert_eq!(members(&save(&mut reopened)), members(&original));
        }
    }
}

#[test]
fn durable_mixed_remove_replace_insert_reopens_and_inverts_exactly() {
    let original = PackageWriter::to_bytes(&source(false, true)).unwrap();
    let package = Package::from_reader(Cursor::new(&original)).unwrap();
    let mut edit = package.edit_ink().unwrap();
    edit.remove(Position::new(0)).unwrap();
    edit.replace(Position::new(1), prepared(), None).unwrap();
    edit.insert(
        Destination::main(Position::new(0)),
        prepared(),
        Style::Base(BaseProfile::InkContent),
    )
    .unwrap();
    let commit = edit.commit().unwrap();
    let durable = durable_patch(&commit);
    apply_and_inverse(&original, &durable, 8);
}

#[test]
fn durable_noop_is_deterministic_exact_and_source_bound() {
    let original = PackageWriter::to_bytes(&source(false, false)).unwrap();
    let package = Package::from_reader(Cursor::new(&original)).unwrap();
    let commit = package.edit_ink().unwrap().commit().unwrap();
    assert!(!commit.changed());
    let durable = durable_patch(&commit);

    let mut package = Package::from_reader(Cursor::new(&original)).unwrap();
    let before = members(&save(&mut package));
    assert_eq!(
        package
            .apply_durable_ink_patch(&durable)
            .unwrap()
            .annotations()
            .len(),
        1
    );
    assert_eq!(members(&save(&mut package)), before);

    let stale = rewrite_member(&original, "word/_rels/document.xml.rels", lexical_comment);
    assert_rejected_unchanged(&stale, &durable);
}

#[test]
fn durable_changed_and_noop_patches_reject_payload_graph_and_lexical_stale_sources() {
    let original = PackageWriter::to_bytes(&source(false, false)).unwrap();
    let baseline = Package::from_reader(Cursor::new(&original)).unwrap();
    let mut edit = baseline.edit_ink().unwrap();
    edit.insert(
        Destination::main(Position::new(0)),
        prepared(),
        Style::Base(BaseProfile::InkContent),
    )
    .unwrap();
    let changed = durable_patch(&edit.commit().unwrap());
    let noop = durable_patch(&baseline.edit_ink().unwrap().commit().unwrap());

    let payload = rewrite_member(&original, "payload/handwriting.xml", |content| {
        replace_once(content, b"1 2, 3 4", b"4 5, 6 7")
    });
    assert_rejected_unchanged(&payload, &changed);
    assert_rejected_unchanged(&payload, &noop);

    let graph_original = source_with_sidecar();
    let graph_baseline = Package::from_reader(Cursor::new(&graph_original)).unwrap();
    let mut graph_edit = graph_baseline.edit_ink().unwrap();
    graph_edit
        .insert(
            Destination::main(Position::new(0)),
            prepared(),
            Style::Base(BaseProfile::InkContent),
        )
        .unwrap();
    let graph_patch = durable_patch(&graph_edit.commit().unwrap());
    let graph = rewrite_member(&graph_original, "word/_rels/document.xml.rels", |content| {
        replace_once(
            content,
            b"Target=\"../payload/sidecar.bin\"",
            b"Target=\"../payload/handwriting.xml\"",
        )
    });
    assert_rejected_unchanged(&graph, &graph_patch);

    for member in [
        "[Content_Types].xml",
        "_rels/.rels",
        "word/_rels/document.xml.rels",
    ] {
        let lexical = rewrite_member(&original, member, lexical_comment);
        assert_rejected_unchanged(&lexical, &changed);
        assert_rejected_unchanged(&lexical, &noop);
    }

    let unrelated_original = source_with_unrelated_relationship();
    let unrelated_baseline = Package::from_reader(Cursor::new(&unrelated_original)).unwrap();
    let mut unrelated_edit = unrelated_baseline.edit_ink().unwrap();
    unrelated_edit
        .insert(
            Destination::main(Position::new(0)),
            prepared(),
            Style::Base(BaseProfile::InkContent),
        )
        .unwrap();
    let unrelated_changed = durable_patch(&unrelated_edit.commit().unwrap());
    let unrelated_noop = durable_patch(&unrelated_baseline.edit_ink().unwrap().commit().unwrap());
    let unrelated_lexical = rewrite_member(
        &unrelated_original,
        "payload/_rels/unrelated.xml.rels",
        lexical_comment,
    );
    assert_rejected_unchanged(&unrelated_lexical, &unrelated_changed);
    assert_rejected_unchanged(&unrelated_lexical, &unrelated_noop);

    let non_part_original = add_member(&original, "custom/untyped.bin", b"opaque source");
    let non_part_baseline = Package::from_reader(Cursor::new(&non_part_original)).unwrap();
    let mut non_part_edit = non_part_baseline.edit_ink().unwrap();
    non_part_edit
        .insert(
            Destination::main(Position::new(0)),
            prepared(),
            Style::Base(BaseProfile::InkContent),
        )
        .unwrap();
    let non_part_changed = durable_patch(&non_part_edit.commit().unwrap());
    let non_part_noop = durable_patch(&non_part_baseline.edit_ink().unwrap().commit().unwrap());
    // Opaque non-part bytes are preserved by the owning archive but are not
    // part of the logical OPC source fingerprint. Name/reason inventory is
    // bound, so topology changes must still be rejected.
    let non_part_lexical = rewrite_member(&non_part_original, "custom/untyped.bin", |content| {
        replace_once(content, b"opaque source", b"opaque forged")
    });
    // The forward topology has added OPC parts, so the owned writer cannot
    // normalize an archive containing an unknown non-part while it is in that
    // intermediate state. Replay and inverse remain source-bound in memory;
    // once the inverse restores the original topology, the opaque bytes must
    // still publish byte-for-byte.
    let mut non_part_package = Package::from_reader(Cursor::new(&non_part_lexical)).unwrap();
    non_part_package
        .apply_durable_ink_patch(&non_part_changed)
        .unwrap();
    assert_eq!(non_part_package.ink().unwrap().annotations().len(), 2);
    non_part_package
        .apply_durable_ink_patch(&non_part_changed.inverse())
        .unwrap();
    assert_eq!(non_part_package.ink().unwrap().annotations().len(), 1);
    assert_eq!(
        members(&save(&mut non_part_package)),
        members(&non_part_lexical)
    );
    let mut non_part_noop_package = Package::from_reader(Cursor::new(&non_part_lexical)).unwrap();
    let non_part_noop_before = members(&save(&mut non_part_noop_package));
    non_part_noop_package
        .apply_durable_ink_patch(&non_part_noop)
        .unwrap();
    assert_eq!(
        members(&save(&mut non_part_noop_package)),
        non_part_noop_before
    );

    let non_part_added = add_member(&non_part_original, "custom/added.bin", b"added");
    let non_part_removed = remove_member(&non_part_original, "custom/untyped.bin");
    let non_part_renamed = add_member(&non_part_removed, "custom/renamed.bin", b"opaque source");
    for (label, stale) in [
        ("non-part added", non_part_added),
        ("non-part removed", non_part_removed),
        ("non-part renamed", non_part_renamed),
    ] {
        assert_rejected_unchanged_named(&stale, &non_part_changed, label);
        assert_rejected_unchanged_named(&stale, &non_part_noop, label);
    }

    let mut empty_opc = source(false, false);
    part(
        &mut empty_opc,
        "/payload/empty.xml",
        "application/xml",
        b"<e:root xmlns:e=\"urn:empty\"/>",
    );
    let empty_original = PackageWriter::to_bytes(&empty_opc).unwrap();
    let empty_baseline = Package::from_reader(Cursor::new(&empty_original)).unwrap();
    let mut empty_edit = empty_baseline.edit_ink().unwrap();
    empty_edit
        .insert(
            Destination::main(Position::new(0)),
            prepared(),
            Style::Base(BaseProfile::InkContent),
        )
        .unwrap();
    let empty_changed = durable_patch(&empty_edit.commit().unwrap());
    let empty_noop = durable_patch(&empty_baseline.edit_ink().unwrap().commit().unwrap());
    let empty_relationships = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>"#;
    let empty_member = add_member(
        &empty_original,
        "payload/_rels/empty.xml.rels",
        empty_relationships,
    );
    assert_rejected_unchanged(&empty_member, &empty_changed);
    assert_rejected_unchanged(&empty_member, &empty_noop);
}

#[test]
fn durable_tampering_and_limits_leave_the_package_unchanged() {
    let original = PackageWriter::to_bytes(&source(false, false)).unwrap();
    let baseline = Package::from_reader(Cursor::new(&original)).unwrap();
    let mut edit = baseline.edit_ink().unwrap();
    edit.insert(
        Destination::main(Position::new(0)),
        prepared(),
        Style::Base(BaseProfile::InkContent),
    )
    .unwrap();
    let commit = edit.commit().unwrap();
    let durable = durable_patch(&commit);
    let wire = durable.to_deterministic_json().unwrap();

    let intent_wire = replace_once(&wire, b"ink.edit", b"ink.edIt");
    let intent =
        CorePatch::<Reversible>::from_deterministic_json(&intent_wire, durable_limits()).unwrap();
    assert_rejected_unchanged(&original, &intent);

    let target_wire = replace_once(&wire, b"\"target\":\"package\"", b"\"target\":\"packagE\"");
    let target =
        CorePatch::<Reversible>::from_deterministic_json(&target_wire, durable_limits()).unwrap();
    assert_rejected_unchanged(&original, &target);

    let target_hash_wire = {
        let marker = b"\"target_sha256\":\"";
        let start = wire
            .windows(marker.len())
            .position(|window| window == marker)
            .unwrap()
            + marker.len();
        let mut tampered = wire.clone();
        tampered[start] = if tampered[start] == b'0' { b'1' } else { b'0' };
        tampered
    };
    let target_hash =
        CorePatch::<Reversible>::from_deterministic_json(&target_hash_wire, durable_limits())
            .unwrap();
    assert_rejected_unchanged(&original, &target_hash);

    let blob_wire = flip_first_blob_digest(&wire);
    assert!(
        CorePatch::<Reversible>::from_deterministic_json(&blob_wire, durable_limits()).is_err()
    );

    let tiny = PatchLimits::new(BlobLimits::new(1, 1, 1), 1024, 1, 8, 128, 1);
    assert!(commit.patch().to_durable(tiny).is_err());
    assert!(CorePatch::<Reversible>::from_deterministic_json(&wire, tiny).is_err());

    let mut package = Package::from_reader(Cursor::new(&original)).unwrap();
    assert_eq!(members(&save(&mut package)), members(&original));
}

#[test]
fn durable_replay_limits_refuse_operations_staging_package_and_payloads_atomically() {
    let mixed_original = PackageWriter::to_bytes(&source(false, true)).unwrap();
    let mixed_baseline = Package::from_reader(Cursor::new(&mixed_original)).unwrap();
    let mut mixed_edit = mixed_baseline.edit_ink().unwrap();
    mixed_edit.remove(Position::new(0)).unwrap();
    mixed_edit
        .replace(Position::new(1), prepared(), None)
        .unwrap();
    mixed_edit
        .insert(
            Destination::main(Position::new(0)),
            prepared(),
            Style::Base(BaseProfile::InkContent),
        )
        .unwrap();
    let mixed = durable_patch(&mixed_edit.commit().unwrap());
    let operation_limits = EditLimits {
        max_operations: 2,
        ..EditLimits::default()
    };
    assert_limit_unchanged(
        &mixed_original,
        &mixed,
        operation_limits,
        Some("durable intent count"),
    );

    let canonical_original = PackageWriter::to_bytes(&source(false, false)).unwrap();
    let canonical_baseline = Package::from_reader(Cursor::new(&canonical_original)).unwrap();
    let canonical_payload = prepared();
    let mut canonical_edit = canonical_baseline.edit_ink().unwrap();
    canonical_edit
        .insert(
            Destination::main(Position::new(0)),
            canonical_payload.clone(),
            Style::Base(BaseProfile::InkContent),
        )
        .unwrap();
    let canonical = durable_patch(&canonical_edit.commit().unwrap());
    let (_, _, intent) = wire_blob(&canonical.to_deterministic_json().unwrap(), "forward_blobs");
    let staged_limits = EditLimits {
        max_staged_bytes: intent.len().saturating_sub(1).max(1),
        ..EditLimits::default()
    };
    assert_limit_unchanged(
        &canonical_original,
        &canonical,
        staged_limits,
        Some("durable intent bytes"),
    );

    let package_limits = EditLimits {
        max_package_bytes: 1,
        ..EditLimits::default()
    };
    let canonical_noop = durable_patch(&canonical_baseline.edit_ink().unwrap().commit().unwrap());
    assert_limit_unchanged(&canonical_original, &canonical, package_limits, None);
    assert_limit_unchanged(&canonical_original, &canonical_noop, package_limits, None);

    let mut payload_limits = EditLimits::default();
    payload_limits.inventory.max_payload_bytes =
        canonical_payload.as_bytes().len().saturating_sub(1).max(1);
    assert_limit_unchanged(&canonical_original, &canonical, payload_limits, None);

    let drawing_original = PackageWriter::to_bytes(&source(false, false)).unwrap();
    let drawing_baseline = Package::from_reader(Cursor::new(&drawing_original)).unwrap();
    let geometry = Geometry::new(914400, 914400).unwrap();
    let fallback_payload = prepared();
    let fallback_image = image();
    let mut drawing_edit = drawing_baseline.edit_ink().unwrap();
    drawing_edit
        .insert(
            Destination::main(Position::new(0)),
            fallback_payload.clone(),
            Style::drawing(Placement::Inline(geometry), fallback_image.clone()),
        )
        .unwrap();
    let drawing = durable_patch(&drawing_edit.commit().unwrap());
    let mut fallback_limits = EditLimits::default();
    fallback_limits.inventory.max_payload_bytes =
        fallback_payload.as_bytes().len().saturating_sub(1).max(1);
    assert_limit_unchanged(
        &drawing_original,
        &drawing,
        fallback_limits,
        Some("payload bytes"),
    );

    let mut canonical_published = Package::from_reader(Cursor::new(&canonical_original)).unwrap();
    canonical_published
        .apply_durable_ink_patch(&canonical)
        .unwrap();
    let canonical_published_bytes = save(&mut canonical_published);
    let inverse = canonical.inverse();
    let (_, _, restore) = wire_blob(&inverse.to_deterministic_json().unwrap(), "forward_blobs");
    let (intent_length, delta_length) = restore_component_lengths(&restore);
    let source_part_bytes: usize = canonical_published
        .opc_package()
        .try_iter_parts()
        .map(|part| part.expect("decode part payload").blob().len())
        .sum();
    assert!(source_part_bytes < delta_length);

    let inverse_package_limits = EditLimits {
        max_package_bytes: source_part_bytes + 1,
        ..EditLimits::default()
    };
    assert_limit_unchanged(
        &canonical_published_bytes,
        &inverse,
        inverse_package_limits,
        Some("durable restore closure bytes"),
    );

    let inverse_staged_limits = EditLimits {
        max_staged_bytes: intent_length.saturating_sub(1).max(1),
        ..EditLimits::default()
    };
    assert_limit_unchanged(
        &canonical_published_bytes,
        &inverse,
        inverse_staged_limits,
        Some("durable restore intent bytes"),
    );

    let mut inverse_topology_limits = EditLimits::default();
    inverse_topology_limits.inventory.stories.max_topology_bytes = 1024;
    assert_limit_unchanged(
        &canonical_published_bytes,
        &inverse,
        inverse_topology_limits,
        Some("Ink graph publication token bytes"),
    );

    let mut drawing_published = Package::from_reader(Cursor::new(&drawing_original)).unwrap();
    drawing_published.apply_durable_ink_patch(&drawing).unwrap();
    let drawing_published_bytes = save(&mut drawing_published);
    let drawing_inverse = drawing.inverse();
    let mut inverse_relationship_limits = EditLimits::default();
    inverse_relationship_limits.inventory.max_relationships = 1;
    assert_limit_unchanged(
        &drawing_published_bytes,
        &drawing_inverse,
        inverse_relationship_limits,
        Some("relationships"),
    );
}

#[test]
fn durable_semantic_intent_tampering_is_rejected_after_recomputing_blob_and_target() {
    let original = PackageWriter::to_bytes(&source(false, false)).unwrap();
    let baseline = Package::from_reader(Cursor::new(&original)).unwrap();
    let replacement = strokes();
    let mut edit = baseline.edit_ink().unwrap();
    edit.replace(Position::new(0), replacement.clone(), None)
        .unwrap();
    let durable = durable_patch(&edit.commit().unwrap());
    let wire = durable.to_deterministic_json().unwrap();
    let (_, _, intent) = wire_blob(&wire, "forward_blobs");
    let tampered_intent = replace_once(&intent, b"0 0", b"9 9");
    let tampered_wire =
        replace_wire_blob(&wire, "forward_blobs", "intent_sha256", &tampered_intent);

    let altered_payload =
        Prepared::from_bytes(&replace_once(replacement.as_bytes(), b"0 0", b"9 9")).unwrap();
    let mut altered = Package::from_reader(Cursor::new(&original)).unwrap();
    let mut altered_edit = altered.edit_ink().unwrap();
    altered_edit
        .replace(Position::new(0), altered_payload, None)
        .unwrap();
    let altered_commit = altered_edit.commit().unwrap();
    altered.apply_ink_patch(altered_commit.patch()).unwrap();
    let altered_bytes = save(&mut altered);
    let altered_target = rewrite_member(&altered_bytes, "word/document.xml", |content| {
        replace_once(
            content,
            b"<u:body>",
            b"<u:body><!-- unrelated target edit -->",
        )
    });
    let tampered_wire = replace_wire_target(&tampered_wire, &artifact_hash(&altered_target));
    let tampered =
        CorePatch::<Reversible>::from_deterministic_json(&tampered_wire, durable_limits()).unwrap();
    assert_rejected_unchanged(&original, &tampered);
}

#[test]
fn durable_coherent_alternate_intent_replays_to_its_matching_target() {
    let original = PackageWriter::to_bytes(&source(false, false)).unwrap();
    let baseline = Package::from_reader(Cursor::new(&original)).unwrap();
    let replacement = strokes();
    let mut edit = baseline.edit_ink().unwrap();
    edit.replace(Position::new(0), replacement.clone(), None)
        .unwrap();
    let durable = durable_patch(&edit.commit().unwrap());
    let wire = durable.to_deterministic_json().unwrap();
    let (_, _, intent) = wire_blob(&wire, "forward_blobs");
    let tampered_intent = replace_once(&intent, b"0 0", b"9 9");
    let tampered_wire =
        replace_wire_blob(&wire, "forward_blobs", "intent_sha256", &tampered_intent);

    let altered_payload =
        Prepared::from_bytes(&replace_once(replacement.as_bytes(), b"0 0", b"9 9")).unwrap();
    let mut expected = Package::from_reader(Cursor::new(&original)).unwrap();
    let mut expected_edit = expected.edit_ink().unwrap();
    expected_edit
        .replace(Position::new(0), altered_payload, None)
        .unwrap();
    let expected_commit = expected_edit.commit().unwrap();
    expected.apply_ink_patch(expected_commit.patch()).unwrap();
    let expected_bytes = save(&mut expected);
    let tampered_wire = replace_wire_target(&tampered_wire, &artifact_hash(&expected_bytes));
    let tampered =
        CorePatch::<Reversible>::from_deterministic_json(&tampered_wire, durable_limits()).unwrap();

    let mut actual = Package::from_reader(Cursor::new(&original)).unwrap();
    actual.apply_durable_ink_patch(&tampered).unwrap();
    assert_eq!(members(&save(&mut actual)), members(&expected_bytes));
}

#[test]
fn durable_restore_tampering_rejects_recomputed_closure_and_intent_variants() {
    let original = source_with_unrelated_text();
    let baseline = Package::from_reader(Cursor::new(&original)).unwrap();
    let mut edit = baseline.edit_ink().unwrap();
    edit.insert(
        Destination::main(Position::new(0)),
        strokes(),
        Style::Base(BaseProfile::InkContent),
    )
    .unwrap();
    let durable = durable_patch(&edit.commit().unwrap());

    let mut published = Package::from_reader(Cursor::new(&original)).unwrap();
    published.apply_durable_ink_patch(&durable).unwrap();
    let published_bytes = save(&mut published);
    let inverse_wire = durable.inverse().to_deterministic_json().unwrap();
    let (_, _, restore) = wire_blob(&inverse_wire, "forward_blobs");

    let altered_source = rewrite_member(&original, "word/document.xml", |content| {
        replace_once(content, b"unrelated", b"tampered!")
    });
    let altered_restore = replace_last(&restore, b"unrelated", b"tampered!");
    let closure_wire = replace_wire_blob(
        &inverse_wire,
        "forward_blobs",
        "restore_sha256",
        &altered_restore,
    );
    let closure_wire = replace_wire_target(&closure_wire, &artifact_hash(&altered_source));
    let closure_patch =
        CorePatch::<Reversible>::from_deterministic_json(&closure_wire, durable_limits()).unwrap();
    assert_rejected_with_message(
        &published_bytes,
        &closure_patch,
        "Ink durable restore forward replay mismatch",
    );

    let altered_intent_restore = mutate_restore_intent(&restore, b"0 0", b"9 9");
    let intent_wire = replace_wire_blob(
        &inverse_wire,
        "forward_blobs",
        "restore_sha256",
        &altered_intent_restore,
    );
    let intent_wire = replace_wire_target(&intent_wire, &artifact_hash(&original));
    let intent_patch =
        CorePatch::<Reversible>::from_deterministic_json(&intent_wire, durable_limits()).unwrap();
    assert_rejected_with_message(
        &published_bytes,
        &intent_patch,
        "Ink durable restore forward replay mismatch",
    );
}

#[test]
fn durable_restore_rejects_existing_part_content_type_tampering() {
    let original = source_with_unrelated_text();
    let baseline = Package::from_reader(Cursor::new(&original)).unwrap();
    let mut edit = baseline.edit_ink().unwrap();
    edit.insert(
        Destination::main(Position::new(0)),
        strokes(),
        Style::Base(BaseProfile::InkContent),
    )
    .unwrap();
    let durable = durable_patch(&edit.commit().unwrap());

    let mut published = Package::from_reader(Cursor::new(&original)).unwrap();
    published.apply_durable_ink_patch(&durable).unwrap();
    let published_bytes = save(&mut published);
    let inverse_wire = durable.inverse().to_deterministic_json().unwrap();
    let (_, _, restore) = wire_blob(&inverse_wire, "forward_blobs");
    let original_content_type =
        b"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
    let altered_content_type =
        b"application/vnd.openxmlformats-officedocument.wordprocessingml.document.mAin+xml";
    assert_eq!(original_content_type.len(), altered_content_type.len());
    let altered_restore = replace_last(&restore, original_content_type, altered_content_type);
    assert_ne!(altered_restore, restore);
    let tampered_wire = replace_wire_blob(
        &inverse_wire,
        "forward_blobs",
        "restore_sha256",
        &altered_restore,
    );
    let tampered =
        CorePatch::<Reversible>::from_deterministic_json(&tampered_wire, durable_limits()).unwrap();
    assert_rejected_with_message(
        &published_bytes,
        &tampered,
        "existing Ink durable part content type changed",
    );
}

#[test]
fn durable_changed_publish_requires_explicit_unsigned_transition() {
    let mut opc = source(false, false);
    part(
        &mut opc,
        "/_xmlsignatures/origin.sigs",
        ct::OPC_DIGITAL_SIGNATURE_ORIGIN,
        b"",
    );
    opc.relate_to("_xmlsignatures/origin.sigs", rt::DIGITAL_SIGNATURE_ORIGIN);
    let original = PackageWriter::to_bytes(&opc).unwrap();
    let mut signed = Package::from_reader(Cursor::new(&original)).unwrap();
    let before = members(&save(&mut signed));
    let mut edit = signed.edit_ink().unwrap();
    edit.insert(
        Destination::main(Position::new(0)),
        prepared(),
        Style::Base(BaseProfile::InkContent),
    )
    .unwrap();
    assert!(edit.commit().is_err());
    assert_eq!(members(&save(&mut signed)), before);
}
