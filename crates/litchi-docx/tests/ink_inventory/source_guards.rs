use std::collections::BTreeMap;

use litchi_core::Position;
use litchi_docx::Package;
use litchi_opc::PackageWriter;
use soapberry_zip::office::{ArchiveReader, StreamingArchiveWriter};

use super::*;

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
        .expect("fixture XML has its closing element");
    let mut result = Vec::with_capacity(content.len() + 32);
    result.extend_from_slice(&content[..index]);
    result.extend_from_slice(b"\n<!-- lexical guard -->\n");
    result.extend_from_slice(&content[index..]);
    result
}

fn lexical_comment_with_padding(content: &[u8], padding: usize) -> Vec<u8> {
    let closing = b"</Relationships>";
    let index = content
        .windows(closing.len())
        .position(|window| window == closing)
        .expect("fixture relationship XML has its closing element");
    let mut result = Vec::with_capacity(content.len() + padding + 16);
    result.extend_from_slice(&content[..index]);
    result.extend_from_slice(b"<!--");
    result.extend(std::iter::repeat_n(b'x', padding));
    result.extend_from_slice(b"-->");
    result.extend_from_slice(&content[index..]);
    result
}

fn removal_patch(package: &Package) -> litchi_docx::ink::Patch {
    let mut edit = package.edit_ink().unwrap();
    edit.remove(Position::new(0)).unwrap();
    edit.commit().unwrap().patch().clone()
}

fn noop_patch(package: &Package) -> litchi_docx::ink::Patch {
    package
        .edit_ink()
        .unwrap()
        .commit()
        .unwrap()
        .patch()
        .clone()
}

fn base_source_with_unrelated_relationship() -> Vec<u8> {
    let mut opc = source(false, false);
    part(
        &mut opc,
        "/payload/unrelated.xml",
        "application/xml",
        b"<u:root xmlns:u=\"urn:unrelated\"/>",
    );
    part(
        &mut opc,
        "/payload/sidecar.bin",
        "application/octet-stream",
        b"sidecar",
    );
    edge(
        &mut opc,
        "/payload/unrelated.xml",
        "sidecar",
        "sidecar.bin",
        "urn:producer:sidecar",
    );
    PackageWriter::to_bytes(&opc).unwrap()
}

fn assert_source_guard_rejects(variant: Vec<u8>, patch: &litchi_docx::ink::Patch) {
    let mut package = Package::from_reader(Cursor::new(&variant)).unwrap();
    let before = members(&save(&mut package));
    assert!(package.apply_ink_patch(patch).is_err());
    assert_eq!(members(&save(&mut package)), before);
}

#[test]
fn changed_and_noop_patches_reject_lexical_only_package_members_atomically() {
    let original = PackageWriter::to_bytes(&source(false, false)).unwrap();
    let baseline = Package::from_reader(Cursor::new(&original)).unwrap();
    let changed = removal_patch(&baseline);
    let noop = noop_patch(&baseline);

    for member in [
        "[Content_Types].xml",
        "_rels/.rels",
        "word/_rels/document.xml.rels",
    ] {
        let variant = rewrite_member(&original, member, lexical_comment);
        assert_source_guard_rejects(variant.clone(), &changed);
        assert_source_guard_rejects(variant, &noop);
    }

    let unrelated = base_source_with_unrelated_relationship();
    let baseline = Package::from_reader(Cursor::new(&unrelated)).unwrap();
    let changed = removal_patch(&baseline);
    let noop = noop_patch(&baseline);
    let variant = rewrite_member(
        &unrelated,
        "payload/_rels/unrelated.xml.rels",
        lexical_comment,
    );
    assert_source_guard_rejects(variant.clone(), &changed);
    assert_source_guard_rejects(variant, &noop);
}

#[test]
fn adding_an_empty_relationship_member_is_a_stale_source_for_changed_and_noop_patches() {
    let mut opc = source(false, false);
    part(
        &mut opc,
        "/payload/empty.xml",
        "application/xml",
        b"<e:root xmlns:e=\"urn:empty\"/>",
    );
    let original = PackageWriter::to_bytes(&opc).unwrap();
    let baseline = Package::from_reader(Cursor::new(&original)).unwrap();
    let changed = removal_patch(&baseline);
    let noop = noop_patch(&baseline);
    let empty_relationships = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>"#;
    let variant = add_member(
        &original,
        "payload/_rels/empty.xml.rels",
        empty_relationships,
    );
    let absent = baseline
        .opc_package()
        .source_relationships(&PackURI::new("/payload/empty.xml").unwrap())
        .unwrap();
    let present_package = Package::from_reader(Cursor::new(&variant)).unwrap();
    let present = present_package
        .opc_package()
        .source_relationships(&PackURI::new("/payload/empty.xml").unwrap())
        .unwrap();
    assert_eq!(absent.bytes(), present.bytes());
    assert!(!absent.member_present());
    assert!(present.member_present());
    assert_source_guard_rejects(variant.clone(), &changed);
    assert_source_guard_rejects(variant, &noop);
}

#[test]
fn relationship_lexical_tokens_count_against_public_ink_topology_limit() {
    let original = PackageWriter::to_bytes(&source(false, false)).unwrap();
    let variant = rewrite_member(&original, "word/_rels/document.xml.rels", |content| {
        lexical_comment_with_padding(content, 8 * 1024)
    });
    let mut package = Package::from_reader(Cursor::new(&variant)).unwrap();
    let semantic_topology_bytes = package
        .story_inventory()
        .unwrap()
        .topology()
        .as_bytes()
        .len();
    let mut limits = Limits::default();
    // The story and graph semantics fit comfortably below this budget. The
    // relationship comment alone is larger, so only publication-token bytes
    // should exhaust the bound.
    limits.stories.max_topology_bytes = 4 * 1024;
    assert!(semantic_topology_bytes < limits.stories.max_topology_bytes);
    let before = members(&save(&mut package));

    let error = package
        .edit_ink_with_limits(litchi_docx::ink::EditLimits {
            inventory: limits,
            ..litchi_docx::ink::EditLimits::default()
        })
        .err()
        .expect("lexical relationship tokens should exceed the semantic-only budget");
    assert!(matches!(
        error,
        Error::InkLimit {
            resource: "Ink graph publication token bytes",
            ..
        }
    ));
    assert_eq!(members(&save(&mut package)), before);
}
