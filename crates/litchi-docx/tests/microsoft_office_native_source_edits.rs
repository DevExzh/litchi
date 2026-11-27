//! Native source-backed story text and inert external-hyperlink publication.
use litchi_core::{OwnedSource, Position};
use litchi_docx::{Package, source_backed};
use std::{collections::BTreeMap, io::Cursor, sync::Arc};
fn native(name: &str) -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/word")
            .join(name),
    )
    .unwrap()
}
fn open(bytes: &[u8]) -> source_backed::Package {
    source_backed::Package::from_read_at(Arc::new(OwnedSource::new(bytes.to_vec()))).unwrap()
}
fn bytes(p: &source_backed::Package) -> Vec<u8> {
    let mut b = Vec::new();
    p.write_to(&mut b).unwrap();
    b
}
fn members(bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
    let archive = litchi_opc::phys_pkg::OwnedPhysPkgReader::from_bytes(
        bytes.to_vec(),
        litchi_opc::SourceLimits::default(),
    )
    .unwrap();
    archive
        .member_names()
        .unwrap()
        .into_iter()
        .map(|name| {
            let bytes = archive.read_member(&name).unwrap();
            (name, bytes)
        })
        .collect()
}
fn unchanged_except(before: &[u8], after: &[u8], parts: &[&str]) {
    let a = members(before);
    let b = members(after);
    assert_eq!(a.keys().collect::<Vec<_>>(), b.keys().collect::<Vec<_>>());
    let changed = a
        .iter()
        .filter(|(n, v)| b[*n] != **v)
        .map(|(n, _)| n.as_str())
        .collect::<Vec<_>>();
    assert_eq!(changed, parts);
}
fn save(p: &source_backed::Package) -> Vec<u8> {
    let mut f = tempfile::NamedTempFile::with_suffix(".docx").unwrap();
    p.write_to(&mut f).unwrap();
    std::fs::read(f.path()).unwrap()
}
fn detached_document(original: &[u8]) -> Vec<u8> {
    let xml = String::from_utf8(members(original)["word/document.xml"].clone()).unwrap();
    let opening = "<w:hyperlink r:id=\"rId11\" w:history=\"1\">";
    assert_eq!(xml.matches(opening).count(), 1);
    assert_eq!(xml.matches("</w:hyperlink>").count(), 1);
    xml.replace(opening, "")
        .replace("</w:hyperlink>", "")
        .into_bytes()
}
fn relationships(xml: &[u8]) -> BTreeMap<String, BTreeMap<String, String>> {
    use quick_xml::events::Event;
    let mut reader = quick_xml::Reader::from_reader(xml);
    let mut out = BTreeMap::new();
    loop {
        match reader.read_event().unwrap() {
            Event::Start(e) | Event::Empty(e) if e.local_name().as_ref() == b"Relationship" => {
                let attrs = e
                    .attributes()
                    .map(|a| {
                        let a = a.unwrap();
                        (
                            String::from_utf8(a.key.as_ref().to_vec()).unwrap(),
                            a.decoded_and_normalized_value(
                                quick_xml::XmlVersion::Explicit1_0,
                                reader.decoder(),
                            )
                            .unwrap()
                            .into_owned(),
                        )
                    })
                    .collect::<BTreeMap<_, _>>();
                assert!(out.insert(attrs["Id"].clone(), attrs).is_none());
            },
            Event::Eof => break,
            _ => {},
        }
    }
    out
}
// Native note paragraphs have a separate leading-space text slot. Whole
// paragraph replacement may redistribute text among those owned slots, but
// the run properties, note markers and every other XML byte must remain.
fn mask_first_text_slots(xml: &[u8], count: usize) -> Vec<u8> {
    use quick_xml::events::Event;
    let mut reader = quick_xml::Reader::from_reader(xml);
    let mut output = Vec::new();
    let mut copied = 0;
    let mut replaced = 0;
    let mut start = None;
    loop {
        let before = usize::try_from(reader.buffer_position()).unwrap();
        let event = reader.read_event().unwrap();
        let after = usize::try_from(reader.buffer_position()).unwrap();
        let span = match event {
            Event::Start(e) if e.name().as_ref() == b"w:t" => {
                start = Some(before);
                None
            },
            Event::End(e) if e.name().as_ref() == b"w:t" => Some((start.take().unwrap(), after)),
            Event::Empty(e) if e.name().as_ref() == b"w:t" => Some((before, after)),
            Event::Eof => panic!("missing owned text slot"),
            _ => None,
        };
        if let Some((start, end)) = span {
            output.extend_from_slice(&xml[copied..start]);
            output.extend_from_slice(b"<text-slot/>");
            copied = end;
            replaced += 1;
            if replaced == count {
                output.extend_from_slice(&xml[copied..]);
                return output;
            }
        }
    }
}
#[test]
fn native_hyperlink_detachment_retains_text_relationship_and_all_other_members() {
    let original = native("word-text-structure.docx");
    let p = open(&original);
    let snapshot = p.external_hyperlink_sanitization_snapshot().unwrap();
    assert_eq!(snapshot.external_hyperlink_count(), 1);
    let edit = p.edit_external_hyperlink_detachment().unwrap();
    let effects = edit.effect_report();
    assert_eq!(effects.detached_hyperlinks(), 1);
    assert_eq!(effects.referenced_external_relationships(), 1);
    assert_eq!(effects.retained_relationships(), 1);
    let commit = edit.commit().unwrap();
    let after = save(commit.package());
    let read = open(&after);
    assert_eq!(
        read.external_hyperlink_sanitization_snapshot()
            .unwrap()
            .external_hyperlink_count(),
        0
    );
    let doc = Package::from_reader(Cursor::new(&after)).unwrap();
    assert_eq!(
        doc.document().unwrap().paragraphs().unwrap()[10]
            .text()
            .unwrap(),
        "Reference link"
    );
    unchanged_except(&original, &after, &["word/document.xml"]);
    assert!(members(&after)["word/document.xml"] == detached_document(&original));
    let noop = read.edit_external_hyperlink_detachment().unwrap();
    assert!(noop.effect_report().is_noop());
    assert!(bytes(noop.commit().unwrap().package()) == after);
    let restored = read
        .apply_external_hyperlink_detachment(&commit.patch().inverse())
        .unwrap();
    unchanged_except(&original, &bytes(restored.package()), &[]);
    assert!(bytes(&p) == original);
}
#[test]
fn native_hyperlink_redaction_removes_only_the_selected_relationship_and_wrapper() {
    let original = native("word-text-structure.docx");
    let p = open(&original);
    let snapshot = p.external_hyperlink_redaction_snapshot().unwrap();
    assert_eq!(snapshot.relationships().len(), 1);
    assert_eq!(snapshot.relationships()[0].relationship_id(), "rId11");
    assert_eq!(
        snapshot.relationships()[0].target_url(),
        "https://example.com/litchi-fixture"
    );
    let edit = p
        .edit_external_hyperlink_redaction(&["https://example.com/litchi-fixture"])
        .unwrap();
    let effects = edit.effect_report();
    assert_eq!(
        (
            effects.selected_targets(),
            effects.removed_relationships(),
            effects.unwrapped_hyperlinks()
        ),
        (1, 1, 1)
    );
    let commit = edit.commit().unwrap();
    let after = save(commit.package());
    let read = open(&after);
    assert!(
        read.external_hyperlink_redaction_snapshot()
            .unwrap()
            .relationships()
            .is_empty()
    );
    let doc = Package::from_reader(Cursor::new(&after)).unwrap();
    assert_eq!(
        doc.document().unwrap().paragraphs().unwrap()[10]
            .text()
            .unwrap(),
        "Reference link"
    );
    unchanged_except(
        &original,
        &after,
        &["word/_rels/document.xml.rels", "word/document.xml"],
    );
    assert!(members(&after)["word/document.xml"] == detached_document(&original));
    let mut expected = relationships(&members(&original)["word/_rels/document.xml.rels"]);
    assert!(expected.remove("rId11").is_some());
    assert_eq!(
        relationships(&members(&after)["word/_rels/document.xml.rels"]),
        expected
    );
    let noop = p.edit_external_hyperlink_redaction(&[]).unwrap();
    assert!(noop.effect_report().is_noop());
    assert!(bytes(noop.commit().unwrap().package()) == original);
    assert!(
        matches!(read.apply_external_hyperlink_redaction(&commit.patch().inverse()),
        Err(litchi_docx::Error::Refused(r)) if r.kind()==litchi_core::edit::RefusalKind::UnsupportedOperation)
    );
    let restored = commit
        .package()
        .apply_external_hyperlink_redaction(&commit.patch().inverse())
        .unwrap();
    assert!(bytes(restored.package()) == original);
    assert!(bytes(&p) == original);
}
#[test]
fn native_secondary_story_text_edits_preserve_anchors_siblings_and_other_parts() {
    use source_backed::story_text::Selector;
    for (name, selector, part, before, after) in [
        (
            "word-text-structure.docx",
            Selector::Footnote(1),
            "word/footnotes.xml",
            " Native footnote text: expected value 42.",
            " Native footnote text: expected value 43.",
        ),
        (
            "word-tables-layout.docx",
            Selector::Endnote(1),
            "word/endnotes.xml",
            " ENDNOTE_TEXT_002: native endnote value 84.",
            " ENDNOTE_TEXT_002: native endnote value 85.",
        ),
        (
            "word-tables-layout.docx",
            Selector::Header(0),
            "word/header1.xml",
            "HEADER_DEFAULT_002",
            "HEADER_CHANGED_002",
        ),
        (
            "word-review-controls.docx",
            Selector::Comment(5),
            "word/comments.xml",
            "MODERN_COMMENT_003: review fixture comment.",
            "MODERN_COMMENT_003: edited fixture comment.",
        ),
    ] {
        let original = native(name);
        let p = open(&original);
        let s = p.story_text_snapshot(selector.clone()).unwrap();
        assert_eq!(s.paragraph_text(0).unwrap().as_deref(), Some(before));
        let mut edit = p.edit_story_text(selector.clone()).unwrap();
        edit.replace_paragraph_text(Position::new(0), after)
            .unwrap();
        let commit = edit.commit().unwrap();
        let published = save(commit.package());
        let read = open(&published);
        assert_eq!(
            read.story_text_snapshot(selector)
                .unwrap()
                .paragraph_text(0)
                .unwrap()
                .as_deref(),
            Some(after)
        );
        unchanged_except(&original, &published, &[part]);
        let a = members(&original);
        let b = members(&published);
        let slots = if before.starts_with(' ') { 2 } else { 1 };
        assert_eq!(
            mask_first_text_slots(&a[part], slots),
            mask_first_text_slots(&b[part], slots),
            "changed XML outside owned text slots: {name} {part}"
        );
        let restored = read.apply_story_text(&commit.patch().inverse()).unwrap();
        unchanged_except(&original, &bytes(restored.package()), &[]);
        assert!(bytes(&p) == original);
    }
}
