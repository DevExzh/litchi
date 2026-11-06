//! Native Word revision dispositions, with immutable inputs and inert metadata.
use litchi_core::Position;
use litchi_docx::{Package, document::RevisionKind};
use std::{collections::BTreeMap, io::Cursor};

fn native(name: &str) -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/word")
            .join(name),
    )
    .unwrap()
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
// Replace one direct body paragraph with a marker so all XML outside the
// selected revision owner can be compared byte-for-byte.
fn outside_paragraph(xml: &[u8], wanted: usize) -> Vec<u8> {
    use quick_xml::events::Event;
    let mut reader = quick_xml::Reader::from_reader(xml);
    let mut depth = 0;
    let mut index = 0;
    let mut start = None;
    loop {
        let before = usize::try_from(reader.buffer_position()).unwrap();
        let event = reader.read_event().unwrap();
        let after = usize::try_from(reader.buffer_position()).unwrap();
        match event {
            Event::Start(e) => {
                if depth == 2 && e.local_name().as_ref() == b"p" {
                    if index == wanted {
                        start = Some(before);
                    }
                    index += 1;
                }
                depth += 1;
            },
            Event::Empty(e) if depth == 2 && e.local_name().as_ref() == b"p" => {
                if index == wanted {
                    return [
                        xml[..before].to_vec(),
                        b"<tested/>".to_vec(),
                        xml[after..].to_vec(),
                    ]
                    .concat();
                }
                index += 1;
            },
            Event::End(_) => {
                if depth == 3
                    && let Some(start) = start
                {
                    return [
                        xml[..start].to_vec(),
                        b"<tested/>".to_vec(),
                        xml[after..].to_vec(),
                    ]
                    .concat();
                }
                depth -= 1;
            },
            Event::Eof => panic!("paragraph {wanted} is absent"),
            _ => {},
        }
    }
}
#[test]
fn native_inline_revision_dispositions_preserve_property_changes_and_review_parts() {
    let bytes = native("word-review-controls.docx");
    let package = Package::from_reader(Cursor::new(&bytes)).unwrap();
    for (insert, delete, expected) in [
        (true, true, "REVIEW_INSERT_003"),
        (false, false, "REVIEW_DELETE_003"),
        (true, false, "REVIEW_INSERT_003REVIEW_DELETE_003"),
        (false, true, ""),
    ] {
        let mut edit = package.edit().unwrap();
        for (kind, accept) in [
            (RevisionKind::Insertion, insert),
            (RevisionKind::Deletion, delete),
        ] {
            if accept {
                edit.accept_revision(Position::new(3), kind, Position::new(0))
                    .unwrap();
            } else {
                edit.reject_revision(Position::new(3), kind, Position::new(0))
                    .unwrap();
            }
        }
        let commit = edit.commit().unwrap();
        let file = tempfile::NamedTempFile::with_suffix(".docx").unwrap();
        commit.package().save(file.path()).unwrap();
        let after = std::fs::read(file.path()).unwrap();
        let read = Package::from_reader(Cursor::new(&after)).unwrap();
        let doc = read.document().unwrap();
        let paragraphs = doc.paragraphs().unwrap();
        assert_eq!(paragraphs.len(), 9);
        assert_eq!(paragraphs[3].text().unwrap(), expected);
        assert_eq!(paragraphs[2].text().unwrap(), "REVIEW_ORIGINAL_003");
        let revisions = doc.revisions().unwrap();
        assert_eq!(revisions.len(), 2);
        for (revision, id) in revisions.iter().zip(["0", "1"]) {
            assert_eq!(revision.id(), id);
            assert_eq!(
                revision.revision_type(),
                litchi_docx::revision::RevisionType::FormatChange
            );
            assert_eq!(revision.author(), Some("Ryker Zhu"));
            assert_eq!(revision.date(), Some("2026-10-15T13:49:00Z"));
            assert_eq!(revision.date_utc(), Some("2026-10-15T05:49:00Z"));
        }
        unchanged_except(&bytes, &after, &["word/document.xml"]);
        assert_eq!(
            outside_paragraph(&members(&bytes)["word/document.xml"], 3),
            outside_paragraph(&members(&after)["word/document.xml"], 3)
        );
        let inverse = read
            .apply(&commit.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap();
        unchanged_except(&bytes, &inverse, &[]);
        assert!(package.to_bytes().unwrap() == bytes);
    }
}
#[test]
fn native_property_and_paragraph_move_owners_expose_their_operation_boundary() {
    for (name, reason) in [
        (
            "word-review-controls.docx",
            "tracked revision property changes are outside the authoring owner",
        ),
        (
            "word-tracked-moves.docx",
            "tracked revision wrapper has an unsupported paragraph or property owner",
        ),
    ] {
        let bytes = native(name);
        let p = Package::from_reader(Cursor::new(&bytes)).unwrap();
        // The semantic reader recognizes all four native revisions. The
        // broader family action owner intentionally excludes these shapes;
        // it must fail before any disposition is staged.
        assert_eq!(p.document().unwrap().revisions().unwrap().len(), 4);
        assert!(
            matches!(p.tracked_revisions(),Err(litchi_docx::Error::InvalidFormat(s)) if s==reason)
        );
        for accept in [true, false] {
            let mut edit = p.edit_tracked_revisions();
            let result = if accept {
                edit.accept_all()
            } else {
                edit.reject_all()
            };
            assert!(matches!(result,Err(litchi_docx::Error::InvalidFormat(s)) if s==reason));
            assert!(edit.commit().unwrap().package().to_bytes().unwrap() == bytes);
        }
        assert!(p.to_bytes().unwrap() == bytes);
    }
}
