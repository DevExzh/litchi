#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "differential assertions intentionally panic on fixture failures"
)]

//! Change 0765: a DOCX part that begins with a UTF-8 byte-order mark reads,
//! edits and saves exactly like the same part without one.
//!
//! quick-xml drops a leading mark before its first event without counting it
//! in its positions, so every span taken from a reader was three bytes early
//! in a marked part. Each test runs one scenario on a package and on its twin
//! whose XML members carry a mark, and requires identical observations and
//! an output that differs only by marks: every member of the marked output is
//! the unmarked output's member, or that member behind one mark, and a mark
//! appears only on a member whose input had one.

use std::collections::{BTreeMap, BTreeSet};
use std::io::Cursor;

use litchi_core::Position;
use litchi_docx::Package;
use soapberry_zip::office::{ArchiveReader, StreamingArchiveWriter};

const BOM: &[u8] = b"\xEF\xBB\xBF";
const ALT_CHUNK_HEADER: &str = concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../test-data/libreoffice-core/sw/qa/writerfilter/dmapper/data/alt-chunk-header.docx"
);

fn is_xml_member(name: &str) -> bool {
    name.ends_with(".xml") || name.ends_with(".rels")
}

fn members(archive: &[u8]) -> BTreeMap<String, Vec<u8>> {
    let reader = ArchiveReader::new(archive).unwrap();
    let names: Vec<String> = reader.file_names().map(str::to_owned).collect();
    names
        .into_iter()
        .map(|name| {
            let bytes = reader.read(&name).unwrap();
            (name, bytes)
        })
        .collect()
}

/// Indent every start tag that directly follows a tag: whitespace-only text
/// between elements, never inside an element that holds only text.
fn indent(xml: &[u8]) -> Vec<u8> {
    let mut output = Vec::with_capacity(xml.len() + xml.len() / 8);
    for (index, byte) in xml.iter().enumerate() {
        if *byte == b'<'
            && index > 0
            && xml[index - 1] == b'>'
            && xml
                .get(index + 1)
                .is_some_and(|next| next.is_ascii_alphabetic())
        {
            output.extend_from_slice(b"\n    ");
        }
        output.push(*byte);
    }
    output
}

/// Rebuild `archive` with every member `mark` selects behind one mark and
/// every other member without one, XML members indented when `indented`.
fn rebuild(
    archive: &[u8],
    mark: &dyn Fn(&str) -> bool,
    indented: bool,
) -> (Vec<u8>, BTreeSet<String>) {
    let mut writer = StreamingArchiveWriter::new();
    let mut marked = BTreeSet::new();
    for (name, bytes) in members(archive) {
        let body = bytes.strip_prefix(BOM).unwrap_or(&bytes);
        let body = if indented && is_xml_member(&name) {
            indent(body)
        } else {
            body.to_vec()
        };
        let bytes = if mark(&name) {
            marked.insert(name.clone());
            [BOM, body.as_slice()].concat()
        } else {
            body
        };
        writer.write_stored(&name, &bytes).unwrap();
    }
    (writer.finish_to_bytes().unwrap(), marked)
}

/// Require the marked output to equal the unmarked output except for marks.
/// Returns the members that kept a mark.
fn assert_equal_but_marks(
    marked: &[u8],
    plain: &[u8],
    marked_input: &BTreeSet<String>,
) -> BTreeSet<String> {
    let marked = members(marked);
    let plain = members(plain);
    assert_eq!(
        marked.keys().collect::<Vec<_>>(),
        plain.keys().collect::<Vec<_>>()
    );
    let mut kept = BTreeSet::new();
    for (name, bytes) in &marked {
        let unmarked = &plain[name];
        assert!(
            !unmarked.starts_with(BOM),
            "{name}: unmarked output gained a mark"
        );
        if let Some(rest) = bytes.strip_prefix(BOM) {
            assert!(
                marked_input.contains(name),
                "{name}: a mark appeared on a member whose input had none"
            );
            assert_eq!(rest, unmarked.as_slice(), "{name}: differs beyond its mark");
            kept.insert(name.clone());
        } else {
            assert_eq!(bytes, unmarked, "{name}: differs");
        }
    }
    kept
}

fn save(package: &mut Package) -> Vec<u8> {
    let mut output = Vec::new();
    package.to_stream(&mut output).unwrap();
    output
}

/// Run `scenario` on the unmarked and the marked twin of `archive`, as given
/// and indented, compare the observations and the saved packages, and return
/// the members that kept a mark (compact run) with the marked input's members.
fn differential(
    archive: &[u8],
    mark: impl Fn(&str) -> bool,
    scenario: impl Fn(&mut Package) -> String,
) -> (BTreeSet<String>, BTreeSet<String>) {
    let mut result = (BTreeSet::new(), BTreeSet::new());
    for indented in [true, false] {
        let (plain, _) = rebuild(archive, &|_| false, indented);
        let (marked, marked_input) = rebuild(archive, &mark, indented);
        assert!(!marked_input.is_empty(), "the scenario must mark a member");
        let mut plain_package = Package::from_reader(Cursor::new(plain)).unwrap();
        let mut marked_package = Package::from_reader(Cursor::new(marked)).unwrap();
        let plain_observed = scenario(&mut plain_package);
        let marked_observed = scenario(&mut marked_package);
        assert_eq!(marked_observed, plain_observed, "indented: {indented}");
        let plain_out = save(&mut plain_package);
        let marked_out = save(&mut marked_package);
        let kept = assert_equal_but_marks(&marked_out, &plain_out, &marked_input);
        result = (kept, marked_input);
    }
    result
}

fn generated_document() -> Vec<u8> {
    let mut package = Package::new().unwrap();
    {
        let document = package.document_mut().unwrap();
        for index in 0..6 {
            document.add_paragraph_with_text(&format!("paragraph {index} & <text>"));
        }
        document.add_heading("Heading", 1).unwrap();
        document
            .add_header_paragraph()
            .add_run_with_text("header text");
        let (_, note) = document.add_footnote();
        note.add_paragraph_with_text("footnote text");
        document.add_comment("Author", "comment text");
    }
    save(&mut package)
}

fn observe(package: &Package) -> String {
    let document = package.document().unwrap();
    let mut observed = format!(
        "{}|{}|",
        document.paragraph_count().unwrap(),
        document.text().unwrap()
    );
    for paragraph in document.paragraphs().unwrap() {
        observed.push_str(&format!("{:?};", paragraph.text()));
    }
    observed
}

#[test]
fn a_fully_marked_document_reads_and_saves_like_the_unmarked_document() {
    let (kept, marked) = differential(&generated_document(), is_xml_member, |package| {
        observe(package)
    });
    assert_eq!(kept, marked, "an unedited save preserves every mark");
}

#[test]
fn a_semantic_edit_of_a_marked_main_part_matches_the_unmarked_edit() {
    let (kept, _) = differential(&generated_document(), is_xml_member, |package| {
        let mut edit = package.edit_document().unwrap();
        edit.replace_paragraph_text(Position::new(1), "edited & <after>")
            .unwrap();
        edit.replace_paragraph_text(Position::new(4), "").unwrap();
        let commit = package.publish_document_edit(edit).unwrap();
        format!("{}|{}", commit.patch().changed(), observe(package))
    });
    assert!(kept.contains("word/document.xml"), "{kept:?}");
}

#[test]
fn a_mutable_edit_of_a_marked_main_part_matches_the_unmarked_edit() {
    let (kept, _) = differential(&generated_document(), is_xml_member, |package| {
        package
            .document_mut()
            .unwrap()
            .add_paragraph_with_text("appended");
        observe(package)
    });
    assert!(kept.contains("word/document.xml"), "{kept:?}");
}

#[test]
fn the_marked_alt_chunk_fixture_reads_and_edits_like_its_unmarked_twin() {
    // Word marks eleven members of this fixture. Its body keeps it off both
    // edit routes (0650's structural refusal); the refusals must be the
    // unmarked twin's, and a save must preserve every member.
    let original = std::fs::read(ALT_CHUNK_HEADER).unwrap();
    let marked_members: BTreeSet<String> = members(&original)
        .into_iter()
        .filter(|(_, bytes)| bytes.starts_with(BOM))
        .map(|(name, _)| name)
        .collect();
    assert!(marked_members.contains("word/document.xml"));
    let (kept, _) = differential(
        &original,
        |name| marked_members.contains(name),
        |package| {
            let before = observe(package);
            let semantic = package
                .edit_document()
                .map(|_| ())
                .map_err(|error| error.to_string());
            let mutable = package
                .document_mut()
                .map(|_| ())
                .map_err(|error| error.to_string());
            format!("{before}|{semantic:?}|{mutable:?}")
        },
    );
    assert_eq!(kept, marked_members, "a save preserves every mark");
}

#[test]
fn settings_patches_of_the_marked_fixture_match_its_unmarked_twin() {
    // Base: Word's marked, pretty-printed settings.xml of this fixture lost
    // three indentation bytes and kept `/>` of the old element as text on
    // `remove_attached_template` and `set_attached_template_uri`.
    let original = std::fs::read(ALT_CHUNK_HEADER).unwrap();
    let marked_members: BTreeSet<String> = members(&original)
        .into_iter()
        .filter(|(_, bytes)| bytes.starts_with(BOM))
        .map(|(name, _)| name)
        .collect();
    assert!(marked_members.contains("word/settings.xml"));
    let scenarios: [&dyn Fn(&mut Package) -> String; 3] = [
        &|package| format!("{:?}", package.remove_attached_template().map(|_| ())),
        &|package| {
            format!(
                "{:?}",
                package
                    .set_attached_template_uri("https://example.invalid/template.dotx")
                    .map(|_| ())
            )
        },
        &|package| {
            let first = package.set_document_variable("first", "one & <two>");
            let second = package.set_document_variable("second", "three");
            let removed = package.remove_document_variable("first");
            format!("{first:?}|{second:?}|{removed:?}")
        },
    ];
    for scenario in scenarios {
        let (kept, _) = differential(&original, |name| marked_members.contains(name), scenario);
        assert!(kept.contains("word/settings.xml"), "{kept:?}");
    }
}

#[test]
fn protection_edits_of_a_marked_settings_part_match_unmarked_edits() {
    differential(&generated_document(), is_xml_member, |package| {
        package
            .document_mut()
            .unwrap()
            .set_protection(litchi_docx::ProtectionType::ReadOnly);
        let first = observe(package);
        package.document_mut().unwrap().remove_protection();
        format!("{first}|{}", observe(package))
    });
}
