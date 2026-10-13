//! Structural DOCX operations through the format-neutral `File` facade.
//!
//! The fixture keeps direct body paragraphs, a table and an unrelated opaque
//! package member together.  Direct paragraphs and runs are the currently
//! admitted write surface; the table makes body-block ordinals observable and
//! the opaque member proves that a structural write does not rebuild away
//! unrelated XML.
#![cfg(feature = "docx")]

use std::io::{Cursor, Read as _, Write as _};

use litchi_core::Position;
use litchi_ops::value::{Color, Length};
use litchi_ops::{Code, File, NewNode, OpenOptions, Placement, SourceMode};

const WORD_MAIN: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
const OPAQUE_MEMBER: &str = "word/customXml/item1.xml";

fn package(document: &str) -> Vec<u8> {
    let content_types = format!(
        r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="{WORD_MAIN}"/></Types>"#
    );
    let root_relationships = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#;
    let opaque = r#"<opaque xmlns="urn:litchi:opaque" marker="preserve-me"><payload>opaque XML must survive structural edits</payload></opaque>"#;

    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let options =
        zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
    for (name, contents) in [
        ("[Content_Types].xml", content_types.as_str()),
        ("_rels/.rels", root_relationships),
        ("word/document.xml", document),
        (OPAQUE_MEMBER, opaque),
    ] {
        zip.start_file(name, options).expect("start fixture member");
        zip.write_all(contents.as_bytes())
            .expect("write fixture member");
    }
    zip.finish().expect("finish fixture package").into_inner()
}

/// Direct body paragraphs are interleaved with a table.  The first paragraph
/// has two runs so run movement and swapping can be checked independently from
/// paragraph movement.  The table's paragraph is intentionally not a direct
/// body paragraph: `/body/p[N]` must not count it.
fn fixture() -> Vec<u8> {
    package(
        r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:paraId="11111111"><w:pPr><w:pStyle w:val="Heading1"/><w:jc w:val="left"/></w:pPr><w:r><w:rPr><w:b w:val="0"/></w:rPr><w:proofErr w:type="spellStart"/><w:t>One</w:t></w:r><w:r><w:t>A</w:t><w:proofErr w:type="spellEnd"/></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p w14:paraId="22222222"><w:r><w:t>Two</w:t></w:r></w:p><w:p w14:paraId="33333333"><w:r><w:t>Drop</w:t></w:r></w:p><w:p w14:paraId="44444444"><w:r><w:t>Drop</w:t></w:r></w:p><w:sectPr/></w:body></w:document>"#,
    )
}

fn interspersed_blocks_fixture() -> Vec<u8> {
    package(
        r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>A</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>T</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p><w:r><w:t>B</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>U</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p><w:r><w:t>C</w:t></w:r></w:p><w:sectPr/></w:body></w:document>"#,
    )
}

fn three_run_fixture() -> Vec<u8> {
    package(
        r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>A</w:t></w:r><w:r><w:t>B</w:t></w:r><w:r><w:t>C</w:t></w:r></w:p><w:sectPr/></w:body></w:document>"#,
    )
}

fn open_source() -> File {
    litchi_ops::open_bytes(fixture()).expect("source-backed fixture")
}

fn open_eager() -> File {
    litchi_ops::open_bytes_with(
        fixture(),
        &OpenOptions::new().source_mode(SourceMode::Eager),
    )
    .expect("eager fixture")
}

fn owners() -> [(&'static str, fn() -> File); 2] {
    [("source", open_source), ("eager", open_eager)]
}

fn open_mode(bytes: Vec<u8>, eager: bool) -> File {
    if eager {
        litchi_ops::open_bytes_with(bytes, &OpenOptions::new().source_mode(SourceMode::Eager))
            .expect("reopen eager structural result")
    } else {
        litchi_ops::open_bytes(bytes).expect("reopen source structural result")
    }
}

fn empty_destination_fixture() -> Vec<u8> {
    package(
        r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Source</w:t></w:r></w:p><w:p/><w:sectPr/></w:body></w:document>"#,
    )
}

fn alternate_namespace_fixture() -> Vec<u8> {
    package(
        r#"<x:document xmlns:x="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><x:body><x:p><x:r><x:t>Prefix</x:t></x:r></x:p><x:p><x:r><x:t>Second</x:t></x:r></x:p><x:sectPr/></x:body></x:document>"#,
    )
}

fn text(file: &File, path: &str) -> String {
    file.get(path)
        .unwrap_or_else(|error| panic!("{path}: {error}"))
        .view()
        .text()
        .unwrap_or_else(|error| panic!("{path} text: {error}"))
        .expect("text value")
        .into_owned()
}

fn cell_text(file: &File) -> String {
    text(file, "/body/tbl[1]/tr[1]/tc[1]/p[1]")
}

fn direct_paragraphs(file: &File) -> Vec<String> {
    (1..=16)
        .filter_map(|ordinal| {
            let path = format!("/body/p[{ordinal}]");
            file.find(&path)
                .unwrap_or_else(|error| panic!("find {ordinal}: {error}"))
                .map(|node| {
                    node.view()
                        .text()
                        .unwrap_or_else(|error| panic!("paragraph {ordinal}: {error}"))
                        .expect("paragraph text")
                        .into_owned()
                })
        })
        .collect()
}

fn has_format(file: &File, path: &str, key: &str, wanted: &str) -> bool {
    file.get(path)
        .unwrap_or_else(|error| panic!("{path}: {error}"))
        .view()
        .format()
        .unwrap_or_else(|error| panic!("{path} format: {error}"))
        .any(|property| property.key() == key && property.to_string() == wanted)
}

fn opaque_member(bytes: &[u8]) -> Vec<u8> {
    member(bytes, OPAQUE_MEMBER)
}

fn member(bytes: &[u8], name: &str) -> Vec<u8> {
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("open DOCX ZIP");
    let mut member = archive.by_name(name).expect("ZIP member present");
    let mut contents = Vec::new();
    member.read_to_end(&mut contents).expect("read ZIP member");
    contents
}

fn document_xml(bytes: &[u8]) -> Vec<u8> {
    member(bytes, "word/document.xml")
}

#[test]
fn run_swaps_keep_opaque_sibling_slots_within_and_between_paragraphs() {
    let bytes = package(
        r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:z="urn:opaque"><w:body><w:p><w:r><w:t>A</w:t></w:r><z:keep/><w:r><w:t>B</w:t></w:r><w:r><w:t>C</w:t></w:r></w:p><w:p><w:r><w:t>D</w:t></w:r><z:keep2/><w:r><w:t>E</w:t></w:r></w:p></w:body></w:document>"#,
    );
    for eager in [false, true] {
        let original = open_mode(bytes.clone(), eager);
        let same = original
            .swap("/body/p[1]/r[1]", "/body/p[1]/r[3]")
            .expect("same paragraph swap");
        assert_xml_order(
            &document_xml(&same.to_bytes().expect("publish same")),
            &[b">C</w:t>", b"<z:keep/>", b">B</w:t>", b">A</w:t>"],
        );
        let cross = original
            .swap("/body/p[1]/r[1]", "/body/p[2]/r[1]")
            .expect("cross paragraph swap");
        assert_xml_order(
            &document_xml(&cross.to_bytes().expect("publish cross")),
            &[
                b">D</w:t>",
                b"<z:keep/>",
                b">B</w:t>",
                b">C</w:t>",
                b">A</w:t>",
                b"<z:keep2/>",
                b">E</w:t>",
            ],
        );
        assert_eq!(original.to_bytes().expect("original"), bytes);
    }
}

fn assert_xml_order(bytes: &[u8], markers: &[&[u8]]) {
    let mut previous = 0;
    for marker in markers {
        let position = bytes
            .windows(marker.len())
            .position(|window| window == *marker)
            .unwrap_or_else(|| {
                panic!("XML marker {:?} is absent", String::from_utf8_lossy(marker))
            });
        assert!(
            position >= previous,
            "XML marker {:?} precedes the previous marker",
            String::from_utf8_lossy(marker)
        );
        previous = position;
    }
}

fn assert_one_of(error: litchi_ops::Error, expected: &[Code]) {
    assert!(
        expected.contains(&error.code()),
        "unexpected structural error code {:?}: {error}",
        error.code()
    );
}

fn assert_code(error: litchi_ops::Error, expected: Code) {
    assert_eq!(
        error.code(),
        expected,
        "unexpected structural error: {error}"
    );
}

#[test]
fn add_paragraph_and_run_reports_canonical_paths_and_properties_for_both_owners() {
    for (owner, open) in owners() {
        let original = open();
        let clone = original.clone();
        let original_bytes = original.to_bytes().expect("publish original");

        let added = original
            .add(
                "/body",
                NewNode::new("paragraph")
                    .prop("text", "Inserted")
                    .prop("style", "Title")
                    .prop("align", "right")
                    .prop("bold", true)
                    .prop("size", Length::from_points(14))
                    .prop("font.latin", "Arial")
                    .prop("color", Color::rgb(0x1F, 0x4E, 0x79))
                    .at(Placement::Start),
            )
            .unwrap_or_else(|error| panic!("{owner} paragraph add: {error}"));
        assert_eq!(added.path.to_string(), "/body/p[1]");
        let with_paragraph = added.into_file();
        assert_eq!(text(&with_paragraph, "/body/p[1]"), "Inserted");
        assert_eq!(
            with_paragraph
                .get("/body/p[1]")
                .expect("added paragraph")
                .view()
                .style()
                .expect("paragraph style"),
            Some("Title".into())
        );
        assert!(has_format(&with_paragraph, "/body/p[1]", "align", "right"));
        for (key, wanted) in [
            ("bold", "true"),
            ("size", "14pt"),
            ("font.latin", "Arial"),
            ("color", "#1F4E79"),
        ] {
            assert!(
                has_format(&with_paragraph, "/body/p[1]/r[1]", key, wanted),
                "missing initial-run {key}={wanted}"
            );
        }

        let added_run = with_paragraph
            .add(
                "/body/p[2]",
                NewNode::new("run")
                    .prop("text", "Inserted run")
                    .prop("bold", true)
                    .at(Placement::Position(Position::new(1))),
            )
            .unwrap_or_else(|error| panic!("{owner} run add: {error}"));
        assert_eq!(added_run.path.to_string(), "/body/p[2]/r[2]");
        let updated = added_run.into_file();
        assert_eq!(text(&updated, "/body/p[2]/r[2]"), "Inserted run");
        assert!(has_format(&updated, "/body/p[2]/r[2]", "bold", "true"));

        assert_eq!(
            original.to_bytes().expect("republish original"),
            original_bytes
        );
        assert_eq!(clone.to_bytes().expect("republish clone"), original_bytes);
        assert_eq!(direct_paragraphs(&original), direct_paragraphs(&clone));

        let reopened = open_mode(
            updated.to_bytes().expect("publish additions"),
            owner == "eager",
        );
        assert_eq!(text(&reopened, "/body/p[1]"), "Inserted");
        assert_eq!(text(&reopened, "/body/p[2]/r[2]"), "Inserted run");
    }
}

#[test]
fn paragraph_move_uses_direct_body_ordinals_and_zero_based_detached_positions() {
    for (owner, open) in owners() {
        let original = open();
        assert_eq!(cell_text(&original), "Cell");
        assert_eq!(text(&original, "/body/p[2]"), "Two");
        let original_bytes = original.to_bytes().expect("publish original");

        let before_table = original
            .move_to(
                "/body/p[3]",
                Placement::before("/body/tbl[1]").expect("table anchor"),
            )
            .unwrap_or_else(|error| panic!("{owner} move before table: {error}"));
        assert_eq!(
            direct_paragraphs(&before_table),
            ["OneA", "Drop", "Two", "Drop"]
        );
        assert_eq!(cell_text(&before_table), "Cell");

        let first_after_detach = original
            .move_to("/body/p[3]", Placement::Position(Position::new(0)))
            .unwrap_or_else(|error| panic!("{owner} zero-based move: {error}"));
        assert_eq!(
            direct_paragraphs(&first_after_detach),
            ["Drop", "OneA", "Two", "Drop"]
        );

        let last = original
            .move_to("/body/p[1]", Placement::End)
            .unwrap_or_else(|error| panic!("{owner} move end: {error}"));
        assert_eq!(direct_paragraphs(&last), ["Two", "Drop", "Drop", "OneA"]);
        assert_eq!(
            original.to_bytes().expect("republish original"),
            original_bytes
        );
    }
}

#[test]
fn run_move_and_swap_preserve_opaque_inline_xml_and_original_snapshots() {
    for (owner, open) in owners() {
        let original = open();
        let clone = original.clone();
        let opaque = opaque_member(&original.to_bytes().expect("publish original"));

        let moved = original
            .move_to("/body/p[1]/r[1]", Placement::End)
            .unwrap_or_else(|error| panic!("{owner} run move: {error}"));
        assert_eq!(text(&moved, "/body/p[1]"), "AOne");
        assert_eq!(
            opaque_member(&moved.to_bytes().expect("publish run move")),
            opaque
        );
        let moved_xml = document_xml(&moved.to_bytes().expect("republish run move"));
        assert!(
            moved_xml
                .windows(b"w:proofErr".len())
                .any(|window| window == b"w:proofErr")
        );

        let swapped = moved
            .swap("/body/p[1]/r[1]", "/body/p[1]/r[2]")
            .unwrap_or_else(|error| panic!("{owner} run swap: {error}"));
        assert_eq!(text(&swapped, "/body/p[1]"), "OneA");
        assert_eq!(
            opaque_member(&swapped.to_bytes().expect("publish run swap")),
            opaque
        );

        let with_third_run = original
            .add(
                "/body/p[1]",
                NewNode::new("run").prop("text", "B").at(Placement::End),
            )
            .unwrap_or_else(|error| panic!("{owner} third-run add: {error}"))
            .into_file();
        let nonadjacent_run_swap = with_third_run
            .swap("/body/p[1]/r[1]", "/body/p[1]/r[3]")
            .unwrap_or_else(|error| panic!("{owner} nonadjacent run swap: {error}"));
        assert_eq!(text(&nonadjacent_run_swap, "/body/p[1]"), "BAOne");
        assert_eq!(
            opaque_member(
                &nonadjacent_run_swap
                    .to_bytes()
                    .expect("publish nonadjacent run swap")
            ),
            opaque
        );
        let nonadjacent_run_xml = document_xml(
            &nonadjacent_run_swap
                .to_bytes()
                .expect("republish nonadjacent run swap"),
        );
        assert_xml_order(
            &nonadjacent_run_xml,
            &[b"<w:t>B</w:t>", b"<w:t>A</w:t>", b"<w:t>One</w:t>"],
        );
        assert!(
            nonadjacent_run_xml
                .windows(b"w:proofErr".len())
                .any(|window| window == b"w:proofErr")
        );

        let cross_moved = original
            .move_to(
                "/body/p[1]/r[1]",
                Placement::before("/body/p[2]/r[1]").expect("cross-run anchor"),
            )
            .unwrap_or_else(|error| panic!("{owner} cross-paragraph run move: {error}"));
        assert_eq!(text(&cross_moved, "/body/p[1]"), "A");
        assert_eq!(text(&cross_moved, "/body/p[2]"), "OneTwo");
        let cross_swapped = cross_moved
            .swap("/body/p[1]/r[1]", "/body/p[2]/r[1]")
            .unwrap_or_else(|error| panic!("{owner} cross-paragraph run swap: {error}"));
        assert_eq!(text(&cross_swapped, "/body/p[1]"), "One");
        assert_eq!(text(&cross_swapped, "/body/p[2]"), "ATwo");
        assert_eq!(
            opaque_member(&cross_swapped.to_bytes().expect("publish cross-run swap")),
            opaque
        );
        let cross_xml = document_xml(&cross_swapped.to_bytes().expect("republish cross-run swap"));
        assert!(
            cross_xml
                .windows(b"w:proofErr".len())
                .any(|window| window == b"w:proofErr")
        );

        assert_eq!(text(&original, "/body/p[1]"), "OneA");
        assert_eq!(text(&clone, "/body/p[1]"), "OneA");
    }
}

#[test]
fn paragraph_swap_and_remove_query_fanout_are_atomic_and_reverse_ordered() {
    for (owner, open) in owners() {
        let original = open();
        let original_bytes = original.to_bytes().expect("publish original");

        let swapped = original
            .swap("/body/p[1]", "/body/p[3]")
            .unwrap_or_else(|error| panic!("{owner} paragraph swap: {error}"));
        assert_eq!(direct_paragraphs(&swapped), ["Drop", "Two", "OneA", "Drop"]);

        let removed = original
            .remove("/body/p[text='Drop']")
            .unwrap_or_else(|error| panic!("{owner} scoped removal: {error}"));
        assert_eq!(direct_paragraphs(&removed), ["OneA", "Two"]);
        assert!(
            removed
                .query("/body/p[text='Drop']")
                .expect("query find")
                .next()
                .is_none()
        );
        assert_eq!(
            original.to_bytes().expect("republish original"),
            original_bytes
        );

        let removed_run = original
            .remove("/body/p[1]/r[2]")
            .unwrap_or_else(|error| panic!("{owner} run removal: {error}"));
        assert_eq!(text(&removed_run, "/body/p[1]"), "One");
        assert_eq!(
            direct_paragraphs(&original),
            ["OneA", "Two", "Drop", "Drop"]
        );
    }
}

#[test]
fn nonadjacent_paragraph_swap_keeps_interspersed_table_body_order() {
    for (owner, eager) in [("source", false), ("eager", true)] {
        let original = open_mode(interspersed_blocks_fixture(), eager);
        assert_eq!(
            original.text().expect("interspersed body text"),
            "A\nT\nB\nU\nC"
        );

        let swapped = original
            .swap("/body/p[1]", "/body/p[3]")
            .unwrap_or_else(|error| panic!("{owner} nonadjacent body swap: {error}"));
        assert_eq!(
            swapped.text().expect("swapped interspersed body text"),
            "C\nT\nB\nU\nA"
        );
        let bytes = swapped.to_bytes().expect("publish interspersed body swap");
        assert_xml_order(
            &document_xml(&bytes),
            &[
                b"<w:t>C</w:t>",
                b"<w:t>T</w:t>",
                b"<w:t>B</w:t>",
                b"<w:t>U</w:t>",
                b"<w:t>A</w:t>",
            ],
        );
        let reopened = open_mode(bytes, eager);
        assert_eq!(
            reopened.text().expect("reopen interspersed body swap"),
            "C\nT\nB\nU\nA"
        );
    }
}

#[test]
fn run_position_is_zero_based_after_detaching_the_source() {
    for (owner, eager) in [("source", false), ("eager", true)] {
        let original = open_mode(three_run_fixture(), eager);
        assert_eq!(text(&original, "/body/p[1]"), "ABC");

        let moved = original
            .move_to("/body/p[1]/r[1]", Placement::Position(Position::new(1)))
            .unwrap_or_else(|error| panic!("{owner} detached run position: {error}"));
        assert_eq!(text(&moved, "/body/p[1]"), "BAC");
        assert_eq!(text(&original, "/body/p[1]"), "ABC");
    }
}

#[test]
fn empty_self_closing_paragraphs_accept_run_add_and_cross_parent_move() {
    for (owner, eager) in [("source", false), ("eager", true)] {
        let original = open_mode(empty_destination_fixture(), eager);
        let moved = original
            .move_to(
                "/body/p[1]/r[1]",
                Placement::inside("/body/p[2]").expect("empty paragraph destination"),
            )
            .unwrap_or_else(|error| panic!("{owner} move into empty paragraph: {error}"));
        assert_eq!(text(&moved, "/body/p[2]"), "Source");
        assert_eq!(
            moved
                .get("/body/p[1]")
                .expect("empty source paragraph")
                .view()
                .text()
                .expect("empty source text")
                .map(|value| value.into_owned())
                .unwrap_or_default(),
            ""
        );

        let added = moved
            .add(
                "/body/p[1]",
                NewNode::new("run").prop("text", "After").at(Placement::End),
            )
            .unwrap_or_else(|error| panic!("{owner} add to emptied paragraph: {error}"));
        assert_eq!(added.path.to_string(), "/body/p[1]/r[1]");
        let updated = added.into_file();
        assert_eq!(text(&updated, "/body/p[1]"), "After");
        assert_eq!(text(&updated, "/body/p[2]"), "Source");
    }
}

#[test]
fn alternate_word_namespace_prefixes_keep_structural_edits_lossless() {
    for (owner, eager) in [("source", false), ("eager", true)] {
        let original = open_mode(alternate_namespace_fixture(), eager);
        let added = original
            .add(
                "/body",
                NewNode::new("paragraph")
                    .prop("text", "Added")
                    .at(Placement::Start),
            )
            .unwrap_or_else(|error| panic!("{owner} prefixed add: {error}"))
            .into_file();
        let moved = added
            .move_to(
                "/body/p[2]",
                Placement::after("/body/p[1]").expect("prefixed paragraph anchor"),
            )
            .unwrap_or_else(|error| panic!("{owner} prefixed move: {error}"));
        assert_eq!(direct_paragraphs(&moved), ["Added", "Prefix", "Second"]);
        let bytes = moved.to_bytes().expect("publish prefixed structural edit");
        let xml = document_xml(&bytes);
        assert!(
            xml.windows(b"x:body".len())
                .any(|window| window == b"x:body")
        );
        assert!(
            !xml.windows(b"<w:body".len())
                .any(|window| window == b"<w:body")
        );
        let reopened = open_mode(bytes, eager);
        assert_eq!(direct_paragraphs(&reopened), ["Added", "Prefix", "Second"]);
    }
}

#[test]
fn invalid_indices_properties_overlaps_and_security_refuse_without_partial_edits() {
    let original = open_source();
    let original_bytes = original.to_bytes().expect("publish original");

    let bad_property = original.add(
        "/body",
        NewNode::new("paragraph")
            .prop("text", "must not commit")
            .prop("noSuchProperty", true),
    );
    assert_one_of(
        bad_property.expect_err("unknown add property"),
        &[
            Code::UnsupportedProperty,
            Code::InvalidValue,
            Code::InvalidArgument,
        ],
    );
    assert_eq!(
        original.to_bytes().expect("republish after bad property"),
        original_bytes
    );

    let bad_index = original.add(
        "/body",
        NewNode::new("paragraph")
            .prop("text", "must not commit")
            .at(Placement::Position(Position::new(99))),
    );
    assert_one_of(
        bad_index.expect_err("out-of-range add position"),
        &[
            Code::OutOfRange,
            Code::RangeTargetNotFound,
            Code::PathNotFound,
            Code::UnsupportedCapability,
        ],
    );

    let bad_remove = original.remove("/body/p[99]");
    assert_one_of(
        bad_remove.expect_err("out-of-range remove"),
        &[Code::PathNotFound, Code::OutOfRange, Code::NotFound],
    );

    let invalid_zero = original.remove("/body/p[0]");
    assert_one_of(
        invalid_zero.expect_err("zero is not a path position"),
        &[Code::InvalidPath, Code::OutOfRange, Code::PathNotFound],
    );

    let same_paragraph = original
        .move_to("/body/p[1]", Placement::Position(Position::new(0)))
        .expect("moving a paragraph to its detached index is a no-op");
    assert_eq!(
        same_paragraph.to_bytes().expect("publish same paragraph"),
        original_bytes
    );
    let same_run = original
        .move_to("/body/p[1]/r[1]", Placement::Position(Position::new(0)))
        .expect("moving a run to its own index is a no-op");
    assert_eq!(
        same_run.to_bytes().expect("publish same run"),
        original_bytes
    );
    let same_swap = original
        .swap("/body/p[1]", "/body/p[1]")
        .expect("swapping a node with itself is a no-op");
    assert_eq!(
        same_swap.to_bytes().expect("publish same swap"),
        original_bytes
    );

    let pruned = original
        .remove("/body/p[1], p[1] > r[1]")
        .unwrap_or_else(|error| panic!("ancestor selection should prune: {error:?}"));
    assert_eq!(direct_paragraphs(&pruned), ["Two", "Drop", "Drop"]);
    assert_eq!(
        original.to_bytes().expect("republish after pruned removal"),
        original_bytes
    );

    // Tables are removable through their element path (record 0251a).
    let without_table = original
        .remove("/body/tbl[1]")
        .unwrap_or_else(|error| panic!("table removal: {error}"));
    assert!(
        without_table
            .find("/body/tbl[1]")
            .expect("table lookup")
            .is_none()
    );
    assert_eq!(
        original.to_bytes().expect("republish after table removal"),
        original_bytes
    );

    let no_retain =
        litchi_ops::open_bytes_with(fixture(), &OpenOptions::new().retain_source(false))
            .expect("open non-retained fixture");
    let refusal = no_retain
        .remove("/body/p[1]")
        .expect_err("non-retained removal");
    assert_eq!(refusal.code(), Code::UnsupportedCapability);
    assert_eq!(
        no_retain.to_bytes().expect("publish non-retained original"),
        original_bytes
    );
}

#[test]
fn structural_verbs_honor_protection_retention_and_source_freshness() {
    let signed_path = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/poi/test-data/xmldsign/hello-world-signed.docx"
    );
    let signed = litchi_ops::open(signed_path).expect("signed fixture");
    assert_code(
        signed
            .add(
                "/body",
                NewNode::new("paragraph").prop("text", "must refuse"),
            )
            .expect_err("signed add"),
        Code::DocumentProtected,
    );
    assert_code(
        signed
            .move_to("/body/p[1]", Placement::End)
            .expect_err("signed move"),
        Code::DocumentProtected,
    );
    assert_code(
        signed
            .swap("/body/p[1]", "/body/p[1]")
            .expect_err("signed swap"),
        Code::DocumentProtected,
    );
    assert_code(
        signed.remove("/body/p[1]").expect_err("signed remove"),
        Code::DocumentProtected,
    );

    let unretained =
        litchi_ops::open_bytes_with(fixture(), &OpenOptions::new().retain_source(false))
            .expect("unretained fixture");
    assert_code(
        unretained
            .add(
                "/body",
                NewNode::new("paragraph").prop("text", "must refuse"),
            )
            .expect_err("unretained add"),
        Code::UnsupportedCapability,
    );
    assert_code(
        unretained
            .move_to("/body/p[1]", Placement::End)
            .expect_err("unretained move"),
        Code::UnsupportedCapability,
    );
    assert_code(
        unretained
            .swap("/body/p[1]", "/body/p[1]")
            .expect_err("unretained swap"),
        Code::UnsupportedCapability,
    );
    assert_code(
        unretained
            .remove("/body/p[1]")
            .expect_err("unretained remove"),
        Code::UnsupportedCapability,
    );

    for operation in ["add", "move", "swap", "remove"] {
        let directory = tempfile::tempdir().expect("stale source directory");
        let path = directory.path().join("stale.docx");
        std::fs::write(&path, fixture()).expect("write stale source");
        let stale = litchi_ops::open(&path).expect("path-backed fixture");
        std::fs::write(&path, b"source changed after open").expect("change source");
        let error = match operation {
            "add" => stale
                .add(
                    "/body",
                    NewNode::new("paragraph").prop("text", "must refuse"),
                )
                .expect_err("stale add"),
            "move" => stale
                .move_to("/body/p[1]", Placement::End)
                .expect_err("stale move"),
            "swap" => stale
                .swap("/body/p[1]", "/body/p[1]")
                .expect_err("stale swap"),
            "remove" => stale.remove("/body/p[1]").expect_err("stale remove"),
            _ => unreachable!("known structural operation"),
        };
        assert_code(error, Code::IoError);
    }
}

#[test]
fn opaque_member_survives_movement_publication_reopen_and_repeated_chains() {
    for (owner, open) in owners() {
        let original = open();
        let opaque = opaque_member(&original.to_bytes().expect("publish original"));

        let added = original
            .add(
                "/body",
                NewNode::new("paragraph")
                    .prop("text", "Chain")
                    .at(Placement::End),
            )
            .unwrap_or_else(|error| panic!("{owner} chain add: {error}"))
            .into_file();
        let moved = added
            .move_to("/body/p[1]", Placement::End)
            .unwrap_or_else(|error| panic!("{owner} chain move: {error}"));
        let swapped = moved
            .swap("/body/p[1]", "/body/p[2]")
            .unwrap_or_else(|error| panic!("{owner} chain swap: {error}"));
        let removed = swapped
            .remove("/body/p[2]")
            .unwrap_or_else(|error| panic!("{owner} chain remove: {error}"));
        let final_added = removed
            .add(
                "/body",
                NewNode::new("paragraph")
                    .prop("text", "Final")
                    .at(Placement::Position(Position::new(1))),
            )
            .unwrap_or_else(|error| panic!("{owner} chain second add: {error}"))
            .into_file();

        let bytes = final_added.to_bytes().expect("publish repeated chain");
        assert_eq!(opaque_member(&bytes), opaque);
        let reopened = open_mode(bytes, owner == "eager");
        assert!(direct_paragraphs(&reopened).contains(&"Final".to_owned()));
        assert_eq!(
            opaque_member(&reopened.to_bytes().expect("republish chain")),
            opaque
        );
    }
}
