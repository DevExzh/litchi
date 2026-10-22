//! Field instructions, bookmarks, tab stops and section page setup by
//! source splicing, and paragraph moves across section breaks (record
//! 0328).

use litchi_core::Position;
use litchi_rtf::edit::{Error, SectionValue};
use litchi_rtf::{Package, PageOrientation, SectionBreakType, TabAlignment, TabLeader, TabStop};

const PARTS: &[u8] = include_bytes!("../../../test-data/unified-ops/0294/parts.rtf");

fn source() -> Package {
    Package::from_bytes(PARTS).unwrap()
}

fn bytes(package: &Package) -> String {
    String::from_utf8(package.to_bytes().unwrap()).unwrap()
}

/// The source with exactly one byte range replaced.
fn assert_one_splice(before: &str, after: &str, old: &str, new: &str) {
    assert_eq!(before.matches(old).count(), 1, "{old} occurs once");
    assert_eq!(after, before.replacen(old, new, 1));
}

#[test]
fn a_field_instruction_set_splices_the_fldinst_text_only() {
    let source = source();
    let toc = source
        .fields()
        .iter()
        .position(|field| field.instruction.trim_start().starts_with("TOC"))
        .unwrap();
    let mut edit = source.edit();
    edit.set_field_instruction(toc, r#"TOC \o "1-2" \h"#)
        .unwrap();
    let commit = edit.commit().unwrap();
    let after = commit.package();
    assert_eq!(after.fields()[toc].instruction.trim(), r#"TOC \o "1-2" \h"#);
    assert_eq!(after.fields()[toc].result, source.fields()[toc].result);
    assert_one_splice(
        &bytes(&source),
        &bytes(after),
        r#"{ TOC \\o "1-3" \\h \\z }"#,
        r#"{ TOC \\o "1-2" \\h }"#,
    );
    // The patch replays durably.
    let limits = litchi_core::patch::PatchLimits::new(
        litchi_core::patch::BlobLimits::new(0, 0, 0),
        1024 * 1024,
        8,
        8,
        64 * 1024,
        256 * 1024,
    );
    let durable = commit.patch().to_durable(limits).unwrap();
    let replayed = source.apply_durable(&durable).unwrap().into_package();
    assert_eq!(bytes(&replayed), bytes(after));
}

#[test]
fn a_bookmark_renames_both_ends_and_removal_keeps_its_text() {
    let source = source();
    let mut edit = source.edit();
    edit.rename_bookmark(0, "Opening").unwrap();
    edit.remove_bookmark(1).unwrap();
    let after = edit.commit().unwrap().into_package();
    let names: Vec<_> = after
        .raw()
        .bookmarks()
        .bookmarks()
        .iter()
        .map(|bookmark| bookmark.name.to_string())
        .collect();
    assert_eq!(names, ["Opening"]);
    assert_eq!(after.text(), source.text());
    let text = bytes(&after);
    assert!(text.contains(r"{\*\bkmkstart Opening}Intro{\*\bkmkend Opening}"));
    assert!(!text.contains("Second}"));
    assert!(text.contains(r"\pard\plain Second paragraph\par"));
}

#[test]
fn an_inserted_bookmark_reads_back_at_its_paragraph() {
    let source = source();
    let mut edit = source.edit();
    edit.insert_bookmark(1, "Pages").unwrap();
    let after = edit.commit().unwrap().into_package();
    let table = after.raw().bookmarks();
    assert!(table.contains("Pages"));
    assert!(bytes(&after).contains(r"\pard\plain {\*\bkmkstart Pages}{\*\bkmkend Pages}Page "));
    let mut edit = source.edit();
    assert!(edit.insert_bookmark(1, "Intro").is_err(), "duplicate name");
    assert!(edit.insert_bookmark(1, "a{b").is_err(), "brace in name");
}

#[test]
fn tab_stops_are_replaced_in_place_and_removed() {
    let source = source();
    let stops = [
        TabStop {
            position: 720,
            alignment: TabAlignment::Center,
            leader: TabLeader::None,
        },
        TabStop {
            position: 2880,
            alignment: TabAlignment::Decimal,
            leader: TabLeader::Hyphen,
        },
    ];
    let mut edit = source.edit();
    edit.set_paragraph_tab_stops(0, &stops).unwrap();
    let after = edit.commit().unwrap().into_package();
    let first = after.body().paragraphs().next().unwrap();
    assert_eq!(first.format().tab_stops(), stops);
    assert_one_splice(
        &bytes(&source),
        &bytes(&after),
        r"\tx1440\tqr\tldot\tx5760 ",
        r"\tqc\tx720\tqdec\tlhyph\tx2880 ",
    );
    let mut edit = after.edit();
    edit.set_paragraph_tab_stops(0, &[]).unwrap();
    let cleared = edit.commit().unwrap().into_package();
    assert!(
        cleared
            .body()
            .paragraphs()
            .next()
            .unwrap()
            .format()
            .tab_stops()
            .is_empty()
    );
    assert!(bytes(&cleared).contains(r"\pard\plain{\*\bkmkstart Intro}"));
    // A paragraph without stops gains them after its paragraph reset.
    let mut edit = source.edit();
    edit.set_paragraph_tab_stops(1, &stops[..1]).unwrap();
    let added = edit.commit().unwrap().into_package();
    assert_eq!(
        added
            .body()
            .paragraphs()
            .nth(1)
            .unwrap()
            .format()
            .tab_stops(),
        &stops[..1]
    );
}

#[test]
fn section_values_replace_and_insert_controls() {
    let source = source();
    let mut edit = source.edit();
    edit.set_section_values(
        1,
        &[
            SectionValue::BreakType(SectionBreakType::Page),
            SectionValue::Orientation(PageOrientation::Portrait),
            SectionValue::PageWidth(12240),
            SectionValue::PageHeight(15840),
            SectionValue::Columns(1),
            SectionValue::MarginLeft(1800),
        ],
    )
    .unwrap();
    let after = edit.commit().unwrap().into_package();
    let section = &after.sections()[1].properties;
    assert_eq!(section.break_type, SectionBreakType::Page);
    assert_eq!(section.orientation, PageOrientation::Portrait);
    assert_eq!(section.page_width, 12240);
    assert_eq!(section.columns.count, 1);
    assert_eq!(section.margin_left, 1800);
    assert_eq!(
        format!("{:?}", after.sections()[0]),
        format!("{:?}", source.sections()[0])
    );
    assert!(bytes(&after).contains(
        r"\sect\sectd\marglsxn1800\sbkpage\pgwsxn12240\pghsxn15840\cols1\colsx720\linebetcol"
    ));
    let mut edit = source.edit();
    assert!(
        edit.set_section_values(2, &[SectionValue::Columns(2)])
            .is_err()
    );
    assert!(
        edit.set_section_values(0, &[SectionValue::Columns(0)])
            .is_err()
    );
}

#[test]
fn marks_and_body_work_do_not_mix() {
    let source = source();
    let mut edit = source.edit();
    edit.remove_paragraph(Position::new(1)).unwrap();
    assert!(matches!(
        edit.set_field_instruction(0, "PAGE"),
        Err(Error::Refused(refusal)) if refusal.kind() == litchi_core::edit::RefusalKind::UnsupportedOperation && refusal.reason() == "body splices and retained-destination edits cannot compose safely"
    ));
    let mut edit = source.edit();
    edit.rename_bookmark(0, "A").unwrap();
    assert!(matches!(
        edit.rename_bookmark(0, "B"),
        Err(Error::Conflict { .. })
    ));
}

/// The section-break follow-up of record 0308: a paragraph moves across a
/// section break, which keeps its place, and moves back byte-exactly.
#[test]
fn a_paragraph_moves_across_a_section_break_and_back() {
    let rtf = "{\\rtf1\\ansi\\deff0{\\fonttbl{\\f0 Times;}}\\sectd\r\n\
\\pard\\plain One\\par\r\n\
\\pard\\plain Two\\par\r\n\
\\sect\\sectd\\sbknone\\cols2\r\n\
\\pard\\plain Three\\par\r\n}";
    let source = Package::parse(rtf).unwrap();
    assert_eq!(source.sections().len(), 2);
    let mut edit = source.edit();
    edit.move_paragraph(Position::new(0), Position::new(2))
        .unwrap();
    let moved = edit.commit().unwrap().into_package();
    let texts: Vec<_> = moved.body().paragraphs().map(|p| p.to_text()).collect();
    assert_eq!(texts, ["Two", "Three", "One"]);
    assert_eq!(moved.sections().len(), 2);
    let mut edit = moved.edit();
    edit.move_paragraph(Position::new(2), Position::new(0))
        .unwrap();
    let back = edit.commit().unwrap().into_package();
    assert_eq!(back.to_bytes().unwrap(), rtf.as_bytes());
    let mut edit = source.edit();
    edit.remove_paragraph(Position::new(0)).unwrap();
    let removed = edit.commit().unwrap().into_package();
    assert_eq!(removed.sections().len(), 2);
    // Removing a section's only paragraph would change how the section
    // reads, which is refused.
    let mut edit = source.edit();
    edit.remove_paragraph(Position::new(2)).unwrap();
    assert!(
        matches!(edit.commit(), Err(Error::Refused(refusal)) if refusal.kind() == litchi_core::edit::RefusalKind::UnsupportedContent)
    );
}
