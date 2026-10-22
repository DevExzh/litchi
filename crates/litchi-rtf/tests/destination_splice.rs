#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design"
)]

//! Exact-source publication of retained destination text edits.

use litchi_rtf::{
    EditOptions, HeaderFooterType, Package, SemanticLimits, TableCellPath,
    edit::{Error, HeaderFooterParagraph},
};

const MIXED: &str = r"{\rtf1\ansi{\fonttbl{\f0 Arial;}}
{\header \pard Head text\par}
\pard Before\par
\trowd\cellx1000\cellx2000\cellx3000\pard\intbl A\cell {\b B}\cell\intbl\cell\row
\pard After{\super\chftn{\footnote\pard\plain\chftn note body}}\par
}";

fn edited(source: &Package, stage: impl FnOnce(&mut litchi_rtf::edit::Edit)) -> Vec<u8> {
    let mut edit = source.edit();
    stage(&mut edit);
    edit.commit().unwrap().package().to_bytes().unwrap()
}

fn replaced(from: &str, to: &str) -> Vec<u8> {
    assert_eq!(MIXED.matches(from).count(), 1, "ambiguous fixture fragment");
    MIXED.replacen(from, to, 1).into_bytes()
}

#[test]
fn cell_text_is_spliced_in_place_without_moving_the_table() {
    let source = Package::parse(MIXED).unwrap();
    let bytes = edited(&source, |edit| {
        edit.set_table_cell_text(TableCellPath::outer(0, 0, 0), "Changed")
            .unwrap();
    });
    assert_eq!(bytes, replaced(r"\intbl A\cell", r"\intbl Changed\cell"));
    let reopened = Package::from_bytes(&bytes).unwrap();
    assert_eq!(reopened.tables()[0].rows()[0].cells()[0].text(), "Changed");
    assert_eq!(reopened.text(), source.text());
}

#[test]
fn formatted_and_empty_cells_keep_their_context() {
    let source = Package::parse(MIXED).unwrap();
    let formatted = edited(&source, |edit| {
        edit.set_table_cell_text(TableCellPath::outer(0, 0, 1), "5 bold")
            .unwrap();
    });
    assert_eq!(formatted, replaced(r"{\b B}", r"{\b 5 bold}"));

    // The empty cell has no text range; the replacement is inserted before
    // its `\cell` mark with a control-word delimiter.
    let empty = edited(&source, |edit| {
        edit.set_table_cell_text(TableCellPath::outer(0, 0, 2), "9 nine")
            .unwrap();
    });
    assert_eq!(
        empty,
        replaced(r"\intbl\cell\row", r"\intbl 9 nine\cell\row")
    );
    assert_eq!(
        Package::from_bytes(&empty).unwrap().tables()[0].rows()[0].cells()[2].text(),
        "9 nine"
    );
}

#[test]
fn escapes_and_unicode_round_trip() {
    let source = Package::parse(MIXED).unwrap();
    let bytes = edited(&source, |edit| {
        edit.set_table_cell_text(TableCellPath::outer(0, 0, 0), "{x}\\ é𝄞\ttab")
            .unwrap();
    });
    let reopened = Package::from_bytes(&bytes).unwrap();
    assert_eq!(
        reopened.tables()[0].rows()[0].cells()[0].text(),
        "{x}\\ é𝄞\ttab"
    );
    assert_eq!(reopened.tables()[0].rows()[0].cells()[1].text(), "B");
}

#[test]
fn header_and_note_text_splice_in_documents_with_tables() {
    let source = Package::parse(MIXED).unwrap();
    let header = edited(&source, |edit| {
        edit.set_header_footer_text(
            HeaderFooterParagraph::new(0, HeaderFooterType::Header, 0),
            "Running head",
        )
        .unwrap();
    });
    assert_eq!(header, replaced("Head text", "Running head"));

    let note = edited(&source, |edit| {
        edit.set_note_text(0, "New note").unwrap();
    });
    assert_eq!(note, replaced("note body", "New note"));
}

#[test]
fn many_destinations_publish_in_one_splice() {
    let mut rtf = String::from(r"{\rtf1\ansi\pard Lead\par");
    for row in 0..250 {
        rtf.push_str(r"\trowd\cellx1000\cellx2000\cellx3000\cellx4000\pard\intbl ");
        for cell in 0..4 {
            rtf.push_str(&format!("r{row}c{cell}\\cell "));
        }
        rtf.push_str(r"\row");
    }
    rtf.push_str(r"\pard Tail\par}");
    let source = Package::parse(&rtf).unwrap();
    assert_eq!(source.tables()[0].row_count(), 250);
    let limits = SemanticLimits::builder()
        .max_operations(1_000)
        .unwrap()
        .build()
        .unwrap();
    let mut edit = source.edit_with_options(EditOptions::new().semantic_limits(limits));
    for row in 0..250 {
        for cell in 0..4 {
            edit.set_table_cell_text(TableCellPath::outer(0, row, cell), format!("v{row}.{cell}"))
                .unwrap();
        }
    }
    let commit = edit.commit().unwrap();
    let bytes = commit.package().to_bytes().unwrap();
    let mut expected = rtf.clone();
    for row in 0..250 {
        for cell in 0..4 {
            expected = expected.replacen(
                &format!(" r{row}c{cell}\\cell"),
                &format!(" v{row}.{cell}\\cell"),
                1,
            );
        }
    }
    assert_eq!(String::from_utf8(bytes).unwrap(), expected);
    assert_eq!(commit.package().text(), source.text());
}

#[test]
fn inverse_patch_restores_exact_source_bytes() {
    let source = Package::parse(MIXED).unwrap();
    let mut edit = source.edit();
    edit.set_table_cell_text(TableCellPath::outer(0, 0, 0), "Changed")
        .unwrap();
    edit.set_note_text(0, "New note").unwrap();
    let commit = edit.commit().unwrap();
    let restored = commit
        .package()
        .apply(&commit.patch().inverse())
        .unwrap()
        .into_package();
    assert_eq!(restored.to_bytes().unwrap(), MIXED.as_bytes());
}

#[test]
fn state_changing_controls_inside_the_story_are_refused() {
    // A paragraph reset inside the cell would be dropped by a splice and
    // change the formatting of the content after it.
    let source = Package::parse(
        r"{\rtf1\ansi\pard Body\par\trowd\cellx1000\cellx2000\pard\intbl A\par\pard\intbl\qc B\cell C\cell\row\pard Tail\par}",
    )
    .unwrap();
    let exact = source.to_bytes().unwrap();
    let mut edit = source.edit();
    edit.set_table_cell_text(TableCellPath::outer(0, 0, 0), "X")
        .unwrap();
    assert!(
        matches!(edit.commit(), Err(Error::Refused(refusal)) if refusal.kind() == litchi_core::edit::RefusalKind::UnsupportedContent)
    );
    assert_eq!(source.to_bytes().unwrap(), exact);

    // The neighbouring plain cell still splices exactly.
    let mut edit = source.edit();
    edit.set_table_cell_text(TableCellPath::outer(0, 0, 1), "Y")
        .unwrap();
    let bytes = edit.commit().unwrap().package().to_bytes().unwrap();
    assert!(
        String::from_utf8(bytes)
            .unwrap()
            .contains(r"\cell Y\cell\row")
    );
}

#[test]
fn header_paragraph_breaks_are_refused_before_staging() {
    let source = Package::parse(MIXED).unwrap();
    let mut edit = source.edit();
    assert!(matches!(
        edit.set_header_footer_text(
            HeaderFooterParagraph::new(0, HeaderFooterType::Header, 0),
            "two\nlines",
        ),
        Err(Error::Refused(refusal)) if refusal.kind() == litchi_core::edit::RefusalKind::UnsupportedContent
    ));
}

#[test]
fn tab_inside_a_cell_belongs_to_the_cell_story() {
    let source = Package::parse(
        r"{\rtf1\ansi\pard Before\par\trowd\cellx1000\pard\intbl x\tab y\cell\row\pard After\par}",
    )
    .unwrap();
    assert_eq!(source.text(), "Before\nAfter\n");
    assert_eq!(source.tables()[0].rows()[0].cells()[0].text(), "x\ty");
}
