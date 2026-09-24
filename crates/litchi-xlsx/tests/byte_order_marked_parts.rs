#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "differential assertions intentionally panic on fixture failures"
)]

//! Change 0765: a workbook part that begins with a UTF-8 byte-order mark
//! reads, edits and saves exactly like the same part without one.
//!
//! quick-xml drops a leading mark before its first event without counting it
//! in its positions, so every span taken from a reader was three bytes early
//! in a marked part. In an indented part those three bytes are whitespace, so
//! a splice at the early span left the old element's last three bytes behind
//! as text and still produced well-formed XML. Each test runs one scenario on
//! a package and on its twin whose XML members carry a mark, in compact and in
//! indented form, and requires identical observations and an output that
//! differs only by marks: every member of the marked output is the unmarked
//! output's member, or that member behind one mark, and a mark appears only on
//! a member whose input had one.

use std::collections::{BTreeMap, BTreeSet};
use std::sync::Arc;

use litchi_core::OwnedSource;
use litchi_xlsx::cell_values::{CellValueEdit, SourceBackedEditor};
use litchi_xlsx::{Address, Hyperlink, HyperlinkReference, Number, Workbook};
use soapberry_zip::office::{ArchiveReader, StreamingArchiveWriter};

const BOM: &[u8] = b"\xEF\xBB\xBF";
const MARKED_RELATIONSHIPS: &str = concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../test-data/libreoffice-core/sc/qa/unit/data/xlsx/tdf167689_xmlMaps_and_xmlColumnPr.xlsx"
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
/// between elements, which SpreadsheetML ignores, never inside an element
/// that holds only text.
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

/// Rebuild `archive`: every member `mark` selects behind one mark, every
/// other member without one, and XML members indented when `indented`.
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
            assert_eq!(
                String::from_utf8_lossy(rest),
                String::from_utf8_lossy(unmarked),
                "{name}: differs beyond its mark"
            );
            kept.insert(name.clone());
        } else {
            assert_eq!(
                String::from_utf8_lossy(bytes),
                String::from_utf8_lossy(unmarked),
                "{name}: differs"
            );
        }
    }
    kept
}

/// Run `scenario` on the unmarked and the marked twin of `archive`, compact
/// and indented, and compare observations and outputs. Returns the members
/// that kept a mark in the compact run.
fn differential(
    archive: &[u8],
    mark: &dyn Fn(&str) -> bool,
    scenario: &dyn Fn(Workbook) -> (String, Vec<u8>),
) -> BTreeSet<String> {
    let mut kept = BTreeSet::new();
    for indented in [true, false] {
        let (plain, _) = rebuild(archive, &|_| false, indented);
        let (marked, marked_input) = rebuild(archive, mark, indented);
        assert!(!marked_input.is_empty(), "the scenario must mark a member");
        let (plain_observed, plain_out) = scenario(Workbook::from_bytes(plain).unwrap());
        let (marked_observed, marked_out) = scenario(Workbook::from_bytes(marked).unwrap());
        assert_eq!(marked_observed, plain_observed, "indented: {indented}");
        kept = assert_equal_but_marks(&marked_out, &plain_out, &marked_input);
    }
    kept
}

fn observe(workbook: &Workbook) -> String {
    let mut observed = String::new();
    for sheet in workbook.sheets() {
        observed.push_str(&format!("[{}]", sheet.name()));
        for (address, cell) in sheet.cells("A1:XFD1048576").unwrap() {
            observed.push_str(&format!("{address:?}={cell:?};"));
        }
        for merge in sheet.merges().unwrap() {
            observed.push_str(&format!("merge {merge:?};"));
        }
    }
    observed
}

fn generated_workbook() -> Vec<u8> {
    generated_workbook_with(true)
}

fn generated_workbook_with(second_sheet: bool) -> Vec<u8> {
    let workbook = Workbook::create().unwrap();
    let mut edit = workbook.edit().unwrap();
    if second_sheet {
        edit.add("Second")
            .unwrap()
            .set(Address::at(0, 0).unwrap(), "second")
            .unwrap();
    }
    {
        let mut sheet = edit.sheet("Sheet1").unwrap().unwrap();
        for row in 0..12_u32 {
            for column in 0..6_u32 {
                let address = Address::at(row, column).unwrap();
                if (row + column) % 3 == 0 {
                    sheet.set(address, format!("text {row}/{column} & <x>").as_str())
                } else {
                    sheet.set(address, i32::try_from(row * 100 + column).unwrap())
                }
                .unwrap();
            }
        }
        sheet.merge("A20:B21").unwrap();
    }
    let commit = edit.commit().unwrap();
    commit.workbook().to_bytes().unwrap()
}

fn committed(edit: litchi_xlsx::Edit) -> (String, Vec<u8>) {
    let commit = edit.commit().unwrap();
    let observed = observe(commit.workbook());
    (observed, commit.workbook().to_bytes().unwrap())
}

#[test]
fn a_fully_marked_workbook_reads_and_saves_like_the_unmarked_workbook() {
    let kept = differential(&generated_workbook(), &is_xml_member, &|workbook| {
        (observe(&workbook), workbook.to_bytes().unwrap())
    });
    assert!(kept.contains("xl/worksheets/sheet1.xml"), "{kept:?}");
    assert!(kept.contains("xl/workbook.xml"), "{kept:?}");
}

#[test]
fn cell_edits_of_marked_indented_worksheets_match_unmarked_edits() {
    // Base: an indented marked worksheet published `…</c>/c>` — the old
    // cell's tail left behind as text — and a later parse ignored it.
    let kept = differential(&generated_workbook(), &is_xml_member, &|workbook| {
        let mut edit = workbook.edit().unwrap();
        {
            let mut sheet = edit.sheet("Sheet1").unwrap().unwrap();
            sheet.set(Address::at(0, 1).unwrap(), 42).unwrap();
            sheet
                .set(Address::at(5, 5).unwrap(), "edited & <v>")
                .unwrap();
            sheet.clear(Address::at(3, 3).unwrap()).unwrap();
            sheet.remove(Address::at(11, 0).unwrap()).unwrap();
            sheet.set(Address::at(40, 9).unwrap(), true).unwrap();
        }
        committed(edit)
    });
    assert!(kept.contains("xl/worksheets/sheet1.xml"), "{kept:?}");
}

#[test]
fn sheet_catalog_view_and_hyperlink_edits_of_marked_parts_match_unmarked_edits() {
    differential(&generated_workbook(), &is_xml_member, &|workbook| {
        let mut edit = workbook.edit().unwrap();
        {
            let mut tab = edit.tab("Second").unwrap().unwrap();
            tab.rename("Renamed").unwrap();
            tab.activate();
        }
        {
            let mut sheet = edit.sheet("Sheet1").unwrap().unwrap();
            sheet
                .put_hyperlink(
                    Hyperlink::external(
                        HyperlinkReference::new("B2").unwrap(),
                        "https://example.invalid/".to_owned(),
                    )
                    .unwrap(),
                )
                .unwrap();
            sheet.merge("D20:E22").unwrap();
            sheet.unmerge("A20").unwrap();
        }
        committed(edit)
    });
}

#[test]
fn source_backed_cell_edits_of_marked_parts_match_unmarked_edits() {
    // The source-backed value editor admits one-worksheet workbooks.
    differential(
        &generated_workbook_with(false),
        &is_xml_member,
        &|workbook| {
            let source = workbook.to_bytes().unwrap();
            let editor =
                SourceBackedEditor::from_read_at(Arc::new(OwnedSource::new(source))).unwrap();
            let mut edit = editor.edit("Sheet1").unwrap();
            edit.apply_batch([
                CellValueEdit::set(Address::at(0, 1).unwrap(), Number::new("10.50").unwrap()),
                CellValueEdit::set(Address::at(2, 2).unwrap(), false),
                CellValueEdit::insert(Address::at(30, 3).unwrap(), Number::new("7").unwrap()),
            ])
            .unwrap();
            let commit = edit.commit().unwrap();
            let mut output = Vec::new();
            editor
                .publish_commit_to_stream(&mut output, &commit)
                .unwrap();
            let published = Workbook::from_bytes(output.clone()).unwrap();
            (observe(&published), output)
        },
    );
}

#[test]
fn a_workbook_with_producer_marked_manifests_edits_like_its_unmarked_twin() {
    let original = std::fs::read(MARKED_RELATIONSHIPS).unwrap();
    let marked_members: BTreeSet<String> = members(&original)
        .into_iter()
        .filter(|(_, bytes)| bytes.starts_with(BOM))
        .map(|(name, _)| name)
        .collect();
    assert!(marked_members.contains("[Content_Types].xml"));
    differential(
        &original,
        &|name| marked_members.contains(name),
        &|workbook| {
            let mut edit = workbook.edit().unwrap();
            {
                let mut sheet = edit.sheet(0).unwrap().unwrap();
                sheet.set(Address::at(0, 0).unwrap(), "edited").unwrap();
            }
            committed(edit)
        },
    );
}
