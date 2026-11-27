//! Native binary Word structure, source-backed edits, and inert saved metadata.
use std::{collections::BTreeMap, io::Cursor};

use litchi_core::Position;
use litchi_doc::body_text::{self, Refusal, SectionValue, TextTarget, TransactionLimits};
use litchi_doc::package::edit::{Error, LimitKind};
use litchi_doc::parts::document_properties::DocumentPropertyVersion;
use litchi_doc::parts::tap::VerticalMergeStatus;
use litchi_doc::{EditOptions, Package};

const SOURCES: [&str; 4] = [
    "word-tables-layout.doc",
    "word-review-controls.doc",
    "word-layout-extensions.doc",
    "word-layout-extensions.dot",
];
fn native(name: &str) -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/word")
            .join(name),
    )
    .unwrap()
}
fn streams(bytes: &[u8]) -> BTreeMap<Vec<String>, Vec<u8>> {
    let mut ole =
        litchi_cfb::OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default())
            .unwrap();
    ole.list_streams()
        .into_iter()
        .map(|path| {
            let names = path.iter().map(String::as_str).collect::<Vec<_>>();
            let bytes = ole.open_stream(&names).unwrap();
            (path, bytes)
        })
        .collect()
}
fn table_part(package: &Package, index: usize) -> Vec<u8> {
    let d = package.document();
    let (offset, size) = d.fib().get_table_pointer(index).unwrap();
    let parts = streams(&package.to_bytes().unwrap());
    let key = vec![format!("{}Table", u8::from(d.fib().which_table_stream()))];
    parts[&key][offset as usize..(offset + size) as usize].to_vec()
}
fn raw_section_cps(package: &Package) -> Vec<u32> {
    let part = table_part(package, 6);
    let count = (part.len() - 4) / 16;
    part[..(count + 1) * 4]
        .chunks_exact(4)
        .map(|b| u32::from_le_bytes(b.try_into().unwrap()))
        .collect()
}
fn cell_text(package: &Package) -> Vec<String> {
    package
        .table_cells()
        .unwrap()
        .iter()
        .map(|c| c.text().to_owned())
        .collect()
}
fn bookmarks(package: &Package) -> Vec<(String, u32, u32)> {
    package
        .document()
        .bookmarks()
        .unwrap()
        .into_iter()
        .map(|b| (b.name, b.start, b.end))
        .collect()
}
fn assert_other_streams(before: &Package, after: &Package) {
    let first = streams(&before.to_bytes().unwrap());
    let second = streams(&after.to_bytes().unwrap());
    assert_eq!(
        first.keys().collect::<Vec<_>>(),
        second.keys().collect::<Vec<_>>()
    );
    for (path, bytes) in first {
        if !matches!(path.as_slice(), [name] if matches!(name.as_str(), "WordDocument" | "0Table" | "1Table"))
        {
            assert_eq!(bytes, second[&path], "{}", path.join("/"));
        }
    }
    // These source records must survive, even when the physical streams grow.
    for index in [1, 15, 31] {
        assert_eq!(table_part(before, index), table_part(after, index));
    }
}
fn save_reopen_inverse(source: &Package, commit: &litchi_doc::package::edit::Commit) -> Package {
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("edited.doc");
    commit.package().save(&path).unwrap();
    let reopened = Package::open(&path).unwrap();
    assert_eq!(
        reopened.to_bytes().unwrap(),
        commit.package().to_bytes().unwrap()
    );
    assert_other_streams(source, &reopened);
    assert_eq!(
        reopened
            .apply(&commit.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
        source.to_bytes().unwrap()
    );
    reopened
}

#[test]
fn native_fib_generation_template_and_saved_statistics_remain_distinct() {
    for (name, main, expected) in [
        (SOURCES[0], 434, [64, 370, 2, 1, 3, 71, 408, 2, 1, 3]),
        (SOURCES[1], 257, [18, 239, 1, 11, 11, 18, 239, 1, 11, 11]),
        (SOURCES[2], 379, [27, 352, 1, 17, 16, 38, 495, 1, 24, 23]),
        (SOURCES[3], 379, [27, 352, 1, 17, 16, 38, 495, 1, 24, 23]),
    ] {
        let bytes = native(name);
        let p = Package::from_bytes(&bytes).unwrap();
        let d = p.document();
        assert_eq!(d.fib().version(), 0x00c1, "base compatibility generation");
        assert_eq!(d.fib().table_pointer_count(), Some(183));
        assert!(d.fib().which_table_stream());
        assert_eq!(d.fib().is_template(), name.ends_with(".dot"));
        assert_eq!(d.fib().get_main_doc_range(), (0, main));
        let parts = streams(&bytes);
        let word = &parts[&vec!["WordDocument".to_owned()]];
        // Counted FIB extension: cswNew=5 and nFibNew=Word 2007 (0x0112).
        assert_eq!(&word[154 + 183 * 8..158 + 183 * 8], &[5, 0, 0x12, 1]);
        let dop = d.document_properties().unwrap().unwrap();
        assert_eq!(dop.version(), DocumentPropertyVersion::Word2013);
        let s = dop.base().statistics();
        // Producer-saved counters; they do not describe the current parsed paragraph count.
        assert_eq!(
            [
                s.words,
                s.characters,
                i32::from(s.pages),
                s.paragraphs,
                s.lines,
                s.words_with_subdocuments,
                s.characters_with_subdocuments,
                i32::from(s.pages_with_subdocuments),
                s.paragraphs_with_subdocuments,
                s.lines_with_subdocuments
            ],
            expected,
            "{name}"
        );
        assert_eq!(
            p.edit()
                .unwrap()
                .commit()
                .unwrap()
                .package()
                .to_bytes()
                .unwrap(),
            bytes
        );
    }
}

#[test]
fn native_tables_keep_row_headers_nested_depth_merges_and_cell_extents() {
    let p = Package::from_bytes(&native(SOURCES[0])).unwrap();
    let d = p.document();
    let tables = d.tables().unwrap();
    assert_eq!(tables.len(), 1);
    let table = &tables[0];
    let extent = table.cp_range().unwrap();
    assert_eq!((extent.start(), extent.end()), (85, 233));
    let rows = table.rows().unwrap();
    assert_eq!(rows.len(), 5);
    for (index, (row, expected)) in rows
        .iter()
        .zip([
            vec!["Item", "Quantity", "Price", "Notes"],
            vec!["Apples", "3", "1.25", "Fruit"],
            vec!["Pears", "5", "2.50", "Unicode 中文"],
            vec!["Tea", "2", "4.75", "Nested table follows\n"],
            vec!["Merged summary\nTABLE_END"],
        ])
        .enumerate()
    {
        let cells = row.cells().unwrap();
        let tap = row.properties().unwrap();
        assert_eq!(
            cells.iter().map(|c| c.text().unwrap()).collect::<Vec<_>>(),
            expected
        );
        assert_eq!(
            tap.cell_count,
            cells.len(),
            "row {index}: no phantom default cell"
        );
        assert_eq!(tap.cell_properties.len(), cells.len());
        assert_eq!(tap.is_header_row, index == 0);
        // The modern sprmPTableProps sequence supplies these boundaries. Its
        // compatibility sprmTDefTable has a different margin representation.
        assert_eq!(
            tap.cell_boundaries,
            if index == 4 {
                vec![0, 9468]
            } else {
                vec![0, 2448, 4788, 7128, 9468]
            }
        );
        let range = row.cp_range().unwrap();
        let mut start = range.start();
        for cell in cells {
            let cp = cell.cp_range().unwrap();
            assert_eq!(cp.start(), start);
            start = cp.end();
        }
        assert_eq!(start + 1, range.end(), "separate TTP after the last cell");
    }
    let paras = d.paragraphs().unwrap();
    let nested = paras
        .iter()
        .filter(|p| p.properties().table_depth() == 2)
        .collect::<Vec<_>>();
    assert_eq!(
        nested.iter().map(|p| p.text().unwrap()).collect::<Vec<_>>(),
        ["N1", "10", "N2", "", "", "20", ""]
    );
    let row_marks = nested
        .iter()
        .filter(|p| p.properties().inner_table_row_end)
        .collect::<Vec<_>>();
    assert_eq!(row_marks.len(), 2);
    for (mark, expected) in row_marks
        .iter()
        .zip([VerticalMergeStatus::First, VerticalMergeStatus::Merged])
    {
        let tap = mark.properties().table_properties.as_ref().unwrap();
        assert_eq!(tap.cell_count, 2);
        assert_eq!(tap.cell_boundaries, [5, 1170, 2227]);
        assert_eq!(tap.cell_properties[0].vertical_merge_status, expected);
    }
    assert!(matches!(
        p.table_cells(),
        Err(Error::Body(body_text::Error::Refused(
            Refusal::ComplexTableCell
        )))
    ));
}

#[test]
fn native_doc_and_dot_cell_resize_excludes_row_marks_and_preserves_other_stories() {
    for name in [SOURCES[2], SOURCES[3]] {
        let bytes = native(name);
        let p = Package::from_bytes(&bytes).unwrap();
        assert_eq!(
            cell_text(&p),
            ["AUTO_TABLE_016", "Amount", "Native AutoCaption", "42"]
        );
        let mut edit = p.edit().unwrap();
        edit.replace_text(TextTarget::TableCell(Position::new(3)), "420")
            .unwrap();
        assert!(matches!(
            edit.replace_text(TextTarget::TableCell(Position::new(4)), "invalid"),
            Err(Error::Body(body_text::Error::Refused(
                Refusal::TargetNotFound
            )))
        ));
        let commit = edit.commit().unwrap();
        let reopened = save_reopen_inverse(&p, &commit);
        assert_eq!(
            cell_text(&reopened),
            ["AUTO_TABLE_016", "Amount", "Native AutoCaption", "420"]
        );
        assert_eq!(
            reopened.document().text().unwrap(),
            p.document().text().unwrap().replacen("42", "420", 1)
        );
        assert_eq!(
            reopened.document().fib().is_template(),
            name.ends_with(".dot")
        );
        assert_eq!(reopened.document().fib().get_main_doc_range().1, 380);
        assert_eq!(reopened.document().sections()[0].end_cp, 380);
        assert_eq!(p.to_bytes().unwrap(), bytes);
    }
}

#[test]
fn native_bookmark_crud_preserves_the_original_reference_and_text() {
    let bytes = native(SOURCES[1]);
    let p = Package::from_bytes(&bytes).unwrap();
    let original = vec![("LITCHI_BOOKMARK_003".to_owned(), 126, 144)];
    assert_eq!(bookmarks(&p), original);
    let mut edit = p.edit().unwrap();
    edit.add_bookmark("Native_New", Position::new(0), 0, 6)
        .unwrap();
    edit.rename_bookmark("Native_New", "Native_Checked")
        .unwrap();
    let commit = edit.commit().unwrap();
    let reopened = save_reopen_inverse(&p, &commit);
    assert_eq!(
        bookmarks(&reopened),
        [("Native_Checked".to_owned(), 0, 6), original[0].clone()]
    );
    assert_eq!(
        reopened.document().text().unwrap(),
        p.document().text().unwrap()
    );
    // Native field and annotation PLCs remain exact; renaming an unrelated
    // bookmark makes no promise to rewrite REF instructions automatically.
    for index in [2, 3, 16, 17, 18, 19, 20] {
        assert_eq!(table_part(&p, index), table_part(&reopened, index));
    }
    let mut edit = reopened.edit().unwrap();
    edit.remove_bookmark("Native_Checked").unwrap();
    let removed = edit.commit().unwrap();
    let final_package = save_reopen_inverse(&reopened, &removed);
    assert_eq!(bookmarks(&final_package), original);
    assert_eq!(
        final_package.document().text().unwrap(),
        p.document().text().unwrap()
    );
    assert_eq!(p.to_bytes().unwrap(), bytes);
}

#[test]
fn native_failed_bookmark_rename_is_atomic_after_invalid_duplicate_and_limit_errors() {
    let bytes = native(SOURCES[1]);
    let p = Package::from_bytes(&bytes).unwrap();
    let mut edit = p.edit().unwrap();
    let before = edit.bookmarks().unwrap();
    assert!(matches!(
        edit.rename_bookmark("LITCHI_BOOKMARK_003", "bad name"),
        Err(Error::Body(body_text::Error::Refused(
            Refusal::BookmarkName
        )))
    ));
    assert_eq!(edit.bookmarks().unwrap(), before);
    assert_eq!(edit.commit().unwrap().package().to_bytes().unwrap(), bytes);
    let mut expected = p.edit().unwrap();
    expected
        .add_bookmark("Existing", Position::new(0), 0, 6)
        .unwrap();
    let expected_bytes = expected.commit().unwrap().package().to_bytes().unwrap();
    for limited in [false, true] {
        let options = EditOptions::default().transaction_limits(TransactionLimits::new(
            if limited { 2 } else { 100 },
            1024,
            4096,
        ));
        let mut edit = p.edit_with_options(options).unwrap();
        edit.add_bookmark("Existing", Position::new(0), 0, 6)
            .unwrap();
        let before = edit.bookmarks().unwrap();
        let error = edit
            .rename_bookmark(
                "LITCHI_BOOKMARK_003",
                if limited { "Renamed" } else { "existing" },
            )
            .unwrap_err();
        if limited {
            assert!(matches!(
                error,
                Error::Limit {
                    kind: LimitKind::Operations,
                    observed: 3,
                    maximum: 2
                }
            ));
        } else {
            assert!(matches!(
                error,
                Error::Body(body_text::Error::Refused(Refusal::BookmarkName))
            ));
        }
        assert_eq!(edit.bookmarks().unwrap(), before);
        assert_eq!(edit.changes().unwrap().len(), 1);
        assert_eq!(
            edit.commit().unwrap().package().to_bytes().unwrap(),
            expected_bytes
        );
    }
}

#[test]
fn native_sections_stop_at_main_story_while_retaining_the_wire_sentinel() {
    for (name, end, sentinel) in [
        (SOURCES[0], 434, 556),
        (SOURCES[1], 257, 346),
        (SOURCES[2], 379, 548),
        (SOURCES[3], 379, 548),
    ] {
        let bytes = native(name);
        let p = Package::from_bytes(&bytes).unwrap();
        let d = p.document();
        assert_eq!(d.sections().last().unwrap().end_cp, end);
        assert!(d.section_at_cp(end - 1).is_some());
        for cp in end..=sentinel {
            assert!(
                d.section_at_cp(cp).is_none(),
                "{name}: secondary story CP {cp}"
            );
        }
        assert_eq!(*raw_section_cps(&p).last().unwrap(), sentinel);
        assert_eq!(p.to_bytes().unwrap(), bytes);
    }
}

#[test]
fn native_section_edit_keeps_landscape_geometry_other_properties_and_stories() {
    use litchi_doc::section::{BreakKind, PageOrientation, columns::Layout};
    let bytes = native(SOURCES[0]);
    let p = Package::from_bytes(&bytes).unwrap();
    let d = p.document();
    let original = d.sections().to_vec();
    assert_eq!(original.len(), 2);
    assert_eq!((original[0].start_cp, original[0].end_cp), (0, 283));
    assert_eq!(original[0].page.orientation, PageOrientation::Portrait);
    assert_eq!(
        (original[0].page.width_twips, original[0].page.height_twips),
        (12240, 15840)
    );
    assert_eq!((original[1].start_cp, original[1].end_cp), (283, 434));
    assert_eq!(original[1].page.orientation, PageOrientation::Landscape);
    assert_eq!(
        (original[1].page.width_twips, original[1].page.height_twips),
        (15840, 12240)
    );
    assert_eq!(original[1].columns, Layout::even(2, 720, false).unwrap());
    assert_eq!(original[1].behavior.revision_save_id, Some(15497745));
    for section in &original {
        assert_eq!(section.break_kind, BreakKind::NewPage);
        assert_eq!(section.page.margins.left_twips, 1440);
    }
    let mut edit = p.edit().unwrap();
    edit.set_section_values(
        Position::new(1),
        &[SectionValue::MarginLeft(1600), SectionValue::Columns(3)],
    )
    .unwrap();
    let commit = edit.commit().unwrap();
    let reopened = save_reopen_inverse(&p, &commit);
    let mut expected = original;
    expected[1].page.margins.left_twips = 1600;
    expected[1].columns = Layout::even(3, 720, false).unwrap();
    assert_eq!(reopened.document().sections(), expected);
    assert_eq!(raw_section_cps(&reopened), [0, 283, 556]);
    assert_eq!(reopened.document().text().unwrap(), d.text().unwrap());
    assert_eq!(p.to_bytes().unwrap(), bytes);
}
