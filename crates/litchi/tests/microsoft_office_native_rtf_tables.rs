//! Native nested-table ownership and the bounded RTF catalog/query surface.
#![cfg(feature = "rtf")]

use litchi::{File, OpenOptions, SourceMode, capabilities::Subject, compat::word::Table, props};

fn catalog(file: &File, fruit: &str) {
    let tables = file.document().unwrap().tables().unwrap();
    assert_eq!(tables.len(), 1);
    assert_eq!(tables[0].row_count().unwrap(), 5);
    for (index, count) in [4, 4, 4, 4, 1].into_iter().enumerate() {
        assert_eq!(
            tables[0]
                .row_at(index)
                .unwrap()
                .unwrap()
                .cell_count()
                .unwrap(),
            count
        );
    }
    assert_eq!(
        tables[0]
            .row_at(1usize)
            .unwrap()
            .unwrap()
            .cell_at(0usize)
            .unwrap()
            .unwrap()
            .text()
            .unwrap(),
        fruit
    );
    let Table::Rtf(retained) = &tables[0] else {
        panic!("expected RTF table")
    };
    let owner = &retained.rows()[3].cells()[3];
    assert_eq!(owner.text(), "Nested table follows\n");
    assert_eq!(owner.nested_tables().len(), 1);
    assert_eq!(owner.nested_tables()[0].text_offset, 21);
    let nested = &owner.nested_tables()[0].table;
    assert_eq!(nested.rows()[0].cells()[0].text(), "N1\n10");
    assert_eq!(nested.rows()[0].cells()[1].text(), "N2");
    assert_eq!(nested.rows()[1].cells()[0].text(), "");
    assert_eq!(nested.rows()[1].cells()[1].text(), "20");
    assert_eq!(
        nested.rows()[0].cells()[0].merge().vertical,
        Some(litchi::rtf::TableCellMergeRole::First)
    );
    assert_eq!(
        nested.rows()[1].cells()[0].merge().vertical,
        Some(litchi::rtf::TableCellMergeRole::Continuation)
    );

    // The public query schema exposes outer cells. The detached format model
    // above retains nested cells without pretending they have public query paths.
    let cells = file
        .query("tc")
        .unwrap()
        .collect::<litchi::Result<Vec<_>>>()
        .unwrap();
    assert_eq!(cells.len(), 17);
    for cell in cells {
        let direct = file.get(cell.path().to_path()).unwrap();
        assert_eq!(direct.path(), cell.path());
        assert_eq!(direct.view().text().unwrap(), cell.view().text().unwrap());
    }
    assert_eq!(
        file.get("/body/tbl[1]/tr[4]/tc[4]")
            .unwrap()
            .view()
            .text()
            .unwrap()
            .as_deref(),
        Some("Nested table follows\n")
    );
    assert_eq!(
        file.get("/body/tbl[1]/tr[5]/tc[1]")
            .unwrap()
            .view()
            .text()
            .unwrap()
            .as_deref(),
        Some("Merged summary\nTABLE_END")
    );
    for subject in [
        Subject::Operation {
            verb: "get",
            element: "body",
            context: "table-order",
        },
        Subject::Operation {
            verb: "set",
            element: "cell",
            context: "nested-table-cells",
        },
    ] {
        assert!(
            file.capabilities()
                .gaps()
                .any(|gap| gap.subject() == subject)
        );
    }
}

#[test]
fn native_rtf_table_views_retain_nested_metadata_while_outer_cell_paths_edit_exactly() {
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/microsoft-office/word/word-tables-layout.rtf");
    let original = std::fs::read(&path).unwrap();
    for mode in [SourceMode::Auto, SourceMode::Eager] {
        let options = OpenOptions::new().source_mode(mode);
        let source = litchi::open_with(&path, &options).unwrap();
        catalog(&source, "Apples");
        let commit = source
            .edit(|edit| edit.set("/body/tbl[1]/tr[2]/tc[1]", props! {"text" => "Peaches"}))
            .unwrap();
        let output = commit.file().to_bytes().unwrap();
        assert_eq!(
            output,
            std::str::from_utf8(&original)
                .unwrap()
                .replace("Apples", "Peaches")
                .as_bytes()
        );
        let directory = tempfile::tempdir().unwrap();
        let saved = directory.path().join("table.rtf");
        commit.file().save_as(&saved).unwrap();
        let reopened = litchi::open_with(saved, &options).unwrap();
        catalog(&reopened, "Peaches");
        assert_eq!(reopened.text().unwrap(), source.text().unwrap());
        assert_eq!(
            commit
                .patch()
                .inverse()
                .apply(&reopened)
                .unwrap()
                .to_bytes()
                .unwrap(),
            original
        );
        assert_eq!(source.to_bytes().unwrap(), original);
    }
    assert_eq!(std::fs::read(path).unwrap(), original);
}
