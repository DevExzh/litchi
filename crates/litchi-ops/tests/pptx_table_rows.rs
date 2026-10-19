//! Unified PPTX table rows (record 0333): `tr` add with seed texts and
//! placement, `set height`, `remove`, merge consistency, on both owners and
//! after a byte round trip.
#![cfg(feature = "pptx")]

use litchi_ops::{Code, File, NewNode, OpenOptions, Placement, SourceMode, props};

const TABLE: &str = "/slide[1]/table[1]";

fn fixture() -> Vec<u8> {
    let mut builder = litchi_pptx::builder::Builder::new().unwrap();
    builder.presentation().unwrap().add_slide().unwrap();
    let package = builder.build().unwrap();
    let mut edit = package.edit().unwrap();
    edit.add_table(
        0usize,
        &[
            vec!["a1".to_owned(), "b1".to_owned()],
            vec!["a2".to_owned(), "b2".to_owned()],
        ],
        (0, 1_000_000, 2_000_000, 600_000),
    )
    .unwrap();
    edit.commit().unwrap().into_package().to_bytes().unwrap()
}

fn open(bytes: &[u8], eager: bool) -> File {
    litchi_ops::open_bytes_with(
        bytes.to_vec(),
        &OpenOptions::new().source_mode(if eager {
            SourceMode::Eager
        } else {
            SourceMode::Auto
        }),
    )
    .unwrap()
}

fn reopen(file: &File, eager: bool) -> File {
    open(&file.to_bytes().unwrap(), eager)
}

fn text(file: &File, path: &str) -> String {
    file.get(path)
        .unwrap()
        .view()
        .text()
        .unwrap()
        .unwrap_or_default()
        .into_owned()
}

fn prop(file: &File, path: &str, key: &str) -> Option<String> {
    let node = file.get(path).unwrap();
    node.view()
        .format()
        .unwrap()
        .find(|property| property.key() == key)
        .map(|property| property.value().to_string())
}

fn code(result: litchi_ops::Result<impl Sized>) -> Code {
    match result {
        Ok(_) => panic!("expected a refusal"),
        Err(error) => error.code(),
    }
}

fn rows(file: &File) -> usize {
    prop(file, TABLE, "rows").unwrap().parse().unwrap()
}

#[test]
fn rows_are_added_with_seed_texts_and_placement() {
    let bytes = fixture();
    for eager in [false, true] {
        let file = open(&bytes, eager);
        let height = prop(&file, &format!("{TABLE}/tr[1]"), "height").unwrap();
        let end = file
            .add(
                TABLE,
                NewNode::new("row").prop("c1", "Bob").prop("c2", "205"),
            )
            .unwrap();
        assert_eq!(end.path.to_string(), "/slide[1]/table[1]/tr[3]");
        let start = end
            .file
            .add(
                TABLE,
                NewNode::new("tr")
                    .prop("c2", "head")
                    .prop("height", "0.5in")
                    .at(Placement::Start),
            )
            .unwrap();
        assert_eq!(start.path.to_string(), "/slide[1]/table[1]/tr[1]");
        let after = start
            .file
            .add(
                TABLE,
                NewNode::new("tr")
                    .prop("cols", 2)
                    .at(Placement::after("/slide[1]/table[1]/tr[2]").unwrap()),
            )
            .unwrap();
        assert_eq!(after.path.to_string(), "/slide[1]/table[1]/tr[3]");
        for file in [&after.file, &reopen(&after.file, eager)] {
            assert_eq!(rows(file), 5, "{eager}");
            assert_eq!(text(file, "/slide[1]/table[1]/tr[1]/tc[2]"), "head");
            assert_eq!(text(file, "/slide[1]/table[1]/tr[1]/tc[1]"), "");
            assert_eq!(text(file, "/slide[1]/table[1]/tr[2]/tc[1]"), "a1");
            assert_eq!(text(file, "/slide[1]/table[1]/tr[3]/tc[1]"), "");
            assert_eq!(text(file, "/slide[1]/table[1]/tr[5]/tc[1]"), "Bob");
            assert_eq!(text(file, "/slide[1]/table[1]/tr[5]/tc[2]"), "205");
            assert_eq!(
                prop(file, "/slide[1]/table[1]/tr[1]", "height").as_deref(),
                Some("1.27cm")
            );
            // A new row takes the first row's height by default.
            assert_eq!(
                prop(file, "/slide[1]/table[1]/tr[5]", "height"),
                Some(height.clone())
            );
        }
        // The input handle is unchanged.
        assert_eq!(rows(&file), 2);
    }
}

#[test]
fn row_seed_texts_and_cols_are_checked() {
    let file = open(&fixture(), false);
    assert_eq!(
        code(file.add(TABLE, NewNode::new("tr").prop("cols", 3))),
        Code::InvalidValue
    );
    assert_eq!(
        code(file.add(TABLE, NewNode::new("tr").prop("c3", "x"))),
        Code::OutOfRange
    );
    assert_eq!(
        code(file.add(TABLE, NewNode::new("tr").prop("c21", "x"))),
        Code::UnsupportedProperty
    );
    assert_eq!(
        code(file.add("/slide[1]", NewNode::new("tr"))),
        Code::UnsupportedType
    );
}

#[test]
fn row_height_is_set_and_the_frame_follows() {
    let bytes = fixture();
    for eager in [false, true] {
        let file = open(&bytes, eager);
        let table_height = |file: &File| prop(file, TABLE, "height");
        let before = table_height(&file);
        let set = file
            .set("/slide[1]/table[1]/tr[2]", props! { "height" => "2in" })
            .unwrap();
        let file = reopen(&set, eager);
        assert_eq!(
            prop(&file, "/slide[1]/table[1]/tr[2]", "height").as_deref(),
            Some("5.08cm")
        );
        assert_ne!(table_height(&file), before);
        // An equal height is an exact no-op.
        let again = file
            .set("/slide[1]/table[1]/tr[2]", props! { "height" => "2in" })
            .unwrap();
        assert_eq!(again.to_bytes().unwrap(), file.to_bytes().unwrap());
        assert_eq!(
            code(file.set("/slide[1]/table[1]/tr[2]", props! { "height" => "-1in" })),
            Code::InvalidValue
        );
    }
}

#[test]
fn rows_are_removed_but_one_is_kept() {
    let bytes = fixture();
    for eager in [false, true] {
        let file = open(&bytes, eager);
        let removed = file.remove("/slide[1]/table[1]/tr[1]").unwrap();
        let file = reopen(&removed, eager);
        assert_eq!(rows(&file), 1);
        assert_eq!(text(&file, "/slide[1]/table[1]/tr[1]/tc[1]"), "a2");
        assert_eq!(
            code(file.remove("/slide[1]/table[1]/tr[1]")),
            Code::UnsupportedCapability
        );
    }
}

#[test]
fn tables_are_added_from_data_or_a_grid() {
    let bytes = fixture();
    for eager in [false, true] {
        let file = open(&bytes, eager);
        let added = file
            .add(
                "/slide[1]",
                NewNode::new("table")
                    .prop("data", "Name,Score;Ann,9;Bob")
                    .prop("x", "1in")
                    .prop("width", "4in")
                    .prop("name", "Scores"),
            )
            .unwrap();
        assert_eq!(added.path.to_string(), "/slide[1]/table[2]");
        let file = reopen(&added.file, eager);
        let table = "/slide[1]/table[2]";
        assert_eq!(prop(&file, table, "rows").as_deref(), Some("3"));
        assert_eq!(prop(&file, table, "cols").as_deref(), Some("2"));
        assert_eq!(prop(&file, table, "name").as_deref(), Some("Scores"));
        assert_eq!(text(&file, "/slide[1]/table[2]/tr[2]/tc[2]"), "9");
        assert_eq!(text(&file, "/slide[1]/table[2]/tr[3]/tc[2]"), "");
        // The default grid is three by three; rows and cols widen data.
        let grid = file.add("/slide[1]", NewNode::new("table")).unwrap();
        assert_eq!(
            prop(&grid.file, "/slide[1]/table[3]", "rows").as_deref(),
            Some("3")
        );
        assert_eq!(
            prop(&grid.file, "/slide[1]/table[3]", "cols").as_deref(),
            Some("3")
        );
        assert_eq!(
            prop(&grid.file, "/slide[1]/table[3]", "name").as_deref(),
            Some("Table 3")
        );
        let wide = file
            .add(
                "/slide[1]",
                NewNode::new("table").prop("data", "a").prop("cols", 4),
            )
            .unwrap();
        assert_eq!(
            prop(&wide.file, "/slide[1]/table[3]", "cols").as_deref(),
            Some("4")
        );
        // The new table takes row and column edits.
        let more = grid
            .file
            .add("/slide[1]/table[3]", NewNode::new("tr").prop("c3", "z"))
            .unwrap();
        assert_eq!(text(&more.file, "/slide[1]/table[3]/tr[4]/tc[3]"), "z");
        assert_eq!(
            code(file.add("/slide[1]", NewNode::new("table").prop("rows", 0))),
            Code::InvalidValue
        );
        assert_eq!(
            code(file.add("/slide[1]/table[1]", NewNode::new("table"))),
            Code::UnsupportedType
        );
    }
}

/// `ppt/tables/tables-nested.sh:86` (record 0342): a table-level set of the
/// frame's name and the six style switches; equal values are a no-op, and
/// add-only keys are refused without a change.
#[test]
fn table_name_and_style_switches_are_set() {
    let bytes = fixture();
    for eager in [false, true] {
        let file = open(&bytes, eager);
        let set = file
            .set(
                TABLE,
                props! {
                    "name" => "QuarterlySales",
                    "firstRow" => true,
                    "lastRow" => true,
                    "firstCol" => true,
                    "lastCol" => false,
                    "bandedRows" => true,
                    "bandedCols" => false,
                },
            )
            .unwrap();
        for read in [reopen(&set, false), reopen(&set, true)] {
            assert_eq!(
                prop(&read, TABLE, "name").as_deref(),
                Some("QuarterlySales")
            );
            assert_eq!(prop(&read, TABLE, "firstRow").as_deref(), Some("true"));
            assert_eq!(prop(&read, TABLE, "lastRow").as_deref(), Some("true"));
            assert_eq!(prop(&read, TABLE, "firstCol").as_deref(), Some("true"));
            assert_eq!(prop(&read, TABLE, "lastCol"), None);
            assert_eq!(prop(&read, TABLE, "bandedRows").as_deref(), Some("true"));
            assert_eq!(prop(&read, TABLE, "bandedCols"), None);
        }
        let off = set.set(TABLE, props! { "bandedRows" => false }).unwrap();
        assert_eq!(prop(&off, TABLE, "bandedRows").as_deref(), Some("false"));
        let same = set
            .set(
                TABLE,
                props! { "lastRow" => true, "name" => "QuarterlySales" },
            )
            .unwrap();
        assert_eq!(same.to_bytes().unwrap(), set.to_bytes().unwrap());
        let before = set.to_bytes().unwrap();
        assert_eq!(
            code(set.set(TABLE, props! { "rows" => 4 })),
            Code::UnsupportedProperty
        );
        assert_eq!(set.to_bytes().unwrap(), before);
    }
}
