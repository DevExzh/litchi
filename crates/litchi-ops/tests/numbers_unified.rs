//! Unified Numbers coverage (record 0297): creation, renames, sheet and
//! table moves, formulas, display formats, comments, merged ranges, cell
//! removal, batch coalescing into one leaf commit per table, and the typed
//! gaps the leaf still has.

#![cfg(feature = "numbers")]

use litchi_core::FileFormat;
use litchi_ops::{BatchItem, BatchMode, Code, File, NewNode, Placement, Props, create, props};

fn fixture(path: &str) -> std::path::PathBuf {
    std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/numbers")
        .join(path)
}

fn read(path: &str) -> File {
    litchi_ops::open_bytes(std::fs::read(fixture(path)).expect("read fixture")).expect("open")
}

fn reopen(file: &File) -> File {
    litchi_ops::open_bytes(file.to_bytes().expect("publish")).expect("reopen")
}

fn prop(file: &File, path: &str, key: &str) -> Option<String> {
    let node = file.get(path).expect("path");
    let format = node.view().format().expect("format");
    format
        .into_iter()
        .find(|prop| prop.key() == key)
        .map(|prop| prop.value().to_string())
}

fn leaf(file: &File) -> litchi_numbers::Package {
    litchi_numbers::Package::from_bytes(&file.to_bytes().expect("publish")).expect("leaf reopen")
}

#[test]
fn create_returns_a_deterministic_minimal_workbook() {
    let file = create(FileFormat::Numbers).expect("create");
    assert!(matches!(file, File::Numbers(_)));
    assert_eq!(prop(&file, "/", "sheetCount").as_deref(), Some("1"));
    assert_eq!(prop(&file, "/sheet[1]", "name").as_deref(), Some("Sheet 1"));
    assert_eq!(
        prop(&file, "/sheet[1]/table[1]", "rowCount").as_deref(),
        Some("10")
    );
    let again = create(FileFormat::Numbers).expect("create again");
    assert_eq!(file.to_bytes().unwrap(), again.to_bytes().unwrap());
    let edited = file
        .set("/sheet[1]/table[1]/A1", props! { "value" => "fresh" })
        .expect("write into the fresh workbook");
    let reopened = reopen(&edited);
    assert_eq!(
        prop(&reopened, "/sheet[1]/table[1]/A1", "value").as_deref(),
        Some("fresh")
    );
    let numeric = reopened
        .set("/sheet[1]/table[1]/A2", props! { "value" => 2 })
        .expect("a number into the fresh workbook");
    let reopened = reopen(&numeric);
    assert_eq!(
        prop(&reopened, "/sheet[1]/table[1]/A2", "value").as_deref(),
        Some("2")
    );
    let merged = reopened
        .add(
            "/sheet[1]/table[1]",
            NewNode::new("mergeCell").prop("ref", "A3:B4"),
        )
        .expect("a merge into the fresh workbook")
        .into_file();
    assert_eq!(
        prop(&reopen(&merged), "/sheet[1]/table[1]", "mergeCount").as_deref(),
        Some("1")
    );
    // Gap: the built blank has no formula-owner dependency record for its
    // table, so the leaf refuses formula authoring (FormulaCache).
    let error = reopened
        .set("/sheet[1]/table[1]/B2", props! { "formula" => "A2*3" })
        .expect_err("formulas need a formula owner");
    assert_eq!(error.code(), Code::UnsupportedCapability);
}

#[test]
fn renames_sheets_and_tables_in_one_names_transaction() {
    let file = read("table-relocation-native.numbers");
    let renamed = file
        .edit(|tx| {
            tx.set("/sheet[1]", props! { "name" => "Inputs" })?;
            tx.set("/sheet[2]/table[1]", props! { "name" => "Results" })?;
            Ok(())
        })
        .expect("renames")
        .file;
    let reopened = reopen(&renamed);
    assert_eq!(
        prop(&reopened, "/sheet[1]", "name").as_deref(),
        Some("Inputs")
    );
    assert_eq!(
        prop(&reopened, "/sheet[2]/table[1]", "name").as_deref(),
        Some("Results")
    );
    assert_eq!(
        prop(&reopened, "/Inputs/Table 1", "name").as_deref(),
        Some("Table 1")
    );
    let error = file
        .set("/sheet[1]", props! { "name" => "" })
        .expect_err("empty name");
    assert_eq!(error.code(), Code::InvalidValue);
    let error = file
        .set("/sheet[1]", props! { "tableCount" => 3 })
        .expect_err("read-only");
    assert_eq!(error.code(), Code::UnsupportedCapability);
}

#[test]
fn moves_sheets_and_relocates_tables() {
    let file = read("table-relocation-native.numbers");
    let moved = file
        .move_to("/sheet[2]", Placement::Start)
        .expect("sheet move");
    let reopened = reopen(&moved);
    assert_eq!(
        prop(&reopened, "/sheet[1]", "name").as_deref(),
        Some("Sheet 2")
    );
    assert_eq!(
        prop(&reopened, "/sheet[2]", "name").as_deref(),
        Some("Sheet 1")
    );

    let relocated = file
        .move_to(
            "/sheet[1]/table[1]",
            Placement::inside("/sheet[2]").expect("placement"),
        )
        .expect("table relocation");
    let reopened = reopen(&relocated);
    assert_eq!(
        prop(&reopened, "/sheet[1]", "tableCount").as_deref(),
        Some("0")
    );
    assert_eq!(
        prop(&reopened, "/sheet[2]/table[2]", "name").as_deref(),
        Some("Table 1")
    );
    let error = file
        .move_to("/sheet[1]/table[1]", Placement::Start)
        .expect_err("in-sheet order");
    assert_eq!(error.code(), Code::UnsupportedCapability);
}

#[test]
fn writes_formulas_through_the_formula_parser() {
    let file = read("basic.numbers");
    let edited = file
        .set("/sheet[1]/table[1]/C3", props! { "formula" => "=B3*2+1" })
        .expect("formula");
    let reopened = reopen(&edited);
    assert_eq!(
        prop(&reopened, "/sheet[1]/table[1]/C3", "type").as_deref(),
        Some("formula")
    );
    let formula = prop(&reopened, "/sheet[1]/table[1]/C3", "formula").expect("formula text");
    assert!(formula.contains("B3"), "{formula}");
    let error = file
        .set("/sheet[1]/table[1]/C3", props! { "formula" => "=1+" })
        .expect_err("malformed");
    assert_eq!(error.code(), Code::InvalidValue);
    let error = file
        .set(
            "/sheet[1]/table[1]/C3",
            props! { "formula" => "=VLOOKUP(A1)" },
        )
        .expect_err("unsupported function");
    assert!(
        matches!(
            error.code(),
            Code::UnsupportedCapability | Code::LimitExceeded
        ),
        "{error}"
    );
}

#[test]
fn sets_display_formats_and_refuses_comments() {
    let file = read("basic.numbers");
    let edited = file
        .set(
            "/sheet[1]/table[1]/B3",
            props! { "numberFormat" => "#,##0.00" },
        )
        .expect("decimal format");
    let package = leaf(&edited);
    let format = package
        .table_cell_number_format(0, 0, litchi_numbers::CellPosition::new(2, 1))
        .expect("read format")
        .expect("explicit format");
    assert_eq!(
        format.thousands_separator(),
        litchi_numbers::cell::data_format::ThousandsSeparator::Shown
    );
    let error = file
        .set("/sheet[1]/table[1]/B3", props! { "numberFormat" => "0.0x" })
        .expect_err("format code");
    assert_eq!(error.code(), Code::InvalidValue);

    // Comments are a declared gap: the leaf's comment transaction refuses
    // natively authored packages, so the unified key is not claimed.
    let error = file
        .set("/sheet[1]/table[1]/B2", props! { "comment" => "check me" })
        .expect_err("comment gap");
    assert_eq!(error.code(), Code::UnsupportedCapability);
}

#[test]
fn adds_and_removes_merged_ranges() {
    let file = read("basic.numbers");
    let added = file
        .add(
            "/sheet[1]/table[1]",
            NewNode::new("mergeCell").prop("ref", "D5:E6"),
        )
        .expect("merge");
    assert_eq!(added.path.to_string(), "/sheet[1]/table[1]/mergeCell[1]");
    let merged = reopen(&added.into_file());
    assert_eq!(
        prop(&merged, "/sheet[1]/table[1]/mergeCell[1]", "ref").as_deref(),
        Some("D5:E6")
    );
    assert_eq!(
        prop(&merged, "/sheet[1]/table[1]", "mergeCount").as_deref(),
        Some("1")
    );
    let removed = merged
        .remove("/sheet[1]/table[1]/mergeCell[1]")
        .expect("unmerge");
    let reopened = reopen(&removed);
    assert_eq!(
        prop(&reopened, "/sheet[1]/table[1]", "mergeCount").as_deref(),
        Some("0")
    );
    assert!(!reopened.to_bytes().unwrap().is_empty());

    let native = read("table-merges-native.numbers");
    assert_eq!(
        prop(&native, "/sheet[1]/table[1]", "mergeCount").as_deref(),
        Some("1")
    );
}

#[test]
fn removing_a_cell_clears_it() {
    let file = read("basic.numbers");
    let cleared = file.remove("/sheet[1]/table[1]/B3").expect("clear");
    let reopened = reopen(&cleared);
    assert!(
        reopened
            .find("/sheet[1]/table[1]/B3")
            .expect("find")
            .is_none_or(
                |node| prop(&reopened, &node.path().to_string(), "type").as_deref()
                    == Some("empty")
            )
    );
}

fn batch_set(path: String, properties: Props<'static>) -> BatchItem<'static> {
    let mut item = BatchItem::new("set");
    item.path = Some(path.into());
    item.props = Some(properties);
    item
}

#[test]
fn batch_cell_sets_share_one_leaf_commit_per_table() {
    let file = read("basic.numbers");
    let items: Vec<_> = (0..40)
        .map(|index| {
            let row = index % 20 + 3;
            let column = ["C", "D"][index / 20];
            batch_set(
                format!("/sheet[1]/table[1]/{column}{row}"),
                Props::new().with("value", index as i64),
            )
        })
        .collect();
    let commit = file.batch(items, BatchMode::Atomic).expect("batch");
    assert_eq!(commit.diagnostics.operations(), 40);
    assert_eq!(commit.diagnostics.commits(), 1);
    let reopened = reopen(&commit.file);
    assert_eq!(
        prop(&reopened, "/sheet[1]/table[1]/D22", "value").as_deref(),
        Some("39")
    );

    // A repeated cell closes the run: two commits, last write wins.
    let commit = file
        .edit(|tx| {
            tx.set("/sheet[1]/table[1]/C4", props! { "value" => 1 })?;
            tx.set("/sheet[1]/table[1]/C4", props! { "value" => 2 })?;
            Ok(())
        })
        .expect("repeated");
    assert_eq!(commit.diagnostics.commits(), 2);
    assert_eq!(
        prop(&commit.file, "/sheet[1]/table[1]/C4", "value").as_deref(),
        Some("2")
    );
    let restored = commit.patch.inverse().apply(&commit.file).expect("inverse");
    assert_eq!(restored.to_bytes().unwrap(), file.to_bytes().unwrap());
}

#[test]
fn lifecycle_operations_are_typed_gaps() {
    let file = read("table-relocation-native.numbers");
    for (parent, kind) in [
        ("/", "sheet"),
        ("/sheet[1]", "table"),
        ("/sheet[1]/table[1]", "row"),
    ] {
        let error = file
            .add(parent, NewNode::new(kind))
            .expect_err("lifecycle add");
        assert_eq!(error.code(), Code::UnsupportedCapability, "{kind}");
    }
    for path in ["/sheet[1]", "/sheet[1]/table[1]"] {
        let error = file.remove(path).expect_err("lifecycle remove");
        assert_eq!(error.code(), Code::UnsupportedCapability, "{path}");
    }
}

/// The tree of a file, as `(path, value)` pairs of every cell and the
/// sheet and table names: what a structural round trip must restore.
fn snapshot(file: &File) -> Vec<(String, Option<String>, Option<String>)> {
    let mut out = Vec::new();
    for node in file.query("sheet, table, cell").expect("query") {
        let node = node.expect("node");
        let path = node.path().to_string();
        out.push((
            path.clone(),
            prop(file, &path, "name"),
            prop(file, &path, "value"),
        ));
    }
    out
}

/// Every Numbers corpus file: a merge added and removed, a rename and its
/// revert, and (on multi-sheet workbooks) a sheet move and its revert either
/// restore the workbook's reads exactly or refuse with a typed code.
#[test]
fn structural_round_trips_over_the_corpus() {
    let mut files = 0;
    let mut merged = 0;
    let mut moved = 0;
    let mut refusals = Vec::new();
    let mut entries: Vec<_> = std::fs::read_dir(fixture(""))
        .expect("corpus")
        .map(|entry| entry.expect("entry").path())
        .filter(|path| path.extension().is_some_and(|ext| ext == "numbers") && path.is_file())
        .collect();
    entries.sort();
    for path in entries {
        let name = path.file_name().unwrap().to_string_lossy().into_owned();
        let file = litchi_ops::open_bytes(std::fs::read(&path).unwrap()).expect("open");
        let before = snapshot(&file);
        files += 1;

        let renamed = file
            .set("/sheet[1]", props! { "name" => "Round trip" })
            .and_then(|file| {
                let original = prop(&file.clone(), "/sheet[1]", "name");
                assert_eq!(original.as_deref(), Some("Round trip"));
                let back = before[0].1.clone().unwrap_or_default();
                file.set("/sheet[1]", props! { "name" => back.as_str() })
            });
        match renamed {
            Ok(file) => assert_eq!(snapshot(&reopen(&file)), before, "{name}: rename"),
            Err(error) => refusals.push(format!("{name} rename: {}", error.code())),
        }

        if prop(&file, "/sheet[1]", "tableCount").as_deref() != Some("0") {
            let rows: u32 = prop(&file, "/sheet[1]/table[1]", "rowCount")
                .and_then(|rows| rows.parse().ok())
                .unwrap_or(0);
            if rows >= 2 {
                let range = format!("A{}:A{rows}", rows - 1);
                let added = file.add(
                    "/sheet[1]/table[1]",
                    NewNode::new("mergeCell").prop("ref", range.as_str()),
                );
                match added.and_then(|added| {
                    let path = added.path.to_string();
                    added.into_file().remove(path.as_str())
                }) {
                    Ok(file) => {
                        assert_eq!(snapshot(&reopen(&file)), before, "{name}: merge");
                        merged += 1;
                    },
                    Err(error) => refusals.push(format!("{name} merge: {}", error.code())),
                }
            }
        }

        if prop(&file, "/", "sheetCount").and_then(|count| count.parse::<usize>().ok()) >= Some(2) {
            match file
                .move_to("/sheet[2]", Placement::Start)
                .and_then(|file| file.move_to("/sheet[2]", Placement::Start))
            {
                Ok(file) => {
                    assert_eq!(snapshot(&reopen(&file)), before, "{name}: sheet move");
                    moved += 1;
                },
                Err(error) => refusals.push(format!("{name} move: {}", error.code())),
            }
        }
    }
    for refusal in &refusals {
        assert!(
            refusal.ends_with("unsupported_capability")
                || refusal.ends_with("document_protected")
                || refusal.ends_with("invalid_input"),
            "untyped refusal {refusal}"
        );
    }
    eprintln!(
        "numbers structural round trips: {files} files, {merged} merge round trips, {moved} sheet-move round trips, refusals: {refusals:?}"
    );
    assert!(
        files >= 30 && merged * 2 >= files,
        "{merged} of {files} merged"
    );
}
