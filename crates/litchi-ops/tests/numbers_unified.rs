//! Unified Numbers coverage (record 0297): creation, renames, sheet and
//! table moves, formulas, display formats, comments, merged ranges, cell
//! removal, batch coalescing into one leaf commit per table, and the typed
//! gaps the leaf still has.

#![cfg(feature = "numbers")]

use litchi_core::FileFormat;
use litchi_ops::error::Source;
use litchi_ops::transaction::{Error as TransactionError, LimitKind, SemanticLimits};
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

fn leaf_bytes(package: &litchi_numbers::Package) -> Vec<u8> {
    let mut bytes = Vec::new();
    package
        .write_to(&mut bytes)
        .expect("serialize leaf package");
    bytes
}

/// The full unified single-operation route: open a source package, dispatch
/// one selector-first write through `File`, and publish it.  The callgrind
/// probe below toggles this function so dispatch, facade, leaf commit and
/// serialization are all inside the measured region.
#[inline(never)]
fn cached_formula_unified_setup(source: &[u8]) -> Vec<u8> {
    litchi_ops::open_bytes(source.to_vec())
        .expect("open unified source")
        .set(
            "/sheet[1]/table[1]/C3",
            props! { "formula" => "=B3+1", "cachedValue" => 2.0 },
        )
        .expect("unified cached formula")
        .to_bytes()
        .expect("publish unified source")
}

/// The direct leaf comparison route for the same single operation and source.
/// Its output is compared with the full `File` route before the ABBA loop.
#[inline(never)]
fn cached_formula_direct_setup(source: &[u8]) -> Vec<u8> {
    let expression =
        litchi_numbers::formula::Expression::parse("=B3+1").expect("formula expression");
    let cached = litchi_numbers::formula::CachedValue::number(2.0).expect("cached number");
    let package = litchi_numbers::Package::from_bytes(source)
        .expect("direct leaf source")
        .edit_table_cells(
            litchi_numbers::SheetSelector::index(0),
            litchi_numbers::TableSelector::index(0),
        )
        .expect("direct leaf editor")
        .set_formula_cached(litchi_numbers::At::new(2, 2), expression, cached)
        .expect("direct leaf cached formula")
        .commit()
        .expect("direct leaf cached commit")
        .into_package();
    leaf_bytes(&package)
}

/// Counted unified wrapper. The setup implementation has a distinct name so
/// exact callgrind globs cannot collect warm-up or output checks.
#[inline(never)]
fn unified_cached_formula_route(source: &[u8]) -> Vec<u8> {
    cached_formula_unified_setup(source)
}

/// Counted direct-leaf wrapper. The setup implementation has a distinct name
/// so exact callgrind globs cannot collect warm-up or output checks.
#[inline(never)]
fn direct_cached_formula_route(source: &[u8]) -> Vec<u8> {
    cached_formula_direct_setup(source)
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
    // The blank's table owner carries the native table extents, so the
    // created workbook authors formulas (record 0309).
    let formula = reopened
        .set("/sheet[1]/table[1]/B2", props! { "formula" => "A2*3" })
        .expect("a formula into the fresh workbook");
    assert_eq!(
        prop(&reopen(&formula), "/sheet[1]/table[1]/B2", "type").as_deref(),
        Some("formula")
    );
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
        .into_file();
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
    // Within one sheet (record 0301): the relocated table goes back first.
    let reordered = relocated
        .edit(|tx| tx.move_to("/sheet[2]/table[2]", Placement::Start))
        .expect("in-sheet order");
    let reopened = reopen(reordered.file());
    assert_eq!(
        prop(&reopened, "/sheet[2]/table[1]", "name").as_deref(),
        Some("Table 1")
    );
    let restored = reordered
        .patch()
        .inverse()
        .apply(reordered.file())
        .expect("inverse");
    assert_eq!(restored.to_bytes().unwrap(), relocated.to_bytes().unwrap());
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
fn writes_cached_formulas_through_the_unified_route() {
    let source = read("formula-budget-native.numbers");
    let source_bytes = source.to_bytes().expect("source bytes");
    let expression =
        litchi_numbers::formula::Expression::parse("=B3+1").expect("formula expression");
    let cached = litchi_numbers::formula::CachedValue::number(2.0).expect("cached number");
    let direct = litchi_numbers::Package::from_bytes(&source_bytes)
        .expect("direct leaf source")
        .edit_table_cells(
            litchi_numbers::SheetSelector::index(0),
            litchi_numbers::TableSelector::index(0),
        )
        .expect("direct leaf editor")
        .set_formula_cached(
            litchi_numbers::At::new(2, 2),
            expression.clone(),
            cached.clone(),
        )
        .expect("direct cached formula")
        .commit()
        .expect("direct cached commit")
        .into_package();
    let direct_bytes = leaf_bytes(&direct);
    let uncached = litchi_numbers::Package::from_bytes(&source_bytes)
        .expect("uncached leaf source")
        .edit_table_cells(
            litchi_numbers::SheetSelector::index(0),
            litchi_numbers::TableSelector::index(0),
        )
        .expect("uncached leaf editor")
        .set_formula(litchi_numbers::At::new(2, 2), expression)
        .expect("direct uncached formula")
        .commit()
        .expect("direct uncached commit")
        .into_package();
    let uncached_bytes = leaf_bytes(&uncached);
    let unified_uncached = source
        .set("/sheet[1]/table[1]/C3", props! { "formula" => "=B3+1" })
        .expect("uncached formula");
    assert_eq!(
        unified_uncached.to_bytes().expect("uncached unified bytes"),
        uncached_bytes,
        "formula-only authoring must remain on the uncached leaf route"
    );

    let edited = source
        .set(
            "/sheet[1]/table[1]/C3",
            props! { "formula" => "=B3+1", "cachedValue" => 2.0 },
        )
        .expect("cached formula");
    assert_eq!(
        edited.to_bytes().expect("unified bytes"),
        direct_bytes,
        "unified cached authoring must use the direct leaf route"
    );
    let reopened = reopen(&edited);
    assert_eq!(
        prop(&reopened, "/sheet[1]/table[1]/C3", "type").as_deref(),
        Some("formula")
    );
    assert!(
        prop(&reopened, "/sheet[1]/table[1]/C3", "formula")
            .expect("formula text")
            .contains("B3")
    );

    // Reapplying the exact formula and cache is a byte-exact no-op.
    let repeated = edited
        .set(
            "/sheet[1]/table[1]/C3",
            props! { "formula" => "=B3+1", "cachedValue" => 2.0 },
        )
        .expect("cached no-op");
    assert_eq!(
        repeated.to_bytes().expect("repeated bytes"),
        edited.to_bytes().expect("edited bytes")
    );

    // The public transaction patch restores the immutable source exactly.
    let transaction = source
        .edit(|edit| {
            edit.set(
                "/sheet[1]/table[1]/C3",
                props! { "formula" => "=B3+1", "cachedValue" => 2.0 },
            )
        })
        .expect("cached transaction");
    let restored = transaction
        .patch()
        .inverse()
        .apply(transaction.file())
        .expect("inverse");
    assert_eq!(restored.to_bytes().expect("restored bytes"), source_bytes);
}

#[test]
fn cached_formula_repeated_casefolded_keys_keep_last_pair_and_reject_interleaved_value() {
    let source = read("formula-budget-native.numbers");
    let source_bytes = source.to_bytes().expect("source bytes");
    let assignments = Props::from_props(props! {
        "FoRmUlA" => "not a formula".to_owned(),
        "cAcHeDvAlUe" => vec!["ignored cache".to_owned()],
        "FORMULA" => "=B3+1".to_owned(),
        "CACHEDVALUE" => 2.0,
    })
    .expect("properties");
    assert_eq!(
        assignments.get("formula").and_then(|value| value.as_text()),
        Some("=B3+1")
    );
    assert!(matches!(
        assignments.get("cachedValue"),
        Some(litchi_ops::PropValue::Number(value)) if *value == 2.0
    ));
    let edited = source
        .set("/sheet[1]/table[1]/C3", &assignments)
        .expect("last case-folded formula/cache pair");
    assert_eq!(
        edited.to_bytes().expect("edited bytes"),
        direct_cached_formula_route(&source_bytes),
        "last case-folded formula/cache pair must win"
    );

    let interleaved = Props::from_props(props! {
        "formula" => "=B3+1",
        "cachedValue" => 2.0,
        "VaLuE" => "interleaved value".to_owned(),
    })
    .expect("interleaved properties");
    assert_eq!(
        interleaved.get("value").and_then(|value| value.as_text()),
        Some("interleaved value")
    );
    let error = source
        .set("/sheet[1]/table[1]/C3", &interleaved)
        .expect_err("a value and formula cache on one cell are not one leaf change");
    assert_eq!(error.code(), Code::InvalidValue);
    assert_eq!(source.to_bytes().expect("unchanged source"), source_bytes);
}

#[test]
fn cached_formula_operation_limit_precedes_parse_and_cache_allocations() {
    let source = read("formula-budget-native.numbers");
    let before = source.to_bytes().expect("source bytes");
    let limits = SemanticLimits::builder()
        .max_operations(1)
        .expect("one-operation limit")
        .build()
        .expect("valid limits");
    let malformed_and_large = props! {
        "formula" => format!("not a formula {}", "x".repeat(128 * 1024)),
        "cachedValue" => "x".repeat(128 * 1024),
    };
    let error = source
        .edit_with_options(
            litchi_ops::EditOptions::new().semantic_limits(limits),
            |edit| edit.set("/sheet[1]/table[1]/C3", malformed_and_large),
        )
        .expect_err("operation admission must precede formula/cache parsing");
    assert_eq!(error.code(), Code::LimitExceeded);
    assert!(matches!(
        error.leaf(),
        Source::Transaction(TransactionError::Limit {
            kind: LimitKind::Operations,
            observed: 2,
            maximum: 1,
        })
    ));
    assert_eq!(source.to_bytes().expect("unchanged source"), before);
}

/// Full `File` versus direct-leaf release instruction probe for one cached
/// formula write.  Run explicitly under pinned callgrind; no wall-clock
/// timing is part of this probe.
#[test]
#[ignore = "release callgrind ABBA probe; run explicitly"]
fn cached_formula_full_file_cost_abba() {
    let source = std::fs::read(fixture("formula-budget-native.numbers"))
        .expect("read formula budget fixture");
    let direct = cached_formula_direct_setup(&source);
    assert_eq!(
        cached_formula_unified_setup(&source),
        direct,
        "full unified output must equal direct leaf output"
    );
    for route in [0_u8, 1, 1, 0] {
        let output = if route == 0 {
            unified_cached_formula_route(&source)
        } else {
            direct_cached_formula_route(&source)
        };
        assert_eq!(output, direct, "cached formula ABBA output must be exact");
        std::hint::black_box(output);
    }
}

#[test]
fn cached_formula_inputs_refuse_without_mutation() {
    let source = read("basic.numbers");
    let before = source.to_bytes().expect("source bytes");
    let cached_only = source
        .set("/sheet[1]/table[1]/C3", props! { "cachedValue" => 43.0 })
        .expect_err("cache without formula");
    assert_eq!(cached_only.code(), Code::InvalidArgument);
    let null_formula = source
        .set(
            "/sheet[1]/table[1]/C3",
            props! {
                "formula" => litchi_ops::PropValue::Null,
                "cachedValue" => 43.0
            },
        )
        .expect_err("null formula with cache");
    assert_eq!(null_formula.code(), Code::InvalidArgument);
    let null_cache = source
        .set(
            "/sheet[1]/table[1]/C3",
            props! {
                "formula" => "=B3+1",
                "cachedValue" => litchi_ops::PropValue::Null
            },
        )
        .expect_err("null cache");
    assert_eq!(null_cache.code(), Code::InvalidValue);
    let unsupported_cache = source
        .set(
            "/sheet[1]/table[1]/C3",
            props! {
                "formula" => "=B3+1",
                "cachedValue" => vec!["not a scalar".to_owned()]
            },
        )
        .expect_err("unsupported cache kind");
    assert_eq!(unsupported_cache.code(), Code::InvalidValue);
    let non_finite = source
        .set(
            "/sheet[1]/table[1]/C3",
            props! { "formula" => "=B3+1", "cachedValue" => f64::NAN },
        )
        .expect_err("non-finite cache");
    assert_eq!(non_finite.code(), Code::InvalidValue);
    assert_eq!(source.to_bytes().expect("unchanged source"), before);
}

#[test]
fn sets_display_formats() {
    let file = read("basic.numbers");
    let edited = file
        .set(
            "/sheet[1]/table[1]/B3",
            props! { "numberFormat" => "#,##0.00" },
        )
        .expect("decimal format");
    let package = leaf(&edited);
    let format = package
        .table_cell_number_format(0, 0, litchi_numbers::At::new(2, 1))
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
}

/// Record 0318: cell comments are read, set, added and removed, on native
/// packages (their lists, tiles and authors in members of their own) and
/// on a created workbook.
#[test]
fn cell_comments_are_read_set_added_and_removed() {
    let native = read("comment-reader-native-created.numbers");
    let commented = native
        .query("cell")
        .expect("query")
        .filter_map(|node| node.ok())
        .find(|node| {
            node.view()
                .format()
                .expect("format")
                .into_iter()
                .any(|prop| prop.key() == "comment")
        })
        .map(|node| node.path().to_string())
        .expect("a natively commented cell");
    // Replace, clear, and set again.
    let replaced = native
        .set(&commented, props! { "comment" => "replaced" })
        .expect("replace a native comment");
    assert_eq!(
        prop(&reopen(&replaced), &commented, "comment").as_deref(),
        Some("replaced")
    );
    let cleared = replaced
        .set(
            &commented,
            props! { "comment" => litchi_ops::PropValue::Null },
        )
        .expect("clear");
    assert_eq!(prop(&reopen(&cleared), &commented, "comment"), None);
    let removed = native
        .remove(&format!("{commented}/comment"))
        .expect("remove the comment");
    assert_eq!(prop(&reopen(&removed), &commented, "comment"), None);
    // Record 0329: an empty cell in a row its tile does not store takes a
    // comment too.
    let basic = read("basic.numbers");
    let last = "/sheet[1]/table[1]/G22";
    let commented_empty = basic
        .set(last, props! { "comment" => "on an empty cell" })
        .expect("comment an empty cell");
    assert_eq!(
        prop(&reopen(&commented_empty), last, "comment").as_deref(),
        Some("on an empty cell")
    );

    for file in [
        read("basic.numbers"),
        create(FileFormat::Numbers)
            .expect("create")
            .set("/sheet[1]/table[1]/B2", props! { "value" => 1 })
            .expect("value"),
    ] {
        let added = file
            .add(
                "/sheet[1]/table[1]/B2",
                NewNode::new("comment").prop("text", "check me"),
            )
            .expect("add a comment");
        assert_eq!(added.path.to_string(), "/sheet[1]/table[1]/B2");
        let reopened = reopen(&added.file);
        assert_eq!(
            prop(&reopened, "/sheet[1]/table[1]/B2", "comment").as_deref(),
            Some("check me")
        );
        let edit = file
            .edit(|tx| {
                tx.add(
                    "/sheet[1]/table[1]/B2",
                    NewNode::new("comment").prop("text", "check me"),
                )
                .map(|_| ())
            })
            .expect("add in a transaction");
        let restored = edit.patch().inverse().apply(edit.file()).expect("inverse");
        assert_eq!(restored.to_bytes().unwrap(), file.to_bytes().unwrap());
        let error = reopened
            .add(
                "/sheet[1]/table[1]/B2",
                NewNode::new("comment").prop("text", "twice"),
            )
            .expect_err("one comment per cell");
        assert_eq!(error.code(), Code::InvalidArgument);
        let edited = reopened
            .set("/sheet[1]/table[1]/B2", props! { "comment" => "edited" })
            .expect("edit");
        assert_eq!(
            prop(&reopen(&edited), "/sheet[1]/table[1]/B2", "comment").as_deref(),
            Some("edited")
        );
        let gone = edited
            .remove("/sheet[1]/table[1]/B2/comment")
            .expect("remove");
        assert_eq!(
            prop(&reopen(&gone), "/sheet[1]/table[1]/B2", "comment"),
            None
        );
    }
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
    assert_eq!(commit.diagnostics().operations(), 40);
    assert_eq!(commit.diagnostics().commits(), 1);
    let reopened = reopen(commit.file());
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
    assert_eq!(commit.diagnostics().commits(), 2);
    assert_eq!(
        prop(commit.file(), "/sheet[1]/table[1]/C4", "value").as_deref(),
        Some("2")
    );
    let restored = commit
        .patch()
        .inverse()
        .apply(commit.file())
        .expect("inverse");
    assert_eq!(restored.to_bytes().unwrap(), file.to_bytes().unwrap());
}

#[test]
fn sheets_and_tables_are_added_and_removed() {
    let file = create(FileFormat::Numbers).expect("create");
    // A sheet, placed first, with a default name.
    let added = file
        .add("/", NewNode::new("sheet").at(Placement::Start))
        .expect("add sheet");
    assert_eq!(added.path.to_string(), "/sheet[1]");
    let reopened = reopen(&added.file);
    assert_eq!(
        prop(&reopened, "/sheet[1]", "name").as_deref(),
        Some("Sheet 2")
    );
    assert_eq!(prop(&reopened, "/", "sheetCount").as_deref(), Some("2"));
    // A table in it, sized and named.
    let table = added
        .file
        .add(
            "/sheet[1]",
            NewNode::new("table")
                .prop("name", "Data")
                .prop("rowCount", 4_i64)
                .prop("columnCount", 3_i64),
        )
        .expect("add table");
    assert_eq!(table.path.to_string(), "/sheet[1]/table[1]");
    let reopened = reopen(&table.file);
    assert_eq!(
        prop(&reopened, "/sheet[1]/table[1]", "name").as_deref(),
        Some("Data")
    );
    assert_eq!(
        prop(&reopened, "/sheet[1]/table[1]", "rowCount").as_deref(),
        Some("4")
    );
    assert_eq!(
        prop(&reopened, "/sheet[1]/table[1]", "columnCount").as_deref(),
        Some("3")
    );
    // It takes values.
    let written = table
        .file
        .set("/sheet[1]/table[1]/B2", props! { "value" => 7_i64 })
        .expect("value");
    assert_eq!(
        prop(&reopen(&written), "/sheet[1]/table[1]/B2", "value").as_deref(),
        Some("7")
    );
    // A second table placed before the first.
    let before = table
        .file
        .add(
            "/sheet[1]",
            NewNode::new("table").at(Placement::before("/sheet[1]/table[1]").expect("placement")),
        )
        .expect("placed table");
    assert_eq!(before.path.to_string(), "/sheet[1]/table[1]");
    assert_eq!(
        prop(&reopen(&before.file), "/sheet[1]/table[2]", "name").as_deref(),
        Some("Data")
    );
    // Remove the table, then the sheet; each inverse restores the bytes.
    let removed = table
        .file
        .edit(|tx| tx.remove("/sheet[1]/table[1]"))
        .expect("remove table");
    assert_eq!(
        prop(&reopen(removed.file()), "/sheet[1]", "tableCount").as_deref(),
        Some("0")
    );
    let restored = removed
        .patch()
        .inverse()
        .apply(removed.file())
        .expect("inverse");
    assert_eq!(restored.to_bytes().unwrap(), table.file.to_bytes().unwrap());
    let gone = table
        .file
        .remove("/sheet[1]")
        .expect("remove sheet with its table");
    let reopened = reopen(&gone);
    assert_eq!(prop(&reopened, "/", "sheetCount").as_deref(), Some("1"));
    // The last sheet stays.
    let error = reopened.remove("/sheet[1]").expect_err("last sheet");
    assert_eq!(error.code(), Code::InvalidArgument);
    let error = file
        .add("/", NewNode::new("sheet").prop("name", "Sheet 1"))
        .expect_err("duplicate");
    assert_eq!(error.code(), Code::InvalidValue);
}

#[test]
fn corpus_tables_round_trip_through_add_and_remove() {
    let mut round_trips = 0;
    for entry in std::fs::read_dir(fixture("")).expect("corpus") {
        let path = entry.expect("entry").path();
        if path
            .extension()
            .is_none_or(|extension| extension != "numbers")
        {
            continue;
        }
        let file = litchi_ops::open_bytes(std::fs::read(&path).expect("read")).expect("open");
        let before = snapshot(&file);
        let added = file
            .add("/sheet[1]", NewNode::new("table"))
            .unwrap_or_else(|error| panic!("{}: {error}", path.display()));
        match added.file.remove(added.path.to_string().as_str()) {
            Ok(removed) => {
                assert_eq!(snapshot(&reopen(&removed)), before, "{}", path.display());
                round_trips += 1;
            },
            Err(error) => assert_eq!(
                error.code(),
                Code::UnsupportedCapability,
                "{}",
                path.display()
            ),
        }
    }
    assert!(round_trips > 30, "{round_trips}");
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

#[test]
fn rows_and_columns_are_inserted_and_removed() {
    let file = create(FileFormat::Numbers).expect("create");
    let filled = file
        .edit(|tx| {
            tx.set("/sheet[1]/table[1]/B2", props! { "value" => "b2" })?;
            tx.set("/sheet[1]/table[1]/B3", props! { "value" => 3 })?;
            tx.set("/sheet[1]/table[1]/D3", props! { "value" => "d3" })?;
            Ok(())
        })
        .expect("values")
        .into_file();
    let rows = prop(&filled, "/sheet[1]/table[1]", "rowCount").expect("rows");
    let rows: usize = rows.parse().expect("count");
    // A row before row 3 (zero-based position 2), then filled.
    let added = filled
        .add(
            "/sheet[1]/table[1]",
            NewNode::new("row").at(Placement::Position(litchi_core::Position::new(2))),
        )
        .expect("insert a row");
    assert_eq!(added.path.to_string(), "/sheet[1]/table[1]/row[3]");
    let reopened = reopen(&added.file);
    assert_eq!(
        prop(&reopened, "/sheet[1]/table[1]", "rowCount"),
        Some((rows + 1).to_string())
    );
    assert_eq!(
        prop(&reopened, "/sheet[1]/table[1]/B4", "value").as_deref(),
        Some("3")
    );
    assert_eq!(
        prop(&reopened, "/sheet[1]/table[1]/B2", "value").as_deref(),
        Some("b2")
    );
    let written = reopened
        .set("/sheet[1]/table[1]/A3", props! { "value" => "new" })
        .expect("the new row takes a value");
    assert_eq!(
        prop(&reopen(&written), "/sheet[1]/table[1]/A3", "value").as_deref(),
        Some("new")
    );
    // The inverse restores the bytes.
    let edited = filled
        .edit(|tx| {
            tx.add("/sheet[1]/table[1]", NewNode::new("row"))
                .map(|_| ())
        })
        .expect("append a row");
    let restored = edited
        .patch()
        .inverse()
        .apply(edited.file())
        .expect("inverse");
    assert_eq!(restored.to_bytes().unwrap(), filled.to_bytes().unwrap());
    // Removing row 3 drops its cells and moves the rest up.
    let removed = filled
        .remove("/sheet[1]/table[1]/row[3]")
        .expect("remove a row");
    let reopened = reopen(&removed);
    assert_eq!(
        prop(&reopened, "/sheet[1]/table[1]", "rowCount"),
        Some((rows - 1).to_string())
    );
    assert!(
        reopened.get("/sheet[1]/table[1]/D3").is_err()
            || prop(&reopened, "/sheet[1]/table[1]/D3", "value").is_none()
    );
    // A column appended, one placed first, and one removed.
    let appended = filled
        .add("/sheet[1]/table[1]", NewNode::new("col"))
        .expect("append a column");
    let columns = prop(&filled, "/sheet[1]/table[1]", "columnCount").expect("columns");
    let columns: usize = columns.parse().expect("count");
    assert_eq!(
        appended.path.to_string(),
        format!("/sheet[1]/table[1]/col[{}]", columns + 1)
    );
    let first = filled
        .add(
            "/sheet[1]/table[1]",
            NewNode::new("column").at(Placement::Start),
        )
        .expect("a header column");
    let reopened = reopen(&first.file);
    assert_eq!(
        prop(&reopened, "/sheet[1]/table[1]/C2", "value").as_deref(),
        Some("b2")
    );
    let dropped = filled
        .remove("/sheet[1]/table[1]/col[2]")
        .expect("remove a column");
    let reopened = reopen(&dropped);
    assert_eq!(
        prop(&reopened, "/sheet[1]/table[1]/C3", "value").as_deref(),
        Some("d3")
    );
    // Typed refusals.
    let error = filled
        .remove(&format!("/sheet[1]/table[1]/row[{}]", rows + 5))
        .expect_err("no such row");
    assert_eq!(error.code(), Code::PathNotFound);
    let error = filled
        .add("/sheet[1]", NewNode::new("row"))
        .expect_err("rows go under a table");
    assert_eq!(error.code(), Code::InvalidArgument);
    let error = read("formula-arena-native.numbers")
        .add("/sheet[1]/table[1]", NewNode::new("row"))
        .expect_err("formulas");
    assert_eq!(error.code(), Code::UnsupportedCapability);
}

/// Record 0318: rows and columns of a table with formulas move the
/// formulas' references; removing a row a formula reads is refused.
#[test]
fn rows_and_columns_move_formula_references() {
    let file = create(FileFormat::Numbers)
        .expect("create")
        .edit(|tx| {
            tx.set("/sheet[1]/table[1]/B2", props! { "value" => 1 })?;
            tx.set("/sheet[1]/table[1]/B3", props! { "value" => 2 })?;
            tx.set("/sheet[1]/table[1]/C3", props! { "value" => 3 })?;
            tx.set(
                "/sheet[1]/table[1]/D2",
                props! { "formula" => "=SUM($B$2:B3)+C3" },
            )?;
            Ok(())
        })
        .expect("formulas")
        .into_file();
    let added = file
        .add(
            "/sheet[1]/table[1]",
            NewNode::new("row").at(Placement::Position(litchi_core::Position::new(2))),
        )
        .expect("insert a row under formulas");
    let reopened = reopen(&added.file);
    assert_eq!(
        prop(&reopened, "/sheet[1]/table[1]/D2", "formula").as_deref(),
        Some("=(SUM($B$2:B4)+C4)")
    );
    let column = file
        .add(
            "/sheet[1]/table[1]",
            NewNode::new("col").at(Placement::Position(litchi_core::Position::new(2))),
        )
        .expect("insert a column under formulas");
    assert_eq!(
        prop(&reopen(&column.file), "/sheet[1]/table[1]/E2", "formula").as_deref(),
        Some("=(SUM($B$2:B3)+D3)")
    );
    let edited = file
        .edit(|tx| {
            tx.add("/sheet[1]/table[1]", NewNode::new("row"))
                .map(|_| ())
        })
        .expect("append a row");
    let restored = edited
        .patch()
        .inverse()
        .apply(edited.file())
        .expect("inverse");
    assert_eq!(restored.to_bytes().unwrap(), file.to_bytes().unwrap());
    // Record 0329: removing row 3, which the formula reads, shrinks its
    // range and turns the reference to C3 into #REF!.
    let removed = file
        .remove("/sheet[1]/table[1]/row[3]")
        .expect("remove a read row");
    assert_eq!(
        prop(&reopen(&removed), "/sheet[1]/table[1]/D2", "formula").as_deref(),
        Some("=(SUM($B$2:B2)+#REF!)")
    );
}

#[test]
fn cell_fonts_are_written() {
    let file = create(FileFormat::Numbers)
        .expect("create")
        .set("/sheet[1]/table[1]/B2", props! { "value" => "styled" })
        .expect("value");
    let styled = file
        .set(
            "/sheet[1]/table[1]/B2",
            props! {
                "font.bold" => true,
                "font.italic" => true,
                "font.size" => "16pt",
                "font.color" => "FF0000",
                "font.name" => "Helvetica",
            },
        )
        .expect("font");
    let package = leaf(&styled);
    let style = package
        .table_cell_text_style(0, 0, litchi_numbers::At::new(1, 1))
        .expect("style");
    assert_eq!(style.bold, Some(true));
    assert_eq!(style.italic, Some(true));
    assert_eq!(style.size, Some(16.0));
    assert_eq!(style.color, Some([255, 0, 0]));
    assert_eq!(style.font_name.as_deref(), Some("Helvetica"));
    assert_eq!(
        prop(&reopen(&styled), "/sheet[1]/table[1]/B2", "value").as_deref(),
        Some("styled")
    );
    let batch = file
        .batch(
            vec![
                batch_set(
                    "/sheet[1]/table[1]/B2".to_owned(),
                    Props::new().with("font.bold", true),
                ),
                batch_set(
                    "/sheet[1]/table[1]/C3".to_owned(),
                    Props::new().with("value", 3_i64),
                ),
            ],
            BatchMode::Atomic,
        )
        .expect("batch");
    let package = leaf(batch.file());
    assert_eq!(
        package
            .table_cell_text_style(0, 0, litchi_numbers::At::new(1, 1))
            .expect("style")
            .bold,
        Some(true)
    );
    // An empty cell takes a style-only cell (record 0370).
    let styled_empty = file
        .set("/sheet[1]/table[1]/D4", props! { "font.bold" => true })
        .expect("empty cell");
    assert_eq!(
        leaf(&styled_empty)
            .table_cell_text_style(0, 0, litchi_numbers::At::new(3, 3))
            .expect("style")
            .bold,
        Some(true)
    );
}

/// Three sheets; the first holds the tables `Table 1`, `B` and `C`, the
/// second none, the third `Z` (record 0347).
fn three_sheets() -> File {
    let mut file = create(FileFormat::Numbers).expect("create");
    for _ in 0..2 {
        file = file
            .add("/", NewNode::new("sheet").at(Placement::End))
            .expect("add sheet")
            .into_file();
    }
    for (sheet, name) in [("/sheet[1]", "B"), ("/sheet[1]", "C"), ("/sheet[3]", "Z")] {
        file = file
            .add(sheet, NewNode::new("table").prop("name", name))
            .expect("add table")
            .into_file();
    }
    file
}

fn names(file: &File, paths: &[&str]) -> Vec<Option<String>> {
    let file = reopen(file);
    paths.iter().map(|path| prop(&file, path, "name")).collect()
}

/// Sheets and tables swap atomically, each swap equals its two moves
/// applied one at a time, and the patch inverts to the exact source bytes
/// (record 0347).
#[test]
fn swaps_sheets_and_tables() {
    let file = three_sheets();
    let source = file.to_bytes().unwrap();
    let sheets = ["/sheet[1]", "/sheet[2]", "/sheet[3]"];

    // Sheets: the first takes the third's place, then the third the first's.
    let swapped = file
        .edit(|tx| tx.swap("/sheet[1]", "/sheet[3]"))
        .expect("swap sheets");
    assert_eq!(
        names(swapped.file(), &sheets),
        [
            Some("Sheet 3".into()),
            Some("Sheet 2".into()),
            Some("Sheet 1".into())
        ]
    );
    let moved = file
        .move_to("/sheet[1]", Placement::End)
        .and_then(|file| file.move_to("/sheet[2]", Placement::Start))
        .expect("two moves");
    assert_eq!(
        swapped.file().to_bytes().unwrap(),
        moved.to_bytes().unwrap()
    );
    let reversed = file
        .swap("/sheet[3]", "/sheet[1]")
        .expect("reversed operands");
    assert_eq!(reversed.to_bytes().unwrap(), moved.to_bytes().unwrap());
    let restored = swapped
        .patch()
        .inverse()
        .apply(swapped.file())
        .expect("inverse");
    assert_eq!(restored.to_bytes().unwrap(), source);
    // An element swapped with itself is an exact no-op.
    let same = file
        .edit(|tx| tx.swap("/sheet[2]", "/Sheet 2"))
        .expect("self swap");
    assert!(!same.diagnostics().changed());
    assert_eq!(same.file().to_bytes().unwrap(), source);

    // Tables of one sheet.
    let tables = [
        "/sheet[1]/table[1]",
        "/sheet[1]/table[2]",
        "/sheet[1]/table[3]",
    ];
    let swapped = file
        .edit(|tx| tx.swap("/sheet[1]/table[1]", "/sheet[1]/table[3]"))
        .expect("swap tables");
    assert_eq!(
        names(swapped.file(), &tables),
        [Some("C".into()), Some("B".into()), Some("Table 1".into())]
    );
    let moved = file
        .move_to("/sheet[1]/table[1]", Placement::End)
        .and_then(|file| file.move_to("/sheet[1]/table[2]", Placement::Start))
        .expect("two moves");
    assert_eq!(
        swapped.file().to_bytes().unwrap(),
        moved.to_bytes().unwrap()
    );
    let restored = swapped
        .patch()
        .inverse()
        .apply(swapped.file())
        .expect("inverse");
    assert_eq!(restored.to_bytes().unwrap(), source);

    // Tables of two sheets: each takes the other's place.
    let swapped = file
        .edit(|tx| tx.swap("/sheet[1]/table[2]", "/sheet[3]/table[1]"))
        .expect("swap across sheets");
    assert_eq!(
        names(
            swapped.file(),
            &[
                "/sheet[1]/table[1]",
                "/sheet[1]/table[2]",
                "/sheet[1]/table[3]",
                "/sheet[3]/table[1]"
            ]
        ),
        [
            Some("Table 1".into()),
            Some("Z".into()),
            Some("C".into()),
            Some("B".into())
        ]
    );
    let moved = file
        .move_to(
            "/sheet[1]/table[2]",
            Placement::inside("/sheet[3]").expect("placement"),
        )
        .and_then(|file| file.move_to("/sheet[3]/table[1]", Placement::inside("/sheet[1]")?))
        .and_then(|file| {
            file.move_to(
                "/sheet[1]/table[3]",
                Placement::Position(litchi_core::Position::new(1)),
            )
        })
        .and_then(|file| file.move_to("/sheet[3]/table[1]", Placement::Start))
        .expect("relocations and moves");
    assert_eq!(
        swapped.file().to_bytes().unwrap(),
        moved.to_bytes().unwrap()
    );
    let restored = swapped
        .patch()
        .inverse()
        .apply(swapped.file())
        .expect("inverse");
    assert_eq!(restored.to_bytes().unwrap(), source);
    // Two tables of one name in two sheets.
    let renamed = file
        .set("/sheet[3]/table[1]", props! { "name" => "Table 1" })
        .expect("rename");
    // A move into a sheet that holds the name is refused, typed.
    let error = renamed
        .move_to(
            "/sheet[1]/table[1]",
            Placement::inside("/sheet[3]").expect("placement"),
        )
        .expect_err("duplicate name");
    assert_eq!(error.code(), Code::InvalidValue);
    let swapped = renamed
        .edit(|tx| tx.swap("/sheet[1]/table[1]", "/sheet[3]/table[1]"))
        .expect("same-named tables");
    let reopened = reopen(swapped.file());
    assert_eq!(
        prop(&reopened, "/sheet[1]", "tableCount").as_deref(),
        Some("3")
    );
    for path in ["/sheet[1]/table[1]", "/sheet[3]/table[1]"] {
        assert_eq!(prop(&reopened, path, "name").as_deref(), Some("Table 1"));
    }
    // The other's tables moved with it: `B` and `C` stay in sheet 1, and
    // the swapped tables carry their cells.
    assert_eq!(
        prop(&reopened, "/sheet[1]/table[2]", "name").as_deref(),
        Some("B")
    );
    let restored = swapped
        .patch()
        .inverse()
        .apply(swapped.file())
        .expect("inverse");
    assert_eq!(restored.to_bytes().unwrap(), renamed.to_bytes().unwrap());
    // A third table of the name in the destination sheet is refused.
    let error = renamed
        .set("/sheet[1]/table[2]", props! { "name" => "Z" })
        .and_then(|file| file.swap("/sheet[1]/table[2]", "/sheet[3]/table[1]"))
        .map(|_| ())
        .expect_err("third name");
    assert_eq!(error.code(), Code::InvalidValue);
}

/// What a Numbers swap refuses, typed and without touching the snapshot.
#[test]
fn swap_refusals_are_typed() {
    let file = three_sheets()
        .set("/sheet[1]/table[1]/A1", props! { "value" => 1_i64 })
        .and_then(|file| file.set("/sheet[1]/table[1]/A2", props! { "value" => 2_i64 }))
        .expect("values");
    let source = file.to_bytes().unwrap();
    for (left, right, code) in [
        ("/sheet[1]", "/sheet[1]/table[1]", Code::InvalidArgument),
        (
            "/sheet[1]/table[1]/row[1]",
            "/sheet[1]/table[1]/row[2]",
            Code::UnsupportedCapability,
        ),
        (
            "/sheet[1]/table[1]/col[1]",
            "/sheet[1]/table[1]/col[2]",
            Code::UnsupportedCapability,
        ),
        (
            "/sheet[1]/table[1]/A1",
            "/sheet[1]/table[1]/A2",
            Code::UnsupportedCapability,
        ),
        ("/sheet[1]", "/sheet[9]", Code::PathNotFound),
    ] {
        let error = file.swap(left, right).expect_err(left);
        assert_eq!(error.code(), code, "{left} {right}: {error}");
    }
    // An atomic batch whose second item fails publishes neither.
    let error = file
        .edit(|tx| {
            tx.swap("/sheet[1]", "/sheet[2]")?;
            tx.swap("/sheet[1]", "/sheet[9]")
        })
        .expect_err("second swap");
    assert_eq!(error.code(), Code::PathNotFound);
    assert_eq!(file.to_bytes().unwrap(), source);
}
