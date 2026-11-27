//! Unified XLSX row and column dimension writes.
#![cfg(feature = "xlsx")]

use litchi_core::selector::{Path as SelectorPath, Query, Scope};
use litchi_ops::transaction::SemanticLimits;
use litchi_ops::{Code, File, OpenOptions, PropValue, Props, SourceMode, props};

const ROW3: &str = "/sheet[1]/row[3]";
const ROW5: &str = "/sheet[1]/row[5]";
const COL_B: &str = "/sheet[1]/col[B]";
const COL_D: &str = "/sheet[1]/col[D]";

fn fixture() -> Vec<u8> {
    let workbook = litchi_xlsx::Package::new().expect("create native workbook");
    let mut edit = workbook.edit().expect("start native workbook edit");
    let mut sheet = edit
        .sheet("Sheet1")
        .expect("find Sheet1")
        .expect("Sheet1 worksheet");
    sheet.set("A1", "row one").expect("seed row one");
    sheet.set("A3", "row three").expect("seed row three");
    sheet.set("B1", "column B").expect("seed column B");
    sheet
        .column("B")
        .expect("seed explicit column B record")
        .width(10.0)
        .expect("seed column B width");
    edit.commit()
        .expect("commit dimension fixture")
        .into_package()
        .to_bytes()
        .expect("serialize dimension fixture")
}

fn open_owner(bytes: &[u8], eager: bool) -> File {
    let mode = if eager {
        SourceMode::Eager
    } else {
        SourceMode::Auto
    };
    litchi_ops::open_bytes_with(bytes.to_vec(), &OpenOptions::new().source_mode(mode))
        .expect("open XLSX owner")
}

fn format_property(file: &File, path: &str, key: &str) -> Option<String> {
    file.get(path)
        .expect("resolve dimension node")
        .view()
        .format()
        .expect("read dimension properties")
        .find(|property| property.key() == key)
        .map(|property| property.value().to_string())
}

fn root_query(selector: &str) -> Query<'_> {
    Query::parse(selector)
        .expect("parse dimension selector")
        .with_scope(Scope::Path(Box::new(
            SelectorPath::parse("/").expect("root selector scope"),
        )))
}

#[test]
fn implicit_rows_and_columns_accept_typed_dimension_writes_and_survive_reopen() {
    let bytes = fixture();
    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let original_bytes = original.to_bytes().expect("serialize original");

        // Row 5 and column D have no source records in the fixture. Their
        // addressing paths still select checked logical dimensions and the
        // native editor materializes only the requested records.
        let row = original
            .set(ROW5, props! { "height" => "24pt", "hidden" => true })
            .expect("write an implicit row dimension");
        assert_eq!(
            format_property(&row, ROW5, "height").as_deref(),
            Some("24pt")
        );
        assert_eq!(
            format_property(&row, ROW5, "hidden").as_deref(),
            Some("true")
        );

        let column = row
            .set(COL_D, props! { "width" => 12.5_f64, "hidden" => true })
            .expect("write an implicit column dimension");
        assert_eq!(
            format_property(&column, COL_D, "width").as_deref(),
            Some("12.5")
        );
        assert_eq!(
            format_property(&column, COL_D, "hidden").as_deref(),
            Some("true")
        );

        let published = column.to_bytes().expect("publish dimension edits");
        let reopened = open_owner(&published, eager);
        assert_eq!(
            format_property(&reopened, ROW5, "height").as_deref(),
            Some("24pt")
        );
        assert_eq!(
            format_property(&reopened, ROW5, "hidden").as_deref(),
            Some("true")
        );
        assert_eq!(
            format_property(&reopened, COL_D, "width").as_deref(),
            Some("12.5")
        );
        assert_eq!(
            format_property(&reopened, COL_D, "hidden").as_deref(),
            Some("true")
        );

        let published_again = column
            .set(ROW5, props! { "height" => "24pt", "hidden" => true })
            .expect("repeat the same row dimensions")
            .set(COL_D, props! { "width" => 12.5_f64, "hidden" => true })
            .expect("repeat the same column dimensions")
            .to_bytes()
            .expect("publish repeated dimensions");
        assert_eq!(
            published_again, published,
            "repeating identical dimensions is an exact no-op"
        );
        assert_ne!(
            original_bytes, published,
            "dimension edits must differ from the untouched source"
        );
    }
}

#[test]
fn scoped_dimension_selectors_deduplicate_rows_and_columns() {
    let bytes = fixture();
    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let rows = root_query("row[3], row[3]");
        let row_matches = original
            .query(rows.clone())
            .expect("run duplicate row selector")
            .collect::<litchi_ops::Result<Vec<_>>>()
            .expect("read duplicate row matches");
        assert_eq!(
            row_matches.len(),
            1,
            "selector unions deduplicate row nodes"
        );
        let row = original
            .set(rows, props! { "height" => "18pt", "hidden" => true })
            .expect("set one deduplicated row target");
        assert_eq!(
            format_property(&row, ROW3, "height").as_deref(),
            Some("18pt")
        );
        assert_eq!(
            format_property(&row, ROW3, "hidden").as_deref(),
            Some("true")
        );

        let columns = root_query("col[B], col[B]");
        let column_matches = row
            .query(columns.clone())
            .expect("run duplicate column selector")
            .collect::<litchi_ops::Result<Vec<_>>>()
            .expect("read duplicate column matches");
        assert_eq!(
            column_matches.len(),
            1,
            "selector unions deduplicate column nodes"
        );
        let column = row
            .set(columns, props! { "width" => 14.0_f64, "hidden" => true })
            .expect("set one deduplicated column target");
        assert_eq!(
            format_property(&column, COL_B, "width").as_deref(),
            Some("14")
        );
        assert_eq!(
            format_property(&column, COL_B, "hidden").as_deref(),
            Some("true")
        );
    }
}

#[test]
fn dimension_transaction_replays_inverses_and_keeps_source_immutable() {
    let bytes = fixture();
    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let original_bytes = original.to_bytes().expect("serialize original");
        let commit = original
            .edit(|transaction| {
                transaction.set(ROW3, props! { "height" => "22pt", "hidden" => true })?;
                transaction.set(COL_B, props! { "width" => 16.0_f64, "hidden" => true })?;
                Ok::<_, litchi_ops::Error>(())
            })
            .expect("stage row and column dimensions");
        assert!(commit.diagnostics().changed());
        assert_eq!(commit.diagnostics().operations(), 2);
        assert_eq!(
            format_property(commit.file(), ROW3, "height").as_deref(),
            Some("22pt")
        );
        assert_eq!(
            format_property(commit.file(), COL_B, "width").as_deref(),
            Some("16")
        );

        let replayed = commit
            .patch()
            .apply(&original)
            .expect("replay dimension patch");
        assert_eq!(
            replayed.to_bytes().expect("serialize replay"),
            commit.file().to_bytes().expect("serialize target")
        );
        let restored = commit
            .patch()
            .inverse()
            .apply(commit.file())
            .expect("inverse dimension patch");
        assert_eq!(
            restored.to_bytes().expect("serialize inverse"),
            original_bytes
        );
        assert_eq!(
            original.to_bytes().expect("serialize unchanged source"),
            original_bytes
        );
    }
}

#[test]
fn dimension_target_limits_and_mixed_families_refuse_atomically() {
    let bytes = fixture();
    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let original_bytes = original.to_bytes().expect("serialize original");

        let limits = SemanticLimits::builder()
            .max_expanded_targets(1)
            .expect("expanded target limit")
            .build()
            .expect("valid semantic limits");
        let too_many = original
            .edit_with_options(
                litchi_ops::EditOptions::new().semantic_limits(limits),
                |transaction| {
                    transaction.set(root_query("row[1], row[3]"), props! { "height" => "20pt" })?;
                    Ok::<_, litchi_ops::Error>(())
                },
            )
            .expect_err("two row targets exceed the configured bound");
        assert_eq!(too_many.code(), Code::LimitExceeded);
        assert_eq!(
            original.to_bytes().expect("serialize after target refusal"),
            original_bytes
        );

        let mixed = original
            .set(
                root_query("row[1], cell"),
                Props::new().with("height", "20pt"),
            )
            .expect_err("row and cell targets cannot share a dimension write");
        assert_eq!(mixed.code(), Code::UnsupportedElement);
        assert_eq!(
            original.to_bytes().expect("serialize after mixed refusal"),
            original_bytes
        );
    }
}

#[test]
fn dimension_outline_aliases_and_collapsed_round_trip_for_rows_and_columns() {
    let bytes = fixture();
    for eager in [false, true] {
        let original = open_owner(&bytes, eager);

        for (alias, level) in [
            ("outline", 2_i32),
            ("group", 3_i32),
            ("outlineLevel", 4_i32),
        ] {
            let row = original
                .set(
                    ROW3,
                    Props::new().with(alias, level).with("collapsed", true),
                )
                .expect("row outline aliases are accepted");
            assert_eq!(
                format_property(&row, ROW3, "outlineLevel"),
                Some(level.to_string())
            );
            assert_eq!(
                format_property(&row, ROW3, "collapsed").as_deref(),
                Some("true")
            );

            let column = original
                .set(
                    COL_B,
                    Props::new().with(alias, level).with("collapsed", true),
                )
                .expect("column outline aliases are accepted");
            assert_eq!(
                format_property(&column, COL_B, "outlineLevel"),
                Some(level.to_string())
            );
            assert_eq!(
                format_property(&column, COL_B, "collapsed").as_deref(),
                Some("true")
            );
        }
    }
}

#[test]
fn null_resets_all_row_and_column_dimension_properties() {
    let bytes = fixture();
    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let configured = original
            .set(
                ROW3,
                props! {
                    "height" => "22pt",
                    "hidden" => true,
                    "outline" => 2,
                    "collapsed" => true,
                },
            )
            .expect("configure row dimensions")
            .set(
                COL_B,
                props! {
                    "width" => 18.0_f64,
                    "hidden" => true,
                    "group" => 3,
                    "collapsed" => true,
                },
            )
            .expect("configure column dimensions");

        let reset = configured
            .set(
                ROW3,
                Props::new()
                    .with("height", PropValue::Null)
                    .with("hidden", PropValue::Null)
                    .with("outlineLevel", PropValue::Null)
                    .with("collapsed", PropValue::Null),
            )
            .expect("reset row dimensions")
            .set(
                COL_B,
                Props::new()
                    .with("width", PropValue::Null)
                    .with("hidden", PropValue::Null)
                    .with("outline", PropValue::Null)
                    .with("collapsed", PropValue::Null),
            )
            .expect("reset column dimensions");

        for key in ["height", "hidden", "outlineLevel", "collapsed"] {
            assert_eq!(format_property(&reset, ROW3, key), None, "row {key} reset");
        }
        for key in ["width", "hidden", "outlineLevel", "collapsed"] {
            assert_eq!(
                format_property(&reset, COL_B, key),
                None,
                "column {key} reset"
            );
        }
    }
}

#[test]
fn duplicate_dimension_aliases_are_invalid_arguments() {
    let bytes = fixture();
    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let row_error = original
            .set(
                ROW3,
                Props::new().with("outline", 1).with("outlineLevel", 2),
            )
            .expect_err("duplicate row outline aliases");
        assert_eq!(row_error.code(), Code::InvalidArgument);

        let column_error = original
            .set(COL_B, Props::new().with("group", 1).with("outline", 2))
            .expect_err("duplicate column outline aliases");
        assert_eq!(column_error.code(), Code::InvalidArgument);
    }
}

#[test]
fn dimension_values_outside_native_domains_are_invalid() {
    let bytes = fixture();
    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let original_bytes = original.to_bytes().expect("serialize original");

        for (path, properties, property) in [
            (ROW3, Props::new().with("height", "410pt"), "height"),
            (COL_B, Props::new().with("width", 256.0_f64), "width"),
            (ROW3, Props::new().with("outline", 8_i32), "outline"),
        ] {
            let error = original
                .set(path, properties)
                .expect_err("out-of-domain dimension value");
            assert_eq!(error.code(), Code::InvalidValue, "{property}");
            assert_eq!(error.property(), Some(property), "{property}");
            assert_eq!(
                original.to_bytes().expect("serialize after invalid value"),
                original_bytes,
                "invalid {property} must not publish a partial write"
            );
        }
    }
}

#[test]
fn dimension_staged_byte_limit_and_unknown_predicate_refuse_atomically() {
    let bytes = fixture();
    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let original_bytes = original.to_bytes().expect("serialize original");
        let staged_limits = SemanticLimits::builder()
            .max_staged_bytes(64)
            .expect("staged limit")
            .build()
            .expect("valid staged limit");
        let staged = original
            .edit_with_options(
                litchi_ops::EditOptions::new().semantic_limits(staged_limits),
                |transaction| {
                    transaction.set(ROW3, props! { "height" => "20pt" })?;
                    Ok::<_, litchi_ops::Error>(())
                },
            )
            .expect_err("dimension payload must exceed a 64-byte staged budget");
        assert_eq!(staged.code(), Code::LimitExceeded);

        let unknown = original
            .set(
                root_query("row[unknownDimension=true]"),
                props! { "height" => "20pt" },
            )
            .expect_err("unknown dynamic-table dimension predicate is not a match");
        assert_eq!(unknown.code(), Code::NotFound);

        let invalid = original
            .set(
                root_query("row[@unknownDimension=true]"),
                props! { "height" => "20pt" },
            )
            .expect_err("unknown explicit dimension predicate cannot widen a write");
        assert_eq!(invalid.code(), Code::InvalidSelector);
        assert_eq!(
            original.to_bytes().expect("serialize after refusals"),
            original_bytes
        );
    }
}
