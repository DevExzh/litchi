//! Regression coverage for the neutral table model and the Numbers facade.
//!
//! These tests exercise the public semantic boundary from both sides.  The
//! common model must remain usable without a Numbers dependency, while the
//! Numbers names and conversions must continue to expose the same lossless
//! sparse-table behavior.

use std::mem::size_of;

use litchi_iwa_common::table::cell::value as shared_value;
use litchi_iwa_common::table::model as shared_model;
use litchi_numbers::cell::{FiniteF64 as NumbersFiniteF64, Value as NumbersValue};
use litchi_numbers::{
    CellPosition as NumbersCellPosition, CellRange as NumbersCellRange,
    Dimensions as NumbersDimensions, Sheet, Table as NumbersTable, TableBuilder, TableSelector,
};

fn finite(value: f64) -> shared_value::FiniteF64 {
    shared_value::FiniteF64::new(value).expect("test scalar should be finite")
}

fn assert_shared_position(_: litchi_iwa_common::table::coordinate::CellPosition) {}

fn assert_shared_range(_: litchi_iwa_common::table::coordinate::CellRange) {}

fn assert_shared_dimensions(_: litchi_iwa_common::table::model::Dimensions) {}

fn assert_shared_cell(_: litchi_iwa_common::table::model::Cell) {}

fn assert_shared_view(_: litchi_iwa_common::table::model::View<'_>) {}

fn shared_table() -> shared_model::Table {
    let mut builder = shared_model::Table::builder("Revenue", NumbersDimensions::new(3, 3));
    builder
        .set_column_headers(["Item", "Amount", "State"])
        .expect("column headers should fit");
    builder
        .set_row_headers(["first", "second", "third"])
        .expect("row headers should fit");
    builder
        .set(
            NumbersCellPosition::new(1, 1),
            shared_value::Value::number(42.5).expect("finite number should construct"),
        )
        .expect("in-bounds cell should insert");
    builder
        .set(
            NumbersCellPosition::new(0, 0),
            shared_value::Value::Text("Coffee".to_owned()),
        )
        .expect("in-bounds cell should insert");
    builder
        .set(
            NumbersCellPosition::new(2, 2),
            shared_value::Value::Boolean(true),
        )
        .expect("in-bounds cell should insert");
    builder.finish().expect("test sparse table should finish")
}

#[test]
fn all_value_variants_keep_shared_type_identity_and_semantics() {
    let values = [
        shared_value::Value::Empty,
        shared_value::Value::Text("text".to_owned()),
        shared_value::Value::Number(finite(12.5)),
        shared_value::Value::Boolean(true),
        shared_value::Value::Date(finite(123.0)),
        shared_value::Value::Duration(finite(45.0)),
        shared_value::Value::Formula("A1+B1".to_owned()),
        shared_value::Value::Error("#VALUE!".to_owned()),
    ];
    let expected_types = [
        shared_value::Type::Empty,
        shared_value::Type::Text,
        shared_value::Type::Number,
        shared_value::Type::Boolean,
        shared_value::Type::Date,
        shared_value::Type::Duration,
        shared_value::Type::Formula,
        shared_value::Type::Error,
    ];

    for (value, expected_type) in values.into_iter().zip(expected_types) {
        // These assignments are compile-time identity checks: Numbers keeps
        // its established public name as a re-export of the shared value.
        let numbers_value: NumbersValue = value;
        let shared_again: shared_value::Value = numbers_value;
        assert_eq!(shared_again.cell_type(), expected_type);
    }

    let shared_finite = finite(3.5);
    let numbers_finite: NumbersFiniteF64 = shared_finite;
    let _: shared_value::FiniteF64 = numbers_finite;
    assert_eq!(numbers_finite.get(), 3.5);
    assert_eq!(size_of::<NumbersValue>(), size_of::<shared_value::Value>());
}

#[test]
fn coordinates_and_ranges_keep_shared_type_identity_and_a1_behavior() {
    let shared_dimensions = NumbersDimensions::new(12, 54);
    let numbers_dimensions: NumbersDimensions = shared_dimensions;
    assert_shared_dimensions(numbers_dimensions);

    let shared_position =
        NumbersCellPosition::from_a1("$bc$12").expect("absolute A1 coordinate should parse");
    let numbers_position: NumbersCellPosition = shared_position;
    assert_shared_position(numbers_position);
    assert_eq!(numbers_position, NumbersCellPosition::new(11, 54));
    assert_eq!(numbers_position.to_string(), "BC12");

    let shared_range = NumbersCellRange::from_a1("B3:D5").expect("inclusive A1 range should parse");
    let numbers_range: NumbersCellRange = shared_range;
    assert_shared_range(numbers_range);
    assert_eq!(numbers_range.start(), NumbersCellPosition::new(2, 1));
    assert_eq!(numbers_range.end(), NumbersCellPosition::new(5, 4));
    assert_eq!(numbers_range.area(), Some(9));
    assert!(numbers_range.contains(NumbersCellPosition::new(4, 3)));
    assert!(!numbers_range.contains(NumbersCellPosition::new(5, 3)));
}

#[test]
fn sparse_model_distinguishes_missing_from_stored_empty() {
    let mut builder = shared_model::Table::builder("Sparse", NumbersDimensions::new(1, 2));
    builder
        .set(NumbersCellPosition::new(0, 0), shared_value::Value::Empty)
        .expect("explicit empty should be materialized");
    let table = builder.finish().expect("sparse table should finish");

    let shared_cell = shared_model::Cell::new(
        NumbersCellPosition::new(0, 1),
        shared_value::Value::Text("alias".to_owned()),
    );
    let numbers_cell: litchi_numbers::Cell = shared_cell;
    assert_shared_cell(numbers_cell);

    assert!(matches!(
        table.view(NumbersCellPosition::new(0, 0)),
        shared_model::View::Stored(value) if value.is_empty()
    ));
    let shared_view = table.view(NumbersCellPosition::new(0, 1));
    assert_shared_view(shared_view);
    assert!(matches!(shared_view, shared_model::View::Missing));
    assert_eq!(table.cell_count(), 1);
    assert_eq!(table.non_empty_cell_count(), 0);
}

#[test]
fn sparse_builder_orders_cells_and_rejects_duplicate_pushes() {
    let mut builder = shared_model::Table::builder("Ordered", NumbersDimensions::new(3, 3));
    builder
        .push(shared_model::Cell::new(
            NumbersCellPosition::new(2, 1),
            shared_value::Value::Text("late".to_owned()),
        ))
        .expect("first pushed cell should fit");
    builder
        .push(shared_model::Cell::new(
            NumbersCellPosition::new(0, 2),
            shared_value::Value::Text("early".to_owned()),
        ))
        .expect("second pushed cell should fit");
    let table = builder.finish().expect("out-of-order cells should sort");
    assert_eq!(
        table
            .iter_cells()
            .map(shared_model::Cell::position)
            .collect::<Vec<_>>(),
        [
            NumbersCellPosition::new(0, 2),
            NumbersCellPosition::new(2, 1)
        ]
    );

    let duplicate_position = NumbersCellPosition::new(0, 0);
    let mut duplicate = shared_model::Table::builder("Duplicate", NumbersDimensions::new(1, 1));
    duplicate
        .push(shared_model::Cell::new(
            duplicate_position,
            shared_value::Value::Empty,
        ))
        .expect("first duplicate candidate should fit");
    duplicate
        .push(shared_model::Cell::new(
            duplicate_position,
            shared_value::Value::Text("replacement".to_owned()),
        ))
        .expect("duplicate detection belongs to finish");
    assert!(matches!(
        duplicate.finish(),
        Err(shared_model::Error::DuplicatePosition { position }) if position == duplicate_position
    ));
}

#[test]
fn sparse_bounds_and_grid_budget_fail_before_partial_results() {
    let dimensions = NumbersDimensions::new(2, 2);
    let mut builder = shared_model::Table::builder("Bounds", dimensions);
    let value = shared_value::Value::Text("rejected".to_owned());
    let rejected = builder
        .set(NumbersCellPosition::new(2, 0), value.clone())
        .expect_err("row two is outside a two-row table");
    assert!(matches!(
        rejected.error(),
        shared_model::Error::OutOfBounds { position, dimensions: actual }
            if *position == NumbersCellPosition::new(2, 0) && *actual == dimensions
    ));
    let (error, returned) = rejected.into_parts();
    assert!(matches!(error, shared_model::Error::OutOfBounds { .. }));
    assert_eq!(returned, value);

    let table = builder.finish().expect("empty bounded table should finish");
    let full_range = NumbersCellRange::new(
        NumbersCellPosition::new(0, 0),
        NumbersCellPosition::new(2, 2),
    )
    .expect("full range should be valid");
    assert!(matches!(
        table.grid(full_range, shared_model::GridBudget::new(3)),
        Err(shared_model::Error::BudgetExceeded {
            requested: 4,
            maximum: 3
        })
    ));
    assert_eq!(
        table
            .grid(full_range, shared_model::GridBudget::new(4))
            .expect("inclusive budget should admit full range")
            .iter()
            .count(),
        4
    );
    let grid = table
        .grid(full_range, shared_model::GridBudget::new(4))
        .expect("grid should remain available through the shared model");
    let _: litchi_numbers::Grid<'_> = grid;
    let outside = NumbersCellRange::new(
        NumbersCellPosition::new(0, 0),
        NumbersCellPosition::new(3, 2),
    )
    .expect("syntactically valid outside range should construct");
    assert!(matches!(
        table.grid(outside, shared_model::GridBudget::new(usize::MAX)),
        Err(shared_model::Error::OutOfBounds { .. })
    ));
}

#[test]
fn numbers_facade_preserves_selector_headers_csv_and_lossless_core_conversion() {
    let shared = shared_table();
    let numbers = NumbersTable::from_shared(shared.clone());

    assert_eq!(numbers.name(), "Revenue");
    assert_eq!(numbers.as_shared(), &shared);
    assert_eq!(numbers.selector(), TableSelector::name("Revenue"));
    assert_eq!(numbers.dimensions(), NumbersDimensions::new(3, 3));
    assert_eq!(
        numbers.column_headers().collect::<Vec<_>>(),
        ["Item", "Amount", "State"]
    );
    assert_eq!(
        numbers.row_headers().collect::<Vec<_>>(),
        ["first", "second", "third"]
    );
    assert_eq!(
        numbers.get_a1("B2"),
        Ok(Some(&shared_value::Value::Number(finite(42.5))))
    );
    assert_eq!(
        numbers.get_a1("C3"),
        Ok(Some(&shared_value::Value::Boolean(true)))
    );
    assert_eq!(numbers.get_a1("C1"), Ok(None));
    let text_pointer_before_move = match numbers
        .get(NumbersCellPosition::new(0, 0))
        .expect("the fixture should contain the first text cell")
    {
        shared_value::Value::Text(value) => value.as_ptr(),
        value => panic!("expected a text cell, got {:?}", value.cell_type()),
    };
    assert_eq!(
        numbers.to_csv(),
        "Item,Amount,State\nfirst,Coffee,,\nsecond,,42.5,\nthird,,,true\n"
    );

    let converted = numbers.into_shared();
    let text_pointer_after_move = match converted
        .get(NumbersCellPosition::new(0, 0))
        .expect("the converted table should retain the first text cell")
    {
        shared_value::Value::Text(value) => value.as_ptr(),
        value => panic!("expected a text cell, got {:?}", value.cell_type()),
    };
    assert_eq!(
        text_pointer_before_move, text_pointer_after_move,
        "Table::from_shared/into_shared must preserve owned text storage"
    );
    assert_eq!(converted, shared);

    let via_from: NumbersTable = shared.clone().into();
    let via_into: shared_model::Table = via_from.into();
    assert_eq!(via_into, shared);

    let mut shared_builder = shared_model::Table::builder("Builder", NumbersDimensions::new(1, 1));
    shared_builder
        .set(
            NumbersCellPosition::new(0, 0),
            shared_value::Value::Text("builder value".to_owned()),
        )
        .expect("shared builder cell should fit");
    let numbers_builder = TableBuilder::from_shared(shared_builder);
    let shared_table_from_builder = numbers_builder
        .into_shared()
        .finish()
        .expect("shared builder conversion should remain lossless");
    assert_eq!(
        shared_table_from_builder.get(NumbersCellPosition::new(0, 0)),
        Some(&shared_value::Value::Text("builder value".to_owned()))
    );

    let rebuilt = TableBuilder::new("Revenue", NumbersDimensions::new(3, 3));
    let rebuilt = rebuilt
        .finish()
        .expect("Numbers builder compatibility should remain available");
    assert_eq!(rebuilt.into_shared().name(), "Revenue");
}

#[test]
fn numbers_facade_keeps_duplicate_name_and_coordinate_error_types() {
    let first = TableBuilder::new("Same", NumbersDimensions::new(1, 1))
        .finish()
        .expect("first table should be valid");
    let second = TableBuilder::new("Same", NumbersDimensions::new(1, 1))
        .finish()
        .expect("second table should be valid");
    let duplicate = Sheet::try_from_tables("Sheet", 0, vec![first, second]);
    assert!(matches!(
        duplicate,
        Err(shared_model::Error::DuplicateTableName { name }) if name == "Same"
    ));

    let position = NumbersCellPosition::new(4, 4);
    let mut builder = TableBuilder::new("Bounded", NumbersDimensions::new(2, 2));
    let rejected = builder
        .set(position, NumbersValue::Empty)
        .expect_err("Numbers facade should preserve out-of-bounds insertion errors");
    assert!(matches!(
        rejected.error(),
        shared_model::Error::OutOfBounds { position: actual, .. } if *actual == position
    ));
}
