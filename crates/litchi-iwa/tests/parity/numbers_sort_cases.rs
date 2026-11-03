//! Numbers table sort of wave W4-H (record 0421).
//!
//! - Persisted sort rules: the host's `table_sort_order`,
//!   `set_table_sort_order` and `clear_table_sort_order` delegate to the
//!   leaf, so the oracle is the host's hidden seams
//!   `__parity_table_sort_order_legacy` (the host's own codec read) and
//!   `__parity_set_table_sort_order_prost` (a Prost re-encoding of the table
//!   model) against `litchi_numbers::Package::{table_sort_order,
//!   edit_table_sort_order}`.
//! - Sort execution (**Sort Now**): the host's own executor
//!   (`Draft::apply_table_sort_order[_to_rows]`) against
//!   `litchi_numbers::Package::edit_table_sort_execution` over the shared
//!   `litchi_iwa_table::sort` executor.

use litchi_iwa::numbers;
use litchi_iwa_common::table::sort::{ColumnIndex, Direction, Order, RowRange, Rule, Scope};
use litchi_numbers::{Package, TableSelector};

use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier, fixture};
use crate::leaf_bytes;

/// Native workbooks whose first table both routes edit, plus a host-built
/// workbook.
const FIXTURES: [&str; 2] = [
    "iwork/numbers/basic.numbers",
    "iwork/numbers/custom-number-native.numbers",
];

const SHEET: usize = 0;
const TABLE: usize = 0;

fn built() -> TestResult<Fixture> {
    Ok(Fixture {
        path: "builder: host 6 × 4 workbook",
        bytes: numbers::Builder::new()
            .table_dimensions(6, 4)
            .build()?
            .to_bytes()?,
    })
}

fn inputs() -> TestResult<Vec<Fixture>> {
    let mut inputs = FIXTURES
        .iter()
        .map(|path| fixture(path))
        .collect::<TestResult<Vec<_>>>()?;
    inputs.push(built()?);
    Ok(inputs)
}

fn orders() -> TestResult<Vec<Order>> {
    Ok(vec![
        Order::new([Rule::new(ColumnIndex::new(0)?, Direction::Ascending)])?,
        Order::new([
            Rule::new(ColumnIndex::new(1)?, Direction::Descending),
            Rule::new(ColumnIndex::new(0)?, Direction::Ascending),
        ])?,
        Order::with_scope(
            Scope::SelectedRows,
            [Rule::new(ColumnIndex::new(1)?, Direction::Ascending)],
        )?,
    ])
}

fn host_read(bytes: &[u8]) -> TestResult<Option<Order>> {
    Ok(numbers::Draft::from_bytes(bytes)?
        .__parity_table_sort_order_legacy(TableSelector::index(TABLE))?)
}

fn leaf_read(bytes: &[u8]) -> TestResult<Option<Order>> {
    Ok(Package::from_bytes(bytes)?.table_sort_order(SHEET, TABLE)?)
}

fn host_write(source: &[u8], order: Option<&Order>) -> TestResult<Vec<u8>> {
    let mut editor = numbers::Draft::from_bytes(source)?;
    editor.__parity_set_table_sort_order_prost(TableSelector::index(TABLE), order)?;
    Ok(editor.to_bytes()?)
}

fn leaf_write(source: &[u8], order: Option<&Order>) -> TestResult<LeafRun> {
    let package = Package::from_bytes(source)?;
    let edit = package.edit_table_sort_order(SHEET, TABLE)?;
    let commit = match order {
        Some(order) => edit.set(order.clone()),
        None => edit.clear(),
    }
    .commit()?;
    let restored = commit
        .package()
        .apply_table_sort_order(&commit.patch().inverse())?;
    Ok(LeafRun::with_inverse(
        leaf_bytes!(commit.package()),
        leaf_bytes!(restored.package()),
    ))
}

fn write_case(
    name: &'static str,
    tier: Tier,
    evidence: Evidence,
    input: &Fixture,
    order: Option<&Order>,
) -> TestResult {
    Case::new(name, Format::Numbers, tier, evidence, input)
        .host(|source| host_write(source, order))
        .leaf(|source| leaf_write(source, order))
        .read("host table_sort_order", host_read)
        .read("leaf table_sort_order", leaf_read)
        .run()?;
    Ok(())
}

/// A rooted model's `sort_order` set on native and built workbooks:
/// one-rule and two-rule entire-table orders and a selected-rows order.
#[test]
fn numbers_table_sort_order_set_parity() -> TestResult {
    for input in inputs()? {
        for order in orders()? {
            write_case(
                "numbers_table_sort_order_set_parity",
                Tier::B,
                Evidence::Queued,
                &input,
                Some(&order),
            )?;
        }
    }
    Ok(())
}

/// Clearing a configured order keeps Numbers' empty-order marker on both
/// routes; clearing a table without an order is a no-op.
#[test]
fn numbers_table_sort_order_clear_parity() -> TestResult {
    for input in inputs()? {
        let sorted = Fixture {
            path: input.path,
            bytes: host_write(&input.bytes, Some(&orders()?[1]))?,
        };
        write_case(
            "numbers_table_sort_order_clear_parity",
            Tier::B,
            Evidence::Queued,
            &sorted,
            None,
        )?;
        assert_eq!(leaf_read(&leaf_write(&sorted.bytes, None)?.output)?, None);
    }
    Ok(())
}

/// The host's codec read and the leaf read agree on unsorted, sorted and
/// cleared tables.
#[test]
fn numbers_table_sort_order_read_parity() -> TestResult {
    for input in inputs()? {
        let mut sources = vec![Fixture {
            path: input.path,
            bytes: input.bytes.clone(),
        }];
        for order in orders()? {
            sources.push(Fixture {
                path: input.path,
                bytes: host_write(&input.bytes, Some(&order))?,
            });
        }
        sources.push(Fixture {
            path: input.path,
            bytes: host_write(&sources[1].bytes, None)?,
        });
        for source in &sources {
            Case::new(
                "numbers_table_sort_order_read_parity",
                Format::Numbers,
                Tier::A,
                Evidence::Queued,
                source,
            )
            .read_only()
            .host(|bytes| Ok(numbers::Draft::from_bytes(bytes)?.to_bytes()?))
            .leaf(|bytes| Ok(LeafRun::output(leaf_bytes!(Package::from_bytes(bytes)?))))
            .read("host table_sort_order", host_read)
            .read("leaf table_sort_order", leaf_read)
            .run()?;
            assert_eq!(host_read(&source.bytes)?, leaf_read(&source.bytes)?);
        }
    }
    Ok(())
}

/// The first table's body rows filled with keys — the first body column
/// numbers (with a duplicate, so the sort's stability shows), the second
/// texts — and a persisted order of `scope` on body column `column`.
fn sortable(
    input: &Fixture,
    scope: Scope,
    column: usize,
    direction: Direction,
) -> TestResult<Fixture> {
    use litchi_numbers::table::At;
    use litchi_numbers::table::cells::Input;
    const NUMBERS: [f64; 7] = [3.0, 7.5, -2.0, 42.0, 7.5, 0.25, 11.0];
    const TEXTS: [&str; 7] = ["kiwi", "pear", "apple", "fig", "date", "lime", "plum"];
    let rows = numbers::Draft::from_bytes(&input.bytes)?.tables()?[TABLE].rows;
    let package = Package::from_bytes(&input.bytes)?;
    let settings = package.table_header_settings(SHEET, TABLE)?;
    let first = settings.header_column_count();
    let body = settings.header_row_count()..rows - settings.footer_row_count();
    let mut edit = package.edit_table_cells(SHEET, TABLE)?;
    for row in body {
        let at = |column: usize| {
            At::new(
                u32::try_from(row).unwrap_or(u32::MAX),
                u32::try_from(first + column).unwrap_or(u32::MAX),
            )
        };
        edit = edit
            .set(
                at(0),
                Input::Number(litchi_numbers::cell::FiniteF64::new(
                    NUMBERS[row % NUMBERS.len()] + (row / NUMBERS.len()) as f64,
                )?),
            )?
            .set(at(1), Input::Text(TEXTS[row % TEXTS.len()].to_owned()))?;
    }
    let filled = edit.commit()?;
    let order = Order::with_scope(
        scope,
        [Rule::new(ColumnIndex::new(first + column)?, direction)],
    )?;
    let sorted = filled
        .package()
        .edit_table_sort_order(SHEET, TABLE)?
        .set(order)
        .commit()?;
    Ok(Fixture {
        path: input.path,
        bytes: leaf_bytes!(sorted.package()),
    })
}

fn host_execute(source: &[u8], rows: Option<RowRange>) -> TestResult<Vec<u8>> {
    let mut editor = numbers::Draft::from_bytes(source)?;
    let _moved = match rows {
        None => editor.apply_table_sort_order(TableSelector::index(TABLE))?,
        Some(rows) => editor.apply_table_sort_order_to_rows(TableSelector::index(TABLE), rows)?,
    };
    Ok(editor.to_bytes()?)
}

fn leaf_execute(source: &[u8], rows: Option<RowRange>) -> TestResult<LeafRun> {
    let package = Package::from_bytes(source)?;
    let mut edit = package.edit_table_sort_execution(SHEET, TABLE)?;
    match rows {
        None => drop(edit.apply()?),
        Some(rows) => drop(edit.apply_to_rows(rows)?),
    }
    let commit = edit.commit()?;
    let restored = commit
        .package()
        .apply_table_sort_execution(&commit.patch().inverse())?;
    Ok(LeafRun::with_inverse(
        leaf_bytes!(commit.package()),
        leaf_bytes!(restored.package()),
    ))
}

fn leaf_cells(bytes: &[u8]) -> TestResult<String> {
    use litchi_iwa_common::table::coordinate::{Area, At};
    let host = numbers::Draft::from_bytes(bytes)?;
    let table = &host.tables()?[TABLE];
    let end = At::new(
        u32::try_from(table.rows - 1)?,
        u32::try_from(table.columns - 1)?,
    );
    Ok(format!(
        "{:?}",
        Package::from_bytes(bytes)?.table_cells(SHEET, TABLE, Area::new(At::new(0, 0), end)?)?
    ))
}

fn execute_case(
    name: &'static str,
    tier: Tier,
    evidence: Evidence,
    input: &Fixture,
    rows: Option<RowRange>,
) -> TestResult {
    Case::new(name, Format::Numbers, tier, evidence, input)
        .host(|source| host_execute(source, rows))
        .leaf(|source| leaf_execute(source, rows))
        .read("leaf table_cells", leaf_cells)
        .read("host table_sort_order", host_read)
        .run()?;
    Ok(())
}

/// **Sort Now** of an entire-table order on native and built workbooks, by
/// a number column and a text column, with host-set cell borders and with
/// user-hidden rows.
#[test]
fn numbers_table_sort_apply_parity() -> TestResult {
    for input in inputs()? {
        for (column, direction) in [(0, Direction::Ascending), (1, Direction::Descending)] {
            execute_case(
                "numbers_table_sort_apply_parity",
                Tier::A,
                Evidence::Queued,
                &sortable(&input, Scope::EntireTable, column, direction)?,
                None,
            )?;
        }
    }
    // Explicit cell borders: a row layer moves with its row and a column
    // layer's runs split with the rows.
    let base = sortable(&built()?, Scope::EntireTable, 0, Direction::Ascending)?;
    let mut editor = numbers::Draft::from_bytes(&base.bytes)?;
    let model = crate::numbers_cases::host_first_table_model_id(&editor)?;
    let stroke = litchi_iwa_common::shape::stroke::Stroke::new(
        litchi_iwa_common::color::Rgba::new(
            0.2,
            0.4,
            0.6,
            1.0,
            litchi_iwa_common::color::RgbColorSpace::Srgb,
        )?,
        litchi_iwa_common::shape::stroke::Width::new(2.0)?,
        litchi_iwa_common::shape::stroke::Pattern::Solid,
    );
    for (row, column, side) in [
        (1, 0, litchi_iwa_common::table::cell::BorderSide::Top),
        (2, 1, litchi_iwa_common::table::cell::BorderSide::Left),
        (3, 1, litchi_iwa_common::table::cell::BorderSide::Left),
        (4, 2, litchi_iwa_common::table::cell::BorderSide::Bottom),
    ] {
        editor.set_table_cell_border(model, row, column, side, stroke)?;
    }
    execute_case(
        "numbers_table_sort_apply_parity",
        Tier::A,
        Evidence::Queued,
        &Fixture {
            path: "builder: host 6 × 4 workbook with keys, a sort order and host-set borders",
            bytes: editor.to_bytes()?,
        },
        None,
    )?;
    // User-hidden rows keep their positions.
    let base = sortable(&built()?, Scope::EntireTable, 0, Direction::Descending)?;
    let hidden = Package::from_bytes(&base.bytes)?
        .edit_table_hidden_axes(SHEET, TABLE)?
        .set(litchi_iwa_common::table::axis::HiddenAxes::new([
            litchi_iwa_common::table::axis::AxisIndex::row(2),
        ])?)
        .commit()?;
    execute_case(
        "numbers_table_sort_apply_parity",
        Tier::A,
        Evidence::Queued,
        &Fixture {
            path: "builder: host 6 × 4 workbook with keys, a sort order and a hidden row",
            bytes: leaf_bytes!(hidden.package()),
        },
        None,
    )?;
    Ok(())
}

/// **Sort Selected Rows** over a body-relative range.
#[test]
fn numbers_table_sort_apply_to_rows_parity() -> TestResult {
    for input in inputs()? {
        execute_case(
            "numbers_table_sort_apply_to_rows_parity",
            Tier::A,
            Evidence::Queued,
            &sortable(&input, Scope::SelectedRows, 0, Direction::Descending)?,
            Some(RowRange::new(1, 4)?),
        )?;
    }
    Ok(())
}

/// Both routes refuse what the executor does not move: a formula body
/// cell, a scope mismatch and a missing order; an already sorted table is
/// an exact no-op on the leaf.
#[test]
fn numbers_table_sort_refusals_match_host() -> TestResult {
    use litchi_numbers::table::sort::execution::{Error, Refusal, Unsupported};
    let input = fixture(FIXTURES[0])?;
    let entire = sortable(&input, Scope::EntireTable, 0, Direction::Ascending)?;
    // Scope mismatch.
    assert!(host_execute(&entire.bytes, Some(RowRange::new(0, 2)?)).is_err());
    let package = Package::from_bytes(&entire.bytes)?;
    let mut edit = package.edit_table_sort_execution(SHEET, TABLE)?;
    edit.apply_to_rows(RowRange::new(0, 2)?)?;
    assert!(matches!(
        edit.commit(),
        Err(Error::Refused(Refusal::ScopeMismatch))
    ));
    // Missing order.
    assert!(host_execute(&input.bytes, None).is_err());
    let package = Package::from_bytes(&input.bytes)?;
    let mut edit = package.edit_table_sort_execution(SHEET, TABLE)?;
    edit.apply()?;
    assert!(matches!(
        edit.commit(),
        Err(Error::Refused(Refusal::NoSortOrder))
    ));
    // A formula body cell (written by the host formula author).
    let mut editor = numbers::Draft::from_bytes(&entire.bytes)?;
    let model = crate::numbers_cases::host_first_table_model_id(&editor)?;
    editor.set_formula(model, 2, 2, numbers::FormulaExpression::Number(1.0))?;
    let formula = editor.to_bytes()?;
    assert!(host_execute(&formula, None).is_err());
    let package = Package::from_bytes(&formula)?;
    let mut edit = package.edit_table_sort_execution(SHEET, TABLE)?;
    edit.apply()?;
    let refused = edit.commit();
    assert!(
        matches!(
            refused,
            Err(Error::Refused(Refusal::Unsupported {
                reason: Unsupported::Formula
            }))
        ),
        "{refused:?}"
    );
    // Already in order: an exact no-op.
    let sorted = leaf_execute(&entire.bytes, None)?.output;
    let package = Package::from_bytes(&sorted)?;
    let mut edit = package.edit_table_sort_execution(SHEET, TABLE)?;
    edit.apply()?;
    let commit = edit.commit()?;
    assert!(!commit.diagnostics().changed());
    assert!(commit.patch().is_noop());
    Ok(())
}

/// Callgrind probe (record 0421): **Sort Now** of the built table by its
/// number column, host executor against leaf, ABBA.
#[test]
#[ignore = "callgrind probe; run explicitly"]
fn numbers_sort_apply_cost_abba() -> TestResult {
    let input = sortable(&built()?, Scope::EntireTable, 0, Direction::Ascending)?;
    crate::pages_cost_probes::abba(
        &input,
        |source| host_execute(source, None),
        |source| {
            let package = Package::from_bytes(source)?;
            let mut edit = package.edit_table_sort_execution(SHEET, TABLE)?;
            edit.apply()?;
            Ok(leaf_bytes!(edit.commit()?.package()))
        },
    )
}

/// Callgrind probe (record 0421): set a two-rule order on the first table of
/// `custom-number-native.numbers`, host oracle against leaf, ABBA.
#[test]
#[ignore = "callgrind probe; run explicitly"]
fn numbers_sort_order_cost_abba() -> TestResult {
    let input = fixture(FIXTURES[1])?;
    let order = orders()?[1].clone();
    crate::pages_cost_probes::abba(
        &input,
        |source| host_write(source, Some(&order)),
        |source| {
            let package = Package::from_bytes(source)?;
            let commit = package
                .edit_table_sort_order(SHEET, TABLE)?
                .set(order.clone())
                .commit()?;
            Ok(leaf_bytes!(commit.package()))
        },
    )
}

/// Callgrind probe (record 0421): the same set, with the host side running
/// the host's own `set_table_sort_order` route step by step (that method is
/// test-only in the host: serialize the draft, open the focused package,
/// commit, write, reopen the draft and read the order back), ABBA.
#[test]
#[ignore = "callgrind probe; run explicitly"]
fn numbers_sort_order_host_route_cost_abba() -> TestResult {
    let input = fixture(FIXTURES[1])?;
    let order = orders()?[1].clone();
    let draft = numbers::Draft::from_bytes(&input.bytes)?;
    crate::pages_cost_probes::abba(
        &input,
        |_source| {
            let bytes = draft.to_bytes()?;
            let commit = Package::from_bytes(&bytes)?
                .edit_table_sort_order(SHEET, TABLE)?
                .set(order.clone())
                .commit()?;
            let written = leaf_bytes!(commit.package());
            let verified = numbers::Draft::from_bytes(&written)?;
            if verified.__parity_table_sort_order_legacy(TableSelector::index(TABLE))?
                != Some(order.clone())
            {
                return Err("round trip".into());
            }
            Ok(verified.to_bytes()?)
        },
        |source| {
            let package = Package::from_bytes(source)?;
            let commit = package
                .edit_table_sort_order(SHEET, TABLE)?
                .set(order.clone())
                .commit()?;
            Ok(leaf_bytes!(commit.package()))
        },
    )
}
