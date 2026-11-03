//! Numbers table sort of wave W4-H (record 0421).
//!
//! - Persisted sort rules: the host's `table_sort_order`,
//!   `set_table_sort_order` and `clear_table_sort_order` delegate to the
//!   leaf, so the oracle is the host's hidden seams
//!   `__parity_table_sort_order_legacy` (the host's own codec read) and
//!   `__parity_set_table_sort_order_prost` (a Prost re-encoding of the table
//!   model) against `litchi_numbers::Package::{table_sort_order,
//!   edit_table_sort_order}`.

use litchi_iwa::numbers;
use litchi_iwa_common::table::sort::{ColumnIndex, Direction, Order, Rule, Scope};
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
