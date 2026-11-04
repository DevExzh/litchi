//! Native pivot table authored in Numbers 14.4 (record 0432).
//!
//! `pivot-all-summary-functions-native.numbers`: sheet "Table 1 Pivot" pivots
//! "Table 1" (rows Region, columns Product) with the values Amount summarized
//! by 13 functions, in the order [`VALUES`] lists them.

use std::io;
use std::path::PathBuf;

use litchi_iwa_archive::package::Catalog;
use litchi_iwa_core::{Archive, SnappyStream};
use litchi_iwa_protos::{tsce, tst};

use litchi_numbers::cell::Value;
use litchi_numbers::formula::{BinaryOperator, CachedValue, Expression};
use litchi_numbers::table::At;
use litchi_numbers::table::cells::{DependencyKind, Error as CellError, Input, Refusal, Storage};

use litchi_numbers::Package;
use litchi_numbers::table::pivot::{Aggregate, Category};
use litchi_numbers_wire::{BncCell, CachedScalar, ListKind};
use prost::Message as _;

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

const PIVOT_SHEET: &str = "Table 1 Pivot";
const PIVOT_TABLE: &str = "Table 1 Pivot";
const PIVOT_TILE: &str = "Index/Tables/Tile-909115.iwa";
const FORMULA_LIST: &str = "Index/Tables/DataList-909120.iwa";

/// The Values order, the header Numbers shows for each value column (a
/// function the owner verified in the "Summarize by" menu), the native code,
/// and the cached result on the Widget/North row, whose amounts are 10 and 70.
const VALUES: [(Aggregate, &str, u32, f64); 13] = [
    (Aggregate::Sum, "Amount (Sum)", 2, 80.0),
    (Aggregate::Average, "Amount (Average)", 18, 40.0),
    (Aggregate::CountAll, "Amount (Count All)", 1, 2.0),
    (Aggregate::Maximum, "Amount (Max.)", 5, 70.0),
    (Aggregate::Minimum, "Amount (Min.)", 4, 10.0),
    (Aggregate::Count, "Amount (Count)", 13, 2.0),
    (Aggregate::CountUnique, "Amount (Unique)", 14, 2.0),
    (Aggregate::Median, "Amount (Median)", 19, 40.0),
    (Aggregate::Product, "Amount (Product)", 21, 700.0),
    (
        Aggregate::StandardDeviation,
        "Amount (STDEV)",
        26,
        42.426_406_871_192_85,
    ),
    (
        Aggregate::StandardDeviationPopulation,
        "Amount (STDEVP)",
        28,
        30.0,
    ),
    (Aggregate::Variance, "Amount (VAR)", 30, 1_800.0),
    (Aggregate::VariancePopulation, "Amount (VARP)", 32, 900.0),
];

fn fixture() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/numbers/pivot-all-summary-functions-native.numbers")
}

fn member(source: &[u8], name: &str) -> TestResult<Archive> {
    let catalog = Catalog::from_bytes(source, Default::default())?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == name)
        .ok_or_else(|| io::Error::other(format!("no member {name}")))?;
    let stream = SnappyStream::decompress(entry.data(), Default::default())?;
    Ok(Archive::parse(stream.as_bytes(), Default::default())?)
}

/// The BNC cell of the pivot tile at `(row, column)`.
fn pivot_cell(source: &[u8], row: u32, column: usize) -> TestResult<BncCell> {
    let archive = member(source, PIVOT_TILE)?;
    let tile = tst::Tile::decode(archive.objects[0].messages[0].data.as_slice())?;
    let row = tile
        .row_infos
        .iter()
        .find(|info| info.tile_row_index == row)
        .ok_or_else(|| io::Error::other("no tile row"))?;
    let offsets = row.cell_offsets.as_deref().unwrap_or_default();
    let slot = |index: usize| -> Option<usize> {
        let raw = offsets.get(index * 2..index * 2 + 2)?;
        let raw = u16::from_le_bytes([raw[0], raw[1]]);
        (raw != u16::MAX).then_some(usize::from(raw))
    };
    let storage = row.cell_storage_buffer.as_deref().unwrap_or_default();
    let start = slot(column).ok_or_else(|| io::Error::other("no cell"))?;
    let end = (column + 1..255).find_map(slot).unwrap_or(storage.len());
    Ok(BncCell::parse(
        &storage[start..end],
        &litchi_numbers_wire::Limits::DEFAULT,
    )?)
}

/// `(formula list key, summary-function code)` of every native pivot formula.
fn formula_codes(source: &[u8]) -> TestResult<Vec<(u32, u32)>> {
    let archive = member(source, FORMULA_LIST)?;
    let list = tst::TableDataList::decode(archive.objects[0].messages[0].data.as_slice())?;
    let mut codes = Vec::new();
    for entry in &list.entries {
        let formula = entry
            .formula
            .as_ref()
            .ok_or_else(|| io::Error::other("no formula"))?;
        let node = &formula.ast_node_array.ast_node[0];
        let category = node
            .ast_category_ref
            .as_ref()
            .ok_or_else(|| io::Error::other("the pivot body cell is not a category"))?;
        codes.push((entry.key, category.category_ref.aggregate_type));
    }
    Ok(codes)
}

fn categories_of(package: &Package) -> TestResult<Vec<Category>> {
    Ok(package.pivot_categories()?)
}

/// The summary-function names read back the 13 functions, in Values order, for
/// every group of both enabled group-bys.
#[test]
fn native_pivot_names_all_thirteen_summary_functions() -> TestResult {
    let package = Package::open(fixture())?;
    let categories = categories_of(&package)?;
    let expected = VALUES.map(|(aggregate, ..)| aggregate);
    // 3 regions under 2 products, 2 products and the root in the first
    // group-by; 3 regions and the root in the second.
    assert_eq!(categories.len(), (6 + 2 + 1 + 3 + 1) * 13);
    for group in categories.chunks(13) {
        let label = group[0].label();
        let aggregates = group
            .iter()
            .map(|category| {
                assert_eq!(category.label(), label);
                category.aggregate()
            })
            .collect::<Vec<_>>();
        assert_eq!(aggregates, expected);
    }
    let labels = categories
        .chunks(13)
        .map(|group| (group[0].label().map(str::to_owned), group[0].group_level()))
        .filter(|(label, _)| label.as_deref() == Some("North"))
        .map(|(_, level)| level)
        .collect::<Vec<_>>();
    assert_eq!(labels, [2, 2, 1]);
    Ok(())
}

/// Each value column's header, the code of its native formula and its cached
/// result agree with the name: the Values order is not the only evidence.
#[test]
fn native_pivot_codes_agree_with_headers_formulas_and_cached_results() -> TestResult {
    let package = Package::open(fixture())?;
    let source = std::fs::read(fixture())?;
    let formulas = formula_codes(&source)?;
    let table = package
        .sheets()
        .iter()
        .find(|sheet| sheet.name() == PIVOT_SHEET)
        .and_then(|sheet| sheet.tables().find(|table| table.name() == PIVOT_TABLE))
        .ok_or_else(|| io::Error::other("no pivot table"))?;
    assert_eq!(
        table.get(At::new(0, 1)),
        Some(&Value::Text("Widget".into()))
    );
    assert_eq!(table.get(At::new(2, 0)), Some(&Value::Text("North".into())));
    for (index, (_, header, code, cached)) in VALUES.iter().enumerate() {
        let column = u32::try_from(index + 1)?;
        assert_eq!(
            table.get(At::new(1, column)),
            Some(&Value::Text((*header).into()))
        );
        let cell = pivot_cell(&source, 2, index + 1)?;
        let key = cell
            .list_keys()
            .find_map(|(kind, key)| (kind == ListKind::Formula).then_some(key))
            .ok_or_else(|| io::Error::other("no formula key"))?;
        let found = formulas
            .iter()
            .find(|(entry, _)| *entry == key)
            .map(|(_, code)| *code);
        assert_eq!(found, Some(*code), "{header}");
        let Some(CachedScalar::Number(value)) = cell.cached_scalar()? else {
            return Err(io::Error::other("no cached number").into());
        };
        let value = f64::try_from(value)?;
        assert!((value - cached).abs() < 1e-9, "{header}: {value}");
    }
    Ok(())
}

fn exact_bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

fn engine(source: &[u8]) -> TestResult<Archive> {
    member(source, "Index/CalculationEngine.iwa")
}

/// Every message of `message_type` in the calculation engine, by object.
fn engine_messages(source: &[u8], message_type: u32) -> TestResult<Vec<(u64, Vec<u8>)>> {
    Ok(engine(source)?
        .objects
        .iter()
        .flat_map(|object| {
            let identifier = object.archive_info.identifier.unwrap_or(0);
            object
                .messages
                .iter()
                .filter(move |message| message.type_ == message_type)
                .map(move |message| (identifier, message.data.clone()))
        })
        .collect())
}

/// `(precedent row, column, owner)` of a cell's dependency record in the
/// record tiles of `owner`.
fn dependency_edges(
    source: &[u8],
    owner: u32,
    (row, column): (u32, u32),
) -> TestResult<Vec<(u32, u32, u32)>> {
    for (_, data) in engine_messages(source, 4_009)? {
        let tile = tsce::CellRecordTileArchive::decode(data.as_slice())?;
        if tile.internal_owner_id != owner {
            continue;
        }
        for record in &tile.cell_records {
            if (record.row, record.column) == (row, column) {
                let edges = record
                    .expanded_edges
                    .as_ref()
                    .ok_or_else(|| io::Error::other("no expanded edges"))?;
                return Ok(edges
                    .edge_with_owner_rows
                    .iter()
                    .zip(&edges.edge_with_owner_columns)
                    .zip(&edges.internal_owner_id_for_edge)
                    .map(|((row, column), owner)| (*row, *column, *owner))
                    .collect());
            }
        }
    }
    Err(io::Error::other("no dependency record").into())
}

/// The category reference of the pivot cell's formula.
fn cell_category(
    source: &[u8],
    row: u32,
    column: usize,
) -> TestResult<tsce::CategoryReferenceArchive> {
    let cell = pivot_cell(source, row, column)?;
    let key = cell
        .list_keys()
        .find_map(|(kind, key)| (kind == ListKind::Formula).then_some(key))
        .ok_or_else(|| io::Error::other("no formula key"))?;
    let archive = member(source, FORMULA_LIST)?;
    let list = tst::TableDataList::decode(archive.objects[0].messages[0].data.as_slice())?;
    let entry = list
        .entries
        .iter()
        .find(|entry| entry.key == key)
        .ok_or_else(|| io::Error::other("no formula entry"))?;
    let formula = entry
        .formula
        .as_ref()
        .ok_or_else(|| io::Error::other("no formula"))?;
    let node = &formula.ast_node_array.ast_node[0];
    Ok(node
        .ast_category_ref
        .as_ref()
        .ok_or_else(|| io::Error::other("no category node"))?
        .category_ref
        .clone())
}

/// The engine's owner map holds entries with no owner object (7 for a pivot's
/// snapshot table); the old equal-count check refused every cell edit in the
/// workbook, on any table.
#[test]
fn native_pivot_source_table_edit_publishes_and_inverts() -> TestResult {
    let package = Package::open(fixture())?;
    let source = exact_bytes(&package)?;
    let commit = package
        .edit_table_cells("Sheet 1", "Table 1")?
        .set(At::new(1, 2), Input::number(55.0)?)?
        .commit()?;
    let edited = commit.package();
    let cell = edited.table_cell("Sheet 1", "Table 1", At::new(1, 2))?;
    assert!(matches!(
        cell.storage(),
        Storage::Stored(Value::Number(value)) if value.get() == 55.0
    ));
    // The pivot's state is Numbers' to refresh: nothing of it moves.
    let bytes = exact_bytes(edited)?;
    for message in [6_370, 6_372, 6_373, 6_382, 6_383] {
        assert_eq!(
            engine_messages(&source, message)?,
            engine_messages(&bytes, message)?,
            "message {message}"
        );
    }
    assert_eq!(cell_category(&bytes, 2, 1)?, cell_category(&source, 2, 1)?);
    assert_eq!(edited.pivot_categories()?, package.pivot_categories()?);
    let restored = edited.apply_table_cells(&commit.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, source);
    Ok(())
}

/// A category formula the leaf writes over a native body cell is the one
/// Numbers wrote there: the same reference and the same dependency edges.
#[test]
fn native_pivot_category_formula_matches_what_numbers_wrote() -> TestResult {
    let package = Package::open(fixture())?;
    let source = exact_bytes(&package)?;
    let native = cell_category(&source, 4, 1)?;
    let native_edges = dependency_edges(&source, 37, (4, 1))?;
    let mut matched = 0;
    for category in package.pivot_categories()?.iter().filter(|category| {
        category.label() == Some("West")
            && category.group_level() == 2
            && category.aggregate() == Aggregate::Sum
    }) {
        let bytes = exact_bytes(
            package
                .edit_table_cells(PIVOT_SHEET, PIVOT_TABLE)?
                .set_formula_cached(
                    At::new(4, 1),
                    Expression::pivot_category(category),
                    CachedValue::number(1.0)?,
                )?
                .commit()?
                .package(),
        )?;
        let written = cell_category(&bytes, 4, 1)?;
        if written == native {
            assert_eq!(dependency_edges(&bytes, 37, (4, 1))?, native_edges);
            matched += 1;
        }
    }
    assert_eq!(matched, 1);
    Ok(())
}

/// Each of the 13 functions of one group can be referenced from the source
/// table, and reopens with the function it was authored with.
#[test]
fn native_pivot_category_formula_on_source_table_keeps_each_function() -> TestResult {
    let package = Package::open(fixture())?;
    let categories = package.pivot_categories()?;
    for category in categories.iter().take(13) {
        let commit = package
            .edit_table_cells("Sheet 1", "Table 1")?
            .set_formula_cached(
                At::new(1, 2),
                Expression::binary(
                    BinaryOperator::Divide,
                    Expression::pivot_category(category),
                    Expression::number(2.0)?,
                )?,
                CachedValue::number(1.0)?,
            )?
            .commit()?;
        let reopened = Package::from_bytes(&exact_bytes(commit.package())?)?;
        let cell = reopened.table_cell("Sheet 1", "Table 1", At::new(1, 2))?;
        assert!(matches!(cell.storage(), Storage::Stored(Value::Formula(_))));
        let restored = commit
            .package()
            .apply_table_cells(&commit.patch().inverse())?;
        assert_eq!(exact_bytes(restored.package())?, exact_bytes(&package)?);
    }
    Ok(())
}

/// A scalar over a pivot body cell, which holds a category formula. The host
/// has no scalar setter; the refusal is typed and nothing is published.
#[test]
fn native_pivot_body_scalar_is_refused_without_publishing() -> TestResult {
    let package = Package::open(fixture())?;
    let before = exact_bytes(&package)?;
    let result = package
        .edit_table_cells(PIVOT_SHEET, PIVOT_TABLE)?
        .set(At::new(2, 1), Input::number(55.0)?)?
        .commit();
    assert!(matches!(
        result,
        Err(CellError::Refused(Refusal::UnsupportedDependency {
            kind: DependencyKind::FormulaCache,
            ..
        }))
    ));
    assert_eq!(exact_bytes(&package)?, before);
    Ok(())
}
