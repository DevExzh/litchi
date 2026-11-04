//! Native pivot table authored in Numbers 14.4 (record 0432).
//!
//! `pivot-all-summary-functions-native.numbers`: sheet "Table 1 Pivot" pivots
//! "Table 1" (rows Region, columns Product) with the values Amount summarized
//! by 13 functions, in the order [`VALUES`] lists them.

use std::io;
use std::path::PathBuf;

use litchi_iwa_archive::package::Catalog;
use litchi_iwa_core::{Archive, SnappyStream};
use litchi_iwa_protos::tst;

use litchi_numbers::cell::Value;
use litchi_numbers::table::At;

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
