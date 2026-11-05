//! Numbers pivot categories and category formulas on one synthetic inline
//! group-by added to `basic.numbers`' calculation engine (string, number,
//! Boolean and date group values, a grand-total root, two aggregated columns),
//! with the values the deleted `litchi-iwa` host agreed with at its last commit,
//! `4d8a3e161f` (records 0421, 0424, 0429).
//!
//! The object is not native data, and Numbers drops it because no owner lists
//! it (record 0434 §6): these tests pin the leaf's answer only.

use litchi_iwa_archive::iwa::{Archive, ArchiveObject, RawMessage, SnappyStream};
use litchi_iwa_archive::package::{Catalog, EntryEdit};
use litchi_iwa_protos::{tsce, tsk, tsp, tst};
use litchi_numbers::Package;
use litchi_numbers::cell::Value;
use litchi_numbers::formula::{BinaryOperator, CachedValue, Expression};
use litchi_numbers::table::At;
use litchi_numbers::table::cells::{Error, Refusal, Storage};
use litchi_numbers::table::pivot::{Aggregate, Category};
use prost::Message;

const BASIC: &[u8] = include_bytes!("../../../test-data/iwork/numbers/basic.numbers");
const ENGINE: &str = "Index/CalculationEngine.iwa";
/// The owner kind of `basic.numbers`' registered category formula owner
/// (internal identifier 13), which its native group-by names.
const NATIVE_OWNER_KIND: i32 = 8;
/// The inline group-by.
const INLINE_GROUP_BY: u64 = 9_100_010;
/// The cell the formula is authored in: C3.
const AT: At = At::new(2, 2);
const FORMULA_MESSAGE_LIST: u32 = 6_005;
const CELL_RECORD_TILE: u32 = 4_009;

/// `North / 2` with `North`'s category of the inline tree: field values of
/// the host's `FormulaArchive`, which the leaf's equalled byte for byte.
const INLINE_NORTH_FORMULA: &str = "0a4c0a2b0842ba02260a240a04080510091204080b1001180120023a08080010011800200052040815100268ffff030a190811210000000000000040d00202d80280808080808080a0300a020804320a08001000180020002801";

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

fn uuid(lower: u64, upper: u64) -> tsp::Uuid {
    tsp::Uuid { lower, upper }
}

fn coordinate(row: u32, column: u32) -> tsce::CellCoordinateArchive {
    tsce::CellCoordinateArchive {
        row: Some(row),
        column: Some(column),
        ..Default::default()
    }
}

fn value(kind: u8) -> tsce::CellValueArchive {
    use tsce::cell_value_archive::CellValueType;
    let mut value = tsce::CellValueArchive::default();
    match kind {
        0 => {
            value.cell_value_type = CellValueType::StringType as i32;
            value.string_value = Some(tsce::StringCellValueArchive {
                value: "North".to_owned(),
                format: tsk::FormatStructArchive::default(),
                ..Default::default()
            });
        },
        1 => {
            value.cell_value_type = CellValueType::NumberType as i32;
            value.number_value = Some(tsce::NumberCellValueArchive {
                value: Some(2.5),
                ..Default::default()
            });
        },
        2 => {
            value.cell_value_type = CellValueType::BooleanType as i32;
            value.boolean_value = Some(tsce::BooleanCellValueArchive {
                value: true,
                ..Default::default()
            });
        },
        _ => {
            value.cell_value_type = CellValueType::DateType as i32;
            value.date_value = Some(tsce::DateCellValueArchive {
                value: 7.25e8,
                format: tsk::FormatStructArchive::default(),
                ..Default::default()
            });
        },
    }
    value
}

fn node(
    group: (u64, u64),
    value: Option<tsce::CellValueArchive>,
    coordinates: &[(u32, u32)],
) -> tst::group_by_archive::GroupNodeArchive {
    tst::group_by_archive::GroupNodeArchive {
        group_uid: uuid(group.0, group.1),
        group_cell_value: value,
        agg_formula_coords: coordinates
            .iter()
            .map(|(row, column)| coordinate(*row, *column))
            .collect(),
        ..Default::default()
    }
}

fn aggregator(column: (u64, u64)) -> tst::group_by_archive::AggregatorArchive {
    tst::group_by_archive::AggregatorArchive {
        column_uid: uuid(column.0, column.1),
        ..Default::default()
    }
}

fn aggregate_type(column: (u64, u64), kind: u32) -> tst::ColumnAggregateArchive {
    tst::ColumnAggregateArchive {
        column_uid: uuid(column.0, column.1),
        level: 0,
        agg_type: kind,
        ..Default::default()
    }
}

/// The inline group-by: (identifier, message type, payload). `first_aggregate`
/// is the summary-function code of its first aggregated column.
fn inline_object(first_aggregate: u32) -> (u64, u32, Vec<u8>) {
    let columns = [(11, 1), (12, 1)];
    let mut root = node((1, 0), None, &[(0, 4), (0, 5)]);
    root.child = vec![
        node((21, 2), Some(value(0)), &[(1, 4), (1, 5)]),
        node((22, 2), Some(value(1)), &[(2, 4)]),
        node((23, 2), Some(value(2)), &[(3, 4), (3, 5)]),
        node((24, 2), Some(value(3)), &[(4, 4), (4, 5)]),
    ];
    root.child[0].child = vec![node((31, 3), None, &[(5, 4), (5, 5)])];
    (
        INLINE_GROUP_BY,
        6_373,
        tst::GroupByArchive {
            group_by_uid: uuid(5, 9),
            group_node_root: Some(root),
            aggregator: columns.iter().map(|column| aggregator(*column)).collect(),
            column_agg_type: vec![
                aggregate_type(columns[0], first_aggregate),
                aggregate_type(columns[1], 4),
            ],
            is_enabled: true,
            grouping_columns_formula: Some(coordinate(9, 1)),
            owner_index: Some(NATIVE_OWNER_KIND),
            ..Default::default()
        }
        .encode_to_vec(),
    )
}

fn with_objects(objects: Vec<(u64, u32, Vec<u8>)>) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(BASIC, Default::default())?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == ENGINE)
        .ok_or("basic.numbers has no calculation engine")?;
    let stream = SnappyStream::decompress(entry.data(), Default::default())?;
    let mut archive = Archive::parse(stream.as_bytes(), Default::default())?;
    for (identifier, message_type, data) in objects {
        archive.objects.push(ArchiveObject::new(
            identifier,
            vec![RawMessage {
                type_: message_type,
                data,
            }],
            Default::default(),
        )?);
    }
    let compressed =
        SnappyStream::compress(&archive.to_bytes(Default::default())?, Default::default())?;
    Ok(catalog.reassemble_to_bytes(&[EntryEdit::new(ENGINE, &compressed)], Default::default())?)
}

/// `basic.numbers` with the inline group-by.
fn inline_fixture() -> TestResult<Vec<u8>> {
    with_objects(vec![inline_object(1)])
}

fn read(bytes: &[u8]) -> TestResult<Vec<(Option<String>, i32, Aggregate)>> {
    Ok(Package::from_bytes(bytes)?
        .pivot_categories()?
        .into_iter()
        .map(|category| {
            (
                category.label().map(str::to_owned),
                category.group_level(),
                category.aggregate(),
            )
        })
        .collect())
}

fn expected(rows: &[(Option<&str>, i32, Aggregate)]) -> Vec<(Option<String>, i32, Aggregate)> {
    rows.iter()
        .map(|(label, level, aggregate)| (label.map(str::to_owned), *level, *aggregate))
        .collect()
}

/// The category labelled `North`.
fn north(package: &Package) -> TestResult<Category> {
    package
        .pivot_categories()?
        .into_iter()
        .find(|category| category.label() == Some("North"))
        .ok_or_else(|| "no category North".into())
}

fn expression(category: &Category) -> TestResult<Expression> {
    Ok(Expression::binary(
        BinaryOperator::Divide,
        Expression::pivot_category(category),
        Expression::number(2.0)?,
    )?)
}

fn exact_bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

fn archives(bytes: &[u8]) -> TestResult<Vec<Archive>> {
    let catalog = Catalog::from_bytes(bytes, Default::default())?;
    let mut archives = Vec::new();
    for entry in catalog
        .iter()
        .filter(|entry| entry.name().ends_with(".iwa"))
    {
        let stream = SnappyStream::decompress(entry.data(), Default::default())?;
        archives.push(Archive::parse(stream.as_bytes(), Default::default())?);
    }
    Ok(archives)
}

/// The one formula with a category node, re-encoded.
fn category_formula(bytes: &[u8]) -> TestResult<Vec<u8>> {
    let mut found = Vec::new();
    for archive in archives(bytes)? {
        for message in archive
            .objects
            .iter()
            .flat_map(|object| &object.messages)
            .filter(|message| message.type_ == FORMULA_MESSAGE_LIST)
        {
            let list = tst::TableDataList::decode(message.data.as_slice())?;
            for formula in list.entries.into_iter().filter_map(|entry| entry.formula) {
                if formula
                    .ast_node_array
                    .ast_node
                    .iter()
                    .any(|node| node.ast_category_ref.is_some())
                {
                    found.push(formula.encode_to_vec());
                }
            }
        }
    }
    match found.as_slice() {
        [one] => Ok(one.clone()),
        _ => Err(format!("expected one category formula, found {}", found.len()).into()),
    }
}

/// C3's `(row, column, owner)` dependency edges.
fn c3_edges(bytes: &[u8]) -> TestResult<Vec<(u32, u32, u32)>> {
    let mut edges = Vec::new();
    for archive in archives(bytes)? {
        for message in archive
            .objects
            .iter()
            .flat_map(|object| &object.messages)
            .filter(|message| message.type_ == CELL_RECORD_TILE)
        {
            let tile = tsce::CellRecordTileArchive::decode(message.data.as_slice())?;
            for record in tile.cell_records {
                if (record.row, record.column) == (2, 2) {
                    let expanded = record.expanded_edges.ok_or("C3 has no expanded edges")?;
                    edges.extend(
                        expanded
                            .edge_with_owner_rows
                            .iter()
                            .zip(&expanded.edge_with_owner_columns)
                            .zip(&expanded.internal_owner_id_for_edge)
                            .map(|((row, column), owner)| (*row, *column, *owner)),
                    );
                }
            }
        }
    }
    Ok(edges)
}

fn hex(bytes: &[u8]) -> String {
    bytes.iter().map(|byte| format!("{byte:02x}")).collect()
}

/// Author `category / 2` in C3 and pin what the host wrote too: the formula
/// archive, C3's edges (both precedents in the registered owner 13), a formula
/// cell after a reopen, and an inverse back to the input's bytes.
fn author(
    input: &[u8],
    cached: Option<f64>,
    formula: &str,
    edges: [(u32, u32, u32); 2],
) -> TestResult {
    let package = Package::from_bytes(input)?;
    let edit = package.edit_table_cells(0usize, 0usize)?;
    let category = north(&package)?;
    let commit = match cached {
        None => edit.set_formula(AT, expression(&category)?)?,
        Some(value) => {
            edit.set_formula_cached(AT, expression(&category)?, CachedValue::number(value)?)?
        },
    }
    .commit()?;
    let bytes = exact_bytes(commit.package())?;
    assert_eq!(hex(&category_formula(&bytes)?), formula);
    assert_eq!(c3_edges(&bytes)?, edges);
    let reopened = Package::from_bytes(&bytes)?;
    assert!(matches!(
        reopened.table_cell(0usize, 0usize, AT)?.storage(),
        Storage::Stored(Value::Formula(_))
    ));
    // Still readable, and the same categories: the edit adds no group.
    assert_eq!(read(&bytes)?, read(input)?);
    let restored = commit
        .package()
        .apply_table_cells(&commit.patch().inverse())?;
    assert_eq!(
        exact_bytes(restored.package())?,
        exact_bytes(&package)?,
        "inverse"
    );
    Ok(())
}

/// A workbook whose only group-by is disabled reads no category.
#[test]
fn plain_workbook_has_no_categories() -> TestResult {
    assert_eq!(read(BASIC)?, []);
    Ok(())
}

#[test]
fn reads_the_inline_tree() -> TestResult {
    use Aggregate::{CountAll, Minimum};
    assert_eq!(
        read(&inline_fixture()?)?,
        expected(&[
            (Some("Grand Total"), 0, CountAll),
            (Some("2.5"), 1, CountAll),
            (Some("725000000"), 1, CountAll),
            (Some("North"), 1, CountAll),
            (Some("TRUE"), 1, CountAll),
            (None, 2, CountAll),
            (Some("Grand Total"), 0, Minimum),
            (Some("725000000"), 1, Minimum),
            (Some("North"), 1, Minimum),
            (Some("TRUE"), 1, Minimum),
            (None, 2, Minimum),
        ])
    );
    Ok(())
}

#[test]
fn authors_a_category_of_the_inline_tree() -> TestResult {
    author(
        &inline_fixture()?,
        Some(21.5),
        INLINE_NORTH_FORMULA,
        [(1, 4, 13), (9, 1, 13)],
    )
}

/// A category formula is not evaluable, so with no supplied cache the leaf
/// writes the same formula and edges, as the host did, and no cached value:
/// its output differs from the output with one.
#[test]
fn authors_a_category_without_a_cache() -> TestResult {
    let input = inline_fixture()?;
    author(&input, None, INLINE_NORTH_FORMULA, [(1, 4, 13), (9, 1, 13)])?;
    let package = Package::from_bytes(&input)?;
    let category = north(&package)?;
    let without = package
        .edit_table_cells(0usize, 0usize)?
        .set_formula(AT, expression(&category)?)?
        .commit()?;
    let with = package
        .edit_table_cells(0usize, 0usize)?
        .set_formula_cached(AT, expression(&category)?, CachedValue::number(21.5)?)?
        .commit()?;
    assert_ne!(
        exact_bytes(without.package())?,
        exact_bytes(with.package())?
    );
    Ok(())
}

/// A category is resolved against the edited package: one read from another
/// workbook is refused when the edited one lacks it.
#[test]
fn a_category_absent_from_the_edited_package_is_refused() -> TestResult {
    let category = north(&Package::from_bytes(&inline_fixture()?)?)?;
    let refused = Package::from_bytes(BASIC)?
        .edit_table_cells(0usize, 0usize)?
        .set_formula_cached(AT, expression(&category)?, CachedValue::number(21.5)?)?
        .commit()
        .expect_err("a category absent from the edited package is refused");
    assert_eq!(refused, Error::Refused(Refusal::PivotCategoryNotFound));
    Ok(())
}

/// A category whose summary function differs in the edited package is refused,
/// and the changed category still reads, with its other function.
#[test]
fn a_category_with_another_function_in_the_edited_package_is_refused() -> TestResult {
    let category = north(&Package::from_bytes(&inline_fixture()?)?)?;
    let target = Package::from_bytes(&with_objects(vec![inline_object(3)])?)?;
    let refused = target
        .edit_table_cells(0usize, 0usize)?
        .set_formula(AT, expression(&category)?)?
        .commit()
        .expect_err("a category with another summary function is refused");
    assert_eq!(refused, Error::Refused(Refusal::PivotCategoryChanged));
    assert_ne!(north(&target)?.aggregate(), category.aggregate());
    Ok(())
}
