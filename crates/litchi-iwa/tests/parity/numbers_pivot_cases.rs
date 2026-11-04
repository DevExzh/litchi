//! Numbers pivot aggregate categories of wave W4-H (record 0421): the host's
//! `numbers::Draft::pivot_categories` (a Prost walk of the calculation
//! engine's group-bys) against `litchi_numbers::Package::pivot_categories`
//! (a bounded wire walk).
//!
//! No tracked native workbook holds a pivot table, so the case adds
//! group-by graphs to `basic.numbers`' calculation engine: an inline tree
//! (string, number, Boolean and date group values, a grand-total root,
//! two aggregated columns), a referenced tree (aggregators and group nodes
//! as separate objects, a cycle through a child reference), a disabled
//! group-by, a group-by whose owner kind names two formula owners, and a
//! category two group-bys describe differently (dropped as ambiguous).

use litchi_iwa::numbers;
use litchi_iwa_archive::iwa::{Archive, ArchiveObject, RawMessage, SnappyStream};
use litchi_iwa_archive::package::{Catalog, EntryEdit};
use litchi_iwa_protos::{tsce, tsk, tsp, tst};
use litchi_numbers::Package;
use prost::Message;

use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier, fixture};
use crate::leaf_bytes;

const ENGINE: &str = "Index/CalculationEngine.iwa";
/// The owner kind of `basic.numbers`' registered category formula owner
/// (internal identifier 13), which its native group-by names.
const NATIVE_OWNER_KIND: i32 = 8;

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

fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
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

fn owner(kind: u32, internal: u32) -> Vec<u8> {
    tsce::FormulaOwnerDependenciesArchive {
        internal_formula_owner_id: internal,
        owner_kind: Some(kind),
        ..Default::default()
    }
    .encode_to_vec()
}

/// The group-by objects: (identifier, message type, payload).
/// `first_aggregate` is the summary-function code of the inline tree's
/// first aggregated column.
fn objects(first_aggregate: u32) -> Vec<(u64, u32, Vec<u8>)> {
    let mut objects = Vec::new();
    // Formula owners: the group-bys name `basic.numbers`' registered
    // category owner (kind 8, the owner its native disabled group-by names;
    // record 0424), and kind 41 names two owners.
    objects.push((9_100_002, 4_008, owner(41, 78)));
    objects.push((9_100_003, 4_008, owner(41, 79)));
    // An inline tree with two aggregated columns.
    let columns = [(11, 1), (12, 1)];
    let mut root = node((1, 0), None, &[(0, 4), (0, 5)]);
    root.child = vec![
        node((21, 2), Some(value(0)), &[(1, 4), (1, 5)]),
        node((22, 2), Some(value(1)), &[(2, 4)]),
        node((23, 2), Some(value(2)), &[(3, 4), (3, 5)]),
        node((24, 2), Some(value(3)), &[(4, 4), (4, 5)]),
    ];
    root.child[0].child = vec![node((31, 3), None, &[(5, 4), (5, 5)])];
    objects.push((
        9_100_010,
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
    ));
    // A referenced tree: aggregator and nodes as objects, with a cycle.
    objects.push((9_100_020, 6_382, aggregator((13, 2)).encode_to_vec()));
    let mut referenced_root = node((1, 0), None, &[(6, 2)]);
    referenced_root.child_ref = vec![reference(9_100_022), reference(9_100_022)];
    objects.push((9_100_021, 6_383, referenced_root.encode_to_vec()));
    let mut child = node((41, 4), Some(value(0)), &[(7, 2)]);
    child.child_ref = vec![reference(9_100_021)];
    objects.push((9_100_022, 6_383, child.encode_to_vec()));
    objects.push((
        9_100_023,
        6_373,
        tst::GroupByArchive {
            group_by_uid: uuid(6, 0),
            group_node_root_ref: Some(reference(9_100_021)),
            aggregator_ref: vec![reference(9_100_020)],
            column_agg_type: vec![aggregate_type((13, 2), 2)],
            is_enabled: true,
            grouping_columns_formula: Some(coordinate(8, 1)),
            owner_index: Some(NATIVE_OWNER_KIND),
            ..Default::default()
        }
        .encode_to_vec(),
    ));
    // Disabled, and an ambiguous owner kind: no categories.
    for (identifier, enabled, kind) in
        [(9_100_030, false, NATIVE_OWNER_KIND), (9_100_031, true, 41)]
    {
        objects.push((
            identifier,
            6_373,
            tst::GroupByArchive {
                group_by_uid: uuid(identifier, 1),
                group_node_root: Some(node((1, 0), None, &[(0, 1)])),
                aggregator: vec![aggregator((14, 1))],
                column_agg_type: vec![aggregate_type((14, 1), 1)],
                is_enabled: enabled,
                grouping_columns_formula: Some(coordinate(1, 1)),
                owner_index: Some(kind),
                ..Default::default()
            }
            .encode_to_vec(),
        ));
    }
    // A second description of the inline tree's first category with a
    // different aggregate cell: both are dropped as ambiguous.
    objects.push((
        9_100_040,
        6_373,
        tst::GroupByArchive {
            group_by_uid: uuid(5, 9),
            group_node_root: Some(node((1, 0), None, &[(9, 9)])),
            aggregator: vec![aggregator(columns[0])],
            column_agg_type: vec![aggregate_type(columns[0], 1)],
            is_enabled: true,
            grouping_columns_formula: Some(coordinate(9, 1)),
            owner_index: Some(NATIVE_OWNER_KIND),
            ..Default::default()
        }
        .encode_to_vec(),
    ));
    objects
}

/// `basic.numbers` with the group-by graphs added to its calculation engine.
pub(crate) fn pivot_fixture() -> TestResult<Fixture> {
    pivot_fixture_with(1)
}

/// The synthetic fixture whose inline tree's first column uses the
/// summary-function code `first_aggregate`.
pub(crate) fn pivot_fixture_with(first_aggregate: u32) -> TestResult<Fixture> {
    pivot_fixture_of(
        objects(first_aggregate),
        "iwork/numbers/basic.numbers with synthetic pivot group-bys",
    )
}

/// `basic.numbers` with only the inline group-by: no reference cycle, no
/// repeated group-by identifier, no second owner kind. Numbers 14.4 drops its
/// AppleEvent connection on the full synthetic fixture even unedited (record
/// 0431); this one separates the fixture from the formula route.
pub(crate) fn inline_pivot_fixture() -> TestResult<Fixture> {
    const INLINE_GROUP_BY: u64 = 9_100_010;
    pivot_fixture_of(
        objects(1)
            .into_iter()
            .filter(|(identifier, _, _)| *identifier == INLINE_GROUP_BY)
            .collect(),
        "iwork/numbers/basic.numbers with one inline synthetic pivot group-by",
    )
}

fn pivot_fixture_of(objects: Vec<(u64, u32, Vec<u8>)>, path: &'static str) -> TestResult<Fixture> {
    let input = fixture("iwork/numbers/basic.numbers")?;
    let catalog = Catalog::from_bytes(&input.bytes, Default::default())?;
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
    Ok(Fixture {
        path,
        bytes: catalog
            .reassemble_to_bytes(&[EntryEdit::new(ENGINE, &compressed)], Default::default())?,
    })
}

/// Number equal values by first occurrence: two lists of summary functions
/// agree when they partition the categories identically (the leaf's
/// `Aggregate` exposes no native code, record 0424).
fn classes<T: PartialEq>(values: &[T]) -> Vec<usize> {
    values
        .iter()
        .map(|value| {
            values
                .iter()
                .position(|candidate| candidate == value)
                .unwrap_or(usize::MAX)
        })
        .collect()
}

/// The categories both routes report: label, level and summary-function
/// class, in order.
fn host_read(bytes: &[u8]) -> TestResult<Vec<(Option<String>, i32, usize)>> {
    let categories = numbers::Draft::from_bytes(bytes)?.pivot_categories()?;
    let aggregates = classes(
        &categories
            .iter()
            .map(|category| category.reference.aggregate_type)
            .collect::<Vec<_>>(),
    );
    Ok(categories
        .into_iter()
        .zip(aggregates)
        .map(|(category, aggregate)| (category.label, category.reference.group_level, aggregate))
        .collect())
}

fn leaf_read(bytes: &[u8]) -> TestResult<Vec<(Option<String>, i32, usize)>> {
    let categories = Package::from_bytes(bytes)?.pivot_categories()?;
    let aggregates = classes(
        &categories
            .iter()
            .map(litchi_numbers::table::pivot::Category::aggregate)
            .collect::<Vec<_>>(),
    );
    Ok(categories
        .into_iter()
        .zip(aggregates)
        .map(|(category, aggregate)| {
            (
                category.label().map(str::to_owned),
                category.group_level(),
                aggregate,
            )
        })
        .collect())
}

/// The host's and the leaf's categories agree on a workbook without
/// pivots and on the synthetic group-bys.
#[test]
fn numbers_pivot_categories_read_parity() -> TestResult {
    let plain = fixture("iwork/numbers/basic.numbers")?;
    let pivot = pivot_fixture()?;
    for source in [&plain, &pivot] {
        Case::new(
            "numbers_pivot_categories_read_parity",
            Format::Numbers,
            Tier::A,
            Evidence::Queued,
            source,
        )
        .read_only()
        .host(|bytes| Ok(numbers::Draft::from_bytes(bytes)?.to_bytes()?))
        .leaf(|bytes| Ok(LeafRun::output(leaf_bytes!(Package::from_bytes(bytes)?))))
        .read("host pivot_categories", host_read)
        .read("leaf pivot_categories", leaf_read)
        .run()?;
        assert_eq!(host_read(&source.bytes)?, leaf_read(&source.bytes)?);
    }
    let categories = leaf_read(&pivot.bytes)?;
    assert!(categories.len() >= 8, "{categories:?}");
    // The synthetic group-bys use three summary-function codes (1, 2, 4).
    let mut distinct = categories
        .iter()
        .map(|(_, _, class)| *class)
        .collect::<Vec<_>>();
    distinct.sort_unstable();
    distinct.dedup();
    assert_eq!(distinct.len(), 3, "{categories:?}");
    for label in ["North", "2.5", "TRUE", "Grand Total"] {
        assert!(
            categories
                .iter()
                .any(|(found, _, _)| found.as_deref() == Some(label)),
            "{label}: {categories:?}"
        );
    }
    Ok(())
}

/// Callgrind probe (record 0421): read the synthetic categories from
/// already opened packages, host against leaf, ABBA.
#[test]
#[ignore = "callgrind probe; run explicitly"]
fn numbers_pivot_categories_cost_abba() -> TestResult {
    let input = pivot_fixture()?;
    let host = numbers::Draft::from_bytes(&input.bytes)?;
    let leaf = Package::from_bytes(&input.bytes)?;
    crate::pages_cost_probes::abba(
        &input,
        |_source| Ok(format!("{:?}", host.pivot_categories()?).into_bytes()),
        |_source| Ok(format!("{:?}", leaf.pivot_categories()?).into_bytes()),
    )
}
