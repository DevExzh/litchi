//! Numbers 14.4 overflows its own stack on a group-node tree with a child
//! reference cycle (record 0433: 1658 frames of one function, under the
//! document's load). Nothing in Litchi writes a child reference; these files
//! must still read, on a stack as small as the one Numbers crashed on.

use litchi_iwa_archive::iwa::{Archive, ArchiveObject, RawMessage, SnappyStream};
use litchi_iwa_archive::package::{Catalog, EntryEdit};
use litchi_iwa_protos::{tsce, tsp, tst};
use litchi_numbers::Package;
use prost::Message;

const BASIC: &[u8] = include_bytes!("../../../test-data/iwork/numbers/basic.numbers");
const ENGINE: &str = "Index/CalculationEngine.iwa";
/// The registered category owner of `basic.numbers` (record 0424).
const OWNER_KIND: i32 = 8;
/// Below the 544 KiB secondary stack of the crash report, to leave no margin.
const SMALL_STACK: usize = 256 * 1024;
const DEEP_INLINE_LEVELS: usize = 4_000;

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

fn uuid(lower: u64, upper: u64) -> tsp::Uuid {
    tsp::Uuid { lower, upper }
}

fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..Default::default()
    }
}

fn node(group: u64, row: u32) -> tst::group_by_archive::GroupNodeArchive {
    tst::group_by_archive::GroupNodeArchive {
        group_uid: uuid(group, 7),
        agg_formula_coords: vec![tsce::CellCoordinateArchive {
            row: Some(row),
            column: Some(2),
            ..Default::default()
        }],
        ..Default::default()
    }
}

fn group_by(
    root: Option<tst::group_by_archive::GroupNodeArchive>,
    root_reference: Option<u64>,
) -> Vec<u8> {
    let column = uuid(13, 2);
    tst::GroupByArchive {
        group_by_uid: uuid(6, 0),
        group_node_root: root,
        group_node_root_ref: root_reference.map(reference),
        aggregator: vec![tst::group_by_archive::AggregatorArchive {
            column_uid: column,
            ..Default::default()
        }],
        column_agg_type: vec![tst::ColumnAggregateArchive {
            column_uid: column,
            level: 0,
            agg_type: 2,
            ..Default::default()
        }],
        is_enabled: true,
        grouping_columns_formula: Some(tsce::CellCoordinateArchive {
            row: Some(8),
            column: Some(1),
            ..Default::default()
        }),
        owner_index: Some(OWNER_KIND),
        ..Default::default()
    }
    .encode_to_vec()
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

/// `read` on a small stack: an overflow aborts the test binary.
fn on_small_stack<T: Send + 'static>(read: impl FnOnce() -> T + Send + 'static) -> TestResult<T> {
    Ok(std::thread::Builder::new()
        .stack_size(SMALL_STACK)
        .spawn(read)?
        .join()
        .map_err(|_| "the read panicked")?)
}

/// A two-node cycle, a self-reference and a repeated child: each node is
/// visited once.
#[test]
fn a_child_reference_cycle_terminates() -> TestResult {
    let mut root = node(1, 6);
    root.child_ref = vec![
        reference(9_200_002),
        reference(9_200_002),
        reference(9_200_001),
    ];
    let mut child = node(41, 7);
    child.child_ref = vec![reference(9_200_001), reference(9_200_002)];
    let bytes = with_objects(vec![
        (9_200_001, 6_383, root.encode_to_vec()),
        (9_200_002, 6_383, child.encode_to_vec()),
        (9_200_003, 6_373, group_by(None, Some(9_200_001))),
    ])?;
    let categories = on_small_stack(move || -> Result<usize, String> {
        let package = Package::from_bytes(&bytes).map_err(|error| error.to_string())?;
        let categories = package
            .pivot_categories()
            .map_err(|error| error.to_string())?;
        // The table reads too: the walk is not the only route over the engine.
        package
            .table_cell(0usize, 0usize, litchi_numbers::table::At::new(1, 1))
            .map_err(|error| error.to_string())?;
        Ok(categories.len())
    })??;
    assert_eq!(categories, 2);
    Ok(())
}

/// An inline tree nested far past any recursion limit: the walk is iterative.
#[test]
fn a_deep_inline_tree_does_not_recurse() -> TestResult {
    let mut root = node(u64::try_from(DEEP_INLINE_LEVELS)? + 1, 1);
    for level in (1..=DEEP_INLINE_LEVELS).rev() {
        let mut parent = node(u64::try_from(level)?, 1);
        parent.child = vec![std::mem::take(&mut root)];
        root = parent;
    }
    let encoded = group_by(Some(root), None);
    let bytes = with_objects(vec![(9_200_003, 6_373, encoded)])?;
    let categories = on_small_stack(move || -> Result<usize, String> {
        let package = Package::from_bytes(&bytes).map_err(|error| error.to_string())?;
        Ok(package
            .pivot_categories()
            .map_err(|error| error.to_string())?
            .len())
    })??;
    assert_eq!(categories, DEEP_INLINE_LEVELS + 1);
    Ok(())
}
