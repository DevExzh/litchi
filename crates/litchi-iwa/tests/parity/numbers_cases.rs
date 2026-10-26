//! Numbers parity cases.

use litchi_iwa::numbers;
use litchi_iwa_common::table::merge::Region;

use crate::harness::{Case, Evidence, Format, LeafRun, Report, TestResult, Tier, fixture};
use crate::leaf_bytes;

const MERGES: &str = "iwork/numbers/table-merges-native.numbers";
const SHEET: &str = "Sheet 1";
const TABLE: &str = "shared-model";
const NATIVE_MERGE: (u32, u32, u32, u32) = (10, 1, 2, 2);

/// The host's native table-model identifier of the fixture's only table.
///
/// The host route takes native identifiers; the harness resolves them from
/// the same selector the leaf route uses (here: the only `TST.TableModel`).
fn host_table_model_id(editor: &numbers::Draft) -> TestResult<u64> {
    let mut ids = Vec::new();
    for name in editor.package().iwa_entry_names() {
        for object in editor.package().archive(name)?.objects {
            if object.messages.iter().any(|message| message.type_ == 6_001) {
                ids.extend(object.archive_info.identifier);
            }
        }
    }
    match ids.as_slice() {
        [id] => Ok(*id),
        _ => Err(format!("expected one table model, found {ids:?}").into()),
    }
}

fn host_merges(bytes: &[u8]) -> TestResult<Vec<Region>> {
    let editor = numbers::Draft::from_bytes(bytes)?;
    let id = host_table_model_id(&editor)?;
    Ok(editor.table_cell_merges(id)?)
}

fn leaf_merges(bytes: &[u8]) -> TestResult<Vec<Region>> {
    Ok(litchi_numbers::Package::from_bytes(bytes)?.table_merges(SHEET, TABLE)?)
}

enum MergeOp {
    Merge(Region),
    Unmerge(Region),
}

/// The operation expressed once, run through both routes.
fn merge_case(name: &'static str, tier: Tier, evidence: Evidence, op: MergeOp) -> TestResult {
    let input = fixture(MERGES)?;
    let op = &op;
    Case::new(name, Format::Numbers, tier, evidence, &input)
        .host(|source| {
            let mut editor = numbers::Draft::from_bytes(source)?;
            let id = host_table_model_id(&editor)?;
            match *op {
                MergeOp::Merge(region) => editor.merge_cells(id, region)?,
                MergeOp::Unmerge(region) => {
                    if !editor.unmerge_cells(id, region)? {
                        return Err("host unmerge found no merge".into());
                    }
                },
            }
            Ok(editor.to_bytes()?)
        })
        .leaf(|source| {
            let package = litchi_numbers::Package::from_bytes(source)?;
            let mut edit = package.edit_table_merges(SHEET, TABLE)?;
            match *op {
                MergeOp::Merge(region) => {
                    edit.merge(region)?;
                },
                MergeOp::Unmerge(region) => {
                    if !edit.unmerge(region)? {
                        return Err("leaf unmerge found no merge".into());
                    }
                },
            }
            let commit = edit.commit()?;
            let restored = commit
                .package()
                .apply_table_merges(&commit.patch().inverse())?;
            Ok(LeafRun::with_inverse(
                leaf_bytes!(commit.package()),
                leaf_bytes!(restored.package()),
            ))
        })
        .read("host table_cell_merges", host_merges)
        .read("leaf table_merges", leaf_merges)
        .run()?;
    Ok(())
}

fn native_region() -> TestResult<Region> {
    let (row, column, rows, columns) = NATIVE_MERGE;
    Ok(Region::new(row, column, rows, columns)?)
}

#[test]
fn numbers_table_merges_read_parity() -> TestResult {
    let input = fixture(MERGES)?;
    Case::new(
        "numbers_table_merges_read_parity",
        Format::Numbers,
        Tier::A,
        Evidence::Native("test-data/iwork/numbers/table-merges-native-receipt.json"),
        &input,
    )
    .read_only()
    .host(|source| Ok(numbers::Draft::from_bytes(source)?.to_bytes()?))
    .leaf(|source| {
        Ok(LeafRun::output(leaf_bytes!(
            litchi_numbers::Package::from_bytes(source)?
        )))
    })
    .read("host table_cell_merges", host_merges)
    .read("leaf table_merges", leaf_merges)
    .read("expected native merge", |bytes| {
        Ok(leaf_merges(bytes)? == vec![native_region()?])
    })
    .run()?;
    Ok(())
}

#[test]
fn numbers_table_unmerge_parity() -> TestResult {
    merge_case(
        "numbers_table_unmerge_parity",
        Tier::A,
        Evidence::Queued,
        MergeOp::Unmerge(native_region()?),
    )
}

#[test]
fn numbers_table_merge_parity() -> TestResult {
    merge_case(
        "numbers_table_merge_parity",
        Tier::A,
        Evidence::Queued,
        MergeOp::Merge(Region::new(2, 0, 2, 2)?),
    )
}

/// `basic.numbers` (source-built by the host) through the row-insertion
/// routes. The crosswalk lists `insert_table_row` as at parity
/// (`edit_table_topology`), but the harness finds the two routes differ on
/// this fixture; until a wave resolves it, this test pins the divergence
/// (record 0369 §4) so the route stays listed, not shipped.
const TOPOLOGY: &str = "iwork/numbers/basic.numbers";

fn leaf_sheets(bytes: &[u8]) -> TestResult<String> {
    let package = litchi_numbers::Package::from_bytes(bytes)?;
    Ok(format!("{:?}", package.sheets()))
}

fn insert_row_case(name: &'static str, tier: Tier, evidence: Evidence) -> TestResult<Report> {
    use litchi_numbers::table::topology::RowInsertion;
    let input = fixture(TOPOLOGY)?;
    Case::new(name, Format::Numbers, tier, evidence, &input)
        .host(|source| {
            let mut editor = numbers::Draft::from_bytes(source)?;
            editor.insert_table_row(
                litchi_numbers::TableSelector::index(0),
                RowInsertion::body(1),
            )?;
            Ok(editor.to_bytes()?)
        })
        .leaf(|source| {
            let package = litchi_numbers::Package::from_bytes(source)?;
            let commit = package
                .edit_table_topology(0usize, 0usize)?
                .insert_row(RowInsertion::body(1))?
                .commit()?;
            let restored = commit
                .package()
                .apply_table_topology(&commit.patch().inverse())?;
            Ok(LeafRun::with_inverse(
                leaf_bytes!(commit.package()),
                leaf_bytes!(restored.package()),
            ))
        })
        .read("leaf sheets", leaf_sheets)
        .run()
}

#[test]
fn numbers_table_insert_row_divergence_is_reported() -> TestResult {
    let Err(error) = insert_row_case(
        "numbers_table_insert_row_divergence_is_reported",
        Tier::B,
        Evidence::Queued,
    ) else {
        return Err("the row-insertion routes now agree: turn this into a parity case".into());
    };
    let text = error.to_string();
    eprintln!("{text}");
    assert!(text.contains("ZIP member sets differ"), "{text}");
    assert!(text.contains("(TST.Tile): known fields differ"), "{text}");
    Ok(())
}
