//! Pages parity cases.

use litchi_iwa::pages;
use litchi_iwa_common::table::merge::Region;

use crate::harness::{Case, Evidence, Format, LeafRun, TestResult, Tier, fixture};
use crate::leaf_bytes;

const MERGES: &str = "iwork/pages/body-table-merges-native.pages";
const TABLE: usize = 0;

/// The host's native model identifier of body table `TABLE` (the host
/// route takes native identifiers; the leaf route takes the selector).
fn host_table_model_id(editor: &pages::Draft) -> TestResult<u64> {
    let tables = editor.tables()?;
    Ok(tables
        .get(TABLE)
        .ok_or("body table missing")?
        .model_object_id)
}

fn host_merges(bytes: &[u8]) -> TestResult<Vec<Region>> {
    let editor = pages::Draft::from_bytes(bytes)?;
    let id = host_table_model_id(&editor)?;
    Ok(editor.table_cell_merges(id)?)
}

fn leaf_merges(bytes: &[u8]) -> TestResult<Vec<Vec<Region>>> {
    let package = litchi_pages::Package::from_bytes(bytes)?;
    Ok(vec![
        package.body_table_merges(0)?,
        package.body_table_merges(1)?,
    ])
}

fn leaf_text(bytes: &[u8]) -> TestResult<String> {
    Ok(litchi_pages::Package::from_bytes(bytes)?
        .text()?
        .to_string())
}

enum MergeOp {
    Merge(Region),
    Unmerge(Region),
}

fn merge_case(name: &'static str, tier: Tier, evidence: Evidence, op: MergeOp) -> TestResult {
    let input = fixture(MERGES)?;
    let op = &op;
    Case::new(name, Format::Pages, tier, evidence, &input)
        .host(|source| {
            let mut editor = pages::Draft::from_bytes(source)?;
            let id = host_table_model_id(&editor)?;
            match *op {
                MergeOp::Merge(region) => editor.merge_table_cells(id, region)?,
                MergeOp::Unmerge(region) => {
                    if !editor.unmerge_table_cells(id, region)? {
                        return Err("host unmerge found no merge".into());
                    }
                },
            }
            Ok(editor.to_bytes()?)
        })
        .leaf(|source| {
            let package = litchi_pages::Package::from_bytes(source)?;
            let mut edit = package.edit_body_table_merges(TABLE)?;
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
                .apply_body_table_merges(&commit.patch().inverse())?;
            Ok(LeafRun::with_inverse(
                leaf_bytes!(commit.package()),
                leaf_bytes!(restored.package()),
            ))
        })
        .read("host table_cell_merges", host_merges)
        .read("leaf body_table_merges", leaf_merges)
        .read("leaf text", leaf_text)
        .run()?;
    Ok(())
}

#[test]
fn pages_body_table_merges_read_parity() -> TestResult {
    let input = fixture(MERGES)?;
    Case::new(
        "pages_body_table_merges_read_parity",
        Format::Pages,
        Tier::A,
        Evidence::Native("test-data/iwork/pages/body-table-merges-native-receipt.json"),
        &input,
    )
    .read_only()
    .host(|source| Ok(pages::Draft::from_bytes(source)?.to_bytes()?))
    .leaf(|source| {
        Ok(LeafRun::output(leaf_bytes!(
            litchi_pages::Package::from_bytes(source)?
        )))
    })
    .read("host table_cell_merges", host_merges)
    .read("leaf body_table_merges", leaf_merges)
    .read("leaf text", leaf_text)
    .run()?;
    Ok(())
}

#[test]
fn pages_body_table_unmerge_parity() -> TestResult {
    merge_case(
        "pages_body_table_unmerge_parity",
        Tier::A,
        Evidence::Queued,
        MergeOp::Unmerge(Region::new(3, 2, 1, 2)?),
    )
}

#[test]
fn pages_body_table_merge_parity() -> TestResult {
    merge_case(
        "pages_body_table_merge_parity",
        Tier::A,
        Evidence::Queued,
        MergeOp::Merge(Region::new(1, 0, 2, 1)?),
    )
}
