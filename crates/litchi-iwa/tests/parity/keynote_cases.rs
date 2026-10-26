//! Keynote parity cases.

use litchi_iwa::keynote;
use litchi_iwa_common::table::merge::Region;

use crate::harness::{Case, Evidence, Format, LeafRun, TestResult, Tier, fixture};
use crate::leaf_bytes;

const MERGES: &str = "iwork/keynote/slide-table-merges-native.key";
const SLIDE: usize = 0;
const TABLE: usize = 0;

/// The host's native model identifier of slide table `TABLE`.
fn host_table_model_id(editor: &keynote::Draft) -> TestResult<u64> {
    let tables = editor.slide_tables(SLIDE)?;
    Ok(tables
        .get(TABLE)
        .ok_or("slide table missing")?
        .model_object_id)
}

fn host_merges(bytes: &[u8]) -> TestResult<Vec<Region>> {
    let editor = keynote::Draft::from_bytes(bytes)?;
    let id = host_table_model_id(&editor)?;
    Ok(editor.slide_table_cell_merges(SLIDE, id)?)
}

fn leaf_merges(bytes: &[u8]) -> TestResult<Vec<Region>> {
    Ok(litchi_keynote::Package::from_bytes(bytes)?.slide_table_merges(SLIDE, TABLE)?)
}

enum MergeOp {
    Merge(Region),
    Unmerge(Region),
}

fn merge_case(name: &'static str, tier: Tier, evidence: Evidence, op: MergeOp) -> TestResult {
    let input = fixture(MERGES)?;
    let op = &op;
    Case::new(name, Format::Keynote, tier, evidence, &input)
        .host(|source| {
            let mut editor = keynote::Draft::from_bytes(source)?;
            let id = host_table_model_id(&editor)?;
            match *op {
                MergeOp::Merge(region) => editor.merge_slide_table_cells(SLIDE, id, region)?,
                MergeOp::Unmerge(region) => {
                    if !editor.unmerge_slide_table_cells(SLIDE, id, region)? {
                        return Err("host unmerge found no merge".into());
                    }
                },
            }
            Ok(editor.to_bytes()?)
        })
        .leaf(|source| {
            let package = litchi_keynote::Package::from_bytes(source)?;
            let mut edit = package.edit_slide_table_merges(SLIDE, TABLE)?;
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
                .apply_slide_table_merges(&commit.patch().inverse())?;
            Ok(LeafRun::with_inverse(
                leaf_bytes!(commit.package()),
                leaf_bytes!(restored.package()),
            ))
        })
        .read("host slide_table_cell_merges", host_merges)
        .read("leaf slide_table_merges", leaf_merges)
        .run()?;
    Ok(())
}

#[test]
fn keynote_slide_table_merges_read_parity() -> TestResult {
    let input = fixture(MERGES)?;
    Case::new(
        "keynote_slide_table_merges_read_parity",
        Format::Keynote,
        Tier::A,
        Evidence::Native("test-data/iwork/keynote/slide-table-merges-native-receipt.json"),
        &input,
    )
    .read_only()
    .host(|source| Ok(keynote::Draft::from_bytes(source)?.to_bytes()?))
    .leaf(|source| {
        Ok(LeafRun::output(leaf_bytes!(
            litchi_keynote::Package::from_bytes(source)?
        )))
    })
    .read("host slide_table_cell_merges", host_merges)
    .read("leaf slide_table_merges", leaf_merges)
    .run()?;
    Ok(())
}

#[test]
fn keynote_slide_table_unmerge_parity() -> TestResult {
    merge_case(
        "keynote_slide_table_unmerge_parity",
        Tier::A,
        Evidence::Queued,
        MergeOp::Unmerge(Region::new(3, 1, 2, 2)?),
    )
}

#[test]
fn keynote_slide_table_merge_parity() -> TestResult {
    merge_case(
        "keynote_slide_table_merge_parity",
        Tier::A,
        Evidence::Queued,
        MergeOp::Merge(Region::new(1, 0, 2, 1)?),
    )
}
