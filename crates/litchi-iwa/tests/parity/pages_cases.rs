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

/// Body-table cell writes (record 0370): the host's `set_table_cell` and
/// `clear_table_cell` against the leaf's `edit_body_table_values`, both over
/// the shared `litchi_numbers_wire::table_write` cell writer. For each body
/// table of each fixture: replace and clear every stored cell (at most 24),
/// and fill up to six empty cells; plus W1-B's case, clearing a row's only
/// unstyled cell after the host stored it (`body-table-catalog-native.pages`
/// D2).
#[test]
fn pages_body_table_cell_write_corpus_parity() -> TestResult {
    use litchi_numbers::cell::{FiniteF64, Value};
    let fixtures = [
        "iwork/pages/body-table-catalog-native.pages",
        "iwork/pages/body-table-cells-native.pages",
        "iwork/pages/body-table-merges-native.pages",
        "iwork/pages/body-table-read-native.pages",
        "iwork/pages/body-table-visible.pages",
    ];
    let mut inputs = Vec::new();
    for path in fixtures {
        inputs.push((path.to_owned(), fixture(path)?.bytes));
    }
    {
        // W1-B's pinned case: the host stores D2, then both routes clear it.
        let source = fixture("iwork/pages/body-table-catalog-native.pages")?;
        let mut editor = pages::Draft::from_bytes(&source.bytes)?;
        let model = editor
            .tables()?
            .first()
            .ok_or("no body table")?
            .model_object_id;
        editor.set_table_cell(model, 1, 3, Value::Text("Zürich".to_owned()))?;
        inputs.push((
            "iwork/pages/body-table-catalog-native.pages + host D2".to_owned(),
            editor.to_bytes()?,
        ));
    }
    let mut failures = Vec::new();
    let mut compared = 0_usize;
    for (path, bytes) in &inputs {
        let editor = pages::Draft::from_bytes(bytes)?;
        for (table_index, info) in editor.tables()?.into_iter().enumerate() {
            let table = editor.table(info.model_object_id)?;
            let stored: Vec<(usize, usize)> = table
                .iter_cells()
                .map(|(position, _)| position)
                .take(24)
                .collect();
            let mut empty = Vec::new();
            for row in 0..info.rows {
                for column in 0..info.columns {
                    if empty.len() < 6 && table.get_cell(row, column).is_none() {
                        empty.push((row, column));
                    }
                }
            }
            let mut ops = Vec::new();
            for cell in stored {
                ops.push(("replace", cell, Value::Text("parity".into())));
                ops.push(("clear", cell, Value::Empty));
            }
            for cell in empty {
                ops.push(("fill", cell, Value::Number(FiniteF64::new(7.0)?)));
            }
            for (label, (row, column), value) in ops {
                let host = (|| -> TestResult<Vec<u8>> {
                    let mut editor = pages::Draft::from_bytes(bytes)?;
                    if value == Value::Empty {
                        editor.clear_table_cell(info.model_object_id, row, column)?;
                    } else {
                        editor.set_table_cell(info.model_object_id, row, column, value.clone())?;
                    }
                    Ok(editor.to_bytes()?)
                })();
                let leaf = (|| -> TestResult<Vec<u8>> {
                    let package = litchi_pages::Package::from_bytes(bytes)?;
                    let mut edit = package.edit_body_table_values(table_index)?;
                    let leaf_value = match &value {
                        Value::Text(text) => litchi_pages::table::values::Value::Text(text.clone()),
                        Value::Number(number) => {
                            litchi_pages::table::values::Value::Number(number.get())
                        },
                        _ => litchi_pages::table::values::Value::Empty,
                    };
                    edit.set(u32::try_from(row)?, u32::try_from(column)?, leaf_value)?;
                    Ok(leaf_bytes!(edit.commit()?.package()))
                })();
                let label = format!("{path} table {table_index} {label} ({row}, {column})");
                match (host, leaf) {
                    (Ok(host), Ok(leaf)) => {
                        compared += 1;
                        if host == leaf {
                            continue;
                        }
                        let diff = crate::archive_diff::compare_with_policies(
                            bytes,
                            &host,
                            &leaf,
                            &crate::known_fields::KnownFields::for_format(Format::Pages),
                            &[],
                            crate::archive_diff::MetadataPolicy::default(),
                        )?;
                        if !diff.violations.is_empty() {
                            failures.push(format!("{label}: tier B violations\n{}", diff.render()));
                        }
                    },
                    (Ok(_), Err(error)) => failures.push(format!(
                        "{label}: the leaf refuses what the host edits: {error}"
                    )),
                    (Err(host), Ok(_)) => {
                        eprintln!("{label}: host refuses ({host}); the leaf edits")
                    },
                    (Err(host), Err(leaf)) => {
                        eprintln!("{label}: both refuse (host {host}; leaf {leaf})");
                    },
                }
            }
        }
    }
    eprintln!("body-table cell writes compared: {compared}");
    if failures.is_empty() {
        Ok(())
    } else {
        Err(failures.join("\n").into())
    }
}
