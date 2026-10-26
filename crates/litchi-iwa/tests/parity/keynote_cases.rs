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

/// Slide-table cell writes (record 0370): the host's `set_slide_table_cell`
/// against the leaf's `edit_slide_drawables().set_table_cell`, both over the
/// shared `litchi_numbers_wire::table_write` cell writer. For each slide
/// table of each fixture: replace and clear every stored cell, and fill up
/// to six empty cells with a number and with text.
#[test]
fn keynote_slide_table_cell_write_corpus_parity() -> TestResult {
    use litchi_numbers::cell::Value;
    let fixtures = [
        "iwork/keynote/table-discovery.key",
        "iwork/keynote/slide-table-merges-native.key",
        "iwork/keynote/slide-table-read-native.key",
    ];
    let mut failures = Vec::new();
    let mut compared = 0_usize;
    for path in fixtures {
        let input = fixture(path)?;
        let editor = keynote::Draft::from_bytes(&input.bytes)?;
        for slide in 0..editor.slides()?.len() {
            for (table_index, info) in editor.slide_tables(slide)?.into_iter().enumerate() {
                let table = editor.slide_table(slide, info.model_object_id)?;
                let stored: Vec<(usize, usize)> =
                    table.iter_cells().map(|(position, _)| position).collect();
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
                    ops.push((
                        "fill",
                        cell,
                        Value::Number(litchi_numbers::cell::FiniteF64::new(7.0)?),
                    ));
                    ops.push(("fill text", cell, Value::Text("new".into())));
                }
                for (label, (row, column), value) in ops {
                    let host = (|| -> TestResult<Vec<u8>> {
                        let mut editor = keynote::Draft::from_bytes(&input.bytes)?;
                        editor.set_slide_table_cell(
                            slide,
                            info.model_object_id,
                            row,
                            column,
                            value.clone(),
                        )?;
                        Ok(editor.to_bytes()?)
                    })();
                    let leaf_value = match &value {
                        Value::Text(text) => {
                            litchi_keynote::slide::drawable::TableCellValue::Text(text.clone())
                        },
                        Value::Number(number) => {
                            litchi_keynote::slide::drawable::TableCellValue::Number(number.get())
                        },
                        _ => litchi_keynote::slide::drawable::TableCellValue::Empty,
                    };
                    let leaf = (|| -> TestResult<Vec<u8>> {
                        let package = litchi_keynote::Package::from_bytes(&input.bytes)?;
                        let mut last = None;
                        for drawable in 0..64 {
                            let mut edit = package.edit_slide_drawables();
                            match edit.set_table_cell(
                                slide,
                                litchi_keynote::DrawableSelector::index(drawable),
                                u32::try_from(row)?,
                                u32::try_from(column)?,
                                leaf_value.clone(),
                            ) {
                                Ok(_) => {},
                                Err(error) => {
                                    last = Some(error.to_string());
                                    continue;
                                },
                            }
                            // The `table_index`-th table drawable of the slide.
                            let _ = table_index;
                            return Ok(leaf_bytes!(edit.commit()?.package()));
                        }
                        Err(last
                            .unwrap_or_else(|| "no table drawable".to_owned())
                            .into())
                    })();
                    let label = format!(
                        "{path} slide {slide} table {table_index} {label} ({row}, {column})"
                    );
                    match (host, leaf) {
                        (Ok(host), Ok(leaf)) => {
                            compared += 1;
                            if host == leaf {
                                continue;
                            }
                            let diff = crate::archive_diff::compare_with_policies(
                                &input.bytes,
                                &host,
                                &leaf,
                                &crate::known_fields::KnownFields::for_format(Format::Keynote),
                                &["preview.jpg", "preview-micro.jpg", "preview-web.jpg"],
                                crate::archive_diff::MetadataPolicy {
                                    save_token: true,
                                    fresh_uuids: true,
                                },
                            )?;
                            if !diff.violations.is_empty() {
                                failures
                                    .push(format!("{label}: tier B violations\n{}", diff.render()));
                            }
                        },
                        (Ok(_), Err(error)) => failures.push(format!(
                            "{label}: the leaf refuses what the host edits: {error}"
                        )),
                        (Err(host), Ok(_)) => {
                            eprintln!("{label}: host refuses ({host}); the leaf edits")
                        },
                        (Err(host), Err(leaf)) => {
                            eprintln!("{label}: both refuse (host {host}; leaf {leaf})")
                        },
                    }
                }
            }
        }
    }
    eprintln!("slide-table cell writes compared: {compared}");
    if failures.is_empty() {
        Ok(())
    } else {
        Err(failures.join("\n").into())
    }
}
