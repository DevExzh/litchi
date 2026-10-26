//! W1-F (record 0376): value writes into Keynote slide tables and Pages body
//! tables that formulas read, through the shared formula-cache refresh
//! (`litchi_numbers_wire::formula_cache::refresh`), against the host's
//! `formula_cache::refresh_formula_caches_after_cell_write`.

use litchi_iwa::{keynote, pages};
use litchi_numbers::cell::{FiniteF64, Value};

use crate::harness::{Case, Evidence, Format, LeafRun, TestResult, Tier, fixture};

/// The cached results stored at the tile-local `row`/`column` of every
/// tile of the package that holds a formula there.
fn formula_caches(bytes: &[u8], row: u32, column: u32) -> TestResult<String> {
    use litchi_numbers_wire::{BncCellView, Limits, StoredValue, table_write};
    let package = litchi_iwa::Package::from_bytes(bytes)?;
    let names: Vec<String> = package.iwa_entry_names().map(str::to_owned).collect();
    let mut caches = Vec::new();
    for name in names {
        for object in package.archive(&name)?.objects {
            for message in object
                .messages
                .iter()
                .filter(|message| message.type_ == 6_002)
            {
                let Some(cell) =
                    table_write::tile_cell(&message.data, row, column, &Limits::DEFAULT)?
                else {
                    continue;
                };
                let view = BncCellView::parse(&cell, &Limits::DEFAULT)?;
                if matches!(view.stored_value(), StoredValue::Formula(_)) {
                    caches.push(format!("{:?}", view.cached_scalar()));
                }
            }
        }
    }
    Ok(caches.join(", "))
}
use crate::leaf_bytes;

/// One slide table with a formula at D2 reading B2.
const KEYNOTE: &str = "iwork/keynote/slide-table-read-native.key";
/// Table 1 sums B2:B3 into B4; table 2 reads table 1.
const PAGES: &str = "iwork/pages/body-table-cells-native.pages";

fn keynote_model(editor: &keynote::Draft) -> TestResult<u64> {
    Ok(editor
        .slide_tables(0)?
        .first()
        .ok_or("slide table missing")?
        .model_object_id)
}

/// The host's view of one cell, cached formula result included.
fn keynote_cell(bytes: &[u8], row: usize, column: usize) -> TestResult<String> {
    let editor = keynote::Draft::from_bytes(bytes)?;
    let table = editor.slide_table(0, keynote_model(&editor)?)?;
    Ok(format!("{:?}", table.get_cell(row, column)))
}

fn keynote_case(name: &'static str, tier: Tier, evidence: Evidence, value: Value) -> TestResult {
    let input = fixture(KEYNOTE)?;
    let value = &value;
    Case::new(name, Format::Keynote, tier, evidence, &input)
        .host(|source| {
            let mut editor = keynote::Draft::from_bytes(source)?;
            let model = keynote_model(&editor)?;
            editor.set_slide_table_cell(0, model, 1, 1, value.clone())?;
            Ok(editor.to_bytes()?)
        })
        .leaf(|source| {
            use litchi_keynote::slide::drawable::TableCellValue;
            let leaf_value = match value {
                Value::Number(number) => TableCellValue::Number(number.get()),
                Value::Text(text) => TableCellValue::Text(text.clone()),
                _ => TableCellValue::Empty,
            };
            let package = litchi_keynote::Package::from_bytes(source)?;
            for drawable in 0..64 {
                let mut edit = package.edit_slide_drawables();
                if edit
                    .set_table_cell(
                        0,
                        litchi_keynote::DrawableSelector::index(drawable),
                        1,
                        1,
                        leaf_value.clone(),
                    )
                    .is_ok()
                {
                    let output = leaf_bytes!(edit.commit()?.package());
                    // The write must refresh D2's cached result.
                    if formula_caches(&output, 1, 3)? == formula_caches(source, 1, 3)? {
                        return Err("the leaf left D2's cached result unchanged".into());
                    }
                    return Ok(LeafRun::output(output));
                }
            }
            Err("no table drawable".into())
        })
        .leaf_invalidates_root_previews()
        .leaf_structural_metadata()
        .read("cached result of D2", |bytes| formula_caches(bytes, 1, 3))
        .read("host formula cell D2", |bytes| keynote_cell(bytes, 1, 3))
        .read("host precedent cell B2", |bytes| keynote_cell(bytes, 1, 1))
        .run()?;
    Ok(())
}

#[test]
fn keynote_slide_table_formula_refresh_parity() -> TestResult {
    keynote_case(
        "keynote_slide_table_formula_refresh_parity",
        Tier::B,
        Evidence::Queued,
        Value::Number(FiniteF64::new(77.0)?),
    )
}

#[test]
fn keynote_slide_table_formula_refresh_clear_parity() -> TestResult {
    keynote_case(
        "keynote_slide_table_formula_refresh_clear_parity",
        Tier::B,
        Evidence::Queued,
        Value::Empty,
    )
}

fn pages_model(editor: &pages::Draft) -> TestResult<u64> {
    Ok(editor
        .tables()?
        .first()
        .ok_or("body table missing")?
        .model_object_id)
}

fn pages_cell(bytes: &[u8], row: usize, column: usize) -> TestResult<String> {
    let editor = pages::Draft::from_bytes(bytes)?;
    let table = editor.table(pages_model(&editor)?)?;
    Ok(format!("{:?}", table.get_cell(row, column)))
}

fn pages_case(
    name: &'static str,
    tier: Tier,
    evidence: Evidence,
    row: u32,
    value: Value,
) -> TestResult {
    let input = fixture(PAGES)?;
    let value = &value;
    Case::new(name, Format::Pages, tier, evidence, &input)
        .host(|source| {
            let mut editor = pages::Draft::from_bytes(source)?;
            let model = pages_model(&editor)?;
            if *value == Value::Empty {
                editor.clear_table_cell(model, usize::try_from(row)?, 1)?;
            } else {
                editor.set_table_cell(model, usize::try_from(row)?, 1, value.clone())?;
            }
            Ok(editor.to_bytes()?)
        })
        .leaf(|source| {
            use litchi_pages::table::values::Value as Leaf;
            let leaf_value = match value {
                Value::Number(number) => Leaf::Number(number.get()),
                Value::Text(text) => Leaf::Text(text.clone()),
                _ => Leaf::Empty,
            };
            let package = litchi_pages::Package::from_bytes(source)?;
            let mut edit = package.edit_body_table_values(0usize)?;
            edit.set(row, 1, leaf_value)?;
            let output = leaf_bytes!(edit.commit()?.package());
            // The write must refresh B4's cached result.
            if formula_caches(&output, 3, 1)? == formula_caches(source, 3, 1)? {
                return Err("the leaf left B4's cached result unchanged".into());
            }
            Ok(LeafRun::output(output))
        })
        .read("cached result of B4", |bytes| formula_caches(bytes, 3, 1))
        .read("host formula cell B4", |bytes| pages_cell(bytes, 3, 1))
        .read("host written cell", |bytes| {
            pages_cell(bytes, usize::try_from(row)?, 1)
        })
        .run()?;
    Ok(())
}

#[test]
fn pages_body_table_formula_refresh_parity() -> TestResult {
    pages_case(
        "pages_body_table_formula_refresh_parity",
        Tier::B,
        Evidence::Queued,
        1,
        Value::Number(FiniteF64::new(41.0)?),
    )
}

#[test]
fn pages_body_table_formula_refresh_clear_parity() -> TestResult {
    pages_case(
        "pages_body_table_formula_refresh_clear_parity",
        Tier::B,
        Evidence::Queued,
        2,
        Value::Empty,
    )
}

#[inline(never)]
fn host_refresh_keynote(source: &[u8]) -> TestResult<Vec<u8>> {
    let mut editor = keynote::Draft::from_bytes(source)?;
    let model = keynote_model(&editor)?;
    editor.set_slide_table_cell(0, model, 1, 1, Value::Number(FiniteF64::new(77.0)?))?;
    Ok(editor.to_bytes()?)
}

#[inline(never)]
fn leaf_refresh_keynote(source: &[u8], drawable: usize) -> TestResult<Vec<u8>> {
    let package = litchi_keynote::Package::from_bytes(source)?;
    let mut edit = package.edit_slide_drawables();
    edit.set_table_cell(
        0,
        litchi_keynote::DrawableSelector::index(drawable),
        1,
        1,
        litchi_keynote::slide::drawable::TableCellValue::Number(77.0),
    )?;
    Ok(leaf_bytes!(edit.commit()?.package()))
}

#[inline(never)]
fn host_refresh_pages(source: &[u8]) -> TestResult<Vec<u8>> {
    let mut editor = pages::Draft::from_bytes(source)?;
    let model = pages_model(&editor)?;
    editor.set_table_cell(model, 1, 1, Value::Number(FiniteF64::new(41.0)?))?;
    Ok(editor.to_bytes()?)
}

#[inline(never)]
fn leaf_refresh_pages(source: &[u8]) -> TestResult<Vec<u8>> {
    let package = litchi_pages::Package::from_bytes(source)?;
    let mut edit = package.edit_body_table_values(0usize)?;
    edit.set(1, 1, litchi_pages::table::values::Value::Number(41.0))?;
    Ok(leaf_bytes!(edit.commit()?.package()))
}

/// Callgrind ABBA probe of the two refresh routes (record 0376 §5):
///
/// ```sh
/// CARGO_PROFILE_RELEASE_DEBUG=1 cargo test --release -p litchi-iwa --test parity --no-run
/// taskset -c 3 valgrind --tool=callgrind '--toggle-collect=*formula_refresh_cases::host_refresh_keynote*' \
///   target/release/deps/parity-<hash> --ignored --exact \
///   formula_refresh_cases::formula_refresh_cost_abba --test-threads=1
/// # likewise leaf_refresh_keynote, host_refresh_pages, leaf_refresh_pages
/// ```
#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn formula_refresh_cost_abba() -> TestResult {
    use std::hint::black_box;
    let keynote_input = fixture(KEYNOTE)?;
    let pages_input = fixture(PAGES)?;
    let drawable = (0..64)
        .find(|drawable| leaf_refresh_keynote(&keynote_input.bytes, *drawable).is_ok())
        .ok_or("no table drawable")?;
    for leg in [false, true, true, false] {
        if leg {
            black_box(leaf_refresh_keynote(&keynote_input.bytes, drawable)?);
            black_box(leaf_refresh_pages(&pages_input.bytes)?);
        } else {
            black_box(host_refresh_keynote(&keynote_input.bytes)?);
            black_box(host_refresh_pages(&pages_input.bytes)?);
        }
    }
    Ok(())
}
