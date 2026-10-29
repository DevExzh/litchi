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
fn leaf_refresh_keynote_impl(source: &[u8], drawable: usize) -> TestResult<Vec<u8>> {
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

/// Counted leaf wrapper. Fixture discovery must call the implementation
/// directly so setup does not contaminate the selected callgrind symbol.
#[inline(never)]
fn leaf_refresh_keynote(source: &[u8], drawable: usize) -> TestResult<Vec<u8>> {
    leaf_refresh_keynote_impl(source, drawable)
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
        .find(|drawable| leaf_refresh_keynote_impl(&keynote_input.bytes, *drawable).is_ok())
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

/// A batch of slide-table writes in one edit: C3, then B2 (read by D2), the
/// second planned over the first's staged state, against the host's
/// one-pass `set_slide_table_cells`.
#[test]
fn keynote_slide_table_formula_refresh_batch_parity() -> TestResult {
    use litchi_keynote::slide::drawable::TableCellValue;
    let input = fixture(KEYNOTE)?;
    Case::new(
        "keynote_slide_table_formula_refresh_batch_parity",
        Format::Keynote,
        Tier::B,
        Evidence::Queued,
        &input,
    )
    .host(|source| {
        let mut editor = keynote::Draft::from_bytes(source)?;
        let model = keynote_model(&editor)?;
        editor.set_slide_table_cells(
            0,
            model,
            [
                keynote::TableCellUpdate::new(
                    1,
                    1,
                    keynote::TableCellValue::Number(30.0.try_into()?),
                ),
                keynote::TableCellUpdate::new(2, 2, keynote::TableCellValue::Boolean(false)),
            ],
        )?;
        Ok(editor.to_bytes()?)
    })
    .leaf(|source| {
        let package = litchi_keynote::Package::from_bytes(source)?;
        for drawable in 0..64 {
            let mut edit = package.edit_slide_drawables();
            let selector = litchi_keynote::DrawableSelector::index(drawable);
            if edit
                .set_table_cell(0, selector, 2, 2, TableCellValue::Boolean(false))
                .is_err()
            {
                continue;
            }
            edit.set_table_cell(0, selector, 1, 1, TableCellValue::Number(30.0))?;
            let output = leaf_bytes!(edit.commit()?.package());
            if formula_caches(&output, 1, 3)? == formula_caches(source, 1, 3)? {
                return Err("the leaf left D2's cached result unchanged".into());
            }
            return Ok(LeafRun::output(output));
        }
        Err("no table drawable".into())
    })
    .leaf_invalidates_root_previews()
    .leaf_structural_metadata()
    .read("cached result of D2", |bytes| formula_caches(bytes, 1, 3))
    .read("host cell B2", |bytes| keynote_cell(bytes, 1, 1))
    .read("host cell C3", |bytes| keynote_cell(bytes, 2, 2))
    .run()?;
    Ok(())
}

/// A write that a formula in another table reads: on
/// `body-table-cross-reference-native.pages`, table 2 A3 = Table 1::B2 + 1
/// and table 1 B4 = SUM(B2:B3); writing table 1 B2 refreshes both, as the
/// host does.
#[test]
fn pages_body_table_formula_refresh_cross_table_parity() -> TestResult {
    let input = fixture("iwork/pages/body-table-cross-reference-native.pages")?;
    Case::new(
        "pages_body_table_formula_refresh_cross_table_parity",
        Format::Pages,
        Tier::B,
        Evidence::Queued,
        &input,
    )
    .host(|source| {
        let mut editor = pages::Draft::from_bytes(source)?;
        let model = pages_model(&editor)?;
        editor.set_table_cell(model, 1, 1, Value::Number(FiniteF64::new(41.0)?))?;
        Ok(editor.to_bytes()?)
    })
    .leaf(|source| {
        let package = litchi_pages::Package::from_bytes(source)?;
        let mut edit = package.edit_body_table_values(0usize)?;
        edit.set(1, 1, litchi_pages::table::values::Value::Number(41.0))?;
        let output = leaf_bytes!(edit.commit()?.package());
        for (row, column) in [(3, 1), (2, 0)] {
            if formula_caches(&output, row, column)? == formula_caches(source, row, column)? {
                return Err(format!("the leaf left ({row}, {column})'s cache unchanged").into());
            }
        }
        Ok(LeafRun::output(output))
    })
    .read("cached result of table 1 B4", |bytes| {
        formula_caches(bytes, 3, 1)
    })
    .read("cached result of table 2 A3", |bytes| {
        formula_caches(bytes, 2, 0)
    })
    .run()?;
    Ok(())
}

/// `C3 = B3 + 1` cached as `cached`.
fn authored_expression() -> litchi_iwa_common::formula::FormulaExpression {
    use litchi_iwa_common::formula::{
        FormulaBinaryOperator, FormulaCellReference, FormulaExpression,
    };
    FormulaExpression::binary(
        FormulaBinaryOperator::Add,
        FormulaExpression::Cell(FormulaCellReference::relative(2, 1)),
        FormulaExpression::Number(1.0),
    )
}

fn keynote_first_table(editor: &keynote::Draft) -> TestResult<(usize, u64)> {
    for slide in 0..editor.slides()?.len() {
        if let Some(table) = editor.slide_tables(slide)?.first() {
            return Ok((slide, table.model_object_id));
        }
    }
    Err("no slide table".into())
}

fn host_keynote_author(source: &[u8]) -> TestResult<Vec<u8>> {
    let mut editor = keynote::Draft::from_bytes(source)?;
    let (slide, model) = keynote_first_table(&editor)?;
    editor.set_slide_table_formula(
        slide,
        model,
        2,
        2,
        authored_expression(),
        litchi_iwa_common::formula::FormulaCachedValue::number(8.0)?,
    )?;
    Ok(editor.to_bytes()?)
}

/// The leaf route on the first slide table: `stage` stages the operation.
fn leaf_keynote(
    source: &[u8],
    stage: &dyn Fn(
        &mut litchi_keynote::slide::drawable::Edit<'_>,
        usize,
        litchi_keynote::DrawableSelector,
    ) -> Result<(), litchi_keynote::slide::drawable::Error>,
) -> TestResult<Vec<u8>> {
    let package = litchi_keynote::Package::from_bytes(source)?;
    let host = keynote::Draft::from_bytes(source)?;
    let (slide, _) = keynote_first_table(&host)?;
    let mut last = None;
    for drawable in 0..64 {
        let mut edit = package.edit_slide_drawables();
        let selector = litchi_keynote::DrawableSelector::index(drawable);
        if let Err(error) = stage(&mut edit, slide, selector) {
            last = Some(error.to_string());
            continue;
        }
        return Ok(leaf_bytes!(edit.commit()?.package()));
    }
    Err(last
        .unwrap_or_else(|| "no table drawable".to_owned())
        .into())
}

fn leaf_keynote_author(source: &[u8]) -> TestResult<Vec<u8>> {
    leaf_keynote(source, &|edit, slide, selector| {
        edit.set_table_formula_cached(
            slide,
            selector,
            2,
            2,
            authored_expression(),
            litchi_iwa_common::formula::FormulaCachedValue::number(8.0).expect("finite"),
        )
        .map(|_| ())
    })
}

fn keynote_formula(bytes: &[u8]) -> TestResult<Option<String>> {
    let editor = keynote::Draft::from_bytes(bytes)?;
    let (slide, model) = keynote_first_table(&editor)?;
    Ok(editor.slide_table_formula(slide, model, 2, 2)?)
}

fn keynote_author_case(
    name: &'static str,
    path: &'static str,
    tier: Tier,
    evidence: Evidence,
) -> TestResult {
    let input = fixture(path)?;
    Case::new(name, Format::Keynote, tier, evidence, &input)
        .host(host_keynote_author)
        .leaf(|source| Ok(LeafRun::output(leaf_keynote_author(source)?)))
        .leaf_invalidates_root_previews()
        .leaf_structural_metadata()
        .read("host slide_table_formula C3", keynote_formula)
        .read("cached result of C3", |bytes| formula_caches(bytes, 2, 2))
        .run()?;
    Ok(())
}

#[test]
fn keynote_slide_table_formula_author_parity() -> TestResult {
    keynote_author_case(
        "keynote_slide_table_formula_author_parity",
        "iwork/keynote/table-discovery.key",
        Tier::B,
        Evidence::Queued,
    )
}

#[test]
fn keynote_slide_table_formula_author_graph_parity() -> TestResult {
    keynote_author_case(
        "keynote_slide_table_formula_author_graph_parity",
        KEYNOTE,
        Tier::B,
        Evidence::Queued,
    )
}

/// Clear a host-authored formula: the host's `clear_slide_table_formula`
/// against the leaf's `set_table_cell(.., Empty)`.
#[test]
fn keynote_slide_table_formula_clear_parity() -> TestResult {
    let authored = host_keynote_author(&fixture("iwork/keynote/table-discovery.key")?.bytes)?;
    let input = crate::harness::Fixture {
        path: "iwork/keynote/table-discovery.key",
        bytes: authored,
    };
    Case::new(
        "keynote_slide_table_formula_clear_parity",
        Format::Keynote,
        Tier::B,
        Evidence::Queued,
        &input,
    )
    .host(|source| {
        let mut editor = keynote::Draft::from_bytes(source)?;
        let (slide, model) = keynote_first_table(&editor)?;
        editor.clear_slide_table_formula(slide, model, 2, 2)?;
        Ok(editor.to_bytes()?)
    })
    .leaf(|source| {
        Ok(LeafRun::output(leaf_keynote(
            source,
            &|edit, slide, selector| {
                edit.set_table_cell(
                    slide,
                    selector,
                    2,
                    2,
                    litchi_keynote::slide::drawable::TableCellValue::Empty,
                )
                .map(|_| ())
            },
        )?))
    })
    .leaf_invalidates_root_previews()
    .leaf_structural_metadata()
    .read("host slide_table_formula C3", keynote_formula)
    .run()?;
    Ok(())
}

fn host_authored_keynote() -> TestResult<crate::harness::Fixture> {
    Ok(crate::harness::Fixture {
        path: "iwork/keynote/table-discovery.key",
        bytes: host_keynote_author(&fixture("iwork/keynote/table-discovery.key")?.bytes)?,
    })
}

/// Replace a host-authored formula with a number: the host's
/// `set_slide_table_cell` against the leaf's `set_table_cell`.
#[test]
fn keynote_slide_table_formula_replace_parity() -> TestResult {
    let input = host_authored_keynote()?;
    Case::new(
        "keynote_slide_table_formula_replace_parity",
        Format::Keynote,
        Tier::B,
        Evidence::Queued,
        &input,
    )
    .host(|source| {
        let mut editor = keynote::Draft::from_bytes(source)?;
        let (slide, model) = keynote_first_table(&editor)?;
        editor.set_slide_table_cell(slide, model, 2, 2, Value::Number(FiniteF64::new(5.0)?))?;
        Ok(editor.to_bytes()?)
    })
    .leaf(|source| {
        Ok(LeafRun::output(leaf_keynote(
            source,
            &|edit, slide, selector| {
                edit.set_table_cell(
                    slide,
                    selector,
                    2,
                    2,
                    litchi_keynote::slide::drawable::TableCellValue::Number(5.0),
                )
                .map(|_| ())
            },
        )?))
    })
    .leaf_invalidates_root_previews()
    .leaf_structural_metadata()
    .read("host slide_table_formula C3", keynote_formula)
    .read("host cell C3", |bytes| {
        let editor = keynote::Draft::from_bytes(bytes)?;
        let (slide, model) = keynote_first_table(&editor)?;
        Ok(format!(
            "{:?}",
            editor.slide_table(slide, model)?.get_cell(2, 2)
        ))
    })
    .run()?;
    Ok(())
}

/// The leaf reads a formula cell as its rendered source, as the host's
/// `slide_table_formula` does: the native D2 (receipted) and a
/// host-authored formula.
#[test]
fn keynote_slide_table_formula_read_parity() -> TestResult {
    let evidence = Evidence::Native("test-data/iwork/keynote/slide-table-read-native-receipt.json");
    for (input, row, column) in [(fixture(KEYNOTE)?, 1, 3), (host_authored_keynote()?, 2, 2)] {
        let host = {
            let editor = keynote::Draft::from_bytes(&input.bytes)?;
            let (slide, model) = keynote_first_table(&editor)?;
            editor
                .slide_table_formula(slide, model, row, column)?
                .ok_or("host reads no formula")?
        };
        Case::new(
            "keynote_slide_table_formula_read_parity",
            Format::Keynote,
            Tier::A,
            evidence,
            &input,
        )
        .read_only()
        .host(|source| Ok(keynote::Draft::from_bytes(source)?.to_bytes()?))
        .leaf(|source| {
            let package = litchi_keynote::Package::from_bytes(source)?;
            let host_editor = keynote::Draft::from_bytes(source)?;
            let (slide, _) = keynote_first_table(&host_editor)?;
            let read = package.slide_table_cells(slide, 0_usize)?;
            let cell = read
                .table()
                .get(litchi_iwa_common::table::coordinate::At::new(
                    u32::try_from(row)?,
                    u32::try_from(column)?,
                ));
            if cell
                != Some(&litchi_iwa_common::table::cell::value::Value::Formula(
                    host.clone(),
                ))
            {
                return Err(format!("leaf reads {cell:?}, host {host:?}").into());
            }
            Ok(LeafRun::output(leaf_bytes!(package)))
        })
        .run()?;
    }
    Ok(())
}

/// The native D2 formula of `slide-table-read-native.key` carries a cell
/// comment: the host refuses to clear it ("expected to remove one protobuf
/// field 11") and the leaf refuses typed, because the shared cell writer
/// does not rewrite a commented cell (record 0376a).
#[test]
fn keynote_native_commented_formula_clear_is_refused() -> TestResult {
    let input = fixture(KEYNOTE)?;
    let host = (|| -> TestResult<Vec<u8>> {
        let mut editor = keynote::Draft::from_bytes(&input.bytes)?;
        let (slide, model) = keynote_first_table(&editor)?;
        editor.clear_slide_table_formula(slide, model, 1, 3)?;
        Ok(editor.to_bytes()?)
    })();
    assert!(host.is_err(), "the host now clears the native formula");
    let leaf = leaf_keynote(&input.bytes, &|edit, slide, selector| {
        edit.set_table_cell(
            slide,
            selector,
            1,
            3,
            litchi_keynote::slide::drawable::TableCellValue::Empty,
        )
        .map(|_| ())
    });
    let error = leaf.expect_err("a commented formula cell is refused");
    assert!(error.to_string().contains("cell storage"), "{error}");
    Ok(())
}

fn host_pages_author(source: &[u8]) -> TestResult<Vec<u8>> {
    let mut editor = pages::Draft::from_bytes(source)?;
    let model = pages_model(&editor)?;
    editor.set_table_formula(
        model,
        2,
        2,
        authored_expression(),
        litchi_iwa_common::formula::FormulaCachedValue::number(8.0)?,
    )?;
    Ok(editor.to_bytes()?)
}

fn leaf_pages(
    source: &[u8],
    stage: &dyn Fn(
        &mut litchi_pages::table::values::Edit<'_>,
    ) -> Result<(), litchi_pages::table::values::Error>,
) -> TestResult<Vec<u8>> {
    let package = litchi_pages::Package::from_bytes(source)?;
    let mut edit = package.edit_body_table_values(0usize)?;
    stage(&mut edit)?;
    Ok(leaf_bytes!(edit.commit()?.package()))
}

fn leaf_pages_author(source: &[u8]) -> TestResult<Vec<u8>> {
    leaf_pages(source, &|edit| {
        edit.set_formula_cached(
            2,
            2,
            authored_expression(),
            litchi_iwa_common::formula::FormulaCachedValue::number(8.0).expect("finite"),
        )
        .map(|_| ())
    })
}

fn pages_formula(bytes: &[u8]) -> TestResult<Option<String>> {
    let editor = pages::Draft::from_bytes(bytes)?;
    Ok(editor.table_formula(pages_model(&editor)?, 2, 2)?)
}

fn pages_author_case(
    name: &'static str,
    path: &'static str,
    tier: Tier,
    evidence: Evidence,
) -> TestResult {
    let input = fixture(path)?;
    Case::new(name, Format::Pages, tier, evidence, &input)
        .host(host_pages_author)
        .leaf(|source| Ok(LeafRun::output(leaf_pages_author(source)?)))
        .leaf_structural_metadata()
        .read("host table_formula C3", pages_formula)
        .read("cached result of C3", |bytes| formula_caches(bytes, 2, 2))
        .run()?;
    Ok(())
}

#[test]
fn pages_body_table_formula_author_parity() -> TestResult {
    pages_author_case(
        "pages_body_table_formula_author_parity",
        "iwork/pages/body-table-catalog-native.pages",
        Tier::B,
        Evidence::Queued,
    )
}

/// `body-table-cells-native.pages`: the table's `SUM(B2:B3)` range is
/// mirrored inline and in a range-precedents tile, which the shared rewrite
/// regenerates together (record 0376a).
#[test]
fn pages_body_table_formula_author_graph_parity() -> TestResult {
    pages_author_case(
        "pages_body_table_formula_author_graph_parity",
        PAGES,
        Tier::B,
        Evidence::Queued,
    )
}

fn host_authored_pages() -> TestResult<crate::harness::Fixture> {
    Ok(crate::harness::Fixture {
        path: "iwork/pages/body-table-catalog-native.pages",
        bytes: host_pages_author(&fixture("iwork/pages/body-table-catalog-native.pages")?.bytes)?,
    })
}

/// Clear a host-authored formula: the host's `clear_table_formula` against
/// the leaf's `set(.., Empty)`.
#[test]
fn pages_body_table_formula_clear_parity() -> TestResult {
    let input = host_authored_pages()?;
    Case::new(
        "pages_body_table_formula_clear_parity",
        Format::Pages,
        Tier::B,
        Evidence::Queued,
        &input,
    )
    .host(|source| {
        let mut editor = pages::Draft::from_bytes(source)?;
        let model = pages_model(&editor)?;
        editor.clear_table_formula(model, 2, 2)?;
        Ok(editor.to_bytes()?)
    })
    .leaf(|source| {
        Ok(LeafRun::output(leaf_pages(source, &|edit| {
            edit.set(2, 2, litchi_pages::table::values::Value::Empty)
                .map(|_| ())
        })?))
    })
    .leaf_structural_metadata()
    .read("host table_formula C3", pages_formula)
    .run()?;
    Ok(())
}

/// Replace a host-authored formula with a number.
#[test]
fn pages_body_table_formula_replace_parity() -> TestResult {
    let input = host_authored_pages()?;
    Case::new(
        "pages_body_table_formula_replace_parity",
        Format::Pages,
        Tier::B,
        Evidence::Queued,
        &input,
    )
    .host(|source| {
        let mut editor = pages::Draft::from_bytes(source)?;
        let model = pages_model(&editor)?;
        editor.set_table_cell(model, 2, 2, Value::Number(FiniteF64::new(5.0)?))?;
        Ok(editor.to_bytes()?)
    })
    .leaf(|source| {
        Ok(LeafRun::output(leaf_pages(source, &|edit| {
            edit.set(2, 2, litchi_pages::table::values::Value::Number(5.0))
                .map(|_| ())
        })?))
    })
    .leaf_structural_metadata()
    .read("host table_formula C3", pages_formula)
    .read("host cell C3", |bytes| pages_cell(bytes, 2, 2))
    .run()?;
    Ok(())
}

/// The leaf reads a formula cell as its rendered source, as the host's
/// `table_formula` does: the native B4 (receipted) and a host-authored
/// formula.
#[test]
fn pages_body_table_formula_read_parity() -> TestResult {
    let evidence = Evidence::Native("test-data/iwork/pages/body-table-cells-native-receipt.json");
    for (input, row, column) in [(fixture(PAGES)?, 3, 1), (host_authored_pages()?, 2, 2)] {
        let host = {
            let editor = pages::Draft::from_bytes(&input.bytes)?;
            editor
                .table_formula(pages_model(&editor)?, row, column)?
                .ok_or("host reads no formula")?
        };
        Case::new(
            "pages_body_table_formula_read_parity",
            Format::Pages,
            Tier::A,
            evidence,
            &input,
        )
        .read_only()
        .host(|source| Ok(pages::Draft::from_bytes(source)?.to_bytes()?))
        .leaf(|source| {
            let package = litchi_pages::Package::from_bytes(source)?;
            let read = package.body_table_cells(litchi_pages::BodyTableSelector::index(0))?;
            let cell = read
                .table()
                .get(litchi_iwa_common::table::coordinate::At::new(
                    u32::try_from(row)?,
                    u32::try_from(column)?,
                ));
            if cell
                != Some(&litchi_iwa_common::table::cell::value::Value::Formula(
                    host.clone(),
                ))
            {
                return Err(format!("leaf reads {cell:?}, host {host:?}").into());
            }
            Ok(LeafRun::output(leaf_bytes!(package)))
        })
        .run()?;
    }
    Ok(())
}

#[inline(never)]
fn host_author_keynote(source: &[u8]) -> TestResult<Vec<u8>> {
    host_keynote_author(source)
}

#[inline(never)]
fn leaf_author_keynote(source: &[u8]) -> TestResult<Vec<u8>> {
    leaf_keynote_author(source)
}

#[inline(never)]
fn host_author_pages(source: &[u8]) -> TestResult<Vec<u8>> {
    host_pages_author(source)
}

#[inline(never)]
fn leaf_author_pages(source: &[u8]) -> TestResult<Vec<u8>> {
    leaf_pages_author(source)
}

/// Callgrind ABBA probe of formula authoring (record 0376a §5): C3 = B3 + 1
/// on `table-discovery.key` and `body-table-catalog-native.pages`, open,
/// edit, commit and serialize; toggle `*formula_refresh_cases::host_author_*`
/// or `*formula_refresh_cases::leaf_author_*`.
#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn formula_author_cost_abba() -> TestResult {
    use std::hint::black_box;
    let keynote_input = fixture("iwork/keynote/table-discovery.key")?;
    let pages_input = fixture("iwork/pages/body-table-catalog-native.pages")?;
    for leg in [false, true, true, false] {
        if leg {
            black_box(leaf_author_keynote(&keynote_input.bytes)?);
            black_box(leaf_author_pages(&pages_input.bytes)?);
        } else {
            black_box(host_author_keynote(&keynote_input.bytes)?);
            black_box(host_author_pages(&pages_input.bytes)?);
        }
    }
    Ok(())
}
