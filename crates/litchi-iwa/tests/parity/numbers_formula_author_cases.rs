//! Numbers formula authoring (record 0376c): the host editor's
//! `set_formula` / `set_formula_with_cached_value` against the leaf's
//! `litchi_numbers::table::cells::Edit::{set_formula, set_formula_cached}`.

use std::hint::black_box;

use litchi_iwa::numbers;
use litchi_iwa_common::formula::{
    FormulaBinaryOperator, FormulaCachedValue, FormulaCellReference, FormulaExpression,
};
use litchi_iwa_protos::tst;
use litchi_numbers::formula::{BinaryOperator, CachedValue, CellReference, Expression};
use litchi_numbers::table::{At, cells::Storage};
use prost::Message;

use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier, fixture};
use crate::known_fields::{Decoded, decoded};
use crate::leaf_bytes;

/// The formula, `B3 + addend` (zero-based row 2, column 1).
#[derive(Clone, Copy, Debug)]
struct Formula {
    addend: f64,
    cached: Option<f64>,
}

fn host_expression(formula: Formula) -> FormulaExpression {
    FormulaExpression::binary(
        FormulaBinaryOperator::Add,
        FormulaExpression::Cell(FormulaCellReference::relative(2, 1)),
        FormulaExpression::Number(formula.addend),
    )
}

fn leaf_expression(formula: Formula) -> TestResult<Expression> {
    Ok(Expression::binary(
        BinaryOperator::Add,
        Expression::cell(CellReference::relative(2, 1)),
        Expression::number(formula.addend)?,
    )?)
}

#[inline(never)]
fn host_route(source: &[u8], (r, c): (usize, usize), formula: Formula) -> TestResult<Vec<u8>> {
    let mut editor = numbers::Draft::from_bytes(source)?;
    let table = crate::numbers_cases::host_first_table_model_id(&editor)?;
    match formula.cached {
        None => editor.set_formula(table, r, c, host_expression(formula))?,
        Some(value) => editor.set_formula_with_cached_value(
            table,
            r,
            c,
            host_expression(formula),
            FormulaCachedValue::number(value)?,
        )?,
    }
    Ok(editor.to_bytes()?)
}

/// The no-cache host setter and the leaf's evaluator intentionally disagree
/// only in the display cache: the host writes zero until Numbers recalculates,
/// while the leaf publishes the evaluated scalar. Normalize only the raw BNC
/// storage bytes in `TST.Tile`; the semantic cell reads below still compare the
/// authored formula and every selected value.
fn decode_tile_without_cell_storage(data: &[u8]) -> Result<Decoded, prost::DecodeError> {
    let mut tile = tst::Tile::decode(data)?;
    for row in &mut tile.row_infos {
        row.cell_storage_buffer_pre_bnc.clear();
        row.cell_offsets_pre_bnc.clear();
        row.cell_storage_buffer = None;
        row.cell_offsets = None;
    }
    Ok(decoded(tile, data))
}

fn cell_value_projection(bytes: &[u8], (row, column): (usize, usize)) -> TestResult<String> {
    let package = litchi_numbers::Package::from_bytes(bytes)?;
    let state = package.table_cell(
        0usize,
        0usize,
        At::new(u32::try_from(row)?, u32::try_from(column)?),
    )?;
    Ok(match state.storage() {
        Storage::Stored(litchi_numbers::cell::Value::Formula(formula)) => {
            format!("formula:{formula}")
        },
        Storage::Stored(value) => format!("value:{value:?}"),
        Storage::Missing => "missing".to_owned(),
        _ => "unknown".to_owned(),
    })
}

fn table_semantic_projection(bytes: &[u8]) -> TestResult<Vec<String>> {
    let package = litchi_numbers::Package::from_bytes(bytes)?;
    let mut cells = Vec::new();
    for sheet in package.sheets() {
        for table in sheet.tables() {
            for cell in table.iter_cells() {
                cells.push(format!("{cell:?}"));
            }
        }
    }
    Ok(cells)
}

fn cell_cache_projection(bytes: &[u8], (row, column): (usize, usize)) -> TestResult<String> {
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
                let Some(cell) = table_write::tile_cell(
                    &message.data,
                    u32::try_from(row)?,
                    u32::try_from(column)?,
                    &Limits::DEFAULT,
                )?
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

fn leaf_route(source: &[u8], (r, c): (usize, usize), formula: Formula) -> TestResult<LeafRun> {
    let package = litchi_numbers::Package::from_bytes(source)?;
    let at = At::new(u32::try_from(r)?, u32::try_from(c)?);
    let edit = package.edit_table_cells(0usize, 0usize)?;
    let edit = match formula.cached {
        None => edit.set_formula(at, leaf_expression(formula)?)?,
        Some(value) => {
            edit.set_formula_cached(at, leaf_expression(formula)?, CachedValue::number(value)?)?
        },
    };
    let commit = edit.commit()?;
    let restored = commit
        .package()
        .apply_table_cells(&commit.patch().inverse())?;
    Ok(LeafRun::with_inverse(
        leaf_bytes!(commit.package()),
        leaf_bytes!(restored.package()),
    ))
}

#[inline(never)]
fn leaf_cost_route(source: &[u8], (r, c): (usize, usize), formula: Formula) -> TestResult<Vec<u8>> {
    let package = litchi_numbers::Package::from_bytes(source)?;
    let at = At::new(u32::try_from(r)?, u32::try_from(c)?);
    let edit = package.edit_table_cells(0usize, 0usize)?;
    let edit = match formula.cached {
        None => edit.set_formula(at, leaf_expression(formula)?)?,
        Some(value) => {
            edit.set_formula_cached(at, leaf_expression(formula)?, CachedValue::number(value)?)?
        },
    };
    let commit = edit.commit()?;
    Ok(leaf_bytes!(commit.package()))
}

fn run(name: &'static str, input: &Fixture, cell: (usize, usize), formula: Formula) -> TestResult {
    let mut case = Case::new(name, Format::Numbers, Tier::B, Evidence::Queued, input)
        .leaf_structural_metadata()
        .leaf_invalidates_root_previews()
        .host(|source| host_route(source, cell, formula))
        .leaf(|source| leaf_route(source, cell, formula))
        .read("formula/value C3", move |bytes| {
            cell_value_projection(bytes, cell)
        })
        .read("precedent/value B3", |bytes| {
            cell_value_projection(bytes, (2, 1))
        })
        .read("all semantic table cells", table_semantic_projection);
    if formula.cached.is_none() {
        case = case.known_fields(|known| {
            known.with(
                6_002,
                "TST.Tile (no-cache BNC policy normalized)",
                decode_tile_without_cell_storage,
            )
        });
    } else {
        case = case.read("display cache C3", move |bytes| {
            cell_cache_projection(bytes, cell)
        });
    }
    case.run()?;
    Ok(())
}

const AUTHOR: Formula = Formula {
    addend: 1.0,
    cached: None,
};

const CACHED: Formula = Formula {
    addend: 1.0,
    cached: Some(43.0),
};

const GRAPH_CACHED: Formula = Formula {
    addend: 1.0,
    cached: Some(2.0),
};

/// A formula in a table whose owner has no dependency tile at the origin
/// (the leaf creates the tile, as the host does).
#[test]
fn numbers_table_formula_author_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    let input = fixture("iwork/numbers/basic.numbers")?;
    run(
        "numbers_table_formula_author_parity",
        &input,
        (2, 2),
        AUTHOR,
    )
}

/// A formula beside a native formula graph.
#[test]
fn numbers_table_formula_author_graph_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    let input = fixture("iwork/numbers/formula-budget-native.numbers")?;
    run(
        "numbers_table_formula_author_graph_parity",
        &input,
        (2, 2),
        GRAPH_CACHED,
    )
}

/// A formula with its display cache.
#[test]
fn numbers_table_formula_cached_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    let input = fixture("iwork/numbers/basic.numbers")?;
    run(
        "numbers_table_formula_cached_parity",
        &input,
        (2, 2),
        CACHED,
    )
}

/// The direct host `set_formula` route is intentionally checked separately
/// because its zero display cache differs from the leaf's evaluated no-cache
/// publication. Formula semantics and the documented cache policy are both
/// asserted from reopened outputs; the normal Tier-B case above normalizes only
/// this one cache-policy difference for structural comparison.
#[test]
fn numbers_table_formula_author_host_set_formula_policy() -> TestResult {
    let input = fixture("iwork/numbers/basic.numbers")?;
    let host = host_route(&input.bytes, (2, 2), AUTHOR)?;
    let leaf = leaf_route(&input.bytes, (2, 2), AUTHOR)?;
    Format::Numbers.reopen(&host)?;
    Format::Numbers.reopen(&leaf.output)?;
    if host == input.bytes {
        return Err("host set_formula produced a no-op".into());
    }
    if leaf.restored.as_deref() != Some(input.bytes.as_slice()) {
        return Err("leaf set_formula inverse did not restore the source".into());
    }
    let host_formula = cell_value_projection(&host, (2, 2))?;
    let leaf_formula = cell_value_projection(&leaf.output, (2, 2))?;
    if host_formula != leaf_formula {
        return Err(format!(
            "host and leaf formula reads differ: host {host_formula:?}, leaf {leaf_formula:?}"
        )
        .into());
    }
    let host_cache = cell_cache_projection(&host, (2, 2))?;
    let leaf_cache = cell_cache_projection(&leaf.output, (2, 2))?;
    if host_cache != "Some(Number(FiniteF64(0.0)))" {
        return Err(format!("host set_formula cache policy changed: {host_cache:?}").into());
    }
    if leaf_cache != "Some(Number(FiniteF64(43.0)))" {
        return Err(format!("leaf set_formula cache policy changed: {leaf_cache:?}").into());
    }
    Ok(())
}

/// Release-mode host-vs-leaf probe for source-built no-cache authoring. Run
/// under pinned-CPU callgrind with `host_route` and `leaf_cost_route` toggles;
/// it deliberately excludes the parity harness' inverse patch readback.
#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn numbers_table_formula_author_cost_abba() -> TestResult {
    let input = fixture("iwork/numbers/basic.numbers")?;
    for route in [0_u8, 1, 1, 0] {
        if route == 0 {
            black_box(host_route(&input.bytes, (2, 2), AUTHOR)?);
        } else {
            black_box(leaf_cost_route(&input.bytes, (2, 2), AUTHOR)?);
        }
    }
    Ok(())
}

/// Release-mode host-vs-leaf probe for explicit cached authoring beside a
/// native formula graph.
#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn numbers_table_formula_author_graph_cost_abba() -> TestResult {
    let input = fixture("iwork/numbers/formula-budget-native.numbers")?;
    for route in [0_u8, 1, 1, 0] {
        if route == 0 {
            black_box(host_route(&input.bytes, (2, 2), GRAPH_CACHED)?);
        } else {
            black_box(leaf_cost_route(&input.bytes, (2, 2), GRAPH_CACHED)?);
        }
    }
    Ok(())
}

/// Release-mode host-vs-leaf probe for explicit cached authoring.
#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn numbers_table_formula_cached_cost_abba() -> TestResult {
    let input = fixture("iwork/numbers/basic.numbers")?;
    for route in [0_u8, 1, 1, 0] {
        if route == 0 {
            black_box(host_route(&input.bytes, (2, 2), CACHED)?);
        } else {
            black_box(leaf_cost_route(&input.bytes, (2, 2), CACHED)?);
        }
    }
    Ok(())
}
