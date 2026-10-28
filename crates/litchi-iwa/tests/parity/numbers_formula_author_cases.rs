//! Numbers formula authoring (record 0376c): the host editor's
//! `set_formula` / `set_formula_with_cached_value` against the leaf's
//! `litchi_numbers::table::cells::Edit::{set_formula, set_formula_cached}`.

use std::hint::black_box;

use litchi_iwa::numbers;
use litchi_iwa_common::formula::{
    FormulaBinaryOperator, FormulaCachedValue, FormulaCellReference, FormulaExpression,
};
use litchi_numbers::formula::{BinaryOperator, CachedValue, CellReference, Expression};
use litchi_numbers::table::At;

use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier, fixture};
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
fn host_route(
    source: &[u8],
    (r, c): (usize, usize),
    formula: Formula,
    evaluated_cache: f64,
) -> TestResult<Vec<u8>> {
    let mut editor = numbers::Draft::from_bytes(source)?;
    let table = crate::numbers_cases::host_first_table_model_id(&editor)?;
    match formula.cached {
        // The host's no-cache setter leaves a zero display cache, while the
        // leaf evaluates the authored scalar and stores that value.  The
        // differential route supplies the evaluated cache so the parity case
        // compares the formula authoring payloads and the shared key/graph
        // transitions rather than that documented cache-policy difference.
        None => editor.set_formula_with_cached_value(
            table,
            r,
            c,
            host_expression(formula),
            FormulaCachedValue::number(evaluated_cache)?,
        )?,
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

fn run(
    name: &'static str,
    input: &Fixture,
    cell: (usize, usize),
    formula: Formula,
    evaluated_cache: f64,
) -> TestResult {
    Case::new(name, Format::Numbers, Tier::B, Evidence::Queued, input)
        .leaf_structural_metadata()
        .leaf_invalidates_root_previews()
        .host(|source| host_route(source, cell, formula, evaluated_cache))
        .leaf(|source| leaf_route(source, cell, formula))
        .run()?;
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

/// A formula in a table whose owner has no dependency tile at the origin
/// (the leaf creates the tile, as the host does).
#[test]
fn numbers_table_formula_author_parity() -> TestResult {
    let input = fixture("iwork/numbers/basic.numbers")?;
    run(
        "numbers_table_formula_author_parity",
        &input,
        (2, 2),
        AUTHOR,
        43.0,
    )
}

/// A formula beside a native formula graph.
#[test]
fn numbers_table_formula_author_graph_parity() -> TestResult {
    let input = fixture("iwork/numbers/formula-budget-native.numbers")?;
    run(
        "numbers_table_formula_author_graph_parity",
        &input,
        (2, 2),
        AUTHOR,
        2.0,
    )
}

/// A formula with its display cache.
#[test]
fn numbers_table_formula_cached_parity() -> TestResult {
    let input = fixture("iwork/numbers/basic.numbers")?;
    run(
        "numbers_table_formula_cached_parity",
        &input,
        (2, 2),
        CACHED,
        43.0,
    )
}

#[test]
#[ignore]
fn scratch_header_probe() -> TestResult {
    use litchi_iwa_protos::tst;
    use prost::Message;
    let source = fixture("iwork/numbers/basic.numbers")?;
    let dump = |bytes: &[u8]| -> TestResult<String> {
        let package = litchi_iwa::Package::from_bytes(bytes)?;
        let mut out = String::new();
        let names: Vec<String> = package.iwa_entry_names().map(str::to_owned).collect();
        for name in names {
            for object in package.archive(&name)?.objects {
                for message in object.messages.iter().filter(|m| m.type_ == 6_006) {
                    let bucket = tst::HeaderStorageBucket::decode(message.data.as_slice())?;
                    let headers: Vec<_> = bucket
                        .headers
                        .iter()
                        .map(|h| (h.index, h.number_of_cells))
                        .collect();
                    out.push_str(&format!(
                        "{name} {:?} {headers:?}\n",
                        object.archive_info.identifier
                    ));
                }
            }
        }
        Ok(out)
    };
    let package = litchi_numbers::Package::from_bytes(&source.bytes)?;
    let value = package
        .edit_table_cells(0usize, 0usize)?
        .set(
            At::new(2, 2),
            litchi_numbers::table::cells::Input::number(4.0)?,
        )?
        .commit()?
        .into_package();
    let formula = leaf_route(&source.bytes, (2, 2), AUTHOR)?.output;
    let host = host_route(&source.bytes, (2, 2), AUTHOR, 43.0)?;
    eprintln!(
        "SOURCE\n{}VALUE\n{}LEAF-FORMULA\n{}HOST-FORMULA\n{}",
        dump(&source.bytes)?,
        dump(&leaf_bytes!(value))?,
        dump(&formula)?,
        dump(&host)?
    );
    Ok(())
}

/// Release-mode host-vs-leaf formula-authoring probe. Run the ignored test
/// under pinned-CPU callgrind with `host_route` and `leaf_cost_route` toggles;
/// it deliberately excludes the parity harness' inverse patch readback.
#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn numbers_table_formula_author_cost_abba() -> TestResult {
    let inputs = [
        (fixture("iwork/numbers/basic.numbers")?, AUTHOR, 43.0),
        (
            fixture("iwork/numbers/formula-budget-native.numbers")?,
            AUTHOR,
            2.0,
        ),
        (fixture("iwork/numbers/basic.numbers")?, CACHED, 43.0),
    ];
    for (input, formula, evaluated_cache) in inputs {
        black_box(host_route(&input.bytes, (2, 2), formula, evaluated_cache)?);
        black_box(leaf_cost_route(&input.bytes, (2, 2), formula)?);
        black_box(leaf_cost_route(&input.bytes, (2, 2), formula)?);
        black_box(host_route(&input.bytes, (2, 2), formula, evaluated_cache)?);
    }
    Ok(())
}

#[test]
#[ignore]
fn scratch_tile_probe() -> TestResult {
    use litchi_iwa_protos::{tsce, tst};
    use prost::Message;
    let source = fixture("iwork/numbers/basic.numbers")?;
    let dump = |bytes: &[u8], label: &str| -> TestResult {
        let package = litchi_iwa::Package::from_bytes(bytes)?;
        let mut out = String::new();
        let names: Vec<String> = package.iwa_entry_names().map(str::to_owned).collect();
        for name in names {
            for object in package.archive(&name)?.objects {
                for message in &object.messages {
                    if message.type_ == 6_002 {
                        out.push_str(&format!(
                            "{name} {:#?}\n",
                            tst::Tile::decode(message.data.as_slice())?
                        ));
                    }
                    if message.type_ == 4_008 {
                        out.push_str(&format!(
                            "{name} {:#?}\n",
                            tsce::FormulaOwnerDependenciesArchive::decode(message.data.as_slice())?
                        ));
                    }
                    if message.type_ == 6_005 {
                        out.push_str(&format!(
                            "{name} {:#?}\n",
                            tst::TableDataList::decode(message.data.as_slice())?
                        ));
                    }
                }
            }
        }
        std::fs::write(
            format!("/home/zhuhe/code/litchi-worktrees/.coord/iwa-w1f/f-{label}.txt"),
            out,
        )?;
        Ok(())
    };
    dump(&source.bytes, "source")?;
    dump(&leaf_route(&source.bytes, (2, 2), AUTHOR)?.output, "leaf")?;
    dump(&host_route(&source.bytes, (2, 2), AUTHOR, 43.0)?, "host")?;
    let budget = fixture("iwork/numbers/formula-budget-native.numbers")?;
    dump(&budget.bytes, "bsource")?;
    dump(&leaf_route(&budget.bytes, (2, 2), AUTHOR)?.output, "bleaf")?;
    dump(&host_route(&budget.bytes, (2, 2), AUTHOR, 2.0)?, "bhost")?;
    Ok(())
}

#[test]
#[ignore]
fn scratch_cell_probe() -> TestResult {
    use litchi_numbers_wire::{Limits, table_write};
    let source = fixture("iwork/numbers/basic.numbers")?;
    let cell = |bytes: &[u8]| -> TestResult<String> {
        let package = litchi_iwa::Package::from_bytes(bytes)?;
        let names: Vec<String> = package.iwa_entry_names().map(str::to_owned).collect();
        let mut out = String::new();
        for name in names {
            for object in package.archive(&name)?.objects {
                for message in object.messages.iter().filter(|m| m.type_ == 6_002) {
                    if let Some(cell) =
                        table_write::tile_cell(&message.data, 2, 2, &Limits::DEFAULT)?
                    {
                        out.push_str(&format!("{cell:02x?}"));
                    }
                }
            }
        }
        Ok(out)
    };
    eprintln!(
        "LEAF {}",
        cell(&leaf_route(&source.bytes, (2, 2), AUTHOR)?.output)?
    );
    eprintln!(
        "HOST {}",
        cell(&host_route(&source.bytes, (2, 2), AUTHOR, 43.0)?)?
    );
    {
        let package = litchi_numbers::Package::from_bytes(&source.bytes)?;
        let result = package
            .edit_table_cells(0usize, 0usize)?
            .set_formula_cached(
                At::new(2, 2),
                leaf_expression(CACHED)?,
                CachedValue::number(8.0)?,
            )
            .and_then(|edit| edit.commit().map(|_| ()));
        eprintln!("LEAFC {result:?}");
    }
    eprintln!(
        "HOSTC {}",
        cell(&host_route(&source.bytes, (2, 2), CACHED, 43.0)?)?
    );
    Ok(())
}
