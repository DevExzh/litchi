//! Numbers formulas that reference a pivot aggregate category (record 0424):
//! the host editor's `set_formula_with_cached_value` with a
//! `FormulaExpression::PivotCategory` against the leaf's
//! `litchi_numbers::table::cells::Edit::set_formula_cached` with
//! `Expression::pivot_category`.
//!
//! No tracked native workbook holds a pivot table, so the cases run on the
//! synthetic group-by graphs of record 0421 (`numbers_pivot_cases`); the
//! route is macOS-queued.

use std::hint::black_box;

use litchi_iwa::numbers;
use litchi_iwa_common::formula::{FormulaBinaryOperator, FormulaCachedValue, FormulaExpression};
use litchi_numbers::formula::{BinaryOperator, CachedValue, Expression};
use litchi_numbers::table::{
    At,
    cells::{Error, Refusal},
    pivot::Category,
};

use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier};
use crate::leaf_bytes;
use crate::numbers_pivot_cases::{inline_pivot_fixture, pivot_fixture};

/// The cell the formula is authored in: C3.
const CELL: (usize, usize) = (2, 2);
/// The formula is `category / 2`, with this display cache (`None`: no
/// supplied cache).
const CACHE: f64 = 21.5;

/// One authored formula: which category and whether a cache is supplied.
#[derive(Clone, Copy, Debug)]
struct Authoring {
    label: &'static str,
    index: usize,
    cached: Option<f64>,
}

const NORTH: Authoring = Authoring {
    label: "North",
    index: 0,
    cached: Some(CACHE),
};

/// The position, in both routes' category order, of the `index`-th category
/// labelled `label` (record 0421 proves the orders agree).
fn position(labels: &[Option<String>], label: &str, index: usize) -> TestResult<usize> {
    labels
        .iter()
        .enumerate()
        .filter(|(_, found)| found.as_deref() == Some(label))
        .nth(index)
        .map(|(position, _)| position)
        .ok_or_else(|| format!("no category {label:?} #{index}").into())
}

fn host_expression(
    source: &numbers::Draft,
    label: &str,
    index: usize,
) -> TestResult<FormulaExpression> {
    let categories = source.pivot_categories()?;
    let labels = categories
        .iter()
        .map(|category| category.label.clone())
        .collect::<Vec<_>>();
    let category = &categories[position(&labels, label, index)?];
    Ok(FormulaExpression::binary(
        FormulaBinaryOperator::Divide,
        FormulaExpression::pivot_category(category.reference),
        FormulaExpression::Number(2.0),
    ))
}

fn leaf_category(
    source: &litchi_numbers::Package,
    label: &str,
    index: usize,
) -> TestResult<Category> {
    let categories = source.pivot_categories()?;
    let labels = categories
        .iter()
        .map(|category| category.label().map(str::to_owned))
        .collect::<Vec<_>>();
    Ok(categories[position(&labels, label, index)?].clone())
}

fn leaf_expression(category: &Category) -> TestResult<Expression> {
    Ok(Expression::binary(
        BinaryOperator::Divide,
        Expression::pivot_category(category),
        Expression::number(2.0)?,
    )?)
}

#[inline(never)]
fn pivot_host_route(source: &[u8], authoring: Authoring) -> TestResult<Vec<u8>> {
    let mut editor = numbers::Draft::from_bytes(source)?;
    let table = crate::numbers_cases::host_first_table_model_id(&editor)?;
    let expression = host_expression(&editor, authoring.label, authoring.index)?;
    match authoring.cached {
        None => editor.set_formula(table, CELL.0, CELL.1, expression)?,
        Some(value) => editor.set_formula_with_cached_value(
            table,
            CELL.0,
            CELL.1,
            expression,
            FormulaCachedValue::number(value)?,
        )?,
    }
    Ok(editor.to_bytes()?)
}

fn leaf_commit(
    source: &[u8],
    authoring: Authoring,
) -> TestResult<litchi_numbers::table::cells::Commit> {
    let package = litchi_numbers::Package::from_bytes(source)?;
    let category = leaf_category(&package, authoring.label, authoring.index)?;
    let at = At::new(u32::try_from(CELL.0)?, u32::try_from(CELL.1)?);
    let edit = package.edit_table_cells(0usize, 0usize)?;
    let expression = leaf_expression(&category)?;
    let edit = match authoring.cached {
        None => edit.set_formula(at, expression)?,
        Some(value) => edit.set_formula_cached(at, expression, CachedValue::number(value)?)?,
    };
    Ok(edit.commit()?)
}

fn leaf_route(source: &[u8], authoring: Authoring) -> TestResult<LeafRun> {
    let commit = leaf_commit(source, authoring)?;
    let restored = commit
        .package()
        .apply_table_cells(&commit.patch().inverse())?;
    Ok(LeafRun::with_inverse(
        leaf_bytes!(commit.package()),
        leaf_bytes!(restored.package()),
    ))
}

#[inline(never)]
fn pivot_leaf_route(source: &[u8], authoring: Authoring) -> TestResult<Vec<u8>> {
    let commit = leaf_commit(source, authoring)?;
    Ok(leaf_bytes!(commit.package()))
}

/// The authored formula archive: the formula list entry the C3 cell names.
fn formula_archive(bytes: &[u8]) -> TestResult<Vec<u8>> {
    use litchi_iwa_protos::{tsce, tst};
    use prost::Message;

    let package = litchi_iwa::Package::from_bytes(bytes)?;
    let names: Vec<String> = package.iwa_entry_names().map(str::to_owned).collect();
    let mut found = Vec::new();
    for name in names {
        for object in package.archive(&name)?.objects {
            for message in object.messages.iter().filter(|m| m.type_ == 6_005) {
                let list = tst::TableDataList::decode(message.data.as_slice())?;
                for entry in list.entries {
                    if let Some(formula) = entry.formula {
                        let bytes = formula.encode_to_vec();
                        let decoded = tsce::FormulaArchive::decode(bytes.as_slice())?;
                        if decoded
                            .ast_node_array
                            .ast_node
                            .iter()
                            .any(|node| node.ast_category_ref.is_some())
                        {
                            found.push(bytes);
                        }
                    }
                }
            }
        }
    }
    match found.as_slice() {
        [one] => Ok(one.clone()),
        _ => Err(format!("expected one category formula, found {}", found.len()).into()),
    }
}

/// The formula's dependency record: C3's expanded edges in its owner.
fn dependency_edges(bytes: &[u8]) -> TestResult<String> {
    use litchi_iwa_protos::tsce;
    use prost::Message;

    let package = litchi_iwa::Package::from_bytes(bytes)?;
    let names: Vec<String> = package.iwa_entry_names().map(str::to_owned).collect();
    let mut edges = Vec::new();
    for name in names {
        for object in package.archive(&name)?.objects {
            for message in object.messages.iter().filter(|m| m.type_ == 4_009) {
                let tile = tsce::CellRecordTileArchive::decode(message.data.as_slice())?;
                for record in tile.cell_records {
                    if (record.row, record.column) == (2, 2) {
                        edges.push(format!("{:?}", record.expanded_edges));
                    }
                }
            }
        }
    }
    Ok(edges.join(" | "))
}

fn cell_projection(bytes: &[u8]) -> TestResult<String> {
    let package = litchi_numbers::Package::from_bytes(bytes)?;
    let state = package.table_cell(0usize, 0usize, At::new(2, 2))?;
    Ok(format!("{:?}", state.storage()))
}

fn run(name: &'static str, input: &Fixture, authoring: Authoring) -> TestResult {
    Case::new(name, Format::Numbers, Tier::B, Evidence::Queued, input)
        .leaf_structural_metadata()
        .leaf_invalidates_root_previews()
        .host(move |source| pivot_host_route(source, authoring))
        .leaf(move |source| leaf_route(source, authoring))
        .read("C3 formula archive", formula_archive)
        .read("C3 dependency edges", dependency_edges)
        .read("C3 cell", cell_projection)
        .run()?;
    // The formula archives are byte-identical: the category node, its
    // translation flags and its summary-function code; so are C3's
    // dependency edges.
    let host = pivot_host_route(&input.bytes, authoring)?;
    let leaf = pivot_leaf_route(&input.bytes, authoring)?;
    assert_eq!(formula_archive(&host)?, formula_archive(&leaf)?);
    assert_eq!(dependency_edges(&host)?, dependency_edges(&leaf)?);
    assert!(dependency_edges(&leaf)?.contains("internal_owner_id_for_edge: [13, 13]"));
    Ok(())
}

/// `North / 2` over the inline group-by's first aggregated column (a
/// level-1 group with a string value).
#[test]
fn numbers_pivot_formula_author_parity() -> TestResult {
    let input = pivot_fixture()?;
    run("numbers_pivot_formula_author_parity", &input, NORTH)
}

/// Without a supplied cache: a category formula is not evaluable, so the
/// leaf leaves the display cache as the host does (record 0424).
#[test]
fn numbers_pivot_formula_author_without_cache_parity() -> TestResult {
    let input = pivot_fixture()?;
    run(
        "numbers_pivot_formula_author_without_cache_parity",
        &input,
        Authoring {
            cached: None,
            ..NORTH
        },
    )
}

/// A category of the referenced group-by tree (aggregators and group
/// nodes as separate objects), whose column has another summary function.
#[test]
fn numbers_pivot_formula_author_referenced_tree_parity() -> TestResult {
    let input = pivot_fixture()?;
    run(
        "numbers_pivot_formula_author_referenced_tree_parity",
        &input,
        Authoring { index: 2, ..NORTH },
    )
}

/// The same route on the fixture without the cycle and the repeated
/// identifiers, to tell a refused fixture from a refused route on a Mac
/// (record 0431).
#[test]
fn numbers_pivot_formula_author_inline_only_parity() -> TestResult {
    let input = inline_pivot_fixture()?;
    run(
        "numbers_pivot_formula_author_inline_only_parity",
        &input,
        NORTH,
    )
}

/// A category is resolved against the edited package, as the host resolves
/// it: a category read from another workbook is refused when the edited one
/// lacks it, and the host refuses the same reference.
#[test]
fn numbers_pivot_formula_missing_category_is_refused_like_host() -> TestResult {
    let pivot = pivot_fixture()?;
    let plain = crate::harness::fixture("iwork/numbers/basic.numbers")?;
    let source = litchi_numbers::Package::from_bytes(&pivot.bytes)?;
    let category = leaf_category(&source, "North", 0)?;
    let target = litchi_numbers::Package::from_bytes(&plain.bytes)?;
    let refused = target
        .edit_table_cells(0usize, 0usize)?
        .set_formula_cached(
            At::new(2, 2),
            leaf_expression(&category)?,
            CachedValue::number(CACHE)?,
        )?
        .commit()
        .expect_err("a category absent from the edited package is refused");
    assert_eq!(refused, Error::Refused(Refusal::PivotCategoryNotFound));

    let host_source = numbers::Draft::from_bytes(&pivot.bytes)?;
    let expression = host_expression(&host_source, "North", 0)?;
    let mut host = numbers::Draft::from_bytes(&plain.bytes)?;
    let table = crate::numbers_cases::host_first_table_model_id(&host)?;
    assert!(
        host.set_formula_with_cached_value(
            table,
            2,
            2,
            expression,
            FormulaCachedValue::number(CACHE)?
        )
        .is_err()
    );
    Ok(())
}

/// A category whose summary function differs in the edited package is
/// refused, as the host refuses the same reference.
#[test]
fn numbers_pivot_formula_changed_category_is_refused_like_host() -> TestResult {
    let pivot = pivot_fixture()?;
    let changed = crate::numbers_pivot_cases::pivot_fixture_with(3)?;
    let source = litchi_numbers::Package::from_bytes(&pivot.bytes)?;
    let category = leaf_category(&source, "North", 0)?;
    let target = litchi_numbers::Package::from_bytes(&changed.bytes)?;
    let refused = target
        .edit_table_cells(0usize, 0usize)?
        .set_formula(At::new(2, 2), leaf_expression(&category)?)?
        .commit()
        .expect_err("a category with another summary function is refused");
    assert_eq!(refused, Error::Refused(Refusal::PivotCategoryChanged));
    // The changed category is still read, with another summary function.
    assert_ne!(
        leaf_category(&target, "North", 0)?.aggregate(),
        category.aggregate()
    );

    let host_source = numbers::Draft::from_bytes(&pivot.bytes)?;
    let expression = host_expression(&host_source, "North", 0)?;
    let mut host = numbers::Draft::from_bytes(&changed.bytes)?;
    let table = crate::numbers_cases::host_first_table_model_id(&host)?;
    assert!(host.set_formula(table, 2, 2, expression).is_err());
    Ok(())
}

/// Callgrind probe (record 0424): author `North / 2` with a cache, host
/// against leaf, ABBA (`--toggle-collect=*pivot_host_route*` /
/// `*pivot_leaf_route*`).
#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn numbers_pivot_formula_author_cost_abba() -> TestResult {
    let input = pivot_fixture()?;
    for route in [0_u8, 1, 1, 0] {
        if route == 0 {
            black_box(pivot_host_route(&input.bytes, NORTH)?);
        } else {
            black_box(pivot_leaf_route(&input.bytes, NORTH)?);
        }
    }
    Ok(())
}
