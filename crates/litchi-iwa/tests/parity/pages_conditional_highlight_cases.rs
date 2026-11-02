//! Pages body-table conditional-highlight parity (record 0412f): the host
//! editor's `set_table_cell_conditional_highlighting`,
//! `clear_table_cell_conditional_highlighting`,
//! `table_cell_conditional_highlighting` and
//! `table_cell_conditional_highlight_rules` against the leaf's
//! `edit_body_table_cell_conditional_highlights` /
//! `body_table_cell_conditional_highlight`, over the shared engine the
//! Numbers family runs.
//!
//! As on Numbers, a table's first rules give its model a fresh
//! conditional-style owner, random on the host and deterministic on the
//! leaf, so the writes are tier B with that owner field dropped.

use litchi_iwa::pages;
use litchi_iwa_common::table::cell::conditional_highlight::Rule;
use litchi_iwa_common::table::coordinate::At;
use litchi_pages::Package;

use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier, fixture};
use crate::known_fields::decode_as;
use crate::leaf_bytes;
use crate::numbers_conditional_highlight_cases::{report, samples, table_model_without_owner};
use crate::pages_table_cases::{CATALOG, TABLE, host_model};

/// The cells each case writes: a stored cell and an empty one.
const CELLS: [(usize, usize); 2] = [(1, 1), (2, 2)];

#[derive(Clone, Debug)]
enum Operation {
    Set(Vec<Rule>),
    Clear,
}

fn host_route(source: &[u8], (r, c): (usize, usize), operation: &Operation) -> TestResult<Vec<u8>> {
    let mut editor = pages::Draft::from_bytes(source)?;
    let model = host_model(&editor)?;
    match operation {
        Operation::Set(rules) => {
            editor.set_table_cell_conditional_highlighting(model, r, c, rules)?;
        },
        Operation::Clear => editor.clear_table_cell_conditional_highlighting(model, r, c)?,
    }
    Ok(editor.to_bytes()?)
}

fn leaf_commit(
    package: &Package,
    (r, c): (usize, usize),
    operation: &Operation,
) -> TestResult<litchi_pages::table::conditional_highlight::Commit> {
    let at = At::new(r as u32, c as u32);
    let edit = package.edit_body_table_cell_conditional_highlights(TABLE)?;
    let edit = match operation {
        Operation::Set(rules) => edit.set(at, rules.iter().cloned())?,
        Operation::Clear => edit.clear(at)?,
    };
    Ok(edit.commit()?)
}

fn leaf_route(source: &[u8], cell: (usize, usize), operation: &Operation) -> TestResult<LeafRun> {
    let package = Package::from_bytes(source)?;
    let commit = leaf_commit(&package, cell, operation)?;
    let restored = commit
        .package()
        .apply_body_table_cell_conditional_highlights(&commit.patch().inverse())?;
    Ok(LeafRun::with_inverse(
        leaf_bytes!(commit.package()),
        leaf_bytes!(restored.package()),
    ))
}

fn host_read(bytes: &[u8], (r, c): (usize, usize)) -> TestResult<String> {
    let editor = pages::Draft::from_bytes(bytes)?;
    let model = host_model(&editor)?;
    let rules = editor
        .table_cell_conditional_highlight_rules(model, r, c)
        .ok()
        .flatten();
    let count = editor
        .table_cell_conditional_highlighting(model, r, c)
        .ok()
        .flatten()
        .map(|info| info.rule_count);
    Ok(format!("{rules:?} {count:?}"))
}

fn leaf_read(bytes: &[u8], (r, c): (usize, usize)) -> TestResult<String> {
    let package = Package::from_bytes(bytes)?;
    let rules = package
        .body_table_cell_conditional_highlight(TABLE, At::new(r as u32, c as u32))
        .ok()
        .flatten();
    let count = rules.as_ref().map(|rules| rules.len() as u32);
    Ok(format!("{rules:?} {count:?}"))
}

fn run(
    name: &'static str,
    tier: Tier,
    evidence: Evidence,
    input: &Fixture,
    cell: (usize, usize),
    operation: &Operation,
) -> TestResult {
    Case::new(name, Format::Pages, tier, evidence, input)
        .leaf_structural_metadata()
        .known_fields(|known| {
            known
                .with(
                    6_001,
                    "TST.TableModelArchive (conditional-style owner dropped)",
                    table_model_without_owner,
                )
                .with(
                    6_010,
                    "TST.ConditionalStyleSetArchive",
                    decode_as::<litchi_iwa_protos::tst::ConditionalStyleSetArchive>,
                )
                .with(
                    6_004,
                    "TST.CellStyleArchive",
                    decode_as::<litchi_iwa_protos::tst::CellStyleArchive>,
                )
                .with(
                    2_022,
                    "TSWP.ParagraphStyleArchive",
                    decode_as::<litchi_iwa_protos::tswp::ParagraphStyleArchive>,
                )
        })
        .host(|source| host_route(source, cell, operation))
        .leaf(|source| leaf_route(source, cell, operation))
        .read("host read", move |bytes| host_read(bytes, cell))
        .read("leaf read", move |bytes| leaf_read(bytes, cell))
        .run()?;
    Ok(())
}

fn with(input: &Fixture, cell: (usize, usize), operation: &Operation) -> TestResult<Fixture> {
    Ok(Fixture {
        path: input.path,
        bytes: host_route(&input.bytes, cell, operation)?,
    })
}

/// Each non-volatile condition kind as a cell's single rule.
#[test]
fn pages_body_table_conditional_highlight_set_parity() -> TestResult {
    let input = fixture(CATALOG)?;
    let mut failures = Vec::new();
    let mut cases = 0;
    for cell in CELLS {
        for rule in samples()? {
            cases += 1;
            let operation = Operation::Set(vec![rule]);
            if let Err(error) = run(
                "pages_body_table_conditional_highlight_set_parity",
                Tier::B,
                Evidence::Queued,
                &input,
                cell,
                &operation,
            ) {
                failures.push(format!("{cell:?} {operation:?}: {error}"));
            }
        }
    }
    report(failures, cases, 50)
}

/// Several rules on one cell; replacing them in place with fewer and more
/// rules; a second cell sharing the table's list.
#[test]
fn pages_body_table_conditional_highlight_replace_parity() -> TestResult {
    let input = fixture(CATALOG)?;
    let rules = samples()?;
    let mut failures = Vec::new();
    let mut cases = 0;
    let three = Operation::Set(rules[12..15].to_vec());
    let mut check = |input: &Fixture, cell, operation: Operation| {
        cases += 1;
        if let Err(error) = run(
            "pages_body_table_conditional_highlight_replace_parity",
            Tier::B,
            Evidence::Queued,
            input,
            cell,
            &operation,
        ) {
            failures.push(format!("{cell:?} {operation:?}: {error}"));
        }
    };
    check(&input, CELLS[0], three.clone());
    let base = with(&input, CELLS[0], &three)?;
    check(&base, CELLS[0], Operation::Set(rules[20..21].to_vec()));
    check(&base, CELLS[0], Operation::Set(rules[0..5].to_vec()));
    check(&base, CELLS[1], Operation::Set(rules[22..24].to_vec()));
    report(failures, cases, 4)
}

/// Clearing a cell's rules (the graph goes with its entry).
#[test]
fn pages_body_table_conditional_highlight_clear_parity() -> TestResult {
    let input = fixture(CATALOG)?;
    let rules = samples()?;
    let mut failures = Vec::new();
    let mut cases = 0;
    for cell in CELLS {
        let base = with(&input, cell, &Operation::Set(rules[12..15].to_vec()))?;
        cases += 1;
        if let Err(error) = run(
            "pages_body_table_conditional_highlight_clear_parity",
            Tier::B,
            Evidence::Queued,
            &base,
            cell,
            &Operation::Clear,
        ) {
            failures.push(format!("{cell:?} clear: {error}"));
        }
    }
    report(failures, cases, 2)
}

/// The rules and rule count read equal on the corpus cells and on every
/// host-written cell.
#[test]
fn pages_body_table_conditional_highlight_read_parity() -> TestResult {
    let (_tier, _evidence) = (Tier::A, Evidence::Queued);
    let input = fixture(CATALOG)?;
    for cell in CELLS {
        assert_eq!(
            leaf_read(&input.bytes, cell)?,
            host_read(&input.bytes, cell)?
        );
    }
    for chunk in samples()?.chunks(4) {
        let written = with(&input, CELLS[0], &Operation::Set(chunk.to_vec()))?;
        assert_eq!(
            leaf_read(&written.bytes, CELLS[0])?,
            host_read(&written.bytes, CELLS[0])?
        );
    }
    Ok(())
}

/// A volatile date rule on a native table whose conditional-style owner is
/// registered without the calculation engine's volatile family: the host
/// refuses it, and so does the leaf, before anything is published.
#[test]
fn pages_body_table_conditional_highlight_volatile_refused_like_host() -> TestResult {
    use litchi_iwa_common::color::{RgbColorSpace, Rgba};
    use litchi_iwa_common::table::cell::conditional_highlight::{Condition, Style};
    let input = fixture(CATALOG)?;
    let operation = Operation::Set(vec![Rule::new(
        Condition::DateIsToday,
        Style::with_fill(Rgba::new(1.0, 0.0, 0.0, 1.0, RgbColorSpace::Srgb)?),
    )]);
    let host = host_route(&input.bytes, CELLS[0], &operation);
    let leaf = leaf_route(&input.bytes, CELLS[0], &operation);
    assert!(host.is_err(), "the host must refuse");
    assert_eq!(
        host.is_err(),
        leaf.is_err(),
        "host {:?} and leaf {:?} must agree",
        host.as_ref().err(),
        leaf.as_ref().err()
    );
    Ok(())
}

#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn pages_body_table_conditional_highlight_cost_abba() -> TestResult {
    let rules = samples()?[12..15].to_vec();
    let operation = Operation::Set(rules);
    crate::pages_cost_probes::abba(
        &fixture(CATALOG)?,
        |source| host_route(source, CELLS[0], &operation),
        |source| {
            let package = Package::from_bytes(source)?;
            Ok(leaf_bytes!(
                leaf_commit(&package, CELLS[0], &operation)?.package()
            ))
        },
    )
}
