//! Numbers conditional-highlight parity (record 0370d): the host editor's
//! `set_cell_conditional_highlighting`, `clear_cell_conditional_highlighting`,
//! `cell_conditional_highlighting` and `cell_conditional_highlight_rules`
//! against the leaf's `edit_table_cell_conditional_highlights` /
//! `table_cell_conditional_highlight`.
//!
//! A table's first rules give its model a fresh conditional-style owner
//! (`TableModelArchive.conditional_style_formula_owner_id`), random on the
//! host and deterministic on the leaf, and register fresh object UUIDs;
//! those writes are tier B, with the model decoder below dropping the
//! owner field (a documented policy, as for the metadata UUIDs).

use litchi_iwa::numbers;
use litchi_iwa_common::color::{RgbColorSpace, Rgba};
use litchi_numbers::cell::conditional_highlight::{
    Condition, Date, DateRange, Number, Range, Rule, Style, Text,
};
use litchi_numbers::table::At;
use prost::Message;

use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier, fixture};
use crate::known_fields::{Decoded, decode_as, decoded};
use crate::leaf_bytes;

const FIXTURES: [&str; 2] = [
    "iwork/numbers/basic.numbers",
    "iwork/numbers/custom-number-native.numbers",
];

/// The cells each fixture's cases write: a text or number cell and an
/// empty one.
const CELLS: [(usize, usize); 3] = [(1, 1), (2, 2), (6, 4)];

#[derive(Clone, Debug)]
enum Operation {
    Set(Vec<Rule>),
    Clear,
}

/// `TST.TableModelArchive` without its conditional-style owner UUID.
fn table_model_without_owner(data: &[u8]) -> Result<Decoded, prost::DecodeError> {
    let mut model = litchi_iwa_protos::tst::TableModelArchive::decode(data)?;
    model.conditional_style_formula_owner_id = None;
    Ok(decoded(model, data))
}

fn host_model(bytes: &[u8]) -> TestResult<u64> {
    crate::numbers_cases::host_first_table_model_id(&numbers::Draft::from_bytes(bytes)?)
}

fn host_route(source: &[u8], (r, c): (usize, usize), operation: &Operation) -> TestResult<Vec<u8>> {
    let mut editor = numbers::Draft::from_bytes(source)?;
    let id = host_model(source)?;
    match operation {
        Operation::Set(rules) => {
            editor.set_cell_conditional_highlighting(id, r, c, rules)?;
        },
        Operation::Clear => editor.clear_cell_conditional_highlighting(id, r, c)?,
    }
    Ok(editor.to_bytes()?)
}

fn leaf_route(source: &[u8], (r, c): (usize, usize), operation: &Operation) -> TestResult<LeafRun> {
    let package = litchi_numbers::Package::from_bytes(source)?;
    let at = At::new(r as u32, c as u32);
    let edit = package.edit_table_cell_conditional_highlights(0usize, 0usize)?;
    let edit = match operation {
        Operation::Set(rules) => edit.set(at, rules.iter().cloned())?,
        Operation::Clear => edit.clear(at)?,
    };
    let commit = edit.commit()?;
    let restored = commit
        .package()
        .apply_table_cell_conditional_highlights(&commit.patch().inverse())?;
    Ok(LeafRun::with_inverse(
        leaf_bytes!(commit.package()),
        leaf_bytes!(restored.package()),
    ))
}

fn host_read(bytes: &[u8], (r, c): (usize, usize)) -> TestResult<String> {
    let editor = numbers::Draft::from_bytes(bytes)?;
    let id = host_model(bytes)?;
    let rules = editor
        .cell_conditional_highlight_rules(id, r, c)
        .ok()
        .flatten();
    let count = editor
        .cell_conditional_highlighting(id, r, c)
        .ok()
        .flatten()
        .map(|info| info.rule_count);
    Ok(format!("{rules:?} {count:?}"))
}

fn leaf_read(bytes: &[u8], (r, c): (usize, usize)) -> TestResult<String> {
    let package = litchi_numbers::Package::from_bytes(bytes)?;
    let rules = package
        .table_cell_conditional_highlight(0usize, 0usize, At::new(r as u32, c as u32))
        .ok()
        .flatten();
    let count = rules.as_ref().map(|rules| rules.len() as u32);
    Ok(format!("{rules:?} {count:?}"))
}

fn run(
    name: &'static str,
    tier: Tier,
    input: &Fixture,
    cell: (usize, usize),
    operation: &Operation,
) -> TestResult {
    Case::new(name, Format::Numbers, tier, Evidence::Queued, input)
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

fn color(red: f32, green: f32, blue: f32) -> TestResult<Rgba> {
    Ok(Rgba::new(red, green, blue, 1.0, RgbColorSpace::Srgb)?)
}

/// One rule per non-volatile condition kind, with fill, text colour and
/// bold styles in turn.
fn samples() -> TestResult<Vec<Rule>> {
    let date = Date::from_ymd(2024, 3, 15)?;
    let later = Date::from_ymd(2024, 4, 1)?;
    let conditions = vec![
        Condition::CellIsBlank,
        Condition::CellIsNotBlank,
        Condition::CheckboxIsChecked,
        Condition::CheckboxIsNotChecked,
        Condition::BooleanIsTrue,
        Condition::BooleanIsFalse,
        Condition::NumberIsPositive,
        Condition::NumberIsNegative,
        Condition::DateIs(date),
        Condition::DateIsBefore(date),
        Condition::DateIsAfter(date),
        Condition::DateIsBetween(DateRange::new(date, later)?),
        Condition::EqualTo(Number::new(5.0)?),
        Condition::NotEqualTo(Number::new(5.0)?),
        Condition::GreaterThan(Number::new(2.5)?),
        Condition::GreaterThanOrEqualTo(Number::new(-1.0)?),
        Condition::LessThan(Number::new(100.0)?),
        Condition::LessThanOrEqualTo(Number::new(0.0)?),
        Condition::Between(Range::new(Number::new(1.0)?, Number::new(10.0)?)?),
        Condition::NotBetween(Range::new(Number::new(-3.0)?, Number::new(3.0)?)?),
        Condition::TextEqualTo(Text::new("Total")?),
        Condition::TextNotEqualTo(Text::new("Total")?),
        Condition::TextStartsWith(Text::new("To")?),
        Condition::TextDoesNotStartWith(Text::new("To")?),
        Condition::TextEndsWith(Text::new("al")?),
        Condition::TextDoesNotEndWith(Text::new("al")?),
        Condition::TextContains(Text::new("ot")?),
        Condition::TextDoesNotContain(Text::new("ot")?),
    ];
    let styles = [
        Style::with_fill(color(1.0, 0.8, 0.2)?),
        Style::with_text_color(color(0.8, 0.1, 0.1)?),
        Style::new(
            Some(color(0.2, 0.6, 0.3)?),
            Some(color(0.0, 0.0, 0.5)?),
            true,
        )?,
    ];
    Ok(conditions
        .into_iter()
        .enumerate()
        .map(|(index, condition)| Rule::new(condition, styles[index % styles.len()]))
        .collect())
}

fn report(failures: Vec<String>, cases: usize, minimum: usize) -> TestResult {
    eprintln!("conditional-highlight cases: {cases}");
    if failures.is_empty() {
        assert!(cases >= minimum, "only {cases} cases ran");
        Ok(())
    } else {
        Err(failures.join("\n").into())
    }
}

/// Each condition kind as a cell's single rule.
#[test]
fn numbers_conditional_highlight_set_parity() -> TestResult {
    let _ = Evidence::Queued;
    let mut failures = Vec::new();
    let mut cases = 0;
    for path in FIXTURES {
        let input = fixture(path)?;
        for cell in CELLS {
            for rule in samples()? {
                cases += 1;
                let operation = Operation::Set(vec![rule]);
                if let Err(error) = run(
                    "numbers_conditional_highlight_set_parity",
                    Tier::B,
                    &input,
                    cell,
                    &operation,
                ) {
                    failures.push(format!("{path} {cell:?} {operation:?}: {error}"));
                }
            }
        }
    }
    report(failures, cases, 100)
}

/// Several rules on one cell; then replacing them in place with fewer and
/// with more rules, and a second cell sharing the table's list.
#[test]
fn numbers_conditional_highlight_replace_parity() -> TestResult {
    let _ = Evidence::Queued;
    let mut failures = Vec::new();
    let mut cases = 0;
    let rules = samples()?;
    for path in FIXTURES {
        let input = fixture(path)?;
        let cell = CELLS[0];
        let three = Operation::Set(rules[12..15].to_vec());
        let mut check = |input: &Fixture, cell, operation: Operation| {
            cases += 1;
            if let Err(error) = run(
                "numbers_conditional_highlight_replace_parity",
                Tier::B,
                input,
                cell,
                &operation,
            ) {
                failures.push(format!("{path} {cell:?} {operation:?}: {error}"));
            }
        };
        check(&input, cell, three.clone());
        let base = with(&input, cell, &three)?;
        check(&base, cell, Operation::Set(rules[20..21].to_vec()));
        check(&base, cell, Operation::Set(rules[0..5].to_vec()));
        check(&base, cell, Operation::Set(rules[18..21].to_vec()));
        check(&base, CELLS[1], Operation::Set(rules[22..24].to_vec()));
    }
    report(failures, cases, 10)
}

/// Clearing a cell's rules (the graph goes with its entry).
#[test]
fn numbers_conditional_highlight_clear_parity() -> TestResult {
    let _ = Evidence::Queued;
    let mut failures = Vec::new();
    let mut cases = 0;
    let rules = samples()?;
    for path in FIXTURES {
        let input = fixture(path)?;
        for cell in CELLS {
            let base = with(&input, cell, &Operation::Set(rules[12..15].to_vec()))?;
            cases += 1;
            if let Err(error) = run(
                "numbers_conditional_highlight_clear_parity",
                Tier::B,
                &base,
                cell,
                &Operation::Clear,
            ) {
                failures.push(format!("{path} {cell:?} clear: {error}"));
            }
        }
    }
    report(failures, cases, 6)
}

/// The rules and rule count read equal on corpus cells and on every
/// written cell.
#[test]
fn numbers_conditional_highlight_read_parity() -> TestResult {
    let (_tier, _evidence) = (Tier::A, Evidence::Queued);
    let mut compared = 0;
    for path in crate::numbers_cases::NUMBERS_CORPUS {
        let input = fixture(path)?;
        if host_model(&input.bytes).is_err() {
            continue;
        }
        for cell in CELLS {
            assert_eq!(
                leaf_read(&input.bytes, cell)?,
                host_read(&input.bytes, cell)?,
                "{path} {cell:?}"
            );
            compared += 1;
        }
    }
    let rules = samples()?;
    for path in FIXTURES {
        let input = fixture(path)?;
        for chunk in rules.chunks(4) {
            let written = with(&input, CELLS[0], &Operation::Set(chunk.to_vec()))?;
            assert_eq!(
                leaf_read(&written.bytes, CELLS[0])?,
                host_read(&written.bytes, CELLS[0])?,
                "{path} {chunk:?}"
            );
            compared += 1;
        }
    }
    eprintln!("conditional-highlight reads compared: {compared}");
    assert!(compared >= 40, "only {compared} reads compared");
    Ok(())
}

/// Volatile date rules are refused by the leaf (the host registers a
/// calculation-engine owner family for them).
#[test]
fn numbers_conditional_highlight_volatile_refusal() -> TestResult {
    let (_tier, _evidence) = (Tier::A, Evidence::Queued);
    let input = fixture(FIXTURES[0])?;
    let package = litchi_numbers::Package::from_bytes(&input.bytes)?;
    let rule = Rule::new(
        Condition::DateIsToday,
        Style::with_fill(color(1.0, 0.0, 0.0)?),
    );
    let refused = package
        .edit_table_cell_conditional_highlights(0usize, 0usize)?
        .set(At::new(1, 1), [rule]);
    assert!(matches!(
        refused,
        Err(litchi_numbers::cell::conditional_highlight::Error::Refused(
            litchi_numbers::cell::conditional_highlight::Refusal::VolatileRule
        ))
    ));
    Ok(())
}
