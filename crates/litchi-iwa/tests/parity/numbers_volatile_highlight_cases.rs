//! Volatile conditional-highlight rules (record 0376a): the host editor's
//! `set_cell_conditional_highlighting` / `clear_cell_conditional_highlighting`
//! with today, yesterday or tomorrow rules, which register the calculation
//! engine's twelve-owner volatile family, against the leaf's
//! `edit_table_cell_conditional_highlights` over
//! `litchi_numbers_wire::conditional_volatile`.
//!
//! The family's conditional-style owner takes the table model's
//! conditional-style owner UUID (random on the host, deterministic on the
//! leaf, record 0370d) and its global-name owner a fresh UUID (random on the
//! host, derived on the leaf): the decoders below drop those UUIDs, and the
//! engine's owner-map UUIDs, before comparing (a documented policy).

use litchi_iwa::numbers;
use litchi_iwa_common::color::{RgbColorSpace, Rgba};
use litchi_iwa_protos::{tsce, tsp, tst};
use litchi_numbers::cell::conditional_highlight::{Condition, Rule, Style};
use litchi_numbers::table::At;
use prost::Message;

use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier, fixture};
use crate::known_fields::{Decoded, decode_as, decoded};
use crate::leaf_bytes;

/// The host registers the family only in a table whose conditional-style
/// owner is not yet registered: every native fixture registers it without
/// the family (the host refuses those, pinned below), so the inputs are host
/// `Builder` documents, as in the host's own volatile tests.
const DOCUMENTS: [(&str, usize, usize); 2] = [
    ("host Builder document, 8 x 6 table", 8, 6),
    ("host Builder document, 12 x 40 table", 12, 40),
];

fn document((path, rows, columns): (&'static str, usize, usize)) -> TestResult<Fixture> {
    Ok(Fixture {
        path,
        bytes: numbers::Builder::new()
            .table_dimensions(rows, columns)
            .build()?
            .to_bytes()?,
    })
}

/// The owner kinds whose UUIDs derive from a random value on the host.
const RANDOM_UUID_KINDS: [u32; 2] = [3, 7];

#[derive(Clone, Debug)]
enum Operation {
    Set(Vec<Rule>),
    Clear,
}

fn rule(condition: Condition) -> TestResult<Rule> {
    Ok(Rule::new(
        condition,
        Style::with_fill(Rgba::new(1.0, 0.0, 0.0, 1.0, RgbColorSpace::Srgb)?),
    ))
}

fn model_without_owner(data: &[u8]) -> Result<Decoded, prost::DecodeError> {
    let mut model = tst::TableModelArchive::decode(data)?;
    model.conditional_style_formula_owner_id = None;
    Ok(decoded(model, data))
}

fn owner_without_random_uuids(data: &[u8]) -> Result<Decoded, prost::DecodeError> {
    let mut owner = tsce::FormulaOwnerDependenciesArchive::decode(data)?;
    if owner
        .owner_kind
        .is_some_and(|kind| RANDOM_UUID_KINDS.contains(&kind))
    {
        owner.formula_owner_uid = tsp::Uuid::default();
        owner.base_owner_uid = owner.base_owner_uid.map(|_| tsp::Uuid::default());
    }
    Ok(decoded(owner, data))
}

fn engine_without_map_uuids(data: &[u8]) -> Result<Decoded, prost::DecodeError> {
    let mut engine = tsce::CalculationEngineArchive::decode(data)?;
    if let Some(map) = engine.dependency_tracker.owner_id_map.as_mut() {
        for entry in &mut map.map_entry {
            entry.owner_id = tsp::CfuuidArchive::default();
        }
    }
    Ok(decoded(engine, data))
}

fn host_route(source: &[u8], (r, c): (usize, usize), operation: &Operation) -> TestResult<Vec<u8>> {
    let mut editor = numbers::Draft::from_bytes(source)?;
    let id = crate::numbers_cases::host_first_table_model_id(&editor)?;
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
    let at = At::new(u32::try_from(r)?, u32::try_from(c)?);
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

fn leaf_read(bytes: &[u8], (r, c): (usize, usize)) -> TestResult<String> {
    let package = litchi_numbers::Package::from_bytes(bytes)?;
    let rules = package.table_cell_conditional_highlight(
        0usize,
        0usize,
        At::new(u32::try_from(r)?, u32::try_from(c)?),
    )?;
    Ok(format!("{rules:?}"))
}

/// The volatile owners of the calculation engine, as (kind, records).
fn volatile_family(bytes: &[u8]) -> TestResult<String> {
    let package = litchi_iwa::Package::from_bytes(bytes)?;
    let names: Vec<String> = package.iwa_entry_names().map(str::to_owned).collect();
    let mut family = Vec::new();
    for name in names {
        for object in package.archive(&name)?.objects {
            for message in object
                .messages
                .iter()
                .filter(|message| message.type_ == 4_008)
            {
                let owner = tsce::FormulaOwnerDependenciesArchive::decode(message.data.as_slice())?;
                if owner.owner_kind == Some(3) {
                    let records = owner
                        .cell_dependencies
                        .map(|cells| {
                            cells
                                .cell_record
                                .iter()
                                .map(|record| (record.row, record.column))
                                .collect::<Vec<_>>()
                        })
                        .unwrap_or_default();
                    family.push(format!("conditional-style records {records:?}"));
                }
            }
        }
    }
    Ok(family.join("; "))
}

fn run(
    name: &'static str,
    tier: Tier,
    evidence: Evidence,
    input: &Fixture,
    cell: (usize, usize),
    operation: &Operation,
) -> TestResult {
    Case::new(name, Format::Numbers, tier, evidence, input)
        .leaf_structural_metadata()
        .known_fields(|known| {
            known
                .with(
                    6_001,
                    "TST.TableModelArchive (conditional-style owner dropped)",
                    model_without_owner,
                )
                .with(
                    4_008,
                    "TSCE.FormulaOwnerDependenciesArchive (random owner UUIDs dropped)",
                    owner_without_random_uuids,
                )
                .with(
                    4_000,
                    "TSCE.CalculationEngineArchive (owner-map UUIDs dropped)",
                    engine_without_map_uuids,
                )
                .with(
                    6_010,
                    "TST.ConditionalStyleSetArchive",
                    decode_as::<tst::ConditionalStyleSetArchive>,
                )
                .with(
                    6_004,
                    "TST.CellStyleArchive",
                    decode_as::<tst::CellStyleArchive>,
                )
                .with(
                    2_022,
                    "TSWP.ParagraphStyleArchive",
                    decode_as::<litchi_iwa_protos::tswp::ParagraphStyleArchive>,
                )
        })
        .host(|source| host_route(source, cell, operation))
        .leaf(|source| leaf_route(source, cell, operation))
        .read("leaf read", move |bytes| leaf_read(bytes, cell))
        .read("volatile family", volatile_family)
        .run()?;
    Ok(())
}

fn with(input: &Fixture, cell: (usize, usize), operation: &Operation) -> TestResult<Fixture> {
    Ok(Fixture {
        path: input.path,
        bytes: host_route(&input.bytes, cell, operation)?,
    })
}

/// A table's first volatile rule registers the owner family.
#[test]
fn numbers_conditional_highlight_volatile_parity() -> TestResult {
    for spec in DOCUMENTS {
        let input = document(spec)?;
        let path = input.path;
        for (cell, condition) in [
            ((1, 1), Condition::DateIsToday),
            ((2, 2), Condition::DateIsYesterday),
            ((6, 4), Condition::DateIsTomorrow),
        ] {
            run(
                "numbers_conditional_highlight_volatile_parity",
                Tier::B,
                Evidence::Queued,
                &input,
                cell,
                &Operation::Set(vec![rule(condition)?]),
            )
            .map_err(|error| format!("{path} {cell:?}: {error}"))?;
        }
    }
    Ok(())
}

/// A second volatile cell joins the family; replacing a volatile rule and
/// clearing one unregister the cell.
#[test]
fn numbers_conditional_highlight_volatile_lifecycle_parity() -> TestResult {
    for spec in DOCUMENTS {
        let input = document(spec)?;
        let path = input.path;
        let first = with(
            &input,
            (1, 1),
            &Operation::Set(vec![rule(Condition::DateIsToday)?]),
        )?;
        for (label, cell, operation) in [
            (
                "second volatile cell",
                (2, 2),
                Operation::Set(vec![rule(Condition::DateIsTomorrow)?]),
            ),
            (
                "non-volatile replacement",
                (1, 1),
                Operation::Set(vec![rule(Condition::CellIsBlank)?]),
            ),
            (
                "volatile replacement",
                (1, 1),
                Operation::Set(vec![rule(Condition::DateIsYesterday)?]),
            ),
            ("clear", (1, 1), Operation::Clear),
        ] {
            run(
                "numbers_conditional_highlight_volatile_lifecycle_parity",
                Tier::B,
                Evidence::Queued,
                &first,
                cell,
                &operation,
            )
            .map_err(|error| format!("{path} {label}: {error}"))?;
        }
    }
    Ok(())
}

/// Numbers-saved tables register the conditional-style owner without the
/// volatile family: the host fails ("dependency tile is missing") and the
/// leaf refuses `UnsupportedDependency`.
#[test]
fn numbers_native_conditional_highlight_volatile_is_refused() -> TestResult {
    for path in [
        "iwork/numbers/basic.numbers",
        "iwork/numbers/custom-number-native.numbers",
    ] {
        let input = fixture(path)?;
        let operation = Operation::Set(vec![rule(Condition::DateIsToday)?]);
        let host = host_route(&input.bytes, (1, 1), &operation);
        assert!(
            host.is_err(),
            "{path}: the host registered the volatile family"
        );
        let package = litchi_numbers::Package::from_bytes(&input.bytes)?;
        let refused = package
            .edit_table_cell_conditional_highlights(0usize, 0usize)?
            .set(At::new(1, 1), [rule(Condition::DateIsToday)?])
            .and_then(|edit| edit.commit().map(|_| ()));
        assert!(
            matches!(
                refused,
                Err(litchi_numbers::cell::conditional_highlight::Error::Refused(
                    litchi_numbers::cell::conditional_highlight::Refusal::UnsupportedDependency
                ))
            ),
            "{path}: {refused:?}"
        );
    }
    Ok(())
}
