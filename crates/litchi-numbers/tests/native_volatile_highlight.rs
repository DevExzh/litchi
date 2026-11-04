//! A volatile conditional highlight authored in Numbers 14.4 (record 0432).
//!
//! `volatile-highlight-native.numbers`: A2 today, A3 yesterday, A4 tomorrow as
//! authored on 2026-11-04, with the rule "Date is today" and the preset style
//! "Red Fill" on A2:A4.

use std::io;
use std::path::PathBuf;

use litchi_iwa_archive::package::Catalog;
use litchi_iwa_common::color::{RgbColorSpace, Rgba};
use litchi_iwa_core::{Archive, SnappyStream};
use litchi_iwa_protos::tsce;
use litchi_numbers::Package;
use litchi_numbers::cell::conditional_highlight::{Condition, Rule, Style};
use litchi_numbers::table::At;
use prost::Message as _;

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

fn fixture() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/numbers/volatile-highlight-native.numbers")
}

fn exact_bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

fn engine(source: &[u8]) -> TestResult<Archive> {
    let catalog = Catalog::from_bytes(source, Default::default())?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == "Index/CalculationEngine.iwa")
        .ok_or_else(|| io::Error::other("no calculation engine"))?;
    let stream = SnappyStream::decompress(entry.data(), Default::default())?;
    Ok(Archive::parse(stream.as_bytes(), Default::default())?)
}

/// The table's volatile-time cells as `(column, rows)`, the rows expanded, and
/// the engine's formula count.
fn volatile_state(source: &[u8]) -> TestResult<(Vec<(u32, Vec<i32>)>, u64)> {
    let engine = engine(source)?;
    let mut cells = None;
    let mut count = None;
    for object in &engine.objects {
        for message in &object.messages {
            if message.type_ == 4_000 {
                let archive = tsce::CalculationEngineArchive::decode(message.data.as_slice())?;
                count = archive.dependency_tracker.number_of_formulas;
            }
            if message.type_ == 4_008 {
                let owner = tsce::FormulaOwnerDependenciesArchive::decode(message.data.as_slice())?;
                if owner.owner_kind == Some(3) {
                    let set = owner
                        .volatile_dependencies
                        .and_then(|set| set.volatile_time_cells)
                        .unwrap_or_default();
                    cells = Some(
                        set.column_entries
                            .iter()
                            .map(|entry| {
                                let rows = entry
                                    .row_set
                                    .entries
                                    .iter()
                                    .flat_map(|range| {
                                        range.range_begin
                                            ..=range.range_end.unwrap_or(range.range_begin)
                                    })
                                    .collect();
                                (entry.column, rows)
                            })
                            .collect::<Vec<_>>(),
                    );
                }
            }
        }
    }
    Ok((
        cells.ok_or_else(|| io::Error::other("no conditional-style owner"))?,
        u64::from(count.ok_or_else(|| io::Error::other("no formula count"))?),
    ))
}

fn native_rule() -> TestResult<Rule> {
    Ok(Rule::new(
        Condition::DateIsToday,
        Style::new(
            Some(Rgba::new(
                1.0,
                0.592_156_9,
                0.505_882_4,
                0.9,
                RgbColorSpace::Srgb,
            )?),
            Some(Rgba::new(0.0, 0.0, 0.0, 1.0, RgbColorSpace::Srgb)?),
            false,
        )?,
    ))
}

fn fill_rule(condition: Condition) -> TestResult<Rule> {
    Ok(Rule::new(
        condition,
        Style::with_fill(Rgba::new(1.0, 0.0, 0.0, 1.0, RgbColorSpace::Srgb)?),
    ))
}

/// The leaf reads Numbers' own rule; A1 and A5 have none.
#[test]
fn native_rule_reads_on_every_cell_of_its_range() -> TestResult {
    let package = Package::open(fixture())?;
    let rule = native_rule()?;
    for row in 1..=3 {
        let rules = package.table_cell_conditional_highlight(0usize, 0usize, At::new(row, 0))?;
        assert_eq!(rules, Some(vec![rule.clone()]), "row {row}");
    }
    for row in [0, 4] {
        let rules = package.table_cell_conditional_highlight(0usize, 0usize, At::new(row, 0))?;
        assert_eq!(rules, None, "row {row}");
    }
    let (cells, count) = volatile_state(&exact_bytes(&package)?)?;
    assert_eq!(cells, [(0, vec![1, 2, 3])]);
    assert_eq!(count, 8);
    Ok(())
}

/// A new volatile rule registers in Numbers' family, and clearing it restores
/// the volatile cells and the formula count.
#[test]
fn native_family_takes_a_new_cell_and_gives_it_back() -> TestResult {
    let package = Package::open(fixture())?;
    let at = At::new(2, 1);
    let commit = package
        .edit_table_cell_conditional_highlights(0usize, 0usize)?
        .set(at, [fill_rule(Condition::DateIsTomorrow)?])?
        .commit()?;
    let bytes = exact_bytes(commit.package())?;
    let reopened = Package::from_bytes(&bytes)?;
    assert_eq!(
        reopened.table_cell_conditional_highlight(0usize, 0usize, at)?,
        Some(vec![fill_rule(Condition::DateIsTomorrow)?])
    );
    assert_eq!(
        reopened.table_cell_conditional_highlight(0usize, 0usize, At::new(1, 0))?,
        Some(vec![native_rule()?])
    );
    let (cells, count) = volatile_state(&bytes)?;
    assert_eq!(cells, [(0, vec![1, 2, 3]), (1, vec![2])]);
    assert_eq!(count, 9);

    let cleared = reopened
        .edit_table_cell_conditional_highlights(0usize, 0usize)?
        .clear(at)?
        .commit()?;
    let bytes = exact_bytes(cleared.package())?;
    let reopened = Package::from_bytes(&bytes)?;
    assert_eq!(
        reopened.table_cell_conditional_highlight(0usize, 0usize, at)?,
        None
    );
    let (cells, count) = volatile_state(&bytes)?;
    assert_eq!(cells, [(0, vec![1, 2, 3])]);
    assert_eq!(count, 8);
    Ok(())
}

/// Numbers stores A2:A4 as one range. Clearing A3 splits it; the other cells
/// keep their rule.
#[test]
fn clearing_inside_a_native_range_splits_it() -> TestResult {
    let package = Package::open(fixture())?;
    let cleared = package
        .edit_table_cell_conditional_highlights(0usize, 0usize)?
        .clear(At::new(2, 0))?
        .commit()?;
    let bytes = exact_bytes(cleared.package())?;
    let reopened = Package::from_bytes(&bytes)?;
    assert_eq!(
        reopened.table_cell_conditional_highlight(0usize, 0usize, At::new(2, 0))?,
        None
    );
    for row in [1, 3] {
        assert_eq!(
            reopened.table_cell_conditional_highlight(0usize, 0usize, At::new(row, 0))?,
            Some(vec![native_rule()?]),
            "row {row}"
        );
    }
    let (cells, count) = volatile_state(&bytes)?;
    assert_eq!(cells, [(0, vec![1, 3])]);
    assert_eq!(count, 7);
    Ok(())
}
