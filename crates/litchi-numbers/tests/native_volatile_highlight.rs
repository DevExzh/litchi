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
use litchi_iwa_protos::{tsce, tst};
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
        count.ok_or_else(|| io::Error::other("no formula count"))?,
    ))
}

/// The engine's `uuid_reference_map`, as `(table UUID lower, owner, cells)`.
fn engine_uuid_map(source: &[u8]) -> TestResult<Vec<(u64, u32, Vec<(u32, Vec<i32>)>)>> {
    let mut output = Vec::new();
    for object in &engine(source)?.objects {
        for message in &object.messages {
            if message.type_ != 4_000 {
                continue;
            }
            let archive = tsce::CalculationEngineArchive::decode(message.data.as_slice())?;
            for entry in archive
                .uuid_reference_map
                .iter()
                .flat_map(|map| &map.cell_refs_for_uuid)
            {
                for owner in entry.cell_refs.iter().flat_map(|refs| &refs.owner_entries) {
                    output.push((
                        entry.uuid.lower,
                        owner.owner_id,
                        owner
                            .coord_set
                            .column_entries
                            .iter()
                            .map(|column| {
                                (
                                    column.column,
                                    column
                                        .row_set
                                        .entries
                                        .iter()
                                        .flat_map(|range| {
                                            range.range_begin
                                                ..=range.range_end.unwrap_or(range.range_begin)
                                        })
                                        .collect(),
                                )
                            })
                            .collect(),
                    ));
                }
            }
        }
    }
    Ok(output)
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

/// The rule Numbers itself added to a cell that had one (record 0433): A2 holds
/// tomorrow's date with "today" and "tomorrow", A3 and A4 keep "today" alone. The
/// second rule's preset fill is read, not assumed.
#[test]
fn native_second_rule_reads_in_order() -> TestResult {
    let path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/numbers/volatile-highlight-native-second.numbers");
    let package = Package::open(path)?;
    let first = native_rule()?;
    let rules = package
        .table_cell_conditional_highlight(0usize, 0usize, At::new(1, 0))?
        .ok_or_else(|| io::Error::other("no rule on A2"))?;
    assert_eq!(rules.len(), 2);
    assert_eq!(rules[0], first);
    assert!(format!("{:?}", rules[1]).contains("DateIsTomorrow"));
    for row in [2, 3] {
        assert_eq!(
            package.table_cell_conditional_highlight(0usize, 0usize, At::new(row, 0))?,
            Some(vec![first.clone()]),
            "row {row}"
        );
    }
    // The volatile set does not grow: A2 was registered by the first rule.
    let (cells, count) = volatile_state(&exact_bytes(&package)?)?;
    assert_eq!(cells, [(0, vec![1, 2, 3])]);
    assert_eq!(count, 8);
    Ok(())
}

/// The two rules the leaf writes on A2 hold the predicates Numbers wrote there:
/// a rule set of its own, the same conditions, and no new volatile cell.
#[test]
fn leaf_second_rule_matches_what_numbers_wrote() -> TestResult {
    let second = std::fs::read(
        PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/iwork/numbers/volatile-highlight-native-second.numbers"),
    )?;
    let at = At::new(1, 0);
    let rules = Package::from_bytes(&second)?
        .table_cell_conditional_highlight(0usize, 0usize, at)?
        .ok_or_else(|| io::Error::other("no rule on A2"))?;
    let package = Package::open(fixture())?;
    let bytes = exact_bytes(
        package
            .edit_table_cell_conditional_highlights(0usize, 0usize)?
            .set(at, rules.clone())?
            .commit()?
            .package(),
    )?;
    assert_eq!(
        Package::from_bytes(&bytes)?.table_cell_conditional_highlight(0usize, 0usize, at)?,
        Some(rules)
    );
    assert_eq!(two_rule_predicates(&bytes)?, two_rule_predicates(&second)?);
    let (cells, count) = volatile_state(&bytes)?;
    assert_eq!(cells, [(0, vec![1, 2, 3])]);
    assert_eq!(count, 8);
    Ok(())
}

/// The predicates of every two-rule set of the file's conditional-style list.
fn two_rule_predicates(source: &[u8]) -> TestResult<Vec<tst::FormulaPredicateArchive>> {
    let catalog = Catalog::from_bytes(source, Default::default())?;
    let mut output = Vec::new();
    for entry in catalog.iter() {
        if !entry.name().ends_with(".iwa") {
            continue;
        }
        let stream = SnappyStream::decompress(entry.data(), Default::default())?;
        let archive = Archive::parse(stream.as_bytes(), Default::default())?;
        for message in archive.objects.iter().flat_map(|object| &object.messages) {
            if message.type_ == 6_010 {
                let set = tst::ConditionalStyleSetArchive::decode(message.data.as_slice())?;
                let rules = set.rules.map(|rules| rules.rule).unwrap_or_default();
                if rules.len() == 2 {
                    output.extend(rules.into_iter().filter_map(|rule| rule.predicate));
                }
            }
        }
    }
    Ok(output)
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
    let map = engine_uuid_map(&bytes)?;
    assert_eq!(map.len(), 1);
    assert_eq!(map[0].2, [(0, vec![1, 2, 3]), (1, vec![2])]);

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
    assert_eq!(
        engine_uuid_map(&bytes)?,
        engine_uuid_map(&std::fs::read(fixture())?)?
    );
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
    assert_eq!(engine_uuid_map(&bytes)?[0].2, [(0, vec![1, 3])]);
    Ok(())
}

/// The file before Numbers wrote its rule is not available, but `basic.numbers`
/// registers the same 13 owners; the first volatile rule gives the
/// conditional-style owner the set, UUID reference and tile Numbers gives it.
#[test]
fn first_volatile_rule_on_a_numbers_saved_table_builds_what_numbers_builds() -> TestResult {
    let basic = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/numbers/basic.numbers");
    let package = Package::open(basic)?;
    let source = exact_bytes(&package)?;
    let at = At::new(1, 1);
    let rule = fill_rule(Condition::DateIsToday)?;
    let commit = package
        .edit_table_cell_conditional_highlights(0usize, 0usize)?
        .set(at, [rule.clone()])?
        .commit()?;
    let bytes = exact_bytes(commit.package())?;
    let reopened = Package::from_bytes(&bytes)?;
    assert_eq!(
        reopened.table_cell_conditional_highlight(0usize, 0usize, at)?,
        Some(vec![rule])
    );
    let owner_of =
        |source: &[u8], kind: u32| -> TestResult<tsce::FormulaOwnerDependenciesArchive> {
            for object in &engine(source)?.objects {
                for message in &object.messages {
                    if message.type_ == 4_008 {
                        let owner =
                            tsce::FormulaOwnerDependenciesArchive::decode(message.data.as_slice())?;
                        if owner.owner_kind == Some(kind) {
                            return Ok(owner);
                        }
                    }
                }
            }
            Err(io::Error::other("no owner").into())
        };
    // The same 13 owners, and only the conditional-style owner changed.
    for kind in [1, 4, 5, 6, 7, 8, 9, 10, 11, 12, 35, 200] {
        assert_eq!(owner_of(&source, kind)?, owner_of(&bytes, kind)?, "{kind}");
    }
    // The shape Numbers wrote for its own first rule.
    let native = owner_of(&std::fs::read(fixture())?, 3)?;
    let written = owner_of(&bytes, 3)?;
    let shape = |owner: &tsce::FormulaOwnerDependenciesArchive| {
        (
            owner.cell_dependencies.as_ref().map(|cells| {
                cells.cell_record[0]
                    .expanded_edges
                    .as_ref()
                    .map(|edges| edges.edge_with_owner_rows.len())
            }),
            owner
                .tiled_cell_dependencies
                .as_ref()
                .map(|tiled| tiled.cell_record_tiles.len()),
            owner
                .uuid_references
                .as_ref()
                .map(|refs| refs.table_refs.len()),
        )
    };
    assert_eq!(shape(&written), shape(&native));
    let (cells, count) = volatile_state(&bytes)?;
    assert_eq!(cells, [(1, vec![1])]);
    assert_eq!(count, volatile_state_count(&source)? + 1);
    // Numbers lists the rule's cell under the table and the conditional-style
    // owner (internal 9 in both files); its re-save keeps whatever the file holds.
    let native_map = engine_uuid_map(&std::fs::read(fixture())?)?;
    assert_eq!(native_map, [(native_map[0].0, 9, vec![(0, vec![1, 2, 3])])]);
    assert_eq!(engine_uuid_map(&source)?, []);
    assert_eq!(
        engine_uuid_map(&bytes)?,
        [(native_map[0].0, 9, vec![(1, vec![1])])]
    );
    // Inverse.
    let restored = commit
        .package()
        .apply_table_cell_conditional_highlights(&commit.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, source);
    Ok(())
}

fn volatile_state_count(source: &[u8]) -> TestResult<u64> {
    for object in &engine(source)?.objects {
        for message in &object.messages {
            if message.type_ == 4_000 {
                let archive = tsce::CalculationEngineArchive::decode(message.data.as_slice())?;
                return archive
                    .dependency_tracker
                    .number_of_formulas
                    .ok_or_else(|| io::Error::other("no formula count").into());
            }
        }
    }
    Err(io::Error::other("no engine").into())
}
