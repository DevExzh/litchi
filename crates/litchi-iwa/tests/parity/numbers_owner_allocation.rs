//! Documented intentional difference (record 0368a): the formula-owner
//! identifier the host allocates when it duplicates a Numbers table.
//!
//! The host takes one above the largest `FormulaOwnerDependenciesArchive`
//! identifier and ignores owner-map entries without an owner object. On
//! `chart-arrangement-native.numbers` and `shared-model-native.numbers` the
//! engine's owner map holds such entries, so a host duplicate registers an
//! internal identifier that is already in the map. The leaf allocator
//! (`litchi_numbers_wire::formula_graph::next_owner_id`) clears the map, so
//! the leaf route's output differs from the host's by design. This is not a
//! Tier A/B case: it pins the host defect on the host's own output and the
//! corrected leaf allocation on the same input. The route is queued for
//! native verification in `docs/unified-ops/0363-mac-queue.md`.

use std::collections::BTreeMap;

use litchi_iwa::numbers;
use litchi_iwa_protos::tsce;
use litchi_numbers_wire::formula_graph::{Limits, next_owner_id};
use prost::Message as _;

use crate::harness::{TestResult, fixture};

const CALCULATION_ENGINE: u32 = 4_000;
const FORMULA_OWNER: u32 = 4_008;

/// The engine payload and every owner payload of a package.
fn engine_and_owners(editor: &numbers::Draft) -> TestResult<(Vec<u8>, Vec<Vec<u8>>)> {
    let mut engine = None;
    let mut owners = Vec::new();
    for name in editor.package().iwa_entry_names() {
        for object in editor.package().archive(name)?.objects {
            for message in object.messages {
                match message.type_ {
                    CALCULATION_ENGINE => engine = Some(message.data),
                    FORMULA_OWNER => owners.push(message.data),
                    _ => {},
                }
            }
        }
    }
    Ok((engine.ok_or("no calculation engine")?, owners))
}

/// Internal identifiers registered more than once in the engine's owner map.
fn repeated_map_identifiers(engine: &[u8]) -> TestResult<Vec<u32>> {
    let engine = tsce::CalculationEngineArchive::decode(engine)?;
    let mut counts = BTreeMap::new();
    for entry in engine
        .dependency_tracker
        .owner_id_map
        .iter()
        .flat_map(|map| &map.map_entry)
    {
        *counts.entry(entry.internal_owner_id).or_insert(0usize) += 1;
    }
    Ok(counts
        .into_iter()
        .filter_map(|(identifier, count)| (count > 1).then_some(identifier))
        .collect())
}

fn mapped_identifiers(engine: &[u8]) -> TestResult<Vec<u32>> {
    let engine = tsce::CalculationEngineArchive::decode(engine)?;
    Ok(engine
        .dependency_tracker
        .owner_id_map
        .iter()
        .flat_map(|map| &map.map_entry)
        .map(|entry| entry.internal_owner_id)
        .collect())
}

/// Returns (host duplicates that succeeded, those with a repeated map entry).
fn intentional_difference(path: &'static str) -> TestResult<(usize, usize)> {
    let input = fixture(path)?;
    let editor = numbers::Draft::from_bytes(&input.bytes)?;
    let (engine, owners) = engine_and_owners(&editor)?;
    assert!(
        repeated_map_identifiers(&engine)?.is_empty(),
        "{path}: input map is valid"
    );

    // Leaf: the map-aware allocation is above every registered identifier.
    let owner_slices = owners.iter().map(Vec::as_slice).collect::<Vec<_>>();
    let leaf_next = next_owner_id(&engine, &owner_slices, &Limits::DEFAULT)?;
    assert!(
        mapped_identifiers(&engine)?
            .iter()
            .all(|identifier| *identifier < leaf_next)
    );

    // Host: some table duplicate registers an identifier the map already has.
    let mut host_duplicates = 0usize;
    let mut host_defects = 0usize;
    // `tables()` may refuse a fixture the duplicate route accepts; try the
    // first table positions directly.
    for index in 0..8 {
        let mut host = numbers::Draft::from_bytes(&input.bytes)?;
        if host
            .duplicate_table(litchi_numbers::TableSelector::index(index))
            .is_err()
        {
            continue;
        }
        host_duplicates += 1;
        let output = host.to_bytes()?;
        let (host_engine, _) = engine_and_owners(&numbers::Draft::from_bytes(&output)?)?;
        let repeated = repeated_map_identifiers(&host_engine)?;
        if !repeated.is_empty() {
            assert!(repeated.iter().all(|identifier| *identifier < leaf_next));
            host_defects += 1;
        }
    }
    Ok((host_duplicates, host_defects))
}

#[test]
fn numbers_table_duplicate_owner_allocation_differs_chart_arrangement() -> TestResult {
    let (duplicates, defects) =
        intentional_difference("iwork/numbers/chart-arrangement-native.numbers")?;
    assert!(
        duplicates > 0 && defects > 0,
        "the host defect no longer reproduces"
    );
    Ok(())
}

#[test]
fn numbers_table_duplicate_owner_allocation_differs_shared_model() -> TestResult {
    // The host's duplicate route refuses this fixture earlier (its
    // table-appearance read), so only its clone primitive collides
    // (`formula_clone::exit_parity::leaf_allocator_avoids_the_host_owner_map_collision`);
    // the leaf allocation is still checked against the map above.
    let (duplicates, defects) =
        intentional_difference("iwork/numbers/shared-model-native.numbers")?;
    assert_eq!((duplicates, defects), (0, 0));
    Ok(())
}
