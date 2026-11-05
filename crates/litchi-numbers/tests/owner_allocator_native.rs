//! The map-aware internal-owner allocator on every tracked Numbers workbook.
//!
//! A workbook's owner map can hold entries with no owner object. The retired
//! host allocated one above the largest owner object, so on some workbooks it
//! took an identifier already in the map, a corrupt map (record 0368a). This is
//! the leaf half of its `leaf_allocator_avoids_the_host_owner_map_collision`
//! (at `4d8a3e161f`): `next_owner_id` is above the map and the owners, the
//! allocation is valid, and the colliding identifier is refused.

use std::collections::HashSet;
use std::path::Path;

use litchi_iwa_archive::iwa::{Archive, SnappyStream};
use litchi_iwa_archive::package::Catalog;
use litchi_iwa_protos::tsce;
use litchi_numbers_wire::formula_graph::{self, Error, Limits, next_owner_id};
use litchi_numbers_wire::formula_owner::Uuid;
use prost::Message;

const ENGINE_MESSAGE: u32 = 4_000;
const OWNER_MESSAGE: u32 = 4_008;
/// Unused by any tracked map.
const NEW_OWNER_UUID: Uuid = Uuid::new(0x0123_4567_89ab_cdef, 0xfedc_ba98_7654_3210);
const NEW_OWNER_OBJECT: u64 = 0x7fff_0000;

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

/// The engine payload and owner payloads of a workbook with an engine.
fn engine_of(bytes: &[u8]) -> TestResult<Option<(Vec<u8>, Vec<Vec<u8>>)>> {
    let catalog = Catalog::from_bytes(bytes, Default::default())?;
    for entry in catalog
        .iter()
        .filter(|entry| entry.name().ends_with(".iwa"))
    {
        let stream = SnappyStream::decompress(entry.data(), Default::default())?;
        let archive = Archive::parse(stream.as_bytes(), Default::default())?;
        let of_type = |wanted: u32| -> Vec<Vec<u8>> {
            archive
                .objects
                .iter()
                .flat_map(|object| &object.messages)
                .filter(|message| message.type_ == wanted)
                .map(|message| message.data.clone())
                .collect()
        };
        if let [engine] = of_type(ENGINE_MESSAGE).as_slice() {
            return Ok(Some((engine.clone(), of_type(OWNER_MESSAGE))));
        }
    }
    Ok(None)
}

#[test]
fn allocator_stays_above_the_owner_map_on_every_workbook() -> TestResult {
    let root = Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data/iwork/numbers");
    let mut paths = std::fs::read_dir(root)?
        .map(|entry| Ok(entry?.path()))
        .collect::<std::io::Result<Vec<_>>>()?;
    paths.retain(|path| path.extension().is_some_and(|ext| ext == "numbers"));
    paths.sort();
    let mut engines = 0;
    let mut colliding = Vec::new();
    for path in paths {
        if !path.is_file() {
            continue;
        }
        let name = path
            .file_name()
            .ok_or("no file name")?
            .to_string_lossy()
            .into_owned();
        let Some((engine, owners)) = engine_of(&std::fs::read(&path)?)? else {
            continue;
        };
        engines += 1;
        let decoded = tsce::CalculationEngineArchive::decode(engine.as_slice())?;
        let mapped = decoded
            .dependency_tracker
            .owner_id_map
            .iter()
            .flat_map(|map| &map.map_entry)
            .map(|entry| entry.internal_owner_id)
            .collect::<HashSet<_>>();
        let largest_owner = owners
            .iter()
            .map(|owner| {
                Ok(
                    tsce::FormulaOwnerDependenciesArchive::decode(owner.as_slice())?
                        .internal_formula_owner_id,
                )
            })
            .collect::<Result<Vec<_>, prost::DecodeError>>()?
            .into_iter()
            .max()
            .unwrap_or(0);
        let owner_slices = owners.iter().map(Vec::as_slice).collect::<Vec<_>>();
        let next = next_owner_id(&engine, &owner_slices, &Limits::DEFAULT)?;
        assert!(next > largest_owner, "{name}");
        assert!(mapped.iter().all(|mapped| *mapped < next), "{name}");
        // The allocation is admitted and stays decodable.
        let appended = formula_graph::append_owners_to_engine(
            &engine,
            &[NEW_OWNER_OBJECT],
            &[(next, NEW_OWNER_UUID)],
            0,
            &Limits::DEFAULT,
        )?;
        tsce::CalculationEngineArchive::decode(appended.as_slice())?;
        // One above the largest owner object: the host's allocation.
        let host_rule = largest_owner + 1;
        if mapped.contains(&host_rule) {
            assert_ne!(host_rule, next, "{name}");
            assert_eq!(
                formula_graph::append_owners_to_engine(
                    &engine,
                    &[NEW_OWNER_OBJECT],
                    &[(host_rule, NEW_OWNER_UUID)],
                    0,
                    &Limits::DEFAULT,
                ),
                Err(Error::InvalidRemap),
                "{name}"
            );
            colliding.push(name);
        } else {
            assert_eq!(host_rule, next, "{name}");
        }
    }
    assert!(engines > 10, "{engines} engines");
    assert_eq!(colliding, COLLIDING);
    Ok(())
}

/// Workbooks where one above the largest owner object is already in the map.
/// The deleted host test listed the second, third and fifth: it looked only at
/// workbooks with a table family it could clone, and this scan has no such filter.
const COLLIDING: [&str; 5] = [
    "cell-value-native.numbers",
    "chart-arrangement-native.numbers",
    "shared-model-native.numbers",
    "table-data-list-native.numbers",
    "volatile-highlight-native.numbers",
];
