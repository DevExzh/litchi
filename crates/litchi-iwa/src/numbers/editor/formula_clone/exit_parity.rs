//! Host-versus-leaf differential parity for the formula-graph routes that
//! record 0368 moved into `litchi_numbers_wire::{formula_graph, formula_owner}`.
//!
//! Every tracked Numbers fixture is opened by the host; for each table's
//! formula-owner family the test builds the remap the host's
//! `clone_table_formula_graph` builds, runs the host wire rewrite and the
//! leaf rewrite on the same source bytes, and requires byte-identical
//! output (Tier A: both routes patch wire bytes in place). The empty table
//! owner the host encodes with Prost must equal the leaf encoder's bytes.
//! Deleted with the host (W3).

#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    clippy::panic,
    reason = "differential parity tests fail loudly on any divergence"
)]

use std::path::{Path, PathBuf};

use litchi_numbers_wire::formula_graph::{self as leaf, Limits, Remap};
use litchi_numbers_wire::formula_owner::{self as leaf_owner, Uuid as LeafUuid};

use super::*;

fn fixtures() -> Vec<PathBuf> {
    let root = Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data/iwork/numbers");
    let mut paths = std::fs::read_dir(&root)
        .unwrap()
        .map(|entry| entry.unwrap().path())
        .filter(|path| {
            path.extension()
                .is_some_and(|extension| extension == "numbers")
        })
        .collect::<Vec<_>>();
    paths.sort();
    paths
}

fn leaf_uuid(uuid: &tsp::Uuid) -> LeafUuid {
    LeafUuid::new(uuid.lower, uuid.upper)
}

/// One table family's host remap, also in the sorted-slice leaf form.
pub(super) struct FamilyCase {
    pub(super) owners: Vec<FormulaOwnerSource>,
    pub(super) tiles: Vec<CellTileSource>,
    pub(super) range_tiles: Vec<RangeTileSource>,
    pub(super) new_table_info_id: u64,
    pub(super) object_remap: HashMap<u64, u64>,
    pub(super) internal_remap: HashMap<u32, u32>,
    pub(super) uuid_remap: HashMap<(u64, u64), tsp::Uuid>,
    pub(super) object_ids: Vec<(u64, u64)>,
    pub(super) owner_ids: Vec<(u32, u32)>,
    pub(super) uuids: Vec<(LeafUuid, LeafUuid)>,
}

impl FamilyCase {
    pub(super) fn remap(&self) -> Remap<'_> {
        Remap::new(&self.object_ids, &self.owner_ids, &self.uuids)
    }
}

/// Build the remap `clone_table_formula_graph` builds for `table_info_id`,
/// with a fixed replacement table UUID and identifiers above the archive's.
///
/// `host_allocation` selects the host's internal-owner allocator (one above
/// the largest owner object); otherwise the allocator also clears every
/// owner-map entry, the policy of `litchi_numbers_wire::table_create`.
pub(super) fn family_case(
    archive: &Archive,
    table_info_id: u64,
    host_allocation: bool,
) -> Option<FamilyCase> {
    let (owners, source_owner_uuid) = formula_owner_family(archive, table_info_id).ok()?;
    if owners
        .iter()
        .any(|source| require_cloneable_formula_owner(&source.owner).is_err())
    {
        return None;
    }
    let tiles = cell_record_tiles(archive, &owners).ok()?;
    let range_tiles = range_precedent_tiles(archive, &owners).ok()?;
    let mut next_identifier = archive
        .objects
        .iter()
        .filter_map(|object| object.archive_info.identifier)
        .max()?
        + 1;
    let new_table_info_id = next_identifier;
    next_identifier += 1;
    let mut object_remap = HashMap::new();
    object_remap.insert(table_info_id, new_table_info_id);
    for identifier in owners
        .iter()
        .map(|source| source.object.archive_info.identifier)
        .chain(
            tiles
                .iter()
                .map(|source| source.object.archive_info.identifier),
        )
        .chain(
            range_tiles
                .iter()
                .map(|source| source.object.archive_info.identifier),
        )
    {
        object_remap.insert(identifier?, next_identifier);
        next_identifier += 1;
    }
    let mut next_internal = next_internal_owner_id(archive).ok()?;
    if !host_allocation {
        let engine =
            tsce::CalculationEngineArchive::decode(engine_payload(archive).as_slice()).ok()?;
        let mapped = engine
            .dependency_tracker
            .owner_id_map
            .iter()
            .flat_map(|map| &map.map_entry)
            .map(|entry| entry.internal_owner_id)
            .max()
            .unwrap_or(0);
        next_internal = next_internal.max(mapped + 1);
    }
    let mut internal_remap = HashMap::new();
    for source in &owners {
        internal_remap.insert(source.owner.internal_formula_owner_id, next_internal);
        next_internal += 1;
    }
    // The source table UUID is the inverse of the owner derivation.
    let source_uuid = formula_owner_uuid_for_table(&source_owner_uuid);
    let new_uuid = tsp::Uuid {
        lower: 0x0123_4567_89ab_cdef,
        upper: 0xfedc_ba98_7654_3210,
    };
    let new_owner_uuid = formula_owner_uuid_for_table(&new_uuid);
    let mut uuid_remap = HashMap::new();
    uuid_remap.insert(uuid_key(&source_uuid), new_uuid);
    uuid_remap.insert(uuid_key(&source_owner_uuid), new_owner_uuid);
    for source in &owners {
        let replacement = translate_owner_uuid(
            &source.owner.formula_owner_uid,
            &source_owner_uuid,
            &new_owner_uuid,
        );
        uuid_remap
            .entry(uuid_key(&source.owner.formula_owner_uid))
            .or_insert(replacement);
    }
    let mut object_ids = object_remap
        .iter()
        .map(|(a, b)| (*a, *b))
        .collect::<Vec<_>>();
    object_ids.sort_unstable();
    let mut owner_ids = internal_remap
        .iter()
        .map(|(a, b)| (*a, *b))
        .collect::<Vec<_>>();
    owner_ids.sort_unstable();
    let mut uuids = uuid_remap
        .iter()
        .map(|(key, value)| (LeafUuid::new(key.0, key.1), leaf_uuid(value)))
        .collect::<Vec<_>>();
    uuids.sort_unstable();
    Some(FamilyCase {
        owners,
        tiles,
        range_tiles,
        new_table_info_id,
        object_remap,
        internal_remap,
        uuid_remap,
        object_ids,
        owner_ids,
        uuids,
    })
}

/// Every table-info identifier with a direct formula owner.
pub(super) fn direct_owner_tables(archive: &Archive) -> Vec<u64> {
    let mut tables = archive
        .objects
        .iter()
        .flat_map(|object| &object.messages)
        .filter(|message| message.type_ == FORMULA_OWNER_MESSAGE_TYPE)
        .filter_map(|message| {
            tsce::FormulaOwnerDependenciesArchive::decode(message.data.as_slice()).ok()
        })
        .filter(|owner| {
            owner.owner_kind == Some(crate::numbers::formula_owner::TABLE_FORMULA_OWNER_KIND)
        })
        .filter_map(|owner| owner.formula_owner.map(|reference| reference.identifier))
        .collect::<Vec<_>>();
    tables.sort_unstable();
    tables.dedup();
    tables
}

/// The calculation-engine archive of a fixture, if it has one.
pub(super) fn engine_archive(path: &Path) -> Option<Archive> {
    let package = HostPackage::from_bytes(&std::fs::read(path).unwrap()).ok()?;
    let entry = package.calculation_engine_entry_name().ok()??.to_owned();
    package.archive(&entry).ok()
}

/// Host route: the per-owner wire rewrite of `clone_table_formula_graph`.
pub(super) fn host_owner(case: &FamilyCase, source: &FormulaOwnerSource) -> Vec<u8> {
    let mut expected = source.owner.clone();
    remap_formula_owner(
        &mut expected,
        case.new_table_info_id,
        &case.object_remap,
        &case.internal_remap,
        &case.uuid_remap,
    );
    remap_formula_owner_wire(
        &source.object.messages[source.message_index].data,
        &source.owner,
        &expected,
        &case.object_remap,
        &case.internal_remap,
        &case.uuid_remap,
    )
    .unwrap()
}

pub(super) fn host_cell_tile(case: &FamilyCase, source: &CellTileSource) -> Vec<u8> {
    let mut expected = source.tile.clone();
    expected.internal_owner_id = case.internal_remap[&source.tile.internal_owner_id];
    remap_cell_records(&mut expected.cell_records, &case.internal_remap);
    remap_cell_tile_wire(
        &source.object.messages[source.message_index].data,
        &source.tile,
        &expected,
        &case.internal_remap,
    )
    .unwrap()
}

pub(super) fn host_range_tile(case: &FamilyCase, source: &RangeTileSource) -> Vec<u8> {
    let mut expected = source.tile.clone();
    expected.to_owner_id = case
        .internal_remap
        .get(&source.tile.to_owner_id)
        .copied()
        .unwrap_or(source.tile.to_owner_id);
    remap_range_tile_wire(
        &source.object.messages[source.message_index].data,
        &source.tile,
        &expected,
    )
    .unwrap()
}

/// The cloned owner identifiers and owner-map entries of a family, in the
/// host's append order.
pub(super) fn engine_inputs(case: &FamilyCase) -> (Vec<u64>, Vec<(u32, LeafUuid)>, u64) {
    let mut owners = Vec::new();
    let mut entries = Vec::new();
    let mut formula_count = 0u64;
    for source in &case.owners {
        owners.push(case.object_remap[&source.object.archive_info.identifier.unwrap()]);
        let uuid = case
            .uuid_remap
            .get(&uuid_key(&source.owner.formula_owner_uid))
            .copied()
            .unwrap_or(source.owner.formula_owner_uid);
        entries.push((
            case.internal_remap[&source.owner.internal_formula_owner_id],
            leaf_uuid(&uuid),
        ));
        formula_count += formula_cell_count(&source.owner, &case.tiles).unwrap();
    }
    (owners, entries, formula_count)
}

pub(super) fn host_engine(engine: &[u8], case: &FamilyCase) -> Vec<u8> {
    let (owners, entries, formula_count) = engine_inputs(case);
    let host_entries = entries
        .iter()
        .map(
            |(internal, uuid)| tsce::owner_id_map_archive::OwnerIdMapArchiveEntry {
                internal_owner_id: *internal,
                owner_id: uuid_as_cfuuid(&tsp::Uuid {
                    lower: uuid.lower(),
                    upper: uuid.upper(),
                }),
            },
        )
        .collect::<Vec<_>>();
    append_formula_owners_to_engine(engine, &owners, &host_entries, formula_count).unwrap()
}

pub(super) fn leaf_engine(
    engine: &[u8],
    case: &FamilyCase,
) -> std::result::Result<Vec<u8>, leaf::Error> {
    let (owners, entries, formula_count) = engine_inputs(case);
    leaf::append_owners_to_engine(engine, &owners, &entries, formula_count, &Limits::DEFAULT)
}

pub(super) fn engine_payload(archive: &Archive) -> Vec<u8> {
    let (identifier, index) = calculation_engine_location(archive).unwrap();
    archive.object(identifier).unwrap().messages[index]
        .data
        .clone()
}

#[derive(Debug, Default)]
struct Tally {
    families: usize,
    owners: usize,
    cell_tiles: usize,
    range_tiles: usize,
    engines: usize,
    leaf_refusals: Vec<String>,
}

#[test]
fn formula_graph_rewrites_are_byte_identical_to_the_host() {
    let mut tally = Tally::default();
    for path in fixtures() {
        let Some(archive) = engine_archive(&path) else {
            continue;
        };
        let name = path.file_name().unwrap().to_string_lossy().into_owned();
        let engine = engine_payload(&archive);
        for table in direct_owner_tables(&archive) {
            let Some(case) = family_case(&archive, table, false) else {
                continue;
            };
            tally.families += 1;
            for source in &case.owners {
                let original = &source.object.messages[source.message_index].data;
                match leaf::remap_owner(original, case.remap(), &Limits::DEFAULT) {
                    Ok(bytes) => {
                        assert_eq!(
                            bytes,
                            host_owner(&case, source),
                            "{name} table {table} owner"
                        );
                        tally.owners += 1;
                    },
                    Err(error) => tally
                        .leaf_refusals
                        .push(format!("{name} table {table} owner: {error}")),
                }
            }
            for source in &case.tiles {
                let original = &source.object.messages[source.message_index].data;
                let bytes = leaf::remap_cell_record_tile(original, case.remap(), &Limits::DEFAULT)
                    .unwrap_or_else(|error| panic!("{name} table {table} cell tile: {error}"));
                assert_eq!(
                    bytes,
                    host_cell_tile(&case, source),
                    "{name} table {table} cell tile"
                );
                tally.cell_tiles += 1;
            }
            for source in &case.range_tiles {
                let original = &source.object.messages[source.message_index].data;
                let bytes =
                    leaf::remap_range_precedents_tile(original, case.remap(), &Limits::DEFAULT)
                        .unwrap_or_else(|error| panic!("{name} table {table} range tile: {error}"));
                assert_eq!(
                    bytes,
                    host_range_tile(&case, source),
                    "{name} table {table} range tile"
                );
                tally.range_tiles += 1;
            }
            match leaf_engine(&engine, &case) {
                Ok(bytes) => {
                    assert_eq!(
                        bytes,
                        host_engine(&engine, &case),
                        "{name} table {table} engine"
                    );
                    tally.engines += 1;
                },
                Err(error) => tally
                    .leaf_refusals
                    .push(format!("{name} table {table} engine: {error}")),
            }
        }
    }
    eprintln!("formula graph parity: {tally:?}");
    assert!(
        tally.families > 0 && tally.owners > 0 && tally.engines > 0,
        "{tally:?}"
    );
    assert!(
        tally.leaf_refusals.is_empty(),
        "leaf refused host-accepted inputs: {tally:?}"
    );
}

#[test]
fn empty_table_owner_is_byte_identical_to_the_host() {
    for (lower, upper, info, internal) in [
        (
            0xa1f5_067b_8152_4a6a_u64,
            0x68b3_84c7_9e7f_4b90_u64,
            0x1234_u64,
            7_u32,
        ),
        (1, 0, 1, 1),
        (u64::MAX, u64::MAX, u64::MAX, u32::MAX),
        (0x0011_2233_4455_6677, 0x8899_aabb_ccdd_eeff, 300, 128),
    ] {
        let table = tsp::Uuid { lower, upper };
        let host = empty_table_formula_owner(&table, info, internal).encode_to_vec();
        let leaf = leaf_owner::encode_empty_table(
            LeafUuid::new(lower, upper),
            info,
            internal,
            &litchi_numbers_wire::Limits::DEFAULT,
        )
        .unwrap();
        assert_eq!(leaf, host, "{table:?} {info} {internal}");
        let host_uuid = formula_owner_uuid_for_table(&table);
        assert_eq!(
            LeafUuid::new(lower, upper).for_table(),
            LeafUuid::new(host_uuid.lower, host_uuid.upper)
        );
        let words = uuid_as_cfuuid(&table);
        assert_eq!(
            LeafUuid::new(lower, upper).cfuuid_words(),
            [
                words.uuid_w0.unwrap(),
                words.uuid_w1.unwrap(),
                words.uuid_w2.unwrap(),
                words.uuid_w3.unwrap()
            ]
        );
        let source = tsp::Uuid {
            lower: 0x1000,
            upper: 0x2000,
        };
        let derived = translate_owner_uuid(&table, &source, &host_uuid);
        assert_eq!(
            LeafUuid::new(lower, upper).translate_from(
                LeafUuid::new(0x1000, 0x2000),
                LeafUuid::new(host_uuid.lower, host_uuid.upper)
            ),
            LeafUuid::new(derived.lower, derived.upper)
        );
    }
}

/// The host allocates internal owner identifiers one above the largest owner
/// object and ignores owner-map entries without an owner object, so on these
/// fixtures it registers duplicate internal identifiers in the engine's owner
/// map. The leaf refuses that remap (`InvalidRemap`) instead of publishing a
/// corrupt map; record 0368 lists the case (stricter admission, §4.2 of 0363).
#[test]
fn leaf_refuses_the_host_owner_map_collision() {
    let mut refused = Vec::new();
    for path in fixtures() {
        let Some(archive) = engine_archive(&path) else {
            continue;
        };
        let engine = engine_payload(&archive);
        for table in direct_owner_tables(&archive) {
            let Some(case) = family_case(&archive, table, true) else {
                continue;
            };
            let decoded = tsce::CalculationEngineArchive::decode(engine.as_slice()).unwrap();
            let mapped = decoded
                .dependency_tracker
                .owner_id_map
                .iter()
                .flat_map(|map| &map.map_entry)
                .map(|entry| entry.internal_owner_id)
                .collect::<HashSet<_>>();
            let collides = case.owner_ids.iter().any(|(_, new)| mapped.contains(new));
            let result = leaf_engine(&engine, &case);
            if collides {
                assert_eq!(result, Err(leaf::Error::InvalidRemap));
                refused.push(path.file_name().unwrap().to_string_lossy().into_owned());
            } else {
                assert_eq!(result.unwrap(), host_engine(&engine, &case));
            }
        }
    }
    refused.sort();
    assert_eq!(
        refused,
        [
            "chart-arrangement-native.numbers",
            "shared-model-native.numbers"
        ]
    );
}

/// The host route decodes every owner and tile with Prost while it collects
/// the family (`formula_owner_family`, `cell_record_tiles`,
/// `range_precedent_tiles`) before it rewrites; the probe charges that decode.
#[inline(never)]
fn host_route(cases: &[(Vec<u8>, FamilyCase)]) -> usize {
    let mut bytes = 0;
    for (engine, case) in cases {
        for source in &case.owners {
            let data = &source.object.messages[source.message_index].data;
            let owner = tsce::FormulaOwnerDependenciesArchive::decode(data.as_slice()).unwrap();
            std::hint::black_box(&owner);
            bytes += host_owner(case, source).len();
        }
        for source in &case.tiles {
            let data = &source.object.messages[source.message_index].data;
            let tile = tsce::CellRecordTileArchive::decode(data.as_slice()).unwrap();
            std::hint::black_box(&tile);
            bytes += host_cell_tile(case, source).len();
        }
        for source in &case.range_tiles {
            let data = &source.object.messages[source.message_index].data;
            let tile = tsce::RangePrecedentsTileArchive::decode(data.as_slice()).unwrap();
            std::hint::black_box(&tile);
            bytes += host_range_tile(case, source).len();
        }
        bytes += host_engine(engine, case).len();
    }
    bytes
}

#[inline(never)]
fn leaf_route(cases: &[(Vec<u8>, FamilyCase)]) -> usize {
    let mut bytes = 0;
    for (engine, case) in cases {
        for source in &case.owners {
            let original = &source.object.messages[source.message_index].data;
            bytes += leaf::remap_owner(original, case.remap(), &Limits::DEFAULT)
                .unwrap()
                .len();
        }
        for source in &case.tiles {
            let original = &source.object.messages[source.message_index].data;
            bytes += leaf::remap_cell_record_tile(original, case.remap(), &Limits::DEFAULT)
                .unwrap()
                .len();
        }
        for source in &case.range_tiles {
            let original = &source.object.messages[source.message_index].data;
            bytes += leaf::remap_range_precedents_tile(original, case.remap(), &Limits::DEFAULT)
                .unwrap()
                .len();
        }
        bytes += leaf_engine(engine, case).unwrap().len();
    }
    bytes
}

/// Paired cost probe for record 0368 (run under callgrind with
/// `--toggle-collect` on `host_route`/`leaf_route`): every fixture family's
/// owner, tile and engine rewrites, in ABBA order.
#[test]
#[ignore = "performance probe; run explicitly under callgrind"]
fn formula_graph_route_cost_abba() {
    let mut cases = Vec::new();
    for path in fixtures() {
        let Some(archive) = engine_archive(&path) else {
            continue;
        };
        let engine = engine_payload(&archive);
        for table in direct_owner_tables(&archive) {
            if let Some(case) = family_case(&archive, table, false) {
                cases.push((engine.clone(), case));
            }
        }
    }
    let a1 = host_route(&cases);
    let b1 = leaf_route(&cases);
    let b2 = leaf_route(&cases);
    let a2 = host_route(&cases);
    assert_eq!((a1, a2), (b1, b2));
}
