//! Pages section and body-table lifecycle parity cases of wave W1-B
//! (record 0371): the host's `{insert,append,remove}_section`, `sections`,
//! `add_table` and `remove_table` against the leaf `section::lifecycle`,
//! `table::creation` and `table::deletion` families.

use std::collections::BTreeMap;

use litchi_core::Position;
use litchi_iwa::pages;
use litchi_pages::table::creation::NewTable;
use litchi_pages::{BodyTableSelector, Package, SectionSelector};

use crate::harness::{Case, Evidence, Format, LeafRun, TestResult, Tier, fixture};
use crate::known_fields::{Decoded, decode_as, decoded};
use crate::leaf_bytes;
use prost::Message as _;

const SECTIONS: &str = "iwork/pages/body-sections-unicode.pages";
const TABLES: &str = "iwork/pages/body-table-catalog-native.pages";
const CELLS: &str = "iwork/pages/body-table-cells-native.pages";

fn host_sections(bytes: &[u8]) -> TestResult<Vec<(u32, Option<String>)>> {
    Ok(pages::Draft::from_bytes(bytes)?
        .sections()
        .iter()
        .map(|section| (section.character_index, section.name.clone()))
        .collect())
}

fn leaf_sections(bytes: &[u8]) -> TestResult<Vec<Option<String>>> {
    Ok(Package::from_bytes(bytes)?
        .sections()
        .iter()
        .map(|section| section.name().map(str::to_owned))
        .collect())
}

fn body_text(bytes: &[u8]) -> TestResult<String> {
    Ok(pages::Draft::from_bytes(bytes)?.body_text()?)
}

fn section_case(name: &'static str, tier: Tier, evidence: Evidence, remove: bool) -> TestResult {
    let input = fixture(SECTIONS)?;
    Case::new(name, Format::Pages, tier, evidence, &input)
        .host(|source| {
            let mut editor = pages::Draft::from_bytes(source)?;
            if remove {
                let id = editor
                    .sections()
                    .get(1)
                    .ok_or("no second section")?
                    .object_id;
                drop(editor.remove_section(id)?);
            } else {
                let id = editor.sections().last().ok_or("no section")?.object_id;
                drop(editor.append_section(id, "Appendix")?);
            }
            Ok(editor.to_bytes()?)
        })
        .leaf(|source| {
            let package = Package::from_bytes(source)?;
            let edit = package.edit_section_lifecycle();
            let edit = if remove {
                edit.remove(SectionSelector::index(1))?
            } else {
                let last = package.sections().len() - 1;
                edit.add(SectionSelector::index(last), Some("Appendix"))?
            };
            let commit = edit.commit()?;
            let restored = commit
                .package()
                .apply_section_lifecycle(&commit.patch().inverse())?;
            Ok(LeafRun::with_inverse(
                leaf_bytes!(commit.package()),
                leaf_bytes!(restored.package()),
            ))
        })
        .read("host sections", host_sections)
        .read("leaf sections", leaf_sections)
        .read("host body text", body_text)
        .leaf_structural_metadata()
        .run()?;
    Ok(())
}

#[test]
fn pages_section_append_parity() -> TestResult {
    section_case(
        "pages_section_append_parity",
        Tier::B,
        Evidence::Queued,
        false,
    )
}

#[test]
fn pages_section_remove_parity() -> TestResult {
    section_case(
        "pages_section_remove_parity",
        Tier::B,
        Evidence::Queued,
        true,
    )
}

fn host_tables(bytes: &[u8]) -> TestResult<Vec<(String, usize, usize, usize)>> {
    Ok(pages::Draft::from_bytes(bytes)?
        .tables()?
        .into_iter()
        .map(|table| {
            (
                table.name,
                table.anchor_character_index,
                table.rows,
                table.columns,
            )
        })
        .collect())
}

fn leaf_tables(bytes: &[u8]) -> TestResult<Vec<Option<String>>> {
    let package = Package::from_bytes(bytes)?;
    let catalog = package.body_tables()?;
    Ok((0..catalog.len())
        .map(|index| catalog.get(index).map(|entry| entry.name().to_owned()))
        .collect())
}

/// `TST.TableModelArchive` without its table UUID: a new table's UUID is
/// a fresh version-4 identity on the leaf and a model-derived one on the
/// host (record 0370f, a documented policy: compared by form).
fn model_without_table_uuid(data: &[u8]) -> Result<Decoded, prost::DecodeError> {
    let mut model = litchi_iwa_protos::tst::TableModelArchive::decode(data)?;
    assert_eq!(model.table_id.len(), 36, "table UUID form");
    model.table_id.clear();
    Ok(decoded(model, data))
}

/// `TSCE.FormulaOwnerDependenciesArchive` without its owner UUID (derived
/// from the table UUID).
fn owner_without_uuid(data: &[u8]) -> Result<Decoded, prost::DecodeError> {
    let mut owner = litchi_iwa_protos::tsce::FormulaOwnerDependenciesArchive::decode(data)?;
    owner.formula_owner_uid = Default::default();
    Ok(decoded(owner, data))
}

/// `TSCE.CalculationEngineArchive` without its owner map's UUIDs (one is
/// derived from the new table UUID).
fn engine_without_owner_uuids(data: &[u8]) -> Result<Decoded, prost::DecodeError> {
    let mut engine = litchi_iwa_protos::tsce::CalculationEngineArchive::decode(data)?;
    if let Some(map) = engine.dependency_tracker.owner_id_map.as_mut() {
        for entry in &mut map.map_entry {
            entry.owner_id = Default::default();
        }
    }
    Ok(decoded(engine, data))
}

/// Adding a 3×4 table at the start of the body (record 0371; the shared
/// `table_create` follows the host graph since record 0370f).
#[test]
fn pages_body_table_add_parity() -> TestResult {
    let input = fixture(TABLES)?;
    Case::new(
        "pages_body_table_add_parity",
        Format::Pages,
        Tier::B,
        Evidence::Queued,
        &input,
    )
    .leaf_structural_metadata()
    .known_fields(|known| {
        known
            .with(
                6_001,
                "TST.TableModelArchive (table UUID by form)",
                model_without_table_uuid,
            )
            .with(
                4_008,
                "TSCE.FormulaOwnerDependenciesArchive (owner UUID by form)",
                owner_without_uuid,
            )
            .with(
                4_000,
                "TSCE.CalculationEngineArchive (owner-map UUIDs by form)",
                engine_without_owner_uuids,
            )
            .with(
                6_000,
                "TST.TableInfoArchive",
                decode_as::<litchi_iwa_protos::tst::TableInfoArchive>,
            )
            .with(
                6_267,
                "TST.ColumnRowUIDMapArchive",
                decode_as::<litchi_iwa_protos::tst::ColumnRowUidMapArchive>,
            )
            .with(
                6_305,
                "TST.StrokeSidecarArchive",
                decode_as::<litchi_iwa_protos::tst::StrokeSidecarArchive>,
            )
    })
    .host(|source| {
        let mut editor = pages::Draft::from_bytes(source)?;
        drop(editor.add_table(0, "Litchi table", 3, 4)?);
        Ok(editor.to_bytes()?)
    })
    .leaf(|source| {
        let package = Package::from_bytes(source)?;
        let commit = package
            .edit_body_table_creation()
            .insert(
                Position::new(0),
                NewTable::new(3, 4).with_name("Litchi table"),
            )?
            .commit()?;
        let restored = commit
            .package()
            .apply_body_table_creation(&commit.patch().inverse())?;
        Ok(LeafRun::with_inverse(
            leaf_bytes!(commit.package()),
            leaf_bytes!(restored.package()),
        ))
    })
    .read("host tables", host_tables)
    .read("leaf body_tables", leaf_tables)
    .run()?;
    Ok(())
}

/// Where a duplicate case anchors its copy.
#[derive(Clone, Copy)]
enum Anchor {
    Start,
    End,
}

/// How a duplicate case compares the cloned formula owners' internal
/// identifiers.
#[derive(Clone, Copy)]
enum Owners {
    /// Equal: the engine's owner map has no entry above its owner objects.
    Equal,
    /// By form: the owner map holds an entry above every owner object, where
    /// the host's allocation collides and the leaf allocates above the map
    /// (policy 0368a); `pages_body_table_duplicate_owner_allocation_differs`
    /// pins the exact relation.
    ByForm,
}

/// `TSCE.FormulaOwnerDependenciesArchive` without its internal owner
/// identifiers: its own, its cell edges' and its range references'.
fn owner_without_internal_identities(data: &[u8]) -> Result<Decoded, prost::DecodeError> {
    let mut owner = litchi_iwa_protos::tsce::FormulaOwnerDependenciesArchive::decode(data)?;
    owner.internal_formula_owner_id = 0;
    if let Some(dependencies) = owner.cell_dependencies.as_mut() {
        for record in &mut dependencies.cell_record {
            if let Some(edges) = record.expanded_edges.as_mut() {
                edges.internal_owner_id_for_edge.clear();
            }
        }
    }
    if let Some(ranges) = owner.range_dependencies.as_mut() {
        for dependency in &mut ranges.back_dependency {
            if let Some(reference) = dependency.internal_range_reference.as_mut() {
                reference.owner_id = 0;
            }
        }
    }
    Ok(decoded(owner, data))
}

/// `TSCE.CalculationEngineArchive` without its owner map's internal
/// identifiers.
fn engine_without_internal_identities(data: &[u8]) -> Result<Decoded, prost::DecodeError> {
    let mut engine = litchi_iwa_protos::tsce::CalculationEngineArchive::decode(data)?;
    if let Some(map) = engine.dependency_tracker.owner_id_map.as_mut() {
        for entry in &mut map.map_entry {
            entry.internal_owner_id = 0;
        }
    }
    Ok(decoded(engine, data))
}

/// `TSCE.CellRecordTileArchive` without its internal owner identifiers.
fn cell_tile_without_identities(data: &[u8]) -> Result<Decoded, prost::DecodeError> {
    let mut tile = litchi_iwa_protos::tsce::CellRecordTileArchive::decode(data)?;
    tile.internal_owner_id = 0;
    for record in &mut tile.cell_records {
        if let Some(edges) = record.expanded_edges.as_mut() {
            edges.internal_owner_id_for_edge.clear();
        }
    }
    Ok(decoded(tile, data))
}

/// `TSCE.RangePrecedentsTileArchive` without its owner identifier.
fn range_tile_without_identity(data: &[u8]) -> Result<Decoded, prost::DecodeError> {
    let mut tile = litchi_iwa_protos::tsce::RangePrecedentsTileArchive::decode(data)?;
    tile.to_owner_id = 0;
    Ok(decoded(tile, data))
}

/// The UTF-16 length of the body story.
fn body_len(source: &[u8]) -> TestResult<usize> {
    Ok(Package::from_bytes(source)?
        .body_storage()?
        .text()
        .encode_utf16()
        .count())
}

/// One body-table duplicate case under `name`: table `table` of `path`,
/// selected by position or by its name.
fn duplicate_case(
    name: &'static str,
    path: &'static str,
    table: usize,
    anchor: Anchor,
    by_name: bool,
    owners: Owners,
) -> TestResult {
    let input = fixture(path)?;
    let anchor = match anchor {
        Anchor::Start => 0,
        Anchor::End => body_len(&input.bytes)?,
    };
    Case::new(name, Format::Pages, Tier::B, Evidence::Queued, &input)
        .leaf_structural_metadata()
        .known_fields(|known| {
            let known = known
                .with(
                    6_000,
                    "TST.TableInfoArchive",
                    decode_as::<litchi_iwa_protos::tst::TableInfoArchive>,
                )
                .with(
                    6_267,
                    "TST.ColumnRowUIDMapArchive",
                    decode_as::<litchi_iwa_protos::tst::ColumnRowUidMapArchive>,
                )
                .with(
                    6_305,
                    "TST.StrokeSidecarArchive",
                    decode_as::<litchi_iwa_protos::tst::StrokeSidecarArchive>,
                );
            match owners {
                Owners::Equal => known,
                Owners::ByForm => known
                    .with(
                        4_000,
                        "TSCE.CalculationEngineArchive (owner-map internal identifiers by form)",
                        engine_without_internal_identities,
                    )
                    .with(
                        4_008,
                        "TSCE.FormulaOwnerDependenciesArchive (internal owner identifiers by form)",
                        owner_without_internal_identities,
                    )
                    .with(
                        4_009,
                        "TSCE.CellRecordTileArchive (internal owner identifiers by form)",
                        cell_tile_without_identities,
                    )
                    .with(
                        4_010,
                        "TSCE.RangePrecedentsTileArchive (owner identifier by form)",
                        range_tile_without_identity,
                    ),
            }
        })
        .host(move |source| host_duplicate(source, table, anchor))
        .leaf(move |source| leaf_duplicate(source, table, anchor, by_name))
        .read("host tables", host_tables)
        .read("leaf body_tables", leaf_tables)
        .run()?;
    Ok(())
}

fn host_duplicate(source: &[u8], table: usize, anchor: usize) -> TestResult<Vec<u8>> {
    let mut editor = pages::Draft::from_bytes(source)?;
    let model = editor
        .tables()?
        .get(table)
        .ok_or("no such table")?
        .model_object_id;
    drop(editor.duplicate_table(model, anchor)?);
    Ok(editor.to_bytes()?)
}

fn leaf_duplicate(
    source: &[u8],
    table: usize,
    anchor: usize,
    by_name: bool,
) -> TestResult<LeafRun> {
    let package = Package::from_bytes(source)?;
    let name = package
        .body_tables()?
        .get(table)
        .ok_or("no such table")?
        .name()
        .to_owned();
    let selector = if by_name {
        BodyTableSelector::name(&name)
    } else {
        BodyTableSelector::index(table)
    };
    let commit = package
        .edit_body_table_creation()
        .duplicate(selector, Position::new(anchor))?
        .commit()?;
    let restored = commit
        .package()
        .apply_body_table_creation(&commit.patch().inverse())?;
    Ok(LeafRun::with_inverse(
        leaf_bytes!(commit.package()),
        leaf_bytes!(restored.package()),
    ))
}

/// Duplicating a populated body table (record 0371, on the shared
/// `litchi_iwa_table::structure::duplicate` engine): at the start and the
/// end of the body, selected by position and by name, with a formula owner
/// family cloned on every input.
#[test]
fn pages_body_table_duplicate_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    let failures: Vec<String> = [
        (CELLS, 0, Anchor::Start, false, Owners::ByForm),
        (TABLES, 0, Anchor::Start, true, Owners::Equal),
        (TABLES, 1, Anchor::End, false, Owners::Equal),
    ]
    .into_iter()
    .filter_map(|(path, table, anchor, by_name, owners)| {
        duplicate_case(
            "pages_body_table_duplicate_parity",
            path,
            table,
            anchor,
            by_name,
            owners,
        )
        .err()
        .map(|error| format!("{path} table {table}: {error}"))
    })
    .collect();
    if failures.is_empty() {
        Ok(())
    } else {
        Err(failures.join("\n").into())
    }
}

/// Every tracked Pages fixture the duplicate sweep visits.
const PAGES_CORPUS: &[&str] = &[
    "iwork/pages/audio-playback-native-resaved.pages",
    "iwork/pages/audio-playback-native.pages",
    "iwork/pages/basic.pages",
    "iwork/pages/body-sections-unicode.pages",
    "iwork/pages/body-table-catalog-native.pages",
    "iwork/pages/body-table-cells-native.pages",
    "iwork/pages/body-table-comment-hidden-native-saved.pages",
    "iwork/pages/body-table-cross-reference-native.pages",
    "iwork/pages/body-table-hidden-axes-cleared-native.pages",
    "iwork/pages/body-table-hidden-axes-focused-native.pages",
    "iwork/pages/body-table-hidden-axes-native.pages",
    "iwork/pages/body-table-merges-native.pages",
    "iwork/pages/body-table-name-focused-native-saved.pages",
    "iwork/pages/body-table-name-focused.pages",
    "iwork/pages/body-table-name-visible-native.pages",
    "iwork/pages/body-table-read-native.pages",
    "iwork/pages/body-table-visible.pages",
    "iwork/pages/chart-arrangement-native.pages",
    "iwork/pages/chart-arrangement-retirement-resaved.pages",
    "iwork/pages/chart-data-edited-native.pages",
    "iwork/pages/chart-data-native.pages",
    "iwork/pages/chart-grid-authored-native.pages",
    "iwork/pages/chart-grid-replaced-native.pages",
    "iwork/pages/image-adjustments-direct-resaved.pages",
    "iwork/pages/image-adjustments-native-resaved.pages",
    "iwork/pages/image-adjustments-native.pages",
    "iwork/pages/image-adjustments-retirement-resaved.pages",
    "iwork/pages/source-built-hidden-axes-before.pages",
    "iwork/pages/source-built-hidden-axes-legacy.pages",
    "iwork/pages/source-built-table-stylesheet-before.pages",
    "iwork/pages/source-built-table-stylesheet-native-saved.pages",
];

/// Whether the engine's owner map holds an internal identifier above every
/// owner object, where the host's and the leaf's allocations part (policy
/// 0368a).
fn owner_map_collides(bytes: &[u8]) -> TestResult<bool> {
    let family = owner_family(bytes)?;
    let Ok(engine) = engine_payload(&family) else {
        return Ok(false);
    };
    let mut owners_maximum = None;
    for (message_type, data) in family.values() {
        if *message_type == 4_008 {
            let owner =
                litchi_iwa_protos::tsce::FormulaOwnerDependenciesArchive::decode(&data[..])?;
            owners_maximum = owners_maximum.max(Some(owner.internal_formula_owner_id));
        }
    }
    let mapped_maximum = owner_map(engine)?
        .iter()
        .map(|(identifier, _)| *identifier)
        .max();
    Ok(mapped_maximum > owners_maximum)
}

/// Duplicating every body table of every corpus document at the start of
/// its body: each copy both sides make is a parity case, and the leaf
/// refuses no table the host copies.
#[test]
fn pages_body_table_duplicate_corpus_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    let mut failures = Vec::new();
    let (mut cases, mut refused) = (0_usize, 0_usize);
    for path in PAGES_CORPUS {
        let input = fixture(path)?;
        let Ok(tables) = pages::Draft::from_bytes(&input.bytes).and_then(|editor| editor.tables())
        else {
            continue;
        };
        let owners = if owner_map_collides(&input.bytes)? {
            Owners::ByForm
        } else {
            Owners::Equal
        };
        for table in 0..tables.len() {
            match (
                host_duplicate(&input.bytes, table, 0),
                leaf_duplicate(&input.bytes, table, 0, false),
            ) {
                (Ok(_), Ok(_)) => {
                    cases += 1;
                    if let Err(error) = duplicate_case(
                        "pages_body_table_duplicate_corpus_parity",
                        path,
                        table,
                        Anchor::Start,
                        false,
                        owners,
                    ) {
                        failures.push(format!("{path} table {table}: {error}"));
                    }
                },
                (Ok(_), Err(error)) => failures.push(format!(
                    "{path} table {table}: leaf refuses where the host copies: {error}"
                )),
                (Err(host), Err(error)) => {
                    refused += 1;
                    eprintln!("{path} table {table}: both refuse (host: {host}; leaf: {error})");
                },
                (Err(host), Ok(_)) => eprintln!("{path} table {table}: host refuses ({host})"),
            }
        }
    }
    eprintln!(
        "pages_body_table_duplicate_corpus_parity: {cases} routes compared, {refused} refused"
    );
    if cases == 0 {
        failures.push("no route compared".to_owned());
    }
    if failures.is_empty() {
        Ok(())
    } else {
        Err(failures.join("\n").into())
    }
}

/// `TSCE.CalculationEngineArchive`, `TSCE.FormulaOwnerDependenciesArchive`,
/// `TSCE.CellRecordTileArchive` and `TSCE.RangePrecedentsTileArchive`.
const OWNER_FAMILY: [u32; 4] = [4_000, 4_008, 4_009, 4_010];

/// The calculation engine and owner-family payloads of a package, by object.
fn owner_family(bytes: &[u8]) -> TestResult<BTreeMap<u64, (u32, Vec<u8>)>> {
    let editor = pages::Draft::from_bytes(bytes)?;
    let mut family = BTreeMap::new();
    for name in editor.package().iwa_entry_names() {
        for object in editor.package().archive(name)?.objects {
            let identifier = object.archive_info.identifier.ok_or("object identifier")?;
            for message in object.messages {
                if OWNER_FAMILY.contains(&message.type_) {
                    family.insert(identifier, (message.type_, message.data));
                }
            }
        }
    }
    Ok(family)
}

/// The engine's owner map: (internal identifier, owner UUID words).
fn owner_map(engine: &[u8]) -> TestResult<Vec<(u32, [Option<u32>; 4])>> {
    let engine = litchi_iwa_protos::tsce::CalculationEngineArchive::decode(engine)?;
    Ok(engine
        .dependency_tracker
        .owner_id_map
        .iter()
        .flat_map(|map| &map.map_entry)
        .map(|entry| {
            let uuid = &entry.owner_id;
            (
                entry.internal_owner_id,
                [uuid.uuid_w0, uuid.uuid_w1, uuid.uuid_w2, uuid.uuid_w3],
            )
        })
        .collect())
}

/// The engine payload of an owner family.
fn engine_payload(family: &BTreeMap<u64, (u32, Vec<u8>)>) -> TestResult<&[u8]> {
    family
        .values()
        .find(|(message_type, _)| *message_type == 4_000)
        .map(|(_, data)| data.as_slice())
        .ok_or_else(|| "no calculation engine".into())
}

/// A cloned owner-family payload with every internal owner identifier at or
/// above `from` moved up by `delta`, re-encoded canonically.
fn shifted(message_type: u32, data: &[u8], from: u32, delta: u32) -> TestResult<Vec<u8>> {
    use litchi_iwa_protos::tsce;
    let shift = |identifier: &mut u32| {
        if *identifier >= from {
            *identifier += delta;
        }
    };
    let edges = |records: &mut Vec<tsce::CellRecordExpandedArchive>| {
        for record in records {
            if let Some(edges) = record.expanded_edges.as_mut() {
                edges.internal_owner_id_for_edge.iter_mut().for_each(shift);
            }
        }
    };
    Ok(match message_type {
        4_008 => {
            let mut owner = tsce::FormulaOwnerDependenciesArchive::decode(data)?;
            shift(&mut owner.internal_formula_owner_id);
            if let Some(dependencies) = owner.cell_dependencies.as_mut() {
                edges(&mut dependencies.cell_record);
            }
            if let Some(ranges) = owner.range_dependencies.as_mut() {
                for dependency in &mut ranges.back_dependency {
                    if let Some(reference) = dependency.internal_range_reference.as_mut() {
                        shift(&mut reference.owner_id);
                    }
                }
            }
            owner.encode_to_vec()
        },
        4_009 => {
            let mut tile = tsce::CellRecordTileArchive::decode(data)?;
            shift(&mut tile.internal_owner_id);
            edges(&mut tile.cell_records);
            tile.encode_to_vec()
        },
        4_010 => {
            let mut tile = tsce::RangePrecedentsTileArchive::decode(data)?;
            shift(&mut tile.to_owner_id);
            tile.encode_to_vec()
        },
        _ => {
            return Err(
                format!("message type {message_type} is not a cloned owner payload").into(),
            );
        },
    })
}

/// The canonical encoding of an owner-family payload.
fn canonical(message_type: u32, data: &[u8]) -> TestResult<Vec<u8>> {
    shifted(message_type, data, u32::MAX, 0)
}

/// Documented intentional difference (policy 0368a, record 0371): the
/// engine's owner map of `body-table-cells-native.pages` holds an internal
/// identifier above every owner object. The host allocates the cloned
/// owners one above the largest owner object, so its copy registers an
/// identifier the map already holds; the leaf allocates above the map
/// (`litchi_numbers_wire::formula_graph::next_owner_id`). Every cloned
/// owner payload and owner-map entry of the leaf equals the host's with the
/// cloned identifiers moved up by one constant.
#[test]
fn pages_body_table_duplicate_owner_allocation_differs() -> TestResult {
    let input = fixture(CELLS)?;
    let mut editor = pages::Draft::from_bytes(&input.bytes)?;
    let model = editor.tables()?.first().ok_or("no table")?.model_object_id;
    drop(editor.duplicate_table(model, 0)?);
    let host = owner_family(&editor.to_bytes()?)?;
    let commit = Package::from_bytes(&input.bytes)?
        .edit_body_table_creation()
        .duplicate(BodyTableSelector::index(0), Position::new(0))?
        .commit()?;
    let leaf = owner_family(&leaf_bytes!(commit.package()))?;
    let source = owner_family(&input.bytes)?;

    // The host repeats a mapped identifier; the leaf allocates above the map.
    let source_map = owner_map(engine_payload(&source)?)?;
    let (host_map, leaf_map) = (
        owner_map(engine_payload(&host)?)?,
        owner_map(engine_payload(&leaf)?)?,
    );
    let repeated = |map: &[(u32, [Option<u32>; 4])]| {
        map.iter()
            .any(|(identifier, _)| map.iter().filter(|(other, _)| other == identifier).count() > 1)
    };
    assert!(repeated(&host_map), "the host defect no longer reproduces");
    assert!(!repeated(&leaf_map));
    let owners: Vec<&[u8]> = source
        .values()
        .filter(|(message_type, _)| *message_type == 4_008)
        .map(|(_, data)| data.as_slice())
        .collect();
    let leaf_first = litchi_numbers_wire::formula_graph::next_owner_id(
        engine_payload(&source)?,
        &owners,
        &litchi_numbers_wire::formula_graph::Limits::DEFAULT,
    )?;
    let mapped_maximum = source_map.iter().map(|(identifier, _)| *identifier).max();
    assert!(mapped_maximum.is_some_and(|maximum| maximum < leaf_first));
    let added = |map: &[(u32, [Option<u32>; 4])]| -> Vec<(u32, [Option<u32>; 4])> {
        map.iter()
            .filter(|entry| !source_map.contains(entry))
            .copied()
            .collect()
    };
    let (host_added, leaf_added) = (added(&host_map), added(&leaf_map));
    let host_first = host_added
        .iter()
        .map(|(identifier, _)| *identifier)
        .min()
        .ok_or("the host clones no owner")?;
    let delta = leaf_first
        .checked_sub(host_first)
        .ok_or("the leaf allocates lower")?;
    assert!(delta > 0);
    assert_eq!(
        leaf_added,
        host_added
            .iter()
            .map(|(identifier, uuid)| (identifier + delta, *uuid))
            .collect::<Vec<_>>()
    );

    // Every cloned owner payload is the host's, shifted.
    let mut compared = 0_usize;
    for (identifier, (message_type, data)) in &host {
        if source.contains_key(identifier) {
            continue;
        }
        let (leaf_type, leaf_data) = leaf.get(identifier).ok_or("a host clone the leaf lacks")?;
        assert_eq!(leaf_type, message_type, "object {identifier}");
        assert_eq!(
            canonical(*leaf_type, leaf_data)?,
            shifted(*message_type, data, host_first, delta)?,
            "object {identifier}"
        );
        compared += 1;
    }
    assert_eq!(compared, leaf.len() - source.len());
    assert!(compared > 0);
    eprintln!(
        "{CELLS}: {} owner-map entries and {compared} cloned owner objects equal the host's \
         moved up by {delta} (host from {host_first}, leaf from {leaf_first})",
        leaf_added.len()
    );
    Ok(())
}

#[test]
fn pages_body_table_remove_parity() -> TestResult {
    let input = fixture(TABLES)?;
    Case::new(
        "pages_body_table_remove_parity",
        Format::Pages,
        Tier::A,
        Evidence::Queued,
        &input,
    )
    .host(|source| {
        let mut editor = pages::Draft::from_bytes(source)?;
        let model = editor
            .tables()?
            .get(1)
            .ok_or("no second table")?
            .model_object_id;
        drop(editor.remove_table(model)?);
        Ok(editor.to_bytes()?)
    })
    .leaf(|source| {
        let package = Package::from_bytes(source)?;
        let commit = package
            .edit_body_table_deletion()
            .remove(BodyTableSelector::index(1))?
            .commit()?;
        let restored = commit
            .package()
            .apply_body_table_deletion(&commit.patch().inverse())?;
        Ok(LeafRun::with_inverse(
            leaf_bytes!(commit.package()),
            leaf_bytes!(restored.package()),
        ))
    })
    .read("host tables", host_tables)
    .read("leaf body_tables", leaf_tables)
    .run()?;
    Ok(())
}

#[test]
#[ignore = "callgrind cost probe (record 0371)"]
fn section_append_cost_abba() -> TestResult {
    crate::pages_cost_probes::abba(
        &fixture(SECTIONS)?,
        |source| {
            let mut editor = pages::Draft::from_bytes(source)?;
            let id = editor.sections().last().ok_or("no section")?.object_id;
            drop(editor.append_section(id, "Appendix")?);
            Ok(editor.to_bytes()?)
        },
        |source| {
            let package = Package::from_bytes(source)?;
            let last = package.sections().len() - 1;
            let commit = package
                .edit_section_lifecycle()
                .add(SectionSelector::index(last), Some("Appendix"))?
                .commit()?;
            Ok(leaf_bytes!(commit.package()))
        },
    )
}

#[test]
#[ignore = "callgrind cost probe (record 0371)"]
fn section_remove_cost_abba() -> TestResult {
    crate::pages_cost_probes::abba(
        &fixture(SECTIONS)?,
        |source| {
            let mut editor = pages::Draft::from_bytes(source)?;
            let id = editor
                .sections()
                .get(1)
                .ok_or("no second section")?
                .object_id;
            drop(editor.remove_section(id)?);
            Ok(editor.to_bytes()?)
        },
        |source| {
            let package = Package::from_bytes(source)?;
            let commit = package
                .edit_section_lifecycle()
                .remove(SectionSelector::index(1))?
                .commit()?;
            Ok(leaf_bytes!(commit.package()))
        },
    )
}

#[test]
#[ignore = "callgrind cost probe (record 0371)"]
fn table_add_cost_abba() -> TestResult {
    crate::pages_cost_probes::abba(
        &fixture(TABLES)?,
        |source| {
            let mut editor = pages::Draft::from_bytes(source)?;
            drop(editor.add_table(0, "Litchi table", 3, 4)?);
            Ok(editor.to_bytes()?)
        },
        |source| {
            let package = Package::from_bytes(source)?;
            let commit = package
                .edit_body_table_creation()
                .insert(
                    Position::new(0),
                    NewTable::new(3, 4).with_name("Litchi table"),
                )?
                .commit()?;
            Ok(leaf_bytes!(commit.package()))
        },
    )
}

#[test]
#[ignore = "callgrind cost probe (record 0371)"]
fn table_remove_cost_abba() -> TestResult {
    crate::pages_cost_probes::abba(
        &fixture(TABLES)?,
        |source| {
            let mut editor = pages::Draft::from_bytes(source)?;
            let model = editor
                .tables()?
                .get(1)
                .ok_or("no second table")?
                .model_object_id;
            drop(editor.remove_table(model)?);
            Ok(editor.to_bytes()?)
        },
        |source| {
            let package = Package::from_bytes(source)?;
            let commit = package
                .edit_body_table_deletion()
                .remove(BodyTableSelector::index(1))?
                .commit()?;
            Ok(leaf_bytes!(commit.package()))
        },
    )
}

#[test]
#[ignore = "callgrind cost probe (record 0371)"]
fn table_duplicate_cost_abba() -> TestResult {
    crate::pages_cost_probes::abba(
        &fixture(TABLES)?,
        |source| host_duplicate(source, 0, 0),
        |source| {
            let package = Package::from_bytes(source)?;
            let commit = package
                .edit_body_table_creation()
                .duplicate(BodyTableSelector::index(0), Position::new(0))?
                .commit()?;
            Ok(leaf_bytes!(commit.package()))
        },
    )
}

/// `insert_section` against the leaf section split: section 0 is copied and
/// the body breaks inside the first section's text.
#[test]
fn pages_section_insert_parity() -> TestResult {
    let input = fixture(SECTIONS)?;
    Case::new(
        "pages_section_insert_parity",
        Format::Pages,
        Tier::B,
        Evidence::Queued,
        &input,
    )
    .host(|source| {
        let mut editor = pages::Draft::from_bytes(source)?;
        let id = editor.sections().first().ok_or("no section")?.object_id;
        drop(editor.insert_section(id, 3, "Split")?);
        Ok(editor.to_bytes()?)
    })
    .leaf(|source| {
        let package = Package::from_bytes(source)?;
        let commit = package
            .edit_section_lifecycle()
            .insert(
                SectionSelector::index(0),
                litchi_pages::TextPosition::from_utf16_index(3)?,
                Some("Split"),
            )?
            .commit()?;
        let restored = commit
            .package()
            .apply_section_lifecycle(&commit.patch().inverse())?;
        Ok(LeafRun::with_inverse(
            leaf_bytes!(commit.package()),
            leaf_bytes!(restored.package()),
        ))
    })
    .read("host sections", host_sections)
    .read("leaf sections", leaf_sections)
    .read("host body text", body_text)
    .leaf_structural_metadata()
    .run()?;
    Ok(())
}

#[test]
#[ignore = "callgrind cost probe (record 0371)"]
fn section_insert_cost_abba() -> TestResult {
    crate::pages_cost_probes::abba(
        &fixture(SECTIONS)?,
        |source| {
            let mut editor = pages::Draft::from_bytes(source)?;
            let id = editor.sections().first().ok_or("no section")?.object_id;
            drop(editor.insert_section(id, 3, "Split")?);
            Ok(editor.to_bytes()?)
        },
        |source| {
            let package = Package::from_bytes(source)?;
            let commit = package
                .edit_section_lifecycle()
                .insert(
                    SectionSelector::index(0),
                    litchi_pages::TextPosition::from_utf16_index(3)?,
                    Some("Split"),
                )?
                .commit()?;
            Ok(leaf_bytes!(commit.package()))
        },
    )
}

/// `body-table-catalog-native.pages` table 0 with column 0 numbers in
/// descending and column 1 text, and a persisted sort order (set through the
/// leaf sort-order family) of `scope` on column `column`.
fn sortable_fixture(
    scope: litchi_iwa_common::table::sort::Scope,
    path: &'static str,
    column: usize,
    direction: litchi_iwa_common::table::sort::Direction,
) -> TestResult<crate::harness::Fixture> {
    use litchi_iwa_common::table::sort::{ColumnIndex, Order, Rule};
    use litchi_pages::table::values::Value;
    let input = fixture(path)?;
    let package = Package::from_bytes(&input.bytes)?;
    let mut edit = package.edit_body_table_values(0)?;
    for (row, number, text) in [
        (0, 3.0, "kiwi"),
        (1, 7.5, "pear"),
        (2, -2.0, "apple"),
        (3, 42.0, "fig"),
        (4, 7.5, "date"),
    ] {
        edit.set(row, 0, Value::Number(number))?;
        edit.set(row, 1, Value::Text(text.to_owned()))?;
    }
    let filled = edit.commit()?;
    let order = Order::with_scope(scope, [Rule::new(ColumnIndex::new(column)?, direction)])?;
    let sorted = filled
        .package()
        .edit_body_table_sort_order(BodyTableSelector::index(0))?
        .set(order)
        .commit()?;
    Ok(crate::harness::Fixture {
        path: "native body table 0 with keys and a sort order",
        bytes: leaf_bytes!(sorted.package()),
    })
}

fn sort_case(
    name: &'static str,
    tier: Tier,
    evidence: Evidence,
    input: &crate::harness::Fixture,
    rows: Option<litchi_iwa_common::table::sort::RowRange>,
) -> TestResult {
    Case::new(name, Format::Pages, tier, evidence, input)
        .host(|source| {
            let mut editor = pages::Draft::from_bytes(source)?;
            let model = editor.tables()?[0].model_object_id;
            let _moved = match rows {
                None => editor.apply_table_sort_order(model)?,
                Some(rows) => editor.apply_table_sort_order_to_rows(model, rows)?,
            };
            Ok(editor.to_bytes()?)
        })
        .leaf(|source| {
            let package = Package::from_bytes(source)?;
            let mut edit = package.edit_body_table_sort_execution(0)?;
            match rows {
                None => drop(edit.apply()?),
                Some(rows) => drop(edit.apply_to_rows(rows)?),
            }
            let commit = edit.commit()?;
            let restored = commit
                .package()
                .apply_body_table_sort_execution(&commit.patch().inverse())?;
            Ok(LeafRun::with_inverse(
                leaf_bytes!(commit.package()),
                leaf_bytes!(restored.package()),
            ))
        })
        .read("host tables", host_tables)
        .read("leaf body_table_cells", |bytes: &[u8]| {
            Ok(format!(
                "{:?}",
                Package::from_bytes(bytes)?.body_table_cells(0)?
            ))
        })
        .run()?;
    Ok(())
}

#[test]
fn pages_body_table_sort_apply_parity() -> TestResult {
    use litchi_iwa_common::table::sort::{Direction, Scope};
    for (column, direction) in [(0, Direction::Ascending), (1, Direction::Descending)] {
        sort_case(
            "pages_body_table_sort_apply_parity",
            Tier::A,
            Evidence::Queued,
            &sortable_fixture(Scope::EntireTable, TABLES, column, direction)?,
            None,
        )?;
    }
    // Explicit cell borders: a row layer moves with its row and a column
    // layer's runs split with the rows.
    let base = sortable_fixture(Scope::EntireTable, TABLES, 0, Direction::Ascending)?;
    let mut editor = pages::Draft::from_bytes(&base.bytes)?;
    let model = editor.tables()?[0].model_object_id;
    let stroke = litchi_iwa_common::shape::stroke::Stroke::new(
        litchi_iwa_common::color::Rgba::new(
            0.2,
            0.4,
            0.6,
            1.0,
            litchi_iwa_common::color::RgbColorSpace::Srgb,
        )?,
        litchi_iwa_common::shape::stroke::Width::new(2.0)?,
        litchi_iwa_common::shape::stroke::Pattern::Solid,
    );
    for (row, column, side) in [
        (1, 0, litchi_iwa_common::table::cell::BorderSide::Top),
        (2, 1, litchi_iwa_common::table::cell::BorderSide::Left),
        (3, 1, litchi_iwa_common::table::cell::BorderSide::Left),
        (4, 2, litchi_iwa_common::table::cell::BorderSide::Bottom),
    ] {
        editor.set_table_cell_border(model, row, column, side, stroke)?;
    }
    let bordered = crate::harness::Fixture {
        path: "native body table 0 with keys, a sort order and host-set cell borders",
        bytes: editor.to_bytes()?,
    };
    sort_case(
        "pages_body_table_sort_apply_parity",
        Tier::A,
        Evidence::Queued,
        &bordered,
        None,
    )?;
    // A table with user-hidden rows and columns: their positions stay.
    sort_case(
        "pages_body_table_sort_apply_parity",
        Tier::A,
        Evidence::Queued,
        &sortable_fixture(
            Scope::EntireTable,
            "iwork/pages/body-table-hidden-axes-native.pages",
            0,
            Direction::Descending,
        )?,
        None,
    )?;
    Ok(())
}

#[test]
fn pages_body_table_sort_apply_to_rows_parity() -> TestResult {
    use litchi_iwa_common::table::sort::{Direction, RowRange, Scope};
    sort_case(
        "pages_body_table_sort_apply_to_rows_parity",
        Tier::A,
        Evidence::Queued,
        &sortable_fixture(Scope::SelectedRows, TABLES, 0, Direction::Descending)?,
        Some(RowRange::new(1, 4)?),
    )
}

#[test]
#[ignore = "callgrind cost probe (record 0371)"]
fn table_sort_apply_cost_abba() -> TestResult {
    use litchi_iwa_common::table::sort::{Direction, Scope};
    crate::pages_cost_probes::abba(
        &sortable_fixture(Scope::EntireTable, TABLES, 0, Direction::Ascending)?,
        |source| {
            let mut editor = pages::Draft::from_bytes(source)?;
            let model = editor.tables()?[0].model_object_id;
            let _moved = editor.apply_table_sort_order(model)?;
            Ok(editor.to_bytes()?)
        },
        |source| {
            let package = Package::from_bytes(source)?;
            let mut edit = package.edit_body_table_sort_execution(0)?;
            edit.apply()?;
            Ok(leaf_bytes!(edit.commit()?.package()))
        },
    )
}

fn resize_case(
    name: &'static str,
    tier: Tier,
    evidence: Evidence,
    input: &crate::harness::Fixture,
    table: usize,
    size: (usize, usize),
) -> TestResult {
    Case::new(name, Format::Pages, tier, evidence, input)
        .host(|source| {
            let mut editor = pages::Draft::from_bytes(source)?;
            let model = editor.tables()?[table].model_object_id;
            editor.resize_table(model, size.0, size.1)?;
            Ok(editor.to_bytes()?)
        })
        .leaf(|source| {
            let package = Package::from_bytes(source)?;
            let mut edit = package.edit_body_table_size(table)?;
            edit.set(size.0, size.1)?;
            let commit = edit.commit()?;
            let restored = commit
                .package()
                .apply_body_table_size(&commit.patch().inverse())?;
            Ok(LeafRun::with_inverse(
                leaf_bytes!(commit.package()),
                leaf_bytes!(restored.package()),
            ))
        })
        .read("host tables", host_tables)
        .read("leaf body_tables", leaf_tables)
        .run()?;
    Ok(())
}

#[test]
fn pages_body_table_resize_parity() -> TestResult {
    let input = fixture(TABLES)?;
    for (table, size) in [(0, (7, 6)), (1, (8, 5))] {
        resize_case(
            "pages_body_table_resize_parity",
            Tier::A,
            Evidence::Queued,
            &input,
            table,
            size,
        )?;
    }
    // Grow, then shrink back to the original size.
    let mut editor = pages::Draft::from_bytes(&input.bytes)?;
    let model = editor.tables()?[0].model_object_id;
    let (rows, columns) = (editor.tables()?[0].rows, editor.tables()?[0].columns);
    editor.resize_table(model, rows + 3, columns + 2)?;
    let grown = crate::harness::Fixture {
        path: "iwork/pages/body-table-catalog-native.pages table 0 grown by the host",
        bytes: editor.to_bytes()?,
    };
    resize_case(
        "pages_body_table_resize_parity",
        Tier::A,
        Evidence::Queued,
        &grown,
        0,
        (rows, columns),
    )
}

#[test]
#[ignore = "callgrind cost probe (record 0371)"]
fn table_resize_cost_abba() -> TestResult {
    crate::pages_cost_probes::abba(
        &fixture(TABLES)?,
        |source| {
            let mut editor = pages::Draft::from_bytes(source)?;
            let model = editor.tables()?[0].model_object_id;
            editor.resize_table(model, 7, 6)?;
            Ok(editor.to_bytes()?)
        },
        |source| {
            let package = Package::from_bytes(source)?;
            let mut edit = package.edit_body_table_size(0)?;
            edit.set(7, 6)?;
            Ok(leaf_bytes!(edit.commit()?.package()))
        },
    )
}
