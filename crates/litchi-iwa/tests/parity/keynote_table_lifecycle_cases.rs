//! Keynote slide-table lifecycle and reads (wave W1-C, record 0372): the
//! host table reads (`slide_tables`, `slide_table` and its `SlideTable`
//! cell, comment and merge accessors) against `Package::{slide_drawable_list,
//! slide_table_cells, slide_table_merges}`; the host `duplicate_slide_table`
//! against `slide::table::lifecycle::Edit::duplicate` on the shared table
//! structure set (`litchi_iwa_table::structure`); and the pinned divergences
//! of `remove_slide_table`, `add_slide_table` and the copied merge owner.

use std::collections::{BTreeMap, BTreeSet};

use litchi_iwa::keynote;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};
use litchi_keynote::slide::drawable::{Geometry, Kind, NewTable};
use litchi_keynote::slide::table::lock::State;
use litchi_keynote::{Package, SlideSelector, TableSelector};

use crate::archive_diff::{self, MetadataPolicy};
use crate::harness::{Case, Evidence, Format, LeafRun, TestResult, Tier, fixture};
use crate::known_fields::KnownFields;
use crate::leaf_bytes;

const DECKS: [&str; 3] = [
    "iwork/keynote/table-discovery.key",
    "iwork/keynote/slide-table-read-native.key",
    "iwork/keynote/slide-table-merges-native.key",
];

/// The deck whose first table has a merged range.
const MERGED: &str = "iwork/keynote/slide-table-read-native.key";

// These are the host-only object islands recorded for the three tracked
// removal fixtures. They are pinned evidence, not an allowance inferred from
// whatever a new leaf output happens to delete.
const TABLE_DISCOVERY_ORPHANS: &[u64] = &[
    2_653_182, 2_653_183, 2_653_185, 2_653_186, 2_653_190, 2_653_193, 2_653_194, 2_653_195,
    2_653_196, 2_653_197, 2_653_198, 2_653_199, 2_653_201, 2_653_202, 2_653_203, 2_653_204,
    2_653_205, 2_653_289, 2_653_290, 2_653_291, 2_653_292, 2_653_293, 2_653_294, 2_653_295,
    2_653_296, 2_653_297, 2_653_298, 2_653_299, 2_653_300, 2_653_469,
];
const READ_NATIVE_ORPHANS: &[u64] = &[
    2_653_182, 2_653_183, 2_653_185, 2_653_186, 2_653_190, 2_653_193, 2_653_194, 2_653_195,
    2_653_196, 2_653_197, 2_653_198, 2_653_199, 2_653_201, 2_653_202, 2_653_203, 2_653_204,
    2_653_205, 2_653_289, 2_653_290, 2_653_291, 2_653_292, 2_653_293, 2_653_294, 2_653_295,
    2_653_296, 2_653_297, 2_653_298, 2_653_299, 2_653_300, 2_653_469, 2_657_477, 2_657_484,
];

const TABLE_DISCOVERY_ORPHAN_MEMBERS: &[&str] = &[
    "Index/Tables/DataList-2653290-2.iwa",
    "Index/Tables/DataList-2653291-2.iwa",
    "Index/Tables/DataList-2653292-2.iwa",
    "Index/Tables/DataList-2653293-2.iwa",
    "Index/Tables/DataList-2653294-2.iwa",
    "Index/Tables/DataList-2653295-2.iwa",
    "Index/Tables/DataList-2653296-2.iwa",
    "Index/Tables/DataList-2653297-2.iwa",
    "Index/Tables/DataList-2653298-2.iwa",
    "Index/Tables/DataList-2653299-2.iwa",
    "Index/Tables/DataList-2653300-2.iwa",
    "Index/Tables/DataList-2653469-2.iwa",
    "Index/Tables/HeaderStorageBucket-2653198-2.iwa",
    "Index/Tables/HeaderStorageBucket-2653199-2.iwa",
    "Index/Tables/Tile-2653289-2.iwa",
];
const READ_NATIVE_ORPHAN_MEMBERS: &[&str] = &[
    "Index/Tables/DataList-2653164-2.iwa",
    "Index/Tables/DataList-2653290-2.iwa",
    "Index/Tables/DataList-2653291-2.iwa",
    "Index/Tables/DataList-2653292-2.iwa",
    "Index/Tables/DataList-2653293-2.iwa",
    "Index/Tables/DataList-2653294-2.iwa",
    "Index/Tables/DataList-2653295-2.iwa",
    "Index/Tables/DataList-2653296-2.iwa",
    "Index/Tables/DataList-2653297-2.iwa",
    "Index/Tables/DataList-2653298-2.iwa",
    "Index/Tables/DataList-2653299-2.iwa",
    "Index/Tables/DataList-2653300-2.iwa",
    "Index/Tables/DataList-2653469-2.iwa",
    "Index/Tables/HeaderStorageBucket-2653198-2.iwa",
    "Index/Tables/HeaderStorageBucket-2653199-2.iwa",
    "Index/Tables/Tile-2653289-2.iwa",
];

fn expected_remove_orphans(path: &str) -> BTreeSet<u64> {
    let ids = if path == "iwork/keynote/slide-table-read-native.key" {
        READ_NATIVE_ORPHANS
    } else {
        TABLE_DISCOVERY_ORPHANS
    };
    ids.iter().copied().collect()
}

fn expected_remove_orphan_members(path: &str) -> &'static [&'static str] {
    if path == "iwork/keynote/slide-table-read-native.key" {
        READ_NATIVE_ORPHAN_MEMBERS
    } else {
        TABLE_DISCOVERY_ORPHAN_MEMBERS
    }
}

/// A read's outcome, with the error reduced to its presence.
fn outcome<T: std::fmt::Debug, E>(result: Result<T, E>) -> String {
    result.map_or_else(|_| "unreadable".to_owned(), |value| format!("{value:?}"))
}

/// Host and leaf agree on every table of every slide: its name, size,
/// frame, appearance and lock state; every stored cell and comment; and its
/// merges, or that both readers refuse them. Returns one row per table.
fn tables_agree(bytes: &[u8]) -> TestResult<Vec<String>> {
    let editor = keynote::Draft::from_bytes(bytes)?;
    let package = Package::from_bytes(bytes)?;
    let mut rows = Vec::new();
    for slide in 0..editor.slides()?.len() {
        let host = editor.slide_tables(slide)?;
        let selector = SlideSelector::index(slide);
        let mut frames = package
            .slide_drawable_list(selector)?
            .into_iter()
            .filter(|drawable| drawable.kind() == Kind::Table)
            .collect::<Vec<_>>();
        frames.sort_by_key(|drawable| drawable.z_position().map(|z| z.get()));
        assert_eq!(host.len(), frames.len(), "slide {slide} table count");
        for (position, (info, frame)) in host.iter().zip(&frames).enumerate() {
            let table = TableSelector::index(position);
            let geometry = frame.geometry().ok_or("leaf table has no frame")?;
            assert_eq!(
                (
                    info.geometry.position.map(|p| (p.x, p.y)),
                    info.geometry.size.map(|s| (s.width, s.height)),
                    info.geometry.angle,
                ),
                (
                    Some((geometry.x, geometry.y)),
                    Some((geometry.width, geometry.height)),
                    geometry.rotation,
                ),
                "slide {slide} table {position} frame"
            );
            assert_eq!(
                info.appearance,
                package.slide_table_appearance(selector, table)?
            );
            assert_eq!(info.lock_state, State::from_locked(frame.is_locked()));
            let read = package.slide_table_cells(selector, table)?;
            assert_eq!(read.table().name(), info.name);
            assert_eq!(
                (
                    usize::try_from(read.table().row_count())?,
                    usize::try_from(read.table().column_count())?
                ),
                (info.rows, info.columns)
            );
            let host_merges = outcome(editor.slide_table_cell_merges(slide, info.model_object_id));
            let leaf_merges = outcome(package.slide_table_merges(selector, table));
            assert_eq!(
                host_merges, leaf_merges,
                "slide {slide} table {position} merges"
            );
            let view = match editor.slide_table(slide, info.model_object_id) {
                Ok(view) => view,
                Err(_) if leaf_merges == "unreadable" => {
                    rows.push(format!(
                        "{} {}x{}: merges unreadable",
                        info.name, info.rows, info.columns
                    ));
                    continue;
                },
                Err(error) => return Err(error.into()),
            };
            assert_eq!(format!("{:?}", view.merges()), leaf_merges);
            assert_eq!(view.cell_count(), read.cell_count());
            assert_eq!(view.comment_count(), read.comment_count());
            let host_cells = view
                .iter_cells()
                .map(|(at, value)| (at, format!("{value:?}")))
                .collect::<Vec<_>>();
            let leaf_cells = read
                .iter_cells()
                .map(|cell| {
                    (
                        (
                            cell.position().row() as usize,
                            cell.position().column() as usize,
                        ),
                        format!("{:?}", cell.value()),
                    )
                })
                .collect::<Vec<_>>();
            assert_eq!(host_cells, leaf_cells);
            for ((row, column), _) in &host_cells {
                let at = litchi_keynote::slide::table::At::new(
                    u32::try_from(*row)?,
                    u32::try_from(*column)?,
                );
                assert_eq!(
                    view.get_cell(*row, *column)
                        .map(|value| format!("{value:?}")),
                    read.get(at).map(|value| format!("{value:?}"))
                );
                assert_eq!(view.get_comment(*row, *column), read.get_comment(at));
            }
            let host_comments = view
                .iter_comments()
                .map(|(at, comment)| (at, format!("{comment:?}")))
                .collect::<Vec<_>>();
            let leaf_comments = read
                .iter_comments()
                .map(|comment| {
                    (
                        (
                            comment.position().row() as usize,
                            comment.position().column() as usize,
                        ),
                        format!("{:?}", comment.comment()),
                    )
                })
                .collect::<Vec<_>>();
            assert_eq!(host_comments, leaf_comments);
            rows.push(format!(
                "{} {}x{} {:?}: {} cells, {} comments, merges {leaf_merges}",
                info.name,
                info.rows,
                info.columns,
                (geometry.x, geometry.y, geometry.width, geometry.height),
                read.cell_count(),
                read.comment_count(),
            ));
        }
    }
    Ok(rows)
}

#[test]
fn keynote_slide_table_read_parity() -> TestResult {
    let evidence = Evidence::Native("test-data/iwork/keynote/slide-table-read-native-receipt.json");
    for path in DECKS {
        let input = fixture(path)?;
        Case::new(
            "keynote_slide_table_read_parity",
            Format::Keynote,
            Tier::A,
            evidence,
            &input,
        )
        .read_only()
        .host(|source| Ok(keynote::Draft::from_bytes(source)?.to_bytes()?))
        .leaf(|source| Ok(LeafRun::output(leaf_bytes!(Package::from_bytes(source)?))))
        .read("host and leaf tables agree", tables_agree)
        .run()?;
    }
    Ok(())
}

/// The host's first table on the first slide.
fn first_table(editor: &keynote::Draft) -> TestResult<u64> {
    Ok(editor
        .slide_tables(0)?
        .first()
        .ok_or("no slide table")?
        .drawable_object_id)
}

/// The leaf route: one lifecycle operation, then its inverse.
fn leaf_lifecycle(
    source: &[u8],
    stage: impl FnOnce(
        litchi_keynote::slide::table::lifecycle::Edit<'_>,
    ) -> Result<
        litchi_keynote::slide::table::lifecycle::Edit<'_>,
        litchi_keynote::slide::table::lifecycle::Error,
    >,
) -> TestResult<LeafRun> {
    let package = Package::from_bytes(source)?;
    let commit = stage(package.edit_slide_table_lifecycle())?.commit()?;
    assert!(commit.diagnostics().changed());
    let restored = commit
        .package()
        .apply_slide_table_lifecycle(&commit.patch().inverse())?;
    Ok(LeafRun::with_inverse(
        leaf_bytes!(commit.package()),
        leaf_bytes!(restored.package()),
    ))
}

#[test]
fn keynote_slide_table_duplicate_parity() -> TestResult {
    for path in DECKS {
        let input = fixture(path)?;
        if !Package::from_bytes(&input.bytes)?
            .slide_table_merges(SlideSelector::index(0), TableSelector::index(0))?
            .is_empty()
        {
            continue;
        }
        Case::new(
            "keynote_slide_table_duplicate_parity",
            Format::Keynote,
            Tier::B,
            Evidence::Queued,
            &input,
        )
        .host(|source| {
            let mut editor = keynote::Draft::from_bytes(source)?;
            let table = first_table(&editor)?;
            editor.duplicate_slide_table(0, table)?;
            Ok(editor.to_bytes()?)
        })
        .leaf(|source| {
            leaf_lifecycle(source, |edit| {
                edit.duplicate(SlideSelector::index(0), TableSelector::index(0))
            })
        })
        .leaf_invalidates_root_previews()
        .leaf_structural_metadata()
        .read("host and leaf tables agree", tables_agree)
        .run()?;
    }
    Ok(())
}

/// Every object of `bytes` by identifier: its component, message types and
/// the objects its header references.
fn objects(bytes: &[u8]) -> TestResult<BTreeMap<u64, (String, Vec<u32>, Vec<u64>)>> {
    let package = litchi_iwa::Package::from_bytes(bytes)?;
    let names = package
        .iwa_entry_names()
        .map(str::to_owned)
        .collect::<Vec<_>>();
    let mut objects = BTreeMap::new();
    for name in names {
        for object in package.archive(&name)?.objects {
            let identifier = object
                .archive_info
                .identifier
                .ok_or("object without identifier")?;
            let types = object
                .messages
                .iter()
                .map(|message| message.type_)
                .collect();
            let references = object
                .archive_info
                .message_infos
                .iter()
                .flat_map(|info| info.object_references.iter().copied())
                .collect();
            objects.insert(identifier, (name.clone(), types, references));
        }
    }
    Ok(objects)
}

/// Pinned host defect (record 0372): the host's `remove_slide_table`
/// removes the drawable, the model, the model's data store and their
/// formula owners, but leaves every other object only the table addressed:
/// its summary model with that model's tile, lists and header buckets, its
/// title and caption stand-ins and its reference tracker, with their
/// CalculationEngine owners. The leaf removes the table's private closure
/// (the rule the Numbers host adopted in record 0370g). Outside that
/// orphaned island both outputs hold the same objects, in the same
/// components, and read the same tables.
#[test]
fn keynote_slide_table_remove_divergence_is_reported() -> TestResult {
    // Tier::B, Evidence::Queued: this host mutation has no native open/save
    // receipt; the object island is an intentional host/leaf difference.
    for path in DECKS {
        let input = fixture(path)?.bytes;
        let mut editor = keynote::Draft::from_bytes(&input)?;
        let table = first_table(&editor)?;
        editor.remove_slide_table(0, table)?;
        let host = editor.to_bytes()?;
        let package = Package::from_bytes(&input)?;
        let commit = package
            .edit_slide_table_lifecycle()
            .remove(SlideSelector::index(0), TableSelector::index(0))?
            .commit()?;
        let leaf = leaf_bytes!(commit.package());
        let restored = commit
            .package()
            .apply_slide_table_lifecycle(&commit.patch().inverse())?;
        assert_eq!(leaf_bytes!(restored.package()), input, "{path}: inverse");

        let host_objects = objects(&host)?;
        let leaf_objects = objects(&leaf)?;
        let orphans = host_objects
            .keys()
            .filter(|identifier| !leaf_objects.contains_key(identifier))
            .copied()
            .collect::<BTreeSet<_>>();
        assert!(!orphans.is_empty(), "{path}: the host keeps no orphan");
        assert_eq!(
            orphans,
            expected_remove_orphans(path),
            "{path}: host-only island changed; update the fixture evidence and closure proof"
        );
        for (identifier, (component, types, references)) in &host_objects {
            if orphans.contains(identifier) {
                continue;
            }
            assert!(
                references.iter().all(|target| !orphans.contains(target)),
                "{path}: kept object {identifier} references an orphan"
            );
            let (leaf_component, leaf_types, leaf_references) = leaf_objects
                .get(identifier)
                .ok_or("a host object is missing from the leaf")?;
            assert_eq!(
                (component, types, references),
                (leaf_component, leaf_types, leaf_references),
                "{path}: surviving object {identifier} metadata/references"
            );
        }
        assert!(
            leaf_objects
                .keys()
                .all(|identifier| host_objects.contains_key(identifier)),
            "{path}: the leaf keeps an object the host removed"
        );
        let mut island = BTreeMap::<u32, usize>::new();
        for identifier in &orphans {
            for message_type in &host_objects[identifier].1 {
                *island.entry(*message_type).or_default() += 1;
            }
        }
        // The summary model (TST.SummaryModelArchive) is among them.
        assert!(island.contains_key(&6_316), "{path}: {island:?}");
        let diff = archive_diff::compare_with_policies_except_objects_and_members(
            &input,
            &host,
            &leaf,
            &KnownFields::for_format(Format::Keynote),
            &["preview.jpg", "preview-micro.jpg", "preview-web.jpg"],
            MetadataPolicy {
                save_token: true,
                fresh_uuids: true,
            },
            &expected_remove_orphans(path),
            expected_remove_orphan_members(path),
        )?;
        assert!(
            diff.violations.is_empty(),
            "{path}: surviving known fields, references, or untouched bytes differ\n{}",
            diff.render()
        );
        assert_eq!(tables_agree(&host)?, tables_agree(&leaf)?);
        if path == DECKS[0] {
            // Negative control: removing an unrelated surviving text box must
            // not disappear behind the first table's pinned orphan allowance.
            let package = Package::from_bytes(&leaf)?;
            let drawables = package.slide_drawable_list(SlideSelector::index(0))?;
            let text_box = drawables
                .into_iter()
                .find(|drawable| drawable.kind() != Kind::Table)
                .ok_or("table-discovery fixture has no surviving text box")?;
            let mut edit = package.edit_slide_drawables();
            edit.remove(SlideSelector::index(0), text_box.selector())?;
            let altered_leaf = leaf_bytes!(edit.commit()?.package());
            let negative = archive_diff::compare_with_policies_except_objects_and_members(
                &input,
                &host,
                &altered_leaf,
                &KnownFields::for_format(Format::Keynote),
                &["preview.jpg", "preview-micro.jpg", "preview-web.jpg"],
                MetadataPolicy {
                    save_token: true,
                    fresh_uuids: true,
                },
                &expected_remove_orphans(path),
                expected_remove_orphan_members(path),
            )?;
            assert!(
                !negative.violations.is_empty(),
                "{path}: fixed orphan allowance masked an unrelated survivor removal"
            );
        }
    }
    Ok(())
}

/// The frame both creation routes get.
const POINT: DrawablePoint = DrawablePoint { x: 120.0, y: 140.0 };
const SIZE: DrawableSize = DrawableSize {
    width: 360.0,
    height: 180.0,
};

/// Pinned divergence (record 0372): the host creates a slide table only
/// from the detached creation scaffold a builder deck carries and refuses a
/// deck without one; the leaf (record 0329) copies the deck's first slide
/// table and refuses a deck without one. No deck admits both routes, so
/// `add_slide_table` stays on the host.
#[test]
fn keynote_slide_table_add_divergence_is_reported() -> TestResult {
    let built = keynote::Builder::new()
        .title("Tables")
        .build()?
        .to_bytes()?;
    let native = fixture("iwork/keynote/table-discovery.key")?.bytes;
    for (bytes, scaffold) in [(built, true), (native, false)] {
        let mut editor = keynote::Draft::from_bytes(&bytes)?;
        let host = editor.add_slide_table(0, "Litchi table", 3, 4, POINT, SIZE);
        assert_eq!(host.is_ok(), scaffold);
        let package = Package::from_bytes(&bytes)?;
        let mut edit = package.edit_slide_drawables();
        let leaf = edit.add_table(
            SlideSelector::index(0),
            NewTable::new(
                3,
                4,
                Geometry::new(POINT.x, POINT.y, SIZE.width, SIZE.height),
            )
            .with_name("Litchi table"),
        );
        assert_eq!(leaf.is_ok(), !scaffold);
    }
    Ok(())
}

/// Pinned host defect (record 0372): the host's table-graph copy copies the
/// model's merge owner verbatim, so the copied table's merge still names the
/// source table. The leaf refuses this dependency before clone; a successful
/// leaf commit must never publish a table that neither merge reader accepts.
#[test]
fn keynote_slide_table_duplicate_merge_owner_divergence_is_reported() -> TestResult {
    let input = fixture(MERGED)?.bytes;
    let mut editor = keynote::Draft::from_bytes(&input)?;
    let table = first_table(&editor)?;
    editor.duplicate_slide_table(0, table)?;
    let host = editor.to_bytes()?;
    let editor = keynote::Draft::from_bytes(&host)?;
    let package = Package::from_bytes(&host)?;
    let tables = editor.slide_tables(0)?;
    assert_eq!(tables.len(), 2);
    assert_eq!(
        editor
            .slide_table_cell_merges(0, tables[0].model_object_id)?
            .len(),
        1
    );
    assert_eq!(
        package
            .slide_table_merges(SlideSelector::index(0), TableSelector::index(0))?
            .len(),
        1
    );
    assert!(
        editor
            .slide_table_cell_merges(0, tables[1].model_object_id)
            .is_err()
    );
    assert!(
        package
            .slide_table_merges(SlideSelector::index(0), TableSelector::index(1))
            .is_err()
    );

    let error = Package::from_bytes(&input)?
        .edit_slide_table_lifecycle()
        .duplicate(SlideSelector::index(0), TableSelector::index(0))
        .expect_err("merged-table duplication must refuse before clone");
    assert!(matches!(
        error,
        litchi_keynote::slide::table::lifecycle::Error::Refused(
            litchi_keynote::slide::table::lifecycle::Refusal::UnsupportedDependency
        )
    ));
    Ok(())
}
