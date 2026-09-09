//! Native Pages hidden-axis lifecycle coverage.
//!
//! The checked-in sources were produced by the focused and legacy hidden-axis
//! routes, then opened, saved, closed, and reopened by Pages 14.4. These tests
//! keep those native sources separate from the synthetic ownership fixtures:
//! semantic reads, exact no-ops, changed transactions, and source-bound
//! patches all operate on producer-shaped packages while owner creation stays
//! outside the qualified native profile.

use std::{env, fs, io, path::Path};

use litchi_iwa_archive::{
    Limits,
    iwa::{Archive, SnappyStream},
    package::{Catalog, EntryEdit},
};
use litchi_iwa_core::ArchiveObject;
use litchi_pages::{
    BodyTableHiddenAxesError, BodyTableSelector, Package,
    table::hidden_axes::{AxisIndex, HiddenAxes},
};

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

const NATIVE_HIDDEN_SOURCE: &str = concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../test-data/iwork/pages/body-table-hidden-axes-native.pages"
);
const NATIVE_VISIBLE_SOURCE: &str = concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../test-data/iwork/pages/body-table-visible.pages"
);
const NATIVE_FOCUSED_SOURCE: &[u8] = include_bytes!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../test-data/iwork/pages/body-table-hidden-axes-focused-native.pages"
));
const NATIVE_CLEARED_SOURCE: &[u8] = include_bytes!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../test-data/iwork/pages/body-table-hidden-axes-cleared-native.pages"
));
const PREVIEWS: [&str; 3] = ["preview.jpg", "preview-micro.jpg", "preview-web.jpg"];
const DEPENDENCY_MESSAGE_TYPES: [u32; 3] = [4_008, 6_204, 6_220];

fn source_bytes(path: &str) -> TestResult<Vec<u8>> {
    Ok(fs::read(path)?)
}

fn exact_bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

fn native_hidden_axes() -> TestResult<HiddenAxes> {
    Ok(HiddenAxes::new([AxisIndex::row(2), AxisIndex::column(1)])?)
}

fn changed_axes() -> TestResult<HiddenAxes> {
    Ok(HiddenAxes::new([
        AxisIndex::row(0),
        AxisIndex::row(4),
        AxisIndex::column(0),
        AxisIndex::column(3),
    ])?)
}

fn export_candidate(name: &str, bytes: &[u8]) -> TestResult {
    let Ok(directory) = env::var("LITCHI_PAGES_FOCUSED_HIDDEN_AXES_DIRECTORY") else {
        return Ok(());
    };
    let directory = Path::new(&directory);
    fs::create_dir_all(directory)?;
    fs::write(directory.join(name), bytes)?;
    Ok(())
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct EntrySnapshot {
    name: String,
    raw_name: Vec<u8>,
    data: Vec<u8>,
    metadata: litchi_iwa_archive::package::EntryMetadata,
}

#[derive(Debug)]
struct SelectedComponent {
    name: String,
    edited_object_ids: Box<[u64]>,
}

fn selected_component(source: &[u8]) -> TestResult<SelectedComponent> {
    let catalog = Catalog::from_bytes(source)?;
    let mut candidates = Vec::new();
    for entry in catalog
        .iter()
        .filter(|entry| entry.name().ends_with(".iwa"))
    {
        let decompressed = SnappyStream::decompress(entry.data())?;
        let archive = Archive::parse(decompressed.as_bytes())?;
        let info_ids = archive
            .objects
            .iter()
            .filter(|object| object.messages.iter().any(|message| message.type_ == 6_000))
            .filter_map(|object| object.archive_info.identifier)
            .collect::<Vec<_>>();
        let model_ids = archive
            .objects
            .iter()
            .filter(|object| object.messages.iter().any(|message| message.type_ == 6_001))
            .filter_map(|object| object.archive_info.identifier)
            .collect::<Vec<_>>();
        if !info_ids.is_empty() && !model_ids.is_empty() {
            candidates.push((entry.name().to_owned(), info_ids, model_ids));
        }
    }
    candidates.sort_unstable();
    candidates.dedup();
    let [(name, info_ids, model_ids)] = candidates.as_slice() else {
        return Err(io::Error::other(format!(
            "expected one native table component, found {candidates:?}"
        ))
        .into());
    };
    let [info_id] = info_ids.as_slice() else {
        return Err(io::Error::other(format!(
            "expected one native TableInfo object, found {info_ids:?}"
        ))
        .into());
    };
    let [model_id] = model_ids.as_slice() else {
        return Err(io::Error::other(format!(
            "expected one native TableModel object, found {model_ids:?}"
        ))
        .into());
    };
    Ok(SelectedComponent {
        name: name.clone(),
        edited_object_ids: Box::new([*info_id, *model_id]),
    })
}

fn unselected_entries(source: &[u8], selected_component: &str) -> TestResult<Vec<EntrySnapshot>> {
    let catalog = Catalog::from_bytes(source)?;
    let mut entries = catalog
        .iter()
        .filter(|entry| entry.name() != selected_component && !PREVIEWS.contains(&entry.name()))
        .map(|entry| EntrySnapshot {
            name: entry.name().to_owned(),
            raw_name: entry.raw_name().to_vec(),
            data: entry.data().to_vec(),
            metadata: entry.metadata().clone(),
        })
        .collect::<Vec<_>>();
    entries.sort_by(|left, right| left.name.cmp(&right.name));
    Ok(entries)
}

fn archive_for_member(source: &[u8], member: &str) -> TestResult<Archive> {
    let catalog = Catalog::from_bytes(source)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == member)
        .ok_or_else(|| io::Error::other(format!("missing selected member {member}")))?;
    let decompressed = SnappyStream::decompress(entry.data())?;
    Ok(Archive::parse(decompressed.as_bytes())?)
}

fn assert_selected_member_locality(
    before_source: &[u8],
    after_source: &[u8],
    selected: &SelectedComponent,
) -> TestResult {
    let before = archive_for_member(before_source, &selected.name)?;
    let after = archive_for_member(after_source, &selected.name)?;
    assert_eq!(before.objects.len(), after.objects.len());
    for (before, after) in before.objects.iter().zip(&after.objects) {
        let before_id = before
            .archive_info
            .identifier
            .ok_or_else(|| io::Error::other("selected member object has no identifier"))?;
        let after_id = after
            .archive_info
            .identifier
            .ok_or_else(|| io::Error::other("selected member object has no identifier"))?;
        assert_eq!(before_id, after_id);
        if selected.edited_object_ids.contains(&before_id) {
            continue;
        }
        assert!(
            before.same_content_ignoring_offsets(after),
            "unselected object {before_id} in {} changed payload, ArchiveInfo, or framing",
            selected.name
        );
    }
    Ok(())
}

#[derive(Debug, Clone)]
struct DependencySnapshot {
    member: String,
    identifier: u64,
    message_type: u32,
    object: ArchiveObject,
}

fn dependency_objects(source: &[u8]) -> TestResult<Vec<DependencySnapshot>> {
    let catalog = Catalog::from_bytes(source)?;
    let mut dependencies = Vec::new();
    for entry in catalog
        .iter()
        .filter(|entry| entry.name().ends_with(".iwa"))
    {
        let decompressed = SnappyStream::decompress(entry.data())?;
        let archive = Archive::parse(decompressed.as_bytes())?;
        for object in archive.objects {
            let Some(message_type) = object.primary_message_type() else {
                continue;
            };
            if !DEPENDENCY_MESSAGE_TYPES.contains(&message_type) {
                continue;
            }
            let identifier = object
                .archive_info
                .identifier
                .ok_or_else(|| io::Error::other("dependency object has no identifier"))?;
            dependencies.push(DependencySnapshot {
                member: entry.name().to_owned(),
                identifier,
                message_type,
                object,
            });
        }
    }
    dependencies.sort_by(|left, right| {
        left.member
            .cmp(&right.member)
            .then(left.identifier.cmp(&right.identifier))
            .then(left.message_type.cmp(&right.message_type))
    });
    Ok(dependencies)
}

fn assert_dependencies_preserved(before: &[DependencySnapshot], after: &[DependencySnapshot]) {
    assert_eq!(before.len(), after.len(), "dependency object count changed");
    assert!(
        before.iter().any(|object| object.message_type == 4_008),
        "native fixture has no 4008 dependency object"
    );
    assert!(
        before.iter().any(|object| object.message_type == 6_204),
        "native fixture has no 6204 dependency object"
    );
    assert!(
        before.iter().any(|object| object.message_type == 6_220),
        "native fixture has no 6220 dependency object"
    );
    for (before, after) in before.iter().zip(after) {
        assert_eq!(before.member, after.member);
        assert_eq!(before.identifier, after.identifier);
        assert_eq!(before.message_type, after.message_type);
        assert!(
            before.object.same_content_ignoring_offsets(&after.object),
            "native dependency {}/{} (type {}) changed payload or ArchiveInfo",
            before.member,
            before.identifier,
            before.message_type
        );
    }
}

fn without_previews(source: &[u8]) -> TestResult<Vec<u8>> {
    Ok(
        Catalog::from_bytes(source)?.reassemble_with_deletions_to_bytes(
            &[],
            &PREVIEWS,
            Limits::default(),
        )?,
    )
}

#[test]
fn native_hidden_axes_read_and_exact_noop_preserve_source() -> TestResult {
    let source = source_bytes(NATIVE_HIDDEN_SOURCE)?;
    let package = Package::from_bytes(&source)?;
    let expected = native_hidden_axes()?;
    assert_eq!(package.body_table_hidden_axes(0usize)?, expected);
    assert_eq!(
        package.body_table_hidden_axes(BodyTableSelector::name("Table 1"))?,
        expected
    );
    assert_eq!(exact_bytes(&package)?, source);

    let noop = package
        .edit_body_table_hidden_axes(0usize)?
        .set(expected)
        .commit()?;
    assert!(noop.patch().is_noop());
    assert!(!noop.diagnostics().changed());
    assert_eq!(noop.diagnostics().touched_components(), 0);
    assert_eq!(exact_bytes(noop.package())?, source);
    let applied = package.apply_body_table_hidden_axes(noop.patch())?;
    assert_eq!(exact_bytes(applied.package())?, source);
    let restored = applied
        .package()
        .apply_body_table_hidden_axes(&noop.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, source);
    Ok(())
}

#[test]
fn focused_native_hidden_axes_read_and_exact_noop_preserve_source() -> TestResult {
    let source = NATIVE_FOCUSED_SOURCE;
    let package = Package::from_bytes(source)?;
    let expected = native_hidden_axes()?;
    assert_eq!(package.body_table_hidden_axes(0usize)?, expected);
    assert_eq!(
        package.body_table_hidden_axes(BodyTableSelector::name("Table 1"))?,
        expected
    );
    assert_eq!(exact_bytes(&package)?, source);

    let noop = package
        .edit_body_table_hidden_axes(0usize)?
        .set(expected)
        .commit()?;
    assert!(noop.patch().is_noop());
    assert!(!noop.diagnostics().changed());
    assert_eq!(noop.diagnostics().touched_components(), 0);
    assert_eq!(exact_bytes(noop.package())?, source);
    Ok(())
}

#[test]
fn focused_native_cleared_hidden_axes_read_noop_and_rehide() -> TestResult {
    let source = NATIVE_CLEARED_SOURCE;
    let package = Package::from_bytes(source)?;
    assert_eq!(package.body_table_hidden_axes(0usize)?, HiddenAxes::empty());
    assert_eq!(
        package.body_table_hidden_axes(BodyTableSelector::name("Table 1"))?,
        HiddenAxes::empty()
    );
    assert_eq!(exact_bytes(&package)?, source);

    let noop = package
        .edit_body_table_hidden_axes(0usize)?
        .set(HiddenAxes::empty())
        .commit()?;
    assert!(noop.patch().is_noop());
    assert!(!noop.diagnostics().changed());
    assert_eq!(noop.diagnostics().touched_components(), 0);
    assert_eq!(exact_bytes(noop.package())?, source);

    let rehidden = package
        .edit_body_table_hidden_axes(0usize)?
        .set(native_hidden_axes()?)
        .commit()?;
    assert_eq!(
        rehidden.package().body_table_hidden_axes(0usize)?,
        native_hidden_axes()?
    );
    assert_eq!(
        Package::from_bytes(&exact_bytes(rehidden.package())?)?.body_table_hidden_axes(0usize)?,
        native_hidden_axes()?
    );
    Ok(())
}

#[test]
fn native_hidden_axes_set_clear_reset_reopen_and_inverse() -> TestResult {
    let source = source_bytes(NATIVE_HIDDEN_SOURCE)?;
    let package = Package::from_bytes(&source)?;
    let requested = changed_axes()?;

    let set = package
        .edit_body_table_hidden_axes(BodyTableSelector::name("Table 1"))?
        .set(requested.clone())
        .commit()?;
    assert!(!set.patch().is_noop());
    assert_eq!(set.diagnostics().touched_components(), 1);
    assert_eq!(set.package().body_table_hidden_axes(0usize)?, requested);
    let set_bytes = exact_bytes(set.package())?;
    assert_eq!(
        Package::from_bytes(&set_bytes)?.body_table_hidden_axes(0usize)?,
        requested
    );
    let restored = set
        .package()
        .apply_body_table_hidden_axes(&set.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, source);
    assert_eq!(
        restored.package().body_table_hidden_axes(0usize)?,
        native_hidden_axes()?
    );

    let clear = package
        .edit_body_table_hidden_axes(0usize)?
        .clear()
        .commit()?;
    assert_eq!(
        clear.package().body_table_hidden_axes(0usize)?,
        HiddenAxes::empty()
    );
    assert_eq!(
        Package::from_bytes(&exact_bytes(clear.package())?)?.body_table_hidden_axes(0usize)?,
        HiddenAxes::empty()
    );
    assert_eq!(
        exact_bytes(
            clear
                .package()
                .apply_body_table_hidden_axes(&clear.patch().inverse())?
                .package(),
        )?,
        source
    );

    let reset = package
        .edit_body_table_hidden_axes(0usize)?
        .reset()
        .commit()?;
    assert_eq!(
        reset.package().body_table_hidden_axes(0usize)?,
        HiddenAxes::empty()
    );
    assert_eq!(exact_bytes(reset.package())?, exact_bytes(clear.package())?);
    Ok(())
}

#[test]
fn native_hidden_axes_preserve_dependencies_locality_and_conflict_fences() -> TestResult {
    let source = source_bytes(NATIVE_HIDDEN_SOURCE)?;
    let package = Package::from_bytes(&source)?;
    let selected_component = selected_component(&source)?;
    let before_entries = unselected_entries(&source, &selected_component.name)?;
    let before_dependencies = dependency_objects(&source)?;
    let commit = package
        .edit_body_table_hidden_axes(0usize)?
        .set(changed_axes()?)
        .commit()?;
    let target = exact_bytes(commit.package())?;
    let after_entries = unselected_entries(&target, &selected_component.name)?;
    let after_dependencies = dependency_objects(&target)?;
    assert_eq!(after_entries, before_entries);
    assert_selected_member_locality(&source, &target, &selected_component)?;
    assert_dependencies_preserved(&before_dependencies, &after_dependencies);
    assert_eq!(commit.diagnostics().touched_components(), 1);
    assert_eq!(commit.diagnostics().deleted_previews(), PREVIEWS.len());
    let catalog = Catalog::from_bytes(&target)?;
    for preview in PREVIEWS {
        assert!(catalog.iter().all(|entry| entry.name() != preview));
    }

    let tampered_source = Catalog::from_bytes(&source)?.reassemble_to_bytes(
        &[EntryEdit::new("preview.jpg", b"tampered-preview")],
        Limits::default(),
    )?;
    let tampered = Package::from_bytes(&tampered_source)?;
    assert!(matches!(
        tampered.apply_body_table_hidden_axes(commit.patch()),
        Err(BodyTableHiddenAxesError::PatchConflict)
    ));
    let foreign = Package::from_bytes(&source_bytes(NATIVE_VISIBLE_SOURCE)?)?;
    assert!(matches!(
        foreign.apply_body_table_hidden_axes(commit.patch()),
        Err(BodyTableHiddenAxesError::PatchConflict)
    ));
    Ok(())
}

#[test]
fn native_hidden_axes_resource_limits_are_atomic() -> TestResult {
    let source = source_bytes(NATIVE_HIDDEN_SOURCE)?;
    let source_without_previews = without_previews(&source)?;
    let requested = changed_axes()?;
    let unrestricted = Package::from_bytes(&source_without_previews)?
        .edit_body_table_hidden_axes(0usize)?
        .set(requested.clone())
        .commit()?;
    let target_len = exact_bytes(unrestricted.package())?.len();
    assert!(
        target_len > source_without_previews.len(),
        "native rewrite should grow the no-preview source"
    );

    let bounded = Limits::new(
        u64::try_from(target_len - 1)?,
        Limits::MAX_ENTRIES,
        Limits::MAX_ENTRY_BYTES,
        Limits::MAX_TOTAL_BYTES,
        Limits::MAX_IWA_STREAM_BYTES,
    )?;
    let package = Package::from_bytes_with_limits(&source_without_previews, bounded)?;
    let before = exact_bytes(&package)?;
    let result = package
        .edit_body_table_hidden_axes(0usize)?
        .set(requested)
        .commit();
    assert!(
        matches!(result, Err(BodyTableHiddenAxesError::LimitExceeded { .. })),
        "native output overrun was accepted: {result:?}"
    );
    assert_eq!(exact_bytes(&package)?, before);

    let input_limited = Limits::new(
        u64::try_from(source.len() - 1)?,
        Limits::MAX_ENTRIES,
        Limits::MAX_ENTRY_BYTES,
        Limits::MAX_TOTAL_BYTES,
        Limits::MAX_IWA_STREAM_BYTES,
    )?;
    assert!(Package::from_bytes_with_limits(&source, input_limited).is_err());
    Ok(())
}

#[test]
fn export_native_hidden_axes_focus_candidates_when_requested() -> TestResult {
    let visible_source = source_bytes(NATIVE_VISIBLE_SOURCE)?;
    let visible = Package::from_bytes(&visible_source)?;
    let focused_axes = native_hidden_axes()?;
    let focused = visible
        .edit_body_table_hidden_axes(0usize)?
        .set(focused_axes.clone())
        .commit()?;
    let focused_bytes = exact_bytes(focused.package())?;
    assert_eq!(
        Package::from_bytes(&focused_bytes)?.body_table_hidden_axes(0usize)?,
        focused_axes
    );
    export_candidate("body-table-hidden-axes-focused.pages", &focused_bytes)?;

    let native_source = source_bytes(NATIVE_HIDDEN_SOURCE)?;
    let native = Package::from_bytes(&native_source)?;
    let clear = native
        .edit_body_table_hidden_axes(0usize)?
        .clear()
        .commit()?;
    let clear_bytes = exact_bytes(clear.package())?;
    assert_eq!(
        Package::from_bytes(&clear_bytes)?.body_table_hidden_axes(0usize)?,
        HiddenAxes::empty()
    );
    export_candidate("body-table-hidden-axes-native-cleared.pages", &clear_bytes)?;

    let reset = native
        .edit_body_table_hidden_axes(0usize)?
        .reset()
        .commit()?;
    let reset_bytes = exact_bytes(reset.package())?;
    assert_eq!(
        Package::from_bytes(&reset_bytes)?.body_table_hidden_axes(0usize)?,
        HiddenAxes::empty()
    );
    export_candidate("body-table-hidden-axes-native-reset.pages", &reset_bytes)?;
    Ok(())
}
