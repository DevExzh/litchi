//! Source-built Pages regression coverage for the legacy comment writers.
//!
//! The public Pages builder supplies a Cities table, an unrelated table, and a
//! shape.  Pages' table-comment editor and the deprecated migration editor then
//! exercise root creation, reply creation, and copy-on-write reply updates.
//! The assertions below inspect only the physical package boundary: untouched
//! ZIP members and archive objects must survive, every newly allocated comment
//! object must have one UUID registry entry, and removed comment objects must
//! not leave stale registrations.

#![allow(deprecated)]

use std::collections::{BTreeMap, BTreeSet};
use std::io;

use litchi_iwa::comments::IWorkDrawableCommentEditor;
use litchi_iwa::pages::{PagesDocumentBuilder, PagesEditor, PagesTableInfo};
use litchi_iwa::shapes::{DrawablePoint, DrawableSize, Preset};
use litchi_iwa_archive::iwa::{Archive, ArchiveObject, SnappyStream};
use litchi_iwa_archive::package::Catalog;
use litchi_iwa_common::comment::{AuthorId, DrawableId, StorageId, Uuid};
use litchi_iwa_protos::{tsd, tsp};
use litchi_pages::{
    BodyTableSelector, Package as FocusedPagesPackage,
    table::hidden_axes::{AxisIndex, HiddenAxes},
};
use prost::Message as _;

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

const METADATA_COMPONENT: &str = "Index/Metadata.iwa";
const DOCUMENT_COMPONENT: &str = "Index/Document.iwa";
const AUTHOR_COMPONENT: &str = "Index/AnnotationAuthorStorage.iwa";
const PACKAGE_METADATA_MESSAGE_TYPE: u32 = 11_006;
const COMMENT_STORAGE_MESSAGE_TYPE: u32 = 3_056;
const ANNOTATION_AUTHOR_MESSAGE_TYPE: u32 = 212;
const ANNOTATION_AUTHOR_STORAGE_MESSAGE_TYPE: u32 = 213;
const NATIVE_COMMENT_BODY_MARKER: &str = "Pages hidden-axis native oracle — 2026-09-05";
const NATIVE_COMMENT_HIDDEN_SOURCE: &[u8] = include_bytes!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../test-data/iwork/pages/body-table-comment-hidden-native-saved.pages"
));
const COMMENT_OBJECT_MESSAGE_TYPES: &[u32] = &[
    COMMENT_STORAGE_MESSAGE_TYPE,
    ANNOTATION_AUTHOR_MESSAGE_TYPE,
    ANNOTATION_AUTHOR_STORAGE_MESSAGE_TYPE,
];

#[derive(Debug, Clone, PartialEq, Eq)]
struct UuidRegistration {
    component_identifier: u64,
    uuid: (u64, u64),
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct MetadataSnapshot {
    last_object_identifier: u64,
    object_uuid_map: BTreeMap<u64, Vec<UuidRegistration>>,
    component_identifiers: BTreeMap<String, u64>,
    component_payloads: BTreeMap<u64, Vec<u8>>,
    component_save_tokens: BTreeMap<u64, Option<u64>>,
    save_token: Option<u64>,
}

#[derive(Debug, Clone)]
struct ObjectRecord {
    component: String,
    object: ArchiveObject,
}

fn pages_shape_package() -> TestResult<(Vec<u8>, DrawableId)> {
    let mut pages = PagesDocumentBuilder::new().body_text("Body text").build()?;
    let shape = pages.add_body_shape(
        4,
        "A source-built shape",
        DrawablePoint { x: 96.0, y: 144.0 },
        DrawableSize {
            width: 180.0,
            height: 90.0,
        },
        Preset::Rectangle,
    )?;
    let drawable_id = DrawableId::from_raw(shape.drawable_object_id)?;
    Ok((pages.to_bytes()?, drawable_id))
}

fn pages_table_package() -> TestResult<(Vec<u8>, u64, u64, PagesTableInfo)> {
    let pages = PagesDocumentBuilder::new()
        .body_text("Body text")
        .body_table("Cities", 5, 2)
        .body_table("Unrelated", 2, 2)
        .build()?;
    let tables = pages.tables()?;
    let cities = tables
        .first()
        .ok_or_else(|| io::Error::other("source-built Cities table is missing"))?;
    let unrelated = tables
        .get(1)
        .ok_or_else(|| io::Error::other("source-built unrelated table is missing"))?;
    assert_eq!(cities.name, "Cities");
    assert_eq!(unrelated.name, "Unrelated");
    Ok((
        pages.to_bytes()?,
        cities.model_object_id,
        unrelated.model_object_id,
        unrelated.clone(),
    ))
}

fn component_archives(source: &[u8]) -> TestResult<BTreeMap<String, Archive>> {
    let catalog = Catalog::from_bytes(source)?;
    let mut archives = BTreeMap::new();
    for entry in catalog
        .iter()
        .filter(|entry| entry.name().ends_with(".iwa"))
    {
        let decompressed = SnappyStream::decompress(entry.data())?.into_bytes();
        let archive = Archive::parse(&decompressed)?;
        if archives.insert(entry.name().to_owned(), archive).is_some() {
            return Err(io::Error::other(format!("duplicate IWA member {}", entry.name())).into());
        }
    }
    Ok(archives)
}

fn object_records(source: &[u8]) -> TestResult<BTreeMap<u64, ObjectRecord>> {
    let mut records = BTreeMap::new();
    for (component, archive) in component_archives(source)? {
        for object in archive.objects {
            let identifier = object
                .archive_info
                .identifier
                .ok_or_else(|| io::Error::other("archive object has no identifier"))?;
            if records
                .insert(
                    identifier,
                    ObjectRecord {
                        component: component.clone(),
                        object,
                    },
                )
                .is_some()
            {
                return Err(io::Error::other(format!(
                    "duplicate native object identifier {identifier}"
                ))
                .into());
            }
        }
    }
    Ok(records)
}

fn object_types(source: &[u8]) -> TestResult<BTreeMap<u64, BTreeSet<u32>>> {
    Ok(object_records(source)?
        .into_iter()
        .map(|(identifier, record)| {
            (
                identifier,
                record
                    .object
                    .messages
                    .into_iter()
                    .map(|message| message.type_)
                    .collect(),
            )
        })
        .collect())
}

fn metadata_payload(source: &[u8]) -> TestResult<Vec<u8>> {
    let archive = component_archives(source)?
        .remove(METADATA_COMPONENT)
        .ok_or_else(|| io::Error::other("missing PackageMetadata component"))?;
    archive
        .objects
        .iter()
        .flat_map(|object| &object.messages)
        .find(|message| message.type_ == PACKAGE_METADATA_MESSAGE_TYPE)
        .map(|message| message.data.clone())
        .ok_or_else(|| io::Error::other("missing PackageMetadata payload").into())
}

fn metadata_snapshot(source: &[u8]) -> TestResult<MetadataSnapshot> {
    let metadata = tsp::PackageMetadata::decode(metadata_payload(source)?.as_slice())?;
    let mut object_uuid_map = BTreeMap::<u64, Vec<UuidRegistration>>::new();
    let mut component_identifiers = BTreeMap::new();
    let mut component_payloads = BTreeMap::new();
    let mut component_save_tokens = BTreeMap::new();
    for component in metadata.components {
        let component_identifier = component.identifier;
        let locator = component
            .locator
            .clone()
            .unwrap_or_else(|| component.preferred_locator.clone());
        component_identifiers.insert(locator, component_identifier);
        component_payloads.insert(component_identifier, component.encode_to_vec());
        component_save_tokens.insert(component_identifier, component.save_token);
        for entry in component.object_uuid_map_entries {
            object_uuid_map
                .entry(entry.identifier)
                .or_default()
                .push(UuidRegistration {
                    component_identifier,
                    uuid: (entry.uuid.lower, entry.uuid.upper),
                });
        }
    }
    for registrations in object_uuid_map.values_mut() {
        registrations.sort_unstable_by_key(|entry| (entry.component_identifier, entry.uuid));
    }
    Ok(MetadataSnapshot {
        last_object_identifier: metadata.last_object_identifier,
        object_uuid_map,
        component_identifiers,
        component_payloads,
        component_save_tokens,
        save_token: metadata.save_token,
    })
}

fn comment_storage_uuid(source: &[u8], identifier: u64) -> TestResult<(u64, u64)> {
    let object = object_records(source)?
        .remove(&identifier)
        .ok_or_else(|| io::Error::other(format!("missing comment object {identifier}")))?;
    let message = object
        .object
        .messages
        .iter()
        .find(|message| message.type_ == COMMENT_STORAGE_MESSAGE_TYPE)
        .ok_or_else(|| io::Error::other("object has no comment-storage payload"))?;
    let storage = tsd::CommentStorageArchive::decode(message.data.as_slice())?;
    let uuid = storage
        .storage_uuid
        .ok_or_else(|| io::Error::other("comment storage has no UUID"))?;
    Ok((uuid.lower, uuid.upper))
}

fn assert_comment_uuid_is_registered(source: &[u8], identifier: u64) -> TestResult<()> {
    let metadata = metadata_snapshot(source)?;
    let registrations = metadata
        .object_uuid_map
        .get(&identifier)
        .ok_or_else(|| io::Error::other(format!("object {identifier} has no UUID binding")))?;
    assert_eq!(
        registrations.len(),
        1,
        "comment object {identifier} has duplicate UUID bindings"
    );
    assert_ne!(registrations[0].uuid, (0, 0));
    let component = object_records(source)?
        .remove(&identifier)
        .ok_or_else(|| io::Error::other(format!("missing object {identifier}")))?
        .component;
    let locator = component
        .strip_prefix("Index/")
        .and_then(|name| name.strip_suffix(".iwa"))
        .ok_or_else(|| io::Error::other(format!("invalid component name {component}")))?;
    let expected_component = metadata
        .component_identifiers
        .get(locator)
        .copied()
        .ok_or_else(|| io::Error::other(format!("metadata lacks component {locator}")))?;
    assert_eq!(
        registrations[0].component_identifier, expected_component,
        "comment object {identifier} is registered under the wrong component"
    );
    assert_ne!(
        comment_storage_uuid(source, identifier)?,
        (0, 0),
        "comment object {identifier} has a zero storage UUID"
    );
    Ok(())
}

fn assert_object_uuid_is_removed(source: &[u8], identifier: u64) -> TestResult<()> {
    assert!(
        !object_records(source)?.contains_key(&identifier),
        "removed object {identifier} remains in an archive"
    );
    assert!(
        !metadata_snapshot(source)?
            .object_uuid_map
            .contains_key(&identifier),
        "removed object {identifier} retains an ObjectUUIDMap entry"
    );
    Ok(())
}

fn assert_comment_object_removed(source: &[u8], identifier: u64) -> TestResult<()> {
    assert_object_uuid_is_removed(source, identifier)
}

fn ids_with_message_type(source: &[u8], message_type: u32) -> TestResult<BTreeSet<u64>> {
    Ok(object_records(source)?
        .into_iter()
        .filter_map(|(identifier, record)| {
            record
                .object
                .messages
                .iter()
                .any(|message| message.type_ == message_type)
                .then_some(identifier)
        })
        .collect())
}

fn assert_new_type_ids_have_uuid_bindings(
    before: &[u8],
    after: &[u8],
    message_type: u32,
) -> TestResult<()> {
    let before_ids = ids_with_message_type(before, message_type)?;
    let after_ids = ids_with_message_type(after, message_type)?;
    let metadata = metadata_snapshot(after)?;
    for identifier in after_ids.difference(&before_ids) {
        let registrations = metadata.object_uuid_map.get(identifier).ok_or_else(|| {
            io::Error::other(format!(
                "new type-{message_type} object {identifier} has no ObjectUUIDMap entry"
            ))
        })?;
        assert_eq!(
            registrations.len(),
            1,
            "new type-{message_type} object {identifier} has duplicate UUID bindings"
        );
        let component = object_records(after)?
            .remove(identifier)
            .ok_or_else(|| io::Error::other(format!("missing object {identifier}")))?
            .component;
        let locator = component
            .strip_prefix("Index/")
            .and_then(|name| name.strip_suffix(".iwa"))
            .ok_or_else(|| io::Error::other(format!("invalid component name {component}")))?;
        let expected_component = metadata
            .component_identifiers
            .get(locator)
            .copied()
            .ok_or_else(|| io::Error::other(format!("metadata lacks component {locator}")))?;
        assert_eq!(
            registrations[0].component_identifier, expected_component,
            "new type-{message_type} object {identifier} is registered under the wrong component"
        );
    }
    Ok(())
}

fn comment_ids(source: &[u8]) -> TestResult<BTreeSet<u64>> {
    Ok(object_types(source)?
        .into_iter()
        .filter_map(|(identifier, types)| {
            types
                .iter()
                .any(|type_| COMMENT_OBJECT_MESSAGE_TYPES.contains(type_))
                .then_some(identifier)
        })
        .collect())
}

fn all_object_ids(source: &[u8]) -> TestResult<BTreeSet<u64>> {
    Ok(object_records(source)?.into_keys().collect())
}

fn assert_watermark_and_registry_shape(before: &[u8], after: &[u8]) -> TestResult<()> {
    let before_objects = all_object_ids(before)?;
    let after_objects = all_object_ids(after)?;
    let before_max = before_objects.iter().copied().max().unwrap_or(0);
    for identifier in after_objects.difference(&before_objects) {
        assert!(
            *identifier > before_max,
            "new object {identifier} was allocated below the source watermark {before_max}"
        );
    }

    let before_metadata = metadata_snapshot(before)?;
    let after_metadata = metadata_snapshot(after)?;
    assert_eq!(
        before_metadata
            .component_payloads
            .keys()
            .collect::<Vec<_>>(),
        after_metadata.component_payloads.keys().collect::<Vec<_>>(),
        "comment mutation changed the PackageMetadata component set"
    );
    if let (Some(before), Some(after)) = (before_metadata.save_token, after_metadata.save_token) {
        assert!(
            after >= before,
            "PackageMetadata save token regressed from {before} to {after}"
        );
    }
    let mut changed_component_identifiers = BTreeSet::new();
    for (identifier, registrations) in &before_metadata.object_uuid_map {
        if after_metadata.object_uuid_map.get(identifier) != Some(registrations) {
            changed_component_identifiers.extend(
                registrations
                    .iter()
                    .map(|registration| registration.component_identifier),
            );
        }
    }
    for (identifier, registrations) in &after_metadata.object_uuid_map {
        if before_metadata.object_uuid_map.get(identifier) != Some(registrations) {
            changed_component_identifiers.extend(
                registrations
                    .iter()
                    .map(|registration| registration.component_identifier),
            );
        }
    }
    for component_identifier in before_metadata.component_payloads.keys() {
        let before_payload = &before_metadata.component_payloads[component_identifier];
        let after_payload = &after_metadata.component_payloads[component_identifier];
        if changed_component_identifiers.contains(component_identifier) {
            if let (Some(before), Some(after)) = (
                before_metadata.component_save_tokens[component_identifier],
                after_metadata.component_save_tokens[component_identifier],
            ) {
                assert!(
                    after >= before,
                    "component {component_identifier} save token regressed from {before} to {after}"
                );
            }
        } else {
            assert_eq!(
                before_payload, after_payload,
                "unrelated PackageMetadata component {component_identifier} changed"
            );
        }
    }
    let after_max = after_objects.iter().copied().max().unwrap_or(0);
    assert!(
        after_metadata.last_object_identifier >= after_max,
        "PackageMetadata watermark {} is below native object {}",
        after_metadata.last_object_identifier,
        after_max
    );

    let before_comments = comment_ids(before)?;
    for (identifier, registrations) in &before_metadata.object_uuid_map {
        if !after_objects.contains(identifier) && before_comments.contains(identifier) {
            assert!(
                !after_metadata.object_uuid_map.contains_key(identifier),
                "removed comment object {identifier} left a stale ObjectUUIDMap entry"
            );
            continue;
        }
        assert_eq!(
            after_metadata.object_uuid_map.get(identifier),
            Some(registrations),
            "unrelated ObjectUUIDMap entry for object {identifier} changed"
        );
    }
    for (identifier, registrations) in &after_metadata.object_uuid_map {
        assert!(
            after_objects.contains(identifier),
            "ObjectUUIDMap contains stale object {identifier}"
        );
        assert_eq!(
            registrations.len(),
            1,
            "object {identifier} has duplicate ObjectUUIDMap entries"
        );
    }
    Ok(())
}

fn assert_unrelated_content_preserved(
    before: &[u8],
    after: &[u8],
    drawable_id: DrawableId,
) -> TestResult<()> {
    let before_catalog = Catalog::from_bytes(before)?;
    let after_catalog = Catalog::from_bytes(after)?;
    let before_entries = before_catalog
        .iter()
        .map(|entry| (entry.name().to_owned(), entry.data().to_vec()))
        .collect::<BTreeMap<_, _>>();
    let after_entries = after_catalog
        .iter()
        .map(|entry| (entry.name().to_owned(), entry.data().to_vec()))
        .collect::<BTreeMap<_, _>>();
    assert_eq!(
        before_entries.keys().collect::<Vec<_>>(),
        after_entries.keys().collect::<Vec<_>>(),
        "comment mutation changed the package member set"
    );
    for (name, data) in before_entries {
        if !matches!(
            name.as_str(),
            METADATA_COMPONENT | DOCUMENT_COMPONENT | AUTHOR_COMPONENT
        ) {
            assert_eq!(
                after_entries.get(&name),
                Some(&data),
                "unrelated package member {name} changed"
            );
        }
    }

    let before_objects = object_records(before)?;
    let after_objects = object_records(after)?;
    for (identifier, before_record) in before_objects {
        let Some(after_record) = after_objects.get(&identifier) else {
            continue;
        };
        if identifier == drawable_id.get()
            || before_record.component == METADATA_COMPONENT
            || before_record
                .object
                .messages
                .iter()
                .any(|message| COMMENT_OBJECT_MESSAGE_TYPES.contains(&message.type_))
            || after_record
                .object
                .messages
                .iter()
                .any(|message| COMMENT_OBJECT_MESSAGE_TYPES.contains(&message.type_))
        {
            continue;
        }
        assert_eq!(
            before_record.component, after_record.component,
            "unrelated object {identifier} moved components"
        );
        assert!(
            before_record
                .object
                .same_content_ignoring_offsets(&after_record.object),
            "unrelated object {identifier} payload or archive metadata changed"
        );
    }
    Ok(())
}

#[derive(Debug, Clone, PartialEq)]
struct TableSnapshot {
    info: PagesTableInfo,
    cells: Vec<((usize, usize), litchi_iwa::pages::PagesCellValue)>,
    comments: Vec<((usize, usize), litchi_iwa_common::comment::Comment)>,
}

fn table_snapshot(editor: &PagesEditor, model_object_id: u64) -> TestResult<TableSnapshot> {
    let table = editor.table(model_object_id)?;
    Ok(TableSnapshot {
        info: table.info.clone(),
        cells: table
            .iter_cells()
            .map(|(position, value)| (position, value.clone()))
            .collect(),
        comments: table
            .iter_comments()
            .map(|(position, comment)| (position, comment.clone()))
            .collect(),
    })
}

fn assert_non_iwa_members_preserved(before: &[u8], after: &[u8]) -> TestResult<()> {
    let before_catalog = Catalog::from_bytes(before)?;
    let after_catalog = Catalog::from_bytes(after)?;
    let before_entries = before_catalog
        .iter()
        .filter(|entry| !entry.name().ends_with(".iwa"))
        .map(|entry| (entry.name().to_owned(), entry.data().to_vec()))
        .collect::<BTreeMap<_, _>>();
    let after_entries = after_catalog
        .iter()
        .filter(|entry| !entry.name().ends_with(".iwa"))
        .map(|entry| (entry.name().to_owned(), entry.data().to_vec()))
        .collect::<BTreeMap<_, _>>();
    assert_eq!(
        before_entries, after_entries,
        "unrelated package members changed"
    );
    Ok(())
}

fn object_references(object: &ArchiveObject) -> BTreeSet<u64> {
    object
        .archive_info
        .message_infos
        .iter()
        .flat_map(|message| {
            message.object_references.iter().copied().chain(
                message
                    .field_infos
                    .iter()
                    .flat_map(|field| field.object_references.iter().copied()),
            )
        })
        .collect()
}

fn reachable_object_ids(source: &[u8], root: u64) -> TestResult<BTreeSet<u64>> {
    let records = object_records(source)?;
    let mut pending = vec![root];
    let mut reachable = BTreeSet::new();
    while let Some(identifier) = pending.pop() {
        if !reachable.insert(identifier) {
            continue;
        }
        let Some(object) = records.get(&identifier) else {
            // ArchiveInfo can carry references into a component that is not
            // materialized in this package projection.  Such an external
            // edge is irrelevant to the physical preservation comparison.
            continue;
        };
        pending.extend(object_references(&object.object));
    }
    Ok(reachable)
}

fn assert_unrelated_table_graph_preserved(
    before: &[u8],
    after: &[u8],
    changed_table_id: u64,
    unrelated_table_id: u64,
) -> TestResult<()> {
    let before_changed = reachable_object_ids(before, changed_table_id)?;
    let before_unrelated = reachable_object_ids(before, unrelated_table_id)?;
    let before_records = object_records(before)?;
    let after_records = object_records(after)?;
    for identifier in before_unrelated.difference(&before_changed) {
        let before_record = before_records
            .get(identifier)
            .ok_or_else(|| io::Error::other(format!("missing source object {identifier}")))?;
        let after_record = after_records.get(identifier).ok_or_else(|| {
            io::Error::other(format!(
                "comment mutation removed unrelated table object {identifier}"
            ))
        })?;
        assert_eq!(
            before_record.component, after_record.component,
            "comment mutation moved unrelated table object {identifier}"
        );
        assert!(
            before_record
                .object
                .same_content_ignoring_offsets(&after_record.object),
            "comment mutation changed unrelated table object {identifier}"
        );
    }
    Ok(())
}

fn assert_thread(
    source: &[u8],
    drawable_id: DrawableId,
    expected_root: &str,
    expected_replies: &[&str],
) -> TestResult<(StorageId, Vec<StorageId>)> {
    let editor = IWorkDrawableCommentEditor::from_bytes(source)?;
    let root = editor
        .comment(drawable_id)?
        .ok_or_else(|| io::Error::other("drawable has no root comment"))?;
    assert_eq!(root.comment.text, expected_root);
    let replies = editor.replies(drawable_id)?;
    assert_eq!(
        replies
            .iter()
            .map(|reply| reply.comment.text.as_str())
            .collect::<Vec<_>>(),
        expected_replies
    );
    Ok((
        root.storage_id,
        replies.into_iter().map(|reply| reply.storage_id).collect(),
    ))
}

#[test]
fn source_built_pages_comment_lifecycle_updates_physical_registry() -> TestResult {
    let (source, drawable_id) = pages_shape_package()?;
    let source_objects = all_object_ids(&source)?;

    let mut editor = IWorkDrawableCommentEditor::from_bytes(&source)?;
    assert_eq!(
        editor.application(),
        litchi_iwa::application::Application::Pages
    );
    let source_drawables = editor.drawables()?;
    assert!(
        source_drawables
            .iter()
            .any(|drawable| drawable.id == drawable_id),
        "source-built shape is not visible to the legacy drawable editor"
    );

    editor.set_comment(drawable_id, "root text")?;
    let root_bytes = editor.to_bytes()?;
    let (root_id, replies) = assert_thread(&root_bytes, drawable_id, "root text", &[])?;
    assert!(replies.is_empty());
    assert!(root_id.get() > source_objects.iter().copied().max().unwrap_or(0));
    assert_comment_uuid_is_registered(&root_bytes, root_id.get())?;
    assert_new_type_ids_have_uuid_bindings(&source, &root_bytes, ANNOTATION_AUTHOR_MESSAGE_TYPE)?;
    assert_watermark_and_registry_shape(&source, &root_bytes)?;
    assert_unrelated_content_preserved(&source, &root_bytes, drawable_id)?;

    editor = IWorkDrawableCommentEditor::from_bytes(&root_bytes)?;
    let reply_id = editor.add_reply(drawable_id, "first reply")?;
    let reply_bytes = editor.to_bytes()?;
    let (reply_root_id, replies) =
        assert_thread(&reply_bytes, drawable_id, "root text", &["first reply"])?;
    assert_ne!(
        reply_root_id, root_id,
        "reply insertion did not clone the root"
    );
    assert_eq!(replies, vec![reply_id]);
    assert_comment_uuid_is_registered(&reply_bytes, reply_root_id.get())?;
    assert_comment_uuid_is_registered(&reply_bytes, reply_id.get())?;
    assert_comment_object_removed(&reply_bytes, root_id.get())?;
    assert_eq!(
        comment_storage_uuid(&reply_bytes, reply_root_id.get())?,
        comment_storage_uuid(&root_bytes, root_id.get())?,
        "reply root clone changed the source storage UUID"
    );
    assert_watermark_and_registry_shape(&root_bytes, &reply_bytes)?;
    assert_unrelated_content_preserved(&root_bytes, &reply_bytes, drawable_id)?;

    editor = IWorkDrawableCommentEditor::from_bytes(&reply_bytes)?;
    let replaced_reply_id = editor.set_reply(drawable_id, reply_id, "updated reply")?;
    let replaced_bytes = editor.to_bytes()?;
    let (replaced_root_id, _replies) = assert_thread(
        &replaced_bytes,
        drawable_id,
        "root text",
        &["updated reply"],
    )?;
    assert_ne!(replaced_root_id, reply_root_id);
    assert_ne!(replaced_reply_id, reply_id);
    assert_comment_uuid_is_registered(&replaced_bytes, replaced_root_id.get())?;
    assert_comment_uuid_is_registered(&replaced_bytes, replaced_reply_id.get())?;
    assert_comment_object_removed(&replaced_bytes, reply_root_id.get())?;
    assert_comment_object_removed(&replaced_bytes, reply_id.get())?;
    assert_watermark_and_registry_shape(&reply_bytes, &replaced_bytes)?;
    assert_unrelated_content_preserved(&reply_bytes, &replaced_bytes, drawable_id)?;

    // The final bytes must be a real reopen, not just the in-memory editor's
    // projection.  This also checks that the newly registered IDs are usable
    // after the package writer has rebuilt the touched archives.
    let reopened = IWorkDrawableCommentEditor::from_bytes(&replaced_bytes)?;
    let reopened_root = reopened
        .comment(drawable_id)?
        .ok_or_else(|| io::Error::other("reopened package lost root comment"))?;
    assert_eq!(reopened_root.storage_id, replaced_root_id);
    assert_eq!(reopened_root.comment.text, "root text");
    assert_eq!(
        reopened.replies(drawable_id)?[0].storage_id,
        replaced_reply_id
    );
    Ok(())
}

#[test]
fn source_built_pages_comment_failures_are_atomic() -> TestResult<()> {
    let (source, drawable_id) = pages_shape_package()?;
    let mut editor = IWorkDrawableCommentEditor::from_bytes(&source)?;
    editor.set_comment(drawable_id, "root text")?;
    let reply_id = editor.add_reply(drawable_id, "reply text")?;
    let before = editor.to_bytes()?;
    let before_metadata = metadata_snapshot(&before)?;

    let invalid_drawable = DrawableId::from_raw(u64::MAX)?;
    assert!(editor.set_comment(invalid_drawable, "must fail").is_err());
    assert_eq!(editor.to_bytes()?, before);
    assert_eq!(metadata_snapshot(&editor.to_bytes()?)?, before_metadata);

    let invalid_reply = StorageId::from_raw(u64::MAX)?;
    assert!(
        editor
            .set_reply(drawable_id, invalid_reply, "must fail")
            .is_err()
    );
    assert_eq!(editor.to_bytes()?, before);
    assert_eq!(metadata_snapshot(&editor.to_bytes()?)?, before_metadata);

    let (root_id, replies) = assert_thread(
        &editor.to_bytes()?,
        drawable_id,
        "root text",
        &["reply text"],
    )?;
    assert_eq!(replies, vec![reply_id]);
    assert_comment_uuid_is_registered(&editor.to_bytes()?, root_id.get())?;
    assert_comment_uuid_is_registered(&editor.to_bytes()?, reply_id.get())?;
    Ok(())
}

#[test]
fn source_built_pages_table_comment_lifecycle_preserves_unrelated_table() -> TestResult<()> {
    let (source, cities_id, unrelated_id, unrelated_info) = pages_table_package()?;
    let baseline_editor = PagesEditor::from_bytes(&source)?;
    let unrelated_before = table_snapshot(&baseline_editor, unrelated_id)?;
    assert_eq!(unrelated_before.info, unrelated_info);

    let mut editor = baseline_editor;
    editor.set_table_cell_comment(cities_id, 1, 1, "Cities root")?;
    let root_bytes = editor.to_bytes()?;
    let root = editor
        .table_cell_comment(cities_id, 1, 1)?
        .ok_or_else(|| io::Error::other("Cities comment was not created"))?;
    assert_eq!(root.comment.text, "Cities root");
    assert_comment_uuid_is_registered(&root_bytes, root.storage_id.get())?;
    assert_new_type_ids_have_uuid_bindings(&source, &root_bytes, ANNOTATION_AUTHOR_MESSAGE_TYPE)?;
    assert_watermark_and_registry_shape(&source, &root_bytes)?;
    assert_non_iwa_members_preserved(&source, &root_bytes)?;
    assert_unrelated_table_graph_preserved(&source, &root_bytes, cities_id, unrelated_id)?;
    assert_eq!(
        table_snapshot(&editor, unrelated_id)?,
        unrelated_before,
        "creating a Cities comment changed the unrelated table"
    );

    let reply_id = editor.add_table_cell_comment_reply(cities_id, 1, 1, "Cities reply")?;
    let reply_bytes = editor.to_bytes()?;
    let reply_root = editor
        .table_cell_comment(cities_id, 1, 1)?
        .ok_or_else(|| io::Error::other("Cities root disappeared after reply creation"))?;
    let replies = editor.table_cell_comment_replies(cities_id, 1, 1)?;
    assert_eq!(replies.len(), 1);
    assert_eq!(replies[0].storage_id.get(), reply_id);
    assert_eq!(replies[0].comment.text, "Cities reply");
    assert_ne!(reply_root.storage_id, root.storage_id);
    assert_comment_uuid_is_registered(&reply_bytes, reply_root.storage_id.get())?;
    assert_comment_uuid_is_registered(&reply_bytes, reply_id)?;
    assert_comment_object_removed(&reply_bytes, root.storage_id.get())?;
    assert_watermark_and_registry_shape(&root_bytes, &reply_bytes)?;
    assert_non_iwa_members_preserved(&root_bytes, &reply_bytes)?;
    assert_unrelated_table_graph_preserved(&root_bytes, &reply_bytes, cities_id, unrelated_id)?;
    assert_eq!(
        table_snapshot(&editor, unrelated_id)?,
        unrelated_before,
        "creating a Cities reply changed the unrelated table"
    );

    let replaced_reply_id =
        editor.set_table_cell_comment_reply(cities_id, 1, 1, reply_id, "Cities reply revised")?;
    let replaced_bytes = editor.to_bytes()?;
    let replaced_root = editor
        .table_cell_comment(cities_id, 1, 1)?
        .ok_or_else(|| io::Error::other("Cities root disappeared after reply update"))?;
    let replaced_replies = editor.table_cell_comment_replies(cities_id, 1, 1)?;
    assert_eq!(replaced_replies.len(), 1);
    assert_eq!(replaced_replies[0].storage_id.get(), replaced_reply_id);
    assert_eq!(replaced_replies[0].comment.text, "Cities reply revised");
    assert_ne!(replaced_root.storage_id, reply_root.storage_id);
    assert_ne!(replaced_reply_id, reply_id);
    assert_comment_uuid_is_registered(&replaced_bytes, replaced_root.storage_id.get())?;
    assert_comment_uuid_is_registered(&replaced_bytes, replaced_reply_id)?;
    assert_comment_object_removed(&replaced_bytes, reply_root.storage_id.get())?;
    assert_comment_object_removed(&replaced_bytes, reply_id)?;
    assert_watermark_and_registry_shape(&reply_bytes, &replaced_bytes)?;
    assert_non_iwa_members_preserved(&reply_bytes, &replaced_bytes)?;
    assert_unrelated_table_graph_preserved(&reply_bytes, &replaced_bytes, cities_id, unrelated_id)?;
    assert_eq!(
        table_snapshot(&editor, unrelated_id)?,
        unrelated_before,
        "updating a Cities reply changed the unrelated table"
    );

    let reopened = PagesEditor::from_bytes(&replaced_bytes)?;
    assert_eq!(table_snapshot(&reopened, unrelated_id)?, unrelated_before);
    let reopened_root = reopened
        .table_cell_comment(cities_id, 1, 1)?
        .ok_or_else(|| io::Error::other("reopened Pages package lost Cities comment"))?;
    assert_eq!(reopened_root.comment.text, "Cities root");
    assert_eq!(reopened_root.storage_id, replaced_root.storage_id);
    assert_eq!(
        reopened.table_cell_comment_replies(cities_id, 1, 1)?[0]
            .storage_id
            .get(),
        replaced_reply_id
    );

    let generated_author_ids =
        ids_with_message_type(&replaced_bytes, ANNOTATION_AUTHOR_MESSAGE_TYPE)?
            .difference(&ids_with_message_type(
                &source,
                ANNOTATION_AUTHOR_MESSAGE_TYPE,
            )?)
            .copied()
            .collect::<BTreeSet<_>>();
    let mut cleared = reopened;
    cleared.clear_table_cell_comment(cities_id, 1, 1)?;
    let cleared_bytes = cleared.to_bytes()?;
    assert!(cleared.table_cell_comment(cities_id, 1, 1)?.is_none());
    for identifier in generated_author_ids {
        assert_object_uuid_is_removed(&cleared_bytes, identifier)?;
    }
    assert_watermark_and_registry_shape(&replaced_bytes, &cleared_bytes)?;
    assert_non_iwa_members_preserved(&replaced_bytes, &cleared_bytes)?;
    Ok(())
}

#[test]
fn source_built_pages_table_comment_failures_are_atomic() -> TestResult<()> {
    let (source, cities_id, unrelated_id, _) = pages_table_package()?;
    let mut editor = PagesEditor::from_bytes(&source)?;
    editor.set_table_cell_comment(cities_id, 1, 1, "Cities root")?;
    let reply_id = editor.add_table_cell_comment_reply(cities_id, 1, 1, "Cities reply")?;
    let before = editor.to_bytes()?;
    let before_metadata = metadata_snapshot(&before)?;
    let before_unrelated = table_snapshot(&editor, unrelated_id)?;

    assert!(
        editor
            .set_table_cell_comment_reply(cities_id, 1, 1, u64::MAX, "must fail")
            .is_err()
    );
    assert_eq!(editor.to_bytes()?, before);
    assert_eq!(metadata_snapshot(&editor.to_bytes()?)?, before_metadata);
    assert_eq!(table_snapshot(&editor, unrelated_id)?, before_unrelated);

    assert!(
        editor
            .remove_table_cell_comment_reply(cities_id, 1, 1, u64::MAX)
            .is_err()
    );
    assert_eq!(editor.to_bytes()?, before);
    assert_eq!(metadata_snapshot(&editor.to_bytes()?)?, before_metadata);
    assert_eq!(table_snapshot(&editor, unrelated_id)?, before_unrelated);
    assert_eq!(
        editor.table_cell_comment_replies(cities_id, 1, 1)?[0]
            .storage_id
            .get(),
        reply_id
    );
    Ok(())
}

#[test]
fn native_saved_pages_hidden_axes_preserve_comment_thread_and_exact_reencode() -> TestResult {
    let source = NATIVE_COMMENT_HIDDEN_SOURCE;
    let focused = FocusedPagesPackage::from_bytes(source)?;
    let expected_axes = HiddenAxes::new([AxisIndex::row(2), AxisIndex::column(1)])?;
    assert_eq!(focused.body_table_hidden_axes(0usize)?, expected_axes);
    assert_eq!(
        focused.body_table_hidden_axes(BodyTableSelector::name("Table 1"))?,
        expected_axes
    );

    let mut focused_bytes = Vec::new();
    focused.write_to(&mut focused_bytes)?;
    assert_eq!(focused_bytes, source);

    let editor = PagesEditor::from_bytes(source)?;
    assert!(editor.body_text()?.contains(NATIVE_COMMENT_BODY_MARKER));
    let tables = editor.tables()?;
    assert_eq!(tables.len(), 1);
    let table = &tables[0];
    assert_eq!(table.name, "Table 1");
    assert_eq!((table.rows, table.columns), (5, 4));

    let root = editor
        .table_cell_comment(table.model_object_id, 1, 1)?
        .ok_or_else(|| io::Error::other("native saved B2 root comment is missing"))?;
    assert_eq!(root.comment.text, "Native hidden-axis control root");
    assert_eq!(root.comment.author_id, Some(AuthorId::from_raw(1_733_524)?));
    assert_eq!(
        root.comment.creation_date_seconds,
        Some(810_628_451.2089901)
    );
    assert_eq!(
        root.comment.storage_uuid,
        Some(Uuid::from_parts(
            9_603_170_737_480_227_315,
            12_444_760_541_495_578_553
        )?),
    );

    let replies = editor.table_cell_comment_replies(table.model_object_id, 1, 1)?;
    assert_eq!(replies.len(), 1);
    let reply = &replies[0];
    assert_eq!(reply.root_storage_id, root.storage_id);
    assert_eq!(reply.comment.text, "Native hidden-axis control reply");
    assert_eq!(
        reply.comment.author_id,
        Some(AuthorId::from_raw(1_733_524)?)
    );
    assert_eq!(
        reply.comment.creation_date_seconds,
        Some(810_628_451.3782489)
    );
    assert_eq!(
        reply.comment.storage_uuid,
        Some(Uuid::from_parts(
            3_983_970_677_132_734_838,
            13_341_971_900_969_325_204
        )?),
    );
    assert_eq!(root.comment.reply_ids.as_ref(), [reply.storage_id]);

    let reopened = PagesEditor::from_bytes(source)?;
    let reopened_root = reopened
        .table_cell_comment(table.model_object_id, 1, 1)?
        .ok_or_else(|| io::Error::other("reopened native saved B2 root comment is missing"))?;
    assert_eq!(reopened_root, root);
    assert_eq!(
        reopened.table_cell_comment_replies(table.model_object_id, 1, 1)?,
        replies
    );
    assert_eq!(reopened.to_bytes()?, source);
    Ok(())
}
