//! Focused-vs-generic parity for Keynote drawable comments.
//!
//! The focused package owns semantic selectors and keeps native identities
//! private.  The deprecated generic editor is used here only as a migration
//! oracle: it exposes the native storage IDs needed to prove COW, physical
//! component ownership, archive census, and PackageMetadata behavior.  The
//! test never calls the seven KeynoteEditor drawable-comment wrappers.

#![allow(deprecated)]

use std::collections::{BTreeMap, BTreeSet};
use std::io;

#[path = "../../litchi-keynote/tests/support/drawable_comment_fixtures.rs"]
mod fixtures;

use fixtures::{DrawableTarget, RelocatedFixture};
use litchi_iwa::comments::IWorkDrawableCommentEditor;
use litchi_iwa_archive::iwa::{Archive, SnappyStream};
use litchi_iwa_archive::package::Catalog;
use litchi_iwa_common::{
    comment::{DrawableId, StorageId},
    wire::WireView,
};
use litchi_iwa_protos::{tsd, tsp};
use litchi_keynote::{DrawableSelector, Package, ReplySelector, SlideSelector};
use prost::Message as _;

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

const METADATA_COMPONENT: &str = "Index/Metadata.iwa";
const COMMENT_OBJECT_MESSAGE_TYPES: &[u32] = &[
    fixtures::COMMENT_STORAGE_MESSAGE_TYPE,
    212, // annotation-author object
    213, // annotation-author storage object
];

#[derive(Debug, Clone, PartialEq, Eq)]
struct MetadataSnapshot {
    last_object_identifier: u64,
    save_token: Option<u64>,
    components: BTreeMap<u64, ComponentMetadataSnapshot>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct ComponentMetadataSnapshot {
    locator: String,
    save_token: Option<u64>,
    external_references: Vec<(u64, Option<u64>, Option<bool>)>,
    versioned_external_references: Vec<(u64, Option<u64>, Option<bool>)>,
    object_uuid_map_entries: Vec<Vec<u8>>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct ComponentMetadataInvariant {
    locator: String,
    external_references: Vec<(u64, Option<u64>, Option<bool>)>,
    versioned_external_references: Vec<(u64, Option<u64>, Option<bool>)>,
    object_uuid_map_entries: Vec<Vec<u8>>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct UuidRegistration {
    component_identifier: u64,
    uuid: (u64, u64),
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct ThreadSnapshot {
    root: Option<(String, Option<u64>, bool)>,
    replies: Vec<(String, Option<u64>, bool)>,
}

fn exact_bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

fn metadata_payload(source: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(source)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == METADATA_COMPONENT)
        .ok_or_else(|| io::Error::other("missing PackageMetadata component"))?;
    let stream = SnappyStream::decompress(entry.data())?.into_bytes();
    let archive = Archive::parse(&stream)?;
    archive
        .objects
        .iter()
        .flat_map(|object| &object.messages)
        .find(|message| message.type_ == fixtures::PACKAGE_METADATA_MESSAGE_TYPE)
        .map(|message| message.data.clone())
        .ok_or_else(|| io::Error::other("missing PackageMetadata payload").into())
}

fn metadata_snapshot(source: &[u8]) -> TestResult<MetadataSnapshot> {
    let metadata = tsp::PackageMetadata::decode(metadata_payload(source)?.as_slice())?;
    let mut components = BTreeMap::new();
    for component in metadata.components {
        let locator = component.locator.unwrap_or(component.preferred_locator);
        let mut external_references = component
            .external_references
            .into_iter()
            .map(|reference| {
                (
                    reference.component_identifier,
                    reference.object_identifier,
                    reference.is_weak,
                )
            })
            .collect::<Vec<_>>();
        external_references.sort_unstable();
        let mut versioned_external_references = component
            .versioned_external_references
            .into_iter()
            .map(|reference| {
                (
                    reference.component_identifier,
                    reference.object_identifier,
                    reference.is_weak,
                )
            })
            .collect::<Vec<_>>();
        versioned_external_references.sort_unstable();
        let mut object_uuid_map_entries = component
            .object_uuid_map_entries
            .into_iter()
            .map(|entry| entry.encode_to_vec())
            .collect::<Vec<_>>();
        object_uuid_map_entries.sort_unstable();
        components.insert(
            component.identifier,
            ComponentMetadataSnapshot {
                locator,
                save_token: component.save_token,
                external_references,
                versioned_external_references,
                object_uuid_map_entries,
            },
        );
    }
    Ok(MetadataSnapshot {
        last_object_identifier: metadata.last_object_identifier,
        save_token: metadata.save_token,
        components,
    })
}

fn uuid_registrations(source: &[u8]) -> TestResult<BTreeMap<u64, Vec<UuidRegistration>>> {
    let metadata = metadata_snapshot(source)?;
    let mut registrations = BTreeMap::<u64, Vec<UuidRegistration>>::new();
    for (component_identifier, component) in metadata.components {
        for encoded in component.object_uuid_map_entries {
            let entry = tsp::ObjectUuidMapEntry::decode(encoded.as_slice())?;
            registrations
                .entry(entry.identifier)
                .or_default()
                .push(UuidRegistration {
                    component_identifier,
                    uuid: (entry.uuid.lower, entry.uuid.upper),
                });
        }
    }
    for entries in registrations.values_mut() {
        entries.sort_unstable_by_key(|entry| (entry.component_identifier, entry.uuid));
    }
    Ok(registrations)
}

fn comment_storage_uuid(source: &[u8], identifier: u64) -> TestResult<Option<(u64, u64)>> {
    let object = fixtures::component_archives(source)?
        .into_iter()
        .find_map(|(_, archive)| archive.object(identifier).cloned())
        .ok_or_else(|| io::Error::other(format!("missing comment object {identifier}")))?;
    let message = object
        .messages
        .iter()
        .find(|message| message.type_ == fixtures::COMMENT_STORAGE_MESSAGE_TYPE)
        .ok_or_else(|| io::Error::other("comment object has no storage payload"))?;
    Ok(tsd::CommentStorageArchive::decode(message.data.as_slice())?
        .storage_uuid
        .map(|uuid| (uuid.lower, uuid.upper)))
}

fn comment_author_identifier(source: &[u8], identifier: u64) -> TestResult<Option<u64>> {
    let object = fixtures::component_archives(source)?
        .into_iter()
        .find_map(|(_, archive)| archive.object(identifier).cloned())
        .ok_or_else(|| io::Error::other(format!("missing comment object {identifier}")))?;
    let message = object
        .messages
        .iter()
        .find(|message| message.type_ == fixtures::COMMENT_STORAGE_MESSAGE_TYPE)
        .ok_or_else(|| io::Error::other("comment object has no storage payload"))?;
    Ok(tsd::CommentStorageArchive::decode(message.data.as_slice())?
        .author
        .map(|author| author.identifier))
}

fn annotation_author_ids(source: &[u8]) -> TestResult<BTreeSet<u64>> {
    Ok(native_object_types(source)?
        .into_iter()
        .filter_map(|(identifier, types)| types.contains(&212).then_some(identifier))
        .collect())
}

fn annotation_author_storage_ids(source: &[u8]) -> TestResult<BTreeSet<u64>> {
    let mut identifiers = BTreeSet::new();
    for (_, archive) in fixtures::component_archives(source)? {
        for object in archive.objects {
            for message in object.messages {
                if message.type_ != 213 {
                    continue;
                }
                for field in WireView::parse(&message.data)?.fields() {
                    if field.number() != 1 {
                        continue;
                    }
                    let reference = tsp::Reference::decode(field.payload())?;
                    if reference.identifier == 0 {
                        return Err(io::Error::other(
                            "annotation-author storage contains a zero identifier",
                        )
                        .into());
                    }
                    identifiers.insert(reference.identifier);
                }
            }
        }
    }
    Ok(identifiers)
}

fn annotation_author_payload(source: &[u8], identifier: u64) -> TestResult<Vec<u8>> {
    fixtures::component_archives(source)?
        .into_iter()
        .find_map(|(_, archive)| {
            archive.object(identifier).and_then(|object| {
                object
                    .messages
                    .iter()
                    .find(|message| message.type_ == 212)
                    .map(|message| message.data.clone())
            })
        })
        .ok_or_else(|| {
            io::Error::other(format!("missing annotation-author object {identifier}")).into()
        })
}

fn selected_thread_author_payloads(
    source: &[u8],
    target: &DrawableTarget,
) -> TestResult<Vec<Vec<u8>>> {
    let editor = IWorkDrawableCommentEditor::from_bytes(source)?;
    let mut identifiers = Vec::new();
    if let Some(comment) = editor.comment(DrawableId::from_raw(target.identifier)?)? {
        if let Some(identifier) = comment.comment.author_id {
            identifiers.push(identifier);
        }
    }
    identifiers.extend(
        editor
            .replies(DrawableId::from_raw(target.identifier)?)?
            .into_iter()
            .filter_map(|reply| reply.comment.author_id),
    );
    identifiers
        .into_iter()
        .map(|identifier| annotation_author_payload(source, identifier.get()))
        .collect()
}

fn assert_selected_thread_author_metadata_equal(
    focused: &[u8],
    generic: &[u8],
    target: &DrawableTarget,
) -> TestResult<()> {
    let focused_authors = selected_thread_author_payloads(focused, target)?;
    let generic_authors = selected_thread_author_payloads(generic, target)?;
    assert_eq!(
        focused_authors, generic_authors,
        "focused and generic selected-thread author metadata differs"
    );
    Ok(())
}

fn assert_selected_thread_author_metadata_present_and_equal(
    focused: &[u8],
    generic: &[u8],
    target: &DrawableTarget,
) -> TestResult<()> {
    let focused_authors = selected_thread_author_payloads(focused, target)?;
    let generic_authors = selected_thread_author_payloads(generic, target)?;
    assert!(
        !focused_authors.is_empty(),
        "selected thread lost all author metadata"
    );
    assert_eq!(
        focused_authors, generic_authors,
        "focused and generic selected-thread author metadata differs"
    );
    Ok(())
}

fn assert_legacy_generic_author_retention(
    focused_before: &[u8],
    generic_before: &[u8],
    focused_after: &[u8],
    generic_after: &[u8],
    author_identifier: u64,
) -> TestResult<()> {
    let focused_before_authors = annotation_author_ids(focused_before)?;
    let generic_before_authors = annotation_author_ids(generic_before)?;
    let focused_authors = annotation_author_ids(focused_after)?;
    let generic_authors = annotation_author_ids(generic_after)?;
    assert!(
        generic_before_authors.contains(&author_identifier),
        "generic remove source does not contain the selected generated author"
    );
    assert!(
        !focused_authors.contains(&author_identifier),
        "focused remove retained the generated annotation author"
    );
    assert!(
        generic_authors.contains(&author_identifier),
        "generic remove did not exhibit the retained generated annotation author"
    );
    let mut unrelated_authors = generic_before_authors;
    unrelated_authors.remove(&author_identifier);
    assert!(
        unrelated_authors.is_subset(&focused_authors),
        "focused remove collected an unrelated native annotation author"
    );
    assert!(
        unrelated_authors.is_subset(&generic_authors),
        "generic remove dropped an unrelated native annotation author"
    );

    let focused_before_registry = annotation_author_storage_ids(focused_before)?;
    let generic_before_registry = annotation_author_storage_ids(generic_before)?;
    let focused_registry = annotation_author_storage_ids(focused_after)?;
    let generic_registry = annotation_author_storage_ids(generic_after)?;
    assert!(
        generic_before_registry.contains(&author_identifier),
        "generic remove source does not register the selected generated author"
    );
    assert!(
        !focused_registry.contains(&author_identifier),
        "focused remove left the generated author in annotation-author storage"
    );
    assert!(
        generic_registry.contains(&author_identifier),
        "generic remove did not retain the generated author registry entry"
    );
    assert_eq!(
        annotation_author_payload(generic_after, author_identifier)?,
        annotation_author_payload(generic_before, author_identifier)?,
        "generic author retention changed the generated author metadata"
    );
    if focused_before_authors.contains(&author_identifier) {
        assert!(
            focused_before_registry.contains(&author_identifier),
            "focused source author registry is inconsistent"
        );
    }
    Ok(())
}

fn assert_registered_uuid(source: &[u8], identifier: u64) -> TestResult<(u64, u64)> {
    let registrations = uuid_registrations(source)?;
    let entries = registrations.get(&identifier).ok_or_else(|| {
        io::Error::other(format!(
            "comment object {identifier} has no UUID registration"
        ))
    })?;
    assert_eq!(
        entries.len(),
        1,
        "comment object {identifier} has duplicate UUID registrations"
    );
    Ok(entries[0].uuid)
}

fn assert_cow_uuid_registration(
    before: &[u8],
    after: &[u8],
    old_identifier: u64,
    new_identifier: u64,
) -> TestResult<()> {
    if old_identifier == new_identifier {
        let before_registrations = uuid_registrations(before)?;
        let after_registrations = uuid_registrations(after)?;
        if let Some(before_registration) = before_registrations.get(&old_identifier) {
            assert_eq!(
                after_registrations.get(&old_identifier),
                Some(before_registration),
                "in-place comment edit changed its ObjectUUIDMap registration"
            );
        }
        assert_eq!(
            comment_storage_uuid(after, old_identifier)?,
            comment_storage_uuid(before, old_identifier)?,
            "in-place comment edit changed its payload UUID"
        );
        return Ok(());
    }
    let before_registrations = uuid_registrations(before)?;
    let after_registrations = uuid_registrations(after)?;
    let old_registrations = before_registrations
        .get(&old_identifier)
        .cloned()
        .unwrap_or_default();
    let old_survives = native_object_ids(after)?.contains(&old_identifier);
    if old_survives {
        if old_registrations.is_empty() {
            assert!(
                !after_registrations.contains_key(&new_identifier),
                "COW invented an ObjectUUIDMap registration for an unregistered source object"
            );
        } else {
            assert_eq!(
                after_registrations.get(&old_identifier),
                Some(&old_registrations),
                "COW did not preserve the registered source ObjectUUIDMap entry"
            );
            let new_registrations = after_registrations
                .get(&new_identifier)
                .ok_or_else(|| {
                    io::Error::other(format!(
                        "COW clone lost its ObjectUUIDMap entry (old={old_identifier}, new={new_identifier})"
                    ))
                })?;
            assert_eq!(
                new_registrations.len(),
                1,
                "COW clone has duplicate ObjectUUIDMap registrations"
            );
            assert_eq!(
                old_registrations.len(),
                1,
                "source comment object has duplicate ObjectUUIDMap registrations"
            );
            assert_eq!(
                new_registrations[0].component_identifier,
                old_registrations[0].component_identifier,
                "COW clone moved to a different component namespace"
            );
            assert_ne!(
                new_registrations[0].uuid, old_registrations[0].uuid,
                "COW clone reused the source ObjectUUIDMap UUID"
            );
        }
        let old_uuid = comment_storage_uuid(before, old_identifier)?;
        let new_uuid = comment_storage_uuid(after, new_identifier)?;
        if let Some(old_uuid) = old_uuid {
            let new_uuid = new_uuid
                .ok_or_else(|| io::Error::other("shared COW clone lost its payload UUID"))?;
            assert_ne!(
                new_uuid, old_uuid,
                "shared COW clone reused the source payload UUID"
            );
        }
    } else if !old_registrations.is_empty() {
        assert!(
            !after_registrations.contains_key(&old_identifier),
            "removed comment object left a dangling ObjectUUIDMap entry"
        );
        let new_registrations = after_registrations
            .get(&new_identifier)
            .ok_or_else(|| io::Error::other("replacement lost its ObjectUUIDMap entry"))?;
        assert_eq!(new_registrations.len(), 1);
        assert_eq!(
            new_registrations[0].component_identifier, old_registrations[0].component_identifier,
            "replacement moved to a different component namespace"
        );
        assert_eq!(
            new_registrations[0].uuid, old_registrations[0].uuid,
            "replacement changed the registered UUID of a removed source object"
        );
        assert_eq!(
            comment_storage_uuid(after, new_identifier)?,
            comment_storage_uuid(before, old_identifier)?,
            "replacement changed the payload UUID of a removed source object"
        );
    } else {
        assert!(
            !after_registrations.contains_key(&new_identifier),
            "replacement invented an ObjectUUIDMap entry for an unregistered source object"
        );
    }
    Ok(())
}

/// The deprecated generic editor predates the focused UUID-map ownership
/// contract.  Keep its observed regression explicit and narrowly scoped: on
/// a registered reply edit it leaves the source object as an unreachable
/// orphan and omits the clone's ObjectUUIDMap registration.  The focused
/// backend continues to use `assert_cow_uuid_registration` above, so this
/// helper cannot mask a focused UUID regression.
fn assert_legacy_generic_reply_cow_uuid_registration(
    before: &[u8],
    after: &[u8],
    old_identifier: u64,
    new_identifier: u64,
) -> TestResult<()> {
    let before_registrations = uuid_registrations(before)?;
    let after_registrations = uuid_registrations(after)?;
    let old_registrations = before_registrations.get(&old_identifier).ok_or_else(|| {
        io::Error::other("legacy UUID regression fixture lost source registration")
    })?;
    assert_eq!(
        old_registrations.len(),
        1,
        "legacy UUID regression fixture has duplicate source registrations"
    );
    assert_eq!(
        after_registrations.get(&old_identifier),
        Some(old_registrations),
        "generic legacy path changed the source ObjectUUIDMap entry"
    );
    assert!(
        native_object_ids(after)?.contains(&old_identifier),
        "generic legacy path no longer exhibits the retained source orphan"
    );
    assert!(
        native_object_ids(after)?.contains(&new_identifier),
        "generic legacy path did not create the replacement reply object"
    );
    assert!(
        !after_registrations.contains_key(&new_identifier),
        "generic legacy path unexpectedly registered its replacement reply"
    );
    assert_eq!(
        comment_storage_uuid(after, old_identifier)?,
        comment_storage_uuid(before, old_identifier)?,
        "generic legacy orphan is not the original registered reply"
    );
    Ok(())
}

fn assert_uuid_registration_removed_if_dead(
    before: &[u8],
    after: &[u8],
    identifier: u64,
) -> TestResult<()> {
    if native_object_ids(after)?.contains(&identifier) {
        return Ok(());
    }
    let before_registrations = uuid_registrations(before)?;
    if before_registrations.contains_key(&identifier) {
        assert!(
            !uuid_registrations(after)?.contains_key(&identifier),
            "removed comment object left a dangling ObjectUUIDMap entry"
        );
    }
    Ok(())
}

fn metadata_invariants(
    source: &[u8],
) -> TestResult<(BTreeMap<u64, ComponentMetadataInvariant>, u64)> {
    let metadata = metadata_snapshot(source)?;
    let object_types = native_object_types(source)?;
    let comment_ids = object_types
        .iter()
        .filter_map(|(identifier, types)| is_comment_object(types).then_some(*identifier))
        .collect::<BTreeSet<_>>();
    let mut components = BTreeMap::new();
    for (identifier, component) in metadata.components {
        let external_references = component
            .external_references
            .into_iter()
            .filter(|(_, object_identifier, _)| {
                object_identifier
                    .is_none_or(|object_identifier| !comment_ids.contains(&object_identifier))
            })
            .collect();
        let versioned_external_references = component
            .versioned_external_references
            .into_iter()
            .filter(|(_, object_identifier, _)| {
                object_identifier
                    .is_none_or(|object_identifier| !comment_ids.contains(&object_identifier))
            })
            .collect();
        let object_uuid_map_entries = component
            .object_uuid_map_entries
            .into_iter()
            .filter(|entry| {
                tsp::ObjectUuidMapEntry::decode(entry.as_slice())
                    .map(|entry| !comment_ids.contains(&entry.identifier))
                    .unwrap_or(false)
            })
            .collect();
        components.insert(
            identifier,
            ComponentMetadataInvariant {
                locator: component.locator,
                external_references,
                versioned_external_references,
                object_uuid_map_entries,
            },
        );
    }
    Ok((components, metadata.last_object_identifier))
}

fn assert_metadata_watermark(source: &[u8]) -> TestResult<()> {
    let (_, watermark) = metadata_invariants(source)?;
    let largest_object = native_object_ids(source)?.into_iter().max().unwrap_or(0);
    assert!(
        watermark >= largest_object,
        "PackageMetadata watermark {watermark} is below native object {largest_object}"
    );
    Ok(())
}

fn archive_snapshots(source: &[u8]) -> TestResult<BTreeMap<String, Vec<u8>>> {
    Ok(fixtures::component_archives(source)?
        .into_iter()
        .map(|(name, archive)| Ok((name, archive.to_bytes()?)))
        .collect::<TestResult<BTreeMap<_, _>>>()?)
}

fn native_object_ids(source: &[u8]) -> TestResult<BTreeSet<u64>> {
    Ok(fixtures::component_archives(source)?
        .into_iter()
        .flat_map(|(_, archive)| archive.objects.into_iter())
        .filter_map(|object| object.archive_info.identifier)
        .collect())
}

fn native_object_types(source: &[u8]) -> TestResult<BTreeMap<u64, BTreeSet<u32>>> {
    let mut objects = BTreeMap::new();
    for (_, archive) in fixtures::component_archives(source)? {
        for object in archive.objects {
            let identifier = object
                .archive_info
                .identifier
                .ok_or_else(|| io::Error::other("archive object has no identifier"))?;
            let types = object
                .messages
                .into_iter()
                .map(|message| message.type_)
                .collect::<BTreeSet<_>>();
            if objects.insert(identifier, types).is_some() {
                return Err(io::Error::other(format!(
                    "duplicate native object identifier {identifier}"
                ))
                .into());
            }
        }
    }
    Ok(objects)
}

fn is_comment_object(types: &BTreeSet<u32>) -> bool {
    types
        .iter()
        .any(|type_| COMMENT_OBJECT_MESSAGE_TYPES.contains(type_))
}

fn referenced_object_ids(source: &[u8]) -> TestResult<BTreeSet<u64>> {
    let mut identifiers = BTreeSet::new();
    for (_, archive) in fixtures::component_archives(source)? {
        for object in archive.objects {
            for message in object.archive_info.message_infos {
                identifiers.extend(message.object_references);
                for field in message.field_infos {
                    identifiers.extend(field.object_references);
                }
            }
        }
    }
    Ok(identifiers)
}

fn referenced_comment_type_counts(
    object_types: &BTreeMap<u64, BTreeSet<u32>>,
    referenced: &BTreeSet<u64>,
) -> BTreeMap<u32, usize> {
    let mut counts = BTreeMap::new();
    for identifier in referenced {
        let Some(types) = object_types.get(identifier) else {
            continue;
        };
        for type_ in types {
            if COMMENT_OBJECT_MESSAGE_TYPES.contains(type_) {
                *counts.entry(*type_).or_default() += 1;
            }
        }
    }
    counts
}

fn comment_type_counts(object_types: &BTreeMap<u64, BTreeSet<u32>>) -> BTreeMap<u32, usize> {
    let mut counts = BTreeMap::new();
    for types in object_types.values() {
        for type_ in types {
            if COMMENT_OBJECT_MESSAGE_TYPES.contains(type_) {
                *counts.entry(*type_).or_default() += 1;
            }
        }
    }
    counts
}

fn unreachable_comment_type_counts(
    object_types: &BTreeMap<u64, BTreeSet<u32>>,
    referenced: &BTreeSet<u64>,
) -> BTreeMap<u32, usize> {
    let mut counts = BTreeMap::new();
    for (identifier, types) in object_types {
        if referenced.contains(identifier) {
            continue;
        }
        for type_ in types {
            if COMMENT_OBJECT_MESSAGE_TYPES.contains(type_) {
                *counts.entry(*type_).or_default() += 1;
            }
        }
    }
    counts
}

fn comment_type_delta(
    before: &BTreeMap<u64, BTreeSet<u32>>,
    after: &BTreeMap<u64, BTreeSet<u32>>,
) -> BTreeMap<u32, isize> {
    let before = comment_type_counts(before);
    let after = comment_type_counts(after);
    COMMENT_OBJECT_MESSAGE_TYPES
        .iter()
        .filter_map(|type_| {
            let delta = after.get(type_).copied().unwrap_or_default() as isize
                - before.get(type_).copied().unwrap_or_default() as isize;
            (delta != 0).then_some((*type_, delta))
        })
        .collect()
}

fn count_delta(
    before: &BTreeMap<u32, usize>,
    after: &BTreeMap<u32, usize>,
) -> BTreeMap<u32, isize> {
    COMMENT_OBJECT_MESSAGE_TYPES
        .iter()
        .filter_map(|type_| {
            let delta = after.get(type_).copied().unwrap_or_default() as isize
                - before.get(type_).copied().unwrap_or_default() as isize;
            (delta != 0).then_some((*type_, delta))
        })
        .collect()
}

fn assert_object_census_parity(
    focused_before: &[u8],
    generic_before: &[u8],
    focused_after: &[u8],
    generic_after: &[u8],
) -> TestResult<()> {
    assert_object_census_parity_with_legacy_author(
        focused_before,
        generic_before,
        focused_after,
        generic_after,
        None,
    )
}

fn assert_object_census_parity_with_legacy_author(
    focused_before: &[u8],
    generic_before: &[u8],
    focused_after: &[u8],
    generic_after: &[u8],
    legacy_generic_author: Option<u64>,
) -> TestResult<()> {
    let focused_before_types = native_object_types(focused_before)?;
    let generic_before_types = native_object_types(generic_before)?;
    let focused_after_types = native_object_types(focused_after)?;
    let generic_after_types = native_object_types(generic_after)?;

    if let Some(author_identifier) = legacy_generic_author {
        assert_legacy_generic_author_retention(
            focused_before,
            generic_before,
            focused_after,
            generic_after,
            author_identifier,
        )?;
    }

    let focused_before_non_comment = focused_before_types
        .iter()
        .filter_map(|(identifier, types)| (!is_comment_object(types)).then_some(*identifier))
        .collect::<BTreeSet<_>>();
    let generic_before_non_comment = generic_before_types
        .iter()
        .filter_map(|(identifier, types)| (!is_comment_object(types)).then_some(*identifier))
        .collect::<BTreeSet<_>>();
    let focused_after_non_comment = focused_after_types
        .iter()
        .filter_map(|(identifier, types)| (!is_comment_object(types)).then_some(*identifier))
        .collect::<BTreeSet<_>>();
    let generic_after_non_comment = generic_after_types
        .iter()
        .filter_map(|(identifier, types)| (!is_comment_object(types)).then_some(*identifier))
        .collect::<BTreeSet<_>>();
    assert_eq!(
        focused_before_non_comment, focused_after_non_comment,
        "focused operation changed non-comment native object census"
    );
    assert_eq!(
        generic_before_non_comment, generic_after_non_comment,
        "generic operation changed non-comment native object census"
    );
    assert_eq!(
        focused_after_non_comment, generic_after_non_comment,
        "focused and generic non-comment native object census differs"
    );
    let focused_before_referenced = referenced_object_ids(focused_before)?;
    let generic_before_referenced = referenced_object_ids(generic_before)?;
    let focused_referenced = referenced_object_ids(focused_after)?;
    let generic_referenced = referenced_object_ids(generic_after)?;
    let focused_before_reachable_counts =
        referenced_comment_type_counts(&focused_before_types, &focused_before_referenced);
    let generic_before_reachable_counts =
        referenced_comment_type_counts(&generic_before_types, &generic_before_referenced);
    let focused_after_reachable_counts =
        referenced_comment_type_counts(&focused_after_types, &focused_referenced);
    let generic_after_reachable_counts =
        referenced_comment_type_counts(&generic_after_types, &generic_referenced);
    let mut generic_before_reachable_counts_for_comparison =
        generic_before_reachable_counts.clone();
    let mut generic_after_reachable_counts_for_comparison = generic_after_reachable_counts.clone();
    if let Some(author_identifier) = legacy_generic_author {
        if !annotation_author_ids(focused_before)?.contains(&author_identifier) {
            let count = generic_before_reachable_counts_for_comparison
                .get_mut(&212)
                .ok_or_else(|| {
                    io::Error::other("legacy author retention has no source type-212 object")
                })?;
            if *count == 0 {
                return Err(io::Error::other("legacy source author count underflow").into());
            }
            *count -= 1;
            if *count == 0 {
                generic_before_reachable_counts_for_comparison.remove(&212);
            }
        }
        let count = generic_after_reachable_counts_for_comparison
            .get_mut(&212)
            .ok_or_else(|| io::Error::other("legacy author retention has no type-212 object"))?;
        if *count == 0 {
            return Err(io::Error::other("legacy author retention count underflow").into());
        }
        *count -= 1;
        if *count == 0 {
            generic_after_reachable_counts_for_comparison.remove(&212);
        }
    }
    assert_eq!(
        focused_before_reachable_counts, generic_before_reachable_counts_for_comparison,
        "focused and generic source reachable comment type counts differ"
    );
    assert_eq!(
        focused_after_reachable_counts, generic_after_reachable_counts_for_comparison,
        "focused and generic reachable comment type counts differ"
    );
    assert_eq!(
        count_delta(
            &focused_before_reachable_counts,
            &focused_after_reachable_counts,
        ),
        count_delta(
            &generic_before_reachable_counts_for_comparison,
            &generic_after_reachable_counts_for_comparison,
        ),
        "focused and generic reachable comment creation/removal counts differ"
    );
    let focused_counts = referenced_comment_type_counts(&focused_after_types, &focused_referenced);
    let generic_counts = generic_after_reachable_counts_for_comparison.clone();
    assert_eq!(
        focused_counts, generic_counts,
        "focused and generic referenced comment object type counts differ"
    );

    let focused_total_counts = comment_type_counts(&focused_after_types);
    let generic_total_counts = comment_type_counts(&generic_after_types);
    if focused_total_counts != generic_total_counts {
        let focused_unreachable =
            unreachable_comment_type_counts(&focused_after_types, &focused_referenced);
        let generic_unreachable =
            unreachable_comment_type_counts(&generic_after_types, &generic_referenced);
        assert_eq!(
            focused_after_reachable_counts, generic_after_reachable_counts_for_comparison,
            "focused and generic comment census differs in reachable objects"
        );
        eprintln!(
            "focused/generic comment census differs only in unreachable objects: focused_total={focused_total_counts:?}, generic_total={generic_total_counts:?}, focused_delta={:?}, generic_delta={:?}, focused_unreachable={focused_unreachable:?}, generic_unreachable={generic_unreachable:?}; retained as a documented legacy orphan difference",
            comment_type_delta(&focused_before_types, &focused_after_types),
            comment_type_delta(&generic_before_types, &generic_after_types),
        );
    }

    let focused_comment_ids = focused_after_types
        .iter()
        .filter_map(|(identifier, types)| is_comment_object(types).then_some(*identifier))
        .collect::<BTreeSet<_>>();
    let generic_comment_ids = generic_after_types
        .iter()
        .filter_map(|(identifier, types)| is_comment_object(types).then_some(*identifier))
        .collect::<BTreeSet<_>>();
    if focused_comment_ids != generic_comment_ids {
        let focused_only = focused_comment_ids
            .difference(&generic_comment_ids)
            .copied()
            .collect::<Vec<_>>();
        let generic_only = generic_comment_ids
            .difference(&focused_comment_ids)
            .copied()
            .collect::<Vec<_>>();
        eprintln!(
            "focused/generic comment census uses distinct storage IDs; retained referenced type counts match, focused-only={focused_only:?}, generic-only={generic_only:?}; this is safe COW/GC identity variance"
        );
    }
    Ok(())
}

fn assert_archive_references_are_live(source: &[u8]) -> TestResult<()> {
    let identifiers = native_object_ids(source)?;
    for (_, archive) in fixtures::component_archives(source)? {
        for object in archive.objects {
            let object_identifier = object
                .archive_info
                .identifier
                .ok_or_else(|| io::Error::other("archive object has no identifier"))?;
            for message in object.archive_info.message_infos {
                for reference in message.object_references {
                    assert!(
                        identifiers.contains(&reference),
                        "object {object_identifier} has dangling object reference {reference}"
                    );
                }
                for field in message.field_infos {
                    for reference in field.object_references {
                        assert!(
                            identifiers.contains(&reference),
                            "object {object_identifier} has dangling field reference {reference}"
                        );
                    }
                }
            }
        }
    }
    Ok(())
}

fn focused_thread(package: &Package, target: &DrawableTarget) -> TestResult<ThreadSnapshot> {
    let slide = SlideSelector::index(0);
    let drawable = DrawableSelector::index(target.position);
    let root = package.slide_drawable_comment(slide, drawable)?;
    let replies = package.slide_drawable_comment_replies(slide, drawable)?;
    Ok(ThreadSnapshot {
        root: root.map(|comment| {
            (
                comment.text().to_owned(),
                comment
                    .timestamp()
                    .map(|timestamp| timestamp.as_f64().to_bits()),
                comment.author().is_some(),
            )
        }),
        replies: replies
            .iter()
            .map(|reply| {
                (
                    reply.text().to_owned(),
                    reply
                        .timestamp()
                        .map(|timestamp| timestamp.as_f64().to_bits()),
                    reply.author().is_some(),
                )
            })
            .collect(),
    })
}

fn generic_thread(
    editor: &IWorkDrawableCommentEditor,
    target: &DrawableTarget,
) -> TestResult<ThreadSnapshot> {
    let drawable = DrawableId::from_raw(target.identifier)?;
    let root = editor.comment(drawable)?;
    let replies = editor.replies(drawable)?;
    Ok(ThreadSnapshot {
        root: root.map(|comment| {
            let comment = comment.comment;
            (
                comment.text,
                comment.creation_date_seconds.map(f64::to_bits),
                comment.author_id.is_some(),
            )
        }),
        replies: replies
            .into_iter()
            .map(|reply| {
                let comment = reply.comment;
                (
                    comment.text,
                    comment.creation_date_seconds.map(f64::to_bits),
                    comment.author_id.is_some(),
                )
            })
            .collect(),
    })
}

fn assert_metadata_parity(
    focused_before: &[u8],
    generic_before: &[u8],
    focused_after: &[u8],
    generic_after: &[u8],
) -> TestResult<()> {
    let (focused_before_invariants, _) = metadata_invariants(focused_before)?;
    let (generic_before_invariants, _) = metadata_invariants(generic_before)?;
    let (focused_after_invariants, _) = metadata_invariants(focused_after)?;
    let (generic_after_invariants, _) = metadata_invariants(generic_after)?;
    assert_eq!(
        focused_before_invariants, focused_after_invariants,
        "focused operation changed non-comment PackageMetadata ownership"
    );
    assert_eq!(
        generic_before_invariants, generic_after_invariants,
        "generic operation changed non-comment PackageMetadata ownership"
    );
    assert_eq!(
        focused_after_invariants, generic_after_invariants,
        "focused and generic non-comment PackageMetadata projections differ"
    );
    assert_metadata_watermark(focused_after)?;
    assert_metadata_watermark(generic_after)?;
    Ok(())
}

fn assert_backend_physical_parity(
    focused_before: &[u8],
    generic_before: &[u8],
    focused_after: &[u8],
    generic_after: &[u8],
) -> TestResult<()> {
    assert_object_census_parity(focused_before, generic_before, focused_after, generic_after)?;
    assert_metadata_parity(focused_before, generic_before, focused_after, generic_after)?;
    assert_metadata_references_are_live(focused_after)?;
    assert_metadata_references_are_live(generic_after)?;
    assert_archive_references_are_live(focused_after)?;
    assert_archive_references_are_live(generic_after)?;
    Ok(())
}

fn assert_backend_physical_parity_with_legacy_author(
    focused_before: &[u8],
    generic_before: &[u8],
    focused_after: &[u8],
    generic_after: &[u8],
    author_identifier: u64,
) -> TestResult<()> {
    assert_object_census_parity_with_legacy_author(
        focused_before,
        generic_before,
        focused_after,
        generic_after,
        Some(author_identifier),
    )?;
    assert_metadata_parity(focused_before, generic_before, focused_after, generic_after)?;
    assert_metadata_references_are_live(focused_after)?;
    assert_metadata_references_are_live(generic_after)?;
    assert_archive_references_are_live(focused_after)?;
    assert_archive_references_are_live(generic_after)?;
    Ok(())
}

fn assert_backend_parity(
    focused_before: &[u8],
    generic_before: &[u8],
    focused_bytes: &[u8],
    generic_bytes: &[u8],
    targets: &[&DrawableTarget],
) -> TestResult<()> {
    let focused_package = Package::from_bytes(focused_bytes)?;
    let generic_editor = IWorkDrawableCommentEditor::from_bytes(generic_bytes)?;
    for target in targets {
        assert_eq!(
            focused_thread(&focused_package, target)?,
            generic_thread(&generic_editor, target)?,
            "focused and generic comment projections differ for drawable {}",
            target.identifier
        );
    }
    assert_backend_physical_parity(focused_before, generic_before, focused_bytes, generic_bytes)
}

fn assert_backend_parity_with_legacy_author(
    focused_before: &[u8],
    generic_before: &[u8],
    focused_bytes: &[u8],
    generic_bytes: &[u8],
    target: &DrawableTarget,
    author_identifier: u64,
) -> TestResult<()> {
    let focused_package = Package::from_bytes(focused_bytes)?;
    let generic_editor = IWorkDrawableCommentEditor::from_bytes(generic_bytes)?;
    assert_eq!(
        focused_thread(&focused_package, target)?,
        generic_thread(&generic_editor, target)?,
        "focused and generic comment projections differ for drawable {}",
        target.identifier
    );
    assert_selected_thread_author_metadata_equal(focused_bytes, generic_bytes, target)?;
    assert_backend_physical_parity_with_legacy_author(
        focused_before,
        generic_before,
        focused_bytes,
        generic_bytes,
        author_identifier,
    )
}

fn assert_valid_generated_timestamp(bits: Option<u64>, label: &str) -> TestResult<()> {
    let bits = bits.ok_or_else(|| io::Error::other(format!("{label} has no timestamp")))?;
    let timestamp = f64::from_bits(bits);
    assert!(
        timestamp.is_finite() && timestamp > 0.0,
        "{label} has an invalid generated timestamp {timestamp}"
    );
    Ok(())
}

fn assert_backend_parity_with_generated_root(
    focused_before: &[u8],
    generic_before: &[u8],
    focused_bytes: &[u8],
    generic_bytes: &[u8],
    target: &DrawableTarget,
) -> TestResult<()> {
    let focused = focused_thread(&Package::from_bytes(focused_bytes)?, target)?;
    let generic = generic_thread(
        &IWorkDrawableCommentEditor::from_bytes(generic_bytes)?,
        target,
    )?;
    let focused_root = focused
        .root
        .as_ref()
        .ok_or_else(|| io::Error::other("focused writer produced no root"))?;
    let generic_root = generic
        .root
        .as_ref()
        .ok_or_else(|| io::Error::other("generic writer produced no root"))?;
    assert_eq!(focused_root.0, generic_root.0);
    assert_eq!(focused_root.2, generic_root.2);
    assert_valid_generated_timestamp(focused_root.1, "focused root")?;
    assert_valid_generated_timestamp(generic_root.1, "generic root")?;
    assert_eq!(focused.replies, generic.replies);
    assert_backend_physical_parity(focused_before, generic_before, focused_bytes, generic_bytes)
}

fn generic_set(source: &[u8], target: &DrawableTarget, text: &str) -> TestResult<Vec<u8>> {
    let mut editor = IWorkDrawableCommentEditor::from_bytes(source)?;
    editor.set_comment(DrawableId::from_raw(target.identifier)?, text)?;
    Ok(editor.to_bytes()?)
}

fn focused_set(source: &[u8], target: &DrawableTarget, text: &str) -> TestResult<Vec<u8>> {
    let package = Package::from_bytes(source)?;
    let commit = package
        .edit_slide_drawable_comment(
            SlideSelector::index(0),
            DrawableSelector::index(target.position),
        )?
        .set(text)?
        .commit()?;
    exact_bytes(commit.package())
}

fn generic_clear(source: &[u8], target: &DrawableTarget) -> TestResult<Vec<u8>> {
    let mut editor = IWorkDrawableCommentEditor::from_bytes(source)?;
    editor.clear_comment(DrawableId::from_raw(target.identifier)?)?;
    Ok(editor.to_bytes()?)
}

fn focused_clear(source: &[u8], target: &DrawableTarget) -> TestResult<Vec<u8>> {
    let package = Package::from_bytes(source)?;
    let commit = package
        .edit_slide_drawable_comment(
            SlideSelector::index(0),
            DrawableSelector::index(target.position),
        )?
        .clear()?
        .commit()?;
    exact_bytes(commit.package())
}

fn generic_set_reply(
    source: &[u8],
    target: &DrawableTarget,
    reply_identifier: u64,
    text: &str,
) -> TestResult<Vec<u8>> {
    let mut editor = with_context(
        "generic_set_reply: parse",
        IWorkDrawableCommentEditor::from_bytes(source),
    )?;
    with_context(
        "generic_set_reply: edit",
        editor.set_reply(
            DrawableId::from_raw(target.identifier)?,
            StorageId::from_raw(reply_identifier)?,
            text,
        ),
    )?;
    with_context("generic_set_reply: serialize", editor.to_bytes())
}

fn focused_set_reply(source: &[u8], target: &DrawableTarget, text: &str) -> TestResult<Vec<u8>> {
    let package = with_context("focused_set_reply: parse", Package::from_bytes(source))?;
    let editor = with_context(
        "focused_set_reply: select",
        package.edit_slide_drawable_comment(
            SlideSelector::index(0),
            DrawableSelector::index(target.position),
        ),
    )?;
    let editor = with_context(
        "focused_set_reply: set_reply",
        editor.set_reply(ReplySelector::index(0), text),
    )?;
    let commit = with_context("focused_set_reply: commit", editor.commit())?;
    with_context(
        "focused_set_reply: serialize",
        exact_bytes(commit.package()),
    )
}

fn generic_remove_reply(
    source: &[u8],
    target: &DrawableTarget,
    reply_identifier: u64,
) -> TestResult<Vec<u8>> {
    let mut editor = IWorkDrawableCommentEditor::from_bytes(source)?;
    editor.remove_reply(
        DrawableId::from_raw(target.identifier)?,
        StorageId::from_raw(reply_identifier)?,
    )?;
    Ok(editor.to_bytes()?)
}

fn focused_remove_reply(source: &[u8], target: &DrawableTarget) -> TestResult<Vec<u8>> {
    let package = Package::from_bytes(source)?;
    let commit = package
        .edit_slide_drawable_comment(
            SlideSelector::index(0),
            DrawableSelector::index(target.position),
        )?
        .remove_reply(ReplySelector::index(0))?
        .commit()?;
    exact_bytes(commit.package())
}

fn normalized_component_locator(locator: &str) -> &str {
    let locator = locator.strip_prefix("Index/").unwrap_or(locator);
    locator.strip_suffix(".iwa").unwrap_or(locator)
}

fn metadata_component_for_archive<'metadata>(
    metadata: &'metadata MetadataSnapshot,
    archive_name: &str,
) -> TestResult<Option<&'metadata ComponentMetadataSnapshot>> {
    let locator = normalized_component_locator(archive_name);
    let mut matches = metadata
        .components
        .values()
        .filter(|component| normalized_component_locator(&component.locator) == locator);
    let Some(component) = matches.next() else {
        return Ok(None);
    };
    if matches.next().is_some() {
        return Err(io::Error::other(format!(
            "PackageMetadata has duplicate current locator {locator}"
        ))
        .into());
    }
    Ok(Some(component))
}

fn metadata_component_id_for_archive(
    metadata: &MetadataSnapshot,
    archive_name: &str,
) -> TestResult<Option<u64>> {
    let locator = normalized_component_locator(archive_name);
    let mut matches = metadata
        .components
        .iter()
        .filter(|(_, component)| normalized_component_locator(&component.locator) == locator);
    let Some((identifier, _)) = matches.next() else {
        return Ok(None);
    };
    if matches.next().is_some() {
        return Err(io::Error::other(format!(
            "PackageMetadata has duplicate current locator {locator}"
        ))
        .into());
    }
    Ok(Some(*identifier))
}

fn assert_metadata_references_are_live(source: &[u8]) -> TestResult<()> {
    let metadata = metadata_snapshot(source)?;
    let current_components = metadata.components.keys().copied().collect::<BTreeSet<_>>();
    let mut object_components = BTreeMap::new();
    for (name, archive) in fixtures::component_archives(source)? {
        let Some(component_identifier) = metadata_component_id_for_archive(&metadata, &name)?
        else {
            continue;
        };
        for object in archive.objects {
            let identifier = object
                .archive_info
                .identifier
                .ok_or_else(|| io::Error::other("archive object has no identifier"))?;
            if object_components
                .insert(identifier, component_identifier)
                .is_some()
            {
                return Err(io::Error::other(format!(
                    "duplicate metadata component ownership for object {identifier}"
                ))
                .into());
            }
        }
    }
    for component in metadata.components.values() {
        for references in [
            &component.external_references,
            &component.versioned_external_references,
        ] {
            for reference in references {
                assert!(
                    current_components.contains(&reference.0),
                    "metadata references missing current component {}",
                    reference.0
                );
                if let Some(object_identifier) = reference.1 {
                    let actual_component =
                        object_components.get(&object_identifier).ok_or_else(|| {
                            io::Error::other(format!(
                                "metadata references missing object {object_identifier}"
                            ))
                        })?;
                    assert_eq!(
                        reference.0, *actual_component,
                        "metadata object edge points at the wrong current component"
                    );
                }
            }
        }
    }
    Ok(())
}

fn component_metadata_changed(
    before: &ComponentMetadataSnapshot,
    after: &ComponentMetadataSnapshot,
) -> bool {
    before.locator != after.locator
        || before.external_references != after.external_references
        || before.versioned_external_references != after.versioned_external_references
        || before.object_uuid_map_entries != after.object_uuid_map_entries
}

#[derive(Clone, Copy)]
enum SaveTokenExpectation {
    FocusedExactlyOnce,
    LegacyMonotonic,
}

fn assert_save_tokens(
    before: &[u8],
    after: &[u8],
    expectation: SaveTokenExpectation,
) -> TestResult<()> {
    let before_archives = archive_snapshots(before)?;
    let after_archives = archive_snapshots(after)?;
    let before_metadata = metadata_snapshot(before)?;
    let after_metadata = metadata_snapshot(after)?;
    let metadata_changed =
        before_archives.get(METADATA_COMPONENT) != after_archives.get(METADATA_COMPONENT);
    if metadata_changed {
        let before_root = before_metadata
            .save_token
            .ok_or_else(|| io::Error::other("source metadata has no root save token"))?;
        let after_root = after_metadata
            .save_token
            .ok_or_else(|| io::Error::other("candidate metadata has no root save token"))?;
        let expected = before_root
            .checked_add(1)
            .ok_or_else(|| io::Error::other("PackageMetadata root save token overflowed"))?;
        match expectation {
            SaveTokenExpectation::FocusedExactlyOnce => assert_eq!(
                after_root, expected,
                "focused backend PackageMetadata root save token did not advance exactly once (before={before_root}, after={after_root})"
            ),
            SaveTokenExpectation::LegacyMonotonic => {
                assert!(
                    after_root >= before_root,
                    "legacy generic PackageMetadata root save token moved backwards"
                );
                if after_root != expected {
                    eprintln!(
                        "legacy generic backend advanced PackageMetadata root save token by {} instead of one",
                        after_root.saturating_sub(before_root)
                    );
                }
            },
        }
    }
    for (name, before_archive) in &before_archives {
        if name == METADATA_COMPONENT {
            continue;
        }
        let after_archive = after_archives
            .get(name)
            .ok_or_else(|| io::Error::other(format!("archive {name} disappeared")))?;
        let before_component = metadata_component_for_archive(&before_metadata, name)?;
        let after_component = metadata_component_for_archive(&after_metadata, name)?;
        let metadata_changed = match (before_component, after_component) {
            (Some(before), Some(after)) => component_metadata_changed(before, after),
            (None, None) => false,
            _ => true,
        };
        let archive_changed = before_archive != after_archive;
        if archive_changed || metadata_changed {
            let before_token = before_component.and_then(|component| component.save_token);
            let after_token = after_component.and_then(|component| component.save_token);
            let before_token = before_token.ok_or_else(|| {
                io::Error::other(format!(
                    "changed component {name} has no current metadata token"
                ))
            })?;
            let after_token = after_token.ok_or_else(|| {
                io::Error::other(format!("changed component {name} lost its metadata token"))
            })?;
            let package_before = before_metadata
                .save_token
                .ok_or_else(|| io::Error::other("source metadata has no root save token"))?;
            let expected = package_before
                .checked_add(1)
                .ok_or_else(|| io::Error::other("component save token overflowed"))?;
            match expectation {
                SaveTokenExpectation::FocusedExactlyOnce => assert_eq!(
                    after_token, expected,
                    "focused backend changed component {name} did not receive the next package token exactly once (component_before={before_token}, package_before={package_before}, after={after_token})"
                ),
                SaveTokenExpectation::LegacyMonotonic => {
                    assert!(
                        after_token >= before_token,
                        "legacy generic component {name} save token moved backwards"
                    );
                    if after_token != expected {
                        eprintln!(
                            "legacy generic backend changed component {name} to token {after_token}; package-root next token is {expected} (component_before={before_token})"
                        );
                    }
                },
            }
        } else if let (Some(before), Some(after)) = (before_component, after_component) {
            assert_eq!(
                after.save_token, before.save_token,
                "unchanged component {name} token changed"
            );
        }
    }
    for name in after_archives.keys() {
        assert!(
            before_archives.contains_key(name),
            "operation introduced an unexpected archive component {name}"
        );
    }
    Ok(())
}

fn assert_save_tokens_advance_once(before: &[u8], after: &[u8]) -> TestResult<()> {
    assert_save_tokens(before, after, SaveTokenExpectation::FocusedExactlyOnce)
}

fn assert_legacy_save_tokens_monotonic(before: &[u8], after: &[u8]) -> TestResult<()> {
    assert_save_tokens(before, after, SaveTokenExpectation::LegacyMonotonic)
}

fn assert_generated_comment_cleanup(before: &[u8], after: &[u8]) -> TestResult<()> {
    let before_metadata = metadata_snapshot(before)?;
    let after_metadata = metadata_snapshot(after)?;
    assert!(
        after_metadata.last_object_identifier >= before_metadata.last_object_identifier,
        "cleanup moved the PackageMetadata watermark backwards"
    );
    let before_components = before_metadata
        .components
        .into_iter()
        .map(|(identifier, component)| {
            (
                identifier,
                (
                    component.locator,
                    component.external_references,
                    component.versioned_external_references,
                    component.object_uuid_map_entries,
                ),
            )
        })
        .collect::<BTreeMap<_, _>>();
    let after_components = after_metadata
        .components
        .into_iter()
        .map(|(identifier, component)| {
            (
                identifier,
                (
                    component.locator,
                    component.external_references,
                    component.versioned_external_references,
                    component.object_uuid_map_entries,
                ),
            )
        })
        .collect::<BTreeMap<_, _>>();
    assert_eq!(
        after_components, before_components,
        "cleanup changed PackageMetadata ownership or UUID membership"
    );
    assert_eq!(
        native_object_ids(after)?,
        native_object_ids(before)?,
        "cleanup left a generated comment or author object in the native census"
    );
    assert_metadata_watermark(after)?;
    assert_archive_references_are_live(after)?;
    Ok(())
}

fn assert_foreign_component(source: &[u8], identifier: u64, expected: &str) -> TestResult<()> {
    let actual = component_for_object(source, identifier)?;
    assert_eq!(actual, expected);
    Ok(())
}

fn component_for_object(source: &[u8], identifier: u64) -> TestResult<String> {
    fixtures::component_archives(source)?
        .into_iter()
        .find_map(|(name, archive)| archive.object(identifier).is_some().then_some(name))
        .ok_or_else(|| io::Error::other(format!("object {identifier} disappeared")).into())
}

fn with_context<T, E>(label: &str, result: Result<T, E>) -> TestResult<T>
where
    E: std::fmt::Display,
{
    result.map_err(|error| io::Error::other(format!("{label}: {error}")).into())
}

fn assert_root_component(
    source: &[u8],
    target: &DrawableTarget,
    expected: &str,
) -> TestResult<u64> {
    let comment = IWorkDrawableCommentEditor::from_bytes(source)?
        .comment(DrawableId::from_raw(target.identifier)?)?
        .ok_or_else(|| io::Error::other("drawable has no root comment"))?;
    let identifier = comment.storage_id.get();
    assert_foreign_component(source, identifier, expected)?;
    Ok(identifier)
}

fn assert_reply_component(
    source: &[u8],
    target: &DrawableTarget,
    position: usize,
    expected: &str,
) -> TestResult<u64> {
    let reply = IWorkDrawableCommentEditor::from_bytes(source)?
        .replies(DrawableId::from_raw(target.identifier)?)?
        .into_iter()
        .nth(position)
        .ok_or_else(|| io::Error::other("drawable has no selected reply"))?;
    let identifier = reply.storage_id.get();
    assert_foreign_component(source, identifier, expected)?;
    Ok(identifier)
}

fn assert_shared_cow_parity(
    fixture: &RelocatedFixture,
    focused_before: &[u8],
    generic_before: &[u8],
    focused_bytes: &[u8],
    generic_bytes: &[u8],
    expected_selected: Option<&str>,
    expected_sibling: Option<&str>,
) -> TestResult<()> {
    let sibling = fixture
        .sibling
        .as_ref()
        .ok_or_else(|| io::Error::other("shared fixture has no sibling"))?;
    assert_backend_parity(
        focused_before,
        generic_before,
        focused_bytes,
        generic_bytes,
        &[&fixture.target, sibling],
    )?;
    let focused_oracle = IWorkDrawableCommentEditor::from_bytes(focused_bytes)?;
    let generic_oracle = IWorkDrawableCommentEditor::from_bytes(generic_bytes)?;
    let selected = DrawableId::from_raw(fixture.target.identifier)?;
    let sibling_id = DrawableId::from_raw(sibling.identifier)?;
    let focused_selected = focused_oracle.comment(selected)?;
    let focused_sibling = focused_oracle.comment(sibling_id)?;
    let generic_selected = generic_oracle.comment(selected)?;
    let generic_sibling = generic_oracle.comment(sibling_id)?;
    assert_eq!(
        focused_selected
            .as_ref()
            .map(|comment| comment.comment.text.as_str()),
        expected_selected
    );
    assert_eq!(
        focused_sibling
            .as_ref()
            .map(|comment| comment.comment.text.as_str()),
        expected_sibling
    );
    assert_eq!(
        generic_selected
            .as_ref()
            .map(|comment| comment.comment.text.as_str()),
        expected_selected
    );
    assert_eq!(
        generic_sibling
            .as_ref()
            .map(|comment| comment.comment.text.as_str()),
        expected_sibling
    );
    if let (Some(selected), Some(sibling)) = (focused_selected, focused_sibling) {
        assert_ne!(selected.storage_id, sibling.storage_id);
    }
    Ok(())
}

#[test]
fn foreign_root_read_and_set_match_generic_backend() -> TestResult {
    let fixture = fixtures::root_foreign_fixture()?;
    let root = fixture
        .root_identifier
        .ok_or_else(|| io::Error::other("foreign-root fixture has no root"))?;
    fixtures::assert_metadata_relocated(fixtures::SOURCE, &fixture.bytes, root)?;
    assert_foreign_component(&fixture.bytes, root, fixtures::AUTHOR_STORAGE_COMPONENT)?;

    let focused = Package::from_bytes(&fixture.bytes)?;
    let generic = IWorkDrawableCommentEditor::from_bytes(&fixture.bytes)?;
    assert_eq!(
        focused_thread(&focused, &fixture.target)?,
        generic_thread(&generic, &fixture.target)?
    );

    let focused_bytes = focused_set(&fixture.bytes, &fixture.target, "focused foreign root")?;
    let generic_bytes = generic_set(&fixture.bytes, &fixture.target, "focused foreign root")?;
    assert_backend_parity(
        &fixture.bytes,
        &fixture.bytes,
        &focused_bytes,
        &generic_bytes,
        &[&fixture.target],
    )?;
    let focused_root = assert_root_component(
        &focused_bytes,
        &fixture.target,
        fixtures::AUTHOR_STORAGE_COMPONENT,
    )?;
    let generic_root = assert_root_component(
        &generic_bytes,
        &fixture.target,
        fixtures::AUTHOR_STORAGE_COMPONENT,
    )?;
    assert_cow_uuid_registration(&fixture.bytes, &focused_bytes, root, focused_root)?;
    assert_cow_uuid_registration(&fixture.bytes, &generic_bytes, root, generic_root)?;
    assert_foreign_component(&focused_bytes, root, fixtures::AUTHOR_STORAGE_COMPONENT)?;
    assert_foreign_component(&generic_bytes, root, fixtures::AUTHOR_STORAGE_COMPONENT)?;
    assert_save_tokens_advance_once(&fixture.bytes, &focused_bytes)?;
    assert_legacy_save_tokens_monotonic(&fixture.bytes, &generic_bytes)?;
    Ok(())
}

#[test]
fn foreign_reply_set_and_remove_match_generic_backend() -> TestResult {
    let fixture = with_context(
        "build reply_foreign_fixture",
        fixtures::reply_foreign_fixture(),
    )?;
    let reply = fixture
        .reply_identifier
        .ok_or_else(|| io::Error::other("foreign-reply fixture has no reply"))?;
    with_context(
        "validate reply relocation metadata",
        fixtures::assert_metadata_relocated(&fixture.metadata_source, &fixture.bytes, reply),
    )?;
    with_context(
        "validate registered reply UUID",
        assert_registered_uuid(&fixture.bytes, reply),
    )?;
    with_context(
        "validate relocated reply component",
        assert_foreign_component(&fixture.bytes, reply, fixtures::AUTHOR_STORAGE_COMPONENT),
    )?;
    let reply_author = with_context(
        "read relocated reply author",
        comment_author_identifier(&fixture.bytes, reply),
    )?
    .ok_or_else(|| io::Error::other("foreign-reply fixture reply has no author"))?;
    let source_root = fixture
        .root_identifier
        .ok_or_else(|| io::Error::other("foreign-reply fixture has no root"))?;
    let source_root_component = with_context(
        "locate source root component",
        component_for_object(&fixture.metadata_source, source_root),
    )?;

    let focused_bytes = with_context(
        "focused_set_reply",
        focused_set_reply(&fixture.bytes, &fixture.target, "focused foreign reply"),
    )?;
    let generic_bytes = with_context(
        "generic_set_reply",
        generic_set_reply(
            &fixture.bytes,
            &fixture.target,
            reply,
            "focused foreign reply",
        ),
    )?;
    with_context(
        "set-reply backend parity",
        assert_backend_parity(
            &fixture.bytes,
            &fixture.bytes,
            &focused_bytes,
            &generic_bytes,
            &[&fixture.target],
        ),
    )?;
    let focused_root =
        assert_root_component(&focused_bytes, &fixture.target, &source_root_component)?;
    let generic_root =
        assert_root_component(&generic_bytes, &fixture.target, &source_root_component)?;
    let focused_reply = assert_reply_component(
        &focused_bytes,
        &fixture.target,
        0,
        fixtures::AUTHOR_STORAGE_COMPONENT,
    )?;
    let generic_reply = assert_reply_component(
        &generic_bytes,
        &fixture.target,
        0,
        fixtures::AUTHOR_STORAGE_COMPONENT,
    )?;
    with_context(
        "focused set-reply root COW UUID",
        assert_cow_uuid_registration(&fixture.bytes, &focused_bytes, source_root, focused_root),
    )?;
    with_context(
        "generic set-reply root COW UUID",
        assert_cow_uuid_registration(&fixture.bytes, &generic_bytes, source_root, generic_root),
    )?;
    with_context(
        "focused set-reply registered-reply COW UUID",
        assert_cow_uuid_registration(&fixture.bytes, &focused_bytes, reply, focused_reply),
    )?;
    with_context(
        "generic set-reply registered-reply legacy UUID behavior",
        assert_legacy_generic_reply_cow_uuid_registration(
            &fixture.bytes,
            &generic_bytes,
            reply,
            generic_reply,
        ),
    )?;
    assert_save_tokens_advance_once(&fixture.bytes, &focused_bytes)?;
    assert_legacy_save_tokens_monotonic(&fixture.bytes, &generic_bytes)?;

    let focused_bytes = with_context(
        "focused_remove_reply",
        focused_remove_reply(&fixture.bytes, &fixture.target),
    )?;
    let generic_bytes = with_context(
        "generic_remove_reply",
        generic_remove_reply(&fixture.bytes, &fixture.target, reply),
    )?;
    with_context(
        "remove selected-thread author metadata",
        assert_selected_thread_author_metadata_present_and_equal(
            &focused_bytes,
            &generic_bytes,
            &fixture.target,
        ),
    )?;
    with_context(
        "remove-reply backend parity",
        assert_backend_parity_with_legacy_author(
            &fixture.bytes,
            &fixture.bytes,
            &focused_bytes,
            &generic_bytes,
            &fixture.target,
            reply_author,
        ),
    )?;
    with_context(
        "focused removed-reply UUID cleanup",
        assert_uuid_registration_removed_if_dead(&fixture.bytes, &focused_bytes, reply),
    )?;
    with_context(
        "generic removed-reply UUID cleanup",
        assert_uuid_registration_removed_if_dead(&fixture.bytes, &generic_bytes, reply),
    )?;
    with_context(
        "focused remove-reply save tokens",
        assert_save_tokens_advance_once(&fixture.bytes, &focused_bytes),
    )?;
    with_context(
        "generic remove-reply save tokens",
        assert_legacy_save_tokens_monotonic(&fixture.bytes, &generic_bytes),
    )?;

    let focused_clear_bytes = with_context(
        "focused_clear after reply removal",
        focused_clear(&focused_bytes, &fixture.target),
    )?;
    let generic_clear_bytes = with_context(
        "generic_clear after reply removal",
        generic_clear(&generic_bytes, &fixture.target),
    )?;
    with_context(
        "clear backend parity",
        assert_backend_parity_with_legacy_author(
            &focused_bytes,
            &generic_bytes,
            &focused_clear_bytes,
            &generic_clear_bytes,
            &fixture.target,
            reply_author,
        ),
    )?;
    with_context(
        "focused clear save tokens",
        assert_save_tokens_advance_once(&focused_bytes, &focused_clear_bytes),
    )?;
    with_context(
        "generic clear save tokens",
        assert_legacy_save_tokens_monotonic(&generic_bytes, &generic_clear_bytes),
    )?;
    Ok(())
}

#[test]
fn shared_foreign_root_cow_and_clear_match_generic_backend() -> TestResult {
    let fixture = fixtures::shared_foreign_root_fixture()?;
    let baseline = IWorkDrawableCommentEditor::from_bytes(&fixture.bytes)?;
    let sibling = fixture
        .sibling
        .as_ref()
        .ok_or_else(|| io::Error::other("shared-root fixture has no sibling"))?;
    let original_sibling_text = baseline
        .comment(DrawableId::from_raw(sibling.identifier)?)?
        .map(|comment| comment.comment.text);
    let focused_set_bytes = focused_set(&fixture.bytes, &fixture.target, "selected root")?;
    let generic_set_bytes = generic_set(&fixture.bytes, &fixture.target, "selected root")?;
    assert_shared_cow_parity(
        &fixture,
        &fixture.bytes,
        &fixture.bytes,
        &focused_set_bytes,
        &generic_set_bytes,
        Some("selected root"),
        original_sibling_text.as_deref(),
    )?;
    let focused_selected_root = assert_root_component(
        &focused_set_bytes,
        &fixture.target,
        fixtures::AUTHOR_STORAGE_COMPONENT,
    )?;
    let generic_selected_root = assert_root_component(
        &generic_set_bytes,
        &fixture.target,
        fixtures::AUTHOR_STORAGE_COMPONENT,
    )?;
    assert_root_component(
        &focused_set_bytes,
        sibling,
        fixtures::AUTHOR_STORAGE_COMPONENT,
    )?;
    assert_root_component(
        &generic_set_bytes,
        sibling,
        fixtures::AUTHOR_STORAGE_COMPONENT,
    )?;
    let root = fixture
        .root_identifier
        .ok_or_else(|| io::Error::other("shared-root fixture has no root"))?;
    assert_cow_uuid_registration(
        &fixture.bytes,
        &focused_set_bytes,
        root,
        focused_selected_root,
    )?;
    assert_cow_uuid_registration(
        &fixture.bytes,
        &generic_set_bytes,
        root,
        generic_selected_root,
    )?;
    assert_save_tokens_advance_once(&fixture.bytes, &focused_set_bytes)?;
    assert_legacy_save_tokens_monotonic(&fixture.bytes, &generic_set_bytes)?;

    let focused_clear_package = Package::from_bytes(&fixture.bytes)?;
    let focused_clear_bytes = exact_bytes(
        focused_clear_package
            .edit_slide_drawable_comment(
                SlideSelector::index(0),
                DrawableSelector::index(fixture.target.position),
            )?
            .clear()?
            .commit()?
            .package(),
    )?;
    let mut generic_clear = IWorkDrawableCommentEditor::from_bytes(&fixture.bytes)?;
    generic_clear.clear_comment(DrawableId::from_raw(fixture.target.identifier)?)?;
    let generic_clear_bytes = generic_clear.to_bytes()?;
    assert_shared_cow_parity(
        &fixture,
        &fixture.bytes,
        &fixture.bytes,
        &focused_clear_bytes,
        &generic_clear_bytes,
        None,
        original_sibling_text.as_deref(),
    )?;
    assert_cow_uuid_registration(&fixture.bytes, &focused_clear_bytes, root, root)?;
    assert_cow_uuid_registration(&fixture.bytes, &generic_clear_bytes, root, root)?;
    assert_save_tokens_advance_once(&fixture.bytes, &focused_clear_bytes)?;
    assert_legacy_save_tokens_monotonic(&fixture.bytes, &generic_clear_bytes)?;
    Ok(())
}

#[test]
fn shared_foreign_reply_clear_preserves_sibling_branch_in_both_backends() -> TestResult {
    let fixture = fixtures::shared_foreign_reply_fixture()?;
    let sibling = fixture
        .sibling
        .as_ref()
        .ok_or_else(|| io::Error::other("shared-reply fixture has no sibling"))?;
    let focused_package = Package::from_bytes(&fixture.bytes)?;
    let focused_bytes = exact_bytes(
        focused_package
            .edit_slide_drawable_comment(
                SlideSelector::index(0),
                DrawableSelector::index(fixture.target.position),
            )?
            .clear()?
            .commit()?
            .package(),
    )?;
    let mut generic = IWorkDrawableCommentEditor::from_bytes(&fixture.bytes)?;
    generic.clear_comment(DrawableId::from_raw(fixture.target.identifier)?)?;
    let generic_bytes = generic.to_bytes()?;
    assert_backend_parity(
        &fixture.bytes,
        &fixture.bytes,
        &focused_bytes,
        &generic_bytes,
        &[&fixture.target, sibling],
    )?;
    let focused_oracle = IWorkDrawableCommentEditor::from_bytes(&focused_bytes)?;
    let generic_oracle = IWorkDrawableCommentEditor::from_bytes(&generic_bytes)?;
    assert!(
        focused_oracle
            .comment(DrawableId::from_raw(fixture.target.identifier)?)?
            .is_none()
    );
    assert!(
        generic_oracle
            .comment(DrawableId::from_raw(fixture.target.identifier)?)?
            .is_none()
    );
    assert_reply_component(
        &focused_bytes,
        sibling,
        0,
        fixtures::AUTHOR_STORAGE_COMPONENT,
    )?;
    assert_reply_component(
        &generic_bytes,
        sibling,
        0,
        fixtures::AUTHOR_STORAGE_COMPONENT,
    )?;
    assert_eq!(
        focused_oracle
            .replies(DrawableId::from_raw(sibling.identifier)?)?
            .len(),
        1
    );
    assert_eq!(
        generic_oracle
            .replies(DrawableId::from_raw(sibling.identifier)?)?
            .len(),
        1
    );
    let root = fixture
        .root_identifier
        .ok_or_else(|| io::Error::other("shared-reply fixture has no root"))?;
    let reply = fixture
        .reply_identifier
        .ok_or_else(|| io::Error::other("shared-reply fixture has no reply"))?;
    assert_uuid_registration_removed_if_dead(&fixture.bytes, &focused_bytes, root)?;
    assert_uuid_registration_removed_if_dead(&fixture.bytes, &generic_bytes, root)?;
    assert_cow_uuid_registration(&fixture.bytes, &focused_bytes, reply, reply)?;
    assert_cow_uuid_registration(&fixture.bytes, &generic_bytes, reply, reply)?;
    assert_save_tokens_advance_once(&fixture.bytes, &focused_bytes)?;
    assert_legacy_save_tokens_monotonic(&fixture.bytes, &generic_bytes)?;
    Ok(())
}

#[test]
fn foreign_drawable_read_and_create_match_generic_backend() -> TestResult {
    let fixture = fixtures::foreign_drawable_fixture()?;
    let focused = Package::from_bytes(&fixture.bytes)?;
    let generic = IWorkDrawableCommentEditor::from_bytes(&fixture.bytes)?;
    assert_eq!(
        focused_thread(&focused, &fixture.target)?,
        generic_thread(&generic, &fixture.target)?
    );
    assert!(focused_thread(&focused, &fixture.target)?.root.is_none());

    let focused_bytes = focused_set(&fixture.bytes, &fixture.target, "foreign drawable")?;
    let generic_bytes = generic_set(&fixture.bytes, &fixture.target, "foreign drawable")?;
    assert_backend_parity_with_generated_root(
        &fixture.bytes,
        &fixture.bytes,
        &focused_bytes,
        &generic_bytes,
        &fixture.target,
    )?;
    let focused_root = IWorkDrawableCommentEditor::from_bytes(&focused_bytes)?
        .comment(DrawableId::from_raw(fixture.target.identifier)?)?
        .ok_or_else(|| io::Error::other("focused writer produced no root"))?
        .storage_id
        .get();
    let generic_root = IWorkDrawableCommentEditor::from_bytes(&generic_bytes)?
        .comment(DrawableId::from_raw(fixture.target.identifier)?)?
        .ok_or_else(|| io::Error::other("generic writer produced no root"))?
        .storage_id
        .get();
    assert_foreign_component(&focused_bytes, focused_root, fixtures::STYLESHEET_COMPONENT)?;
    assert_foreign_component(&generic_bytes, generic_root, fixtures::STYLESHEET_COMPONENT)?;
    assert_save_tokens_advance_once(&fixture.bytes, &focused_bytes)?;
    assert_legacy_save_tokens_monotonic(&fixture.bytes, &generic_bytes)?;

    let focused_clear_bytes = focused_clear(&focused_bytes, &fixture.target)?;
    let generic_clear_bytes = generic_clear(&generic_bytes, &fixture.target)?;
    assert_backend_parity(
        &focused_bytes,
        &generic_bytes,
        &focused_clear_bytes,
        &generic_clear_bytes,
        &[&fixture.target],
    )?;
    assert_save_tokens_advance_once(&focused_bytes, &focused_clear_bytes)?;
    assert_legacy_save_tokens_monotonic(&generic_bytes, &generic_clear_bytes)?;
    assert_eq!(
        native_object_ids(&focused_clear_bytes)?,
        native_object_ids(&fixture.bytes)?,
        "focused author/root cleanup left an object behind"
    );
    assert_eq!(
        native_object_ids(&generic_clear_bytes)?,
        native_object_ids(&fixture.bytes)?,
        "generic author/root cleanup left an object behind"
    );
    assert_generated_comment_cleanup(&fixture.bytes, &focused_clear_bytes)?;
    assert_generated_comment_cleanup(&fixture.bytes, &generic_clear_bytes)?;
    assert_archive_references_are_live(&focused_clear_bytes)?;
    assert_archive_references_are_live(&generic_clear_bytes)?;
    Ok(())
}

#[test]
fn stale_focused_patch_rejects_atomically_after_generic_foreign_edit() -> TestResult {
    let fixture = fixtures::root_foreign_fixture()?;
    let source_package = Package::from_bytes(&fixture.bytes)?;
    let staged = source_package
        .edit_slide_drawable_comment(
            SlideSelector::index(0),
            DrawableSelector::index(fixture.target.position),
        )?
        .set("staged focused edit")?
        .commit()?;
    let generic_bytes = generic_set(&fixture.bytes, &fixture.target, "generic edit")?;
    let generic_package = Package::from_bytes(&generic_bytes)?;
    let before = exact_bytes(&generic_package)?;
    let inverse = staged.patch().inverse();
    let error = generic_package.apply_slide_drawable_comment(&inverse);
    assert!(error.is_err(), "stale cross-component patch was accepted");
    assert_eq!(exact_bytes(&generic_package)?, before);
    Ok(())
}
