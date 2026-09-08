//! Native drawable-comment graph fixtures with component relocation.
//!
//! The focused comment API intentionally hides archive identifiers.  These
//! helpers are test-only physical oracles: they discover the identifiers from
//! the checked-in native package, relocate one archive object, and update the
//! package metadata witnesses which make that relocation valid.  No component
//! or object identifier is part of the public test contract.

use std::{
    collections::{BTreeMap, BTreeSet},
    io,
};

use litchi_iwa_archive::iwa::{Archive, ArchiveObject, RawMessage, SnappyStream};
use litchi_iwa_archive::{Limits, package::Catalog};
use litchi_iwa_common::wire::{WireView, append_length_delimited_field};
use litchi_iwa_protos::{kn, tsd, tsp};
use litchi_keynote::{DrawableSelector, Package, SlideSelector};
use prost::Message as _;

pub(crate) type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

pub(crate) const SOURCE: &[u8] =
    include_bytes!("../../../../test-data/iwork/keynote/drawable-comments-source-native.key");
pub(crate) const COMMENT_STORAGE_MESSAGE_TYPE: u32 = 3_056;
pub(crate) const SLIDE_MESSAGE_TYPE: u32 = 5;
pub(crate) const PACKAGE_METADATA_MESSAGE_TYPE: u32 = 11_006;

/// Existing native component destinations.  They are present in the checked-
/// in package and therefore do not require synthetic PackageMetadata roots.
pub(crate) const AUTHOR_STORAGE_COMPONENT: &str = "Index/AnnotationAuthorStorage.iwa";
pub(crate) const STYLESHEET_COMPONENT: &str = "Index/DocumentStylesheet.iwa";

const ROUTES: &[(u32, &[u32])] = &[
    (3_002, &[6]),
    (3_004, &[1, 6]),
    (3_005, &[1, 6]),
    (3_006, &[1, 6]),
    (3_007, &[1, 6]),
    (3_008, &[1, 6]),
    (5_021, &[1, 6]),
    (6_000, &[1, 6]),
    (3_009, &[1, 1, 6]),
    (2_011, &[1, 1, 6]),
    (6_007, &[1, 1, 6]),
    (2_014, &[1, 1, 1, 6]),
    (7, &[1, 1, 1, 6]),
    (12, &[1, 1, 1, 6]),
];

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct DrawableTarget {
    pub position: usize,
    pub identifier: u64,
    message_index: usize,
    message_type: u32,
    comment_path: &'static [u32],
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct RelocatedFixture {
    /// The valid same-component source immediately before the physical move.
    /// This is retained so tests can compare metadata tokens and UUID shape.
    pub metadata_source: Vec<u8>,
    pub bytes: Vec<u8>,
    pub target: DrawableTarget,
    pub sibling: Option<DrawableTarget>,
    pub root_identifier: Option<u64>,
    pub reply_identifier: Option<u64>,
}

#[derive(Debug, Clone)]
struct Profile {
    drawables: Vec<DrawableTarget>,
    commented: usize,
    empty: Vec<usize>,
}

#[derive(Debug, Clone)]
struct LocatedObject {
    component: String,
    object: ArchiveObject,
}

fn exact_bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

fn route_path(message_type: u32) -> Option<&'static [u32]> {
    ROUTES
        .iter()
        .find_map(|(kind, path)| (*kind == message_type).then_some(*path))
}

pub(crate) fn component_archives(source: &[u8]) -> TestResult<Vec<(String, Archive)>> {
    let mut archives = Vec::new();
    for entry in Catalog::from_bytes(source)?.iter() {
        if !entry.name().ends_with(".iwa") {
            continue;
        }
        let stream = match SnappyStream::decompress(entry.data()) {
            Ok(stream) => stream.into_bytes(),
            Err(_) => continue,
        };
        let archive = match Archive::parse(&stream) {
            Ok(archive) => archive,
            Err(_) => continue,
        };
        archives.push((entry.name().to_owned(), archive));
    }
    Ok(archives)
}

fn component_containing_object(source: &[u8], identifier: u64) -> TestResult<LocatedObject> {
    component_archives(source)?
        .into_iter()
        .find_map(|(component, archive)| {
            archive
                .object(identifier)
                .cloned()
                .map(|object| LocatedObject { component, object })
        })
        .ok_or_else(|| io::Error::other(format!("missing archive object {identifier}")).into())
}

fn replace_component(source: &[u8], component_name: &str, archive: Archive) -> TestResult<Vec<u8>> {
    let bytes = SnappyStream::compress(&archive.to_bytes()?)?;
    let catalog = Catalog::from_bytes(source)?;
    let entries = catalog
        .iter()
        .map(|entry| {
            if entry.name() == component_name {
                (entry.name(), bytes.as_slice())
            } else {
                (entry.name(), entry.data())
            }
        })
        .collect::<Vec<_>>();
    Ok(litchi_iwa_archive::package::to_bytes(
        entries,
        Limits::default(),
    )?)
}

fn first_slide(source: &[u8]) -> TestResult<(String, u64, Vec<u64>)> {
    for (component, archive) in component_archives(source)? {
        if !component.starts_with("Index/Slide-") {
            continue;
        }
        for object in &archive.objects {
            let Some(message) = object
                .messages
                .iter()
                .find(|message| message.type_ == SLIDE_MESSAGE_TYPE)
            else {
                continue;
            };
            let slide = kn::SlideArchive::decode(message.data.as_slice())?;
            if slide.owned_drawables.is_empty() {
                continue;
            }
            let identifier = object
                .archive_info
                .identifier
                .ok_or_else(|| io::Error::other("slide object has no identifier"))?;
            return Ok((
                component,
                identifier,
                slide
                    .owned_drawables
                    .into_iter()
                    .map(|reference| reference.identifier)
                    .collect(),
            ));
        }
    }
    Err(io::Error::other("native fixture has no rooted slide").into())
}

fn nested_reference(payload: &[u8], path: &[u32]) -> TestResult<Option<u64>> {
    let mut current = payload;
    for field_number in path {
        let fields = WireView::parse(current)?
            .fields()
            .filter(|field| field.number() == *field_number)
            .collect::<Vec<_>>();
        if fields.len() > 1 {
            return Err(io::Error::other("drawable route repeats a reference field").into());
        }
        let Some(field) = fields.first() else {
            return Ok(None);
        };
        if field.wire_type() != 2 {
            return Err(io::Error::other("drawable reference is not length-delimited").into());
        }
        current = field.payload();
    }
    let fields = WireView::parse(current)?
        .fields()
        .filter(|field| field.number() == 1)
        .collect::<Vec<_>>();
    if fields.is_empty() {
        return Err(io::Error::other("drawable reference has no identifier").into());
    }
    if fields.len() != 1 || fields[0].wire_type() != 0 {
        return Err(io::Error::other("drawable reference has duplicate identifier").into());
    }
    let reference = tsp::Reference::decode(current)?;
    if reference.identifier == 0 {
        return Err(io::Error::other("drawable reference has zero identifier").into());
    }
    Ok(Some(reference.identifier))
}

fn profile(source: &[u8]) -> TestResult<Profile> {
    let (_, _, identifiers) = first_slide(source)?;
    let mut drawables = Vec::new();
    for (position, identifier) in identifiers.into_iter().enumerate() {
        let located = component_containing_object(source, identifier)?;
        let Some((message_index, message)) = located
            .object
            .messages
            .iter()
            .enumerate()
            .find(|(_, message)| route_path(message.type_).is_some())
        else {
            continue;
        };
        let comment_path = route_path(message.type_).expect("route selected above");
        let _ = nested_reference(&message.data, comment_path)?;
        drawables.push(DrawableTarget {
            position,
            identifier,
            message_index,
            message_type: message.type_,
            comment_path,
        });
    }
    let package = Package::from_bytes(source)?;
    let summaries = package.slide_drawables(SlideSelector::index(0))?;
    if summaries.len() != drawables.len() {
        return Err(io::Error::other(format!(
            "native drawable census {} != focused census {}",
            drawables.len(),
            summaries.len()
        ))
        .into());
    }
    let commented = summaries
        .iter()
        .position(|summary| summary.has_comment())
        .ok_or_else(|| io::Error::other("native fixture has no comment"))?;
    let empty = summaries
        .iter()
        .enumerate()
        .filter_map(|(index, summary)| (!summary.has_comment()).then_some(index))
        .collect::<Vec<_>>();
    if empty.len() < 2 {
        return Err(io::Error::other("native fixture has fewer than two empty drawables").into());
    }
    Ok(Profile {
        drawables,
        commented,
        empty,
    })
}

fn target_comment_id(source: &[u8], target: &DrawableTarget) -> TestResult<Option<u64>> {
    let object = component_containing_object(source, target.identifier)?.object;
    nested_reference(
        &object
            .messages
            .get(target.message_index)
            .ok_or_else(|| io::Error::other("drawable message disappeared"))?
            .data,
        target.comment_path,
    )
}

fn append_comment_reference(payload: &[u8], path: &[u32], identifier: u64) -> TestResult<Vec<u8>> {
    let Some((&field_number, remainder)) = path.split_first() else {
        return Err(io::Error::other("empty comment route").into());
    };
    let mut output = payload.to_vec();
    if remainder.is_empty() {
        let reference = tsp::Reference {
            identifier,
            ..Default::default()
        }
        .encode_to_vec();
        append_length_delimited_field(&mut output, field_number, &reference)?;
        return Ok(output);
    }
    let fields = WireView::parse(payload)?
        .fields()
        .filter(|field| field.number() == field_number)
        .collect::<Vec<_>>();
    if fields.len() != 1 || fields[0].wire_type() != 2 {
        return Err(io::Error::other("drawable route has no unique envelope").into());
    }
    let rewritten = append_comment_reference(fields[0].payload(), remainder, identifier)?;
    Ok(litchi_iwa_common::wire::transform_length_delimited_field(
        payload,
        field_number,
        |_| -> TestResult<Vec<u8>> { Ok(rewritten.clone()) },
    )?)
}

fn attach_comment(source: &[u8], target: &DrawableTarget, comment: u64) -> TestResult<Vec<u8>> {
    let located = component_containing_object(source, target.identifier)?;
    let mut archive = component_archives(source)?
        .into_iter()
        .find_map(|(name, archive)| (name == located.component).then_some(archive))
        .ok_or_else(|| io::Error::other("drawable component disappeared"))?;
    let object = archive
        .object_mut(target.identifier)
        .ok_or_else(|| io::Error::other("drawable disappeared"))?;
    let message = object
        .messages
        .get(target.message_index)
        .ok_or_else(|| io::Error::other("drawable message disappeared"))?;
    if nested_reference(&message.data, target.comment_path)?.is_some() {
        return Err(io::Error::other("drawable already has a comment").into());
    }
    let payload = append_comment_reference(&message.data, target.comment_path, comment)?;
    let info = object
        .archive_info
        .message_infos
        .get_mut(target.message_index)
        .ok_or_else(|| io::Error::other("drawable message metadata disappeared"))?;
    if !info.object_references.contains(&comment) {
        info.object_references.push(comment);
    }
    object.replace_message(
        target.message_index,
        RawMessage {
            type_: target.message_type,
            data: payload,
        },
    )?;
    replace_component(source, &located.component, archive)
}

fn comment_object(
    source: &[u8],
    identifier: u64,
) -> TestResult<(String, Archive, usize, tsd::CommentStorageArchive)> {
    let located = component_containing_object(source, identifier)?;
    let (message_index, payload) = located
        .object
        .messages
        .iter()
        .enumerate()
        .find(|(_, message)| message.type_ == COMMENT_STORAGE_MESSAGE_TYPE)
        .map(|(index, message)| (index, message.data.clone()))
        .ok_or_else(|| io::Error::other("comment object has no storage message"))?;
    let archive = component_archives(source)?
        .into_iter()
        .find_map(|(name, archive)| (name == located.component).then_some(archive))
        .ok_or_else(|| io::Error::other("comment component disappeared"))?;
    Ok((
        located.component,
        archive,
        message_index,
        tsd::CommentStorageArchive::decode(payload.as_slice())?,
    ))
}

fn add_reply(source: &[u8], target: &DrawableTarget) -> TestResult<(Vec<u8>, u64)> {
    target_comment_id(source, target)?
        .ok_or_else(|| io::Error::other("reply fixture requires an existing root"))?;
    let package = Package::from_bytes(source)?;
    let commit = package
        .edit_slide_drawable_comment(
            SlideSelector::index(0),
            DrawableSelector::index(target.position),
        )?
        .add_reply("cross-component fixture reply")?
        .commit()?;
    let bytes = exact_bytes(commit.package())?;
    let root = target_comment_id(&bytes, target)?
        .ok_or_else(|| io::Error::other("reply writer lost its root"))?;
    let (_, _, _, comment) = comment_object(&bytes, root)?;
    let reply = comment
        .replies
        .last()
        .map(|reference| reference.identifier)
        .ok_or_else(|| io::Error::other("reply writer produced no reply"))?;
    Ok((bytes, reply))
}

fn metadata_payload(source: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(source)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == "Index/Metadata.iwa")
        .ok_or_else(|| io::Error::other("missing metadata component"))?;
    let stream = SnappyStream::decompress(entry.data())?.into_bytes();
    let archive = Archive::parse(&stream)?;
    archive
        .objects
        .iter()
        .flat_map(|object| &object.messages)
        .find(|message| message.type_ == PACKAGE_METADATA_MESSAGE_TYPE)
        .map(|message| message.data.clone())
        .ok_or_else(|| io::Error::other("missing PackageMetadata payload").into())
}

fn replace_metadata_payload(source: &[u8], metadata: &tsp::PackageMetadata) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(source)?;
    let metadata_entry = catalog
        .iter()
        .find(|entry| entry.name() == "Index/Metadata.iwa")
        .ok_or_else(|| io::Error::other("metadata component disappeared"))?;
    let metadata_stream = SnappyStream::decompress(metadata_entry.data())?.into_bytes();
    let mut metadata_archive = Archive::parse(&metadata_stream)?;
    let object = metadata_archive
        .objects
        .iter_mut()
        .find(|object| {
            object
                .messages
                .iter()
                .any(|message| message.type_ == PACKAGE_METADATA_MESSAGE_TYPE)
        })
        .ok_or_else(|| io::Error::other("metadata root disappeared"))?;
    let message = object
        .messages
        .iter_mut()
        .find(|message| message.type_ == PACKAGE_METADATA_MESSAGE_TYPE)
        .ok_or_else(|| io::Error::other("metadata payload disappeared"))?;
    message.data = metadata.encode_to_vec();
    replace_component(source, "Index/Metadata.iwa", metadata_archive)
}

/// Register a generated comment object using the UUID carried by its own
/// `CommentStorageArchive` payload.  Native-created replies may carry that
/// identity without a PackageMetadata map entry; this fixture makes the
/// registered path explicit before relocation while rejecting any collision.
fn register_object_uuid(source: &[u8], identifier: u64) -> TestResult<Vec<u8>> {
    let located = component_containing_object(source, identifier)?;
    let (_, _, _, comment) = comment_object(source, identifier)?;
    let uuid = comment
        .storage_uuid
        .clone()
        .ok_or_else(|| io::Error::other("comment object has no storage UUID"))?;
    let mut metadata = tsp::PackageMetadata::decode(metadata_payload(source)?.as_slice())?;
    if metadata
        .components
        .iter()
        .flat_map(|component| &component.object_uuid_map_entries)
        .any(|entry| entry.identifier == identifier)
    {
        return Err(io::Error::other("generated comment UUID already registered").into());
    }
    let component_identifier = component_id_for_name(&metadata, &located.component)?;
    let component = metadata
        .components
        .iter_mut()
        .find(|component| component.identifier == component_identifier)
        .ok_or_else(|| io::Error::other("comment component metadata disappeared"))?;
    component
        .object_uuid_map_entries
        .push(tsp::ObjectUuidMapEntry { identifier, uuid });
    let registered = replace_metadata_payload(source, &metadata)?;
    let registered_metadata =
        tsp::PackageMetadata::decode(metadata_payload(&registered)?.as_slice())?;
    assert_eq!(
        metadata.last_object_identifier,
        registered_metadata.last_object_identifier
    );
    assert_eq!(
        metadata.save_token, registered_metadata.save_token,
        "UUID registration must preserve the package watermark"
    );
    let registered_component = registered_metadata
        .components
        .iter()
        .find(|component| component.identifier == component_identifier)
        .ok_or_else(|| io::Error::other("registered component disappeared"))?;
    let registered_entry = registered_component
        .object_uuid_map_entries
        .iter()
        .find(|entry| entry.identifier == identifier)
        .ok_or_else(|| io::Error::other("registered UUID entry disappeared"))?;
    assert_eq!(
        (registered_entry.uuid.lower, registered_entry.uuid.upper),
        (uuid.lower, uuid.upper)
    );
    Ok(registered)
}

fn metadata_locators(metadata: &tsp::PackageMetadata) -> BTreeMap<String, u64> {
    metadata
        .components
        .iter()
        .map(|component| {
            (
                component
                    .locator
                    .as_deref()
                    .unwrap_or(&component.preferred_locator)
                    .to_owned(),
                component.identifier,
            )
        })
        .collect()
}

fn physical_component_locator(name: &str) -> &str {
    let without_prefix = name.strip_prefix("Index/").unwrap_or(name);
    without_prefix
        .strip_suffix(".iwa")
        .unwrap_or(without_prefix)
}

fn component_id_for_name(metadata: &tsp::PackageMetadata, name: &str) -> TestResult<u64> {
    let locators = metadata_locators(metadata);
    let expected = physical_component_locator(name);
    locators
        .iter()
        .find_map(|(locator, identifier)| {
            (physical_component_locator(locator) == expected).then_some(*identifier)
        })
        .ok_or_else(|| io::Error::other(format!("metadata has no component {name}")).into())
}

fn object_component_map(source: &[u8]) -> TestResult<BTreeMap<u64, String>> {
    let mut result = BTreeMap::new();
    for (component, archive) in component_archives(source)? {
        for object in archive.objects {
            let identifier = object
                .archive_info
                .identifier
                .ok_or_else(|| io::Error::other("archive object has no identifier"))?;
            if result.insert(identifier, component.clone()).is_some() {
                return Err(
                    io::Error::other(format!("duplicate archive object {identifier}")).into(),
                );
            }
        }
    }
    Ok(result)
}

fn object_refs(object: &ArchiveObject) -> BTreeSet<u64> {
    object
        .archive_info
        .message_infos
        .iter()
        .flat_map(|info| {
            info.object_references.iter().copied().chain(
                info.field_infos
                    .iter()
                    .flat_map(|field| field.object_references.iter().copied()),
            )
        })
        .filter(|identifier| *identifier != 0)
        .collect()
}

fn add_current_edge(component: &mut tsp::ComponentInfo, target: u64, object: Option<u64>) {
    if component.external_references.iter().any(|reference| {
        reference.component_identifier == target && reference.object_identifier == object
    }) {
        return;
    }
    component
        .external_references
        .push(tsp::ComponentExternalReference {
            component_identifier: target,
            object_identifier: object,
            is_weak: None,
        });
}

fn rewrite_component_edges(
    metadata: &mut tsp::PackageMetadata,
    old_component: u64,
    new_component: u64,
    object_identifier: u64,
    owner_components: &BTreeSet<u64>,
    outgoing_components: &BTreeSet<(u64, Option<u64>)>,
) {
    for component in &mut metadata.components {
        for reference in &mut component.external_references {
            if reference.component_identifier == old_component
                && reference.object_identifier == Some(object_identifier)
            {
                reference.component_identifier = new_component;
            }
        }
    }
    for owner in owner_components {
        if *owner != new_component {
            if let Some(component) = metadata
                .components
                .iter_mut()
                .find(|component| component.identifier == *owner)
            {
                add_current_edge(component, new_component, Some(object_identifier));
            }
        }
    }
    if let Some(destination) = metadata
        .components
        .iter_mut()
        .find(|component| component.identifier == new_component)
    {
        for (target, object) in outgoing_components {
            if *target != new_component {
                add_current_edge(destination, *target, *object);
            }
        }
    }
}

fn rewrite_metadata_for_relocation(
    source_before: &[u8],
    source_after: &[u8],
    object_identifier: u64,
    source_component: &str,
    destination_component: &str,
    moved_object: &ArchiveObject,
) -> TestResult<Vec<u8>> {
    let mut metadata = tsp::PackageMetadata::decode(metadata_payload(source_before)?.as_slice())?;
    let old_component = component_id_for_name(&metadata, source_component)?;
    let new_component = component_id_for_name(&metadata, destination_component)?;
    if old_component == new_component {
        return Err(io::Error::other("relocation destination equals source").into());
    }
    let mut owner_components = BTreeSet::new();
    let mut outgoing_components = BTreeSet::new();
    for (component, archive) in component_archives(source_before)? {
        let component_id = component_id_for_name(&metadata, &component).ok();
        for object in &archive.objects {
            let Some(identifier) = object.archive_info.identifier else {
                continue;
            };
            if identifier != object_identifier && object_refs(object).contains(&object_identifier) {
                if let Some(component_id) = component_id {
                    owner_components.insert(component_id);
                }
            }
        }
    }
    let object_components_after = object_component_map(source_after)?;
    for reference in object_refs(moved_object) {
        let target_name = object_components_after.get(&reference).ok_or_else(|| {
            io::Error::other(format!(
                "moved object references missing object {reference}"
            ))
        })?;
        let target_component = component_id_for_name(&metadata, target_name)?;
        outgoing_components.insert((target_component, Some(reference)));
    }
    rewrite_component_edges(
        &mut metadata,
        old_component,
        new_component,
        object_identifier,
        &owner_components,
        &outgoing_components,
    );

    // Preserve an existing ObjectUUIDMap registration, including its UUID,
    // while moving it between the two current component namespaces.  A
    // native object without a registration stays unregistered.
    let mut moved_uuid = None;
    if let Some(component) = metadata
        .components
        .iter_mut()
        .find(|component| component.identifier == old_component)
    {
        if let Some(index) = component
            .object_uuid_map_entries
            .iter()
            .position(|entry| entry.identifier == object_identifier)
        {
            moved_uuid = Some(component.object_uuid_map_entries.remove(index));
        }
    }
    if let Some(entry) = moved_uuid {
        if let Some(component) = metadata
            .components
            .iter_mut()
            .find(|component| component.identifier == new_component)
        {
            if component
                .object_uuid_map_entries
                .iter()
                .any(|existing| existing.identifier == object_identifier)
            {
                return Err(
                    io::Error::other("destination already owns moved ObjectUUIDMap entry").into(),
                );
            }
            component.object_uuid_map_entries.push(entry);
        }
    }

    replace_metadata_payload(source_after, &metadata)
}

/// Move an existing archive object and repair all PackageMetadata witnesses.
///
/// The returned package remains source-order equivalent; only the physical
/// archive component, current object-edge targets, and optional UUID namespace
/// ownership change. Historical/versioned metadata records are retained as
/// source witnesses and are never rewritten.
pub(crate) fn relocate_object(
    source: &[u8],
    identifier: u64,
    destination_component: &str,
) -> TestResult<Vec<u8>> {
    let located = component_containing_object(source, identifier)?;
    if located.component == destination_component {
        return Err(io::Error::other("relocation destination equals source").into());
    }
    let mut source_archive = component_archives(source)?
        .into_iter()
        .find_map(|(name, archive)| (name == located.component).then_some(archive))
        .ok_or_else(|| io::Error::other("source archive disappeared"))?;
    let moved = source_archive
        .remove_object(identifier)
        .ok_or_else(|| io::Error::other("source object disappeared"))?;
    let source_after_remove = replace_component(source, &located.component, source_archive)?;
    let mut destination_archive = component_archives(&source_after_remove)?
        .into_iter()
        .find_map(|(name, archive)| (name == destination_component).then_some(archive))
        .ok_or_else(|| {
            io::Error::other(format!(
                "destination archive {destination_component} missing"
            ))
        })?;
    destination_archive.insert_object(moved.clone())?;
    let source_after_insert = replace_component(
        &source_after_remove,
        destination_component,
        destination_archive,
    )?;
    rewrite_metadata_for_relocation(
        source,
        &source_after_insert,
        identifier,
        &located.component,
        destination_component,
        &moved,
    )
}

/// Return the first native commented drawable and the first empty drawable.
pub(crate) fn primary_targets(source: &[u8]) -> TestResult<(DrawableTarget, DrawableTarget)> {
    let profile = profile(source)?;
    let commented = profile.drawables[profile.commented].clone();
    let empty_position = profile
        .empty
        .iter()
        .copied()
        .find(|position| {
            matches!(
                profile.drawables[*position].message_type,
                2_011 | 2_014 | 3_002 | 3_004
            )
        })
        .unwrap_or(profile.empty[0]);
    let empty = profile.drawables[empty_position].clone();
    Ok((commented, empty))
}

/// Build a root-comment candidate whose storage object lives in the existing
/// author-storage component.
pub(crate) fn root_foreign_fixture() -> TestResult<RelocatedFixture> {
    let (target, sibling) = primary_targets(SOURCE)?;
    let root = target_comment_id(SOURCE, &target)?
        .ok_or_else(|| io::Error::other("native fixture has no root"))?;
    let bytes = relocate_object(SOURCE, root, AUTHOR_STORAGE_COMPONENT)?;
    Ok(RelocatedFixture {
        metadata_source: SOURCE.to_vec(),
        bytes,
        target,
        sibling: Some(sibling),
        root_identifier: Some(root),
        reply_identifier: None,
    })
}

/// Build a reply candidate by using the semantic writer for allocation, then
/// relocating only the newly written reply object across components.
pub(crate) fn reply_foreign_fixture() -> TestResult<RelocatedFixture> {
    let (target, sibling) = primary_targets(SOURCE)?;
    let (with_reply, reply) = add_reply(SOURCE, &target)?;
    let with_registered_reply = register_object_uuid(&with_reply, reply)?;
    let root = target_comment_id(&with_registered_reply, &target)?
        .ok_or_else(|| io::Error::other("reply candidate lost root"))?;
    let bytes = relocate_object(&with_registered_reply, reply, AUTHOR_STORAGE_COMPONENT)?;
    Ok(RelocatedFixture {
        metadata_source: with_registered_reply,
        bytes,
        target,
        sibling: Some(sibling),
        root_identifier: Some(root),
        reply_identifier: Some(reply),
    })
}

/// Build two drawable references to one root and relocate that root.  The
/// second drawable remains a source-order sibling and is used by COW tests.
pub(crate) fn shared_foreign_root_fixture() -> TestResult<RelocatedFixture> {
    let (target, sibling) = primary_targets(SOURCE)?;
    let root = target_comment_id(SOURCE, &target)?
        .ok_or_else(|| io::Error::other("native fixture has no root"))?;
    let attached = attach_comment(SOURCE, &sibling, root)?;
    let bytes = relocate_object(&attached, root, AUTHOR_STORAGE_COMPONENT)?;
    Ok(RelocatedFixture {
        metadata_source: attached,
        bytes,
        target,
        sibling: Some(sibling),
        root_identifier: Some(root),
        reply_identifier: None,
    })
}

/// Build two roots sharing one reply, then relocate that reply.  The second
/// root is allocated through the semantic API, so its metadata identity and
/// author closure remain native-valid before the shared edge is introduced.
pub(crate) fn shared_foreign_reply_fixture() -> TestResult<RelocatedFixture> {
    let (target, sibling) = primary_targets(SOURCE)?;
    let (with_reply, reply) = add_reply(SOURCE, &target)?;
    let package = Package::from_bytes(&with_reply)?;
    let second = package
        .edit_slide_drawable_comment(
            SlideSelector::index(0),
            DrawableSelector::index(sibling.position),
        )?
        .set("cross-component sibling root")?
        .commit()?;
    let with_second = exact_bytes(second.package())?;
    let second_root = target_comment_id(&with_second, &sibling)?
        .ok_or_else(|| io::Error::other("second root writer produced no root"))?;
    let with_shared_reply = attach_reply_reference(&with_second, second_root, reply)?;
    let root = target_comment_id(&with_shared_reply, &target)?
        .ok_or_else(|| io::Error::other("shared reply lost primary root"))?;
    let bytes = relocate_object(&with_shared_reply, reply, AUTHOR_STORAGE_COMPONENT)?;
    Ok(RelocatedFixture {
        metadata_source: with_shared_reply,
        bytes,
        target,
        sibling: Some(sibling),
        root_identifier: Some(root),
        reply_identifier: Some(reply),
    })
}

/// Move an existing empty slide-owned drawable to the stylesheet component.
/// The slide's source-order reference remains the semantic owner witness.
pub(crate) fn foreign_drawable_fixture() -> TestResult<RelocatedFixture> {
    let (commented, target) = primary_targets(SOURCE)?;
    let bytes = relocate_object(SOURCE, target.identifier, STYLESHEET_COMPONENT)?;
    Ok(RelocatedFixture {
        metadata_source: SOURCE.to_vec(),
        bytes,
        target,
        sibling: Some(commented),
        root_identifier: None,
        reply_identifier: None,
    })
}

fn attach_reply_reference(
    source: &[u8],
    root_identifier: u64,
    reply_identifier: u64,
) -> TestResult<Vec<u8>> {
    let located = component_containing_object(source, root_identifier)?;
    let mut archive = component_archives(source)?
        .into_iter()
        .find_map(|(name, archive)| (name == located.component).then_some(archive))
        .ok_or_else(|| io::Error::other("root component disappeared"))?;
    let object = archive
        .object_mut(root_identifier)
        .ok_or_else(|| io::Error::other("root object disappeared"))?;
    let message_index = object
        .messages
        .iter()
        .position(|message| message.type_ == COMMENT_STORAGE_MESSAGE_TYPE)
        .ok_or_else(|| io::Error::other("root message disappeared"))?;
    let mut root =
        tsd::CommentStorageArchive::decode(object.messages[message_index].data.as_slice())?;
    if root
        .replies
        .iter()
        .any(|reference| reference.identifier == reply_identifier)
    {
        return Err(io::Error::other("reply already attached").into());
    }
    root.replies.push(tsp::Reference {
        identifier: reply_identifier,
        ..Default::default()
    });
    object.archive_info.message_infos[message_index]
        .object_references
        .push(reply_identifier);
    object.replace_message(
        message_index,
        RawMessage {
            type_: COMMENT_STORAGE_MESSAGE_TYPE,
            data: root.encode_to_vec(),
        },
    )?;
    replace_component(source, &located.component, archive)
}

/// Assert that a relocation kept the metadata shape and moved the selected
/// current UUID identity only once.  This is intentionally semantic: native
/// PackageMetadata may contain opaque fields which the focused codec preserves
/// independently of the test fixture's typed view.
pub(crate) fn assert_metadata_relocated(
    source: &[u8],
    candidate: &[u8],
    moved: u64,
) -> TestResult<()> {
    let before = tsp::PackageMetadata::decode(metadata_payload(source)?.as_slice())?;
    let after = tsp::PackageMetadata::decode(metadata_payload(candidate)?.as_slice())?;
    assert_eq!(before.last_object_identifier, after.last_object_identifier);
    assert_eq!(before.components.len(), after.components.len());
    assert_eq!(before.versioned_components, after.versioned_components);
    let before_tokens = before
        .components
        .iter()
        .map(|component| (component.identifier, component.save_token))
        .collect::<BTreeMap<_, _>>();
    let after_tokens = after
        .components
        .iter()
        .map(|component| (component.identifier, component.save_token))
        .collect::<BTreeMap<_, _>>();
    assert_eq!(before_tokens, after_tokens);
    let before_locators = before
        .components
        .iter()
        .map(|component| {
            (
                component.identifier,
                (
                    component.preferred_locator.clone(),
                    component.locator.clone(),
                ),
            )
        })
        .collect::<BTreeMap<_, _>>();
    let after_locators = after
        .components
        .iter()
        .map(|component| {
            (
                component.identifier,
                (
                    component.preferred_locator.clone(),
                    component.locator.clone(),
                ),
            )
        })
        .collect::<BTreeMap<_, _>>();
    assert_eq!(before_locators, after_locators);
    let before_versioned_edges = before
        .components
        .iter()
        .map(|component| {
            (
                component.identifier,
                component.versioned_external_references.clone(),
            )
        })
        .collect::<BTreeMap<_, _>>();
    let after_versioned_edges = after
        .components
        .iter()
        .map(|component| {
            (
                component.identifier,
                component.versioned_external_references.clone(),
            )
        })
        .collect::<BTreeMap<_, _>>();
    assert_eq!(before_versioned_edges, after_versioned_edges);
    let before_uuid_count = before
        .components
        .iter()
        .flat_map(|component| &component.object_uuid_map_entries)
        .filter(|entry| entry.identifier == moved)
        .count();
    let before_uuid = before
        .components
        .iter()
        .flat_map(|component| &component.object_uuid_map_entries)
        .find(|entry| entry.identifier == moved)
        .map(|entry| (entry.uuid.lower, entry.uuid.upper));
    let after_uuid = after
        .components
        .iter()
        .flat_map(|component| &component.object_uuid_map_entries)
        .find(|entry| entry.identifier == moved)
        .map(|entry| (entry.uuid.lower, entry.uuid.upper));
    assert_eq!(before_uuid, after_uuid);
    let after_uuid_count = after
        .components
        .iter()
        .flat_map(|component| &component.object_uuid_map_entries)
        .filter(|entry| entry.identifier == moved)
        .count();
    assert_eq!(before_uuid_count, after_uuid_count);
    Ok(())
}
