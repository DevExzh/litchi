//! Atomic authoring of inert embedded objects, OLE payloads, and images.

use crate::constants;
use crate::core::OwnedPackage;
use crate::package::charts::{
    Addition, EmbeddedChartHost, ObjectSpan, insert_at_host, locate_objects, rebuild_package,
    splice,
};
use base64::Engine as _;
use base64::engine::general_purpose::STANDARD as BASE64_STANDARD;
use litchi_core::{Error, Result};
use litchi_odf_common::drawing::Part;
use litchi_odf_common::embedded::{Kind, Object, Root, Source, scan_package};
use litchi_odf_common::package::{is_linked_href, resolve_package_path};
use quick_xml::XmlVersion;
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::{BTreeMap, HashSet};

const DRAW_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const OFFICE_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const DRAW_URI: &str = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const OFFICE_URI: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const XLINK_URI: &str = "http://www.w3.org/1999/xlink";
const MAX_FILES: usize = 16_384;
const MAX_BYTES: usize = 64 * 1024 * 1024;
const MAX_INLINE_BYTES: usize = 16 * 1024 * 1024;
pub(crate) const MAX_BATCH_CHANGES: usize = 256;

/// Element kind for an authored embedded package resource.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum EmbeddedResourceKind {
    Object,
    ObjectOle,
    Image,
}

/// One file in an authored embedded `OpenDocument` subdocument.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct EmbeddedResourceFile {
    /// Path relative to the subdocument root, such as `content.xml`.
    pub path: String,
    pub bytes: Vec<u8>,
    /// Exact manifest media type for this file.
    pub media_type: String,
}

/// Storage for a newly authored inert resource.
#[derive(Clone, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum EmbeddedResourceSource {
    /// An external target retained as an inert link. It is never fetched or executed.
    Linked { href: String },
    /// One opaque package file.
    PackageFile {
        bytes: Vec<u8>,
        media_type: String,
        preferred_path: Option<String>,
    },
    /// An embedded `OpenDocument` package rooted at one package directory.
    PackageSubdocument {
        files: Vec<EmbeddedResourceFile>,
        media_type: String,
        preferred_root: Option<String>,
    },
    /// A complete inline `office:document` or `math:math` element.
    InlineXml { root: Root, xml: String },
    /// Base64-encoded into an `office:binary-data` child.
    InlineBinary {
        bytes: Vec<u8>,
        media_type: Option<String>,
    },
}

/// Typed input for adding or replacing an embedded package resource.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct EmbeddedResource {
    pub kind: EmbeddedResourceKind,
    pub source: EmbeddedResourceSource,
    pub frame_name: Option<String>,
    pub xml_id: Option<String>,
    /// Optional OLE class identifier. Only valid for `ObjectOle`.
    pub class_id: Option<String>,
}

/// A base-snapshot embedded-resource selector.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum EmbeddedResourceSelector {
    /// An inert object or OLE object in `content.xml` document order.
    Object(crate::transaction::Position),
    /// An image in `content.xml` document order.
    Image(crate::transaction::Position),
}

/// One base-snapshot change in an atomic embedded-resource batch.
#[derive(Clone, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum EmbeddedResourceChange {
    /// Append one new resource owner to the text body.
    Add(EmbeddedResource),
    /// Replace one selected inner resource while preserving the source frame,
    /// including its name, attributes, and untracked children. A displaced
    /// package payload at a different path is retained; the same path is
    /// overwritten by the replacement payload.
    Replace {
        selector: EmbeddedResourceSelector,
        resource: EmbeddedResource,
    },
    /// Remove one selected resource owner. Package payloads are always retained.
    Remove(EmbeddedResourceSelector),
}

impl EmbeddedResourceChange {
    /// Append one resource.
    #[must_use]
    pub fn add(resource: &EmbeddedResource) -> Self {
        Self::Add(resource.clone())
    }

    /// Replace an object. A displaced payload at a different path is retained;
    /// the same path is overwritten by the replacement payload.
    #[must_use]
    pub fn replace_object(
        position: crate::transaction::Position,
        resource: &EmbeddedResource,
    ) -> Self {
        Self::Replace {
            selector: EmbeddedResourceSelector::Object(position),
            resource: resource.clone(),
        }
    }

    /// Replace an image. A displaced payload at a different path is retained;
    /// the same path is overwritten by the replacement payload.
    #[must_use]
    pub fn replace_image(
        position: crate::transaction::Position,
        resource: &EmbeddedResource,
    ) -> Self {
        Self::Replace {
            selector: EmbeddedResourceSelector::Image(position),
            resource: resource.clone(),
        }
    }

    /// Remove an object and retain any displaced package payload.
    #[must_use]
    pub const fn remove_object(position: crate::transaction::Position) -> Self {
        Self::Remove(EmbeddedResourceSelector::Object(position))
    }

    /// Remove an image and retain any displaced package payload.
    #[must_use]
    pub const fn remove_image(position: crate::transaction::Position) -> Self {
        Self::Remove(EmbeddedResourceSelector::Image(position))
    }
}

#[derive(Clone, Copy, PartialEq, Eq)]
pub(crate) enum ResourceTarget {
    Object,
    Image,
}

struct BuiltResource {
    element_xml: String,
    frame_xml: String,
    additions: Vec<Addition>,
    directories: Vec<(String, String)>,
}

#[derive(Clone)]
struct FrameOwner {
    span: ObjectSpan,
    inner_start: usize,
    inner_end: usize,
}

struct PlannedRemoval {
    span: ObjectSpan,
    owner: usize,
}

#[derive(Clone)]
enum StoredLocation {
    File(String),
    Directory(String),
}

pub(crate) fn add(
    package: &OwnedPackage,
    content: &str,
    styles: Option<&str>,
    host: EmbeddedChartHost<'_>,
    resource: &EmbeddedResource,
) -> Result<(Vec<u8>, usize)> {
    let target = target(resource.kind);
    let index = target_count(package, content, styles, target)?;
    let built = build_resource(package, resource, None, &HashSet::new())?;
    let updated = insert_at_host(content, host, &built.frame_xml)?;
    let bytes = rebuild_package(
        package,
        &updated,
        built.additions,
        built.directories,
        Vec::new(),
        Vec::new(),
    )?;
    Ok((bytes, index))
}

pub(crate) fn replace(
    package: &OwnedPackage,
    content: &str,
    styles: Option<&str>,
    index: usize,
    target: ResourceTarget,
    resource: &EmbeddedResource,
) -> Result<Vec<u8>> {
    if self::target(resource.kind) != target {
        return invalid("replacement resource kind does not match the selected API");
    }
    let spans = selected_spans(package, content, styles, target)?;
    let locations = selected_locations(package, content, styles, target)?;
    let span = spans
        .get(index)
        .cloned()
        .ok_or_else(|| bounds(index, spans.len()))?;
    let old = locations.get(index).cloned().ok_or_else(|| {
        Error::InvalidFormat("embedded-resource span/location scan disagreement".to_string())
    })?;
    let mut built = build_resource(package, resource, old.as_ref(), &HashSet::new())?;
    let updated = splice(content, span.start, span.end, &built.element_xml)?;
    remove_exact_additions(package, &mut built.additions)?;
    remove_exact_directories(package, &mut built.directories)?;
    if updated == content && built.additions.is_empty() && built.directories.is_empty() {
        return Ok(package.as_bytes().to_vec());
    }
    rebuild_package(
        package,
        &updated,
        built.additions,
        built.directories,
        Vec::<String>::new(),
        Vec::<String>::new(),
    )
}

pub(crate) fn remove(
    package: &OwnedPackage,
    content: &str,
    styles: Option<&str>,
    index: usize,
    target: ResourceTarget,
) -> Result<Vec<u8>> {
    let spans = selected_spans(package, content, styles, target)?;
    let span = spans
        .get(index)
        .cloned()
        .ok_or_else(|| bounds(index, spans.len()))?;
    let frames = locate_frames(content)?;
    let owner = *resource_owner_indices(&frames, &spans)?
        .get(index)
        .ok_or_else(|| bounds(index, spans.len()))?;
    let mut planned = Vec::new();
    plan_owner_removals(
        content,
        &frames,
        &HashSet::new(),
        vec![PlannedRemoval { span, owner }],
        &mut planned,
    )?;
    let removal = planned.pop().ok_or_else(|| {
        Error::InvalidFormat("embedded-resource removal produced no owner splice".to_string())
    })?;
    let updated = splice(content, removal.0.start, removal.0.end, "")?;
    rebuild_package(
        package,
        &updated,
        Vec::new(),
        Vec::new(),
        Vec::<String>::new(),
        Vec::<String>::new(),
    )
}

pub(crate) fn reorder(
    package: &OwnedPackage,
    content: &str,
    styles: Option<&str>,
    from: usize,
    to: usize,
    target: ResourceTarget,
) -> Result<Vec<u8>> {
    let spans = selected_spans(package, content, styles, target)?;
    let first = spans.get(from).ok_or_else(|| bounds(from, spans.len()))?;
    let second = spans.get(to).ok_or_else(|| bounds(to, spans.len()))?;
    if from == to {
        return rebuild_package(
            package,
            content,
            Vec::new(),
            Vec::new(),
            Vec::new(),
            Vec::new(),
        );
    }
    let mut updated = String::with_capacity(content.len());
    if first.start < second.start {
        updated.push_str(&content[..first.start]);
        updated.push_str(&content[first.end..second.end]);
        updated.push_str(&content[first.start..first.end]);
        updated.push_str(&content[second.end..]);
    } else {
        updated.push_str(&content[..second.start]);
        updated.push_str(&content[first.start..first.end]);
        updated.push_str(&content[second.start..first.start]);
        updated.push_str(&content[first.end..]);
    }
    rebuild_package(
        package,
        &updated,
        Vec::new(),
        Vec::new(),
        Vec::new(),
        Vec::new(),
    )
}

/// Apply one base-snapshot resource batch with one package publication.
pub(crate) fn apply_batch(
    package: &OwnedPackage,
    content: &str,
    styles: Option<&str>,
    changes: &[EmbeddedResourceChange],
) -> Result<(Vec<u8>, Vec<usize>)> {
    validate_batch_limits(changes)?;
    if changes.is_empty() {
        return Ok((package.as_bytes().to_vec(), Vec::new()));
    }

    let object_spans = selected_spans(package, content, styles, ResourceTarget::Object)?;
    let object_locations = selected_locations(package, content, styles, ResourceTarget::Object)?;
    let image_spans = selected_spans(package, content, styles, ResourceTarget::Image)?;
    let image_locations = selected_locations(package, content, styles, ResourceTarget::Image)?;
    let frames = locate_frames(content)?;
    let object_owners = resource_owner_indices(&frames, &object_spans)?;
    let image_owners = resource_owner_indices(&frames, &image_spans)?;
    let (removed_objects, removed_images) = removed_resource_counts(changes)?;
    let mut selected_targets = HashSet::new();
    let mut reserved_paths = HashSet::new();
    let mut additions = Vec::new();
    let mut directories = Vec::new();
    let mut splices = Vec::new();
    let mut removals = Vec::new();
    let mut replacement_owners = HashSet::new();
    let mut appended_frames = String::new();
    let mut added_indices = Vec::new();
    let mut next_object = object_spans
        .len()
        .checked_sub(removed_objects)
        .ok_or_else(|| bounds(removed_objects, object_spans.len()))?;
    let mut next_image = image_spans
        .len()
        .checked_sub(removed_images)
        .ok_or_else(|| bounds(removed_images, image_spans.len()))?;

    for change in changes {
        match change {
            EmbeddedResourceChange::Add(resource) => {
                let built = build_resource(package, resource, None, &reserved_paths)?;
                reserve_built_paths(&built, &mut reserved_paths)?;
                match target(resource.kind) {
                    ResourceTarget::Object => {
                        added_indices.push(next_object);
                        next_object = next_object.checked_add(1).ok_or_else(|| {
                            Error::InvalidFormat("embedded-object batch index overflow".to_string())
                        })?;
                    },
                    ResourceTarget::Image => {
                        added_indices.push(next_image);
                        next_image = next_image.checked_add(1).ok_or_else(|| {
                            Error::InvalidFormat("embedded-image batch index overflow".to_string())
                        })?;
                    },
                }
                appended_frames.push_str(&built.frame_xml);
                additions.extend(built.additions);
                directories.extend(built.directories);
            },
            EmbeddedResourceChange::Replace { selector, resource } => {
                if !selected_targets.insert(*selector) {
                    return invalid(
                        "embedded-resource batch selects one base owner more than once",
                    );
                }
                let (span, owner, old, expected) = batch_selection(
                    *selector,
                    &object_spans,
                    &object_owners,
                    &object_locations,
                    &image_spans,
                    &image_owners,
                    &image_locations,
                )?;
                if target(resource.kind) != expected {
                    return invalid(
                        "replacement resource kind does not match its base-snapshot selector",
                    );
                }
                let built = build_resource(package, resource, old.as_ref(), &reserved_paths)?;
                reserve_built_paths(&built, &mut reserved_paths)?;
                replacement_owners.insert(owner);
                splices.push((span, built.element_xml.clone()));
                additions.extend(built.additions);
                directories.extend(built.directories);
            },
            EmbeddedResourceChange::Remove(selector) => {
                if !selected_targets.insert(*selector) {
                    return invalid(
                        "embedded-resource batch selects one base owner more than once",
                    );
                }
                let (span, owner, _old, _) = batch_selection(
                    *selector,
                    &object_spans,
                    &object_owners,
                    &object_locations,
                    &image_spans,
                    &image_owners,
                    &image_locations,
                )?;
                removals.push(PlannedRemoval { span, owner });
            },
        }
    }

    plan_owner_removals(
        content,
        &frames,
        &replacement_owners,
        removals,
        &mut splices,
    )?;

    splices.sort_by_key(|(span, _)| span.start);
    for pair in splices.windows(2) {
        if pair[0].0.end > pair[1].0.start {
            return invalid("embedded-resource batch selects overlapping resource owners");
        }
    }
    let mut updated = content.to_owned();
    for (span, replacement) in splices.into_iter().rev() {
        updated = splice(&updated, span.start, span.end, &replacement)?;
    }
    if !appended_frames.is_empty() {
        updated = insert_at_host(&updated, EmbeddedChartHost::Text, &appended_frames)?;
    }

    remove_exact_additions(package, &mut additions)?;
    remove_exact_directories(package, &mut directories)?;
    if updated == content && additions.is_empty() && directories.is_empty() {
        return Ok((package.as_bytes().to_vec(), added_indices));
    }
    let bytes = rebuild_package(
        package,
        &updated,
        additions,
        directories,
        Vec::<String>::new(),
        Vec::<String>::new(),
    )?;
    Ok((bytes, added_indices))
}

pub(crate) fn validate_batch_limits(changes: &[EmbeddedResourceChange]) -> Result<()> {
    if changes.len() > MAX_BATCH_CHANGES {
        return invalid(format!(
            "embedded-resource batch exceeds {MAX_BATCH_CHANGES} changes"
        ));
    }
    let mut total_bytes = 0usize;
    let mut total_files = 0usize;
    for change in changes {
        let resource = match change {
            EmbeddedResourceChange::Add(resource)
            | EmbeddedResourceChange::Replace { resource, .. } => resource,
            EmbeddedResourceChange::Remove(_) => continue,
        };
        for value in [
            resource.frame_name.as_deref(),
            resource.xml_id.as_deref(),
            resource.class_id.as_deref(),
        ]
        .into_iter()
        .flatten()
        {
            total_bytes = batch_byte_sum(total_bytes, value.len())?;
        }
        let (payload_bytes, files) = match &resource.source {
            EmbeddedResourceSource::Linked { href } => (href.len(), 0),
            EmbeddedResourceSource::PackageFile {
                bytes,
                media_type,
                preferred_path,
            } => {
                total_bytes = batch_byte_sum(total_bytes, media_type.len())?;
                if let Some(path) = preferred_path {
                    total_bytes = batch_byte_sum(total_bytes, path.len())?;
                }
                (bytes.len(), 1)
            },
            EmbeddedResourceSource::InlineBinary { bytes, media_type } => {
                if let Some(media_type) = media_type {
                    total_bytes = batch_byte_sum(total_bytes, media_type.len())?;
                }
                (bytes.len(), 1)
            },
            EmbeddedResourceSource::PackageSubdocument {
                files,
                media_type,
                preferred_root,
            } => {
                total_bytes = batch_byte_sum(total_bytes, media_type.len())?;
                if let Some(root) = preferred_root {
                    total_bytes = batch_byte_sum(total_bytes, root.len())?;
                }
                let mut bytes = 0usize;
                for file in files {
                    bytes = bytes.checked_add(file.bytes.len()).ok_or_else(|| {
                        Error::InvalidFormat(
                            "embedded-resource batch byte count overflow".to_string(),
                        )
                    })?;
                    bytes = batch_byte_sum(bytes, file.path.len())?;
                    bytes = batch_byte_sum(bytes, file.media_type.len())?;
                }
                (bytes, files.len())
            },
            EmbeddedResourceSource::InlineXml { xml, .. } => (xml.len(), 0),
        };
        total_bytes = batch_byte_sum(total_bytes, payload_bytes)?;
        total_files = total_files.checked_add(files).ok_or_else(|| {
            Error::InvalidFormat("embedded-resource batch file count overflow".to_string())
        })?;
        if total_bytes > MAX_BYTES {
            return invalid("embedded-resource batch payloads exceed the aggregate size limit");
        }
        if total_files > MAX_FILES {
            return invalid("embedded-resource batch exceeds the aggregate package-file limit");
        }
    }
    Ok(())
}

fn batch_byte_sum(current: usize, added: usize) -> Result<usize> {
    current.checked_add(added).ok_or_else(|| {
        Error::InvalidFormat("embedded-resource batch byte count overflow".to_string())
    })
}

fn batch_selection(
    selector: EmbeddedResourceSelector,
    object_spans: &[ObjectSpan],
    object_owners: &[usize],
    object_locations: &[Option<StoredLocation>],
    image_spans: &[ObjectSpan],
    image_owners: &[usize],
    image_locations: &[Option<StoredLocation>],
) -> Result<(ObjectSpan, usize, Option<StoredLocation>, ResourceTarget)> {
    let (position, spans, owners, locations, target) = match selector {
        EmbeddedResourceSelector::Object(position) => (
            position.get(),
            object_spans,
            object_owners,
            object_locations,
            ResourceTarget::Object,
        ),
        EmbeddedResourceSelector::Image(position) => (
            position.get(),
            image_spans,
            image_owners,
            image_locations,
            ResourceTarget::Image,
        ),
    };
    let span = spans
        .get(position)
        .cloned()
        .ok_or_else(|| bounds(position, spans.len()))?;
    let location = locations.get(position).cloned().ok_or_else(|| {
        Error::InvalidFormat("embedded-resource span/location scan disagreement".to_string())
    })?;
    let owner = owners.get(position).copied().ok_or_else(|| {
        Error::InvalidFormat("embedded-resource owner/span scan disagreement".to_string())
    })?;
    Ok((span, owner, location, target))
}

fn resource_owner_indices(frames: &[FrameOwner], resources: &[ObjectSpan]) -> Result<Vec<usize>> {
    let mut owners = Vec::new();
    owners
        .try_reserve_exact(resources.len())
        .map_err(|source| Error::Allocation {
            resource: "embedded-resource owner spans",
            source,
        })?;
    for resource in resources {
        let (owner, _) = frames
            .iter()
            .enumerate()
            .filter(|(_, frame)| {
                frame.span.start <= resource.start && resource.end <= frame.span.end
            })
            .min_by_key(|(_, frame)| frame.span.end - frame.span.start)
            .ok_or_else(|| {
                Error::InvalidFormat(
                    "selected embedded resource has no supported draw:frame owner".to_string(),
                )
            })?;
        owners.push(owner);
    }
    Ok(owners)
}

fn locate_frames(xml: &str) -> Result<Vec<FrameOwner>> {
    struct Active {
        depth: usize,
        start: usize,
        inner_start: usize,
    }
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut active = Vec::new();
    let mut spans = Vec::new();
    loop {
        let start = usize::try_from(reader.buffer_position())
            .map_err(|_error| Error::InvalidFormat("XML position overflow".to_string()))?;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid frame-owner XML: {error}")))?;
        let is_draw =
            matches!(namespace, ResolveResult::Bound(Namespace(value)) if value == DRAW_NS);
        let end = usize::try_from(reader.buffer_position())
            .map_err(|_error| Error::InvalidFormat("XML position overflow".to_string()))?;
        match event {
            Event::Start(element) => {
                if is_draw && element.local_name().as_ref() == b"frame" {
                    active.push(Active {
                        depth,
                        start,
                        inner_start: end,
                    });
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| Error::InvalidFormat("XML depth overflow".to_string()))?;
            },
            Event::Empty(element) if is_draw && element.local_name().as_ref() == b"frame" => {
                spans.push(FrameOwner {
                    span: ObjectSpan {
                        start,
                        end,
                        inline_payload: None,
                    },
                    inner_start: end,
                    inner_end: end,
                });
            },
            Event::End(element) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| Error::InvalidFormat("XML depth underflow".to_string()))?;
                if is_draw && element.local_name().as_ref() == b"frame" {
                    let owner = active.pop().ok_or_else(|| {
                        Error::InvalidFormat("frame-owner XML stack underflow".to_string())
                    })?;
                    if owner.depth != depth {
                        return invalid("crossed draw:frame ownership is unsupported");
                    }
                    spans.push(FrameOwner {
                        span: ObjectSpan {
                            start: owner.start,
                            end,
                            inline_payload: None,
                        },
                        inner_start: owner.inner_start,
                        inner_end: start,
                    });
                }
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if !active.is_empty() {
        return invalid("unterminated draw:frame owner");
    }
    Ok(spans)
}

fn removed_resource_counts(changes: &[EmbeddedResourceChange]) -> Result<(usize, usize)> {
    let mut selected = HashSet::new();
    let mut objects = 0usize;
    let mut images = 0usize;
    for selector in changes.iter().filter_map(|change| match change {
        EmbeddedResourceChange::Remove(selector) => Some(*selector),
        _ => None,
    }) {
        if !selected.insert(selector) {
            return invalid("embedded-resource batch selects one base owner more than once");
        }
        match selector {
            EmbeddedResourceSelector::Object(_) => objects = objects.saturating_add(1),
            EmbeddedResourceSelector::Image(_) => images = images.saturating_add(1),
        }
    }
    Ok((objects, images))
}

fn plan_owner_removals(
    content: &str,
    frames: &[FrameOwner],
    replacement_owners: &HashSet<usize>,
    removals: Vec<PlannedRemoval>,
    splices: &mut Vec<(ObjectSpan, String)>,
) -> Result<()> {
    let mut by_owner = BTreeMap::<usize, Vec<ObjectSpan>>::new();
    for removal in removals {
        by_owner
            .entry(removal.owner)
            .or_default()
            .push(removal.span);
    }
    for (owner_index, mut members) in by_owner {
        let owner = frames.get(owner_index).ok_or_else(|| {
            Error::InvalidFormat("embedded-resource frame owner index is invalid".to_string())
        })?;
        members.sort_by_key(|span| span.start);
        for pair in members.windows(2) {
            if pair[0].end > pair[1].start {
                return invalid("embedded-resource batch selects overlapping resource members");
            }
        }
        if !replacement_owners.contains(&owner_index)
            && frame_is_empty_after(content, owner, &members)?
        {
            splices.push((owner.span.clone(), String::new()));
        } else {
            splices.extend(members.into_iter().map(|span| (span, String::new())));
        }
    }
    Ok(())
}

fn frame_is_empty_after(
    content: &str,
    owner: &FrameOwner,
    removals: &[ObjectSpan],
) -> Result<bool> {
    if owner.inner_start > owner.inner_end || owner.inner_end > content.len() {
        return invalid("embedded-resource frame inner range is invalid");
    }
    let mut cursor = owner.inner_start;
    for removal in removals {
        if removal.start < cursor || removal.end > owner.inner_end {
            return invalid("embedded-resource member escapes its selected frame owner");
        }
        if !content[cursor..removal.start].trim().is_empty() {
            return Ok(false);
        }
        cursor = removal.end;
    }
    Ok(content[cursor..owner.inner_end].trim().is_empty())
}

fn reserve_built_paths(built: &BuiltResource, reserved_paths: &mut HashSet<String>) -> Result<()> {
    for addition in &built.additions {
        if reserved_paths
            .iter()
            .any(|path| package_paths_conflict(path, &addition.path))
            || !reserved_paths.insert(addition.path.clone())
        {
            return invalid(format!(
                "embedded-resource batch package path '{}' collides",
                addition.path
            ));
        }
    }
    Ok(())
}

fn remove_exact_additions(package: &OwnedPackage, additions: &mut Vec<Addition>) -> Result<()> {
    let archive = package.package()?;
    additions.retain(|addition| {
        archive.get_file(&addition.path).ok().is_none_or(|bytes| {
            bytes != addition.bytes
                || archive.manifest().get_media_type(&addition.path)
                    != Some(addition.media_type.as_str())
        })
    });
    Ok(())
}

fn remove_exact_directories(
    package: &OwnedPackage,
    directories: &mut Vec<(String, String)>,
) -> Result<()> {
    let archive = package.package()?;
    directories.retain(|(path, media_type)| {
        archive.manifest().get_media_type(path) != Some(media_type.as_str())
    });
    Ok(())
}

fn build_resource(
    package: &OwnedPackage,
    resource: &EmbeddedResource,
    replacing: Option<&StoredLocation>,
    reserved_paths: &HashSet<String>,
) -> Result<BuiltResource> {
    validate_metadata(resource)?;
    let mut additions = Vec::new();
    let mut directories = Vec::new();
    let element_xml = match &resource.source {
        EmbeddedResourceSource::Linked { href } => {
            if !is_linked_href(href) {
                return invalid(
                    "linked resources must use an external, fragment, or otherwise inert target",
                );
            }
            href_element(resource, href, None)?
        },
        EmbeddedResourceSource::PackageFile {
            bytes,
            media_type,
            preferred_path,
        } => {
            validate_payload(bytes.len())?;
            validate_media_type(resource.kind, media_type, false)?;
            let path = match preferred_path {
                Some(path) => validate_available_path(package, path, replacing, reserved_paths)?,
                None => unused_file_path(package, resource.kind, media_type, reserved_paths)?,
            };
            additions.push(Addition {
                path: path.clone(),
                bytes: bytes.clone(),
                media_type: media_type.clone(),
            });
            href_element(resource, &path, Some(media_type))?
        },
        EmbeddedResourceSource::PackageSubdocument {
            files,
            media_type,
            preferred_root,
        } => {
            if resource.kind != EmbeddedResourceKind::Object {
                return invalid("only draw:object supports package subdocuments");
            }
            validate_media_type(resource.kind, media_type, true)?;
            if files.is_empty() || files.len() > MAX_FILES {
                return invalid("embedded subdocument file count is outside the allowed range");
            }
            let root = match preferred_root {
                Some(root) => validate_available_root(package, root, replacing, reserved_paths)?,
                None => unused_root(package, reserved_paths)?,
            };
            let mut names = HashSet::new();
            let mut total = 0usize;
            let mut has_content = false;
            for file in files {
                let relative = validate_relative_path(&file.path)?;
                if !names.insert(relative.clone()) {
                    return invalid(format!("duplicate embedded subdocument path '{relative}'"));
                }
                if relative == constants::ODF_CONTENT {
                    has_content = true;
                    if file.media_type != "text/xml" {
                        return invalid(
                            "embedded content.xml must have manifest media type text/xml",
                        );
                    }
                    validate_xml_document(&file.bytes, "embedded content.xml")?;
                }
                validate_manifest_media_type(&file.media_type)?;
                total = total.checked_add(file.bytes.len()).ok_or_else(|| {
                    Error::InvalidFormat("embedded resource size overflow".to_string())
                })?;
                if total > MAX_BYTES {
                    return invalid("embedded subdocument exceeds size limit");
                }
                additions.push(Addition {
                    path: format!("{root}{relative}"),
                    bytes: file.bytes.clone(),
                    media_type: file.media_type.clone(),
                });
            }
            if !has_content {
                return invalid("embedded subdocument is missing content.xml");
            }
            directories.push((root.clone(), media_type.clone()));
            href_element(resource, root.trim_end_matches('/'), None)?
        },
        EmbeddedResourceSource::InlineXml { root, xml } => {
            if resource.kind != EmbeddedResourceKind::Object {
                return invalid("inline XML is only valid for draw:object");
            }
            validate_inline_xml(*root, xml)?;
            element(resource, xml, None)?
        },
        EmbeddedResourceSource::InlineBinary { bytes, media_type } => {
            if resource.kind == EmbeddedResourceKind::Object {
                return invalid("draw:object inline payloads must be XML");
            }
            validate_payload(bytes.len())?;
            if let Some(media_type) = media_type {
                validate_media_type(resource.kind, media_type, false)?;
            }
            let encoded = BASE64_STANDARD.encode(bytes);
            element(
                resource,
                &format!("<office:binary-data>{encoded}</office:binary-data>"),
                media_type.as_deref(),
            )?
        },
    };
    let mut frame = format!(
        "<draw:frame xmlns:draw=\"{DRAW_URI}\" xmlns:office=\"{OFFICE_URI}\" xmlns:xlink=\"{XLINK_URI}\""
    );
    attribute(&mut frame, "draw:name", resource.frame_name.as_deref())?;
    frame.push('>');
    frame.push_str(&element_xml);
    frame.push_str("</draw:frame>");
    Ok(BuiltResource {
        element_xml,
        frame_xml: frame,
        additions,
        directories,
    })
}

fn href_element(
    resource: &EmbeddedResource,
    href: &str,
    media_type: Option<&str>,
) -> Result<String> {
    let mut attrs = String::new();
    attribute(&mut attrs, "xlink:href", Some(href))?;
    attrs.push_str(" xlink:type=\"simple\" xlink:show=\"embed\" xlink:actuate=\"onLoad\"");
    element(resource, "", media_type).map(|element| {
        let point = element.find('>').unwrap_or(element.len());
        format!("{}{}{}", &element[..point], attrs, &element[point..])
    })
}

fn element(resource: &EmbeddedResource, payload: &str, media_type: Option<&str>) -> Result<String> {
    let local = match resource.kind {
        EmbeddedResourceKind::Object => "object",
        EmbeddedResourceKind::ObjectOle => "object-ole",
        EmbeddedResourceKind::Image => "image",
    };
    let mut out = format!("<draw:{local}");
    attribute(&mut out, "xml:id", resource.xml_id.as_deref())?;
    attribute(&mut out, "draw:class-id", resource.class_id.as_deref())?;
    if resource.kind == EmbeddedResourceKind::Image {
        attribute(&mut out, "draw:mime-type", media_type)?;
    }
    out.push('>');
    out.push_str(payload);
    out.push_str("</draw:");
    out.push_str(local);
    out.push('>');
    Ok(out)
}

fn validate_metadata(resource: &EmbeddedResource) -> Result<()> {
    if resource.kind != EmbeddedResourceKind::ObjectOle && resource.class_id.is_some() {
        return invalid("draw:class-id is only valid for ObjectOle resources");
    }
    for value in [
        resource.frame_name.as_deref(),
        resource.xml_id.as_deref(),
        resource.class_id.as_deref(),
    ]
    .into_iter()
    .flatten()
    {
        validate_xml_chars(value)?;
    }
    Ok(())
}

fn validate_inline_xml(root: Root, xml: &str) -> Result<()> {
    if xml.len() > MAX_INLINE_BYTES {
        return invalid("inline XML exceeds size limit");
    }
    if xml.contains("<!DOCTYPE")
        || xml.contains("<!ENTITY")
        || xml.contains("<?")
        || xml.contains("office:scripts")
        || xml.contains("script:event-listener")
        || xml.contains("urn:oasis:names:tc:opendocument:xmlns:script")
    {
        return invalid("inline XML contains active or prohibited markup");
    }
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid inline XML: {error}")))?;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                let expected = match root {
                    Root::OpenDocument => b"document".as_slice(),
                    Root::MathMl => b"math".as_slice(),
                    _ => return invalid("unsupported inline XML root"),
                };
                if element.local_name().as_ref() != expected {
                    return invalid("inline XML root does not match its declared kind");
                }
                if root == Root::OpenDocument {
                    let valid_ns = matches!(namespace, ResolveResult::Bound(Namespace(value)) if value == OFFICE_NS);
                    if !valid_ns {
                        return invalid("inline OpenDocument root has the wrong namespace");
                    }
                    let mut mime = None;
                    for attr in element.attributes() {
                        let attr = attr.map_err(|error| {
                            Error::InvalidFormat(format!("invalid inline XML attribute: {error}"))
                        })?;
                        if attr.key.local_name().as_ref() == b"mimetype" {
                            mime = Some(
                                attr.decoded_and_normalized_value(
                                    XmlVersion::Implicit1_0,
                                    reader.decoder(),
                                )
                                .map_err(|error| {
                                    Error::InvalidFormat(format!(
                                        "invalid inline XML attribute: {error}"
                                    ))
                                })?
                                .into_owned(),
                            );
                        }
                    }
                    let mime = mime.ok_or_else(|| {
                        Error::InvalidFormat(
                            "inline OpenDocument root is missing office:mimetype".to_string(),
                        )
                    })?;
                    if !constants::is_odf_mime_type(&mime) {
                        return invalid("inline OpenDocument has an invalid media type");
                    }
                }
                return Ok(());
            },
            Event::DocType(_) | Event::PI(_) => {
                return invalid("inline XML contains a prohibited declaration");
            },
            Event::Text(value) => {
                let bytes: &[u8] = value.as_ref();
                if !bytes.iter().all(u8::is_ascii_whitespace) {
                    return invalid("invalid content before inline XML root");
                }
            },
            Event::Comment(_) | Event::Decl(_) => {},
            Event::Eof => return invalid("inline XML has no root element"),
            _ => return invalid("invalid content before inline XML root"),
        }
        buffer.clear();
    }
}

fn validate_xml_document(bytes: &[u8], label: &str) -> Result<()> {
    if bytes.len() > MAX_INLINE_BYTES {
        return invalid(format!("{label} exceeds size limit"));
    }
    let text = std::str::from_utf8(bytes)
        .map_err(|_error| Error::InvalidFormat(format!("{label} is not UTF-8")))?;
    if text.contains("<!DOCTYPE") || text.contains("<!ENTITY") {
        return invalid(format!("{label} contains a DTD"));
    }
    let mut reader = NsReader::from_str(text);
    let mut buffer = Vec::new();
    loop {
        match reader
            .read_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid {label}: {error}")))?
        {
            Event::Eof => return Ok(()),
            _ => buffer.clear(),
        }
    }
}

fn validate_media_type(
    kind: EmbeddedResourceKind,
    media_type: &str,
    subdocument: bool,
) -> Result<()> {
    validate_manifest_media_type(media_type)?;
    let lower = media_type.to_ascii_lowercase();
    if lower.contains("javascript")
        || lower.contains("ecmascript")
        || lower.contains("x-executable")
        || lower.contains("x-msdownload")
        || lower.contains("portable-executable")
        || lower.contains("x-sharedlib")
        || lower.contains("x-shellscript")
    {
        return invalid("active or executable embedded media types are prohibited");
    }
    if kind == EmbeddedResourceKind::Image && !lower.starts_with("image/") {
        return invalid("draw:image package payloads require an image media type");
    }
    if subdocument && !constants::is_odf_mime_type(media_type) {
        return invalid("embedded subdocuments require an OpenDocument media type");
    }
    Ok(())
}

fn validate_manifest_media_type(value: &str) -> Result<()> {
    if value.is_empty()
        || value.trim() != value
        || value.chars().any(char::is_control)
        || !value.contains('/')
        || value.contains(';')
    {
        return invalid("invalid manifest media type");
    }
    Ok(())
}

fn validate_payload(len: usize) -> Result<()> {
    if len > MAX_BYTES {
        invalid("embedded payload exceeds size limit")
    } else {
        Ok(())
    }
}

fn validate_relative_path(path: &str) -> Result<String> {
    let path = resolve_package_path(path)?;
    if path.is_empty() || path.ends_with('/') || path.starts_with("META-INF/") || path == "mimetype"
    {
        return invalid("invalid embedded subdocument file path");
    }
    Ok(path)
}

fn validate_available_path(
    package: &OwnedPackage,
    path: &str,
    replacing: Option<&StoredLocation>,
    reserved_paths: &HashSet<String>,
) -> Result<String> {
    let path = resolve_package_path(path)?;
    if path.is_empty() || path.ends_with('/') || protected_path(&path) {
        return invalid("invalid embedded package file path");
    }
    let allowed = matches!(replacing, Some(StoredLocation::File(old)) if old == &path);
    if reserved_paths
        .iter()
        .any(|reserved| package_paths_conflict(reserved, &path))
        || package_namespace_paths(package)?.iter().any(|existing| {
            package_paths_conflict(existing, &path) && !(allowed && existing == &path)
        })
    {
        return invalid(format!("package path '{path}' already exists"));
    }
    Ok(path)
}

fn validate_available_root(
    package: &OwnedPackage,
    root: &str,
    replacing: Option<&StoredLocation>,
    reserved_paths: &HashSet<String>,
) -> Result<String> {
    let mut root = resolve_package_path(root)?;
    if root.is_empty() || protected_path(&root) {
        return invalid("invalid embedded subdocument root");
    }
    root.push('/');
    let allowed = matches!(replacing, Some(StoredLocation::Directory(old)) if old == &root);
    if reserved_paths
        .iter()
        .any(|path| package_paths_conflict(path, &root))
        || package_namespace_paths(package)?.iter().any(|path| {
            package_paths_conflict(path, &root) && !(allowed && path.starts_with(&root))
        })
    {
        return invalid(format!("package root '{root}' already exists"));
    }
    Ok(root)
}

fn protected_path(path: &str) -> bool {
    path == "mimetype"
        || path == constants::ODF_CONTENT
        || path == constants::ODF_STYLES
        || path == constants::ODF_META
        || path.starts_with("META-INF/")
}

fn package_paths_conflict(left: &str, right: &str) -> bool {
    let left = left.trim_end_matches('/');
    let right = right.trim_end_matches('/');
    left == right
        || left
            .strip_prefix(right)
            .is_some_and(|suffix| suffix.starts_with('/'))
        || right
            .strip_prefix(left)
            .is_some_and(|suffix| suffix.starts_with('/'))
}

fn package_namespace_paths(package: &OwnedPackage) -> Result<Vec<String>> {
    let archive = package.package()?;
    let mut paths = archive.files()?;
    paths.extend(archive.manifest().entries.keys().cloned());
    paths.sort();
    paths.dedup();
    Ok(paths)
}

fn unused_file_path(
    package: &OwnedPackage,
    kind: EmbeddedResourceKind,
    media_type: &str,
    reserved_paths: &HashSet<String>,
) -> Result<String> {
    let files = package_namespace_paths(package)?;
    let extension = if kind == EmbeddedResourceKind::Image {
        match media_type.to_ascii_lowercase().as_str() {
            "image/png" => "png",
            "image/jpeg" => "jpg",
            "image/gif" => "gif",
            "image/svg+xml" => "svg",
            "image/webp" => "webp",
            _ => "bin",
        }
    } else {
        "bin"
    };
    for index in 1..=100_000usize {
        let path = if kind == EmbeddedResourceKind::Image {
            format!("Pictures/Image_{index}.{extension}")
        } else {
            format!("Object_{index}.{extension}")
        };
        if !reserved_paths
            .iter()
            .any(|reserved| package_paths_conflict(reserved, &path))
            && !files
                .iter()
                .any(|existing| package_paths_conflict(existing, &path))
        {
            return Ok(path);
        }
    }
    invalid("no collision-free embedded package path is available")
}

fn unused_root(package: &OwnedPackage, reserved_paths: &HashSet<String>) -> Result<String> {
    let files = package_namespace_paths(package)?;
    for index in 1..=100_000usize {
        let root = format!("Object_{index}/");
        if !reserved_paths
            .iter()
            .any(|path| package_paths_conflict(path, &root))
            && !files.iter().any(|path| package_paths_conflict(path, &root))
        {
            return Ok(root);
        }
    }
    invalid("no collision-free embedded package root is available")
}

fn selected_spans(
    package: &OwnedPackage,
    content: &str,
    styles: Option<&str>,
    target: ResourceTarget,
) -> Result<Vec<ObjectSpan>> {
    match target {
        ResourceTarget::Object => {
            let all_spans = locate_objects(content)?;
            let objects = scan_objects(package, content, styles)?;
            let content_objects: Vec<_> = objects
                .iter()
                .filter(|object| object.part == Part::Content)
                .collect();
            if all_spans.len() != content_objects.len() {
                return invalid("embedded-object XML scan disagreement");
            }
            Ok(all_spans
                .into_iter()
                .zip(content_objects)
                .filter_map(|(span, object)| {
                    matches!(object.kind, Kind::Object | Kind::ObjectOle).then_some(span)
                })
                .collect())
        },
        ResourceTarget::Image => locate_images(content),
    }
}

fn selected_locations(
    package: &OwnedPackage,
    content: &str,
    styles: Option<&str>,
    target: ResourceTarget,
) -> Result<Vec<Option<StoredLocation>>> {
    match target {
        ResourceTarget::Object => Ok(scan_objects(package, content, styles)?
            .into_iter()
            .filter(|object| object.part == Part::Content)
            .filter(|object| matches!(object.kind, Kind::Object | Kind::ObjectOle))
            .map(|object| match object.source {
                Source::PackageFile { path, .. } => Some(StoredLocation::File(path)),
                Source::PackageSubdocument { root_path, .. } => {
                    Some(StoredLocation::Directory(root_path))
                },
                _ => None,
            })
            .collect()),
        ResourceTarget::Image => Ok(scan_images(package, content, styles)?
            .into_iter()
            .filter(|image| image.part == Part::Content)
            .map(|image| match image.source {
                litchi_odf_common::media::Source::PackagePart { path, .. } => {
                    Some(StoredLocation::File(path))
                },
                _ => None,
            })
            .collect()),
    }
}

fn target_count(
    package: &OwnedPackage,
    content: &str,
    styles: Option<&str>,
    target: ResourceTarget,
) -> Result<usize> {
    Ok(selected_spans(package, content, styles, target)?.len())
}

fn scan_objects(
    package: &OwnedPackage,
    content: &str,
    styles: Option<&str>,
) -> Result<Vec<Object>> {
    let archive = package.package()?;
    scan_package(content, styles, &archive)
}

fn scan_images(
    package: &OwnedPackage,
    content: &str,
    styles: Option<&str>,
) -> Result<Vec<crate::Image>> {
    let archive = package.package()?;
    crate::media::scan_package(content, styles, &archive)
}

fn locate_images(xml: &str) -> Result<Vec<ObjectSpan>> {
    struct Active {
        depth: usize,
        start: usize,
    }
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut active: Option<Active> = None;
    let mut spans = Vec::new();
    loop {
        let start = usize::try_from(reader.buffer_position())
            .map_err(|_error| Error::InvalidFormat("XML position overflow".to_string()))?;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid image host XML: {error}")))?;
        let is_draw =
            matches!(namespace, ResolveResult::Bound(Namespace(value)) if value == DRAW_NS);
        let end = usize::try_from(reader.buffer_position())
            .map_err(|_error| Error::InvalidFormat("XML position overflow".to_string()))?;
        match event {
            Event::Start(element) => {
                if is_draw && element.local_name().as_ref() == b"image" {
                    if active.is_some() {
                        return invalid("nested draw:image elements are not supported");
                    }
                    active = Some(Active { depth, start });
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| Error::InvalidFormat("XML depth overflow".to_string()))?;
            },
            Event::Empty(element) if is_draw && element.local_name().as_ref() == b"image" => spans
                .push(ObjectSpan {
                    start,
                    end,
                    inline_payload: None,
                }),
            Event::End(element) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| Error::InvalidFormat("XML depth underflow".to_string()))?;
                if is_draw
                    && element.local_name().as_ref() == b"image"
                    && active.as_ref().is_some_and(|item| item.depth == depth)
                {
                    let item = active.take().ok_or_else(|| {
                        Error::InvalidFormat(
                            "image host XML lost its active draw:image element".to_string(),
                        )
                    })?;
                    spans.push(ObjectSpan {
                        start: item.start,
                        end,
                        inline_payload: None,
                    });
                }
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if active.is_some() {
        return invalid("unterminated draw:image element");
    }
    Ok(spans)
}

fn target(kind: EmbeddedResourceKind) -> ResourceTarget {
    match kind {
        EmbeddedResourceKind::Image => ResourceTarget::Image,
        _ => ResourceTarget::Object,
    }
}

fn attribute(out: &mut String, name: &str, value: Option<&str>) -> Result<()> {
    let Some(value) = value else {
        return Ok(());
    };
    validate_xml_chars(value)?;
    out.push(' ');
    out.push_str(name);
    out.push_str("=\"");
    for character in value.chars() {
        match character {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '"' => out.push_str("&quot;"),
            _ => out.push(character),
        }
    }
    out.push('"');
    Ok(())
}

fn validate_xml_chars(value: &str) -> Result<()> {
    if value.chars().any(|character| matches!(character as u32, 0..=8 | 11 | 12 | 14..=31 | 0xD800..=0xDFFF | 0xFFFE | 0xFFFF)) {
        invalid("value contains a character forbidden by XML 1.0")
    } else { Ok(()) }
}

fn bounds(index: usize, len: usize) -> Error {
    Error::InvalidFormat(format!(
        "embedded resource index {index} is out of range for {len} resources"
    ))
}

fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::InvalidFormat(message.into()))
}
