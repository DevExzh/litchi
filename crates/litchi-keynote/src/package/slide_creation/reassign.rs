//! Layout reassignment of a live slide (a port of the migration host's
//! `set_slide_layout`, record 0372).
//!
//! The slide keeps its text, notes, builds, transition and ordinary
//! drawables. The selected layout supplies the slide style and template,
//! the title/body placeholder presentation (geometry, style, path) and
//! visibility, and the slide node's template UUID; the slide preview is
//! invalidated and the slide component's external references follow the
//! references its objects now hold. Layout-owned media (images, live
//! video, media placeholders) materialized from the old layout is replaced
//! by the new layout's, matched by content as the host matches it; an
//! ambiguous match is refused with [`Unsupported::LayoutMedia`].

use std::collections::{BTreeSet, HashMap};

use litchi_iwa_core::ArchiveObject;

use super::graph::{self, payload, reference_field, repeated_references};
use super::layout::{self, Media, Record};
use super::{Error, Lift, Unsupported};
use crate::package::structure::metadata::ExternalReference;
use crate::package::structure::wire::{self, Fields, LENGTH_DELIMITED};
use crate::package::structure::workspace::Workspace;
use crate::package::{PLACEHOLDER_MESSAGE_TYPE, SLIDE_MESSAGE_TYPE, SLIDE_NODE_MESSAGE_TYPE};

const SLIDE_STYLE: u32 = 1;
const SLIDE_TITLE_PLACEHOLDER: u32 = 5;
const SLIDE_BODY_PLACEHOLDER: u32 = 6;
const SLIDE_OWNED_DRAWABLES: u32 = 7;
const SLIDE_TEMPLATE: u32 = 17;
const SLIDE_DRAWABLES_Z_ORDER: u32 = 42;
const NODE_TEMPLATE_UUID: u32 = 29;
const IMAGE_MESSAGE_TYPE: u32 = 3_005;
const MOVIE_MESSAGE_TYPE: u32 = 3_007;
/// `KN.PlaceholderArchive.super(ShapeInfo).super(Shape)`: geometry is
/// `super(Drawable).geometry`, then `style` and `pathsource`.
const PLACEHOLDER_SHAPE: [u32; 2] = [1, 1];
const SHAPE_DRAWABLE: u32 = 1;
const DRAWABLE_GEOMETRY: u32 = 1;
const SHAPE_STYLE: u32 = 2;
const SHAPE_PATH_SOURCE: u32 = 3;
const DRAWABLE_SUPER: u32 = 1;
const DRAWABLE_PARENT: u32 = 2;
const DRAWABLE_TITLE: u32 = 10;
const DRAWABLE_CAPTION: u32 = 11;
const IMAGE_MASK: u32 = 5;
const IMAGE_FLAGS: u32 = 7;
const MOVIE_IS_LIVE_VIDEO: u32 = 30;

/// What a reassignment changed.
pub(in crate::package) struct Reassigned {
    /// Whether the slide already used the layout (nothing was staged).
    pub(in crate::package) unchanged: bool,
}

struct PlaceholderPlan {
    current: u64,
    target: Option<u64>,
    current_visible: bool,
    target_visible: bool,
}

fn visible(fields: &Fields, placeholder: u64) -> Result<bool, Error> {
    let owned = repeated_references(fields, SLIDE_OWNED_DRAWABLES)?;
    let z_order = repeated_references(fields, SLIDE_DRAWABLES_Z_ORDER)?;
    let owned = owned.iter().filter(|id| **id == placeholder).count();
    let z_order = z_order.iter().filter(|id| **id == placeholder).count();
    if owned > 1 || z_order > 1 || owned != z_order {
        return Err(Error::InvalidSource);
    }
    Ok(owned == 1)
}

fn plan(current: &Fields, target: &Fields, field: u32) -> Result<Option<PlaceholderPlan>, Error> {
    let target_id = reference_field(target, field)?;
    let Some(current_id) = reference_field(current, field)? else {
        if target_id.is_none() {
            return Ok(None);
        }
        // The live slide lost the placeholder the layout needs.
        return Err(Error::Refused(
            crate::slide::creation::Refusal::UnsupportedTopology,
        ));
    };
    Ok(Some(PlaceholderPlan {
        current: current_id,
        target: target_id,
        current_visible: visible(current, current_id)?,
        target_visible: match target_id {
            Some(id) => visible(target, id)?,
            None => false,
        },
    }))
}

/// Replace one reference in a message header the way the host does: in
/// place in the aggregate list (or removed when `new` is already listed),
/// and in every field attribution that lists `old`.
fn replace_header_reference(
    object: &mut ArchiveObject,
    index: usize,
    old: Option<u64>,
    new: Option<u64>,
) -> Result<(), Error> {
    if old == new {
        return Ok(());
    }
    let info = object
        .archive_info
        .message_infos
        .get_mut(index)
        .ok_or(Error::InvalidSource)?;
    replace_in_list(&mut info.object_references, old, new)?;
    for field in &mut info.field_infos {
        if old.is_some_and(|old| field.object_references.contains(&old)) {
            replace_in_list(&mut field.object_references, old, new)?;
        }
    }
    Ok(())
}

fn replace_in_list(list: &mut Vec<u64>, old: Option<u64>, new: Option<u64>) -> Result<(), Error> {
    let count = old.map_or(0, |old| list.iter().filter(|id| **id == old).count());
    if count > 1 {
        return Err(Error::InvalidSource);
    }
    match (old, new) {
        (Some(old), Some(new)) if count == 1 => {
            let position = list
                .iter()
                .position(|id| *id == old)
                .ok_or(Error::InvalidSource)?;
            if list.contains(&new) {
                list.remove(position);
            } else {
                list[position] = new;
            }
        },
        (Some(old), None) => list.retain(|id| *id != old),
        (_, Some(new)) if !list.contains(&new) => list.push(new),
        _ => {},
    }
    Ok(())
}

/// Replace the unique `message_type` payload of `identifier`, keeping the
/// header, then apply `references` (old, new) pairs to it.
fn rewrite(
    workspace: &mut Workspace<'_>,
    identifier: u64,
    message_type: u32,
    data: Vec<u8>,
    references: &[(Option<u64>, Option<u64>)],
) -> Result<(), Error> {
    let limits = workspace.archive_limits();
    let object = workspace.object_mut(identifier).lift()?;
    let index = graph::message_index(object, message_type)?;
    object
        .replace_message_preserving_header(
            index,
            litchi_iwa_core::RawMessage {
                type_: message_type,
                data,
            },
            limits,
        )
        .map_err(crate::package::structure::StructureError::from)
        .lift()?;
    for (old, new) in references {
        replace_header_reference(object, index, *old, *new)?;
    }
    Ok(())
}

/// `source`'s nested message at `path`, if present.
fn nested(
    source: &[u8],
    path: &[u32],
    wire: litchi_iwa_common::WireLimits,
) -> Result<Option<Vec<u8>>, Error> {
    let mut current = source.to_vec();
    for (depth, number) in path.iter().enumerate() {
        let fields = Fields::parse(&current, wire).lift()?;
        match fields.unique(*number).lift()? {
            Some(field) => current = field.message().lift()?.to_vec(),
            None if depth + 1 == path.len() => return Ok(None),
            None => return Err(Error::InvalidSource),
        }
    }
    Ok(Some(current))
}

/// Set or remove the nested message at `path` (every parent must exist).
fn set_nested(
    payload: &[u8],
    path: &[u32],
    value: Option<&[u8]>,
    wire: litchi_iwa_common::WireLimits,
) -> Result<Vec<u8>, Error> {
    let (first, rest) = path.split_first().ok_or(Error::InvalidSource)?;
    let mut fields = Fields::parse(payload, wire).lift()?;
    if rest.is_empty() {
        match value {
            Some(value) => fields.set_message(*first, value),
            None => {
                fields.remove(*first);
            },
        }
    } else {
        let inner = fields
            .unique(*first)
            .lift()?
            .ok_or(Error::InvalidSource)?
            .message()
            .lift()?
            .to_vec();
        let edited = set_nested(&inner, rest, value, wire)?;
        fields.set_message(*first, &edited);
    }
    Ok(fields.encode())
}

fn style_of(placeholder: &[u8], wire: litchi_iwa_common::WireLimits) -> Result<Option<u64>, Error> {
    let path = [PLACEHOLDER_SHAPE[0], PLACEHOLDER_SHAPE[1], SHAPE_STYLE];
    nested(placeholder, &path, wire)?
        .map(|reference| wire::reference_identifier(&reference, wire).lift())
        .transpose()
}

/// Copy the layout placeholder's geometry, style and path onto the live
/// placeholder (host `patch_placeholder_presentation`).
fn patch_presentation(
    workspace: &mut Workspace<'_>,
    current: u64,
    target: u64,
) -> Result<(), Error> {
    let wire = workspace.wire;
    let source = payload(workspace, current, PLACEHOLDER_MESSAGE_TYPE)?.to_vec();
    let layout = payload(workspace, target, PLACEHOLDER_MESSAGE_TYPE)?.to_vec();
    let geometry = [
        PLACEHOLDER_SHAPE[0],
        PLACEHOLDER_SHAPE[1],
        SHAPE_DRAWABLE,
        DRAWABLE_GEOMETRY,
    ];
    let style = [PLACEHOLDER_SHAPE[0], PLACEHOLDER_SHAPE[1], SHAPE_STYLE];
    let path = [
        PLACEHOLDER_SHAPE[0],
        PLACEHOLDER_SHAPE[1],
        SHAPE_PATH_SOURCE,
    ];
    let mut data = source.clone();
    for slot in [&geometry[..], &style[..], &path[..]] {
        let value = nested(&layout, slot, wire)?;
        data = set_nested(&data, slot, value.as_deref(), wire)?;
    }
    if data == source {
        return Ok(());
    }
    let (old, new) = (style_of(&source, wire)?, style_of(&layout, wire)?);
    rewrite(
        workspace,
        current,
        PLACEHOLDER_MESSAGE_TYPE,
        data,
        &[(old, new)],
    )
}

/// The media roots a slide owns, by kind: images, live videos, media
/// placeholders (template-only movies are rendered by the layout and never
/// materialized).
#[derive(Default)]
struct MediaRoots {
    images: Vec<u64>,
    live_videos: Vec<u64>,
    placeholders: Vec<u64>,
}

impl MediaRoots {
    fn all(&self) -> Vec<u64> {
        self.images
            .iter()
            .chain(&self.live_videos)
            .chain(&self.placeholders)
            .copied()
            .collect()
    }

    fn is_empty(&self) -> bool {
        self.images.is_empty() && self.live_videos.is_empty() && self.placeholders.is_empty()
    }
}

fn is_live_video(workspace: &Workspace<'_>, movie: u64) -> Result<bool, Error> {
    let fields = Fields::parse(
        payload(workspace, movie, MOVIE_MESSAGE_TYPE)?,
        workspace.wire,
    )
    .lift()?;
    Ok(match fields.unique(MOVIE_IS_LIVE_VIDEO).lift()? {
        Some(field) => field.varint().lift()? != 0,
        None => false,
    })
}

fn media_roots(workspace: &Workspace<'_>, slide: u64) -> Result<MediaRoots, Error> {
    let fields = Fields::parse(
        payload(workspace, slide, SLIDE_MESSAGE_TYPE)?,
        workspace.wire,
    )
    .lift()?;
    let mut roots = MediaRoots::default();
    for drawable in repeated_references(&fields, SLIDE_OWNED_DRAWABLES)? {
        match layout::classify(workspace, drawable)? {
            Media::Image => roots.images.push(drawable),
            Media::Movie if is_live_video(workspace, drawable)? => roots.live_videos.push(drawable),
            Media::Movie => roots.placeholders.push(drawable),
            Media::None | Media::TemplateOnly => {},
        }
    }
    Ok(roots)
}

/// A media drawable's payload without its instance references (parent,
/// title, caption) and, for images, its mask and flags: two drawables with
/// equal signatures are the same layout media (host `image_signature`,
/// `live_video_signature`, `movie_placeholder_signature`).
fn signature(workspace: &Workspace<'_>, drawable: u64) -> Result<Vec<u8>, Error> {
    let wire = workspace.wire;
    let object = workspace.require(drawable).lift()?;
    let (message_type, data) = object
        .messages
        .iter()
        .find(|message| matches!(message.type_, IMAGE_MESSAGE_TYPE | MOVIE_MESSAGE_TYPE))
        .map(|message| (message.type_, message.data.clone()))
        .ok_or(Error::InvalidSource)?;
    let mut fields = Fields::parse(&data, wire).lift()?;
    let drawable_fields = fields
        .unique(DRAWABLE_SUPER)
        .lift()?
        .ok_or(Error::InvalidSource)?
        .message()
        .lift()?
        .to_vec();
    let mut inner = Fields::parse(&drawable_fields, wire).lift()?;
    for number in [DRAWABLE_PARENT, DRAWABLE_TITLE, DRAWABLE_CAPTION] {
        inner.remove(number);
    }
    fields.set_message(DRAWABLE_SUPER, &inner.encode());
    if message_type == IMAGE_MESSAGE_TYPE {
        fields.remove(IMAGE_MASK);
        fields.remove(IMAGE_FLAGS);
    }
    let mut signature = message_type.to_le_bytes().to_vec();
    signature.extend(fields.encode());
    Ok(signature)
}

/// Match each layout root to exactly one live candidate of equal signature
/// (host `match_roots`).
fn match_roots(
    workspace: &Workspace<'_>,
    templates: &[u64],
    candidates: &[u64],
) -> Result<Vec<u64>, Error> {
    let mut available = candidates
        .iter()
        .map(|candidate| Ok((*candidate, signature(workspace, *candidate)?)))
        .collect::<Result<Vec<_>, Error>>()?;
    let mut matched = Vec::with_capacity(templates.len());
    for template in templates {
        let expected = signature(workspace, *template)?;
        let positions = available
            .iter()
            .enumerate()
            .filter(|(_, (_, candidate))| *candidate == expected)
            .map(|(position, _)| position)
            .collect::<Vec<_>>();
        let [position] = positions.as_slice() else {
            return Err(Error::Refused(
                crate::slide::creation::Refusal::UnsupportedContent {
                    reason: Unsupported::LayoutMedia,
                },
            ));
        };
        matched.push(available.remove(*position).0);
    }
    Ok(matched)
}

/// The private closure of `roots` inside `member`, without `owner`.
fn media_graph(
    workspace: &Workspace<'_>,
    member: &str,
    roots: &[u64],
    owner: u64,
) -> Result<Vec<u64>, Error> {
    if roots.is_empty() {
        return Ok(Vec::new());
    }
    Ok(
        layout::private_closure(workspace, member, roots.iter().copied(), &[owner])?
            .into_iter()
            .filter(|object| *object != owner)
            .collect(),
    )
}

/// Place the materialized roots in a slide's drawable list the way the host
/// does (`insert_target_media`): before the live placeholder matching the
/// next layout placeholder, else at the end when the layout has a
/// placeholder anchor, else at the front in layout order.
fn insert_target_media(
    identifiers: &mut Vec<u64>,
    target_owned: &[u64],
    anchors: &[(u64, Option<u64>)],
    replacements: &HashMap<u64, u64>,
) -> Result<(), Error> {
    let anchor = |target: u64| {
        anchors
            .iter()
            .find(|(layout, _)| *layout == target)
            .and_then(|(_, live)| *live)
    };
    let has_anchor = target_owned.iter().any(|item| anchor(*item).is_some());
    let mut leading = 0;
    for (index, item) in target_owned.iter().enumerate() {
        let Some(new) = replacements.get(item).copied() else {
            continue;
        };
        if identifiers.contains(&new) {
            return Err(Error::InvalidSource);
        }
        let next = target_owned[index + 1..]
            .iter()
            .find_map(|candidate| anchor(*candidate))
            .filter(|candidate| identifiers.contains(candidate));
        if let Some(next) = next {
            let position = identifiers
                .iter()
                .position(|identifier| *identifier == next)
                .ok_or(Error::InvalidSource)?;
            identifiers.insert(position, new);
        } else if has_anchor {
            identifiers.push(new);
        } else {
            identifiers.insert(leading, new);
            leading += 1;
        }
    }
    Ok(())
}

/// Replace the slide's materialized layout media with the target layout's
/// (host `slide_layout_media::materialize`).
#[allow(clippy::too_many_arguments, reason = "one step of the host port")]
fn materialize_media(
    workspace: &mut Workspace<'_>,
    member: &str,
    slide: u64,
    slide_component: u64,
    old_template: Option<u64>,
    record: &Record,
    current: &Fields,
    target: &Fields,
) -> Result<(), Error> {
    let target_slide = record.slide();
    let target_roots = media_roots(workspace, target_slide)?;
    let old_roots = match old_template {
        Some(template) => media_roots(workspace, template)?,
        None => MediaRoots::default(),
    };
    let candidates = media_roots(workspace, slide)?;
    let live = MediaRoots {
        images: match_roots(workspace, &old_roots.images, &candidates.images)?,
        live_videos: match_roots(workspace, &old_roots.live_videos, &candidates.live_videos)?,
        placeholders: match_roots(workspace, &old_roots.placeholders, &candidates.placeholders)?,
    };
    if target_roots.is_empty() && live.is_empty() {
        return Ok(());
    }
    let target_root_ids = target_roots.all();
    let live_root_ids = live.all();
    let old_graph = media_graph(workspace, member, &live_root_ids, slide)?;
    let source_graph = media_graph(workspace, record.member(), &target_root_ids, target_slide)?;
    let mut remap = graph::allocate_remap(workspace, &source_graph)?;
    remap.insert(target_slide, slide);
    graph::clone_into(workspace, member, &source_graph, &remap)?;
    for root in &target_roots.images {
        layout::materialize_image(workspace, remap[root], slide)?;
    }
    for root in target_roots
        .live_videos
        .iter()
        .chain(&target_roots.placeholders)
    {
        let movie = payload(workspace, remap[root], MOVIE_MESSAGE_TYPE)?;
        if layout::media_parent(movie, workspace)? != Some(slide) {
            return Err(Error::Refused(
                crate::slide::creation::Refusal::UnsupportedContent {
                    reason: Unsupported::LayoutMedia,
                },
            ));
        }
    }

    // The slide's drawable lists.
    let target_owned = repeated_references(target, SLIDE_OWNED_DRAWABLES)?;
    let anchors = [SLIDE_TITLE_PLACEHOLDER, SLIDE_BODY_PLACEHOLDER]
        .into_iter()
        .map(|number| {
            Ok((
                reference_field(target, number)?,
                reference_field(current, number)?,
            ))
        })
        .collect::<Result<Vec<_>, Error>>()?
        .into_iter()
        .filter_map(|(layout, live)| layout.map(|layout| (layout, live)))
        .collect::<Vec<_>>();
    let replacements = target_root_ids
        .iter()
        .map(|root| (*root, remap[root]))
        .collect::<HashMap<_, _>>();
    let wire = workspace.wire;
    let mut fields = Fields::parse(payload(workspace, slide, SLIDE_MESSAGE_TYPE)?, wire).lift()?;
    for number in [SLIDE_OWNED_DRAWABLES, SLIDE_DRAWABLES_Z_ORDER] {
        let mut preserved = HashMap::new();
        let mut identifiers = Vec::new();
        for field in fields.all(number) {
            let raw = field.message().lift()?.to_vec();
            let identifier = wire::reference_identifier(&raw, wire).lift()?;
            if preserved.insert(identifier, raw).is_some() {
                return Err(Error::InvalidSource);
            }
            identifiers.push(identifier);
        }
        for removed in &live_root_ids {
            if identifiers.iter().filter(|id| *id == removed).count() != 1 {
                return Err(Error::InvalidSource);
            }
            identifiers.retain(|id| id != removed);
        }
        insert_target_media(&mut identifiers, &target_owned, &anchors, &replacements)?;
        let values = identifiers
            .iter()
            .map(|id| {
                preserved
                    .get(id)
                    .cloned()
                    .unwrap_or_else(|| wire::reference(*id))
            })
            .collect::<Vec<_>>();
        fields.replace_all(number, LENGTH_DELIMITED, &values);
    }
    graph::rewrite(workspace, slide, SLIDE_MESSAGE_TYPE, fields.encode(), None)?;
    for object in &old_graph {
        workspace.remove_object(*object).lift()?;
    }

    // Registry: UUIDs and data owners of the swapped graphs.
    let target_component = workspace
        .registry
        .component_for_member(record.member())
        .ok_or(Error::InvalidSource)?;
    // In graph (archive) order, as the host registers them.
    let source_uuids = source_graph
        .iter()
        .filter_map(|object| {
            target_component
                .uuid_of(*object)
                .map(|uuid| (*object, uuid))
        })
        .collect::<Vec<_>>();
    let mut fresh = Vec::with_capacity(source_uuids.len());
    for (object, uuid) in source_uuids {
        let target = remap[&object];
        fresh.push((target, workspace.clone_uuid(Some(uuid), target).lift()?));
    }
    let cloned = source_graph
        .iter()
        .map(|object| remap[object])
        .collect::<Vec<_>>();
    let new_data = data_owned(workspace, &cloned)?;
    let component = workspace
        .registry
        .component_mut(slide_component)
        .ok_or(Error::InvalidSource)?;
    component.remove_object_uuids(&old_graph);
    for (object, uuid) in fresh {
        component.add_object_uuid(object, uuid);
    }
    for object in &old_graph {
        component.remove_data_owner(*object);
    }
    for (data, object) in new_data {
        component.add_data_owner(data, object);
    }
    Ok(())
}

/// The (data, object) pairs the headers of `objects` list.
fn data_owned(workspace: &Workspace<'_>, objects: &[u64]) -> Result<Vec<(u64, u64)>, Error> {
    let mut owned = Vec::new();
    for object in objects {
        let archive_object = workspace.require(*object).lift()?;
        for info in &archive_object.archive_info.message_infos {
            for data in info.data_references.iter().chain(
                info.field_infos
                    .iter()
                    .flat_map(|field| field.data_references.iter()),
            ) {
                owned.push((*data, *object));
            }
        }
    }
    Ok(owned)
}

/// Every object the component `member` references outside itself (archive
/// headers).
fn external_objects(workspace: &Workspace<'_>, member: &str) -> Result<BTreeSet<u64>, Error> {
    let objects = workspace.member_objects(member);
    let internal = objects.iter().copied().collect::<BTreeSet<_>>();
    let mut found = BTreeSet::new();
    for identifier in &objects {
        let object = workspace.require(*identifier).lift()?;
        for info in &object.archive_info.message_infos {
            for target in info.object_references.iter().chain(
                info.field_infos
                    .iter()
                    .flat_map(|f| f.object_references.iter()),
            ) {
                if !internal.contains(target) {
                    found.insert(*target);
                }
            }
        }
    }
    Ok(found)
}

/// The component whose object-UUID map registers `object`, if exactly one.
fn uuid_component(workspace: &Workspace<'_>, object: u64) -> Result<Option<u64>, Error> {
    let mut found = None;
    for component in workspace.registry.components() {
        if component.uuid_of(object).is_some() {
            if found.is_some() {
                return Err(Error::InvalidSource);
            }
            found = Some(component.identifier);
        }
    }
    Ok(found)
}

fn component_of_object(workspace: &Workspace<'_>, object: u64) -> Result<u64, Error> {
    let member = workspace.component_of(object).ok_or(Error::InvalidSource)?;
    Ok(workspace
        .registry
        .component_for_member(&member)
        .ok_or(Error::InvalidSource)?
        .identifier)
}

/// Reassign the live slide owned by `node` to the layout `record`.
pub(in crate::package) fn stage(
    workspace: &mut Workspace<'_>,
    node: u64,
    record: &Record,
) -> Result<Reassigned, Error> {
    let wire = workspace.wire;
    let slide = graph::node_slide(workspace, node)?;
    let target_slide = record.slide();
    let current = Fields::parse(payload(workspace, slide, SLIDE_MESSAGE_TYPE)?, wire).lift()?;
    let old_template = reference_field(&current, SLIDE_TEMPLATE)?;
    if old_template == Some(target_slide) {
        return Ok(Reassigned { unchanged: true });
    }
    let target =
        Fields::parse(payload(workspace, target_slide, SLIDE_MESSAGE_TYPE)?, wire).lift()?;
    let member = workspace.component_of(slide).ok_or(Error::InvalidSource)?;
    let slide_component = component_of_object(workspace, slide)?;
    let old_external = external_objects(workspace, &member)?;
    let old_template_component = old_template
        .map(|template| component_of_object(workspace, template))
        .transpose()?;
    let title = plan(&current, &target, SLIDE_TITLE_PLACEHOLDER)?;
    let body = plan(&current, &target, SLIDE_BODY_PLACEHOLDER)?;

    // Placeholder presentation.
    for plan in [&title, &body].into_iter().flatten() {
        if let Some(target) = plan.target {
            patch_presentation(workspace, plan.current, target)?;
        }
    }

    // Slide style and template (host `patch_slide_relationship`), then the
    // placeholder visibility (host `placeholder_ownership::patch`).
    let old_style = reference_field(&current, SLIDE_STYLE)?.ok_or(Error::InvalidSource)?;
    let target_style_raw = target
        .unique(SLIDE_STYLE)
        .lift()?
        .ok_or(Error::InvalidSource)?
        .message()
        .lift()?
        .to_vec();
    let new_style = wire::reference_identifier(&target_style_raw, wire).lift()?;
    let mut fields = current.clone();
    fields.set_message(SLIDE_STYLE, &target_style_raw);
    fields.set_message(SLIDE_TEMPLATE, &wire::reference(target_slide));
    for plan in [&title, &body].into_iter().flatten() {
        if plan.current_visible == plan.target_visible {
            continue;
        }
        let raw = wire::reference(plan.current);
        for number in [SLIDE_OWNED_DRAWABLES, SLIDE_DRAWABLES_Z_ORDER] {
            let mut values = fields
                .all(number)
                .map(|field| field.message().map(<[u8]>::to_vec))
                .collect::<Result<Vec<_>, _>>()
                .lift()?;
            if plan.target_visible {
                values.push(raw.clone());
            } else {
                let mut kept = Vec::with_capacity(values.len());
                for value in values {
                    if wire::reference_identifier(&value, wire).lift()? != plan.current {
                        kept.push(value);
                    }
                }
                values = kept;
            }
            fields.replace_all(number, LENGTH_DELIMITED, &values);
        }
    }
    rewrite(
        workspace,
        slide,
        SLIDE_MESSAGE_TYPE,
        fields.encode(),
        &[
            (Some(old_style), Some(new_style)),
            (old_template, Some(target_slide)),
        ],
    )?;

    materialize_media(
        workspace,
        &member,
        slide,
        slide_component,
        old_template,
        record,
        &current,
        &target,
    )?;

    // Slide-node template UUID, then the slide preview.
    let node_payload = payload(workspace, node, SLIDE_NODE_MESSAGE_TYPE)?.to_vec();
    let layout_node = payload(workspace, record.node(), SLIDE_NODE_MESSAGE_TYPE)?.to_vec();
    let uuid = nested(&layout_node, &[NODE_TEMPLATE_UUID], wire)?;
    let patched = set_nested(&node_payload, &[NODE_TEMPLATE_UUID], uuid.as_deref(), wire)?;
    if patched != node_payload {
        rewrite(workspace, node, SLIDE_NODE_MESSAGE_TYPE, patched, &[])?;
    }
    let archive_limits = workspace.archive_limits();
    let object = workspace.object_mut(node).lift()?;
    crate::package::slide_preview::invalidate(object, archive_limits, wire)
        .map_err(|_error| Error::InvalidSource)?;

    // External references follow the slide component's references (host
    // `reconcile_slide_layout_dependencies`).
    let target_component = component_of_object(workspace, target_slide)?;
    let final_external = external_objects(workspace, &member)?;
    let mut final_components = BTreeSet::new();
    let mut additions = Vec::new();
    for object in &final_external {
        if *object == target_slide {
            continue;
        }
        let Some(component) = uuid_component(workspace, *object)? else {
            continue;
        };
        if component == slide_component {
            continue;
        }
        final_components.insert(component);
        additions.push(ExternalReference {
            component,
            object: Some(*object),
            is_weak: None,
        });
    }
    let mut removals = Vec::new();
    for object in old_external.difference(&final_external) {
        if let Some(component) = uuid_component(workspace, *object)?
            && component != slide_component
        {
            removals.push((component, *object));
        }
    }
    let registry_component = workspace
        .registry
        .component_mut(slide_component)
        .ok_or(Error::InvalidSource)?;
    let mut references = registry_component.external_references().to_vec();
    for addition in additions {
        if !references.contains(&addition) {
            references.push(addition);
        }
    }
    references.retain(|reference| {
        !removals.iter().any(|(component, object)| {
            reference.component == *component && reference.object == Some(*object)
        })
    });
    if target_component != slide_component {
        let link = ExternalReference {
            component: target_component,
            object: None,
            is_weak: None,
        };
        if !references.contains(&link) {
            references.push(link);
        }
    }
    if let Some(old) = old_template_component
        && old != slide_component
        && old != target_component
        && !final_components.contains(&old)
    {
        references.retain(|reference| {
            !(reference.component == old
                && reference.object.is_none()
                && reference.is_weak.is_none())
        });
    }
    registry_component.set_external_references(references);
    Ok(Reassigned { unchanged: false })
}
