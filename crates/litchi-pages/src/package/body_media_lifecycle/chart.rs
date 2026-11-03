//! Body chart creation (record 0418; the host's `add_body_chart`): the
//! shared `litchi_iwa_style::chart_lifecycle::author` graph with the Pages
//! profile in the body component, a body attachment and U+FFFC at the
//! anchor, the z-order (created on first use), the styles in the document
//! stylesheet and the preset in the theme's chart presets, with the host's
//! cross-component references and UUIDs.
//!
//! Duplication and removal (record 0426; the host's `duplicate_body_chart`
//! and `remove_body_chart`) work on the chart's private graph in the host's
//! discovery order (`litchi_iwa_style::chart_lifecycle::graph`): the
//! drawable, its caption objects and title, its own preset and the styles
//! only it references, and its attachment. A copy clones the graph under
//! consecutive fresh identifiers in that order, moves the frame 12 pt,
//! registers the copied styles and preset, and anchors the copy with a
//! U+FFFC and attachment entry; a removal clears a direct comment, deletes
//! the U+FFFC, z-order entry and graph, and releases the styles, preset,
//! cross-component references, UUIDs, data owners and orphaned data.

use std::collections::HashSet;
use std::num::NonZeroU64;
use std::sync::Arc;

use litchi_iwa_common::WireLimits;
use litchi_iwa_common::chart::data::ChartData;
use litchi_iwa_common::chart::kind::Kind;
use litchi_iwa_common::wire::transform_length_delimited_field;
use litchi_iwa_core::{Archive, RawMessage};
use litchi_iwa_style::chart_lifecycle::author::{self, Ids, Profile, Request};
use litchi_iwa_style::chart_lifecycle::graph as chart_graph;
use litchi_iwa_table::structure::references;
use litchi_iwa_text::TextPosition;

use super::{
    Budget, CreationIds, DOCUMENT_COMPONENT_IDENTIFIER, Diagnostics, Error, METADATA_MESSAGE_TYPE,
    ObjectIndex, Package, Pb, Refusal, STANDARD_VERSION, THEME_MESSAGE_TYPE, append_zorder,
    create_zorder, drawable_payload, editable, graph, insert_anchor, new_object, objects,
    release_orphans, zorder_without,
};
use crate::selector::BodyChartSelector;

const STYLESHEET_MESSAGE_TYPE: u32 = 401;
/// `TSS.ThemeArchive` (field 1 of the Pages theme): its document
/// stylesheet, first paragraph style preset and chart-preset list.
const THEME_STYLESHEET: [u32; 2] = [1, 4];
const THEME_PARAGRAPH_PRESETS: [u32; 3] = [1, 110, 7];
const THEME_CHART_PRESETS: [u32; 3] = [1, 120, 1];
const STYLESHEET_STYLES: [u32; 1] = [1];

/// One staged body chart.
#[derive(Debug, Clone, PartialEq)]
pub(super) struct NewChart {
    pub(super) kind: Kind,
    pub(super) data: ChartData,
    pub(super) position: crate::drawable::Point,
    pub(super) size: crate::drawable::Size,
}

impl NewChart {
    pub(super) fn validate(&self) -> Result<(), Refusal> {
        if !litchi_iwa_style::chart_lifecycle::assignable_kind(self.kind)
            || author::series_count(self.kind, &self.data).is_none()
        {
            return Err(Refusal::InvalidChart);
        }
        if !self.position.x.is_finite()
            || !self.position.y.is_finite()
            || !self.size.width.is_finite()
            || !self.size.height.is_finite()
            || self.size.width <= 0.0
            || self.size.height <= 0.0
        {
            return Err(Refusal::InvalidGeometry);
        }
        Ok(())
    }
}

fn fault(_fault: litchi_iwa_table::structure::Fault) -> Error {
    Error::InvalidSource
}

fn has_path(payload: &[u8], path: &[u32]) -> Result<bool, Error> {
    let mut current = payload;
    for number in path {
        let records = references::records(current, &[*number]).map_err(fault)?;
        match records.as_slice() {
            [] => return Ok(false),
            [inner] => current = inner,
            _ => return Err(Error::InvalidSource),
        }
    }
    Ok(true)
}

fn first(payload: &[u8], path: &[u32]) -> Result<Option<u64>, Error> {
    if !has_path(payload, &path[..path.len() - 1])? {
        return Ok(None);
    }
    Ok(references::identifiers(payload, path)
        .map_err(fault)?
        .into_iter()
        .find(|identifier| *identifier != 0))
}

/// Append `added` to the reference list at `path` of `object`'s message
/// `message_type` in `archive`, keeping the header in step.
fn append(
    archive: &mut Archive,
    object: u64,
    message_type: u32,
    path: &[u32],
    added: &[u64],
    limits: litchi_iwa_core::Limits,
) -> Result<(), Error> {
    let target = archive.object_mut(object).ok_or(Error::InvalidSource)?;
    let index = target
        .messages
        .iter()
        .position(|message| message.type_ == message_type)
        .ok_or(Error::InvalidSource)?;
    let mut data = target.messages[index].data.clone();
    if !has_path(&data, &path[..path.len() - 1])? {
        // The theme's chart-preset list is created empty on first use.
        data = transform_length_delimited_field(&data, path[0], |inner| {
            let mut extended = inner.to_vec();
            extended.extend_from_slice(&[0xc2, 0x07, 0x00]);
            Ok::<_, litchi_iwa_common::Error>(extended)
        })
        .map_err(|_error| Error::InvalidSource)?;
    }
    let mut after = references::identifiers(&data, path).map_err(fault)?;
    if added.iter().any(|identifier| after.contains(identifier)) {
        return Err(Error::InvalidSource);
    }
    after.extend_from_slice(added);
    let data = references::rewrite(&data, path, &after).map_err(fault)?;
    references::replace(target, index, data, &[], added, path, limits).map_err(fault)
}

/// Build the chart at UTF-16 body position `anchor`.
#[allow(
    clippy::too_many_lines,
    reason = "one staged creation: graph, attachment, registrations, registry"
)]
pub(super) fn add(
    source: &Package,
    anchor: TextPosition,
    chart: &NewChart,
) -> Result<(Package, Diagnostics), Error> {
    let anchor = anchor.utf16_index();
    let archive_limits = source
        .state
        .source
        .limits()
        .effective_archive_limits()
        .map_err(graph::map_archive_error)?;
    let metadata = objects::metadata(source)?.ok_or(Error::InvalidSource)?;
    let (body, left_margin, zorder, body_member, theme) = {
        let mut budget = Budget::new(source)?;
        let index = ObjectIndex::new(source, &mut budget)?;
        let body = graph::resolve_body(&index, &mut budget)?;
        let zorder = graph::root_zorder(source, &mut budget)?;
        let member = index.component_name(index.get(body.root.body)?.component_index)?;
        let (theme, _) = graph::root_theme(source, &mut budget)?;
        (
            body.root.body.get(),
            body.root.left_margin,
            zorder,
            member,
            theme.ok_or(Error::InvalidSource)?,
        )
    };
    let (theme_member, theme_payload) = {
        let mut budget = Budget::new(source)?;
        let index = ObjectIndex::new(source, &mut budget)?;
        let entry = index.get(theme)?;
        let (_, message) = graph::unique_message(entry.object, THEME_MESSAGE_TYPE)?;
        (
            index.component_name(entry.component_index)?,
            message.data.clone(),
        )
    };
    let stylesheet = first(&theme_payload, &THEME_STYLESHEET)?.ok_or(Error::InvalidSource)?;
    let paragraph =
        first(&theme_payload, &THEME_PARAGRAPH_PRESETS)?.ok_or(Refusal::MissingThemeStyle)?;
    let (stylesheet_member, paragraph_member) = {
        let mut budget = Budget::new(source)?;
        let index = ObjectIndex::new(source, &mut budget)?;
        let member = |identifier: u64| -> Result<Arc<str>, Error> {
            let identifier = NonZeroU64::new(identifier).ok_or(Error::InvalidSource)?;
            Ok(index.component_name(index.get(identifier)?.component_index)?)
        };
        (member(stylesheet)?, member(paragraph)?)
    };

    let first_identifier = objects::next_identifier(source, Some(&metadata))?;
    let mut next = first_identifier + u64::from(zorder.is_none());
    let ids = Ids::allocate(Profile::Pages, || -> Result<u64, Error> {
        let identifier = next;
        next = next.checked_add(1).ok_or(Error::InvalidSource)?;
        Ok(identifier)
    })?;
    let attachment = ids.last().checked_add(1).ok_or(Error::InvalidSource)?;
    let creation = CreationIds {
        zorder: zorder.is_none().then_some(first_identifier),
        drawable: ids.drawable,
        title: ids.title,
        caption: ids.caption,
        attachment,
    };
    let drawable = drawable_payload(&creation, body, chart.position, chart.size, false);
    let request = Request {
        profile: Profile::Pages,
        kind: chart.kind,
        data: &chart.data,
        drawable: &drawable,
        stylesheet,
        paragraph_style: paragraph,
    };
    let graph_objects =
        author::objects(&ids, &request, archive_limits).map_err(|error| match error {
            litchi_iwa_style::chart_lifecycle::Error::InvalidData => {
                Error::Refused(Refusal::InvalidChart)
            },
            _ => Error::InvalidSource,
        })?;
    let attachment_payload = Pb::default()
        .reference(1, ids.drawable)
        .varint(2, 0)
        .float(3, chart.position.x - left_margin.unwrap_or_default())
        .varint(4, 1)
        .float(5, chart.position.y)
        .0;
    let mut archives: Vec<(Arc<str>, Archive)> = Vec::new();
    {
        let archive = editable(source, &mut archives, &body_member)?;
        if let Some(zorder) = creation.zorder {
            create_zorder(source, archive, zorder, archive_limits)?;
        }
        for object in graph_objects {
            archive
                .insert_object(object, archive_limits)
                .map_err(graph::map_core_error)?;
        }
        archive
            .insert_object(
                new_object(
                    source,
                    attachment,
                    graph::DRAWABLE_ATTACHMENT_MESSAGE_TYPE,
                    attachment_payload,
                    STANDARD_VERSION,
                    vec![ids.drawable],
                    Vec::new(),
                )?,
                archive_limits,
            )
            .map_err(graph::map_core_error)?;
    }
    let styles = ids.styles();
    append(
        editable(source, &mut archives, &stylesheet_member)?,
        stylesheet,
        STYLESHEET_MESSAGE_TYPE,
        &STYLESHEET_STYLES,
        &styles,
        archive_limits,
    )?;
    insert_anchor(source, &mut archives, anchor, attachment, archive_limits)?;
    let zorder_identifier = zorder.map_or(creation.zorder.unwrap_or(0), NonZeroU64::get);
    append_zorder(
        source,
        &mut archives,
        zorder_identifier,
        &body_member,
        ids.drawable,
        archive_limits,
    )?;
    append(
        editable(source, &mut archives, &theme_member)?,
        theme.get(),
        THEME_MESSAGE_TYPE,
        &THEME_CHART_PRESETS,
        &[ids.preset],
        archive_limits,
    )?;

    // Registry: UUIDs of the chart graph, the watermark at the attachment,
    // then the host's cross-component references.
    let owner = metadata
        .component_for_entry(&body_member)?
        .unwrap_or(DOCUMENT_COMPONENT_IDENTIFIER);
    let mut payload = metadata.registered(source, owner, &ids.all(), attachment)?;
    let component = |member: &str| -> Result<u64, Error> {
        metadata
            .component_for_entry(member)?
            .ok_or(Error::InvalidSource)
    };
    let stylesheet_component = component(&stylesheet_member)?;
    let theme_component = component(&theme_member)?;
    let paragraph_component = component(&paragraph_member)?;
    let mut edges = Vec::new();
    if stylesheet_component != owner {
        edges.push((owner, stylesheet_component, stylesheet));
        edges.extend(
            styles
                .iter()
                .map(|style| (stylesheet_component, owner, *style)),
        );
    }
    if theme_component != owner {
        edges.push((theme_component, owner, ids.preset));
    }
    if paragraph_component != owner {
        edges.push((owner, paragraph_component, paragraph));
    }
    for (from, to, object) in edges {
        payload = objects::add_external_reference(&payload, from, to, object)?;
    }
    {
        let archive = editable(source, &mut archives, &metadata.component)?;
        archive
            .object_mut(metadata.object)
            .ok_or(Error::InvalidSource)?
            .replace_message_preserving_header(
                metadata.message_index,
                RawMessage {
                    type_: METADATA_MESSAGE_TYPE,
                    data: payload,
                },
                archive_limits,
            )
            .map_err(graph::map_core_error)?;
    }
    let candidate = graph::publish_changes(source, archives, &[], &[], &[])?;
    if !super::resolve_drawable(
        &candidate,
        crate::selector::DrawableSelector::shape(0),
        NonZeroU64::new(ids.drawable),
    )? {
        return Err(Error::Verification);
    }
    Ok((
        candidate,
        Diagnostics {
            changed: true,
            created_objects: ids.all().len() + 1 + usize::from(creation.zorder.is_some()),
            ..Diagnostics::default()
        },
    ))
}

// ---------------------------------------------------------------------------
// Duplicate and remove (record 0426; the host's `duplicate_body_chart` and
// `remove_body_chart`).

/// `TSCH.ChartDrawableArchive` and `TSCH.ChartStylePreset`.
pub(super) const CHART_MESSAGE_TYPE: u32 = author::CHART_MESSAGE_TYPE;
pub(super) const PRESET_MESSAGE_TYPE: u32 = author::PRESET_MESSAGE_TYPE;
const CAPTION_PLACEMENT_MESSAGE_TYPE: u32 = 634;
const SHAPE_STYLE_MESSAGE_TYPE: u32 = 2_025;
const STORAGE_MESSAGE_TYPE: u32 = 2_001;
/// `TSA.CaptionInfoArchive.child_info_kind` of a caption.
const CAPTION_KIND: i32 = 1;
/// `TSS.ThemeArchive` field of the theme's chart presets.
const THEME_CHART_PRESETS_FIELD: u32 = 120;
/// `TSP.PackageMetadata` components and their data references.
const METADATA_COMPONENTS: u32 = 3;
const COMPONENT_IDENTIFIER: u32 = 1;
const COMPONENT_DATA_REFERENCES: u32 = 7;
const DATA_REFERENCE_DATA: u32 = 1;
const DATA_REFERENCE_OWNERS: u32 = 2;
const OWNER_OBJECT: u32 = 1;
const OWNER_COUNT: u32 = 2;

/// `message` (of type `message_type`, whose header records `recorded`)
/// with its chart-graph references remapped, refusing a private reference
/// outside the message's known reference paths (the host's
/// `remap_references` and `remap_chart_preset_wire`).
pub(super) fn remap_message(
    message_type: u32,
    data: &[u8],
    recorded: &[u64],
    remap: &std::collections::HashMap<u64, u64>,
) -> Result<Vec<u8>, Error> {
    let paths = chart_graph::reference_paths(message_type).ok_or(Error::InvalidSource)?;
    let (output, seen) = chart_graph::remap(data, paths, remap, WireLimits::DEFAULT)
        .map_err(|_error| Error::InvalidSource)?;
    if recorded
        .iter()
        .any(|identifier| remap.contains_key(identifier) && !seen.contains(identifier))
    {
        return Err(Refusal::SharedGraph.into());
    }
    Ok(output)
}

/// One object of a chart graph: the member that stores it, its registry
/// component and its identifier.
#[derive(Debug, Clone)]
struct GraphObject {
    member: Arc<str>,
    component: u64,
    identifier: u64,
}

/// The theme and document stylesheet a chart graph registers with.
struct ThemeContext {
    theme: u64,
    theme_member: Arc<str>,
    theme_component: u64,
    stylesheet: u64,
    stylesheet_member: Arc<str>,
}

/// The private graph of one body chart (the host's `body_chart_graph`).
struct ChartGraph {
    drawable: u64,
    drawable_message_index: usize,
    attachment: u64,
    character_index: u32,
    member: Arc<str>,
    component: u64,
    /// Drawable, caption objects, title, private preset, private styles,
    /// attachment (the host's order).
    objects: Vec<GraphObject>,
    /// The registered styles among `objects`.
    styles: Vec<u64>,
    preset: Option<u64>,
    /// The objects registered with a UUID (in their own component).
    uuids: Vec<u64>,
    commented: bool,
    zorder: NonZeroU64,
    frame: crate::drawable::Geometry,
    theme: ThemeContext,
}

impl ChartGraph {
    fn identifiers(&self) -> Vec<u64> {
        self.objects
            .iter()
            .map(|object| object.identifier)
            .collect()
    }

    /// The graph's members in first-use order, with their component and
    /// objects.
    fn groups(&self) -> Vec<(Arc<str>, u64, Vec<u64>)> {
        let mut groups: Vec<(Arc<str>, u64, Vec<u64>)> = Vec::new();
        for object in &self.objects {
            match groups
                .iter_mut()
                .find(|(member, _, _)| *member == object.member)
            {
                Some((_, _, objects)) => objects.push(object.identifier),
                None => groups.push((
                    Arc::clone(&object.member),
                    object.component,
                    vec![object.identifier],
                )),
            }
        }
        groups
    }
}

fn arrangement_error(error: super::super::body_chart_arrangement::Error) -> Error {
    use super::super::body_chart_arrangement::Error as Arrangement;
    match error {
        Arrangement::Refused(crate::chart::arrangement::Refusal::ChartNotFound { position }) => {
            Refusal::NotFound {
                kind: "chart",
                position,
            }
            .into()
        },
        Arrangement::Limit {
            observed, maximum, ..
        } => Error::Limit {
            kind: super::LimitKind::WireWork,
            observed,
            maximum,
        },
        Arrangement::Allocation { amount } => Error::Allocation { amount },
        _ => Error::InvalidSource,
    }
}

fn theme_context(source: &Package, metadata: &objects::Metadata) -> Result<ThemeContext, Error> {
    let mut budget = Budget::new(source)?;
    let index = ObjectIndex::new(source, &mut budget)?;
    let (theme, _) = graph::root_theme(source, &mut budget)?;
    let theme = theme.ok_or(Error::InvalidSource)?;
    let entry = index.get(theme)?;
    let (_, message) = graph::unique_message(entry.object, THEME_MESSAGE_TYPE)?;
    let stylesheet = first(&message.data, &THEME_STYLESHEET)?.ok_or(Error::InvalidSource)?;
    let theme_member = index.component_name(entry.component_index)?;
    let stylesheet_member = index.component_name(
        index
            .get(NonZeroU64::new(stylesheet).ok_or(Error::InvalidSource)?)?
            .component_index,
    )?;
    Ok(ThemeContext {
        theme: theme.get(),
        theme_component: metadata
            .component_for_entry(&theme_member)?
            .ok_or(Error::InvalidSource)?,
        theme_member,
        stylesheet,
        stylesheet_member,
    })
}

/// The number of chart payloads that reference each object.
fn chart_owner_counts(source: &Package) -> Result<std::collections::HashMap<u64, usize>, Error> {
    let mut counts = std::collections::HashMap::new();
    for component in source.state.source.components().iter() {
        for object in &component.archive().objects {
            for message in &object.messages {
                if message.type_ != CHART_MESSAGE_TYPE {
                    continue;
                }
                let references = chart_graph::references(&message.data, WireLimits::DEFAULT)
                    .map_err(|_error| Error::InvalidSource)?;
                for identifier in references.typed {
                    *counts.entry(identifier).or_insert(0usize) += 1;
                }
            }
        }
    }
    Ok(counts)
}

/// The caption slot's private objects (the host's `drawable_caption_slot`):
/// a stand-in, or the caption, its style when the drawable's member holds
/// it, its storage and its placement.
fn caption_objects(
    index: &ObjectIndex<'_>,
    member: &str,
    drawable: u64,
    caption: u64,
) -> Result<Vec<u64>, Error> {
    use litchi_iwa_protos::pages_movie_caption_codec as codec;
    let in_member = |identifier: u64| -> Result<bool, Error> {
        let located = index.get(NonZeroU64::new(identifier).ok_or(Error::InvalidSource)?)?;
        Ok(index.component_name(located.component_index)?.as_ref() == member)
    };
    let located = index.get(NonZeroU64::new(caption).ok_or(Error::InvalidSource)?)?;
    if !in_member(caption)? {
        return Err(Error::InvalidSource);
    }
    let Some((info_index, info)) =
        graph::unique_optional_message(located.object, super::CAPTION_INFO_MESSAGE_TYPE)?
    else {
        graph::unique_message(located.object, author::STANDIN_MESSAGE_TYPE)?;
        return Ok(vec![caption]);
    };
    graph::validate_message_metadata(located.object, info_index)?;
    let snapshot =
        codec::decode_caption_info(&info.data, codec::DecodeOptions::for_source(&info.data))
            .map_err(|_error| Error::InvalidSource)?;
    let storage = snapshot
        .owned_storage_identifier()
        .filter(|identifier| *identifier != 0)
        .ok_or(Error::InvalidSource)?;
    let style = snapshot
        .style_identifier()
        .filter(|identifier| *identifier != 0)
        .ok_or(Error::InvalidSource)?;
    let placement = snapshot
        .placement_identifier()
        .filter(|identifier| *identifier != 0)
        .ok_or(Error::InvalidSource)?;
    if snapshot.child_info_kind() != Some(CAPTION_KIND)
        || snapshot.parent_identifier() != drawable
        || snapshot.deprecated_storage_identifier() != Some(storage)
        || snapshot.is_text_box() != Some(true)
    {
        return Err(Error::InvalidSource);
    }
    for (identifier, message_type) in [
        (style, SHAPE_STYLE_MESSAGE_TYPE),
        (storage, STORAGE_MESSAGE_TYPE),
        (placement, CAPTION_PLACEMENT_MESSAGE_TYPE),
    ] {
        let object = index.get(NonZeroU64::new(identifier).ok_or(Error::InvalidSource)?)?;
        graph::unique_message(object.object, message_type)?;
    }
    if !in_member(storage)? || !in_member(placement)? {
        return Err(Error::InvalidSource);
    }
    let mut found = vec![caption];
    if in_member(style)? {
        found.push(style);
    }
    found.extend([storage, placement]);
    Ok(found)
}

fn zorder_count(source: &Package, zorder: NonZeroU64, drawable: u64) -> Result<usize, Error> {
    let mut budget = Budget::new(source)?;
    let index = ObjectIndex::new(source, &mut budget)?;
    let (_, message) =
        graph::unique_message(index.get(zorder)?.object, super::ZORDER_MESSAGE_TYPE)?;
    Ok(
        references::identifiers(&message.data, &[super::ZORDER_DRAWABLES_FIELD])
            .map_err(fault)?
            .into_iter()
            .filter(|identifier| *identifier == drawable)
            .count(),
    )
}

#[allow(
    clippy::too_many_lines,
    reason = "one graph proof in the host's discovery order"
)]
fn resolve_chart(
    source: &Package,
    selector: BodyChartSelector,
    metadata: &objects::Metadata,
) -> Result<ChartGraph, Error> {
    use super::super::body_chart_arrangement::{self as arrangement, ArrangementBudget};
    let target = {
        let mut budget = ArrangementBudget::new(source).map_err(arrangement_error)?;
        arrangement::resolve_target(source, selector, &mut budget).map_err(arrangement_error)?
    };
    let payload =
        arrangement::selected_chart_payload(source, &target).map_err(arrangement_error)?;
    let found = chart_graph::references(payload, WireLimits::DEFAULT)
        .map_err(|_error| Error::InvalidSource)?;
    let frame = super::super::body_chart_lifecycle::geometry(payload)
        .map_err(|_error| Error::InvalidSource)?
        .filter(|frame| frame.position.is_some())
        .ok_or(Error::InvalidSource)?;
    let drawable = target.drawable_identifier.get();
    let member = Arc::clone(&target.component_name);
    if found.parent != Some(target.body_identifier.get()) {
        return Err(Error::InvalidSource);
    }
    if found.mediator.is_some() {
        return Err(Refusal::SharedGraph.into());
    }
    let mut budget = Budget::new(source)?;
    let index = ObjectIndex::new(source, &mut budget)?;
    let locate = |identifier: u64| -> Result<(graph::Located<'_>, Arc<str>), Error> {
        let located = index.get(NonZeroU64::new(identifier).ok_or(Error::InvalidSource)?)?;
        let name = index.component_name(located.component_index)?;
        Ok((located, name))
    };
    let caption = found
        .caption
        .filter(|identifier| *identifier != 0)
        .ok_or(Error::InvalidSource)?;
    let title = found
        .title
        .filter(|identifier| *identifier != 0)
        .ok_or(Error::InvalidSource)?;
    let mut identifiers = vec![drawable];
    identifiers.extend(caption_objects(&index, &member, drawable, caption)?);
    identifiers.push(title);
    let mut styles = Vec::new();
    let preset = found.owned_preset;
    if let Some(preset) = preset {
        if found.preset != Some(preset) {
            return Err(Error::InvalidSource);
        }
        let counts = chart_owner_counts(source)?;
        let (located, preset_member) = locate(preset)?;
        let presets = located
            .object
            .messages
            .iter()
            .filter(|message| message.type_ == PRESET_MESSAGE_TYPE)
            .count();
        if preset_member != member || presets != 1 {
            return Err(Error::InvalidSource);
        }
        if counts.get(&preset) != Some(&1) {
            return Err(Refusal::SharedGraph.into());
        }
        identifiers.push(preset);
        let mut seen = std::collections::HashMap::new();
        for style in &found.styles {
            if let Some(kind) = seen.insert(style.identifier, style.message_type) {
                if kind != style.message_type {
                    return Err(Error::InvalidSource);
                }
                continue;
            }
            let (located, style_member) = locate(style.identifier)?;
            let mut messages = located
                .object
                .messages
                .iter()
                .filter(|message| message.type_ == style.message_type);
            let (Some(message), None) = (messages.next(), messages.next()) else {
                return Err(Error::InvalidSource);
            };
            if counts.get(&style.identifier) != Some(&1) {
                continue;
            }
            let registered = chart_graph::registered_style(
                style.message_type,
                &message.data,
                WireLimits::DEFAULT,
            )
            .map_err(|_error| Error::InvalidSource)?;
            if style_member == member || registered {
                identifiers.push(style.identifier);
            }
            if registered {
                styles.push(style.identifier);
            }
        }
    }
    {
        let (located, title_member) = locate(title)?;
        if title_member != member {
            return Err(Error::InvalidSource);
        }
        graph::unique_message(located.object, author::STANDIN_MESSAGE_TYPE)?;
    }
    let attachment = target.attachment_identifier.get();
    identifiers.push(attachment);
    let unique = identifiers.iter().collect::<HashSet<_>>();
    if unique.len() != identifiers.len() {
        return Err(Refusal::SharedGraph.into());
    }
    let component = metadata
        .component_for_entry(&member)?
        .ok_or(Error::InvalidSource)?;
    let mut graph_objects = Vec::with_capacity(identifiers.len());
    let mut uuids = Vec::new();
    for identifier in &identifiers {
        let (_, object_member) = locate(*identifier)?;
        if !styles.contains(identifier) && object_member != member {
            return Err(Refusal::SharedGraph.into());
        }
        let object_component = metadata
            .component_for_entry(&object_member)?
            .ok_or(Error::InvalidSource)?;
        if *identifier != attachment {
            match metadata.uuid_owner(*identifier)? {
                Some(owner) if owner == object_component => uuids.push(*identifier),
                Some(_) => return Err(Error::InvalidSource),
                None => {},
            }
        }
        graph_objects.push(GraphObject {
            member: object_member,
            component: object_component,
            identifier: *identifier,
        });
    }
    // Source-built graphs register every object but the attachment; native
    // caption graphs may leave part of their graph out (the host's rule).
    let caption_graph = {
        let (located, _) = locate(caption)?;
        graph::unique_optional_message(located.object, super::CAPTION_INFO_MESSAGE_TYPE)?.is_some()
    };
    if preset.is_some() && !caption_graph && uuids.len() + 1 != identifiers.len() {
        return Err(Error::InvalidSource);
    }
    let theme = theme_context(source, metadata)?;
    if !styles.is_empty() {
        let (located, _) = locate(theme.stylesheet)?;
        let (_, message) = graph::unique_message(located.object, STYLESHEET_MESSAGE_TYPE)?;
        let registered =
            references::identifiers(&message.data, &STYLESHEET_STYLES).map_err(fault)?;
        if styles.iter().any(|style| {
            registered
                .iter()
                .filter(|identifier| *identifier == style)
                .count()
                != 1
        }) {
            return Err(Error::InvalidSource);
        }
    }
    let zorder = graph::root_zorder(source, &mut budget)?.ok_or(Error::InvalidSource)?;
    if zorder_count(source, zorder, drawable)? != 1 {
        return Err(Refusal::SharedGraph.into());
    }
    Ok(ChartGraph {
        drawable,
        drawable_message_index: target.drawable_message_index,
        attachment,
        character_index: target.character_index,
        member,
        component,
        objects: graph_objects,
        styles,
        preset,
        uuids,
        commented: found.comment.is_some_and(|identifier| identifier != 0),
        zorder,
        frame,
        theme,
    })
}

/// Rewrite the reference list at `path` of `object`'s message
/// `message_type` in `archive`: `removed` leave it, `added` join it.
fn relist(
    archive: &mut Archive,
    object: u64,
    message_type: u32,
    path: &[u32],
    (removed, added): (&[u64], &[u64]),
    limits: litchi_iwa_core::Limits,
) -> Result<(), Error> {
    if removed.is_empty() {
        return append(archive, object, message_type, path, added, limits);
    }
    let target = archive.object_mut(object).ok_or(Error::InvalidSource)?;
    let index = target
        .messages
        .iter()
        .position(|message| message.type_ == message_type)
        .ok_or(Error::InvalidSource)?;
    let data = target.messages[index].data.clone();
    let mut after = references::identifiers(&data, path).map_err(fault)?;
    for identifier in removed {
        if after.iter().filter(|listed| *listed == identifier).count() != 1 {
            return Err(Error::InvalidSource);
        }
    }
    after.retain(|identifier| !removed.contains(identifier));
    after.extend_from_slice(added);
    let data = references::rewrite(&data, path, &after).map_err(fault)?;
    references::replace(target, index, data, removed, added, path, limits).map_err(fault)
}

/// Drop the theme's chart-preset list when it no longer lists a preset (the
/// host clears `extensions.chart`).
fn drop_empty_presets(
    archive: &mut Archive,
    theme: u64,
    limits: litchi_iwa_core::Limits,
) -> Result<(), Error> {
    let target = archive.object_mut(theme).ok_or(Error::InvalidSource)?;
    let index = target
        .messages
        .iter()
        .position(|message| message.type_ == THEME_MESSAGE_TYPE)
        .ok_or(Error::InvalidSource)?;
    let data = &target.messages[index].data;
    if !has_path(data, &THEME_CHART_PRESETS[..2])?
        || !references::identifiers(data, &THEME_CHART_PRESETS)
            .map_err(fault)?
            .is_empty()
    {
        return Ok(());
    }
    let data = transform_length_delimited_field(data, THEME_CHART_PRESETS[0], |theme| {
        litchi_iwa_common::wire::patch_length_delimited_field(
            theme,
            THEME_CHART_PRESETS_FIELD,
            true,
            None,
        )
    })
    .map_err(|_error| Error::InvalidSource)?;
    target
        .replace_message_preserving_header(
            index,
            RawMessage {
                type_: THEME_MESSAGE_TYPE,
                data,
            },
            limits,
        )
        .map_err(graph::map_core_error)?;
    Ok(())
}

/// Every `(data, object, count)` data-reference owner `component` registers
/// for an object `wanted` accepts.
fn owned_data(
    payload: &[u8],
    component: u64,
    wanted: impl Fn(u64) -> bool,
) -> Result<Vec<(u64, u64, u64)>, Error> {
    use litchi_iwa_common::wire::repeated_length_delimited_payloads as repeated;
    let varint = |payload: &[u8], number: u32| -> Result<Option<u64>, Error> {
        let view = litchi_iwa_common::wire::WireView::parse(payload, WireLimits::DEFAULT)
            .map_err(|_error| Error::InvalidSource)?;
        Ok(view
            .fields()
            .find(|field| field.number() == number)
            .and_then(|field| litchi_iwa_common::decode_varint_from_bytes(field.payload()).ok())
            .map(|(value, _)| value))
    };
    let mut owned = Vec::new();
    for info in repeated(payload, METADATA_COMPONENTS).map_err(|_error| Error::InvalidSource)? {
        if varint(info, COMPONENT_IDENTIFIER)? != Some(component) {
            continue;
        }
        for reference in
            repeated(info, COMPONENT_DATA_REFERENCES).map_err(|_error| Error::InvalidSource)?
        {
            let data = varint(reference, DATA_REFERENCE_DATA)?.ok_or(Error::InvalidSource)?;
            for owner in
                repeated(reference, DATA_REFERENCE_OWNERS).map_err(|_error| Error::InvalidSource)?
            {
                let object = varint(owner, OWNER_OBJECT)?.ok_or(Error::InvalidSource)?;
                if wanted(object) {
                    owned.push((data, object, varint(owner, OWNER_COUNT)?.unwrap_or(1)));
                }
            }
        }
    }
    Ok(owned)
}

fn write_metadata(
    source: &Package,
    archives: &mut Vec<(Arc<str>, Archive)>,
    metadata: &objects::Metadata,
    payload: Vec<u8>,
    limits: litchi_iwa_core::Limits,
) -> Result<(), Error> {
    if payload == metadata.payload {
        return Ok(());
    }
    editable(source, archives, &metadata.component)?
        .object_mut(metadata.object)
        .ok_or(Error::InvalidSource)?
        .replace_message_preserving_header(
            metadata.message_index,
            RawMessage {
                type_: METADATA_MESSAGE_TYPE,
                data: payload,
            },
            limits,
        )
        .map_err(graph::map_core_error)?;
    Ok(())
}

/// The ordinal of `drawable` among the candidate's body charts.
fn chart_position(candidate: &Package, drawable: u64) -> Result<Option<usize>, Error> {
    use super::super::body_chart_arrangement::{self as arrangement, ArrangementBudget};
    let mut budget = ArrangementBudget::new(candidate).map_err(arrangement_error)?;
    Ok(arrangement::resolve_targets(candidate, &mut budget)
        .map_err(arrangement_error)?
        .iter()
        .position(|target| target.drawable_identifier.get() == drawable))
}

/// Copy the selected chart under fresh identifiers anchored at UTF-16 body
/// position `anchor`, offset 12 pt.
#[allow(
    clippy::too_many_lines,
    reason = "one staged copy in the host's order: graph, styles, anchor, preset, registry"
)]
pub(super) fn duplicate(
    source: &Package,
    selector: BodyChartSelector,
    anchor: TextPosition,
) -> Result<(Package, Diagnostics), Error> {
    let anchor = anchor.utf16_index();
    let limits = source
        .state
        .source
        .limits()
        .effective_archive_limits()
        .map_err(graph::map_archive_error)?;
    let metadata = objects::metadata(source)?.ok_or(Error::InvalidSource)?;
    let chart = resolve_chart(source, selector, &metadata)?;
    let first_identifier = objects::next_identifier(source, Some(&metadata))?;
    let mut remap = std::collections::HashMap::with_capacity(chart.objects.len());
    for (offset, object) in chart.objects.iter().enumerate() {
        let identifier = first_identifier
            .checked_add(offset as u64)
            .ok_or(Error::InvalidSource)?;
        remap.insert(object.identifier, identifier);
    }
    let new_drawable = remap[&chart.drawable];
    let new_attachment = remap[&chart.attachment];
    let mut frame = chart.frame;
    let position = frame.position.ok_or(Error::InvalidSource)?;
    let position = crate::drawable::Point {
        x: position.x + super::BODY_DUPLICATE_OFFSET,
        y: position.y + super::BODY_DUPLICATE_OFFSET,
    };
    if !position.x.is_finite() || !position.y.is_finite() {
        return Err(Refusal::InvalidGeometry.into());
    }
    frame.position = Some(position);
    let left_margin = {
        let mut budget = Budget::new(source)?;
        let index = ObjectIndex::new(source, &mut budget)?;
        graph::resolve_body(&index, &mut budget)?
            .root
            .left_margin
            .unwrap_or_default()
    };
    let horizontal = position.x - left_margin;
    if !horizontal.is_finite() {
        return Err(Refusal::InvalidGeometry.into());
    }

    let mut archives: Vec<(Arc<str>, Archive)> = Vec::new();
    for object in &chart.objects {
        let original = source
            .state
            .source
            .components()
            .get(object.member.as_ref())
            .and_then(|component| component.archive().object(object.identifier))
            .ok_or(Error::InvalidSource)?;
        let mut copy = super::clone_object(source, original, &remap)?;
        if object.identifier == chart.drawable {
            let index = chart.drawable_message_index;
            let message = copy.messages.get(index).ok_or(Error::InvalidSource)?;
            if message.type_ != CHART_MESSAGE_TYPE {
                return Err(Error::InvalidSource);
            }
            let data = super::super::body_chart_lifecycle::with_geometry(&message.data, frame)
                .map_err(|_error| Error::InvalidSource)?;
            copy.replace_message_preserving_header(
                index,
                RawMessage {
                    type_: CHART_MESSAGE_TYPE,
                    data,
                },
                limits,
            )
            .map_err(graph::map_core_error)?;
        }
        if object.identifier == chart.attachment {
            let (index, message) =
                graph::unique_message(&copy, graph::DRAWABLE_ATTACHMENT_MESSAGE_TYPE)?;
            let data = litchi_iwa_common::wire::patch_fixed32_field(
                &message.data,
                super::ATTACHMENT_H_OFFSET_FIELD,
                true,
                Some(horizontal.to_bits()),
            )
            .and_then(|data| {
                litchi_iwa_common::wire::patch_fixed32_field(
                    &data,
                    super::ATTACHMENT_V_OFFSET_FIELD,
                    true,
                    Some(position.y.to_bits()),
                )
            })
            .map_err(graph::map_wire_error)?;
            copy.replace_message_preserving_header(
                index,
                RawMessage {
                    type_: graph::DRAWABLE_ATTACHMENT_MESSAGE_TYPE,
                    data,
                },
                limits,
            )
            .map_err(graph::map_core_error)?;
        }
        editable(source, &mut archives, &object.member)?
            .insert_object(copy, limits)
            .map_err(graph::map_core_error)?;
    }

    // Styles join the stylesheet per member, with the host's
    // cross-component references; the U+FFFC, attachment entry and
    // z-order; the preset joins the theme.
    let mut payload = metadata.payload.clone();
    let stylesheet_component = metadata
        .component_for_entry(&chart.theme.stylesheet_member)?
        .ok_or(Error::InvalidSource)?;
    for (_, component, objects) in chart.groups() {
        let new_styles = objects
            .iter()
            .filter(|identifier| chart.styles.contains(identifier))
            .map(|identifier| remap[identifier])
            .collect::<Vec<_>>();
        if new_styles.is_empty() {
            continue;
        }
        relist(
            editable(source, &mut archives, &chart.theme.stylesheet_member)?,
            chart.theme.stylesheet,
            STYLESHEET_MESSAGE_TYPE,
            &STYLESHEET_STYLES,
            (&[], &new_styles),
            limits,
        )?;
        if component != stylesheet_component {
            payload = objects::add_external_reference(
                &payload,
                component,
                stylesheet_component,
                chart.theme.stylesheet,
            )?;
            for style in &new_styles {
                payload = objects::add_external_reference(
                    &payload,
                    stylesheet_component,
                    component,
                    *style,
                )?;
            }
        }
        if component != chart.component {
            for style in &new_styles {
                payload =
                    objects::add_external_reference(&payload, chart.component, component, *style)?;
            }
        }
    }
    insert_anchor(source, &mut archives, anchor, new_attachment, limits)?;
    append_zorder(
        source,
        &mut archives,
        chart.zorder.get(),
        &chart.member,
        new_drawable,
        limits,
    )?;
    if let Some(preset) = chart.preset {
        let new_preset = remap[&preset];
        append(
            editable(source, &mut archives, &chart.theme.theme_member)?,
            chart.theme.theme,
            THEME_MESSAGE_TYPE,
            &THEME_CHART_PRESETS,
            &[new_preset],
            limits,
        )?;
        if chart.theme.theme_component != chart.component {
            payload = objects::add_external_reference(
                &payload,
                chart.theme.theme_component,
                chart.component,
                new_preset,
            )?;
        }
    }

    // Registry: the watermark and UUIDs, then the copied data owners.
    let last = remap.values().copied().max().ok_or(Error::InvalidSource)?;
    let groups = chart.groups();
    let uuid_groups = groups
        .iter()
        .map(|(_, component, objects)| {
            (
                *component,
                objects
                    .iter()
                    .filter(|identifier| chart.uuids.contains(identifier))
                    .map(|identifier| remap[identifier])
                    .collect::<Vec<_>>(),
            )
        })
        .filter(|(_, objects)| !objects.is_empty())
        .collect::<Vec<_>>();
    let borrowed = uuid_groups
        .iter()
        .map(|(component, objects)| (*component, objects.as_slice()))
        .collect::<Vec<_>>();
    payload = objects::Metadata::with_payload(&metadata, payload)
        .registered_in(source, &borrowed, last)?;
    for (_, component, objects) in &groups {
        for (data, object, count) in
            owned_data(&payload, *component, |object| objects.contains(&object))?
        {
            for _ in 0..count {
                payload = objects::adjust_data_reference(
                    &payload,
                    *component,
                    data,
                    remap[&object],
                    true,
                )?;
            }
        }
    }
    write_metadata(source, &mut archives, &metadata, payload, limits)?;
    let candidate = graph::publish_archives(source, archives)?;
    let copy = resolve_chart(
        &candidate,
        BodyChartSelector::index(
            chart_position(&candidate, new_drawable)?.ok_or(Error::Verification)?,
        ),
        &objects::metadata(&candidate)?.ok_or(Error::Verification)?,
    )
    .map_err(|_error| Error::Verification)?;
    if copy.character_index != anchor
        || copy.frame != frame
        || copy.identifiers()
            != chart
                .objects
                .iter()
                .map(|object| remap[&object.identifier])
                .collect::<Vec<_>>()
    {
        return Err(Error::Verification);
    }
    Ok((
        candidate,
        Diagnostics {
            changed: true,
            created_objects: chart.objects.len(),
            ..Diagnostics::default()
        },
    ))
}

/// Remove the selected chart, its attachment and U+FFFC, its private
/// styles and preset; a direct comment is cleared first.
#[allow(
    clippy::too_many_lines,
    reason = "one staged removal in the host's order: text, z-order, styles, preset, registry"
)]
pub(super) fn remove(
    source: &Package,
    selector: BodyChartSelector,
) -> Result<(Package, Diagnostics), Error> {
    let metadata = objects::metadata(source)?.ok_or(Error::InvalidSource)?;
    let chart = resolve_chart(source, selector, &metadata)?;
    if chart.commented {
        let cleared = super::super::text_box_drawable_comment::clear_chart_comment(
            source,
            chart.drawable,
            chart.drawable_message_index,
        )
        .map_err(super::comment_error)?
        .ok_or(Error::InvalidSource)?;
        let (package, mut diagnostics) = remove(&cleared, selector)?;
        diagnostics.changed = true;
        return Ok((package, diagnostics));
    }
    let limits = source
        .state
        .source
        .limits()
        .effective_archive_limits()
        .map_err(graph::map_archive_error)?;
    let mut archives: Vec<(Arc<str>, Archive)> = Vec::new();

    // The anchoring U+FFFC and its attachment entry.
    let (mut storage, _) = objects::body_storage(source)?;
    let at = chart.character_index;
    if storage.text.encode_utf16().nth(at as usize) != Some(0xfffc) {
        return Err(Error::InvalidSource);
    }
    let removed = objects::splice_text(source, &mut storage, at, at + 1, "")?;
    if removed != [chart.attachment] {
        return Err(Refusal::SharedGraph.into());
    }
    editable(source, &mut archives, &storage.component)?
        .object_mut(storage.identifier.get())
        .ok_or(Error::InvalidSource)?
        .replace_message_pruning_object_references_preserving_header(
            storage.message_index,
            RawMessage {
                type_: storage.message_type,
                data: storage.payload.clone(),
            },
            &[chart.attachment],
            limits,
        )
        .map_err(graph::map_core_error)?;

    // The z-order, the stylesheet and the theme's chart presets.
    {
        let mut budget = Budget::new(source)?;
        let index = ObjectIndex::new(source, &mut budget)?;
        let member = index.component_name(index.get(chart.zorder)?.component_index)?;
        zorder_without(
            editable(source, &mut archives, &member)?,
            chart.zorder,
            chart.drawable,
            limits,
        )?;
    }
    if !chart.styles.is_empty() {
        relist(
            editable(source, &mut archives, &chart.theme.stylesheet_member)?,
            chart.theme.stylesheet,
            STYLESHEET_MESSAGE_TYPE,
            &STYLESHEET_STYLES,
            (&chart.styles, &[]),
            limits,
        )?;
    }
    if let Some(preset) = chart.preset {
        let archive = editable(source, &mut archives, &chart.theme.theme_member)?;
        relist(
            archive,
            chart.theme.theme,
            THEME_MESSAGE_TYPE,
            &THEME_CHART_PRESETS,
            (&[preset], &[]),
            limits,
        )?;
        drop_empty_presets(archive, chart.theme.theme, limits)?;
    }

    // The graph objects; nothing outside the graph may reference them.
    let identifiers = chart.identifiers();
    for object in &chart.objects {
        editable(source, &mut archives, &object.member)?
            .remove_object(object.identifier)
            .ok_or(Error::InvalidSource)?;
    }
    let removed_set = identifiers.iter().copied().collect::<HashSet<_>>();
    let references_removed = |archive: &Archive| {
        archive.objects.iter().any(|object| {
            object.archive_info.message_infos.iter().any(|info| {
                info.object_references
                    .iter()
                    .chain(
                        info.field_infos
                            .iter()
                            .flat_map(|field| &field.object_references),
                    )
                    .any(|identifier| removed_set.contains(identifier))
            })
        })
    };
    for component in source.state.source.components().iter() {
        let staged = archives
            .iter()
            .find(|(name, _)| name.as_ref() == component.name())
            .map_or(component.archive(), |(_, archive)| archive);
        if references_removed(staged) {
            return Err(Refusal::SharedGraph.into());
        }
    }

    // Registry: external references to the graph, its data owners, UUIDs,
    // orphaned data and the watermark.
    let mut payload = metadata.payload.clone();
    let mut candidates = Vec::new();
    for (_, component, objects) in chart.groups() {
        for identifier in &objects {
            payload = objects::remove_external_references(&payload, component, *identifier)?;
        }
        for (data, object, count) in
            owned_data(&payload, component, |object| objects.contains(&object))?
        {
            for _ in 0..count {
                payload = objects::adjust_data_reference(&payload, component, data, object, false)?;
            }
            candidates.push(data);
        }
    }
    for object in &chart.objects {
        if chart.uuids.contains(&object.identifier) {
            payload = objects::remove_object_uuid(&payload, object.component, object.identifier)?;
        }
    }
    let (released, deletions, removed_assets) =
        release_orphans(source, &archives, &metadata, payload, candidates)?;
    let remaining = objects::max_identifier(source, &archives);
    let payload = objects::released(&released, metadata.last, &identifiers, remaining)?;
    write_metadata(source, &mut archives, &metadata, payload, limits)?;
    let deletions = deletions.iter().map(String::as_str).collect::<Vec<_>>();
    let candidate = graph::publish_changes(source, archives, &[], &[], &deletions)?;
    if chart_position(&candidate, chart.drawable)?.is_some() {
        return Err(Error::Verification);
    }
    {
        let mut budget = Budget::new(&candidate)?;
        let index = ObjectIndex::new(&candidate, &mut budget)?;
        if identifiers
            .iter()
            .filter_map(|identifier| NonZeroU64::new(*identifier))
            .any(|identifier| index.get(identifier).is_ok())
        {
            return Err(Error::Verification);
        }
    }
    Ok((
        candidate,
        Diagnostics {
            changed: true,
            removed_objects: identifiers.len(),
            removed_assets,
            ..Diagnostics::default()
        },
    ))
}
