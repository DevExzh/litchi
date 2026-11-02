//! Body chart creation (record 0418; the host's `add_body_chart`): the
//! shared `litchi_iwa_style::chart_lifecycle::author` graph with the Pages
//! profile in the body component, a body attachment and U+FFFC at the
//! anchor, the z-order (created on first use), the styles in the document
//! stylesheet and the preset in the theme's chart presets, with the host's
//! cross-component references and UUIDs.

use std::sync::Arc;

use litchi_iwa_common::chart::data::ChartData;
use litchi_iwa_common::chart::kind::Kind;
use litchi_iwa_common::wire::transform_length_delimited_field;
use litchi_iwa_core::{Archive, RawMessage};
use litchi_iwa_style::chart_lifecycle::author::{self, Ids, Profile, Request};
use litchi_iwa_table::structure::references;
use litchi_iwa_text::TextPosition;

use super::{
    Budget, CreationIds, DOCUMENT_COMPONENT_IDENTIFIER, Diagnostics, Error, METADATA_MESSAGE_TYPE,
    ObjectIndex, Package, Pb, Refusal, STANDARD_VERSION, THEME_MESSAGE_TYPE, append_zorder,
    create_zorder, drawable_payload, editable, graph, insert_anchor, new_object, objects,
};

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
            let identifier = std::num::NonZeroU64::new(identifier).ok_or(Error::InvalidSource)?;
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
    let zorder_identifier = zorder.map_or(creation.zorder.unwrap_or(0), std::num::NonZeroU64::get);
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
        std::num::NonZeroU64::new(ids.drawable),
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
