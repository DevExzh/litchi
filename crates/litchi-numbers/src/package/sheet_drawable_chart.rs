//! Sheet chart creation, duplication and removal inside the structural
//! sheet-drawable transaction (record 0418; the host's `add_sheet_chart`,
//! `duplicate_sheet_chart` and `remove_sheet_chart`).
//!
//! A Numbers chart's private graph lives in the CalculationEngine
//! component: the drawable, its caption graph and title stand-in, its
//! mediator, its theme preset and its sixteen styles. A new graph comes
//! from the shared `litchi_iwa_style::chart_lifecycle::author`; its styles
//! join the document stylesheet and its preset the theme's chart presets,
//! with the host's cross-component references and UUIDs. A copy clones the
//! graph under fresh identifiers in the host's order; a removal detaches
//! the styles and preset and deletes the graph.

use std::collections::{BTreeSet, HashMap};

use litchi_iwa_common::chart::data::ChartData;
use litchi_iwa_common::chart::kind::Kind as ChartKind;
use litchi_iwa_common::wire::transform_length_delimited_field;
use litchi_iwa_style::chart_lifecycle::author::{self, Ids, Profile, Request};

use super::sheet_drawable::{Error, Refusal};
use super::sheet_drawable_author::Writer;
use super::sheet_drawable_duplicate as duplicate;
use super::structure::{
    self, Fault, FaultResult, Working,
    metadata::{self, Registry, Transition},
    references,
};
use crate::sheet::drawable::Geometry;

/// `TN.ThemeArchive` and `TSS.StylesheetArchive`.
const THEME_MESSAGE_TYPE: u32 = 12_009;
const STYLESHEET_MESSAGE_TYPE: u32 = 401;
const DOCUMENT_THEME_FIELD: u32 = 6;
/// `TSS.ThemeArchive` (field 1 of the Numbers theme): its document
/// stylesheet, its text presets (`TSWP.ThemePresetsArchive` at 110, first
/// paragraph style preset at 7) and its chart presets (120, list at 1).
const THEME_SUPER: u32 = 1;
const THEME_STYLESHEET: u32 = 4;
const THEME_TEXT_PRESETS: u32 = 110;
const THEME_PARAGRAPH_PRESETS: u32 = 7;
const THEME_CHART_PRESETS: u32 = 120;
const STYLESHEET_STYLES: u32 = 1;
const ENGINE_SUFFIX: &str = "CalculationEngine.iwa";
const WRAP_TYPE_SQUARE: u64 = 4;
const WRAP_DIRECTION_BOTH: u64 = 2;
const WRAP_FIT_TEXT: u64 = 1;
const WRAP_MARGIN: f32 = 12.0;
const WRAP_ALPHA: f32 = 0.5;
const DEFAULT_FLAGS: u64 = 3;

/// One staged chart creation.
#[derive(Debug, Clone, PartialEq)]
pub(super) struct NewChart {
    pub(super) kind: ChartKind,
    pub(super) data: ChartData,
    pub(super) geometry: Geometry,
}

impl NewChart {
    /// Whether the kind is assignable, the data forms it and the frame has a
    /// finite position and a positive size.
    pub(super) fn validate(&self) -> Result<(), Refusal> {
        if !litchi_iwa_style::chart_lifecycle::assignable_kind(self.kind)
            || author::series_count(self.kind, &self.data).is_none()
        {
            return Err(Refusal::InvalidChart);
        }
        let valid = self.geometry.is_valid()
            && self.geometry.position.is_some()
            && self
                .geometry
                .size
                .is_some_and(|size| size.width > 0.0 && size.height > 0.0);
        if valid {
            Ok(())
        } else {
            Err(Refusal::InvalidGeometry)
        }
    }
}

/// The theme, stylesheet and paragraph style a chart graph references, and
/// the CalculationEngine member that owns chart graphs.
struct Context {
    theme: u64,
    stylesheet: u64,
    paragraph: u64,
    engine: String,
}

fn payload_of<'w>(
    working: &'w Working<'_>,
    object: u64,
    message_type: u32,
) -> FaultResult<&'w [u8]> {
    let object = working.object(object)?;
    let index = structure::unique_message(object, &[message_type])?;
    Ok(&object.messages[index].data)
}

fn first(payload: &[u8], path: &[u32]) -> FaultResult<Option<u64>> {
    references::records(payload, path)?
        .first()
        .map(|record| references::reference_identifier(record))
        .transpose()
}

fn has_path(payload: &[u8], path: &[u32]) -> FaultResult<bool> {
    let mut current = payload;
    for number in path {
        let records = references::records(current, &[*number])?;
        match records.as_slice() {
            [] => return Ok(false),
            [inner] => current = inner,
            _ => return Err(Fault::Invalid),
        }
    }
    Ok(true)
}

impl Context {
    fn read(working: &Working<'_>) -> FaultResult<Self> {
        let document = payload_of(working, 1, 1)?;
        let theme = first(document, &[DOCUMENT_THEME_FIELD])?.ok_or(Fault::Invalid)?;
        let payload = payload_of(working, theme, THEME_MESSAGE_TYPE)?;
        let stylesheet = first(payload, &[THEME_SUPER, THEME_STYLESHEET])?.ok_or(Fault::Invalid)?;
        let paragraph = if has_path(payload, &[THEME_SUPER, THEME_TEXT_PRESETS])? {
            references::identifiers(
                payload,
                &[THEME_SUPER, THEME_TEXT_PRESETS, THEME_PARAGRAPH_PRESETS],
            )?
            .into_iter()
            .find(|identifier| *identifier != 0)
        } else {
            None
        }
        .ok_or(Fault::Unsupported)?;
        let engine = working
            .component_names()
            .into_iter()
            .find(|name| name.ends_with(ENGINE_SUFFIX))
            .ok_or(Fault::Unsupported)?;
        Ok(Self {
            theme,
            stylesheet,
            paragraph,
            engine,
        })
    }
}

fn component(working: &Working<'_>, registry: &Registry, object: u64) -> FaultResult<u64> {
    let member = working.locate(object)?.ok_or(Fault::Invalid)?;
    Ok(registry.component_for_entry(&member)?.identifier)
}

/// Rewrite the repeated reference list at `path` of `object`'s message
/// `message_type` to `after`, keeping the header in step.
fn relist(
    working: &mut Working<'_>,
    object: u64,
    message_type: u32,
    path: &[u32],
    (removed, appended): (&[u64], &[u64]),
) -> FaultResult<()> {
    let limits = working.archive_limits();
    let target = working.object_mut(object)?;
    let index = structure::unique_message(target, &[message_type])?;
    let mut data = target.messages[index].data.clone();
    if !has_path(&data, &path[..path.len() - 1])? {
        // The theme's chart-preset list is created empty on first use.
        let empty = [0xc2, 0x07, 0x00];
        data = transform_length_delimited_field(&data, path[0], |inner| {
            let mut extended = inner.to_vec();
            extended.extend_from_slice(&empty);
            Ok::<_, litchi_iwa_common::Error>(extended)
        })
        .map_err(|error| structure::map_wire(&error))?;
    }
    let mut after = references::identifiers(&data, path)?;
    after.retain(|identifier| !removed.contains(identifier));
    for identifier in appended {
        if after.contains(identifier) {
            return Err(Fault::Invalid);
        }
        structure::reserve(&mut after, 1)?;
        after.push(*identifier);
    }
    let data = references::rewrite(&data, path, &after)?;
    references::replace(target, index, data, removed, appended, path, limits)
}

fn preset_path() -> [u32; 3] {
    [THEME_SUPER, THEME_CHART_PRESETS, 1]
}

/// The `TSD.DrawableArchive` of a new chart (the host's square text wrap,
/// unlocked, with title and caption stand-ins).
fn drawable_archive(geometry: Geometry, sheet: u64, title: u64, caption: u64) -> Vec<u8> {
    let (x, y) = geometry
        .position
        .map_or((0.0, 0.0), |point| (point.x, point.y));
    let (width, height) = geometry
        .size
        .map_or((0.0, 0.0), |size| (size.width, size.height));
    let point = |x: f32, y: f32| Writer::new().float(1, x).float(2, y);
    Writer::new()
        .message(
            1,
            &Writer::new()
                .message(1, &point(x, y))
                .message(2, &point(width, height))
                .varint(3, DEFAULT_FLAGS)
                .float(4, 0.0),
        )
        .reference(2, sheet)
        .message(
            3,
            &Writer::new()
                .varint(1, WRAP_TYPE_SQUARE)
                .varint(2, WRAP_DIRECTION_BOTH)
                .varint(3, WRAP_FIT_TEXT)
                .float(4, WRAP_MARGIN)
                .float(5, WRAP_ALPHA)
                .boolean(6, false),
        )
        .boolean(5, false)
        .boolean(7, false)
        .reference(10, title)
        .reference(11, caption)
        .boolean(12, false)
        .boolean(13, false)
        .finish()
}

/// Author `chart` on sheet `sheet` (member `sheet_member`); returns its
/// drawable.
pub(super) fn add(
    working: &mut Working<'_>,
    registry: &Registry,
    transition: &mut Transition,
    (sheet, sheet_member): (u64, &str),
    chart: &NewChart,
) -> Result<u64, Error> {
    let context = Context::read(working)?;
    let allocated = working.allocate(21)?;
    let mut next = allocated.into_iter();
    let ids = Ids::allocate(Profile::Numbers, || next.next().ok_or(Fault::Invalid))?;
    let drawable = drawable_archive(chart.geometry, sheet, ids.title, ids.caption);
    let request = Request {
        profile: Profile::Numbers,
        kind: chart.kind,
        data: &chart.data,
        drawable: &drawable,
        stylesheet: context.stylesheet,
        paragraph_style: context.paragraph,
    };
    let objects =
        author::objects(&ids, &request, working.archive_limits()).map_err(|error| match error {
            litchi_iwa_style::chart_lifecycle::Error::InvalidData => {
                Error::Refused(Refusal::InvalidChart)
            },
            _ => Error::InvalidSource,
        })?;
    for object in objects {
        working.insert(&context.engine, object)?;
    }
    let styles = ids.styles();
    relist(
        working,
        context.stylesheet,
        STYLESHEET_MESSAGE_TYPE,
        &[STYLESHEET_STYLES],
        (&[], &styles),
    )?;
    relist(
        working,
        context.theme,
        THEME_MESSAGE_TYPE,
        &preset_path(),
        (&[], &[ids.preset]),
    )?;
    // Registry edges in the host's order, then the graph's UUIDs.
    let owner = registry.component_for_entry(&context.engine)?.identifier;
    let sheet_component = registry.component_for_entry(sheet_member)?.identifier;
    let stylesheet_component = component(working, registry, context.stylesheet)?;
    let theme_component = component(working, registry, context.theme)?;
    let paragraph_component = component(working, registry, context.paragraph)?;
    if stylesheet_component != owner {
        transition.reference(registry, owner, stylesheet_component, context.stylesheet)?;
        for style in &styles {
            transition.reference(registry, stylesheet_component, owner, *style)?;
        }
    }
    transition.reference(registry, sheet_component, owner, ids.drawable)?;
    transition.reference(registry, theme_component, owner, ids.preset)?;
    transition.reference(registry, owner, paragraph_component, context.paragraph)?;
    for identifier in ids.all() {
        let seed = metadata::identity_seed(working, identifier)?;
        transition.register_uuid(registry, owner, identifier, seed)?;
    }
    Ok(ids.drawable)
}

fn outgoing(working: &Working<'_>, object: u64) -> FaultResult<Vec<u64>> {
    let mut found = Vec::new();
    for info in &working.object(object)?.archive_info.message_infos {
        for identifier in info.object_references.iter().chain(
            info.field_infos
                .iter()
                .flat_map(|field| &field.object_references),
        ) {
            if !found.contains(identifier) {
                structure::reserve(&mut found, 1)?;
                found.push(*identifier);
            }
        }
    }
    Ok(found)
}

/// The chart's private graph inside its member, in the host's
/// `chart_graph` order: the drawable, its caption graph, its title and
/// mediator, then its preset and styles in `TSCH.ChartArchive` field
/// order; anything else of the closure follows in identifier order.
fn chart_graph(working: &Working<'_>, chart: u64, sheet: u64) -> FaultResult<(String, Vec<u64>)> {
    let member = working.locate(chart)?.ok_or(Fault::Invalid)?;
    let reachable = |root: u64| -> FaultResult<Vec<u64>> {
        let mut found = vec![root];
        let mut seen = BTreeSet::from([root, sheet]);
        let mut index = 0;
        while index < found.len() {
            for target in outgoing(working, found[index])? {
                if seen.insert(target) && working.locate(target)?.as_deref() == Some(&*member) {
                    found.push(target);
                }
            }
            index += 1;
        }
        Ok(found)
    };
    let closure: BTreeSet<u64> = reachable(chart)?.into_iter().collect();
    let payload = payload_of(working, chart, author::CHART_MESSAGE_TYPE)?;
    let mut ordered = vec![chart];
    let push = |identifier: u64, ordered: &mut Vec<u64>| {
        if closure.contains(&identifier) && !ordered.contains(&identifier) {
            ordered.push(identifier);
        }
    };
    if let Some(caption) = first(payload, &[1, 11])?
        && closure.contains(&caption)
    {
        for identifier in reachable(caption)? {
            push(identifier, &mut ordered);
        }
    }
    for path in [[1, 10], [10_000, 8]] {
        if let Some(identifier) = first(payload, &path)? {
            push(identifier, &mut ordered);
        }
    }
    for number in [4, 9, 10, 11, 12, 13, 14, 15, 16, 17] {
        for identifier in references::identifiers(payload, &[10_000, number])? {
            push(identifier, &mut ordered);
        }
    }
    for identifier in &closure {
        push(*identifier, &mut ordered);
    }
    Ok((member, ordered))
}

/// The chart's objects among the stylesheet's styles and the theme's chart
/// presets.
fn registered(
    working: &Working<'_>,
    context: &Context,
    objects: &[u64],
) -> FaultResult<(Vec<u64>, Vec<u64>)> {
    let styles = references::identifiers(
        payload_of(working, context.stylesheet, STYLESHEET_MESSAGE_TYPE)?,
        &[STYLESHEET_STYLES],
    )?;
    let theme = payload_of(working, context.theme, THEME_MESSAGE_TYPE)?;
    let presets = if has_path(theme, &[THEME_SUPER, THEME_CHART_PRESETS])? {
        references::identifiers(theme, &preset_path())?
    } else {
        Vec::new()
    };
    Ok((
        objects
            .iter()
            .copied()
            .filter(|object| styles.contains(object))
            .collect(),
        objects
            .iter()
            .copied()
            .filter(|object| presets.contains(object))
            .collect(),
    ))
}

/// The `TSP.Reference` paths of a chart graph's cloned messages; `None` for
/// a message whose private references the copy cannot follow.
pub(super) fn reference_paths(message_type: u32) -> Option<&'static [&'static [u32]]> {
    const CHART: &[&[u32]] = &[
        &[1, 2],
        &[1, 9],
        &[1, 10],
        &[1, 11],
        &[10_000, 4],
        &[10_000, 8],
        &[10_000, 9],
        &[10_000, 10],
        &[10_000, 11],
        &[10_000, 12],
        &[10_000, 13],
        &[10_000, 14],
        &[10_000, 15],
        &[10_000, 16],
        &[10_000, 17],
        &[10_000, 20],
        &[10_000, 23],
    ];
    const PRESET: &[&[u32]] = &[&[1], &[2], &[3], &[4], &[5], &[6]];
    const MEDIATOR: &[&[u32]] = &[&[1, 1]];
    match message_type {
        author::CHART_MESSAGE_TYPE => Some(CHART),
        author::PRESET_MESSAGE_TYPE => Some(PRESET),
        author::MEDIATOR_MESSAGE_TYPE => Some(MEDIATOR),
        _ => None,
    }
}

/// `object` cloned under `remap`, refusing a private reference outside the
/// message's known reference paths (the host's "unrecognized private
/// reference" check).
fn clone(
    working: &Working<'_>,
    object: u64,
    remap: &HashMap<u64, u64>,
) -> Result<litchi_iwa_core::ArchiveObject, Error> {
    let source = working.object(object)?;
    for (message, info) in source
        .messages
        .iter()
        .zip(&source.archive_info.message_infos)
    {
        let Some(paths) = reference_paths(message.type_) else {
            continue;
        };
        let mut known = Vec::new();
        for path in paths {
            if has_path(&message.data, &path[..path.len() - 1])? {
                known.extend(references::identifiers(&message.data, path)?);
            }
        }
        if info
            .object_references
            .iter()
            .any(|identifier| remap.contains_key(identifier) && !known.contains(identifier))
        {
            return Err(Refusal::UnsupportedDependency.into());
        }
    }
    duplicate::clone_object(source, remap, working.archive_limits()).map_err(|fault| match fault {
        Fault::Dependency => Error::Refused(Refusal::UnsupportedDependency),
        fault => fault.into(),
    })
}

/// Copy chart `chart` of sheet `sheet` (member `sheet_member`) under fresh
/// identifiers, offset like Numbers' Duplicate; returns the copy.
pub(super) fn duplicate(
    working: &mut Working<'_>,
    registry: &Registry,
    transition: &mut Transition,
    (sheet, sheet_member): (u64, &str),
    chart: u64,
) -> Result<u64, Error> {
    let context = Context::read(working)?;
    let (member, objects) = chart_graph(working, chart, sheet)?;
    let (styles, presets) = registered(working, &context, &objects)?;
    let identifiers = working.allocate(objects.len())?;
    let remap: HashMap<u64, u64> = objects
        .iter()
        .copied()
        .zip(identifiers.iter().copied())
        .collect();
    for object in &objects {
        let mut copy = clone(working, *object, &remap)?;
        if *object == chart {
            let data = duplicate::offset(
                &copy.messages[0].data,
                author::CHART_MESSAGE_TYPE,
                duplicate::OFFSET,
            )?;
            copy.replace_message_preserving_header(
                0,
                litchi_iwa_core::RawMessage {
                    type_: author::CHART_MESSAGE_TYPE,
                    data,
                },
                working.archive_limits(),
            )
            .map_err(|error| structure::map_core(&error))?;
        }
        working.insert(&member, copy)?;
    }
    let new_styles: Vec<u64> = styles.iter().map(|style| remap[style]).collect();
    let new_presets: Vec<u64> = presets.iter().map(|preset| remap[preset]).collect();
    relist(
        working,
        context.stylesheet,
        STYLESHEET_MESSAGE_TYPE,
        &[STYLESHEET_STYLES],
        (&[], &new_styles),
    )?;
    if !new_presets.is_empty() {
        relist(
            working,
            context.theme,
            THEME_MESSAGE_TYPE,
            &preset_path(),
            (&[], &new_presets),
        )?;
    }
    let owner = registry.component_for_entry(&member)?.identifier;
    let sheet_component = registry.component_for_entry(sheet_member)?.identifier;
    let stylesheet_component = component(working, registry, context.stylesheet)?;
    let theme_component = component(working, registry, context.theme)?;
    if stylesheet_component != owner && !new_styles.is_empty() {
        transition.reference(registry, owner, stylesheet_component, context.stylesheet)?;
        for style in &new_styles {
            transition.reference(registry, stylesheet_component, owner, *style)?;
        }
    }
    let copy = remap[&chart];
    transition.reference(registry, sheet_component, owner, copy)?;
    for preset in &new_presets {
        transition.reference(registry, theme_component, owner, *preset)?;
    }
    for object in &objects {
        if registry
            .uuids
            .iter()
            .any(|binding| binding.component == owner && binding.object == *object)
        {
            let new = remap[object];
            let seed = metadata::identity_seed(working, new)?;
            transition.register_uuid(registry, owner, new, seed)?;
        }
    }
    for owner_record in &registry.data_references {
        if owner_record.component == owner && objects.contains(&owner_record.object) {
            structure::reserve(&mut transition.data_owner_additions, 1)?;
            transition.data_owner_additions.push(metadata::DataOwner {
                object: remap[&owner_record.object],
                ..*owner_record
            });
        }
    }
    Ok(copy)
}

/// Detach and delete the private graphs of `charts` (sheet `sheet`);
/// returns the deleted objects.
pub(super) fn remove(
    working: &mut Working<'_>,
    registry: &Registry,
    transition: &mut Transition,
    sheet: u64,
    charts: &[u64],
) -> Result<Vec<u64>, Error> {
    if charts.is_empty() {
        return Ok(Vec::new());
    }
    let context = Context::read(working)?;
    let mut deleted = Vec::new();
    let mut members = Vec::new();
    for chart in charts {
        let (member, objects) = chart_graph(working, *chart, sheet)?;
        let (styles, presets) = registered(working, &context, &objects)?;
        relist(
            working,
            context.stylesheet,
            STYLESHEET_MESSAGE_TYPE,
            &[STYLESHEET_STYLES],
            (&styles, &[]),
        )?;
        if !presets.is_empty() {
            relist(
                working,
                context.theme,
                THEME_MESSAGE_TYPE,
                &preset_path(),
                (&presets, &[]),
            )?;
        }
        if !members.contains(&member) {
            members.push(member);
        }
        deleted.extend(objects);
    }
    // No surviving object outside the graphs may reference them.
    let removed: BTreeSet<u64> = deleted.iter().copied().collect();
    for name in working.component_names() {
        for object in &working.archive(&name)?.objects {
            let identifier = object.archive_info.identifier.ok_or(Fault::Invalid)?;
            if removed.contains(&identifier) || identifier == sheet {
                continue;
            }
            if outgoing(working, identifier)?
                .iter()
                .any(|target| removed.contains(target))
            {
                return Err(Refusal::UnsupportedDependency.into());
            }
        }
    }
    transition.remove_objects(registry, &deleted)?;
    transition.release_suffix(&deleted)?;
    for identifier in &deleted {
        working.remove(*identifier)?;
    }
    // A component left without charts drops its stylesheet edges.
    let stylesheet_component = component(working, registry, context.stylesheet)?;
    let paragraph_component = component(working, registry, context.paragraph)?;
    for member in members {
        let still_charts = working.archive(&member)?.objects.iter().any(|object| {
            object
                .messages
                .iter()
                .any(|message| message.type_ == author::CHART_MESSAGE_TYPE)
        });
        if still_charts {
            continue;
        }
        let owner = registry.component_for_entry(&member)?.identifier;
        transition.remove_reference(registry, owner, paragraph_component, context.paragraph)?;
        transition.remove_reference(registry, owner, stylesheet_component, context.stylesheet)?;
    }
    transition.release_unreferenced_data(working, registry)?;
    Ok(deleted)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::sheet::drawable::{Point, Size};

    #[test]
    fn a_new_chart_needs_a_formable_kind_and_a_frame() {
        let data = ChartData::new(
            vec!["a".to_owned()],
            vec!["x".to_owned()],
            vec![vec![Some(1.0)]],
        )
        .expect("data");
        let geometry = Geometry::new(
            Point { x: 10.0, y: 10.0 },
            Size {
                width: 100.0,
                height: 80.0,
            },
        );
        let chart = NewChart {
            kind: ChartKind::Bar2d,
            data: data.clone(),
            geometry,
        };
        assert_eq!(chart.validate(), Ok(()));
        let bubble = NewChart {
            kind: ChartKind::Bubble2d,
            ..chart.clone()
        };
        assert_eq!(bubble.validate(), Err(Refusal::InvalidChart));
        let empty = NewChart {
            geometry: Geometry::new(
                Point { x: 10.0, y: 10.0 },
                Size {
                    width: 0.0,
                    height: 80.0,
                },
            ),
            ..chart
        };
        assert_eq!(empty.validate(), Err(Refusal::InvalidGeometry));
    }
}
