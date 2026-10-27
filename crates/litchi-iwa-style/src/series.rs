//! The style-variation graph engine of per-series chart overrides (record
//! 0375a): the sparse `TSCH.ChartSeriesNonStyleArchive` objects a chart
//! references from `series_non_styles`, one per series that overrides a
//! behavioral property (connection lines, wedge explosions, label
//! distances, …).
//!
//! [`set_values`] writes one property for every series, as the migration
//! host's editors do byte for byte: an override object is patched in place
//! when only this chart owns it, copied when another chart shares it,
//! allocated when the series had none, and removed (with its stylesheet
//! membership, UUID binding and object-container slot) when it becomes empty.
//! New styled overrides join the chart's document stylesheet; new unstyled
//! (geometry-only) overrides live in a fresh `TSP.ObjectContainer`
//! component, as iWork saves them. The chart's sparse array, the package
//! metadata registry and the issued-identifier watermark follow.
//!
//! A [`Property`] maps one value type to the override payload; the
//! properties of the Format panel are in [`property`].

use std::collections::{BTreeMap, HashSet};

use litchi_iwa_common::wire;
use litchi_iwa_core::{Archive, ArchiveObject, RawMessage};
use litchi_iwa_protos::chart_style_codec::{self as codec, Limits as CodecLimits};

use crate::graph::{
    self, DOCUMENT_METADATA_MEMBER, Error, Location, METADATA_MEMBER, Result, Workspace,
    push_message_field, push_packed_field, push_varint_field,
};

const CHART_MESSAGE_TYPE: u32 = 5_021;
const SERIES_STYLE_MESSAGE_TYPE: u32 = 5_028;
const SERIES_NON_STYLE_MESSAGE_TYPE: u32 = 5_029;
const STYLESHEET_MESSAGE_TYPE: u32 = 401;
const OBJECT_CONTAINER_MESSAGE_TYPE: u32 = 11_008;
const CHART_STYLE_MESSAGE_TYPES: [u32; 8] =
    [5_022, 5_023, 5_024, 5_025, 5_026, 5_027, 5_028, 5_029];

const STYLE_PARENT_FIELD: u32 = 1;
const STYLE_PARENT_REFERENCE_FIELD: u32 = 3;
const STYLE_STYLESHEET_FIELD: u32 = 5;
const REFERENCE_IDENTIFIER_FIELD: u32 = 1;
const STYLESHEET_STYLES_FIELD: u32 = 1;
const CONTAINER_IDENTIFIER_FIELD: u32 = 1;
const CONTAINER_OBJECTS_FIELD: u32 = 2;
const PRIMARY_PACKAGE_IDENTIFIER: u64 = 1;
const OBJECT_CONTAINER_LOCATOR: &str = "ObjectContainer";
const OBJECT_CONTAINER_VERSION: [u32; 3] = [2, 0, 0];
const STANDARD_MESSAGE_VERSION: [u32; 3] = [1, 0, 5];
const CAPABILITY_FIELDS: [u32; 3] = [10_001, 10_002, 10_003];

/// The style parent of an override object allocated for a series that had
/// none.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Base {
    /// Joins the chart's document stylesheet (labels, formats, lines).
    Styled,
    /// An empty style parent in a standalone object container (pure
    /// geometry: wedge explosions and label distances).
    Unstyled,
}

/// One per-series override property.
pub trait Property {
    /// The value of one series.
    type Value: Clone + PartialEq;

    /// Where a newly allocated override lives.
    fn base(&self) -> Base;

    /// The value of a series without an override.
    fn default_value(&self) -> Self::Value;

    /// Read the value from an override payload.
    ///
    /// # Errors
    ///
    /// The payload is invalid or stores an unsupported value.
    fn read(&self, payload: &[u8]) -> Result<Self::Value>;

    /// Store `value` in an override payload, keeping every other byte.
    ///
    /// # Errors
    ///
    /// The payload is invalid or `value` cannot be stored.
    fn patch(&self, payload: &[u8], value: &Self::Value) -> Result<Vec<u8>>;
}

#[derive(Debug, Clone)]
struct Slot {
    location: Location,
    object: u64,
    message: usize,
}

struct SeriesGraph {
    chart_message: usize,
    slots: Vec<Option<Slot>>,
}

/// Read `property` for each of the `series_count` series of chart drawable
/// `chart`, in series order.
///
/// # Errors
///
/// An invalid chart, sparse array or override, or a value the property
/// rejects.
pub fn values<P: Property>(
    workspace: &Workspace<'_>,
    chart: u64,
    series_count: usize,
    property: &P,
) -> Result<Vec<P::Value>> {
    let graph = series_graph(workspace, chart, series_count)?;
    read_values(workspace, &graph, property)
}

fn read_values<P: Property>(
    workspace: &Workspace<'_>,
    graph: &SeriesGraph,
    property: &P,
) -> Result<Vec<P::Value>> {
    graph
        .slots
        .iter()
        .map(|slot| match slot {
            Some(slot) => property.read(slot_payload(workspace, slot)?),
            None => Ok(property.default_value()),
        })
        .collect()
}

fn slot_payload<'w>(workspace: &'w Workspace<'_>, slot: &Slot) -> Result<&'w [u8]> {
    let object = workspace.object(slot.object)?;
    let message = object
        .messages
        .get(slot.message)
        .ok_or(Error::Invalid("a series override payload index changed"))?;
    if message.type_ != SERIES_NON_STYLE_MESSAGE_TYPE {
        return Err(Error::Invalid("a series override payload type changed"));
    }
    Ok(message.data.as_slice())
}

fn chart_payload<'w>(workspace: &'w Workspace<'_>, chart: u64) -> Result<(usize, &'w [u8])> {
    workspace.unique_payload(
        chart,
        CHART_MESSAGE_TYPE,
        "a chart must have exactly one chart payload",
    )
}

fn series_graph(workspace: &Workspace<'_>, chart: u64, series_count: usize) -> Result<SeriesGraph> {
    let (chart_message, payload) = chart_payload(workspace, chart)?;
    let references = codec::references(payload, CodecLimits::DEFAULT)?;
    let mut slots = vec![None; series_count];
    if let Some(sparse) = references.series_non_styles() {
        if usize::try_from(sparse.count()).ok() != Some(sparse.entries().len()) {
            return Err(Error::Invalid(
                "a chart's series_non_styles count disagrees with its entries",
            ));
        }
        let mut seen = HashSet::new();
        for &(index, object) in sparse.entries() {
            let index = usize::try_from(index)
                .map_err(|_| Error::Invalid("a series index exceeds usize"))?;
            if index >= series_count {
                return Err(Error::Invalid(
                    "a series override index exceeds the series count",
                ));
            }
            if !seen.insert(object) || slots[index].is_some() {
                return Err(Error::Invalid("a series override is repeated"));
            }
            let location = workspace.locate(object)?;
            let (message, _) = workspace.unique_payload(
                object,
                SERIES_NON_STYLE_MESSAGE_TYPE,
                "a series override must have exactly one payload",
            )?;
            slots[index] = Some(Slot {
                location,
                object,
                message,
            });
        }
    }
    Ok(SeriesGraph {
        chart_message,
        slots,
    })
}

/// The document stylesheet of the chart's first series theme style.
fn chart_stylesheet(workspace: &Workspace<'_>, chart: u64) -> Result<u64> {
    let (_, payload) = chart_payload(workspace, chart)?;
    let style = codec::references(payload, CodecLimits::DEFAULT)?
        .series_theme_styles()
        .first()
        .copied()
        .filter(|identifier| *identifier != 0)
        .ok_or(Error::Invalid("the chart has no series theme style"))?;
    let (_, style_payload) = workspace.unique_payload(
        style,
        SERIES_STYLE_MESSAGE_TYPE,
        "a series style must have exactly one payload",
    )?;
    stylesheet_of(style_payload)?
        .filter(|identifier| *identifier != 0)
        .ok_or(Error::Invalid(
            "the series theme style has no document stylesheet",
        ))
}

/// The stylesheet reference of a style payload's `TSS.StyleArchive` parent
/// (`Some(0)` for a reference without identifier).
fn stylesheet_of(payload: &[u8]) -> Result<Option<u64>> {
    let parent = graph::repeated(payload, STYLE_PARENT_FIELD)?;
    let [parent] = parent.as_slice() else {
        return Err(Error::Invalid("a chart style has no single style parent"));
    };
    let stylesheet = graph::repeated(parent, STYLE_STYLESHEET_FIELD)?;
    Ok(match stylesheet.last() {
        Some(reference) => Some(graph::varint(reference, REFERENCE_IDENTIFIER_FIELD)?.unwrap_or(0)),
        None => None,
    })
}

fn style_parent(stylesheet: Option<u64>) -> Vec<u8> {
    let mut style = Vec::new();
    if let Some(identifier) = stylesheet {
        let mut reference = Vec::with_capacity(11);
        push_varint_field(&mut reference, REFERENCE_IDENTIFIER_FIELD, identifier);
        push_message_field(&mut style, STYLE_STYLESHEET_FIELD, &reference);
    }
    style
}

fn override_with_parent(stylesheet: Option<u64>) -> Vec<u8> {
    let mut payload = Vec::new();
    push_message_field(&mut payload, STYLE_PARENT_FIELD, &style_parent(stylesheet));
    payload
}

fn append_capabilities(payload: &mut Vec<u8>) {
    for field in CAPABILITY_FIELDS {
        push_varint_field(payload, field, 1);
    }
}

fn canonical_empty(stylesheet: Option<u64>) -> Vec<u8> {
    let mut payload = override_with_parent(stylesheet);
    append_capabilities(&mut payload);
    payload
}

fn set_stylesheet(payload: &[u8], stylesheet: u64) -> Result<Vec<u8>> {
    if stylesheet_of(payload)? == Some(stylesheet) {
        return Ok(payload.to_vec());
    }
    Ok(wire::patch_length_delimited_field(
        payload,
        STYLE_PARENT_FIELD,
        true,
        Some(&style_parent(Some(stylesheet))),
    )?)
}

fn new_override(
    identifier: u64,
    data: Vec<u8>,
    workspace: &Workspace<'_>,
) -> Result<ArchiveObject> {
    let mut object = ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_: SERIES_NON_STYLE_MESSAGE_TYPE,
            data,
        }],
        workspace.limits(),
    )?;
    object.archive_info.message_infos[0].versions = STANDARD_MESSAGE_VERSION.to_vec().into();
    Ok(object)
}

fn allocate(next: &mut u64) -> Result<u64> {
    let identifier = *next;
    *next = next.checked_add(1).ok_or(Error::IdentifierOverflow)?;
    Ok(identifier)
}

/// How many chart payloads reference override `object`.
fn owner_count(workspace: &Workspace<'_>, object: u64) -> Result<usize> {
    let mut owners = 0usize;
    for location in workspace.locations() {
        for candidate in &workspace.archive_at(location)?.objects {
            for message in &candidate.messages {
                if message.type_ != CHART_MESSAGE_TYPE {
                    continue;
                }
                let references = codec::references(&message.data, CodecLimits::DEFAULT)?;
                if references
                    .series_non_styles()
                    .is_some_and(|sparse| sparse.entries().iter().any(|(_, id)| *id == object))
                {
                    owners += 1;
                }
            }
        }
    }
    Ok(owners)
}

struct Container {
    component: u64,
    name: String,
    locator: String,
}

fn reserve_container(workspace: &Workspace<'_>, next: &mut u64) -> Result<Container> {
    let component = allocate(next)?;
    let mut suffix: Option<u64> = None;
    loop {
        let locator = suffix.map_or_else(
            || OBJECT_CONTAINER_LOCATOR.to_owned(),
            |value| format!("{OBJECT_CONTAINER_LOCATOR}-{value}"),
        );
        let name = format!("Index/{locator}.iwa");
        if !workspace.has_member(&name) && workspace.component_for_member(&name)?.is_none() {
            return Ok(Container {
                component,
                name,
                locator,
            });
        }
        suffix = Some(match suffix {
            None => component,
            Some(value) => value.checked_add(1).ok_or(Error::IdentifierOverflow)?,
        });
    }
}

fn encode_container(objects: &[u64]) -> Vec<u8> {
    let mut payload = Vec::new();
    push_varint_field(
        &mut payload,
        CONTAINER_IDENTIFIER_FIELD,
        PRIMARY_PACKAGE_IDENTIFIER,
    );
    for &object in objects {
        let mut reference = Vec::with_capacity(11);
        push_varint_field(&mut reference, REFERENCE_IDENTIFIER_FIELD, object);
        push_message_field(&mut payload, CONTAINER_OBJECTS_FIELD, &reference);
    }
    payload
}

fn insert_container(
    workspace: &mut Workspace<'_>,
    chart_member: &str,
    container: Container,
    objects: Vec<ArchiveObject>,
) -> Result<()> {
    let identifiers = objects
        .iter()
        .map(|object| {
            object
                .archive_info
                .identifier
                .ok_or(Error::Invalid("a container member has no identifier"))
        })
        .collect::<Result<Vec<_>>>()?;
    let limits = workspace.limits();
    let mut archive_objects = Vec::with_capacity(objects.len() + 1);
    archive_objects.push(ArchiveObject::new(
        container.component,
        vec![RawMessage {
            type_: OBJECT_CONTAINER_MESSAGE_TYPE,
            data: encode_container(&identifiers),
        }],
        limits,
    )?);
    archive_objects.extend(objects);
    let archive = Archive {
        objects: archive_objects,
    };
    let anchor = if workspace.has_member(DOCUMENT_METADATA_MEMBER) {
        DOCUMENT_METADATA_MEMBER
    } else {
        METADATA_MEMBER
    };
    workspace.create_member(&container.name, archive, Some(anchor))?;
    if !workspace.has_registry() {
        return Ok(());
    }
    let mut registration = Vec::new();
    push_varint_field(&mut registration, 1, container.component);
    push_message_field(&mut registration, 2, container.locator.as_bytes());
    push_packed_field(&mut registration, 4, &OBJECT_CONTAINER_VERSION);
    push_packed_field(&mut registration, 5, &OBJECT_CONTAINER_VERSION);
    push_varint_field(&mut registration, 10, 0);
    if let Some(token) = workspace.save_token()? {
        push_varint_field(&mut registration, 12, token);
    }
    if workspace.has_data_metadata_map()? {
        push_varint_field(&mut registration, 21, PRIMARY_PACKAGE_IDENTIFIER);
    }
    workspace.add_component(container.component, &registration)?;
    workspace.add_object_uuids(container.component, &identifiers)?;
    let source = workspace
        .component_for_member(chart_member)?
        .ok_or(Error::Invalid("the chart component is not registered"))?;
    for identifier in identifiers {
        workspace.add_external_reference(source, container.component, identifier)?;
    }
    Ok(())
}

/// `(container object, payload index, member objects)` of an object
/// container component.
fn container_of(archive: &Archive) -> Result<Option<(u64, usize, Vec<u64>)>> {
    let mut found = None;
    for object in &archive.objects {
        for (index, message) in object.messages.iter().enumerate() {
            if message.type_ == OBJECT_CONTAINER_MESSAGE_TYPE {
                let identifier = object
                    .archive_info
                    .identifier
                    .ok_or(Error::Invalid("an object container has no identifier"))?;
                if found
                    .replace((identifier, index, message.data.as_slice()))
                    .is_some()
                {
                    return Err(Error::Invalid(
                        "a component holds several object containers",
                    ));
                }
            }
        }
    }
    let Some((identifier, index, payload)) = found else {
        return Ok(None);
    };
    if graph::varint(payload, CONTAINER_IDENTIFIER_FIELD)? != Some(PRIMARY_PACKAGE_IDENTIFIER) {
        return Err(Error::Invalid(
            "an object container is not in the primary package",
        ));
    }
    let mut members = Vec::new();
    let mut seen = HashSet::new();
    for reference in graph::repeated(payload, CONTAINER_OBJECTS_FIELD)? {
        let member = graph::varint(reference, REFERENCE_IDENTIFIER_FIELD)?.unwrap_or(0);
        if member == 0 || !seen.insert(member) || archive.object(member).is_none() {
            return Err(Error::Invalid("an object container member is invalid"));
        }
        members.push(member);
    }
    Ok(Some((identifier, index, members)))
}

/// Remove `objects` from their container; returns the container's
/// identifier when it became empty (the component is then deleted).
fn remove_from_container(
    workspace: &mut Workspace<'_>,
    location: Location,
    objects: &[u64],
) -> Result<Option<u64>> {
    let member = workspace.member(location).to_owned();
    let (container, message, mut members) = container_of(workspace.archive_at(location)?)?.ok_or(
        Error::Invalid("a series override container is not an object container"),
    )?;
    let requested = objects.iter().copied().collect::<HashSet<_>>();
    if requested.len() != objects.len() || !requested.iter().all(|id| members.contains(id)) {
        return Err(Error::Invalid(
            "an object container does not own every removed override",
        ));
    }
    members.retain(|identifier| !requested.contains(identifier));
    let component = workspace.component_for_member(&member)?;
    if component.is_some_and(|component| component != container) {
        return Err(Error::Invalid(
            "an object container disagrees with its component",
        ));
    }
    if members.is_empty() {
        if let Some(component) = component {
            let locator = workspace
                .component_locator(component)?
                .ok_or(Error::Invalid(
                    "an object container component has no locator",
                ))?;
            workspace.remove_component(component, &locator)?;
        }
        workspace.delete_member(location)?;
        return Ok(Some(container));
    }
    if let Some(component) = component {
        let registered = workspace.component_uuid_objects(component)?;
        let bound = objects
            .iter()
            .copied()
            .filter(|identifier| registered.contains(identifier))
            .collect::<Vec<_>>();
        workspace.remove_object_uuids(component, &bound)?;
        for &identifier in objects {
            workspace.remove_external_references_to(component, identifier)?;
        }
    }
    for &identifier in objects {
        workspace.remove(identifier)?;
    }
    workspace.replace_payload(container, message, encode_container(&members))?;
    Ok(None)
}

fn style_member(workspace: &Workspace<'_>, owner: Location, styles: &[u64]) -> Result<()> {
    let unique = styles.iter().copied().collect::<HashSet<_>>();
    if unique.len() != styles.len() || unique.contains(&0) {
        return Err(Error::Invalid(
            "chart style identifiers must be unique and nonzero",
        ));
    }
    let archive = workspace.archive_at(owner)?;
    for &style in styles {
        let object = archive.object(style).ok_or(Error::Invalid(
            "a chart style is not in its owner component",
        ))?;
        let payloads = object
            .messages
            .iter()
            .filter(|message| CHART_STYLE_MESSAGE_TYPES.contains(&message.type_))
            .count();
        if payloads != 1 {
            return Err(Error::Invalid(
                "a chart style must have exactly one style payload",
            ));
        }
    }
    Ok(())
}

fn register_styles(
    workspace: &mut Workspace<'_>,
    stylesheet: u64,
    owner: Location,
    styles: &[u64],
) -> Result<()> {
    style_member(workspace, owner, styles)?;
    let sheet_location = workspace.locate(stylesheet)?;
    let owner_member = workspace.member(owner).to_owned();
    let sheet_member = workspace.member(sheet_location).to_owned();
    let owner_component = workspace.component_for_member(&owner_member)?;
    let sheet_component = workspace.component_for_member(&sheet_member)?;
    let (message, payload) = workspace.unique_payload(
        stylesheet,
        STYLESHEET_MESSAGE_TYPE,
        "a stylesheet must have exactly one payload",
    )?;
    let mut data = payload.to_vec();
    let registered = graph::repeated(&data, STYLESHEET_STYLES_FIELD)?
        .into_iter()
        .map(|reference| graph::varint(reference, REFERENCE_IDENTIFIER_FIELD))
        .collect::<Result<Vec<_>>>()?;
    for &style in styles {
        if registered.contains(&Some(style)) {
            return Err(Error::Invalid("the stylesheet already lists a new style"));
        }
        let mut reference = Vec::with_capacity(11);
        push_varint_field(&mut reference, REFERENCE_IDENTIFIER_FIELD, style);
        data = wire::append_repeated_length_delimited_field(
            &data,
            STYLESHEET_STYLES_FIELD,
            &reference,
        )?;
    }
    workspace.replace_payload(stylesheet, message, data)?;
    let info = workspace
        .object_mut(stylesheet)?
        .archive_info
        .message_infos
        .get_mut(message)
        .ok_or(Error::Invalid("a stylesheet payload has no metadata"))?;
    for &style in styles {
        if info.object_references.contains(&style) {
            return Err(Error::Invalid(
                "the stylesheet metadata already references a new style",
            ));
        }
        info.object_references.push(style);
    }
    if let (Some(owner), Some(sheet)) = (owner_component, sheet_component)
        && owner != sheet
    {
        workspace.add_external_reference(owner, sheet, stylesheet)?;
        for &style in styles {
            workspace.add_external_reference(sheet, owner, style)?;
        }
    }
    Ok(())
}

fn unregister_styles(
    workspace: &mut Workspace<'_>,
    stylesheet: u64,
    owner: Location,
    styles: &[u64],
) -> Result<()> {
    let unique = styles.iter().copied().collect::<HashSet<_>>();
    if unique.len() != styles.len() || unique.contains(&0) {
        return Err(Error::Invalid(
            "chart style identifiers must be unique and nonzero",
        ));
    }
    let sheet_location = workspace.locate(stylesheet)?;
    let local = sheet_location == owner;
    let owner_member = workspace.member(owner).to_owned();
    let sheet_member = workspace.member(sheet_location).to_owned();
    let owner_component = workspace.component_for_member(&owner_member)?;
    let sheet_component = workspace.component_for_member(&sheet_member)?;
    let (message, payload) = workspace.unique_payload(
        stylesheet,
        STYLESHEET_MESSAGE_TYPE,
        "a stylesheet must have exactly one payload",
    )?;
    let mut kept = Vec::new();
    let mut counts = BTreeMap::<u64, usize>::new();
    for reference in graph::repeated(payload, STYLESHEET_STYLES_FIELD)? {
        let identifier = graph::varint(reference, REFERENCE_IDENTIFIER_FIELD)?.unwrap_or(0);
        if unique.contains(&identifier) {
            *counts.entry(identifier).or_default() += 1;
        } else {
            kept.push(reference.to_vec());
        }
    }
    if styles.iter().any(|style| counts.get(style) != Some(&1)) {
        return Err(Error::Invalid(
            "the stylesheet must list each removed style once",
        ));
    }
    let data =
        wire::rewrite_repeated_length_delimited_fields(payload, STYLESHEET_STYLES_FIELD, &kept)?;
    workspace.replace_payload(stylesheet, message, data)?;
    let info = workspace
        .object_mut(stylesheet)?
        .archive_info
        .message_infos
        .get_mut(message)
        .ok_or(Error::Invalid("a stylesheet payload has no metadata"))?;
    for &style in styles {
        let listed = info
            .object_references
            .iter()
            .filter(|identifier| **identifier == style)
            .count();
        if listed > 1 || (!local && listed != 1) {
            return Err(Error::Invalid(
                "the stylesheet metadata has invalid style ownership",
            ));
        }
        info.object_references
            .retain(|identifier| *identifier != style);
        for field in &mut info.field_infos {
            field
                .object_references
                .retain(|identifier| *identifier != style);
        }
    }
    if let (Some(owner), Some(sheet)) = (owner_component, sheet_component)
        && owner != sheet
    {
        for &style in styles {
            workspace.remove_external_reference(sheet, owner, style)?;
        }
    }
    Ok(())
}

fn patch_chart_references(
    workspace: &mut Workspace<'_>,
    chart: u64,
    message: usize,
    previous: &[u64],
    identifiers: &[Option<u64>],
) -> Result<()> {
    let entries = identifiers
        .iter()
        .enumerate()
        .filter_map(|(index, identifier)| identifier.map(|identifier| (index, identifier)))
        .map(|(index, identifier)| {
            u32::try_from(index)
                .map(|index| (index, identifier))
                .map_err(|_| Error::Invalid("a series index exceeds u32"))
        })
        .collect::<Result<Vec<_>>>()?;
    let payload = workspace.object(chart)?.messages[message].data.as_slice();
    let data = codec::rewrite_series_non_styles(payload, &entries, CodecLimits::DEFAULT)?
        .into_bytes()
        .into_owned();
    workspace.replace_payload(chart, message, data)?;
    let info = workspace
        .object_mut(chart)?
        .archive_info
        .message_infos
        .get_mut(message)
        .ok_or(Error::Invalid("a chart payload has no metadata"))?;
    info.object_references
        .retain(|identifier| !previous.contains(identifier));
    for identifier in identifiers.iter().copied().flatten() {
        if !info.object_references.contains(&identifier) {
            info.object_references.push(identifier);
        }
    }
    Ok(())
}

/// Write `property` for every series of chart drawable `chart`, in series
/// order (`expected.len()` is the series count). Returns `false` when every
/// series already has its value (nothing is staged).
///
/// # Errors
///
/// An empty `expected`; an invalid chart, sparse array, override,
/// stylesheet, object container or registry; a value the property rejects;
/// a write that does not read back.
#[allow(
    clippy::too_many_lines,
    reason = "one transaction in the migration host's order, kept in one place for review"
)]
pub fn set_values<P: Property>(
    workspace: &mut Workspace<'_>,
    chart: u64,
    property: &P,
    expected: &[P::Value],
) -> Result<bool> {
    if expected.is_empty() {
        return Err(Error::Invalid("a chart series write names no series"));
    }
    let base = property.base();
    let default = property.default_value();
    let graph = series_graph(workspace, chart, expected.len())?;
    let stylesheet = chart_stylesheet(workspace, chart)?;
    let current = read_values(workspace, &graph, property)?;
    if current == expected {
        return Ok(false);
    }
    let chart_location = workspace.locate(chart)?;
    let chart_member = workspace.member(chart_location).to_owned();

    let styled_empty = canonical_empty(Some(stylesheet));
    let unstyled_empty = canonical_empty(None);
    let unstyled_base = override_with_parent(None);
    let new_base = match base {
        Base::Styled => styled_empty.as_slice(),
        Base::Unstyled => unstyled_base.as_slice(),
    };
    let mut next = workspace.next_object_identifier()?;
    let mut final_ids = graph
        .slots
        .iter()
        .map(|slot| slot.as_ref().map(|slot| slot.object))
        .collect::<Vec<_>>();
    let mut updates = Vec::new();
    let mut removals = Vec::new();
    let mut styled_creations = Vec::new();
    let mut chart_created = Vec::new();
    let mut unstyled_creations = Vec::new();
    let mut container = None;
    let mut created = Vec::new();
    let mut registrations = BTreeMap::<Location, Vec<u64>>::new();
    let mut unregistrations = BTreeMap::<Location, Vec<u64>>::new();

    for (index, ((slot, value), replacement)) in
        graph.slots.iter().zip(&current).zip(expected).enumerate()
    {
        if value == replacement {
            continue;
        }
        if let Some(slot) = slot {
            let payload = slot_payload(workspace, slot)?;
            let mut patched = property.patch(payload, replacement)?;
            let registered = stylesheet_of(payload)?;
            if let Some(registered) = registered
                && registered != stylesheet
            {
                return Err(Error::Invalid(
                    "a series override belongs to another stylesheet",
                ));
            }
            let needs_registration =
                base == Base::Styled && replacement != &default && registered.is_none();
            if needs_registration {
                patched = set_stylesheet(&patched, stylesheet)?;
            }
            let is_empty =
                patched == styled_empty || patched == unstyled_empty || patched == unstyled_base;
            let owners = owner_count(workspace, slot.object)?;
            if owners == 0 {
                return Err(Error::Invalid("a series override has no owner"));
            }
            if owners > 1 {
                if is_empty {
                    final_ids[index] = None;
                    continue;
                }
                let allocation = if registered.is_some() {
                    Base::Styled
                } else {
                    base
                };
                if allocation == Base::Unstyled && container.is_none() {
                    container = Some(reserve_container(workspace, &mut next)?);
                }
                let identifier = allocate(&mut next)?;
                final_ids[index] = Some(identifier);
                let object = new_override(identifier, patched, workspace)?;
                match allocation {
                    Base::Styled => {
                        registrations
                            .entry(chart_location)
                            .or_default()
                            .push(identifier);
                        chart_created.push(identifier);
                        styled_creations.push(object);
                    },
                    Base::Unstyled => unstyled_creations.push(object),
                }
                created.push(identifier);
            } else if is_empty {
                final_ids[index] = None;
                if registered.is_some() {
                    unregistrations
                        .entry(slot.location)
                        .or_default()
                        .push(slot.object);
                }
                removals.push(slot.clone());
            } else {
                if needs_registration {
                    registrations
                        .entry(slot.location)
                        .or_default()
                        .push(slot.object);
                }
                updates.push((slot.clone(), patched));
            }
        } else if replacement != &default {
            let mut data = property.patch(new_base, replacement)?;
            if data.as_slice() == new_base {
                return Err(Error::Invalid("a series write produced no native override"));
            }
            if base == Base::Unstyled {
                append_capabilities(&mut data);
            }
            if base == Base::Unstyled && container.is_none() {
                container = Some(reserve_container(workspace, &mut next)?);
            }
            let identifier = allocate(&mut next)?;
            final_ids[index] = Some(identifier);
            let object = new_override(identifier, data, workspace)?;
            match base {
                Base::Styled => {
                    registrations
                        .entry(chart_location)
                        .or_default()
                        .push(identifier);
                    chart_created.push(identifier);
                    styled_creations.push(object);
                },
                Base::Unstyled => unstyled_creations.push(object),
            }
            created.push(identifier);
        }
    }

    for (slot, data) in updates {
        workspace.replace_payload(slot.object, slot.message, data)?;
    }
    for (location, styles) in &unregistrations {
        unregister_styles(workspace, stylesheet, *location, styles)?;
    }
    // Removal bookkeeping uses each override's member name before removal.
    let removal_members = removals
        .iter()
        .map(|slot| (slot.object, workspace.member(slot.location).to_owned()))
        .collect::<Vec<_>>();
    let mut contained = BTreeMap::<Location, Vec<u64>>::new();
    for slot in &removals {
        if container_of(workspace.archive_at(slot.location)?)?.is_some() {
            contained
                .entry(slot.location)
                .or_default()
                .push(slot.object);
        } else {
            workspace.remove(slot.object)?;
        }
    }
    let mut released_containers = Vec::new();
    for (location, objects) in contained {
        if let Some(identifier) = remove_from_container(workspace, location, &objects)? {
            released_containers.push(identifier);
        }
    }
    for object in styled_creations {
        workspace.insert(chart_location, object)?;
    }
    if !unstyled_creations.is_empty() {
        let container = container.take().ok_or(Error::Invalid(
            "an unstyled override lost its object container",
        ))?;
        insert_container(workspace, &chart_member, container, unstyled_creations)?;
    }
    for (location, styles) in &registrations {
        register_styles(workspace, stylesheet, *location, styles)?;
    }
    let previous = graph
        .slots
        .iter()
        .flatten()
        .map(|slot| slot.object)
        .collect::<Vec<_>>();
    patch_chart_references(workspace, chart, graph.chart_message, &previous, &final_ids)?;

    let mut unbound = BTreeMap::<u64, Vec<u64>>::new();
    for (object, member) in &removal_members {
        if let Some(component) = workspace.component_for_member(member)? {
            if workspace
                .component_uuid_objects(component)?
                .contains(object)
            {
                unbound.entry(component).or_default().push(*object);
            }
        }
    }
    for (component, objects) in unbound {
        workspace.remove_object_uuids(component, &objects)?;
    }
    if !chart_created.is_empty()
        && let Some(component) = workspace.component_for_member(&chart_member)?
    {
        workspace.add_object_uuids(component, &chart_created)?;
    }
    if let Some(last) = created.last().copied() {
        workspace.set_last_object_identifier(last)?;
    }
    let mut released = removals.iter().map(|slot| slot.object).collect::<Vec<_>>();
    released.extend(released_containers);
    if !released.is_empty() {
        workspace.release_identifier_suffix(&released)?;
    }
    if values(workspace, chart, expected.len(), property)? != expected {
        return Err(Error::Invalid("a chart series write does not read back"));
    }
    Ok(true)
}

/// One per-series property stored in the series styles (inherited through
/// the `TSS.StyleArchive.parent` chain).
pub trait StyleProperty {
    /// The stored value.
    type Item: Clone + PartialEq;

    /// The value a style payload defines itself (`None`: inherited).
    ///
    /// # Errors
    ///
    /// The payload is invalid or stores an unsupported value.
    fn read_local(&self, payload: &[u8]) -> Result<Option<Self::Item>>;

    /// Define (`Some`) or clear (`None`) the value in a style payload,
    /// keeping every other byte.
    ///
    /// # Errors
    ///
    /// The payload is invalid or the value cannot be stored.
    fn patch(&self, payload: &[u8], value: Option<&Self::Item>) -> Result<Vec<u8>>;
}

/// The effective series style of each of the first `series_count` series:
/// the private override of `series_private_styles`, else the theme style.
fn style_slots(workspace: &Workspace<'_>, chart: u64, series_count: usize) -> Result<Vec<u64>> {
    let (_, payload) = chart_payload(workspace, chart)?;
    let references = codec::references(payload, CodecLimits::DEFAULT)?;
    let mut styles = references.series_theme_styles().to_vec();
    if styles.is_empty() {
        return Err(Error::Invalid("the chart has no series styles"));
    }
    let mut seen = HashSet::new();
    if styles
        .iter()
        .any(|identifier| *identifier == 0 || !seen.insert(*identifier))
    {
        return Err(Error::Invalid(
            "a series theme style is invalid or repeated",
        ));
    }
    if let Some(private) = references.series_private_styles() {
        if usize::try_from(private.count()).ok() != Some(private.entries().len()) {
            return Err(Error::Invalid(
                "a chart's series_private_styles count disagrees with its entries",
            ));
        }
        let mut indices = HashSet::new();
        for &(index, identifier) in private.entries() {
            let index = usize::try_from(index)
                .map_err(|_| Error::Invalid("a series index exceeds usize"))?;
            if index >= styles.len() || !indices.insert(index) || !seen.insert(identifier) {
                return Err(Error::Invalid("a private series style is invalid"));
            }
            styles[index] = identifier;
        }
    }
    if styles.len() < series_count {
        return Err(Error::Invalid(
            "the chart has fewer series styles than series",
        ));
    }
    styles.truncate(series_count);
    for &style in &styles {
        workspace.unique_payload(
            style,
            SERIES_STYLE_MESSAGE_TYPE,
            "a series style must have exactly one payload",
        )?;
    }
    Ok(styles)
}

fn inherited<P: StyleProperty>(
    workspace: &Workspace<'_>,
    style: u64,
    property: &P,
) -> Result<Option<P::Item>> {
    let mut identifier = style;
    let mut visited = HashSet::new();
    loop {
        if !visited.insert(identifier) {
            return Err(Error::Invalid("a series style parent chain has a cycle"));
        }
        let (_, payload) = workspace.unique_payload(
            identifier,
            SERIES_STYLE_MESSAGE_TYPE,
            "a series style must have exactly one payload",
        )?;
        if let Some(value) = property.read_local(payload)? {
            return Ok(Some(value));
        }
        let parent = graph::repeated(payload, STYLE_PARENT_FIELD)?;
        let parent = match parent.last() {
            Some(style) => match graph::repeated(style, STYLE_PARENT_REFERENCE_FIELD)?.last() {
                Some(reference) => graph::varint(reference, REFERENCE_IDENTIFIER_FIELD)?
                    .filter(|identifier| *identifier != 0),
                None => None,
            },
            None => None,
        };
        let Some(parent) = parent else {
            return Ok(None);
        };
        identifier = parent;
    }
}

/// Read `property` (inherited) for each of the `series_count` series.
///
/// # Errors
///
/// An invalid chart or series style, or a value the property rejects.
pub fn style_values<P: StyleProperty>(
    workspace: &Workspace<'_>,
    chart: u64,
    series_count: usize,
    property: &P,
) -> Result<Vec<Option<P::Item>>> {
    style_slots(workspace, chart, series_count)?
        .into_iter()
        .map(|style| inherited(workspace, style, property))
        .collect()
}

/// Refuse a write to a series style another chart also uses.
fn ensure_exclusive(workspace: &Workspace<'_>, style: u64) -> Result<()> {
    let mut owners = 0usize;
    for location in workspace.locations() {
        for candidate in &workspace.archive_at(location)?.objects {
            for message in &candidate.messages {
                if message.type_ != CHART_MESSAGE_TYPE {
                    continue;
                }
                let references = codec::references(&message.data, CodecLimits::DEFAULT)?;
                let private = references
                    .series_private_styles()
                    .is_some_and(|sparse| sparse.entries().iter().any(|(_, id)| *id == style));
                if private || references.series_theme_styles().contains(&style) {
                    owners += 1;
                }
            }
        }
    }
    if owners != 1 {
        return Err(Error::Invalid("a series style is shared by several charts"));
    }
    Ok(())
}

/// Write `property` in the effective series style of every series whose
/// inherited value differs (`expected.len()` is the series count). Returns
/// `false` when nothing changes.
///
/// # Errors
///
/// An invalid chart or series style, a style another chart shares, a value
/// the property rejects, or a write that does not read back.
pub fn set_style_values<P: StyleProperty>(
    workspace: &mut Workspace<'_>,
    chart: u64,
    property: &P,
    expected: &[Option<P::Item>],
) -> Result<bool> {
    let styles = style_slots(workspace, chart, expected.len())?;
    let mut changed = false;
    for (&style, replacement) in styles.iter().zip(expected) {
        if inherited(workspace, style, property)? == *replacement {
            continue;
        }
        ensure_exclusive(workspace, style)?;
        let (message, payload) = workspace.unique_payload(
            style,
            SERIES_STYLE_MESSAGE_TYPE,
            "a series style must have exactly one payload",
        )?;
        let data = property.patch(payload, replacement.as_ref())?;
        workspace.replace_payload(style, message, data)?;
        changed = true;
    }
    if changed && style_values(workspace, chart, expected.len(), property)? != expected {
        return Err(Error::Invalid("a series style write does not read back"));
    }
    Ok(changed)
}

/// Write the data symbol of every series: visibility and shape in the
/// overrides, size in the series styles (`expected.len()` is the series
/// count; `None` hides the symbol). Returns `false` when nothing changes.
///
/// # Errors
///
/// The [`set_values`] and [`set_style_values`] errors.
pub fn set_symbols(
    workspace: &mut Workspace<'_>,
    chart: u64,
    family: litchi_iwa_common::chart::series::SymbolFamily,
    expected: &[Option<litchi_iwa_common::chart::series::Symbol>],
) -> Result<bool> {
    let marks = expected
        .iter()
        .map(|symbol| {
            symbol.map_or(property::Mark::HIDDEN, |symbol| property::Mark {
                visible: true,
                shape: symbol.shape,
            })
        })
        .collect::<Vec<_>>();
    let marked = set_values(workspace, chart, &property::SymbolMarks(family), &marks)?;
    let sizes = expected
        .iter()
        .map(|symbol| symbol.and_then(|symbol| symbol.size))
        .collect::<Vec<_>>();
    let sized = set_style_values(workspace, chart, &property::SymbolSizes(family), &sizes)?;
    if symbols(workspace, chart, family, expected.len())? != expected {
        return Err(Error::Invalid("a data-symbol write does not read back"));
    }
    Ok(marked || sized)
}

/// Read the data symbol of each of the `series_count` series.
///
/// # Errors
///
/// The [`values`] and [`style_values`] errors.
pub fn symbols(
    workspace: &Workspace<'_>,
    chart: u64,
    family: litchi_iwa_common::chart::series::SymbolFamily,
    series_count: usize,
) -> Result<Vec<Option<litchi_iwa_common::chart::series::Symbol>>> {
    let marks = values(
        workspace,
        chart,
        series_count,
        &property::SymbolMarks(family),
    )?;
    let sizes = style_values(
        workspace,
        chart,
        series_count,
        &property::SymbolSizes(family),
    )?;
    Ok(marks
        .into_iter()
        .zip(sizes)
        .map(|(mark, size)| {
            mark.visible
                .then_some(litchi_iwa_common::chart::series::Symbol {
                    shape: mark.shape,
                    size,
                })
        })
        .collect())
}

/// The per-series properties of the chart Format panel.
pub mod property {
    use litchi_iwa_common::chart::error_bar::{
        CustomValue, CustomValues, Direction as ErrorBarDirection, FixedValue, Percentage,
        Series as ErrorBarSeries, StandardDeviationCount,
    };
    use litchi_iwa_common::chart::kind::Kind;
    use litchi_iwa_common::chart::number_format::{LabelAffixes, NumberFormat};
    use litchi_iwa_common::chart::pie::{LabelDistance, LabelVisibility, WedgeExplosion};
    use litchi_iwa_common::chart::series::{
        ConnectionLine, ConnectionLineFamily, SymbolFamily, SymbolShape, SymbolSize, Trendline,
        TrendlineType,
    };
    use litchi_iwa_common::chart::series_labels::Visibility;
    use litchi_iwa_common::wire;
    use litchi_iwa_protos::chart_style_codec::{
        self as codec, Archive, Limits as CodecLimits, field, property as rules,
    };

    use super::{Base, Property, StyleProperty};
    use crate::graph::{Error, Result};
    use crate::number_format::{self, Fields};

    fn extension(payload: &[u8]) -> Result<Option<&[u8]>> {
        Ok(codec::decode(payload, Archive::SeriesNonStyle, CodecLimits::DEFAULT)?.extension())
    }

    fn strict_varint(extension: &[u8], number: u32) -> Result<Option<u64>> {
        let mut found = None;
        for candidate in crate::graph::fields(extension)? {
            if candidate.number() != number {
                continue;
            }
            if found.is_some() {
                return Err(Error::Invalid("a singular override field is repeated"));
            }
            if candidate.wire_type() != 0 {
                return Err(Error::Invalid("an override varint has the wrong wire type"));
            }
            let (value, _) = litchi_iwa_common::decode_varint_from_bytes(
                &extension[candidate.key_end()..candidate.end()],
            )
            .map_err(|_| Error::Invalid("an override varint is malformed"))?;
            found = Some(value);
        }
        Ok(found)
    }

    fn strict_bool(extension: &[u8], number: u32) -> Result<Option<bool>> {
        strict_varint(extension, number)?
            .map(|value| {
                if value > 1 {
                    return Err(Error::Invalid("an override boolean is not 0 or 1"));
                }
                Ok(value == 1)
            })
            .transpose()
    }

    fn strict_enum(extension: &[u8], number: u32) -> Result<Option<i32>> {
        strict_varint(extension, number)?
            .map(|value| {
                i32::try_from(value).map_err(|_| Error::Invalid("an override enum exceeds i32"))
            })
            .transpose()
    }

    fn rewrite(payload: &[u8], write: &rules::PropertyWrite<'_>) -> Result<Vec<u8>> {
        Ok(codec::rewrite(
            payload,
            write.archive(),
            write.writes(),
            write.empty(),
            CodecLimits::DEFAULT,
        )?
        .into_bytes()
        .into_owned())
    }

    /// How a series connects its data points (line, radar and scatter
    /// charts).
    #[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
    pub struct ConnectionLines(pub ConnectionLineFamily);

    impl Property for ConnectionLines {
        type Value = ConnectionLine;

        fn base(&self) -> Base {
            Base::Styled
        }

        fn default_value(&self) -> ConnectionLine {
            self.0.default_value()
        }

        fn read(&self, payload: &[u8]) -> Result<ConnectionLine> {
            let Some(extension) = extension(payload)? else {
                return Ok(self.0.default_value());
            };
            let line_type = strict_enum(extension, self.0.field())?;
            let show = if self.0 == ConnectionLineFamily::Scatter2d {
                strict_bool(extension, ConnectionLineFamily::SCATTER_SHOW_FIELD)?
            } else {
                None
            };
            self.0
                .read(show, line_type)
                .map_err(|_| Error::Invalid("an unsupported stored connection line"))
        }

        fn patch(&self, payload: &[u8], value: &ConnectionLine) -> Result<Vec<u8>> {
            let family = self.0;
            family
                .validate(*value)
                .map_err(|_| Error::Invalid("the chart kind cannot hide connection lines"))?;
            let existing = extension(payload)?;
            let source = existing.unwrap_or_default();
            let type_present = strict_enum(source, family.field())?.is_some();
            let native = |value: ConnectionLine| {
                u64::try_from(value.native_value())
                    .map_err(|_| Error::Invalid("a negative connection line"))
            };
            let scatter = family == ConnectionLineFamily::Scatter2d;
            let (show, native) = match value {
                ConnectionLine::Hidden => (None, None),
                ConnectionLine::Straight if !scatter => (None, None),
                value if scatter => (Some(1), Some(native(*value)?)),
                value => (None, Some(native(*value)?)),
            };
            let mut extension = source.to_vec();
            if family == ConnectionLineFamily::Scatter2d {
                let show_present =
                    strict_bool(source, ConnectionLineFamily::SCATTER_SHOW_FIELD)?.is_some();
                extension = wire::patch_varint_field(
                    &extension,
                    ConnectionLineFamily::SCATTER_SHOW_FIELD,
                    show_present,
                    show,
                )?;
            }
            let extension =
                wire::patch_varint_field(&extension, family.field(), type_present, native)?;
            let patched = wire::patch_length_delimited_field(
                payload,
                codec::EXTENSION_FIELD,
                existing.is_some(),
                (!extension.is_empty()).then_some(extension.as_slice()),
            )?;
            if self.read(&patched)? != *value {
                return Err(Error::Invalid("a connection-line write does not read back"));
            }
            Ok(patched)
        }
    }

    /// The distance of each pie or donut wedge from the chart center.
    #[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
    pub struct WedgeExplosions;

    impl Property for WedgeExplosions {
        type Value = WedgeExplosion;

        fn base(&self) -> Base {
            Base::Unstyled
        }

        fn default_value(&self) -> WedgeExplosion {
            WedgeExplosion::ZERO
        }

        fn read(&self, payload: &[u8]) -> Result<WedgeExplosion> {
            let properties = codec::decode(payload, Archive::SeriesNonStyle, CodecLimits::DEFAULT)?;
            properties
                .float(field::series_non_style::PIE_WEDGE_EXPLOSION)?
                .map(WedgeExplosion::from_native)
                .transpose()
                .map_err(|_| Error::Invalid("a stored wedge explosion is out of range"))
                .map(Option::unwrap_or_default)
        }

        fn patch(&self, payload: &[u8], value: &WedgeExplosion) -> Result<Vec<u8>> {
            rewrite(
                payload,
                &rules::pie_wedge_explosion(value.native_fraction()),
            )
        }
    }

    /// The distance of each pie or donut wedge's label from the center.
    #[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
    pub struct LabelDistances;

    impl Property for LabelDistances {
        type Value = LabelDistance;

        fn base(&self) -> Base {
            Base::Unstyled
        }

        fn default_value(&self) -> LabelDistance {
            LabelDistance::DEFAULT
        }

        fn read(&self, payload: &[u8]) -> Result<LabelDistance> {
            let properties = codec::decode(payload, Archive::SeriesNonStyle, CodecLimits::DEFAULT)?;
            let current = properties.float(field::series_non_style::PIE_LABEL_EXPLOSION)?;
            let legacy = properties.float(field::series_non_style::PIE_LEGACY_LABEL_EXPLOSION)?;
            if let (Some(current), Some(legacy)) = (current, legacy)
                && current.to_bits() != legacy.to_bits()
            {
                return Err(Error::Invalid("the stored label distances disagree"));
            }
            legacy
                .or(current)
                .map(LabelDistance::from_percent)
                .transpose()
                .map_err(|_| Error::Invalid("a stored label distance is out of range"))
                .map(|value| value.unwrap_or(LabelDistance::DEFAULT))
        }

        fn patch(&self, payload: &[u8], value: &LabelDistance) -> Result<Vec<u8>> {
            rewrite(payload, &rules::pie_label_distance(value.percent()))
        }
    }

    /// A series' data-symbol visibility and shape (the override half of a
    /// data symbol; the size lives in the series style).
    #[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
    pub struct Mark {
        /// Whether the symbol is drawn.
        pub visible: bool,
        /// The marker shape.
        pub shape: SymbolShape,
    }

    impl Mark {
        /// No symbol (the native default).
        pub const HIDDEN: Self = Self {
            visible: false,
            shape: SymbolShape::Circle,
        };
    }

    /// Data-symbol visibility and shape (area, line, radar and scatter
    /// charts).
    #[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
    pub struct SymbolMarks(pub SymbolFamily);

    impl Property for SymbolMarks {
        type Value = Mark;

        fn base(&self) -> Base {
            Base::Styled
        }

        fn default_value(&self) -> Mark {
            Mark::HIDDEN
        }

        fn read(&self, payload: &[u8]) -> Result<Mark> {
            let Some(extension) = extension(payload)? else {
                return Ok(Mark::HIDDEN);
            };
            let visible = strict_bool(extension, self.0.show_field())?.unwrap_or(false);
            let shape =
                SymbolShape::from_native(strict_enum(extension, self.0.type_field())?.unwrap_or(0))
                    .map_err(|_| Error::Invalid("an unsupported stored symbol shape"))?;
            Ok(Mark { visible, shape })
        }

        fn patch(&self, payload: &[u8], value: &Mark) -> Result<Vec<u8>> {
            let existing = extension(payload)?;
            let source = existing.unwrap_or_default();
            let show_present = strict_bool(source, self.0.show_field())?.is_some();
            let type_present = strict_enum(source, self.0.type_field())?.is_some();
            let extension = wire::patch_varint_field(
                source,
                self.0.show_field(),
                show_present,
                value.visible.then_some(1),
            )?;
            let native = u64::try_from(value.shape.native_value())
                .map_err(|_| Error::Invalid("a negative symbol shape"))?;
            let extension = wire::patch_varint_field(
                &extension,
                self.0.type_field(),
                type_present,
                (value.shape != SymbolShape::Circle).then_some(native),
            )?;
            let patched = wire::patch_length_delimited_field(
                payload,
                codec::EXTENSION_FIELD,
                existing.is_some(),
                (!extension.is_empty()).then_some(extension.as_slice()),
            )?;
            if self.read(&patched)? != *value {
                return Err(Error::Invalid("a data-symbol write does not read back"));
            }
            Ok(patched)
        }
    }

    /// Data-symbol size in the series style (`None`: automatic).
    #[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
    pub struct SymbolSizes(pub SymbolFamily);

    impl StyleProperty for SymbolSizes {
        type Item = SymbolSize;

        fn read_local(&self, payload: &[u8]) -> Result<Option<SymbolSize>> {
            codec::decode(payload, Archive::SeriesStyle, CodecLimits::DEFAULT)?
                .float(self.0.size_field())?
                .map(SymbolSize::new)
                .transpose()
                .map_err(|_| Error::Invalid("a stored symbol size is not positive"))
        }

        fn patch(&self, payload: &[u8], value: Option<&SymbolSize>) -> Result<Vec<u8>> {
            rewrite(
                payload,
                &rules::symbol_size(self.0.size_field(), value.map(|size| size.points())),
            )
        }
    }

    const SHOW_TRENDLINE: u32 = 37;
    const TRENDLINE_LABEL: u32 = 54;
    const POLYNOMIAL_ORDER: u32 = 55;
    const MOVING_AVERAGE_PERIOD: u32 = 56;
    const SHOW_EQUATION: u32 = 59;
    const SHOW_LABEL: u32 = 60;
    const SHOW_R_SQUARED: u32 = 61;
    const TRENDLINE_TYPE: u32 = 62;

    fn strict_bytes(extension: &[u8], number: u32) -> Result<Option<&[u8]>> {
        let mut found = None;
        for candidate in crate::graph::fields(extension)? {
            if candidate.number() != number {
                continue;
            }
            if found.is_some() {
                return Err(Error::Invalid("a singular override field is repeated"));
            }
            if candidate.wire_type() != 2 {
                return Err(Error::Invalid("an override string has the wrong wire type"));
            }
            found = Some(&extension[candidate.payload_start()..candidate.end()]);
        }
        Ok(found)
    }

    fn patch_varint(data: &[u8], number: u32, value: Option<u64>) -> Result<Vec<u8>> {
        let present = strict_varint(data, number)?.is_some();
        Ok(wire::patch_varint_field(data, number, present, value)?)
    }

    fn patch_bytes(data: &[u8], number: u32, value: Option<&[u8]>) -> Result<Vec<u8>> {
        let present = strict_bytes(data, number)?.is_some();
        Ok(wire::patch_length_delimited_field(
            data, number, present, value,
        )?)
    }

    /// The trendline of each series.
    #[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
    pub struct Trendlines;

    impl Property for Trendlines {
        type Value = Trendline;

        fn base(&self) -> Base {
            Base::Styled
        }

        fn default_value(&self) -> Trendline {
            Trendline::none()
        }

        fn read(&self, payload: &[u8]) -> Result<Trendline> {
            let Some(extension) = extension(payload)? else {
                return Ok(Trendline::none());
            };
            let label = strict_bytes(extension, TRENDLINE_LABEL)?
                .map(|bytes| {
                    std::str::from_utf8(bytes)
                        .map(str::to_owned)
                        .map_err(|_| Error::Invalid("a trendline label is not UTF-8"))
                })
                .transpose()?;
            Trendline::from_native(
                strict_varint(extension, SHOW_TRENDLINE)?,
                strict_varint(extension, TRENDLINE_TYPE)?,
                strict_varint(extension, POLYNOMIAL_ORDER)?,
                strict_varint(extension, MOVING_AVERAGE_PERIOD)?,
                strict_varint(extension, SHOW_LABEL)?,
                strict_varint(extension, SHOW_EQUATION)?,
                strict_varint(extension, SHOW_R_SQUARED)?,
                label,
            )
            .map_err(|_| Error::Invalid("an unsupported stored trendline"))
        }

        fn patch(&self, payload: &[u8], value: &Trendline) -> Result<Vec<u8>> {
            value
                .validate()
                .map_err(|_| Error::Invalid("an invalid trendline"))?;
            let existing = extension(payload)?;
            if existing.is_none() && *value == Trendline::none() {
                return Ok(payload.to_vec());
            }
            let kind = value.trendline_type();
            let visible = kind.is_visible();
            let mut extension = existing.unwrap_or_default().to_vec();
            extension = patch_varint(&extension, SHOW_TRENDLINE, visible.then_some(1))?;
            extension = patch_varint(
                &extension,
                TRENDLINE_TYPE,
                visible.then_some(i64::from(kind.native_value()) as u64),
            )?;
            if visible {
                extension = patch_bytes(
                    &extension,
                    TRENDLINE_LABEL,
                    value.custom_name().map(str::as_bytes),
                )?;
                extension = patch_varint(
                    &extension,
                    SHOW_LABEL,
                    value.shows_name_in_legend().then_some(1),
                )?;
                match kind {
                    TrendlineType::Polynomial(order) => {
                        extension = patch_varint(
                            &extension,
                            POLYNOMIAL_ORDER,
                            Some(u64::from(order.get())),
                        )?;
                    },
                    TrendlineType::MovingAverage(period) => {
                        extension = patch_varint(
                            &extension,
                            MOVING_AVERAGE_PERIOD,
                            Some(u64::from(period.get())),
                        )?;
                    },
                    _ => {},
                }
                if kind.supports_equation() {
                    extension = patch_varint(
                        &extension,
                        SHOW_EQUATION,
                        value.shows_equation().then_some(1),
                    )?;
                    extension = patch_varint(
                        &extension,
                        SHOW_R_SQUARED,
                        value.shows_r_squared().then_some(1),
                    )?;
                }
            }
            let patched = wire::patch_length_delimited_field(
                payload,
                codec::EXTENSION_FIELD,
                existing.is_some(),
                (!extension.is_empty()).then_some(extension.as_slice()),
            )?;
            if self.read(&patched)? != *value {
                return Err(Error::Invalid("a trendline write does not read back"));
            }
            Ok(patched)
        }
    }

    const ERROR_CUSTOM_NEGATIVE: u32 = 2;
    const ERROR_CUSTOM_POSITIVE: u32 = 4;
    const ERROR_FIXED_VALUE: u32 = 6;
    const ERROR_PERCENTAGE: u32 = 8;
    const ERROR_DIRECTION: u32 = 10;
    const ERROR_STANDARD_DEVIATION: u32 = 12;
    const ERROR_VALUE_TYPE: u32 = 14;
    const SHOW_ERROR_BARS: u32 = 27;
    const CUSTOM_ARRAY_VALUE: u32 = 1;
    const NATIVE_DIRECTION_BOTH: i32 = 1;
    const NATIVE_FIXED_VALUE: i32 = 1;
    const NATIVE_PERCENTAGE: i32 = 2;
    const NATIVE_STANDARD_DEVIATION: i32 = 3;
    const NATIVE_STANDARD_ERROR: i32 = 4;
    const NATIVE_CUSTOM_VALUES: i32 = 5;

    fn strict_f32(extension: &[u8], number: u32) -> Result<Option<f32>> {
        let mut found = None;
        for candidate in crate::graph::fields(extension)? {
            if candidate.number() != number {
                continue;
            }
            if found.is_some() {
                return Err(Error::Invalid("a singular override field is repeated"));
            }
            if candidate.wire_type() != 5 {
                return Err(Error::Invalid("an override float has the wrong wire type"));
            }
            let bytes: [u8; 4] = extension[candidate.payload_start()..candidate.end()]
                .try_into()
                .map_err(|_| Error::Invalid("a truncated override float"))?;
            let value = f32::from_le_bytes(bytes);
            if !value.is_finite() {
                return Err(Error::Invalid("an override float is not finite"));
            }
            found = Some(value);
        }
        Ok(found)
    }

    fn patch_fixed32(data: &[u8], number: u32, value: Option<u32>) -> Result<Vec<u8>> {
        let present = strict_f32(data, number)?.is_some();
        Ok(wire::patch_fixed32_field(data, number, present, value)?)
    }

    fn whole(value: f32, maximum: f32) -> Result<f32> {
        if !value.is_finite() || value.fract() != 0.0 || value < 0.0 || value > maximum {
            return Err(Error::Invalid(
                "an error-bar count is not an unsigned integer",
            ));
        }
        Ok(value)
    }

    fn custom_values(extension: &[u8], number: u32) -> Result<Vec<CustomValue>> {
        let Some(nested) = strict_bytes(extension, number)? else {
            return Ok(Vec::new());
        };
        wire::repeated_fixed64_values(nested, CUSTOM_ARRAY_VALUE)?
            .into_iter()
            .map(|bits| {
                CustomValue::new(f64::from_bits(bits))
                    .map_err(|_| Error::Invalid("an invalid custom error value"))
            })
            .collect()
    }

    fn patch_custom_values(
        extension: &[u8],
        number: u32,
        values: &[CustomValue],
    ) -> Result<Vec<u8>> {
        let existing = strict_bytes(extension, number)?;
        let nested = wire::rewrite_repeated_fixed64_fields(
            existing.unwrap_or_default(),
            CUSTOM_ARRAY_VALUE,
            &values
                .iter()
                .map(|value| value.value().to_bits())
                .collect::<Vec<_>>(),
        )?;
        Ok(wire::patch_length_delimited_field(
            extension,
            number,
            existing.is_some(),
            (!nested.is_empty()).then_some(nested.as_slice()),
        )?)
    }

    /// The error bars of each series.
    #[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
    pub struct ErrorBars;

    impl Property for ErrorBars {
        type Value = ErrorBarSeries;

        fn base(&self) -> Base {
            Base::Styled
        }

        fn default_value(&self) -> ErrorBarSeries {
            ErrorBarSeries::None
        }

        fn read(&self, payload: &[u8]) -> Result<ErrorBarSeries> {
            let invalid = |_| Error::Invalid("an unsupported stored error bar");
            let Some(extension) = extension(payload)? else {
                return Ok(ErrorBarSeries::None);
            };
            let visible = strict_bool(extension, SHOW_ERROR_BARS)?.unwrap_or(false);
            let direction =
                strict_enum(extension, ERROR_DIRECTION)?.unwrap_or(NATIVE_DIRECTION_BOTH);
            let native = strict_enum(extension, ERROR_VALUE_TYPE)?.unwrap_or(NATIVE_FIXED_VALUE);
            let fixed = strict_f32(extension, ERROR_FIXED_VALUE)?;
            let percentage = strict_f32(extension, ERROR_PERCENTAGE)?;
            let deviations = strict_f32(extension, ERROR_STANDARD_DEVIATION)?;
            let negative = custom_values(extension, ERROR_CUSTOM_NEGATIVE)?;
            let positive = custom_values(extension, ERROR_CUSTOM_POSITIVE)?;
            if !visible {
                return Ok(ErrorBarSeries::None);
            }
            let direction = ErrorBarDirection::from_native(direction);
            let value = match native {
                NATIVE_FIXED_VALUE => ErrorBarSeries::FixedValue {
                    direction,
                    value: FixedValue::new(fixed.unwrap_or(FixedValue::DEFAULT.value()))
                        .map_err(invalid)?,
                },
                NATIVE_PERCENTAGE => ErrorBarSeries::Percentage {
                    direction,
                    percentage: Percentage::new(whole(
                        percentage.unwrap_or(f32::from(Percentage::DEFAULT.value())),
                        f32::from(u8::MAX),
                    )? as u8)
                    .map_err(invalid)?,
                },
                NATIVE_STANDARD_DEVIATION => ErrorBarSeries::StandardDeviation {
                    direction,
                    deviations: StandardDeviationCount::new(whole(
                        deviations.unwrap_or(StandardDeviationCount::DEFAULT.value() as f32),
                        u32::MAX as f32,
                    )? as u32)
                    .map_err(invalid)?,
                },
                NATIVE_STANDARD_ERROR => ErrorBarSeries::StandardError { direction },
                NATIVE_CUSTOM_VALUES => ErrorBarSeries::CustomValues {
                    direction,
                    values: CustomValues::from_validated(positive, negative).map_err(invalid)?,
                },
                native => ErrorBarSeries::unsupported(direction, native).map_err(invalid)?,
            };
            value.validate().map_err(invalid)?;
            Ok(value)
        }

        fn patch(&self, payload: &[u8], value: &ErrorBarSeries) -> Result<Vec<u8>> {
            value
                .validate()
                .map_err(|_| Error::Invalid("an invalid error bar"))?;
            let existing = extension(payload)?;
            if existing.is_none() && *value == ErrorBarSeries::None {
                return Ok(payload.to_vec());
            }
            let mut extension = existing.unwrap_or_default().to_vec();
            let direction = value.direction();
            extension = patch_varint(
                &extension,
                SHOW_ERROR_BARS,
                direction.is_some().then_some(1),
            )?;
            if let Some(direction) = direction {
                extension = patch_varint(
                    &extension,
                    ERROR_DIRECTION,
                    Some(i64::from(direction.native_value()) as u64),
                )?;
                extension = patch_varint(
                    &extension,
                    ERROR_VALUE_TYPE,
                    value.native_type().map(|native| i64::from(native) as u64),
                )?;
                match value {
                    ErrorBarSeries::FixedValue { value, .. } => {
                        extension = patch_fixed32(
                            &extension,
                            ERROR_FIXED_VALUE,
                            Some(value.value().to_bits()),
                        )?;
                    },
                    ErrorBarSeries::Percentage { percentage, .. } => {
                        extension = patch_fixed32(
                            &extension,
                            ERROR_PERCENTAGE,
                            Some(f32::from(percentage.value()).to_bits()),
                        )?;
                    },
                    ErrorBarSeries::StandardDeviation { deviations, .. } => {
                        extension = patch_fixed32(
                            &extension,
                            ERROR_STANDARD_DEVIATION,
                            Some((deviations.value() as f32).to_bits()),
                        )?;
                    },
                    ErrorBarSeries::CustomValues { values, .. } => {
                        extension = patch_custom_values(
                            &extension,
                            ERROR_CUSTOM_NEGATIVE,
                            values.negative(),
                        )?;
                        extension = patch_custom_values(
                            &extension,
                            ERROR_CUSTOM_POSITIVE,
                            values.positive(),
                        )?;
                    },
                    _ => {},
                }
            }
            let patched = wire::patch_length_delimited_field(
                payload,
                codec::EXTENSION_FIELD,
                existing.is_some(),
                (!extension.is_empty()).then_some(extension.as_slice()),
            )?;
            if self.read(&patched)? != *value {
                return Err(Error::Invalid("an error-bar write does not read back"));
            }
            Ok(patched)
        }
    }

    const PIE_SHOW_NAMES: u32 = 31;
    const PIE_SHOW_VALUES: u32 = 44;

    /// Data-point-name and value label visibility of each pie or donut
    /// wedge.
    #[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
    pub struct PieLabels;

    impl Property for PieLabels {
        type Value = LabelVisibility;

        fn base(&self) -> Base {
            Base::Styled
        }

        fn default_value(&self) -> LabelVisibility {
            LabelVisibility::DEFAULT
        }

        fn read(&self, payload: &[u8]) -> Result<LabelVisibility> {
            let Some(extension) = extension(payload)? else {
                return Ok(LabelVisibility::DEFAULT);
            };
            Ok(LabelVisibility::new(
                strict_bool(extension, PIE_SHOW_NAMES)?.unwrap_or(false),
                strict_bool(extension, PIE_SHOW_VALUES)?.unwrap_or(true),
            ))
        }

        fn patch(&self, payload: &[u8], value: &LabelVisibility) -> Result<Vec<u8>> {
            let existing = extension(payload)?;
            if existing.is_none() && *value == LabelVisibility::DEFAULT {
                return Ok(payload.to_vec());
            }
            let source = existing.unwrap_or_default();
            let names_present = strict_bool(source, PIE_SHOW_NAMES)?.is_some();
            let values_present = strict_bool(source, PIE_SHOW_VALUES)?.is_some();
            let extension = wire::patch_varint_field(
                source,
                PIE_SHOW_NAMES,
                names_present,
                value.data_point_names_visible().then_some(1),
            )?;
            let extension = wire::patch_varint_field(
                &extension,
                PIE_SHOW_VALUES,
                values_present,
                (!value.values_visible()).then_some(0),
            )?;
            let patched = wire::patch_length_delimited_field(
                payload,
                codec::EXTENSION_FIELD,
                existing.is_some(),
                (!extension.is_empty()).then_some(extension.as_slice()),
            )?;
            if self.read(&patched)? != *value {
                return Err(Error::Invalid("a pie-label write does not read back"));
            }
            Ok(patched)
        }
    }

    /// Value-label visibility of each series, in the chart family's field.
    #[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
    pub struct ValueLabels {
        field: u32,
        default: Visibility,
    }

    impl ValueLabels {
        /// The value-label storage of a chart kind, if it has value labels.
        #[must_use]
        pub fn for_chart_kind(kind: Kind) -> Option<Self> {
            let (field, default) = match kind {
                Kind::Area2d | Kind::Area3d | Kind::StackedArea2d | Kind::StackedArea3d => {
                    (38, Visibility::Hidden)
                },
                Kind::Column2d
                | Kind::Column3d
                | Kind::Bar2d
                | Kind::Bar3d
                | Kind::StackedColumn2d
                | Kind::StackedColumn3d
                | Kind::StackedBar2d
                | Kind::StackedBar3d
                | Kind::MultiDataColumn2d
                | Kind::MultiDataBar2d => (39, Visibility::Hidden),
                Kind::Bubble2d | Kind::MultiDataBubble2d => (40, Visibility::Hidden),
                Kind::Line2d | Kind::Line3d => (42, Visibility::Hidden),
                Kind::Mixed2d | Kind::TwoAxis2d => (43, Visibility::Hidden),
                Kind::Pie2d | Kind::Pie3d | Kind::Donut2d | Kind::Donut3d => {
                    (PIE_SHOW_VALUES, Visibility::Visible)
                },
                Kind::Radar2d => (162, Visibility::Hidden),
                Kind::Scatter2d | Kind::MultiDataScatter2d => (45, Visibility::Hidden),
                _ => return None,
            };
            Some(Self { field, default })
        }
    }

    impl Property for ValueLabels {
        type Value = Visibility;

        fn base(&self) -> Base {
            Base::Styled
        }

        fn default_value(&self) -> Visibility {
            self.default
        }

        fn read(&self, payload: &[u8]) -> Result<Visibility> {
            let Some(extension) = extension(payload)? else {
                return Ok(self.default);
            };
            Ok(strict_bool(extension, self.field)?.map_or(self.default, Visibility::from))
        }

        fn patch(&self, payload: &[u8], value: &Visibility) -> Result<Vec<u8>> {
            let existing = extension(payload)?;
            if existing.is_none() && *value == self.default {
                return Ok(payload.to_vec());
            }
            let source = existing.unwrap_or_default();
            let present = strict_bool(source, self.field)?.is_some();
            let extension = wire::patch_varint_field(
                source,
                self.field,
                present,
                (*value != self.default).then_some(u64::from(value.is_visible())),
            )?;
            let patched = wire::patch_length_delimited_field(
                payload,
                codec::EXTENSION_FIELD,
                existing.is_some(),
                (!extension.is_empty()).then_some(extension.as_slice()),
            )?;
            if self.read(&patched)? != *value {
                return Err(Error::Invalid("a value-label write does not read back"));
            }
            Ok(patched)
        }
    }

    fn supported(kind: Kind) -> bool {
        kind != Kind::Undefined && !kind.is_unsupported()
    }

    fn store_extension(
        payload: &[u8],
        existing: Option<&[u8]>,
        extension: Option<Vec<u8>>,
    ) -> Result<Vec<u8>> {
        let Some(extension) = extension else {
            return Ok(payload.to_vec());
        };
        Ok(wire::patch_length_delimited_field(
            payload,
            codec::EXTENSION_FIELD,
            existing.is_some(),
            Some(extension.as_slice()),
        )?)
    }

    /// The value-label prefix and suffix of each series.
    #[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
    pub struct ValueLabelAffixes(());

    impl ValueLabelAffixes {
        /// The property for a chart kind with value labels.
        #[must_use]
        pub fn for_chart_kind(kind: Kind) -> Option<Self> {
            supported(kind).then_some(Self(()))
        }
    }

    impl Property for ValueLabelAffixes {
        type Value = LabelAffixes;

        fn base(&self) -> Base {
            Base::Styled
        }

        fn default_value(&self) -> LabelAffixes {
            LabelAffixes::default()
        }

        fn read(&self, payload: &[u8]) -> Result<LabelAffixes> {
            number_format::affixes(extension(payload)?, Fields::SERIES_VALUE_LABEL)
        }

        fn patch(&self, payload: &[u8], value: &LabelAffixes) -> Result<Vec<u8>> {
            let existing = extension(payload)?;
            if existing.is_none() && value.is_empty() {
                return Ok(payload.to_vec());
            }
            let next = number_format::patch_affixes(
                existing.unwrap_or_default(),
                Fields::SERIES_VALUE_LABEL,
                value,
                NumberFormat::SERIES_VALUE_LABEL_NATIVE_DEFAULT,
            )?;
            let patched = store_extension(payload, existing, next)?;
            if self.read(&patched)? != *value {
                return Err(Error::Invalid(
                    "a value-label affix write does not read back",
                ));
            }
            Ok(patched)
        }
    }

    /// The value-label number format of each series.
    #[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
    pub struct ValueLabelNumberFormats(());

    impl ValueLabelNumberFormats {
        /// The property for a chart kind with value labels.
        #[must_use]
        pub fn for_chart_kind(kind: Kind) -> Option<Self> {
            supported(kind).then_some(Self(()))
        }
    }

    impl Property for ValueLabelNumberFormats {
        type Value = NumberFormat;

        fn base(&self) -> Base {
            Base::Styled
        }

        fn default_value(&self) -> NumberFormat {
            NumberFormat::SERIES_VALUE_LABEL_NATIVE_DEFAULT
        }

        fn read(&self, payload: &[u8]) -> Result<NumberFormat> {
            number_format::number_format(
                extension(payload)?,
                Fields::SERIES_VALUE_LABEL,
                NumberFormat::SERIES_VALUE_LABEL_NATIVE_DEFAULT,
            )
        }

        fn patch(&self, payload: &[u8], value: &NumberFormat) -> Result<Vec<u8>> {
            let existing = extension(payload)?;
            if existing.is_none() && *value == NumberFormat::SERIES_VALUE_LABEL_NATIVE_DEFAULT {
                return Ok(payload.to_vec());
            }
            let next = number_format::patch_number_format(
                existing.unwrap_or_default(),
                Fields::SERIES_VALUE_LABEL,
                *value,
                NumberFormat::SERIES_VALUE_LABEL_NATIVE_DEFAULT,
            )?;
            let patched = store_extension(payload, existing, next)?;
            if self.read(&patched)? != *value {
                return Err(Error::Invalid(
                    "a value-label number-format write does not read back",
                ));
            }
            Ok(patched)
        }
    }
}
