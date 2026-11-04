//! The reference-line graph engine (record 0375a): the ordered
//! `TSCH.ReferenceLineNonStyleArchive` objects a chart's `reference_lines`
//! extension (`TSCH.ChartReferenceLinesArchive`, field 10005 of the chart
//! archive) lists for its primary value axis, each with a stable UUID.
//!
//! [`set_lines`] is the host's `set_chart_reference_lines` ported operation
//! for operation over a [`Workspace`]: retained lines are patched in place
//! (an object another chart lists is refused), surplus lines are removed,
//! new lines are allocated in the chart component with a fresh UUID; then
//! the chart's extension (non-style items, the primary axis's style entries
//! trimmed to the new count), the chart message's object references, the
//! component UUID bindings and the watermark follow. New item UUIDs are
//! derived deterministically where the host draws random ones.

use std::collections::HashSet;

use litchi_iwa_common::chart::reference_line::{Kind, Line, Value};
use litchi_iwa_common::wire;
use litchi_iwa_core::{ArchiveObject, RawMessage};
use litchi_iwa_protos::chart_style_codec::{self as codec, Limits as CodecLimits};

use crate::graph::{
    self, Error, Location, Result, Workspace, push_message_field, push_varint, push_varint_field,
};

const CHART_MESSAGE_TYPE: u32 = 5_021;
/// `TSCH.ReferenceLineNonStyleArchive`.
pub const MESSAGE_TYPE: u32 = 5_031;
const MESSAGE_VERSION: [u32; 3] = [1, 0, 5];
const EXTENSION_FIELD: u32 = 10_000;
const TYPE_FIELD: u32 = 1;
const SHOW_LINE_FIELD: u32 = 2;
const SHOW_NAME_FIELD: u32 = 3;
const SHOW_VALUE_FIELD: u32 = 4;
const NAME_FIELD: u32 = 5;
const CUSTOM_VALUE_FIELD: u32 = 6;
const NATIVE_MINIMUM: i32 = 1;
const NATIVE_MAXIMUM: i32 = 2;
const NATIVE_AVERAGE: i32 = 3;
const NATIVE_MEDIAN: i32 = 4;
const NATIVE_CUSTOM: i32 = 5;
/// The most reference lines one value axis holds.
pub const MAX_LINES: usize = 5;

const NON_STYLES_MAP: u32 = 1;
const STYLES_MAP: u32 = 2;
const AXIS_ID: u32 = 1;
const AXIS_ITEMS: u32 = 2;
const AXIS_STYLES: u32 = 2;
const AXIS_TYPE: u32 = 1;
const AXIS_ORDINAL: u32 = 2;
const AXIS_TYPE_Y: u64 = 2;
const ITEM_NON_STYLE: u32 = 1;
const ITEM_UUID: u32 = 2;

#[derive(Debug, Clone)]
struct Slot {
    object: u64,
    location: Location,
    message: usize,
    uuid: (u64, u64),
    value: Line,
}

struct LineGraph {
    chart_message: usize,
    extension: Option<Vec<u8>>,
    slots: Vec<Slot>,
}

fn strict(extension: &[u8], number: u32, wire_type: u8) -> Result<Option<&[u8]>> {
    let mut found = None;
    for field in graph::fields(extension)? {
        if field.number() != number {
            continue;
        }
        if found.is_some() {
            return Err(Error::Invalid(
                "a singular reference-line field is repeated",
            ));
        }
        if field.wire_type() != wire_type {
            return Err(Error::Invalid(
                "a reference-line field has the wrong wire type",
            ));
        }
        let start = if wire_type == 2 {
            field.payload_start()
        } else {
            field.key_end()
        };
        found = Some(&extension[start..field.end()]);
    }
    Ok(found)
}

fn strict_varint(extension: &[u8], number: u32) -> Result<Option<u64>> {
    strict(extension, number, 0)?
        .map(|bytes| {
            litchi_iwa_common::decode_varint_from_bytes(bytes)
                .map(|(value, _)| value)
                .map_err(|_| Error::Invalid("a reference-line varint is malformed"))
        })
        .transpose()
}

fn strict_bool(extension: &[u8], number: u32) -> Result<Option<bool>> {
    strict_varint(extension, number)?
        .map(|value| match value {
            0 => Ok(false),
            1 => Ok(true),
            _ => Err(Error::Invalid("a reference-line bool is not 0 or 1")),
        })
        .transpose()
}

fn custom_value(bytes: &[u8]) -> Result<Value> {
    let raw =
        strict(bytes, 1, 1)?.ok_or(Error::Invalid("a custom reference value has no number"))?;
    let raw: [u8; 8] = raw
        .try_into()
        .map_err(|_| Error::Invalid("a custom reference value is truncated"))?;
    Value::new(f64::from_le_bytes(raw))
        .map_err(|_| Error::Invalid("a custom reference value is not finite"))
}

fn line_extension(payload: &[u8]) -> Result<&[u8]> {
    strict(payload, EXTENSION_FIELD, 2)?.ok_or(Error::Invalid(
        "a reference-line non-style has no generated payload",
    ))
}

/// The line a `TSCH.ReferenceLineNonStyleArchive` payload stores.
///
/// # Errors
///
/// A malformed payload, a hidden line or an invalid value.
pub fn read(payload: &[u8]) -> Result<Line> {
    let invalid = |_| Error::Invalid("an invalid stored reference line");
    let extension = line_extension(payload)?;
    let native = match strict_varint(extension, TYPE_FIELD)? {
        Some(raw) => i32::try_from(raw as i64)
            .map_err(|_| Error::Invalid("a reference-line type exceeds i32"))?,
        None => 0,
    };
    let custom = strict(extension, CUSTOM_VALUE_FIELD, 2)?
        .map(custom_value)
        .transpose()?;
    let known = |kind: Kind| {
        if custom.is_some() {
            return Err(Error::Invalid(
                "a reference line stores an unexpected custom value",
            ));
        }
        Ok(kind)
    };
    let kind = match native {
        NATIVE_MINIMUM => known(Kind::Minimum)?,
        NATIVE_MAXIMUM => known(Kind::Maximum)?,
        NATIVE_AVERAGE => known(Kind::Average)?,
        NATIVE_MEDIAN => known(Kind::Median)?,
        NATIVE_CUSTOM => {
            Kind::Custom(custom.ok_or(Error::Invalid("a custom reference line has no value"))?)
        },
        native => Kind::unsupported(native, custom).map_err(invalid)?,
    };
    if !strict_bool(extension, SHOW_LINE_FIELD)?.unwrap_or(true) {
        return Err(Error::Invalid(
            "a stored reference line is hidden instead of removed",
        ));
    }
    let show_name = strict_bool(extension, SHOW_NAME_FIELD)?.unwrap_or(true);
    let show_value =
        strict_bool(extension, SHOW_VALUE_FIELD)?.unwrap_or_else(|| kind.default_shows_value());
    let mut line = Line::from_kind(kind)
        .map_err(invalid)?
        .with_name_visibility(show_name)
        .with_value_visibility(show_value);
    if let Some(name) = strict(extension, NAME_FIELD, 2)? {
        let name = std::str::from_utf8(name)
            .map_err(|_| Error::Invalid("a reference-line name is not UTF-8"))?;
        line = line.try_with_name(name).map_err(invalid)?;
    }
    line.validate().map_err(invalid)?;
    Ok(line)
}

fn number_archive(value: Value) -> Vec<u8> {
    let mut bytes = Vec::with_capacity(9);
    push_varint(&mut bytes, (1 << 3) | 1);
    bytes.extend_from_slice(&value.value().to_le_bytes());
    bytes
}

/// The payload a new reference-line object stores for `line` (an empty
/// style parent and the generated extension, as the host writes it).
///
/// # Errors
///
/// An invalid line.
pub fn encode(line: &Line) -> Result<Vec<u8>> {
    line.validate()
        .map_err(|_| Error::Invalid("an invalid reference line"))?;
    let kind = line.kind();
    let mut extension = Vec::new();
    push_varint_field(
        &mut extension,
        TYPE_FIELD,
        i64::from(kind.native_value()) as u64,
    );
    push_varint_field(&mut extension, SHOW_LINE_FIELD, 1);
    push_varint_field(
        &mut extension,
        SHOW_NAME_FIELD,
        u64::from(line.shows_name()),
    );
    push_varint_field(
        &mut extension,
        SHOW_VALUE_FIELD,
        u64::from(line.shows_value()),
    );
    push_message_field(&mut extension, NAME_FIELD, line.name().as_bytes());
    if let Some(value) = kind.custom_value() {
        push_message_field(&mut extension, CUSTOM_VALUE_FIELD, &number_archive(value));
    }
    let mut payload = vec![0x0a, 0x00];
    push_message_field(&mut payload, EXTENSION_FIELD, &extension);
    if read(&payload)? != *line {
        return Err(Error::Invalid(
            "a reference-line encoding does not read back",
        ));
    }
    Ok(payload)
}

fn patch_varint(data: &[u8], number: u32, value: Option<u64>) -> Result<Vec<u8>> {
    let present = strict_varint(data, number)?.is_some();
    Ok(wire::patch_varint_field(data, number, present, value)?)
}

fn patch_message(data: &[u8], number: u32, value: Option<&[u8]>) -> Result<Vec<u8>> {
    let present = strict(data, number, 2)?.is_some();
    Ok(wire::patch_length_delimited_field(
        data, number, present, value,
    )?)
}

/// Rewrite a stored reference-line payload to `line`, keeping every other
/// byte (the host's `patch_reference_line`).
///
/// # Errors
///
/// A malformed payload or an invalid line.
pub fn patch(payload: &[u8], line: &Line) -> Result<Vec<u8>> {
    line.validate()
        .map_err(|_| Error::Invalid("an invalid reference line"))?;
    read(payload)?;
    let extension = line_extension(payload)?;
    let kind = line.kind();
    let mut patched = patch_varint(
        extension,
        TYPE_FIELD,
        Some(i64::from(kind.native_value()) as u64),
    )?;
    patched = patch_varint(&patched, SHOW_LINE_FIELD, Some(1))?;
    patched = patch_varint(
        &patched,
        SHOW_NAME_FIELD,
        Some(u64::from(line.shows_name())),
    )?;
    patched = patch_varint(
        &patched,
        SHOW_VALUE_FIELD,
        Some(u64::from(line.shows_value())),
    )?;
    patched = patch_message(&patched, NAME_FIELD, Some(line.name().as_bytes()))?;
    let custom = match kind.custom_value() {
        Some(value) => Some(match strict(extension, CUSTOM_VALUE_FIELD, 2)? {
            Some(existing) => {
                let present = strict(existing, 1, 1)?.is_some();
                wire::patch_fixed64_field(existing, 1, present, Some(value.value().to_bits()))?
            },
            None => number_archive(value),
        }),
        None => None,
    };
    patched = patch_message(&patched, CUSTOM_VALUE_FIELD, custom.as_deref())?;
    let output =
        wire::patch_length_delimited_field(payload, EXTENSION_FIELD, true, Some(&patched))?;
    if read(&output)? != *line {
        return Err(Error::Invalid("a reference-line patch does not read back"));
    }
    Ok(output)
}

/// One `ChartAxis…Archive` entry of the extension, kept raw.
struct AxisEntry<'a> {
    raw: &'a [u8],
    primary: bool,
}

fn is_primary(entry: &[u8]) -> Result<bool> {
    let Some(axis) = strict(entry, AXIS_ID, 2)? else {
        return Err(Error::Invalid("a reference-line axis entry has no axis"));
    };
    Ok(graph::varint(axis, AXIS_TYPE)? == Some(AXIS_TYPE_Y)
        && graph::varint(axis, AXIS_ORDINAL)?.unwrap_or(0) == 0)
}

fn axis_entries(extension: &[u8], number: u32) -> Result<Vec<AxisEntry<'_>>> {
    let entries = graph::repeated(extension, number)?
        .into_iter()
        .map(|raw| {
            Ok(AxisEntry {
                raw,
                primary: is_primary(raw)?,
            })
        })
        .collect::<Result<Vec<_>>>()?;
    if entries.iter().filter(|entry| entry.primary).count() > 1 {
        return Err(Error::Invalid(
            "a chart repeats its primary value-axis reference-line map",
        ));
    }
    Ok(entries)
}

/// `(object, uuid)` of each item of a non-style axis entry.
fn items(entry: &[u8]) -> Result<Vec<(u64, (u64, u64))>> {
    graph::repeated(entry, AXIS_ITEMS)?
        .into_iter()
        .map(|item| {
            let reference = strict(item, ITEM_NON_STYLE, 2)?
                .ok_or(Error::Invalid("a reference-line item has no object"))?;
            let uuid = strict(item, ITEM_UUID, 2)?
                .ok_or(Error::Invalid("a reference-line item has no UUID"))?;
            Ok((
                graph::varint(reference, 1)?.unwrap_or(0),
                (
                    graph::varint(uuid, 1)?.unwrap_or(0),
                    graph::varint(uuid, 2)?.unwrap_or(0),
                ),
            ))
        })
        .collect()
}

/// `(index, object, raw entry)` of a style axis entry's sparse array.
fn style_entries(entry: &[u8]) -> Result<Option<Vec<(u64, u64, Vec<u8>)>>> {
    let Some(sparse) = strict(entry, AXIS_STYLES, 2)? else {
        return Ok(None);
    };
    let entries = graph::repeated(sparse, 2)?
        .into_iter()
        .map(|raw| {
            let index = graph::varint(raw, 1)?
                .ok_or(Error::Invalid("a reference-line style entry has no index"))?;
            let reference = strict(raw, 2, 2)?.ok_or(Error::Invalid(
                "a reference-line style entry has no reference",
            ))?;
            Ok((
                index,
                graph::varint(reference, 1)?.unwrap_or(0),
                raw.to_vec(),
            ))
        })
        .collect::<Result<Vec<_>>>()?;
    Ok(Some(entries))
}

fn managed(extension: Option<&[u8]>) -> Result<Vec<u64>> {
    let mut identifiers = Vec::new();
    let Some(extension) = extension else {
        return Ok(identifiers);
    };
    for entry in graph::repeated(extension, NON_STYLES_MAP)? {
        for (object, _) in items(entry)? {
            if object != 0 && !identifiers.contains(&object) {
                identifiers.push(object);
            }
        }
    }
    for entry in graph::repeated(extension, STYLES_MAP)? {
        if let Some(entries) = style_entries(entry)? {
            for (_, object, _) in entries {
                if object != 0 && !identifiers.contains(&object) {
                    identifiers.push(object);
                }
            }
        }
    }
    Ok(identifiers)
}

fn validate_collection(lines: &[Line]) -> Result<()> {
    if lines.len() > MAX_LINES {
        return Err(Error::Invalid(
            "a value axis holds at most five reference lines",
        ));
    }
    let mut kinds = HashSet::new();
    for line in lines {
        line.validate()
            .map_err(|_| Error::Invalid("an invalid reference line"))?;
        if !kinds.insert(line.kind().native_value()) {
            return Err(Error::Invalid(
                "a reference-line type occurs more than once",
            ));
        }
    }
    Ok(())
}

fn line_graph(workspace: &Workspace<'_>, chart: u64) -> Result<LineGraph> {
    let (chart_message, payload) = workspace.unique_payload(
        chart,
        CHART_MESSAGE_TYPE,
        "a chart must have exactly one chart payload",
    )?;
    let extension = codec::reference_lines(payload, CodecLimits::DEFAULT)?.map(<[u8]>::to_vec);
    let mut slots = Vec::new();
    if let Some(extension) = extension.as_deref() {
        let primary = axis_entries(extension, NON_STYLES_MAP)?
            .into_iter()
            .find(|entry| entry.primary);
        let listed = match primary {
            Some(entry) => items(entry.raw)?,
            None => Vec::new(),
        };
        if listed.len() > MAX_LINES {
            return Err(Error::Invalid("a chart has more than five reference lines"));
        }
        let references = &workspace
            .object(chart)?
            .archive_info
            .message_infos
            .get(chart_message)
            .ok_or(Error::Invalid("a chart payload has no metadata"))?
            .object_references;
        let mut objects = HashSet::new();
        let mut uuids = HashSet::new();
        for (object, uuid) in listed {
            if object == 0 || !objects.insert(object) || !uuids.insert(uuid) {
                return Err(Error::Invalid(
                    "a chart has an invalid or repeated reference line",
                ));
            }
            if references
                .iter()
                .filter(|candidate| **candidate == object)
                .count()
                != 1
            {
                return Err(Error::Invalid(
                    "a chart's metadata must reference each line once",
                ));
            }
            let (message, line) = workspace.unique_payload(
                object,
                MESSAGE_TYPE,
                "a reference line must have exactly one non-style payload",
            )?;
            slots.push(Slot {
                object,
                location: workspace.locate(object)?,
                message,
                uuid,
                value: read(line)?,
            });
        }
        let styles = axis_entries(extension, STYLES_MAP)?;
        if let Some(entry) = styles.iter().find(|entry| entry.primary)
            && let Some(entries) = style_entries(entry.raw)?
        {
            if entries.len() > MAX_LINES {
                return Err(Error::Invalid(
                    "a chart's reference-line styles are inconsistent",
                ));
            }
            let mut indexes = HashSet::new();
            for (index, object, _) in &entries {
                if usize::try_from(*index).map_or(true, |index| index >= slots.len())
                    || *object == 0
                    || !indexes.insert(*index)
                {
                    return Err(Error::Invalid(
                        "a chart has an invalid reference-line style entry",
                    ));
                }
            }
        }
    }
    validate_collection(
        &slots
            .iter()
            .map(|slot| slot.value.clone())
            .collect::<Vec<_>>(),
    )?;
    Ok(LineGraph {
        chart_message,
        extension,
        slots,
    })
}

/// The ordered reference lines of chart drawable `chart`'s primary value
/// axis.
///
/// # Errors
///
/// An invalid chart, extension or reference-line object.
pub fn lines(workspace: &Workspace<'_>, chart: u64) -> Result<Vec<Line>> {
    Ok(line_graph(workspace, chart)?
        .slots
        .into_iter()
        .map(|slot| slot.value)
        .collect())
}

fn extension_charts(workspace: &Workspace<'_>) -> Result<Vec<Vec<u8>>> {
    let mut extensions = Vec::new();
    for location in workspace.locations() {
        for object in &workspace.archive_at(location)?.objects {
            for message in &object.messages {
                if message.type_ == CHART_MESSAGE_TYPE
                    && let Some(extension) =
                        codec::reference_lines(&message.data, CodecLimits::DEFAULT)?
                {
                    extensions.push(extension.to_vec());
                }
            }
        }
    }
    Ok(extensions)
}

fn ensure_exclusive(workspace: &Workspace<'_>, object: u64) -> Result<()> {
    let mut owners = 0usize;
    for extension in extension_charts(workspace)? {
        let mut listed = false;
        for entry in graph::repeated(&extension, NON_STYLES_MAP)? {
            listed |= items(entry)?
                .iter()
                .any(|(candidate, _)| *candidate == object);
        }
        owners += usize::from(listed);
    }
    if owners != 1 {
        return Err(Error::Invalid(
            "a reference line is shared by several charts",
        ));
    }
    Ok(())
}

fn used_uuids(workspace: &Workspace<'_>) -> Result<HashSet<(u64, u64)>> {
    let mut uuids = HashSet::new();
    for extension in extension_charts(workspace)? {
        for entry in graph::repeated(&extension, NON_STYLES_MAP)? {
            for (_, uuid) in items(entry)? {
                if !uuids.insert(uuid) {
                    return Err(Error::Invalid("a reference-line UUID is repeated"));
                }
            }
        }
    }
    Ok(uuids)
}

fn encode_item(object: u64, (lower, upper): (u64, u64)) -> Vec<u8> {
    let mut reference = Vec::with_capacity(11);
    push_varint_field(&mut reference, 1, object);
    let mut uuid = Vec::with_capacity(22);
    push_varint_field(&mut uuid, 1, lower);
    push_varint_field(&mut uuid, 2, upper);
    let mut item = Vec::with_capacity(40);
    push_message_field(&mut item, ITEM_NON_STYLE, &reference);
    push_message_field(&mut item, ITEM_UUID, &uuid);
    item
}

/// The next extension for the primary axis's `items` (`None`: no extension
/// is left), keeping every other axis entry and field raw.
fn next_extension(current: Option<&[u8]>, items: &[(u64, (u64, u64))]) -> Result<Option<Vec<u8>>> {
    let current = current.unwrap_or_default();
    let mut non_styles = Vec::new();
    let mut found = false;
    for entry in axis_entries(current, NON_STYLES_MAP)? {
        if !entry.primary {
            non_styles.push(entry.raw.to_vec());
            continue;
        }
        found = true;
        if items.is_empty() {
            continue;
        }
        let mut rebuilt = Vec::new();
        for field in graph::fields(entry.raw)? {
            if field.number() == AXIS_ITEMS {
                continue;
            }
            rebuilt.extend_from_slice(field.raw(entry.raw)?);
            if field.number() == AXIS_ID {
                for (object, uuid) in items {
                    push_message_field(&mut rebuilt, AXIS_ITEMS, &encode_item(*object, *uuid));
                }
            }
        }
        non_styles.push(rebuilt);
    }
    if !found && !items.is_empty() {
        let mut axis = Vec::new();
        push_varint_field(&mut axis, AXIS_TYPE, AXIS_TYPE_Y);
        push_varint_field(&mut axis, AXIS_ORDINAL, 0);
        let mut entry = Vec::new();
        push_message_field(&mut entry, AXIS_ID, &axis);
        for (object, uuid) in items {
            push_message_field(&mut entry, AXIS_ITEMS, &encode_item(*object, *uuid));
        }
        non_styles.push(entry);
    }
    let mut styles = Vec::new();
    for entry in axis_entries(current, STYLES_MAP)? {
        if !entry.primary {
            styles.push(entry.raw.to_vec());
            continue;
        }
        let Some(entries) = style_entries(entry.raw)? else {
            continue;
        };
        let kept = entries
            .into_iter()
            .filter(|(index, _, _)| usize::try_from(*index).is_ok_and(|index| index < items.len()))
            .collect::<Vec<_>>();
        if kept.is_empty() {
            continue;
        }
        let mut sparse = Vec::new();
        // The array length, one past the highest index (record 0431).
        let count = kept
            .iter()
            .map(|(index, _, _)| index + 1)
            .max()
            .unwrap_or(0);
        push_varint_field(&mut sparse, 1, count);
        for (_, _, raw) in &kept {
            push_message_field(&mut sparse, 2, raw);
        }
        let mut rebuilt = Vec::new();
        for field in graph::fields(entry.raw)? {
            if field.number() == AXIS_STYLES {
                push_message_field(&mut rebuilt, AXIS_STYLES, &sparse);
            } else {
                rebuilt.extend_from_slice(field.raw(entry.raw)?);
            }
        }
        styles.push(rebuilt);
    }
    let mut others = Vec::new();
    for field in graph::fields(current)? {
        if field.number() != NON_STYLES_MAP && field.number() != STYLES_MAP {
            others.extend_from_slice(field.raw(current)?);
        }
    }
    if non_styles.is_empty() && styles.is_empty() && others.is_empty() {
        return Ok(None);
    }
    let mut output = Vec::new();
    for entry in &non_styles {
        push_message_field(&mut output, NON_STYLES_MAP, entry);
    }
    for entry in &styles {
        push_message_field(&mut output, STYLES_MAP, entry);
    }
    output.extend_from_slice(&others);
    Ok(Some(output))
}

/// Replace chart drawable `chart`'s primary value-axis reference lines with
/// `expected`, in order. Returns `false` when they already are.
///
/// # Errors
///
/// More than five lines or a repeated type; an invalid chart, extension or
/// line object; a line another chart lists; an invalid registry; a write
/// that does not read back.
pub fn set_lines(workspace: &mut Workspace<'_>, chart: u64, expected: &[Line]) -> Result<bool> {
    validate_collection(expected)?;
    let graph = line_graph(workspace, chart)?;
    let current = graph
        .slots
        .iter()
        .map(|slot| slot.value.clone())
        .collect::<Vec<_>>();
    if current == expected {
        return Ok(false);
    }
    let chart_location = workspace.locate(chart)?;
    let chart_member = workspace.member(chart_location).to_owned();
    let retained = graph.slots.len().min(expected.len());
    for (slot, replacement) in graph.slots.iter().take(retained).zip(expected) {
        if &slot.value == replacement {
            continue;
        }
        ensure_exclusive(workspace, slot.object)?;
        let payload = workspace.object(slot.object)?.messages[slot.message]
            .data
            .clone();
        workspace.replace_payload(slot.object, slot.message, patch(&payload, replacement)?)?;
    }
    let removed = graph
        .slots
        .iter()
        .skip(retained)
        .cloned()
        .collect::<Vec<_>>();
    for slot in &removed {
        ensure_exclusive(workspace, slot.object)?;
    }
    let removal_members = removed
        .iter()
        .map(|slot| (slot.object, workspace.member(slot.location).to_owned()))
        .collect::<Vec<_>>();
    for slot in &removed {
        workspace.remove(slot.object)?;
    }
    let mut next = workspace.next_object_identifier()?;
    let mut uuids = used_uuids(workspace)?;
    let mut created = Vec::new();
    let mut listed = graph
        .slots
        .iter()
        .take(retained)
        .map(|slot| (slot.object, slot.uuid))
        .collect::<Vec<_>>();
    for replacement in expected.iter().skip(retained) {
        let identifier = next;
        next = next.checked_add(1).ok_or(Error::IdentifierOverflow)?;
        let uuid = workspace.fresh_uuid(identifier, &mut uuids);
        let mut object = ArchiveObject::new(
            identifier,
            vec![RawMessage {
                type_: MESSAGE_TYPE,
                data: encode(replacement)?,
            }],
            workspace.limits(),
        )?;
        object.archive_info.message_infos[0].versions = MESSAGE_VERSION.to_vec().into();
        workspace.insert(chart_location, object)?;
        created.push(identifier);
        listed.push((identifier, uuid));
    }
    let previous = managed(graph.extension.as_deref())?;
    let extension = next_extension(graph.extension.as_deref(), &listed)?;
    let following = managed(extension.as_deref())?;
    let payload = workspace.object(chart)?.messages[graph.chart_message]
        .data
        .clone();
    let data =
        codec::rewrite_reference_lines(&payload, extension.as_deref(), CodecLimits::DEFAULT)?
            .into_bytes()
            .into_owned();
    workspace.replace_payload(chart, graph.chart_message, data)?;
    let references = &mut workspace
        .object_mut(chart)?
        .archive_info
        .message_infos
        .get_mut(graph.chart_message)
        .ok_or(Error::Invalid("a chart payload has no metadata"))?
        .object_references;
    references.retain(|identifier| !previous.contains(identifier));
    for identifier in following {
        if !references.contains(&identifier) {
            references.push(identifier);
        }
    }
    for (object, member) in &removal_members {
        if let Some(component) = workspace.component_for_member(member)?
            && workspace
                .component_uuid_objects(component)?
                .contains(object)
        {
            workspace.remove_object_uuids(component, &[*object])?;
        }
    }
    if let Some(last) = created.last().copied() {
        if let Some(component) = workspace.component_for_member(&chart_member)? {
            workspace.add_object_uuids(component, &created)?;
        }
        workspace.set_last_object_identifier(last)?;
    }
    if !removed.is_empty() {
        let released = removed.iter().map(|slot| slot.object).collect::<Vec<_>>();
        workspace.release_identifier_suffix(&released)?;
    }
    if lines(workspace, chart)? != expected {
        return Err(Error::Invalid("a reference-line write does not read back"));
    }
    Ok(true)
}

/// The chart archive field holding `TSCH.ChartReferenceLinesArchive`. A
/// facade's locality check admits it beside the generated extension when a
/// reference-line write's scope names it among the chart fields.
pub const REFERENCE_LINES_FIELD: u32 = 10_005;

/// The line objects of chart drawable `chart`, with their members.
fn line_objects(workspace: &Workspace<'_>, chart: u64) -> Result<Vec<(u64, String)>> {
    Ok(line_graph(workspace, chart)?
        .slots
        .into_iter()
        .map(|slot| (slot.object, workspace.member(slot.location).to_owned()))
        .collect())
}

/// [`lines`] under one residual profile (the bounded read the format
/// facades run).
///
/// # Errors
///
/// An invalid chart, extension or reference-line object, or an exhausted
/// limit.
pub fn read_lines(
    workspace: &Workspace<'_>,
    chart: u64,
    limits: crate::series::Limits,
) -> crate::series::SymbolResult<(Vec<Line>, crate::series::Report)> {
    let mut budget = crate::series::SymbolBudget::new(limits);
    budget.work(workspace.source_work())?;
    let value = lines(workspace, chart)?;
    budget.work(value.len())?;
    Ok((value, budget.report(false)))
}

/// [`set_lines`] under one residual profile, returning whether anything
/// changed and the report with the mutation scope: the chart and every
/// retained, removed and created line object, the chart's and metadata
/// members, the chart and line message types, [`REFERENCE_LINES_FIELD`]
/// among the chart fields, the created objects planned new and the removed
/// ones planned removed.
///
/// # Errors
///
/// As [`set_lines`], or an exhausted limit.
pub fn write_lines(
    workspace: &mut Workspace<'_>,
    chart: u64,
    expected: &[Line],
    limits: crate::series::Limits,
) -> crate::series::SymbolResult<(bool, crate::series::Report)> {
    let mut budget = crate::series::SymbolBudget::new(limits);
    budget.work(workspace.source_work())?;
    let before = line_objects(workspace, chart)?;
    let chart_member = workspace.member(workspace.locate(chart)?).to_owned();
    if !set_lines(workspace, chart, expected)? {
        return Ok((false, budget.report(false)));
    }
    let after = line_objects(workspace, chart)?;
    budget.work(before.len().saturating_add(after.len()))?;
    let mut scope = crate::series::MutationScope::in_place(
        chart,
        vec![chart],
        vec![chart_member.clone()],
        vec![CHART_MESSAGE_TYPE, MESSAGE_TYPE, 11_006],
        // The chart's generated extension and reference lines; the
        // package-metadata UUID registrations and watermark.
        vec![
            EXTENSION_FIELD,
            REFERENCE_LINES_FIELD,
            1,
            2,
            3,
            6,
            8,
            10,
            11,
            12,
        ],
        vec![REFERENCE_LINES_FIELD],
    );
    for member in [graph::METADATA_MEMBER, graph::DOCUMENT_METADATA_MEMBER] {
        if workspace.has_member(member) {
            scope.allow_member(member);
        }
    }
    for (object, member) in &before {
        scope.allow(*object, member, MESSAGE_TYPE);
        if after.iter().all(|(kept, _)| kept != object) {
            scope.plan(*object, member, MESSAGE_TYPE, false);
        }
    }
    for (object, member) in &after {
        if before.iter().all(|(kept, _)| kept != object) {
            scope.plan(*object, member, MESSAGE_TYPE, true);
        }
    }
    budget.set_scope(scope);
    Ok((true, budget.report(true)))
}
