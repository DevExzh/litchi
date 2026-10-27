//! The shape-style variation graph engine (record 0375a): the copy-on-write
//! `TSWP.ShapeStyleArchive` variations through which iWork gives one shape
//! its own fill, stroke or other drawing property.
//!
//! [`set`] and [`reset`] are the host's `set_shape_*` / `reset_shape_*`
//! transactions ported operation for operation over a
//! [`Workspace`](crate::graph::Workspace):
//!
//! - a shape whose style is a minimal variation only it uses (no name, a
//!   parent, `is_variation`, its stylesheet, nothing but recognised direct
//!   overrides) has that variation rewritten in place;
//! - otherwise a new variation is allocated with the one override, parented
//!   to the current style, registered in the stylesheet (`styles` and the
//!   parent's children), given a UUID in the stylesheet component and an
//!   external reference from the drawable's component, and the shape is
//!   repointed to it;
//! - a reset that empties an exclusive variation collapses it: the shape is
//!   repointed to the parent, the variation unregistered and removed, its
//!   UUID and external references released and the watermark lowered.
//!
//! A [`ShapeProperty`] maps one value to its `TSD.ShapeStylePropertiesArchive`
//! field; [`property`] holds the fill and the stroke. Image fills, which also
//! move data references and media, stay with the format families.

use std::collections::{BTreeMap, HashSet};

use litchi_iwa_common::wire;
use litchi_iwa_core::{ArchiveObject, RawMessage};
use litchi_iwa_protos::chart_style_codec::Limits as CodecLimits;
use litchi_iwa_protos::shape_style_codec;

use crate::graph::{
    self, Error, Location, Result, Workspace, push_message_field, push_varint, push_varint_field,
};

const SHAPE_INFO_MESSAGE_TYPE: u32 = 2_011;
const SHAPE_STYLE_MESSAGE_TYPE: u32 = 2_025;
const STYLESHEET_MESSAGE_TYPE: u32 = 401;
const STANDARD_MESSAGE_VERSION: [u32; 3] = [1, 0, 5];
const MAX_INHERITANCE_DEPTH: usize = 64;

const STYLE_SUPER: u32 = 1;
const STYLE_OVERRIDE_COUNT: u32 = 10;
const STYLE_PROPERTIES: u32 = 11;
const TSS_PARENT: u32 = 3;
const TSS_VARIATION: u32 = 4;
const TSS_STYLESHEET: u32 = 5;
const SHAPE_INFO_SUPER: u32 = 1;
const DRAWABLE_STYLE: u32 = 2;
const REFERENCE_IDENTIFIER: u32 = 1;
const STYLESHEET_STYLES: u32 = 1;
const STYLESHEET_PARENT_CHILDREN: u32 = 5;
const CHILDREN_PARENT: u32 = 1;
const CHILDREN: u32 = 2;
const HEAD_LINE_END: u32 = 6;
const TAIL_LINE_END: u32 = 7;
const SHAPE_PROPERTY_FIELDS: std::ops::RangeInclusive<u32> = 1..=7;
const TEXT_PROPERTY_FIELDS: std::ops::RangeInclusive<u32> = 1..=6;

/// One drawing property of `TSD.ShapeStylePropertiesArchive`.
pub trait ShapeProperty {
    /// The semantic value.
    type Value: Clone + PartialEq;

    /// The `TSD.ShapeStylePropertiesArchive` field (a message field).
    fn field(&self) -> u32;

    /// The value of a shape whose style chain defines none.
    fn absent(&self) -> Self::Value;

    /// The value of a stored field payload.
    ///
    /// # Errors
    ///
    /// The payload is invalid or the value unsupported.
    fn read(&self, payload: &[u8]) -> Result<Self::Value>;

    /// The stored field payload of a value.
    ///
    /// # Errors
    ///
    /// The value cannot be stored by this engine.
    fn encode(&self, value: &Self::Value) -> Result<Vec<u8>>;

    /// Refuse to replace or remove a stored payload this engine cannot
    /// release (an image fill's data reference).
    ///
    /// # Errors
    ///
    /// The stored payload is not replaceable here.
    fn replaceable(&self, payload: &[u8]) -> Result<()> {
        let _ = payload;
        Ok(())
    }
}

/// One raw override: wire type and value bytes (a length-delimited payload,
/// a varint's bytes or a fixed32's four bytes).
type Raw = (u8, Vec<u8>);

/// The direct overrides of a minimal variation.
#[derive(Debug, Clone, Default)]
struct Overrides {
    shape: BTreeMap<u32, Raw>,
    text: BTreeMap<u32, Raw>,
}

impl Overrides {
    fn count(&self) -> u64 {
        (self.shape.len() + self.text.len()) as u64
    }

    fn is_empty(&self) -> bool {
        self.count() == 0
    }
}

fn push_raw(output: &mut Vec<u8>, number: u32, (wire_type, bytes): &Raw) {
    push_varint(output, (u64::from(number) << 3) | u64::from(*wire_type));
    if *wire_type == 2 {
        push_varint(output, bytes.len() as u64);
    }
    output.extend_from_slice(bytes);
}

fn raw_of(data: &[u8], field: wire::WireField) -> Result<Raw> {
    let start = if field.wire_type() == 2 {
        field.payload_start()
    } else {
        field.key_end()
    };
    Ok((
        field.wire_type(),
        data.get(start..field.end())
            .ok_or(Error::Invalid("a style field is truncated"))?
            .to_vec(),
    ))
}

fn single(data: &[u8], number: u32) -> Result<Option<&[u8]>> {
    let payloads = graph::repeated(data, number)?;
    match payloads.as_slice() {
        [] => Ok(None),
        [payload] => Ok(Some(payload)),
        _ => Err(Error::Invalid("a singular style field is repeated")),
    }
}

fn reference_identifier(reference: &[u8]) -> Result<u64> {
    graph::varint(reference, REFERENCE_IDENTIFIER)?
        .ok_or(Error::Invalid("a reference has no identifier"))
}

fn field_numbers(data: &[u8]) -> Result<Vec<u32>> {
    let mut numbers = graph::fields(data)?
        .into_iter()
        .map(wire::WireField::number)
        .collect::<Vec<_>>();
    numbers.sort_unstable();
    Ok(numbers)
}

/// The parts of a `TSWP.ShapeStyleArchive` payload the engine reads.
struct Style<'a> {
    parent: Option<u64>,
    stylesheet: Option<u64>,
    shape_properties: Option<&'a [u8]>,
}

fn style_parts(payload: &[u8]) -> Result<Style<'_>> {
    let tsd = single(payload, STYLE_SUPER)?
        .ok_or(Error::Invalid("a shape style has no drawing style"))?;
    let tss =
        single(tsd, STYLE_SUPER)?.ok_or(Error::Invalid("a shape style has no style parent"))?;
    Ok(Style {
        parent: single(tss, TSS_PARENT)?
            .map(reference_identifier)
            .transpose()?,
        stylesheet: single(tss, TSS_STYLESHEET)?
            .map(reference_identifier)
            .transpose()?,
        shape_properties: single(tsd, STYLE_PROPERTIES)?,
    })
}

/// The direct overrides when the style is a minimal variation (the host's
/// `direct_shape_style_overrides`).
fn direct_overrides(payload: &[u8]) -> Result<Option<Overrides>> {
    if field_numbers(payload)? != [STYLE_SUPER, STYLE_OVERRIDE_COUNT, STYLE_PROPERTIES] {
        return Ok(None);
    }
    let tsd = single(payload, STYLE_SUPER)?
        .ok_or(Error::Invalid("a shape style has no drawing style"))?;
    if field_numbers(tsd)? != [STYLE_SUPER, STYLE_OVERRIDE_COUNT, STYLE_PROPERTIES] {
        return Ok(None);
    }
    let tss =
        single(tsd, STYLE_SUPER)?.ok_or(Error::Invalid("a shape style has no style parent"))?;
    if field_numbers(tss)? != [TSS_PARENT, TSS_VARIATION, TSS_STYLESHEET]
        || graph::varint(tss, TSS_VARIATION)? != Some(1)
    {
        return Ok(None);
    }
    let mut overrides = Overrides::default();
    let shape = single(tsd, STYLE_PROPERTIES)?.unwrap_or_default();
    for field in graph::fields(shape)? {
        if !SHAPE_PROPERTY_FIELDS.contains(&field.number())
            || overrides
                .shape
                .insert(field.number(), raw_of(shape, field)?)
                .is_some()
        {
            return Ok(None);
        }
    }
    let text = single(payload, STYLE_PROPERTIES)?.unwrap_or_default();
    for field in graph::fields(text)? {
        if !TEXT_PROPERTY_FIELDS.contains(&field.number())
            || overrides
                .text
                .insert(field.number(), raw_of(text, field)?)
                .is_some()
        {
            return Ok(None);
        }
    }
    let count = overrides.count();
    let paired = overrides.shape.contains_key(&HEAD_LINE_END)
        == overrides.shape.contains_key(&TAIL_LINE_END);
    let semantic = count > 0
        && paired
        && graph::varint(payload, STYLE_OVERRIDE_COUNT)? == Some(count)
        && graph::varint(tsd, STYLE_OVERRIDE_COUNT)? == Some(count);
    Ok(semantic.then_some(overrides))
}

fn reference(identifier: u64) -> Vec<u8> {
    let mut reference = Vec::with_capacity(11);
    push_varint_field(&mut reference, REFERENCE_IDENTIFIER, identifier);
    reference
}

/// The variation payload a protobuf encoder writes for `overrides` (the
/// host's `shape_style_variation_object`).
fn variation_payload(parent: u64, stylesheet: u64, overrides: &Overrides) -> Result<Vec<u8>> {
    let count = overrides.count();
    if count == 0 {
        return Err(Error::Invalid("a shape-style variation needs an override"));
    }
    let mut tss = Vec::new();
    push_message_field(&mut tss, TSS_PARENT, &reference(parent));
    push_varint_field(&mut tss, TSS_VARIATION, 1);
    push_message_field(&mut tss, TSS_STYLESHEET, &reference(stylesheet));
    let mut shape = Vec::new();
    for (number, raw) in &overrides.shape {
        push_raw(&mut shape, *number, raw);
    }
    let mut tsd = Vec::new();
    push_message_field(&mut tsd, STYLE_SUPER, &tss);
    push_varint_field(&mut tsd, STYLE_OVERRIDE_COUNT, count);
    push_message_field(&mut tsd, STYLE_PROPERTIES, &shape);
    let mut text = Vec::new();
    for (number, raw) in &overrides.text {
        push_raw(&mut text, *number, raw);
    }
    let mut payload = Vec::new();
    push_message_field(&mut payload, STYLE_SUPER, &tsd);
    push_varint_field(&mut payload, STYLE_OVERRIDE_COUNT, count);
    push_message_field(&mut payload, STYLE_PROPERTIES, &text);
    Ok(payload)
}

fn shape_style_reference(workspace: &Workspace<'_>, drawable: u64) -> Result<(usize, u64)> {
    let (message, payload) = workspace.unique_payload(
        drawable,
        SHAPE_INFO_MESSAGE_TYPE,
        "a shape must have exactly one ShapeInfo payload",
    )?;
    let style = shape_style_codec::style_reference(payload, CodecLimits::DEFAULT)?
        .ok_or(Error::Invalid("a shape has no style"))?;
    Ok((message, style))
}

fn style_payload<'w>(workspace: &'w Workspace<'_>, style: u64) -> Result<(usize, &'w [u8])> {
    workspace.unique_payload(
        style,
        SHAPE_STYLE_MESSAGE_TYPE,
        "a shape style must have exactly one ShapeStyle payload",
    )
}

fn inherited<P: ShapeProperty>(
    workspace: &Workspace<'_>,
    first: u64,
    property: &P,
) -> Result<P::Value> {
    let mut visited = HashSet::new();
    let mut style = Some(first);
    for _ in 0..MAX_INHERITANCE_DEPTH {
        let Some(identifier) = style else {
            return Ok(property.absent());
        };
        if !visited.insert(identifier) {
            return Err(Error::Invalid("a shape-style parent chain has a cycle"));
        }
        let (_, payload) = style_payload(workspace, identifier)?;
        let parts = style_parts(payload)?;
        if let Some(properties) = parts.shape_properties
            && let Some(value) = single(properties, property.field())?
        {
            return property.read(value);
        }
        style = parts.parent;
    }
    Err(Error::Invalid("a shape-style parent chain is too deep"))
}

/// The effective value of `property` for shape drawable `drawable`.
///
/// # Errors
///
/// An invalid shape or style chain, or a value the property rejects.
pub fn value<P: ShapeProperty>(
    workspace: &Workspace<'_>,
    drawable: u64,
    property: &P,
) -> Result<P::Value> {
    let (_, style) = shape_style_reference(workspace, drawable)?;
    inherited(workspace, style, property)
}

/// Whether exactly one shape uses `style` and no style inherits from it.
fn is_exclusive(workspace: &Workspace<'_>, style: u64) -> Result<bool> {
    let mut shapes = 0usize;
    for location in workspace.locations() {
        for object in &workspace.archive_at(location)?.objects {
            for message in &object.messages {
                match message.type_ {
                    SHAPE_INFO_MESSAGE_TYPE => {
                        if shape_style_codec::style_reference(&message.data, CodecLimits::DEFAULT)?
                            == Some(style)
                        {
                            shapes += 1;
                        }
                    },
                    SHAPE_STYLE_MESSAGE_TYPE => {
                        if style_parts(&message.data)?.parent == Some(style) {
                            return Ok(false);
                        }
                    },
                    _ => {},
                }
            }
        }
    }
    Ok(shapes == 1)
}

struct Context {
    drawable: u64,
    drawable_location: Location,
    style: u64,
    style_location: Location,
    stylesheet: u64,
}

fn context(workspace: &Workspace<'_>, drawable: u64, style: u64) -> Result<Context> {
    let (_, payload) = style_payload(workspace, style)?;
    let stylesheet = style_parts(payload)?
        .stylesheet
        .ok_or(Error::Invalid("a shape style has no stylesheet"))?;
    let style_location = workspace.locate(style)?;
    if workspace.locate(stylesheet)? != style_location {
        return Err(Error::Invalid(
            "a shape style is not stored with its stylesheet",
        ));
    }
    Ok(Context {
        drawable,
        drawable_location: workspace.locate(drawable)?,
        style,
        style_location,
        stylesheet,
    })
}

fn repoint(workspace: &mut Workspace<'_>, drawable: u64, from: u64, to: u64) -> Result<()> {
    let (message, payload) = workspace.unique_payload(
        drawable,
        SHAPE_INFO_MESSAGE_TYPE,
        "a shape must have exactly one ShapeInfo payload",
    )?;
    let data = wire::transform_length_delimited_field(payload, SHAPE_INFO_SUPER, |shape| {
        wire::transform_length_delimited_field(shape, DRAWABLE_STYLE, |style| -> Result<Vec<u8>> {
            if reference_identifier(style)? != from {
                return Err(Error::Invalid("a shape's style reference changed"));
            }
            Ok(wire::patch_varint_field(
                style,
                REFERENCE_IDENTIFIER,
                true,
                Some(to),
            )?)
        })
    })?;
    workspace.replace_payload(drawable, message, data)?;
    let info = workspace
        .object_mut(drawable)?
        .archive_info
        .message_infos
        .get_mut(message)
        .ok_or(Error::Invalid("a ShapeInfo payload has no metadata"))?;
    let mut replaced = 0usize;
    for identifier in &mut info.object_references {
        if *identifier == from {
            *identifier = to;
            replaced += 1;
        }
    }
    for field in &mut info.field_infos {
        for identifier in &mut field.object_references {
            if *identifier == from {
                *identifier = to;
            }
        }
    }
    if replaced != 1 {
        return Err(Error::Invalid(
            "a shape's metadata must reference its style once",
        ));
    }
    Ok(())
}

fn stylesheet_message(workspace: &Workspace<'_>, stylesheet: u64) -> Result<(usize, Vec<u8>)> {
    let (message, payload) = workspace.unique_payload(
        stylesheet,
        STYLESHEET_MESSAGE_TYPE,
        "a stylesheet must have exactly one payload",
    )?;
    Ok((message, payload.to_vec()))
}

fn style_ids(data: &[u8], number: u32) -> Result<Vec<u64>> {
    graph::repeated(data, number)?
        .into_iter()
        .map(reference_identifier)
        .collect()
}

fn register(workspace: &mut Workspace<'_>, stylesheet: u64, parent: u64, style: u64) -> Result<()> {
    let (message, data) = stylesheet_message(workspace, stylesheet)?;
    let mut children_ids = Vec::new();
    for entry in graph::repeated(&data, STYLESHEET_PARENT_CHILDREN)? {
        children_ids.extend(style_ids(entry, CHILDREN)?);
    }
    if style_ids(&data, STYLESHEET_STYLES)?.contains(&style) || children_ids.contains(&style) {
        return Err(Error::Invalid("the stylesheet already lists a new style"));
    }
    let data =
        wire::append_repeated_length_delimited_field(&data, STYLESHEET_STYLES, &reference(style))?;
    let mut matches = 0usize;
    let mut entries = Vec::new();
    for entry in graph::repeated(&data, STYLESHEET_PARENT_CHILDREN)? {
        let entry_parent = single(entry, CHILDREN_PARENT)?
            .map(reference_identifier)
            .transpose()?;
        if entry_parent == Some(parent) {
            matches += 1;
            entries.push(wire::append_repeated_length_delimited_field(
                entry,
                CHILDREN,
                &reference(style),
            )?);
        } else {
            entries.push(entry.to_vec());
        }
    }
    if matches > 1 {
        return Err(Error::Invalid("the stylesheet repeats a parent style"));
    }
    if matches == 0 {
        let mut entry = Vec::new();
        push_message_field(&mut entry, CHILDREN_PARENT, &reference(parent));
        push_message_field(&mut entry, CHILDREN, &reference(style));
        entries.push(entry);
    }
    let data = wire::rewrite_repeated_length_delimited_fields(
        &data,
        STYLESHEET_PARENT_CHILDREN,
        &entries,
    )?;
    workspace.replace_payload(stylesheet, message, data)?;
    let info = workspace
        .object_mut(stylesheet)?
        .archive_info
        .message_infos
        .get_mut(message)
        .ok_or(Error::Invalid("a stylesheet payload has no metadata"))?;
    if info.object_references.contains(&style) {
        return Err(Error::Invalid(
            "the stylesheet metadata already references a new style",
        ));
    }
    info.object_references.push(style);
    Ok(())
}

fn unregister(
    workspace: &mut Workspace<'_>,
    stylesheet: u64,
    parent: u64,
    style: u64,
) -> Result<()> {
    let (message, data) = stylesheet_message(workspace, stylesheet)?;
    let mut kept = Vec::new();
    let mut listed = 0usize;
    for entry in graph::repeated(&data, STYLESHEET_STYLES)? {
        if reference_identifier(entry)? == style {
            listed += 1;
        } else {
            kept.push(entry.to_vec());
        }
    }
    if listed != 1 {
        return Err(Error::Invalid(
            "the stylesheet must list a removed style once",
        ));
    }
    let data = wire::rewrite_repeated_length_delimited_fields(&data, STYLESHEET_STYLES, &kept)?;
    let mut removed = 0usize;
    let mut entries = Vec::new();
    for entry in graph::repeated(&data, STYLESHEET_PARENT_CHILDREN)? {
        let entry_parent = single(entry, CHILDREN_PARENT)?
            .map(reference_identifier)
            .transpose()?;
        let mut children = Vec::new();
        let mut found = false;
        for child in graph::repeated(entry, CHILDREN)? {
            if reference_identifier(child)? == style {
                found = true;
                removed += 1;
            } else {
                children.push(child.to_vec());
            }
        }
        if found && entry_parent != Some(parent) {
            return Err(Error::Invalid(
                "the stylesheet maps a removed style under another parent",
            ));
        }
        if !children.is_empty() {
            entries.push(wire::rewrite_repeated_length_delimited_fields(
                entry, CHILDREN, &children,
            )?);
        }
    }
    if removed != 1 {
        return Err(Error::Invalid(
            "the stylesheet must map a removed style once",
        ));
    }
    let data = wire::rewrite_repeated_length_delimited_fields(
        &data,
        STYLESHEET_PARENT_CHILDREN,
        &entries,
    )?;
    workspace.replace_payload(stylesheet, message, data)?;
    let info = workspace
        .object_mut(stylesheet)?
        .archive_info
        .message_infos
        .get_mut(message)
        .ok_or(Error::Invalid("a stylesheet payload has no metadata"))?;
    if info
        .object_references
        .iter()
        .filter(|id| **id == style)
        .count()
        != 1
    {
        return Err(Error::Invalid(
            "the stylesheet metadata must reference a removed style once",
        ));
    }
    info.object_references.retain(|id| *id != style);
    for field in &mut info.field_infos {
        field.object_references.retain(|id| *id != style);
    }
    Ok(())
}

fn replace_variation(
    workspace: &mut Workspace<'_>,
    style: u64,
    parent: u64,
    stylesheet: u64,
    overrides: &Overrides,
) -> Result<()> {
    let data = variation_payload(parent, stylesheet, overrides)?;
    let (message, _) = style_payload(workspace, style)?;
    workspace.replace_payload(style, message, data)
}

fn insert_variation(
    workspace: &mut Workspace<'_>,
    context: &Context,
    overrides: &Overrides,
) -> Result<()> {
    let identifier = workspace.next_object_identifier()?;
    let mut object = ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_: SHAPE_STYLE_MESSAGE_TYPE,
            data: variation_payload(context.style, context.stylesheet, overrides)?,
        }],
        workspace.limits(),
    )?;
    let info = &mut object.archive_info.message_infos[0];
    info.versions = STANDARD_MESSAGE_VERSION.to_vec().into();
    info.object_references.push(context.style);
    repoint(workspace, context.drawable, context.style, identifier)?;
    register(workspace, context.stylesheet, context.style, identifier)?;
    workspace.insert(context.style_location, object)?;
    let style_member = workspace.member(context.style_location).to_owned();
    let drawable_member = workspace.member(context.drawable_location).to_owned();
    if let Some(style_component) = workspace.component_for_member(&style_member)? {
        workspace.add_object_uuids(style_component, &[identifier])?;
        if let Some(drawable_component) = workspace.component_for_member(&drawable_member)?
            && drawable_component != style_component
        {
            workspace.add_external_reference(drawable_component, style_component, identifier)?;
        }
    }
    workspace.set_last_object_identifier(identifier)
}

fn collapse(workspace: &mut Workspace<'_>, context: &Context, parent: u64) -> Result<()> {
    repoint(workspace, context.drawable, context.style, parent)?;
    unregister(workspace, context.stylesheet, parent, context.style)?;
    workspace.remove(context.style)?;
    let style_member = workspace.member(context.style_location).to_owned();
    let drawable_member = workspace.member(context.drawable_location).to_owned();
    if let Some(style_component) = workspace.component_for_member(&style_member)? {
        workspace.remove_object_uuids(style_component, &[context.style])?;
        workspace.remove_external_references_to(style_component, context.style)?;
        if let Some(drawable_component) = workspace.component_for_member(&drawable_member)?
            && drawable_component != style_component
        {
            workspace.add_external_reference(drawable_component, style_component, parent)?;
        }
    }
    workspace.release_identifier_suffix(&[context.style])
}

fn message_override(field_value: Vec<u8>) -> Raw {
    (2, field_value)
}

/// Give shape drawable `drawable` its own `value` of `property`. Returns
/// `false` when the shape already shows it.
///
/// # Errors
///
/// An invalid shape, style, stylesheet or registry; a value the property
/// cannot store; a write that does not read back.
pub fn set<P: ShapeProperty>(
    workspace: &mut Workspace<'_>,
    drawable: u64,
    property: &P,
    value: &P::Value,
) -> Result<bool> {
    let (_, style) = shape_style_reference(workspace, drawable)?;
    if inherited(workspace, style, property)? == *value {
        return Ok(false);
    }
    let encoded = property.encode(value)?;
    let context = context(workspace, drawable, style)?;
    let (_, payload) = style_payload(workspace, style)?;
    let parent = style_parts(payload)?.parent;
    let direct = direct_overrides(payload)?;
    let exclusive = direct.is_some() && is_exclusive(workspace, style)?;
    match direct {
        Some(mut overrides) if exclusive => {
            let parent = parent.ok_or(Error::Invalid("a shape-style variation has no parent"))?;
            if let Some((_, stored)) = overrides.shape.get(&property.field()) {
                property.replaceable(stored)?;
            }
            overrides
                .shape
                .insert(property.field(), message_override(encoded));
            replace_variation(workspace, style, parent, context.stylesheet, &overrides)?;
        },
        _ => {
            let mut overrides = Overrides::default();
            overrides
                .shape
                .insert(property.field(), message_override(encoded));
            insert_variation(workspace, &context, &overrides)?;
        },
    }
    if self::value(workspace, drawable, property)? != *value {
        return Err(Error::Invalid("a shape-style write does not read back"));
    }
    Ok(true)
}

/// Remove shape drawable `drawable`'s own value of `property`, restoring
/// the inherited one. Returns `false` when its style defines none.
///
/// # Errors
///
/// The [`set`] errors; a style with the property and unsupported direct
/// overrides.
pub fn reset<P: ShapeProperty>(
    workspace: &mut Workspace<'_>,
    drawable: u64,
    property: &P,
) -> Result<bool> {
    let (_, style) = shape_style_reference(workspace, drawable)?;
    let (_, payload) = style_payload(workspace, style)?;
    let parts = style_parts(payload)?;
    let defines = match parts.shape_properties {
        Some(properties) => single(properties, property.field())?.is_some(),
        None => false,
    };
    if !defines {
        return Ok(false);
    }
    if let Some(properties) = parts.shape_properties
        && let Some(stored) = single(properties, property.field())?
    {
        property.replaceable(stored)?;
    }
    let mut direct = direct_overrides(payload)?.ok_or(Error::Invalid(
        "a shape style has the property plus unsupported direct overrides",
    ))?;
    let parent = parts
        .parent
        .ok_or(Error::Invalid("a shape-style variation has no parent"))?;
    let context = context(workspace, drawable, style)?;
    let inherited_value = inherited(workspace, parent, property)?;
    if is_exclusive(workspace, style)? {
        direct.shape.remove(&property.field());
        if direct.is_empty() {
            collapse(workspace, &context, parent)?;
        } else {
            replace_variation(workspace, style, parent, context.stylesheet, &direct)?;
        }
    } else {
        let mut overrides = Overrides::default();
        overrides.shape.insert(
            property.field(),
            message_override(property.encode(&inherited_value)?),
        );
        insert_variation(workspace, &context, &overrides)?;
    }
    if value(workspace, drawable, property)? != inherited_value {
        return Err(Error::Invalid("a shape-style reset does not read back"));
    }
    Ok(true)
}

/// The fill and the stroke.
pub mod property {
    use litchi_iwa_common::shape::fill::Fill;
    use litchi_iwa_common::shape::stroke::Stroke;
    use litchi_iwa_protos::chart_style_codec::Limits as CodecLimits;
    use litchi_iwa_protos::shape_style_codec::native as stored;

    use super::ShapeProperty;
    use crate::drawing;
    use crate::graph::{Error, Result};

    /// The shape fill (`TSD.ShapeStylePropertiesArchive.fill`). Image fills
    /// move data references and media and are refused here.
    #[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
    pub struct Fills;

    impl ShapeProperty for Fills {
        type Value = Fill;

        fn field(&self) -> u32 {
            1
        }

        fn absent(&self) -> Fill {
            Fill::None
        }

        fn read(&self, payload: &[u8]) -> Result<Fill> {
            drawing::fill(&stored::fill(payload, CodecLimits::DEFAULT)?)
                .map_err(|_| Error::Invalid("an unsupported stored fill"))
        }

        fn encode(&self, value: &Fill) -> Result<Vec<u8>> {
            if matches!(value, Fill::Image(_)) {
                return Err(Error::Invalid(
                    "image fills are written by the format families (data references and media)",
                ));
            }
            Ok(drawing::encode_fill(value))
        }

        fn replaceable(&self, payload: &[u8]) -> Result<()> {
            if matches!(self.read(payload)?, Fill::Image(_)) {
                return Err(Error::Invalid(
                    "image fills are released by the format families (data references and media)",
                ));
            }
            Ok(())
        }
    }

    /// The shape stroke (`TSD.ShapeStylePropertiesArchive.stroke`; `None`
    /// is the hidden stroke).
    #[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
    pub struct Strokes;

    impl ShapeProperty for Strokes {
        type Value = Option<Stroke>;

        fn field(&self) -> u32 {
            2
        }

        fn absent(&self) -> Option<Stroke> {
            None
        }

        fn read(&self, payload: &[u8]) -> Result<Option<Stroke>> {
            drawing::stroke(&stored::stroke(payload, CodecLimits::DEFAULT)?)
                .map_err(|_| Error::Invalid("an unsupported stored stroke"))
        }

        fn encode(&self, value: &Option<Stroke>) -> Result<Vec<u8>> {
            Ok(drawing::encode_stroke(*value))
        }
    }
}
