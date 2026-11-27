//! Wire-level inventory of one slide's owned drawables.
//!
//! The inventory reads `KN.SlideArchive` owned drawables (field 7) and
//! z-order (field 42), classifies each drawable by its native message type
//! and projects geometry, text, name and lock state from the raw payloads.
//! It reads through [`Objects`], so the same code inspects the immutable
//! base package, a staged workspace and a reopened candidate.

use litchi_iwa_common::WireLimits;
use litchi_iwa_core::ArchiveObject;

use super::super::Package;
use super::super::structure::StructureError;
use super::super::structure::wire::{Fields, reference_identifier};
use super::super::structure::workspace::Workspace;
use super::{Geometry, Kind};

pub(super) const SLIDE: u32 = 5;
pub(super) const SLIDE_OWNED_DRAWABLES: u32 = 7;
pub(super) const SLIDE_Z_ORDER: u32 = 42;
pub(super) const PLACEHOLDER: u32 = 7;
pub(super) const SHAPE_INFO: u32 = 2_011;
pub(super) const STORAGE: u32 = 2_001;
pub(super) const IMAGE: u32 = 3_005;
pub(super) const MOVIE: u32 = 3_007;
pub(super) const GROUP: u32 = 3_008;
pub(super) const CONNECTION_LINE: u32 = 3_009;
pub(super) const CHART: u32 = 5_021;
pub(super) const TABLE: u32 = 6_000;

/// Read access to native objects, implemented by the base package and by a
/// staged workspace.
pub(super) trait Objects {
    fn get(&self, identifier: u64) -> Option<&ArchiveObject>;
}

impl Objects for Package {
    fn get(&self, identifier: u64) -> Option<&ArchiveObject> {
        self.object(identifier)
    }
}

impl Objects for Workspace<'_> {
    fn get(&self, identifier: u64) -> Option<&ArchiveObject> {
        self.object(identifier)
    }
}

/// One classified drawable with its private native identity.
#[derive(Debug, Clone, PartialEq)]
pub(super) struct Item {
    pub(super) identifier: u64,
    pub(super) kind: Kind,
    pub(super) message_index: usize,
    pub(super) message_type: u32,
    pub(super) geometry: Option<Geometry>,
    pub(super) storage: Option<u64>,
    pub(super) text: Option<String>,
    pub(super) name: Option<String>,
    pub(super) locked: bool,
}

/// A slide's owned drawables (source order) and z-order (back to front).
#[derive(Debug, Clone, PartialEq)]
pub(super) struct SlideDrawables {
    pub(super) slide: u64,
    pub(super) items: Vec<Item>,
    pub(super) z_order: Vec<u64>,
}

impl SlideDrawables {
    pub(super) fn item(&self, identifier: u64) -> Option<&Item> {
        self.items.iter().find(|item| item.identifier == identifier)
    }
}

/// Bounds applied while reading an inventory.
#[derive(Debug, Clone, Copy)]
pub(super) struct ReadLimits {
    pub(super) wire: WireLimits,
    pub(super) max_drawables: usize,
    pub(super) max_text_bytes: usize,
}

/// Why an inventory could not be read.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) enum ReadFailure {
    Structure(StructureError),
    Drawables { observed: usize, maximum: usize },
    TextBytes { observed: usize, maximum: usize },
}

impl From<StructureError> for ReadFailure {
    fn from(error: StructureError) -> Self {
        Self::Structure(error)
    }
}

/// The unique payload of `message_type` in `object`, with its index.
pub(super) fn unique_message(
    object: &ArchiveObject,
    message_type: u32,
) -> Result<(usize, &[u8]), StructureError> {
    let mut found = object
        .messages
        .iter()
        .enumerate()
        .filter(|(_, message)| message.type_ == message_type);
    let (index, message) = found
        .next()
        .ok_or(StructureError::Unexpected("message type"))?;
    if found.next().is_some() {
        return Err(StructureError::Unexpected("duplicate message"));
    }
    Ok((index, &message.data))
}

/// The object identifiers held by a repeated reference field.
pub(super) fn references(
    fields: &Fields,
    number: u32,
    limits: WireLimits,
) -> Result<Vec<u64>, StructureError> {
    fields
        .all(number)
        .map(|field| reference_identifier(field.message()?, limits))
        .collect()
}

/// The payload reached by following singular message fields `path`.
pub(super) fn descend(
    payload: &[u8],
    path: &[u32],
    limits: WireLimits,
) -> Result<Option<Vec<u8>>, StructureError> {
    let mut current = payload.to_vec();
    for number in path {
        let fields = Fields::parse(&current, limits)?;
        let Some(field) = fields.unique(*number)? else {
            return Ok(None);
        };
        current = field.message()?.to_vec();
    }
    Ok(Some(current))
}

/// Field path from a drawable payload to its `TSD.DrawableArchive`.
pub(super) const fn drawable_path(message_type: u32) -> &'static [u32] {
    match message_type {
        SHAPE_INFO | CONNECTION_LINE => &[1, 1],
        PLACEHOLDER => &[1, 1, 1],
        _ => &[1],
    }
}

/// Field path from a drawable payload to its `TSWP.ShapeInfoArchive`.
const fn shape_info_path(message_type: u32) -> Option<&'static [u32]> {
    match message_type {
        SHAPE_INFO => Some(&[]),
        PLACEHOLDER => Some(&[1]),
        _ => None,
    }
}

const KNOWN: [u32; 9] = [
    SHAPE_INFO,
    PLACEHOLDER,
    IMAGE,
    MOVIE,
    GROUP,
    CONNECTION_LINE,
    CHART,
    TABLE,
    3_006,
];

/// Read the owned drawables of slide object `slide`.
pub(super) fn read_slide(
    source: &impl Objects,
    slide: u64,
    limits: ReadLimits,
) -> Result<SlideDrawables, ReadFailure> {
    let wire = limits.wire;
    let object = source.get(slide).ok_or(StructureError::MissingObject)?;
    let (_, payload) = unique_message(object, SLIDE)?;
    let fields = Fields::parse(payload, wire)?;
    let owned = references(&fields, SLIDE_OWNED_DRAWABLES, wire)?;
    let z_order = references(&fields, SLIDE_Z_ORDER, wire)?;
    if owned.len() > limits.max_drawables {
        return Err(ReadFailure::Drawables {
            observed: owned.len(),
            maximum: limits.max_drawables,
        });
    }
    let mut sorted = owned.clone();
    sorted.sort_unstable();
    if sorted.windows(2).any(|pair| pair[0] == pair[1]) {
        return Err(StructureError::Unexpected("duplicate owned drawable").into());
    }
    let mut text_bytes = 0_usize;
    let mut items = Vec::with_capacity(owned.len());
    for identifier in owned {
        let item = read_item(source, identifier, wire)?;
        text_bytes = text_bytes.saturating_add(item.text.as_ref().map_or(0, String::len));
        if text_bytes > limits.max_text_bytes {
            return Err(ReadFailure::TextBytes {
                observed: text_bytes,
                maximum: limits.max_text_bytes,
            });
        }
        items.push(item);
    }
    Ok(SlideDrawables {
        slide,
        items,
        z_order,
    })
}

fn read_item(
    source: &impl Objects,
    identifier: u64,
    wire: WireLimits,
) -> Result<Item, StructureError> {
    let object = source
        .get(identifier)
        .ok_or(StructureError::MissingObject)?;
    let Some((message_index, message)) = object
        .messages
        .iter()
        .enumerate()
        .find(|(_, message)| KNOWN.contains(&message.type_))
    else {
        return Ok(Item {
            identifier,
            kind: Kind::Other,
            message_index: 0,
            message_type: object.messages.first().map_or(0, |message| message.type_),
            geometry: None,
            storage: None,
            text: None,
            name: None,
            locked: false,
        });
    };
    let message_type = message.type_;
    let payload = message.data.as_slice();
    let drawable = descend(payload, drawable_path(message_type), wire)?;
    let (geometry, locked) = match &drawable {
        Some(drawable) => drawable_facts(drawable, wire)?,
        None => (None, false),
    };
    let shape = shape_info_path(message_type)
        .map(|path| descend(payload, path, wire))
        .transpose()?
        .flatten();
    let storage = match &shape {
        Some(shape) => {
            let fields = Fields::parse(shape, wire)?;
            match fields.unique(4)?.or(fields.unique(2)?) {
                Some(field) => Some(reference_identifier(field.message()?, wire)?),
                None => None,
            }
        },
        None => None,
    };
    let text = storage
        .map(|storage| storage_text(source, storage, wire))
        .transpose()?;
    let kind = classify(message_type, payload, shape.as_deref(), wire)?;
    let name = match &shape {
        Some(shape) => user_defined_name(shape, wire)?,
        None => None,
    };
    Ok(Item {
        identifier,
        kind,
        message_index,
        message_type,
        geometry,
        storage,
        text,
        name,
        locked,
    })
}

fn classify(
    message_type: u32,
    payload: &[u8],
    shape: Option<&[u8]>,
    wire: WireLimits,
) -> Result<Kind, StructureError> {
    Ok(match message_type {
        SHAPE_INFO => {
            let fields = Fields::parse(payload, wire)?;
            if fields.unique(6)?.map(|field| field.varint()).transpose()? == Some(1) {
                Kind::TextBox
            } else if shape.is_some_and(|shape| is_straight_line(shape, wire).unwrap_or(false)) {
                Kind::Line
            } else {
                Kind::Shape
            }
        },
        PLACEHOLDER => {
            let fields = Fields::parse(payload, wire)?;
            match fields.unique(2)?.map(|field| field.varint()).transpose()? {
                Some(1) => Kind::SlideNumberPlaceholder,
                Some(2) => Kind::TitlePlaceholder,
                Some(3) => Kind::BodyPlaceholder,
                Some(4) => Kind::ObjectPlaceholder,
                _ => Kind::Placeholder,
            }
        },
        IMAGE => Kind::Image,
        MOVIE => {
            let fields = Fields::parse(payload, wire)?;
            if fields.unique(9)?.map(|field| field.varint()).transpose()? == Some(1) {
                Kind::Audio
            } else {
                Kind::Movie
            }
        },
        GROUP => Kind::Group,
        CONNECTION_LINE => Kind::Line,
        CHART => Kind::Chart,
        TABLE => Kind::Table,
        _ => Kind::Other,
    })
}

/// Whether a `TSWP.ShapeInfoArchive` path is a two-point Bézier line.
fn is_straight_line(shape: &[u8], wire: WireLimits) -> Result<bool, StructureError> {
    let Some(path) = descend(shape, &[1, 3, 5, 3], wire)? else {
        return Ok(false);
    };
    let fields = Fields::parse(&path, wire)?;
    let mut types = Vec::new();
    for element in fields.all(1) {
        let element = Fields::parse(element.message()?, wire)?;
        types.push(element.unique(1)?.map(|field| field.varint()).transpose()?);
    }
    Ok(types == [Some(1), Some(2)])
}

fn user_defined_name(shape: &[u8], wire: WireLimits) -> Result<Option<String>, StructureError> {
    let Some(path_source) = descend(shape, &[1, 3], wire)? else {
        return Ok(None);
    };
    let fields = Fields::parse(&path_source, wire)?;
    fields
        .unique(10)?
        .map(|field| {
            String::from_utf8(field.message()?.to_vec())
                .map_err(|_| StructureError::Unexpected("name UTF-8"))
        })
        .transpose()
        .map(|name| name.filter(|name| !name.is_empty()))
}

fn drawable_facts(
    drawable: &[u8],
    wire: WireLimits,
) -> Result<(Option<Geometry>, bool), StructureError> {
    let fields = Fields::parse(drawable, wire)?;
    let locked = fields.unique(5)?.map(|field| field.varint()).transpose()? == Some(1);
    let Some(geometry) = fields.unique(1)? else {
        return Ok((None, locked));
    };
    Ok((read_geometry(geometry.message()?, wire)?, locked))
}

/// Project a `TSD.GeometryArchive`; `None` without position or size.
pub(super) fn read_geometry(
    payload: &[u8],
    wire: WireLimits,
) -> Result<Option<Geometry>, StructureError> {
    let fields = Fields::parse(payload, wire)?;
    let pair = |number: u32| -> Result<Option<(f32, f32)>, StructureError> {
        let Some(field) = fields.unique(number)? else {
            return Ok(None);
        };
        let inner = Fields::parse(field.message()?, wire)?;
        let first = inner.unique(1)?.map(|field| field.float()).transpose()?;
        let second = inner.unique(2)?.map(|field| field.float()).transpose()?;
        Ok(first.zip(second))
    };
    let (Some((x, y)), Some((width, height))) = (pair(1)?, pair(2)?) else {
        return Ok(None);
    };
    let rotation = fields.unique(4)?.map(|field| field.float()).transpose()?;
    Ok(Some(Geometry {
        x,
        y,
        width,
        height,
        rotation,
    }))
}

/// The concatenated text of one `TSWP.StorageArchive`.
pub(super) fn storage_text(
    source: &impl Objects,
    storage: u64,
    wire: WireLimits,
) -> Result<String, StructureError> {
    let object = source.get(storage).ok_or(StructureError::MissingObject)?;
    let (_, payload) = unique_message(object, STORAGE)?;
    let fields = Fields::parse(payload, wire)?;
    let mut text = String::new();
    for field in fields.all(3) {
        text.push_str(
            std::str::from_utf8(field.message()?)
                .map_err(|_| StructureError::Unexpected("storage text UTF-8"))?,
        );
    }
    Ok(text)
}
