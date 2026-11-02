//! Wire-level authoring of new sheet text boxes and preset shapes, after the
//! legacy host's typed constructors (`numbers/editor/text_box_create.rs`,
//! `sheet_shapes.rs`, `shapes/path.rs`): the drawable, its title and caption
//! stand-ins and its writable storage, styled from the document theme's
//! presets.

use litchi_iwa_common::WireLimits;
use litchi_iwa_common::shape::path::Preset;
use litchi_iwa_common::wire::WireView;
use litchi_iwa_core::ArchiveObject;

use super::Package;
use super::sheet_drawable_graph::{self as graph, SHAPE_INFO_MESSAGE_TYPE};
use super::structure::{self, Fault, FaultResult, references};
use crate::sheet::drawable::Geometry;

/// `TN.ThemeArchive`.
const THEME_MESSAGE_TYPE: u32 = 12_009;
/// `TSWP.StorageArchive`.
const STORAGE_MESSAGE_TYPE: u32 = 2_001;
/// `TSD.StandinCaptionArchive`.
const STANDIN_CAPTION_MESSAGE_TYPE: u32 = 3_097;
const DOCUMENT_STYLESHEET_FIELD: u32 = 4;
const DOCUMENT_THEME_FIELD: u32 = 6;
const DOCUMENT_SUPER_FIELD: u32 = 8;
const DOCUMENT_LANGUAGE_FIELD: u32 = 3;
const STANDARD_VERSION: [u32; 3] = [1, 0, 5];
const STANDIN_VERSION: [u32; 3] = [10, 1, 0];
const WRAP_TYPE_SQUARE: u64 = 4;
const WRAP_DIRECTION_BOTH: u64 = 2;
const WRAP_FIT_TEXT: u64 = 1;
const WRAP_MARGIN: f32 = 12.0;
const WRAP_ALPHA: f32 = 0.5;

/// A minimal canonical protobuf writer (fields in call order).
#[derive(Default)]
pub(super) struct Writer(Vec<u8>);

impl Writer {
    pub(super) fn new() -> Self {
        Self::default()
    }

    fn key(&mut self, number: u32, wire_type: u8) {
        self.raw_varint((u64::from(number) << 3) | u64::from(wire_type));
    }

    fn raw_varint(&mut self, mut value: u64) {
        loop {
            #[allow(
                clippy::cast_possible_truncation,
                reason = "the value is masked to seven bits"
            )]
            let byte = (value & 0x7f) as u8;
            value >>= 7;
            if value == 0 {
                self.0.push(byte);
                return;
            }
            self.0.push(byte | 0x80);
        }
    }

    pub(super) fn varint(mut self, number: u32, value: u64) -> Self {
        self.key(number, 0);
        self.raw_varint(value);
        self
    }

    pub(super) fn boolean(self, number: u32, value: bool) -> Self {
        self.varint(number, u64::from(value))
    }

    pub(super) fn float(mut self, number: u32, value: f32) -> Self {
        self.key(number, 5);
        self.0.extend_from_slice(&value.to_le_bytes());
        self
    }

    pub(super) fn bytes(mut self, number: u32, payload: &[u8]) -> Self {
        self.key(number, 2);
        self.raw_varint(payload.len() as u64);
        self.0.extend_from_slice(payload);
        self
    }

    pub(super) fn message(self, number: u32, inner: &Self) -> Self {
        self.bytes(number, &inner.0)
    }

    pub(super) fn reference(self, number: u32, identifier: u64) -> Self {
        self.message(number, &Self::new().varint(1, identifier))
    }

    pub(super) fn finish(self) -> Vec<u8> {
        self.0
    }
}

/// The document styles a new text box or shape takes (the host's
/// `text_box_theme_styles`) and the document language.
pub(super) struct Context {
    /// The theme's first text-box style preset (shape style).
    pub(super) shape_style: u64,
    stylesheet: u64,
    paragraph: Option<u64>,
    list: Option<u64>,
    language: Option<String>,
}

fn first_reference(payload: &[u8], number: u32) -> FaultResult<Option<u64>> {
    references::records(payload, &[number])?
        .first()
        .map(|record| references::reference_identifier(record))
        .transpose()
}

fn unique_message_payload(object: &ArchiveObject, message_type: u32) -> FaultResult<&[u8]> {
    let index = structure::unique_message(object, &[message_type])?;
    Ok(&object.messages[index].data)
}

/// A nested message at `path`, if present.
fn descend<'p>(payload: &'p [u8], path: &[u32]) -> FaultResult<Option<&'p [u8]>> {
    let mut current = payload;
    for number in path {
        let view = WireView::parse(current, WireLimits::DEFAULT)
            .map_err(|error| structure::map_wire(&error))?;
        match graph::unique_field(&view, *number, 2)? {
            Some(field) => current = field.payload(),
            None => return Ok(None),
        }
    }
    Ok(Some(current))
}

impl Context {
    /// Resolve the theme presets, stylesheet and language of `package`.
    pub(super) fn read(package: &Package) -> FaultResult<Self> {
        let document = graph::object(package, 1)?;
        let document = unique_message_payload(document, 1)?;
        let stylesheet =
            first_reference(document, DOCUMENT_STYLESHEET_FIELD)?.ok_or(Fault::Invalid)?;
        let theme = first_reference(document, DOCUMENT_THEME_FIELD)?.ok_or(Fault::Invalid)?;
        let language = match descend(document, &[DOCUMENT_SUPER_FIELD])? {
            Some(super_) => {
                let view = WireView::parse(super_, WireLimits::DEFAULT)
                    .map_err(|error| structure::map_wire(&error))?;
                graph::unique_field(&view, DOCUMENT_LANGUAGE_FIELD, 2)?
                    .map(|field| {
                        String::from_utf8(field.payload().to_vec()).map_err(|_| Fault::Invalid)
                    })
                    .transpose()?
            },
            None => None,
        };
        let theme = unique_message_payload(graph::object(package, theme)?, THEME_MESSAGE_TYPE)?;
        let drawing = descend(theme, &[1, 100])?.ok_or(Fault::Unsupported)?;
        let shape_style = first_reference(drawing, 6)?.ok_or(Fault::Unsupported)?;
        let (paragraph, list) = match descend(theme, &[1, 110])? {
            Some(text) => (first_reference(text, 7)?, first_reference(text, 1)?),
            None => (None, None),
        };
        Ok(Self {
            shape_style,
            stylesheet,
            paragraph,
            list,
            language,
        })
    }
}

fn object_table(object: Option<u64>) -> Writer {
    let mut entry = Writer::new().varint(1, 0);
    if let Some(object) = object {
        entry = entry.reference(2, object);
    }
    Writer::new().message(1, &entry)
}

fn zero_para_table() -> Writer {
    Writer::new().message(1, &Writer::new().varint(1, 0).varint(2, 0).varint(3, 0))
}

/// The host's `text_box_storage` and its object references.
fn storage_payload(text: &str, context: &Context) -> (Vec<u8>, Vec<u64>) {
    let mut storage = Writer::new()
        .reference(2, context.stylesheet)
        .bytes(3, text.as_bytes())
        .message(5, &object_table(context.paragraph))
        .message(6, &zero_para_table())
        .message(7, &object_table(context.list))
        .varint(10, 1)
        .message(14, &zero_para_table());
    if let Some(language) = &context.language {
        storage = storage.message(
            19,
            &Writer::new().message(1, &Writer::new().varint(1, 0).bytes(2, language.as_bytes())),
        );
    }
    storage = storage
        .message(24, &zero_para_table())
        .message(28, &object_table(None));
    let mut references: Vec<u64> = [Some(context.stylesheet), context.paragraph, context.list]
        .into_iter()
        .flatten()
        .collect();
    references.sort_unstable();
    references.dedup();
    (storage.finish(), references)
}

fn point(x: f32, y: f32) -> Writer {
    Writer::new().float(1, x).float(2, y)
}

/// The objects of one new text box or shape, in insertion order.
pub(super) struct Authored {
    /// drawable, caption, title, storage.
    pub(super) identifiers: [u64; 4],
    pub(super) objects: Vec<ArchiveObject>,
}

fn native_object(
    identifier: u64,
    message_type: u32,
    payload: Vec<u8>,
    versions: &[u32],
    object_references: &[u64],
    limits: litchi_iwa_core::Limits,
) -> FaultResult<ArchiveObject> {
    let mut object = structure::new_object(
        identifier,
        None,
        message_type,
        payload,
        object_references.to_vec(),
        limits,
    )?;
    object.archive_info.message_infos[0].versions = versions.to_vec().into();
    Ok(object)
}

/// The path a new shape draws.
#[derive(Debug, Clone, Copy, PartialEq)]
pub(super) enum Path {
    /// A preset drawn in the frame.
    Preset(Preset),
    /// A native straight line of this length (the frame's width).
    Line(f32),
}

/// Author a text box (`text_box`), preset shape or straight line at `geometry` with
/// `text`, owned by `sheet`, under the identifiers `identifiers`
/// (drawable, caption, title, storage). `None` when the preset does not fit
/// the frame.
pub(super) fn author(
    context: &Context,
    identifiers: [u64; 4],
    sheet: u64,
    geometry: Geometry,
    path: Path,
    text: &str,
    text_box: bool,
    limits: litchi_iwa_core::Limits,
) -> FaultResult<Option<Authored>> {
    let [drawable, caption, title, storage] = identifiers;
    let (Some(position), Some(size)) = (geometry.position, geometry.size) else {
        return Err(Fault::Invalid);
    };
    let path_source = match path {
        Path::Preset(preset) => {
            let Some(source) =
                litchi_iwa_style::shape_path::path_source(preset, size.width, size.height)
            else {
                return Ok(None);
            };
            source
        },
        Path::Line(length) => litchi_iwa_style::shape_path::line_path_source(length),
    };
    let frame = Writer::new()
        .message(1, &point(position.x, position.y))
        .message(2, &point(size.width, size.height))
        .varint(3, u64::from(geometry.flags.map_or(3, |flags| flags.bits())))
        .float(4, geometry.angle.unwrap_or(0.0));
    let drawable_archive = Writer::new()
        .message(1, &frame)
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
        .boolean(13, false);
    let shape = Writer::new()
        .message(1, &drawable_archive)
        .reference(2, context.shape_style)
        .bytes(3, &path_source)
        .float(6, 0.0);
    let shape_info = Writer::new()
        .message(1, &shape)
        .reference(2, storage)
        .reference(4, storage)
        .boolean(6, text_box)
        .finish();
    let (storage_payload, storage_references) = storage_payload(text, context);
    let objects = vec![
        native_object(
            drawable,
            SHAPE_INFO_MESSAGE_TYPE,
            shape_info,
            &STANDARD_VERSION,
            &[caption, title, context.shape_style, storage],
            limits,
        )?,
        native_object(
            caption,
            STANDIN_CAPTION_MESSAGE_TYPE,
            Vec::new(),
            &STANDIN_VERSION,
            &[],
            limits,
        )?,
        native_object(
            title,
            STANDIN_CAPTION_MESSAGE_TYPE,
            Vec::new(),
            &STANDIN_VERSION,
            &[],
            limits,
        )?,
        native_object(
            storage,
            STORAGE_MESSAGE_TYPE,
            storage_payload,
            &STANDARD_VERSION,
            &storage_references,
            limits,
        )?,
    ];
    Ok(Some(Authored {
        identifiers,
        objects,
    }))
}
