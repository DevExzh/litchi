//! Wire-level authoring of new text boxes, preset shapes and images, after
//! the legacy host's typed constructors (`text_box_create.rs`,
//! `slide_shapes.rs`, `slide_images/graph.rs`, `shapes/path.rs`).

use litchi_iwa_common::WireLimits;
use litchi_iwa_core::{ArchiveObject, RawMessage};
use sha1::{Digest, Sha1};

use super::super::structure::StructureError;
use super::super::structure::metadata::ExternalReference;
use super::super::structure::wire::{self, Fields, LENGTH_DELIMITED, Writer};
use super::super::structure::workspace::Workspace;
use super::inventory::{
    self, IMAGE, SHAPE_INFO, SLIDE, SLIDE_OWNED_DRAWABLES, SLIDE_Z_ORDER, STORAGE, descend,
    unique_message,
};
use super::{Error, Geometry, ImageFormat, Preset, Unsupported, lift};

const DOCUMENT_OBJECT: u64 = 1;
const DOCUMENT: u32 = 1;
const SHOW: u32 = 2;
const THEME: u32 = 10;
const STYLESHEET: u32 = 401;
const MEDIA_STYLE: u32 = 3_016;
const STANDIN_CAPTION: u32 = 3_097;
const PLACEHOLDER: u32 = 7;
const DEFAULT_DRAWABLE_FLAGS: u64 = 3;
const DEFAULT_ROTATION: f32 = 0.0;
const WRAP_TYPE_SQUARE: u64 = 4;
const WRAP_DIRECTION_BOTH: u64 = 2;
const WRAP_FIT_TEXT: u64 = 1;
const WRAP_MARGIN: f32 = 12.0;
const WRAP_ALPHA: f32 = 0.5;
const STANDARD_VERSION: [u32; 3] = [1, 0, 5];
const STANDIN_VERSION: [u32; 3] = [10, 1, 0];
const ELLIPSE_CONTROL_FACTOR: f32 = 0.552_284_8;

/// Theme and document context shared by every authored drawable.
pub(super) struct Context {
    stylesheet: u64,
    textbox_style: Option<u64>,
    paragraph_style: Option<u64>,
    list_style: Option<u64>,
    media_style: Option<u64>,
    language: Option<String>,
}

fn first_reference(
    payload: Option<&[u8]>,
    number: u32,
    wire: WireLimits,
) -> Result<Option<u64>, StructureError> {
    let Some(payload) = payload else {
        return Ok(None);
    };
    let fields = Fields::parse(payload, wire)?;
    fields
        .all(number)
        .next()
        .map(|field| wire::reference_identifier(field.message()?, wire))
        .transpose()
}

fn reference_field(payload: &[u8], number: u32, wire: WireLimits) -> Result<u64, StructureError> {
    first_reference(Some(payload), number, wire)?
        .ok_or(StructureError::Unexpected("required reference"))
}

impl Context {
    /// Resolve the show stylesheet, theme presets and document language.
    pub(super) fn read(workspace: &Workspace<'_>) -> Result<Self, StructureError> {
        let wire = workspace.wire;
        let document = workspace.payload(DOCUMENT_OBJECT, DOCUMENT)?.to_vec();
        let show_id = reference_field(&document, 2, wire)?;
        let language = match descend(&document, &[3], wire)? {
            Some(super_) => Fields::parse(&super_, wire)?
                .unique(3)?
                .map(|field| {
                    String::from_utf8(field.message()?.to_vec())
                        .map_err(|_| StructureError::Unexpected("language UTF-8"))
                })
                .transpose()?,
            None => None,
        };
        let show = workspace.payload(show_id, SHOW)?.to_vec();
        let theme_id = reference_field(&show, 2, wire)?;
        let stylesheet = reference_field(&show, 5, wire)?;
        let theme = workspace.payload(theme_id, THEME)?.to_vec();
        let drawing = descend(&theme, &[1, 100], wire)?;
        let text = descend(&theme, &[1, 110], wire)?;
        let styles = workspace.payload(stylesheet, STYLESHEET)?.to_vec();
        let mut media_style = None;
        for field in Fields::parse(&styles, wire)?.all(1) {
            let style = wire::reference_identifier(field.message()?, wire)?;
            if workspace.object(style).is_some_and(|object| {
                object
                    .messages
                    .iter()
                    .any(|message| message.type_ == MEDIA_STYLE)
            }) {
                media_style = Some(style);
                break;
            }
        }
        Ok(Self {
            stylesheet,
            textbox_style: first_reference(drawing.as_deref(), 6, wire)?,
            paragraph_style: first_reference(text.as_deref(), 7, wire)?,
            list_style: first_reference(text.as_deref(), 1, wire)?,
            media_style,
            language,
        })
    }
}

/// The storage styles a new text body inherits.
struct StorageStyles {
    stylesheet: u64,
    paragraph: Option<u64>,
    list: Option<u64>,
    drop_cap: Option<u64>,
    language: Option<String>,
}

fn first_table_object(
    fields: &Fields,
    number: u32,
    wire: WireLimits,
) -> Result<Option<u64>, StructureError> {
    let Some(table) = fields.unique(number)? else {
        return Ok(None);
    };
    for entry in Fields::parse(table.message()?, wire)?.all(1) {
        if let Some(object) = Fields::parse(entry.message()?, wire)?.unique(2)? {
            return Ok(Some(wire::reference_identifier(object.message()?, wire)?));
        }
    }
    Ok(None)
}

/// Styles from the slide's body or title placeholder storage (the host's
/// `slide_text_storage_template`), falling back to the theme presets.
fn storage_styles(
    workspace: &Workspace<'_>,
    slide_payload: &[u8],
    context: &Context,
) -> Result<StorageStyles, StructureError> {
    let wire = workspace.wire;
    let slide = Fields::parse(slide_payload, wire)?;
    let mut template = None;
    for number in [6, 5] {
        let Some(field) = slide.unique(number)? else {
            continue;
        };
        let placeholder = wire::reference_identifier(field.message()?, wire)?;
        let Some(object) = workspace.object(placeholder) else {
            continue;
        };
        let Ok((_, payload)) = unique_message(object, PLACEHOLDER) else {
            continue;
        };
        let Some(shape) = descend(payload, &[1], wire)? else {
            continue;
        };
        let shape = Fields::parse(&shape, wire)?;
        let Some(storage) = shape.unique(4)?.or(shape.unique(2)?) else {
            continue;
        };
        let storage = wire::reference_identifier(storage.message()?, wire)?;
        if let Some(object) = workspace.object(storage)
            && let Ok((_, payload)) = unique_message(object, STORAGE)
        {
            template = Some(Fields::parse(payload, wire)?);
            break;
        }
    }
    let Some(template) = template else {
        return Ok(StorageStyles {
            stylesheet: context.stylesheet,
            paragraph: context.paragraph_style,
            list: context.list_style,
            drop_cap: None,
            language: context.language.clone(),
        });
    };
    let stylesheet = match template.unique(2)? {
        Some(field) => wire::reference_identifier(field.message()?, wire)?,
        None => context.stylesheet,
    };
    let mut language = None;
    if let Some(table) = template.unique(19)? {
        for entry in Fields::parse(table.message()?, wire)?.all(1) {
            if let Some(object) = Fields::parse(entry.message()?, wire)?.unique(2)? {
                language = Some(
                    String::from_utf8(object.message()?.to_vec())
                        .map_err(|_| StructureError::Unexpected("language UTF-8"))?,
                );
                break;
            }
        }
    }
    Ok(StorageStyles {
        stylesheet,
        paragraph: first_table_object(&template, 5, wire)?.or(context.paragraph_style),
        list: first_table_object(&template, 7, wire)?.or(context.list_style),
        drop_cap: first_table_object(&template, 28, wire)?,
        language: language.or_else(|| context.language.clone()),
    })
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

fn storage_payload(text: &str, styles: &StorageStyles) -> (Vec<u8>, Vec<u64>) {
    let mut storage = Writer::new()
        .reference(2, styles.stylesheet)
        .string(3, text)
        .message(5, &object_table(styles.paragraph))
        .message(6, &zero_para_table())
        .message(7, &object_table(styles.list))
        .varint(10, 1)
        .message(14, &zero_para_table());
    if let Some(language) = &styles.language {
        storage = storage.message(
            19,
            &Writer::new().message(1, &Writer::new().varint(1, 0).string(2, language)),
        );
    }
    storage = storage
        .message(24, &zero_para_table())
        .message(28, &object_table(styles.drop_cap));
    let mut references = [
        Some(styles.stylesheet),
        styles.paragraph,
        styles.list,
        styles.drop_cap,
    ]
    .into_iter()
    .flatten()
    .collect::<Vec<_>>();
    references.sort_unstable();
    references.dedup();
    (storage.finish(), references)
}

fn point(x: f32, y: f32) -> Writer {
    Writer::new().float(1, x).float(2, y)
}

pub(super) fn drawable_archive(geometry: Geometry, slide: u64, title: u64, caption: u64) -> Writer {
    Writer::new()
        .message(
            1,
            &Writer::new()
                .message(1, &point(geometry.x, geometry.y))
                .message(2, &point(geometry.width, geometry.height))
                .varint(3, DEFAULT_DRAWABLE_FLAGS)
                .float(4, geometry.rotation.unwrap_or(DEFAULT_ROTATION)),
        )
        .reference(2, slide)
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
}

fn element(kind: u64, points: &[(f32, f32)]) -> Writer {
    let mut element = Writer::new().varint(1, kind);
    for (x, y) in points {
        element = element.message(2, &point(*x, *y));
    }
    element
}

fn path(elements: &[Writer]) -> Writer {
    let mut path = Writer::new();
    for item in elements {
        path = path.message(1, item);
    }
    path
}

/// The host's `shape_path_source` for every preset.
pub(super) fn path_source(preset: Preset, width: f32, height: f32) -> Result<Writer, Error> {
    let size = point(width, height);
    let source = Writer::new().boolean(1, false).boolean(2, false);
    Ok(match preset {
        Preset::Rectangle => source.message(
            5,
            &Writer::new().message(2, &size).message(
                3,
                &path(&[
                    element(1, &[(0.0, 0.0)]),
                    element(2, &[(width, 0.0)]),
                    element(2, &[(width, height)]),
                    element(2, &[(0.0, height)]),
                    element(5, &[]),
                    element(1, &[(0.0, 0.0)]),
                ]),
            ),
        ),
        Preset::RoundedRectangle { radius } => {
            if !radius.is_finite() || radius < 0.0 || radius > width.min(height) / 2.0 {
                return Err(Error::Refused(
                    crate::slide::drawable::Refusal::InvalidGeometry,
                ));
            }
            source.message(
                4,
                &Writer::new()
                    .varint(1, 0)
                    .float(2, radius)
                    .message(3, &size)
                    .boolean(4, true),
            )
        },
        Preset::Ellipse => {
            let (cx, cy) = (width / 2.0, height / 2.0);
            let (kx, ky) = (cx * ELLIPSE_CONTROL_FACTOR, cy * ELLIPSE_CONTROL_FACTOR);
            source.message(
                5,
                &Writer::new().message(2, &size).message(
                    3,
                    &path(&[
                        element(1, &[(cx, 0.0)]),
                        element(4, &[(cx + kx, 0.0), (width, cy - ky), (width, cy)]),
                        element(4, &[(width, cy + ky), (cx + kx, height), (cx, height)]),
                        element(4, &[(cx - kx, height), (0.0, cy + ky), (0.0, cy)]),
                        element(4, &[(0.0, cy - ky), (cx - kx, 0.0), (cx, 0.0)]),
                        element(5, &[]),
                        element(1, &[(cx, 0.0)]),
                    ]),
                ),
            )
        },
        Preset::LeftArrow => {
            source.message(3, &arrow(0, SINGLE_ARROW_HEAD_START_RATIO, width, &size))
        },
        Preset::RightArrow => {
            source.message(3, &arrow(1, SINGLE_ARROW_HEAD_START_RATIO, width, &size))
        },
        Preset::DoubleArrow => {
            source.message(3, &arrow(10, DOUBLE_ARROW_HEAD_END_RATIO, width, &size))
        },
        Preset::RegularPolygon { sides } => source.message(
            4,
            &Writer::new()
                .varint(1, 1)
                .float(2, f32::from(sides.count()))
                .message(3, &size)
                .boolean(4, false),
        ),
        Preset::Star {
            points,
            inner_radius,
        } => source.message(
            3,
            &Writer::new()
                .varint(1, 100)
                .message(2, &point(f32::from(points.count()), inner_radius.ratio()))
                .message(3, &size),
        ),
    })
}

/// The host's native arrow proportions (`TSD.PointPathSourceArchive`).
const SINGLE_ARROW_HEAD_START_RATIO: f32 = 0.64;
const DOUBLE_ARROW_HEAD_END_RATIO: f32 = 0.40;
const ARROW_SHAFT_RATIO: f32 = 0.34;

/// A `TSD.PointPathSourceArchive` arrow of native `kind`.
fn arrow(kind: u64, head_ratio: f32, width: f32, size: &Writer) -> Writer {
    Writer::new()
        .varint(1, kind)
        .message(2, &point(width * head_ratio, ARROW_SHAFT_RATIO))
        .message(3, size)
}

fn native_object(
    workspace: &Workspace<'_>,
    identifier: u64,
    message_type: u32,
    payload: Vec<u8>,
    versions: &[u32],
    object_references: &[u64],
    data_references: &[u64],
) -> Result<ArchiveObject, StructureError> {
    let mut object = ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_: message_type,
            data: payload,
        }],
        workspace.archive_limits(),
    )?;
    let info = object
        .archive_info
        .message_infos
        .first_mut()
        .ok_or(StructureError::Unexpected("message info"))?;
    info.versions = versions.to_vec().into();
    info.object_references = object_references.to_vec();
    info.data_references = data_references.to_vec();
    Ok(object)
}

/// Append `drawable` to the slide's owned drawables and z-order (front).
pub(super) fn attach(
    workspace: &mut Workspace<'_>,
    slide: u64,
    drawable: u64,
) -> Result<(), StructureError> {
    let wire = workspace.wire;
    let limits = workspace.archive_limits();
    let object = workspace.require(slide)?;
    let (index, payload) = unique_message(object, SLIDE)?;
    let aggregate = object
        .archive_info
        .message_infos
        .get(index)
        .ok_or(StructureError::Unexpected("message info"))?
        .object_references
        .clone();
    let mut fields = Fields::parse(payload, wire)?;
    for number in [SLIDE_OWNED_DRAWABLES, SLIDE_Z_ORDER] {
        let mut values = inventory::references(&fields, number, wire)?;
        if values.contains(&drawable) {
            return Err(StructureError::Unexpected("drawable already attached"));
        }
        values.push(drawable);
        let encoded = values
            .iter()
            .map(|id| wire::reference(*id))
            .collect::<Vec<_>>();
        fields.replace_all(number, LENGTH_DELIMITED, &encoded);
    }
    let mut after = aggregate.clone();
    if !after.contains(&drawable) {
        after.push(drawable);
    }
    workspace
        .object_mut(slide)?
        .replace_message_transitioning_object_references_preserving_header(
            index,
            RawMessage {
                type_: SLIDE,
                data: fields.encode(),
            },
            litchi_iwa_core::archive::ObjectReferenceTransition {
                aggregate_before: &aggregate,
                aggregate_after: &after,
                fields: &[],
            },
            limits,
        )?;
    Ok(())
}

/// Register new objects' UUIDs and an external style reference in the
/// slide component.
fn register(
    workspace: &mut Workspace<'_>,
    member: &str,
    objects: &[u64],
    style: u64,
) -> Result<u64, StructureError> {
    let component = workspace
        .registry
        .component_for_member(member)
        .ok_or(StructureError::Registry("slide component"))?
        .identifier;
    let style_component = workspace
        .component_of(style)
        .and_then(|member| workspace.registry.component_for_member(&member))
        .map(|component| component.identifier);
    let mut uuids = Vec::with_capacity(objects.len());
    for identifier in objects {
        uuids.push((*identifier, workspace.clone_uuid(None, *identifier)?));
    }
    let entry = workspace
        .registry
        .component_mut(component)
        .ok_or(StructureError::Registry("slide component"))?;
    for (identifier, uuid) in uuids {
        entry.add_object_uuid(identifier, uuid);
    }
    if let Some(style_component) = style_component
        && style_component != component
    {
        entry.add_external_reference(ExternalReference {
            component: style_component,
            object: Some(style),
            is_weak: None,
        });
    }
    Ok(component)
}

fn slide_member(workspace: &Workspace<'_>, slide: u64) -> Result<String, Error> {
    workspace.component_of(slide).ok_or(Error::InvalidSource)
}

/// Add a text box (`is_text_box`) or preset shape with its own storage.
pub(super) fn add_text_shape(
    workspace: &mut Workspace<'_>,
    context: &Context,
    slide: u64,
    preset: Option<Preset>,
    geometry: Geometry,
    text: &str,
) -> Result<u64, Error> {
    let shape_path = path_source(
        preset.unwrap_or(Preset::Rectangle),
        geometry.width,
        geometry.height,
    )?;
    add_path_shape(
        workspace,
        context,
        slide,
        geometry,
        &shape_path,
        preset.is_none(),
        text,
    )
}

/// The host's `line_path_source`: a two-element Bézier of the line's length.
fn line_path_source(length: f32) -> Writer {
    Writer::new().boolean(1, false).boolean(2, false).message(
        5,
        &Writer::new().message(2, &point(length, 0.0)).message(
            3,
            &path(&[element(1, &[(0.0, 0.0)]), element(2, &[(length, 0.0)])]),
        ),
    )
}

/// Add a native straight line along `segment` (host `add_slide_line`).
pub(super) fn add_line(
    workspace: &mut Workspace<'_>,
    context: &Context,
    slide: u64,
    segment: litchi_iwa_common::shape::line::Segment,
) -> Result<u64, Error> {
    let (position, size, angle) = segment.geometry();
    let geometry = Geometry {
        x: position.x,
        y: position.y,
        width: size.width,
        height: size.height,
        rotation: Some(angle),
    };
    add_path_shape(
        workspace,
        context,
        slide,
        geometry,
        &line_path_source(segment.length()),
        false,
        "",
    )
}

fn add_path_shape(
    workspace: &mut Workspace<'_>,
    context: &Context,
    slide: u64,
    geometry: Geometry,
    shape_path: &Writer,
    text_box: bool,
    text: &str,
) -> Result<u64, Error> {
    let style = context.textbox_style.ok_or(Error::Refused(
        crate::slide::drawable::Refusal::Unsupported {
            reason: Unsupported::MissingThemeStyle,
        },
    ))?;
    let member = slide_member(workspace, slide)?;
    let slide_payload = workspace.payload(slide, SLIDE).map_err(lift)?.to_vec();
    let styles = storage_styles(workspace, &slide_payload, context).map_err(lift)?;
    let drawable = workspace.allocate().map_err(lift)?;
    let caption = workspace.allocate().map_err(lift)?;
    let title = workspace.allocate().map_err(lift)?;
    let storage = workspace.allocate().map_err(lift)?;
    let shape = Writer::new()
        .message(
            1,
            &Writer::new()
                .message(1, &drawable_archive(geometry, slide, title, caption))
                .reference(2, style)
                .message(3, shape_path)
                .float(6, 0.0),
        )
        .reference(2, storage)
        .reference(4, storage)
        .boolean(6, text_box)
        .finish();
    let (storage_payload, storage_references) = storage_payload(text, &styles);
    let objects = [
        native_object(
            workspace,
            drawable,
            SHAPE_INFO,
            shape,
            &STANDARD_VERSION,
            &[caption, title, style, storage],
            &[],
        ),
        native_object(
            workspace,
            caption,
            STANDIN_CAPTION,
            Vec::new(),
            &STANDIN_VERSION,
            &[],
            &[],
        ),
        native_object(
            workspace,
            title,
            STANDIN_CAPTION,
            Vec::new(),
            &STANDIN_VERSION,
            &[],
            &[],
        ),
        native_object(
            workspace,
            storage,
            STORAGE,
            storage_payload,
            &STANDARD_VERSION,
            &storage_references,
            &[],
        ),
    ];
    for object in objects {
        workspace
            .insert(&member, object.map_err(lift)?)
            .map_err(lift)?;
    }
    attach(workspace, slide, drawable).map_err(lift)?;
    register(
        workspace,
        &member,
        &[drawable, caption, title, storage],
        style,
    )
    .map_err(lift)?;
    Ok(drawable)
}

/// Sniff a supported raster format and its pixel size.
pub(super) fn sniff(bytes: &[u8]) -> Option<(ImageFormat, Option<(u32, u32)>)> {
    let be16 = |at: usize| {
        bytes
            .get(at..at + 2)
            .map(|b| u32::from(u16::from_be_bytes([b[0], b[1]])))
    };
    if bytes.starts_with(b"\x89PNG\r\n\x1a\n") {
        let size = bytes.get(16..24).map(|b| {
            (
                u32::from_be_bytes([b[0], b[1], b[2], b[3]]),
                u32::from_be_bytes([b[4], b[5], b[6], b[7]]),
            )
        });
        return Some((ImageFormat::Png, size));
    }
    if bytes.starts_with(b"GIF87a") || bytes.starts_with(b"GIF89a") {
        let size = bytes.get(6..10).map(|b| {
            (
                u32::from(u16::from_le_bytes([b[0], b[1]])),
                u32::from(u16::from_le_bytes([b[2], b[3]])),
            )
        });
        return Some((ImageFormat::Gif, size));
    }
    if bytes.starts_with(&[0xFF, 0xD8, 0xFF]) {
        let mut at = 2;
        let mut size = None;
        while at + 4 <= bytes.len() {
            if bytes[at] != 0xFF {
                break;
            }
            let marker = bytes[at + 1];
            if marker == 0xFF {
                at += 1;
                continue;
            }
            let Some(length) = be16(at + 2) else { break };
            if matches!(marker, 0xC0..=0xCF) && !matches!(marker, 0xC4 | 0xC8 | 0xCC) {
                size = be16(at + 7).zip(be16(at + 5));
                break;
            }
            at += 2 + length as usize;
        }
        return Some((ImageFormat::Jpeg, size));
    }
    None
}

fn data_file_name(preferred: &str, identifier: u64) -> Result<String, Error> {
    let path = std::path::Path::new(preferred);
    let stem = path
        .file_stem()
        .and_then(|stem| stem.to_str())
        .filter(|stem| !stem.is_empty())
        .ok_or(Error::Refused(
            crate::slide::drawable::Refusal::InvalidImage,
        ))?;
    let name = match path.extension().and_then(|extension| extension.to_str()) {
        Some(extension) if !extension.is_empty() => format!("{stem}-{identifier}.{extension}"),
        _ => format!("{stem}-{identifier}"),
    };
    if name.contains(['/', '\\', '\0']) || name.starts_with('.') {
        return Err(Error::Refused(
            crate::slide::drawable::Refusal::InvalidImage,
        ));
    }
    Ok(name)
}

/// Add an image drawable with a new `Data/` member and registry records.
pub(super) fn add_image(
    workspace: &mut Workspace<'_>,
    context: &Context,
    slide: u64,
    bytes: &[u8],
    preferred_file_name: &str,
    geometry: Geometry,
) -> Result<u64, Error> {
    let style = context.media_style.ok_or(Error::Refused(
        crate::slide::drawable::Refusal::Unsupported {
            reason: Unsupported::MissingThemeStyle,
        },
    ))?;
    let (_, pixels) = sniff(bytes).ok_or(Error::Refused(
        crate::slide::drawable::Refusal::Unsupported {
            reason: Unsupported::ImageFormat,
        },
    ))?;
    let natural = pixels
        .filter(|(width, height)| *width > 0 && *height > 0)
        .map_or((geometry.width, geometry.height), |(width, height)| {
            (width as f32, height as f32)
        });
    let member = slide_member(workspace, slide)?;
    let data = workspace
        .registry
        .data_identifiers(workspace.wire)
        .map_err(lift)?
        .into_iter()
        .max()
        .unwrap_or(0)
        .checked_add(1)
        .ok_or(Error::Limit {
            kind: super::LimitKind::Identifiers,
            observed: u64::MAX,
            maximum: u64::MAX,
        })?;
    let file_name = data_file_name(preferred_file_name, data)?;
    let digest = Sha1::digest(bytes);
    workspace
        .insert_member(&format!("Data/{file_name}"), bytes.to_vec())
        .map_err(lift)?;
    workspace.registry.add_data_info(
        data,
        &digest,
        preferred_file_name,
        &file_name,
        bytes.len() as u64,
    );
    let drawable = workspace.allocate().map_err(lift)?;
    let title = workspace.allocate().map_err(lift)?;
    let caption = workspace.allocate().map_err(lift)?;
    let image = Writer::new()
        .message(1, &drawable_archive(geometry, slide, title, caption))
        .reference(3, style)
        .message(4, &point(natural.0, natural.1))
        .varint(7, 0)
        .message(9, &point(natural.0, natural.1))
        .message(11, &Writer::new().varint(1, data))
        .finish();
    let objects = [
        native_object(
            workspace,
            drawable,
            IMAGE,
            image,
            &STANDARD_VERSION,
            &[title, caption, style],
            &[data],
        ),
        native_object(
            workspace,
            title,
            STANDIN_CAPTION,
            Vec::new(),
            &STANDIN_VERSION,
            &[],
            &[],
        ),
        native_object(
            workspace,
            caption,
            STANDIN_CAPTION,
            Vec::new(),
            &STANDIN_VERSION,
            &[],
            &[],
        ),
    ];
    for object in objects {
        workspace
            .insert(&member, object.map_err(lift)?)
            .map_err(lift)?;
    }
    attach(workspace, slide, drawable).map_err(lift)?;
    let component =
        register(workspace, &member, &[drawable, title, caption], style).map_err(lift)?;
    workspace
        .registry
        .component_mut(component)
        .ok_or(Error::InvalidSource)?
        .add_data_owner(data, drawable);
    Ok(drawable)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn sniff_reads_formats_and_pixel_sizes() {
        let mut png = b"\x89PNG\r\n\x1a\n\0\0\0\rIHDR".to_vec();
        png.extend_from_slice(&[0, 0, 0, 2, 0, 0, 0, 3]);
        assert_eq!(sniff(&png), Some((ImageFormat::Png, Some((2, 3)))));
        let gif = b"GIF89a\x05\x00\x07\x00";
        assert_eq!(sniff(gif), Some((ImageFormat::Gif, Some((5, 7)))));
        let jpeg = [
            0xFF, 0xD8, 0xFF, 0xE0, 0x00, 0x04, 0x00, 0x00, 0xFF, 0xC0, 0x00, 0x11, 0x08, 0x00,
            0x09, 0x00, 0x0B,
        ];
        assert_eq!(sniff(&jpeg), Some((ImageFormat::Jpeg, Some((11, 9)))));
        assert_eq!(sniff(b"BM"), None);
        assert_eq!(
            data_file_name("photo.jpeg", 4).ok().as_deref(),
            Some("photo-4.jpeg")
        );
        assert_eq!(data_file_name("plain", 4).ok().as_deref(), Some("plain-4"));
        assert!(data_file_name("", 4).is_err());
    }

    #[test]
    fn rounded_rectangle_radius_is_bounded() {
        assert!(path_source(Preset::RoundedRectangle { radius: 5.0 }, 10.0, 20.0).is_ok());
        assert!(matches!(
            path_source(Preset::RoundedRectangle { radius: 5.1 }, 10.0, 20.0),
            Err(Error::Refused(
                crate::slide::drawable::Refusal::InvalidGeometry
            ))
        ));
        assert!(path_source(Preset::Ellipse, 10.0, 20.0).is_ok());
    }
}
