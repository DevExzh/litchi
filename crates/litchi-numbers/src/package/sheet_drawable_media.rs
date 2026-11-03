//! Authoring of new sheet images, movies and audio clips for
//! [`crate::sheet::drawable::Edit::add_image`] and its siblings (record
//! 0378): the host editor's `add_sheet_{image,movie,audio}` graphs (the
//! drawable and its title and caption stand-ins, encoded as the host's
//! generated messages encode them), their embedded data (`DataInfo` records
//! and `Data/` members named `stem-identifier.extension`) and their package
//! metadata (UUIDs, data-reference owners, the style's cross-component
//! reference).

use litchi_iwa_common::media::Type;
use litchi_iwa_core::{ArchiveObject, RawMessage};

use super::Package;
use super::sheet_drawable_author::Writer;
use super::sheet_drawable_graph as graph;
use super::structure::{
    self, Fault, FaultResult, Working,
    metadata::{self, DataAddition, DataOwner, Registry, Transition},
    references,
};
use crate::sheet::drawable::{AudioOptions, ImageOptions, MovieOptions, Point, Size};

const IMAGE_MESSAGE_TYPE: u32 = 3_005;
const MOVIE_MESSAGE_TYPE: u32 = 3_007;
const STANDIN_MESSAGE_TYPE: u32 = 3_097;
/// `TN.ThemeArchive`.
const THEME_MESSAGE_TYPE: u32 = 12_009;
/// `TSS.ThemeArchive` drawing presets and their image and movie style lists.
const THEME_DRAWING_FIELD: u32 = 100;
const IMAGE_STYLE_PRESETS_FIELD: u32 = 7;
const MOVIE_STYLE_PRESETS_FIELD: u32 = 8;
const DEFAULT_DRAWABLE_FLAGS: u32 = 3;
const TEXT_WRAP_SQUARE: u32 = 4;
const TEXT_WRAP_BOTH_SIDES: u32 = 2;
const TEXT_WRAP_FIT_TEXT: u32 = 1;
const TEXT_WRAP_MARGIN_POINTS: f32 = 12.0;
const TEXT_WRAP_ALPHA_THRESHOLD: f32 = 0.5;
const STANDARD_VERSION: [u32; 3] = [1, 0, 5];
const STANDIN_VERSION: [u32; 3] = [10, 1, 0];

/// One staged new media drawable.
#[derive(Debug, Clone, PartialEq)]
pub(super) enum NewMedia {
    Image {
        file_name: String,
        data: Vec<u8>,
        options: ImageOptions,
    },
    Movie {
        movie: (String, Vec<u8>),
        poster: (String, Vec<u8>),
        options: MovieOptions,
    },
    Audio {
        file_name: String,
        data: Vec<u8>,
        options: AudioOptions,
    },
}

/// Why a new media drawable is refused.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum Invalid {
    /// The file name, its type or the data's signature does not fit.
    Media,
    /// The geometry, natural size or duration is out of range.
    Placement,
}

/// The validation of one new asset (the host's `validate_new_media` and
/// the media-type check of its `add_sheet_*` route).
fn check(file_name: &str, data: &[u8], expected: Type) -> Result<(), Invalid> {
    let path = std::path::Path::new(file_name);
    if data.is_empty()
        || path.file_name().and_then(|name| name.to_str()) != Some(file_name)
        || file_name.contains(['\0', '\\'])
        || file_name.starts_with('.')
    {
        return Err(Invalid::Media);
    }
    let declared = path
        .extension()
        .and_then(|extension| extension.to_str())
        .map_or(Type::Unknown, Type::from_extension);
    let detected = Type::from_bytes(data);
    if declared != expected || (detected != Type::Unknown && detected != declared) {
        return Err(Invalid::Media);
    }
    Ok(())
}

fn finite(point: Point) -> bool {
    point.x.is_finite() && point.y.is_finite()
}

fn non_negative(size: Size) -> bool {
    size.width.is_finite() && size.height.is_finite() && size.width >= 0.0 && size.height >= 0.0
}

fn positive(size: Size) -> bool {
    non_negative(size) && size.width > 0.0 && size.height > 0.0
}

fn seconds(duration: std::time::Duration) -> Result<f32, Invalid> {
    let seconds = duration.as_secs_f64();
    if seconds == 0.0 || seconds > f64::from(f32::MAX) {
        return Err(Invalid::Placement);
    }
    #[allow(
        clippy::cast_possible_truncation,
        reason = "bounded by f32::MAX above, as the host narrows it"
    )]
    Ok(seconds as f32)
}

impl NewMedia {
    /// Validate the staged drawable as the host's route does.
    pub(super) fn validate(&self) -> Result<(), Invalid> {
        match self {
            Self::Image {
                file_name,
                data,
                options,
            } => {
                check(file_name, data, Type::Image)?;
                if !finite(options.position)
                    || !non_negative(options.size)
                    || !positive(options.natural_size)
                {
                    return Err(Invalid::Placement);
                }
            },
            Self::Movie {
                movie,
                poster,
                options,
            } => {
                check(&movie.0, &movie.1, Type::Video)?;
                check(&poster.0, &poster.1, Type::Image)?;
                if !finite(options.position)
                    || !non_negative(options.size)
                    || !positive(options.natural_size)
                {
                    return Err(Invalid::Placement);
                }
                seconds(options.duration)?;
            },
            Self::Audio {
                file_name,
                data,
                options,
            } => {
                check(file_name, data, Type::Audio)?;
                if !finite(options.position) {
                    return Err(Invalid::Placement);
                }
                seconds(options.duration)?;
            },
        }
        Ok(())
    }

    fn assets(&self) -> Vec<(&str, &[u8])> {
        match self {
            Self::Image {
                file_name, data, ..
            }
            | Self::Audio {
                file_name, data, ..
            } => vec![(file_name.as_str(), data.as_slice())],
            Self::Movie { movie, poster, .. } => vec![
                (movie.0.as_str(), movie.1.as_slice()),
                (poster.0.as_str(), poster.1.as_slice()),
            ],
        }
    }
}

/// The first image or movie style preset of the document theme (the host's
/// `image_style_id` / `movie_style_id`).
fn style_preset(package: &Package, field: u32) -> FaultResult<u64> {
    let document = graph::object(package, 1)?;
    let document = &document.messages[structure::document_message(document)?].data;
    let theme = references::records(document, &[6])?;
    let [theme] = theme.as_slice() else {
        return Err(Fault::Invalid);
    };
    let theme = graph::object(package, references::reference_identifier(theme)?)?;
    let theme = &theme.messages[structure::unique_message(theme, &[THEME_MESSAGE_TYPE])?].data;
    let presets = references::records(theme, &[1, THEME_DRAWING_FIELD, field])
        .map_err(|_error| Fault::Unsupported)?;
    let first = presets.first().ok_or(Fault::Unsupported)?;
    references::reference_identifier(first)
}

/// A new asset's `Data/` file name (the host's `materialized_file_name`).
fn materialized_name(preferred: &str, identifier: u64) -> FaultResult<String> {
    let path = std::path::Path::new(preferred);
    let stem = path
        .file_stem()
        .and_then(|stem| stem.to_str())
        .filter(|stem| !stem.is_empty())
        .ok_or(Fault::Invalid)?;
    Ok(
        match path.extension().and_then(|extension| extension.to_str()) {
            Some(extension) if !extension.is_empty() => format!("{stem}-{identifier}.{extension}"),
            _ => format!("{stem}-{identifier}"),
        },
    )
}

fn size(writer: Writer, number: u32, size: Size) -> Writer {
    writer.message(
        number,
        &Writer::new().float(1, size.width).float(2, size.height),
    )
}

fn data_reference(identifier: u64) -> Writer {
    Writer::new().varint(1, identifier)
}

/// The host's `DrawableArchive` of a new media drawable, in its generated
/// encoder's field order.
fn drawable(
    sheet: u64,
    (title, caption): (u64, u64),
    position: Point,
    displayed: Size,
    aspect_ratio_locked: bool,
) -> Writer {
    let frame = size(
        Writer::new().message(1, &Writer::new().float(1, position.x).float(2, position.y)),
        2,
        displayed,
    )
    .varint(3, u64::from(DEFAULT_DRAWABLE_FLAGS))
    .float(4, 0.0);
    Writer::new()
        .message(1, &frame)
        .reference(2, sheet)
        .message(
            3,
            &Writer::new()
                .varint(1, u64::from(TEXT_WRAP_SQUARE))
                .varint(2, u64::from(TEXT_WRAP_BOTH_SIDES))
                .varint(3, u64::from(TEXT_WRAP_FIT_TEXT))
                .float(4, TEXT_WRAP_MARGIN_POINTS)
                .float(5, TEXT_WRAP_ALPHA_THRESHOLD)
                .boolean(6, false),
        )
        .boolean(5, false)
        .boolean(7, aspect_ratio_locked)
        .reference(10, title)
        .reference(11, caption)
        .boolean(12, false)
        .boolean(13, false)
}

/// The host's `TSD.MovieArchive` of a new movie (`poster`) or audio clip.
fn movie(
    drawable: &Writer,
    style: u64,
    data: u64,
    poster: Option<u64>,
    duration: f32,
    natural: Size,
) -> Writer {
    let mut movie = Writer::new()
        .message(1, drawable)
        .float(3, 0.0)
        .float(4, duration)
        .float(5, 0.0)
        .float(7, 1.0)
        .boolean(9, poster.is_none())
        .varint(13, 0)
        .message(14, &data_reference(data));
    if let Some(poster) = poster {
        movie = movie.message(15, &data_reference(poster));
    }
    let movie = movie.boolean(18, false).reference(19, style);
    size(size(movie, 20, natural), 21, natural)
        .boolean(23, poster.is_some())
        .varint(24, 0)
}

fn new_object(
    identifier: u64,
    message_type: u32,
    data: Vec<u8>,
    version: [u32; 3],
    object_references: Vec<u64>,
    data_references: Vec<u64>,
    limits: litchi_iwa_core::Limits,
) -> FaultResult<ArchiveObject> {
    let mut object = ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_: message_type,
            data,
        }],
        limits,
    )
    .map_err(|error| structure::map_core(&error))?;
    let info = object
        .archive_info
        .message_infos
        .first_mut()
        .ok_or(Fault::Invalid)?;
    info.versions = version.to_vec().into();
    info.object_references = object_references;
    info.data_references = data_references;
    Ok(object)
}

/// Author `media` on the sheet `sheet` stored in `sheet_member`: its objects
/// join the working set, its assets, UUIDs, owners and style reference the
/// transition. Returns the new drawable's identifier.
pub(super) fn author(
    working: &mut Working<'_>,
    registry: &Registry,
    transition: &mut Transition,
    (sheet, sheet_member): (u64, &str),
    media: &NewMedia,
) -> FaultResult<u64> {
    let source = working.source();
    let style = style_preset(
        source,
        if matches!(media, NewMedia::Image { .. }) {
            IMAGE_STYLE_PRESETS_FIELD
        } else {
            MOVIE_STYLE_PRESETS_FIELD
        },
    )?;
    let ids = working.allocate(3)?;
    let [drawable_id, title, caption] = ids.as_slice() else {
        return Err(Fault::Invalid);
    };
    let (drawable_id, title, caption) = (*drawable_id, *title, *caption);
    // The assets.
    let mut data_ids = Vec::new();
    for (file_name, data) in media.assets() {
        use sha1::Digest as _;
        let identifier = metadata::next_data_identifier(working, transition)?;
        let file = materialized_name(file_name, identifier)?;
        let entry = format!("Data/{file}");
        let exists = structure::exact_source(working.source())?
            .package()
            .iter()
            .any(|existing| existing.name() == entry);
        if exists {
            return Err(Fault::Invalid);
        }
        structure::reserve(&mut transition.data_additions, 1)?;
        transition.data_additions.push(DataAddition {
            identifier,
            digest: sha1::Sha1::digest(data).to_vec(),
            preferred: file_name.to_owned(),
            file,
            length: u64::try_from(data.len()).map_err(|_error| Fault::Invalid)?,
        });
        structure::reserve(&mut transition.entry_insertions, 1)?;
        transition
            .entry_insertions
            .push((entry, structure::copy_bytes(data)?));
        structure::reserve(&mut data_ids, 1)?;
        data_ids.push(identifier);
    }
    let labels = (title, caption);
    let payload = match media {
        NewMedia::Image { options, .. } => {
            let drawable = drawable(sheet, labels, options.position, options.size, false);
            let image = Writer::new().message(1, &drawable).reference(3, style);
            size(
                size(image, 4, options.natural_size).varint(7, 0),
                9,
                options.natural_size,
            )
            .message(11, &data_reference(data_ids[0]))
            .finish()
        },
        NewMedia::Movie { options, .. } => movie(
            &drawable(sheet, labels, options.position, options.size, true),
            style,
            data_ids[0],
            Some(data_ids[1]),
            seconds(options.duration).map_err(|_error| Fault::Invalid)?,
            options.natural_size,
        )
        .finish(),
        NewMedia::Audio { options, .. } => {
            let zero = Size {
                width: 0.0,
                height: 0.0,
            };
            movie(
                &drawable(sheet, labels, options.position, zero, true),
                style,
                data_ids[0],
                None,
                seconds(options.duration).map_err(|_error| Fault::Invalid)?,
                zero,
            )
            .finish()
        },
    };
    let message_type = if matches!(media, NewMedia::Image { .. }) {
        IMAGE_MESSAGE_TYPE
    } else {
        MOVIE_MESSAGE_TYPE
    };
    let limits = working.archive_limits();
    let objects = [
        new_object(
            drawable_id,
            message_type,
            payload,
            STANDARD_VERSION,
            vec![title, caption, style],
            data_ids.clone(),
            limits,
        )?,
        new_object(
            title,
            STANDIN_MESSAGE_TYPE,
            Vec::new(),
            STANDIN_VERSION,
            Vec::new(),
            Vec::new(),
            limits,
        )?,
        new_object(
            caption,
            STANDIN_MESSAGE_TYPE,
            Vec::new(),
            STANDIN_VERSION,
            Vec::new(),
            Vec::new(),
            limits,
        )?,
    ];
    for object in objects {
        working.insert(sheet_member, object)?;
    }
    let component = registry.component_for_entry(sheet_member)?.identifier;
    for identifier in [drawable_id, title, caption] {
        let seed = metadata::identity_seed(working, identifier)?;
        transition.register_uuid(registry, component, identifier, seed)?;
    }
    for data in data_ids {
        structure::reserve(&mut transition.data_owner_additions, 1)?;
        transition.data_owner_additions.push(DataOwner {
            component,
            data,
            object: drawable_id,
            count: 1,
        });
    }
    let style_member = working.locate(style)?.ok_or(Fault::Invalid)?;
    let style_component = registry.component_for_entry(&style_member)?.identifier;
    transition.reference(registry, component, style_component, style)?;
    Ok(drawable_id)
}
