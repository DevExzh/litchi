//! Typed construction and strict discovery of sheet-owned Numbers movie graphs.

use super::*;
use crate::IWorkThemeArchive;
use crate::media_playback::media_playback_settings;
use crate::shapes::{
    DrawableProperties, drawable_properties, geometry_from_drawable, patch_drawable_geometry,
    patch_wrapped_drawable_properties,
};

const NUMBERS_THEME_MESSAGE_TYPE: u32 = 12_009;
const MOVIE_MESSAGE_TYPE: u32 = 3_007;
const STANDIN_CAPTION_MESSAGE_TYPE: u32 = 3_097;
const DEFAULT_DRAWABLE_FLAGS: u32 = 3;
const DEFAULT_MOVIE_FLAGS: u32 = 0;
const DEFAULT_MOVIE_ROTATION_DEGREES: f32 = 0.0;
const DEFAULT_MOVIE_VOLUME: f32 = 1.0;
const DEFAULT_TEXT_WRAP_MARGIN_POINTS: f32 = 12.0;
const DEFAULT_TEXT_WRAP_ALPHA_THRESHOLD: f32 = 0.5;
const STANDARD_MESSAGE_VERSION: [u32; 3] = [1, 0, 5];
const STANDIN_CAPTION_MESSAGE_VERSION: [u32; 3] = [10, 1, 0];

#[derive(Debug, Clone, Copy)]
#[repr(u32)]
enum TextWrapType {
    Square = 4,
}

#[derive(Debug, Clone, Copy)]
#[repr(u32)]
enum TextWrapDirection {
    BothSides = 2,
}

#[derive(Debug, Clone, Copy)]
#[repr(u32)]
enum TextWrapFit {
    Text = 1,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(in crate::numbers::editor) struct MovieObjectIds {
    pub(in crate::numbers::editor) drawable: u64,
    pub(in crate::numbers::editor) title: u64,
    pub(in crate::numbers::editor) caption: u64,
}

impl MovieObjectIds {
    pub(in crate::numbers::editor) fn allocate(first: u64) -> Result<Self> {
        let identifier = |offset: u64| {
            first
                .checked_add(offset)
                .ok_or_else(|| Error::ParseError("iWork object identifier overflow".to_owned()))
        };
        Ok(Self {
            drawable: identifier(0)?,
            title: identifier(1)?,
            caption: identifier(2)?,
        })
    }

    pub(in crate::numbers::editor) const fn last(self) -> u64 {
        self.caption
    }

    pub(in crate::numbers::editor) const fn all(self) -> [u64; 3] {
        [self.drawable, self.title, self.caption]
    }
}

pub(in crate::numbers::editor) struct MovieCreationContext {
    pub(in crate::numbers::editor) archive_name: String,
    pub(in crate::numbers::editor) component_id: u64,
    pub(in crate::numbers::editor) style_id: u64,
    pub(in crate::numbers::editor) stylesheet_component_id: u64,
}

pub(super) struct SheetMovieGraph {
    pub(super) sheet_id: u64,
    pub(super) archive_name: String,
    pub(super) component_id: u64,
    pub(super) info: NumbersSheetMovieInfo,
    pub(super) object_ids: Vec<u64>,
    pub(super) uuid_object_ids: Vec<u64>,
    pub(super) data_references: Vec<(u64, u64)>,
}

pub(super) fn movie_creation_values(
    options: NumbersSheetMovieOptions,
) -> Result<(DrawableGeometry, f32)> {
    if !options.natural_size.width.is_finite()
        || !options.natural_size.height.is_finite()
        || options.natural_size.width <= 0.0
        || options.natural_size.height <= 0.0
    {
        return Err(Error::ParseError(
            "Numbers movie natural size must be finite and greater than zero".to_owned(),
        ));
    }
    let duration_seconds = options.duration.as_secs_f64();
    if duration_seconds == 0.0 || duration_seconds > f64::from(f32::MAX) {
        return Err(Error::ParseError(
            "Numbers movie duration must be greater than zero and fit in f32 seconds".to_owned(),
        ));
    }
    let geometry = DrawableGeometry {
        position: Some(options.position),
        size: Some(options.size),
        flags: Some(DEFAULT_DRAWABLE_FLAGS),
        angle: Some(DEFAULT_MOVIE_ROTATION_DEGREES),
    }
    .validate()?;
    Ok((geometry, duration_seconds as f32))
}

pub(in crate::numbers::editor) fn movie_creation_context(
    editor: &NumbersEditor,
    sheet_id: u64,
) -> Result<MovieCreationContext> {
    let (archive_name, _, _) = numbers_sheet(editor.package(), sheet_id)?;
    let document = numbers_document(editor.package())?;
    let style_id = movie_style_id(editor.package(), document.theme.identifier)?;
    let component_id = component_identifier_for_entry(editor.package(), &archive_name)?
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers sheet component {archive_name} is not registered"
            ))
        })?;
    let style_archive = object_locations(editor.package())?
        .get(&style_id)
        .cloned()
        .ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers movie style {style_id} is missing"))
        })?;
    let stylesheet_component_id = component_identifier_for_entry(editor.package(), &style_archive)?
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers stylesheet component {style_archive} is not registered"
            ))
        })?;
    Ok(MovieCreationContext {
        archive_name,
        component_id,
        style_id,
        stylesheet_component_id,
    })
}

fn movie_style_id(package: &IWorkPackage, theme_id: u64) -> Result<u64> {
    let locations = object_locations(package)?;
    let archive_name = locations.get(&theme_id).ok_or_else(|| {
        Error::InvalidFormat(format!("Numbers theme object {theme_id} is missing"))
    })?;
    let archive = package.archive(archive_name)?;
    let object = archive.object(theme_id).ok_or_else(|| {
        Error::InvalidFormat(format!("Numbers theme object {theme_id} is missing"))
    })?;
    let messages = object
        .messages
        .iter()
        .filter(|message| message.type_ == NUMBERS_THEME_MESSAGE_TYPE)
        .collect::<Vec<_>>();
    let [message] = messages.as_slice() else {
        return Err(Error::InvalidFormat(format!(
            "Numbers theme object {theme_id} must have exactly one theme payload"
        )));
    };
    IWorkThemeArchive::decode(&message.data)?
        .extensions
        .drawing
        .and_then(|presets| presets.movie_style_presets.into_iter().next())
        .map(|reference| reference.identifier)
        .ok_or_else(|| Error::InvalidFormat("Numbers theme has no movie style preset".to_owned()))
}

pub(super) fn movie_infos(
    editor: &NumbersEditor,
    sheet_id: u64,
) -> Result<Vec<NumbersSheetMovieInfo>> {
    let (_, _, sheet) = numbers_sheet(editor.package(), sheet_id)?;
    let locations = object_locations(editor.package())?;
    let mut movies = Vec::new();
    for reference in sheet.drawable_infos {
        let archive_name = locations.get(&reference.identifier).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers sheet {sheet_id} drawable {} is missing",
                reference.identifier
            ))
        })?;
        let archive = editor.package().archive(archive_name)?;
        let object = archive.object(reference.identifier).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers sheet {sheet_id} drawable {} is missing",
                reference.identifier
            ))
        })?;
        let movie_messages = object
            .messages
            .iter()
            .filter(|message| message.type_ == MOVIE_MESSAGE_TYPE)
            .collect::<Vec<_>>();
        if movie_messages.is_empty() {
            continue;
        }
        if movie_messages.len() != 1 {
            return Err(Error::InvalidFormat(format!(
                "Numbers drawable {} has multiple movie payloads",
                reference.identifier
            )));
        }
        let movie = tsd::MovieArchive::decode(movie_messages[0].data.as_slice())?;
        if movie.audio_only == Some(true) || movie.is_live_video == Some(true) {
            continue;
        }
        movies.push(movie_info(
            editor.package(),
            sheet_id,
            reference.identifier,
        )?);
    }
    Ok(movies)
}

pub(super) fn movie_graph(
    editor: &NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
) -> Result<SheetMovieGraph> {
    let (archive_name, _, sheet) = numbers_sheet(editor.package(), sheet_id)?;
    if sheet
        .drawable_infos
        .iter()
        .filter(|reference| reference.identifier == drawable_object_id)
        .count()
        != 1
    {
        return Err(Error::ParseError(format!(
            "Numbers sheet {sheet_id} does not own movie {drawable_object_id} exactly once"
        )));
    }
    let locations = object_locations(editor.package())?;
    if locations.get(&drawable_object_id).map(String::as_str) != Some(archive_name.as_str()) {
        return Err(Error::InvalidFormat(format!(
            "Numbers movie {drawable_object_id} is outside sheet component {archive_name}"
        )));
    }
    let archive = editor.package().archive(&archive_name)?;
    let object = archive.object(drawable_object_id).ok_or_else(|| {
        Error::InvalidFormat(format!("Numbers movie {drawable_object_id} is missing"))
    })?;
    let messages = object
        .messages
        .iter()
        .filter(|message| message.type_ == MOVIE_MESSAGE_TYPE)
        .collect::<Vec<_>>();
    let [message] = messages.as_slice() else {
        return Err(Error::ParseError(format!(
            "Numbers drawable {drawable_object_id} is not an ordinary movie"
        )));
    };
    let movie = tsd::MovieArchive::decode(message.data.as_slice())?;
    if movie.audio_only == Some(true) || movie.is_live_video == Some(true) {
        return Err(Error::ParseError(format!(
            "Numbers drawable {drawable_object_id} is not an ordinary file-backed movie"
        )));
    }
    if movie.super_.parent.map(|parent| parent.identifier) != Some(sheet_id) {
        return Err(Error::InvalidFormat(format!(
            "Numbers movie {drawable_object_id} is not owned by sheet {sheet_id}"
        )));
    }
    let title_id = required_reference(drawable_object_id, movie.super_.title, "title stand-in")?;
    let caption_id =
        required_reference(drawable_object_id, movie.super_.caption, "caption stand-in")?;
    let style_id = required_reference(drawable_object_id, movie.style, "movie style")?;
    if !locations.contains_key(&style_id) {
        return Err(Error::InvalidFormat(format!(
            "Numbers movie {drawable_object_id} style {style_id} is missing"
        )));
    }
    let object_ids = vec![drawable_object_id, title_id, caption_id];
    if object_ids.iter().copied().collect::<HashSet<_>>().len() != object_ids.len() {
        return Err(Error::InvalidFormat(format!(
            "Numbers movie {drawable_object_id} reuses private graph identifiers"
        )));
    }
    for identifier in [title_id, caption_id] {
        if locations.get(&identifier).map(String::as_str) != Some(archive_name.as_str()) {
            return Err(Error::InvalidFormat(format!(
                "Numbers movie {drawable_object_id} private graph spans multiple archives"
            )));
        }
        let standin = archive.object(identifier).ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers movie stand-in {identifier} is missing"))
        })?;
        if standin
            .messages
            .iter()
            .filter(|message| message.type_ == STANDIN_CAPTION_MESSAGE_TYPE)
            .count()
            != 1
        {
            return Err(Error::InvalidFormat(format!(
                "Numbers movie stand-in {identifier} is malformed"
            )));
        }
    }
    let mut allowed_references = [sheet_id, title_id, caption_id, style_id]
        .into_iter()
        .collect::<HashSet<_>>();
    allowed_references.extend(movie.super_.comment.map(|reference| reference.identifier));
    let unexpected_references = object
        .archive_info
        .message_infos
        .iter()
        .flat_map(|info| {
            info.object_references.iter().chain(
                info.field_infos
                    .iter()
                    .flat_map(|field| &field.object_references),
            )
        })
        .copied()
        .filter(|identifier| !allowed_references.contains(identifier))
        .collect::<HashSet<_>>();
    if !unexpected_references.is_empty() {
        return Err(Error::ParseError(format!(
            "Numbers movie {drawable_object_id} has unsupported private references {unexpected_references:?}"
        )));
    }
    let component_id = component_identifier_for_entry(editor.package(), &archive_name)?
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers sheet component {archive_name} is not registered"
            ))
        })?;
    let registered =
        component_uuid_identifiers(editor.package(), component_id)?.unwrap_or_default();
    let uuid_object_ids = object_ids
        .iter()
        .copied()
        .filter(|identifier| registered.contains(identifier))
        .collect::<Vec<_>>();
    let mut data_references = object
        .archive_info
        .message_infos
        .iter()
        .flat_map(|info| {
            info.data_references.iter().chain(
                info.field_infos
                    .iter()
                    .flat_map(|field| &field.data_references),
            )
        })
        .copied()
        .map(|data_identifier| (data_identifier, drawable_object_id))
        .collect::<Vec<_>>();
    data_references.sort_unstable();
    data_references.dedup();
    let info = movie_info(editor.package(), sheet_id, drawable_object_id)?;
    for identifier in [
        info.movie_data_identifier,
        info.poster_image_data_identifier,
    ] {
        if !data_references.contains(&(identifier, drawable_object_id)) {
            return Err(Error::InvalidFormat(format!(
                "Numbers movie {drawable_object_id} data {identifier} is missing from archive metadata"
            )));
        }
    }
    Ok(SheetMovieGraph {
        sheet_id,
        archive_name,
        component_id,
        info,
        object_ids,
        uuid_object_ids,
        data_references,
    })
}

fn movie_info(
    package: &IWorkPackage,
    sheet_id: u64,
    identifier: u64,
) -> Result<NumbersSheetMovieInfo> {
    let locations = object_locations(package)?;
    let archive_name = locations
        .get(&identifier)
        .ok_or_else(|| Error::InvalidFormat(format!("Numbers movie {identifier} is missing")))?;
    let archive = package.archive(archive_name)?;
    let object = archive
        .object(identifier)
        .ok_or_else(|| Error::InvalidFormat(format!("Numbers movie {identifier} is missing")))?;
    let messages = object
        .messages
        .iter()
        .filter(|message| message.type_ == MOVIE_MESSAGE_TYPE)
        .collect::<Vec<_>>();
    let [message] = messages.as_slice() else {
        return Err(Error::InvalidFormat(format!(
            "Numbers movie {identifier} must have exactly one movie payload"
        )));
    };
    let movie = tsd::MovieArchive::decode(message.data.as_slice())?;
    if movie.audio_only == Some(true) || movie.is_live_video == Some(true) {
        return Err(Error::ParseError(format!(
            "Numbers drawable {identifier} is not an ordinary file-backed movie"
        )));
    }
    if movie.super_.parent.map(|parent| parent.identifier) != Some(sheet_id) {
        return Err(Error::InvalidFormat(format!(
            "Numbers movie {identifier} is not owned by sheet {sheet_id}"
        )));
    }
    let movie_data_identifier = movie
        .movie_data
        .ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers movie {identifier} has no video data"))
        })?
        .identifier;
    let poster_image_data_identifier = movie
        .poster_image_data
        .ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers movie {identifier} has no poster data"))
        })?
        .identifier;
    let playback = media_playback_settings(&movie).map_err(|error| {
        Error::InvalidFormat(format!(
            "Numbers movie {identifier} has invalid playback settings: {error}"
        ))
    })?;
    Ok(NumbersSheetMovieInfo {
        sheet_id,
        drawable_object_id: identifier,
        movie_data_identifier,
        poster_image_data_identifier,
        geometry: geometry_from_drawable(&movie.super_)?,
        properties: drawable_properties(&movie.super_),
        playback,
        original_size: movie.original_size.map(drawable_size),
        natural_size: movie.natural_size.map(drawable_size),
        duration: playback.duration(),
    })
}

#[allow(deprecated)]
pub(super) fn movie_objects(
    ids: MovieObjectIds,
    sheet_id: u64,
    style_id: u64,
    movie_data_identifier: u64,
    poster_data_identifier: u64,
    geometry: DrawableGeometry,
    natural_size: DrawableSize,
    duration_seconds: f32,
) -> Result<[ArchiveObject; 3]> {
    let position = geometry.position.ok_or_else(|| {
        Error::InvalidFormat("validated Numbers movie geometry has no position".to_owned())
    })?;
    let size = geometry.size.ok_or_else(|| {
        Error::InvalidFormat("validated Numbers movie geometry has no size".to_owned())
    })?;
    let movie = tsd::MovieArchive {
        super_: tsd::DrawableArchive {
            geometry: Some(tsd::GeometryArchive {
                position: Some(tsp::Point {
                    x: position.x,
                    y: position.y,
                }),
                size: Some(tsp::Size {
                    width: size.width,
                    height: size.height,
                }),
                flags: geometry.flags,
                angle: geometry.angle,
            }),
            parent: Some(reference(sheet_id)),
            exterior_text_wrap: Some(tsd::ExteriorTextWrapArchive {
                r#type: Some(TextWrapType::Square as u32),
                direction: Some(TextWrapDirection::BothSides as u32),
                fit_type: Some(TextWrapFit::Text as u32),
                margin: Some(DEFAULT_TEXT_WRAP_MARGIN_POINTS),
                alpha_threshold: Some(DEFAULT_TEXT_WRAP_ALPHA_THRESHOLD),
                is_html_wrap: Some(false),
            }),
            locked: Some(false),
            aspect_ratio_locked: Some(true),
            title: Some(reference(ids.title)),
            caption: Some(reference(ids.caption)),
            title_hidden: Some(false),
            caption_hidden: Some(false),
            ..Default::default()
        },
        movie_data: Some(tsp::DataReference {
            identifier: movie_data_identifier,
        }),
        start_time: Some(0.0),
        end_time: Some(duration_seconds),
        poster_time: Some(0.0),
        loop_option: Some(tsd::movie_archive::MovieLoopOption::None as i32),
        volume: Some(DEFAULT_MOVIE_VOLUME),
        audio_only: Some(false),
        streaming: Some(false),
        poster_image_data: Some(tsp::DataReference {
            identifier: poster_data_identifier,
        }),
        poster_image_generated_with_alpha_support: Some(true),
        flags: Some(DEFAULT_MOVIE_FLAGS),
        style: Some(reference(style_id)),
        original_size: Some(tsp::Size {
            width: natural_size.width,
            height: natural_size.height,
        }),
        natural_size: Some(tsp::Size {
            width: natural_size.width,
            height: natural_size.height,
        }),
        ..Default::default()
    };
    Ok([
        numbers_movie_object(
            ids.drawable,
            MOVIE_MESSAGE_TYPE,
            movie,
            &STANDARD_MESSAGE_VERSION,
            &[ids.title, ids.caption, style_id],
            &[movie_data_identifier, poster_data_identifier],
        )?,
        numbers_movie_object(
            ids.title,
            STANDIN_CAPTION_MESSAGE_TYPE,
            tsd::StandinCaptionArchive::default(),
            &STANDIN_CAPTION_MESSAGE_VERSION,
            &[],
            &[],
        )?,
        numbers_movie_object(
            ids.caption,
            STANDIN_CAPTION_MESSAGE_TYPE,
            tsd::StandinCaptionArchive::default(),
            &STANDIN_CAPTION_MESSAGE_VERSION,
            &[],
            &[],
        )?,
    ])
}

pub(in crate::numbers::editor) fn set_movie_geometry(
    package: &mut IWorkPackage,
    archive_name: &str,
    movie_id: u64,
    geometry: DrawableGeometry,
) -> Result<()> {
    geometry.validate()?;
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(movie_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers movie object {movie_id} is missing"))
        })?;
        let indexes = object
            .messages
            .iter()
            .enumerate()
            .filter(|(_, message)| message.type_ == MOVIE_MESSAGE_TYPE)
            .map(|(index, _)| index)
            .collect::<Vec<_>>();
        let [message_index] = indexes.as_slice() else {
            return Err(Error::InvalidFormat(format!(
                "Numbers movie {movie_id} must have exactly one MovieArchive payload"
            )));
        };
        let data = transform_length_delimited_field(
            object.messages[*message_index].data.as_slice(),
            1,
            |drawable| patch_drawable_geometry(drawable, geometry),
        )?;
        object.replace_message(
            *message_index,
            RawMessage {
                type_: MOVIE_MESSAGE_TYPE,
                data,
            },
        )?;
        Ok(())
    })
}

pub(in crate::numbers::editor) fn set_movie_properties(
    package: &mut IWorkPackage,
    archive_name: &str,
    movie_id: u64,
    properties: &DrawableProperties,
) -> Result<()> {
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(movie_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers movie object {movie_id} is missing"))
        })?;
        let indexes = object
            .messages
            .iter()
            .enumerate()
            .filter(|(_, message)| message.type_ == MOVIE_MESSAGE_TYPE)
            .map(|(index, _)| index)
            .collect::<Vec<_>>();
        let [message_index] = indexes.as_slice() else {
            return Err(Error::InvalidFormat(format!(
                "Numbers movie {movie_id} must have exactly one MovieArchive payload"
            )));
        };
        let original = object.messages[*message_index].data.as_slice();
        let current = drawable_properties(&tsd::MovieArchive::decode(original)?.super_);
        let data = patch_wrapped_drawable_properties(original, &current, properties)?;
        let verified = tsd::MovieArchive::decode(data.as_slice())?;
        if drawable_properties(&verified.super_) != *properties {
            return Err(Error::InvalidFormat(
                "Numbers movie properties patch failed validation".to_owned(),
            ));
        }
        object.replace_message(
            *message_index,
            RawMessage {
                type_: MOVIE_MESSAGE_TYPE,
                data,
            },
        )?;
        Ok(())
    })
}

fn required_reference(
    movie_id: u64,
    reference: Option<tsp::Reference>,
    label: &str,
) -> Result<u64> {
    reference
        .map(|reference| reference.identifier)
        .ok_or_else(|| Error::InvalidFormat(format!("Numbers movie {movie_id} has no {label}")))
}

fn numbers_movie_object(
    identifier: u64,
    message_type: u32,
    message: impl Message,
    versions: &[u32],
    object_references: &[u64],
    data_references: &[u64],
) -> Result<ArchiveObject> {
    let mut object = ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_: message_type,
            data: message.encode_to_vec(),
        }],
    )?;
    let info = &mut object.archive_info.message_infos[0];
    info.versions = versions.to_vec();
    info.object_references = object_references.to_vec();
    info.data_references = data_references.to_vec();
    Ok(object)
}

fn drawable_size(size: tsp::Size) -> DrawableSize {
    DrawableSize {
        width: size.width,
        height: size.height,
    }
}

fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..Default::default()
    }
}
