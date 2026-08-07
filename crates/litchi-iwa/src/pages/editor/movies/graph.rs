//! Typed construction and strict discovery of body-anchored Pages movie graphs.

use super::*;
use crate::IWorkThemeArchive;
use crate::media_playback::media_playback_settings;
use crate::shapes::{
    DrawableProperties, drawable_properties, geometry_from_drawable, patch_drawable_geometry,
    patch_wrapped_drawable_properties,
};
use litchi_pages::movie::Options as PagesMovieOptions;

const THEME_MESSAGE_TYPE: u32 = 10_001;
const DRAWABLE_Z_ORDER_MESSAGE_TYPE: u32 = 10_015;
const MOVIE_MESSAGE_TYPE: u32 = 3_007;
const STANDIN_CAPTION_MESSAGE_TYPE: u32 = 3_097;
const MOVIE_DRAWABLE_FIELD: u32 = 1;
const ATTACHMENT_HORIZONTAL_OFFSET_FIELD: u32 = 3;
const ATTACHMENT_VERTICAL_OFFSET_FIELD: u32 = 5;
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

#[derive(Debug, Clone, Copy)]
#[repr(u32)]
enum HorizontalAnchorBasis {
    BodyMargin = 0,
}

#[derive(Debug, Clone, Copy)]
#[repr(u32)]
enum VerticalAnchorBasis {
    Page = 1,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct MovieObjectIds {
    pub(super) drawable: u64,
    title: u64,
    caption: u64,
    pub(super) attachment: u64,
}

impl MovieObjectIds {
    pub(super) fn allocate(first: u64) -> Result<Self> {
        let identifier = |offset: u64| {
            first
                .checked_add(offset)
                .ok_or_else(|| Error::ParseError("iWork object identifier overflow".to_owned()))
        };
        Ok(Self {
            drawable: identifier(0)?,
            title: identifier(1)?,
            caption: identifier(2)?,
            attachment: identifier(3)?,
        })
    }

    pub(super) const fn last(self) -> u64 {
        self.attachment
    }

    pub(super) const fn all(self) -> [u64; 4] {
        [self.drawable, self.title, self.caption, self.attachment]
    }

    pub(super) const fn uuid_objects(self) -> [u64; 3] {
        [self.drawable, self.title, self.caption]
    }
}

pub(super) struct BodyMovieGraph {
    pub(super) archive_name: String,
    pub(super) attachment_id: u64,
    pub(super) info: PagesMovieInfo,
    pub(super) object_ids: Vec<u64>,
    pub(super) uuid_object_ids: Vec<u64>,
    pub(super) data_references: Vec<(u64, u64)>,
}

pub(super) fn movie_creation_values(options: PagesMovieOptions) -> Result<(DrawableGeometry, f32)> {
    let geometry = DrawableGeometry {
        position: Some(options.position()),
        size: Some(options.size()),
        flags: Some(DEFAULT_DRAWABLE_FLAGS),
        angle: Some(DEFAULT_MOVIE_ROTATION_DEGREES),
    }
    .validate()?;
    Ok((geometry, options.duration_seconds()))
}

pub(super) fn movie_style_id(package: &IWorkPackage, root: &DocumentArchive) -> Result<u64> {
    let theme_id = root
        .theme
        .as_ref()
        .ok_or_else(|| Error::InvalidFormat("Pages document has no theme".to_owned()))?
        .identifier;
    let archive_name = find_object_archive(package, theme_id)?;
    let archive = package.archive(&archive_name)?;
    let object = archive
        .object(theme_id)
        .ok_or_else(|| Error::InvalidFormat(format!("Pages theme {theme_id} is missing")))?;
    let messages = object
        .messages
        .iter()
        .filter(|message| message.type_ == THEME_MESSAGE_TYPE)
        .collect::<Vec<_>>();
    let [message] = messages.as_slice() else {
        return Err(Error::InvalidFormat(format!(
            "Pages theme {theme_id} must have exactly one theme payload"
        )));
    };
    IWorkThemeArchive::decode(&message.data)?
        .extensions
        .drawing
        .and_then(|presets| presets.movie_style_presets.into_iter().next())
        .map(|reference| reference.identifier)
        .ok_or_else(|| Error::InvalidFormat("Pages theme has no movie style preset".to_owned()))
}

pub(super) fn body_movie_infos(editor: &PagesEditor) -> Result<Vec<PagesMovieInfo>> {
    let body: StorageArchive = decode_typed_package_object(
        editor.package(),
        editor.body_storage_id.get(),
        editor.body_storage()?.message_type,
        "TSWP.StorageArchive",
    )?;
    let mut movies = Vec::new();
    for entry in body
        .table_attachment
        .as_ref()
        .into_iter()
        .flat_map(|table| &table.entries)
    {
        let Some(attachment_reference) = entry.object else {
            continue;
        };
        if !object_has_message_type(
            editor.package(),
            attachment_reference.identifier,
            DRAWABLE_ATTACHMENT_MESSAGE_TYPE,
        )? {
            continue;
        }
        let attachment: DrawableAttachmentArchive = decode_typed_package_object(
            editor.package(),
            attachment_reference.identifier,
            DRAWABLE_ATTACHMENT_MESSAGE_TYPE,
            "TSWP.DrawableAttachmentArchive",
        )?;
        let Some(drawable) = attachment.drawable else {
            continue;
        };
        let archive_name = find_object_archive(editor.package(), drawable.identifier)?;
        let archive = editor.package().archive(&archive_name)?;
        let object = archive.object(drawable.identifier).ok_or_else(|| {
            Error::InvalidFormat(format!("Pages drawable {} is missing", drawable.identifier))
        })?;
        let messages = object
            .messages
            .iter()
            .filter(|message| message.type_ == MOVIE_MESSAGE_TYPE)
            .collect::<Vec<_>>();
        if messages.is_empty() {
            continue;
        }
        let [message] = messages.as_slice() else {
            return Err(Error::InvalidFormat(format!(
                "Pages drawable {} has multiple movie payloads",
                drawable.identifier
            )));
        };
        let movie = tsd::MovieArchive::decode(message.data.as_slice())?;
        if movie.audio_only == Some(true) || movie.is_live_video == Some(true) {
            continue;
        }
        movies.push(movie_info(
            editor.package(),
            editor.body_storage_id.get(),
            drawable.identifier,
            entry.character_index,
        )?);
    }
    movies.sort_by_key(|movie| movie.anchor_character_index);
    let mut drawable_ids = HashSet::with_capacity(movies.len());
    if let Some(duplicate) = movies
        .iter()
        .map(|movie| movie.drawable_object_id)
        .find(|identifier| !drawable_ids.insert(*identifier))
    {
        return Err(Error::InvalidFormat(format!(
            "Pages movie {duplicate} has multiple body attachments"
        )));
    }
    Ok(movies)
}

pub(super) fn body_movie_graph(
    editor: &PagesEditor,
    drawable_object_id: u64,
) -> Result<BodyMovieGraph> {
    let info = body_movie_infos(editor)?
        .into_iter()
        .find(|movie| movie.drawable_object_id == drawable_object_id)
        .ok_or_else(|| {
            Error::ParseError(format!(
                "Pages drawable {drawable_object_id} is not a body-anchored movie"
            ))
        })?;
    let body: StorageArchive = decode_typed_package_object(
        editor.package(),
        editor.body_storage_id.get(),
        editor.body_storage()?.message_type,
        "TSWP.StorageArchive",
    )?;
    let mut attachments = Vec::new();
    for entry in body
        .table_attachment
        .as_ref()
        .into_iter()
        .flat_map(|table| &table.entries)
        .filter(|entry| entry.character_index == info.anchor_character_index)
    {
        let Some(reference) = entry.object else {
            continue;
        };
        if !object_has_message_type(
            editor.package(),
            reference.identifier,
            DRAWABLE_ATTACHMENT_MESSAGE_TYPE,
        )? {
            continue;
        }
        let attachment: DrawableAttachmentArchive = decode_typed_package_object(
            editor.package(),
            reference.identifier,
            DRAWABLE_ATTACHMENT_MESSAGE_TYPE,
            "TSWP.DrawableAttachmentArchive",
        )?;
        if attachment
            .drawable
            .is_some_and(|drawable| drawable.identifier == drawable_object_id)
        {
            attachments.push((reference.identifier, attachment));
        }
    }
    let [(attachment_id, attachment)] = attachments.as_slice() else {
        return Err(Error::InvalidFormat(format!(
            "Pages movie {drawable_object_id} has {} body attachments; expected one",
            attachments.len()
        )));
    };
    let body_units = editor.body_text()?.encode_utf16().collect::<Vec<_>>();
    if body_units.get(info.anchor_character_index as usize) != Some(&0xfffc) {
        return Err(Error::InvalidFormat(format!(
            "Pages movie {drawable_object_id} attachment is not backed by an object-replacement character"
        )));
    }

    let movie: tsd::MovieArchive = decode_typed_package_object(
        editor.package(),
        drawable_object_id,
        MOVIE_MESSAGE_TYPE,
        "TSD.MovieArchive",
    )?;
    if movie.audio_only == Some(true) || movie.is_live_video == Some(true) {
        return Err(Error::ParseError(format!(
            "Pages drawable {drawable_object_id} is not an ordinary file-backed movie"
        )));
    }
    if movie.super_.parent.map(|parent| parent.identifier) != Some(editor.body_storage_id.get()) {
        return Err(Error::InvalidFormat(format!(
            "Pages movie {drawable_object_id} is not owned by the body storage"
        )));
    }

    if attachment.h_offset_type != Some(HorizontalAnchorBasis::BodyMargin as u32)
        || attachment.v_offset_type != Some(VerticalAnchorBasis::Page as u32)
    {
        return Err(Error::ParseError(format!(
            "Pages movie {drawable_object_id} uses unsupported attachment offset bases"
        )));
    }
    let h_offset = attachment.h_offset.ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Pages movie {drawable_object_id} attachment has no horizontal offset"
        ))
    })?;
    let v_offset = attachment.v_offset.ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Pages movie {drawable_object_id} attachment has no vertical offset"
        ))
    })?;
    if !h_offset.is_finite() || !v_offset.is_finite() {
        return Err(Error::InvalidFormat(format!(
            "Pages movie {drawable_object_id} attachment offsets must be finite"
        )));
    }

    let document = root_document(editor.package())?;
    let attachment_position = DrawablePoint {
        x: h_offset + document.left_margin.unwrap_or_default(),
        y: v_offset,
    };
    if info.geometry.position != Some(attachment_position) {
        return Err(Error::InvalidFormat(format!(
            "Pages movie {drawable_object_id} geometry and attachment positions disagree"
        )));
    }
    let z_order_id = document.drawables_zorder.ok_or_else(|| {
        Error::InvalidFormat("Pages document has no drawable z-order object".to_owned())
    })?;
    let z_order: tp::DrawablesZOrderArchive = decode_typed_package_object(
        editor.package(),
        z_order_id.identifier,
        DRAWABLE_Z_ORDER_MESSAGE_TYPE,
        "TP.DrawablesZOrderArchive",
    )?;
    let z_order_count = z_order
        .drawables
        .iter()
        .filter(|reference| reference.identifier == drawable_object_id)
        .count();
    if z_order_count != 1 {
        return Err(Error::InvalidFormat(format!(
            "Pages movie {drawable_object_id} occurs {z_order_count} times in drawable z-order"
        )));
    }

    let title = movie_caption_slot_from_reference(
        editor.package(),
        drawable_object_id,
        movie.super_.title,
        crate::image_caption::DrawableCaptionKind::Title,
    )?;
    let caption = movie_caption_slot_from_reference(
        editor.package(),
        drawable_object_id,
        movie.super_.caption,
        crate::image_caption::DrawableCaptionKind::Caption,
    )?;
    let style_id = required_reference(drawable_object_id, movie.style, "movie style")?;
    find_object_archive(editor.package(), style_id)?;
    let mut object_ids = vec![drawable_object_id];
    object_ids.extend(title.object_ids.iter().copied());
    object_ids.extend(caption.object_ids.iter().copied());
    object_ids.push(*attachment_id);
    if object_ids.iter().copied().collect::<HashSet<_>>().len() != object_ids.len() {
        return Err(Error::InvalidFormat(format!(
            "Pages movie {drawable_object_id} reuses private graph identifiers"
        )));
    }
    let archive_name = find_object_archive(editor.package(), drawable_object_id)?;
    for identifier in &object_ids {
        if find_object_archive(editor.package(), *identifier)? != archive_name {
            return Err(Error::InvalidFormat(format!(
                "Pages movie {drawable_object_id} private graph spans multiple archives"
            )));
        }
    }
    let archive = editor.package().archive(&archive_name)?;
    let drawable = archive.object(drawable_object_id).ok_or_else(|| {
        Error::InvalidFormat(format!("Pages movie {drawable_object_id} is missing"))
    })?;
    let mut allowed_references = [
        editor.body_storage_id.get(),
        title.reference_id,
        caption.reference_id,
        style_id,
    ]
    .into_iter()
    .collect::<HashSet<_>>();
    allowed_references.extend(movie.super_.comment.map(|reference| reference.identifier));
    let unexpected_references = drawable
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
            "Pages movie {drawable_object_id} has unsupported private references {unexpected_references:?}"
        )));
    }

    let registered =
        component_uuid_identifiers(editor.package(), DOCUMENT_OBJECT_ID)?.unwrap_or_default();
    let uuid_object_ids = object_ids
        .iter()
        .copied()
        .filter(|identifier| registered.contains(identifier))
        .collect::<Vec<_>>();
    let mut data_references = Vec::new();
    for identifier in &object_ids {
        let object = archive.object(*identifier).ok_or_else(|| {
            Error::InvalidFormat(format!("Pages movie object {identifier} is missing"))
        })?;
        data_references.extend(
            object
                .archive_info
                .message_infos
                .iter()
                .flat_map(|message| {
                    message
                        .data_references
                        .iter()
                        .chain(
                            message
                                .field_infos
                                .iter()
                                .flat_map(|field| &field.data_references),
                        )
                        .map(|data| (*data, *identifier))
                }),
        );
    }
    data_references.sort_unstable();
    data_references.dedup();
    for identifier in [
        info.movie_data_identifier.get(),
        info.poster_image_data_identifier.get(),
    ] {
        if !data_references.contains(&(identifier, drawable_object_id)) {
            return Err(Error::InvalidFormat(format!(
                "Pages movie {drawable_object_id} data {identifier} is missing from archive metadata"
            )));
        }
    }
    Ok(BodyMovieGraph {
        archive_name,
        attachment_id: *attachment_id,
        info,
        object_ids,
        uuid_object_ids,
        data_references,
    })
}

fn movie_info(
    package: &IWorkPackage,
    body_storage_id: u64,
    identifier: u64,
    anchor_character_index: u32,
) -> Result<PagesMovieInfo> {
    let movie: tsd::MovieArchive =
        decode_typed_package_object(package, identifier, MOVIE_MESSAGE_TYPE, "TSD.MovieArchive")?;
    if movie.audio_only == Some(true) || movie.is_live_video == Some(true) {
        return Err(Error::ParseError(format!(
            "Pages drawable {identifier} is not an ordinary file-backed movie"
        )));
    }
    if movie.super_.parent.map(|parent| parent.identifier) != Some(body_storage_id) {
        return Err(Error::InvalidFormat(format!(
            "Pages movie {identifier} is not owned by the body storage"
        )));
    }
    let movie_data_identifier = MediaAssetId::try_from(
        movie
            .movie_data
            .ok_or_else(|| {
                Error::InvalidFormat(format!("Pages movie {identifier} has no video data"))
            })?
            .identifier,
    )?;
    let poster_image_data_identifier = MediaAssetId::try_from(
        movie
            .poster_image_data
            .ok_or_else(|| {
                Error::InvalidFormat(format!("Pages movie {identifier} has no poster data"))
            })?
            .identifier,
    )?;
    let playback = media_playback_settings(&movie).map_err(|error| {
        Error::InvalidFormat(format!(
            "Pages movie {identifier} has invalid playback settings: {error}"
        ))
    })?;
    Ok(PagesMovieInfo {
        drawable_object_id: identifier,
        anchor_character_index,
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
#[allow(clippy::too_many_arguments)]
pub(super) fn movie_objects(
    ids: MovieObjectIds,
    body_storage_id: u64,
    style_id: u64,
    movie_data_identifier: u64,
    poster_data_identifier: u64,
    geometry: DrawableGeometry,
    natural_size: DrawableSize,
    duration_seconds: f32,
    left_margin: f32,
) -> Result<[ArchiveObject; 4]> {
    let position = geometry.position.ok_or_else(|| {
        Error::InvalidFormat("validated Pages movie geometry has no position".to_owned())
    })?;
    let size = geometry.size.ok_or_else(|| {
        Error::InvalidFormat("validated Pages movie geometry has no size".to_owned())
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
            parent: Some(reference(body_storage_id)),
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
    let attachment = DrawableAttachmentArchive {
        drawable: Some(reference(ids.drawable)),
        h_offset_type: Some(HorizontalAnchorBasis::BodyMargin as u32),
        h_offset: Some(position.x - left_margin),
        v_offset_type: Some(VerticalAnchorBasis::Page as u32),
        v_offset: Some(position.y),
    };
    Ok([
        pages_movie_object(
            ids.drawable,
            MOVIE_MESSAGE_TYPE,
            movie,
            &STANDARD_MESSAGE_VERSION,
            &[ids.title, ids.caption, style_id],
            &[movie_data_identifier, poster_data_identifier],
        )?,
        pages_movie_object(
            ids.title,
            STANDIN_CAPTION_MESSAGE_TYPE,
            tsd::StandinCaptionArchive::default(),
            &STANDIN_CAPTION_MESSAGE_VERSION,
            &[],
            &[],
        )?,
        pages_movie_object(
            ids.caption,
            STANDIN_CAPTION_MESSAGE_TYPE,
            tsd::StandinCaptionArchive::default(),
            &STANDIN_CAPTION_MESSAGE_VERSION,
            &[],
            &[],
        )?,
        pages_movie_object(
            ids.attachment,
            DRAWABLE_ATTACHMENT_MESSAGE_TYPE,
            attachment,
            &STANDARD_MESSAGE_VERSION,
            &[ids.drawable],
            &[],
        )?,
    ])
}

pub(super) fn set_movie_geometry(
    package: &mut IWorkPackage,
    archive_name: &str,
    movie_id: u64,
    geometry: DrawableGeometry,
) -> Result<()> {
    geometry.validate()?;
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(movie_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Pages movie object {movie_id} is missing"))
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
                "Pages movie {movie_id} must have exactly one MovieArchive payload"
            )));
        };
        let data = transform_length_delimited_field(
            object.messages[*message_index].data.as_slice(),
            MOVIE_DRAWABLE_FIELD,
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

pub(super) fn set_movie_properties(
    package: &mut IWorkPackage,
    archive_name: &str,
    movie_id: u64,
    properties: &DrawableProperties,
) -> Result<()> {
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(movie_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Pages movie object {movie_id} is missing"))
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
                "Pages movie {movie_id} must have exactly one MovieArchive payload"
            )));
        };
        let original = object.messages[*message_index].data.as_slice();
        let current = drawable_properties(&tsd::MovieArchive::decode(original)?.super_);
        let data = patch_wrapped_drawable_properties(original, &current, properties)?;
        let verified = tsd::MovieArchive::decode(data.as_slice())?;
        if drawable_properties(&verified.super_) != *properties {
            return Err(Error::InvalidFormat(
                "Pages movie properties patch failed validation".to_owned(),
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

pub(super) fn set_movie_attachment_position(
    package: &mut IWorkPackage,
    archive_name: &str,
    attachment_id: u64,
    position: DrawablePoint,
    left_margin: f32,
) -> Result<()> {
    if !position.x.is_finite() || !position.y.is_finite() || !left_margin.is_finite() {
        return Err(Error::ParseError(
            "Pages movie attachment position and left margin must be finite".to_owned(),
        ));
    }
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(attachment_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Pages movie attachment {attachment_id} is missing"))
        })?;
        let indexes = object
            .messages
            .iter()
            .enumerate()
            .filter(|(_, message)| message.type_ == DRAWABLE_ATTACHMENT_MESSAGE_TYPE)
            .map(|(index, _)| index)
            .collect::<Vec<_>>();
        let [message_index] = indexes.as_slice() else {
            return Err(Error::InvalidFormat(format!(
                "Pages movie attachment {attachment_id} must have exactly one payload"
            )));
        };
        let mut data = object.messages[*message_index].data.clone();
        data = patch_fixed32_field(
            &data,
            ATTACHMENT_HORIZONTAL_OFFSET_FIELD,
            true,
            Some((position.x - left_margin).to_bits()),
        )?;
        data = patch_fixed32_field(
            &data,
            ATTACHMENT_VERTICAL_OFFSET_FIELD,
            true,
            Some(position.y.to_bits()),
        )?;
        object.replace_message(
            *message_index,
            RawMessage {
                type_: DRAWABLE_ATTACHMENT_MESSAGE_TYPE,
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
        .ok_or_else(|| Error::InvalidFormat(format!("Pages movie {movie_id} has no {label}")))
}

fn pages_movie_object(
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

fn object_has_message_type(
    package: &IWorkPackage,
    identifier: u64,
    message_type: u32,
) -> Result<bool> {
    let archive_name = find_object_archive(package, identifier)?;
    let archive = package.archive(&archive_name)?;
    let object = archive.object(identifier).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Pages object {identifier} is missing from {archive_name}"
        ))
    })?;
    Ok(object
        .messages
        .iter()
        .any(|message| message.type_ == message_type))
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
