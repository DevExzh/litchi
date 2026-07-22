//! Typed construction and strict discovery of sheet-owned Numbers audio graphs.

use super::*;
use crate::media_playback::media_playback_settings;
use crate::shapes::{DrawableSize, drawable_properties, geometry_from_drawable};

const AUDIO_MESSAGE_TYPE: u32 = 3_007;
const STANDIN_CAPTION_MESSAGE_TYPE: u32 = 3_097;
const DEFAULT_DRAWABLE_FLAGS: u32 = 3;
const DEFAULT_AUDIO_FLAGS: u32 = 0;
const DEFAULT_AUDIO_ROTATION_DEGREES: f32 = 0.0;
const DEFAULT_AUDIO_VOLUME: f32 = 1.0;
const DEFAULT_TEXT_WRAP_MARGIN_POINTS: f32 = 12.0;
const DEFAULT_TEXT_WRAP_ALPHA_THRESHOLD: f32 = 0.5;
const STANDARD_MESSAGE_VERSION: [u32; 3] = [1, 0, 5];
const STANDIN_CAPTION_MESSAGE_VERSION: [u32; 3] = [10, 1, 0];
const ZERO_SIZE: DrawableSize = DrawableSize {
    width: 0.0,
    height: 0.0,
};

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

pub(super) struct SheetAudioGraph {
    pub(super) sheet_id: u64,
    pub(super) archive_name: String,
    pub(super) component_id: u64,
    pub(super) info: NumbersSheetAudioInfo,
    pub(super) geometry: DrawableGeometry,
    pub(super) object_ids: Vec<u64>,
    pub(super) uuid_object_ids: Vec<u64>,
    pub(super) data_references: Vec<(u64, u64)>,
}

pub(super) fn audio_creation_values(
    options: NumbersSheetAudioOptions,
) -> Result<(DrawableGeometry, f32)> {
    let duration_seconds = options.duration.as_secs_f64();
    if duration_seconds == 0.0 || duration_seconds > f64::from(f32::MAX) {
        return Err(Error::ParseError(
            "Numbers audio duration must be greater than zero and fit in f32 seconds".to_owned(),
        ));
    }
    let geometry = DrawableGeometry {
        position: Some(options.position),
        size: Some(ZERO_SIZE),
        flags: Some(DEFAULT_DRAWABLE_FLAGS),
        angle: Some(DEFAULT_AUDIO_ROTATION_DEGREES),
    }
    .validate()?;
    Ok((geometry, duration_seconds as f32))
}

pub(super) fn audio_infos(
    editor: &NumbersEditor,
    sheet_id: u64,
) -> Result<Vec<NumbersSheetAudioInfo>> {
    let (_, _, sheet) = numbers_sheet(editor.package(), sheet_id)?;
    let locations = object_locations(editor.package())?;
    let mut audio = Vec::new();
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
        let messages = object
            .messages
            .iter()
            .filter(|message| message.type_ == AUDIO_MESSAGE_TYPE)
            .collect::<Vec<_>>();
        if messages.is_empty() {
            continue;
        }
        let [message] = messages.as_slice() else {
            return Err(Error::InvalidFormat(format!(
                "Numbers drawable {} has multiple media payloads",
                reference.identifier
            )));
        };
        let media = tsd::MovieArchive::decode(message.data.as_slice())?;
        if media.audio_only != Some(true) || media.is_live_video == Some(true) {
            continue;
        }
        audio.push(audio_info(
            editor.package(),
            sheet_id,
            reference.identifier,
        )?);
    }
    let mut drawable_ids = HashSet::with_capacity(audio.len());
    if let Some(duplicate) = audio
        .iter()
        .map(|item| item.drawable_object_id)
        .find(|identifier| !drawable_ids.insert(*identifier))
    {
        return Err(Error::InvalidFormat(format!(
            "Numbers audio {duplicate} occurs multiple times in its sheet"
        )));
    }
    Ok(audio)
}

pub(super) fn audio_graph(
    editor: &NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
) -> Result<SheetAudioGraph> {
    let (archive_name, _, sheet) = numbers_sheet(editor.package(), sheet_id)?;
    if sheet
        .drawable_infos
        .iter()
        .filter(|reference| reference.identifier == drawable_object_id)
        .count()
        != 1
    {
        return Err(Error::ParseError(format!(
            "Numbers sheet {sheet_id} does not own audio {drawable_object_id} exactly once"
        )));
    }
    let locations = object_locations(editor.package())?;
    if locations.get(&drawable_object_id).map(String::as_str) != Some(archive_name.as_str()) {
        return Err(Error::InvalidFormat(format!(
            "Numbers audio {drawable_object_id} is outside sheet component {archive_name}"
        )));
    }
    let archive = editor.package().archive(&archive_name)?;
    let object = archive.object(drawable_object_id).ok_or_else(|| {
        Error::InvalidFormat(format!("Numbers audio {drawable_object_id} is missing"))
    })?;
    let messages = object
        .messages
        .iter()
        .filter(|message| message.type_ == AUDIO_MESSAGE_TYPE)
        .collect::<Vec<_>>();
    let [message] = messages.as_slice() else {
        return Err(Error::ParseError(format!(
            "Numbers drawable {drawable_object_id} is not ordinary audio"
        )));
    };
    let audio = tsd::MovieArchive::decode(message.data.as_slice())?;
    if audio.audio_only != Some(true) || audio.is_live_video == Some(true) {
        return Err(Error::ParseError(format!(
            "Numbers drawable {drawable_object_id} is not ordinary file-backed audio"
        )));
    }
    if audio.super_.parent.map(|parent| parent.identifier) != Some(sheet_id) {
        return Err(Error::InvalidFormat(format!(
            "Numbers audio {drawable_object_id} is not owned by sheet {sheet_id}"
        )));
    }
    let geometry = geometry_from_drawable(&audio.super_)?;
    if geometry.size != Some(ZERO_SIZE) {
        return Err(Error::InvalidFormat(format!(
            "Numbers audio {drawable_object_id} must use zero-size control geometry"
        )));
    }
    let title_id = required_reference(drawable_object_id, audio.super_.title, "title stand-in")?;
    let caption_id =
        required_reference(drawable_object_id, audio.super_.caption, "caption stand-in")?;
    let style_id = required_reference(drawable_object_id, audio.style, "media style")?;
    if !locations.contains_key(&style_id) {
        return Err(Error::InvalidFormat(format!(
            "Numbers audio {drawable_object_id} style {style_id} is missing"
        )));
    }
    let object_ids = vec![drawable_object_id, title_id, caption_id];
    if object_ids.iter().copied().collect::<HashSet<_>>().len() != object_ids.len() {
        return Err(Error::InvalidFormat(format!(
            "Numbers audio {drawable_object_id} reuses private graph identifiers"
        )));
    }
    for identifier in [title_id, caption_id] {
        if locations.get(&identifier).map(String::as_str) != Some(archive_name.as_str()) {
            return Err(Error::InvalidFormat(format!(
                "Numbers audio {drawable_object_id} private graph spans multiple archives"
            )));
        }
        let standin = archive.object(identifier).ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers audio stand-in {identifier} is missing"))
        })?;
        if standin
            .messages
            .iter()
            .filter(|message| message.type_ == STANDIN_CAPTION_MESSAGE_TYPE)
            .count()
            != 1
        {
            return Err(Error::InvalidFormat(format!(
                "Numbers audio stand-in {identifier} is malformed"
            )));
        }
    }

    let mut allowed_references = [sheet_id, title_id, caption_id, style_id]
        .into_iter()
        .collect::<HashSet<_>>();
    allowed_references.extend(audio.super_.comment.map(|reference| reference.identifier));
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
            "Numbers audio {drawable_object_id} has unsupported private references {unexpected_references:?}"
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
    let info = audio_info(editor.package(), sheet_id, drawable_object_id)?;
    if data_references != [(info.audio_data_identifier, drawable_object_id)] {
        return Err(Error::InvalidFormat(format!(
            "Numbers audio {drawable_object_id} has inconsistent data-reference metadata {data_references:?}"
        )));
    }
    Ok(SheetAudioGraph {
        sheet_id,
        archive_name,
        component_id,
        info,
        geometry,
        object_ids,
        uuid_object_ids,
        data_references,
    })
}

fn audio_info(
    package: &IWorkPackage,
    sheet_id: u64,
    identifier: u64,
) -> Result<NumbersSheetAudioInfo> {
    let locations = object_locations(package)?;
    let archive_name = locations
        .get(&identifier)
        .ok_or_else(|| Error::InvalidFormat(format!("Numbers audio {identifier} is missing")))?;
    let archive = package.archive(archive_name)?;
    let object = archive
        .object(identifier)
        .ok_or_else(|| Error::InvalidFormat(format!("Numbers audio {identifier} is missing")))?;
    let messages = object
        .messages
        .iter()
        .filter(|message| message.type_ == AUDIO_MESSAGE_TYPE)
        .collect::<Vec<_>>();
    let [message] = messages.as_slice() else {
        return Err(Error::InvalidFormat(format!(
            "Numbers audio {identifier} must have exactly one media payload"
        )));
    };
    let audio = tsd::MovieArchive::decode(message.data.as_slice())?;
    if audio.audio_only != Some(true) || audio.is_live_video == Some(true) {
        return Err(Error::ParseError(format!(
            "Numbers drawable {identifier} is not ordinary file-backed audio"
        )));
    }
    if audio.super_.parent.map(|parent| parent.identifier) != Some(sheet_id) {
        return Err(Error::InvalidFormat(format!(
            "Numbers audio {identifier} is not owned by sheet {sheet_id}"
        )));
    }
    let audio_data_identifier = audio
        .movie_data
        .ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers audio {identifier} has no data reference"))
        })?
        .identifier;
    if audio.poster_image_data.is_some() {
        return Err(Error::InvalidFormat(format!(
            "Numbers audio {identifier} unexpectedly references a poster image"
        )));
    }
    let geometry = geometry_from_drawable(&audio.super_)?;
    if geometry.size != Some(ZERO_SIZE) {
        return Err(Error::InvalidFormat(format!(
            "Numbers audio {identifier} must use zero-size control geometry"
        )));
    }
    let position = geometry.position.ok_or_else(|| {
        Error::InvalidFormat(format!("Numbers audio {identifier} has no position"))
    })?;
    let playback = media_playback_settings(&audio).map_err(|error| {
        Error::InvalidFormat(format!(
            "Numbers audio {identifier} has invalid playback settings: {error}"
        ))
    })?;
    Ok(NumbersSheetAudioInfo {
        sheet_id,
        drawable_object_id: identifier,
        audio_data_identifier,
        position,
        properties: drawable_properties(&audio.super_),
        playback,
        duration: playback.duration(),
    })
}

#[allow(deprecated)]
pub(super) fn audio_objects(
    ids: MovieObjectIds,
    sheet_id: u64,
    style_id: u64,
    audio_data_identifier: u64,
    geometry: DrawableGeometry,
    duration_seconds: f32,
) -> Result<[ArchiveObject; 3]> {
    let position = geometry.position.ok_or_else(|| {
        Error::InvalidFormat("validated Numbers audio geometry has no position".to_owned())
    })?;
    let size = geometry.size.ok_or_else(|| {
        Error::InvalidFormat("validated Numbers audio geometry has no size".to_owned())
    })?;
    let audio = tsd::MovieArchive {
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
            identifier: audio_data_identifier,
        }),
        start_time: Some(0.0),
        end_time: Some(duration_seconds),
        poster_time: Some(0.0),
        loop_option: Some(tsd::movie_archive::MovieLoopOption::None as i32),
        volume: Some(DEFAULT_AUDIO_VOLUME),
        audio_only: Some(true),
        streaming: Some(false),
        poster_image_generated_with_alpha_support: Some(false),
        flags: Some(DEFAULT_AUDIO_FLAGS),
        style: Some(reference(style_id)),
        original_size: Some(tsp::Size {
            width: ZERO_SIZE.width,
            height: ZERO_SIZE.height,
        }),
        natural_size: Some(tsp::Size {
            width: ZERO_SIZE.width,
            height: ZERO_SIZE.height,
        }),
        ..Default::default()
    };
    Ok([
        numbers_audio_object(
            ids.drawable,
            AUDIO_MESSAGE_TYPE,
            audio,
            &STANDARD_MESSAGE_VERSION,
            &[ids.title, ids.caption, style_id],
            &[audio_data_identifier],
        )?,
        numbers_audio_object(
            ids.title,
            STANDIN_CAPTION_MESSAGE_TYPE,
            tsd::StandinCaptionArchive::default(),
            &STANDIN_CAPTION_MESSAGE_VERSION,
            &[],
            &[],
        )?,
        numbers_audio_object(
            ids.caption,
            STANDIN_CAPTION_MESSAGE_TYPE,
            tsd::StandinCaptionArchive::default(),
            &STANDIN_CAPTION_MESSAGE_VERSION,
            &[],
            &[],
        )?,
    ])
}

fn required_reference(
    audio_id: u64,
    reference: Option<tsp::Reference>,
    label: &str,
) -> Result<u64> {
    reference
        .map(|reference| reference.identifier)
        .ok_or_else(|| Error::InvalidFormat(format!("Numbers audio {audio_id} has no {label}")))
}

fn numbers_audio_object(
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

fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..Default::default()
    }
}
