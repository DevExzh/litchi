//! Typed construction of source-built Keynote movie graphs.

use super::*;
use litchi_iwa_protos::keynote_media_creation_codec;
use litchi_keynote::slide::movie::Options as SlideMovieOptions;

const STYLESHEET_MESSAGE_TYPE: u32 = 401;
const MEDIA_STYLE_MESSAGE_TYPE: u32 = 3_016;
const STANDIN_CAPTION_MESSAGE_TYPE: u32 = 3_097;
const DEFAULT_DRAWABLE_FLAGS: u32 = 3;
const DEFAULT_MOVIE_ROTATION_DEGREES: f32 = 0.0;
const STANDARD_MESSAGE_VERSION: [u32; 3] = [1, 0, 5];
const STANDIN_CAPTION_MESSAGE_VERSION: [u32; 3] = [10, 1, 0];

#[derive(Debug, Clone, Copy)]
pub(in crate::keynote::editor) struct MovieObjectIds {
    pub(in crate::keynote::editor) drawable: u64,
    title: u64,
    caption: u64,
}

impl MovieObjectIds {
    pub(in crate::keynote::editor) fn allocate(first: u64) -> Result<Self> {
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

    pub(in crate::keynote::editor) const fn last(self) -> u64 {
        self.caption
    }

    pub(in crate::keynote::editor) const fn all(self) -> [u64; 3] {
        [self.drawable, self.title, self.caption]
    }
}

pub(in crate::keynote::editor) struct MovieCreationContext {
    pub(in crate::keynote::editor) slide_id: u64,
    pub(in crate::keynote::editor) component_id: u64,
    pub(in crate::keynote::editor) archive_name: String,
    pub(in crate::keynote::editor) style_id: u64,
    pub(in crate::keynote::editor) stylesheet_component_id: u64,
}

pub(in crate::keynote::editor) fn movie_creation_values(
    options: SlideMovieOptions,
) -> Result<(DrawableGeometry, f32)> {
    let geometry = DrawableGeometry {
        position: Some(options.position()),
        size: Some(options.size()),
        flags: Some(DEFAULT_DRAWABLE_FLAGS),
        angle: Some(DEFAULT_MOVIE_ROTATION_DEGREES),
    }
    .validate()?;
    Ok((geometry, options.duration_seconds()))
}

pub(in crate::keynote::editor) fn movie_creation_context(
    editor: &KeynoteEditor,
    slide_index: usize,
) -> Result<MovieCreationContext> {
    let slides = editor.slides()?;
    let slide = slides.get(slide_index).ok_or_else(|| {
        Error::ParseError(format!(
            "Keynote slide index {slide_index} is out of range for {} slides",
            slides.len()
        ))
    })?;
    let slide_id = slide.native_ids()?.slide.get();
    let graph = ObjectGraph::read(editor.package())?;
    let document: kn::DocumentArchive = graph.decode(1, "KN.DocumentArchive")?;
    let show: kn::ShowArchive = graph.decode(document.show.identifier, "KN.ShowArchive")?;
    let stylesheet_id = show.stylesheet.identifier;
    let stylesheet: tss::StylesheetArchive = graph.decode_type(
        stylesheet_id,
        STYLESHEET_MESSAGE_TYPE,
        "TSS.StylesheetArchive",
    )?;
    let style_id = stylesheet
        .styles
        .iter()
        .map(|style| style.identifier)
        .find(|identifier| {
            graph.objects.get(identifier).is_some_and(|messages| {
                messages
                    .iter()
                    .any(|message| message.type_ == MEDIA_STYLE_MESSAGE_TYPE)
            })
        })
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Keynote stylesheet {stylesheet_id} has no media style"
            ))
        })?;
    let archive_name = graph.archive_name(slide_id)?.to_owned();
    let component_id = component_identifier_for_entry(editor.package(), &archive_name)?
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Keynote slide component {archive_name} is not registered"
            ))
        })?;
    let stylesheet_archive = graph.archive_name(stylesheet_id)?;
    let stylesheet_component_id =
        component_identifier_for_entry(editor.package(), stylesheet_archive)?.ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Keynote stylesheet component {stylesheet_archive} is not registered"
            ))
        })?;
    Ok(MovieCreationContext {
        slide_id,
        component_id,
        archive_name,
        style_id,
        stylesheet_component_id,
    })
}

pub(in crate::keynote::editor) fn movie_objects(
    ids: MovieObjectIds,
    slide_id: u64,
    style_id: u64,
    movie_data_identifier: u64,
    poster_data_identifier: u64,
    geometry: DrawableGeometry,
    natural_size: DrawableSize,
    duration_seconds: f32,
) -> Result<[ArchiveObject; 3]> {
    media_objects(
        ids,
        slide_id,
        style_id,
        movie_data_identifier,
        geometry,
        duration_seconds,
        poster_data_identifier,
        natural_size,
    )
}

fn media_objects(
    ids: MovieObjectIds,
    slide_id: u64,
    style_id: u64,
    data_identifier: u64,
    geometry: DrawableGeometry,
    duration_seconds: f32,
    poster_data_identifier: u64,
    natural_size: DrawableSize,
) -> Result<[ArchiveObject; 3]> {
    let position = geometry.position.ok_or_else(|| {
        Error::InvalidFormat("validated Keynote movie geometry has no position".to_owned())
    })?;
    let size = geometry.size.ok_or_else(|| {
        Error::InvalidFormat("validated Keynote movie geometry has no size".to_owned())
    })?;
    let geometry = keynote_media_creation_codec::Geometry::new(
        keynote_media_creation_codec::Point::new(position.x, position.y),
        keynote_media_creation_codec::Size::new(size.width, size.height),
        geometry.flags,
        geometry.angle,
    );
    let write = keynote_media_creation_codec::MediaArchiveWrite::movie(
        slide_id,
        style_id,
        ids.title,
        ids.caption,
        data_identifier,
        Some(poster_data_identifier),
        geometry,
        duration_seconds,
        keynote_media_creation_codec::Size::new(natural_size.width, natural_size.height),
        true,
    );
    let data_references = vec![poster_data_identifier, data_identifier];
    let movie = keynote_media_creation_codec::encode_media_archive(
        &write,
        keynote_media_creation_codec::EncodeOptions::for_write(&write),
    )
    .map_err(|error| Error::InvalidFormat(format!("invalid Keynote media creation: {error}")))?;
    Ok([
        keynote_object(
            ids.drawable,
            MOVIE_MESSAGE_TYPE,
            movie,
            &STANDARD_MESSAGE_VERSION,
            &[ids.caption, ids.title, style_id],
            &data_references,
        )?,
        keynote_object(
            ids.title,
            STANDIN_CAPTION_MESSAGE_TYPE,
            keynote_media_creation_codec::canonical_standin_payload().to_vec(),
            &STANDIN_CAPTION_MESSAGE_VERSION,
            &[],
            &[],
        )?,
        keynote_object(
            ids.caption,
            STANDIN_CAPTION_MESSAGE_TYPE,
            keynote_media_creation_codec::canonical_standin_payload().to_vec(),
            &STANDIN_CAPTION_MESSAGE_VERSION,
            &[],
            &[],
        )?,
    ])
}

fn keynote_object(
    identifier: u64,
    message_type: u32,
    data: Vec<u8>,
    versions: &[u32],
    object_references: &[u64],
    data_references: &[u64],
) -> Result<ArchiveObject> {
    let mut object = ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_: message_type,
            data,
        }],
    )?;
    let info = &mut object.archive_info.message_infos[0];
    info.versions = versions.to_vec();
    info.object_references = object_references.to_vec();
    info.data_references = data_references.to_vec();
    Ok(object)
}
