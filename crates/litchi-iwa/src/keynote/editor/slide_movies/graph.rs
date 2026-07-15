//! Typed construction of source-built Keynote movie graphs.

use super::*;

const STYLESHEET_MESSAGE_TYPE: u32 = 401;
const MEDIA_STYLE_MESSAGE_TYPE: u32 = 3_016;
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

#[derive(Debug, Clone, Copy)]
pub(super) struct MovieObjectIds {
    pub(super) drawable: u64,
    title: u64,
    caption: u64,
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
        })
    }

    pub(super) const fn last(self) -> u64 {
        self.caption
    }

    pub(super) const fn all(self) -> [u64; 3] {
        [self.drawable, self.title, self.caption]
    }
}

pub(super) struct MovieCreationContext {
    pub(super) slide_id: u64,
    pub(super) component_id: u64,
    pub(super) archive_name: String,
    pub(super) style_id: u64,
    pub(super) stylesheet_component_id: u64,
}

pub(super) fn movie_creation_values(
    options: KeynoteSlideMovieOptions,
) -> Result<(DrawableGeometry, f32)> {
    if options.size.width <= 0.0 || options.size.height <= 0.0 {
        return Err(Error::ParseError(
            "Keynote movie display size must be greater than zero".to_owned(),
        ));
    }
    if !options.natural_size.width.is_finite()
        || !options.natural_size.height.is_finite()
        || options.natural_size.width <= 0.0
        || options.natural_size.height <= 0.0
    {
        return Err(Error::ParseError(
            "Keynote movie natural size must be finite and greater than zero".to_owned(),
        ));
    }
    let duration_seconds = options.duration.as_secs_f64();
    if duration_seconds == 0.0 || duration_seconds > f64::from(f32::MAX) {
        return Err(Error::ParseError(
            "Keynote movie duration must be greater than zero and fit in f32 seconds".to_owned(),
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

pub(super) fn movie_creation_context(
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
    let archive_name = graph.archive_name(slide.slide_id)?.to_owned();
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
        slide_id: slide.slide_id,
        component_id,
        archive_name,
        style_id,
        stylesheet_component_id,
    })
}

#[allow(deprecated)]
pub(super) fn movie_objects(
    ids: MovieObjectIds,
    slide_id: u64,
    style_id: u64,
    movie_data_identifier: u64,
    poster_data_identifier: u64,
    geometry: DrawableGeometry,
    natural_size: DrawableSize,
    duration_seconds: f32,
) -> Result<[ArchiveObject; 3]> {
    let position = geometry.position.ok_or_else(|| {
        Error::InvalidFormat("validated Keynote movie geometry has no position".to_owned())
    })?;
    let size = geometry.size.ok_or_else(|| {
        Error::InvalidFormat("validated Keynote movie geometry has no size".to_owned())
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
            parent: Some(reference(slide_id)),
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
        plays_across_slides: Some(true),
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
        keynote_object(
            ids.drawable,
            MOVIE_MESSAGE_TYPE,
            movie,
            &STANDARD_MESSAGE_VERSION,
            &[ids.caption, ids.title, style_id],
            &[poster_data_identifier, movie_data_identifier],
        )?,
        keynote_object(
            ids.title,
            STANDIN_CAPTION_MESSAGE_TYPE,
            tsd::StandinCaptionArchive::default(),
            &STANDIN_CAPTION_MESSAGE_VERSION,
            &[],
            &[],
        )?,
        keynote_object(
            ids.caption,
            STANDIN_CAPTION_MESSAGE_TYPE,
            tsd::StandinCaptionArchive::default(),
            &STANDIN_CAPTION_MESSAGE_VERSION,
            &[],
            &[],
        )?,
    ])
}

fn keynote_object(
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
