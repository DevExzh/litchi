//! Typed native archive construction shared by image title and caption editors.

use prost::Message;

use crate::archive::{ArchiveObject, RawMessage};
use crate::protobuf::{tsa, tsd, tsp, tss, tswp};
use crate::wire::patch_length_delimited_field;
use crate::{Error, Result};

pub(crate) const CAPTION_INFO_MESSAGE_TYPE: u32 = 633;
pub(crate) const CAPTION_PLACEMENT_MESSAGE_TYPE: u32 = 634;
pub(crate) const SHAPE_STYLE_MESSAGE_TYPE: u32 = 2_025;
pub(crate) const STORAGE_MESSAGE_TYPE: u32 = 2_001;
pub(crate) const STANDIN_CAPTION_MESSAGE_TYPE: u32 = 3_097;

const STANDARD_MESSAGE_VERSION: [u32; 3] = [1, 0, 5];
const STANDIN_CAPTION_MESSAGE_VERSION: [u32; 3] = [10, 1, 0];
const DRAWABLE_TITLE_FIELD: u32 = 10;
const DRAWABLE_CAPTION_FIELD: u32 = 11;
const DEFAULT_CAPTION_GEOMETRY_FLAGS: u32 = 1;
const DEFAULT_CAPTION_ROTATION_DEGREES: f32 = 0.0;
const DEFAULT_TEXT_WRAP_MARGIN_POINTS: f32 = 12.0;
const DEFAULT_TEXT_WRAP_ALPHA_THRESHOLD: f32 = 0.5;
const DEFAULT_CAPTION_STYLE_OVERRIDE_COUNT: u32 = 7;
const DEFAULT_CAPTION_STYLE_PADDING_POINTS: f32 = 4.0;
const DEFAULT_STROKE_WIDTH_POINTS: f32 = 1.0;
const DEFAULT_STROKE_MITER_LIMIT: f32 = 4.0;
const DEFAULT_SHADOW_ANGLE_DEGREES: f32 = 315.0;
const DEFAULT_SHADOW_OFFSET_POINTS: f32 = 5.0;
const DEFAULT_SHADOW_RADIUS_POINTS: i32 = 1;
const DEFAULT_SHADOW_OPACITY: f32 = 1.0;
const NORMALIZED_PATH_EXTENT: f32 = 100.0;
const DEFAULT_STROKE_PATTERN_ELEMENT_COUNT: usize = 6;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum DrawableCaptionKind {
    Caption,
    Title,
}

impl DrawableCaptionKind {
    pub(crate) const fn drawable_field(self) -> u32 {
        match self {
            Self::Caption => DRAWABLE_CAPTION_FIELD,
            Self::Title => DRAWABLE_TITLE_FIELD,
        }
    }

    pub(crate) const fn native_kind(self) -> i32 {
        match self {
            Self::Caption => tsd::CaptionOrTitleKind::Caption as i32,
            Self::Title => tsd::CaptionOrTitleKind::Title as i32,
        }
    }

    const fn style_identifier(self) -> &'static str {
        match self {
            Self::Caption => "captions-0-shapestyle-Object Caption",
            Self::Title => "captions-0-shapestyle-Object Title",
        }
    }

    const fn placement(self) -> tsa::CaptionPlacementArchive {
        match self {
            Self::Caption => tsa::CaptionPlacementArchive {
                caption_anchor_location: Some(CaptionAnchorLocation::Bottom as i32),
                drawable_anchor_location: Some(CaptionAnchorLocation::Top as i32),
            },
            Self::Title => tsa::CaptionPlacementArchive {
                caption_anchor_location: Some(CaptionAnchorLocation::Top as i32),
                drawable_anchor_location: Some(CaptionAnchorLocation::Bottom as i32),
            },
        }
    }
}

/// Existing theme objects needed to materialize a native image title or caption.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct CaptionThemeStyle {
    pub(crate) stylesheet_id: u64,
    pub(crate) paragraph_style_id: u64,
}

/// Fresh identifiers for one native title or caption graph.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct CaptionObjectIds {
    pub(crate) style: u64,
    pub(crate) info: u64,
    pub(crate) storage: u64,
    pub(crate) placement: u64,
}

impl CaptionObjectIds {
    pub(crate) fn allocate(first: u64) -> Result<Self> {
        let identifier = |offset: u64| {
            first
                .checked_add(offset)
                .ok_or_else(|| Error::ParseError("iWork object identifier overflow".to_owned()))
        };
        Ok(Self {
            style: identifier(0)?,
            info: identifier(1)?,
            storage: identifier(2)?,
            placement: identifier(3)?,
        })
    }

    pub(crate) const fn last(self) -> u64 {
        self.placement
    }

    pub(crate) const fn all(self) -> [u64; 4] {
        [self.style, self.info, self.storage, self.placement]
    }
}

/// Create all four native objects that back one image title or caption.
#[allow(deprecated)]
pub(crate) fn caption_objects(
    ids: CaptionObjectIds,
    image_id: u64,
    image_width: f32,
    text: &str,
    kind: DrawableCaptionKind,
    theme: CaptionThemeStyle,
    language: Option<&str>,
) -> Result<[ArchiveObject; 4]> {
    if !image_width.is_finite() || image_width <= 0.0 {
        return Err(Error::InvalidFormat(
            "native image title/caption width must be finite and greater than zero".to_owned(),
        ));
    }
    let style = caption_shape_style(theme, kind);
    let storage = caption_storage(text, theme, language);
    let info = tsa::CaptionInfoArchive {
        super_: tswp::ShapeInfoArchive {
            super_: tsd::ShapeArchive {
                super_: tsd::DrawableArchive {
                    geometry: Some(tsd::GeometryArchive {
                        position: Some(tsp::Point { x: 0.0, y: 0.0 }),
                        size: Some(tsp::Size {
                            width: image_width,
                            height: 0.0,
                        }),
                        flags: Some(DEFAULT_CAPTION_GEOMETRY_FLAGS),
                        angle: Some(DEFAULT_CAPTION_ROTATION_DEGREES),
                    }),
                    parent: Some(reference(image_id)),
                    exterior_text_wrap: Some(tsd::ExteriorTextWrapArchive {
                        r#type: Some(TextWrapType::Square as u32),
                        direction: Some(TextWrapDirection::BothSides as u32),
                        fit_type: Some(TextWrapFit::Text as u32),
                        margin: Some(DEFAULT_TEXT_WRAP_MARGIN_POINTS),
                        alpha_threshold: Some(DEFAULT_TEXT_WRAP_ALPHA_THRESHOLD),
                        is_html_wrap: Some(false),
                    }),
                    locked: Some(false),
                    aspect_ratio_locked: Some(false),
                    title_hidden: Some(false),
                    caption_hidden: Some(false),
                    ..Default::default()
                },
                style: Some(reference(ids.style)),
                pathsource: Some(caption_path_source(image_width)),
                stroke_pattern_offset_distance: Some(0.0),
                ..Default::default()
            },
            deprecated_storage: Some(reference(ids.storage)),
            owned_storage: Some(reference(ids.storage)),
            is_text_box: Some(true),
            ..Default::default()
        },
        placement: Some(reference(ids.placement)),
        child_info_kind: Some(kind.native_kind()),
    };
    Ok([
        archive_object(
            ids.style,
            SHAPE_STYLE_MESSAGE_TYPE,
            style,
            &[theme.paragraph_style_id],
        )?,
        archive_object(
            ids.info,
            CAPTION_INFO_MESSAGE_TYPE,
            info,
            &[ids.style, ids.storage, ids.placement],
        )?,
        archive_object(
            ids.storage,
            STORAGE_MESSAGE_TYPE,
            storage,
            &[theme.paragraph_style_id],
        )?,
        archive_object(
            ids.placement,
            CAPTION_PLACEMENT_MESSAGE_TYPE,
            kind.placement(),
            &[],
        )?,
    ])
}

/// Create a fresh empty stand-in when a native title or caption is removed.
pub(crate) fn standin_caption_object(identifier: u64) -> Result<ArchiveObject> {
    let mut object = ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_: STANDIN_CAPTION_MESSAGE_TYPE,
            data: tsd::StandinCaptionArchive::default().encode_to_vec(),
        }],
    )?;
    object.archive_info.message_infos[0].versions = STANDIN_CAPTION_MESSAGE_VERSION.to_vec();
    Ok(object)
}

/// Rewrite just one reference inside a drawable while retaining all unknown fields.
pub(crate) fn patch_drawable_caption_reference(
    data: &[u8],
    kind: DrawableCaptionKind,
    replacement: u64,
) -> Result<Vec<u8>> {
    let drawable = tsd::DrawableArchive::decode(data)?;
    let was_present = match kind {
        DrawableCaptionKind::Caption => drawable.caption.is_some(),
        DrawableCaptionKind::Title => drawable.title.is_some(),
    };
    let replacement = reference(replacement).encode_to_vec();
    let data =
        patch_length_delimited_field(data, kind.drawable_field(), was_present, Some(&replacement))?;
    let verified = tsd::DrawableArchive::decode(data.as_slice())?;
    let actual = match kind {
        DrawableCaptionKind::Caption => verified.caption,
        DrawableCaptionKind::Title => verified.title,
    }
    .map(|reference| reference.identifier);
    if actual != Some(tsp::Reference::decode(replacement.as_slice())?.identifier) {
        return Err(Error::InvalidFormat(
            "native image title/caption reference patch failed validation".to_owned(),
        ));
    }
    Ok(data)
}

/// Replace one archive metadata reference without perturbing unrelated ordering.
pub(crate) fn replace_object_reference(references: &mut Vec<u64>, old: u64, new: u64) {
    if let Some(index) = references.iter().position(|identifier| *identifier == old) {
        references[index] = new;
        let mut seen_new = false;
        references.retain(|identifier| {
            if *identifier != new {
                return true;
            }
            let keep = !seen_new;
            seen_new = true;
            keep
        });
    } else if !references.contains(&new) {
        references.push(new);
    }
}

fn caption_shape_style(
    theme: CaptionThemeStyle,
    kind: DrawableCaptionKind,
) -> tswp::ShapeStyleArchive {
    tswp::ShapeStyleArchive {
        super_: tsd::ShapeStyleArchive {
            super_: tss::StyleArchive {
                style_identifier: Some(kind.style_identifier().to_owned()),
                stylesheet: Some(reference(theme.stylesheet_id)),
                ..Default::default()
            },
            override_count: Some(DEFAULT_CAPTION_STYLE_OVERRIDE_COUNT),
            shape_properties: Some(tsd::ShapeStylePropertiesArchive {
                fill: Some(transparent_fill()),
                stroke: Some(empty_stroke()),
                opacity: Some(1.0),
                shadow: Some(disabled_shadow()),
                reflection: Some(tsd::ReflectionArchive::default()),
                ..Default::default()
            }),
        },
        override_count: Some(DEFAULT_CAPTION_STYLE_OVERRIDE_COUNT),
        shape_properties: Some(tswp::ShapeStylePropertiesArchive {
            padding: Some(tswp::PaddingArchive {
                left: Some(DEFAULT_CAPTION_STYLE_PADDING_POINTS),
                top: Some(DEFAULT_CAPTION_STYLE_PADDING_POINTS),
                right: Some(DEFAULT_CAPTION_STYLE_PADDING_POINTS),
                bottom: Some(DEFAULT_CAPTION_STYLE_PADDING_POINTS),
            }),
            paragraph_style: Some(reference(theme.paragraph_style_id)),
            ..Default::default()
        }),
    }
}

fn caption_storage(
    text: &str,
    theme: CaptionThemeStyle,
    language: Option<&str>,
) -> tswp::StorageArchive {
    tswp::StorageArchive {
        style_sheet: Some(reference(theme.stylesheet_id)),
        text: vec![text.to_owned()],
        in_document: Some(true),
        table_para_style: Some(object_attribute_table(Some(reference(
            theme.paragraph_style_id,
        )))),
        table_para_data: Some(zero_para_data()),
        table_para_starts: Some(zero_para_data()),
        table_language: language.map(|language| tswp::StringAttributeTable {
            entries: vec![tswp::string_attribute_table::StringAttribute {
                character_index: 0,
                object: Some(language.to_owned()),
            }],
        }),
        table_para_bidi: Some(zero_para_data()),
        table_drop_cap_style: Some(object_attribute_table(None)),
        ..Default::default()
    }
}

fn caption_path_source(width: f32) -> tsd::PathSourceArchive {
    use tsp::path::ElementType;

    let point = |x, y| tsp::Point { x, y };
    let element = |r#type, points| tsp::path::Element {
        r#type: r#type as i32,
        points,
    };
    tsd::PathSourceArchive {
        horizontal_flip: Some(false),
        vertical_flip: Some(false),
        bezier_path_source: Some(tsd::BezierPathSourceArchive {
            natural_size: Some(tsp::Size { width, height: 0.0 }),
            path: Some(tsp::Path {
                elements: vec![
                    element(ElementType::MoveTo, vec![point(0.0, 0.0)]),
                    element(
                        ElementType::LineTo,
                        vec![point(NORMALIZED_PATH_EXTENT, 0.0)],
                    ),
                    element(
                        ElementType::LineTo,
                        vec![point(NORMALIZED_PATH_EXTENT, NORMALIZED_PATH_EXTENT)],
                    ),
                    element(
                        ElementType::LineTo,
                        vec![point(0.0, NORMALIZED_PATH_EXTENT)],
                    ),
                    element(ElementType::CloseSubpath, Vec::new()),
                    element(ElementType::MoveTo, vec![point(0.0, 0.0)]),
                ],
            }),
            ..Default::default()
        }),
        ..Default::default()
    }
}

fn transparent_fill() -> tsd::FillArchive {
    tsd::FillArchive {
        color: Some(color(0.0)),
        ..Default::default()
    }
}

fn empty_stroke() -> tsd::StrokeArchive {
    tsd::StrokeArchive {
        color: Some(color(1.0)),
        width: Some(DEFAULT_STROKE_WIDTH_POINTS),
        cap: Some(tsd::stroke_archive::LineCap::ButtCap as i32),
        join: Some(tsd::LineJoin::MiterJoin as i32),
        miter_limit: Some(DEFAULT_STROKE_MITER_LIMIT),
        pattern: Some(tsd::StrokePatternArchive {
            r#type: Some(tsd::stroke_pattern_archive::StrokePatternType::TsdEmptyPattern as i32),
            phase: Some(0.0),
            count: Some(0),
            pattern: vec![0.0; DEFAULT_STROKE_PATTERN_ELEMENT_COUNT],
        }),
        ..Default::default()
    }
}

fn disabled_shadow() -> tsd::ShadowArchive {
    tsd::ShadowArchive {
        color: Some(color(1.0)),
        angle: Some(DEFAULT_SHADOW_ANGLE_DEGREES),
        offset: Some(DEFAULT_SHADOW_OFFSET_POINTS),
        radius: Some(DEFAULT_SHADOW_RADIUS_POINTS),
        opacity: Some(DEFAULT_SHADOW_OPACITY),
        is_enabled: Some(false),
        r#type: Some(tsd::shadow_archive::ShadowType::TsdDropShadow as i32),
        ..Default::default()
    }
}

fn color(alpha: f32) -> tsp::Color {
    tsp::Color {
        model: tsp::color::ColorModel::Rgb as i32,
        r: Some(0.0),
        g: Some(0.0),
        b: Some(0.0),
        rgbspace: Some(tsp::color::RgbColorSpace::Srgb as i32),
        a: Some(alpha),
        ..Default::default()
    }
}

fn object_attribute_table(object: Option<tsp::Reference>) -> tswp::ObjectAttributeTable {
    tswp::ObjectAttributeTable {
        entries: vec![tswp::object_attribute_table::ObjectAttribute {
            character_index: 0,
            object,
        }],
    }
}

fn zero_para_data() -> tswp::ParaDataAttributeTable {
    tswp::ParaDataAttributeTable {
        entries: vec![tswp::para_data_attribute_table::ParaDataAttribute {
            character_index: 0,
            first: 0,
            second: 0,
        }],
    }
}

fn archive_object(
    identifier: u64,
    message_type: u32,
    message: impl Message,
    object_references: &[u64],
) -> Result<ArchiveObject> {
    let mut object = ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_: message_type,
            data: message.encode_to_vec(),
        }],
    )?;
    let info = &mut object.archive_info.message_infos[0];
    info.versions = STANDARD_MESSAGE_VERSION.to_vec();
    info.object_references = object_references.to_vec();
    Ok(object)
}

fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..Default::default()
    }
}

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

/// Native `CaptionPlacementArchive` anchors used by title and caption children.
#[derive(Debug, Clone, Copy)]
#[repr(i32)]
enum CaptionAnchorLocation {
    Top = 1,
    Bottom = 7,
}
