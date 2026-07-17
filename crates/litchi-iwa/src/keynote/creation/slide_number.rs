//! Source-built modern Keynote slide-number placeholders.

use super::*;

const NATIVE_PLACEHOLDER_SIZE: f32 = 100.0;
const NATIVE_GEOMETRY_FLAGS: u32 = 0;
const STORAGELESS_OBJECT_IDENTIFIER: u64 = 0;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum PlaceholderContext {
    Template,
    Live,
}

/// Construct the storage-less placeholder Keynote creates for modern themes.
#[allow(deprecated)]
pub(super) fn placeholder(parent_id: u64, context: PlaceholderContext) -> kn::PlaceholderArchive {
    let geometry = match context {
        PlaceholderContext::Template => tsd::GeometryArchive::default(),
        PlaceholderContext::Live => tsd::GeometryArchive {
            position: Some(tsp::Point { x: 0.0, y: 0.0 }),
            size: Some(tsp::Size {
                width: 0.0,
                height: 0.0,
            }),
            flags: Some(NATIVE_GEOMETRY_FLAGS),
            angle: Some(0.0),
        },
    };
    let pathsource = match context {
        PlaceholderContext::Template => template_path_source(),
        PlaceholderContext::Live => rectangle_path_source(&tsp::Size {
            width: NATIVE_PLACEHOLDER_SIZE,
            height: NATIVE_PLACEHOLDER_SIZE,
        }),
    };
    kn::PlaceholderArchive {
        super_: tswp::ShapeInfoArchive {
            super_: tsd::ShapeArchive {
                super_: tsd::DrawableArchive {
                    geometry: Some(geometry),
                    parent: Some(reference(parent_id)),
                    exterior_text_wrap: Some(tsd::ExteriorTextWrapArchive {
                        r#type: Some(TEXT_WRAP_TYPE),
                        direction: Some(TEXT_WRAP_DIRECTION),
                        fit_type: Some(TEXT_WRAP_FIT_TYPE),
                        margin: Some(TEXT_WRAP_MARGIN),
                        alpha_threshold: Some(TEXT_WRAP_ALPHA_THRESHOLD),
                        is_html_wrap: Some(false),
                    }),
                    locked: Some(false),
                    aspect_ratio_locked: Some(false),
                    title_hidden: Some(false),
                    caption_hidden: Some(false),
                    ..Default::default()
                },
                style: Some(reference(SHAPE_STYLE)),
                pathsource: Some(pathsource),
                stroke_pattern_offset_distance: Some(0.0),
                ..Default::default()
            },
            deprecated_storage: Some(reference(STORAGELESS_OBJECT_IDENTIFIER)),
            owned_storage: Some(reference(STORAGELESS_OBJECT_IDENTIFIER)),
            is_text_box: Some(false),
            ..Default::default()
        },
        kind: Some(kn::placeholder_archive::Kind::KKindSlideNumberPlaceholder as i32),
    }
}

fn template_path_source() -> tsd::PathSourceArchive {
    use tsp::path::{Element, ElementType};

    tsd::PathSourceArchive {
        horizontal_flip: Some(false),
        vertical_flip: Some(false),
        bezier_path_source: Some(tsd::BezierPathSourceArchive {
            natural_size: Some(tsp::Size {
                width: NATIVE_PLACEHOLDER_SIZE,
                height: NATIVE_PLACEHOLDER_SIZE,
            }),
            path: Some(tsp::Path {
                elements: vec![
                    Element {
                        r#type: ElementType::MoveTo as i32,
                        points: vec![tsp::Point { x: 0.0, y: 0.0 }],
                    },
                    Element {
                        r#type: ElementType::LineTo as i32,
                        points: vec![tsp::Point {
                            x: NATIVE_PLACEHOLDER_SIZE,
                            y: NATIVE_PLACEHOLDER_SIZE,
                        }],
                    },
                ],
            }),
            ..Default::default()
        }),
        ..Default::default()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn contexts_match_native_storageless_placeholder_shapes() {
        let template = placeholder(TEMPLATE_SLIDE, PlaceholderContext::Template);
        assert_eq!(
            template.kind,
            Some(kn::placeholder_archive::Kind::KKindSlideNumberPlaceholder as i32)
        );
        assert_eq!(
            template.super_.owned_storage,
            Some(reference(STORAGELESS_OBJECT_IDENTIFIER))
        );
        assert_eq!(
            template.super_.super_.super_.geometry,
            Some(tsd::GeometryArchive::default())
        );

        let live = placeholder(LIVE_SLIDE, PlaceholderContext::Live);
        let geometry = live.super_.super_.super_.geometry.unwrap();
        assert_eq!(geometry.position, Some(tsp::Point { x: 0.0, y: 0.0 }));
        assert_eq!(
            geometry.size,
            Some(tsp::Size {
                width: 0.0,
                height: 0.0,
            })
        );
        assert_eq!(geometry.flags, Some(NATIVE_GEOMETRY_FLAGS));
        assert_eq!(live.super_.is_text_box, Some(false));
    }
}
