//! SVG style and XML serialization helpers.

use super::super::constants::{brush, pen};
use super::state::{Brush, MappingState, Pen};
use super::transform::CoordinateTransform;
use crate::svg_utils::{write_color_hex, write_num};

/// Fill attributes for an SVG paint value under the polygon fill mode.
pub(super) fn fill_attrs(paint: &str, poly_fill_mode: u16) -> String {
    let mut attrs = String::with_capacity(48 + paint.len());
    attrs.push_str(r#" fill=""#);
    attrs.push_str(paint);
    attrs.push('"');
    attrs.push_str(if poly_fill_mode == 2 {
        r#" fill-rule="nonzero""#
    } else {
        r#" fill-rule="evenodd""#
    });
    attrs
}

/// The SVG paint value of a `COLORREF`.
pub(super) fn color_paint(color: u32) -> String {
    let mut paint = String::with_capacity(7);
    write_color_hex(&mut paint, color);
    paint
}

pub(super) fn stroke_attrs(
    pen_value: &Pen,
    mapping: &MappingState,
    transform: &CoordinateTransform,
) -> String {
    let base_style = pen_value.style & pen::PS_STYLE_MASK;
    if base_style == pen::PS_NULL {
        return r#" stroke="none""#.to_owned();
    }

    // [MS-WMF] 2.2.1.4: the x-coordinate of the Width point is the pen
    // width and the y-coordinate is ignored. Programs leave arbitrary values
    // there, so reading it would draw strokes far wider than Windows does.
    // GDI draws every pen, and a zero-width cosmetic pen in particular, at
    // least one device pixel wide. The SVG canvas stands for the playback
    // device, so one user unit is one device pixel: the metafile's own
    // viewport is often a token 1-by-1 extent that cannot define a pixel.
    let (width_x, _) = transform.logical_vector(mapping, f64::from(pen_value.width.0).abs(), 0.0);
    let width = width_x.max(1.0);

    let mut attrs = String::with_capacity(128);
    attrs.push_str(r#" stroke=""#);
    write_color_hex(&mut attrs, pen_value.color);
    attrs.push_str(r#"" stroke-width=""#);
    write_num(&mut attrs, width);
    attrs.push('"');

    match pen_value.style & pen::PS_ENDCAP_MASK {
        pen::PS_ENDCAP_ROUND => attrs.push_str(r#" stroke-linecap="round""#),
        pen::PS_ENDCAP_SQUARE => attrs.push_str(r#" stroke-linecap="square""#),
        _ => {}, // SVG's default butt cap is PS_ENDCAP_FLAT.
    }
    match pen_value.style & pen::PS_JOIN_MASK {
        pen::PS_JOIN_ROUND => attrs.push_str(r#" stroke-linejoin="round""#),
        pen::PS_JOIN_BEVEL => attrs.push_str(r#" stroke-linejoin="bevel""#),
        _ => {}, // SVG's default miter join is PS_JOIN_MITER.
    }

    let pattern: &[f64] = match base_style {
        pen::PS_DASH => &[3.0, 1.0],
        pen::PS_DOT | pen::PS_ALTERNATE => &[1.0, 1.0],
        pen::PS_DASHDOT => &[3.0, 1.0, 1.0, 1.0],
        pen::PS_DASHDOTDOT => &[3.0, 1.0, 1.0, 1.0, 1.0, 1.0],
        _ => &[],
    };
    if !pattern.is_empty() {
        attrs.push_str(r#" stroke-dasharray=""#);
        for (index, multiplier) in pattern.iter().enumerate() {
            if index != 0 {
                attrs.push(',');
            }
            write_num(&mut attrs, width * multiplier);
        }
        attrs.push('"');
    }
    attrs
}

/// Hatch pattern definition; `background` is `None` for a transparent
/// background.
pub(super) fn hatch_definition(id: &str, brush_value: &Brush, background: Option<u32>) -> String {
    let mut definition = String::with_capacity(320);
    definition.push_str(r#"<pattern id=""#);
    escape_xml_attr_into(&mut definition, id);
    definition.push_str(r#"" patternUnits="userSpaceOnUse" width="8" height="8">"#);
    if let Some(background_color) = background {
        definition.push_str(r#"<rect width="8" height="8" fill=""#);
        write_color_hex(&mut definition, background_color);
        definition.push_str(r#""/>"#);
    }
    definition.push_str(r#"<path d=""#);
    definition.push_str(match brush_value.hatch {
        brush::HS_HORIZONTAL => "M0 4H8",
        brush::HS_VERTICAL => "M4 0V8",
        brush::HS_FDIAGONAL => "M-2 8L8-2M2 10L10 2",
        brush::HS_BDIAGONAL => "M-2 0L8 10M2-2L10 6",
        brush::HS_CROSS => "M0 4H8M4 0V8",
        brush::HS_DIAGCROSS => "M-2 0L8 10M2-2L10 6M-2 8L8-2M2 10L10 2",
        _ => "M0 4H8",
    });
    definition.push_str(r#"" fill="none" stroke=""#);
    write_color_hex(&mut definition, brush_value.color);
    definition.push_str(r#"" stroke-width="1"/></pattern>"#);
    definition
}

pub(super) fn escape_xml_text_into(output: &mut String, value: &str) {
    for character in value.chars() {
        match character {
            '&' => output.push_str("&amp;"),
            '<' => output.push_str("&lt;"),
            '>' => output.push_str("&gt;"),
            '\t' | '\n' | '\r' => output.push(character),
            character
                if character >= '\u{20}' && character != '\u{fffe}' && character != '\u{ffff}' =>
            {
                output.push(character)
            },
            _ => output.push('\u{fffd}'),
        }
    }
}

pub(super) fn escape_xml_attr_into(output: &mut String, value: &str) {
    for character in value.chars() {
        match character {
            '"' => output.push_str("&quot;"),
            '\'' => output.push_str("&apos;"),
            _ => escape_xml_text_into(output, &character.to_string()),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn xml_helpers_remove_forbidden_controls_and_escape_attributes() {
        let mut output = String::new();
        escape_xml_attr_into(&mut output, "A&\"<\u{1}");
        assert_eq!(output, "A&amp;&quot;&lt;�");
    }

    #[test]
    fn pen_cap_join_and_dash_flags_are_serialized() {
        let pen = Pen {
            style: pen::PS_DASH | pen::PS_ENDCAP_SQUARE | pen::PS_JOIN_BEVEL,
            width: (2, 0),
            color: 0xff,
        };
        let attrs = stroke_attrs(
            &pen,
            &MappingState::default(),
            &CoordinateTransform::new((0.0, 0.0, 100.0, 100.0), 100.0, 100.0),
        );
        assert!(attrs.contains(r#"stroke-linecap="square""#));
        assert!(attrs.contains(r#"stroke-linejoin="bevel""#));
        assert!(attrs.contains(r#"stroke-dasharray="6,2""#));
    }

    #[test]
    fn pen_is_at_least_one_canvas_unit_wide() {
        // One logical unit is a hundredth of a canvas unit here.
        let transform = CoordinateTransform::new((0.0, 0.0, 10_000.0, 10_000.0), 100.0, 100.0);
        for width in [(0, 0), (20, 0)] {
            let pen = Pen {
                style: pen::PS_SOLID,
                width,
                color: 0,
            };
            assert!(
                stroke_attrs(&pen, &MappingState::default(), &transform)
                    .contains(r#"stroke-width="1""#)
            );
        }
    }

    #[test]
    fn pen_width_ignores_the_y_coordinate() {
        let transform = CoordinateTransform::new((0.0, 0.0, 100.0, 100.0), 100.0, 100.0);
        let wide_y = Pen {
            style: pen::PS_SOLID,
            width: (3, -12_860),
            color: 0,
        };
        let only_y = Pen {
            width: (0, 3_711),
            ..wide_y.clone()
        };

        assert!(
            stroke_attrs(&wide_y, &MappingState::default(), &transform)
                .contains(r#"stroke-width="3""#)
        );
        assert!(
            stroke_attrs(&only_y, &MappingState::default(), &transform)
                .contains(r#"stroke-width="1""#)
        );
    }
}
