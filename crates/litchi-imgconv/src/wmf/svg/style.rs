//! SVG style and XML serialization helpers.

use super::super::constants::{bk_mode, brush, pen};
use super::state::{Brush, MappingState, Pen};
use super::transform::CoordinateTransform;
use crate::svg_utils::{write_color_hex, write_num};

pub(super) fn fill_attrs(
    brush_value: &Brush,
    poly_fill_mode: u16,
    pattern: Option<&str>,
) -> String {
    let mut attrs = String::with_capacity(64);
    match brush_value.style {
        brush::BS_NULL => attrs.push_str(r#" fill="none""#),
        brush::BS_HATCHED if pattern.is_some() => {
            attrs.push_str(r#" fill="url(#"#);
            attrs.push_str(pattern.unwrap_or_default());
            attrs.push_str(r#")""#);
        },
        brush::BS_SOLID | brush::BS_HATCHED => {
            attrs.push_str(r#" fill=""#);
            write_color_hex(&mut attrs, brush_value.color);
            attrs.push('"');
        },
        _ => attrs.push_str(r#" fill="none""#),
    }
    attrs.push_str(if poly_fill_mode == 2 {
        r#" fill-rule="nonzero""#
    } else {
        r#" fill-rule="evenodd""#
    });
    attrs
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

    let (width_x, width_y) = transform.logical_vector(
        mapping,
        f64::from(pen_value.width.0).abs(),
        f64::from(pen_value.width.1).abs(),
    );
    let width = match (width_x > 0.0, width_y > 0.0) {
        (true, true) => (width_x + width_y) / 2.0,
        (true, false) => width_x,
        (false, true) => width_y,
        (false, false) => transform.device_width(1.0).max(0.01),
    };

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

pub(super) fn hatch_definition(
    id: &str,
    brush_value: &Brush,
    background_mode: u16,
    background_color: u32,
) -> String {
    let mut definition = String::with_capacity(320);
    definition.push_str(r#"<pattern id=""#);
    escape_xml_attr_into(&mut definition, id);
    definition.push_str(r#"" patternUnits="userSpaceOnUse" width="8" height="8">"#);
    if background_mode == bk_mode::OPAQUE {
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

pub(super) fn map_font_family(name: &str) -> &str {
    if name.eq_ignore_ascii_case("Times New Roman")
        || name.eq_ignore_ascii_case("Times")
        || name.eq_ignore_ascii_case("Georgia")
        || name.eq_ignore_ascii_case("Garamond")
    {
        "serif"
    } else if name.eq_ignore_ascii_case("Arial")
        || name.eq_ignore_ascii_case("Helvetica")
        || name.eq_ignore_ascii_case("Verdana")
        || name.eq_ignore_ascii_case("Tahoma")
        || name.eq_ignore_ascii_case("Trebuchet MS")
        || name.eq_ignore_ascii_case("Arial Black")
    {
        "sans-serif"
    } else if name.eq_ignore_ascii_case("Courier New")
        || name.eq_ignore_ascii_case("Courier")
        || name.eq_ignore_ascii_case("Consolas")
        || name.eq_ignore_ascii_case("Monaco")
        || name.eq_ignore_ascii_case("Lucida Console")
    {
        "monospace"
    } else {
        name
    }
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
}
