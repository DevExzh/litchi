//! SVG style attribute generation
//!
//! Converts WMF pen, brush, and font properties to minimal SVG attributes.
//! Only includes non-default attributes to minimize output size.

use super::state::{Brush, Pen};
use super::transform::CoordinateTransform;

/// Format a number with minimal precision (no trailing zeros)
pub fn fmt_num(n: f64) -> String {
    if n.fract() == 0.0 && n.abs() < 1e10 {
        format!("{}", n as i64)
    } else {
        let rounded = (n * 100.0).round() / 100.0;
        let s = format!("{:.2}", rounded);
        s.trim_end_matches('0').trim_end_matches('.').to_string()
    }
}

/// Convert COLORREF to #RRGGBB
pub fn color_hex(c: u32) -> String {
    format!(
        "#{:02x}{:02x}{:02x}",
        c & 0xFF,
        (c >> 8) & 0xFF,
        (c >> 16) & 0xFF
    )
}

/// Generate fill attribute (only if non-default)
pub fn fill_attr(brush: &Brush, poly_fill_mode: u16) -> Option<String> {
    let mut attrs = String::with_capacity(48);

    if brush.style == 1 {
        // BS_NULL - no fill
        attrs.push_str(r#" fill="none""#);
    } else {
        // Solid or patterned fill
        attrs.push_str(&format!(r#" fill="{}""#, color_hex(brush.color)));

        // Add fill-rule based on poly_fill_mode (matches libwmf)
        // 1=ALTERNATE (evenodd), 2=WINDING (nonzero)
        if poly_fill_mode == 2 {
            attrs.push_str(r#" fill-rule="nonzero""#);
        } else if poly_fill_mode == 1 {
            attrs.push_str(r#" fill-rule="evenodd""#);
        }
    }

    Some(attrs)
}

/// Generate stroke attributes (matching libwmf behavior)
pub fn stroke_attrs(pen: &Pen, transform: &CoordinateTransform) -> String {
    let style = pen.style & 0x0F;

    // PS_NULL (5) - no stroke
    if style == 5 {
        return r#" stroke="none""#.to_string();
    }

    let mut attrs = String::with_capacity(96);

    // Stroke color
    attrs.push_str(&format!(r#" stroke="{}""#, color_hex(pen.color)));

    // Stroke width (average of width and height like libwmf)
    let width = transform.width(pen.width.max(1) as f64);
    attrs.push_str(&format!(r#" stroke-width="{}""#, fmt_num(width)));

    // Line cap
    let endcap = (pen.style >> 8) & 0x0F;
    let cap = match endcap {
        0x01 => "square", // PS_ENDCAP_SQUARE
        0x00 => "round",  // PS_ENDCAP_ROUND
        _ => "butt",      // PS_ENDCAP_FLAT (default)
    };
    if cap != "butt" {
        attrs.push_str(&format!(r#" stroke-linecap="{}""#, cap));
    }

    // Line join
    let join = (pen.style >> 12) & 0x0F;
    let join_style = match join {
        0x01 => "bevel", // PS_JOIN_BEVEL
        0x02 => "round", // PS_JOIN_ROUND
        _ => "miter",    // PS_JOIN_MITER (default)
    };
    if join_style != "miter" {
        attrs.push_str(&format!(r#" stroke-linejoin="{}""#, join_style));
    }

    // Dash array (scaled by pen width like libwmf)
    match style {
        1 => {
            // PS_DASH - dashed line (10x width dash + 10x width gap)
            let dash = width * 10.0;
            attrs.push_str(&format!(
                r#" stroke-dasharray="{},{}"#,
                fmt_num(dash),
                fmt_num(dash)
            ));
        },
        2 | 7 => {
            // PS_DOT or PS_ALTERNATE - dotted line (width dash + 2x width gap)
            let dash = width;
            let gap = width * 2.0;
            attrs.push_str(&format!(
                r#" stroke-dasharray="{},{}"#,
                fmt_num(dash),
                fmt_num(gap)
            ));
        },
        3 => {
            // PS_DASHDOT - dash-dot pattern
            let long = width * 10.0;
            let short = width;
            let gap = width * 2.0;
            attrs.push_str(&format!(
                r#" stroke-dasharray="{},{},{},{}"#,
                fmt_num(long),
                fmt_num(gap),
                fmt_num(short),
                fmt_num(gap)
            ));
        },
        4 => {
            // PS_DASHDOTDOT - dash-dot-dot pattern
            let long = width * 10.0;
            let short = width;
            let gap = width * 2.0;
            attrs.push_str(&format!(
                r#" stroke-dasharray="{},{},{},{},{},{}"#,
                fmt_num(long),
                fmt_num(gap),
                fmt_num(short),
                fmt_num(gap),
                fmt_num(short),
                fmt_num(gap)
            ));
        },
        _ => {}, // PS_SOLID (0) or PS_INSIDEFRAME (6) - no dasharray
    }

    attrs
}

/// Map WMF font name to generic family or keep specific
///
/// Maps common Windows fonts to generic CSS font families for better compatibility
/// and smaller SVG output. Follows common font fallback patterns used in libwmf.
pub fn map_font_family(name: &str) -> &str {
    match name {
        // Serif fonts
        "Times New Roman" | "Times" | "Georgia" | "Garamond" => "serif",
        // Sans-serif fonts
        "Arial" | "Helvetica" | "Verdana" | "Tahoma" | "Trebuchet MS" | "Arial Black" => {
            "sans-serif"
        },
        // Monospace fonts
        "Courier New" | "Courier" | "Consolas" | "Monaco" | "Lucida Console" => "monospace",
        // Cursive fonts
        "Comic Sans MS" | "Brush Script MT" => "cursive",
        // Fantasy fonts
        "Impact" | "Papyrus" => "fantasy",
        // Keep original name for other fonts
        _ => name,
    }
}
