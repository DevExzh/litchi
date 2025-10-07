// OMML property parsing
//
// This module handles parsing of OMML property elements and attributes.
// Properties control styling, spacing, alignment, and other formatting aspects.

use crate::formula::omml::elements::ElementProperties;
use crate::formula::omml::attributes::*;

/// Parse run properties (m:rPr)
pub fn parse_run_properties(attrs: &[quick_xml::events::attributes::Attribute]) -> ElementProperties {
    let mut properties = ElementProperties::default();

    for attr in attrs {
        if let Ok(key) = std::str::from_utf8(attr.key.as_ref())
            && let Ok(value) = std::str::from_utf8(&attr.value) {
                match key {
                    "scr" | "m:scr" => {
                        // Script/style type
                        properties.style = Some(value.to_string());
                    }
                    "sty" | "m:sty" => {
                        // Math style (display/text)
                        properties.style = Some(value.to_string());
                    }
                    "nor" | "m:nor" => {
                        // Normal text
                        properties.font = Some(value.to_string());
                    }
                    "lit" | "m:lit" => {
                        // Literal text flag
                        properties.hide = Some(value == "1" || value == "true");
                    }
                    _ => {}
                }
            }
    }

    properties
}

/// Parse fraction properties (m:fPr)
pub fn parse_fraction_properties(attrs: &[quick_xml::events::attributes::Attribute]) -> ElementProperties {
    let mut properties = ElementProperties::default();

    for attr in attrs {
        if let Ok(key) = std::str::from_utf8(attr.key.as_ref())
            && let Ok(value) = std::str::from_utf8(&attr.value) {
                match key {
                    "type" | "m:type" => {
                        // Fraction type (bar, noBar, etc.)
                        properties.style = Some(value.to_string());
                    }
                    _ => {}
                }
            }
    }

    properties
}

/// Parse delimiter properties (m:dPr)
pub fn parse_delimiter_properties(attrs: &[quick_xml::events::attributes::Attribute]) -> ElementProperties {
    let mut properties = ElementProperties::default();

    for attr in attrs {
        if let Ok(key) = std::str::from_utf8(attr.key.as_ref())
            && let Ok(value) = std::str::from_utf8(&attr.value) {
                match key {
                    "begChr" | "m:begChr" => {
                        // Beginning character
                        properties.delimiter_open_char = Some(value.to_string());
                    }
                    "endChr" | "m:endChr" => {
                        // Ending character
                        properties.delimiter_close_char = Some(value.to_string());
                    }
                    "sepChr" | "m:sepChr" => {
                        // Separator character
                        properties.delimiter_separator_char = Some(value.to_string());
                    }
                    "grow" | "m:grow" => {
                        // Grow to fit content
                        properties.hide = Some(value == "1" || value == "true");
                    }
                    "shp" | "m:shp" => {
                        // Shape (centered, match, etc.)
                        properties.alignment = Some(value.to_string());
                    }
                    _ => {}
                }
            }
    }

    properties
}

/// Parse n-ary operator properties (m:naryPr)
pub fn parse_nary_properties(attrs: &[quick_xml::events::attributes::Attribute]) -> ElementProperties {
    let mut properties = ElementProperties::default();

    for attr in attrs {
        if let Ok(key) = std::str::from_utf8(attr.key.as_ref())
            && let Ok(value) = std::str::from_utf8(&attr.value) {
                match key {
                    "chr" | "m:chr" => {
                        // Operator character
                        properties.chr = Some(value.to_string());
                    }
                    "grow" | "m:grow" => {
                        // Grow to fit content
                        properties.hide = Some(value == "1" || value == "true");
                    }
                    "subHide" | "m:subHide" => {
                        // Hide subscript
                        properties.hide = Some(value == "1" || value == "true");
                    }
                    "supHide" | "m:supHide" => {
                        // Hide superscript
                        properties.hide = Some(value == "1" || value == "true");
                    }
                    _ => {}
                }
            }
    }

    properties
}

/// Parse accent properties (m:accPr)
pub fn parse_accent_properties(attrs: &[quick_xml::events::attributes::Attribute]) -> ElementProperties {
    let mut properties = ElementProperties::default();

    for attr in attrs {
        if let Ok(key) = std::str::from_utf8(attr.key.as_ref())
            && let Ok(value) = std::str::from_utf8(&attr.value) {
                match key {
                    "chr" | "m:chr" => {
                        // Accent character
                        properties.chr = Some(value.to_string());
                    }
                    _ => {}
                }
            }
    }

    properties
}

/// Parse matrix properties (m:mPr)
pub fn parse_matrix_properties(attrs: &[quick_xml::events::attributes::Attribute]) -> ElementProperties {
    let mut properties = ElementProperties::default();

    for attr in attrs {
        if let Ok(key) = std::str::from_utf8(attr.key.as_ref())
            && let Ok(value) = std::str::from_utf8(&attr.value) {
                match key {
                    "baseJc" | "m:baseJc" => {
                        // Baseline justification
                        properties.alignment = Some(value.to_string());
                    }
                    "plcHide" | "m:plcHide" => {
                        // Hide placeholder
                        properties.hide = Some(value == "1" || value == "true");
                    }
                    "rSp" | "m:rSp" => {
                        // Row spacing
                        properties.spacing = Some(value.to_string());
                    }
                    "cSp" | "m:cSp" => {
                        // Column spacing
                        properties.spacing = Some(value.to_string());
                    }
                    "cGp" | "m:cGp" => {
                        // Column gap
                        properties.spacing = Some(value.to_string());
                    }
                    "mcs" | "m:mcs" => {
                        // Matrix column spacing
                        properties.spacing = Some(value.to_string());
                    }
                    "mcsJc" | "m:mcsJc" => {
                        // Matrix column spacing justification
                        properties.alignment = Some(value.to_string());
                    }
                    _ => {}
                }
            }
    }

    properties
}

/// Parse group character properties (m:groupChrPr)
pub fn parse_group_char_properties(attrs: &[quick_xml::events::attributes::Attribute]) -> ElementProperties {
    let mut properties = ElementProperties::default();

    for attr in attrs {
        if let Ok(key) = std::str::from_utf8(attr.key.as_ref())
            && let Ok(value) = std::str::from_utf8(&attr.value) {
                match key {
                    "chr" | "m:chr" => {
                        // Group character
                        properties.chr = Some(value.to_string());
                    }
                    "pos" | "m:pos" => {
                        // Position (top/bot)
                        properties.alignment = Some(value.to_string());
                    }
                    "vertJc" | "m:vertJc" => {
                        // Vertical justification
                        properties.alignment = Some(value.to_string());
                    }
                    _ => {}
                }
            }
    }

    properties
}

/// Parse equation array properties (m:eqArrPr)
pub fn parse_eq_arr_properties(attrs: &[quick_xml::events::attributes::Attribute]) -> ElementProperties {
    let mut properties = ElementProperties::default();

    for attr in attrs {
        if let Ok(key) = std::str::from_utf8(attr.key.as_ref())
            && let Ok(value) = std::str::from_utf8(&attr.value) {
                match key {
                    "baseJc" | "m:baseJc" => {
                        // Baseline justification
                        properties.alignment = Some(value.to_string());
                    }
                    "maxDist" | "m:maxDist" => {
                        // Maximum distance
                        properties.spacing = Some(value.to_string());
                    }
                    "objDist" | "m:objDist" => {
                        // Object distance
                        properties.spacing = Some(value.to_string());
                    }
                    "rSp" | "m:rSp" => {
                        // Row spacing
                        properties.spacing = Some(value.to_string());
                    }
                    "rSpRule" | "m:rSpRule" => {
                        // Row spacing rule
                        properties.spacing = Some(value.to_string());
                    }
                    _ => {}
                }
            }
    }

    properties
}

/// Parse limit properties (m:lim)
pub fn parse_limit_properties(attrs: &[quick_xml::events::attributes::Attribute]) -> ElementProperties {
    let mut properties = ElementProperties::default();

    for attr in attrs {
        if let Ok(key) = std::str::from_utf8(attr.key.as_ref())
            && let Ok(value) = std::str::from_utf8(&attr.value) {
                match key {
                    "type" | "m:type" => {
                        // Limit type (undOvr, subSup)
                        properties.style = Some(value.to_string());
                    }
                    _ => {}
                }
            }
    }

    properties
}

/// Parse bar properties (m:barPr)
pub fn parse_bar_properties(attrs: &[quick_xml::events::attributes::Attribute]) -> ElementProperties {
    let mut properties = ElementProperties::default();

    for attr in attrs {
        if let Ok(key) = std::str::from_utf8(attr.key.as_ref())
            && let Ok(value) = std::str::from_utf8(&attr.value) {
                match key {
                    "pos" | "m:pos" => {
                        // Position (top/bot)
                        properties.alignment = Some(value.to_string());
                    }
                    _ => {}
                }
            }
    }

    properties
}

/// Parse box properties (m:boxPr)
pub fn parse_box_properties(attrs: &[quick_xml::events::attributes::Attribute]) -> ElementProperties {
    let mut properties = ElementProperties::default();

    for attr in attrs {
        if let Ok(key) = std::str::from_utf8(attr.key.as_ref())
            && let Ok(value) = std::str::from_utf8(&attr.value) {
                match key {
                    "opEmu" | "m:opEmu" => {
                        // Operator emulation
                        properties.hide = Some(value == "1" || value == "true");
                    }
                    "noBreak" | "m:noBreak" => {
                        // No break
                        properties.hide = Some(value == "1" || value == "true");
                    }
                    "diff" | "m:diff" => {
                        // Differential
                        properties.hide = Some(value == "1" || value == "true");
                    }
                    "brk" | "m:brk" => {
                        // Break
                        properties.hide = Some(value == "1" || value == "true");
                    }
                    "aln" | "m:aln" => {
                        // Alignment
                        properties.alignment = Some(value.to_string());
                    }
                    _ => {}
                }
            }
    }

    properties
}

/// Parse border box properties (m:borderBoxPr)
pub fn parse_border_box_properties(attrs: &[quick_xml::events::attributes::Attribute]) -> ElementProperties {
    let mut properties = ElementProperties::default();

    for attr in attrs {
        if let Ok(key) = std::str::from_utf8(attr.key.as_ref())
            && let Ok(value) = std::str::from_utf8(&attr.value) {
                match key {
                    "hideTop" | "m:hideTop" => {
                        properties.hide = Some(value == "1" || value == "true");
                    }
                    "hideBot" | "m:hideBot" => {
                        properties.hide = Some(value == "1" || value == "true");
                    }
                    "hideLeft" | "m:hideLeft" => {
                        properties.hide = Some(value == "1" || value == "true");
                    }
                    "hideRight" | "m:hideRight" => {
                        properties.hide = Some(value == "1" || value == "true");
                    }
                    "strikeH" | "m:strikeH" => {
                        properties.strike_through = Some(value == "1" || value == "true");
                    }
                    "strikeV" | "m:strikeV" => {
                        properties.double_strike_through = Some(value == "1" || value == "true");
                    }
                    "strikeBLTR" | "m:strikeBLTR" => {
                        // Strike bottom-left to top-right
                    }
                    "strikeTLBR" | "m:strikeTLBR" => {
                        // Strike top-left to bottom-right
                    }
                    _ => {}
                }
            }
    }

    properties
}

/// Parse phantom properties (m:phantPr)
pub fn parse_phantom_properties(attrs: &[quick_xml::events::attributes::Attribute]) -> ElementProperties {
    let mut properties = ElementProperties::default();

    for attr in attrs {
        if let Ok(key) = std::str::from_utf8(attr.key.as_ref())
            && let Ok(value) = std::str::from_utf8(&attr.value) {
                match key {
                    "show" | "m:show" => {
                        // Show phantom content
                        properties.hide = Some(!(value == "1" || value == "true"));
                    }
                    "zeroWid" | "m:zeroWid" => {
                        // Zero width
                        properties.hide = Some(value == "1" || value == "true");
                    }
                    "zeroAsc" | "m:zeroAsc" => {
                        // Zero ascent
                        properties.hide = Some(value == "1" || value == "true");
                    }
                    "zeroDesc" | "m:zeroDesc" => {
                        // Zero descent
                        properties.hide = Some(value == "1" || value == "true");
                    }
                    "transp" | "m:transp" => {
                        // Transparent
                        properties.hide = Some(value == "1" || value == "true");
                    }
                    _ => {}
                }
            }
    }

    properties
}

/// Parse spacing properties (m:sPre, m:sPost, etc.)
pub fn parse_spacing_properties(attrs: &[quick_xml::events::attributes::Attribute]) -> ElementProperties {
    let mut properties = ElementProperties::default();

    for attr in attrs {
        if let Ok(key) = std::str::from_utf8(attr.key.as_ref())
            && let Ok(value) = std::str::from_utf8(&attr.value) {
                match key {
                    "val" | "m:val" => {
                        properties.spacing = Some(value.to_string());
                    }
                    _ => {}
                }
            }
    }

    properties
}

/// Parse general element properties from attributes
pub fn parse_general_properties(attrs: &[quick_xml::events::attributes::Attribute]) -> ElementProperties {
    let mut properties = ElementProperties::default();

    for attr in attrs {
        if let Ok(key) = std::str::from_utf8(attr.key.as_ref())
            && let Ok(value) = std::str::from_utf8(&attr.value) {
                match key {
                    // Style and formatting
                    "scr" | "m:scr" => properties.math_variant = Some(value.to_string()),
                    "sty" | "m:sty" => properties.display_style = Some(matches!(value, "d" | "display" | "1" | "true")),

                    // Size and scaling
                    "sz" | "m:sz" => properties.size = Some(value.to_string()),
                    "minSz" | "m:minSz" => properties.min_size = Some(value.to_string()),
                    "maxSz" | "m:maxSz" => properties.max_size = Some(value.to_string()),
                    "scrLvl" | "m:scrLvl" => {
                        if let Some(lvl) = parse_int_simd(value) {
                            properties.script_level = Some(lvl);
                        }
                    }

                    // Color and font
                    "color" | "m:color" => properties.color = Some(value.to_string()),
                    "font" | "m:font" => properties.font = Some(value.to_string()),
                    "nor" | "m:nor" => properties.font = Some(value.to_string()),

                    // Layout and positioning
                    "aln" | "m:aln" => properties.alignment = Some(value.to_string()),
                    "alnScr" | "m:alnScr" => properties.alignment = Some(value.to_string()),
                    "vertJc" | "m:vertJc" => properties.vertical_alignment = Some(value.to_string()),
                    "baseJc" | "m:baseJc" => properties.alignment = Some(value.to_string()),

                    // Characters and symbols
                    "chr" | "m:chr" => properties.chr = Some(value.to_string()),

                    // Spacing
                    "val" | "m:val" => properties.spacing = Some(value.to_string()),

                    // Fraction properties
                    "type" | "m:type" => properties.fraction_type = Some(value.to_string()),
                    "lnThick" | "m:lnThick" => properties.fraction_line_thickness = Some(value.to_string()),

                    // Matrix properties
                    "rSp" | "m:rSp" => properties.matrix_row_spacing = Some(value.to_string()),
                    "cSp" | "m:cSp" => properties.matrix_column_spacing = Some(value.to_string()),

                    // Accent properties
                    "pos" | "m:pos" => properties.accent_position = Some(value.to_string()),

                    // Box properties
                    "diff" | "m:diff" => properties.box_differential = Some(matches!(value, "1" | "true")),
                    "opEmu" | "m:opEmu" => properties.box_operator_emulation = Some(matches!(value, "1" | "true")),
                    "brk" | "m:brk" => properties.box_break = Some(matches!(value, "1" | "true")),
                    "noBreak" | "m:noBreak" => properties.box_no_break = Some(matches!(value, "1" | "true")),

                    // Phantom properties
                    "show" | "m:show" => properties.phantom_show = Some(matches!(value, "1" | "true")),
                    "zeroWid" | "m:zeroWid" => properties.phantom_zero_width = Some(matches!(value, "1" | "true")),
                    "zeroAsc" | "m:zeroAsc" => properties.phantom_zero_ascent = Some(matches!(value, "1" | "true")),
                    "zeroDesc" | "m:zeroDesc" => properties.phantom_zero_descent = Some(matches!(value, "1" | "true")),
                    "transp" | "m:transp" => properties.phantom_transparent = Some(matches!(value, "1" | "true")),

                    // Border box properties
                    "hideTop" | "m:hideTop" => properties.border_hide_top = Some(matches!(value, "1" | "true")),
                    "hideBot" | "m:hideBot" => properties.border_hide_bottom = Some(matches!(value, "1" | "true")),
                    "hideLeft" | "m:hideLeft" => properties.border_hide_left = Some(matches!(value, "1" | "true")),
                    "hideRight" | "m:hideRight" => properties.border_hide_right = Some(matches!(value, "1" | "true")),
                    "strikeH" | "m:strikeH" => properties.border_strike_horizontal = Some(matches!(value, "1" | "true")),
                    "strikeV" | "m:strikeV" => properties.border_strike_vertical = Some(matches!(value, "1" | "true")),
                    "strikeBLTR" | "m:strikeBLTR" => properties.border_strike_bltr = Some(matches!(value, "1" | "true")),
                    "strikeTLBR" | "m:strikeTLBR" => properties.border_strike_tlbr = Some(matches!(value, "1" | "true")),

                    // Equation array properties
                    "maxDist" | "m:maxDist" => properties.eq_arr_max_distance = Some(value.to_string()),
                    "objDist" | "m:objDist" => properties.eq_arr_object_distance = Some(value.to_string()),
                    "rSpRule" | "m:rSpRule" => properties.eq_arr_row_spacing_rule = Some(value.to_string()),

                    // N-ary operator properties
                    "subHide" | "m:subHide" => properties.nary_hide_sub = Some(matches!(value, "1" | "true")),
                    "supHide" | "m:supHide" => properties.nary_hide_sup = Some(matches!(value, "1" | "true")),
                    "grow" | "m:grow" => properties.nary_operator_grow = Some(matches!(value, "1" | "true")),

                    // Delimiter properties
                    "sepChr" | "m:sepChr" => properties.delimiter_separator_char = Some(value.to_string()),
                    "begChr" | "m:begChr" => properties.delimiter_open_char = Some(value.to_string()),
                    "endChr" | "m:endChr" => properties.delimiter_close_char = Some(value.to_string()),
                    "shp" | "m:shp" => properties.delimiter_shape = Some(value.to_string()),

                    // Radical properties
                    "degHide" | "m:degHide" => properties.radical_hide_degree = Some(matches!(value, "1" | "true")),

                    // Run properties
                    "lit" | "m:lit" => properties.run_literal = Some(matches!(value, "1" | "true")),

                    // Visibility and rendering
                    "hide" | "m:hide" => properties.hide = Some(matches!(value, "1" | "true")),
                    "strike" | "m:strike" => properties.strike_through = Some(matches!(value, "1" | "true")),
                    "dstrike" | "m:dstrike" => properties.double_strike_through = Some(matches!(value, "1" | "true")),

                    // Line styles
                    "u" | "m:u" => properties.underline = Some(value.to_string()),
                    "o" | "m:o" => properties.overline = Some(value.to_string()),

                    // Special positioning attributes
                    "den" | "m:den" => properties.alignment = Some("denominator".to_string()),
                    "num" | "m:num" => properties.alignment = Some("numerator".to_string()),

                    _ => {} // Ignore unknown attributes
                }
            }
    }

    properties
}
