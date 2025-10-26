use crate::formula::ast::{
    BreakType, FractionType, LineStyle, Position, ShapeType, StrikeStyle, VerticalAlignment, *,
};
use crate::formula::omml::elements::ElementProperties;
use crate::formula::omml::lookup::*;

/// SIMD-accelerated numeric parsing functions
/// Fast integer parsing using atoi_simd
pub fn parse_int_simd(value: &str) -> Option<i32> {
    atoi_simd::parse(value.as_bytes()).ok()
}

/// Fast float parsing using fast_float2
pub fn parse_float_simd(value: &str) -> Option<f32> {
    fast_float2::parse(value).ok()
}

/// Fast boolean parsing with lookup
pub fn parse_bool_fast(value: &str) -> Option<bool> {
    parse_bool_value(value)
}

/// Fast alignment parsing with lookup
#[allow(dead_code)] // Reserved for future optimizations
pub fn parse_alignment_fast(value: &str) -> Option<Alignment> {
    parse_alignment_value(value)
}

/// Fast style parsing with lookup
#[allow(dead_code)] // Reserved for future optimizations
pub fn parse_style_fast(value: &str) -> Option<StyleType> {
    parse_style_value(value)
}

/// Parse a fence type from OMML attribute values
pub fn parse_fence_type(open: Option<&str>, close: Option<&str>) -> (Option<Fence>, Option<Fence>) {
    // Fast fence parsing using direct character matching
    let open_fence = open.and_then(|s| match s {
        "(" | "&#40;" | "paren" => Some(Fence::Paren),
        "[" | "&#91;" | "bracket" => Some(Fence::Bracket),
        "{" | "&#123;" | "brace" => Some(Fence::Brace),
        "⟨" | "&#10216;" | "&#8240;" | "langle" => Some(Fence::Angle),
        "|" | "pipe" => Some(Fence::Pipe),
        "‖" | "&#8214;" | "||" | "doublepipe" => Some(Fence::DoublePipe),
        "⌊" | "&#8970;" | "lfloor" => Some(Fence::Floor),
        "⌈" | "&#8971;" | "lceil" => Some(Fence::Ceiling),
        "⟪" | "&#10218;" => Some(Fence::AngleBracket),
        "⟦" | "&#10214;" => Some(Fence::SquareBracket),
        "⦃" => Some(Fence::CurlyBrace),
        _ => None,
    });

    let close_fence = close.and_then(|s| match s {
        ")" | "&#41;" | "paren" => Some(Fence::Paren),
        "]" | "&#93;" | "bracket" => Some(Fence::Bracket),
        "}" | "&#125;" | "brace" => Some(Fence::Brace),
        "⟩" | "&#10217;" | "&#8241;" | "rangle" => Some(Fence::Angle),
        "|" | "pipe" => Some(Fence::Pipe),
        "‖" | "&#8214;" | "||" | "doublepipe" => Some(Fence::DoublePipe),
        "⌋" | "&#8971;" | "rfloor" => Some(Fence::Floor),
        "⌉" | "&#8972;" | "rceil" => Some(Fence::Ceiling),
        "⟫" | "&#10219;" => Some(Fence::AngleBracket),
        "⟧" | "&#10215;" => Some(Fence::SquareBracket),
        "⦄" => Some(Fence::CurlyBrace),
        _ => None,
    });

    (open_fence, close_fence)
}

/// Parse a large operator type from OMML attribute values
pub fn parse_large_operator(op: Option<&str>) -> Option<LargeOperator> {
    op.and_then(get_large_operator)
}

/// Parse an accent type from OMML attribute values
pub fn parse_accent_type(acc: Option<&str>) -> Option<AccentType> {
    acc.and_then(get_accent_type)
}

/// Parse a matrix fence type from OMML attribute values
pub fn parse_matrix_fence(mcs: Option<&str>) -> Option<MatrixFence> {
    match mcs {
        Some("&#40;") | Some("(") | Some("paren") => Some(MatrixFence::Paren),
        Some("&#91;") | Some("[") | Some("bracket") => Some(MatrixFence::Bracket),
        Some("&#123;") | Some("{") | Some("brace") => Some(MatrixFence::Brace),
        Some("|") | Some("pipe") => Some(MatrixFence::Pipe),
        Some("&#8214;") | Some("||") | Some("doublepipe") => Some(MatrixFence::DoublePipe),
        _ => None,
    }
}

/// Parse style type from OMML attribute values
#[allow(dead_code)] // Reserved for future use in property parsing
pub fn parse_style_type(scr: Option<&str>) -> Option<StyleType> {
    scr.and_then(parse_style_fast)
}

/// Parse element properties from OMML attributes
/// This is a general-purpose property parser used as a fallback
#[allow(dead_code)] // Used indirectly through batch parsing functions
pub fn parse_element_properties(
    attrs: &[quick_xml::events::attributes::Attribute],
) -> ElementProperties {
    let mut properties = ElementProperties::default();

    for attr in attrs {
        if let Ok(key) = std::str::from_utf8(attr.key.as_ref())
            && let Ok(value) = std::str::from_utf8(&attr.value)
        {
            match key {
                "val" | "m:val" => {
                    // This could be various values depending on context
                    // Store as string for now, will be interpreted by caller
                },
                "scr" | "m:scr" => {
                    properties.style = Some(value.to_string());
                },
                "sty" | "m:sty" => {
                    // Math style (display/text)
                    properties.style = Some(value.to_string());
                },
                "nor" | "m:nor" => {
                    // Normal text
                    properties.font = Some(value.to_string());
                },
                "lit" | "m:lit" => {
                    // Literal text
                },
                "aln" | "m:aln" => {
                    properties.alignment = Some(value.to_string());
                },
                "alnScr" | "m:alnScr" => {
                    // Alignment script
                },
                "den" | "m:den" => {
                    // Denominator alignment
                },
                "num" | "m:num" => {
                    // Numerator alignment
                },
                "chr" | "m:chr" => {
                    properties.chr = Some(value.to_string());
                },
                _ => {},
            }
        }
    }

    properties
}

/// Extract attribute value as string (optimized for performance)
///
/// This function is called extensively during OMML parsing, so it's heavily optimized:
/// - Uses byte-level comparison to avoid UTF-8 decoding when possible
/// - Pre-computes the m: prefix to avoid format! allocation in the loop
/// - Inlined for better performance
/// - Returns owned String only when a match is found
#[inline]
pub fn get_attribute_value(
    attrs: &[quick_xml::events::attributes::Attribute],
    key: &str,
) -> Option<String> {
    // Pre-compute the m: prefixed key once to avoid allocation in the loop
    let key_bytes = key.as_bytes();
    let mut m_prefixed = smallvec::SmallVec::<[u8; 32]>::with_capacity(key.len() + 2);
    m_prefixed.extend_from_slice(b"m:");
    m_prefixed.extend_from_slice(key_bytes);

    // Fast path: iterate attributes with byte-level comparison
    for attr in attrs {
        let attr_key = attr.key.as_ref();

        // Fast byte comparison (avoids UTF-8 validation overhead)
        if attr_key == key_bytes || attr_key == m_prefixed.as_slice() {
            // Only decode UTF-8 when we have a match
            if let Ok(value) = std::str::from_utf8(&attr.value) {
                return Some(value.to_string());
            }
        }
    }
    None
}

/// Extract attribute value as integer with SIMD acceleration
pub fn get_attribute_value_int(
    attrs: &[quick_xml::events::attributes::Attribute],
    key: &str,
) -> Option<i32> {
    get_attribute_value(attrs, key).and_then(|s| parse_int_simd(&s))
}

/// Extract attribute value as float with SIMD acceleration
pub fn get_attribute_value_float(
    attrs: &[quick_xml::events::attributes::Attribute],
    key: &str,
) -> Option<f32> {
    get_attribute_value(attrs, key).and_then(|s| parse_float_simd(&s))
}

/// Extract attribute value as boolean with fast lookup
/// Part of the attribute extraction API for element handlers
#[allow(dead_code)]
pub fn get_attribute_value_bool(
    attrs: &[quick_xml::events::attributes::Attribute],
    key: &str,
) -> Option<bool> {
    get_attribute_value(attrs, key).and_then(|s| parse_bool_fast(&s))
}

/// Extract attribute value as space type
/// Part of the attribute extraction API for spacing elements
#[allow(dead_code)]
pub fn get_attribute_value_space(
    attrs: &[quick_xml::events::attributes::Attribute],
    key: &str,
) -> Option<SpaceType> {
    get_attribute_value(attrs, key).and_then(|s| parse_space_type(Some(&s)))
}

/// Extract attribute value as alignment
/// Part of the attribute extraction API for positioning
#[allow(dead_code)]
pub fn get_attribute_value_alignment(
    attrs: &[quick_xml::events::attributes::Attribute],
    key: &str,
) -> Option<Alignment> {
    get_attribute_value(attrs, key).and_then(|s| parse_alignment_value(&s))
}

/// Extract attribute value as vertical alignment
/// Part of the attribute extraction API for vertical positioning
#[allow(dead_code)]
pub fn get_attribute_value_vertical_alignment(
    attrs: &[quick_xml::events::attributes::Attribute],
    key: &str,
) -> Option<VerticalAlignment> {
    get_attribute_value(attrs, key).and_then(|s| parse_vertical_alignment(Some(&s)))
}

/// Extract attribute value as position
/// Part of the attribute extraction API for position properties
#[allow(dead_code)]
pub fn get_attribute_value_position(
    attrs: &[quick_xml::events::attributes::Attribute],
    key: &str,
) -> Option<Position> {
    get_attribute_value(attrs, key).and_then(|s| parse_position_type(Some(&s)))
}

/// Extract attribute value as fraction type
/// Part of the attribute extraction API for fraction elements
#[allow(dead_code)]
pub fn get_attribute_value_fraction_type(
    attrs: &[quick_xml::events::attributes::Attribute],
    key: &str,
) -> Option<FractionType> {
    get_attribute_value(attrs, key).and_then(|s| parse_fraction_type(Some(&s)))
}

/// Extract attribute value as shape type
/// Part of the attribute extraction API for shape properties
#[allow(dead_code)]
pub fn get_attribute_value_shape(
    attrs: &[quick_xml::events::attributes::Attribute],
    key: &str,
) -> Option<ShapeType> {
    get_attribute_value(attrs, key).and_then(|s| parse_shape_type(Some(&s)))
}

/// Extract attribute value as break type
/// Part of the attribute extraction API for break properties
#[allow(dead_code)]
pub fn get_attribute_value_break(
    attrs: &[quick_xml::events::attributes::Attribute],
    key: &str,
) -> Option<BreakType> {
    get_attribute_value(attrs, key).and_then(|s| parse_break_type(Some(&s)))
}

/// Extract attribute value as line style
/// Part of the attribute extraction API for line styling
#[allow(dead_code)]
pub fn get_attribute_value_line_style(
    attrs: &[quick_xml::events::attributes::Attribute],
    key: &str,
) -> Option<LineStyle> {
    get_attribute_value(attrs, key).and_then(|s| parse_line_style(Some(&s)))
}

/// Extract attribute value as strike style
/// Part of the attribute extraction API for strike-through styling
#[allow(dead_code)]
pub fn get_attribute_value_strike_style(
    attrs: &[quick_xml::events::attributes::Attribute],
    key: &str,
) -> Option<StrikeStyle> {
    get_attribute_value(attrs, key).and_then(|s| parse_strike_style(Some(&s)))
}

/// Parse spacing type from OMML attribute values
pub fn parse_space_type(val: Option<&str>) -> Option<SpaceType> {
    match val {
        Some("1") | Some("thinmathspace") | Some("thin") => Some(SpaceType::Thin),
        Some("2") | Some("mediummathspace") | Some("medium") => Some(SpaceType::Medium),
        Some("3") | Some("thickmathspace") | Some("thick") => Some(SpaceType::Thick),
        Some("4") | Some("quad") => Some(SpaceType::Quad),
        Some("5") | Some("qquad") | Some("doublequad") => Some(SpaceType::QQuad),
        Some("-1") | Some("negative") | Some("negthinmathspace") => Some(SpaceType::Negative),
        _ => None,
    }
}

/// Parse alignment type from OMML attribute values
#[allow(dead_code)] // API for future alignment parsing enhancements
pub fn parse_alignment_type(val: Option<&str>) -> Option<String> {
    match val {
        Some("left") | Some("l") => Some("left".to_string()),
        Some("center") | Some("c") => Some("center".to_string()),
        Some("right") | Some("r") => Some("right".to_string()),
        Some("top") | Some("t") => Some("top".to_string()),
        Some("bottom") | Some("b") => Some("bottom".to_string()),
        Some("baseline") | Some("base") => Some("baseline".to_string()),
        Some("axis") | Some("ax") => Some("axis".to_string()),
        Some("centered") | Some("cen") => Some("centered".to_string()),
        Some("match") | Some("mat") => Some("match".to_string()),
        _ => val.map(|s| s.to_string()),
    }
}

/// Parse operator form from OMML attribute values
#[allow(dead_code)] // API for future operator property parsing
pub fn parse_operator_form(val: Option<&str>) -> Option<String> {
    match val {
        Some("pre") | Some("prefix") => Some("prefix".to_string()),
        Some("post") | Some("postfix") => Some("postfix".to_string()),
        Some("in") | Some("infix") => Some("infix".to_string()),
        _ => val.map(|s| s.to_string()),
    }
}

/// Parse math variant from OMML attribute values
pub fn parse_math_variant(val: Option<&str>) -> Option<String> {
    match val {
        Some("normal") | Some("nor") => Some("normal".to_string()),
        Some("bold") | Some("b") => Some("bold".to_string()),
        Some("italic") | Some("i") => Some("italic".to_string()),
        Some("bold-italic") | Some("bi") => Some("bold-italic".to_string()),
        Some("double-struck") | Some("ds") => Some("double-struck".to_string()),
        Some("bold-fraktur") | Some("bfr") => Some("bold-fraktur".to_string()),
        Some("script") | Some("sc") => Some("script".to_string()),
        Some("bold-script") | Some("bsc") => Some("bold-script".to_string()),
        Some("fraktur") | Some("fr") => Some("fraktur".to_string()),
        Some("sans-serif") | Some("ss") => Some("sans-serif".to_string()),
        Some("sans-serif-bold") | Some("ssb") => Some("sans-serif-bold".to_string()),
        Some("sans-serif-italic") | Some("ssi") => Some("sans-serif-italic".to_string()),
        Some("sans-serif-bold-italic") | Some("ssbi") => Some("sans-serif-bold-italic".to_string()),
        Some("monospace") | Some("m") => Some("monospace".to_string()),
        _ => val.map(|s| s.to_string()),
    }
}

/// Parse display style from OMML attribute values
pub fn parse_display_style(val: Option<&str>) -> Option<bool> {
    match val {
        Some("d") | Some("display") | Some("1") | Some("true") => Some(true),
        Some("t") | Some("text") | Some("0") | Some("false") => Some(false),
        _ => None,
    }
}

/// Parse script level from OMML attribute values
pub fn parse_script_level(val: Option<&str>) -> Option<i32> {
    val.and_then(|s| s.parse().ok())
}

/// Parse boolean attribute with multiple possible true values
#[allow(dead_code)] // API for general attribute parsing
pub fn parse_bool_attribute(val: Option<&str>) -> Option<bool> {
    match val {
        Some("1") | Some("true") | Some("on") | Some("yes") => Some(true),
        Some("0") | Some("false") | Some("off") | Some("no") => Some(false),
        _ => None,
    }
}

/// Parse color attribute (hex, named colors, etc.)
#[allow(dead_code)] // API for color attribute parsing
pub fn parse_color_attribute(val: Option<&str>) -> Option<String> {
    val.map(|s| s.to_string())
}

/// Parse font size attribute
#[allow(dead_code)] // API for font size parsing
pub fn parse_font_size(val: Option<&str>) -> Option<String> {
    val.map(|s| s.to_string())
}

/// Parse underline style
#[allow(dead_code)] // API for underline style parsing
pub fn parse_underline_style(val: Option<&str>) -> Option<String> {
    match val {
        Some("single") | Some("s") => Some("single".to_string()),
        Some("double") | Some("d") => Some("double".to_string()),
        Some("thick") | Some("th") => Some("thick".to_string()),
        Some("dotted") | Some("dot") => Some("dotted".to_string()),
        Some("dashed") | Some("dash") => Some("dashed".to_string()),
        Some("wave") | Some("w") => Some("wave".to_string()),
        _ => val.map(|s| s.to_string()),
    }
}

/// Parse line style for underline/overline
pub fn parse_line_style(val: Option<&str>) -> Option<LineStyle> {
    match val {
        Some("single") | Some("s") => Some(LineStyle::Single),
        Some("double") | Some("d") => Some(LineStyle::Double),
        Some("thick") | Some("th") => Some(LineStyle::Thick),
        Some("dotted") | Some("dot") => Some(LineStyle::Dotted),
        Some("dashed") | Some("dash") => Some(LineStyle::Dashed),
        Some("wave") | Some("w") => Some(LineStyle::Wave),
        _ => None,
    }
}

/// Parse strike style
#[allow(dead_code)] // API for strike-through style parsing
pub fn parse_strike_style(val: Option<&str>) -> Option<StrikeStyle> {
    match val {
        Some("single") | Some("s") => Some(StrikeStyle::Single),
        Some("double") | Some("d") => Some(StrikeStyle::Double),
        _ => None,
    }
}

/// Parse shape type
#[allow(dead_code)] // API for shape type parsing
pub fn parse_shape_type(val: Option<&str>) -> Option<ShapeType> {
    match val {
        Some("centered") | Some("cen") => Some(ShapeType::Centered),
        Some("match") | Some("mat") => Some(ShapeType::Match),
        _ => None,
    }
}

/// Parse break type
#[allow(dead_code)] // API for break type parsing
pub fn parse_break_type(val: Option<&str>) -> Option<BreakType> {
    match val {
        Some("line") | Some("ln") => Some(BreakType::Line),
        Some("page") | Some("pg") => Some(BreakType::Page),
        Some("none") | Some("no") => Some(BreakType::None),
        _ => None,
    }
}

/// Parse vertical alignment
#[allow(dead_code)] // API for vertical alignment parsing
pub fn parse_vertical_alignment(val: Option<&str>) -> Option<VerticalAlignment> {
    match val {
        Some("top") | Some("t") => Some(VerticalAlignment::Top),
        Some("bottom") | Some("bot") => Some(VerticalAlignment::Bottom),
        Some("center") | Some("cen") => Some(VerticalAlignment::Center),
        Some("baseline") | Some("base") => Some(VerticalAlignment::Baseline),
        Some("axis") | Some("ax") => Some(VerticalAlignment::Axis),
        _ => None,
    }
}

/// Parse position type
#[allow(dead_code)] // API for position type parsing
pub fn parse_position_type(val: Option<&str>) -> Option<Position> {
    match val {
        Some("pre") | Some("prefix") => Some(Position::Prefix),
        Some("post") | Some("postfix") => Some(Position::Postfix),
        Some("in") | Some("infix") => Some(Position::Infix),
        Some("top") => Some(Position::Top),
        Some("bot") | Some("bottom") => Some(Position::Bottom),
        _ => None,
    }
}

/// Parse fraction type
#[allow(dead_code)] // API for fraction type parsing
pub fn parse_fraction_type(val: Option<&str>) -> Option<FractionType> {
    match val {
        Some("bar") | Some("normal") => Some(FractionType::Bar),
        Some("noBar") | Some("linear") => Some(FractionType::NoBar),
        Some("skw") | Some("skewed") => Some(FractionType::Skewed),
        _ => None,
    }
}

/// Parse overline style
#[allow(dead_code)] // API for overline style parsing
pub fn parse_overline_style(val: Option<&str>) -> Option<String> {
    parse_underline_style(val) // Same as underline for overline
}

/// Fast attribute lookup without allocations
///
/// This uses direct linear search which is faster than HashMap for small attribute counts.
/// OMML elements typically have 2-5 attributes, so linear search is O(n) where n is small,
/// and avoids the overhead of HashMap allocation, hashing, and deallocation.
///
/// Performance characteristics:
/// - Zero heap allocations (stack-only)
/// - O(n) lookup where n is typically 2-5
/// - No HashMap overhead (allocation, hashing, drop)
/// - Cache-friendly (sequential access)
pub struct AttributeCache<'a> {
    attrs: &'a [quick_xml::events::attributes::Attribute<'a>],
}

impl<'a> AttributeCache<'a> {
    #[inline]
    pub fn new(attrs: &'a [quick_xml::events::attributes::Attribute]) -> Self {
        Self { attrs }
    }

    /// Get attribute value directly without caching
    ///
    /// For small attribute counts (typical in OMML), direct linear search is faster
    /// than HashMap lookup because it avoids allocation and hashing overhead.
    #[inline]
    pub fn get(&mut self, key: &str) -> Option<String> {
        get_attribute_value(self.attrs, key)
    }

    /// Get attribute as boolean
    #[allow(dead_code)] // Part of the AttributeCache API
    pub fn get_bool(&mut self, key: &str) -> Option<bool> {
        self.get(key).and_then(|s| parse_bool_fast(&s))
    }

    /// Get attribute as integer with SIMD acceleration
    #[allow(dead_code)] // Part of the AttributeCache API
    pub fn get_int(&mut self, key: &str) -> Option<i32> {
        self.get(key).and_then(|s| parse_int_simd(&s))
    }

    /// Get attribute as float with SIMD acceleration
    #[allow(dead_code)] // Part of the AttributeCache API
    pub fn get_float(&mut self, key: &str) -> Option<f32> {
        self.get(key).and_then(|s| parse_float_simd(&s))
    }

    /// Get attribute as space type
    #[allow(dead_code)] // Part of the AttributeCache API
    pub fn get_space(&mut self, key: &str) -> Option<SpaceType> {
        self.get(key).and_then(|s| parse_space_type(Some(&s)))
    }

    /// Get attribute as alignment
    #[allow(dead_code)] // Part of the AttributeCache API
    pub fn get_alignment(&mut self, key: &str) -> Option<Alignment> {
        self.get(key).and_then(|s| parse_alignment_value(&s))
    }

    /// Get attribute as vertical alignment
    #[allow(dead_code)] // Part of the AttributeCache API
    pub fn get_vertical_alignment(&mut self, key: &str) -> Option<VerticalAlignment> {
        self.get(key)
            .and_then(|s| parse_vertical_alignment(Some(&s)))
    }

    /// Get attribute as position type
    #[allow(dead_code)] // Part of the AttributeCache API
    pub fn get_position(&mut self, key: &str) -> Option<Position> {
        self.get(key).and_then(|s| parse_position_type(Some(&s)))
    }

    /// Get attribute as fraction type
    #[allow(dead_code)] // Part of the AttributeCache API
    pub fn get_fraction_type(&mut self, key: &str) -> Option<FractionType> {
        self.get(key).and_then(|s| parse_fraction_type(Some(&s)))
    }

    /// Get attribute as shape type
    #[allow(dead_code)] // Part of the AttributeCache API
    pub fn get_shape(&mut self, key: &str) -> Option<ShapeType> {
        self.get(key).and_then(|s| parse_shape_type(Some(&s)))
    }

    /// Get attribute as break type
    #[allow(dead_code)] // Part of the AttributeCache API
    pub fn get_break(&mut self, key: &str) -> Option<BreakType> {
        self.get(key).and_then(|s| parse_break_type(Some(&s)))
    }

    /// Get attribute as line style
    #[allow(dead_code)] // Part of the AttributeCache API
    pub fn get_line_style(&mut self, key: &str) -> Option<LineStyle> {
        self.get(key).and_then(|s| parse_line_style(Some(&s)))
    }

    /// Get attribute as strike style
    #[allow(dead_code)] // Part of the AttributeCache API
    pub fn get_strike_style(&mut self, key: &str) -> Option<StrikeStyle> {
        self.get(key).and_then(|s| parse_strike_style(Some(&s)))
    }
}

/// Batch attribute parsing for performance
#[allow(clippy::field_reassign_with_default)]
pub fn parse_attributes_batch(
    attrs: &[quick_xml::events::attributes::Attribute],
) -> ElementProperties {
    let mut cache = AttributeCache::new(attrs);
    let mut properties = ElementProperties::default();

    // Style and formatting
    properties.style = cache.get("scr");
    properties.math_variant = cache.get("scr").and_then(|s| parse_math_variant(Some(&s)));
    properties.display_style = cache.get("sty").and_then(|s| parse_display_style(Some(&s)));
    properties.script_level = cache
        .get("scrLvl")
        .and_then(|s| parse_script_level(Some(&s)));

    // Size and scaling
    properties.size = cache.get("sz");
    properties.min_size = cache.get("minSz");
    properties.max_size = cache.get("maxSz");

    // Color and font
    properties.color = cache.get("color");
    properties.font = cache.get("font");

    // Layout and positioning
    properties.alignment = cache.get("aln");
    properties.vertical_alignment = cache.get("vertJc");

    // Visibility and rendering
    properties.hide = cache.get_bool("hide");
    properties.strike_through = cache.get_bool("strike");
    properties.double_strike_through = cache.get_bool("dstrike");
    properties.underline = cache
        .get_line_style("u")
        .map(|s| format!("{:?}", s).to_lowercase());
    properties.overline = cache
        .get_line_style("o")
        .map(|s| format!("{:?}", s).to_lowercase());

    // Characters and symbols
    properties.chr = cache.get("chr");

    // Spacing
    properties.spacing = cache.get("val");

    // Fraction properties
    properties.fraction_type = cache.get("type");
    properties.fraction_line_thickness = cache.get("lnThick");

    // Matrix properties
    properties.matrix_alignment = cache.get("baseJc");
    properties.matrix_row_spacing = cache.get("rSp");
    properties.matrix_column_spacing = cache.get("cSp");

    // Accent properties
    properties.accent_position = cache.get("pos");

    // Box properties
    properties.box_alignment = cache.get("aln");
    properties.box_differential = cache.get_bool("diff");
    properties.box_operator_emulation = cache.get_bool("opEmu");
    properties.box_break = cache.get_bool("brk");
    properties.box_no_break = cache.get_bool("noBreak");

    // Phantom properties
    properties.phantom_show = cache.get_bool("show");
    properties.phantom_zero_width = cache.get_bool("zeroWid");
    properties.phantom_zero_ascent = cache.get_bool("zeroAsc");
    properties.phantom_zero_descent = cache.get_bool("zeroDesc");
    properties.phantom_transparent = cache.get_bool("transp");

    // Border box properties
    properties.border_hide_top = cache.get_bool("hideTop");
    properties.border_hide_bottom = cache.get_bool("hideBot");
    properties.border_hide_left = cache.get_bool("hideLeft");
    properties.border_hide_right = cache.get_bool("hideRight");
    properties.border_strike_horizontal = cache.get_bool("strikeH");
    properties.border_strike_vertical = cache.get_bool("strikeV");
    properties.border_strike_bltr = cache.get_bool("strikeBLTR");
    properties.border_strike_tlbr = cache.get_bool("strikeTLBR");

    // Equation array properties
    properties.eq_arr_base_alignment = cache.get("baseJc");
    properties.eq_arr_max_distance = cache.get("maxDist");
    properties.eq_arr_object_distance = cache.get("objDist");
    properties.eq_arr_row_spacing = cache.get("rSp");
    properties.eq_arr_row_spacing_rule = cache.get("rSpRule");

    // N-ary operator properties
    properties.nary_hide_sub = cache.get_bool("subHide");
    properties.nary_hide_sup = cache.get_bool("supHide");
    properties.nary_operator_grow = cache.get_bool("grow");

    // Delimiter properties
    properties.delimiter_grow = cache.get_bool("grow");
    properties.delimiter_shape = cache
        .get_shape("shp")
        .map(|s| format!("{:?}", s).to_lowercase());
    properties.delimiter_separator_char = cache.get("sepChr");
    properties.delimiter_open_char = cache.get("begChr");
    properties.delimiter_close_char = cache.get("endChr");

    // Radical properties
    properties.radical_hide_degree = cache.get_bool("degHide");

    // Run properties
    properties.run_literal = cache.get_bool("lit");
    properties.run_normal_text = cache.get("nor");
    properties.run_math_style = cache.get("sty");

    properties
}

/// Batch attribute parsing with caching for performance
#[allow(clippy::field_reassign_with_default)]
pub fn parse_attributes_batch_with_cache(cache: &mut AttributeCache) -> ElementProperties {
    let mut properties = ElementProperties::default();

    // Style and formatting
    properties.style = cache.get("scr");
    properties.math_variant = cache.get("scr").and_then(|s| parse_math_variant(Some(&s)));
    properties.display_style = cache.get("sty").and_then(|s| parse_display_style(Some(&s)));
    properties.script_level = cache
        .get("scrLvl")
        .and_then(|s| parse_script_level(Some(&s)));

    // Size and scaling
    properties.size = cache.get("sz");
    properties.min_size = cache.get("minSz");
    properties.max_size = cache.get("maxSz");

    // Color and font
    properties.color = cache.get("color");
    properties.font = cache.get("font");

    // Layout and positioning
    properties.alignment = cache.get("aln");
    properties.vertical_alignment = cache.get("vertJc");

    // Visibility and rendering
    properties.hide = cache.get_bool("hide");
    properties.strike_through = cache.get_bool("strike");
    properties.double_strike_through = cache.get_bool("dstrike");
    properties.underline = cache
        .get_line_style("u")
        .map(|s| format!("{:?}", s).to_lowercase());
    properties.overline = cache
        .get_line_style("o")
        .map(|s| format!("{:?}", s).to_lowercase());

    // Characters and symbols
    properties.chr = cache.get("chr");

    // Spacing
    properties.spacing = cache.get("val");

    // Fraction properties
    properties.fraction_type = cache.get("type");
    properties.fraction_line_thickness = cache.get("lnThick");

    // Matrix properties
    properties.matrix_alignment = cache.get("baseJc");
    properties.matrix_row_spacing = cache.get("rSp");
    properties.matrix_column_spacing = cache.get("cSp");

    // Accent properties
    properties.accent_position = cache.get("pos");

    // Box properties
    properties.box_alignment = cache.get("aln");
    properties.box_differential = cache.get_bool("diff");
    properties.box_operator_emulation = cache.get_bool("opEmu");
    properties.box_break = cache.get_bool("brk");
    properties.box_no_break = cache.get_bool("noBreak");

    // Phantom properties
    properties.phantom_show = cache.get_bool("show");
    properties.phantom_zero_width = cache.get_bool("zeroWid");
    properties.phantom_zero_ascent = cache.get_bool("zeroAsc");
    properties.phantom_zero_descent = cache.get_bool("zeroDesc");
    properties.phantom_transparent = cache.get_bool("transp");

    // Border box properties
    properties.border_hide_top = cache.get_bool("hideTop");
    properties.border_hide_bottom = cache.get_bool("hideBot");
    properties.border_hide_left = cache.get_bool("hideLeft");
    properties.border_hide_right = cache.get_bool("hideRight");
    properties.border_strike_horizontal = cache.get_bool("strikeH");
    properties.border_strike_vertical = cache.get_bool("strikeV");
    properties.border_strike_bltr = cache.get_bool("strikeBLTR");
    properties.border_strike_tlbr = cache.get_bool("strikeTLBR");

    // Equation array properties
    properties.eq_arr_base_alignment = cache.get("baseJc");
    properties.eq_arr_max_distance = cache.get("maxDist");
    properties.eq_arr_object_distance = cache.get("objDist");
    properties.eq_arr_row_spacing = cache.get("rSp");
    properties.eq_arr_row_spacing_rule = cache.get("rSpRule");

    // N-ary operator properties
    properties.nary_hide_sub = cache.get_bool("subHide");
    properties.nary_hide_sup = cache.get_bool("supHide");
    properties.nary_operator_grow = cache.get_bool("grow");

    // Delimiter properties
    properties.delimiter_grow = cache.get_bool("grow");
    properties.delimiter_shape = cache
        .get_shape("shp")
        .map(|s| format!("{:?}", s).to_lowercase());
    properties.delimiter_separator_char = cache.get("sepChr");
    properties.delimiter_open_char = cache.get("begChr");
    properties.delimiter_close_char = cache.get("endChr");

    // Radical properties
    properties.radical_hide_degree = cache.get_bool("degHide");

    // Run properties
    properties.run_literal = cache.get_bool("lit");
    properties.run_normal_text = cache.get("nor");
    properties.run_math_style = cache.get("sty");

    properties
}
