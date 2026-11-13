use crate::ast::{LineStyle, Position, ShapeType, VerticalAlignment, *};
use crate::omml::elements::ElementProperties;
use crate::omml::lookup::*;

/// SIMD-accelerated numeric parsing functions
/// Fast integer parsing using atoi_simd
pub fn parse_int_simd(value: &str) -> Option<i32> {
    atoi_simd::parse::<_, false, false>(value.as_bytes()).ok()
}

/// Fast float parsing using fast_float2
pub fn parse_float_simd(value: &str) -> Option<f32> {
    fast_float2::parse(value).ok()
}

/// Fast boolean parsing with lookup
pub fn parse_bool_fast(value: &str) -> Option<bool> {
    parse_bool_value(value)
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

/// Extract attribute value as string (optimized for performance)
///
/// This function is called extensively during OMML parsing, so it's heavily optimized:
/// Extract attribute value with zero-copy optimization (borrowed version)
///
/// Returns `Cow<str>` to avoid unnecessary allocations. The returned value
/// borrows from the XML attribute when possible.
///
/// # Performance
/// - Zero-copy in most cases (no `_platform_memmove`)
/// - Uses stack-allocated array instead of SmallVec to avoid memcpy
/// - Uses byte-level comparison to avoid UTF-8 validation overhead
/// - Optimized for OMML keys which are typically short (<30 bytes)
///
/// Use this version when you don't need to store the value long-term.
#[inline]
pub fn get_attribute_value_borrowed<'a>(
    attrs: &'a [quick_xml::events::attributes::Attribute<'a>],
    key: &str,
) -> Option<std::borrow::Cow<'a, str>> {
    let key_bytes = key.as_bytes();
    let key_len = key_bytes.len();

    // Optimize: Use stack array instead of SmallVec to avoid memcpy
    // Most OMML keys are short (e.g., "val", "chr", "pos") so 32 bytes is plenty
    debug_assert!(key_len <= 30, "OMML attribute key too long: {}", key);

    if key_len > 30 {
        // Fallback for unexpectedly long keys (should never happen in practice)
        return None;
    }

    // Pre-compute the m: prefixed key on the stack (zero heap allocation)
    let mut m_prefixed_buf = [0u8; 32]; // "m:" + key (max 30 bytes)
    m_prefixed_buf[0] = b'm';
    m_prefixed_buf[1] = b':';
    m_prefixed_buf[2..2 + key_len].copy_from_slice(key_bytes);
    let m_prefixed_len = 2 + key_len;
    let m_prefixed = &m_prefixed_buf[..m_prefixed_len];

    // Fast path: iterate attributes with byte-level comparison
    // Typically only 2-5 attributes per element, so linear search is optimal
    for attr in attrs {
        let attr_key = attr.key.as_ref();

        // Fast byte comparison (avoids UTF-8 validation overhead)
        if attr_key == key_bytes || attr_key == m_prefixed {
            // Only decode UTF-8 when we have a match (borrowing from attr.value)
            return Some(std::borrow::Cow::Borrowed(
                std::str::from_utf8(&attr.value).ok()?,
            ));
        }
    }
    None
}

/// Extract attribute value as owned String
///
/// This is a convenience wrapper around `get_attribute_value_borrowed` that
/// converts the result to an owned String for storage in structures.
///
/// For better performance when the value doesn't need to be stored,
/// use `get_attribute_value_borrowed` instead.
#[inline]
pub fn get_attribute_value(
    attrs: &[quick_xml::events::attributes::Attribute],
    key: &str,
) -> Option<String> {
    get_attribute_value_borrowed(attrs, key).map(|cow| cow.into_owned())
}

/// Extract attribute value as integer with SIMD acceleration
pub fn get_attribute_value_int(
    attrs: &[quick_xml::events::attributes::Attribute],
    key: &str,
) -> Option<i32> {
    get_attribute_value_borrowed(attrs, key).and_then(|s| parse_int_simd(&s))
}

/// Extract attribute value as float with SIMD acceleration
pub fn get_attribute_value_float(
    attrs: &[quick_xml::events::attributes::Attribute],
    key: &str,
) -> Option<f32> {
    get_attribute_value_borrowed(attrs, key).and_then(|s| parse_float_simd(&s))
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

/// Parse shape type
pub fn parse_shape_type(val: Option<&str>) -> Option<ShapeType> {
    match val {
        Some("centered") | Some("cen") => Some(ShapeType::Centered),
        Some("match") | Some("mat") => Some(ShapeType::Match),
        _ => None,
    }
}

/// Parse vertical alignment
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

    /// Get attribute as borrowed Cow (zero-copy when possible)
    ///
    /// Performance-optimized version that borrows from the XML attribute
    /// instead of allocating. Use this for immediate parsing operations.
    #[inline]
    pub fn get_borrowed(&self, key: &str) -> Option<std::borrow::Cow<'a, str>> {
        get_attribute_value_borrowed(self.attrs, key)
    }

    /// Get attribute as boolean (optimized: zero-copy)
    ///
    /// Uses borrowed string to avoid allocation overhead.
    /// This is called very frequently during OMML parsing.
    #[inline]
    pub fn get_bool(&mut self, key: &str) -> Option<bool> {
        // Optimize: Use borrowed version to avoid String allocation
        self.get_borrowed(key).and_then(|s| parse_bool_fast(&s))
    }

    /// Get attribute as shape type (optimized: zero-copy)
    #[inline]
    pub fn get_shape(&mut self, key: &str) -> Option<ShapeType> {
        self.get_borrowed(key)
            .and_then(|s| parse_shape_type(Some(&s)))
    }

    /// Get attribute as line style (optimized: zero-copy)
    #[inline]
    pub fn get_line_style(&mut self, key: &str) -> Option<LineStyle> {
        self.get_borrowed(key)
            .and_then(|s| parse_line_style(Some(&s)))
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
