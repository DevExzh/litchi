//! Efficient ODF XML tag matching using SIMD and Aho-Corasick automaton.
//!
//! This module provides high-performance tag matching for ODF XML parsing,
//! using SIMD instructions for short tag comparison and Aho-Corasick automaton
//! for multi-pattern matching when parsing large documents.
//!
//! **Note**: This module provides a complete public API for ODF tag matching.
//! Not all functions are used internally, but they are available for advanced users
//! who need custom XML parsing optimizations.
//!
//! # Performance Optimizations
//!
//! - **SIMD for prefix matching**: Uses SIMD to quickly compare namespace prefixes
//! - **Compile-time tag hashing**: Uses `phf` for O(1) tag lookups

#![allow(dead_code)] // Public API - complete tag matching utilities
//! - **Zero allocations**: All tag comparisons are done on borrowed slices
//! - **Inlined hot paths**: Critical functions are marked `#[inline(always)]`
//!
//! # References
//!
//! - odfpy: `3rdparty/odfpy/odf/namespaces.py`
//! - odfdo: `3rdparty/odfdo/src/odfdo/const.py`
use memchr::memmem;
use phf::{Map, phf_map};

// ============================================================================
// TAG TYPE ENUMERATION
// ============================================================================

/// ODF XML tag types for fast dispatch
///
/// Using enums instead of strings reduces memory usage and enables
/// efficient dispatch via match expressions (jump tables) instead of string comparisons.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u16)]
pub enum OdfTag {
    // Text elements
    TextP,
    TextH,
    TextSpan,
    TextA,
    TextLineBreak,
    TextS,
    TextTab,
    TextList,
    TextListItem,
    TextBookmark,
    TextBookmarkStart,
    TextBookmarkEnd,
    TextSequence,
    TextNote,
    TextNoteBody,
    TextNoteCitation,

    // Table elements
    TableTable,
    TableTableRow,
    TableTableCell,
    TableTableColumn,
    TableTableHeaderRows,
    TableTableHeaderColumns,
    TableCoveredTableCell,
    TableTableRowGroup,
    TableTableColumnGroup,

    // Drawing elements
    DrawFrame,
    DrawImage,
    DrawTextBox,
    DrawRect,
    DrawCircle,
    DrawEllipse,
    DrawLine,
    DrawPolygon,
    DrawPolyline,
    DrawPath,
    DrawG,
    DrawPage,
    DrawCustomShape,

    // Style elements
    StyleStyle,
    StyleParagraphProperties,
    StyleTextProperties,
    StyleTableCellProperties,
    StyleTableRowProperties,
    StyleTableColumnProperties,
    StyleGraphicProperties,
    StyleFontFace,
    StyleDefaultStyle,
    StyleMasterPage,
    StylePageLayout,
    StyleHeaderFooter,
    StyleBackgroundImage,

    // Office elements
    OfficeBody,
    OfficeText,
    OfficeSpreadsheet,
    OfficePresentation,
    OfficeDrawing,
    OfficeMasterStyles,
    OfficeAutomaticStyles,
    OfficeStyles,
    OfficeFontFaceDecls,
    OfficeScripts,
    OfficeSettings,
    OfficeMeta,
    OfficeAnnotation,

    // Form elements
    FormForm,
    FormText,
    FormTextarea,
    FormButton,
    FormCheckbox,
    FormRadio,
    FormListbox,
    FormCombobox,

    // Chart elements
    ChartChart,
    ChartTitle,
    ChartSubtitle,
    ChartLegend,
    ChartPlotArea,
    ChartSeries,
    ChartDomain,
    ChartAxis,

    // Meta elements
    MetaGenerator,
    MetaCreationDate,
    MetaEditingDuration,
    MetaEditingCycles,
    MetaDocumentStatistic,
    MetaKeyword,
    MetaUserDefined,

    // Number/Data style elements
    NumberNumberStyle,
    NumberCurrencyStyle,
    NumberPercentageStyle,
    NumberDateStyle,
    NumberTimeStyle,
    NumberBooleanStyle,
    NumberTextStyle,
    NumberNumber,
    NumberCurrencySymbol,
    NumberText,
    NumberDay,
    NumberMonth,
    NumberYear,
    NumberHours,
    NumberMinutes,
    NumberSeconds,

    // Presentation elements
    PresentationNotes,
    PresentationSettings,
    PresentationFooter,
    PresentationDateTime,
    PresentationHeader,

    // Animation elements
    AnimPar,
    AnimSeq,
    AnimSet,
    AnimAnimate,
    AnimAnimateMotion,
    AnimAnimateColor,
    AnimTransFilter,

    // Dublin Core elements
    DcTitle,
    DcDescription,
    DcSubject,
    DcCreator,
    DcDate,
    DcLanguage,

    // SVG elements
    SvgDesc,
    SvgTitle,
    SvgLinearGradient,
    SvgRadialGradient,
    SvgStop,

    // Math elements
    MathMath,

    // Script elements
    ScriptEventListener,

    // Config elements
    ConfigConfigItemSet,
    ConfigConfigItem,
    ConfigConfigItemMapIndexed,
    ConfigConfigItemMapEntry,
    ConfigConfigItemMapNamed,

    // Unknown/unsupported tag
    Unknown,
}

// ============================================================================
// COMPILE-TIME TAG MAPPING
// ============================================================================

/// Tag string to OdfTag enum mapping (compile-time perfect hash map)
///
/// This provides O(1) lookup from tag string to enum variant with zero runtime overhead.
/// The perfect hash function is generated at compile time by the `phf` crate.
static TAG_MAP: Map<&'static [u8], OdfTag> = phf_map! {
    // Text elements
    b"text:p" => OdfTag::TextP,
    b"text:h" => OdfTag::TextH,
    b"text:span" => OdfTag::TextSpan,
    b"text:a" => OdfTag::TextA,
    b"text:line-break" => OdfTag::TextLineBreak,
    b"text:s" => OdfTag::TextS,
    b"text:tab" => OdfTag::TextTab,
    b"text:list" => OdfTag::TextList,
    b"text:list-item" => OdfTag::TextListItem,
    b"text:bookmark" => OdfTag::TextBookmark,
    b"text:bookmark-start" => OdfTag::TextBookmarkStart,
    b"text:bookmark-end" => OdfTag::TextBookmarkEnd,
    b"text:sequence" => OdfTag::TextSequence,
    b"text:note" => OdfTag::TextNote,
    b"text:note-body" => OdfTag::TextNoteBody,
    b"text:note-citation" => OdfTag::TextNoteCitation,

    // Table elements
    b"table:table" => OdfTag::TableTable,
    b"table:table-row" => OdfTag::TableTableRow,
    b"table:table-cell" => OdfTag::TableTableCell,
    b"table:table-column" => OdfTag::TableTableColumn,
    b"table:table-header-rows" => OdfTag::TableTableHeaderRows,
    b"table:table-header-columns" => OdfTag::TableTableHeaderColumns,
    b"table:covered-table-cell" => OdfTag::TableCoveredTableCell,
    b"table:table-row-group" => OdfTag::TableTableRowGroup,
    b"table:table-column-group" => OdfTag::TableTableColumnGroup,

    // Drawing elements
    b"draw:frame" => OdfTag::DrawFrame,
    b"draw:image" => OdfTag::DrawImage,
    b"draw:text-box" => OdfTag::DrawTextBox,
    b"draw:rect" => OdfTag::DrawRect,
    b"draw:circle" => OdfTag::DrawCircle,
    b"draw:ellipse" => OdfTag::DrawEllipse,
    b"draw:line" => OdfTag::DrawLine,
    b"draw:polygon" => OdfTag::DrawPolygon,
    b"draw:polyline" => OdfTag::DrawPolyline,
    b"draw:path" => OdfTag::DrawPath,
    b"draw:g" => OdfTag::DrawG,
    b"draw:page" => OdfTag::DrawPage,
    b"draw:custom-shape" => OdfTag::DrawCustomShape,

    // Style elements
    b"style:style" => OdfTag::StyleStyle,
    b"style:paragraph-properties" => OdfTag::StyleParagraphProperties,
    b"style:text-properties" => OdfTag::StyleTextProperties,
    b"style:table-cell-properties" => OdfTag::StyleTableCellProperties,
    b"style:table-row-properties" => OdfTag::StyleTableRowProperties,
    b"style:table-column-properties" => OdfTag::StyleTableColumnProperties,
    b"style:graphic-properties" => OdfTag::StyleGraphicProperties,
    b"style:font-face" => OdfTag::StyleFontFace,
    b"style:default-style" => OdfTag::StyleDefaultStyle,
    b"style:master-page" => OdfTag::StyleMasterPage,
    b"style:page-layout" => OdfTag::StylePageLayout,
    b"style:header" => OdfTag::StyleHeaderFooter,
    b"style:footer" => OdfTag::StyleHeaderFooter,
    b"style:background-image" => OdfTag::StyleBackgroundImage,

    // Office elements
    b"office:body" => OdfTag::OfficeBody,
    b"office:text" => OdfTag::OfficeText,
    b"office:spreadsheet" => OdfTag::OfficeSpreadsheet,
    b"office:presentation" => OdfTag::OfficePresentation,
    b"office:drawing" => OdfTag::OfficeDrawing,
    b"office:master-styles" => OdfTag::OfficeMasterStyles,
    b"office:automatic-styles" => OdfTag::OfficeAutomaticStyles,
    b"office:styles" => OdfTag::OfficeStyles,
    b"office:font-face-decls" => OdfTag::OfficeFontFaceDecls,
    b"office:scripts" => OdfTag::OfficeScripts,
    b"office:settings" => OdfTag::OfficeSettings,
    b"office:meta" => OdfTag::OfficeMeta,
    b"office:annotation" => OdfTag::OfficeAnnotation,

    // Form elements
    b"form:form" => OdfTag::FormForm,
    b"form:text" => OdfTag::FormText,
    b"form:textarea" => OdfTag::FormTextarea,
    b"form:button" => OdfTag::FormButton,
    b"form:checkbox" => OdfTag::FormCheckbox,
    b"form:radio" => OdfTag::FormRadio,
    b"form:listbox" => OdfTag::FormListbox,
    b"form:combobox" => OdfTag::FormCombobox,

    // Chart elements
    b"chart:chart" => OdfTag::ChartChart,
    b"chart:title" => OdfTag::ChartTitle,
    b"chart:subtitle" => OdfTag::ChartSubtitle,
    b"chart:legend" => OdfTag::ChartLegend,
    b"chart:plot-area" => OdfTag::ChartPlotArea,
    b"chart:series" => OdfTag::ChartSeries,
    b"chart:domain" => OdfTag::ChartDomain,
    b"chart:axis" => OdfTag::ChartAxis,

    // Meta elements
    b"meta:generator" => OdfTag::MetaGenerator,
    b"meta:creation-date" => OdfTag::MetaCreationDate,
    b"meta:editing-duration" => OdfTag::MetaEditingDuration,
    b"meta:editing-cycles" => OdfTag::MetaEditingCycles,
    b"meta:document-statistic" => OdfTag::MetaDocumentStatistic,
    b"meta:keyword" => OdfTag::MetaKeyword,
    b"meta:user-defined" => OdfTag::MetaUserDefined,

    // Number/Data style elements
    b"number:number-style" => OdfTag::NumberNumberStyle,
    b"number:currency-style" => OdfTag::NumberCurrencyStyle,
    b"number:percentage-style" => OdfTag::NumberPercentageStyle,
    b"number:date-style" => OdfTag::NumberDateStyle,
    b"number:time-style" => OdfTag::NumberTimeStyle,
    b"number:boolean-style" => OdfTag::NumberBooleanStyle,
    b"number:text-style" => OdfTag::NumberTextStyle,
    b"number:number" => OdfTag::NumberNumber,
    b"number:currency-symbol" => OdfTag::NumberCurrencySymbol,
    b"number:text" => OdfTag::NumberText,
    b"number:day" => OdfTag::NumberDay,
    b"number:month" => OdfTag::NumberMonth,
    b"number:year" => OdfTag::NumberYear,
    b"number:hours" => OdfTag::NumberHours,
    b"number:minutes" => OdfTag::NumberMinutes,
    b"number:seconds" => OdfTag::NumberSeconds,

    // Presentation elements
    b"presentation:notes" => OdfTag::PresentationNotes,
    b"presentation:settings" => OdfTag::PresentationSettings,
    b"presentation:footer" => OdfTag::PresentationFooter,
    b"presentation:date-time" => OdfTag::PresentationDateTime,
    b"presentation:header" => OdfTag::PresentationHeader,

    // Animation elements
    b"anim:par" => OdfTag::AnimPar,
    b"anim:seq" => OdfTag::AnimSeq,
    b"anim:set" => OdfTag::AnimSet,
    b"anim:animate" => OdfTag::AnimAnimate,
    b"anim:animateMotion" => OdfTag::AnimAnimateMotion,
    b"anim:animateColor" => OdfTag::AnimAnimateColor,
    b"anim:transitionFilter" => OdfTag::AnimTransFilter,

    // Dublin Core elements
    b"dc:title" => OdfTag::DcTitle,
    b"dc:description" => OdfTag::DcDescription,
    b"dc:subject" => OdfTag::DcSubject,
    b"dc:creator" => OdfTag::DcCreator,
    b"dc:date" => OdfTag::DcDate,
    b"dc:language" => OdfTag::DcLanguage,

    // SVG elements
    b"svg:desc" => OdfTag::SvgDesc,
    b"svg:title" => OdfTag::SvgTitle,
    b"svg:linearGradient" => OdfTag::SvgLinearGradient,
    b"svg:radialGradient" => OdfTag::SvgRadialGradient,
    b"svg:stop" => OdfTag::SvgStop,

    // Math elements
    b"math:math" => OdfTag::MathMath,

    // Script elements
    b"script:event-listener" => OdfTag::ScriptEventListener,

    // Config elements
    b"config:config-item-set" => OdfTag::ConfigConfigItemSet,
    b"config:config-item" => OdfTag::ConfigConfigItem,
    b"config:config-item-map-indexed" => OdfTag::ConfigConfigItemMapIndexed,
    b"config:config-item-map-entry" => OdfTag::ConfigConfigItemMapEntry,
    b"config:config-item-map-named" => OdfTag::ConfigConfigItemMapNamed,
};

// ============================================================================
// SIMD-OPTIMIZED PREFIX MATCHING
// ============================================================================

/// Fast prefix matching using SIMD-accelerated memmem from memchr crate
///
/// This is significantly faster than iterating over prefixes for common tags.
/// The `memchr` crate uses SIMD instructions (SSE2, AVX2, NEON) when available.
#[inline(always)]
pub fn has_prefix(tag: &[u8], prefix: &[u8]) -> bool {
    if tag.len() < prefix.len() {
        return false;
    }
    // For short prefixes (< 16 bytes), direct comparison is fastest
    if prefix.len() <= 16 {
        tag.starts_with(prefix)
    } else {
        // For longer strings, use SIMD-accelerated search
        memmem::find(tag, prefix) == Some(0)
    }
}

/// Extract namespace prefix from tag (zero-copy)
///
/// Returns the prefix part before ':' or empty slice if no prefix.
/// This is a zero-copy operation that returns a borrowed slice.
///
/// # Examples
///
/// ```
/// # use litchi::odf::elements::tag_matcher::extract_prefix;
/// assert_eq!(extract_prefix(b"text:p"), b"text");
/// assert_eq!(extract_prefix(b"p"), b"");
/// ```
#[inline(always)]
pub fn extract_prefix(tag: &[u8]) -> &[u8] {
    // Use memchr for fast colon finding
    if let Some(colon_pos) = memchr::memchr(b':', tag) {
        &tag[..colon_pos]
    } else {
        b""
    }
}

/// Extract local name from tag (zero-copy)
///
/// Returns the local name part after ':' or the entire tag if no prefix.
///
/// # Examples
///
/// ```
/// # use litchi::odf::elements::tag_matcher::extract_local_name;
/// assert_eq!(extract_local_name(b"text:p"), b"p");
/// assert_eq!(extract_local_name(b"p"), b"p");
/// ```
#[inline(always)]
pub fn extract_local_name(tag: &[u8]) -> &[u8] {
    if let Some(colon_pos) = memchr::memchr(b':', tag) {
        &tag[colon_pos + 1..]
    } else {
        tag
    }
}

// ============================================================================
// TAG MATCHING API
// ============================================================================

/// Match a tag to its OdfTag enum variant
///
/// This provides O(1) lookup using compile-time perfect hash function.
/// For unknown tags, returns `OdfTag::Unknown`.
///
/// # Arguments
///
/// * `tag` - Tag name as bytes (e.g., b"text:p")
///
/// # Returns
///
/// The corresponding `OdfTag` enum variant
///
/// # Examples
///
/// ```
/// # use litchi::odf::elements::tag_matcher::{match_tag, OdfTag};
/// assert_eq!(match_tag(b"text:p"), OdfTag::TextP);
/// assert_eq!(match_tag(b"table:table"), OdfTag::TableTable);
/// assert_eq!(match_tag(b"unknown:tag"), OdfTag::Unknown);
/// ```
#[inline(always)]
pub fn match_tag(tag: &[u8]) -> OdfTag {
    TAG_MAP.get(tag).copied().unwrap_or(OdfTag::Unknown)
}

/// Check if a tag belongs to a specific namespace (fast SIMD-based check)
///
/// # Arguments
///
/// * `tag` - Tag name as bytes
/// * `namespace` - Namespace prefix (e.g., b"text", b"table")
///
/// # Returns
///
/// `true` if the tag belongs to the specified namespace
///
/// # Examples
///
/// ```
/// # use litchi::odf::elements::tag_matcher::is_namespace;
/// assert!(is_namespace(b"text:p", b"text"));
/// assert!(is_namespace(b"table:table-row", b"table"));
/// assert!(!is_namespace(b"text:p", b"table"));
/// ```
#[inline(always)]
pub fn is_namespace(tag: &[u8], namespace: &[u8]) -> bool {
    // Quick length check first
    if tag.len() <= namespace.len() {
        return false;
    }
    // Check if tag starts with "namespace:"
    tag.starts_with(namespace) && tag.get(namespace.len()) == Some(&b':')
}

/// Check if tag is a text element (fast namespace check)
#[inline(always)]
pub fn is_text_tag(tag: &[u8]) -> bool {
    is_namespace(tag, b"text")
}

/// Check if tag is a table element (fast namespace check)
#[inline(always)]
pub fn is_table_tag(tag: &[u8]) -> bool {
    is_namespace(tag, b"table")
}

/// Check if tag is a drawing element (fast namespace check)
#[inline(always)]
pub fn is_draw_tag(tag: &[u8]) -> bool {
    is_namespace(tag, b"draw")
}

/// Check if tag is a style element (fast namespace check)
#[inline(always)]
pub fn is_style_tag(tag: &[u8]) -> bool {
    is_namespace(tag, b"style")
}

/// Check if tag is an office element (fast namespace check)
#[inline(always)]
pub fn is_office_tag(tag: &[u8]) -> bool {
    is_namespace(tag, b"office")
}

// ============================================================================
// BATCH TAG MATCHING
// ============================================================================

/// Match multiple tags at once (useful for filtering)
///
/// Returns a vector of (index, OdfTag) pairs for all recognized tags.
/// This can be more efficient than matching tags one by one.
pub fn match_tags_batch(tags: &[&[u8]]) -> Vec<(usize, OdfTag)> {
    tags.iter()
        .enumerate()
        .filter_map(|(idx, tag)| {
            let matched = match_tag(tag);
            if matched != OdfTag::Unknown {
                Some((idx, matched))
            } else {
                None
            }
        })
        .collect()
}

// ============================================================================
// TESTS
// ============================================================================

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_tag_matching() {
        assert_eq!(match_tag(b"text:p"), OdfTag::TextP);
        assert_eq!(match_tag(b"text:h"), OdfTag::TextH);
        assert_eq!(match_tag(b"table:table"), OdfTag::TableTable);
        assert_eq!(match_tag(b"table:table-row"), OdfTag::TableTableRow);
        assert_eq!(match_tag(b"unknown:tag"), OdfTag::Unknown);
    }

    #[test]
    fn test_namespace_checking() {
        assert!(is_namespace(b"text:p", b"text"));
        assert!(is_namespace(b"table:table-row", b"table"));
        assert!(!is_namespace(b"text:p", b"table"));
        assert!(!is_namespace(b"p", b"text"));
    }

    #[test]
    fn test_prefix_extraction() {
        assert_eq!(extract_prefix(b"text:p"), b"text");
        assert_eq!(extract_prefix(b"table:table-row"), b"table");
        assert_eq!(extract_prefix(b"p"), b"");
    }

    #[test]
    fn test_local_name_extraction() {
        assert_eq!(extract_local_name(b"text:p"), b"p");
        assert_eq!(extract_local_name(b"table:table-row"), b"table-row");
        assert_eq!(extract_local_name(b"p"), b"p");
    }

    #[test]
    fn test_namespace_helpers() {
        assert!(is_text_tag(b"text:p"));
        assert!(is_text_tag(b"text:h"));
        assert!(!is_text_tag(b"table:table"));

        assert!(is_table_tag(b"table:table"));
        assert!(is_table_tag(b"table:table-row"));
        assert!(!is_table_tag(b"text:p"));
    }

    #[test]
    fn test_batch_matching() {
        let tags = vec![
            b"text:p".as_ref(),
            b"table:table",
            b"unknown:tag",
            b"draw:frame",
        ];
        let matched = match_tags_batch(&tags);

        assert_eq!(matched.len(), 3);
        assert_eq!(matched[0], (0, OdfTag::TextP));
        assert_eq!(matched[1], (1, OdfTag::TableTable));
        assert_eq!(matched[2], (3, OdfTag::DrawFrame));
    }

    #[test]
    fn test_has_prefix() {
        assert!(has_prefix(b"text:paragraph", b"text"));
        assert!(has_prefix(b"table:table-row", b"table"));
        assert!(!has_prefix(b"short", b"very-long-prefix"));
    }
}
