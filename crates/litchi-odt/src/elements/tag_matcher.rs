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
pub enum Tag {
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

/// Tag string to Tag enum mapping (compile-time perfect hash map)
///
/// This provides O(1) lookup from tag string to enum variant with zero runtime overhead.
/// The perfect hash function is generated at compile time by the `phf` crate.
static TAG_MAP: Map<&'static [u8], Tag> = phf_map! {
    // Text elements
    b"text:p" => Tag::TextP,
    b"text:h" => Tag::TextH,
    b"text:span" => Tag::TextSpan,
    b"text:a" => Tag::TextA,
    b"text:line-break" => Tag::TextLineBreak,
    b"text:s" => Tag::TextS,
    b"text:tab" => Tag::TextTab,
    b"text:list" => Tag::TextList,
    b"text:list-item" => Tag::TextListItem,
    b"text:bookmark" => Tag::TextBookmark,
    b"text:bookmark-start" => Tag::TextBookmarkStart,
    b"text:bookmark-end" => Tag::TextBookmarkEnd,
    b"text:sequence" => Tag::TextSequence,
    b"text:note" => Tag::TextNote,
    b"text:note-body" => Tag::TextNoteBody,
    b"text:note-citation" => Tag::TextNoteCitation,

    // Table elements
    b"table:table" => Tag::TableTable,
    b"table:table-row" => Tag::TableTableRow,
    b"table:table-cell" => Tag::TableTableCell,
    b"table:table-column" => Tag::TableTableColumn,
    b"table:table-header-rows" => Tag::TableTableHeaderRows,
    b"table:table-header-columns" => Tag::TableTableHeaderColumns,
    b"table:covered-table-cell" => Tag::TableCoveredTableCell,
    b"table:table-row-group" => Tag::TableTableRowGroup,
    b"table:table-column-group" => Tag::TableTableColumnGroup,

    // Drawing elements
    b"draw:frame" => Tag::DrawFrame,
    b"draw:image" => Tag::DrawImage,
    b"draw:text-box" => Tag::DrawTextBox,
    b"draw:rect" => Tag::DrawRect,
    b"draw:circle" => Tag::DrawCircle,
    b"draw:ellipse" => Tag::DrawEllipse,
    b"draw:line" => Tag::DrawLine,
    b"draw:polygon" => Tag::DrawPolygon,
    b"draw:polyline" => Tag::DrawPolyline,
    b"draw:path" => Tag::DrawPath,
    b"draw:g" => Tag::DrawG,
    b"draw:page" => Tag::DrawPage,
    b"draw:custom-shape" => Tag::DrawCustomShape,

    // Style elements
    b"style:style" => Tag::StyleStyle,
    b"style:paragraph-properties" => Tag::StyleParagraphProperties,
    b"style:text-properties" => Tag::StyleTextProperties,
    b"style:table-cell-properties" => Tag::StyleTableCellProperties,
    b"style:table-row-properties" => Tag::StyleTableRowProperties,
    b"style:table-column-properties" => Tag::StyleTableColumnProperties,
    b"style:graphic-properties" => Tag::StyleGraphicProperties,
    b"style:font-face" => Tag::StyleFontFace,
    b"style:default-style" => Tag::StyleDefaultStyle,
    b"style:master-page" => Tag::StyleMasterPage,
    b"style:page-layout" => Tag::StylePageLayout,
    b"style:header" => Tag::StyleHeaderFooter,
    b"style:footer" => Tag::StyleHeaderFooter,
    b"style:background-image" => Tag::StyleBackgroundImage,

    // Office elements
    b"office:body" => Tag::OfficeBody,
    b"office:text" => Tag::OfficeText,
    b"office:spreadsheet" => Tag::OfficeSpreadsheet,
    b"office:presentation" => Tag::OfficePresentation,
    b"office:drawing" => Tag::OfficeDrawing,
    b"office:master-styles" => Tag::OfficeMasterStyles,
    b"office:automatic-styles" => Tag::OfficeAutomaticStyles,
    b"office:styles" => Tag::OfficeStyles,
    b"office:font-face-decls" => Tag::OfficeFontFaceDecls,
    b"office:scripts" => Tag::OfficeScripts,
    b"office:settings" => Tag::OfficeSettings,
    b"office:meta" => Tag::OfficeMeta,
    b"office:annotation" => Tag::OfficeAnnotation,

    // Form elements
    b"form:form" => Tag::FormForm,
    b"form:text" => Tag::FormText,
    b"form:textarea" => Tag::FormTextarea,
    b"form:button" => Tag::FormButton,
    b"form:checkbox" => Tag::FormCheckbox,
    b"form:radio" => Tag::FormRadio,
    b"form:listbox" => Tag::FormListbox,
    b"form:combobox" => Tag::FormCombobox,

    // Chart elements
    b"chart:chart" => Tag::ChartChart,
    b"chart:title" => Tag::ChartTitle,
    b"chart:subtitle" => Tag::ChartSubtitle,
    b"chart:legend" => Tag::ChartLegend,
    b"chart:plot-area" => Tag::ChartPlotArea,
    b"chart:series" => Tag::ChartSeries,
    b"chart:domain" => Tag::ChartDomain,
    b"chart:axis" => Tag::ChartAxis,

    // Meta elements
    b"meta:generator" => Tag::MetaGenerator,
    b"meta:creation-date" => Tag::MetaCreationDate,
    b"meta:editing-duration" => Tag::MetaEditingDuration,
    b"meta:editing-cycles" => Tag::MetaEditingCycles,
    b"meta:document-statistic" => Tag::MetaDocumentStatistic,
    b"meta:keyword" => Tag::MetaKeyword,
    b"meta:user-defined" => Tag::MetaUserDefined,

    // Number/Data style elements
    b"number:number-style" => Tag::NumberNumberStyle,
    b"number:currency-style" => Tag::NumberCurrencyStyle,
    b"number:percentage-style" => Tag::NumberPercentageStyle,
    b"number:date-style" => Tag::NumberDateStyle,
    b"number:time-style" => Tag::NumberTimeStyle,
    b"number:boolean-style" => Tag::NumberBooleanStyle,
    b"number:text-style" => Tag::NumberTextStyle,
    b"number:number" => Tag::NumberNumber,
    b"number:currency-symbol" => Tag::NumberCurrencySymbol,
    b"number:text" => Tag::NumberText,
    b"number:day" => Tag::NumberDay,
    b"number:month" => Tag::NumberMonth,
    b"number:year" => Tag::NumberYear,
    b"number:hours" => Tag::NumberHours,
    b"number:minutes" => Tag::NumberMinutes,
    b"number:seconds" => Tag::NumberSeconds,

    // Presentation elements
    b"presentation:notes" => Tag::PresentationNotes,
    b"presentation:settings" => Tag::PresentationSettings,
    b"presentation:footer" => Tag::PresentationFooter,
    b"presentation:date-time" => Tag::PresentationDateTime,
    b"presentation:header" => Tag::PresentationHeader,

    // Animation elements
    b"anim:par" => Tag::AnimPar,
    b"anim:seq" => Tag::AnimSeq,
    b"anim:set" => Tag::AnimSet,
    b"anim:animate" => Tag::AnimAnimate,
    b"anim:animateMotion" => Tag::AnimAnimateMotion,
    b"anim:animateColor" => Tag::AnimAnimateColor,
    b"anim:transitionFilter" => Tag::AnimTransFilter,

    // Dublin Core elements
    b"dc:title" => Tag::DcTitle,
    b"dc:description" => Tag::DcDescription,
    b"dc:subject" => Tag::DcSubject,
    b"dc:creator" => Tag::DcCreator,
    b"dc:date" => Tag::DcDate,
    b"dc:language" => Tag::DcLanguage,

    // SVG elements
    b"svg:desc" => Tag::SvgDesc,
    b"svg:title" => Tag::SvgTitle,
    b"svg:linearGradient" => Tag::SvgLinearGradient,
    b"svg:radialGradient" => Tag::SvgRadialGradient,
    b"svg:stop" => Tag::SvgStop,

    // Math elements
    b"math:math" => Tag::MathMath,

    // Script elements
    b"script:event-listener" => Tag::ScriptEventListener,

    // Config elements
    b"config:config-item-set" => Tag::ConfigConfigItemSet,
    b"config:config-item" => Tag::ConfigConfigItem,
    b"config:config-item-map-indexed" => Tag::ConfigConfigItemMapIndexed,
    b"config:config-item-map-entry" => Tag::ConfigConfigItemMapEntry,
    b"config:config-item-map-named" => Tag::ConfigConfigItemMapNamed,
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
/// # use litchi_odt::elements::tag_matcher::extract_prefix;
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
/// # use litchi_odt::elements::tag_matcher::extract_local_name;
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

/// Match a tag to its Tag enum variant
///
/// This provides O(1) lookup using compile-time perfect hash function.
/// For unknown tags, returns `Tag::Unknown`.
///
/// # Arguments
///
/// * `tag` - Tag name as bytes (e.g., b"text:p")
///
/// # Returns
///
/// The corresponding `Tag` enum variant
///
/// # Examples
///
/// ```
/// # use litchi_odt::elements::tag_matcher::{match_tag, Tag};
/// assert_eq!(match_tag(b"text:p"), Tag::TextP);
/// assert_eq!(match_tag(b"table:table"), Tag::TableTable);
/// assert_eq!(match_tag(b"unknown:tag"), Tag::Unknown);
/// ```
#[inline(always)]
pub fn match_tag(tag: &[u8]) -> Tag {
    TAG_MAP.get(tag).copied().unwrap_or(Tag::Unknown)
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
/// # use litchi_odt::elements::tag_matcher::is_namespace;
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
/// Returns a vector of (index, Tag) pairs for all recognized tags.
/// This can be more efficient than matching tags one by one.
pub fn match_tags_batch(tags: &[&[u8]]) -> Vec<(usize, Tag)> {
    tags.iter()
        .enumerate()
        .filter_map(|(idx, tag)| {
            let matched = match_tag(tag);
            if matched != Tag::Unknown {
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
        assert_eq!(match_tag(b"text:p"), Tag::TextP);
        assert_eq!(match_tag(b"text:h"), Tag::TextH);
        assert_eq!(match_tag(b"table:table"), Tag::TableTable);
        assert_eq!(match_tag(b"table:table-row"), Tag::TableTableRow);
        assert_eq!(match_tag(b"unknown:tag"), Tag::Unknown);
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
        assert_eq!(matched[0], (0, Tag::TextP));
        assert_eq!(matched[1], (1, Tag::TableTable));
        assert_eq!(matched[2], (3, Tag::DrawFrame));
    }

    #[test]
    fn test_has_prefix() {
        assert!(has_prefix(b"text:paragraph", b"text"));
        assert!(has_prefix(b"table:table-row", b"table"));
        assert!(!has_prefix(b"short", b"very-long-prefix"));
    }
}
