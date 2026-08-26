#![cfg_attr(not(test), deny(clippy::indexing_slicing))]

//! RTF parser that builds document structure from tokens.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "items stay grouped by RTF feature area rather than by item kind"
)]
use super::super::error::{RtfError, RtfResult};
use super::super::lexer::{ControlWord, Token};
use super::super::limits::ParseLimits;
use super::super::types::{
    Alignment, AnimatedTextEffect, AssociatedCharacterFormatting, CharacterGrid, CharacterType,
    Color, ColorRef, ColorTable, EmbeddedFont, EmbeddedFontFormat, EmphasisMark, FitText, Font,
    FontCharset, FontFamily, FontPage, FontPitch, FontRef, FontTable, FontTheme, Formatting,
    MAX_PARAGRAPH_DROP_CAP_LINES, Paragraph, ParagraphDropCap, ParagraphDropCapKind,
    ParagraphFontAlignment, ParagraphWrapping, RevisionMetadata, StyleBlock, TextDirection,
    UnderlineStyle,
};
use bumpalo::Bump;
use litchi_codepage::Mbcs;
use smallvec::SmallVec;
use std::borrow::Cow;
use std::cell::RefCell;
use std::collections::HashMap;
use std::mem::size_of;
use std::num::NonZeroU16;
use std::ops::Range;

const BODY_BLOCK_RESERVE_SOURCE_MULTIPLIER: usize = 24;
const MAX_BODY_BLOCK_RESERVE_BYTES: usize = 16 * 1_048_576;
const MIN_BODY_BLOCK_RESERVE_SOURCE_BYTES: usize = 64 * 1_024;

fn initial_body_block_capacity(
    source_len: Option<usize>,
    potential_blocks: usize,
    max_tokens: usize,
) -> usize {
    let source_len = source_len.unwrap_or(0);
    if source_len < MIN_BODY_BLOCK_RESERVE_SOURCE_BYTES {
        return 0;
    }
    let source_relative_bytes = source_len.saturating_mul(BODY_BLOCK_RESERVE_SOURCE_MULTIPLIER);
    let capacity_ceiling =
        source_relative_bytes.min(MAX_BODY_BLOCK_RESERVE_BYTES) / size_of::<StyleBlock<'static>>();
    potential_blocks.min(capacity_ceiling).min(max_tokens)
}

fn disables_body_block_reservation(control: &ControlWord<'_>) -> bool {
    matches!(
        control,
        ControlWord::TableNestingLevel(_)
            | ControlWord::TableRowDefaults
            | ControlWord::TableRow
            | ControlWord::TableCell
            | ControlWord::InTable
            | ControlWord::CellX(_)
            | ControlWord::NestedTableCell(_)
            | ControlWord::NestedTableRow(_)
            | ControlWord::NestedTableProperties(_)
            | ControlWord::Deleted(true)
    )
}

#[derive(Debug, Clone, Copy)]
enum RtfEncoding {
    Standard(Mbcs),
    Cp437,
    Cp850,
}

#[derive(Default)]
struct DeferredText {
    parts: Vec<DeferredTextPart>,
    source_len: usize,
}

enum DeferredTextPart {
    Bytes(Vec<u8>),
    Unicode(String),
}

impl DeferredText {
    fn push_transport(&mut self, text: &str) -> RtfResult<()> {
        let mut incoming = Vec::with_capacity(text.len());
        append_transport_bytes(&mut incoming, text)?;
        self.source_len = self.source_len.saturating_add(incoming.len());
        if let Some(DeferredTextPart::Bytes(bytes)) = self.parts.last_mut() {
            bytes.extend_from_slice(&incoming);
        } else if !incoming.is_empty() {
            self.parts.push(DeferredTextPart::Bytes(incoming));
        }
        Ok(())
    }

    fn push_unicode(&mut self, text: &str) {
        self.source_len = self.source_len.saturating_add(text.len());
        if let Some(DeferredTextPart::Unicode(value)) = self.parts.last_mut() {
            value.push_str(text);
        } else if !text.is_empty() {
            self.parts.push(DeferredTextPart::Unicode(text.to_string()));
        }
    }

    const fn source_len(&self) -> usize {
        self.source_len
    }

    fn has_non_ascii_transport(&self) -> bool {
        self.parts
            .iter()
            .any(|part| matches!(part, DeferredTextPart::Bytes(bytes) if !bytes.is_ascii()))
    }

    fn decode(self, encoding: RtfEncoding, context: &str) -> RtfResult<String> {
        let mut output = String::with_capacity(self.source_len);
        for part in self.parts {
            match part {
                DeferredTextPart::Bytes(bytes) => {
                    output.push_str(&encoding.decode_strict(&bytes, context)?);
                },
                DeferredTextPart::Unicode(text) => output.push_str(&text),
            }
        }
        Ok(output)
    }
}

fn strict_paragraph_toggle(value: Option<i32>, name: &str) -> Result<bool, RtfError> {
    match value {
        None | Some(1) => Ok(true),
        Some(0) => Ok(false),
        Some(_) => Err(RtfError::MalformedDocument(format!(
            "RTF {name} accepts only 0 or 1"
        ))),
    }
}

#[cold]
pub(super) fn parser_classification_error() -> RtfError {
    RtfError::ParserError("RTF parser control classification invariant failed".to_string())
}

fn strict_paragraph_selector(value: Option<i32>, name: &str) -> Result<(), RtfError> {
    if value.is_some() {
        return Err(RtfError::MalformedDocument(format!(
            "RTF {name} must not have a numeric parameter"
        )));
    }
    Ok(())
}

fn required_paragraph_bool(value: Option<i32>, name: &str) -> Result<bool, RtfError> {
    match value {
        Some(0) => Ok(false),
        Some(1) => Ok(true),
        None => Err(RtfError::MalformedDocument(format!(
            "RTF {name} requires 0 or 1"
        ))),
        Some(_) => Err(RtfError::MalformedDocument(format!(
            "RTF {name} accepts only 0 or 1"
        ))),
    }
}

fn required_list_spacing(raw_value: Option<i32>, name: &str) -> Result<u32, RtfError> {
    let value = raw_value.ok_or_else(|| {
        RtfError::MalformedDocument(format!("RTF {name} requires a numeric parameter"))
    })?;
    u32::try_from(value)
        .ok()
        .filter(|parsed| *parsed <= 1_000_000)
        .ok_or_else(|| RtfError::MalformedDocument(format!("RTF {name} must be in 0..=1000000")))
}
fn required_paragraph_indent(raw_value: Option<i32>, name: &str) -> Result<i32, RtfError> {
    let value = raw_value.ok_or_else(|| {
        RtfError::MalformedDocument(format!("RTF {name} requires a numeric parameter"))
    })?;
    if value.unsigned_abs() > 10_000_000 {
        return Err(RtfError::MalformedDocument(format!(
            "RTF {name} exceeds the supported range"
        )));
    }
    Ok(value)
}

impl RtfEncoding {
    fn decode(self, bytes: &[u8]) -> Cow<'_, str> {
        match self {
            Self::Standard(page) => page.decode_lossy(bytes),
            Self::Cp437 => decode_dos_codepage(bytes, &CP437_HIGH),
            Self::Cp850 => decode_dos_codepage(bytes, &CP850_HIGH),
        }
    }

    fn decode_strict<'b>(self, bytes: &'b [u8], context: &str) -> RtfResult<Cow<'b, str>> {
        match self {
            Self::Standard(page) => page.decode(bytes).map_err(|error| {
                RtfError::MalformedDocument(format!("invalid RTF {context}: {error}"))
            }),
            Self::Cp437 => Ok(decode_dos_codepage(bytes, &CP437_HIGH)),
            Self::Cp850 => Ok(decode_dos_codepage(bytes, &CP850_HIGH)),
        }
    }

    const fn from_font_page(page: FontPage) -> Self {
        match page {
            FontPage::Mbcs(code_page) => Self::Standard(code_page),
            FontPage::Cp437 => Self::Cp437,
            FontPage::Cp850 => Self::Cp850,
        }
    }
}

fn decode_dos_codepage<'a>(bytes: &'a [u8], high: &[char; 128]) -> Cow<'a, str> {
    if bytes.iter().all(u8::is_ascii) {
        return String::from_utf8_lossy(bytes);
    }
    let mut output = String::with_capacity(bytes.len());
    for &byte in bytes {
        if byte.is_ascii() {
            output.push(char::from(byte));
        } else {
            output.push(
                high.get(usize::from(byte - 0x80))
                    .copied()
                    .unwrap_or('\u{FFFD}'),
            );
        }
    }
    Cow::Owned(output)
}

const CP437_HIGH: [char; 128] = [
    '\u{00C7}', '\u{00FC}', '\u{00E9}', '\u{00E2}', '\u{00E4}', '\u{00E0}', '\u{00E5}', '\u{00E7}',
    '\u{00EA}', '\u{00EB}', '\u{00E8}', '\u{00EF}', '\u{00EE}', '\u{00EC}', '\u{00C4}', '\u{00C5}',
    '\u{00C9}', '\u{00E6}', '\u{00C6}', '\u{00F4}', '\u{00F6}', '\u{00F2}', '\u{00FB}', '\u{00F9}',
    '\u{00FF}', '\u{00D6}', '\u{00DC}', '\u{00A2}', '\u{00A3}', '\u{00A5}', '\u{20A7}', '\u{0192}',
    '\u{00E1}', '\u{00ED}', '\u{00F3}', '\u{00FA}', '\u{00F1}', '\u{00D1}', '\u{00AA}', '\u{00BA}',
    '\u{00BF}', '\u{2310}', '\u{00AC}', '\u{00BD}', '\u{00BC}', '\u{00A1}', '\u{00AB}', '\u{00BB}',
    '\u{2591}', '\u{2592}', '\u{2593}', '\u{2502}', '\u{2524}', '\u{2561}', '\u{2562}', '\u{2556}',
    '\u{2555}', '\u{2563}', '\u{2551}', '\u{2557}', '\u{255D}', '\u{255C}', '\u{255B}', '\u{2510}',
    '\u{2514}', '\u{2534}', '\u{252C}', '\u{251C}', '\u{2500}', '\u{253C}', '\u{255E}', '\u{255F}',
    '\u{255A}', '\u{2554}', '\u{2569}', '\u{2566}', '\u{2560}', '\u{2550}', '\u{256C}', '\u{2567}',
    '\u{2568}', '\u{2564}', '\u{2565}', '\u{2559}', '\u{2558}', '\u{2552}', '\u{2553}', '\u{256B}',
    '\u{256A}', '\u{2518}', '\u{250C}', '\u{2588}', '\u{2584}', '\u{258C}', '\u{2590}', '\u{2580}',
    '\u{03B1}', '\u{00DF}', '\u{0393}', '\u{03C0}', '\u{03A3}', '\u{03C3}', '\u{00B5}', '\u{03C4}',
    '\u{03A6}', '\u{0398}', '\u{03A9}', '\u{03B4}', '\u{221E}', '\u{03C6}', '\u{03B5}', '\u{2229}',
    '\u{2261}', '\u{00B1}', '\u{2265}', '\u{2264}', '\u{2320}', '\u{2321}', '\u{00F7}', '\u{2248}',
    '\u{00B0}', '\u{2219}', '\u{00B7}', '\u{221A}', '\u{207F}', '\u{00B2}', '\u{25A0}', '\u{00A0}',
];

const CP850_HIGH: [char; 128] = [
    '\u{00C7}', '\u{00FC}', '\u{00E9}', '\u{00E2}', '\u{00E4}', '\u{00E0}', '\u{00E5}', '\u{00E7}',
    '\u{00EA}', '\u{00EB}', '\u{00E8}', '\u{00EF}', '\u{00EE}', '\u{00EC}', '\u{00C4}', '\u{00C5}',
    '\u{00C9}', '\u{00E6}', '\u{00C6}', '\u{00F4}', '\u{00F6}', '\u{00F2}', '\u{00FB}', '\u{00F9}',
    '\u{00FF}', '\u{00D6}', '\u{00DC}', '\u{00F8}', '\u{00A3}', '\u{00D8}', '\u{00D7}', '\u{0192}',
    '\u{00E1}', '\u{00ED}', '\u{00F3}', '\u{00FA}', '\u{00F1}', '\u{00D1}', '\u{00AA}', '\u{00BA}',
    '\u{00BF}', '\u{00AE}', '\u{00AC}', '\u{00BD}', '\u{00BC}', '\u{00A1}', '\u{00AB}', '\u{00BB}',
    '\u{2591}', '\u{2592}', '\u{2593}', '\u{2502}', '\u{2524}', '\u{00C1}', '\u{00C2}', '\u{00C0}',
    '\u{00A9}', '\u{2563}', '\u{2551}', '\u{2557}', '\u{255D}', '\u{00A2}', '\u{00A5}', '\u{2510}',
    '\u{2514}', '\u{2534}', '\u{252C}', '\u{251C}', '\u{2500}', '\u{253C}', '\u{00E3}', '\u{00C3}',
    '\u{255A}', '\u{2554}', '\u{2569}', '\u{2566}', '\u{2560}', '\u{2550}', '\u{256C}', '\u{00A4}',
    '\u{00F0}', '\u{00D0}', '\u{00CA}', '\u{00CB}', '\u{00C8}', '\u{0131}', '\u{00CD}', '\u{00CE}',
    '\u{00CF}', '\u{2518}', '\u{250C}', '\u{2588}', '\u{2584}', '\u{00A6}', '\u{00CC}', '\u{2580}',
    '\u{00D3}', '\u{00DF}', '\u{00D4}', '\u{00D2}', '\u{00F5}', '\u{00D5}', '\u{00B5}', '\u{00FE}',
    '\u{00DE}', '\u{00DA}', '\u{00DB}', '\u{00D9}', '\u{00FD}', '\u{00DD}', '\u{00AF}', '\u{00B4}',
    '\u{00AD}', '\u{00B1}', '\u{2017}', '\u{00BE}', '\u{00B6}', '\u{00A7}', '\u{00F7}', '\u{00B8}',
    '\u{00B0}', '\u{00A8}', '\u{00B7}', '\u{00B9}', '\u{00B3}', '\u{00B2}', '\u{25A0}', '\u{00A0}',
];

fn append_transport_bytes(buffer: &mut impl Extend<u8>, text: &str) -> RtfResult<()> {
    if text.is_ascii() {
        buffer.extend(text.bytes());
        return Ok(());
    }

    for character in text.chars() {
        let byte = u8::try_from(character as u32).map_err(|_err| {
            RtfError::InvalidUnicode(
                "RTF source text is not a byte-preserving transport string".to_string(),
            )
        })?;
        buffer.extend(std::iter::once(byte));
    }
    Ok(())
}

#[allow(
    clippy::wildcard_enum_match_arm,
    reason = "remaining variants share the same fallback by design"
)]
fn control_symbol_text(control: &ControlWord<'_>) -> Option<&'static str> {
    match control {
        ControlWord::NonBreakingSpace => Some("\u{00A0}"),
        ControlWord::OptionalHyphen => Some("\u{00AD}"),
        ControlWord::NonBreakingHyphen => Some("\u{2011}"),
        ControlWord::EmDash => Some("\u{2014}"),
        ControlWord::EnDash => Some("\u{2013}"),
        ControlWord::EmSpace => Some("\u{2003}"),
        ControlWord::EnSpace => Some("\u{2002}"),
        ControlWord::QuarterEmSpace => Some("\u{2005}"),
        ControlWord::Bullet => Some("\u{2022}"),
        ControlWord::LeftSingleQuote => Some("\u{2018}"),
        ControlWord::RightSingleQuote => Some("\u{2019}"),
        ControlWord::LeftDoubleQuote => Some("\u{201C}"),
        ControlWord::RightDoubleQuote => Some("\u{201D}"),
        ControlWord::LeftToRightMark => Some("\u{200E}"),
        ControlWord::RightToLeftMark => Some("\u{200F}"),
        ControlWord::ZeroWidthJoiner => Some("\u{200D}"),
        ControlWord::ZeroWidthNonJoiner => Some("\u{200C}"),
        ControlWord::ZeroWidthBreakOpportunity => Some("\u{200B}"),
        ControlWord::ZeroWidthNoBreakOpportunity => Some("\u{FEFF}"),
        _ => None,
    }
}

fn duplicate_mail_merge(name: &str) -> RtfError {
    RtfError::MalformedDocument(format!(
        "RTF mail-merge destination contains duplicate {name} metadata"
    ))
}

fn nonnegative_mail_merge(value: i32, name: &str) -> RtfResult<u32> {
    u32::try_from(value).map_err(|_err| {
        RtfError::MalformedDocument(format!("RTF mail-merge {name} cannot be negative"))
    })
}

fn set_mail_merge_text<'a>(
    slot: &mut Option<Cow<'a, str>>,
    value: Cow<'a, str>,
    name: &str,
) -> RtfResult<()> {
    if slot.replace(value).is_some() {
        return Err(duplicate_mail_merge(name));
    }
    Ok(())
}

/// RTF destination type - determines if we're in document body or header
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Destination {
    /// Main document body - text should be extracted
    DocumentBody,
    /// Font table - should be skipped
    FontTable,
    /// Color table - should be skipped
    ColorTable,
    /// Stylesheet - should be skipped
    StyleSheet,
    /// Document info - should be skipped
    Info,
    /// Picture data - extract and process embedded images
    Picture,
    /// Result of embedded object rendering - should be skipped
    Result,
    /// Field instruction
    FieldInstruction,
    /// Field result
    FieldResult,
    /// Header content
    Header,
    /// Footer content
    Footer,
    /// Footnote content
    Footnote,
    /// Endnote content
    Endnote,
    /// End-defined row properties for a nested table.
    NestedTableProperties,
    /// Revision/track changes
    #[allow(dead_code, reason = "reserved for revision-destination tracking")]
    Revision,
    /// Other destinations - should be skipped
    Other,
}

#[derive(Clone, Copy)]
enum InfoTextField {
    Title,
    Subject,
    Author,
    Manager,
    Company,
    Operator,
    Category,
    Keywords,
    Comment,
    DocumentComment,
    HyperlinkBase,
}

#[derive(Clone, Copy)]
enum InfoTimeField {
    Creation,
    Revision,
    Print,
    Backup,
}

const MAX_INFO_TEXT_BYTES: usize = 1_048_576;
const MAX_BOOKMARKS: usize = 65_536;
const MAX_BOOKMARK_NAME_BYTES: usize = 65_536;
const MAX_ANNOTATIONS: usize = super::super::annotation::MAX_ANNOTATIONS;
const MAX_ANNOTATION_TEXT_BYTES: usize = 4 * 1_048_576;
const MAX_SECTIONS: usize = 4_096;
use super::super::list::{MAX_LIST_LEVELS, MAX_LIST_TABS, MAX_LIST_TEXT_BYTES, MAX_LISTS};
use super::super::stylesheet::{MAX_STYLE_NAME_BYTES, MAX_STYLES};
const MAX_REVISION_AUTHORS: usize = super::super::annotation::MAX_REVISION_AUTHORS;
const MAX_REVISION_AUTHOR_BYTES: usize = 65_536;
const MAX_REVISIONS: usize = super::super::annotation::MAX_REVISIONS;
const MAX_SHAPES: usize = 65_536;
const MAX_SHAPE_GROUPS: usize = 16_384;
const MAX_SHAPES_PER_GROUP: usize = 65_536;
const MAX_GROUPS_PER_GROUP: usize = 16_384;
/// Hard ceiling on how deeply RTF groups (`{` ... `}`) may nest.
///
/// Group parsing is recursive, so an unbounded nesting depth lets a hostile or
/// merely corrupt file exhaust the call stack and abort the process rather than
/// surface a recoverable error. The deepest nesting observed anywhere in the
/// real-world compatibility corpus is 15 levels, so this leaves ample room for
/// genuine documents while keeping the worst-case stack cost comfortably inside
/// a default 2 MiB thread stack even in unoptimised builds.
pub(crate) const MAX_GROUP_NESTING_DEPTH: usize = 32;
const MAX_SHAPE_GROUP_DEPTH: usize = 64;
const MAX_STORY_GROUP_DEPTH: usize = 64;
const MAX_SHAPE_PROPERTIES: usize = 65_536;
const MAX_SHAPE_PROPERTY_BYTES: usize = 1_048_576;
const MAX_SHAPE_TEXT_BYTES: usize = 16 * 1_048_576;
/// Bound temporary transport/Unicode buffers before they can spill into a
/// large heap allocation.  Larger retained stories use their own aggregate
/// limits, but decoder intermediates stay bounded independently.
pub(crate) const MAX_TEXT_INTERMEDIATE_BYTES: usize = 64 * 1_024;
pub(crate) const MAX_OBJECTS: usize = 65_536;
const MAX_OBJECT_TEXT_BYTES: usize = 1_048_576;
pub(crate) const MAX_OBJECT_DATA_BYTES: usize = crate::object::MAX_OBJECT_DATA_BYTES;
pub(crate) const MAX_PICTURE_DATA_BYTES: usize = crate::picture::MAX_PICTURE_WRITE_BYTES;
use super::super::document_variable::{
    DocumentVariable, MAX_DOCUMENT_VARIABLE_NAME_BYTES, MAX_DOCUMENT_VARIABLE_TEXT_BYTES,
    MAX_DOCUMENT_VARIABLE_VALUE_BYTES, MAX_DOCUMENT_VARIABLES,
};
use super::super::navigation_entry::{
    IndexEntry, IndexPageReference, MAX_NAVIGATION_ENTRIES, MAX_NAVIGATION_ENTRY_DEPTH,
    MAX_NAVIGATION_ENTRY_TEXT_BYTES, MAX_NAVIGATION_ENTRY_TEXT_TOTAL_BYTES, NavigationEntry,
    TableOfContentsEntry,
};
use super::super::user_property::{
    MAX_USER_PROPERTIES, MAX_USER_PROPERTY_NAME_BYTES, MAX_USER_PROPERTY_TEXT_BYTES,
    MAX_USER_PROPERTY_VALUE_BYTES, UserProperty, UserPropertyValue,
};

struct OpenBookmark {
    name: String,
    position: usize,
    first_column: Option<i32>,
    last_column: Option<i32>,
    is_public: bool,
    order: usize,
}

struct BookmarkSpan {
    bookmark: OpenBookmark,
    end: usize,
}

struct OpenCustomXmlTag {
    name: String,
    namespace: Option<u32>,
    attributes: Vec<(String, String)>,
    position: usize,
    order: usize,
}

struct CustomXmlSpan {
    tag: OpenCustomXmlTag,
    end: usize,
}

struct OpenProtectionRange {
    id: String,
    position: usize,
    order: usize,
}

struct ProtectionRangeSpan {
    range: OpenProtectionRange,
    end: usize,
}

struct OpenEditableRegion {
    position: usize,
    order: usize,
}

struct EditableRegionSpan {
    position: usize,
    order: usize,
    end: usize,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum ParagraphBorderSide {
    Top,
    Bottom,
    Left,
    Right,
    Bar,
    Between,
    Box,
}

impl ParagraphBorderSide {
    const fn bit(self) -> u8 {
        match self {
            Self::Top => 1,
            Self::Bottom => 2,
            Self::Left => 4,
            Self::Right => 8,
            Self::Bar => 16,
            Self::Between => 32,
            Self::Box => 64,
        }
    }
}

#[derive(Debug, Clone, Copy)]
enum ParsedBodyStoryEvent {
    Resolved(crate::BodyStoryEvent),
    AnnotationStart(i32),
    AnnotationEnd(i32),
    RevisionStart(usize),
    RevisionEnd(usize),
    RevisionDeletion(usize),
}

#[allow(
    clippy::struct_excessive_bools,
    reason = "independent RTF feature flags stay flat for direct access"
)]
/// Parser state for tracking formatting context.
#[derive(Debug, Clone)]
struct State {
    /// Current character formatting
    formatting: Formatting,
    /// Whether generic border controls currently belong to `chbrdr`.
    character_border_active: bool,
    /// Components already supplied for the current character border.
    character_border_seen: u8,
    /// Current paragraph properties
    paragraph: Paragraph,
    /// Border segment currently collecting properties (`\brdrt`, `\box`, ...).
    paragraph_border_side: Option<ParagraphBorderSide>,
    /// Paragraph border segments already declared in this paragraph.
    paragraph_border_seen: u8,
    /// Parsed `dropcapt` value retained independently until the pair is complete.
    drop_cap_kind: Option<ParagraphDropCapKind>,
    /// Parsed `dropcapli` value retained independently until the pair is complete.
    drop_cap_lines: Option<u8>,
    /// Optional alignment selector for the next `tx` tab terminator.
    pending_tab_alignment: Option<super::super::border::TabAlignment>,
    /// Optional leader selector for the next `tx` or `tb` tab terminator.
    pending_tab_leader: Option<super::super::border::TabLeader>,
    /// Unicode skip count (characters to skip after \u)
    unicode_skip: i32,
    /// Whether we're inside a table
    in_table: bool,
    table_nesting_level: u8,
    /// Cell boundaries for current row (in twips)
    cell_boundaries: SmallVec<[i32; 8]>,
    table_style: Option<u16>,
    table_rsid: Option<u32>,
    table_row_revision: RevisionMetadata,
    table_row_padding: crate::TableEdgeDistances,
    table_row_spacing: crate::TableEdgeDistances,
    table_row_positioning: crate::FloatingTablePosition,
    table_row_direction: Option<TextDirection>,
    table_row_layout: crate::TableRowLayout,
    table_row_borders: crate::TableRowBorders,
    table_row_shading: crate::TableShading,
    table_row_geometry: crate::TableRowGeometry,
    table_default_borders: crate::TableStyleDefaultBorders,
    table_default_padding: crate::TableEdgeDistances,
    table_default_spacing: crate::TableEdgeDistances,
    table_default_width_unit: Option<crate::TablePreferredWidthUnit>,
    table_default_width_value: Option<i32>,
    table_autoformat_flags: crate::TableAutoformatFlags,
    table_row_banding: crate::TableRowBanding,
    table_row_index_seen: bool,
    table_row_band_index_seen: bool,
    table_last_row_seen: bool,
    table_width_unit: Option<crate::TablePreferredWidthUnit>,
    table_width_value: Option<i32>,
    table_leading_width_unit: Option<crate::TablePreferredWidthUnit>,
    table_leading_width_value: Option<i32>,
    table_trailing_width_unit: Option<crate::TablePreferredWidthUnit>,
    table_trailing_width_value: Option<i32>,
    table_indent_value: Option<i32>,
    table_indent_unit: Option<crate::TableIndentUnit>,
    pending_cell_padding: crate::TableEdgeDistances,
    pending_cell_spacing: crate::TableEdgeDistances,
    pending_cell_layout: crate::TableCellLayout,
    pending_cell_merge: crate::TableCellMergeState,
    pending_cell_revision: Option<crate::CellRevision>,
    pending_cell_borders: crate::TableCellBorders,
    pending_cell_shading: crate::TableShading,
    pending_cell_width_unit: Option<crate::TablePreferredWidthUnit>,
    pending_cell_width_value: Option<i32>,
    table_row_shading_seen: u8,
    pending_cell_shading_seen: u8,
    active_table_border: Option<crate::table::TableBorderTarget>,
    active_table_border_seen: u8,
    cell_distances: SmallVec<[(crate::TableEdgeDistances, crate::TableEdgeDistances); 8]>,
    cell_layouts: SmallVec<[crate::TableCellLayout; 8]>,
    cell_merges: SmallVec<[crate::TableCellMergeState; 8]>,
    cell_revisions: SmallVec<[Option<crate::CellRevision>; 8]>,
    cell_decorations: SmallVec<[(crate::TableCellBorders, crate::TableShading); 8]>,
    cell_widths: SmallVec<[Option<crate::TablePreferredWidth>; 8]>,
    /// Current destination (for skipping non-document content)
    destination: Destination,
    /// Whether this body-flow group is an explicit section-format snapshot.
    visible_section_format: bool,
    /// One-based explicit section column selected by `colno` in this group.
    section_column_number: Option<u16>,
    /// Current text encoding
    encoding: RtfEncoding,
    /// Active tracked-change kind for text emitted in this state
    revision_type: Option<super::super::annotation::RevisionType>,
    /// Revision-author table index
    revision_author_id: Option<i32>,
    /// Packed RTF revision timestamp
    revision_date: Option<i32>,
    revision_event_id: Option<usize>,
    paragraph_content_started: bool,
    paragraph_numbering_declared: bool,
}

fn is_section_control(control: &ControlWord<'_>) -> bool {
    matches!(
        control,
        ControlWord::SectionDefault
            | ControlWord::SectionStyle(_)
            | ControlWord::SectionBreak
            | ControlWord::TitlePage
            | ControlWord::SectionRsid(_)
            | ControlWord::SectionContinuous
            | ControlWord::SectionColumn
            | ControlWord::SectionPage
            | ControlWord::SectionEvenPage
            | ControlWord::SectionOddPage
            | ControlWord::PageWidth(_)
            | ControlWord::PageHeight(_)
            | ControlWord::MarginLeft(_)
            | ControlWord::MarginRight(_)
            | ControlWord::MarginTop(_)
            | ControlWord::MarginBottom(_)
            | ControlWord::MarginGutter(_)
            | ControlWord::PaperSourceFirst(_)
            | ControlWord::PaperSourceOther(_)
            | ControlWord::HeaderDistance(_)
            | ControlWord::FooterDistance(_)
            | ControlWord::Landscape
            | ControlWord::Columns(_)
            | ControlWord::ColumnSpace(_)
            | ControlWord::ColumnNumber(_)
            | ControlWord::ColumnWidth(_)
            | ControlWord::ColumnSpaceRight(_)
            | ControlWord::ColumnSeparator(_)
            | ControlWord::PageNumberStart(_)
            | ControlWord::PageNumberFormat(_)
            | ControlWord::PageNumberRestart(_)
            | ControlWord::PageNumberOffsetX(_)
            | ControlWord::PageNumberOffsetY(_)
            | ControlWord::PageNumberHeadingLevel(_)
            | ControlWord::PageNumberHeadingSeparator(_)
            | ControlWord::SectionLineGrid(_)
            | ControlWord::SectionDocumentGrid(_)
            | ControlWord::SectionRevisionAuthor(_)
            | ControlWord::SectionRevisionDate(_)
            | ControlWord::VerticalAlignTop
            | ControlWord::VerticalAlignCenter
            | ControlWord::VerticalAlignJustify
            | ControlWord::VerticalAlignBottom
            | ControlWord::LineNumbering(_)
            | ControlWord::LineNumberDistance(_)
            | ControlWord::LineNumberStart(_)
            | ControlWord::LineNumberRestartSection
            | ControlWord::LineNumberRestartPage
            | ControlWord::LineNumberContinuous
            | ControlWord::LeftToRightSection
            | ControlWord::RightToLeftSection
            | ControlWord::SectionVerticalRendering(_)
            | ControlWord::SectionHorizontalRendering(_)
            | ControlWord::SectionNoColumnBalance(_)
            | ControlWord::SectionDefaultColumns(_)
            | ControlWord::SectionFootnotePlacement(_)
            | ControlWord::SectionEndnoteHere
            | ControlWord::SectionFootnoteStart(_)
            | ControlWord::SectionEndnoteStart(_)
            | ControlWord::SectionFootnoteRestart(_)
            | ControlWord::SectionEndnoteRestart(_)
            | ControlWord::SectionFootnoteNumbering(_)
            | ControlWord::SectionEndnoteNumbering(_)
            | ControlWord::PageBorderTop
            | ControlWord::PageBorderLeft
            | ControlWord::PageBorderBottom
            | ControlWord::PageBorderRight
            | ControlWord::PageBorderOptions(_)
            | ControlWord::PageBorderSurroundHeader
            | ControlWord::PageBorderSurroundFooter
            | ControlWord::PageBorderSnap
    )
}

#[allow(
    clippy::cast_possible_truncation,
    clippy::cast_sign_loss,
    reason = "adjacent range checks bound each narrowing conversion to the target type's range"
)]
fn apply_table_distance_side(
    distances: &mut crate::TableEdgeDistances,
    edge: crate::TableEdge,
    parameter: Option<i32>,
    unit: bool,
) -> RtfResult<()> {
    let value = parameter.ok_or_else(|| {
        RtfError::MalformedDocument("RTF table distance control requires a parameter".to_string())
    })?;
    let side = distances.side_mut(edge);
    if unit {
        side.unit = Some(match value {
            0 => crate::TableDistanceUnit::Null,
            3 => crate::TableDistanceUnit::Twips,
            _ => {
                return Err(RtfError::MalformedDocument(
                    "RTF table distance unit must be 0 or 3".to_string(),
                ));
            },
        });
    } else {
        if !(0..=crate::MAX_TABLE_DISTANCE_TWIPS).contains(&value) {
            return Err(RtfError::MalformedDocument(
                "RTF table distance value is out of range".to_string(),
            ));
        }
        side.value = Some(value as u16);
    }
    Ok(())
}

fn apply_table_distance(
    state: &mut State,
    target: crate::TableDistanceTarget,
    parameter: Option<i32>,
    unit: bool,
) -> RtfResult<()> {
    let distances = match (target.scope, target.kind) {
        (crate::TableDistanceScope::Row, crate::TableDistanceKind::Padding) => {
            &mut state.table_row_padding
        },
        (crate::TableDistanceScope::Row, crate::TableDistanceKind::Spacing) => {
            &mut state.table_row_spacing
        },
        (crate::TableDistanceScope::Cell, crate::TableDistanceKind::Padding) => {
            &mut state.pending_cell_padding
        },
        (crate::TableDistanceScope::Cell, crate::TableDistanceKind::Spacing) => {
            &mut state.pending_cell_spacing
        },
    };
    apply_table_distance_side(distances, target.edge, parameter, unit)
}

fn table_width_unit(parameter: Option<i32>) -> RtfResult<crate::TablePreferredWidthUnit> {
    match parameter {
        Some(0) => Ok(crate::TablePreferredWidthUnit::Null),
        Some(1) => Ok(crate::TablePreferredWidthUnit::Auto),
        Some(2) => Ok(crate::TablePreferredWidthUnit::Percent),
        Some(3) => Ok(crate::TablePreferredWidthUnit::Twips),
        None => Err(RtfError::MalformedDocument(
            "RTF preferred-width unit requires a parameter".to_string(),
        )),
        Some(_) => Err(RtfError::MalformedDocument(
            "RTF preferred-width unit must be in 0..=3".to_string(),
        )),
    }
}
fn table_indent_unit(parameter: Option<i32>) -> RtfResult<crate::TableIndentUnit> {
    match parameter {
        Some(0) => Ok(crate::TableIndentUnit::Auto),
        Some(1) => Ok(crate::TableIndentUnit::Twips),
        Some(2) => Ok(crate::TableIndentUnit::Nil),
        Some(3) => Ok(crate::TableIndentUnit::Percent),
        None => Err(RtfError::MalformedDocument(
            "RTF tblindtype requires a parameter".to_string(),
        )),
        Some(_) => Err(RtfError::MalformedDocument(
            "RTF tblindtype must be in 0..=3".to_string(),
        )),
    }
}
#[allow(
    clippy::cast_possible_truncation,
    clippy::cast_sign_loss,
    reason = "the crate width constants are defined within the u16 range"
)]
fn resolve_preferred_width(
    unit: Option<crate::TablePreferredWidthUnit>,
    value: Option<i32>,
) -> RtfResult<Option<crate::TablePreferredWidth>> {
    let Some(width_unit) = unit else {
        return if value.is_none() {
            Ok(None)
        } else {
            Err(RtfError::MalformedDocument(
                "RTF preferred-width value lacks its unit control".to_string(),
            ))
        };
    };
    let width_value = match width_unit {
        crate::TablePreferredWidthUnit::Null | crate::TablePreferredWidthUnit::Auto => {
            if value.is_some_and(|v| v != 0) {
                return Err(RtfError::MalformedDocument(
                    "RTF null or auto preferred width must omit its value or use zero".to_string(),
                ));
            }
            None
        },
        crate::TablePreferredWidthUnit::Percent => Some(required_table_value(
            value,
            "preferred width percentage",
            crate::MAX_TABLE_WIDTH_PERCENT as u16,
        )?),
        crate::TablePreferredWidthUnit::Twips => Some(required_table_value(
            value,
            "preferred width",
            crate::MAX_TABLE_GEOMETRY_TWIPS as u16,
        )?),
    };
    Ok(Some(crate::TablePreferredWidth::new(
        width_unit,
        width_value,
    )?))
}
#[allow(
    clippy::cast_possible_truncation,
    clippy::cast_sign_loss,
    reason = "the crate width constants are defined within the u16 range"
)]
fn resolve_invisible_width(
    unit: Option<crate::TablePreferredWidthUnit>,
    value: Option<i32>,
    side: &str,
) -> RtfResult<Option<crate::TablePreferredWidth>> {
    let Some(width_unit) = unit else {
        return if value.is_none() {
            Ok(None)
        } else {
            Err(RtfError::MalformedDocument(format!(
                "RTF {side} invisible-width value lacks its unit control"
            )))
        };
    };
    let width_value = match width_unit {
        crate::TablePreferredWidthUnit::Null | crate::TablePreferredWidthUnit::Auto => {
            if value.is_some_and(|v| v != 0) {
                return Err(RtfError::MalformedDocument(format!(
                    "RTF null or auto {side} invisible width must omit its value or use zero"
                )));
            }
            None
        },
        crate::TablePreferredWidthUnit::Percent => Some(required_table_value(
            Some(value.unwrap_or(0)),
            &format!("{side} invisible width percentage"),
            crate::MAX_TABLE_WIDTH_PERCENT as u16,
        )?),
        crate::TablePreferredWidthUnit::Twips => Some(required_table_value(
            Some(value.unwrap_or(0)),
            &format!("{side} invisible width"),
            crate::MAX_TABLE_GEOMETRY_TWIPS as u16,
        )?),
    };
    Ok(Some(crate::TablePreferredWidth::new(
        width_unit,
        width_value,
    )?))
}
fn resolve_row_geometry(state: &State) -> RtfResult<crate::TableRowGeometry> {
    let mut geometry = state.table_row_geometry;
    geometry.set_preferred_width(resolve_preferred_width(
        state.table_width_unit,
        state.table_width_value,
    )?);
    geometry.set_leading_invisible_width(resolve_invisible_width(
        state.table_leading_width_unit,
        state.table_leading_width_value,
        "leading",
    )?);
    geometry.set_trailing_invisible_width(resolve_invisible_width(
        state.table_trailing_width_unit,
        state.table_trailing_width_value,
        "trailing",
    )?);
    if state.table_indent_value.is_some() || state.table_indent_unit.is_some() {
        let unit = state
            .table_indent_unit
            .unwrap_or(crate::TableIndentUnit::Twips);
        geometry.set_indent(Some(crate::TableIndent::new(
            unit,
            state.table_indent_value.unwrap_or(0),
        )?));
    }
    geometry.validate()?;
    Ok(geometry)
}
fn table_geometry_twips(parameter: Option<i32>, name: &str, signed: bool) -> RtfResult<i32> {
    let value = parameter
        .ok_or_else(|| RtfError::MalformedDocument(format!("RTF {name} requires a parameter")))?;
    let valid = if signed {
        value.unsigned_abs() <= crate::MAX_TABLE_GEOMETRY_TWIPS as u32
    } else {
        (0..=crate::MAX_TABLE_GEOMETRY_TWIPS).contains(&value)
    };
    if valid {
        Ok(value)
    } else {
        Err(RtfError::MalformedDocument(format!(
            "RTF {name} is out of range"
        )))
    }
}
#[allow(
    clippy::cast_possible_truncation,
    clippy::cast_sign_loss,
    reason = "adjacent range checks bound each narrowing conversion to the target type's range"
)]
fn table_row_height(parameter: Option<i32>) -> RtfResult<crate::TableRowHeight> {
    let value = parameter
        .ok_or_else(|| RtfError::MalformedDocument("RTF trrh requires a parameter".to_string()))?;
    if value.unsigned_abs() > crate::MAX_TABLE_GEOMETRY_TWIPS as u32 {
        return Err(RtfError::MalformedDocument(
            "RTF trrh is out of range".to_string(),
        ));
    }
    Ok(match value.cmp(&0) {
        std::cmp::Ordering::Equal => crate::TableRowHeight::Automatic,
        std::cmp::Ordering::Greater => crate::TableRowHeight::Minimum(value as u16),
        std::cmp::Ordering::Less => crate::TableRowHeight::Exact(value.unsigned_abs() as u16),
    })
}

fn require_parameterless(parameter: Option<i32>, name: &str) -> RtfResult<()> {
    if parameter.is_some() {
        return Err(RtfError::MalformedDocument(format!(
            "RTF {name} does not accept a parameter"
        )));
    }
    Ok(())
}

fn required_table_value(value: Option<i32>, name: &str, maximum: u16) -> RtfResult<u16> {
    let parameter = value.ok_or_else(|| {
        RtfError::MalformedDocument(format!("RTF {name} requires a numeric parameter"))
    })?;
    let table_value = u16::try_from(parameter).map_err(|_err| {
        RtfError::MalformedDocument(format!("RTF {name} value must be in 0..={maximum}"))
    })?;
    if table_value > maximum {
        return Err(RtfError::MalformedDocument(format!(
            "RTF {name} value must be in 0..={maximum}"
        )));
    }
    Ok(table_value)
}
fn floating_table_offset(parameter: Option<i32>, negative: bool, axis: &str) -> RtfResult<i32> {
    let value = parameter.ok_or_else(|| {
        RtfError::MalformedDocument(format!(
            "RTF floating-table {axis} offset requires a parameter"
        ))
    })?;
    let valid = if negative {
        (-crate::MAX_FLOATING_TABLE_DISTANCE_TWIPS..=-1).contains(&value)
    } else {
        (0..=crate::MAX_FLOATING_TABLE_DISTANCE_TWIPS).contains(&value)
    };
    if !valid {
        return Err(RtfError::MalformedDocument(format!(
            "RTF floating-table {axis} offset is out of range"
        )));
    }
    Ok(value)
}
#[allow(
    clippy::cast_possible_truncation,
    clippy::cast_sign_loss,
    reason = "adjacent range checks bound each narrowing conversion to the target type's range"
)]
fn floating_table_wrap_distance(parameter: Option<i32>) -> RtfResult<u16> {
    let value = parameter.ok_or_else(|| {
        RtfError::MalformedDocument(
            "RTF floating-table wrap distance requires a parameter".to_string(),
        )
    })?;
    if !(0..=crate::MAX_FLOATING_TABLE_DISTANCE_TWIPS).contains(&value) {
        return Err(RtfError::MalformedDocument(
            "RTF floating-table wrap distance is out of range".to_string(),
        ));
    }
    Ok(value as u16)
}
const MAX_LOGICAL_TABLES: usize = 4096;
const MAX_LOGICAL_TABLE_ROWS: usize = 65_536;

#[derive(Default)]
struct DrawingStoryCapture<'a> {
    shapes: Vec<crate::Shape<'a>>,
    shape_groups: Vec<crate::ShapeGroup<'a>>,
    drawing_order: Vec<crate::StoryDrawing>,
    story_events: Vec<crate::StoryEvent>,
    story_offset: usize,
}

struct NestedTableBuilder<'a> {
    level: u8,
    table: super::super::table::Table<'a>,
    row: super::super::table::Row<'a>,
    cell_text: SmallVec<[u8; 128]>,
    cell_nested: Vec<crate::CellNestedTable<'a>>,
    cell_drawings: DrawingStoryCapture<'a>,
    cell_story_events: Vec<crate::CellStoryEvent>,
}
impl NestedTableBuilder<'_> {
    fn new(level: u8) -> Self {
        Self {
            level,
            table: super::super::table::Table::new(),
            row: super::super::table::Row::new(),
            cell_text: SmallVec::new(),
            cell_nested: Vec::new(),
            cell_drawings: DrawingStoryCapture::default(),
            cell_story_events: Vec::new(),
        }
    }
}

#[derive(Clone, Copy)]
enum RootDrawingOwner {
    FieldResult,
    NoteSeparator,
    Note,
    HeaderFooter,
    Cell(u8),
    Body,
}

fn associated_font_ref(raw_value: Option<i32>) -> RtfResult<FontRef> {
    let value = raw_value.ok_or_else(|| {
        RtfError::MalformedDocument("RTF af control requires a numeric parameter".to_string())
    })?;
    u16::try_from(value).map_err(|_err| {
        RtfError::MalformedDocument("RTF af value must be in 0..=65535".to_string())
    })
}

fn character_type_selector(
    parameter: Option<i32>,
    name: &str,
    value: CharacterType,
) -> RtfResult<CharacterType> {
    require_parameterless(parameter, name)?;
    Ok(value)
}

fn complex_script_selector(value: Option<i32>) -> RtfResult<bool> {
    match value {
        Some(0) => Ok(false),
        Some(1) => Ok(true),
        _ => Err(RtfError::MalformedDocument(
            "RTF fcs control requires a numeric parameter of 0 or 1".to_string(),
        )),
    }
}

fn character_grid(value: Option<i32>) -> RtfResult<CharacterGrid> {
    match value {
        None => Ok(CharacterGrid::Parameterless),
        Some(raw) => i16::try_from(raw)
            .map(CharacterGrid::Value)
            .map_err(|_err| {
                RtfError::MalformedDocument("RTF cgrid value must be in -32768..=32767".to_string())
            }),
    }
}

fn animated_text(raw_value: Option<i32>) -> RtfResult<AnimatedTextEffect> {
    let value = raw_value.ok_or_else(|| {
        RtfError::MalformedDocument("RTF animtext control requires a numeric parameter".to_string())
    })?;
    AnimatedTextEffect::from_rtf(value).ok_or_else(|| {
        RtfError::MalformedDocument(format!(
            "RTF animtext value must be in 0..={}",
            AnimatedTextEffect::MAX_RTF_VALUE
        ))
    })
}

fn fit_text(raw_value: Option<i32>) -> RtfResult<FitText> {
    let value = raw_value.ok_or_else(|| {
        RtfError::MalformedDocument("RTF fittext control requires a numeric parameter".to_string())
    })?;
    FitText::from_rtf(value).ok_or_else(|| {
        RtfError::MalformedDocument(format!(
            "RTF fittext value must be -1 or 0..={}",
            FitText::MAX_TWIPS
        ))
    })
}

fn emphasis_mark(mark: EmphasisMark, value: Option<i32>) -> RtfResult<EmphasisMark> {
    require_parameterless(value, mark.control_word())?;
    Ok(mark)
}

fn paper_source_bin(raw_value: Option<i32>, name: &str) -> RtfResult<u16> {
    let value = raw_value.ok_or_else(|| {
        RtfError::MalformedDocument(format!("RTF {name} control requires a numeric parameter"))
    })?;
    u16::try_from(value).map_err(|_err| {
        RtfError::MalformedDocument(format!(
            "RTF {name} value must be in 0..={}",
            super::super::section::MAX_SECTION_PAPER_BIN
        ))
    })
}

fn associated_font_size(raw_value: Option<i32>) -> RtfResult<NonZeroU16> {
    let value = raw_value.ok_or_else(|| {
        RtfError::MalformedDocument("RTF afs control requires a numeric parameter".to_string())
    })?;
    let size = u16::try_from(value).map_err(|_err| {
        RtfError::MalformedDocument("RTF afs value must be in 1..=65535".to_string())
    })?;
    NonZeroU16::new(size).ok_or_else(|| {
        RtfError::MalformedDocument("RTF afs value must be in 1..=65535".to_string())
    })
}

fn font_size(raw_value: i32) -> RtfResult<NonZeroU16> {
    let size = u16::try_from(raw_value).map_err(|_err| {
        RtfError::MalformedDocument("RTF fs value must be in 1..=65535".to_string())
    })?;
    NonZeroU16::new(size)
        .ok_or_else(|| RtfError::MalformedDocument("RTF fs value must be in 1..=65535".to_string()))
}

fn associated_language(raw_value: Option<i32>) -> RtfResult<crate::LanguageId> {
    let value = raw_value.ok_or_else(|| {
        RtfError::MalformedDocument("RTF alang control requires a numeric parameter".to_string())
    })?;
    crate::LanguageId::from_rtf(value)
}

fn character_style_reference(raw_value: Option<i32>) -> RtfResult<u16> {
    let value = raw_value.ok_or_else(|| {
        RtfError::MalformedDocument("RTF cs control requires a numeric style handle".to_string())
    })?;
    u16::try_from(value).map_err(|_err| {
        RtfError::MalformedDocument("RTF cs style handle must be in 0..=65535".to_string())
    })
}

fn paragraph_style_reference(raw_value: Option<i32>) -> RtfResult<u16> {
    let value = raw_value.ok_or_else(|| {
        RtfError::MalformedDocument("RTF s control requires a numeric style handle".to_string())
    })?;
    u16::try_from(value).map_err(|_err| {
        RtfError::MalformedDocument("RTF s style handle must be in 0..=65535".to_string())
    })
}

fn section_style_reference(raw_value: Option<i32>) -> RtfResult<u16> {
    let value = raw_value.ok_or_else(|| {
        RtfError::MalformedDocument("RTF ds control requires a numeric style handle".to_string())
    })?;
    u16::try_from(value).map_err(|_err| {
        RtfError::MalformedDocument("RTF ds style handle must be in 0..=65535".to_string())
    })
}

fn table_style_reference(raw_value: Option<i32>) -> RtfResult<u16> {
    let value = raw_value.ok_or_else(|| {
        RtfError::MalformedDocument("RTF ts control requires a numeric style handle".to_string())
    })?;
    u16::try_from(value).map_err(|_err| {
        RtfError::MalformedDocument("RTF ts style handle must be in 0..=65535".to_string())
    })
}

/// Validate a structural revision-author index (`prauth`, `srauth`,
/// `trauth`, or a `\cl*auth` control).
fn nonnegative_author_index(value: i32, name: &str) -> RtfResult<i32> {
    if value < 0 {
        return Err(RtfError::MalformedDocument(format!(
            "RTF {name} revision author index cannot be negative"
        )));
    }
    Ok(value)
}

/// The pending cell revision of the given kind, requiring the matching
/// `\clins`, `\cldel`, or `\clmrgd` marker to have appeared first.
fn pending_cell_revision<'a>(
    state: &'a mut State,
    kind: crate::CellRevisionKind,
    name: &str,
) -> RtfResult<&'a mut crate::CellRevision> {
    state
        .pending_cell_revision
        .as_mut()
        .filter(|revision| revision.kind == kind)
        .ok_or_else(|| {
            RtfError::MalformedDocument(format!(
                "RTF {name} requires a preceding \\{} cell revision marker",
                kind.control_word(),
            ))
        })
}

fn associated_toggle(value: Option<i32>, name: &str) -> RtfResult<bool> {
    match value {
        None | Some(1) => Ok(true),
        Some(0) => Ok(false),
        Some(_) => Err(RtfError::MalformedDocument(format!(
            "RTF {name} toggle parameter must be 0 or 1"
        ))),
    }
}

#[allow(
    clippy::cast_possible_truncation,
    clippy::cast_sign_loss,
    reason = "adjacent range checks bound each narrowing conversion to the target type's range"
)]
fn associated_required_u16(raw_value: Option<i32>, name: &str, maximum: i32) -> RtfResult<u16> {
    let value = raw_value.ok_or_else(|| {
        RtfError::MalformedDocument(format!("RTF {name} control requires a numeric parameter"))
    })?;
    if !(0..=maximum).contains(&value) {
        return Err(RtfError::MalformedDocument(format!(
            "RTF {name} value must be in 0..={maximum}"
        )));
    }
    Ok(value as u16)
}

#[allow(
    clippy::wildcard_enum_match_arm,
    reason = "remaining variants share the same fallback by design"
)]
#[allow(
    clippy::cast_possible_truncation,
    clippy::cast_sign_loss,
    reason = "adjacent range checks bound each narrowing conversion to the target type's range"
)]
fn apply_associated_character_control(
    formatting: &mut AssociatedCharacterFormatting,
    control: &ControlWord<'_>,
) -> RtfResult<bool> {
    use crate::{AssociatedCharacterBaseline as Baseline, AssociatedUnderlineStyle as Underline};

    match control {
        ControlWord::AssociatedBold(value) => {
            formatting.bold = Some(associated_toggle(*value, "ab")?);
        },
        ControlWord::AssociatedAllCaps(value) => {
            formatting.all_caps = Some(associated_toggle(*value, "acaps")?);
        },
        ControlWord::AssociatedColor(value) => {
            formatting.color_ref =
                Some(associated_required_u16(*value, "acf", i32::from(u16::MAX))?);
        },
        ControlWord::AssociatedBaselineDown(value) => {
            let value =
                associated_required_u16(*value, "adn", crate::MAX_CHARACTER_BASELINE_HALF_POINTS)?;
            formatting.baseline = (value != 0).then_some(Baseline::LoweredHalfPoints(value));
        },
        ControlWord::AssociatedExpansion(value) => {
            let expansion = value.ok_or_else(|| {
                RtfError::MalformedDocument(
                    "RTF aexpnd control requires a numeric parameter".to_string(),
                )
            })?;
            if !(-crate::MAX_CHARACTER_EXPANSION..=crate::MAX_CHARACTER_EXPANSION)
                .contains(&expansion)
            {
                return Err(RtfError::MalformedDocument(
                    "RTF aexpnd value must be in -31680..=31680".to_string(),
                ));
            }
            formatting.expansion_quarter_points = Some(expansion as i16);
        },
        ControlWord::AssociatedFontNumber(value) => {
            formatting.font_ref = Some(associated_font_ref(*value)?);
        },
        ControlWord::AssociatedFontSize(value) => {
            formatting.font_size = Some(associated_font_size(*value)?);
        },
        ControlWord::AssociatedItalic(value) => {
            formatting.italic = Some(associated_toggle(*value, "ai")?);
        },
        ControlWord::AssociatedLanguage(value) => {
            formatting.language = Some(associated_language(*value)?);
        },
        ControlWord::AssociatedOutline(value) => {
            formatting.outline = Some(associated_toggle(*value, "aoutl")?);
        },
        ControlWord::AssociatedSmallCaps(value) => {
            formatting.small_caps = Some(associated_toggle(*value, "ascaps")?);
        },
        ControlWord::AssociatedShadow(value) => {
            formatting.shadow = Some(associated_toggle(*value, "ashad")?);
        },
        ControlWord::AssociatedStrike(value) => {
            formatting.strike = Some(associated_toggle(*value, "astrike")?);
        },
        ControlWord::AssociatedUnderline(value) => {
            formatting.underline = Some(if associated_toggle(*value, "aul")? {
                Underline::Single
            } else {
                Underline::None
            });
        },
        ControlWord::AssociatedUnderlineDotted(value) => {
            require_parameterless(*value, "auld")?;
            formatting.underline = Some(Underline::Dotted);
        },
        ControlWord::AssociatedUnderlineDouble(value) => {
            require_parameterless(*value, "auldb")?;
            formatting.underline = Some(Underline::Double);
        },
        ControlWord::AssociatedUnderlineNone(value) => {
            require_parameterless(*value, "aulnone")?;
            formatting.underline = Some(Underline::None);
        },
        ControlWord::AssociatedUnderlineWords(value) => {
            require_parameterless(*value, "aulw")?;
            formatting.underline = Some(Underline::Words);
        },
        ControlWord::AssociatedBaselineUp(value) => {
            let value =
                associated_required_u16(*value, "aup", crate::MAX_CHARACTER_BASELINE_HALF_POINTS)?;
            formatting.baseline = (value != 0).then_some(Baseline::RaisedHalfPoints(value));
        },
        _ => return Ok(false),
    }
    Ok(true)
}

impl Default for State {
    fn default() -> Self {
        Self {
            formatting: Formatting::default(),
            character_border_active: false,
            character_border_seen: 0,
            paragraph: Paragraph::default(),
            paragraph_border_side: None,
            paragraph_border_seen: 0,
            drop_cap_kind: None,
            drop_cap_lines: None,
            pending_tab_alignment: None,
            pending_tab_leader: None,
            unicode_skip: 1,
            in_table: false,
            table_nesting_level: 0,
            cell_boundaries: SmallVec::new(),
            table_style: None,
            table_rsid: None,
            table_row_revision: RevisionMetadata::default(),
            table_row_padding: crate::TableEdgeDistances::default(),
            table_row_spacing: crate::TableEdgeDistances::default(),
            table_row_positioning: crate::FloatingTablePosition::default(),
            table_row_direction: None,
            table_row_layout: crate::TableRowLayout::default(),
            table_row_borders: crate::TableRowBorders::default(),
            table_row_shading: crate::TableShading::default(),
            table_row_geometry: crate::TableRowGeometry::default(),
            table_default_borders: crate::TableStyleDefaultBorders::default(),
            table_default_padding: crate::TableEdgeDistances::default(),
            table_default_spacing: crate::TableEdgeDistances::default(),
            table_default_width_unit: None,
            table_default_width_value: None,
            table_autoformat_flags: crate::TableAutoformatFlags::default(),
            table_row_banding: crate::TableRowBanding::default(),
            table_row_index_seen: false,
            table_row_band_index_seen: false,
            table_last_row_seen: false,
            table_width_unit: None,
            table_width_value: None,
            table_leading_width_unit: None,
            table_leading_width_value: None,
            table_trailing_width_unit: None,
            table_trailing_width_value: None,
            table_indent_value: None,
            table_indent_unit: None,
            pending_cell_padding: crate::TableEdgeDistances::default(),
            pending_cell_spacing: crate::TableEdgeDistances::default(),
            pending_cell_layout: crate::TableCellLayout::default(),
            pending_cell_merge: crate::TableCellMergeState::default(),
            pending_cell_revision: None,
            pending_cell_borders: crate::TableCellBorders::default(),
            pending_cell_shading: crate::TableShading::default(),
            pending_cell_width_unit: None,
            pending_cell_width_value: None,
            table_row_shading_seen: 0,
            pending_cell_shading_seen: 0,
            active_table_border: None,
            active_table_border_seen: 0,
            cell_distances: SmallVec::new(),
            cell_layouts: SmallVec::new(),
            cell_merges: SmallVec::new(),
            cell_revisions: SmallVec::new(),
            cell_decorations: SmallVec::new(),
            cell_widths: SmallVec::new(),
            destination: Destination::DocumentBody,
            visible_section_format: false,
            section_column_number: None,
            encoding: RtfEncoding::Standard(Mbcs::WINDOWS_1252),
            revision_type: None,
            revision_author_id: None,
            revision_date: None,
            revision_event_id: None,
            paragraph_content_started: false,
            paragraph_numbering_declared: false,
        }
    }
}

#[allow(
    clippy::struct_excessive_bools,
    reason = "independent RTF feature flags stay flat for direct access"
)]
/// RTF Parser.
pub(crate) struct Parser<'a> {
    /// Token stream
    tokens: &'a [Token<'a>],
    token_spans: Option<&'a [Range<usize>]>,
    source: Option<&'a str>,
    limits: ParseLimits,
    parse_provenance: crate::validation::ParseProvenance,
    opaque_nodes: Vec<crate::opaque::Node>,
    opaque_bytes: usize,
    unknown_syntax_markers: usize,
    /// Current position in token stream
    pos: usize,
    /// State stack (for handling groups)
    states: Vec<State>,
    /// Font table
    font_table: RefCell<FontTable<'a>>,
    saw_font_table: bool,
    file_table: Option<crate::FileTable<'a>>,
    unicode_alternate_depth: usize,
    /// Color table
    color_table: RefCell<ColorTable>,
    /// Parsed style blocks
    blocks: Vec<StyleBlock<'a>>,
    /// Bounded first-allocation hint derived during structural validation.
    body_block_capacity_hint: usize,
    /// Arena for temporary allocations
    arena: &'a Bump,
    /// Extracted tables
    tables: Vec<super::super::table::Table<'a>>,
    /// Current table being built
    current_table: Option<super::super::table::Table<'a>>,
    /// Current row being built
    current_row: Option<super::super::table::Row<'a>>,
    /// Current cell text buffer
    current_cell_text: SmallVec<[u8; 128]>,
    current_cell_nested: Vec<crate::CellNestedTable<'a>>,
    current_cell_drawings: DrawingStoryCapture<'a>,
    current_cell_story_events: Vec<crate::CellStoryEvent>,
    nested_table_builders: Vec<NestedTableBuilder<'a>>,
    logical_table_count: usize,
    /// Extracted pictures
    pictures: Vec<super::super::picture::Picture<'a>>,
    /// Positional body picture-wrapper records referencing `pictures`.
    picture_compatibility_records: Vec<crate::PictureCompatibilityRecord>,
    /// Extracted fields
    fields: Vec<super::super::field::Field<'a>>,
    field_safety: Vec<crate::validation::FieldSafety>,
    /// Bounded recursion guard for generic fields nested in field results.
    field_nesting_depth: usize,
    field_drawing_captures: Vec<DrawingStoryCapture<'a>>,
    form_fields: Vec<super::super::form_field::FormField<'a>>,
    form_field_text_bytes: usize,
    generator: Option<crate::DocumentGenerator<'a>>,
    revision_save_ids: Vec<u32>,
    saw_revision_save_table: bool,
    revision_save_root: Option<u32>,
    saw_revision_save_root: bool,
    xml_namespaces: Vec<crate::XmlNamespace<'a>>,
    saw_xml_namespace_table: bool,
    xml_namespace_text_bytes: usize,
    custom_xml_tags: Vec<crate::CustomXmlTag<'a>>,
    open_custom_xml_tags: Vec<OpenCustomXmlTag>,
    custom_xml_spans: Vec<CustomXmlSpan>,
    pending_custom_xml_attribute: Option<String>,
    next_custom_xml_order: usize,
    custom_xml_text_bytes: usize,
    math_zones: Vec<crate::MathZone<'a>>,
    math_text_bytes: usize,
    protection_ranges: Vec<crate::ProtectionRange<'a>>,
    open_protection_ranges: HashMap<String, Vec<OpenProtectionRange>>,
    protection_range_spans: Vec<ProtectionRangeSpan>,
    next_protection_range_order: usize,
    editable_regions: Vec<crate::EditableRegion<'a>>,
    open_editable_regions: Vec<OpenEditableRegion>,
    editable_region_spans: Vec<EditableRegionSpan>,
    next_editable_region_order: usize,
    protection_users: Vec<crate::ProtectionUser<'a>>,
    saw_protection_user_table: bool,
    protection_user_text_bytes: usize,
    hyphenation: crate::DocumentHyphenation,
    hyphenation_seen: u8,
    external_references: crate::DocumentExternalReferences<'a>,
    external_reference_spans: crate::metadata::DocumentExternalReferenceSpans,
    document_view: crate::DocumentView,
    document_view_seen: u8,
    review_display: crate::DocumentReviewDisplay,
    review_display_seen: u8,
    window_caption: Option<crate::DocumentWindowCaption<'a>>,
    kinsoku: crate::DocumentKinsoku<'a>,
    xsl_transform: Option<crate::DocumentXslTransform<'a>>,
    xsl_transform_usage: crate::DocumentXslTransformUsage,
    use_xsl_transform_seen: bool,
    style_list_filter: Option<crate::DocumentStyleListFilter>,
    style_sort_method: Option<crate::DocumentStyleSortMethod>,
    style_sort_method_seen: bool,
    save_preferences: crate::DocumentSavePreferences,
    save_preferences_seen: u8,
    write_reservations: crate::DocumentWriteReservations<'a>,
    origin_metadata: crate::DocumentOriginMetadata,
    file_settings: crate::DocumentFileSettings,
    file_settings_seen: u8,
    output_settings: crate::DocumentOutputSettings,
    output_settings_seen: u8,
    rendering_settings: crate::DocumentRenderingSettings,
    rendering_settings_seen: u8,
    processing_settings: crate::DocumentProcessingSettings,
    processing_settings_seen: u8,
    drawing_grid: crate::DocumentDrawingGrid,
    drawing_grid_seen: u8,
    print_layout_settings: crate::DocumentPrintLayoutSettings,
    print_layout_settings_seen: u8,
    section_gutter_overrides: Vec<bool>,
    theme_languages: crate::DocumentThemeLanguages,
    theme_languages_seen: u8,
    xml_policies: crate::DocumentXmlPolicies,
    xml_policies_seen: u8,
    embedding_policies: crate::DocumentEmbeddingPolicies,
    embedding_policies_seen: u8,
    revision_policies: crate::DocumentRevisionPolicies,
    revision_policies_seen: u8,
    style_policies: crate::DocumentStylePolicies,
    style_policies_seen: u8,
    style_restrictions: crate::DocumentStyleRestrictions,
    style_restrictions_seen: u8,
    booklet_printing: crate::DocumentBookletPrinting,
    booklet_printing_seen: u8,
    privacy_policies: crate::DocumentPrivacyPolicies,
    privacy_policies_seen: u8,
    line_spacing_compatibility: crate::DocumentLineSpacingCompatibility,
    line_spacing_compatibility_seen: u8,
    east_asian_compatibility: crate::DocumentEastAsianCompatibility,
    east_asian_compatibility_seen: u8,
    table_layout_compatibility: crate::DocumentTableLayoutCompatibility,
    table_layout_compatibility_seen: u8,
    legacy_layout_compatibility: crate::DocumentLegacyLayoutCompatibility,
    legacy_layout_compatibility_seen: u8,
    asian_grid_compatibility: crate::DocumentAsianGridCompatibility,
    asian_grid_compatibility_seen: u8,
    compatibility_policy: crate::DocumentCompatibilityPolicy,
    compatibility_policy_seen: u8,
    word_2003_compatibility: crate::DocumentWord2003Compatibility,
    word_2003_compatibility_seen: u16,
    theme_data: Option<Vec<u8>>,
    saw_theme_data: bool,
    color_scheme_mapping: Option<Vec<u8>>,
    saw_color_scheme_mapping: bool,
    latent_styles: Option<crate::LatentStyles<'a>>,
    data_store: Option<Vec<u8>>,
    saw_data_store: bool,
    mail_merge: Option<crate::MailMerge<'a>>,
    math_properties: Option<crate::DocumentMathProperties>,
    default_tab_width_twips: Option<u32>,
    language_defaults: crate::DocumentLanguageDefaults,
    default_formatting: crate::DocumentDefaultFormatting,
    default_font_selectors_seen: u8,
    saw_info_group: bool,
    document_direction: Option<TextDirection>,
    gutter_on_right: bool,
    /// Embedded and linked objects
    objects: Vec<super::super::object::EmbeddedObject<'a>>,
    /// Ordered inert document variables
    document_variables: Vec<DocumentVariable<'a>>,
    /// Aggregate decoded document-variable text size
    document_variable_text_bytes: usize,
    /// Ordered inert user-defined properties
    user_properties: Vec<UserProperty<'a>>,
    /// Aggregate decoded user-property text size
    user_property_text_bytes: usize,
    /// Ordered inert index and table-of-contents source marks.
    navigation_entries: Vec<NavigationEntry<'a>>,
    /// Aggregate decoded source-mark text.
    navigation_entry_text_bytes: usize,
    /// Ordered inert generated list markers.
    generated_list_markers: Vec<crate::GeneratedListMarker<'a>>,
    /// Aggregate decoded generated-marker text.
    generated_list_marker_text_bytes: usize,
    /// Whether the unique user-properties destination has been seen
    saw_user_properties: bool,
    /// List table
    list_table: super::super::list::ListTable<'a>,
    saw_list_table: bool,
    /// List override table
    list_override_table: super::super::list::ListOverrideTable,
    saw_list_override_table: bool,
    legacy_section_numbering: crate::LegacySectionNumbering<'a>,
    legacy_paragraph_numbering: Vec<crate::LegacyParagraphNumbering<'a>>,
    paragraph_group_table: Option<crate::ParagraphGroupPropertyTable>,
    /// Sections
    sections: Vec<super::super::section::Section<'a>>,
    /// Whether section-specific properties are currently active.
    section_properties_active: bool,
    /// Whether header/footer or body content has closed the active section-format prefix.
    section_note_options_closed: bool,
    /// Whether the root body is in an explicit late section-format run.
    root_section_format_run: bool,
    /// Bookmarks
    bookmarks: super::super::bookmark::BookmarkTable<'a>,
    /// Open bookmark ranges, indexed by name.
    open_bookmarks: HashMap<String, Vec<OpenBookmark>>,
    /// Completed bookmark ranges awaiting content reconstruction.
    bookmark_spans: Vec<BookmarkSpan>,
    /// UTF-8 byte length of body text emitted into style blocks.
    body_text_len: usize,
    /// Accepted visible root-body paragraph breaks.
    body_paragraph_breaks: usize,
    /// UTF-8 position immediately after the last accepted paragraph break.
    body_after_last_paragraph_break: usize,
    /// Structural paragraph and line breaks retained independently from text.
    body_boundaries: Vec<crate::story::Boundary>,
    /// Stable source order for bookmark ranges.
    next_bookmark_order: usize,
    /// Shapes
    shapes: Vec<super::super::shape::Shape<'a>>,
    /// Exact source order of non-background root drawings in the body story.
    drawing_order: Vec<crate::StoryDrawing>,
    body_story_events: Vec<ParsedBodyStoryEvent>,
    revision_event_indices: Vec<Option<usize>>,
    /// Index in `shapes` owned by the unique document-background destination.
    background_shape_index: Option<usize>,
    /// Inert legacy drawing text boxes.
    legacy_text_boxes: Vec<crate::LegacyTextBox<'a>>,
    /// Inert legacy drawing primitives other than top-level compatibility text boxes.
    legacy_drawings: Vec<crate::LegacyDrawing<'a>>,
    legacy_text_box_text_bytes: usize,
    legacy_drawing_primitives: usize,
    legacy_drawing_points: usize,
    /// Shape groups
    shape_groups: Vec<super::super::shape::ShapeGroup<'a>>,
    /// Stylesheet
    stylesheet: super::super::stylesheet::StyleSheet<'a>,
    /// Whether the unique root stylesheet destination was seen.
    saw_stylesheet: bool,
    /// Document information
    info: super::super::info::DocumentInfo<'a>,
    /// Annotations
    annotations: Vec<super::super::annotation::Annotation<'a>>,
    /// Parsed annotation reference ranges by numeric identifier.
    annotation_ranges: HashMap<i32, (usize, Option<usize>)>,
    /// Author metadata immediately preceding an annotation destination.
    pending_annotation_author: String,
    pending_annotation_author_seen: bool,
    /// Author initials immediately preceding an annotation destination.
    pending_annotation_initials: String,
    pending_annotation_initials_seen: bool,
    pending_annotation_mark: bool,
    /// Footnotes and endnotes
    notes: Vec<super::super::section::Note<'a>>,
    note_options: crate::NoteOptions,
    note_options_closed: bool,
    note_separators: crate::NoteSeparatorTable<'a>,
    current_note_separator_active: bool,
    current_note_separator_elements: Vec<crate::NoteSeparatorElement<'a>>,
    current_note_separator_drawings: DrawingStoryCapture<'a>,
    /// Track changes/revisions
    revisions: Vec<super::super::annotation::Revision<'a>>,
    /// Authors referenced by tracked-change author indices
    revision_authors: Vec<super::super::annotation::RevisionAuthor<'a>>,
    /// Whether the unique revision-author table has been parsed.
    saw_revision_table: bool,
    /// Aggregate decoded author-table text.
    revision_author_text_bytes: usize,
    /// Aggregate decoded tracked-change text.
    revision_text_bytes: usize,
    /// Current header/footer being parsed
    #[allow(dead_code, reason = "retained for header/footer reassembly")]
    current_header_footer: Option<super::super::section::HeaderFooter<'a>>,
    /// Current note being parsed (content buffer)
    current_note_buffer: SmallVec<[u8; 256]>,
    /// Root shapes captured from the active footnote/endnote story.
    current_note_shapes: Vec<super::super::shape::Shape<'a>>,
    /// Root shape groups captured from the active footnote/endnote story.
    current_note_shape_groups: Vec<super::super::shape::ShapeGroup<'a>>,
    current_note_drawing_order: Vec<crate::StoryDrawing>,
    current_note_story_events: Vec<crate::StoryEvent>,
    /// Drawings captured from the active header/footer story.
    current_hf_shapes: Vec<super::super::shape::Shape<'a>>,
    current_hf_shape_groups: Vec<super::super::shape::ShapeGroup<'a>>,
    current_hf_drawing_order: Vec<crate::StoryDrawing>,
    current_hf_story_events: Vec<crate::StoryEvent>,
    current_hf_story_offset: usize,
    /// Current header/footer type being parsed
    current_hf_type: Option<super::super::section::HeaderFooterType>,
}

#[derive(Default)]
struct FormFieldBuilder {
    field_type: Option<super::super::form_field::FormFieldType>,
    text_type: Option<super::super::form_field::FormTextType>,
    name: Option<String>,
    max_length: Option<u16>,
    format: Option<String>,
    default_text: Option<String>,
    default_result: Option<i32>,
    result: Option<i32>,
    half_point_size: Option<i32>,
    protected: Option<bool>,
    calculate_on_exit: Option<bool>,
    size_automatically: Option<bool>,
    own_help: Option<bool>,
    own_status: Option<bool>,
    help_text: Option<String>,
    status_text: Option<String>,
    entry_macro: Option<String>,
    exit_macro: Option<String>,
    list_entries: Vec<String>,
    has_list_box: Option<bool>,
}

#[derive(Default)]
struct LegacyTextBoxBuilder {
    saw_text_box: bool,
    text: Option<String>,
    shapes: Vec<super::super::shape::Shape<'static>>,
    shape_groups: Vec<super::super::shape::ShapeGroup<'static>>,
    drawing_order: Vec<crate::StoryDrawing>,
    story_events: Vec<crate::StoryEvent>,
    horizontal_anchor: Option<crate::LegacyHorizontalAnchor>,
    vertical_anchor: Option<crate::LegacyVerticalAnchor>,
    x: Option<i32>,
    y: Option<i32>,
    width: Option<i32>,
    height: Option<i32>,
    margin: Option<i32>,
    z_order: Option<i32>,
    direction: Option<crate::LegacyTextDirection>,
}

#[derive(Clone, Copy)]
enum LegacySimpleKind {
    Line,
    Rectangle,
    Ellipse,
    Polyline,
    Arc,
}

#[derive(Default)]
struct LegacyColorBuilder {
    gray: Option<i32>,
    red: Option<i32>,
    green: Option<i32>,
    blue: Option<i32>,
    palette: bool,
}

impl LegacyColorBuilder {
    fn is_empty(&self) -> bool {
        self.gray.is_none()
            && self.red.is_none()
            && self.green.is_none()
            && self.blue.is_none()
            && !self.palette
    }

    fn finish(&self, name: &str) -> RtfResult<Option<crate::LegacyDrawingColor>> {
        if let Some(gray) = self.gray {
            if self.red.is_some() || self.green.is_some() || self.blue.is_some() || self.palette {
                return Err(Parser::legacy_error(&format!(
                    "{name} mixes grayscale and RGB"
                )));
            }
            return Ok(Some(crate::LegacyDrawingColor::gray_half_percent(gray)?));
        }
        if self.is_empty() {
            return Ok(None);
        }
        let component = |value: Option<i32>| -> RtfResult<u8> {
            u8::try_from(
                value
                    .ok_or_else(|| Parser::legacy_error(&format!("incomplete {name} RGB color")))?,
            )
            .map_err(|_err| {
                Parser::legacy_error(&format!("{name} RGB component is outside 0..=255"))
            })
        };
        Ok(Some(crate::LegacyDrawingColor::Rgb {
            red: component(self.red)?,
            green: component(self.green)?,
            blue: component(self.blue)?,
            palette: self.palette,
        }))
    }
}

#[derive(Default)]
struct LegacyArrowBuilder {
    fill: Option<crate::LegacyDrawingArrowFill>,
    length: Option<i32>,
    width: Option<i32>,
}

impl LegacyArrowBuilder {
    fn finish(&self, name: &str) -> RtfResult<Option<crate::LegacyDrawingArrow>> {
        if self.fill.is_none() && self.length.is_none() && self.width.is_none() {
            return Ok(None);
        }
        Ok(Some(crate::LegacyDrawingArrow {
            fill: self
                .fill
                .ok_or_else(|| Parser::legacy_error(&format!("incomplete {name} arrow")))?,
            length: crate::LegacyDrawingArrowSize::try_from(
                self.length
                    .ok_or_else(|| Parser::legacy_error(&format!("incomplete {name} arrow")))?,
            )?,
            width: crate::LegacyDrawingArrowSize::try_from(
                self.width
                    .ok_or_else(|| Parser::legacy_error(&format!("incomplete {name} arrow")))?,
            )?,
        }))
    }
}

#[derive(Default)]
struct LegacyPropertiesBuilder {
    line_style: Option<crate::LegacyDrawingLineStyle>,
    line_color: LegacyColorBuilder,
    line_width: Option<i32>,
    fill_foreground: LegacyColorBuilder,
    fill_background: LegacyColorBuilder,
    fill_pattern: Option<i32>,
    start_arrow: LegacyArrowBuilder,
    end_arrow: LegacyArrowBuilder,
    shadow: bool,
    shadow_x: Option<i32>,
    shadow_y: Option<i32>,
}

impl LegacyPropertiesBuilder {
    fn set<T>(slot: &mut Option<T>, value: T, name: &str) -> RtfResult<()> {
        Parser::set_legacy_once(slot, value, name)
    }

    fn flag(slot: &mut bool, name: &str) -> RtfResult<()> {
        if *slot {
            return Err(Parser::legacy_error(&format!("duplicate {name}")));
        }
        *slot = true;
        Ok(())
    }

    #[allow(
        clippy::wildcard_enum_match_arm,
        reason = "remaining variants share the same fallback by design"
    )]
    fn apply(&mut self, control: &ControlWord<'_>) -> RtfResult<bool> {
        match control {
            ControlWord::LegacyDrawingLineStyle(value) => {
                Self::set(&mut self.line_style, *value, "line style")?;
            },
            ControlWord::LegacyDrawingLineGray(value) => {
                Self::set(&mut self.line_color.gray, *value, "line grayscale")?;
            },
            ControlWord::LegacyDrawingLineRed(value) => {
                Self::set(&mut self.line_color.red, *value, "line red")?;
            },
            ControlWord::LegacyDrawingLineGreen(value) => {
                Self::set(&mut self.line_color.green, *value, "line green")?;
            },
            ControlWord::LegacyDrawingLineBlue(value) => {
                Self::set(&mut self.line_color.blue, *value, "line blue")?;
            },
            ControlWord::LegacyDrawingLinePalette => {
                Self::flag(&mut self.line_color.palette, "line palette")?;
            },
            ControlWord::LegacyDrawingLineWidth(value) => {
                Self::set(&mut self.line_width, *value, "line width")?;
            },
            ControlWord::LegacyDrawingFillForegroundGray(value) => Self::set(
                &mut self.fill_foreground.gray,
                *value,
                "foreground grayscale",
            )?,
            ControlWord::LegacyDrawingFillForegroundRed(value) => {
                Self::set(&mut self.fill_foreground.red, *value, "foreground red")?;
            },
            ControlWord::LegacyDrawingFillForegroundGreen(value) => {
                Self::set(&mut self.fill_foreground.green, *value, "foreground green")?;
            },
            ControlWord::LegacyDrawingFillForegroundBlue(value) => {
                Self::set(&mut self.fill_foreground.blue, *value, "foreground blue")?;
            },
            ControlWord::LegacyDrawingFillForegroundPalette => {
                Self::flag(&mut self.fill_foreground.palette, "foreground palette")?;
            },
            ControlWord::LegacyDrawingFillBackgroundGray(value) => Self::set(
                &mut self.fill_background.gray,
                *value,
                "background grayscale",
            )?,
            ControlWord::LegacyDrawingFillBackgroundRed(value) => {
                Self::set(&mut self.fill_background.red, *value, "background red")?;
            },
            ControlWord::LegacyDrawingFillBackgroundGreen(value) => {
                Self::set(&mut self.fill_background.green, *value, "background green")?;
            },
            ControlWord::LegacyDrawingFillBackgroundBlue(value) => {
                Self::set(&mut self.fill_background.blue, *value, "background blue")?;
            },
            ControlWord::LegacyDrawingFillBackgroundPalette => {
                Self::flag(&mut self.fill_background.palette, "background palette")?;
            },
            ControlWord::LegacyDrawingFillPattern(value) => {
                Self::set(&mut self.fill_pattern, *value, "fill pattern")?;
            },
            ControlWord::LegacyDrawingStartArrowFill(value) => {
                Self::set(&mut self.start_arrow.fill, *value, "start arrow fill")?;
            },
            ControlWord::LegacyDrawingStartArrowLength(value) => {
                Self::set(&mut self.start_arrow.length, *value, "start arrow length")?;
            },
            ControlWord::LegacyDrawingStartArrowWidth(value) => {
                Self::set(&mut self.start_arrow.width, *value, "start arrow width")?;
            },
            ControlWord::LegacyDrawingEndArrowFill(value) => {
                Self::set(&mut self.end_arrow.fill, *value, "end arrow fill")?;
            },
            ControlWord::LegacyDrawingEndArrowLength(value) => {
                Self::set(&mut self.end_arrow.length, *value, "end arrow length")?;
            },
            ControlWord::LegacyDrawingEndArrowWidth(value) => {
                Self::set(&mut self.end_arrow.width, *value, "end arrow width")?;
            },
            ControlWord::LegacyDrawingShadow => Self::flag(&mut self.shadow, "shadow")?,
            ControlWord::LegacyDrawingShadowX(value) => {
                Self::set(&mut self.shadow_x, *value, "shadow x")?;
            },
            ControlWord::LegacyDrawingShadowY(value) => {
                Self::set(&mut self.shadow_y, *value, "shadow y")?;
            },
            _ => return Ok(false),
        }
        Ok(true)
    }

    fn finish(self) -> RtfResult<crate::LegacyDrawingProperties> {
        let line_color = self.line_color.finish("line")?;
        let line = if self.line_style.is_some() || line_color.is_some() || self.line_width.is_some()
        {
            let style = self
                .line_style
                .unwrap_or(crate::LegacyDrawingLineStyle::Solid);
            Some(crate::LegacyDrawingLine {
                style,
                color: line_color.unwrap_or(crate::LegacyDrawingColor::Rgb {
                    red: 0,
                    green: 0,
                    blue: 0,
                    palette: false,
                }),
                width: self.line_width.unwrap_or(0),
            })
        } else {
            None
        };
        let foreground = self.fill_foreground.finish("foreground fill")?;
        let background = self.fill_background.finish("background fill")?;
        let fill = if foreground.is_some() || background.is_some() || self.fill_pattern.is_some() {
            Some(crate::LegacyDrawingFill {
                foreground: foreground.unwrap_or(crate::LegacyDrawingColor::Rgb {
                    red: 255,
                    green: 255,
                    blue: 255,
                    palette: false,
                }),
                background: background.unwrap_or(crate::LegacyDrawingColor::Rgb {
                    red: 0,
                    green: 0,
                    blue: 0,
                    palette: false,
                }),
                pattern: crate::LegacyDrawingFillPattern::try_from(
                    self.fill_pattern
                        .ok_or_else(|| Parser::legacy_error("fill colors lack dpfillpat"))?,
                )?,
            })
        } else {
            None
        };
        let shadow = if self.shadow || self.shadow_x.is_some() || self.shadow_y.is_some() {
            if !self.shadow {
                return Err(Parser::legacy_error("shadow offsets lack dpshadow"));
            }
            Some(crate::LegacyDrawingShadow {
                x_offset: self
                    .shadow_x
                    .ok_or_else(|| Parser::legacy_error("dpshadow lacks dpshadx"))?,
                y_offset: self
                    .shadow_y
                    .ok_or_else(|| Parser::legacy_error("dpshadow lacks dpshady"))?,
            })
        } else {
            None
        };
        Ok(crate::LegacyDrawingProperties {
            line,
            fill,
            start_arrow: self.start_arrow.finish("start")?,
            end_arrow: self.end_arrow.finish("end")?,
            shadow,
        })
    }
}

#[derive(Default)]
struct LatentStyleExceptionBuilder {
    locked: Option<bool>,
    semi_hidden: Option<bool>,
    unhide_when_used: Option<bool>,
    quick_format: Option<bool>,
    priority: Option<u8>,
}

/// Parsed RTF document.
///
/// This is an intermediate representation produced by the parser
/// before being converted into the final `RtfDocument` structure.
/// All fields are public to allow direct access during document construction.
pub(crate) struct ParsedDocument<'a> {
    /// Root-level body range proven during the parser's structural preflight.
    pub ordinary_body_source_span: Option<Range<usize>>,
    /// Font table
    pub font_table: FontTable<'a>,
    pub file_table: Option<crate::FileTable<'a>>,
    /// Color table
    pub color_table: ColorTable,
    /// Style blocks
    pub blocks: Vec<StyleBlock<'a>>,
    /// Ordered unsupported syntax retained as inert transport fragments.
    pub opaque_nodes: Vec<crate::opaque::Node>,
    /// Count of content-free markers for syntax not safely interpreted by a
    /// focused destination parser.
    pub unknown_syntax_markers: usize,
    /// Extracted tables
    pub tables: Vec<super::super::table::Table<'a>>,
    /// Extracted pictures
    pub pictures: Vec<super::super::picture::Picture<'a>>,
    /// Positional body picture-wrapper records referencing `pictures`.
    pub picture_compatibility_records: Vec<crate::PictureCompatibilityRecord>,
    /// Extracted fields
    pub fields: Vec<super::super::field::Field<'a>>,
    pub field_safety: Vec<crate::validation::FieldSafety>,
    pub parse_provenance: crate::validation::ParseProvenance,
    pub form_fields: Vec<super::super::form_field::FormField<'a>>,
    pub generator: Option<crate::DocumentGenerator<'a>>,
    pub revision_save: Option<crate::RevisionSaveMetadata>,
    pub xml_namespaces: Vec<crate::XmlNamespace<'a>>,
    pub saw_xml_namespace_table: bool,
    /// Ordered inert custom XML markup tags spanning body text.
    pub custom_xml_tags: Vec<crate::CustomXmlTag<'a>>,
    /// Ordered inert math zones anchored in the body story.
    pub math_zones: Vec<crate::MathZone<'a>>,
    /// Ordered inert protection-exception ranges spanning body text.
    pub protection_ranges: Vec<crate::ProtectionRange<'a>>,
    /// Ordered inert editable regions spanning body text.
    pub editable_regions: Vec<crate::EditableRegion<'a>>,
    pub protection_user_table: Option<crate::ProtectionUserTable<'a>>,
    pub hyphenation: crate::DocumentHyphenation,
    pub external_references: crate::DocumentExternalReferences<'a>,
    pub external_reference_spans: crate::metadata::DocumentExternalReferenceSpans,
    pub document_view: crate::DocumentView,
    pub review_display: crate::DocumentReviewDisplay,
    pub window_caption: Option<crate::DocumentWindowCaption<'a>>,
    /// Inert custom kinsoku character sets.
    pub kinsoku: crate::DocumentKinsoku<'a>,
    pub xsl_transform: Option<crate::DocumentXslTransform<'a>>,
    pub xsl_transform_usage: crate::DocumentXslTransformUsage,
    pub style_list_filter: Option<crate::DocumentStyleListFilter>,
    pub style_sort_method: Option<crate::DocumentStyleSortMethod>,
    pub save_preferences: crate::DocumentSavePreferences,
    pub write_reservations: crate::DocumentWriteReservations<'a>,
    pub origin_metadata: crate::DocumentOriginMetadata,
    pub file_settings: crate::DocumentFileSettings,
    pub output_settings: crate::DocumentOutputSettings,
    pub rendering_settings: crate::DocumentRenderingSettings,
    pub processing_settings: crate::DocumentProcessingSettings,
    pub drawing_grid: crate::DocumentDrawingGrid,
    pub print_layout_settings: crate::DocumentPrintLayoutSettings,
    pub theme_languages: crate::DocumentThemeLanguages,
    pub xml_policies: crate::DocumentXmlPolicies,
    pub embedding_policies: crate::DocumentEmbeddingPolicies,
    pub revision_policies: crate::DocumentRevisionPolicies,
    pub style_policies: crate::DocumentStylePolicies,
    pub style_restrictions: crate::DocumentStyleRestrictions,
    pub booklet_printing: crate::DocumentBookletPrinting,
    pub privacy_policies: crate::DocumentPrivacyPolicies,
    pub line_spacing_compatibility: crate::DocumentLineSpacingCompatibility,
    pub east_asian_compatibility: crate::DocumentEastAsianCompatibility,
    pub table_layout_compatibility: crate::DocumentTableLayoutCompatibility,
    pub legacy_layout_compatibility: crate::DocumentLegacyLayoutCompatibility,
    pub asian_grid_compatibility: crate::DocumentAsianGridCompatibility,
    pub compatibility_policy: crate::DocumentCompatibilityPolicy,
    pub word_2003_compatibility: crate::DocumentWord2003Compatibility,
    pub theme: Option<crate::DocumentTheme<'a>>,
    pub latent_styles: Option<crate::LatentStyles<'a>>,
    pub data_store: Option<crate::DocumentDataStore<'a>>,
    pub mail_merge: Option<crate::MailMerge<'a>>,
    pub math_properties: Option<crate::DocumentMathProperties>,
    pub default_tab_width_twips: Option<u32>,
    pub language_defaults: crate::DocumentLanguageDefaults,
    pub default_formatting: crate::DocumentDefaultFormatting,
    pub document_direction: Option<TextDirection>,
    pub gutter_on_right: bool,
    /// Embedded and linked objects
    pub objects: Vec<super::super::object::EmbeddedObject<'a>>,
    /// Ordered inert document-variable metadata
    pub document_variables: Vec<DocumentVariable<'a>>,
    /// Ordered inert user-defined document properties
    pub user_properties: Vec<UserProperty<'a>>,
    /// Ordered inert index and table-of-contents source marks.
    pub navigation_entries: Vec<NavigationEntry<'a>>,
    /// Ordered inert generated list markers.
    pub generated_list_markers: Vec<crate::GeneratedListMarker<'a>>,
    /// List table
    pub list_table: super::super::list::ListTable<'a>,
    /// List override table
    pub list_override_table: super::super::list::ListOverrideTable,
    pub legacy_section_numbering: crate::LegacySectionNumbering<'a>,
    pub legacy_paragraph_numbering: Vec<crate::LegacyParagraphNumbering<'a>>,
    pub paragraph_group_table: Option<crate::ParagraphGroupPropertyTable>,
    /// Sections
    pub sections: Vec<super::super::section::Section<'a>>,
    /// Bookmarks
    pub bookmarks: super::super::bookmark::BookmarkTable<'a>,
    /// Shapes
    pub shapes: Vec<super::super::shape::Shape<'a>>,
    /// Exact source order of non-background root drawings in the body story.
    pub drawing_order: Vec<crate::StoryDrawing>,
    pub body_paragraph_count: usize,
    pub body_boundaries: Vec<crate::story::Boundary>,
    pub body_story_events: Vec<crate::BodyStoryEvent>,
    /// Index in `shapes` owned by the unique document-background destination.
    pub background_shape_index: Option<usize>,
    /// Inert legacy drawing text boxes.
    pub legacy_text_boxes: Vec<crate::LegacyTextBox<'a>>,
    pub legacy_drawings: Vec<crate::LegacyDrawing<'a>>,
    /// Shape groups
    pub shape_groups: Vec<super::super::shape::ShapeGroup<'a>>,
    /// Stylesheet
    pub stylesheet: super::super::stylesheet::StyleSheet<'a>,
    /// Document information
    pub info: super::super::info::DocumentInfo<'a>,
    /// Annotations
    pub annotations: Vec<super::super::annotation::Annotation<'a>>,
    /// Footnotes and endnotes
    pub notes: Vec<super::super::section::Note<'a>>,
    /// Explicit document-level footnote and endnote configuration.
    pub note_options: crate::NoteOptions,
    pub note_separators: crate::NoteSeparatorTable<'a>,
    /// Track changes/revisions
    pub revisions: Vec<super::super::annotation::Revision<'a>>,
    /// Ordered inert revision-author table.
    pub revision_authors: Vec<super::super::annotation::RevisionAuthor<'a>>,
}

mod annotations;
mod api;
mod content;
mod controls;
mod dispatch;
mod document;
mod fields;
mod groups;
mod info;
mod legacy_drawing;
mod markup;
mod math;
mod navigation;
mod numbering;
mod objects;
mod pictures;
mod properties;
mod protection;
mod resources;
mod section;
mod shapes;
mod stories;
mod styles;
mod tables;
mod unicode;

#[cfg(test)]
mod tests {
    use super::{
        BODY_BLOCK_RESERVE_SOURCE_MULTIPLIER, ControlWord, MAX_BODY_BLOCK_RESERVE_BYTES,
        MIN_BODY_BLOCK_RESERVE_SOURCE_BYTES, ParseLimits, Parser, RtfError, StyleBlock,
        append_transport_bytes, disables_body_block_reservation, initial_body_block_capacity,
    };
    use crate::codec::lexer::Lexer;
    use bumpalo::Bump;
    use std::mem::size_of;

    #[derive(Default)]
    struct CountingBuffer {
        bytes: Vec<u8>,
        extend_calls: usize,
    }

    impl Extend<u8> for CountingBuffer {
        fn extend<T: IntoIterator<Item = u8>>(&mut self, iter: T) {
            self.extend_calls += 1;
            self.bytes.extend(iter);
        }
    }

    #[test]
    fn ascii_transport_is_extended_in_one_batch() {
        let mut buffer = CountingBuffer::default();

        append_transport_bytes(&mut buffer, "ASCII transport").expect("ASCII transport bytes");

        assert_eq!(buffer.bytes, b"ASCII transport");
        assert_eq!(buffer.extend_calls, 1);
    }

    #[test]
    fn byte_valued_non_ascii_transport_keeps_the_fallback() {
        let mut buffer = CountingBuffer::default();

        append_transport_bytes(&mut buffer, "\u{e9}").expect("byte-valued transport character");

        assert_eq!(buffer.bytes, [0xe9]);
        assert_eq!(buffer.extend_calls, 1);
    }

    #[test]
    fn invalid_unicode_keeps_the_valid_prefix_and_error() {
        let mut buffer = Vec::new();

        let error = append_transport_bytes(&mut buffer, "ok\u{100}")
            .expect_err("non-byte Unicode must be rejected");

        assert_eq!(buffer, b"ok");
        assert!(matches!(error, RtfError::InvalidUnicode(_)));
    }

    #[test]
    fn body_block_capacity_is_bounded_by_source_candidates_and_limits() {
        assert_eq!(initial_body_block_capacity(None, 10, 10), 0);
        assert_eq!(initial_body_block_capacity(Some(10), 0, 10), 0);
        assert_eq!(
            initial_body_block_capacity(Some(MIN_BODY_BLOCK_RESERVE_SOURCE_BYTES - 1), 10, 10),
            0
        );

        let enough_source = MIN_BODY_BLOCK_RESERVE_SOURCE_BYTES.max(
            size_of::<StyleBlock<'static>>()
                .saturating_mul(10)
                .div_ceil(BODY_BLOCK_RESERVE_SOURCE_MULTIPLIER),
        );
        assert_eq!(initial_body_block_capacity(Some(enough_source), 5, 10), 5);
        assert_eq!(initial_body_block_capacity(Some(enough_source), 10, 3), 3);

        let hard_capacity = MAX_BODY_BLOCK_RESERVE_BYTES / size_of::<StyleBlock<'static>>();
        assert_eq!(
            initial_body_block_capacity(Some(usize::MAX), usize::MAX, usize::MAX),
            hard_capacity
        );
    }

    #[test]
    fn table_and_deletion_controls_disable_body_block_reservation() {
        for control in [
            ControlWord::TableNestingLevel(Some(2)),
            ControlWord::TableRowDefaults,
            ControlWord::TableRow,
            ControlWord::TableCell,
            ControlWord::InTable,
            ControlWord::CellX(1_000),
            ControlWord::NestedTableCell(None),
            ControlWord::NestedTableRow(None),
            ControlWord::NestedTableProperties(None),
            ControlWord::Deleted(true),
        ] {
            assert!(disables_body_block_reservation(&control));
        }
        assert!(!disables_body_block_reservation(&ControlWord::Par));
        assert!(!disables_body_block_reservation(&ControlWord::Deleted(
            false
        )));
    }

    #[test]
    fn dense_plain_body_uses_the_bounded_preflight_capacity() {
        const PARAGRAPHS: usize = 1_536;
        let mut source = String::from("{\\rtf1 ");
        for index in 0..PARAGRAPHS {
            if index != 0 {
                source.push_str("\\par ");
            }
            source.push_str("paragraph-abcdefghijklmnopqrstuvwxyz-0123456789-");
            source.push_str(&index.to_string());
        }
        source.push('}');

        let arena = Bump::new();
        let limits = ParseLimits::default();
        let (tokens, spans) = Lexer::new_with_limits(&source, &arena, limits)
            .tokenize_with_spans()
            .expect("dense body tokens");
        let expected_capacity =
            initial_body_block_capacity(Some(source.len()), PARAGRAPHS, tokens.len());
        assert_eq!(expected_capacity, PARAGRAPHS);

        let parsed = Parser::new_with_source(&tokens, &spans, &source, &arena, limits)
            .parse()
            .expect("dense body parse");

        assert_eq!(parsed.blocks.len(), PARAGRAPHS);
        assert!(parsed.blocks.capacity() >= expected_capacity);
        assert_eq!(
            parsed.blocks.first().map(|block| block.text.as_ref()),
            Some("paragraph-abcdefghijklmnopqrstuvwxyz-0123456789-0")
        );
        assert!(
            parsed
                .blocks
                .last()
                .is_some_and(|block| block.text.ends_with("1535"))
        );
    }

    #[test]
    fn oversized_optional_hint_falls_back_to_one_block() {
        let source = r"{\rtf1 body}";
        let arena = Bump::new();
        let limits = ParseLimits::default();
        let (tokens, spans) = Lexer::new_with_limits(source, &arena, limits)
            .tokenize_with_spans()
            .expect("body tokens");
        let mut parser = Parser::new_with_source(&tokens, &spans, source, &arena, limits);
        parser.body_block_capacity_hint = usize::MAX;

        parser
            .prepare_body_block_push()
            .expect("one-block fallback reserve");

        assert!(parser.blocks.capacity() >= 1);
        assert_eq!(parser.body_block_capacity_hint, 0);
    }

    #[test]
    fn table_heavy_body_keeps_lazy_block_growth() {
        const CELLS: usize = 256;
        let filler = "x".repeat(256);
        let mut source = String::from("{\\rtf1\\trowd");
        for cell in 1..=CELLS {
            source.push_str("\\cellx");
            source.push_str(&(cell * 100).to_string());
        }
        source.push_str("\\intbl ");
        for cell in 0..CELLS {
            source.push_str("cell-");
            source.push_str(&cell.to_string());
            source.push_str(&filler);
            source.push_str("\\cell ");
        }
        source.push_str("\\row\\pard body}");

        let arena = Bump::new();
        let limits = ParseLimits::default();
        let (tokens, spans) = Lexer::new_with_limits(&source, &arena, limits)
            .tokenize_with_spans()
            .expect("table-heavy tokens");
        let unsafe_capacity =
            initial_body_block_capacity(Some(source.len()), CELLS + 1, tokens.len());
        let parsed = Parser::new_with_source(&tokens, &spans, &source, &arena, limits)
            .parse()
            .expect("table-heavy parse");

        assert_eq!(parsed.blocks.len(), 1);
        assert_eq!(parsed.blocks[0].text, "body");
        assert!(parsed.blocks.capacity() < unsafe_capacity);
        assert_eq!(parsed.tables.len(), 1);
        assert_eq!(parsed.tables[0].rows()[0].cells().len(), CELLS);
    }

    #[test]
    fn deletion_heavy_body_keeps_lazy_block_growth() {
        const DELETED_RUNS: usize = 256;
        let filler = "x".repeat(256);
        let mut source = String::from("{\\rtf1{\\*\\revtbl Unknown;}\\deleted\\revauthdel0 ");
        for run in 0..DELETED_RUNS {
            source.push_str("hidden-");
            source.push_str(&run.to_string());
            source.push_str(&filler);
            source.push_str(if run % 2 == 0 { "\\b " } else { "\\b0 " });
        }
        source.push_str("\\deleted0 body}");

        let arena = Bump::new();
        let limits = ParseLimits::default();
        let (tokens, spans) = Lexer::new_with_limits(&source, &arena, limits)
            .tokenize_with_spans()
            .expect("deletion-heavy tokens");
        let unsafe_capacity =
            initial_body_block_capacity(Some(source.len()), DELETED_RUNS + 1, tokens.len());
        let parsed = Parser::new_with_source(&tokens, &spans, &source, &arena, limits)
            .parse()
            .expect("deletion-heavy parse");

        assert_eq!(parsed.blocks.len(), 1);
        assert_eq!(parsed.blocks[0].text, "body");
        assert!(parsed.blocks.capacity() < unsafe_capacity);
        assert_eq!(parsed.revisions.len(), 1);
        let expected_deleted_len = (0..DELETED_RUNS)
            .map(|run| "hidden-".len() + run.to_string().len() + filler.len())
            .sum::<usize>();
        assert_eq!(parsed.revisions[0].content.len(), expected_deleted_len);
    }
}
