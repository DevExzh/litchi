//! RTF parser that builds document structure from tokens.

use super::error::{RtfError, RtfResult};
use super::lexer::{ControlWord, Token};
use super::types::*;
use bumpalo::Bump;
use encoding_rs::Encoding;
use litchi_core::encoding::codepage_to_encoding;
use smallvec::SmallVec;
use std::borrow::Cow;
use std::cell::RefCell;
use std::collections::HashMap;
use std::num::NonZeroU16;

#[derive(Debug, Clone, Copy)]
enum RtfEncoding {
    Standard(&'static Encoding),
    Cp437,
    Cp850,
}

fn strict_paragraph_toggle(
    value: Option<i32>,
    name: &str,
) -> std::result::Result<bool, RtfError> {
    match value {
        None | Some(1) => Ok(true),
        Some(0) => Ok(false),
        Some(_) => Err(RtfError::MalformedDocument(format!(
            "RTF {name} accepts only 0 or 1"
        ))),
    }
}

fn strict_paragraph_selector(
    value: Option<i32>,
    name: &str,
) -> std::result::Result<(), RtfError> {
    if value.is_some() {
        return Err(RtfError::MalformedDocument(format!(
            "RTF {name} must not have a numeric parameter"
        )));
    }
    Ok(())
}

fn required_paragraph_bool(value: Option<i32>, name: &str) -> std::result::Result<bool, RtfError> {
    match value {
        Some(0) => Ok(false),
        Some(1) => Ok(true),
        None => Err(RtfError::MalformedDocument(format!("RTF {name} requires 0 or 1"))),
        Some(_) => Err(RtfError::MalformedDocument(format!("RTF {name} accepts only 0 or 1"))),
    }
}

fn required_list_spacing(value: Option<i32>, name: &str) -> std::result::Result<u32, RtfError> {
    let value = value.ok_or_else(|| RtfError::MalformedDocument(format!("RTF {name} requires a numeric parameter")))?;
    u32::try_from(value).ok().filter(|value| *value <= 1_000_000).ok_or_else(|| {
        RtfError::MalformedDocument(format!("RTF {name} must be in 0..=1000000"))
    })
}
fn required_paragraph_indent(value:Option<i32>,name:&str)->std::result::Result<i32,RtfError>{let value=value.ok_or_else(||RtfError::MalformedDocument(format!("RTF {name} requires a numeric parameter")))?;if value.unsigned_abs()>10_000_000{return Err(RtfError::MalformedDocument(format!("RTF {name} exceeds the supported range")))}Ok(value)}

impl RtfEncoding {
    fn decode(self, bytes: &[u8]) -> Cow<'_, str> {
        match self {
            Self::Standard(encoding) => encoding.decode(bytes).0,
            Self::Cp437 => decode_dos_codepage(bytes, &CP437_HIGH),
            Self::Cp850 => decode_dos_codepage(bytes, &CP850_HIGH),
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
            output.push(high[usize::from(byte - 0x80)]);
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
    for character in text.chars() {
        let byte = u8::try_from(character as u32).map_err(|_| {
            RtfError::InvalidUnicode(
                "RTF source text is not a byte-preserving transport string".to_string(),
            )
        })?;
        buffer.extend(std::iter::once(byte));
    }
    Ok(())
}

fn control_symbol_text(control: &ControlWord<'_>) -> Option<&'static str> {
    match control {
        ControlWord::NonBreakingSpace => Some("\u{00A0}"),
        ControlWord::OptionalHyphen => Some("\u{00AD}"),
        ControlWord::NonBreakingHyphen => Some("\u{2011}"),
        _ => None,
    }
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
    #[allow(dead_code)]
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
const MAX_ANNOTATIONS: usize = super::annotation::MAX_ANNOTATIONS;
const MAX_ANNOTATION_TEXT_BYTES: usize = 4 * 1_048_576;
const MAX_SECTIONS: usize = 4_096;
use super::stylesheet::{MAX_STYLES, MAX_STYLE_NAME_BYTES};
use super::list::{MAX_LISTS, MAX_LIST_LEVELS, MAX_LIST_TEXT_BYTES, MAX_LIST_TABS};
const MAX_REVISION_AUTHORS: usize = super::annotation::MAX_REVISION_AUTHORS;
const MAX_REVISION_AUTHOR_BYTES: usize = 65_536;
const MAX_REVISIONS: usize = super::annotation::MAX_REVISIONS;
const MAX_SHAPES: usize = 65_536;
const MAX_SHAPE_GROUPS: usize = 16_384;
const MAX_SHAPES_PER_GROUP: usize = 65_536;
const MAX_GROUPS_PER_GROUP: usize = 16_384;
const MAX_SHAPE_GROUP_DEPTH: usize = 64;
const MAX_SHAPE_PROPERTIES: usize = 65_536;
const MAX_SHAPE_PROPERTY_BYTES: usize = 1_048_576;
const MAX_SHAPE_TEXT_BYTES: usize = 16 * 1_048_576;
const MAX_OBJECTS: usize = 65_536;
const MAX_OBJECT_TEXT_BYTES: usize = 1_048_576;
const MAX_OBJECT_DATA_BYTES: usize = 256 * 1_048_576;
const MAX_PICTURE_DATA_BYTES: usize = 256 * 1_048_576;
use super::document_variable::{
    DocumentVariable, MAX_DOCUMENT_VARIABLES, MAX_DOCUMENT_VARIABLE_NAME_BYTES,
    MAX_DOCUMENT_VARIABLE_TEXT_BYTES, MAX_DOCUMENT_VARIABLE_VALUE_BYTES,
};
use super::user_property::{
    MAX_USER_PROPERTIES, MAX_USER_PROPERTY_NAME_BYTES, MAX_USER_PROPERTY_TEXT_BYTES,
    MAX_USER_PROPERTY_VALUE_BYTES, UserProperty, UserPropertyValue,
};
use super::navigation_entry::{
    IndexEntry, IndexPageReference, MAX_NAVIGATION_ENTRIES, MAX_NAVIGATION_ENTRY_DEPTH,
    MAX_NAVIGATION_ENTRY_TEXT_BYTES, MAX_NAVIGATION_ENTRY_TEXT_TOTAL_BYTES, NavigationEntry,
    TableOfContentsEntry,
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
    /// Optional alignment selector for the next `tx` tab terminator.
    pending_tab_alignment: Option<super::border::TabAlignment>,
    /// Optional leader selector for the next `tx` or `tb` tab terminator.
    pending_tab_leader: Option<super::border::TabLeader>,
    /// Unicode skip count (characters to skip after \u)
    unicode_skip: i32,
    /// Whether we're inside a table
    in_table: bool,
    table_nesting_level: u8,
    /// Cell boundaries for current row (in twips)
    cell_boundaries: SmallVec<[i32; 8]>,
    table_row_padding: crate::TableEdgeDistances,
    table_row_spacing: crate::TableEdgeDistances,
    table_row_positioning: crate::FloatingTablePosition,
    table_row_direction: Option<TextDirection>,
    table_row_layout: crate::TableRowLayout,
    table_row_borders: crate::TableRowBorders,
    table_row_shading: crate::TableShading,
    table_row_geometry: crate::TableRowGeometry,
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
    revision_type: Option<super::annotation::RevisionType>,
    /// Revision-author table index
    revision_author_id: Option<i32>,
    /// Packed RTF revision timestamp
    revision_date: Option<i32>,
}

fn is_section_control(control: &ControlWord<'_>) -> bool {
    matches!(
        control,
        ControlWord::SectionDefault
            | ControlWord::SectionBreak
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
            | ControlWord::PageNumberDecimal
            | ControlWord::PageNumberUpperRoman
            | ControlWord::PageNumberLowerRoman
            | ControlWord::PageNumberUpperLetter
            | ControlWord::PageNumberLowerLetter
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
            | ControlWord::SectionFootnotePlacement(_)
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

fn apply_table_distance(state:&mut State,target:crate::TableDistanceTarget,parameter:Option<i32>,unit:bool)->RtfResult<()> {
    let value=parameter.ok_or_else(||RtfError::MalformedDocument("RTF table distance control requires a parameter".to_string()))?;
    let distances=match (target.scope,target.kind){(crate::TableDistanceScope::Row,crate::TableDistanceKind::Padding)=>&mut state.table_row_padding,(crate::TableDistanceScope::Row,crate::TableDistanceKind::Spacing)=>&mut state.table_row_spacing,(crate::TableDistanceScope::Cell,crate::TableDistanceKind::Padding)=>&mut state.pending_cell_padding,(crate::TableDistanceScope::Cell,crate::TableDistanceKind::Spacing)=>&mut state.pending_cell_spacing};
    let side=distances.side_mut(target.edge);if unit{side.unit=Some(match value{0=>crate::TableDistanceUnit::Null,3=>crate::TableDistanceUnit::Twips,_=>return Err(RtfError::MalformedDocument("RTF table distance unit must be 0 or 3".to_string()))});}else{if !(0..=crate::MAX_TABLE_DISTANCE_TWIPS).contains(&value){return Err(RtfError::MalformedDocument("RTF table distance value is out of range".to_string()))}side.value=Some(value as u16);}Ok(())
}

fn table_width_unit(parameter:Option<i32>)->RtfResult<crate::TablePreferredWidthUnit>{match parameter{Some(0)=>Ok(crate::TablePreferredWidthUnit::Null),Some(1)=>Ok(crate::TablePreferredWidthUnit::Auto),Some(2)=>Ok(crate::TablePreferredWidthUnit::Percent),Some(3)=>Ok(crate::TablePreferredWidthUnit::Twips),None=>Err(RtfError::MalformedDocument("RTF preferred-width unit requires a parameter".to_string())),Some(_)=>Err(RtfError::MalformedDocument("RTF preferred-width unit must be in 0..=3".to_string()))}}
fn table_indent_unit(parameter:Option<i32>)->RtfResult<crate::TableIndentUnit>{match parameter{Some(0)=>Ok(crate::TableIndentUnit::Auto),Some(1)=>Ok(crate::TableIndentUnit::Twips),Some(2)=>Ok(crate::TableIndentUnit::Nil),Some(3)=>Ok(crate::TableIndentUnit::Percent),None=>Err(RtfError::MalformedDocument("RTF tblindtype requires a parameter".to_string())),Some(_)=>Err(RtfError::MalformedDocument("RTF tblindtype must be in 0..=3".to_string()))}}
fn resolve_preferred_width(unit:Option<crate::TablePreferredWidthUnit>,value:Option<i32>)->RtfResult<Option<crate::TablePreferredWidth>>{let Some(unit)=unit else{return if value.is_none(){Ok(None)}else{Err(RtfError::MalformedDocument("RTF preferred-width value lacks its unit control".to_string()))}};let value=match unit{crate::TablePreferredWidthUnit::Null|crate::TablePreferredWidthUnit::Auto=>{if value.is_some_and(|value|value!=0){return Err(RtfError::MalformedDocument("RTF null or auto preferred width must omit its value or use zero".to_string()))}None},crate::TablePreferredWidthUnit::Percent=>Some(required_table_value(value,"preferred width percentage",crate::MAX_TABLE_WIDTH_PERCENT as u16)?),crate::TablePreferredWidthUnit::Twips=>Some(required_table_value(value,"preferred width",crate::MAX_TABLE_GEOMETRY_TWIPS as u16)?)};Ok(Some(crate::TablePreferredWidth::new(unit,value)?))}
fn resolve_invisible_width(unit:Option<crate::TablePreferredWidthUnit>,value:Option<i32>,side:&str)->RtfResult<Option<crate::TablePreferredWidth>>{let Some(unit)=unit else{return if value.is_none(){Ok(None)}else{Err(RtfError::MalformedDocument(format!("RTF {side} invisible-width value lacks its unit control")))}};let value=match unit{crate::TablePreferredWidthUnit::Null|crate::TablePreferredWidthUnit::Auto=>{if value.is_some_and(|value|value!=0){return Err(RtfError::MalformedDocument(format!("RTF null or auto {side} invisible width must omit its value or use zero")))}None},crate::TablePreferredWidthUnit::Percent=>Some(required_table_value(Some(value.unwrap_or(0)),&format!("{side} invisible width percentage"),crate::MAX_TABLE_WIDTH_PERCENT as u16)?),crate::TablePreferredWidthUnit::Twips=>Some(required_table_value(Some(value.unwrap_or(0)),&format!("{side} invisible width"),crate::MAX_TABLE_GEOMETRY_TWIPS as u16)?)};Ok(Some(crate::TablePreferredWidth::new(unit,value)?))}
fn resolve_row_geometry(state:&State)->RtfResult<crate::TableRowGeometry>{let mut geometry=state.table_row_geometry;geometry.set_preferred_width(resolve_preferred_width(state.table_width_unit,state.table_width_value)?);geometry.set_leading_invisible_width(resolve_invisible_width(state.table_leading_width_unit,state.table_leading_width_value,"leading")?);geometry.set_trailing_invisible_width(resolve_invisible_width(state.table_trailing_width_unit,state.table_trailing_width_value,"trailing")?);if state.table_indent_value.is_some()||state.table_indent_unit.is_some(){let unit=state.table_indent_unit.unwrap_or(crate::TableIndentUnit::Twips);geometry.set_indent(Some(crate::TableIndent::new(unit,state.table_indent_value.unwrap_or(0))?));}geometry.validate()?;Ok(geometry)}
fn table_geometry_twips(parameter:Option<i32>,name:&str,signed:bool)->RtfResult<i32>{let value=parameter.ok_or_else(||RtfError::MalformedDocument(format!("RTF {name} requires a parameter")))?;let valid=if signed{value.unsigned_abs()<=crate::MAX_TABLE_GEOMETRY_TWIPS as u32}else{(0..=crate::MAX_TABLE_GEOMETRY_TWIPS).contains(&value)};if valid{Ok(value)}else{Err(RtfError::MalformedDocument(format!("RTF {name} is out of range")))}}
fn table_row_height(parameter:Option<i32>)->RtfResult<crate::TableRowHeight>{let value=parameter.ok_or_else(||RtfError::MalformedDocument("RTF trrh requires a parameter".to_string()))?;if value.unsigned_abs()>crate::MAX_TABLE_GEOMETRY_TWIPS as u32{return Err(RtfError::MalformedDocument("RTF trrh is out of range".to_string()))}Ok(if value==0{crate::TableRowHeight::Automatic}else if value>0{crate::TableRowHeight::Minimum(value as u16)}else{crate::TableRowHeight::Exact(value.unsigned_abs() as u16)})}

fn require_parameterless(parameter:Option<i32>,name:&str)->RtfResult<()>{if parameter.is_some(){return Err(RtfError::MalformedDocument(format!("RTF {name} does not accept a parameter")))}Ok(())}

fn required_table_value(value:Option<i32>,name:&str,maximum:u16)->RtfResult<u16>{let value=value.ok_or_else(||RtfError::MalformedDocument(format!("RTF {name} requires a numeric parameter")))?;let value=u16::try_from(value).map_err(|_|RtfError::MalformedDocument(format!("RTF {name} value must be in 0..={maximum}")))?;if value>maximum{return Err(RtfError::MalformedDocument(format!("RTF {name} value must be in 0..={maximum}")))}Ok(value)}
fn floating_table_offset(parameter:Option<i32>,negative:bool,axis:&str)->RtfResult<i32>{let value=parameter.ok_or_else(||RtfError::MalformedDocument(format!("RTF floating-table {axis} offset requires a parameter")))?;let valid=if negative{(-crate::MAX_FLOATING_TABLE_DISTANCE_TWIPS..=-1).contains(&value)}else{(0..=crate::MAX_FLOATING_TABLE_DISTANCE_TWIPS).contains(&value)};if !valid{return Err(RtfError::MalformedDocument(format!("RTF floating-table {axis} offset is out of range")))}Ok(value)}
fn floating_table_wrap_distance(parameter:Option<i32>)->RtfResult<u16>{let value=parameter.ok_or_else(||RtfError::MalformedDocument("RTF floating-table wrap distance requires a parameter".to_string()))?;if !(0..=crate::MAX_FLOATING_TABLE_DISTANCE_TWIPS).contains(&value){return Err(RtfError::MalformedDocument("RTF floating-table wrap distance is out of range".to_string()))}Ok(value as u16)}
const MAX_LOGICAL_TABLES:usize=4096;
const MAX_LOGICAL_TABLE_ROWS:usize=65_536;

struct NestedTableBuilder<'a>{level:u8,table:super::table::Table<'a>,row:super::table::Row<'a>,cell_text:SmallVec<[u8;128]>,cell_nested:Vec<crate::CellNestedTable<'a>>}
impl<'a> NestedTableBuilder<'a>{fn new(level:u8)->Self{Self{level,table:super::table::Table::new(),row:super::table::Row::new(),cell_text:SmallVec::new(),cell_nested:Vec::new()}}}

fn associated_font_ref(value: Option<i32>) -> RtfResult<FontRef> {
    let value = value.ok_or_else(|| {
        RtfError::MalformedDocument("RTF af control requires a numeric parameter".to_string())
    })?;
    u16::try_from(value).map_err(|_| {
        RtfError::MalformedDocument("RTF af value must be in 0..=65535".to_string())
    })
}

fn associated_font_size(value: Option<i32>) -> RtfResult<NonZeroU16> {
    let value = value.ok_or_else(|| {
        RtfError::MalformedDocument("RTF afs control requires a numeric parameter".to_string())
    })?;
    let value = u16::try_from(value).map_err(|_| {
        RtfError::MalformedDocument("RTF afs value must be in 1..=65535".to_string())
    })?;
    NonZeroU16::new(value).ok_or_else(|| {
        RtfError::MalformedDocument("RTF afs value must be in 1..=65535".to_string())
    })
}

fn associated_language(value: Option<i32>) -> RtfResult<crate::LanguageId> {
    let value = value.ok_or_else(|| {
        RtfError::MalformedDocument("RTF alang control requires a numeric parameter".to_string())
    })?;
    crate::LanguageId::from_rtf(value)
}

impl Default for State {
    fn default() -> Self {
        Self {
            formatting: Formatting::default(),
            character_border_active: false,
            character_border_seen: 0,
            paragraph: Paragraph::default(),
            pending_tab_alignment: None,
            pending_tab_leader: None,
            unicode_skip: 1,
            in_table: false,
            table_nesting_level: 0,
            cell_boundaries: SmallVec::new(),
            table_row_padding: Default::default(), table_row_spacing: Default::default(), table_row_positioning: Default::default(), table_row_direction:None, table_row_layout:Default::default(), table_row_borders:Default::default(), table_row_shading:Default::default(), table_row_geometry:Default::default(), table_width_unit:None, table_width_value:None, table_leading_width_unit:None, table_leading_width_value:None, table_trailing_width_unit:None, table_trailing_width_value:None, table_indent_value:None, table_indent_unit:None,
            pending_cell_padding: Default::default(), pending_cell_spacing: Default::default(), pending_cell_layout:Default::default(), pending_cell_merge:Default::default(), pending_cell_borders:Default::default(), pending_cell_shading:Default::default(), pending_cell_width_unit:None, pending_cell_width_value:None, table_row_shading_seen:0, pending_cell_shading_seen:0, active_table_border:None, active_table_border_seen:0, cell_distances: SmallVec::new(), cell_layouts:SmallVec::new(), cell_merges:SmallVec::new(), cell_decorations:SmallVec::new(), cell_widths:SmallVec::new(),
            destination: Destination::DocumentBody,
            visible_section_format: false,
            section_column_number: None,
            encoding: RtfEncoding::Standard(encoding_rs::WINDOWS_1252),
            revision_type: None,
            revision_author_id: None,
            revision_date: None,
        }
    }
}

/// RTF Parser.
pub struct Parser<'a> {
    /// Token stream
    tokens: &'a [Token<'a>],
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
    /// Arena for temporary allocations
    arena: &'a Bump,
    /// Extracted tables
    tables: Vec<super::table::Table<'a>>,
    /// Current table being built
    current_table: Option<super::table::Table<'a>>,
    /// Current row being built
    current_row: Option<super::table::Row<'a>>,
    /// Current cell text buffer
    current_cell_text: SmallVec<[u8; 128]>,
    current_cell_nested: Vec<crate::CellNestedTable<'a>>,
    nested_table_builders: Vec<NestedTableBuilder<'a>>,
    logical_table_count: usize,
    /// Extracted pictures
    pictures: Vec<super::picture::Picture<'a>>,
    /// Extracted fields
    fields: Vec<super::field::Field<'a>>,
    form_fields: Vec<super::form_field::FormField<'a>>,
    form_field_text_bytes: usize,
    generator: Option<crate::DocumentGenerator<'a>>,
    revision_save_ids: Vec<u32>,
    saw_revision_save_table: bool,
    revision_save_root: Option<u32>,
    saw_revision_save_root: bool,
    xml_namespaces: Vec<crate::XmlNamespace<'a>>,
    saw_xml_namespace_table: bool,
    xml_namespace_text_bytes: usize,
    theme_data: Option<Vec<u8>>,
    saw_theme_data: bool,
    color_scheme_mapping: Option<Vec<u8>>,
    saw_color_scheme_mapping: bool,
    latent_styles: Option<crate::LatentStyles<'a>>,
    data_store: Option<Vec<u8>>,
    saw_data_store: bool,
    math_properties: Option<crate::DocumentMathProperties>,
    language_defaults: crate::DocumentLanguageDefaults,
    saw_info_group: bool,
    document_direction: Option<crate::TextDirection>,
    gutter_on_right: bool,
    /// Embedded and linked objects
    objects: Vec<super::object::EmbeddedObject<'a>>,
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
    list_table: super::list::ListTable<'a>,
    saw_list_table: bool,
    /// List override table
    list_override_table: super::list::ListOverrideTable,
    saw_list_override_table: bool,
    legacy_section_numbering: crate::LegacySectionNumbering<'a>,
    paragraph_group_table: Option<crate::ParagraphGroupPropertyTable>,
    /// Sections
    sections: Vec<super::section::Section<'a>>,
    /// Whether section-specific properties are currently active.
    section_properties_active: bool,
    /// Whether header/footer or body content has closed the active section-format prefix.
    section_note_options_closed: bool,
    /// Whether the root body is in an explicit late section-format run.
    root_section_format_run: bool,
    /// Bookmarks
    bookmarks: super::bookmark::BookmarkTable<'a>,
    /// Open bookmark ranges, indexed by name.
    open_bookmarks: HashMap<String, Vec<OpenBookmark>>,
    /// Completed bookmark ranges awaiting content reconstruction.
    bookmark_spans: Vec<BookmarkSpan>,
    /// UTF-8 byte length of body text emitted into style blocks.
    body_text_len: usize,
    /// Stable source order for bookmark ranges.
    next_bookmark_order: usize,
    /// Shapes
    shapes: Vec<super::shape::Shape<'a>>,
    /// Inert legacy drawing text boxes.
    legacy_text_boxes: Vec<crate::LegacyTextBox<'a>>,
    legacy_text_box_text_bytes: usize,
    /// Shape groups
    shape_groups: Vec<super::shape::ShapeGroup<'a>>,
    /// Stylesheet
    stylesheet: super::stylesheet::StyleSheet<'a>,
    /// Whether the unique root stylesheet destination was seen.
    saw_stylesheet: bool,
    /// Document information
    info: super::info::DocumentInfo<'a>,
    /// Annotations
    annotations: Vec<super::annotation::Annotation<'a>>,
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
    notes: Vec<super::section::Note<'a>>,
    note_options: crate::NoteOptions,
    note_options_closed: bool,
    note_separators: crate::NoteSeparatorTable<'a>,
    /// Track changes/revisions
    revisions: Vec<super::annotation::Revision<'a>>,
    /// Authors referenced by tracked-change author indices
    revision_authors: Vec<super::annotation::RevisionAuthor<'a>>,
    /// Whether the unique revision-author table has been parsed.
    saw_revision_table: bool,
    /// Aggregate decoded author-table text.
    revision_author_text_bytes: usize,
    /// Aggregate decoded tracked-change text.
    revision_text_bytes: usize,
    /// Current header/footer being parsed
    #[allow(dead_code)]
    current_header_footer: Option<super::section::HeaderFooter<'a>>,
    /// Current note being parsed (content buffer)
    current_note_buffer: SmallVec<[u8; 256]>,
    /// Current header/footer type being parsed
    current_hf_type: Option<super::section::HeaderFooterType>,
}

#[derive(Default)]
struct FormFieldBuilder {
    field_type: Option<super::form_field::FormFieldType>,
    text_type: Option<super::form_field::FormTextType>,
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

#[derive(Default)]
struct LatentStyleExceptionBuilder {
    locked: Option<bool>,
    semi_hidden: Option<bool>,
    unhide_when_used: Option<bool>,
    quick_format: Option<bool>,
    priority: Option<u8>,
}

impl<'a> Parser<'a> {
    /// Create a new parser.
    pub fn new(tokens: &'a [Token<'a>], arena: &'a Bump) -> Self {
        Self {
            tokens,
            pos: 0,
            states: vec![State::default()],
            font_table: RefCell::new(FontTable::new()),
            saw_font_table: false,
            file_table: None,
            unicode_alternate_depth: 0,
            color_table: RefCell::new(ColorTable::new()),
            blocks: Vec::new(),
            arena,
            tables: Vec::new(),
            current_table: None,
            current_row: None,
            current_cell_text: SmallVec::new(),
            current_cell_nested: Vec::new(),
            nested_table_builders: Vec::new(),
            logical_table_count: 0,
            pictures: Vec::new(),
            fields: Vec::new(),
            form_fields: Vec::new(),
            form_field_text_bytes: 0,
            generator: None,
            revision_save_ids: Vec::new(),
            saw_revision_save_table: false,
            revision_save_root: None,
            saw_revision_save_root: false,
            xml_namespaces: Vec::new(),
            saw_xml_namespace_table: false,
            xml_namespace_text_bytes: 0,
            theme_data: None,
            saw_theme_data: false,
            color_scheme_mapping: None,
            saw_color_scheme_mapping: false,
            latent_styles: None,
            data_store: None,
            saw_data_store: false,
            math_properties: None,
            language_defaults: crate::DocumentLanguageDefaults::default(),
            saw_info_group: false,
            document_direction: None,
            gutter_on_right: false,
            objects: Vec::new(),
            document_variables: Vec::new(),
            document_variable_text_bytes: 0,
            user_properties: Vec::new(),
            user_property_text_bytes: 0,
            navigation_entries: Vec::new(),
            navigation_entry_text_bytes: 0,
            generated_list_markers: Vec::new(),
            generated_list_marker_text_bytes: 0,
            saw_user_properties: false,
            list_table: super::list::ListTable::new(),
            saw_list_table: false,
            list_override_table: super::list::ListOverrideTable::new(),
            saw_list_override_table: false,
            legacy_section_numbering: crate::LegacySectionNumbering::new(),
            paragraph_group_table: None,
            sections: Vec::new(),
            section_properties_active: false,
            section_note_options_closed: false,
            root_section_format_run: false,
            bookmarks: super::bookmark::BookmarkTable::new(),
            open_bookmarks: HashMap::new(),
            bookmark_spans: Vec::new(),
            body_text_len: 0,
            next_bookmark_order: 0,
            shapes: Vec::new(),
            legacy_text_boxes: Vec::new(),
            legacy_text_box_text_bytes: 0,
            shape_groups: Vec::new(),
            stylesheet: super::stylesheet::StyleSheet::new(),
            saw_stylesheet: false,
            info: super::info::DocumentInfo::new(),
            annotations: Vec::new(),
            annotation_ranges: HashMap::new(),
            pending_annotation_author: String::new(),
            pending_annotation_author_seen: false,
            pending_annotation_initials: String::new(),
            pending_annotation_initials_seen: false,
            pending_annotation_mark: false,
            notes: Vec::new(),
            note_options: crate::NoteOptions::default(),
            note_options_closed: false,
            note_separators: crate::NoteSeparatorTable::new(),
            revisions: Vec::new(),
            revision_authors: Vec::new(),
            saw_revision_table: false,
            revision_author_text_bytes: 0,
            revision_text_bytes: 0,
            current_header_footer: None,
            current_note_buffer: SmallVec::new(),
            current_hf_type: None,
        }
    }

    /// Parse the token stream into a document.
    pub fn parse(mut self) -> RtfResult<ParsedDocument<'a>> {
        // Validate document structure
        if self.tokens.is_empty() {
            return Err(RtfError::MalformedDocument(
                "Empty token stream".to_string(),
            ));
        }
        #[derive(Clone, Copy)]
        struct NoteGuardContext {
            body_flow: bool,
            visible_field_result: bool,
            visible_section_format: bool,
            direct_header_footer: bool,
            direct_field_instruction: bool,
            inert_section_format: bool,
        }

        let mut contexts: Vec<NoteGuardContext> = Vec::new();
        for (index, token) in self.tokens.iter().enumerate() {
            match token {
                Token::OpenBrace => {
                    if let Some(parent) = contexts.last_mut() {
                        parent.inert_section_format = false;
                    }
                    let parent = contexts.last().copied().unwrap_or(NoteGuardContext {
                        body_flow: true,
                        visible_field_result: false,
                        visible_section_format: false,
                        direct_header_footer: false,
                        direct_field_instruction: false,
                        inert_section_format: false,
                    });
                    let mut destination_index = index + 1;
                    let starred = matches!(
                        self.tokens.get(destination_index),
                        Some(Token::Control(ControlWord::IgnorableDestination))
                    );
                    if starred {
                        destination_index += 1;
                    }
                    let destination = match self.tokens.get(destination_index) {
                        Some(Token::Control(control)) => Some(control),
                        _ => None,
                    };
                    let context = match destination {
                        Some(ControlWord::FieldResult) if parent.body_flow => NoteGuardContext {
                            body_flow: true,
                            visible_field_result: true,
                            visible_section_format: false,
                            direct_header_footer: false,
                            direct_field_instruction: false,
                            inert_section_format: false,
                        },
                        Some(ControlWord::Field) => NoteGuardContext {
                            body_flow: parent.body_flow,
                            visible_field_result: false,
                            visible_section_format: false,
                            direct_header_footer: false,
                            direct_field_instruction: false,
                            inert_section_format: false,
                        },
                        Some(ControlWord::SectionDefault) if parent.body_flow && !starred => {
                            NoteGuardContext {
                                body_flow: true,
                                visible_field_result: parent.visible_field_result,
                                visible_section_format: true,
                                direct_header_footer: false,
                                direct_field_instruction: false,
                                inert_section_format: false,
                            }
                        },
                        Some(
                            ControlWord::Header
                            | ControlWord::HeaderFirst
                            | ControlWord::HeaderLeft
                            | ControlWord::HeaderRight
                            | ControlWord::Footer
                            | ControlWord::FooterFirst
                            | ControlWord::FooterLeft
                            | ControlWord::FooterRight,
                        ) => NoteGuardContext {
                            body_flow: false,
                            visible_field_result: false,
                            visible_section_format: false,
                            direct_header_footer: true,
                            direct_field_instruction: false,
                            inert_section_format: false,
                        },
                        Some(ControlWord::FieldInstruction) => NoteGuardContext {
                            body_flow: false,
                            visible_field_result: false,
                            visible_section_format: false,
                            direct_header_footer: false,
                            direct_field_instruction: true,
                            inert_section_format: false,
                        },
                        Some(
                            ControlWord::Annotation
                            | ControlWord::Footnote
                            | ControlWord::Endnote
                            | ControlWord::Object
                            | ControlWord::Result
                            | ControlWord::Picture
                            | ControlWord::Shape
                            | ControlWord::ShapeGroup,
                        ) => NoteGuardContext {
                            body_flow: false,
                            visible_field_result: false,
                            visible_section_format: false,
                            direct_header_footer: false,
                            direct_field_instruction: false,
                            inert_section_format: false,
                        },
                        _ if starred => NoteGuardContext {
                            body_flow: false,
                            visible_field_result: false,
                            visible_section_format: false,
                            direct_header_footer: false,
                            direct_field_instruction: false,
                            inert_section_format: false,
                        },
                        _ => NoteGuardContext {
                            body_flow: parent.body_flow,
                            visible_field_result: parent.visible_field_result,
                            visible_section_format: parent.visible_section_format,
                            direct_header_footer: false,
                            direct_field_instruction: false,
                            inert_section_format: false,
                        },
                    };
                    contexts.push(context);
                },
                Token::CloseBrace => {
                    contexts.pop();
                    if contexts.is_empty() {
                        break;
                    }
                },
                Token::Control(ControlWord::SectionDefault) => {
                    if let Some(context) = contexts.last_mut()
                        && (context.direct_header_footer || context.direct_field_instruction)
                    {
                        context.inert_section_format = true;
                    }
                },
                Token::Control(ControlWord::SectionBreak) => {
                    if let Some(context) = contexts.last_mut() {
                        context.inert_section_format = false;
                    }
                },
                Token::Text(text) if !text.is_empty() => {
                    if let Some(context) = contexts.last_mut() {
                        context.inert_section_format = false;
                    }
                },
                Token::Control(control)
                    if matches!(
                        control,
                        ControlWord::Par
                            | ControlWord::Line
                            | ControlWord::Tab
                            | ControlWord::Unicode(_)
                    ) || control_symbol_text(control).is_some() =>
                {
                    if let Some(context) = contexts.last_mut() {
                        context.inert_section_format = false;
                    }
                },
                Token::Control(
                    ControlWord::NoteKinds(_)
                    | ControlWord::FootnotePlacement(_)
                    | ControlWord::EndnotePlacement(_)
                    | ControlWord::FootnoteStart(_)
                    | ControlWord::EndnoteStart(_)
                    | ControlWord::FootnoteRestart(_)
                    | ControlWord::EndnoteRestart(_)
                    | ControlWord::FootnoteNumbering(_)
                    | ControlWord::EndnoteNumbering(_),
                ) if contexts.len() != 1 => {
                    return Err(RtfError::MalformedDocument(
                        "RTF note options must be root document-format controls".to_string(),
                    ));
                },
                Token::Control(
                    ControlWord::SectionFootnotePlacement(_)
                    | ControlWord::SectionFootnoteStart(_)
                    | ControlWord::SectionEndnoteStart(_)
                    | ControlWord::SectionFootnoteRestart(_)
                    | ControlWord::SectionEndnoteRestart(_)
                    | ControlWord::SectionFootnoteNumbering(_)
                    | ControlWord::SectionEndnoteNumbering(_),
                ) if contexts.len() != 1
                    && !contexts
                        .last()
                        .is_some_and(|context| {
                            context.visible_field_result
                                || context.visible_section_format
                                || context.inert_section_format
                        }) =>
                {
                    return Err(RtfError::MalformedDocument(
                        "RTF section note options must be root controls or visible field-result formatting"
                            .to_string(),
                    ));
                },
                _ => {},
            }
        }

        // Expect opening brace
        if !matches!(self.tokens.first(), Some(Token::OpenBrace)) {
            return Err(RtfError::MalformedDocument(
                "Document must start with {".to_string(),
            ));
        }

        // Parse document content
        self.parse_group()?;

        // Finalize any remaining table
        self.finalize_table()?;
        self.finalize_bookmarks()?;
        self.finalize_annotations()?;

        let revision_save = if self.saw_revision_save_table || self.saw_revision_save_root {
            Some(crate::RevisionSaveMetadata::new(
                self.revision_save_ids,
                self.revision_save_root,
            )?)
        } else {
            None
        };
        let theme = match (self.theme_data, self.color_scheme_mapping) {
            (Some(data), mapping) => Some(crate::DocumentTheme::new(
                Cow::Owned(data),
                mapping.map(Cow::Owned),
            )?),
            (None, Some(_)) => {
                return Err(RtfError::MalformedDocument(
                    "RTF color-scheme mapping is orphaned without theme data".to_string(),
                ));
            },
            (None, None) => None,
        };
        let data_store = self
            .data_store
            .map(|data| crate::DocumentDataStore::new(Cow::Owned(data)))
            .transpose()?;

        for section in &self.sections {
            section.properties.columns.validate()?;
        }

        Ok(ParsedDocument {
            font_table: self.font_table.into_inner(),
            file_table: self.file_table,
            color_table: self.color_table.into_inner(),
            blocks: self.blocks,
            tables: self.tables,
            pictures: self.pictures,
            fields: self.fields,
            form_fields: self.form_fields,
            generator: self.generator,
            revision_save,
            xml_namespaces: self.xml_namespaces,
            saw_xml_namespace_table: self.saw_xml_namespace_table,
            theme,
            latent_styles: self.latent_styles,
            data_store,
            math_properties: self.math_properties,
            language_defaults: self.language_defaults,
            document_direction: self.document_direction,
            gutter_on_right: self.gutter_on_right,
            objects: self.objects,
            document_variables: self.document_variables,
            user_properties: self.user_properties,
            navigation_entries: self.navigation_entries,
            generated_list_markers: self.generated_list_markers,
            list_table: self.list_table,
            list_override_table: self.list_override_table,
            legacy_section_numbering: self.legacy_section_numbering,
            paragraph_group_table: self.paragraph_group_table,
            sections: self.sections,
            bookmarks: self.bookmarks,
            shapes: self.shapes,
            legacy_text_boxes: self.legacy_text_boxes,
            shape_groups: self.shape_groups,
            stylesheet: self.stylesheet,
            info: self.info,
            annotations: self.annotations,
            notes: self.notes,
            note_options: self.note_options,
            note_separators: self.note_separators,
            revisions: self.revisions,
            revision_authors: self.revision_authors,
        })
    }

    /// Parse a group (content between braces).
    fn parse_group(&mut self) -> RtfResult<()> {
        self.expect_token(Token::OpenBrace)?;

        if self.states.len() == 2 {
            self.root_section_format_run = false;
        }
        let starts_visible_section_format = self
            .states
            .last()
            .is_some_and(|state| state.destination == Destination::DocumentBody)
            && matches!(
                self.tokens.get(self.pos),
                Some(Token::Control(ControlWord::SectionDefault))
            );

        // Push new state (inherit from parent)
        if let Some(current) = self.states.last() {
            self.states.push(current.clone());
        } else {
            self.states.push(State::default());
        }
        if starts_visible_section_format {
            self.current_state_mut()?.visible_section_format = true;
        }

        let nested_destination=match (self.tokens.get(self.pos),self.tokens.get(self.pos+1)){
            (Some(Token::Control(ControlWord::NestedTableProperties(param))),_)=>Some((true,*param,1)),
            (Some(Token::Control(ControlWord::NoNestedTables(param))),_)=>Some((false,*param,1)),
            (Some(Token::Control(ControlWord::IgnorableDestination)),Some(Token::Control(ControlWord::NestedTableProperties(param))))=>Some((true,*param,2)),
            (Some(Token::Control(ControlWord::IgnorableDestination)),Some(Token::Control(ControlWord::NoNestedTables(param))))=>Some((false,*param,2)),
            _=>None,
        };
        if let Some((properties,param,consumed))=nested_destination{
            require_parameterless(param,if properties{"nesttableprops"}else{"nonesttables"})?;self.pos+=consumed;
            if properties{self.current_state_mut()?.destination=Destination::NestedTableProperties;self.parse_content()?;}else{self.current_state_mut()?.destination=Destination::Other;self.skip_until_close_brace()?;}
            self.states.pop();return Ok(());
        }

        if self.current_state()?.revision_type.is_some()
            && matches!(
                self.tokens.get(self.pos),
                Some(Token::Control(
                    ControlWord::IgnorableDestination
                        | ControlWord::UserProperties
                        | ControlWord::IndexEntry
                        | ControlWord::TableOfContentsEntry
                        | ControlWord::TableOfContentsEntryNoPage
                        | ControlWord::FontTable
                        | ControlWord::ColorTable
                        | ControlWord::StyleSheet
                        | ControlWord::ListTable
                        | ControlWord::ListOverrideTable
                        | ControlWord::RevisionTable
                        | ControlWord::Info
                        | ControlWord::Shape
                        | ControlWord::ShapeGroup
                        | ControlWord::Picture
                        | ControlWord::Object
                        | ControlWord::Result
                        | ControlWord::Field
                        | ControlWord::Header
                        | ControlWord::HeaderFirst
                        | ControlWord::HeaderLeft
                        | ControlWord::HeaderRight
                        | ControlWord::Footer
                        | ControlWord::FooterFirst
                        | ControlWord::FooterLeft
                        | ControlWord::FooterRight
                        | ControlWord::Footnote
                        | ControlWord::Endnote
                ))
            )
        {
            return Err(RtfError::MalformedDocument(
                "RTF revision text cannot contain active or external destinations".to_string(),
            ));
        }

        // Check if this is a special group (header, destination, etc.)
        if self.pos < self.tokens.len() {
            match &self.tokens[self.pos] {
                Token::Control(
                    ControlWord::FileTable | ControlWord::FileEntry | ControlWord::BlipUid,
                ) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF file-table destinations are misplaced or not starred".to_string(),
                    ));
                },
                Token::Control(ControlWord::GeneratedListText) => {
                    self.parse_generated_list_marker(crate::GeneratedListMarkerKind::Modern)?;
                    self.states.pop();
                    return Ok(());
                },
                Token::Control(ControlWord::LegacyGeneratedListText) => {
                    self.parse_generated_list_marker(crate::GeneratedListMarkerKind::Legacy)?;
                    self.states.pop();
                    return Ok(());
                },
                Token::Control(ControlWord::LegacyDrawingObject) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF legacy drawing-object destination must be starred".to_string(),
                    ));
                },
                Token::Control(ControlWord::UnicodeAlternate) => {
                    self.parse_unicode_alternate_group()?;
                    self.states.pop();
                    return Ok(());
                },
                Token::Control(ControlWord::FontTable) => {
                    let valid_scope = self.states.len() == 3
                        || (self.unicode_alternate_depth == 1 && self.states.len() == 4);
                    if self.saw_font_table
                        || !valid_scope
                        || self.blocks.iter().any(|block| !block.text.trim().is_empty())
                    {
                        return Err(RtfError::MalformedDocument(
                            "RTF font table must occur exactly once at document scope before body text".to_string(),
                        ));
                    }
                    self.saw_font_table = true;
                    // Mark this as font table destination
                    if let Some(state) = self.states.last_mut() {
                        state.destination = Destination::FontTable;
                    }
                    self.parse_font_table()?;
                    self.skip_until_close_brace()?;
                    self.states.pop();
                    return Ok(());
                },
                Token::Control(ControlWord::ColorTable) => {
                    // Mark this as color table destination
                    if let Some(state) = self.states.last_mut() {
                        state.destination = Destination::ColorTable;
                    }
                    self.parse_color_table()?;
                    self.skip_until_close_brace()?;
                    self.states.pop();
                    return Ok(());
                },
                Token::Control(ControlWord::UserProperties) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF userprops destination must be starred".to_string(),
                    ));
                },
                Token::Control(ControlWord::DocumentVariable) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF docvar destination must be starred".to_string(),
                    ));
                },
                Token::Control(
                    ControlWord::IndexEntry
                    | ControlWord::TableOfContentsEntry
                    | ControlWord::TableOfContentsEntryNoPage,
                ) => {
                    self.parse_navigation_entry_destination()?;
                    self.states.pop();
                    return Ok(());
                },
                Token::Control(ControlWord::IgnorableDestination) => {
                    if matches!(
                        self.tokens.get(self.pos + 1),
                        Some(Token::Control(
                            ControlWord::BookmarkStart | ControlWord::BookmarkEnd
                        ))
                    ) {
                        self.parse_bookmark_destination()?;
                        self.states.pop();
                        return Ok(());
                    }
                    match self.tokens.get(self.pos + 1) {
                        Some(Token::Control(ControlWord::FileTable)) => {
                            if self.file_table.is_some() {
                                return Err(RtfError::MalformedDocument(
                                    "RTF contains multiple filetbl destinations".to_string(),
                                ));
                            }
                            self.file_table = Some(self.parse_file_table()?);
                            self.states.pop();
                            return Ok(());
                        },
                        Some(Token::Control(ControlWord::BlipUid)) => {
                            return Err(RtfError::MalformedDocument(
                                "RTF blipuid destination may occur only inside pict".to_string(),
                            ));
                        },
                        Some(Token::Control(
                            ControlWord::GeneratedListText
                            | ControlWord::LegacyGeneratedListText,
                        )) => {
                            return Err(RtfError::MalformedDocument(
                                "RTF generated list-marker destinations must not be starred"
                                    .to_string(),
                            ));
                        },
                        Some(Token::Control(
                            ControlWord::LegacyTextBox | ControlWord::LegacyTextBoxText,
                        )) => {
                            return Err(RtfError::MalformedDocument(
                                "RTF legacy text-box controls must occur inside a starred do destination"
                                    .to_string(),
                            ));
                        },
                        Some(Token::Control(
                            ControlWord::NoteKinds(_)
                            | ControlWord::FootnotePlacement(_)
                            | ControlWord::EndnotePlacement(_)
                            | ControlWord::FootnoteStart(_)
                            | ControlWord::EndnoteStart(_)
                            | ControlWord::FootnoteRestart(_)
                            | ControlWord::EndnoteRestart(_)
                            | ControlWord::FootnoteNumbering(_)
                            | ControlWord::EndnoteNumbering(_),
                        )) => {
                            return Err(RtfError::MalformedDocument(
                                "RTF note options must be unstarred root document-format controls"
                                    .to_string(),
                            ));
                        },
                        Some(Token::Control(ControlWord::LegacyDrawingObject)) => {
                            if let Some(text_box) = self.parse_legacy_text_box()? {
                                self.legacy_text_boxes.push(text_box);
                            }
                            self.states.pop();
                            return Ok(());
                        },
                        Some(Token::Control(control @ (
                            ControlWord::FootnoteSeparator
                            | ControlWord::FootnoteContinuationSeparator
                            | ControlWord::FootnoteContinuationNotice
                            | ControlWord::EndnoteSeparator
                            | ControlWord::EndnoteContinuationSeparator
                            | ControlWord::EndnoteContinuationNotice
                        ))) => {
                            let kind = match control {
                                ControlWord::FootnoteSeparator => crate::NoteSeparatorKind::FootnoteSeparator,
                                ControlWord::FootnoteContinuationSeparator => crate::NoteSeparatorKind::FootnoteContinuationSeparator,
                                ControlWord::FootnoteContinuationNotice => crate::NoteSeparatorKind::FootnoteContinuationNotice,
                                ControlWord::EndnoteSeparator => crate::NoteSeparatorKind::EndnoteSeparator,
                                ControlWord::EndnoteContinuationSeparator => crate::NoteSeparatorKind::EndnoteContinuationSeparator,
                                _ => crate::NoteSeparatorKind::EndnoteContinuationNotice,
                            };
                            let separator = self.parse_note_separator_destination(kind)?;
                            self.note_separators.add(separator)?;
                            self.states.pop();
                            return Ok(());
                        },
                        Some(Token::Control(ControlWord::UnicodeAlternateDestination)) => {
                            return Err(RtfError::MalformedDocument(
                                "RTF ud destination must be the Unicode branch of upr".to_string(),
                            ));
                        },
                        Some(Token::Control(ControlWord::ListTable)) => {
                            self.pos += 1;
                            self.parse_list_table()?;
                            self.states.pop();
                            return Ok(());
                        },
                        Some(Token::Control(ControlWord::ListOverrideTable)) => {
                            self.pos += 1;
                            self.parse_list_override_table()?;
                            self.states.pop();
                            return Ok(());
                        },
                        Some(Token::Control(ControlWord::RevisionTable)) => {
                            self.pos += 1;
                            self.parse_revision_table()?;
                            self.states.pop();
                            return Ok(());
                        },
                        Some(Token::Control(ControlWord::FormField | ControlWord::DataField)) => {
                            return Err(RtfError::MalformedDocument(
                                "orphan RTF formfield/datafield destination".to_string(),
                            ));
                        },
                        Some(Token::Control(ControlWord::Generator)) => {
                            if self.generator.is_some() {
                                return Err(RtfError::MalformedDocument(
                                    "RTF contains multiple generator destinations".to_string(),
                                ));
                            }
                            self.generator = Some(self.parse_generator_destination()?);
                            self.states.pop();
                            return Ok(());
                        },
                        Some(Token::Control(ControlWord::RevisionSaveTable)) => {
                            self.pos += 1;
                            self.parse_revision_save_table()?;
                            self.states.pop();
                            return Ok(());
                        },
                        Some(Token::Control(ControlWord::XmlNamespaceTable)) => {
                            self.pos += 1;
                            self.parse_xml_namespace_table()?;
                            self.states.pop();
                            return Ok(());
                        },
                        Some(Token::Control(ControlWord::ThemeData)) => {
                            if self.saw_theme_data {
                                return Err(RtfError::MalformedDocument(
                                    "RTF contains multiple theme-data destinations".to_string(),
                                ));
                            }
                            self.saw_theme_data = true;
                            self.theme_data = Some(self.parse_theme_hex_destination(
                                ControlWord::ThemeData,
                                crate::theme::MAX_THEME_DATA_BYTES,
                            )?);
                            self.states.pop();
                            return Ok(());
                        },
                        Some(Token::Control(ControlWord::ColorSchemeMapping)) => {
                            if self.saw_color_scheme_mapping {
                                return Err(RtfError::MalformedDocument(
                                    "RTF contains multiple color-scheme mappings".to_string(),
                                ));
                            }
                            self.saw_color_scheme_mapping = true;
                            self.color_scheme_mapping = Some(self.parse_theme_hex_destination(
                                ControlWord::ColorSchemeMapping,
                                crate::theme::MAX_COLOR_SCHEME_MAPPING_BYTES,
                            )?);
                            self.states.pop();
                            return Ok(());
                        },
                        Some(Token::Control(ControlWord::LatentStyles)) => {
                            if self.latent_styles.is_some() {
                                return Err(RtfError::MalformedDocument(
                                    "RTF contains multiple latentstyles destinations".to_string(),
                                ));
                            }
                            self.latent_styles = Some(self.parse_latent_styles()?);
                            self.states.pop();
                            return Ok(());
                        },
                        Some(Token::Control(ControlWord::LegacySectionNumberingLevel(_))) => {
                            self.parse_legacy_section_numbering_level()?;
                            self.states.pop();
                            return Ok(());
                        },
                        Some(Token::Control(ControlWord::ParagraphGroupTable)) => {
                            if self.paragraph_group_table.is_some() {
                                return Err(RtfError::MalformedDocument(
                                    "RTF contains multiple pgptbl destinations".to_string(),
                                ));
                            }
                            self.paragraph_group_table = Some(self.parse_paragraph_group_table()?);
                            self.states.pop();
                            return Ok(());
                        },
                        Some(Token::Control(ControlWord::DataStore)) => {
                            if self.saw_data_store {
                                return Err(RtfError::MalformedDocument(
                                    "RTF contains multiple datastore destinations".to_string(),
                                ));
                            }
                            self.saw_data_store = true;
                            self.data_store = Some(self.parse_data_store_destination()?);
                            self.states.pop();
                            return Ok(());
                        },
                        Some(Token::Control(ControlWord::MathProperties)) => {
                            if self.math_properties.is_some() {
                                return Err(RtfError::MalformedDocument(
                                    "RTF contains multiple math-properties destinations"
                                        .to_string(),
                                ));
                            }
                            self.math_properties = Some(self.parse_math_properties_destination()?);
                            self.states.pop();
                            return Ok(());
                        },
                        Some(Token::Control(ControlWord::DocumentVariable)) => {
                            self.parse_document_variable_destination()?;
                            self.states.pop();
                            return Ok(());
                        },
                        Some(Token::Control(ControlWord::UserProperties)) => {
                            self.parse_user_properties_destination()?;
                            self.states.pop();
                            return Ok(());
                        },
                        Some(Token::Control(ControlWord::AnnotationAuthor)) => {
                            if self.pending_annotation_author_seen {
                                return Err(RtfError::MalformedDocument(
                                    "duplicate pending RTF annotation author".to_string(),
                                ));
                            }
                            self.pending_annotation_author =
                                self.parse_ignorable_text_destination()?;
                            self.pending_annotation_author_seen = true;
                            self.states.pop();
                            return Ok(());
                        },
                        Some(Token::Control(ControlWord::AnnotationInitials)) => {
                            if self.pending_annotation_initials_seen {
                                return Err(RtfError::MalformedDocument(
                                    "duplicate pending RTF annotation initials".to_string(),
                                ));
                            }
                            self.pending_annotation_initials =
                                self.parse_ignorable_text_destination()?;
                            self.pending_annotation_initials_seen = true;
                            self.states.pop();
                            return Ok(());
                        },
                        Some(Token::Control(ControlWord::AnnotationRangeStart)) => {
                            self.parse_annotation_range_marker(true)?;
                            self.states.pop();
                            return Ok(());
                        },
                        Some(Token::Control(ControlWord::AnnotationRangeEnd)) => {
                            self.parse_annotation_range_marker(false)?;
                            self.states.pop();
                            return Ok(());
                        },
                        Some(Token::Control(ControlWord::Annotation)) => {
                            self.parse_annotation_destination()?;
                            self.states.pop();
                            return Ok(());
                        },
                        Some(Token::Control(ControlWord::BackgroundDestination)) => {
                            self.pos += 2;
                            if let Some(state) = self.states.last_mut() {
                                state.destination = Destination::Other;
                            }
                            self.parse_content()?;
                            self.states.pop();
                            return Ok(());
                        },
                        _ => {},
                    }
                    // Mark as other destination and skip
                    if let Some(state) = self.states.last_mut() {
                        state.destination = Destination::Other;
                    }
                    self.skip_until_close_brace()?;
                    self.states.pop();
                    return Ok(());
                },
                Token::Control(ControlWord::StyleSheet) => {
                    // Parse style definitions without adding their names to body text.
                    if let Some(state) = self.states.last_mut() {
                        state.destination = Destination::StyleSheet;
                    }
                    self.parse_stylesheet()?;
                    self.states.pop();
                    return Ok(());
                },
                Token::Control(ControlWord::ListTable) => {
                    self.parse_list_table()?;
                    self.states.pop();
                    return Ok(());
                },
                Token::Control(ControlWord::ListOverrideTable) => {
                    self.parse_list_override_table()?;
                    self.states.pop();
                    return Ok(());
                },
                Token::Control(ControlWord::RevisionTable) => {
                    self.parse_revision_table()?;
                    self.states.pop();
                    return Ok(());
                },
                Token::Control(ControlWord::FormField | ControlWord::DataField) => {
                    return Err(RtfError::MalformedDocument(
                        "orphan RTF formfield/datafield destination".to_string(),
                    ));
                },
                Token::Control(ControlWord::Generator) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF generator destination must be starred".to_string(),
                    ));
                },
                Token::Control(ControlWord::RevisionSaveTable) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF revision-save table must be starred".to_string(),
                    ));
                },
                Token::Control(ControlWord::XmlNamespaceTable) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF XML namespace table must be starred".to_string(),
                    ));
                },
                Token::Control(ControlWord::ThemeData | ControlWord::ColorSchemeMapping) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF theme destinations must be starred".to_string(),
                    ));
                },
                Token::Control(ControlWord::LatentStyles | ControlWord::LatentStyleExceptions) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF latent-style destinations are misplaced or not starred".to_string(),
                    ));
                },
                Token::Control(ControlWord::LegacySectionNumberingLevel(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF pnseclvl destination is misplaced or not starred".to_string(),
                    ));
                },
                Token::Control(ControlWord::ParagraphGroupTable | ControlWord::ParagraphGroup) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF paragraph-group destination is misplaced or not starred".to_string(),
                    ));
                },
                Token::Control(ControlWord::DataStore) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF datastore destination must be starred".to_string(),
                    ));
                },
                Token::Control(ControlWord::MathProperties) => {
                    if self.math_properties.is_some() {
                        return Err(RtfError::MalformedDocument(
                            "RTF contains multiple math-properties destinations".to_string(),
                        ));
                    }
                    self.math_properties = Some(self.parse_math_properties_destination()?);
                    self.states.pop();
                    return Ok(());
                },
                Token::Control(ControlWord::Info) => {
                    // Parse document metadata without adding it to body text.
                    if let Some(state) = self.states.last_mut() {
                        state.destination = Destination::Info;
                    }
                    self.parse_info()?;
                    self.states.pop();
                    return Ok(());
                },
                Token::Control(ControlWord::Shape) => {
                    if self.shapes.len() >= MAX_SHAPES {
                        return Err(RtfError::MalformedDocument(
                            "RTF shape count exceeds the safety limit".to_string(),
                        ));
                    }
                    let shape = self.parse_shape_destination()?;
                    self.shapes.push(shape);
                    self.states.pop();
                    return Ok(());
                },
                Token::Control(ControlWord::ShapeGroup) => {
                    if self.shape_groups.len() >= MAX_SHAPE_GROUPS {
                        return Err(RtfError::MalformedDocument(
                            "RTF shape group count exceeds the safety limit".to_string(),
                        ));
                    }
                    let group = self.parse_shape_group_destination()?;
                    self.shape_groups.push(group);
                    self.states.pop();
                    return Ok(());
                },
                Token::Control(ControlWord::Picture) => {
                    // Mark as picture destination and extract
                    if let Some(state) = self.states.last_mut() {
                        state.destination = Destination::Picture;
                    }
                    self.parse_picture()?;
                    self.skip_until_close_brace()?;
                    self.states.pop();
                    return Ok(());
                },
                Token::Control(ControlWord::Object) => {
                    if self.objects.len() >= MAX_OBJECTS {
                        return Err(RtfError::MalformedDocument(
                            "RTF embedded object count exceeds the safety limit".to_string(),
                        ));
                    }
                    let object = self.parse_object_destination()?;
                    self.objects.push(object);
                    self.states.pop();
                    return Ok(());
                },
                Token::Control(ControlWord::Result) => {
                    // Mark as result destination and skip
                    // This contains the rendered result of an embedded object
                    if let Some(state) = self.states.last_mut() {
                        state.destination = Destination::Result;
                    }
                    self.skip_until_close_brace()?;
                    self.states.pop();
                    return Ok(());
                },
                Token::Control(ControlWord::Field) => {
                    // Parse field group
                    self.parse_field()?;
                    self.skip_until_close_brace()?;
                    self.states.pop();
                    return Ok(());
                },
                Token::Control(ControlWord::Header) => {
                    self.pos += 1;
                    if let Some(state) = self.states.last_mut() {
                        state.destination = Destination::Header;
                    }
                    self.current_hf_type = Some(super::section::HeaderFooterType::Header);
                    self.parse_header_footer_content()?;
                    self.states.pop();
                    return Ok(());
                },
                Token::Control(ControlWord::HeaderFirst) => {
                    self.pos += 1;
                    if let Some(state) = self.states.last_mut() {
                        state.destination = Destination::Header;
                    }
                    self.current_hf_type = Some(super::section::HeaderFooterType::HeaderFirst);
                    self.parse_header_footer_content()?;
                    self.states.pop();
                    return Ok(());
                },
                Token::Control(ControlWord::HeaderLeft) => {
                    self.pos += 1;
                    if let Some(state) = self.states.last_mut() {
                        state.destination = Destination::Header;
                    }
                    self.current_hf_type = Some(super::section::HeaderFooterType::HeaderLeft);
                    self.parse_header_footer_content()?;
                    self.states.pop();
                    return Ok(());
                },
                Token::Control(ControlWord::HeaderRight) => {
                    self.pos += 1;
                    if let Some(state) = self.states.last_mut() {
                        state.destination = Destination::Header;
                    }
                    self.current_hf_type = Some(super::section::HeaderFooterType::HeaderRight);
                    self.parse_header_footer_content()?;
                    self.states.pop();
                    return Ok(());
                },
                Token::Control(ControlWord::Footer) => {
                    self.pos += 1;
                    if let Some(state) = self.states.last_mut() {
                        state.destination = Destination::Footer;
                    }
                    self.current_hf_type = Some(super::section::HeaderFooterType::Footer);
                    self.parse_header_footer_content()?;
                    self.states.pop();
                    return Ok(());
                },
                Token::Control(ControlWord::FooterFirst) => {
                    self.pos += 1;
                    if let Some(state) = self.states.last_mut() {
                        state.destination = Destination::Footer;
                    }
                    self.current_hf_type = Some(super::section::HeaderFooterType::FooterFirst);
                    self.parse_header_footer_content()?;
                    self.states.pop();
                    return Ok(());
                },
                Token::Control(ControlWord::FooterLeft) => {
                    self.pos += 1;
                    if let Some(state) = self.states.last_mut() {
                        state.destination = Destination::Footer;
                    }
                    self.current_hf_type = Some(super::section::HeaderFooterType::FooterLeft);
                    self.parse_header_footer_content()?;
                    self.states.pop();
                    return Ok(());
                },
                Token::Control(ControlWord::FooterRight) => {
                    self.pos += 1;
                    if let Some(state) = self.states.last_mut() {
                        state.destination = Destination::Footer;
                    }
                    self.current_hf_type = Some(super::section::HeaderFooterType::FooterRight);
                    self.parse_header_footer_content()?;
                    self.states.pop();
                    return Ok(());
                },
                Token::Control(ControlWord::Footnote) => {
                    self.pos += 1;
                    if let Some(state) = self.states.last_mut() {
                        state.destination = Destination::Footnote;
                    }
                    self.parse_note(true)?;
                    self.states.pop();
                    return Ok(());
                },
                Token::Control(
                    ControlWord::FootnoteSeparator
                    | ControlWord::FootnoteContinuationSeparator
                    | ControlWord::FootnoteContinuationNotice
                    | ControlWord::EndnoteSeparator
                    | ControlWord::EndnoteContinuationSeparator
                    | ControlWord::EndnoteContinuationNotice,
                ) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF note-separator destinations must be starred".to_string(),
                    ));
                },
                Token::Control(ControlWord::Endnote) => {
                    self.pos += 1;
                    if let Some(state) = self.states.last_mut() {
                        state.destination = Destination::Endnote;
                    }
                    self.parse_note(false)?;
                    self.states.pop();
                    return Ok(());
                },
                _ => {},
            }
        }

        // Parse group content. A scoped revision marker without any inert text
        // is an orphan rather than an empty tracked change.
        let revision_text_bytes_before = self.revision_text_bytes;
        self.parse_content()?;
        if self.current_state()?.revision_type.is_some()
            && self.revision_text_bytes == revision_text_bytes_before
        {
            return Err(RtfError::MalformedDocument(
                "RTF revision marker has no tracked text".to_string(),
            ));
        }

        // Pop state
        self.states.pop();

        Ok(())
    }

    fn parse_unicode_alternate_group(&mut self) -> RtfResult<()> {
        self.pos += 1; // upr
        if !matches!(self.tokens.get(self.pos), Some(Token::OpenBrace)) {
            return Err(RtfError::MalformedDocument(
                "RTF upr lacks its ANSI fallback group".to_string(),
            ));
        }
        // A Unicode-aware reader must ignore the first (ANSI) representation.
        self.skip_group()?;
        if !matches!(
            self.tokens.get(self.pos..self.pos + 3),
            Some([
                Token::OpenBrace,
                Token::Control(ControlWord::IgnorableDestination),
                Token::Control(ControlWord::UnicodeAlternateDestination),
            ])
        ) {
            return Err(RtfError::MalformedDocument(
                "RTF upr lacks its starred ud destination".to_string(),
            ));
        }
        self.pos += 3;
        self.unicode_alternate_depth = self
            .unicode_alternate_depth
            .checked_add(1)
            .ok_or_else(|| RtfError::MalformedDocument("RTF upr nesting overflow".to_string()))?;
        if self.unicode_alternate_depth > 8 {
            return Err(RtfError::MalformedDocument(
                "RTF upr nesting exceeds the safety limit".to_string(),
            ));
        }
        let parsed = self.parse_content();
        self.unicode_alternate_depth -= 1;
        parsed?;
        self.expect_token(Token::CloseBrace)?; // outer upr group
        Ok(())
    }

    fn parse_note_separator_destination(
        &mut self,
        kind: crate::NoteSeparatorKind,
    ) -> RtfResult<crate::NoteSeparator<'a>> {
        if self.states.len() != 3
            || self.blocks.iter().any(|block| !block.text.trim().is_empty())
        {
            return Err(RtfError::MalformedDocument(
                "RTF note separators must occur at document scope before body text".to_string(),
            ));
        }
        self.pos += 2; // ignorable marker and destination
        let mut elements = Vec::new();
        let mut unicode_skip = self.current_state()?.unicode_skip.max(0);
        self.parse_note_separator_elements(&mut elements, &mut unicode_skip, 0)?;
        let separator = crate::NoteSeparator { kind, elements };
        separator.validate()?;
        Ok(separator)
    }

    fn parse_note_separator_elements(
        &mut self,
        elements: &mut Vec<crate::NoteSeparatorElement<'a>>,
        unicode_skip: &mut i32,
        depth: usize,
    ) -> RtfResult<()> {
        if depth > 16 {
            return Err(RtfError::MalformedDocument(
                "RTF note-separator nesting exceeds the safety limit".to_string(),
            ));
        }
        while self.pos < self.tokens.len() {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    return Ok(());
                },
                Some(Token::OpenBrace) => {
                    let direct = self.tokens.get(self.pos + 1);
                    let starred = self.tokens.get(self.pos + 2);
                    if matches!(direct, Some(Token::Control(
                        ControlWord::Field
                        | ControlWord::Object
                        | ControlWord::Picture
                        | ControlWord::Shape
                        | ControlWord::ShapeGroup
                        | ControlWord::Footnote
                        | ControlWord::Endnote
                    ))) || (matches!(direct, Some(Token::Control(ControlWord::IgnorableDestination)))
                        && matches!(starred, Some(Token::Control(
                            ControlWord::Field
                            | ControlWord::Object
                            | ControlWord::Picture
                            | ControlWord::Shape
                            | ControlWord::ShapeGroup
                        ))))
                    {
                        return Err(RtfError::MalformedDocument(
                            "RTF note separator cannot contain fields, objects, pictures, or active destinations".to_string(),
                        ));
                    }
                    self.pos += 1;
                    self.parse_note_separator_elements(elements, unicode_skip, depth + 1)?;
                    continue;
                },
                Some(Token::Text(text)) => {
                    let decoded = self.decode_transport_text(text)?;
                    Self::push_note_separator_text(elements, decoded);
                },
                Some(Token::Control(ControlWord::Unicode(first))) => {
                    let decoded = self.parse_style_unicode(*first, *unicode_skip)?;
                    Self::push_note_separator_text(elements, decoded);
                    continue;
                },
                Some(Token::Control(ControlWord::UnicodeSkip(value))) => *unicode_skip = (*value).max(0),
                Some(Token::Control(ControlWord::NoteSeparatorCharacter)) => {
                    elements.push(crate::NoteSeparatorElement::SeparatorMark)
                },
                Some(Token::Control(ControlWord::NoteContinuationSeparatorCharacter)) => {
                    elements.push(crate::NoteSeparatorElement::ContinuationSeparatorMark)
                },
                Some(Token::Control(ControlWord::Par)) => elements.push(crate::NoteSeparatorElement::ParagraphBreak),
                Some(Token::Control(ControlWord::Line)) => elements.push(crate::NoteSeparatorElement::LineBreak),
                Some(Token::Control(ControlWord::Tab)) => Self::push_note_separator_text(elements, "\t".to_string()),
                Some(Token::Control(control)) if control_symbol_text(control).is_some() => {
                    Self::push_note_separator_text(elements, control_symbol_text(control).unwrap_or_default().to_string())
                },
                Some(Token::Binary(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF note separator cannot contain binary data".to_string(),
                    ));
                },
                Some(Token::Control(_)) => {}, // formatting is inert
                None => return Err(RtfError::UnexpectedEof),
            }
            self.pos += 1;
            if elements.len() > crate::note_separator::MAX_NOTE_SEPARATOR_ELEMENTS {
                return Err(RtfError::MalformedDocument(
                    "RTF note separator contains too many elements".to_string(),
                ));
            }
            let text_bytes = elements.iter().map(|element| match element {
                crate::NoteSeparatorElement::Text(text) => text.len(),
                _ => 0,
            }).sum::<usize>();
            if text_bytes > crate::note_separator::MAX_NOTE_SEPARATOR_TEXT_BYTES {
                return Err(RtfError::MalformedDocument(
                    "RTF note-separator text exceeds the safety limit".to_string(),
                ));
            }
        }
        Err(RtfError::UnexpectedEof)
    }

    fn push_note_separator_text(
        elements: &mut Vec<crate::NoteSeparatorElement<'a>>,
        text: String,
    ) {
        if text.is_empty() {
            return;
        }
        if let Some(crate::NoteSeparatorElement::Text(existing)) = elements.last_mut() {
            existing.to_mut().push_str(&text);
        } else {
            elements.push(crate::NoteSeparatorElement::Text(Cow::Owned(text)));
        }
    }

    fn parse_navigation_entry_destination(&mut self) -> RtfResult<()> {
        if self.navigation_entries.len() >= MAX_NAVIGATION_ENTRIES {
            return Err(RtfError::MalformedDocument(
                "RTF navigation-entry count limit exceeded".to_string(),
            ));
        }
        if self.current_state()?.in_table {
            return Err(RtfError::MalformedDocument(
                "RTF positional navigation entries inside tables are unsupported".to_string(),
            ));
        }
        let entry = match self.tokens.get(self.pos) {
            Some(Token::Control(ControlWord::IndexEntry)) => self.parse_index_entry()?,
            Some(Token::Control(ControlWord::TableOfContentsEntry)) => {
                self.parse_table_of_contents_entry(false)?
            },
            Some(Token::Control(ControlWord::TableOfContentsEntryNoPage)) => {
                self.parse_table_of_contents_entry(true)?
            },
            _ => {
                return Err(RtfError::MalformedDocument(
                    "invalid RTF navigation-entry destination".to_string(),
                ));
            },
        };
        entry.validate()?;
        let added = entry.text_bytes().ok_or_else(|| {
            RtfError::MalformedDocument("RTF navigation-entry size overflow".to_string())
        })?;
        self.navigation_entry_text_bytes = self
            .navigation_entry_text_bytes
            .checked_add(added)
            .ok_or_else(|| {
                RtfError::MalformedDocument("RTF navigation-entry size overflow".to_string())
            })?;
        if self.navigation_entry_text_bytes > MAX_NAVIGATION_ENTRY_TEXT_TOTAL_BYTES {
            return Err(RtfError::MalformedDocument(
                "RTF navigation-entry aggregate text limit exceeded".to_string(),
            ));
        }
        self.navigation_entries.push(entry);
        Ok(())
    }

    fn parse_generated_list_marker(
        &mut self,
        kind: crate::GeneratedListMarkerKind,
    ) -> RtfResult<()> {
        if self.current_state()?.destination != Destination::DocumentBody {
            return Err(RtfError::MalformedDocument(
                "RTF generated list marker may occur only in the visible document body"
                    .to_string(),
            ));
        }
        if self.generated_list_markers.len()
            >= crate::generated_list_marker::MAX_GENERATED_LIST_MARKERS
        {
            return Err(RtfError::MalformedDocument(
                "RTF generated list-marker count exceeds the safety limit".to_string(),
            ));
        }

        self.pos += 1;
        let mut depth = 0usize;
        let mut unicode_skip = self.current_state()?.unicode_skip.max(0);
        let mut text = String::new();
        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) if depth == 0 => {
                    self.pos += 1;
                    let marker = crate::GeneratedListMarker {
                        kind,
                        text: Cow::Borrowed(self.arena.alloc_str(&text) as &str),
                        position: self.body_text_len,
                    };
                    marker.validate()?;
                    if self.generated_list_markers.last().is_some_and(|previous| {
                        previous.position == marker.position && previous.kind == marker.kind
                    }) {
                        return Err(RtfError::MalformedDocument(
                            "RTF contains duplicate generated list markers at one body position"
                                .to_string(),
                        ));
                    }
                    self.generated_list_marker_text_bytes = self
                        .generated_list_marker_text_bytes
                        .checked_add(marker.text.len())
                        .ok_or_else(|| {
                            RtfError::MalformedDocument(
                                "RTF generated list-marker text size overflow".to_string(),
                            )
                        })?;
                    if self.generated_list_marker_text_bytes
                        > crate::generated_list_marker::MAX_GENERATED_LIST_MARKER_TOTAL_BYTES
                    {
                        return Err(RtfError::MalformedDocument(
                            "RTF generated list-marker text exceeds the aggregate safety limit"
                                .to_string(),
                        ));
                    }
                    self.generated_list_markers.push(marker);
                    return Ok(());
                },
                Some(Token::CloseBrace) => {
                    depth -= 1;
                    self.pos += 1;
                },
                Some(Token::OpenBrace) => {
                    if matches!(
                        self.tokens.get(self.pos + 1),
                        Some(Token::Control(ControlWord::IgnorableDestination))
                    ) || matches!(
                        self.tokens.get(self.pos + 1),
                        Some(Token::Control(
                            ControlWord::Field
                                | ControlWord::Object
                                | ControlWord::Picture
                                | ControlWord::Shape
                                | ControlWord::ShapeGroup
                                | ControlWord::FormField
                                | ControlWord::DataField
                        ))
                    ) {
                        return Err(RtfError::MalformedDocument(
                            "RTF generated list marker contains an active nested destination"
                                .to_string(),
                        ));
                    }
                    depth = depth.checked_add(1).ok_or_else(|| {
                        RtfError::MalformedDocument(
                            "RTF generated list-marker nesting depth overflow".to_string(),
                        )
                    })?;
                    self.pos += 1;
                },
                Some(Token::Control(ControlWord::Unicode(code))) => {
                    text.push_str(&self.parse_style_unicode(*code, unicode_skip)?);
                },
                Some(Token::Control(ControlWord::UnicodeSkip(value))) => {
                    unicode_skip = (*value).max(0);
                    self.pos += 1;
                },
                Some(Token::Control(ControlWord::Tab)) => {
                    text.push('\t');
                    self.pos += 1;
                },
                Some(Token::Control(control)) if control_symbol_text(control).is_some() => {
                    text.push_str(control_symbol_text(control).unwrap_or_default());
                    self.pos += 1;
                },
                Some(Token::Control(
                    ControlWord::Field
                    | ControlWord::Object
                    | ControlWord::Picture
                    | ControlWord::Shape
                    | ControlWord::ShapeGroup
                    | ControlWord::FormField
                    | ControlWord::DataField
                    | ControlWord::Par
                    | ControlWord::Line,
                )) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF generated list marker contains active or structural content"
                            .to_string(),
                    ));
                },
                Some(Token::Control(_)) => self.pos += 1,
                Some(Token::Text(value)) => {
                    text.push_str(&self.decode_transport_text(value)?);
                    self.pos += 1;
                },
                Some(Token::Binary(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF generated list marker cannot contain binary data".to_string(),
                    ));
                },
                None => return Err(RtfError::UnexpectedEof),
            }
            if text.len() > crate::generated_list_marker::MAX_GENERATED_LIST_MARKER_BYTES {
                return Err(RtfError::MalformedDocument(
                    "RTF generated list marker exceeds the safety limit".to_string(),
                ));
            }
        }
    }

    fn parse_index_entry(&mut self) -> RtfResult<NavigationEntry<'a>> {
        let position = self.body_text_len;
        self.pos += 1; // \xe
        let mut text = String::new();
        let mut index_id = None;
        let mut bold_page_number = false;
        let mut italic_page_number = false;
        let mut page_reference = IndexPageReference::CurrentPage;
        let mut yomi = None;
        let mut saw_yomi = false;
        let mut saw_text = false;
        let mut saw_reference = false;

        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    break;
                },
                Some(Token::OpenBrace) => match (self.tokens.get(self.pos + 1), self.tokens.get(self.pos + 2)) {
                    (Some(Token::Control(ControlWord::IndexReplacementText)), _) => {
                        if !saw_text || saw_reference {
                            return Err(RtfError::MalformedDocument(
                                "RTF index entry has a misplaced or duplicate txe/rxe destination".to_string(),
                            ));
                        }
                        page_reference = IndexPageReference::ReplacementText(Cow::Owned(
                            self.parse_navigation_subdestination(false)?,
                        ));
                        saw_reference = true;
                    },
                    (Some(Token::Control(ControlWord::IndexBookmarkRange)), _) => {
                        if !saw_text || saw_reference {
                            return Err(RtfError::MalformedDocument(
                                "RTF index entry has a misplaced or duplicate txe/rxe destination".to_string(),
                            ));
                        }
                        page_reference = IndexPageReference::BookmarkRange(Cow::Owned(
                            self.parse_navigation_subdestination(false)?,
                        ));
                        saw_reference = true;
                    },
                    (
                        Some(Token::Control(ControlWord::IgnorableDestination)),
                        Some(Token::Control(ControlWord::IndexPronunciation)),
                    ) => {
                        if !saw_yomi || yomi.is_some() {
                            return Err(RtfError::MalformedDocument(
                                "RTF pxe pronunciation requires one preceding yxe".to_string(),
                            ));
                        }
                        yomi = Some(Cow::Owned(self.parse_navigation_subdestination(true)?));
                    },
                    (Some(Token::Control(ControlWord::IndexYomi)), _) => {
                        if !saw_text || saw_yomi || yomi.is_some() {
                            return Err(RtfError::MalformedDocument(
                                "RTF index entry has a misplaced or duplicate yxe group"
                                    .to_string(),
                            ));
                        }
                        yomi = Some(Cow::Owned(self.parse_index_yomi_group()?));
                        saw_yomi = true;
                    },
                    _ => {
                        if saw_reference || yomi.is_some() {
                            return Err(RtfError::MalformedDocument(
                                "RTF index entry text must precede its subdestinations".to_string(),
                            ));
                        }
                        self.parse_navigation_text_group(&mut text, true, 1)?;
                        saw_text = !text.is_empty();
                    },
                },
                Some(Token::Control(ControlWord::IndexIdentifier(value))) => {
                    if saw_text || index_id.is_some() {
                        return Err(RtfError::MalformedDocument(
                            "RTF xef must occur once before index text".to_string(),
                        ));
                    }
                    let value = value.ok_or_else(|| {
                        RtfError::MalformedDocument("RTF xef requires a parameter".to_string())
                    })?;
                    index_id = Some(u8::try_from(value).map_err(|_| {
                        RtfError::MalformedDocument("RTF xef parameter is out of range".to_string())
                    })?);
                    self.pos += 1;
                },
                Some(Token::Control(ControlWord::IndexBold(value))) => {
                    if saw_text || bold_page_number {
                        return Err(RtfError::MalformedDocument(
                            "RTF bxe must occur once before index text".to_string(),
                        ));
                    }
                    bold_page_number = *value;
                    self.pos += 1;
                },
                Some(Token::Control(ControlWord::IndexItalic(value))) => {
                    if saw_text || italic_page_number {
                        return Err(RtfError::MalformedDocument(
                            "RTF ixe must occur once before index text".to_string(),
                        ));
                    }
                    italic_page_number = *value;
                    self.pos += 1;
                },
                Some(Token::Control(ControlWord::IndexYomi)) => {
                    if !saw_text || saw_yomi {
                        return Err(RtfError::MalformedDocument(
                            "RTF yxe must occur once after index text".to_string(),
                        ));
                    }
                    saw_yomi = true;
                    self.pos += 1;
                },
                Some(_) => {
                    if saw_reference || yomi.is_some() {
                        return Err(RtfError::MalformedDocument(
                            "RTF index text cannot follow a subdestination".to_string(),
                        ));
                    }
                    self.parse_navigation_text_token(&mut text, true, 1)?;
                    saw_text = !text.is_empty();
                },
                None => return Err(RtfError::UnexpectedEof),
            }
        }
        if saw_yomi != yomi.is_some() {
            return Err(RtfError::MalformedDocument(
                "RTF yxe and pxe pronunciation controls must occur together".to_string(),
            ));
        }
        let mut entry = IndexEntry::new(position, Cow::Owned(text))?;
        entry.index_id = index_id;
        entry.bold_page_number = bold_page_number;
        entry.italic_page_number = italic_page_number;
        entry.page_reference = page_reference;
        entry.yomi = yomi;
        entry.validate()?;
        Ok(NavigationEntry::Index(entry))
    }

    fn parse_index_yomi_group(&mut self) -> RtfResult<String> {
        self.pos += 2; // group open and \yxe
        let state = self.current_state()?.clone();
        self.states.push(state);
        if !matches!(
            (self.tokens.get(self.pos), self.tokens.get(self.pos + 1)),
            (
                Some(Token::OpenBrace),
                Some(Token::Control(ControlWord::IgnorableDestination))
            )
        ) || !matches!(
            self.tokens.get(self.pos + 2),
            Some(Token::Control(ControlWord::IndexPronunciation))
        ) {
            return Err(RtfError::MalformedDocument(
                "RTF yxe group must contain one immediate starred pxe destination".to_string(),
            ));
        }
        let value = self.parse_navigation_subdestination(true)?;
        if !matches!(self.tokens.get(self.pos), Some(Token::CloseBrace)) {
            return Err(RtfError::MalformedDocument(
                "RTF yxe group must contain only its pxe destination".to_string(),
            ));
        }
        self.pos += 1;
        self.states.pop();
        Ok(value)
    }

    fn parse_table_of_contents_entry(
        &mut self,
        suppress_page_number: bool,
    ) -> RtfResult<NavigationEntry<'a>> {
        let position = self.body_text_len;
        self.pos += 1; // \tc or \tcn
        let mut text = String::new();
        let mut table_id = b'C';
        let mut level = 1u8;
        let mut saw_table = false;
        let mut saw_level = false;
        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    break;
                },
                Some(Token::Control(ControlWord::TableOfContentsTable(value))) => {
                    if !text.is_empty() || saw_table {
                        return Err(RtfError::MalformedDocument(
                            "RTF tcf must occur once before TOC-entry text".to_string(),
                        ));
                    }
                    table_id = u8::try_from(*value).map_err(|_| {
                        RtfError::MalformedDocument("RTF tcf parameter is out of range".to_string())
                    })?;
                    saw_table = true;
                    self.pos += 1;
                },
                Some(Token::Control(ControlWord::TableOfContentsLevel(value))) => {
                    if !text.is_empty() || saw_level {
                        return Err(RtfError::MalformedDocument(
                            "RTF tcl must occur once before TOC-entry text".to_string(),
                        ));
                    }
                    level = u8::try_from(*value).map_err(|_| {
                        RtfError::MalformedDocument("RTF tcl parameter is out of range".to_string())
                    })?;
                    saw_level = true;
                    self.pos += 1;
                },
                Some(Token::OpenBrace) => self.parse_navigation_text_group(&mut text, true, 1)?,
                Some(_) => self.parse_navigation_text_token(&mut text, true, 1)?,
                None => return Err(RtfError::UnexpectedEof),
            }
        }
        let mut entry = TableOfContentsEntry::new(position, Cow::Owned(text))?;
        entry.table_id = table_id;
        entry.level = level;
        entry.suppress_page_number = suppress_page_number;
        entry.validate()?;
        Ok(NavigationEntry::TableOfContents(entry))
    }

    fn parse_navigation_subdestination(&mut self, starred: bool) -> RtfResult<String> {
        self.pos += if starred { 3 } else { 2 }; // group, optional star, destination
        let state = self.current_state()?.clone();
        self.states.push(state);
        let mut value = String::new();
        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    self.states.pop();
                    break;
                },
                Some(Token::OpenBrace) => {
                    self.parse_navigation_text_group(&mut value, false, 1)?;
                },
                Some(_) => self.parse_navigation_text_token(&mut value, false, 1)?,
                None => return Err(RtfError::UnexpectedEof),
            }
        }
        if value.is_empty() {
            return Err(RtfError::MalformedDocument(
                "RTF navigation subdestination text cannot be empty".to_string(),
            ));
        }
        Ok(value)
    }

    fn parse_navigation_text_group(
        &mut self,
        output: &mut String,
        visible: bool,
        depth: usize,
    ) -> RtfResult<()> {
        if depth > MAX_NAVIGATION_ENTRY_DEPTH {
            return Err(RtfError::MalformedDocument(
                "RTF navigation-entry nesting limit exceeded".to_string(),
            ));
        }
        self.pos += 1; // group open
        let state = self.current_state()?.clone();
        self.states.push(state);
        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    self.states.pop();
                    return Ok(());
                },
                Some(Token::OpenBrace) => {
                    self.parse_navigation_text_group(output, visible, depth + 1)?;
                },
                Some(_) => self.parse_navigation_text_token(output, visible, depth)?,
                None => return Err(RtfError::UnexpectedEof),
            }
        }
    }

    fn parse_navigation_text_token(
        &mut self,
        output: &mut String,
        visible: bool,
        _depth: usize,
    ) -> RtfResult<()> {
        let decoded = match self.tokens.get(self.pos) {
            Some(Token::Text(text)) => {
                let decoded = self.decode_transport_text(text)?;
                self.pos += 1;
                Some(decoded)
            },
            Some(Token::Control(ControlWord::Unicode(code))) => {
                Some(self.parse_navigation_unicode_sequence(*code)?)
            },
            Some(Token::Control(ControlWord::Par | ControlWord::Line)) => {
                self.pos += 1;
                Some("\n".to_string())
            },
            Some(Token::Control(ControlWord::Tab)) => {
                self.pos += 1;
                Some("\t".to_string())
            },
            Some(Token::Control(control)) if control_symbol_text(control).is_some() => {
                let decoded = control_symbol_text(control).unwrap_or_default().to_string();
                self.pos += 1;
                Some(decoded)
            },
            Some(Token::Control(control)) => {
                if Self::forbidden_navigation_control(control) {
                    return Err(RtfError::MalformedDocument(
                        "RTF navigation entries cannot contain active or nested destinations"
                            .to_string(),
                    ));
                }
                let control = *control;
                self.pos += 1;
                self.apply_control_word(&control)?;
                None
            },
            Some(Token::Binary(_)) => {
                return Err(RtfError::MalformedDocument(
                    "RTF navigation entries cannot contain binary data".to_string(),
                ));
            },
            Some(Token::OpenBrace | Token::CloseBrace) => {
                return Err(RtfError::MalformedDocument(
                    "invalid RTF navigation-entry group structure".to_string(),
                ));
            },
            None => return Err(RtfError::UnexpectedEof),
        };
        if let Some(decoded) = decoded {
            let new_len = output.len().checked_add(decoded.len()).ok_or_else(|| {
                RtfError::MalformedDocument("RTF navigation-entry size overflow".to_string())
            })?;
            if new_len > MAX_NAVIGATION_ENTRY_TEXT_BYTES {
                return Err(RtfError::MalformedDocument(
                    "RTF navigation-entry text limit exceeded".to_string(),
                ));
            }
            output.push_str(&decoded);
            if visible && !self.current_state()?.formatting.hidden {
                self.append_semantic_text(&decoded)?;
            }
        }
        Ok(())
    }

    fn parse_navigation_unicode_sequence(&mut self, first_code: i32) -> RtfResult<String> {
        let skip_count = self.current_state()?.unicode_skip.max(0) as usize;
        let mut utf16 = SmallVec::<[u16; 4]>::new();
        let mut code = first_code;
        let mut remainder = String::new();
        loop {
            utf16.push(code as u16);
            self.pos += 1;
            let mut fallback_skip = skip_count;
            while fallback_skip > 0 {
                match self.tokens.get(self.pos) {
                    Some(Token::Text(text)) => {
                        let count = text.chars().count();
                        if count <= fallback_skip {
                            fallback_skip -= count;
                        } else {
                            remainder.extend(text.chars().skip(fallback_skip));
                            fallback_skip = 0;
                        }
                        self.pos += 1;
                    },
                    Some(Token::Control(control)) if control_symbol_text(control).is_some() => {
                        fallback_skip -= 1;
                        self.pos += 1;
                    },
                    _ => break,
                }
            }
            if !remainder.is_empty() {
                break;
            }
            match self.tokens.get(self.pos) {
                Some(Token::Control(ControlWord::Unicode(next))) => code = *next,
                _ => break,
            }
        }
        let mut decoded = String::from_utf16(&utf16).map_err(|error| {
            RtfError::InvalidUnicode(format!("invalid navigation-entry Unicode: {error}"))
        })?;
        decoded.push_str(&self.decode_transport_text(&remainder)?);
        Ok(decoded)
    }

    fn forbidden_navigation_control(control: &ControlWord<'_>) -> bool {
        matches!(
            control,
            ControlWord::IgnorableDestination
                | ControlWord::Field
                | ControlWord::FieldInstruction
                | ControlWord::FieldResult
                | ControlWord::Object
                | ControlWord::Result
                | ControlWord::Picture
                | ControlWord::Shape
                | ControlWord::ShapeGroup
                | ControlWord::DocumentVariable
                | ControlWord::UserProperties
                | ControlWord::Annotation
                | ControlWord::Footnote
                | ControlWord::Endnote
                | ControlWord::Header
                | ControlWord::HeaderFirst
                | ControlWord::HeaderLeft
                | ControlWord::HeaderRight
                | ControlWord::Footer
                | ControlWord::FooterFirst
                | ControlWord::FooterLeft
                | ControlWord::FooterRight
                | ControlWord::FontTable
                | ControlWord::ColorTable
                | ControlWord::StyleSheet
                | ControlWord::ListTable
                | ControlWord::ListOverrideTable
                | ControlWord::RevisionTable
                | ControlWord::IndexEntry
                | ControlWord::IndexIdentifier(_)
                | ControlWord::IndexBold(_)
                | ControlWord::IndexItalic(_)
                | ControlWord::IndexReplacementText
                | ControlWord::IndexBookmarkRange
                | ControlWord::IndexYomi
                | ControlWord::IndexPronunciation
                | ControlWord::TableOfContentsEntry
                | ControlWord::TableOfContentsEntryNoPage
                | ControlWord::TableOfContentsTable(_)
                | ControlWord::TableOfContentsLevel(_)
        )
    }

    fn parse_document_variable_destination(&mut self) -> RtfResult<()> {
        // The destination group is one level below the RTF root. Body text is
        // flushed before nested groups, so a nonzero body length also rejects
        // document variables that appear after body content has begun.
        if self.states.len() != 3
            || self
                .blocks
                .iter()
                .any(|block| !block.text.trim().is_empty())
        {
            return Err(RtfError::MalformedDocument(
                "RTF docvar destination must occur in the root document header".to_string(),
            ));
        }
        if self.document_variables.len() >= MAX_DOCUMENT_VARIABLES {
            return Err(RtfError::MalformedDocument(format!(
                "RTF document exceeds {MAX_DOCUMENT_VARIABLES} document variables"
            )));
        }
        self.pos += 2; // \* and \docvar
        let name = self.parse_document_variable_text_group(MAX_DOCUMENT_VARIABLE_NAME_BYTES)?;
        let value = self.parse_document_variable_text_group(MAX_DOCUMENT_VARIABLE_VALUE_BYTES)?;
        if !matches!(self.tokens.get(self.pos), Some(Token::CloseBrace)) {
            return Err(RtfError::MalformedDocument(
                "RTF document variable must contain exactly two immediate text groups".to_string(),
            ));
        }
        self.pos += 1;
        let variable = DocumentVariable::new(Cow::Owned(name), Cow::Owned(value))?;
        let added = variable
            .name
            .len()
            .checked_add(variable.value.len())
            .ok_or_else(|| RtfError::MalformedDocument("document-variable size overflow".to_string()))?;
        self.document_variable_text_bytes = self
            .document_variable_text_bytes
            .checked_add(added)
            .ok_or_else(|| RtfError::MalformedDocument("document-variable size overflow".to_string()))?;
        if self.document_variable_text_bytes > MAX_DOCUMENT_VARIABLE_TEXT_BYTES {
            return Err(RtfError::MalformedDocument(format!(
                "RTF document-variable text exceeds {MAX_DOCUMENT_VARIABLE_TEXT_BYTES} bytes"
            )));
        }
        self.document_variables.push(variable);
        Ok(())
    }

    fn parse_user_properties_destination(&mut self) -> RtfResult<()> {
        if self.saw_user_properties {
            return Err(RtfError::MalformedDocument(
                "RTF document contains multiple userprops destinations".to_string(),
            ));
        }
        self.saw_user_properties = true;
        self.pos += 2; // \* and \userprops
        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    return Ok(());
                },
                Some(Token::OpenBrace) => self.parse_user_property()?,
                Some(Token::Text(text)) if text.as_bytes().iter().all(u8::is_ascii_whitespace) => {
                    self.pos += 1;
                },
                Some(_) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF userprops may contain only immediate propinfo groups".to_string(),
                    ));
                },
                None => {
                    return Err(RtfError::MalformedDocument(
                        "unterminated RTF userprops destination".to_string(),
                    ));
                },
            }
        }
    }

    fn parse_user_property(&mut self) -> RtfResult<()> {
        if self.user_properties.len() >= MAX_USER_PROPERTIES {
            return Err(RtfError::MalformedDocument(format!(
                "RTF document exceeds {MAX_USER_PROPERTIES} user properties"
            )));
        }
        let name = self.parse_user_property_text_group(
            ControlWord::PropertyName,
            MAX_USER_PROPERTY_NAME_BYTES,
        )?;
        self.skip_user_property_whitespace();
        let type_code = match self.tokens.get(self.pos) {
            Some(Token::Control(ControlWord::PropertyType(Some(type_code)))) => *type_code,
            _ => {
                return Err(RtfError::MalformedDocument(
                    "RTF propinfo requires an immediate proptype parameter".to_string(),
                ));
            },
        };
        self.pos += 1;
        self.skip_user_property_whitespace();
        let lexical = self.parse_user_property_text_group(
            ControlWord::StaticValue,
            MAX_USER_PROPERTY_VALUE_BYTES,
        )?;
        self.skip_user_property_whitespace();
        let link_value = if matches!(
            self.tokens.get(self.pos..self.pos + 2),
            Some([Token::OpenBrace, Token::Control(ControlWord::LinkValue)])
        ) {
            Some(self.parse_user_property_text_group(
                ControlWord::LinkValue,
                MAX_USER_PROPERTY_VALUE_BYTES,
            )?)
        } else {
            None
        };
        if self.user_properties.iter().any(|property| property.name == name) {
            return Err(RtfError::MalformedDocument(format!(
                "duplicate RTF user-property name: {name}"
            )));
        }
        let property = UserProperty::new(
            Cow::Owned(name),
            UserPropertyValue::from_lexical(type_code, Cow::Owned(lexical))?,
            link_value.map(Cow::Owned),
        )?;
        self.user_property_text_bytes = self
            .user_property_text_bytes
            .checked_add(property.text_bytes().ok_or_else(|| {
                RtfError::MalformedDocument("user-property size overflow".to_string())
            })?)
            .ok_or_else(|| RtfError::MalformedDocument("user-property size overflow".to_string()))?;
        if self.user_property_text_bytes > MAX_USER_PROPERTY_TEXT_BYTES {
            return Err(RtfError::MalformedDocument(format!(
                "RTF user-property text exceeds {MAX_USER_PROPERTY_TEXT_BYTES} bytes"
            )));
        }
        self.user_properties.push(property);
        Ok(())
    }

    fn skip_user_property_whitespace(&mut self) {
        while matches!(
            self.tokens.get(self.pos),
            Some(Token::Text(text)) if text.as_bytes().iter().all(u8::is_ascii_whitespace)
        ) {
            self.pos += 1;
        }
    }

    fn parse_user_property_text_group(
        &mut self,
        destination: ControlWord<'a>,
        limit: usize,
    ) -> RtfResult<String> {
        self.expect_token(Token::OpenBrace)?;
        self.expect_token(Token::Control(destination))?;
        self.parse_inert_text_group_contents(limit, "user-property")
    }

    fn parse_document_variable_text_group(&mut self, limit: usize) -> RtfResult<String> {
        self.expect_token(Token::OpenBrace)?;
        self.parse_inert_text_group_contents(limit, "document-variable")
    }

    fn parse_inert_text_group_contents(
        &mut self,
        limit: usize,
        kind: &str,
    ) -> RtfResult<String> {
        let mut bytes = SmallVec::<[u8; 128]>::new();
        let mut output = String::new();
        let mut unicode_skip = self.states.last().map_or(1, |state| state.unicode_skip);
        let mut skip_fallback = 0i32;
        let mut pending_high_surrogate = None;
        loop {
            let token = self.tokens.get(self.pos).ok_or_else(|| {
                RtfError::MalformedDocument(format!("unterminated RTF {kind} text group"))
            })?;
            match token {
                Token::CloseBrace => {
                    self.pos += 1;
                    break;
                },
                Token::OpenBrace => {
                    return Err(RtfError::MalformedDocument(
                        format!("nested groups are not allowed in RTF {kind} text"),
                    ));
                },
                Token::Binary(_) => {
                    return Err(RtfError::MalformedDocument(
                        format!("binary data is not allowed in RTF {kind} text"),
                    ));
                },
                Token::Text(text) => {
                    let mut transport = SmallVec::<[u8; 128]>::new();
                    append_transport_bytes(&mut transport, text)?;
                    let skip = usize::try_from(skip_fallback.max(0)).unwrap_or(usize::MAX);
                    let skipped = skip.min(transport.len());
                    skip_fallback -= i32::try_from(skipped).unwrap_or(i32::MAX);
                    bytes.extend_from_slice(&transport[skipped..]);
                    self.pos += 1;
                },
                Token::Control(ControlWord::UnicodeSkip(value)) => {
                    unicode_skip = *value;
                    self.pos += 1;
                },
                Token::Control(ControlWord::Unicode(value)) => {
                    if !bytes.is_empty() {
                        let decoded = self
                            .states
                            .last()
                            .map_or(RtfEncoding::Standard(encoding_rs::WINDOWS_1252), |state| state.encoding)
                            .decode(&bytes);
                        output.push_str(&decoded);
                        bytes.clear();
                    }
                    let unit = *value as i16 as u16;
                    if (0xD800..=0xDBFF).contains(&unit) {
                        if pending_high_surrogate.replace(unit).is_some() {
                            output.push('\u{FFFD}');
                        }
                    } else if let Some(high) = pending_high_surrogate.take() {
                        output.push(
                            char::decode_utf16([high, unit])
                                .next()
                                .expect("two UTF-16 units")
                                .unwrap_or('\u{FFFD}'),
                        );
                    } else {
                        output.push(
                            char::decode_utf16([unit])
                                .next()
                                .expect("one UTF-16 unit")
                                .unwrap_or('\u{FFFD}'),
                        );
                    }
                    skip_fallback = unicode_skip.max(0);
                    self.pos += 1;
                },
                Token::Control(control) => {
                    if let Some(text) = control_symbol_text(control) {
                        if !bytes.is_empty() {
                            let decoded = self
                                .states
                                .last()
                                .map_or(RtfEncoding::Standard(encoding_rs::WINDOWS_1252), |state| state.encoding)
                                .decode(&bytes);
                            output.push_str(&decoded);
                            bytes.clear();
                        }
                        output.push_str(text);
                        self.pos += 1;
                    } else {
                        return Err(RtfError::MalformedDocument(
                            format!("active controls are not allowed in RTF {kind} text"),
                        ));
                    }
                },
            }
            if output.len().saturating_add(bytes.len()) > limit {
                return Err(RtfError::MalformedDocument(format!(
                    "RTF {kind} text group exceeds {limit} bytes"
                )));
            }
        }
        if !bytes.is_empty() {
            let decoded = self
                .states
                .last()
                .map_or(RtfEncoding::Standard(encoding_rs::WINDOWS_1252), |state| state.encoding)
                .decode(&bytes);
            output.push_str(&decoded);
        }
        if pending_high_surrogate.is_some() {
            output.push('\u{FFFD}');
        }
        if output.len() > limit {
            return Err(RtfError::MalformedDocument(format!(
                "RTF {kind} text group exceeds {limit} bytes"
            )));
        }
        Ok(output)
    }

    /// Parse group content (text and control words).
    fn parse_content(&mut self) -> RtfResult<()> {
        let mut text_buffer = SmallVec::<[u8; 256]>::new();

        while self.pos < self.tokens.len() {
            match &self.tokens[self.pos] {
                Token::CloseBrace => {
                    // Flush any buffered text
                    if !text_buffer.is_empty() {
                        self.flush_text_buffer(&mut text_buffer)?;
                    }
                    self.pos += 1;
                    return Ok(());
                },
                Token::OpenBrace => {
                    // Flush text before entering nested group
                    if !text_buffer.is_empty() {
                        self.flush_text_buffer(&mut text_buffer)?;
                    }
                    self.current_state_mut()?.character_border_active = false;
                    self.parse_group()?;
                },
                Token::Control(control) => {
                    match control {
                        ControlWord::Par | ControlWord::Line => {
                            let structural_table_boundary=self.finalize_table_before_non_table_body_content(true)?;
                            self.pos += 1;
                            // Paragraph break - flush current text
                            if !text_buffer.is_empty() {
                                self.flush_text_buffer(&mut text_buffer)?;
                            }
                            if !structural_table_boundary{text_buffer.push(b'\n');}
                        },
                        ControlWord::Tab => {
                            self.finalize_table_before_non_table_body_content(true)?;
                            self.pos += 1;
                            text_buffer.push(b'\t');
                        },
                        ControlWord::Unicode(code) => {
                            self.finalize_table_before_non_table_body_content(true)?;
                            // Handle Unicode character with potential fallback
                            if self.states.last().is_some_and(|state| {
                                state.destination == Destination::DocumentBody
                            }) {
                                self.section_note_options_closed = true;
                                self.root_section_format_run = false;
                            }
                            if !text_buffer.is_empty() {
                                self.flush_text_buffer(&mut text_buffer)?;
                            }
                            self.parse_unicode_sequence(*code)?;
                        },
                        ControlWord::Ansi
                        | ControlWord::AnsiCodePage(_)
                        | ControlWord::Mac
                        | ControlWord::Pc
                        | ControlWord::Pca => {
                            if !text_buffer.is_empty() {
                                self.flush_text_buffer(&mut text_buffer)?;
                            }
                            self.pos += 1;
                            self.apply_control_word(control)?;
                        },
                        ControlWord::NonBreakingSpace
                        | ControlWord::OptionalHyphen
                        | ControlWord::NonBreakingHyphen => {
                            if !text_buffer.is_empty() {
                                self.flush_text_buffer(&mut text_buffer)?;
                            }
                            let text = control_symbol_text(control).ok_or_else(|| {
                                RtfError::MalformedDocument(
                                    "missing RTF control-symbol text".to_string(),
                                )
                            })?;
                            self.pos += 1;
                            self.append_semantic_text(text)?;
                        },
                        ControlWord::AnnotationMark => {
                            if !text_buffer.is_empty() {
                                self.flush_text_buffer(&mut text_buffer)?;
                            }
                            if self.pending_annotation_mark {
                                return Err(RtfError::MalformedDocument(
                                    "duplicate pending RTF annotation marker".to_string(),
                                ));
                            }
                            self.pending_annotation_mark = true;
                            self.pos += 1;
                        },
                        ControlWord::Revised(_)
                        | ControlWord::Deleted(_)
                        | ControlWord::RevisionAuthor(_)
                        | ControlWord::DeletedRevisionAuthor(_)
                        | ControlWord::RevisionDate(_)
                        | ControlWord::DeletedRevisionDate(_) => {
                            if !text_buffer.is_empty() {
                                self.flush_text_buffer(&mut text_buffer)?;
                            }
                            self.pos += 1;
                            self.apply_control_word(control)?;
                        },
                        ControlWord::FormProtection(_)
                        | ControlWord::AnnotationProtection(_)
                        | ControlWord::RevisionProtection(_)
                        | ControlWord::ReadOnlyProtection(_)
                        | ControlWord::AllProtection(_)
                        | ControlWord::EnforceProtection(_)
                        | ControlWord::ProtectionLevel(_) => {
                            if !text_buffer.is_empty() {
                                self.flush_text_buffer(&mut text_buffer)?;
                            }
                            self.pos += 1;
                            self.apply_control_word(control)?;
                        },
                        _ => {
                            self.pos += 1;
                            // Apply formatting changes
                            self.apply_control_word(control)?;
                        },
                    }
                },
                Token::Text(text) => {
                    self.pos += 1;
                    self.current_state_mut()?.character_border_active = false;
                    // Skip empty text tokens
                    if text.is_empty() {
                        continue;
                    }
                    self.finalize_table_before_non_table_body_content(!text.trim().is_empty())?;
                    if self.states.last().is_some_and(|state| {
                        state.destination == Destination::DocumentBody
                    }) && !text.trim().is_empty()
                    {
                        self.note_options_closed = true;
                        self.section_note_options_closed = true;
                        self.root_section_format_run = false;
                    }
                    // Check if we're in a table
                    if self.current_state().is_ok_and(|s|s.destination==Destination::DocumentBody&&(s.in_table||s.table_nesting_level>=2)) {
                        let state=self.current_state()?.clone();let encoding = state.encoding;
                        let mut bytes = SmallVec::<[u8; 64]>::new();
                        append_transport_bytes(&mut bytes, text)?;
                        self.append_table_text(encoding.decode(&bytes).as_bytes(),state.table_nesting_level)?;
                    } else if self.current_state().is_ok_and(|s|s.destination==Destination::DocumentBody) {
                        append_transport_bytes(&mut text_buffer, text)?;
                    }
                },
                Token::Binary(_) => {
                    if self.current_state()?.revision_type.is_some() {
                        return Err(RtfError::MalformedDocument(
                            "RTF revision text cannot contain binary data".to_string(),
                        ));
                    }
                    self.pos += 1;
                },
            }
        }

        Err(RtfError::UnexpectedEof)
    }

    /// Flush text buffer to a style block.
    fn flush_text_buffer(&mut self, buffer: &mut SmallVec<[u8; 256]>) -> RtfResult<()> {
        if buffer.is_empty() {
            return Ok(());
        }

        let state = self.current_state()?.clone();

        // Only create blocks for text in the document body
        // Skip text from font tables, color tables, stylesheets, etc.
        if state.destination == Destination::DocumentBody {
            let decoded_str = state.encoding.decode(buffer);

            // Allocate in arena and create block
            let text = self.arena.alloc_str(&decoded_str);
            let start = self.body_text_len;
            if state.revision_type == Some(super::annotation::RevisionType::Deletion) {
                self.append_revision_text(&state, text, start, start)?;
                buffer.clear();
                return Ok(());
            }
            let block = StyleBlock::new(Cow::Borrowed(text), state.formatting, state.paragraph);
            self.body_text_len = self.body_text_len.checked_add(text.len()).ok_or_else(|| {
                RtfError::MalformedDocument("RTF body text length overflow".to_string())
            })?;
            self.blocks.push(block);
            self.append_revision_text(&state, text, start, self.body_text_len)?;
        }

        buffer.clear();
        Ok(())
    }

    fn decode_transport_text(&self, text: &str) -> RtfResult<String> {
        let mut bytes = SmallVec::<[u8; 64]>::new();
        append_transport_bytes(&mut bytes, text)?;
        Ok(self.current_state()?.encoding.decode(&bytes).into_owned())
    }

    fn append_semantic_text(&mut self, text: &str) -> RtfResult<()> {
        self.finalize_table_before_non_table_body_content(!text.is_empty())?;
        let state = self.current_state()?.clone();
        if state.destination != Destination::DocumentBody { return Ok(()); }
        if state.in_table||state.table_nesting_level>=2 {
            self.append_table_text(text.as_bytes(),state.table_nesting_level)?;
            return Ok(());
        }
        let text = self.arena.alloc_str(text);
        let start = self.body_text_len;
        if state.revision_type == Some(super::annotation::RevisionType::Deletion) {
            return self.append_revision_text(&state, text, start, start);
        }
        self.body_text_len = self.body_text_len.checked_add(text.len()).ok_or_else(|| {
            RtfError::MalformedDocument("RTF body text length overflow".to_string())
        })?;
        self.blocks.push(StyleBlock::new(
            Cow::Borrowed(text),
            state.formatting,
            state.paragraph,
        ));
        self.append_revision_text(&state, text, start, self.body_text_len)
    }

    fn append_revision_text(
        &mut self,
        state: &State,
        text: &str,
        start: usize,
        end: usize,
    ) -> RtfResult<()> {
        let Some(revision_type) = state.revision_type else {
            return Ok(());
        };
        let id = state.revision_author_id.ok_or_else(|| {
            RtfError::MalformedDocument("RTF revision text is missing an author index".to_string())
        })?;
        let index = usize::try_from(id).map_err(|_| {
            RtfError::MalformedDocument("RTF revision author index cannot be negative".to_string())
        })?;
        let author = self.revision_authors.get(index).ok_or_else(|| {
            RtfError::MalformedDocument("RTF revision author index is outside revtbl".to_string())
        })?;
        let author = author.name.clone();
        let date = state
            .revision_date
            .map(|value| Cow::Owned(value.to_string()));

        self.revision_text_bytes = self
            .revision_text_bytes
            .checked_add(text.len())
            .ok_or_else(|| {
                RtfError::MalformedDocument(
                    "RTF aggregate revision text size overflow".to_string(),
                )
            })?;
        if self.revision_text_bytes > super::annotation::MAX_REVISION_TEXT_TOTAL_BYTES {
            return Err(RtfError::MalformedDocument(
                "RTF aggregate revision text exceeds the safety limit".to_string(),
            ));
        }

        if let Some(previous) = self.revisions.last_mut()
            && previous.revision_type == revision_type
            && previous.id == id
            && previous.author == author
            && previous.date == date
            && previous.range_end == start
            && (revision_type != super::annotation::RevisionType::Deletion
                || previous.position == start)
        {
            if previous.content.len().saturating_add(text.len())
                > super::annotation::MAX_REVISION_TEXT_BYTES
            {
                return Err(RtfError::MalformedDocument(
                    "RTF revision text exceeds the safety limit".to_string(),
                ));
            }
            previous.content.to_mut().push_str(text);
            previous.range_end = end;
            return Ok(());
        }
        if self.revisions.len() >= MAX_REVISIONS {
            return Err(RtfError::MalformedDocument(
                "RTF revision count exceeds the safety limit".to_string(),
            ));
        }
        let revision = super::annotation::Revision {
            revision_type,
            author,
            date,
            id,
            content: Cow::Owned(text.to_string()),
            position: start,
            range_end: end,
        };
        revision.validate()?;
        self.revisions.push(revision);
        Ok(())
    }

    /// Apply a control word to the current state.
    fn apply_control_word(&mut self, control: &ControlWord) -> RtfResult<()> {
        if let ControlWord::TableNestingLevel(parameter)=control{let value=parameter.ok_or_else(||RtfError::MalformedDocument("RTF itap requires a numeric parameter".to_string()))?;let level=u8::try_from(value).map_err(|_|RtfError::MalformedDocument("RTF itap is outside 0..=32".to_string()))?;if usize::from(level)>crate::MAX_TABLE_NESTING_DEPTH{return Err(RtfError::MalformedDocument("RTF itap is outside 0..=32".to_string()))}let previous=self.current_state()?.table_nesting_level;self.current_state_mut()?.table_nesting_level=level;let previous=if previous>=2{previous}else{1};let effective=if level>=2{level}else{1};if effective<previous{self.drain_nested_to(effective)?;}return Ok(())}
        match control {
            ControlWord::RevisionSaveRoot(value) => {
                if self.states.len() != 2 || self.saw_revision_save_root {
                    return Err(RtfError::MalformedDocument(
                        "RTF rsidroot must occur exactly once at document scope".to_string(),
                    ));
                }
                let value = u32::try_from(*value).map_err(|_| {
                    RtfError::MalformedDocument(
                        "RTF revision root must be a positive signed integer".to_string(),
                    )
                })?;
                if value == 0 {
                    return Err(RtfError::MalformedDocument(
                        "RTF revision root must be a positive signed integer".to_string(),
                    ));
                }
                self.saw_revision_save_root = true;
                self.revision_save_root = Some(value);
                return Ok(());
            },
            ControlWord::RevisionSaveId(_) => {
                return Err(RtfError::MalformedDocument(
                    "orphan RTF rsid control outside rsidtbl".to_string(),
                ));
            },
            ControlWord::BlipTag(_) | ControlWord::BlipUnitsPerInch(_) => {
                return Err(RtfError::MalformedDocument(
                    "orphan RTF picture identity control outside pict".to_string(),
                ));
            },
            control @ (ControlWord::NoteKinds(_)
            | ControlWord::FootnotePlacement(_)
            | ControlWord::EndnotePlacement(_)
            | ControlWord::FootnoteStart(_)
            | ControlWord::EndnoteStart(_)
            | ControlWord::FootnoteRestart(_)
            | ControlWord::EndnoteRestart(_)
            | ControlWord::FootnoteNumbering(_)
            | ControlWord::EndnoteNumbering(_)) => {
                if self.states.len() != 2
                    || self.note_options_closed
                    || self.blocks.iter().any(|block| !block.text.is_empty())
                {
                    return Err(RtfError::MalformedDocument(
                        "RTF note options must precede body text at document root".to_string(),
                    ));
                }
                match control {
                    ControlWord::NoteKinds(value) => {
                        self.note_options.present_kinds = Some(match value {
                            0 => crate::PresentNoteKinds::FootnotesOnly,
                            1 => crate::PresentNoteKinds::EndnotesOnly,
                            2 => crate::PresentNoteKinds::FootnotesAndEndnotes,
                            _ => {
                                return Err(RtfError::MalformedDocument(
                                    "RTF fet value must be between 0 and 2".to_string(),
                                ));
                            },
                        });
                    },
                    ControlWord::FootnotePlacement(value) => {
                        self.note_options.footnote_placement = Some(*value);
                    },
                    ControlWord::EndnotePlacement(value) => {
                        self.note_options.endnote_placement = Some(*value);
                    },
                    ControlWord::FootnoteStart(value) => {
                        if *value <= 0 {
                            return Err(RtfError::MalformedDocument(
                                "RTF footnote starting number must be positive".to_string(),
                            ));
                        }
                        self.note_options.footnote_start = Some(*value);
                    },
                    ControlWord::EndnoteStart(value) => {
                        if *value <= 0 {
                            return Err(RtfError::MalformedDocument(
                                "RTF endnote starting number must be positive".to_string(),
                            ));
                        }
                        self.note_options.endnote_start = Some(*value);
                    },
                    ControlWord::FootnoteRestart(value) => {
                        self.note_options.footnote_restart = Some(*value);
                    },
                    ControlWord::EndnoteRestart(value) => {
                        self.note_options.endnote_restart = Some(*value);
                    },
                    ControlWord::FootnoteNumbering(value) => {
                        self.note_options.footnote_numbering = Some(*value);
                    },
                    ControlWord::EndnoteNumbering(value) => {
                        self.note_options.endnote_numbering = Some(*value);
                    },
                    _ => unreachable!(),
                }
                return Ok(());
            },
            ControlWord::LegacyDrawingObject
            | ControlWord::LegacyTextBox
            | ControlWord::LegacyTextBoxText
            | ControlWord::LegacyAnchorXPage
            | ControlWord::LegacyAnchorXMargin
            | ControlWord::LegacyAnchorXColumn
            | ControlWord::LegacyAnchorYPage
            | ControlWord::LegacyAnchorYMargin
            | ControlWord::LegacyAnchorYParagraph
            | ControlWord::LegacyDrawingHeight(_)
            | ControlWord::LegacyTextBoxMargin(_)
            | ControlWord::LegacyDrawingX(_)
            | ControlWord::LegacyDrawingY(_)
            | ControlWord::LegacyDrawingWidth(_)
            | ControlWord::LegacyDrawingHeightSize(_)
            | ControlWord::LegacyTextLeftRightTopBottom
            | ControlWord::LegacyTextLeftRightTopBottomVertical
            | ControlWord::LegacyTextTopBottomRightLeft
            | ControlWord::LegacyTextTopBottomRightLeftVertical
            | ControlWord::LegacyTextBottomTopLeftRight => {
                return Err(RtfError::MalformedDocument(
                    "orphan RTF legacy drawing control outside do".to_string(),
                ));
            },
            ControlWord::GeneratedListText | ControlWord::LegacyGeneratedListText => {
                return Err(RtfError::MalformedDocument(
                    "RTF generated list marker must be a grouped body destination".to_string(),
                ));
            },
            ControlWord::XmlNamespace(_) => {
                return Err(RtfError::MalformedDocument(
                    "orphan RTF xmlns control outside xmlnstbl".to_string(),
                ));
            },
            ControlWord::ThemeData | ControlWord::ColorSchemeMapping => {
                return Err(RtfError::MalformedDocument(
                    "orphan RTF theme destination control".to_string(),
                ));
            },
            ControlWord::LatentStyles
            | ControlWord::LatentStyleMax(_)
            | ControlWord::LatentStyleLockedDefault(_)
            | ControlWord::LatentStyleSemiHiddenDefault(_)
            | ControlWord::LatentStyleUnhideUsedDefault(_)
            | ControlWord::LatentStyleQuickFormatDefault(_)
            | ControlWord::LatentStylePriorityDefault(_)
            | ControlWord::LatentStyleExceptions
            | ControlWord::LatentStyleLocked(_)
            | ControlWord::LatentStyleSemiHidden(_)
            | ControlWord::LatentStyleUnhideUsed(_)
            | ControlWord::LatentStyleQuickFormat(_)
            | ControlWord::LatentStylePriority(_) => {
                return Err(RtfError::MalformedDocument(
                    "orphan RTF latent-style control".to_string(),
                ));
            },
            ControlWord::DataStore => {
                return Err(RtfError::MalformedDocument(
                    "orphan RTF datastore destination control".to_string(),
                ));
            },
            ControlWord::MathProperties
            | ControlWord::MathBreakBinary(_)
            | ControlWord::MathBreakBinarySubtraction(_)
            | ControlWord::MathDefaultJustification(_)
            | ControlWord::MathDisplayDefaults(_)
            | ControlWord::MathInterEquationSpacing(_)
            | ControlWord::MathIntegralLimitPlacement(_)
            | ControlWord::MathIntraEquationSpacing(_)
            | ControlWord::MathLeftMargin(_)
            | ControlWord::MathFont(_)
            | ControlWord::MathNaryLimitPlacement(_)
            | ControlWord::MathPostSpacing(_)
            | ControlWord::MathPreSpacing(_)
            | ControlWord::MathRightMargin(_)
            | ControlWord::MathSmallFractions(_)
            | ControlWord::MathWrapIndent(_)
            | ControlWord::MathWrapRight(_) => {
                return Err(RtfError::MalformedDocument(
                    "orphan RTF document math-properties control".to_string(),
                ));
            },
            ControlWord::DefaultLanguage(value) => {
                let language = crate::LanguageId::from_rtf(*value)?;
                self.language_defaults.primary = Some(language);
                let state = self.current_state_mut()?;
                state.formatting.language = Some(language);
                state.formatting.language_no_proof = Some(language);
                return Ok(());
            },
            ControlWord::DefaultLanguageEastAsian(value) => {
                let language = crate::LanguageId::from_rtf(*value)?;
                self.language_defaults.east_asian = Some(language);
                let state = self.current_state_mut()?;
                state.formatting.east_asian_language = Some(language);
                state.formatting.east_asian_language_no_proof = Some(language);
                return Ok(());
            },
            ControlWord::DefaultLanguageComplexScript(value) => {
                let language = crate::LanguageId::from_rtf(*value)?;
                self.language_defaults.complex_script = Some(language);
                self.current_state_mut()?.formatting.associated.language = Some(language);
                return Ok(());
            },
            ControlWord::LeftToRightDocument => {
                self.document_direction = Some(TextDirection::LeftToRight);
                return Ok(());
            },
            ControlWord::RightToLeftDocument => {
                self.document_direction = Some(TextDirection::RightToLeft);
                return Ok(());
            },
            ControlWord::RightGutter(value) => {
                self.gutter_on_right = *value;
                return Ok(());
            },
            ControlWord::FormProtection(value) => {
                self.ensure_protection_scope()?;
                Self::set_protection_flag(&mut self.info.protection.forms, *value, "formprot")?;
                return Ok(());
            },
            ControlWord::AnnotationProtection(value) => {
                self.ensure_protection_scope()?;
                Self::set_protection_flag(&mut self.info.protection.annotations, *value, "annotprot")?;
                return Ok(());
            },
            ControlWord::RevisionProtection(value) => {
                self.ensure_protection_scope()?;
                Self::set_protection_flag(&mut self.info.protection.revisions, *value, "revprot")?;
                return Ok(());
            },
            ControlWord::ReadOnlyProtection(value) => {
                self.ensure_protection_scope()?;
                Self::set_protection_flag(&mut self.info.protection.read_only, *value, "readprot")?;
                return Ok(());
            },
            ControlWord::AllProtection(value) => {
                self.ensure_protection_scope()?;
                Self::set_protection_flag(&mut self.info.protection.all, *value, "allprot")?;
                return Ok(());
            },
            ControlWord::EnforceProtection(Some(value)) => {
                self.ensure_protection_scope()?;
                Self::set_required_protection_flag(&mut self.info.protection.enforced, *value, "enforceprot")?;
                return Ok(());
            },
            ControlWord::EnforceProtection(None) => {
                return Err(RtfError::MalformedDocument(
                    "RTF enforceprot requires a numeric parameter".to_string(),
                ));
            },
            ControlWord::ProtectionLevel(Some(value)) => {
                self.ensure_protection_scope()?;
                if self.info.protection.level.is_some() {
                    return Err(RtfError::MalformedDocument(
                        "duplicate RTF protlevel control".to_string(),
                    ));
                }
                self.info.protection.level = Some(crate::ProtectionLevel::from_rtf(*value)?);
                return Ok(());
            },
            ControlWord::ProtectionLevel(None) => {
                return Err(RtfError::MalformedDocument(
                    "RTF protlevel requires a numeric parameter".to_string(),
                ));
            },
            ControlWord::Password => {
                return Err(RtfError::MalformedDocument(
                    "RTF password hash is misplaced or not a starred info destination".to_string(),
                ));
            },
            _ => {},
        }
        if self.apply_section_control(control)? {
            return Ok(());
        }
        let language_defaults = self.language_defaults;
        let state = self.current_state_mut()?;

        if Self::apply_paragraph_tab_control(state, control)? {
            return Ok(());
        }

        if Self::apply_table_decoration_control(state, control)? {
            return Ok(());
        }

        if Self::apply_character_decoration_control(state, control)? {
            return Ok(());
        }

        match control {
            // Font formatting
            ControlWord::FontNumber(n) => {
                state.formatting.font_ref = *n as FontRef;
            },
            ControlWord::Language(value) => {
                state.formatting.language = Some(crate::LanguageId::from_rtf(*value)?);
            },
            ControlWord::LanguageEastAsian(value) => {
                state.formatting.east_asian_language =
                    Some(crate::LanguageId::from_rtf(*value)?);
            },
            ControlWord::LanguageNoProof(value) => {
                state.formatting.language_no_proof =
                    Some(crate::LanguageId::from_rtf(*value)?);
            },
            ControlWord::LanguageEastAsianNoProof(value) => {
                state.formatting.east_asian_language_no_proof =
                    Some(crate::LanguageId::from_rtf(*value)?);
            },
            ControlWord::NoProof(value) => state.formatting.no_proof = *value,
            ControlWord::LeftToRightCharacter => {
                state.formatting.direction = Some(TextDirection::LeftToRight);
            },
            ControlWord::RightToLeftCharacter => {
                state.formatting.direction = Some(TextDirection::RightToLeft);
            },
            ControlWord::FontSize(size) => {
                if let Some(nz) = NonZeroU16::new((*size).max(0) as u16) {
                    state.formatting.font_size = nz;
                }
            },
            ControlWord::AssociatedFontNumber(value) => {
                state.formatting.associated.font_ref = Some(associated_font_ref(*value)?);
            },
            ControlWord::AssociatedFontSize(value) => {
                state.formatting.associated.font_size = Some(associated_font_size(*value)?);
            },
            ControlWord::AssociatedLanguage(value) => {
                state.formatting.associated.language = Some(associated_language(*value)?);
            },
            ControlWord::AssociatedBold(value) => {
                state.formatting.associated.bold = Some(*value);
            },
            ControlWord::AssociatedItalic(value) => {
                state.formatting.associated.italic = Some(*value);
            },
            ControlWord::ColorForeground(c) => {
                state.formatting.color_ref = *c as ColorRef;
            },

            // Character formatting
            ControlWord::Bold(b) => state.formatting.bold = *b,
            ControlWord::Italic(b) => state.formatting.italic = *b,
            ControlWord::Underline(b) => {
                state.formatting.underline = if *b {
                    super::types::UnderlineStyle::Single
                } else {
                    super::types::UnderlineStyle::None
                }
            },
            ControlWord::UnderlineNone => {
                state.formatting.underline = super::types::UnderlineStyle::None
            },
            ControlWord::UnderlineDouble => {
                state.formatting.underline = super::types::UnderlineStyle::Double
            },
            ControlWord::UnderlineDotted => {
                state.formatting.underline = super::types::UnderlineStyle::Dotted
            },
            ControlWord::UnderlineDashed => {
                state.formatting.underline = super::types::UnderlineStyle::Dashed
            },
            ControlWord::UnderlineDashDot => {
                state.formatting.underline = super::types::UnderlineStyle::DashDot
            },
            ControlWord::UnderlineDashDotDot => {
                state.formatting.underline = super::types::UnderlineStyle::DashDotDot
            },
            ControlWord::UnderlineWords => {
                state.formatting.underline = super::types::UnderlineStyle::Words
            },
            ControlWord::UnderlineThick => {
                state.formatting.underline = super::types::UnderlineStyle::Thick
            },
            ControlWord::UnderlineWave => {
                state.formatting.underline = super::types::UnderlineStyle::Wave
            },
            ControlWord::Strike(b) => state.formatting.strike = *b,
            ControlWord::DoubleStrike(b) => state.formatting.double_strike = *b,
            ControlWord::Superscript(b) => { state.formatting.superscript = *b; if *b { state.formatting.subscript = false; } state.formatting.character_positioning.set_superscript(*b); },
            ControlWord::Subscript(b) => { state.formatting.subscript = *b; if *b { state.formatting.superscript = false; } state.formatting.character_positioning.set_subscript(*b); },
            ControlWord::NoSuperSub => { state.formatting.superscript = false; state.formatting.subscript = false; state.formatting.character_positioning.clear_baseline(); },
            ControlWord::BaselineUp(value) => { state.formatting.superscript = false; state.formatting.subscript = false; state.formatting.character_positioning.set_raised(*value)?; },
            ControlWord::BaselineDown(value) => { state.formatting.superscript = false; state.formatting.subscript = false; state.formatting.character_positioning.set_lowered(*value)?; },
            ControlWord::SmallCaps(b) => state.formatting.smallcaps = *b,
            ControlWord::AllCaps(b) => state.formatting.all_caps = *b,
            ControlWord::Hidden(b) => state.formatting.hidden = *b,
            ControlWord::Outline(b) => state.formatting.outline = *b,
            ControlWord::Shadow(b) => state.formatting.shadow = *b,
            ControlWord::Emboss(b) => state.formatting.emboss = *b,
            ControlWord::Imprint(b) => state.formatting.imprint = *b,
            ControlWord::CharSpacing(n) => { state.formatting.character_positioning.set_quarter_point_expansion(*n)?; state.formatting.char_spacing = *n; },
            ControlWord::CharSpacingTwips(n) => { state.formatting.character_positioning.set_twip_expansion(*n)?; state.formatting.char_spacing = *n; },
            ControlWord::CharScale(n) => { state.formatting.character_positioning.set_scale(*n)?; state.formatting.char_scale = *n; },
            ControlWord::Kerning(n) => { state.formatting.character_positioning.set_kerning(*n)?; state.formatting.kerning = *n; },
            ControlWord::Highlight(c) => state.formatting.highlight_color = Some(*c as ColorRef),
            ControlWord::Plain => {
                // Reset to default formatting
                state.formatting = Formatting::default();
                state.formatting.language = language_defaults.primary;
                state.formatting.east_asian_language = language_defaults.east_asian;
                state.formatting.language_no_proof = language_defaults.primary;
                state.formatting.east_asian_language_no_proof = language_defaults.east_asian;
                state.formatting.associated.language = language_defaults.complex_script;
                state.character_border_active = false;
                state.character_border_seen = 0;
            },

            // Paragraph alignment
            ControlWord::LeftAlign => state.paragraph.alignment = Alignment::Left,
            ControlWord::RightAlign => state.paragraph.alignment = Alignment::Right,
            ControlWord::Center => state.paragraph.alignment = Alignment::Center,
            ControlWord::Justify => state.paragraph.alignment = Alignment::Justify,
            ControlWord::LeftToRightParagraph => {
                state.paragraph.direction = Some(TextDirection::LeftToRight);
            },
            ControlWord::RightToLeftParagraph => {
                state.paragraph.direction = Some(TextDirection::RightToLeft);
            },
            ControlWord::Pard => {
                // Reset to default paragraph properties
                state.paragraph = Paragraph::default();
                state.pending_tab_alignment = None;
                state.pending_tab_leader = None;
                state.in_table = false;
            },

            // Paragraph spacing
            ControlWord::SpaceBefore(n) => state.paragraph.spacing.before = *n,
            ControlWord::SpaceAfter(n) => state.paragraph.spacing.after = *n,
            ControlWord::SpaceBetween(n) => state.paragraph.spacing.line = *n,
            ControlWord::LineMultiple(b) => state.paragraph.spacing.line_multiple = *b,
            ControlWord::SpaceBeforeAuto(value) => state.paragraph.spacing_policy.automatic_before = required_paragraph_bool(*value, "sbauto")?,
            ControlWord::SpaceAfterAuto(value) => state.paragraph.spacing_policy.automatic_after = required_paragraph_bool(*value, "saauto")?,
            ControlWord::ListSpaceBefore(value) => state.paragraph.spacing_policy.list_before = Some(required_list_spacing(*value, "lisb")?),
            ControlWord::ListSpaceAfter(value) => state.paragraph.spacing_policy.list_after = Some(required_list_spacing(*value, "lisa")?),
            ControlWord::NoSnapLineGrid(value) => { strict_paragraph_selector(*value, "nosnaplinegrid")?; state.paragraph.spacing_policy.snap_to_line_grid = false; },
            ControlWord::ContextualSpacing(value) => { strict_paragraph_selector(*value, "contextualspace")?; state.paragraph.spacing_policy.contextual_spacing = true; },

            // Paragraph indentation
            ControlWord::LeftIndent(n) => state.paragraph.indentation.left = *n,
            ControlWord::RightIndent(n) => state.paragraph.indentation.right = *n,
            ControlWord::FirstLineIndent(n) => state.paragraph.indentation.first_line = *n,
            ControlWord::LogicalLeftIndent(v)=>state.paragraph.logical_indentation.start=Some(required_paragraph_indent(*v,"lin")?), ControlWord::LogicalRightIndent(v)=>state.paragraph.logical_indentation.end=Some(required_paragraph_indent(*v,"rin")?), ControlWord::CharacterFirstLineIndent(v)=>state.paragraph.logical_indentation.first_line_character_units=Some(required_paragraph_indent(*v,"cufi")?), ControlWord::CharacterLeftIndent(v)=>state.paragraph.logical_indentation.left_character_units=Some(required_paragraph_indent(*v,"culi")?), ControlWord::CharacterRightIndent(v)=>state.paragraph.logical_indentation.right_character_units=Some(required_paragraph_indent(*v,"curi")?), ControlWord::MirrorIndents(v)=>{strict_paragraph_selector(*v,"indmirror")?;state.paragraph.logical_indentation.mirrored=true;},

            // Paragraph additional properties
            ControlWord::KeepTogether => state.paragraph.keep_together = true,
            ControlWord::KeepNext => state.paragraph.keep_next = true,
            ControlWord::PageBreakBefore => state.paragraph.page_break_before = true,
            ControlWord::WidowControl => state.paragraph.widow_control = true,
            ControlWord::ParagraphHyphenation(value) => state.paragraph.line_breaking.automatic_hyphenation = strict_paragraph_toggle(*value, "hyphpar")?,
            ControlWord::AutoSpaceAlphabetic(value) => state.paragraph.line_breaking.auto_space_alphabetic = strict_paragraph_toggle(*value, "aspalpha")?,
            ControlWord::AutoSpaceNumbers(value) => state.paragraph.line_breaking.auto_space_numbers = strict_paragraph_toggle(*value, "aspnum")?,
            ControlWord::AdjustRightIndent(value) => state.paragraph.line_breaking.adjust_right_indent = strict_paragraph_toggle(*value, "adjustright")?,
            ControlWord::WrapDefault(value) => { strict_paragraph_selector(*value, "wrapdefault")?; state.paragraph.line_breaking.wrapping = crate::ParagraphWrapping::Default; },
            ControlWord::NoCharacterWrap(value) => { strict_paragraph_selector(*value, "nocwrap")?; state.paragraph.line_breaking.wrapping = crate::ParagraphWrapping::NoCharacterWrap; },
            ControlWord::NoWordWrap(value) => { strict_paragraph_selector(*value, "nowwrap")?; state.paragraph.line_breaking.wrapping = crate::ParagraphWrapping::NoWordWrap; },
            ControlWord::NoOverflow(value) => { strict_paragraph_selector(*value, "nooverflow")?; state.paragraph.line_breaking.wrapping = crate::ParagraphWrapping::NoOverflow; },
            ControlWord::FontAlignAuto(value) => { strict_paragraph_selector(*value, "faauto")?; state.paragraph.line_breaking.font_alignment = crate::ParagraphFontAlignment::Auto; },
            ControlWord::FontAlignHanging(value) => { strict_paragraph_selector(*value, "fahang")?; state.paragraph.line_breaking.font_alignment = crate::ParagraphFontAlignment::Hanging; },
            ControlWord::FontAlignCenter(value) => { strict_paragraph_selector(*value, "facenter")?; state.paragraph.line_breaking.font_alignment = crate::ParagraphFontAlignment::Center; },
            ControlWord::FontAlignRoman(value) => { strict_paragraph_selector(*value, "faroman")?; state.paragraph.line_breaking.font_alignment = crate::ParagraphFontAlignment::Roman; },
            ControlWord::FontAlignVariable(value) => { strict_paragraph_selector(*value, "favar")?; state.paragraph.line_breaking.font_alignment = crate::ParagraphFontAlignment::Variable; },
            ControlWord::FontAlignFixed(value) => { strict_paragraph_selector(*value, "fafixed")?; state.paragraph.line_breaking.font_alignment = crate::ParagraphFontAlignment::Fixed; },
            ControlWord::ListOverrideIndex(value) => {
                state.paragraph.list_override = Some(*value);
            },
            ControlWord::ListLevelIndex(value) => {
                let level = u8::try_from(*value).map_err(|_| {
                    RtfError::MalformedDocument(
                        "RTF paragraph list level is outside the supported range".to_string(),
                    )
                })?;
                if level > 8 {
                    return Err(RtfError::MalformedDocument(
                        "RTF paragraph list level exceeds the nine-level specification limit"
                            .to_string(),
                    ));
                }
                state.paragraph.list_level = Some(level);
            },

            // Tracked revisions
            ControlWord::Revised(value) => {
                if *value {
                    if state.in_table {
                        return Err(RtfError::MalformedDocument(
                            "RTF positional revisions inside tables are unsupported".to_string(),
                        ));
                    }
                    if state.revision_type.is_some() {
                        return Err(RtfError::MalformedDocument(
                            "conflicting or duplicate RTF revision marker".to_string(),
                        ));
                    }
                    state.revision_type = Some(super::annotation::RevisionType::Insertion);
                } else if state.revision_type == Some(super::annotation::RevisionType::Insertion) {
                    state.revision_type = None;
                    state.revision_author_id = None;
                    state.revision_date = None;
                }
            },
            ControlWord::Deleted(value) => {
                if *value {
                    if state.in_table {
                        return Err(RtfError::MalformedDocument(
                            "RTF positional revisions inside tables are unsupported".to_string(),
                        ));
                    }
                    if state.revision_type.is_some() {
                        return Err(RtfError::MalformedDocument(
                            "conflicting or duplicate RTF revision marker".to_string(),
                        ));
                    }
                    state.revision_type = Some(super::annotation::RevisionType::Deletion);
                } else if state.revision_type == Some(super::annotation::RevisionType::Deletion) {
                    state.revision_type = None;
                    state.revision_author_id = None;
                    state.revision_date = None;
                }
            },
            ControlWord::RevisionAuthor(value) => {
                if state.revision_type.is_none() || state.revision_author_id.is_some() {
                    return Err(RtfError::MalformedDocument(
                        "RTF revauth requires one active revision marker".to_string(),
                    ));
                }
                state.revision_author_id = Some(*value);
            },
            ControlWord::DeletedRevisionAuthor(value) => {
                if state.revision_type != Some(super::annotation::RevisionType::Deletion)
                    || state.revision_author_id.is_some()
                {
                    return Err(RtfError::MalformedDocument(
                        "RTF revauthdel requires one active deletion marker".to_string(),
                    ));
                }
                state.revision_author_id = Some(*value);
            },
            ControlWord::RevisionDate(value) => {
                if state.revision_type.is_none() || state.revision_date.is_some() {
                    return Err(RtfError::MalformedDocument(
                        "RTF revdttm requires one active revision marker".to_string(),
                    ));
                }
                state.revision_date = Some(*value);
            },
            ControlWord::DeletedRevisionDate(value) => {
                if state.revision_type != Some(super::annotation::RevisionType::Deletion)
                    || state.revision_date.is_some()
                {
                    return Err(RtfError::MalformedDocument(
                        "RTF revdttmdel requires one active deletion marker".to_string(),
                    ));
                }
                state.revision_date = Some(*value);
            },

            // Unicode
            ControlWord::UnicodeSkip(n) => state.unicode_skip = *n,
            ControlWord::Unicode(code) => {
                // Unicode characters are handled separately during text parsing
                // since they may span multiple tokens with fallback characters
                // The control word itself doesn't add text here
                let _ = code; // Suppress unused warning
            },

            // Character encoding
            ControlWord::Ansi => {
                state.encoding = RtfEncoding::Standard(encoding_rs::WINDOWS_1252);
            },
            ControlWord::AnsiCodePage(cp) => {
                state.encoding = match *cp {
                    437 => RtfEncoding::Cp437,
                    850 => RtfEncoding::Cp850,
                    _ => {
                        if let Some(encoding) = codepage_to_encoding(*cp as u32) {
                            RtfEncoding::Standard(encoding)
                        } else {
                            state.encoding
                        }
                    },
                }
            },
            ControlWord::Mac => {
                state.encoding = RtfEncoding::Standard(encoding_rs::MACINTOSH);
            },
            ControlWord::Pc => state.encoding = RtfEncoding::Cp437,
            ControlWord::Pca => state.encoding = RtfEncoding::Cp850,

            // Table control words
            ControlWord::InTable => {
                state.in_table = true;
            },
            ControlWord::TableRowDefaults => {
                // Start a new row definition
                state.cell_boundaries.clear();
                state.table_row_padding=Default::default();state.table_row_spacing=Default::default();state.table_row_positioning=Default::default();state.table_row_direction=None;state.table_row_layout=Default::default();state.table_row_borders=Default::default();state.table_row_shading=Default::default();state.table_row_geometry=Default::default();state.table_width_unit=None;state.table_width_value=None;state.table_leading_width_unit=None;state.table_leading_width_value=None;state.table_trailing_width_unit=None;state.table_trailing_width_value=None;state.table_indent_value=None;state.table_indent_unit=None;state.pending_cell_padding=Default::default();state.pending_cell_spacing=Default::default();state.pending_cell_layout=Default::default();state.pending_cell_merge=Default::default();state.pending_cell_borders=Default::default();state.pending_cell_shading=Default::default();state.pending_cell_width_unit=None;state.pending_cell_width_value=None;state.table_row_shading_seen=0;state.pending_cell_shading_seen=0;state.active_table_border=None;state.active_table_border_seen=0;state.cell_distances.clear();state.cell_layouts.clear();state.cell_merges.clear();state.cell_decorations.clear();state.cell_widths.clear();
                let destination=state.destination;let level=state.table_nesting_level;let _=state;if destination==Destination::NestedTableProperties{if level<2{return Err(RtfError::MalformedDocument("RTF nesttableprops lacks itap level 2 or greater".to_string()))}let row=&mut self.ensure_nested_builder(level)?.row;row.set_direction(None);row.set_layout(Default::default());}else{self.drain_nested_to(1)?;self.start_table_if_needed();if let Some(row)=&mut self.current_row{row.set_direction(None);row.set_layout(Default::default());}}
            },
            ControlWord::LeftToRightRow(param) => {require_parameterless(*param,"ltrrow")?;state.table_row_direction=Some(TextDirection::LeftToRight)},
            ControlWord::RightToLeftRow(param) => {require_parameterless(*param,"rtlrow")?;state.table_row_direction=Some(TextDirection::RightToLeft)},
            ControlWord::TableRowGap(param)=>{let value=table_geometry_twips(*param,"trgaph",false)?;state.table_row_geometry.set_half_gap_twips(Some(value as u16));},
            ControlWord::TableRowLeft(param)=>{let value=table_geometry_twips(*param,"trleft",true)?;state.table_row_geometry.set_left_edge_twips(Some(value));},
            ControlWord::TableRowHeight(param)=>state.table_row_geometry.set_height(table_row_height(*param)?),
            ControlWord::TablePreferredWidthUnit(scope,param)=>{let unit=table_width_unit(*param)?;let target=match scope{crate::TableDistanceScope::Row=>&mut state.table_width_unit,crate::TableDistanceScope::Cell=>&mut state.pending_cell_width_unit};if target.replace(unit).is_some(){return Err(RtfError::MalformedDocument("RTF preferred-width unit is duplicated".to_string()))}},
            ControlWord::TablePreferredWidthValue(scope,param)=>{let value=param.ok_or_else(||RtfError::MalformedDocument("RTF preferred-width value requires a parameter".to_string()))?;let target=match scope{crate::TableDistanceScope::Row=>&mut state.table_width_value,crate::TableDistanceScope::Cell=>&mut state.pending_cell_width_value};if target.replace(value).is_some(){return Err(RtfError::MalformedDocument("RTF preferred-width value is duplicated".to_string()))}},
            ControlWord::TableInvisibleWidthUnit(trailing,param)=>{let unit=table_width_unit(*param)?;let target=if *trailing{&mut state.table_trailing_width_unit}else{&mut state.table_leading_width_unit};if target.replace(unit).is_some(){return Err(RtfError::MalformedDocument("RTF invisible-width unit is duplicated".to_string()))}},
            ControlWord::TableInvisibleWidthValue(trailing,param)=>{let value=param.ok_or_else(||RtfError::MalformedDocument("RTF invisible-width value requires a parameter".to_string()))?;let target=if *trailing{&mut state.table_trailing_width_value}else{&mut state.table_leading_width_value};if target.replace(value).is_some(){return Err(RtfError::MalformedDocument("RTF invisible-width value is duplicated".to_string()))}},
            ControlWord::TableAutoFit(param)=>state.table_row_geometry.set_auto_fit(match param{Some(0)=>false,Some(1)=>true,None=>return Err(RtfError::MalformedDocument("RTF trautofit requires 0 or 1".to_string())),Some(_)=>return Err(RtfError::MalformedDocument("RTF trautofit accepts only 0 or 1".to_string()))}),
            ControlWord::TableIndentValue(param)=>{let value=match param{None=>0,Some(_)=>table_geometry_twips(*param,"tblind",true)?};if state.table_indent_value.replace(value).is_some(){return Err(RtfError::MalformedDocument("RTF tblind is duplicated".to_string()))}},
            ControlWord::TableIndentUnit(param)=>{let unit=table_indent_unit(*param)?;if state.table_indent_unit.replace(unit).is_some(){return Err(RtfError::MalformedDocument("RTF tblindtype is duplicated".to_string()))}},
            ControlWord::TableRowHeader(param)=>{require_parameterless(*param,"trhdr")?;state.table_row_layout.header=true},
            ControlWord::TableRowKeep(param)=>{require_parameterless(*param,"trkeep")?;state.table_row_layout.keep_together=true},
            ControlWord::TableRowKeepFollow(param)=>{require_parameterless(*param,"trkeepfollow")?;state.table_row_layout.keep_with_following=true},
            ControlWord::TableRowAlignment(value,param)=>{require_parameterless(*param,"table row alignment")?;state.table_row_layout.alignment=Some(*value)},
            ControlWord::TableCellVerticalAlignment(value,param)=>{require_parameterless(*param,"cell vertical alignment")?;state.pending_cell_layout.vertical_alignment=Some(*value)},
            ControlWord::TableCellTextFlow(value,param)=>{require_parameterless(*param,"cell text flow")?;state.pending_cell_layout.text_flow=Some(*value)},
            ControlWord::TableCellFitText(param)=>{require_parameterless(*param,"clFitText")?;state.pending_cell_layout.fit_text=true},
            ControlWord::TableCellNoWrap(param)=>{require_parameterless(*param,"clNoWrap")?;state.pending_cell_layout.no_wrap=true},
            ControlWord::TableCellHideMark(param)=>{require_parameterless(*param,"clhidemark")?;state.pending_cell_layout.hide_mark=true},
            ControlWord::TableCellMerge(axis,role,param)=>{require_parameterless(*param,"table cell merge")?;let pending=match axis{crate::TableCellMergeAxis::Horizontal=>&mut state.pending_cell_merge.horizontal,crate::TableCellMergeAxis::Vertical=>&mut state.pending_cell_merge.vertical};if pending.replace(*role).is_some(){return Err(RtfError::MalformedDocument("RTF cell definition has duplicate or conflicting merge roles on one axis".to_string()))}},
            ControlWord::TableRightToLeft(value) => {
                let direction=Some(if *value{TextDirection::RightToLeft}else{TextDirection::LeftToRight});let destination=state.destination;let level=state.table_nesting_level;let _=state;if destination==Destination::NestedTableProperties{self.ensure_nested_builder(level)?.table.set_direction(direction);}else{self.start_table_if_needed();if let Some(table)=&mut self.current_table{table.set_direction(direction);}}
            },
            ControlWord::CellX(boundary) => {
                // Cell boundary definition
                state.cell_boundaries.push(*boundary);
                if state.cell_distances.len()>=crate::MAX_TABLE_CELLS_PER_ROW{return Err(RtfError::MalformedDocument("RTF row exceeds 4096 cell definitions".to_string()))}let width=resolve_preferred_width(state.pending_cell_width_unit.take(),state.pending_cell_width_value.take())?;state.cell_widths.push(width);state.cell_distances.push((std::mem::take(&mut state.pending_cell_padding),std::mem::take(&mut state.pending_cell_spacing)));state.cell_layouts.push(std::mem::take(&mut state.pending_cell_layout));state.cell_merges.push(std::mem::take(&mut state.pending_cell_merge));state.cell_decorations.push((std::mem::take(&mut state.pending_cell_borders),std::mem::take(&mut state.pending_cell_shading)));state.pending_cell_shading_seen=0;state.active_table_border=None;state.active_table_border_seen=0;
            },
            ControlWord::TableDistanceValue(target,value)=>apply_table_distance(state,*target,*value,false)?,
            ControlWord::TableDistanceUnit(target,value)=>apply_table_distance(state,*target,*value,true)?,
            ControlWord::TableHorizontalReference(value,param)=>if matches!(state.destination,Destination::DocumentBody|Destination::NestedTableProperties){require_parameterless(*param,"floating-table horizontal reference")?;state.table_row_positioning.horizontal_reference=Some(*value)},
            ControlWord::TableVerticalReference(value,param)=>if matches!(state.destination,Destination::DocumentBody|Destination::NestedTableProperties){require_parameterless(*param,"floating-table vertical reference")?;state.table_row_positioning.vertical_reference=Some(*value)},
            ControlWord::TableHorizontalPosition(value,param)=>if matches!(state.destination,Destination::DocumentBody|Destination::NestedTableProperties){require_parameterless(*param,"floating-table horizontal position")?;state.table_row_positioning.horizontal_position=Some(*value)},
            ControlWord::TableVerticalPosition(value,param)=>if matches!(state.destination,Destination::DocumentBody|Destination::NestedTableProperties){require_parameterless(*param,"floating-table vertical position")?;state.table_row_positioning.vertical_position=Some(*value)},
            ControlWord::TableHorizontalOffset(negative,param)=>if matches!(state.destination,Destination::DocumentBody|Destination::NestedTableProperties){let value=floating_table_offset(*param,*negative,"horizontal")?;state.table_row_positioning.horizontal_position=Some(if *negative{crate::TableHorizontalPosition::NegativeOffset(value)}else{crate::TableHorizontalPosition::Offset(value)})},
            ControlWord::TableVerticalOffset(negative,param)=>if matches!(state.destination,Destination::DocumentBody|Destination::NestedTableProperties){let value=floating_table_offset(*param,*negative,"vertical")?;state.table_row_positioning.vertical_position=Some(if *negative{crate::TableVerticalPosition::NegativeOffset(value)}else{crate::TableVerticalPosition::Offset(value)})},
            ControlWord::TableWrapDistance(edge,param)=>if matches!(state.destination,Destination::DocumentBody|Destination::NestedTableProperties){*state.table_row_positioning.wrap_distances.side_mut(*edge)=Some(floating_table_wrap_distance(*param)?)},
            ControlWord::TableNoOverlap(param)=>if matches!(state.destination,Destination::DocumentBody|Destination::NestedTableProperties){state.table_row_positioning.no_overlap=match *param{None|Some(1)=>true,Some(0)=>false,Some(_)=>return Err(RtfError::MalformedDocument("RTF tabsnoovrlp accepts only 0 or 1".to_string()))}},
            ControlWord::NestedTableCell(param)=>{require_parameterless(*param,"nestcell")?;let destination=state.destination;let level=state.table_nesting_level;let _=state;if destination!=Destination::DocumentBody||level<2{return Err(RtfError::MalformedDocument("RTF nestcell requires visible nested-table text and itap 2 or greater".to_string()))}self.finalize_nested_cell(level)?;},
            ControlWord::NestedTableRow(param)=>{require_parameterless(*param,"nestrow")?;let destination=state.destination;let level=state.table_nesting_level;let _=state;if destination!=Destination::NestedTableProperties||level<2{return Err(RtfError::MalformedDocument("RTF nestrow requires a nesttableprops destination and itap 2 or greater".to_string()))}self.finalize_nested_row(level)?;},
            ControlWord::NestedTableProperties(_)|ControlWord::NoNestedTables(_)=>return Err(RtfError::MalformedDocument("RTF nested-table destination control is misplaced".to_string())),
            ControlWord::TableCell => {
                // Cell break - finalize current cell
                self.start_table_if_needed();
                self.finalize_cell(true)?;
            },
            ControlWord::TableRow => {
                // Row break - finalize current row
                let row_geometry=resolve_row_geometry(state)?;let row_padding=state.table_row_padding.clone();let row_spacing=state.table_row_spacing.clone();let row_positioning=state.table_row_positioning.clone();let row_direction=state.table_row_direction;let row_layout=state.table_row_layout;let row_borders=state.table_row_borders.clone();let row_shading=state.table_row_shading;let boundaries=state.cell_boundaries.clone();let cell_distances=state.cell_distances.clone();let cell_layouts=state.cell_layouts.clone();let cell_merges=state.cell_merges.clone();let cell_decorations=state.cell_decorations.clone();let cell_widths=state.cell_widths.clone();let _=state;self.drain_nested_to(1)?;self.finalize_cell(false)?;if let Some(row)=&mut self.current_row{if !boundaries.is_empty()&&boundaries.len()!=row.cell_count(){return Err(RtfError::MalformedDocument("RTF row cell boundaries do not match cell count".to_string()))}for(index,cell)in row.cells_mut().iter_mut().enumerate(){if let Some((padding,spacing))=cell_distances.get(index){cell.set_padding(padding.clone());cell.set_spacing(spacing.clone());}if let Some(layout)=cell_layouts.get(index){cell.set_layout(*layout);}if let Some(merge)=cell_merges.get(index){cell.set_merge(*merge);}cell.set_right_boundary(boundaries.get(index).copied());cell.set_preferred_width(cell_widths.get(index).copied().flatten());if let Some((borders,shading))=cell_decorations.get(index){cell.set_borders(borders.clone());cell.set_shading(*shading);}}row.set_direction(row_direction);row.set_layout(row_layout);row.set_padding(row_padding);row.set_spacing(row_spacing);row.set_positioning(row_positioning);row.set_borders(row_borders);row.set_shading(row_shading);row.set_geometry(row_geometry);}
                self.finalize_row()?;
            },

            _ => {
                // Ignore unknown or unhandled control words
            },
        }

        Ok(())
    }

    fn apply_section_control(&mut self, control: &ControlWord<'_>) -> RtfResult<bool> {
        use super::section::{
            PageNumberFormat, PageOrientation, SectionBreakType, VerticalAlignment,
        };

        let is_line_numbering_control = matches!(
            control,
            ControlWord::LineNumbering(_)
                | ControlWord::LineNumberDistance(_)
                | ControlWord::LineNumberStart(_)
                | ControlWord::LineNumberRestartSection
                | ControlWord::LineNumberRestartPage
                | ControlWord::LineNumberContinuous
        );
        if is_line_numbering_control
            && self.current_state()?.destination != Destination::DocumentBody
        {
            return Ok(true);
        }

        if let Some(side) = match control {
            ControlWord::PageBorderTop => Some(crate::PageBorderSide::Top),
            ControlWord::PageBorderLeft => Some(crate::PageBorderSide::Left),
            ControlWord::PageBorderBottom => Some(crate::PageBorderSide::Bottom),
            ControlWord::PageBorderRight => Some(crate::PageBorderSide::Right),
            _ => None,
        } {
            let border = self.parse_page_border_run()?;
            if self.sections.is_empty() {
                if self.sections.len() >= MAX_SECTIONS {
                    return Err(RtfError::MalformedDocument("RTF section count exceeds limit".to_string()));
                }
                self.sections.push(super::section::Section::new());
            }
            let section = self.sections.last_mut().ok_or_else(|| RtfError::MalformedDocument("no active RTF section".to_string()))?;
            if section.properties.page_borders.get(side).is_some() {
                return Err(RtfError::MalformedDocument("duplicate RTF page-border edge".to_string()));
            }
            section.properties.page_borders.set(side, border);
            self.section_properties_active = true;
            return Ok(true);
        }

        if matches!(control, ControlWord::Section) {
            self.current_state_mut()?.section_column_number = None;
            self.section_properties_active = false;
            self.section_note_options_closed = false;
            self.root_section_format_run = false;
            return Ok(true);
        }
        let is_section_note_control = matches!(
            control,
            ControlWord::SectionFootnotePlacement(_)
                | ControlWord::SectionFootnoteStart(_)
                | ControlWord::SectionEndnoteStart(_)
                | ControlWord::SectionFootnoteRestart(_)
                | ControlWord::SectionEndnoteRestart(_)
                | ControlWord::SectionFootnoteNumbering(_)
                | ControlWord::SectionEndnoteNumbering(_)
        );
        if !is_section_control(control) {
            return Ok(false);
        }
        let in_root_document_body = self.states.len() == 2
            && self
                .states
                .last()
                .is_some_and(|state| state.destination == Destination::DocumentBody);
        if matches!(control, ControlWord::SectionDefault) && in_root_document_body {
            self.root_section_format_run = true;
        } else if matches!(control, ControlWord::SectionBreak) {
            self.root_section_format_run = false;
        }
        let in_visible_section_format = self.states.last().is_some_and(|state| {
            state.destination == Destination::DocumentBody && state.visible_section_format
        });
        let in_root_section_prefix = self.states.len() == 2
            && !self.section_note_options_closed
            && self
                .sections
                .last()
                .is_none_or(|section| section.headers_footers.is_empty());
        let in_root_section_format_run =
            in_root_document_body && self.root_section_format_run;
        if is_section_note_control
            && !in_root_section_prefix
            && !in_visible_section_format
            && !in_root_section_format_run
        {
            return Err(RtfError::MalformedDocument(
                "RTF section note options must precede section content at document root"
                    .to_string(),
            ));
        }

        if !self.section_properties_active {
            if self.sections.len() >= MAX_SECTIONS {
                return Err(RtfError::MalformedDocument(
                    "RTF section count exceeds the safety limit".to_string(),
                ));
            }
            let inherited = self
                .sections
                .last()
                .map(|section| section.properties.clone())
                .unwrap_or_default();
            let mut section = super::section::Section::new();
            section.properties = inherited;
            self.sections.push(section);
            self.section_properties_active = true;
        }
        let section = self.sections.last_mut().ok_or_else(|| {
            RtfError::ParserError("failed to create RTF section state".to_string())
        })?;
        let properties = &mut section.properties;
        match control {
            ControlWord::SectionDefault => {
                *properties = super::section::SectionProperties::default();
                self.states.last_mut().ok_or_else(|| {
                    RtfError::ParserError("missing RTF parser state".to_string())
                })?.section_column_number = None;
            },
            ControlWord::PageBorderOptions(value) => {
                let value = value.ok_or_else(|| RtfError::MalformedDocument("RTF pgbrdropt requires a numeric parameter".to_string()))?;
                properties.page_borders.set_option_value(value)?;
            },
            ControlWord::PageBorderSurroundHeader => properties.page_borders.surround_header = true,
            ControlWord::PageBorderSurroundFooter => properties.page_borders.surround_footer = true,
            ControlWord::PageBorderSnap => properties.page_borders.snap_to_text_borders = true,
            ControlWord::SectionFootnotePlacement(value) => {
                properties.note_options.footnote_placement = Some(*value);
            },
            ControlWord::SectionFootnoteStart(value) => {
                if *value <= 0 {
                    return Err(RtfError::MalformedDocument(
                        "RTF section footnote starting number must be positive".to_string(),
                    ));
                }
                properties.note_options.footnote_start = Some(*value);
            },
            ControlWord::SectionEndnoteStart(value) => {
                if *value <= 0 {
                    return Err(RtfError::MalformedDocument(
                        "RTF section endnote starting number must be positive".to_string(),
                    ));
                }
                properties.note_options.endnote_start = Some(*value);
            },
            ControlWord::SectionFootnoteRestart(value) => {
                properties.note_options.footnote_restart = Some(*value);
            },
            ControlWord::SectionEndnoteRestart(value) => {
                properties.note_options.endnote_restart = Some(*value);
            },
            ControlWord::SectionFootnoteNumbering(value) => {
                properties.note_options.footnote_numbering = Some(*value);
            },
            ControlWord::SectionEndnoteNumbering(value) => {
                properties.note_options.endnote_numbering = Some(*value);
            },
            ControlWord::LeftToRightSection => {
                properties.direction = Some(TextDirection::LeftToRight);
            },
            ControlWord::RightToLeftSection => {
                properties.direction = Some(TextDirection::RightToLeft);
            },
            ControlWord::SectionBreak | ControlWord::SectionPage => {
                properties.break_type = SectionBreakType::Page;
            },
            ControlWord::SectionContinuous => {
                properties.break_type = SectionBreakType::Continuous;
            },
            ControlWord::SectionColumn => properties.break_type = SectionBreakType::Column,
            ControlWord::SectionEvenPage => properties.break_type = SectionBreakType::EvenPage,
            ControlWord::SectionOddPage => properties.break_type = SectionBreakType::OddPage,
            ControlWord::PageWidth(value) => properties.page_width = *value,
            ControlWord::PageHeight(value) => properties.page_height = *value,
            ControlWord::MarginLeft(value) => properties.margin_left = *value,
            ControlWord::MarginRight(value) => properties.margin_right = *value,
            ControlWord::MarginTop(value) => properties.margin_top = *value,
            ControlWord::MarginBottom(value) => properties.margin_bottom = *value,
            ControlWord::MarginGutter(value) => properties.margin_gutter = *value,
            ControlWord::HeaderDistance(value) => properties.header_distance = *value,
            ControlWord::FooterDistance(value) => properties.footer_distance = *value,
            ControlWord::Landscape => properties.orientation = PageOrientation::Landscape,
            ControlWord::Columns(value) => {
                let value = value.unwrap_or(1);
                let count = u16::try_from(value).map_err(|_| {
                    RtfError::MalformedDocument(format!(
                        "RTF section-column count must be in 1..={}",
                        super::section::MAX_SECTION_COLUMNS
                    ))
                })?;
                if !(1..=super::section::MAX_SECTION_COLUMNS).contains(&count) {
                    return Err(RtfError::MalformedDocument(format!(
                        "RTF section-column count must be in 1..={}",
                        super::section::MAX_SECTION_COLUMNS
                    )));
                }
                properties.columns.count = count;
                properties.columns.explicit.clear();
                self.states.last_mut().ok_or_else(|| {
                    RtfError::ParserError("missing RTF parser state".to_string())
                })?.section_column_number = None;
            },
            ControlWord::ColumnSpace(value) => {
                let value = value.unwrap_or(720);
                if !(0..=super::section::MAX_SECTION_COLUMN_TWIPS).contains(&value) {
                    return Err(RtfError::MalformedDocument(
                        "RTF section-column default spacing must be in 0..=31680 twips".to_string(),
                    ));
                }
                properties.columns.default_spacing = value;
            },
            ControlWord::ColumnSeparator(value) => properties.columns.separator = *value,
            ControlWord::ColumnNumber(value) => {
                let value = value.ok_or_else(|| RtfError::MalformedDocument(
                    "RTF colno requires a numeric parameter".to_string(),
                ))?;
                let number = u16::try_from(value).map_err(|_| RtfError::MalformedDocument(
                    "RTF colno must select an existing one-based section column".to_string(),
                ))?;
                let expected = u16::try_from(properties.columns.explicit.len() + 1)
                    .unwrap_or(u16::MAX);
                if number != expected || number > properties.columns.count {
                    return Err(RtfError::MalformedDocument(
                        "RTF explicit section columns must use sequential one-based colno values"
                            .to_string(),
                    ));
                }
                properties.columns.explicit.push(super::section::SectionColumn {
                    width: 0,
                    space_after: None,
                });
                self.states.last_mut().ok_or_else(|| {
                    RtfError::ParserError("missing RTF parser state".to_string())
                })?.section_column_number = Some(number);
            },
            ControlWord::ColumnWidth(value) => {
                let value = value.ok_or_else(|| RtfError::MalformedDocument(
                    "RTF colw requires a numeric parameter".to_string(),
                ))?;
                if !(1..=super::section::MAX_SECTION_COLUMN_TWIPS).contains(&value) {
                    return Err(RtfError::MalformedDocument(
                        "RTF section-column width must be in 1..=31680 twips".to_string(),
                    ));
                }
                let number = self.states.last().and_then(|state| state.section_column_number)
                    .ok_or_else(|| RtfError::MalformedDocument(
                        "RTF colw requires a preceding colno in the active group".to_string(),
                    ))?;
                let column = properties.columns.explicit.get_mut(usize::from(number - 1))
                    .ok_or_else(|| RtfError::MalformedDocument(
                        "RTF colw refers to an undefined section column".to_string(),
                    ))?;
                if column.width != 0 {
                    return Err(RtfError::MalformedDocument(
                        "duplicate RTF colw for one section column".to_string(),
                    ));
                }
                column.width = value;
            },
            ControlWord::ColumnSpaceRight(value) => {
                let value = value.ok_or_else(|| RtfError::MalformedDocument(
                    "RTF colsr requires a numeric parameter".to_string(),
                ))?;
                if !(0..=super::section::MAX_SECTION_COLUMN_TWIPS).contains(&value) {
                    return Err(RtfError::MalformedDocument(
                        "RTF section-column spacing must be in 0..=31680 twips".to_string(),
                    ));
                }
                let number = self.states.last().and_then(|state| state.section_column_number)
                    .ok_or_else(|| RtfError::MalformedDocument(
                        "RTF colsr requires a preceding colno in the active group".to_string(),
                    ))?;
                let column = properties.columns.explicit.get_mut(usize::from(number - 1))
                    .ok_or_else(|| RtfError::MalformedDocument(
                        "RTF colsr refers to an undefined section column".to_string(),
                    ))?;
                if column.width == 0 {
                    return Err(RtfError::MalformedDocument(
                        "RTF colsr must follow colw for its section column".to_string(),
                    ));
                }
                if column.space_after.replace(value).is_some() {
                    return Err(RtfError::MalformedDocument(
                        "duplicate RTF colsr for one section column".to_string(),
                    ));
                }
            },
            ControlWord::PageNumberStart(value) => properties.page_number_start = *value,
            ControlWord::PageNumberDecimal => {
                properties.page_number_format = PageNumberFormat::Decimal;
            },
            ControlWord::PageNumberUpperRoman => {
                properties.page_number_format = PageNumberFormat::UpperRoman;
            },
            ControlWord::PageNumberLowerRoman => {
                properties.page_number_format = PageNumberFormat::LowerRoman;
            },
            ControlWord::PageNumberUpperLetter => {
                properties.page_number_format = PageNumberFormat::UpperLetter;
            },
            ControlWord::PageNumberLowerLetter => {
                properties.page_number_format = PageNumberFormat::LowerLetter;
            },
            ControlWord::VerticalAlignTop => {
                properties.vertical_alignment = VerticalAlignment::Top;
            },
            ControlWord::VerticalAlignCenter => {
                properties.vertical_alignment = VerticalAlignment::Center;
            },
            ControlWord::VerticalAlignJustify => {
                properties.vertical_alignment = VerticalAlignment::Justify;
            },
            ControlWord::VerticalAlignBottom => {
                properties.vertical_alignment = VerticalAlignment::Bottom;
            },
            ControlWord::LineNumbering(value) => {
                let value = value.unwrap_or(1);
                if value < 0 || value > i32::from(super::section::MAX_SECTION_LINE_INCREMENT) {
                    return Err(RtfError::MalformedDocument(
                        "RTF line-number increment must be in 0..=65535".to_string(),
                    ));
                }
                properties.line_numbering.increment = if value == 0 {
                    None
                } else {
                    Some(value as u16)
                };
            },
            ControlWord::LineNumberDistance(value) => {
                let value = value.unwrap_or(360);
                if !(0..=super::section::MAX_SECTION_LINE_DISTANCE).contains(&value) {
                    return Err(RtfError::MalformedDocument(
                        "RTF line-number distance must be in 0..=31680 twips".to_string(),
                    ));
                }
                properties.line_numbering.distance = Some(value);
            },
            ControlWord::LineNumberStart(value) => {
                let value = value.unwrap_or(1);
                if value <= 0 || value as u32 > super::section::MAX_SECTION_LINE_START {
                    return Err(RtfError::MalformedDocument(format!(
                        "RTF starting line number must be in 1..={}",
                        super::section::MAX_SECTION_LINE_START
                    )));
                }
                properties.line_numbering.start = Some(value as u32);
            },
            ControlWord::LineNumberRestartSection => {
                properties.line_numbering.restart =
                    Some(super::section::SectionLineNumberRestart::Section);
            },
            ControlWord::LineNumberRestartPage => {
                properties.line_numbering.restart =
                    Some(super::section::SectionLineNumberRestart::Page);
            },
            ControlWord::LineNumberContinuous => {
                properties.line_numbering.restart =
                    Some(super::section::SectionLineNumberRestart::Continuous);
            },
            _ => {},
        }
        Ok(true)
    }

    fn parse_page_border_run(&mut self) -> RtfResult<crate::PageBorder> {
        let mut border = crate::PageBorder::default();
        let mut saw_style = false;
        let mut seen = 0u8;
        loop {
            let Some(Token::Control(control)) = self.tokens.get(self.pos) else { break; };
            let style = match control {
                ControlWord::BorderNone => Some(crate::PageBorderStyle::None),
                ControlWord::BorderSingle => Some(crate::PageBorderStyle::Single),
                ControlWord::BorderThick => Some(crate::PageBorderStyle::Thick),
                ControlWord::BorderDotted => Some(crate::PageBorderStyle::Dotted),
                ControlWord::BorderDashed => Some(crate::PageBorderStyle::Dashed),
                ControlWord::BorderDashSmall => Some(crate::PageBorderStyle::DashSmallGap),
                ControlWord::BorderDotDash => Some(crate::PageBorderStyle::DotDash),
                ControlWord::BorderDotDotDash => Some(crate::PageBorderStyle::DotDotDash),
                ControlWord::BorderDouble => Some(crate::PageBorderStyle::Double),
                ControlWord::BorderTriple => Some(crate::PageBorderStyle::Triple),
                ControlWord::BorderThinThickSmall => Some(crate::PageBorderStyle::ThinThickSmallGap),
                ControlWord::BorderThickThinSmall => Some(crate::PageBorderStyle::ThickThinSmallGap),
                ControlWord::BorderThinThickThinSmall => Some(crate::PageBorderStyle::ThinThickThinSmallGap),
                ControlWord::BorderThinThickMedium => Some(crate::PageBorderStyle::ThinThickMediumGap),
                ControlWord::BorderThickThinMedium => Some(crate::PageBorderStyle::ThickThinMediumGap),
                ControlWord::BorderThinThickThinMedium => Some(crate::PageBorderStyle::ThinThickThinMediumGap),
                ControlWord::BorderThinThickLarge => Some(crate::PageBorderStyle::ThinThickLargeGap),
                ControlWord::BorderThickThinLarge => Some(crate::PageBorderStyle::ThickThinLargeGap),
                ControlWord::BorderThinThickThinLarge => Some(crate::PageBorderStyle::ThinThickThinLargeGap),
                ControlWord::BorderWave => Some(crate::PageBorderStyle::Wavy),
                ControlWord::BorderWavyDouble => Some(crate::PageBorderStyle::DoubleWavy),
                ControlWord::BorderStriped => Some(crate::PageBorderStyle::Striped),
                ControlWord::BorderEmbossed => Some(crate::PageBorderStyle::Embossed),
                ControlWord::BorderEngraved => Some(crate::PageBorderStyle::Engraved),
                ControlWord::BorderOutset => Some(crate::PageBorderStyle::Outset),
                ControlWord::BorderInset => Some(crate::PageBorderStyle::Inset),
                _ => None,
            };
            if let Some(style) = style {
                if saw_style { return Err(RtfError::MalformedDocument("duplicate RTF page-border style".to_string())); }
                saw_style = true;
                border.style = style;
                self.pos += 1;
                continue;
            }
            match control {
                ControlWord::PageBorderArt(value) => {
                    if saw_style { return Err(RtfError::MalformedDocument("duplicate RTF page-border style/art".to_string())); }
                    let value = value.ok_or_else(|| RtfError::MalformedDocument("RTF brdrart requires a numeric parameter".to_string()))?;
                    border.art = Some(u8::try_from(value).map_err(|_| RtfError::MalformedDocument("invalid RTF page-border art".to_string()))?);
                    saw_style = true;
                },
                ControlWord::BorderWidth(value) => {
                    if !saw_style || seen & 1 != 0 { return Err(RtfError::MalformedDocument("invalid or duplicate RTF page-border width".to_string())); }
                    border.width = u8::try_from(value.ok_or_else(|| RtfError::MalformedDocument("RTF page brdrw requires a numeric parameter".to_string()))?).map_err(|_| RtfError::MalformedDocument("invalid RTF page-border width".to_string()))?;
                    seen |= 1;
                },
                ControlWord::BorderColor(value) => {
                    if !saw_style || seen & 2 != 0 { return Err(RtfError::MalformedDocument("invalid or duplicate RTF page-border color".to_string())); }
                    border.color_ref = u16::try_from(value.ok_or_else(|| RtfError::MalformedDocument("RTF page brdrcf requires a numeric parameter".to_string()))?).map_err(|_| RtfError::MalformedDocument("invalid RTF page-border color".to_string()))?;
                    seen |= 2;
                },
                ControlWord::BorderSpace(value) => {
                    if !saw_style || seen & 4 != 0 { return Err(RtfError::MalformedDocument("invalid or duplicate RTF page-border spacing".to_string())); }
                    border.space = u16::try_from(value.ok_or_else(|| RtfError::MalformedDocument("RTF page brsp requires a numeric parameter".to_string()))?).map_err(|_| RtfError::MalformedDocument("invalid RTF page-border spacing".to_string()))?;
                    seen |= 4;
                },
                ControlWord::BorderShadow => {
                    if !saw_style || seen & 8 != 0 { return Err(RtfError::MalformedDocument("invalid or duplicate RTF page-border shadow".to_string())); }
                    border.shadow = true; seen |= 8;
                },
                ControlWord::BorderFrame => {
                    if !saw_style || seen & 16 != 0 { return Err(RtfError::MalformedDocument("invalid or duplicate RTF page-border frame".to_string())); }
                    border.frame = true; seen |= 16;
                },
                _ => break,
            }
            self.pos += 1;
        }
        if !saw_style { return Err(RtfError::MalformedDocument("RTF page-border edge requires a style or art control".to_string())); }
        border.validate()?;
        Ok(border)
    }

    fn parse_revision_table(&mut self) -> RtfResult<()> {
        if self.saw_revision_table {
            return Err(RtfError::MalformedDocument(
                "RTF contains multiple revision-author tables".to_string(),
            ));
        }
        self.saw_revision_table = true;
        self.pos += 1; // `revtbl`
        let mut direct_text = String::new();
        let mut unicode_skip = self.current_state()?.unicode_skip.max(0);
        while self.pos < self.tokens.len() {
            match self.tokens.get(self.pos) {
                Some(Token::OpenBrace) => {
                    self.push_direct_revision_authors(&mut direct_text)?;
                    let author = self.parse_revision_author_group()?;
                    self.push_revision_author(author)?;
                    continue;
                },
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    self.push_direct_revision_authors(&mut direct_text)?;
                    if !direct_text.trim().is_empty() {
                        self.push_revision_author(direct_text.trim().to_string())?;
                    }
                    return Ok(());
                },
                Some(Token::Text(text)) => direct_text.push_str(&self.decode_transport_text(text)?),
                Some(Token::Control(ControlWord::Unicode(first))) => {
                    let decoded = self.parse_style_unicode(*first, unicode_skip)?;
                    direct_text.push_str(&decoded);
                    continue;
                },
                Some(Token::Control(ControlWord::UnicodeSkip(value))) => {
                    unicode_skip = (*value).max(0);
                },
                Some(Token::Control(control)) if control_symbol_text(control).is_some() => {
                    direct_text.push_str(control_symbol_text(control).unwrap_or_default());
                },
                Some(Token::Control(_)) | Some(Token::Binary(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF revision-author table contains a non-text control or binary data"
                            .to_string(),
                    ));
                },
                None => break,
            }
            self.pos += 1;
            if direct_text.len() > MAX_REVISION_AUTHOR_BYTES {
                return Err(RtfError::MalformedDocument(
                    "RTF revision author exceeds the safety limit".to_string(),
                ));
            }
            self.push_direct_revision_authors(&mut direct_text)?;
        }
        Err(RtfError::UnexpectedEof)
    }

    fn parse_revision_save_table(&mut self) -> RtfResult<()> {
        if self.saw_revision_save_table {
            return Err(RtfError::MalformedDocument(
                "RTF contains multiple revision-save tables".to_string(),
            ));
        }
        self.saw_revision_save_table = true;
        self.pos += 1; // rsidtbl
        while self.pos < self.tokens.len() {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    return Ok(());
                },
                Some(Token::Control(ControlWord::RevisionSaveId(value))) => {
                    let value = u32::try_from(*value).map_err(|_| {
                        RtfError::MalformedDocument(
                            "RTF revision-save IDs must be positive signed integers".to_string(),
                        )
                    })?;
                    if value == 0 {
                        return Err(RtfError::MalformedDocument(
                            "RTF revision-save IDs must be positive signed integers".to_string(),
                        ));
                    }
                    if self.revision_save_ids.len()
                        >= crate::revision_save::MAX_REVISION_SAVE_IDS
                    {
                        return Err(RtfError::MalformedDocument(
                            "RTF revision-save ID count exceeds the safety limit".to_string(),
                        ));
                    }
                    if self.revision_save_ids.contains(&value) {
                        return Err(RtfError::MalformedDocument(
                            "RTF revision-save IDs must be unique".to_string(),
                        ));
                    }
                    self.revision_save_ids.push(value);
                    self.pos += 1;
                },
                Some(Token::Text(text)) if self.decode_transport_text(text)?.trim().is_empty() => {
                    self.pos += 1;
                },
                Some(Token::OpenBrace | Token::Binary(_))
                | Some(Token::Control(_))
                | Some(Token::Text(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF revision-save table contains text, nesting, binary data, or an unsupported control"
                            .to_string(),
                    ));
                },
                None => return Err(RtfError::UnexpectedEof),
            }
        }
        Err(RtfError::UnexpectedEof)
    }

    fn parse_data_store_destination(&mut self) -> RtfResult<Vec<u8>> {
        self.pos += 1; // ignorable-destination marker
        if !matches!(
            self.tokens.get(self.pos),
            Some(Token::Control(ControlWord::DataStore))
        ) {
            return Err(RtfError::MalformedDocument(
                "invalid RTF datastore destination".to_string(),
            ));
        }
        self.pos += 1;
        let mut data = Vec::new();
        let mut high = None;
        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    if high.is_some() {
                        return Err(RtfError::MalformedDocument(
                            "RTF data-store payload has an odd hexadecimal digit count"
                                .to_string(),
                        ));
                    }
                    if data.is_empty() {
                        return Err(RtfError::MalformedDocument(
                            "RTF data-store payload cannot be empty".to_string(),
                        ));
                    }
                    return Ok(data);
                },
                Some(Token::Text(text)) => {
                    for byte in text.as_bytes() {
                        if byte.is_ascii_whitespace() {
                            continue;
                        }
                        let nibble = match byte {
                            b'0'..=b'9' => byte - b'0',
                            b'a'..=b'f' => byte - b'a' + 10,
                            b'A'..=b'F' => byte - b'A' + 10,
                            _ => {
                                return Err(RtfError::MalformedDocument(
                                    "RTF data-store payload contains a non-hexadecimal character"
                                        .to_string(),
                                ));
                            },
                        };
                        if let Some(first) = high.take() {
                            data.push(first << 4 | nibble);
                        } else {
                            high = Some(nibble);
                        }
                    }
                    self.pos += 1;
                },
                Some(Token::OpenBrace | Token::Binary(_)) | Some(Token::Control(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF datastore cannot contain controls, nesting, or binary data"
                            .to_string(),
                    ));
                },
                None => return Err(RtfError::UnexpectedEof),
            }
            if data.len() > crate::data_store::MAX_DATA_STORE_BYTES {
                return Err(RtfError::MalformedDocument(
                    "RTF data-store payload exceeds the safety limit".to_string(),
                ));
            }
        }
    }

    fn parse_math_properties_destination(
        &mut self,
    ) -> RtfResult<crate::DocumentMathProperties> {
        if matches!(
            self.tokens.get(self.pos),
            Some(Token::Control(ControlWord::IgnorableDestination))
        ) {
            self.pos += 1;
        }
        if !matches!(
            self.tokens.get(self.pos),
            Some(Token::Control(ControlWord::MathProperties))
        ) {
            return Err(RtfError::MalformedDocument(
                "invalid RTF math-properties destination".to_string(),
            ));
        }
        self.pos += 1;
        let mut properties = crate::DocumentMathProperties::default();

        macro_rules! set_once {
            ($field:ident, $value:expr, $name:literal) => {{
                if properties.$field.is_some() {
                    return Err(RtfError::MalformedDocument(concat!(
                        "duplicate RTF math property ",
                        $name
                    )
                    .to_string()));
                }
                properties.$field = Some($value);
            }};
        }

        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    properties.validate()?;
                    return Ok(properties);
                },
                Some(Token::Text(text)) if self.decode_transport_text(text)?.trim().is_empty() => {
                    self.pos += 1;
                },
                Some(Token::Control(control)) => {
                    match *control {
                        ControlWord::MathBreakBinary(value) => set_once!(
                            binary_operator_break,
                            crate::MathBinaryOperatorBreak::from_rtf(value),
                            "mbrkBin"
                        ),
                        ControlWord::MathBreakBinarySubtraction(value) => set_once!(
                            binary_subtraction_break,
                            crate::MathBinarySubtractionBreak::from_rtf(value),
                            "mbrkBinSub"
                        ),
                        ControlWord::MathDefaultJustification(value) => set_once!(
                            default_justification,
                            crate::MathJustification::from_rtf(value),
                            "mdefJc"
                        ),
                        ControlWord::MathDisplayDefaults(value) => set_once!(
                            display_defaults,
                            crate::MathFlag::from_rtf(value),
                            "mdispDef"
                        ),
                        ControlWord::MathInterEquationSpacing(value) => set_once!(
                            inter_equation_spacing,
                            value,
                            "minterSp"
                        ),
                        ControlWord::MathIntegralLimitPlacement(value) => set_once!(
                            integral_limit_placement,
                            crate::MathLimitPlacement::from_rtf(value),
                            "mintLim"
                        ),
                        ControlWord::MathIntraEquationSpacing(value) => set_once!(
                            intra_equation_spacing,
                            value,
                            "mintraSp"
                        ),
                        ControlWord::MathLeftMargin(value) => {
                            set_once!(left_margin, value, "mlMargin")
                        },
                        ControlWord::MathFont(value) => {
                            let value = u32::try_from(value).map_err(|_| {
                                RtfError::MalformedDocument(
                                    "RTF math font index cannot be negative".to_string(),
                                )
                            })?;
                            set_once!(math_font, value, "mmathFont");
                        },
                        ControlWord::MathNaryLimitPlacement(value) => set_once!(
                            nary_limit_placement,
                            crate::MathLimitPlacement::from_rtf(value),
                            "mnaryLim"
                        ),
                        ControlWord::MathPostSpacing(value) => {
                            set_once!(post_spacing, value, "mpostSp")
                        },
                        ControlWord::MathPreSpacing(value) => {
                            set_once!(pre_spacing, value, "mpreSp")
                        },
                        ControlWord::MathRightMargin(value) => {
                            set_once!(right_margin, value, "mrMargin")
                        },
                        ControlWord::MathSmallFractions(value) => set_once!(
                            small_fractions,
                            crate::MathFlag::from_rtf(value),
                            "msmallFrac"
                        ),
                        ControlWord::MathWrapIndent(value) => {
                            set_once!(wrap_indent, value, "mwrapIndent")
                        },
                        ControlWord::MathWrapRight(value) => set_once!(
                            wrap_right,
                            crate::MathFlag::from_rtf(value),
                            "mwrapRight"
                        ),
                        _ => {
                            return Err(RtfError::MalformedDocument(
                                "RTF math-properties destination contains an unsupported control"
                                    .to_string(),
                            ));
                        },
                    }
                    self.pos += 1;
                },
                Some(Token::OpenBrace | Token::Binary(_)) | Some(Token::Text(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF math-properties destination contains active, nested, or binary data"
                            .to_string(),
                    ));
                },
                None => return Err(RtfError::UnexpectedEof),
            }
        }
    }

    fn parse_paragraph_group_table(&mut self) -> RtfResult<crate::ParagraphGroupPropertyTable> {
        if self.states.len() != 3
            || self.blocks.iter().any(|block| !block.text.trim().is_empty())
        {
            return Err(RtfError::MalformedDocument(
                "RTF pgptbl must occur at document scope before body text".to_string(),
            ));
        }
        self.pos += 1; // ignorable-destination marker
        if !matches!(self.tokens.get(self.pos), Some(Token::Control(ControlWord::ParagraphGroupTable))) {
            return Err(RtfError::MalformedDocument("invalid RTF pgptbl destination".to_string()));
        }
        self.pos += 1;
        let mut table = crate::ParagraphGroupPropertyTable::new();
        while self.pos < self.tokens.len() {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    table.validate()?;
                    return Ok(table);
                },
                Some(Token::OpenBrace)
                    if matches!(self.tokens.get(self.pos + 1), Some(Token::Control(ControlWord::ParagraphGroup))) =>
                {
                    let id = u32::try_from(table.entries().len() + 1).map_err(|_| {
                        RtfError::MalformedDocument("RTF paragraph-group ID overflow".to_string())
                    })?;
                    let entry = self.parse_paragraph_group_property(id)?;
                    table.push(entry)?;
                    continue;
                },
                Some(Token::Text(text)) if text.trim().is_empty() => {},
                Some(Token::OpenBrace) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF pgptbl cannot contain fields, objects, or unknown destinations".to_string(),
                    ));
                },
                Some(_) => {
                    return Err(RtfError::MalformedDocument(
                        "invalid content in RTF pgptbl destination".to_string(),
                    ));
                },
                None => return Err(RtfError::UnexpectedEof),
            }
            self.pos += 1;
        }
        Err(RtfError::UnexpectedEof)
    }

    fn parse_paragraph_group_property(
        &mut self,
        id: u32,
    ) -> RtfResult<crate::ParagraphGroupProperty> {
        self.pos += 2; // opening brace and pgp
        let mut parent_id = None;
        let mut nesting = None;
        let mut left = None;
        let mut right = None;
        let mut before = None;
        let mut after = None;
        let mut borders = crate::Borders::new();
        let mut current_border = None;
        let mut seen = std::collections::HashSet::new();
        while self.pos < self.tokens.len() {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    let entry = crate::ParagraphGroupProperty {
                        id,
                        parent_id: u32::try_from(parent_id.ok_or_else(|| RtfError::MalformedDocument("RTF pgp entry lacks ipgp".to_string()))?)
                            .map_err(|_| RtfError::MalformedDocument("invalid RTF ipgp reference".to_string()))?,
                        table_nesting_level: u8::try_from(nesting.ok_or_else(|| RtfError::MalformedDocument("RTF pgp entry lacks itap".to_string()))?)
                            .map_err(|_| RtfError::MalformedDocument("invalid RTF pgp itap value".to_string()))?,
                        left_indent: left.ok_or_else(|| RtfError::MalformedDocument("RTF pgp entry lacks li".to_string()))?,
                        right_indent: right.ok_or_else(|| RtfError::MalformedDocument("RTF pgp entry lacks ri".to_string()))?,
                        space_before: before.ok_or_else(|| RtfError::MalformedDocument("RTF pgp entry lacks sb".to_string()))?,
                        space_after: after.ok_or_else(|| RtfError::MalformedDocument("RTF pgp entry lacks sa".to_string()))?,
                        borders,
                    };
                    entry.validate()?;
                    return Ok(entry);
                },
                Some(Token::Control(control)) => match control {
                    ControlWord::ParagraphGroupParent(value) => {
                        if !seen.insert("ipgp") { return Err(RtfError::MalformedDocument("duplicate RTF pgp ipgp".to_string())); }
                        parent_id = Some(*value);
                    },
                    ControlWord::TableNestingLevel(value) => {
                        if !seen.insert("itap") { return Err(RtfError::MalformedDocument("duplicate RTF pgp itap".to_string())); }
                        nesting = Some(value.ok_or_else(||RtfError::MalformedDocument("RTF pgp itap requires a numeric parameter".to_string()))?);
                    },
                    ControlWord::LeftIndent(value) => {
                        if !seen.insert("li") { return Err(RtfError::MalformedDocument("duplicate RTF pgp li".to_string())); }
                        left = Some(*value);
                    },
                    ControlWord::RightIndent(value) => {
                        if !seen.insert("ri") { return Err(RtfError::MalformedDocument("duplicate RTF pgp ri".to_string())); }
                        right = Some(*value);
                    },
                    ControlWord::SpaceBefore(value) => {
                        if !seen.insert("sb") { return Err(RtfError::MalformedDocument("duplicate RTF pgp sb".to_string())); }
                        before = Some(*value);
                    },
                    ControlWord::SpaceAfter(value) => {
                        if !seen.insert("sa") { return Err(RtfError::MalformedDocument("duplicate RTF pgp sa".to_string())); }
                        after = Some(*value);
                    },
                    ControlWord::BorderTop => {
                        if !seen.insert("brdrt") { return Err(RtfError::MalformedDocument("duplicate RTF pgp top border".to_string())); }
                        current_border = Some(0u8);
                    },
                    ControlWord::BorderBottom => {
                        if !seen.insert("brdrb") { return Err(RtfError::MalformedDocument("duplicate RTF pgp bottom border".to_string())); }
                        current_border = Some(1u8);
                    },
                    ControlWord::BorderLeft => {
                        if !seen.insert("brdrl") { return Err(RtfError::MalformedDocument("duplicate RTF pgp left border".to_string())); }
                        current_border = Some(2u8);
                    },
                    ControlWord::BorderRight => {
                        if !seen.insert("brdrr") { return Err(RtfError::MalformedDocument("duplicate RTF pgp right border".to_string())); }
                        current_border = Some(3u8);
                    },
                    ControlWord::BorderNone
                    | ControlWord::BorderSingle
                    | ControlWord::BorderDotted
                    | ControlWord::BorderDashed
                    | ControlWord::BorderDouble
                    | ControlWord::BorderTriple
                    | ControlWord::BorderWave => {
                        let border = match current_border {
                            Some(0) => &mut borders.top,
                            Some(1) => &mut borders.bottom,
                            Some(2) => &mut borders.left,
                            Some(3) => &mut borders.right,
                            _ => return Err(RtfError::MalformedDocument("RTF pgp border style has no side".to_string())),
                        };
                        if border.style != crate::BorderStyle::None
                            && !matches!(control, ControlWord::BorderNone)
                        {
                            return Err(RtfError::MalformedDocument("duplicate RTF pgp border style".to_string()));
                        }
                        border.style = match control {
                            ControlWord::BorderNone => crate::BorderStyle::None,
                            ControlWord::BorderSingle => crate::BorderStyle::Single,
                            ControlWord::BorderDotted => crate::BorderStyle::Dotted,
                            ControlWord::BorderDashed => crate::BorderStyle::Dashed,
                            ControlWord::BorderDouble => crate::BorderStyle::Double,
                            ControlWord::BorderTriple => crate::BorderStyle::Triple,
                            _ => crate::BorderStyle::Wavy,
                        };
                    },
                    ControlWord::BorderWidth(value) => {
                        let border = match current_border { Some(0) => &mut borders.top, Some(1) => &mut borders.bottom, Some(2) => &mut borders.left, Some(3) => &mut borders.right, _ => return Err(RtfError::MalformedDocument("RTF pgp border width has no side".to_string())) };
                        border.width = value.ok_or_else(|| RtfError::MalformedDocument("RTF pgp brdrw requires a numeric parameter".to_string()))?;
                    },
                    ControlWord::BorderColor(value) => {
                        let border = match current_border { Some(0) => &mut borders.top, Some(1) => &mut borders.bottom, Some(2) => &mut borders.left, Some(3) => &mut borders.right, _ => return Err(RtfError::MalformedDocument("RTF pgp border color has no side".to_string())) };
                        border.color_ref = u16::try_from(value.ok_or_else(|| RtfError::MalformedDocument("RTF pgp brdrcf requires a numeric parameter".to_string()))?).map_err(|_| RtfError::MalformedDocument("invalid RTF pgp border color".to_string()))?;
                    },
                    ControlWord::BorderSpace(value) => {
                        let border = match current_border { Some(0) => &mut borders.top, Some(1) => &mut borders.bottom, Some(2) => &mut borders.left, Some(3) => &mut borders.right, _ => return Err(RtfError::MalformedDocument("RTF pgp border space has no side".to_string())) };
                        border.space = value.ok_or_else(|| RtfError::MalformedDocument("RTF pgp brsp requires a numeric parameter".to_string()))?;
                    },
                    _ => return Err(RtfError::MalformedDocument("unsupported control in RTF pgp entry".to_string())),
                },
                Some(Token::Text(text)) if text.trim().is_empty() => {},
                Some(Token::OpenBrace) => return Err(RtfError::MalformedDocument("RTF pgp entry cannot contain nested destinations".to_string())),
                Some(_) => return Err(RtfError::MalformedDocument("invalid content in RTF pgp entry".to_string())),
                None => return Err(RtfError::UnexpectedEof),
            }
            self.pos += 1;
        }
        Err(RtfError::UnexpectedEof)
    }

    fn parse_legacy_section_numbering_level(&mut self) -> RtfResult<()> {
        if self.states.len() != 3
            || self.blocks.iter().any(|block| !block.text.trim().is_empty())
        {
            return Err(RtfError::MalformedDocument(
                "RTF pnseclvl destinations must occur at document scope before body text"
                    .to_string(),
            ));
        }
        self.pos += 1; // ignorable-destination marker
        let level_index = match self.tokens.get(self.pos) {
            Some(Token::Control(ControlWord::LegacySectionNumberingLevel(value))) => {
                u8::try_from(*value).map_err(|_| {
                    RtfError::MalformedDocument(
                        "RTF pnseclvl index must be between 1 and 9".to_string(),
                    )
                })?
            },
            _ => {
                return Err(RtfError::MalformedDocument(
                    "invalid RTF pnseclvl destination".to_string(),
                ));
            },
        };
        self.pos += 1;
        let mut format = None;
        let mut start_at = None;
        let mut indent = None;
        let mut space = None;
        let mut hanging = false;
        let mut previous = false;
        let mut alignment = None;
        let mut font_ref = None;
        let mut text_before = String::new();
        let mut text_after = String::new();
        let mut seen = std::collections::HashSet::new();
        while self.pos < self.tokens.len() {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    let format = format.ok_or_else(|| {
                        RtfError::MalformedDocument(
                            "RTF pnseclvl destination has no numbering format".to_string(),
                        )
                    })?;
                    let mut level = crate::LegacySectionNumberingLevel::new(level_index, format);
                    level.start_at = start_at;
                    level.indent = indent;
                    level.space = space;
                    level.hanging = hanging;
                    level.previous = previous;
                    level.alignment = alignment;
                    level.font_ref = font_ref;
                    level.text_before = Cow::Owned(text_before);
                    level.text_after = Cow::Owned(text_after);
                    return self.legacy_section_numbering.add(level);
                },
                Some(Token::OpenBrace)
                    if matches!(
                        self.tokens.get(self.pos + 1),
                        Some(Token::Control(ControlWord::LegacyNumberingTextBefore))
                    ) =>
                {
                    if !seen.insert("text-before") {
                        return Err(RtfError::MalformedDocument(
                            "duplicate RTF pntxtb destination".to_string(),
                        ));
                    }
                    text_before = self.parse_legacy_numbering_text(
                        ControlWord::LegacyNumberingTextBefore,
                    )?;
                    continue;
                },
                Some(Token::OpenBrace)
                    if matches!(
                        self.tokens.get(self.pos + 1),
                        Some(Token::Control(ControlWord::LegacyNumberingTextAfter))
                    ) =>
                {
                    if !seen.insert("text-after") {
                        return Err(RtfError::MalformedDocument(
                            "duplicate RTF pntxta destination".to_string(),
                        ));
                    }
                    text_after = self.parse_legacy_numbering_text(
                        ControlWord::LegacyNumberingTextAfter,
                    )?;
                    continue;
                },
                Some(Token::OpenBrace) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF pnseclvl cannot contain nested fields, objects, or destinations"
                            .to_string(),
                    ));
                },
                Some(Token::Control(control)) => {
                    let (key, new_format) = match control {
                        ControlWord::LegacyNumberingDecimal => ("format", Some(crate::LegacyNumberingFormat::Decimal)),
                        ControlWord::LegacyNumberingUpperRoman => ("format", Some(crate::LegacyNumberingFormat::UpperRoman)),
                        ControlWord::LegacyNumberingLowerRoman => ("format", Some(crate::LegacyNumberingFormat::LowerRoman)),
                        ControlWord::LegacyNumberingUpperLetter => ("format", Some(crate::LegacyNumberingFormat::UpperLetter)),
                        ControlWord::LegacyNumberingLowerLetter => ("format", Some(crate::LegacyNumberingFormat::LowerLetter)),
                        ControlWord::LegacyNumberingStart(value) => {
                            if !seen.insert("start") { return Err(RtfError::MalformedDocument("duplicate RTF pnstart".to_string())); }
                            start_at = Some(*value);
                            self.pos += 1;
                            continue;
                        },
                        ControlWord::LegacyNumberingIndent(value) => {
                            if !seen.insert("indent") { return Err(RtfError::MalformedDocument("duplicate RTF pnindent".to_string())); }
                            indent = Some(*value);
                            self.pos += 1;
                            continue;
                        },
                        ControlWord::LegacyNumberingSpace(value) => {
                            if !seen.insert("space") { return Err(RtfError::MalformedDocument("duplicate RTF pnsp".to_string())); }
                            space = Some(*value);
                            self.pos += 1;
                            continue;
                        },
                        ControlWord::LegacyNumberingHanging => {
                            if !seen.insert("hanging") { return Err(RtfError::MalformedDocument("duplicate RTF pnhang".to_string())); }
                            hanging = true;
                            self.pos += 1;
                            continue;
                        },
                        ControlWord::LegacyNumberingPrevious => {
                            if !seen.insert("previous") { return Err(RtfError::MalformedDocument("duplicate RTF pnprev".to_string())); }
                            previous = true;
                            self.pos += 1;
                            continue;
                        },
                        ControlWord::LegacyNumberingAlignLeft => ("alignment", None),
                        ControlWord::LegacyNumberingAlignCenter => ("alignment-center", None),
                        ControlWord::LegacyNumberingAlignRight => ("alignment-right", None),
                        ControlWord::LegacyNumberingFont(value) => {
                            if !seen.insert("font") { return Err(RtfError::MalformedDocument("duplicate RTF pnf".to_string())); }
                            font_ref = Some(u16::try_from(*value).map_err(|_| RtfError::MalformedDocument("invalid RTF pnf reference".to_string()))?);
                            self.pos += 1;
                            continue;
                        },
                        _ => {
                            return Err(RtfError::MalformedDocument(
                                "unsupported control in RTF pnseclvl destination".to_string(),
                            ));
                        },
                    };
                    if key == "format" {
                        if !seen.insert(key) {
                            return Err(RtfError::MalformedDocument(
                                "duplicate RTF pnseclvl numbering format".to_string(),
                            ));
                        }
                        format = new_format;
                    } else {
                        if !seen.insert("alignment") {
                            return Err(RtfError::MalformedDocument(
                                "duplicate RTF pnseclvl alignment".to_string(),
                            ));
                        }
                        alignment = Some(match control {
                            ControlWord::LegacyNumberingAlignCenter => crate::LegacyNumberingAlignment::Center,
                            ControlWord::LegacyNumberingAlignRight => crate::LegacyNumberingAlignment::Right,
                            _ => crate::LegacyNumberingAlignment::Left,
                        });
                    }
                },
                Some(Token::Text(text)) if text.trim().is_empty() => {},
                Some(_) => {
                    return Err(RtfError::MalformedDocument(
                        "invalid content in RTF pnseclvl destination".to_string(),
                    ));
                },
                None => return Err(RtfError::UnexpectedEof),
            }
            self.pos += 1;
        }
        Err(RtfError::UnexpectedEof)
    }

    fn parse_legacy_numbering_text(&mut self, expected: ControlWord<'a>) -> RtfResult<String> {
        if !matches!(self.tokens.get(self.pos), Some(Token::OpenBrace))
            || self.tokens.get(self.pos + 1) != Some(&Token::Control(expected))
        {
            return Err(RtfError::MalformedDocument(
                "invalid RTF legacy-numbering text destination".to_string(),
            ));
        }
        self.pos += 2;
        let mut value = String::new();
        let mut unicode_skip = self.current_state()?.unicode_skip.max(0);
        while self.pos < self.tokens.len() {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    return Ok(value);
                },
                Some(Token::OpenBrace) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF legacy-numbering text cannot contain nested destinations".to_string(),
                    ));
                },
                Some(Token::Text(text)) => value.push_str(&self.decode_transport_text(text)?),
                Some(Token::Control(ControlWord::Unicode(first))) => {
                    value.push_str(&self.parse_style_unicode(*first, unicode_skip)?);
                    continue;
                },
                Some(Token::Control(ControlWord::UnicodeSkip(value))) => {
                    unicode_skip = (*value).max(0);
                },
                Some(Token::Control(control)) if control_symbol_text(control).is_some() => {
                    value.push_str(control_symbol_text(control).unwrap_or_default());
                },
                // Character-type selectors affect only which run font decodes
                // the following text. `dbch` is emitted by Word in pnseclvl
                // punctuation destinations and carries no textual payload.
                Some(Token::Control(ControlWord::Unknown("dbch", None))) => {},
                Some(Token::Control(_)) | Some(Token::Binary(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF legacy-numbering text contains a non-text control".to_string(),
                    ));
                },
                None => return Err(RtfError::UnexpectedEof),
            }
            self.pos += 1;
            if value.len() > crate::legacy_numbering::MAX_LEGACY_NUMBERING_TEXT_BYTES {
                return Err(RtfError::MalformedDocument(
                    "RTF pnseclvl text exceeds the safety limit".to_string(),
                ));
            }
        }
        Err(RtfError::UnexpectedEof)
    }

    fn parse_latent_styles(&mut self) -> RtfResult<crate::LatentStyles<'a>> {
        self.pos += 1; // ignorable-destination marker
        if !matches!(
            self.tokens.get(self.pos),
            Some(Token::Control(ControlWord::LatentStyles))
        ) {
            return Err(RtfError::MalformedDocument(
                "invalid RTF latentstyles destination".to_string(),
            ));
        }
        self.pos += 1;
        let mut max_style_index = None;
        let mut locked_default = None;
        let mut semi_hidden_default = None;
        let mut unhide_when_used_default = None;
        let mut quick_format_default = None;
        let mut priority_default = None;
        let mut exceptions = None;
        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    let styles = crate::LatentStyles {
                        max_style_index: max_style_index.ok_or_else(|| {
                            RtfError::MalformedDocument(
                                "RTF latentstyles is missing lsdstimax".to_string(),
                            )
                        })?,
                        locked_default,
                        semi_hidden_default,
                        unhide_when_used_default,
                        quick_format_default,
                        priority_default,
                        exceptions: exceptions.unwrap_or_default(),
                    };
                    styles.validate()?;
                    return Ok(styles);
                },
                Some(Token::OpenBrace) => {
                    if exceptions.is_some()
                        || !matches!(
                            self.tokens.get(self.pos + 1),
                            Some(Token::Control(ControlWord::LatentStyleExceptions))
                        )
                    {
                        return Err(RtfError::MalformedDocument(
                            "RTF latentstyles contains a duplicate or active nested destination"
                                .to_string(),
                        ));
                    }
                    exceptions = Some(self.parse_latent_style_exceptions()?);
                },
                Some(Token::Control(control)) => {
                    macro_rules! set_once {
                        ($slot:expr, $value:expr, $name:literal) => {{
                            if $slot.is_some() {
                                return Err(RtfError::MalformedDocument(concat!(
                                    "duplicate RTF latent-style ",
                                    $name
                                )
                                .to_string()));
                            }
                            $slot = Some($value);
                        }};
                    }
                    match control {
                        ControlWord::LatentStyleMax(value) => {
                            let value = u32::try_from(*value).map_err(|_| {
                                RtfError::MalformedDocument(
                                    "RTF lsdstimax cannot be negative".to_string(),
                                )
                            })?;
                            if value > crate::latent_style::MAX_LATENT_STYLE_INDEX {
                                return Err(RtfError::MalformedDocument(
                                    "RTF lsdstimax exceeds 65535".to_string(),
                                ));
                            }
                            set_once!(max_style_index, value, "lsdstimax");
                        },
                        ControlWord::LatentStyleLockedDefault(value) => set_once!(
                            locked_default,
                            Self::parse_latent_style_bool(*value)?,
                            "lsdlockeddef"
                        ),
                        ControlWord::LatentStyleSemiHiddenDefault(value) => set_once!(
                            semi_hidden_default,
                            Self::parse_latent_style_bool(*value)?,
                            "lsdsemihiddendef"
                        ),
                        ControlWord::LatentStyleUnhideUsedDefault(value) => set_once!(
                            unhide_when_used_default,
                            Self::parse_latent_style_bool(*value)?,
                            "lsdunhideuseddef"
                        ),
                        ControlWord::LatentStyleQuickFormatDefault(value) => set_once!(
                            quick_format_default,
                            Self::parse_latent_style_bool(*value)?,
                            "lsdqformatdef"
                        ),
                        ControlWord::LatentStylePriorityDefault(value) => set_once!(
                            priority_default,
                            Self::parse_latent_style_priority(*value)?,
                            "lsdprioritydef"
                        ),
                        _ => {
                            return Err(RtfError::MalformedDocument(
                                "RTF latentstyles contains an unsupported control".to_string(),
                            ));
                        },
                    }
                    self.pos += 1;
                },
                Some(Token::Text(text)) if self.decode_transport_text(text)?.trim().is_empty() => {
                    self.pos += 1;
                },
                Some(Token::Binary(_)) | Some(Token::Text(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF latentstyles contains orphan text or binary data".to_string(),
                    ));
                },
                None => return Err(RtfError::UnexpectedEof),
            }
        }
    }

    fn parse_latent_style_exceptions(
        &mut self,
    ) -> RtfResult<Vec<crate::LatentStyleException<'a>>> {
        self.expect_token(Token::OpenBrace)?;
        if !matches!(
            self.tokens.get(self.pos),
            Some(Token::Control(ControlWord::LatentStyleExceptions))
        ) {
            return Err(RtfError::MalformedDocument(
                "invalid RTF lsdlockedexcept destination".to_string(),
            ));
        }
        self.pos += 1;
        let mut entries = Vec::new();
        let mut builder = LatentStyleExceptionBuilder::default();
        let mut name = String::new();
        let mut unicode_skip = self.current_state()?.unicode_skip.max(0);
        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    if !name.trim().is_empty()
                        || builder.locked.is_some()
                        || builder.semi_hidden.is_some()
                        || builder.unhide_when_used.is_some()
                        || builder.quick_format.is_some()
                        || builder.priority.is_some()
                    {
                        return Err(RtfError::MalformedDocument(
                            "RTF latent-style exception is missing its terminating semicolon"
                                .to_string(),
                        ));
                    }
                    return Ok(entries);
                },
                Some(Token::Control(control)) => {
                    if matches!(
                        control,
                        ControlWord::LatentStyleLocked(_)
                            | ControlWord::LatentStyleSemiHidden(_)
                            | ControlWord::LatentStyleUnhideUsed(_)
                            | ControlWord::LatentStyleQuickFormat(_)
                            | ControlWord::LatentStylePriority(_)
                    ) && !name.trim().is_empty()
                    {
                        return Err(RtfError::MalformedDocument(
                            "RTF latent-style properties must precede the style name".to_string(),
                        ));
                    }
                    macro_rules! set_once {
                        ($slot:expr, $value:expr, $name:literal) => {{
                            if $slot.is_some() {
                                return Err(RtfError::MalformedDocument(concat!(
                                    "duplicate RTF latent-style exception ",
                                    $name
                                )
                                .to_string()));
                            }
                            $slot = Some($value);
                        }};
                    }
                    match control {
                        ControlWord::LatentStyleLocked(value) => set_once!(
                            builder.locked,
                            Self::parse_latent_style_bool(*value)?,
                            "lsdlocked"
                        ),
                        ControlWord::LatentStyleSemiHidden(value) => set_once!(
                            builder.semi_hidden,
                            Self::parse_latent_style_bool(*value)?,
                            "lsdsemihidden"
                        ),
                        ControlWord::LatentStyleUnhideUsed(value) => set_once!(
                            builder.unhide_when_used,
                            Self::parse_latent_style_bool(*value)?,
                            "lsdunhideused"
                        ),
                        ControlWord::LatentStyleQuickFormat(value) => set_once!(
                            builder.quick_format,
                            Self::parse_latent_style_bool(*value)?,
                            "lsdqformat"
                        ),
                        ControlWord::LatentStylePriority(value) => set_once!(
                            builder.priority,
                            Self::parse_latent_style_priority(*value)?,
                            "lsdpriority"
                        ),
                        ControlWord::Unicode(first) => {
                            name.push_str(&self.parse_style_unicode(*first, unicode_skip)?);
                            while let Some(separator) = name.find(';') {
                                let remainder = name.split_off(separator + 1);
                                let entry_name = name[..separator].trim();
                                if entry_name.is_empty() {
                                    return Err(RtfError::MalformedDocument(
                                        "RTF latent-style exception name cannot be empty"
                                            .to_string(),
                                    ));
                                }
                                if entries.len()
                                    >= crate::latent_style::MAX_LATENT_STYLE_EXCEPTIONS
                                {
                                    return Err(RtfError::MalformedDocument(
                                        "RTF latent-style exception count exceeds the safety limit"
                                            .to_string(),
                                    ));
                                }
                                entries.push(crate::LatentStyleException {
                                    name: Cow::Borrowed(self.arena.alloc_str(entry_name)),
                                    locked: builder.locked.take(),
                                    semi_hidden: builder.semi_hidden.take(),
                                    unhide_when_used: builder.unhide_when_used.take(),
                                    quick_format: builder.quick_format.take(),
                                    priority: builder.priority.take(),
                                });
                                name = remainder;
                            }
                            continue;
                        },
                        ControlWord::UnicodeSkip(value) => {
                            unicode_skip = (*value).max(0);
                            self.pos += 1;
                            continue;
                        },
                        control if control_symbol_text(control).is_some() => {
                            name.push_str(control_symbol_text(control).unwrap_or_default());
                            self.pos += 1;
                            continue;
                        },
                        _ => {
                            return Err(RtfError::MalformedDocument(
                                "RTF latent-style exception contains an unsupported control"
                                    .to_string(),
                            ));
                        },
                    }
                    self.pos += 1;
                },
                Some(Token::Text(text)) => {
                    name.push_str(&self.decode_transport_text(text)?);
                    self.pos += 1;
                    while let Some(separator) = name.find(';') {
                        let remainder = name.split_off(separator + 1);
                        let entry_name = name[..separator].trim();
                        if entry_name.is_empty() {
                            return Err(RtfError::MalformedDocument(
                                "RTF latent-style exception name cannot be empty".to_string(),
                            ));
                        }
                        if entries.len()
                            >= crate::latent_style::MAX_LATENT_STYLE_EXCEPTIONS
                        {
                            return Err(RtfError::MalformedDocument(
                                "RTF latent-style exception count exceeds the safety limit"
                                    .to_string(),
                            ));
                        }
                        entries.push(crate::LatentStyleException {
                            name: Cow::Borrowed(self.arena.alloc_str(entry_name)),
                            locked: builder.locked.take(),
                            semi_hidden: builder.semi_hidden.take(),
                            unhide_when_used: builder.unhide_when_used.take(),
                            quick_format: builder.quick_format.take(),
                            priority: builder.priority.take(),
                        });
                        name = remainder;
                    }
                },
                Some(Token::OpenBrace | Token::Binary(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF latent-style exceptions cannot contain nesting or binary data"
                            .to_string(),
                    ));
                },
                None => return Err(RtfError::UnexpectedEof),
            }
            if name.len() > crate::latent_style::MAX_LATENT_STYLE_NAME_BYTES {
                return Err(RtfError::MalformedDocument(
                    "RTF latent-style exception name exceeds the safety limit".to_string(),
                ));
            }
        }
    }

    fn parse_latent_style_bool(value: i32) -> RtfResult<bool> {
        match value {
            0 => Ok(false),
            1 => Ok(true),
            _ => Err(RtfError::MalformedDocument(
                "RTF latent-style Boolean values must be 0 or 1".to_string(),
            )),
        }
    }

    fn parse_latent_style_priority(value: i32) -> RtfResult<u8> {
        u8::try_from(value)
            .ok()
            .filter(|priority| *priority <= 99)
            .ok_or_else(|| {
                RtfError::MalformedDocument(
                    "RTF latent-style priority must be in 0..=99".to_string(),
                )
            })
    }

    fn parse_theme_hex_destination(
        &mut self,
        expected: ControlWord<'a>,
        limit: usize,
    ) -> RtfResult<Vec<u8>> {
        self.pos += 1; // ignorable-destination marker
        let matches_expected = matches!(
            (&expected, self.tokens.get(self.pos)),
            (ControlWord::ThemeData, Some(Token::Control(ControlWord::ThemeData)))
                | (
                    ControlWord::ColorSchemeMapping,
                    Some(Token::Control(ControlWord::ColorSchemeMapping))
                )
        );
        if !matches_expected {
            return Err(RtfError::MalformedDocument(
                "invalid RTF theme destination".to_string(),
            ));
        }
        self.pos += 1;
        let mut data = Vec::new();
        let mut high = None;
        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    if high.is_some() {
                        return Err(RtfError::MalformedDocument(
                            "RTF theme payload has an odd hexadecimal digit count".to_string(),
                        ));
                    }
                    if data.is_empty() {
                        return Err(RtfError::MalformedDocument(
                            "RTF theme payload cannot be empty".to_string(),
                        ));
                    }
                    return Ok(data);
                },
                Some(Token::Text(text)) => {
                    for byte in text.as_bytes() {
                        if byte.is_ascii_whitespace() {
                            continue;
                        }
                        let nibble = match byte {
                            b'0'..=b'9' => byte - b'0',
                            b'a'..=b'f' => byte - b'a' + 10,
                            b'A'..=b'F' => byte - b'A' + 10,
                            _ => {
                                return Err(RtfError::MalformedDocument(
                                    "RTF theme payload contains a non-hexadecimal character"
                                        .to_string(),
                                ));
                            },
                        };
                        if let Some(first) = high.take() {
                            data.push(first << 4 | nibble);
                        } else {
                            high = Some(nibble);
                        }
                    }
                    self.pos += 1;
                },
                Some(Token::OpenBrace | Token::Binary(_)) | Some(Token::Control(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF theme payload cannot contain controls, nesting, or binary data"
                            .to_string(),
                    ));
                },
                None => return Err(RtfError::UnexpectedEof),
            }
            if data.len() > limit {
                return Err(RtfError::MalformedDocument(
                    "RTF theme payload exceeds the safety limit".to_string(),
                ));
            }
        }
    }

    fn parse_xml_namespace_table(&mut self) -> RtfResult<()> {
        if self.saw_xml_namespace_table {
            return Err(RtfError::MalformedDocument(
                "RTF contains multiple XML namespace tables".to_string(),
            ));
        }
        self.saw_xml_namespace_table = true;
        self.pos += 1; // xmlnstbl
        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    return Ok(());
                },
                Some(Token::OpenBrace) => self.parse_xml_namespace_entry()?,
                Some(Token::Text(text)) if self.decode_transport_text(text)?.trim().is_empty() => {
                    self.pos += 1;
                },
                Some(Token::Binary(_))
                | Some(Token::Control(_))
                | Some(Token::Text(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF XML namespace table contains ungrouped, active, or binary data"
                            .to_string(),
                    ));
                },
                None => return Err(RtfError::UnexpectedEof),
            }
        }
    }

    fn parse_xml_namespace_entry(&mut self) -> RtfResult<()> {
        if self.xml_namespaces.len() >= crate::xml_namespace::MAX_XML_NAMESPACES {
            return Err(RtfError::MalformedDocument(
                "RTF XML namespace count exceeds the safety limit".to_string(),
            ));
        }
        self.expect_token(Token::OpenBrace)?;
        let id = match self.tokens.get(self.pos) {
            Some(Token::Control(ControlWord::XmlNamespace(value))) => {
                let value = u32::try_from(*value).map_err(|_| {
                    RtfError::MalformedDocument(
                        "RTF XML namespace ID must be a positive signed integer".to_string(),
                    )
                })?;
                if value == 0 {
                    return Err(RtfError::MalformedDocument(
                        "RTF XML namespace ID must be a positive signed integer".to_string(),
                    ));
                }
                value
            },
            _ => {
                return Err(RtfError::MalformedDocument(
                    "RTF XML namespace entry is missing xmlnsN".to_string(),
                ));
            },
        };
        if self.xml_namespaces.iter().any(|entry| entry.id == id) {
            return Err(RtfError::MalformedDocument(
                "RTF XML namespace IDs must be unique".to_string(),
            ));
        }
        self.pos += 1;
        let mut namespace = String::new();
        let mut unicode_skip = self.current_state()?.unicode_skip.max(0);
        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    let namespace = namespace.trim();
                    let entry = crate::XmlNamespace::new(
                        id,
                        Cow::Borrowed(self.arena.alloc_str(namespace)),
                    )?;
                    self.xml_namespace_text_bytes = self
                        .xml_namespace_text_bytes
                        .checked_add(entry.namespace.len())
                        .ok_or_else(|| {
                            RtfError::MalformedDocument(
                                "RTF XML namespace aggregate size overflow".to_string(),
                            )
                        })?;
                    if self.xml_namespace_text_bytes
                        > crate::xml_namespace::MAX_XML_NAMESPACE_TOTAL_BYTES
                    {
                        return Err(RtfError::MalformedDocument(
                            "RTF XML namespace aggregate text exceeds the safety limit"
                                .to_string(),
                        ));
                    }
                    self.xml_namespaces.push(entry);
                    return Ok(());
                },
                Some(Token::Text(text)) => {
                    namespace.extend(
                        self.decode_transport_text(text)?
                            .chars()
                            .filter(|character| !matches!(character, '\r' | '\n')),
                    );
                    self.pos += 1;
                },
                Some(Token::Control(ControlWord::Unicode(first))) => {
                    namespace.push_str(&self.parse_style_unicode(*first, unicode_skip)?);
                },
                Some(Token::Control(ControlWord::UnicodeSkip(value))) => {
                    unicode_skip = (*value).max(0);
                    self.pos += 1;
                },
                Some(Token::Control(control)) if control_symbol_text(control).is_some() => {
                    namespace.push_str(control_symbol_text(control).unwrap_or_default());
                    self.pos += 1;
                },
                Some(Token::OpenBrace | Token::Binary(_)) | Some(Token::Control(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF XML namespace entry contains active, nested, or binary data"
                            .to_string(),
                    ));
                },
                None => return Err(RtfError::UnexpectedEof),
            }
            if namespace.len() > crate::xml_namespace::MAX_XML_NAMESPACE_BYTES {
                return Err(RtfError::MalformedDocument(
                    "RTF XML namespace identifier exceeds the safety limit".to_string(),
                ));
            }
        }
    }

    fn parse_generator_destination(&mut self) -> RtfResult<crate::DocumentGenerator<'a>> {
        self.pos += 1; // ignorable-destination marker
        if !matches!(
            self.tokens.get(self.pos),
            Some(Token::Control(ControlWord::Generator))
        ) {
            return Err(RtfError::MalformedDocument(
                "invalid RTF generator destination".to_string(),
            ));
        }
        self.pos += 1;
        let mut value = String::new();
        let mut unicode_skip = self.current_state()?.unicode_skip.max(0);
        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    let value = value.trim();
                    let value = value.strip_suffix(';').unwrap_or(value).trim_end();
                    let value = self.arena.alloc_str(value);
                    return crate::DocumentGenerator::new(Cow::Borrowed(value));
                },
                Some(Token::Text(text)) => {
                    value.push_str(&self.decode_transport_text(text)?);
                    self.pos += 1;
                },
                Some(Token::Control(ControlWord::Unicode(first))) => {
                    value.push_str(&self.parse_style_unicode(*first, unicode_skip)?);
                },
                Some(Token::Control(ControlWord::UnicodeSkip(count))) => {
                    unicode_skip = (*count).max(0);
                    self.pos += 1;
                },
                Some(Token::Control(control)) if control_symbol_text(control).is_some() => {
                    value.push_str(control_symbol_text(control).unwrap_or_default());
                    self.pos += 1;
                },
                Some(Token::OpenBrace | Token::Binary(_)) | Some(Token::Control(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF generator destination contains active, nested, or binary data"
                            .to_string(),
                    ));
                },
                None => return Err(RtfError::UnexpectedEof),
            }
            if value.len() > crate::generator::MAX_GENERATOR_BYTES {
                return Err(RtfError::MalformedDocument(
                    "RTF generator value exceeds the safety limit".to_string(),
                ));
            }
        }
    }

    fn parse_revision_author_group(&mut self) -> RtfResult<String> {
        self.pos += 1; // opening brace
        let mut author = String::new();
        let mut unicode_skip = self.current_state()?.unicode_skip.max(0);
        while self.pos < self.tokens.len() {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    return Ok(author
                        .trim_end_matches(['\r', '\n', ' '])
                        .strip_suffix(';')
                        .unwrap_or(author.trim_end_matches(['\r', '\n', ' ']))
                        .trim()
                        .to_string());
                },
                Some(Token::OpenBrace) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF revision author contains a nested destination".to_string(),
                    ));
                },
                Some(Token::Text(text)) => author.push_str(&self.decode_transport_text(text)?),
                Some(Token::Control(ControlWord::Unicode(first))) => {
                    let decoded = self.parse_style_unicode(*first, unicode_skip)?;
                    author.push_str(&decoded);
                    continue;
                },
                Some(Token::Control(ControlWord::UnicodeSkip(value))) => {
                    unicode_skip = (*value).max(0);
                },
                Some(Token::Control(control)) if control_symbol_text(control).is_some() => {
                    author.push_str(control_symbol_text(control).unwrap_or_default());
                },
                Some(Token::Control(_)) | Some(Token::Binary(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF revision author contains a non-text control or binary data"
                            .to_string(),
                    ));
                },
                None => break,
            }
            self.pos += 1;
            if author.len() > MAX_REVISION_AUTHOR_BYTES {
                return Err(RtfError::MalformedDocument(
                    "RTF revision author exceeds the safety limit".to_string(),
                ));
            }
        }
        Err(RtfError::UnexpectedEof)
    }

    fn push_direct_revision_authors(&mut self, text: &mut String) -> RtfResult<()> {
        while let Some(separator) = text.find(';') {
            let remainder = text.split_off(separator + 1);
            let author = text[..separator].trim().to_string();
            self.push_revision_author(author)?;
            *text = remainder;
        }
        Ok(())
    }

    fn push_revision_author(&mut self, author: String) -> RtfResult<()> {
        if self.revision_authors.len() >= MAX_REVISION_AUTHORS {
            return Err(RtfError::MalformedDocument(
                "RTF revision author count exceeds the safety limit".to_string(),
            ));
        }
        self.revision_author_text_bytes = self
            .revision_author_text_bytes
            .checked_add(author.len())
            .ok_or_else(|| {
                RtfError::MalformedDocument(
                    "RTF aggregate revision-author size overflow".to_string(),
                )
            })?;
        if self.revision_author_text_bytes
            > super::annotation::MAX_REVISION_AUTHOR_TEXT_TOTAL_BYTES
        {
            return Err(RtfError::MalformedDocument(
                "RTF aggregate revision-author text exceeds the safety limit".to_string(),
            ));
        }
        let author = super::annotation::RevisionAuthor::new(Cow::Borrowed(
            self.arena.alloc_str(&author),
        ))?;
        author.validate()?;
        self.revision_authors.push(author);
        Ok(())
    }

    fn parse_list_table(&mut self) -> RtfResult<()> {
        if self.saw_list_table {
            return Err(RtfError::MalformedDocument("RTF document contains multiple list tables".to_string()));
        }
        if self.states.len() != 3 || self.blocks.iter().any(|block| !block.text.trim().is_empty()) {
            return Err(RtfError::MalformedDocument("RTF list table must occur in the root header".to_string()));
        }
        self.saw_list_table = true;
        self.pos += 1; // `listtable`
        while self.pos < self.tokens.len() {
            match self.tokens.get(self.pos) {
                Some(Token::OpenBrace)
                    if matches!(
                        self.tokens.get(self.pos + 1),
                        Some(Token::Control(ControlWord::List))
                    ) =>
                {
                    self.parse_list_definition()?;
                },
                Some(Token::OpenBrace)
                    if matches!(
                        self.tokens.get(self.pos..self.pos + 3),
                        Some([
                            Token::OpenBrace,
                            Token::Control(ControlWord::IgnorableDestination),
                            Token::Control(ControlWord::ListPicture),
                        ])
                    ) =>
                {
                    self.list_table.picture_bullet_count = self
                        .list_table
                        .picture_bullet_count
                        .checked_add(1)
                        .ok_or_else(|| RtfError::MalformedDocument("RTF list-picture count overflow".to_string()))?;
                    self.skip_group()?;
                },
                Some(Token::OpenBrace) => self.skip_group()?,
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    self.list_table.validate()?;
                    return Ok(());
                },
                Some(_) => self.pos += 1,
                None => break,
            }
        }
        Err(RtfError::UnexpectedEof)
    }

    fn parse_list_definition(&mut self) -> RtfResult<()> {
        self.pos += 2; // opening brace and `list`
        let mut list = super::list::List::new(0);
        list.simple = false;
        let mut has_id = false;
        let mut has_template_id = false;
        let mut closed = false;
        while self.pos < self.tokens.len() {
            match self.tokens.get(self.pos) {
                Some(Token::OpenBrace)
                    if matches!(
                        self.tokens.get(self.pos + 1),
                        Some(Token::Control(ControlWord::ListLevel))
                    ) =>
                {
                    if list.levels.len() >= MAX_LIST_LEVELS {
                        return Err(RtfError::MalformedDocument(
                            "RTF list exceeds the nine-level specification limit".to_string(),
                        ));
                    }
                    let level = self.parse_list_level(list.levels.len() as u8)?;
                    list.add_level(level);
                    continue;
                },
                Some(Token::OpenBrace)
                    if matches!(
                        self.tokens.get(self.pos + 1),
                        Some(Token::Control(ControlWord::ListName))
                    ) =>
                {
                    let name = self.parse_list_text_group(true, false)?;
                    list.name = Cow::Borrowed(self.arena.alloc_str(&name));
                    continue;
                },
                Some(Token::OpenBrace)
                    if matches!(
                        self.tokens.get(self.pos..self.pos + 3),
                        Some([
                            Token::OpenBrace,
                            Token::Control(ControlWord::IgnorableDestination),
                            Token::Control(ControlWord::ListStyleName),
                        ])
                    ) =>
                {
                    let name = self.parse_list_text_group(true, false)?;
                    list.style_name = Cow::Borrowed(self.arena.alloc_str(&name));
                    continue;
                },
                Some(Token::OpenBrace) => {
                    self.skip_group()?;
                    continue;
                },
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    closed = true;
                    break;
                },
                Some(Token::Control(control)) => match control {
                    ControlWord::ListTemplateId(value) => {
                        list.template_id = *value;
                        has_template_id = true;
                    },
                    ControlWord::ListSimple(value) => list.simple = *value,
                    ControlWord::ListHybrid(value) => list.hybrid = *value,
                    ControlWord::ListId(value) => {
                        list.id = *value;
                        has_id = true;
                    },
                    ControlWord::StylePriority(value) => list.style_priority = Some(*value),
                    _ => {},
                },
                Some(_) => {},
                None => return Err(RtfError::UnexpectedEof),
            }
            self.pos += 1;
        }
        if !closed {
            return Err(RtfError::UnexpectedEof);
        }
        if list.simple && (list.hybrid || list.levels.len() > 1) {
            return Err(RtfError::MalformedDocument(
                "invalid simple RTF list definition".to_string(),
            ));
        }
        if !has_template_id {
            list.template_id = list.id;
        }
        if has_id {
            if self.list_table.lists().len() >= MAX_LISTS {
                return Err(RtfError::MalformedDocument(
                    "RTF list count exceeds the safety limit".to_string(),
                ));
            }
            self.list_table.add(list);
        }
        Ok(())
    }

    fn parse_list_level(&mut self, level_index: u8) -> RtfResult<super::list::ListLevel<'a>> {
        self.pos += 2; // opening brace and `listlevel`
        let mut level = super::list::ListLevel::new(level_index);
        let mut explicit_indent = false;
        while self.pos < self.tokens.len() {
            match self.tokens.get(self.pos) {
                Some(Token::OpenBrace)
                    if matches!(
                        self.tokens.get(self.pos + 1),
                        Some(Token::Control(ControlWord::ListNumberText))
                    ) =>
                {
                    let text = self.parse_list_text_group(false, true)?;
                    level.number_text = Cow::Borrowed(self.arena.alloc_str(&text));
                    continue;
                },
                Some(Token::OpenBrace)
                    if matches!(
                        self.tokens.get(self.pos + 1),
                        Some(Token::Control(ControlWord::ListLevelNumbers))
                    ) =>
                {
                    let positions = self.parse_list_text_group(false, false)?;
                    level.number_positions = Cow::Borrowed(self.arena.alloc_str(&positions));
                    continue;
                },
                Some(Token::OpenBrace) => {
                    self.skip_group()?;
                    continue;
                },
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    return Ok(level);
                },
                Some(Token::Control(control)) => match control {
                    ControlWord::ListLevelType(value) => {
                        level.level_type = Self::list_level_type(*value);
                    },
                    ControlWord::ListLevelJustification(value) => {
                        level.justification = match value {
                            1 => super::list::ListJustification::Center,
                            2 => super::list::ListJustification::Right,
                            _ => super::list::ListJustification::Left,
                        };
                    },
                    ControlWord::ListLevelFollow(value) => {
                        level.follow = match value {
                            1 => super::list::ListFollow::Space,
                            2 => super::list::ListFollow::Nothing,
                            _ => super::list::ListFollow::Tab,
                        };
                        level.follow_previous = *value != 0;
                    },
                    ControlWord::ListLevelStartAt(value) => level.start_at = *value,
                    ControlWord::ListLevelSpace(value) => level.space = *value,
                    ControlWord::ListLevelIndent(value) => {
                        level.indent = *value;
                        explicit_indent = true;
                    },
                    ControlWord::FontNumber(value) => {
                        level.font_ref = u16::try_from(*value).map_err(|_| {
                            RtfError::MalformedDocument(
                                "RTF list font reference is outside the supported range"
                                    .to_string(),
                            )
                        })?;
                    },
                    ControlWord::LeftIndent(value) => {
                        level.left_indent = Some(*value);
                        if !explicit_indent {
                            level.indent = *value;
                        }
                    },
                    ControlWord::FirstLineIndent(value) => level.first_line_indent = Some(*value),
                    ControlWord::TabPosition(value) => {
                        let value = value.ok_or_else(|| {
                            RtfError::MalformedDocument(
                                "RTF list-level tx control requires a numeric parameter"
                                    .to_string(),
                            )
                        })?;
                        if level.tabs.len() >= MAX_LIST_TABS {
                            return Err(RtfError::MalformedDocument("RTF list level has too many tabs".to_string()));
                        }
                        level.tabs.push(value);
                    },
                    ControlWord::ListLevelPicture(value) => {
                        level.picture_index = Some(u32::try_from(*value).map_err(|_| {
                            RtfError::MalformedDocument("RTF list picture index cannot be negative".to_string())
                        })?);
                    },
                    _ => {},
                },
                Some(_) => {},
                None => break,
            }
            self.pos += 1;
        }
        Err(RtfError::UnexpectedEof)
    }

    fn parse_list_text_group(&mut self, is_name: bool, strip_length: bool) -> RtfResult<String> {
        self.pos += if matches!(
            self.tokens.get(self.pos + 1),
            Some(Token::Control(ControlWord::IgnorableDestination))
        ) { 3 } else { 2 };
        let mut value = String::new();
        let mut unicode_skip = self.current_state()?.unicode_skip.max(0);
        while self.pos < self.tokens.len() {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    let trimmed = value.trim_end_matches(['\r', '\n', ' ']);
                    let trimmed = trimmed.strip_suffix(';').unwrap_or(trimmed);
                    if is_name {
                        return Ok(trimmed.trim().to_string());
                    }
                    if strip_length {
                        let mut chars = trimmed.chars();
                        if chars.next().is_some_and(|ch| u32::from(ch) <= u8::MAX.into()) {
                            return Ok(chars.collect());
                        }
                    }
                    return Ok(trimmed.to_string());
                },
                Some(Token::OpenBrace) => {
                    self.skip_group()?;
                    continue;
                },
                Some(Token::Text(text)) => value.push_str(&self.decode_transport_text(text)?),
                Some(Token::Control(ControlWord::Unicode(first))) => {
                    let decoded = self.parse_style_unicode(*first, unicode_skip)?;
                    value.push_str(&decoded);
                    if value.len() > MAX_LIST_TEXT_BYTES {
                        return Err(RtfError::MalformedDocument(
                            "RTF list text exceeds the safety limit".to_string(),
                        ));
                    }
                    continue;
                },
                Some(Token::Control(ControlWord::UnicodeSkip(value))) => {
                    unicode_skip = (*value).max(0);
                },
                Some(Token::Control(control)) if control_symbol_text(control).is_some() => {
                    value.push_str(control_symbol_text(control).unwrap_or_default());
                },
                Some(_) => {},
                None => break,
            }
            self.pos += 1;
            if value.len() > MAX_LIST_TEXT_BYTES {
                return Err(RtfError::MalformedDocument(
                    "RTF list text exceeds the safety limit".to_string(),
                ));
            }
        }
        Err(RtfError::UnexpectedEof)
    }

    fn list_level_type(value: i32) -> super::list::ListLevelType {
        match value {
            0 => super::list::ListLevelType::Decimal,
            1 => super::list::ListLevelType::UpperRoman,
            2 => super::list::ListLevelType::LowerRoman,
            3 => super::list::ListLevelType::UpperLetter,
            4 => super::list::ListLevelType::LowerLetter,
            5 => super::list::ListLevelType::Ordinal,
            6 => super::list::ListLevelType::CardinalText,
            7 => super::list::ListLevelType::OrdinalText,
            23 => super::list::ListLevelType::Bullet,
            255 => super::list::ListLevelType::None,
            other => super::list::ListLevelType::Other(other),
        }
    }

    fn parse_list_override_table(&mut self) -> RtfResult<()> {
        if self.saw_list_override_table {
            return Err(RtfError::MalformedDocument("RTF document contains multiple list override tables".to_string()));
        }
        if !self.saw_list_table || self.states.len() != 3
            || self.blocks.iter().any(|block| !block.text.trim().is_empty())
        {
            return Err(RtfError::MalformedDocument("RTF list override table must follow listtable in the root header".to_string()));
        }
        self.saw_list_override_table = true;
        self.pos += 1; // `listoverridetable`
        while self.pos < self.tokens.len() {
            match self.tokens.get(self.pos) {
                Some(Token::OpenBrace)
                    if matches!(
                        self.tokens.get(self.pos + 1),
                        Some(Token::Control(ControlWord::ListOverride))
                    ) =>
                {
                    self.parse_list_override()?;
                },
                Some(Token::OpenBrace) => self.skip_group()?,
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    self.list_override_table.validate(&self.list_table)?;
                    return Ok(());
                },
                Some(_) => self.pos += 1,
                None => break,
            }
        }
        Err(RtfError::UnexpectedEof)
    }

    fn parse_list_override(&mut self) -> RtfResult<()> {
        self.pos += 2; // opening brace and `listoverride`
        let mut list_id = None;
        let mut index = None;
        let mut level_count = None;
        let mut start_at = None;
        let mut override_levels = Vec::new();
        let mut closed = false;
        while self.pos < self.tokens.len() {
            match self.tokens.get(self.pos) {
                Some(Token::OpenBrace)
                    if matches!(
                        self.tokens.get(self.pos + 1),
                        Some(Token::Control(ControlWord::ListOverrideLevel))
                    ) =>
                {
                    self.pos += 2;
                    let mut has_start_override = false;
                    let mut has_format_override = false;
                    let mut level_start_at = None;
                    let override_index = u8::try_from(override_levels.len()).map_err(|_| {
                        RtfError::MalformedDocument("RTF list override has too many levels".to_string())
                    })?;
                    while self.pos < self.tokens.len() {
                        match self.tokens.get(self.pos) {
                            Some(Token::CloseBrace) => {
                                self.pos += 1;
                                break;
                            },
                            Some(Token::Control(ControlWord::ListOverrideStartAt(value))) => {
                                has_start_override = *value;
                            },
                            Some(Token::Control(ControlWord::ListOverrideFormat(value))) => {
                                has_format_override = *value;
                            },
                            Some(Token::Control(ControlWord::ListLevelStartAt(value)))
                                if has_start_override =>
                            {
                                level_start_at = Some(*value);
                                start_at = Some(*value);
                            },
                            Some(Token::OpenBrace) => {
                                self.skip_group()?;
                                continue;
                            },
                            Some(_) => {},
                            None => return Err(RtfError::UnexpectedEof),
                        }
                        self.pos += 1;
                    }
                    let level_start = if has_start_override {
                        level_start_at
                    } else {
                        None
                    };
                    override_levels.push(super::list::ListOverrideLevel {
                        level: override_index,
                        start_at: level_start,
                        format_override: has_format_override,
                    });
                    continue;
                },
                Some(Token::OpenBrace) => {
                    self.skip_group()?;
                    continue;
                },
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    closed = true;
                    break;
                },
                Some(Token::Control(ControlWord::ListId(value))) => list_id = Some(*value),
                Some(Token::Control(ControlWord::ListOverrideIndex(value))) => index = Some(*value),
                Some(Token::Control(ControlWord::ListOverrideCount(value))) => {
                    level_count = Some(u8::try_from(*value).map_err(|_| {
                        RtfError::MalformedDocument(
                            "RTF list override count is outside the supported range".to_string(),
                        )
                    })?);
                },
                Some(_) => {},
                None => return Err(RtfError::UnexpectedEof),
            }
            self.pos += 1;
        }
        if !closed {
            return Err(RtfError::UnexpectedEof);
        }
        if let (Some(index), Some(list_id)) = (index, list_id) {
            if self.list_override_table.overrides().len() >= MAX_LISTS {
                return Err(RtfError::MalformedDocument(
                    "RTF list override count exceeds the safety limit".to_string(),
                ));
            }
            let mut entry = super::list::ListOverride::new(index, list_id);
            entry.level_count_override = level_count;
            entry.start_at_override = start_at;
            entry.levels = override_levels;
            self.list_override_table.add(entry);
        }
        Ok(())
    }

    /// Parse the standard RTF stylesheet destination.
    fn parse_stylesheet(&mut self) -> RtfResult<()> {
        if self.saw_stylesheet {
            return Err(RtfError::MalformedDocument(
                "RTF document contains multiple stylesheet destinations".to_string(),
            ));
        }
        if self.states.len() != 3
            || self
                .blocks
                .iter()
                .any(|block| !block.text.trim().is_empty())
        {
            return Err(RtfError::MalformedDocument(
                "RTF stylesheet must occur in the root document header".to_string(),
            ));
        }
        self.saw_stylesheet = true;
        self.pos += 1; // `stylesheet`
        while self.pos < self.tokens.len() {
            match self.tokens.get(self.pos) {
                Some(Token::OpenBrace) => self.parse_style_entry()?,
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    self.stylesheet.validate()?;
                    return Ok(());
                },
                Some(_) => self.pos += 1,
                None => break,
            }
        }
        Err(RtfError::UnexpectedEof)
    }

    fn parse_style_entry(&mut self) -> RtfResult<()> {
        self.pos += 1; // opening brace
        let starred = matches!(
            self.tokens.get(self.pos),
            Some(Token::Control(ControlWord::IgnorableDestination))
        );
        if starred {
            self.pos += 1;
        }
        let mut style_type = None;
        let mut id = None;
        let inherited_unicode_skip = self.current_state()?.unicode_skip;
        let mut state = State::default();
        state.unicode_skip = inherited_unicode_skip;
        let mut name = String::new();
        let mut name_complete = false;
        let mut based_on = None;
        let mut next_style = None;
        let mut linked_style = None;
        let mut additive = false;
        let mut auto_update = false;
        let mut hidden = false;
        let mut locked = false;
        let mut semi_hidden = false;
        let mut unhide_when_used = false;
        let mut quick_format = false;
        let mut priority = None;
        let mut revision_id = None;
        let mut personal = false;
        let mut compose = false;
        let mut reply = false;
        let mut seen_metadata = std::collections::HashSet::new();
        let mut saw_content_before_selector = false;
        macro_rules! set_style_once {
            ($key:literal, $target:expr, $value:expr) => {{
                if !seen_metadata.insert($key) {
                    return Err(RtfError::MalformedDocument(format!(
                        "duplicate RTF style metadata control: {}",
                        $key
                    )));
                }
                $target = $value;
            }};
        }

        while self.pos < self.tokens.len() {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    break;
                },
                Some(Token::OpenBrace) => {
                    // Nested extension groups do not form part of the style name.
                    self.skip_group()?;
                    continue;
                },
                Some(Token::Text(text)) if !name_complete => {
                    let decoded = self.decode_transport_text(text)?;
                    if style_type.is_none() && name.is_empty() && decoded.trim().is_empty() {
                        self.pos += 1;
                        continue;
                    }
                    saw_content_before_selector = true;
                    Self::append_style_name(&mut name, &decoded, &mut name_complete);
                },
                Some(Token::Control(ControlWord::Unicode(first))) if !name_complete => {
                    saw_content_before_selector = true;
                    let decoded = self.parse_style_unicode(*first, state.unicode_skip)?;
                    Self::append_style_name(&mut name, &decoded, &mut name_complete);
                    if name.len() > MAX_STYLE_NAME_BYTES {
                        return Err(RtfError::MalformedDocument(
                            "RTF style name exceeds the safety limit".to_string(),
                        ));
                    }
                    continue;
                },
                Some(Token::Control(control)) => match control {
                    control if !name_complete && control_symbol_text(control).is_some() => {
                        Self::append_style_name(
                            &mut name,
                            control_symbol_text(control).unwrap_or_default(),
                            &mut name_complete,
                        );
                    },
                    ControlWord::ParagraphStyle(value) => {
                        if style_type.is_some() || saw_content_before_selector {
                            return Err(RtfError::MalformedDocument(
                                "RTF style selector is duplicated or out of order".to_string(),
                            ));
                        }
                        style_type = Some(super::stylesheet::StyleType::Paragraph);
                        id = Some(Self::style_id(*value, "style")?);
                    },
                    ControlWord::CharacterStyle(value) => {
                        if style_type.is_some() || saw_content_before_selector {
                            return Err(RtfError::MalformedDocument(
                                "RTF style selector is duplicated or out of order".to_string(),
                            ));
                        }
                        style_type = Some(super::stylesheet::StyleType::Character);
                        id = Some(Self::style_id(*value, "style")?);
                    },
                    ControlWord::SectionStyle(value) => {
                        if style_type.is_some() || saw_content_before_selector {
                            return Err(RtfError::MalformedDocument(
                                "RTF style selector is duplicated or out of order".to_string(),
                            ));
                        }
                        style_type = Some(super::stylesheet::StyleType::Section);
                        id = Some(Self::style_id(*value, "style")?);
                    },
                    ControlWord::TableStyle(value) => {
                        if style_type.is_some() || saw_content_before_selector {
                            return Err(RtfError::MalformedDocument(
                                "RTF style selector is duplicated or out of order".to_string(),
                            ));
                        }
                        style_type = Some(super::stylesheet::StyleType::Table);
                        id = Some(Self::style_id(*value, "style")?);
                    },
                    ControlWord::StyleBasedOn(value) => {
                        if !seen_metadata.insert("sbasedon") {
                            return Err(RtfError::MalformedDocument("duplicate RTF sbasedon".to_string()));
                        }
                        based_on = Some(Self::style_id(*value, "based-on style")?);
                    },
                    ControlWord::StyleNext(value) => {
                        if !seen_metadata.insert("snext") {
                            return Err(RtfError::MalformedDocument("duplicate RTF snext".to_string()));
                        }
                        next_style = Some(Self::style_id(*value, "next style")?);
                    },
                    ControlWord::StyleLink(value) => {
                        if !seen_metadata.insert("slink") {
                            return Err(RtfError::MalformedDocument("duplicate RTF slink".to_string()));
                        }
                        linked_style = Some(Self::style_id(*value, "linked style")?);
                    },
                    ControlWord::StyleAdditive(value) => set_style_once!("additive", additive, *value),
                    ControlWord::StyleAutoUpdate(value) => set_style_once!("sautoupd", auto_update, *value),
                    ControlWord::StyleHidden(value) => set_style_once!("shidden", hidden, *value),
                    ControlWord::StyleLocked(value) => set_style_once!("slocked", locked, *value),
                    ControlWord::StyleSemiHidden(value) => set_style_once!("ssemihidden", semi_hidden, *value),
                    ControlWord::StyleUnhideWhenUsed(value) => set_style_once!("sunhideused", unhide_when_used, *value),
                    ControlWord::StyleQuickFormat(value) => set_style_once!("sqformat", quick_format, *value),
                    ControlWord::StylePriority(value) => set_style_once!("spriority", priority, Some(*value)),
                    ControlWord::StyleRevisionId(value) => {
                        if !seen_metadata.insert("styrsid") {
                            return Err(RtfError::MalformedDocument("duplicate RTF styrsid".to_string()));
                        }
                        revision_id = Some(*value);
                    },
                    ControlWord::StylePersonal(value) => set_style_once!("spersonal", personal, *value),
                    ControlWord::StyleCompose(value) => set_style_once!("scompose", compose, *value),
                    ControlWord::StyleReply(value) => set_style_once!("sreply", reply, *value),
                    ControlWord::UnicodeSkip(value) => state.unicode_skip = (*value).max(0),
                    _ => {
                        saw_content_before_selector = style_type.is_none();
                        Self::apply_style_property(&mut state, control)?;
                    },
                },
                Some(_) => {},
                None => break,
            }
            self.pos += 1;
            if name.len() > MAX_STYLE_NAME_BYTES {
                return Err(RtfError::MalformedDocument(
                    "RTF style name exceeds the safety limit".to_string(),
                ));
            }
        }

        if style_type.is_none() && !starred && name_complete {
            style_type = Some(super::stylesheet::StyleType::Paragraph);
            id = Some(0);
        }
        let (Some(style_type), Some(id)) = (style_type, id) else {
            // Unknown starred extension groups are permitted inside a stylesheet.
            return Ok(());
        };
        if !name_complete {
            return Err(RtfError::MalformedDocument(
                "RTF style name must end with a semicolon".to_string(),
            ));
        }
        if style_type != super::stylesheet::StyleType::Paragraph && !starred {
            return Err(RtfError::MalformedDocument(
                "RTF non-paragraph style entries must be starred".to_string(),
            ));
        }
        if self.stylesheet.styles().len() >= MAX_STYLES {
            return Err(RtfError::MalformedDocument(
                "RTF style count exceeds the safety limit".to_string(),
            ));
        }
        let name = name.trim().to_string();
        let allocated = self.arena.alloc_str(&name);
        let mut style = match style_type {
            super::stylesheet::StyleType::Paragraph => {
                super::stylesheet::Style::paragraph(id, Cow::Borrowed(allocated))
            },
            super::stylesheet::StyleType::Character => {
                super::stylesheet::Style::character(id, Cow::Borrowed(allocated))
            },
            super::stylesheet::StyleType::Section => {
                super::stylesheet::Style::section(id, Cow::Borrowed(allocated))
            },
            super::stylesheet::StyleType::Table => {
                super::stylesheet::Style::table(id, Cow::Borrowed(allocated))
            },
        };
        style.based_on = based_on;
        style.next_style = next_style;
        style.linked_style = linked_style;
        style.formatting = state.formatting;
        if style_type == super::stylesheet::StyleType::Paragraph {
            style.paragraph = Some(state.paragraph);
        }
        style.hidden = hidden;
        style.additive = additive;
        style.auto_update = auto_update;
        style.locked = locked;
        style.semi_hidden = semi_hidden;
        style.unhide_when_used = unhide_when_used;
        style.quick_format = quick_format;
        style.priority = priority;
        style.revision_id = revision_id;
        style.personal = personal;
        style.compose = compose;
        style.reply = reply;
        self.stylesheet.add(style);
        Ok(())
    }

    fn style_id(value: i32, field: &str) -> RtfResult<u16> {
        u16::try_from(value).map_err(|_| {
            RtfError::MalformedDocument(format!("RTF {field} ID is outside the supported range"))
        })
    }

    fn append_style_name(name: &mut String, text: &str, complete: &mut bool) {
        if let Some((prefix, _)) = text.split_once(';') {
            name.push_str(prefix);
            *complete = true;
        } else {
            name.push_str(text);
        }
    }

    fn parse_style_unicode(&mut self, first_code: i32, unicode_skip: i32) -> RtfResult<String> {
        let mut utf16 = SmallVec::<[u16; 4]>::new();
        utf16.push(first_code as u16);
        self.pos += 1;
        while let Some(Token::Control(ControlWord::Unicode(code))) = self.tokens.get(self.pos) {
            utf16.push(*code as u16);
            self.pos += 1;
        }

        let mut fallback_skip = (unicode_skip.max(0) as usize).saturating_mul(utf16.len());
        let mut remainder = String::new();
        while fallback_skip > 0 && self.pos < self.tokens.len() {
            match self.tokens.get(self.pos) {
                Some(Token::Text(text)) => {
                    let count = text.chars().count();
                    if count <= fallback_skip {
                        fallback_skip -= count;
                    } else {
                        remainder.extend(text.chars().skip(fallback_skip));
                        fallback_skip = 0;
                    }
                    self.pos += 1;
                },
                Some(Token::Control(ControlWord::Unicode(_))) => break,
                Some(_) => {
                    fallback_skip -= 1;
                    self.pos += 1;
                },
                None => break,
            }
        }
        let mut decoded = String::from_utf16(&utf16)
            .map_err(|error| RtfError::InvalidUnicode(format!("invalid style name: {error}")))?;
        decoded.push_str(&self.decode_transport_text(&remainder)?);
        Ok(decoded)
    }

    fn required_character_value(
        value: Option<i32>,
        control: &str,
        maximum: u16,
    ) -> RtfResult<u16> {
        let value = value.ok_or_else(|| {
            RtfError::MalformedDocument(format!(
                "RTF {control} requires a numeric parameter"
            ))
        })?;
        let value = u16::try_from(value).map_err(|_| {
            RtfError::MalformedDocument(format!(
                "RTF {control} value must be in 0..={maximum}"
            ))
        })?;
        if value > maximum {
            return Err(RtfError::MalformedDocument(format!(
                "RTF {control} value must be in 0..={maximum}"
            )));
        }
        Ok(value)
    }

    fn table_border_style(control:&ControlWord<'_>)->Option<crate::BorderStyle>{use crate::BorderStyle as Style;Some(match control{ControlWord::BorderNone=>Style::None,ControlWord::BorderSingle=>Style::Single,ControlWord::BorderThick=>Style::Thick,ControlWord::BorderDotted=>Style::Dotted,ControlWord::BorderDashed=>Style::Dashed,ControlWord::BorderDashSmall=>Style::DashSmallGap,ControlWord::BorderDotDash=>Style::DotDash,ControlWord::BorderDotDotDash=>Style::DotDotDash,ControlWord::BorderDouble=>Style::Double,ControlWord::BorderTriple=>Style::Triple,ControlWord::BorderThinThickSmall=>Style::ThinThickSmall,ControlWord::BorderThickThinSmall=>Style::ThickThinSmall,ControlWord::BorderThinThickThinSmall=>Style::ThinThickThinSmall,ControlWord::BorderThinThickMedium=>Style::ThinThickMedium,ControlWord::BorderThickThinMedium=>Style::ThickThinMedium,ControlWord::BorderThinThickThinMedium=>Style::ThinThickThinMedium,ControlWord::BorderThinThickLarge=>Style::ThinThickLarge,ControlWord::BorderThickThinLarge=>Style::ThickThinLarge,ControlWord::BorderThinThickThinLarge=>Style::ThinThickThinLarge,ControlWord::BorderWave=>Style::Wavy,ControlWord::BorderWavyDouble=>Style::WavyDouble,ControlWord::BorderStriped=>Style::Striped,ControlWord::BorderEmbossed=>Style::Embossed,ControlWord::BorderEngraved=>Style::Engraved,ControlWord::BorderOutset=>Style::Outset,ControlWord::BorderInset=>Style::Inset,_=>return None})}

    fn apply_table_decoration_control(state:&mut State,control:&ControlWord<'_>)->RtfResult<bool>{
        const STYLE:u8=1;const WIDTH:u8=2;const COLOR:u8=4;const SPACE:u8=8;const SHADOW:u8=16;const FRAME:u8=32;
        if let ControlWord::TableBorder(target,param)=control{require_parameterless(*param,"table border selector")?;state.active_table_border=Some(*target);state.active_table_border_seen=0;let slot=match target{crate::table::TableBorderTarget::Row(side)=>state.table_row_borders.side_mut(*side),crate::table::TableBorderTarget::Cell(side)=>state.pending_cell_borders.side_mut(*side)};*slot=Some(crate::Border::default());return Ok(true)}
        let shading_control=matches!(control,ControlWord::TableShadingAmount(..)|ControlWord::TableShadingForeground(..)|ControlWord::TableShadingBackground(..)|ControlWord::TableShadingPattern(..)|ControlWord::TableRowShadingPatternIndex(..));
        if shading_control{state.active_table_border=None;state.active_table_border_seen=0;let scope=match control{ControlWord::TableShadingAmount(scope,_)|ControlWord::TableShadingForeground(scope,_)|ControlWord::TableShadingBackground(scope,_)|ControlWord::TableShadingPattern(scope,_,_)=>*scope,ControlWord::TableRowShadingPatternIndex(_)=>crate::TableDistanceScope::Row,_=>unreachable!()};let(shading,seen)=match scope{crate::TableDistanceScope::Row=>(&mut state.table_row_shading,&mut state.table_row_shading_seen),crate::TableDistanceScope::Cell=>(&mut state.pending_cell_shading,&mut state.pending_cell_shading_seen)};let bit=match control{ControlWord::TableShadingAmount(..)=>1,ControlWord::TableShadingForeground(..)=>2,ControlWord::TableShadingBackground(..)=>4,ControlWord::TableShadingPattern(..)|ControlWord::TableRowShadingPatternIndex(..)=>8,_=>unreachable!()};if *seen&bit!=0{return Err(RtfError::MalformedDocument("duplicate RTF table shading component".to_string()))}*seen|=bit;match control{ControlWord::TableShadingAmount(_,value)=>shading.amount=Some(required_table_value(*value,"table shading",10_000)?),ControlWord::TableShadingForeground(_,value)=>shading.foreground_color=Some(required_table_value(*value,"table shading foreground color",u16::MAX)?),ControlWord::TableShadingBackground(_,value)=>shading.background_color=Some(required_table_value(*value,"table shading background color",u16::MAX)?),ControlWord::TableShadingPattern(_,pattern,param)=>{require_parameterless(*param,"table shading pattern")?;shading.pattern=Some(*pattern)},ControlWord::TableRowShadingPatternIndex(value)=>shading.pattern_index=Some(required_table_value(*value,"trpat",u16::MAX)?),_=>unreachable!()}return Ok(true)}
        let Some(target)=state.active_table_border else{return Ok(false)};let(component,name)=if Self::table_border_style(control).is_some(){(STYLE,"style")}else{match control{ControlWord::BorderWidth(_)=>(WIDTH,"width"),ControlWord::BorderColor(_)=>(COLOR,"color"),ControlWord::BorderSpace(_)=>(SPACE,"spacing"),ControlWord::BorderShadow=>(SHADOW,"shadow"),ControlWord::BorderFrame=>(FRAME,"frame"),_=>{state.active_table_border=None;state.active_table_border_seen=0;return Ok(false)}}};if state.active_table_border_seen&component!=0{return Err(RtfError::MalformedDocument(format!("duplicate RTF table-border {name}")))}if component!=STYLE&&state.active_table_border_seen&STYLE==0{return Err(RtfError::MalformedDocument(format!("RTF table-border {name} precedes its style")))}state.active_table_border_seen|=component;let border=match target{crate::table::TableBorderTarget::Row(side)=>state.table_row_borders.side_mut(side),crate::table::TableBorderTarget::Cell(side)=>state.pending_cell_borders.side_mut(side)}.as_mut().expect("active table border");if let Some(style)=Self::table_border_style(control){border.style=style}else{match control{ControlWord::BorderWidth(value)=>border.width=i32::from(required_table_value(*value,"brdrw",75)?),ControlWord::BorderColor(value)=>border.color_ref=required_table_value(*value,"brdrcf",u16::MAX)?,ControlWord::BorderSpace(value)=>border.space=i32::from(required_table_value(*value,"brsp",crate::MAX_TABLE_DISTANCE_TWIPS as u16)?),ControlWord::BorderShadow=>border.shadow=true,ControlWord::BorderFrame=>border.frame=true,_=>unreachable!()}}Ok(true)
    }

    fn character_border_style(
        control: &ControlWord<'_>,
    ) -> Option<crate::CharacterBorderStyle> {
        use crate::CharacterBorderStyle as Style;
        Some(match control {
            ControlWord::BorderNone => Style::None,
            ControlWord::BorderSingle => Style::Single,
            ControlWord::BorderThick => Style::Thick,
            ControlWord::BorderDotted => Style::Dotted,
            ControlWord::BorderDashed => Style::Dashed,
            ControlWord::BorderDashSmall => Style::DashSmallGap,
            ControlWord::BorderDotDash => Style::DotDash,
            ControlWord::BorderDotDotDash => Style::DotDotDash,
            ControlWord::BorderDouble => Style::Double,
            ControlWord::BorderTriple => Style::Triple,
            ControlWord::BorderThinThickSmall => Style::ThinThickSmallGap,
            ControlWord::BorderThickThinSmall => Style::ThickThinSmallGap,
            ControlWord::BorderThinThickThinSmall => Style::ThinThickThinSmallGap,
            ControlWord::BorderThinThickMedium => Style::ThinThickMediumGap,
            ControlWord::BorderThickThinMedium => Style::ThickThinMediumGap,
            ControlWord::BorderThinThickThinMedium => Style::ThinThickThinMediumGap,
            ControlWord::BorderThinThickLarge => Style::ThinThickLargeGap,
            ControlWord::BorderThickThinLarge => Style::ThickThinLargeGap,
            ControlWord::BorderThinThickThinLarge => Style::ThinThickThinLargeGap,
            ControlWord::BorderWave => Style::Wavy,
            ControlWord::BorderWavyDouble => Style::DoubleWavy,
            ControlWord::BorderStriped => Style::Striped,
            ControlWord::BorderEmbossed => Style::Embossed,
            ControlWord::BorderEngraved => Style::Engraved,
            ControlWord::BorderOutset => Style::Outset,
            ControlWord::BorderInset => Style::Inset,
            _ => return None,
        })
    }

    fn apply_character_decoration_control(
        state: &mut State,
        control: &ControlWord<'_>,
    ) -> RtfResult<bool> {
        const STYLE: u8 = 1;
        const WIDTH: u8 = 2;
        const COLOR: u8 = 4;
        const SPACE: u8 = 8;
        const SHADOW: u8 = 16;
        const FRAME: u8 = 32;

        match control {
            ControlWord::CharacterBorder(parameter) => {
                if parameter.is_some() {
                    return Err(RtfError::MalformedDocument(
                        "RTF chbrdr must not have a numeric parameter".to_string(),
                    ));
                }
                state.formatting.character_border = Some(crate::CharacterBorder::default());
                state.character_border_active = true;
                state.character_border_seen = 0;
                return Ok(true);
            },
            ControlWord::CharacterShading(value) => {
                state.character_border_active = false;
                let amount = Self::required_character_value(*value, "chshdng", 10_000)?;
                state
                    .formatting
                    .character_shading
                    .get_or_insert_default()
                    .amount = amount;
                return Ok(true);
            },
            ControlWord::CharacterForegroundPattern(value) => {
                state.character_border_active = false;
                let color = Self::required_character_value(*value, "chcfpat", u16::MAX)?;
                state
                    .formatting
                    .character_shading
                    .get_or_insert_default()
                    .foreground_color = color;
                return Ok(true);
            },
            ControlWord::CharacterBackgroundPattern(value) => {
                state.character_border_active = false;
                let color = Self::required_character_value(*value, "chcbpat", u16::MAX)?;
                state
                    .formatting
                    .character_shading
                    .get_or_insert_default()
                    .background_color = color;
                return Ok(true);
            },
            _ => {},
        }

        if !state.character_border_active {
            return Ok(false);
        }
        let (component, duplicate_name) = if Self::character_border_style(control).is_some() {
            (STYLE, "style")
        } else {
            match control {
                ControlWord::BorderWidth(_) => (WIDTH, "width"),
                ControlWord::BorderColor(_) => (COLOR, "color"),
                ControlWord::BorderSpace(_) => (SPACE, "space"),
                ControlWord::BorderShadow => (SHADOW, "shadow"),
                ControlWord::BorderFrame => (FRAME, "frame"),
                _ => {
                    state.character_border_active = false;
                    return Ok(false);
                },
            }
        };
        if state.character_border_seen & component != 0 {
            return Err(RtfError::MalformedDocument(format!(
                "duplicate RTF character-border {duplicate_name}"
            )));
        }
        state.character_border_seen |= component;
        let border = state
            .formatting
            .character_border
            .as_mut()
            .expect("active character border");
        if let Some(style) = Self::character_border_style(control) {
            border.style = style;
        } else {
            match control {
                ControlWord::BorderWidth(value) => {
                    border.width = Self::required_character_value(*value, "brdrw", 75)?;
                },
                ControlWord::BorderColor(value) => {
                    border.color_ref =
                        Self::required_character_value(*value, "brdrcf", u16::MAX)?;
                },
                ControlWord::BorderSpace(value) => {
                    border.space = Self::required_character_value(*value, "brsp", u16::MAX)?;
                },
                ControlWord::BorderShadow => border.shadow = true,
                ControlWord::BorderFrame => border.frame = true,
                _ => unreachable!("classified character-border component"),
            }
        }
        Ok(true)
    }

    fn apply_style_property(state: &mut State, control: &ControlWord<'_>) -> RtfResult<()> {
        if Self::apply_paragraph_tab_control(state, control)? {
            return Ok(());
        }
        if Self::apply_character_decoration_control(state, control)? {
            return Ok(());
        }
        match control {
            ControlWord::FontNumber(value) => state.formatting.font_ref = *value as FontRef,
            ControlWord::FontSize(value) => {
                if let Some(size) = NonZeroU16::new((*value).clamp(1, i32::from(u16::MAX)) as u16) {
                    state.formatting.font_size = size;
                }
            },
            ControlWord::AssociatedFontNumber(value) => {
                state.formatting.associated.font_ref = Some(associated_font_ref(*value)?);
            },
            ControlWord::AssociatedFontSize(value) => {
                state.formatting.associated.font_size = Some(associated_font_size(*value)?);
            },
            ControlWord::AssociatedLanguage(value) => {
                state.formatting.associated.language = Some(associated_language(*value)?);
            },
            ControlWord::AssociatedBold(value) => {
                state.formatting.associated.bold = Some(*value);
            },
            ControlWord::AssociatedItalic(value) => {
                state.formatting.associated.italic = Some(*value);
            },
            ControlWord::ColorForeground(value) => {
                state.formatting.color_ref = *value as ColorRef;
            },
            ControlWord::Highlight(value) => {
                state.formatting.highlight_color = Some(*value as ColorRef);
            },
            ControlWord::Bold(value) => state.formatting.bold = *value,
            ControlWord::Italic(value) => state.formatting.italic = *value,
            ControlWord::Underline(value) => {
                state.formatting.underline = if *value {
                    UnderlineStyle::Single
                } else {
                    UnderlineStyle::None
                };
            },
            ControlWord::UnderlineNone => state.formatting.underline = UnderlineStyle::None,
            ControlWord::UnderlineDouble => state.formatting.underline = UnderlineStyle::Double,
            ControlWord::UnderlineDotted => state.formatting.underline = UnderlineStyle::Dotted,
            ControlWord::UnderlineDashed => state.formatting.underline = UnderlineStyle::Dashed,
            ControlWord::UnderlineDashDot => state.formatting.underline = UnderlineStyle::DashDot,
            ControlWord::UnderlineDashDotDot => {
                state.formatting.underline = UnderlineStyle::DashDotDot;
            },
            ControlWord::UnderlineWords => state.formatting.underline = UnderlineStyle::Words,
            ControlWord::UnderlineThick => state.formatting.underline = UnderlineStyle::Thick,
            ControlWord::UnderlineWave => state.formatting.underline = UnderlineStyle::Wave,
            ControlWord::Strike(value) => state.formatting.strike = *value,
            ControlWord::DoubleStrike(value) => state.formatting.double_strike = *value,
            ControlWord::Superscript(value) => { state.formatting.superscript = *value; if *value { state.formatting.subscript = false; } state.formatting.character_positioning.set_superscript(*value); },
            ControlWord::Subscript(value) => { state.formatting.subscript = *value; if *value { state.formatting.superscript = false; } state.formatting.character_positioning.set_subscript(*value); },
            ControlWord::NoSuperSub => { state.formatting.superscript = false; state.formatting.subscript = false; state.formatting.character_positioning.clear_baseline(); },
            ControlWord::BaselineUp(value) => { state.formatting.superscript = false; state.formatting.subscript = false; state.formatting.character_positioning.set_raised(*value)?; },
            ControlWord::BaselineDown(value) => { state.formatting.superscript = false; state.formatting.subscript = false; state.formatting.character_positioning.set_lowered(*value)?; },
            ControlWord::SmallCaps(value) => state.formatting.smallcaps = *value,
            ControlWord::AllCaps(value) => state.formatting.all_caps = *value,
            ControlWord::Hidden(value) => state.formatting.hidden = *value,
            ControlWord::Outline(value) => state.formatting.outline = *value,
            ControlWord::Shadow(value) => state.formatting.shadow = *value,
            ControlWord::Emboss(value) => state.formatting.emboss = *value,
            ControlWord::Imprint(value) => state.formatting.imprint = *value,
            ControlWord::CharSpacing(value) => { state.formatting.character_positioning.set_quarter_point_expansion(*value)?; state.formatting.char_spacing = *value; },
            ControlWord::CharSpacingTwips(value) => { state.formatting.character_positioning.set_twip_expansion(*value)?; state.formatting.char_spacing = *value; },
            ControlWord::CharScale(value) => { state.formatting.character_positioning.set_scale(*value)?; state.formatting.char_scale = *value; },
            ControlWord::Kerning(value) => { state.formatting.character_positioning.set_kerning(*value)?; state.formatting.kerning = *value; },
            ControlWord::Language(value) => {
                state.formatting.language = crate::LanguageId::from_rtf(*value).ok();
            },
            ControlWord::LanguageEastAsian(value) => {
                state.formatting.east_asian_language = crate::LanguageId::from_rtf(*value).ok();
            },
            ControlWord::LanguageNoProof(value) => {
                state.formatting.language_no_proof = crate::LanguageId::from_rtf(*value).ok();
            },
            ControlWord::LanguageEastAsianNoProof(value) => {
                state.formatting.east_asian_language_no_proof =
                    crate::LanguageId::from_rtf(*value).ok();
            },
            ControlWord::NoProof(value) => state.formatting.no_proof = *value,
            ControlWord::LeftToRightCharacter => {
                state.formatting.direction = Some(TextDirection::LeftToRight);
            },
            ControlWord::RightToLeftCharacter => {
                state.formatting.direction = Some(TextDirection::RightToLeft);
            },
            ControlWord::Plain => {
                state.formatting = Formatting::default();
                state.character_border_active = false;
                state.character_border_seen = 0;
            },
            ControlWord::LeftAlign => state.paragraph.alignment = Alignment::Left,
            ControlWord::RightAlign => state.paragraph.alignment = Alignment::Right,
            ControlWord::Center => state.paragraph.alignment = Alignment::Center,
            ControlWord::Justify => state.paragraph.alignment = Alignment::Justify,
            ControlWord::LeftToRightParagraph => {
                state.paragraph.direction = Some(TextDirection::LeftToRight);
            },
            ControlWord::RightToLeftParagraph => {
                state.paragraph.direction = Some(TextDirection::RightToLeft);
            },
            ControlWord::Pard => {
                state.paragraph = Paragraph::default();
                state.pending_tab_alignment = None;
                state.pending_tab_leader = None;
            },
            ControlWord::SpaceBefore(value) => state.paragraph.spacing.before = *value,
            ControlWord::SpaceAfter(value) => state.paragraph.spacing.after = *value,
            ControlWord::SpaceBetween(value) => state.paragraph.spacing.line = *value,
            ControlWord::LineMultiple(value) => state.paragraph.spacing.line_multiple = *value,
            ControlWord::SpaceBeforeAuto(value) => state.paragraph.spacing_policy.automatic_before = required_paragraph_bool(*value, "sbauto")?,
            ControlWord::SpaceAfterAuto(value) => state.paragraph.spacing_policy.automatic_after = required_paragraph_bool(*value, "saauto")?,
            ControlWord::ListSpaceBefore(value) => state.paragraph.spacing_policy.list_before = Some(required_list_spacing(*value, "lisb")?),
            ControlWord::ListSpaceAfter(value) => state.paragraph.spacing_policy.list_after = Some(required_list_spacing(*value, "lisa")?),
            ControlWord::NoSnapLineGrid(value) => { strict_paragraph_selector(*value, "nosnaplinegrid")?; state.paragraph.spacing_policy.snap_to_line_grid = false; },
            ControlWord::ContextualSpacing(value) => { strict_paragraph_selector(*value, "contextualspace")?; state.paragraph.spacing_policy.contextual_spacing = true; },
            ControlWord::LeftIndent(value) => state.paragraph.indentation.left = *value,
            ControlWord::RightIndent(value) => state.paragraph.indentation.right = *value,
            ControlWord::FirstLineIndent(value) => {
                state.paragraph.indentation.first_line = *value;
            },
            ControlWord::LogicalLeftIndent(v)=>state.paragraph.logical_indentation.start=Some(required_paragraph_indent(*v,"lin")?), ControlWord::LogicalRightIndent(v)=>state.paragraph.logical_indentation.end=Some(required_paragraph_indent(*v,"rin")?), ControlWord::CharacterFirstLineIndent(v)=>state.paragraph.logical_indentation.first_line_character_units=Some(required_paragraph_indent(*v,"cufi")?), ControlWord::CharacterLeftIndent(v)=>state.paragraph.logical_indentation.left_character_units=Some(required_paragraph_indent(*v,"culi")?), ControlWord::CharacterRightIndent(v)=>state.paragraph.logical_indentation.right_character_units=Some(required_paragraph_indent(*v,"curi")?), ControlWord::MirrorIndents(v)=>{strict_paragraph_selector(*v,"indmirror")?;state.paragraph.logical_indentation.mirrored=true;},
            ControlWord::KeepTogether => state.paragraph.keep_together = true,
            ControlWord::KeepNext => state.paragraph.keep_next = true,
            ControlWord::PageBreakBefore => state.paragraph.page_break_before = true,
            ControlWord::WidowControl => state.paragraph.widow_control = true,
            ControlWord::ParagraphHyphenation(value) => state.paragraph.line_breaking.automatic_hyphenation = strict_paragraph_toggle(*value, "hyphpar")?,
            ControlWord::AutoSpaceAlphabetic(value) => state.paragraph.line_breaking.auto_space_alphabetic = strict_paragraph_toggle(*value, "aspalpha")?,
            ControlWord::AutoSpaceNumbers(value) => state.paragraph.line_breaking.auto_space_numbers = strict_paragraph_toggle(*value, "aspnum")?,
            ControlWord::AdjustRightIndent(value) => state.paragraph.line_breaking.adjust_right_indent = strict_paragraph_toggle(*value, "adjustright")?,
            ControlWord::WrapDefault(value) => { strict_paragraph_selector(*value, "wrapdefault")?; state.paragraph.line_breaking.wrapping = crate::ParagraphWrapping::Default; },
            ControlWord::NoCharacterWrap(value) => { strict_paragraph_selector(*value, "nocwrap")?; state.paragraph.line_breaking.wrapping = crate::ParagraphWrapping::NoCharacterWrap; },
            ControlWord::NoWordWrap(value) => { strict_paragraph_selector(*value, "nowwrap")?; state.paragraph.line_breaking.wrapping = crate::ParagraphWrapping::NoWordWrap; },
            ControlWord::NoOverflow(value) => { strict_paragraph_selector(*value, "nooverflow")?; state.paragraph.line_breaking.wrapping = crate::ParagraphWrapping::NoOverflow; },
            ControlWord::FontAlignAuto(value) => { strict_paragraph_selector(*value, "faauto")?; state.paragraph.line_breaking.font_alignment = crate::ParagraphFontAlignment::Auto; },
            ControlWord::FontAlignHanging(value) => { strict_paragraph_selector(*value, "fahang")?; state.paragraph.line_breaking.font_alignment = crate::ParagraphFontAlignment::Hanging; },
            ControlWord::FontAlignCenter(value) => { strict_paragraph_selector(*value, "facenter")?; state.paragraph.line_breaking.font_alignment = crate::ParagraphFontAlignment::Center; },
            ControlWord::FontAlignRoman(value) => { strict_paragraph_selector(*value, "faroman")?; state.paragraph.line_breaking.font_alignment = crate::ParagraphFontAlignment::Roman; },
            ControlWord::FontAlignVariable(value) => { strict_paragraph_selector(*value, "favar")?; state.paragraph.line_breaking.font_alignment = crate::ParagraphFontAlignment::Variable; },
            ControlWord::FontAlignFixed(value) => { strict_paragraph_selector(*value, "fafixed")?; state.paragraph.line_breaking.font_alignment = crate::ParagraphFontAlignment::Fixed; },
            ControlWord::ListOverrideIndex(value) => {
                state.paragraph.list_override = Some(*value);
            },
            ControlWord::ListLevelIndex(value) => {
                if let Ok(level @ 0..=8) = u8::try_from(*value) {
                    state.paragraph.list_level = Some(level);
                }
            },
            _ => {},
        }
        Ok(())
    }

    fn apply_paragraph_tab_control(
        state: &mut State,
        control: &ControlWord<'_>,
    ) -> RtfResult<bool> {
        use super::border::{TabAlignment, TabLeader, TabStop};

        fn require_flag(parameter: Option<i32>, name: &str) -> RtfResult<()> {
            if parameter.is_some() {
                return Err(RtfError::MalformedDocument(format!(
                    "RTF {name} tab selector does not accept a numeric parameter"
                )));
            }
            Ok(())
        }

        fn select_alignment(
            state: &mut State,
            parameter: Option<i32>,
            name: &str,
            alignment: TabAlignment,
        ) -> RtfResult<()> {
            require_flag(parameter, name)?;
            if state.pending_tab_alignment.is_some() || state.pending_tab_leader.is_some() {
                return Err(RtfError::MalformedDocument(
                    "RTF tab alignment must occur once and before its leader".to_string(),
                ));
            }
            state.pending_tab_alignment = Some(alignment);
            Ok(())
        }

        fn select_leader(
            state: &mut State,
            parameter: Option<i32>,
            name: &str,
            leader: TabLeader,
        ) -> RtfResult<()> {
            require_flag(parameter, name)?;
            if state.pending_tab_leader.is_some() {
                return Err(RtfError::MalformedDocument(
                    "RTF tab definition contains multiple leader selectors".to_string(),
                ));
            }
            state.pending_tab_leader = Some(leader);
            Ok(())
        }

        fn append(state: &mut State, position: Option<i32>, bar: bool) -> RtfResult<()> {
            let position = position.ok_or_else(|| {
                RtfError::MalformedDocument(format!(
                    "RTF {} control requires a numeric parameter",
                    if bar { "tb" } else { "tx" }
                ))
            })?;
            if bar && state.pending_tab_alignment.is_some() {
                return Err(RtfError::MalformedDocument(
                    "RTF bar tab cannot have a tab-alignment selector".to_string(),
                ));
            }
            let tab = TabStop {
                position,
                alignment: if bar {
                    TabAlignment::Bar
                } else {
                    state.pending_tab_alignment.unwrap_or(TabAlignment::Left)
                },
                leader: state.pending_tab_leader.unwrap_or(TabLeader::None),
            };
            state.paragraph.tab_stops.push(tab).map_err(|_| {
                RtfError::MalformedDocument(
                    "RTF paragraph exceeds the 64-tab safety limit".to_string(),
                )
            })?;
            state.pending_tab_alignment = None;
            state.pending_tab_leader = None;
            Ok(())
        }

        match control {
            ControlWord::TabLeft(parameter) => {
                select_alignment(state, *parameter, "tql", TabAlignment::Left)?;
            },
            ControlWord::TabRight(parameter) => {
                select_alignment(state, *parameter, "tqr", TabAlignment::Right)?;
            },
            ControlWord::TabCenter(parameter) => {
                select_alignment(state, *parameter, "tqc", TabAlignment::Center)?;
            },
            ControlWord::TabDecimal(parameter) => {
                select_alignment(state, *parameter, "tqdec", TabAlignment::Decimal)?;
            },
            ControlWord::TabLeaderDot(parameter) => {
                select_leader(state, *parameter, "tldot", TabLeader::Dot)?;
            },
            ControlWord::TabLeaderMiddleDot(parameter) => {
                select_leader(state, *parameter, "tlmdot", TabLeader::MiddleDot)?;
            },
            ControlWord::TabLeaderHyphen(parameter) => {
                select_leader(state, *parameter, "tlhyph", TabLeader::Hyphen)?;
            },
            ControlWord::TabLeaderUnderscore(parameter) => {
                select_leader(state, *parameter, "tlul", TabLeader::Underscore)?;
            },
            ControlWord::TabLeaderThick(parameter) => {
                select_leader(state, *parameter, "tlth", TabLeader::ThickLine)?;
            },
            ControlWord::TabLeaderEqual(parameter) => {
                select_leader(state, *parameter, "tleq", TabLeader::Equal)?;
            },
            ControlWord::TabPosition(position) => append(state, *position, false)?,
            ControlWord::TabBar(position) => append(state, *position, true)?,
            _ => return Ok(false),
        }
        Ok(true)
    }

    fn parse_file_table(&mut self) -> RtfResult<crate::FileTable<'a>> {
        if self.states.len() != 3
            || self.blocks.iter().any(|block| !block.text.trim().is_empty())
        {
            return Err(RtfError::MalformedDocument(
                "RTF filetbl must occur at document scope before body text".to_string(),
            ));
        }
        self.pos += 1; // ignorable-destination marker
        if !matches!(self.tokens.get(self.pos), Some(Token::Control(ControlWord::FileTable))) {
            return Err(RtfError::MalformedDocument("invalid RTF filetbl destination".to_string()));
        }
        self.pos += 1;
        let mut table = crate::FileTable::new();
        while self.pos < self.tokens.len() {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    table.validate()?;
                    return Ok(table);
                },
                Some(Token::OpenBrace)
                    if matches!(self.tokens.get(self.pos + 1), Some(Token::Control(ControlWord::FileEntry))) =>
                {
                    let entry = self.parse_file_table_entry()?;
                    table.add(entry)?;
                    continue;
                },
                Some(Token::Text(text)) if text.trim().is_empty() => {},
                Some(Token::OpenBrace) => return Err(RtfError::MalformedDocument(
                    "RTF filetbl cannot contain fields, objects, or unknown destinations".to_string(),
                )),
                Some(_) => return Err(RtfError::MalformedDocument("invalid content in RTF filetbl".to_string())),
                None => return Err(RtfError::UnexpectedEof),
            }
            self.pos += 1;
        }
        Err(RtfError::UnexpectedEof)
    }

    fn parse_file_table_entry(&mut self) -> RtfResult<crate::FileTableEntry<'a>> {
        self.pos += 2; // opening brace and file
        let mut id = None;
        let mut relative = None;
        let mut operating_system = None;
        let mut valid_on = crate::FileSystemValidity::default();
        let mut location = crate::FileLocation::Local;
        let mut name = String::new();
        let mut unicode_skip = self.current_state()?.unicode_skip.max(0);
        let mut seen = std::collections::HashSet::new();
        while self.pos < self.tokens.len() {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    let trimmed = name.trim_end_matches(['\r', '\n', ' ']);
                    let name = trimmed.strip_suffix(';').ok_or_else(|| {
                        RtfError::MalformedDocument("RTF file-table name lacks its semicolon terminator".to_string())
                    })?.trim();
                    let mut entry = crate::FileTableEntry::new(
                        id.ok_or_else(|| RtfError::MalformedDocument("RTF file entry lacks fid".to_string()))?,
                        Cow::Owned(name.to_string()),
                    );
                    entry.relative_path_level = relative;
                    entry.operating_system = operating_system;
                    entry.valid_on = valid_on;
                    entry.location = location;
                    entry.validate()?;
                    return Ok(entry);
                },
                Some(Token::OpenBrace) | Some(Token::Binary(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF file entry cannot contain fields, objects, nested destinations, or binary data".to_string(),
                    ));
                },
                Some(Token::Text(text)) => name.push_str(&self.decode_transport_text(text)?),
                Some(Token::Control(ControlWord::Unicode(first))) => {
                    name.push_str(&self.parse_style_unicode(*first, unicode_skip)?);
                    continue;
                },
                Some(Token::Control(ControlWord::UnicodeSkip(value))) => unicode_skip = (*value).max(0),
                Some(Token::Control(control)) if control_symbol_text(control).is_some() => {
                    name.push_str(control_symbol_text(control).unwrap_or_default())
                },
                Some(Token::Control(control)) => match control {
                    ControlWord::FileId(value) => {
                        if !seen.insert("fid") { return Err(RtfError::MalformedDocument("duplicate RTF fid".to_string())); }
                        id = Some(u32::try_from(*value).map_err(|_| RtfError::MalformedDocument("invalid RTF fid".to_string()))?);
                    },
                    ControlWord::FileRelative(value) => {
                        if !seen.insert("frelative") { return Err(RtfError::MalformedDocument("duplicate RTF frelative".to_string())); }
                        relative = Some(u8::try_from(*value).map_err(|_| RtfError::MalformedDocument("invalid RTF frelative".to_string()))?);
                    },
                    ControlWord::FileOperatingSystem(value) => {
                        if !seen.insert("fosnum") { return Err(RtfError::MalformedDocument("duplicate RTF fosnum".to_string())); }
                        operating_system = Some(u8::try_from(*value).map_err(|_| RtfError::MalformedDocument("invalid RTF fosnum".to_string()))?);
                    },
                    ControlWord::FileValidMac => {
                        if !seen.insert("fvalidmac") { return Err(RtfError::MalformedDocument("duplicate RTF fvalidmac".to_string())); }
                        valid_on.mac = true;
                    },
                    ControlWord::FileValidDos => {
                        if !seen.insert("fvaliddos") { return Err(RtfError::MalformedDocument("duplicate RTF fvaliddos".to_string())); }
                        valid_on.dos = true;
                    },
                    ControlWord::FileValidNtfs => {
                        if !seen.insert("fvalidntfs") { return Err(RtfError::MalformedDocument("duplicate RTF fvalidntfs".to_string())); }
                        valid_on.ntfs = true;
                    },
                    ControlWord::FileValidHpfs => {
                        if !seen.insert("fvalidhpfs") { return Err(RtfError::MalformedDocument("duplicate RTF fvalidhpfs".to_string())); }
                        valid_on.hpfs = true;
                    },
                    ControlWord::FileNetwork => {
                        if !seen.insert("location") { return Err(RtfError::MalformedDocument("conflicting RTF file locations".to_string())); }
                        location = crate::FileLocation::Network;
                    },
                    ControlWord::FileNonFileSystem => {
                        if !seen.insert("location") { return Err(RtfError::MalformedDocument("conflicting RTF file locations".to_string())); }
                        location = crate::FileLocation::NonFileSystem;
                    },
                    _ => return Err(RtfError::MalformedDocument("unsupported control in RTF file entry".to_string())),
                },
                None => return Err(RtfError::UnexpectedEof),
            }
            self.pos += 1;
            if name.len() > crate::file_table::MAX_FILE_NAME_BYTES {
                return Err(RtfError::MalformedDocument("RTF file-table name exceeds the safety limit".to_string()));
            }
        }
        Err(RtfError::UnexpectedEof)
    }

    /// Parse font table.
    fn parse_font_table(&mut self) -> RtfResult<()> {
        self.pos += 1; // Skip \fonttbl

        while self.pos < self.tokens.len() {
            match &self.tokens[self.pos] {
                Token::CloseBrace => {
                    self.font_table.borrow().validate()?;
                    return Ok(());
                },
                Token::OpenBrace => {
                    self.parse_font_entry()?;
                },
                _ => {
                    self.pos += 1;
                },
            }
        }

        Ok(())
    }

    /// Parse a single font table entry.
    fn parse_font_entry(&mut self) -> RtfResult<()> {
        self.pos += 1; // Skip {

        let mut font_num = None;
        let mut font_family = FontFamily::Nil;
        let mut charset = None;
        let mut pitch = crate::FontPitch::Default;
        let mut code_page = None;
        let mut alternate_name = None;
        let mut non_tagged_name = None;
        let mut panose = None;
        let mut name = String::new();
        let mut unicode_skip = self.current_state()?.unicode_skip.max(0);
        let mut seen = std::collections::HashSet::new();

        while self.pos < self.tokens.len() {
            match &self.tokens[self.pos] {
                Token::CloseBrace => {
                    self.pos += 1;
                    break;
                },
                Token::OpenBrace
                    if matches!(
                        self.tokens.get(self.pos + 1..self.pos + 3),
                        Some([Token::Control(ControlWord::IgnorableDestination), Token::Control(ControlWord::FontAlternateName)])
                    ) => {
                        if !seen.insert("falt") { return Err(RtfError::MalformedDocument("duplicate RTF falt destination".to_string())); }
                        alternate_name = Some(self.parse_font_name_destination(ControlWord::FontAlternateName)?);
                    },
                Token::OpenBrace
                    if matches!(
                        self.tokens.get(self.pos + 1..self.pos + 3),
                        Some([Token::Control(ControlWord::IgnorableDestination), Token::Control(ControlWord::FontNonTaggedName)])
                    ) => {
                        if !seen.insert("fname") { return Err(RtfError::MalformedDocument("duplicate RTF fname destination".to_string())); }
                        non_tagged_name = Some(self.parse_font_name_destination(ControlWord::FontNonTaggedName)?);
                    },
                Token::OpenBrace
                    if matches!(
                        self.tokens.get(self.pos + 1..self.pos + 3),
                        Some([Token::Control(ControlWord::IgnorableDestination), Token::Control(ControlWord::FontPanose)])
                    ) => {
                        if !seen.insert("panose") { return Err(RtfError::MalformedDocument("duplicate RTF panose destination".to_string())); }
                        panose = Some(self.parse_font_panose_destination()?);
                    },
                Token::OpenBrace => {
                    if matches!(
                        self.tokens.get(self.pos + 1),
                        Some(Token::Control(ControlWord::Field | ControlWord::Object))
                    ) {
                        return Err(RtfError::MalformedDocument(
                            "RTF font entry cannot contain fields or objects".to_string(),
                        ));
                    }
                    self.skip_group()?;
                },
                Token::Control(ControlWord::FontNumber(n)) => {
                    if !seen.insert("font-number") { return Err(RtfError::MalformedDocument("duplicate RTF font ID".to_string())); }
                    font_num = Some(FontRef::try_from(*n).map_err(|_| RtfError::MalformedDocument("invalid RTF font ID".to_string()))?);
                    self.pos += 1;
                },
                Token::Control(ControlWord::FontFamily(family)) => {
                    if !seen.insert("family") { return Err(RtfError::MalformedDocument("duplicate RTF font family".to_string())); }
                    font_family = match *family {
                        "roman" => FontFamily::Roman,
                        "swiss" => FontFamily::Swiss,
                        "modern" => FontFamily::Modern,
                        "script" => FontFamily::Script,
                        "decor" => FontFamily::Decor,
                        "tech" => FontFamily::Tech,
                        _ => FontFamily::Nil,
                    };
                    self.pos += 1;
                },
                Token::Control(ControlWord::FontCharset(cs)) => {
                    if !seen.insert("charset") { return Err(RtfError::MalformedDocument("duplicate RTF font charset".to_string())); }
                    charset = Some(u8::try_from(*cs).map_err(|_| RtfError::MalformedDocument("invalid RTF font charset".to_string()))?);
                    self.pos += 1;
                },
                Token::Control(ControlWord::FontPitch(value)) => {
                    if !seen.insert("pitch") { return Err(RtfError::MalformedDocument("duplicate RTF font pitch".to_string())); }
                    pitch = match *value {
                        0 => crate::FontPitch::Default,
                        1 => crate::FontPitch::Fixed,
                        2 => crate::FontPitch::Variable,
                        _ => return Err(RtfError::MalformedDocument("invalid RTF font pitch".to_string())),
                    };
                    self.pos += 1;
                },
                Token::Control(ControlWord::FontCodePage(value)) => {
                    if !seen.insert("code-page") { return Err(RtfError::MalformedDocument("duplicate RTF font code page".to_string())); }
                    code_page = Some(u16::try_from(*value).map_err(|_| RtfError::MalformedDocument("invalid RTF font code page".to_string()))?);
                    self.pos += 1;
                },
                Token::Text(text) => {
                    name.push_str(&self.decode_transport_text(text)?);
                    self.pos += 1;
                },
                Token::Control(ControlWord::Unicode(first)) => {
                    name.push_str(&self.parse_style_unicode(*first, unicode_skip)?);
                },
                Token::Control(ControlWord::UnicodeSkip(value)) => {
                    unicode_skip = (*value).max(0);
                    self.pos += 1;
                },
                Token::Control(control) if control_symbol_text(control).is_some() => {
                    name.push_str(control_symbol_text(control).unwrap_or_default());
                    self.pos += 1;
                },
                _ => {
                    self.pos += 1;
                },
            }
            if name.len() > 4_096 {
                return Err(RtfError::MalformedDocument("RTF font name exceeds the safety limit".to_string()));
            }
        }

        let font_num = font_num.ok_or_else(|| RtfError::MalformedDocument("RTF font entry lacks an ID".to_string()))?;
        let name = name.trim().strip_suffix(';').unwrap_or(name.trim()).trim();
        let mut font = Font::new(Cow::Owned(name.to_string()), font_family, charset.unwrap_or(0));
        font.alternate_name = alternate_name.map(Cow::Owned);
        font.non_tagged_name = non_tagged_name.map(Cow::Owned);
        font.panose = panose;
        font.pitch = pitch;
        font.code_page = code_page;
        font.validate()?;
        if let Some(existing) = self.font_table.borrow().get(font_num) {
            if existing == &font {
                return Ok(());
            }
            return Err(RtfError::MalformedDocument(
                "conflicting duplicate RTF font ID".to_string(),
            ));
        }
        self.font_table.borrow_mut().insert(font_num, font);

        Ok(())
    }

    fn parse_font_name_destination(&mut self, expected: ControlWord<'a>) -> RtfResult<String> {
        if !matches!(self.tokens.get(self.pos), Some(Token::OpenBrace))
            || !matches!(self.tokens.get(self.pos + 1), Some(Token::Control(ControlWord::IgnorableDestination)))
            || self.tokens.get(self.pos + 2) != Some(&Token::Control(expected))
        {
            return Err(RtfError::MalformedDocument("invalid RTF font-name destination".to_string()));
        }
        self.pos += 3;
        let mut value = String::new();
        let mut unicode_skip = self.current_state()?.unicode_skip.max(0);
        while self.pos < self.tokens.len() {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    let value = value.trim().strip_suffix(';').unwrap_or(value.trim()).trim().to_string();
                    if value.is_empty() || value.len() > 4_096 {
                        return Err(RtfError::MalformedDocument("invalid or oversized RTF alternate font name".to_string()));
                    }
                    return Ok(value);
                },
                Some(Token::Text(text)) => value.push_str(&self.decode_transport_text(text)?),
                Some(Token::Control(ControlWord::Unicode(first))) => {
                    value.push_str(&self.parse_style_unicode(*first, unicode_skip)?);
                    continue;
                },
                Some(Token::Control(ControlWord::UnicodeSkip(count))) => unicode_skip = (*count).max(0),
                Some(Token::Control(control)) if control_symbol_text(control).is_some() => value.push_str(control_symbol_text(control).unwrap_or_default()),
                Some(Token::OpenBrace) | Some(Token::Control(_)) | Some(Token::Binary(_)) => {
                    return Err(RtfError::MalformedDocument("RTF font-name destination contains non-text content".to_string()));
                },
                None => return Err(RtfError::UnexpectedEof),
            }
            self.pos += 1;
            if value.len() > 4_096 {
                return Err(RtfError::MalformedDocument("RTF alternate font name exceeds the safety limit".to_string()));
            }
        }
        Err(RtfError::UnexpectedEof)
    }

    fn parse_font_panose_destination(&mut self) -> RtfResult<[u8; 10]> {
        self.pos += 3; // opening brace, ignorable marker, panose
        let mut digits = String::new();
        while self.pos < self.tokens.len() {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    let compact: String = digits.chars().filter(|ch| !ch.is_whitespace()).collect();
                    if compact.len() != 20 || !compact.bytes().all(|byte| byte.is_ascii_hexdigit()) {
                        return Err(RtfError::MalformedDocument("RTF panose must contain exactly ten hexadecimal bytes".to_string()));
                    }
                    let mut panose = [0u8; 10];
                    for (index, byte) in panose.iter_mut().enumerate() {
                        *byte = u8::from_str_radix(&compact[index * 2..index * 2 + 2], 16)
                            .map_err(|_| RtfError::MalformedDocument("invalid RTF panose payload".to_string()))?;
                    }
                    return Ok(panose);
                },
                Some(Token::Text(text)) => digits.push_str(text),
                Some(Token::OpenBrace) | Some(Token::Control(_)) | Some(Token::Binary(_)) => {
                    return Err(RtfError::MalformedDocument("RTF panose contains non-hexadecimal content".to_string()));
                },
                None => return Err(RtfError::UnexpectedEof),
            }
            self.pos += 1;
            if digits.len() > 64 {
                return Err(RtfError::MalformedDocument("RTF panose payload exceeds the safety limit".to_string()));
            }
        }
        Err(RtfError::UnexpectedEof)
    }

    /// Parse color table.
    fn parse_color_table(&mut self) -> RtfResult<()> {
        self.pos += 1; // Skip \colortbl

        let mut current_red = 0;
        let mut current_green = 0;
        let mut current_blue = 0;

        while self.pos < self.tokens.len() {
            match &self.tokens[self.pos] {
                Token::CloseBrace => {
                    // Add final color if any
                    let color = Color::new(current_red, current_green, current_blue);
                    self.color_table.borrow_mut().add(color);
                    return Ok(());
                },
                Token::Control(ControlWord::Red(r)) => {
                    current_red = (*r).clamp(0, 255) as u8;
                    self.pos += 1;
                },
                Token::Control(ControlWord::Green(g)) => {
                    current_green = (*g).clamp(0, 255) as u8;
                    self.pos += 1;
                },
                Token::Control(ControlWord::Blue(b)) => {
                    current_blue = (*b).clamp(0, 255) as u8;
                    self.pos += 1;
                },
                Token::Text(text) if text.trim() == ";" => {
                    // Color separator - add current color
                    let color = Color::new(current_red, current_green, current_blue);
                    self.color_table.borrow_mut().add(color);
                    current_red = 0;
                    current_green = 0;
                    current_blue = 0;
                    self.pos += 1;
                },
                _ => {
                    self.pos += 1;
                },
            }
        }

        Ok(())
    }

    /// Parse the standard RTF `info` destination.
    fn parse_info(&mut self) -> RtfResult<()> {
        if self.saw_info_group {
            return Err(RtfError::MalformedDocument(
                "RTF contains multiple info groups".to_string(),
            ));
        }
        self.saw_info_group = true;
        self.pos += 1; // `info`
        while self.pos < self.tokens.len() {
            match self.tokens.get(self.pos) {
                Some(Token::OpenBrace) => {
                    self.pos += 1;
                    let control = self.tokens.get(self.pos).cloned();
                    match control {
                        Some(Token::Control(ControlWord::Title)) => {
                            self.parse_info_text(InfoTextField::Title)?;
                        },
                        Some(Token::Control(ControlWord::Subject)) => {
                            self.parse_info_text(InfoTextField::Subject)?;
                        },
                        Some(Token::Control(ControlWord::Author)) => {
                            self.parse_info_text(InfoTextField::Author)?;
                        },
                        Some(Token::Control(ControlWord::Manager)) => {
                            self.parse_info_text(InfoTextField::Manager)?;
                        },
                        Some(Token::Control(ControlWord::Company)) => {
                            self.parse_info_text(InfoTextField::Company)?;
                        },
                        Some(Token::Control(ControlWord::Operator)) => {
                            self.parse_info_text(InfoTextField::Operator)?;
                        },
                        Some(Token::Control(ControlWord::Category)) => {
                            self.parse_info_text(InfoTextField::Category)?;
                        },
                        Some(Token::Control(ControlWord::Keywords)) => {
                            self.parse_info_text(InfoTextField::Keywords)?;
                        },
                        Some(Token::Control(ControlWord::Comment)) => {
                            self.parse_info_text(InfoTextField::Comment)?;
                        },
                        Some(Token::Control(ControlWord::DocComment)) => {
                            self.parse_info_text(InfoTextField::DocumentComment)?;
                        },
                        Some(Token::Control(ControlWord::HyperlinkBase)) => {
                            self.parse_info_text(InfoTextField::HyperlinkBase)?;
                        },
                        Some(Token::Control(ControlWord::CreationTime)) => {
                            self.parse_info_time(InfoTimeField::Creation)?;
                        },
                        Some(Token::Control(ControlWord::RevisionTime)) => {
                            self.parse_info_time(InfoTimeField::Revision)?;
                        },
                        Some(Token::Control(ControlWord::PrintTime)) => {
                            self.parse_info_time(InfoTimeField::Print)?;
                        },
                        Some(Token::Control(ControlWord::BackupTime)) => {
                            self.parse_info_time(InfoTimeField::Backup)?;
                        },
                        Some(Token::Control(ControlWord::IgnorableDestination))
                            if matches!(
                                self.tokens.get(self.pos + 1),
                                Some(Token::Control(ControlWord::Password))
                            ) =>
                        {
                            self.parse_info_password()?;
                        },
                        Some(Token::Control(ControlWord::Password)) => {
                            return Err(RtfError::MalformedDocument(
                                "RTF password hash destination must be starred".to_string(),
                            ));
                        },
                        _ => self.skip_open_info_group()?,
                    }
                },
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    return Ok(());
                },
                Some(Token::Control(control)) => {
                    match control {
                        ControlWord::InfoVersion(value) => Self::set_info_number(&mut self.info.version, *value, "version")?,
                        ControlWord::InfoRevision(value) => Self::set_info_number(&mut self.info.revision, *value, "vern")?,
                        ControlWord::EditingTime(value) => Self::set_info_number(&mut self.info.editing_time, *value, "edmins")?,
                        ControlWord::NumberOfPages(value) => Self::set_info_number(&mut self.info.pages, *value, "nofpages")?,
                        ControlWord::NumberOfWords(value) => Self::set_info_number(&mut self.info.words, *value, "nofwords")?,
                        ControlWord::NumberOfCharacters(value) => {
                            Self::set_info_number(&mut self.info.characters, *value, "nofchars")?;
                        },
                        ControlWord::NumberOfCharactersWithSpaces(value) => {
                            Self::set_info_number(&mut self.info.characters_with_spaces, *value, "nofcharsws")?;
                        },
                        ControlWord::DocumentId(value) => Self::set_info_number(&mut self.info.id, *value, "id")?,
                        _ => {},
                    }
                    self.pos += 1;
                },
                Some(Token::Text(text)) if self.decode_transport_text(text)?.trim().is_empty() => self.pos += 1,
                Some(Token::Text(_) | Token::Binary(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF info group contains active text or binary data".to_string(),
                    ));
                },
                None => break,
            }
        }
        Err(RtfError::UnexpectedEof)
    }

    fn parse_bookmark_destination(&mut self) -> RtfResult<()> {
        self.pos += 1; // ignorable-destination marker
        let is_start = match self.tokens.get(self.pos) {
            Some(Token::Control(ControlWord::BookmarkStart)) => true,
            Some(Token::Control(ControlWord::BookmarkEnd)) => false,
            _ => {
                return Err(RtfError::MalformedDocument(
                    "invalid bookmark destination".into(),
                ));
            },
        };
        self.pos += 1;

        let mut name = String::new();
        let mut first_column = None;
        let mut last_column = None;
        let mut is_public = false;
        let mut depth = 1usize;
        let mut unicode_skip = self.current_state()?.unicode_skip.max(0) as usize;
        let mut fallback_skip = 0usize;
        while self.pos < self.tokens.len() && depth > 0 {
            match self.tokens.get(self.pos) {
                Some(Token::OpenBrace) => depth += 1,
                Some(Token::CloseBrace) => depth -= 1,
                Some(Token::Text(text)) => {
                    let skipped = fallback_skip.min(text.chars().count());
                    fallback_skip -= skipped;
                    let remainder: String = text.chars().skip(skipped).collect();
                    name.push_str(&self.decode_transport_text(&remainder)?);
                },
                Some(Token::Control(ControlWord::BookmarkFirstColumn(value))) => {
                    first_column = Some(*value);
                },
                Some(Token::Control(ControlWord::BookmarkLastColumn(value))) => {
                    last_column = Some(*value);
                },
                Some(Token::Control(ControlWord::BookmarkPublic)) => is_public = true,
                Some(Token::Control(ControlWord::Unicode(_))) => {
                    let mut utf16 = SmallVec::<[u16; 4]>::new();
                    while let Some(Token::Control(ControlWord::Unicode(code))) =
                        self.tokens.get(self.pos)
                    {
                        utf16.push(*code as u16);
                        self.pos += 1;
                    }
                    name.push_str(&String::from_utf16(&utf16).map_err(|error| {
                        RtfError::InvalidUnicode(format!("invalid Unicode bookmark name: {error}"))
                    })?);
                    fallback_skip = unicode_skip.saturating_mul(utf16.len());
                    continue;
                },
                Some(Token::Control(ControlWord::UnicodeSkip(value))) => {
                    unicode_skip = (*value).max(0) as usize;
                },
                Some(Token::Control(control)) if control_symbol_text(control).is_some() => {
                    name.push_str(control_symbol_text(control).unwrap_or_default());
                },
                _ => {},
            }
            self.pos += 1;
            if name.len() > MAX_BOOKMARK_NAME_BYTES {
                return Err(RtfError::MalformedDocument(
                    "RTF bookmark name exceeds the safety limit".to_string(),
                ));
            }
        }
        if depth != 0 {
            return Err(RtfError::UnexpectedEof);
        }
        let name = name.trim_end_matches(['\r', '\n']).to_string();
        if name.is_empty() {
            return Ok(());
        }

        if is_start {
            if self.next_bookmark_order >= MAX_BOOKMARKS {
                return Err(RtfError::MalformedDocument(
                    "RTF bookmark count exceeds the safety limit".to_string(),
                ));
            }
            let bookmark = OpenBookmark {
                name: name.clone(),
                position: self.body_text_len,
                first_column,
                last_column,
                is_public,
                order: self.next_bookmark_order,
            };
            self.next_bookmark_order += 1;
            self.open_bookmarks.entry(name).or_default().push(bookmark);
        } else if let Some(open) = self.open_bookmarks.get_mut(&name).and_then(Vec::pop) {
            self.bookmark_spans.push(BookmarkSpan {
                bookmark: open,
                end: self.body_text_len,
            });
        }
        Ok(())
    }

    fn finalize_bookmarks(&mut self) -> RtfResult<()> {
        for bookmarks in self.open_bookmarks.values_mut() {
            for bookmark in bookmarks.drain(..) {
                self.bookmark_spans.push(BookmarkSpan {
                    bookmark,
                    end: self.body_text_len,
                });
            }
        }
        self.bookmark_spans
            .sort_unstable_by_key(|span| span.bookmark.order);
        if self.bookmark_spans.is_empty() {
            return Ok(());
        }

        let mut body = String::with_capacity(self.body_text_len);
        for block in &self.blocks {
            body.push_str(block.text.as_ref());
        }
        for span in self.bookmark_spans.drain(..) {
            let content = body.get(span.bookmark.position..span.end).ok_or_else(|| {
                RtfError::MalformedDocument("bookmark does not align to body text".to_string())
            })?;
            self.bookmarks.add(super::bookmark::Bookmark {
                name: Cow::Owned(span.bookmark.name),
                position: span.bookmark.position,
                content: Cow::Owned(content.to_string()),
                first_column: span.bookmark.first_column,
                last_column: span.bookmark.last_column,
                is_public: span.bookmark.is_public,
            });
        }
        Ok(())
    }

    fn parse_ignorable_text_destination(&mut self) -> RtfResult<String> {
        self.pos += 2; // ignorable marker and destination control word
        let mut value = String::new();
        let mut depth = 1usize;
        let mut unicode_skip = self.current_state()?.unicode_skip.max(0) as usize;
        let mut fallback_skip = 0usize;
        while self.pos < self.tokens.len() && depth > 0 {
            match self.tokens.get(self.pos) {
                Some(Token::OpenBrace) => depth += 1,
                Some(Token::CloseBrace) => depth -= 1,
                Some(Token::Text(text)) => {
                    let skipped = fallback_skip.min(text.chars().count());
                    fallback_skip -= skipped;
                    let remainder: String = text.chars().skip(skipped).collect();
                    value.push_str(&self.decode_transport_text(&remainder)?);
                },
                Some(Token::Control(ControlWord::Unicode(_))) => {
                    let mut utf16 = SmallVec::<[u16; 4]>::new();
                    while let Some(Token::Control(ControlWord::Unicode(code))) =
                        self.tokens.get(self.pos)
                    {
                        utf16.push(*code as u16);
                        self.pos += 1;
                    }
                    value.push_str(&String::from_utf16(&utf16).map_err(|error| {
                        RtfError::InvalidUnicode(format!(
                            "invalid Unicode annotation metadata: {error}"
                        ))
                    })?);
                    fallback_skip = unicode_skip.saturating_mul(utf16.len());
                    continue;
                },
                Some(Token::Control(ControlWord::UnicodeSkip(count))) => {
                    unicode_skip = (*count).max(0) as usize;
                },
                Some(Token::Control(control)) if control_symbol_text(control).is_some() => {
                    value.push_str(control_symbol_text(control).unwrap_or_default());
                },
                _ => {},
            }
            self.pos += 1;
            if value.len() > MAX_BOOKMARK_NAME_BYTES {
                return Err(RtfError::MalformedDocument(
                    "RTF annotation destination exceeds the safety limit".to_string(),
                ));
            }
        }
        if depth != 0 {
            return Err(RtfError::UnexpectedEof);
        }
        Ok(value.trim_end_matches(['\r', '\n']).to_string())
    }

    fn parse_annotation_range_marker(&mut self, is_start: bool) -> RtfResult<()> {
        let value = self.parse_ignorable_text_destination()?;
        let reference = value.trim().parse::<i32>().map_err(|_| {
            RtfError::MalformedDocument(
                "RTF annotation range reference must be a signed integer".to_string(),
            )
        })?;
        if !self.annotation_ranges.contains_key(&reference)
            && self.annotation_ranges.len() >= MAX_ANNOTATIONS
        {
            return Err(RtfError::MalformedDocument(
                "RTF annotation range count exceeds the safety limit".to_string(),
            ));
        }
        if is_start {
            if self.annotation_ranges.contains_key(&reference) {
                return Err(RtfError::MalformedDocument(
                    "duplicate RTF annotation range start".to_string(),
                ));
            }
            self.annotation_ranges
                .insert(reference, (self.body_text_len, None));
        } else {
            let range = self.annotation_ranges.get_mut(&reference).ok_or_else(|| {
                RtfError::MalformedDocument(
                    "RTF annotation range end has no matching start".to_string(),
                )
            })?;
            if range.1.is_some() {
                return Err(RtfError::MalformedDocument(
                    "duplicate RTF annotation range end".to_string(),
                ));
            }
            range.1 = Some(self.body_text_len);
        }
        Ok(())
    }

    fn parse_annotation_destination(&mut self) -> RtfResult<()> {
        if self.annotations.len() >= MAX_ANNOTATIONS {
            return Err(RtfError::MalformedDocument(
                "RTF annotation count exceeds the safety limit".to_string(),
            ));
        }
        if !self.pending_annotation_mark {
            return Err(RtfError::MalformedDocument(
                "RTF annotation destination requires a preceding chatn marker".to_string(),
            ));
        }
        self.pending_annotation_mark = false;
        self.pos += 2; // ignorable marker and annotation destination
        let mut reference = None;
        let mut date = None;
        let mut parent_id = None;
        let mut icon = None;
        let mut time = None;
        let mut text = String::new();
        let mut depth = 1usize;
        let mut fallback_skip = 0usize;
        while self.pos < self.tokens.len() && depth > 0 {
            match self.tokens.get(self.pos) {
                Some(Token::OpenBrace) => {
                    let nested =
                        match (self.tokens.get(self.pos + 1), self.tokens.get(self.pos + 2)) {
                            (
                                Some(Token::Control(ControlWord::IgnorableDestination)),
                                Some(Token::Control(control)),
                            ) => Some(*control),
                            _ => None,
                        };
                    match nested {
                        Some(ControlWord::AnnotationReference) => {
                            if reference.is_some() {
                                return Err(RtfError::MalformedDocument(
                                    "duplicate RTF annotation reference".to_string(),
                                ));
                            }
                            let value = self.parse_nested_annotation_value()?;
                            reference = Some(value.trim().parse::<i32>().map_err(|_| {
                                RtfError::MalformedDocument(
                                    "RTF annotation reference must be a signed integer".to_string(),
                                )
                            })?);
                        },
                        Some(ControlWord::AnnotationDate) => {
                            if date.is_some() {
                                return Err(RtfError::MalformedDocument(
                                    "duplicate RTF annotation date".to_string(),
                                ));
                            }
                            date = Some(self.parse_nested_annotation_value()?);
                        },
                        Some(ControlWord::AnnotationParent) => {
                            if parent_id.is_some() {
                                return Err(RtfError::MalformedDocument(
                                    "duplicate RTF annotation parent".to_string(),
                                ));
                            }
                            parent_id = Some(self.parse_nested_annotation_value()?);
                        },
                        Some(ControlWord::AnnotationIcon) => {
                            if icon.is_some() {
                                return Err(RtfError::MalformedDocument(
                                    "duplicate RTF annotation icon".to_string(),
                                ));
                            }
                            icon = Some(self.parse_nested_annotation_value()?);
                        },
                        Some(ControlWord::AnnotationTime) => {
                            if time.is_some() {
                                return Err(RtfError::MalformedDocument(
                                    "duplicate RTF annotation time".to_string(),
                                ));
                            }
                            time = Some(self.parse_nested_annotation_value()?);
                        },
                        Some(control) if Self::forbidden_annotation_control(&control) => {
                            return Err(RtfError::MalformedDocument(
                                "RTF annotation body cannot contain active data".to_string(),
                            ));
                        },
                        Some(_) => {
                            self.skip_group()?;
                        },
                        _ => {
                            depth += 1;
                            self.pos += 1;
                        },
                    }
                    continue;
                },
                Some(Token::CloseBrace) => depth -= 1,
                Some(Token::Text(value)) => {
                    let skipped = fallback_skip.min(value.chars().count());
                    fallback_skip -= skipped;
                    let remainder: String = value.chars().skip(skipped).collect();
                    text.push_str(&self.decode_transport_text(&remainder)?);
                },
                Some(Token::Control(ControlWord::Unicode(_))) => {
                    let code = match self.tokens.get(self.pos) {
                        Some(Token::Control(ControlWord::Unicode(code))) => *code,
                        _ => unreachable!(),
                    };
                    text.push_str(&self.parse_navigation_unicode_sequence(code)?);
                    fallback_skip = 0;
                    continue;
                },
                Some(Token::Control(ControlWord::UnicodeSkip(count))) => {
                    self.current_state_mut()?.unicode_skip = (*count).max(0);
                },
                Some(Token::Control(ControlWord::Par | ControlWord::Line)) => text.push('\n'),
                Some(Token::Control(ControlWord::Tab)) => text.push('\t'),
                Some(Token::Control(control)) if control_symbol_text(control).is_some() => {
                    text.push_str(control_symbol_text(control).unwrap_or_default());
                },
                Some(Token::Control(control)) if Self::forbidden_annotation_control(control) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF annotation body cannot contain active data".to_string(),
                    ));
                },
                Some(Token::Binary(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF annotation body cannot contain binary data".to_string(),
                    ));
                },
                _ => {},
            }
            self.pos += 1;
            if text.len() > MAX_ANNOTATION_TEXT_BYTES {
                return Err(RtfError::MalformedDocument(
                    "RTF annotation text exceeds the safety limit".to_string(),
                ));
            }
        }
        if depth != 0 {
            return Err(RtfError::UnexpectedEof);
        }

        let has_reference = reference.is_some();
        let id = reference.unwrap_or(0);
        if has_reference
            && self
                .annotations
                .iter()
                .any(|annotation| annotation.has_reference && annotation.id == id)
        {
            return Err(RtfError::MalformedDocument(
                "duplicate RTF annotation reference".to_string(),
            ));
        }
        let (position, range_end) = match self.annotation_ranges.remove(&id) {
            Some((start, Some(end))) if start <= end => (start, end),
            Some((_start, None)) => {
                return Err(RtfError::MalformedDocument(
                    "RTF annotation range has no matching end".to_string(),
                ));
            },
            Some(_) => {
                return Err(RtfError::MalformedDocument(
                    "RTF annotation range end precedes its start".to_string(),
                ));
            },
            None => (self.body_text_len, self.body_text_len),
        };
        let annotation = super::annotation::Annotation {
            annotation_type: super::annotation::AnnotationType::Comment,
            id,
            has_reference,
            author: Cow::Owned(std::mem::take(&mut self.pending_annotation_author)),
            initials: Cow::Owned(std::mem::take(&mut self.pending_annotation_initials)),
            date: date.map(Cow::Owned),
            text: Cow::Owned(text.trim_end_matches(['\r', '\n']).to_string()),
            position,
            range_end,
            parent_id: parent_id.map(Cow::Owned),
            icon: icon.map(Cow::Owned),
            time: time.map(Cow::Owned),
        };
        self.pending_annotation_author_seen = false;
        self.pending_annotation_initials_seen = false;
        annotation.validate()?;
        self.annotations.push(annotation);
        Ok(())
    }

    fn parse_nested_annotation_value(&mut self) -> RtfResult<String> {
        self.pos += 3; // opening brace, ignorable marker, destination
        let mut value = String::new();
        let mut depth = 1usize;
        while self.pos < self.tokens.len() && depth > 0 {
            match self.tokens.get(self.pos) {
                Some(Token::OpenBrace) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF annotation metadata cannot contain nested groups".to_string(),
                    ));
                },
                Some(Token::CloseBrace) => depth -= 1,
                Some(Token::Text(text)) => value.push_str(&self.decode_transport_text(text)?),
                Some(Token::Control(ControlWord::Unicode(code))) => {
                    value.push_str(&self.parse_navigation_unicode_sequence(*code)?);
                    continue;
                },
                Some(Token::Control(ControlWord::UnicodeSkip(count))) => {
                    self.current_state_mut()?.unicode_skip = (*count).max(0);
                },
                Some(Token::Control(control)) if control_symbol_text(control).is_some() => {
                    value.push_str(control_symbol_text(control).unwrap_or_default());
                },
                Some(Token::Binary(_)) | Some(Token::Control(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF annotation metadata contains active or invalid controls".to_string(),
                    ));
                },
                None => break,
            }
            self.pos += 1;
            if value.len() > MAX_BOOKMARK_NAME_BYTES {
                return Err(RtfError::MalformedDocument(
                    "RTF annotation metadata exceeds the safety limit".to_string(),
                ));
            }
        }
        if depth != 0 {
            return Err(RtfError::UnexpectedEof);
        }
        Ok(value.trim_end_matches(['\r', '\n']).to_string())
    }

    fn forbidden_annotation_control(control: &ControlWord<'_>) -> bool {
        matches!(
            control,
            ControlWord::Field
                | ControlWord::FieldInstruction
                | ControlWord::FieldResult
                | ControlWord::Object
                | ControlWord::Result
                | ControlWord::Picture
                | ControlWord::Shape
                | ControlWord::ShapeGroup
                | ControlWord::DocumentVariable
                | ControlWord::UserProperties
                | ControlWord::Annotation
                | ControlWord::Footnote
                | ControlWord::Endnote
        )
    }

    fn finalize_annotations(&self) -> RtfResult<()> {
        if !self.annotation_ranges.is_empty() {
            return Err(RtfError::MalformedDocument(
                "RTF document contains an orphan annotation range".to_string(),
            ));
        }
        if self.pending_annotation_author_seen
            || self.pending_annotation_initials_seen
            || self.pending_annotation_mark
        {
            return Err(RtfError::MalformedDocument(
                "RTF document contains orphan annotation metadata".to_string(),
            ));
        }
        Ok(())
    }

    fn parse_info_text(&mut self, field: InfoTextField) -> RtfResult<()> {
        let duplicate = match field {
            InfoTextField::Title => self.info.title.is_some(),
            InfoTextField::Subject => self.info.subject.is_some(),
            InfoTextField::Author => self.info.author.is_some(),
            InfoTextField::Manager => self.info.manager.is_some(),
            InfoTextField::Company => self.info.company.is_some(),
            InfoTextField::Operator => self.info.operator.is_some(),
            InfoTextField::Category => self.info.category.is_some(),
            InfoTextField::Keywords => self.info.keywords.is_some(),
            InfoTextField::Comment => self.info.comment.is_some(),
            InfoTextField::DocumentComment => self.info.document_comment.is_some(),
            InfoTextField::HyperlinkBase => self.info.hyperlink_base.is_some(),
        };
        if duplicate {
            return Err(RtfError::MalformedDocument(
                "RTF info text destination occurs more than once".to_string(),
            ));
        }
        self.pos += 1; // destination control word
        let mut value = String::new();
        let mut depth = 1usize;
        let mut fallback_skip = 0usize;
        while self.pos < self.tokens.len() && depth > 0 {
            match self.tokens.get(self.pos) {
                Some(Token::OpenBrace) => {
                    depth += 1;
                    self.pos += 1;
                },
                Some(Token::CloseBrace) => {
                    depth -= 1;
                    self.pos += 1;
                },
                Some(Token::Text(text)) => {
                    let skipped = fallback_skip.min(text.chars().count());
                    fallback_skip -= skipped;
                    let remainder: String = text.chars().skip(skipped).collect();
                    value.push_str(&self.decode_transport_text(&remainder)?);
                    self.pos += 1;
                },
                Some(Token::Control(ControlWord::Unicode(_))) => {
                    let mut utf16 = SmallVec::<[u16; 4]>::new();
                    while let Some(Token::Control(ControlWord::Unicode(code))) =
                        self.tokens.get(self.pos)
                    {
                        utf16.push(*code as u16);
                        self.pos += 1;
                    }
                    value.push_str(&String::from_utf16(&utf16).map_err(|error| {
                        RtfError::InvalidUnicode(format!("Invalid info Unicode: {error}"))
                    })?);
                    fallback_skip =
                        self.current_state()?.unicode_skip.max(0) as usize * utf16.len();
                },
                Some(Token::Control(ControlWord::UnicodeSkip(count))) => {
                    self.current_state_mut()?.unicode_skip = *count;
                    self.pos += 1;
                },
                Some(Token::Control(control)) if control_symbol_text(control).is_some() => {
                    value.push_str(control_symbol_text(control).unwrap_or_default());
                    self.pos += 1;
                },
                Some(_) => self.pos += 1,
                None => break,
            }
            if value.len() > MAX_INFO_TEXT_BYTES {
                return Err(RtfError::MalformedDocument(
                    "RTF info text exceeds the metadata safety limit".to_string(),
                ));
            }
        }
        if depth != 0 {
            return Err(RtfError::UnexpectedEof);
        }
        let allocated = self.arena.alloc_str(value.trim_end_matches(['\r', '\n']));
        let value = Some(Cow::Borrowed(&*allocated));
        match field {
            InfoTextField::Title => self.info.title = value,
            InfoTextField::Subject => self.info.subject = value,
            InfoTextField::Author => self.info.author = value,
            InfoTextField::Manager => self.info.manager = value,
            InfoTextField::Company => self.info.company = value,
            InfoTextField::Operator => self.info.operator = value,
            InfoTextField::Category => self.info.category = value,
            InfoTextField::Keywords => self.info.keywords = value,
            InfoTextField::Comment => self.info.comment = value,
            InfoTextField::DocumentComment => self.info.document_comment = value,
            InfoTextField::HyperlinkBase => self.info.hyperlink_base = value,
        }
        Ok(())
    }

    fn parse_info_time(&mut self, field: InfoTimeField) -> RtfResult<()> {
        let duplicate = match field {
            InfoTimeField::Creation => self.info.creation_timestamp.is_some(),
            InfoTimeField::Revision => self.info.revision_timestamp.is_some(),
            InfoTimeField::Print => self.info.print_timestamp.is_some(),
            InfoTimeField::Backup => self.info.backup_timestamp.is_some(),
        };
        if duplicate {
            return Err(RtfError::MalformedDocument(
                "RTF info timestamp destination occurs more than once".to_string(),
            ));
        }
        self.pos += 1; // time destination
        let mut timestamp = crate::RtfTimestamp::default();
        let mut depth = 1usize;
        while self.pos < self.tokens.len() && depth > 0 {
            match self.tokens.get(self.pos) {
                Some(Token::OpenBrace) => depth += 1,
                Some(Token::CloseBrace) => depth -= 1,
                Some(Token::Control(control)) => match control {
                    ControlWord::Year(value) => timestamp.year = Some(*value),
                    ControlWord::Month(value) => timestamp.month = Some(*value),
                    ControlWord::Day(value) => timestamp.day = Some(*value),
                    ControlWord::Hour(value) => timestamp.hour = Some(*value),
                    ControlWord::Minute(value) => timestamp.minute = Some(*value),
                    ControlWord::Second(value) => timestamp.second = Some(*value),
                    _ => {},
                },
                _ => {},
            }
            self.pos += 1;
        }
        if depth != 0 {
            return Err(RtfError::UnexpectedEof);
        }
        let serialized = timestamp.legacy_string();
        let allocated = self.arena.alloc_str(&serialized);
        let value = Some(Cow::Borrowed(&*allocated));
        match field {
            InfoTimeField::Creation => { self.info.creation_time = value; self.info.creation_timestamp = Some(timestamp); },
            InfoTimeField::Revision => { self.info.revision_time = value; self.info.revision_timestamp = Some(timestamp); },
            InfoTimeField::Print => { self.info.print_time = value; self.info.print_timestamp = Some(timestamp); },
            InfoTimeField::Backup => { self.info.backup_time = value; self.info.backup_timestamp = Some(timestamp); },
        }
        Ok(())
    }

    fn set_info_number(slot: &mut Option<u32>, value: i32, name: &str) -> RtfResult<()> {
        if slot.is_some() {
            return Err(RtfError::MalformedDocument(format!(
                "RTF info numeric control {name} occurs more than once"
            )));
        }
        *slot = Some(u32::try_from(value).map_err(|_| {
            RtfError::MalformedDocument(format!("RTF info numeric control {name} cannot be negative"))
        })?);
        Ok(())
    }

    fn parse_info_password(&mut self) -> RtfResult<()> {
        if self.info.protection.password_hash.is_some() {
            return Err(RtfError::MalformedDocument(
                "duplicate RTF protection password hash".to_string(),
            ));
        }
        self.pos += 2; // ignorable marker and password destination
        let value = self.parse_inert_text_group_contents(
            crate::info::PROTECTION_PASSWORD_HASH_BYTES,
            "protection password hash",
        )?;
        self.info.protection.password_hash = Some(Cow::Owned(value));
        self.info.protection.validate()
    }

    fn ensure_protection_scope(&self) -> RtfResult<()> {
        if self.states.len() != 2 || self.body_text_len != 0 {
            return Err(RtfError::MalformedDocument(
                "RTF document protection controls must occur in the root header".to_string(),
            ));
        }
        Ok(())
    }

    fn set_protection_flag(
        slot: &mut Option<bool>,
        value: Option<i32>,
        name: &str,
    ) -> RtfResult<()> {
        let value = value.unwrap_or(1);
        Self::set_required_protection_flag(slot, value, name)
    }

    fn set_required_protection_flag(
        slot: &mut Option<bool>,
        value: i32,
        name: &str,
    ) -> RtfResult<()> {
        if slot.is_some() {
            return Err(RtfError::MalformedDocument(format!(
                "duplicate RTF {name} control"
            )));
        }
        *slot = Some(match value {
            0 => false,
            1 => true,
            _ => {
                return Err(RtfError::MalformedDocument(format!(
                    "RTF {name} parameter must be 0 or 1"
                )));
            },
        });
        Ok(())
    }

    fn skip_open_info_group(&mut self) -> RtfResult<()> {
        let mut depth = 1usize;
        while self.pos < self.tokens.len() && depth > 0 {
            match self.tokens.get(self.pos) {
                Some(Token::OpenBrace) => depth += 1,
                Some(Token::CloseBrace) => depth -= 1,
                _ => {},
            }
            self.pos += 1;
        }
        (depth == 0).then_some(()).ok_or(RtfError::UnexpectedEof)
    }

    /// Skip tokens until closing brace.
    fn skip_until_close_brace(&mut self) -> RtfResult<()> {
        let mut depth = 1;

        while self.pos < self.tokens.len() && depth > 0 {
            match &self.tokens[self.pos] {
                Token::OpenBrace => depth += 1,
                Token::CloseBrace => depth -= 1,
                _ => {},
            }
            self.pos += 1;
        }

        (depth == 0).then_some(()).ok_or(RtfError::UnexpectedEof)
    }

    /// Skip an entire group starting from the OpenBrace token.
    fn skip_group(&mut self) -> RtfResult<()> {
        // Must be positioned at OpenBrace
        if !matches!(self.tokens.get(self.pos), Some(Token::OpenBrace)) {
            return Ok(());
        }

        self.pos += 1; // Skip the OpenBrace
        let mut depth = 1;

        while self.pos < self.tokens.len() && depth > 0 {
            match &self.tokens[self.pos] {
                Token::OpenBrace => depth += 1,
                Token::CloseBrace => depth -= 1,
                _ => {},
            }
            self.pos += 1;
        }

        (depth == 0).then_some(()).ok_or(RtfError::UnexpectedEof)
    }

    /// Expect a specific token.
    fn expect_token(&mut self, expected: Token) -> RtfResult<()> {
        if self.pos >= self.tokens.len() {
            return Err(RtfError::UnexpectedEof);
        }

        if self.tokens[self.pos] != expected {
            return Err(RtfError::ParserError(format!(
                "Expected {:?}, found {:?}",
                expected, self.tokens[self.pos]
            )));
        }

        self.pos += 1;
        Ok(())
    }

    /// Get current state (mutable).
    fn current_state_mut(&mut self) -> RtfResult<&mut State> {
        self.states
            .last_mut()
            .ok_or_else(|| RtfError::ParserError("No parser state available".to_string()))
    }

    /// Get current state (immutable).
    fn current_state(&self) -> RtfResult<&State> {
        self.states
            .last()
            .ok_or_else(|| RtfError::ParserError("No parser state available".to_string()))
    }

    /// Parse Unicode character sequence with fallback handling.
    ///
    /// RTF Unicode format: `\uN` where N is a signed 16-bit decimal value
    /// Followed by `\ucN` fallback characters (usually ANSI representation)
    ///
    /// Handles compound Unicode characters (surrogate pairs for emoji, etc.)
    fn parse_unicode_sequence(&mut self, first_code: i32) -> RtfResult<()> {
        let skip_count = self.current_state()?.unicode_skip as usize;

        // Collect all consecutive unicode values (for surrogate pairs)
        let mut unicode_values = SmallVec::<[u16; 4]>::new();

        // Convert signed 16-bit value to unsigned
        unicode_values.push(first_code as u16);
        self.pos += 1;

        // Look ahead for additional Unicode characters (compound characters)
        while self.pos < self.tokens.len() {
            if let Token::Control(ControlWord::Unicode(code)) = &self.tokens[self.pos] {
                unicode_values.push(*code as u16);
                self.pos += 1;
            } else {
                break;
            }
        }

        // Skip fallback characters based on unicode_skip count
        // Fallback chars are for non-Unicode readers (usually hex escapes or plain ASCII)
        let mut fallback_skip = skip_count * unicode_values.len();
        let mut fallback_remainder = None;

        // Handle fallback: skip the next N characters/tokens
        while fallback_skip > 0 && self.pos < self.tokens.len() {
            match &self.tokens[self.pos] {
                Token::Text(text) => {
                    let character_count = text.chars().count();
                    if character_count <= fallback_skip {
                        fallback_skip -= character_count;
                        self.pos += 1;
                    } else {
                        fallback_remainder =
                            Some(text.chars().skip(fallback_skip).collect::<String>());
                        fallback_skip = 0;
                        self.pos += 1;
                    }
                },
                Token::Control(ControlWord::Unicode(_)) => {
                    // Next unicode, don't skip
                    break;
                },
                _ => {
                    // Treat other tokens as single character
                    fallback_skip = fallback_skip.saturating_sub(1);
                    self.pos += 1;
                },
            }
        }

        // Convert Unicode values to UTF-8 string
        let unicode_str = String::from_utf16(&unicode_values)
            .map_err(|e| RtfError::InvalidUnicode(format!("Invalid Unicode sequence: {}", e)))?;

        let state = self.current_state()?.clone();
        if state.destination==Destination::DocumentBody&&(state.in_table||state.table_nesting_level>=2) {
            self.append_table_text(unicode_str.as_bytes(),state.table_nesting_level)?;
            if let Some(remainder) = fallback_remainder {
                self.append_table_text(remainder.as_bytes(),state.table_nesting_level)?;
            }
        } else if state.destination==Destination::DocumentBody {
            // Add the Unicode sequence to the document as its own formatted block.
            let allocated = self.arena.alloc_str(&unicode_str);
            let start = self.body_text_len;
            if state.revision_type == Some(super::annotation::RevisionType::Deletion) {
                self.append_revision_text(&state, allocated, start, start)?;
            } else {
                let block =
                    StyleBlock::new(Cow::Borrowed(allocated), state.formatting, state.paragraph);
                self.body_text_len = self
                    .body_text_len
                    .checked_add(allocated.len())
                    .ok_or_else(|| {
                        RtfError::MalformedDocument("RTF body text length overflow".into())
                    })?;
                self.blocks.push(block);
                self.append_revision_text(&state, allocated, start, self.body_text_len)?;
            }

            // A fallback and subsequent text often share one lexer token. Preserve
            // the portion after the configured fallback character count.
            if let Some(remainder) = fallback_remainder {
                let mut buffer = SmallVec::<[u8; 256]>::new();
                append_transport_bytes(&mut buffer, &remainder)?;
                self.flush_text_buffer(&mut buffer)?;
            }
        }

        Ok(())
    }

    /// Start a table if not already started.
    fn start_table_if_needed(&mut self) {
        if self.current_table.is_none() {
            self.current_table = Some(super::table::Table::new());
        }
        if self.current_row.is_none() {
            self.current_row = Some(super::table::Row::new());
        }
    }

    fn ensure_nested_builder(&mut self,level:u8)->RtfResult<&mut NestedTableBuilder<'a>>{
        if !(2..=crate::MAX_TABLE_NESTING_DEPTH as u8).contains(&level){return Err(RtfError::MalformedDocument("RTF nested-table level is outside 2..=32".to_string()))}
        let index=usize::from(level-2);if self.nested_table_builders.len()<index{return Err(RtfError::MalformedDocument("RTF nested-table level transition skips a parent level".to_string()))}
        if self.nested_table_builders.len()==index{if level==2{self.start_table_if_needed();}self.nested_table_builders.push(NestedTableBuilder::new(level));}
        Ok(&mut self.nested_table_builders[index])
    }

    fn append_table_text(&mut self,text:&[u8],raw_level:u8)->RtfResult<()>{let level=if raw_level>=2{raw_level}else{1};self.drain_nested_to(level)?;if level==1{self.current_cell_text.extend_from_slice(text);}else{self.ensure_nested_builder(level)?.cell_text.extend_from_slice(text);}Ok(())}

    fn drain_nested_to(&mut self,parent_level:u8)->RtfResult<()>{while self.nested_table_builders.last().is_some_and(|builder|builder.level>parent_level){let builder=self.nested_table_builders.pop().unwrap();if !builder.cell_text.is_empty()||!builder.cell_nested.is_empty()||builder.row.cell_count()>0{return Err(RtfError::MalformedDocument("RTF nested-table level ended before nestcell/nestrow".to_string()))}if builder.table.row_count()==0{return Err(RtfError::MalformedDocument("RTF nested table has no completed rows".to_string()))}builder.table.validate_merges().map_err(RtfError::MalformedDocument)?;if self.logical_table_count>=MAX_LOGICAL_TABLES{return Err(RtfError::MalformedDocument("RTF document exceeds 4096 logical tables".to_string()))}self.logical_table_count+=1;let entry=crate::CellNestedTable{text_offset:if parent_level==1{self.current_cell_text.len()}else{self.nested_table_builders.last().map_or(0,|parent|parent.cell_text.len())},table:builder.table};if parent_level==1{self.current_cell_nested.push(entry);}else{let parent=self.nested_table_builders.last_mut().ok_or_else(||RtfError::MalformedDocument("RTF nested table lacks a parent table".to_string()))?;if parent.level!=parent_level{return Err(RtfError::MalformedDocument("RTF nested-table parent level mismatch".to_string()))}parent.cell_nested.push(entry);}}Ok(())}

    fn finalize_nested_cell(&mut self,level:u8)->RtfResult<()>{self.drain_nested_to(level)?;let arena=self.arena;let builder=self.ensure_nested_builder(level)?;if builder.row.cell_count()>=crate::MAX_TABLE_CELLS_PER_ROW{return Err(RtfError::MalformedDocument("RTF table row exceeds 4096 cells".to_string()))}let text=std::str::from_utf8(&builder.cell_text).map_err(|_|RtfError::MalformedDocument("invalid UTF-8 in nested table cell".to_string()))?;let mut cell=crate::Cell::new(Cow::Borrowed(arena.alloc_str(text)));cell.nested_tables_mut().append(&mut builder.cell_nested);builder.row.add_cell(cell);builder.cell_text.clear();Ok(())}

    fn finalize_nested_row(&mut self,level:u8)->RtfResult<()>{self.drain_nested_to(level)?;let state=self.current_state()?.clone();let geometry=resolve_row_geometry(&state)?;let builder=self.ensure_nested_builder(level)?;if !builder.cell_text.is_empty()||!builder.cell_nested.is_empty(){return Err(RtfError::MalformedDocument("RTF nestrow encountered an unterminated nested cell".to_string()))}if builder.row.cell_count()==0{return Err(RtfError::MalformedDocument("RTF nestrow has no nestcell".to_string()))}if !state.cell_boundaries.is_empty()&&state.cell_boundaries.len()!=builder.row.cell_count(){return Err(RtfError::MalformedDocument("RTF nested row cell boundaries do not match nestcell count".to_string()))}for(index,cell)in builder.row.cells_mut().iter_mut().enumerate(){if let Some((padding,spacing))=state.cell_distances.get(index){cell.set_padding(padding.clone());cell.set_spacing(spacing.clone());}if let Some(layout)=state.cell_layouts.get(index){cell.set_layout(*layout);}if let Some(merge)=state.cell_merges.get(index){cell.set_merge(*merge);}cell.set_right_boundary(state.cell_boundaries.get(index).copied());cell.set_preferred_width(state.cell_widths.get(index).copied().flatten());if let Some((borders,shading))=state.cell_decorations.get(index){cell.set_borders(borders.clone());cell.set_shading(*shading);}}builder.row.set_direction(state.table_row_direction);builder.row.set_layout(state.table_row_layout);builder.row.set_padding(state.table_row_padding.clone());builder.row.set_spacing(state.table_row_spacing.clone());builder.row.set_positioning(state.table_row_positioning.clone());builder.row.set_borders(state.table_row_borders.clone());builder.row.set_shading(state.table_row_shading);builder.row.set_geometry(geometry);if builder.table.row_count()>=MAX_LOGICAL_TABLE_ROWS{return Err(RtfError::MalformedDocument("RTF logical table exceeds 65536 rows".to_string()))}if builder.table.rows().first().is_some_and(|first|first.positioning()!=builder.row.positioning()){return Err(RtfError::MalformedDocument("RTF positioned-table properties must be identical for all rows in one logical table".to_string()))}builder.table.add_row(std::mem::take(&mut builder.row));Ok(())}

    /// Finalize the current cell and add it to the current row.
    fn finalize_cell(&mut self, explicit:bool)->RtfResult<()> {
        self.drain_nested_to(1)?;
        if explicit
            || !self.current_cell_text.is_empty()
            || !self.current_cell_nested.is_empty()
        {
            if self.current_row.as_ref().map_or(0,|row|row.cell_count())>=crate::MAX_TABLE_CELLS_PER_ROW{return Err(RtfError::MalformedDocument("RTF table row exceeds 4096 cells".to_string()))}
            // Convert cell text to string
            if let Ok(text_str) = std::str::from_utf8(&self.current_cell_text) {
                let allocated = self.arena.alloc_str(text_str);
                let index=self.current_row.as_ref().map_or(0,|row|row.cell_count());let (padding,spacing)=self.current_state().ok().and_then(|state|state.cell_distances.get(index)).cloned().unwrap_or_default();let layout=self.current_state().ok().and_then(|state|state.cell_layouts.get(index)).copied().unwrap_or_default();let merge=self.current_state().ok().and_then(|state|state.cell_merges.get(index)).copied().unwrap_or_default();let boundary=self.current_state().ok().and_then(|state|state.cell_boundaries.get(index)).copied();let width=self.current_state().ok().and_then(|state|state.cell_widths.get(index)).copied().flatten();let(borders,shading)=self.current_state().ok().and_then(|state|state.cell_decorations.get(index)).cloned().unwrap_or_default();let mut cell = super::table::Cell::with_distances(Cow::Borrowed(allocated),padding,spacing);cell.set_layout(layout);cell.set_merge(merge);cell.set_right_boundary(boundary);cell.set_preferred_width(width);cell.set_borders(borders);cell.set_shading(shading);cell.nested_tables_mut().append(&mut self.current_cell_nested);

                // Add cell to current row
                if let Some(row) = &mut self.current_row {
                    row.add_cell(cell);
                }
            }

        }
        self.current_cell_text.clear();
        Ok(())
    }

    /// Finalize the current row and add it to the current table.
    fn finalize_row(&mut self) -> RtfResult<()> {
        // Finalize any pending cell
        self.finalize_cell(false)?;

        // Add row to table
        if let (Some(table), Some(row)) = (&mut self.current_table, self.current_row.take())
            && row.cell_count() > 0
        {
            if table.row_count() >= MAX_LOGICAL_TABLE_ROWS {
                return Err(RtfError::MalformedDocument("RTF logical table exceeds 65536 rows".to_string()));
            }
            if table.rows().first().is_some_and(|first| first.positioning() != row.positioning()) {
                return Err(RtfError::MalformedDocument("RTF positioned-table properties must be identical for all rows in one logical table".to_string()));
            }
            table.add_row(row);
        }

        // Start a new row for next cells
        self.current_row = Some(super::table::Row::new());
        Ok(())
    }

    /// Finalize the current table and add it to the tables list.
    fn finalize_table(&mut self) -> RtfResult<()> {
        self.drain_nested_to(1)?;
        // Finalize any pending row
        if self.current_row.is_some() {
            self.finalize_row()?;
        }

        // Add table to tables list
        if let Some(table) = self.current_table.take()
            && table.row_count() > 0
        {
            table.validate_merges().map_err(RtfError::MalformedDocument)?;
            if self.logical_table_count >= MAX_LOGICAL_TABLES {
                return Err(RtfError::MalformedDocument("RTF document exceeds 4096 logical tables".to_string()));
            }
            self.logical_table_count+=1;self.tables.push(table);
        }
        Ok(())
    }

    fn finalize_table_before_non_table_body_content(&mut self, meaningful:bool)->RtfResult<bool> {
        if meaningful
            && self.current_table.as_ref().is_some_and(|table|table.row_count()>0)
            && self.current_state().is_ok_and(|state|state.destination==Destination::DocumentBody&&!state.in_table)
        {
            self.finalize_table()?;
            return Ok(true);
        }
        Ok(false)
    }

    /// Parse an `object` destination without activating or updating its content.
    fn parse_object_destination(&mut self) -> RtfResult<super::object::EmbeddedObject<'a>> {
        use super::object::ObjectKind;

        let mut object = super::object::EmbeddedObject::new();
        let mut depth = 0usize;
        self.pos += 1; // consume \object

        while self.pos < self.tokens.len() {
            match &self.tokens[self.pos] {
                Token::OpenBrace
                    if self.nested_control_word() == Some(ControlWord::ObjectClass) =>
                {
                    object.class_name = Cow::Owned(self.parse_object_text_destination()?);
                },
                Token::OpenBrace if self.nested_control_word() == Some(ControlWord::ObjectName) => {
                    object.name = Cow::Owned(self.parse_object_text_destination()?);
                },
                Token::OpenBrace if self.nested_control_word() == Some(ControlWord::ObjectData) => {
                    object.data = self.parse_object_hex_destination()?;
                },
                Token::OpenBrace if self.nested_control_word() == Some(ControlWord::Result) => {
                    let (text, pictures) = self.parse_object_result_destination()?;
                    object.result_text = Cow::Owned(text);
                    object.result_picture_indices = pictures;
                },
                Token::OpenBrace => {
                    depth += 1;
                    self.pos += 1;
                },
                Token::CloseBrace if depth == 0 => {
                    self.pos += 1;
                    return Ok(object);
                },
                Token::CloseBrace => {
                    depth -= 1;
                    self.pos += 1;
                },
                Token::Control(ControlWord::ObjectEmbedded) => {
                    object.kind = ObjectKind::Embedded;
                    self.pos += 1;
                },
                Token::Control(ControlWord::ObjectLink) => {
                    object.kind = ObjectKind::Link;
                    self.pos += 1;
                },
                Token::Control(ControlWord::ObjectAutoLink) => {
                    object.kind = ObjectKind::AutoLink;
                    self.pos += 1;
                },
                Token::Control(ControlWord::ObjectHtml) => {
                    object.kind = ObjectKind::Html;
                    self.pos += 1;
                },
                Token::Control(ControlWord::ObjectWidth(value)) => {
                    object.width = *value;
                    self.pos += 1;
                },
                Token::Control(ControlWord::ObjectHeight(value)) => {
                    object.height = *value;
                    self.pos += 1;
                },
                Token::Control(ControlWord::ObjectLocked(value)) => {
                    object.locked = *value;
                    self.pos += 1;
                },
                Token::Control(ControlWord::ObjectUpdate(value)) => {
                    object.update_requested = *value;
                    self.pos += 1;
                },
                Token::Control(ControlWord::ObjectSetSize(value)) => {
                    object.set_size = *value;
                    self.pos += 1;
                },
                _ => self.pos += 1,
            }
        }
        Err(RtfError::UnexpectedEof)
    }

    fn parse_object_result_destination(&mut self) -> RtfResult<(String, Vec<usize>)> {
        let mut text = String::new();
        let mut picture_indices = Vec::new();
        self.pos += 1; // opening brace
        if matches!(
            self.tokens.get(self.pos),
            Some(Token::Control(ControlWord::IgnorableDestination))
        ) {
            self.pos += 1;
        }
        if !matches!(
            self.tokens.get(self.pos),
            Some(Token::Control(ControlWord::Result))
        ) {
            return Err(RtfError::MalformedDocument(
                "invalid RTF object result destination".to_string(),
            ));
        }
        self.pos += 1;

        while self.pos < self.tokens.len() {
            match &self.tokens[self.pos] {
                Token::CloseBrace => {
                    self.pos += 1;
                    return Ok((text.trim().to_string(), picture_indices));
                },
                Token::OpenBrace if self.nested_control_word() == Some(ControlWord::Picture) => {
                    let first_picture = self.pictures.len();
                    self.parse_group()?;
                    picture_indices.extend(first_picture..self.pictures.len());
                },
                Token::OpenBrace => self.skip_group()?,
                Token::Control(ControlWord::Unicode(code)) => {
                    text.push_str(&self.parse_destination_unicode_sequence(*code)?);
                },
                Token::Control(ControlWord::Par | ControlWord::Line) => {
                    text.push('\n');
                    self.pos += 1;
                },
                Token::Control(ControlWord::Tab) => {
                    text.push('\t');
                    self.pos += 1;
                },
                Token::Control(control) if control_symbol_text(control).is_some() => {
                    text.push_str(control_symbol_text(control).unwrap_or_default());
                    self.pos += 1;
                },
                Token::Text(value) => {
                    text.push_str(&self.decode_transport_text(value)?);
                    self.pos += 1;
                },
                _ => self.pos += 1,
            }
            if text.len() > MAX_OBJECT_TEXT_BYTES {
                return Err(RtfError::MalformedDocument(
                    "RTF object result text exceeds the safety limit".to_string(),
                ));
            }
        }
        Err(RtfError::UnexpectedEof)
    }

    fn parse_object_text_destination(&mut self) -> RtfResult<String> {
        let mut text = String::new();
        let mut depth = 0usize;
        self.pos += 1; // opening brace
        while self.pos < self.tokens.len() {
            match &self.tokens[self.pos] {
                Token::OpenBrace => {
                    depth += 1;
                    self.pos += 1;
                },
                Token::CloseBrace if depth == 0 => {
                    self.pos += 1;
                    return Ok(text.trim().to_string());
                },
                Token::CloseBrace => {
                    depth -= 1;
                    self.pos += 1;
                },
                Token::Control(ControlWord::Unicode(code)) => {
                    text.push_str(&self.parse_destination_unicode_sequence(*code)?);
                },
                Token::Control(control) if control_symbol_text(control).is_some() => {
                    text.push_str(control_symbol_text(control).unwrap_or_default());
                    self.pos += 1;
                },
                Token::Text(value) => {
                    text.push_str(&self.decode_transport_text(value)?);
                    self.pos += 1;
                },
                _ => self.pos += 1,
            }
            if text.len() > MAX_OBJECT_TEXT_BYTES {
                return Err(RtfError::MalformedDocument(
                    "RTF embedded object metadata exceeds the safety limit".to_string(),
                ));
            }
        }
        Err(RtfError::UnexpectedEof)
    }

    fn parse_object_hex_destination(&mut self) -> RtfResult<Vec<u8>> {
        let mut data = Vec::new();
        let mut high_nibble = None;
        let mut depth = 0usize;
        self.pos += 1; // opening brace
        while self.pos < self.tokens.len() {
            match &self.tokens[self.pos] {
                Token::OpenBrace => {
                    depth += 1;
                    self.pos += 1;
                },
                Token::CloseBrace if depth == 0 => {
                    self.pos += 1;
                    if high_nibble.is_some() {
                        return Err(RtfError::MalformedDocument(
                            "RTF objdata contains an odd number of hexadecimal digits".to_string(),
                        ));
                    }
                    return Ok(data);
                },
                Token::CloseBrace => {
                    depth -= 1;
                    self.pos += 1;
                },
                Token::Text(text) => {
                    for byte in text.bytes().filter(|byte| !byte.is_ascii_whitespace()) {
                        let nibble = Self::hex_nibble(byte).ok_or_else(|| {
                            RtfError::MalformedDocument(
                                "RTF objdata contains a non-hexadecimal character".to_string(),
                            )
                        })?;
                        if let Some(high) = high_nibble.take() {
                            data.push((high << 4) | nibble);
                            if data.len() > MAX_OBJECT_DATA_BYTES {
                                return Err(RtfError::MalformedDocument(
                                    "RTF embedded object data exceeds the safety limit".to_string(),
                                ));
                            }
                        } else {
                            high_nibble = Some(nibble);
                        }
                    }
                    self.pos += 1;
                },
                Token::Binary(bytes) => {
                    if high_nibble.is_some() {
                        return Err(RtfError::MalformedDocument(
                            "RTF objdata binary payload splits a hexadecimal byte".to_string(),
                        ));
                    }
                    data.extend_from_slice(bytes);
                    if data.len() > MAX_OBJECT_DATA_BYTES {
                        return Err(RtfError::MalformedDocument(
                            "RTF embedded object data exceeds the safety limit".to_string(),
                        ));
                    }
                    self.pos += 1;
                },
                _ => self.pos += 1,
            }
        }
        Err(RtfError::UnexpectedEof)
    }

    fn hex_nibble(byte: u8) -> Option<u8> {
        match byte {
            b'0'..=b'9' => Some(byte - b'0'),
            b'a'..=b'f' => Some(byte - b'a' + 10),
            b'A'..=b'F' => Some(byte - b'A' + 10),
            _ => None,
        }
    }

    /// Parse a `shp` destination and its nested shape-property groups.
    fn parse_shape_destination(&mut self) -> RtfResult<super::shape::Shape<'a>> {
        use super::shape::{Shape, ShapeType};

        let mut shape = Shape::new(ShapeType::Unknown);
        let mut text = String::new();
        let mut depth = 0usize;
        let mut text_depth = None;
        let mut right = None;
        let mut bottom = None;
        let mut closed = false;
        self.pos += 1; // consume \shp

        while self.pos < self.tokens.len() {
            match &self.tokens[self.pos] {
                Token::OpenBrace
                    if self.nested_control_word() == Some(ControlWord::ShapeProperty) =>
                {
                    if shape.properties.len() >= MAX_SHAPE_PROPERTIES {
                        return Err(RtfError::MalformedDocument(
                            "RTF shape property count exceeds the safety limit".to_string(),
                        ));
                    }
                    let (name, value) = self.parse_shape_property_group()?;
                    shape.properties.push(super::shape::ShapeProperty::new(
                        Cow::Owned(name),
                        Cow::Owned(value),
                    ));
                },
                Token::OpenBrace => {
                    depth += 1;
                    self.pos += 1;
                },
                Token::CloseBrace if depth == 0 => {
                    self.pos += 1;
                    closed = true;
                    break;
                },
                Token::CloseBrace => {
                    if text_depth == Some(depth) {
                        text_depth = None;
                    }
                    depth -= 1;
                    self.pos += 1;
                },
                Token::Control(ControlWord::ShapeType(value)) => {
                    shape.shape_type = Self::shape_type_from_rtf(*value);
                    self.pos += 1;
                },
                Token::Control(ControlWord::ShapeLeft(value)) => {
                    shape.geometry.x = *value;
                    self.pos += 1;
                },
                Token::Control(ControlWord::ShapeTop(value)) => {
                    shape.geometry.y = *value;
                    self.pos += 1;
                },
                Token::Control(ControlWord::ShapeRight(value)) => {
                    right = Some(*value);
                    self.pos += 1;
                },
                Token::Control(ControlWord::ShapeBottom(value)) => {
                    bottom = Some(*value);
                    self.pos += 1;
                },
                Token::Control(ControlWord::ShapeWidth(value)) => {
                    shape.geometry.width = *value;
                    self.pos += 1;
                },
                Token::Control(ControlWord::ShapeHeight(value)) => {
                    shape.geometry.height = *value;
                    self.pos += 1;
                },
                Token::Control(ControlWord::ShapeRotation(value)) => {
                    shape.geometry.rotation = *value;
                    self.pos += 1;
                },
                Token::Control(ControlWord::ShapeZOrder(value)) => {
                    shape.geometry.z_order = *value;
                    self.pos += 1;
                },
                Token::Control(ControlWord::ShapeWrap(value)) => {
                    shape.wrap_mode = match value {
                        1 => super::shape::WrapMode::None,
                        2 => super::shape::WrapMode::Square,
                        4 => super::shape::WrapMode::Tight,
                        3 | 5 => super::shape::WrapMode::Through,
                        _ => shape.wrap_mode,
                    };
                    self.pos += 1;
                },
                Token::Control(ControlWord::ShapeBelowText(value)) => {
                    shape.behind_doc = *value;
                    if *value {
                        shape.wrap_mode = super::shape::WrapMode::Behind;
                    }
                    self.pos += 1;
                },
                Token::Control(ControlWord::ShapeLockAnchor) => {
                    shape.locked = true;
                    self.pos += 1;
                },
                Token::Control(ControlWord::ShapeText) => {
                    text_depth = Some(depth);
                    self.pos += 1;
                },
                Token::Control(ControlWord::Unicode(code)) if text_depth.is_some() => {
                    text.push_str(&self.parse_destination_unicode_sequence(*code)?);
                    if text.len() > MAX_SHAPE_TEXT_BYTES {
                        return Err(RtfError::MalformedDocument(
                            "RTF shape text exceeds the safety limit".to_string(),
                        ));
                    }
                },
                Token::Control(ControlWord::Par | ControlWord::Line) if text_depth.is_some() => {
                    text.push('\n');
                    self.pos += 1;
                },
                Token::Control(ControlWord::Tab) if text_depth.is_some() => {
                    text.push('\t');
                    self.pos += 1;
                },
                Token::Control(control)
                    if text_depth.is_some() && control_symbol_text(control).is_some() =>
                {
                    text.push_str(control_symbol_text(control).unwrap_or_default());
                    self.pos += 1;
                },
                Token::Text(value) if text_depth.is_some() => {
                    text.push_str(&self.decode_transport_text(value)?);
                    if text.len() > MAX_SHAPE_TEXT_BYTES {
                        return Err(RtfError::MalformedDocument(
                            "RTF shape text exceeds the safety limit".to_string(),
                        ));
                    }
                    self.pos += 1;
                },
                _ => self.pos += 1,
            }
        }
        if !closed {
            return Err(RtfError::UnexpectedEof);
        }
        if !text.is_empty() {
            shape.text = Cow::Owned(text);
            shape.text_formatting = self.current_state().ok().map(|state| state.formatting);
        }
        Self::apply_shape_properties(&mut shape);
        if let Some(right) = right {
            shape.geometry.width = right.saturating_sub(shape.geometry.x);
        }
        if let Some(bottom) = bottom {
            shape.geometry.height = bottom.saturating_sub(shape.geometry.y);
        }
        Ok(shape)
    }

    fn parse_shape_property_group(&mut self) -> RtfResult<(String, String)> {
        #[derive(Clone, Copy, PartialEq, Eq)]
        enum PropertyPart {
            Name,
            Value,
        }

        let mut name = String::new();
        let mut value = String::new();
        let mut part = None;
        let mut part_depth = None;
        let mut depth = 0usize;
        self.pos += 1; // consume the opening brace
        while self.pos < self.tokens.len() {
            match &self.tokens[self.pos] {
                Token::OpenBrace => {
                    depth += 1;
                    self.pos += 1;
                },
                Token::CloseBrace if depth == 0 => {
                    self.pos += 1;
                    return Ok((name.trim().to_string(), value.trim().to_string()));
                },
                Token::CloseBrace => {
                    if part_depth == Some(depth) {
                        part = None;
                        part_depth = None;
                    }
                    depth -= 1;
                    self.pos += 1;
                },
                Token::Control(ControlWord::ShapePropertyName) => {
                    part = Some(PropertyPart::Name);
                    part_depth = Some(depth);
                    self.pos += 1;
                },
                Token::Control(ControlWord::ShapePropertyValue) => {
                    part = Some(PropertyPart::Value);
                    part_depth = Some(depth);
                    self.pos += 1;
                },
                Token::Control(ControlWord::Unicode(code)) if part.is_some() => {
                    let decoded = self.parse_destination_unicode_sequence(*code)?;
                    match part {
                        Some(PropertyPart::Name) => name.push_str(&decoded),
                        Some(PropertyPart::Value) => value.push_str(&decoded),
                        None => {},
                    }
                },
                Token::Control(control)
                    if part.is_some() && control_symbol_text(control).is_some() =>
                {
                    let decoded = control_symbol_text(control).unwrap_or_default();
                    match part {
                        Some(PropertyPart::Name) => name.push_str(decoded),
                        Some(PropertyPart::Value) => value.push_str(decoded),
                        None => {},
                    }
                    self.pos += 1;
                },
                Token::Text(text) => {
                    let decoded = self.decode_transport_text(text)?;
                    match part {
                        Some(PropertyPart::Name) => name.push_str(&decoded),
                        Some(PropertyPart::Value) => value.push_str(&decoded),
                        None => {},
                    }
                    self.pos += 1;
                },
                _ => self.pos += 1,
            }
            if name.len().saturating_add(value.len()) > MAX_SHAPE_PROPERTY_BYTES {
                return Err(RtfError::MalformedDocument(
                    "RTF shape property exceeds the safety limit".to_string(),
                ));
            }
        }
        Err(RtfError::UnexpectedEof)
    }

    fn apply_shape_property(shape: &mut super::shape::Shape<'a>, name: &str, value: &str) {
        match name {
            "shapeType" => {
                if let Ok(value) = value.parse() {
                    shape.shape_type = Self::shape_type_from_rtf(value);
                }
            },
            "wzName" => shape.name = Cow::Owned(value.to_string()),
            "fBehindDocument" => {
                if let Some(value) = Self::parse_shape_bool(value) {
                    shape.behind_doc = value;
                }
            },
            "fBackground" => {
                if let Some(value) = Self::parse_shape_bool(value) {
                    shape.is_background = value;
                }
            },
            "fLockPosition" | "fLockAgainstGrouping" => {
                if let Some(value) = Self::parse_shape_bool(value) {
                    shape.locked |= value;
                }
            },
            "fillType" => {
                if let Ok(value) = value.parse::<i32>() {
                    shape.fill.fill_type = match value {
                        0 => super::shape::FillType::Solid,
                        1 => super::shape::FillType::Pattern,
                        2 => super::shape::FillType::Texture,
                        3 => super::shape::FillType::Picture,
                        4..=8 => super::shape::FillType::Gradient,
                        9 => super::shape::FillType::Background,
                        _ => shape.fill.fill_type,
                    };
                }
            },
            "fillColor" => {
                if let Some(value) = Self::parse_office_art_u32(value) {
                    shape.fill.color = super::shape::OfficeArtColor(value);
                }
            },
            "fillBackColor" => {
                if let Some(value) = Self::parse_office_art_u32(value) {
                    shape.fill.color2 = Some(super::shape::OfficeArtColor(value));
                }
            },
            "fillOpacity" => {
                if let Some(value) = Self::parse_office_art_u32(value) {
                    shape.fill.opacity = super::shape::OfficeArtOpacity(value);
                }
            },
            "lineColor" => {
                if let Some(value) = Self::parse_office_art_u32(value) {
                    shape.line.color = super::shape::OfficeArtColor(value);
                }
            },
            "lineWidth" => {
                if let Ok(value) = value.parse() {
                    shape.line.width_emu = value;
                }
            },
            "rotation" => {
                if let Ok(value) = value.parse::<i32>() {
                    shape.geometry.rotation = value / 65_536;
                }
            },
            _ => {},
        }
    }

    fn apply_shape_properties(shape: &mut super::shape::Shape<'a>) {
        for index in 0..shape.properties.len() {
            let name = shape.properties[index].name.to_string();
            let value = shape.properties[index].value.to_string();
            Self::apply_shape_property(shape, &name, &value);
        }

        if let Some(value) = shape
            .properties
            .iter()
            .rev()
            .find(|property| property.name == "fFilled")
            .and_then(|property| Self::parse_shape_bool(&property.value))
        {
            if value {
                if shape.fill.fill_type == super::shape::FillType::None {
                    shape.fill.fill_type = super::shape::FillType::Solid;
                }
            } else {
                shape.fill.fill_type = super::shape::FillType::None;
            }
        }

        if let Some(value) = shape
            .properties
            .iter()
            .rev()
            .find(|property| property.name == "fLine")
            .and_then(|property| Self::parse_shape_bool(&property.value))
        {
            shape.line.visible = value;
        }
    }

    fn parse_shape_bool(value: &str) -> Option<bool> {
        value.trim().parse::<i32>().ok().map(|value| value != 0)
    }

    fn parse_office_art_u32(value: &str) -> Option<u32> {
        let value = value.trim();
        value
            .parse::<u32>()
            .ok()
            .or_else(|| value.parse::<i32>().ok().map(|value| value as u32))
    }

    fn parse_shape_group_destination(&mut self) -> RtfResult<super::shape::ShapeGroup<'a>> {
        self.parse_shape_group_destination_at_depth(0)
    }

    fn parse_shape_group_destination_at_depth(
        &mut self,
        nesting_depth: usize,
    ) -> RtfResult<super::shape::ShapeGroup<'a>> {
        if nesting_depth >= MAX_SHAPE_GROUP_DEPTH {
            return Err(RtfError::MalformedDocument(
                "RTF shape group nesting exceeds the safety limit".to_string(),
            ));
        }
        let mut group = super::shape::ShapeGroup::new();
        let mut depth = 0usize;
        let mut right = None;
        let mut bottom = None;
        let mut closed = false;
        self.pos += 1; // consume \shpgrp

        while self.pos < self.tokens.len() {
            match &self.tokens[self.pos] {
                Token::OpenBrace
                    if self.nested_control_word() == Some(ControlWord::ShapeProperty) =>
                {
                    if group.properties.len() >= MAX_SHAPE_PROPERTIES {
                        return Err(RtfError::MalformedDocument(
                            "RTF shape group property count exceeds the safety limit".to_string(),
                        ));
                    }
                    let (name, value) = self.parse_shape_property_group()?;
                    if name == "wzName" {
                        group.name = Cow::Owned(value.clone());
                    }
                    group.properties.push(super::shape::ShapeProperty::new(
                        Cow::Owned(name),
                        Cow::Owned(value),
                    ));
                },
                Token::OpenBrace if self.nested_shape_control() == Some(ControlWord::Shape) => {
                    if group.shapes.len() >= MAX_SHAPES_PER_GROUP {
                        return Err(RtfError::MalformedDocument(
                            "RTF shape group child count exceeds the safety limit".to_string(),
                        ));
                    }
                    self.pos += 1;
                    if matches!(
                        self.tokens.get(self.pos),
                        Some(Token::Control(ControlWord::IgnorableDestination))
                    ) {
                        self.pos += 1;
                    }
                    group.add_shape(self.parse_shape_destination()?);
                },
                Token::OpenBrace
                    if self.nested_shape_control() == Some(ControlWord::ShapeGroup) =>
                {
                    if group.groups.len() >= MAX_GROUPS_PER_GROUP {
                        return Err(RtfError::MalformedDocument(
                            "RTF nested shape group count exceeds the safety limit".to_string(),
                        ));
                    }
                    self.pos += 1;
                    if matches!(
                        self.tokens.get(self.pos),
                        Some(Token::Control(ControlWord::IgnorableDestination))
                    ) {
                        self.pos += 1;
                    }
                    let nested = self
                        .parse_shape_group_destination_at_depth(nesting_depth.saturating_add(1))?;
                    group.add_group(nested);
                },
                Token::OpenBrace => {
                    depth += 1;
                    self.pos += 1;
                },
                Token::CloseBrace if depth == 0 => {
                    self.pos += 1;
                    closed = true;
                    break;
                },
                Token::CloseBrace => {
                    depth -= 1;
                    self.pos += 1;
                },
                Token::Control(ControlWord::ShapeLeft(value)) => {
                    group.geometry.x = *value;
                    self.pos += 1;
                },
                Token::Control(ControlWord::ShapeTop(value)) => {
                    group.geometry.y = *value;
                    self.pos += 1;
                },
                Token::Control(ControlWord::ShapeRight(value)) => {
                    right = Some(*value);
                    self.pos += 1;
                },
                Token::Control(ControlWord::ShapeBottom(value)) => {
                    bottom = Some(*value);
                    self.pos += 1;
                },
                Token::Control(ControlWord::ShapeWidth(value)) => {
                    group.geometry.width = *value;
                    self.pos += 1;
                },
                Token::Control(ControlWord::ShapeHeight(value)) => {
                    group.geometry.height = *value;
                    self.pos += 1;
                },
                Token::Control(ControlWord::ShapeRotation(value)) => {
                    group.geometry.rotation = *value;
                    self.pos += 1;
                },
                Token::Control(ControlWord::ShapeZOrder(value)) => {
                    group.geometry.z_order = *value;
                    self.pos += 1;
                },
                _ => self.pos += 1,
            }
        }
        if !closed {
            return Err(RtfError::UnexpectedEof);
        }
        if let Some(right) = right {
            group.geometry.width = right.saturating_sub(group.geometry.x);
        }
        if let Some(bottom) = bottom {
            group.geometry.height = bottom.saturating_sub(group.geometry.y);
        }
        Ok(group)
    }

    fn nested_shape_control(&self) -> Option<ControlWord<'a>> {
        match self.nested_control_word()? {
            control @ (ControlWord::Shape | ControlWord::ShapeGroup) => Some(control),
            _ => None,
        }
    }

    fn nested_control_word(&self) -> Option<ControlWord<'a>> {
        let mut index = self.pos.checked_add(1)?;
        if matches!(
            self.tokens.get(index),
            Some(Token::Control(ControlWord::IgnorableDestination))
        ) {
            index += 1;
        }
        match self.tokens.get(index) {
            Some(Token::Control(control)) => Some(*control),
            _ => None,
        }
    }

    fn shape_type_from_rtf(value: i32) -> super::shape::ShapeType {
        use super::shape::ShapeType;
        match value {
            1 => ShapeType::Rectangle,
            2 => ShapeType::RoundRectangle,
            3 => ShapeType::Ellipse,
            19 => ShapeType::Arc,
            20 => ShapeType::Line,
            75 => ShapeType::PictureFrame,
            202 => ShapeType::TextBox,
            0 => ShapeType::Group,
            value => ShapeType::Custom(value),
        }
    }

    fn apply_legacy_text_box_control(
        builder: &mut LegacyTextBoxBuilder,
        control: &ControlWord,
    ) -> RtfResult<bool> {
        macro_rules! set_once {
            ($slot:expr, $value:expr, $name:literal) => {{
                if $slot.is_some() {
                    return Err(RtfError::MalformedDocument(
                        concat!("duplicate RTF legacy text-box ", $name).to_string(),
                    ));
                }
                $slot = Some($value);
                true
            }};
        }
        Ok(match control {
            ControlWord::LegacyAnchorXPage => set_once!(
                builder.horizontal_anchor,
                crate::LegacyHorizontalAnchor::Page,
                "horizontal anchor"
            ),
            ControlWord::LegacyAnchorXMargin => set_once!(
                builder.horizontal_anchor,
                crate::LegacyHorizontalAnchor::Margin,
                "horizontal anchor"
            ),
            ControlWord::LegacyAnchorXColumn => set_once!(
                builder.horizontal_anchor,
                crate::LegacyHorizontalAnchor::Column,
                "horizontal anchor"
            ),
            ControlWord::LegacyAnchorYPage => set_once!(
                builder.vertical_anchor,
                crate::LegacyVerticalAnchor::Page,
                "vertical anchor"
            ),
            ControlWord::LegacyAnchorYMargin => set_once!(
                builder.vertical_anchor,
                crate::LegacyVerticalAnchor::Margin,
                "vertical anchor"
            ),
            ControlWord::LegacyAnchorYParagraph => set_once!(
                builder.vertical_anchor,
                crate::LegacyVerticalAnchor::Paragraph,
                "vertical anchor"
            ),
            ControlWord::LegacyDrawingX(value) => set_once!(builder.x, *value, "x"),
            ControlWord::LegacyDrawingY(value) => set_once!(builder.y, *value, "y"),
            ControlWord::LegacyDrawingWidth(value) => {
                set_once!(builder.width, *value, "width")
            },
            ControlWord::LegacyDrawingHeightSize(value) => {
                set_once!(builder.height, *value, "height")
            },
            ControlWord::LegacyTextBoxMargin(value) => {
                set_once!(builder.margin, *value, "margin")
            },
            ControlWord::LegacyDrawingHeight(value) => {
                set_once!(builder.z_order, *value, "z-order")
            },
            ControlWord::LegacyTextLeftRightTopBottom => set_once!(
                builder.direction,
                crate::LegacyTextDirection::LeftToRightTopToBottom,
                "direction"
            ),
            ControlWord::LegacyTextLeftRightTopBottomVertical => set_once!(
                builder.direction,
                crate::LegacyTextDirection::LeftToRightTopToBottomVertical,
                "direction"
            ),
            ControlWord::LegacyTextTopBottomRightLeft => set_once!(
                builder.direction,
                crate::LegacyTextDirection::TopToBottomRightToLeft,
                "direction"
            ),
            ControlWord::LegacyTextTopBottomRightLeftVertical => set_once!(
                builder.direction,
                crate::LegacyTextDirection::TopToBottomRightToLeftVertical,
                "direction"
            ),
            ControlWord::LegacyTextBottomTopLeftRight => set_once!(
                builder.direction,
                crate::LegacyTextDirection::BottomToTopLeftToRight,
                "direction"
            ),
            _ => false,
        })
    }

    fn parse_legacy_text_box(&mut self) -> RtfResult<Option<crate::LegacyTextBox<'a>>> {
        if self.current_state()?.destination != Destination::DocumentBody {
            return Err(RtfError::MalformedDocument(
                "RTF legacy drawing text box may occur only in the document body".to_string(),
            ));
        }
        self.pos += 2; // ignorable marker and do
        let mut builder = LegacyTextBoxBuilder::default();
        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    if !builder.saw_text_box {
                        return Ok(None);
                    }
                    let text = builder.text.ok_or_else(|| {
                        RtfError::MalformedDocument(
                            "RTF legacy text box lacks dptxbxtext".to_string(),
                        )
                    })?;
                    let text_box = crate::LegacyTextBox {
                        text: Cow::Borrowed(self.arena.alloc_str(&text) as &str),
                        position: self.body_text_len,
                        horizontal_anchor: builder.horizontal_anchor,
                        vertical_anchor: builder.vertical_anchor,
                        x: builder.x,
                        y: builder.y,
                        width: builder.width,
                        height: builder.height,
                        margin: builder.margin,
                        z_order: builder.z_order,
                        direction: builder.direction.unwrap_or_default(),
                    };
                    text_box.validate()?;
                    if self.legacy_text_boxes.len()
                        >= crate::legacy_text_box::MAX_LEGACY_TEXT_BOXES
                    {
                        return Err(RtfError::MalformedDocument(
                            "RTF legacy text-box count exceeds the safety limit".to_string(),
                        ));
                    }
                    self.legacy_text_box_text_bytes = self
                        .legacy_text_box_text_bytes
                        .checked_add(text_box.text.len())
                        .ok_or_else(|| {
                            RtfError::MalformedDocument(
                                "RTF legacy text-box text size overflow".to_string(),
                            )
                        })?;
                    if self.legacy_text_box_text_bytes
                        > crate::legacy_text_box::MAX_LEGACY_TEXT_BOX_TOTAL_BYTES
                    {
                        return Err(RtfError::MalformedDocument(
                            "RTF legacy text-box text exceeds the aggregate safety limit"
                                .to_string(),
                        ));
                    }
                    return Ok(Some(text_box));
                },
                Some(Token::OpenBrace)
                    if matches!(
                        self.tokens.get(self.pos + 1),
                        Some(Token::Control(ControlWord::LegacyTextBoxText))
                    ) =>
                {
                    if !builder.saw_text_box {
                        return Err(RtfError::MalformedDocument(
                            "RTF legacy drawing dptxbxtext must follow dptxbx".to_string(),
                        ));
                    }
                    if builder.text.is_some() {
                        return Err(RtfError::MalformedDocument(
                            "RTF legacy text box contains duplicate dptxbxtext".to_string(),
                        ));
                    }
                    builder.text = Some(self.parse_legacy_text_box_text(&mut builder)?);
                },
                Some(Token::OpenBrace) => self.skip_group()?,
                Some(Token::Control(ControlWord::LegacyTextBox)) => {
                    if builder.saw_text_box {
                        return Err(RtfError::MalformedDocument(
                            "RTF legacy drawing contains duplicate dptxbx".to_string(),
                        ));
                    }
                    builder.saw_text_box = true;
                    self.pos += 1;
                },
                Some(Token::Control(control)) => {
                    Self::apply_legacy_text_box_control(&mut builder, control)?;
                    self.pos += 1;
                },
                Some(Token::Text(text)) if text.trim().is_empty() => self.pos += 1,
                Some(Token::Text(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF legacy drawing contains orphan text".to_string(),
                    ));
                },
                Some(Token::Binary(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF legacy drawing cannot contain binary data".to_string(),
                    ));
                },
                None => return Err(RtfError::UnexpectedEof),
            }
        }
    }

    fn parse_legacy_text_box_text(
        &mut self,
        builder: &mut LegacyTextBoxBuilder,
    ) -> RtfResult<String> {
        self.pos += 2; // opening brace and dptxbxtext
        let mut depth = 0usize;
        let mut unicode_skip = self.current_state()?.unicode_skip.max(0);
        let mut text = String::new();
        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) if depth == 0 => {
                    self.pos += 1;
                    return Ok(text);
                },
                Some(Token::CloseBrace) => {
                    depth -= 1;
                    self.pos += 1;
                },
                Some(Token::OpenBrace) => {
                    if matches!(
                        self.tokens.get(self.pos + 1),
                        Some(Token::Control(ControlWord::IgnorableDestination))
                    ) || matches!(
                        self.tokens.get(self.pos + 1),
                        Some(Token::Control(
                            ControlWord::Field
                                | ControlWord::Object
                                | ControlWord::Picture
                                | ControlWord::Shape
                                | ControlWord::FormField
                        ))
                    ) {
                        return Err(RtfError::MalformedDocument(
                            "RTF legacy text box contains an active nested destination"
                                .to_string(),
                        ));
                    }
                    depth += 1;
                    self.pos += 1;
                },
                Some(Token::Control(control))
                    if Self::apply_legacy_text_box_control(builder, control)? =>
                {
                    self.pos += 1;
                },
                Some(Token::Control(ControlWord::Unicode(code))) => {
                    text.push_str(&self.parse_style_unicode(*code, unicode_skip)?);
                },
                Some(Token::Control(ControlWord::UnicodeSkip(value))) => {
                    unicode_skip = (*value).max(0);
                    self.pos += 1;
                },
                Some(Token::Control(ControlWord::Tab)) => {
                    text.push('\t');
                    self.pos += 1;
                },
                Some(Token::Control(ControlWord::Par | ControlWord::Line)) => {
                    text.push('\n');
                    self.pos += 1;
                },
                Some(Token::Control(control)) if control_symbol_text(control).is_some() => {
                    text.push_str(control_symbol_text(control).unwrap_or_default());
                    self.pos += 1;
                },
                Some(Token::Control(
                    ControlWord::Field
                    | ControlWord::Object
                    | ControlWord::Picture
                    | ControlWord::Shape
                    | ControlWord::FormField,
                )) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF legacy text box contains active content".to_string(),
                    ));
                },
                Some(Token::Control(_)) => self.pos += 1,
                Some(Token::Text(value)) => {
                    text.push_str(&self.decode_transport_text(value)?);
                    self.pos += 1;
                },
                Some(Token::Binary(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF legacy text box cannot contain binary data".to_string(),
                    ));
                },
                None => return Err(RtfError::UnexpectedEof),
            }
            if text.len() > crate::legacy_text_box::MAX_LEGACY_TEXT_BOX_BYTES {
                return Err(RtfError::MalformedDocument(
                    "RTF legacy text-box text exceeds the safety limit".to_string(),
                ));
            }
        }
    }

    /// Parse picture/image content.
    ///
    /// Pictures in RTF have the format:
    /// {\pict\emfblip\picw<width>\pich<height>...<hex data>}
    fn parse_picture(&mut self) -> RtfResult<()> {
        self.pos += 1; // Skip \pict

        let mut image_type = super::picture::ImageType::Unknown;
        let mut width = None;
        let mut height = None;
        let mut goal_width = None;
        let mut goal_height = None;
        let mut scale_x = None;
        let mut scale_y = None;
        let mut blip_tag = None;
        let mut blip_upi = None;
        let mut blip_uid = None;
        let mut identity_stage = 0u8;
        let mut data_started = false;
        let mut data = Vec::new();
        let mut high_nibble = None;

        // Parse picture properties and data
        while self.pos < self.tokens.len() {
            match &self.tokens[self.pos] {
                Token::CloseBrace => {
                    break;
                },
                Token::Control(control) => {
                    self.pos += 1;
                    match control {
                        ControlWord::Emfblip => image_type = super::picture::ImageType::Emf,
                        ControlWord::Pngblip => image_type = super::picture::ImageType::Png,
                        ControlWord::Jpegblip => image_type = super::picture::ImageType::Jpeg,
                        ControlWord::Macpict => image_type = super::picture::ImageType::Pict,
                        ControlWord::Wmetafile(_) | ControlWord::Pmmetafile(_) => {
                            image_type = super::picture::ImageType::Wmf
                        },
                        ControlWord::Dibitmap(_) | ControlWord::Wbitmap(_) => {
                            image_type = super::picture::ImageType::Dib
                        },
                        ControlWord::PictureWidth(w) => width = Some(*w),
                        ControlWord::PictureHeight(h) => height = Some(*h),
                        ControlWord::PictureGoalWidth(w) => goal_width = Some(*w),
                        ControlWord::PictureGoalHeight(h) => goal_height = Some(*h),
                        ControlWord::PictureScaleX(s) => scale_x = Some(*s),
                        ControlWord::PictureScaleY(s) => scale_y = Some(*s),
                        ControlWord::BlipTag(value) => {
                            if data_started || blip_tag.is_some() || identity_stage > 1 {
                                return Err(RtfError::MalformedDocument(
                                    "RTF bliptag is duplicated, late, or out of order".to_string(),
                                ));
                            }
                            blip_tag = Some(*value);
                            identity_stage = 1;
                        },
                        ControlWord::BlipUnitsPerInch(value) => {
                            if data_started || blip_upi.is_some() || identity_stage > 2 {
                                return Err(RtfError::MalformedDocument(
                                    "RTF blipupi is duplicated, late, or out of order".to_string(),
                                ));
                            }
                            let value = u16::try_from(*value).map_err(|_| {
                                RtfError::MalformedDocument(
                                    "RTF blipupi is outside 1..=65535".to_string(),
                                )
                            })?;
                            if value == 0 {
                                return Err(RtfError::MalformedDocument(
                                    "RTF blipupi must be positive".to_string(),
                                ));
                            }
                            blip_upi = Some(value);
                            identity_stage = 2;
                        },
                        ControlWord::BlipUid => {
                            return Err(RtfError::MalformedDocument(
                                "RTF blipuid destination must be starred and grouped".to_string(),
                            ));
                        },
                        _ => {},
                    }
                },
                Token::Text(text) => {
                    data_started |= text.bytes().any(|byte| !byte.is_ascii_whitespace());
                    for byte in text.bytes().filter(|byte| !byte.is_ascii_whitespace()) {
                        let nibble = Self::hex_nibble(byte).ok_or_else(|| {
                            RtfError::MalformedDocument(
                                "RTF picture contains a non-hexadecimal character".to_string(),
                            )
                        })?;
                        if let Some(high) = high_nibble.take() {
                            data.push((high << 4) | nibble);
                        } else {
                            high_nibble = Some(nibble);
                        }
                    }
                    if data.len() > MAX_PICTURE_DATA_BYTES {
                        return Err(RtfError::MalformedDocument(
                            "RTF picture data exceeds the safety limit".to_string(),
                        ));
                    }
                    self.pos += 1;
                },
                Token::Binary(bytes) => {
                    data_started = true;
                    if high_nibble.is_some() {
                        return Err(RtfError::MalformedDocument(
                            "RTF picture binary payload splits a hexadecimal byte".to_string(),
                        ));
                    }
                    data.extend_from_slice(bytes);
                    if data.len() > MAX_PICTURE_DATA_BYTES {
                        return Err(RtfError::MalformedDocument(
                            "RTF picture data exceeds the safety limit".to_string(),
                        ));
                    }
                    self.pos += 1;
                },
                Token::OpenBrace => {
                    if matches!(
                        self.tokens.get(self.pos..self.pos + 3),
                        Some([
                            Token::OpenBrace,
                            Token::Control(ControlWord::IgnorableDestination),
                            Token::Control(ControlWord::BlipUid),
                        ])
                    ) {
                        if data_started || blip_uid.is_some() {
                            return Err(RtfError::MalformedDocument(
                                "RTF blipuid is duplicated or occurs after picture data"
                                    .to_string(),
                            ));
                        }
                        blip_uid = Some(self.parse_picture_uid()?);
                        identity_stage = 3;
                    } else if matches!(
                        self.tokens.get(self.pos..self.pos + 2),
                        Some([Token::OpenBrace, Token::Control(ControlWord::BlipUid)])
                    ) {
                        return Err(RtfError::MalformedDocument(
                            "RTF blipuid destination must be starred".to_string(),
                        ));
                    } else {
                        self.skip_group()?;
                    }
                },
            }
        }

        if high_nibble.is_some() {
            return Err(RtfError::MalformedDocument(
                "RTF picture contains an odd number of hexadecimal digits".to_string(),
            ));
        }

        if !data.is_empty() {
            // If type not specified, try to detect from data
            if image_type == super::picture::ImageType::Unknown {
                image_type = super::picture::detect_image_type(&data);
            }

            // Allocate in arena and create picture
            let data_alloc = self.arena.alloc_slice_copy(&data);
            let mut picture = super::picture::Picture::new(image_type, Cow::Borrowed(data_alloc));
            picture.width = width;
            picture.height = height;
            picture.goal_width = goal_width;
            picture.goal_height = goal_height;
            picture.scale_x = scale_x;
            picture.scale_y = scale_y;
            if blip_tag.is_some() || blip_upi.is_some() || blip_uid.is_some() {
                let identity = super::picture::PictureIdentity {
                    tag: blip_tag,
                    units_per_inch: blip_upi,
                    uid: blip_uid.map(|uid| {
                        Cow::Borrowed(self.arena.alloc_slice_copy(&uid) as &[u8])
                    }),
                };
                identity.validate()?;
                picture.identity = Some(identity);
            }

            self.pictures.push(picture);
        }

        Ok(())
    }

    fn parse_picture_uid(&mut self) -> RtfResult<Vec<u8>> {
        self.pos += 3; // opening brace, ignorable marker, and blipuid
        let mut bytes = Vec::with_capacity(16);
        let mut high_nibble = None;
        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    if high_nibble.is_some() {
                        return Err(RtfError::MalformedDocument(
                            "RTF blipuid contains an odd number of hexadecimal digits"
                                .to_string(),
                        ));
                    }
                    if !bytes.is_empty() && bytes.len() != 16 {
                        return Err(RtfError::MalformedDocument(
                            "RTF blipuid must contain exactly 16 bytes or be empty".to_string(),
                        ));
                    }
                    return Ok(bytes);
                },
                Some(Token::Text(text)) => {
                    for byte in text.bytes().filter(|byte| !byte.is_ascii_whitespace()) {
                        let nibble = Self::hex_nibble(byte).ok_or_else(|| {
                            RtfError::MalformedDocument(
                                "RTF blipuid contains a non-hexadecimal character".to_string(),
                            )
                        })?;
                        if let Some(high) = high_nibble.take() {
                            bytes.push((high << 4) | nibble);
                            if bytes.len() > 16 {
                                return Err(RtfError::MalformedDocument(
                                    "RTF blipuid exceeds 16 bytes".to_string(),
                                ));
                            }
                        } else {
                            high_nibble = Some(nibble);
                        }
                    }
                    self.pos += 1;
                },
                Some(Token::OpenBrace | Token::Control(_) | Token::Binary(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF blipuid contains active, nested, or binary content".to_string(),
                    ));
                },
                None => return Err(RtfError::UnexpectedEof),
            }
        }
    }

    /// Parse field content.
    ///
    /// Fields in RTF have the format:
    /// {\field{\*\fldinst INSTRUCTION}{\fldrslt RESULT}}
    fn parse_field(&mut self) -> RtfResult<()> {
        let field_position = self.body_text_len;
        let enclosing_destination = self.current_state()?.destination;
        let field_in_table = self.current_state()?.in_table;
        self.pos += 1; // Skip \field

        let mut instruction = SmallVec::<[u8; 128]>::new();
        let mut result = SmallVec::<[u8; 128]>::new();
        let mut form_field = None;
        let mut data_field = None;
        let mut in_instruction;
        let mut in_result;

        // Parse field groups
        while self.pos < self.tokens.len() {
            match &self.tokens[self.pos] {
                Token::CloseBrace => {
                    // End of outer field group
                    break;
                },
                Token::OpenBrace => {
                    self.pos += 1;
                    // Check for fldinst or fldrslt
                    if self.pos < self.tokens.len() {
                        // Look for \*\fldinst or \fldrslt
                        let is_ignorable = matches!(
                            self.tokens.get(self.pos),
                            Some(Token::Control(ControlWord::IgnorableDestination))
                        );
                        if is_ignorable {
                            self.pos += 1;
                        }

                        if let Some(Token::Control(ControlWord::FieldInstruction)) =
                            self.tokens.get(self.pos)
                        {
                            self.pos += 1;
                            in_instruction = true;
                            in_result = false;
                            if let Some(state) = self.states.last_mut() {
                                state.destination = Destination::FieldInstruction;
                            }
                        } else if let Some(Token::Control(ControlWord::FieldResult)) =
                            self.tokens.get(self.pos)
                        {
                            self.pos += 1;
                            in_instruction = false;
                            in_result = true;
                            if let Some(state) = self.states.last_mut() {
                                state.destination = Destination::FieldResult;
                            }
                        } else {
                            // Skip unknown nested groups
                            self.skip_until_close_brace()?;
                            continue;
                        }

                        // Collect text until the destination's closing brace. Producers often
                        // wrap or split field instructions in formatting groups; those groups
                        // do not change the field-code text and must not discard it.
                        let mut nested_depth = 0usize;
                        while self.pos < self.tokens.len() {
                            match &self.tokens[self.pos] {
                                Token::CloseBrace if nested_depth == 0 => {
                                    self.pos += 1;
                                    break;
                                },
                                Token::CloseBrace => {
                                    nested_depth -= 1;
                                    self.pos += 1;
                                },
                                Token::Text(text) => {
                                    let decoded = self.decode_transport_text(text)?;
                                    if in_instruction {
                                        instruction.extend_from_slice(decoded.as_bytes());
                                    } else if in_result {
                                        result.extend_from_slice(decoded.as_bytes());
                                    }
                                    self.pos += 1;
                                },
                                Token::Control(ControlWord::Unicode(first)) => {
                                    let decoded = self.parse_style_unicode(
                                        *first,
                                        self.current_state()?.unicode_skip.max(0),
                                    )?;
                                    if in_instruction {
                                        instruction.extend_from_slice(decoded.as_bytes());
                                    } else if in_result {
                                        result.extend_from_slice(decoded.as_bytes());
                                    }
                                },
                                Token::Control(ControlWord::UnicodeSkip(value)) => {
                                    self.current_state_mut()?.unicode_skip = (*value).max(0);
                                    self.pos += 1;
                                },
                                Token::Control(ControlWord::Par | ControlWord::Line)
                                    if in_result =>
                                {
                                    result.push(b'\n');
                                    self.pos += 1;
                                },
                                Token::Control(ControlWord::Tab) if in_result => {
                                    result.push(b'\t');
                                    self.pos += 1;
                                },
                                Token::Control(control)
                                    if control_symbol_text(control).is_some() =>
                                {
                                    let decoded = control_symbol_text(control).unwrap_or_default();
                                    if in_instruction {
                                        instruction.extend_from_slice(decoded.as_bytes());
                                    } else if in_result {
                                        result.extend_from_slice(decoded.as_bytes());
                                    }
                                    self.pos += 1;
                                },
                                Token::OpenBrace if in_instruction => {
                                    let destination = match (
                                        self.tokens.get(self.pos + 1),
                                        self.tokens.get(self.pos + 2),
                                    ) {
                                        (
                                            Some(Token::Control(
                                                ControlWord::IgnorableDestination,
                                            )),
                                            Some(Token::Control(control)),
                                        ) => Some(control),
                                        (Some(Token::Control(control)), _) => Some(control),
                                        _ => None,
                                    };
                                    match destination {
                                        Some(ControlWord::FormField) => {
                                            if form_field.is_some() {
                                                return Err(RtfError::MalformedDocument(
                                                    "RTF field contains multiple formfield destinations"
                                                        .to_string(),
                                                ));
                                            }
                                            form_field =
                                                Some(self.parse_form_field_destination()?);
                                        },
                                        Some(ControlWord::DataField) => {
                                            if data_field.is_some() {
                                                return Err(RtfError::MalformedDocument(
                                                    "RTF field contains multiple datafield destinations"
                                                        .to_string(),
                                                ));
                                            }
                                            data_field = Some(self.parse_data_field_destination()?);
                                        },
                                        _ => {
                                            nested_depth = nested_depth.checked_add(1).ok_or_else(
                                                || {
                                                    RtfError::MalformedDocument(
                                                        "field instruction nesting depth overflow"
                                                            .to_string(),
                                                    )
                                                },
                                            )?;
                                            self.pos += 1;
                                        },
                                    }
                                },
                                Token::OpenBrace
                                    if matches!(
                                        self.tokens.get(self.pos + 1),
                                        Some(Token::Control(ControlWord::Field))
                                    ) =>
                                {
                                    self.pos += 1;
                                    self.parse_field()?;
                                    self.skip_until_close_brace()?;
                                },
                                Token::OpenBrace => {
                                    nested_depth = nested_depth.checked_add(1).ok_or_else(|| {
                                        RtfError::MalformedDocument(
                                            "field instruction nesting depth overflow".to_string(),
                                        )
                                    })?;
                                    self.pos += 1;
                                },
                                _ => {
                                    self.pos += 1;
                                },
                            }
                            if instruction.len() > super::form_field::MAX_FORM_FIELD_STRING_BYTES
                                || result.len()
                                    > super::form_field::MAX_FORM_FIELD_STRING_BYTES
                            {
                                return Err(RtfError::MalformedDocument(
                                    "RTF field instruction or result exceeds the safety limit"
                                        .to_string(),
                                ));
                            }
                        }
                    }
                },
                _ => {
                    self.pos += 1;
                },
            }
        }

        let result_text = std::str::from_utf8(&result).map_err(|_| {
            RtfError::MalformedDocument("RTF field result is not valid UTF-8".to_string())
        })?;

        // Create the generic field record if we have an instruction.
        if !instruction.is_empty()
            && let Ok(inst_str) = std::str::from_utf8(&instruction)
        {
            // Allocate instruction in arena first
            let inst_alloc = self.arena.alloc_str(inst_str);

            // Parse field type from allocated instruction
            let mut field = super::field::Field::parse_instruction(inst_alloc);
            field.instruction = Cow::Borrowed(inst_alloc);

            // Add result if available
            if !result.is_empty()
                && let Ok(res_str) = std::str::from_utf8(&result)
            {
                let res_alloc = self.arena.alloc_str(res_str);
                field.result = Cow::Borrowed(res_alloc);
            }

            self.fields.push(field);
        }

        self.current_state_mut()?.destination = enclosing_destination;
        if form_field.is_some() && !result_text.is_empty() {
            self.append_semantic_text(result_text)?;
        }

        if let Some(builder) = form_field {
            if self.form_fields.len() >= super::form_field::MAX_FORM_FIELDS {
                return Err(RtfError::MalformedDocument(
                    "RTF form-field count exceeds the safety limit".to_string(),
                ));
            }
            let field_type = builder.field_type.ok_or_else(|| {
                RtfError::MalformedDocument(
                    "RTF formfield destination is missing fftype".to_string(),
                )
            })?;
            let to_cow = |value: Option<String>| {
                value.map(|text| Cow::Borrowed(self.arena.alloc_str(&text) as &str))
            };
            let form_field = super::form_field::FormField {
                field_type,
                text_type: builder.text_type,
                name: to_cow(builder.name),
                max_length: builder.max_length,
                format: to_cow(builder.format),
                default_text: to_cow(builder.default_text),
                default_result: builder.default_result,
                result: builder.result,
                half_point_size: builder.half_point_size,
                protected: builder.protected.unwrap_or(false),
                calculate_on_exit: builder.calculate_on_exit.unwrap_or(false),
                size_automatically: builder.size_automatically.unwrap_or(false),
                own_help: builder.own_help.unwrap_or(false),
                own_status: builder.own_status.unwrap_or(false),
                help_text: to_cow(builder.help_text),
                status_text: to_cow(builder.status_text),
                entry_macro: to_cow(builder.entry_macro),
                exit_macro: to_cow(builder.exit_macro),
                list_entries: builder
                    .list_entries
                    .into_iter()
                    .map(|text| Cow::Borrowed(self.arena.alloc_str(&text) as &str))
                    .collect(),
                has_list_box: builder.has_list_box.unwrap_or(false),
                data: Cow::Borrowed(self.arena.alloc_slice_copy(
                    data_field.as_deref().unwrap_or_default(),
                )),
                result_text: Cow::Borrowed(
                    self.arena
                        .alloc_str(if field_in_table { "" } else { result_text }),
                ),
                position: field_position,
                range_end: if field_in_table {
                    field_position
                } else {
                    self.body_text_len
                },
            };
            form_field.validate()?;
            let added = form_field.text_bytes().ok_or_else(|| {
                RtfError::MalformedDocument("RTF form-field aggregate size overflow".to_string())
            })?;
            self.form_field_text_bytes = self
                .form_field_text_bytes
                .checked_add(added)
                .ok_or_else(|| {
                    RtfError::MalformedDocument(
                        "RTF form-field aggregate size overflow".to_string(),
                    )
                })?;
            if self.form_field_text_bytes > super::form_field::MAX_FORM_FIELD_TOTAL_BYTES {
                return Err(RtfError::MalformedDocument(
                    "RTF form-field aggregate text exceeds the safety limit".to_string(),
                ));
            }
            self.form_fields.push(form_field);
        } else if data_field.is_some() {
            // Data fields attached to non-form fields are inert legacy payloads and
            // are intentionally not exposed as executable/external content.
        }

        Ok(())
    }

    fn parse_form_field_destination(&mut self) -> RtfResult<FormFieldBuilder> {
        self.expect_token(Token::OpenBrace)?;
        if !matches!(
            self.tokens.get(self.pos),
            Some(Token::Control(ControlWord::IgnorableDestination))
        ) {
            return Err(RtfError::MalformedDocument(
                "RTF formfield destination must be starred".to_string(),
            ));
        }
        self.pos += 1;
        if !matches!(
            self.tokens.get(self.pos),
            Some(Token::Control(ControlWord::FormField))
        ) {
            return Err(RtfError::MalformedDocument(
                "invalid RTF formfield destination".to_string(),
            ));
        }
        self.pos += 1;
        let mut builder = FormFieldBuilder::default();
        let mut depth = 1usize;
        while self.pos < self.tokens.len() {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    depth -= 1;
                    if depth == 0 {
                        return Ok(builder);
                    }
                },
                Some(Token::OpenBrace) => {
                    let starred = matches!(
                        self.tokens.get(self.pos + 1),
                        Some(Token::Control(ControlWord::IgnorableDestination))
                    );
                    let control = match (
                        self.tokens.get(self.pos + 1),
                        self.tokens.get(self.pos + 2),
                    ) {
                        (
                            Some(Token::Control(ControlWord::IgnorableDestination)),
                            Some(Token::Control(control)),
                        ) => Some(control),
                        (Some(Token::Control(control)), _) => Some(control),
                        _ => None,
                    };
                    if !starred
                        && matches!(
                            control,
                            Some(
                                ControlWord::FormFieldName
                                    | ControlWord::FormFieldFormat
                                    | ControlWord::FormFieldDefaultText
                                    | ControlWord::FormFieldHelpText
                                    | ControlWord::FormFieldStatusText
                                    | ControlWord::FormFieldEntryMacro
                                    | ControlWord::FormFieldExitMacro
                                    | ControlWord::FormFieldListEntry
                            )
                        )
                    {
                        return Err(RtfError::MalformedDocument(
                            "RTF formfield text destinations must be starred".to_string(),
                        ));
                    }
                    let target = match control {
                        Some(ControlWord::FormFieldName) => &mut builder.name,
                        Some(ControlWord::FormFieldFormat) => &mut builder.format,
                        Some(ControlWord::FormFieldDefaultText) => &mut builder.default_text,
                        Some(ControlWord::FormFieldHelpText) => &mut builder.help_text,
                        Some(ControlWord::FormFieldStatusText) => &mut builder.status_text,
                        Some(ControlWord::FormFieldEntryMacro) => &mut builder.entry_macro,
                        Some(ControlWord::FormFieldExitMacro) => &mut builder.exit_macro,
                        Some(ControlWord::FormFieldListEntry) => {
                            if builder.list_entries.len()
                                >= super::form_field::MAX_FORM_FIELD_LIST_ENTRIES
                            {
                                return Err(RtfError::MalformedDocument(
                                    "RTF form-field list exceeds 25 entries".to_string(),
                                ));
                            }
                            builder
                                .list_entries
                                .push(self.parse_form_field_text_destination()?);
                            continue;
                        },
                        Some(
                            ControlWord::FormFieldType(_)
                            | ControlWord::FormFieldTextType(_)
                            | ControlWord::FormFieldMaxLength(_)
                            | ControlWord::FormFieldProtected(_)
                            | ControlWord::FormFieldRecalculate(_)
                            | ControlWord::FormFieldAutomaticSize(_)
                            | ControlWord::FormFieldDefaultResult(_)
                            | ControlWord::FormFieldResult(_)
                            | ControlWord::FormFieldHalfPointSize(_)
                            | ControlWord::FormFieldOwnHelp(_)
                            | ControlWord::FormFieldOwnStatus(_)
                            | ControlWord::FormFieldHasListBox(_),
                        ) => {
                            self.pos += 1;
                            depth = depth.checked_add(1).ok_or_else(|| {
                                RtfError::MalformedDocument(
                                    "RTF formfield nesting depth overflow".to_string(),
                                )
                            })?;
                            continue;
                        },
                        Some(_) => {
                            return Err(RtfError::MalformedDocument(
                                "RTF formfield contains an active or unknown nested destination"
                                    .to_string(),
                            ));
                        },
                        None => {
                            self.pos += 1;
                            depth = depth.checked_add(1).ok_or_else(|| {
                                RtfError::MalformedDocument(
                                    "RTF formfield nesting depth overflow".to_string(),
                                )
                            })?;
                            continue;
                        },
                    };
                    if target.is_some() {
                        return Err(RtfError::MalformedDocument(
                            "RTF formfield contains a duplicate text destination".to_string(),
                        ));
                    }
                    *target = Some(self.parse_form_field_text_destination()?);
                },
                Some(Token::Control(control)) => {
                    macro_rules! set_once {
                        ($slot:expr, $value:expr, $name:literal) => {{
                            if $slot.is_some() {
                                return Err(RtfError::MalformedDocument(concat!(
                                    "duplicate RTF formfield ",
                                    $name
                                )
                                .to_string()));
                            }
                            $slot = Some($value);
                        }};
                    }
                    match control {
                        ControlWord::FormFieldType(value) => set_once!(
                            builder.field_type,
                            super::form_field::FormFieldType::from_rtf(*value)?,
                            "fftype"
                        ),
                        ControlWord::FormFieldTextType(value) => set_once!(
                            builder.text_type,
                            super::form_field::FormTextType::from_rtf(*value)?,
                            "fftypetxt"
                        ),
                        ControlWord::FormFieldMaxLength(value) => set_once!(
                            builder.max_length,
                            u16::try_from(*value).map_err(|_| RtfError::MalformedDocument(
                                "RTF ffmaxlen is outside 0..=65535".to_string()
                            ))?,
                            "ffmaxlen"
                        ),
                        ControlWord::FormFieldProtected(value) => {
                            set_once!(builder.protected, *value, "ffprot")
                        },
                        ControlWord::FormFieldRecalculate(value) => {
                            set_once!(builder.calculate_on_exit, *value, "ffrecalc")
                        },
                        ControlWord::FormFieldAutomaticSize(value) => {
                            set_once!(builder.size_automatically, *value, "ffsize")
                        },
                        ControlWord::FormFieldDefaultResult(value) => {
                            set_once!(builder.default_result, *value, "ffdefres")
                        },
                        ControlWord::FormFieldResult(value) => {
                            set_once!(builder.result, *value, "ffres")
                        },
                        ControlWord::FormFieldHalfPointSize(value) => {
                            set_once!(builder.half_point_size, *value, "ffhps")
                        },
                        ControlWord::FormFieldOwnHelp(value) => {
                            set_once!(builder.own_help, *value, "ffownhelp")
                        },
                        ControlWord::FormFieldOwnStatus(value) => {
                            set_once!(builder.own_status, *value, "ffownstat")
                        },
                        ControlWord::FormFieldHasListBox(value) => {
                            set_once!(builder.has_list_box, *value, "ffhaslistbox")
                        },
                        _ => {
                            return Err(RtfError::MalformedDocument(
                                "RTF formfield contains an unsupported control".to_string(),
                            ));
                        },
                    }
                    self.pos += 1;
                },
                Some(Token::Text(text)) => {
                    if !self.decode_transport_text(text)?.trim().is_empty() {
                        return Err(RtfError::MalformedDocument(
                            "RTF formfield contains orphan text".to_string(),
                        ));
                    }
                    self.pos += 1;
                },
                Some(Token::Binary(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF formfield cannot contain binary data".to_string(),
                    ));
                },
                None => return Err(RtfError::UnexpectedEof),
            }
        }
        Err(RtfError::UnexpectedEof)
    }

    fn parse_form_field_text_destination(&mut self) -> RtfResult<String> {
        self.expect_token(Token::OpenBrace)?;
        if matches!(
            self.tokens.get(self.pos),
            Some(Token::Control(ControlWord::IgnorableDestination))
        ) {
            self.pos += 1;
        }
        self.pos += 1; // destination control, classified by caller
        let mut text = String::new();
        let mut unicode_skip = self.current_state()?.unicode_skip.max(0);
        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    return Ok(text.trim_end_matches(['\r', '\n']).to_string());
                },
                Some(Token::Text(value)) => {
                    text.push_str(&self.decode_transport_text(value)?);
                    self.pos += 1;
                },
                Some(Token::Control(ControlWord::Unicode(first))) => {
                    text.push_str(&self.parse_style_unicode(*first, unicode_skip)?);
                },
                Some(Token::Control(ControlWord::UnicodeSkip(value))) => {
                    unicode_skip = (*value).max(0);
                    self.pos += 1;
                },
                Some(Token::Control(control)) if control_symbol_text(control).is_some() => {
                    text.push_str(control_symbol_text(control).unwrap_or_default());
                    self.pos += 1;
                },
                Some(Token::OpenBrace | Token::Binary(_)) | Some(Token::Control(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF form-field text contains active, nested, or binary data".to_string(),
                    ));
                },
                None => return Err(RtfError::UnexpectedEof),
            }
            if text.len() > super::form_field::MAX_FORM_FIELD_STRING_BYTES {
                return Err(RtfError::MalformedDocument(
                    "RTF form-field string exceeds the safety limit".to_string(),
                ));
            }
        }
    }

    fn parse_data_field_destination(&mut self) -> RtfResult<Vec<u8>> {
        self.expect_token(Token::OpenBrace)?;
        if matches!(
            self.tokens.get(self.pos),
            Some(Token::Control(ControlWord::IgnorableDestination))
        ) {
            self.pos += 1;
        }
        if !matches!(
            self.tokens.get(self.pos),
            Some(Token::Control(ControlWord::DataField))
        ) {
            return Err(RtfError::MalformedDocument(
                "invalid RTF datafield destination".to_string(),
            ));
        }
        self.pos += 1;
        let mut high = None;
        let mut data = Vec::new();
        loop {
            match self.tokens.get(self.pos) {
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    if high.is_some() {
                        return Err(RtfError::MalformedDocument(
                            "RTF datafield has an odd hexadecimal digit count".to_string(),
                        ));
                    }
                    return Ok(data);
                },
                Some(Token::Text(text)) => {
                    for byte in text.as_bytes() {
                        if byte.is_ascii_whitespace() {
                            continue;
                        }
                        let nibble = match byte {
                            b'0'..=b'9' => byte - b'0',
                            b'a'..=b'f' => byte - b'a' + 10,
                            b'A'..=b'F' => byte - b'A' + 10,
                            _ => {
                                return Err(RtfError::MalformedDocument(
                                    "RTF datafield contains a non-hexadecimal character"
                                        .to_string(),
                                ));
                            },
                        };
                        if let Some(first) = high.take() {
                            data.push(first << 4 | nibble);
                        } else {
                            high = Some(nibble);
                        }
                    }
                    self.pos += 1;
                },
                Some(Token::OpenBrace | Token::Binary(_)) | Some(Token::Control(_)) => {
                    return Err(RtfError::MalformedDocument(
                        "RTF datafield cannot contain controls, nesting, or binary data"
                            .to_string(),
                    ));
                },
                None => return Err(RtfError::UnexpectedEof),
            }
            if data.len() > super::form_field::MAX_FORM_FIELD_DATA_BYTES {
                return Err(RtfError::MalformedDocument(
                    "RTF datafield exceeds the safety limit".to_string(),
                ));
            }
        }
    }

    /// Parse header or footer content.
    fn parse_header_footer_content(&mut self) -> RtfResult<()> {
        let hf_type = self
            .current_hf_type
            .ok_or_else(|| RtfError::MalformedDocument("Header/footer type not set".to_string()))?;

        let mut hf = super::section::HeaderFooter::new(hf_type);
        let mut text_buffer = SmallVec::<[u8; 256]>::new();
        let default_state = State::default();
        let mut inert_section_format = false;

        while self.pos < self.tokens.len() {
            match &self.tokens[self.pos] {
                Token::CloseBrace => {
                    if !text_buffer.is_empty() {
                        if let Ok(text) = std::str::from_utf8(&text_buffer) {
                            let state = self.current_state().ok().unwrap_or(&default_state);
                            let text_alloc = self.arena.alloc_str(text);
                            let para = super::section::HeaderFooterParagraph::new(
                                Cow::Borrowed(text_alloc),
                                state.formatting,
                                state.paragraph,
                            );
                            hf.add_paragraph(para);
                        }
                        text_buffer.clear();
                    }
                    self.pos += 1;
                    break;
                },
                Token::OpenBrace => {
                    inert_section_format = false;
                    if !text_buffer.is_empty() {
                        if let Ok(text) = std::str::from_utf8(&text_buffer) {
                            let state = self.current_state().ok().unwrap_or(&default_state);
                            let text_alloc = self.arena.alloc_str(text);
                            let para = super::section::HeaderFooterParagraph::new(
                                Cow::Borrowed(text_alloc),
                                state.formatting,
                                state.paragraph,
                            );
                            hf.add_paragraph(para);
                        }
                        text_buffer.clear();
                    }
                    self.parse_group()?;
                },
                Token::Control(ControlWord::Par | ControlWord::Line) => {
                    inert_section_format = false;
                    self.pos += 1;
                    if !text_buffer.is_empty() {
                        if let Ok(text) = std::str::from_utf8(&text_buffer) {
                            let state = self.current_state().ok().unwrap_or(&default_state);
                            let text_alloc = self.arena.alloc_str(text);
                            let para = super::section::HeaderFooterParagraph::new(
                                Cow::Borrowed(text_alloc),
                                state.formatting,
                                state.paragraph,
                            );
                            hf.add_paragraph(para);
                        }
                        text_buffer.clear();
                    }
                },
                Token::Control(ControlWord::Tab) => {
                    inert_section_format = false;
                    self.pos += 1;
                    text_buffer.push(b'\t');
                },
                Token::Control(ControlWord::Unicode(code)) => {
                    inert_section_format = false;
                    let decoded = self.parse_destination_unicode_sequence(*code)?;
                    text_buffer.extend_from_slice(decoded.as_bytes());
                },
                Token::Control(control) if control_symbol_text(control).is_some() => {
                    inert_section_format = false;
                    self.pos += 1;
                    text_buffer.extend_from_slice(
                        control_symbol_text(control).unwrap_or_default().as_bytes(),
                    );
                },
                Token::Control(ControlWord::SectionDefault) => {
                    self.pos += 1;
                    inert_section_format = true;
                },
                Token::Control(ControlWord::SectionBreak) if inert_section_format => {
                    self.pos += 1;
                    inert_section_format = false;
                },
                Token::Control(control)
                    if inert_section_format && is_section_control(control) =>
                {
                    self.pos += 1;
                },
                Token::Control(control) => {
                    self.pos += 1;
                    self.apply_control_word(control)?;
                },
                Token::Text(text) => {
                    let decoded = self.decode_transport_text(text)?;
                    self.pos += 1;
                    if !decoded.is_empty() {
                        inert_section_format = false;
                    }
                    text_buffer.extend_from_slice(decoded.as_bytes());
                },
                _ => {
                    self.pos += 1;
                },
            }
        }

        // Add header/footer to the current section or create a new section
        if let Some(section) = self.sections.last_mut() {
            section.add_header_footer(hf);
        } else {
            let mut section = super::section::Section::new();
            section.add_header_footer(hf);
            self.sections.push(section);
        }

        self.current_hf_type = None;
        Ok(())
    }

    fn parse_destination_unicode_sequence(&mut self, first_code: i32) -> RtfResult<String> {
        let skip_count = self.current_state()?.unicode_skip.max(0) as usize;
        let mut utf16 = SmallVec::<[u16; 4]>::new();
        utf16.push(first_code as u16);
        self.pos += 1;
        while let Some(Token::Control(ControlWord::Unicode(code))) = self.tokens.get(self.pos) {
            utf16.push(*code as u16);
            self.pos += 1;
        }

        let mut fallback_skip = skip_count.saturating_mul(utf16.len());
        let mut remainder = String::new();
        while fallback_skip > 0 && self.pos < self.tokens.len() {
            match self.tokens.get(self.pos) {
                Some(Token::Text(text)) => {
                    let count = text.chars().count();
                    if count <= fallback_skip {
                        fallback_skip -= count;
                    } else {
                        remainder.extend(text.chars().skip(fallback_skip));
                        fallback_skip = 0;
                    }
                    self.pos += 1;
                },
                Some(Token::Control(ControlWord::Unicode(_))) => break,
                Some(_) => {
                    fallback_skip = fallback_skip.saturating_sub(1);
                    self.pos += 1;
                },
                None => break,
            }
        }
        let mut decoded = String::from_utf16(&utf16).map_err(|error| {
            RtfError::InvalidUnicode(format!("invalid destination Unicode: {error}"))
        })?;
        decoded.push_str(&self.decode_transport_text(&remainder)?);
        Ok(decoded)
    }

    /// Parse footnote or endnote content.
    fn parse_note(&mut self, is_footnote: bool) -> RtfResult<()> {
        self.current_note_buffer.clear();
        let mut reference = String::from(if is_footnote { "1" } else { "i" });

        while self.pos < self.tokens.len() {
            match &self.tokens[self.pos] {
                Token::CloseBrace => {
                    self.pos += 1;
                    break;
                },
                Token::OpenBrace => {
                    self.parse_group()?;
                },
                Token::Control(ControlWord::FootnoteNumber(n)) => {
                    self.pos += 1;
                    reference = n.to_string();
                },
                Token::Control(ControlWord::Tab) => {
                    self.pos += 1;
                    self.current_note_buffer.push(b'\t');
                },
                Token::Control(control) if control_symbol_text(control).is_some() => {
                    self.pos += 1;
                    self.current_note_buffer.extend_from_slice(
                        control_symbol_text(control).unwrap_or_default().as_bytes(),
                    );
                },
                Token::Control(control) => {
                    self.pos += 1;
                    self.apply_control_word(control)?;
                },
                Token::Text(text) => {
                    let decoded = self.decode_transport_text(text)?;
                    self.pos += 1;
                    self.current_note_buffer
                        .extend_from_slice(decoded.as_bytes());
                },
                _ => {
                    self.pos += 1;
                },
            }
        }

        if !self.current_note_buffer.is_empty()
            && let Ok(content) = std::str::from_utf8(&self.current_note_buffer)
        {
            let content_alloc = self.arena.alloc_str(content);
            let mut note = if is_footnote {
                super::section::Note::footnote(Cow::Owned(reference), Cow::Borrowed(content_alloc))
            } else {
                super::section::Note::endnote(Cow::Owned(reference), Cow::Borrowed(content_alloc))
            };

            if let Ok(state) = self.current_state() {
                note.formatting = state.formatting;
            }

            self.notes.push(note);
        }

        Ok(())
    }
}

/// Parsed RTF document.
///
/// This is an intermediate representation produced by the parser
/// before being converted into the final `RtfDocument` structure.
/// All fields are public to allow direct access during document construction.
pub struct ParsedDocument<'a> {
    /// Font table
    pub font_table: FontTable<'a>,
    pub file_table: Option<crate::FileTable<'a>>,
    /// Color table
    pub color_table: ColorTable,
    /// Style blocks
    pub blocks: Vec<StyleBlock<'a>>,
    /// Extracted tables
    pub tables: Vec<super::table::Table<'a>>,
    /// Extracted pictures
    pub pictures: Vec<super::picture::Picture<'a>>,
    /// Extracted fields
    pub fields: Vec<super::field::Field<'a>>,
    pub form_fields: Vec<super::form_field::FormField<'a>>,
    pub generator: Option<crate::DocumentGenerator<'a>>,
    pub revision_save: Option<crate::RevisionSaveMetadata>,
    pub xml_namespaces: Vec<crate::XmlNamespace<'a>>,
    pub saw_xml_namespace_table: bool,
    pub theme: Option<crate::DocumentTheme<'a>>,
    pub latent_styles: Option<crate::LatentStyles<'a>>,
    pub data_store: Option<crate::DocumentDataStore<'a>>,
    pub math_properties: Option<crate::DocumentMathProperties>,
    pub language_defaults: crate::DocumentLanguageDefaults,
    pub document_direction: Option<crate::TextDirection>,
    pub gutter_on_right: bool,
    /// Embedded and linked objects
    pub objects: Vec<super::object::EmbeddedObject<'a>>,
    /// Ordered inert document-variable metadata
    pub document_variables: Vec<DocumentVariable<'a>>,
    /// Ordered inert user-defined document properties
    pub user_properties: Vec<UserProperty<'a>>,
    /// Ordered inert index and table-of-contents source marks.
    pub navigation_entries: Vec<NavigationEntry<'a>>,
    /// Ordered inert generated list markers.
    pub generated_list_markers: Vec<crate::GeneratedListMarker<'a>>,
    /// List table
    pub list_table: super::list::ListTable<'a>,
    /// List override table
    pub list_override_table: super::list::ListOverrideTable,
    pub legacy_section_numbering: crate::LegacySectionNumbering<'a>,
    pub paragraph_group_table: Option<crate::ParagraphGroupPropertyTable>,
    /// Sections
    pub sections: Vec<super::section::Section<'a>>,
    /// Bookmarks
    pub bookmarks: super::bookmark::BookmarkTable<'a>,
    /// Shapes
    pub shapes: Vec<super::shape::Shape<'a>>,
    /// Inert legacy drawing text boxes.
    pub legacy_text_boxes: Vec<crate::LegacyTextBox<'a>>,
    /// Shape groups
    pub shape_groups: Vec<super::shape::ShapeGroup<'a>>,
    /// Stylesheet
    pub stylesheet: super::stylesheet::StyleSheet<'a>,
    /// Document information
    pub info: super::info::DocumentInfo<'a>,
    /// Annotations
    pub annotations: Vec<super::annotation::Annotation<'a>>,
    /// Footnotes and endnotes
    pub notes: Vec<super::section::Note<'a>>,
    /// Explicit document-level footnote and endnote configuration.
    pub note_options: crate::NoteOptions,
    pub note_separators: crate::NoteSeparatorTable<'a>,
    /// Track changes/revisions
    pub revisions: Vec<super::annotation::Revision<'a>>,
    /// Ordered inert revision-author table.
    pub revision_authors: Vec<super::annotation::RevisionAuthor<'a>>,
}
