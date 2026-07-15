/// Paragraph Properties (PAP) parser for DOC files.
///
/// PAP structures define paragraph-level formatting such as:
/// - Alignment (left, right, center, justified)
/// - Indentation (left, right, first line)
/// - Spacing (before, after, line spacing)
/// - Borders and shading
/// - Tab stops
/// - Table nesting information
///
/// Based on Apache POI's ParagraphSprmUncompressor and ParagraphProperties.
use super::super::package::{DocError, Result};
use super::styles::StyleSheet;
use super::tap::TableProperties;
pub use super::tap::{CellShading as Shading, ShadingPattern};
use crate::sprm::{Sprm, parse_sprms};
use crate::sprm_operations::*;
use litchi_core::binary::{read_i16_le, read_u16_le};

/// Paragraph Properties structure.
///
/// Contains formatting information for a paragraph.
/// Based on Apache POI's ParagraphProperties implementation.
#[derive(Debug, Clone, Default)]
pub struct ParagraphProperties {
    /// Justification/alignment
    pub justification: Justification,
    /// Left indent in twips (1/1440 inch)
    pub indent_left: Option<i32>,
    /// Right indent in twips
    pub indent_right: Option<i32>,
    /// First line indent in twips
    pub indent_first_line: Option<i32>,
    /// Logical left indent in hundredths of a character
    pub indent_left_chars: Option<i16>,
    /// Logical right indent in hundredths of a character
    pub indent_right_chars: Option<i16>,
    /// First-line indent in hundredths of a character
    pub indent_first_line_chars: Option<i16>,
    /// Space before paragraph in twips
    pub space_before: Option<u16>,
    /// Space after paragraph in twips
    pub space_after: Option<u16>,
    /// Space before paragraph in hundredths of a line
    pub space_before_lines: Option<i16>,
    /// Space after paragraph in hundredths of a line
    pub space_after_lines: Option<i16>,
    /// Whether automatic space-before is enabled
    pub space_before_auto: bool,
    /// Whether automatic space-after is enabled
    pub space_after_auto: bool,
    /// Line spacing value
    pub line_spacing: Option<i16>,
    /// Line spacing type
    pub line_spacing_type: LineSpacingType,
    /// Keep paragraph on one page
    pub keep_on_page: bool,
    /// Keep with next paragraph
    pub keep_with_next: bool,
    /// Page break before paragraph
    pub page_break_before: bool,
    /// Widow/orphan control
    pub widow_control: bool,
    /// Side-by-side paragraphs
    pub side_by_side: bool,
    /// No line numbering
    pub no_line_numbering: bool,
    /// No auto hyphenation
    pub no_auto_hyph: bool,
    /// Prevent overlapping floating objects anchored to this paragraph
    pub no_allow_overlap: bool,
    /// Suppress spacing between paragraphs with the same style
    pub contextual_spacing: bool,
    /// Mirror left and right indents on facing pages
    pub mirror_indents: bool,
    /// Tight-wrap mode for text boxes
    pub text_box_tight_wrap: Option<TextBoxTightWrap>,
    /// Tab stops
    pub tab_stops: Vec<TabStop>,
    /// Borders
    pub borders: Borders,
    /// Background shading
    pub shading: Option<Shading>,
    /// Paragraph is inside a table
    pub in_table: bool,
    /// Paragraph is a table row end marker
    pub is_table_row_end: bool,
    /// Table nesting level (itap: 0 = not in table, 1+ = nested level)
    pub table_nesting_level: i32,
    /// Inner table cell flag
    pub inner_table_cell: bool,
    /// Inner table row end flag
    pub inner_table_row_end: bool,
    /// This paragraph is terminated by a top-level table cell mark.
    pub is_table_cell_end: bool,
    /// The table cell mark remained displayed immediately after a nested table
    pub open_table_cell_mark: bool,
    /// Parsed row-level TAP properties when table SPRMs are present.
    pub table_properties: Option<TableProperties>,
    /// Outline level (0-9, where 0-8 are heading levels)
    pub outline_level: Option<u8>,
    /// Style index (istd)
    pub style_index: Option<u16>,
    /// List level (ilvl)
    pub list_level: Option<u8>,
    /// List format override index (ilfo)
    pub list_format_override: Option<i16>,
    /// Bi-directional paragraph
    pub bi_directional: bool,
    /// Whether the paragraph follows vertical document-grid settings
    pub use_page_setup_settings: Option<bool>,
    /// Whether the right indent adjusts automatically to the document grid
    pub adjust_right_indent: Option<bool>,
    /// Locked paragraph
    pub locked: bool,
    /// Kinsoku (Asian typography)
    pub kinsoku: bool,
    /// Word wrap
    pub word_wrap: bool,
    /// Overflow punctuation (Asian)
    pub overflow_punct: bool,
    /// Top line punctuation (Asian)
    pub top_line_punct: bool,
    /// Auto space DE (Asian)
    pub auto_space_de: bool,
    /// Auto space DN (Asian)
    pub auto_space_dn: bool,
    /// Font alignment
    pub font_align: Option<u16>,
    /// Frame text flow
    pub frame_text_flow: Option<u16>,
    /// Absolute horizontal position (for positioned paragraphs)
    pub dxa_abs: Option<i16>,
    /// Absolute vertical position
    pub dya_abs: Option<i16>,
    /// Absolute width
    pub dxa_width: Option<i16>,
    /// Row height (for table rows)
    pub row_height: Option<u16>,
    /// Text wrapping
    pub text_wrap: Option<u8>,
    /// Horizontal distance from text
    pub dxa_from_text: Option<i16>,
    /// Vertical distance from text
    pub dya_from_text: Option<i16>,
    /// Whether this paragraph has a tracked property change.
    pub has_formatting_revision: Option<bool>,
    /// Paragraph revision author index in `SttbfRMark`.
    pub formatting_revision_author_index: Option<u16>,
    /// Packed paragraph revision DTTM.
    pub formatting_revision_timestamp: Option<u32>,
    /// Whether a numbered list was applied after the previous revision.
    pub numbering_revision_list_applied: Option<bool>,
    /// Numbering state retained by a numbering revision mark.
    pub numbering_revision: Option<NumberingRevisionProperties>,
    /// Whether the containing table row has a tracked property change.
    pub has_table_formatting_revision: Option<bool>,
    /// Table-row revision author index in `SttbfRMark`.
    pub table_formatting_revision_author_index: Option<u16>,
    /// Packed table-row revision DTTM.
    pub table_formatting_revision_timestamp: Option<u32>,
    /// Whether table properties before the tracked change are preserved.
    pub table_properties_preserved_for_revision: bool,
    /// Whether paragraph properties before a tracked change are preserved.
    pub properties_preserved_for_revision: bool,
    /// Nonzero `PGPInfo.ipgpSelf` associated with this paragraph.
    pub paragraph_group_id: Option<u32>,
    /// Revision save ID associated with paragraph formatting.
    pub revision_save_id: Option<u32>,
}

/// Parsed `NumRM` numbering revision state.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NumberingRevisionProperties {
    /// Whether the paragraph was already numbered when tracking began.
    pub was_numbered: bool,
    /// Revision author index in `SttbfRMark`.
    pub author_index: u16,
    /// Packed revision DTTM.
    pub timestamp: u32,
    /// Placeholder positions for the nine numbering levels.
    pub placeholder_positions: [u8; 9],
    /// MSONFC values for the nine numbering levels.
    pub number_formats: [u8; 9],
    /// Numeric values for the nine numbering levels.
    pub numbers: [u32; 9],
    /// Numbering format string.
    pub format_string: String,
}

/// Paragraph justification/alignment.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum Justification {
    /// Left aligned
    #[default]
    Left,
    /// Center aligned
    Center,
    /// Right aligned
    Right,
    /// Justified (full width)
    Justified,
    /// Distributed (Asian typography)
    Distributed,
    /// Medium Kashida or medium character compression.
    MediumKashida,
    /// Indented justification.
    Indented,
    /// High Kashida or high character compression.
    HighKashida,
    /// Low Kashida or high character compression.
    LowKashida,
    /// Thai distributed or low character compression.
    ThaiDistributed,
}

impl TryFrom<u8> for Justification {
    type Error = u8;

    fn try_from(value: u8) -> std::result::Result<Self, Self::Error> {
        match value {
            0 => Ok(Self::Left),
            1 => Ok(Self::Center),
            2 => Ok(Self::Right),
            3 => Ok(Self::Justified),
            4 => Ok(Self::Distributed),
            5 => Ok(Self::MediumKashida),
            6 => Ok(Self::Indented),
            7 => Ok(Self::HighKashida),
            8 => Ok(Self::LowKashida),
            9 => Ok(Self::ThaiDistributed),
            invalid => Err(invalid),
        }
    }
}

/// Line spacing type.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum LineSpacingType {
    /// Single line spacing (lspd.fMultLineSp = 1, lspd.dyaLine = 240)
    #[default]
    Single,
    /// 1.5 line spacing (lspd.fMultLineSp = 1, lspd.dyaLine = 360)
    OnePointFive,
    /// Double line spacing (lspd.fMultLineSp = 1, lspd.dyaLine = 480)
    Double,
    /// At least N twips (lspd.fMultLineSp = 0, lspd.dyaLine > 0)
    AtLeast,
    /// Exactly N twips (lspd.fMultLineSp = 0, lspd.dyaLine < 0)
    Exactly,
    /// Multiple (value in 240ths of a line) (lspd.fMultLineSp = 1)
    Multiple,
}

/// Lines in a text box whose edges permit tight wrapping by surrounding text.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u8)]
pub enum TextBoxTightWrap {
    /// No lines permit tight wrapping.
    None = 0,
    /// All lines permit tight wrapping.
    AllLines = 1,
    /// Only the first and last lines permit tight wrapping.
    FirstAndLastLine = 2,
    /// Only the first line permits tight wrapping.
    FirstLineOnly = 3,
    /// Only the last line permits tight wrapping.
    LastLineOnly = 4,
}

impl TryFrom<u8> for TextBoxTightWrap {
    type Error = u8;

    fn try_from(value: u8) -> std::result::Result<Self, Self::Error> {
        match value {
            0 => Ok(Self::None),
            1 => Ok(Self::AllLines),
            2 => Ok(Self::FirstAndLastLine),
            3 => Ok(Self::FirstLineOnly),
            4 => Ok(Self::LastLineOnly),
            invalid => Err(invalid),
        }
    }
}

/// Tab stop definition.
#[derive(Debug, Clone, Copy)]
pub struct TabStop {
    /// Position in twips
    pub position: i32,
    /// Tab alignment
    pub alignment: TabAlignment,
    /// Leader character
    pub leader: TabLeader,
}

/// Tab alignment.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TabAlignment {
    /// Left aligned
    Left,
    /// Center aligned
    Center,
    /// Right aligned
    Right,
    /// Decimal aligned
    Decimal,
    /// Bar (vertical line)
    Bar,
}

/// Tab leader characters.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TabLeader {
    /// No leader
    None,
    /// Dots
    Dots,
    /// Hyphens
    Hyphens,
    /// Underline
    Underline,
    /// Heavy line
    Heavy,
    /// Middle dot
    MiddleDot,
}

/// Paragraph borders.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Borders {
    /// Top border
    pub top: Option<Border>,
    /// Left border
    pub left: Option<Border>,
    /// Bottom border
    pub bottom: Option<Border>,
    /// Right border
    pub right: Option<Border>,
    /// Between border (for multi-column layouts)
    pub between: Option<Border>,
    /// Bar border
    pub bar: Option<Border>,
}

/// Border definition.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Border {
    /// Border style
    pub style: BorderStyle,
    /// Border width in eighths of a point
    pub width: u8,
    /// Border color (RGB), or `None` for automatic color
    pub color: Option<(u8, u8, u8)>,
    /// Distance from text to the border in points
    pub spacing: u8,
    /// Whether the border has a shadow effect
    pub shadow: bool,
    /// Whether the border has a frame effect
    pub frame: bool,
}

/// Border styles.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BorderStyle {
    None,
    Single,
    Thick,
    Double,
    Dotted,
    Dashed,
    DotDash,
    DotDotDash,
    Triple,
    ThinThickSmallGap,
    ThickThinSmallGap,
    ThinThickThinSmallGap,
    ThinThickMediumGap,
    ThickThinMediumGap,
    ThinThickThinMediumGap,
    ThinThickLargeGap,
    ThickThinLargeGap,
    ThinThickThinLargeGap,
    Wave,
    DoubleWave,
    DashSmallGap,
    DashDotStroked,
    ThreeDEmboss,
    ThreeDEngrave,
    Outset,
    Inset,
}

impl ParagraphProperties {
    /// Create a new ParagraphProperties with default values.
    pub fn new() -> Self {
        Self::default()
    }

    /// Parse paragraph properties from SPRM (Single Property Modifier) data.
    ///
    /// SPRMs are variable-length records that modify properties.
    ///
    /// Based on Apache POI's ParagraphSprmUncompressor.
    ///
    /// # Arguments
    ///
    /// * `grpprl` - Group of SPRMs (property modifications)
    pub fn from_sprm(grpprl: &[u8]) -> Result<Self> {
        Self::from_sprm_context(grpprl, None)
    }

    pub(crate) fn from_sprm_with_stylesheet(
        grpprl: &[u8],
        stylesheet: &StyleSheet,
    ) -> Result<Self> {
        Self::from_sprm_context(grpprl, Some(stylesheet))
    }

    fn from_sprm_context(grpprl: &[u8], stylesheet: Option<&StyleSheet>) -> Result<Self> {
        let mut pap = Self::default();
        let sprms = parse_sprms(grpprl);
        let consumed = sprms.last().map_or(0, |sprm| sprm.offset + sprm.size);
        if consumed != grpprl.len() {
            return Err(DocError::Corrupted(
                "PAP grpprl does not contain a whole number of SPRMs".to_string(),
            ));
        }

        for sprm in &sprms {
            // Only process PAP SPRMs (type = 1)
            if get_sprm_type(sprm.opcode) == 1 {
                Self::apply_sprm(&mut pap, sprm)?;
            } else if get_sprm_type(sprm.opcode) == 5 {
                Self::apply_table_revision_sprm(&mut pap, sprm)?;
            }
        }

        if sprms.iter().any(|sprm| get_sprm_type(sprm.opcode) == 5) {
            let arena = bumpalo::Bump::new();
            let parser = super::tap_parser::TapParser::new(&arena);
            pap.table_properties = Some(if let Some(stylesheet) = stylesheet {
                parser.parse_tap_with_stylesheet(grpprl, stylesheet)?
            } else {
                parser.parse_tap(grpprl)?
            });
        }

        Ok(pap)
    }

    pub(crate) fn cascade_styles(
        initial_style_index: Option<u16>,
        direct_sprms: &[u8],
        stylesheet: &StyleSheet,
    ) -> Result<Self> {
        let mut current = Self::paragraph_style_baseline(initial_style_index, stylesheet)?;
        let sprms = parse_sprms(direct_sprms);
        let consumed = sprms.last().map_or(0, |sprm| sprm.offset + sprm.size);
        if consumed != direct_sprms.len() {
            return Err(DocError::Corrupted(
                "PAPX grpprl does not contain a whole number of SPRMs".to_string(),
            ));
        }
        for sprm in &sprms {
            if get_sprm_type(sprm.opcode) != 1 {
                continue;
            }
            if sprm.opcode == 0x4600 {
                let requested = sprm.operand_word().ok_or_else(|| {
                    DocError::Corrupted("sprmPIstd is missing its style index".to_string())
                })?;
                let mut styled = Self::paragraph_style_baseline(Some(requested), stylesheet)?;
                styled.style_index = Some(requested);
                Self::preserve_style_state(&current, &mut styled);
                current = styled;
            } else {
                Self::apply_sprm(&mut current, sprm)?;
            }
        }

        let table_state = Self::from_sprm_with_stylesheet(direct_sprms, stylesheet)?;
        current.table_properties = table_state.table_properties;
        current.has_table_formatting_revision = table_state.has_table_formatting_revision;
        current.table_formatting_revision_author_index =
            table_state.table_formatting_revision_author_index;
        current.table_formatting_revision_timestamp =
            table_state.table_formatting_revision_timestamp;
        current.table_properties_preserved_for_revision =
            table_state.table_properties_preserved_for_revision;
        Ok(current)
    }

    fn paragraph_style_baseline(style_index: Option<u16>, stylesheet: &StyleSheet) -> Result<Self> {
        let Some(requested) = style_index else {
            return Ok(Self::default());
        };
        let (effective, paragraph, _) = stylesheet.resolve_paragraph_style_sprms(requested)?;
        let mut baseline = Self::from_sprm(&paragraph)?;
        baseline.style_index = Some(requested);
        if effective.is_some() && (1..=9).contains(&requested) {
            baseline.outline_level = Some((requested - 1) as u8);
        }
        Ok(baseline)
    }

    fn preserve_style_state(previous: &Self, styled: &mut Self) {
        styled.in_table = previous.in_table;
        styled.is_table_row_end = previous.is_table_row_end;
        styled.table_nesting_level = previous.table_nesting_level;
        styled.inner_table_cell = previous.inner_table_cell;
        styled.inner_table_row_end = previous.inner_table_row_end;
        styled.is_table_cell_end = previous.is_table_cell_end;
        styled.open_table_cell_mark = previous.open_table_cell_mark;
        styled.table_properties = previous.table_properties.clone();
        styled.paragraph_group_id = previous.paragraph_group_id;
        styled.properties_preserved_for_revision = previous.properties_preserved_for_revision;
        styled.revision_save_id = previous.revision_save_id;
        styled.has_formatting_revision = previous.has_formatting_revision;
        styled.formatting_revision_author_index = previous.formatting_revision_author_index;
        styled.formatting_revision_timestamp = previous.formatting_revision_timestamp;
        styled.numbering_revision_list_applied = previous.numbering_revision_list_applied;
        styled.numbering_revision = previous.numbering_revision.clone();
    }

    fn apply_table_revision_sprm(pap: &mut ParagraphProperties, sprm: &Sprm) -> Result<()> {
        match sprm.opcode {
            0xD667 => {
                let operand = sprm.operand_bytes();
                if operand.len() != 7 {
                    return Err(DocError::Corrupted(
                        "sprmTPropRMark operand must contain exactly 7 bytes".to_string(),
                    ));
                }
                pap.has_table_formatting_revision = Some(match operand[0] {
                    0 => false,
                    1 => true,
                    _ => {
                        return Err(DocError::Corrupted(
                            "sprmTPropRMark must begin with a Boolean8 value".to_string(),
                        ));
                    },
                });
                let author = i16::from_le_bytes([operand[1], operand[2]]);
                pap.table_formatting_revision_author_index =
                    Some(u16::try_from(author).map_err(|_| {
                        DocError::Corrupted("sprmTPropRMark author index is negative".to_string())
                    })?);
                let timestamp =
                    u32::from_le_bytes([operand[3], operand[4], operand[5], operand[6]]);
                crate::doc::revision::decode_dttm(timestamp)?;
                pap.table_formatting_revision_timestamp = Some(timestamp);
            },
            0x3668 => {
                let operand = sprm.operand_bytes();
                if operand.len() != 1 {
                    return Err(DocError::Corrupted(
                        "sprmTWall operand must contain exactly 1 byte".to_string(),
                    ));
                }
                pap.table_properties_preserved_for_revision = match operand[0] {
                    0 => false,
                    1 => true,
                    _ => {
                        return Err(DocError::Corrupted(
                            "sprmTWall must contain a Boolean8 value".to_string(),
                        ));
                    },
                };
            },
            _ => {},
        }
        Ok(())
    }

    /// Apply a single SPRM operation to paragraph properties.
    ///
    /// Based on Apache POI's ParagraphSprmUncompressor.unCompressPAPOperation().
    ///
    /// # Arguments
    ///
    /// * `pap` - The paragraph properties to modify
    /// * `sprm` - The SPRM operation to apply
    fn apply_sprm(pap: &mut ParagraphProperties, sprm: &Sprm) -> Result<()> {
        match sprm.opcode {
            SPRM_P_DXC_RIGHT => {
                pap.indent_right_chars = Some(Self::required_i16(sprm, "sprmPDxcRight")?);
                return Ok(());
            },
            SPRM_P_DXC_LEFT => {
                pap.indent_left_chars = Some(Self::required_i16(sprm, "sprmPDxcLeft")?);
                return Ok(());
            },
            SPRM_P_DXC_LEFT1 => {
                pap.indent_first_line_chars = Some(Self::required_i16(sprm, "sprmPDxcLeft1")?);
                return Ok(());
            },
            SPRM_P_DYL_BEFORE => {
                pap.space_before_lines = Some(Self::line_hundredths(sprm, "sprmPDylBefore")?);
                return Ok(());
            },
            SPRM_P_DYL_AFTER => {
                pap.space_after_lines = Some(Self::line_hundredths(sprm, "sprmPDylAfter")?);
                return Ok(());
            },
            SPRM_P_F_OPEN_TCH => {
                pap.open_table_cell_mark = Self::strict_bool8(sprm, "sprmPFOpenTch")?;
                return Ok(());
            },
            SPRM_P_F_DYA_BEFORE_AUTO => {
                pap.space_before_auto = Self::strict_bool8(sprm, "sprmPFDyaBeforeAuto")?;
                return Ok(());
            },
            SPRM_P_F_DYA_AFTER_AUTO => {
                pap.space_after_auto = Self::strict_bool8(sprm, "sprmPFDyaAfterAuto")?;
                return Ok(());
            },
            SPRM_P_DXA_RIGHT_2000 => {
                pap.indent_right = Some(i32::from(Self::required_i16(sprm, "sprmPDxaRight")?));
                return Ok(());
            },
            SPRM_P_DXA_LEFT_2000 => {
                pap.indent_left = Some(i32::from(Self::required_i16(sprm, "sprmPDxaLeft")?));
                return Ok(());
            },
            SPRM_P_NEST_2000 => {
                let delta = i32::from(Self::required_i16(sprm, "sprmPNest")?);
                pap.indent_left = Some(pap.indent_left.unwrap_or(0) + delta);
                return Ok(());
            },
            SPRM_P_DXA_LEFT1_2000 => {
                pap.indent_first_line = Some(i32::from(Self::required_i16(sprm, "sprmPDxaLeft1")?));
                return Ok(());
            },
            SPRM_P_JC_LOGICAL => {
                let code = sprm.operand_byte().ok_or_else(|| {
                    DocError::Corrupted("sprmPJc is missing its justification".to_string())
                })?;
                pap.justification = Justification::try_from(code).map_err(|invalid| {
                    DocError::Corrupted(format!(
                        "sprmPJc has invalid logical justification {invalid}"
                    ))
                })?;
                return Ok(());
            },
            SPRM_P_BRC_TOP => {
                pap.borders.top = Self::parse_current_border(sprm)?;
                return Ok(());
            },
            SPRM_P_BRC_LEFT => {
                pap.borders.left = Self::parse_current_border(sprm)?;
                return Ok(());
            },
            SPRM_P_BRC_BOTTOM => {
                pap.borders.bottom = Self::parse_current_border(sprm)?;
                return Ok(());
            },
            SPRM_P_BRC_RIGHT => {
                pap.borders.right = Self::parse_current_border(sprm)?;
                return Ok(());
            },
            SPRM_P_BRC_BETWEEN => {
                pap.borders.between = Self::parse_current_border(sprm)?;
                return Ok(());
            },
            SPRM_P_BRC_BAR => {
                pap.borders.bar = Self::parse_current_border(sprm)?;
                return Ok(());
            },
            SPRM_P_F_NO_ALLOW_OVERLAP => {
                pap.no_allow_overlap = Self::strict_bool8(sprm, "sprmPFNoAllowOverlap")?;
                return Ok(());
            },
            SPRM_P_WALL => {
                pap.properties_preserved_for_revision = Self::strict_bool8(sprm, "sprmPWall")?;
                return Ok(());
            },
            SPRM_P_IPGP => {
                let group_id = sprm.operand_dword().ok_or_else(|| {
                    DocError::Corrupted("sprmPIpgp is missing its PGPInfo index".to_string())
                })?;
                if group_id == 0 {
                    return Err(DocError::Corrupted(
                        "sprmPIpgp must contain a nonzero PGPInfo index".to_string(),
                    ));
                }
                pap.paragraph_group_id = Some(group_id);
                return Ok(());
            },
            SPRM_P_RSID => {
                pap.revision_save_id = Some(sprm.operand_dword().ok_or_else(|| {
                    DocError::Corrupted("sprmPRsid is missing its revision save ID".to_string())
                })?);
                return Ok(());
            },
            SPRM_P_F_CONTEXTUAL_SPACING => {
                pap.contextual_spacing = Self::strict_bool8(sprm, "sprmPFContextualSpacing")?;
                return Ok(());
            },
            SPRM_P_F_MIRROR_INDENTS => {
                pap.mirror_indents = Self::strict_bool8(sprm, "sprmPFMirrorIndents")?;
                return Ok(());
            },
            SPRM_P_TTWO => {
                let tight_wrap = sprm.operand_byte().ok_or_else(|| {
                    DocError::Corrupted("sprmPTtwo is missing its tight-wrap mode".to_string())
                })?;
                pap.text_box_tight_wrap =
                    Some(TextBoxTightWrap::try_from(tight_wrap).map_err(|invalid| {
                        DocError::Corrupted(format!(
                            "sprmPTtwo has invalid tight-wrap mode {invalid}"
                        ))
                    })?);
                return Ok(());
            },
            _ => {},
        }
        let operation = get_sprm_operation(sprm.opcode);

        match operation {
            // Operation 0x00: sprmPIstd - Paragraph style
            0x00 => {
                if let Some(istd) = sprm.operand_word() {
                    pap.style_index = Some(istd);
                }
            },
            // Operation 0x01: sprmPIstdPermute - Style permutation
            0x01 => {
                // Used only for piece table grpprl's, not for PAPX
            },
            // Operation 0x02: sprmPIncLvl - Increment outline level
            0x02 => {
                if let Some(param) = sprm.operand_byte()
                    && pap.style_index.unwrap_or(0) <= 9
                    && pap.style_index.unwrap_or(0) >= 1
                {
                    let param_signed = param as i8;
                    let istd = pap.style_index.unwrap_or(0) as i16 + param_signed as i16;
                    let lvl = pap.outline_level.unwrap_or(0) as i16 + param_signed as i16;

                    pap.style_index = if (param_signed >> 7) & 0x01 == 1 {
                        Some(istd.max(1) as u16)
                    } else {
                        Some(istd.min(9) as u16)
                    };
                    pap.outline_level = Some(lvl as u8);
                }
            },
            // Operation 0x03: sprmPJc - Paragraph justification
            0x03 => {
                let jc = sprm.operand_byte().ok_or_else(|| {
                    DocError::Corrupted("sprmPJc80 is missing its justification".to_string())
                })?;
                if jc > 4 {
                    return Err(DocError::Corrupted(format!(
                        "sprmPJc80 has invalid justification {jc}"
                    )));
                }
                pap.justification = Justification::try_from(jc).expect("values 0 through 4 map");
            },
            // Operation 0x04: sprmPFSideBySide - Side-by-side
            0x04 => {
                if let Some(val) = sprm.operand_byte() {
                    pap.side_by_side = val != 0;
                }
            },
            // Operation 0x05: sprmPFKeep - Keep paragraph intact
            0x05 => {
                if let Some(val) = sprm.operand_byte() {
                    pap.keep_on_page = val != 0;
                }
            },
            // Operation 0x06: sprmPFKeepFollow - Keep with next
            0x06 => {
                if let Some(val) = sprm.operand_byte() {
                    pap.keep_with_next = val != 0;
                }
            },
            // Operation 0x07: sprmPFPageBreakBefore - Page break before
            0x07 => {
                if let Some(val) = sprm.operand_byte() {
                    pap.page_break_before = val != 0;
                }
            },
            // Operation 0x08: sprmPBrcl - Border location
            0x08 => {
                // Border location code - not commonly used
            },
            // Operation 0x09: sprmPBrcp - Border position
            0x09 => {
                // Border position - not commonly used
            },
            // Operation 0x0A: sprmPIlvl - List level
            0x0A => {
                if let Some(ilvl) = sprm.operand_byte() {
                    pap.list_level = Some(ilvl);
                }
            },
            // Operation 0x0B: sprmPIlfo - List format override
            0x0B => {
                if let Some(ilfo) = sprm.operand_i16() {
                    pap.list_format_override = Some(ilfo);
                }
            },
            // Operation 0x0C: sprmPFNoLineNumb - No line numbering
            0x0C => {
                if let Some(val) = sprm.operand_byte() {
                    pap.no_line_numbering = val != 0;
                }
            },
            // Operation 0x0D: sprmPChgTabsPapx - Tab stops
            0x0D => {
                Self::handle_tabs(pap, sprm);
            },
            // Operation 0x0E: sprmPDxaRight - Right indent
            0x0E => {
                if let Some(val) = sprm.operand_i16() {
                    pap.indent_right = Some(val as i32);
                }
            },
            // Operation 0x0F: sprmPDxaLeft - Left indent
            0x0F => {
                if let Some(val) = sprm.operand_i16() {
                    pap.indent_left = Some(val as i32);
                }
            },
            // Operation 0x10: sprmPNest - Nested indent
            0x10 => {
                if let Some(val) = sprm.operand_i16() {
                    let current = pap.indent_left.unwrap_or(0);
                    pap.indent_left = Some((current + val as i32).max(0));
                }
            },
            // Operation 0x11: sprmPDxaLeft1 - First line indent
            0x11 => {
                if let Some(val) = sprm.operand_i16() {
                    pap.indent_first_line = Some(val as i32);
                }
            },
            // Operation 0x12: sprmPDyaLine - Line spacing
            0x12 => {
                if sprm.operand.len() >= 4
                    && let Ok(dya_line) = read_i16_le(&sprm.operand, 0)
                    && let Ok(f_mult) = read_u16_le(&sprm.operand, 2)
                {
                    pap.line_spacing = Some(dya_line);
                    if f_mult != 0 {
                        // Multiple line spacing
                        pap.line_spacing_type = match dya_line {
                            240 => LineSpacingType::Single,
                            360 => LineSpacingType::OnePointFive,
                            480 => LineSpacingType::Double,
                            _ => LineSpacingType::Multiple,
                        };
                    } else if dya_line > 0 {
                        pap.line_spacing_type = LineSpacingType::AtLeast;
                    } else {
                        pap.line_spacing_type = LineSpacingType::Exactly;
                    }
                }
            },
            // Operation 0x13: sprmPDyaBefore - Space before
            0x13 => {
                if let Some(val) = sprm.operand_word() {
                    pap.space_before = Some(val);
                }
            },
            // Operation 0x14: sprmPDyaAfter - Space after
            0x14 => {
                if let Some(val) = sprm.operand_word() {
                    pap.space_after = Some(val);
                }
            },
            // Operation 0x15: sprmPChgTabs - Change tabs (fast saved)
            0x15 => {
                // Fast saved only - not commonly used
            },
            // Operation 0x16: sprmPFInTable - In table flag
            0x16 => {
                if let Some(val) = sprm.operand_byte() {
                    pap.in_table = val != 0;
                }
            },
            // Operation 0x17: sprmPFTtp - Table row end
            0x17 => {
                if let Some(val) = sprm.operand_byte() {
                    pap.is_table_row_end = val != 0;
                }
            },
            // Operation 0x18: sprmPDxaAbs - Absolute horizontal position
            0x18 => {
                if let Some(val) = sprm.operand_i16() {
                    pap.dxa_abs = Some(val);
                }
            },
            // Operation 0x19: sprmPDyaAbs - Absolute vertical position
            0x19 => {
                if let Some(val) = sprm.operand_i16() {
                    pap.dya_abs = Some(val);
                }
            },
            // Operation 0x1A: sprmPDxaWidth - Absolute width
            0x1A => {
                if let Some(val) = sprm.operand_i16() {
                    pap.dxa_width = Some(val);
                }
            },
            // Operation 0x1B: sprmPPc - Positioning code
            0x1B => {
                if let Some(param) = sprm.operand_byte() {
                    let pc_vert = (param & 0x0C) >> 2;
                    let pc_horz = param & 0x03;
                    // Store positioning codes if needed
                    let _ = (pc_vert, pc_horz);
                }
            },
            // Operations 0x1C-0x21: Old border formats (Word 6.0)
            0x1C..=0x21 => {
                // BrcXXX10 - older version borders
            },
            // Operation 0x22: sprmPDxaFromText10 - Distance from text (Word 6.0)
            0x22 => {
                if let Some(val) = sprm.operand_i16() {
                    pap.dxa_from_text = Some(val);
                }
            },
            // Operation 0x23: sprmPWr - Text wrapping
            0x23 => {
                if let Some(val) = sprm.operand_byte() {
                    pap.text_wrap = Some(val);
                }
            },
            // Operations 0x24-0x29: Word 97 Brc80 borders
            0x24 => pap.borders.top = Self::parse_border80(sprm)?,
            0x25 => pap.borders.left = Self::parse_border80(sprm)?,
            0x26 => pap.borders.bottom = Self::parse_border80(sprm)?,
            0x27 => pap.borders.right = Self::parse_border80(sprm)?,
            0x28 => pap.borders.between = Self::parse_border80(sprm)?,
            0x29 => pap.borders.bar = Self::parse_border80(sprm)?,
            // Operation 0x2A: sprmPFNoAutoHyph - No auto hyphenation
            0x2A => {
                if let Some(val) = sprm.operand_byte() {
                    pap.no_auto_hyph = val != 0;
                }
            },
            // Operation 0x2B: sprmPWHeightAbs - Row height (for table rows)
            0x2B => {
                if let Some(val) = sprm.operand_word() {
                    pap.row_height = Some(val);
                }
            },
            // Operation 0x2C: sprmPDcs - Drop cap
            0x2C => {
                // Drop cap specifier - not commonly used
            },
            // Operation 0x2D: sprmPShd80 - Shading (Word 97-2000)
            0x2D => {
                let shd = sprm.operand_word().ok_or_else(|| {
                    DocError::Corrupted("sprmPShd80 is missing its Shd80".to_string())
                })?;
                pap.shading = Self::parse_shd80(shd)?;
            },
            // Operation 0x2E: sprmPDyaFromText - Vertical distance from text
            0x2E => {
                if let Some(val) = sprm.operand_i16() {
                    pap.dya_from_text = Some(val);
                }
            },
            // Operation 0x2F: sprmPDxaFromText - Horizontal distance from text
            0x2F => {
                if let Some(val) = sprm.operand_i16() {
                    pap.dxa_from_text = Some(val);
                }
            },
            // Operation 0x30: sprmPFLocked - Locked paragraph
            0x30 => {
                if let Some(val) = sprm.operand_byte() {
                    pap.locked = val != 0;
                }
            },
            // Operation 0x31: sprmPFWidowControl - Widow/orphan control
            0x31 => {
                if let Some(val) = sprm.operand_byte() {
                    pap.widow_control = val != 0;
                }
            },
            // Operation 0x33: sprmPFKinsoku - Kinsoku
            0x33 => {
                if let Some(val) = sprm.operand_byte() {
                    pap.kinsoku = val != 0;
                }
            },
            // Operation 0x34: sprmPFWordWrap - Word wrap
            0x34 => {
                if let Some(val) = sprm.operand_byte() {
                    pap.word_wrap = val != 0;
                }
            },
            // Operation 0x35: sprmPFOverflowPunct - Overflow punctuation
            0x35 => {
                if let Some(val) = sprm.operand_byte() {
                    pap.overflow_punct = val != 0;
                }
            },
            // Operation 0x36: sprmPFTopLinePunct - Top line punctuation
            0x36 => {
                if let Some(val) = sprm.operand_byte() {
                    pap.top_line_punct = val != 0;
                }
            },
            // Operation 0x37: sprmPFAutoSpaceDE - Auto space DE
            0x37 => {
                if let Some(val) = sprm.operand_byte() {
                    pap.auto_space_de = val != 0;
                }
            },
            // Operation 0x38: sprmPFAutoSpaceDN - Auto space DN
            0x38 => {
                if let Some(val) = sprm.operand_byte() {
                    pap.auto_space_dn = val != 0;
                }
            },
            // Operation 0x39: sprmPWAlignFont - Font alignment
            0x39 => {
                if let Some(val) = sprm.operand_word() {
                    pap.font_align = Some(val);
                }
            },
            // Operation 0x3A: sprmPFrameTextFlow - Frame text flow
            0x3A => {
                if let Some(val) = sprm.operand_word() {
                    pap.frame_text_flow = Some(val);
                }
            },
            // Operation 0x3B: sprmPISnapBaseLine - Snap to baseline
            0x3B => {
                // Not commonly used
            },
            // Operation 0x3E: sprmPAnld - Autonumber list data
            0x3E => {
                // Autonumber list data - complex structure
            },
            // Versioned sprmPPropRMark property revision marks.
            0x3F | 0x65 | 0x6F => Self::apply_property_revision(pap, sprm)?,
            // Operation 0x40: sprmPOutLvl - Outline level
            0x40 => {
                if let Some(lvl) = sprm.operand_byte() {
                    pap.outline_level = Some(lvl);
                }
            },
            // Operation 0x41: sprmPFBiDi - Bi-directional paragraph
            0x41 => {
                if let Some(val) = sprm.operand_byte() {
                    pap.bi_directional = val != 0;
                }
            },
            // Operation 0x43: sprmPFNumRMIns - Numbering revision insert
            0x43 => {
                pap.numbering_revision_list_applied =
                    Some(Self::strict_bool8(sprm, "sprmPFNumRMIns")?);
            },
            // Operation 0x44: sprmPCrLf - CR/LF
            0x44 => {
                // Not commonly used
            },
            // Operation 0x45: sprmPNumRM - Numbering revision mark
            0x45 => pap.numbering_revision = Some(Self::parse_numbering_revision(sprm)?),
            // Operation 0x47: sprmPFUsePgsuSettings - Use page setup settings
            0x47 => {
                pap.use_page_setup_settings =
                    Some(Self::strict_bool8(sprm, "sprmPFUsePgsuSettings")?);
            },
            // Operation 0x48: sprmPFAdjustRight - Adjust right
            0x48 => {
                pap.adjust_right_indent = Some(Self::strict_bool8(sprm, "sprmPFAdjustRight")?);
            },
            // Operation 0x49: sprmPItap - Table nesting level
            0x49 => {
                let depth = Self::required_i32(sprm, "sprmPItap")?;
                if depth < 0 {
                    return Err(DocError::Corrupted(
                        "sprmPItap table depth must be non-negative".to_string(),
                    ));
                }
                pap.table_nesting_level = depth;
            },
            // Operation 0x4A: sprmPDtap - Table nesting delta
            0x4A => {
                let delta = Self::required_i32(sprm, "sprmPDtap")?;
                let depth = pap.table_nesting_level.checked_add(delta).ok_or_else(|| {
                    DocError::Corrupted("sprmPDtap table depth overflowed".to_string())
                })?;
                if depth < 0 {
                    return Err(DocError::Corrupted(
                        "sprmPDtap produced a negative table depth".to_string(),
                    ));
                }
                pap.table_nesting_level = depth;
            },
            // Operation 0x4B: sprmPFInnerTableCell - Inner table cell
            0x4B => {
                let value = Self::strict_bool8(sprm, "sprmPFInnerTableCell")?;
                if value && pap.table_nesting_level <= 1 {
                    return Err(DocError::Corrupted(
                        "sprmPFInnerTableCell requires table depth greater than 1".to_string(),
                    ));
                }
                pap.inner_table_cell = value;
            },
            // Operation 0x4C: sprmPFInnerTtp - Inner table row end
            0x4C => {
                let value = Self::strict_bool8(sprm, "sprmPFInnerTtp")?;
                if value && pap.table_nesting_level <= 1 {
                    return Err(DocError::Corrupted(
                        "sprmPFInnerTtp requires table depth greater than 1".to_string(),
                    ));
                }
                pap.inner_table_row_end = value;
            },
            // Operation 0x4D: sprmPShd - Shading (Word 2002+)
            0x4D => pap.shading = Self::parse_shading_descriptor(sprm)?,
            // Operation 0x67: sprmPRsid - Revision save ID
            0x67 => {
                // Revision save ID - not commonly used
            },
            // Default: Unknown or unsupported SPRM
            _ => {
                // Silently ignore unknown SPRMs
            },
        }
        Ok(())
    }

    fn apply_property_revision(pap: &mut ParagraphProperties, sprm: &Sprm) -> Result<()> {
        let operand = sprm.operand_bytes();
        if operand.len() != 7 {
            return Err(DocError::Corrupted(
                "sprmPPropRMark operand must contain exactly 7 bytes".to_string(),
            ));
        }
        pap.has_formatting_revision = Some(match operand[0] {
            0 => false,
            1 => true,
            _ => {
                return Err(DocError::Corrupted(
                    "sprmPPropRMark must begin with a Boolean8 value".to_string(),
                ));
            },
        });
        let author = i16::from_le_bytes([operand[1], operand[2]]);
        pap.formatting_revision_author_index = Some(u16::try_from(author).map_err(|_| {
            DocError::Corrupted("sprmPPropRMark author index is negative".to_string())
        })?);
        let timestamp = u32::from_le_bytes([operand[3], operand[4], operand[5], operand[6]]);
        crate::doc::revision::decode_dttm(timestamp)?;
        pap.formatting_revision_timestamp = Some(timestamp);
        Ok(())
    }

    fn strict_bool8(sprm: &Sprm, name: &str) -> Result<bool> {
        match sprm.operand_byte() {
            Some(0) => Ok(false),
            Some(1) => Ok(true),
            _ => Err(DocError::Corrupted(format!(
                "{name} must contain a Boolean8 value"
            ))),
        }
    }

    fn required_i16(sprm: &Sprm, name: &str) -> Result<i16> {
        sprm.operand_i16()
            .ok_or_else(|| DocError::Corrupted(format!("{name} is missing its 16-bit operand")))
    }

    fn required_i32(sprm: &Sprm, name: &str) -> Result<i32> {
        let bytes: [u8; 4] = sprm
            .operand_bytes()
            .try_into()
            .map_err(|_| DocError::Corrupted(format!("{name} is missing its 32-bit operand")))?;
        Ok(i32::from_le_bytes(bytes))
    }

    fn line_hundredths(sprm: &Sprm, name: &str) -> Result<i16> {
        let value = Self::required_i16(sprm, name)?;
        if !(-20..=31_680).contains(&value) {
            return Err(DocError::Corrupted(format!(
                "{name} value {value} is outside -20..=31680"
            )));
        }
        Ok(value)
    }

    fn parse_numbering_revision(sprm: &Sprm) -> Result<NumberingRevisionProperties> {
        let operand = sprm.operand_bytes();
        if operand.len() != 128 {
            return Err(DocError::Corrupted(
                "sprmPNumRM operand must contain exactly 128 bytes".to_string(),
            ));
        }
        let was_numbered = match operand[0] {
            0 => false,
            1 => true,
            _ => {
                return Err(DocError::Corrupted(
                    "NumRM.fNumRM must be a Boolean8 value".to_string(),
                ));
            },
        };
        let author = i16::from_le_bytes([operand[2], operand[3]]);
        let author_index = u16::try_from(author)
            .map_err(|_| DocError::Corrupted("NumRM author index is negative".to_string()))?;
        let timestamp = u32::from_le_bytes([operand[4], operand[5], operand[6], operand[7]]);
        let placeholder_positions: [u8; 9] = operand[8..17].try_into().expect("fixed NumRM slice");
        let number_formats: [u8; 9] = operand[17..26].try_into().expect("fixed NumRM slice");
        let numbers = std::array::from_fn(|index| {
            let offset = 28 + index * 4;
            u32::from_le_bytes(
                operand[offset..offset + 4]
                    .try_into()
                    .expect("fixed NumRM integer"),
            )
        });
        let string_length = usize::from(u16::from_le_bytes([operand[64], operand[65]]));
        if string_length > 31 {
            return Err(DocError::Corrupted(
                "NumRM format string exceeds its 31-code-unit field".to_string(),
            ));
        }
        let units = (0..string_length)
            .map(|index| {
                let offset = 66 + index * 2;
                u16::from_le_bytes([operand[offset], operand[offset + 1]])
            })
            .collect::<Vec<_>>();
        let format_string = String::from_utf16(&units).map_err(|_| {
            DocError::Corrupted("NumRM format string is invalid UTF-16".to_string())
        })?;
        if placeholder_positions
            .iter()
            .any(|position| usize::from(*position) > string_length)
        {
            return Err(DocError::Corrupted(
                "NumRM placeholder position exceeds its format string".to_string(),
            ));
        }
        Ok(NumberingRevisionProperties {
            was_numbered,
            author_index,
            timestamp,
            placeholder_positions,
            number_formats,
            numbers,
            format_string,
        })
    }

    /// Handle tab stops (sprmPChgTabsPapx).
    ///
    /// Tab stops are stored as:
    /// - 1 byte: number of tabs to delete (delSize)
    /// - delSize * 2 bytes: positions to delete
    /// - 1 byte: number of tabs to add (addSize)
    /// - addSize * 2 bytes: positions to add
    /// - addSize bytes: tab descriptors (jc + tlc)
    fn handle_tabs(pap: &mut ParagraphProperties, sprm: &Sprm) {
        let bytes = sprm.operand_bytes();
        if bytes.is_empty() {
            return;
        }

        let mut offset = 0;

        // Read delete count
        let del_size = bytes[offset] as usize;
        offset += 1;

        // Create a map of existing tabs
        let mut tab_map: std::collections::HashMap<i32, TabStop> =
            pap.tab_stops.iter().map(|t| (t.position, *t)).collect();

        // Delete tabs
        for _ in 0..del_size {
            if offset + 1 < bytes.len() {
                if let Ok(pos) = read_i16_le(bytes, offset) {
                    tab_map.remove(&(pos as i32));
                }
                offset += 2;
            }
        }

        // Read add count
        if offset >= bytes.len() {
            return;
        }
        let add_size = bytes[offset] as usize;
        offset += 1;

        // Read new tab positions
        let positions_start = offset;
        offset += add_size * 2;

        // Read tab descriptors and add tabs
        for i in 0..add_size {
            if positions_start + i * 2 + 1 < bytes.len()
                && offset < bytes.len()
                && let Ok(pos) = read_i16_le(bytes, positions_start + i * 2)
            {
                let tbd = bytes[offset];
                let jc = tbd & 0x07;
                let tlc = (tbd >> 3) & 0x07;

                let alignment = match jc {
                    0 => TabAlignment::Left,
                    1 => TabAlignment::Center,
                    2 => TabAlignment::Right,
                    3 => TabAlignment::Decimal,
                    4 => TabAlignment::Bar,
                    _ => TabAlignment::Left,
                };

                let leader = match tlc {
                    0 => TabLeader::None,
                    1 => TabLeader::Dots,
                    2 => TabLeader::Hyphens,
                    3 => TabLeader::Underline,
                    4 => TabLeader::Heavy,
                    5 => TabLeader::MiddleDot,
                    _ => TabLeader::None,
                };

                tab_map.insert(
                    pos as i32,
                    TabStop {
                        position: pos as i32,
                        alignment,
                        leader,
                    },
                );

                offset += 1;
            }
        }

        // Convert map back to sorted vector
        let mut tabs: Vec<TabStop> = tab_map.into_values().collect();
        tabs.sort_by_key(|t| t.position);
        pap.tab_stops = tabs;
    }

    /// Parse a Word 97 `Brc80` paragraph border.
    fn parse_border80(sprm: &Sprm) -> Result<Option<Border>> {
        let data = sprm.operand_bytes();
        if data.len() != 4 {
            return Err(DocError::Corrupted(
                "DOC paragraph Brc80 must contain exactly 4 bytes".to_string(),
            ));
        }
        let Some(style) = Self::parse_border_style(data[1], false)? else {
            return Ok(None);
        };
        let color = match data[2] {
            0 => None,
            index @ 1..=16 => Some(Self::get_ico_color(index)),
            invalid => {
                return Err(DocError::Corrupted(format!(
                    "DOC paragraph Brc80 has invalid color index {invalid}"
                )));
            },
        };
        Ok(Some(Border {
            style,
            width: data[0],
            color,
            spacing: data[3] & 0x1F,
            shadow: data[3] & 0x20 != 0,
            frame: data[3] & 0x40 != 0,
        }))
    }

    /// Parse a current 8-byte `Brc` wrapped by a `BrcOperand`.
    fn parse_current_border(sprm: &Sprm) -> Result<Option<Border>> {
        let data = sprm.operand_bytes();
        if data.len() != 8 {
            return Err(DocError::Corrupted(
                "DOC paragraph BrcOperand must contain exactly 8 bytes".to_string(),
            ));
        }
        let Some(style) = Self::parse_border_style(data[5], true)? else {
            return Ok(None);
        };
        let color = match data[3] {
            0 => Some((data[0], data[1], data[2])),
            0xFF => None,
            invalid => {
                return Err(DocError::Corrupted(format!(
                    "DOC paragraph Brc has invalid automatic-color flag {invalid:#04x}"
                )));
            },
        };
        Ok(Some(Border {
            style,
            width: data[4],
            color,
            spacing: data[6] & 0x1F,
            shadow: data[6] & 0x20 != 0,
            frame: data[6] & 0x40 != 0,
        }))
    }

    fn parse_border_style(code: u8, current: bool) -> Result<Option<BorderStyle>> {
        Ok(Some(match code {
            0 => return Ok(None),
            1 => BorderStyle::Single,
            3 => BorderStyle::Double,
            5 => BorderStyle::Thick,
            6 => BorderStyle::Dotted,
            7 => BorderStyle::Dashed,
            8 => BorderStyle::DotDash,
            9 => BorderStyle::DotDotDash,
            10 => BorderStyle::Triple,
            11 => BorderStyle::ThinThickSmallGap,
            12 => BorderStyle::ThickThinSmallGap,
            13 => BorderStyle::ThinThickThinSmallGap,
            14 => BorderStyle::ThinThickMediumGap,
            15 => BorderStyle::ThickThinMediumGap,
            16 => BorderStyle::ThinThickThinMediumGap,
            17 => BorderStyle::ThinThickLargeGap,
            18 => BorderStyle::ThickThinLargeGap,
            19 => BorderStyle::ThinThickThinLargeGap,
            20 => BorderStyle::Wave,
            21 => BorderStyle::DoubleWave,
            22 => BorderStyle::DashSmallGap,
            23 => BorderStyle::DashDotStroked,
            24 => BorderStyle::ThreeDEmboss,
            25 => BorderStyle::ThreeDEngrave,
            26 if current => BorderStyle::Outset,
            27 if current => BorderStyle::Inset,
            invalid => {
                return Err(DocError::Corrupted(format!(
                    "DOC paragraph border has invalid type {invalid:#04x}"
                )));
            },
        }))
    }

    /// Parse shading from Shd80 (2 bytes).
    fn parse_shd80(shd: u16) -> Result<Option<Shading>> {
        if shd == u16::MAX {
            return Ok(None);
        }
        let ico_fore = (shd & 0x1F) as u8;
        let ico_back = ((shd >> 5) & 0x1F) as u8;
        let ipat = ((shd >> 10) & 0x3F) as u8;
        let pattern = ShadingPattern::from_u8(ipat).ok_or_else(|| {
            DocError::Corrupted(format!("sprmPShd80 has invalid pattern {ipat:#04x}"))
        })?;
        if pattern == ShadingPattern::Auto {
            return Ok(None);
        }
        let palette_color = |index| match index {
            0 => Ok(None),
            value @ 1..=16 => Ok(Some(Self::get_ico_color(value))),
            invalid => Err(DocError::Corrupted(format!(
                "sprmPShd80 has invalid color index {invalid}"
            ))),
        };
        Ok(Some(Shading {
            foreground_color: palette_color(ico_fore)?,
            background_color: palette_color(ico_back)?,
            pattern,
        }))
    }

    /// Parse shading from ShadingDescriptor (10 bytes).
    fn parse_shading_descriptor(sprm: &Sprm) -> Result<Option<Shading>> {
        let data = sprm.operand_bytes();
        if data.len() != 10 {
            return Err(DocError::Corrupted(
                "sprmPShd SHDOperand must contain exactly 10 bytes".to_string(),
            ));
        }
        let pattern_code = read_u16_le(data, 8).map_err(|error| {
            DocError::Corrupted(format!("sprmPShd has invalid pattern: {error}"))
        })?;
        let pattern = u8::try_from(pattern_code)
            .ok()
            .and_then(ShadingPattern::from_u8)
            .ok_or_else(|| {
                DocError::Corrupted(format!("sprmPShd has invalid pattern {pattern_code:#06x}"))
            })?;
        if pattern == ShadingPattern::Auto {
            return Ok(None);
        }
        let colorref = |bytes: &[u8]| match bytes[3] {
            0 => Ok(Some((bytes[0], bytes[1], bytes[2]))),
            0xFF => Ok(None),
            invalid => Err(DocError::Corrupted(format!(
                "sprmPShd has invalid automatic-color flag {invalid:#04x}"
            ))),
        };
        Ok(Some(Shading {
            foreground_color: colorref(&data[..4])?,
            background_color: colorref(&data[4..8])?,
            pattern,
        }))
    }

    /// Get color from ico index.
    fn get_ico_color(ico: u8) -> (u8, u8, u8) {
        match ico {
            0 => (0, 0, 0),        // Auto/Black
            1 => (0, 0, 0),        // Black
            2 => (0, 0, 255),      // Blue
            3 => (0, 255, 255),    // Cyan
            4 => (0, 255, 0),      // Green
            5 => (255, 0, 255),    // Magenta
            6 => (255, 0, 0),      // Red
            7 => (255, 255, 0),    // Yellow
            8 => (255, 255, 255),  // White
            9 => (0, 0, 128),      // Dark Blue
            10 => (0, 128, 128),   // Dark Cyan
            11 => (0, 128, 0),     // Dark Green
            12 => (128, 0, 128),   // Dark Magenta
            13 => (128, 0, 0),     // Dark Red
            14 => (128, 128, 0),   // Dark Yellow
            15 => (128, 128, 128), // Dark Gray
            16 => (192, 192, 192), // Light Gray
            _ => (0, 0, 0),
        }
    }

    /// Check if any formatting is applied.
    pub fn has_formatting(&self) -> bool {
        self.justification != Justification::Left
            || self.indent_left.is_some()
            || self.indent_right.is_some()
            || self.indent_first_line.is_some()
            || self.indent_left_chars.is_some()
            || self.indent_right_chars.is_some()
            || self.indent_first_line_chars.is_some()
            || self.space_before.is_some()
            || self.space_after.is_some()
            || self.space_before_lines.is_some()
            || self.space_after_lines.is_some()
            || self.space_before_auto
            || self.space_after_auto
            || self.line_spacing.is_some()
            || self.keep_on_page
            || self.keep_with_next
            || self.page_break_before
            || self.widow_control
            || self.use_page_setup_settings.is_some()
            || self.adjust_right_indent.is_some()
            || self.no_allow_overlap
            || self.contextual_spacing
            || self.mirror_indents
            || self.text_box_tight_wrap.is_some()
            || self.borders != Borders::default()
            || self.shading.is_some()
            || !self.tab_stops.is_empty()
    }

    /// Get indent in inches.
    pub fn get_indent_left_inches(&self) -> f32 {
        self.indent_left.map(|v| v as f32 / 1440.0).unwrap_or(0.0)
    }

    /// Get right indent in inches.
    pub fn get_indent_right_inches(&self) -> f32 {
        self.indent_right.map(|v| v as f32 / 1440.0).unwrap_or(0.0)
    }

    /// Get first line indent in inches.
    pub fn get_indent_first_line_inches(&self) -> f32 {
        self.indent_first_line
            .map(|v| v as f32 / 1440.0)
            .unwrap_or(0.0)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_default_pap() {
        let pap = ParagraphProperties::new();
        assert_eq!(pap.justification, Justification::Left);
        assert!(!pap.keep_on_page);
        assert!(!pap.has_formatting());
    }

    #[test]
    fn test_justification() {
        let left = Justification::Left;
        let center = Justification::Center;
        assert_ne!(left, center);
        assert_eq!(left, Justification::Left);
    }

    #[test]
    fn test_line_spacing_type() {
        let single = LineSpacingType::Single;
        let double = LineSpacingType::Double;
        assert_ne!(single, double);
    }

    #[test]
    fn test_indent_conversion() {
        let mut pap = ParagraphProperties::new();
        pap.indent_left = Some(1440); // 1 inch in twips
        assert_eq!(pap.get_indent_left_inches(), 1.0);
    }

    #[test]
    fn parses_all_paragraph_formatting_revision_sprms_strictly() {
        let timestamp =
            30u32 | (14u32 << 6) | (15u32 << 11) | (7u32 << 16) | (126u32 << 20) | (3u32 << 29);
        for opcode in [
            SPRM_P_PROP_RMARK,
            SPRM_P_PROP_RMARK90,
            SPRM_P_PROP_RMARK_CURRENT,
        ] {
            let mut grpprl = opcode.to_le_bytes().to_vec();
            grpprl.push(7);
            grpprl.push(1);
            grpprl.extend_from_slice(&2i16.to_le_bytes());
            grpprl.extend_from_slice(&timestamp.to_le_bytes());
            let properties = ParagraphProperties::from_sprm(&grpprl).unwrap();
            assert_eq!(properties.has_formatting_revision, Some(true));
            assert_eq!(properties.formatting_revision_author_index, Some(2));
            assert_eq!(properties.formatting_revision_timestamp, Some(timestamp));
        }

        for operand in [
            vec![2, 0, 0, 0, 0, 0, 0],
            vec![1, 0xFF, 0xFF, 0, 0, 0, 0],
            vec![1, 0, 0, 0, 0, 0],
            vec![1, 0, 0, 0x3F, 0, 0, 0],
        ] {
            let mut grpprl = SPRM_P_PROP_RMARK_CURRENT.to_le_bytes().to_vec();
            grpprl.push(operand.len() as u8);
            grpprl.extend_from_slice(&operand);
            assert!(ParagraphProperties::from_sprm(&grpprl).is_err());
        }
    }

    #[test]
    fn parses_table_row_revision_state_strictly() {
        let timestamp =
            30u32 | (14u32 << 6) | (15u32 << 11) | (7u32 << 16) | (126u32 << 20) | (3u32 << 29);
        let mut grpprl = 0xD667u16.to_le_bytes().to_vec();
        grpprl.push(7);
        grpprl.push(1);
        grpprl.extend_from_slice(&2i16.to_le_bytes());
        grpprl.extend_from_slice(&timestamp.to_le_bytes());
        grpprl.extend_from_slice(&0x3668u16.to_le_bytes());
        grpprl.push(1);
        let properties = ParagraphProperties::from_sprm(&grpprl).unwrap();
        assert_eq!(properties.has_table_formatting_revision, Some(true));
        assert_eq!(properties.table_formatting_revision_author_index, Some(2));
        assert_eq!(
            properties.table_formatting_revision_timestamp,
            Some(timestamp)
        );
        assert!(properties.table_properties_preserved_for_revision);

        for operand in [
            vec![2, 0, 0, 0, 0, 0, 0],
            vec![1, 0xFF, 0xFF, 0, 0, 0, 0],
            vec![1, 0, 0, 0, 0, 0],
            vec![1, 0, 0, 0x3F, 0, 0, 0],
        ] {
            let mut invalid = 0xD667u16.to_le_bytes().to_vec();
            invalid.push(operand.len() as u8);
            invalid.extend_from_slice(&operand);
            assert!(ParagraphProperties::from_sprm(&invalid).is_err());
        }

        let invalid_wall = [0x68, 0x36, 2];
        assert!(ParagraphProperties::from_sprm(&invalid_wall).is_err());
    }

    #[test]
    fn parses_numbering_revision_state_strictly() {
        let timestamp =
            30u32 | (14u32 << 6) | (15u32 << 11) | (7u32 << 16) | (126u32 << 20) | (3u32 << 29);
        let mut numrm = [0u8; 128];
        numrm[0] = 1;
        numrm[2..4].copy_from_slice(&1i16.to_le_bytes());
        numrm[4..8].copy_from_slice(&timestamp.to_le_bytes());
        numrm[8] = 1;
        numrm[17] = 0;
        numrm[28..32].copy_from_slice(&12u32.to_le_bytes());
        numrm[64..66].copy_from_slice(&2u16.to_le_bytes());
        numrm[66..68].copy_from_slice(&('%' as u16).to_le_bytes());
        numrm[68..70].copy_from_slice(&('.' as u16).to_le_bytes());

        let mut grpprl = SPRM_P_F_NUM_RM_INS.to_le_bytes().to_vec();
        grpprl.push(1);
        grpprl.extend_from_slice(&SPRM_P_NUM_RM.to_le_bytes());
        grpprl.push(128);
        grpprl.extend_from_slice(&numrm);
        let properties = ParagraphProperties::from_sprm(&grpprl).unwrap();
        assert_eq!(properties.numbering_revision_list_applied, Some(true));
        let revision = properties.numbering_revision.unwrap();
        assert!(revision.was_numbered);
        assert_eq!(revision.author_index, 1);
        assert_eq!(revision.timestamp, timestamp);
        assert_eq!(revision.placeholder_positions[0], 1);
        assert_eq!(revision.numbers[0], 12);
        assert_eq!(revision.format_string, "%.");

        let mut invalid_bool = SPRM_P_F_NUM_RM_INS.to_le_bytes().to_vec();
        invalid_bool.push(2);
        assert!(ParagraphProperties::from_sprm(&invalid_bool).is_err());

        for mutate in [0usize, 2, 8, 64] {
            let mut invalid = numrm;
            match mutate {
                0 => invalid[0] = 2,
                2 => invalid[2..4].copy_from_slice(&(-1i16).to_le_bytes()),
                8 => invalid[8] = 3,
                64 => invalid[64..66].copy_from_slice(&32u16.to_le_bytes()),
                _ => unreachable!(),
            }
            let mut grpprl = SPRM_P_NUM_RM.to_le_bytes().to_vec();
            grpprl.push(128);
            grpprl.extend_from_slice(&invalid);
            assert!(ParagraphProperties::from_sprm(&grpprl).is_err());
        }
    }

    #[test]
    fn parses_current_paragraph_identity_and_revision_state_strictly() {
        let mut grpprl = Vec::new();
        grpprl.extend_from_slice(&SPRM_P_WALL.to_le_bytes());
        grpprl.push(1);
        grpprl.extend_from_slice(&SPRM_P_IPGP.to_le_bytes());
        grpprl.extend_from_slice(&9u32.to_le_bytes());
        grpprl.extend_from_slice(&SPRM_P_RSID.to_le_bytes());
        grpprl.extend_from_slice(&0x1122_3344u32.to_le_bytes());
        grpprl.extend_from_slice(&SPRM_P_F_NO_ALLOW_OVERLAP.to_le_bytes());
        grpprl.push(1);
        grpprl.extend_from_slice(&SPRM_P_F_CONTEXTUAL_SPACING.to_le_bytes());
        grpprl.push(1);
        grpprl.extend_from_slice(&SPRM_P_F_MIRROR_INDENTS.to_le_bytes());
        grpprl.push(1);
        grpprl.extend_from_slice(&SPRM_P_TTWO.to_le_bytes());
        grpprl.push(4);

        let properties = ParagraphProperties::from_sprm(&grpprl).unwrap();
        assert!(properties.properties_preserved_for_revision);
        assert_eq!(properties.paragraph_group_id, Some(9));
        assert_eq!(properties.revision_save_id, Some(0x1122_3344));
        assert!(properties.no_allow_overlap);
        assert!(properties.contextual_spacing);
        assert!(properties.mirror_indents);
        assert_eq!(
            properties.text_box_tight_wrap,
            Some(TextBoxTightWrap::LastLineOnly)
        );

        let invalid_bool = [SPRM_P_WALL.to_le_bytes().as_slice(), &[2]].concat();
        assert!(ParagraphProperties::from_sprm(&invalid_bool).is_err());

        let invalid_group = [
            SPRM_P_IPGP.to_le_bytes().as_slice(),
            0u32.to_le_bytes().as_slice(),
        ]
        .concat();
        assert!(ParagraphProperties::from_sprm(&invalid_group).is_err());

        let truncated_rsid = [SPRM_P_RSID.to_le_bytes().as_slice(), &[1, 2]].concat();
        assert!(ParagraphProperties::from_sprm(&truncated_rsid).is_err());

        let invalid_tight_wrap = [SPRM_P_TTWO.to_le_bytes().as_slice(), &[5]].concat();
        assert!(ParagraphProperties::from_sprm(&invalid_tight_wrap).is_err());
    }

    #[test]
    fn parses_current_character_relative_paragraph_layout_strictly() {
        let mut grpprl = Vec::new();
        for (opcode, value) in [
            (SPRM_P_DXC_RIGHT, -125i16),
            (SPRM_P_DXC_LEFT, 250),
            (SPRM_P_DXC_LEFT1, -50),
            (SPRM_P_DYL_BEFORE, -20),
            (SPRM_P_DYL_AFTER, 31_680),
            (SPRM_P_DXA_LEFT_2000, 100),
            (SPRM_P_NEST_2000, -20),
        ] {
            grpprl.extend_from_slice(&opcode.to_le_bytes());
            grpprl.extend_from_slice(&value.to_le_bytes());
        }
        for opcode in [
            SPRM_P_F_OPEN_TCH,
            SPRM_P_F_DYA_BEFORE_AUTO,
            SPRM_P_F_DYA_AFTER_AUTO,
        ] {
            grpprl.extend_from_slice(&opcode.to_le_bytes());
            grpprl.push(1);
        }
        grpprl.extend_from_slice(&SPRM_P_JC_LOGICAL.to_le_bytes());
        grpprl.push(9);

        let properties = ParagraphProperties::from_sprm(&grpprl).unwrap();
        assert_eq!(properties.indent_right_chars, Some(-125));
        assert_eq!(properties.indent_left_chars, Some(250));
        assert_eq!(properties.indent_first_line_chars, Some(-50));
        assert_eq!(properties.space_before_lines, Some(-20));
        assert_eq!(properties.space_after_lines, Some(31_680));
        assert_eq!(properties.indent_left, Some(80));
        assert!(properties.open_table_cell_mark);
        assert!(properties.space_before_auto);
        assert!(properties.space_after_auto);
        assert_eq!(properties.justification, Justification::ThaiDistributed);

        for (opcode, value) in [(SPRM_P_DYL_BEFORE, -21i16), (SPRM_P_DYL_AFTER, 31_681)] {
            let invalid = [opcode.to_le_bytes(), value.to_le_bytes()].concat();
            assert!(ParagraphProperties::from_sprm(&invalid).is_err());
        }

        let invalid_bool = [SPRM_P_F_OPEN_TCH.to_le_bytes().as_slice(), &[2]].concat();
        assert!(ParagraphProperties::from_sprm(&invalid_bool).is_err());

        let invalid_logical_jc = [SPRM_P_JC_LOGICAL.to_le_bytes().as_slice(), &[10]].concat();
        assert!(ParagraphProperties::from_sprm(&invalid_logical_jc).is_err());

        let invalid_legacy_jc = [SPRM_P_JC.to_le_bytes().as_slice(), &[5]].concat();
        assert!(ParagraphProperties::from_sprm(&invalid_legacy_jc).is_err());
    }

    #[test]
    fn parses_current_and_word97_paragraph_borders_strictly() {
        let mut grpprl = SPRM_P_BRC_TOP.to_le_bytes().to_vec();
        grpprl.push(8);
        grpprl.extend_from_slice(&[0x11, 0x22, 0x33, 0, 12, 27, 0x67, 0]);
        grpprl.extend_from_slice(&SPRM_P_BRC_LEFT80.to_le_bytes());
        grpprl.extend_from_slice(&[8, 3, 2, 0x24]);

        let properties = ParagraphProperties::from_sprm(&grpprl).unwrap();
        assert_eq!(
            properties.borders.top,
            Some(Border {
                style: BorderStyle::Inset,
                width: 12,
                color: Some((0x11, 0x22, 0x33)),
                spacing: 7,
                shadow: true,
                frame: true,
            })
        );
        assert_eq!(
            properties.borders.left,
            Some(Border {
                style: BorderStyle::Double,
                width: 8,
                color: Some((0, 0, 255)),
                spacing: 4,
                shadow: true,
                frame: false,
            })
        );

        for operand in [
            vec![0, 0, 0, 1, 8, 1, 0, 0],
            vec![0, 0, 0, 0, 8, 2, 0, 0],
            vec![0; 7],
        ] {
            let mut invalid = SPRM_P_BRC_TOP.to_le_bytes().to_vec();
            invalid.push(operand.len() as u8);
            invalid.extend_from_slice(&operand);
            assert!(ParagraphProperties::from_sprm(&invalid).is_err());
        }

        for operand in [[8, 1, 17, 0], [8, 26, 1, 0]] {
            let invalid = [SPRM_P_BRC_TOP80.to_le_bytes().as_slice(), &operand].concat();
            assert!(ParagraphProperties::from_sprm(&invalid).is_err());
        }
    }

    #[test]
    fn parses_current_grid_table_depth_and_shading_strictly() {
        let mut grpprl = Vec::new();
        for opcode in [SPRM_P_F_USE_PGSU_SETTINGS, SPRM_P_F_ADJUST_RIGHT] {
            grpprl.extend_from_slice(&opcode.to_le_bytes());
            grpprl.push(1);
        }
        grpprl.extend_from_slice(&SPRM_P_ITAP.to_le_bytes());
        grpprl.extend_from_slice(&3i32.to_le_bytes());
        grpprl.extend_from_slice(&SPRM_P_DTAP.to_le_bytes());
        grpprl.extend_from_slice(&(-1i32).to_le_bytes());
        for opcode in [SPRM_P_F_INNER_TABLE_CELL, SPRM_P_F_INNER_TTP] {
            grpprl.extend_from_slice(&opcode.to_le_bytes());
            grpprl.push(1);
        }
        grpprl.extend_from_slice(&SPRM_P_SHD.to_le_bytes());
        grpprl.push(10);
        grpprl.extend_from_slice(&[1, 2, 3, 0, 4, 5, 6, 0, 0x19, 0]);

        let properties = ParagraphProperties::from_sprm(&grpprl).unwrap();
        assert_eq!(properties.use_page_setup_settings, Some(true));
        assert_eq!(properties.adjust_right_indent, Some(true));
        assert_eq!(properties.table_nesting_level, 2);
        assert!(properties.inner_table_cell);
        assert!(properties.inner_table_row_end);
        assert_eq!(
            properties.shading,
            Some(Shading {
                foreground_color: Some((1, 2, 3)),
                background_color: Some((4, 5, 6)),
                pattern: ShadingPattern::DiagonalCross,
            })
        );

        let invalid_i32 = |opcode: u16, value: i32| {
            [
                opcode.to_le_bytes().as_slice(),
                value.to_le_bytes().as_slice(),
            ]
            .concat()
        };
        assert!(ParagraphProperties::from_sprm(&invalid_i32(SPRM_P_ITAP, -1)).is_err());
        assert!(ParagraphProperties::from_sprm(&invalid_i32(SPRM_P_DTAP, -1)).is_err());
        assert!(
            ParagraphProperties::from_sprm(
                &[SPRM_P_F_USE_PGSU_SETTINGS.to_le_bytes().as_slice(), &[2]].concat()
            )
            .is_err()
        );

        let inner_at_depth_one = [
            SPRM_P_ITAP.to_le_bytes().as_slice(),
            1i32.to_le_bytes().as_slice(),
            SPRM_P_F_INNER_TABLE_CELL.to_le_bytes().as_slice(),
            &[1],
        ]
        .concat();
        assert!(ParagraphProperties::from_sprm(&inner_at_depth_one).is_err());

        for operand in [
            vec![1, 2, 3, 2, 4, 5, 6, 0, 1, 0],
            vec![1, 2, 3, 0, 4, 5, 6, 0, 0x1A, 0],
            vec![0; 9],
        ] {
            let mut invalid = SPRM_P_SHD.to_le_bytes().to_vec();
            invalid.push(operand.len() as u8);
            invalid.extend_from_slice(&operand);
            assert!(ParagraphProperties::from_sprm(&invalid).is_err());
        }
    }
}
