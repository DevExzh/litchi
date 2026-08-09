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
/// Based on Apache POI's `ParagraphSprmUncompressor` and `ParagraphProperties`.
use crate::parts::numbering::NumberFormat;
pub use crate::parts::tap::{CellShading as Shading, ShadingPattern};
use crate::parts::tap::{TableProperties, TableStyleCondition};

/// Paragraph Properties structure.
///
/// Contains formatting information for a paragraph.
/// Based on Apache POI's `ParagraphProperties` implementation.
#[derive(Debug, Clone, Default)]
pub struct ParagraphProperties {
    /// Justification/alignment
    pub justification: Justification,
    /// Legacy physical justification, when set by `sprmPJc80`
    pub physical_justification: Option<PhysicalJustification>,
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
    /// Obsolete line style selected by `sprmPBrcl`
    pub legacy_border_style: Option<LegacyBorderStyle>,
    /// Obsolete border placement selected by `sprmPBrcp`
    pub legacy_border_position: Option<LegacyBorderPosition>,
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
    /// Conditional paragraph formatting definitions carried by a table style
    pub conditional_formats: Vec<ParagraphConditionalFormatting>,
    /// List level (0 through 8, or 12 when list numbering skips this paragraph)
    pub list_level: Option<u8>,
    /// Signed list format override encoding (negative values preserve paragraph indents)
    pub list_format_override: Option<i16>,
    /// Legacy Word autonumber descriptor (`sprmPAnld`)
    pub legacy_autonumbering: Option<LegacyAutoNumbering>,
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
    /// Vertical font alignment
    pub font_align: Option<FontAlignment>,
    /// Frame text flow
    pub frame_text_flow: Option<FrameTextFlow>,
    /// Horizontal position of the paragraph frame
    pub frame_horizontal_position: Option<FrameHorizontalPosition>,
    /// Vertical position of the paragraph frame
    pub frame_vertical_position: Option<FrameVerticalPosition>,
    /// Frame width in twips, where zero means automatic
    pub frame_width: Option<u16>,
    /// Anchors used to interpret the frame positions
    pub frame_anchor: Option<FrameAnchor>,
    /// Height constraint for the paragraph frame
    pub frame_height: Option<FrameHeight>,
    /// Text wrapping
    pub text_wrap: Option<FrameTextWrap>,
    /// Drop-cap formatting
    pub drop_cap: Option<DropCap>,
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
    /// Paragraph state immediately before the active `sprmPWall` boundary.
    pub preserved_properties_for_revision: Option<Box<ParagraphProperties>>,
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

/// Conditional paragraph formatting carried by `sprmPCnf` in a table style.
#[derive(Debug, Clone)]
pub struct ParagraphConditionalFormatting {
    /// Table location or band for which the nested properties apply.
    pub condition: TableStyleCondition,
    /// Typed paragraph properties decoded from the nested grpprl.
    pub properties: Box<ParagraphProperties>,
    /// Exact nested grpprl retained for lossless preservation.
    pub raw_grpprl: Vec<u8>,
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

    fn try_from(value: u8) -> Result<Self, Self::Error> {
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

/// Physical paragraph justification used by Word 97-compatible `sprmPJc80` records.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PhysicalJustification {
    Left,
    Center,
    Right,
    LowCompression,
    MediumCompression,
    HighCompression,
}

impl TryFrom<u8> for PhysicalJustification {
    type Error = u8;

    fn try_from(value: u8) -> Result<Self, Self::Error> {
        match value {
            0 => Ok(Self::Left),
            1 => Ok(Self::Center),
            2 => Ok(Self::Right),
            3 => Ok(Self::LowCompression),
            4 => Ok(Self::MediumCompression),
            5 => Ok(Self::HighCompression),
            invalid => Err(invalid),
        }
    }
}

impl PhysicalJustification {
    pub(super) fn normalized(self) -> Justification {
        match self {
            Self::Left => Justification::Left,
            Self::Center => Justification::Center,
            Self::Right => Justification::Right,
            Self::LowCompression => Justification::Justified,
            Self::MediumCompression => Justification::MediumKashida,
            Self::HighCompression => Justification::HighKashida,
        }
    }
}

/// Alignment of a legacy Word autonumber label.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AutoNumberAlignment {
    Left,
    Center,
    Right,
    Justified,
}

/// Legacy ANLD autonumbering descriptor used by pre-list-table documents.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct LegacyAutoNumbering {
    pub number_format: NumberFormat,
    pub alignment: AutoNumberAlignment,
    pub include_previous_levels: bool,
    pub hanging_indent: bool,
    pub set_bold: bool,
    pub set_italic: bool,
    pub set_small_caps: bool,
    pub set_caps: bool,
    pub set_strike: bool,
    pub set_underline: bool,
    pub prefix_space: bool,
    pub bold: bool,
    pub italic: bool,
    pub small_caps: bool,
    pub caps: bool,
    pub strike: bool,
    pub underline: u8,
    pub color_index: u8,
    pub font_index: u16,
    pub font_size_half_points: u16,
    pub start_at: u16,
    pub indent_twips: i16,
    pub space_twips: u16,
    pub number_once_per_cell: bool,
    pub number_across_cells: bool,
    pub restart_each_section: bool,
    pub prefix: String,
    pub suffix: String,
}

impl Default for LegacyAutoNumbering {
    fn default() -> Self {
        Self {
            number_format: NumberFormat::Arabic,
            alignment: AutoNumberAlignment::Left,
            include_previous_levels: false,
            hanging_indent: false,
            set_bold: false,
            set_italic: false,
            set_small_caps: false,
            set_caps: false,
            set_strike: false,
            set_underline: false,
            prefix_space: false,
            bold: false,
            italic: false,
            small_caps: false,
            caps: false,
            strike: false,
            underline: 0,
            color_index: 0,
            font_index: 0,
            font_size_half_points: 0,
            start_at: 1,
            indent_twips: 0,
            space_twips: 0,
            number_once_per_cell: false,
            number_across_cells: false,
            restart_each_section: false,
            prefix: String::new(),
            suffix: String::new(),
        }
    }
}

/// Vertical alignment of characters within a line.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u16)]
pub enum FontAlignment {
    Top = 0,
    Center = 1,
    Baseline = 2,
    Bottom = 3,
    Auto = 4,
}

impl TryFrom<u16> for FontAlignment {
    type Error = u16;

    fn try_from(value: u16) -> Result<Self, Self::Error> {
        match value {
            0 => Ok(Self::Top),
            1 => Ok(Self::Center),
            2 => Ok(Self::Baseline),
            3 => Ok(Self::Bottom),
            4 => Ok(Self::Auto),
            invalid => Err(invalid),
        }
    }
}

/// Direction and glyph rotation used by text in a paragraph frame.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct FrameTextFlow {
    pub vertical: bool,
    pub backwards: bool,
    pub rotate_font: bool,
}

/// Text wrapping around a paragraph frame.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u8)]
pub enum FrameTextWrap {
    Auto = 0,
    NotBeside = 1,
    Around = 2,
    None = 3,
    Tight = 4,
    Through = 5,
}

impl TryFrom<u8> for FrameTextWrap {
    type Error = u8;

    fn try_from(value: u8) -> Result<Self, Self::Error> {
        match value {
            0 => Ok(Self::Auto),
            1 => Ok(Self::NotBeside),
            2 => Ok(Self::Around),
            3 => Ok(Self::None),
            4 => Ok(Self::Tight),
            5 => Ok(Self::Through),
            invalid => Err(invalid),
        }
    }
}

/// Height constraint for a paragraph frame.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct FrameHeight {
    pub height_twips: u16,
    pub minimum: bool,
}

/// Drop-cap placement.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DropCapType {
    Regular,
    Margin,
}

/// Drop-cap placement and number of occupied lines.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DropCap {
    pub kind: DropCapType,
    pub lines: u8,
}

/// Horizontal position of a paragraph frame.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FrameHorizontalPosition {
    Left,
    Center,
    Right,
    Inside,
    Outside,
    Offset(i16),
}

/// Vertical position of a paragraph frame.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FrameVerticalPosition {
    Inline,
    Top,
    Center,
    Bottom,
    Inside,
    Outside,
    Offset(i16),
}

/// Vertical reference for an absolutely positioned frame.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FrameVerticalAnchor {
    Margin,
    Page,
    Paragraph,
    None,
}

/// Horizontal reference for an absolutely positioned frame.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FrameHorizontalAnchor {
    Column,
    Margin,
    Page,
    None,
}

/// Reference points used by paragraph-frame coordinates.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct FrameAnchor {
    pub vertical: FrameVerticalAnchor,
    pub horizontal: FrameHorizontalAnchor,
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

    fn try_from(value: u8) -> Result<Self, Self::Error> {
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
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
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
    /// List tab
    List,
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
    /// Default leader behavior (equivalent to no leader)
    DefaultLeader,
}

/// Line style used by the obsolete `sprmPBrcl` paragraph property.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u8)]
pub enum LegacyBorderStyle {
    Single = 0,
    Thick = 1,
    Double = 2,
    Shadow = 3,
}

impl TryFrom<u8> for LegacyBorderStyle {
    type Error = u8;

    fn try_from(value: u8) -> Result<Self, Self::Error> {
        match value {
            0 => Ok(Self::Single),
            1 => Ok(Self::Thick),
            2 => Ok(Self::Double),
            3 => Ok(Self::Shadow),
            invalid => Err(invalid),
        }
    }
}

/// Placement used by the obsolete `sprmPBrcp` paragraph property.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u8)]
pub enum LegacyBorderPosition {
    None = 0,
    Above = 1,
    Below = 2,
    Box = 15,
    LeftBar = 16,
}

impl TryFrom<u8> for LegacyBorderPosition {
    type Error = u8;

    fn try_from(value: u8) -> Result<Self, Self::Error> {
        match value {
            0 => Ok(Self::None),
            1 => Ok(Self::Above),
            2 => Ok(Self::Below),
            15 => Ok(Self::Box),
            16 => Ok(Self::LeftBar),
            invalid => Err(invalid),
        }
    }
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
    /// Check if any formatting is applied.
    #[must_use]
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
            || self.frame_height.is_some()
            || self.frame_horizontal_position.is_some()
            || self.frame_vertical_position.is_some()
            || self.frame_width.is_some()
            || self.frame_anchor.is_some()
            || self.text_wrap.is_some()
            || self.drop_cap.is_some()
            || self.dxa_from_text.is_some()
            || self.dya_from_text.is_some()
            || self.no_auto_hyph
            || self.no_allow_overlap
            || self.contextual_spacing
            || self.mirror_indents
            || self.text_box_tight_wrap.is_some()
            || self.borders != Borders::default()
            || self.legacy_border_style.is_some()
            || self.legacy_border_position.is_some()
            || self.shading.is_some()
            || !self.tab_stops.is_empty()
            || self.legacy_autonumbering.is_some()
            || !self.conditional_formats.is_empty()
    }

    /// Get indent in inches.
    #[must_use]
    pub fn get_indent_left_inches(&self) -> f32 {
        self.indent_left.map_or(0.0, |v| v as f32 / 1440.0)
    }

    /// Get right indent in inches.
    #[must_use]
    pub fn get_indent_right_inches(&self) -> f32 {
        self.indent_right.map_or(0.0, |v| v as f32 / 1440.0)
    }

    /// Get first line indent in inches.
    #[must_use]
    pub fn get_indent_first_line_inches(&self) -> f32 {
        self.indent_first_line.map_or(0.0, |v| v as f32 / 1440.0)
    }
}
