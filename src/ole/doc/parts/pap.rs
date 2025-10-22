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
use super::super::package::Result;
use crate::common::binary::{read_i16_le, read_u16_le, read_u32_le};
use crate::ole::sprm::{Sprm, parse_sprms};
use crate::ole::sprm_operations::*;

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
    /// Space before paragraph in twips
    pub space_before: Option<u16>,
    /// Space after paragraph in twips
    pub space_after: Option<u16>,
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
#[derive(Debug, Clone, Default)]
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
#[derive(Debug, Clone, Copy)]
pub struct Border {
    /// Border style
    pub style: BorderStyle,
    /// Border width in eighths of a point
    pub width: u8,
    /// Border color (RGB)
    pub color: (u8, u8, u8),
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
}

/// Paragraph shading.
#[derive(Debug, Clone, Copy)]
pub struct Shading {
    /// Background color (RGB)
    pub background_color: (u8, u8, u8),
    /// Foreground color (RGB) for patterns
    pub foreground_color: (u8, u8, u8),
    /// Shading pattern
    pub pattern: ShadingPattern,
}

/// Shading patterns.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ShadingPattern {
    Clear,
    Solid,
    Percent5,
    Percent10,
    Percent20,
    Percent25,
    Percent30,
    Percent40,
    Percent50,
    Percent60,
    Percent70,
    Percent75,
    Percent80,
    Percent90,
    DarkHorizontal,
    DarkVertical,
    DarkForwardDiagonal,
    DarkBackwardDiagonal,
    DarkCross,
    DarkDiagonalCross,
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
        let mut pap = Self::default();
        let sprms = parse_sprms(grpprl);

        for sprm in &sprms {
            // Only process PAP SPRMs (type = 1)
            if get_sprm_type(sprm.opcode) == 1 {
                Self::apply_sprm(&mut pap, sprm);
            }
        }

        Ok(pap)
    }

    /// Apply a single SPRM operation to paragraph properties.
    ///
    /// Based on Apache POI's ParagraphSprmUncompressor.unCompressPAPOperation().
    ///
    /// # Arguments
    ///
    /// * `pap` - The paragraph properties to modify
    /// * `sprm` - The SPRM operation to apply
    fn apply_sprm(pap: &mut ParagraphProperties, sprm: &Sprm) {
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
                if let Some(jc) = sprm.operand_byte() {
                    pap.justification = match jc {
                        0 => Justification::Left,
                        1 => Justification::Center,
                        2 => Justification::Right,
                        3 => Justification::Justified,
                        4 => Justification::Distributed,
                        _ => Justification::Left,
                    };
                }
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
            // Operation 0x24: sprmPBrcTop - Top border
            0x24 => {
                // Parse BorderCode structure (4 bytes)
                if sprm.operand.len() >= 4 {
                    pap.borders.top = Self::parse_border(&sprm.operand);
                }
            },
            // Operation 0x25: sprmPBrcLeft - Left border
            0x25 => {
                if sprm.operand.len() >= 4 {
                    pap.borders.left = Self::parse_border(&sprm.operand);
                }
            },
            // Operation 0x26: sprmPBrcBottom - Bottom border
            0x26 => {
                if sprm.operand.len() >= 4 {
                    pap.borders.bottom = Self::parse_border(&sprm.operand);
                }
            },
            // Operation 0x27: sprmPBrcRight - Right border
            0x27 => {
                if sprm.operand.len() >= 4 {
                    pap.borders.right = Self::parse_border(&sprm.operand);
                }
            },
            // Operation 0x28: sprmPBrcBetween - Between border
            0x28 => {
                if sprm.operand.len() >= 4 {
                    pap.borders.between = Self::parse_border(&sprm.operand);
                }
            },
            // Operation 0x29: sprmPBrcBar - Bar border
            0x29 => {
                if sprm.operand.len() >= 4 {
                    pap.borders.bar = Self::parse_border(&sprm.operand);
                }
            },
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
                if let Some(shd) = sprm.operand_word() {
                    pap.shading = Self::parse_shd80(shd);
                }
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
            // Operation 0x3F: sprmPPropRMark - Property revision mark
            0x3F => {
                // Revision mark properties - not commonly used
            },
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
                // Numbering revision - not commonly used
            },
            // Operation 0x44: sprmPCrLf - CR/LF
            0x44 => {
                // Not commonly used
            },
            // Operation 0x45: sprmPNumRM - Numbering revision mark
            0x45 => {
                // Numbering revision mark - complex structure
            },
            // Operation 0x47: sprmPFUsePgsuSettings - Use page setup settings
            0x47 => {
                // Use page setup settings - not commonly used
            },
            // Operation 0x48: sprmPFAdjustRight - Adjust right
            0x48 => {
                // Adjust right - not commonly used
            },
            // Operation 0x49: sprmPItap - Table nesting level
            0x49 => {
                if let Some(itap) = sprm.operand_dword() {
                    pap.table_nesting_level = itap as i32;
                }
            },
            // Operation 0x4A: sprmPDtap - Table nesting delta
            0x4A => {
                if let Some(dtap) = sprm.operand_dword() {
                    pap.table_nesting_level += dtap as i32;
                }
            },
            // Operation 0x4B: sprmPFInnerTableCell - Inner table cell
            0x4B => {
                if let Some(val) = sprm.operand_byte() {
                    pap.inner_table_cell = val != 0;
                }
            },
            // Operation 0x4C: sprmPFInnerTtp - Inner table row end
            0x4C => {
                if let Some(val) = sprm.operand_byte() {
                    pap.inner_table_row_end = val != 0;
                }
            },
            // Operation 0x4D: sprmPShd - Shading (Word 2002+)
            0x4D => {
                // Parse ShadingDescriptor structure
                if sprm.operand.len() >= 10 {
                    pap.shading = Self::parse_shading_descriptor(&sprm.operand);
                }
            },
            // Operations 0x4E-0x53: Borders v80
            0x4E..=0x53 => {
                // BrcXXX80 - Word 97-2000 borders
            },
            // Operation 0x5D: sprmPDxaRight (alternative)
            0x5D => {
                if let Some(val) = sprm.operand_i16() {
                    pap.indent_right = Some(val as i32);
                }
            },
            // Operation 0x5E: sprmPDxaLeft (alternative)
            0x5E => {
                if let Some(val) = sprm.operand_i16() {
                    pap.indent_left = Some(val as i32);
                }
            },
            // Operation 0x60: sprmPDxaLeft1 (alternative)
            0x60 => {
                if let Some(val) = sprm.operand_i16() {
                    pap.indent_first_line = Some(val as i32);
                }
            },
            // Operation 0x61: sprmPJc (logical justification for bi-di)
            0x61 => {
                if let Some(jc) = sprm.operand_byte() {
                    pap.justification = match jc {
                        0 => Justification::Left,
                        1 => Justification::Center,
                        2 => Justification::Right,
                        3 => Justification::Justified,
                        4 => Justification::Distributed,
                        _ => Justification::Left,
                    };
                }
            },
            // Operation 0x67: sprmPRsid - Revision save ID
            0x67 => {
                // Revision save ID - not commonly used
            },
            // Default: Unknown or unsupported SPRM
            _ => {
                // Silently ignore unknown SPRMs
            },
        }
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

    /// Parse a border from BorderCode structure (4 bytes).
    fn parse_border(data: &[u8]) -> Option<Border> {
        if data.len() < 4 {
            return None;
        }

        // BorderCode structure (simplified)
        let dpt_line_width = data[0];
        let brc_type = data[1];
        let ico = data[2];

        if brc_type == 0 || brc_type == 255 {
            return None; // No border
        }

        let style = match brc_type {
            1 => BorderStyle::Single,
            2 => BorderStyle::Thick,
            3 => BorderStyle::Double,
            5 => BorderStyle::Dotted,
            6 => BorderStyle::Dashed,
            7 => BorderStyle::DotDash,
            8 => BorderStyle::DotDotDash,
            9 => BorderStyle::Triple,
            _ => BorderStyle::Single,
        };

        let color = match ico {
            1 => (0, 0, 0),       // Black
            2 => (0, 0, 255),     // Blue
            3 => (0, 255, 255),   // Cyan
            4 => (0, 255, 0),     // Green
            5 => (255, 0, 255),   // Magenta
            6 => (255, 0, 0),     // Red
            7 => (255, 255, 0),   // Yellow
            8 => (255, 255, 255), // White
            _ => (0, 0, 0),       // Auto/Black
        };

        Some(Border {
            style,
            width: dpt_line_width,
            color,
        })
    }

    /// Parse shading from Shd80 (2 bytes).
    fn parse_shd80(shd: u16) -> Option<Shading> {
        // Simplified Shd80 parsing
        let ico_fore = (shd & 0x1F) as u8;
        let ico_back = ((shd >> 5) & 0x1F) as u8;
        let ipat = ((shd >> 10) & 0x3F) as u8;

        if ipat == 0 {
            return None;
        }

        let fg_color = Self::get_ico_color(ico_fore);
        let bg_color = Self::get_ico_color(ico_back);

        let pattern = match ipat {
            0 => ShadingPattern::Clear,
            1 => ShadingPattern::Solid,
            2 => ShadingPattern::Percent5,
            3 => ShadingPattern::Percent10,
            4 => ShadingPattern::Percent20,
            5 => ShadingPattern::Percent25,
            6 => ShadingPattern::Percent30,
            7 => ShadingPattern::Percent40,
            8 => ShadingPattern::Percent50,
            9 => ShadingPattern::Percent60,
            10 => ShadingPattern::Percent70,
            11 => ShadingPattern::Percent75,
            12 => ShadingPattern::Percent80,
            13 => ShadingPattern::Percent90,
            _ => ShadingPattern::Clear,
        };

        Some(Shading {
            foreground_color: fg_color,
            background_color: bg_color,
            pattern,
        })
    }

    /// Parse shading from ShadingDescriptor (10 bytes).
    fn parse_shading_descriptor(data: &[u8]) -> Option<Shading> {
        if data.len() < 10 {
            return None;
        }

        // ShadingDescriptor structure (simplified)
        let cv_fore = read_u32_le(data, 0).ok()?;
        let cv_back = read_u32_le(data, 4).ok()?;
        let ipat = read_u16_le(data, 8).ok()?;

        let fg_color = (
            (cv_fore & 0xFF) as u8,
            ((cv_fore >> 8) & 0xFF) as u8,
            ((cv_fore >> 16) & 0xFF) as u8,
        );
        let bg_color = (
            (cv_back & 0xFF) as u8,
            ((cv_back >> 8) & 0xFF) as u8,
            ((cv_back >> 16) & 0xFF) as u8,
        );

        let pattern = match ipat {
            0 => ShadingPattern::Clear,
            1 => ShadingPattern::Solid,
            2 => ShadingPattern::Percent5,
            3 => ShadingPattern::Percent10,
            4 => ShadingPattern::Percent20,
            5 => ShadingPattern::Percent25,
            6 => ShadingPattern::Percent30,
            7 => ShadingPattern::Percent40,
            8 => ShadingPattern::Percent50,
            9 => ShadingPattern::Percent60,
            10 => ShadingPattern::Percent70,
            11 => ShadingPattern::Percent75,
            12 => ShadingPattern::Percent80,
            13 => ShadingPattern::Percent90,
            _ => ShadingPattern::Clear,
        };

        Some(Shading {
            foreground_color: fg_color,
            background_color: bg_color,
            pattern,
        })
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
            || self.space_before.is_some()
            || self.space_after.is_some()
            || self.line_spacing.is_some()
            || self.keep_on_page
            || self.keep_with_next
            || self.page_break_before
            || self.widow_control
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
}
