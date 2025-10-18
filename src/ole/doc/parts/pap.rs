/// Paragraph Properties (PAP) parser for DOC files.
///
/// PAP structures define paragraph-level formatting such as:
/// - Alignment (left, right, center, justified)
/// - Indentation (left, right, first line)
/// - Spacing (before, after, line spacing)
/// - Borders and shading
/// - Tab stops
use super::super::package::Result;
use crate::ole::binary::{read_i16_le, read_u16_le};

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
    pub space_before: Option<i32>,
    /// Space after paragraph in twips
    pub space_after: Option<i32>,
    /// Line spacing (in twips or as a multiple)
    pub line_spacing: Option<i32>,
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
    /// Tab stops
    pub tab_stops: Vec<TabStop>,
    /// Borders
    pub borders: Borders,
    /// Background shading
    pub shading: Option<Shading>,
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
    /// Single line spacing
    #[default]
    Single,
    /// 1.5 line spacing
    OnePointFive,
    /// Double line spacing
    Double,
    /// At least N twips
    AtLeast,
    /// Exactly N twips
    Exactly,
    /// Multiple (value in 240ths of a line)
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
    /// Format: opcode (1 or 2 bytes) + operand (variable length)
    ///
    /// # Arguments
    ///
    /// * `grpprl` - Group of SPRMs (property modifications)
    ///
    /// Based on Apache POI's PAP parsing logic in ParagraphSprmUncompressor.
    pub fn from_sprm(grpprl: &[u8]) -> Result<Self> {
        let mut pap = Self::default();
        let mut offset = 0;

        while offset < grpprl.len() {
            if offset + 1 > grpprl.len() {
                break;
            }

            // Read SPRM opcode (2 bytes in Word 97+)
            let sprm = read_u16_le(grpprl, offset).unwrap_or(0);
            offset += 2;

            // Parse SPRM based on opcode using match for idiomatic Rust
            match sprm {
                // Justification (sprmPJc)
                0x2403 | 0x0003 => {
                    if offset < grpprl.len() {
                        pap.justification = match grpprl[offset] {
                            0 => Justification::Left,
                            1 => Justification::Center,
                            2 => Justification::Right,
                            3 => Justification::Justified,
                            4 => Justification::Distributed,
                            _ => Justification::Left,
                        };
                        offset += 1;
                    }
                }
                // Left indent (sprmPDxaLeft)
                0x840F | 0x000F => {
                    if offset + 1 < grpprl.len() {
                        pap.indent_left = Some(read_i16_le(grpprl, offset).unwrap_or(0) as i32);
                        offset += 2;
                    }
                }
                // Right indent (sprmPDxaRight)
                0x8411 | 0x0011 => {
                    if offset + 1 < grpprl.len() {
                        pap.indent_right = Some(read_i16_le(grpprl, offset).unwrap_or(0) as i32);
                        offset += 2;
                    }
                }
                // First line indent (sprmPDxaLeft1)
                0x8416 | 0x0016 => {
                    if offset + 1 < grpprl.len() {
                        pap.indent_first_line = Some(read_i16_le(grpprl, offset).unwrap_or(0) as i32);
                        offset += 2;
                    }
                }
                // Space before (sprmPDyaBefore)
                0xA413 | 0x0013 => {
                    if offset + 1 < grpprl.len() {
                        pap.space_before = Some(read_u16_le(grpprl, offset).unwrap_or(0) as i32);
                        offset += 2;
                    }
                }
                // Space after (sprmPDyaAfter)
                0xA414 | 0x0014 => {
                    if offset + 1 < grpprl.len() {
                        pap.space_after = Some(read_u16_le(grpprl, offset).unwrap_or(0) as i32);
                        offset += 2;
                    }
                }
                // Line spacing (sprmPDyaLine)
                0x6412 | 0x0012 => {
                    if offset + 3 < grpprl.len() {
                        pap.line_spacing = Some(read_i16_le(grpprl, offset).unwrap_or(0) as i32);
                        // Line spacing type is in the next 2 bytes
                        let spacing_type = read_u16_le(grpprl, offset + 2).unwrap_or(0);
                        pap.line_spacing_type = match spacing_type {
                            0 => LineSpacingType::Single,
                            1 => LineSpacingType::OnePointFive,
                            2 => LineSpacingType::Double,
                            3 => LineSpacingType::AtLeast,
                            4 => LineSpacingType::Exactly,
                            5 => LineSpacingType::Multiple,
                            _ => LineSpacingType::Single,
                        };
                        offset += 4;
                    }
                }
                // Keep on page (sprmPFKeep)
                0x2405 | 0x0005 => {
                    if offset < grpprl.len() {
                        pap.keep_on_page = grpprl[offset] != 0;
                        offset += 1;
                    }
                }
                // Keep with next (sprmPFKeepFollow)
                0x2406 | 0x0006 => {
                    if offset < grpprl.len() {
                        pap.keep_with_next = grpprl[offset] != 0;
                        offset += 1;
                    }
                }
                // Page break before (sprmPFPageBreakBefore)
                0x2407 | 0x0007 => {
                    if offset < grpprl.len() {
                        pap.page_break_before = grpprl[offset] != 0;
                        offset += 1;
                    }
                }
                // Widow control (sprmPFWidowControl)
                0x240E | 0x000E => {
                    if offset < grpprl.len() {
                        pap.widow_control = grpprl[offset] != 0;
                        offset += 1;
                    }
                }
                // Tab stops (sprmPChgTabsPapx)
                0xC615 | 0x0015 => {
                    // Tab stops are complex - simplified parsing
                    if offset < grpprl.len() {
                        let tab_count = grpprl[offset] as usize;
                        offset += 1;
                        // Each tab is 4 bytes: position (2) + alignment (1) + leader (1)
                        for _ in 0..tab_count.min(64) {
                            if offset + 3 < grpprl.len() {
                                let position = read_i16_le(grpprl, offset).unwrap_or(0) as i32;
                                let alignment_val = grpprl[offset + 2];
                                let leader_val = grpprl[offset + 3];

                                let alignment = match alignment_val {
                                    0 => TabAlignment::Left,
                                    1 => TabAlignment::Center,
                                    2 => TabAlignment::Right,
                                    3 => TabAlignment::Decimal,
                                    4 => TabAlignment::Bar,
                                    _ => TabAlignment::Left,
                                };

                                let leader = match leader_val {
                                    0 => TabLeader::None,
                                    1 => TabLeader::Dots,
                                    2 => TabLeader::Hyphens,
                                    3 => TabLeader::Underline,
                                    4 => TabLeader::Heavy,
                                    5 => TabLeader::MiddleDot,
                                    _ => TabLeader::None,
                                };

                                pap.tab_stops.push(TabStop {
                                    position,
                                    alignment,
                                    leader,
                                });

                                offset += 4;
                            }
                        }
                    }
                }
                // Unknown SPRM - skip based on size
                _ => {
                    let size = Self::get_sprm_size(sprm);
                    offset += size;
                }
            }
        }

        Ok(pap)
    }

    /// Get the size of an SPRM operand based on its opcode.
    fn get_sprm_size(sprm: u16) -> usize {
        // Extract type from SPRM (bits 0-2)
        let sprm_type = sprm & 0x07;
        match sprm_type {
            0 | 1 => 1,  // 1-byte operand
            2 | 4 | 5 => 2,  // 2-byte operand
            3 => 4,      // 4-byte operand
            6 => 1,      // Variable - default to 1
            7 => 3,      // 3-byte operand
            _ => 1,
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
        self.indent_first_line.map(|v| v as f32 / 1440.0).unwrap_or(0.0)
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

