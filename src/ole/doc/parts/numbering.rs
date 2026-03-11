/// Numbering and list structures parser for Word binary format.
///
/// Based on Apache POI's ListTables and LibreOffice's implementation.
/// Lists in DOC files are defined by:
/// - List Format Override (LFO) structures
/// - List Format (LF) structures  
/// - List Level Format (LVL) structures
use super::super::package::{DocError, Result};
use super::fib::FileInformationBlock;
use crate::common::binary;

/// Number format for list levels
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u8)]
pub enum NumberFormat {
    /// Arabic numerals (1, 2, 3...)
    Arabic = 0,
    /// Uppercase Roman (I, II, III...)
    UpperRoman = 1,
    /// Lowercase Roman (i, ii, iii...)
    LowerRoman = 2,
    /// Uppercase letters (A, B, C...)
    UpperLetter = 3,
    /// Lowercase letters (a, b, c...)
    LowerLetter = 4,
    /// Ordinal numbers (1st, 2nd, 3rd...)
    Ordinal = 5,
    /// Cardinal text (One, Two, Three...)
    CardinalText = 6,
    /// Ordinal text (First, Second, Third...)
    OrdinalText = 7,
    /// Bullet
    Bullet = 23,
    /// No numbering
    None = 255,
}

impl NumberFormat {
    /// Convert from u8 value
    pub fn from_u8(value: u8) -> Self {
        match value {
            0 => NumberFormat::Arabic,
            1 => NumberFormat::UpperRoman,
            2 => NumberFormat::LowerRoman,
            3 => NumberFormat::UpperLetter,
            4 => NumberFormat::LowerLetter,
            5 => NumberFormat::Ordinal,
            6 => NumberFormat::CardinalText,
            7 => NumberFormat::OrdinalText,
            23 => NumberFormat::Bullet,
            255 => NumberFormat::None,
            _ => NumberFormat::Arabic, // Default to Arabic for unknown values
        }
    }
}

/// Alignment for list numbers/bullets
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ListAlignment {
    Left = 0,
    Center = 1,
    Right = 2,
}

impl ListAlignment {
    pub fn from_u8(value: u8) -> Self {
        match value {
            0 => ListAlignment::Left,
            1 => ListAlignment::Center,
            2 => ListAlignment::Right,
            _ => ListAlignment::Left,
        }
    }
}

/// List level format (LVLF structure)
#[derive(Debug, Clone)]
pub struct ListLevel {
    /// Start-at value
    pub start_at: u32,
    /// Number format
    pub number_format: NumberFormat,
    /// Alignment
    pub alignment: ListAlignment,
    /// Level number (0-8)
    pub level: u8,
    /// Follow character after number (tab, space, nothing)
    pub follow_char: u8,
    /// Indentation in twips
    pub indent_left: i32,
    /// Hanging indent in twips
    pub indent_hanging: i32,
    /// Number text (format string with placeholders)
    pub number_text: String,
}

impl ListLevel {
    /// Parse a list level from LVLF structure (28 bytes minimum)
    pub fn from_bytes(data: &[u8], level: u8) -> Result<Self> {
        if data.len() < 28 {
            return Err(DocError::InvalidFormat("LVLF too short".to_string()));
        }

        let start_at = binary::read_u32_le(data, 0)
            .map_err(|e| DocError::InvalidFormat(format!("Failed to read start_at: {}", e)))?;
        let number_format = NumberFormat::from_u8(data[4]);
        let alignment = ListAlignment::from_u8(data[5]);
        let follow_char = data[7];

        // Read indentation values (signed 32-bit)
        let indent_left = binary::read_i32_le(data, 12)
            .map_err(|e| DocError::InvalidFormat(format!("Failed to read indent_left: {}", e)))?;
        let indent_hanging = binary::read_i32_le(data, 16).map_err(|e| {
            DocError::InvalidFormat(format!("Failed to read indent_hanging: {}", e))
        })?;

        // Number text follows the fixed structure
        let number_text = if data.len() > 28 {
            // Read cbGrpprlChpx and cbGrpprlPapx to skip SPRM data
            let cb_chpx = data.get(25).copied().unwrap_or(0) as usize;
            let cb_papx = data.get(26).copied().unwrap_or(0) as usize;

            // Number text length is at offset 27
            let text_len = data.get(27).copied().unwrap_or(0) as usize;
            let text_offset = 28 + cb_chpx + cb_papx;

            if text_offset + text_len * 2 <= data.len() {
                // Number text is UTF-16LE
                let text_bytes = &data[text_offset..text_offset + text_len * 2];
                <String as Utf16LeExt>::from_utf16le_lossy(text_bytes)
            } else {
                String::new()
            }
        } else {
            String::new()
        };

        Ok(Self {
            start_at,
            number_format,
            alignment,
            level,
            follow_char,
            indent_left,
            indent_hanging,
            number_text,
        })
    }

    /// Check if this is a bullet list
    pub fn is_bullet(&self) -> bool {
        self.number_format == NumberFormat::Bullet
    }

    /// Check if this is a numbered list
    pub fn is_numbered(&self) -> bool {
        !self.is_bullet() && self.number_format != NumberFormat::None
    }
}

/// List structure (LST - List Structure)
#[derive(Debug, Clone)]
pub struct ListStructure {
    /// List ID (lsid)
    pub list_id: u32,
    /// Template ID (tplc)
    pub template_id: u32,
    /// Simple list flag
    pub is_simple: bool,
    /// List levels (up to 9 levels)
    pub levels: Vec<ListLevel>,
}

impl ListStructure {
    /// Parse a list structure from LST
    pub fn from_bytes(data: &[u8]) -> Result<Self> {
        if data.len() < 28 {
            return Err(DocError::InvalidFormat("LST too short".to_string()));
        }

        let list_id = binary::read_u32_le(data, 0)
            .map_err(|e| DocError::InvalidFormat(format!("Failed to read list_id: {}", e)))?;
        let template_id = binary::read_u32_le(data, 4)
            .map_err(|e| DocError::InvalidFormat(format!("Failed to read template_id: {}", e)))?;

        // Flags byte at offset 26
        let flags = data[26];
        let is_simple = (flags & 0x01) != 0;

        // Parse levels (LVL structures follow LST)
        let mut levels = Vec::new();
        let mut offset = 28;

        for level in 0..9 {
            if offset + 28 <= data.len() {
                if let Ok(lvl) = ListLevel::from_bytes(&data[offset..], level) {
                    levels.push(lvl);

                    // Calculate actual LVLF size to advance offset
                    // This is approximate - in reality we need to parse cbGrpprlChpx, cbGrpprlPapx
                    let cb_chpx = data.get(offset + 25).copied().unwrap_or(0) as usize;
                    let cb_papx = data.get(offset + 26).copied().unwrap_or(0) as usize;
                    let text_len = data.get(offset + 27).copied().unwrap_or(0) as usize;

                    offset += 28 + cb_chpx + cb_papx + text_len * 2;
                } else {
                    break;
                }
            } else {
                break;
            }
        }

        Ok(Self {
            list_id,
            template_id,
            is_simple,
            levels,
        })
    }

    /// Get a specific level
    pub fn level(&self, level: u8) -> Option<&ListLevel> {
        self.levels.get(level as usize)
    }
}

/// List Format Override (LFO structure)
#[derive(Debug, Clone)]
pub struct ListFormatOverride {
    /// List ID this override applies to
    pub list_id: u32,
    /// Override count
    pub override_count: u8,
    /// LFO ID (used to reference this from paragraphs)
    pub lfo_id: u32,
}

impl ListFormatOverride {
    /// Parse an LFO structure (12 bytes)
    pub fn from_bytes(data: &[u8]) -> Result<Self> {
        if data.len() < 12 {
            return Err(DocError::InvalidFormat("LFO too short".to_string()));
        }

        let list_id = binary::read_u32_le(data, 0)
            .map_err(|e| DocError::InvalidFormat(format!("Failed to read list_id: {}", e)))?;
        let override_count = data[8];
        let lfo_id = binary::read_u32_le(data, 8)
            .map_err(|e| DocError::InvalidFormat(format!("Failed to read lfo_id: {}", e)))?; // Actually at offset 8-11

        Ok(Self {
            list_id,
            override_count,
            lfo_id,
        })
    }
}

/// List tables parser
pub struct ListTables {
    /// All list structures
    list_structures: Vec<ListStructure>,
    /// All list format overrides
    list_overrides: Vec<ListFormatOverride>,
}

impl ListTables {
    /// Parse list tables from the table stream
    ///
    /// # Arguments
    ///
    /// * `fib` - File Information Block
    /// * `table_stream` - The table stream (0Table or 1Table)
    pub fn parse(fib: &FileInformationBlock, table_stream: &[u8]) -> Result<Self> {
        let mut list_structures = Vec::new();
        let mut list_overrides = Vec::new();

        // Parse PlfLst (List Table) - FIB index 27
        if let Some((offset, length)) = fib.get_table_pointer(27)
            && length > 0
            && (offset as usize) < table_stream.len()
        {
            let plf_data = &table_stream[offset as usize..];
            let plf_len = length.min((table_stream.len() - offset as usize) as u32) as usize;

            list_structures = Self::parse_plflst(&plf_data[..plf_len])?;
        }

        // Parse PlfLfo (List Format Override Table) - FIB index 28
        if let Some((offset, length)) = fib.get_table_pointer(28)
            && length > 0
            && (offset as usize) < table_stream.len()
        {
            let plf_data = &table_stream[offset as usize..];
            let plf_len = length.min((table_stream.len() - offset as usize) as u32) as usize;

            list_overrides = Self::parse_plflfo(&plf_data[..plf_len])?;
        }

        Ok(Self {
            list_structures,
            list_overrides,
        })
    }

    /// Parse PlfLst (List Table)
    fn parse_plflst(data: &[u8]) -> Result<Vec<ListStructure>> {
        if data.len() < 2 {
            return Ok(Vec::new());
        }

        let count = binary::read_u16_le(data, 0)
            .map_err(|e| DocError::InvalidFormat(format!("Failed to read count: {}", e)))?
            as usize;
        let mut structures = Vec::with_capacity(count);
        let mut offset = 2;

        for _ in 0..count {
            if offset >= data.len() {
                break;
            }

            // Each LST is variable length, parse and advance
            if let Ok(lst) = ListStructure::from_bytes(&data[offset..]) {
                // Calculate size (this is approximate)
                let mut lst_size = 28;
                for _level in &lst.levels {
                    lst_size += 28; // Base LVLF
                    // Add SPRM and text sizes (simplified)
                }

                structures.push(lst);
                offset += lst_size.min(data.len() - offset);
            } else {
                break;
            }
        }

        Ok(structures)
    }

    /// Parse PlfLfo (List Format Override Table)
    fn parse_plflfo(data: &[u8]) -> Result<Vec<ListFormatOverride>> {
        if data.len() < 4 {
            return Ok(Vec::new());
        }

        let count = binary::read_u32_le(data, 0)
            .map_err(|e| DocError::InvalidFormat(format!("Failed to read count: {}", e)))?
            as usize;
        let mut overrides = Vec::with_capacity(count);
        let mut offset = 4;

        for _ in 0..count {
            if offset + 12 > data.len() {
                break;
            }

            if let Ok(lfo) = ListFormatOverride::from_bytes(&data[offset..]) {
                overrides.push(lfo);
                offset += 12;
            } else {
                break;
            }
        }

        Ok(overrides)
    }

    /// Get all list structures
    pub fn structures(&self) -> &[ListStructure] {
        &self.list_structures
    }

    /// Get all list format overrides
    pub fn overrides(&self) -> &[ListFormatOverride] {
        &self.list_overrides
    }

    /// Find a list structure by ID
    pub fn find_structure(&self, list_id: u32) -> Option<&ListStructure> {
        self.list_structures
            .iter()
            .find(|lst| lst.list_id == list_id)
    }

    /// Find a list override by LFO ID
    pub fn find_override(&self, lfo_id: u32) -> Option<&ListFormatOverride> {
        self.list_overrides.iter().find(|lfo| lfo.lfo_id == lfo_id)
    }

    /// Get the list structure for a given LFO ID
    pub fn get_list_for_lfo(&self, lfo_id: u32) -> Option<&ListStructure> {
        self.find_override(lfo_id)
            .and_then(|lfo| self.find_structure(lfo.list_id))
    }
}

/// Helper trait for UTF-16LE string conversion
trait Utf16LeExt {
    fn from_utf16le_lossy(bytes: &[u8]) -> String;
}

impl Utf16LeExt for String {
    fn from_utf16le_lossy(bytes: &[u8]) -> String {
        let mut u16_vec = Vec::with_capacity(bytes.len() / 2);
        for chunk in bytes.chunks_exact(2) {
            let val = u16::from_le_bytes([chunk[0], chunk[1]]);
            u16_vec.push(val);
        }
        String::from_utf16_lossy(&u16_vec)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_number_format() {
        assert_eq!(NumberFormat::from_u8(0), NumberFormat::Arabic);
        assert_eq!(NumberFormat::from_u8(23), NumberFormat::Bullet);
        assert_eq!(NumberFormat::from_u8(255), NumberFormat::None);
    }

    #[test]
    fn test_list_alignment() {
        assert_eq!(ListAlignment::from_u8(0), ListAlignment::Left);
        assert_eq!(ListAlignment::from_u8(1), ListAlignment::Center);
        assert_eq!(ListAlignment::from_u8(2), ListAlignment::Right);
    }

    #[test]
    fn test_number_format_all_variants() {
        assert_eq!(NumberFormat::from_u8(0), NumberFormat::Arabic);
        assert_eq!(NumberFormat::from_u8(1), NumberFormat::UpperRoman);
        assert_eq!(NumberFormat::from_u8(2), NumberFormat::LowerRoman);
        assert_eq!(NumberFormat::from_u8(3), NumberFormat::UpperLetter);
        assert_eq!(NumberFormat::from_u8(4), NumberFormat::LowerLetter);
        assert_eq!(NumberFormat::from_u8(5), NumberFormat::Ordinal);
        assert_eq!(NumberFormat::from_u8(6), NumberFormat::CardinalText);
        assert_eq!(NumberFormat::from_u8(7), NumberFormat::OrdinalText);
        assert_eq!(NumberFormat::from_u8(23), NumberFormat::Bullet);
        assert_eq!(NumberFormat::from_u8(255), NumberFormat::None);
    }

    #[test]
    fn test_number_format_default_for_unknown() {
        // Unknown values default to Arabic
        assert_eq!(NumberFormat::from_u8(8), NumberFormat::Arabic);
        assert_eq!(NumberFormat::from_u8(100), NumberFormat::Arabic);
        assert_eq!(NumberFormat::from_u8(254), NumberFormat::Arabic);
    }

    #[test]
    fn test_number_format_clone() {
        let fmt = NumberFormat::Bullet;
        let cloned = fmt.clone();
        assert_eq!(fmt, cloned);
    }

    #[test]
    fn test_number_format_copy() {
        let fmt = NumberFormat::UpperRoman;
        let copied = fmt;
        assert_eq!(fmt, copied);
    }

    #[test]
    fn test_number_format_debug() {
        let fmt = NumberFormat::Arabic;
        let debug_str = format!("{:?}", fmt);
        assert!(debug_str.contains("Arabic"));
    }

    #[test]
    fn test_number_format_equality() {
        assert_eq!(NumberFormat::Arabic, NumberFormat::Arabic);
        assert_ne!(NumberFormat::Arabic, NumberFormat::Bullet);
    }

    #[test]
    fn test_list_alignment_default_for_unknown() {
        assert_eq!(ListAlignment::from_u8(3), ListAlignment::Left);
        assert_eq!(ListAlignment::from_u8(100), ListAlignment::Left);
    }

    #[test]
    fn test_list_alignment_clone() {
        let align = ListAlignment::Center;
        let cloned = align.clone();
        assert_eq!(align, cloned);
    }

    #[test]
    fn test_list_level_creation() {
        let level = ListLevel {
            start_at: 1,
            number_format: NumberFormat::Arabic,
            alignment: ListAlignment::Left,
            level: 0,
            follow_char: 0,
            indent_left: 720,
            indent_hanging: 360,
            number_text: "%1.".to_string(),
        };

        assert_eq!(level.start_at, 1);
        assert_eq!(level.number_format, NumberFormat::Arabic);
        assert_eq!(level.alignment, ListAlignment::Left);
        assert_eq!(level.level, 0);
        assert_eq!(level.indent_left, 720);
        assert_eq!(level.indent_hanging, 360);
        assert_eq!(level.number_text, "%1.");
        assert!(level.is_numbered());
        assert!(!level.is_bullet());
    }

    #[test]
    fn test_list_level_bullet() {
        let level = ListLevel {
            start_at: 1,
            number_format: NumberFormat::Bullet,
            alignment: ListAlignment::Left,
            level: 0,
            follow_char: 0,
            indent_left: 720,
            indent_hanging: 360,
            number_text: "\u{2022}".to_string(),
        };

        assert!(level.is_bullet());
        assert!(!level.is_numbered());
    }

    #[test]
    fn test_list_level_none() {
        let level = ListLevel {
            start_at: 0,
            number_format: NumberFormat::None,
            alignment: ListAlignment::Left,
            level: 0,
            follow_char: 0,
            indent_left: 0,
            indent_hanging: 0,
            number_text: String::new(),
        };

        assert!(!level.is_bullet());
        assert!(!level.is_numbered());
    }

    #[test]
    fn test_list_level_clone() {
        let level = ListLevel {
            start_at: 1,
            number_format: NumberFormat::LowerRoman,
            alignment: ListAlignment::Right,
            level: 2,
            follow_char: 1,
            indent_left: 1440,
            indent_hanging: 720,
            number_text: "(%2)".to_string(),
        };
        let cloned = level.clone();

        assert_eq!(cloned.start_at, level.start_at);
        assert_eq!(cloned.number_format, level.number_format);
        assert_eq!(cloned.alignment, level.alignment);
        assert_eq!(cloned.level, level.level);
        assert_eq!(cloned.number_text, level.number_text);
    }

    #[test]
    fn test_list_level_debug() {
        let level = ListLevel {
            start_at: 1,
            number_format: NumberFormat::Arabic,
            alignment: ListAlignment::Left,
            level: 0,
            follow_char: 0,
            indent_left: 720,
            indent_hanging: 360,
            number_text: "%1.".to_string(),
        };
        let debug_str = format!("{:?}", level);
        assert!(debug_str.contains("ListLevel"));
        assert!(debug_str.contains("Arabic"));
    }

    #[test]
    fn test_list_structure_creation() {
        let levels = vec![ListLevel {
            start_at: 1,
            number_format: NumberFormat::Arabic,
            alignment: ListAlignment::Left,
            level: 0,
            follow_char: 0,
            indent_left: 720,
            indent_hanging: 360,
            number_text: "%1.".to_string(),
        }];

        let lst = ListStructure {
            list_id: 12345,
            template_id: 67890,
            is_simple: false,
            levels,
        };

        assert_eq!(lst.list_id, 12345);
        assert_eq!(lst.template_id, 67890);
        assert!(!lst.is_simple);
        assert_eq!(lst.levels.len(), 1);
    }

    #[test]
    fn test_list_structure_simple() {
        let lst = ListStructure {
            list_id: 1,
            template_id: 1,
            is_simple: true,
            levels: Vec::new(),
        };

        assert!(lst.is_simple);
    }

    #[test]
    fn test_list_structure_level_accessor() {
        let levels = vec![
            ListLevel {
                start_at: 1,
                number_format: NumberFormat::Arabic,
                alignment: ListAlignment::Left,
                level: 0,
                follow_char: 0,
                indent_left: 720,
                indent_hanging: 360,
                number_text: "%1.".to_string(),
            },
            ListLevel {
                start_at: 1,
                number_format: NumberFormat::LowerLetter,
                alignment: ListAlignment::Left,
                level: 1,
                follow_char: 0,
                indent_left: 1440,
                indent_hanging: 360,
                number_text: "%1.%2.".to_string(),
            },
        ];

        let lst = ListStructure {
            list_id: 1,
            template_id: 1,
            is_simple: false,
            levels,
        };

        assert!(lst.level(0).is_some());
        assert!(lst.level(1).is_some());
        assert!(lst.level(2).is_none());
        assert_eq!(lst.level(0).unwrap().number_format, NumberFormat::Arabic);
        assert_eq!(
            lst.level(1).unwrap().number_format,
            NumberFormat::LowerLetter
        );
    }

    #[test]
    fn test_list_structure_clone() {
        let lst = ListStructure {
            list_id: 100,
            template_id: 200,
            is_simple: false,
            levels: vec![ListLevel {
                start_at: 1,
                number_format: NumberFormat::Bullet,
                alignment: ListAlignment::Left,
                level: 0,
                follow_char: 0,
                indent_left: 720,
                indent_hanging: 360,
                number_text: "\u{2022}".to_string(),
            }],
        };
        let cloned = lst.clone();

        assert_eq!(cloned.list_id, lst.list_id);
        assert_eq!(cloned.template_id, lst.template_id);
        assert_eq!(cloned.levels.len(), lst.levels.len());
    }

    #[test]
    fn test_list_structure_debug() {
        let lst = ListStructure {
            list_id: 1,
            template_id: 2,
            is_simple: false,
            levels: Vec::new(),
        };
        let debug_str = format!("{:?}", lst);
        assert!(debug_str.contains("ListStructure"));
    }

    #[test]
    fn test_list_format_override_creation() {
        let lfo = ListFormatOverride {
            list_id: 12345,
            override_count: 1,
            lfo_id: 1,
        };

        assert_eq!(lfo.list_id, 12345);
        assert_eq!(lfo.override_count, 1);
        assert_eq!(lfo.lfo_id, 1);
    }

    #[test]
    fn test_list_format_override_clone() {
        let lfo = ListFormatOverride {
            list_id: 100,
            override_count: 2,
            lfo_id: 5,
        };
        let cloned = lfo.clone();

        assert_eq!(cloned.list_id, lfo.list_id);
        assert_eq!(cloned.override_count, lfo.override_count);
        assert_eq!(cloned.lfo_id, lfo.lfo_id);
    }

    #[test]
    fn test_list_format_override_debug() {
        let lfo = ListFormatOverride {
            list_id: 1,
            override_count: 0,
            lfo_id: 1,
        };
        let debug_str = format!("{:?}", lfo);
        assert!(debug_str.contains("ListFormatOverride"));
    }

    #[test]
    fn test_list_tables_empty() {
        let tables = ListTables {
            list_structures: Vec::new(),
            list_overrides: Vec::new(),
        };

        assert!(tables.structures().is_empty());
        assert!(tables.overrides().is_empty());
    }

    #[test]
    fn test_list_tables_with_data() {
        let tables = ListTables {
            list_structures: vec![ListStructure {
                list_id: 1,
                template_id: 1,
                is_simple: false,
                levels: Vec::new(),
            }],
            list_overrides: vec![ListFormatOverride {
                list_id: 1,
                override_count: 0,
                lfo_id: 1,
            }],
        };

        assert_eq!(tables.structures().len(), 1);
        assert_eq!(tables.overrides().len(), 1);
    }

    #[test]
    fn test_list_tables_find_structure() {
        let tables = ListTables {
            list_structures: vec![
                ListStructure {
                    list_id: 100,
                    template_id: 1,
                    is_simple: false,
                    levels: Vec::new(),
                },
                ListStructure {
                    list_id: 200,
                    template_id: 2,
                    is_simple: true,
                    levels: Vec::new(),
                },
            ],
            list_overrides: Vec::new(),
        };

        assert!(tables.find_structure(100).is_some());
        assert!(tables.find_structure(200).is_some());
        assert!(tables.find_structure(999).is_none());
    }

    #[test]
    fn test_list_tables_find_override() {
        let tables = ListTables {
            list_structures: Vec::new(),
            list_overrides: vec![
                ListFormatOverride {
                    list_id: 1,
                    override_count: 0,
                    lfo_id: 10,
                },
                ListFormatOverride {
                    list_id: 2,
                    override_count: 1,
                    lfo_id: 20,
                },
            ],
        };

        assert!(tables.find_override(10).is_some());
        assert!(tables.find_override(20).is_some());
        assert!(tables.find_override(999).is_none());
    }

    #[test]
    fn test_list_tables_get_list_for_lfo() {
        let tables = ListTables {
            list_structures: vec![ListStructure {
                list_id: 100,
                template_id: 1,
                is_simple: false,
                levels: Vec::new(),
            }],
            list_overrides: vec![ListFormatOverride {
                list_id: 100,
                override_count: 0,
                lfo_id: 1,
            }],
        };

        let lst = tables.get_list_for_lfo(1);
        assert!(lst.is_some());
        assert_eq!(lst.unwrap().list_id, 100);

        assert!(tables.get_list_for_lfo(999).is_none());
    }

    #[test]
    fn test_list_tables_get_list_for_lfo_no_override() {
        let tables = ListTables {
            list_structures: vec![ListStructure {
                list_id: 100,
                template_id: 1,
                is_simple: false,
                levels: Vec::new(),
            }],
            list_overrides: Vec::new(),
        };

        assert!(tables.get_list_for_lfo(1).is_none());
    }

    #[test]
    fn test_utf16le_ext_empty() {
        let result = <String as Utf16LeExt>::from_utf16le_lossy(b"");
        assert_eq!(result, "");
    }

    #[test]
    fn test_utf16le_ext_single_char() {
        // 'A' in UTF-16LE
        let result = <String as Utf16LeExt>::from_utf16le_lossy(b"A\0");
        assert_eq!(result, "A");
    }

    #[test]
    fn test_utf16le_ext_multiple_chars() {
        // "ABC" in UTF-16LE
        let bytes = vec!['A' as u16, 'B' as u16, 'C' as u16]
            .iter()
            .flat_map(|c| c.to_le_bytes())
            .collect::<Vec<_>>();
        let result = <String as Utf16LeExt>::from_utf16le_lossy(&bytes);
        assert_eq!(result, "ABC");
    }

    #[test]
    fn test_utf16le_ext_unicode() {
        // Unicode test in UTF-16LE
        let bytes: Vec<u8> = "Test"
            .encode_utf16()
            .flat_map(|c| c.to_le_bytes())
            .collect();
        let result = <String as Utf16LeExt>::from_utf16le_lossy(&bytes);
        assert_eq!(result, "Test");
    }

    #[test]
    fn test_list_level_from_bytes_too_short() {
        let data = vec![0u8; 10];
        let result = ListLevel::from_bytes(&data, 0);
        assert!(result.is_err());
    }

    #[test]
    fn test_list_level_from_bytes_minimal() {
        // Create a minimal valid LVLF structure (28 bytes minimum)
        let mut data = vec![0u8; 28];
        // start_at at offset 0
        data[0] = 1; // start_at = 1
        // number_format at offset 4
        data[4] = 0; // Arabic
        // alignment at offset 5
        data[5] = 0; // Left
        // follow_char at offset 7
        data[7] = 0;
        // indent_left at offset 12
        data[12] = 0xD0; // 720 in little-endian
        data[13] = 0x02;
        // indent_hanging at offset 16
        data[16] = 0x68; // 360 in little-endian
        data[17] = 0x01;

        let result = ListLevel::from_bytes(&data, 0);
        assert!(result.is_ok());

        let level = result.unwrap();
        assert_eq!(level.start_at, 1);
        assert_eq!(level.number_format, NumberFormat::Arabic);
        assert_eq!(level.alignment, ListAlignment::Left);
        assert_eq!(level.level, 0);
        assert_eq!(level.indent_left, 720);
        assert_eq!(level.indent_hanging, 360);
        assert_eq!(level.number_text, "");
    }

    #[test]
    fn test_list_level_from_bytes_bullet() {
        let mut data = vec![0u8; 28];
        data[4] = 23; // Bullet format

        let level = ListLevel::from_bytes(&data, 0).unwrap();
        assert!(level.is_bullet());
        assert!(!level.is_numbered());
    }

    #[test]
    fn test_list_level_from_bytes_with_text() {
        let mut data = vec![0u8; 40];
        // Fixed part
        data[0] = 1; // start_at
        data[4] = 0; // Arabic
        data[5] = 0; // Left
        data[7] = 0; // follow_char
        // cbGrpprlChpx at offset 25
        data[25] = 0;
        // cbGrpprlPapx at offset 26
        data[26] = 0;
        // text length at offset 27
        data[27] = 3; // 3 characters

        // Add UTF-16LE text at offset 28: "%1."
        let text = "%1.";
        let text_offset = 28;
        for (i, c) in text.encode_utf16().enumerate() {
            let bytes = c.to_le_bytes();
            data[text_offset + i * 2] = bytes[0];
            data[text_offset + i * 2 + 1] = bytes[1];
        }

        let level = ListLevel::from_bytes(&data, 0).unwrap();
        assert_eq!(level.number_text, "%1.");
    }

    #[test]
    fn test_list_structure_from_bytes_too_short() {
        let data = vec![0u8; 10];
        let result = ListStructure::from_bytes(&data);
        assert!(result.is_err());
    }

    #[test]
    fn test_list_structure_from_bytes_minimal() {
        let mut data = vec![0u8; 28];
        // list_id at offset 0
        data[0] = 0x39; // 57 in little-endian
        data[1] = 0x00;
        data[2] = 0x00;
        data[3] = 0x00;
        // template_id at offset 4
        data[4] = 0x30; // 48 in little-endian
        data[5] = 0x00;
        // flags at offset 26 - simple flag
        data[26] = 0x01; // is_simple = true

        let result = ListStructure::from_bytes(&data);
        assert!(result.is_ok());

        let lst = result.unwrap();
        assert_eq!(lst.list_id, 57);
        assert_eq!(lst.template_id, 48);
        assert!(lst.is_simple);
        assert!(lst.levels.is_empty());
    }

    #[test]
    fn test_list_format_override_from_bytes_too_short() {
        let data = vec![0u8; 5];
        let result = ListFormatOverride::from_bytes(&data);
        assert!(result.is_err());
    }

    #[test]
    fn test_list_format_override_from_bytes_valid() {
        let mut data = vec![0u8; 12];
        // list_id at offset 0
        data[0] = 0x39;
        data[1] = 0x00;
        data[2] = 0x00;
        data[3] = 0x00;
        // override_count at offset 8
        data[8] = 2;
        // lfo_id at offset 8-11 (overlaps with override_count in this structure)

        let result = ListFormatOverride::from_bytes(&data);
        assert!(result.is_ok());

        let lfo = result.unwrap();
        assert_eq!(lfo.list_id, 57);
        assert_eq!(lfo.override_count, 2);
    }

    #[test]
    fn test_numbering_with_unicode_number_text() {
        let level = ListLevel {
            start_at: 1,
            number_format: NumberFormat::Bullet,
            alignment: ListAlignment::Left,
            level: 0,
            follow_char: 0,
            indent_left: 720,
            indent_hanging: 360,
            number_text: "\u{2022} \u{25ba} \u{2192}".to_string(), // bullet, pointer, arrow
        };

        assert_eq!(level.number_text, "\u{2022} \u{25ba} \u{2192}");
    }

    #[test]
    fn test_list_level_negative_indent() {
        let mut data = vec![0u8; 28];
        // indent_left at offset 12 (signed 32-bit)
        data[12] = 0xF0; // -16 in little-endian two's complement
        data[13] = 0xFF;
        data[14] = 0xFF;
        data[15] = 0xFF;

        let level = ListLevel::from_bytes(&data, 0).unwrap();
        assert_eq!(level.indent_left, -16);
    }
}
