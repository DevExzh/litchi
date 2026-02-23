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
}
