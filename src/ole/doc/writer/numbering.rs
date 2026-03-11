//! List numbering writer for DOC files
//!
//! Generates list structures (LST, LVL) and format overrides (LFO, LFOLVL).

use std::io::Write;

/// Number format type
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum NumberFormat {
    Decimal = 0,
    UpperRoman = 1,
    LowerRoman = 2,
    UpperLetter = 3,
    LowerLetter = 4,
    Ordinal = 5,
    Bullet = 23,
}

/// List level definition
#[derive(Debug, Clone)]
pub struct ListLevel {
    /// Starting number
    pub start_at: u32,
    /// Number format
    pub number_format: NumberFormat,
    /// Number text (e.g., "%1." for "1.")
    pub number_text: String,
    /// Left indent in twips
    pub indent_left: i32,
    /// Hanging indent in twips
    pub indent_hanging: i32,
}

impl ListLevel {
    /// Create a new list level
    pub fn new(start_at: u32, number_format: NumberFormat) -> Self {
        Self {
            start_at,
            number_format,
            number_text: String::from("%1."),
            indent_left: 720,     // 0.5 inch
            indent_hanging: -360, // -0.25 inch
        }
    }

    /// Serialize to LVL structure per MS-DOC spec [2.9.150].
    ///
    /// A complete LVL consists of:
    /// 1. LVLF (28 bytes fixed) — level format descriptor
    /// 2. grpprlPapx (cbGrpprlPapx bytes) — paragraph SPRMs
    /// 3. grpprlChpx (cbGrpprlChpx bytes) — character SPRMs
    /// 4. xst — number text as counted string: cch (u16 LE) + cch UTF-16LE chars
    ///
    /// The number text uses placeholder characters 0x0000–0x0008 for levels 0–8.
    /// For example, `"%1."` becomes `[0x0000, u'.']` (level 0 counter + period).
    pub fn to_bytes(&self) -> Vec<u8> {
        // Convert user-facing number_text to internal xst format.
        // "%1" → char 0x0000 (level 0 placeholder), "%2" → 0x0001, etc.
        let mut xst_chars: Vec<u16> = Vec::new();
        let mut rgbxch_nums = [0u8; 9]; // 1-based positions of level placeholders in xst
        let src: Vec<char> = self.number_text.chars().collect();
        let mut i = 0;
        while i < src.len() {
            if src[i] == '%' && i + 1 < src.len() && src[i + 1].is_ascii_digit() {
                let level_1based = (src[i + 1] as u8) - b'0'; // 1-based level number
                if (1..=9).contains(&level_1based) {
                    let level_idx = (level_1based - 1) as usize; // 0-based
                    // Record 1-based position in xst for this level placeholder
                    rgbxch_nums[level_idx] = (xst_chars.len() + 1) as u8;
                    xst_chars.push(level_idx as u16); // placeholder char = 0-based level
                }
                i += 2;
            } else {
                xst_chars.push(src[i] as u16);
                i += 1;
            }
        }

        // For bullet format, override with bullet character (no level placeholder)
        if self.number_format == NumberFormat::Bullet {
            xst_chars.clear();
            xst_chars.push(0x2022); // •
            rgbxch_nums = [0u8; 9]; // no level placeholders for bullets
        }

        // No grpprl SPRMs for simplicity (LVL-level indents come from paragraph SPRMs)
        let cb_grpprl_papx: u8 = 0;
        let cb_grpprl_chpx: u8 = 0;

        let mut buf = Vec::with_capacity(28 + 2 + xst_chars.len() * 2);

        // === LVLF (exactly 28 bytes) per MS-DOC 2.9.150 ===
        // Offset 0: iStartAt (4 bytes)
        buf.write_all(&self.start_at.to_le_bytes()).unwrap();
        // Offset 4: nfc (1 byte) — number format code
        buf.push(self.number_format as u8);
        // Offset 5: jc:2 + flags:6 (1 byte) — left-aligned (jc=0), no flags
        buf.push(0x00);
        // Offset 6: rgbxchNums[9] (9 bytes) — placeholder positions
        buf.write_all(&rgbxch_nums).unwrap();
        // Offset 15: ixchFollow (1 byte) — 0=tab, 1=space, 2=nothing
        buf.push(0x00); // tab follow
        // Offset 16: dxaIndentSav (4 bytes, i32 LE)
        buf.write_all(&0i32.to_le_bytes()).unwrap();
        // Offset 20: reserved2 (4 bytes)
        buf.write_all(&0u32.to_le_bytes()).unwrap();
        // Offset 24: cbGrpprlChpx (1 byte)
        buf.push(cb_grpprl_chpx);
        // Offset 25: cbGrpprlPapx (1 byte)
        buf.push(cb_grpprl_papx);
        // Offset 26: ixchLim (1 byte, unused, must be 0)
        buf.push(0);
        // Offset 27: nfcOrig (1 byte, unused)
        buf.push(self.number_format as u8);

        // grpprlPapx (cbGrpprlPapx bytes) — empty for now
        // grpprlChpx (cbGrpprlChpx bytes) — empty for now

        // xst: cch (u16 LE) + cch UTF-16LE characters
        buf.write_all(&(xst_chars.len() as u16).to_le_bytes())
            .unwrap();
        for &ch in &xst_chars {
            buf.write_all(&ch.to_le_bytes()).unwrap();
        }

        buf
    }
}

/// List structure definition
#[derive(Debug, Clone)]
pub struct ListStructure {
    /// List ID (unique identifier)
    pub list_id: u32,
    /// Template ID
    pub template_id: u32,
    /// List levels (up to 9)
    pub levels: Vec<ListLevel>,
}

impl ListStructure {
    /// Create a new list structure
    pub fn new(list_id: u32) -> Self {
        Self {
            list_id,
            template_id: list_id,
            levels: Vec::new(),
        }
    }

    /// Add a level
    pub fn add_level(&mut self, level: ListLevel) {
        if self.levels.len() < 9 {
            self.levels.push(level);
        }
    }

    /// Serialize to LSTF structure (fixed 28 bytes, per MS-DOC spec).
    ///
    /// This does NOT include the variable-length LVL data — use
    /// [`levels_to_bytes`] for that.
    pub fn to_bytes(&self) -> Vec<u8> {
        let mut buf = Vec::with_capacity(28);

        // List ID (4 bytes)
        buf.write_all(&self.list_id.to_le_bytes()).unwrap();

        // Template ID (4 bytes)
        buf.write_all(&self.template_id.to_le_bytes()).unwrap();

        // 9 RGISTDs (18 bytes) - style IDs for each level (0xFFFF = no style)
        buf.write_all(&[0xff; 18]).unwrap();

        // Flags (1 byte):
        //   bit 0: fSimpleList (1 = single-level, 0 = multi-level)
        //   bit 1: fRestartHdn
        //   bit 2: fAutoNum (unused)
        //   bits 3-7: reserved
        let f_simple = if self.levels.len() <= 1 {
            0x01u8
        } else {
            0x00u8
        };
        buf.push(f_simple);

        // grfhic (1 byte) — reserved/compatibility, set to 0
        buf.push(0);

        buf
    }

    /// Serialize the LVL array for this list structure.
    ///
    /// Per MS-DOC spec, LVLs are appended after all LSTFs in the PlfLst
    /// and are NOT counted in `lcbPlfLst`.
    pub fn levels_to_bytes(&self) -> Vec<u8> {
        let mut buf = Vec::new();
        for level in &self.levels {
            buf.extend_from_slice(&level.to_bytes());
        }
        buf
    }
}

/// List format override
#[derive(Debug, Clone)]
pub struct ListFormatOverride {
    /// List ID this override applies to
    pub list_id: u32,
    /// Override ID
    pub lfo_id: u32,
}

impl ListFormatOverride {
    /// Create a new list format override
    pub fn new(list_id: u32, lfo_id: u32) -> Self {
        Self { list_id, lfo_id }
    }

    /// Serialize to LFO structure (16 bytes)
    pub fn to_bytes(&self) -> Vec<u8> {
        let mut buf = Vec::new();

        // List ID (4 bytes)
        buf.write_all(&self.list_id.to_le_bytes()).unwrap();

        // Reserved (8 bytes)
        buf.write_all(&[0; 8]).unwrap();

        // Level count (1 byte) - 0 means use all from LST
        buf.push(0);

        // Reserved (3 bytes)
        buf.write_all(&[0; 3]).unwrap();

        buf
    }
}

/// Numbering writer for list tables
#[derive(Debug)]
pub struct NumberingWriter {
    list_structures: Vec<ListStructure>,
    list_overrides: Vec<ListFormatOverride>,
}

impl NumberingWriter {
    /// Create a new numbering writer
    pub fn new() -> Self {
        Self {
            list_structures: Vec::new(),
            list_overrides: Vec::new(),
        }
    }

    /// Add a list structure
    pub fn add_list(&mut self, list: ListStructure) {
        self.list_structures.push(list);
    }

    /// Add a list format override
    pub fn add_override(&mut self, lfo: ListFormatOverride) {
        self.list_overrides.push(lfo);
    }

    /// Get number of list structures
    pub fn list_count(&self) -> usize {
        self.list_structures.len()
    }

    /// Generate PlfLst (List Table).
    ///
    /// Returns `(plflst_for_lcb, lvl_data)` where:
    /// - `plflst_for_lcb` = cLst (u16) + LSTF array (28 bytes each) — this is what
    ///   `lcbPlfLst` should cover.
    /// - `lvl_data` = LVL array for all lists — appended immediately after but
    ///   NOT counted in `lcbPlfLst` per MS-DOC spec / Apache POI.
    pub fn build_plflst(&self) -> (Vec<u8>, Vec<u8>) {
        let mut header_buf = Vec::new();
        let mut lvl_buf = Vec::new();

        // Count of lists (2 bytes)
        header_buf
            .write_all(&(self.list_structures.len() as u16).to_le_bytes())
            .unwrap();

        // Each LSTF (fixed 28 bytes)
        for list in &self.list_structures {
            header_buf.extend_from_slice(&list.to_bytes());
        }

        // LVL data for all lists
        for list in &self.list_structures {
            lvl_buf.extend_from_slice(&list.levels_to_bytes());
        }

        (header_buf, lvl_buf)
    }

    /// Generate PlfLfo (List Format Override Table)
    pub fn build_plflfo(&self) -> Vec<u8> {
        let mut buf = Vec::new();

        // Count of overrides (4 bytes)
        buf.write_all(&(self.list_overrides.len() as u32).to_le_bytes())
            .unwrap();

        // Each override
        for lfo in &self.list_overrides {
            buf.extend_from_slice(&lfo.to_bytes());
        }

        buf
    }

    /// Check if empty
    pub fn is_empty(&self) -> bool {
        self.list_structures.is_empty()
    }
}

impl Default for NumberingWriter {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_number_format_variants() {
        assert_eq!(NumberFormat::Decimal as u8, 0);
        assert_eq!(NumberFormat::UpperRoman as u8, 1);
        assert_eq!(NumberFormat::LowerRoman as u8, 2);
        assert_eq!(NumberFormat::UpperLetter as u8, 3);
        assert_eq!(NumberFormat::LowerLetter as u8, 4);
        assert_eq!(NumberFormat::Ordinal as u8, 5);
        assert_eq!(NumberFormat::Bullet as u8, 23);
    }

    #[test]
    fn test_list_level_new() {
        let level = ListLevel::new(1, NumberFormat::Decimal);
        assert_eq!(level.start_at, 1);
        assert_eq!(level.number_format, NumberFormat::Decimal);
        assert_eq!(level.number_text, "%1.");
        assert_eq!(level.indent_left, 720);
        assert_eq!(level.indent_hanging, -360);
    }

    #[test]
    fn test_list_level_to_bytes_basic() {
        let level = ListLevel::new(1, NumberFormat::Decimal);
        let bytes = level.to_bytes();

        // LVLF is 28 bytes + xst length
        assert!(bytes.len() >= 30); // 28 + at least 2 for cch

        // Check iStartAt (offset 0, 4 bytes)
        assert_eq!(
            u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]),
            1
        );

        // Check nfc (offset 4)
        assert_eq!(bytes[4], 0); // Decimal = 0
    }

    #[test]
    fn test_list_level_bullet_format() {
        let mut level = ListLevel::new(1, NumberFormat::Bullet);
        level.number_text = "•".to_string();
        let bytes = level.to_bytes();

        // Bullet format should have bullet character 0x2022
        assert_eq!(bytes[4], 23); // Bullet = 23
    }

    #[test]
    fn test_list_level_with_level_placeholder() {
        let mut level = ListLevel::new(1, NumberFormat::Decimal);
        level.number_text = "%1.".to_string();
        let bytes = level.to_bytes();

        // Should generate valid LVL structure
        assert!(bytes.len() > 28);
    }

    #[test]
    fn test_list_level_multi_level_placeholder() {
        let mut level = ListLevel::new(1, NumberFormat::Decimal);
        level.number_text = "%1.%2.%3.".to_string();
        let bytes = level.to_bytes();

        // Should handle multiple level placeholders
        assert!(bytes.len() >= 28);
    }

    #[test]
    fn test_list_structure_new() {
        let list = ListStructure::new(42);
        assert_eq!(list.list_id, 42);
        assert_eq!(list.template_id, 42);
        assert!(list.levels.is_empty());
    }

    #[test]
    fn test_list_structure_add_level() {
        let mut list = ListStructure::new(1);
        let level = ListLevel::new(1, NumberFormat::Decimal);
        list.add_level(level);

        assert_eq!(list.levels.len(), 1);
    }

    #[test]
    fn test_list_structure_max_levels() {
        let mut list = ListStructure::new(1);
        for i in 0..15 {
            list.add_level(ListLevel::new(i as u32 + 1, NumberFormat::Decimal));
        }
        // Should only have 9 levels max
        assert_eq!(list.levels.len(), 9);
    }

    #[test]
    fn test_list_structure_to_bytes() {
        let mut list = ListStructure::new(0x12345678);
        list.add_level(ListLevel::new(1, NumberFormat::Decimal));

        let bytes = list.to_bytes();
        assert_eq!(bytes.len(), 28); // Fixed LSTF size

        // Check list ID
        assert_eq!(
            u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]),
            0x12345678
        );
    }

    #[test]
    fn test_list_structure_simple_list_flag() {
        let mut list_single = ListStructure::new(1);
        list_single.add_level(ListLevel::new(1, NumberFormat::Decimal));

        let bytes_single = list_single.to_bytes();
        // Offset 26: flags byte (4 + 4 + 18 = 26), bit 0 = fSimpleList
        assert_eq!(bytes_single[26] & 0x01, 1);

        let mut list_multi = ListStructure::new(2);
        list_multi.add_level(ListLevel::new(1, NumberFormat::Decimal));
        list_multi.add_level(ListLevel::new(1, NumberFormat::Decimal));

        let bytes_multi = list_multi.to_bytes();
        assert_eq!(bytes_multi[26] & 0x01, 0);
    }

    #[test]
    fn test_list_structure_levels_to_bytes() {
        let mut list = ListStructure::new(1);
        list.add_level(ListLevel::new(1, NumberFormat::Decimal));
        list.add_level(ListLevel::new(1, NumberFormat::Bullet));

        let bytes = list.levels_to_bytes();
        // Should contain bytes from both levels
        assert!(!bytes.is_empty());
        assert!(bytes.len() >= 56); // At least 28 bytes per level
    }

    #[test]
    fn test_list_format_override_new() {
        let lfo = ListFormatOverride::new(100, 1);
        assert_eq!(lfo.list_id, 100);
        assert_eq!(lfo.lfo_id, 1);
    }

    #[test]
    fn test_list_format_override_to_bytes() {
        let lfo = ListFormatOverride::new(0x12345678, 5);
        let bytes = lfo.to_bytes();
        assert_eq!(bytes.len(), 16);

        // Check list ID
        assert_eq!(
            u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]),
            0x12345678
        );
    }

    #[test]
    fn test_numbering_writer_new() {
        let writer = NumberingWriter::new();
        assert!(writer.is_empty());
        assert_eq!(writer.list_count(), 0);
    }

    #[test]
    fn test_numbering_writer_default() {
        let writer: NumberingWriter = Default::default();
        assert!(writer.is_empty());
    }

    #[test]
    fn test_numbering_writer_add_list() {
        let mut writer = NumberingWriter::new();
        let list = ListStructure::new(1);
        writer.add_list(list);

        assert_eq!(writer.list_count(), 1);
        assert!(!writer.is_empty());
    }

    #[test]
    fn test_numbering_writer_add_override() {
        let mut writer = NumberingWriter::new();
        let lfo = ListFormatOverride::new(100, 1);
        writer.add_override(lfo);

        assert_eq!(writer.list_overrides.len(), 1);
    }

    #[test]
    fn test_build_plflst_empty() {
        let writer = NumberingWriter::new();
        let (header, lvl_data) = writer.build_plflst();

        // Should have just count (0)
        assert_eq!(header.len(), 2);
        assert_eq!(u16::from_le_bytes([header[0], header[1]]), 0);
        assert!(lvl_data.is_empty());
    }

    #[test]
    fn test_build_plflst_with_lists() {
        let mut writer = NumberingWriter::new();
        let mut list = ListStructure::new(1);
        list.add_level(ListLevel::new(1, NumberFormat::Decimal));
        writer.add_list(list);

        let (header, lvl_data) = writer.build_plflst();

        // Header: 2 bytes count + 28 bytes LSTF
        assert_eq!(header.len(), 30);
        assert!(!lvl_data.is_empty());
    }

    #[test]
    fn test_build_plflfo_empty() {
        let writer = NumberingWriter::new();
        let bytes = writer.build_plflfo();

        // Just count (0)
        assert_eq!(bytes.len(), 4);
        assert_eq!(
            u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]),
            0
        );
    }

    #[test]
    fn test_build_plflfo_with_overrides() {
        let mut writer = NumberingWriter::new();
        writer.add_override(ListFormatOverride::new(100, 1));
        writer.add_override(ListFormatOverride::new(200, 2));

        let bytes = writer.build_plflfo();

        // 4 bytes count + 2 * 16 bytes LFO
        assert_eq!(bytes.len(), 36);
        assert_eq!(
            u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]),
            2
        );
    }

    #[test]
    fn test_list_level_clone() {
        let level = ListLevel::new(1, NumberFormat::Decimal);
        let cloned = level.clone();
        assert_eq!(level.start_at, cloned.start_at);
        assert_eq!(level.number_format, cloned.number_format);
    }

    #[test]
    fn test_list_structure_clone() {
        let mut list = ListStructure::new(42);
        list.add_level(ListLevel::new(1, NumberFormat::Decimal));
        let cloned = list.clone();
        assert_eq!(list.list_id, cloned.list_id);
        assert_eq!(list.levels.len(), cloned.levels.len());
    }

    #[test]
    fn test_list_format_override_clone() {
        let lfo = ListFormatOverride::new(100, 1);
        let cloned = lfo.clone();
        assert_eq!(lfo.list_id, cloned.list_id);
        assert_eq!(lfo.lfo_id, cloned.lfo_id);
    }

    #[test]
    fn test_list_level_debug() {
        let level = ListLevel::new(1, NumberFormat::Decimal);
        let debug_str = format!("{:?}", level);
        assert!(debug_str.contains("ListLevel"));
    }

    #[test]
    fn test_list_structure_debug() {
        let list = ListStructure::new(1);
        let debug_str = format!("{:?}", list);
        assert!(debug_str.contains("ListStructure"));
    }

    #[test]
    fn test_numbering_writer_debug() {
        let writer = NumberingWriter::new();
        let debug_str = format!("{:?}", writer);
        assert!(debug_str.contains("NumberingWriter"));
    }

    #[test]
    fn test_all_number_formats_to_bytes() {
        let formats = vec![
            NumberFormat::Decimal,
            NumberFormat::UpperRoman,
            NumberFormat::LowerRoman,
            NumberFormat::UpperLetter,
            NumberFormat::LowerLetter,
            NumberFormat::Ordinal,
            NumberFormat::Bullet,
        ];

        for format in formats {
            let level = ListLevel::new(1, format);
            let bytes = level.to_bytes();
            assert!(!bytes.is_empty(), "Failed for format {:?}", format);
            assert_eq!(bytes[4], format as u8);
        }
    }

    #[test]
    fn test_list_level_custom_indent() {
        let mut level = ListLevel::new(1, NumberFormat::Decimal);
        level.indent_left = 1440; // 1 inch
        level.indent_hanging = -720; // -0.5 inch

        assert_eq!(level.indent_left, 1440);
        assert_eq!(level.indent_hanging, -720);
    }

    #[test]
    fn test_multiple_lists() {
        let mut writer = NumberingWriter::new();

        let mut list1 = ListStructure::new(1);
        list1.add_level(ListLevel::new(1, NumberFormat::Decimal));

        let mut list2 = ListStructure::new(2);
        list2.add_level(ListLevel::new(1, NumberFormat::Bullet));

        writer.add_list(list1);
        writer.add_list(list2);

        assert_eq!(writer.list_count(), 2);

        let (header, _) = writer.build_plflst();
        assert_eq!(header.len(), 2 + 2 * 28); // count + 2 LSTFs
    }
}
