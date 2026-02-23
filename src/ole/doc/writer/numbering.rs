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
