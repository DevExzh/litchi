//! List numbering writer for DOC files
//!
//! Generates list structures (LST, LVL) and format overrides (LFO, LFOLVL).

use super::core::WriteError;
#[cfg(test)]
use crate::parts::numbering::ListLevelOverrideMetadata;
pub use crate::parts::numbering::NumberFormat;
use crate::parts::numbering::{
    ListFormatOverrideMetadata, ListLevelMetadata, ListStructureMetadata,
};
use crate::parts::{list_names::ListNamesTable, list_templates::ListTemplateTable};
use std::io::Write;

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
    #[must_use]
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
    pub fn to_bytes(&self) -> Result<Vec<u8>, WriteError> {
        self.to_bytes_with_metadata(None)
    }

    fn to_bytes_with_metadata(
        &self,
        metadata: Option<&ListLevelMetadata>,
    ) -> Result<Vec<u8>, WriteError> {
        if matches!(
            self.number_format,
            NumberFormat::Hex
                | NumberFormat::Chicago
                | NumberFormat::DecimalHalfWidth
                | NumberFormat::DecimalFullWidth2
        ) {
            return Err(WriteError::InvalidData(format!(
                "MSONFC {:#04x} is forbidden in a DOC list level",
                self.number_format as u8
            )));
        }
        if self.number_format != NumberFormat::Bullet
            && self.number_format != NumberFormat::None
            && self.start_at > 0x7FFF
        {
            return Err(WriteError::InvalidData(format!(
                "DOC list start value {} exceeds 32767",
                self.start_at
            )));
        }
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
                let mut encoded = [0u16; 2];
                xst_chars.extend_from_slice(src[i].encode_utf16(&mut encoded));
                i += 1;
            }
        }

        // For bullet format, override with bullet character (no level placeholder)
        if self.number_format == NumberFormat::Bullet {
            xst_chars.clear();
            xst_chars.push(0x2022); // •
            rgbxch_nums = [0u8; 9]; // no level placeholders for bullets
        }

        if let Some(metadata) = metadata {
            rgbxch_nums = metadata.placeholder_positions;
        }

        let paragraph_properties =
            metadata.map_or(&[][..], |value| value.paragraph_properties.as_slice());
        let number_properties =
            metadata.map_or(&[][..], |value| value.number_properties.as_slice());
        let cb_grpprl_papx = u8::try_from(paragraph_properties.len()).map_err(|_| {
            WriteError::InvalidData("LVL paragraph properties exceed 255 bytes".to_string())
        })?;
        let cb_grpprl_chpx = u8::try_from(number_properties.len()).map_err(|_| {
            WriteError::InvalidData("LVL number properties exceed 255 bytes".to_string())
        })?;

        let mut buf = Vec::with_capacity(
            28 + paragraph_properties.len() + number_properties.len() + 2 + xst_chars.len() * 2,
        );

        // === LVLF (exactly 28 bytes) per MS-DOC 2.9.150 ===
        // Offset 0: iStartAt (4 bytes)
        buf.write_all(&self.start_at.to_le_bytes()).unwrap();
        // Offset 4: nfc (1 byte) — number format code
        buf.push(self.number_format as u8);
        // Offset 5: jc:2 + typed flags:6. Writer-facing levels remain left aligned.
        let flags = metadata.map_or(0, |value| {
            (value.ignored_flags & 0x40)
                | u8::from(value.legal_numbering) << 2
                | u8::from(value.no_restart) << 3
                | u8::from(value.saved_indent.is_some()) << 4
                | u8::from(value.converted) << 5
                | u8::from(value.tentative) << 7
        });
        buf.push(flags);
        // Offset 6: rgbxchNums[9] (9 bytes) — placeholder positions
        buf.write_all(&rgbxch_nums).unwrap();
        // Offset 15: ixchFollow (1 byte) — 0=tab, 1=space, 2=nothing
        buf.push(metadata.map_or(0, |value| value.follow_character as u8));
        // Offset 16: dxaIndentSav (4 bytes, i32 LE)
        let saved_indent = metadata.map_or(0, |value| {
            value.saved_indent.unwrap_or(value.ignored_saved_indent)
        });
        buf.write_all(&saved_indent.to_le_bytes()).unwrap();
        // Offset 20: reserved2 (4 bytes)
        buf.write_all(&metadata.map_or(0, |value| value.unused_value).to_le_bytes())
            .unwrap();
        // Offset 24: cbGrpprlChpx (1 byte)
        buf.push(cb_grpprl_chpx);
        // Offset 25: cbGrpprlPapx (1 byte)
        buf.push(cb_grpprl_papx);
        // Offset 26: ilvlRestartLim.
        buf.push(metadata.map_or(0, |value| {
            value.restart_limit.unwrap_or(value.ignored_restart_limit)
        }));
        // Offset 27: grfhic.
        buf.push(metadata.map_or(0, |value| value.html_compatibility.raw()));

        buf.extend_from_slice(paragraph_properties);
        buf.extend_from_slice(number_properties);

        // xst: cch (u16 LE) + cch UTF-16LE characters
        buf.write_all(&(xst_chars.len() as u16).to_le_bytes())
            .unwrap();
        for &ch in &xst_chars {
            buf.write_all(&ch.to_le_bytes()).unwrap();
        }

        Ok(buf)
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

/// Lossless metadata paired with one writer list definition.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct ListDefinitionMetadata {
    pub definition: ListStructureMetadata,
    pub levels: Vec<ListLevelMetadata>,
}

impl ListStructure {
    /// Create a new list structure
    #[must_use]
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
    /// [`Self::levels_to_bytes`] for that.
    #[must_use]
    pub fn to_bytes(&self) -> Vec<u8> {
        self.to_bytes_with_metadata(None)
    }

    fn to_bytes_with_metadata(&self, metadata: Option<&ListStructureMetadata>) -> Vec<u8> {
        let mut buf = Vec::with_capacity(28);

        // List ID (4 bytes)
        buf.write_all(&self.list_id.to_le_bytes()).unwrap();

        // Template ID (4 bytes)
        buf.write_all(&self.template_id.to_le_bytes()).unwrap();

        for index in 0..9 {
            let style = metadata
                .and_then(|value| value.style_links[index])
                .map_or(0x0FFF, crate::parts::numbering::ListStyleIndex::get);
            buf.write_all(&style.to_le_bytes()).unwrap();
        }

        // Flags (1 byte):
        //   bit 0: fSimpleList (1 = single-level, 0 = multi-level)
        //   bit 1: fRestartHdn
        //   bit 2: fAutoNum (unused)
        //   bits 3-7: reserved
        let f_simple = u8::from(self.levels.len() <= 1);
        let flags = metadata.map_or(f_simple, |value| {
            f_simple
                | (value.ignored_flags & 0xEA)
                | u8::from(value.automatic_numbering) << 2
                | u8::from(value.hybrid) << 4
        });
        buf.push(flags);

        // grfhic (1 byte) — reserved/compatibility, set to 0
        buf.push(metadata.map_or(0, |value| value.html_compatibility.raw()));

        buf
    }

    /// Serialize the LVL array for this list structure.
    ///
    /// Per MS-DOC spec, LVLs are appended after all LSTFs in the `PlfLst`
    /// and are NOT counted in `lcbPlfLst`.
    pub fn levels_to_bytes(&self) -> Result<Vec<u8>, WriteError> {
        self.levels_to_bytes_with_metadata(None)
    }

    fn levels_to_bytes_with_metadata(
        &self,
        metadata: Option<&[ListLevelMetadata]>,
    ) -> Result<Vec<u8>, WriteError> {
        let mut buf = Vec::new();
        let level_count = if self.levels.len() <= 1 { 1 } else { 9 };
        for level_index in 0..level_count {
            if let Some(level) = self.levels.get(level_index) {
                buf.extend_from_slice(
                    &level.to_bytes_with_metadata(
                        metadata.and_then(|values| values.get(level_index)),
                    )?,
                );
            } else {
                let mut level = ListLevel::new(1, NumberFormat::Decimal);
                level.number_text = format!("%{}.", level_index + 1);
                buf.extend_from_slice(
                    &level.to_bytes_with_metadata(
                        metadata.and_then(|values| values.get(level_index)),
                    )?,
                );
            }
        }
        Ok(buf)
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

/// Writer representation of one `LFOLVL` entry.
#[derive(Debug, Clone)]
pub struct ListLevelOverride {
    pub level: u8,
    pub start_at: Option<u32>,
    pub format: Option<ListLevel>,
}

/// Complete typed `LFOData` and `LFOLVL` payload for one override.
#[derive(Debug, Clone)]
pub struct ListFormatOverrideData {
    pub metadata: ListFormatOverrideMetadata,
    pub levels: Vec<ListLevelOverride>,
}

impl ListFormatOverride {
    /// Create a new list format override
    #[must_use]
    pub fn new(list_id: u32, lfo_id: u32) -> Self {
        Self { list_id, lfo_id }
    }

    /// Serialize to LFO structure (16 bytes)
    #[must_use]
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

    fn encode_with_data(
        &self,
        data: &ListFormatOverrideData,
    ) -> Result<(Vec<u8>, Vec<u8>), WriteError> {
        if data.levels.len() > 9 || data.metadata.levels.len() != data.levels.len() {
            return Err(WriteError::InvalidData(
                "LFOData must contain matching metadata for at most nine levels".to_string(),
            ));
        }
        let mut seen = [false; 9];
        let mut header = Vec::with_capacity(16);
        header.extend_from_slice(&self.list_id.to_le_bytes());
        header.extend_from_slice(&data.metadata.unused1.to_le_bytes());
        header.extend_from_slice(&data.metadata.unused2.to_le_bytes());
        header.push(data.levels.len() as u8);
        header.push(data.metadata.field as u8);
        header.push(data.metadata.html_compatibility.raw());
        header.push(data.metadata.unused3);

        let mut body = Vec::new();
        body.extend_from_slice(
            &data
                .metadata
                .first_paragraph_cp
                .unwrap_or(u32::MAX)
                .to_le_bytes(),
        );
        for (level, metadata) in data.levels.iter().zip(&data.metadata.levels) {
            let index = usize::from(level.level);
            if index >= seen.len() || seen[index] {
                return Err(WriteError::InvalidData(format!(
                    "LFO level {} is out of range or duplicated",
                    level.level
                )));
            }
            seen[index] = true;
            if level.start_at.is_some_and(|value| value > 0x7FFF) {
                return Err(WriteError::InvalidData(format!(
                    "LFO start value for level {} exceeds 32767",
                    level.level
                )));
            }
            if level.format.is_some() != metadata.formatting.is_some() {
                return Err(WriteError::InvalidData(format!(
                    "LFO level {} formatting and metadata disagree",
                    level.level
                )));
            }
            body.extend_from_slice(
                &level
                    .start_at
                    .unwrap_or(metadata.unused_start_at)
                    .to_le_bytes(),
            );
            let flags = u32::from(level.level)
                | u32::from(level.start_at.is_some()) << 4
                | u32::from(level.format.is_some()) << 5
                | u32::from(metadata.html_compatibility.raw()) << 6
                | (metadata.ignored_flags & 0xFFFF_C000);
            body.extend_from_slice(&flags.to_le_bytes());
            if let (Some(format), Some(format_metadata)) =
                (level.format.as_ref(), metadata.formatting.as_ref())
            {
                body.extend_from_slice(&format.to_bytes_with_metadata(Some(format_metadata))?);
            }
        }
        Ok((header, body))
    }
}

/// Numbering writer for list tables
#[derive(Debug)]
pub struct NumberingWriter {
    list_structures: Vec<ListStructure>,
    list_metadata: Vec<Option<ListDefinitionMetadata>>,
    list_overrides: Vec<ListFormatOverride>,
    override_headers: Vec<Option<Vec<u8>>>,
    override_data: Vec<Option<Vec<u8>>>,
    list_names: Option<ListNamesTable>,
    list_templates: Option<ListTemplateTable>,
}

impl NumberingWriter {
    /// Create a new numbering writer
    #[must_use]
    pub fn new() -> Self {
        Self {
            list_structures: Vec::new(),
            list_metadata: Vec::new(),
            list_overrides: Vec::new(),
            override_headers: Vec::new(),
            override_data: Vec::new(),
            list_names: None,
            list_templates: None,
        }
    }

    /// Add a list structure
    pub fn add_list(&mut self, list: ListStructure) {
        self.list_structures.push(list);
        self.list_metadata.push(None);
    }

    /// Add a list with lossless LSTF/LVLF metadata and property payloads.
    pub fn add_list_with_metadata(
        &mut self,
        list: ListStructure,
        metadata: ListDefinitionMetadata,
    ) {
        self.list_structures.push(list);
        self.list_metadata.push(Some(metadata));
    }

    /// Add a list format override
    pub fn add_override(&mut self, lfo: ListFormatOverride) {
        self.list_overrides.push(lfo);
        self.override_headers.push(None);
        self.override_data.push(None);
    }

    /// Add a complete LFO/LFOData entry, validating and encoding it transactionally.
    pub fn add_override_with_data(
        &mut self,
        lfo: ListFormatOverride,
        data: ListFormatOverrideData,
    ) -> Result<(), WriteError> {
        let (header, body) = lfo.encode_with_data(&data)?;
        self.list_overrides.push(lfo);
        self.override_headers.push(Some(header));
        self.override_data.push(Some(body));
        Ok(())
    }

    /// Set names parallel to `PlfLst.rgLstf` for `LISTNUM` fields.
    pub fn set_list_names(&mut self, table: ListNamesTable) {
        self.list_names = Some(table);
    }

    /// Set list-level template codes parallel to `PlfLst.rgLstf`.
    pub fn set_list_templates(&mut self, table: ListTemplateTable) {
        self.list_templates = Some(table);
    }

    pub fn build_sttb_list_names(&self) -> Result<Option<Vec<u8>>, WriteError> {
        self.list_names
            .as_ref()
            .map(ListNamesTable::to_bytes)
            .transpose()
            .map_err(|error| WriteError::InvalidData(error.to_string()))
    }

    pub fn build_sttb_rgtplc(&self) -> Result<Option<Vec<u8>>, WriteError> {
        self.list_templates
            .as_ref()
            .map(ListTemplateTable::to_bytes)
            .transpose()
            .map_err(|error| WriteError::InvalidData(error.to_string()))
    }

    /// Get number of list structures
    #[must_use]
    pub fn list_count(&self) -> usize {
        self.list_structures.len()
    }

    /// Generate `PlfLst` (List Table).
    ///
    /// Returns `(plflst_for_lcb, lvl_data)` where:
    /// - `plflst_for_lcb` = cLst (u16) + LSTF array (28 bytes each) — this is what
    ///   `lcbPlfLst` should cover.
    /// - `lvl_data` = LVL array for all lists — appended immediately after but
    ///   NOT counted in `lcbPlfLst` per MS-DOC spec / Apache POI.
    pub fn build_plflst(&self) -> Result<(Vec<u8>, Vec<u8>), WriteError> {
        let mut header_buf = Vec::new();
        let mut lvl_buf = Vec::new();

        // Count of lists (2 bytes)
        header_buf
            .write_all(&(self.list_structures.len() as u16).to_le_bytes())
            .unwrap();

        // Each LSTF (fixed 28 bytes)
        for (index, list) in self.list_structures.iter().enumerate() {
            header_buf.extend_from_slice(
                &list.to_bytes_with_metadata(
                    self.list_metadata
                        .get(index)
                        .and_then(Option::as_ref)
                        .map(|value| &value.definition),
                ),
            );
        }

        // LVL data for all lists
        for (index, list) in self.list_structures.iter().enumerate() {
            lvl_buf.extend_from_slice(
                &list.levels_to_bytes_with_metadata(
                    self.list_metadata
                        .get(index)
                        .and_then(Option::as_ref)
                        .map(|value| value.levels.as_slice()),
                )?,
            );
        }

        Ok((header_buf, lvl_buf))
    }

    /// Generate `PlfLfo` (List Format Override Table)
    pub fn build_plflfo(&self) -> Vec<u8> {
        let mut buf = Vec::new();

        // Count of overrides (4 bytes)
        buf.write_all(&(self.list_overrides.len() as u32).to_le_bytes())
            .unwrap();

        // Each override
        for (index, lfo) in self.list_overrides.iter().enumerate() {
            if let Some(header) = self.override_headers.get(index).and_then(Option::as_ref) {
                buf.extend_from_slice(header);
            } else {
                buf.extend_from_slice(&lfo.to_bytes());
            }
        }

        // Parallel LFOData array. With clfolvl=0 each entry contains only its main-story CP.
        for index in 0..self.list_overrides.len() {
            if let Some(data) = self.override_data.get(index).and_then(Option::as_ref) {
                buf.extend_from_slice(data);
            } else {
                buf.extend_from_slice(&u32::MAX.to_le_bytes());
            }
        }

        buf
    }

    /// Check if empty
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.list_structures.is_empty()
            && self.list_names.is_none()
            && self.list_templates.is_none()
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
        let bytes = level.to_bytes().unwrap();

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
        let bytes = level.to_bytes().unwrap();

        // Bullet format should have bullet character 0x2022
        assert_eq!(bytes[4], 23); // Bullet = 23
    }

    #[test]
    fn test_list_level_with_level_placeholder() {
        let mut level = ListLevel::new(1, NumberFormat::Decimal);
        level.number_text = "%1.".to_string();
        let bytes = level.to_bytes().unwrap();

        // Should generate valid LVL structure
        assert!(bytes.len() > 28);
    }

    #[test]
    fn test_list_level_multi_level_placeholder() {
        let mut level = ListLevel::new(1, NumberFormat::Decimal);
        level.number_text = "%1.%2.%3.".to_string();
        let bytes = level.to_bytes().unwrap();

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

        let bytes = list.levels_to_bytes().unwrap();
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
        let (header, lvl_data) = writer.build_plflst().unwrap();

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

        let (header, lvl_data) = writer.build_plflst().unwrap();

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

        // 4 bytes count + 2 * 16 bytes LFO + 2 * 4 bytes LFOData
        assert_eq!(bytes.len(), 44);
        assert_eq!(
            u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]),
            2
        );
    }

    #[test]
    fn writes_lossless_list_definition_metadata_and_property_payloads() {
        let mut list = ListStructure::new(0x1122_3344);
        list.add_level(ListLevel::new(3, NumberFormat::Decimal));
        let mut definition = ListStructureMetadata::default();
        definition.style_links[0] = Some(crate::parts::numbering::ListStyleIndex::new(7).unwrap());
        definition.automatic_numbering = true;
        definition.hybrid = true;
        definition.ignored_flags = 0x40;
        definition.html_compatibility =
            crate::parts::numbering::HtmlCompatibilityFlags::from_raw(0xA5);
        let level = ListLevelMetadata {
            legal_numbering: true,
            saved_indent: Some(-240),
            placeholder_positions: [1, 0, 0, 0, 0, 0, 0, 0, 0],
            follow_character: crate::parts::numbering::ListFollowCharacter::Space,
            unused_value: 0x5566_7788,
            html_compatibility: crate::parts::numbering::HtmlCompatibilityFlags::from_raw(0x5A),
            paragraph_properties: vec![1, 2],
            number_properties: vec![3, 4, 5],
            ..ListLevelMetadata::default()
        };
        let mut writer = NumberingWriter::new();
        writer.add_list_with_metadata(
            list,
            ListDefinitionMetadata {
                definition,
                levels: vec![level],
            },
        );
        let (header, levels) = writer.build_plflst().unwrap();
        assert_eq!(u16::from_le_bytes([header[10], header[11]]), 7);
        assert_eq!(header[28], 0x55);
        assert_eq!(header[29], 0xA5);
        assert_eq!(levels[5], 0x14);
        assert_eq!(&levels[16..20], &(-240i32).to_le_bytes());
        assert_eq!(&levels[20..24], &0x5566_7788u32.to_le_bytes());
        assert_eq!(&levels[28..33], &[1, 2, 3, 4, 5]);
    }

    #[test]
    fn writes_complete_lfo_data_and_formatting_override() {
        let format_metadata = ListLevelMetadata {
            paragraph_properties: vec![0xAA],
            number_properties: vec![0xBB],
            ..ListLevelMetadata::default()
        };
        let override_metadata = ListLevelOverrideMetadata {
            unused_start_at: 99,
            html_compatibility: crate::parts::numbering::HtmlCompatibilityFlags::from_raw(0x12),
            ignored_flags: 0x8000_0000,
            formatting: Some(format_metadata),
        };
        let data = ListFormatOverrideData {
            metadata: ListFormatOverrideMetadata {
                unused1: 1,
                unused2: 2,
                field: crate::parts::numbering::AutomaticNumberingField::AutoNumber,
                html_compatibility: crate::parts::numbering::HtmlCompatibilityFlags::from_raw(3),
                unused3: 4,
                first_paragraph_cp: Some(123),
                levels: vec![override_metadata],
            },
            levels: vec![ListLevelOverride {
                level: 2,
                start_at: Some(7),
                format: Some(ListLevel::new(7, NumberFormat::Decimal)),
            }],
        };
        let mut writer = NumberingWriter::new();
        writer
            .add_override_with_data(ListFormatOverride::new(42, 1), data)
            .unwrap();
        let bytes = writer.build_plflfo();
        assert_eq!(&bytes[8..12], &1u32.to_le_bytes());
        assert_eq!(&bytes[12..16], &2u32.to_le_bytes());
        assert_eq!(&bytes[16..20], &[1, 0xFE, 3, 4]);
        assert_eq!(&bytes[20..24], &123u32.to_le_bytes());
        assert_eq!(&bytes[24..28], &7u32.to_le_bytes());
        assert_eq!(
            u32::from_le_bytes(bytes[28..32].try_into().unwrap()),
            0x8000_04B2
        );
        assert_eq!(&bytes[60..62], &[0xAA, 0xBB]);
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
        for value in 0..=59 {
            let format = NumberFormat::try_from(value).unwrap();
            let level = ListLevel::new(1, format);
            if matches!(value, 8 | 9 | 15 | 19) {
                assert!(level.to_bytes().is_err());
            } else {
                let bytes = level.to_bytes().unwrap();
                assert!(!bytes.is_empty(), "Failed for format {:?}", format);
                assert_eq!(bytes[4], value);
            }
        }
        let none = ListLevel::new(1, NumberFormat::None).to_bytes().unwrap();
        assert_eq!(none[4], 0xFF);
        assert!(
            ListLevel::new(32_768, NumberFormat::Decimal)
                .to_bytes()
                .is_err()
        );
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

        let (header, _) = writer.build_plflst().unwrap();
        assert_eq!(header.len(), 2 + 2 * 28); // count + 2 LSTFs
    }
}
