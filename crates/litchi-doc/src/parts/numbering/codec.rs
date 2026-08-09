//! Binary codecs for DOC list tables and their FIB routing.

use super::model::{
    AutomaticNumberingField, HtmlCompatibilityFlags, ListFollowCharacter, ListFormatOverride,
    ListFormatOverrideMetadata, ListLevel, ListLevelMetadata, ListLevelOverride,
    ListLevelOverrideMetadata, ListStructure, ListStructureMetadata, ListStyleIndex, ListTables,
    ListTablesMetadata, NumberFormat, ParagraphListBinding,
};
use super::validation;
use crate::package::{Error as PackageError, Result};
use crate::parts::fib::FileInformationBlock;
use litchi_core::binary;

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

        // Parse PlfLst (List Table) - FibRgFcLcb97 index 73.
        if let Some((offset, length)) = fib.get_table_pointer(73)
            && length > 0
            && (offset as usize) < table_stream.len()
        {
            let offset = offset as usize;
            let header_end = offset.checked_add(length as usize).ok_or_else(|| {
                PackageError::InvalidFormat("PlfLst table range overflows".to_string())
            })?;
            if header_end > table_stream.len() {
                return Err(PackageError::InvalidFormat(
                    "PlfLst header extends beyond the table stream".to_string(),
                ));
            }
            let level_end = fib
                .get_table_pointer(74)
                .map(|(lfo_offset, _)| lfo_offset as usize)
                .filter(|&lfo_offset| lfo_offset >= header_end)
                .unwrap_or(table_stream.len());
            if level_end > table_stream.len() {
                return Err(PackageError::InvalidFormat(
                    "PlfLst level range extends beyond the table stream".to_string(),
                ));
            }

            list_structures = Self::parse_plflst(
                &table_stream[offset..header_end],
                &table_stream[header_end..level_end],
            )?;
        }

        // Parse PlfLfo (List Format Override Table) - FibRgFcLcb97 index 74.
        if let Some((offset, length)) = fib.get_table_pointer(74)
            && length > 0
            && (offset as usize) < table_stream.len()
        {
            let plf_data = &table_stream[offset as usize..];
            let plf_len = length.min((table_stream.len() - offset as usize) as u32) as usize;

            list_overrides = Self::parse_plflfo(&plf_data[..plf_len])?;
        }

        let metadata = Self::parse_metadata(fib, table_stream, &list_structures, &list_overrides)?;

        Ok(Self {
            list_structures,
            list_overrides,
            metadata,
        })
    }

    fn parse_metadata(
        fib: &FileInformationBlock,
        table_stream: &[u8],
        structures: &[ListStructure],
        overrides: &[ListFormatOverride],
    ) -> Result<ListTablesMetadata> {
        let mut metadata = ListTablesMetadata::default();
        if let Some((offset, length)) = fib.get_table_pointer(73).filter(|(_, length)| *length > 0)
        {
            let start = usize::try_from(offset).map_err(|_| {
                PackageError::InvalidFormat("PlfLst offset exceeds usize".to_string())
            })?;
            let header_end = start.checked_add(length as usize).ok_or_else(|| {
                PackageError::InvalidFormat("PlfLst metadata range overflows".to_string())
            })?;
            let level_end = fib
                .get_table_pointer(74)
                .map(|(offset, _)| offset as usize)
                .filter(|&offset| offset >= header_end)
                .unwrap_or(table_stream.len());
            let header = table_stream.get(start..header_end).ok_or_else(|| {
                PackageError::InvalidFormat("PlfLst metadata header is truncated".to_string())
            })?;
            let levels = table_stream.get(header_end..level_end).ok_or_else(|| {
                PackageError::InvalidFormat("PlfLst metadata levels are truncated".to_string())
            })?;
            let mut level_offset = 0usize;
            for (index, structure) in structures.iter().enumerate() {
                let lstf = &header[2 + index * 28..2 + (index + 1) * 28];
                let mut definition = ListStructureMetadata::default();
                for style in 0..9 {
                    let raw = u16::from_le_bytes([lstf[8 + style * 2], lstf[9 + style * 2]]);
                    definition.style_links[style] = match raw {
                        0x0FFF | 0xFFFF => None,
                        value => Some(ListStyleIndex::new(value).map_err(|invalid| {
                            PackageError::InvalidFormat(format!(
                                "LSTF has invalid linked style index {invalid:#06x}"
                            ))
                        })?),
                    };
                }
                let flags = lstf[26];
                definition.automatic_numbering = flags & 0x04 != 0;
                definition.hybrid = flags & 0x10 != 0;
                definition.ignored_flags = flags & 0xEA;
                definition.html_compatibility = HtmlCompatibilityFlags::from_raw(lstf[27]);
                metadata.definitions.push(definition);

                let mut level_metadata = Vec::with_capacity(structure.levels.len());
                for level in &structure.levels {
                    let (parsed, size) = Self::parse_level_metadata(
                        levels.get(level_offset..).ok_or_else(|| {
                            PackageError::InvalidFormat(
                                "LVL metadata offset is invalid".to_string(),
                            )
                        })?,
                        level.level,
                        level.number_format,
                    )?;
                    level_offset = level_offset.checked_add(size).ok_or_else(|| {
                        PackageError::InvalidFormat("LVL metadata size overflows".to_string())
                    })?;
                    level_metadata.push(parsed);
                }
                metadata.levels.push(level_metadata);
            }
        }

        if let Some((offset, length)) = fib.get_table_pointer(74).filter(|(_, length)| *length > 0)
        {
            let start = offset as usize;
            let data = table_stream
                .get(start..start.saturating_add(length as usize))
                .ok_or_else(|| {
                    PackageError::InvalidFormat("PlfLfo metadata is truncated".to_string())
                })?;
            let mut data_offset = 4usize
                .checked_add(overrides.len().checked_mul(16).ok_or_else(|| {
                    PackageError::InvalidFormat("PlfLfo metadata count overflows".to_string())
                })?)
                .ok_or_else(|| {
                    PackageError::InvalidFormat("PlfLfo metadata overflows".to_string())
                })?;
            for (index, lfo) in overrides.iter().enumerate() {
                let raw = &data[4 + index * 16..4 + (index + 1) * 16];
                let first_cp = binary::read_u32_le(data, data_offset).map_err(|e| {
                    PackageError::InvalidFormat(format!("Failed to read LFOData CP: {e}"))
                })?;
                data_offset += 4;
                let mut parsed = ListFormatOverrideMetadata {
                    unused1: u32::from_le_bytes(raw[4..8].try_into().expect("LFO unused1")),
                    unused2: u32::from_le_bytes(raw[8..12].try_into().expect("LFO unused2")),
                    field: AutomaticNumberingField::try_from(raw[13]).map_err(|invalid| {
                        PackageError::InvalidFormat(format!(
                            "LFO has invalid automatic-number field {invalid:#04x}"
                        ))
                    })?,
                    html_compatibility: HtmlCompatibilityFlags::from_raw(raw[14]),
                    unused3: raw[15],
                    first_paragraph_cp: (first_cp != u32::MAX).then_some(first_cp),
                    levels: Vec::with_capacity(lfo.level_overrides.len()),
                };
                for level_override in &lfo.level_overrides {
                    let flags = binary::read_u32_le(data, data_offset + 4).map_err(|e| {
                        PackageError::InvalidFormat(format!("Failed to read LFOLVL flags: {e}"))
                    })?;
                    data_offset += 8;
                    let formatting = if let Some(level) = level_override.format.as_ref() {
                        let (metadata, size) = Self::parse_level_metadata(
                            &data[data_offset..],
                            level.level,
                            level.number_format,
                        )?;
                        data_offset += size;
                        Some(metadata)
                    } else {
                        None
                    };
                    parsed.levels.push(ListLevelOverrideMetadata {
                        unused_start_at: if flags & 0x10 == 0 {
                            binary::read_u32_le(data, data_offset - 8).map_err(|e| {
                                PackageError::InvalidFormat(format!(
                                    "Failed to read LFOLVL ignored start: {e}"
                                ))
                            })?
                        } else {
                            0
                        },
                        html_compatibility: HtmlCompatibilityFlags::from_raw(
                            ((flags >> 6) & 0xFF) as u8,
                        ),
                        ignored_flags: flags & 0xFFFF_C000,
                        formatting,
                    });
                }
                metadata.overrides.push(parsed);
            }
        }
        Ok(metadata)
    }

    fn parse_level_metadata(
        data: &[u8],
        level: u8,
        number_format: NumberFormat,
    ) -> Result<(ListLevelMetadata, usize)> {
        if data.len() < 30 {
            return Err(PackageError::InvalidFormat(
                "LVL metadata is truncated".to_string(),
            ));
        }
        let flags = data[5];
        let cb_chpx = usize::from(data[24]);
        let cb_papx = usize::from(data[25]);
        let text_offset = 28usize
            .checked_add(cb_papx)
            .and_then(|value| value.checked_add(cb_chpx))
            .ok_or_else(|| {
                PackageError::InvalidFormat("LVL metadata size overflows".to_string())
            })?;
        let text_len = usize::from(binary::read_u16_le(data, text_offset).map_err(|e| {
            PackageError::InvalidFormat(format!("Failed to read LVL metadata XST: {e}"))
        })?);
        let total = text_offset
            .checked_add(2)
            .and_then(|value| value.checked_add(text_len.checked_mul(2)?))
            .ok_or_else(|| PackageError::InvalidFormat("LVL metadata XST overflows".to_string()))?;
        if total > data.len() {
            return Err(PackageError::InvalidFormat(
                "LVL metadata XST is truncated".to_string(),
            ));
        }
        let placeholders: [u8; 9] = data[6..15].try_into().expect("LVLF placeholders");
        for position in placeholders.into_iter().filter(|position| *position != 0) {
            if usize::from(position) > text_len {
                return Err(PackageError::InvalidFormat(format!(
                    "LVLF placeholder position {position} exceeds XST length {text_len}"
                )));
            }
            let offset = text_offset + 2 + (usize::from(position) - 1) * 2;
            let placeholder = u16::from_le_bytes([data[offset], data[offset + 1]]);
            if placeholder > u16::from(level) {
                return Err(PackageError::InvalidFormat(format!(
                    "LVL placeholder level {placeholder} exceeds level {level}"
                )));
            }
        }
        if number_format == NumberFormat::Bullet
            && (text_len != 1 || placeholders.iter().any(|position| *position != 0))
        {
            return Err(PackageError::InvalidFormat(
                "bullet LVL must contain one character and no placeholders".to_string(),
            ));
        }
        let no_restart = flags & 0x08 != 0;
        let restart = data[26];
        if no_restart && restart > level {
            return Err(PackageError::InvalidFormat(format!(
                "LVLF restart limit {restart} exceeds level {level}"
            )));
        }
        Ok((
            ListLevelMetadata {
                legal_numbering: flags & 0x04 != 0,
                no_restart,
                saved_indent: (flags & 0x10 != 0).then(|| {
                    i32::from_le_bytes(data[16..20].try_into().expect("LVLF saved indent"))
                }),
                ignored_saved_indent: if flags & 0x10 == 0 {
                    i32::from_le_bytes(data[16..20].try_into().expect("LVLF ignored indent"))
                } else {
                    0
                },
                converted: flags & 0x20 != 0,
                tentative: flags & 0x80 != 0,
                ignored_flags: flags & 0x40,
                placeholder_positions: placeholders,
                follow_character: ListFollowCharacter::try_from(data[15]).map_err(|invalid| {
                    PackageError::InvalidFormat(format!(
                        "LVLF has invalid follow character {invalid}"
                    ))
                })?,
                unused_value: u32::from_le_bytes(data[20..24].try_into().expect("LVLF unused2")),
                restart_limit: no_restart.then_some(restart),
                ignored_restart_limit: if no_restart { 0 } else { restart },
                html_compatibility: HtmlCompatibilityFlags::from_raw(data[27]),
                paragraph_properties: data[28..28 + cb_papx].to_vec(),
                number_properties: data[28 + cb_papx..text_offset].to_vec(),
            },
            total,
        ))
    }

    /// Parse `PlfLst` (List Table)
    pub(super) fn parse_plflst(
        header_data: &[u8],
        level_data: &[u8],
    ) -> Result<Vec<ListStructure>> {
        if header_data.len() < 2 {
            return Err(PackageError::InvalidFormat(
                "PlfLst is too short".to_string(),
            ));
        }

        let count = binary::read_u16_le(header_data, 0)
            .map_err(|e| PackageError::InvalidFormat(format!("Failed to read count: {e}")))?
            as usize;
        validation::count(count, "PlfLst structure count")?;
        let expected_header_len = 2usize
            .checked_add(count.checked_mul(28).ok_or_else(|| {
                PackageError::InvalidFormat("PlfLst structure count overflows".to_string())
            })?)
            .ok_or_else(|| PackageError::InvalidFormat("PlfLst size overflows".to_string()))?;
        if header_data.len() != expected_header_len {
            return Err(PackageError::InvalidFormat(format!(
                "PlfLst header length is {}, expected {expected_header_len}",
                header_data.len()
            )));
        }

        let mut structures = Vec::with_capacity(count);
        for index in 0..count {
            let offset = 2 + index * 28;
            structures.push(ListStructure::from_bytes(
                &header_data[offset..offset + 28],
            )?);
        }

        let mut level_offset = 0usize;
        for structure in &mut structures {
            let level_count = if structure.is_simple { 1 } else { 9 };
            structure.levels.reserve(level_count);
            for level in 0..level_count {
                let (parsed, size) = ListLevel::parse_with_size(
                    level_data.get(level_offset..).ok_or_else(|| {
                        PackageError::InvalidFormat("PlfLst LVL offset is invalid".to_string())
                    })?,
                    level as u8,
                )?;
                level_offset = level_offset.checked_add(size).ok_or_else(|| {
                    PackageError::InvalidFormat("PlfLst LVL size overflows".to_string())
                })?;
                structure.levels.push(parsed);
            }
        }

        Ok(structures)
    }

    /// Parse `PlfLfo` (List Format Override Table)
    pub(super) fn parse_plflfo(data: &[u8]) -> Result<Vec<ListFormatOverride>> {
        if data.len() < 4 {
            return Ok(Vec::new());
        }

        let count = binary::read_u32_le(data, 0)
            .map_err(|e| PackageError::InvalidFormat(format!("Failed to read count: {e}")))?
            as usize;
        validation::count(count, "PlfLfo count")?;
        let mut overrides = Vec::with_capacity(count);
        let lfo_bytes = count
            .checked_mul(16)
            .ok_or_else(|| PackageError::InvalidFormat("PlfLfo count overflows".to_string()))?;
        let lfo_data_start = 4usize
            .checked_add(lfo_bytes)
            .ok_or_else(|| PackageError::InvalidFormat("PlfLfo size overflows".to_string()))?;
        if lfo_data_start > data.len() {
            return Err(PackageError::InvalidFormat(
                "PlfLfo LFO array is truncated".to_string(),
            ));
        }
        let mut offset = 4;

        for index in 0..count {
            overrides.push(ListFormatOverride::from_bytes_with_id(
                &data[offset..offset + 16],
                u32::try_from(index + 1).map_err(|_| {
                    PackageError::InvalidFormat("PlfLfo index exceeds u32".to_string())
                })?,
            )?);
            offset += 16;
        }

        /// Mask for the zero-based `iLvl` field of an `LFOLVL`.
        const LFOLVL_ILVL_MASK: u32 = 0x0F;
        /// `fStartAt` flag of an `LFOLVL`.
        const LFOLVL_F_START_AT: u32 = 0x10;
        /// `fFormatting` flag of an `LFOLVL`.
        const LFOLVL_F_FORMATTING: u32 = 0x20;
        /// Maximum permitted `iStartAt` override value ([MS-DOC] 2.9.133).
        const LFOLVL_MAX_START_AT: u32 = 0x7FFF;

        let mut data_offset = lfo_data_start;
        for lfo in &mut overrides {
            data_offset = data_offset.checked_add(4).ok_or_else(|| {
                PackageError::InvalidFormat("PlfLfo LFOData size overflows".to_string())
            })?;
            if data_offset > data.len() {
                return Err(PackageError::InvalidFormat(
                    "PlfLfo LFOData array is truncated".to_string(),
                ));
            }
            for _ in 0..lfo.override_count {
                let base_end = data_offset.checked_add(8).ok_or_else(|| {
                    PackageError::InvalidFormat("LFOLVL size overflows".to_string())
                })?;
                if base_end > data.len() {
                    return Err(PackageError::InvalidFormat(
                        "LFOLVL is truncated".to_string(),
                    ));
                }
                let start_at = binary::read_u32_le(data, data_offset).map_err(|e| {
                    PackageError::InvalidFormat(format!("Failed to read LFOLVL iStartAt: {e}"))
                })?;
                let flags = binary::read_u32_le(data, data_offset + 4).map_err(|e| {
                    PackageError::InvalidFormat(format!("Failed to read LFOLVL flags: {e}"))
                })?;
                data_offset = base_end;
                let level = (flags & LFOLVL_ILVL_MASK) as u8;
                if level > 8 {
                    return Err(PackageError::InvalidFormat(format!(
                        "LFOLVL has invalid iLvl {level}"
                    )));
                }
                let overrides_start_at = flags & LFOLVL_F_START_AT != 0;
                let overrides_formatting = flags & LFOLVL_F_FORMATTING != 0;
                let mut level_override = ListLevelOverride {
                    level,
                    start_at: None,
                    format: None,
                };
                if overrides_formatting {
                    let (parsed, size) = ListLevel::parse_with_size(&data[data_offset..], level)?;
                    data_offset = data_offset.checked_add(size).ok_or_else(|| {
                        PackageError::InvalidFormat("LFOLVL formatting size overflows".to_string())
                    })?;
                    level_override.format = Some(parsed);
                } else if overrides_start_at {
                    // iStartAt is only meaningful when fFormatting is clear.
                    if start_at > LFOLVL_MAX_START_AT {
                        return Err(PackageError::InvalidFormat(format!(
                            "LFOLVL start value {start_at} exceeds 32767"
                        )));
                    }
                    level_override.start_at = Some(start_at);
                }
                lfo.level_overrides.push(level_override);
            }
        }
        if data_offset != data.len() {
            return Err(PackageError::InvalidFormat(format!(
                "PlfLfo has {} trailing bytes",
                data.len() - data_offset
            )));
        }

        Ok(overrides)
    }

    /// Get all list structures
    #[must_use]
    pub fn structures(&self) -> &[ListStructure] {
        &self.list_structures
    }

    /// Get all list format overrides
    #[must_use]
    pub fn overrides(&self) -> &[ListFormatOverride] {
        &self.list_overrides
    }

    /// Lossless typed metadata for the list tables.
    #[must_use]
    pub fn metadata(&self) -> &ListTablesMetadata {
        &self.metadata
    }

    /// Find a list structure by ID
    #[must_use]
    pub fn find_structure(&self, list_id: u32) -> Option<&ListStructure> {
        self.list_structures
            .iter()
            .find(|lst| lst.list_id == list_id)
    }

    /// Find a list override by LFO ID
    #[must_use]
    pub fn find_override(&self, lfo_id: u32) -> Option<&ListFormatOverride> {
        self.list_overrides.iter().find(|lfo| lfo.lfo_id == lfo_id)
    }

    /// Get the list structure for a given LFO ID
    #[must_use]
    pub fn get_list_for_lfo(&self, lfo_id: u32) -> Option<&ListStructure> {
        self.find_override(lfo_id)
            .and_then(|lfo| self.find_structure(lfo.list_id))
    }

    /// Resolve the signed paragraph list reference and level without cloning.
    ///
    /// `sprmPIlfo` is one-based. Negative values select the same absolute LFO
    /// index while requesting that paragraph indents be preserved. Level 12 is
    /// the specification's "skip numbering" sentinel and does not bind a list.
    #[must_use]
    pub fn bind_paragraph(&self, signed_lfo: i16, level: u8) -> Option<ParagraphListBinding<'_>> {
        if signed_lfo == 0 || signed_lfo == i16::MIN || level > 8 {
            return None;
        }
        let lfo_id = u32::from(signed_lfo.unsigned_abs());
        let format_override = self.find_override(lfo_id)?;
        let definition = self.find_structure(format_override.list_id)?;
        let base_level = definition.level(level)?;
        Some(ParagraphListBinding {
            lfo_id,
            level,
            preserve_indents: signed_lfo.is_negative(),
            definition,
            format_override,
            base_level,
            level_override: format_override.level_override(level),
        })
    }

    /// Resolve the effective level formatting for an LFO ID and zero-based
    /// level, applying any `LFOLVL` start-at or formatting overrides.
    #[must_use]
    pub fn resolve_level(&self, lfo_id: u32, level: u8) -> Option<ListLevel> {
        let signed_lfo = i16::try_from(lfo_id).ok()?;
        let binding = self.bind_paragraph(signed_lfo, level)?;
        let mut resolved = binding.effective_level().clone();
        resolved.start_at = binding.effective_start_at();
        Some(resolved)
    }
}
