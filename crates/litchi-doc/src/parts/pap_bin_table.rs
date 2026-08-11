//! Paragraph-property bin-table reconstruction for legacy Word documents.
//!
//! `PlcfBtePapx` is a page index, not a paragraph PLCF. Each indexed 512-byte
//! `PapxFkp` supplies exact FC boundaries and a `GrpPrlAndIstd` for each run.
//! This module resolves that two-level structure and maps physical FC ranges
//! through the piece table into logical CP ranges.

use std::collections::HashSet;

use litchi_core::binary::{read_u16_le, read_u32_le};

use super::super::package::{Error as PackageError, Result};
use super::fkp::{PapxFkp, ParagraphHeight};
use super::pap::ParagraphProperties;
use super::piece_table::PieceTable;
use super::styles::StyleSheet;
use crate::sprm::parse_sprms;
use crate::sprm_operations::{SPRM_P_HUGE_PAPX, SPRM_P_TABLE_PROPS};

const FKP_PAGE_SIZE: usize = 512;
const MAX_DATA_INDIRECTION_DEPTH: usize = 64;
// Some older producer tables use operation 0x62, so accept it on input while
// emitting and exposing the current [MS-DOC] operation 0x6B.
const SPRM_P_TABLE_PROPS_LEGACY: u16 = 0x6462;

/// A contiguous logical paragraph-property run.
#[derive(Debug, Clone)]
pub struct ParagraphRun {
    /// Inclusive logical character position.
    pub start_cp: u32,
    /// Exclusive logical character position.
    pub end_cp: u32,
    /// Direct paragraph properties stored in the PAPX.
    pub properties: ParagraphProperties,
    /// Expanded direct PAPX followed by piece modifiers, retained for later
    /// table-style cascading once row and cell context is known.
    pub(crate) direct_grpprl: Vec<u8>,
    /// Paragraph style selected by the PAPX header before direct SPRMs run.
    pub(crate) initial_style_index: Option<u16>,
    /// Version-specific paragraph-height metadata from `BxPap`.
    pub paragraph_height: Option<ParagraphHeight>,
}

/// Parsed `PlcfBtePapx` and all reachable PAPX FKP runs.
#[derive(Debug)]
pub struct PapBinTable {
    runs: Vec<ParagraphRun>,
}

impl PapBinTable {
    #[cfg(test)]
    pub(crate) fn from_runs_for_test(runs: Vec<ParagraphRun>) -> Self {
        Self { runs }
    }

    /// Reconstruct a paragraph bin table from its page index.
    pub fn parse(
        plcf_bte_papx_data: &[u8],
        word_document: &[u8],
        data_stream: Option<&[u8]>,
        piece_table: &PieceTable,
        stylesheet: Option<&StyleSheet>,
    ) -> Result<Option<Self>> {
        // PlcBtePapx = (n + 1) FCs followed by n four-byte PnFkpPapx values.
        if plcf_bte_papx_data.len() < 12 || !(plcf_bte_papx_data.len() - 4).is_multiple_of(8) {
            return Ok(None);
        }
        let page_count = (plcf_bte_papx_data.len() - 4) / 8;
        if page_count == 0 {
            return Ok(None);
        }

        let mut runs = Vec::with_capacity(page_count.saturating_mul(10));
        let mut style_baseline_cache = None;
        let pn_array_offset = (page_count + 1) * 4;

        for index in 0..page_count {
            let pn_raw =
                read_u32_le(plcf_bte_papx_data, pn_array_offset + index * 4).map_err(|error| {
                    PackageError::Corrupted(format!("invalid PAP bin-table page: {error}"))
                })?;
            let page_number = pn_raw & 0x003f_ffff;
            if page_number == 0 {
                continue;
            }
            let page_offset = (page_number as usize)
                .checked_mul(FKP_PAGE_SIZE)
                .ok_or_else(|| {
                    PackageError::Corrupted("PAP FKP page offset overflowed".to_string())
                })?;
            let page_end = page_offset.checked_add(FKP_PAGE_SIZE).ok_or_else(|| {
                PackageError::Corrupted("PAP FKP page range overflowed".to_string())
            })?;
            let page = word_document.get(page_offset..page_end).ok_or_else(|| {
                PackageError::Corrupted("PAP FKP page extends beyond WordDocument".to_string())
            })?;
            let fkp = PapxFkp::parse(page, data_stream.unwrap_or_default())
                .ok_or_else(|| PackageError::Corrupted("PAP FKP page is malformed".to_string()))?;

            for entry_index in 0..fkp.count() {
                let entry = fkp.entry(entry_index).ok_or_else(|| {
                    PackageError::Corrupted("PAP FKP entry is malformed".to_string())
                })?;
                for (start_cp, end_cp) in piece_table.fc_range_to_cp_ranges(entry.fc, entry.end_fc)
                {
                    let piece_modifier = piece_table
                        .piece_for_cp(start_cp)
                        .map(super::piece_table::TextPiece::property_modifier)
                        .unwrap_or_default();
                    let (properties, direct_grpprl, initial_style_index) =
                        Self::parse_properties_with_direct_cached(
                            &entry.grpprl,
                            piece_modifier,
                            data_stream,
                            stylesheet,
                            &mut style_baseline_cache,
                        )?;
                    runs.push(ParagraphRun {
                        start_cp,
                        end_cp,
                        properties,
                        direct_grpprl,
                        initial_style_index,
                        paragraph_height: entry.paragraph_height,
                    });
                }
            }
        }

        runs.sort_unstable_by_key(|run| (run.start_cp, run.end_cp));
        let mut last_end_cp = 0;
        runs.retain_mut(|run| {
            if run.start_cp < last_end_cp {
                if run.end_cp <= last_end_cp {
                    return false;
                }
                run.start_cp = last_end_cp;
            }
            if run.start_cp >= run.end_cp {
                return false;
            }
            last_end_cp = run.end_cp;
            true
        });

        Ok(Some(Self { runs }))
    }

    #[cfg(test)]
    fn parse_properties(
        grpprl_and_istd: &[u8],
        piece_modifier: &[u8],
        data_stream: Option<&[u8]>,
        stylesheet: Option<&StyleSheet>,
    ) -> Result<ParagraphProperties> {
        Self::parse_properties_with_direct(grpprl_and_istd, piece_modifier, data_stream, stylesheet)
            .map(|(properties, _, _)| properties)
    }

    #[cfg(test)]
    fn parse_properties_with_direct(
        grpprl_and_istd: &[u8],
        piece_modifier: &[u8],
        data_stream: Option<&[u8]>,
        stylesheet: Option<&StyleSheet>,
    ) -> Result<(ParagraphProperties, Vec<u8>, Option<u16>)> {
        Self::parse_properties_with_direct_cached(
            grpprl_and_istd,
            piece_modifier,
            data_stream,
            stylesheet,
            &mut None,
        )
    }

    fn parse_properties_with_direct_cached(
        grpprl_and_istd: &[u8],
        piece_modifier: &[u8],
        data_stream: Option<&[u8]>,
        stylesheet: Option<&StyleSheet>,
        style_baseline_cache: &mut Option<(u16, ParagraphProperties)>,
    ) -> Result<(ParagraphProperties, Vec<u8>, Option<u16>)> {
        if grpprl_and_istd.is_empty() {
            let properties = stylesheet.map_or_else(
                || ParagraphProperties::from_sprm(piece_modifier),
                |styles| ParagraphProperties::from_sprm_with_stylesheet(piece_modifier, styles),
            )?;
            return Ok((properties, piece_modifier.to_vec(), None));
        }

        let (style_index, direct_sprms) = if grpprl_and_istd.len() >= 2 {
            (read_u16_le(grpprl_and_istd, 0).ok(), &grpprl_and_istd[2..])
        } else {
            (Some(u16::from(grpprl_and_istd[0])), &[][..])
        };

        let expanded;
        let sprms = if let Some(data) = data_stream {
            let mut visited = HashSet::new();
            expanded = Self::expand_data_indirections(direct_sprms, data, &mut visited, 0)?;
            expanded.as_deref().unwrap_or(direct_sprms)
        } else {
            if parse_sprms(direct_sprms)?.iter().any(|sprm| {
                matches!(
                    sprm.opcode,
                    SPRM_P_HUGE_PAPX | SPRM_P_TABLE_PROPS | SPRM_P_TABLE_PROPS_LEGACY
                )
            }) {
                return Err(PackageError::Corrupted(
                    "PAPX data indirection requires a Data Stream".to_string(),
                ));
            }
            direct_sprms
        };

        let direct_grpprl = [sprms, piece_modifier].concat();
        let mut properties = match (stylesheet, style_index) {
            (Some(styles), Some(index)) => {
                if style_baseline_cache
                    .as_ref()
                    .is_none_or(|(cached_index, _)| *cached_index != index)
                {
                    let baseline =
                        ParagraphProperties::resolve_style_baseline(Some(index), styles)?;
                    *style_baseline_cache = Some((index, baseline));
                }
                let baseline = &style_baseline_cache
                    .as_ref()
                    .expect("the adjacent style baseline was just populated")
                    .1;
                ParagraphProperties::cascade_styles_from_resolved_baseline(
                    baseline,
                    &direct_grpprl,
                    styles,
                )?
            },
            (Some(styles), None) => {
                ParagraphProperties::cascade_styles(None, &direct_grpprl, styles)?
            },
            (None, _) => ParagraphProperties::from_sprm(&direct_grpprl)?,
        };
        if properties.style_index.is_none() {
            properties.style_index = style_index;
        }
        Ok((properties, direct_grpprl, style_index))
    }

    /// Expand `sprmPHugePapx`/`sprmPTableProps` `PrcData` references.
    ///
    /// Malformed, cyclic, or excessively deep chains are reported as corruption.
    fn expand_data_indirections(
        grpprl: &[u8],
        data_stream: &[u8],
        visited: &mut HashSet<u32>,
        depth: usize,
    ) -> Result<Option<Vec<u8>>> {
        if depth >= MAX_DATA_INDIRECTION_DEPTH {
            return Err(PackageError::Corrupted(
                "PAPX data indirection exceeds the depth limit".to_string(),
            ));
        }

        for sprm in parse_sprms(grpprl)? {
            let is_huge = sprm.opcode == SPRM_P_HUGE_PAPX;
            let is_table_props =
                matches!(sprm.opcode, SPRM_P_TABLE_PROPS | SPRM_P_TABLE_PROPS_LEGACY);
            if !is_huge && !is_table_props {
                continue;
            }
            // A huge PAPX is valid only as the first Prl in its array.
            if is_huge && sprm.offset != 0 {
                continue;
            }

            let data_offset = sprm.operand_dword().ok_or_else(|| {
                PackageError::Corrupted("PAPX data indirection lacks an offset".to_string())
            })?;
            if !visited.insert(data_offset) {
                return Err(PackageError::Corrupted(
                    "PAPX data indirection contains a cycle".to_string(),
                ));
            }
            let offset = usize::try_from(data_offset).map_err(|_| {
                PackageError::Corrupted("PAPX data offset does not fit in memory".to_string())
            })?;
            let size = usize::from(read_u16_le(data_stream, offset).map_err(|error| {
                PackageError::Corrupted(format!("invalid PAPX PrcData length: {error}"))
            })?);
            if size < 10 {
                return Err(PackageError::Corrupted(
                    "PAPX PrcData is shorter than 10 bytes".to_string(),
                ));
            }
            let content_start = offset.checked_add(2).ok_or_else(|| {
                PackageError::Corrupted("PAPX PrcData start overflowed".to_string())
            })?;
            let content_end = content_start.checked_add(size).ok_or_else(|| {
                PackageError::Corrupted("PAPX PrcData range overflowed".to_string())
            })?;
            let referenced = data_stream.get(content_start..content_end).ok_or_else(|| {
                PackageError::Corrupted("PAPX PrcData extends beyond the Data Stream".to_string())
            })?;

            let nested =
                Self::expand_data_indirections(referenced, data_stream, visited, depth + 1)?;
            let resolved = nested.as_deref().unwrap_or(referenced);
            let mut combined = Vec::with_capacity(sprm.offset.saturating_add(resolved.len()));
            combined.extend_from_slice(&grpprl[..sprm.offset]);
            combined.extend_from_slice(resolved);
            return Ok(Some(combined));
        }

        Ok(None)
    }

    /// All reconstructed paragraph-property runs.
    #[inline]
    #[must_use]
    pub fn runs(&self) -> &[ParagraphRun] {
        &self.runs
    }

    /// Properties covering a logical character position.
    #[must_use]
    pub fn properties_at(&self, cp: u32) -> Option<&ParagraphProperties> {
        self.run_at(cp).map(|run| &run.properties)
    }

    /// Property run covering a logical character position.
    pub(crate) fn run_at(&self, cp: u32) -> Option<&ParagraphRun> {
        let index = self.runs.partition_point(|run| run.start_cp <= cp);
        let run = self.runs.get(index.checked_sub(1)?)?;
        (cp < run.end_cp).then_some(run)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::leniency::Leniency;

    fn style_record(
        invariant_id: u16,
        kind: u16,
        base: u16,
        next: u16,
        name: &str,
        property_sets: &[&[u8]],
    ) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&invariant_id.to_le_bytes());
        data.extend_from_slice(&(kind | (base << 4)).to_le_bytes());
        data.extend_from_slice(&((property_sets.len() as u16) | (next << 4)).to_le_bytes());
        data.extend_from_slice(&0_u16.to_le_bytes());
        data.extend_from_slice(&0_u16.to_le_bytes());
        let units = name.encode_utf16().collect::<Vec<_>>();
        data.extend_from_slice(&(units.len() as u16).to_le_bytes());
        data.extend(units.into_iter().flat_map(u16::to_le_bytes));
        data.extend_from_slice(&0_u16.to_le_bytes());
        for property_set in property_sets {
            data.extend_from_slice(&(property_set.len() as u16).to_le_bytes());
            data.extend_from_slice(property_set);
            if property_set.len() % 2 != 0 {
                data.push(0);
            }
        }
        let size = data.len() as u16;
        data[6..8].copy_from_slice(&size.to_le_bytes());
        data
    }

    fn paragraph_stylesheet() -> StyleSheet {
        let mut slots = vec![None; 17];
        slots[0] = Some(style_record(0, 1, 0x0fff, 0, "Normal", &[&[], &[]]));
        slots[15] = Some(style_record(
            0x0ffe,
            1,
            0,
            0,
            "Base Paragraph",
            &[&[15, 0, 0x03, 0x24, 2], &[]],
        ));
        slots[16] = Some(style_record(
            0x0ffe,
            1,
            15,
            0,
            "Derived Paragraph",
            &[&[0x03, 0x24, 1], &[]],
        ));

        let mut data = Vec::new();
        data.extend_from_slice(&18_u16.to_le_bytes());
        data.extend_from_slice(&(slots.len() as u16).to_le_bytes());
        data.extend_from_slice(&10_u16.to_le_bytes());
        data.extend_from_slice(&1_u16.to_le_bytes());
        data.extend_from_slice(&15_u16.to_le_bytes());
        data.extend_from_slice(&15_u16.to_le_bytes());
        data.extend_from_slice(&0_u16.to_le_bytes());
        data.extend_from_slice(&0_i16.to_le_bytes());
        data.extend_from_slice(&0_i16.to_le_bytes());
        data.extend_from_slice(&0_i16.to_le_bytes());
        for slot in slots {
            if let Some(record) = slot {
                data.extend_from_slice(&(record.len() as u16).to_le_bytes());
                data.extend_from_slice(&record);
                if record.len() % 2 != 0 {
                    data.push(0);
                }
            } else {
                data.extend_from_slice(&0_u16.to_le_bytes());
            }
        }
        StyleSheet::parse_data(&data, 0, Leniency::Strict).unwrap()
    }

    #[test]
    fn resolves_huge_papx_and_preserves_style() {
        let mut data = vec![0; 64];
        let direct = [
            0x12, 0x64, 0xf0, 0x00, 0x01, 0x00, 0x12, 0x64, 0xf0, 0x00, 0x01, 0x00,
        ];
        data[20..22].copy_from_slice(&(direct.len() as u16).to_le_bytes());
        data[22..22 + direct.len()].copy_from_slice(&direct);

        let mut papx = 7u16.to_le_bytes().to_vec();
        papx.extend_from_slice(&SPRM_P_HUGE_PAPX.to_le_bytes());
        papx.extend_from_slice(&20u32.to_le_bytes());

        let (properties, direct_grpprl, initial_style_index) =
            PapBinTable::parse_properties_with_direct(&papx, &[], Some(&data), None).unwrap();
        assert_eq!(properties.style_index, Some(7));
        assert_eq!(properties.line_spacing, Some(240));
        assert_eq!(direct_grpprl, direct);
        assert_eq!(initial_style_index, Some(7));
    }

    #[test]
    fn rejects_cyclic_data_indirections() {
        let mut data = vec![0; 64];
        let mut reference = SPRM_P_HUGE_PAPX.to_le_bytes().to_vec();
        reference.extend_from_slice(&20u32.to_le_bytes());
        reference.resize(10, 0);
        data[20..22].copy_from_slice(&(reference.len() as u16).to_le_bytes());
        data[22..22 + reference.len()].copy_from_slice(&reference);

        let mut visited = HashSet::new();
        let result = PapBinTable::expand_data_indirections(&reference, &data, &mut visited, 0);
        assert!(result.is_err());
    }

    #[test]
    fn rejects_missing_or_truncated_papx_data_indirections() {
        let mut papx = 7u16.to_le_bytes().to_vec();
        papx.extend_from_slice(&SPRM_P_HUGE_PAPX.to_le_bytes());
        papx.extend_from_slice(&20u32.to_le_bytes());
        assert!(PapBinTable::parse_properties(&papx, &[], None, None).is_err());

        let mut truncated_data = vec![0; 24];
        truncated_data[20..22].copy_from_slice(&10u16.to_le_bytes());
        assert!(PapBinTable::parse_properties(&papx, &[], Some(&truncated_data), None).is_err());
    }

    #[test]
    fn applies_piece_modifiers_after_fkp_properties() {
        let papx = [0x00, 0x00, 0x03, 0x24, 0x00];
        let piece_modifier = [0x03, 0x24, 0x02];
        let properties = PapBinTable::parse_properties(&papx, &piece_modifier, None, None).unwrap();
        assert_eq!(
            properties.justification,
            super::super::pap::Justification::Right
        );
    }

    #[test]
    fn adjacent_style_cache_matches_scalar_cascade_and_rekeys() {
        let stylesheet = paragraph_stylesheet();
        let mut cache = None;

        let derived_with_piece_override = [16, 0, 0x03, 0x24, 0];
        let piece_modifier = [0x03, 0x24, 2];
        let cached = PapBinTable::parse_properties_with_direct_cached(
            &derived_with_piece_override,
            &piece_modifier,
            None,
            Some(&stylesheet),
            &mut cache,
        )
        .unwrap();
        let scalar = PapBinTable::parse_properties_with_direct(
            &derived_with_piece_override,
            &piece_modifier,
            None,
            Some(&stylesheet),
        )
        .unwrap();
        assert_eq!(format!("{:?}", cached.0), format!("{:?}", scalar.0));
        assert_eq!(cached.1, scalar.1);
        assert_eq!(cached.2, scalar.2);
        assert_eq!(cache.as_ref().map(|(index, _)| *index), Some(16));

        let switch_to_base = [16, 0, 0x00, 0x46, 15, 0];
        let cached = PapBinTable::parse_properties_with_direct_cached(
            &switch_to_base,
            &[],
            None,
            Some(&stylesheet),
            &mut cache,
        )
        .unwrap();
        let scalar = PapBinTable::parse_properties_with_direct(
            &switch_to_base,
            &[],
            None,
            Some(&stylesheet),
        )
        .unwrap();
        assert_eq!(format!("{:?}", cached.0), format!("{:?}", scalar.0));
        assert_eq!(cached.0.style_index, Some(15));
        assert_eq!(cache.as_ref().map(|(index, _)| *index), Some(16));

        PapBinTable::parse_properties_with_direct_cached(
            &[15, 0],
            &[],
            None,
            Some(&stylesheet),
            &mut cache,
        )
        .unwrap();
        assert_eq!(cache.as_ref().map(|(index, _)| *index), Some(15));
    }
}
