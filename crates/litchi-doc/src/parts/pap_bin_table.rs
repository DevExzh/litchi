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
use crate::sprm::{Sprm, parse_sprms};
use crate::sprm_operations::{SPRM_P_HUGE_PAPX, SPRM_P_TABLE_PROPS};

const FKP_PAGE_SIZE: usize = 512;
const MAX_DATA_INDIRECTION_DEPTH: usize = 64;
/// Upper bound on resolved style baselines cached per `PapBinTable::parse`.
///
/// Real documents draw on a handful of paragraph styles, so this bound is not
/// reached; it keeps the cache finite for a hostile stylesheet that names
/// thousands. Eviction is least-recently-used, so the baseline a run of
/// same-style paragraphs needs always stays resident and no input resolves a
/// baseline more often than the one-entry cache of change 0051 did.
const MAX_CACHED_STYLE_BASELINES: usize = 64;
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

/// Bounded cache of style baselines resolved during one `PapBinTable::parse`.
///
/// `ParagraphProperties::resolve_style_baseline` is a pure function of the style
/// index and the immutable stylesheet, so a baseline resolved for one run is the
/// baseline every later run with that index would resolve. Only successful
/// resolutions are retained, exactly as in change 0051.
#[derive(Debug, Default)]
struct StyleBaselineCache {
    /// `(istd, resolved baseline, last-use stamp)`, at most
    /// [`MAX_CACHED_STYLE_BASELINES`] entries.
    entries: Vec<(u16, ParagraphProperties, u64)>,
    /// Monotonic use counter driving least-recently-used eviction.
    clock: u64,
}

impl StyleBaselineCache {
    fn baseline(&mut self, index: u16, stylesheet: &StyleSheet) -> Result<&ParagraphProperties> {
        self.clock = self.clock.wrapping_add(1);
        let position = match self
            .entries
            .iter()
            .position(|(cached, _, _)| *cached == index)
        {
            Some(position) => position,
            None => {
                let baseline =
                    ParagraphProperties::resolve_style_baseline(Some(index), stylesheet)?;
                if self.entries.len() < MAX_CACHED_STYLE_BASELINES {
                    self.entries.push((index, baseline, self.clock));
                    self.entries.len() - 1
                } else {
                    let victim = self
                        .entries
                        .iter()
                        .enumerate()
                        .min_by_key(|(_, (_, _, used))| *used)
                        .map_or(0, |(position, _)| position);
                    self.entries[victim] = (index, baseline, self.clock);
                    victim
                }
            },
        };
        self.entries[position].2 = self.clock;
        Ok(&self.entries[position].1)
    }
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
        let mut style_baseline_cache = StyleBaselineCache::default();
        // One `PrcData` visit set, cleared per entry, rather than one per entry.
        let mut visited = HashSet::new();
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
                // The checked-in MS-DOC grammar says these bytes form a
                // GrpPrlAndIstd of whole Prl elements and does not name a
                // PAPX pad. A repeated producer witness nevertheless has an
                // odd SPRM prefix plus one zero byte. Keep the shared SPRM
                // parser strict and remove exactly that byte only when it is
                // otherwise proven to be a final incomplete opcode.
                let grpprl = Self::trim_papx_word_alignment_pad(&entry.grpprl);
                for (start_cp, end_cp) in piece_table.fc_range_to_cp_ranges(entry.fc, entry.end_fc)
                {
                    let piece_modifier = piece_table
                        .piece_for_cp(start_cp)
                        .map(super::piece_table::TextPiece::property_modifier)
                        .unwrap_or_default();
                    let (properties, direct_grpprl, initial_style_index) =
                        Self::parse_properties_with_direct_cached(
                            grpprl,
                            piece_modifier,
                            data_stream,
                            stylesheet,
                            &mut style_baseline_cache,
                            &mut visited,
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

    /// Remove the single producer-compatibility byte after an odd PAPX SPRM
    /// sequence. The local MS-DOC grammar does not name this byte, and
    /// nonzero `cb` encodings are odd-sized by definition, so an even-sized
    /// `GrpPrlAndIstd` can only come from the `cb=0`/`cb'` form. The byte is
    /// accepted only when strict SPRM parsing identifies every preceding byte
    /// as a complete sequence and the final zero as exactly one incomplete
    /// opcode byte. Any other malformed input remains untouched and is rejected
    /// by the normal typed parser.
    fn trim_papx_word_alignment_pad(grpprl_and_istd: &[u8]) -> &[u8] {
        if grpprl_and_istd.len() < 3
            || !grpprl_and_istd.len().is_multiple_of(2)
            || grpprl_and_istd.last() != Some(&0)
        {
            return grpprl_and_istd;
        }

        let direct_sprms = &grpprl_and_istd[2..];
        match parse_sprms(direct_sprms) {
            Err(crate::sprm::Error::Opcode { at, remaining: 1 })
                if at.saturating_add(1) == direct_sprms.len() =>
            {
                &grpprl_and_istd[..grpprl_and_istd.len() - 1]
            },
            _ => grpprl_and_istd,
        }
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
            &mut StyleBaselineCache::default(),
            &mut HashSet::new(),
        )
    }

    fn parse_properties_with_direct_cached(
        grpprl_and_istd: &[u8],
        piece_modifier: &[u8],
        data_stream: Option<&[u8]>,
        stylesheet: Option<&StyleSheet>,
        style_baseline_cache: &mut StyleBaselineCache,
        visited: &mut HashSet<u32>,
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

        // Both arms below opened by parsing `direct_sprms`, so the entry is
        // parsed once here and every later consumer reads that one parse.
        let parsed = parse_sprms(direct_sprms)?;
        let expanded = if let Some(data) = data_stream {
            visited.clear();
            Self::expand_parsed_indirections(direct_sprms, &parsed, data, visited, 0)?
        } else {
            if parsed.iter().any(|sprm| {
                matches!(
                    sprm.opcode,
                    SPRM_P_HUGE_PAPX | SPRM_P_TABLE_PROPS | SPRM_P_TABLE_PROPS_LEGACY
                )
            }) {
                return Err(PackageError::Corrupted(
                    "PAPX data indirection requires a Data Stream".to_string(),
                ));
            }
            None
        };
        let sprms = expanded.as_deref().unwrap_or(direct_sprms);

        // With no expansion and no piece modifier the concatenation reproduces
        // `direct_sprms` byte for byte, so `parsed` still describes it exactly
        // and the cascade below need not parse the same bytes again.
        let unmodified = expanded.is_none() && piece_modifier.is_empty();
        let direct_grpprl = [sprms, piece_modifier].concat();
        let pre_parsed = unmodified.then_some(parsed.as_slice());
        let mut properties = match (stylesheet, style_index) {
            (Some(styles), Some(index)) => {
                let baseline = style_baseline_cache.baseline(index, styles)?;
                ParagraphProperties::cascade_styles_from_resolved_baseline(
                    baseline,
                    &direct_grpprl,
                    styles,
                    pre_parsed,
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

        let sprms = parse_sprms(grpprl)?;
        Self::expand_parsed_indirections(grpprl, &sprms, data_stream, visited, depth)
    }

    /// [`Self::expand_data_indirections`] over an already parsed `grpprl`.
    ///
    /// `sprms` must be `parse_sprms(grpprl)` for the very same bytes. The depth
    /// limit is checked by the recursive entry point before it parses, so a
    /// chain that is too deep is still reported as such rather than as whatever
    /// its innermost `grpprl` happens to be.
    fn expand_parsed_indirections(
        grpprl: &[u8],
        sprms: &[Sprm],
        data_stream: &[u8],
        visited: &mut HashSet<u32>,
        depth: usize,
    ) -> Result<Option<Vec<u8>>> {
        for sprm in sprms {
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
    fn trims_one_zero_byte_after_a_complete_odd_papx_sprm_sequence() {
        let grpprl = [0x00, 0x00, 0x31, 0x24, 0x00, 0x00];
        assert_eq!(
            PapBinTable::trim_papx_word_alignment_pad(&grpprl),
            &grpprl[..5]
        );
    }

    #[test]
    fn keeps_non_padding_trailing_bytes_for_typed_rejection() {
        let grpprl = [0x00, 0x00, 0x31, 0x24, 0x00, 0x01];
        assert_eq!(PapBinTable::trim_papx_word_alignment_pad(&grpprl), &grpprl);
    }

    #[test]
    fn adjacent_style_cache_matches_scalar_cascade_and_rekeys() {
        let stylesheet = paragraph_stylesheet();
        let mut cache = StyleBaselineCache::default();
        let mut visited = HashSet::new();
        let keys = |cache: &StyleBaselineCache| {
            cache
                .entries
                .iter()
                .map(|(index, _, _)| *index)
                .collect::<Vec<_>>()
        };

        let derived_with_piece_override = [16, 0, 0x03, 0x24, 0];
        let piece_modifier = [0x03, 0x24, 2];
        let cached = PapBinTable::parse_properties_with_direct_cached(
            &derived_with_piece_override,
            &piece_modifier,
            None,
            Some(&stylesheet),
            &mut cache,
            &mut visited,
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
        assert_eq!(keys(&cache), vec![16]);

        let switch_to_base = [16, 0, 0x00, 0x46, 15, 0];
        let cached = PapBinTable::parse_properties_with_direct_cached(
            &switch_to_base,
            &[],
            None,
            Some(&stylesheet),
            &mut cache,
            &mut visited,
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
        // A direct `sprmPIstd` never re-keys the cache: only the PAPX header's
        // initial style does, exactly as in change 0051.
        assert_eq!(keys(&cache), vec![16]);

        PapBinTable::parse_properties_with_direct_cached(
            &[15, 0],
            &[],
            None,
            Some(&stylesheet),
            &mut cache,
            &mut visited,
        )
        .unwrap();
        assert_eq!(keys(&cache), vec![16, 15]);

        // Returning to style 16 is now a hit rather than a re-resolution, and
        // both baselines stay resident.
        PapBinTable::parse_properties_with_direct_cached(
            &derived_with_piece_override,
            &piece_modifier,
            None,
            Some(&stylesheet),
            &mut cache,
            &mut visited,
        )
        .unwrap();
        assert_eq!(keys(&cache), vec![16, 15]);
    }

    #[test]
    fn style_baseline_cache_is_bounded_and_evicts_least_recently_used() {
        let stylesheet = paragraph_stylesheet();
        let mut cache = StyleBaselineCache::default();
        let baseline = ParagraphProperties::default();

        // Fill the cache with synthetic entries so the bound is exercised
        // without needing a stylesheet that names thousands of styles.
        for index in 0..MAX_CACHED_STYLE_BASELINES {
            cache.clock += 1;
            cache
                .entries
                .push((1000 + index as u16, baseline.clone(), cache.clock));
        }
        assert_eq!(cache.entries.len(), MAX_CACHED_STYLE_BASELINES);

        // Touch the oldest entry so it is no longer the eviction victim.
        cache.baseline(1000, &stylesheet).unwrap();
        // A new style evicts entry 1001, the least recently used, and never
        // grows the cache past its bound.
        cache.baseline(15, &stylesheet).unwrap();
        assert_eq!(cache.entries.len(), MAX_CACHED_STYLE_BASELINES);
        let keys = cache
            .entries
            .iter()
            .map(|(index, _, _)| *index)
            .collect::<Vec<_>>();
        assert!(
            keys.contains(&1000),
            "the freshly used entry stays resident"
        );
        assert!(keys.contains(&15), "the new baseline is cached");
        assert!(
            !keys.contains(&1001),
            "the least recently used entry is gone"
        );
    }
}
