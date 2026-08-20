/// Proper paragraph extraction from binary structures.
///
/// Based on Apache POI's HWPF paragraph parsing logic, this module
/// extracts paragraphs using reconstructed paragraph and character FKP runs.
use super::super::package::Result;
use super::chp::CharacterProperties;
use super::chp_bin_table::ChpBinTable;
use super::pap::ParagraphProperties;
use super::pap_bin_table::PapBinTable;
use super::styles::StyleSheet;
use super::tap::{TableStyleCondition, matching_table_style_conditions};
use crate::sprm::parse_sprms;
use crate::sprm_operations::{SPRM_C_ISTD, get_sprm_type};
use std::sync::Arc;

/// Type alias for extracted paragraph data: (text, properties, runs).
pub(crate) type ExtractedParagraph = (
    String,
    ParagraphProperties,
    Vec<(String, CharacterProperties)>,
);

#[derive(Debug)]
struct TableTextContext {
    style_index: u16,
    conditions: Vec<TableStyleCondition>,
}

/// Paragraph extractor using binary structures.
///
/// Based on Apache POI's `ParagraphPropertiesTable` (`PAPBinTable`) and
/// `CharacterPropertiesTable` (`CHPBinTable`).
pub struct ParagraphExtractor<'a> {
    /// Paragraph property bin table (shared reference to avoid re-parsing)
    pap_bin_table: Option<&'a PapBinTable>,
    /// Character property bin table (shared reference to avoid re-parsing)
    chp_bin_table: Option<&'a ChpBinTable>,
    /// The extracted text (shared via Arc to avoid cloning, thread-safe)
    text: Arc<String>,
    /// Text piece character positions
    text_ranges: Vec<usize>, // UTF-16 CP boundary -> UTF-8 byte offset
    /// Character position range to extract (for subdocuments)
    cp_range: Option<(u32, u32)>,
    /// Stylesheet used to apply paragraph-style character properties.
    stylesheet: Option<&'a StyleSheet>,
}

impl<'a> ParagraphExtractor<'a> {
    /// Create a new paragraph extractor.
    ///
    /// # Arguments
    ///
    /// * `text` - Extracted document text (shared via Arc, thread-safe)
    /// * `pap_bin_table` - Pre-parsed paragraph property bin table
    /// * `chp_bin_table` - Pre-parsed character property bin table (avoids re-parsing)
    pub fn new(
        text: Arc<String>,
        pap_bin_table: Option<&'a PapBinTable>,
        chp_bin_table: Option<&'a ChpBinTable>,
    ) -> Result<Self> {
        // Build text ranges for mapping CPs to text offsets
        let text_ranges = Self::build_text_ranges(&text);

        Ok(Self {
            pap_bin_table,
            chp_bin_table,
            text,
            text_ranges,
            cp_range: None,
            stylesheet: None,
        })
    }

    /// Create a new paragraph extractor for a specific character position range.
    ///
    /// This is used to extract paragraphs from subdocuments (footnotes, headers, etc.).
    ///
    /// # Arguments
    ///
    /// * `text` - Extracted document text (shared via Arc, thread-safe)
    /// * `pap_bin_table` - Pre-parsed paragraph property bin table
    /// * `chp_bin_table` - Pre-parsed character property bin table (avoids re-parsing)
    /// * `cp_range` - Character position range (`start_cp`, `end_cp`)
    pub fn new_with_range(
        text: Arc<String>,
        pap_bin_table: Option<&'a PapBinTable>,
        chp_bin_table: Option<&'a ChpBinTable>,
        cp_range: (u32, u32),
    ) -> Result<Self> {
        let mut extractor = Self::new(text, pap_bin_table, chp_bin_table)?;
        extractor.cp_range = Some(cp_range);
        Ok(extractor)
    }

    /// Create a range extractor that applies paragraph and character styles.
    pub(crate) fn new_with_range_and_stylesheet(
        text: Arc<String>,
        pap_bin_table: Option<&'a PapBinTable>,
        chp_bin_table: Option<&'a ChpBinTable>,
        cp_range: (u32, u32),
        stylesheet: Option<&'a StyleSheet>,
    ) -> Result<Self> {
        let mut extractor = Self::new_with_range(text, pap_bin_table, chp_bin_table, cp_range)?;
        extractor.stylesheet = stylesheet;
        Ok(extractor)
    }

    /// Build mapping from character positions to text offsets.
    fn build_text_ranges(text: &str) -> Vec<usize> {
        let mut ranges = Vec::with_capacity(text.encode_utf16().count() + 1);
        for (offset, ch) in text.char_indices() {
            for _ in 0..ch.len_utf16() {
                ranges.push(offset);
            }
        }
        ranges.push(text.len());
        ranges
    }

    /// Extract paragraphs with properties.
    ///
    /// Returns a vector of (text, `paragraph_properties`, `character_runs`) tuples.
    ///
    /// Based on MS-DOC specification and Apache POI's approach:
    /// Paragraphs in Word documents are delimited by CR (\r = 0x000D) characters.
    /// The PAP PLCF stores formatting properties, but doesn't define paragraph boundaries.
    pub fn extract_paragraphs(&self) -> Result<Vec<ExtractedParagraph>> {
        let mut paragraphs = Vec::new();

        // Determine the CP range to process
        let doc_start_cp = self.cp_range.map_or(0, |(start, _)| start);
        let doc_end_cp = self.cp_range.map_or_else(
            || self.text_ranges.len().saturating_sub(1) as u32,
            |(_, end)| end,
        );

        // Find all paragraph breaks (CR characters) in the text
        // CR (0x000D / '\r') marks the end of each paragraph in Word documents
        let mut para_boundaries = vec![doc_start_cp];
        let mut current_cp = 0u32;

        for c in self.text.chars() {
            let next_cp = current_cp + c.len_utf16() as u32;
            if next_cp <= doc_start_cp {
                current_cp = next_cp;
                continue;
            }
            if current_cp >= doc_end_cp {
                break;
            }
            if matches!(c, '\r' | '\u{7}') && next_cp <= doc_end_cp {
                para_boundaries.push(next_cp); // Position after paragraph/cell marker
            }
            current_cp = next_cp;
        }

        // Ensure we have an end boundary
        if para_boundaries.last() != Some(&doc_end_cp) && current_cp > doc_start_cp {
            para_boundaries.push(current_cp.min(doc_end_cp));
        }

        // Extract each paragraph
        for i in 0..para_boundaries.len().saturating_sub(1) {
            let para_start = para_boundaries[i];
            let para_end = para_boundaries[i + 1];

            if para_start >= para_end {
                continue;
            }

            // Extract paragraph text (excluding the CR marker itself)
            let terminator = self.extract_text_range(para_end - 1, para_end);
            let mut para_text = self.extract_text_range(para_start, para_end);
            // Structural terminators are not paragraph content.
            if matches!(terminator.as_str(), "\r" | "\u{7}") {
                para_text.pop();
            }

            // Find matching PAP properties for this paragraph
            let mut para_props = self
                .pap_bin_table
                .and_then(|table| table.properties_at(para_start))
                .cloned()
                .unwrap_or_default();
            let table_context = self.table_text_context(para_start, &para_props);
            let (table_papx, table_chpx) = if let (Some(stylesheet), Some(context)) =
                (self.stylesheet, table_context.as_ref())
            {
                let (_, paragraph, character) = stylesheet
                    .resolve_table_text_style_sprms_for_conditions(
                        context.style_index,
                        &context.conditions,
                    )?;
                (paragraph, character)
            } else {
                (Vec::new(), Vec::new())
            };
            if let (Some(stylesheet), Some(paragraph_run)) = (
                self.stylesheet,
                self.pap_bin_table
                    .and_then(|table| table.run_at(para_start)),
            ) && !table_papx.is_empty()
            {
                para_props = ParagraphProperties::cascade_table_style(
                    &table_papx,
                    paragraph_run.initial_style_index,
                    &paragraph_run.direct_grpprl,
                    stylesheet,
                )?;
            }
            para_props.is_table_cell_end = terminator == "\u{7}";

            // Extract character runs within this paragraph (excluding the CR)
            let para_text_end = if matches!(terminator.as_str(), "\r" | "\u{7}") {
                para_end - 1
            } else {
                para_end
            };
            let runs = self.extract_runs(
                para_start,
                para_text_end,
                para_props.style_index,
                &table_chpx,
            )?;

            paragraphs.push((para_text, para_props, runs));
        }

        // Fallback if no paragraphs were found
        if paragraphs.is_empty() && !self.text.is_empty() {
            let runs = self.extract_runs(doc_start_cp, doc_end_cp, None, &[])?;
            paragraphs.push((
                self.text.as_ref().clone(),
                ParagraphProperties::default(),
                runs,
            ));
        }

        Ok(paragraphs)
    }

    /// Count the logical paragraphs in a UTF-16 character-position range.
    ///
    /// Paragraph counting is intentionally independent from semantic extraction: it only
    /// classifies the structural paragraph and table-cell terminators and therefore does not
    /// parse PAP/CHP runs, styles, images, or formulas. The fallback for non-empty source text
    /// with no structural terminator matches [`Self::extract_paragraphs`] exactly.
    pub(crate) fn count_paragraphs_in_range(text: &str, cp_range: (u32, u32)) -> usize {
        let (doc_start_cp, doc_end_cp) = cp_range;
        let mut paragraph_count = 0usize;
        let mut current_cp = 0u32;
        let mut has_text = false;
        let mut last_boundary = doc_start_cp;

        for character in text.chars() {
            let next_cp = current_cp.saturating_add(character.len_utf16() as u32);
            if next_cp <= doc_start_cp {
                current_cp = next_cp;
                continue;
            }
            if current_cp >= doc_end_cp {
                break;
            }
            has_text = true;
            if matches!(character, '\r' | '\u{7}') && next_cp <= doc_end_cp {
                paragraph_count = paragraph_count.saturating_add(1);
                last_boundary = next_cp;
            }
            current_cp = next_cp;
        }

        // `extract_paragraphs` appends a final boundary for trailing text, including text that
        // ends before an over-large requested range. A range with only structural markers already
        // has one paragraph per marker and must not receive an additional fallback paragraph.
        if paragraph_count == 0 && !text.is_empty() {
            return 1;
        }
        if has_text
            && current_cp.min(doc_end_cp) > doc_start_cp
            && current_cp.min(doc_end_cp) > last_boundary
        {
            paragraph_count = paragraph_count.saturating_add(1);
        }

        paragraph_count
    }

    /// Extract text for a character position range.
    fn extract_text_range(&self, cp_start: u32, cp_end: u32) -> String {
        // Clamp CPs to valid range
        let max_cp = self.text_ranges.len().saturating_sub(1) as u32;
        let cp_start_clamped = cp_start.min(max_cp);
        let cp_end_clamped = cp_end.min(max_cp);

        if cp_start_clamped >= cp_end_clamped {
            return String::new();
        }

        let start_idx = cp_start_clamped as usize;
        let end_idx = cp_end_clamped as usize;

        if start_idx < self.text_ranges.len() {
            let start_offset = self.text_ranges[start_idx];
            let end_offset = self.text_ranges[end_idx];

            if start_offset <= end_offset {
                self.text[start_offset..end_offset].to_string()
            } else {
                String::new()
            }
        } else {
            String::new()
        }
    }

    fn table_text_context(
        &self,
        cp: u32,
        properties: &ParagraphProperties,
    ) -> Option<TableTextContext> {
        if !properties.in_table {
            return None;
        }
        let level = properties.table_nesting_level;
        let table = self.pap_bin_table?;
        let runs = table.runs();
        let current = runs
            .partition_point(|run| run.start_cp <= cp)
            .checked_sub(1)?;
        if cp >= runs.get(current)?.end_cp {
            return None;
        }

        let mut table_start = current;
        while table_start > 0 {
            let candidate = &runs[table_start - 1].properties;
            if !candidate.in_table || candidate.table_nesting_level < level {
                break;
            }
            table_start -= 1;
        }
        let mut table_end = current + 1;
        while let Some(run) = runs.get(table_end) {
            if !run.properties.in_table || run.properties.table_nesting_level < level {
                break;
            }
            table_end += 1;
        }

        let row_ends = (table_start..table_end)
            .filter(|&index| {
                let properties = &runs[index].properties;
                properties.table_nesting_level == level && properties.is_table_row_end
            })
            .collect::<Vec<_>>();
        let row_position = row_ends.iter().position(|&index| index >= current)?;
        let row_end_index = row_ends[row_position];
        let row_properties = runs[row_end_index].properties.table_properties.as_ref()?;
        let style_index = row_properties.table_style_index?;

        let row_start_cp = row_position
            .checked_sub(1)
            .map_or(runs[table_start].start_cp, |previous| {
                runs[row_ends[previous]].end_cp
            });
        let mut cell_index = 0usize;
        for marker_cp in row_start_cp..cp {
            if self.extract_text_range(marker_cp, marker_cp + 1) != "\u{7}" {
                continue;
            }
            if table
                .properties_at(marker_cp)
                .is_some_and(|marker| marker.table_nesting_level == level)
            {
                cell_index += 1;
            }
        }
        let table_rows = row_ends
            .iter()
            .map(|&index| runs[index].properties.table_properties.as_ref())
            .collect::<Option<Vec<_>>>()?;
        let conditions = matching_table_style_conditions(&table_rows, row_position, cell_index);

        Some(TableTextContext {
            style_index,
            conditions,
        })
    }

    /// Extract character runs (formatted text segments) within a paragraph.
    fn extract_runs(
        &self,
        para_start: u32,
        para_end: u32,
        paragraph_style_index: Option<u16>,
        table_style_chpx: &[u8],
    ) -> Result<Vec<(String, CharacterProperties)>> {
        let mut runs = Vec::new();
        let paragraph_style_chpx = self
            .stylesheet
            .zip(paragraph_style_index)
            .map(|(styles, index)| styles.resolve_paragraph_style_sprms(index))
            .transpose()?
            .map(|(_, _, character)| character)
            .unwrap_or_default();

        if let Some(chp_bin_table) = self.chp_bin_table {
            // Get runs that overlap with this paragraph
            let overlapping_runs = chp_bin_table.runs_in_range(para_start, para_end);

            for run in overlapping_runs {
                // Calculate actual run boundaries within paragraph
                let actual_start = run.start_cp.max(para_start);
                let actual_end = run.end_cp.min(para_end);

                if actual_start >= actual_end {
                    continue;
                }

                // Extract run text
                let run_text = self.extract_text_range(actual_start, actual_end);

                // Skip empty runs
                if run_text.is_empty() {
                    continue;
                }

                let properties = cascade_character_properties_with_table(
                    self.stylesheet,
                    table_style_chpx,
                    &paragraph_style_chpx,
                    &run.properties,
                    &run.direct_grpprl,
                )?;
                runs.push((run_text, properties));
            }
        }

        // If no runs found, return the whole paragraph as one run
        if runs.is_empty() {
            let para_text = self.extract_text_range(para_start, para_end);
            if !para_text.is_empty() {
                let properties = cascade_character_properties_with_table(
                    self.stylesheet,
                    table_style_chpx,
                    &paragraph_style_chpx,
                    &CharacterProperties::default(),
                    &[],
                )?;
                runs.push((para_text, properties));
            }
        } else {
            // Filter out empty runs before consolidation to prevent style markers without text
            runs.retain(|(text, _)| !text.is_empty());

            // Consolidate consecutive runs with identical formatting
            // This prevents markdown like **a****b****c** instead of **abc**
            if !runs.is_empty() {
                runs = Self::consolidate_runs(runs);
            }
        }

        Ok(runs)
    }

    /// Consolidate consecutive runs with identical formatting properties.
    ///
    /// This prevents creating separate runs for each character when they have
    /// the same styling, which would result in markdown like `**a****b****c**`
    /// instead of the desired `**abc**`.
    ///
    /// # Note
    ///
    /// Empty runs are skipped to prevent style markers without text (e.g., `****`).
    fn consolidate_runs(
        runs: Vec<(String, CharacterProperties)>,
    ) -> Vec<(String, CharacterProperties)> {
        if runs.is_empty() {
            return runs;
        }

        let mut consolidated = Vec::new();
        let mut current_text = String::new();
        let mut current_props = runs[0].1.clone();

        for (text, props) in runs {
            // Skip empty runs entirely to prevent empty style markers
            if text.is_empty() {
                continue;
            }

            // Every CHP field affects the meaning of a run. In particular,
            // revision and object-reference metadata must remain attached to
            // its exact text range even when the visible formatting matches.
            if props == current_props {
                // Same formatting - append to current run
                current_text.push_str(&text);
            } else {
                // Different formatting - save current run and start new one
                if !current_text.is_empty() {
                    consolidated.push((current_text.clone(), current_props.clone()));
                }
                current_text = text;
                current_props = props;
            }
        }

        // Don't forget the last run (only if non-empty)
        if !current_text.is_empty() {
            consolidated.push((current_text, current_props));
        }

        consolidated
    }
}

#[cfg(test)]
pub(crate) fn cascade_character_properties(
    stylesheet: Option<&StyleSheet>,
    paragraph_style_chpx: &[u8],
    direct_properties: &CharacterProperties,
    direct_grpprl: &[u8],
) -> Result<CharacterProperties> {
    cascade_character_properties_with_table(
        stylesheet,
        &[],
        paragraph_style_chpx,
        direct_properties,
        direct_grpprl,
    )
}

pub(crate) fn cascade_character_properties_with_table(
    stylesheet: Option<&StyleSheet>,
    table_style_chpx: &[u8],
    paragraph_style_chpx: &[u8],
    direct_properties: &CharacterProperties,
    direct_grpprl: &[u8],
) -> Result<CharacterProperties> {
    if stylesheet.is_none() && table_style_chpx.is_empty() && paragraph_style_chpx.is_empty() {
        return Ok(direct_properties.clone());
    }
    let mut paragraph_baseline = CharacterProperties::from_sprm(table_style_chpx)?;
    for sprm in parse_sprms(paragraph_style_chpx)? {
        CharacterProperties::apply_sprm(&mut paragraph_baseline, &sprm)?;
    }
    let mut current = paragraph_baseline.clone();
    let sprms = parse_sprms(direct_grpprl)?;
    let consumed = sprms.last().map_or(0, |sprm| sprm.offset + sprm.size);
    if consumed != direct_grpprl.len() {
        return Err(super::super::package::Error::Corrupted(
            "CHPX grpprl does not contain a whole number of SPRMs".to_string(),
        ));
    }
    for sprm in &sprms {
        if get_sprm_type(sprm.opcode) != 2 {
            continue;
        }
        if sprm.opcode == SPRM_C_ISTD {
            let requested = sprm.operand_word().ok_or_else(|| {
                super::super::package::Error::Corrupted(
                    "sprmCIstd is missing its style index".to_string(),
                )
            })?;
            let style_chpx = stylesheet
                .map(|styles| styles.resolve_character_style_sprms(requested))
                .transpose()?
                .map(|(_, character)| character)
                .unwrap_or_default();
            let mut styled = paragraph_baseline.clone();
            for style_sprm in parse_sprms(&style_chpx)? {
                CharacterProperties::apply_sprm(&mut styled, &style_sprm)?;
            }
            styled.style_index = Some(requested);
            preserve_character_style_state(&current, &mut styled);
            current = styled;
        } else {
            CharacterProperties::apply_sprm(&mut current, sprm)?;
        }
    }
    Ok(current)
}

fn preserve_character_style_state(
    previous: &CharacterProperties,
    styled: &mut CharacterProperties,
) {
    styled.is_revision_deleted = previous.is_revision_deleted;
    styled.deletion_author_index = previous.deletion_author_index;
    styled.deletion_timestamp = previous.deletion_timestamp;
    styled.deletion_revision_id = previous.deletion_revision_id;
    styled.deletion_revision_save_id = previous.deletion_revision_save_id;
    styled.is_bidi = previous.is_bidi;
    styled.is_complex_scripts = previous.is_complex_scripts;
    styled.is_revision_inserted = previous.is_revision_inserted;
    styled.revision_author_index = previous.revision_author_index;
    styled.revision_timestamp = previous.revision_timestamp;
    styled.revision_id = previous.revision_id;
    styled.is_spec = previous.is_spec;
    styled.is_data = previous.is_data;
    styled.is_ole2 = previous.is_ole2;
    styled.is_obj = previous.is_obj;
    styled.pic_offset = previous.pic_offset;
    styled.obj_offset = previous.obj_offset;
    styled.has_formatting_revision = previous.has_formatting_revision;
    styled.formatting_revision_author_index = previous.formatting_revision_author_index;
    styled.formatting_revision_timestamp = previous.formatting_revision_timestamp;
    styled.script_hint = previous.script_hint;
    styled.highlight = previous.highlight;
    styled.insertion_revision_save_id = previous.insertion_revision_save_id;
    styled.formatting_revision_save_id = previous.formatting_revision_save_id;
    styled.display_field_revision = previous.display_field_revision.clone();
}

#[cfg(test)]
mod tests {
    use super::super::pap_bin_table::ParagraphRun;
    use super::super::tap::{TableLook, TableLookFlags, TableProperties, TableStyleDefaults};
    use super::*;
    use crate::sprm_operations::{SPRM_C_F_BOLD, SPRM_C_F_ITALIC};

    #[test]
    fn consolidation_preserves_non_visual_character_metadata_boundaries() {
        let inserted = CharacterProperties {
            is_revision_inserted: Some(true),
            revision_author_index: Some(1),
            ..CharacterProperties::default()
        };
        let deleted = CharacterProperties {
            is_revision_deleted: Some(true),
            deletion_author_index: Some(2),
            ..CharacterProperties::default()
        };

        let runs = ParagraphExtractor::consolidate_runs(vec![
            ("inserted ".to_string(), inserted.clone()),
            ("deleted".to_string(), deleted.clone()),
        ]);

        assert_eq!(runs.len(), 2);
        assert_eq!(runs[0], ("inserted ".to_string(), inserted));
        assert_eq!(runs[1], ("deleted".to_string(), deleted));
    }

    #[test]
    fn cell_and_row_marks_form_structural_paragraph_boundaries() {
        let extractor =
            ParagraphExtractor::new(Arc::new("Cell\u{7}\u{7}".to_string()), None, None).unwrap();
        let paragraphs = extractor.extract_paragraphs().unwrap();
        assert_eq!(paragraphs.len(), 2);
        assert_eq!(paragraphs[0].0, "Cell");
        assert!(paragraphs[0].1.is_table_cell_end);
        assert_eq!(paragraphs[1].0, "");
        assert!(paragraphs[1].1.is_table_cell_end);
    }

    #[test]
    fn paragraph_count_matches_extraction_without_parsing_runs() {
        let text = "first\rsecond 😀\rCell\u{7}\u{7}tail";
        let extractor = ParagraphExtractor::new(Arc::new(text.to_owned()), None, None).unwrap();
        let utf16_len = text.encode_utf16().count() as u32;
        let ranges = [
            (0, utf16_len),
            (0, 6),
            (6, utf16_len),
            (7, 17),
            (17, utf16_len),
            (utf16_len, utf16_len),
            (utf16_len + 1, utf16_len + 2),
            (utf16_len + 2, utf16_len + 1),
        ];

        for range in ranges {
            let ranged =
                ParagraphExtractor::new_with_range(Arc::new(text.to_owned()), None, None, range)
                    .unwrap();
            assert_eq!(
                ParagraphExtractor::count_paragraphs_in_range(text, range),
                ranged.extract_paragraphs().unwrap().len(),
                "paragraph count mismatch for range {range:?}"
            );
        }

        assert_eq!(
            ParagraphExtractor::count_paragraphs_in_range(text, (0, utf16_len)),
            extractor.extract_paragraphs().unwrap().len()
        );

        let terminated = "first\r";
        let terminated_extractor =
            ParagraphExtractor::new(Arc::new(terminated.to_owned()), None, None).unwrap();
        let overlarge_range = (0, terminated.encode_utf16().count() as u32 + 5);
        assert_eq!(
            ParagraphExtractor::count_paragraphs_in_range(terminated, overlarge_range),
            terminated_extractor.extract_paragraphs().unwrap().len()
        );
    }

    #[test]
    fn cascades_table_character_defaults_before_paragraph_and_direct_properties() {
        let table = [SPRM_C_F_BOLD.to_le_bytes().as_slice(), &[1]].concat();
        let paragraph = [SPRM_C_F_ITALIC.to_le_bytes().as_slice(), &[1]].concat();
        let direct_grpprl = [SPRM_C_F_BOLD.to_le_bytes().as_slice(), &[0]].concat();
        let direct = CharacterProperties::from_sprm(&direct_grpprl).unwrap();

        let properties = cascade_character_properties_with_table(
            None,
            &table,
            &paragraph,
            &direct,
            &direct_grpprl,
        )
        .unwrap();
        assert_eq!(properties.is_bold, Some(false));
        assert_eq!(properties.is_italic, Some(true));
    }

    #[test]
    fn selects_table_text_conditions_in_specified_precedence() {
        fn cell_properties() -> ParagraphProperties {
            ParagraphProperties {
                in_table: true,
                table_nesting_level: 1,
                ..ParagraphProperties::default()
            }
        }

        fn row_end_properties() -> ParagraphProperties {
            let table = TableProperties {
                cell_count: 2,
                style_defaults: TableStyleDefaults {
                    horizontal_band_size: Some(1),
                    vertical_band_size: Some(1),
                    ..TableStyleDefaults::default()
                },
                table_style_index: Some(15),
                table_look: Some(TableLook {
                    autoformat_index: -1,
                    flags: TableLookFlags::HEADER_ROW
                        | TableLookFlags::LAST_ROW
                        | TableLookFlags::HEADER_COLUMN
                        | TableLookFlags::LAST_COLUMN,
                }),
                ..TableProperties::default()
            };
            ParagraphProperties {
                in_table: true,
                is_table_row_end: true,
                table_nesting_level: 1,
                table_properties: Some(table),
                ..ParagraphProperties::default()
            }
        }

        fn run(start_cp: u32, end_cp: u32, properties: ParagraphProperties) -> ParagraphRun {
            ParagraphRun {
                start_cp,
                end_cp,
                properties,
                direct_grpprl: Vec::new(),
                initial_style_index: None,
                paragraph_height: None,
            }
        }

        let table = PapBinTable::from_runs_for_test(vec![
            run(0, 2, cell_properties()),
            run(2, 4, cell_properties()),
            run(4, 5, row_end_properties()),
            run(5, 7, cell_properties()),
            run(7, 9, cell_properties()),
            run(9, 10, row_end_properties()),
        ]);
        let extractor = ParagraphExtractor::new(
            Arc::new("a\u{7}b\u{7}\u{7}c\u{7}d\u{7}\u{7}".to_string()),
            Some(&table),
            None,
        )
        .unwrap();

        let top_left = extractor
            .table_text_context(0, table.properties_at(0).unwrap())
            .unwrap();
        assert_eq!(top_left.style_index, 15);
        assert_eq!(
            top_left.conditions,
            [
                TableStyleCondition::FirstColumn,
                TableStyleCondition::HeaderRow,
                TableStyleCondition::TopLeftCell,
            ]
        );
        let top_right = extractor
            .table_text_context(2, table.properties_at(2).unwrap())
            .unwrap();
        assert_eq!(
            top_right.conditions,
            [
                TableStyleCondition::LastColumn,
                TableStyleCondition::HeaderRow,
                TableStyleCondition::TopRightCell,
            ]
        );
        let bottom_left = extractor
            .table_text_context(5, table.properties_at(5).unwrap())
            .unwrap();
        assert_eq!(
            bottom_left.conditions,
            [
                TableStyleCondition::FirstColumn,
                TableStyleCondition::FooterRow,
                TableStyleCondition::BottomLeftCell,
            ]
        );

        let mut no_look_end = row_end_properties();
        no_look_end.table_properties.as_mut().unwrap().table_look = None;
        let no_look_table = PapBinTable::from_runs_for_test(vec![
            run(0, 2, cell_properties()),
            run(2, 3, no_look_end),
        ]);
        let no_look_extractor = ParagraphExtractor::new(
            Arc::new("a\u{7}\u{7}".to_string()),
            Some(&no_look_table),
            None,
        )
        .unwrap();
        let context = no_look_extractor
            .table_text_context(0, no_look_table.properties_at(0).unwrap())
            .unwrap();
        assert_eq!(context.style_index, 15);
        assert!(context.conditions.is_empty());
    }
}
