/// Proper paragraph extraction from binary structures.
///
/// Based on Apache POI's HWPF paragraph parsing logic, this module
/// extracts paragraphs using reconstructed paragraph and character FKP runs.
use super::super::package::Result;
use super::chp::CharacterProperties;
use super::chp_bin_table::ChpBinTable;
use super::pap::ParagraphProperties;
use super::pap_bin_table::PapBinTable;
use std::sync::Arc;

/// Type alias for extracted paragraph data: (text, properties, runs).
pub(crate) type ExtractedParagraph = (
    String,
    ParagraphProperties,
    Vec<(String, CharacterProperties)>,
);

/// Paragraph extractor using binary structures.
///
/// Based on Apache POI's ParagraphPropertiesTable (PAPBinTable) and
/// CharacterPropertiesTable (CHPBinTable).
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
    /// * `cp_range` - Character position range (start_cp, end_cp)
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
    /// Returns a vector of (text, paragraph_properties, character_runs) tuples.
    ///
    /// Based on MS-DOC specification and Apache POI's approach:
    /// Paragraphs in Word documents are delimited by CR (\r = 0x000D) characters.
    /// The PAP PLCF stores formatting properties, but doesn't define paragraph boundaries.
    pub fn extract_paragraphs(&self) -> Result<Vec<ExtractedParagraph>> {
        let mut paragraphs = Vec::new();

        // Determine the CP range to process
        let doc_start_cp = self.cp_range.map(|(start, _)| start).unwrap_or(0);
        let doc_end_cp = self
            .cp_range
            .map(|(_, end)| end)
            .unwrap_or_else(|| self.text_ranges.len().saturating_sub(1) as u32);

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
            if c == '\r' && next_cp <= doc_end_cp {
                para_boundaries.push(next_cp); // Position after CR
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
            let mut para_text = self.extract_text_range(para_start, para_end);
            // Remove trailing CR if present
            if para_text.ends_with('\r') {
                para_text.pop();
            }

            // Find matching PAP properties for this paragraph
            let para_props = self
                .pap_bin_table
                .and_then(|table| table.properties_at(para_start))
                .cloned()
                .unwrap_or_default();

            // Extract character runs within this paragraph (excluding the CR)
            let para_text_end = if para_end > para_start
                && self.extract_text_range(para_end - 1, para_end) == "\r"
            {
                para_end - 1
            } else {
                para_end
            };
            let runs = self.extract_runs(para_start, para_text_end)?;

            paragraphs.push((para_text, para_props, runs));
        }

        // Fallback if no paragraphs were found
        if paragraphs.is_empty() && !self.text.is_empty() {
            let runs = self.extract_runs(doc_start_cp, doc_end_cp)?;
            paragraphs.push((
                self.text.as_ref().clone(),
                ParagraphProperties::default(),
                runs,
            ));
        }

        Ok(paragraphs)
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

    /// Extract character runs (formatted text segments) within a paragraph.
    fn extract_runs(
        &self,
        para_start: u32,
        para_end: u32,
    ) -> Result<Vec<(String, CharacterProperties)>> {
        let mut runs = Vec::new();

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

                runs.push((run_text, run.properties.clone()));
            }
        }

        // If no runs found, return the whole paragraph as one run
        if runs.is_empty() {
            let para_text = self.extract_text_range(para_start, para_end);
            if !para_text.is_empty() {
                runs.push((para_text, CharacterProperties::default()));
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
mod tests {
    use super::*;

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
}
