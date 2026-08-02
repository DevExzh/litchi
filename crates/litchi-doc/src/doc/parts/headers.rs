/// Headers and footers parser for Word binary format.
///
/// Based on Apache POI's HeaderStories and LibreOffice's implementation.
/// Headers and footers in DOC files are stored as a subdocument with character positions
/// defined in the FIB, and their mapping to sections is defined in a PLCF structure.
use super::super::package::{DocError, Result};
use super::fib::FileInformationBlock;

/// Header/Footer types based on section properties
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum HeaderFooterType {
    /// First page header
    FirstPageHeader = 0,
    /// First page footer
    FirstPageFooter = 1,
    /// Even page header
    EvenPageHeader = 2,
    /// Even page footer
    EvenPageFooter = 3,
    /// Odd page header (default)
    OddPageHeader = 4,
    /// Odd page footer (default)
    OddPageFooter = 5,
}

impl HeaderFooterType {
    /// Get all header/footer types in stable semantic order.
    ///
    /// This differs from the interleaved storage order used by `PlcfHdd`.
    pub fn all_types() -> &'static [HeaderFooterType] {
        &[
            HeaderFooterType::FirstPageHeader,
            HeaderFooterType::FirstPageFooter,
            HeaderFooterType::EvenPageHeader,
            HeaderFooterType::EvenPageFooter,
            HeaderFooterType::OddPageHeader,
            HeaderFooterType::OddPageFooter,
        ]
    }

    /// Check if this is a header type
    pub fn is_header(&self) -> bool {
        matches!(
            self,
            HeaderFooterType::FirstPageHeader
                | HeaderFooterType::EvenPageHeader
                | HeaderFooterType::OddPageHeader
        )
    }

    /// Check if this is a footer type
    pub fn is_footer(&self) -> bool {
        !self.is_header()
    }
}

/// A header or footer story (text content)
#[derive(Debug, Clone)]
pub struct HeaderFooterStory {
    /// Type of header/footer
    pub story_type: HeaderFooterType,
    /// Character position range in the header subdocument
    pub start_cp: u32,
    pub end_cp: u32,
}

impl HeaderFooterStory {
    /// Create a new header/footer story
    pub fn new(story_type: HeaderFooterType, start_cp: u32, end_cp: u32) -> Self {
        Self {
            story_type,
            start_cp,
            end_cp,
        }
    }

    /// Get the length in characters
    pub fn length(&self) -> u32 {
        self.end_cp.saturating_sub(self.start_cp)
    }

    /// Check if this story is empty
    pub fn is_empty(&self) -> bool {
        self.length() == 0
    }
}

/// Headers and footers table parser
///
/// Headers/footers are stored in a special subdocument. The FIB contains:
/// - ccpHdd: Character count for header/footer subdocument (at FIB offset 0x54)
/// - plcfHdd: PLCF mapping character positions to header/footer boundaries
pub struct HeadersTable {
    /// All header/footer stories extracted from the subdocument
    stories: Vec<HeaderFooterStory>,
}

impl HeadersTable {
    /// Parse headers/footers from the FIB and table stream
    ///
    /// # Arguments
    ///
    /// * `fib` - File Information Block containing character counts
    /// * `table_stream` - The table stream (0Table or 1Table)
    ///
    /// # Returns
    ///
    /// A parsed HeadersTable or an error
    pub fn parse(fib: &FileInformationBlock, table_stream: &[u8]) -> Result<Self> {
        let mut stories = Vec::new();

        // Check if headers/footers exist
        if let Some((start_cp, end_cp)) = fib.get_header_range() {
            // Get the PLCF for headers/footers (plcfHdd)
            // FIB index 11: fcPlcfHdd and lcbPlcfHdd
            if let Some((offset, length)) = fib.get_table_pointer(11)
                && length > 0
                && (offset as usize) < table_stream.len()
            {
                let plcf_data = &table_stream[offset as usize..];
                let plcf_len = length.min((table_stream.len() - offset as usize) as u32) as usize;

                if plcf_len >= 4 {
                    stories = Self::parse_header_plcf(&plcf_data[..plcf_len], start_cp, end_cp)?;
                }
            }
        }

        Ok(Self { stories })
    }

    /// Parse the PLCF structure for headers/footers
    ///
    /// The plcfHdd PLCF has element_size = 0 (just character positions).
    /// It contains character positions that divide the header subdocument into stories.
    /// The first six stories are footnote/endnote separators. Each section then contributes
    /// six stories in even-header, odd-header, even-footer, odd-footer, first-header,
    /// first-footer order.
    fn parse_header_plcf(
        data: &[u8],
        subdoc_start: u32,
        subdoc_end: u32,
    ) -> Result<Vec<HeaderFooterStory>> {
        // Parse as PLCF with element_size = 0 (only CPs, no properties)
        // We need to manually parse this since PlcfParser expects element_size > 0
        if data.len() < 56 || !data.len().is_multiple_of(4) {
            return Err(DocError::Corrupted(
                "PlcfHdd must contain the six separator and six section story ranges".to_string(),
            ));
        }

        // Count of CPs = data.len() / 4
        let cp_count = data.len() / 4;
        if cp_count < 14 || !(cp_count - 8).is_multiple_of(6) {
            return Err(DocError::Corrupted(
                "PlcfHdd story count is inconsistent with its section groups".to_string(),
            ));
        }

        let mut cps = Vec::with_capacity(cp_count);
        for i in 0..cp_count {
            if let Ok(cp) = litchi_core::binary::read_u32_le(data, i * 4) {
                cps.push(cp);
            } else {
                break;
            }
        }

        let subdoc_len = subdoc_end.checked_sub(subdoc_start).ok_or_else(|| {
            DocError::Corrupted("header subdocument range is reversed".to_string())
        })?;
        for pair in cps[..cps.len() - 1].windows(2) {
            if pair[0] > pair[1] {
                return Err(DocError::Corrupted(
                    "PlcfHdd character positions are not monotonic".to_string(),
                ));
            }
        }
        if subdoc_len == 0 || cps[..cps.len() - 1].iter().any(|&cp| cp >= subdoc_len) {
            return Err(DocError::Corrupted(
                "PlcfHdd character position exceeds the header subdocument".to_string(),
            ));
        }
        if cps[cps.len() - 2] != subdoc_len - 1 {
            return Err(DocError::Corrupted(
                "PlcfHdd story terminator must equal ccpHdd - 1".to_string(),
            ));
        }

        // Build stories from consecutive CP pairs. Slots 0-5 are separator stories,
        // not document headers or footers, so the public table begins at slot 6.
        let mut stories = Vec::new();
        for i in 6..(cps.len() - 2) {
            let start = cps[i];
            let end = cps[i + 1];

            // Convert relative CPs to absolute CPs in the text stream
            let abs_start = subdoc_start.checked_add(start).ok_or_else(|| {
                DocError::Corrupted("PlcfHdd start character position overflows".to_string())
            })?;
            let abs_end = subdoc_start.checked_add(end).ok_or_else(|| {
                DocError::Corrupted("PlcfHdd end character position overflows".to_string())
            })?;

            let story_type = match (i - 6) % 6 {
                0 => HeaderFooterType::EvenPageHeader,
                1 => HeaderFooterType::OddPageHeader,
                2 => HeaderFooterType::EvenPageFooter,
                3 => HeaderFooterType::OddPageFooter,
                4 => HeaderFooterType::FirstPageHeader,
                5 => HeaderFooterType::FirstPageFooter,
                _ => unreachable!(),
            };

            stories.push(HeaderFooterStory::new(story_type, abs_start, abs_end));
        }

        Ok(stories)
    }

    /// Get all header/footer stories
    pub fn stories(&self) -> &[HeaderFooterStory] {
        &self.stories
    }

    /// Get stories of a specific type
    pub fn stories_by_type(&self, story_type: HeaderFooterType) -> Vec<&HeaderFooterStory> {
        self.stories
            .iter()
            .filter(|s| s.story_type == story_type)
            .collect()
    }

    /// Get all header stories
    pub fn headers(&self) -> Vec<&HeaderFooterStory> {
        self.stories
            .iter()
            .filter(|s| s.story_type.is_header())
            .collect()
    }

    /// Get all footer stories
    pub fn footers(&self) -> Vec<&HeaderFooterStory> {
        self.stories
            .iter()
            .filter(|s| s.story_type.is_footer())
            .collect()
    }

    /// Get the total count of header/footer stories
    pub fn count(&self) -> usize {
        self.stories.len()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn encode_cps(cps: &[u32]) -> Vec<u8> {
        cps.iter().flat_map(|cp| cp.to_le_bytes()).collect()
    }

    #[test]
    fn test_header_footer_type() {
        assert_eq!(HeaderFooterType::all_types().len(), 6);
        assert!(HeaderFooterType::OddPageHeader.is_header());
        assert!(HeaderFooterType::OddPageFooter.is_footer());
    }

    #[test]
    fn test_header_footer_type_all_variants() {
        let all = HeaderFooterType::all_types();
        assert_eq!(all.len(), 6);

        // Verify order
        assert_eq!(all[0], HeaderFooterType::FirstPageHeader);
        assert_eq!(all[1], HeaderFooterType::FirstPageFooter);
        assert_eq!(all[2], HeaderFooterType::EvenPageHeader);
        assert_eq!(all[3], HeaderFooterType::EvenPageFooter);
        assert_eq!(all[4], HeaderFooterType::OddPageHeader);
        assert_eq!(all[5], HeaderFooterType::OddPageFooter);
    }

    #[test]
    fn test_header_footer_type_discriminants() {
        assert_eq!(HeaderFooterType::FirstPageHeader as u8, 0);
        assert_eq!(HeaderFooterType::FirstPageFooter as u8, 1);
        assert_eq!(HeaderFooterType::EvenPageHeader as u8, 2);
        assert_eq!(HeaderFooterType::EvenPageFooter as u8, 3);
        assert_eq!(HeaderFooterType::OddPageHeader as u8, 4);
        assert_eq!(HeaderFooterType::OddPageFooter as u8, 5);
    }

    #[test]
    fn test_header_footer_type_is_header() {
        assert!(HeaderFooterType::FirstPageHeader.is_header());
        assert!(HeaderFooterType::EvenPageHeader.is_header());
        assert!(HeaderFooterType::OddPageHeader.is_header());

        assert!(!HeaderFooterType::FirstPageFooter.is_header());
        assert!(!HeaderFooterType::EvenPageFooter.is_header());
        assert!(!HeaderFooterType::OddPageFooter.is_header());
    }

    #[test]
    fn test_header_footer_type_is_footer() {
        assert!(HeaderFooterType::FirstPageFooter.is_footer());
        assert!(HeaderFooterType::EvenPageFooter.is_footer());
        assert!(HeaderFooterType::OddPageFooter.is_footer());

        assert!(!HeaderFooterType::FirstPageHeader.is_footer());
        assert!(!HeaderFooterType::EvenPageHeader.is_footer());
        assert!(!HeaderFooterType::OddPageHeader.is_footer());
    }

    #[test]
    fn test_header_footer_story() {
        let story = HeaderFooterStory::new(HeaderFooterType::OddPageHeader, 100, 200);
        assert_eq!(story.length(), 100);
        assert!(!story.is_empty());

        let empty_story = HeaderFooterStory::new(HeaderFooterType::OddPageFooter, 100, 100);
        assert!(empty_story.is_empty());
    }

    #[test]
    fn test_header_footer_story_new() {
        let story = HeaderFooterStory::new(HeaderFooterType::EvenPageHeader, 50, 150);

        assert_eq!(story.story_type, HeaderFooterType::EvenPageHeader);
        assert_eq!(story.start_cp, 50);
        assert_eq!(story.end_cp, 150);
        assert_eq!(story.length(), 100);
    }

    #[test]
    fn test_header_footer_story_length() {
        let story = HeaderFooterStory::new(HeaderFooterType::FirstPageHeader, 0, 100);
        assert_eq!(story.length(), 100);

        let story2 = HeaderFooterStory::new(HeaderFooterType::FirstPageFooter, 500, 600);
        assert_eq!(story2.length(), 100);

        let empty = HeaderFooterStory::new(HeaderFooterType::OddPageHeader, 100, 100);
        assert_eq!(empty.length(), 0);
    }

    #[test]
    fn test_header_footer_story_is_empty() {
        let empty = HeaderFooterStory::new(HeaderFooterType::OddPageFooter, 100, 100);
        assert!(empty.is_empty());

        let non_empty = HeaderFooterStory::new(HeaderFooterType::OddPageHeader, 100, 101);
        assert!(!non_empty.is_empty());
    }

    #[test]
    fn test_header_footer_story_all_types() {
        let types = [
            HeaderFooterType::FirstPageHeader,
            HeaderFooterType::FirstPageFooter,
            HeaderFooterType::EvenPageHeader,
            HeaderFooterType::EvenPageFooter,
            HeaderFooterType::OddPageHeader,
            HeaderFooterType::OddPageFooter,
        ];

        for (i, story_type) in types.iter().enumerate() {
            let story = HeaderFooterStory::new(*story_type, i as u32 * 100, (i as u32 + 1) * 100);
            assert_eq!(story.story_type, *story_type);
            assert_eq!(story.length(), 100);
        }
    }

    #[test]
    fn test_header_footer_story_large_ranges() {
        let story = HeaderFooterStory::new(HeaderFooterType::OddPageHeader, 0, u32::MAX);
        assert_eq!(story.length(), u32::MAX);
        assert!(!story.is_empty());
    }

    #[test]
    fn test_header_footer_story_zero_length() {
        let story = HeaderFooterStory::new(HeaderFooterType::FirstPageHeader, 1000, 1000);
        assert_eq!(story.length(), 0);
        assert!(story.is_empty());
    }

    #[test]
    fn test_header_footer_story_clone() {
        let story = HeaderFooterStory::new(HeaderFooterType::EvenPageFooter, 200, 300);
        let cloned = story.clone();

        assert_eq!(cloned.story_type, story.story_type);
        assert_eq!(cloned.start_cp, story.start_cp);
        assert_eq!(cloned.end_cp, story.end_cp);
        assert_eq!(cloned.length(), story.length());
    }

    #[test]
    fn test_header_footer_type_debug() {
        let header_type = HeaderFooterType::OddPageHeader;
        let debug_str = format!("{:?}", header_type);
        assert!(debug_str.contains("OddPageHeader") || debug_str.contains("Odd Page Header"));
    }

    #[test]
    fn test_header_footer_story_debug() {
        let story = HeaderFooterStory::new(HeaderFooterType::FirstPageHeader, 100, 200);
        let debug_str = format!("{:?}", story);
        assert!(debug_str.contains("HeaderFooterStory"));
        assert!(debug_str.contains("FirstPageHeader") || debug_str.contains("First Page Header"));
    }

    #[test]
    fn test_header_footer_type_equality() {
        assert_eq!(
            HeaderFooterType::OddPageHeader,
            HeaderFooterType::OddPageHeader
        );
        assert_ne!(
            HeaderFooterType::OddPageHeader,
            HeaderFooterType::OddPageFooter
        );
        assert_ne!(
            HeaderFooterType::FirstPageHeader,
            HeaderFooterType::EvenPageHeader
        );
    }

    #[test]
    fn test_header_footer_type_copy() {
        let header = HeaderFooterType::OddPageHeader;
        let copied = header;
        // After copy, original should still be valid
        assert_eq!(header, HeaderFooterType::OddPageHeader);
        assert_eq!(copied, HeaderFooterType::OddPageHeader);
    }

    #[test]
    fn parses_section_slots_after_separator_stories() {
        let data = encode_cps(&[0, 1, 2, 3, 4, 5, 6, 7, 10, 11, 12, 13, 14, 15]);
        let stories = HeadersTable::parse_header_plcf(&data, 100, 115).unwrap();

        assert_eq!(stories.len(), 6);
        assert_eq!(stories[0].story_type, HeaderFooterType::EvenPageHeader);
        assert_eq!((stories[0].start_cp, stories[0].end_cp), (106, 107));
        assert_eq!(stories[1].story_type, HeaderFooterType::OddPageHeader);
        assert_eq!((stories[1].start_cp, stories[1].end_cp), (107, 110));
        assert_eq!(stories[2].story_type, HeaderFooterType::EvenPageFooter);
        assert_eq!(stories[3].story_type, HeaderFooterType::OddPageFooter);
        assert_eq!(stories[4].story_type, HeaderFooterType::FirstPageHeader);
        assert_eq!(stories[5].story_type, HeaderFooterType::FirstPageFooter);
    }

    #[test]
    fn rejects_malformed_header_character_positions() {
        assert!(HeadersTable::parse_header_plcf(&[0; 7], 0, 10).is_err());
        assert!(HeadersTable::parse_header_plcf(&encode_cps(&[0, 2, 1]), 0, 10).is_err());
        assert!(HeadersTable::parse_header_plcf(&encode_cps(&[0, 11]), 0, 10).is_err());
        assert!(HeadersTable::parse_header_plcf(&encode_cps(&[0, 0, 0]), 10, 9).is_err());
        assert!(HeadersTable::parse_header_plcf(&encode_cps(&[0, 8, 10]), 0, 10).is_err());
    }
}
