/// Headers and footers parser for Word binary format.
///
/// Based on Apache POI's HeaderStories and LibreOffice's implementation.
/// Headers and footers in DOC files are stored as a subdocument with character positions
/// defined in the FIB, and their mapping to sections is defined in a PLCF structure.
use super::super::package::Result;
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
    /// Get all header/footer types in the order they appear in the subdocument
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
            // FIB index 9: fcPlcfHdd and lcbPlcfHdd
            if let Some((offset, length)) = fib.get_table_pointer(9)
                && length > 0
                && (offset as usize) < table_stream.len()
            {
                let plcf_data = &table_stream[offset as usize..];
                let plcf_len = length.min((table_stream.len() - offset as usize) as u32) as usize;

                if plcf_len >= 4 {
                    stories = Self::parse_header_plcf(&plcf_data[..plcf_len], start_cp, end_cp);
                }
            }
        }

        Ok(Self { stories })
    }

    /// Parse the PLCF structure for headers/footers
    ///
    /// The plcfHdd PLCF has element_size = 0 (just character positions).
    /// It contains character positions that divide the header subdocument into stories.
    /// Each section can have up to 6 stories (first page header/footer, even page, odd page).
    fn parse_header_plcf(
        data: &[u8],
        subdoc_start: u32,
        _subdoc_end: u32,
    ) -> Vec<HeaderFooterStory> {
        // Parse as PLCF with element_size = 0 (only CPs, no properties)
        // We need to manually parse this since PlcfParser expects element_size > 0
        if data.len() < 8 {
            return Vec::new();
        }

        // Count of CPs = data.len() / 4
        let cp_count = data.len() / 4;
        if cp_count < 2 {
            return Vec::new();
        }

        let mut cps = Vec::with_capacity(cp_count);
        for i in 0..cp_count {
            if let Ok(cp) = crate::common::binary::read_u32_le(data, i * 4) {
                cps.push(cp);
            } else {
                break;
            }
        }

        // Build stories from consecutive CP pairs
        // Each pair of CPs defines one header/footer story
        let mut stories = Vec::new();
        let header_types = HeaderFooterType::all_types();

        for i in 0..(cps.len() - 1) {
            let start = cps[i];
            let end = cps[i + 1];

            // Convert relative CPs to absolute CPs in the text stream
            let abs_start = subdoc_start + start;
            let abs_end = subdoc_start + end;

            // Determine story type based on position
            // Each section contributes up to 6 stories in order
            let type_index = i % header_types.len();
            let story_type = header_types[type_index];

            stories.push(HeaderFooterStory::new(story_type, abs_start, abs_end));
        }

        stories
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

    #[test]
    fn test_header_footer_type() {
        assert_eq!(HeaderFooterType::all_types().len(), 6);
        assert!(HeaderFooterType::OddPageHeader.is_header());
        assert!(HeaderFooterType::OddPageFooter.is_footer());
    }

    #[test]
    fn test_header_footer_story() {
        let story = HeaderFooterStory::new(HeaderFooterType::OddPageHeader, 100, 200);
        assert_eq!(story.length(), 100);
        assert!(!story.is_empty());

        let empty_story = HeaderFooterStory::new(HeaderFooterType::OddPageFooter, 100, 100);
        assert!(empty_story.is_empty());
    }
}
