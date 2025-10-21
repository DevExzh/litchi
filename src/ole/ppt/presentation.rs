use super::super::OleFile;
/// High-performance Presentation API with zero-copy slide parsing.
use super::package::{PptError, Result};
use super::parsers::PptRecordParser;
use super::persist::PersistMapping;
use super::slide::{Slide, SlideFactory};
use std::io::{Read, Seek};

/// A PowerPoint presentation (.ppt) with high-performance zero-copy parsing.
///
/// # Performance
///
/// - Document data loaded once and borrowed for all slides
/// - Slides parsed lazily using persist mapping
/// - Shapes loaded on-demand
///
/// # Examples
///
/// ```rust,no_run
/// use litchi::ppt::Package;
///
/// let pkg = Package::open("presentation.ppt")?;
/// let pres = pkg.presentation()?;
///
/// // Get slides (zero-copy, lazy evaluation)
/// for slide in pres.slides()? {
///     println!("Slide {}: {}", slide.slide_number(), slide.text()?);
///     println!("  Shapes: {}", slide.shape_count()?);
/// }
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
pub struct Presentation {
    /// The main document stream data (owned for lifetime management)
    powerpoint_document: Vec<u8>,
    /// Parsed record structure (reserved for future advanced parsing)
    #[allow(dead_code)]
    pub(crate) parser: PptRecordParser,
    /// Persist ID to offset mapping
    pub(crate) persist_mapping: PersistMapping,
}

impl Presentation {
    /// Create a new Presentation from an OLE file.
    pub(crate) fn from_ole<R: Read + Seek>(ole: &mut OleFile<R>) -> Result<Self> {
        // Read the PowerPoint Document stream
        let powerpoint_document = Self::read_powerpoint_document(ole)?;

        // Parse document structure
        let mut parser = PptRecordParser::new();
        parser.parse_document(&powerpoint_document)?;

        // Build persist mapping for slide lookup (collect all records recursively)
        let all_records = parser.find_records(crate::ole::consts::PptRecordType::Unknown);
        let persist_mapping = PersistMapping::build_from_records(&all_records);

        Ok(Self {
            powerpoint_document,
            parser,
            persist_mapping,
        })
    }

    /// Read the PowerPoint Document stream from OLE file.
    fn read_powerpoint_document<R: Read + Seek>(ole: &mut OleFile<R>) -> Result<Vec<u8>> {
        // Try primary location
        if let Ok(data) = ole.open_stream(&["PowerPoint Document"]) {
            return Ok(data);
        }

        // Try alternate location
        if let Ok(data) = ole.open_stream(&["PP97_DUALSTORAGE", "PowerPoint Document"]) {
            return Ok(data);
        }

        Err(PptError::InvalidFormat(
            "PowerPoint Document stream not found".to_string(),
        ))
    }

    /// Get iterator over all slides with zero-copy borrowing.
    ///
    /// # Performance
    ///
    /// - Returns lazy iterator (slides parsed on iteration)
    /// - Zero-copy: slides borrow from document data
    /// - Each slide lazily loads its shapes
    pub fn slides(&self) -> Result<Vec<Slide<'_>>> {
        let factory = SlideFactory::new(&self.powerpoint_document, &self.persist_mapping);

        factory
            .slides()
            .enumerate()
            .map(|(idx, slide_result)| {
                slide_result.map(|slide_data| Slide::from_slide_data(slide_data, idx + 1))
            })
            .collect()
    }

    /// Get the number of slides (actual Slide records only).
    #[inline]
    pub fn slide_count(&self) -> usize {
        let factory = SlideFactory::new(&self.powerpoint_document, &self.persist_mapping);
        factory.slide_ids().len()
    }

    /// Extract all text from the presentation.
    ///
    /// # Performance
    ///
    /// - Iterates through all slides
    /// - Each slide extracts text lazily
    /// - Text is collected and joined
    pub fn text(&self) -> Result<String> {
        let slides = self.slides()?;
        let text_parts: Vec<String> = slides
            .iter()
            .filter_map(|slide| slide.text().ok().map(|s| s.to_string()))
            .filter(|text| !text.is_empty())
            .collect();

        Ok(if text_parts.is_empty() {
            String::from("No text content found in presentation")
        } else {
            text_parts.join("\n\n")
        })
    }
}
