use super::super::OleFile;
/// High-performance Presentation API with zero-copy slide parsing.
use super::package::{PptError, Result};
use super::parsers::PptRecordParser;
use super::persist::PersistMapping;
use super::slide::{Slide, SlideFactory};
#[cfg(feature = "imgconv")]
use crate::images::{BlipStore, ExtractedImage, ImageExtractor};
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
    /// Pictures stream data (for image extraction)
    #[cfg(feature = "imgconv")]
    pictures_data: Option<Vec<u8>>,
    /// BLIP store (image metadata index)
    #[cfg(feature = "imgconv")]
    blip_store: Option<BlipStore<'static>>,
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
        // Use zero-copy reference collection to avoid cloning all record data
        let all_records_ref = parser.find_records_ref();
        let persist_mapping = PersistMapping::build_from_records_ref(&all_records_ref);

        // Try to read Pictures stream for image extraction
        #[cfg(feature = "imgconv")]
        let (pictures_data, blip_store) = if let Ok(pictures) = ole.open_stream(&["Pictures"]) {
            // Extract BLIP store from pictures data
            let store = ImageExtractor::extract_blip_store(&pictures)
                .ok()
                .map(|store| store.into_owned()); // Convert to 'static lifetime
            (Some(pictures), store)
        } else {
            (None, None)
        };

        Ok(Self {
            powerpoint_document,
            parser,
            persist_mapping,
            #[cfg(feature = "imgconv")]
            pictures_data,
            #[cfg(feature = "imgconv")]
            blip_store,
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

    /// Fast text extraction that skips shape parsing.
    ///
    /// This is optimized for cases where only text is needed (e.g., markdown conversion)
    /// and shape information is not required.
    ///
    /// # Performance
    ///
    /// - Directly extracts text from slide records without parsing shapes
    /// - Significantly faster than `slides()` + `text()` for large presentations
    /// - No shape object allocation or geometry calculations
    /// - Pre-allocated string buffer
    ///
    /// # Returns
    ///
    /// Vector of (slide_number, text) tuples for each slide
    pub(crate) fn extract_text_fast(&self) -> Result<Vec<(usize, String)>> {
        let factory = SlideFactory::new(&self.powerpoint_document, &self.persist_mapping);

        let mut results = Vec::with_capacity(factory.slide_ids().len());

        for (idx, slide_result) in factory.slides().enumerate() {
            let slide_data = slide_result?;

            // Pre-allocate string buffer
            let mut text = String::with_capacity(512);

            // Extract text from slide records without parsing shapes
            if let Ok(record_text) = slide_data.record.extract_text() {
                let trimmed = record_text.trim();
                if !trimmed.is_empty() {
                    text.push_str(trimmed);
                }
            }

            // Extract text from Escher/PPDrawing using the optimized path
            if let Some(ppdrawing) = slide_data
                .record
                .find_child(crate::ole::consts::PptRecordType::PPDrawing)
                && let Ok(escher_text) = super::escher::extract_text_from_escher(&ppdrawing.data)
            {
                let trimmed = escher_text.trim();
                if !trimmed.is_empty() {
                    if !text.is_empty() {
                        text.push('\n');
                    }
                    text.push_str(trimmed);
                }
            }

            results.push((idx + 1, text));
        }

        Ok(results)
    }

    /// Extract all images from the presentation
    ///
    /// This extracts all embedded images from the Pictures stream.
    ///
    /// # Returns
    /// Vector of all extracted images with metadata
    ///
    /// # Example
    /// ```no_run
    /// use litchi::ole::ppt::Package;
    ///
    /// let mut pkg = Package::open("presentation.ppt")?;
    /// let pres = pkg.presentation()?;
    ///
    /// for image in pres.extract_all_images()? {
    ///     let png_data = image.to_png(None, None)?;
    ///     std::fs::write(image.suggested_filename(), png_data)?;
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    #[cfg(feature = "imgconv")]
    pub fn extract_all_images(&self) -> Result<Vec<ExtractedImage<'static>>> {
        if let Some(ref pictures_data) = self.pictures_data {
            ImageExtractor::extract_from_pictures_stream(pictures_data)
                .map_err(|e| PptError::Corrupted(format!("Failed to extract images: {}", e)))
        } else {
            Ok(Vec::new())
        }
    }

    /// Extract an image by BLIP ID
    ///
    /// This method is used internally by PictureShape to resolve
    /// BLIP ID references to actual image data.
    ///
    /// # Arguments
    /// * `blip_id` - The BLIP ID from the shape's Escher properties
    ///
    /// # Returns
    /// The extracted image, or None if not found
    #[cfg(feature = "imgconv")]
    pub(crate) fn extract_image_by_blip_id(
        &self,
        blip_id: u32,
    ) -> Result<Option<ExtractedImage<'static>>> {
        // Extract all images and find the one matching the BLIP ID
        let images = self.extract_all_images()?;

        // BLIP ID is 1-based index
        let index = (blip_id.saturating_sub(1)) as usize;

        Ok(images.into_iter().nth(index))
    }

    /// Get the BLIP store (image metadata index)
    ///
    /// This provides access to image metadata without extracting the full image data.
    #[cfg(feature = "imgconv")]
    pub fn blip_store(&self) -> Option<&BlipStore<'static>> {
        self.blip_store.as_ref()
    }

    /// Check if the presentation has a Pictures stream
    #[cfg(feature = "imgconv")]
    pub fn has_pictures(&self) -> bool {
        self.pictures_data.is_some()
    }
}
