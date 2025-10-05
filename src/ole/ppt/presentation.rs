/// Presentation - the main API for working with PowerPoint presentation content.
use super::package::{PptError, Result};
use super::slide::Slide;
use super::record_parser::PptRecordParser;
use super::super::OleFile;
use std::fs::File;

/// A PowerPoint presentation (.ppt).
///
/// This is the main API for reading and manipulating legacy PowerPoint presentation content.
/// It provides access to slides, metadata, and other presentation elements.
///
/// # Examples
///
/// ```rust,no_run
/// use litchi::ppt::Package;
///
/// let mut pkg = Package::open("presentation.ppt")?;
/// let pres = pkg.presentation()?;
///
/// // Extract all text
/// let text = pres.text()?;
/// println!("Presentation text: {}", text);
///
/// // Get slide count
/// let count = pres.slide_count()?;
/// println!("Number of slides: {}", count);
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
pub struct Presentation {
    /// The main document stream data
    powerpoint_document: Vec<u8>,
    /// Metadata from the OLE file
    metadata: Option<super::super::OleMetadata>,
}

impl Presentation {
    /// Create a new Presentation from an OLE file.
    ///
    /// This is typically called internally by `Package::presentation()`.
    pub(crate) fn from_ole(ole: &mut OleFile<File>) -> Result<Self> {
        // Read the PowerPoint Document stream (main presentation stream)
        let powerpoint_document = ole
            .open_stream(&["PowerPoint Document"])
            .map_err(|_| PptError::StreamNotFound("PowerPoint Document".to_string()))?;

        // Try to read metadata if available
        let metadata = ole.get_metadata().ok();

        Ok(Self {
            powerpoint_document,
            metadata,
        })
    }

    /// Get all text content from the presentation.
    ///
    /// This extracts all text from all slides in the presentation, concatenated together.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ppt::Package;
    ///
    /// let mut pkg = Package::open("presentation.ppt")?;
    /// let pres = pkg.presentation()?;
    /// let text = pres.text()?;
    /// println!("{}", text);
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn text(&self) -> Result<String> {
        // Parse the PowerPoint document and extract text
        let mut text_parts = Vec::new();

        // Parse the document structure
        let mut parser = PptRecordParser::new();
        parser.parse_document(&self.powerpoint_document)?;

        // Extract text from all slides
        for slide_data in parser.slides() {
            if let Ok(slide_text) = PptRecordParser::extract_text_from_slide_data(slide_data) {
                if !slide_text.is_empty() {
                    text_parts.push(slide_text);
                }
            }
        }

        if text_parts.is_empty() {
            Ok("No text content found in presentation".to_string())
        } else {
            Ok(text_parts.join("\n\n"))
        }
    }

    /// Get the number of slides in the presentation.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ppt::Package;
    ///
    /// let mut pkg = Package::open("presentation.ppt")?;
    /// let pres = pkg.presentation()?;
    /// let count = pres.slide_count()?;
    /// println!("Slides: {}", count);
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn slide_count(&self) -> Result<usize> {
        // Parse the document structure and count slides
        let mut parser = PptRecordParser::new();
        parser.parse_document(&self.powerpoint_document)?;
        Ok(parser.slide_count())
    }

    /// Get all slides in the presentation.
    ///
    /// Returns a vector of `Slide` objects representing slides
    /// in the presentation.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ppt::Package;
    ///
    /// let mut pkg = Package::open("presentation.ppt")?;
    /// let pres = pkg.presentation()?;
    ///
    /// for slide in pres.slides()? {
    ///     println!("Slide: {}", slide.text()?);
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn slides(&self) -> Result<Vec<Slide>> {
        // Parse the document structure and create slide objects
        let mut parser = PptRecordParser::new();
        parser.parse_document(&self.powerpoint_document)?;

        let mut slides = Vec::new();
        for (i, slide_data) in parser.slides().iter().enumerate() {
            // Create a slide with the parsed data
            let slide = Slide::new(slide_data.clone(), i)?;
            slides.push(slide);
        }

        Ok(slides)
    }

    /// Get the presentation's metadata.
    ///
    /// Returns metadata such as title, author, subject, etc. if available.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ppt::Package;
    ///
    /// let mut pkg = Package::open("presentation.ppt")?;
    /// let pres = pkg.presentation()?;
    ///
    /// if let Some(metadata) = pres.metadata()? {
    ///     if let Some(title) = &metadata.title {
    ///         println!("Title: {}", title);
    ///     }
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn metadata(&self) -> Result<Option<&super::super::OleMetadata>> {
        Ok(self.metadata.as_ref())
    }

    /// Get all placeholders across all slides.
    ///
    /// Returns a vector of all placeholders found in the presentation.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ppt::Package;
    ///
    /// let mut pkg = Package::open("presentation.ppt")?;
    /// let pres = pkg.presentation()?;
    ///
    /// for slide in pres.slides()? {
    ///     for placeholder in slide.placeholders() {
    ///         println!("Placeholder: {}", placeholder.placeholder_type());
    ///     }
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn get_all_placeholders(&self) -> Result<Vec<&super::shapes::placeholder::Placeholder>> {
        // For now, return empty vector - this would need to be implemented
        // to collect placeholders from all slides
        Ok(Vec::new())
    }

    /// Get placeholders of a specific type across all slides.
    ///
    /// # Arguments
    ///
    /// * `placeholder_type` - The type of placeholder to find
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ppt::{Package, shapes::PlaceholderType};
    ///
    /// let mut pkg = Package::open("presentation.ppt")?;
    /// let pres = pkg.presentation()?;
    ///
    /// for slide in pres.slides()? {
    ///     for title_placeholder in slide.get_placeholders_by_type(PlaceholderType::Title) {
    ///         println!("Found title placeholder");
    ///     }
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn get_placeholders_by_type(&self, _placeholder_type: super::shapes::placeholder::PlaceholderType) -> Result<Vec<&super::shapes::placeholder::Placeholder>> {
        // For now, return empty vector - this would need to be implemented
        Ok(Vec::new())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ole::OleFile;
    use std::fs::File;

    #[test]
    fn test_create_presentation() {
        let file = File::open("test.ppt").unwrap();
        let mut ole = OleFile::open(file).unwrap();
        let result = Presentation::from_ole(&mut ole);
        assert!(result.is_ok());
    }

    #[test]
    #[ignore] // Requires test file
    fn test_presentation_text() {
        let file = File::open("test.ppt").unwrap();
        let mut ole = OleFile::open(file).unwrap();
        let pres = Presentation::from_ole(&mut ole).unwrap();
        let text = pres.text().unwrap();
        assert!(!text.is_empty());
    }

}
