//! Main Presentation structure and implementation.

use super::Slide;
use crate::common::{Error, Metadata, Result};
use crate::odf::core::{Content, Meta, Package, Styles};
use std::io::Cursor;
use std::path::Path;

/// An OpenDocument presentation (.odp).
///
/// This struct represents a complete ODP presentation and provides methods to access
/// its slides and metadata.
///
/// # Examples
///
/// ```no_run
/// use litchi::odf::Presentation;
///
/// # fn main() -> litchi::Result<()> {
/// let mut presentation = Presentation::open("slides.odp")?;
///
/// // Get slide count
/// println!("Slides: {}", presentation.slide_count()?);
///
/// // Access slides
/// let slides = presentation.slides()?;
/// for slide in slides {
///     println!("Slide {}: {}", slide.index() + 1, slide.text()?);
/// }
/// # Ok(())
/// # }
/// ```
pub struct Presentation {
    package: Package<Cursor<Vec<u8>>>,
    #[allow(dead_code)]
    content: Content,
    #[allow(dead_code)]
    styles: Option<Styles>,
    meta: Option<Meta>,
}

impl Presentation {
    /// Open an ODP presentation from a file path.
    ///
    /// # Arguments
    ///
    /// * `path` - Path to the .odp file
    ///
    /// # Errors
    ///
    /// Returns an error if the file cannot be read or is not a valid ODP file.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi::odf::Presentation;
    ///
    /// # fn main() -> litchi::Result<()> {
    /// let presentation = Presentation::open("slides.odp")?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self> {
        let bytes = std::fs::read(path.as_ref())?;
        Self::from_bytes(bytes)
    }

    /// Create a Presentation from a byte buffer.
    ///
    /// # Arguments
    ///
    /// * `bytes` - Complete ODP file contents as bytes
    ///
    /// # Errors
    ///
    /// Returns an error if the bytes do not represent a valid ODP file.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi::odf::Presentation;
    ///
    /// # fn main() -> litchi::Result<()> {
    /// let bytes = std::fs::read("slides.odp")?;
    /// let presentation = Presentation::from_bytes(bytes)?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        let cursor = Cursor::new(bytes);
        let package = Package::from_reader(cursor)?;

        // Verify this is a presentation
        let mime_type = package.mimetype();
        if !mime_type.contains("opendocument.presentation") {
            return Err(Error::InvalidFormat(format!(
                "Not an ODP file: MIME type is {}",
                mime_type
            )));
        }

        // Parse core components
        let content_bytes = package.get_file("content.xml")?;
        let content = Content::from_bytes(&content_bytes)?;

        let styles = if package.has_file("styles.xml") {
            let styles_bytes = package.get_file("styles.xml")?;
            Some(Styles::from_bytes(&styles_bytes)?)
        } else {
            None
        };

        let meta = if package.has_file("meta.xml") {
            let meta_bytes = package.get_file("meta.xml")?;
            Some(Meta::from_bytes(&meta_bytes)?)
        } else {
            None
        };

        Ok(Self {
            package,
            content,
            styles,
            meta,
        })
    }

    /// Create an ODP presentation from an already-parsed ZIP archive.
    ///
    /// This is used for single-pass parsing where the ZIP archive has already
    /// been parsed during format detection. It avoids double-parsing.
    pub fn from_zip_archive(
        zip_archive: zip::ZipArchive<std::io::Cursor<Vec<u8>>>,
    ) -> Result<Self> {
        let package = Package::from_zip_archive(zip_archive)?;

        // Verify this is a presentation
        let mime_type = package.mimetype();
        if !mime_type.contains("opendocument.presentation") {
            return Err(Error::InvalidFormat(format!(
                "Not an ODP file: MIME type is {}",
                mime_type
            )));
        }

        // Parse core components
        let content_bytes = package.get_file("content.xml")?;
        let content = Content::from_bytes(&content_bytes)?;

        let styles = if package.has_file("styles.xml") {
            let styles_bytes = package.get_file("styles.xml")?;
            Some(Styles::from_bytes(&styles_bytes)?)
        } else {
            None
        };

        let meta = if package.has_file("meta.xml") {
            let meta_bytes = package.get_file("meta.xml")?;
            Some(Meta::from_bytes(&meta_bytes)?)
        } else {
            None
        };

        Ok(Self {
            package,
            content,
            styles,
            meta,
        })
    }

    /// Get the number of slides in the presentation.
    pub fn slide_count(&self) -> Result<usize> {
        let slides = self.slides()?;
        Ok(slides.len())
    }

    /// Get all slides in the presentation.
    ///
    /// Returns a vector of `Slide` objects representing all slides in the document.
    pub fn slides(&self) -> Result<Vec<Slide>> {
        use super::parser::OdpParser;

        let content_bytes = self.package.get_file("content.xml")?;
        let content = Content::from_bytes(&content_bytes)?;

        OdpParser::parse_slides(content.xml_content())
    }

    /// Get a slide by index.
    ///
    /// Returns `Some(slide)` if a slide exists at the given index, `None` otherwise.
    ///
    /// # Arguments
    ///
    /// * `index` - 0-based index of the slide
    pub fn slide(&self, index: usize) -> Result<Option<Slide>> {
        let slides = self.slides()?;
        Ok(slides.into_iter().nth(index))
    }

    /// Extract all text content from the presentation.
    ///
    /// Returns text from all slides, separated by double newlines.
    pub fn text(&self) -> Result<String> {
        let slides = self.slides()?;
        let mut all_text = Vec::new();

        for slide in slides {
            if !slide.text.trim().is_empty() {
                all_text.push(slide.text.trim().to_string());
            }
        }

        Ok(all_text.join("\n\n"))
    }

    /// Get document metadata.
    ///
    /// Extracts metadata from the meta.xml file.
    pub fn metadata(&self) -> Result<Metadata> {
        if let Some(meta) = &self.meta {
            Ok(meta.extract_metadata())
        } else {
            Ok(Metadata::default())
        }
    }

    // Note: For presentation modification operations, see `MutablePresentation` which provides
    // full CRUD operations on slides and shapes including add/remove/update slides, add/remove
    // shapes, and clear operations.

    /// Save the presentation to a new file.
    ///
    /// This method saves the current presentation state to a new file.
    ///
    /// # Arguments
    ///
    /// * `path` - Path where the ODP file should be saved
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi::odf::Presentation;
    ///
    /// # fn main() -> litchi::Result<()> {
    /// let presentation = Presentation::open("input.odp")?;
    /// presentation.save("output.odp")?;
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// # Note
    ///
    /// Full presentation modification support is planned for future releases. For now,
    /// to modify a presentation, use `PresentationBuilder` to create a new one with
    /// the desired content.
    pub fn save<P: AsRef<std::path::Path>>(&self, path: P) -> Result<()> {
        let bytes = self.to_bytes()?;
        std::fs::write(path, bytes)?;
        Ok(())
    }

    /// Convert the presentation to bytes.
    ///
    /// This method serializes the presentation to an ODF-compliant ZIP archive.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi::odf::Presentation;
    ///
    /// # fn main() -> litchi::Result<()> {
    /// let presentation = Presentation::open("slides.odp")?;
    /// let bytes = presentation.to_bytes()?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        use crate::odf::core::PackageWriter;

        let mut writer = PackageWriter::new();

        // Set MIME type
        writer.set_mimetype(self.package.mimetype())?;

        // Add content.xml
        let content_xml = self.content.xml_content();
        writer.add_file("content.xml", content_xml.as_bytes())?;

        // Add styles.xml if present
        if let Some(ref styles) = self.styles {
            let styles_xml = styles.xml_content();
            writer.add_file("styles.xml", styles_xml.as_bytes())?;
        }

        // Add meta.xml if present
        if let Some(ref meta) = self.meta {
            let meta_xml = meta.xml_content();
            writer.add_file("meta.xml", meta_xml.as_bytes())?;
        }

        // Copy settings.xml if present
        if self.package.has_file("settings.xml") {
            let settings_bytes = self.package.get_file("settings.xml")?;
            writer.add_file("settings.xml", &settings_bytes)?;
        }

        // Copy all media files (images, videos, etc.) from the original package
        let media_files = self.package.media_files()?;
        for media_path in media_files {
            if let Ok(media_bytes) = self.package.get_file(&media_path) {
                writer.add_file(&media_path, &media_bytes)?;
            }
        }

        // Copy other common ODF files if they exist
        let other_files = vec!["Thumbnails/thumbnail.png", "Configurations2/"];
        for file_path in other_files {
            if self.package.has_file(file_path)
                && let Ok(file_bytes) = self.package.get_file(file_path)
            {
                writer.add_file(file_path, &file_bytes)?;
            }
        }

        writer.finish_to_bytes()
    }

    // Note: DELETE operations are available via `MutablePresentation`. To modify this presentation:
    //   1. Convert: `let mut mutable = MutablePresentation::from_presentation(presentation)?`
    //   2. Modify: `mutable.remove_slide(0)?`, `mutable.add_shape(0, shape)?`, etc.
    //   3. Save: `mutable.save("output.odp")?`
    // Available methods: remove_slide, remove_shape, update_slide, clear_slide, clear_slides, etc.
}
