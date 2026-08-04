//! Main Presentation structure and implementation.

use crate::codec::Parser;
use crate::core::{Content, Meta, OwnedPackage, Styles};
use crate::model::{
    MediaReference, PresentationDeclarations, PresentationPageLayouts,
    PresentationPageMetadataCollection, PresentationSettings, Slide,
    parse_presentation_declarations, parse_presentation_page_layouts,
    parse_presentation_page_metadata, parse_presentation_settings,
};
use litchi_core::{Error, Metadata, Result};
use std::path::Path;

/// An OpenDocument presentation (.odp).
///
/// This struct represents a complete ODP presentation and provides methods to access
/// its slides and metadata.
///
/// # Examples
///
/// ```no_run
/// use litchi_odp::Presentation;
///
/// # fn main() -> litchi_core::Result<()> {
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
    package: OwnedPackage,
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
    /// use litchi_odp::Presentation;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let presentation = Presentation::open("slides.odp")?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self> {
        let bytes = std::fs::read(path.as_ref())?;
        Self::from_bytes(bytes)
    }

    /// Open a password-encrypted ODP presentation.
    pub fn open_with_password<P: AsRef<Path>>(
        path: P,
        password: impl Into<String>,
    ) -> Result<Self> {
        let bytes = std::fs::read(path.as_ref())?;
        Self::from_bytes_with_password(bytes, password)
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
    /// use litchi_odp::Presentation;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let bytes = std::fs::read("slides.odp")?;
    /// let presentation = Presentation::from_bytes(bytes)?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        let owned_package = OwnedPackage::from_bytes(bytes)?;
        Self::from_owned_package(owned_package)
    }

    /// Create a presentation from password-encrypted ODP bytes.
    pub fn from_bytes_with_password(bytes: Vec<u8>, password: impl Into<String>) -> Result<Self> {
        Self::from_owned_package(OwnedPackage::from_bytes_with_password(bytes, password)?)
    }

    fn from_owned_package(owned_package: OwnedPackage) -> Result<Self> {
        let package = owned_package.package()?;

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
            package: owned_package,
            content,
            styles,
            meta,
        })
    }

    /// Create an ODP presentation from raw bytes (ZIP archive data).
    ///
    /// This is used for single-pass parsing where the ZIP archive has already
    /// been validated during format detection. It avoids double-parsing.
    pub fn from_archive_bytes(bytes: Vec<u8>) -> Result<Self> {
        Self::from_bytes(bytes)
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
        let package = self.package.package()?;
        let content_bytes = package.get_file("content.xml")?;
        let content = Content::from_bytes(&content_bytes)?;

        Parser::parse_slides_with_styles(
            content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
        )
    }

    /// Inspect inert slide-show settings and ordered custom shows.
    pub fn settings(&self) -> Result<Option<PresentationSettings>> {
        parse_presentation_settings(self.content.xml_content())
    }

    /// Inspect inert header, footer, date-time, and page-binding declarations.
    pub fn declarations(&self) -> Result<PresentationDeclarations> {
        parse_presentation_declarations(self.content.xml_content())
    }

    /// Inspect static page names, IDs, and layout/master references.
    pub fn page_metadata(&self) -> Result<PresentationPageMetadataCollection> {
        parse_presentation_page_metadata(self.content.xml_content())
    }

    /// Inspect named presentation page layouts and their typed placeholders.
    pub fn page_layouts(&self) -> Result<PresentationPageLayouts> {
        match self.styles.as_ref() {
            Some(styles) => parse_presentation_page_layouts(styles.xml_content()),
            None => Ok(PresentationPageLayouts::default()),
        }
    }

    /// Inspect named legacy and SVG drawing gradients without resolving style use sites.
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

    /// Read a package-contained media payload without fetching external URLs.
    ///
    /// Returns `None` for external links, fragment links, unsafe paths, and
    /// package-relative references whose payload is absent.
    pub fn media_data(&self, media: &MediaReference) -> Result<Option<Vec<u8>>> {
        let Some(path) = media.package_path() else {
            return Ok(None);
        };
        let package = self.package.package()?;
        if !package.has_file(path) {
            return Ok(None);
        }
        package.get_file(path).map(Some)
    }

    /// Extract all text content from the presentation.
    ///
    /// Returns text from all slides, separated by double newlines.
    pub fn text(&self) -> Result<String> {
        let slides = self.slides()?;
        let mut all_text = Vec::new();

        for slide in slides {
            let text = slide.all_text();
            if !text.is_empty() {
                all_text.push(text);
            }
        }

        Ok(all_text.join("\n\n"))
    }

    /// Get document metadata.
    ///
    /// Extracts metadata from the meta.xml file.
    pub fn metadata(&self) -> Result<Metadata> {
        if let Some(meta) = &self.meta {
            meta.try_extract_metadata()
        } else {
            Ok(Metadata::default())
        }
    }

    /// Read all inert RDF metadata graphs in package order.
    pub fn rdf_graphs(&self) -> Result<Vec<crate::rdf::Graph>> {
        crate::rdf::graphs(&self.package)
    }

    /// Add a graph and atomically replace this snapshot with the rebuilt package.
    pub fn add_rdf_graph(
        &mut self,
        preferred_path: Option<&str>,
        triples: &[crate::rdf::Triple],
    ) -> Result<String> {
        let (bytes, path) = crate::rdf::add_graph(&self.package, preferred_path, triples)?;
        *self = Self::from_bytes(bytes)?;
        Ok(path)
    }

    /// Replace one complete RDF graph and atomically publish the result.
    pub fn replace_rdf_graph(&mut self, path: &str, triples: &[crate::rdf::Triple]) -> Result<()> {
        let bytes = crate::rdf::replace_graph(&self.package, path, triples)?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }

    /// Remove one RDF graph after validating that no remaining graph references it.
    pub fn remove_rdf_graph(&mut self, path: &str) -> Result<()> {
        let bytes = crate::rdf::remove_graph(&self.package, path)?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }

    /// Append one triple to an existing graph and return its committed index.
    pub fn add_rdf_triple(&mut self, path: &str, triple: &crate::rdf::Triple) -> Result<usize> {
        let index = self
            .rdf_graphs()?
            .into_iter()
            .find(|graph| graph.path == path)
            .ok_or_else(|| Error::InvalidFormat(format!("RDF graph '{path}' was not found")))?
            .triples
            .len();
        let (bytes, _) = crate::rdf::add_triple(&self.package, path, triple)?;
        *self = Self::from_bytes(bytes)?;
        Ok(index)
    }

    /// Replace one triple while preserving its description subject.
    pub fn replace_rdf_triple(
        &mut self,
        path: &str,
        index: usize,
        triple: &crate::rdf::Triple,
    ) -> Result<()> {
        let bytes = crate::rdf::replace_triple(&self.package, path, index, triple)?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }

    /// Remove one triple from a graph.
    pub fn remove_rdf_triple(&mut self, path: &str, index: usize) -> Result<()> {
        let bytes = crate::rdf::remove_triple(&self.package, path, index)?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }

    /// Move one triple within its RDF description.
    pub fn move_rdf_triple(&mut self, path: &str, from: usize, to: usize) -> Result<()> {
        let bytes = crate::rdf::move_triple(&self.package, path, from, to)?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }

    /// Get the complete format-specific OpenDocument metadata model.
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
    /// use litchi_odp::Presentation;
    ///
    /// # fn main() -> litchi_core::Result<()> {
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
    pub fn save<P: AsRef<Path>>(&self, path: P) -> Result<()> {
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
    /// use litchi_odp::Presentation;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let presentation = Presentation::open("slides.odp")?;
    /// let bytes = presentation.to_bytes()?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        Ok(self.package.as_bytes().to_vec())
    }

    // Note: DELETE operations are available via `MutablePresentation`. To modify this presentation:
    //   1. Convert: `let mut mutable = MutablePresentation::from_presentation(presentation)?`
    //   2. Modify: `mutable.remove_slide(0)?`, `mutable.add_shape(0, shape)?`, etc.
    //   3. Save: `mutable.save("output.odp")?`
    // Available methods: remove_slide, remove_shape, update_slide, clear_slide, clear_slides, etc.
}
