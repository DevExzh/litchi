//! Main Presentation structure and implementation.

use super::Slide;
use crate::core::{Content, Meta, OwnedPackage, Styles};
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
/// use litchi_odf::Presentation;
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
    crate::script_package::script_facade_methods!();
    crate::annotation_package::annotation_facade_methods!(Presentation);

    pub(crate) fn package_ref(&self) -> &OwnedPackage {
        &self.package
    }

    pub(crate) fn content_xml(&self) -> &str {
        self.content.xml_content()
    }

    pub(crate) fn into_package(self) -> OwnedPackage {
        self.package
    }

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
    /// use litchi_odf::Presentation;
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
    /// use litchi_odf::Presentation;
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

    /// Discover referenced, inline, missing, and inert linked images.
    pub fn images(&self) -> Result<Vec<crate::Image>> {
        let package = self.package.package()?;
        crate::media::scan_packaged_images(
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            |path| package.has_file(path),
            |path| package.manifest().get_media_type(path).map(str::to_string),
        )
    }

    /// Inspect classic forms without executing bindings, events, or external resources.
    pub fn forms(&self) -> Result<crate::OdfForms> {
        let mut parts = vec![(self.content.xml_content(), crate::OdfFormPart::Content)];
        if let Some(styles) = self.styles.as_ref().map(Styles::xml_content) {
            parts.push((styles, crate::OdfFormPart::Styles));
        }
        crate::form::parse_form_parts(&parts)
    }

    pub fn rdf_graphs(&self) -> Result<Vec<crate::OdfRdfGraph>> {
        crate::rdf_package::graphs(&self.package)
    }
    pub fn add_rdf_graph(
        &mut self,
        preferred_path: Option<&str>,
        triples: &[crate::OdfRdfTriple],
    ) -> Result<String> {
        let (bytes, path) = crate::rdf_package::add_graph(&self.package, preferred_path, triples)?;
        *self = Self::from_bytes(bytes)?;
        Ok(path)
    }
    pub fn replace_rdf_graph(&mut self, path: &str, triples: &[crate::OdfRdfTriple]) -> Result<()> {
        let bytes = crate::rdf_package::replace_graph(&self.package, path, triples)?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }
    pub fn remove_rdf_graph(&mut self, path: &str) -> Result<()> {
        let bytes = crate::rdf_package::remove_graph(&self.package, path)?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }
    pub fn add_rdf_triple(&mut self, path: &str, triple: &crate::OdfRdfTriple) -> Result<usize> {
        let index = self
            .rdf_graphs()?
            .into_iter()
            .find(|graph| graph.path == path)
            .ok_or_else(|| Error::InvalidFormat(format!("RDF graph '{path}' was not found")))?
            .triples
            .len();
        let (bytes, _) = crate::rdf_package::add_triple(&self.package, path, triple)?;
        *self = Self::from_bytes(bytes)?;
        Ok(index)
    }
    pub fn replace_rdf_triple(
        &mut self,
        path: &str,
        index: usize,
        triple: &crate::OdfRdfTriple,
    ) -> Result<()> {
        let bytes = crate::rdf_package::replace_triple(&self.package, path, index, triple)?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }
    pub fn remove_rdf_triple(&mut self, path: &str, index: usize) -> Result<()> {
        let bytes = crate::rdf_package::remove_triple(&self.package, path, index)?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }
    pub fn move_rdf_triple(&mut self, path: &str, from: usize, to: usize) -> Result<()> {
        let bytes = crate::rdf_package::move_triple(&self.package, path, from, to)?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }

    pub fn add_form(&mut self, group_index: usize, form: &crate::OdfAuthoredForm) -> Result<usize> {
        let (bytes, index) = crate::form_package::add_form(
            &self.package,
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            crate::form_package::FormHost::Presentation,
            group_index,
            None,
            form,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(index)
    }
    pub fn add_nested_form(
        &mut self,
        parent_form: usize,
        form: &crate::OdfAuthoredForm,
    ) -> Result<usize> {
        let (bytes, index) = crate::form_package::add_form(
            &self.package,
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            crate::form_package::FormHost::Presentation,
            0,
            Some(parent_form),
            form,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(index)
    }
    pub fn replace_form(&mut self, index: usize, form: &crate::OdfAuthoredForm) -> Result<()> {
        let bytes = crate::form_package::replace_form(
            &self.package,
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            index,
            form,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }
    pub fn remove_form(&mut self, index: usize) -> Result<()> {
        let bytes = crate::form_package::remove_form(
            &self.package,
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            index,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }
    pub fn move_form(&mut self, from: usize, to: usize) -> Result<()> {
        let bytes = crate::form_package::move_form(
            &self.package,
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            from,
            to,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }
    pub fn add_form_control(
        &mut self,
        form_index: usize,
        control: &crate::OdfAuthoredFormControl,
    ) -> Result<usize> {
        let (bytes, index) = crate::form_package::add_control(
            &self.package,
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            form_index,
            control,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(index)
    }
    pub fn replace_form_control(
        &mut self,
        index: usize,
        control: &crate::OdfAuthoredFormControl,
    ) -> Result<()> {
        let bytes = crate::form_package::replace_control(
            &self.package,
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            index,
            control,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }
    pub fn remove_form_control(&mut self, index: usize) -> Result<()> {
        let bytes = crate::form_package::remove_control(
            &self.package,
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            index,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }
    pub fn move_form_control(&mut self, from: usize, to: usize) -> Result<()> {
        let bytes = crate::form_package::move_control(
            &self.package,
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            from,
            to,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }

    /// Inspect ordered ODF variable declarations without evaluating fields or formulas.
    pub fn variable_declarations(&self) -> Result<crate::OdfVariableDeclarations> {
        let mut parts = vec![(self.content.xml_content(), crate::OdfVariablePart::Content)];
        if let Some(styles) = self.styles.as_ref().map(Styles::xml_content) {
            parts.push((styles, crate::OdfVariablePart::Styles));
        }
        crate::variable_declaration::parse_variable_declaration_parts(&parts)
    }

    /// Discover package, inline, missing, and inert linked embedded objects.
    pub fn embedded_objects(&self) -> Result<Vec<crate::OdfEmbeddedObject>> {
        let package = self.package.package()?;
        crate::embedded_object::scan_packaged_objects(
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            |path| package.has_file(path),
            |path| package.manifest().get_media_type(path).map(str::to_string),
        )
    }

    pub fn embedded_chart(&self, index: usize) -> Result<crate::ChartDocument> {
        crate::embedded_chart::open_embedded_chart(
            &self.package,
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            index,
        )
    }

    pub fn add_embedded_chart(
        &mut self,
        page_name: &str,
        definition: &crate::ChartDefinition,
    ) -> Result<usize> {
        self.add_embedded_chart_with_storage(
            page_name,
            definition,
            crate::OdfEmbeddedChartStorage::PackageSubdocument,
        )
    }

    pub fn add_embedded_chart_with_storage(
        &mut self,
        page_name: &str,
        definition: &crate::ChartDefinition,
        storage: crate::OdfEmbeddedChartStorage,
    ) -> Result<usize> {
        let (bytes, index) = crate::embedded_chart::add_embedded_chart(
            &self.package,
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            crate::embedded_chart::EmbeddedChartHost::Page(page_name),
            storage,
            definition,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(index)
    }

    pub fn replace_embedded_chart(
        &mut self,
        index: usize,
        definition: &crate::ChartDefinition,
    ) -> Result<()> {
        let bytes = crate::embedded_chart::replace_embedded_chart(
            &self.package,
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            index,
            definition,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }

    pub fn remove_embedded_chart(&mut self, index: usize) -> Result<()> {
        let bytes = crate::embedded_chart::remove_embedded_chart(
            &self.package,
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            index,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }

    pub fn add_embedded_resource(
        &mut self,
        page_name: &str,
        resource: &crate::OdfEmbeddedResource,
    ) -> Result<usize> {
        let (bytes, index) = crate::embedded_package::add(
            &self.package,
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            crate::embedded_chart::EmbeddedChartHost::Page(page_name),
            resource,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(index)
    }

    pub fn replace_embedded_object(
        &mut self,
        index: usize,
        resource: &crate::OdfEmbeddedResource,
    ) -> Result<()> {
        let bytes = crate::embedded_package::replace(
            &self.package,
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            index,
            crate::embedded_package::ResourceTarget::Object,
            resource,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }

    pub fn replace_embedded_image(
        &mut self,
        index: usize,
        resource: &crate::OdfEmbeddedResource,
    ) -> Result<()> {
        let bytes = crate::embedded_package::replace(
            &self.package,
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            index,
            crate::embedded_package::ResourceTarget::Image,
            resource,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }

    pub fn remove_embedded_object(&mut self, index: usize) -> Result<()> {
        let bytes = crate::embedded_package::remove(
            &self.package,
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            index,
            crate::embedded_package::ResourceTarget::Object,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }

    pub fn remove_embedded_image(&mut self, index: usize) -> Result<()> {
        let bytes = crate::embedded_package::remove(
            &self.package,
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            index,
            crate::embedded_package::ResourceTarget::Image,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }

    pub fn move_embedded_object(&mut self, from: usize, to: usize) -> Result<()> {
        let bytes = crate::embedded_package::reorder(
            &self.package,
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            from,
            to,
            crate::embedded_package::ResourceTarget::Object,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }

    pub fn move_embedded_image(&mut self, from: usize, to: usize) -> Result<()> {
        let bytes = crate::embedded_package::reorder(
            &self.package,
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            from,
            to,
            crate::embedded_package::ResourceTarget::Image,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }

    /// Return bytes only for inline or verified package-contained images.
    /// Linked images remain inert and are never fetched.
    pub fn image_bytes(&self, image: &crate::Image) -> Result<Option<Vec<u8>>> {
        match &image.source {
            crate::ImageSource::Inline { bytes, .. } => Ok(Some(bytes.clone())),
            crate::ImageSource::PackagePart { path, .. } => self.package.get_file(path).map(Some),
            _ => Ok(None),
        }
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

        let package = self.package.package()?;
        let content_bytes = package.get_file("content.xml")?;
        let content = Content::from_bytes(&content_bytes)?;

        OdpParser::parse_slides_with_styles(
            content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
        )
    }

    /// Inspect inert slide-show settings and ordered custom shows.
    pub fn settings(&self) -> Result<Option<super::PresentationSettings>> {
        super::parse_presentation_settings(self.content.xml_content())
    }

    /// Inspect inert header, footer, date-time, and page-binding declarations.
    pub fn declarations(&self) -> Result<super::PresentationDeclarations> {
        super::parse_presentation_declarations(self.content.xml_content())
    }

    /// Inspect static page names, IDs, and layout/master references.
    pub fn page_metadata(&self) -> Result<super::PresentationPageMetadataCollection> {
        super::parse_presentation_page_metadata(self.content.xml_content())
    }

    /// Inspect named presentation page layouts and their typed placeholders.
    pub fn page_layouts(&self) -> Result<super::PresentationPageLayouts> {
        match self.styles.as_ref() {
            Some(styles) => super::parse_presentation_page_layouts(styles.xml_content()),
            None => Ok(super::PresentationPageLayouts::default()),
        }
    }

    /// Inspect named legacy and SVG drawing gradients without resolving style use sites.
    pub fn drawing_gradients(&self) -> Result<crate::drawing_gradient::OdfDrawingGradients> {
        match self.styles.as_ref() {
            Some(styles) => crate::drawing_gradient::parse_drawing_gradients(styles.xml_content()),
            None => Ok(crate::drawing_gradient::OdfDrawingGradients::default()),
        }
    }

    /// Inspect named drawing hatch resources without resolving style use sites.
    pub fn drawing_hatches(&self) -> Result<crate::drawing_hatch::OdfDrawingHatches> {
        match self.styles.as_ref() {
            Some(styles) => crate::drawing_hatch::parse_drawing_hatches(styles.xml_content()),
            None => Ok(crate::drawing_hatch::OdfDrawingHatches::default()),
        }
    }

    /// Inspect named fill-image definitions without resolving style use sites.
    ///
    /// Links remain stored metadata: this does not follow them, load linked
    /// resources, or render images.
    pub fn drawing_fill_images(&self) -> Result<crate::drawing_fill_image::OdfDrawingFillImages> {
        match self.styles.as_ref() {
            Some(styles) => {
                crate::drawing_fill_image::parse_drawing_fill_images(styles.xml_content())
            },
            None => Ok(crate::drawing_fill_image::OdfDrawingFillImages::default()),
        }
    }

    /// Inspect named drawing marker definitions without resolving style use sites.
    ///
    /// This does not render marker paths.
    pub fn drawing_markers(&self) -> Result<crate::drawing_marker::OdfDrawingMarkers> {
        match self.styles.as_ref() {
            Some(styles) => crate::drawing_marker::parse_drawing_markers(styles.xml_content()),
            None => Ok(crate::drawing_marker::OdfDrawingMarkers::default()),
        }
    }

    /// Inspect named drawing opacity definitions without resolving style use sites.
    ///
    /// This does not render opacity gradients.
    pub fn drawing_opacities(&self) -> Result<crate::drawing_opacity::OdfDrawingOpacities> {
        match self.styles.as_ref() {
            Some(styles) => crate::drawing_opacity::parse_drawing_opacities(styles.xml_content()),
            None => Ok(crate::drawing_opacity::OdfDrawingOpacities::default()),
        }
    }

    /// Inspect named drawing stroke-dash definitions without resolving style use sites.
    ///
    /// This does not render strokes.
    pub fn drawing_stroke_dashes(
        &self,
    ) -> Result<crate::drawing_stroke_dash::OdfDrawingStrokeDashes> {
        match self.styles.as_ref() {
            Some(styles) => {
                crate::drawing_stroke_dash::parse_drawing_stroke_dashes(styles.xml_content())
            },
            None => Ok(crate::drawing_stroke_dash::OdfDrawingStrokeDashes::default()),
        }
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

    /// Read a package-contained media payload without fetching external URLs.
    ///
    /// Returns `None` for external links, fragment links, unsafe paths, and
    /// package-relative references whose payload is absent.
    pub fn media_data(&self, media: &super::MediaReference) -> Result<Option<Vec<u8>>> {
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

    /// Get the complete format-specific OpenDocument metadata model.
    pub fn odf_metadata(&self) -> Result<Option<crate::OdfMetadata>> {
        self.meta.as_ref().map(Meta::odf_metadata).transpose()
    }

    pub(crate) fn styles_xml(&self) -> Option<&str> {
        self.styles.as_ref().map(Styles::xml_content)
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
    /// use litchi_odf::Presentation;
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
    /// use litchi_odf::Presentation;
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
