//! ODT package lifecycle, mutation, and lossless byte access.

use super::model::Document;
use crate::core::{Content, Meta, OwnedPackage, Styles};
use crate::elements::style::{StyleElements, StyleRegistry};
use litchi_core::{Error, Result};
use std::path::Path;

impl Document {
    pub(crate) fn into_package(self) -> OwnedPackage {
        take(self)
    }

    /// Open an ODT document from a file path.
    ///
    /// This method reads the entire file into memory and parses it. For large files,
    /// consider using `from_bytes` with a streaming reader if memory is constrained.
    ///
    /// # Arguments
    ///
    /// * `path` - Path to the .odt file
    ///
    /// # Errors
    ///
    /// Returns an error if:
    /// - The file cannot be read
    /// - The file is not a valid ZIP archive
    /// - The file is not a valid ODT document
    /// - Required XML components are malformed
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odt::Document;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let doc = Document::open("my_document.odt")?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self> {
        let bytes = std::fs::read(path.as_ref())?;
        Self::from_bytes(bytes)
    }

    /// Open a password-encrypted ODT document.
    pub fn open_with_password<P: AsRef<Path>>(
        path: P,
        password: impl Into<String>,
    ) -> Result<Self> {
        let bytes = std::fs::read(path.as_ref())?;
        Self::from_bytes_with_password(bytes, password)
    }

    /// Create a Document from a byte buffer.
    ///
    /// This is useful when you have the document data in memory already,
    /// such as from network transfers or embedded resources.
    ///
    /// # Arguments
    ///
    /// * `bytes` - Complete ODT file contents as bytes
    ///
    /// # Errors
    ///
    /// Returns an error if the bytes do not represent a valid ODT document.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odt::Document;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let bytes = std::fs::read("document.odt")?;
    /// let doc = Document::from_bytes(bytes)?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        let owned_package = OwnedPackage::from_bytes(bytes)?;
        Self::from_owned_package(owned_package)
    }

    /// Create a document from password-encrypted ODT bytes.
    pub fn from_bytes_with_password(bytes: Vec<u8>, password: impl Into<String>) -> Result<Self> {
        Self::from_owned_package(OwnedPackage::from_bytes_with_password(bytes, password)?)
    }

    pub(crate) fn from_owned_package(owned_package: OwnedPackage) -> Result<Self> {
        let package = owned_package.package()?;

        // Verify this is a text document.
        validate_mimetype(package.mimetype())?;

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

        // Initialize style registry
        let mut style_registry = StyleRegistry::default();

        // Parse styles from styles.xml if available
        if let Some(ref styles_part) = styles
            && let Ok(registry) = StyleElements::parse_styles(styles_part.xml_content())
        {
            style_registry = registry;
        }

        // Also parse styles from content.xml (automatic styles)
        if let Ok(content_registry) = StyleElements::parse_styles(content.xml_content()) {
            // Merge content styles into main registry (content styles take precedence)
            for (_name, style) in content_registry.styles {
                style_registry.add_style(style);
            }
        }

        Ok(Self {
            package: owned_package,
            content,
            styles,
            meta,
            style_registry,
        })
    }

    pub fn original_bytes(&self) -> &[u8] {
        original_bytes(self)
    }

    /// Create an ODT document from raw bytes (ZIP archive data).
    ///
    /// This is used for single-pass parsing where the ZIP archive has already
    /// been validated during format detection. It avoids double-parsing.
    pub fn from_archive_bytes(bytes: Vec<u8>) -> Result<Self> {
        Self::from_bytes(bytes)
    }

    // Note: For document modification operations, see `MutableDocument` which provides
    // full CRUD operations (Create, Read, Update, Delete) on document content including
    // adding, updating, and removing paragraphs and tables while preserving insertion order.

    /// Save the document to a new file.
    ///
    /// This method saves the current document state to a new file. Note that this
    /// creates a copy of the original document; modifications are not yet supported.
    ///
    /// # Arguments
    ///
    /// * `path` - Path where the ODT file should be saved
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odt::Document;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let doc = Document::open("input.odt")?;
    /// doc.save("output.odt")?;
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// # Note
    ///
    /// Full document modification support is planned for future releases. For now,
    /// to modify a document, use `Builder` to create a new document with
    /// the desired content.
    pub fn save<P: AsRef<Path>>(&self, path: P) -> Result<()> {
        let bytes = self.to_bytes()?;
        std::fs::write(path, bytes)?;
        Ok(())
    }

    /// Convert the document to bytes.
    ///
    /// This method serializes the document to an ODF-compliant ZIP archive.
    /// All embedded media files (images, etc.) are automatically copied.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odt::Document;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let doc = Document::open("document.odt")?;
    /// let bytes = doc.to_bytes()?;
    /// // Use bytes for network transfer, etc.
    /// # Ok(())
    /// # }
    /// ```
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        to_bytes(self)
    }

    /// Replace document interaction and protection policy metadata.
    ///
    /// Only `settings.xml` is changed; all other package parts remain under
    /// the package writer's lossless auxiliary-copy policy.
    pub fn set_protection(&mut self, policy: &crate::protection::Policy) -> Result<()> {
        let before = self.protection()?;
        if &before == policy {
            return Ok(());
        }
        let mimetype = self.package.mimetype()?;
        let bytes = crate::protection::rewrite_owned_package(&self.package, &mimetype, policy)?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }

    pub fn add_rdf_graph(
        &mut self,
        preferred_path: Option<&str>,
        triples: &[crate::rdf::Triple],
    ) -> Result<String> {
        let (bytes, path) = crate::rdf::add_graph(&self.package, preferred_path, triples)?;
        *self = Self::from_bytes(bytes)?;
        Ok(path)
    }

    pub fn replace_rdf_graph(&mut self, path: &str, triples: &[crate::rdf::Triple]) -> Result<()> {
        let bytes = crate::rdf::replace_graph(&self.package, path, triples)?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }

    pub fn remove_rdf_graph(&mut self, path: &str) -> Result<()> {
        let bytes = crate::rdf::remove_graph(&self.package, path)?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }

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

    pub fn remove_rdf_triple(&mut self, path: &str, index: usize) -> Result<()> {
        let bytes = crate::rdf::remove_triple(&self.package, path, index)?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }

    pub fn move_rdf_triple(&mut self, path: &str, from: usize, to: usize) -> Result<()> {
        let bytes = crate::rdf::move_triple(&self.package, path, from, to)?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }

    pub fn add_form(
        &mut self,
        group_index: usize,
        form: &crate::package::forms::AuthoredForm,
    ) -> Result<usize> {
        let (bytes, index) = crate::package::forms::add_form(
            &self.package,
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            crate::package::forms::FormHost::Text,
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
        form: &crate::package::forms::AuthoredForm,
    ) -> Result<usize> {
        let (bytes, index) = crate::package::forms::add_form(
            &self.package,
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            crate::package::forms::FormHost::Text,
            0,
            Some(parent_form),
            form,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(index)
    }

    pub fn replace_form(
        &mut self,
        index: usize,
        form: &crate::package::forms::AuthoredForm,
    ) -> Result<()> {
        let bytes = crate::package::forms::replace_form(
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
        let bytes = crate::package::forms::remove_form(
            &self.package,
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            index,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }

    pub fn move_form(&mut self, from: usize, to: usize) -> Result<()> {
        let bytes = crate::package::forms::move_form(
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
        control: &crate::package::forms::AuthoredFormControl,
    ) -> Result<usize> {
        let (bytes, index) = crate::package::forms::add_control(
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
        control: &crate::package::forms::AuthoredFormControl,
    ) -> Result<()> {
        let bytes = crate::package::forms::replace_control(
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
        let bytes = crate::package::forms::remove_control(
            &self.package,
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            index,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }

    pub fn move_form_control(&mut self, from: usize, to: usize) -> Result<()> {
        let bytes = crate::package::forms::move_control(
            &self.package,
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            from,
            to,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }

    /// Append a packaged chart object to the text body.
    pub fn add_embedded_chart(&mut self, definition: &crate::odc::Definition) -> Result<usize> {
        self.add_embedded_chart_with_storage(
            definition,
            crate::EmbeddedChartStorage::PackageSubdocument,
        )
    }

    /// Append a chart object using an explicit storage form.
    pub fn add_embedded_chart_with_storage(
        &mut self,
        definition: &crate::odc::Definition,
        storage: crate::EmbeddedChartStorage,
    ) -> Result<usize> {
        let (bytes, index) = crate::package::charts::add_embedded_chart(
            &self.package,
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            crate::package::charts::EmbeddedChartHost::Text,
            storage,
            definition,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(index)
    }

    pub fn replace_embedded_chart(
        &mut self,
        index: usize,
        definition: &crate::odc::Definition,
    ) -> Result<()> {
        let bytes = crate::package::charts::replace_embedded_chart(
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
        let bytes = crate::package::charts::remove_embedded_chart(
            &self.package,
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            index,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }

    /// Append an inert embedded object or image to the text body.
    pub fn add_embedded_resource(&mut self, resource: &crate::EmbeddedResource) -> Result<usize> {
        let (bytes, index) = crate::package::embedded::add(
            &self.package,
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            crate::package::charts::EmbeddedChartHost::Text,
            resource,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(index)
    }

    pub fn replace_embedded_object(
        &mut self,
        index: usize,
        resource: &crate::EmbeddedResource,
    ) -> Result<()> {
        let bytes = crate::package::embedded::replace(
            &self.package,
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            index,
            crate::package::embedded::ResourceTarget::Object,
            resource,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }

    pub fn replace_embedded_image(
        &mut self,
        index: usize,
        resource: &crate::EmbeddedResource,
    ) -> Result<()> {
        let bytes = crate::package::embedded::replace(
            &self.package,
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            index,
            crate::package::embedded::ResourceTarget::Image,
            resource,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }

    pub fn remove_embedded_object(&mut self, index: usize) -> Result<()> {
        let bytes = crate::package::embedded::remove(
            &self.package,
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            index,
            crate::package::embedded::ResourceTarget::Object,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }

    pub fn remove_embedded_image(&mut self, index: usize) -> Result<()> {
        let bytes = crate::package::embedded::remove(
            &self.package,
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            index,
            crate::package::embedded::ResourceTarget::Image,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }

    pub fn move_embedded_object(&mut self, from: usize, to: usize) -> Result<()> {
        let bytes = crate::package::embedded::reorder(
            &self.package,
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            from,
            to,
            crate::package::embedded::ResourceTarget::Object,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }

    pub fn move_embedded_image(&mut self, from: usize, to: usize) -> Result<()> {
        let bytes = crate::package::embedded::reorder(
            &self.package,
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            from,
            to,
            crate::package::embedded::ResourceTarget::Image,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }

    /// Return bytes only for inline or verified package-contained images.
    /// Linked images remain inert and are never fetched.
    pub fn image_bytes(&self, image: &crate::Image) -> Result<Option<Vec<u8>>> {
        match &image.source {
            litchi_odf_common::media::Source::Inline { bytes, .. } => Ok(Some(bytes.clone())),
            litchi_odf_common::media::Source::PackagePart { path, .. } => {
                self.package.get_file(path).map(Some)
            },
            _ => Ok(None),
        }
    }

    /// Get a file from the ODF package (useful for extracting images)
    ///
    /// # Arguments
    ///
    /// * `path` - Path to the file within the package (e.g., "Pictures/image1.png")
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odt::Document;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let doc = Document::open("document.odt")?;
    /// let images = doc.image_paths()?;
    /// if let Some(first_image) = images.first() {
    ///     let image_bytes = doc.get_file(first_image)?;
    ///     std::fs::write("extracted_image.png", image_bytes)?;
    /// }
    /// # Ok(())
    /// # }
    /// ```
    pub fn get_file(&self, path: &str) -> Result<Vec<u8>> {
        get_file(self, path)
    }

    // Note: DELETE operations are available via `MutableDocument`. To modify this document:
    //   1. Convert to MutableDocument:  `let mut mutable = MutableDocument::from_document(doc)?`
    //   2. Perform modifications: `mutable.remove_paragraph(0)?`, `mutable.remove_table(1)?`, etc.
    //   3. Save: `mutable.save("output.odt")?`
    // Available methods: remove_paragraph, remove_table, update_paragraph, clear_content, etc.
}

pub(super) fn take(document: Document) -> OwnedPackage {
    document.package
}

pub(super) fn validate_mimetype(mimetype: &str) -> Result<()> {
    if mimetype.contains("opendocument.text") {
        Ok(())
    } else {
        Err(Error::InvalidFormat(format!(
            "Not an ODT file: MIME type is {mimetype}"
        )))
    }
}

pub(super) fn original_bytes(document: &Document) -> &[u8] {
    document.package.as_bytes()
}

pub(super) fn to_bytes(document: &Document) -> Result<Vec<u8>> {
    Ok(document.package.as_bytes().to_vec())
}

pub(super) fn get_file(document: &Document, path: &str) -> Result<Vec<u8>> {
    document.package.get_file(path)
}
