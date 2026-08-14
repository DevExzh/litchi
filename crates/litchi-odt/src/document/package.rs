//! ODT package lifecycle, mutation, and lossless byte access.

use super::model::Document;
use crate::core::{Content, Meta, OwnedPackage, PreparedPackage, Styles};
use crate::elements::style::{StyleElements, StyleRegistry};
use litchi_core::{Error, Result};
use std::path::Path;

impl Document {
    pub(crate) fn into_package(self) -> OwnedPackage {
        take(self)
    }

    pub(crate) fn transaction_package(&self) -> &OwnedPackage {
        &self.package
    }

    pub(crate) fn transaction_content_xml(&self) -> &str {
        self.content.xml_content()
    }

    pub(crate) fn transaction_styles_xml(&self) -> Option<&str> {
        self.styles.as_ref().map(Styles::xml_content)
    }

    pub(crate) fn replace_transaction_bytes(&mut self, bytes: Vec<u8>) -> Result<()> {
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }

    pub(super) fn publish_transaction(&mut self, commit: crate::transaction::Commit) -> Result<()> {
        *self = commit.into_snapshot().document()?;
        Ok(())
    }

    pub(super) fn publish_transaction_path(
        &mut self,
        commit: crate::transaction::Commit,
    ) -> Result<String> {
        let path = match commit.results().last() {
            Some(crate::transaction::OperationResult::Path(path)) => path.clone(),
            _ => {
                return Err(Error::InvalidFormat(
                    "ODT transaction did not return an allocated path".to_string(),
                ));
            },
        };
        self.publish_transaction(commit)?;
        Ok(path)
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

    /// Open a prepared packaged ODF result produced by smart detection.
    ///
    /// The retained ZIP index is transferred into this semantic owner. The
    /// concrete ODT MIME and content contracts are still checked here before
    /// the document model is constructed.
    pub fn from_prepared_package(prepared: PreparedPackage) -> Result<Self> {
        if prepared.format() != litchi_core::detection::FileFormat::Odt {
            return Err(Error::InvalidFormat(
                "prepared ODF package is not an ODT family document".to_string(),
            ));
        }
        Self::from_owned_package(prepared.into_package())
    }

    /// Alias for [`Self::from_prepared_package`].
    #[inline]
    pub fn from_prepared(prepared: PreparedPackage) -> Result<Self> {
        Self::from_prepared_package(prepared)
    }

    /// Return the identity of the archive index retained by this document.
    #[doc(hidden)]
    #[must_use]
    pub fn prepared_index_identity(&self) -> usize {
        self.package.prepared_index_identity()
    }

    pub(crate) fn from_owned_package(owned_package: OwnedPackage) -> Result<Self> {
        let package = owned_package.package()?;

        // Verify this is a text document.
        validate_mimetype(package.mimetype())?;

        // Parse core components
        let content_bytes = package.get_file("content.xml")?;
        let content_xml = std::str::from_utf8(&content_bytes).map_err(|error| {
            Error::InvalidFormat(format!("ODT content.xml is not UTF-8: {error}"))
        })?;
        litchi_odf_common::core::validate_content_document_part(
            content_xml,
            "<office:text",
            "ODT",
        )?;
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

    /// Captures this document as an immutable, source-bound transaction snapshot.
    pub fn snapshot(&self) -> Result<crate::transaction::Snapshot> {
        crate::transaction::Snapshot::from_document(self)
    }

    /// Starts a detached packaged-document transaction.
    ///
    /// This is the preferred mutation boundary for RDF graphs, form trees,
    /// embedded charts and resources, and lossless inline text operations.
    pub fn edit(&self) -> Result<crate::transaction::Edit> {
        Ok(self.snapshot()?.edit())
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
    #[deprecated(
        since = "0.0.1",
        note = "use Document::edit().set_protection() and commit"
    )]
    pub fn set_protection(&mut self, policy: &crate::protection::Policy) -> Result<()> {
        let mut edit = self.edit()?;
        edit.set_protection(policy)?;
        self.publish_transaction(edit.commit()?)
    }

    #[deprecated(
        since = "0.0.1",
        note = "use Document::edit().add_rdf_graph() and commit"
    )]
    pub fn add_rdf_graph(
        &mut self,
        preferred_path: Option<&str>,
        triples: &[crate::rdf::Triple],
    ) -> Result<String> {
        let (bytes, path) = crate::rdf::add_graph(&self.package, preferred_path, triples)?;
        *self = Self::from_bytes(bytes)?;
        Ok(path)
    }

    #[deprecated(
        since = "0.0.1",
        note = "use Document::edit().replace_rdf_graph() and commit"
    )]
    pub fn replace_rdf_graph(&mut self, path: &str, triples: &[crate::rdf::Triple]) -> Result<()> {
        let bytes = crate::rdf::replace_graph(&self.package, path, triples)?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }

    #[deprecated(
        since = "0.0.1",
        note = "use Document::edit().remove_rdf_graph() and commit"
    )]
    pub fn remove_rdf_graph(&mut self, path: &str) -> Result<()> {
        let bytes = crate::rdf::remove_graph(&self.package, path)?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }

    #[deprecated(
        since = "0.0.1",
        note = "use Document::edit().add_rdf_triple() and Commit::results()"
    )]
    pub fn add_rdf_triple(&mut self, path: &str, triple: &crate::rdf::Triple) -> Result<usize> {
        let mut edit = self.edit()?;
        edit.add_rdf_triple(path, triple)?;
        let commit = edit.commit()?;
        let index = transaction_index(&commit)?;
        self.publish_transaction(commit)?;
        Ok(index)
    }

    #[deprecated(
        since = "0.0.1",
        note = "use Document::edit().replace_rdf_triple() and commit"
    )]
    pub fn replace_rdf_triple(
        &mut self,
        path: &str,
        index: usize,
        triple: &crate::rdf::Triple,
    ) -> Result<()> {
        let mut edit = self.edit()?;
        edit.replace_rdf_triple(path, index, triple)?;
        self.publish_transaction(edit.commit()?)
    }

    #[deprecated(
        since = "0.0.1",
        note = "use Document::edit().remove_rdf_triple_at() and commit"
    )]
    pub fn remove_rdf_triple(&mut self, path: &str, index: usize) -> Result<()> {
        let mut edit = self.edit()?;
        edit.remove_rdf_triple_at(path, crate::transaction::Position::new(index))?;
        self.publish_transaction(edit.commit()?)
    }

    #[deprecated(
        since = "0.0.1",
        note = "use Document::edit().move_rdf_triple_to() and commit"
    )]
    pub fn move_rdf_triple(&mut self, path: &str, from: usize, to: usize) -> Result<()> {
        let mut edit = self.edit()?;
        edit.move_rdf_triple_to(
            path,
            crate::transaction::Position::new(from),
            crate::transaction::Position::new(to),
        )?;
        self.publish_transaction(edit.commit()?)
    }

    #[deprecated(since = "0.0.1", note = "use Document::edit().add_form() and commit")]
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

    #[deprecated(
        since = "0.0.1",
        note = "use Document::edit().add_nested_form() and Commit::results()"
    )]
    pub fn add_nested_form(
        &mut self,
        parent_form: usize,
        form: &crate::package::forms::AuthoredForm,
    ) -> Result<usize> {
        let mut edit = self.edit()?;
        edit.add_nested_form(parent_form, form)?;
        let commit = edit.commit()?;
        let index = transaction_index(&commit)?;
        self.publish_transaction(commit)?;
        Ok(index)
    }

    #[deprecated(
        since = "0.0.1",
        note = "use Document::edit().replace_form() and commit"
    )]
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

    #[deprecated(
        since = "0.0.1",
        note = "use Document::edit().remove_form() and commit"
    )]
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

    #[deprecated(since = "0.0.1", note = "use Document::edit().move_form() and commit")]
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

    #[deprecated(
        since = "0.0.1",
        note = "use Document::edit().add_form_control() and Commit::results()"
    )]
    pub fn add_form_control(
        &mut self,
        form_index: usize,
        control: &crate::package::forms::AuthoredFormControl,
    ) -> Result<usize> {
        let mut edit = self.edit()?;
        edit.add_form_control(form_index, control)?;
        let commit = edit.commit()?;
        let index = transaction_index(&commit)?;
        self.publish_transaction(commit)?;
        Ok(index)
    }

    #[deprecated(
        since = "0.0.1",
        note = "use Document::edit().replace_form_control() and commit"
    )]
    pub fn replace_form_control(
        &mut self,
        index: usize,
        control: &crate::package::forms::AuthoredFormControl,
    ) -> Result<()> {
        let mut edit = self.edit()?;
        edit.replace_form_control(index, control)?;
        self.publish_transaction(edit.commit()?)
    }

    #[deprecated(
        since = "0.0.1",
        note = "use Document::edit().remove_form_control_at() and commit"
    )]
    pub fn remove_form_control(&mut self, index: usize) -> Result<()> {
        let mut edit = self.edit()?;
        edit.remove_form_control_at(crate::transaction::Position::new(index))?;
        self.publish_transaction(edit.commit()?)
    }

    #[deprecated(
        since = "0.0.1",
        note = "use Document::edit().move_form_control_to() and commit"
    )]
    pub fn move_form_control(&mut self, from: usize, to: usize) -> Result<()> {
        let mut edit = self.edit()?;
        edit.move_form_control_to(
            crate::transaction::Position::new(from),
            crate::transaction::Position::new(to),
        )?;
        self.publish_transaction(edit.commit()?)
    }

    /// Append a packaged chart object to the text body.
    #[deprecated(
        since = "0.0.1",
        note = "use Document::edit().add_embedded_chart() and commit"
    )]
    #[allow(
        deprecated,
        reason = "compatibility wrapper delegates to explicit storage"
    )]
    pub fn add_embedded_chart(&mut self, definition: &crate::odc::Definition) -> Result<usize> {
        self.add_embedded_chart_with_storage(
            definition,
            crate::EmbeddedChartStorage::PackageSubdocument,
        )
    }

    /// Append a chart object using an explicit storage form.
    #[deprecated(
        since = "0.0.1",
        note = "use Document::edit().add_embedded_chart_with_storage() and Commit::results()"
    )]
    pub fn add_embedded_chart_with_storage(
        &mut self,
        definition: &crate::odc::Definition,
        storage: crate::EmbeddedChartStorage,
    ) -> Result<usize> {
        let mut edit = self.edit()?;
        edit.add_embedded_chart_with_storage(definition, storage)?;
        let commit = edit.commit()?;
        let index = transaction_index(&commit)?;
        self.publish_transaction(commit)?;
        Ok(index)
    }

    #[deprecated(
        since = "0.0.1",
        note = "use Document::edit().replace_embedded_chart() and commit"
    )]
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

    #[deprecated(
        since = "0.0.1",
        note = "use Document::edit().remove_embedded_chart() and commit"
    )]
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
    #[deprecated(
        since = "0.0.1",
        note = "use Document::edit().add_embedded_resource() and commit"
    )]
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

    #[deprecated(
        since = "0.0.1",
        note = "use Document::edit().replace_embedded_object() and commit"
    )]
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

    #[deprecated(
        since = "0.0.1",
        note = "use Document::edit().replace_embedded_image() and commit"
    )]
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

    #[deprecated(
        since = "0.0.1",
        note = "use Document::edit().remove_embedded_object() and commit"
    )]
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

    #[deprecated(
        since = "0.0.1",
        note = "use Document::edit().remove_embedded_image() and commit"
    )]
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

    #[deprecated(
        since = "0.0.1",
        note = "use Document::edit().move_embedded_object() and commit"
    )]
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

    #[deprecated(
        since = "0.0.1",
        note = "use Document::edit().move_embedded_image() and commit"
    )]
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

    // Ordinary paragraph edits should use `Document::edit()` and commit an
    // immutable transaction. `MutableDocument` remains the compatibility
    // surface for structural families that have not migrated yet.
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

fn transaction_index(commit: &crate::transaction::Commit) -> Result<usize> {
    match commit.results().last() {
        Some(crate::transaction::OperationResult::Index(index)) => Ok(*index),
        _ => Err(Error::InvalidFormat(
            "ODT transaction did not return an allocated index".to_string(),
        )),
    }
}
