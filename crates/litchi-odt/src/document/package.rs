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
        // Let the prepared detector perform the local MIME probe exactly
        // once.  A matching ODT result transfers its indexed archive; a
        // different ODF family still follows the historical package owner so
        // its MIME error is reported at the same boundary; rejected probes
        // recover the original allocation for the ordinary parser.
        match litchi_odf_common::detect::prepared_or_original(bytes) {
            Ok(prepared) if prepared.format() == litchi_core::detection::FileFormat::Odt => {
                Self::from_prepared_package(prepared)
            },
            Ok(prepared) => Self::from_owned_package(prepared.into_package()),
            Err(bytes) => Self::from_owned_package(OwnedPackage::from_bytes(bytes)?),
        }
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
        // One fused tokenization validates the content structure and collects
        // the automatic styles (the same `OpenParse` the source-backed facade
        // uses). Error precedence is identical to the historical sequential
        // passes: content validation (mid-stream and end-of-stream) returns
        // here, before `Content` adoption and before styles.xml/meta.xml are
        // fetched; a content-styles error is only recorded and surfaces via
        // `finish` after the styles.xml parse, and a `try_extend` conflict
        // still loses to it — the historical call order.
        let content_styles = super::open_parse::OpenParse::run(content_xml)?;
        let content = Content::from_vec(content_bytes)?;

        let styles = if package.has_file("styles.xml") {
            let styles_bytes = package.get_file("styles.xml")?;
            Some(Styles::from_vec(styles_bytes)?)
        } else {
            None
        };

        let meta = if package.has_file("meta.xml") {
            let meta_bytes = package.get_file("meta.xml")?;
            Some(Meta::from_vec(meta_bytes)?)
        } else {
            None
        };

        // Initialize style registry
        let mut style_registry = StyleRegistry::default();

        // Parse styles from styles.xml if available
        if let Some(ref styles_part) = styles {
            style_registry = StyleElements::parse_styles(styles_part.xml_content())?;
        }

        // Also parse styles from content.xml (automatic styles), collected by
        // the fused pass above; its recorded error surfaces at this point,
        // after the styles.xml parse.
        // Merge content styles into main registry (content styles take precedence)
        style_registry.try_extend(content_styles.finish()?)?;

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

#[cfg(test)]
mod tests {
    use super::Document;
    use crate::core::OwnedPackage;
    use crate::elements::style::StyleElements;
    use litchi_core::Result;
    use soapberry_zip::ZipArchiveWriter;

    const MIMETYPE: &str = "application/vnd.oasis.opendocument.text";
    const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
    const STYLE: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
    const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";

    /// Pre-0222 oracle: the historical sequential owned-path open — full
    /// `validate_content_document_part` scan, then `styles.xml` parse, then
    /// the `content.xml` automatic-styles rescan, then `try_extend` —
    /// retained verbatim to cross-check the fused `OpenParse` open on error
    /// precedence, error messages, and the resulting document state.
    fn from_owned_package_sequential_oracle(owned_package: OwnedPackage) -> Result<Document> {
        let package = owned_package.package()?;

        // Verify this is a text document.
        super::validate_mimetype(package.mimetype())?;

        // Parse core components
        let content_bytes = package.get_file("content.xml")?;
        let content_xml = std::str::from_utf8(&content_bytes).map_err(|error| {
            litchi_core::Error::InvalidFormat(format!("ODT content.xml is not UTF-8: {error}"))
        })?;
        litchi_odf_common::core::validate_content_document_part(
            content_xml,
            "<office:text",
            "ODT",
        )?;
        let content = crate::core::Content::from_vec(content_bytes)?;

        let styles = if package.has_file("styles.xml") {
            let styles_bytes = package.get_file("styles.xml")?;
            Some(crate::core::Styles::from_vec(styles_bytes)?)
        } else {
            None
        };

        let meta = if package.has_file("meta.xml") {
            let meta_bytes = package.get_file("meta.xml")?;
            Some(crate::core::Meta::from_vec(meta_bytes)?)
        } else {
            None
        };

        // Initialize style registry
        let mut style_registry = crate::elements::style::StyleRegistry::default();

        // Parse styles from styles.xml if available
        if let Some(ref styles_part) = styles {
            style_registry = StyleElements::parse_styles(styles_part.xml_content())?;
        }

        // Also parse styles from content.xml (automatic styles)
        // Merge content styles into main registry (content styles take precedence)
        style_registry.try_extend(StyleElements::parse_styles(content.xml_content())?)?;

        Ok(Document {
            package: owned_package,
            content,
            styles,
            meta,
            style_registry,
        })
    }

    fn oracle_from_bytes(bytes: Vec<u8>) -> Result<Document> {
        from_owned_package_sequential_oracle(OwnedPackage::from_bytes(bytes)?)
    }

    /// Deterministic comparison projection: error string on failure, or the
    /// sorted style names with their families plus the full extracted text
    /// on success.
    fn projection(result: Result<Document>) -> std::result::Result<String, String> {
        result
            .and_then(|document| {
                let mut styles: Vec<String> = document
                    .style_registry
                    .styles
                    .iter()
                    .map(|(name, style)| format!("{name}|{:?}", style.family()))
                    .collect();
                styles.sort_unstable();
                Ok(format!("{}\n{}", styles.join(","), document.text()?))
            })
            .map_err(|error| error.to_string())
    }

    fn assert_open_parity(label: &str, bytes: &[u8]) {
        let expected = projection(oracle_from_bytes(bytes.to_vec()));
        let actual = projection(Document::from_bytes(bytes.to_vec()));
        assert_eq!(
            expected, actual,
            "{label}: fused owned open and sequential oracle disagree"
        );
    }

    fn odt_package_with_mimetype(
        mimetype: &str,
        content: &[u8],
        styles: Option<&[u8]>,
        meta: Option<&[u8]>,
    ) -> Vec<u8> {
        // Stored members written directly: `PackageWriter` validates XML on
        // publication, which would reject the deliberately malformed
        // fixtures below.
        let mut writer = ZipArchiveWriter::new(Vec::new());
        let mut built = writer.write_stored_file("mimetype", mimetype.as_bytes());
        let mut extra_members: Vec<(&str, &[u8])> = vec![("content.xml", content)];
        if let Some(styles) = styles {
            extra_members.push(("styles.xml", styles));
        }
        if let Some(meta) = meta {
            extra_members.push(("meta.xml", meta));
        }
        let mut manifest = String::from(
            r#"<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">"#,
        );
        manifest.push_str(&format!(
            r#"<manifest:file-entry manifest:full-path="/" manifest:media-type="{mimetype}"/>"#
        ));
        for (path, _) in &extra_members {
            manifest.push_str(&format!(
                r#"<manifest:file-entry manifest:full-path="{path}" manifest:media-type="text/xml"/>"#
            ));
        }
        manifest.push_str("</manifest:manifest>");
        for (path, bytes) in extra_members {
            built = built.and_then(|()| writer.write_stored_file(path, bytes));
        }
        built = built
            .and_then(|()| writer.write_stored_file("META-INF/manifest.xml", manifest.as_bytes()));
        match built.and_then(|()| writer.finish()) {
            Ok(bytes) => bytes,
            Err(error) => panic!("synthetic ODT package must build: {error}"),
        }
    }

    fn odt_package(content: &str, styles: Option<&str>, meta: Option<&[u8]>) -> Vec<u8> {
        odt_package_with_mimetype(
            MIMETYPE,
            content.as_bytes(),
            styles.map(str::as_bytes),
            meta,
        )
    }

    fn content(body: &str) -> String {
        format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="{OFFICE}" xmlns:style="{STYLE}" xmlns:text="{TEXT}" office:version="1.4">{body}</office:document-content>"#
        )
    }

    fn styles_document(styles: &str) -> String {
        format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><office:document-styles xmlns:office="{OFFICE}" xmlns:style="{STYLE}" office:version="1.4">{styles}</office:document-styles>"#
        )
    }

    const VALID_BODY: &str =
        r#"<office:body><office:text><text:p>hello</text:p></office:text></office:body>"#;
    const STYLE_A: &str = r#"<style:style style:name="A" style:family="paragraph"></style:style>"#;
    // A duplicate style attribute: the styles parse fails with
    // "invalid ODT style attribute: …".
    const DUP_ATTR_STYLE: &str = r#"<style:style style:name="a" style:name="b"></style:style>"#;
    // A mismatched end tag: the styles parse fails with
    // "invalid ODT style XML: …".
    const BROKEN_STYLES_XML: &str = "<a></b>";

    #[test]
    fn fused_owned_open_matches_sequential_oracle_on_synthetic_packages() {
        let with_styles = styles_document(STYLE_A);
        // Each case: label, package bytes, expected outcome — `None` for a
        // successful open, `Some(fragment)` for an error whose message must
        // contain the fragment. The fragments make the cross-stage error
        // precedence observable, not just parity on identical failures.
        let cases: Vec<(&str, Vec<u8>, Option<&str>)> = vec![
            (
                "valid-full",
                odt_package(
                    &content(&format!(
                        r#"<office:automatic-styles><style:style style:name="B" style:family="paragraph"></style:style></office:automatic-styles>{VALID_BODY}"#
                    )),
                    Some(&with_styles),
                    Some(br#"<?xml version="1.0"?><office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:meta/></office:document-meta>"#),
                ),
                None,
            ),
            (
                "valid-bare",
                odt_package(&content(VALID_BODY), None, None),
                None,
            ),
            // styles.xml and content.xml both define style "A": content
            // styles take precedence (try_extend overwrites), no error.
            (
                "same-style-name-in-both",
                odt_package(
                    &content(&format!(
                        r#"<office:automatic-styles>{STYLE_A}</office:automatic-styles>{VALID_BODY}"#
                    )),
                    Some(&with_styles),
                    None,
                ),
                None,
            ),
            // A content-styles error loses to the styles.xml parse error.
            (
                "styles-error-beats-content-style-error",
                odt_package(
                    &content(&format!(
                        r#"<office:automatic-styles>{DUP_ATTR_STYLE}</office:automatic-styles>{VALID_BODY}"#
                    )),
                    Some(BROKEN_STYLES_XML),
                    None,
                ),
                Some("invalid ODT style XML:"),
            ),
            // A content-styles error alone surfaces after a clean styles.xml
            // parse.
            (
                "content-style-error-alone",
                odt_package(
                    &content(&format!(
                        r#"<office:automatic-styles>{DUP_ATTR_STYLE}</office:automatic-styles>{VALID_BODY}"#
                    )),
                    Some(&with_styles),
                    None,
                ),
                Some("invalid ODT style attribute:"),
            ),
            // A content validation error beats any styles.xml failure: the
            // historical validation pass early-returned before styles.xml
            // was even fetched.
            (
                "validation-error-beats-styles-error",
                odt_package(
                    &content(r#"<office:body><office:spreadsheet/></office:body>"#),
                    Some(BROKEN_STYLES_XML),
                    None,
                ),
                Some("has the wrong office body"),
            ),
            // A style error recorded before a later validation failure still
            // loses to the validation error.
            (
                "early-style-error-loses-to-late-validation-error",
                odt_package(
                    &content(&format!(
                        r#"<office:automatic-styles>{DUP_ATTR_STYLE}</office:automatic-styles><office:body><office:spreadsheet/></office:body>"#
                    )),
                    None,
                    None,
                ),
                Some("has the wrong office body"),
            ),
            // Deep validation failure (mismatched end tag) maps to the
            // tokenizer message.
            (
                "mismatched-end-tag-deep",
                odt_package(
                    &content(
                        r#"<office:body><office:text><text:p></text:q></office:text></office:body>"#,
                    ),
                    Some(&with_styles),
                    None,
                ),
                Some("invalid ODT content.xml:"),
            ),
            (
                "wrong-root",
                odt_package(
                    &content(VALID_BODY).replace("document-content", "document"),
                    Some(&with_styles),
                    None,
                ),
                Some("has the wrong root"),
            ),
            (
                "missing-body",
                odt_package(&content(""), None, None),
                Some("has no complete expected body"),
            ),
            (
                "meta-error-after-clean-styles",
                odt_package(&content(VALID_BODY), Some(&with_styles), Some(b"\xff\xfe")),
                Some("Invalid UTF-8 in XML content"),
            ),
        ];
        for (label, bytes, expected) in &cases {
            assert_open_parity(label, bytes);
            let outcome = projection(Document::from_bytes(bytes.clone()));
            match (expected, &outcome) {
                (None, Ok(_)) => {},
                (Some(fragment), Err(error)) => assert!(
                    error.contains(fragment),
                    "{label}: error {error:?} does not contain {fragment:?}"
                ),
                _ => panic!("{label}: outcome {outcome:?} misses expectation {expected:?}"),
            }
        }

        // Mimetype and container-level failures agree as well.
        let wrong_mime = odt_package_with_mimetype(
            "application/vnd.oasis.opendocument.presentation",
            content(VALID_BODY).as_bytes(),
            None,
            None,
        );
        assert_open_parity("wrong-mimetype", &wrong_mime);
        let outcome = projection(Document::from_bytes(wrong_mime));
        match outcome {
            Err(error) => assert!(error.contains("Not an ODT file"), "unexpected: {error}"),
            Ok(_) => panic!("wrong mimetype accepted"),
        }
        assert_open_parity("not-a-zip", b"this is not a package");
        match projection(Document::from_bytes(b"this is not a package".to_vec())) {
            Err(_) => {},
            Ok(_) => panic!("non-ZIP bytes accepted"),
        }
    }

    #[test]
    fn fused_owned_open_matches_sequential_oracle_on_odt_corpus() {
        let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../..");
        let mut files = Vec::new();
        collect_odt(&root.join("test-data"), &mut files);
        files.sort();
        assert!(!files.is_empty(), "no .odt corpus fixtures discovered");
        for path in &files {
            let Ok(bytes) = std::fs::read(path) else {
                continue;
            };
            assert_open_parity(&path.display().to_string(), &bytes);
        }
    }

    fn collect_odt(directory: &std::path::Path, files: &mut Vec<std::path::PathBuf>) {
        let Ok(entries) = std::fs::read_dir(directory) else {
            return;
        };
        for entry in entries.flatten() {
            let path = entry.path();
            if path.is_dir() {
                collect_odt(&path, files);
            } else if path.extension().is_some_and(|extension| extension == "odt") {
                files.push(path);
            }
        }
    }
}
