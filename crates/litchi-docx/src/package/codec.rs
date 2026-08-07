//! OPC construction, validation, and publication codecs for DOCX packages.

use super::model::*;
#[cfg(feature = "encryption")]
pub(super) use super::model::{Limits, Mode};

impl Package {
    /// Create a new empty .docx package.
    ///
    /// Creates a minimal valid Word document with default styles and settings.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_docx::Package;
    ///
    /// let mut pkg = Package::new()?;
    /// // Add content to the document...
    /// pkg.save("new_document.docx")?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn new() -> Result<Self> {
        use crate::template;
        use litchi_opc::constants::content_type as ct;
        use litchi_opc::constants::relationship_type as rt;
        use litchi_opc::packuri::PackURI;
        use litchi_opc::part::BlobPart;

        let mut opc = OpcPackage::new();

        // Create document.xml part
        let doc_partname = PackURI::new("/word/document.xml")
            .map_err(|e| Error::InvalidUri(format!("document partname: {}", e)))?;
        let doc_part = BlobPart::new(
            doc_partname.clone(),
            ct::WML_DOCUMENT_MAIN.to_string(),
            template::default_document_xml().as_bytes().to_vec(),
        );

        // Create relationship from package to document (use relative path for package-level rels)
        opc.relate_to("word/document.xml", rt::OFFICE_DOCUMENT);
        opc.add_part(Box::new(doc_part));

        // Create styles.xml part with dynamic style generation
        let styles_partname = PackURI::new("/word/styles.xml")
            .map_err(|e| Error::InvalidUri(format!("styles partname: {}", e)))?;

        // Generate default styles dynamically
        use crate::writer::style::{MutableStyle, generate_styles_xml};
        let default_styles = vec![
            MutableStyle::normal(),
            MutableStyle::heading_1(),
            MutableStyle::heading_2(),
            MutableStyle::heading_3(),
            MutableStyle::heading_4(),
            MutableStyle::heading_5(),
            MutableStyle::heading_6(),
            MutableStyle::heading_7(),
            MutableStyle::heading_8(),
            MutableStyle::heading_9(),
            MutableStyle::title(),
            MutableStyle::default_paragraph_font(),
            MutableStyle::toc_heading(),
            MutableStyle::toc1(),
            MutableStyle::toc2(),
            MutableStyle::toc3(),
            MutableStyle::hyperlink(),
            MutableStyle::header(),
            MutableStyle::footer(),
            MutableStyle::footnote_text(),
            MutableStyle::endnote_text(),
        ];
        let styles_xml = generate_styles_xml(&default_styles)?;

        let styles_part = BlobPart::new(
            styles_partname.clone(),
            ct::WML_STYLES.to_string(),
            styles_xml.as_bytes().to_vec(),
        );

        // Add relationship from document to styles (use relative path)
        if let Ok(doc_part) = opc.get_part_mut(&doc_partname) {
            doc_part.relate_to("styles.xml", rt::STYLES);
        }
        opc.add_part(Box::new(styles_part));

        // Create settings.xml part
        let settings_partname = PackURI::new("/word/settings.xml")
            .map_err(|e| Error::InvalidUri(format!("settings partname: {}", e)))?;
        let settings_part = BlobPart::new(
            settings_partname,
            ct::WML_SETTINGS.to_string(),
            template::default_settings_xml().as_bytes().to_vec(),
        );

        if let Ok(doc_part) = opc.get_part_mut(&doc_partname) {
            doc_part.relate_to("settings.xml", rt::SETTINGS);
        }
        opc.add_part(Box::new(settings_part));

        // Create fontTable.xml part
        let font_table_partname = PackURI::new("/word/fontTable.xml")
            .map_err(|e| Error::InvalidUri(format!("fontTable partname: {}", e)))?;
        let font_table_part = BlobPart::new(
            font_table_partname,
            ct::WML_FONT_TABLE.to_string(),
            template::default_font_table_xml().as_bytes().to_vec(),
        );

        if let Ok(doc_part) = opc.get_part_mut(&doc_partname) {
            doc_part.relate_to("fontTable.xml", rt::FONT_TABLE);
        }
        opc.add_part(Box::new(font_table_part));

        // Create webSettings.xml part
        let web_settings_partname = PackURI::new("/word/webSettings.xml")
            .map_err(|e| Error::InvalidUri(format!("webSettings partname: {}", e)))?;
        let web_settings_xml = docx_web::write(
            &docx_web::Settings::default(),
            docx_web::Conformance::Transitional,
        )?;
        let web_settings_part = BlobPart::new(
            web_settings_partname,
            ct::WML_WEB_SETTINGS.to_string(),
            web_settings_xml,
        );

        if let Ok(doc_part) = opc.get_part_mut(&doc_partname) {
            doc_part.relate_to("webSettings.xml", rt::WEB_SETTINGS);
        }
        opc.add_part(Box::new(web_settings_part));

        // Create core.xml part (core properties)
        let core_props_partname = PackURI::new("/docProps/core.xml")
            .map_err(|e| Error::InvalidUri(format!("core.xml partname: {}", e)))?;
        let core_props_part = BlobPart::new(
            core_props_partname,
            ct::OPC_CORE_PROPERTIES.to_string(),
            template::default_core_props_xml().as_bytes().to_vec(),
        );

        opc.relate_to("docProps/core.xml", rt::CORE_PROPERTIES);
        opc.add_part(Box::new(core_props_part));

        // Create app.xml part (extended properties)
        let app_props_partname = PackURI::new("/docProps/app.xml")
            .map_err(|e| Error::InvalidUri(format!("app.xml partname: {}", e)))?;
        let app_props_part = BlobPart::new(
            app_props_partname,
            ct::OFC_EXTENDED_PROPERTIES.to_string(),
            template::default_app_props_xml().as_bytes().to_vec(),
        );

        opc.relate_to("docProps/app.xml", rt::EXTENDED_PROPERTIES);
        opc.add_part(Box::new(app_props_part));

        // Create theme1.xml part
        let theme_partname = PackURI::new("/word/theme/theme1.xml")
            .map_err(|e| Error::InvalidUri(format!("theme partname: {}", e)))?;
        let theme_part = BlobPart::new(
            theme_partname,
            ct::OFC_THEME.to_string(),
            template::default_theme_xml().as_bytes().to_vec(),
        );

        // Add relationship from document to theme (use relative path)
        if let Ok(doc_part) = opc.get_part_mut(&doc_partname) {
            doc_part.relate_to("theme/theme1.xml", rt::THEME);
        }
        opc.add_part(Box::new(theme_part));

        // Create numbering.xml part
        let numbering_partname = PackURI::new("/word/numbering.xml")
            .map_err(|e| Error::InvalidUri(format!("numbering partname: {}", e)))?;
        let numbering_part = BlobPart::new(
            numbering_partname,
            ct::WML_NUMBERING.to_string(),
            template::default_numbering_xml().as_bytes().to_vec(),
        );

        // Add relationship from document to numbering (use relative path)
        if let Ok(doc_part) = opc.get_part_mut(&doc_partname) {
            doc_part.relate_to("numbering.xml", rt::NUMBERING);
        }
        opc.add_part(Box::new(numbering_part));

        // Create a mutable document for writing
        let mutable_doc = Some(MutableDocument::new());

        // Initialize document properties
        let properties = Slot::load(&opc)?;

        // Initialize custom properties
        let custom_props = CustomProps::new();

        Ok(Self {
            opc,
            mutable_doc,
            raw_edit_committed: false,
            properties,
            custom_props,
            custom_props_dirty: false,
            #[cfg(feature = "fonts")]
            font_embedding: None,
            #[cfg(feature = "encryption")]
            source_encryption: None,
        })
    }

    /// Create a new empty macro-free Word template (`.dotx`) package.
    ///
    /// Template packages are the native container for reusable AutoText and
    /// other building blocks authored through [`Self::put_glossary`].
    pub fn new_template() -> Result<Self> {
        let mut package = Self::new()?;
        let main = package.opc.main_document_part()?.partname().clone();
        package
            .opc
            .get_part_mut(&main)?
            .set_content_type(ct::WML_TEMPLATE_MAIN.to_owned())?;
        Ok(package)
    }

    /// Open a .docx, .docm, .dotx, or .dotm package from a file path.
    ///
    /// # Arguments
    ///
    /// * `path` - Path to the .docx file
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self> {
        Self::open_with_limits(path, litchi_opc::ReadLimits::default())
    }

    /// Open a DOCX package from a file path with explicit OPC resource limits.
    pub fn open_with_limits<P: AsRef<Path>>(
        path: P,
        limits: litchi_opc::ReadLimits,
    ) -> Result<Self> {
        Self::from_opc_package(OpcPackage::open_with_limits(path, limits)?)
    }

    #[cfg(feature = "encryption")]
    pub fn open_with_password<P: AsRef<Path>>(path: P, password: &str) -> Result<Self> {
        Self::open_with_password_and_limits(
            path,
            password,
            &Limits::default(),
            litchi_opc::ReadLimits::default(),
        )
    }

    /// Open with explicit outer-encryption limits.
    ///
    /// The decrypted OPC archive uses [`litchi_opc::ReadLimits::default`].
    /// Use [`Self::open_with_password_and_limits`] to select both independent
    /// resource policies.
    #[cfg(feature = "encryption")]
    pub fn open_with<P: AsRef<Path>>(path: P, password: &str, limits: &Limits) -> Result<Self> {
        Self::open_with_password_and_limits(
            path,
            password,
            limits,
            litchi_opc::ReadLimits::default(),
        )
    }

    /// Open an encrypted DOCX with independent outer-encryption and inner-OPC
    /// resource policies.
    ///
    /// `encryption_limits` bounds encrypted input and decryption. `opc_limits`
    /// is applied to the decrypted archive before it is adopted as a DOCX
    /// package.
    #[cfg(feature = "encryption")]
    pub fn open_with_password_and_limits<P: AsRef<Path>>(
        path: P,
        password: &str,
        encryption_limits: &Limits,
        opc_limits: litchi_opc::ReadLimits,
    ) -> Result<Self> {
        let file = std::fs::File::open(path.as_ref()).map_err(Error::Io)?;
        let opened = crate::encryption::load_with(file, password, encryption_limits)?;
        Self::from_opened_with_limits(opened, opc_limits)
    }

    /// Create a Package from an already-parsed OPC package.
    ///
    /// This is used for single-pass parsing where the OPC package has already
    /// been parsed during format detection. It avoids double-parsing.
    ///
    /// # Arguments
    ///
    /// * `opc` - An already-parsed OPC package
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_docx::Package;
    /// use litchi_opc::OpcPackage;
    /// use std::io::Cursor;
    ///
    /// let bytes = std::fs::read("document.docx")?;
    /// let opc = OpcPackage::from_reader(Cursor::new(bytes))?;
    /// let pkg = Package::from_opc_package(opc)?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn from_opc_package(opc: OpcPackage) -> Result<Self> {
        // Verify it's a Word document by checking the main part's content type
        let main_part = opc
            .main_document_part()
            .map_err(|e| Error::PartNotFound(format!("main document part: {}", e)))?;

        validate_document_main_content_type(main_part.content_type())?;

        let custom_props = CustomProps::read_for(&opc, CustomPropsHost::Word)?;
        let properties = Slot::load(&opc)?;

        Ok(Self {
            opc,
            mutable_doc: None,
            raw_edit_committed: false,
            properties,
            custom_props,
            custom_props_dirty: false,
            #[cfg(feature = "fonts")]
            font_embedding: None,
            #[cfg(feature = "encryption")]
            source_encryption: None,
        })
    }

    /// Create a .docx, .docm, or .dotm package from a reader.
    ///
    /// # Arguments
    ///
    /// * `reader` - A reader containing the .docx file data (must implement Read + Seek)
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_docx::Package;
    /// use std::io::Cursor;
    ///
    /// let data = std::fs::read("document.docx")?;
    /// let cursor = Cursor::new(data);
    /// let pkg = Package::from_reader(cursor)?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn from_reader<R: Read + Seek>(reader: R) -> Result<Self> {
        Self::from_reader_with_limits(reader, litchi_opc::ReadLimits::default())
    }

    /// Create a DOCX package from a reader with explicit OPC resource limits.
    pub fn from_reader_with_limits<R: Read + Seek>(
        reader: R,
        limits: litchi_opc::ReadLimits,
    ) -> Result<Self> {
        Self::from_opc_package(OpcPackage::from_reader_with_limits(reader, limits)?)
    }

    #[cfg(feature = "encryption")]
    pub fn from_reader_with_password<R: Read>(reader: R, password: &str) -> Result<Self> {
        Self::from_reader_with_password_and_limits(
            reader,
            password,
            &Limits::default(),
            litchi_opc::ReadLimits::default(),
        )
    }

    /// Open a reader with explicit outer-encryption limits.
    ///
    /// The decrypted OPC archive uses [`litchi_opc::ReadLimits::default`].
    /// Use [`Self::from_reader_with_password_and_limits`] to select both
    /// independent resource policies.
    #[cfg(feature = "encryption")]
    pub fn from_reader_with<R: Read>(reader: R, password: &str, limits: &Limits) -> Result<Self> {
        Self::from_reader_with_password_and_limits(
            reader,
            password,
            limits,
            litchi_opc::ReadLimits::default(),
        )
    }

    /// Open an encrypted DOCX reader with independent outer-encryption and
    /// inner-OPC resource policies.
    #[cfg(feature = "encryption")]
    pub fn from_reader_with_password_and_limits<R: Read>(
        reader: R,
        password: &str,
        encryption_limits: &Limits,
        opc_limits: litchi_opc::ReadLimits,
    ) -> Result<Self> {
        let opened = crate::encryption::load_with(reader, password, encryption_limits)?;
        Self::from_opened_with_limits(opened, opc_limits)
    }

    #[cfg(feature = "encryption")]
    fn from_opened_with_limits(
        opened: crate::encryption::Opened,
        limits: litchi_opc::ReadLimits,
    ) -> Result<Self> {
        let source_encryption = opened.mode();
        let opc = OpcPackage::from_vec_with_limits(opened.into_bytes(), limits)?;
        let mut package = Self::from_opc_package(opc)?;
        package.source_encryption = source_encryption;
        Ok(package)
    }
}

impl Package {
    /// Save the package to a file.
    ///
    /// Writes the complete Word document including all parts, relationships,
    /// and content types to a .docx file.
    ///
    /// # Arguments
    /// * `path` - Path where the .docx file should be written
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_docx::Package;
    ///
    /// let mut pkg = Package::new()?;
    /// // Modify document...
    /// pkg.save("output.docx")?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn save<P: AsRef<Path>>(&mut self, path: P) -> Result<()> {
        self.ensure_plain_output("save")?;
        self.save_plain_impl(path)
    }

    /// Explicitly save a plaintext package, even when the source was encrypted.
    pub fn save_plain<P: AsRef<Path>>(&mut self, path: P) -> Result<()> {
        self.save_plain_impl(path)
    }

    fn save_plain_impl<P: AsRef<Path>>(&mut self, path: P) -> Result<()> {
        litchi_opc::atomic::replace_with::<Error>(path.as_ref(), |temporary| {
            self.write_plain(temporary)
        })
    }

    /// Save the package to a stream.
    ///
    /// Writes the complete Word document including all parts, relationships,
    /// and content types to a writer stream.
    ///
    /// # Arguments
    /// * `writer` - A writer that implements Write + Seek
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_docx::Package;
    /// use std::io::Cursor;
    ///
    /// let mut pkg = Package::new()?;
    /// // Modify document...
    /// let mut cursor = Cursor::new(Vec::new());
    /// pkg.to_stream(&mut cursor)?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn to_stream<W: Write + Seek>(&mut self, writer: W) -> Result<()> {
        self.ensure_plain_output("to_stream")?;
        self.write_plain(writer)
    }

    /// Explicitly write a plaintext package to a stream.
    pub fn to_plain_stream<W: Write + Seek>(&mut self, writer: W) -> Result<()> {
        self.write_plain(writer)
    }

    /// Serialize and encrypt this package entirely in memory.
    #[cfg(feature = "encryption")]
    pub fn to_encrypted(&mut self, password: &str, mode: Mode) -> Result<Vec<u8>> {
        let mut output = std::io::Cursor::new(Vec::new());
        self.write_plain(&mut output)?;
        crate::encryption::encrypt(output.into_inner(), password, mode).map_err(Into::into)
    }

    /// Serialize and encrypt using the source package's retained profile.
    #[cfg(feature = "encryption")]
    pub fn to_reencrypted(&mut self, password: &str) -> Result<Vec<u8>> {
        let mode = self.preserved_mode("to_reencrypted")?;
        self.to_encrypted(password, mode)
    }

    /// Save with an explicit encryption profile and a borrowed password.
    #[cfg(feature = "encryption")]
    pub fn save_encrypted<P: AsRef<Path>>(
        &mut self,
        path: P,
        password: &str,
        mode: Mode,
    ) -> Result<()> {
        let output = self.to_encrypted(password, mode)?;
        litchi_opc::atomic::replace(path.as_ref(), |temporary| {
            temporary.write_all(&output)?;
            Ok(())
        })?;
        self.source_encryption = Some(mode);
        Ok(())
    }

    /// Save using the encrypted source's retained profile.
    #[cfg(feature = "encryption")]
    pub fn save_reencrypted<P: AsRef<Path>>(&mut self, path: P, password: &str) -> Result<()> {
        let mode = self.preserved_mode("save_reencrypted")?;
        self.save_encrypted(path, password, mode)
    }

    /// Encryption profile of the opened or most recently encrypted package.
    #[cfg(feature = "encryption")]
    pub const fn encryption(&self) -> Option<Mode> {
        self.source_encryption
    }

    pub(super) fn ensure_opc_current(&self, operation: &'static str) -> Result<()> {
        #[cfg(feature = "encryption")]
        if self.source_encryption.is_some() {
            return Err(Error::UnsafeEdit {
                format: "DOCX",
                operation,
                reason: "raw OPC access would expose an encrypted source as plaintext; use the managed encryption or explicit plaintext APIs",
            });
        }

        if self
            .mutable_doc
            .as_ref()
            .is_some_and(MutableDocument::is_modified)
        {
            return Err(Error::UnsafeEdit {
                format: "DOCX",
                operation,
                reason: "the legacy document writer has unmaterialized changes; use a managed save or to_plain_stream first",
            });
        }
        if self.properties.is_dirty() {
            return Err(Error::UnsafeEdit {
                format: "DOCX",
                operation,
                reason: "core properties have unmaterialized changes; use a managed save or to_plain_stream first",
            });
        }
        if self.custom_props_dirty {
            return Err(Error::UnsafeEdit {
                format: "DOCX",
                operation,
                reason: "custom properties have unmaterialized changes; use a managed save or to_plain_stream first",
            });
        }

        Ok(())
    }

    fn ensure_plain_output(&self, _operation: &'static str) -> Result<()> {
        #[cfg(feature = "encryption")]
        if self.source_encryption.is_some() {
            return Err(Error::UnsafeEdit {
                format: "DOCX",
                operation: _operation,
                reason: "the source package was encrypted; use save_reencrypted or save_plain",
            });
        }
        Ok(())
    }

    #[cfg(feature = "encryption")]
    fn preserved_mode(&self, operation: &'static str) -> Result<Mode> {
        self.source_encryption.ok_or(Error::UnsafeEdit {
            format: "DOCX",
            operation,
            reason: "the source package has no encryption profile; supply an explicit Mode",
        })
    }

    fn write_plain<W: Write + Seek>(&mut self, writer: W) -> Result<()> {
        use crate::writer::relmap::RelationshipMapper;
        use litchi_opc::constants::relationship_type as rt;

        // Keep both the source graph and the mutable semantic document
        // available until the complete publication succeeds. Materializing a
        // document rebuilds many related parts; an error half-way through must
        // leave the caller with the same retryable edit rather than a dropped
        // writer and a partially rewritten package.
        let mut rollback = WriteRollbackGuard::new(self);
        // A sink or a late writer hook may unwind instead of returning an
        // error. Catch the unwind long enough to put both owned pieces of the
        // host state back before resuming it; the staged-properties guard is
        // dropped while unwinding the closure and therefore keeps its dirty
        // intent for the next attempt.
        let result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| -> Result<()> {
            // If we have a mutable document, update the document.xml part
            if let Some(mutable_doc) = rollback.mutable_doc_mut()
                && mutable_doc.is_modified()
            {
                // Generate TOC if configured (must happen before serialization)
                mutable_doc.generate_toc_if_needed()?;

                // Step 1: Collect all content that needs relationships
                let hyperlink_urls = mutable_doc.collect_hyperlink_urls();
                let images = mutable_doc.collect_images();
                let ole_objects = mutable_doc.collect_ole_objects();
                let smart_arts = mutable_doc.collect_smart_arts();
                let has_header = mutable_doc.has_header();
                let has_footer = mutable_doc.has_footer();
                let section_header_footer_parts =
                    mutable_doc.collect_section_header_footer_parts()?;
                let explicit_section_relationships =
                    mutable_doc.collect_explicit_section_header_footer_relationships()?;
                let mut planned_section_parts = Vec::new();
                for (index, (header, part)) in section_header_footer_parts.into_iter().enumerate() {
                    let stem = if header {
                        "headerSection"
                    } else {
                        "footerSection"
                    };
                    let filename = format!("{stem}{}.xml", index + 1);
                    let uri = PackURI::new(format!("/word/{filename}"))
                        .map_err(|error| Error::InvalidUri(error.to_string()))?;
                    if self.opc.get_part(&uri).is_ok() {
                        return Err(Error::InvalidFormat(format!(
                            "section header/footer part {uri} already exists"
                        )));
                    }
                    planned_section_parts.push((header, part, uri, filename));
                }

                // Step 2: Create a relationship mapper and add relationships
                let mut rel_mapper = RelationshipMapper::new();

                // Create the document part first (we'll update it later)
                let doc_uri = PackURI::new("/word/document.xml")
                    .map_err(|e| Error::InvalidUri(format!("document URI: {}", e)))?;

                if !explicit_section_relationships.is_empty() {
                    let existing_document = self.opc.get_part(&doc_uri).map_err(|_| {
                        Error::InvalidFormat(
                            "section references exist without a document part".to_string(),
                        )
                    })?;
                    for (id, header) in &explicit_section_relationships {
                        let relationship = existing_document.rels().get(id).ok_or_else(|| {
                            Error::InvalidFormat(format!("section relationship {id:?} is missing"))
                        })?;
                        let expected_type = if *header {
                            [
                                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header",
                                "http://purl.oclc.org/ooxml/officeDocument/relationships/header",
                            ]
                        } else {
                            [
                                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer",
                                "http://purl.oclc.org/ooxml/officeDocument/relationships/footer",
                            ]
                        };
                        if relationship.is_external()
                            || !expected_type.contains(&relationship.reltype())
                        {
                            return Err(Error::InvalidFormat(format!(
                                "section relationship {id:?} has the wrong type or target mode"
                            )));
                        }
                        let target = relationship.target_partname().map_err(|error| {
                            Error::InvalidFormat(format!(
                                "invalid section relationship {id:?}: {error}"
                            ))
                        })?;
                        let part = self.opc.get_part(&target).map_err(|_| {
                            Error::InvalidFormat(format!(
                                "section relationship {id:?} targets a missing part"
                            ))
                        })?;
                        let expected_content_type = if *header {
                            ct::WML_HEADER
                        } else {
                            ct::WML_FOOTER
                        };
                        if part.content_type() != expected_content_type {
                            return Err(Error::InvalidFormat(format!(
                                "section relationship {id:?} targets the wrong content type"
                            )));
                        }
                    }
                }

                // Get or create the document part to add relationships to
                let content_type = self
                    .opc
                    .get_part(&doc_uri)
                    .map(|p| p.content_type().to_string())
                    .unwrap_or_else(|_| ct::WML_DOCUMENT_MAIN.to_string());

                // Create new temporary part for relationships
                use litchi_opc::part::{BlobPart, Part};
                let mut temp_part =
                    BlobPart::new(doc_uri.clone(), content_type.clone(), Vec::new());

                // Copy existing relationships from the original document part (styles, settings, etc.)
                if let Ok(existing_part) = self.opc.get_part(&doc_uri) {
                    for rel in existing_part.rels().iter() {
                        // Skip relationships we're going to recreate dynamically
                        if !matches!(
                            rel.reltype(),
                            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
                                | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
                                | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
                                | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes"
                        ) {
                            temp_part.rels_mut().add_relationship(
                                rel.reltype().to_string(),
                                rel.target_ref().to_string(),
                                rel.r_id().to_string(),
                                rel.is_external(),
                            );
                        }
                    }
                }

                for (header, part, _, filename) in &planned_section_parts {
                    let relationship_type = if *header { rt::HEADER } else { rt::FOOTER };
                    let rid = temp_part.relate_to(filename, relationship_type);
                    rel_mapper.add_section_header_footer_id(part.key.clone(), rid);
                }

                // Add hyperlink relationships (external)
                for (i, url) in hyperlink_urls.iter().enumerate() {
                    let rid = temp_part.relate_to_ext(url, rt::HYPERLINK);
                    rel_mapper.add_hyperlink(i, rid);
                }

                // Add image parts and relationships
                for (i, (image_data, image_format)) in images.iter().enumerate() {
                    let image_num = i + 1;
                    let ext = image_format.extension();
                    let image_partname = format!("/word/media/image{}.{}", image_num, ext);
                    let image_uri = PackURI::new(&image_partname)
                        .map_err(|e| Error::InvalidUri(format!("image URI: {}", e)))?;

                    // Create and add image part
                    let image_part = BlobPart::new(
                        image_uri,
                        image_format.mime_type().to_string(),
                        image_data.to_vec(),
                    );
                    self.opc.add_part(Box::new(image_part));

                    // Create relationship from document to image
                    let rid = temp_part.relate_to(&image_partname, rt::IMAGE);
                    rel_mapper.add_image(i, rid);
                }

                // Add embedded OLE object parts and relationships. Payloads
                // are stored verbatim as inert binary parts; optional
                // previews are stored as ordinary media parts.
                for (i, object) in ole_objects.iter().enumerate() {
                    let object_num = i + 1;
                    let object_partname = format!("/word/embeddings/oleObject{object_num}.bin");
                    let object_uri = PackURI::new(&object_partname)
                        .map_err(|e| Error::InvalidUri(format!("OLE object URI: {}", e)))?;
                    self.opc.add_part(Box::new(BlobPart::new(
                        object_uri,
                        ct::OFC_OLE_OBJECT.to_string(),
                        object.payload().to_vec(),
                    )));
                    let rid = temp_part.relate_to(&object_partname, rt::OLE_OBJECT);
                    rel_mapper.add_ole_object(object.shape_id(), rid);

                    if let Some((preview_data, preview_format)) = object.preview() {
                        let preview_partname = format!(
                            "/word/media/oleObjectPreview{object_num}.{}",
                            preview_format.extension()
                        );
                        let preview_uri = PackURI::new(&preview_partname)
                            .map_err(|e| Error::InvalidUri(format!("OLE preview URI: {}", e)))?;
                        self.opc.add_part(Box::new(BlobPart::new(
                            preview_uri,
                            preview_format.mime_type().to_string(),
                            preview_data.to_vec(),
                        )));
                        let rid = temp_part.relate_to(&preview_partname, rt::IMAGE);
                        rel_mapper.add_ole_preview(object.shape_id(), rid);
                    }
                }

                // Add SmartArt diagram parts (data, layout, quick style,
                // colors) and their relationships. The optional pre-rendered
                // drawing part is not generated; Word and LibreOffice
                // re-render from the layout and data parts.
                let mut diagram_index = 0u32;
                for smartart in &smart_arts {
                    // Allocate non-colliding part names under /word/diagrams/.
                    let (data_name, layout_name, quick_style_name, colors_name) = loop {
                        diagram_index = diagram_index.checked_add(1).ok_or_else(|| {
                            Error::InvalidFormat(
                                "SmartArt diagram part name space exhausted".to_string(),
                            )
                        })?;
                        let names = (
                            format!("/word/diagrams/data{diagram_index}.xml"),
                            format!("/word/diagrams/layout{diagram_index}.xml"),
                            format!("/word/diagrams/quickStyle{diagram_index}.xml"),
                            format!("/word/diagrams/colors{diagram_index}.xml"),
                        );
                        let taken = [&names.0, &names.1, &names.2, &names.3].iter().any(|name| {
                            PackURI::new(*name)
                                .map(|uri| self.opc.get_part(&uri).is_ok())
                                .unwrap_or(true)
                        });
                        if !taken {
                            break names;
                        }
                    };
                    let parts = smartart.generate_parts();
                    for (partname, content_type, xml) in [
                        (&data_name, ct::DML_DIAGRAM_DATA, parts.data_xml),
                        (&layout_name, ct::DML_DIAGRAM_LAYOUT, parts.layout_xml),
                        (
                            &quick_style_name,
                            ct::DML_DIAGRAM_STYLE,
                            parts.quick_style_xml,
                        ),
                        (&colors_name, ct::DML_DIAGRAM_COLORS, parts.colors_xml),
                    ] {
                        let uri = PackURI::new(partname)
                            .map_err(|e| Error::InvalidUri(format!("diagram URI: {}", e)))?;
                        self.opc.add_part(Box::new(BlobPart::new(
                            uri,
                            content_type.to_string(),
                            xml.into_bytes(),
                        )));
                    }
                    let rel_ids = crate::writer::smartart::SmartArtRelIds {
                        data: temp_part.relate_to(&data_name, DIAGRAM_DATA_REL),
                        layout: temp_part.relate_to(&layout_name, DIAGRAM_LAYOUT_REL),
                        quick_style: temp_part
                            .relate_to(&quick_style_name, DIAGRAM_QUICK_STYLE_REL),
                        colors: temp_part.relate_to(&colors_name, DIAGRAM_COLORS_REL),
                    };
                    rel_mapper.add_smart_art(smartart.anchor_key(), rel_ids);
                }

                // Add header/footer parts and relationships
                // Note: If watermark exists, headers will be handled by update_watermark_headers
                // which merges user content with watermark
                if has_header
                    && !mutable_doc.has_watermark()
                    && let Some(header_xml) = mutable_doc.generate_header_xml()?
                {
                    let header_uri = PackURI::new("/word/header1.xml")
                        .map_err(|e| Error::InvalidUri(format!("header URI: {}", e)))?;
                    let header_part = BlobPart::new(
                        header_uri,
                        ct::WML_HEADER.to_string(),
                        header_xml.into_bytes(),
                    );
                    self.opc.add_part(Box::new(header_part));
                    // Use relative path for relationship (relative to document.xml location)
                    let rid = temp_part.relate_to("header1.xml", rt::HEADER);
                    rel_mapper.set_header_id(rid);
                }

                if has_footer && let Some(footer_xml) = mutable_doc.generate_footer_xml()? {
                    let footer_uri = PackURI::new("/word/footer1.xml")
                        .map_err(|e| Error::InvalidUri(format!("footer URI: {}", e)))?;
                    let footer_part = BlobPart::new(
                        footer_uri,
                        ct::WML_FOOTER.to_string(),
                        footer_xml.into_bytes(),
                    );
                    self.opc.add_part(Box::new(footer_part));
                    // Use relative path for relationship (relative to document.xml location)
                    let rid = temp_part.relate_to("footer1.xml", rt::FOOTER);
                    rel_mapper.set_footer_id(rid);
                }

                // Add footnotes parts and relationships BEFORE document XML generation
                if let Some(footnotes_xml) = mutable_doc.generate_footnotes_xml()? {
                    let footnotes_uri = PackURI::new("/word/footnotes.xml")
                        .map_err(|e| Error::InvalidUri(format!("footnotes URI: {}", e)))?;
                    let footnotes_part = BlobPart::new(
                        footnotes_uri,
                        ct::WML_FOOTNOTES.to_string(),
                        footnotes_xml.into_bytes(),
                    );
                    self.opc.add_part(Box::new(footnotes_part));
                    let rid = temp_part.relate_to("footnotes.xml", rt::FOOTNOTES);
                    rel_mapper.set_footnotes_id(rid);
                }

                // Add endnotes parts and relationships BEFORE document XML generation
                if let Some(endnotes_xml) = mutable_doc.generate_endnotes_xml()? {
                    let endnotes_uri = PackURI::new("/word/endnotes.xml")
                        .map_err(|e| Error::InvalidUri(format!("endnotes URI: {}", e)))?;
                    let endnotes_part = BlobPart::new(
                        endnotes_uri,
                        ct::WML_ENDNOTES.to_string(),
                        endnotes_xml.into_bytes(),
                    );
                    self.opc.add_part(Box::new(endnotes_part));
                    let rid = temp_part.relate_to("endnotes.xml", rt::ENDNOTES);
                    rel_mapper.set_endnotes_id(rid);
                }

                // Handle watermark headers before generating document XML
                // This ensures header relationships are properly set up
                if mutable_doc.has_watermark() || mutable_doc.has_image_watermark() {
                    // Generate user header content if exists (will be merged with watermark)
                    let user_header_content = if mutable_doc.has_header() {
                        mutable_doc.generate_header_xml()?
                    } else {
                        None
                    };

                    // Store the watermark image as an ordinary media part,
                    // shared by all three headers.
                    let image_media_name =
                        if let Some(image_watermark) = mutable_doc.image_watermark.as_ref() {
                            let media_name = format!(
                                "/word/media/watermarkImage1.{}",
                                image_watermark.format().extension()
                            );
                            let media_uri = PackURI::new(&media_name).map_err(|e| {
                                Error::InvalidUri(format!("watermark image URI: {}", e))
                            })?;
                            self.opc.add_part(Box::new(BlobPart::new(
                                media_uri,
                                image_watermark.format().mime_type().to_string(),
                                image_watermark.data().to_vec(),
                            )));
                            Some(media_name)
                        } else {
                            None
                        };

                    // Create three headers (default, first, even) with watermark
                    let header_types = [
                        ("/word/header1.xml", "header1.xml"),
                        ("/word/header2.xml", "header2.xml"),
                        ("/word/header3.xml", "header3.xml"),
                    ];

                    for (idx, (header_uri_path, header_filename)) in header_types.iter().enumerate()
                    {
                        let mut watermark_xml = String::new();
                        if let Some(wm) = mutable_doc.watermark.as_ref() {
                            watermark_xml.push_str(&wm.to_header_xml((idx + 1) as u32)?);
                        }

                        let header_uri = PackURI::new(*header_uri_path)
                            .map_err(|e| Error::InvalidUri(format!("header URI: {}", e)))?;
                        let mut header_part =
                            BlobPart::new(header_uri, ct::WML_HEADER.to_string(), Vec::new());

                        // The image watermark references the media part
                        // through a relationship owned by this header part.
                        if let (Some(image_watermark), Some(media_name)) = (
                            mutable_doc.image_watermark.as_ref(),
                            image_media_name.as_deref(),
                        ) {
                            let media_target =
                                media_name.strip_prefix("/word/").unwrap_or(media_name);
                            let rel_id = header_part.relate_to(media_target, rt::IMAGE);
                            watermark_xml.push_str(
                                &image_watermark.to_header_xml((idx + 1) as u32, &rel_id)?,
                            );
                        }

                        // Merge user header content with watermark for the default header
                        let header_xml = if idx == 0
                            && let Some(ref user_content) = user_header_content
                        {
                            // Extract user paragraphs from the <w:hdr>...</w:hdr> wrapper
                            let user_paragraphs = if let Some(start) = user_content.find("<w:p") {
                                if let Some(end) = user_content.rfind("</w:hdr>") {
                                    &user_content[start..end]
                                } else {
                                    ""
                                }
                            } else {
                                ""
                            };

                            // Combine watermark and user content
                            format!(
                                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">{}{}</w:hdr>"#,
                                watermark_xml, user_paragraphs
                            )
                        } else {
                            // Just watermark for first and even headers
                            format!(
                                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">{}</w:hdr>"#,
                                watermark_xml
                            )
                        };

                        header_part.set_blob(header_xml.into_bytes());
                        self.opc.add_part(Box::new(header_part));

                        // Add relationship for the default header
                        if idx == 0 {
                            let rid = temp_part.relate_to(header_filename, rt::HEADER);
                            rel_mapper.set_header_id(rid);
                        } else {
                            // Other headers are added but not set in rel_mapper (they're referenced in sectPr)
                            temp_part.relate_to(header_filename, rt::HEADER);
                        }
                    }
                }

                // Step 3: Generate XML with actual relationship IDs
                let xml = mutable_doc.to_xml_with_rels(&rel_mapper)?;

                // Step 4: Update the document part with final XML and relationships
                for (header, part, uri, _) in planned_section_parts {
                    let content_type = if header {
                        ct::WML_HEADER
                    } else {
                        ct::WML_FOOTER
                    };
                    self.opc.add_part(Box::new(BlobPart::new(
                        uri,
                        content_type.to_string(),
                        part.xml.into_bytes(),
                    )));
                }
                temp_part.set_blob(xml.into_bytes());
                self.opc.add_part(Box::new(temp_part));

                // Note: Footnotes and endnotes are already handled above (before document XML generation)
                // so they appear in sectPr with proper relationship IDs

                // Update comments if present
                if let Some(comments_xml) = mutable_doc.generate_comments_xml()? {
                    self.update_comments_part(comments_xml)?;
                }

                // Patch only explicitly changed protection, preserving every other setting.
                if mutable_doc.protection_is_dirty() {
                    let settings_uri = PackURI::new("/word/settings.xml")
                        .map_err(|error| Error::InvalidUri(format!("settings URI: {error}")))?;
                    let existing_settings = self
                        .opc
                        .get_part(&settings_uri)
                        .ok()
                        .map(|part| part.blob().to_vec());
                    let settings_xml =
                        mutable_doc.generate_settings_xml(existing_settings.as_deref())?;
                    self.update_settings_part(settings_xml)?;
                }

                // Update theme if present
                if let Some(theme_xml) = mutable_doc.generate_theme_xml()? {
                    self.update_theme_part(theme_xml)?;
                }
            }

            // Update or remove the custom-properties package graph atomically.
            self.custom_props
                .write_for(&mut self.opc, CustomPropsHost::Word)?;

            // Embed fonts if feature enabled and requested in options
            #[cfg(feature = "fonts")]
            {
                if let Some(mutable_doc) = rollback.mutable_doc_mut() {
                    self.embed_fonts_for_document(mutable_doc)?;
                } else {
                    self.embed_fonts()?;
                }
            }

            // Stage only an explicitly edited core-properties slot. The guard
            // keeps edit intent dirty until the output sink accepts the complete
            // package, so a failed stream remains retryable.
            let staged_properties = self.properties.stage(&mut self.opc)?;

            self.opc.to_stream(writer).map_err(|e| {
                Error::Io(std::io::Error::other(format!(
                    "Failed to save package: {}",
                    e
                )))
            })?;
            staged_properties.commit();
            self.custom_props_dirty = false;
            Ok(())
        }));

        match result {
            Ok(Ok(())) => {
                rollback.publish(self);
                Ok(())
            },
            Ok(Err(error)) => {
                rollback.rollback(self);
                Err(error)
            },
            Err(payload) => {
                rollback.rollback(self);
                std::panic::resume_unwind(payload);
            },
        }
    }
}
