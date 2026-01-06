use crate::ooxml::common::DocumentProperties;
use crate::ooxml::custom_properties::CustomProperties;
use crate::ooxml::docx::document::Document;
use crate::ooxml::docx::parts::DocumentPart;
use crate::ooxml::docx::writer::MutableDocument;
/// Package implementation for Word documents.
use crate::ooxml::error::{OoxmlError, Result};
use crate::ooxml::opc::OpcPackage;
use crate::ooxml::opc::constants::content_type as ct;
use crate::ooxml::opc::packuri::PackURI;
use std::io::{Read, Seek};
use std::path::Path;

/// A Word (.docx) package.
///
/// This is the main entry point for working with Word documents.
/// It wraps an OPC package and provides Word-specific functionality.
///
/// # Examples
///
/// ## Reading an existing document
///
/// ```rust,no_run
/// use litchi::ooxml::docx::Package;
///
/// // Open an existing document
/// let pkg = Package::open("document.docx")?;
///
/// // Get the main document
/// let doc = pkg.document()?;
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
///
/// ## Creating a new document
///
/// ```rust,no_run
/// use litchi::ooxml::docx::Package;
///
/// // Create a new document
/// let mut pkg = Package::new()?;
/// let mut doc = pkg.document_mut()?;
///
/// // Add content
/// doc.add_paragraph_with_text("Hello, World!");
/// doc.add_heading("Chapter 1", 1)?;
///
/// // Save the document
/// pkg.save("output.docx")?;
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
pub struct Package {
    /// The underlying OPC package
    opc: OpcPackage,
    /// Mutable document for writing (cached)
    mutable_doc: Option<MutableDocument>,
    /// Document properties (metadata)
    properties: DocumentProperties,
    /// Custom document properties
    custom_properties: CustomProperties,
}

#[cfg(feature = "fonts")]
use crate::fonts::CollectGlyphs;
#[cfg(feature = "fonts")]
use crate::ooxml::fonts::{EmbedFonts, embed_fonts_in_package};
#[cfg(feature = "fonts")]
use roaring::RoaringBitmap;
#[cfg(feature = "fonts")]
use std::collections::HashMap;

#[cfg(feature = "fonts")]
impl CollectGlyphs for Package {
    fn collect_glyphs(&self) -> HashMap<String, RoaringBitmap> {
        if let Some(doc) = &self.mutable_doc {
            doc.collect_glyphs()
        } else if let Ok(_doc) = self.document() {
            // For now, only mutable documents support scanning as they are in-memory.
            // Future enhancement: parse document.xml part directly for glyphs.
            HashMap::new()
        } else {
            HashMap::new()
        }
    }
}

#[cfg(feature = "fonts")]
impl EmbedFonts for Package {
    fn embed_fonts(&mut self) -> Result<()> {
        let glyphs = self.collect_glyphs();
        let font_table_uri = PackURI::new("/word/fontTable.xml")
            .map_err(|e| OoxmlError::Other(format!("Invalid fontTable URI: {}", e)))?;

        // Embed fonts and get relationship IDs with fontKey
        let embedded_fonts =
            embed_fonts_in_package(glyphs, &mut self.opc, "/word/fonts", &font_table_uri)?;

        if embedded_fonts.is_empty() {
            return Ok(());
        }

        // Update settings.xml to include embedTrueTypeFonts flag
        let settings_uri = PackURI::new("/word/settings.xml")
            .map_err(|e| OoxmlError::Other(format!("Invalid settings URI: {}", e)))?;

        if let Ok(settings_part) = self.opc.get_part_mut(&settings_uri) {
            let xml_content = std::str::from_utf8(settings_part.blob())
                .map_err(|e| OoxmlError::Other(format!("Invalid settings.xml: {}", e)))?;

            // Check if embedTrueTypeFonts already exists
            if !xml_content.contains("<w:embedTrueTypeFonts") {
                let mut updated_xml = xml_content.to_string();

                // Insert after <w:settings> opening tag or before </w:settings>
                if let Some(pos) = updated_xml.find("</w:settings>") {
                    updated_xml.insert_str(pos, "<w:embedTrueTypeFonts/>");
                    settings_part.set_blob(updated_xml.into_bytes());
                }
            }
        }

        // Update fontTable.xml content with embedded font references
        if let Ok(font_table_part) = self.opc.get_part_mut(&font_table_uri) {
            let xml_content = std::str::from_utf8(font_table_part.blob())
                .map_err(|e| OoxmlError::Other(format!("Invalid fontTable.xml: {}", e)))?;

            let mut updated_xml = xml_content.to_string();

            for (font_name, info) in embedded_fonts {
                // Find the <w:font w:name="Font Name"> element
                let search_pattern = format!("w:name=\"{}\"", font_name);
                if let Some(pos) = updated_xml.find(&search_pattern) {
                    // Find the closing tag of this font entry or the next property
                    if let Some(font_end_pos) = updated_xml[pos..].find("</w:font>") {
                        let absolute_end_pos = pos + font_end_pos;
                        // Include w:fontKey attribute (GUID) - required for Office to recognize embedded fonts
                        let embed_xml = format!(
                            "<w:embedRegular r:id=\"{}\" w:fontKey=\"{}\"/>",
                            info.relationship_id, info.font_key
                        );
                        // Insert before </w:font>
                        updated_xml.insert_str(absolute_end_pos, &embed_xml);
                    }
                } else {
                    // Font not in table, append new entry before </w:fonts>
                    if let Some(fonts_end_pos) = updated_xml.rfind("</w:fonts>") {
                        let mut new_font_xml = format!("<w:font w:name=\"{}\">", font_name);

                        // Add font properties if available (required for Office recognition)
                        if let Some(ref props) = info.properties {
                            if let Some(ref panose) = props.panose {
                                new_font_xml
                                    .push_str(&format!("<w:panose1 w:val=\"{}\"/>", panose));
                            }
                            if let Some(ref charset) = props.charset {
                                new_font_xml
                                    .push_str(&format!("<w:charset w:val=\"{}\"/>", charset));
                            }
                            if let Some(ref family) = props.family {
                                new_font_xml.push_str(&format!("<w:family w:val=\"{}\"/>", family));
                            }
                            if let Some(ref pitch) = props.pitch {
                                new_font_xml.push_str(&format!("<w:pitch w:val=\"{}\"/>", pitch));
                            }
                            if let Some(ref sig) = props.sig {
                                new_font_xml.push_str(&format!(
                                    "<w:sig w:usb0=\"{}\" w:usb1=\"{}\" w:usb2=\"{}\" w:usb3=\"{}\" w:csb0=\"{}\" w:csb1=\"{}\"/>",
                                    sig.0, sig.1, sig.2, sig.3, sig.4, sig.5
                                ));
                            }
                        }

                        new_font_xml.push_str(&format!(
                            "<w:embedRegular r:id=\"{}\" w:fontKey=\"{}\"/></w:font>",
                            info.relationship_id, info.font_key
                        ));
                        updated_xml.insert_str(fonts_end_pos, &new_font_xml);
                    }
                }
            }

            font_table_part.set_blob(updated_xml.into_bytes());
        }

        Ok(())
    }
}

impl Package {
    /// Create a new empty .docx package.
    ///
    /// Creates a minimal valid Word document with default styles and settings.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::docx::Package;
    ///
    /// let pkg = Package::new()?;
    /// // Add content to the document...
    /// pkg.save("new_document.docx")?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn new() -> Result<Self> {
        use crate::ooxml::docx::template;
        use crate::ooxml::opc::constants::content_type as ct;
        use crate::ooxml::opc::constants::relationship_type as rt;
        use crate::ooxml::opc::packuri::PackURI;
        use crate::ooxml::opc::part::BlobPart;

        let mut opc = OpcPackage::new();

        // Create document.xml part
        let doc_partname = PackURI::new("/word/document.xml")
            .map_err(|e| OoxmlError::InvalidUri(format!("document partname: {}", e)))?;
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
            .map_err(|e| OoxmlError::InvalidUri(format!("styles partname: {}", e)))?;

        // Generate default styles dynamically
        use crate::ooxml::docx::writer::style::{MutableStyle, generate_styles_xml};
        let default_styles = vec![
            MutableStyle::normal(),
            MutableStyle::heading_1(),
            MutableStyle::heading_2(),
            MutableStyle::heading_3(),
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
            .map_err(|e| OoxmlError::InvalidUri(format!("settings partname: {}", e)))?;
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
            .map_err(|e| OoxmlError::InvalidUri(format!("fontTable partname: {}", e)))?;
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
            .map_err(|e| OoxmlError::InvalidUri(format!("webSettings partname: {}", e)))?;
        let web_settings_part = BlobPart::new(
            web_settings_partname,
            ct::WML_WEB_SETTINGS.to_string(),
            template::default_web_settings_xml().as_bytes().to_vec(),
        );

        if let Ok(doc_part) = opc.get_part_mut(&doc_partname) {
            doc_part.relate_to("webSettings.xml", rt::WEB_SETTINGS);
        }
        opc.add_part(Box::new(web_settings_part));

        // Create core.xml part (core properties)
        let core_props_partname = PackURI::new("/docProps/core.xml")
            .map_err(|e| OoxmlError::InvalidUri(format!("core.xml partname: {}", e)))?;
        let core_props_part = BlobPart::new(
            core_props_partname,
            ct::OPC_CORE_PROPERTIES.to_string(),
            template::default_core_props_xml().as_bytes().to_vec(),
        );

        opc.relate_to("docProps/core.xml", rt::CORE_PROPERTIES);
        opc.add_part(Box::new(core_props_part));

        // Create app.xml part (extended properties)
        let app_props_partname = PackURI::new("/docProps/app.xml")
            .map_err(|e| OoxmlError::InvalidUri(format!("app.xml partname: {}", e)))?;
        let app_props_part = BlobPart::new(
            app_props_partname,
            ct::OFC_EXTENDED_PROPERTIES.to_string(),
            template::default_app_props_xml().as_bytes().to_vec(),
        );

        opc.relate_to("docProps/app.xml", rt::EXTENDED_PROPERTIES);
        opc.add_part(Box::new(app_props_part));

        // Create theme1.xml part
        let theme_partname = PackURI::new("/word/theme/theme1.xml")
            .map_err(|e| OoxmlError::InvalidUri(format!("theme partname: {}", e)))?;
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
            .map_err(|e| OoxmlError::InvalidUri(format!("numbering partname: {}", e)))?;
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
        let properties = DocumentProperties::new();

        // Initialize custom properties
        let custom_properties = CustomProperties::new();

        Ok(Self {
            opc,
            mutable_doc,
            properties,
            custom_properties,
        })
    }

    /// Open a .docx package from a file path.
    ///
    /// # Arguments
    ///
    /// * `path` - Path to the .docx file
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self> {
        let opc = OpcPackage::open(path)?;

        // Verify it's a Word document by checking the main part's content type
        let main_part = opc
            .main_document_part()
            .map_err(|e| OoxmlError::PartNotFound(format!("main document part: {}", e)))?;

        let content_type = main_part.content_type();
        if content_type != ct::WML_DOCUMENT_MAIN {
            return Err(OoxmlError::InvalidContentType {
                expected: ct::WML_DOCUMENT_MAIN.to_string(),
                got: content_type.to_string(),
            });
        }

        // Try to extract custom properties
        let custom_properties = crate::ooxml::custom_properties::extract_custom_properties(&opc)
            .unwrap_or_else(|_| CustomProperties::new());

        Ok(Self {
            opc,
            mutable_doc: None,
            properties: DocumentProperties::new(),
            custom_properties,
        })
    }

    #[cfg(feature = "ooxml_encryption")]
    pub fn open_with_password<P: AsRef<Path>>(path: P, password: &str) -> Result<Self> {
        let data = std::fs::read(path.as_ref()).map_err(OoxmlError::Io)?;
        let decrypted = crate::ooxml::crypto::decrypt_ooxml_if_encrypted(&data, password)?;
        let opc = OpcPackage::from_bytes(&decrypted.package_bytes)?;
        Self::from_opc_package(opc)
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
    /// use litchi::ooxml::{OpcPackage, docx::Package};
    /// use std::io::Cursor;
    ///
    /// let bytes = std::fs::read("document.docx")?;
    /// let opc = OpcPackage::from_reader(Cursor::new(bytes))?;
    /// let pkg = Package::from_opc_package(opc)?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn from_opc_package(opc: OpcPackage) -> Result<Self> {
        // Verify it's a Word document by checking the main part's content type
        let main_part = opc
            .main_document_part()
            .map_err(|e| OoxmlError::PartNotFound(format!("main document part: {}", e)))?;

        let content_type = main_part.content_type();
        if content_type != ct::WML_DOCUMENT_MAIN {
            return Err(OoxmlError::InvalidContentType {
                expected: ct::WML_DOCUMENT_MAIN.to_string(),
                got: content_type.to_string(),
            });
        }

        // Try to extract custom properties
        let custom_properties = crate::ooxml::custom_properties::extract_custom_properties(&opc)
            .unwrap_or_else(|_| CustomProperties::new());

        Ok(Self {
            opc,
            mutable_doc: None,
            properties: DocumentProperties::new(),
            custom_properties,
        })
    }

    /// Create a .docx package from a reader.
    ///
    /// # Arguments
    ///
    /// * `reader` - A reader containing the .docx file data (must implement Read + Seek)
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::docx::Package;
    /// use std::io::Cursor;
    ///
    /// let data = std::fs::read("document.docx")?;
    /// let cursor = Cursor::new(data);
    /// let pkg = Package::from_reader(cursor)?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn from_reader<R: Read + Seek>(reader: R) -> Result<Self> {
        let opc = OpcPackage::from_reader(reader)?;

        // Verify it's a Word document by checking the main part's content type
        let main_part = opc
            .main_document_part()
            .map_err(|e| OoxmlError::PartNotFound(format!("main document part: {}", e)))?;

        let content_type = main_part.content_type();
        if content_type != ct::WML_DOCUMENT_MAIN {
            return Err(OoxmlError::InvalidContentType {
                expected: ct::WML_DOCUMENT_MAIN.to_string(),
                got: content_type.to_string(),
            });
        }

        // Try to extract custom properties
        let custom_properties = crate::ooxml::custom_properties::extract_custom_properties(&opc)
            .unwrap_or_else(|_| CustomProperties::new());

        Ok(Self {
            opc,
            mutable_doc: None,
            properties: DocumentProperties::new(),
            custom_properties,
        })
    }

    #[cfg(feature = "ooxml_encryption")]
    pub fn from_reader_with_password<R: Read + Seek>(
        mut reader: R,
        password: &str,
    ) -> Result<Self> {
        let mut data = Vec::new();
        reader.read_to_end(&mut data).map_err(OoxmlError::Io)?;
        let decrypted = crate::ooxml::crypto::decrypt_ooxml_if_encrypted(&data, password)?;
        let opc = OpcPackage::from_bytes(&decrypted.package_bytes)?;
        Self::from_opc_package(opc)
    }

    /// Get the main document for reading.
    ///
    /// Returns the `Document` object which provides access to the document's
    /// content, styles, and other features.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn document(&self) -> Result<Document<'_>> {
        let main_part = self
            .opc
            .main_document_part()
            .map_err(|e| OoxmlError::PartNotFound(format!("main document part: {}", e)))?;

        // Create DocumentPart wrapper
        let doc_part = DocumentPart::from_part(main_part)?;

        // Create and return Document with reference to OPC package
        Ok(Document::new(doc_part, &self.opc))
    }

    /// Get the underlying OPC package.
    ///
    /// This provides access to lower-level package operations.
    #[inline]
    pub fn opc_package(&self) -> &OpcPackage {
        &self.opc
    }

    /// Get mutable access to the underlying OPC package.
    ///
    /// This provides access to lower-level package operations for modification.
    #[inline]
    pub fn opc_package_mut(&mut self) -> &mut OpcPackage {
        &mut self.opc
    }

    /// Get a mutable document for writing and modification.
    ///
    /// This returns a `MutableDocument` that allows you to add and modify
    /// paragraphs, tables, and other document elements.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::docx::Package;
    ///
    /// let mut pkg = Package::new()?;
    /// let mut doc = pkg.document_mut()?;
    ///
    /// // Add content
    /// doc.add_paragraph_with_text("Hello, World!");
    /// let para = doc.add_paragraph();
    /// para.add_run_with_text("Bold text").bold(true);
    ///
    /// // Add a table
    /// let table = doc.add_table(3, 2);
    /// table.cell(0, 0).unwrap().set_text("Header 1");
    ///
    /// pkg.save("output.docx")?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn document_mut(&mut self) -> Result<&mut MutableDocument> {
        // If we don't have a mutable document, try to load it from the package
        if self.mutable_doc.is_none() {
            let doc_uri = PackURI::new("/word/document.xml")
                .map_err(|e| OoxmlError::InvalidUri(format!("document URI: {}", e)))?;

            // Try to get existing document content
            if let Ok(part) = self.opc.get_part(&doc_uri) {
                let xml = std::str::from_utf8(part.blob())
                    .map_err(|e| OoxmlError::InvalidFormat(format!("Invalid UTF-8: {}", e)))?;
                self.mutable_doc = Some(MutableDocument::from_xml(xml)?);
            } else {
                // Create a new empty document
                self.mutable_doc = Some(MutableDocument::new());
            }
        }

        Ok(self.mutable_doc.as_mut().unwrap())
    }

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
    /// use litchi::ooxml::docx::Package;
    ///
    /// let mut pkg = Package::new()?;
    /// // Modify document...
    /// pkg.save("output.docx")?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn save<P: AsRef<Path>>(&mut self, path: P) -> Result<()> {
        use crate::ooxml::docx::writer::relmap::RelationshipMapper;
        use crate::ooxml::opc::constants::relationship_type as rt;

        // If we have a mutable document, update the document.xml part
        if let Some(mut mutable_doc) = self.mutable_doc.take() {
            if mutable_doc.is_modified() {
                // Generate TOC if configured (must happen before serialization)
                mutable_doc.generate_toc_if_needed()?;

                // Step 1: Collect all content that needs relationships
                let hyperlink_urls = mutable_doc.collect_hyperlink_urls();
                let images = mutable_doc.collect_images();
                let has_header = mutable_doc.has_header();
                let has_footer = mutable_doc.has_footer();

                // Step 2: Create a relationship mapper and add relationships
                let mut rel_mapper = RelationshipMapper::new();

                // Create the document part first (we'll update it later)
                let doc_uri = PackURI::new("/word/document.xml")
                    .map_err(|e| OoxmlError::InvalidUri(format!("document URI: {}", e)))?;

                // Get or create the document part to add relationships to
                let content_type = self
                    .opc
                    .get_part(&doc_uri)
                    .map(|p| p.content_type().to_string())
                    .unwrap_or_else(|_| ct::WML_DOCUMENT_MAIN.to_string());

                // Create new temporary part for relationships
                use crate::ooxml::opc::part::{BlobPart, Part};
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
                                | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
                                | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
                                | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
                                | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes"
                        ) {
                            if rel.is_external() {
                                temp_part.relate_to_ext(rel.target_ref(), rel.reltype());
                            } else {
                                temp_part.relate_to(rel.target_ref(), rel.reltype());
                            }
                        }
                    }
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
                        .map_err(|e| OoxmlError::InvalidUri(format!("image URI: {}", e)))?;

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

                // Add header/footer parts and relationships
                // Note: If watermark exists, headers will be handled by update_watermark_headers
                // which merges user content with watermark
                if has_header
                    && !mutable_doc.has_watermark()
                    && let Some(header_xml) = mutable_doc.generate_header_xml()?
                {
                    let header_uri = PackURI::new("/word/header1.xml")
                        .map_err(|e| OoxmlError::InvalidUri(format!("header URI: {}", e)))?;
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
                        .map_err(|e| OoxmlError::InvalidUri(format!("footer URI: {}", e)))?;
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
                        .map_err(|e| OoxmlError::InvalidUri(format!("footnotes URI: {}", e)))?;
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
                        .map_err(|e| OoxmlError::InvalidUri(format!("endnotes URI: {}", e)))?;
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
                if mutable_doc.has_watermark() {
                    // Generate user header content if exists (will be merged with watermark)
                    let user_header_content = if mutable_doc.has_header() {
                        mutable_doc.generate_header_xml()?
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
                        if let Some(wm) = mutable_doc.watermark.as_ref() {
                            let watermark_xml = wm.to_header_xml((idx + 1) as u32)?;

                            // Merge user header content with watermark for the default header
                            let header_xml = if idx == 0
                                && let Some(ref user_content) = user_header_content
                            {
                                // Extract user paragraphs from the <w:hdr>...</w:hdr> wrapper
                                let user_paragraphs = if let Some(start) = user_content.find("<w:p")
                                {
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
                                    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">{}{}</w:hdr>"#,
                                    watermark_xml, user_paragraphs
                                )
                            } else {
                                // Just watermark for first and even headers
                                format!(
                                    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">{}</w:hdr>"#,
                                    watermark_xml
                                )
                            };

                            let header_uri = PackURI::new(*header_uri_path).map_err(|e| {
                                OoxmlError::InvalidUri(format!("header URI: {}", e))
                            })?;

                            let header_part = BlobPart::new(
                                header_uri,
                                ct::WML_HEADER.to_string(),
                                header_xml.into_bytes(),
                            );

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
                }

                // Step 3: Generate XML with actual relationship IDs
                let xml = mutable_doc.to_xml_with_rels(&rel_mapper)?;

                // Step 4: Update the document part with final XML and relationships
                temp_part.set_blob(xml.into_bytes());
                self.opc.add_part(Box::new(temp_part));

                // Note: Footnotes and endnotes are already handled above (before document XML generation)
                // so they appear in sectPr with proper relationship IDs

                // Update comments if present
                if let Some(comments_xml) = mutable_doc.generate_comments_xml()? {
                    self.update_comments_part(comments_xml)?;
                }

                // Update settings.xml with protection if modified
                let settings_xml = mutable_doc.generate_settings_xml()?;
                self.update_settings_part(settings_xml)?;

                // Update theme if present
                if let Some(theme_xml) = mutable_doc.generate_theme_xml()? {
                    self.update_theme_part(theme_xml)?;
                }
            }
            // Put the document back
            self.mutable_doc = Some(mutable_doc);
        }

        // Update core properties
        self.update_core_properties()?;

        // Update custom properties
        self.update_custom_properties()?;

        // Embed fonts if feature enabled and requested in options
        #[cfg(feature = "fonts")]
        {
            self.embed_fonts()?;
        }

        self.opc.save(path).map_err(|e| {
            OoxmlError::IoError(std::io::Error::other(format!(
                "Failed to save package: {}",
                e
            )))
        })
    }

    /// Get a reference to the document properties.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let props = pkg.properties();
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn properties(&self) -> &DocumentProperties {
        &self.properties
    }

    /// Get a mutable reference to the document properties.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::docx::Package;
    ///
    /// let mut pkg = Package::new()?;
    /// pkg.properties_mut().title = Some("My Document".to_string());
    /// pkg.properties_mut().creator = Some("John Doe".to_string());
    /// pkg.save("document.docx")?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn properties_mut(&mut self) -> &mut DocumentProperties {
        &mut self.properties
    }

    /// Get a reference to the custom document properties.
    ///
    /// Custom properties allow you to attach arbitrary typed metadata to documents.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let custom_props = pkg.custom_properties();
    ///
    /// if let Some(value) = custom_props.get_property("ProjectName") {
    ///     println!("Project: {:?}", value);
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn custom_properties(&self) -> &CustomProperties {
        &self.custom_properties
    }

    /// Get a mutable reference to the custom document properties.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::docx::Package;
    /// use litchi::ooxml::custom_properties::PropertyValue;
    ///
    /// let mut pkg = Package::new()?;
    /// let custom_props = pkg.custom_properties_mut();
    ///
    /// custom_props.add_property("ProjectName", PropertyValue::String("MyProject".to_string()));
    /// custom_props.add_property("Version", PropertyValue::Integer(1));
    /// custom_props.add_property("Budget", PropertyValue::Double(50000.0));
    ///
    /// pkg.save("document.docx")?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn custom_properties_mut(&mut self) -> &mut CustomProperties {
        &mut self.custom_properties
    }

    /// Update the core.xml properties part.
    fn update_core_properties(&mut self) -> Result<()> {
        use crate::ooxml::opc::part::BlobPart;

        let core_uri = PackURI::new("/docProps/core.xml")
            .map_err(|e| OoxmlError::InvalidUri(format!("core.xml URI: {}", e)))?;

        // Generate XML from properties
        let xml = self.properties.to_xml();

        // Create or update the core properties part
        let core_part = BlobPart::new(
            core_uri,
            ct::OPC_CORE_PROPERTIES.to_string(),
            xml.into_bytes(),
        );

        self.opc.add_part(Box::new(core_part));

        Ok(())
    }

    /// Update the custom.xml properties part.
    fn update_custom_properties(&mut self) -> Result<()> {
        use crate::ooxml::opc::constants::relationship_type as rt;
        use crate::ooxml::opc::part::BlobPart;

        // Only create custom properties part if there are custom properties
        if self.custom_properties.is_empty() {
            return Ok(());
        }

        let custom_uri = PackURI::new("/docProps/custom.xml")
            .map_err(|e| OoxmlError::InvalidUri(format!("custom.xml URI: {}", e)))?;

        // Generate XML from custom properties
        let xml = self.custom_properties.to_xml()?;

        // Create or update the custom properties part
        let custom_part = BlobPart::new(
            custom_uri.clone(),
            ct::OFC_CUSTOM_PROPERTIES.to_string(),
            xml.into_bytes(),
        );

        self.opc.add_part(Box::new(custom_part));

        // Ensure relationship exists
        self.opc
            .relate_to("docProps/custom.xml", rt::CUSTOM_PROPERTIES);

        Ok(())
    }

    /// Update the footnotes.xml part with new content.
    #[allow(unused)] // Kept for future use
    fn update_footnotes_part(&mut self, xml: String) -> Result<()> {
        use crate::ooxml::opc::constants::content_type as ct;
        use crate::ooxml::opc::constants::relationship_type as rt;
        use crate::ooxml::opc::part::BlobPart;

        let footnotes_uri = PackURI::new("/word/footnotes.xml")
            .map_err(|e| OoxmlError::InvalidUri(format!("footnotes URI: {}", e)))?;

        let content_type = ct::WML_FOOTNOTES.to_string();
        let footnotes_part = BlobPart::new(footnotes_uri.clone(), content_type, xml.into_bytes());

        // Add the footnotes part
        self.opc.add_part(Box::new(footnotes_part));

        // Create relationship from document to footnotes (use relative path)
        let doc_uri = PackURI::new("/word/document.xml")
            .map_err(|e| OoxmlError::InvalidUri(format!("document URI: {}", e)))?;

        if let Ok(doc_part) = self.opc.get_part_mut(&doc_uri) {
            let _ = doc_part.relate_to("footnotes.xml", rt::FOOTNOTES);
        }

        Ok(())
    }

    /// Update the endnotes.xml part with new content.
    #[allow(unused)] // Kept for future use
    fn update_endnotes_part(&mut self, xml: String) -> Result<()> {
        use crate::ooxml::opc::constants::content_type as ct;
        use crate::ooxml::opc::constants::relationship_type as rt;
        use crate::ooxml::opc::part::BlobPart;

        let endnotes_uri = PackURI::new("/word/endnotes.xml")
            .map_err(|e| OoxmlError::InvalidUri(format!("endnotes URI: {}", e)))?;

        let content_type = ct::WML_ENDNOTES.to_string();
        let endnotes_part = BlobPart::new(endnotes_uri.clone(), content_type, xml.into_bytes());

        // Add the endnotes part
        self.opc.add_part(Box::new(endnotes_part));

        // Create relationship from document to endnotes (use relative path)
        let doc_uri = PackURI::new("/word/document.xml")
            .map_err(|e| OoxmlError::InvalidUri(format!("document URI: {}", e)))?;

        if let Ok(doc_part) = self.opc.get_part_mut(&doc_uri) {
            let _ = doc_part.relate_to("endnotes.xml", rt::ENDNOTES);
        }

        Ok(())
    }

    /// Update or create the comments part with the given XML content.
    fn update_comments_part(&mut self, xml: String) -> Result<()> {
        use crate::ooxml::opc::constants::content_type as ct;
        use crate::ooxml::opc::constants::relationship_type as rt;
        use crate::ooxml::opc::part::BlobPart;

        let comments_uri = PackURI::new("/word/comments.xml")
            .map_err(|e| OoxmlError::InvalidUri(format!("comments URI: {}", e)))?;

        let content_type = ct::WML_COMMENTS.to_string();
        let comments_part = BlobPart::new(comments_uri.clone(), content_type, xml.into_bytes());

        // Add the comments part
        self.opc.add_part(Box::new(comments_part));

        // Create relationship from document to comments
        let doc_uri = PackURI::new("/word/document.xml")
            .map_err(|e| OoxmlError::InvalidUri(format!("document URI: {}", e)))?;

        if let Ok(doc_part) = self.opc.get_part_mut(&doc_uri) {
            let _ = doc_part.relate_to("/word/comments.xml", rt::COMMENTS);
        }

        Ok(())
    }

    /// Update the settings.xml part with new content.
    fn update_settings_part(&mut self, xml: String) -> Result<()> {
        use crate::ooxml::opc::constants::content_type as ct;
        use crate::ooxml::opc::part::BlobPart;

        let settings_uri = PackURI::new("/word/settings.xml")
            .map_err(|e| OoxmlError::InvalidUri(format!("settings URI: {}", e)))?;

        let content_type = ct::WML_SETTINGS.to_string();
        let settings_part = BlobPart::new(settings_uri, content_type, xml.into_bytes());

        // Add/replace the settings part
        self.opc.add_part(Box::new(settings_part));

        Ok(())
    }

    fn update_theme_part(&mut self, xml: String) -> Result<()> {
        use crate::ooxml::opc::part::BlobPart;

        let theme_uri = PackURI::new("/word/theme/theme1.xml")
            .map_err(|e| OoxmlError::InvalidUri(format!("theme URI: {}", e)))?;

        let content_type = "application/vnd.openxmlformats-officedocument.theme+xml".to_string();
        let theme_part = BlobPart::new(theme_uri.clone(), content_type, xml.into_bytes());

        // Add/replace the theme part
        self.opc.add_part(Box::new(theme_part));

        // Add relationship from document to theme if not exists
        let doc_uri = PackURI::new("/word/document.xml")
            .map_err(|e| OoxmlError::InvalidUri(format!("document URI: {}", e)))?;

        if let Ok(doc_part) = self.opc.get_part_mut(&doc_uri) {
            // Check if theme relationship already exists
            let has_theme_rel = doc_part.rels().iter().any(|rel| {
                rel.reltype()
                    == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
            });

            if !has_theme_rel {
                doc_part.relate_to(
                    "theme/theme1.xml",
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
                );
            }
        }

        Ok(())
    }

    #[allow(unused)] // Kept for future use
    fn update_watermark_headers(
        &mut self,
        mutable_doc: &crate::ooxml::docx::writer::MutableDocument,
    ) -> Result<()> {
        use crate::ooxml::opc::constants::content_type as ct;
        use crate::ooxml::opc::constants::relationship_type as rt;
        use crate::ooxml::opc::part::BlobPart;

        // Get watermark if present
        // Access watermark through a temporary reference
        let has_watermark = mutable_doc.has_watermark();
        if !has_watermark {
            return Ok(());
        }

        // Get user header content if it exists
        let user_header_content = if mutable_doc.has_header() {
            mutable_doc.generate_header_xml()?
        } else {
            None
        };

        // Create three headers (default, first, even) with watermark
        let header_types = [
            ("/word/header1.xml", "default"),
            ("/word/header2.xml", "first"),
            ("/word/header3.xml", "even"),
        ];

        let doc_uri = PackURI::new("/word/document.xml")
            .map_err(|e| OoxmlError::InvalidUri(format!("document URI: {}", e)))?;

        for (idx, (header_path, _header_type)) in header_types.iter().enumerate() {
            // Generate watermark XML for this header - need to get watermark again each iteration
            let watermark_xml = if let Some(wm) = mutable_doc.watermark.as_ref() {
                wm.to_header_xml((idx + 1) as u32)?
            } else {
                continue;
            };

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
                    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">{}{}</w:hdr>"#,
                    watermark_xml, user_paragraphs
                )
            } else {
                // Just watermark for first and even headers
                format!(
                    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">{}</w:hdr>"#,
                    watermark_xml
                )
            };

            let header_uri = PackURI::new(*header_path)
                .map_err(|e| OoxmlError::InvalidUri(format!("header URI: {}", e)))?;

            let header_part = BlobPart::new(
                header_uri,
                ct::WML_HEADER.to_string(),
                header_xml.into_bytes(),
            );

            self.opc.add_part(Box::new(header_part));

            // Add relationship from document to header (use relative path)
            // Extract filename from the absolute path (e.g., "/word/header1.xml" -> "header1.xml")
            let header_filename = header_path.rsplit('/').next().unwrap_or(header_path);
            if let Ok(doc_part) = self.opc.get_part_mut(&doc_uri) {
                doc_part.relate_to(header_filename, rt::HEADER);
            }
        }

        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    #[ignore] // Requires test file
    fn test_open_package() {
        let result = Package::open("test.docx");
        assert!(result.is_ok());
    }
}
