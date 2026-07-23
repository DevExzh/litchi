use crate::common::DocumentProperties;
use crate::custom_properties::CustomProperties;
use crate::custom_xml_data::{
    CustomXmlDataItem, CustomXmlDataProperties, NewCustomXmlDataItem, add_custom_xml_data,
    discover_custom_xml_data, validate_custom_xml_content_type, validate_custom_xml_payload,
    validate_custom_xml_properties, write_custom_xml_properties,
};
use crate::docx::alt_chunk::{
    AltChunk, AltChunkNamespace, AlternativeFormatImport, STRICT_ALTERNATIVE_FORMAT_IMPORT,
    is_alternative_format_relationship,
};
use crate::docx::bibliography::{
    BibliographySource, BibliographySourceStore, discover_bibliography_source_stores,
};
use crate::docx::document::Document;
use crate::docx::mail_merge::{
    MailMergeConformance, MailMergeRecipients, MailMergeSettings, MailMergeSource,
    MailMergeTarget,
};
use crate::docx::content_control::ContentControl;
use crate::docx::custom_xml::{CustomXmlBinding, NewCustomXmlDataStore};
use crate::docx::font_table::{FontTable, is_font_table_relationship};
use crate::docx::glossary::{
    GlossaryDocument, GlossaryEntry, GlossaryPackage, load_from_package,
    load_package_from_package, remove_from_package, store_in_package,
};
use crate::docx::parts::DocumentPart;
use crate::docx::settings::{
    ATTACHED_TEMPLATE_RELATIONSHIP, AttachedTemplate, DocumentSettings,
    patch_attached_template, patch_document_variables, patch_mail_merge,
    validate_attached_template_target,
};
use crate::docx::variables::DocumentVariables;
use crate::docx::vba_project::{VbaProject, discover_vba_project};
use crate::docx::web_settings::{WebSettings, is_web_settings_relationship};
use crate::docx::writer::MutableDocument;
/// Package implementation for Word documents.
use crate::error::{OoxmlError, Result};
use litchi_opc::OpcPackage;
use litchi_opc::constants::content_type as ct;
use litchi_opc::packuri::PackURI;
use litchi_opc::part::{BlobPart, Part};
use litchi_opc::rel::TargetMode;
use std::io::{Read, Seek, Write};
use std::path::Path;

fn validate_document_main_content_type(content_type: &str) -> Result<()> {
    if matches!(
        content_type,
        ct::WML_DOCUMENT_MAIN | ct::WML_DOCUMENT_MACRO_MAIN | ct::WML_TEMPLATE_MACRO_MAIN
    ) {
        return Ok(());
    }

    Err(OoxmlError::InvalidContentType {
        expected: format!(
            "{}, {}, or {}",
            ct::WML_DOCUMENT_MAIN,
            ct::WML_DOCUMENT_MACRO_MAIN,
            ct::WML_TEMPLATE_MACRO_MAIN,
        ),
        got: content_type.to_string(),
    })
}

/// A WordprocessingML (.docx, .docm, or .dotm) package.
///
/// This is the main entry point for working with Word documents.
/// It wraps an OPC package and provides Word-specific functionality.
///
/// # Examples
///
/// ## Reading an existing document
///
/// ```rust,no_run
/// use litchi_ooxml::docx::Package;
///
/// // Open an existing document
/// let pkg = Package::open("document.docx")?;
///
/// // Get the main document
/// let doc = pkg.document()?;
/// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
/// ```
///
/// ## Creating a new document
///
/// ```rust,no_run
/// use litchi_ooxml::docx::Package;
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
/// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
/// ```
pub struct Package {
    /// The underlying OPC package
    opc: OpcPackage,
    /// Mutable document for writing (cached)
    mutable_doc: Option<MutableDocument>,
    /// Lazily loaded web-output settings for explicit edits.
    mutable_web_settings: Option<WebSettings>,
    /// Whether the web-settings part must be rewritten.
    web_settings_dirty: bool,
    /// Document properties (metadata)
    properties: DocumentProperties,
    /// Custom document properties
    custom_properties: CustomProperties,
}

struct StoredRelationship {
    reltype: String,
    target: String,
    id: String,
    external: bool,
}

struct SettingsPartSnapshot {
    document_uri: PackURI,
    target: PackURI,
    relationship_exists: bool,
    content_type: String,
    xml: Vec<u8>,
    relationships: Vec<StoredRelationship>,
}

#[cfg(feature = "fonts")]
use crate::fonts::{EmbedFonts, embed_fonts_in_package};
#[cfg(feature = "fonts")]
use litchi_fonts::CollectGlyphs;
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
    /// use litchi_ooxml::docx::Package;
    ///
    /// let mut pkg = Package::new()?;
    /// // Add content to the document...
    /// pkg.save("new_document.docx")?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn new() -> Result<Self> {
        use crate::docx::template;
        use litchi_opc::constants::content_type as ct;
        use litchi_opc::constants::relationship_type as rt;
        use litchi_opc::packuri::PackURI;
        use litchi_opc::part::BlobPart;

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
        use crate::docx::writer::style::{MutableStyle, generate_styles_xml};
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
            mutable_web_settings: None,
            web_settings_dirty: false,
            properties,
            custom_properties,
        })
    }

    /// Open a .docx, .docm, or .dotm package from a file path.
    ///
    /// # Arguments
    ///
    /// * `path` - Path to the .docx file
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self> {
        let opc = OpcPackage::open(path)?;

        // Verify it's a Word document by checking the main part's content type
        let main_part = opc
            .main_document_part()
            .map_err(|e| OoxmlError::PartNotFound(format!("main document part: {}", e)))?;

        validate_document_main_content_type(main_part.content_type())?;

        // Try to extract custom properties
        let custom_properties = crate::custom_properties::extract_custom_properties(&opc)
            .unwrap_or_else(|_| CustomProperties::new());

        Ok(Self {
            opc,
            mutable_doc: None,
            mutable_web_settings: None,
            web_settings_dirty: false,
            properties: DocumentProperties::new(),
            custom_properties,
        })
    }

    #[cfg(feature = "encryption")]
    pub fn open_with_password<P: AsRef<Path>>(path: P, password: &str) -> Result<Self> {
        let data = std::fs::read(path.as_ref()).map_err(OoxmlError::Io)?;
        let decrypted = crate::crypto::decrypt_ooxml_if_encrypted(&data, password)?;
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
    /// use litchi_ooxml::{OpcPackage, docx::Package};
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
            .map_err(|e| OoxmlError::PartNotFound(format!("main document part: {}", e)))?;

        validate_document_main_content_type(main_part.content_type())?;

        // Try to extract custom properties
        let custom_properties = crate::custom_properties::extract_custom_properties(&opc)
            .unwrap_or_else(|_| CustomProperties::new());

        Ok(Self {
            opc,
            mutable_doc: None,
            mutable_web_settings: None,
            web_settings_dirty: false,
            properties: DocumentProperties::new(),
            custom_properties,
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
    /// use litchi_ooxml::docx::Package;
    /// use std::io::Cursor;
    ///
    /// let data = std::fs::read("document.docx")?;
    /// let cursor = Cursor::new(data);
    /// let pkg = Package::from_reader(cursor)?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn from_reader<R: Read + Seek>(reader: R) -> Result<Self> {
        let opc = OpcPackage::from_reader(reader)?;

        // Verify it's a Word document by checking the main part's content type
        let main_part = opc
            .main_document_part()
            .map_err(|e| OoxmlError::PartNotFound(format!("main document part: {}", e)))?;

        validate_document_main_content_type(main_part.content_type())?;

        // Try to extract custom properties
        let custom_properties = crate::custom_properties::extract_custom_properties(&opc)
            .unwrap_or_else(|_| CustomProperties::new());

        Ok(Self {
            opc,
            mutable_doc: None,
            mutable_web_settings: None,
            web_settings_dirty: false,
            properties: DocumentProperties::new(),
            custom_properties,
        })
    }

    #[cfg(feature = "encryption")]
    pub fn from_reader_with_password<R: Read + Seek>(
        mut reader: R,
        password: &str,
    ) -> Result<Self> {
        let mut data = Vec::new();
        reader.read_to_end(&mut data).map_err(OoxmlError::Io)?;
        let decrypted = crate::crypto::decrypt_ooxml_if_encrypted(&data, password)?;
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
    /// use litchi_ooxml::docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
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

    /// Discover the attached MS-OFFMACRO2 VBA project without inspecting its payloads.
    ///
    /// This validates only the declared OPC relationship graph and content
    /// types. It does not inspect, parse, decompress, or execute the binary
    /// VBA project or Word supplemental-data bytes.
    pub fn vba_project(&self) -> Result<Option<VbaProject>> {
        let document = self.opc.main_document_part()?;
        discover_vba_project(&self.opc, document)
    }

    /// Load typed font metadata and inert embedded-font resources.
    pub fn font_table(&self) -> Result<Option<FontTable>> {
        let main_part = self.opc.main_document_part()?;
        let mut matches = main_part
            .rels()
            .iter()
            .filter(|relationship| is_font_table_relationship(relationship.reltype()));
        let Some(relationship) = matches.next() else {
            return Ok(None);
        };
        if matches.next().is_some() {
            return Err(OoxmlError::InvalidFormat(
                "document has multiple font-table relationships".into(),
            ));
        }
        if relationship.is_external() {
            return Err(OoxmlError::InvalidFormat(
                "font-table relationship cannot be external".into(),
            ));
        }
        let target = relationship.target_partname()?;
        let part = self.opc.get_part(&target)?;
        Ok(Some(FontTable::extract_from_part(part, &self.opc)?))
    }

    /// Get the underlying OPC package.
    ///
    /// This provides access to lower-level package operations.
    #[inline]
    pub fn opc_package(&self) -> &OpcPackage {
        &self.opc
    }

    /// Verify package signatures without making a PKI trust determination.
    pub fn verify_digital_signatures(
        &self,
        policy: &litchi_opc::SignatureVerificationPolicy,
    ) -> litchi_opc::signature::Result<Vec<litchi_opc::DigitalSignatureVerification>> {
        self.opc.verify_digital_signatures(policy)
    }

    /// Sign the current, fully materialized package while preserving valid signatures.
    pub fn add_digital_signature(
        &mut self,
        signer: &litchi_opc::PackageSigner,
    ) -> litchi_opc::signature::Result<PackURI> {
        self.opc.add_digital_signature(signer)
    }

    /// Replace all package signatures with one new signature.
    pub fn resign_digital_signature(
        &mut self,
        signer: &litchi_opc::PackageSigner,
    ) -> litchi_opc::signature::Result<PackURI> {
        self.opc.resign_digital_signature(signer)
    }

    /// Remove all package digital signatures.
    pub fn clear_digital_signatures(&mut self) -> litchi_opc::signature::Result<()> {
        self.opc.clear_digital_signatures()
    }

    /// Discover inert embedded-object and embedded-package relationships.
    pub fn embedded_parts(&self) -> Result<Vec<crate::EmbeddedPart<'_>>> {
        crate::embedded_object::discover_embedded_parts(&self.opc)
    }

    /// Load the bounded, inert classic-chart graph owned by the main document.
    pub fn chart_graph(&self) -> Result<crate::docx::chart::DocxChartGraph> {
        let document = self.opc.main_document_part()?.partname().clone();
        crate::docx::chart::load_chart_graph(&self.opc, &document)
    }

    /// Deterministically store an already coherent classic-chart graph.
    pub fn store_chart_graph(&mut self, graph: &crate::docx::chart::DocxChartGraph) -> Result<()> {
        let document = self.opc.main_document_part()?.partname().clone();
        crate::docx::chart::store_chart_graph(&mut self.opc, &document, graph)
    }

    /// Get mutable access to the underlying OPC package.
    ///
    /// This provides access to lower-level package operations for modification.
    #[inline]
    pub fn opc_package_mut(&mut self) -> &mut OpcPackage {
        let _ = self.opc.clear_digital_signatures();
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
    /// use litchi_ooxml::docx::Package;
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
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
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

    /// Append a package-backed alternative-format import to the document body.
    pub fn add_alt_chunk(
        &mut self,
        import: AlternativeFormatImport,
        match_source: Option<bool>,
    ) -> Result<AltChunk> {
        let index = self.document_mut()?.alt_chunks().len();
        self.insert_alt_chunk(index, import, match_source)
    }

    /// Insert a package-backed alternative-format import by anchor-relative index.
    ///
    /// Part, relationship, and body mutations are rolled back together on error.
    pub fn insert_alt_chunk(
        &mut self,
        index: usize,
        import: AlternativeFormatImport,
        match_source: Option<bool>,
    ) -> Result<AltChunk> {
        let count = self.document_mut()?.alt_chunks().len();
        if index > count {
            return Err(OoxmlError::InvalidFormat(format!(
                "altChunk index {index} is out of range"
            )));
        }
        let namespace = self.alt_chunk_namespace()?;
        let (chunk, installed_part) = self.install_alt_chunk_target(import, match_source, namespace)?;
        if let Err(error) = self
            .document_mut()?
            .insert_alt_chunk(index, chunk.clone(), namespace)
        {
            self.rollback_alt_chunk_target(chunk.relationship_id(), installed_part.as_ref())?;
            return Err(error);
        }
        Ok(chunk)
    }

    /// Replace an anchor and its relationship as one package mutation.
    pub fn replace_alt_chunk(
        &mut self,
        index: usize,
        import: AlternativeFormatImport,
        match_source: Option<bool>,
    ) -> Result<AltChunk> {
        let old = self
            .document_mut()?
            .alt_chunks()
            .get(index)
            .cloned()
            .ok_or_else(|| {
                OoxmlError::InvalidFormat(format!("altChunk index {index} is out of range"))
            })?;
        let old_target = self.validate_alt_chunk_relationship(&old)?;
        let namespace = self.alt_chunk_namespace()?;
        let (new, installed_part) = self.install_alt_chunk_target(import, match_source, namespace)?;
        if let Err(error) = self
            .document_mut()?
            .replace_alt_chunk(index, new.clone(), namespace)
        {
            self.rollback_alt_chunk_target(new.relationship_id(), installed_part.as_ref())?;
            return Err(error);
        }
        self.remove_alt_chunk_relationship(&old, old_target.as_ref())?;
        Ok(old)
    }

    /// Remove an anchor, its relationship, and an unreachable internal payload.
    pub fn remove_alt_chunk(&mut self, index: usize) -> Result<AltChunk> {
        let old = self
            .document_mut()?
            .alt_chunks()
            .get(index)
            .cloned()
            .ok_or_else(|| {
                OoxmlError::InvalidFormat(format!("altChunk index {index} is out of range"))
            })?;
        let old_target = self.validate_alt_chunk_relationship(&old)?;
        self.document_mut()?.remove_alt_chunk(index)?;
        self.remove_alt_chunk_relationship(&old, old_target.as_ref())?;
        Ok(old)
    }

    /// Reorder body anchors without changing their package relationships.
    pub fn move_alt_chunk(&mut self, from: usize, to: usize) -> Result<()> {
        self.document_mut()?.move_alt_chunk(from, to)
    }

    fn alt_chunk_namespace(&self) -> Result<AltChunkNamespace> {
        let document_uri = PackURI::new("/word/document.xml")
            .map_err(|error| OoxmlError::InvalidUri(format!("document URI: {error}")))?;
        let strict = self
            .opc
            .get_part(&document_uri)
            .map(|part| {
                part.blob()
                    .windows(b"http://purl.oclc.org/ooxml/wordprocessingml/main".len())
                    .any(|window| window == b"http://purl.oclc.org/ooxml/wordprocessingml/main")
            })
            .unwrap_or(false);
        Ok(if strict {
            AltChunkNamespace::Strict
        } else {
            AltChunkNamespace::Transitional
        })
    }

    fn install_alt_chunk_target(
        &mut self,
        import: AlternativeFormatImport,
        match_source: Option<bool>,
        namespace: AltChunkNamespace,
    ) -> Result<(AltChunk, Option<PackURI>)> {
        let document_uri = PackURI::new("/word/document.xml")
            .map_err(|error| OoxmlError::InvalidUri(format!("document URI: {error}")))?;
        let document = self.opc.get_part(&document_uri)?;
        let relationship_id = (1usize..)
            .map(|number| format!("rIdAltChunk{number}"))
            .find(|id| document.rels().get(id).is_none())
            .expect("the relationship ID space is unbounded");
        let relationship_type = match namespace {
            AltChunkNamespace::Transitional => {
                litchi_opc::constants::relationship_type::MS_ALTERNATIVE_FORMAT_IMPORT
            },
            AltChunkNamespace::Strict => STRICT_ALTERNATIVE_FORMAT_IMPORT,
        };
        let (target_ref, target_mode, installed_part) = match import {
            AlternativeFormatImport::External(uri) => {
                if uri.is_empty() || uri.len() > 32_768 || uri.chars().any(char::is_control) {
                    return Err(OoxmlError::InvalidFormat(
                        "external altChunk target is empty or unsafe".to_string(),
                    ));
                }
                (uri, TargetMode::External, None)
            },
            AlternativeFormatImport::Internal(data) => {
                if data.bytes().len() > 128 * 1024 * 1024 {
                    return Err(OoxmlError::InvalidFormat(
                        "alternative-format part exceeds the 128 MiB authoring limit".to_string(),
                    ));
                }
                let (uri, target_ref) = (1usize..)
                    .find_map(|number| {
                        let target_ref = format!("afchunk{number}.{}", data.extension());
                        let uri = PackURI::new(&format!("/word/{target_ref}")).ok()?;
                        self.opc.get_part(&uri).is_err().then_some((uri, target_ref))
                    })
                    .expect("the alternative-format part-name space is unbounded");
                self.opc.try_add_part(Box::new(BlobPart::new(
                    uri.clone(),
                    data.content_type().to_string(),
                    data.bytes().to_vec(),
                )))?;
                (target_ref, TargetMode::Internal, Some(uri))
            },
        };
        let relation_result = self
            .opc
            .get_part_mut(&document_uri)?
            .rels_mut()
            .try_add_relationship(
                relationship_type.to_string(),
                target_ref,
                relationship_id.clone(),
                target_mode,
            );
        if let Err(error) = relation_result {
            if let Some(uri) = &installed_part {
                self.opc.remove_part(uri);
            }
            return Err(error.into());
        }
        Ok((AltChunk::new(relationship_id, match_source)?, installed_part))
    }

    fn validate_alt_chunk_relationship(&self, chunk: &AltChunk) -> Result<Option<PackURI>> {
        let document_uri = PackURI::new("/word/document.xml")
            .map_err(|error| OoxmlError::InvalidUri(format!("document URI: {error}")))?;
        let relationship = self
            .opc
            .get_part(&document_uri)?
            .rels()
            .get(chunk.relationship_id())
            .ok_or_else(|| {
                OoxmlError::InvalidFormat(format!(
                    "altChunk relationship {:?} is missing",
                    chunk.relationship_id()
                ))
            })?;
        if !is_alternative_format_relationship(relationship.reltype()) {
            return Err(OoxmlError::InvalidFormat(format!(
                "relationship {:?} is not an alternative-format import",
                chunk.relationship_id()
            )));
        }
        if relationship.is_external() {
            return Ok(None);
        }
        let target = relationship.target_partname().map_err(|error| {
            OoxmlError::InvalidFormat(format!(
                "invalid altChunk relationship {:?}: {error}",
                chunk.relationship_id()
            ))
        })?;
        self.opc.get_part(&target).map_err(|_| {
            OoxmlError::InvalidFormat(format!(
                "altChunk relationship {:?} targets a missing part",
                chunk.relationship_id()
            ))
        })?;
        Ok(Some(target))
    }

    fn rollback_alt_chunk_target(&mut self, id: &str, part: Option<&PackURI>) -> Result<()> {
        let document_uri = PackURI::new("/word/document.xml")
            .map_err(|error| OoxmlError::InvalidUri(format!("document URI: {error}")))?;
        self.opc.get_part_mut(&document_uri)?.rels_mut().remove(id);
        if let Some(part) = part {
            self.opc.remove_part(part);
        }
        Ok(())
    }

    fn remove_alt_chunk_relationship(
        &mut self,
        chunk: &AltChunk,
        target: Option<&PackURI>,
    ) -> Result<()> {
        if self
            .mutable_doc
            .as_ref()
            .is_some_and(|document| {
                document
                    .alt_chunks()
                    .iter()
                    .any(|remaining| remaining.relationship_id() == chunk.relationship_id())
            })
        {
            return Ok(());
        }
        let document_uri = PackURI::new("/word/document.xml")
            .map_err(|error| OoxmlError::InvalidUri(format!("document URI: {error}")))?;
        self.opc
            .get_part_mut(&document_uri)?
            .rels_mut()
            .remove(chunk.relationship_id());
        let Some(target) = target else {
            return Ok(());
        };
        let package_reference = self.opc.rels().iter().any(|relationship| {
            !relationship.is_external()
                && relationship.target_partname().is_ok_and(|part| &part == target)
        });
        let part_reference = self.opc.iter_parts().any(|part| {
            part.rels().iter().any(|relationship| {
                !relationship.is_external()
                    && relationship.target_partname().is_ok_and(|part| &part == target)
            })
        });
        if !package_reference && !part_reference {
            self.opc.remove_part(target);
        }
        Ok(())
    }

    /// Discover every validated Custom XML Data Storage relationship occurrence.
    pub fn custom_xml_data_stores(&self) -> Result<Vec<CustomXmlDataItem>> {
        discover_custom_xml_data(&self.opc)
    }

    /// Discover typed, inert bibliography source stores from Custom XML.
    ///
    /// Word stores its current bibliography source list in a document Custom
    /// XML data store. This method exposes stored source values and style
    /// metadata only. It never matches source tags to citations, resolves
    /// schemas or styles, runs transforms, refreshes fields, or changes data.
    pub fn bibliography_source_stores(&self) -> Result<Vec<BibliographySourceStore>> {
        let items = discover_custom_xml_data(&self.opc)?;
        discover_bibliography_source_stores(&items)
    }

    /// Discover typed, inert bibliography sources in package and XML order.
    ///
    /// This flattens [`Self::bibliography_source_stores`] without resolving
    /// `CITATION` fields or applying bibliography style rules.
    pub fn bibliography_sources(&self) -> Result<Vec<BibliographySource>> {
        let stores = self.bibliography_source_stores()?;
        Ok(stores
            .iter()
            .flat_map(|store| store.sources().iter().cloned())
            .collect())
    }

    /// Return the number of typed, inert bibliography sources.
    pub fn bibliography_source_count(&self) -> Result<usize> {
        Ok(self.bibliography_sources()?.len())
    }

    /// Find a Custom XML data store by its case-insensitive datastore item GUID.
    pub fn find_custom_xml_data_store(&self, item_id: &str) -> Result<Option<CustomXmlDataItem>> {
        Ok(discover_custom_xml_data(&self.opc)?.into_iter().find(|item| {
            item.properties
                .as_ref()
                .is_some_and(|properties| properties.item_id.eq_ignore_ascii_case(item_id))
        }))
    }

    /// Add a collision-safe `/customXml/itemN.xml` data store to the main document.
    pub fn add_custom_xml_data_store(
        &mut self,
        store: NewCustomXmlDataStore,
    ) -> Result<CustomXmlDataItem> {
        validate_custom_xml_content_type(&store.content_type)?;
        validate_custom_xml_payload(&store.xml)?;
        let properties = CustomXmlDataProperties {
            item_id: store.item_id,
            schema_references: store.schema_references,
        };
        validate_custom_xml_properties(&properties)?;
        let source_part_name = self.opc.main_document_part()?.partname().clone();
        let source = self.opc.get_part(&source_part_name)?;
        let relationship_id = (1usize..)
            .map(|number| format!("rIdCustomXml{number}"))
            .find(|id| source.rels().get(id).is_none())
            .expect("the relationship ID space is unbounded");
        let (data_part_name, properties_part_name) = (1usize..)
            .find_map(|number| {
                let data = PackURI::new(format!("/customXml/item{number}.xml")).ok()?;
                let properties = PackURI::new(format!("/customXml/itemProps{number}.xml")).ok()?;
                let conflict = self.opc.iter_parts().any(|part| {
                    part.partname().as_str().eq_ignore_ascii_case(data.as_str())
                        || part
                            .partname()
                            .as_str()
                            .eq_ignore_ascii_case(properties.as_str())
                });
                (!conflict).then_some((data, properties))
            })
            .expect("the Custom XML part-name space is unbounded");
        add_custom_xml_data(
            &mut self.opc,
            NewCustomXmlDataItem {
                source_part_name,
                relationship_id,
                data_part_name: data_part_name.clone(),
                content_type: store.content_type,
                xml: store.xml,
                properties_part_name: Some(properties_part_name),
                properties_relationship_id: "rIdProps1".to_string(),
                properties: Some(properties),
                conformance: store.conformance,
            },
        )?;
        let _ = self.opc.clear_digital_signatures();
        discover_custom_xml_data(&self.opc)?
            .into_iter()
            .find(|item| item.data_part_name == data_part_name)
            .ok_or_else(|| {
                OoxmlError::InvalidFormat("new Custom XML data store was not discoverable".into())
            })
    }

    /// Replace only the inert XML payload of a data store.
    pub fn update_custom_xml_data_store(&mut self, item_id: &str, xml: Vec<u8>) -> Result<()> {
        validate_custom_xml_payload(&xml)?;
        let item = self
            .find_custom_xml_data_store(item_id)?
            .ok_or_else(|| OoxmlError::PartNotFound(format!("Custom XML itemID '{item_id}'")))?;
        self.opc.get_part_mut(&item.data_part_name)?.set_blob(xml);
        let _ = self.opc.clear_digital_signatures();
        Ok(())
    }

    /// Replace payload, content type, schema references, and canonical properties.
    pub fn replace_custom_xml_data_store(
        &mut self,
        item_id: &str,
        replacement: NewCustomXmlDataStore,
    ) -> Result<()> {
        validate_custom_xml_content_type(&replacement.content_type)?;
        validate_custom_xml_payload(&replacement.xml)?;
        if !replacement.item_id.eq_ignore_ascii_case(item_id) {
            return Err(OoxmlError::InvalidFormat(
                "replacement itemID must identify the existing data store".into(),
            ));
        }
        let properties = CustomXmlDataProperties {
            item_id: replacement.item_id,
            schema_references: replacement.schema_references,
        };
        let properties_xml = write_custom_xml_properties(&properties, replacement.conformance)?;
        let item = self
            .find_custom_xml_data_store(item_id)?
            .ok_or_else(|| OoxmlError::PartNotFound(format!("Custom XML itemID '{item_id}'")))?;
        let properties_part_name = item.properties_part_name.ok_or_else(|| {
            OoxmlError::InvalidFormat("Custom XML data store has no properties part".into())
        })?;
        let existing_relationships = self
            .opc
            .get_part(&item.data_part_name)?
            .rels()
            .iter()
            .map(|relationship| {
                (
                    relationship.reltype().to_string(),
                    relationship.target_ref().to_string(),
                    relationship.r_id().to_string(),
                    relationship.is_external(),
                )
            })
            .collect::<Vec<_>>();
        let mut data_part = BlobPart::new(
            item.data_part_name.clone(),
            replacement.content_type,
            replacement.xml,
        );
        for (reltype, target, id, external) in existing_relationships {
            data_part
                .rels_mut()
                .add_relationship(reltype, target, id, external);
        }
        self.opc.add_part(Box::new(data_part));
        self.opc
            .get_part_mut(&properties_part_name)?
            .set_blob(properties_xml);
        let _ = self.opc.clear_digital_signatures();
        Ok(())
    }

    /// Remove a data store unless an SDT still binds to its item GUID.
    pub fn remove_custom_xml_data_store(&mut self, item_id: &str) -> Result<bool> {
        let items = discover_custom_xml_data(&self.opc)?;
        let matching = items
            .iter()
            .filter(|item| {
                item.properties
                    .as_ref()
                    .is_some_and(|properties| properties.item_id.eq_ignore_ascii_case(item_id))
            })
            .cloned()
            .collect::<Vec<_>>();
        if matching.is_empty() {
            return Ok(false);
        }
        if self.custom_xml_bindings()?.iter().any(|binding| {
            binding.store_item_id.eq_ignore_ascii_case(item_id)
        }) {
            return Err(OoxmlError::InvalidFormat(format!(
                "Custom XML itemID '{item_id}' is still referenced by a content control"
            )));
        }
        for item in &matching {
            self.opc
                .get_part_mut(&item.source_part_name)?
                .rels_mut()
                .remove(&item.relationship_id);
        }
        let data_part_name = matching[0].data_part_name.clone();
        let properties_part_name = matching[0].properties_part_name.clone();
        if !self.part_is_referenced(&data_part_name) {
            self.opc.remove_part(&data_part_name);
            if let Some(properties_part_name) = properties_part_name
                && !self.part_is_referenced(&properties_part_name)
            {
                self.opc.remove_part(&properties_part_name);
            }
        }
        let _ = self.opc.clear_digital_signatures();
        Ok(true)
    }

    /// Reorder main-document data-store relationships by item GUID.
    pub fn reorder_custom_xml_data_stores(&mut self, ordered_item_ids: &[String]) -> Result<()> {
        let source_part_name = self.opc.main_document_part()?.partname().clone();
        let items = discover_custom_xml_data(&self.opc)?
            .into_iter()
            .filter(|item| item.source_part_name == source_part_name)
            .collect::<Vec<_>>();
        if items.len() != ordered_item_ids.len() {
            return Err(OoxmlError::InvalidFormat(
                "reorder list must contain every main-document Custom XML item".into(),
            ));
        }
        let mut by_id = std::collections::HashMap::new();
        for item in &items {
            let id = item.properties.as_ref().ok_or_else(|| {
                OoxmlError::InvalidFormat("Custom XML item has no datastore itemID".into())
            })?.item_id.to_ascii_lowercase();
            if by_id.insert(id, item).is_some() {
                return Err(OoxmlError::InvalidFormat(
                    "main-document Custom XML items are not uniquely reorderable".into(),
                ));
            }
        }
        let mut ordered = Vec::with_capacity(items.len());
        let mut seen = std::collections::HashSet::new();
        for item_id in ordered_item_ids {
            let key = item_id.to_ascii_lowercase();
            if !seen.insert(key.clone()) {
                return Err(OoxmlError::InvalidFormat("duplicate reorder itemID".into()));
            }
            let item = *by_id.get(&key).ok_or_else(|| {
                OoxmlError::InvalidFormat(format!("unknown reorder itemID '{item_id}'"))
            })?;
            let reltype = self
                .opc
                .get_part(&source_part_name)?
                .rels()
                .get(&item.relationship_id)
                .expect("discovery validated the relationship")
                .reltype()
                .to_string();
            ordered.push((item, reltype));
        }
        let source = self.opc.get_part(&source_part_name)?;
        let reserved = source
            .rels()
            .iter()
            .filter(|relationship| {
                !items
                    .iter()
                    .any(|item| item.relationship_id == relationship.r_id())
            })
            .map(|relationship| relationship.r_id().to_string())
            .collect::<std::collections::HashSet<_>>();
        let ids = (1usize..)
            .filter_map(|batch| {
                let candidates = (0..ordered.len())
                    .map(|index| format!("rIdCustomXmlOrder{batch:04}_{index:06}"))
                    .collect::<Vec<_>>();
                candidates.iter().all(|id| !reserved.contains(id)).then_some(candidates)
            })
            .next()
            .expect("the relationship ID space is unbounded");
        let source = self.opc.get_part_mut(&source_part_name)?;
        let source_base_uri = source.partname().base_uri().to_string();
        for item in &items {
            source.rels_mut().remove(&item.relationship_id);
        }
        for ((item, reltype), id) in ordered.into_iter().zip(ids) {
            source.rels_mut().add_relationship(
                reltype,
                item.data_part_name.relative_ref(&source_base_uri),
                id,
                false,
            );
        }
        let _ = self.opc.clear_digital_signatures();
        Ok(())
    }

    /// Collect and lexically validate SDT bindings from every permitted Word container.
    pub fn custom_xml_bindings(&self) -> Result<Vec<CustomXmlBinding>> {
        let permitted = [
            ct::WML_DOCUMENT_MAIN,
            ct::WML_DOCUMENT_GLOSSARY,
            ct::WML_HEADER,
            ct::WML_FOOTER,
            ct::WML_FOOTNOTES,
            ct::WML_ENDNOTES,
        ];
        let mut bindings = Vec::new();
        for part in self
            .opc
            .iter_parts()
            .filter(|part| permitted.contains(&part.content_type()))
        {
            for control in ContentControl::extract_from_document(part.blob())? {
                control.validate_data_binding()?;
                if let (Some(xpath), Some(store_item_id)) = (
                    control.data_binding_xpath(),
                    control.data_binding_store_item_id(),
                ) {
                    bindings.push(CustomXmlBinding {
                        source_part_name: part.partname().clone(),
                        content_control_id: control.id(),
                        xpath: xpath.to_string(),
                        store_item_id: store_item_id.to_string(),
                        prefix_mappings: control
                            .data_binding_prefix_mappings()
                            .map(str::to_string),
                    });
                }
            }
        }
        bindings.sort_unstable_by(|left, right| {
            left.source_part_name
                .as_str()
                .cmp(right.source_part_name.as_str())
                .then_with(|| left.content_control_id.cmp(&right.content_control_id))
        });
        Ok(bindings)
    }

    /// Validate that every permitted SDT binding resolves to a datastore item GUID.
    pub fn validate_custom_xml_binding_integrity(&self) -> Result<()> {
        let item_ids = discover_custom_xml_data(&self.opc)?
            .into_iter()
            .filter_map(|item| item.properties.map(|properties| properties.item_id.to_ascii_lowercase()))
            .collect::<std::collections::HashSet<_>>();
        for binding in self.custom_xml_bindings()? {
            if !item_ids.contains(&binding.store_item_id.to_ascii_lowercase()) {
                return Err(OoxmlError::InvalidFormat(format!(
                    "content control {} in '{}' references missing Custom XML itemID '{}'",
                    binding.content_control_id,
                    binding.source_part_name.as_str(),
                    binding.store_item_id
                )));
            }
        }
        Ok(())
    }

    fn part_is_referenced(&self, target: &PackURI) -> bool {
        self.opc.rels().iter().any(|relationship| {
            !relationship.is_external()
                && relationship.target_partname().is_ok_and(|part| &part == target)
        }) || self.opc.iter_parts().any(|part| {
            part.rels().iter().any(|relationship| {
                !relationship.is_external()
                    && relationship.target_partname().is_ok_and(|part| &part == target)
            })
        })
    }

    /// Return the validated inert mail-merge settings, if configured.
    pub fn mail_merge_settings(&self) -> Result<Option<MailMergeSettings>> {
        let snapshot = self.settings_part_snapshot()?;
        let part = settings_part_from_snapshot(&snapshot, snapshot.xml.clone(), None);
        Ok(DocumentSettings::extract_from_part(&part)?.mail_merge().cloned())
    }

    /// Resolve a mail-merge relationship without opening or fetching its target.
    pub fn mail_merge_target(&self, relationship_id: &str) -> Result<MailMergeTarget> {
        let snapshot = self.settings_part_snapshot()?;
        let part = self.opc.get_part(&snapshot.target)?;
        let relationship = part.rels().get(relationship_id).ok_or_else(|| {
            OoxmlError::InvalidFormat(format!(
                "mail-merge relationship '{relationship_id}' is missing"
            ))
        })?;
        if !is_mail_merge_relationship_type(relationship.reltype()) {
            return Err(OoxmlError::InvalidFormat(format!(
                "relationship '{relationship_id}' is not a mail-merge source"
            )));
        }
        if relationship.is_external() {
            return Ok(MailMergeTarget::External(relationship.target_ref().to_string()));
        }
        let target = relationship.target_partname()?;
        let target_part = self.opc.get_part(&target)?;
        Ok(MailMergeTarget::Internal {
            part_name: target,
            bytes: target_part.blob().to_vec(),
            content_type: target_part.content_type().to_string(),
        })
    }

    /// Set or replace the complete mail-merge graph atomically.
    pub fn set_mail_merge(
        &mut self,
        mut settings: MailMergeSettings,
        data_source: Option<MailMergeSource>,
        header_source: Option<MailMergeSource>,
        recipients: Option<MailMergeRecipients>,
        conformance: MailMergeConformance,
    ) -> Result<()> {
        let snapshot = self.settings_part_snapshot()?;
        let original = settings_part_from_snapshot(&snapshot, snapshot.xml.clone(), None);
        DocumentSettings::extract_from_part(&original)?;
        let old_targets = self.mail_merge_internal_targets(&snapshot)?;
        let mut used_ids = snapshot
            .relationships
            .iter()
            .filter(|relationship| !is_mail_merge_relationship_type(&relationship.reltype))
            .map(|relationship| relationship.id.clone())
            .collect::<std::collections::HashSet<_>>();
        let mut staged_parts = Vec::new();
        let mut staged_relationships = Vec::new();

        let data_id = if let Some(source) = data_source {
            let (relationship, part) = self.stage_mail_merge_source(
                source,
                "Data",
                "mailMergeSource",
                conformance,
                &mut used_ids,
            )?;
            let id = relationship.id.clone();
            staged_relationships.push(relationship);
            if let Some(part) = part { staged_parts.push(part); }
            Some(id)
        } else { None };
        let header_id = if let Some(source) = header_source {
            let (relationship, part) = self.stage_mail_merge_source(
                source,
                "Header",
                "mailMergeHeaderSource",
                conformance,
                &mut used_ids,
            )?;
            let id = relationship.id.clone();
            staged_relationships.push(relationship);
            if let Some(part) = part { staged_parts.push(part); }
            Some(id)
        } else { None };
        let recipient_id = if let Some(recipients) = recipients {
            let xml = recipients.to_xml(conformance)?.into_bytes();
            let id = allocate_mail_merge_relationship_id("Recipients", &mut used_ids);
            let uri = self.allocate_mail_merge_part_name("recipientData", "xml")?;
            let target = uri.relative_ref(snapshot.target.base_uri());
            staged_parts.push(BlobPart::new(
                uri,
                MailMergeRecipients::content_type().to_string(),
                xml,
            ));
            staged_relationships.push(StoredRelationship {
                reltype: mail_merge_relationship_type(conformance, "recipientData"),
                target,
                id: id.clone(),
                external: false,
            });
            Some(id)
        } else { None };
        settings.assign_package_relationships(data_id, header_id, recipient_id);
        let patched = patch_mail_merge(&snapshot.xml, Some(&settings), conformance)?;
        let mut replacement = settings_part_from_snapshot(&snapshot, patched, None);
        let old_ids = replacement
            .rels()
            .iter()
            .filter(|relationship| is_mail_merge_relationship_type(relationship.reltype()))
            .map(|relationship| relationship.r_id().to_string())
            .collect::<Vec<_>>();
        for id in old_ids { replacement.rels_mut().remove(&id); }
        for relationship in staged_relationships {
            replacement.rels_mut().add_relationship(
                relationship.reltype,
                relationship.target,
                relationship.id,
                relationship.external,
            );
        }
        DocumentSettings::extract_from_part(&replacement)?;

        let mut installed = Vec::new();
        for part in staged_parts {
            let name = part.partname().clone();
            if let Err(error) = self.opc.try_add_part(Box::new(part)) {
                for installed_name in installed { self.opc.remove_part(&installed_name); }
                return Err(error.into());
            }
            installed.push(name);
        }
        if let Err(error) = self.commit_settings_part(&snapshot, replacement) {
            for installed_name in installed { self.opc.remove_part(&installed_name); }
            return Err(error);
        }
        for old_target in old_targets {
            if !self.part_is_referenced(&old_target) { self.opc.remove_part(&old_target); }
        }
        let _ = self.opc.clear_digital_signatures();
        Ok(())
    }

    /// Update settings and sources using the same atomic replacement semantics.
    pub fn update_mail_merge(
        &mut self,
        settings: MailMergeSettings,
        data_source: Option<MailMergeSource>,
        header_source: Option<MailMergeSource>,
        recipients: Option<MailMergeRecipients>,
        conformance: MailMergeConformance,
    ) -> Result<()> {
        self.set_mail_merge(settings, data_source, header_source, recipients, conformance)
    }

    /// Replace recipient inclusion flags while retaining inert source targets and settings.
    pub fn update_mail_merge_recipients(
        &mut self,
        recipients: MailMergeRecipients,
        conformance: MailMergeConformance,
    ) -> Result<()> {
        let settings = self.mail_merge_settings()?.ok_or_else(|| {
            OoxmlError::InvalidFormat("document has no mail-merge settings".into())
        })?;
        let data_source = settings
            .data_source_relationship_id()
            .map(|id| self.mail_merge_target(id).map(mail_merge_target_as_source))
            .transpose()?;
        let header_source = settings
            .header_source_relationship_id()
            .map(|id| self.mail_merge_target(id).map(mail_merge_target_as_source))
            .transpose()?;
        self.set_mail_merge(
            settings,
            data_source,
            header_source,
            Some(recipients),
            conformance,
        )
    }

    /// Clear mail-merge XML, relationships, and unreachable owned targets.
    pub fn clear_mail_merge(&mut self) -> Result<bool> {
        let snapshot = self.settings_part_snapshot()?;
        let original = settings_part_from_snapshot(&snapshot, snapshot.xml.clone(), None);
        if DocumentSettings::extract_from_part(&original)?.mail_merge().is_none() {
            return Ok(false);
        }
        let old_targets = self.mail_merge_internal_targets(&snapshot)?;
        let patched = patch_mail_merge(
            &snapshot.xml,
            None,
            MailMergeConformance::Transitional,
        )?;
        let mut replacement = settings_part_from_snapshot(&snapshot, patched, None);
        let ids = replacement
            .rels()
            .iter()
            .filter(|relationship| is_mail_merge_relationship_type(relationship.reltype()))
            .map(|relationship| relationship.r_id().to_string())
            .collect::<Vec<_>>();
        for id in ids { replacement.rels_mut().remove(&id); }
        DocumentSettings::extract_from_part(&replacement)?;
        self.commit_settings_part(&snapshot, replacement)?;
        for old_target in old_targets {
            if !self.part_is_referenced(&old_target) { self.opc.remove_part(&old_target); }
        }
        let _ = self.opc.clear_digital_signatures();
        Ok(true)
    }

    fn stage_mail_merge_source(
        &self,
        source: MailMergeSource,
        label: &str,
        relationship_suffix: &str,
        conformance: MailMergeConformance,
        used_ids: &mut std::collections::HashSet<String>,
    ) -> Result<(StoredRelationship, Option<BlobPart>)> {
        let id = allocate_mail_merge_relationship_id(label, used_ids);
        let settings_target = self.settings_part_snapshot()?.target;
        match source {
            MailMergeSource::External(uri) => {
                validate_mail_merge_external_uri(&uri)?;
                Ok((StoredRelationship {
                    reltype: mail_merge_relationship_type(conformance, relationship_suffix),
                    target: uri,
                    id,
                    external: true,
                }, None))
            },
            MailMergeSource::Internal { bytes, content_type, extension } => {
                validate_mail_merge_internal_source(&bytes, &content_type, &extension)?;
                let uri = self.allocate_mail_merge_part_name(label, &extension)?;
                let target = uri.relative_ref(settings_target.base_uri());
                let part = BlobPart::new(uri, content_type, bytes);
                Ok((StoredRelationship {
                    reltype: mail_merge_relationship_type(conformance, relationship_suffix),
                    target,
                    id,
                    external: false,
                }, Some(part)))
            },
        }
    }

    fn allocate_mail_merge_part_name(&self, stem: &str, extension: &str) -> Result<PackURI> {
        for number in 1usize.. {
            let candidate = PackURI::new(format!("/word/mailMerge/{stem}{number}.{extension}"))
                .map_err(OoxmlError::InvalidUri)?;
            if self.opc.iter_parts().all(|part| {
                !part.partname().as_str().eq_ignore_ascii_case(candidate.as_str())
            }) {
                return Ok(candidate);
            }
        }
        unreachable!("the mail-merge part-name space is unbounded")
    }

    fn mail_merge_internal_targets(&self, snapshot: &SettingsPartSnapshot) -> Result<Vec<PackURI>> {
        let Ok(part) = self.opc.get_part(&snapshot.target) else { return Ok(Vec::new()); };
        part.rels()
            .iter()
            .filter(|relationship| {
                is_mail_merge_relationship_type(relationship.reltype()) && !relationship.is_external()
            })
            .map(|relationship| relationship.target_partname().map_err(Into::into))
            .collect()
    }

    /// Get web-output settings for explicit package mutation.
    ///
    /// Existing settings are loaded lazily. If the package has no web-settings
    /// relationship, this starts with an empty typed model and creates the part
    /// on the next save. Merely editing document content does not rewrite the
    /// existing web-settings part.
    pub fn web_settings_mut(&mut self) -> Result<&mut WebSettings> {
        if self.mutable_web_settings.is_none() {
            self.mutable_web_settings = Some(self.load_web_settings()?.unwrap_or_default());
        }
        self.web_settings_dirty = true;
        Ok(self
            .mutable_web_settings
            .as_mut()
            .expect("web settings were initialized above"))
    }

    /// Replace the package's complete typed web-output settings.
    pub fn set_web_settings(&mut self, settings: WebSettings) -> &mut Self {
        self.mutable_web_settings = Some(settings);
        self.web_settings_dirty = true;
        self
    }

    /// Inspect the external template associated with this document without dereferencing it.
    pub fn attached_template(&self) -> Result<Option<AttachedTemplate>> {
        let snapshot = self.settings_part_snapshot()?;
        if !snapshot.relationship_exists {
            return Ok(None);
        }
        let part = settings_part_from_snapshot(&snapshot, snapshot.xml.clone(), None);
        Ok(DocumentSettings::extract_from_part(&part)?
            .attached_template()
            .cloned())
    }

    /// Associate this document with an external template URI.
    ///
    /// The URI is recorded inertly and is never fetched or executed.
    pub fn set_attached_template_uri(
        &mut self,
        target_uri: impl Into<String>,
    ) -> Result<&mut Self> {
        use litchi_opc::part::Part;

        let target_uri = target_uri.into();
        validate_attached_template_target(&target_uri)?;
        let snapshot = self.settings_part_snapshot()?;
        let original = settings_part_from_snapshot(&snapshot, snapshot.xml.clone(), None);
        let settings = DocumentSettings::extract_from_part(&original)?;
        let old_id = settings
            .attached_template()
            .map(|template| template.relationship_id().to_owned());

        let mut replacement = settings_part_from_snapshot(
            &snapshot,
            snapshot.xml.clone(),
            old_id.as_deref(),
        );
        let relationship_id = if let Some(id) = old_id {
            replacement.rels_mut().add_relationship(
                ATTACHED_TEMPLATE_RELATIONSHIP.to_owned(),
                target_uri,
                id.clone(),
                true,
            );
            id
        } else {
            replacement.relate_to_ext(&target_uri, ATTACHED_TEMPLATE_RELATIONSHIP)
        };
        replacement.set_blob(patch_attached_template(
            &snapshot.xml,
            Some(&relationship_id),
        )?);
        DocumentSettings::extract_from_part(&replacement)?;
        self.commit_settings_part(&snapshot, replacement)?;
        Ok(self)
    }

    /// Remove the attached-template element and its referenced relationship.
    pub fn remove_attached_template(&mut self) -> Result<Option<AttachedTemplate>> {
        let snapshot = self.settings_part_snapshot()?;
        if !snapshot.relationship_exists {
            return Ok(None);
        }
        let original = settings_part_from_snapshot(&snapshot, snapshot.xml.clone(), None);
        let settings = DocumentSettings::extract_from_part(&original)?;
        let Some(attached_template) = settings.attached_template().cloned() else {
            return Ok(None);
        };
        let replacement = settings_part_from_snapshot(
            &snapshot,
            patch_attached_template(&snapshot.xml, None)?,
            Some(attached_template.relationship_id()),
        );
        DocumentSettings::extract_from_part(&replacement)?;
        self.commit_settings_part(&snapshot, replacement)?;
        Ok(Some(attached_template))
    }

    /// Read the document variables stored in `settings.xml`.
    pub fn document_variables(&self) -> Result<Option<DocumentVariables>> {
        let snapshot = self.settings_part_snapshot()?;
        if !snapshot.relationship_exists {
            return Ok(None);
        }
        let part = settings_part_from_snapshot(&snapshot, snapshot.xml.clone(), None);
        Ok(Some(DocumentVariables::extract_from_settings_part(&part)?))
    }

    /// Insert or replace one document variable atomically.
    pub fn set_document_variable(
        &mut self,
        name: impl Into<String>,
        value: impl Into<String>,
    ) -> Result<Option<String>> {
        let snapshot = self.settings_part_snapshot()?;
        let original = settings_part_from_snapshot(&snapshot, snapshot.xml.clone(), None);
        DocumentSettings::extract_from_part(&original)?;
        let mut variables = DocumentVariables::extract_from_settings_part(&original)?;
        let previous = variables.insert(name, value)?;
        let xml = patch_document_variables(&snapshot.xml, &variables)?;
        let replacement = settings_part_from_snapshot(&snapshot, xml, None);
        DocumentSettings::extract_from_part(&replacement)?;
        DocumentVariables::extract_from_settings_part(&replacement)?;
        self.commit_settings_part(&snapshot, replacement)?;
        Ok(previous)
    }

    /// Remove one document variable atomically.
    pub fn remove_document_variable(&mut self, name: &str) -> Result<Option<String>> {
        let snapshot = self.settings_part_snapshot()?;
        if !snapshot.relationship_exists {
            return Ok(None);
        }
        let original = settings_part_from_snapshot(&snapshot, snapshot.xml.clone(), None);
        DocumentSettings::extract_from_part(&original)?;
        let mut variables = DocumentVariables::extract_from_settings_part(&original)?;
        let Some(previous) = variables.remove(name) else {
            return Ok(None);
        };
        let xml = patch_document_variables(&snapshot.xml, &variables)?;
        let replacement = settings_part_from_snapshot(&snapshot, xml, None);
        DocumentSettings::extract_from_part(&replacement)?;
        DocumentVariables::extract_from_settings_part(&replacement)?;
        self.commit_settings_part(&snapshot, replacement)?;
        Ok(Some(previous))
    }

    /// Remove every document variable atomically and return the number removed.
    pub fn clear_document_variables(&mut self) -> Result<usize> {
        let snapshot = self.settings_part_snapshot()?;
        if !snapshot.relationship_exists {
            return Ok(0);
        }
        let original = settings_part_from_snapshot(&snapshot, snapshot.xml.clone(), None);
        DocumentSettings::extract_from_part(&original)?;
        let mut variables = DocumentVariables::extract_from_settings_part(&original)?;
        let count = variables.count();
        if count == 0 {
            return Ok(0);
        }
        variables.clear();
        let xml = patch_document_variables(&snapshot.xml, &variables)?;
        let replacement = settings_part_from_snapshot(&snapshot, xml, None);
        DocumentSettings::extract_from_part(&replacement)?;
        DocumentVariables::extract_from_settings_part(&replacement)?;
        self.commit_settings_part(&snapshot, replacement)?;
        Ok(count)
    }

    fn load_web_settings(&self) -> Result<Option<WebSettings>> {
        let main_part = self.opc.main_document_part()?;
        let mut matches = main_part
            .rels()
            .iter()
            .filter(|relationship| is_web_settings_relationship(relationship.reltype()));
        let Some(relationship) = matches.next() else {
            return Ok(None);
        };
        if matches.next().is_some() {
            return Err(OoxmlError::InvalidFormat(
                "document has multiple web-settings relationships".into(),
            ));
        }
        if relationship.is_external() {
            return Err(OoxmlError::InvalidFormat(
                "web-settings relationship cannot be external".into(),
            ));
        }
        let target = relationship.target_partname()?;
        let part = self.opc.get_part(&target)?;
        Ok(Some(WebSettings::extract_from_part(part)?))
    }

    /// Load the typed glossary/building-block document, if present.
    pub fn glossary_document(&self) -> Result<Option<GlossaryDocument>> {
        load_from_package(&self.opc)
            .map_err(|error| OoxmlError::InvalidFormat(error.to_string()))
    }

    /// Load the glossary together with its owned auxiliary relationship graph.
    pub fn glossary_package(&self) -> Result<Option<GlossaryPackage>> {
        load_package_from_package(&self.opc)
            .map_err(|error| OoxmlError::InvalidFormat(error.to_string()))
    }

    /// Atomically install or replace a glossary while preserving its existing graph.
    pub fn set_glossary_document(&mut self, document: GlossaryDocument) -> Result<()> {
        let mut package = self.glossary_package()?.unwrap_or_else(|| {
            let strict = self
                .opc
                .main_document_part()
                .is_ok_and(|part| {
                    std::str::from_utf8(part.blob()).is_ok_and(|xml| {
                        xml.contains("http://purl.oclc.org/ooxml/wordprocessingml/main")
                    })
                });
            GlossaryPackage::new(GlossaryDocument::default(), strict)
        });
        package.document = document;
        store_in_package(&mut self.opc, package)
            .map_err(|error| OoxmlError::InvalidFormat(error.to_string()))
    }

    /// Atomically install a complete glossary and auxiliary OPC graph.
    pub fn set_glossary_package(&mut self, package: GlossaryPackage) -> Result<()> {
        store_in_package(&mut self.opc, package)
            .map_err(|error| OoxmlError::InvalidFormat(error.to_string()))
    }

    /// Atomically edit the glossary document while preserving auxiliary parts.
    pub fn update_glossary_document<F>(&mut self, update: F) -> Result<()>
    where
        F: FnOnce(&mut GlossaryDocument) -> Result<()>,
    {
        let mut package = self
            .glossary_package()?
            .ok_or_else(|| OoxmlError::PartNotFound("glossary document".into()))?;
        update(&mut package.document)?;
        store_in_package(&mut self.opc, package)
            .map_err(|error| OoxmlError::InvalidFormat(error.to_string()))
    }

    /// Add one building block and return its insertion-order index.
    pub fn add_glossary_entry(&mut self, entry: GlossaryEntry) -> Result<usize> {
        let mut index = 0;
        if self.glossary_document()?.is_some() {
            self.update_glossary_document(|document| {
                index = document
                    .add_entry(entry)
                    .map_err(|error| OoxmlError::InvalidFormat(error.to_string()))?;
                Ok(())
            })?;
        } else {
            let mut document = GlossaryDocument::default();
            index = document
                .add_entry(entry)
                .map_err(|error| OoxmlError::InvalidFormat(error.to_string()))?;
            self.set_glossary_document(document)?;
        }
        Ok(index)
    }

    /// Replace one building block atomically.
    pub fn replace_glossary_entry(
        &mut self,
        index: usize,
        entry: GlossaryEntry,
    ) -> Result<GlossaryEntry> {
        let mut previous = None;
        self.update_glossary_document(|document| {
            previous = Some(
                document
                    .replace_entry(index, entry)
                    .map_err(|error| OoxmlError::InvalidFormat(error.to_string()))?,
            );
            Ok(())
        })?;
        previous.ok_or_else(|| OoxmlError::Other("glossary replacement failed".into()))
    }

    /// Atomically update one building block in place.
    pub fn update_glossary_entry<F>(&mut self, index: usize, update: F) -> Result<()>
    where
        F: FnOnce(&mut GlossaryEntry) -> Result<()>,
    {
        let mut update = Some(update);
        self.update_glossary_document(|document| {
            document
                .update_entry(index, |entry| {
                    update.take().expect("glossary update closure called once")(entry)
                        .map_err(|error| Box::new(error) as Box<dyn std::error::Error + Send + Sync>)
                })
                .map_err(|error| OoxmlError::InvalidFormat(error.to_string()))
        })
    }

    /// Reorder one building block while preserving all other insertion positions.
    pub fn move_glossary_entry(&mut self, from: usize, to: usize) -> Result<()> {
        self.update_glossary_document(|document| {
            document
                .move_entry(from, to)
                .map_err(|error| OoxmlError::InvalidFormat(error.to_string()))
        })
    }

    /// Find a building block by case-insensitive name.
    pub fn find_glossary_entry(&self, name: &str) -> Result<Option<(usize, GlossaryEntry)>> {
        Ok(self.glossary_document()?.and_then(|document| {
            document
                .find_entry(name)
                .map(|(index, entry)| (index, entry.clone()))
        }))
    }

    /// Remove one building block while preserving the remaining order.
    pub fn remove_glossary_entry(&mut self, index: usize) -> Result<GlossaryEntry> {
        let mut removed = None;
        self.update_glossary_document(|document| {
            removed = Some(
                document
                    .remove_entry(index)
                    .map_err(|error| OoxmlError::InvalidFormat(error.to_string()))?,
            );
            Ok(())
        })?;
        removed.ok_or_else(|| OoxmlError::Other("glossary removal failed".into()))
    }

    /// Remove every building block but retain the glossary part and its graph.
    pub fn clear_glossary_entries(&mut self) -> Result<usize> {
        let mut count = 0;
        self.update_glossary_document(|document| {
            count = document.clear_entries();
            Ok(())
        })?;
        Ok(count)
    }

    /// Remove the glossary relationship and only reachable glossary-owned parts.
    pub fn remove_glossary_document(&mut self) -> Result<Option<GlossaryDocument>> {
        remove_from_package(&mut self.opc)
            .map_err(|error| OoxmlError::InvalidFormat(error.to_string()))
    }

    /// Alias for removing the complete glossary graph.
    pub fn clear_glossary_document(&mut self) -> Result<Option<GlossaryDocument>> {
        self.remove_glossary_document()
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
    /// use litchi_ooxml::docx::Package;
    ///
    /// let mut pkg = Package::new()?;
    /// // Modify document...
    /// pkg.save("output.docx")?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn save<P: AsRef<Path>>(&mut self, path: P) -> Result<()> {
        let mut file = std::fs::OpenOptions::new()
            .write(true)
            .create(true)
            .truncate(true)
            .open(path)?;
        self.to_stream(&mut file)
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
    /// use litchi_ooxml::docx::Package;
    /// use std::io::Cursor;
    ///
    /// let mut pkg = Package::new()?;
    /// // Modify document...
    /// let mut cursor = Cursor::new(Vec::new());
    /// pkg.to_stream(&mut cursor)?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn to_stream<W: Write + Seek>(&mut self, writer: W) -> Result<()> {
        use crate::docx::writer::relmap::RelationshipMapper;
        use litchi_opc::constants::relationship_type as rt;

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
                let section_header_footer_parts =
                    mutable_doc.collect_section_header_footer_parts()?;
                let explicit_section_relationships =
                    mutable_doc.collect_explicit_section_header_footer_relationships()?;
                let mut planned_section_parts = Vec::new();
                for (index, (header, part)) in section_header_footer_parts.into_iter().enumerate() {
                    let stem = if header { "headerSection" } else { "footerSection" };
                    let filename = format!("{stem}{}.xml", index + 1);
                    let uri = PackURI::new(format!("/word/{filename}"))
                        .map_err(|error| OoxmlError::InvalidUri(error.to_string()))?;
                    if self.opc.get_part(&uri).is_ok() {
                        return Err(OoxmlError::InvalidFormat(format!(
                            "section header/footer part {uri} already exists"
                        )));
                    }
                    planned_section_parts.push((header, part, uri, filename));
                }

                // Step 2: Create a relationship mapper and add relationships
                let mut rel_mapper = RelationshipMapper::new();

                // Create the document part first (we'll update it later)
                let doc_uri = PackURI::new("/word/document.xml")
                    .map_err(|e| OoxmlError::InvalidUri(format!("document URI: {}", e)))?;

                if !explicit_section_relationships.is_empty() {
                    let existing_document = self.opc.get_part(&doc_uri).map_err(|_| {
                        OoxmlError::InvalidFormat(
                            "section references exist without a document part".to_string(),
                        )
                    })?;
                    for (id, header) in &explicit_section_relationships {
                        let relationship = existing_document.rels().get(id).ok_or_else(|| {
                            OoxmlError::InvalidFormat(format!(
                                "section relationship {id:?} is missing"
                            ))
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
                            return Err(OoxmlError::InvalidFormat(format!(
                                "section relationship {id:?} has the wrong type or target mode"
                            )));
                        }
                        let target = relationship.target_partname().map_err(|error| {
                            OoxmlError::InvalidFormat(format!(
                                "invalid section relationship {id:?}: {error}"
                            ))
                        })?;
                        let part = self.opc.get_part(&target).map_err(|_| {
                            OoxmlError::InvalidFormat(format!(
                                "section relationship {id:?} targets a missing part"
                            ))
                        })?;
                        let expected_content_type = if *header {
                            ct::WML_HEADER
                        } else {
                            ct::WML_FOOTER
                        };
                        if part.content_type() != expected_content_type {
                            return Err(OoxmlError::InvalidFormat(format!(
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
                for (header, part, uri, _) in planned_section_parts {
                    let content_type = if header { ct::WML_HEADER } else { ct::WML_FOOTER };
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
                    let settings_uri = PackURI::new("/word/settings.xml").map_err(|error| {
                        OoxmlError::InvalidUri(format!("settings URI: {error}"))
                    })?;
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
            // Put the document back
            self.mutable_doc = Some(mutable_doc);
        }

        if self.web_settings_dirty {
            let xml = self
                .mutable_web_settings
                .as_ref()
                .expect("dirty web settings must have a typed model")
                .to_xml()?
                .into_bytes();
            self.update_web_settings_part(xml)?;
            self.web_settings_dirty = false;
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

        self.opc.to_stream(writer).map_err(|e| {
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
    /// use litchi_ooxml::docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let props = pkg.properties();
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn properties(&self) -> &DocumentProperties {
        &self.properties
    }

    /// Get a mutable reference to the document properties.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::docx::Package;
    ///
    /// let mut pkg = Package::new()?;
    /// pkg.properties_mut().title = Some("My Document".to_string());
    /// pkg.properties_mut().creator = Some("John Doe".to_string());
    /// pkg.save("document.docx")?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
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
    /// use litchi_ooxml::docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let custom_props = pkg.custom_properties();
    ///
    /// if let Some(value) = custom_props.get_property("ProjectName") {
    ///     println!("Project: {:?}", value);
    /// }
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn custom_properties(&self) -> &CustomProperties {
        &self.custom_properties
    }

    /// Get a mutable reference to the custom document properties.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::docx::Package;
    /// use litchi_ooxml::custom_properties::PropertyValue;
    ///
    /// let mut pkg = Package::new()?;
    /// let custom_props = pkg.custom_properties_mut();
    ///
    /// custom_props.add_property("ProjectName", PropertyValue::String("MyProject".to_string()));
    /// custom_props.add_property("Version", PropertyValue::Integer(1));
    /// custom_props.add_property("Budget", PropertyValue::Double(50000.0));
    ///
    /// pkg.save("document.docx")?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn custom_properties_mut(&mut self) -> &mut CustomProperties {
        &mut self.custom_properties
    }

    /// Update the core.xml properties part.
    fn update_core_properties(&mut self) -> Result<()> {
        use litchi_opc::part::BlobPart;

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
        use litchi_opc::constants::relationship_type as rt;
        use litchi_opc::part::BlobPart;

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
        use litchi_opc::constants::content_type as ct;
        use litchi_opc::constants::relationship_type as rt;
        use litchi_opc::part::BlobPart;

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
        use litchi_opc::constants::content_type as ct;
        use litchi_opc::constants::relationship_type as rt;
        use litchi_opc::part::BlobPart;

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
        use litchi_opc::constants::content_type as ct;
        use litchi_opc::constants::relationship_type as rt;
        use litchi_opc::part::BlobPart;

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
    fn update_settings_part(&mut self, xml: Vec<u8>) -> Result<()> {
        let snapshot = self.settings_part_snapshot()?;
        let part = settings_part_from_snapshot(&snapshot, xml, None);
        DocumentSettings::extract_from_part(&part)?;
        self.commit_settings_part(&snapshot, part)
    }

    fn settings_part_snapshot(&self) -> Result<SettingsPartSnapshot> {
        use litchi_opc::constants::relationship_type as rt;

        const STRICT_SETTINGS_RELATIONSHIP: &str =
            "http://purl.oclc.org/ooxml/officeDocument/relationships/settings";
        let document = self.opc.main_document_part()?;
        let document_uri = document.partname().clone();
        let mut matches = document.rels().iter().filter(|relationship| {
            matches!(relationship.reltype(), rt::SETTINGS | STRICT_SETTINGS_RELATIONSHIP)
        });
        let relationship = matches.next();
        if matches.next().is_some() {
            return Err(OoxmlError::InvalidFormat(
                "document has multiple settings relationships".into(),
            ));
        }
        let (target, relationship_exists) = match relationship {
            Some(relationship) if relationship.is_external() => {
                return Err(OoxmlError::InvalidFormat(
                    "settings relationship cannot be external".into(),
                ));
            },
            Some(relationship) => (relationship.target_partname()?, true),
            None => (
                PackURI::new("/word/settings.xml")
                    .map_err(|error| OoxmlError::InvalidUri(error.to_string()))?,
                false,
            ),
        };

        let (content_type, xml, relationships) = match self.opc.get_part(&target) {
            Ok(part) => {
                if part.content_type() != ct::WML_SETTINGS {
                    return Err(OoxmlError::InvalidFormat(format!(
                        "settings part has content type {:?}, expected {:?}",
                        part.content_type(),
                        ct::WML_SETTINGS
                    )));
                }
                (
                    part.content_type().to_owned(),
                    part.blob().to_vec(),
                    part.rels()
                        .iter()
                        .map(|relationship| StoredRelationship {
                            reltype: relationship.reltype().to_owned(),
                            target: relationship.target_ref().to_owned(),
                            id: relationship.r_id().to_owned(),
                            external: relationship.is_external(),
                        })
                        .collect(),
                )
            },
            Err(_) if relationship_exists => {
                return Err(OoxmlError::PartNotFound(format!("settings part {target}")));
            },
            Err(_) => (
                ct::WML_SETTINGS.to_owned(),
                crate::docx::template::default_settings_xml()
                    .as_bytes()
                    .to_vec(),
                Vec::new(),
            ),
        };
        Ok(SettingsPartSnapshot {
            document_uri,
            target,
            relationship_exists,
            content_type,
            xml,
            relationships,
        })
    }

    fn commit_settings_part(
        &mut self,
        snapshot: &SettingsPartSnapshot,
        part: litchi_opc::part::BlobPart,
    ) -> Result<()> {
        use litchi_opc::constants::relationship_type as rt;

        if !snapshot.relationship_exists {
            // Acquire the only fallible mutable reference before changing package state.
            self.opc
                .get_part_mut(&snapshot.document_uri)?
                .relate_to("settings.xml", rt::SETTINGS);
        }
        self.opc.add_part(Box::new(part));
        Ok(())
    }

    fn update_web_settings_part(&mut self, xml: Vec<u8>) -> Result<()> {
        use litchi_opc::constants::relationship_type as rt;
        use litchi_opc::part::{BlobPart, Part};

        let doc_uri = PackURI::new("/word/document.xml")
            .map_err(|error| OoxmlError::InvalidUri(format!("document URI: {error}")))?;
        let (target, relationship_exists) = {
            let document_part = self.opc.get_part(&doc_uri)?;
            let mut matches = document_part
                .rels()
                .iter()
                .filter(|relationship| is_web_settings_relationship(relationship.reltype()));
            let relationship = matches.next();
            if matches.next().is_some() {
                return Err(OoxmlError::InvalidFormat(
                    "document has multiple web-settings relationships".into(),
                ));
            }
            match relationship {
                Some(relationship) if relationship.is_external() => {
                    return Err(OoxmlError::InvalidFormat(
                        "web-settings relationship cannot be external".into(),
                    ));
                },
                Some(relationship) => (relationship.target_partname()?, true),
                None => (
                    PackURI::new("/word/webSettings.xml").map_err(|error| {
                        OoxmlError::InvalidUri(format!("webSettings URI: {error}"))
                    })?,
                    false,
                ),
            }
        };

        let (content_type, relationships) = match self.opc.get_part(&target) {
            Ok(part) => (
                part.content_type().to_owned(),
                part.rels()
                    .iter()
                    .map(|relationship| {
                        (
                            relationship.reltype().to_owned(),
                            relationship.target_ref().to_owned(),
                            relationship.r_id().to_owned(),
                            relationship.is_external(),
                        )
                    })
                    .collect::<Vec<_>>(),
            ),
            Err(_) if relationship_exists => {
                return Err(OoxmlError::PartNotFound(format!(
                    "web-settings part {target}"
                )));
            },
            Err(_) => (ct::WML_WEB_SETTINGS.to_owned(), Vec::new()),
        };

        let mut part = BlobPart::new(target, content_type, xml);
        for (reltype, target_ref, id, external) in relationships {
            part.rels_mut()
                .add_relationship(reltype, target_ref, id, external);
        }
        WebSettings::extract_from_part(&part)?;
        self.opc.add_part(Box::new(part));

        if !relationship_exists {
            self.opc
                .get_part_mut(&doc_uri)?
                .relate_to("webSettings.xml", rt::WEB_SETTINGS);
        }
        Ok(())
    }

    fn update_theme_part(&mut self, xml: String) -> Result<()> {
        use litchi_opc::part::BlobPart;

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
        mutable_doc: &crate::docx::writer::MutableDocument,
    ) -> Result<()> {
        use litchi_opc::constants::content_type as ct;
        use litchi_opc::constants::relationship_type as rt;
        use litchi_opc::part::BlobPart;

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

fn settings_part_from_snapshot(
    snapshot: &SettingsPartSnapshot,
    xml: Vec<u8>,
    omitted_relationship_id: Option<&str>,
) -> litchi_opc::part::BlobPart {
    use litchi_opc::part::{BlobPart, Part};

    let mut part = BlobPart::new(
        snapshot.target.clone(),
        snapshot.content_type.clone(),
        xml,
    );
    for relationship in &snapshot.relationships {
        if omitted_relationship_id == Some(relationship.id.as_str()) {
            continue;
        }
        part.rels_mut().add_relationship(
            relationship.reltype.clone(),
            relationship.target.clone(),
            relationship.id.clone(),
            relationship.external,
        );
    }
    part
}

fn is_mail_merge_relationship_type(value: &str) -> bool {
    ["mailMergeSource", "mailMergeHeaderSource", "recipientData"]
        .iter()
        .any(|suffix| {
            value == format!(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/{suffix}"
            ) || value == format!(
                "http://purl.oclc.org/ooxml/officeDocument/relationships/{suffix}"
            )
        })
}

fn mail_merge_relationship_type(
    conformance: MailMergeConformance,
    suffix: &str,
) -> String {
    let base = match conformance {
        MailMergeConformance::Transitional => {
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        },
        MailMergeConformance::Strict => {
            "http://purl.oclc.org/ooxml/officeDocument/relationships"
        },
    };
    format!("{base}/{suffix}")
}

fn allocate_mail_merge_relationship_id(
    label: &str,
    used: &mut std::collections::HashSet<String>,
) -> String {
    (1usize..)
        .map(|number| format!("rIdMailMerge{label}{number}"))
        .find(|id| used.insert(id.clone()))
        .expect("the mail-merge relationship ID space is unbounded")
}

fn validate_mail_merge_external_uri(uri: &str) -> Result<()> {
    if uri.is_empty() || uri.len() > 32 * 1024 || uri.chars().any(char::is_control) {
        return Err(OoxmlError::InvalidFormat(
            "mail-merge external target is empty or exceeds URI limits".into(),
        ));
    }
    Ok(())
}

fn validate_mail_merge_internal_source(
    bytes: &[u8],
    content_type: &str,
    extension: &str,
) -> Result<()> {
    if bytes.len() > 128 * 1024 * 1024 {
        return Err(OoxmlError::InvalidFormat(
            "mail-merge source exceeds the 128 MiB authoring limit".into(),
        ));
    }
    if content_type.is_empty()
        || content_type.len() > 1024
        || content_type.chars().any(char::is_control)
    {
        return Err(OoxmlError::InvalidFormat(
            "mail-merge source content type is invalid".into(),
        ));
    }
    if extension.is_empty()
        || extension.len() > 16
        || !extension.bytes().all(|byte| byte.is_ascii_alphanumeric())
    {
        return Err(OoxmlError::InvalidFormat(
            "mail-merge source extension is invalid".into(),
        ));
    }
    Ok(())
}

fn mail_merge_target_as_source(target: MailMergeTarget) -> MailMergeSource {
    match target {
        MailMergeTarget::External(uri) => MailMergeSource::External(uri),
        MailMergeTarget::Internal {
            part_name,
            bytes,
            content_type,
        } => {
            let extension = part_name
                .as_str()
                .rsplit_once('.')
                .map(|(_, extension)| extension)
                .filter(|extension| {
                    !extension.is_empty()
                        && extension.len() <= 16
                        && extension.bytes().all(|byte| byte.is_ascii_alphanumeric())
                })
                .unwrap_or("bin")
                .to_string();
            MailMergeSource::Internal {
                bytes,
                content_type,
                extension,
            }
        },
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Cursor;
    use tempfile::NamedTempFile;

    #[test]
    fn saves_and_reopens_package() {
        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        package
            .document_mut()
            .unwrap()
            .add_paragraph_with_text("round-trip text");
        package.save(file.path()).unwrap();

        let mut reopened = Package::open(file.path()).unwrap();
        assert!(
            reopened
                .document()
                .unwrap()
                .text()
                .unwrap()
                .contains("round-trip text")
        );

        reopened
            .document_mut()
            .unwrap()
            .add_paragraph_with_text("appended after reopen");
        reopened.save(file.path()).unwrap();
        let reopened_again = Package::open(file.path()).unwrap();
        let text = reopened_again.document().unwrap().text().unwrap();
        assert!(text.contains("round-trip text"));
        assert!(text.contains("appended after reopen"));
    }

    #[test]
    fn saves_and_reopens_inline_and_display_office_math() {
        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let inline = crate::docx::OfficeMath::text("x + y");
        let display = crate::docx::OfficeMath::from_xml(
            "<m:oMath><m:r><m:t>z</m:t></m:r></m:oMath>",
        )
        .unwrap();
        let mut package = Package::new().unwrap();
        {
            let document = package.document_mut().unwrap();
            document
                .add_paragraph()
                .add_inline_office_math(inline.clone());
            document.add_display_office_math(display.clone());
        }
        package.save(file.path()).unwrap();

        let reopened = Package::open(file.path()).unwrap();
        let document_uri = PackURI::new("/word/document.xml").unwrap();
        let document_xml = std::str::from_utf8(
            reopened
                .opc
                .get_part(&document_uri)
                .unwrap()
                .blob(),
        )
        .unwrap();
        let document_opening = &document_xml[..document_xml.find("><w:body>").unwrap()];
        assert!(document_opening.contains(
            "xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\""
        ));
        let paragraphs = reopened.document().unwrap().paragraphs().unwrap();
        assert_eq!(paragraphs[0].inline_office_math().unwrap(), vec![inline]);
        assert_eq!(paragraphs[1].display_office_math().unwrap(), vec![display]);
    }

    #[test]
    fn writes_and_rediscovers_distinct_watermarks() {
        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        let mut watermark = crate::docx::Watermark::text("INTERNAL");
        watermark.set_font("Aptos");
        watermark.set_color("808080");
        package.document_mut().unwrap().set_watermark(watermark);
        package.save(file.path()).unwrap();

        let reopened = Package::open(file.path()).unwrap();
        let watermarks = reopened.document().unwrap().watermarks().unwrap();

        assert_eq!(watermarks.len(), 1);
        assert_eq!(watermarks[0].get_text(), "INTERNAL");
        assert_eq!(watermarks[0].font(), "Aptos");
        assert_eq!(watermarks[0].color(), "#808080");
    }

    #[test]
    fn writes_and_discovers_typed_table_of_contents_fields() {
        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        {
            let document = package.document_mut().unwrap();
            document.add_heading("Overview", 1).unwrap();
            document
                .add_toc(
                    crate::docx::TableOfContents::new()
                        .heading_levels(1, 4)
                        .hyperlinks(true),
                )
                .unwrap();
        }
        package.save(file.path()).unwrap();

        let reopened = Package::open(file.path()).unwrap();
        let document = reopened.document().unwrap();
        let toc = document.table_of_contents().unwrap();
        assert_eq!(document.table_of_contents_count().unwrap(), 1);
        assert_eq!(toc.len(), 1);
        assert!(toc[0].includes_hyperlinks());
        assert!(toc[0].hides_page_numbers_in_web_layout());
        assert_eq!(
            toc[0].heading_style_levels().unwrap(),
            vec![crate::docx::TableOfContentsLevelRange::new(1, 4).unwrap()]
        );
    }

    #[test]
    fn writes_and_discovers_typed_table_of_contents_entry_fields() {
        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        {
            let document = package.document_mut().unwrap();
            let entry = document.add_paragraph();
            entry.add_field(crate::docx::writer::MutableField::with_result(
                r#"TC "Illustration 1" \f i \l 4 \n"#.to_string(),
                "cached entry".to_string(),
            ));
        }
        package.save(file.path()).unwrap();

        let reopened = Package::open(file.path()).unwrap();
        let document = reopened.document().unwrap();
        let entries = document.table_of_contents_entries().unwrap();
        assert_eq!(document.table_of_contents_entry_count().unwrap(), 1);
        assert_eq!(entries.len(), 1);
        assert_eq!(entries[0].entry(), "Illustration 1");
        assert_eq!(entries[0].cached_result(), Some("cached entry"));
        assert_eq!(entries[0].list_identifier().unwrap(), Some("i"));
        assert_eq!(entries[0].level().unwrap(), Some("4"));
        assert!(entries[0].omits_page_number());
    }

    #[test]
    fn writes_and_discovers_typed_table_of_authorities_fields() {
        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        {
            let document = package.document_mut().unwrap();
            let authority_table = document.add_paragraph();
            authority_table.add_field(crate::docx::writer::MutableField::with_result(
                r#"TOA \c 2 \b "Authorities" \p \h"#.to_string(),
                "Statutes\t3".to_string(),
            ));
            let entry = document.add_paragraph();
            entry.add_field(crate::docx::writer::MutableField::with_result(
                r#"TA \l "Example Statute" \s "Example" \c 2 \b"#.to_string(),
                "hidden marker".to_string(),
            ));
        }
        package.save(file.path()).unwrap();

        let reopened = Package::open(file.path()).unwrap();
        let document = reopened.document().unwrap();
        let authorities = document.tables_of_authorities().unwrap();
        assert_eq!(document.table_of_authorities_count().unwrap(), 1);
        assert_eq!(authorities.len(), 1);
        assert_eq!(authorities[0].category().unwrap(), Some(2));
        assert_eq!(authorities[0].bookmark().unwrap(), Some("Authorities"));
        assert!(authorities[0].uses_passim());
        assert!(authorities[0].includes_category_headers());

        let entries = document.table_of_authorities_entries().unwrap();
        assert_eq!(document.table_of_authorities_entry_count().unwrap(), 1);
        assert_eq!(entries.len(), 1);
        assert_eq!(entries[0].long_citation().unwrap(), Some("Example Statute"));
        assert_eq!(entries[0].short_citation().unwrap(), Some("Example"));
        assert_eq!(entries[0].category().unwrap(), Some(2));
        assert!(entries[0].is_bold());
    }

    #[test]
    fn writes_and_discovers_typed_citation_and_bibliography_fields() {
        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        {
            let document = package.document_mut().unwrap();
            let citation = document.add_paragraph();
            let mut primary = crate::docx::CitationSource::new("Doe2024").unwrap();
            primary.set_prefix(Some("qtd. in".to_string())).unwrap();
            let mut citation_spec = crate::docx::CitationFieldSpec::new(primary);
            citation_spec.set_locale(Some(1033));
            citation_spec
                .add_source(crate::docx::CitationSource::new("Smith2025").unwrap())
                .unwrap();
            citation_spec
                .set_cached_result(Some("(Doe, 2024; Smith, 2025)".to_string()))
                .unwrap();
            citation_spec.set_dirty(false);
            citation
                .add_field(crate::docx::writer::MutableField::citation(&citation_spec).unwrap());
            let bibliography = document.add_paragraph();
            let mut bibliography_spec = crate::docx::BibliographyFieldSpec::new();
            bibliography_spec.set_locale(Some(1033));
            bibliography_spec.set_filter(Some(crate::docx::BibliographyFilter::Locale(1036)));
            bibliography_spec.add_source_tag("Doe2024").unwrap();
            bibliography_spec.add_source_tag("Smith2025").unwrap();
            bibliography_spec
                .set_cached_result(Some("Doe. Example work.".to_string()))
                .unwrap();
            bibliography_spec.set_dirty(false);
            bibliography.add_field(
                crate::docx::writer::MutableField::bibliography(&bibliography_spec).unwrap(),
            );
        }
        package.save(file.path()).unwrap();

        let reopened = Package::open(file.path()).unwrap();
        let document = reopened.document().unwrap();
        let citations = document.citations().unwrap();
        assert_eq!(document.citation_count().unwrap(), 1);
        assert_eq!(citations.len(), 1);
        assert_eq!(citations[0].primary_source_tag(), "Doe2024");
        assert_eq!(citations[0].source_tags(), ["Doe2024", "Smith2025"]);
        assert!(citations[0].has_switch('l'));
        assert!(citations[0].has_switch('f'));
        assert!(!citations[0].is_dirty());
        assert_eq!(
            citations[0].cached_result(),
            Some("(Doe, 2024; Smith, 2025)")
        );

        let bibliographies = document.bibliographies().unwrap();
        assert_eq!(document.bibliography_count().unwrap(), 1);
        assert_eq!(bibliographies.len(), 1);
        assert_eq!(
            bibliographies[0].cached_result(),
            Some("Doe. Example work.")
        );
        assert!(bibliographies[0].has_switch('l'));
        assert!(bibliographies[0].has_switch('f'));
        assert!(bibliographies[0].has_switch('m'));
        assert!(!bibliographies[0].is_dirty());
        assert_eq!(bibliographies[0].switches()[0].argument(), Some("1033"));
        assert_eq!(bibliographies[0].switches()[1].argument(), Some("1036"));
    }

    #[test]
    fn writes_and_discovers_inert_document_variable_fields_without_resolution() {
        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        {
            let document = package.document_mut().unwrap();
            let paragraph = document.add_paragraph();
            paragraph.add_field(crate::docx::writer::MutableField::with_result(
                r#"DOCVARIABLE CustomerName \* MERGEFORMAT"#.to_string(),
                "cached customer".to_string(),
            ));
        }
        package.save(file.path()).unwrap();

        let reopened = Package::open(file.path()).unwrap();
        let document = reopened.document().unwrap();
        let fields = document.document_variable_fields().unwrap();
        assert_eq!(document.document_variable_field_count().unwrap(), 1);
        assert_eq!(fields.len(), 1);
        assert_eq!(fields[0].variable_name(), "CustomerName");
        assert_eq!(fields[0].cached_result(), Some("cached customer"));
        assert!(fields[0].has_switch('*'));
        assert_eq!(fields[0].switches()[0].argument(), Some("MERGEFORMAT"));
        assert!(
            document
                .document_variables()
                .unwrap()
                .is_none_or(|variables| variables.get("CustomerName").is_none())
        );
    }

    #[test]
    fn writes_and_discovers_inert_document_property_fields_without_resolution() {
        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        {
            let document = package.document_mut().unwrap();
            let paragraph = document.add_paragraph();
            paragraph.add_field(crate::docx::writer::MutableField::with_result(
                r#"DOCPROPERTY "Project Name" \* MERGEFORMAT \@ "MMMM d, yyyy""#.to_string(),
                "cached project".to_string(),
            ));
        }
        package.save(file.path()).unwrap();

        let reopened = Package::open(file.path()).unwrap();
        let document = reopened.document().unwrap();
        let fields = document.document_property_fields().unwrap();
        assert_eq!(document.document_property_field_count().unwrap(), 1);
        assert_eq!(fields.len(), 1);
        assert_eq!(fields[0].property_name(), "Project Name");
        assert_eq!(fields[0].cached_result(), Some("cached project"));
        assert!(fields[0].has_switch('*'));
        assert!(fields[0].has_switch('@'));
        assert_eq!(fields[0].switches()[0].argument(), Some("MERGEFORMAT"));
        assert_eq!(fields[0].switches()[1].argument(), Some("MMMM d, yyyy"));
    }

    #[test]
    fn writes_and_discovers_inert_document_information_fields_without_resolution() {
        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        {
            let document = package.document_mut().unwrap();
            let title = document.add_paragraph();
            title.add_field(crate::docx::writer::MutableField::with_result(
                r#"TITLE \* MERGEFORMAT"#.to_string(),
                "cached title".to_string(),
            ));
            let author = document.add_paragraph();
            author.add_field(crate::docx::writer::MutableField::with_result(
                r#"AUTHOR \@ "opaque format""#.to_string(),
                "cached author".to_string(),
            ));
        }
        package.save(file.path()).unwrap();

        let reopened = Package::open(file.path()).unwrap();
        let document = reopened.document().unwrap();
        let fields = document.document_information_fields().unwrap();
        assert_eq!(document.document_information_field_count().unwrap(), 2);
        assert_eq!(fields.len(), 2);
        assert_eq!(
            fields[0].kind(),
            crate::docx::DocumentInformationFieldKind::Title
        );
        assert_eq!(fields[0].cached_result(), Some("cached title"));
        assert!(fields[0].has_switch('*'));
        assert_eq!(fields[0].switches()[0].argument(), Some("MERGEFORMAT"));
        assert_eq!(
            fields[1].kind(),
            crate::docx::DocumentInformationFieldKind::Author
        );
        assert_eq!(fields[1].cached_result(), Some("cached author"));
        assert!(fields[1].has_switch('@'));
        assert_eq!(fields[1].switches()[0].argument(), Some("opaque format"));
    }

    #[test]
    fn writes_and_discovers_inert_document_context_fields_without_resolution() {
        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        {
            let document = package.document_mut().unwrap();
            let file_name = document.add_paragraph();
            file_name.add_field(crate::docx::writer::MutableField::with_result(
                r#"FILENAME \p"#.to_string(),
                "cached file name".to_string(),
            ));
            let page = document.add_paragraph();
            page.add_field(crate::docx::writer::MutableField::with_result(
                r#"PAGE \* MERGEFORMAT"#.to_string(),
                "cached page".to_string(),
            ));
        }
        package.save(file.path()).unwrap();

        let reopened = Package::open(file.path()).unwrap();
        let document = reopened.document().unwrap();
        let fields = document.document_context_fields().unwrap();
        assert_eq!(document.document_context_field_count().unwrap(), 2);
        assert_eq!(fields.len(), 2);
        assert_eq!(
            fields[0].kind(),
            crate::docx::DocumentContextFieldKind::FileName
        );
        assert_eq!(fields[0].cached_result(), Some("cached file name"));
        assert!(fields[0].has_switch('p'));
        assert_eq!(
            fields[1].kind(),
            crate::docx::DocumentContextFieldKind::Page
        );
        assert_eq!(fields[1].cached_result(), Some("cached page"));
        assert!(fields[1].has_switch('*'));
        assert_eq!(fields[1].switches()[0].argument(), Some("MERGEFORMAT"));
    }

    #[test]
    fn writes_and_discovers_typed_inert_merge_fields_without_merging() {
        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        {
            let document = package.document_mut().unwrap();
            let paragraph = document.add_paragraph();
            paragraph.add_field(crate::docx::writer::MutableField::with_result(
                r#"MERGEFIELD "Customer Region" \b "Dear " \f "!" \m \v \* MERGEFORMAT"#
                    .to_string(),
                "cached customer".to_string(),
            ));
        }
        package.save(file.path()).unwrap();

        let reopened = Package::open(file.path()).unwrap();
        assert!(reopened.mail_merge_settings().unwrap().is_none());
        let document = reopened.document().unwrap();
        let fields = document.typed_merge_fields().unwrap();
        assert_eq!(document.typed_merge_field_count().unwrap(), 1);
        assert_eq!(fields.len(), 1);
        assert_eq!(fields[0].field_name(), "Customer Region");
        assert_eq!(fields[0].cached_result(), Some("cached customer"));
        assert!(fields[0].has_switch('b'));
        assert!(fields[0].has_switch('f'));
        assert!(fields[0].has_switch('m'));
        assert!(fields[0].has_switch('v'));
        assert_eq!(fields[0].switches()[0].argument(), Some("Dear "));
        assert_eq!(fields[0].switches()[1].argument(), Some("!"));
    }

    #[test]
    fn writes_and_discovers_typed_inert_mail_merge_counters_without_merging() {
        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        {
            let document = package.document_mut().unwrap();
            document
                .add_paragraph()
                .add_field(crate::docx::writer::MutableField::with_result(
                    "MERGEREC".to_string(),
                    "12".to_string(),
                ));
            document
                .add_paragraph()
                .add_field(crate::docx::writer::MutableField::with_result(
                    "MERGESEQ".to_string(),
                    "3".to_string(),
                ));
        }
        package.save(file.path()).unwrap();

        let reopened = Package::open(file.path()).unwrap();
        assert!(reopened.mail_merge_settings().unwrap().is_none());
        let document = reopened.document().unwrap();
        let counters = document.mail_merge_counters().unwrap();
        assert_eq!(document.mail_merge_counter_count().unwrap(), 2);
        assert_eq!(counters.len(), 2);
        assert_eq!(
            counters[0].kind(),
            crate::docx::MailMergeCounterKind::Record
        );
        assert_eq!(counters[0].cached_result(), Some("12"));
        assert_eq!(
            counters[1].kind(),
            crate::docx::MailMergeCounterKind::Sequence
        );
        assert_eq!(counters[1].cached_result(), Some("3"));
    }

    #[test]
    fn writes_and_discovers_inert_mail_merge_next_fields_without_advancing_records() {
        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        package.document_mut().unwrap().add_paragraph().add_field(
            crate::docx::writer::MutableField::with_result(
                "NEXT".to_string(),
                "cached next".to_string(),
            ),
        );
        package.save(file.path()).unwrap();

        let reopened = Package::open(file.path()).unwrap();
        assert!(reopened.mail_merge_settings().unwrap().is_none());
        let document = reopened.document().unwrap();
        let fields = document.mail_merge_next_fields().unwrap();
        assert_eq!(document.mail_merge_next_field_count().unwrap(), 1);
        assert_eq!(fields.len(), 1);
        assert_eq!(fields[0].cached_result(), Some("cached next"));
    }

    #[test]
    fn writes_and_discovers_inert_conditional_mail_merge_controls_without_merging() {
        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        package.document_mut().unwrap().add_paragraph().add_field(
            crate::docx::writer::MutableField::with_result(
                r#"SKIPIF MERGEFIELD Order < 100"#.to_string(),
                "cached skipif".to_string(),
            ),
        );
        package.save(file.path()).unwrap();

        let reopened = Package::open(file.path()).unwrap();
        assert!(reopened.mail_merge_settings().unwrap().is_none());
        let document = reopened.document().unwrap();
        let controls = document.mail_merge_conditional_controls().unwrap();
        assert_eq!(document.mail_merge_conditional_control_count().unwrap(), 1);
        assert_eq!(controls.len(), 1);
        assert_eq!(
            controls[0].kind(),
            crate::docx::MailMergeConditionalControlKind::SkipIf
        );
        assert_eq!(controls[0].comparison(), "MERGEFIELD Order < 100");
        assert_eq!(controls[0].cached_result(), Some("cached skipif"));
    }

    #[test]
    fn writes_and_discovers_inert_if_fields_without_evaluation() {
        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        package.document_mut().unwrap().add_paragraph().add_field(
            crate::docx::writer::MutableField::with_result(
                r#"IF 1 = 1 "yes" "no""#.to_string(),
                "yes".to_string(),
            ),
        );
        package.save(file.path()).unwrap();

        let reopened = Package::open(file.path()).unwrap();
        let document = reopened.document().unwrap();
        let fields = document.if_fields().unwrap();
        assert_eq!(document.if_field_count().unwrap(), 1);
        assert_eq!(fields.len(), 1);
        assert_eq!(fields[0].expression(), r#"1 = 1 "yes" "no""#);
        assert_eq!(fields[0].cached_result(), Some("yes"));
    }

    #[test]
    fn writes_and_discovers_inert_document_state_fields() {
        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        {
            let document = package.document_mut().unwrap();
            for (instruction, cached_result) in [
                (
                    r#"SET RecipientName "North America" \* MERGEFORMAT"#,
                    "cached recipient",
                ),
                (r#"SEQ Figure FigureChapter \r 3 \* ARABIC"#, "3"),
                (r#"=SUM(ABOVE) \* MERGEFORMAT"#, "42"),
                (r#"STYLEREF "Heading 1" \n \p"#, "1 above"),
            ] {
                document.add_paragraph().add_field(
                    crate::docx::writer::MutableField::with_result(
                        instruction.to_string(),
                        cached_result.to_string(),
                    ),
                );
            }
        }
        package.save(file.path()).unwrap();

        let reopened = Package::open(file.path()).unwrap();
        let document = reopened.document().unwrap();

        let sets = document.set_fields().unwrap();
        assert_eq!(document.set_field_count().unwrap(), 1);
        assert_eq!(sets.len(), 1);
        assert_eq!(sets[0].target_name(), "RecipientName");
        assert_eq!(sets[0].expression(), r#""North America" \* MERGEFORMAT"#);
        assert_eq!(sets[0].cached_result(), Some("cached recipient"));

        let sequences = document.sequence_fields().unwrap();
        assert_eq!(document.sequence_field_count().unwrap(), 1);
        assert_eq!(sequences.len(), 1);
        assert_eq!(sequences[0].identifier(), "Figure");
        assert_eq!(sequences[0].bookmark(), Some("FigureChapter"));
        assert_eq!(sequences[0].tail(), r#"\r 3 \* ARABIC"#);
        assert_eq!(sequences[0].cached_result(), Some("3"));

        let formulas = document.formula_fields().unwrap();
        assert_eq!(document.formula_field_count().unwrap(), 1);
        assert_eq!(formulas.len(), 1);
        assert_eq!(formulas[0].formula(), r#"SUM(ABOVE) \* MERGEFORMAT"#);
        assert_eq!(formulas[0].cached_result(), Some("42"));

        let style_references = document.style_reference_fields().unwrap();
        assert_eq!(document.style_reference_field_count().unwrap(), 1);
        assert_eq!(style_references.len(), 1);
        assert_eq!(style_references[0].style_name(), "Heading 1");
        assert_eq!(
            style_references[0].options(),
            &[
                crate::docx::StyleReferenceFieldOption::ParagraphNumber,
                crate::docx::StyleReferenceFieldOption::RelativePosition,
            ]
        );
        assert_eq!(style_references[0].cached_result(), Some("1 above"));
    }

    #[test]
    fn writes_and_discovers_inert_bookmark_reference_fields() {
        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        {
            let document = package.document_mut().unwrap();
            for (instruction, cached_result) in [
                (
                    r#"REF "Target Bookmark" \d "-" \f \h \n \p \r \t \w"#,
                    "cached reference",
                ),
                (r#"PAGEREF PageTarget \h \p"#, "12 above"),
                (r#"FTNREF FootnoteTarget \p \f"#, "1 above"),
                (r#"NOTEREF EndnoteTarget \p \f"#, "i above"),
            ] {
                document.add_paragraph().add_field(
                    crate::docx::writer::MutableField::with_result(
                        instruction.to_string(),
                        cached_result.to_string(),
                    ),
                );
            }
        }
        package.save(file.path()).unwrap();

        let reopened = Package::open(file.path()).unwrap();
        let document = reopened.document().unwrap();
        let references = document.reference_fields().unwrap();
        assert_eq!(document.reference_field_count().unwrap(), 4);
        assert_eq!(references.len(), 4);

        assert_eq!(
            references[0].kind(),
            crate::docx::ReferenceFieldKind::Reference
        );
        assert_eq!(references[0].bookmark(), "Target Bookmark");
        assert_eq!(
            references[0].options(),
            &[
                crate::docx::ReferenceFieldOption::SequencePageSeparator("-".to_string()),
                crate::docx::ReferenceFieldOption::ReferencedNoteContent,
                crate::docx::ReferenceFieldOption::Hyperlink,
                crate::docx::ReferenceFieldOption::ParagraphNumberWithoutContext,
                crate::docx::ReferenceFieldOption::RelativePosition,
                crate::docx::ReferenceFieldOption::ParagraphNumberRelativeContext,
                crate::docx::ReferenceFieldOption::SuppressNonNumberText,
                crate::docx::ReferenceFieldOption::ParagraphNumberFullContext,
            ]
        );
        assert_eq!(references[0].cached_result(), Some("cached reference"));

        assert_eq!(
            references[1].kind(),
            crate::docx::ReferenceFieldKind::PageReference
        );
        assert_eq!(references[1].bookmark(), "PageTarget");
        assert_eq!(
            references[1].options(),
            &[
                crate::docx::ReferenceFieldOption::Hyperlink,
                crate::docx::ReferenceFieldOption::RelativePosition,
            ]
        );
        assert_eq!(references[1].cached_result(), Some("12 above"));

        assert_eq!(
            references[2].kind(),
            crate::docx::ReferenceFieldKind::FootnoteReference
        );
        assert_eq!(references[2].bookmark(), "FootnoteTarget");
        assert_eq!(
            references[2].options(),
            &[
                crate::docx::ReferenceFieldOption::RelativePosition,
                crate::docx::ReferenceFieldOption::NoteMarkFormatting,
            ]
        );
        assert_eq!(references[2].cached_result(), Some("1 above"));

        assert_eq!(
            references[3].kind(),
            crate::docx::ReferenceFieldKind::NoteReference
        );
        assert_eq!(references[3].bookmark(), "EndnoteTarget");
        assert_eq!(
            references[3].options(),
            &[
                crate::docx::ReferenceFieldOption::RelativePosition,
                crate::docx::ReferenceFieldOption::NoteMarkFormatting,
            ]
        );
        assert_eq!(references[3].cached_result(), Some("i above"));
    }

    #[test]
    fn writes_and_discovers_inert_equation_fields_without_calculation_or_rendering() {
        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        {
            let document = package.document_mut().unwrap();
            for (instruction, cached_result) in [
                (r#"EQ \o\ac(\fs24 Q,\fs16 R)"#, "cached equation"),
                (r#"EQ \f(1,2)"#, "1/2"),
                ("EQ", ""),
            ] {
                document.add_paragraph().add_field(
                    crate::docx::writer::MutableField::with_result(
                        instruction.to_string(),
                        cached_result.to_string(),
                    ),
                );
            }
        }
        package.save(file.path()).unwrap();

        let reopened = Package::open(file.path()).unwrap();
        let document = reopened.document().unwrap();
        let equations = document.equations().unwrap();
        assert_eq!(document.equation_count().unwrap(), 3);
        assert_eq!(equations.len(), 3);
        assert_eq!(equations[0].expression(), r#"\o\ac(\fs24 Q,\fs16 R)"#);
        assert_eq!(equations[0].cached_result(), Some("cached equation"));
        assert_eq!(equations[1].expression(), r#"\f(1,2)"#);
        assert_eq!(equations[1].cached_result(), Some("1/2"));
        assert_eq!(equations[2].expression(), "");
    }

    #[test]
    fn writes_and_discovers_inert_prompt_fields_without_displaying_prompts() {
        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        {
            let document = package.document_mut().unwrap();
            let ask = document.add_paragraph();
            ask.add_field(crate::docx::writer::MutableField::with_result(
                r#"ASK AskResponse "What is your first name?" \d "" \o"#.to_string(),
                "cached ask response".to_string(),
            ));
            let fill_in = document.add_paragraph();
            fill_in.add_field(crate::docx::writer::MutableField::with_result(
                r#"FILLIN "Enter appointment time" \d "09:00""#.to_string(),
                "10:30".to_string(),
            ));
        }
        package.save(file.path()).unwrap();

        let reopened = Package::open(file.path()).unwrap();
        let document = reopened.document().unwrap();
        let fields = document.prompt_fields().unwrap();
        assert_eq!(document.prompt_field_count().unwrap(), 2);
        assert_eq!(fields.len(), 2);
        assert_eq!(fields[0].kind(), crate::docx::PromptFieldKind::Ask);
        assert_eq!(fields[0].bookmark(), Some("AskResponse"));
        assert_eq!(fields[0].default_response(), Some(""));
        assert!(fields[0].prompts_once_per_mail_merge());
        assert_eq!(fields[0].cached_result(), Some("cached ask response"));
        assert_eq!(fields[1].kind(), crate::docx::PromptFieldKind::FillIn);
        assert_eq!(fields[1].bookmark(), None);
        assert_eq!(fields[1].prompt(), Some("Enter appointment time"));
        assert_eq!(fields[1].default_response(), Some("09:00"));
        assert_eq!(fields[1].cached_result(), Some("10:30"));
    }

    #[test]
    fn writes_and_discovers_inert_macro_button_fields_without_execution() {
        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        {
            let document = package.document_mut().unwrap();
            let paragraph = document.add_paragraph();
            paragraph.add_field(crate::docx::writer::MutableField::with_result(
                r#"MACROBUTTON NeverRun "Click here""#.to_string(),
                "cached button".to_string(),
            ));
        }
        package.save(file.path()).unwrap();

        let reopened = Package::open(file.path()).unwrap();
        let document = reopened.document().unwrap();
        let fields = document.macro_button_fields().unwrap();
        assert_eq!(document.macro_button_field_count().unwrap(), 1);
        assert_eq!(fields.len(), 1);
        assert_eq!(fields[0].macro_name(), "NeverRun");
        assert_eq!(fields[0].display_text(), "Click here");
        assert_eq!(fields[0].cached_result(), Some("cached button"));
    }

    #[test]
    fn writes_and_discovers_inert_active_and_building_block_fields() {
        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        {
            let document = package.document_mut().unwrap();
            for (instruction, cached_result) in [
                ("ADDIN opaque-add-in-data", "cached add-in"),
                ("CONTROL opaque-control-data", "cached control"),
                ("HTMLCONTROL opaque-html-data", "cached HTML control"),
                (r#"GLOSSARY "Legacy Clause" \* MERGEFORMAT"#, "cached glossary"),
                (r#"AUTOTEXT "Reusable Clause" \q opaque"#, "cached auto text"),
                (
                    r#"AUTOTEXTLIST "Choose a name" \s "Name Style" \t "Select one""#,
                    "cached auto text list",
                ),
            ] {
                document.add_paragraph().add_field(
                    crate::docx::writer::MutableField::with_result(
                        instruction.to_string(),
                        cached_result.to_string(),
                    ),
                );
            }
        }
        package.save(file.path()).unwrap();

        let reopened = Package::open(file.path()).unwrap();
        let document = reopened.document().unwrap();

        let active_content = document.active_content_fields().unwrap();
        assert_eq!(document.active_content_field_count().unwrap(), 3);
        assert_eq!(active_content.len(), 3);
        assert_eq!(
            active_content[0].kind(),
            crate::docx::ActiveContentFieldKind::AddIn
        );
        assert_eq!(
            active_content[1].kind(),
            crate::docx::ActiveContentFieldKind::OcxControl
        );
        assert_eq!(
            active_content[2].kind(),
            crate::docx::ActiveContentFieldKind::HtmlControl
        );
        assert_eq!(active_content[2].cached_result(), Some("cached HTML control"));

        let auto_text = document.auto_text_fields().unwrap();
        assert_eq!(document.auto_text_field_count().unwrap(), 2);
        assert_eq!(auto_text.len(), 2);
        assert_eq!(auto_text[0].kind(), crate::docx::AutoTextFieldKind::Glossary);
        assert_eq!(auto_text[0].entry_name(), "Legacy Clause");
        assert_eq!(auto_text[1].kind(), crate::docx::AutoTextFieldKind::AutoText);
        assert_eq!(auto_text[1].entry_name(), "Reusable Clause");

        let auto_text_lists = document.auto_text_list_fields().unwrap();
        assert_eq!(document.auto_text_list_field_count().unwrap(), 1);
        assert_eq!(auto_text_lists.len(), 1);
        assert_eq!(auto_text_lists[0].display_text(), Some("Choose a name"));
        assert_eq!(auto_text_lists[0].cached_result(), Some("cached auto text list"));
    }

    #[test]
    fn writes_and_discovers_typed_inert_dde_fields() {
        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        {
            let document = package.document_mut().unwrap();
            let manual = document.add_paragraph();
            manual.add_field(crate::docx::writer::MutableField::with_result(
                r#"DDE Excel "missing.xlsx" "Sheet1!A1" \a \p"#.to_string(),
                "cached DDE link".to_string(),
            ));
            let automatic = document.add_paragraph();
            automatic.add_field(crate::docx::writer::MutableField::with_result(
                r#"DDEAUTO Excel "missing.xlsx" "Sheet1!A2" \t"#.to_string(),
                "cached DDE auto link".to_string(),
            ));
        }
        package.save(file.path()).unwrap();

        let reopened = Package::open(file.path()).unwrap();
        let document = reopened.document().unwrap();
        let links = document.dde_links().unwrap();
        assert_eq!(document.dde_link_count().unwrap(), 2);
        assert_eq!(links.len(), 2);
        assert_eq!(links[0].kind(), crate::docx::DdeFieldKind::Dde);
        assert_eq!(links[0].application(), "Excel");
        assert_eq!(links[0].source(), "missing.xlsx");
        assert_eq!(links[0].item(), Some("Sheet1!A1"));
        assert!(links[0].requests_automatic_updates());
        assert_eq!(
            links[0].representation(),
            Some(crate::docx::DdeRepresentation::Picture)
        );
        assert_eq!(links[0].cached_result(), Some("cached DDE link"));
        assert_eq!(links[1].kind(), crate::docx::DdeFieldKind::DdeAuto);
        assert_eq!(links[1].item(), Some("Sheet1!A2"));
        assert!(links[1].requests_automatic_updates());
        assert_eq!(
            links[1].representation(),
            Some(crate::docx::DdeRepresentation::Text)
        );
        assert_eq!(links[1].cached_result(), Some("cached DDE auto link"));
    }

    #[test]
    fn writes_and_discovers_typed_inert_external_include_fields() {
        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        {
            let document = package.document_mut().unwrap();
            let text = document.add_paragraph();
            text.add_field(crate::docx::writer::MutableField::with_result(
                r#"INCLUDETEXT "file:///no-contact/source.docx" Summary \! \c Word8 \x /resume/name"#
                    .to_string(),
                "cached included text".to_string(),
            ));
            let picture = document.add_paragraph();
            picture.add_field(crate::docx::writer::MutableField::with_result(
                r#"INCLUDEPICTURE "file:///no-contact/picture.gif" \c Pictim32 \d"#.to_string(),
                "cached picture".to_string(),
            ));
            let legacy_text = document.add_paragraph();
            legacy_text.add_field(crate::docx::writer::MutableField::with_result(
                r#"INCLUDE "file:///no-contact/legacy.docx" LegacySection \!"#.to_string(),
                "cached legacy text".to_string(),
            ));
            let legacy_picture = document.add_paragraph();
            legacy_picture.add_field(crate::docx::writer::MutableField::with_result(
                r#"IMPORT "file:///no-contact/legacy.wmf" \c GraphicsFilter \d"#.to_string(),
                "cached legacy picture".to_string(),
            ));
        }
        package.save(file.path()).unwrap();

        let reopened = Package::open(file.path()).unwrap();
        let document = reopened.document().unwrap();
        let includes = document.external_includes().unwrap();
        assert_eq!(document.external_include_count().unwrap(), 4);
        assert_eq!(includes.len(), 4);
        assert_eq!(includes[0].kind(), crate::docx::IncludeFieldKind::Text);
        assert_eq!(includes[0].source(), "file:///no-contact/source.docx");
        assert_eq!(includes[0].bookmark(), Some("Summary"));
        assert!(includes[0].suppresses_nested_field_updates());
        assert_eq!(
            includes[0].options(),
            &[
                crate::docx::ExternalIncludeOption::Converter("Word8".to_string()),
                crate::docx::ExternalIncludeOption::XPath("/resume/name".to_string()),
            ]
        );
        assert_eq!(includes[0].cached_result(), Some("cached included text"));
        assert_eq!(includes[1].kind(), crate::docx::IncludeFieldKind::Picture);
        assert_eq!(includes[1].source(), "file:///no-contact/picture.gif");
        assert!(includes[1].omits_picture_data());
        assert_eq!(
            includes[1].options(),
            &[crate::docx::ExternalIncludeOption::Converter(
                "Pictim32".to_string()
            )]
        );
        assert_eq!(includes[1].cached_result(), Some("cached picture"));
        assert_eq!(includes[2].kind(), crate::docx::IncludeFieldKind::Text);
        assert_eq!(includes[2].source(), "file:///no-contact/legacy.docx");
        assert_eq!(includes[2].bookmark(), Some("LegacySection"));
        assert!(includes[2].suppresses_nested_field_updates());
        assert_eq!(includes[2].cached_result(), Some("cached legacy text"));
        assert_eq!(includes[3].kind(), crate::docx::IncludeFieldKind::Picture);
        assert_eq!(includes[3].source(), "file:///no-contact/legacy.wmf");
        assert!(includes[3].omits_picture_data());
        assert_eq!(
            includes[3].options(),
            &[crate::docx::ExternalIncludeOption::Converter(
                "GraphicsFilter".to_string()
            )]
        );
        assert_eq!(includes[3].cached_result(), Some("cached legacy picture"));
    }

    #[test]
    fn writes_and_discovers_typed_inert_referenced_document_fields() {
        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        {
            let document = package.document_mut().unwrap();
            let relative = document.add_paragraph();
            relative.add_field(crate::docx::writer::MutableField::with_result(
                r#"RD "C:\\Manual\\Chapters\\Chapter 1.docx" \p"#.to_string(),
                "cached relative reference".to_string(),
            ));
            let absolute = document.add_paragraph();
            absolute.add_field(crate::docx::writer::MutableField::with_result(
                r#"RD "file:///no-contact/appendix.docx""#.to_string(),
                "cached absolute reference".to_string(),
            ));
        }
        package.save(file.path()).unwrap();

        let reopened = Package::open(file.path()).unwrap();
        let document = reopened.document().unwrap();
        let references = document.referenced_documents().unwrap();
        assert_eq!(document.referenced_document_count().unwrap(), 2);
        assert_eq!(references.len(), 2);
        assert_eq!(references[0].source(), r"C:\Manual\Chapters\Chapter 1.docx");
        assert!(references[0].uses_relative_path());
        assert_eq!(
            references[0].cached_result(),
            Some("cached relative reference")
        );
        assert_eq!(references[1].source(), "file:///no-contact/appendix.docx");
        assert!(!references[1].uses_relative_path());
        assert_eq!(
            references[1].cached_result(),
            Some("cached absolute reference")
        );
    }

    #[test]
    fn writes_and_discovers_typed_inert_link_fields() {
        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        {
            let document = package.document_mut().unwrap();
            let spreadsheet = document.add_paragraph();
            spreadsheet.add_field(crate::docx::writer::MutableField::with_result(
                r#"LINK Excel.Sheet.8 "missing.xlsx" "Sheet1!A1" \a \f 4 \p"#.to_string(),
                "cached spreadsheet link".to_string(),
            ));
            let text = document.add_paragraph();
            text.add_field(crate::docx::writer::MutableField::with_result(
                r#"LINK Word.Document.8 "missing.docx" Bookmark \t"#.to_string(),
                "cached text link".to_string(),
            ));
        }
        package.save(file.path()).unwrap();

        let reopened = Package::open(file.path()).unwrap();
        let document = reopened.document().unwrap();
        let links = document.link_fields().unwrap();
        assert_eq!(document.link_field_count().unwrap(), 2);
        assert_eq!(links.len(), 2);
        assert_eq!(links[0].application_type(), "Excel.Sheet.8");
        assert_eq!(links[0].source(), "missing.xlsx");
        assert_eq!(links[0].item(), Some("Sheet1!A1"));
        assert!(links[0].requests_automatic_updates());
        assert_eq!(
            links[0].formatting_modes(),
            &[crate::docx::LinkFormatting::SpreadsheetSource]
        );
        assert_eq!(
            links[0].effective_result_option(),
            Some(crate::docx::LinkResultOption::Picture)
        );
        assert_eq!(links[0].cached_result(), Some("cached spreadsheet link"));
        assert_eq!(links[1].application_type(), "Word.Document.8");
        assert_eq!(links[1].item(), Some("Bookmark"));
        assert_eq!(
            links[1].effective_result_option(),
            Some(crate::docx::LinkResultOption::Text)
        );
        assert_eq!(links[1].cached_result(), Some("cached text link"));
    }

    #[test]
    fn saves_and_discovers_typed_inert_bibliography_source_stores() {
        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        package
            .add_custom_xml_data_store(NewCustomXmlDataStore {
                xml: br#"<b:Sources xmlns:b="http://schemas.openxmlformats.org/officeDocument/2006/bibliography" SelectedStyle="/APA.XSL" StyleName="APA"><b:Source><b:Tag>Doe2024</b:Tag><b:SourceType>Book</b:SourceType><b:Title>Stored source</b:Title></b:Source></b:Sources>"#.to_vec(),
                content_type: "application/xml".to_string(),
                item_id: "{22222222-2222-2222-2222-222222222222}".to_string(),
                schema_references: vec![
                    crate::docx::OOXML_BIBLIOGRAPHY_NAMESPACE.to_string(),
                ],
                conformance: crate::custom_xml_data::CustomXmlConformance::Transitional,
            })
            .unwrap();
        package.save(file.path()).unwrap();

        let reopened = Package::open(file.path()).unwrap();
        let stores = reopened.bibliography_source_stores().unwrap();
        assert_eq!(stores.len(), 1);
        assert_eq!(
            stores[0].data_store_item_id(),
            Some("{22222222-2222-2222-2222-222222222222}")
        );
        assert_eq!(stores[0].selected_style(), Some("/APA.XSL"));
        assert_eq!(stores[0].style_name(), Some("APA"));
        assert_eq!(stores[0].source_count(), 1);
        assert_eq!(stores[0].sources()[0].tag(), Some("Doe2024"));
        assert_eq!(stores[0].sources()[0].source_type(), Some("Book"));
        assert_eq!(stores[0].sources()[0].title(), Some("Stored source"));

        let sources = reopened.bibliography_sources().unwrap();
        assert_eq!(reopened.bibliography_source_count().unwrap(), 1);
        assert_eq!(sources.len(), 1);
        assert_eq!(sources[0].tag(), Some("Doe2024"));
    }

    #[test]
    fn writes_and_discovers_typed_index_fields() {
        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        {
            let document = package.document_mut().unwrap();
            let index = document.add_paragraph();
            index.add_field(crate::docx::writer::MutableField::with_result(
                r#"INDEX \c 2 \f "topics" \r"#.to_string(),
                "Topic\t3".to_string(),
            ));
            let entry = document.add_paragraph();
            entry.add_field(crate::docx::writer::MutableField::with_result(
                r#"XE "Topic" \f "topics" \r TopicRange \b"#.to_string(),
                "hidden marker".to_string(),
            ));
        }
        package.save(file.path()).unwrap();

        let reopened = Package::open(file.path()).unwrap();
        let document = reopened.document().unwrap();
        let indexes = document.indexes().unwrap();
        assert_eq!(document.index_count().unwrap(), 1);
        assert_eq!(indexes.len(), 1);
        assert_eq!(indexes[0].columns().unwrap(), Some(2));
        assert_eq!(indexes[0].entry_identifier().unwrap(), Some("topics"));
        assert!(indexes[0].runs_subentries_inline());

        let entries = document.index_entries().unwrap();
        assert_eq!(document.index_entry_count().unwrap(), 1);
        assert_eq!(entries.len(), 1);
        assert_eq!(entries[0].entry(), "Topic");
        assert_eq!(entries[0].entry_identifier().unwrap(), Some("topics"));
        assert_eq!(entries[0].page_range_bookmark().unwrap(), Some("TopicRange"));
        assert!(entries[0].is_bold());
    }

    #[test]
    fn body_edits_preserve_settings_part_byte_for_byte() {
        let mut package = Package::new().unwrap();
        let settings_uri = PackURI::new("/word/settings.xml").unwrap();
        let before = package.opc.get_part(&settings_uri).unwrap().blob().to_vec();

        package
            .document_mut()
            .unwrap()
            .add_paragraph_with_text("body-only edit");
        package.to_stream(Cursor::new(Vec::new())).unwrap();

        assert_eq!(package.opc.get_part(&settings_uri).unwrap().blob(), before);
    }

    #[test]
    fn body_edits_preserve_web_settings_part_byte_for_byte() {
        let mut package = Package::new().unwrap();
        let web_settings_uri = PackURI::new("/word/webSettings.xml").unwrap();
        let before = package
            .opc
            .get_part(&web_settings_uri)
            .unwrap()
            .blob()
            .to_vec();

        package
            .document_mut()
            .unwrap()
            .add_paragraph_with_text("body-only edit");
        package.to_stream(Cursor::new(Vec::new())).unwrap();

        assert_eq!(
            package.opc.get_part(&web_settings_uri).unwrap().blob(),
            before
        );
    }

    #[test]
    fn edits_web_settings_without_rewriting_document_content() {
        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        package
            .web_settings_mut()
            .unwrap()
            .set_allow_png(false)
            .set_optimize_for_browser(true)
            .set_target_screen_size(crate::docx::web_settings::TargetScreenSize::Pixels1600x1200);
        package.save(file.path()).unwrap();

        let reopened = Package::open(file.path()).unwrap();
        let settings = reopened
            .document()
            .unwrap()
            .web_settings()
            .unwrap()
            .unwrap();
        assert_eq!(settings.allow_png(), Some(false));
        assert_eq!(settings.optimize_for_browser(), Some(true));
        assert_eq!(
            settings.target_screen_size(),
            Some(crate::docx::web_settings::TargetScreenSize::Pixels1600x1200)
        );
    }

    #[test]
    fn web_settings_updates_preserve_frame_relationship_ids() {
        use crate::docx::web_settings::{FrameLayout, Frameset};

        const FRAME_RELATIONSHIP: &str =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/frame";
        let mut package = Package::new().unwrap();
        let web_settings_uri = PackURI::new("/word/webSettings.xml").unwrap();
        let relationship_id = package
            .opc
            .get_part_mut(&web_settings_uri)
            .unwrap()
            .relate_to("frame1.html", FRAME_RELATIONSHIP);

        let mut frameset = Frameset::default();
        frameset.set_layout(FrameLayout::Rows);
        frameset
            .add_frame()
            .set_name("main")
            .set_source_file_relationship_id(&relationship_id);
        package.web_settings_mut().unwrap().set_frameset(frameset);
        package.to_stream(Cursor::new(Vec::new())).unwrap();

        let part = package.opc.get_part(&web_settings_uri).unwrap();
        let relationship = part.rels().get(&relationship_id).unwrap();
        assert_eq!(relationship.reltype(), FRAME_RELATIONSHIP);
        assert_eq!(relationship.target_ref(), "frame1.html");
        assert!(
            package
                .document()
                .unwrap()
                .web_settings()
                .unwrap()
                .is_some()
        );
    }

    #[test]
    fn creates_a_web_settings_relationship_when_missing() {
        use litchi_opc::constants::relationship_type as rt;

        let mut package = Package::new().unwrap();
        let doc_uri = PackURI::new("/word/document.xml").unwrap();
        let relationship_id = package
            .opc
            .get_part(&doc_uri)
            .unwrap()
            .rels()
            .part_with_reltype(rt::WEB_SETTINGS)
            .unwrap()
            .r_id()
            .to_owned();
        package
            .opc
            .get_part_mut(&doc_uri)
            .unwrap()
            .rels_mut()
            .remove(&relationship_id);

        package.web_settings_mut().unwrap().set_encoding("utf-8");
        package.to_stream(Cursor::new(Vec::new())).unwrap();

        let relationship = package
            .opc
            .get_part(&doc_uri)
            .unwrap()
            .rels()
            .part_with_reltype(rt::WEB_SETTINGS)
            .unwrap();
        assert_eq!(relationship.target_ref(), "webSettings.xml");
        assert_eq!(
            package
                .document()
                .unwrap()
                .web_settings()
                .unwrap()
                .unwrap()
                .encoding(),
            Some("utf-8")
        );
    }

    #[test]
    fn reads_and_updates_strict_web_settings_relationships() {
        use crate::docx::web_settings::STRICT_WEB_SETTINGS_RELATIONSHIP;
        use litchi_opc::constants::relationship_type as rt;

        let mut package = Package::new().unwrap();
        let doc_uri = PackURI::new("/word/document.xml").unwrap();
        let (relationship_id, target_ref) = {
            let relationship = package
                .opc
                .get_part(&doc_uri)
                .unwrap()
                .rels()
                .part_with_reltype(rt::WEB_SETTINGS)
                .unwrap();
            (
                relationship.r_id().to_owned(),
                relationship.target_ref().to_owned(),
            )
        };
        let document_part = package.opc.get_part_mut(&doc_uri).unwrap();
        document_part.rels_mut().remove(&relationship_id);
        document_part.rels_mut().add_relationship(
            STRICT_WEB_SETTINGS_RELATIONSHIP.to_owned(),
            target_ref,
            relationship_id.clone(),
            false,
        );

        assert!(
            package
                .document()
                .unwrap()
                .web_settings()
                .unwrap()
                .is_some()
        );
        package
            .web_settings_mut()
            .unwrap()
            .set_save_smart_tags_as_xml(true);
        package.to_stream(Cursor::new(Vec::new())).unwrap();

        let relationship = package
            .opc
            .get_part(&doc_uri)
            .unwrap()
            .rels()
            .get(&relationship_id)
            .unwrap();
        assert_eq!(relationship.reltype(), STRICT_WEB_SETTINGS_RELATIONSHIP);
        assert_eq!(
            package
                .document()
                .unwrap()
                .web_settings()
                .unwrap()
                .unwrap()
                .save_smart_tags_as_xml(),
            Some(true)
        );
    }

    #[test]
    fn rejects_ambiguous_or_external_web_settings_relationships() {
        use crate::docx::web_settings::STRICT_WEB_SETTINGS_RELATIONSHIP;
        use litchi_opc::constants::relationship_type as rt;

        let mut duplicate = Package::new().unwrap();
        let doc_uri = PackURI::new("/word/document.xml").unwrap();
        duplicate
            .opc
            .get_part_mut(&doc_uri)
            .unwrap()
            .rels_mut()
            .add_relationship(
                STRICT_WEB_SETTINGS_RELATIONSHIP.to_owned(),
                "webSettings.xml".to_owned(),
                "rIdDuplicateWebSettings".to_owned(),
                false,
            );
        assert!(duplicate.document().unwrap().web_settings().is_err());
        assert!(duplicate.web_settings_mut().is_err());

        let mut external = Package::new().unwrap();
        let relationship_id = external
            .opc
            .get_part(&doc_uri)
            .unwrap()
            .rels()
            .part_with_reltype(rt::WEB_SETTINGS)
            .unwrap()
            .r_id()
            .to_owned();
        let document_part = external.opc.get_part_mut(&doc_uri).unwrap();
        document_part.rels_mut().remove(&relationship_id);
        document_part.rels_mut().add_relationship(
            rt::WEB_SETTINGS.to_owned(),
            "https://example.invalid/webSettings.xml".to_owned(),
            relationship_id,
            true,
        );
        assert!(external.document().unwrap().web_settings().is_err());
        assert!(external.web_settings_mut().is_err());
    }

    fn settings_state(package: &Package) -> (Vec<u8>, String) {
        let settings_uri = PackURI::new("/word/settings.xml").unwrap();
        let part = package.opc.get_part(&settings_uri).unwrap();
        (part.blob().to_vec(), part.rels().to_xml())
    }

    #[test]
    fn adds_replaces_removes_and_reopens_attached_template() {
        use crate::docx::settings::is_attached_template_relationship;

        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        package
            .set_attached_template_uri("file:///templates/Corporate.dotx")
            .unwrap();
        let attached = package.attached_template().unwrap().unwrap();
        let relationship_id = attached.relationship_id().to_owned();
        assert_eq!(attached.target_uri(), "file:///templates/Corporate.dotx");

        package
            .set_attached_template_uri("https://example.test/New.dotx?a=1&b=2")
            .unwrap();
        let replacement = package.attached_template().unwrap().unwrap();
        assert_eq!(replacement.relationship_id(), relationship_id);
        assert_eq!(
            replacement.target_uri(),
            "https://example.test/New.dotx?a=1&b=2"
        );
        let settings_uri = PackURI::new("/word/settings.xml").unwrap();
        let settings_part = package.opc.get_part(&settings_uri).unwrap();
        assert_eq!(
            settings_part
                .rels()
                .iter()
                .filter(|relationship| is_attached_template_relationship(relationship.reltype()))
                .count(),
            1
        );

        package.save(file.path()).unwrap();
        let mut reopened = Package::open(file.path()).unwrap();
        assert_eq!(
            reopened
                .attached_template()
                .unwrap()
                .unwrap()
                .target_uri(),
            "https://example.test/New.dotx?a=1&b=2"
        );
        let removed = reopened.remove_attached_template().unwrap().unwrap();
        assert_eq!(removed.relationship_id(), relationship_id);
        assert!(reopened.attached_template().unwrap().is_none());
        let part = reopened.opc.get_part(&settings_uri).unwrap();
        assert!(!String::from_utf8_lossy(part.blob()).contains("attachedTemplate"));
        assert!(!part
            .rels()
            .iter()
            .any(|relationship| is_attached_template_relationship(relationship.reltype())));
    }

    #[test]
    fn attached_template_mutation_preserves_unrelated_xml_and_relationships() {
        let mut package = Package::new().unwrap();
        let settings_uri = PackURI::new("/word/settings.xml").unwrap();
        let part = package.opc.get_part_mut(&settings_uri).unwrap();
        part.set_blob(br#"<?xml version="1.0"?><q:settings xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x="urn:opaque"><!--keep--><q:zoom q:percent="137"/><x:opaque><![CDATA[a < b]]></x:opaque></q:settings>"#.to_vec());
        part.rels_mut().add_relationship(
            "urn:unrelated".to_owned(),
            "https://example.test/keep?a=1&b=2".to_owned(),
            "customRelationship".to_owned(),
            true,
        );

        package
            .set_attached_template_uri("file:///templates/Keep.dotx")
            .unwrap();
        let part = package.opc.get_part(&settings_uri).unwrap();
        let xml = String::from_utf8_lossy(part.blob());
        assert!(xml.contains(r#"<!--keep--><q:zoom q:percent="137"/><x:opaque><![CDATA[a < b]]></x:opaque>"#));
        let unrelated = part.rels().get("customRelationship").unwrap();
        assert_eq!(unrelated.reltype(), "urn:unrelated");
        assert_eq!(unrelated.target_ref(), "https://example.test/keep?a=1&b=2");
    }

    #[test]
    fn reads_strict_prefixed_attached_template_and_rewrites_word_compatible_type() {
        use crate::docx::settings::STRICT_ATTACHED_TEMPLATE_RELATIONSHIP;

        let mut package = Package::new().unwrap();
        let settings_uri = PackURI::new("/word/settings.xml").unwrap();
        let part = package.opc.get_part_mut(&settings_uri).unwrap();
        part.set_blob(br#"<s:settings xmlns:s="http://purl.oclc.org/ooxml/wordprocessingml/main" xmlns:rel="http://purl.oclc.org/ooxml/officeDocument/relationships"><s:attachedTemplate rel:id="arbitrary-id"/></s:settings>"#.to_vec());
        part.rels_mut().add_relationship(
            STRICT_ATTACHED_TEMPLATE_RELATIONSHIP.to_owned(),
            "file:///strict.dotx".to_owned(),
            "arbitrary-id".to_owned(),
            true,
        );
        assert_eq!(
            package.attached_template().unwrap().unwrap().target_uri(),
            "file:///strict.dotx"
        );

        package
            .set_attached_template_uri("file:///compatible.dotx")
            .unwrap();
        let part = package.opc.get_part(&settings_uri).unwrap();
        let relationship = part.rels().get("arbitrary-id").unwrap();
        assert_eq!(relationship.reltype(), ATTACHED_TEMPLATE_RELATIONSHIP);
        assert!(relationship.is_external());
        assert!(String::from_utf8_lossy(part.blob())
            .contains(r#"<s:attachedTemplate rel:id="arbitrary-id"/>"#));
    }

    #[test]
    fn attached_template_failures_are_atomic() {
        let mut invalid_target = Package::new().unwrap();
        let before = settings_state(&invalid_target);
        assert!(invalid_target
            .set_attached_template_uri("file:///bad path.dotx")
            .is_err());
        assert_eq!(settings_state(&invalid_target), before);

        let mut malformed = Package::new().unwrap();
        malformed
            .set_attached_template_uri("file:///valid.dotx")
            .unwrap();
        let settings_uri = PackURI::new("/word/settings.xml").unwrap();
        malformed
            .opc
            .get_part_mut(&settings_uri)
            .unwrap()
            .rels_mut()
            .add_relationship(
                ATTACHED_TEMPLATE_RELATIONSHIP.to_owned(),
                "file:///duplicate.dotx".to_owned(),
                "duplicate-id".to_owned(),
                true,
            );
        let before = settings_state(&malformed);
        assert!(malformed
            .set_attached_template_uri("file:///replacement.dotx")
            .is_err());
        assert_eq!(settings_state(&malformed), before);
        assert!(malformed.remove_attached_template().is_err());
        assert_eq!(settings_state(&malformed), before);
    }

    #[test]
    fn protection_rewrite_preserves_attached_template_relationship() {
        use crate::docx::settings::ProtectionType;

        let mut package = Package::new().unwrap();
        package
            .set_attached_template_uri("file:///templates/Protected.dotx")
            .unwrap();
        let relationship_id = package
            .attached_template()
            .unwrap()
            .unwrap()
            .relationship_id()
            .to_owned();
        package
            .document_mut()
            .unwrap()
            .set_protection(ProtectionType::ReadOnly);
        package.to_stream(Cursor::new(Vec::new())).unwrap();

        let attached = package.attached_template().unwrap().unwrap();
        assert_eq!(attached.relationship_id(), relationship_id);
        assert_eq!(attached.target_uri(), "file:///templates/Protected.dotx");
        let settings_uri = PackURI::new("/word/settings.xml").unwrap();
        assert!(package
            .opc
            .get_part(&settings_uri)
            .unwrap()
            .rels()
            .get(&relationship_id)
            .is_some());
    }

    #[test]
    fn document_variable_package_lifecycle_is_deterministic_and_reopens() {
        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        assert_eq!(
            package
                .set_document_variable("Company & Team", "A < B")
                .unwrap(),
            None
        );
        package.set_document_variable("second", "two").unwrap();
        assert_eq!(
            package
                .set_document_variable("Company & Team", "updated")
                .unwrap(),
            Some("A < B".into())
        );
        assert_eq!(
            package.document_variables().unwrap().unwrap().names(),
            vec!["Company & Team", "second"]
        );
        package.save(file.path()).unwrap();

        let mut reopened = Package::open(file.path()).unwrap();
        let variables = reopened.document().unwrap().document_variables().unwrap().unwrap();
        assert_eq!(variables.get("Company & Team"), Some("updated"));
        assert_eq!(variables.get("second"), Some("two"));
        assert_eq!(
            reopened.remove_document_variable("Company & Team").unwrap(),
            Some("updated".into())
        );
        assert_eq!(reopened.clear_document_variables().unwrap(), 1);
        assert!(reopened.document_variables().unwrap().unwrap().is_empty());
        let settings_uri = PackURI::new("/word/settings.xml").unwrap();
        assert!(!String::from_utf8_lossy(
            reopened.opc.get_part(&settings_uri).unwrap().blob()
        )
        .contains("docVars"));
    }

    #[test]
    fn document_variable_mutation_preserves_xml_relationships_and_protection() {
        use crate::docx::settings::ProtectionType;

        let mut package = Package::new().unwrap();
        let settings_uri = PackURI::new("/word/settings.xml").unwrap();
        package.opc.get_part_mut(&settings_uri).unwrap().set_blob(
            br#"<?xml version="1.0"?><q:settings xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x="urn:opaque"><!--keep--><q:zoom q:percent="133"/><x:opaque><![CDATA[a < b]]></x:opaque><q:rsids/></q:settings>"#.to_vec(),
        );
        package
            .opc
            .get_part_mut(&settings_uri)
            .unwrap()
            .rels_mut()
            .add_relationship(
                "urn:unrelated".to_owned(),
                "https://example.test/keep".to_owned(),
                "keep-id".to_owned(),
                true,
            );
        package
            .set_attached_template_uri("file:///templates/Variables.dotx")
            .unwrap();
        let attached_id = package
            .attached_template()
            .unwrap()
            .unwrap()
            .relationship_id()
            .to_owned();
        package.set_document_variable("project", "Litchi").unwrap();
        let part = package.opc.get_part(&settings_uri).unwrap();
        let xml = String::from_utf8_lossy(part.blob());
        assert!(xml.contains(
            r#"<!--keep--><q:zoom q:percent="133"/><x:opaque><![CDATA[a < b]]></x:opaque>"#
        ));
        assert!(part.rels().get("keep-id").is_some());
        assert!(part.rels().get(&attached_id).is_some());

        package
            .document_mut()
            .unwrap()
            .set_protection(ProtectionType::ReadOnly);
        package.to_stream(Cursor::new(Vec::new())).unwrap();
        assert_eq!(
            package
                .document_variables()
                .unwrap()
                .unwrap()
                .get("project"),
            Some("Litchi")
        );
        let part = package.opc.get_part(&settings_uri).unwrap();
        assert!(part.rels().get("keep-id").is_some());
        assert!(part.rels().get(&attached_id).is_some());
    }

    #[test]
    fn document_variable_mutation_failures_are_atomic() {
        let mut package = Package::new().unwrap();
        let before = settings_state(&package);
        assert!(package.set_document_variable("", "invalid").is_err());
        assert_eq!(settings_state(&package), before);
        assert!(package
            .set_document_variable("too-long", "x".repeat(65_281))
            .is_err());
        assert_eq!(settings_state(&package), before);

        let settings_uri = PackURI::new("/word/settings.xml").unwrap();
        package.opc.get_part_mut(&settings_uri).unwrap().set_blob(
            br#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docVars><w:docVar w:name="duplicate" w:val="one"/><w:docVar w:name="duplicate" w:val="two"/></w:docVars></w:settings>"#.to_vec(),
        );
        let malformed = settings_state(&package);
        assert!(package.set_document_variable("new", "value").is_err());
        assert_eq!(settings_state(&package), malformed);
        assert!(package.remove_document_variable("duplicate").is_err());
        assert_eq!(settings_state(&package), malformed);
        assert!(package.clear_document_variables().is_err());
        assert_eq!(settings_state(&package), malformed);
    }

    #[test]
    fn reads_strict_settings_relationship_and_mce_fallback_variables() {
        use litchi_opc::constants::relationship_type as rt;

        const STRICT_SETTINGS_RELATIONSHIP: &str =
            "http://purl.oclc.org/ooxml/officeDocument/relationships/settings";
        let mut package = Package::new().unwrap();
        let doc_uri = PackURI::new("/word/document.xml").unwrap();
        let relationship_id = package
            .opc
            .get_part(&doc_uri)
            .unwrap()
            .rels()
            .part_with_reltype(rt::SETTINGS)
            .unwrap()
            .r_id()
            .to_owned();
        let document = package.opc.get_part_mut(&doc_uri).unwrap();
        let target = document
            .rels_mut()
            .remove(&relationship_id)
            .unwrap()
            .target_ref()
            .to_owned();
        document.rels_mut().add_relationship(
            STRICT_SETTINGS_RELATIONSHIP.to_owned(),
            target,
            relationship_id,
            false,
        );
        let settings_uri = PackURI::new("/word/settings.xml").unwrap();
        package.opc.get_part_mut(&settings_uri).unwrap().set_blob(
            br#"<s:settings xmlns:s="http://purl.oclc.org/ooxml/wordprocessingml/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:u="urn:unsupported" mc:Ignorable="u"><mc:AlternateContent><mc:Choice Requires="u"><s:docVars><s:docVar s:name="choice" s:val="ignored"/></s:docVars></mc:Choice><mc:Fallback><s:docVars><s:docVar s:name="fallback" s:val="selected"/></s:docVars></mc:Fallback></mc:AlternateContent></s:settings>"#.to_vec(),
        );

        let package_variables = package.document_variables().unwrap().unwrap();
        assert_eq!(package_variables.get("fallback"), Some("selected"));
        assert!(!package_variables.contains("choice"));
        let document_variables = package
            .document()
            .unwrap()
            .document_variables()
            .unwrap()
            .unwrap();
        assert_eq!(document_variables.get("fallback"), Some("selected"));
    }
}
