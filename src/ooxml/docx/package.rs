use crate::ooxml::docx::document::Document;
use crate::ooxml::docx::parts::DocumentPart;
/// Package implementation for Word documents.
use crate::ooxml::error::{OoxmlError, Result};
use crate::ooxml::opc::OpcPackage;
use crate::ooxml::opc::constants::content_type as ct;
use std::io::{Read, Seek};
use std::path::Path;

/// A Word (.docx) package.
///
/// This is the main entry point for working with Word documents.
/// It wraps an OPC package and provides Word-specific functionality.
///
/// # Examples
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
pub struct Package {
    /// The underlying OPC package
    opc: OpcPackage,
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

        // Create styles.xml part
        let styles_partname = PackURI::new("/word/styles.xml")
            .map_err(|e| OoxmlError::InvalidUri(format!("styles partname: {}", e)))?;
        let styles_part = BlobPart::new(
            styles_partname.clone(),
            ct::WML_STYLES.to_string(),
            template::default_styles_xml().as_bytes().to_vec(),
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

        Ok(Self { opc })
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

        Ok(Self { opc })
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

        Ok(Self { opc })
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

        Ok(Self { opc })
    }

    /// Get the main document.
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
    pub fn save<P: AsRef<Path>>(&self, path: P) -> Result<()> {
        self.opc.save(path).map_err(|e| {
            OoxmlError::IoError(std::io::Error::other(format!(
                "Failed to save package: {}",
                e
            )))
        })
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
