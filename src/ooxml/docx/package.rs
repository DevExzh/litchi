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
