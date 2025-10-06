/// Package implementation for PowerPoint presentations.
use crate::ooxml::error::{OoxmlError, Result};
use crate::ooxml::opc::constants::content_type as ct;
use crate::ooxml::opc::OpcPackage;
use crate::ooxml::pptx::presentation::Presentation;
use crate::ooxml::pptx::parts::PresentationPart;
use std::io::{Read, Seek};
use std::path::Path;

/// A PowerPoint (.pptx) package.
///
/// This is the main entry point for working with PowerPoint presentations.
/// It wraps an OPC package and provides PowerPoint-specific functionality.
///
/// # Examples
///
/// ```rust,no_run
/// use litchi::ooxml::pptx::Package;
///
/// // Open an existing presentation
/// let pkg = Package::open("presentation.pptx")?;
///
/// // Get the main presentation
/// let pres = pkg.presentation()?;
/// 
/// // Access slides
/// println!("Presentation has {} slides", pres.slide_count()?);
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
pub struct Package {
    /// The underlying OPC package
    opc: OpcPackage,
}

impl Package {
    /// Open a .pptx package from a file path.
    ///
    /// # Arguments
    ///
    /// * `path` - Path to the .pptx file
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::pptx::Package;
    ///
    /// let pkg = Package::open("presentation.pptx")?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self> {
        let opc = OpcPackage::open(path)?;

        // Verify it's a PowerPoint presentation by checking the main part's content type
        let main_part = opc
            .main_document_part()
            .map_err(|e| OoxmlError::PartNotFound(format!("main presentation part: {}", e)))?;

        let content_type = main_part.content_type();
        // Support both regular and macro-enabled presentations
        if content_type != ct::PML_PRESENTATION_MAIN && content_type != ct::PML_PRES_MACRO_MAIN {
            return Err(OoxmlError::InvalidContentType {
                expected: format!(
                    "{} or {}",
                    ct::PML_PRESENTATION_MAIN,
                    ct::PML_PRES_MACRO_MAIN
                ),
                got: content_type.to_string(),
            });
        }

        Ok(Self { opc })
    }

    /// Create a .pptx package from a reader.
    ///
    /// # Arguments
    ///
    /// * `reader` - A reader containing the .pptx file data (must implement Read + Seek)
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::pptx::Package;
    /// use std::io::Cursor;
    ///
    /// let data = std::fs::read("presentation.pptx")?;
    /// let cursor = Cursor::new(data);
    /// let pkg = Package::from_reader(cursor)?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn from_reader<R: Read + Seek>(reader: R) -> Result<Self> {
        let opc = OpcPackage::from_reader(reader)?;

        // Verify it's a PowerPoint presentation by checking the main part's content type
        let main_part = opc
            .main_document_part()
            .map_err(|e| OoxmlError::PartNotFound(format!("main presentation part: {}", e)))?;

        let content_type = main_part.content_type();
        // Support both regular and macro-enabled presentations
        if content_type != ct::PML_PRESENTATION_MAIN && content_type != ct::PML_PRES_MACRO_MAIN {
            return Err(OoxmlError::InvalidContentType {
                expected: format!(
                    "{} or {}",
                    ct::PML_PRESENTATION_MAIN,
                    ct::PML_PRES_MACRO_MAIN
                ),
                got: content_type.to_string(),
            });
        }

        Ok(Self { opc })
    }

    /// Get the main presentation.
    ///
    /// Returns the `Presentation` object which provides access to the presentation's
    /// content, slides, and other features.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::pptx::Package;
    ///
    /// let pkg = Package::open("presentation.pptx")?;
    /// let pres = pkg.presentation()?;
    /// 
    /// // Access slides
    /// for slide in pres.slides()? {
    ///     println!("Slide text: {}", slide.text()?);
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn presentation(&self) -> Result<Presentation> {
        let main_part = self
            .opc
            .main_document_part()
            .map_err(|e| OoxmlError::PartNotFound(format!("main presentation part: {}", e)))?;

        // Create PresentationPart wrapper
        let pres_part = PresentationPart::from_part(main_part)?;

        // Create and return Presentation
        Ok(Presentation::new(pres_part, &self.opc))
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
        let result = Package::open("test.pptx");
        assert!(result.is_ok());
    }
}

