//! ODT package assembly for the builder.

use super::model::Builder;
use crate::core::PackageWriter;
use litchi_core::Result;
use std::path::Path;

impl Builder {
    /// Build the document and return as bytes
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odt::Builder;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut builder = Builder::new();
    /// builder.add_paragraph("Hello, World!")?;
    /// let bytes = builder.build()?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn build(self) -> Result<Vec<u8>> {
        self.build_package(crate::constants::ODF_TEXT)
    }

    /// Build the document package with an explicit root MIME type.
    ///
    /// Used by the web-template authoring model to emit the legacy
    /// `application/vnd.oasis.opendocument.text-web` MIME type.
    pub(crate) fn build_package(self, mimetype: &str) -> Result<Vec<u8>> {
        let mut writer = PackageWriter::new();

        // Set MIME type
        writer.set_mimetype(mimetype)?;

        // Add content.xml
        let content_xml = self.generate_content_xml();
        writer.add_file("content.xml", content_xml.as_bytes())?;

        // Add styles.xml with list styles
        let styles_xml = self.generate_styles_xml();
        writer.add_file("styles.xml", styles_xml.as_bytes())?;

        // Add meta.xml
        let meta_xml = self.generate_meta_xml();
        writer.add_file("meta.xml", meta_xml.as_bytes())?;

        // Finish and return bytes
        writer.finish_to_bytes()
    }

    /// Build and save the document to a file
    ///
    /// # Arguments
    ///
    /// * `path` - Path where the ODT file should be saved
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odt::Builder;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut builder = Builder::new();
    /// builder.add_paragraph("Hello, World!")?;
    /// builder.save("output.odt")?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn save<P: AsRef<Path>>(self, path: P) -> Result<()> {
        let bytes = self.build()?;
        std::fs::write(path, bytes)?;
        Ok(())
    }
}
