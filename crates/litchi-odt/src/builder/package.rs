//! ODT package assembly for the builder.

use super::model::Builder;
use crate::Package;
use crate::core::PackageWriter;
use litchi_core::Result;

impl Builder {
    /// Build the document into an immutable [`Package`].
    ///
    /// Output (`save`, `save_with_durability`, `to_bytes`, `write_to`) lives
    /// on the returned package. Building never consults the clock or any
    /// other ambient state: two builders with equal inputs produce
    /// byte-identical packages.
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi_odt::Builder;
    ///
    /// # fn main() -> Result<(), litchi_odt::Error> {
    /// let mut builder = Builder::new();
    /// builder.add_paragraph("Hello, World!")?;
    /// let package = builder.build()?;
    /// assert_eq!(package.document()?.text()?, "Hello, World!");
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// # Errors
    ///
    /// Returns an error when an authored part cannot be serialized.
    pub fn build(self) -> std::result::Result<Package, crate::Error> {
        Ok(Package::from_authored(
            self.build_package(crate::constants::ODF_TEXT)?,
        )?)
    }

    /// Build the document package bytes with an explicit root MIME type.
    ///
    /// Used by the web-template authoring model to emit the legacy
    /// `application/vnd.oasis.opendocument.text-web` MIME type.
    pub(crate) fn build_package(self, mimetype: &str) -> Result<Vec<u8>> {
        let mut writer =
            PackageWriter::new(litchi_odf_common::core::StreamingArchiveLimits::default());

        // Set MIME type
        writer.set_mimetype(mimetype)?;

        // Add content.xml
        let content_xml = self.generate_content_xml()?;
        writer.add_file("content.xml", content_xml.as_bytes())?;

        // Add styles.xml with list styles
        let styles_xml = self.generate_styles_xml()?;
        writer.add_file("styles.xml", styles_xml.as_bytes())?;

        // Add meta.xml
        let meta_xml = self.generate_meta_xml();
        writer.add_file("meta.xml", meta_xml.as_bytes())?;

        // Finish and return bytes
        writer.finish_to_bytes()
    }
}
