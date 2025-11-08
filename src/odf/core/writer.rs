//! ODF package writing functionality.
//!
//! This module provides utilities for creating and writing ODF files as ZIP archives,
//! including generating manifests and proper file structure.

use crate::common::{Error, Result};
use std::io::Write;
use zip::write::{SimpleFileOptions, ZipWriter};

/// Builder for creating ODF packages (ZIP archives)
///
/// This struct helps create valid ODF files by managing the ZIP archive structure,
/// manifest, and required files.
///
/// # Examples
///
/// ```no_run
/// # use litchi::odf::core::PackageWriter;
/// # use litchi::Result;
/// # fn example() -> Result<()> {
/// let mut writer = PackageWriter::new();
/// writer.set_mimetype("application/vnd.oasis.opendocument.text")?;
/// writer.add_file("content.xml", b"<office:document-content>...</office:document-content>")?;
/// writer.add_file("styles.xml", b"<office:document-styles>...</office:document-styles>")?;
/// writer.add_file("meta.xml", b"<office:document-meta>...</office:document-meta>")?;
///
/// let bytes = writer.finish()?;
/// std::fs::write("document.odt", bytes)?;
/// # Ok(())
/// # }
/// ```
pub struct PackageWriter<W: Write + std::io::Seek> {
    zip_writer: ZipWriter<W>,
    mimetype: Option<String>,
    manifest_entries: Vec<ManifestEntry>,
}

/// Entry in the ODF manifest
#[derive(Debug, Clone)]
struct ManifestEntry {
    full_path: String,
    media_type: String,
}

impl PackageWriter<std::io::Cursor<Vec<u8>>> {
    /// Create a new package writer that writes to memory
    pub fn new() -> Self {
        Self {
            zip_writer: ZipWriter::new(std::io::Cursor::new(Vec::new())),
            mimetype: None,
            manifest_entries: Vec::new(),
        }
    }
}

impl<W: Write + std::io::Seek> PackageWriter<W> {
    /// Create a new package writer with a custom writer
    #[allow(dead_code)] // Reserved for future use
    pub fn with_writer(writer: W) -> Self {
        Self {
            zip_writer: ZipWriter::new(writer),
            mimetype: None,
            manifest_entries: Vec::new(),
        }
    }

    /// Set the MIME type for the document
    ///
    /// This sets both the mimetype file and the root manifest entry.
    ///
    /// # Arguments
    ///
    /// * `mimetype` - MIME type string (e.g., "application/vnd.oasis.opendocument.text")
    pub fn set_mimetype(&mut self, mimetype: &str) -> Result<()> {
        self.mimetype = Some(mimetype.to_string());

        // Add root manifest entry
        self.manifest_entries.push(ManifestEntry {
            full_path: "/".to_string(),
            media_type: mimetype.to_string(),
        });

        Ok(())
    }

    /// Add a file to the package
    ///
    /// # Arguments
    ///
    /// * `path` - Path within the ZIP archive (e.g., "content.xml", "Pictures/image1.png")
    /// * `content` - File content as bytes
    ///
    /// # Note
    ///
    /// This method automatically adds the file to the manifest with an appropriate media type.
    pub fn add_file(&mut self, path: &str, content: &[u8]) -> Result<()> {
        // Determine media type based on file extension
        let media_type = Self::guess_media_type(path);

        // Add to manifest
        self.manifest_entries.push(ManifestEntry {
            full_path: path.to_string(),
            media_type: media_type.to_string(),
        });

        // Add to ZIP with no compression for mimetype, normal compression for others
        let options = if path == "mimetype" {
            SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored)
        } else {
            SimpleFileOptions::default().compression_method(zip::CompressionMethod::Deflated)
        };

        self.zip_writer.start_file(path, options)?;
        self.zip_writer.write_all(content)?;

        Ok(())
    }

    /// Add a file to the package with a specific media type
    ///
    /// # Arguments
    ///
    /// * `path` - Path within the ZIP archive
    /// * `content` - File content as bytes
    /// * `media_type` - MIME type for the manifest entry
    #[allow(dead_code)] // Reserved for future use
    pub fn add_file_with_media_type(
        &mut self,
        path: &str,
        content: &[u8],
        media_type: &str,
    ) -> Result<()> {
        // Add to manifest
        self.manifest_entries.push(ManifestEntry {
            full_path: path.to_string(),
            media_type: media_type.to_string(),
        });

        // Add to ZIP
        let options = if path == "mimetype" {
            SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored)
        } else {
            SimpleFileOptions::default().compression_method(zip::CompressionMethod::Deflated)
        };

        self.zip_writer.start_file(path, options)?;
        self.zip_writer.write_all(content)?;

        Ok(())
    }

    /// Generate the manifest.xml content
    fn generate_manifest(&self) -> String {
        let mut manifest = String::from(
            r#"<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3">
"#,
        );

        // Add manifest entries
        for entry in &self.manifest_entries {
            manifest.push_str(&format!(
                r#"  <manifest:file-entry manifest:full-path="{}" manifest:media-type="{}"/>
"#,
                Self::escape_xml(&entry.full_path),
                Self::escape_xml(&entry.media_type)
            ));
        }

        manifest.push_str("</manifest:manifest>\n");
        manifest
    }

    /// Escape XML special characters
    fn escape_xml(text: &str) -> String {
        text.replace('&', "&amp;")
            .replace('<', "&lt;")
            .replace('>', "&gt;")
            .replace('"', "&quot;")
            .replace('\'', "&apos;")
    }

    /// Guess media type from file path
    fn guess_media_type(path: &str) -> &'static str {
        if path.ends_with(".xml") {
            "text/xml"
        } else if path.ends_with(".png") {
            "image/png"
        } else if path.ends_with(".jpg") || path.ends_with(".jpeg") {
            "image/jpeg"
        } else if path.ends_with(".gif") {
            "image/gif"
        } else if path.ends_with(".svg") {
            "image/svg+xml"
        } else if path.ends_with("/") {
            "" // Directory entry
        } else {
            "application/octet-stream"
        }
    }

    /// Finish writing the package and return the result
    ///
    /// This method writes the mimetype file, manifest, and finalizes the ZIP archive.
    ///
    /// # Errors
    ///
    /// Returns an error if:
    /// - No MIME type has been set
    /// - Writing to the ZIP archive fails
    pub fn finish(mut self) -> Result<W> {
        let mimetype = self
            .mimetype
            .as_ref()
            .ok_or_else(|| Error::InvalidFormat("MIME type not set".to_string()))?
            .clone();

        // First, write the mimetype file (must be first and uncompressed per ODF spec)
        let options =
            SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
        self.zip_writer.start_file("mimetype", options)?;
        self.zip_writer.write_all(mimetype.as_bytes())?;

        // Add META-INF directory to manifest
        self.manifest_entries.push(ManifestEntry {
            full_path: "META-INF/".to_string(),
            media_type: String::new(),
        });

        // Generate and write manifest
        let manifest_content = self.generate_manifest();
        let options =
            SimpleFileOptions::default().compression_method(zip::CompressionMethod::Deflated);
        self.zip_writer
            .start_file("META-INF/manifest.xml", options)?;
        self.zip_writer.write_all(manifest_content.as_bytes())?;

        // Finish ZIP archive
        let writer = self.zip_writer.finish()?;
        Ok(writer)
    }
}

impl Default for PackageWriter<std::io::Cursor<Vec<u8>>> {
    fn default() -> Self {
        Self::new()
    }
}

impl PackageWriter<std::io::Cursor<Vec<u8>>> {
    /// Finish writing and return the bytes
    pub fn finish_to_bytes(self) -> Result<Vec<u8>> {
        let cursor = self.finish()?;
        Ok(cursor.into_inner())
    }
}

/// Helper to create standard ODF directory structure
pub struct OdfStructure;

impl OdfStructure {
    /// Generate a default content.xml skeleton
    #[allow(dead_code)] // Reserved for future use
    pub fn default_content_xml(office_type: &str) -> String {
        format!(
            r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
                          xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
                          xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
                          xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
                          xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
                          xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
                          xmlns:xlink="http://www.w3.org/1999/xlink"
                          xmlns:dc="http://purl.org/dc/elements/1.1/"
                          xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"
                          xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"
                          xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"
                          xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
                          xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0"
                          xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0"
                          xmlns:math="http://www.w3.org/1998/Math/MathML"
                          xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0"
                          xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0"
                          xmlns:ooo="http://openoffice.org/2004/office"
                          xmlns:ooow="http://openoffice.org/2004/writer"
                          xmlns:oooc="http://openoffice.org/2004/calc"
                          xmlns:dom="http://www.w3.org/2001/xml-events"
                          xmlns:xforms="http://www.w3.org/2002/xforms"
                          xmlns:xsd="http://www.w3.org/2001/XMLSchema"
                          xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
                          xmlns:rpt="http://openoffice.org/2005/report"
                          xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2"
                          xmlns:xhtml="http://www.w3.org/1999/xhtml"
                          xmlns:grddl="http://www.w3.org/2003/g/data-view#"
                          xmlns:tableooo="http://openoffice.org/2009/table"
                          xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0"
                          xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0"
                          xmlns:field="urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0"
                          xmlns:formx="urn:openoffice:names:experimental:ooxml-odf-interop:xmlns:form:1.0"
                          xmlns:css3t="http://www.w3.org/TR/css3-text/"
                          office:version="1.3">
  <office:scripts/>
  <office:font-face-decls/>
  <office:automatic-styles/>
  <office:body>
    <{office_type}>
    </{office_type}>
  </office:body>
</office:document-content>
"#
        )
    }

    /// Generate a default styles.xml skeleton
    pub fn default_styles_xml() -> String {
        r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
                         xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
                         xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
                         xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
                         xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
                         xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
                         xmlns:xlink="http://www.w3.org/1999/xlink"
                         xmlns:dc="http://purl.org/dc/elements/1.1/"
                         xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"
                         xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"
                         xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
                         xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0"
                         xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0"
                         xmlns:math="http://www.w3.org/1998/Math/MathML"
                         xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0"
                         xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0"
                         xmlns:ooo="http://openoffice.org/2004/office"
                         xmlns:ooow="http://openoffice.org/2004/writer"
                         xmlns:oooc="http://openoffice.org/2004/calc"
                         xmlns:dom="http://www.w3.org/2001/xml-events"
                         xmlns:rpt="http://openoffice.org/2005/report"
                         xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2"
                         xmlns:xhtml="http://www.w3.org/1999/xhtml"
                         xmlns:grddl="http://www.w3.org/2003/g/data-view#"
                         xmlns:tableooo="http://openoffice.org/2009/table"
                         xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0"
                         xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0"
                         xmlns:field="urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0"
                         xmlns:formx="urn:openoffice:names:experimental:ooxml-odf-interop:xmlns:form:1.0"
                         xmlns:css3t="http://www.w3.org/TR/css3-text/"
                         office:version="1.3">
  <office:font-face-decls/>
  <office:styles/>
  <office:automatic-styles/>
  <office:master-styles/>
</office:document-styles>
"#.to_string()
    }

    /// Generate a default meta.xml skeleton
    #[allow(dead_code)] // Reserved for future use
    pub fn default_meta_xml() -> String {
        let now = chrono::Utc::now().to_rfc3339();
        format!(
            r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
                       xmlns:xlink="http://www.w3.org/1999/xlink"
                       xmlns:dc="http://purl.org/dc/elements/1.1/"
                       xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"
                       xmlns:ooo="http://openoffice.org/2004/office"
                       xmlns:grddl="http://www.w3.org/2003/g/data-view#"
                       office:version="1.3">
  <office:meta>
    <meta:generator>Litchi/0.0.1</meta:generator>
    <meta:creation-date>{}</meta:creation-date>
    <dc:date>{}</dc:date>
  </office:meta>
</office:document-meta>
"#,
            now, now
        )
    }

    /// Generate a default settings.xml skeleton
    #[allow(dead_code)] // Will be used for future enhancements
    pub fn default_settings_xml() -> String {
        r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-settings xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
                           xmlns:xlink="http://www.w3.org/1999/xlink"
                           xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0"
                           xmlns:ooo="http://openoffice.org/2004/office"
                           office:version="1.3">
  <office:settings>
    <config:config-item-set config:name="ooo:view-settings">
      <config:config-item config:name="ViewAreaTop" config:type="long">0</config:config-item>
      <config:config-item config:name="ViewAreaLeft" config:type="long">0</config:config-item>
      <config:config-item config:name="ViewAreaWidth" config:type="long">1</config:config-item>
      <config:config-item config:name="ViewAreaHeight" config:type="long">1</config:config-item>
    </config:config-item-set>
  </office:settings>
</office:document-settings>
"#
        .to_string()
    }
}
