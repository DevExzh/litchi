use crate::common::unit::{emu_to_pt_f64, emu_to_px_96};
/// Image reading support for DOCX documents.
///
/// This module provides structures and functions for extracting images from Word documents.
/// Images in DOCX are embedded within `<w:drawing>` elements in paragraphs, and the actual
/// binary data is stored in separate media parts linked via relationships.
///
/// # Architecture
///
/// - `InlineImage`: Represents an inline image within a paragraph run
/// - Image data is accessed through relationships using the `r:embed` attribute
/// - Supports all standard image formats (PNG, JPEG, GIF, BMP, TIFF, EMF, WMF)
///
/// # Example
///
/// ```rust,no_run
/// use litchi::ooxml::docx::Package;
///
/// let pkg = Package::open("document.docx")?;
/// let doc = pkg.document()?;
///
/// // Extract all images from the document
/// for para in doc.paragraphs()? {
///     for image in para.images()? {
///         println!("Image: {} ({}x{} EMUs)",
///             image.description(),
///             image.width_emu(),
///             image.height_emu()
///         );
///         let data = image.data()?;
///         // Process image data...
///     }
/// }
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
use crate::ooxml::docx::format::ImageFormat;
use crate::ooxml::error::{OoxmlError, Result};
use crate::ooxml::opc::OpcPackage;
use crate::ooxml::opc::rel::Relationships;
use quick_xml::Reader;
use quick_xml::events::Event;
use smallvec::SmallVec;
use std::borrow::Cow;

/// An inline image embedded in a Word document.
///
/// Represents an image within a `<w:drawing>` element. The image contains
/// metadata such as dimensions and description, with the actual binary data
/// stored in a related media part.
///
/// # Performance
///
/// Image metadata is stored inline for fast access. Binary data is loaded
/// lazily on demand via `data()` to minimize memory usage.
///
/// # Field Ordering
///
/// Fields are ordered to maximize CPU cache line utilization and minimize padding:
/// - Strings (24 bytes each on 64-bit systems) are placed together
/// - i64 values (8 bytes each) are placed together
#[derive(Debug, Clone)]
pub struct InlineImage {
    /// Relationship ID referencing the image part (e.g., "rId5")
    r_embed: String,

    /// Image description/alt text
    description: String,

    /// Image name/title
    name: String,

    /// Width in EMUs (English Metric Units, 1 inch = 914400 EMUs)
    width_emu: i64,

    /// Height in EMUs
    height_emu: i64,
}

impl InlineImage {
    /// Create a new InlineImage from parsed attributes.
    ///
    /// This is typically called internally during XML parsing.
    ///
    /// # Arguments
    ///
    /// * `r_embed` - Relationship ID for the image part
    /// * `width_emu` - Width in EMUs
    /// * `height_emu` - Height in EMUs
    /// * `description` - Image description/alt text
    /// * `name` - Image name/title
    #[inline]
    pub fn new(
        r_embed: String,
        width_emu: i64,
        height_emu: i64,
        description: String,
        name: String,
    ) -> Self {
        Self {
            r_embed,
            width_emu,
            height_emu,
            description,
            name,
        }
    }

    /// Get the relationship ID for this image.
    #[inline]
    pub fn r_embed(&self) -> &str {
        &self.r_embed
    }

    /// Get the width in EMUs (English Metric Units).
    ///
    /// EMUs are used throughout Office Open XML. 1 inch = 914400 EMUs.
    #[inline]
    pub fn width_emu(&self) -> i64 {
        self.width_emu
    }

    /// Get the height in EMUs (English Metric Units).
    #[inline]
    pub fn height_emu(&self) -> i64 {
        self.height_emu
    }

    /// Get the width in pixels (assuming 96 DPI).
    #[inline]
    pub fn width_px(&self) -> u32 {
        emu_to_px_96(self.width_emu)
    }

    /// Get the height in pixels (assuming 96 DPI).
    #[inline]
    pub fn height_px(&self) -> u32 {
        emu_to_px_96(self.height_emu)
    }

    /// Get the width in points.
    #[inline]
    pub fn width_pt(&self) -> f64 {
        emu_to_pt_f64(self.width_emu)
    }

    /// Get the height in points.
    #[inline]
    pub fn height_pt(&self) -> f64 {
        emu_to_pt_f64(self.height_emu)
    }

    /// Get the image description/alt text.
    #[inline]
    pub fn description(&self) -> &str {
        &self.description
    }

    /// Get the image name/title.
    #[inline]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Load the binary image data from the package.
    ///
    /// This resolves the relationship ID and loads the actual image bytes
    /// from the media part.
    ///
    /// # Arguments
    ///
    /// * `opc` - Reference to the OPC package
    /// * `rels` - Relationships from the document part
    ///
    /// # Performance
    ///
    /// Image data is loaded lazily to minimize memory usage. The binary
    /// data is borrowed from the package without copying when possible.
    pub fn data<'a>(&self, opc: &'a OpcPackage, rels: &Relationships) -> Result<Cow<'a, [u8]>> {
        // Resolve the relationship to get the image part name
        let rel = rels.get(&self.r_embed).ok_or_else(|| {
            OoxmlError::InvalidFormat(format!("Image relationship not found: {}", self.r_embed))
        })?;

        // Get the target part name
        let part_name = rel.target_partname().map_err(|e| {
            OoxmlError::InvalidFormat(format!("Invalid image part reference: {}", e))
        })?;

        // Get the part from the package
        let part = opc
            .get_part(&part_name)
            .map_err(|e| OoxmlError::InvalidFormat(format!("Image part not found: {}", e)))?;

        // Return the binary data as a borrowed slice (zero-copy)
        Ok(Cow::Borrowed(part.blob()))
    }

    /// Detect the image format from binary data.
    ///
    /// # Arguments
    ///
    /// * `opc` - Reference to the OPC package
    /// * `rels` - Relationships from the document part
    pub fn format(&self, opc: &OpcPackage, rels: &Relationships) -> Result<ImageFormat> {
        let data = self.data(opc, rels)?;
        ImageFormat::detect_from_bytes(&data)
            .ok_or_else(|| OoxmlError::InvalidFormat("Unable to detect image format".to_string()))
    }
}

/// Parse inline images from paragraph XML.
///
/// Extracts all `<w:drawing>` elements containing `<wp:inline>` images
/// from the paragraph XML.
///
/// # Arguments
///
/// * `xml_bytes` - The raw XML bytes of the paragraph
///
/// # Performance
///
/// Uses streaming XML parsing with pre-allocated SmallVec for efficient
/// storage of typically small image collections. Avoids unnecessary allocations
/// by reusing buffers.
///
/// # Example XML Structure
///
/// ```xml
/// <w:drawing>
///   <wp:inline>
///     <wp:extent cx="914400" cy="914400"/>
///     <wp:docPr name="Picture 1" descr="Description"/>
///     <a:graphic>
///       <a:graphicData>
///         <pic:pic>
///           <pic:blipFill>
///             <a:blip r:embed="rId5"/>
///           </pic:blipFill>
///         </pic:pic>
///       </a:graphicData>
///     </a:graphic>
///   </wp:inline>
/// </w:drawing>
/// ```
pub(crate) fn parse_inline_images(xml_bytes: &[u8]) -> Result<SmallVec<[InlineImage; 4]>> {
    let mut reader = Reader::from_reader(xml_bytes);
    reader.config_mut().trim_text(true);

    // Use SmallVec for efficient storage of typically small image collections
    let mut images = SmallVec::new();

    // State tracking for parsing
    let mut in_drawing = false;
    let mut in_inline = false;
    let mut in_blip_fill = false;

    // Image attributes being built
    let mut width_emu: i64 = 914400; // Default 1 inch
    let mut height_emu: i64 = 914400; // Default 1 inch
    let mut description = String::new();
    let mut name = String::new();
    let mut r_embed = String::new();

    loop {
        match reader.read_event() {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                let local_name_ref = e.local_name();
                let local_name = local_name_ref.as_ref();

                match local_name {
                    b"drawing" => {
                        in_drawing = true;
                    },
                    b"inline" if in_drawing => {
                        in_inline = true;
                        // Reset state for new image
                        width_emu = 914400;
                        height_emu = 914400;
                        description.clear();
                        name.clear();
                        r_embed.clear();
                    },
                    b"extent" if in_inline => {
                        // Parse width and height from extent element
                        // <wp:extent cx="914400" cy="914400"/>
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"cx" => {
                                    if let Ok(s) = std::str::from_utf8(&attr.value) {
                                        width_emu = s.parse().unwrap_or(914400);
                                    }
                                },
                                b"cy" => {
                                    if let Ok(s) = std::str::from_utf8(&attr.value) {
                                        height_emu = s.parse().unwrap_or(914400);
                                    }
                                },
                                _ => {},
                            }
                        }
                    },
                    b"docPr" if in_inline => {
                        // Parse name and description from docPr element
                        // <wp:docPr id="1" name="Picture" descr="Description"/>
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"name" => {
                                    if let Ok(s) = std::str::from_utf8(&attr.value) {
                                        name = s.to_string();
                                    }
                                },
                                b"descr" => {
                                    if let Ok(s) = std::str::from_utf8(&attr.value) {
                                        description = s.to_string();
                                    }
                                },
                                _ => {},
                            }
                        }
                    },
                    b"blipFill" if in_inline => {
                        in_blip_fill = true;
                    },
                    b"blip" if in_blip_fill => {
                        // Parse r:embed attribute from blip element
                        // <a:blip r:embed="rId5"/>
                        for attr in e.attributes().flatten() {
                            let key = attr.key.as_ref();
                            // Check for r:embed (with namespace prefix)
                            if (key == b"r:embed" || key.ends_with(b":embed"))
                                && let Ok(s) = std::str::from_utf8(&attr.value)
                            {
                                r_embed = s.to_string();
                            }
                        }
                    },
                    _ => {},
                }
            },
            Ok(Event::End(e)) => {
                let local_name_ref = e.local_name();
                let local_name = local_name_ref.as_ref();

                match local_name {
                    b"drawing" => {
                        in_drawing = false;
                    },
                    b"inline" if in_inline => {
                        // Finished parsing an inline image
                        in_inline = false;

                        // Only add if we found a valid r:embed
                        if !r_embed.is_empty() {
                            images.push(InlineImage::new(
                                r_embed.clone(),
                                width_emu,
                                height_emu,
                                description.clone(),
                                name.clone(),
                            ));
                        }
                    },
                    b"blipFill" => {
                        in_blip_fill = false;
                    },
                    _ => {},
                }
            },
            Ok(Event::Eof) => break,
            Err(e) => return Err(OoxmlError::Xml(e.to_string())),
            _ => {},
        }
    }

    Ok(images)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_inline_image_dimensions() {
        let img = InlineImage::new(
            "rId5".to_string(),
            914400,  // 1 inch
            1828800, // 2 inches
            "Test Image".to_string(),
            "Picture 1".to_string(),
        );

        assert_eq!(img.width_emu(), 914400);
        assert_eq!(img.height_emu(), 1828800);
        assert_eq!(img.width_px(), 96); // 1 inch at 96 DPI
        assert_eq!(img.height_px(), 192); // 2 inches at 96 DPI
        assert!((img.width_pt() - 72.0).abs() < 0.1); // 1 inch = 72 points
        assert!((img.height_pt() - 144.0).abs() < 0.1); // 2 inches = 144 points
    }

    #[test]
    fn test_parse_inline_images_empty() {
        let xml = b"<w:p><w:r><w:t>Text only</w:t></w:r></w:p>";
        let images = parse_inline_images(xml).unwrap();
        assert_eq!(images.len(), 0);
    }

    #[test]
    fn test_parse_inline_images_single() {
        let xml = br#"<w:p>
            <w:r>
                <w:drawing>
                    <wp:inline>
                        <wp:extent cx="1000000" cy="2000000"/>
                        <wp:docPr id="1" name="MyImage" descr="Test description"/>
                        <a:graphic>
                            <a:graphicData>
                                <pic:pic>
                                    <pic:blipFill>
                                        <a:blip r:embed="rId7"/>
                                    </pic:blipFill>
                                </pic:pic>
                            </a:graphicData>
                        </a:graphic>
                    </wp:inline>
                </w:drawing>
            </w:r>
        </w:p>"#;

        let images = parse_inline_images(xml).unwrap();
        assert_eq!(images.len(), 1);

        let img = &images[0];
        assert_eq!(img.r_embed(), "rId7");
        assert_eq!(img.width_emu(), 1000000);
        assert_eq!(img.height_emu(), 2000000);
        assert_eq!(img.name(), "MyImage");
        assert_eq!(img.description(), "Test description");
    }

    #[test]
    fn test_parse_inline_images_multiple() {
        let xml = br#"<w:p>
            <w:r>
                <w:drawing>
                    <wp:inline>
                        <wp:extent cx="1000000" cy="1000000"/>
                        <wp:docPr name="Image1"/>
                        <pic:blipFill><a:blip r:embed="rId1"/></pic:blipFill>
                    </wp:inline>
                </w:drawing>
            </w:r>
            <w:r>
                <w:drawing>
                    <wp:inline>
                        <wp:extent cx="2000000" cy="2000000"/>
                        <wp:docPr name="Image2"/>
                        <pic:blipFill><a:blip r:embed="rId2"/></pic:blipFill>
                    </wp:inline>
                </w:drawing>
            </w:r>
        </w:p>"#;

        let images = parse_inline_images(xml).unwrap();
        assert_eq!(images.len(), 2);
        assert_eq!(images[0].r_embed(), "rId1");
        assert_eq!(images[1].r_embed(), "rId2");
    }
}
