//! Embedded OLE/package object authoring for DOCX documents.
//!
//! An embedded object is serialized as a `<w:object>` element wrapping a VML
//! shape (`v:shape`, optionally with a `v:imagedata` preview) and an
//! `o:OLEObject` descriptor carrying the ProgID, shape identity, and the
//! relationship ID of the payload part. The payload is stored verbatim as
//! `/word/embeddings/oleObjectN.bin` (content type
//! `application/vnd.openxmlformats-officedocument.oleObject`) with an
//! `oleObject` relationship from the main document part, matching the inert
//! discovery contract of [`litchi_ooxml_common::embedded`].
//!
//! Everything is inert: payloads are never parsed, activated, or executed,
//! and the optional preview is a plain image part.

use crate::docx::format::ImageFormat;
use crate::error::{OoxmlError, Result};
use litchi_core::unit::{EMUS_PER_INCH, EMUS_PER_PT, EMUS_PER_TWIP};
use litchi_core::xml::escape_xml;
use std::fmt::Write as FmtWrite;

/// Maximum embedded payload size accepted by the authoring API (64 MiB).
pub const MAX_OLE_PAYLOAD_BYTES: usize = 64 * 1024 * 1024;
/// Maximum ProgID length per the COM ProgID contract.
const MAX_PROG_ID_LENGTH: usize = 39;
/// Maximum VML shape ID length accepted by the authoring API.
const MAX_SHAPE_ID_LENGTH: usize = 255;

/// A mutable embedded OLE/package object being authored in a document.
///
/// Built with [`MutableOleObject::new`], then embedded via
/// [`crate::docx::writer::MutableDocument::add_ole_object`], which assigns the
/// VML shape identity and validates uniqueness. The payload bytes are stored
/// verbatim and never interpreted.
///
/// # Examples
///
/// ```rust,no_run
/// use litchi_ooxml::docx::{MutableOleObject, Package};
///
/// let mut package = Package::new()?;
/// let object = MutableOleObject::new("Package", b"opaque bytes".to_vec())?;
/// package.document_mut()?.add_ole_object(object)?;
/// package.save("out.docx")?;
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
#[derive(Debug)]
pub struct MutableOleObject {
    /// COM ProgID of the embedded object (e.g. `Package`, `Excel.Sheet.12`).
    pub(crate) prog_id: String,
    /// VML shape ID (`v:shape@id` / `o:OLEObject@ShapeID`); assigned by
    /// `MutableDocument::add_ole_object` when left empty.
    pub(crate) shape_id: String,
    /// OLE object ID (`o:OLEObject@ObjectID`); assigned alongside the shape ID.
    pub(crate) object_id: u32,
    /// Payload bytes stored verbatim as the embedded object part.
    pub(crate) payload: Vec<u8>,
    /// Shape width in EMUs (English Metric Units, 1 inch = 914400 EMUs).
    pub(crate) width_emu: i64,
    /// Shape height in EMUs.
    pub(crate) height_emu: i64,
    /// Optional preview image (data and detected format) shown in place of
    /// the object; stored as an ordinary media part.
    pub(crate) preview: Option<(Vec<u8>, ImageFormat)>,
}

impl MutableOleObject {
    /// Create an embedded object with a ProgID and verbatim payload bytes.
    ///
    /// The shape defaults to one square inch; resize with
    /// [`Self::set_size_emu`]. The payload must not exceed
    /// [`MAX_OLE_PAYLOAD_BYTES`].
    ///
    /// # Arguments
    ///
    /// * `prog_id` - COM ProgID (1–39 chars, letters/digits/`.`, starting
    ///   with a letter; e.g. `Package`, `Word.Document.12`)
    /// * `payload` - Opaque payload bytes, stored verbatim and never activated
    pub fn new(prog_id: impl Into<String>, payload: Vec<u8>) -> Result<Self> {
        let prog_id = prog_id.into();
        validate_prog_id(&prog_id)?;
        if payload.len() > MAX_OLE_PAYLOAD_BYTES {
            return Err(OoxmlError::InvalidFormat(format!(
                "OLE object payload exceeds {MAX_OLE_PAYLOAD_BYTES} bytes"
            )));
        }
        Ok(Self {
            prog_id,
            shape_id: String::new(),
            object_id: 0,
            payload,
            width_emu: EMUS_PER_INCH,
            height_emu: EMUS_PER_INCH,
            preview: None,
        })
    }

    /// Get the ProgID.
    pub fn prog_id(&self) -> &str {
        &self.prog_id
    }

    /// Get the payload bytes.
    pub fn payload(&self) -> &[u8] {
        &self.payload
    }

    /// Get the optional preview image bytes and detected format.
    pub fn preview(&self) -> Option<(&[u8], ImageFormat)> {
        self.preview
            .as_ref()
            .map(|(data, format)| (data.as_slice(), *format))
    }

    /// Get the assigned VML shape ID (empty until the object is added to a
    /// document or explicitly set).
    pub fn shape_id(&self) -> &str {
        &self.shape_id
    }

    /// Set the shape extents in EMUs (both must be positive).
    pub fn set_size_emu(&mut self, width_emu: i64, height_emu: i64) -> Result<&mut Self> {
        if width_emu <= 0 || height_emu <= 0 {
            return Err(OoxmlError::InvalidFormat(
                "OLE object extents must be positive EMU values".to_string(),
            ));
        }
        self.width_emu = width_emu;
        self.height_emu = height_emu;
        Ok(self)
    }

    /// Set an explicit VML shape ID instead of accepting the allocated one.
    ///
    /// The ID must be unique within the document;
    /// [`crate::docx::writer::MutableDocument::add_ole_object`] rejects
    /// collisions with existing shapes.
    pub fn set_shape_id(&mut self, shape_id: impl Into<String>) -> Result<&mut Self> {
        let shape_id = shape_id.into();
        validate_shape_id(&shape_id)?;
        self.shape_id = shape_id;
        Ok(self)
    }

    /// Set an optional preview image rendered in place of the object.
    ///
    /// The image format is detected from the bytes (PNG, JPEG, GIF, EMF, and
    /// the other formats supported by [`ImageFormat`]); the image is stored
    /// as an ordinary media part.
    pub fn set_preview(&mut self, data: Vec<u8>) -> Result<&mut Self> {
        let format = ImageFormat::detect_from_bytes(&data)
            .ok_or_else(|| OoxmlError::InvalidFormat("Unknown preview image format".to_string()))?;
        self.preview = Some((data, format));
        Ok(self)
    }

    /// Serialize the `<w:object>` element.
    ///
    /// `ole_rid` is the relationship ID of the payload part and `preview_rid`
    /// the relationship ID of the preview image part; when a relationship ID
    /// is unavailable (placeholder serialization), a deterministic
    /// `{{OLE_OBJECT_*}}` placeholder is emitted, matching the writer's
    /// hyperlink/image placeholder convention. The enclosing `<w:r>` wrapper
    /// is emitted by the paragraph writer.
    pub(crate) fn to_xml(
        &self,
        xml: &mut String,
        ole_rid: Option<&str>,
        preview_rid: Option<&str>,
    ) -> Result<()> {
        if self.shape_id.is_empty() {
            return Err(OoxmlError::InvalidFormat(
                "OLE object has no shape ID; add it via MutableDocument::add_ole_object"
                    .to_string(),
            ));
        }
        let ole_rid = ole_rid
            .map(str::to_owned)
            .unwrap_or_else(|| format!("{{{{OLE_OBJECT_{}}}}}", self.shape_id));
        let shape_id = escape_xml(&self.shape_id);
        let width_twips = self.width_emu / EMUS_PER_TWIP;
        let height_twips = self.height_emu / EMUS_PER_TWIP;
        let width_pt = self.width_emu as f64 / EMUS_PER_PT as f64;
        let height_pt = self.height_emu as f64 / EMUS_PER_PT as f64;
        write!(
            xml,
            r#"<w:object xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" w:dxaOrig="{width_twips}" w:dyaOrig="{height_twips}">"#,
        )
        .map_err(|error| OoxmlError::Xml(error.to_string()))?;
        write!(
            xml,
            r#"<v:shape id="{shape_id}" style="width:{width_pt}pt;height:{height_pt}pt" o:ole="">"#
        )
        .map_err(|error| OoxmlError::Xml(error.to_string()))?;
        if self.preview.is_some() {
            let preview_rid = preview_rid
                .map(str::to_owned)
                .unwrap_or_else(|| format!("{{{{OLE_PREVIEW_{}}}}}", self.shape_id));
            write!(xml, r#"<v:imagedata r:id="{preview_rid}" o:title=""/>"#)
                .map_err(|error| OoxmlError::Xml(error.to_string()))?;
        }
        xml.push_str("</v:shape>");
        write!(
            xml,
            r#"<o:OLEObject Type="Embed" ProgID="{}" ShapeID="{shape_id}" DrawAspect="Content" ObjectID="_{}" r:id="{ole_rid}"/>"#,
            escape_xml(&self.prog_id),
            self.object_id,
        )
        .map_err(|error| OoxmlError::Xml(error.to_string()))?;
        xml.push_str("</w:object>");
        Ok(())
    }
}

/// Validate a COM ProgID: 1–39 characters, ASCII letters/digits/`.`,
/// starting with a letter.
fn validate_prog_id(prog_id: &str) -> Result<()> {
    let mut characters = prog_id.chars();
    let valid = !prog_id.is_empty()
        && prog_id.len() <= MAX_PROG_ID_LENGTH
        && characters
            .next()
            .is_some_and(|first| first.is_ascii_alphabetic())
        && characters.all(|c| c.is_ascii_alphanumeric() || c == '.');
    if !valid {
        return Err(OoxmlError::InvalidFormat(format!(
            "invalid OLE ProgID '{prog_id}'"
        )));
    }
    Ok(())
}

/// Validate a VML shape ID: 1–255 characters of ASCII alphanumerics,
/// underscores, and hyphens.
pub(crate) fn validate_shape_id(shape_id: &str) -> Result<()> {
    let valid = !shape_id.is_empty()
        && shape_id.len() <= MAX_SHAPE_ID_LENGTH
        && shape_id
            .bytes()
            .all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'_' | b'-'));
    if !valid {
        return Err(OoxmlError::InvalidFormat(format!(
            "invalid OLE shape ID '{shape_id}'"
        )));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_ooxml_common::embedded::{Kind, Target};

    #[test]
    fn validates_prog_ids() {
        assert!(MutableOleObject::new("Package", vec![1]).is_ok());
        assert!(MutableOleObject::new("Excel.Sheet.12", vec![1]).is_ok());
        assert!(MutableOleObject::new("", vec![1]).is_err());
        assert!(MutableOleObject::new("1Package", vec![1]).is_err());
        assert!(MutableOleObject::new("Has Space", vec![1]).is_err());
        assert!(MutableOleObject::new("X".repeat(40), vec![1]).is_err());
    }

    #[test]
    fn rejects_oversized_payloads_and_bad_shape_ids() {
        assert!(MutableOleObject::new("Package", vec![0; MAX_OLE_PAYLOAD_BYTES + 1]).is_err());
        let mut object = MutableOleObject::new("Package", vec![1]).unwrap();
        assert!(object.set_shape_id("_x0000_i1025").is_ok());
        assert!(object.set_shape_id("bad id").is_err());
        assert!(object.set_shape_id("").is_err());
        assert!(object.set_size_emu(0, 914400).is_err());
        assert!(object.set_size_emu(914400, -1).is_err());
    }

    #[test]
    fn serializes_object_element_with_relationships() {
        let mut object = MutableOleObject::new("Package", b"payload".to_vec()).unwrap();
        object.shape_id = "_x0000_i1025".to_string();
        object.object_id = 1025;
        let mut xml = String::new();
        object.to_xml(&mut xml, Some("rId9"), None).unwrap();
        assert!(xml.contains(r#"<w:object xmlns:v="urn:schemas-microsoft-com:vml""#));
        assert!(xml.contains(r#"<v:shape id="_x0000_i1025" "#));
        assert!(xml.contains(r#"o:ole="">"#));
        assert!(xml.contains(
            r#"<o:OLEObject Type="Embed" ProgID="Package" ShapeID="_x0000_i1025" DrawAspect="Content" ObjectID="_1025" r:id="rId9"/>"#
        ));
        assert!(!xml.contains("imagedata"));
    }

    #[test]
    fn serializes_preview_and_placeholders() {
        const PNG_HEADER: &[u8] = &[0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A];
        let mut object = MutableOleObject::new("Package", b"payload".to_vec()).unwrap();
        object.set_preview(PNG_HEADER.to_vec()).unwrap();
        object.shape_id = "_x0000_i1026".to_string();
        object.object_id = 1026;

        let mut xml = String::new();
        object.to_xml(&mut xml, None, None).unwrap();
        assert!(xml.contains(r#"r:id="{{OLE_OBJECT__x0000_i1026}}""#));
        assert!(xml.contains(r#"<v:imagedata r:id="{{OLE_PREVIEW__x0000_i1026}}" o:title=""/>"#));

        let mut xml = String::new();
        object
            .to_xml(&mut xml, Some("rId9"), Some("rId10"))
            .unwrap();
        assert!(xml.contains(r#"<v:imagedata r:id="rId10" o:title=""/>"#));
    }

    #[test]
    fn rejects_serialization_without_shape_id() {
        let object = MutableOleObject::new("Package", b"payload".to_vec()).unwrap();
        assert!(object.to_xml(&mut String::new(), None, None).is_err());
    }

    #[test]
    fn round_trips_embedded_object_through_saved_package() {
        use crate::docx::Package;
        use tempfile::NamedTempFile;

        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let payload: Vec<u8> = (0u8..=255).collect();
        let mut package = Package::new().unwrap();
        package
            .document_mut()
            .unwrap()
            .add_paragraph_with_text("host paragraph");
        let object = MutableOleObject::new("Package", payload.clone()).unwrap();
        let object = package
            .document_mut()
            .unwrap()
            .add_ole_object(object)
            .unwrap();
        assert_eq!(object.shape_id(), "_x0000_i1025");
        package.save(file.path()).unwrap();

        let reopened = Package::open(file.path()).unwrap();
        let entries = reopened.embedded().unwrap();
        assert_eq!(entries.len(), 1);
        let entry = &entries[0];
        assert_eq!(entry.kind(), Kind::Object);
        assert_eq!(entry.source().as_str(), "/word/document.xml");
        let Target::Internal(discovered) = entry.target() else {
            panic!("expected an internal embedded payload")
        };
        assert_eq!(
            discovered.part().as_str(),
            "/word/embeddings/oleObject1.bin"
        );
        assert_eq!(
            discovered.content_type(),
            "application/vnd.openxmlformats-officedocument.oleObject"
        );
        assert_eq!(discovered.bytes(), payload.as_slice());
    }

    #[test]
    fn allocates_distinct_identities_for_multiple_objects() {
        use crate::docx::Package;
        use tempfile::NamedTempFile;

        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        {
            let document = package.document_mut().unwrap();
            let first = MutableOleObject::new("Package", b"first".to_vec()).unwrap();
            let first = document.add_ole_object(first).unwrap();
            assert_eq!(first.shape_id(), "_x0000_i1025");
            let second = MutableOleObject::new("Word.Document.12", b"second".to_vec()).unwrap();
            let second = document.add_ole_object(second).unwrap();
            assert_eq!(second.shape_id(), "_x0000_i1026");

            // An explicit shape ID colliding with an assigned one is rejected.
            let mut duplicate = MutableOleObject::new("Package", b"dup".to_vec()).unwrap();
            duplicate.set_shape_id("_x0000_i1025").unwrap();
            assert!(document.add_ole_object(duplicate).is_err());
        }
        package.save(file.path()).unwrap();

        let reopened = Package::open(file.path()).unwrap();
        let entries = reopened.embedded().unwrap();
        assert_eq!(entries.len(), 2);
        let mut part_names: Vec<String> = entries
            .iter()
            .map(|entry| match entry.target() {
                Target::Internal(payload) => payload.part().as_str().to_owned(),
                Target::External(_) => panic!("expected internal payloads"),
            })
            .collect();
        part_names.sort_unstable();
        assert_eq!(
            part_names,
            [
                "/word/embeddings/oleObject1.bin".to_owned(),
                "/word/embeddings/oleObject2.bin".to_owned()
            ]
        );
    }

    #[test]
    fn coexists_with_text_boxes_and_images() {
        use crate::docx::textbox::TextBoxAnchor;
        use crate::docx::{MutableTextBox, Package};
        use tempfile::NamedTempFile;

        const PNG_HEADER: &[u8] = &[0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A];

        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        {
            let document = package.document_mut().unwrap();

            let mut text_box = MutableTextBox::new("Companion Box", 914400, 457200).unwrap();
            text_box.add_run("box story");
            document.add_text_box(text_box);

            document
                .add_paragraph()
                .add_picture_from_bytes(PNG_HEADER.to_vec(), Some(914400), Some(914400))
                .unwrap();

            let mut object = MutableOleObject::new("Package", b"coexist payload".to_vec()).unwrap();
            object.set_preview(PNG_HEADER.to_vec()).unwrap();
            document.add_ole_object(object).unwrap();
        }
        package.save(file.path()).unwrap();

        let reopened = Package::open(file.path()).unwrap();
        let document = reopened.document().unwrap();

        let entries = reopened.embedded().unwrap();
        assert_eq!(entries.len(), 1);
        let Target::Internal(discovered) = entries[0].target() else {
            panic!("expected an internal embedded payload")
        };
        assert_eq!(discovered.bytes(), b"coexist payload");

        let text_boxes = document.text_boxes().unwrap();
        assert_eq!(text_boxes.len(), 1);
        assert_eq!(text_boxes[0].anchor, TextBoxAnchor::Inline);
        assert_eq!(text_boxes[0].text(), "box story");

        let picture_count = document
            .paragraphs()
            .unwrap()
            .iter()
            .map(|paragraph| {
                paragraph
                    .drawing_objects()
                    .unwrap()
                    .iter()
                    .filter(|drawing| drawing.name() == "Picture")
                    .count()
            })
            .sum::<usize>();
        assert_eq!(picture_count, 1);
    }
}
