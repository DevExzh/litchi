#![expect(
    clippy::module_name_repetitions,
    reason = "public names retain established OOXML facade terminology"
)]
//! Embedded OLE object (`<w:object>`) reading support for DOCX documents.
//!
//! Word stores OLE-embedded objects (equation editor output, embedded
//! packages, linked documents) as a `<w:object>` run child wrapping a VML
//! shape with an optional `v:imagedata` preview and an `o:OLEObject`
//! descriptor carrying the `ProgID`, placement metadata, and the relationship
//! ID of the payload part (typically `word/embeddings/oleObjectN.bin`).
//!
//! # Architecture
//!
//! - `EmbeddedObject`: the typed reader-side view of one `<w:object>` element
//! - The VML preview reuses [`InlineImage`], so its bytes and format resolve
//!   through the same relationship and OPC plumbing as `<w:drawing>` images
//! - Payloads are inert: bytes are exposed verbatim and are never parsed,
//!   activated, or executed
//!
//! # Example XML Structure
//!
//! ```xml
//! <w:object>
//!   <v:shape id="_x0000_i1025" type="#_x0000_t75" alt="equation" style="width:16.7pt;height:14pt" o:ole="">
//!     <v:imagedata r:id="rId7" o:title="equation"/>
//!   </v:shape>
//!   <o:OLEObject Type="Embed" ProgID="Equation.DSMT4" ShapeID="_x0000_i1025" DrawAspect="Content" ObjectID="_1468075725" r:id="rId8"/>
//! </w:object>
//! ```
use crate::error::{Error, Result};
use crate::image::InlineImage;
use litchi_core::unit::{EMUS_PER_CM, EMUS_PER_INCH, EMUS_PER_MM, EMUS_PER_PT};
use litchi_ooxml_common::xml::attributes::first_wins;
use litchi_opc::OpcPackage;
use litchi_opc::rel::Relationships;
use quick_xml::Reader;
use quick_xml::events::Event;
use std::borrow::Cow;

/// EMUs per CSS pixel at 96 DPI.
const EMUS_PER_PX: i64 = EMUS_PER_INCH / 96;

/// An OLE object embedded in a Word document run (`<w:object>`).
///
/// The object couples an `o:OLEObject` descriptor (program ID, embedding
/// type, draw aspect, and the payload relationship) with a VML shape whose
/// `v:imagedata` supplies the preview image Word renders in place of the
/// object. The payload is inert: it is exposed verbatim and never activated.
#[derive(Debug, Clone)]
pub struct EmbeddedObject {
    /// VML preview image shown in place of the object, when the source
    /// declares a `v:imagedata` relationship.
    preview: Option<InlineImage>,

    /// COM program ID (`o:OLEObject` `ProgID`), when declared.
    prog_id: Option<String>,

    /// Embedding type (`o:OLEObject` `Type`), when declared.
    object_type: Option<String>,

    /// Draw aspect (`o:OLEObject` `DrawAspect`), when declared.
    draw_aspect: Option<String>,

    /// Relationship ID of the embedded payload part (`o:OLEObject` `r:id`),
    /// when declared.
    ole_r_id: Option<String>,
}

impl EmbeddedObject {
    /// Get the VML preview image rendered in place of the object.
    ///
    /// The returned [`InlineImage`] resolves its bytes and detected format
    /// through the document relationships exactly like a `<w:drawing>` image;
    /// its description is the `v:shape` `alt` text, its name the
    /// `v:imagedata` `o:title`, and its extents the `v:shape` `style`
    /// width/height (zero when the style cannot be read).
    #[inline]
    #[must_use]
    pub const fn preview_image(&self) -> Option<&InlineImage> {
        self.preview.as_ref()
    }

    /// Get the COM program ID (`o:OLEObject` `ProgID`), when declared.
    ///
    /// For example `Equation.DSMT4` for Equation Editor objects or `Package`
    /// for embedded packages.
    #[inline]
    #[must_use]
    pub fn prog_id(&self) -> Option<&str> {
        self.prog_id.as_deref()
    }

    /// Get the embedding type (`o:OLEObject` `Type`), when declared.
    ///
    /// Word writes `Embed` for embedded payloads and `Link` for linked ones.
    #[inline]
    #[must_use]
    pub fn object_type(&self) -> Option<&str> {
        self.object_type.as_deref()
    }

    /// Get the draw aspect (`o:OLEObject` `DrawAspect`), when declared.
    ///
    /// Word writes `Content` when the object renders its contents and `Icon`
    /// when it renders as an icon.
    #[inline]
    #[must_use]
    pub fn draw_aspect(&self) -> Option<&str> {
        self.draw_aspect.as_deref()
    }

    /// Get the relationship ID of the embedded payload part (`o:OLEObject`
    /// `r:id`), when declared.
    ///
    /// The relationship typically targets `word/embeddings/oleObjectN.bin`.
    #[inline]
    #[must_use]
    pub fn ole_r_id(&self) -> Option<&str> {
        self.ole_r_id.as_deref()
    }

    /// Load the embedded payload bytes from the package.
    ///
    /// Returns `Ok(None)` when the object declares no payload relationship.
    /// The bytes are borrowed from the package without copying and are never
    /// interpreted.
    ///
    /// # Arguments
    ///
    /// * `opc` - Reference to the OPC package
    /// * `rels` - Relationships from the document part
    ///
    /// # Errors
    ///
    /// Returns an error when the declared relationship or payload part is
    /// missing from the package.
    pub fn ole_data<'a>(
        &self,
        opc: &'a OpcPackage,
        rels: &Relationships,
    ) -> Result<Option<Cow<'a, [u8]>>> {
        let Some(r_id) = &self.ole_r_id else {
            return Ok(None);
        };

        // Resolve the relationship to get the payload part name
        let rel = rels.get(r_id).ok_or_else(|| {
            Error::InvalidFormat(format!("OLE object relationship not found: {r_id}"))
        })?;

        // Get the target part name
        let part_name = rel
            .target_partname()
            .map_err(|e| Error::InvalidFormat(format!("Invalid OLE object part reference: {e}")))?;

        // Get the part from the package
        let part = opc
            .get_part(&part_name)
            .map_err(|e| Error::InvalidFormat(format!("OLE object part not found: {e}")))?;

        // Return the payload as a borrowed slice (zero-copy)
        Ok(Some(Cow::Borrowed(part.blob())))
    }
}

/// Parse one `<w:object>` run child into a typed embedded object.
///
/// Returns `Ok(None)` when the fragment carries neither an `o:OLEObject`
/// descriptor nor a `v:imagedata` preview relationship, so the caller can
/// retain the element as opaque XML instead.
///
/// # Arguments
///
/// * `xml_bytes` - The raw XML bytes of one `<w:object>` element
///
/// # Errors
///
/// Returns an error for malformed XML.
pub(crate) fn parse_embedded_object(xml_bytes: &[u8]) -> Result<Option<EmbeddedObject>> {
    let mut reader = Reader::from_reader(xml_bytes);
    reader.config_mut().trim_text(true);

    // VML shape and preview state; the first element carrying each value wins.
    let mut shape_alt: Option<String> = None;
    let mut shape_style: Option<String> = None;
    let mut preview_r_id: Option<String> = None;
    let mut preview_title: Option<String> = None;

    // o:OLEObject descriptor state.
    let mut prog_id: Option<String> = None;
    let mut object_type: Option<String> = None;
    let mut draw_aspect: Option<String> = None;
    let mut ole_r_id: Option<String> = None;

    loop {
        match reader.read_event() {
            Ok(Event::Start(e) | Event::Empty(e)) => match e.local_name().as_ref() {
                b"shape" => {
                    for attr in first_wins(&e).flatten() {
                        let value = match std::str::from_utf8(&attr.value) {
                            Ok(value) => value,
                            Err(_) => continue,
                        };
                        match attr.key.as_ref() {
                            b"alt" if shape_alt.is_none() => {
                                shape_alt = Some(value.to_owned());
                            },
                            b"style" if shape_style.is_none() => {
                                shape_style = Some(value.to_owned());
                            },
                            _ => {},
                        }
                    }
                },
                b"imagedata" => {
                    for attr in first_wins(&e).flatten() {
                        let key = attr.key.as_ref();
                        let value = match std::str::from_utf8(&attr.value) {
                            Ok(value) => value,
                            Err(_) => continue,
                        };
                        // Check for r:id / r:embed (with namespace prefix) and o:title
                        if (key.ends_with(b":id") || key.ends_with(b":embed"))
                            && preview_r_id.is_none()
                        {
                            preview_r_id = Some(value.to_owned());
                        } else if key.ends_with(b":title") && preview_title.is_none() {
                            preview_title = Some(value.to_owned());
                        }
                    }
                },
                b"OLEObject" => {
                    for attr in first_wins(&e).flatten() {
                        let key = attr.key.as_ref();
                        let value = match std::str::from_utf8(&attr.value) {
                            Ok(value) => value,
                            Err(_) => continue,
                        };
                        match key {
                            b"Type" if object_type.is_none() => {
                                object_type = Some(value.to_owned());
                            },
                            b"ProgID" if prog_id.is_none() => {
                                prog_id = Some(value.to_owned());
                            },
                            b"DrawAspect" if draw_aspect.is_none() => {
                                draw_aspect = Some(value.to_owned());
                            },
                            // Check for r:id (with namespace prefix)
                            _ if key.ends_with(b":id") && ole_r_id.is_none() => {
                                ole_r_id = Some(value.to_owned());
                            },
                            _ => {},
                        }
                    }
                },
                _ => {},
            },
            Ok(Event::Eof) => break,
            Err(e) => return Err(Error::Xml(e.to_string())),
            _ => {},
        }
    }

    let has_descriptor =
        prog_id.is_some() || object_type.is_some() || draw_aspect.is_some() || ole_r_id.is_some();
    let preview = preview_r_id.map(|r_id| {
        let (width_emu, height_emu) = shape_style.as_deref().map_or((0, 0), shape_extent_emu);
        InlineImage::new(
            r_id,
            width_emu,
            height_emu,
            shape_alt.unwrap_or_default(),
            preview_title.unwrap_or_default(),
        )
    });
    if !has_descriptor && preview.is_none() {
        return Ok(None);
    }
    Ok(Some(EmbeddedObject {
        preview,
        prog_id,
        object_type,
        draw_aspect,
        ole_r_id,
    }))
}

/// Read the `width`/`height` declarations of a VML `style` attribute as EMU.
///
/// Missing or unreadable declarations yield a zero extent; the preview
/// relationship remains the authoritative payload either way.
fn shape_extent_emu(style: &str) -> (i64, i64) {
    let mut width_emu = 0;
    let mut height_emu = 0;
    for declaration in style.split(';') {
        let Some((name, value)) = declaration.split_once(':') else {
            continue;
        };
        let value = value.trim();
        let unit_start = value
            .find(|c: char| !(c.is_ascii_digit() || matches!(c, '.' | '-' | '+')))
            .unwrap_or(value.len());
        let Ok(number) = value[..unit_start].parse::<f64>() else {
            continue;
        };
        if !number.is_finite() {
            continue;
        }
        let scale = match value[unit_start..].trim() {
            "" | "pt" => EMUS_PER_PT,
            "in" => EMUS_PER_INCH,
            "cm" => EMUS_PER_CM,
            "mm" => EMUS_PER_MM,
            "px" => EMUS_PER_PX,
            _ => continue,
        };
        // Rust float-to-int casts saturate, so hostile extents cannot wrap.
        let emu = (number * scale as f64) as i64;
        match name.trim() {
            "width" => width_emu = emu,
            "height" => height_emu = emu,
            _ => {},
        }
    }
    (width_emu, height_emu)
}

#[cfg(test)]
mod tests {
    use super::*;

    const OBJECT: &str = r##"<w:object w:dxaOrig="1440" w:dyaOrig="1080">
        <v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75" filled="f" stroked="f">
            <v:stroke joinstyle="miter"/>
        </v:shapetype>
        <v:shape id="_x0000_i1025" type="#_x0000_t75" alt="equation one" style="width:16.7pt;height:14pt" o:ole="">
            <v:imagedata r:id="rId7" o:title="equation title"/>
            <o:lock v:ext="edit" aspectratio="t"/>
        </v:shape>
        <o:OLEObject Type="Embed" ProgID="Equation.DSMT4" ShapeID="_x0000_i1025" DrawAspect="Content" ObjectID="_1468075725" r:id="rId8"/>
    </w:object>"##;

    #[test]
    fn parses_descriptor_and_preview() {
        let object = parse_embedded_object(OBJECT.as_bytes())
            .unwrap()
            .expect("typed object");
        assert_eq!(object.prog_id(), Some("Equation.DSMT4"));
        assert_eq!(object.object_type(), Some("Embed"));
        assert_eq!(object.draw_aspect(), Some("Content"));
        assert_eq!(object.ole_r_id(), Some("rId8"));

        let preview = object.preview_image().expect("preview image");
        assert_eq!(preview.r_embed(), "rId7");
        assert_eq!(preview.description(), "equation one");
        assert_eq!(preview.name(), "equation title");
        assert_eq!(preview.width_emu(), 212_090); // 16.7pt
        assert_eq!(preview.height_emu(), 177_800); // 14pt
    }

    #[test]
    fn parses_object_without_preview() {
        let xml = r#"<w:object><v:shape id="_x0000_i1025" o:ole=""/><o:OLEObject Type="Embed" ProgID="Package" ShapeID="_x0000_i1025" DrawAspect="Icon" ObjectID="_1" r:id="rId9"/></w:object>"#;
        let object = parse_embedded_object(xml.as_bytes())
            .unwrap()
            .expect("typed object");
        assert_eq!(object.prog_id(), Some("Package"));
        assert_eq!(object.draw_aspect(), Some("Icon"));
        assert_eq!(object.ole_r_id(), Some("rId9"));
        assert!(object.preview_image().is_none());
    }

    #[test]
    fn parses_preview_only_object() {
        let xml = r#"<w:object><v:shape id="s" style="width:1in;height:96px"><v:imagedata r:id="rId3"/></v:shape></w:object>"#;
        let object = parse_embedded_object(xml.as_bytes())
            .unwrap()
            .expect("typed object");
        assert!(object.prog_id().is_none());
        let preview = object.preview_image().expect("preview image");
        assert_eq!(preview.r_embed(), "rId3");
        assert_eq!(preview.width_emu(), EMUS_PER_INCH);
        assert_eq!(preview.height_emu(), 96 * EMUS_PER_PX);
        assert!(preview.description().is_empty());
        assert!(preview.name().is_empty());
    }

    #[test]
    fn object_without_modeled_content_stays_opaque() {
        let xml = br#"<w:object><v:shape id="s"/></w:object>"#;
        assert!(parse_embedded_object(xml).unwrap().is_none());
    }

    #[test]
    fn rejects_malformed_xml() {
        let xml = br#"<w:object><v:shape"#;
        assert!(parse_embedded_object(xml).is_err());
    }

    #[test]
    fn ole_data_requires_a_declared_relationship() {
        let xml = br#"<w:object><v:shape id="s"><v:imagedata r:id="rId3"/></v:shape></w:object>"#;
        let object = parse_embedded_object(xml).unwrap().expect("typed object");
        let package = OpcPackage::new();
        let relationships = Relationships::new("/word/document.xml".to_owned());
        assert!(object.ole_data(&package, &relationships).unwrap().is_none());
    }
}
