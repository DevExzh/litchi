//! VML shape authoring for DOCX documents.
//!
//! Legacy VML shapes are serialized inside `w:pict` as `v:rect`,
//! `v:roundrect`, `v:oval`, or `v:line` elements with a CSS-like `style`
//! attribute (extents, floating position) and optional fill/stroke colors.
//! An optional `v:textbox` carries a plain word-processing story, which is
//! discoverable after save and reopen through
//! [`crate::Document::text_boxes`] (anchor
//! [`crate::textbox::TextBoxAnchor::Vml`]), matching the read model in
//! [`crate::textbox`].
//!
//! Everything is inert metadata: VML has no relationships, scripts, or
//! executable content.

use crate::error::{Error, Result};
use crate::textbox::{TextBoxParagraph, TextBoxRun};
use litchi_core::unit::{EMUS_PER_INCH, EMUS_PER_PT};
use litchi_core::xml::escape_xml;
use std::fmt::Write as FmtWrite;

/// Default shape width (1 inch) when no explicit size is set.
const DEFAULT_WIDTH_EMU: i64 = EMUS_PER_INCH;
/// Default shape height (1 inch) when no explicit size is set.
const DEFAULT_HEIGHT_EMU: i64 = EMUS_PER_INCH;

/// The VML preset shape element to emit.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
#[repr(u8)]
pub enum VmlShapeKind {
    /// `v:rect` — rectangle.
    #[default]
    Rectangle,
    /// `v:roundrect` — rounded rectangle.
    RoundRectangle,
    /// `v:oval` — ellipse/circle.
    Ellipse,
    /// `v:line` — straight line from the top-left to the bottom-right of the
    /// shape box.
    Line,
}

impl VmlShapeKind {
    /// The VML element name for this shape kind.
    fn element_name(self) -> &'static str {
        match self {
            Self::Rectangle => "v:rect",
            Self::RoundRectangle => "v:roundrect",
            Self::Ellipse => "v:oval",
            Self::Line => "v:line",
        }
    }
}

/// How a VML shape is positioned in the document flow.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub enum VmlShapePosition {
    /// The shape flows with the surrounding text (default).
    #[default]
    Inline,
    /// The shape floats at an absolute offset from the page margin, expressed
    /// in EMUs via the `mso-position-horizontal/vertical` style properties.
    Floating {
        /// Horizontal offset from the margin (`margin-left`).
        x_emu: i64,
        /// Vertical offset from the margin (`margin-top`).
        y_emu: i64,
    },
}

/// A mutable VML shape being authored in a document.
///
/// Built with [`MutableVmlShape::new`], then attached to a document via
/// [`crate::writer::MutableDocument::add_vml_shape`], which assigns the
/// shape identity and validates uniqueness.
///
/// # Examples
///
/// ```rust,no_run
/// use litchi_docx::{MutableVmlShape, Package, VmlShapeKind};
///
/// let mut package = Package::new()?;
/// let mut shape = MutableVmlShape::new(VmlShapeKind::RoundRectangle, 1828800, 914400)?;
/// shape.set_fill_color("#E5F1F8")?;
/// shape.add_run("legacy box").bold = Some(true);
/// package.document_mut()?.add_vml_shape(shape)?;
/// package.save("out.docx")?;
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
#[derive(Debug)]
pub struct MutableVmlShape {
    /// The VML preset shape element.
    pub(crate) kind: VmlShapeKind,
    /// VML shape ID (`v:shape@id`); assigned by
    /// `MutableDocument::add_vml_shape` when left empty.
    pub(crate) id: String,
    /// Shape width in EMUs (English Metric Units, 1 inch = 914400 EMUs).
    pub(crate) width_emu: i64,
    /// Shape height in EMUs.
    pub(crate) height_emu: i64,
    /// Inline or floating positioning.
    pub(crate) position: VmlShapePosition,
    /// Fill color as normalized `#RRGGBB`.
    pub(crate) fill_color: Option<String>,
    /// Stroke (outline) color as normalized `#RRGGBB`.
    pub(crate) stroke_color: Option<String>,
    /// The optional `v:textbox` story as paragraphs with runs.
    pub(crate) paragraphs: Vec<TextBoxParagraph>,
}

impl MutableVmlShape {
    /// Create a shape of the given kind with the given EMU extents.
    pub fn new(kind: VmlShapeKind, width_emu: i64, height_emu: i64) -> Result<Self> {
        if width_emu <= 0 || height_emu <= 0 {
            return Err(Error::InvalidFormat(
                "VML shape extents must be positive EMU values".to_string(),
            ));
        }
        Ok(Self {
            kind,
            id: String::new(),
            width_emu,
            height_emu,
            position: VmlShapePosition::Inline,
            fill_color: None,
            stroke_color: None,
            paragraphs: Vec::new(),
        })
    }

    /// Create a rectangle with default (one square inch) extents.
    pub fn rectangle() -> Self {
        Self::new(
            VmlShapeKind::Rectangle,
            DEFAULT_WIDTH_EMU,
            DEFAULT_HEIGHT_EMU,
        )
        .expect("default extents are positive")
    }

    /// Get the VML shape kind.
    pub fn kind(&self) -> VmlShapeKind {
        self.kind
    }

    /// Get the assigned VML shape ID (empty until the shape is added to a
    /// document).
    pub fn id(&self) -> &str {
        &self.id
    }

    /// Set an explicit VML shape ID instead of accepting the allocated one.
    ///
    /// The ID must be unique within the document;
    /// [`crate::writer::MutableDocument::add_vml_shape`] rejects
    /// collisions with existing shapes.
    pub fn set_shape_id(&mut self, id: impl Into<String>) -> Result<&mut Self> {
        self.id = id.into();
        super::ole_object::validate_shape_id(&self.id)?;
        Ok(self)
    }

    /// Get the positioning mode.
    pub fn position(&self) -> VmlShapePosition {
        self.position
    }

    /// Float the shape at an absolute offset from the page margin.
    pub fn set_floating(&mut self, x_emu: i64, y_emu: i64) -> &mut Self {
        self.position = VmlShapePosition::Floating { x_emu, y_emu };
        self
    }

    /// Set the fill color (`#RRGGBB`, with or without the `#` prefix).
    pub fn set_fill_color(&mut self, color: &str) -> Result<&mut Self> {
        self.fill_color = Some(normalize_color(color)?);
        Ok(self)
    }

    /// Set the stroke (outline) color (`#RRGGBB`, with or without the `#`
    /// prefix).
    pub fn set_stroke_color(&mut self, color: &str) -> Result<&mut Self> {
        self.stroke_color = Some(normalize_color(color)?);
        Ok(self)
    }

    /// Append a paragraph of plain text to the `v:textbox` story.
    pub fn add_paragraph_with_text(&mut self, text: &str) -> &mut Self {
        self.paragraphs.push(TextBoxParagraph::default());
        self.add_run(text);
        self
    }

    /// Append a run with the given text to the current story paragraph,
    /// starting a new paragraph when the story is empty or already closed.
    ///
    /// Returns the new run so basic formatting can be toggled through its
    /// public fields (`bold`, `italic`, `underline`).
    pub fn add_run(&mut self, text: &str) -> &mut TextBoxRun {
        if self.paragraphs.is_empty() {
            self.paragraphs.push(TextBoxParagraph::default());
        }
        let paragraph = self.paragraphs.last_mut().expect("paragraph checked above");
        paragraph.runs.push(TextBoxRun {
            text: text.to_string(),
            ..TextBoxRun::default()
        });
        paragraph.runs.last_mut().expect("run pushed above")
    }

    /// Get the story paragraphs.
    pub fn paragraphs(&self) -> &[TextBoxParagraph] {
        &self.paragraphs
    }

    /// Serialize the `<w:pict>` element.
    ///
    /// VML shapes carry no relationships, so serialization is identical in
    /// both writer modes. The enclosing `<w:r>` wrapper is emitted by the
    /// paragraph writer.
    pub(crate) fn to_xml(&self, xml: &mut String) -> Result<()> {
        if self.id.is_empty() {
            return Err(Error::InvalidFormat(
                "VML shape has no ID; add it via MutableDocument::add_vml_shape".to_string(),
            ));
        }
        let width_pt = self.width_emu as f64 / EMUS_PER_PT as f64;
        let height_pt = self.height_emu as f64 / EMUS_PER_PT as f64;
        let mut style = format!("width:{width_pt}pt;height:{height_pt}pt");
        if let VmlShapePosition::Floating { x_emu, y_emu } = self.position {
            let x_pt = x_emu as f64 / EMUS_PER_PT as f64;
            let y_pt = y_emu as f64 / EMUS_PER_PT as f64;
            write!(
                style,
                ";position:absolute;margin-left:{x_pt}pt;margin-top:{y_pt}pt;\
                 mso-position-horizontal:absolute;mso-position-horizontal-relative:margin;\
                 mso-position-vertical:absolute;mso-position-vertical-relative:margin"
            )
            .map_err(|error| Error::Xml(error.to_string()))?;
        }

        let element = self.kind.element_name();
        write!(
            xml,
            r#"<w:pict xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">"#
        )
        .map_err(|error| Error::Xml(error.to_string()))?;
        if self.kind == VmlShapeKind::Line {
            write!(
                xml,
                r#"<{element} id="{}" from="0,0" to="{width_pt},{height_pt}" style="{style}""#,
                escape_xml(&self.id),
            )
            .map_err(|error| Error::Xml(error.to_string()))?;
        } else {
            write!(
                xml,
                r#"<{element} id="{}" style="{style}""#,
                escape_xml(&self.id),
            )
            .map_err(|error| Error::Xml(error.to_string()))?;
        }
        if let Some(fill_color) = &self.fill_color {
            write!(xml, r#" fillcolor="{}""#, escape_xml(fill_color))
                .map_err(|error| Error::Xml(error.to_string()))?;
        }
        if let Some(stroke_color) = &self.stroke_color {
            write!(xml, r#" strokecolor="{}""#, escape_xml(stroke_color))
                .map_err(|error| Error::Xml(error.to_string()))?;
        }
        xml.push('>');
        if !self.paragraphs.is_empty() {
            xml.push_str("<v:textbox><w:txbxContent>");
            super::textbox::write_story_xml(xml, &self.paragraphs)?;
            xml.push_str("</w:txbxContent></v:textbox>");
        }
        write!(xml, "</{element}></w:pict>").map_err(|error| Error::Xml(error.to_string()))?;
        Ok(())
    }
}

/// Validate and normalize a color to `#RRGGBB`.
fn normalize_color(color: &str) -> Result<String> {
    let digits = color.strip_prefix('#').unwrap_or(color);
    if digits.len() != 6 || !digits.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        return Err(Error::InvalidFormat(format!(
            "invalid VML color '{color}'; expected RRGGBB hex digits"
        )));
    }
    Ok(format!("#{}", digits.to_ascii_uppercase()))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::textbox::{TextBoxAnchor, TextBoxBodyProperties, load_text_boxes};

    fn shape_document_xml(shape: &MutableVmlShape) -> String {
        let mut xml = String::new();
        shape.to_xml(&mut xml).unwrap();
        format!(
            "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" \
             xmlns:v=\"urn:schemas-microsoft-com:vml\" \
             xmlns:o=\"urn:schemas-microsoft-com:office:office\">\
             <w:body><w:p><w:r>{xml}</w:r></w:p></w:body></w:document>"
        )
    }

    fn identified(shape: &mut MutableVmlShape) {
        shape.id = "_x0000_s1025".to_string();
    }

    #[test]
    fn serializes_rect_with_colors_and_text() {
        let mut shape = MutableVmlShape::new(VmlShapeKind::Rectangle, 1828800, 914400).unwrap();
        identified(&mut shape);
        shape.set_fill_color("e5f1f8").unwrap();
        shape.set_stroke_color("#007ab9").unwrap();
        shape.add_run("legacy ").bold = Some(true);
        shape.add_run("box");

        let mut xml = String::new();
        shape.to_xml(&mut xml).unwrap();
        assert!(xml.contains(r#"<v:rect id="_x0000_s1025" style="width:144pt;height:72pt""#));
        assert!(xml.contains(r##"fillcolor="#E5F1F8""##));
        assert!(xml.contains(r##"strokecolor="#007AB9""##));
        assert!(xml.contains("<v:textbox><w:txbxContent>"));
        assert!(!xml.contains("position:absolute"));

        let inventory = load_text_boxes(shape_document_xml(&shape).as_bytes()).unwrap();
        assert_eq!(inventory.len(), 1);
        let parsed = &inventory[0];
        assert_eq!(parsed.anchor, TextBoxAnchor::Vml);
        assert_eq!(parsed.name.as_deref(), Some("_x0000_s1025"));
        assert_eq!(parsed.text(), "legacy box");
        assert_eq!(parsed.paragraphs[0].runs[0].bold, Some(true));
    }

    #[test]
    fn serializes_each_shape_kind() {
        for (kind, element) in [
            (VmlShapeKind::Rectangle, "<v:rect "),
            (VmlShapeKind::RoundRectangle, "<v:roundrect "),
            (VmlShapeKind::Ellipse, "<v:oval "),
        ] {
            let mut shape = MutableVmlShape::new(kind, 914400, 914400).unwrap();
            identified(&mut shape);
            let mut xml = String::new();
            shape.to_xml(&mut xml).unwrap();
            assert!(xml.contains(element), "missing {element} in {xml}");
            assert!(!xml.contains("<v:textbox>"));
        }

        let mut line = MutableVmlShape::new(VmlShapeKind::Line, 914400, 457200).unwrap();
        identified(&mut line);
        line.set_stroke_color("FF0000").unwrap();
        let mut xml = String::new();
        line.to_xml(&mut xml).unwrap();
        assert!(xml.contains(r#"<v:line id="_x0000_s1025" from="0,0" to="72,36""#));
        assert!(xml.contains(r##"strokecolor="#FF0000""##));
    }

    #[test]
    fn serializes_floating_position() {
        let mut shape = MutableVmlShape::new(VmlShapeKind::Rectangle, 914400, 914400).unwrap();
        identified(&mut shape);
        shape.set_floating(914400, 457200);
        let mut xml = String::new();
        shape.to_xml(&mut xml).unwrap();
        assert!(xml.contains("position:absolute"));
        assert!(xml.contains("margin-left:72pt"));
        assert!(xml.contains("margin-top:36pt"));
        assert!(xml.contains("mso-position-horizontal:absolute"));
        assert!(xml.contains("mso-position-vertical-relative:margin"));
    }

    #[test]
    fn validates_colors_extents_and_identity() {
        let mut shape = MutableVmlShape::new(VmlShapeKind::Rectangle, 914400, 914400).unwrap();
        assert!(shape.set_fill_color("red").is_err());
        assert!(shape.set_fill_color("#12345").is_err());
        assert!(shape.set_fill_color("#GGGGGG").is_err());
        assert!(shape.set_fill_color("A0B1C2").is_ok());
        assert!(MutableVmlShape::new(VmlShapeKind::Rectangle, 0, 914400).is_err());
        assert!(MutableVmlShape::new(VmlShapeKind::Rectangle, 914400, -1).is_err());
        // Serialization requires an assigned ID.
        assert!(shape.to_xml(&mut String::new()).is_err());
    }

    #[test]
    fn textless_shapes_are_not_text_box_inventory_entries() {
        let mut shape = MutableVmlShape::new(VmlShapeKind::Ellipse, 914400, 914400).unwrap();
        identified(&mut shape);
        shape.set_fill_color("#00FF00").unwrap();
        let inventory = load_text_boxes(shape_document_xml(&shape).as_bytes()).unwrap();
        assert!(inventory.is_empty());
    }

    #[test]
    fn round_trips_vml_shape_through_saved_package() {
        use crate::Package;
        use tempfile::NamedTempFile;

        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        {
            let document = package.document_mut().unwrap();
            document.add_paragraph_with_text("before the shape");
            let mut shape =
                MutableVmlShape::new(VmlShapeKind::RoundRectangle, 1828800, 914400).unwrap();
            shape.set_fill_color("#E5F1F8").unwrap();
            shape.set_stroke_color("#007AB9").unwrap();
            shape.set_floating(914400, 457200);
            shape.add_run("legacy ").bold = Some(true);
            shape.add_run("box");
            shape.add_paragraph_with_text("second line");
            let shape = document.add_vml_shape(shape).unwrap();
            assert_eq!(shape.id(), "_x0000_s1025");
        }
        package.save(file.path()).unwrap();

        let reopened = Package::open(file.path()).unwrap();
        let document = reopened.document().unwrap();
        assert!(document.text().unwrap().contains("before the shape"));

        let inventory = document.text_boxes().unwrap();
        assert_eq!(inventory.len(), 1);
        let parsed = &inventory[0];
        assert_eq!(parsed.anchor, TextBoxAnchor::Vml);
        assert_eq!(parsed.name.as_deref(), Some("_x0000_s1025"));
        assert_eq!(parsed.text(), "legacy box\nsecond line");
        assert_eq!(parsed.paragraphs[0].runs[0].bold, Some(true));
        // VML fallbacks carry no bodyPr; the typed defaults apply.
        assert_eq!(parsed.body, TextBoxBodyProperties::default());
        assert!(!parsed.is_word_art());
    }

    #[test]
    fn multiple_shapes_get_distinct_ids_and_reject_collisions() {
        use crate::Package;
        use tempfile::NamedTempFile;

        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        {
            let document = package.document_mut().unwrap();
            let mut first = MutableVmlShape::new(VmlShapeKind::Rectangle, 914400, 457200).unwrap();
            first.add_run("first");
            let first = document.add_vml_shape(first).unwrap();
            assert_eq!(first.id(), "_x0000_s1025");

            let mut second = MutableVmlShape::new(VmlShapeKind::Ellipse, 457200, 457200).unwrap();
            second.add_run("second");
            let second = document.add_vml_shape(second).unwrap();
            assert_eq!(second.id(), "_x0000_s1026");

            // An explicit shape ID colliding with an assigned one is rejected.
            let mut duplicate =
                MutableVmlShape::new(VmlShapeKind::Rectangle, 914400, 914400).unwrap();
            duplicate.set_shape_id("_x0000_s1025").unwrap();
            assert!(document.add_vml_shape(duplicate).is_err());
        }
        package.save(file.path()).unwrap();

        let reopened = Package::open(file.path()).unwrap();
        let inventory = reopened.document().unwrap().text_boxes().unwrap();
        assert_eq!(inventory.len(), 2);
        let mut entries: Vec<(String, String)> = inventory
            .iter()
            .map(|shape| (shape.name.as_deref().unwrap().to_owned(), shape.text()))
            .collect();
        entries.sort();
        assert_eq!(
            entries,
            [
                ("_x0000_s1025".to_owned(), "first".to_owned()),
                ("_x0000_s1026".to_owned(), "second".to_owned())
            ]
        );
    }

    #[test]
    fn coexists_with_drawingml_text_boxes_and_images() {
        use crate::{MutableTextBox, Package};
        use tempfile::NamedTempFile;

        const PNG_HEADER: &[u8] = &[0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A];

        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        {
            let document = package.document_mut().unwrap();

            let mut text_box = MutableTextBox::new("DrawingML Box", 914400, 457200).unwrap();
            text_box.add_run("drawingml story").unwrap();
            document.add_text_box(text_box);

            document
                .add_paragraph()
                .add_picture_from_bytes(PNG_HEADER.to_vec(), Some(914400), Some(914400))
                .unwrap();

            let mut shape = MutableVmlShape::new(VmlShapeKind::Rectangle, 914400, 457200).unwrap();
            shape.add_run("vml story");
            document.add_vml_shape(shape).unwrap();
        }
        package.save(file.path()).unwrap();

        let reopened = Package::open(file.path()).unwrap();
        let document = reopened.document().unwrap();
        let inventory = document.text_boxes().unwrap();
        assert_eq!(inventory.len(), 2);

        let drawingml = inventory
            .iter()
            .find(|shape| shape.name.as_deref() == Some("DrawingML Box"))
            .expect("DrawingML text box survives the round trip");
        assert_eq!(drawingml.anchor, TextBoxAnchor::Inline);
        assert_eq!(drawingml.text(), "drawingml story");

        let vml = inventory
            .iter()
            .find(|shape| shape.anchor == TextBoxAnchor::Vml)
            .expect("VML shape survives the round trip");
        assert_eq!(vml.text(), "vml story");

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
