//! Text-box authoring for DOCX documents.
//!
//! A text box is serialized as a `<w:drawing>` inline wordprocessing shape
//! (`wps:wsp`) carrying a `wps:txbx` story and a `wps:bodyPr` body-properties
//! element, mirroring the read model in [`crate::docx::textbox`]. Text boxes
//! authored here reappear in the [`crate::docx::Document::text_boxes`]
//! inventory with identical semantics after save and reopen.
//!
//! WordArt authoring (text warps and run-level text fill/outline/effect
//! styling) is deliberately not supported; those remain read-only metadata.

use crate::docx::drawing::ShapeType;
use crate::docx::textbox::{
    TextBoxAutofit, TextBoxBodyProperties, TextBoxParagraph, TextBoxRun, TextDirection,
    TextVerticalAnchor, TextWrap,
};
use crate::error::{OoxmlError, Result};
use litchi_core::unit::EMUS_PER_INCH;
use litchi_core::xml::escape_xml;
use std::fmt::Write as FmtWrite;

/// Maximum paragraphs in an authored text-box story.
pub(crate) const MAX_PARAGRAPHS: usize = 1024;
/// Maximum aggregate story text bytes in an authored text box.
pub(crate) const MAX_TEXT_BYTES: usize = 1024 * 1024;

/// A mutable text box being authored in a document.
///
/// Built with [`MutableTextBox::new`] or [`MutableTextBox::with_shape`], then
/// attached to a paragraph via
/// [`crate::docx::writer::MutableParagraph::add_text_box`].
///
/// # Examples
///
/// ```rust,no_run
/// use litchi_ooxml::docx::{MutableTextBox, Package, TextBoxAutofit};
///
/// let mut package = Package::new()?;
/// let mut text_box = MutableTextBox::new("Greeting", 1828800, 914400)?;
/// text_box.add_run("Hello ").bold = Some(true);
/// text_box.add_run("world");
/// text_box.body_properties_mut().autofit = TextBoxAutofit::ShapeAutofit;
/// package.document_mut()?.add_text_box(text_box);
/// package.save("out.docx")?;
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
#[derive(Debug)]
pub struct MutableTextBox {
    /// Drawing element ID written to `wp:docPr@id`.
    pub(crate) id: u32,
    /// Shape name written to `wp:docPr@name`.
    pub(crate) name: String,
    /// Preset geometry of the shape.
    pub(crate) shape_type: ShapeType,
    /// Shape width in EMUs (English Metric Units, 1 inch = 914400 EMUs).
    pub(crate) width_emu: i64,
    /// Shape height in EMUs.
    pub(crate) height_emu: i64,
    /// Text-body properties (`wps:bodyPr`).
    pub(crate) body: TextBoxBodyProperties,
    /// The text-box story as paragraphs with runs.
    pub(crate) paragraphs: Vec<TextBoxParagraph>,
}

impl MutableTextBox {
    /// Create a rectangular text box with the given name and EMU extents.
    ///
    /// # Arguments
    ///
    /// * `name` - Shape name (written to `wp:docPr@name`)
    /// * `width_emu` - Width in EMUs (must be positive)
    /// * `height_emu` - Height in EMUs (must be positive)
    pub fn new(name: impl Into<String>, width_emu: i64, height_emu: i64) -> Result<Self> {
        Self::with_shape(name, ShapeType::Rectangle, width_emu, height_emu)
    }

    /// Create a text box with a preset shape geometry, name, and EMU extents.
    pub fn with_shape(
        name: impl Into<String>,
        shape_type: ShapeType,
        width_emu: i64,
        height_emu: i64,
    ) -> Result<Self> {
        if width_emu <= 0 || height_emu <= 0 {
            return Err(OoxmlError::InvalidFormat(
                "text box extents must be positive EMU values".to_string(),
            ));
        }
        Ok(Self {
            // Matches the writer convention for inline pictures (`wp:docPr
            // id="1"`); callers composing several shapes should assign
            // distinct IDs with `set_id`.
            id: 1,
            name: name.into(),
            shape_type,
            width_emu,
            height_emu,
            body: TextBoxBodyProperties::default(),
            paragraphs: Vec::new(),
        })
    }

    /// Set the drawing element ID (`wp:docPr@id`).
    pub fn set_id(&mut self, id: u32) -> &mut Self {
        self.id = id;
        self
    }

    /// Set the shape name (`wp:docPr@name`).
    pub fn set_name(&mut self, name: impl Into<String>) -> &mut Self {
        self.name = name.into();
        self
    }

    /// Get the shape name.
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Get the preset shape geometry.
    pub fn shape_type(&self) -> &ShapeType {
        &self.shape_type
    }

    /// Get the text-body properties.
    pub fn body_properties(&self) -> &TextBoxBodyProperties {
        &self.body
    }

    /// Get mutable access to the text-body properties (insets, vertical
    /// anchor, direction, wrap, autofit, columns).
    pub fn body_properties_mut(&mut self) -> &mut TextBoxBodyProperties {
        &mut self.body
    }

    /// Append a paragraph of plain text to the story.
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

    /// Convert extents from inches to EMUs.
    pub fn inches_to_emu(inches: f64) -> i64 {
        (inches * EMUS_PER_INCH as f64) as i64
    }

    /// Serialize the text box as a `<w:drawing>` inline wordprocessing shape.
    ///
    /// The enclosing `<w:r>` wrapper is emitted by the paragraph writer,
    /// matching the inline-picture serialization path. The bounded story
    /// limits are validated by [`write_story_xml`].
    pub(crate) fn to_xml(&self, xml: &mut String) -> Result<()> {
        let name = escape_xml(&self.name);
        write!(
            xml,
            r#"<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="{}" cy="{}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="{}" name="{}"/><wp:cNvGraphicFramePr/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{}" cy="{}"/></a:xfrm><a:prstGeom prst="{}"><a:avLst/></a:prstGeom></wps:spPr>"#,
            self.width_emu,
            self.height_emu,
            self.id,
            name,
            self.width_emu,
            self.height_emu,
            escape_xml(self.shape_type.to_preset()),
        )
        .map_err(|error| OoxmlError::Xml(error.to_string()))?;

        xml.push_str("<wps:txbx><w:txbxContent>");
        write_story_xml(xml, &self.paragraphs)?;
        xml.push_str("</w:txbxContent></wps:txbx>");

        write_body_properties(xml, &self.body)?;
        xml.push_str("</wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing>");
        Ok(())
    }
}

/// Serialize story paragraphs (`w:p` with runs and basic run properties),
/// validating the bounded story limits. Shared by the DrawingML text-box and
/// VML shape writers.
pub(crate) fn write_story_xml(xml: &mut String, paragraphs: &[TextBoxParagraph]) -> Result<()> {
    if paragraphs.len() > MAX_PARAGRAPHS {
        return Err(OoxmlError::InvalidFormat(
            "text box paragraph limit exceeded".to_string(),
        ));
    }
    let text_bytes: usize = paragraphs
        .iter()
        .flat_map(|paragraph| paragraph.runs.iter())
        .map(|run| run.text.len())
        .sum();
    if text_bytes > MAX_TEXT_BYTES {
        return Err(OoxmlError::InvalidFormat(
            "text box story text limit exceeded".to_string(),
        ));
    }
    for paragraph in paragraphs {
        xml.push_str("<w:p>");
        for run in &paragraph.runs {
            xml.push_str("<w:r>");
            write_run_properties(xml, run)?;
            write!(
                xml,
                r#"<w:t xml:space="preserve">{}</w:t>"#,
                escape_xml(&run.text)
            )
            .map_err(|error| OoxmlError::Xml(error.to_string()))?;
            xml.push_str("</w:r>");
        }
        xml.push_str("</w:p>");
    }
    Ok(())
}

/// Write the run properties of a story run (`w:b`, `w:i`, `w:u`).
fn write_run_properties(xml: &mut String, run: &TextBoxRun) -> Result<()> {
    if run.bold.is_none() && run.italic.is_none() && run.underline.is_none() {
        return Ok(());
    }
    xml.push_str("<w:rPr>");
    if let Some(bold) = run.bold {
        xml.push_str(if bold {
            "<w:b/>"
        } else {
            r#"<w:b w:val="0"/>"#
        });
    }
    if let Some(italic) = run.italic {
        xml.push_str(if italic {
            "<w:i/>"
        } else {
            r#"<w:i w:val="0"/>"#
        });
    }
    if let Some(underline) = run.underline {
        xml.push_str(if underline {
            r#"<w:u w:val="single"/>"#
        } else {
            r#"<w:u w:val="none"/>"#
        });
    }
    xml.push_str("</w:rPr>");
    Ok(())
}

/// Write the `wps:bodyPr` element; non-default values only.
fn write_body_properties(xml: &mut String, body: &TextBoxBodyProperties) -> Result<()> {
    write!(
        xml,
        r#"<wps:bodyPr lIns="{}" tIns="{}" rIns="{}" bIns="{}""#,
        body.insets.left_emu, body.insets.top_emu, body.insets.right_emu, body.insets.bottom_emu,
    )
    .map_err(|error| OoxmlError::Xml(error.to_string()))?;
    if body.vertical_anchor != TextVerticalAnchor::Top {
        write!(xml, r#" anchor="{}""#, body.vertical_anchor.as_token())
            .map_err(|error| OoxmlError::Xml(error.to_string()))?;
    }
    if body.anchor_center {
        xml.push_str(r#" anchorCtr="1""#);
    }
    if body.direction != TextDirection::Horizontal {
        write!(xml, r#" vert="{}""#, body.direction.as_token())
            .map_err(|error| OoxmlError::Xml(error.to_string()))?;
    }
    if body.wrap != TextWrap::Square {
        write!(xml, r#" wrap="{}""#, body.wrap.as_token())
            .map_err(|error| OoxmlError::Xml(error.to_string()))?;
    }
    if body.column_count > 1 {
        write!(xml, r#" numCol="{}""#, body.column_count)
            .map_err(|error| OoxmlError::Xml(error.to_string()))?;
    }
    if body.space_first_last_paragraph {
        xml.push_str(r#" spcFirstLastPara="1""#);
    }
    let autofit = match body.autofit {
        TextBoxAutofit::NoAutofit => "<a:noAutofit/>",
        TextBoxAutofit::ShapeAutofit => "<a:spAutoFit/>",
        TextBoxAutofit::NormalAutofit => "<a:normAutofit/>",
    };
    write!(xml, ">{autofit}</wps:bodyPr>").map_err(|error| OoxmlError::Xml(error.to_string()))?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::docx::textbox::{TextBoxAnchor, load_text_boxes};

    fn authored_xml(text_box: &MutableTextBox) -> String {
        let mut xml = String::new();
        text_box.to_xml(&mut xml).unwrap();
        format!(
            "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" \
             xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" \
             xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" \
             xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\">\
             <w:body><w:p><w:r>{xml}</w:r></w:p></w:body></w:document>"
        )
    }

    #[test]
    fn serializes_inline_text_box_with_body_properties() {
        let mut text_box = MutableTextBox::new("My Box", 1828800, 914400).unwrap();
        text_box.set_id(42);
        text_box.add_run("rich text");
        {
            let body = text_box.body_properties_mut();
            body.insets.left_emu = 182880;
            body.vertical_anchor = TextVerticalAnchor::Center;
            body.direction = TextDirection::Vertical270;
            body.wrap = TextWrap::None;
            body.autofit = TextBoxAutofit::ShapeAutofit;
            body.column_count = 2;
            body.space_first_last_paragraph = true;
        }
        let document_xml = authored_xml(&text_box);
        assert!(document_xml.contains("<wps:txbx><w:txbxContent>"));
        assert!(document_xml.contains(r#"<wps:bodyPr lIns="182880""#));
        assert!(document_xml.contains(r#"anchor="ctr""#));
        assert!(document_xml.contains(r#"vert="vert270""#));
        assert!(document_xml.contains(r#"<a:spAutoFit/>"#));

        let inventory = load_text_boxes(document_xml.as_bytes()).unwrap();
        assert_eq!(inventory.len(), 1);
        let parsed = &inventory[0];
        assert_eq!(parsed.id, Some(42));
        assert_eq!(parsed.name.as_deref(), Some("My Box"));
        assert_eq!(parsed.anchor, TextBoxAnchor::Inline);
        assert_eq!(parsed.shape_type, Some(ShapeType::Rectangle));
        assert_eq!(parsed.text(), "rich text");
        assert_eq!(parsed.body, *text_box.body_properties());
    }

    #[test]
    fn serializes_preset_shape_and_formatted_runs() {
        let mut text_box =
            MutableTextBox::with_shape("Round", ShapeType::RoundRectangle, 914400, 457200)
                .unwrap();
        text_box.add_run("bold").bold = Some(true);
        text_box.add_run("italic-off").italic = Some(false);
        text_box.add_run("under").underline = Some(true);
        text_box.add_paragraph_with_text("second");

        let document_xml = authored_xml(&text_box);
        assert!(document_xml.contains(r#"<a:prstGeom prst="roundRect">"#));
        assert!(document_xml.contains("<w:b/>"));
        assert!(document_xml.contains(r#"<w:i w:val="0"/>"#));
        assert!(document_xml.contains(r#"<w:u w:val="single"/>"#));

        let inventory = load_text_boxes(document_xml.as_bytes()).unwrap();
        let parsed = &inventory[0];
        assert_eq!(parsed.shape_type, Some(ShapeType::RoundRectangle));
        assert_eq!(parsed.text(), "bolditalic-offunder\nsecond");
        let runs = &parsed.paragraphs[0].runs;
        assert_eq!(runs[0].bold, Some(true));
        assert_eq!(runs[1].italic, Some(false));
        assert_eq!(runs[2].underline, Some(true));
    }

    #[test]
    fn defaults_match_read_model_defaults() {
        let text_box = MutableTextBox::new("Plain", 914400, 914400).unwrap();
        text_box.to_xml(&mut String::new()).unwrap();
        let inventory = load_text_boxes(authored_xml(&text_box).as_bytes()).unwrap();
        let parsed = &inventory[0];
        assert_eq!(parsed.body, TextBoxBodyProperties::default());
        assert!(!parsed.is_word_art());
        assert_eq!(parsed.text(), "");
    }

    #[test]
    fn rejects_invalid_extents() {
        assert!(MutableTextBox::new("Bad", 0, 914400).is_err());
        assert!(MutableTextBox::new("Bad", 914400, -1).is_err());
    }

    #[test]
    fn escapes_markup_in_text_and_name() {
        let mut text_box = MutableTextBox::new("A & <B>", 914400, 914400).unwrap();
        text_box.add_run("x < y & \"z\"");
        let document_xml = authored_xml(&text_box);
        assert!(document_xml.contains("name=\"A &amp; &lt;B&gt;\""));
        let inventory = load_text_boxes(document_xml.as_bytes()).unwrap();
        assert_eq!(inventory[0].name.as_deref(), Some("A & <B>"));
        assert_eq!(inventory[0].text(), "x < y & \"z\"");
    }

    #[test]
    fn round_trips_text_box_through_saved_package() {
        use crate::docx::Package;
        use tempfile::NamedTempFile;

        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        {
            let document = package.document_mut().unwrap();
            document.add_paragraph_with_text("before the box");
            let mut text_box = MutableTextBox::new("Round Trip Box", 1828800, 914400).unwrap();
            text_box.set_id(11);
            text_box.add_run("Hello ").bold = Some(true);
            text_box.add_run("world").italic = Some(true);
            text_box.add_paragraph_with_text("second paragraph");
            {
                let body = text_box.body_properties_mut();
                body.insets.left_emu = 182880;
                body.insets.right_emu = 182880;
                body.vertical_anchor = TextVerticalAnchor::Center;
                body.autofit = TextBoxAutofit::ShapeAutofit;
            }
            let expected_body = *text_box.body_properties();
            document.add_text_box(text_box);
            document.add_paragraph_with_text("after the box");

            package.save(file.path()).unwrap();

            let reopened = Package::open(file.path()).unwrap();
            let document = reopened.document().unwrap();
            // Ordinary body text extraction is unaffected by the shape.
            let text = document.text().unwrap();
            assert!(text.contains("before the box"));
            assert!(text.contains("after the box"));

            let inventory = document.text_boxes().unwrap();
            assert_eq!(inventory.len(), 1);
            let parsed = &inventory[0];
            assert_eq!(parsed.id, Some(11));
            assert_eq!(parsed.name.as_deref(), Some("Round Trip Box"));
            assert_eq!(parsed.anchor, TextBoxAnchor::Inline);
            assert_eq!(parsed.shape_type, Some(ShapeType::Rectangle));
            assert_eq!(parsed.text(), "Hello world\nsecond paragraph");
            assert_eq!(parsed.body, expected_body);
            assert!(!parsed.is_word_art());
            let runs = &parsed.paragraphs[0].runs;
            assert_eq!(runs[0].bold, Some(true));
            assert_eq!(runs[1].italic, Some(true));
        }
    }

    #[test]
    fn round_trips_multiple_text_boxes_and_coexists_with_images() {
        use crate::docx::Package;
        use tempfile::NamedTempFile;

        // Minimal PNG header accepted by the image writer.
        const PNG_HEADER: &[u8] = &[0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A];

        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        {
            let document = package.document_mut().unwrap();
            let mut first = MutableTextBox::new("First Box", 914400, 457200).unwrap();
            first.set_id(21);
            first.add_run("first story");
            document.add_text_box(first);

            document
                .add_paragraph()
                .add_picture_from_bytes(PNG_HEADER.to_vec(), Some(914400), Some(914400))
                .unwrap();

            let mut second =
                MutableTextBox::with_shape("Second Box", ShapeType::Ellipse, 457200, 457200)
                    .unwrap();
            second.set_id(22);
            second.body_properties_mut().wrap = TextWrap::None;
            second.add_paragraph_with_text("second story");
            document.add_text_box(second);

            package.save(file.path()).unwrap();

            let reopened = Package::open(file.path()).unwrap();
            let inventory = reopened.document().unwrap().text_boxes().unwrap();
            assert_eq!(inventory.len(), 2);

            let first = inventory
                .iter()
                .find(|text_box| text_box.name.as_deref() == Some("First Box"))
                .expect("first text box survives the round trip");
            assert_eq!(first.id, Some(21));
            assert_eq!(first.text(), "first story");

            let second = inventory
                .iter()
                .find(|text_box| text_box.name.as_deref() == Some("Second Box"))
                .expect("second text box survives the round trip");
            assert_eq!(second.id, Some(22));
            assert_eq!(second.shape_type, Some(ShapeType::Ellipse));
            assert_eq!(second.body.wrap, TextWrap::None);
            assert_eq!(second.text(), "second story");

            // The inline image is preserved alongside the text boxes.
            let document = reopened.document().unwrap();
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
}
