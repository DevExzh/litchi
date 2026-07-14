//! OpenDocument Presentation builder.
//!
//! This module provides a builder pattern for creating new ODP presentations from scratch.

use crate::core::{OdfStructure, PackageWriter};
use crate::odp::Slide;
use litchi_core::{Metadata, Result, xml::escape_xml};
use std::path::Path;

/// Builder for creating new ODP presentations.
///
/// This builder allows you to create ODP presentations programmatically by adding
/// slides with text and shapes, then saving them to a file or bytes.
///
/// # Examples
///
/// ```no_run
/// use litchi_odf::PresentationBuilder;
///
/// # fn main() -> litchi_core::Result<()> {
/// let mut builder = PresentationBuilder::new();
/// builder.add_slide_with_title("Welcome", "This is my presentation")?;
/// builder.add_slide_with_title("Slide 2", "More content here")?;
/// builder.save("presentation.odp")?;
/// # Ok(())
/// # }
/// ```
pub struct PresentationBuilder {
    slides: Vec<Slide>,
    metadata: Metadata,
}

fn encode_text_content(text: &str) -> String {
    fn flush_plain(output: &mut String, plain: &mut String) {
        if !plain.is_empty() {
            output.push_str(&escape_xml(plain));
            plain.clear();
        }
    }

    let mut output = String::with_capacity(text.len());
    let mut plain = String::new();
    let mut characters = text.chars().peekable();
    while let Some(character) = characters.next() {
        match character {
            ' ' => {
                flush_plain(&mut output, &mut plain);
                let mut count = 1usize;
                while characters.next_if_eq(&' ').is_some() {
                    count += 1;
                }
                if count == 1 && !output.is_empty() && characters.peek().is_some() {
                    output.push(' ');
                } else if count == 1 {
                    output.push_str("<text:s/>");
                } else {
                    output.push_str(&format!(r#"<text:s text:c="{count}"/>"#));
                }
            },
            '\t' => {
                flush_plain(&mut output, &mut plain);
                output.push_str("<text:tab/>");
            },
            '\r' => {
                flush_plain(&mut output, &mut plain);
                output.push_str("<text:line-break/>");
            },
            _ => plain.push(character),
        }
    }
    flush_plain(&mut output, &mut plain);
    output
}

pub(super) fn generate_text_paragraphs(text: &str, style_name: Option<&str>) -> String {
    let escaped_style = style_name.map(escape_xml);
    let mut output = String::with_capacity(text.len() + 32);
    for paragraph in text.split('\n') {
        output.push_str("<text:p");
        if let Some(style) = escaped_style.as_deref() {
            output.push_str(r#" text:style-name="#);
            output.push('"');
            output.push_str(style);
            output.push('"');
        }
        output.push('>');
        output.push_str(&encode_text_content(paragraph));
        output.push_str("</text:p>");
    }
    output
}

fn push_optional_attribute(output: &mut String, name: &str, value: Option<&str>) {
    if let Some(value) = value {
        output.push(' ');
        output.push_str(name);
        output.push_str("=\"");
        output.push_str(&escape_xml(value));
        output.push('"');
    }
}

pub(super) fn slide_style_name(slide: &Slide, index: usize) -> String {
    if slide
        .transition
        .as_ref()
        .is_some_and(|value| !value.is_empty())
    {
        format!("dpTransition{}", index + 1)
    } else {
        "dp1".to_string()
    }
}

pub(super) fn generate_transition_styles(slides: &[Slide]) -> String {
    let mut output = String::from(
        r#"<style:style style:name="dp1" style:family="drawing-page"><style:drawing-page-properties/></style:style>"#,
    );
    for (index, slide) in slides.iter().enumerate() {
        let Some(transition) = slide.transition.as_ref().filter(|value| !value.is_empty()) else {
            continue;
        };
        output.push_str(r#"<style:style style:name=""#);
        output.push_str(&slide_style_name(slide, index));
        output.push_str(r#"" style:family="drawing-page"><style:drawing-page-properties"#);
        push_optional_attribute(
            &mut output,
            "presentation:transition-type",
            transition.transition_type.map(|value| value.as_str()),
        );
        push_optional_attribute(
            &mut output,
            "presentation:transition-style",
            transition.style.as_ref().map(|value| value.as_str()),
        );
        push_optional_attribute(
            &mut output,
            "presentation:transition-speed",
            transition.speed.map(|value| value.as_str()),
        );
        push_optional_attribute(&mut output, "smil:type", transition.smil_type.as_deref());
        push_optional_attribute(
            &mut output,
            "smil:subtype",
            transition.smil_subtype.as_deref(),
        );
        push_optional_attribute(
            &mut output,
            "smil:direction",
            transition.direction.map(|value| value.as_str()),
        );
        push_optional_attribute(
            &mut output,
            "smil:fadeColor",
            transition.fade_color.as_deref(),
        );
        push_optional_attribute(
            &mut output,
            "presentation:duration",
            transition.duration.as_deref(),
        );
        if let Some(sound) = transition.sound.as_ref() {
            output.push('>');
            output.push_str(r#"<presentation:sound xlink:type="simple" xlink:href=""#);
            output.push_str(&escape_xml(&sound.href));
            output.push('"');
            if sound.actuate_on_request {
                output.push_str(r#" xlink:actuate="onRequest""#);
            }
            push_optional_attribute(
                &mut output,
                "xlink:show",
                sound.show.map(|value| value.as_str()),
            );
            push_optional_attribute(&mut output, "xml:id", sound.xml_id.as_deref());
            push_optional_attribute(
                &mut output,
                "presentation:play-full",
                sound
                    .play_full
                    .map(|value| if value { "true" } else { "false" }),
            );
            output.push_str("/></style:drawing-page-properties>");
        } else {
            output.push_str("/>");
        }
        output.push_str("</style:style>");
    }
    output
}

impl Default for PresentationBuilder {
    fn default() -> Self {
        Self::new()
    }
}

impl PresentationBuilder {
    /// Create a new presentation builder
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi_odf::PresentationBuilder;
    ///
    /// let builder = PresentationBuilder::new();
    /// ```
    pub fn new() -> Self {
        Self {
            slides: Vec::new(),
            metadata: Metadata::default(),
        }
    }

    /// Set document metadata
    ///
    /// # Arguments
    ///
    /// * `metadata` - Document metadata (title, author, etc.)
    pub fn set_metadata(&mut self, metadata: Metadata) {
        self.metadata = metadata;
    }

    /// Add a slide with title and text content
    ///
    /// # Arguments
    ///
    /// * `title` - Title for the slide
    /// * `text` - Text content for the slide
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi_odf::PresentationBuilder;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut builder = PresentationBuilder::new();
    /// builder.add_slide_with_title("Introduction", "Welcome to our presentation")?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn add_slide_with_title(&mut self, title: &str, text: &str) -> Result<&mut Self> {
        let slide = Slide {
            title: Some(title.to_string()),
            text: text.to_string(),
            index: self.slides.len(),
            notes: None,
            transition: None,
            shapes: Vec::new(),
        };
        self.slides.push(slide);
        Ok(self)
    }

    /// Add a slide with only text content (no title)
    ///
    /// # Arguments
    ///
    /// * `text` - Text content for the slide
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi_odf::PresentationBuilder;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut builder = PresentationBuilder::new();
    /// builder.add_slide("Simple slide content")?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn add_slide(&mut self, text: &str) -> Result<&mut Self> {
        let slide = Slide {
            title: None,
            text: text.to_string(),
            index: self.slides.len(),
            notes: None,
            transition: None,
            shapes: Vec::new(),
        };
        self.slides.push(slide);
        Ok(self)
    }

    /// Add a Slide element directly
    ///
    /// # Arguments
    ///
    /// * `slide` - A complete `Slide` element to add
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi_odf::{PresentationBuilder, Slide, Shape};
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut builder = PresentationBuilder::new();
    /// let slide = Slide {
    ///     title: Some("Custom Slide".to_string()),
    ///     text: "Custom content".to_string(),
    ///     index: 0,
    ///     notes: Some("Speaker notes".to_string()),
    ///     transition: None,
    ///     shapes: vec![],
    /// };
    /// builder.add_slide_element(slide)?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn add_slide_element(&mut self, mut slide: Slide) -> Result<&mut Self> {
        slide.index = self.slides.len();
        self.slides.push(slide);
        Ok(self)
    }

    /// Generate XML for a shape
    pub(super) fn generate_shape_xml(shape: &crate::odp::Shape, idx: usize) -> Result<String> {
        use litchi_core::ShapeType;

        // Determine default position and size if not provided
        let x = shape.x.as_deref().unwrap_or("2cm");
        let y = shape.y.as_deref().unwrap_or("8cm");
        let width = shape.width.as_deref().unwrap_or("10cm");
        let height = shape.height.as_deref().unwrap_or("5cm");
        let default_name = format!("Shape{}", idx + 1);
        let name = shape.name.as_deref().unwrap_or(&default_name);
        let style_name = shape.style_name.as_deref().unwrap_or("gr3");
        let escaped_name = escape_xml(name);
        let escaped_style_name = escape_xml(style_name);
        let escaped_x = escape_xml(x);
        let escaped_y = escape_xml(y);
        let escaped_width = escape_xml(width);
        let escaped_height = escape_xml(height);

        let xml = match shape.shape_type {
            ShapeType::TextBox | ShapeType::Placeholder => {
                let presentation_class = if shape.shape_type == ShapeType::Placeholder {
                    r#" presentation:class="object""#
                } else {
                    ""
                };
                if shape.has_text() {
                    format!(
                        r#"<draw:frame draw:name="{}" draw:style-name="{}" draw:layer="layout"{} svg:x="{}" svg:y="{}" svg:width="{}" svg:height="{}"><draw:text-box>{}</draw:text-box></draw:frame>"#,
                        escaped_name,
                        escaped_style_name,
                        presentation_class,
                        escaped_x,
                        escaped_y,
                        escaped_width,
                        escaped_height,
                        generate_text_paragraphs(&shape.text, Some("P2"))
                    )
                } else {
                    // Empty frame
                    format!(
                        r#"<draw:frame draw:name="{}" draw:style-name="{}" draw:layer="layout"{} svg:x="{}" svg:y="{}" svg:width="{}" svg:height="{}"/>"#,
                        escaped_name,
                        escaped_style_name,
                        presentation_class,
                        escaped_x,
                        escaped_y,
                        escaped_width,
                        escaped_height
                    )
                }
            },
            ShapeType::AutoShape => {
                if shape.has_text() {
                    format!(
                        r#"<draw:rect draw:name="{}" draw:style-name="{}" draw:layer="layout" svg:x="{}" svg:y="{}" svg:width="{}" svg:height="{}">{}</draw:rect>"#,
                        escaped_name,
                        escaped_style_name,
                        escaped_x,
                        escaped_y,
                        escaped_width,
                        escaped_height,
                        generate_text_paragraphs(&shape.text, Some("P2"))
                    )
                } else {
                    format!(
                        r#"<draw:rect draw:name="{}" draw:style-name="{}" draw:layer="layout" svg:x="{}" svg:y="{}" svg:width="{}" svg:height="{}"/>"#,
                        escaped_name,
                        escaped_style_name,
                        escaped_x,
                        escaped_y,
                        escaped_width,
                        escaped_height
                    )
                }
            },
            ShapeType::Picture => {
                let image = shape.image_href.as_deref().map_or_else(
                    || "<draw:image/>".to_string(),
                    |href| {
                        format!(
                            r#"<draw:image xlink:href="{}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/>"#,
                            escape_xml(href)
                        )
                    },
                );
                format!(
                    r#"<draw:frame draw:name="{}" draw:style-name="{}" draw:layer="layout" svg:x="{}" svg:y="{}" svg:width="{}" svg:height="{}">{}</draw:frame>"#,
                    escaped_name,
                    escaped_style_name,
                    escaped_x,
                    escaped_y,
                    escaped_width,
                    escaped_height,
                    image
                )
            },
            ShapeType::Line | ShapeType::Connector => {
                // Line shape - use x,y as start and width,height as end offsets
                let x2 = shape.width.as_deref().unwrap_or("12cm");
                let y2 = shape.height.as_deref().unwrap_or("8cm");
                let element_name = if shape.shape_type == ShapeType::Connector {
                    "draw:connector"
                } else {
                    "draw:line"
                };
                format!(
                    r#"<{} draw:name="{}" draw:style-name="{}" draw:layer="layout" svg:x1="{}" svg:y1="{}" svg:x2="{}" svg:y2="{}"/>"#,
                    element_name,
                    escaped_name,
                    escaped_style_name,
                    escaped_x,
                    escaped_y,
                    escape_xml(x2),
                    escape_xml(y2)
                )
            },
            ShapeType::Group | ShapeType::Table | ShapeType::GraphicFrame | ShapeType::Unknown => {
                return Err(litchi_core::Error::InvalidFormat(format!(
                    "ODP serialization does not have enough data to write {} shape '{}'",
                    shape.shape_type, name
                )));
            },
            _ => {
                return Err(litchi_core::Error::InvalidFormat(format!(
                    "unsupported ODP shape type '{}' for shape '{}'",
                    shape.shape_type, name
                )));
            },
        };
        Ok(xml)
    }

    pub(super) fn generate_notes_xml(notes: Option<&str>) -> String {
        notes
            .filter(|notes| !notes.is_empty())
            .map(|notes| {
                format!(
                    r#"<presentation:notes><draw:frame draw:layer="layout" presentation:class="notes"><draw:text-box>{}</draw:text-box></draw:frame></presentation:notes>"#,
                    generate_text_paragraphs(notes, None)
                )
            })
            .unwrap_or_default()
    }

    /// Generate the content.xml body for presentation
    fn generate_content_body(&self) -> Result<String> {
        let shape_count = self.slides.iter().map(|s| s.shapes.len()).sum::<usize>();
        let mut estimated = 256usize;
        estimated += self.slides.len() * 128;
        estimated += shape_count * 192;
        estimated += self
            .slides
            .iter()
            .map(|s| s.text.len() + s.title.as_ref().map(|t| t.len()).unwrap_or(0))
            .sum::<usize>();
        estimated += self
            .slides
            .iter()
            .flat_map(|s| s.shapes.iter())
            .map(|sh| sh.text.len() + sh.name.as_ref().map(|n| n.len()).unwrap_or(0))
            .sum::<usize>();

        let mut body = String::with_capacity(estimated);

        for (i, slide) in self.slides.iter().enumerate() {
            let slide_style = slide_style_name(slide, i);
            body.push_str(&format!(
                r#"<draw:page draw:name="page{}" draw:style-name="{}" draw:master-page-name="Default">"#,
                i + 1,
                slide_style
            ));

            // Add title frame if title exists
            if let Some(ref title) = slide.title {
                body.push_str(&format!(
                    r#"<draw:frame draw:style-name="gr1" draw:text-style-name="P1" draw:layer="layout" presentation:class="title" svg:width="25.199cm" svg:height="3.506cm" svg:x="1.4cm" svg:y="0.962cm"><draw:text-box>{}</draw:text-box></draw:frame>"#,
                    generate_text_paragraphs(title, Some("P1"))
                ));
            }

            // Add text frame
            if !slide.text.is_empty() {
                let y_position = if slide.title.is_some() {
                    "5.0cm"
                } else {
                    "2.0cm"
                };
                body.push_str(&format!(
                    r#"<draw:frame draw:style-name="gr2" draw:text-style-name="P2" draw:layer="layout" presentation:class="object" svg:width="25.199cm" svg:height="10cm" svg:x="1.4cm" svg:y="{}"><draw:text-box>{}</draw:text-box></draw:frame>"#,
                    y_position,
                    generate_text_paragraphs(&slide.text, Some("P2"))
                ));
            }

            // Add custom shapes
            for (shape_idx, shape) in slide.shapes.iter().enumerate() {
                body.push_str(&Self::generate_shape_xml(shape, shape_idx)?);
            }

            body.push_str(&Self::generate_notes_xml(slide.notes.as_deref()));

            body.push_str("</draw:page>");
        }

        Ok(body)
    }

    /// Generate the complete content.xml for presentation
    fn generate_content_xml(&self) -> Result<String> {
        let body = self.generate_content_body()?;
        let transition_styles = generate_transition_styles(&self.slides);

        Ok(format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:smil="urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0" xmlns:math="http://www.w3.org/1998/Math/MathML" xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0" xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0" xmlns:ooo="http://openoffice.org/2004/office" office:version="1.3"><office:scripts/><office:font-face-decls/><office:automatic-styles>{}</office:automatic-styles><office:body><office:presentation>{}</office:presentation></office:body></office:document-content>"#,
            transition_styles, body
        ))
    }

    /// Generate meta.xml with metadata
    fn generate_meta_xml(&self) -> String {
        let now = chrono::Utc::now().to_rfc3339();

        let mut meta = format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" office:version="1.3"><office:meta><meta:generator>Litchi/0.0.1</meta:generator><meta:creation-date>{}</meta:creation-date><dc:date>{}</dc:date>"#,
            now, now
        );

        // Add optional metadata fields
        if let Some(ref title) = self.metadata.title {
            meta.push_str(&format!("<dc:title>{}</dc:title>", escape_xml(title)));
        }

        if let Some(ref author) = self.metadata.author {
            meta.push_str(&format!("<dc:creator>{}</dc:creator>", escape_xml(author)));
        }

        meta.push_str("</office:meta>");
        meta.push_str("</office:document-meta>");

        meta
    }

    /// Build the presentation and return as bytes
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odf::PresentationBuilder;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut builder = PresentationBuilder::new();
    /// builder.add_slide("Slide content")?;
    /// let bytes = builder.build()?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn build(self) -> Result<Vec<u8>> {
        let mut writer = PackageWriter::new();

        // Set MIME type
        writer.set_mimetype("application/vnd.oasis.opendocument.presentation")?;

        // Add content.xml
        let content_xml = self.generate_content_xml()?;
        writer.add_file("content.xml", content_xml.as_bytes())?;

        // Add styles.xml
        let styles_xml = OdfStructure::default_styles_xml();
        writer.add_file("styles.xml", styles_xml.as_bytes())?;

        // Add meta.xml
        let meta_xml = self.generate_meta_xml();
        writer.add_file("meta.xml", meta_xml.as_bytes())?;

        // Finish and return bytes
        writer.finish_to_bytes()
    }

    /// Build and save the presentation to a file
    ///
    /// # Arguments
    ///
    /// * `path` - Path where the ODP file should be saved
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odf::PresentationBuilder;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut builder = PresentationBuilder::new();
    /// builder.add_slide("Slide content")?;
    /// builder.save("output.odp")?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn save<P: AsRef<Path>>(self, path: P) -> Result<()> {
        let bytes = self.build()?;
        std::fs::write(path, bytes)?;
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::odp::{
        Presentation, Shape, TransitionDirection, TransitionSound, TransitionSoundShow,
        TransitionSpeed, TransitionStyle, TransitionType,
    };
    use litchi_core::ShapeType;

    #[test]
    fn writes_native_rectangles_and_connectors() {
        let mut rectangle = Shape::new();
        rectangle.name = Some("Box & label".to_string());
        rectangle.text = "Visible <text>".to_string();
        let rectangle_xml = PresentationBuilder::generate_shape_xml(&rectangle, 0).unwrap();
        assert!(rectangle_xml.starts_with("<draw:rect"));
        assert!(rectangle_xml.contains("Visible &lt;text&gt;"));

        let mut connector = Shape::new();
        connector.shape_type = ShapeType::Connector;
        let connector_xml = PresentationBuilder::generate_shape_xml(&connector, 1).unwrap();
        assert!(connector_xml.starts_with("<draw:connector"));
    }

    #[test]
    fn rejects_shapes_without_enough_serializable_data() {
        for shape_type in [
            ShapeType::Group,
            ShapeType::Table,
            ShapeType::GraphicFrame,
            ShapeType::Unknown,
        ] {
            let mut shape = Shape::new();
            shape.shape_type = shape_type;
            assert!(PresentationBuilder::generate_shape_xml(&shape, 0).is_err());
        }
    }

    #[test]
    fn writes_paragraphs_and_explicit_odf_whitespace() {
        assert_eq!(
            generate_text_paragraphs(" leading  text\ttab\nnext ", Some("P&1")),
            r#"<text:p text:style-name="P&amp;1"><text:s/>leading<text:s text:c="2"/>text<text:tab/>tab</text:p><text:p text:style-name="P&amp;1">next<text:s/></text:p>"#
        );
    }

    #[test]
    fn transition_configuration_round_trips_through_a_package() {
        let mut builder = PresentationBuilder::new();
        builder.add_slide("Transition slide").unwrap();
        let transition = builder.slides[0].transition_mut();
        transition
            .set_transition_type(Some(TransitionType::Automatic))
            .set_style(Some(TransitionStyle::new("fade-from-left").unwrap()))
            .set_speed(Some(TransitionSpeed::Fast))
            .set_smil_type(Some("fade & dissolve"))
            .set_smil_subtype(Some("crossfade"))
            .set_direction(Some(TransitionDirection::Reverse));
        transition.set_fade_color(Some("#102030")).unwrap();
        transition.set_duration(Some("PT6.5S")).unwrap();
        let mut sound = TransitionSound::new("Sounds/a&b.ogg");
        sound.play_full = Some(true);
        sound.actuate_on_request = true;
        sound.show = Some(TransitionSoundShow::Replace);
        sound.xml_id = Some("transitionSound1".to_string());
        transition.set_sound(Some(sound));

        let presentation = Presentation::from_bytes(builder.build().unwrap()).unwrap();
        let slide = presentation.slides().unwrap().remove(0);
        let transition = slide.transition().unwrap();
        assert_eq!(
            transition.transition_type(),
            Some(TransitionType::Automatic)
        );
        assert_eq!(transition.style().unwrap().as_str(), "fade-from-left");
        assert_eq!(transition.speed(), Some(TransitionSpeed::Fast));
        assert_eq!(transition.smil_type(), Some("fade & dissolve"));
        assert_eq!(transition.smil_subtype(), Some("crossfade"));
        assert_eq!(transition.direction(), Some(TransitionDirection::Reverse));
        assert_eq!(transition.fade_color(), Some("#102030"));
        assert_eq!(transition.duration(), Some("PT6.5S"));
        let sound = transition.sound().unwrap();
        assert_eq!(sound.href, "Sounds/a&b.ogg");
        assert_eq!(sound.play_full, Some(true));
        assert!(sound.actuate_on_request);
        assert_eq!(sound.show, Some(TransitionSoundShow::Replace));
        assert_eq!(sound.xml_id.as_deref(), Some("transitionSound1"));
    }
}
