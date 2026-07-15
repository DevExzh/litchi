//! Watermark support for DOCX documents.
//!
//! Based on Apache POI's XWPFHeaderFooterPolicy watermark implementation.
use crate::error::{OoxmlError, Result};
use litchi_core::xml::escape_xml;
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;

const VML_NAMESPACE: &[u8] = b"urn:schemas-microsoft-com:vml";

/// A watermark for a Word document.
///
/// Watermarks are typically displayed as diagonal text behind the document content.
///
/// # Examples
///
/// ```rust,ignore
/// use litchi_ooxml::docx::writer::Watermark;
///
/// let watermark = Watermark::text("CONFIDENTIAL");
/// ```
#[derive(Debug, Clone, PartialEq)]
pub struct Watermark {
    /// Watermark text
    text: String,
    /// Font family (default: Cambria)
    font: String,
    /// Font size in points (default: 1pt for shape-based watermarks)
    font_size: f64,
    /// Color (RGB hex, default: "000000" black)
    color: String,
}

impl Watermark {
    /// Create a text watermark.
    ///
    /// # Arguments
    ///
    /// * `text` - Text to display as watermark
    ///
    /// # Examples
    ///
    /// ```rust,ignore
    /// let wm = Watermark::text("CONFIDENTIAL");
    /// ```
    pub fn text(text: impl Into<String>) -> Self {
        Self {
            text: text.into(),
            font: "Cambria".to_string(),
            font_size: 1.0,
            color: "black".to_string(),
        }
    }

    /// Set the font family.
    pub fn set_font(&mut self, font: impl Into<String>) {
        self.font = font.into();
    }

    /// Set the font size in points.
    pub fn set_font_size(&mut self, size: u32) {
        self.font_size = f64::from(size);
    }

    /// Set a potentially fractional VML text-path font size in points.
    pub fn set_font_size_points(&mut self, size: f64) -> Result<()> {
        if !size.is_finite() || size < 0.0 {
            return Err(OoxmlError::InvalidFormat(
                "watermark font size must be finite and non-negative".into(),
            ));
        }
        self.font_size = size;
        Ok(())
    }

    /// Set the color.
    pub fn set_color(&mut self, color: impl Into<String>) {
        self.color = color.into();
    }

    /// Get the watermark text.
    #[inline]
    pub fn get_text(&self) -> &str {
        &self.text
    }

    /// Get the watermark font family.
    #[inline]
    pub fn font(&self) -> &str {
        &self.font
    }

    /// Get the VML text-path font size in points.
    #[inline]
    pub fn font_size(&self) -> f64 {
        self.font_size
    }

    /// Get the VML fill color.
    #[inline]
    pub fn color(&self) -> &str {
        &self.color
    }

    /// Parse text watermarks from a Word header part.
    pub(crate) fn from_header_xml(xml: &[u8]) -> Result<Vec<Self>> {
        let mut reader = NsReader::from_reader(xml);
        reader.config_mut().check_end_names = true;
        let mut depth = 0usize;
        let mut shapes: Vec<(usize, Option<WatermarkCandidate>)> = Vec::new();
        let mut watermarks = Vec::new();

        loop {
            let decoder = reader.decoder();
            let event = reader
                .read_event()
                .map_err(|error| OoxmlError::Xml(error.to_string()))?
                .into_owned();
            let resolver = reader.resolver().clone();
            let (namespace, event) = resolver.resolve_event(event);

            match event {
                Event::Start(element) => {
                    depth = depth.checked_add(1).ok_or_else(|| {
                        OoxmlError::InvalidFormat("Word header XML nesting is too deep".into())
                    })?;
                    if depth > 256 {
                        return Err(OoxmlError::InvalidFormat(
                            "Word header XML nesting exceeds 256 elements".into(),
                        ));
                    }
                    if is_vml(&namespace) && element.local_name().as_ref() == b"shape" {
                        shapes.push((
                            depth,
                            WatermarkCandidate::from_shape(&element, decoder, &resolver)?,
                        ));
                    } else if is_vml(&namespace) && element.local_name().as_ref() == b"textpath" {
                        apply_text_path(&mut shapes, &element, decoder, &resolver)?;
                    }
                },
                Event::Empty(element)
                    if is_vml(&namespace) && element.local_name().as_ref() == b"textpath" =>
                {
                    apply_text_path(&mut shapes, &element, decoder, &resolver)?;
                },
                Event::End(element)
                    if is_vml(&namespace) && element.local_name().as_ref() == b"shape" =>
                {
                    let Some((shape_depth, candidate)) = shapes.pop() else {
                        return Err(OoxmlError::InvalidFormat(
                            "VML shape end has no matching start".into(),
                        ));
                    };
                    if shape_depth != depth {
                        return Err(OoxmlError::InvalidFormat(
                            "invalid nested VML shape structure".into(),
                        ));
                    }
                    if let Some(watermark) = candidate.and_then(WatermarkCandidate::finish) {
                        watermarks.push(watermark);
                    }
                    depth = depth.checked_sub(1).ok_or_else(|| {
                        OoxmlError::InvalidFormat("invalid Word header XML nesting".into())
                    })?;
                },
                Event::End(_) => {
                    depth = depth.checked_sub(1).ok_or_else(|| {
                        OoxmlError::InvalidFormat("invalid Word header XML nesting".into())
                    })?;
                },
                Event::Eof if depth != 0 || !shapes.is_empty() => {
                    return Err(OoxmlError::InvalidFormat(
                        "unterminated Word header XML".into(),
                    ));
                },
                Event::Eof => break,
                _ => {},
            }
        }

        Ok(watermarks)
    }

    /// Generate watermark XML for a header paragraph.
    ///
    /// This creates a VML shape-based watermark that appears behind document content.
    /// The watermark is embedded in a header paragraph.
    ///
    /// # Arguments
    ///
    /// * `idx` - Index for unique ID generation (1, 2, or 3 for default/first/even)
    pub(crate) fn to_header_xml(&self, idx: u32) -> Result<String> {
        let mut xml = String::with_capacity(2048);

        // Start paragraph with Header style
        xml.push_str(r#"<w:p><w:pPr><w:pStyle w:val="Header"/></w:pPr>"#);

        // Start run with picture
        xml.push_str(r#"<w:r><w:rPr><w:noProof/></w:rPr><w:pict>"#);

        // Shapetype for diagonal text (NO v:group wrapper for watermarks!)
        xml.push_str(r#"<v:shapetype id="_x0000_t136" coordsize="21600,21600" o:spt="136" adj="10800" path="m@7,l@8,m@5,21600l@6,21600e">"#);

        // Formulas for shape
        xml.push_str(r#"<v:formulas>"#);
        xml.push_str(r#"<v:f eqn="sum #0 0 10800"/>"#);
        xml.push_str(r#"<v:f eqn="prod #0 2 1"/>"#);
        xml.push_str(r#"<v:f eqn="sum 21600 0 @1"/>"#);
        xml.push_str(r#"<v:f eqn="sum 0 0 @2"/>"#);
        xml.push_str(r#"<v:f eqn="sum 21600 0 @3"/>"#);
        xml.push_str(r#"<v:f eqn="if @0 @3 0"/>"#);
        xml.push_str(r#"<v:f eqn="if @0 21600 @1"/>"#);
        xml.push_str(r#"<v:f eqn="if @0 0 @2"/>"#);
        xml.push_str(r#"<v:f eqn="if @0 @4 21600"/>"#);
        xml.push_str(r#"<v:f eqn="mid @5 @6"/>"#);
        xml.push_str(r#"<v:f eqn="mid @8 @5"/>"#);
        xml.push_str(r#"<v:f eqn="mid @7 @8"/>"#);
        xml.push_str(r#"<v:f eqn="mid @6 @7"/>"#);
        xml.push_str(r#"<v:f eqn="sum @6 0 @5"/>"#);
        xml.push_str(r#"</v:formulas>"#);

        xml.push_str(r#"<v:path textpathok="t" o:connecttype="custom" "#);
        xml.push_str(r#"o:connectlocs="@9,0;@10,10800;@11,21600;@12,10800" "#);
        xml.push_str(r#"o:connectangles="270,180,90,0"/>"#);

        xml.push_str(r#"<v:textpath on="t" fitshape="t"/>"#);

        xml.push_str(r#"<v:handles>"#);
        xml.push_str(r##"<v:h position="#0,bottomRight" xrange="6629,14971"/>"##);
        xml.push_str(r#"</v:handles>"#);

        xml.push_str(r#"<o:lock v:ext="edit" text="t" shapetype="t"/>"#);

        xml.push_str(r#"</v:shapetype>"#);

        // Main shape
        xml.push_str(&format!(
            r#"<v:shape id="PowerPlusWaterMarkObject{}" "#,
            idx
        ));
        xml.push_str(&format!(r##"o:spid="_x0000_s102{}" "##, 4 + idx));
        xml.push_str(r##"type="#_x0000_t136" "##);
        xml.push_str(
            r#"style="position:absolute;margin-left:0;margin-top:0;width:439.9pt;height:219.95pt;"#,
        );
        xml.push_str(r#"rotation:315;z-index:-251655168;"#); // 315 degrees = diagonal
        xml.push_str(r#"mso-position-horizontal:center;mso-position-horizontal-relative:margin;"#);
        xml.push_str(r#"mso-position-vertical:center;mso-position-vertical-relative:margin" "#);
        xml.push_str(r#"o:allowincell="f" "#);
        // Add # prefix to color if not already present
        let color_with_hash = if self.color.starts_with('#')
            || !self.color.bytes().all(|byte| byte.is_ascii_hexdigit())
        {
            self.color.clone()
        } else {
            format!("#{}", self.color)
        };
        xml.push_str(&format!(r#"fillcolor="{}" "#, color_with_hash));
        xml.push_str(r#"stroked="f">"#);

        // Fill with opacity (separate element, not attribute!)
        xml.push_str(r#"<v:fill opacity=".5"/>"#);

        // Text path
        // Shape dimensions control the visual scale; the text-path point size
        // is still preserved for producer compatibility.
        xml.push_str(&format!(
            r#"<v:textpath style="font-family:&quot;{}&quot;;font-size:{}pt" "#,
            self.font, self.font_size
        ));
        xml.push_str(&format!(r#"string="{}"/>"#, escape_xml(&self.text)));

        xml.push_str(r#"</v:shape>"#);

        // End picture and run
        xml.push_str(r#"</w:pict></w:r></w:p>"#);

        Ok(xml)
    }
}

#[derive(Default)]
struct WatermarkCandidate {
    color: Option<String>,
    text: Option<String>,
    font: Option<String>,
    font_size: Option<f64>,
}

impl WatermarkCandidate {
    fn from_shape(
        element: &BytesStart<'_>,
        decoder: quick_xml::encoding::Decoder,
        resolver: &NamespaceResolver,
    ) -> Result<Option<Self>> {
        let id = unqualified_attribute(element, b"id", decoder, resolver)?;
        let shape_type = unqualified_attribute(element, b"type", decoder, resolver)?;
        let style = unqualified_attribute(element, b"style", decoder, resolver)?;
        let canonical_id = id
            .as_deref()
            .is_some_and(|value| value.starts_with("PowerPlusWaterMarkObject"));
        let background_word_art = shape_type.as_deref() == Some("#_x0000_t136")
            && style.as_deref().is_some_and(|value| {
                value.contains("z-index:-")
                    && value.contains("mso-position-horizontal:center")
                    && value.contains("mso-position-vertical:center")
            });
        if !canonical_id && !background_word_art {
            return Ok(None);
        }
        Ok(Some(Self {
            color: unqualified_attribute(element, b"fillcolor", decoder, resolver)?,
            ..Self::default()
        }))
    }

    fn finish(self) -> Option<Watermark> {
        Some(Watermark {
            text: self.text?,
            font: self.font.unwrap_or_else(|| "Cambria".to_string()),
            font_size: self.font_size.unwrap_or(1.0),
            color: self.color.unwrap_or_else(|| "black".to_string()),
        })
    }
}

fn is_vml(namespace: &ResolveResult<'_>) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == VML_NAMESPACE)
}

fn apply_text_path(
    shapes: &mut [(usize, Option<WatermarkCandidate>)],
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    resolver: &NamespaceResolver,
) -> Result<()> {
    let Some((_, Some(candidate))) = shapes.last_mut() else {
        return Ok(());
    };
    candidate.text = unqualified_attribute(element, b"string", decoder, resolver)?;
    if let Some(style) = unqualified_attribute(element, b"style", decoder, resolver)? {
        for declaration in style.split(';') {
            let Some((name, value)) = declaration.split_once(':') else {
                continue;
            };
            match name.trim().to_ascii_lowercase().as_str() {
                "font-family" => {
                    let value = value.trim().trim_matches(['\'', '"']);
                    if !value.is_empty() {
                        candidate.font = Some(value.to_string());
                    }
                },
                "font-size" => {
                    let value = value.trim().strip_suffix("pt").unwrap_or(value.trim());
                    if let Ok(size) = value.parse::<f64>()
                        && size.is_finite()
                        && size >= 0.0
                    {
                        candidate.font_size = Some(size);
                    }
                },
                _ => {},
            }
        }
    }
    Ok(())
}

fn unqualified_attribute(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: quick_xml::encoding::Decoder,
    resolver: &NamespaceResolver,
) -> Result<Option<String>> {
    let mut value = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| OoxmlError::Xml(error.to_string()))?;
        if attribute.key.local_name().as_ref() != name
            || !matches!(
                resolver.resolve_attribute(attribute.key).0,
                ResolveResult::Unbound
            )
        {
            continue;
        }
        if value.is_some() {
            return Err(OoxmlError::InvalidFormat(format!(
                "duplicate VML attribute '{}'",
                String::from_utf8_lossy(name)
            )));
        }
        value = Some(
            attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .map_err(|error| OoxmlError::Xml(error.to_string()))?
                .into_owned(),
        );
    }
    Ok(value)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_create_watermark() {
        let wm = Watermark::text("CONFIDENTIAL");
        assert_eq!(wm.get_text(), "CONFIDENTIAL");
    }

    #[test]
    fn test_watermark_customization() {
        let mut wm = Watermark::text("DRAFT");
        wm.set_font("Arial");
        wm.set_color("red");
        wm.set_font_size(2);

        let xml = wm.to_header_xml(1).unwrap();
        assert!(xml.contains("DRAFT"));
        assert!(xml.contains("Arial"));
    }

    #[test]
    fn test_watermark_xml() {
        let wm = Watermark::text("TEST");
        let xml = wm.to_header_xml(1).unwrap();

        assert!(xml.contains("<w:p>"));
        assert!(xml.contains("<w:pict>"));
        assert!(xml.contains("PowerPlusWaterMarkObject"));
        assert!(xml.contains("TEST"));
    }

    #[test]
    fn parses_generated_watermark_from_header_xml() {
        let mut expected = Watermark::text("DRAFT & REVIEW");
        expected.set_font("Arial");
        expected.set_font_size_points(2.5).unwrap();
        expected.set_color("C0FFEE");
        let body = expected.to_header_xml(1).unwrap();
        let xml = format!(
            r#"<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">{body}</w:hdr>"#
        );

        let parsed = Watermark::from_header_xml(xml.as_bytes()).unwrap();

        assert_eq!(parsed.len(), 1);
        assert_eq!(parsed[0].get_text(), "DRAFT & REVIEW");
        assert_eq!(parsed[0].font(), "Arial");
        assert_eq!(parsed[0].font_size(), 2.5);
        assert_eq!(parsed[0].color(), "#C0FFEE");
    }

    #[test]
    fn recognizes_aliased_vml_and_ignores_ordinary_or_spoofed_shapes() {
        let xml = br#"<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                xmlns:q="urn:schemas-microsoft-com:vml" xmlns:false="urn:not-vml">
            <q:shape id="ordinary"><q:textpath string="ordinary"/></q:shape>
            <false:shape id="PowerPlusWaterMarkObject0"><false:textpath string="spoofed"/></false:shape>
            <q:shape id="PowerPlusWaterMarkObject1" fillcolor="navy">
                <q:textpath style="font-family:'Aptos';font-size:3pt" string="REAL"/>
            </q:shape>
        </w:hdr>"#;

        let parsed = Watermark::from_header_xml(xml).unwrap();

        assert_eq!(parsed.len(), 1);
        assert_eq!(parsed[0].get_text(), "REAL");
        assert_eq!(parsed[0].font(), "Aptos");
        assert_eq!(parsed[0].font_size(), 3.0);
        assert_eq!(parsed[0].color(), "navy");
    }

    #[test]
    fn rejects_unterminated_watermark_header_xml() {
        let xml = br#"<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:v="urn:schemas-microsoft-com:vml"><v:shape id="PowerPlusWaterMarkObject1">"#;
        assert!(Watermark::from_header_xml(xml).is_err());
    }

    #[test]
    fn rejects_invalid_fractional_font_sizes() {
        let mut watermark = Watermark::text("DRAFT");
        assert!(watermark.set_font_size_points(f64::NAN).is_err());
        assert!(watermark.set_font_size_points(-1.0).is_err());
        assert_eq!(watermark.font_size(), 1.0);
    }
}
