//! Watermark support for DOCX documents.
//!
//! Based on Apache POI's XWPFHeaderFooterPolicy watermark implementation.
use crate::common::xml::escape_xml;
use crate::ooxml::error::Result;

/// A watermark for a Word document.
///
/// Watermarks are typically displayed as diagonal text behind the document content.
///
/// # Examples
///
/// ```rust,ignore
/// use litchi::ooxml::docx::writer::Watermark;
///
/// let watermark = Watermark::text("CONFIDENTIAL");
/// ```
#[derive(Debug, Clone)]
pub struct Watermark {
    /// Watermark text
    text: String,
    /// Font family (default: Cambria)
    font: String,
    /// Font size in points (default: 1pt for shape-based watermarks)
    font_size: u32,
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
            font_size: 1,
            color: "black".to_string(),
        }
    }

    /// Set the font family.
    pub fn set_font(&mut self, font: impl Into<String>) {
        self.font = font.into();
    }

    /// Set the font size in points.
    pub fn set_font_size(&mut self, size: u32) {
        self.font_size = size;
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
        let color_with_hash = if self.color.starts_with('#') {
            self.color.clone()
        } else {
            format!("#{}", self.color)
        };
        xml.push_str(&format!(r#"fillcolor="{}" "#, color_with_hash));
        xml.push_str(r#"stroked="f">"#);

        // Fill with opacity (separate element, not attribute!)
        xml.push_str(r#"<v:fill opacity=".5"/>"#);

        // Text path
        // NOTE: font-size should always be 1pt for watermarks - the actual size is controlled by shape dimensions
        xml.push_str(&format!(
            r#"<v:textpath style="font-family:&quot;{}&quot;;font-size:1pt" "#,
            self.font
        ));
        xml.push_str(&format!(r#"string="{}"/>"#, escape_xml(&self.text)));

        xml.push_str(r#"</v:shape>"#);

        // End picture and run
        xml.push_str(r#"</w:pict></w:r></w:p>"#);

        Ok(xml)
    }
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
}
