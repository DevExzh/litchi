/// Theme writer support for DOCX documents.
///
/// Themes define the color scheme, fonts, and effects used in a document.
/// Based on Apache POI's XWPFTheme implementation.
use crate::ooxml::error::Result;
use std::fmt::Write as FmtWrite;

/// A mutable theme for document styling.
///
/// Themes provide coordinated colors, fonts, and effects that can be applied
/// throughout a document to maintain consistent styling.
///
/// # Examples
///
/// ```rust,ignore
/// use litchi::ooxml::docx::writer::MutableTheme;
///
/// let mut theme = MutableTheme::new("Office Theme");
/// theme.set_major_font("Calibri Light");
/// theme.set_minor_font("Calibri");
/// ```
#[derive(Debug, Clone)]
pub struct MutableTheme {
    /// Theme name
    name: String,
    /// Major font (for headings)
    major_font: String,
    /// Minor font (for body text)
    minor_font: String,
    /// Color scheme
    color_scheme: ColorScheme,
}

/// Color scheme for a theme.
///
/// Defines the 12 theme colors: dark1, light1, dark2, light2, accent1-6, hyperlink, followed hyperlink.
#[derive(Debug, Clone)]
pub struct ColorScheme {
    /// Color scheme name
    name: String,
    /// Dark 1 color (usually for text)
    dk1: String,
    /// Light 1 color (usually for background)
    lt1: String,
    /// Dark 2 color (secondary text)
    dk2: String,
    /// Light 2 color (secondary background)
    lt2: String,
    /// Accent colors (6 colors)
    accents: [String; 6],
    /// Hyperlink color
    hlink: String,
    /// Followed hyperlink color
    fol_hlink: String,
}

impl MutableTheme {
    /// Create a new theme with the given name.
    ///
    /// # Arguments
    ///
    /// * `name` - Theme name (e.g., "Office Theme")
    ///
    /// # Examples
    ///
    /// ```rust,ignore
    /// let theme = MutableTheme::new("My Custom Theme");
    /// ```
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            major_font: "Calibri Light".to_string(),
            minor_font: "Calibri".to_string(),
            color_scheme: ColorScheme::default(),
        }
    }

    /// Create the default Office theme.
    pub fn office_theme() -> Self {
        let mut theme = Self::new("Office Theme");
        theme.major_font = "Calibri Light".to_string();
        theme.minor_font = "Calibri".to_string();
        theme.color_scheme = ColorScheme::office();
        theme
    }

    /// Get the theme name.
    #[inline]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Set the theme name.
    pub fn set_name(&mut self, name: impl Into<String>) {
        self.name = name.into();
    }

    /// Get the major font.
    #[inline]
    pub fn major_font(&self) -> &str {
        &self.major_font
    }

    /// Set the major font (used for headings).
    pub fn set_major_font(&mut self, font: impl Into<String>) {
        self.major_font = font.into();
    }

    /// Get the minor font.
    #[inline]
    pub fn minor_font(&self) -> &str {
        &self.minor_font
    }

    /// Set the minor font (used for body text).
    pub fn set_minor_font(&mut self, font: impl Into<String>) {
        self.minor_font = font.into();
    }

    /// Get a mutable reference to the color scheme.
    pub fn color_scheme_mut(&mut self) -> &mut ColorScheme {
        &mut self.color_scheme
    }

    /// Generate theme1.xml content.
    pub(crate) fn to_xml(&self) -> Result<String> {
        let mut xml = String::with_capacity(4096);

        xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
        xml.push_str(
            r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" "#,
        );
        write!(&mut xml, r#"name="{}">"#, escape_xml(&self.name))?;

        // Theme elements
        xml.push_str("<a:themeElements>");

        // Color scheme
        self.color_scheme.to_xml(&mut xml)?;

        // Font scheme
        write!(&mut xml, r#"<a:fontScheme name="Office">"#)?;
        xml.push_str("<a:majorFont>");
        write!(
            &mut xml,
            r#"<a:latin typeface="{}"/>"#,
            escape_xml(&self.major_font)
        )?;
        write!(&mut xml, r#"<a:ea typeface=""/>"#)?;
        write!(&mut xml, r#"<a:cs typeface=""/>"#)?;
        xml.push_str("</a:majorFont>");

        xml.push_str("<a:minorFont>");
        write!(
            &mut xml,
            r#"<a:latin typeface="{}"/>"#,
            escape_xml(&self.minor_font)
        )?;
        write!(&mut xml, r#"<a:ea typeface=""/>"#)?;
        write!(&mut xml, r#"<a:cs typeface=""/>"#)?;
        xml.push_str("</a:minorFont>");
        xml.push_str("</a:fontScheme>");

        // Format scheme (minimal)
        xml.push_str(r#"<a:fmtScheme name="Office">"#);
        xml.push_str("<a:fillStyleLst>");
        xml.push_str(r#"<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>"#);
        xml.push_str(r#"<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="35000"><a:schemeClr val="phClr"><a:tint val="37000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="15000"/><a:satMod val="350000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="16200000" scaled="1"/></a:gradFill>"#);
        xml.push_str(r#"<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="100000"/><a:shade val="100000"/><a:satMod val="130000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="50000"/><a:shade val="100000"/><a:satMod val="350000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="16200000" scaled="0"/></a:gradFill>"#);
        xml.push_str("</a:fillStyleLst>");
        xml.push_str("<a:lnStyleLst>");
        xml.push_str(r#"<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"><a:shade val="95000"/><a:satMod val="105000"/></a:schemeClr></a:solidFill><a:prstDash val="solid"/></a:ln>"#);
        xml.push_str(r#"<a:ln w="25400" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>"#);
        xml.push_str(r#"<a:ln w="38100" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>"#);
        xml.push_str("</a:lnStyleLst>");
        xml.push_str("<a:effectStyleLst>");
        xml.push_str("<a:effectStyle><a:effectLst><a:outerShdw blurRad=\"40000\" dist=\"20000\" dir=\"5400000\" rotWithShape=\"0\"><a:srgbClr val=\"000000\"><a:alpha val=\"38000\"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle>");
        xml.push_str("<a:effectStyle><a:effectLst><a:outerShdw blurRad=\"40000\" dist=\"23000\" dir=\"5400000\" rotWithShape=\"0\"><a:srgbClr val=\"000000\"><a:alpha val=\"35000\"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle>");
        xml.push_str("<a:effectStyle><a:effectLst><a:outerShdw blurRad=\"40000\" dist=\"23000\" dir=\"5400000\" rotWithShape=\"0\"><a:srgbClr val=\"000000\"><a:alpha val=\"35000\"/></a:srgbClr></a:outerShdw></a:effectLst><a:scene3d><a:camera prst=\"orthographicFront\"><a:rot lat=\"0\" lon=\"0\" rev=\"0\"/></a:camera><a:lightRig rig=\"threePt\" dir=\"t\"><a:rot lat=\"0\" lon=\"0\" rev=\"1200000\"/></a:lightRig></a:scene3d><a:sp3d><a:bevelT w=\"63500\" h=\"25400\"/></a:sp3d></a:effectStyle>");
        xml.push_str("</a:effectStyleLst>");
        xml.push_str("<a:bgFillStyleLst>");
        xml.push_str(r#"<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>"#);
        xml.push_str(r#"<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="40000"/><a:satMod val="350000"/></a:schemeClr></a:gs><a:gs pos="40000"><a:schemeClr val="phClr"><a:tint val="45000"/><a:shade val="99000"/><a:satMod val="350000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="20000"/><a:satMod val="255000"/></a:schemeClr></a:gs></a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="-80000" r="50000" b="180000"/></a:path></a:gradFill>"#);
        xml.push_str(r#"<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="80000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="30000"/><a:satMod val="200000"/></a:schemeClr></a:gs></a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path></a:gradFill>"#);
        xml.push_str("</a:bgFillStyleLst>");
        xml.push_str("</a:fmtScheme>");

        xml.push_str("</a:themeElements>");

        // Object defaults (minimal)
        xml.push_str("<a:objectDefaults/>");

        // Extra color scheme list (empty)
        xml.push_str("<a:extraClrSchemeLst/>");

        xml.push_str("</a:theme>");

        Ok(xml)
    }
}

impl ColorScheme {
    /// Create a new color scheme with the given name.
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            dk1: "000000".to_string(),
            lt1: "FFFFFF".to_string(),
            dk2: "44546A".to_string(),
            lt2: "E7E6E6".to_string(),
            accents: [
                "4472C4".to_string(),
                "ED7D31".to_string(),
                "A5A5A5".to_string(),
                "FFC000".to_string(),
                "5B9BD5".to_string(),
                "70AD47".to_string(),
            ],
            hlink: "0563C1".to_string(),
            fol_hlink: "954F72".to_string(),
        }
    }

    /// Create the default Office color scheme.
    pub fn office() -> Self {
        Self::new("Office")
    }

    /// Set an accent color by index (0-5).
    pub fn set_accent(&mut self, index: usize, color: impl Into<String>) {
        if index < 6 {
            self.accents[index] = color.into();
        }
    }

    /// Generate color scheme XML.
    fn to_xml(&self, xml: &mut String) -> Result<()> {
        write!(xml, r#"<a:clrScheme name="{}">"#, escape_xml(&self.name))?;

        // Dark 1 and Light 1
        write!(xml, r#"<a:dk1><a:srgbClr val="{}"/></a:dk1>"#, self.dk1)?;
        write!(xml, r#"<a:lt1><a:srgbClr val="{}"/></a:lt1>"#, self.lt1)?;

        // Dark 2 and Light 2
        write!(xml, r#"<a:dk2><a:srgbClr val="{}"/></a:dk2>"#, self.dk2)?;
        write!(xml, r#"<a:lt2><a:srgbClr val="{}"/></a:lt2>"#, self.lt2)?;

        // Accent colors
        for (i, accent) in self.accents.iter().enumerate() {
            write!(
                xml,
                r#"<a:accent{0}><a:srgbClr val="{1}"/></a:accent{0}>"#,
                i + 1,
                accent
            )?;
        }

        // Hyperlinks
        write!(
            xml,
            r#"<a:hlink><a:srgbClr val="{}"/></a:hlink>"#,
            self.hlink
        )?;
        write!(
            xml,
            r#"<a:folHlink><a:srgbClr val="{}"/></a:folHlink>"#,
            self.fol_hlink
        )?;

        xml.push_str("</a:clrScheme>");

        Ok(())
    }
}

impl Default for ColorScheme {
    fn default() -> Self {
        Self::office()
    }
}

/// Escape XML special characters.
fn escape_xml(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_create_theme() {
        let theme = MutableTheme::new("Test Theme");
        assert_eq!(theme.name(), "Test Theme");
        assert_eq!(theme.major_font(), "Calibri Light");
        assert_eq!(theme.minor_font(), "Calibri");
    }

    #[test]
    fn test_office_theme() {
        let theme = MutableTheme::office_theme();
        assert_eq!(theme.name(), "Office Theme");
    }

    #[test]
    fn test_set_fonts() {
        let mut theme = MutableTheme::new("Custom");
        theme.set_major_font("Arial");
        theme.set_minor_font("Times New Roman");

        assert_eq!(theme.major_font(), "Arial");
        assert_eq!(theme.minor_font(), "Times New Roman");
    }

    #[test]
    fn test_color_scheme() {
        let mut scheme = ColorScheme::new("Custom Colors");
        scheme.set_accent(0, "FF0000");
        scheme.set_accent(5, "00FF00");

        assert_eq!(scheme.accents[0], "FF0000");
        assert_eq!(scheme.accents[5], "00FF00");
    }

    #[test]
    fn test_xml_generation() {
        let theme = MutableTheme::office_theme();
        let xml = theme.to_xml().unwrap();

        assert!(xml.contains("<?xml version"));
        assert!(xml.contains(r#"name="Office Theme""#));
        assert!(xml.contains("a:themeElements"));
        assert!(xml.contains("a:clrScheme"));
        assert!(xml.contains("a:fontScheme"));
        assert!(xml.contains("a:fmtScheme"));
    }
}
