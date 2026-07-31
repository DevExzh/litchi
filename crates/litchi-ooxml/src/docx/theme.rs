/// Theme support for Word documents.
///
/// Themes define the color scheme, fonts, and effects used in a document.
use litchi_ooxml_common::xml::{is_drawingml_name, unqualified_attribute_value};
use crate::error::{OoxmlError, Result};
use litchi_opc::part::Part;
use quick_xml::encoding::Decoder;
use quick_xml::events::BytesStart;
use quick_xml::events::Event;
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;

/// Document theme containing color scheme, font scheme, and format scheme.
///
/// # Examples
///
/// ```rust,no_run
/// use litchi_ooxml::docx::Package;
///
/// let pkg = Package::open("document.docx")?;
/// let doc = pkg.document()?;
///
/// if let Some(theme) = doc.theme()? {
///     if let Some(name) = theme.name() {
///         println!("Theme: {}", name);
///     }
/// }
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
#[derive(Debug, Clone)]
pub struct Theme {
    /// Theme name
    name: Option<String>,
    /// Major font (for headings)
    major_font: Option<String>,
    /// Minor font (for body text)
    minor_font: Option<String>,
    /// Color scheme name
    color_scheme: Option<String>,
}

impl Theme {
    /// Create a new empty Theme.
    pub fn new() -> Self {
        Self {
            name: None,
            major_font: None,
            minor_font: None,
            color_scheme: None,
        }
    }

    /// Get the theme name.
    #[inline]
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    /// Get the major font (for headings).
    #[inline]
    pub fn major_font(&self) -> Option<&str> {
        self.major_font.as_deref()
    }

    /// Get the minor font (for body text).
    #[inline]
    pub fn minor_font(&self) -> Option<&str> {
        self.minor_font.as_deref()
    }

    /// Get the color scheme name.
    #[inline]
    pub fn color_scheme(&self) -> Option<&str> {
        self.color_scheme.as_deref()
    }

    /// Extract theme from a theme part.
    pub(crate) fn extract_from_part(part: &dyn Part) -> Result<Self> {
        let xml_bytes = litchi_ooxml_common::mce::process_part(part)?;
        let mut reader = NsReader::from_reader(xml_bytes.as_ref());

        let mut theme = Self::new();
        let mut in_major_font = false;
        let mut in_minor_font = false;
        let mut depth = 0usize;

        loop {
            let decoder = reader.decoder();
            let (namespace, event) = reader
                .read_resolved_event()
                .map_err(|error| OoxmlError::Xml(error.to_string()))?;
            match event {
                Event::Start(element) => {
                    depth = depth.checked_add(1).ok_or_else(|| {
                        OoxmlError::InvalidFormat("DrawingML nesting is too deep".to_string())
                    })?;
                    parse_theme_element(
                        &mut theme,
                        &namespace,
                        &element,
                        decoder,
                        in_major_font,
                        in_minor_font,
                    )?;
                    if is_drawingml_name(&namespace, element.name(), b"majorFont") {
                        in_major_font = true;
                    } else if is_drawingml_name(&namespace, element.name(), b"minorFont") {
                        in_minor_font = true;
                    }
                },
                Event::Empty(element) => parse_theme_element(
                    &mut theme,
                    &namespace,
                    &element,
                    decoder,
                    in_major_font,
                    in_minor_font,
                )?,
                Event::End(element) => {
                    if is_drawingml_name(&namespace, element.name(), b"majorFont") {
                        in_major_font = false;
                    } else if is_drawingml_name(&namespace, element.name(), b"minorFont") {
                        in_minor_font = false;
                    }
                    depth = depth.checked_sub(1).ok_or_else(|| {
                        OoxmlError::InvalidFormat("invalid DrawingML nesting".to_string())
                    })?;
                },
                Event::Eof if depth != 0 => {
                    return Err(OoxmlError::InvalidFormat(
                        "unterminated DrawingML theme XML".to_string(),
                    ));
                },
                Event::Eof => break,
                _ => {},
            }
        }

        Ok(theme)
    }
}

fn parse_theme_element(
    theme: &mut Theme,
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    in_major_font: bool,
    in_minor_font: bool,
) -> Result<()> {
    if is_drawingml_name(namespace, element.name(), b"theme") {
        theme.name = unqualified_attribute_value(element, b"name", decoder)?;
    } else if is_drawingml_name(namespace, element.name(), b"clrScheme") {
        theme.color_scheme = unqualified_attribute_value(element, b"name", decoder)?;
    } else if is_drawingml_name(namespace, element.name(), b"latin")
        && let Some(font) = unqualified_attribute_value(element, b"typeface", decoder)?
    {
        if in_major_font && theme.major_font.is_none() {
            theme.major_font = Some(font);
        } else if in_minor_font && theme.minor_font.is_none() {
            theme.minor_font = Some(font);
        }
    }
    Ok(())
}

impl Default for Theme {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::PackURI;
    use litchi_opc::part::BlobPart;

    fn parse_theme(xml: &[u8]) -> Result<Theme> {
        let part = BlobPart::new(
            PackURI::new("/word/theme/theme1.xml").unwrap(),
            "application/xml".to_string(),
            xml.to_vec(),
        );
        Theme::extract_from_part(&part)
    }

    #[test]
    fn test_theme_creation() {
        let theme = Theme::new();
        assert!(theme.name().is_none());
        assert!(theme.major_font().is_none());
        assert!(theme.minor_font().is_none());
        assert!(theme.color_scheme().is_none());
    }

    #[test]
    fn parses_aliased_theme_names_and_fonts() {
        let theme = parse_theme(
            br#"<d:theme xmlns:d="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:false="urn:not-drawingml" name="A &amp; B">
                <d:themeElements><d:clrScheme name="Office &amp; More"/>
                <d:fontScheme><false:majorFont><d:latin typeface="ignored"/></false:majorFont>
                <d:majorFont><false:latin typeface="ignored"/><d:latin typeface="Major &amp; Co"/></d:majorFont>
                <d:minorFont><d:latin typeface="Minor"/></d:minorFont></d:fontScheme></d:themeElements>
            </d:theme>"#,
        )
        .unwrap();
        assert_eq!(theme.name(), Some("A & B"));
        assert_eq!(theme.color_scheme(), Some("Office & More"));
        assert_eq!(theme.major_font(), Some("Major & Co"));
        assert_eq!(theme.minor_font(), Some("Minor"));
    }

    #[test]
    fn parses_strict_theme_and_rejects_malformed_xml() {
        let strict = parse_theme(
            br#"<s:theme xmlns:s="http://purl.oclc.org/ooxml/drawingml/main" name="Strict"><s:themeElements><s:fontScheme><s:majorFont><s:latin typeface="Heading"/></s:majorFont></s:fontScheme></s:themeElements></s:theme>"#,
        )
        .unwrap();
        assert_eq!(strict.name(), Some("Strict"));
        assert_eq!(strict.major_font(), Some("Heading"));

        assert!(parse_theme(
            br#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="one" name="two"/>"#
        )
        .is_err());
        assert!(parse_theme(
            br#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements>"#
        )
        .is_err());
    }
}
