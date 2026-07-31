/// Theme parts for PowerPoint presentations.
///
/// This module provides types for working with themes in PPTX files.
use litchi_ooxml_common::xml::{is_drawingml_name, unqualified_attribute_value};
use crate::error::{OoxmlError, Result};
use litchi_opc::part::Part;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{QName, ResolveResult};
use quick_xml::reader::NsReader;

/// Color information from a theme.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ThemeColor {
    /// Color name (e.g., "accent1", "dk1", "lt1")
    pub name: String,
    /// RGB color value if available (format: "RRGGBB")
    pub rgb: Option<String>,
    /// System color if available
    pub system_color: Option<String>,
}

/// Font information from a theme.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ThemeFont {
    /// Font typeface name
    pub typeface: String,
    /// Font character set
    pub charset: Option<String>,
}

/// Theme information extracted from a theme part.
#[derive(Debug, Clone)]
pub struct Theme {
    /// Theme name
    pub name: String,
    /// Major (heading) font
    pub major_font: Option<ThemeFont>,
    /// Minor (body) font
    pub minor_font: Option<ThemeFont>,
    /// Color scheme colors
    pub colors: Vec<ThemeColor>,
}

/// Theme part - defines the visual styling of a presentation.
///
/// Corresponds to `/ppt/theme/themeN.xml` in the package.
pub struct ThemePart<'a> {
    /// The underlying OPC part
    part: &'a dyn Part,
}

impl<'a> ThemePart<'a> {
    /// Create a ThemePart from an OPC Part.
    pub fn from_part(part: &'a dyn Part) -> Result<Self> {
        Ok(Self { part })
    }

    /// Get the XML bytes of the theme.
    #[inline]
    fn xml_bytes(&self) -> &[u8] {
        self.part.blob()
    }

    /// Parse and return the theme information.
    ///
    /// # Examples
    ///
    /// ```rust,ignore
    /// let theme_part = ThemePart::from_part(part)?;
    /// let theme = theme_part.theme()?;
    /// println!("Theme name: {}", theme.name);
    /// ```
    pub fn theme(&self) -> Result<Theme> {
        let xml = litchi_ooxml_common::mce::process_ooxml(self.xml_bytes())?;
        let mut reader = NsReader::from_reader(xml.as_ref());
        let mut theme = Theme {
            name: String::new(),
            major_font: None,
            minor_font: None,
            colors: Vec::new(),
        };
        let mut context = ThemeContext::default();
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
                    parse_theme_element(&mut theme, &namespace, &element, decoder, &mut context)?;
                    if is_drawingml_name(&namespace, element.name(), b"clrScheme") {
                        context.in_color_scheme = true;
                    } else if is_drawingml_name(&namespace, element.name(), b"majorFont") {
                        context.in_major_font = true;
                    } else if is_drawingml_name(&namespace, element.name(), b"minorFont") {
                        context.in_minor_font = true;
                    }
                },
                Event::Empty(element) => {
                    parse_theme_element(&mut theme, &namespace, &element, decoder, &mut context)?;
                    if color_slot_name(&namespace, element.name()).is_some() {
                        context.current_color_name = None;
                    }
                },
                Event::End(element) => {
                    if is_drawingml_name(&namespace, element.name(), b"clrScheme") {
                        context.in_color_scheme = false;
                        context.current_color_name = None;
                    } else if is_drawingml_name(&namespace, element.name(), b"majorFont") {
                        context.in_major_font = false;
                    } else if is_drawingml_name(&namespace, element.name(), b"minorFont") {
                        context.in_minor_font = false;
                    } else if color_slot_name(&namespace, element.name()).is_some() {
                        context.current_color_name = None;
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
    /// Get the underlying OPC part.
    #[inline]
    pub fn part(&self) -> &'a dyn Part {
        self.part
    }
}

#[derive(Default)]
struct ThemeContext {
    in_major_font: bool,
    in_minor_font: bool,
    in_color_scheme: bool,
    current_color_name: Option<&'static str>,
}

fn parse_theme_element(
    theme: &mut Theme,
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    context: &mut ThemeContext,
) -> Result<()> {
    if is_drawingml_name(namespace, element.name(), b"theme") {
        if let Some(name) = unqualified_attribute_value(element, b"name", decoder)? {
            theme.name = name;
        }
    } else if is_drawingml_name(namespace, element.name(), b"latin")
        && (context.in_major_font || context.in_minor_font)
        && let Some(typeface) = unqualified_attribute_value(element, b"typeface", decoder)?
    {
        let font = ThemeFont {
            typeface,
            charset: unqualified_attribute_value(element, b"charset", decoder)?,
        };
        if context.in_major_font && theme.major_font.is_none() {
            theme.major_font = Some(font);
        } else if context.in_minor_font && theme.minor_font.is_none() {
            theme.minor_font = Some(font);
        }
    } else if context.in_color_scheme
        && let Some(color_name) = color_slot_name(namespace, element.name())
    {
        context.current_color_name = Some(color_name);
    } else if context.in_color_scheme
        && is_drawingml_name(namespace, element.name(), b"srgbClr")
        && let Some(color_name) = context.current_color_name.take()
    {
        let rgb = required_color_attribute(element, b"val", decoder, "srgbClr")?;
        validate_rgb(&rgb)?;
        theme.colors.push(ThemeColor {
            name: color_name.to_string(),
            rgb: Some(rgb),
            system_color: None,
        });
    } else if context.in_color_scheme
        && is_drawingml_name(namespace, element.name(), b"sysClr")
        && let Some(color_name) = context.current_color_name.take()
    {
        let system_color = required_color_attribute(element, b"val", decoder, "sysClr")?;
        let rgb = unqualified_attribute_value(element, b"lastClr", decoder)?;
        if let Some(rgb) = &rgb {
            validate_rgb(rgb)?;
        }
        theme.colors.push(ThemeColor {
            name: color_name.to_string(),
            rgb,
            system_color: Some(system_color),
        });
    }
    Ok(())
}

fn color_slot_name(namespace: &ResolveResult<'_>, name: QName<'_>) -> Option<&'static str> {
    let slot = match name.local_name().as_ref() {
        b"dk1" => "dk1",
        b"lt1" => "lt1",
        b"dk2" => "dk2",
        b"lt2" => "lt2",
        b"accent1" => "accent1",
        b"accent2" => "accent2",
        b"accent3" => "accent3",
        b"accent4" => "accent4",
        b"accent5" => "accent5",
        b"accent6" => "accent6",
        b"hlink" => "hlink",
        b"folHlink" => "folHlink",
        _ => return None,
    };
    is_drawingml_name(namespace, name, slot.as_bytes()).then_some(slot)
}

fn required_color_attribute(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    element_name: &str,
) -> Result<String> {
    unqualified_attribute_value(element, name, decoder)?.ok_or_else(|| {
        OoxmlError::InvalidFormat(format!(
            "DrawingML {element_name} is missing its required {} attribute",
            String::from_utf8_lossy(name)
        ))
    })
}

fn validate_rgb(rgb: &str) -> Result<()> {
    if rgb.len() == 6 && rgb.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        return Ok(());
    }
    Err(OoxmlError::InvalidFormat(format!(
        "invalid DrawingML RGB color '{rgb}'"
    )))
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::PackURI;
    use litchi_opc::part::BlobPart;

    fn parse_theme(xml: &[u8]) -> Result<Theme> {
        let part = BlobPart::new(
            PackURI::new("/ppt/theme/theme1.xml").unwrap(),
            "application/xml".to_string(),
            xml.to_vec(),
        );
        ThemePart::from_part(&part)?.theme()
    }

    #[test]
    fn parses_aliased_theme_fonts_and_colors() {
        let theme = parse_theme(
            br#"<d:theme xmlns:d="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:false="urn:not-drawingml" name="A &amp; B">
                <d:themeElements>
                    <d:clrScheme name="Office">
                        <d:dk1><d:sysClr val="windowText" lastClr="000000"/></d:dk1>
                        <d:lt1><d:srgbClr val="FFFFFF"/></d:lt1>
                        <false:accent1><d:srgbClr val="111111"/></false:accent1>
                        <d:accent1><false:srgbClr val="222222"/><d:srgbClr val="4F81BD"/></d:accent1>
                    </d:clrScheme>
                    <d:fontScheme><d:majorFont><d:latin typeface="Major &amp; Co" charset="01"/></d:majorFont>
                    <d:minorFont><d:latin typeface="Minor"/></d:minorFont></d:fontScheme>
                </d:themeElements>
            </d:theme>"#,
        )
        .unwrap();

        assert_eq!(theme.name, "A & B");
        assert_eq!(
            theme.major_font,
            Some(ThemeFont {
                typeface: "Major & Co".to_string(),
                charset: Some("01".to_string()),
            })
        );
        assert_eq!(
            theme.minor_font,
            Some(ThemeFont {
                typeface: "Minor".to_string(),
                charset: None,
            })
        );
        assert_eq!(
            theme.colors,
            vec![
                ThemeColor {
                    name: "dk1".to_string(),
                    rgb: Some("000000".to_string()),
                    system_color: Some("windowText".to_string()),
                },
                ThemeColor {
                    name: "lt1".to_string(),
                    rgb: Some("FFFFFF".to_string()),
                    system_color: None,
                },
                ThemeColor {
                    name: "accent1".to_string(),
                    rgb: Some("4F81BD".to_string()),
                    system_color: None,
                },
            ]
        );
    }

    #[test]
    fn parses_strict_theme_and_rejects_invalid_colors_or_xml() {
        let strict = parse_theme(
            br#"<s:theme xmlns:s="http://purl.oclc.org/ooxml/drawingml/main" name="Strict"><s:themeElements><s:clrScheme name="Strict"><s:accent2><s:srgbClr val="ABCDEF"/></s:accent2></s:clrScheme></s:themeElements></s:theme>"#,
        )
        .unwrap();
        assert_eq!(strict.name, "Strict");
        assert_eq!(strict.colors[0].rgb.as_deref(), Some("ABCDEF"));

        assert!(parse_theme(
            br#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:clrScheme name="x"><a:accent1><a:srgbClr val="XYZ"/></a:accent1></a:clrScheme></a:themeElements></a:theme>"#
        )
        .is_err());
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
