/// Theme parts for PowerPoint presentations.
///
/// This module provides types for working with themes in PPTX files.
use crate::ooxml::error::{OoxmlError, Result};
use crate::ooxml::opc::part::Part;
use quick_xml::Reader;
use quick_xml::events::Event;

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
        let mut reader = Reader::from_reader(self.xml_bytes());
        reader.config_mut().trim_text(true);

        let mut theme_name = String::new();
        let mut major_font: Option<ThemeFont> = None;
        let mut minor_font: Option<ThemeFont> = None;
        let mut colors = Vec::new();

        let mut in_major_font = false;
        let mut in_minor_font = false;
        let mut in_color_scheme = false;
        let mut current_color_name = String::new();

        loop {
            match reader.read_event() {
                Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                    let tag_name = e.local_name();

                    match tag_name.as_ref() {
                        b"theme" => {
                            // Extract theme name
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"name" {
                                    theme_name = std::str::from_utf8(&attr.value)
                                        .map(|s| s.to_string())
                                        .unwrap_or_default();
                                }
                            }
                        },
                        b"clrScheme" => {
                            in_color_scheme = true;
                        },
                        b"majorFont" => {
                            in_major_font = true;
                        },
                        b"minorFont" => {
                            in_minor_font = true;
                        },
                        b"latin" if in_major_font || in_minor_font => {
                            // Extract font typeface
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"typeface" {
                                    let typeface = std::str::from_utf8(&attr.value)
                                        .map(|s| s.to_string())
                                        .unwrap_or_default();

                                    let font = ThemeFont {
                                        typeface,
                                        charset: None,
                                    };

                                    if in_major_font {
                                        major_font = Some(font);
                                    } else if in_minor_font {
                                        minor_font = Some(font);
                                    }
                                }
                            }
                        },
                        // Color elements in color scheme
                        b"dk1" | b"lt1" | b"dk2" | b"lt2" | b"accent1" | b"accent2"
                        | b"accent3" | b"accent4" | b"accent5" | b"accent6" | b"hlink"
                        | b"folHlink"
                            if in_color_scheme =>
                        {
                            current_color_name = std::str::from_utf8(tag_name.as_ref())
                                .unwrap_or("")
                                .to_string();
                        },
                        b"srgbClr" if in_color_scheme && !current_color_name.is_empty() => {
                            // RGB color
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"val" {
                                    let rgb = std::str::from_utf8(&attr.value)
                                        .map(|s| s.to_string())
                                        .ok();

                                    colors.push(ThemeColor {
                                        name: current_color_name.clone(),
                                        rgb,
                                        system_color: None,
                                    });
                                    current_color_name.clear();
                                }
                            }
                        },
                        b"sysClr" if in_color_scheme && !current_color_name.is_empty() => {
                            // System color
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"val" {
                                    let sys_color = std::str::from_utf8(&attr.value)
                                        .map(|s| s.to_string())
                                        .ok();

                                    colors.push(ThemeColor {
                                        name: current_color_name.clone(),
                                        rgb: None,
                                        system_color: sys_color,
                                    });
                                    current_color_name.clear();
                                }
                            }
                        },
                        _ => {},
                    }
                },
                Ok(Event::End(e)) => {
                    let tag_name = e.local_name();
                    match tag_name.as_ref() {
                        b"clrScheme" => in_color_scheme = false,
                        b"majorFont" => in_major_font = false,
                        b"minorFont" => in_minor_font = false,
                        _ => {},
                    }
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
        }

        Ok(Theme {
            name: theme_name,
            major_font,
            minor_font,
            colors,
        })
    }

    /// Get the underlying OPC part.
    #[inline]
    pub fn part(&self) -> &'a dyn Part {
        self.part
    }
}
