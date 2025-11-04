/// Theme support for Word documents.
///
/// Themes define the color scheme, fonts, and effects used in a document.
use crate::ooxml::error::{OoxmlError, Result};
use crate::ooxml::opc::part::Part;
use quick_xml::Reader;
use quick_xml::events::Event;

/// Document theme containing color scheme, font scheme, and format scheme.
///
/// # Examples
///
/// ```rust,no_run
/// use litchi::ooxml::docx::Package;
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
        let xml_bytes = part.blob();
        let mut reader = Reader::from_reader(xml_bytes);
        reader.config_mut().trim_text(true);

        let mut theme = Self::new();
        let mut in_major_font = false;
        let mut in_minor_font = false;
        let mut buf = Vec::with_capacity(1024);

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => match e.local_name().as_ref() {
                    b"theme" => {
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"name" {
                                theme.name =
                                    Some(String::from_utf8_lossy(&attr.value).into_owned());
                            }
                        }
                    },
                    b"clrScheme" => {
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"name" {
                                theme.color_scheme =
                                    Some(String::from_utf8_lossy(&attr.value).into_owned());
                            }
                        }
                    },
                    b"majorFont" => {
                        in_major_font = true;
                    },
                    b"minorFont" => {
                        in_minor_font = true;
                    },
                    b"latin" => {
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"typeface" {
                                let font = String::from_utf8_lossy(&attr.value).into_owned();
                                if in_major_font && theme.major_font.is_none() {
                                    theme.major_font = Some(font);
                                } else if in_minor_font && theme.minor_font.is_none() {
                                    theme.minor_font = Some(font);
                                }
                            }
                        }
                    },
                    _ => {},
                },
                Ok(Event::End(e)) => match e.local_name().as_ref() {
                    b"majorFont" => in_major_font = false,
                    b"minorFont" => in_minor_font = false,
                    _ => {},
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
            buf.clear();
        }

        Ok(theme)
    }
}

impl Default for Theme {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_theme_creation() {
        let theme = Theme::new();
        assert!(theme.name().is_none());
        assert!(theme.major_font().is_none());
        assert!(theme.minor_font().is_none());
        assert!(theme.color_scheme().is_none());
    }
}
