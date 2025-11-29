/// Styles - document styles and formatting definitions.
use crate::ooxml::docx::enums::WdStyleType;
use crate::ooxml::error::{OoxmlError, Result};
use crate::ooxml::opc::part::Part;
use quick_xml::Reader;
use quick_xml::events::Event;
use smallvec::SmallVec;

/// A collection of styles defined in a Word document.
///
/// Provides access to paragraph, character, table, and list styles.
/// Supports iteration and lookup by style ID or name.
///
/// # Examples
///
/// ```rust,ignore
/// use litchi::ooxml::docx::Package;
///
/// let pkg = Package::open("document.docx")?;
/// let doc = pkg.document()?;
/// let styles = doc.styles()?;
///
/// println!("Document has {} styles", styles.len());
/// for style in styles.iter() {
///     println!("Style: {} (type: {})", style.name(), style.style_type());
/// }
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
pub struct Styles<'a> {
    /// Reference to the styles part
    part: &'a dyn Part,
    /// Cached list of styles
    style_list: Option<SmallVec<[Style; 32]>>,
}

impl<'a> std::fmt::Debug for Styles<'a> {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("Styles")
            .field(
                "style_count",
                &self.style_list.as_ref().map(|s| s.len()).unwrap_or(0),
            )
            .finish()
    }
}

impl<'a> Styles<'a> {
    /// Create a new Styles object from a styles part.
    ///
    /// This is typically called internally when accessing document styles.
    #[inline]
    pub fn from_part(part: &'a dyn Part) -> Self {
        Self {
            part,
            style_list: None,
        }
    }

    /// Get the number of styles in the document.
    pub fn len(&mut self) -> Result<usize> {
        self.ensure_styles_loaded()?;
        Ok(self.style_list.as_ref().map_or(0, |list| list.len()))
    }

    /// Check if there are no styles defined.
    pub fn is_empty(&mut self) -> Result<bool> {
        Ok(self.len()? == 0)
    }

    /// Get an iterator over all styles.
    pub fn iter(&mut self) -> Result<std::slice::Iter<'_, Style>> {
        self.ensure_styles_loaded()?;
        Ok(self
            .style_list
            .as_ref()
            .map_or([].iter(), |list| list.iter()))
    }

    /// Get a style by its ID.
    ///
    /// Returns `None` if no style with the given ID is found.
    pub fn get_by_id(&mut self, style_id: &str) -> Result<Option<&Style>> {
        self.ensure_styles_loaded()?;
        Ok(self
            .style_list
            .as_ref()
            .and_then(|list| list.iter().find(|s| s.style_id == style_id)))
    }

    /// Get a style by its name.
    ///
    /// Returns `None` if no style with the given name is found.
    pub fn get_by_name(&mut self, name: &str) -> Result<Option<&Style>> {
        self.ensure_styles_loaded()?;
        Ok(self
            .style_list
            .as_ref()
            .and_then(|list| list.iter().find(|s| s.name.as_deref() == Some(name))))
    }

    /// Get the default style for a given style type.
    ///
    /// Returns `None` if no default style is defined for that type.
    pub fn get_default(&mut self, style_type: WdStyleType) -> Result<Option<&Style>> {
        self.ensure_styles_loaded()?;
        Ok(self.style_list.as_ref().and_then(|list| {
            list.iter()
                .find(|s| s.is_default && s.style_type == style_type)
        }))
    }

    /// Ensure styles are loaded from XML.
    fn ensure_styles_loaded(&mut self) -> Result<()> {
        if self.style_list.is_some() {
            return Ok(());
        }

        let xml_bytes = self.part.blob();
        let mut reader = Reader::from_reader(xml_bytes);
        reader.config_mut().trim_text(true);

        let mut styles = SmallVec::new();
        let mut current_style: Option<StyleBuilder> = None;

        loop {
            match reader.read_event() {
                Ok(Event::Start(e)) if e.local_name().as_ref() == b"style" => {
                    // Start a new style
                    let mut builder = StyleBuilder::default();

                    // Parse attributes
                    for attr in e.attributes().flatten() {
                        match attr.key.local_name().as_ref() {
                            b"type" => {
                                if let Ok(value) = attr.decode_and_unescape_value(reader.decoder())
                                {
                                    builder.style_type =
                                        WdStyleType::from_xml(&value).unwrap_or_default();
                                }
                            },
                            b"styleId" => {
                                if let Ok(value) = attr.decode_and_unescape_value(reader.decoder())
                                {
                                    builder.style_id = Some(value.to_string());
                                }
                            },
                            b"default" => {
                                if let Ok(value) = attr.decode_and_unescape_value(reader.decoder())
                                {
                                    builder.is_default = value == "1" || value == "true";
                                }
                            },
                            b"customStyle" => {
                                if let Ok(value) = attr.decode_and_unescape_value(reader.decoder())
                                {
                                    builder.is_custom = value == "1" || value == "true";
                                }
                            },
                            _ => {},
                        }
                    }

                    current_style = Some(builder);
                },
                Ok(Event::Empty(e)) if current_style.is_some() => {
                    let builder = current_style.as_mut().unwrap();
                    match e.local_name().as_ref() {
                        b"name" => {
                            // Parse name attribute
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"val"
                                    && let Ok(value) =
                                        attr.decode_and_unescape_value(reader.decoder())
                                {
                                    builder.name = Some(value.to_string());
                                }
                            }
                        },
                        b"basedOn" => {
                            // Parse basedOn attribute
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"val"
                                    && let Ok(value) =
                                        attr.decode_and_unescape_value(reader.decoder())
                                {
                                    builder.based_on = Some(value.to_string());
                                }
                            }
                        },
                        b"uiPriority" => {
                            // Parse UI priority
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"val"
                                    && let Ok(value) =
                                        attr.decode_and_unescape_value(reader.decoder())
                                    && let Ok(priority) = value.parse::<i32>()
                                {
                                    builder.priority = Some(priority);
                                }
                            }
                        },
                        b"qFormat" => {
                            builder.is_quick_style = true;
                        },
                        b"semiHidden" => {
                            builder.is_hidden = true;
                        },
                        b"locked" => {
                            builder.is_locked = true;
                        },
                        _ => {},
                    }
                },
                Ok(Event::End(e)) if e.local_name().as_ref() == b"style" => {
                    // Finish current style
                    if let Some(builder) = current_style.take()
                        && let Some(style_id) = builder.style_id
                    {
                        styles.push(Style {
                            style_id,
                            name: builder.name,
                            style_type: builder.style_type,
                            is_default: builder.is_default,
                            is_custom: builder.is_custom,
                            based_on: builder.based_on,
                            priority: builder.priority,
                            is_quick_style: builder.is_quick_style,
                            is_hidden: builder.is_hidden,
                            is_locked: builder.is_locked,
                        });
                    }
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
        }

        self.style_list = Some(styles);
        Ok(())
    }
}

/// Builder for constructing Style objects during XML parsing.
#[derive(Debug, Default)]
struct StyleBuilder {
    style_id: Option<String>,
    name: Option<String>,
    style_type: WdStyleType,
    is_default: bool,
    is_custom: bool,
    based_on: Option<String>,
    priority: Option<i32>,
    is_quick_style: bool,
    is_hidden: bool,
    is_locked: bool,
}

/// A single style definition in a Word document.
///
/// Represents a `<w:style>` element with its properties.
/// Can be a paragraph, character, table, or list style.
#[derive(Debug, Clone)]
pub struct Style {
    /// Style identifier (required)
    style_id: String,
    /// UI-visible name
    name: Option<String>,
    /// Type of style (paragraph, character, table, or list)
    style_type: WdStyleType,
    /// Whether this is the default style for its type
    is_default: bool,
    /// Whether this is a custom (user-defined) style
    is_custom: bool,
    /// ID of the style this is based on
    based_on: Option<String>,
    /// UI priority for display ordering
    priority: Option<i32>,
    /// Whether to show in quick style gallery
    is_quick_style: bool,
    /// Whether hidden from UI
    is_hidden: bool,
    /// Whether locked (formatting protection)
    is_locked: bool,
}

impl Style {
    /// Get the style identifier.
    #[inline]
    pub fn style_id(&self) -> &str {
        &self.style_id
    }

    /// Get the style name.
    ///
    /// Returns `None` if no name is defined.
    #[inline]
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    /// Get the style type.
    #[inline]
    pub fn style_type(&self) -> WdStyleType {
        self.style_type
    }

    /// Check if this is the default style for its type.
    #[inline]
    pub fn is_default(&self) -> bool {
        self.is_default
    }

    /// Check if this is a built-in style.
    ///
    /// Returns `true` if this is a built-in Word style, `false` for custom styles.
    #[inline]
    pub fn is_builtin(&self) -> bool {
        !self.is_custom
    }

    /// Check if this is a custom (user-defined) style.
    #[inline]
    pub fn is_custom(&self) -> bool {
        self.is_custom
    }

    /// Get the ID of the style this is based on.
    #[inline]
    pub fn based_on(&self) -> Option<&str> {
        self.based_on.as_deref()
    }

    /// Get the UI priority for this style.
    ///
    /// Lower values appear first in style lists.
    #[inline]
    pub fn priority(&self) -> Option<i32> {
        self.priority
    }

    /// Check if this style appears in the quick style gallery.
    #[inline]
    pub fn is_quick_style(&self) -> bool {
        self.is_quick_style
    }

    /// Check if this style is hidden from the UI.
    #[inline]
    pub fn is_hidden(&self) -> bool {
        self.is_hidden
    }

    /// Check if this style is locked.
    ///
    /// Locked styles cannot be applied when formatting protection is enabled.
    #[inline]
    pub fn is_locked(&self) -> bool {
        self.is_locked
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_style_type_default() {
        let style_type = WdStyleType::default();
        assert_eq!(style_type, WdStyleType::Paragraph);
    }
}
