/// Styles - document styles and formatting definitions.
use crate::docx::enums::WdStyleType;
use crate::docx::numbering::ParagraphNumbering;
use crate::error::{OoxmlError, Result};
use litchi_opc::part::Part;
use quick_xml::events::Event;
use quick_xml::{Reader, XmlVersion};
use smallvec::SmallVec;
use std::collections::HashSet;

/// A Word paragraph outline level tied to a built-in heading rank.
///
/// Word stores Heading 1 as wire level `0` and Heading 9 as wire level `8`.
/// The enum keeps every public value inside that closed domain.
#[repr(u8)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub enum Outline {
    H1 = 0,
    H2 = 1,
    H3 = 2,
    H4 = 3,
    H5 = 4,
    H6 = 5,
    H7 = 6,
    H8 = 7,
    H9 = 8,
}

impl Outline {
    /// Convert a wire level into its typed heading rank.
    pub const fn new(level: u8) -> Option<Self> {
        match level {
            0 => Some(Self::H1),
            1 => Some(Self::H2),
            2 => Some(Self::H3),
            3 => Some(Self::H4),
            4 => Some(Self::H5),
            5 => Some(Self::H6),
            6 => Some(Self::H7),
            7 => Some(Self::H8),
            8 => Some(Self::H9),
            _ => None,
        }
    }

    /// Return the zero-based WordprocessingML wire level.
    pub const fn level(self) -> u8 {
        self as u8
    }
}

/// A collection of styles defined in a Word document.
///
/// Provides access to paragraph, character, table, and list styles.
/// Supports iteration and lookup by style ID or name.
///
/// # Examples
///
/// ```rust,ignore
/// use litchi_ooxml::docx::Package;
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

    /// Resolve inherited paragraph numbering with a bounded, cycle-checked `basedOn` walk.
    pub fn resolved_numbering(&mut self, style_id: &str) -> Result<Option<ParagraphNumbering>> {
        self.ensure_styles_loaded()?;
        let styles = self.style_list.as_ref().expect("styles loaded");
        let mut current = style_id;
        let mut visited = HashSet::new();
        loop {
            if !visited.insert(current.to_owned()) {
                return Err(OoxmlError::InvalidFormat(format!(
                    "style basedOn cycle at '{current}'"
                )));
            }
            if visited.len() > styles.len().saturating_add(1) {
                return Err(OoxmlError::InvalidFormat(
                    "style inheritance exceeds the style table".to_owned(),
                ));
            }
            let Some(style) = styles.iter().find(|style| style.style_id == current) else {
                return Ok(None);
            };
            if let Some(numbering) = style.numbering {
                return Ok(Some(numbering));
            }
            let Some(parent) = style.based_on.as_deref() else {
                return Ok(None);
            };
            current = parent;
        }
    }

    /// Ensure styles are loaded from XML.
    fn ensure_styles_loaded(&mut self) -> Result<()> {
        if self.style_list.is_some() {
            return Ok(());
        }

        let xml_bytes = litchi_ooxml_common::mce::process_part(self.part)?;
        let mut reader = Reader::from_reader(xml_bytes.as_ref());
        reader.config_mut().trim_text(true);

        let mut styles = SmallVec::new();
        let mut current_style: Option<StyleBuilder> = None;
        let mut in_ppr = false;
        let mut in_num_pr = false;
        let mut pending_num_id = None;
        let mut pending_level = None;

        loop {
            match reader.read_event() {
                Ok(Event::Start(e)) if e.local_name().as_ref() == b"style" => {
                    // Start a new style
                    let mut builder = StyleBuilder::default();

                    // Parse attributes
                    for attr in e.attributes().flatten() {
                        match attr.key.local_name().as_ref() {
                            b"type" => {
                                if let Ok(value) = attr.decoded_and_normalized_value(
                                    XmlVersion::Implicit1_0,
                                    reader.decoder(),
                                ) {
                                    builder.style_type =
                                        WdStyleType::from_xml(&value).unwrap_or_default();
                                }
                            },
                            b"styleId" => {
                                if let Ok(value) = attr.decoded_and_normalized_value(
                                    XmlVersion::Implicit1_0,
                                    reader.decoder(),
                                ) {
                                    builder.style_id = Some(value.to_string());
                                }
                            },
                            b"default" => {
                                if let Ok(value) = attr.decoded_and_normalized_value(
                                    XmlVersion::Implicit1_0,
                                    reader.decoder(),
                                ) {
                                    builder.is_default = value == "1" || value == "true";
                                }
                            },
                            b"customStyle" => {
                                if let Ok(value) = attr.decoded_and_normalized_value(
                                    XmlVersion::Implicit1_0,
                                    reader.decoder(),
                                ) {
                                    builder.is_custom = value == "1" || value == "true";
                                }
                            },
                            _ => {},
                        }
                    }

                    current_style = Some(builder);
                },
                Ok(Event::Start(e))
                    if current_style.is_some() && e.local_name().as_ref() == b"pPr" =>
                {
                    in_ppr = true
                },
                Ok(Event::Start(e))
                    if current_style.is_some() && in_ppr && e.local_name().as_ref() == b"numPr" =>
                {
                    if in_num_pr {
                        return Err(OoxmlError::InvalidFormat(
                            "style has nested numPr".to_owned(),
                        ));
                    }
                    in_num_pr = true;
                    pending_num_id = None;
                    pending_level = None;
                },
                Ok(Event::Empty(e)) if current_style.is_some() => {
                    let builder = current_style.as_mut().unwrap();
                    match e.local_name().as_ref() {
                        b"numId" if in_num_pr => {
                            if pending_num_id.is_some() {
                                return Err(OoxmlError::InvalidFormat(
                                    "style has duplicate numId".to_owned(),
                                ));
                            }
                            let raw = required_style_value(&e, reader.decoder())?;
                            pending_num_id = Some(raw.parse::<u32>().map_err(|_| {
                                OoxmlError::InvalidFormat(format!("invalid style numId '{raw}'"))
                            })?);
                        },
                        b"ilvl" if in_num_pr => {
                            if pending_level.is_some() {
                                return Err(OoxmlError::InvalidFormat(
                                    "style has duplicate ilvl".to_owned(),
                                ));
                            }
                            let raw = required_style_value(&e, reader.decoder())?;
                            pending_level = Some(
                                raw.parse::<u8>()
                                    .ok()
                                    .filter(|value| *value <= 8)
                                    .ok_or_else(|| {
                                        OoxmlError::InvalidFormat(format!(
                                            "invalid style ilvl '{raw}'"
                                        ))
                                    })?,
                            );
                        },
                        b"outlineLvl" if in_ppr => {
                            if builder.outline.is_some() {
                                return Err(OoxmlError::InvalidFormat(
                                    "style has duplicate outlineLvl".to_owned(),
                                ));
                            }
                            let raw = required_style_value(&e, reader.decoder())?;
                            let level = raw.parse::<u8>().map_err(|_| {
                                OoxmlError::InvalidFormat(format!(
                                    "invalid style outlineLvl '{raw}'"
                                ))
                            })?;
                            builder.outline = Some(Outline::new(level).ok_or_else(|| {
                                OoxmlError::InvalidFormat(format!(
                                    "invalid style outlineLvl '{raw}'"
                                ))
                            })?);
                        },
                        b"name" => {
                            // Parse name attribute
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"val"
                                    && let Ok(value) = attr.decoded_and_normalized_value(
                                        XmlVersion::Implicit1_0,
                                        reader.decoder(),
                                    )
                                {
                                    builder.name = Some(value.to_string());
                                }
                            }
                        },
                        b"basedOn" => {
                            // Parse basedOn attribute
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"val"
                                    && let Ok(value) = attr.decoded_and_normalized_value(
                                        XmlVersion::Implicit1_0,
                                        reader.decoder(),
                                    )
                                {
                                    builder.based_on = Some(value.to_string());
                                }
                            }
                        },
                        b"uiPriority" => {
                            // Parse UI priority
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"val"
                                    && let Ok(value) = attr.decoded_and_normalized_value(
                                        XmlVersion::Implicit1_0,
                                        reader.decoder(),
                                    )
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
                Ok(Event::End(e)) if e.local_name().as_ref() == b"numPr" && in_num_pr => {
                    let num_id = pending_num_id.take().ok_or_else(|| {
                        OoxmlError::InvalidFormat("style numPr is missing numId".to_owned())
                    })?;
                    current_style.as_mut().expect("style checked").numbering =
                        Some(ParagraphNumbering {
                            num_id,
                            level: pending_level.take().unwrap_or(0),
                        });
                    in_num_pr = false;
                },
                Ok(Event::End(e)) if e.local_name().as_ref() == b"pPr" && in_ppr => {
                    if in_num_pr {
                        return Err(OoxmlError::InvalidFormat(
                            "unterminated style numPr".to_owned(),
                        ));
                    }
                    in_ppr = false;
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
                            numbering: builder.numbering,
                            outline: builder.outline,
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
    numbering: Option<ParagraphNumbering>,
    outline: Option<Outline>,
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
    numbering: Option<ParagraphNumbering>,
    outline: Option<Outline>,
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

    /// Direct paragraph numbering declared by this style.
    #[inline]
    pub fn numbering(&self) -> Option<ParagraphNumbering> {
        self.numbering
    }

    /// Direct outline level declared by this paragraph style.
    #[inline]
    pub const fn outline(&self) -> Option<Outline> {
        self.outline
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

fn required_style_value(
    element: &quick_xml::events::BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
) -> Result<String> {
    for attribute in element.attributes().with_checks(false) {
        let attribute = attribute.map_err(|error| OoxmlError::Xml(error.to_string()))?;
        if attribute.key.local_name().as_ref() == b"val" {
            return attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                .map(|value| value.into_owned())
                .map_err(|error| OoxmlError::Xml(error.to_string()));
        }
    }
    Err(OoxmlError::InvalidFormat(
        "style property is missing val".to_owned(),
    ))
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::PackURI;
    use litchi_opc::part::BlobPart;

    #[test]
    fn test_style_type_default() {
        let style_type = WdStyleType::default();
        assert_eq!(style_type, WdStyleType::Paragraph);
    }

    fn with_styles<T>(xml: &[u8], inspect: impl FnOnce(&mut Styles<'_>) -> T) -> T {
        let part = BlobPart::new(
            PackURI::new("/word/styles.xml").unwrap(),
            "application/xml".to_owned(),
            xml.to_vec(),
        );
        let mut styles = Styles::from_part(&part);
        inspect(&mut styles)
    }

    #[test]
    fn resolves_inherited_numbering_and_retains_cancellation() {
        with_styles(br#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:styleId="Base"><w:pPr><w:numPr><w:ilvl w:val="2"/><w:numId w:val="7"/></w:numPr></w:pPr></w:style><w:style w:type="paragraph" w:styleId="Child"><w:basedOn w:val="Base"/></w:style><w:style w:type="paragraph" w:styleId="Cancel"><w:pPr><w:numPr><w:numId w:val="0"/></w:numPr></w:pPr></w:style></w:styles>"#, |value| {
            assert_eq!(
                value.resolved_numbering("Child").unwrap(),
                Some(ParagraphNumbering { num_id: 7, level: 2 })
            );
            assert_eq!(
                value.resolved_numbering("Cancel").unwrap(),
                Some(ParagraphNumbering { num_id: 0, level: 0 })
            );
        });
    }

    #[test]
    fn rejects_based_on_cycles_and_malformed_num_pr() {
        with_styles(br#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:styleId="A"><w:basedOn w:val="B"/></w:style><w:style w:type="paragraph" w:styleId="B"><w:basedOn w:val="A"/></w:style></w:styles>"#, |cycle| {
            assert!(cycle.resolved_numbering("A").is_err());
        });
        with_styles(br#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:styleId="A"><w:pPr><w:numPr><w:ilvl w:val="9"/></w:numPr></w:pPr></w:style></w:styles>"#, |malformed| {
            assert!(malformed.len().is_err());
        });
    }

    #[test]
    fn outline_levels_are_typed_and_malformed_values_are_rejected() {
        assert_eq!(Outline::new(0), Some(Outline::H1));
        assert_eq!(Outline::new(8), Some(Outline::H9));
        assert_eq!(Outline::new(9), None);

        with_styles(br#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:styleId="Heading1"><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style></w:styles>"#, |styles| {
            assert_eq!(
                styles.get_by_id("Heading1").unwrap().unwrap().outline(),
                Some(Outline::H1)
            );
        });
        with_styles(br#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:styleId="Bad"><w:pPr><w:outlineLvl w:val="9"/></w:pPr></w:style></w:styles>"#, |styles| {
            assert!(styles.len().is_err());
        });
    }
}
