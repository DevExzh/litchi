//! Document settings and protection support.
//!
//! This module provides types and methods for accessing document settings
//! and protection status.

use crate::docx::namespace::{is_wordprocessing_namespace, word_attribute_value};
use crate::error::{OoxmlError, Result};
use litchi_opc::part::Part;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;

/// Document settings including protection status.
///
/// # Examples
///
/// ```rust,no_run
/// use litchi_ooxml::docx::Package;
///
/// let pkg = Package::open("document.docx")?;
/// let doc = pkg.document()?;
///
/// if let Some(settings) = doc.settings()? {
///     if settings.is_protected() {
///         println!("Document is protected");
///         println!("Protection type: {:?}", settings.protection_type());
///     }
/// }
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
#[derive(Debug, Clone)]
pub struct DocumentSettings {
    /// Whether document is protected
    protected: bool,
    /// Type of protection
    protection_type: Option<ProtectionType>,
    /// Whether to track revisions
    track_revisions: bool,
    /// Zoom percentage
    zoom_percent: Option<u32>,
    /// Smart-tag type declarations.
    smart_tag_types: Vec<SmartTagType>,
    /// Whether applications should omit embedded smart-tag data when saving.
    do_not_embed_smart_tags: bool,
}

/// A smart-tag vocabulary declaration from `settings.xml`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SmartTagType {
    namespace_uri: String,
    name: String,
    url: String,
}

impl SmartTagType {
    /// Return the smart-tag vocabulary namespace URI.
    #[inline]
    pub fn namespace_uri(&self) -> &str {
        &self.namespace_uri
    }

    /// Return the smart-tag type name.
    #[inline]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Return the vocabulary download URL.
    #[inline]
    pub fn url(&self) -> &str {
        &self.url
    }
}

/// Type of document protection.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ProtectionType {
    /// No editing allowed
    ReadOnly,
    /// Only comments allowed
    Comments,
    /// Only tracked changes allowed
    TrackedChanges,
    /// Only form fields allowed
    Forms,
}

impl ProtectionType {
    /// Parse protection type from XML value.
    fn from_xml(s: &str) -> Option<Self> {
        match s {
            "readOnly" => Some(Self::ReadOnly),
            "comments" => Some(Self::Comments),
            "trackedChanges" => Some(Self::TrackedChanges),
            "forms" => Some(Self::Forms),
            _ => None,
        }
    }

    /// Get XML value for this protection type.
    pub const fn to_xml(self) -> &'static str {
        match self {
            Self::ReadOnly => "readOnly",
            Self::Comments => "comments",
            Self::TrackedChanges => "trackedChanges",
            Self::Forms => "forms",
        }
    }
}

impl DocumentSettings {
    /// Create a new DocumentSettings with default values.
    pub fn new() -> Self {
        Self {
            protected: false,
            protection_type: None,
            track_revisions: false,
            zoom_percent: None,
            smart_tag_types: Vec::new(),
            do_not_embed_smart_tags: false,
        }
    }

    /// Check if the document is protected.
    #[inline]
    pub fn is_protected(&self) -> bool {
        self.protected
    }

    /// Get the type of protection applied.
    #[inline]
    pub fn protection_type(&self) -> Option<ProtectionType> {
        self.protection_type
    }

    /// Check if track revisions is enabled.
    #[inline]
    pub fn track_revisions(&self) -> bool {
        self.track_revisions
    }

    /// Get the zoom percentage.
    #[inline]
    pub fn zoom_percent(&self) -> Option<u32> {
        self.zoom_percent
    }

    /// Return the declared smart-tag vocabularies in document order.
    #[inline]
    pub fn smart_tag_types(&self) -> &[SmartTagType] {
        &self.smart_tag_types
    }

    /// Whether embedded smart-tag data should be omitted when saving.
    #[inline]
    pub fn do_not_embed_smart_tags(&self) -> bool {
        self.do_not_embed_smart_tags
    }

    /// Extract settings from a settings.xml part.
    ///
    /// # Arguments
    ///
    /// * `part` - The settings part
    ///
    /// # Returns
    ///
    /// A DocumentSettings object
    pub(crate) fn extract_from_part(part: &dyn Part) -> Result<Self> {
        let xml = crate::common::mce::process_part(part)?;
        Self::extract_from_xml(xml.as_ref())
    }

    fn extract_from_xml(xml_bytes: &[u8]) -> Result<Self> {
        let mut reader = NsReader::from_reader(xml_bytes);

        let mut settings = Self::new();
        let mut depth = 0usize;
        let mut saw_root = false;
        let mut saw_do_not_embed_smart_tags = false;

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
                        OoxmlError::InvalidFormat("Word settings XML nesting is too deep".into())
                    })?;
                    if depth == 1 {
                        validate_settings_root(&namespace, &element, saw_root)?;
                        saw_root = true;
                    } else if depth == 2 && saw_root && is_wordprocessing_namespace(&namespace) {
                        parse_setting(
                            &element,
                            decoder,
                            &resolver,
                            &mut settings,
                            &mut saw_do_not_embed_smart_tags,
                        )?;
                    }
                },
                Event::Empty(element) => {
                    let child_depth = depth.checked_add(1).ok_or_else(|| {
                        OoxmlError::InvalidFormat("Word settings XML nesting is too deep".into())
                    })?;
                    if child_depth == 1 {
                        validate_settings_root(&namespace, &element, saw_root)?;
                        saw_root = true;
                    } else if child_depth == 2
                        && saw_root
                        && is_wordprocessing_namespace(&namespace)
                    {
                        parse_setting(
                            &element,
                            decoder,
                            &resolver,
                            &mut settings,
                            &mut saw_do_not_embed_smart_tags,
                        )?;
                    }
                },
                Event::End(_) => {
                    depth = depth.checked_sub(1).ok_or_else(|| {
                        OoxmlError::InvalidFormat("invalid Word settings XML nesting".into())
                    })?;
                },
                Event::Eof if depth != 0 => {
                    return Err(OoxmlError::InvalidFormat(
                        "unterminated Word settings XML".into(),
                    ));
                },
                Event::Eof => break,
                _ => {},
            }
        }

        if !saw_root {
            return Err(OoxmlError::InvalidFormat(
                "settings part has no settings root".into(),
            ));
        }
        Ok(settings)
    }
}

fn validate_settings_root(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    saw_root: bool,
) -> Result<()> {
    if saw_root
        || !is_wordprocessing_namespace(namespace)
        || element.local_name().as_ref() != b"settings"
    {
        return Err(OoxmlError::InvalidFormat(
            "settings part has an invalid or trailing root element".into(),
        ));
    }
    Ok(())
}

fn parse_setting(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    settings: &mut DocumentSettings,
    saw_do_not_embed_smart_tags: &mut bool,
) -> Result<()> {
    match element.local_name().as_ref() {
        b"documentProtection" => {
            settings.protected = true;
            if let Some(value) = word_attribute_value(element, b"edit", decoder, resolver)? {
                settings.protection_type = ProtectionType::from_xml(&value);
            }
            if let Some(value) = word_attribute_value(element, b"enforcement", decoder, resolver)? {
                settings.protected = parse_on_off_value(&value)?;
            }
        },
        b"trackRevisions" => {
            settings.track_revisions = parse_on_off(element, decoder, resolver)?;
        },
        b"zoom" => {
            if let Some(value) = word_attribute_value(element, b"percent", decoder, resolver)? {
                settings.zoom_percent =
                    atoi_simd::parse::<u32, false, false>(value.as_bytes()).ok();
            }
        },
        b"smartTagType" => {
            settings.smart_tag_types.push(SmartTagType {
                namespace_uri: required_attribute(
                    element,
                    b"namespaceuri",
                    decoder,
                    resolver,
                    "smart-tag namespace URI",
                )?,
                name: required_attribute(
                    element,
                    b"name",
                    decoder,
                    resolver,
                    "smart-tag type name",
                )?,
                url: required_attribute(
                    element,
                    b"url",
                    decoder,
                    resolver,
                    "smart-tag vocabulary URL",
                )?,
            });
        },
        b"doNotEmbedSmartTags" => {
            if std::mem::replace(saw_do_not_embed_smart_tags, true) {
                return Err(OoxmlError::InvalidFormat(
                    "duplicate doNotEmbedSmartTags setting".into(),
                ));
            }
            settings.do_not_embed_smart_tags = parse_on_off(element, decoder, resolver)?;
        },
        _ => {},
    }
    Ok(())
}

fn required_attribute(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    resolver: &NamespaceResolver,
    description: &str,
) -> Result<String> {
    word_attribute_value(element, name, decoder, resolver)?.ok_or_else(|| {
        OoxmlError::InvalidFormat(format!("Word {description} attribute is required"))
    })
}

fn parse_on_off(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<bool> {
    word_attribute_value(element, b"val", decoder, resolver)?
        .as_deref()
        .map_or(Ok(true), parse_on_off_value)
}

fn parse_on_off_value(value: &str) -> Result<bool> {
    match value {
        "true" | "1" | "on" => Ok(true),
        "false" | "0" | "off" => Ok(false),
        _ => Err(OoxmlError::InvalidFormat(format!(
            "invalid Word on/off value '{value}'"
        ))),
    }
}

impl Default for DocumentSettings {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_settings_creation() {
        let settings = DocumentSettings::new();
        assert!(!settings.is_protected());
        assert!(settings.protection_type().is_none());
        assert!(!settings.track_revisions());
    }

    #[test]
    fn test_protection_type() {
        assert_eq!(
            ProtectionType::from_xml("readOnly"),
            Some(ProtectionType::ReadOnly)
        );
        assert_eq!(
            ProtectionType::from_xml("comments"),
            Some(ProtectionType::Comments)
        );
        assert_eq!(
            ProtectionType::from_xml("trackedChanges"),
            Some(ProtectionType::TrackedChanges)
        );
        assert_eq!(
            ProtectionType::from_xml("forms"),
            Some(ProtectionType::Forms)
        );
        assert_eq!(ProtectionType::from_xml("invalid"), None);
    }

    #[test]
    fn parses_smart_tag_settings_with_strict_namespaces() {
        let xml = br#"<s:settings xmlns:s="http://purl.oclc.org/ooxml/wordprocessingml/main"
            xmlns:false="urn:not-wordprocessingml">
            <s:trackRevisions s:val="on"/>
            <s:smartTagType s:namespaceuri="urn:contacts" s:name="person"
                s:url="https://example.test/schema?a=1&amp;b=2"/>
            <false:smartTagType false:namespaceuri="urn:false" false:name="ignored" false:url="ignored"/>
            <s:doNotEmbedSmartTags s:val="off"/>
        </s:settings>"#;

        let settings = DocumentSettings::extract_from_xml(xml).unwrap();
        assert!(settings.track_revisions());
        assert!(!settings.do_not_embed_smart_tags());
        assert_eq!(settings.smart_tag_types().len(), 1);
        assert_eq!(
            settings.smart_tag_types()[0].namespace_uri(),
            "urn:contacts"
        );
        assert_eq!(settings.smart_tag_types()[0].name(), "person");
        assert_eq!(
            settings.smart_tag_types()[0].url(),
            "https://example.test/schema?a=1&b=2"
        );
    }

    #[test]
    fn validates_smart_tag_settings() {
        let enabled = br#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:doNotEmbedSmartTags/></w:settings>"#;
        assert!(
            DocumentSettings::extract_from_xml(enabled)
                .unwrap()
                .do_not_embed_smart_tags()
        );

        let missing_url = br#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:smartTagType w:namespaceuri="urn:test" w:name="test"/></w:settings>"#;
        assert!(DocumentSettings::extract_from_xml(missing_url).is_err());

        let invalid_on_off = br#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:doNotEmbedSmartTags w:val="maybe"/></w:settings>"#;
        assert!(DocumentSettings::extract_from_xml(invalid_on_off).is_err());

        let duplicate = br#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:doNotEmbedSmartTags/><w:doNotEmbedSmartTags/></w:settings>"#;
        assert!(DocumentSettings::extract_from_xml(duplicate).is_err());
    }
}
