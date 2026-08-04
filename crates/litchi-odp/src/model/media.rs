//! Inert multimedia references used by ODF presentation frames.

use litchi_core::{Error, Result, xml::escape_xml};
use std::collections::BTreeMap;

/// XLink `show` behavior stored on a presentation media plugin.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Show {
    /// Open in a new presentation context.
    New,
    /// Replace the current context.
    Replace,
    /// Embed the target in the current context.
    Embed,
    /// Application-defined behavior.
    Other,
    /// No requested show behavior.
    None,
}

impl Show {
    pub(crate) fn parse(value: &str) -> Result<Self> {
        match value {
            "new" => Ok(Self::New),
            "replace" => Ok(Self::Replace),
            "embed" => Ok(Self::Embed),
            "other" => Ok(Self::Other),
            "none" => Ok(Self::None),
            _ => Err(Error::InvalidFormat(format!(
                "invalid draw:plugin xlink:show value '{value}'"
            ))),
        }
    }

    const fn as_str(self) -> &'static str {
        match self {
            Self::New => "new",
            Self::Replace => "replace",
            Self::Embed => "embed",
            Self::Other => "other",
            Self::None => "none",
        }
    }
}

/// XLink `actuate` behavior stored on a presentation media plugin.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Actuate {
    /// Load the media with its containing document.
    OnLoad,
    /// Load the media only when requested.
    OnRequest,
    /// Application-defined behavior.
    Other,
    /// No requested activation behavior.
    None,
}

impl Actuate {
    pub(crate) fn parse(value: &str) -> Result<Self> {
        match value {
            "onLoad" => Ok(Self::OnLoad),
            "onRequest" => Ok(Self::OnRequest),
            "other" => Ok(Self::Other),
            "none" => Ok(Self::None),
            _ => Err(Error::InvalidFormat(format!(
                "invalid draw:plugin xlink:actuate value '{value}'"
            ))),
        }
    }

    const fn as_str(self) -> &'static str {
        match self {
            Self::OnLoad => "onLoad",
            Self::OnRequest => "onRequest",
            Self::Other => "other",
            Self::None => "none",
        }
    }
}

/// An inert name/value parameter belonging to a media plugin.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Parameter {
    name: String,
    value: String,
}

impl Parameter {
    /// Create a plugin parameter.
    pub fn new(name: impl Into<String>, value: impl Into<String>) -> Result<Self> {
        let name = name.into();
        let value = value.into();
        validate_bounded_xml_value(&name, "media parameter name")?;
        validate_bounded_xml_value(&value, "media parameter value")?;
        Ok(Self { name, value })
    }

    /// Return the parameter name.
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Return the parameter value.
    pub fn value(&self) -> &str {
        &self.value
    }
}

/// An inert `draw:plugin` media reference.
///
/// This type records package-local or external audio/video links. Litchi does
/// not load external URLs, play media, or interpret plugin parameters.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Reference {
    href: String,
    mime_type: Option<String>,
    show: Option<Show>,
    actuate: Option<Actuate>,
    xml_id: Option<String>,
    parameters: Vec<Parameter>,
}

impl Reference {
    /// Create the minimal schema-valid plugin reference.
    ///
    /// The serialized XLink type is always `simple`, as required by ODF.
    pub fn new(href: impl Into<String>) -> Result<Self> {
        let href = href.into();
        validate_href(&href)?;
        Ok(Self {
            href,
            mime_type: None,
            show: None,
            actuate: None,
            xml_id: None,
            parameters: Vec::new(),
        })
    }

    /// Return the unescaped XLink target.
    pub fn href(&self) -> &str {
        &self.href
    }

    /// Return a safe package-relative path, or `None` for an external/fragment link.
    pub fn package_path(&self) -> Option<&str> {
        let path = self.href.strip_prefix("./").unwrap_or(&self.href);
        let first_component = path.split('/').next()?;
        if path.is_empty()
            || path.starts_with('/')
            || path.ends_with('/')
            || path.contains('\\')
            || path.contains('?')
            || path.contains('#')
            || first_component.contains(':')
            || path
                .split('/')
                .any(|component| component.is_empty() || component == "." || component == "..")
        {
            None
        } else {
            Some(path)
        }
    }

    /// Replace the XLink target.
    pub fn set_href(&mut self, href: impl Into<String>) -> Result<()> {
        let href = href.into();
        validate_href(&href)?;
        self.href = href;
        Ok(())
    }

    /// Return the optional plugin MIME type.
    pub fn mime_type(&self) -> Option<&str> {
        self.mime_type.as_deref()
    }

    /// Set the plugin MIME type.
    pub fn set_mime_type(&mut self, mime_type: impl Into<String>) -> Result<()> {
        let mime_type = mime_type.into();
        validate_media_type(&mime_type)?;
        self.mime_type = Some(mime_type);
        Ok(())
    }

    /// Remove the plugin MIME type.
    pub fn clear_mime_type(&mut self) {
        self.mime_type = None;
    }

    /// Return the optional XLink show behavior.
    pub fn show(&self) -> Option<Show> {
        self.show
    }

    /// Set or remove the XLink show behavior.
    pub fn set_show(&mut self, show: Option<Show>) {
        self.show = show;
    }

    /// Return the optional XLink activation behavior.
    pub fn actuate(&self) -> Option<Actuate> {
        self.actuate
    }

    /// Set or remove the XLink activation behavior.
    pub fn set_actuate(&mut self, actuate: Option<Actuate>) {
        self.actuate = actuate;
    }

    /// Return the optional XML ID.
    pub fn xml_id(&self) -> Option<&str> {
        self.xml_id.as_deref()
    }

    /// Set the XML ID.
    pub fn set_xml_id(&mut self, xml_id: impl Into<String>) -> Result<()> {
        let xml_id = xml_id.into();
        validate_ncname(&xml_id, "media XML ID")?;
        self.xml_id = Some(xml_id);
        Ok(())
    }

    /// Remove the XML ID.
    pub fn clear_xml_id(&mut self) {
        self.xml_id = None;
    }

    /// Return the inert plugin parameters.
    pub fn parameters(&self) -> &[Parameter] {
        &self.parameters
    }

    /// Return mutable plugin parameters.
    pub fn parameters_mut(&mut self) -> &mut Vec<Parameter> {
        &mut self.parameters
    }

    /// Add an inert plugin parameter.
    pub fn add_parameter(&mut self, parameter: Parameter) -> Result<()> {
        if self.parameters.len() >= 1024 {
            return Err(Error::InvalidFormat(
                "ODP media plugin exceeds 1024 parameters".to_string(),
            ));
        }
        self.parameters.push(parameter);
        Ok(())
    }

    pub(crate) fn validate(&self) -> Result<()> {
        validate_href(&self.href)?;
        if let Some(mime_type) = &self.mime_type {
            validate_media_type(mime_type)?;
        }
        if let Some(xml_id) = &self.xml_id {
            validate_ncname(xml_id, "media XML ID")?;
        }
        if self.parameters.len() > 1024 {
            return Err(Error::InvalidFormat(
                "ODP media plugin exceeds 1024 parameters".to_string(),
            ));
        }
        for parameter in &self.parameters {
            validate_bounded_xml_value(&parameter.name, "media parameter name")?;
            validate_bounded_xml_value(&parameter.value, "media parameter value")?;
        }
        Ok(())
    }

    pub(crate) fn write_xml(&self, output: &mut String) -> Result<()> {
        self.validate()?;
        output.push_str("<draw:plugin xlink:href=\"");
        output.push_str(&escape_xml(&self.href));
        output.push_str("\" xlink:type=\"simple\"");
        if let Some(mime_type) = &self.mime_type {
            output.push_str(" draw:mime-type=\"");
            output.push_str(&escape_xml(mime_type));
            output.push('"');
        }
        if let Some(show) = self.show {
            output.push_str(" xlink:show=\"");
            output.push_str(show.as_str());
            output.push('"');
        }
        if let Some(actuate) = self.actuate {
            output.push_str(" xlink:actuate=\"");
            output.push_str(actuate.as_str());
            output.push('"');
        }
        if let Some(xml_id) = &self.xml_id {
            output.push_str(" xml:id=\"");
            output.push_str(xml_id);
            output.push('"');
        }
        if self.parameters.is_empty() {
            output.push_str("/>");
            return Ok(());
        }
        output.push('>');
        for parameter in &self.parameters {
            output.push_str("<draw:param draw:name=\"");
            output.push_str(&escape_xml(&parameter.name));
            output.push_str("\" draw:value=\"");
            output.push_str(&escape_xml(&parameter.value));
            output.push_str("\"/>");
        }
        output.push_str("</draw:plugin>");
        Ok(())
    }
}

#[derive(Debug, Clone)]
pub(crate) struct EmbeddedMedia {
    pub(crate) bytes: Vec<u8>,
    pub(crate) media_type: String,
}

pub(crate) fn embed_media(
    files: &mut BTreeMap<String, EmbeddedMedia>,
    path: impl Into<String>,
    bytes: impl Into<Vec<u8>>,
    media_type: impl Into<String>,
) -> Result<Reference> {
    let path = path.into();
    let media_type = media_type.into();
    validate_package_media_path(&path)?;
    validate_media_type(&media_type)?;
    if files.len() >= 65_536 {
        return Err(Error::InvalidFormat(
            "ODP package exceeds 65536 newly embedded media files".to_string(),
        ));
    }
    if files.contains_key(&path) {
        return Err(Error::InvalidFormat(format!(
            "duplicate embedded ODP media path '{path}'"
        )));
    }
    let mut reference = Reference::new(path.clone())?;
    reference.set_mime_type(media_type.clone())?;
    reference.set_show(Some(Show::Embed));
    reference.set_actuate(Some(Actuate::OnLoad));
    files.insert(
        path,
        EmbeddedMedia {
            bytes: bytes.into(),
            media_type,
        },
    );
    Ok(reference)
}

pub(crate) fn validate_package_media_path(path: &str) -> Result<()> {
    if path.is_empty()
        || path.starts_with('/')
        || path.ends_with('/')
        || path.contains('\\')
        || path.contains('?')
        || path.contains('#')
        || path.chars().any(char::is_control)
        || path
            .split('/')
            .any(|component| component.is_empty() || component == "." || component == "..")
        || matches!(
            path,
            "mimetype" | "content.xml" | "styles.xml" | "meta.xml" | "settings.xml"
        )
        || path.starts_with("META-INF/")
    {
        return Err(Error::InvalidFormat(format!(
            "invalid embedded ODP media path '{path}'"
        )));
    }
    validate_bounded_xml_value(path, "embedded media path")
}

pub(crate) fn validate_media_type(media_type: &str) -> Result<()> {
    validate_bounded_xml_value(media_type, "embedded media type")?;
    let mut segments = media_type.split(';');
    let essence = segments.next().unwrap_or_default().trim();
    let mut parts = essence.split('/');
    let top_level = parts.next().unwrap_or_default();
    let subtype = parts.next();
    let valid_token = |token: &str| {
        !token.is_empty()
            && token.bytes().all(|byte| {
                byte.is_ascii_alphanumeric()
                    || matches!(
                        byte,
                        b'!' | b'#'
                            | b'$'
                            | b'%'
                            | b'&'
                            | b'\''
                            | b'*'
                            | b'+'
                            | b'-'
                            | b'.'
                            | b'^'
                            | b'_'
                            | b'`'
                            | b'|'
                            | b'~'
                    )
            })
    };
    if !valid_token(top_level)
        || !subtype.is_some_and(valid_token)
        || parts.next().is_some()
        || segments.any(|parameter| parameter.trim().is_empty())
    {
        return Err(Error::InvalidFormat(format!(
            "invalid embedded media type '{media_type}'"
        )));
    }
    Ok(())
}

pub(crate) fn validate_href(href: &str) -> Result<()> {
    if href.is_empty() {
        return Err(Error::InvalidFormat(
            "ODP media href cannot be empty".to_string(),
        ));
    }
    validate_bounded_xml_value(href, "ODP media href")
}

pub(crate) fn validate_bounded_xml_value(value: &str, description: &str) -> Result<()> {
    if value.len() > 1_048_576 {
        return Err(Error::InvalidFormat(format!("{description} exceeds 1 MiB")));
    }
    if value.chars().any(|character| {
        !matches!(
            character,
            '\u{0009}' | '\u{000A}' | '\u{000D}' | '\u{0020}'..='\u{D7FF}' | '\u{E000}'..='\u{FFFD}' | '\u{10000}'..='\u{10FFFF}'
        )
    }) {
        return Err(Error::InvalidFormat(format!(
            "{description} contains a character forbidden by XML 1.0"
        )));
    }
    Ok(())
}

pub(crate) fn validate_ncname(value: &str, description: &str) -> Result<()> {
    let mut chars = value.chars();
    let Some(first) = chars.next() else {
        return Err(Error::InvalidFormat(format!(
            "{description} cannot be empty"
        )));
    };
    if !(first == '_' || first.is_alphabetic())
        || !chars.all(|character| {
            character == '_'
                || character == '-'
                || character == '.'
                || character.is_alphanumeric()
                || character == '\u{00B7}'
                || ('\u{0300}'..='\u{036F}').contains(&character)
                || ('\u{203F}'..='\u{2040}').contains(&character)
        })
    {
        return Err(Error::InvalidFormat(format!(
            "invalid {description} '{value}'"
        )));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn writes_escaped_inert_plugin_metadata() {
        let mut media = Reference::new("https://example.test/a?x=1&y=2").unwrap();
        media.set_mime_type("video/mp4; codecs=avc1").unwrap();
        media.set_show(Some(Show::New));
        media.set_actuate(Some(Actuate::OnRequest));
        media.set_xml_id("video_1").unwrap();
        media
            .add_parameter(Parameter::new("caption", "A < B").unwrap())
            .unwrap();
        let mut xml = String::new();
        media.write_xml(&mut xml).unwrap();
        assert!(xml.contains("https://example.test/a?x=1&amp;y=2"));
        assert!(xml.contains(r#"draw:mime-type="video/mp4; codecs=avc1""#));
        assert!(xml.contains(r#"draw:value="A &lt; B""#));
        assert_eq!(media.package_path(), None);
    }

    #[test]
    fn recognizes_only_safe_package_relative_targets() {
        assert_eq!(
            Reference::new("./Media/clip.ogg").unwrap().package_path(),
            Some("Media/clip.ogg")
        );
        for href in [
            "../clip.ogg",
            "/Media/clip.ogg",
            "file:///clip.ogg",
            "Media/../clip.ogg",
            "Media/clip.ogg#part",
        ] {
            assert_eq!(Reference::new(href).unwrap().package_path(), None);
        }
    }

    #[test]
    fn validates_names_values_and_media_types() {
        assert!(Reference::new("").is_err());
        let mut media = Reference::new("Media/a.ogg").unwrap();
        assert!(media.set_xml_id("1-invalid").is_err());
        assert!(media.set_mime_type("not a media type").is_err());
        assert!(Parameter::new("bad\0name", "value").is_err());
    }
}
