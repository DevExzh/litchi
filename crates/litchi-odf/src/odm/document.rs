//! Typed access to OpenDocument text master documents.

use crate::OdfMetadata;
use crate::constants;
use crate::core::OwnedPackage;
use crate::odt::Document;
use litchi_core::{Error, Metadata, Result};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::io::Read;
use std::path::Path;

const OFFICE_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TEXT_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const XLINK_NAMESPACE: &str = "http://www.w3.org/1999/xlink";
const XML_NAMESPACE: &str = "http://www.w3.org/XML/1998/namespace";
const MAX_DEPTH: usize = 128;
const MAX_SECTIONS: usize = 65_536;
const MAX_ATTRIBUTE_BYTES: usize = 1_048_576;

/// One linked section in an OpenDocument master document.
///
/// The link and its cached section content remain inert. Litchi never opens or
/// refreshes the referenced document automatically.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MasterSubdocument {
    section_name: String,
    style_name: Option<String>,
    protected: Option<bool>,
    protection_key: Option<String>,
    protection_key_digest_algorithm: Option<String>,
    display: Option<String>,
    condition: Option<String>,
    xml_id: Option<String>,
    href: Option<String>,
    source_section_name: Option<String>,
    filter_name: Option<String>,
    xlink_type: Option<String>,
    xlink_show: Option<String>,
}

impl MasterSubdocument {
    /// Create a linked subdocument. The URI remains inert and is never fetched.
    pub fn new(section_name: impl Into<String>, href: impl Into<String>) -> Result<Self> {
        let value = Self {
            section_name: section_name.into(),
            style_name: None,
            protected: None,
            protection_key: None,
            protection_key_digest_algorithm: None,
            display: None,
            condition: None,
            xml_id: None,
            href: Some(href.into()),
            source_section_name: None,
            filter_name: None,
            xlink_type: Some("simple".to_string()),
            xlink_show: Some("embed".to_string()),
        };
        value.validate()?;
        Ok(value)
    }

    pub fn set_style_name(&mut self, value: Option<String>) -> Result<&mut Self> {
        let mut candidate = self.clone();
        candidate.style_name = value;
        candidate.validate()?;
        *self = candidate;
        Ok(self)
    }

    pub fn set_xml_id(&mut self, value: Option<String>) -> Result<&mut Self> {
        let mut candidate = self.clone();
        candidate.xml_id = value;
        candidate.validate()?;
        *self = candidate;
        Ok(self)
    }

    pub fn set_source_section_name(&mut self, value: Option<String>) -> Result<&mut Self> {
        let mut candidate = self.clone();
        candidate.source_section_name = value;
        candidate.validate()?;
        *self = candidate;
        Ok(self)
    }

    pub fn set_filter_name(&mut self, value: Option<String>) -> Result<&mut Self> {
        let mut candidate = self.clone();
        candidate.filter_name = value;
        candidate.validate()?;
        *self = candidate;
        Ok(self)
    }

    pub fn set_protection(
        &mut self,
        protected: bool,
        key: Option<String>,
        digest_algorithm: Option<String>,
    ) -> Result<&mut Self> {
        let mut candidate = self.clone();
        candidate.protected = Some(protected);
        candidate.protection_key = key;
        candidate.protection_key_digest_algorithm = digest_algorithm;
        candidate.validate()?;
        *self = candidate;
        Ok(self)
    }

    pub fn validate(&self) -> Result<()> {
        if self.href.as_deref().is_none_or(str::is_empty) {
            return Err(Error::InvalidFormat(
                "master subdocument requires a non-empty inert href".to_string(),
            ));
        }
        if self.xlink_type.as_deref().is_some_and(|value| value != "simple")
            || self.xlink_show.as_deref().is_some_and(|value| value != "embed")
        {
            return Err(Error::InvalidFormat(
                "master subdocument requires simple/embed XLink behavior".to_string(),
            ));
        }
        self.to_section().validate()
    }

    pub(crate) fn to_section(&self) -> crate::Section {
        crate::Section {
            name: self.section_name.clone(),
            style: self.style_name.clone(),
            protected: self.protected.unwrap_or(false),
            xml_id: self.xml_id.clone(),
            protection_key: self.protection_key.clone(),
            protection_key_digest_algorithm: self.protection_key_digest_algorithm.clone(),
            display: match self.display.as_deref() {
                Some("none") => crate::SectionDisplay::Hidden,
                Some("condition") => crate::SectionDisplay::Condition,
                _ => crate::SectionDisplay::Visible,
            },
            condition: self.condition.clone(),
            source: Some(crate::SectionSource {
                href: self.href.clone(),
                section_name: self.source_section_name.clone(),
                filter_name: self.filter_name.clone(),
            }),
            dde_source: None,
            content: String::new(),
        }
    }

    /// Return the required containing `text:section` name.
    pub fn section_name(&self) -> &str {
        &self.section_name
    }

    /// Return the containing section's style reference.
    pub fn style_name(&self) -> Option<&str> {
        self.style_name.as_deref()
    }

    /// Return the optional section protection flag.
    pub fn protected(&self) -> Option<bool> {
        self.protected
    }

    /// Return the exact optional protection key.
    pub fn protection_key(&self) -> Option<&str> {
        self.protection_key.as_deref()
    }

    /// Return the protection-key digest algorithm URI.
    pub fn protection_key_digest_algorithm(&self) -> Option<&str> {
        self.protection_key_digest_algorithm.as_deref()
    }

    /// Return `text:display` (`true`, `none`, or `condition`).
    pub fn display(&self) -> Option<&str> {
        self.display.as_deref()
    }

    /// Return the display condition expression when present.
    pub fn condition(&self) -> Option<&str> {
        self.condition.as_deref()
    }

    /// Return the containing section's `xml:id`.
    pub fn xml_id(&self) -> Option<&str> {
        self.xml_id.as_deref()
    }

    /// Return the inert linked-document URI.
    pub fn href(&self) -> Option<&str> {
        self.href.as_deref()
    }

    /// Return the named section selected inside the linked document.
    pub fn source_section_name(&self) -> Option<&str> {
        self.source_section_name.as_deref()
    }

    /// Return the producer filter name used for the linked document.
    pub fn filter_name(&self) -> Option<&str> {
        self.filter_name.as_deref()
    }

    /// Return the exact optional XLink type.
    pub fn xlink_type(&self) -> Option<&str> {
        self.xlink_type.as_deref()
    }

    /// Return the exact optional XLink show behavior.
    pub fn xlink_show(&self) -> Option<&str> {
        self.xlink_show.as_deref()
    }
}

struct SectionBuilder {
    open_depth: usize,
    section: MasterSubdocument,
    source_seen: bool,
    content_seen: bool,
}

/// A validated OpenDocument Master document (`.odm`) or template (`.otm`).
///
/// The complete ODT semantic reader is available through [`Self::document`].
pub struct MasterDocument {
    document: Document,
    mimetype: String,
    template: bool,
    global: Option<bool>,
    subdocuments: Vec<MasterSubdocument>,
}

impl MasterDocument {
    /// Open a master document from a path.
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        let file = std::fs::File::open(path)?;
        Self::from_reader(file)
    }

    /// Read a master document from a stream.
    pub fn from_reader(mut reader: impl Read) -> Result<Self> {
        let mut bytes = Vec::new();
        reader.read_to_end(&mut bytes)?;
        Self::from_bytes(bytes)
    }

    /// Validate a master document from owned package bytes.
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        Self::from_owned_package(OwnedPackage::from_bytes(bytes)?)
    }

    /// Validate an encrypted master package with the supplied password.
    pub fn from_bytes_with_password(bytes: Vec<u8>, password: impl Into<String>) -> Result<Self> {
        Self::from_owned_package(OwnedPackage::from_bytes_with_password(bytes, password)?)
    }

    /// Open and validate an encrypted master package.
    pub fn open_with_password(
        path: impl AsRef<Path>,
        password: impl Into<String>,
    ) -> Result<Self> {
        Self::from_bytes_with_password(std::fs::read(path)?, password)
    }

    fn from_owned_package(package: OwnedPackage) -> Result<Self> {
        let mimetype = package.mimetype()?;
        let template = match mimetype.as_str() {
            constants::ODF_MASTER => false,
            constants::ODF_MASTER_TEMPLATE => true,
            _ => {
                return Err(Error::InvalidFormat(format!(
                    "not an OpenDocument master: MIME type is '{mimetype}'"
                )));
            },
        };
        if package.package()?.manifest().get_media_type("/") != Some(mimetype.as_str()) {
            return Err(Error::InvalidFormat(
                "master package mimetype and root manifest media type differ".to_string(),
            ));
        }
        let content = package.get_file(constants::ODF_CONTENT)?;
        let content = std::str::from_utf8(&content)
            .map_err(|_| Error::InvalidFormat("invalid UTF-8 in content.xml".to_string()))?;
        let (global, subdocuments) = validate_master_content_xml(content)?;
        let document = Document::from_owned_package(package)?;
        Ok(Self {
            document,
            mimetype,
            template,
            global,
            subdocuments,
        })
    }

    /// Whether this is an `.otm` master template.
    pub fn is_template(&self) -> bool {
        self.template
    }

    /// Return the exact package MIME type.
    pub fn mimetype(&self) -> &str {
        &self.mimetype
    }

    /// Return the optional `text:global` flag from `office:text`.
    ///
    /// LibreOffice commonly identifies master documents by MIME type and omits
    /// this optional schema attribute.
    pub fn global(&self) -> Option<bool> {
        self.global
    }

    /// Return linked section sources in document order.
    pub fn subdocuments(&self) -> &[MasterSubdocument] {
        &self.subdocuments
    }

    /// Return the complete text-document semantic reader.
    pub fn document(&self) -> &Document {
        &self.document
    }

    /// Inspect ordered ODF variable declarations without evaluating fields or formulas.
    pub fn variable_declarations(&self) -> Result<crate::OdfVariableDeclarations> {
        self.document.variable_declarations()
    }

    /// Extract cached visible text without opening linked documents.
    pub fn text(&self) -> Result<String> {
        self.document.text()
    }

    /// Extract common package metadata.
    pub fn metadata(&self) -> Result<Metadata> {
        self.document.metadata()
    }

    /// Extract complete OpenDocument metadata.
    pub fn odf_metadata(&self) -> Result<Option<OdfMetadata>> {
        self.document.odf_metadata()
    }

    /// Return the exact original package bytes.
    pub fn as_bytes(&self) -> &[u8] {
        self.document.original_bytes()
    }

    /// Clone the exact original package bytes.
    pub fn to_bytes(&self) -> Vec<u8> {
        self.as_bytes().to_vec()
    }

    /// Save without reconstructing or refreshing linked content.
    pub fn save(&self, path: impl AsRef<Path>) -> Result<()> {
        std::fs::write(path, self.as_bytes())?;
        Ok(())
    }

    pub(crate) fn into_owned_package(self) -> OwnedPackage {
        self.document.into_package()
    }
}

pub(crate) fn validate_master_content_xml(
    xml: &str,
) -> Result<(Option<bool>, Vec<MasterSubdocument>)> {
    if xml.len() > 64 * 1024 * 1024 {
        return Err(Error::InvalidFormat(
            "master content.xml exceeds 64 MiB".to_string(),
        ));
    }
    let parsed = parse_master_content(xml)?;
    let mut names = std::collections::HashSet::new();
    for subdocument in &parsed.1 {
        subdocument.validate()?;
    }
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut ids = std::collections::HashSet::new();
    loop {
        match reader.read_event_into(&mut buffer).map_err(|error| {
            Error::InvalidFormat(format!("invalid master XML: {error}"))
        })? {
            Event::Start(element) | Event::Empty(element) => {
                if element_matches(&reader, &element, TEXT_NAMESPACE, "section") {
                    let parsed_section = parse_section(&reader, &element)?;
                    let mut section = parsed_section.to_section();
                    section.source = None;
                    section.validate()?;
                    if !names.insert(section.name.clone()) {
                        return Err(Error::InvalidFormat(format!(
                            "duplicate master section name '{}'",
                            section.name
                        )));
                    }
                }
                if let Some(id) = optional_attribute(&reader, &element, XML_NAMESPACE, "id")? {
                    validate_xml_id(&id)?;
                    if !ids.insert(id.clone()) {
                        return Err(Error::InvalidFormat(format!(
                            "duplicate master xml:id '{id}'"
                        )));
                    }
                }
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(Error::InvalidFormat(
                    "active XML declarations are prohibited in master content".to_string(),
                ));
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    Ok(parsed)
}

fn element_matches(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    namespace: &str,
    local: &str,
) -> bool {
    let (resolved, found_local) = reader.resolver().resolve_element(element.name());
    matches!(resolved, ResolveResult::Bound(Namespace(uri)) if uri == namespace.as_bytes())
        && found_local.as_ref() == local.as_bytes()
}

fn validate_xml_id(value: &str) -> Result<()> {
    let mut chars = value.chars();
    let Some(first) = chars.next() else {
        return Err(Error::InvalidFormat("master xml:id cannot be empty".to_string()));
    };
    if !(first == '_' || first.is_alphabetic())
        || chars.any(|character| {
            !(character == '_' || character == '-' || character == '.' || character.is_alphanumeric())
        })
    {
        return Err(Error::InvalidFormat(format!(
            "master xml:id '{value}' is not an XML NCName"
        )));
    }
    Ok(())
}

fn parse_master_content(xml: &str) -> Result<(Option<bool>, Vec<MasterSubdocument>)> {
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut body_seen = false;
    let mut body_depth = None;
    let mut text_seen = false;
    let mut text_depth = None;
    let mut global = None;
    let mut sections = Vec::new();
    let mut stack: Vec<SectionBuilder> = Vec::new();
    let mut source_depth = None;

    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid master XML: {error}")))?;
        match event {
            Event::Start(ref element) => {
                let namespace_uri = namespace_uri(&namespace)?;
                let local = decode_utf8(element.local_name().as_ref(), "element name")?;
                if source_depth.is_some() {
                    return Err(Error::InvalidFormat(
                        "text:section-source cannot contain elements".to_string(),
                    ));
                }
                validate_master_container(
                    &reader,
                    element,
                    namespace_uri.as_deref(),
                    &local,
                    depth,
                    &mut root_seen,
                    root_closed,
                    &mut body_seen,
                    &mut body_depth,
                    &mut text_seen,
                    &mut text_depth,
                    &mut global,
                )?;
                let is_source =
                    namespace_uri.as_deref() == Some(TEXT_NAMESPACE) && local == "section-source";
                if let Some(section) = stack.last_mut()
                    && depth == section.open_depth
                    && !is_source
                {
                    section.content_seen = true;
                }
                if text_depth.is_some()
                    && namespace_uri.as_deref() == Some(TEXT_NAMESPACE)
                    && local == "section"
                {
                    if stack.len() >= MAX_DEPTH || sections.len() + stack.len() >= MAX_SECTIONS {
                        return Err(Error::InvalidFormat(
                            "master document has excessive section nesting or count".to_string(),
                        ));
                    }
                    stack.push(SectionBuilder {
                        open_depth: depth + 1,
                        section: parse_section(&reader, element)?,
                        source_seen: false,
                        content_seen: false,
                    });
                } else if text_depth.is_some()
                    && namespace_uri.as_deref() == Some(TEXT_NAMESPACE)
                    && local == "section-source"
                {
                    sections.push(attach_source(&reader, element, depth, &mut stack)?);
                    source_depth = Some(depth + 1);
                }
                depth = depth.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("master XML nesting overflow".to_string())
                })?;
                if depth > MAX_DEPTH {
                    return Err(Error::InvalidFormat(format!(
                        "master XML nesting exceeds {MAX_DEPTH} levels"
                    )));
                }
            },
            Event::Empty(ref element) => {
                let namespace_uri = namespace_uri(&namespace)?;
                let local = decode_utf8(element.local_name().as_ref(), "element name")?;
                let is_source =
                    namespace_uri.as_deref() == Some(TEXT_NAMESPACE) && local == "section-source";
                if let Some(section) = stack.last_mut()
                    && depth == section.open_depth
                    && !is_source
                {
                    section.content_seen = true;
                }
                if depth == 0 {
                    return Err(Error::InvalidFormat(
                        "master content root cannot be empty".to_string(),
                    ));
                }
                if depth == 1
                    && namespace_uri.as_deref() == Some(OFFICE_NAMESPACE)
                    && local == "body"
                {
                    if body_seen {
                        return Err(Error::InvalidFormat("duplicate office:body".to_string()));
                    }
                    body_seen = true;
                } else if depth == 2 && body_depth == Some(2) {
                    if namespace_uri.as_deref() != Some(OFFICE_NAMESPACE)
                        || local != "text"
                        || text_seen
                    {
                        return Err(Error::InvalidFormat(
                            "master body must contain exactly one office:text".to_string(),
                        ));
                    }
                    text_seen = true;
                    global = optional_bool_attribute(&reader, element, TEXT_NAMESPACE, "global")?;
                } else if text_depth.is_some()
                    && namespace_uri.as_deref() == Some(TEXT_NAMESPACE)
                    && local == "section-source"
                {
                    sections.push(attach_source(&reader, element, depth, &mut stack)?);
                }
            },
            Event::End(ref element) => {
                let namespace_uri = namespace_uri(&namespace)?;
                let local = decode_utf8(element.local_name().as_ref(), "element name")?;
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("unexpected master XML closing tag".to_string())
                })?;
                if source_depth == Some(depth + 1) {
                    source_depth = None;
                }
                if namespace_uri.as_deref() == Some(TEXT_NAMESPACE) && local == "section" {
                    let section = stack.pop().ok_or_else(|| {
                        Error::InvalidFormat("master section stack underflow".to_string())
                    })?;
                    if section.open_depth != depth + 1 {
                        return Err(Error::InvalidFormat(
                            "master section hierarchy is inconsistent".to_string(),
                        ));
                    }
                }
                if namespace_uri.as_deref() == Some(OFFICE_NAMESPACE)
                    && local == "text"
                    && depth == 2
                {
                    text_depth = None;
                } else if namespace_uri.as_deref() == Some(OFFICE_NAMESPACE)
                    && local == "body"
                    && depth == 1
                {
                    body_depth = None;
                }
                if depth == 0 {
                    root_closed = true;
                }
            },
            Event::Text(ref text)
                if source_depth.is_some() && !text.iter().all(u8::is_ascii_whitespace) =>
            {
                return Err(Error::InvalidFormat(
                    "text:section-source cannot contain text".to_string(),
                ));
            },
            Event::Text(ref text)
                if stack
                    .last()
                    .is_some_and(|section| depth == section.open_depth)
                    && !text.iter().all(u8::is_ascii_whitespace) =>
            {
                return Err(Error::InvalidFormat(
                    "text:section cannot contain direct character data".to_string(),
                ));
            },
            Event::Text(ref text) if depth == 0 && !text.iter().all(u8::is_ascii_whitespace) => {
                return Err(Error::InvalidFormat(
                    "text is not allowed outside the master root".to_string(),
                ));
            },
            Event::CData(_) | Event::GeneralRef(_) if depth == 0 || source_depth.is_some() => {
                return Err(Error::InvalidFormat(
                    "content is not allowed in this master XML position".to_string(),
                ));
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(Error::InvalidFormat(
                    "active XML declarations are prohibited in master content".to_string(),
                ));
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if !root_seen
        || !root_closed
        || depth != 0
        || !body_seen
        || body_depth.is_some()
        || !text_seen
        || text_depth.is_some()
        || !stack.is_empty()
        || source_depth.is_some()
    {
        return Err(Error::InvalidFormat(
            "incomplete OpenDocument master structure".to_string(),
        ));
    }
    Ok((global, sections))
}

#[allow(clippy::too_many_arguments)]
fn validate_master_container(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    namespace_uri: Option<&str>,
    local: &str,
    depth: usize,
    root_seen: &mut bool,
    root_closed: bool,
    body_seen: &mut bool,
    body_depth: &mut Option<usize>,
    text_seen: &mut bool,
    text_depth: &mut Option<usize>,
    global: &mut Option<bool>,
) -> Result<()> {
    if depth == 0 {
        if *root_seen
            || root_closed
            || namespace_uri != Some(OFFICE_NAMESPACE)
            || local != "document-content"
        {
            return Err(Error::InvalidFormat(
                "master content must have one office:document-content root".to_string(),
            ));
        }
        *root_seen = true;
    } else if depth == 1 && namespace_uri == Some(OFFICE_NAMESPACE) && local == "body" {
        if *body_seen || body_depth.is_some() {
            return Err(Error::InvalidFormat("duplicate office:body".to_string()));
        }
        *body_seen = true;
        *body_depth = Some(2);
    } else if depth == 2 && *body_depth == Some(2) {
        if namespace_uri != Some(OFFICE_NAMESPACE) || local != "text" || *text_seen {
            return Err(Error::InvalidFormat(
                "master body must contain exactly one office:text".to_string(),
            ));
        }
        *text_seen = true;
        *text_depth = Some(3);
        *global = optional_bool_attribute(reader, element, TEXT_NAMESPACE, "global")?;
    }
    Ok(())
}

fn parse_section(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<MasterSubdocument> {
    let section_name = required_attribute(reader, element, TEXT_NAMESPACE, "name")?;
    let display = optional_attribute(reader, element, TEXT_NAMESPACE, "display")?;
    let condition = optional_attribute(reader, element, TEXT_NAMESPACE, "condition")?;
    if display
        .as_deref()
        .is_some_and(|value| !matches!(value, "true" | "none" | "condition"))
        || (display.as_deref() == Some("condition")) != condition.is_some()
    {
        return Err(Error::InvalidFormat(
            "invalid text:section display/condition attributes".to_string(),
        ));
    }
    Ok(MasterSubdocument {
        section_name,
        style_name: optional_attribute(reader, element, TEXT_NAMESPACE, "style-name")?,
        protected: optional_bool_attribute(reader, element, TEXT_NAMESPACE, "protected")?,
        protection_key: optional_attribute(reader, element, TEXT_NAMESPACE, "protection-key")?,
        protection_key_digest_algorithm: optional_attribute(
            reader,
            element,
            TEXT_NAMESPACE,
            "protection-key-digest-algorithm",
        )?,
        display,
        condition,
        xml_id: optional_attribute(reader, element, XML_NAMESPACE, "id")?,
        href: None,
        source_section_name: None,
        filter_name: None,
        xlink_type: None,
        xlink_show: None,
    })
}

fn attach_source(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    depth: usize,
    stack: &mut [SectionBuilder],
) -> Result<MasterSubdocument> {
    let section = stack.last_mut().ok_or_else(|| {
        Error::InvalidFormat("text:section-source must be inside text:section".to_string())
    })?;
    if depth != section.open_depth || section.source_seen || section.content_seen {
        return Err(Error::InvalidFormat(
            "text:section-source must be the unique direct section source".to_string(),
        ));
    }
    let xlink_type = optional_attribute(reader, element, XLINK_NAMESPACE, "type")?;
    let xlink_show = optional_attribute(reader, element, XLINK_NAMESPACE, "show")?;
    if xlink_type.as_deref().is_some_and(|value| value != "simple")
        || xlink_show.as_deref().is_some_and(|value| value != "embed")
    {
        return Err(Error::InvalidFormat(
            "invalid text:section-source XLink behavior".to_string(),
        ));
    }
    section.section.href = optional_attribute(reader, element, XLINK_NAMESPACE, "href")?;
    section.section.source_section_name =
        optional_attribute(reader, element, TEXT_NAMESPACE, "section-name")?;
    section.section.filter_name =
        optional_attribute(reader, element, TEXT_NAMESPACE, "filter-name")?;
    section.section.xlink_type = xlink_type;
    section.section.xlink_show = xlink_show;
    section.source_seen = true;
    Ok(section.section.clone())
}

fn required_attribute(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    namespace: &str,
    local: &str,
) -> Result<String> {
    optional_attribute(reader, element, namespace, local)?
        .ok_or_else(|| Error::InvalidFormat(format!("text:section is missing {local}")))
}

fn optional_bool_attribute(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    namespace: &str,
    local: &str,
) -> Result<Option<bool>> {
    optional_attribute(reader, element, namespace, local)?
        .map(|value| match value.as_str() {
            "true" | "1" => Ok(true),
            "false" | "0" => Ok(false),
            _ => Err(Error::InvalidFormat(format!(
                "invalid boolean {local} value '{value}'"
            ))),
        })
        .transpose()
}

fn optional_attribute(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    namespace: &str,
    local: &str,
) -> Result<Option<String>> {
    let mut found = None;
    for attribute in element.attributes() {
        let attribute = attribute
            .map_err(|error| Error::InvalidFormat(format!("invalid master attribute: {error}")))?;
        let (resolved, resolved_local) = reader.resolver().resolve_attribute(attribute.key);
        if matches!(resolved, ResolveResult::Bound(Namespace(uri)) if uri == namespace.as_bytes())
            && resolved_local.as_ref() == local.as_bytes()
        {
            if found.is_some() {
                return Err(Error::InvalidFormat(format!(
                    "duplicate expanded master attribute '{local}'"
                )));
            }
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                .map_err(|error| {
                    Error::InvalidFormat(format!("invalid master attribute value: {error}"))
                })?
                .into_owned();
            if value.len() > MAX_ATTRIBUTE_BYTES {
                return Err(Error::InvalidFormat(
                    "master attribute exceeds 1 MiB".to_string(),
                ));
            }
            found = Some(value);
        }
    }
    Ok(found)
}

fn namespace_uri(namespace: &ResolveResult<'_>) -> Result<Option<String>> {
    match namespace {
        ResolveResult::Unbound => Ok(None),
        ResolveResult::Bound(Namespace(uri)) => decode_utf8(uri, "namespace URI").map(Some),
        ResolveResult::Unknown(prefix) => Err(Error::InvalidFormat(format!(
            "unknown master namespace prefix '{}'",
            String::from_utf8_lossy(prefix)
        ))),
    }
}

fn decode_utf8(bytes: &[u8], kind: &str) -> Result<String> {
    std::str::from_utf8(bytes)
        .map(str::to_string)
        .map_err(|_| Error::InvalidFormat(format!("non-UTF-8 master {kind}")))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::core::PackageWriter;
    use std::io::Cursor;

    fn package(mimetype: &str, content: &str) -> Vec<u8> {
        let mut writer = PackageWriter::new();
        writer.set_mimetype(mimetype).unwrap();
        writer
            .add_file(constants::ODF_CONTENT, content.as_bytes())
            .unwrap();
        writer.finish_to_bytes().unwrap()
    }

    fn master_xml() -> &'static str {
        r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
 xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
 xmlns:xlink="http://www.w3.org/1999/xlink">
 <office:automatic-styles/><office:body><office:text text:global="true">
  <text:p>Master introduction</text:p>
  <text:section text:name="Chapter One" text:style-name="Sect1" text:protected="1" text:display="condition" text:condition="ooow:chapter" xml:id="chapter-one">
   <text:section-source xlink:href="../chapter1.odt" xlink:type="simple" xlink:show="embed" text:section-name="Body" text:filter-name="writer8"></text:section-source>
   <text:h text:outline-level="1">Cached chapter</text:h>
  </text:section>
  <text:section text:name="Appendix"><text:section-source xlink:href="https://example.invalid/a.odt"/></text:section>
 </office:text></office:body>
</office:document-content>"#
    }

    #[test]
    fn parses_libreoffice_style_master_links_and_reuses_text_model_losslessly() {
        let bytes = package(constants::ODF_MASTER, master_xml());
        let document = MasterDocument::from_bytes(bytes.clone()).unwrap();
        assert!(!document.is_template());
        assert_eq!(document.global(), Some(true));
        assert_eq!(document.subdocuments().len(), 2);
        let chapter = &document.subdocuments()[0];
        assert_eq!(chapter.section_name(), "Chapter One");
        assert_eq!(chapter.style_name(), Some("Sect1"));
        assert_eq!(chapter.protected(), Some(true));
        assert_eq!(chapter.display(), Some("condition"));
        assert_eq!(chapter.condition(), Some("ooow:chapter"));
        assert_eq!(chapter.xml_id(), Some("chapter-one"));
        assert_eq!(chapter.href(), Some("../chapter1.odt"));
        assert_eq!(chapter.source_section_name(), Some("Body"));
        assert_eq!(chapter.filter_name(), Some("writer8"));
        assert_eq!(chapter.xlink_type(), Some("simple"));
        assert_eq!(chapter.xlink_show(), Some("embed"));
        assert!(document.text().unwrap().contains("Master introduction"));
        assert!(document.text().unwrap().contains("Cached chapter"));
        assert_eq!(document.as_bytes(), bytes);
        assert_eq!(document.to_bytes(), bytes);
    }

    #[test]
    fn accepts_master_templates_readers_and_omitted_global_flag() {
        let xml = r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><o:body><o:text/></o:body></o:document-content>"#;
        let bytes = package(constants::ODF_MASTER_TEMPLATE, xml);
        let document = MasterDocument::from_reader(Cursor::new(bytes.clone())).unwrap();
        assert!(document.is_template());
        assert_eq!(document.global(), None);
        assert!(document.subdocuments().is_empty());
        assert_eq!(document.as_bytes(), bytes);
    }

    #[test]
    fn reused_text_model_accepts_arbitrary_prefixes_and_decodes_text() {
        let xml = format!(
            r#"<o:document-content xmlns:o="{OFFICE_NAMESPACE}" xmlns:t="{TEXT_NAMESPACE}"><o:body><o:text t:global="true"><t:h t:outline-level="2">Master &amp; Cache</t:h><t:p>A<t:span>B</t:span>C<t:s t:c="2"/>D</t:p></o:text></o:body></o:document-content>"#
        );
        let document = MasterDocument::from_bytes(package(constants::ODF_MASTER, &xml)).unwrap();
        assert_eq!(document.text().unwrap(), "Master & Cache\nABC  D");
        assert_eq!(document.document().paragraph_count().unwrap(), 1);
        assert_eq!(
            document.document().paragraphs().unwrap()[0].text().unwrap(),
            "ABC  D"
        );
    }

    #[test]
    fn rejects_other_families_and_invalid_master_hierarchy() {
        assert!(MasterDocument::from_bytes(package(constants::ODF_TEXT, master_xml())).is_err());
        for xml in [
            r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><o:body><o:spreadsheet/></o:body></o:document-content>"#,
            r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><o:body/></o:document-content>"#,
            r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><o:body><o:text/></o:body>"#,
            r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><o:body><o:text/></o:body></o:document-content><o:x/>"#,
        ] {
            assert!(
                MasterDocument::from_bytes(package(constants::ODF_MASTER, xml)).is_err(),
                "accepted {xml}"
            );
        }
    }

    #[test]
    fn rejects_malformed_section_sources_and_attributes() {
        for body in [
            r#"<t:section><t:section-source x:href="a.odt"/></t:section>"#,
            r#"<t:section t:name="A"><t:p>cached</t:p><t:section-source x:href="a.odt"/></t:section>"#,
            r#"<t:section t:name="A"><t:section-source x:type="extended" x:href="a.odt"/></t:section>"#,
            r#"<t:section t:name="A"><t:section-source x:show="new" x:href="a.odt"/></t:section>"#,
            r#"<t:section t:name="A"><t:section-source x:href="a.odt"/><t:section-source x:href="b.odt"/></t:section>"#,
            r#"<t:section t:name="A"><t:section-source x:href="a.odt">bad</t:section-source></t:section>"#,
            r#"<t:section t:name="A" t:protected="yes"><t:section-source/></t:section>"#,
            r#"<t:section t:name="A" t:display="condition"><t:section-source/></t:section>"#,
        ] {
            let xml = format!(
                r#"<o:document-content xmlns:o="{OFFICE_NAMESPACE}" xmlns:t="{TEXT_NAMESPACE}" xmlns:x="{XLINK_NAMESPACE}"><o:body><o:text>{body}</o:text></o:body></o:document-content>"#
            );
            assert!(
                MasterDocument::from_bytes(package(constants::ODF_MASTER, &xml)).is_err(),
                "accepted {body}"
            );
        }
    }

    #[test]
    fn rejects_invalid_global_and_excessive_nesting() {
        let invalid = format!(
            r#"<o:document-content xmlns:o="{OFFICE_NAMESPACE}" xmlns:t="{TEXT_NAMESPACE}"><o:body><o:text t:global="yes"/></o:body></o:document-content>"#
        );
        assert!(MasterDocument::from_bytes(package(constants::ODF_MASTER, &invalid)).is_err());

        let nested =
            "<t:section t:name=\"A\">".repeat(MAX_DEPTH) + &"</t:section>".repeat(MAX_DEPTH);
        let deep = format!(
            r#"<o:document-content xmlns:o="{OFFICE_NAMESPACE}" xmlns:t="{TEXT_NAMESPACE}"><o:body><o:text>{nested}</o:text></o:body></o:document-content>"#
        );
        assert!(MasterDocument::from_bytes(package(constants::ODF_MASTER, &deep)).is_err());
    }
}
