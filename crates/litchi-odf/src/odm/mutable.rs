use super::{
    MasterDocument, MasterSubdocument,
    builder::{validate_xml_part, write_master_package},
    document::validate_master_content_xml,
};
use crate::{
    OdfDocumentSigner, OdfEncryptionProfile, OwnedPackage, Section, TextIndex, constants,
};
use litchi_core::{Error, Result, xml::escape_xml};
use quick_xml::{XmlVersion, events::Event, name::{Namespace, ResolveResult}, reader::NsReader};
use std::collections::{BTreeMap, HashMap, HashSet};
use std::ops::Range;
use std::path::Path;

const OFFICE_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TEXT_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:text:1.0";

pub struct MutableMasterDocument {
    mimetype: String,
    content_xml: String,
    styles_xml: Option<String>,
    meta_xml: Option<String>,
    settings_xml: Option<String>,
    source_package: Option<OwnedPackage>,
    auxiliary: BTreeMap<String, (Vec<u8>, Option<String>)>,
    encryption: Option<(String, OdfEncryptionProfile)>,
    signer: Option<OdfDocumentSigner>,
}

impl MutableMasterDocument {
    pub fn new() -> Self {
        let content_xml = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><office:document-content xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" office:version=\"1.3\"><office:automatic-styles/><office:body><office:text text:global=\"true\"></office:text></office:body></office:document-content>".to_string();
        Self {
            mimetype: constants::ODF_MASTER.to_string(),
            content_xml,
            styles_xml: None,
            meta_xml: None,
            settings_xml: None,
            source_package: None,
            auxiliary: BTreeMap::new(),
            encryption: None,
            signer: None,
        }
    }

    pub fn from_document(document: MasterDocument) -> Result<Self> {
        let mimetype = document.mimetype().to_string();
        let package = document.into_owned_package();
        let content_xml = decode_part(&package, constants::ODF_CONTENT)?.ok_or_else(|| {
            Error::InvalidFormat("master package has no content.xml".to_string())
        })?;
        let styles_xml = decode_part(&package, constants::ODF_STYLES)?;
        let meta_xml = decode_part(&package, constants::ODF_META)?;
        let settings_xml = decode_part(&package, constants::ODF_SETTINGS)?;
        validate_master_content_xml(&content_xml)?;
        Ok(Self {
            mimetype,
            content_xml,
            styles_xml,
            meta_xml,
            settings_xml,
            source_package: Some(package),
            auxiliary: BTreeMap::new(),
            encryption: None,
            signer: None,
        })
    }

    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        Self::from_document(MasterDocument::from_bytes(bytes)?)
    }

    pub fn from_bytes_with_password(bytes: Vec<u8>, password: impl Into<String>) -> Result<Self> {
        Self::from_document(MasterDocument::from_bytes_with_password(bytes, password)?)
    }

    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        Self::from_bytes(std::fs::read(path)?)
    }

    pub fn content_xml(&self) -> &str {
        &self.content_xml
    }

    pub fn is_template(&self) -> bool {
        self.mimetype == constants::ODF_MASTER_TEMPLATE
    }

    pub fn set_template(&mut self, template: bool) {
        self.mimetype = if template {
            constants::ODF_MASTER_TEMPLATE
        } else {
            constants::ODF_MASTER
        }
        .to_string();
    }

    pub fn subdocuments(&self) -> Result<Vec<MasterSubdocument>> {
        validate_master_content_xml(&self.content_xml).map(|(_, values)| values)
    }

    pub fn set_styles_xml(&mut self, xml: Option<String>) -> Result<()> {
        validate_xml_part(xml.as_deref(), "styles.xml")?;
        self.styles_xml = xml;
        Ok(())
    }

    pub fn set_meta_xml(&mut self, xml: Option<String>) -> Result<()> {
        validate_xml_part(xml.as_deref(), "meta.xml")?;
        self.meta_xml = xml;
        Ok(())
    }

    pub fn set_settings_xml(&mut self, xml: Option<String>) -> Result<()> {
        validate_xml_part(xml.as_deref(), "settings.xml")?;
        self.settings_xml = xml;
        Ok(())
    }

    pub fn set_encryption(
        &mut self,
        password: impl Into<String>,
        profile: OdfEncryptionProfile,
    ) -> Result<()> {
        let password = password.into();
        if password.is_empty() {
            return Err(Error::InvalidFormat(
                "master package encryption password cannot be empty".to_string(),
            ));
        }
        self.encryption = Some((password, profile));
        Ok(())
    }

    pub fn set_document_signer(&mut self, signer: OdfDocumentSigner) {
        self.signer = Some(signer);
    }

    pub fn add_paragraph(&mut self, text: impl AsRef<str>) -> Result<()> {
        let fragment = format!("<text:p>{}</text:p>", escape_xml(text.as_ref()));
        self.insert_fragment(fragment)
    }

    pub fn add_section(&mut self, section: &Section) -> Result<()> {
        if section.source.is_some() || section.dde_source.is_some() {
            return Err(Error::InvalidFormat(
                "use add_subdocument for linked master sections".to_string(),
            ));
        }
        section.validate()?;
        self.insert_fragment(section.to_xml_fragment()?)
    }

    pub fn update_section(&mut self, name: &str, replacement: &Section) -> Result<()> {
        if replacement.source.is_some() || replacement.dde_source.is_some() {
            return Err(Error::InvalidFormat(
                "use update_subdocument for linked master sections".to_string(),
            ));
        }
        replacement.validate()?;
        let site = find_section_site(&self.content_xml, name)?;
        self.replace_candidate(apply_edits(
            &self.content_xml,
            vec![(site.span, replacement.to_xml_fragment()?)],
        )?)
    }

    pub fn remove_section(&mut self, name: &str) -> Result<()> {
        let site = find_section_site(&self.content_xml, name)?;
        self.replace_candidate(apply_edits(
            &self.content_xml,
            vec![(site.span, String::new())],
        )?)
    }

    pub fn add_index(&mut self, index: &TextIndex) -> Result<()> {
        self.insert_fragment(index.to_xml_fragment()?)
    }

    pub fn update_index(&mut self, name: &str, replacement: &TextIndex) -> Result<()> {
        let site = find_direct_named_site(&self.content_xml, name, true)?;
        self.replace_candidate(apply_edits(
            &self.content_xml,
            vec![(site, replacement.to_xml_fragment()?)],
        )?)
    }

    pub fn remove_index(&mut self, name: &str) -> Result<()> {
        let site = find_direct_named_site(&self.content_xml, name, true)?;
        self.replace_candidate(apply_edits(
            &self.content_xml,
            vec![(site, String::new())],
        )?)
    }

    pub fn add_subdocument(&mut self, subdocument: &MasterSubdocument) -> Result<()> {
        subdocument.validate()?;
        self.insert_fragment(subdocument.to_section().to_xml_fragment()?)
    }

    pub fn update_subdocument(
        &mut self,
        section_name: &str,
        replacement: &MasterSubdocument,
    ) -> Result<()> {
        if !self
            .subdocuments()?
            .iter()
            .any(|value| value.section_name() == section_name)
        {
            return Err(Error::InvalidFormat(format!(
                "linked master section '{section_name}' was not found"
            )));
        }
        replacement.validate()?;
        let site = find_section_site(&self.content_xml, section_name)?;
        let replacement = replacement.to_section();
        let canonical = replacement.to_xml_fragment()?;
        let content_start = canonical.find("<text:p>").ok_or_else(|| {
            Error::InvalidFormat("invalid canonical linked section".to_string())
        })?;
        let prefix = &canonical[..content_start];
        let suffix = "</text:section>";
        let mut fragment = String::with_capacity(site.span.len() + prefix.len());
        fragment.push_str(prefix);
        fragment.push_str(&self.content_xml[site.body_start..site.close_start]);
        fragment.push_str(suffix);
        self.replace_candidate(apply_edits(
            &self.content_xml,
            vec![(site.span, fragment)],
        )?)
    }

    pub fn remove_subdocument(&mut self, section_name: &str) -> Result<MasterSubdocument> {
        let old = self
            .subdocuments()?
            .into_iter()
            .find(|value| value.section_name() == section_name)
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "linked master section '{section_name}' was not found"
                ))
            })?;
        let site = find_section_site(&self.content_xml, section_name)?;
        let candidate = apply_edits(&self.content_xml, vec![(site.span, String::new())])?;
        self.replace_candidate(candidate)?;
        Ok(old)
    }

    pub fn move_body_element(&mut self, from: usize, to: usize) -> Result<()> {
        let scan = scan_body(&self.content_xml)?;
        if from >= scan.sites.len() || to >= scan.sites.len() {
            return Err(Error::InvalidFormat(
                "master body reorder index is out of bounds".to_string(),
            ));
        }
        let mut fragments = scan
            .sites
            .iter()
            .map(|site| self.content_xml[site.span.clone()].to_string())
            .collect::<Vec<_>>();
        let fragment = fragments.remove(from);
        fragments.insert(to, fragment);
        let edits = scan
            .sites
            .iter()
            .zip(fragments)
            .map(|(site, fragment)| (site.span.clone(), fragment))
            .collect();
        self.replace_candidate(apply_edits(&self.content_xml, edits)?)
    }

    pub fn reorder_subdocuments(&mut self, section_names: &[String]) -> Result<()> {
        let scan = scan_body(&self.content_xml)?;
        let linked = scan
            .sites
            .iter()
            .filter(|site| site.linked)
            .collect::<Vec<_>>();
        if linked.len() != section_names.len() {
            return Err(Error::InvalidFormat(
                "linked subdocument reorder must name every direct linked section".to_string(),
            ));
        }
        let mut by_name = HashMap::new();
        for site in &linked {
            let name = site.section_name.as_deref().ok_or_else(|| {
                Error::InvalidFormat("linked section has no name".to_string())
            })?;
            if by_name
                .insert(name, self.content_xml[site.span.clone()].to_string())
                .is_some()
            {
                return Err(Error::InvalidFormat(
                    "linked section names are ambiguous".to_string(),
                ));
            }
        }
        let mut requested = HashSet::new();
        let mut fragments = Vec::with_capacity(section_names.len());
        for name in section_names {
            if !requested.insert(name.as_str()) {
                return Err(Error::InvalidFormat(format!(
                    "duplicate linked section reorder name '{name}'"
                )));
            }
            fragments.push(by_name.get(name.as_str()).cloned().ok_or_else(|| {
                Error::InvalidFormat(format!("unknown linked section reorder name '{name}'"))
            })?);
        }
        let edits = linked
            .into_iter()
            .zip(fragments)
            .map(|(site, fragment)| (site.span.clone(), fragment))
            .collect();
        self.replace_candidate(apply_edits(&self.content_xml, edits)?)
    }

    pub fn to_bytes(self) -> Result<Vec<u8>> {
        write_master_package(
            &self.mimetype,
            &self.content_xml,
            self.styles_xml.as_deref(),
            self.meta_xml.as_deref(),
            self.settings_xml.as_deref(),
            &self.auxiliary,
            self.source_package.as_ref(),
            self.encryption,
            self.signer,
        )
    }

    pub fn save(self, path: impl AsRef<Path>) -> Result<()> {
        std::fs::write(path, self.to_bytes()?)?;
        Ok(())
    }

    fn insert_fragment(&mut self, fragment: String) -> Result<()> {
        let scan = scan_body(&self.content_xml)?;
        let candidate = if let Some((span, qname)) = scan.empty_text {
            let raw = &self.content_xml[span.clone()];
            let opening = raw.strip_suffix("/>").ok_or_else(|| {
                Error::InvalidFormat("invalid empty office:text element".to_string())
            })?;
            apply_edits(
                &self.content_xml,
                vec![(span, format!("{opening}>{fragment}</{qname}>"))],
            )?
        } else {
            apply_edits(
                &self.content_xml,
                vec![(scan.close_offset..scan.close_offset, fragment)],
            )?
        };
        self.replace_candidate(candidate)
    }

    fn replace_candidate(&mut self, candidate: String) -> Result<()> {
        validate_master_content_xml(&candidate)?;
        self.content_xml = candidate;
        Ok(())
    }
}

impl Default for MutableMasterDocument {
    fn default() -> Self {
        Self::new()
    }
}

fn decode_part(package: &OwnedPackage, path: &str) -> Result<Option<String>> {
    if !package.has_file(path)? {
        return Ok(None);
    }
    String::from_utf8(package.get_file(path)?)
        .map(Some)
        .map_err(|_| Error::InvalidFormat(format!("invalid UTF-8 in {path}")))
}

#[derive(Debug)]
struct BodySite {
    span: Range<usize>,
    section_name: Option<String>,
    text_name: Option<String>,
    local_name: Option<String>,
    linked: bool,
}

struct BodyScan {
    close_offset: usize,
    empty_text: Option<(Range<usize>, String)>,
    sites: Vec<BodySite>,
}

fn scan_body(xml: &str) -> Result<BodyScan> {
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut text_depth = None;
    let mut close_offset = None;
    let mut empty_text = None;
    let mut active: Option<BodySite> = None;
    let mut sites = Vec::new();
    loop {
        let before = usize::try_from(reader.buffer_position())
            .map_err(|_| Error::InvalidFormat("master XML offset overflow".to_string()))?;
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid master XML: {error}")))?;
        let after = usize::try_from(reader.buffer_position())
            .map_err(|_| Error::InvalidFormat("master XML offset overflow".to_string()))?;
        match event {
            Event::Start(element) => {
                if element_is(&reader, &element, OFFICE_NS, b"text") {
                    text_depth = Some(depth + 1);
                } else if text_depth == Some(depth) {
                    let text_name = text_attribute(&reader, &element, b"name")?;
                    active = Some(BodySite {
                        span: before..after,
                        section_name: element_is(&reader, &element, TEXT_NS, b"section")
                            .then(|| text_name.clone())
                            .flatten(),
                        text_name,
                        local_name: resolved_text_local(&reader, &element),
                        linked: false,
                    });
                } else if text_depth.is_some_and(|value| depth == value + 1)
                    && element_is(&reader, &element, TEXT_NS, b"section-source")
                {
                    if let Some(site) = &mut active {
                        site.linked = true;
                    }
                }
                depth = depth.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("master XML depth overflow".to_string())
                })?;
            },
            Event::Empty(element) => {
                if element_is(&reader, &element, OFFICE_NS, b"text") {
                    empty_text = Some((before..after, qname(&element)?));
                    close_offset = Some(before);
                } else if text_depth == Some(depth) {
                    let text_name = text_attribute(&reader, &element, b"name")?;
                    sites.push(BodySite {
                        span: before..after,
                        section_name: element_is(&reader, &element, TEXT_NS, b"section")
                            .then(|| text_name.clone())
                            .flatten(),
                        text_name,
                        local_name: resolved_text_local(&reader, &element),
                        linked: false,
                    });
                } else if text_depth.is_some_and(|value| depth == value + 1)
                    && element_is(&reader, &element, TEXT_NS, b"section-source")
                {
                    if let Some(site) = &mut active {
                        site.linked = true;
                    }
                }
            },
            Event::End(element) => {
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("master XML depth underflow".to_string())
                })?;
                if element_end_is(&reader, &element, OFFICE_NS, b"text")
                    && text_depth == Some(depth + 1)
                {
                    close_offset = Some(before);
                    text_depth = None;
                } else if text_depth == Some(depth) {
                    if let Some(mut site) = active.take() {
                        site.span.end = after;
                        sites.push(site);
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
    Ok(BodyScan {
        close_offset: close_offset.ok_or_else(|| {
            Error::InvalidFormat("master document has no office:text body".to_string())
        })?,
        empty_text,
        sites,
    })
}

fn find_direct_named_site(xml: &str, name: &str, index: bool) -> Result<Range<usize>> {
    let scan = scan_body(xml)?;
    let mut matches = scan.sites.into_iter().filter(|site| {
        site.section_name.as_deref() == Some(name)
            || (index
                && site
                    .local_name
                    .as_deref()
                    .is_some_and(|local| local.ends_with("index") || local == "table-of-content")
                && site.text_name.as_deref() == Some(name))
    });
    let site = matches.next().ok_or_else(|| {
        Error::InvalidFormat(format!("master body item '{name}' was not found"))
    })?;
    if matches.next().is_some() {
        return Err(Error::InvalidFormat(format!(
            "master body item '{name}' is ambiguous"
        )));
    }
    Ok(site.span)
}

#[derive(Debug)]
struct SectionSite {
    span: Range<usize>,
    body_start: usize,
    close_start: usize,
}

struct OpenSection {
    name: String,
    start: usize,
    open_end: usize,
    content_depth: usize,
    source_end: Option<usize>,
    source_depth: Option<usize>,
}

fn find_section_site(xml: &str, name: &str) -> Result<SectionSite> {
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut stack: Vec<OpenSection> = Vec::new();
    let mut found = None;
    loop {
        let before = usize::try_from(reader.buffer_position())
            .map_err(|_| Error::InvalidFormat("master XML offset overflow".to_string()))?;
        let event = reader.read_event_into(&mut buffer).map_err(|error| {
            Error::InvalidFormat(format!("invalid master XML: {error}"))
        })?;
        let after = usize::try_from(reader.buffer_position())
            .map_err(|_| Error::InvalidFormat("master XML offset overflow".to_string()))?;
        match event {
            Event::Start(element) => {
                if element_is(&reader, &element, TEXT_NS, b"section") {
                    stack.push(OpenSection {
                        name: text_attribute(&reader, &element, b"name")?.ok_or_else(|| {
                            Error::InvalidFormat("master section has no name".to_string())
                        })?,
                        start: before,
                        open_end: after,
                        content_depth: depth + 1,
                        source_end: None,
                        source_depth: None,
                    });
                } else if element_is(&reader, &element, TEXT_NS, b"section-source") {
                    if let Some(section) = stack.last_mut()
                        && depth == section.content_depth
                    {
                        section.source_depth = Some(depth + 1);
                    }
                }
                depth += 1;
            },
            Event::Empty(element) => {
                if element_is(&reader, &element, TEXT_NS, b"section-source") {
                    if let Some(section) = stack.last_mut()
                        && depth == section.content_depth
                    {
                        section.source_end = Some(after);
                    }
                } else if element_is(&reader, &element, TEXT_NS, b"section") {
                    let section_name = text_attribute(&reader, &element, b"name")?.ok_or_else(|| {
                        Error::InvalidFormat("master section has no name".to_string())
                    })?;
                    if section_name == name {
                        if found.is_some() {
                            return Err(Error::InvalidFormat(format!(
                                "master section '{name}' is ambiguous"
                            )));
                        }
                        found = Some(SectionSite {
                            span: before..after,
                            body_start: after,
                            close_start: after,
                        });
                    }
                }
            },
            Event::End(element) => {
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("master XML depth underflow".to_string())
                })?;
                if element_end_is(&reader, &element, TEXT_NS, b"section-source") {
                    if let Some(section) = stack.last_mut()
                        && section.source_depth == Some(depth + 1)
                    {
                        section.source_end = Some(after);
                        section.source_depth = None;
                    }
                } else if element_end_is(&reader, &element, TEXT_NS, b"section") {
                    let section = stack.pop().ok_or_else(|| {
                        Error::InvalidFormat("master section stack underflow".to_string())
                    })?;
                    if section.name == name {
                        if found.is_some() {
                            return Err(Error::InvalidFormat(format!(
                                "master section '{name}' is ambiguous"
                            )));
                        }
                        found = Some(SectionSite {
                            span: section.start..after,
                            body_start: section.source_end.unwrap_or(section.open_end),
                            close_start: before,
                        });
                    }
                }
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    found.ok_or_else(|| Error::InvalidFormat(format!("master section '{name}' was not found")))
}

fn resolved_text_local(
    reader: &NsReader<&[u8]>,
    element: &quick_xml::events::BytesStart<'_>,
) -> Option<String> {
    let (resolved, local) = reader.resolver().resolve_element(element.name());
    matches!(resolved, ResolveResult::Bound(found) if found == Namespace(TEXT_NS))
        .then(|| String::from_utf8_lossy(local.as_ref()).into_owned())
}

fn apply_edits(xml: &str, mut edits: Vec<(Range<usize>, String)>) -> Result<String> {
    edits.sort_by(|left, right| right.0.start.cmp(&left.0.start));
    let mut prior_start = xml.len();
    let mut output = xml.to_string();
    for (span, replacement) in edits {
        if span.start > span.end || span.end > prior_start || span.end > output.len() {
            return Err(Error::InvalidFormat(
                "overlapping or invalid master XML edits".to_string(),
            ));
        }
        output.replace_range(span.clone(), &replacement);
        prior_start = span.start;
    }
    Ok(output)
}

fn element_is(
    reader: &NsReader<&[u8]>,
    element: &quick_xml::events::BytesStart<'_>,
    namespace: &[u8],
    local: &[u8],
) -> bool {
    let (resolved, found_local) = reader.resolver().resolve_element(element.name());
    matches!(resolved, ResolveResult::Bound(found) if found == Namespace(namespace))
        && found_local.as_ref() == local
}

fn element_end_is(
    reader: &NsReader<&[u8]>,
    element: &quick_xml::events::BytesEnd<'_>,
    namespace: &[u8],
    local: &[u8],
) -> bool {
    let (resolved, found_local) = reader.resolver().resolve_element(element.name());
    matches!(resolved, ResolveResult::Bound(found) if found == Namespace(namespace))
        && found_local.as_ref() == local
}

fn text_attribute(
    reader: &NsReader<&[u8]>,
    element: &quick_xml::events::BytesStart<'_>,
    local: &[u8],
) -> Result<Option<String>> {
    let mut value = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| {
            Error::InvalidFormat(format!("invalid master attribute: {error}"))
        })?;
        let (resolved, found_local) = reader.resolver().resolve_attribute(attribute.key);
        if matches!(resolved, ResolveResult::Bound(found) if found == Namespace(TEXT_NS))
            && found_local.as_ref() == local
        {
            if value.is_some() {
                return Err(Error::InvalidFormat(
                    "duplicate master section name attribute".to_string(),
                ));
            }
            value = Some(
                attribute
                    .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid master attribute: {error}"))
                    })?
                    .into_owned(),
            );
        }
    }
    Ok(value)
}

fn qname(element: &quick_xml::events::BytesStart<'_>) -> Result<String> {
    std::str::from_utf8(element.name().as_ref())
        .map(str::to_string)
        .map_err(|_| Error::InvalidFormat("non-UTF-8 master element name".to_string()))
}
