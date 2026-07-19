use super::{MasterDocument, MasterSubdocument, document::validate_master_content_xml};
use crate::{
    OdfDocumentSigner, OdfEncryptionProfile, OdfStructure, PackageWriter, Section, TextIndex,
    constants,
};
use litchi_core::{Error, Result, xml::escape_xml};
use quick_xml::events::Event;
use quick_xml::reader::Reader;
use std::collections::{BTreeMap, HashSet};
use std::path::Path;

const MAX_ELEMENTS: usize = 65_536;
const MAX_DEPTH: usize = 128;

#[derive(Debug, Clone)]
pub enum MasterDocumentElement {
    Paragraph(String),
    Section(MasterSection),
    Index(TextIndex),
    Subdocument(MasterSubdocument),
}

#[derive(Debug, Clone)]
pub struct MasterSection {
    pub section: Section,
    pub children: Vec<MasterDocumentElement>,
}

impl MasterSection {
    pub fn new(section: Section) -> Result<Self> {
        if section.source.is_some() || section.dde_source.is_some() {
            return Err(Error::InvalidFormat(
                "local master sections cannot contain a source declaration".to_string(),
            ));
        }
        section.validate()?;
        Ok(Self {
            section,
            children: Vec::new(),
        })
    }

    pub fn push(&mut self, element: MasterDocumentElement) -> Result<&mut Self> {
        if self.children.len() >= MAX_ELEMENTS {
            return Err(Error::InvalidFormat(
                "master section exceeds 65536 children".to_string(),
            ));
        }
        self.children.push(element);
        Ok(self)
    }
}

pub struct MasterDocumentBuilder {
    mimetype: String,
    global: Option<bool>,
    elements: Vec<MasterDocumentElement>,
    styles_xml: Option<String>,
    meta_xml: Option<String>,
    settings_xml: Option<String>,
    auxiliary: BTreeMap<String, (Vec<u8>, Option<String>)>,
    encryption: Option<(String, OdfEncryptionProfile)>,
    signer: Option<OdfDocumentSigner>,
}

impl MasterDocumentBuilder {
    pub fn new() -> Self {
        Self {
            mimetype: constants::ODF_MASTER.to_string(),
            global: Some(true),
            elements: Vec::new(),
            styles_xml: None,
            meta_xml: None,
            settings_xml: None,
            auxiliary: BTreeMap::new(),
            encryption: None,
            signer: None,
        }
    }

    pub fn template() -> Self {
        let mut value = Self::new();
        value.mimetype = constants::ODF_MASTER_TEMPLATE.to_string();
        value
    }

    pub fn set_template(&mut self, template: bool) -> &mut Self {
        self.mimetype = if template {
            constants::ODF_MASTER_TEMPLATE
        } else {
            constants::ODF_MASTER
        }
        .to_string();
        self
    }

    pub fn set_global(&mut self, global: Option<bool>) -> &mut Self {
        self.global = global;
        self
    }

    pub fn push(&mut self, element: MasterDocumentElement) -> Result<&mut Self> {
        if self.elements.len() >= MAX_ELEMENTS {
            return Err(Error::InvalidFormat(
                "master document exceeds 65536 body elements".to_string(),
            ));
        }
        self.elements.push(element);
        Ok(self)
    }

    pub fn add_paragraph(&mut self, text: impl Into<String>) -> Result<&mut Self> {
        self.push(MasterDocumentElement::Paragraph(text.into()))
    }

    pub fn add_section(&mut self, section: MasterSection) -> Result<&mut Self> {
        self.push(MasterDocumentElement::Section(section))
    }

    pub fn add_index(&mut self, index: TextIndex) -> Result<&mut Self> {
        self.push(MasterDocumentElement::Index(index))
    }

    pub fn add_subdocument(&mut self, subdocument: MasterSubdocument) -> Result<&mut Self> {
        subdocument.validate()?;
        self.push(MasterDocumentElement::Subdocument(subdocument))
    }

    pub fn move_element(&mut self, from: usize, to: usize) -> Result<&mut Self> {
        if from >= self.elements.len() || to >= self.elements.len() {
            return Err(Error::InvalidFormat(
                "master body reorder index is out of bounds".to_string(),
            ));
        }
        if from != to {
            let element = self.elements.remove(from);
            self.elements.insert(to, element);
        }
        Ok(self)
    }

    pub fn set_styles_xml(&mut self, xml: Option<String>) -> Result<&mut Self> {
        validate_xml_part(xml.as_deref(), "styles.xml")?;
        self.styles_xml = xml;
        Ok(self)
    }

    pub fn set_meta_xml(&mut self, xml: Option<String>) -> Result<&mut Self> {
        validate_xml_part(xml.as_deref(), "meta.xml")?;
        self.meta_xml = xml;
        Ok(self)
    }

    pub fn set_settings_xml(&mut self, xml: Option<String>) -> Result<&mut Self> {
        validate_xml_part(xml.as_deref(), "settings.xml")?;
        self.settings_xml = xml;
        Ok(self)
    }

    pub fn add_auxiliary_file(
        &mut self,
        path: impl Into<String>,
        bytes: Vec<u8>,
        media_type: Option<String>,
    ) -> Result<&mut Self> {
        let path = path.into();
        validate_auxiliary_path(&path)?;
        if self.auxiliary.contains_key(&path) {
            return Err(Error::InvalidFormat(format!(
                "duplicate master package path '{path}'"
            )));
        }
        self.auxiliary.insert(path, (bytes, media_type));
        Ok(self)
    }

    pub fn set_encryption(
        &mut self,
        password: impl Into<String>,
        profile: OdfEncryptionProfile,
    ) -> Result<&mut Self> {
        let password = password.into();
        if password.is_empty() {
            return Err(Error::InvalidFormat(
                "master package encryption password cannot be empty".to_string(),
            ));
        }
        self.encryption = Some((password, profile));
        Ok(self)
    }

    pub fn set_document_signer(&mut self, signer: OdfDocumentSigner) -> &mut Self {
        self.signer = Some(signer);
        self
    }

    pub fn build(self) -> Result<Vec<u8>> {
        let content = self.content_xml()?;
        write_master_package(
            &self.mimetype,
            &content,
            self.styles_xml.as_deref(),
            self.meta_xml.as_deref(),
            self.settings_xml.as_deref(),
            &self.auxiliary,
            None,
            self.encryption,
            self.signer,
        )
    }

    pub fn save(self, path: impl AsRef<Path>) -> Result<()> {
        std::fs::write(path, self.build()?)?;
        Ok(())
    }

    fn content_xml(&self) -> Result<String> {
        let mut body = String::new();
        let mut count = 0usize;
        let mut names = HashSet::new();
        let mut ids = HashSet::new();
        for element in &self.elements {
            write_element(element, 0, &mut count, &mut names, &mut ids, &mut body)?;
        }
        let global = self
            .global
            .map(|value| format!(" text:global=\"{}\"", if value { "true" } else { "false" }))
            .unwrap_or_default();
        let xml = format!(
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?><office:document-content xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" office:version=\"1.3\"><office:automatic-styles/><office:body><office:text{global}>{body}</office:text></office:body></office:document-content>"
        );
        validate_master_content_xml(&xml)?;
        Ok(xml)
    }
}

impl Default for MasterDocumentBuilder {
    fn default() -> Self {
        Self::new()
    }
}

fn write_element(
    element: &MasterDocumentElement,
    depth: usize,
    count: &mut usize,
    names: &mut HashSet<String>,
    ids: &mut HashSet<String>,
    output: &mut String,
) -> Result<()> {
    if depth >= MAX_DEPTH || *count >= MAX_ELEMENTS {
        return Err(Error::InvalidFormat(
            "master body nesting or element count exceeds its limit".to_string(),
        ));
    }
    *count += 1;
    match element {
        MasterDocumentElement::Paragraph(text) => {
            output.push_str("<text:p>");
            output.push_str(&escape_xml(text));
            output.push_str("</text:p>");
        },
        MasterDocumentElement::Index(index) => output.push_str(&index.to_xml_fragment()?),
        MasterDocumentElement::Subdocument(subdocument) => {
            subdocument.validate()?;
            unique_section(&subdocument.to_section(), names, ids)?;
            output.push_str(&subdocument.to_section().to_xml_fragment()?);
        },
        MasterDocumentElement::Section(section) => {
            if section.section.source.is_some() || section.section.dde_source.is_some() {
                return Err(Error::InvalidFormat(
                    "local master section cannot have a source".to_string(),
                ));
            }
            section.section.validate()?;
            unique_section(&section.section, names, ids)?;
            let shell = section.section.to_xml_fragment()?;
            let open_end = shell.find('>').ok_or_else(|| {
                Error::InvalidFormat("invalid canonical section fragment".to_string())
            })?;
            let close = shell.rfind("</text:section>").ok_or_else(|| {
                Error::InvalidFormat("invalid canonical section close".to_string())
            })?;
            output.push_str(&shell[..=open_end]);
            output.push_str(&shell[open_end + 1..close]);
            for child in &section.children {
                write_element(child, depth + 1, count, names, ids, output)?;
            }
            output.push_str("</text:section>");
        },
    }
    Ok(())
}

fn unique_section(
    section: &Section,
    names: &mut HashSet<String>,
    ids: &mut HashSet<String>,
) -> Result<()> {
    if !names.insert(section.name.clone()) {
        return Err(Error::InvalidFormat(format!(
            "duplicate master section name '{}'",
            section.name
        )));
    }
    if let Some(id) = &section.xml_id
        && !ids.insert(id.clone())
    {
        return Err(Error::InvalidFormat(format!(
            "duplicate master xml:id '{id}'"
        )));
    }
    Ok(())
}

pub(crate) fn validate_xml_part(xml: Option<&str>, name: &str) -> Result<()> {
    let Some(xml) = xml else { return Ok(()); };
    if xml.len() > 64 * 1024 * 1024 {
        return Err(Error::InvalidFormat(format!("{name} exceeds 64 MiB")));
    }
    let mut reader = Reader::from_str(xml);
    let mut buffer = Vec::new();
    let mut roots = 0usize;
    loop {
        match reader.read_event_into(&mut buffer).map_err(|error| {
            Error::InvalidFormat(format!("invalid {name}: {error}"))
        })? {
            Event::Start(_) => roots += usize::from(reader.buffer_position() > 0 && roots == 0),
            Event::Empty(_) if roots == 0 => roots = 1,
            Event::DocType(_) | Event::PI(_) => {
                return Err(Error::InvalidFormat(format!(
                    "active XML declarations are prohibited in {name}"
                )));
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if roots == 0 {
        return Err(Error::InvalidFormat(format!("{name} has no root element")));
    }
    Ok(())
}

fn validate_auxiliary_path(path: &str) -> Result<()> {
    if path.is_empty()
        || path.starts_with('/')
        || path.ends_with('/')
        || path.split('/').any(|part| part.is_empty() || part == "." || part == "..")
        || matches!(path, "mimetype" | "content.xml" | "styles.xml" | "meta.xml" | "settings.xml")
    {
        return Err(Error::InvalidFormat(format!(
            "unsafe or reserved master package path '{path}'"
        )));
    }
    Ok(())
}

#[allow(clippy::too_many_arguments)]
pub(crate) fn write_master_package(
    mimetype: &str,
    content: &str,
    styles: Option<&str>,
    meta: Option<&str>,
    settings: Option<&str>,
    auxiliary: &BTreeMap<String, (Vec<u8>, Option<String>)>,
    source: Option<&crate::OwnedPackage>,
    encryption: Option<(String, OdfEncryptionProfile)>,
    signer: Option<OdfDocumentSigner>,
) -> Result<Vec<u8>> {
    if !matches!(mimetype, constants::ODF_MASTER | constants::ODF_MASTER_TEMPLATE) {
        return Err(Error::InvalidFormat(
            "invalid master document MIME type".to_string(),
        ));
    }
    validate_master_content_xml(content)?;
    validate_xml_part(styles, "styles.xml")?;
    validate_xml_part(meta, "meta.xml")?;
    validate_xml_part(settings, "settings.xml")?;
    let encrypted = encryption.is_some();
    let mut writer = PackageWriter::new();
    writer.set_mimetype(mimetype)?;
    if let Some((password, profile)) = encryption {
        writer.set_encryption(password, profile)?;
    }
    if let Some(signer) = signer {
        writer.set_document_signer(signer)?;
    }
    writer.add_file(constants::ODF_CONTENT, content.as_bytes())?;
    let default_styles;
    let styles = match styles {
        Some(value) => value,
        None => {
            default_styles = OdfStructure::default_styles_xml();
            &default_styles
        },
    };
    writer.add_file(constants::ODF_STYLES, styles.as_bytes())?;
    let default_meta = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><office:document-meta xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:meta=\"urn:oasis:names:tc:opendocument:xmlns:meta:1.0\" office:version=\"1.3\"><office:meta><meta:generator>Litchi/0.0.1</meta:generator></office:meta></office:document-meta>";
    writer.add_file(constants::ODF_META, meta.unwrap_or(default_meta).as_bytes())?;
    if let Some(settings) = settings {
        writer.add_file(constants::ODF_SETTINGS, settings.as_bytes())?;
    }
    for (path, (bytes, media_type)) in auxiliary {
        validate_auxiliary_path(path)?;
        match media_type {
            Some(media_type) => writer.add_file_with_media_type(path, bytes, media_type)?,
            None => writer.add_file(path, bytes)?,
        }
    }
    if let Some(source) = source {
        writer.copy_auxiliary_files_from_except(
            source,
            &[constants::ODF_SETTINGS.to_string()],
            &[],
        )?;
    }
    let bytes = writer.finish_to_bytes()?;
    if !encrypted {
        MasterDocument::from_bytes(bytes.clone())?;
    }
    Ok(bytes)
}
