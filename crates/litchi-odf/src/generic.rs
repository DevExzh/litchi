//! Format-neutral access to any packaged OpenDocument document family.

use crate::constants;
use crate::core::{Meta, OwnedPackage};
use litchi_core::{Error, Metadata, Result};
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::io::Read;
use std::path::Path;

/// Standard packaged OpenDocument document family.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum OpenDocumentFamily {
    /// Text document or template.
    Text,
    /// Spreadsheet document or template.
    Spreadsheet,
    /// Presentation document or template.
    Presentation,
    /// Drawing document or template.
    Drawing,
    /// Standalone chart document or template.
    Chart,
    /// Mathematical formula document or template.
    Formula,
    /// Image document or template.
    Image,
    /// Text master document or template.
    Master,
    /// Legacy producer-specific web-oriented text template.
    Web,
    /// Database front-end document.
    Database,
}

/// Validated, format-neutral OpenDocument package.
///
/// This provides lossless package access for every standard OpenDocument
/// family, including document types that do not yet have a specialized object
/// model. Saving an unmodified package returns the original bytes exactly.
pub struct OpenDocumentPackage {
    package: OwnedPackage,
    family: OpenDocumentFamily,
    template: bool,
    mimetype: String,
}

/// Validated flat OpenDocument XML file.
///
/// Flat documents combine content, styles, settings, and metadata under one
/// `office:document` root and are conventionally stored as `.fodt`, `.fods`,
/// `.fodp`, `.fodg`, `.fodc`, or `.fodi`. The `.fodf` extension is also
/// accepted for compatibility with odfdo's non-standard `office:formula`
/// convention; conforming packaged `.odf` formulas use a direct MathML root.
pub struct FlatOpenDocument {
    xml: String,
    family: OpenDocumentFamily,
    mimetype: String,
}

impl FlatOpenDocument {
    /// Parses the optional flat-document `office:settings` inventory.
    pub fn settings(&self) -> Result<crate::OdfSettings> {
        crate::settings::parse_settings(self.xml(), crate::settings::SettingsDocumentKind::Flat)
    }

    /// Open and validate a flat OpenDocument XML file.
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        let file = std::fs::File::open(path)?;
        Self::from_reader(file)
    }

    /// Read and validate a flat OpenDocument XML stream.
    pub fn from_reader(mut reader: impl Read) -> Result<Self> {
        let mut bytes = Vec::new();
        reader.read_to_end(&mut bytes)?;
        Self::from_bytes(bytes)
    }

    /// Validate flat OpenDocument XML from owned bytes.
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        let mimetype = crate::detect::flat_mime(&bytes)
            .ok_or_else(|| Error::InvalidFormat("invalid flat OpenDocument root".to_string()))?;
        let (family, template) = classify_mimetype(&mimetype).ok_or_else(|| {
            Error::InvalidFormat(format!("unsupported OpenDocument mimetype '{mimetype}'"))
        })?;
        if template
            || matches!(
                family,
                OpenDocumentFamily::Master | OpenDocumentFamily::Web | OpenDocumentFamily::Database
            )
        {
            return Err(Error::InvalidFormat(format!(
                "mimetype '{mimetype}' has no standard flat OpenDocument form"
            )));
        }
        let xml = String::from_utf8(bytes)
            .map_err(|_| Error::InvalidFormat("invalid UTF-8 in flat OpenDocument".to_string()))?;
        validate_flat_document(&xml, family)?;
        Ok(Self {
            xml,
            family,
            mimetype,
        })
    }

    /// Return the document family.
    pub fn family(&self) -> OpenDocumentFamily {
        self.family
    }

    /// Return the root `office:mimetype` value.
    pub fn mimetype(&self) -> &str {
        &self.mimetype
    }

    /// Return the conventional flat OpenDocument extension.
    pub fn extension(&self) -> &'static str {
        match self.family {
            OpenDocumentFamily::Text => "fodt",
            OpenDocumentFamily::Spreadsheet => "fods",
            OpenDocumentFamily::Presentation => "fodp",
            OpenDocumentFamily::Drawing => "fodg",
            OpenDocumentFamily::Chart => "fodc",
            OpenDocumentFamily::Formula => "fodf",
            OpenDocumentFamily::Image => "fodi",
            OpenDocumentFamily::Master | OpenDocumentFamily::Web | OpenDocumentFamily::Database => {
                unreachable!("master and web flat documents are rejected")
            },
        }
    }

    /// Return the complete flat XML document.
    pub fn xml(&self) -> &str {
        &self.xml
    }

    /// Extract common document metadata from the combined XML document.
    pub fn metadata(&self) -> Result<Metadata> {
        Meta::from_bytes(self.xml.as_bytes())?.try_extract_metadata()
    }

    /// Extract the complete format-specific metadata model.
    pub fn odf_metadata(&self) -> Result<crate::OdfMetadata> {
        Meta::from_bytes(self.xml.as_bytes())?.odf_metadata()
    }

    /// Discover inline and inert linked images in the flat document.
    pub fn images(&self) -> Result<Vec<crate::OdfImage>> {
        crate::media::scan_flat_images(&self.xml)
    }

    /// Inspect classic forms without executing bindings, events, or external resources.
    pub fn forms(&self) -> Result<crate::OdfForms> {
        crate::form::parse_form_parts(&[(self.xml(), crate::OdfFormPart::Flat)])
    }

    /// Inspect ordered ODF variable declarations without evaluating fields or formulas.
    pub fn variable_declarations(&self) -> Result<crate::OdfVariableDeclarations> {
        crate::variable_declaration::parse_variable_declaration_parts(&[(
            self.xml(),
            crate::OdfVariablePart::Flat,
        )])
    }

    /// Atomically insert or replace one variable declaration container.
    ///
    /// The group must target the flat part. Formulas and cached values remain
    /// inert; this method only updates XML metadata and never evaluates fields.
    pub fn set_variable_declaration_group(
        &mut self,
        group: &crate::OdfVariableDeclarationGroup,
    ) -> Result<Option<crate::OdfVariableDeclarationGroup>> {
        if group.part != crate::OdfVariablePart::Flat {
            return Err(Error::InvalidFormat(
                "FlatOpenDocument requires OdfVariablePart::Flat".to_string(),
            ));
        }
        let current = self.variable_declarations()?;
        let old = current
            .groups
            .iter()
            .find(|candidate| candidate.scope == group.scope && candidate.kind == group.kind)
            .cloned();
        let updated = crate::set_variable_declaration_group_xml(&self.xml, group)?;
        validate_flat_document(&updated, self.family)?;
        crate::variable_declaration::parse_variable_declaration_parts(&[(
            updated.as_str(),
            crate::OdfVariablePart::Flat,
        )])?;
        self.xml = updated;
        Ok(old)
    }

    /// Atomically remove one variable declaration container.
    ///
    /// Removal fails without mutation if any remaining field references a
    /// declaration owned by the container.
    pub fn remove_variable_declaration_group(
        &mut self,
        scope: &crate::OdfVariableScope,
        kind: crate::OdfVariableKind,
    ) -> Result<Option<crate::OdfVariableDeclarationGroup>> {
        let current = self.variable_declarations()?;
        let Some(old) = current
            .groups
            .iter()
            .find(|candidate| candidate.scope == *scope && candidate.kind == kind)
            .cloned()
        else {
            return Ok(None);
        };
        let updated = crate::remove_variable_declaration_group_xml(&self.xml, scope, kind)?;
        validate_flat_document(&updated, self.family)?;
        crate::variable_declaration::parse_variable_declaration_parts(&[(
            updated.as_str(),
            crate::OdfVariablePart::Flat,
        )])?;
        self.xml = updated;
        Ok(Some(old))
    }

    /// Discover inert inline and linked embedded objects.
    pub fn embedded_objects(&self) -> Result<Vec<crate::OdfEmbeddedObject>> {
        crate::embedded_object::scan_flat_objects(&self.xml)
    }

    /// Return the exact original bytes.
    pub fn as_bytes(&self) -> &[u8] {
        self.xml.as_bytes()
    }

    /// Clone the exact original bytes.
    pub fn to_bytes(&self) -> Vec<u8> {
        self.as_bytes().to_vec()
    }

    /// Consume this wrapper and return the exact original bytes.
    pub fn into_bytes(self) -> Vec<u8> {
        self.xml.into_bytes()
    }

    /// Save the flat document without reconstructing its XML.
    pub fn save(&self, path: impl AsRef<Path>) -> Result<()> {
        std::fs::write(path, self.as_bytes())?;
        Ok(())
    }
}

fn validate_flat_document(xml: &str, family: OpenDocumentFamily) -> Result<()> {
    const OFFICE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";

    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut body_seen = false;
    let mut family_body_seen = false;
    let expected_body = match family {
        OpenDocumentFamily::Text => b"text".as_slice(),
        OpenDocumentFamily::Spreadsheet => b"spreadsheet".as_slice(),
        OpenDocumentFamily::Presentation => b"presentation".as_slice(),
        OpenDocumentFamily::Drawing => b"drawing".as_slice(),
        OpenDocumentFamily::Chart => b"chart".as_slice(),
        OpenDocumentFamily::Formula => b"formula".as_slice(),
        OpenDocumentFamily::Image => b"image".as_slice(),
        OpenDocumentFamily::Master | OpenDocumentFamily::Web | OpenDocumentFamily::Database => {
            unreachable!()
        },
    };
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid flat OpenDocument XML: {error}"))
            })?;
        match event {
            Event::Start(element) => {
                if depth == 0 {
                    if root_seen
                        || !matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == OFFICE_NAMESPACE)
                        || element.local_name().as_ref() != b"document"
                    {
                        return Err(Error::InvalidFormat(
                            "flat OpenDocument must contain one office:document root".to_string(),
                        ));
                    }
                    root_seen = true;
                } else if depth == 1
                    && matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == OFFICE_NAMESPACE)
                    && element.local_name().as_ref() == b"body"
                {
                    body_seen = true;
                } else if depth == 2
                    && matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == OFFICE_NAMESPACE)
                    && element.local_name().as_ref() == expected_body
                {
                    family_body_seen = true;
                }
                depth = depth.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("flat OpenDocument nesting overflow".to_string())
                })?;
            },
            Event::Empty(element) => {
                if depth == 0 {
                    return Err(Error::InvalidFormat(
                        "flat OpenDocument root cannot be empty".to_string(),
                    ));
                }
                if depth == 1
                    && matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == OFFICE_NAMESPACE)
                    && element.local_name().as_ref() == b"body"
                {
                    body_seen = true;
                } else if depth == 2
                    && matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == OFFICE_NAMESPACE)
                    && element.local_name().as_ref() == expected_body
                {
                    family_body_seen = true;
                }
            },
            Event::End(_) => {
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("unexpected flat OpenDocument closing tag".to_string())
                })?;
                if depth == 0 {
                    root_closed = true;
                }
            },
            Event::Text(text) if depth == 0 && !text.iter().all(u8::is_ascii_whitespace) => {
                return Err(Error::InvalidFormat(
                    "text is not allowed outside the flat OpenDocument root".to_string(),
                ));
            },
            Event::CData(_) | Event::GeneralRef(_) if depth == 0 => {
                return Err(Error::InvalidFormat(
                    "content is not allowed outside the flat OpenDocument root".to_string(),
                ));
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if !root_seen || !root_closed || depth != 0 || !body_seen || !family_body_seen {
        return Err(Error::InvalidFormat(
            "flat OpenDocument is missing its complete family-specific body".to_string(),
        ));
    }
    Ok(())
}

impl OpenDocumentPackage {
    pub(crate) fn owned_package(&self) -> &OwnedPackage {
        &self.package
    }

    /// Stage a `content.xml` replacement without mutating this package.
    ///
    /// Optional core parts and reproducible auxiliary entries are preserved;
    /// invalidated signatures are omitted and encrypted packages are rejected
    /// by the package writer.
    pub(crate) fn with_replaced_content_xml(&self, content: String) -> Result<Self> {
        std::str::from_utf8(content.as_bytes())
            .map_err(|_| Error::InvalidFormat("invalid UTF-8 in content.xml".to_string()))?;

        let mut writer = crate::core::PackageWriter::new();
        writer.set_mimetype(&self.mimetype)?;
        writer.add_file(constants::ODF_CONTENT, content.as_bytes())?;
        for path in [
            constants::ODF_STYLES,
            constants::ODF_META,
            constants::ODF_SETTINGS,
        ] {
            if self.package.has_file(path)? {
                let bytes = self.package.get_file(path)?;
                writer.add_file(path, &bytes)?;
            }
        }
        writer.copy_auxiliary_files_from_except(
            &self.package,
            &[constants::ODF_SETTINGS.to_string()],
            &[],
        )?;
        OpenDocumentPackage::from_bytes(writer.finish_to_bytes()?)
    }

    /// Replace `content.xml` while preserving optional core parts and every
    /// auxiliary package entry that can be rewritten safely.
    pub(crate) fn replace_content_xml(&mut self, content: String) -> Result<()> {
        *self = self.with_replaced_content_xml(content)?;
        Ok(())
    }

    /// Open and validate a packaged OpenDocument file.
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        let file = std::fs::File::open(path)?;
        Self::from_reader(file)
    }

    /// Open and validate a password-encrypted packaged OpenDocument file.
    pub fn open_with_password(path: impl AsRef<Path>, password: impl Into<String>) -> Result<Self> {
        let file = std::fs::File::open(path)?;
        Self::from_reader_with_password(file, password)
    }

    /// Read and validate a packaged OpenDocument file.
    pub fn from_reader(reader: impl Read) -> Result<Self> {
        let package = OwnedPackage::from_reader(reader)?;
        Self::from_owned(package)
    }

    /// Read and validate a password-encrypted packaged OpenDocument file.
    pub fn from_reader_with_password(
        reader: impl Read,
        password: impl Into<String>,
    ) -> Result<Self> {
        Self::from_owned(OwnedPackage::from_reader_with_password(reader, password)?)
    }

    /// Validate a packaged OpenDocument file from owned bytes.
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        let package = OwnedPackage::from_bytes(bytes)?;
        Self::from_owned(package)
    }

    /// Validate password-encrypted packaged OpenDocument bytes.
    pub fn from_bytes_with_password(bytes: Vec<u8>, password: impl Into<String>) -> Result<Self> {
        Self::from_owned(OwnedPackage::from_bytes_with_password(bytes, password)?)
    }

    fn from_owned(package: OwnedPackage) -> Result<Self> {
        // Constructing the borrowed package validates the mimetype file and
        // META-INF/manifest.xml together.
        let mimetype = package.mimetype()?;
        let (family, template) = classify_mimetype(&mimetype).ok_or_else(|| {
            Error::InvalidFormat(format!("unsupported OpenDocument mimetype '{mimetype}'"))
        })?;
        if !package.has_file(constants::ODF_CONTENT)? {
            return Err(Error::InvalidFormat(
                "OpenDocument package has no content.xml".to_string(),
            ));
        }
        // ODF XML parts are UTF-8 XML. Validate the mandatory part eagerly so
        // callers never receive a nominally valid package with binary content.
        let content = package.get_file(constants::ODF_CONTENT)?;
        std::str::from_utf8(&content)
            .map_err(|_| Error::InvalidFormat("invalid UTF-8 in content.xml".to_string()))?;

        Ok(Self {
            package,
            family,
            template,
            mimetype,
        })
    }

    /// Return the standard document family.
    pub fn family(&self) -> OpenDocumentFamily {
        self.family
    }

    /// Whether this package uses a template MIME type.
    pub fn is_template(&self) -> bool {
        self.template
    }

    /// Whether the package manifest contains password-encrypted entries.
    pub fn is_encrypted(&self) -> Result<bool> {
        Ok(self.package.package()?.manifest().has_encrypted_entries())
    }

    /// Return the exact package MIME type.
    pub fn mimetype(&self) -> &str {
        &self.mimetype
    }

    /// Return the mandatory `content.xml` part.
    pub fn content_xml(&self) -> Result<String> {
        decode_xml_part(
            self.package.get_file(constants::ODF_CONTENT)?,
            constants::ODF_CONTENT,
        )
    }

    /// Return the optional `styles.xml` part.
    pub fn styles_xml(&self) -> Result<Option<String>> {
        self.optional_xml_part(constants::ODF_STYLES)
    }

    /// Return the optional `settings.xml` part.
    pub fn settings_xml(&self) -> Result<Option<String>> {
        self.optional_xml_part(constants::ODF_SETTINGS)
    }

    /// Parses the package settings part without evaluating any stored values.
    pub fn settings(&self) -> Result<Option<crate::OdfSettings>> {
        let Some(xml) = self.settings_xml()? else {
            return Ok(None);
        };
        crate::settings::parse_settings(&xml, crate::settings::SettingsDocumentKind::Package)
            .map(Some)
    }

    /// Extract common document metadata, or an empty value when `meta.xml` is absent.
    pub fn metadata(&self) -> Result<Metadata> {
        let Some(xml) = self.optional_xml_part(constants::ODF_META)? else {
            return Ok(Metadata::default());
        };
        Meta::from_bytes(xml.as_bytes())?.try_extract_metadata()
    }

    /// Extract the complete format-specific metadata model, if `meta.xml` exists.
    pub fn odf_metadata(&self) -> Result<Option<crate::OdfMetadata>> {
        let Some(xml) = self.optional_xml_part(constants::ODF_META)? else {
            return Ok(None);
        };
        Ok(Some(Meta::from_bytes(xml.as_bytes())?.odf_metadata()?))
    }

    fn optional_xml_part(&self, path: &str) -> Result<Option<String>> {
        if !self.package.has_file(path)? {
            return Ok(None);
        }
        self.package
            .get_file(path)
            .and_then(|bytes| decode_xml_part(bytes, path))
            .map(Some)
    }

    /// List every path stored in the package.
    pub fn files(&self) -> Result<Vec<String>> {
        self.package.files()
    }

    /// Return whether a package path exists.
    pub fn has_file(&self, path: &str) -> Result<bool> {
        self.package.has_file(path)
    }

    /// Extract an arbitrary package part without interpreting it.
    pub fn get_file(&self, path: &str) -> Result<Vec<u8>> {
        self.package.get_file(path)
    }

    /// List package paths that appear to contain embedded media.
    pub fn media_files(&self) -> Result<Vec<String>> {
        self.package.media_files()
    }

    /// Discover referenced, inline, missing, and inert linked images.
    pub fn images(&self) -> Result<Vec<crate::OdfImage>> {
        let content = self.content_xml()?;
        let styles = self.styles_xml()?;
        let package = self.package.package()?;
        crate::media::scan_packaged_images(
            &content,
            styles.as_deref(),
            |path| package.has_file(path),
            |path| package.manifest().get_media_type(path).map(str::to_string),
        )
    }

    /// Inspect classic forms in content and styles without executing behavior.
    pub fn forms(&self) -> Result<crate::OdfForms> {
        let content = self.content_xml()?;
        let styles = self.styles_xml()?;
        let mut parts = vec![(content.as_str(), crate::OdfFormPart::Content)];
        if let Some(styles) = styles.as_deref() {
            parts.push((styles, crate::OdfFormPart::Styles));
        }
        crate::form::parse_form_parts(&parts)
    }

    /// Inspect ordered ODF variable declarations in content and styles.
    pub fn variable_declarations(&self) -> Result<crate::OdfVariableDeclarations> {
        let content = self.content_xml()?;
        let styles = self.styles_xml()?;
        let mut parts = vec![(content.as_str(), crate::OdfVariablePart::Content)];
        if let Some(styles) = styles.as_deref() {
            parts.push((styles, crate::OdfVariablePart::Styles));
        }
        crate::variable_declaration::parse_variable_declaration_parts(&parts)
    }

    /// Atomically insert or replace one variable declaration container.
    ///
    /// The group must target `content.xml` or `styles.xml`. Auxiliary package
    /// entries remain byte-for-byte intact, and formulas remain inert.
    pub fn set_variable_declaration_group(
        &mut self,
        group: &crate::OdfVariableDeclarationGroup,
    ) -> Result<Option<crate::OdfVariableDeclarationGroup>> {
        if group.part == crate::OdfVariablePart::Flat {
            return Err(Error::InvalidFormat(
                "OpenDocumentPackage cannot write OdfVariablePart::Flat".to_string(),
            ));
        }

        let current = self.variable_declarations()?;
        let old = current
            .groups
            .iter()
            .find(|candidate| {
                candidate.part == group.part
                    && candidate.scope == group.scope
                    && candidate.kind == group.kind
            })
            .cloned();
        let content = self.content_xml()?;
        let styles = self.styles_xml()?;
        let (content, styles) = match group.part {
            crate::OdfVariablePart::Content => (
                crate::set_variable_declaration_group_xml(&content, group)?,
                styles,
            ),
            crate::OdfVariablePart::Styles => {
                let styles = styles.ok_or_else(|| {
                    Error::InvalidFormat(
                        "cannot write a styles declaration without styles.xml".to_string(),
                    )
                })?;
                (
                    content,
                    Some(crate::set_variable_declaration_group_xml(&styles, group)?),
                )
            },
            crate::OdfVariablePart::Flat => unreachable!(),
        };

        self.replace_variable_xml(content, styles, old)
    }

    /// Atomically remove one variable declaration container.
    ///
    /// Removal fails without mutation if any remaining field references a
    /// declaration owned by the container.
    pub fn remove_variable_declaration_group(
        &mut self,
        part: crate::OdfVariablePart,
        scope: &crate::OdfVariableScope,
        kind: crate::OdfVariableKind,
    ) -> Result<Option<crate::OdfVariableDeclarationGroup>> {
        if part == crate::OdfVariablePart::Flat {
            return Err(Error::InvalidFormat(
                "OpenDocumentPackage cannot remove OdfVariablePart::Flat".to_string(),
            ));
        }

        let current = self.variable_declarations()?;
        let Some(old) = current
            .groups
            .iter()
            .find(|candidate| {
                candidate.part == part && candidate.scope == *scope && candidate.kind == kind
            })
            .cloned()
        else {
            return Ok(None);
        };
        let content = self.content_xml()?;
        let styles = self.styles_xml()?;
        let (content, styles) = match part {
            crate::OdfVariablePart::Content => (
                crate::remove_variable_declaration_group_xml(&content, scope, kind)?,
                styles,
            ),
            crate::OdfVariablePart::Styles => {
                let styles = styles.ok_or_else(|| {
                    Error::InvalidFormat(
                        "cannot remove a styles declaration without styles.xml".to_string(),
                    )
                })?;
                (
                    content,
                    Some(crate::remove_variable_declaration_group_xml(
                        &styles, scope, kind,
                    )?),
                )
            },
            crate::OdfVariablePart::Flat => unreachable!(),
        };

        self.replace_variable_xml(content, styles, Some(old))
    }

    fn replace_variable_xml(
        &mut self,
        content: String,
        styles: Option<String>,
        old: Option<crate::OdfVariableDeclarationGroup>,
    ) -> Result<Option<crate::OdfVariableDeclarationGroup>> {
        let mut parts = vec![(content.as_str(), crate::OdfVariablePart::Content)];
        if let Some(styles) = styles.as_deref() {
            parts.push((styles, crate::OdfVariablePart::Styles));
        }
        crate::variable_declaration::parse_variable_declaration_parts(&parts)?;

        let mut writer = crate::core::PackageWriter::new();
        writer.set_mimetype(&self.mimetype)?;
        writer.add_file(constants::ODF_CONTENT, content.as_bytes())?;
        if let Some(styles) = styles.as_deref() {
            writer.add_file(constants::ODF_STYLES, styles.as_bytes())?;
        }
        if self.package.has_file(constants::ODF_META)? {
            let bytes = self.package.get_file(constants::ODF_META)?;
            writer.add_file(constants::ODF_META, &bytes)?;
        }
        writer.copy_auxiliary_files_from(&self.package)?;
        let replacement = OpenDocumentPackage::from_bytes(writer.finish_to_bytes()?)?;
        *self = replacement;
        Ok(old)
    }

    /// Discover package, inline, missing, and inert linked embedded objects.
    pub fn embedded_objects(&self) -> Result<Vec<crate::OdfEmbeddedObject>> {
        let content = self.content_xml()?;
        let styles = self.styles_xml()?;
        let package = self.package.package()?;
        crate::embedded_object::scan_packaged_objects(
            &content,
            styles.as_deref(),
            |path| package.has_file(path),
            |path| package.manifest().get_media_type(path).map(str::to_string),
        )
    }

    /// Return bytes only for inline or verified package-contained images.
    /// Linked images remain inert and are never fetched.
    pub fn image_bytes(&self, image: &crate::OdfImage) -> Result<Option<Vec<u8>>> {
        match &image.source {
            crate::OdfImageSource::Inline { bytes, .. } => Ok(Some(bytes.clone())),
            crate::OdfImageSource::PackagePart { path, .. } => {
                self.package.get_file(path).map(Some)
            },
            _ => Ok(None),
        }
    }

    /// Return the original package bytes.
    pub fn as_bytes(&self) -> &[u8] {
        self.package.as_bytes()
    }

    /// Clone the original package bytes for lossless saving or transfer.
    pub fn to_bytes(&self) -> Vec<u8> {
        self.package.as_bytes().to_vec()
    }

    /// Consume this wrapper and return the original package bytes.
    pub fn into_bytes(self) -> Vec<u8> {
        self.package.into_inner()
    }

    /// Save the package without reconstructing or discarding any parts.
    pub fn save(&self, path: impl AsRef<Path>) -> Result<()> {
        std::fs::write(path, self.as_bytes())?;
        Ok(())
    }
}

fn decode_xml_part(bytes: Vec<u8>, path: &str) -> Result<String> {
    String::from_utf8(bytes).map_err(|_| Error::InvalidFormat(format!("invalid UTF-8 in {path}")))
}

fn classify_mimetype(mimetype: &str) -> Option<(OpenDocumentFamily, bool)> {
    Some(match mimetype {
        constants::ODF_TEXT => (OpenDocumentFamily::Text, false),
        constants::ODF_TEXT_TEMPLATE => (OpenDocumentFamily::Text, true),
        constants::ODF_SPREADSHEET => (OpenDocumentFamily::Spreadsheet, false),
        constants::ODF_SPREADSHEET_TEMPLATE => (OpenDocumentFamily::Spreadsheet, true),
        constants::ODF_PRESENTATION => (OpenDocumentFamily::Presentation, false),
        constants::ODF_PRESENTATION_TEMPLATE => (OpenDocumentFamily::Presentation, true),
        constants::ODF_DRAWING => (OpenDocumentFamily::Drawing, false),
        constants::ODF_DRAWING_TEMPLATE => (OpenDocumentFamily::Drawing, true),
        constants::ODF_CHART => (OpenDocumentFamily::Chart, false),
        constants::ODF_CHART_TEMPLATE => (OpenDocumentFamily::Chart, true),
        constants::ODF_FORMULA => (OpenDocumentFamily::Formula, false),
        constants::ODF_FORMULA_TEMPLATE => (OpenDocumentFamily::Formula, true),
        constants::ODF_IMAGE => (OpenDocumentFamily::Image, false),
        constants::ODF_IMAGE_TEMPLATE => (OpenDocumentFamily::Image, true),
        constants::ODF_MASTER => (OpenDocumentFamily::Master, false),
        constants::ODF_MASTER_TEMPLATE => (OpenDocumentFamily::Master, true),
        constants::ODF_WEB => (OpenDocumentFamily::Web, true),
        constants::ODF_DATABASE => (OpenDocumentFamily::Database, false),
        _ => return None,
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::core::PackageWriter;

    fn package(mimetype: &str) -> Vec<u8> {
        let mut writer = PackageWriter::new();
        writer.set_mimetype(mimetype).unwrap();
        writer
            .add_file(
                constants::ODF_CONTENT,
                br#"<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:body/></office:document-content>"#,
            )
            .unwrap();
        writer
            .add_file_with_media_type("Pictures/pixel.png", b"PNG", "image/png")
            .unwrap();
        writer.finish_to_bytes().unwrap()
    }

    #[test]
    fn opens_every_document_family_and_template_losslessly() {
        for (mimetype, family, template) in [
            (constants::ODF_TEXT, OpenDocumentFamily::Text, false),
            (constants::ODF_TEXT_TEMPLATE, OpenDocumentFamily::Text, true),
            (
                constants::ODF_SPREADSHEET,
                OpenDocumentFamily::Spreadsheet,
                false,
            ),
            (
                constants::ODF_SPREADSHEET_TEMPLATE,
                OpenDocumentFamily::Spreadsheet,
                true,
            ),
            (
                constants::ODF_PRESENTATION,
                OpenDocumentFamily::Presentation,
                false,
            ),
            (
                constants::ODF_PRESENTATION_TEMPLATE,
                OpenDocumentFamily::Presentation,
                true,
            ),
            (constants::ODF_DRAWING, OpenDocumentFamily::Drawing, false),
            (
                constants::ODF_DRAWING_TEMPLATE,
                OpenDocumentFamily::Drawing,
                true,
            ),
            (constants::ODF_CHART, OpenDocumentFamily::Chart, false),
            (
                constants::ODF_CHART_TEMPLATE,
                OpenDocumentFamily::Chart,
                true,
            ),
            (constants::ODF_FORMULA, OpenDocumentFamily::Formula, false),
            (
                constants::ODF_FORMULA_TEMPLATE,
                OpenDocumentFamily::Formula,
                true,
            ),
            (constants::ODF_IMAGE, OpenDocumentFamily::Image, false),
            (
                constants::ODF_IMAGE_TEMPLATE,
                OpenDocumentFamily::Image,
                true,
            ),
            (constants::ODF_MASTER, OpenDocumentFamily::Master, false),
            (
                constants::ODF_MASTER_TEMPLATE,
                OpenDocumentFamily::Master,
                true,
            ),
            (constants::ODF_WEB, OpenDocumentFamily::Web, true),
            (constants::ODF_DATABASE, OpenDocumentFamily::Database, false),
        ] {
            let bytes = package(mimetype);
            let document = OpenDocumentPackage::from_bytes(bytes.clone()).unwrap();
            assert_eq!(document.family(), family);
            assert_eq!(document.is_template(), template);
            assert_eq!(document.mimetype(), mimetype);
            assert!(document.content_xml().unwrap().contains("office:body"));
            assert!(document.odf_metadata().unwrap().is_none());
            assert_eq!(document.media_files().unwrap(), ["Pictures/pixel.png"]);
            assert_eq!(document.to_bytes(), bytes);
            assert_eq!(document.into_bytes(), bytes);
        }
    }

    #[test]
    fn rejects_non_odf_missing_content_and_invalid_xml_bytes() {
        let mut writer = PackageWriter::new();
        writer.set_mimetype("application/zip").unwrap();
        writer.add_file(constants::ODF_CONTENT, b"<x/>").unwrap();
        assert!(OpenDocumentPackage::from_bytes(writer.finish_to_bytes().unwrap()).is_err());

        let mut writer = PackageWriter::new();
        writer.set_mimetype(constants::ODF_DRAWING).unwrap();
        assert!(OpenDocumentPackage::from_bytes(writer.finish_to_bytes().unwrap()).is_err());

        let mut writer = PackageWriter::new();
        writer.set_mimetype(constants::ODF_CHART).unwrap();
        writer.add_file(constants::ODF_CONTENT, &[0xff]).unwrap();
        assert!(OpenDocumentPackage::from_bytes(writer.finish_to_bytes().unwrap()).is_err());
    }

    #[test]
    fn opens_standard_and_odfdo_compatible_flat_documents_losslessly() {
        for (mimetype, body, family, extension) in [
            (
                constants::ODF_TEXT,
                "text",
                OpenDocumentFamily::Text,
                "fodt",
            ),
            (
                constants::ODF_SPREADSHEET,
                "spreadsheet",
                OpenDocumentFamily::Spreadsheet,
                "fods",
            ),
            (
                constants::ODF_PRESENTATION,
                "presentation",
                OpenDocumentFamily::Presentation,
                "fodp",
            ),
            (
                constants::ODF_DRAWING,
                "drawing",
                OpenDocumentFamily::Drawing,
                "fodg",
            ),
            (
                constants::ODF_CHART,
                "chart",
                OpenDocumentFamily::Chart,
                "fodc",
            ),
            (
                constants::ODF_FORMULA,
                "formula",
                OpenDocumentFamily::Formula,
                "fodf",
            ),
            (
                constants::ODF_IMAGE,
                "image",
                OpenDocumentFamily::Image,
                "fodi",
            ),
        ] {
            let xml = format!(
                r#"<?xml version="1.0"?><!-- keep --><o:document xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" o:mimetype="{mimetype}" o:version="1.3"><o:body><o:{body}/></o:body></o:document>"#
            );
            let document = FlatOpenDocument::from_bytes(xml.clone().into_bytes()).unwrap();
            assert_eq!(document.family(), family);
            assert_eq!(document.mimetype(), mimetype);
            assert_eq!(document.extension(), extension);
            assert_eq!(document.xml(), xml);
            assert_eq!(document.to_bytes(), xml.as_bytes());
            assert_eq!(document.into_bytes(), xml.into_bytes());
        }
    }

    #[test]
    fn rejects_flat_mimetype_body_mismatch_and_incomplete_xml() {
        for xml in [
            r#"<o:document xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" o:mimetype="application/vnd.oasis.opendocument.text"><o:body><o:spreadsheet/></o:body></o:document>"#,
            r#"<o:document xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" o:mimetype="application/vnd.oasis.opendocument.text"><o:body><o:text/></o:body>"#,
            r#"<o:document xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" o:mimetype="application/vnd.oasis.opendocument.text-template"><o:body><o:text/></o:body></o:document>"#,
            r#"<o:document xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" o:mimetype="application/vnd.oasis.opendocument.text"><o:body><o:text/></o:body></o:document><o:document/>"#,
        ] {
            assert!(
                FlatOpenDocument::from_bytes(xml.as_bytes().to_vec()).is_err(),
                "accepted invalid flat document {xml}"
            );
        }
    }

    #[test]
    fn flat_document_exposes_namespace_aware_metadata() {
        let xml = br#"<o:document xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
            xmlns:d="http://purl.org/dc/elements/1.1/"
            o:mimetype="application/vnd.oasis.opendocument.text">
            <o:meta><d:title>A &amp; B</d:title></o:meta>
            <o:body><o:text/></o:body>
        </o:document>"#;
        let document = FlatOpenDocument::from_bytes(xml.to_vec()).unwrap();
        assert_eq!(
            document.odf_metadata().unwrap().title.as_deref(),
            Some("A & B")
        );
        assert_eq!(document.metadata().unwrap().title.as_deref(), Some("A & B"));
    }

    #[test]
    fn flat_variable_declarations_expand_replace_and_remove_atomically() {
        let xml = format!(
            r#"<o:document xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" o:mimetype="{}"><o:body><o:text/></o:body></o:document>"#,
            constants::ODF_TEXT,
        );
        let mut document = FlatOpenDocument::from_bytes(xml.into_bytes()).unwrap();
        let scope = crate::OdfVariableScope::Body(crate::OdfVariableBody::Text);
        let first = crate::OdfVariableDeclarationGroup {
            kind: crate::OdfVariableKind::Simple,
            part: crate::OdfVariablePart::Flat,
            scope: scope.clone(),
            declarations: vec![crate::OdfVariableDeclaration::Simple {
                name: "counter".to_string(),
                value_type: crate::OdfVariableValueType::Float,
            }],
        };
        assert!(
            document
                .set_variable_declaration_group(&first)
                .unwrap()
                .is_none()
        );
        assert!(document.xml().contains("<o:text><text:variable-decls"));
        assert!(
            document
                .variable_declarations()
                .unwrap()
                .find(crate::OdfVariableKind::Simple, "counter")
                .is_some()
        );

        let second = crate::OdfVariableDeclarationGroup {
            declarations: vec![crate::OdfVariableDeclaration::Simple {
                name: "replacement".to_string(),
                value_type: crate::OdfVariableValueType::String,
            }],
            ..first.clone()
        };
        assert_eq!(
            document.set_variable_declaration_group(&second).unwrap(),
            Some(first.clone()),
        );
        assert!(
            document
                .variable_declarations()
                .unwrap()
                .find(crate::OdfVariableKind::Simple, "replacement")
                .is_some()
        );
        assert_eq!(
            document
                .remove_variable_declaration_group(&scope, crate::OdfVariableKind::Simple)
                .unwrap(),
            Some(second),
        );
        assert!(document.variable_declarations().unwrap().groups.is_empty());
    }
}
