//! Packaged-document snapshot and transactional package edits.

use super::codec::{classify_mimetype, decode_xml_part};
use super::model::{Family, Package};
use crate::constants;
use crate::core::{Meta, OwnedPackage};
use litchi_core::{Error, Metadata, Result};
use std::io::Read;
use std::path::Path;

impl Package {
    pub fn owned_package(&self) -> &OwnedPackage {
        &self.package
    }

    /// Stage a `content.xml` replacement without mutating this package.
    ///
    /// Optional core parts and reproducible auxiliary entries are preserved;
    /// invalidated signatures are omitted and encrypted packages are rejected
    /// by the package writer.
    pub(crate) fn with_replaced_content_xml(&self, content: String) -> Result<Self> {
        std::str::from_utf8(content.as_bytes())
            .map_err(|_error| Error::InvalidFormat("invalid UTF-8 in content.xml".to_string()))?;

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
        Package::from_bytes(writer.finish_to_bytes()?)
    }

    /// Replace `content.xml` while preserving optional core parts and every
    /// auxiliary package entry that can be rewritten safely.
    pub(crate) fn replace_content_xml(&mut self, content: String) -> Result<()> {
        *self = self.with_replaced_content_xml(content)?;
        Ok(())
    }

    /// Open and validate a packaged `OpenDocument` file.
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        let file = std::fs::File::open(path)?;
        Self::from_reader(file)
    }

    /// Open and validate a password-encrypted packaged `OpenDocument` file.
    pub fn open_with_password(path: impl AsRef<Path>, password: impl Into<String>) -> Result<Self> {
        let file = std::fs::File::open(path)?;
        Self::from_reader_with_password(file, password)
    }

    /// Read and validate a packaged `OpenDocument` file.
    pub fn from_reader(reader: impl Read) -> Result<Self> {
        let package = OwnedPackage::from_reader(reader)?;
        Self::from_owned(package)
    }

    /// Read and validate a password-encrypted packaged `OpenDocument` file.
    pub fn from_reader_with_password(
        reader: impl Read,
        password: impl Into<String>,
    ) -> Result<Self> {
        Self::from_owned(OwnedPackage::from_reader_with_password(reader, password)?)
    }

    /// Validate a packaged `OpenDocument` file from owned bytes.
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        let package = OwnedPackage::from_bytes(bytes)?;
        Self::from_owned(package)
    }

    /// Validate password-encrypted packaged `OpenDocument` bytes.
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
            .map_err(|_error| Error::InvalidFormat("invalid UTF-8 in content.xml".to_string()))?;

        Ok(Self {
            package,
            family,
            template,
            mimetype,
        })
    }

    /// Return the standard document family.
    pub fn family(&self) -> Family {
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
    pub fn settings(&self) -> Result<Option<crate::Settings>> {
        let Some(xml) = self.settings_xml()? else {
            return Ok(None);
        };
        crate::settings::parse_package(&xml).map(Some)
    }

    /// Extract common document metadata, or an empty value when `meta.xml` is absent.
    pub fn metadata(&self) -> Result<Metadata> {
        let Some(xml) = self.optional_xml_part(constants::ODF_META)? else {
            return Ok(Metadata::default());
        };
        Meta::from_bytes(xml.as_bytes())?.try_extract_metadata()
    }

    /// Extract the complete format-specific metadata model, if `meta.xml` exists.
    pub fn odf_metadata(&self) -> Result<Option<crate::Metadata>> {
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
    pub fn images(&self) -> Result<Vec<crate::Image>> {
        let content = self.content_xml()?;
        let styles = self.styles_xml()?;
        let package = self.package.package()?;
        crate::media::scan_package(&content, styles.as_deref(), &package)
    }

    /// Inspect classic forms in content and styles without executing behavior.
    pub fn forms(&self) -> Result<crate::form::Forms> {
        let content = self.content_xml()?;
        let styles = self.styles_xml()?;
        let mut parts = vec![(content.as_str(), crate::form::Part::Content)];
        if let Some(styles) = styles.as_deref() {
            parts.push((styles, crate::form::Part::Styles));
        }
        crate::form::parse_form_parts(&parts)
    }

    /// Inspect ordered ODF variable declarations in content and styles.
    pub fn variable_declarations(&self) -> Result<crate::variable_declaration::Declarations> {
        let content = self.content_xml()?;
        let styles = self.styles_xml()?;
        let mut parts = vec![(content.as_str(), crate::variable_declaration::Part::Content)];
        if let Some(styles) = styles.as_deref() {
            parts.push((styles, crate::variable_declaration::Part::Styles));
        }
        crate::variable_declaration::parse_parts(&parts)
    }

    /// Atomically insert or replace one variable declaration container.
    ///
    /// The group must target `content.xml` or `styles.xml`. Auxiliary package
    /// entries remain byte-for-byte intact, and formulas remain inert.
    pub fn set_variable_declaration_group(
        &mut self,
        group: &crate::variable_declaration::Group,
    ) -> Result<Option<crate::variable_declaration::Group>> {
        if group.part == crate::variable_declaration::Part::Flat {
            return Err(Error::InvalidFormat(
                "Package cannot write Part::Flat".to_string(),
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
            crate::variable_declaration::Part::Content => (
                crate::variable_declaration::set_xml(&content, group)?,
                styles,
            ),
            crate::variable_declaration::Part::Styles => {
                let styles = styles.ok_or_else(|| {
                    Error::InvalidFormat(
                        "cannot write a styles declaration without styles.xml".to_string(),
                    )
                })?;
                (
                    content,
                    Some(crate::variable_declaration::set_xml(&styles, group)?),
                )
            },
            crate::variable_declaration::Part::Flat => unreachable!(),
        };

        self.replace_variable_xml(content, styles, old, group.part, &group.scope, group.kind)
    }

    /// Atomically remove one variable declaration container.
    ///
    /// Removal fails without mutation if any remaining field references a
    /// declaration owned by the container.
    pub fn remove_variable_declaration_group(
        &mut self,
        part: crate::variable_declaration::Part,
        scope: &crate::variable_declaration::Scope,
        kind: crate::variable_declaration::Kind,
    ) -> Result<Option<crate::variable_declaration::Group>> {
        if part == crate::variable_declaration::Part::Flat {
            return Err(Error::InvalidFormat(
                "Package cannot remove Part::Flat".to_string(),
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
            crate::variable_declaration::Part::Content => (
                crate::variable_declaration::remove_xml(&content, scope, kind)?,
                styles,
            ),
            crate::variable_declaration::Part::Styles => {
                let styles = styles.ok_or_else(|| {
                    Error::InvalidFormat(
                        "cannot remove a styles declaration without styles.xml".to_string(),
                    )
                })?;
                (
                    content,
                    Some(crate::variable_declaration::remove_xml(
                        &styles, scope, kind,
                    )?),
                )
            },
            crate::variable_declaration::Part::Flat => unreachable!(),
        };

        self.replace_variable_xml(content, styles, Some(old), part, scope, kind)
    }

    fn replace_variable_xml(
        &mut self,
        content: String,
        styles: Option<String>,
        old: Option<crate::variable_declaration::Group>,
        changed_part: crate::variable_declaration::Part,
        scope: &crate::variable_declaration::Scope,
        kind: crate::variable_declaration::Kind,
    ) -> Result<Option<crate::variable_declaration::Group>> {
        let mut parts = vec![(content.as_str(), crate::variable_declaration::Part::Content)];
        if let Some(styles) = styles.as_deref() {
            parts.push((styles, crate::variable_declaration::Part::Styles));
        }
        crate::variable_declaration::parse_parts(&parts)?;

        let mut writer = crate::core::PackageWriter::new();
        writer.set_mimetype(&self.mimetype)?;
        if changed_part == crate::variable_declaration::Part::Content {
            crate::variable_declaration::splice_publication(
                &self.package,
                constants::ODF_CONTENT,
                &content,
                scope,
                kind,
            )?
            .publish(&mut writer)?;
        } else {
            crate::core::XmlSplicePublication::new(crate::core::XmlSourcePart::load(
                &self.package,
                constants::ODF_CONTENT,
            )?)
            .publish(&mut writer)?;
        }
        if let Some(styles) = styles.as_deref() {
            if changed_part == crate::variable_declaration::Part::Styles {
                crate::variable_declaration::splice_publication(
                    &self.package,
                    constants::ODF_STYLES,
                    styles,
                    scope,
                    kind,
                )?
                .publish(&mut writer)?;
            } else {
                crate::core::XmlSplicePublication::new(crate::core::XmlSourcePart::load(
                    &self.package,
                    constants::ODF_STYLES,
                )?)
                .publish(&mut writer)?;
            }
        }
        if self.package.has_file(constants::ODF_META)? {
            crate::core::XmlSplicePublication::new(crate::core::XmlSourcePart::load(
                &self.package,
                constants::ODF_META,
            )?)
            .publish(&mut writer)?;
        }
        writer.copy_auxiliary_files_from(&self.package)?;
        let replacement = Package::from_bytes(writer.finish_to_bytes()?)?;
        *self = replacement;
        Ok(old)
    }

    /// Discover package, inline, missing, and inert linked embedded objects.
    pub fn embedded_objects(&self) -> Result<Vec<crate::Object>> {
        let content = self.content_xml()?;
        let styles = self.styles_xml()?;
        let package = self.package.package()?;
        crate::embedded::scan_package(&content, styles.as_deref(), &package)
    }

    /// Return bytes only for inline or verified package-contained images.
    /// Linked images remain inert and are never fetched.
    pub fn image_bytes(&self, image: &crate::Image) -> Result<Option<Vec<u8>>> {
        match &image.source {
            litchi_odf_common::media::Source::Inline { bytes, .. } => Ok(Some(bytes.clone())),
            litchi_odf_common::media::Source::PackagePart { path, .. } => {
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
