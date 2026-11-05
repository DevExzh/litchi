//! Packaged-document snapshot and transactional package edits.

use super::codec::{classify_mimetype, decode_xml_part};
use super::model::{Family, Package};
use crate::constants;
use crate::core::{Meta, OwnedPackage};
use litchi_core::{Error, Metadata, Result};
use litchi_odf_common::package::replace_content_xml_with_payload_verification;
use litchi_odf_common::{Error as OdfError, XmlStreamLimits};

type FamilyResult<T> = std::result::Result<T, OdfError>;

impl Package {
    pub(crate) fn xml_limits_value(&self) -> XmlStreamLimits {
        self.xml_limits
    }

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

        let bytes = replace_content_xml_with_payload_verification(&self.package, &content)?;
        Package::from_vec_with_xml_limits(bytes, self.xml_limits)
    }

    /// Replace `content.xml` while preserving optional core parts and every
    /// auxiliary package entry that can be rewritten safely.
    pub(crate) fn replace_content_xml(&mut self, content: String) -> Result<()> {
        if self.content_xml()? == content {
            return Ok(());
        }
        *self = self.with_replaced_content_xml(content)?;
        Ok(())
    }

    /// Open and validate a packaged `OpenDocument` document from a path.
    ///
    /// # Errors
    ///
    /// Returns [`litchi_odf_common::Error::Limit`] for an exceeded dimension
    /// and [`litchi_odf_common::Error::Core`] for I/O or validation failures.
    pub fn open(path: impl AsRef<std::path::Path>) -> std::result::Result<Self, OdfError> {
        Self::open_with_options(path, litchi_odf_common::OpenOptions::new())
    }

    /// Open and validate a packaged `OpenDocument` document from a path under caller-selected limits.
    ///
    /// # Errors
    ///
    /// As [`Self::open`].
    pub fn open_with_options(
        path: impl AsRef<std::path::Path>,
        options: litchi_odf_common::OpenOptions,
    ) -> std::result::Result<Self, OdfError> {
        let file = std::fs::File::open(path)?;
        Self::from_reader_with_options(file, options)
    }

    /// Open and validate password-encrypted a packaged `OpenDocument` document from a path.
    ///
    /// # Errors
    ///
    /// As [`Self::open`].
    pub fn open_with_password(
        path: impl AsRef<std::path::Path>,
        password: impl Into<String>,
    ) -> std::result::Result<Self, OdfError> {
        Self::open_with_password_and_options(path, password, litchi_odf_common::OpenOptions::new())
    }

    /// Open and validate password-encrypted a packaged `OpenDocument` document from a path under
    /// caller-selected limits.
    ///
    /// # Errors
    ///
    /// As [`Self::open`].
    pub fn open_with_password_and_options(
        path: impl AsRef<std::path::Path>,
        password: impl Into<String>,
        options: litchi_odf_common::OpenOptions,
    ) -> std::result::Result<Self, OdfError> {
        let password = zeroize::Zeroizing::new(password.into());
        let file = std::fs::File::open(path)?;
        let package =
            OwnedPackage::from_reader_with_password_and_options(file, password.as_str(), options)?;
        Ok(Self::from_owned(package, options.xml_limits_value())?)
    }

    /// Read and validate a packaged `OpenDocument` document from a reader.
    ///
    /// # Errors
    ///
    /// As [`Self::open`].
    pub fn from_reader(reader: impl std::io::Read) -> std::result::Result<Self, OdfError> {
        Self::from_reader_with_options(reader, litchi_odf_common::OpenOptions::new())
    }

    /// Read and validate a packaged `OpenDocument` document from a reader under caller-selected limits.
    ///
    /// # Errors
    ///
    /// As [`Self::open`].
    pub fn from_reader_with_options(
        reader: impl std::io::Read,
        options: litchi_odf_common::OpenOptions,
    ) -> std::result::Result<Self, OdfError> {
        let package = OwnedPackage::from_reader_with_options(reader, options)?;
        Ok(Self::from_owned(package, options.xml_limits_value())?)
    }

    /// Validate a packaged `OpenDocument` document from bytes (the bytes are copied).
    ///
    /// # Errors
    ///
    /// As [`Self::open`].
    pub fn from_bytes(bytes: &[u8]) -> std::result::Result<Self, OdfError> {
        Self::from_bytes_with_options(bytes, litchi_odf_common::OpenOptions::new())
    }

    /// Validate a packaged `OpenDocument` document from bytes under caller-selected limits.
    ///
    /// # Errors
    ///
    /// As [`Self::open`].
    pub fn from_bytes_with_options(
        bytes: &[u8],
        options: litchi_odf_common::OpenOptions,
    ) -> std::result::Result<Self, OdfError> {
        Self::from_reader_with_options(bytes, options)
    }

    /// Validate password-encrypted a packaged `OpenDocument` document from bytes.
    ///
    /// # Errors
    ///
    /// As [`Self::open`].
    pub fn from_bytes_with_password(
        bytes: &[u8],
        password: impl Into<String>,
    ) -> std::result::Result<Self, OdfError> {
        Self::from_bytes_with_password_and_options(
            bytes,
            password,
            litchi_odf_common::OpenOptions::new(),
        )
    }

    /// Validate password-encrypted a packaged `OpenDocument` document from bytes under caller-selected
    /// limits.
    ///
    /// # Errors
    ///
    /// As [`Self::open`].
    pub fn from_bytes_with_password_and_options(
        bytes: &[u8],
        password: impl Into<String>,
        options: litchi_odf_common::OpenOptions,
    ) -> std::result::Result<Self, OdfError> {
        let password = zeroize::Zeroizing::new(password.into());
        let package =
            OwnedPackage::from_reader_with_password_and_options(bytes, password.as_str(), options)?;
        Ok(Self::from_owned(package, options.xml_limits_value())?)
    }

    /// Validate complete package bytes produced inside the crate.
    pub(crate) fn from_vec(bytes: Vec<u8>) -> Result<Self> {
        Self::from_owned(
            OwnedPackage::from_bytes(bytes)?,
            XmlStreamLimits::owner_default_value(),
        )
    }

    fn from_vec_with_xml_limits(bytes: Vec<u8>, xml_limits: XmlStreamLimits) -> Result<Self> {
        Self::from_owned(OwnedPackage::from_bytes(bytes)?, xml_limits)
    }

    fn from_owned(package: OwnedPackage, xml_limits: XmlStreamLimits) -> Result<Self> {
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
            xml_limits,
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
    pub fn is_encrypted(&self) -> FamilyResult<bool> {
        Ok(self.package.package()?.manifest().has_encrypted_entries())
    }

    /// Return the exact package MIME type.
    pub fn mimetype(&self) -> &str {
        &self.mimetype
    }

    /// Return the mandatory `content.xml` part.
    pub fn content_xml(&self) -> FamilyResult<String> {
        Ok(decode_xml_part(
            self.package.get_file(constants::ODF_CONTENT)?,
            constants::ODF_CONTENT,
        )?)
    }

    /// Return the optional `styles.xml` part.
    pub fn styles_xml(&self) -> FamilyResult<Option<String>> {
        Ok(self.optional_xml_part(constants::ODF_STYLES)?)
    }

    /// Return the optional `settings.xml` part.
    pub fn settings_xml(&self) -> FamilyResult<Option<String>> {
        Ok(self.optional_xml_part(constants::ODF_SETTINGS)?)
    }

    /// Parses the package settings part without evaluating any stored values.
    pub fn settings(&self) -> FamilyResult<Option<crate::Settings>> {
        let Some(xml) = self.settings_xml()? else {
            return Ok(None);
        };
        Ok(crate::settings::parse_package(&xml).map(Some)?)
    }

    /// Extract common document metadata, or an empty value when `meta.xml` is absent.
    pub fn metadata(&self) -> FamilyResult<Metadata> {
        let Some(xml) = self.optional_xml_part(constants::ODF_META)? else {
            return Ok(Metadata::default());
        };
        Ok(Meta::from_bytes(xml.as_bytes(), self.xml_limits)?.try_extract_metadata()?)
    }

    /// Extract the complete format-specific metadata model, if `meta.xml` exists.
    pub fn odf_metadata(&self) -> FamilyResult<Option<crate::Metadata>> {
        let Some(xml) = self.optional_xml_part(constants::ODF_META)? else {
            return Ok(None);
        };
        Ok(Some(
            Meta::from_bytes(xml.as_bytes(), self.xml_limits)?.odf_metadata()?,
        ))
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
    pub fn files(&self) -> FamilyResult<Vec<String>> {
        Ok(self.package.files()?)
    }

    /// Return whether a package path exists.
    pub fn has_file(&self, path: &str) -> FamilyResult<bool> {
        Ok(self.package.has_file(path)?)
    }

    /// Extract an arbitrary package part without interpreting it.
    pub fn get_file(&self, path: &str) -> FamilyResult<Vec<u8>> {
        Ok(self.package.get_file(path)?)
    }

    /// List package paths that appear to contain embedded media.
    pub fn media_files(&self) -> FamilyResult<Vec<String>> {
        Ok(self.package.media_files()?)
    }

    /// Discover referenced, inline, missing, and inert linked images.
    pub fn images(&self) -> FamilyResult<Vec<crate::Image>> {
        let content = self.content_xml()?;
        let styles = self.styles_xml()?;
        let package = self.package.package()?;
        Ok(crate::media::scan_package(
            &content,
            styles.as_deref(),
            &package,
        )?)
    }

    /// Inspect classic forms in content and styles without executing behavior.
    pub fn forms(&self) -> FamilyResult<crate::form::Forms> {
        let content = self.content_xml()?;
        let styles = self.styles_xml()?;
        let mut parts = vec![(content.as_str(), crate::form::Part::Content)];
        if let Some(styles) = styles.as_deref() {
            parts.push((styles, crate::form::Part::Styles));
        }
        Ok(crate::form::parse_form_parts(&parts)?)
    }

    /// Inspect in-content RDFa and inline `text:meta` values in content and
    /// styles without materializing unrelated package members.
    pub fn in_content_metadata(&self) -> FamilyResult<crate::ContentMetadata> {
        let content = self.content_xml()?;
        let styles = self.styles_xml()?;
        let mut parts = vec![(content.as_str(), crate::MetadataPart::Content)];
        if let Some(styles) = styles.as_deref() {
            parts.push((styles, crate::MetadataPart::Styles));
        }
        Ok(crate::content_metadata::parse_parts_with_xml_limits(
            &parts,
            self.xml_limits,
        )?)
    }

    /// Inspect inert XForms model declarations in `office:forms`.
    pub fn xforms_models(&self) -> FamilyResult<Vec<crate::xforms::Model>> {
        Ok(crate::xforms::parse_models_with_xml_limits(
            &self.content_xml()?,
            self.xml_limits,
        )?)
    }

    /// Set or clear paragraph RDFa and atomically publish `content.xml`.
    pub(crate) fn set_paragraph_rdfa(
        &mut self,
        position: litchi_core::Position,
        value: &crate::RdfaAttributes,
    ) -> Result<()> {
        let content = crate::content_metadata::set_paragraph_rdfa_with_xml_limits(
            &self.content_xml()?,
            position,
            value,
            self.xml_limits,
        )?;
        self.replace_content_xml(content)
    }

    /// Set or clear RDFa on a uniquely named bookmark start and publish it.
    pub(crate) fn set_bookmark_rdfa(
        &mut self,
        name: &str,
        value: &crate::RdfaAttributes,
    ) -> Result<()> {
        let content = crate::content_metadata::set_bookmark_rdfa_with_xml_limits(
            &self.content_xml()?,
            name,
            value,
            self.xml_limits,
        )?;
        self.replace_content_xml(content)
    }

    /// Set or clear RDFa on one inline `text:meta` occurrence and publish it.
    pub(crate) fn set_text_meta_rdfa(
        &mut self,
        position: litchi_core::Position,
        value: &crate::RdfaAttributes,
    ) -> Result<()> {
        let content = crate::content_metadata::set_text_meta_rdfa_with_xml_limits(
            &self.content_xml()?,
            position,
            value,
            self.xml_limits,
        )?;
        self.replace_content_xml(content)
    }

    /// Insert an inline `text:meta` and atomically publish `content.xml`.
    pub(crate) fn insert_text_meta(
        &mut self,
        paragraph: litchi_core::Position,
        value: &crate::TextMeta,
    ) -> Result<()> {
        let content = crate::content_metadata::insert_text_meta_with_xml_limits(
            &self.content_xml()?,
            paragraph,
            value,
            self.xml_limits,
        )?;
        self.replace_content_xml(content)
    }

    /// Replace one inline `text:meta` and atomically publish `content.xml`.
    pub(crate) fn replace_text_meta(
        &mut self,
        position: litchi_core::Position,
        value: &crate::TextMeta,
    ) -> Result<()> {
        let content = crate::content_metadata::replace_text_meta_with_xml_limits(
            &self.content_xml()?,
            position,
            value,
            self.xml_limits,
        )?;
        self.replace_content_xml(content)
    }

    /// Remove one inline `text:meta` and atomically publish `content.xml`.
    pub(crate) fn remove_text_meta(&mut self, position: litchi_core::Position) -> Result<()> {
        let content = crate::content_metadata::remove_text_meta_with_xml_limits(
            &self.content_xml()?,
            position,
            self.xml_limits,
        )?;
        self.replace_content_xml(content)
    }

    /// Replace one inert XForms model and atomically publish `content.xml`.
    pub(crate) fn replace_xforms_model(
        &mut self,
        position: litchi_core::Position,
        model: &crate::xforms::Model,
    ) -> Result<()> {
        let content = crate::xforms::replace_model_with_xml_limits(
            &self.content_xml()?,
            position.get(),
            model,
            self.xml_limits,
        )?;
        self.replace_content_xml(content)
    }

    /// Insert an inert XForms model into the existing `office:forms` container.
    pub(crate) fn insert_xforms_model(&mut self, model: &crate::xforms::Model) -> Result<()> {
        let content = crate::xforms::insert_model_with_xml_limits(
            &self.content_xml()?,
            model,
            self.xml_limits,
        )?;
        self.replace_content_xml(content)
    }

    /// Remove one inert XForms model from `office:forms` and publish it.
    pub(crate) fn remove_xforms_model(&mut self, position: litchi_core::Position) -> Result<()> {
        let content = crate::xforms::remove_model_with_xml_limits(
            &self.content_xml()?,
            position.get(),
            self.xml_limits,
        )?;
        self.replace_content_xml(content)
    }

    /// Inspect ordered ODF variable declarations in content and styles.
    pub fn variable_declarations(&self) -> FamilyResult<crate::variable_declaration::Declarations> {
        let content = self.content_xml()?;
        let styles = self.styles_xml()?;
        let mut parts = vec![(content.as_str(), crate::variable_declaration::Part::Content)];
        if let Some(styles) = styles.as_deref() {
            parts.push((styles, crate::variable_declaration::Part::Styles));
        }
        Ok(crate::variable_declaration::parse_parts_with_xml_limits(
            &parts,
            self.xml_limits,
        )?)
    }

    /// Atomically insert or replace one variable declaration container.
    ///
    /// The group must target `content.xml` or `styles.xml`. Auxiliary package
    /// entries remain byte-for-byte intact, and formulas remain inert.
    pub(crate) fn set_variable_declaration_group(
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
                crate::variable_declaration::set_xml_with_xml_limits(
                    &content,
                    group,
                    self.xml_limits,
                )?,
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
                    Some(crate::variable_declaration::set_xml_with_xml_limits(
                        &styles,
                        group,
                        self.xml_limits,
                    )?),
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
    pub(crate) fn remove_variable_declaration_group(
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
                crate::variable_declaration::remove_xml_with_xml_limits(
                    &content,
                    scope,
                    kind,
                    self.xml_limits,
                )?,
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
                    Some(crate::variable_declaration::remove_xml_with_xml_limits(
                        &styles,
                        scope,
                        kind,
                        self.xml_limits,
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
        crate::variable_declaration::parse_parts_with_xml_limits(&parts, self.xml_limits)?;

        let mut writer = crate::core::PackageWriter::new(
            litchi_odf_common::core::StreamingArchiveLimits::default(),
        );
        writer.set_mimetype(&self.mimetype)?;
        if changed_part == crate::variable_declaration::Part::Content {
            crate::variable_declaration::splice_publication(
                &self.package,
                constants::ODF_CONTENT,
                &content,
                scope,
                kind,
                self.xml_limits,
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
                    self.xml_limits,
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
        let replacement =
            Package::from_vec_with_xml_limits(writer.finish_to_bytes()?, self.xml_limits)?;
        *self = replacement;
        Ok(old)
    }

    /// Discover package, inline, missing, and inert linked embedded objects.
    pub fn embedded_objects(&self) -> FamilyResult<Vec<crate::Object>> {
        let content = self.content_xml()?;
        let styles = self.styles_xml()?;
        let package = self.package.package()?;
        Ok(crate::embedded::scan_package(
            &content,
            styles.as_deref(),
            &package,
        )?)
    }

    /// Return bytes only for inline or verified package-contained images.
    /// Linked images remain inert and are never fetched.
    pub fn image_bytes(&self, image: &crate::Image) -> FamilyResult<Option<Vec<u8>>> {
        match &image.source {
            litchi_odf_common::media::Source::Inline { bytes, .. } => Ok(Some(bytes.clone())),
            litchi_odf_common::media::Source::PackagePart { path, .. } => {
                Ok(Some(self.package.get_file(path)?))
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

    /// Atomically replace `path` with the exact package bytes (full
    /// durability). The artifact is staged in a sibling temporary file and
    /// published with one rename; a failure before publication leaves the
    /// destination unchanged.
    ///
    /// # Errors
    ///
    /// Returns [`litchi_odf_common::Error::Core`] carrying the I/O failure.
    pub fn save(
        &self,
        path: impl AsRef<std::path::Path>,
    ) -> std::result::Result<(), litchi_odf_common::Error> {
        self.save_with_durability(path, litchi_core::Durability::Full)
    }

    /// Atomically replace `path` with the exact package bytes under explicit
    /// durability.
    ///
    /// # Errors
    ///
    /// As [`Self::save`].
    pub fn save_with_durability(
        &self,
        path: impl AsRef<std::path::Path>,
        durability: litchi_core::Durability,
    ) -> std::result::Result<(), litchi_odf_common::Error> {
        litchi_odf_common::publish::save_bytes(path.as_ref(), self.as_bytes(), durability)
    }
}
