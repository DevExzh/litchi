//! Immutable, positional-source ODT reads.
//!
//! [`SourceBackedDocument`] is an additive read-only owner.  The existing
//! [`super::Document`] remains the complete owned and mutable facade; callers
//! explicitly call [`SourceBackedDocument::materialize`] before entering that
//! mutation boundary.  The ZIP index and the validated core XML parts are
//! retained, while unrelated package members remain cold until selected.

use std::{fmt, sync::Arc};

#[cfg(any(unix, windows))]
use std::path::Path;

#[cfg(any(unix, windows))]
use litchi_core::FileSource;
use litchi_core::{Error, Metadata, ReadAt, Result, SourceVersion};
use litchi_odf_common::core::{
    Content, Meta, SourceBackedPackage, SourcePackageLimits, Styles, validate_content_document_part,
};
use litchi_odf_common::package::{is_media_path, resolve_package_path};
use zeroize::Zeroizing;

use crate::elements::style::{StyleElements, StyleRegistry};
use crate::elements::table::Table as ElementTable;
use crate::elements::text::{Paragraph as ElementParagraph, TextElements};

/// The bounded ZIP-index policy used by positional ODF facades.
pub type ReadLimits = SourcePackageLimits;

const FAMILY_NAME: &str = "ODT";
const CONTENT_ROOT: &str = "<office:text";

/// Read-only ODT access over an immutable positional source.
///
/// Opening validates the ZIP archive, MIME member, manifest, content root,
/// and UTF-8 core XML parts, then retains only `content.xml`, the optional
/// `styles.xml` and `meta.xml` projections, and the style registry.  Package
/// members such as pictures, embedded objects, and scripts are read only when
/// selected through the explicit package-member methods.  Every public
/// operation checks the captured [`SourceVersion`] and reports
/// [`Error::SourceChanged`] when the source no longer identifies the same
/// snapshot.
pub struct SourceBackedDocument {
    package: SourceBackedPackage,
    source: Arc<dyn ReadAt>,
    source_version: SourceVersion,
    content: Content,
    styles: Option<Styles>,
    meta: Option<Meta>,
    style_registry: StyleRegistry,
}

impl fmt::Debug for SourceBackedDocument {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SourceBackedDocument")
            .field("source_version", &self.source_version)
            .field("has_styles", &self.styles.is_some())
            .field("has_meta", &self.meta.is_some())
            .field("styles", &self.style_registry.styles.len())
            .finish_non_exhaustive()
    }
}

impl SourceBackedDocument {
    /// Open an ODT from a regular filesystem file without slurping it.
    #[cfg(any(unix, windows))]
    pub fn from_path(path: impl AsRef<Path>) -> Result<Self> {
        Self::from_read_at(file_source(path)?)
    }

    /// Open an ODT from a regular filesystem file with explicit ZIP limits.
    #[cfg(any(unix, windows))]
    pub fn from_path_with_limits(path: impl AsRef<Path>, limits: ReadLimits) -> Result<Self> {
        Self::from_read_at_with_limits(file_source(path)?, limits)
    }

    /// Open a password-protected ODT from a regular filesystem file.
    #[cfg(any(unix, windows))]
    pub fn from_path_with_password(
        path: impl AsRef<Path>,
        password: impl Into<String>,
    ) -> Result<Self> {
        // Enter the zeroizing owner before any fallible source-backed work.
        let mut password = Zeroizing::new(password.into());
        let source = file_source(path)?;
        Self::from_read_at_with_password(source, std::mem::take(&mut *password))
    }

    /// Open a password-protected ODT from a regular filesystem file with
    /// explicit ZIP limits.
    #[cfg(any(unix, windows))]
    pub fn from_path_with_limits_and_password(
        path: impl AsRef<Path>,
        limits: ReadLimits,
        password: impl Into<String>,
    ) -> Result<Self> {
        let mut password = Zeroizing::new(password.into());
        let source = file_source(path)?;
        Self::from_read_at_with_limits_and_password(source, limits, std::mem::take(&mut *password))
    }

    /// Alias for [`Self::from_path`].
    #[cfg(any(unix, windows))]
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        Self::from_path(path)
    }

    /// Alias for [`Self::from_path_with_limits`].
    #[cfg(any(unix, windows))]
    pub fn open_with_limits(path: impl AsRef<Path>, limits: ReadLimits) -> Result<Self> {
        Self::from_path_with_limits(path, limits)
    }

    /// Alias for [`Self::from_path_with_password`].
    #[cfg(any(unix, windows))]
    pub fn open_with_password(path: impl AsRef<Path>, password: impl Into<String>) -> Result<Self> {
        Self::from_path_with_password(path, password)
    }

    /// Alias for [`Self::from_path_with_limits_and_password`].
    #[cfg(any(unix, windows))]
    pub fn open_with_limits_and_password(
        path: impl AsRef<Path>,
        limits: ReadLimits,
        password: impl Into<String>,
    ) -> Result<Self> {
        Self::from_path_with_limits_and_password(path, limits, password)
    }

    /// Open an ODT from a caller-provided positional source.
    pub fn from_read_at(source: Arc<dyn ReadAt>) -> Result<Self> {
        Self::from_read_at_with_limits(source, ReadLimits::default())
    }

    /// Open a password-protected ODT from a caller-provided positional source.
    pub fn from_read_at_with_password(
        source: Arc<dyn ReadAt>,
        password: impl Into<String>,
    ) -> Result<Self> {
        Self::from_read_at_inner_with_password(source, ReadLimits::default(), password)
    }

    /// Open an ODT from a positional source with explicit finite limits.
    pub fn from_read_at_with_limits(source: Arc<dyn ReadAt>, limits: ReadLimits) -> Result<Self> {
        Self::from_read_at_inner(source, limits)
    }

    /// Open a password-protected ODT from a positional source with explicit
    /// finite limits.
    pub fn from_read_at_with_limits_and_password(
        source: Arc<dyn ReadAt>,
        limits: ReadLimits,
        password: impl Into<String>,
    ) -> Result<Self> {
        Self::from_read_at_inner_with_password(source, limits, password)
    }

    fn from_read_at_inner(source: Arc<dyn ReadAt>, limits: ReadLimits) -> Result<Self> {
        let package = SourceBackedPackage::from_read_at_with_limits(Arc::clone(&source), limits)?;
        Self::from_package(source, package)
    }

    fn from_read_at_inner_with_password(
        source: Arc<dyn ReadAt>,
        limits: ReadLimits,
        password: impl Into<String>,
    ) -> Result<Self> {
        // SourceBackedPackage takes ownership into Zeroizing before any
        // archive work.  This also keeps encrypted payloads lazy.
        let package = SourceBackedPackage::from_read_at_with_limits_and_password(
            Arc::clone(&source),
            limits,
            password,
        )?;
        Self::from_package(source, package)
    }

    fn from_package(source: Arc<dyn ReadAt>, package: SourceBackedPackage) -> Result<Self> {
        let source_version = package.source_version()?;
        let parsed = (|| {
            let mimetype = package.mimetype()?;
            super::package::validate_mimetype(mimetype)?;

            let content_bytes = package.get_file(crate::constants::ODF_CONTENT)?;
            let content = Content::from_bytes(&content_bytes)?;
            validate_content_document_part(content.xml_content(), CONTENT_ROOT, FAMILY_NAME)?;

            let styles = if package.has_file(crate::constants::ODF_STYLES)? {
                Some(Styles::from_bytes(
                    &package.get_file(crate::constants::ODF_STYLES)?,
                )?)
            } else {
                None
            };

            // Keep meta.xml's UTF-8 validation eager, like the owned facade,
            // while retaining semantic metadata validation for metadata()
            // and odf_metadata().
            let meta = if package.has_file(crate::constants::ODF_META)? {
                Some(Meta::from_bytes(
                    &package.get_file(crate::constants::ODF_META)?,
                )?)
            } else {
                None
            };

            let mut style_registry = StyleRegistry::default();
            if let Some(styles_part) = styles.as_ref()
                && let Ok(registry) = StyleElements::parse_styles(styles_part.xml_content())
            {
                style_registry = registry;
            }
            if let Ok(content_registry) = StyleElements::parse_styles(content.xml_content()) {
                for (_name, style) in content_registry.styles {
                    style_registry.add_style(style);
                }
            }

            Ok((content, styles, meta, style_registry))
        })();
        let (content, styles, meta, style_registry) =
            prefer_current(source.as_ref(), source_version, parsed)?;

        Ok(Self {
            package,
            source,
            source_version,
            content,
            styles,
            meta,
            style_registry,
        })
    }

    /// Check the source identity and revision without reading a member.
    pub fn check_source(&self) -> Result<()> {
        ensure_current(self.source_version, self.source.version()?)
    }

    /// Return the source identity captured during open.
    #[must_use = "the captured source version identifies this source snapshot"]
    pub fn source_version(&self) -> Result<SourceVersion> {
        self.check_source()?;
        let value = self.source_version;
        self.check_source()?;
        Ok(value)
    }

    /// Return the captured physical source length.
    pub fn source_len(&self) -> Result<u64> {
        self.check_source()?;
        let value = self.package.len();
        self.check_source()?;
        Ok(value)
    }

    /// Return the validated package MIME type.
    pub fn mimetype(&self) -> Result<&str> {
        self.check_source()?;
        let value = self.package.mimetype()?;
        self.check_source()?;
        Ok(value)
    }

    /// Borrow the validated `content.xml` snapshot.
    pub fn content_xml(&self) -> Result<&str> {
        self.check_source()?;
        let value = self.content.xml_content();
        self.check_source()?;
        Ok(value)
    }

    /// Borrow the optional validated `styles.xml` snapshot.
    pub fn styles_xml(&self) -> Result<Option<&str>> {
        self.check_source()?;
        let value = self.styles.as_ref().map(Styles::xml_content);
        self.check_source()?;
        Ok(value)
    }

    /// Borrow the optional validated `meta.xml` snapshot.
    pub fn meta_xml(&self) -> Result<Option<&str>> {
        self.check_source()?;
        let value = self.meta.as_ref().map(Meta::xml_content);
        self.check_source()?;
        Ok(value)
    }

    /// Extract all document text while retaining paragraph separators.
    pub fn text(&self) -> Result<String> {
        self.check_source()?;
        let result = TextElements::extract_text(self.content.xml_content());
        prefer_current(self.source.as_ref(), self.source_version, result)
    }

    /// Return the number of paragraphs in document order.
    pub fn paragraph_count(&self) -> Result<usize> {
        self.check_source()?;
        let result =
            TextElements::parse_paragraphs(self.content.xml_content()).map(|items| items.len());
        prefer_current(self.source.as_ref(), self.source_version, result)
    }

    /// Return all paragraphs in document order.
    pub fn paragraphs(&self) -> Result<Vec<ElementParagraph>> {
        self.check_source()?;
        let result = TextElements::parse_paragraphs(self.content.xml_content());
        prefer_current(self.source.as_ref(), self.source_version, result)
    }

    /// Return one paragraph by zero-based paragraph position.
    pub fn paragraph(&self, index: usize) -> Result<Option<ElementParagraph>> {
        self.check_source()?;
        let result = TextElements::parse_paragraph_at(self.content.xml_content(), index);
        prefer_current(self.source.as_ref(), self.source_version, result)
    }

    /// Return all tables in document order.
    pub fn tables(&self) -> Result<Vec<ElementTable>> {
        self.check_source()?;
        let result = crate::elements::table::TableElements::parse_tables_from_content(
            self.content.xml_content(),
        );
        prefer_current(self.source.as_ref(), self.source_version, result)
    }

    /// Return paragraphs, headings, lists, and tables in source document order.
    pub fn elements(&self) -> Result<Vec<crate::elements::parser::OrderElement>> {
        self.check_source()?;
        let result =
            crate::elements::parser::Parser::parse_elements_in_order(self.content.xml_content());
        prefer_current(self.source.as_ref(), self.source_version, result)
    }

    /// Return all tables with repeated rows and cells expanded.
    pub fn tables_expanded(&self) -> Result<Vec<ElementTable>> {
        self.check_source()?;
        let result = (|| {
            let tables = crate::elements::table::TableElements::parse_tables_from_content(
                self.content.xml_content(),
            )?;
            crate::elements::table_expansion::TableExpander::expand_tables(tables)
        })();
        prefer_current(self.source.as_ref(), self.source_version, result)
    }

    /// Return common document metadata.  Invalid metadata values remain a
    /// lazy error, matching [`super::Document::metadata`].
    pub fn metadata(&self) -> Result<Metadata> {
        self.check_source()?;
        let result = self
            .meta
            .as_ref()
            .map_or_else(|| Ok(Metadata::default()), Meta::try_extract_metadata);
        prefer_current(self.source.as_ref(), self.source_version, result)
    }

    /// Return the complete format-specific ODF metadata model.
    pub fn odf_metadata(&self) -> Result<Option<crate::Metadata>> {
        self.check_source()?;
        let result = self.meta.as_ref().map(Meta::odf_metadata).transpose();
        prefer_current(self.source.as_ref(), self.source_version, result)
    }

    /// Return the retained style registry.
    pub fn styles(&self) -> Result<&StyleRegistry> {
        self.check_source()?;
        let value = &self.style_registry;
        self.check_source()?;
        Ok(value)
    }

    /// Return resolved properties for one style name.
    pub fn get_style_properties(
        &self,
        style_name: &str,
    ) -> Result<crate::elements::style::StyleProperties<'static>> {
        self.check_source()?;
        let value = self.style_registry.get_resolved_properties(style_name);
        self.check_source()?;
        Ok(value)
    }

    /// Discover referenced, inline, missing, and inert linked images.
    pub fn images(&self) -> Result<Vec<crate::Image>> {
        self.check_source()?;
        let result = crate::media::scan_package(
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            &self.package,
        );
        prefer_current(self.source.as_ref(), self.source_version, result)
    }

    /// Read bytes for an inline or verified package-contained image.
    /// Linked and missing images remain inert and return `None`.
    pub fn image_bytes(&self, image: &crate::Image) -> Result<Option<Vec<u8>>> {
        self.check_source()?;
        let result = match &image.source {
            litchi_odf_common::media::Source::Inline { bytes, .. } => Ok(Some(bytes.clone())),
            litchi_odf_common::media::Source::PackagePart { path, .. } => self.media_data(path),
            litchi_odf_common::media::Source::MissingPackagePart { .. }
            | litchi_odf_common::media::Source::Linked { .. }
            | litchi_odf_common::media::Source::Missing
            | _ => Ok(None),
        };
        prefer_current(self.source.as_ref(), self.source_version, result)
    }

    /// List package members without reading their payloads.
    pub fn files(&self) -> Result<Vec<String>> {
        self.check_source()?;
        let result = self.package.files();
        prefer_current(self.source.as_ref(), self.source_version, result)
    }

    /// List package media members without reading their payloads.
    pub fn media_files(&self) -> Result<Vec<String>> {
        self.check_source()?;
        let result = self.package.media_files();
        prefer_current(self.source.as_ref(), self.source_version, result)
    }

    /// Read one safe package-contained member on demand.
    pub fn member_data(&self, path: &str) -> Result<Option<Vec<u8>>> {
        self.check_source()?;
        let result = (|| {
            let path = resolve_package_path(path)?;
            if !self.package.has_file(&path)? {
                return Ok(None);
            }
            self.package.get_file(&path).map(Some)
        })();
        prefer_current(self.source.as_ref(), self.source_version, result)
    }

    /// Read one package media member on demand.
    pub fn media_data(&self, path: &str) -> Result<Option<Vec<u8>>> {
        self.check_source()?;
        let result = (|| {
            let path = resolve_package_path(path)?;
            if !is_media_path(&path) {
                return Ok(None);
            }
            self.member_data(&path)
        })();
        prefer_current(self.source.as_ref(), self.source_version, result)
    }

    /// Return inert document interaction and protection metadata.
    pub fn protection(&self) -> Result<crate::protection::Policy> {
        self.check_source()?;
        let result = (|| {
            if !self.package.has_file(crate::constants::ODF_SETTINGS)? {
                return Ok(crate::protection::Policy::default());
            }
            crate::protection::parse_package(
                &self.package.get_file(crate::constants::ODF_SETTINGS)?,
            )
        })();
        prefer_current(self.source.as_ref(), self.source_version, result)
    }

    /// Explicitly materialize this source into the existing mutable ODT owner.
    pub fn materialize(self) -> Result<super::Document> {
        let source = Arc::clone(&self.source);
        let source_version = self.source_version;
        ensure_current(source_version, source.version()?)?;
        let result = self
            .package
            .materialize()
            .and_then(super::Document::from_owned_package);
        prefer_current(source.as_ref(), source_version, result)
    }
}

#[cfg(any(unix, windows))]
fn file_source(path: impl AsRef<Path>) -> Result<Arc<dyn ReadAt>> {
    Ok(Arc::new(FileSource::open(path)?))
}

fn ensure_current(expected: SourceVersion, observed: SourceVersion) -> Result<()> {
    if expected == observed {
        Ok(())
    } else {
        Err(Error::SourceChanged { expected, observed })
    }
}

fn prefer_current<T>(source: &dyn ReadAt, expected: SourceVersion, result: Result<T>) -> Result<T> {
    let observed = source.version()?;
    ensure_current(expected, observed)?;
    result
}
