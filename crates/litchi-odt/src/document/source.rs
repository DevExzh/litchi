//! Immutable, positional-source ODT reads.
//!
//! [`SourceBackedDocument`] is an additive read-only owner.  The existing
//! [`super::Document`] remains the complete owned and mutable facade; callers
//! explicitly call [`SourceBackedDocument::materialize`] before entering that
//! mutation boundary.  The ZIP index and the validated core XML parts are
//! retained, while unrelated package members remain cold until selected.

use std::{
    fmt,
    io::Write,
    sync::{
        Arc, OnceLock,
        atomic::{AtomicUsize, Ordering},
    },
};

#[cfg(any(unix, windows))]
use std::path::Path;

#[cfg(any(unix, windows))]
use litchi_core::FileSource;
use litchi_core::{Error, Metadata, ReadAt, Result, SourceVersion};
use litchi_odf_common::core::{Content, Meta, SourceBackedPackage, SourcePackageLimits, Styles};
use litchi_odf_common::package::{is_media_path, resolve_package_path};
use zeroize::Zeroizing;

use super::codec::parse_hyperlinks;
use crate::elements::style::{StyleElements, StyleRegistry};
use crate::elements::table::Table as ElementTable;
use crate::elements::text::{Paragraph as ElementParagraph, TextElements};

/// The bounded ZIP-index policy used by positional ODF facades.
pub type ReadLimits = SourcePackageLimits;

pub(super) const FAMILY_NAME: &str = "ODT";
const MAX_CONTENT_BYTES: u64 = 256 * 1024 * 1024;
const TEXT_CACHE_QUERY_THRESHOLD: usize = 2;
const TEXT_CACHE_MAX_BYTES: usize = 16 * 1024 * 1024;

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
    text_queries: AtomicUsize,
    text_cache: OnceLock<Option<String>>,
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
        let mut password = Zeroizing::new(password.into());
        Self::from_read_at_inner_with_password(
            source,
            ReadLimits::default(),
            std::mem::take(&mut *password),
        )
    }

    /// Open an ODT from a positional source with explicit finite limits.
    pub fn from_read_at_with_limits(source: Arc<dyn ReadAt>, limits: ReadLimits) -> Result<Self> {
        Self::from_read_at_inner(source, limits)
    }

    /// Adopt an already validated positional ODF package without rebuilding
    /// its ZIP index.
    ///
    /// This hidden handoff is used by the unified format detector after it has
    /// arbitrated any competing OOXML catalog on the same physical source.
    #[doc(hidden)]
    pub fn from_source_package(package: SourceBackedPackage) -> Result<Self> {
        let source = package.source_arc();
        Self::from_package(source, package)
    }

    /// Open a password-protected ODT from a positional source with explicit
    /// finite limits.
    pub fn from_read_at_with_limits_and_password(
        source: Arc<dyn ReadAt>,
        limits: ReadLimits,
        password: impl Into<String>,
    ) -> Result<Self> {
        let mut password = Zeroizing::new(password.into());
        Self::from_read_at_inner_with_password(source, limits, std::mem::take(&mut *password))
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

            if package
                .member_materialized_size(crate::constants::ODF_CONTENT)?
                .is_some_and(|size| size > MAX_CONTENT_BYTES)
            {
                return Err(Error::InvalidFormat(format!(
                    "{FAMILY_NAME} content.xml exceeds the family limit"
                )));
            }
            let content_bytes = package.get_file(crate::constants::ODF_CONTENT)?;
            let content = Content::from_vec(content_bytes)?;
            // One fused tokenization validates the content structure and
            // collects the automatic styles; a validation error returns
            // here, before styles.xml is fetched, exactly as the historical
            // first pass early-returned.
            let content_styles = super::open_parse::OpenParse::run(content.xml_content())?;

            let styles = if package.has_file(crate::constants::ODF_STYLES)? {
                Some(Styles::from_vec(
                    package.get_file(crate::constants::ODF_STYLES)?,
                )?)
            } else {
                None
            };

            // Keep meta.xml's UTF-8 validation eager, like the owned facade,
            // while retaining semantic metadata validation for metadata()
            // and odf_metadata().
            let meta = if package.has_file(crate::constants::ODF_META)? {
                Some(Meta::from_vec(
                    package.get_file(crate::constants::ODF_META)?,
                )?)
            } else {
                None
            };

            let mut style_registry = StyleRegistry::default();
            if let Some(styles_part) = styles.as_ref() {
                style_registry = StyleElements::parse_styles(styles_part.xml_content())?;
            }
            style_registry.try_extend(content_styles.finish()?)?;

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
            text_queries: AtomicUsize::new(0),
            text_cache: OnceLock::new(),
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
        if let Some(cache) = self.text_cache.get() {
            self.check_source()?;
            return self.text_from_cache(cache);
        }

        // Keep the first call on the established parser path.  Once the
        // finite threshold is reached, retain one bounded projection when it
        // is safe to do so.  Saturating the counter avoids an eventual wrap
        // changing the cache state after an impractical number of calls.
        let previous = self
            .text_queries
            .fetch_update(Ordering::Relaxed, Ordering::Relaxed, |count| {
                Some(count.saturating_add(1))
            })
            .unwrap_or(usize::MAX);
        if previous.saturating_add(1) < TEXT_CACHE_QUERY_THRESHOLD {
            return self.text_uncached();
        }

        let text = self.text_uncached()?;
        // Do not retain a projection unless the source is still current after
        // the complete validating parser has finished.  The trailing check
        // below still covers the publication window itself.
        self.check_source()?;
        if self.text_cache.get().is_none() {
            self.publish_text_cache(&text);
        }
        let final_check = self.check_source();
        match final_check {
            Err(error) => Err(error),
            Ok(()) => Ok(text),
        }
    }

    fn text_uncached(&self) -> Result<String> {
        self.check_source()?;
        let result = TextElements::extract_text(self.content.xml_content());
        prefer_current(self.source.as_ref(), self.source_version, result)
    }

    fn text_from_cache(&self, cache: &Option<String>) -> Result<String> {
        // `None` is a terminal no-retention marker for an oversized or
        // refused cache construction.  Keep returning the ordinary parser
        // result without retrying the cache construction.
        let Some(text) = cache else {
            return self.text_uncached();
        };
        let result = clone_text_result(text);
        let final_check = self.check_source();
        match final_check {
            Err(error) => Err(error),
            Ok(()) => result,
        }
    }

    fn publish_text_cache(&self, text: &str) {
        if text.len() > TEXT_CACHE_MAX_BYTES {
            let _ = self.text_cache.set(None);
            return;
        }

        // Reserve the retained copy fallibly.  The original parser result is
        // still returned when the bounded optional cache cannot be allocated.
        let mut retained = String::new();
        if retained.try_reserve_exact(text.len()).is_err() {
            let _ = self.text_cache.set(None);
            return;
        }
        retained.push_str(text);
        let _ = self.text_cache.set(Some(retained));
    }

    /// Write visible paragraphs and headings as bounded UTF-8 text to a
    /// caller-owned sequential, non-seek sink without using the text cache.
    ///
    /// Headings are paragraph-like objects. Formatting is omitted, inline
    /// `text:line-break` controls remain `\n`, and separators come from
    /// `options`. Tracked-change definitions, note bodies, and ruby
    /// pronunciation runs are excluded from visible text. The ODT parser
    /// applies a hard 64 MiB decoded-text ceiling in addition to the shared
    /// output limits in `options`.
    ///
    /// Output may be partial when parsing, a resource limit, or the sink
    /// fails. This method never flushes or rolls back the caller-owned sink,
    /// and it has no cancellation hook. A source revision observed before or
    /// after parsing takes precedence over another failure while preserving
    /// the sink progress accumulated so far.
    pub fn write_text_to<W: Write + ?Sized>(
        &self,
        output: &mut W,
        options: litchi_core::TextOutputOptions<'_>,
    ) -> std::result::Result<litchi_core::TextOutputReport, litchi_core::TextOutputError<Error>>
    {
        let mut writer = litchi_core::SequentialTextWriter::new(output, options);
        let parse_result = match self.check_source() {
            Ok(()) => {
                TextElements::write_text_to_with_writer(self.content.xml_content(), &mut writer)
            },
            Err(error) => Err(writer.document_error(error)),
        };

        match self.check_source() {
            Err(error) => Err(writer.document_error(error)),
            Ok(()) => parse_result.map(|()| writer.finish()),
        }
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

    /// Return all semantic notes in document order without materializing the
    /// remaining package members.
    pub fn notes(&self) -> Result<Vec<crate::Note>> {
        self.check_source()?;
        let result = crate::note::parse_notes(self.content.xml_content());
        prefer_current(self.source.as_ref(), self.source_version, result)
    }

    /// Return semantic footnotes in document order.
    pub fn footnotes(&self) -> Result<Vec<crate::Note>> {
        let result = self.notes().and_then(|notes| {
            let mut filtered = Vec::new();
            filtered
                .try_reserve_exact(notes.len())
                .map_err(|source| Error::Allocation {
                    resource: "ODT source footnote projection",
                    source,
                })?;
            for note in notes {
                if note.class() == crate::NoteClass::Footnote {
                    filtered.push(note);
                }
            }
            Ok(filtered)
        });
        prefer_current(self.source.as_ref(), self.source_version, result)
    }

    /// Return semantic endnotes in document order.
    pub fn endnotes(&self) -> Result<Vec<crate::Note>> {
        let result = self.notes().and_then(|notes| {
            let mut filtered = Vec::new();
            filtered
                .try_reserve_exact(notes.len())
                .map_err(|source| Error::Allocation {
                    resource: "ODT source endnote projection",
                    source,
                })?;
            for note in notes {
                if note.class() == crate::NoteClass::Endnote {
                    filtered.push(note);
                }
            }
            Ok(filtered)
        });
        prefer_current(self.source.as_ref(), self.source_version, result)
    }

    /// Extract hyperlinks from the retained content projection.
    pub fn hyperlinks(&self) -> Result<Vec<(String, String)>> {
        self.check_source()?;
        let result = parse_hyperlinks(self.content.xml_content());
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
        let value = self.style_registry.try_get_resolved_properties(style_name);
        prefer_current(self.source.as_ref(), self.source_version, value)
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
            litchi_odf_common::media::Source::Inline { bytes, .. } => (|| {
                let mut copy = Vec::new();
                copy.try_reserve_exact(bytes.len())
                    .map_err(|source| Error::Allocation {
                        resource: "ODT inline image bytes",
                        source,
                    })?;
                copy.extend_from_slice(bytes);
                Ok(Some(copy))
            })(),
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

fn clone_text_result(text: &str) -> Result<String> {
    let mut result = String::new();
    result
        .try_reserve_exact(text.len())
        .map_err(|source| Error::Allocation {
            resource: "ODT source text result",
            source,
        })?;
    result.push_str(text);
    Ok(result)
}

fn prefer_current<T>(source: &dyn ReadAt, expected: SourceVersion, result: Result<T>) -> Result<T> {
    let observed = source.version()?;
    ensure_current(expected, observed)?;
    result
}
