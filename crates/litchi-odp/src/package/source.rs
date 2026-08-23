//! Immutable, positional-source ODP reads.
//!
//! [`SourceBackedPresentation`] is deliberately additive: the existing
//! [`super::Presentation`] keeps its owned-byte and mutation behavior.  This
//! facade retains only the validated XML needed for ordinary semantic reads;
//! package media remains cold until [`SourceBackedPresentation::media_data`]
//! selects a member.

use std::sync::{
    Arc, OnceLock,
    atomic::{AtomicUsize, Ordering},
};

#[cfg(any(unix, windows))]
use std::path::Path;

#[cfg(any(unix, windows))]
use litchi_core::FileSource;
use litchi_core::{Error, ReadAt, Result, SourceVersion};
use litchi_odf_common::core::{
    Meta, SourceBackedPackage, SourcePackageLimits, validate_content_part,
};
use zeroize::Zeroizing;

use crate::codec::Parser;
use crate::model::{Reference, Slide};

const ODF_PRESENTATION: &str = "application/vnd.oasis.opendocument.presentation";
const BODY_MARKER: &str = "<office:presentation";
const FAMILY_NAME: &str = "ODP";
const SLIDES_CACHE_QUERY_THRESHOLD: usize = 2;
// The source-backed owner already retains bounded content.xml bytes.  Keep a
// much smaller ceiling for the optional semantic projection so a repeated
// query cannot turn an otherwise lazy source view into a large second model.
const SLIDES_CACHE_MAX_SOURCE_BYTES: usize = 4 * 1024 * 1024;
const SLIDES_CACHE_MAX_COUNT: usize = 256;
const TEXT_CACHE_QUERY_THRESHOLD: usize = 2;
const TEXT_CACHE_MAX_BYTES: usize = 16 * 1024 * 1024;

/// The bounded ZIP-index limits accepted by source-backed ODF facades.
pub type ReadLimits = SourcePackageLimits;

#[cfg(any(unix, windows))]
fn file_source(path: impl AsRef<Path>) -> Result<Arc<dyn ReadAt>> {
    Ok(Arc::new(FileSource::open(path)?))
}

/// Read-only ODP access over an immutable positional source.
///
/// Opening validates the ZIP archive, MIME member, and family content root,
/// then loads only `content.xml` and the optional `styles.xml`.  Other package
/// members—including pictures and audio/video payloads—remain deferred until
/// [`Self::media_data`] is called.  Every public operation checks the captured
/// [`SourceVersion`] before and after work, returning
/// [`Error::SourceChanged`] when the source revision is no longer current.
///
/// This type has no edit or output methods.  Use [`super::Presentation`] for
/// the established owned-byte/mutation API.
pub struct SourceBackedPresentation {
    package: SourceBackedPackage,
    source: Arc<dyn ReadAt>,
    source_version: SourceVersion,
    content_xml: String,
    styles_xml: Option<String>,
    metadata: Option<litchi_core::Metadata>,
    slides_queries: AtomicUsize,
    slides_cache: OnceLock<Option<Arc<[Slide]>>>,
    text_queries: AtomicUsize,
    text_cache: OnceLock<Option<String>>,
}

impl SourceBackedPresentation {
    /// Open an ODP from a regular filesystem file without slurping it.
    #[cfg(any(unix, windows))]
    pub fn from_path(path: impl AsRef<Path>) -> Result<Self> {
        Self::from_read_at(file_source(path)?)
    }

    /// Open an ODP from a regular filesystem file with explicit ZIP limits.
    #[cfg(any(unix, windows))]
    pub fn from_path_with_limits(path: impl AsRef<Path>, limits: ReadLimits) -> Result<Self> {
        Self::from_read_at_with_limits(file_source(path)?, limits)
    }

    /// Open a password-protected ODP from a regular filesystem file.
    #[cfg(any(unix, windows))]
    pub fn from_path_with_password(
        path: impl AsRef<Path>,
        password: impl Into<String>,
    ) -> Result<Self> {
        let mut password = Zeroizing::new(password.into());
        let source = file_source(path)?;
        Self::from_read_at_with_password(source, std::mem::take(&mut *password))
    }

    /// Open a password-protected ODP from a regular filesystem file with
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

    /// Open an ODP from a caller-provided positional source.
    ///
    /// # Errors
    ///
    /// Returns an error when source I/O, ZIP validation, MIME validation, or
    /// family content validation fails.
    pub fn from_read_at(source: Arc<dyn ReadAt>) -> Result<Self> {
        Self::from_read_at_with_limits(source, ReadLimits::default())
    }

    /// Open a password-protected ODP from a caller-provided positional source.
    pub fn from_read_at_with_password(
        source: Arc<dyn ReadAt>,
        password: impl Into<String>,
    ) -> Result<Self> {
        Self::from_read_at_inner_with_password(source, ReadLimits::default(), password)
    }

    /// Open an ODP from a positional source with explicit bounded ZIP limits.
    ///
    /// The archive index is retained by the common ODF owner.  Payloads not
    /// needed for mandatory family validation are not read by this constructor.
    pub fn from_read_at_with_limits(source: Arc<dyn ReadAt>, limits: ReadLimits) -> Result<Self> {
        Self::from_read_at_inner(source, limits)
    }

    /// Open a password-protected ODP from a positional source with explicit
    /// bounded ZIP limits.
    pub fn from_read_at_with_limits_and_password(
        source: Arc<dyn ReadAt>,
        limits: ReadLimits,
        password: impl Into<String>,
    ) -> Result<Self> {
        Self::from_read_at_inner_with_password(source, limits, password)
    }

    fn from_read_at_inner(source: Arc<dyn ReadAt>, limits: ReadLimits) -> Result<Self> {
        let package = SourceBackedPackage::from_read_at_with_limits(Arc::clone(&source), limits)?;
        Self::from_read_at_package(source, package)
    }

    fn from_read_at_inner_with_password(
        source: Arc<dyn ReadAt>,
        limits: ReadLimits,
        password: impl Into<String>,
    ) -> Result<Self> {
        let package = SourceBackedPackage::from_read_at_with_limits_and_password(
            Arc::clone(&source),
            limits,
            password,
        )?;
        Self::from_read_at_package(source, package)
    }

    fn from_read_at_package(source: Arc<dyn ReadAt>, package: SourceBackedPackage) -> Result<Self> {
        // Reuse the common owner's captured revision.  This both preserves
        // its archive/source consistency window and avoids a plaintext
        // password surviving any facade-side probe before zeroization.
        let source_version = package.source_version()?;
        let mimetype = package.mimetype()?;
        if mimetype != ODF_PRESENTATION {
            return Err(Error::InvalidFormat(format!(
                "expected {FAMILY_NAME} package MIME type '{ODF_PRESENTATION}', found '{mimetype}'"
            )));
        }

        let content_xml = String::from_utf8(package.get_file("content.xml")?).map_err(|error| {
            Error::InvalidFormat(format!("{FAMILY_NAME} content.xml is not UTF-8: {error}"))
        })?;
        validate_content_part(&content_xml, BODY_MARKER, FAMILY_NAME)?;

        let styles_xml = if package.has_file("styles.xml")? {
            Some(
                String::from_utf8(package.get_file("styles.xml")?).map_err(|error| {
                    Error::InvalidFormat(format!("{FAMILY_NAME} styles.xml is not UTF-8: {error}"))
                })?,
            )
        } else {
            None
        };

        // Match the ordinary family package's mandatory metadata validation:
        // a malformed optional meta.xml must not become observable only after
        // a source-backed facade is opened.
        let metadata = if package.has_file("meta.xml")? {
            Some(Meta::from_bytes(&package.get_file("meta.xml")?)?.try_extract_metadata()?)
        } else {
            None
        };

        let observed = source.version()?;
        ensure_current(source_version, observed)?;

        Ok(Self {
            package,
            source,
            source_version,
            content_xml,
            styles_xml,
            metadata,
            slides_queries: AtomicUsize::new(0),
            slides_cache: OnceLock::new(),
            text_queries: AtomicUsize::new(0),
            text_cache: OnceLock::new(),
        })
    }

    /// Check the source identity and revision without reading a package part.
    pub fn check_source(&self) -> Result<()> {
        ensure_current(self.source_version, self.source.version()?)
    }

    /// Return the source identity captured during open.
    #[must_use = "the captured source version is needed to identify this source"]
    pub fn source_version(&self) -> Result<SourceVersion> {
        self.check_source()?;
        Ok(self.source_version)
    }

    /// Borrow the validated `content.xml` snapshot.
    pub fn content_xml(&self) -> Result<&str> {
        self.check_source()?;
        Ok(&self.content_xml)
    }

    /// Borrow the optional validated `styles.xml` snapshot.
    pub fn styles_xml(&self) -> Result<Option<&str>> {
        self.check_source()?;
        Ok(self.styles_xml.as_deref())
    }

    /// Return common document metadata projected from the optional
    /// `meta.xml` part, matching [`super::Presentation::metadata`].
    pub fn metadata(&self) -> Result<litchi_core::Metadata> {
        self.check_source()?;
        Ok(self.metadata.clone().unwrap_or_default())
    }

    /// Return the number of slides after running the ordinary parser's full
    /// validation path.
    pub fn slide_count(&self) -> Result<usize> {
        let slides = self.slides()?;
        Ok(slides.len())
    }

    /// Parse all slides with the same semantic parser used by [`super::Presentation`].
    pub fn slides(&self) -> Result<Vec<Slide>> {
        if let Some(cache) = self.slides_cache.get() {
            self.check_source()?;
            return self.slides_from_cache(cache);
        }

        let previous = self
            .slides_queries
            .fetch_update(Ordering::Relaxed, Ordering::Relaxed, |count| {
                Some(count.saturating_add(1))
            })
            .unwrap_or(usize::MAX);
        if previous.saturating_add(1) < SLIDES_CACHE_QUERY_THRESHOLD {
            return self.slides_uncached();
        }

        let slides = self.slides_uncached()?;
        self.check_source()?;
        if self.slides_cache.get().is_none() {
            self.publish_slides_cache(&slides);
        }
        self.check_source()?;
        Ok(slides)
    }

    fn slides_uncached(&self) -> Result<Vec<Slide>> {
        self.check_source()?;
        let result =
            Parser::parse_slides_with_styles(&self.content_xml, self.styles_xml.as_deref());
        let observed = self.source.version()?;
        ensure_current(self.source_version, observed)?;
        result
    }

    /// Parse one slide while retaining the parser's complete-document
    /// validation semantics.
    pub fn slide(&self, index: usize) -> Result<Option<Slide>> {
        if let Some(cache) = self.slides_cache.get() {
            self.check_source()?;
            return self.slide_from_cache(cache, index);
        }

        let previous = self
            .slides_queries
            .fetch_update(Ordering::Relaxed, Ordering::Relaxed, |count| {
                Some(count.saturating_add(1))
            })
            .unwrap_or(usize::MAX);
        if previous.saturating_add(1) < SLIDES_CACHE_QUERY_THRESHOLD {
            return self.slide_uncached(index);
        }

        // Build the shared projection only after the finite threshold.  This
        // keeps a one-off indexed query on the specialized parser while
        // allowing repeated list/index queries to share one validated result.
        let slides = self.slides_uncached()?;
        let selected = slides.get(index).cloned();
        self.check_source()?;
        if self.slides_cache.get().is_none() {
            self.publish_slides_cache(&slides);
        }
        self.check_source()?;
        Ok(selected)
    }

    fn slide_uncached(&self, index: usize) -> Result<Option<Slide>> {
        self.check_source()?;
        let result = Parser::parse_slide_with_styles_at(
            &self.content_xml,
            self.styles_xml.as_deref(),
            index,
        );
        let observed = self.source.version()?;
        ensure_current(self.source_version, observed)?;
        result
    }

    fn slides_from_cache(&self, cache: &Option<Arc<[Slide]>>) -> Result<Vec<Slide>> {
        let Some(slides) = cache else {
            return self.slides_uncached();
        };
        let result = slides.to_vec();
        self.check_source()?;
        Ok(result)
    }

    fn slide_from_cache(
        &self,
        cache: &Option<Arc<[Slide]>>,
        index: usize,
    ) -> Result<Option<Slide>> {
        let Some(slides) = cache else {
            return self.slide_uncached(index);
        };
        let result = slides.get(index).cloned();
        self.check_source()?;
        Ok(result)
    }

    fn publish_slides_cache(&self, slides: &[Slide]) {
        if self.content_xml.len() > SLIDES_CACHE_MAX_SOURCE_BYTES
            || slides.len() > SLIDES_CACHE_MAX_COUNT
        {
            let _ = self.slides_cache.set(None);
            return;
        }

        // The projection is optional.  Retaining it must not turn a valid
        // source query into an allocation failure, and the source-size/count
        // gates keep its aggregate footprint finite before cloning strings.
        let mut retained = Vec::new();
        if retained.try_reserve_exact(slides.len()).is_err() {
            let _ = self.slides_cache.set(None);
            return;
        }
        retained.extend(slides.iter().cloned());
        let _ = self.slides_cache.set(Some(retained.into()));
    }

    /// Extract all visible presentation text with the ordinary facade's
    /// ordering and separator semantics.
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
        // Keep the text selector's established threshold and freshness shape
        // independent of the broader slide projection cache.
        let slides = self.slides_uncached()?;
        let mut all_text = Vec::new();
        for slide in slides {
            let text = slide.all_text();
            if !text.is_empty() {
                all_text.push(text);
            }
        }
        self.check_source()?;
        Ok(all_text.join("\n\n"))
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

    /// Read one package-contained media payload without following external
    /// URLs, fragments, or unsafe relative paths.
    ///
    /// The selected member is decompressed only when this method is called;
    /// opening and semantic slide queries do not materialize media payloads.
    pub fn media_data(&self, media: &Reference) -> Result<Option<Vec<u8>>> {
        self.check_source()?;
        let Some(path) = media.package_path() else {
            self.check_source()?;
            return Ok(None);
        };
        if !self.package.has_file(path)? {
            self.check_source()?;
            return Ok(None);
        }
        let bytes = self.package.get_file(path)?;
        let observed = self.source.version()?;
        ensure_current(self.source_version, observed)?;
        Ok(Some(bytes))
    }

    /// Materialize the exact source into the established mutable presentation
    /// owner.  This is an explicit transition; read-only source operations do
    /// not allocate a complete package buffer.
    pub fn materialize(self) -> Result<super::Presentation> {
        self.package
            .materialize()
            .and_then(super::Presentation::from_owned_package)
    }
}

fn clone_text_result(text: &str) -> Result<String> {
    let mut result = String::new();
    result
        .try_reserve_exact(text.len())
        .map_err(|source| Error::Allocation {
            resource: "ODP source text result",
            source,
        })?;
    result.push_str(text);
    Ok(result)
}

fn ensure_current(expected: SourceVersion, observed: SourceVersion) -> Result<()> {
    if expected == observed {
        Ok(())
    } else {
        Err(Error::SourceChanged { expected, observed })
    }
}
