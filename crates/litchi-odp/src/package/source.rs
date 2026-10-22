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
use litchi_core::{
    Error, ReadAt, Result, SequentialTextWriter, SourceVersion, TextOutputError, TextOutputOptions,
    TextOutputReport,
};
use litchi_odf_common::core::source_publication::{
    SourceContentInsertionError, SourceContentInsertionPlan,
};
use litchi_odf_common::core::{
    Meta, SourceBackedPackage, SourceContentPublicationOptions, SourceContentPublicationReport,
    private::ContentDocumentValidator, validate_content_part,
};
use litchi_odf_common::{Error as OdfError, OpenOptions};
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
    pub(super) package: Arc<SourceBackedPackage>,
    source: Arc<dyn ReadAt>,
    source_version: SourceVersion,
    content_xml: Arc<str>,
    styles_xml: Option<Arc<str>>,
    metadata: Option<litchi_core::Metadata>,
    slides_queries: AtomicUsize,
    slides_cache: OnceLock<Option<Arc<[Slide]>>>,
    text_queries: AtomicUsize,
    text_cache: OnceLock<Option<String>>,
    /// Package repairs and the repairs applied while `content.xml` was
    /// loaded at open.
    open_tolerance: litchi_odf_common::leniency::ToleranceReport,
    /// [`Self::open_tolerance`] followed by the slide scan's content
    /// repairs, computed by the first [`Self::tolerance_report`] or
    /// complete slide scan.
    tolerance: OnceLock<litchi_odf_common::leniency::ToleranceReport>,
    /// The committed output of a source-backed edit, written when this
    /// presentation is written; reads materialize it once (record 0358a).
    deferred: Option<Arc<Deferred>>,
}

/// The pending output of a committed source-backed presentation.
pub(super) struct Deferred {
    plan: SourceContentInsertionPlan,
    options: SourceContentPublicationOptions,
    materialized: OnceLock<SourceBackedPresentation>,
}

impl SourceBackedPresentation {
    /// Open an ODP presentation from a regular filesystem file without slurping it.
    ///
    /// # Errors
    ///
    /// Returns [`litchi_odf_common::Error::Limit`] for an exceeded dimension
    /// and [`litchi_odf_common::Error::Core`] for I/O or validation failures.
    #[cfg(any(unix, windows))]
    pub fn open(path: impl AsRef<Path>) -> std::result::Result<Self, OdfError> {
        Self::open_with_options(path, OpenOptions::new())
    }

    /// Open an ODP presentation from a filesystem file under caller-selected limits.
    ///
    /// # Errors
    ///
    /// As [`Self::open`].
    #[cfg(any(unix, windows))]
    pub fn open_with_options(
        path: impl AsRef<Path>,
        options: OpenOptions,
    ) -> std::result::Result<Self, OdfError> {
        Self::from_read_at_with_options(file_source(path)?, options)
    }

    /// Open password-protected an ODP presentation from a filesystem file.
    ///
    /// # Errors
    ///
    /// As [`Self::open`].
    #[cfg(any(unix, windows))]
    pub fn open_with_password(
        path: impl AsRef<Path>,
        password: impl Into<String>,
    ) -> std::result::Result<Self, OdfError> {
        Self::open_with_password_and_options(path, password, OpenOptions::new())
    }

    /// Open password-protected an ODP presentation from a filesystem file under
    /// caller-selected limits.
    ///
    /// # Errors
    ///
    /// As [`Self::open`].
    #[cfg(any(unix, windows))]
    pub fn open_with_password_and_options(
        path: impl AsRef<Path>,
        password: impl Into<String>,
        options: OpenOptions,
    ) -> std::result::Result<Self, OdfError> {
        let mut password = Zeroizing::new(password.into());
        let source = file_source(path)?;
        Self::from_read_at_with_password_and_options(
            source,
            std::mem::take(&mut *password),
            options,
        )
    }

    /// Open an ODP presentation from a caller-provided positional source.
    ///
    /// # Errors
    ///
    /// As [`Self::open`].
    pub fn from_read_at(source: Arc<dyn ReadAt>) -> std::result::Result<Self, OdfError> {
        Self::from_read_at_with_options(source, OpenOptions::new())
    }

    /// Open an ODP presentation from a positional source under caller-selected limits.
    ///
    /// # Errors
    ///
    /// As [`Self::open`].
    pub fn from_read_at_with_options(
        source: Arc<dyn ReadAt>,
        options: OpenOptions,
    ) -> std::result::Result<Self, OdfError> {
        Self::from_read_at_inner(source, options, None)
    }

    /// Open password-protected an ODP presentation from a caller-provided positional
    /// source.
    ///
    /// # Errors
    ///
    /// As [`Self::open`].
    pub fn from_read_at_with_password(
        source: Arc<dyn ReadAt>,
        password: impl Into<String>,
    ) -> std::result::Result<Self, OdfError> {
        Self::from_read_at_with_password_and_options(source, password, OpenOptions::new())
    }

    /// Open password-protected an ODP presentation from a positional source under
    /// caller-selected limits.
    ///
    /// # Errors
    ///
    /// As [`Self::open`].
    pub fn from_read_at_with_password_and_options(
        source: Arc<dyn ReadAt>,
        password: impl Into<String>,
        options: OpenOptions,
    ) -> std::result::Result<Self, OdfError> {
        let mut password = Zeroizing::new(password.into());
        Self::from_read_at_inner(source, options, Some(std::mem::take(&mut *password)))
    }

    fn from_read_at_inner(
        source: Arc<dyn ReadAt>,
        options: OpenOptions,
        password: Option<String>,
    ) -> std::result::Result<Self, OdfError> {
        let limits = options.source_limits_value();
        let package = match password {
            Some(password) => {
                let mut password = Zeroizing::new(password);
                SourceBackedPackage::from_read_at_with_password_and_options(
                    Arc::clone(&source),
                    std::mem::take(&mut *password),
                    OpenOptions::new().source_limits(limits),
                )?
            },
            None => SourceBackedPackage::from_read_at_with_options(
                Arc::clone(&source),
                OpenOptions::new().source_limits(limits),
            )?,
        };
        check_semantic_part_sizes(&package, options.semantic_limits_value())?;
        Ok(Self::from_read_at_package(source, package)?)
    }

    fn from_read_at_package(source: Arc<dyn ReadAt>, package: SourceBackedPackage) -> Result<Self> {
        // Reuse the common owner's captured revision.  This both preserves
        // its archive/source consistency window and avoids a plaintext
        // password surviving any facade-side probe before zeroization.
        let source_version = package.source_version()?;
        let parsed = (|| {
            let mimetype = package.mimetype()?;
            if mimetype != ODF_PRESENTATION {
                return Err(Error::InvalidFormat(format!(
                    "expected {FAMILY_NAME} package MIME type '{ODF_PRESENTATION}', found '{mimetype}'"
                )));
            }

            // Check the ZIP or manifest-declared materialized size before
            // reading content.xml. For encrypted entries this is plaintext
            // metadata, so a hostile ciphertext cannot force allocation first.
            ContentDocumentValidator::check_materialized_size(
                package.member_materialized_size("content.xml")?,
                FAMILY_NAME,
            )?;
            let mut content_xml =
                String::from_utf8(package.get_file("content.xml")?).map_err(|error| {
                    Error::InvalidFormat(format!("{FAMILY_NAME} content.xml is not UTF-8: {error}"))
                })?;
            let doctype =
                litchi_odf_common::core::private::neutralize_prologue_doctype(&mut content_xml);
            validate_content_part(&content_xml, BODY_MARKER, FAMILY_NAME)?;

            let styles_xml = if package.has_file("styles.xml")? {
                Some(
                    String::from_utf8(package.get_file("styles.xml")?).map_err(|error| {
                        Error::InvalidFormat(format!(
                            "{FAMILY_NAME} styles.xml is not UTF-8: {error}"
                        ))
                    })?,
                )
            } else {
                None
            };

            // Match the ordinary family package's mandatory metadata
            // validation: a malformed optional meta.xml must not become
            // observable only after a source-backed facade is opened.
            let metadata = if package.has_file("meta.xml")? {
                // ODP owners retain no raw XML policy yet; meta.xml is
                // admitted under the documented owner default.
                Some(
                    Meta::from_bytes(
                        &package.get_file("meta.xml")?,
                        litchi_odf_common::XmlStreamLimits::owner_default_value(),
                    )?
                    .try_extract_metadata()?,
                )
            } else {
                None
            };

            Ok((content_xml, styles_xml, metadata, doctype))
        })();

        // Reconcile the source before exposing any secondary parse error. This
        // keeps stale-source errors authoritative across the whole open pass,
        // including pure validation and UTF-8/metadata parsing work.
        let (content_xml, styles_xml, metadata, doctype) =
            reconcile_source(source.as_ref(), source_version, parsed)?;
        let mut open_tolerance = package.tolerance_report().clone();
        if doctype {
            open_tolerance.record(
                litchi_odf_common::leniency::ContentDefect::DoctypeDeclaration,
                0,
            );
        }

        Ok(Self {
            package: Arc::new(package),
            source,
            source_version,
            content_xml: Arc::from(content_xml),
            styles_xml: styles_xml.map(Arc::from),
            metadata,
            slides_queries: AtomicUsize::new(0),
            slides_cache: OnceLock::new(),
            open_tolerance,
            tolerance: OnceLock::new(),
            text_queries: AtomicUsize::new(0),
            text_cache: OnceLock::new(),
            deferred: None,
        })
    }

    /// A share of this presentation's source, retained parts and pending
    /// output, with fresh caches.
    pub(crate) fn share(&self) -> Self {
        Self {
            package: Arc::clone(&self.package),
            source: Arc::clone(&self.source),
            source_version: self.source_version,
            content_xml: Arc::clone(&self.content_xml),
            styles_xml: self.styles_xml.clone(),
            metadata: self.metadata.clone(),
            slides_queries: AtomicUsize::new(0),
            slides_cache: OnceLock::new(),
            text_queries: AtomicUsize::new(0),
            text_cache: OnceLock::new(),
            open_tolerance: self.open_tolerance.clone(),
            tolerance: OnceLock::new(),
            deferred: self.deferred.clone(),
        }
    }

    /// A committed presentation over this one whose output `plan` writes
    /// under `options`. `self` must not itself be committed.
    pub(super) fn deferred(
        &self,
        plan: SourceContentInsertionPlan,
        options: SourceContentPublicationOptions,
    ) -> Self {
        let mut committed = self.share();
        committed.deferred = Some(Arc::new(Deferred {
            plan,
            options,
            materialized: OnceLock::new(),
        }));
        committed
    }

    /// The presentation that serves reads and edits: the materialized output
    /// of a committed presentation, published into memory once, or `self`.
    pub(crate) fn effective(&self) -> Result<&Self> {
        let Some(deferred) = &self.deferred else {
            return Ok(self);
        };
        if let Some(materialized) = deferred.materialized.get() {
            return Ok(materialized);
        }
        let mut bytes = Vec::new();
        deferred
            .plan
            .write_to(&mut bytes, &deferred.options)
            .map_err(|error| Error::from(OdfError::from(error)))?;
        let materialized = Self::from_read_at(Arc::new(litchi_core::OwnedSource::new(bytes)))
            .map_err(Error::from)?;
        Ok(deferred.materialized.get_or_init(|| materialized))
    }

    /// Whether two presentations are shares of one source and pending output.
    pub(super) fn same_share(&self, other: &Self) -> bool {
        match (&self.deferred, &other.deferred) {
            (Some(left), Some(right)) => Arc::ptr_eq(left, right),
            (None, None) => Arc::ptr_eq(&self.package, &other.package),
            _ => false,
        }
    }

    /// Write this presentation as a complete ODP package: the source archive,
    /// or the committed output of a source-backed edit, which copies every
    /// untouched member raw from the retained source (record 0358a).
    ///
    /// # Errors
    ///
    /// Returns the typed publication failure with the sink progress it
    /// observed; a source that changed since open is refused before any byte
    /// is written.
    pub fn write_to<W: std::io::Write>(
        &self,
        writer: W,
    ) -> std::result::Result<SourceContentPublicationReport, SourceContentInsertionError> {
        let options = match &self.deferred {
            Some(deferred) => deferred.options.clone(),
            None => SourceContentPublicationOptions::default(),
        };
        self.write_to_with_options(writer, options)
    }

    /// [`Self::write_to`] under explicit output ceilings, cancellation,
    /// payload verification and execution policy.
    ///
    /// # Errors
    ///
    /// As [`Self::write_to`].
    pub fn write_to_with_options<W: std::io::Write>(
        &self,
        writer: W,
        options: SourceContentPublicationOptions,
    ) -> std::result::Result<SourceContentPublicationReport, SourceContentInsertionError> {
        match &self.deferred {
            Some(deferred) => match deferred.materialized.get() {
                Some(materialized) => materialized.write_to_with_options(writer, options),
                None => deferred.plan.write_to(writer, options),
            },
            None => Ok(self.package.write_content_xml_to_stream_with_options(
                writer,
                self.content_xml.as_bytes(),
                options,
            )?),
        }
    }

    /// Atomically replace `path` with this presentation (full durability).
    ///
    /// # Errors
    ///
    /// Returns the publication or I/O failure; the destination is unchanged
    /// unless the error is a committed-I/O payload.
    #[cfg(any(unix, windows))]
    pub fn save(&self, path: impl AsRef<Path>) -> std::result::Result<(), OdfError> {
        self.save_with_durability(path, litchi_core::Durability::Full)
    }

    /// Atomically replace `path` with this presentation under explicit
    /// durability.
    ///
    /// # Errors
    ///
    /// As [`Self::save`].
    #[cfg(any(unix, windows))]
    pub fn save_with_durability(
        &self,
        path: impl AsRef<Path>,
        durability: litchi_core::Durability,
    ) -> std::result::Result<(), OdfError> {
        litchi_core::atomic::save(
            path.as_ref(),
            durability,
            litchi_core::atomic::Overwrite::Replace,
            |writer| self.write_to(writer).map(drop).map_err(OdfError::from),
            || self.check_source().map_err(OdfError::from),
        )
    }

    /// Check the source identity and revision without reading a package part.
    pub fn check_source(&self) -> Result<()> {
        ensure_current(self.source_version, self.source.version()?)
    }

    /// Everything a tolerant read repaired while this presentation was opened.
    ///
    /// An empty report means the document conformed; see
    /// [`litchi_odf_common::leniency::ToleranceReport`]. Package repairs come
    /// first, then the content repairs of the slide scan (LENIENCY_AUDIT
    /// F9.10). Slides are parsed lazily, so the first call to this method
    /// (or a complete slide scan) runs that scan once.
    #[inline]
    #[must_use]
    pub fn tolerance_report(&self) -> &litchi_odf_common::leniency::ToleranceReport {
        if self.deferred.is_some()
            && let Ok(materialized) = self.effective()
        {
            return materialized.tolerance_report();
        }
        self.tolerance.get_or_init(|| {
            let mut merged = self.open_tolerance.clone();
            // Content repairs (LENIENCY_AUDIT F9.10) are found by the slide
            // scan; a scan that fails outright reports its error through
            // `slides()`, and the report then lists the package repairs only.
            if let Ok((_slides, content)) =
                Parser::parse_slides_with_report(&self.content_xml, self.styles_xml.as_deref())
            {
                merged.append(&content);
            }
            merged
        })
    }

    /// Return the source identity captured during open.
    #[must_use = "the captured source version is needed to identify this source"]
    pub fn source_version(&self) -> Result<SourceVersion> {
        self.check_source()?;
        Ok(self.source_version)
    }

    /// Borrow the validated `content.xml` snapshot.
    pub fn content_xml(&self) -> Result<&str> {
        if self.deferred.is_some() {
            return self.effective()?.content_xml();
        }
        self.check_source()?;
        Ok(&self.content_xml)
    }

    /// Borrow the optional validated `styles.xml` snapshot.
    pub fn styles_xml(&self) -> Result<Option<&str>> {
        if self.deferred.is_some() {
            return self.effective()?.styles_xml();
        }
        self.check_source()?;
        Ok(self.styles_xml.as_deref())
    }

    /// Return common document metadata projected from the optional
    /// `meta.xml` part, matching [`super::Presentation::metadata`].
    pub fn metadata(&self) -> Result<litchi_core::Metadata> {
        if self.deferred.is_some() {
            return self.effective()?.metadata();
        }
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
        if self.deferred.is_some() {
            return self.effective()?.slides();
        }
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
            Parser::parse_slides_with_report(&self.content_xml, self.styles_xml.as_deref());
        let observed = self.source.version()?;
        ensure_current(self.source_version, observed)?;
        let (slides, content) = result?;
        if self.tolerance.get().is_none() {
            let mut merged = self.open_tolerance.clone();
            merged.append(&content);
            let _ = self.tolerance.set(merged);
        }
        Ok(slides)
    }

    /// Parse one slide while retaining the parser's complete-document
    /// validation semantics.
    pub fn slide(&self, index: usize) -> Result<Option<Slide>> {
        if self.deferred.is_some() {
            return self.effective()?.slide(index);
        }
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
        if self.deferred.is_some() {
            return self.effective()?.text();
        }
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

    /// Stream one semantic text object per slide into `output` in document order.
    ///
    /// The retained `content.xml` snapshot is used directly; package media is
    /// not read. Output may be partial when parsing, limits, or the sink fail.
    /// This method does not flush or roll back the sink and provides no
    /// cancellation hook. The paragraph separator applies to all paragraph,
    /// title/body, and shape-text boundaries within each slide; the slide
    /// separator applies between slide objects. Source freshness is checked
    /// before and after the pass, with a source-change error taking precedence while preserving
    /// exact progress.
    /// Semantic output parity is defined for documents accepted by the ordinary
    /// parser; the fused sink validates only its bounded XML/text contexts and
    /// does not promise full parser validation or error-precedence parity. Use
    /// `with_empty_objects(false)` to match [`super::Presentation::text`]
    /// filtering of empty slides.
    ///
    /// # Errors
    /// Returns a typed error containing exact partial-output progress.
    pub fn write_text_to<W: std::io::Write + ?Sized>(
        &self,
        output: &mut W,
        options: TextOutputOptions<'_>,
    ) -> std::result::Result<TextOutputReport, TextOutputError<Error>> {
        if self.deferred.is_some() {
            match self.effective() {
                Ok(materialized) => return materialized.write_text_to(output, options),
                Err(source) => {
                    return Err(TextOutputError::Document {
                        source,
                        progress: TextOutputReport::default(),
                    });
                },
            }
        }
        self.check_source()
            .map_err(|source| TextOutputError::Document {
                source,
                progress: TextOutputReport::default(),
            })?;
        let paragraph_separator = options.paragraph_separator();
        let mut writer = SequentialTextWriter::new(output, options);
        let result =
            Parser::write_slides_text_to(&self.content_xml, &mut writer, paragraph_separator);
        let progress = writer.progress();
        if let Err(source) = self.check_source() {
            return Err(TextOutputError::Document { source, progress });
        }
        result.map(|()| writer.finish())
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
        if self.deferred.is_some() {
            return self.effective()?.media_data(media);
        }
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

    /// Read one package member (an embedded object's part) on demand, or
    /// `None` when the package has no member at `path` (record 0328).
    ///
    /// # Errors
    /// Returns an error when the source changed or the member cannot be read.
    pub fn member_data(&self, path: &str) -> Result<Option<Vec<u8>>> {
        if self.deferred.is_some() {
            return self.effective()?.member_data(path);
        }
        self.check_source()?;
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
    ///
    /// # Errors
    ///
    /// Returns a source-changed, archive or family-validation failure.
    pub fn materialize(self) -> Result<crate::Package> {
        if self.deferred.is_some() {
            return self.effective()?.share().materialize();
        }
        self.package
            .materialize()
            .and_then(super::Presentation::from_owned_package)
            .map(crate::Package::from_presentation)
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

fn reconcile_source<T>(
    source: &dyn ReadAt,
    expected: SourceVersion,
    result: Result<T>,
) -> Result<T> {
    let observed = source.version()?;
    ensure_current(expected, observed)?;
    result
}

fn ensure_current(expected: SourceVersion, observed: SourceVersion) -> Result<()> {
    if expected == observed {
        Ok(())
    } else {
        Err(Error::SourceChanged { expected, observed })
    }
}

/// Check the caller's semantic part-size limits against the declared member
/// sizes before any part is inflated.
fn check_semantic_part_sizes(
    package: &SourceBackedPackage,
    limits: litchi_odf_common::SemanticLimits,
) -> std::result::Result<(), OdfError> {
    if limits.is_default() {
        return Ok(());
    }
    if let Some(size) = package.member_materialized_size("content.xml")? {
        limits.check(litchi_odf_common::SemanticLimitKind::ContentBytes, size)?;
    }
    if let Some(size) = package.member_materialized_size("styles.xml")? {
        limits.check(litchi_odf_common::SemanticLimitKind::StylesBytes, size)?;
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use std::io;

    use super::*;

    struct FixedSource;

    impl ReadAt for FixedSource {
        fn len(&self) -> io::Result<u64> {
            Ok(0)
        }

        fn read_at(&self, _offset: u64, _output: &mut [u8]) -> io::Result<usize> {
            Ok(0)
        }

        fn version(&self) -> io::Result<SourceVersion> {
            Ok(SourceVersion::new(0x4f44_5301, 1))
        }
    }

    #[test]
    fn source_reconciliation_precedes_secondary_parse_errors() {
        let expected = SourceVersion::new(0x4f44_5301, 0);
        let secondary: Result<()> = Err(Error::InvalidFormat("secondary parse error".to_string()));

        assert!(matches!(
            reconcile_source(&FixedSource, expected, secondary),
            Err(Error::SourceChanged { .. })
        ));
    }
}
