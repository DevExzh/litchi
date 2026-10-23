//! Immutable, positional-source ODP reads.
//!
//! [`SourceBackedPresentation`] is deliberately additive: the existing
//! [`super::Presentation`] keeps its owned-byte and mutation behavior.  This
//! facade validates the XML needed for ordinary semantic reads at open, then
//! keeps the source-backed projection lazy; package media remains cold until
//! [`SourceBackedPresentation::media_data`] selects a member.

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
    Meta, SourceBackedPackage, SourceContentPublicationError, SourceContentPublicationOptions,
    SourceContentPublicationReport, SourceMemberReaderError, XmlStreamLimits,
    private::ContentDocumentValidator,
};
use litchi_odf_common::{Error as OdfError, OpenOptions};
use zeroize::Zeroizing;

use super::source_append::ScanFacts;
use crate::codec::Parser;
use crate::model::{Reference, Slide};

const ODF_PRESENTATION: &str = "application/vnd.oasis.opendocument.presentation";
const BODY_MARKER: &str = "<office:presentation";
const FAMILY_NAME: &str = "ODP";
const SLIDES_CACHE_QUERY_THRESHOLD: usize = 2;
// Keep a ceiling for the optional semantic projection so a repeated query
// cannot turn an otherwise lazy source view into a large second model.
const SLIDES_CACHE_MAX_SOURCE_BYTES: usize = 4 * 1024 * 1024;
const SLIDES_CACHE_MAX_COUNT: usize = 256;
const TEXT_CACHE_QUERY_THRESHOLD: usize = 2;
const TEXT_CACHE_MAX_BYTES: usize = 16 * 1024 * 1024;
// A small source may safely retain its validated XML so repeated text writes
// reuse decompression. Larger sources stay reader-backed even when a caller
// repeats a query; this ceiling is independent of the public content_xml()
// accessor, which remains an explicit materialization request.
const SEMANTIC_CONTENT_MEMO_MAX_SOURCE_BYTES: u64 = 1024 * 1024;
const STREAM_PROBE_WINDOW_BYTES: usize = 4096;
const OFFICE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";

#[cfg(any(unix, windows))]
fn file_source(path: impl AsRef<Path>) -> Result<Arc<dyn ReadAt>> {
    Ok(Arc::new(FileSource::open(path)?))
}

#[derive(Debug)]
enum StreamProbeError {
    Invalid(Error),
    BodyMissing,
}

#[derive(Debug, Default)]
struct Utf8Probe {
    pending: [u8; 3],
    pending_len: usize,
}

impl Utf8Probe {
    fn feed(&mut self, bytes: &[u8], part: &str) -> std::result::Result<(), StreamProbeError> {
        let mut rest = bytes;
        loop {
            if self.pending_len != 0 {
                if rest.is_empty() {
                    return Ok(());
                }
                // A UTF-8 scalar is at most four bytes.  Check the pending
                // prefix with only a tiny stack boundary buffer so admission
                // does not allocate a second copy of every streamed chunk.
                let take = rest.len().min(3);
                let mut boundary = [0_u8; 6];
                let pending_len = self.pending_len;
                boundary[..pending_len].copy_from_slice(&self.pending[..pending_len]);
                boundary[pending_len..pending_len + take].copy_from_slice(&rest[..take]);
                match std::str::from_utf8(&boundary[..pending_len + take]) {
                    Ok(_) => {
                        self.pending_len = 0;
                        rest = &rest[take..];
                    },
                    Err(error) if error.error_len().is_none() => {
                        let tail = &boundary[error.valid_up_to()..pending_len + take];
                        debug_assert!(tail.len() <= self.pending.len());
                        self.pending[..tail.len()].copy_from_slice(tail);
                        self.pending_len = tail.len();
                        rest = &rest[take..];
                    },
                    Err(error) => {
                        return Err(StreamProbeError::Invalid(Error::InvalidFormat(format!(
                            "{FAMILY_NAME} {part} is not UTF-8: {error}"
                        ))));
                    },
                }
                continue;
            }

            match std::str::from_utf8(rest) {
                Ok(_) => return Ok(()),
                Err(error) if error.error_len().is_none() => {
                    let tail = &rest[error.valid_up_to()..];
                    debug_assert!(tail.len() <= self.pending.len());
                    self.pending[..tail.len()].copy_from_slice(tail);
                    self.pending_len = tail.len();
                    return Ok(());
                },
                Err(error) => {
                    return Err(StreamProbeError::Invalid(Error::InvalidFormat(format!(
                        "{FAMILY_NAME} {part} is not UTF-8: {error}"
                    ))));
                },
            }
        }
    }

    fn finish(self, part: &str) -> std::result::Result<(), StreamProbeError> {
        if self.pending_len == 0 {
            Ok(())
        } else {
            Err(StreamProbeError::Invalid(Error::InvalidFormat(format!(
                "{FAMILY_NAME} {part} is not valid UTF-8"
            ))))
        }
    }
}

#[derive(Debug)]
struct BodyProbe {
    window: [u8; STREAM_PROBE_WINDOW_BYTES],
    window_len: usize,
    alias_marker: Option<Vec<u8>>,
    found: bool,
    root_seen: bool,
    doctype: bool,
}

impl BodyProbe {
    fn new() -> Self {
        Self {
            window: [0; STREAM_PROBE_WINDOW_BYTES],
            window_len: 0,
            alias_marker: None,
            found: false,
            root_seen: false,
            doctype: false,
        }
    }

    fn feed(&mut self, bytes: &[u8]) {
        let mut bytes = bytes;
        while !bytes.is_empty() && !self.found {
            // Retain each incoming subchunk while it is scanned. A single
            // archive fill may be larger than the probe window; taking at
            // most half keeps both the previous suffix and incoming prefix
            // available for markers crossing that boundary.
            let take = bytes.len().min(STREAM_PROBE_WINDOW_BYTES / 2);
            let chunk = &bytes[..take];
            let total = self.window_len + chunk.len();
            if total > STREAM_PROBE_WINDOW_BYTES {
                let remove = total - STREAM_PROBE_WINDOW_BYTES;
                self.window.copy_within(remove..self.window_len, 0);
                self.window_len -= remove;
            }
            self.window[self.window_len..self.window_len + chunk.len()].copy_from_slice(chunk);
            self.window_len += chunk.len();

            if !self.root_seen {
                let window = &self.window[..self.window_len];
                let doctype_start = find_subslice(window, b"<!DOCTYPE");
                let root_start = root_tag_position(window);
                if doctype_start.is_some_and(|doctype| root_start.is_none_or(|root| doctype < root))
                {
                    self.doctype = true;
                }
                self.root_seen = root_start.is_some();
            }
            let body_found =
                find_subslice(&self.window[..self.window_len], BODY_MARKER.as_bytes()).is_some();
            if body_found {
                self.found = true;
            } else {
                self.find_alias_marker();
                let window = &self.window[..self.window_len];
                if let Some(marker) = self.alias_marker.as_deref()
                    && find_subslice(window, marker).is_some()
                {
                    self.found = true;
                }
            }
            bytes = &bytes[take..];
        }
    }

    fn find_alias_marker(&mut self) {
        if self.alias_marker.is_some() {
            return;
        }
        let window = &self.window[..self.window_len];
        let Some(uri_start) = find_subslice(window, OFFICE_NAMESPACE) else {
            return;
        };
        let Some(namespace_start) = find_subslice(&window[..uri_start], b"xmlns:") else {
            return;
        };
        let declaration = &window[namespace_start + b"xmlns:".len()..uri_start];
        let Some(equals) = declaration.iter().rposition(|byte| *byte == b'=') else {
            return;
        };
        let mut prefix = &declaration[..equals];
        while prefix.last().is_some_and(u8::is_ascii_whitespace) {
            prefix = &prefix[..prefix.len() - 1];
        }
        if prefix.is_empty()
            || !prefix
                .iter()
                .all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'_' | b'-' | b'.'))
        {
            return;
        }
        let mut marker = Vec::with_capacity(1 + prefix.len() + 1 + b"presentation".len());
        marker.push(b'<');
        marker.extend_from_slice(prefix);
        marker.extend_from_slice(b":presentation");
        self.alias_marker = Some(marker);
    }

    fn finish(self) -> std::result::Result<bool, StreamProbeError> {
        if self.found {
            Ok(self.doctype)
        } else {
            Err(StreamProbeError::BodyMissing)
        }
    }
}

fn find_subslice(haystack: &[u8], needle: &[u8]) -> Option<usize> {
    if needle.is_empty() {
        return Some(0);
    }
    haystack
        .windows(needle.len())
        .position(|window| window == needle)
}

fn root_tag_position(bytes: &[u8]) -> Option<usize> {
    let mut search_from = 0;
    while let Some(local) = find_subslice(&bytes[search_from..], b":document-content") {
        let local = search_from + local;
        let open = bytes[..local].iter().rposition(|byte| *byte == b'<')?;
        let prefix = &bytes[open + 1..local];
        if !prefix.is_empty()
            && prefix
                .iter()
                .all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'_' | b'-' | b'.'))
        {
            return Some(open);
        }
        search_from = local + 1;
    }
    None
}

fn encrypted_member_cannot_stream(error: &Error) -> bool {
    matches!(
        error,
        Error::InvalidFormat(message)
            if message.contains("cannot be streamed without decryption")
    )
}

fn source_reader_error(error: SourceMemberReaderError<Error>) -> Error {
    match error {
        SourceMemberReaderError::Core { error, .. } | SourceMemberReaderError::Callback(error) => {
            error
        },
        _ => Error::InvalidFormat("ODP content.xml source read failed".to_string()),
    }
}

fn validate_source_content(package: &SourceBackedPackage) -> Result<(u64, bool)> {
    let streamed = package.with_verified_member_reader("content.xml", |reader| {
        let mut bytes = 0_u64;
        let mut utf8 = Utf8Probe::default();
        let mut body = BodyProbe::new();
        loop {
            let chunk = reader
                .fill_buf()
                .map_err(|error| StreamProbeError::Invalid(Error::Io(error)))?;
            if chunk.is_empty() {
                break;
            }
            bytes = bytes.saturating_add(u64::try_from(chunk.len()).unwrap_or(u64::MAX));
            ContentDocumentValidator::check_materialized_size(Some(bytes), FAMILY_NAME)
                .map_err(StreamProbeError::Invalid)?;
            utf8.feed(chunk, "content.xml")?;
            body.feed(chunk);
            let consumed = chunk.len();
            reader.consume(consumed);
        }
        utf8.finish("content.xml")?;
        let doctype = body.finish()?;
        Ok((bytes, doctype))
    });

    match streamed {
        Ok((bytes, doctype)) => Ok((bytes, doctype)),
        Err(SourceMemberReaderError::Callback(StreamProbeError::BodyMissing)) => {
            validate_materialized_content(package)
        },
        Err(SourceMemberReaderError::Callback(StreamProbeError::Invalid(error))) => Err(error),
        Err(SourceMemberReaderError::Core {
            error,
            callback_error: _,
        }) if encrypted_member_cannot_stream(&error) => validate_materialized_content(package),
        Err(SourceMemberReaderError::Core { error, .. }) => Err(error),
        Err(_) => Err(Error::InvalidFormat(
            "ODP content.xml source validation failed".to_string(),
        )),
    }
}

fn validate_materialized_content(package: &SourceBackedPackage) -> Result<(u64, bool)> {
    let mut content_xml = String::from_utf8(package.get_file("content.xml")?).map_err(|error| {
        Error::InvalidFormat(format!("{FAMILY_NAME} content.xml is not UTF-8: {error}"))
    })?;
    let doctype = litchi_odf_common::core::private::neutralize_prologue_doctype(&mut content_xml);
    litchi_odf_common::core::validate_content_part(&content_xml, BODY_MARKER, FAMILY_NAME)?;
    Ok((
        u64::try_from(content_xml.len()).unwrap_or(u64::MAX),
        doctype,
    ))
}

fn validate_source_styles(package: &SourceBackedPackage) -> Result<()> {
    let streamed = package.with_verified_member_reader("styles.xml", |reader| {
        let mut utf8 = Utf8Probe::default();
        loop {
            let chunk = reader
                .fill_buf()
                .map_err(|error| StreamProbeError::Invalid(Error::Io(error)))?;
            if chunk.is_empty() {
                break;
            }
            utf8.feed(chunk, "styles.xml")?;
            let consumed = chunk.len();
            reader.consume(consumed);
        }
        utf8.finish("styles.xml")
    });

    match streamed {
        Ok(()) => Ok(()),
        Err(SourceMemberReaderError::Callback(StreamProbeError::Invalid(error))) => Err(error),
        Err(SourceMemberReaderError::Callback(StreamProbeError::BodyMissing)) => Err(
            Error::InvalidFormat("ODP styles.xml validation failed".to_string()),
        ),
        Err(SourceMemberReaderError::Core {
            error,
            callback_error: _,
        }) if encrypted_member_cannot_stream(&error) => {
            String::from_utf8(package.get_file("styles.xml")?).map_err(|error| {
                Error::InvalidFormat(format!("{FAMILY_NAME} styles.xml is not UTF-8: {error}"))
            })?;
            Ok(())
        },
        Err(SourceMemberReaderError::Core { error, .. }) => Err(error),
        Err(_) => Err(Error::InvalidFormat(
            "ODP styles.xml source validation failed".to_string(),
        )),
    }
}

/// Read-only ODP access over an immutable positional source.
///
/// Opening validates the ZIP archive, MIME member, and family content root,
/// then validates only `content.xml` and the optional `styles.xml`.  Their
/// semantic text is materialized on first use and retained only as a shared
/// cache. Other package members—including pictures and audio/video payloads—
/// remain deferred until [`Self::media_data`] is called. Every public
/// operation checks the captured [`SourceVersion`] before and after work,
/// returning [`Error::SourceChanged`] when the source revision is no longer
/// current.
///
/// This type has no edit or output methods.  Use [`super::Presentation`] for
/// the established owned-byte/mutation API.
pub struct SourceBackedPresentation {
    pub(super) package: Arc<SourceBackedPackage>,
    source: Arc<dyn ReadAt>,
    source_version: SourceVersion,
    content_xml_bytes: u64,
    append_facts: Arc<OnceLock<ScanFacts>>,
    content_xml: OnceLock<Arc<String>>,
    styles_xml: OnceLock<Option<Arc<String>>>,
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
            let (content_xml_bytes, doctype) = validate_source_content(&package)?;

            if package.has_file("styles.xml")? {
                // Validate the optional styles part at open, but leave its
                // semantic projection in the source until a caller asks for
                // it.  This keeps the open and first-read contracts strict
                // without retaining another potentially large XML string.
                validate_source_styles(&package)?;
            }

            // Match the ordinary family package's mandatory metadata
            // validation: a malformed optional meta.xml must not become
            // observable only after a source-backed facade is opened.
            let metadata = if package.has_file("meta.xml")? {
                // ODP owners retain no raw XML policy yet; meta.xml is
                // admitted under the documented owner default.
                Some(
                    Meta::from_bytes(
                        &package.get_file("meta.xml")?,
                        XmlStreamLimits::owner_default_value(),
                    )?
                    .try_extract_metadata()?,
                )
            } else {
                None
            };

            Ok((content_xml_bytes, metadata, doctype))
        })();

        // Reconcile the source before exposing any secondary parse error. This
        // keeps stale-source errors authoritative across the whole open pass,
        // including pure validation and UTF-8/metadata parsing work.
        let (content_xml_bytes, metadata, doctype) =
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
            content_xml_bytes,
            append_facts: Arc::new(OnceLock::new()),
            content_xml: OnceLock::new(),
            styles_xml: OnceLock::new(),
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
        let content_xml = clone_once_lock(&self.content_xml);
        let styles_xml = clone_once_lock(&self.styles_xml);
        Self {
            package: Arc::clone(&self.package),
            source: Arc::clone(&self.source),
            source_version: self.source_version,
            content_xml_bytes: self.content_xml_bytes,
            append_facts: Arc::clone(&self.append_facts),
            content_xml,
            styles_xml,
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
        self.effective_with_odf_error().map_err(Error::from)
    }

    /// The materialized presentation while retaining the common ODF
    /// publication chain for source-backed edit callers.
    pub(crate) fn effective_with_odf_error(&self) -> std::result::Result<&Self, OdfError> {
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
            .map_err(OdfError::from)?;
        let materialized = Self::from_read_at(Arc::new(litchi_core::OwnedSource::new(bytes)))?;
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

    fn read_content_xml(&self) -> Result<String> {
        let result = (|| {
            let mut content_xml = String::from_utf8(self.package.get_file("content.xml")?)
                .map_err(|error| {
                    Error::InvalidFormat(format!("{FAMILY_NAME} content.xml is not UTF-8: {error}"))
                })?;
            litchi_odf_common::core::private::neutralize_prologue_doctype(&mut content_xml);
            Ok(content_xml)
        })();
        reconcile_source(self.source.as_ref(), self.source_version, result)
    }

    fn load_content_xml(&self) -> Result<Arc<String>> {
        self.read_content_xml().map(Arc::new)
    }

    fn content_xml_snapshot(&self) -> Result<&str> {
        self.check_source()?;
        if self.content_xml.get().is_none() {
            let content_xml = self.load_content_xml()?;
            let _ = self.content_xml.set(content_xml);
        }
        self.check_source()?;
        Ok(self
            .content_xml
            .get()
            .expect("content.xml cache is initialized")
            .as_ref())
    }

    fn read_styles_xml(&self) -> Result<Option<String>> {
        let result = (|| {
            if !self.package.has_file("styles.xml")? {
                return Ok(None);
            }
            let styles_xml =
                String::from_utf8(self.package.get_file("styles.xml")?).map_err(|error| {
                    Error::InvalidFormat(format!("{FAMILY_NAME} styles.xml is not UTF-8: {error}"))
                })?;
            Ok(Some(styles_xml))
        })();
        reconcile_source(self.source.as_ref(), self.source_version, result)
    }

    fn load_styles_xml(&self) -> Result<Option<Arc<String>>> {
        self.read_styles_xml().map(|styles| styles.map(Arc::new))
    }

    fn styles_xml_snapshot(&self) -> Result<Option<&str>> {
        self.check_source()?;
        if self.styles_xml.get().is_none() {
            let styles_xml = self.load_styles_xml()?;
            let _ = self.styles_xml.set(styles_xml);
        }
        self.check_source()?;
        Ok(self
            .styles_xml
            .get()
            .expect("styles.xml cache is initialized")
            .as_ref()
            .map(Arc::as_ref)
            .map(String::as_str))
    }

    fn parse_slides_from_source(
        &self,
    ) -> Result<(Vec<Slide>, litchi_odf_common::leniency::ToleranceReport)> {
        self.check_source()?;
        // The optional styles part was validated during open.  Reuse its
        // bounded lazy projection for semantic reads so repeated slide/text
        // queries do not reopen and decode the same member on every call.
        let styles_xml = self.styles_xml_snapshot()?;
        let result = match self
            .package
            .with_verified_member_reader("content.xml", |reader| {
                Parser::parse_slides_with_report_from_reader(reader, styles_xml)
            }) {
            Ok(result) => Ok(result),
            Err(SourceMemberReaderError::Core { error, .. })
                if encrypted_member_cannot_stream(&error) =>
            {
                let content_xml = self.read_content_xml()?;
                Parser::parse_slides_with_report(&content_xml, styles_xml)
            },
            Err(error) => Err(source_reader_error(error)),
        };
        let observed = self.source.version()?;
        ensure_current(self.source_version, observed)?;
        result
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
            None => {
                let content_xml = self.content_xml_snapshot().map_err(|error| {
                    SourceContentInsertionError::from(SourceContentPublicationError::from(error))
                })?;
                Ok(self.package.write_content_xml_to_stream_with_options(
                    writer,
                    content_xml.as_bytes(),
                    options,
                )?)
            },
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

    /// Reuse the bounded structural facts collected during strict open when
    /// the append operation has no per-operation accounting or cancellation
    /// policy.  The common insertion plan still rechecks source freshness and
    /// proves the complete source hash and exact splice.
    pub(super) fn cached_append_facts(
        &self,
        max_content_xml_bytes: u64,
        max_slides: usize,
        options: &SourceContentPublicationOptions,
    ) -> Result<Option<ScanFacts>> {
        // Keep the cache path's error timing aligned with `scan_source`: an
        // out-of-envelope source limit must fail before authored-fragment
        // construction, even when the owner already has structural facts.
        if XmlStreamLimits::new(max_content_xml_bytes, 256, 4_000_000, 1_000_000, 64 * 1024)
            .is_err()
        {
            return Ok(None);
        }
        if options.execution_context().is_some() || options.cancellation().is_some() {
            return Ok(None);
        }
        self.check_source()?;
        let Some(facts) = self.append_facts.get() else {
            return Ok(None);
        };
        if facts.content_bytes > max_content_xml_bytes || facts.slide_count >= max_slides {
            return Ok(None);
        }
        Ok(Some(facts.clone()))
    }

    /// Retain append facts after a successful, uncancelled source scan. The
    /// facts contain only bounded offsets and counts; the common insertion
    /// plan still performs its independent freshness and byte proof.
    pub(super) fn remember_append_facts(
        &self,
        facts: &ScanFacts,
        options: &SourceContentPublicationOptions,
    ) {
        if options.execution_context().is_some() || options.cancellation().is_some() {
            return;
        }
        let _ = self.append_facts.set(facts.clone());
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
            if let Ok((_slides, content)) = self.parse_slides_from_source() {
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
        self.content_xml_snapshot()
    }

    /// Borrow the optional validated `styles.xml` snapshot.
    pub fn styles_xml(&self) -> Result<Option<&str>> {
        if self.deferred.is_some() {
            return self.effective()?.styles_xml();
        }
        self.styles_xml_snapshot()
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
        let (slides, content) = self.parse_slides_from_source()?;
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
        // Keep indexed semantic reads on the same validated lazy styles
        // projection as full slide scans.  This also memoizes an absent
        // styles.xml member, avoiding repeated archive admission work.
        let styles_xml = self.styles_xml_snapshot()?;
        let result = match self
            .package
            .with_verified_member_reader("content.xml", |reader| {
                Parser::parse_slide_with_styles_at_recording_from_reader(
                    reader,
                    styles_xml,
                    index,
                    &mut litchi_odf_common::leniency::ToleranceReport::default(),
                )
            }) {
            Ok(result) => Ok(result),
            Err(SourceMemberReaderError::Core { error, .. })
                if encrypted_member_cannot_stream(&error) =>
            {
                let content_xml = self.read_content_xml()?;
                Parser::parse_slide_with_styles_at(&content_xml, styles_xml, index)
            },
            Err(error) => Err(source_reader_error(error)),
        };
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
        if self.content_xml_bytes > SLIDES_CACHE_MAX_SOURCE_BYTES as u64
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
    /// The validated source content is decoded for this pass; when its
    /// declared size fits the bounded small-source memo it may be retained
    /// for a repeated text write, while larger sources remain streaming.
    /// Package media is not read. Output may be partial when parsing, limits,
    /// or the sink fail.
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
        if let Some(content_xml) =
            self.small_content_xml_snapshot()
                .map_err(|source| TextOutputError::Document {
                    source,
                    progress: writer.progress(),
                })?
        {
            let result =
                Parser::write_slides_text_to(content_xml, &mut writer, paragraph_separator);
            let progress = writer.progress();
            if let Err(source) = self.check_source() {
                return Err(TextOutputError::Document { source, progress });
            }
            return result.map(|()| writer.finish());
        }
        let result = self
            .package
            .with_verified_member_reader("content.xml", |reader| {
                Parser::write_slides_text_from_reader(reader, &mut writer, paragraph_separator)
            });
        let result = match result {
            Ok(()) => Ok(()),
            Err(SourceMemberReaderError::Core { error, .. })
                if encrypted_member_cannot_stream(&error) =>
            {
                match self.read_content_xml() {
                    Ok(content_xml) => {
                        Parser::write_slides_text_to(&content_xml, &mut writer, paragraph_separator)
                    },
                    Err(source) => Err(TextOutputError::Document {
                        source,
                        progress: writer.progress(),
                    }),
                }
            },
            Err(SourceMemberReaderError::Callback(error)) => Err(error),
            Err(SourceMemberReaderError::Core {
                error,
                callback_error,
            }) => Err(TextOutputError::Document {
                source: error,
                progress: callback_error
                    .as_ref()
                    .map(TextOutputError::progress)
                    .unwrap_or_else(|| writer.progress()),
            }),
            Err(_) => Err(TextOutputError::Document {
                source: Error::InvalidFormat("ODP content.xml source text read failed".to_string()),
                progress: writer.progress(),
            }),
        };
        let progress = writer.progress();
        if let Err(source) = self.check_source() {
            return Err(TextOutputError::Document { source, progress });
        }
        result.map(|()| writer.finish())
    }

    /// Reuse one bounded decoded content member for repeated text writes.
    /// The source-size check is performed before materialization, and the
    /// existing content snapshot retains source-version reconciliation around
    /// the read and publication.
    fn small_content_xml_snapshot(&self) -> Result<Option<&str>> {
        if self.content_xml_bytes > SEMANTIC_CONTENT_MEMO_MAX_SOURCE_BYTES {
            return Ok(None);
        }
        self.content_xml_snapshot().map(Some)
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

fn clone_once_lock<T: Clone>(source: &OnceLock<T>) -> OnceLock<T> {
    let cloned = OnceLock::new();
    if let Some(value) = source.get() {
        let _ = cloned.set(value.clone());
    }
    cloned
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

    use litchi_core::OwnedSource;
    use litchi_odf_common::constants::ODF_PRESENTATION;
    use litchi_odf_common::core::PackageWriter;

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

    #[test]
    fn streamed_admission_probes_keep_utf8_boundaries_without_heap_scratch() {
        let bytes = "prefix € and 𐍈 suffix".as_bytes();
        for split in 0..=bytes.len() {
            let mut probe = Utf8Probe::default();
            probe.feed(&bytes[..split], "content.xml").unwrap();
            probe.feed(&bytes[split..], "content.xml").unwrap();
            probe.finish("content.xml").unwrap();
        }

        let mut invalid = Utf8Probe::default();
        invalid.feed(&[0xe2], "content.xml").unwrap();
        assert!(matches!(
            invalid.feed(&[0x28], "content.xml"),
            Err(StreamProbeError::Invalid(Error::InvalidFormat(message)))
                if message.contains("not UTF-8")
        ));
    }

    #[test]
    fn streamed_body_probe_scans_a_fill_larger_than_its_window() {
        let mut probe = BodyProbe::new();
        let mut bytes = vec![b'x'; STREAM_PROBE_WINDOW_BYTES * 3];
        let marker = b"<office:body><office:presentation";
        bytes[128..128 + marker.len()].copy_from_slice(marker);
        probe.feed(&bytes);
        assert!(!probe.finish().unwrap());
    }

    #[test]
    fn source_owner_drops_large_validation_buffer_before_semantic_read() {
        let content = format!(
            r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><office:body><office:presentation><draw:page><draw:frame><draw:text-box><text:p>{}</text:p></draw:text-box></draw:frame></draw:page></office:presentation></office:body></office:document-content>"#,
            "x".repeat(1024 * 1024),
        );
        let mut writer =
            PackageWriter::new(litchi_odf_common::core::StreamingArchiveLimits::default());
        writer.set_mimetype(ODF_PRESENTATION).unwrap();
        writer.add_file("content.xml", content.as_bytes()).unwrap();
        let bytes = writer.finish_to_bytes().unwrap();
        let owner =
            SourceBackedPresentation::from_read_at(Arc::new(OwnedSource::new(bytes))).unwrap();

        assert!(owner.content_xml.get().is_none());
        assert_eq!(owner.content_xml_bytes, content.len() as u64);

        assert!(owner.slide(0).unwrap().is_some());
        let mut output = Vec::new();
        owner
            .write_text_to(&mut output, TextOutputOptions::default())
            .unwrap();
        assert!(
            owner.content_xml.get().is_none(),
            "semantic projections must not retain the raw XML cache"
        );
        assert!(owner.content_xml().unwrap().contains(&"x".repeat(1024)));
        assert!(owner.content_xml.get().is_some());
    }

    #[test]
    fn append_facts_are_deferred_until_a_successful_prepare() {
        let content = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"><office:body><office:presentation><draw:page/></office:presentation></office:body></office:document-content>"#;
        let mut writer =
            PackageWriter::new(litchi_odf_common::core::StreamingArchiveLimits::default());
        writer.set_mimetype(ODF_PRESENTATION).unwrap();
        writer.add_file("content.xml", content.as_bytes()).unwrap();
        let owner = SourceBackedPresentation::from_read_at(Arc::new(OwnedSource::new(
            writer.finish_to_bytes().unwrap(),
        )))
        .unwrap();
        assert!(owner.append_facts.get().is_none());

        owner
            .edit_tail_append("title", "body")
            .unwrap()
            .prepare()
            .unwrap();
        assert!(owner.append_facts.get().is_some());
    }
}
