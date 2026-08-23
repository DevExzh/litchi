//! Best-effort `OpenDocument Format` (`ODF`) detection.
//!
//! Detection is inert: it reads the standardized `MIME` type from a flat `XML`
//! root or packaged `mimetype` member without constructing a document model.
//!
//! ```rust
//! use litchi_odf_common::detect::{self, Format};
//!
//! assert_eq!(
//!     detect::mime(b"application/vnd.oasis.opendocument.text"),
//!     Some(Format::Odt),
//! );
//! ```

use crate::constants;
use crate::core::{OwnedPackage, PreparedPackage, package::read_owned_input};
use litchi_core::{Error, ReadAt, Resource, ResourceLimit, Result, SourceVersion};
use quick_xml::XmlVersion;
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::HashSet;
use std::io::{Read, Seek, SeekFrom};

/// Neutral file classification returned by the detector.
pub use litchi_core::FileFormat as Format;

const OFFICE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const ZIP_SIGNATURE: &[u8] = b"PK\x03\x04";
const LOCAL_HEADER_BYTES: usize = 30;
const MIMETYPE_PATH: &str = "mimetype";
const MIMETYPE_NAME: &[u8] = MIMETYPE_PATH.as_bytes();
const OOXML_CATALOG_NAME: &str = "[Content_Types].xml";
const MAX_MIMETYPE_BYTES: usize = 256;

/// Resource ceilings for the metadata-only OPC catalog probe.
///
/// The probe is deliberately conservative: a ceiling hit returns `None` so
/// the caller can run its normal bounded OPC path. The fields mirror the ZIP
/// portion of the facade's OPC `ReadLimits` but remain owned by this crate to
/// avoid a dependency cycle.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct CatalogProbeLimits {
    max_input_bytes: u64,
    max_members: usize,
    max_member_name_bytes: u64,
    max_metadata_bytes: u64,
    max_compressed_bytes: u64,
    max_entry_bytes: u64,
    max_total_bytes: u64,
}

impl CatalogProbeLimits {
    /// Construct a catalog-probe policy from the input and ZIP ceilings.
    #[must_use]
    pub const fn new(
        max_input_bytes: u64,
        max_members: usize,
        max_member_name_bytes: u64,
        max_metadata_bytes: u64,
        max_compressed_bytes: u64,
        max_entry_bytes: u64,
        max_total_bytes: u64,
    ) -> Self {
        Self {
            max_input_bytes,
            max_members,
            max_member_name_bytes,
            max_metadata_bytes,
            max_compressed_bytes,
            max_entry_bytes,
            max_total_bytes,
        }
    }
}

impl Default for CatalogProbeLimits {
    fn default() -> Self {
        Self::new(
            512 * 1024 * 1024,
            100_000,
            4 * 1024,
            64 * 1024 * 1024,
            512 * 1024 * 1024,
            512 * 1024 * 1024,
            2 * 1024 * 1024 * 1024,
        )
    }
}

/// Maximum input accepted by the compatibility detector defaults.
pub const DEFAULT_MAX_INPUT_BYTES: u64 = 2 * 1024 * 1024 * 1024;

/// Finite limits for ODF detection and stream ingress.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    max_input_bytes: u64,
}

impl Limits {
    /// Construct a detector profile with an explicit input ceiling.
    #[must_use]
    pub const fn new(max_input_bytes: u64) -> Self {
        Self { max_input_bytes }
    }

    /// Return the maximum source bytes accepted by this profile.
    #[must_use]
    pub const fn max_input_bytes(self) -> u64 {
        self.max_input_bytes
    }

    /// Return a copy with a different input ceiling.
    #[must_use]
    pub const fn with_max_input_bytes(mut self, maximum: u64) -> Self {
        self.max_input_bytes = maximum;
        self
    }

    fn validate(self) -> Result<()> {
        if self.max_input_bytes == 0 || self.max_input_bytes > DEFAULT_MAX_INPUT_BYTES {
            return Err(Error::InvalidFormat(format!(
                "ODF detector input bytes must be between 1 and {DEFAULT_MAX_INPUT_BYTES}"
            )));
        }
        Ok(())
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self::new(DEFAULT_MAX_INPUT_BYTES)
    }
}

/// Classify the raw contents of an ODF `mimetype` member.
///
/// Leading and trailing ASCII whitespace is ignored without allocating.
/// Unknown MIME types and invalid UTF-8 return `None`.
#[inline]
#[must_use]
pub fn mime(value: &[u8]) -> Option<Format> {
    match std::str::from_utf8(trim_ascii(value)).ok()? {
        constants::ODF_TEXT | constants::ODF_TEXT_TEMPLATE => Some(Format::Odt),
        constants::ODF_SPREADSHEET | constants::ODF_SPREADSHEET_TEMPLATE => Some(Format::Ods),
        constants::ODF_PRESENTATION | constants::ODF_PRESENTATION_TEMPLATE => Some(Format::Odp),
        constants::ODF_DRAWING | constants::ODF_DRAWING_TEMPLATE => Some(Format::Odg),
        constants::ODF_CHART | constants::ODF_CHART_TEMPLATE => Some(Format::Odc),
        constants::ODF_FORMULA | constants::ODF_FORMULA_TEMPLATE => Some(Format::Odf),
        constants::ODF_IMAGE | constants::ODF_IMAGE_TEMPLATE => Some(Format::Odi),
        constants::ODF_MASTER | constants::ODF_MASTER_TEMPLATE => Some(Format::Odm),
        constants::ODF_WEB => Some(Format::Oth),
        constants::ODF_DATABASE => Some(Format::Odb),
        _ => None,
    }
}

/// Read a recognized `office:mimetype` value from a flat `ODF` `XML` root.
///
/// The root must be `office:document` in the ODF office namespace. Namespace
/// prefixes are resolved semantically. The returned value is decoded and XML
/// attribute whitespace is normalized before classification, then copied into
/// an owned string.
#[must_use]
pub fn flat_mime(value: &[u8]) -> Option<String> {
    flat_mime_with_limits(value, Limits::default())
        .ok()
        .flatten()
}

/// Read a recognized flat `ODF` MIME value under explicit finite limits.
pub fn flat_mime_with_limits(value: &[u8], limits: Limits) -> Result<Option<String>> {
    check_input_len(value.len(), limits, "ODF flat detector input")?;
    with_flat_mime(value, |raw_mimetype| {
        let trimmed_mimetype = trim_ascii(raw_mimetype);
        if mime(trimmed_mimetype).is_none() {
            return Ok(None);
        }
        let Some(mimetype_text) = std::str::from_utf8(trimmed_mimetype).ok() else {
            return Ok(None);
        };
        let mut owned = String::new();
        owned
            .try_reserve_exact(mimetype_text.len())
            .map_err(|source| Error::Allocation {
                resource: "ODF flat detector mimetype",
                source,
            })?;
        owned.push_str(mimetype_text);
        Ok(Some(owned))
    })
}

fn with_flat_mime<T>(
    value: &[u8],
    classify: impl FnOnce(&[u8]) -> Result<Option<T>>,
) -> Result<Option<T>> {
    let mut reader = NsReader::from_reader(value);
    loop {
        let (event_namespace, event) = match reader.read_resolved_event() {
            Ok(event) => event,
            Err(_) => return Ok(None),
        };
        match event {
            Event::Start(element) | Event::Empty(element) => {
                if !matches!(event_namespace, ResolveResult::Bound(Namespace(uri)) if uri == OFFICE_NAMESPACE)
                    || element.local_name().as_ref() != b"document"
                {
                    return Ok(None);
                }
                for raw_attribute in element.attributes() {
                    let Ok(attribute) = raw_attribute else {
                        return Ok(None);
                    };
                    let (attribute_namespace, local_name) =
                        reader.resolver().resolve_attribute(attribute.key);
                    if matches!(attribute_namespace, ResolveResult::Bound(Namespace(uri)) if uri == OFFICE_NAMESPACE)
                        && local_name.as_ref() == b"mimetype"
                    {
                        let Ok(decoded_mimetype) = attribute.decoded_and_normalized_value(
                            XmlVersion::Implicit1_0,
                            reader.decoder(),
                        ) else {
                            return Ok(None);
                        };
                        return classify(decoded_mimetype.as_bytes());
                    }
                }
                return Ok(None);
            },
            Event::Decl(_) | Event::Comment(_) | Event::DocType(_) | Event::PI(_) => {},
            Event::Text(text) if text.iter().all(u8::is_ascii_whitespace) => {},
            Event::Eof => return Ok(None),
            Event::End(_) | Event::Text(_) | Event::CData(_) | Event::GeneralRef(_) => {
                return Ok(None);
            },
        }
    }
}

/// Detect a flat `OpenDocument` `XML` document.
#[inline]
#[must_use]
pub fn flat(value: &[u8]) -> Option<Format> {
    flat_with_limits(value, Limits::default()).ok().flatten()
}

/// Detect a flat `OpenDocument` XML document under explicit finite limits.
pub fn flat_with_limits(value: &[u8], limits: Limits) -> Result<Option<Format>> {
    check_input_len(value.len(), limits, "ODF flat detector input")?;
    with_flat_mime(value, |raw_mimetype| Ok(mime(raw_mimetype)))
}

/// Detect a packaged or flat `OpenDocument` document from complete bytes.
///
/// Conforming ODF packages place an uncompressed `mimetype` entry first. The
/// local ZIP header and payload are checked in place without allocating a
/// decompression buffer. The central archive structure is also validated.
#[must_use]
pub fn bytes(value: &[u8]) -> Option<Format> {
    bytes_with_limits(value, Limits::default()).ok().flatten()
}

/// Detect a packaged or flat `OpenDocument` document under explicit finite
/// limits.
pub fn bytes_with_limits(value: &[u8], limits: Limits) -> Result<Option<Format>> {
    check_input_len(value.len(), limits, "ODF detector input")?;
    if value.starts_with(ZIP_SIGNATURE) {
        let Some(format) = packaged_mime_with_limits(value, limits)? else {
            return Ok(None);
        };
        let archive =
            soapberry_zip::office::ArchiveReader::new(value).map_err(map_detector_zip_error)?;
        if !archive
            .is_stored(MIMETYPE_PATH)
            .map_err(map_detector_zip_error)?
        {
            return Ok(None);
        }
        return Ok(Some(format));
    }
    flat_with_limits(value, limits)
}

/// Classify a packaged ODF candidate from its local `mimetype` entry only.
///
/// Unlike [`bytes`], this helper does not inspect the central directory. It is
/// intended for an owner that will immediately run [`prepared_or_original`]
/// and build the bounded archive index exactly once. A nonconforming local
/// header, unknown MIME type, or non-ZIP input returns `None`.
#[must_use]
pub fn packaged_mime(value: &[u8]) -> Option<Format> {
    packaged_mime_with_limits(value, Limits::default())
        .ok()
        .flatten()
}

/// Classify a packaged ODF candidate from its local `mimetype` entry under
/// explicit finite limits.
pub fn packaged_mime_with_limits(value: &[u8], limits: Limits) -> Result<Option<Format>> {
    check_input_len(value.len(), limits, "ODF detector input")?;
    Ok(packaged_mime_bytes(value).and_then(mime))
}

/// Check whether a ZIP candidate has the reserved OPC content-types member.
///
/// This inspects only the central-directory catalog. It does not read or
/// decompress a member and does not build an owned archive index, so a normal
/// ODF package can avoid an unrelated full OPC probe. `Some(false)` means the
/// ZIP catalog was valid and had no exact content-types member; `None` means
/// the candidate was not a canonical, in-budget ZIP catalog. Any malformed,
/// aliased, duplicate, encrypted, data-descriptor, ZIP64, prefixed, or
/// trailing layout is deliberately unknown so the caller can invoke its
/// existing bounded OPC path.
#[must_use]
pub fn packaged_has_ooxml_catalog(value: &[u8]) -> Option<bool> {
    packaged_has_ooxml_catalog_with_limits(value, CatalogProbeLimits::default())
}

/// Check the OPC content-types member with explicit input and ZIP budgets.
///
/// A successful `Some(false)` is reserved for a canonical catalog that stays
/// within every supplied ceiling. Returning `None` is intentionally
/// conservative: callers must continue with their ordinary bounded OPC
/// detector rather than treating the candidate as ordinary ODF.
#[must_use]
pub fn packaged_has_ooxml_catalog_with_limits(
    value: &[u8],
    limits: CatalogProbeLimits,
) -> Option<bool> {
    let length = u64::try_from(value.len()).ok()?;
    if length > limits.max_input_bytes {
        return None;
    }
    if !value.starts_with(ZIP_SIGNATURE) {
        return None;
    }

    let archive = soapberry_zip::ZipArchive::from_slice(value).ok()?;
    if archive.is_zip64() || archive.end_offset() != length {
        return None;
    }

    let mut seen = HashSet::<&[u8]>::new();
    let mut state = CatalogProbeState::default();
    let mut entries = archive.entries();
    while let Some(entry) = entries.next_entry().ok()? {
        let raw = entry.file_path().as_bytes();
        let normalized = entry.file_path().try_normalize().ok()?;
        if normalized.as_str().as_bytes() != raw {
            return None;
        }
        if !catalog_probe_range_is_bounded(&entry, archive.directory_offset()) {
            return None;
        }
        if !catalog_probe_entry_with_limits(&mut state, &entry, limits) {
            return None;
        }
        if !entry.is_dir() {
            seen.try_reserve(1).ok()?;
            if !seen.insert(raw) {
                return None;
            }
        }
        if !entry.is_dir() && normalized.as_str().eq_ignore_ascii_case(OOXML_CATALOG_NAME) {
            state.found_catalog = true;
        }
    }
    state.finish()
}

/// Check the reserved OPC content-types member from a positional ZIP source.
///
/// This is the source-backed counterpart to [`packaged_has_ooxml_catalog`]. It
/// reads only ZIP metadata through [`ReadAt`], retaining the caller's source
/// identity check and never materializing or decompressing a member. ZIP
/// layout errors are reported as `Ok(None)` so an owning facade can retain its
/// existing fallback/error policy for malformed candidates.
///
/// # Errors
///
/// Returns source I/O, source-change, or bounded scratch-buffer allocation
/// errors.
pub fn packaged_has_ooxml_catalog_read_at(
    source: &dyn ReadAt,
) -> Result<Option<bool>> {
    packaged_has_ooxml_catalog_read_at_with_limits(source, CatalogProbeLimits::default())
}

/// Check the OPC content-types member from a positional source under bounds.
///
/// ZIP layout errors become `Ok(None)`, while source-provider I/O and source
/// identity changes remain typed errors. This distinction prevents a failing
/// source from being silently reclassified as an ordinary ODF package.
pub fn packaged_has_ooxml_catalog_read_at_with_limits(
    source: &dyn ReadAt,
    limits: CatalogProbeLimits,
) -> Result<Option<bool>> {
    let expected = source.version()?;
    let detected = (|| {
        let length = source.len()?;
        if length > limits.max_input_bytes {
            return Ok(None);
        }
        if length < u64::try_from(ZIP_SIGNATURE.len()).expect("ZIP signature fits in u64") {
            return Ok(None);
        }

        let mut signature = [0_u8; ZIP_SIGNATURE.len()];
        if !read_at_catalog_exact(source, 0, &mut signature)? {
            return Ok(None);
        }
        if signature.as_slice() != ZIP_SIGNATURE {
            return Ok(None);
        }

        let mut buffer = Vec::new();
        buffer
            .try_reserve_exact(soapberry_zip::RECOMMENDED_BUFFER_SIZE)
            .map_err(|source| Error::Allocation {
                resource: "ODF ZIP catalog probe",
                source,
            })?;
        buffer.resize(soapberry_zip::RECOMMENDED_BUFFER_SIZE, 0);
        let archive = match soapberry_zip::ZipLocator::new().locate_in_reader(
            ReadAtZipSource { source },
            &mut buffer,
            length,
        ) {
            Ok(archive) => archive,
            Err((_reader, error)) => return map_catalog_probe_zip_error(error),
        };
        if archive.is_zip64() || archive.end_offset() != length {
            return Ok(None);
        }

        let mut seen = HashSet::<Vec<u8>>::new();
        let mut state = CatalogProbeState::default();
        let mut entries = archive.entries(&mut buffer);
        while let Some(entry) = match entries.next_entry() {
            Ok(entry) => entry,
            Err(error) => return map_catalog_probe_zip_error(error),
        } {
            let raw = entry.file_path().as_bytes();
            let normalized = match entry.file_path().try_normalize() {
                Ok(normalized) => normalized,
                Err(_error) => return Ok(None),
            };
            if normalized.as_str().as_bytes() != raw {
                return Ok(None);
            }
            if !catalog_probe_range_is_bounded(&entry, archive.directory_offset()) {
                return Ok(None);
            }
            if !catalog_probe_entry_with_limits(&mut state, &entry, limits) {
                return Ok(None);
            }
            if !entry.is_dir() {
                seen.try_reserve(1).map_err(|source| Error::Allocation {
                    resource: "ODF ZIP catalog names",
                    source,
                })?;
                let mut owned = Vec::new();
                owned
                    .try_reserve_exact(raw.len())
                    .map_err(|source| Error::Allocation {
                        resource: "ODF ZIP catalog name",
                        source,
                    })?;
                owned.extend_from_slice(raw);
                if !seen.insert(owned) {
                    return Ok(None);
                }
                if normalized.as_str().eq_ignore_ascii_case(OOXML_CATALOG_NAME) {
                    state.found_catalog = true;
                }
            }
        }
        Ok(state.finish())
    })();
    let observed = source.version()?;
    ensure_source_current(expected, observed)?;
    detected
}

/// Check the reserved OPC content-types member from a seekable reader.
///
/// The reader's original position is restored before returning. As with the
/// other catalog probes, this reads ZIP central-directory metadata only and
/// returns `None` for non-ZIP or malformed candidates.
pub fn packaged_has_ooxml_catalog_from_reader<R: Read + Seek>(reader: &mut R) -> Option<bool> {
    packaged_has_ooxml_catalog_from_reader_with_limits(reader, CatalogProbeLimits::default())
}

/// Check the OPC content-types member from a seekable reader under bounds.
///
/// The reader's cursor is restored even when the bounded metadata probe is
/// uncertain or an underlying read/seek operation fails.
pub fn packaged_has_ooxml_catalog_from_reader_with_limits<R: Read + Seek>(
    reader: &mut R,
    limits: CatalogProbeLimits,
) -> Option<bool> {
    let original = reader.stream_position().ok()?;
    let detected = (|| {
        let end = reader.seek(SeekFrom::End(0)).ok()?;
        if end > limits.max_input_bytes {
            return None;
        }
        reader.seek(SeekFrom::Start(0)).ok()?;
        let mut signature = [0_u8; ZIP_SIGNATURE.len()];
        reader.read_exact(&mut signature).ok()?;
        if signature.as_slice() != ZIP_SIGNATURE {
            return None;
        }
        let mut buffer = Vec::new();
        buffer
            .try_reserve_exact(soapberry_zip::RECOMMENDED_BUFFER_SIZE)
            .ok()?;
        buffer.resize(soapberry_zip::RECOMMENDED_BUFFER_SIZE, 0);
        let archive = soapberry_zip::ZipArchive::from_seekable(&mut *reader, &mut buffer).ok()?;
        if archive.is_zip64() || archive.end_offset() != end {
            return None;
        }
        let mut seen = HashSet::<Vec<u8>>::new();
        let mut state = CatalogProbeState::default();
        let mut entries = archive.entries(&mut buffer);
        while let Some(entry) = entries.next_entry().ok()? {
            let raw = entry.file_path().as_bytes();
            let normalized = entry.file_path().try_normalize().ok()?;
            if normalized.as_str().as_bytes() != raw {
                return None;
            }
            if !catalog_probe_range_is_bounded(&entry, archive.directory_offset()) {
                return None;
            }
            if !catalog_probe_entry_with_limits(&mut state, &entry, limits) {
                return None;
            }
            if !entry.is_dir() {
                seen.try_reserve(1).ok()?;
                let mut owned = Vec::new();
                owned.try_reserve_exact(raw.len()).ok()?;
                owned.extend_from_slice(raw);
                if !seen.insert(owned) {
                    return None;
                }
                if normalized.as_str().eq_ignore_ascii_case(OOXML_CATALOG_NAME) {
                    state.found_catalog = true;
                }
            }
        }
        state.finish()
    })();
    reader.seek(SeekFrom::Start(original)).ok()?;
    detected
}

/// Detect a packaged ODF MIME type from a positional source.
///
/// This probe reads only the fixed local header, the `mimetype` name, and its
/// bounded stored payload.  It does not allocate or inspect the central
/// directory, and it leaves ownership and complete package validation to the
/// family facade.  The source identity is checked before and after the probe;
/// a source that changes while it is being read is rejected instead of being
/// classified from a mixed snapshot.
pub fn packaged_mime_read_at(source: &dyn ReadAt) -> Result<Option<Format>> {
    packaged_mime_read_at_with_limits(source, Limits::default())
}

/// Detect a packaged ODF document from a positional source under explicit
/// finite limits.
pub fn packaged_mime_read_at_with_limits(
    source: &dyn ReadAt,
    limits: Limits,
) -> Result<Option<Format>> {
    limits.validate()?;
    let expected = source.version()?;
    let detected = (|| {
        let length = source.len()?;
        if length > limits.max_input_bytes {
            return Err(Error::ResourceLimit(ResourceLimit {
                resource: Resource::InputBytes,
                observed: length,
                limit: limits.max_input_bytes,
                scope: "ODF detector input".into(),
            }));
        }
        let local_name_end = u64::try_from(LOCAL_HEADER_BYTES + MIMETYPE_NAME.len())
            .expect("fixed local header size fits in u64");
        if length < local_name_end {
            return Ok(None);
        }

        let mut header = [0_u8; LOCAL_HEADER_BYTES];
        source.read_exact_at(0, &mut header)?;
        let mut name = [0_u8; MIMETYPE_NAME.len()];
        source.read_exact_at(LOCAL_HEADER_BYTES as u64, &mut name)?;
        let Some((data_start, compressed, expected_crc)) = packaged_mime_layout(&header, &name)
        else {
            return Ok(None);
        };

        let data_start = u64::try_from(data_start).expect("fixed local header size fits in u64");
        let data_end = data_start
            .checked_add(u64::try_from(compressed).expect("bounded MIME size fits in u64"))
            .ok_or_else(|| {
                std::io::Error::new(std::io::ErrorKind::InvalidData, "mimetype range overflow")
            })?;
        if data_end > length {
            return Ok(None);
        }

        // The local grammar bounds this to MAX_MIMETYPE_BYTES before any
        // source read.  A fixed array keeps this detector allocation-free.
        let mut data = [0_u8; MAX_MIMETYPE_BYTES];
        source.read_exact_at(data_start, &mut data[..compressed])?;
        if soapberry_zip::crc32(&data[..compressed]) != expected_crc {
            return Ok(None);
        }
        Ok(mime(&data[..compressed]))
    })();
    let observed = source.version()?;
    ensure_source_current(expected, observed)?;
    detected
}

/// Detect a packaged ODF document while retaining its validated ZIP index.
///
/// This is the ownership-taking counterpart to [`bytes`]. The local
/// `mimetype` framing contract is checked before the bounded archive index is
/// built, and the retained index is transferred to a concrete family facade
/// without a second central-directory scan.
#[must_use]
pub fn prepared(value: Vec<u8>) -> Option<PreparedPackage> {
    prepared_or_original(value).ok()
}

/// Detect a packaged ODF document while preserving ownership on rejection.
///
/// A successful result retains the one validated ZIP index built by the
/// detector. A rejected candidate returns the caller's original `Vec`
/// allocation so a lower-precedence package detector can inspect it without a
/// full-input clone.
#[must_use = "inspect the prepared package or recover the original bytes"]
pub fn prepared_or_original(value: Vec<u8>) -> std::result::Result<PreparedPackage, Vec<u8>> {
    let Some(format) = packaged_mime(&value) else {
        return Err(value);
    };
    let package = OwnedPackage::from_prepared_bytes_or_recover(value)?;
    if !package.is_stored(MIMETYPE_PATH).unwrap_or(false) {
        return Err(package.into_inner());
    }
    Ok(PreparedPackage::new(package, format))
}

/// Compatibility spelling for callers that prefer the full detector name.
#[inline]
#[must_use]
pub fn prepared_package(value: Vec<u8>) -> Option<PreparedPackage> {
    prepared(value)
}

/// Detect a packaged or flat `OpenDocument` stream.
///
/// Detection reads the complete stream from its beginning and restores the
/// caller's original cursor position on every success or failure path. If the
/// original position cannot be restored, this function returns `None`.
pub fn reader<R: Read + Seek>(value: &mut R) -> Option<Format> {
    reader_with_limits(value, Limits::default()).ok().flatten()
}

/// Detect an ODF stream under explicit finite limits while restoring its
/// original cursor position on every success or failure path.
pub fn reader_with_limits<R: Read + Seek>(value: &mut R, limits: Limits) -> Result<Option<Format>> {
    limits.validate()?;
    let original = value.stream_position()?;
    let detected = (|| {
        value.seek(SeekFrom::Start(0))?;
        let data = read_owned_input(&mut *value, limits.max_input_bytes, "ODF detector input")?;
        bytes_with_limits(&data, limits)
    })();
    value.seek(SeekFrom::Start(original))?;
    detected
}

fn check_input_len(length: usize, limits: Limits, scope: &'static str) -> Result<()> {
    limits.validate()?;
    let observed = u64::try_from(length)
        .map_err(|_| Error::InvalidFormat("ODF detector input exceeds platform limits".into()))?;
    if observed > limits.max_input_bytes {
        return Err(Error::ResourceLimit(ResourceLimit {
            resource: Resource::InputBytes,
            observed,
            limit: limits.max_input_bytes,
            scope: scope.into(),
        }));
    }
    Ok(())
}

fn map_detector_zip_error(error: soapberry_zip::Error) -> Error {
    match crate::core::package::map_zip_error(error) {
        error @ (Error::Allocation { .. } | Error::Io(_) | Error::ResourceLimit(_)) => error,
        error => Error::InvalidFormat(error.to_string()),
    }
}

fn packaged_mime_bytes(value: &[u8]) -> Option<&[u8]> {
    let header = value.get(..LOCAL_HEADER_BYTES)?;
    let name = value.get(LOCAL_HEADER_BYTES..LOCAL_HEADER_BYTES + MIMETYPE_NAME.len())?;
    let (data_start, compressed, expected_crc) = packaged_mime_layout(header, name)?;
    let data_end = data_start.checked_add(compressed)?;
    let data = value.get(data_start..data_end)?;
    (soapberry_zip::crc32(data) == expected_crc).then_some(data)
}

/// Validate the strict local `mimetype` header grammar shared by complete-byte
/// and positional detection.  The caller supplies the already-read filename
/// bytes so no temporary concatenation is needed for a `ReadAt` probe.
fn packaged_mime_layout(header: &[u8], name: &[u8]) -> Option<(usize, usize, u32)> {
    if header.get(..4)? != ZIP_SIGNATURE {
        return None;
    }
    let flags = little_u16(header, 6)?;
    let compression = little_u16(header, 8)?;
    let expected_crc = little_u32(header, 14)?;
    let compressed = usize::try_from(little_u32(header, 18)?).ok()?;
    let uncompressed = usize::try_from(little_u32(header, 22)?).ok()?;
    let name_len = usize::from(little_u16(header, 26)?);
    let extra_len = usize::from(little_u16(header, 28)?);

    // ODF permits the UTF-8 file-name flag, but its mimetype entry must be
    // stored, sized in the local header, unencrypted, and have no extra field.
    if flags & !(1 << 11) != 0
        || compression != 0
        || compressed != uncompressed
        || uncompressed > MAX_MIMETYPE_BYTES
        || name_len != MIMETYPE_NAME.len()
        || extra_len != 0
        || name != MIMETYPE_NAME
    {
        return None;
    }
    let data_start = LOCAL_HEADER_BYTES
        .checked_add(name_len)?
        .checked_add(extra_len)?;
    Some((data_start, compressed, expected_crc))
}

fn ensure_source_current(expected: SourceVersion, observed: SourceVersion) -> Result<()> {
    if expected == observed {
        Ok(())
    } else {
        Err(Error::SourceChanged { expected, observed })
    }
}

fn little_u16(value: &[u8], offset: usize) -> Option<u16> {
    Some(u16::from_le_bytes(
        value.get(offset..offset + 2)?.try_into().ok()?,
    ))
}

fn little_u32(value: &[u8], offset: usize) -> Option<u32> {
    Some(u32::from_le_bytes(
        value.get(offset..offset + 4)?.try_into().ok()?,
    ))
}

fn trim_ascii(mut value: &[u8]) -> &[u8] {
    while value.first().is_some_and(u8::is_ascii_whitespace) {
        value = &value[1..];
    }
    while value.last().is_some_and(u8::is_ascii_whitespace) {
        value = &value[..value.len() - 1];
    }
    value
}

#[derive(Default)]
struct CatalogProbeState {
    file_count: usize,
    metadata_bytes: u64,
    total_uncompressed_bytes: u64,
    first_local_header_offset: Option<u64>,
    found_catalog: bool,
}

impl CatalogProbeState {
    fn finish(self) -> Option<bool> {
        (self.first_local_header_offset == Some(0)).then_some(self.found_catalog)
    }
}

fn catalog_probe_entry_with_limits(
    state: &mut CatalogProbeState,
    entry: &soapberry_zip::ZipFileHeaderRecord<'_>,
    limits: CatalogProbeLimits,
) -> bool {
    let Ok(name_bytes) = u64::try_from(entry.file_path().as_bytes().len()) else {
        return false;
    };
    if name_bytes > limits.max_member_name_bytes {
        return false;
    }
    let Some(metadata_bytes) = state.metadata_bytes.checked_add(entry.metadata_size_hint()) else {
        return false;
    };
    if metadata_bytes > limits.max_metadata_bytes {
        return false;
    }
    state.metadata_bytes = metadata_bytes;

    state.first_local_header_offset = Some(
        state
            .first_local_header_offset
            .map_or(entry.local_header_offset(), |offset| {
                offset.min(entry.local_header_offset())
            }),
    );

    // These layouts require the full bounded OPC path to validate before a
    // facade can make any precedence decision.
    if entry.is_zip64() || entry.has_data_descriptor() || entry.is_encrypted() {
        return false;
    }

    if entry.is_dir() {
        return true;
    }
    if state.file_count >= limits.max_members
        || entry.compressed_size_hint() > limits.max_compressed_bytes
        || entry.uncompressed_size_hint() > limits.max_entry_bytes
    {
        return false;
    }
    let Some(total_uncompressed_bytes) = state
        .total_uncompressed_bytes
        .checked_add(entry.uncompressed_size_hint())
    else {
        return false;
    };
    if total_uncompressed_bytes > limits.max_total_bytes {
        return false;
    }
    state.file_count += 1;
    state.total_uncompressed_bytes = total_uncompressed_bytes;
    true
}

fn catalog_probe_range_is_bounded(
    entry: &soapberry_zip::ZipFileHeaderRecord<'_>,
    directory_offset: u64,
) -> bool {
    let Some(end) = entry
        .local_header_offset()
        .checked_add(30)
        .and_then(|offset| offset.checked_add(entry.metadata_size_hint()))
        .and_then(|offset| offset.checked_add(entry.compressed_size_hint()))
    else {
        return false;
    };
    end <= directory_offset
}

fn map_catalog_probe_zip_error(error: soapberry_zip::Error) -> Result<Option<bool>> {
    match error.into_kind() {
        soapberry_zip::ErrorKind::IO(error) | soapberry_zip::ErrorKind::Io(error) => {
            map_catalog_probe_io_error(error)
        },
        soapberry_zip::ErrorKind::Allocation { resource, source } => {
            Err(Error::Allocation { resource, source })
        },
        _ => Ok(None),
    }
}

fn map_catalog_probe_io_error(error: std::io::Error) -> Result<Option<bool>> {
    if error
        .get_ref()
        .is_some_and(|source| source.is::<ReadAtProviderError>())
    {
        let source = error
            .into_inner()
            .expect("a provider marker has an inner error")
            .downcast::<ReadAtProviderError>()
            .expect("provider marker type was checked above");
        return Err(Error::Io(source.0));
    }
    if error.kind() == std::io::ErrorKind::UnexpectedEof {
        return Ok(None);
    }
    Err(Error::Io(error))
}

fn read_at_catalog_exact(
    source: &dyn ReadAt,
    mut offset: u64,
    mut output: &mut [u8],
) -> Result<bool> {
    while !output.is_empty() {
        match source.read_at(offset, output) {
            Ok(0) => return Ok(false),
            Ok(read) if read > output.len() => {
                return Err(Error::Io(std::io::Error::new(
                    std::io::ErrorKind::InvalidData,
                    "positional source reported more bytes than requested",
                )));
            },
            Ok(read) => {
                let read_u64 = u64::try_from(read).map_err(|_| {
                    Error::Io(std::io::Error::new(
                        std::io::ErrorKind::InvalidInput,
                        "source read length does not fit u64",
                    ))
                })?;
                offset = offset.checked_add(read_u64).ok_or_else(|| {
                    Error::Io(std::io::Error::new(
                        std::io::ErrorKind::InvalidInput,
                        "source offset overflow",
                    ))
                })?;
                output = &mut output[read..];
            },
            Err(error) if error.kind() == std::io::ErrorKind::Interrupted => {},
            Err(error) => return Err(Error::Io(error)),
        }
    }
    Ok(true)
}

#[derive(Debug)]
struct ReadAtProviderError(std::io::Error);

impl std::fmt::Display for ReadAtProviderError {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        self.0.fmt(formatter)
    }
}

impl std::error::Error for ReadAtProviderError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        Some(&self.0)
    }
}

struct ReadAtZipSource<'a> {
    source: &'a dyn ReadAt,
}

impl soapberry_zip::ReaderAt for ReadAtZipSource<'_> {
    fn read_at(&self, buffer: &mut [u8], offset: u64) -> std::io::Result<usize> {
        self.source
            .read_at(offset, buffer)
            .map_err(|error| std::io::Error::new(error.kind(), ReadAtProviderError(error)))
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::{Cursor, Write};

    fn zip_with_entries(entries: &[(&str, &[u8])]) -> Vec<u8> {
        let mut output = Cursor::new(Vec::new());
        let mut writer = zip::ZipWriter::new(&mut output);
        let options = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Stored);
        for (path, value) in entries {
            writer
                .start_file(*path, options)
                .unwrap_or_else(|error| panic!("test ZIP entry must start: {error}"));
            writer
                .write_all(value)
                .unwrap_or_else(|error| panic!("test ZIP entry must write: {error}"));
        }
        writer
            .finish()
            .unwrap_or_else(|error| panic!("test ZIP must finish: {error}"));
        output.into_inner()
    }

    fn central_record(bytes: &[u8], target: &[u8]) -> usize {
        let mut offset = bytes
            .windows(4)
            .position(|signature| signature == b"PK\x01\x02")
            .expect("test ZIP must have a central directory");
        loop {
            let name_len = usize::from(u16::from_le_bytes(
                bytes[offset + 28..offset + 30]
                    .try_into()
                    .expect("central name length"),
            ));
            let extra_len = usize::from(u16::from_le_bytes(
                bytes[offset + 30..offset + 32]
                    .try_into()
                    .expect("central extra length"),
            ));
            let comment_len = usize::from(u16::from_le_bytes(
                bytes[offset + 32..offset + 34]
                    .try_into()
                    .expect("central comment length"),
            ));
            if &bytes[offset + 46..offset + 46 + name_len] == target {
                return offset;
            }
            offset += 46 + name_len + extra_len + comment_len;
            assert_eq!(&bytes[offset..offset + 4], b"PK\x01\x02");
        }
    }

    fn local_record(bytes: &[u8], target: &[u8]) -> usize {
        let mut offset = 0;
        while let Some(relative) = bytes[offset..]
            .windows(4)
            .position(|signature| signature == b"PK\x03\x04")
        {
            offset += relative;
            let name_len = usize::from(u16::from_le_bytes(
                bytes[offset + 26..offset + 28]
                    .try_into()
                    .expect("local name length"),
            ));
            if &bytes[offset + 30..offset + 30 + name_len] == target {
                return offset;
            }
            let extra_len = usize::from(u16::from_le_bytes(
                bytes[offset + 28..offset + 30]
                    .try_into()
                    .expect("local extra length"),
            ));
            let compressed_len = usize::try_from(u32::from_le_bytes(
                bytes[offset + 18..offset + 22]
                    .try_into()
                    .expect("local compressed length"),
            ))
            .expect("test ZIP size");
            offset += 30 + name_len + extra_len + compressed_len;
        }
        panic!("test ZIP local record not found");
    }

    #[test]
    fn classifies_every_packaged_family_and_template_without_lossy_text() {
        for (value, expected) in [
            (constants::ODF_TEXT, Format::Odt),
            (constants::ODF_TEXT_TEMPLATE, Format::Odt),
            (constants::ODF_SPREADSHEET, Format::Ods),
            (constants::ODF_SPREADSHEET_TEMPLATE, Format::Ods),
            (constants::ODF_PRESENTATION, Format::Odp),
            (constants::ODF_PRESENTATION_TEMPLATE, Format::Odp),
            (constants::ODF_DRAWING, Format::Odg),
            (constants::ODF_DRAWING_TEMPLATE, Format::Odg),
            (constants::ODF_CHART, Format::Odc),
            (constants::ODF_CHART_TEMPLATE, Format::Odc),
            (constants::ODF_FORMULA, Format::Odf),
            (constants::ODF_FORMULA_TEMPLATE, Format::Odf),
            (constants::ODF_IMAGE, Format::Odi),
            (constants::ODF_IMAGE_TEMPLATE, Format::Odi),
            (constants::ODF_MASTER, Format::Odm),
            (constants::ODF_MASTER_TEMPLATE, Format::Odm),
            (constants::ODF_WEB, Format::Oth),
            (constants::ODF_DATABASE, Format::Odb),
        ] {
            assert_eq!(mime(value.as_bytes()), Some(expected), "{value}");
        }
        assert_eq!(
            mime(b" \napplication/vnd.oasis.opendocument.text\t"),
            Some(Format::Odt)
        );
        assert_eq!(mime(b"application/pdf"), None);
        assert_eq!(mime(b"\xff"), None);
    }

    #[test]
    fn detects_flat_documents_with_semantic_namespace_resolution() {
        for (body, mimetype, expected) in [
            ("text", constants::ODF_TEXT, Format::Odt),
            ("spreadsheet", constants::ODF_SPREADSHEET, Format::Ods),
            ("presentation", constants::ODF_PRESENTATION, Format::Odp),
            ("drawing", constants::ODF_DRAWING, Format::Odg),
            ("chart", constants::ODF_CHART, Format::Odc),
            ("formula", constants::ODF_FORMULA, Format::Odf),
            ("image", constants::ODF_IMAGE, Format::Odi),
        ] {
            let xml = format!(
                r#"<?xml version="1.0"?><!--flat--><o:document xmlns:o="{}" o:mimetype="{mimetype}"><o:body><o:{body}/></o:body></o:document>"#,
                String::from_utf8_lossy(OFFICE_NAMESPACE),
            );
            assert_eq!(flat_mime(xml.as_bytes()).as_deref(), Some(mimetype));
            assert_eq!(flat(xml.as_bytes()), Some(expected));
            assert_eq!(bytes(xml.as_bytes()), Some(expected));
        }

        let padded = format!(
            r#"<o:document xmlns:o="{}" o:mimetype="  {}  "><o:body><o:text/></o:body></o:document>"#,
            String::from_utf8_lossy(OFFICE_NAMESPACE),
            constants::ODF_TEXT,
        );
        assert_eq!(
            flat_mime(padded.as_bytes()).as_deref(),
            Some(constants::ODF_TEXT)
        );
    }

    #[test]
    fn rejects_non_flat_roots_and_unknown_mimetypes() {
        for xml in [
            r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" office:mimetype="application/vnd.oasis.opendocument.text"/>"#,
            r#"<office:document xmlns:office="urn:wrong" office:mimetype="application/vnd.oasis.opendocument.text"/>"#,
            r#"<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" office:mimetype="application/xml"/>"#,
            r#"<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"/>"#,
        ] {
            assert_eq!(flat(xml.as_bytes()), None);
        }
    }

    #[test]
    fn detects_packaged_documents_and_restores_nonzero_reader_position() {
        let mut writer = crate::core::PackageWriter::new();
        writer
            .set_mimetype(constants::ODF_TEXT)
            .unwrap_or_else(|error| panic!("test package mimetype must be accepted: {error}"));
        let package = writer
            .finish_to_bytes()
            .unwrap_or_else(|error| panic!("test package must be writable: {error}"));
        assert_eq!(bytes(&package), Some(Format::Odt));

        let mut input = Cursor::new(package);
        input.set_position(7);
        assert_eq!(reader(&mut input), Some(Format::Odt));
        assert_eq!(input.position(), 7);

        let xml = format!(
            r#"<o:document xmlns:o="{}" o:mimetype="{}"><o:body><o:text/></o:body></o:document>"#,
            String::from_utf8_lossy(OFFICE_NAMESPACE),
            constants::ODF_TEXT,
        );
        let mut flat_input = Cursor::new(xml);
        flat_input.set_position(9);
        assert_eq!(reader(&mut flat_input), Some(Format::Odt));
        assert_eq!(flat_input.position(), 9);

        let mut invalid = Cursor::new(b"not an OpenDocument file".to_vec());
        invalid.set_position(4);
        assert_eq!(reader(&mut invalid), None);
        assert_eq!(invalid.position(), 4);
    }

    #[test]
    fn bounded_reader_uses_max_plus_one_and_restores_position() {
        let mut reader = Cursor::new(b"xx".to_vec());
        reader.set_position(1);
        let error = reader_with_limits(&mut reader, Limits::new(1)).unwrap_err();
        assert!(matches!(error, Error::ResourceLimit(_)));
        assert_eq!(reader.position(), 1);
        assert!(matches!(
            reader_with_limits(&mut reader, Limits::new(0)),
            Err(Error::InvalidFormat(_))
        ));
    }

    #[test]
    fn detector_preserves_zip_file_count_and_metadata_limits() {
        for (resource, expected_resource) in [
            (soapberry_zip::LimitResource::FileCount, Resource::Objects),
            (
                soapberry_zip::LimitResource::MetadataBytes,
                Resource::InputBytes,
            ),
        ] {
            let error = map_detector_zip_error(soapberry_zip::Error::from(
                soapberry_zip::ErrorKind::LimitExceeded {
                    resource,
                    actual: 2,
                    maximum: 1,
                },
            ));
            assert!(matches!(
                error,
                Error::ResourceLimit(ResourceLimit {
                    resource,
                    observed: 2,
                    limit: 1,
                    ..
                }) if resource == expected_resource
            ));
        }
    }

    #[test]
    fn packaged_detection_rejects_nonconforming_local_mimetype_entries() {
        let mut writer = crate::core::PackageWriter::new();
        writer
            .set_mimetype(constants::ODF_TEXT)
            .unwrap_or_else(|error| panic!("test package mimetype must be accepted: {error}"));
        let package = writer
            .finish_to_bytes()
            .unwrap_or_else(|error| panic!("test package must be writable: {error}"));

        let mut compressed = package.clone();
        compressed[8..10].copy_from_slice(&8_u16.to_le_bytes());
        assert_eq!(bytes(&compressed), None);

        let mut corrupt = package.clone();
        let payload = LOCAL_HEADER_BYTES + MIMETYPE_NAME.len();
        corrupt[payload] ^= 1;
        assert_eq!(bytes(&corrupt), None);

        assert_eq!(bytes(&package[..LOCAL_HEADER_BYTES]), None);
        let local_entry_end = LOCAL_HEADER_BYTES + MIMETYPE_NAME.len() + constants::ODF_TEXT.len();
        assert_eq!(bytes(&package[..local_entry_end]), None);
    }

    #[test]
    fn central_catalog_probe_is_exact_and_conservative() {
        let ordinary = zip_with_entries(&[
            ("mimetype", constants::ODF_TEXT.as_bytes()),
            ("content.xml", b"[Content_Types].xml"),
        ]);
        assert_eq!(packaged_has_ooxml_catalog(&ordinary), Some(false));

        let case_insensitive = zip_with_entries(&[
            ("mimetype", constants::ODF_TEXT.as_bytes()),
            ("[content_types].xml", b"<broken>"),
        ]);
        assert_eq!(packaged_has_ooxml_catalog(&case_insensitive), Some(true));

        let normalized = zip_with_entries(&[
            ("mimetype", constants::ODF_TEXT.as_bytes()),
            ("/[Content_Types].xml", b"<broken>"),
        ]);
        assert_eq!(packaged_has_ooxml_catalog(&normalized), None);

        let directory = zip_with_entries(&[
            ("mimetype", constants::ODF_TEXT.as_bytes()),
            ("[Content_Types].xml/", b""),
        ]);
        assert_eq!(packaged_has_ooxml_catalog(&directory), Some(false));

        let mut malformed = ordinary;
        malformed.truncate(malformed.len() - 1);
        assert_eq!(packaged_has_ooxml_catalog(&malformed), None);
        assert_eq!(packaged_has_ooxml_catalog(b"not a ZIP"), None);
    }

    #[test]
    fn central_catalog_probe_applies_caller_budgets() {
        let ordinary = zip_with_entries(&[
            ("mimetype", constants::ODF_TEXT.as_bytes()),
            ("content.xml", b"<content/>"),
        ]);
        let limits = CatalogProbeLimits::new(
            u64::MAX,
            1,
            u64::MAX,
            u64::MAX,
            u64::MAX,
            u64::MAX,
            u64::MAX,
        );
        assert_eq!(
            packaged_has_ooxml_catalog_with_limits(&ordinary, limits),
            None
        );

        let input_limited = CatalogProbeLimits::new(
            u64::try_from(ordinary.len() - 1).expect("test ZIP length"),
            usize::MAX,
            u64::MAX,
            u64::MAX,
            u64::MAX,
            u64::MAX,
            u64::MAX,
        );
        assert_eq!(
            packaged_has_ooxml_catalog_with_limits(&ordinary, input_limited),
            None
        );
    }

    #[derive(Debug)]
    struct ProviderFailure;

    impl ReadAt for ProviderFailure {
        fn len(&self) -> std::io::Result<u64> {
            Ok(128)
        }

        fn read_at(&self, _offset: u64, _output: &mut [u8]) -> std::io::Result<usize> {
            Err(std::io::Error::new(
                std::io::ErrorKind::PermissionDenied,
                "catalog provider failed",
            ))
        }

        fn version(&self) -> std::io::Result<SourceVersion> {
            Ok(SourceVersion::new(17, 0))
        }
    }

    struct CentralProviderFailure {
        signature: [u8; ZIP_SIGNATURE.len()],
        first_read: std::sync::atomic::AtomicBool,
    }

    impl ReadAt for CentralProviderFailure {
        fn len(&self) -> std::io::Result<u64> {
            Ok(128)
        }

        fn read_at(&self, offset: u64, output: &mut [u8]) -> std::io::Result<usize> {
            if offset == 0
                && self
                    .first_read
                    .swap(false, std::sync::atomic::Ordering::Relaxed)
            {
                let count = output.len().min(self.signature.len());
                output[..count].copy_from_slice(&self.signature[..count]);
                return Ok(count);
            }
            Err(std::io::Error::new(
                std::io::ErrorKind::ConnectionReset,
                "central catalog provider failed",
            ))
        }

        fn version(&self) -> std::io::Result<SourceVersion> {
            Ok(SourceVersion::new(20, 0))
        }
    }

    struct ShortReadAt {
        length: u64,
    }

    impl ReadAt for ShortReadAt {
        fn len(&self) -> std::io::Result<u64> {
            Ok(self.length)
        }

        fn read_at(&self, _offset: u64, _output: &mut [u8]) -> std::io::Result<usize> {
            Ok(0)
        }

        fn version(&self) -> std::io::Result<SourceVersion> {
            Ok(SourceVersion::new(19, 0))
        }
    }

    struct ChangingSource {
        bytes: Vec<u8>,
        revision: std::sync::atomic::AtomicU64,
    }

    impl ReadAt for ChangingSource {
        fn len(&self) -> std::io::Result<u64> {
            u64::try_from(self.bytes.len()).map_err(|_| std::io::Error::other("test source length"))
        }

        fn read_at(&self, offset: u64, output: &mut [u8]) -> std::io::Result<usize> {
            let offset = usize::try_from(offset).map_err(|_| {
                std::io::Error::new(std::io::ErrorKind::InvalidInput, "test offset")
            })?;
            let Some(input) = self.bytes.get(offset..) else {
                return Ok(0);
            };
            let count = input.len().min(output.len());
            output[..count].copy_from_slice(&input[..count]);
            Ok(count)
        }

        fn version(&self) -> std::io::Result<SourceVersion> {
            Ok(SourceVersion::new(
                18,
                self.revision
                    .fetch_add(1, std::sync::atomic::Ordering::Relaxed),
            ))
        }
    }

    #[test]
    fn positional_catalog_probe_preserves_provider_errors_and_source_changes() {
        assert_eq!(
            packaged_has_ooxml_catalog_read_at(&ShortReadAt { length: 128 })
                .expect("short source is a layout uncertainty"),
            None
        );
        assert!(matches!(
            packaged_has_ooxml_catalog_read_at(&ProviderFailure),
            Err(Error::Io(error)) if error.kind() == std::io::ErrorKind::PermissionDenied
        ));
        assert!(matches!(
            packaged_has_ooxml_catalog_read_at(&CentralProviderFailure {
                signature: *b"PK\x03\x04",
                first_read: std::sync::atomic::AtomicBool::new(true),
            }),
            Err(Error::Io(error)) if error.kind() == std::io::ErrorKind::ConnectionReset
        ));

        let ordinary = zip_with_entries(&[
            ("mimetype", constants::ODF_TEXT.as_bytes()),
            ("content.xml", b"<content/>"),
        ]);
        let source = ChangingSource {
            bytes: ordinary,
            revision: std::sync::atomic::AtomicU64::new(0),
        };
        assert!(matches!(
            packaged_has_ooxml_catalog_read_at(&source),
            Err(Error::SourceChanged { .. })
        ));
    }

    struct ShortReadReader {
        cursor: Cursor<Vec<u8>>,
    }

    impl Read for ShortReadReader {
        fn read(&mut self, output: &mut [u8]) -> std::io::Result<usize> {
            if output.is_empty() {
                return Ok(0);
            }
            let mut one = [0_u8; 1];
            let read = self.cursor.read(&mut one)?;
            if read == 1 {
                output[0] = one[0];
            }
            Ok(read)
        }
    }

    impl Seek for ShortReadReader {
        fn seek(&mut self, position: SeekFrom) -> std::io::Result<u64> {
            self.cursor.seek(position)
        }
    }

    struct RestoreFailReader {
        cursor: Cursor<Vec<u8>>,
        refused_position: u64,
    }

    impl Read for RestoreFailReader {
        fn read(&mut self, output: &mut [u8]) -> std::io::Result<usize> {
            self.cursor.read(output)
        }
    }

    impl Seek for RestoreFailReader {
        fn seek(&mut self, position: SeekFrom) -> std::io::Result<u64> {
            if let SeekFrom::Start(offset) = position
                && offset == self.refused_position
            {
                return Err(std::io::Error::other("cursor restore refused"));
            }
            self.cursor.seek(position)
        }
    }

    #[test]
    fn reader_catalog_probe_handles_short_reads_and_restore_failures() {
        let ordinary = zip_with_entries(&[
            ("mimetype", constants::ODF_TEXT.as_bytes()),
            ("content.xml", b"<content/>"),
        ]);
        let mut short = ShortReadReader {
            cursor: Cursor::new(ordinary.clone()),
        };
        short.seek(SeekFrom::Start(7)).expect("test cursor");
        assert_eq!(
            packaged_has_ooxml_catalog_from_reader(&mut short),
            Some(false)
        );
        assert_eq!(short.cursor.position(), 7);

        let mut failing = RestoreFailReader {
            cursor: Cursor::new(ordinary),
            refused_position: 11,
        };
        failing.cursor.set_position(failing.refused_position);
        assert_eq!(packaged_has_ooxml_catalog_from_reader(&mut failing), None);
    }

    #[test]
    fn positional_packaged_probe_reuses_strict_local_mimetype_grammar() {
        let mut writer = crate::core::PackageWriter::new();
        writer
            .set_mimetype(constants::ODF_PRESENTATION)
            .expect("test package MIME");
        writer
            .add_file("content.xml", b"<office:document-content/>")
            .expect("test content");
        let package = writer.finish_to_bytes().expect("test package");
        let source = litchi_core::OwnedSource::new(package.clone());

        assert_eq!(
            packaged_mime_read_at(&source).expect("positional probe"),
            packaged_mime(&package)
        );
        assert_eq!(
            packaged_has_ooxml_catalog_read_at(&source).expect("positional catalog probe"),
            packaged_has_ooxml_catalog(&package)
        );
        let mut seekable = Cursor::new(package.clone());
        seekable.set_position(5);
        assert_eq!(
            packaged_has_ooxml_catalog_from_reader(&mut seekable),
            Some(false)
        );
        assert_eq!(seekable.position(), 5);
        let mut malformed = package.clone();
        malformed.truncate(malformed.len() - 1);
        assert_eq!(
            packaged_has_ooxml_catalog_read_at(&litchi_core::OwnedSource::new(malformed))
                .expect("malformed positional catalog probe"),
            None
        );
        assert_eq!(
            packaged_mime_read_at(&litchi_core::OwnedSource::new(
                package[..LOCAL_HEADER_BYTES].to_vec()
            ))
            .expect("short positional probe"),
            None
        );
    }

    #[test]
    fn prepared_detection_retains_one_bounded_index_and_rejects_hostile_archives() {
        crate::package::reset_index_build_count();
        let mut writer = crate::core::PackageWriter::new();
        writer
            .set_mimetype(constants::ODF_TEXT)
            .unwrap_or_else(|error| panic!("test package mimetype must be accepted: {error}"));
        let package = writer
            .finish_to_bytes()
            .unwrap_or_else(|error| panic!("test package must be writable: {error}"));
        let retained = prepared(package).expect("valid package must prepare");
        assert_eq!(retained.format(), Format::Odt);
        assert_ne!(retained.prepared_index_identity(), 0);
        assert_eq!(crate::package::index_build_count(), 1);
        let _semantic_package = retained
            .package()
            .package()
            .expect("prepared package must expose its indexed view");
        assert_eq!(crate::package::index_build_count(), 1);

        let duplicate = zip_with_entries(&[
            ("mimetype", constants::ODF_TEXT.as_bytes()),
            ("Pictures/a", b"one"),
            ("./Pictures/a", b"two"),
        ]);
        assert!(prepared(duplicate).is_none());

        let traversal = zip_with_entries(&[
            ("mimetype", constants::ODF_TEXT.as_bytes()),
            ("../Pictures/a", b"one"),
            ("Pictures/a", b"two"),
        ]);
        assert!(prepared(traversal).is_none());

        let oversized_name = format!("{}.xml", "x".repeat(4 * 1024));
        let oversized = zip_with_entries(&[
            ("mimetype", constants::ODF_TEXT.as_bytes()),
            (&oversized_name, b"one"),
        ]);
        assert!(prepared(oversized).is_none());
        assert!(prepared(b"PK\x03\x04truncated".to_vec()).is_none());

        let mut forged = zip_with_entries(&[
            ("mimetype", constants::ODF_TEXT.as_bytes()),
            ("content.xml", b"<office:document-content/>"),
        ]);
        let content_local_offset = local_record(&forged, b"content.xml");
        let mimetype_central_offset = central_record(&forged, b"mimetype");
        forged[mimetype_central_offset + 42..mimetype_central_offset + 46]
            .copy_from_slice(&(content_local_offset as u32).to_le_bytes());
        assert!(prepared(forged).is_none());

        let mut invalid_utf8 = zip_with_entries(&[
            ("mimetype", constants::ODF_TEXT.as_bytes()),
            ("content.xml", b"<office:document-content/>"),
        ]);
        let invalid_local = local_record(&invalid_utf8, b"content.xml");
        invalid_utf8[invalid_local + 30] = 0xff;
        let invalid_central = central_record(&invalid_utf8, b"content.xml");
        invalid_utf8[invalid_central + 46] = 0xff;
        assert!(prepared(invalid_utf8).is_none());

        let traversal = zip_with_entries(&[
            ("mimetype", constants::ODF_TEXT.as_bytes()),
            ("../content.xml", b"junk"),
        ]);
        assert!(prepared(traversal).is_none());
    }

    #[test]
    fn rejected_prepared_detection_returns_the_original_allocation() {
        let mut writer = crate::core::PackageWriter::new();
        writer
            .set_mimetype(constants::ODF_TEXT)
            .expect("test package mimetype must be accepted");
        let mut invalid = writer
            .finish_to_bytes()
            .expect("test package must be writable");
        let central = central_record(&invalid, MIMETYPE_NAME);
        invalid[central] = 0;
        invalid.reserve(64);
        let pointer = invalid.as_ptr();
        let capacity = invalid.capacity();

        let recovered = match prepared_or_original(invalid) {
            Err(recovered) => recovered,
            Ok(_) => panic!("malformed ODF index must return its source"),
        };

        assert_eq!(recovered.as_ptr(), pointer);
        assert_eq!(recovered.capacity(), capacity);
    }
}
