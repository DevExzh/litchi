//! Catalog-first, source-backed ODS reads.
//!
//! [`SourceBackedSpreadsheetCatalog`] is an opt-in lifecycle for callers that
//! need worksheet discovery and one worksheet at a time. Opening validates
//! the package MIME/manifest and the ODS content hierarchy, retains only
//! bounded worksheet names, and defers typed worksheet allocation until a
//! worksheet selector is used. Unrelated package members remain opaque and
//! cold until explicitly selected.

use std::{
    collections::HashSet,
    fmt,
    ops::Range,
    sync::{
        Arc,
        atomic::{AtomicU64, Ordering},
    },
};

#[cfg(any(unix, windows))]
use std::path::Path;

#[cfg(any(unix, windows))]
use litchi_core::FileSource;
use litchi_core::{Error, ReadAt, Result, SourceVersion};
use litchi_odf_common::{
    core::SourceBackedPackage,
    package::{is_media_path, resolve_package_path},
};
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::{Namespace, ResolveResult},
    reader::NsReader,
};
use zeroize::Zeroizing;

use super::ReadLimits;
use crate::{
    Spreadsheet,
    authoring::{ValidateHandler, is_office_namespace, validate_size},
    worksheet::{Sheet, codec::OFFICE_NAMESPACE, codec::TABLE_NAMESPACE},
};

const ODF_SPREADSHEET: &str = "application/vnd.oasis.opendocument.spreadsheet";
const FAMILY_NAME: &str = "ODS";

/// A bounded snapshot of positional reads performed by one catalog owner.
///
/// The counters include ZIP indexing, mandatory package reads, and deferred
/// member reads. Version checks are deliberately not counted as payload
/// reads. Values are observational when calls overlap concurrently.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct SourceReadMetrics {
    /// Number of positional `read_at` calls made by this owner.
    pub read_calls: u64,
    /// Number of bytes returned successfully by positional reads.
    pub read_bytes: u64,
}

impl SourceReadMetrics {
    /// Return the number of positional reads.
    #[must_use]
    pub const fn calls(self) -> u64 {
        self.read_calls
    }

    /// Return the number of bytes returned by positional reads.
    #[must_use]
    pub const fn bytes(self) -> u64 {
        self.read_bytes
    }
}

#[derive(Debug, Default)]
struct ReadMeter {
    calls: AtomicU64,
    bytes: AtomicU64,
}

impl ReadMeter {
    fn snapshot(&self) -> SourceReadMetrics {
        SourceReadMetrics {
            read_calls: self.calls.load(Ordering::Relaxed),
            read_bytes: self.bytes.load(Ordering::Relaxed),
        }
    }

    fn reset(&self) {
        self.calls.store(0, Ordering::Relaxed);
        self.bytes.store(0, Ordering::Relaxed);
    }

    fn add_call(&self) {
        let _unused = self
            .calls
            .fetch_update(Ordering::Relaxed, Ordering::Relaxed, |value| {
                Some(value.saturating_add(1))
            });
    }

    fn add_bytes(&self, bytes: usize) {
        let Ok(bytes) = u64::try_from(bytes) else {
            return;
        };
        let _unused = self
            .bytes
            .fetch_update(Ordering::Relaxed, Ordering::Relaxed, |value| {
                Some(value.saturating_add(bytes))
            });
    }
}

struct MeteredSource {
    source: Arc<dyn ReadAt>,
    meter: Arc<ReadMeter>,
}

impl ReadAt for MeteredSource {
    fn len(&self) -> std::io::Result<u64> {
        self.source.len()
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> std::io::Result<usize> {
        self.meter.add_call();
        let read = self.source.read_at(offset, output)?;
        if read <= output.len() {
            self.meter.add_bytes(read);
        }
        Ok(read)
    }

    fn version(&self) -> std::io::Result<SourceVersion> {
        self.source.version()
    }
}

/// One worksheet entry in a catalog-first ODS owner.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SheetCatalogEntry {
    index: usize,
    name: String,
}

impl SheetCatalogEntry {
    /// Return the zero-based worksheet position.
    #[must_use]
    pub const fn index(&self) -> usize {
        self.index
    }

    /// Return the exact ODF worksheet name.
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }
}

/// Immutable, catalog-first ODS access over a positional source.
///
/// Opening reads the ZIP index, `mimetype`, the manifest, and `content.xml`
/// once to validate the ODS hierarchy and retain worksheet names. It does
/// not parse or retain any [`Sheet`], styles, metadata, media, embedded
/// object, or signature payload. [`Self::sheet_at`] and [`Self::sheet`] are
/// the explicit semantic-read boundary. [`Self::materialize`] is the explicit
/// transition to the complete owned/mutable [`Spreadsheet`] facade.
pub struct SourceBackedSpreadsheetCatalog {
    package: SourceBackedPackage,
    source: Arc<dyn ReadAt>,
    source_version: SourceVersion,
    entries: Arc<[SheetCatalogEntry]>,
    meter: Arc<ReadMeter>,
}

impl fmt::Debug for SourceBackedSpreadsheetCatalog {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SourceBackedSpreadsheetCatalog")
            .field("source_version", &self.source_version)
            .field("sheets", &self.entries.len())
            .finish_non_exhaustive()
    }
}

impl SourceBackedSpreadsheetCatalog {
    /// Open an ODS catalog from a filesystem path without slurping the source.
    #[cfg(any(unix, windows))]
    pub fn from_path(path: impl AsRef<Path>) -> Result<Self> {
        Self::from_read_at(file_source(path)?)
    }

    /// Open an ODS catalog from a filesystem path with explicit limits.
    #[cfg(any(unix, windows))]
    pub fn from_path_with_limits(path: impl AsRef<Path>, limits: ReadLimits) -> Result<Self> {
        Self::from_read_at_with_limits(file_source(path)?, limits)
    }

    /// Open an encrypted ODS catalog from a filesystem path.
    #[cfg(any(unix, windows))]
    pub fn from_path_with_password(
        path: impl AsRef<Path>,
        password: impl Into<String>,
    ) -> Result<Self> {
        let mut password = Zeroizing::new(password.into());
        Self::from_read_at_with_password(file_source(path)?, std::mem::take(&mut *password))
    }

    /// Open an encrypted ODS catalog from a filesystem path with limits.
    #[cfg(any(unix, windows))]
    pub fn from_path_with_limits_and_password(
        path: impl AsRef<Path>,
        limits: ReadLimits,
        password: impl Into<String>,
    ) -> Result<Self> {
        let mut password = Zeroizing::new(password.into());
        Self::from_read_at_with_limits_and_password(
            file_source(path)?,
            limits,
            std::mem::take(&mut *password),
        )
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

    /// Open an ODS catalog from a caller-provided positional source.
    pub fn from_read_at(source: Arc<dyn ReadAt>) -> Result<Self> {
        Self::from_read_at_with_limits(source, ReadLimits::default())
    }

    /// Open an encrypted ODS catalog from a caller-provided source.
    pub fn from_read_at_with_password(
        source: Arc<dyn ReadAt>,
        password: impl Into<String>,
    ) -> Result<Self> {
        Self::from_read_at_inner(source, ReadLimits::default(), Some(password.into()))
    }

    /// Open an ODS catalog with explicit finite ZIP limits.
    pub fn from_read_at_with_limits(source: Arc<dyn ReadAt>, limits: ReadLimits) -> Result<Self> {
        Self::from_read_at_inner(source, limits, None)
    }

    /// Open an encrypted ODS catalog with explicit finite ZIP limits.
    pub fn from_read_at_with_limits_and_password(
        source: Arc<dyn ReadAt>,
        limits: ReadLimits,
        password: impl Into<String>,
    ) -> Result<Self> {
        Self::from_read_at_inner(source, limits, Some(password.into()))
    }

    fn from_read_at_inner(
        source: Arc<dyn ReadAt>,
        limits: ReadLimits,
        password: Option<String>,
    ) -> Result<Self> {
        let meter = Arc::new(ReadMeter::default());
        let metered: Arc<dyn ReadAt> = Arc::new(MeteredSource {
            source,
            meter: Arc::clone(&meter),
        });
        let package = match password {
            Some(password) => {
                let mut password = Zeroizing::new(password);
                SourceBackedPackage::from_read_at_with_limits_and_password(
                    Arc::clone(&metered),
                    limits,
                    std::mem::take(&mut *password),
                )?
            },
            None => SourceBackedPackage::from_read_at_with_limits(Arc::clone(&metered), limits)?,
        };
        let source_version = package.source_version()?;
        let entries = (|| {
            let mimetype = package.mimetype()?;
            if mimetype != ODF_SPREADSHEET {
                return Err(Error::InvalidFormat(format!(
                    "expected {FAMILY_NAME} package MIME type '{ODF_SPREADSHEET}', found '{mimetype}'"
                )));
            }
            let content = read_content(&package)?;
            scan_catalog(&content)
        })();
        let entries = prefer_current(metered.as_ref(), source_version, entries)?;
        Ok(Self {
            package,
            source: metered,
            source_version,
            entries: Arc::from(entries),
            meter,
        })
    }

    /// Check the positional source identity without reading a member.
    pub fn check_source(&self) -> Result<()> {
        ensure_current(self.source_version, self.source.version()?)
    }

    /// Return the source version captured at open.
    pub fn source_version(&self) -> Result<SourceVersion> {
        self.check_source()?;
        Ok(self.source_version)
    }

    /// Return the physical source length captured at open.
    pub fn source_len(&self) -> Result<u64> {
        self.check_source()?;
        let length = self.package.len();
        prefer_current(self.source.as_ref(), self.source_version, Ok(length))
    }

    /// Return a snapshot of positional source-read counters.
    pub fn source_read_metrics(&self) -> Result<SourceReadMetrics> {
        self.check_source()?;
        let metrics = self.meter.snapshot();
        prefer_current(self.source.as_ref(), self.source_version, Ok(metrics))
    }

    /// Reset positional source-read counters after a preparation phase.
    ///
    /// Resetting counters is an observational instrumentation operation and
    /// is intended for a quiescent owner or an explicitly coordinated query.
    pub fn reset_source_read_metrics(&self) -> Result<()> {
        self.check_source()?;
        self.meter.reset();
        self.check_source()
    }

    /// Borrow the retained worksheet catalog.
    pub fn catalog(&self) -> Result<&[SheetCatalogEntry]> {
        self.check_source()?;
        let entries = self.entries.as_ref();
        prefer_current(self.source.as_ref(), self.source_version, Ok(entries))
    }

    /// Return worksheet names in document order.
    pub fn sheet_names(&self) -> Result<Vec<String>> {
        self.check_source()?;
        let mut names = Vec::new();
        names
            .try_reserve_exact(self.entries.len())
            .map_err(|source| Error::Allocation {
                resource: "ODS source catalog worksheet names",
                source,
            })?;
        names.extend(self.entries.iter().map(|entry| entry.name.clone()));
        prefer_current(self.source.as_ref(), self.source_version, Ok(names))
    }

    /// Return the number of worksheets without parsing worksheet payloads.
    pub fn sheet_count(&self) -> Result<usize> {
        self.check_source()?;
        let count = self.entries.len();
        prefer_current(self.source.as_ref(), self.source_version, Ok(count))
    }

    /// Read one worksheet selected by its zero-based document position.
    pub fn sheet_at(&self, index: usize) -> Result<Option<Sheet>> {
        self.check_source()?;
        if index >= self.entries.len() {
            return Ok(None);
        }
        let content = read_content(&self.package)?;
        let result = parse_selected_sheet(&content, index);
        prefer_current(self.source.as_ref(), self.source_version, result)
    }

    /// Read one worksheet selected by its exact ODF name.
    pub fn sheet(&self, name: &str) -> Result<Option<Sheet>> {
        self.check_source()?;
        let Some(entry) = self.entries.iter().find(|entry| entry.name == name) else {
            return Ok(None);
        };
        self.sheet_at(entry.index)
    }

    /// List package members without reading their payloads.
    pub fn files(&self) -> Result<Vec<String>> {
        let result = self.package.files();
        prefer_current(self.source.as_ref(), self.source_version, result)
    }

    /// List media members without reading their payloads.
    pub fn media_files(&self) -> Result<Vec<String>> {
        let result = self.package.media_files();
        prefer_current(self.source.as_ref(), self.source_version, result)
    }

    /// Read one safe package member on demand.
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

    /// Read one safe package media member on demand.
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

    /// Read inert document and macro signature metadata without executing it.
    pub fn digital_signatures(&self) -> Result<litchi_odf_common::signature::DigitalSignatures> {
        let result = self.package.digital_signatures();
        prefer_current(self.source.as_ref(), self.source_version, result)
    }

    /// Materialize the exact source into the complete owned/mutable ODS owner.
    pub fn materialize(self) -> Result<Spreadsheet> {
        self.check_source()?;
        let source = Arc::clone(&self.source);
        let version = self.source_version;
        let package = prefer_current(source.as_ref(), version, self.package.materialize())?;
        prefer_current(
            source.as_ref(),
            version,
            Spreadsheet::from_owned_package(package),
        )
    }
}

#[cfg(any(unix, windows))]
fn file_source(path: impl AsRef<Path>) -> Result<Arc<dyn ReadAt>> {
    Ok(Arc::new(FileSource::open(path)?))
}

fn read_content(package: &SourceBackedPackage) -> Result<String> {
    let bytes = package.get_file("content.xml")?;
    String::from_utf8(bytes).map_err(|error| {
        Error::InvalidFormat(format!("{FAMILY_NAME} content.xml is not UTF-8: {error}"))
    })
}

fn scan_catalog(xml: &str) -> Result<Vec<SheetCatalogEntry>> {
    // Keep the mandatory ODS hierarchy validation and the worksheet catalog
    // scan on one tokenizer.  The catalog is deliberately narrower than the
    // semantic worksheet parser, but it must retain the same structural
    // refusal boundary as the established owner.  Errors from the catalog
    // are delayed until the validation handler reaches EOF so a malformed XML
    // stream retains the historical validation/parser error precedence.
    validate_size(xml)?;
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut validate = ValidateHandler::default();
    let mut catalog_error = None;
    let mut depth = 0usize;
    let mut spreadsheet_depth = None;
    let mut entries = Vec::new();
    let mut names = HashSet::<String>::new();

    loop {
        let (resolved, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid ODS content.xml: {error}")))?;
        validate.on_event(is_office_namespace(&resolved), &event)?;
        let namespace = NamespaceKind::from_resolved(&resolved);
        match event {
            Event::Start(element) => {
                let local = element.local_name();
                if namespace == NamespaceKind::Office
                    && local.as_ref() == b"spreadsheet"
                    && depth == 2
                {
                    spreadsheet_depth = Some(depth);
                }
                if namespace == NamespaceKind::Table && local.as_ref() == b"table" {
                    if spreadsheet_depth != depth.checked_sub(1) {
                        catalog_error.get_or_insert_with(|| {
                            Error::InvalidFormat(
                                "table:table must be a direct child of office:spreadsheet"
                                    .to_string(),
                            )
                        });
                    } else if catalog_error.is_none() {
                        if let Err(error) =
                            add_catalog_entry(&mut entries, &mut names, &element, &reader)
                        {
                            catalog_error = Some(error);
                        }
                    }
                }
                depth = depth.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("ODS XML nesting depth overflows usize".to_string())
                })?;
            },
            Event::Empty(element) => {
                let local = element.local_name();
                if namespace == NamespaceKind::Office
                    && local.as_ref() == b"spreadsheet"
                    && depth == 2
                {
                    spreadsheet_depth = Some(depth);
                }
                if namespace == NamespaceKind::Table && local.as_ref() == b"table" {
                    if spreadsheet_depth != depth.checked_sub(1) {
                        catalog_error.get_or_insert_with(|| {
                            Error::InvalidFormat(
                                "table:table must be a direct child of office:spreadsheet"
                                    .to_string(),
                            )
                        });
                    } else if catalog_error.is_none() {
                        if let Err(error) =
                            add_catalog_entry(&mut entries, &mut names, &element, &reader)
                        {
                            catalog_error = Some(error);
                        }
                    }
                }
            },
            Event::End(_) => {
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("ODS XML element stack underflow".to_string())
                })?;
                if spreadsheet_depth == Some(depth) {
                    spreadsheet_depth = None;
                }
            },
            Event::Eof => break,
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
        buffer.clear();
    }
    validate.finish()?;
    if let Some(error) = catalog_error {
        return Err(error);
    }
    Ok(entries)
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum NamespaceKind {
    Office,
    Table,
    Other,
}

impl NamespaceKind {
    fn from_resolved(value: &ResolveResult<'_>) -> Self {
        match value {
            ResolveResult::Bound(Namespace(uri)) if *uri == OFFICE_NAMESPACE.as_bytes() => {
                Self::Office
            },
            ResolveResult::Bound(Namespace(uri)) if *uri == TABLE_NAMESPACE.as_bytes() => {
                Self::Table
            },
            ResolveResult::Unbound | ResolveResult::Bound(_) | ResolveResult::Unknown(_) => {
                Self::Other
            },
        }
    }
}

fn add_catalog_entry(
    entries: &mut Vec<SheetCatalogEntry>,
    names: &mut HashSet<String>,
    element: &BytesStart<'_>,
    reader: &NsReader<&[u8]>,
) -> Result<()> {
    if entries.len() >= crate::worksheet::validation::MAX_PHYSICAL_RUNS {
        return Err(Error::InvalidFormat(format!(
            "ODS sheet count exceeds the {} safety limit",
            crate::worksheet::validation::MAX_PHYSICAL_RUNS
        )));
    }
    let name = table_name(element, reader)?;
    crate::worksheet::validation::validate_text(&name, "sheet name")?;
    if name.is_empty() {
        return Err(Error::InvalidFormat(
            "ODS sheet names must be non-empty".to_string(),
        ));
    }
    names.try_reserve(1).map_err(|source| Error::Allocation {
        resource: "ODS source worksheet catalog names",
        source,
    })?;
    if !names.insert(name.clone()) {
        return Err(Error::InvalidFormat(format!(
            "ODS sheet name '{name}' is duplicated"
        )));
    }
    entries.try_reserve(1).map_err(|source| Error::Allocation {
        resource: "ODS source worksheet catalog",
        source,
    })?;
    entries.push(SheetCatalogEntry {
        index: entries.len(),
        name,
    });
    Ok(())
}

fn table_name(element: &BytesStart<'_>, reader: &NsReader<&[u8]>) -> Result<String> {
    let mut name = None;
    for raw in element.attributes().with_checks(true) {
        let raw =
            raw.map_err(|error| Error::InvalidFormat(format!("invalid ODS attribute: {error}")))?;
        let (namespace, local) = reader.resolver().resolve_attribute(raw.key);
        let value = raw
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid ODS attribute value: {error}"))
            })?;
        let namespace = match namespace {
            ResolveResult::Bound(Namespace(uri)) => uri,
            ResolveResult::Unbound => b"",
            ResolveResult::Unknown(prefix) => {
                return Err(Error::InvalidFormat(format!(
                    "unbound ODS attribute prefix '{}'",
                    String::from_utf8_lossy(prefix.as_ref())
                )));
            },
        };
        if namespace == TABLE_NAMESPACE.as_bytes() && local.as_ref() == b"name" {
            name = Some(value.into_owned());
        }
    }
    Ok(name.unwrap_or_else(|| "Sheet1".to_string()))
}

fn parse_selected_sheet(xml: &str, index: usize) -> Result<Option<Sheet>> {
    let (range, wrapper) = find_sheet_range(xml, index)?;
    let Some(range) = range else {
        return Ok(None);
    };
    let table = xml.get(range).ok_or_else(|| {
        Error::InvalidFormat("ODS selected worksheet source range is invalid".to_string())
    })?;
    let wrapped = wrapper.wrap(table)?;
    let mut sheets = crate::worksheet::codec::parse(&wrapped)?;
    if sheets.len() != 1 {
        return Err(Error::InvalidFormat(
            "ODS selected worksheet parser returned an unexpected sheet count".to_string(),
        ));
    }
    Ok(sheets.pop())
}

struct Wrapper {
    root: Vec<u8>,
    body: Vec<u8>,
    spreadsheet: Vec<u8>,
    root_name: Vec<u8>,
    body_name: Vec<u8>,
    spreadsheet_name: Vec<u8>,
}

impl Wrapper {
    fn wrap(&self, table: &str) -> Result<String> {
        let root_name = std::str::from_utf8(&self.root_name).map_err(|error| {
            Error::InvalidFormat(format!("ODS root name is not UTF-8: {error}"))
        })?;
        let body_name = std::str::from_utf8(&self.body_name).map_err(|error| {
            Error::InvalidFormat(format!("ODS body name is not UTF-8: {error}"))
        })?;
        let spreadsheet_name = std::str::from_utf8(&self.spreadsheet_name).map_err(|error| {
            Error::InvalidFormat(format!("ODS spreadsheet name is not UTF-8: {error}"))
        })?;
        let suffix_size = spreadsheet_name
            .len()
            .checked_add(body_name.len())
            .and_then(|size| size.checked_add(root_name.len()))
            .and_then(|size| size.checked_add(9))
            .ok_or_else(|| {
                Error::InvalidFormat("ODS selected worksheet suffix overflows".to_string())
            })?;
        let mut suffix = String::new();
        suffix
            .try_reserve_exact(suffix_size)
            .map_err(|source| Error::Allocation {
                resource: "ODS selected worksheet suffix",
                source,
            })?;
        suffix.push_str("</");
        suffix.push_str(spreadsheet_name);
        suffix.push_str("></");
        suffix.push_str(body_name);
        suffix.push_str("></");
        suffix.push_str(root_name);
        suffix.push('>');
        let size = self
            .root
            .len()
            .checked_add(self.body.len())
            .and_then(|size| size.checked_add(self.spreadsheet.len()))
            .and_then(|size| size.checked_add(table.len()))
            .and_then(|size| size.checked_add(suffix.len()))
            .ok_or_else(|| {
                Error::InvalidFormat("ODS selected worksheet wrapper overflows".to_string())
            })?;
        let mut output = String::new();
        output
            .try_reserve_exact(size)
            .map_err(|source| Error::Allocation {
                resource: "ODS selected worksheet wrapper",
                source,
            })?;
        output.push_str(std::str::from_utf8(&self.root).map_err(|error| {
            Error::InvalidFormat(format!("ODS root start tag is not UTF-8: {error}"))
        })?);
        output.push_str(std::str::from_utf8(&self.body).map_err(|error| {
            Error::InvalidFormat(format!("ODS body start tag is not UTF-8: {error}"))
        })?);
        output.push_str(std::str::from_utf8(&self.spreadsheet).map_err(|error| {
            Error::InvalidFormat(format!("ODS spreadsheet start tag is not UTF-8: {error}"))
        })?);
        output.push_str(table);
        output.push_str(&suffix);
        Ok(output)
    }
}

fn find_sheet_range(xml: &str, target: usize) -> Result<(Option<Range<usize>>, Wrapper)> {
    // `scan_catalog` validated this exact source before the catalog was
    // published.  The source version is checked before and after every public
    // read, so re-validating the entire document here would only repeat the
    // structural pass without widening the safety boundary.
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut spreadsheet_depth = None;
    let mut root = None;
    let mut body = None;
    let mut spreadsheet = None;
    let mut root_name = None;
    let mut body_name = None;
    let mut spreadsheet_name = None;
    let mut table_start = None;
    let mut table_depth = None;
    let mut table_index = 0usize;
    let mut selected_range = None;

    loop {
        let before = reader.buffer_position();
        let (resolved, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid ODS XML: {error}")))?;
        let namespace = NamespaceKind::from_resolved(&resolved);
        let after = reader.buffer_position();
        match event {
            Event::Start(element) => {
                let local = element.local_name();
                if depth == 0 {
                    root = Some(start_tag(&element)?);
                    root_name = Some(copy_bytes(
                        element.name().as_ref(),
                        "ODS selected worksheet root name",
                    )?);
                } else if depth == 1
                    && namespace == NamespaceKind::Office
                    && local.as_ref() == b"body"
                {
                    body = Some(start_tag(&element)?);
                    body_name = Some(copy_bytes(
                        element.name().as_ref(),
                        "ODS selected worksheet body name",
                    )?);
                } else if depth == 2
                    && namespace == NamespaceKind::Office
                    && local.as_ref() == b"spreadsheet"
                {
                    spreadsheet_depth = Some(depth);
                    spreadsheet = Some(start_tag(&element)?);
                    spreadsheet_name = Some(copy_bytes(
                        element.name().as_ref(),
                        "ODS selected worksheet spreadsheet name",
                    )?);
                }
                if namespace == NamespaceKind::Table
                    && local.as_ref() == b"table"
                    && spreadsheet_depth == depth.checked_sub(1)
                {
                    if table_index == target {
                        table_start = Some(before as usize);
                        table_depth = Some(depth + 1);
                    }
                    table_index = table_index.checked_add(1).ok_or_else(|| {
                        Error::InvalidFormat("ODS worksheet index overflows usize".to_string())
                    })?;
                }
                depth = depth.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("ODS XML nesting depth overflows usize".to_string())
                })?;
            },
            Event::Empty(element) => {
                let local = element.local_name();
                if namespace == NamespaceKind::Table
                    && local.as_ref() == b"table"
                    && spreadsheet_depth == depth.checked_sub(1)
                {
                    if table_index == target {
                        selected_range = Some(before as usize..after as usize);
                    }
                    table_index = table_index.checked_add(1).ok_or_else(|| {
                        Error::InvalidFormat("ODS worksheet index overflows usize".to_string())
                    })?;
                }
            },
            Event::End(_) => {
                if table_depth == Some(depth) {
                    table_depth = None;
                    if let Some(start) = table_start.take() {
                        selected_range = Some(start..after as usize);
                    }
                }
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("ODS XML element stack underflow".to_string())
                })?;
                if spreadsheet_depth == Some(depth) {
                    spreadsheet_depth = None;
                }
            },
            Event::Eof => break,
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
        buffer.clear();
    }

    let wrapper = Wrapper {
        root: root
            .ok_or_else(|| Error::InvalidFormat("ODS root start tag is missing".to_string()))?,
        body: body
            .ok_or_else(|| Error::InvalidFormat("ODS body start tag is missing".to_string()))?,
        spreadsheet: spreadsheet.ok_or_else(|| {
            Error::InvalidFormat("ODS spreadsheet start tag is missing".to_string())
        })?,
        root_name: root_name
            .ok_or_else(|| Error::InvalidFormat("ODS root name is missing".to_string()))?,
        body_name: body_name
            .ok_or_else(|| Error::InvalidFormat("ODS body name is missing".to_string()))?,
        spreadsheet_name: spreadsheet_name
            .ok_or_else(|| Error::InvalidFormat("ODS spreadsheet name is missing".to_string()))?,
    };
    Ok((selected_range, wrapper))
}

fn start_tag(element: &BytesStart<'_>) -> Result<Vec<u8>> {
    let raw: &[u8] = element.as_ref();
    let size = raw
        .len()
        .checked_add(2)
        .ok_or_else(|| Error::InvalidFormat("ODS start tag size overflows".to_string()))?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(size)
        .map_err(|source| Error::Allocation {
            resource: "ODS selected worksheet start tag",
            source,
        })?;
    output.push(b'<');
    output.extend_from_slice(raw);
    output.push(b'>');
    Ok(output)
}

fn copy_bytes(value: &[u8], resource: &'static str) -> Result<Vec<u8>> {
    let mut output = Vec::new();
    output
        .try_reserve_exact(value.len())
        .map_err(|source| Error::Allocation { resource, source })?;
    output.extend_from_slice(value);
    Ok(output)
}

fn ensure_current(expected: SourceVersion, observed: SourceVersion) -> Result<()> {
    if expected == observed {
        Ok(())
    } else {
        Err(Error::SourceChanged { expected, observed })
    }
}

fn prefer_current<T>(source: &dyn ReadAt, expected: SourceVersion, result: Result<T>) -> Result<T> {
    ensure_current(expected, source.version()?)?;
    result
}
