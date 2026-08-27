//! Catalog-first, source-backed ODP reads.
//!
//! [`SourceBackedPresentationCatalog`] is an additive lifecycle for callers
//! that need slide discovery and one slide at a time. Opening retains only
//! the validated positional package and bounded slide descriptors. The
//! `content.xml` catalog scan, selected slide XML, and optional `styles.xml`
//! are temporary projections and are never retained by this owner.

use std::{fmt, sync::Arc};

#[cfg(any(unix, windows))]
use std::path::Path;

#[cfg(any(unix, windows))]
use litchi_core::FileSource;
use litchi_core::{Error, ReadAt, Result, SourceVersion};
use litchi_odf_common::{
    core::{SourceBackedPackage, validate_content_document_part},
    package::resolve_package_path,
};
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::{Namespace, ResolveResult},
    reader::NsReader,
};
use zeroize::Zeroizing;

use super::ReadLimits;
use crate::{Reference, Slide, codec::Parser};

const ODF_PRESENTATION: &str = "application/vnd.oasis.opendocument.presentation";
const FAMILY_NAME: &str = "ODP";
const OFFICE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const DRAW_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const MAX_PAGES: usize = 65_536;
const MAX_NAME_BYTES: usize = 1024 * 1024;

/// One slide entry in a catalog-first ODP owner.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SlideCatalogEntry {
    index: usize,
    name: Option<String>,
}

impl SlideCatalogEntry {
    /// Return the zero-based slide position.
    #[must_use]
    pub const fn index(&self) -> usize {
        self.index
    }

    /// Return the exact producer-supplied `draw:name`, when present.
    #[must_use]
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }
}

/// Immutable, catalog-first ODP access over a positional source.
///
/// Opening validates the source package, MIME type, and direct ODP content
/// hierarchy, then retains only slide positions and optional names. It does
/// not retain `content.xml`, `styles.xml`, metadata, slide models, or media.
/// [`Self::slide_at`] is the explicit semantic-read boundary and
/// [`Self::materialize`] is the explicit transition to the complete mutable
/// [`super::Presentation`] owner.
pub struct SourceBackedPresentationCatalog {
    package: SourceBackedPackage,
    source: Arc<dyn ReadAt>,
    source_version: SourceVersion,
    entries: Arc<[SlideCatalogEntry]>,
}

impl fmt::Debug for SourceBackedPresentationCatalog {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SourceBackedPresentationCatalog")
            .field("source_version", &self.source_version)
            .field("slides", &self.entries.len())
            .finish_non_exhaustive()
    }
}

impl SourceBackedPresentationCatalog {
    /// Open an ODP catalog from a filesystem path without slurping the source.
    #[cfg(any(unix, windows))]
    pub fn from_path(path: impl AsRef<Path>) -> Result<Self> {
        Self::from_read_at(file_source(path)?)
    }

    /// Open an ODP catalog from a filesystem path with explicit limits.
    #[cfg(any(unix, windows))]
    pub fn from_path_with_limits(path: impl AsRef<Path>, limits: ReadLimits) -> Result<Self> {
        Self::from_read_at_with_limits(file_source(path)?, limits)
    }

    /// Open an encrypted ODP catalog from a filesystem path.
    #[cfg(any(unix, windows))]
    pub fn from_path_with_password(
        path: impl AsRef<Path>,
        password: impl Into<String>,
    ) -> Result<Self> {
        let mut password = Zeroizing::new(password.into());
        let source = file_source(path)?;
        Self::from_read_at_with_password(source, std::mem::take(&mut *password))
    }

    /// Open an encrypted ODP catalog from a filesystem path with limits.
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

    /// Open an ODP catalog from a caller-provided positional source.
    pub fn from_read_at(source: Arc<dyn ReadAt>) -> Result<Self> {
        Self::from_read_at_with_limits(source, ReadLimits::default())
    }

    /// Open an encrypted ODP catalog from a caller-provided source.
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

    /// Open an ODP catalog with explicit finite ZIP limits.
    pub fn from_read_at_with_limits(source: Arc<dyn ReadAt>, limits: ReadLimits) -> Result<Self> {
        Self::from_read_at_inner(source, limits)
    }

    /// Open an encrypted ODP catalog with explicit finite ZIP limits.
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
            if mimetype != ODF_PRESENTATION {
                return Err(Error::InvalidFormat(format!(
                    "expected {FAMILY_NAME} package MIME type '{ODF_PRESENTATION}', found '{mimetype}'"
                )));
            }
            let content = read_content(&package)?;
            // The shared validator derives only the expected local name from
            // this API marker and resolves the office namespace, so producer
            // documents may use any valid prefix for `presentation`.
            validate_content_document_part(&content, "<office:presentation", FAMILY_NAME)?;
            scan_catalog(&content)
        })();
        let entries = prefer_current(source.as_ref(), source_version, parsed)?;

        Ok(Self {
            package,
            source,
            source_version,
            entries: Arc::from(entries),
        })
    }

    /// Check the positional source identity without reading a member.
    pub fn check_source(&self) -> Result<()> {
        ensure_current(self.source_version, self.source.version()?)
    }

    /// Return the source version captured at open.
    #[must_use = "the captured source version identifies this source snapshot"]
    pub fn source_version(&self) -> Result<SourceVersion> {
        self.check_source()?;
        let value = self.source_version;
        self.check_source()?;
        Ok(value)
    }

    /// Return the physical source length captured at open.
    pub fn source_len(&self) -> Result<u64> {
        self.check_source()?;
        let length = self.package.len();
        prefer_current(self.source.as_ref(), self.source_version, Ok(length))
    }

    /// Borrow the retained slide catalog.
    pub fn catalog(&self) -> Result<&[SlideCatalogEntry]> {
        self.check_source()?;
        let entries = self.entries.as_ref();
        prefer_current(self.source.as_ref(), self.source_version, Ok(entries))
    }

    /// Return the number of slides without materializing slide models.
    pub fn slide_count(&self) -> Result<usize> {
        self.check_source()?;
        let count = self.entries.len();
        prefer_current(self.source.as_ref(), self.source_version, Ok(count))
    }

    /// Read one slide selected by its zero-based document position.
    ///
    /// The selected query temporarily reads `content.xml` and optional
    /// `styles.xml`, then uses the established parser. That parser retains
    /// full-document validation semantics, so malformed unselected content
    /// can still make this operation fail.
    pub fn slide_at(&self, index: usize) -> Result<Option<Slide>> {
        self.check_source()?;
        if index >= self.entries.len() {
            self.check_source()?;
            return Ok(None);
        }
        let result = (|| {
            let content = read_content(&self.package)?;
            let styles = read_styles(&self.package)?;
            Parser::parse_slide_with_styles_at(&content, styles.as_deref(), index)
        })();
        prefer_current(self.source.as_ref(), self.source_version, result)
    }

    /// Read the first slide with an exact producer-supplied `draw:name`.
    pub fn slide(&self, name: &str) -> Result<Option<Slide>> {
        self.check_source()?;
        let index = self
            .entries
            .iter()
            .find(|entry| entry.name.as_deref() == Some(name))
            .map(SlideCatalogEntry::index);
        let Some(index) = index else {
            self.check_source()?;
            return Ok(None);
        };
        self.slide_at(index)
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
        let result = (|| {
            self.check_source()?;
            let path = resolve_package_path(path)?;
            if !self.package.has_file(&path)? {
                return Ok(None);
            }
            Ok(Some(self.package.get_file(&path)?))
        })();
        prefer_current(self.source.as_ref(), self.source_version, result)
    }

    /// Read one package-contained media payload on demand.
    pub fn media_data(&self, media: &Reference) -> Result<Option<Vec<u8>>> {
        let result = (|| {
            self.check_source()?;
            let Some(path) = media.package_path() else {
                return Ok(None);
            };
            if !self.package.has_file(path)? {
                return Ok(None);
            }
            Ok(Some(self.package.get_file(path)?))
        })();
        prefer_current(self.source.as_ref(), self.source_version, result)
    }

    /// Read inert document and macro signature metadata.
    pub fn digital_signatures(&self) -> Result<litchi_odf_common::signature::DigitalSignatures> {
        let result = self.package.digital_signatures();
        prefer_current(self.source.as_ref(), self.source_version, result)
    }

    /// Materialize the exact source into the established mutable presentation owner.
    pub fn materialize(self) -> Result<super::Presentation> {
        self.package
            .materialize()
            .and_then(super::Presentation::from_owned_package)
    }
}

#[cfg(any(unix, windows))]
fn file_source(path: impl AsRef<Path>) -> Result<Arc<dyn ReadAt>> {
    Ok(Arc::new(FileSource::open(path)?))
}

fn read_content(package: &SourceBackedPackage) -> Result<String> {
    String::from_utf8(package.get_file("content.xml")?).map_err(|error| {
        Error::InvalidFormat(format!("{FAMILY_NAME} content.xml is not UTF-8: {error}"))
    })
}

fn read_styles(package: &SourceBackedPackage) -> Result<Option<String>> {
    if !package.has_file("styles.xml")? {
        return Ok(None);
    }
    String::from_utf8(package.get_file("styles.xml")?)
        .map(Some)
        .map_err(|error| {
            Error::InvalidFormat(format!("{FAMILY_NAME} styles.xml is not UTF-8: {error}"))
        })
}

fn scan_catalog(xml: &str) -> Result<Vec<SlideCatalogEntry>> {
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut root_seen = false;
    let mut body_seen = false;
    let mut body_depth = None;
    let mut presentation_seen = false;
    let mut presentation_depth = None;
    let mut entries = Vec::new();

    loop {
        match reader.read_event_into(&mut buffer).map_err(xml_error)? {
            Event::Start(element) => {
                let parent_depth = depth;
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("ODP content.xml nesting overflow"))?;
                if parent_depth == 0 {
                    if root_seen
                        || !element_is(&reader, &element, OFFICE_NAMESPACE, b"document-content")
                    {
                        return Err(invalid(
                            "ODP content.xml root is not office:document-content",
                        ));
                    }
                    root_seen = true;
                } else if parent_depth == 1
                    && element_is(&reader, &element, OFFICE_NAMESPACE, b"body")
                {
                    if body_seen {
                        return Err(invalid("duplicate office:body element in ODP content.xml"));
                    }
                    body_seen = true;
                    body_depth = Some(depth);
                } else if body_depth == Some(parent_depth)
                    && element_is(&reader, &element, OFFICE_NAMESPACE, b"presentation")
                {
                    if presentation_seen {
                        return Err(invalid(
                            "duplicate office:presentation element in ODP content.xml",
                        ));
                    }
                    presentation_seen = true;
                    presentation_depth = Some(depth);
                } else if presentation_depth == Some(parent_depth)
                    && element_is(&reader, &element, DRAW_NAMESPACE, b"page")
                {
                    push_entry(&mut entries, &reader, &element)?;
                } else if body_depth == Some(parent_depth) {
                    return Err(invalid("ODP office:body must contain office:presentation"));
                }
            },
            Event::Empty(element) => {
                let parent_depth = depth;
                if parent_depth == 0 {
                    return Err(invalid(
                        "ODP content.xml root office:document-content cannot be empty",
                    ));
                }
                if parent_depth == 1 && element_is(&reader, &element, OFFICE_NAMESPACE, b"body") {
                    if body_seen {
                        return Err(invalid("duplicate office:body element in ODP content.xml"));
                    }
                    body_seen = true;
                } else if body_depth == Some(parent_depth)
                    && element_is(&reader, &element, OFFICE_NAMESPACE, b"presentation")
                {
                    if presentation_seen {
                        return Err(invalid(
                            "duplicate office:presentation element in ODP content.xml",
                        ));
                    }
                    presentation_seen = true;
                } else if presentation_depth == Some(parent_depth)
                    && element_is(&reader, &element, DRAW_NAMESPACE, b"page")
                {
                    push_entry(&mut entries, &reader, &element)?;
                } else if body_depth == Some(parent_depth) {
                    return Err(invalid("ODP office:body must contain office:presentation"));
                }
            },
            Event::End(element) => {
                if depth == 0 {
                    return Err(invalid("ODP content.xml depth underflow"));
                }
                if presentation_depth == Some(depth)
                    && end_is(&reader, &element, OFFICE_NAMESPACE, b"presentation")
                {
                    presentation_depth = None;
                }
                if body_depth == Some(depth) && end_is(&reader, &element, OFFICE_NAMESPACE, b"body")
                {
                    body_depth = None;
                }
                depth -= 1;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("active XML declarations are prohibited"));
            },
            Event::Eof => break,
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::GeneralRef(_) => {},
        }
        buffer.clear();
    }

    if depth != 0 {
        return Err(invalid("unterminated ODP content.xml element"));
    }
    if !root_seen {
        return Err(invalid(
            "ODP content.xml has no office:document-content root",
        ));
    }
    if !body_seen {
        return Err(invalid("ODP content.xml has no office:body"));
    }
    if !presentation_seen {
        return Err(invalid("ODP content.xml has no office:presentation"));
    }
    Ok(entries)
}

fn push_entry(
    entries: &mut Vec<SlideCatalogEntry>,
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<()> {
    if entries.len() >= MAX_PAGES {
        return Err(invalid(format!(
            "ODP content.xml holds more than {MAX_PAGES} slides"
        )));
    }
    let mut name = None;
    for raw_attribute in element.attributes() {
        let attribute = raw_attribute.map_err(xml_error)?;
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        if matches!(namespace, ResolveResult::Bound(found) if found == Namespace(DRAW_NAMESPACE))
            && local.as_ref() == b"name"
        {
            if name.is_some() {
                return Err(invalid("duplicate draw:name presentation page attribute"));
            }
            if attribute.value.len() > MAX_NAME_BYTES {
                return Err(invalid("ODP draw:name exceeds 1 MiB"));
            }
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                .map_err(xml_error)?
                .into_owned();
            validate_name(&value)?;
            name = Some(value);
        }
    }
    entries.try_reserve(1).map_err(|source| Error::Allocation {
        resource: "ODP source catalog slide entries",
        source,
    })?;
    entries.push(SlideCatalogEntry {
        index: entries.len(),
        name,
    });
    Ok(())
}

fn validate_name(value: &str) -> Result<()> {
    if value.is_empty() {
        return Err(invalid("draw:name cannot be empty"));
    }
    if value.len() > MAX_NAME_BYTES {
        return Err(invalid("ODP draw:name exceeds 1 MiB"));
    }
    if value.chars().any(|character| {
        character == '\0' || (character.is_control() && !matches!(character, '\t' | '\n' | '\r'))
    }) {
        return Err(invalid("draw:name contains invalid XML characters"));
    }
    Ok(())
}

fn element_is(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    namespace: &[u8],
    local: &[u8],
) -> bool {
    let (resolved, local_name) = reader.resolver().resolve_element(element.name());
    matches!(resolved, ResolveResult::Bound(found) if found == Namespace(namespace))
        && local_name.as_ref() == local
}

fn end_is(
    reader: &NsReader<&[u8]>,
    element: &quick_xml::events::BytesEnd<'_>,
    namespace: &[u8],
    local: &[u8],
) -> bool {
    let (resolved, local_name) = reader.resolver().resolve_element(element.name());
    matches!(resolved, ResolveResult::Bound(found) if found == Namespace(namespace))
        && local_name.as_ref() == local
}

fn read_xml_error(error: impl fmt::Display) -> Error {
    invalid(format!("ODP content.xml XML parsing error: {error}"))
}

fn xml_error(error: impl fmt::Display) -> Error {
    read_xml_error(error)
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

fn ensure_current(expected: SourceVersion, observed: SourceVersion) -> Result<()> {
    if expected == observed {
        Ok(())
    } else {
        Err(Error::SourceChanged { expected, observed })
    }
}

fn prefer_current<T>(source: &dyn ReadAt, expected: SourceVersion, result: Result<T>) -> Result<T> {
    match source.version() {
        Err(error) => Err(error.into()),
        Ok(observed) if observed != expected => Err(Error::SourceChanged { expected, observed }),
        Ok(_) => result,
    }
}
