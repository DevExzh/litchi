//! Catalog-first, source-backed ODT reads.
//!
//! [`SourceBackedDocumentCatalog`] is an additive read-only lifecycle for
//! callers that need to discover visible text blocks and select one block at
//! a time. The ZIP index, package manifest, and compact block descriptors are
//! retained; `content.xml` is a temporary projection and styles, metadata,
//! media, and semantic models remain cold until explicitly requested.

use std::{fmt, sync::Arc};

#[cfg(any(unix, windows))]
use std::path::Path;

#[cfg(any(unix, windows))]
use litchi_core::FileSource;
use litchi_core::{Error, ReadAt, Result, SourceVersion};
use litchi_odf_common::{
    core::{SourceBackedPackage, validate_content_document_part},
    package::{is_media_path, resolve_package_path},
};
use zeroize::Zeroizing;

use super::ReadLimits;
use crate::elements::text::{Block, Kind, TextElements};

const ODF_TEXT: &str = "application/vnd.oasis.opendocument.text";
const FAMILY_NAME: &str = "ODT";
const MAX_CONTENT_BYTES: u64 = 256 * 1024 * 1024;

/// One visible paragraph or heading in a catalog-first ODT owner.
///
/// The index follows the same start order and suppression rules as the
/// established [`TextElements`] parser. No XML range or source implementation
/// detail is exposed because the current package owner rereads complete
/// compressed members for semantic selection.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct TextBlockCatalogEntry {
    index: usize,
    kind: Kind,
}

impl TextBlockCatalogEntry {
    /// Return the zero-based visible text-block position.
    #[must_use]
    pub const fn index(self) -> usize {
        self.index
    }

    /// Return whether this entry is a paragraph or heading.
    #[must_use]
    pub const fn kind(self) -> Kind {
        self.kind
    }
}

/// Immutable, catalog-first ODT access over a positional source.
///
/// Opening validates the source package, ODT MIME type, complete ODT content
/// hierarchy, and the visible paragraph/heading catalog. It retains only the
/// bounded descriptors. `content.xml` is dropped after opening and reread
/// temporarily by [`Self::block_at`]. Styles, metadata, media, and semantic
/// models are not read by this owner unless selected through an explicit
/// package-member method. [`Self::materialize`] is the explicit transition to
/// the complete owned and mutable [`super::Document`] owner.
pub struct SourceBackedDocumentCatalog {
    package: SourceBackedPackage,
    source: Arc<dyn ReadAt>,
    source_version: SourceVersion,
    entries: Arc<[TextBlockCatalogEntry]>,
}

impl fmt::Debug for SourceBackedDocumentCatalog {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SourceBackedDocumentCatalog")
            .field("source_version", &self.source_version)
            .field("text_blocks", &self.entries.len())
            .finish_non_exhaustive()
    }
}

impl SourceBackedDocumentCatalog {
    /// Open an ODT catalog from a filesystem path without slurping the source.
    #[cfg(any(unix, windows))]
    pub fn from_path(path: impl AsRef<Path>) -> Result<Self> {
        Self::from_read_at(file_source(path)?)
    }

    /// Open an ODT catalog from a filesystem path with explicit limits.
    #[cfg(any(unix, windows))]
    pub fn from_path_with_limits(path: impl AsRef<Path>, limits: ReadLimits) -> Result<Self> {
        Self::from_read_at_with_limits(file_source(path)?, limits)
    }

    /// Open an encrypted ODT catalog from a filesystem path.
    #[cfg(any(unix, windows))]
    pub fn from_path_with_password(
        path: impl AsRef<Path>,
        password: impl Into<String>,
    ) -> Result<Self> {
        let mut password = Zeroizing::new(password.into());
        Self::from_read_at_with_password(file_source(path)?, std::mem::take(&mut *password))
    }

    /// Open an encrypted ODT catalog from a filesystem path with limits.
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

    /// Open an ODT catalog from a caller-provided positional source.
    pub fn from_read_at(source: Arc<dyn ReadAt>) -> Result<Self> {
        Self::from_read_at_with_limits(source, ReadLimits::default())
    }

    /// Open an encrypted ODT catalog from a caller-provided source.
    pub fn from_read_at_with_password(
        source: Arc<dyn ReadAt>,
        password: impl Into<String>,
    ) -> Result<Self> {
        let mut password = Zeroizing::new(password.into());
        Self::from_read_at_inner(
            source,
            ReadLimits::default(),
            Some(std::mem::take(&mut *password)),
        )
    }

    /// Open an ODT catalog with explicit finite package limits.
    pub fn from_read_at_with_limits(source: Arc<dyn ReadAt>, limits: ReadLimits) -> Result<Self> {
        Self::from_read_at_inner(source, limits, None)
    }

    /// Open an encrypted ODT catalog with explicit finite package limits.
    pub fn from_read_at_with_limits_and_password(
        source: Arc<dyn ReadAt>,
        limits: ReadLimits,
        password: impl Into<String>,
    ) -> Result<Self> {
        let mut password = Zeroizing::new(password.into());
        Self::from_read_at_inner(source, limits, Some(std::mem::take(&mut *password)))
    }

    fn from_read_at_inner(
        source: Arc<dyn ReadAt>,
        limits: ReadLimits,
        password: Option<String>,
    ) -> Result<Self> {
        let package = match password {
            Some(password) => {
                let mut password = Zeroizing::new(password);
                SourceBackedPackage::from_read_at_with_limits_and_password(
                    Arc::clone(&source),
                    limits,
                    std::mem::take(&mut *password),
                )?
            },
            None => SourceBackedPackage::from_read_at_with_limits(Arc::clone(&source), limits)?,
        };
        let source_version = package.source_version()?;
        let parsed = (|| {
            let mimetype = package.mimetype()?;
            if mimetype != ODF_TEXT {
                return Err(Error::InvalidFormat(format!(
                    "expected {FAMILY_NAME} package MIME type '{ODF_TEXT}', found '{mimetype}'"
                )));
            }
            let content = read_content(&package)?;
            validate_content_document_part(&content, "<office:text", FAMILY_NAME)?;
            let kinds = TextElements::scan_block_kinds(&content)?;
            let mut entries = Vec::new();
            entries
                .try_reserve_exact(kinds.len())
                .map_err(|source| Error::Allocation {
                    resource: "ODT source text-block catalog",
                    source,
                })?;
            for (index, kind) in kinds.into_iter().enumerate() {
                entries.push(TextBlockCatalogEntry { index, kind });
            }
            Ok(entries)
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
        let value = self.package.len();
        prefer_current(self.source.as_ref(), self.source_version, Ok(value))
    }

    /// Return the validated ODT package MIME type.
    pub fn mimetype(&self) -> Result<&str> {
        self.check_source()?;
        let value = self.package.mimetype()?;
        self.check_source()?;
        Ok(value)
    }

    /// Borrow the retained visible text-block catalog.
    pub fn catalog(&self) -> Result<&[TextBlockCatalogEntry]> {
        self.check_source()?;
        let entries = self.entries.as_ref();
        prefer_current(self.source.as_ref(), self.source_version, Ok(entries))
    }

    /// Return the number of visible paragraphs and headings.
    pub fn text_block_count(&self) -> Result<usize> {
        self.check_source()?;
        let count = self.entries.len();
        prefer_current(self.source.as_ref(), self.source_version, Ok(count))
    }

    /// Read one visible paragraph or heading selected by document position.
    ///
    /// The selected parser rereads and scans all of `content.xml` so malformed
    /// unselected text still retains the established full-document validation
    /// behavior. Only the selected semantic block is retained.
    pub fn block_at(&self, index: usize) -> Result<Option<Block>> {
        self.check_source()?;
        if index >= self.entries.len() {
            self.check_source()?;
            return Ok(None);
        }
        let result = (|| {
            let content = read_content(&self.package)?;
            TextElements::parse_block_at(&content, index)
        })();
        prefer_current(self.source.as_ref(), self.source_version, result)
    }

    /// List package members without reading their payloads.
    pub fn files(&self) -> Result<Vec<String>> {
        let result = self.package.files();
        prefer_current(self.source.as_ref(), self.source_version, result)
    }

    /// List package media members without reading their payloads.
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
            if !is_media_path(&path) || !self.package.has_file(&path)? {
                return Ok(None);
            }
            self.package.get_file(&path).map(Some)
        })();
        prefer_current(self.source.as_ref(), self.source_version, result)
    }

    /// Read inert document and macro signature metadata without executing it.
    pub fn digital_signatures(&self) -> Result<litchi_odf_common::signature::DigitalSignatures> {
        let result = self.package.digital_signatures();
        prefer_current(self.source.as_ref(), self.source_version, result)
    }

    /// Materialize the exact source into the established mutable ODT owner.
    pub fn materialize(self) -> Result<super::Document> {
        self.check_source()?;
        let source = Arc::clone(&self.source);
        let source_version = self.source_version;
        let package = prefer_current(source.as_ref(), source_version, self.package.materialize())?;
        prefer_current(
            source.as_ref(),
            source_version,
            super::Document::from_owned_package(package),
        )
    }
}

#[cfg(any(unix, windows))]
fn file_source(path: impl AsRef<Path>) -> Result<Arc<dyn ReadAt>> {
    Ok(Arc::new(FileSource::open(path)?))
}

fn read_content(package: &SourceBackedPackage) -> Result<String> {
    if package
        .member_materialized_size("content.xml")?
        .is_some_and(|size| size > MAX_CONTENT_BYTES)
    {
        return Err(Error::InvalidFormat(format!(
            "{FAMILY_NAME} content.xml exceeds the family limit"
        )));
    }
    String::from_utf8(package.get_file("content.xml")?).map_err(|error| {
        Error::InvalidFormat(format!("{FAMILY_NAME} content.xml is not UTF-8: {error}"))
    })
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
