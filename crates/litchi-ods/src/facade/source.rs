//! Immutable, positional-source ODS reads and focused existing-cell edits.
//!
//! [`SourceBackedSpreadsheet`] retains a ZIP index and validated semantic XML
//! while unrelated package members stay cold until selected. Focused ordinary
//! existing-cell edits publish through the source-backed `content.xml` path;
//! callers explicitly call [`SourceBackedSpreadsheet::materialize`] before
//! entering the complete owned/mutable [`super::Spreadsheet`] boundary.

use std::{
    fmt,
    sync::{
        Arc, OnceLock,
        atomic::{AtomicUsize, Ordering},
    },
};

#[cfg(any(unix, windows))]
use litchi_core::FileSource;
use litchi_core::{Error, Metadata, ReadAt, Result, SourceVersion};
use litchi_odf_common::{
    core::{SourceBackedPackage, SourceContentProof, SourcePackageLimits, validate_content_part},
    package::{is_media_path, resolve_package_path},
};
#[cfg(any(unix, windows))]
use std::path::Path;
use zeroize::Zeroizing;

use super::CellSelector;
use crate::{
    codec::names,
    model::names::Definition,
    settings::Settings,
    worksheet::{CellView, Sheet},
};

const ODF_SPREADSHEET: &str = "application/vnd.oasis.opendocument.spreadsheet";
const BODY_MARKER: &str = "<office:spreadsheet";
const FAMILY_NAME: &str = "ODS";

/// The bounded ZIP-index policy used by positional ODF facades.
pub type ReadLimits = SourcePackageLimits;

/// ODS access over an immutable positional source.
///
/// Opening validates the ZIP archive, MIME member, manifest, content root,
/// and worksheet graph, then retains only the required XML projections.
/// Package members such as pictures and embedded objects are read only when
/// selected through [`Self::member_data`] or [`Self::media_data`].  Every
/// public operation checks the captured [`SourceVersion`] and reports
/// [`Error::SourceChanged`] when the source no longer identifies the same
/// snapshot. [`Self::edit_cells`] adds a narrow existing-cell transaction
/// without materializing the complete package.
pub struct SourceBackedSpreadsheet {
    pub(super) package: SourceBackedPackage,
    source: Arc<dyn ReadAt>,
    pub(super) source_version: SourceVersion,
    pub(super) content_xml: Arc<str>,
    pub(super) content_proof: SourceContentProof,
    pub(super) styles_xml: Option<Arc<str>>,
    definitions: Vec<Definition>,
    pub(super) sheets: Arc<[Sheet]>,
    metadata: crate::metadata::Snapshot,
    settings: Option<Settings>,
    cell_queries: AtomicUsize,
    cell_locator: OnceLock<Option<super::cell_locator::CellLocator>>,
}

impl fmt::Debug for SourceBackedSpreadsheet {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SourceBackedSpreadsheet")
            .field("source_version", &self.source_version)
            .field("sheets", &self.sheets.len())
            .field("definitions", &self.definitions.len())
            .finish_non_exhaustive()
    }
}

impl SourceBackedSpreadsheet {
    /// Open an ODS from a regular filesystem file without slurping it.
    #[cfg(any(unix, windows))]
    pub fn from_path(path: impl AsRef<Path>) -> Result<Self> {
        Self::from_read_at(file_source(path)?)
    }

    /// Open an ODS from a regular filesystem file with explicit ZIP limits.
    #[cfg(any(unix, windows))]
    pub fn from_path_with_limits(path: impl AsRef<Path>, limits: ReadLimits) -> Result<Self> {
        Self::from_read_at_with_limits(file_source(path)?, limits)
    }

    /// Open a password-protected ODS from a regular filesystem file.
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

    /// Open a password-protected ODS from a regular filesystem file with
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

    /// Open an ODS from a caller-provided positional source.
    pub fn from_read_at(source: Arc<dyn ReadAt>) -> Result<Self> {
        Self::from_read_at_with_limits(source, ReadLimits::default())
    }

    /// Open a password-protected ODS from a caller-provided positional source.
    pub fn from_read_at_with_password(
        source: Arc<dyn ReadAt>,
        password: impl Into<String>,
    ) -> Result<Self> {
        Self::from_read_at_inner_with_password(source, ReadLimits::default(), password)
    }

    /// Open an ODS from a positional source with explicit finite limits.
    pub fn from_read_at_with_limits(source: Arc<dyn ReadAt>, limits: ReadLimits) -> Result<Self> {
        Self::from_read_at_inner(source, limits)
    }

    /// Open a password-protected ODS from a positional source with explicit
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
        // The common owner immediately wraps the credential in Zeroizing.
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
            if mimetype != ODF_SPREADSHEET {
                return Err(Error::InvalidFormat(format!(
                    "expected {FAMILY_NAME} package MIME type '{ODF_SPREADSHEET}', found '{mimetype}'"
                )));
            }

            let (content_bytes, content_proof) = package.get_content_xml_with_source_proof()?;
            let content_xml = String::from_utf8(content_bytes).map_err(|error| {
                Error::InvalidFormat(format!("{FAMILY_NAME} content.xml is not UTF-8: {error}"))
            })?;
            validate_content_part(&content_xml, BODY_MARKER, FAMILY_NAME)?;
            crate::authoring::validate_content_xml(&content_xml)?;
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

            let metadata_source = if package.has_file("meta.xml")? {
                Some(
                    String::from_utf8(package.get_file("meta.xml")?).map_err(|error| {
                        Error::InvalidFormat(format!(
                            "{FAMILY_NAME} meta.xml is not UTF-8: {error}"
                        ))
                    })?,
                )
            } else {
                None
            };
            let metadata = crate::metadata::Snapshot::from_source(metadata_source)?;
            let settings = crate::settings::Snapshot::from_content_xml(&content_xml)?
                .calculation()
                .cloned();
            let definitions = names::parse(&content_xml)?;
            let sheets = crate::worksheet::codec::parse(&content_xml)?;
            Ok((
                content_xml,
                content_proof,
                styles_xml,
                definitions,
                sheets,
                metadata,
                settings,
            ))
        })();
        let (content_xml, content_proof, styles_xml, definitions, sheets, metadata, settings) =
            prefer_current(source.as_ref(), source_version, parsed)?;

        Ok(Self {
            package,
            source,
            source_version,
            content_xml: Arc::from(content_xml),
            content_proof,
            styles_xml: styles_xml.map(Arc::from),
            definitions,
            sheets: Arc::from(sheets),
            metadata,
            settings,
            cell_queries: AtomicUsize::new(0),
            cell_locator: OnceLock::new(),
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

    /// Return the captured source length.
    pub fn source_len(&self) -> Result<u64> {
        self.check_source()?;
        let value = self.package.len();
        self.check_source()?;
        Ok(value)
    }

    /// Borrow the validated `content.xml` snapshot.
    pub fn content_xml(&self) -> Result<&str> {
        self.check_source()?;
        let value = self.content_xml.as_ref();
        self.check_source()?;
        Ok(value)
    }

    /// Borrow the optional validated `styles.xml` snapshot.
    pub fn styles_xml(&self) -> Result<Option<&str>> {
        self.check_source()?;
        let value = self.styles_xml.as_deref();
        self.check_source()?;
        Ok(value)
    }

    /// Return compact cross-format metadata projected from `meta.xml`.
    pub fn metadata(&self) -> Result<&Metadata> {
        self.check_source()?;
        let value = self.metadata.value();
        self.check_source()?;
        Ok(value)
    }

    /// Return the complete typed ODF metadata projection.
    pub fn odf_metadata(&self) -> Result<&crate::metadata::Metadata> {
        self.check_source()?;
        let value = self.metadata.odf();
        self.check_source()?;
        Ok(value)
    }

    /// Return the retained metadata snapshot, including optional source XML.
    pub fn metadata_snapshot(&self) -> Result<&crate::metadata::Snapshot> {
        self.check_source()?;
        let value = &self.metadata;
        self.check_source()?;
        Ok(value)
    }

    /// Return spreadsheet calculation settings, if declared.
    pub fn settings(&self) -> Result<Option<&Settings>> {
        self.check_source()?;
        let value = self.settings.as_ref();
        self.check_source()?;
        Ok(value)
    }

    /// Alias for [`Self::settings`].
    pub fn calculation_settings(&self) -> Result<Option<&Settings>> {
        self.settings()
    }

    /// Return named definitions in document order.
    pub fn definitions(&self) -> Result<&[Definition]> {
        self.check_source()?;
        let value = self.definitions.as_slice();
        self.check_source()?;
        Ok(value)
    }

    /// Return typed worksheets in document order.
    pub fn sheets(&self) -> Result<&[Sheet]> {
        self.check_source()?;
        let value = self.sheets.as_ref();
        self.check_source()?;
        Ok(value)
    }

    /// Return the number of worksheets.
    pub fn sheet_count(&self) -> Result<usize> {
        self.check_source()?;
        let value = self.sheets.len();
        self.check_source()?;
        Ok(value)
    }

    /// Return worksheet names in document order.
    pub fn sheet_names(&self) -> Result<Vec<String>> {
        self.check_source()?;
        let result = (|| {
            let mut names = Vec::new();
            names
                .try_reserve_exact(self.sheets.len())
                .map_err(|source| Error::Allocation {
                    resource: "ODS source worksheet names",
                    source,
                })?;
            names.extend(self.sheets.iter().map(|sheet| sheet.name.clone()));
            Ok(names)
        })();
        prefer_current(self.source.as_ref(), self.source_version, result)
    }

    /// Select one worksheet by exact ODF name.
    pub fn sheet(&self, name: &str) -> Result<Option<&Sheet>> {
        self.check_source()?;
        let value = self.sheets.iter().find(|sheet| sheet.name == name);
        self.check_source()?;
        Ok(value)
    }

    /// Select one worksheet by checked zero-based position.
    pub fn sheet_at(&self, index: usize) -> Result<Option<&Sheet>> {
        self.check_source()?;
        let value = self.sheets.get(index);
        self.check_source()?;
        Ok(value)
    }

    /// Select one physical cell run by worksheet name and logical position.
    pub fn cell(
        &self,
        sheet_name: &str,
        row: usize,
        column: usize,
    ) -> Result<Option<CellView<'_>>> {
        self.check_source()?;
        let value = self.cell_unchecked(CellSelector::new(sheet_name, row, column));
        self.check_source()?;
        Ok(value)
    }

    fn cell_unchecked(&self, selector: CellSelector<'_>) -> Option<CellView<'_>> {
        let sheet_index = self
            .sheets
            .iter()
            .position(|sheet| sheet.name == selector.sheet_name)?;
        let direct = || self.sheets[sheet_index].cell_view(selector.row, selector.column);

        if let Some(locator) = self.cell_locator.get() {
            return Some(locator.as_ref().map_or_else(direct, |locator| {
                locator.cell_view(&self.sheets, sheet_index, selector.row, selector.column)
            }));
        }

        let previous = self
            .cell_queries
            .fetch_update(Ordering::Relaxed, Ordering::Relaxed, |count| {
                Some(count.saturating_add(1))
            })
            .unwrap_or(usize::MAX);
        if previous.saturating_add(1) >= super::cell_locator::BUILD_QUERY_THRESHOLD {
            let locator = self
                .cell_locator
                .get_or_init(|| super::cell_locator::CellLocator::try_build(&self.sheets));
            return Some(locator.as_ref().map_or_else(direct, |locator| {
                locator.cell_view(&self.sheets, sheet_index, selector.row, selector.column)
            }));
        }

        Some(direct())
    }

    /// Look up an ordered batch of logical cells with one source freshness
    /// check before and one after the complete operation.
    ///
    /// A missing worksheet produces `None` for that selector, while an
    /// existing worksheet with no physical cell at the coordinate produces
    /// `Some(CellView::Missing)`, exactly matching [`Self::cell`].  Results
    /// retain selector order and duplicate selectors are allowed.  The
    /// lookup and any adaptive index construction perform no source I/O.
    ///
    /// # Errors
    ///
    /// Returns [`Error::SourceChanged`] if the source is stale before or during
    /// the batch, a typed allocation error when the result vector cannot be
    /// reserved, or an invalid-format error when the selector bound is
    /// exceeded.  A stale source takes precedence over an operation error
    /// observed while finalizing the batch.
    pub fn cell_batch(&self, selectors: &[CellSelector<'_>]) -> Result<Vec<Option<CellView<'_>>>> {
        self.check_source()?;
        let result = self.cell_batch_unchecked(selectors);
        let final_check = self.check_source();
        match final_check {
            Err(error) => Err(error),
            Ok(()) => result,
        }
    }

    fn cell_batch_unchecked(
        &self,
        selectors: &[CellSelector<'_>],
    ) -> Result<Vec<Option<CellView<'_>>>> {
        super::validate_cell_batch_len(selectors.len())?;
        let mut values = Vec::new();
        values
            .try_reserve_exact(selectors.len())
            .map_err(|source| Error::Allocation {
                resource: "ODS source cell lookup batch results",
                source,
            })?;
        for &selector in selectors {
            values.push(self.cell_unchecked(selector));
        }
        Ok(values)
    }

    /// Alias for [`Self::cell_batch`].
    pub fn cells(&self, selectors: &[CellSelector<'_>]) -> Result<Vec<Option<CellView<'_>>>> {
        self.cell_batch(selectors)
    }

    /// Extract displayed worksheet text using tab-separated cells and
    /// newline-separated rows, preserving sheet order.
    pub fn text(&self) -> Result<String> {
        self.check_source()?;
        let result = project_text(&self.sheets);
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

    /// Read one safe, package-contained member on demand.
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

    /// Read one package media member selected by the established ODF image
    /// inventory without following links or fragments.
    pub fn image_data(&self, image: &litchi_odf_common::media::Image) -> Result<Option<Vec<u8>>> {
        let Some(path) = image.package_path() else {
            self.check_source()?;
            return Ok(None);
        };
        // `Image::package_path` is already a manifest-aware, safe package
        // part selected by the ODF media scanner.  Do not apply the
        // Pictures/uppercase-sensitive media inventory heuristic here.
        self.member_data(path)
    }

    /// Explicitly materialize this source into the existing mutable ODS
    /// spreadsheet owner.
    pub fn materialize(self) -> Result<super::Spreadsheet> {
        self.check_source()?;
        let source = Arc::clone(&self.source);
        let source_version = self.source_version;
        let package = prefer_current(source.as_ref(), source_version, self.package.materialize())?;
        let result = super::Spreadsheet::from_owned_package(package);
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

pub(crate) fn project_text(sheets: &[Sheet]) -> Result<String> {
    let capacity = text_capacity(sheets)?;
    let mut output = String::new();
    output
        .try_reserve_exact(capacity)
        .map_err(|source| Error::Allocation {
            resource: "ODS source text projection",
            source,
        })?;
    for (sheet_index, sheet) in sheets.iter().enumerate() {
        if sheet_index != 0 {
            output.push('\n');
        }
        for (row_index, row) in sheet.rows.iter().enumerate() {
            if row_index != 0 {
                output.push('\n');
            }
            for logical_row in 0..row.repeat() {
                if logical_row != 0 {
                    output.push('\n');
                }
                let mut first = true;
                for cell in &row.cells {
                    for _ in 0..cell.repeat() {
                        if !first {
                            output.push('\t');
                        }
                        first = false;
                        output.push_str(&cell.text);
                    }
                }
            }
        }
    }
    Ok(output)
}

fn text_capacity(sheets: &[Sheet]) -> Result<usize> {
    let mut total = 0usize;
    for (sheet_index, sheet) in sheets.iter().enumerate() {
        if sheet_index != 0 {
            total = total
                .checked_add(1)
                .ok_or_else(|| Error::InvalidFormat("ODS source text size overflow".to_string()))?;
            if total > crate::worksheet::validation::MAX_TEXT_BYTES {
                return Err(Error::InvalidFormat(
                    "ODS source text projection exceeds the safety limit".to_string(),
                ));
            }
        }
        for (row_index, row) in sheet.rows.iter().enumerate() {
            if row_index != 0 {
                total = total.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("ODS source text size overflow".to_string())
                })?;
            }
            let row_repeat = row.repeat();
            let mut row_size = 0usize;
            let mut cells = 0usize;
            for cell in &row.cells {
                let repeat = cell.repeat();
                let payload = cell.text.len().checked_mul(repeat).ok_or_else(|| {
                    Error::InvalidFormat("ODS source text size overflow".to_string())
                })?;
                row_size = row_size.checked_add(payload).ok_or_else(|| {
                    Error::InvalidFormat("ODS source text size overflow".to_string())
                })?;
                cells = cells.checked_add(repeat).ok_or_else(|| {
                    Error::InvalidFormat("ODS source text size overflow".to_string())
                })?;
            }
            row_size = row_size
                .checked_add(cells.saturating_sub(1))
                .ok_or_else(|| Error::InvalidFormat("ODS source text size overflow".to_string()))?;
            let repeated_row_separators = row_repeat.checked_sub(1).ok_or_else(|| {
                Error::InvalidFormat("ODS row repetition must be positive".to_string())
            })?;
            total = total
                .checked_add(row_size.checked_mul(row_repeat).ok_or_else(|| {
                    Error::InvalidFormat("ODS source text size overflow".to_string())
                })?)
                .ok_or_else(|| Error::InvalidFormat("ODS source text size overflow".to_string()))?;
            total = total
                .checked_add(repeated_row_separators)
                .ok_or_else(|| Error::InvalidFormat("ODS source text size overflow".to_string()))?;
            if total > crate::worksheet::validation::MAX_TEXT_BYTES {
                return Err(Error::InvalidFormat(
                    "ODS source text projection exceeds the safety limit".to_string(),
                ));
            }
        }
    }
    Ok(total)
}

#[cfg(test)]
mod tests {
    use super::{ReadLimits, SourceBackedSpreadsheet, project_text};
    use crate::worksheet::{Cell, CellValue, Row, Sheet};
    use crate::{CellSelector, MAX_CELL_SELECTORS};
    use litchi_core::{OwnedSource, ReadAt, SourceVersion};
    use std::io::{Cursor, Read, Write};
    use std::ptr;
    use std::sync::{
        Arc,
        atomic::{AtomicU64, AtomicUsize, Ordering},
    };
    use zip::{ZipArchive, ZipWriter, write::SimpleFileOptions};

    const MIME: &str = "application/vnd.oasis.opendocument.spreadsheet";
    const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
    const TABLE: &str = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
    const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";

    fn package() -> Vec<u8> {
        let content = format!(
            r#"<office:document-content xmlns:office="{OFFICE}" xmlns:table="{TABLE}" xmlns:text="{TEXT}"><office:body><office:spreadsheet><table:table table:name="Sheet1"><table:table-row table:number-rows-repeated="2"><table:table-cell office:value-type="string"><text:p>hello</text:p></table:table-cell><table:table-cell office:value-type="string"><text:p>world</text:p></table:table-cell></table:table-row></table:table><table:table table:name="Sheet2"/></office:spreadsheet></office:body></office:document-content>"#
        );
        let mut writer = litchi_odf_common::core::PackageWriter::new();
        writer.set_mimetype(MIME).expect("ODS MIME");
        writer
            .add_file("content.xml", content.as_bytes())
            .expect("ODS content");
        writer
            .add_file_with_media_type("Pictures/sample.bin", b"media", "application/octet-stream")
            .expect("ODS media");
        writer.finish_to_bytes().expect("ODS package")
    }

    #[test]
    fn source_projection_matches_owned_semantics_and_materializes_explicitly() {
        let bytes = package();
        let source =
            SourceBackedSpreadsheet::from_read_at(Arc::new(OwnedSource::new(bytes.clone())))
                .expect("source ODS");
        assert_eq!(source.sheet_names().unwrap(), ["Sheet1", "Sheet2"]);
        assert_eq!(source.sheet_count().unwrap(), 2);
        assert!(matches!(
            source.cell("Sheet1", 0, 0).unwrap(),
            Some(crate::worksheet::CellView::Stored(cell)) if cell.text == "hello"
        ));
        assert_eq!(source.text().unwrap(), "hello\tworld\nhello\tworld\n");
        assert_eq!(source.media_files().unwrap(), ["Pictures/sample.bin"]);
        assert_eq!(
            source.media_data("Pictures/sample.bin").unwrap(),
            Some(b"media".to_vec())
        );
        assert!(
            source
                .media_data("../outside.bin")
                .unwrap_err()
                .to_string()
                .contains("package href escapes the package root")
        );

        let owned = crate::Spreadsheet::from_bytes(bytes).expect("owned ODS");
        assert_eq!(owned.sheets().len(), source.sheet_count().unwrap());
        let materialized = source.materialize().expect("materialize ODS");
        assert_eq!(materialized.sheets(), owned.sheets());
    }

    #[test]
    fn repeated_row_projection_counts_emitted_separators_in_the_bound() {
        let mut sheet = Sheet::new("Sheet1").expect("sheet");
        let mut row = Row::repeated(crate::worksheet::validation::MAX_TEXT_BYTES / 2 + 1)
            .expect("repeated row");
        row.push_cell(Cell::new(CellValue::Text("x".to_owned()), "x"))
            .expect("cell");
        sheet.rows.push(row);
        assert!(matches!(
            project_text(&[sheet]),
            Err(litchi_core::Error::InvalidFormat(message))
                if message.contains("exceeds the safety limit")
        ));
    }

    #[test]
    fn source_checks_revision_and_is_send_sync() {
        fn assert_send_sync<T: Send + Sync>() {}
        assert_send_sync::<SourceBackedSpreadsheet>();

        let bytes = package();
        let source = Arc::new(MutableSource::new(bytes));
        let spreadsheet =
            SourceBackedSpreadsheet::from_read_at(source.clone()).expect("source ODS");
        source.bump();
        assert!(matches!(
            spreadsheet.sheet_count(),
            Err(litchi_core::Error::SourceChanged { .. })
        ));
        assert!(matches!(
            spreadsheet.text(),
            Err(litchi_core::Error::SourceChanged { .. })
        ));
        assert!(matches!(
            spreadsheet.member_data("Pictures/sample.bin"),
            Err(litchi_core::Error::SourceChanged { .. })
        ));
        assert!(matches!(
            spreadsheet.materialize(),
            Err(litchi_core::Error::SourceChanged { .. })
        ));
    }

    #[test]
    fn source_cell_locator_is_lazy_and_builds_at_threshold_without_source_reads() {
        let source = Arc::new(CountingSource::new(package()));
        let spreadsheet =
            SourceBackedSpreadsheet::from_read_at(source.clone()).expect("source ODS");
        let after_open = source.reads();

        assert!(spreadsheet.cell_locator.get().is_none());
        for _ in 0..(super::super::cell_locator::BUILD_QUERY_THRESHOLD - 1) {
            assert!(matches!(
                spreadsheet.cell("Sheet1", 0, 0),
                Ok(Some(crate::worksheet::CellView::Stored(_)))
            ));
        }
        assert!(spreadsheet.cell_locator.get().is_none());
        assert_eq!(source.reads(), after_open);

        assert!(matches!(
            spreadsheet.cell("Sheet1", 0, 0),
            Ok(Some(crate::worksheet::CellView::Stored(_)))
        ));
        assert!(matches!(spreadsheet.cell_locator.get(), Some(Some(_))));
        assert_eq!(source.reads(), after_open);
    }

    #[test]
    fn source_cell_batch_matches_scalar_and_observes_version_once_at_each_boundary() {
        let source = Arc::new(CountingSource::new(package()));
        let spreadsheet =
            SourceBackedSpreadsheet::from_read_at(source.clone()).expect("source ODS");
        let owned = crate::Spreadsheet::from_bytes(package()).expect("owned ODS");
        let after_open_reads = source.reads();
        let after_open_versions = source.versions();
        let selectors = [
            CellSelector::new("Sheet1", 1, 1),
            CellSelector::new("Sheet1", 0, 0),
            CellSelector::new("Sheet1", 2, 0),
            CellSelector::new("Sheet2", 0, 0),
            CellSelector::new("Missing", 0, 0),
            CellSelector::new("Sheet1", 0, 0),
            CellSelector::new("Sheet1", usize::MAX, usize::MAX),
        ];

        let batch = spreadsheet
            .cell_batch(&selectors)
            .expect("source cell batch");
        let expected = selectors
            .iter()
            .map(|selector| owned.cell(selector.sheet_name(), selector.row(), selector.column()))
            .collect::<Vec<_>>();
        assert_eq!(batch, expected);
        assert_eq!(source.versions() - after_open_versions, 2);
        assert_eq!(source.reads(), after_open_reads);
    }

    #[test]
    fn source_cell_batch_empty_and_selector_bound_are_atomic_and_read_free() {
        let source = Arc::new(CountingSource::new(package()));
        let spreadsheet =
            SourceBackedSpreadsheet::from_read_at(source.clone()).expect("source ODS");
        let after_open_reads = source.reads();
        assert!(
            spreadsheet
                .cell_batch(&[])
                .expect("empty source cell batch")
                .is_empty()
        );
        assert_eq!(source.reads(), after_open_reads);

        let exact = (0..MAX_CELL_SELECTORS)
            .map(|index| match index % 4 {
                0 => CellSelector::new("Sheet1", 0, 0),
                1 => CellSelector::new("Sheet1", 0, 2),
                2 => CellSelector::new("Sheet2", 0, 0),
                _ => CellSelector::new("Missing", 0, 0),
            })
            .collect::<Vec<_>>();
        let before_exact_reads = source.reads();
        let before_exact_versions = source.versions();
        let values = spreadsheet
            .cell_batch(&exact)
            .expect("exact selector bound should succeed");
        assert_eq!(values.len(), MAX_CELL_SELECTORS);
        for (index, value) in values.iter().enumerate() {
            match index % 4 {
                0 => assert!(matches!(
                    value,
                    Some(crate::worksheet::CellView::Stored(cell)) if cell.text == "hello"
                )),
                1 | 2 => assert!(matches!(value, Some(crate::worksheet::CellView::Missing))),
                _ => assert_eq!(*value, None),
            }
        }
        assert_eq!(source.versions() - before_exact_versions, 2);
        assert_eq!(source.reads(), before_exact_reads);
        assert!(matches!(spreadsheet.cell_locator.get(), Some(Some(_))));

        let bounded_source = Arc::new(CountingSource::new(package()));
        let bounded =
            SourceBackedSpreadsheet::from_read_at(bounded_source.clone()).expect("source ODS");
        let bounded_reads = bounded_source.reads();
        let selectors = vec![CellSelector::new("Sheet1", 0, 0); MAX_CELL_SELECTORS + 1];
        let error = bounded.cell_batch(&selectors).expect_err("bound must fail");
        assert!(matches!(
            error,
            litchi_core::Error::InvalidFormat(message)
                if message.contains("selector safety limit")
        ));
        assert_eq!(bounded_source.reads(), bounded_reads);
        assert!(bounded.cell_locator.get().is_none());
    }

    #[test]
    fn source_cell_batch_prefers_stale_before_and_during_lookup() {
        let source = Arc::new(MutableSource::new(package()));
        let spreadsheet =
            SourceBackedSpreadsheet::from_read_at(source.clone()).expect("source ODS");
        let selectors = [CellSelector::new("Sheet1", 0, 0)];

        source.bump();
        assert!(matches!(
            spreadsheet.cell_batch(&selectors),
            Err(litchi_core::Error::SourceChanged { .. })
        ));

        let source = Arc::new(MutableSource::new(package()));
        let spreadsheet =
            SourceBackedSpreadsheet::from_read_at(source.clone()).expect("source ODS");
        source.flip_after_next_version();
        assert!(matches!(
            spreadsheet.cell_batch(&selectors),
            Err(litchi_core::Error::SourceChanged { .. })
        ));
    }

    #[test]
    fn source_cell_batch_concurrent_lookup_preserves_pointer_identity() {
        fn assert_send_sync<T: Send + Sync>() {}
        assert_send_sync::<SourceBackedSpreadsheet>();

        let source = Arc::new(OwnedSource::new(package()));
        let spreadsheet =
            Arc::new(SourceBackedSpreadsheet::from_read_at(source).expect("source ODS"));
        let expected = match spreadsheet.sheets[0].cell_view(0, 0) {
            crate::worksheet::CellView::Stored(cell) => ptr::from_ref(cell) as usize,
            crate::worksheet::CellView::Missing => panic!("fixture cell"),
        };
        let selector = CellSelector::new("Sheet1", 0, 0);
        let threads = (0..8)
            .map(|_| {
                let spreadsheet = Arc::clone(&spreadsheet);
                std::thread::spawn(move || {
                    for _ in 0..super::super::cell_locator::BUILD_QUERY_THRESHOLD {
                        let result = spreadsheet
                            .cell_batch(&[selector])
                            .expect("source cell batch");
                        let Some(crate::worksheet::CellView::Stored(cell)) =
                            result.first().copied().flatten()
                        else {
                            panic!("fixture cell");
                        };
                        assert_eq!(ptr::from_ref(cell) as usize, expected);
                    }
                })
            })
            .collect::<Vec<_>>();
        for thread in threads {
            thread.join().expect("cell batch thread");
        }
        assert!(matches!(spreadsheet.cell_locator.get(), Some(Some(_))));
    }

    #[test]
    fn source_indexed_lookup_matches_linear_runs_for_repeated_empty_and_boundary_rows() {
        let spreadsheet =
            SourceBackedSpreadsheet::from_read_at(Arc::new(OwnedSource::new(package())))
                .expect("source ODS");

        for _ in 0..super::super::cell_locator::BUILD_QUERY_THRESHOLD {
            assert!(spreadsheet.cell("Sheet1", 0, 0).unwrap().is_some());
        }
        assert!(matches!(spreadsheet.cell_locator.get(), Some(Some(_))));

        for (sheet_name, row, column) in [
            ("Sheet1", 0, 0),
            ("Sheet1", 1, 1),
            ("Sheet1", 2, 0),
            ("Sheet1", 0, 2),
            ("Sheet1", usize::MAX, usize::MAX),
            ("Sheet2", 0, 0),
        ] {
            let sheet_index = spreadsheet
                .sheets
                .iter()
                .position(|sheet| sheet.name == sheet_name)
                .expect("fixture sheet");
            let expected = spreadsheet.sheets[sheet_index].cell_view(row, column);
            let actual = spreadsheet
                .cell(sheet_name, row, column)
                .expect("source cell")
                .expect("fixture sheet");
            match (expected, actual) {
                (crate::worksheet::CellView::Missing, crate::worksheet::CellView::Missing) => {},
                (
                    crate::worksheet::CellView::Stored(expected),
                    crate::worksheet::CellView::Stored(actual),
                ) => assert!(ptr::eq(expected, actual)),
                _ => panic!("indexed and direct cell views differ"),
            }
        }

        assert_eq!(spreadsheet.cell("missing", 0, 0).unwrap(), None);
    }

    #[test]
    fn source_cell_locator_zero_budget_falls_back_permanently_to_linear_lookup() {
        let spreadsheet =
            SourceBackedSpreadsheet::from_read_at(Arc::new(OwnedSource::new(package())))
                .expect("source ODS");
        let locator =
            super::super::cell_locator::CellLocator::try_build_with_budget(&spreadsheet.sheets, 0);
        assert!(locator.is_none());
        assert!(spreadsheet.cell_locator.set(locator).is_ok());

        for _ in 0..(super::super::cell_locator::BUILD_QUERY_THRESHOLD * 2) {
            assert!(matches!(
                spreadsheet.cell("Sheet1", 0, 0),
                Ok(Some(crate::worksheet::CellView::Stored(_)))
            ));
        }
        assert!(matches!(spreadsheet.cell_locator.get(), Some(None)));
        assert_eq!(spreadsheet.cell_queries.load(Ordering::Relaxed), 0);
    }

    #[test]
    fn source_concurrent_first_cell_locator_build_is_shared_and_preserves_identity() {
        fn assert_send_sync<T: Send + Sync>() {}
        assert_send_sync::<SourceBackedSpreadsheet>();

        let spreadsheet = Arc::new(
            SourceBackedSpreadsheet::from_read_at(Arc::new(OwnedSource::new(package())))
                .expect("source ODS"),
        );
        let expected = match spreadsheet.sheets[0].cell_view(0, 0) {
            crate::worksheet::CellView::Stored(cell) => ptr::from_ref(cell) as usize,
            crate::worksheet::CellView::Missing => panic!("fixture cell"),
        };

        let threads = (0..8)
            .map(|_| {
                let spreadsheet = Arc::clone(&spreadsheet);
                std::thread::spawn(move || {
                    for _ in 0..super::super::cell_locator::BUILD_QUERY_THRESHOLD {
                        let Ok(Some(crate::worksheet::CellView::Stored(cell))) =
                            spreadsheet.cell("Sheet1", 0, 0)
                        else {
                            panic!("fixture cell");
                        };
                        assert_eq!(ptr::from_ref(cell) as usize, expected);
                    }
                })
            })
            .collect::<Vec<_>>();
        for thread in threads {
            thread.join().expect("cell lookup thread");
        }
        assert!(matches!(spreadsheet.cell_locator.get(), Some(Some(_))));
    }

    #[test]
    fn source_indexed_cell_lookup_rejects_stale_source() {
        let source = Arc::new(MutableSource::new(package()));
        let spreadsheet =
            SourceBackedSpreadsheet::from_read_at(source.clone()).expect("source ODS");
        for _ in 0..super::super::cell_locator::BUILD_QUERY_THRESHOLD {
            assert!(spreadsheet.cell("Sheet1", 0, 0).unwrap().is_some());
        }
        assert!(matches!(spreadsheet.cell_locator.get(), Some(Some(_))));

        source.bump();
        assert!(matches!(
            spreadsheet.cell("Sheet1", 0, 0),
            Err(litchi_core::Error::SourceChanged { .. })
        ));
    }

    #[test]
    fn selected_media_read_is_deferred_to_member_access() {
        let source = Arc::new(CountingSource::new(package()));
        let spreadsheet =
            SourceBackedSpreadsheet::from_read_at(source.clone()).expect("source ODS");
        let after_open = source.reads();
        let range_start = source.range_count();
        assert_eq!(spreadsheet.sheet_count().unwrap(), 2);
        assert_eq!(source.reads(), after_open);
        assert_eq!(
            spreadsheet.media_data("Pictures/sample.bin").unwrap(),
            Some(b"media".to_vec())
        );
        assert!(source.reads() > after_open);
        let media_ranges = source.ranges_since(range_start);
        assert!(!media_ranges.is_empty());
        assert!(media_ranges.iter().any(|(offset, length)| {
            *offset > 0 && *length > 0 && *length < source.bytes.len().unwrap() as usize
        }));
    }

    #[test]
    fn source_limits_and_passwords_fail_closed() {
        let bytes = package();
        let limits = ReadLimits::default().with_max_source_bytes(1);
        assert!(matches!(
            SourceBackedSpreadsheet::from_read_at_with_limits(
                Arc::new(OwnedSource::new(bytes.clone())),
                limits,
            ),
            Err(litchi_core::Error::ResourceLimit(_))
        ));

        let mut writer = litchi_odf_common::core::PackageWriter::new();
        writer.set_mimetype(MIME).expect("ODS MIME");
        writer
            .set_encryption(
                "source-password",
                litchi_odf_common::core::Profile::compatible(),
            )
            .expect("encryption profile");
        let content = format!(
            r#"<office:document-content xmlns:office="{OFFICE}" xmlns:table="{TABLE}" xmlns:text="{TEXT}"><office:body><office:spreadsheet><table:table table:name="Sheet1"/></office:spreadsheet></office:body></office:document-content>"#
        );
        writer
            .add_file("content.xml", content.as_bytes())
            .expect("encrypted content");
        let encrypted = writer.finish_to_bytes().expect("encrypted package");
        assert!(
            SourceBackedSpreadsheet::from_read_at(Arc::new(OwnedSource::new(encrypted.clone(),)))
                .is_err()
        );
        assert!(
            SourceBackedSpreadsheet::from_read_at_with_password(
                Arc::new(OwnedSource::new(encrypted.clone())),
                "wrong",
            )
            .is_err()
        );
        assert!(
            SourceBackedSpreadsheet::from_read_at_with_password(
                Arc::new(OwnedSource::new(encrypted)),
                "source-password",
            )
            .is_ok()
        );
    }

    #[test]
    fn owned_and_source_reject_manifest_aliases_before_password_use() {
        let mut writer = litchi_odf_common::core::PackageWriter::new();
        writer.set_mimetype(MIME).expect("ODS MIME");
        writer
            .set_encryption(
                "manifest-password",
                litchi_odf_common::core::Profile::compatible(),
            )
            .expect("encryption profile");
        let content = format!(
            r#"<office:document-content xmlns:office="{OFFICE}" xmlns:table="{TABLE}" xmlns:text="{TEXT}"><office:body><office:spreadsheet><table:table table:name="Sheet1"/></office:spreadsheet></office:body></office:document-content>"#
        );
        writer
            .add_file("content.xml", content.as_bytes())
            .expect("encrypted content");
        let canonical = writer.finish_to_bytes().expect("encrypted package");

        for alias in [
            "/content.xml",
            "./content.xml",
            "foo/../content.xml",
            "content%2Exml",
            "content.xml?cache=1",
            "content.xml#fragment",
            "C:content.xml",
            "foo:content.xml",
        ] {
            let aliased = rewrite_manifest_content_path(&canonical, alias);
            assert!(
                SourceBackedSpreadsheet::from_read_at(Arc::new(OwnedSource::new(aliased.clone())))
                    .is_err(),
                "source open must reject manifest alias {alias}"
            );
            assert!(
                SourceBackedSpreadsheet::from_read_at_with_password(
                    Arc::new(OwnedSource::new(aliased.clone())),
                    "wrong",
                )
                .is_err(),
                "source wrong-password open must reject manifest alias {alias}"
            );
            assert!(
                crate::Spreadsheet::from_bytes(aliased.clone()).is_err(),
                "owned open must reject manifest alias {alias}"
            );
            assert!(
                crate::Spreadsheet::from_bytes_with_password(aliased, "wrong").is_err(),
                "owned wrong-password open must reject manifest alias {alias}"
            );
        }

        assert!(
            SourceBackedSpreadsheet::from_read_at_with_password(
                Arc::new(OwnedSource::new(canonical.clone())),
                "manifest-password",
            )
            .is_ok()
        );
        assert!(
            crate::Spreadsheet::from_bytes_with_password(canonical, "manifest-password").is_ok()
        );
    }

    fn rewrite_manifest_content_path(bytes: &[u8], path: &str) -> Vec<u8> {
        let mut archive = ZipArchive::new(Cursor::new(bytes)).expect("encrypted ZIP");
        let mut output = Cursor::new(Vec::new());
        {
            let mut writer = ZipWriter::new(&mut output);
            let options =
                SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
            for index in 0..archive.len() {
                let (name, mut data) = {
                    let mut entry = archive.by_index(index).expect("ZIP entry");
                    let name = entry.name().to_owned();
                    let mut data = Vec::new();
                    entry.read_to_end(&mut data).expect("ZIP entry bytes");
                    (name, data)
                };
                if name == "META-INF/manifest.xml" {
                    let manifest = String::from_utf8(data).expect("manifest UTF-8");
                    let needle = r#"manifest:full-path="content.xml""#;
                    let replacement = format!(r#"manifest:full-path="{path}""#);
                    assert!(manifest.contains(needle));
                    data = manifest.replace(needle, &replacement).into_bytes();
                }
                writer.start_file(name, options).expect("ZIP entry");
                writer.write_all(&data).expect("ZIP entry bytes");
            }
            writer.finish().expect("rewritten ZIP");
        }
        output.into_inner()
    }

    struct CountingSource {
        bytes: OwnedSource,
        reads: AtomicUsize,
        versions: AtomicUsize,
        ranges: std::sync::Mutex<Vec<(u64, usize)>>,
    }

    impl CountingSource {
        fn new(bytes: Vec<u8>) -> Self {
            Self {
                bytes: OwnedSource::new(bytes),
                reads: AtomicUsize::new(0),
                versions: AtomicUsize::new(0),
                ranges: std::sync::Mutex::new(Vec::new()),
            }
        }

        fn reads(&self) -> usize {
            self.reads.load(Ordering::Relaxed)
        }

        fn versions(&self) -> usize {
            self.versions.load(Ordering::Relaxed)
        }

        fn range_count(&self) -> usize {
            self.ranges.lock().expect("range lock").len()
        }

        fn ranges_since(&self, start: usize) -> Vec<(u64, usize)> {
            self.ranges
                .lock()
                .expect("range lock")
                .get(start..)
                .unwrap_or_default()
                .to_vec()
        }
    }

    impl ReadAt for CountingSource {
        fn len(&self) -> std::io::Result<u64> {
            self.bytes.len()
        }

        fn read_at(&self, offset: u64, output: &mut [u8]) -> std::io::Result<usize> {
            self.reads.fetch_add(1, Ordering::Relaxed);
            let read = self.bytes.read_at(offset, output)?;
            self.ranges
                .lock()
                .map_err(|_| std::io::Error::other("range lock poisoned"))?
                .push((offset, read));
            Ok(read)
        }

        fn version(&self) -> std::io::Result<SourceVersion> {
            self.versions.fetch_add(1, Ordering::Relaxed);
            self.bytes.version()
        }
    }

    struct MutableSource {
        bytes: Arc<Vec<u8>>,
        revision: AtomicU64,
        flip_after_version: std::sync::atomic::AtomicBool,
    }

    impl MutableSource {
        fn new(bytes: Vec<u8>) -> Self {
            Self {
                bytes: Arc::new(bytes),
                revision: AtomicU64::new(0),
                flip_after_version: std::sync::atomic::AtomicBool::new(false),
            }
        }

        fn bump(&self) {
            self.revision.fetch_add(1, Ordering::Relaxed);
        }

        fn flip_after_next_version(&self) {
            self.flip_after_version.store(true, Ordering::Relaxed);
        }
    }

    impl ReadAt for MutableSource {
        fn len(&self) -> std::io::Result<u64> {
            u64::try_from(self.bytes.len())
                .map_err(|_| std::io::Error::other("source length does not fit u64"))
        }

        fn read_at(&self, offset: u64, output: &mut [u8]) -> std::io::Result<usize> {
            let start = usize::try_from(offset).unwrap_or(usize::MAX);
            let Some(input) = self.bytes.get(start..) else {
                return Ok(0);
            };
            let count = input.len().min(output.len());
            output[..count].copy_from_slice(&input[..count]);
            Ok(count)
        }

        fn version(&self) -> std::io::Result<SourceVersion> {
            let revision = self.revision.load(Ordering::Relaxed);
            let value = SourceVersion::new(0x4f44, revision);
            if self.flip_after_version.swap(false, Ordering::Relaxed) {
                self.revision.fetch_add(1, Ordering::Relaxed);
            }
            Ok(value)
        }
    }
}
