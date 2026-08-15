//! Immutable, positional-source ODS reads.
//!
//! [`SourceBackedSpreadsheet`] is an additive read-only owner.  The existing
//! [`super::Spreadsheet`] remains the complete owned/mutable facade; callers
//! explicitly call [`SourceBackedSpreadsheet::materialize`] before entering
//! that mutation boundary.  The ZIP index and the validated semantic XML are
//! retained, while unrelated package members stay cold until selected.

use std::{fmt, sync::Arc};

#[cfg(any(unix, windows))]
use litchi_core::FileSource;
use litchi_core::{Error, Metadata, ReadAt, Result, SourceVersion};
use litchi_odf_common::{
    core::{SourceBackedPackage, SourcePackageLimits, validate_content_part},
    package::{is_media_path, resolve_package_path},
};
#[cfg(any(unix, windows))]
use std::path::Path;
use zeroize::Zeroizing;

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

/// Read-only ODS access over an immutable positional source.
///
/// Opening validates the ZIP archive, MIME member, manifest, content root,
/// and worksheet graph, then retains only the required XML projections.
/// Package members such as pictures and embedded objects are read only when
/// selected through [`Self::member_data`] or [`Self::media_data`].  Every
/// public operation checks the captured [`SourceVersion`] and reports
/// [`Error::SourceChanged`] when the source no longer identifies the same
/// snapshot.
pub struct SourceBackedSpreadsheet {
    package: SourceBackedPackage,
    source: Arc<dyn ReadAt>,
    source_version: SourceVersion,
    content_xml: String,
    styles_xml: Option<String>,
    definitions: Vec<Definition>,
    sheets: Vec<Sheet>,
    metadata: crate::metadata::Snapshot,
    settings: Option<Settings>,
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

            let content_xml =
                String::from_utf8(package.get_file("content.xml")?).map_err(|error| {
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
                styles_xml,
                definitions,
                sheets,
                metadata,
                settings,
            ))
        })();
        let (content_xml, styles_xml, definitions, sheets, metadata, settings) =
            prefer_current(source.as_ref(), source_version, parsed)?;

        Ok(Self {
            package,
            source,
            source_version,
            content_xml,
            styles_xml,
            definitions,
            sheets,
            metadata,
            settings,
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
        let value = self.content_xml.as_str();
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
        let value = self.sheets.as_slice();
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
        let value = self
            .sheets
            .iter()
            .find(|sheet| sheet.name == sheet_name)
            .map(|sheet| sheet.cell_view(row, column));
        self.check_source()?;
        Ok(value)
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
    use litchi_core::{OwnedSource, ReadAt, SourceVersion};
    use std::io::{Cursor, Read, Write};
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
        ranges: std::sync::Mutex<Vec<(u64, usize)>>,
    }

    impl CountingSource {
        fn new(bytes: Vec<u8>) -> Self {
            Self {
                bytes: OwnedSource::new(bytes),
                reads: AtomicUsize::new(0),
                ranges: std::sync::Mutex::new(Vec::new()),
            }
        }

        fn reads(&self) -> usize {
            self.reads.load(Ordering::Relaxed)
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
            self.bytes.version()
        }
    }

    struct MutableSource {
        bytes: Arc<Vec<u8>>,
        revision: AtomicU64,
    }

    impl MutableSource {
        fn new(bytes: Vec<u8>) -> Self {
            Self {
                bytes: Arc::new(bytes),
                revision: AtomicU64::new(0),
            }
        }

        fn bump(&self) {
            self.revision.fetch_add(1, Ordering::Relaxed);
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
            Ok(SourceVersion::new(
                0x4f44,
                self.revision.load(Ordering::Relaxed),
            ))
        }
    }
}
