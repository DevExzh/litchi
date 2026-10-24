//! Cell style references and readable formatting on the source-backed owner.

use crate::style::StyleKey;
use std::sync::Arc;

use litchi_sheet::{Area, At, Cell as Address};

use super::{SourceBackedWorkbook, SourceCell, SourceInner, SourceWorksheet};
use crate::cell::{Cell, Stored};
use crate::error::{Result, allocation};
use crate::style::cell_format::{Format, FormatKey, FormatTable};

/// One stored cell record of a source-backed worksheet, with the local
/// shared-format reference of the record.
///
/// This is [`SourceCell`] plus the record's `c/@s` attribute. It is an owned
/// value and may outlive the worksheet handle that produced it.
#[derive(Clone, PartialEq, Eq)]
#[non_exhaustive]
pub struct Record {
    /// Checked worksheet coordinate of the stored record.
    pub address: Address,
    /// Owned semantic cell state.
    pub cell: Cell,
    /// The record's local shared format. `None` means the record has no
    /// local style, so the workbook's base format applies; pass the value to
    /// [`SourceWorksheet::format`] either way.
    pub style: Option<StyleKey>,
}

impl std::fmt::Debug for Record {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("Record")
            .field("address", &self.address)
            .field("has_style", &self.style.is_some())
            .finish_non_exhaustive()
    }
}

impl Record {
    fn from_stored(entry: &Stored, lineage: &Arc<crate::style::StyleLineage>) -> Self {
        Self {
            address: entry.address,
            cell: entry.cell.clone(),
            style: entry
                .style
                .map(|raw| StyleKey::new(raw, Arc::clone(lineage))),
        }
    }
}

impl From<Record> for SourceCell {
    fn from(record: Record) -> Self {
        Self {
            address: record.address,
            cell: record.cell,
        }
    }
}

impl SourceBackedWorkbook {
    /// Readable local formatting of a shared cell format.
    ///
    /// `style` is a record's [`Record::style`]; `None` selects the
    /// base format, the one a cell without a local style uses. The result is
    /// `None` when the workbook has no styles part.
    ///
    /// The styles payload is parsed once per source revision into a bounded,
    /// evictable cache owned by this workbook; later calls are table lookups.
    ///
    /// # Errors
    ///
    /// Returns an error when the styles part cannot be read, a property value
    /// lies outside its schema domain, the key is outside this workbook's
    /// format table, belongs to another workbook, or the source changed or the
    /// read was cancelled.
    pub fn format(&self, style: Option<&StyleKey>) -> Result<Option<Format>> {
        self.view()?.inner.format(style)
    }

    /// Release the parsed format table retained for [`Self::format`], unless
    /// an active [`Format`] still uses it, and return whether a table was
    /// released.
    ///
    /// The table is a clean cached value: releasing it changes no result, and
    /// the next formatting read parses the styles part again.
    pub fn release_cached_formats(&self) -> bool {
        self.current().formats.trim_clean()
    }
}

impl SourceWorksheet {
    /// Read the opaque local style key of one stored cell without cloning its
    /// semantic cell value.
    ///
    /// `None` means that no cell record is stored at the checked coordinate or
    /// that the record has no local style and therefore uses the workbook's
    /// base format. An explicit style key for the base format (index zero) is
    /// retained, so callers can distinguish it from an unstyled record when
    /// they need the leaf's exact local-style state.
    ///
    /// # Errors
    ///
    /// Returns a coordinate error for invalid A1 or out-of-domain coordinates,
    /// a non-cell-sheet error, a worksheet parse error, or a source/execution
    /// fence error.
    pub fn style_key<'a>(&self, at: impl Into<At<'a>>) -> Result<Option<StyleKey>> {
        let address = match at.into().resolve() {
            Ok(address) => address,
            Err(error) => return self.finish_result(Err(error.into())),
        };
        let key = self.store().map(|store| {
            store.entry(address).and_then(|entry| {
                entry
                    .style
                    .map(|raw| StyleKey::new(raw, Arc::clone(&self.owner.style_lineage)))
            })
        });
        self.finish_result(key)
    }

    /// Readable local formatting of a shared cell format of this sheet's
    /// workbook; see [`SourceBackedWorkbook::format`].
    ///
    /// ```no_run
    /// use litchi_xlsx::SourceBackedWorkbook;
    ///
    /// let workbook = SourceBackedWorkbook::open("report.xlsx")?;
    /// let sheet = workbook.sheet("Summary")?.expect("sheet");
    /// for record in sheet.cell_records("A1:D10")? {
    ///     let format = sheet.format(record.style.as_ref())?;
    ///     let bold = format
    ///         .as_ref()
    ///         .and_then(|format| format.font())
    ///         .and_then(|font| font.bold());
    ///     println!("{}: {bold:?}", record.address);
    /// }
    /// # Ok::<(), litchi_xlsx::Error>(())
    /// ```
    ///
    /// # Errors
    ///
    /// As for [`SourceBackedWorkbook::format`].
    pub fn format(&self, style: Option<&StyleKey>) -> Result<Option<Format>> {
        self.owner.format(style)
    }

    /// The stored record at one coordinate, with its local style.
    ///
    /// `None` means no `<c>` record is stored there. Unlike [`Self::cell`],
    /// the record of a merged follower coordinate is returned as stored.
    /// The worksheet is materialized and retained on first use, so later
    /// reads are lookups in the parsed store.
    ///
    /// # Errors
    ///
    /// Returns an error for a coordinate outside the grid, a non-cell sheet,
    /// a malformed worksheet, or a changed source or cancelled read.
    pub fn cell_record<'a>(&self, at: impl Into<At<'a>>) -> Result<Option<Record>> {
        let address = match at.into().resolve() {
            Ok(address) => address,
            Err(error) => return self.finish_result(Err(error.into())),
        };
        let record = self.store().map(|store| {
            store
                .entry(address)
                .map(|entry| Record::from_stored(entry, &self.owner.style_lineage))
        });
        self.finish_result(record)
    }

    /// Every stored record inside a checked range, with their local styles,
    /// in row-major order.
    ///
    /// This is [`Self::cells`] with each record's style. It reads the
    /// materialized worksheet, which is parsed and retained on first use.
    ///
    /// # Errors
    ///
    /// As for [`Self::cell_record`], plus an allocation failure for the
    /// result.
    pub fn cell_records<'a>(&self, area: impl Into<Area<'a>>) -> Result<Vec<Record>> {
        let range = match area.into().resolve() {
            Ok(range) => range,
            Err(error) => return self.finish_result(Err(error.into())),
        };
        let records = self.store().and_then(|store| {
            let count = store.entries_in(range).count();
            let mut records = Vec::new();
            records
                .try_reserve_exact(count)
                .map_err(|source| allocation("source-backed cell records", source))?;
            records.extend(
                store
                    .entries_in(range)
                    .map(|entry| Record::from_stored(entry, &self.owner.style_lineage)),
            );
            Ok(records)
        });
        self.finish_result(records)
    }
}

impl SourceInner {
    fn format(&self, style: Option<&StyleKey>) -> Result<Option<Format>> {
        self.execution_check()?;
        self.package.source_version()?;
        if style.is_some_and(|key| !key.belongs_to(&self.style_lineage)) {
            return Err(crate::Error::ForeignStyle);
        }
        let Some(table) = self.format_table()? else {
            return Ok(None);
        };
        Format::new(table, style.map_or(0, StyleKey::raw)).map(Some)
    }

    /// The parsed format table of this workbook's styles part, through the
    /// workbook's cache. `None` when the workbook has no styles part.
    fn format_table(&self) -> Result<Option<Arc<FormatTable>>> {
        self.execution_check()?;
        let Some(uri) = self.styles_uri.as_ref() else {
            return Ok(None);
        };
        // Resolving the part checks the captured source version and the
        // managed package's execution policy even when the table is cached.
        let part = self.package.part(uri)?;
        let version = self.package.source_version()?;
        let limits = self.package.read_limits();
        let table = self.formats.get_or_build(FormatKey::source(version), || {
            let data = part.data()?;
            let table = FormatTable::parse_with_limits(data.as_bytes(), limits)?;
            // `get_or_build` publishes the successful callback result before
            // returning. Fence the value while the cache still owns only the
            // local result, so cancellation or source replacement cannot
            // publish a stale format table.
            self.execution_check()?;
            self.package.source_version()?;
            Ok(table)
        })?;
        self.execution_check()?;
        self.package.source_version()?;
        Ok(Some(table))
    }
}

#[cfg(test)]
mod tests {
    #![allow(clippy::unwrap_used, reason = "focused source-owner fixtures")]

    use std::io;
    use std::sync::Arc;
    use std::sync::atomic::{AtomicU64, Ordering};

    use litchi_core::{ReadAt, SourceVersion};
    use litchi_opc::constants::content_type as ct;
    use litchi_opc::{OpcError, SourceLimitKind, SourceLimits};
    use soapberry_zip::office::StreamingArchiveWriter;

    use super::{Record, StyleKey};
    use crate::cell::{Cell, Value};
    use crate::{Error, SourceBackedWorkbook};

    const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    fn workbook_bytes(styles: Option<&str>) -> Vec<u8> {
        let mut writer =
            StreamingArchiveWriter::new(soapberry_zip::office::StreamingArchiveLimits::default());
        let styles_override = styles.map_or(String::new(), |_| {
            format!(
                r#"<Override PartName="/xl/styles.xml" ContentType="{}"/>"#,
                ct::SML_STYLES
            )
        });
        writer
            .write_stored(
                "[Content_Types].xml",
                format!(
                    r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="{}"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="{}"/>{styles_override}</Types>"#,
                    ct::SML_SHEET_MAIN,
                    ct::SML_WORKSHEET,
                )
                .as_bytes(),
            )
            .unwrap();
        writer
            .write_stored(
                "_rels/.rels",
                format!(
                    r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>"#
                )
                .as_bytes(),
            )
            .unwrap();
        writer
            .write_stored(
                "xl/workbook.xml",
                format!(
                    r#"<workbook xmlns="{SML}" xmlns:r="{REL}"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>"#
                )
                .as_bytes(),
            )
            .unwrap();
        let styles_relationship = styles.map_or(String::new(), |_| {
            format!(r#"<Relationship Id="rId2" Type="{REL}/styles" Target="styles.xml"/>"#)
        });
        writer
            .write_stored(
                "xl/_rels/workbook.xml.rels",
                format!(
                    r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{REL}/worksheet" Target="worksheets/sheet1.xml"/>{styles_relationship}</Relationships>"#
                )
                .as_bytes(),
            )
            .unwrap();
        let styled = if styles.is_some() { r#" s="1""# } else { "" };
        writer
            .write_stored(
                "xl/worksheets/sheet1.xml",
                format!(
                    r#"<worksheet xmlns="{SML}"><sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"{styled}><v>2</v></c></row></sheetData></worksheet>"#
                )
                .as_bytes(),
            )
            .unwrap();
        if let Some(styles) = styles {
            writer
                .write_stored("xl/styles.xml", styles.as_bytes())
                .unwrap();
        }
        writer.finish_to_bytes().unwrap()
    }

    fn styles() -> String {
        format!(
            r#"<styleSheet xmlns="{SML}"><fonts><font><sz val="11"/></font><font><b/><sz val="9"/></font></fonts><cellXfs count="2"><xf/><xf fontId="1"/></cellXfs></styleSheet>"#
        )
    }

    struct Source {
        bytes: Vec<u8>,
        revision: AtomicU64,
    }

    impl ReadAt for Source {
        fn len(&self) -> io::Result<u64> {
            Ok(self.bytes.len() as u64)
        }

        fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
            let offset = usize::try_from(offset)
                .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "offset"))?;
            let available = self.bytes.get(offset..).unwrap_or_default();
            let count = output.len().min(available.len());
            output[..count].copy_from_slice(&available[..count]);
            Ok(count)
        }

        fn version(&self) -> io::Result<SourceVersion> {
            Ok(SourceVersion::new(
                4_242,
                self.revision.load(Ordering::SeqCst),
            ))
        }
    }

    fn open(bytes: Vec<u8>) -> (Arc<Source>, SourceBackedWorkbook) {
        open_with_limits(bytes, SourceLimits::default())
    }

    fn open_with_limits(
        bytes: Vec<u8>,
        limits: SourceLimits,
    ) -> (Arc<Source>, SourceBackedWorkbook) {
        let source = Arc::new(Source {
            bytes,
            revision: AtomicU64::new(0),
        });
        let workbook = SourceBackedWorkbook::from_read_at_with_options(
            source.clone(),
            crate::OpenOptions::new().source_limits(limits),
        )
        .unwrap();
        (source, workbook)
    }

    fn assert_depth_limit(error: Error) {
        assert!(matches!(
            error,
            Error::Package(OpcError::Limit {
                kind: SourceLimitKind::Depth,
                maximum: 3,
                ..
            })
        ));
    }

    #[test]
    fn records_carry_local_styles_and_resolve_formats_once() {
        let (_source, workbook) = open(workbook_bytes(Some(&styles())));
        let sheet = workbook.sheet("Data").unwrap().unwrap();
        let records = sheet.cell_records("A1:B1").unwrap();
        assert_eq!(
            sheet.style_key("A1").unwrap().as_ref(),
            records[0].style.as_ref()
        );
        assert_eq!(
            sheet.style_key("B1").unwrap().as_ref(),
            records[1].style.as_ref()
        );
        assert_eq!(
            records,
            [
                Record {
                    address: litchi_sheet::Cell::at(0, 0).unwrap(),
                    cell: Cell::Value(Value::Number(crate::Number::new("1").unwrap())),
                    style: None,
                },
                Record {
                    address: litchi_sheet::Cell::at(0, 1).unwrap(),
                    cell: Cell::Value(Value::Number(crate::Number::new("2").unwrap())),
                    style: Some(StyleKey::new(1, Arc::clone(&workbook.inner.style_lineage))),
                },
            ]
        );
        let base = sheet.format(records[0].style.as_ref()).unwrap().unwrap();
        assert!(base.is_base());
        assert_eq!(base.font().unwrap().bold(), None);
        let bold = sheet.format(records[1].style.as_ref()).unwrap().unwrap();
        assert_eq!(bold.font().unwrap().bold(), Some(true));
        assert_eq!(bold.font().unwrap().size(), Some(9.0));

        let diagnostics = workbook.inner.formats.diagnostics();
        assert_eq!((diagnostics.builds, diagnostics.hits), (1, 1));
        assert!(diagnostics.pinned);

        let outside = sheet.format(Some(&StyleKey::new(
            2,
            Arc::clone(&workbook.inner.style_lineage),
        )));
        assert!(matches!(outside, Err(Error::Invalid(_))), "{outside:?}");
    }

    #[test]
    fn style_keys_are_bound_to_the_owner_lineage() {
        let (_, first) = open(workbook_bytes(Some(&styles())));
        let (_, second) = open(workbook_bytes(Some(&styles())));
        let direct_key = first
            .sheet("Data")
            .unwrap()
            .unwrap()
            .style_key("B1")
            .unwrap();
        assert_eq!(direct_key.as_ref().map(StyleKey::raw), Some(1));
        let record = first
            .sheet("Data")
            .unwrap()
            .unwrap()
            .cell_record("B1")
            .unwrap()
            .unwrap();
        assert!(
            first
                .clone()
                .format(record.style.as_ref())
                .unwrap()
                .is_some()
        );
        assert!(matches!(
            second.format(record.style.as_ref()),
            Err(Error::ForeignStyle)
        ));
        assert!(matches!(
            second.format(direct_key.as_ref()),
            Err(Error::ForeignStyle)
        ));
        let debug = format!("{record:?}");
        assert!(!debug.contains("Number"), "{debug}");
        assert!(!debug.contains("raw"), "{debug}");
    }

    #[test]
    fn a_workbook_without_styles_has_no_formats() {
        let (_source, workbook) = open(workbook_bytes(None));
        let sheet = workbook.sheet("Data").unwrap().unwrap();
        assert!(workbook.format(None).unwrap().is_none());
        let record = sheet.cell_record("B1").unwrap().unwrap();
        assert_eq!(record.style, None);
        assert!(sheet.format(record.style.as_ref()).unwrap().is_none());
        assert!(sheet.cell_record("C1").unwrap().is_none());
    }

    #[test]
    fn source_changes_refuse_cached_formats_and_records() {
        let (source, workbook) = open(workbook_bytes(Some(&styles())));
        let sheet = workbook.sheet("Data").unwrap().unwrap();
        assert!(workbook.format(None).unwrap().is_some());
        assert!(sheet.cell_record("B1").unwrap().is_some());
        source.revision.store(1, Ordering::SeqCst);
        assert!(matches!(workbook.format(None), Err(Error::Package(_))));
        assert!(matches!(sheet.cell_record("B1"), Err(Error::Package(_))));
        assert!(matches!(sheet.style_key("B1"), Err(Error::Package(_))));
        assert!(matches!(
            sheet.cell_records("A1:B1"),
            Err(Error::Package(_))
        ));
    }

    #[test]
    fn invalid_coordinates_are_typed_errors() {
        let (_source, workbook) = open(workbook_bytes(Some(&styles())));
        let sheet = workbook.sheet("Data").unwrap().unwrap();
        assert!(matches!(sheet.cell_record("A0"), Err(Error::Coordinate(_))));
        assert!(matches!(sheet.style_key("A0"), Err(Error::Coordinate(_))));
        assert!(matches!(sheet.cell_records("B2:A1"), Err(Error::Range(_))));
    }

    #[test]
    fn caller_depth_limits_reach_eager_and_source_format_reads() {
        let bytes = workbook_bytes(Some(&styles()));
        let limits = SourceLimits::builder()
            .max_depth(3)
            .unwrap()
            .build()
            .unwrap();

        let eager = crate::Package::from_bytes_with_options(
            &bytes,
            crate::OpenOptions::new().source_limits(limits),
        )
        .unwrap();
        let eager_style = eager.styles().unwrap().get(1).unwrap();
        assert_depth_limit(eager_style.format().unwrap_err());

        let (_source, source) = open_with_limits(bytes, limits);
        let source_style = StyleKey::new(1, Arc::clone(&source.inner.style_lineage));
        assert_depth_limit(source.format(Some(&source_style)).unwrap_err());
    }
}
