//! Immutable XLSX reads backed by a caller-provided positional source.
//!
//! This facade intentionally does not adapt into [`super::Workbook`]: that
//! snapshot owns a mutable OPC graph, while this type must keep ordinary part
//! payloads deferred. It exposes only semantic catalog and worksheet reads.

use std::collections::HashMap;
use std::sync::{Arc, OnceLock};

use litchi_core::{ReadAt, Selector as CoreSelector};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{PackURI, PartView, ReadLimits, SourceBackedPackage};
use litchi_sheet::{Area, At, Cell as Address, Rect};

use super::{DateSystem, Flavor, Selector, Visibility, WorksheetKind, codec};
use crate::cell::{Cell, Store, Text, View};
use crate::error::{Error, Result, allocation, invalid};
use crate::raw;

const CHARTSHEET_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet";
const STRICT_CHARTSHEET_REL: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/chartsheet";
const DIALOGSHEET_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/dialogsheet";
const MACROSHEET_REL: &str = "http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet";
const INTL_MACROSHEET_REL: &str =
    "http://schemas.microsoft.com/office/2006/relationships/xlIntlMacrosheet";
const CHARTSHEET_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml";

struct SourceSheetData {
    position: usize,
    name: String,
    name_key: Box<str>,
    kind: WorksheetKind,
    visibility: Visibility,
    part_uri: PackURI,
    cells: OnceLock<Store>,
}

struct SourceInner {
    package: SourceBackedPackage,
    shared_strings_uri: Option<PackURI>,
    shared_strings: OnceLock<Box<[Text]>>,
    styles_uri: Option<PackURI>,
    styles: OnceLock<raw::styles::Catalog>,
    flavor: Flavor,
    date_system: DateSystem,
    active_sheet: Option<usize>,
    sheets: Box<[Arc<SourceSheetData>]>,
}

/// Read-only XLSX catalog and worksheet access over a positional source.
///
/// Opening validates the OPC catalog, package relationships, workbook part,
/// and workbook-to-sheet graph. Worksheet, shared-string, and style payloads
/// are not materialized or parsed until a selected semantic read requires
/// them. ZIP central-directory discovery can still perform bounded physical
/// tail reads through the positional source; laziness here means deferred
/// logical part extraction and semantic parsing. The type has no edit or
/// output APIs.
#[derive(Clone)]
pub struct SourceBackedWorkbook {
    inner: Arc<SourceInner>,
}

/// A lifetime-free read-only worksheet handle from [`SourceBackedWorkbook`].
#[derive(Clone)]
pub struct SourceWorksheet {
    owner: Arc<SourceInner>,
    data: Arc<SourceSheetData>,
}

/// Owned semantic state at one coordinate in a source-backed worksheet.
///
/// This mirrors [`crate::cell::View`] without borrowing a lazily retained
/// worksheet store.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum SourceCellView {
    /// No cell record or covering merge exists at this coordinate.
    Missing,
    /// The coordinate is covered by a merge whose anchor is `range.start()`.
    Covered(Rect),
    /// One physical cell record is stored at this coordinate.
    Stored(Cell),
}

impl SourceCellView {
    /// Borrow the owned stored cell, when this coordinate owns one.
    #[must_use]
    pub const fn stored(&self) -> Option<&Cell> {
        match self {
            Self::Stored(cell) => Some(cell),
            Self::Missing | Self::Covered(_) => None,
        }
    }

    /// Covering merged range, if this is a non-anchor coordinate.
    #[must_use]
    pub const fn merge(&self) -> Option<Rect> {
        match self {
            Self::Covered(range) => Some(*range),
            Self::Missing | Self::Stored(_) => None,
        }
    }
}

/// One owning sparse cell returned by [`SourceWorksheet::cells`].
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SourceCell {
    /// Checked worksheet coordinate of the stored cell.
    pub address: Address,
    /// Owned semantic cell state.
    pub cell: Cell,
}

impl SourceBackedWorkbook {
    /// Open an ordinary XLSX package from a caller-provided positional source.
    pub fn from_read_at(source: Arc<dyn ReadAt>) -> Result<Self> {
        Self::from_read_at_with_limits(source, ReadLimits::default())
    }

    /// Open from a positional source with explicit OPC resource limits.
    pub fn from_read_at_with_limits(source: Arc<dyn ReadAt>, limits: ReadLimits) -> Result<Self> {
        Self::from_source_backed_package(SourceBackedPackage::from_read_at_with_limits(
            source, limits,
        )?)
    }

    /// Build the read-only XLSX facade from a validated deferred OPC package.
    pub fn from_source_backed_package(package: SourceBackedPackage) -> Result<Self> {
        let workbook = package.main_document_part()?;
        let flavor = codec::flavor(workbook.content_type()).ok_or_else(|| {
            invalid(format!(
                "main part '{}' has non-XLSX content type '{}'",
                workbook.partname(),
                workbook.content_type()
            ))
        })?;
        let catalog_bytes = workbook.data()?;
        let catalog = raw::parse_catalog(catalog_bytes.as_bytes())?;
        let sheet_parts = validate_sheet_graph(&package, &workbook, &catalog.sheets)?;
        let shared_strings_uri = validate_shared_strings(&package, &workbook)?;
        let styles_uri = validate_styles(&package, &workbook)?;

        let active_sheet = (!catalog.sheets.is_empty()).then_some(catalog.active_sheet_index);
        let mut sheets = Vec::new();
        sheets
            .try_reserve_exact(catalog.sheets.len())
            .map_err(|source| allocation("source-backed workbook sheets", source))?;
        for (position, (sheet, part)) in catalog.sheets.into_iter().zip(sheet_parts).enumerate() {
            sheets.push(Arc::new(SourceSheetData {
                position,
                name_key: crate::sheet::key(&sheet.name),
                name: sheet.name,
                kind: part.kind,
                visibility: codec::visibility(sheet.visibility),
                part_uri: part.uri,
                cells: OnceLock::new(),
            }));
        }
        let sheets = sheets.into_boxed_slice();

        Ok(Self {
            inner: Arc::new(SourceInner {
                package,
                shared_strings_uri,
                shared_strings: OnceLock::new(),
                styles_uri,
                styles: OnceLock::new(),
                flavor,
                date_system: if catalog.uses_1904_date_system {
                    DateSystem::Excel1904
                } else {
                    DateSystem::Excel1900
                },
                active_sheet,
                sheets,
            }),
        })
    }

    /// Workbook flavor derived from the mandatory workbook catalog.
    #[must_use]
    pub fn flavor(&self) -> Flavor {
        self.inner.flavor
    }

    /// Date serial system derived from the mandatory workbook catalog.
    #[must_use]
    pub fn date_system(&self) -> DateSystem {
        self.inner.date_system
    }

    /// Number of logical workbook sheets.
    #[must_use]
    pub fn len(&self) -> usize {
        self.inner.sheets.len()
    }

    /// Whether the workbook catalog contains no sheets.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.inner.sheets.is_empty()
    }

    /// Iterate lightweight sheet handles without reading worksheet payloads.
    #[must_use]
    pub fn sheets(
        &self,
    ) -> impl ExactSizeIterator<Item = SourceWorksheet> + DoubleEndedIterator + '_ {
        self.inner
            .sheets
            .iter()
            .cloned()
            .map(|data| SourceWorksheet {
                owner: Arc::clone(&self.inner),
                data,
            })
    }

    /// Look up a sheet by developer-facing name or checked zero-based position.
    pub fn sheet<'a>(&self, selector: impl Into<Selector<'a>>) -> Result<Option<SourceWorksheet>> {
        let data = match selector.into() {
            CoreSelector::Position(position) => self.inner.sheets.get(position.get()).cloned(),
            CoreSelector::Name(name) => {
                let key = crate::sheet::key(&name);
                self.inner
                    .sheets
                    .iter()
                    .find(|sheet| sheet.name_key == key)
                    .cloned()
            },
            CoreSelector::Id(never) => match never {},
            _ => return Err(Error::UnsupportedSelector),
        };
        Ok(data.map(|data| SourceWorksheet {
            owner: Arc::clone(&self.inner),
            data,
        }))
    }

    /// Return the active sheet, if the workbook catalog contains one.
    #[must_use]
    pub fn active_sheet(&self) -> Option<SourceWorksheet> {
        let data = self
            .inner
            .active_sheet
            .and_then(|position| self.inner.sheets.get(position))
            .cloned()?;
        Some(SourceWorksheet {
            owner: Arc::clone(&self.inner),
            data,
        })
    }
}

impl SourceWorksheet {
    /// Developer-facing sheet name.
    #[must_use]
    pub fn name(&self) -> &str {
        &self.data.name
    }

    /// Checked zero-based workbook position.
    #[must_use]
    pub fn position(&self) -> usize {
        self.data.position
    }

    /// Semantic sheet kind resolved from its workbook relationship.
    #[must_use]
    pub fn kind(&self) -> WorksheetKind {
        self.data.kind
    }

    /// Retained visibility state.
    #[must_use]
    pub fn visibility(&self) -> &Visibility {
        &self.data.visibility
    }

    /// Whether this is the active sheet in its source-backed workbook catalog.
    #[must_use]
    pub fn is_active(&self) -> bool {
        self.owner.active_sheet == Some(self.data.position)
    }

    /// Look up an exact logical cell state, returning an owning semantic value.
    pub fn cell<'a>(&self, at: impl Into<At<'a>>) -> Result<SourceCellView> {
        let address = at.into().resolve()?;
        Ok(match self.store()?.view(address) {
            View::Missing => SourceCellView::Missing,
            View::Covered(range) => SourceCellView::Covered(range),
            View::Stored(cell) => SourceCellView::Stored(cell.clone()),
        })
    }

    /// Read every stored cell selected by a checked range into owning values.
    ///
    /// The selected worksheet XML is loaded and parsed on first use. The
    /// returned cells are independent of the source and may outlive this
    /// worksheet handle.
    pub fn cells<'a>(&self, area: impl Into<Area<'a>>) -> Result<Vec<SourceCell>> {
        let range = area.into().resolve()?;
        let mut values = Vec::new();
        for (address, cell) in self.store()?.cells(range) {
            values
                .try_reserve(1)
                .map_err(|source| allocation("source-backed selected cells", source))?;
            values.push(SourceCell {
                address,
                cell: cell.clone(),
            });
        }
        Ok(values)
    }

    fn store(&self) -> Result<&Store> {
        if self.data.kind != WorksheetKind::Worksheet {
            return Err(Error::NotWorksheet {
                sheet: self.data.name.clone(),
            });
        }
        // `SourceBackedPackage::part` checks the captured source version even
        // when this worksheet's parsed semantic store is already retained.
        // Never return cached semantic data after a caller-visible source
        // change has invalidated the snapshot.
        let part = self.owner.package.part(&self.data.part_uri)?;
        if let Some(store) = self.data.cells.get() {
            return Ok(store);
        }

        let data = part.data()?;
        let parsed = raw::worksheet::parse(data.as_bytes(), || self.owner.shared_strings())?;
        self.owner.validate_styles(&parsed)?;
        let _publish_result = self.data.cells.set(parsed);
        self.data.cells.get().ok_or_else(|| {
            invalid("source-backed worksheet cache initialization did not publish a value")
        })
    }
}

impl SourceInner {
    fn shared_strings(&self) -> Result<Option<&[Text]>> {
        let Some(uri) = self.shared_strings_uri.as_ref() else {
            return Ok(None);
        };
        if let Some(strings) = self.shared_strings.get() {
            return Ok(Some(strings));
        }

        let data = self.package.part(uri)?.data()?;
        let parsed = raw::strings::parse(data.as_bytes())?;
        let _publish_result = self.shared_strings.set(parsed);
        self.shared_strings
            .get()
            .map(|strings| Some(strings.as_ref()))
            .ok_or_else(|| {
                invalid("source-backed shared-string cache initialization did not publish a value")
            })
    }

    fn style_count(&self) -> Result<u32> {
        let Some(uri) = self.styles_uri.as_ref() else {
            return Ok(0);
        };
        if let Some(styles) = self.styles.get() {
            return Ok(styles.len());
        }

        let data = self.package.part(uri)?.data()?;
        let parsed = raw::styles::parse(data.as_bytes())?;
        let _publish_result = self.styles.set(parsed);
        self.styles
            .get()
            .map(raw::styles::Catalog::len)
            .ok_or_else(|| {
                invalid("source-backed style cache initialization did not publish a value")
            })
    }

    fn validate_styles(&self, store: &Store) -> Result<()> {
        if !store.entries().iter().any(|entry| entry.style.is_some())
            && !store
                .row_entries()
                .iter()
                .any(|entry| entry.properties.style.is_some())
            && !store
                .column_entries()
                .iter()
                .any(|entry| entry.properties.style.is_some())
        {
            return Ok(());
        }
        let len = self.style_count()?;
        if let Some(entry) = store
            .entries()
            .iter()
            .find(|entry| entry.style.is_some_and(|key| key >= len))
        {
            return Err(invalid(format!(
                "worksheet cell {} references shared style {}, but the workbook contains {len} cell formats",
                entry.address,
                entry.style.unwrap_or_default()
            )));
        }
        if let Some(entry) = store
            .row_entries()
            .iter()
            .find(|entry| entry.properties.style.is_some_and(|key| key >= len))
        {
            return Err(invalid(format!(
                "worksheet row {} references shared style {}, but the workbook contains {len} cell formats",
                entry.index,
                entry.properties.style.unwrap_or_default()
            )));
        }
        if let Some(entry) = store
            .column_entries()
            .iter()
            .find(|entry| entry.properties.style.is_some_and(|key| key >= len))
        {
            return Err(invalid(format!(
                "worksheet column {} references shared style {}, but the workbook contains {len} cell formats",
                entry.first,
                entry.properties.style.unwrap_or_default()
            )));
        }
        Ok(())
    }
}

pub(crate) struct SheetPart {
    pub(crate) kind: WorksheetKind,
    pub(crate) uri: PackURI,
}

pub(crate) fn validate_sheet_graph(
    package: &SourceBackedPackage,
    workbook: &PartView<'_>,
    sheets: &[raw::Sheet],
) -> Result<Vec<SheetPart>> {
    let mut parts = Vec::new();
    parts
        .try_reserve_exact(sheets.len())
        .map_err(|source| allocation("source-backed workbook sheet graph", source))?;
    let mut targets = HashMap::<PackURI, usize>::new();
    targets
        .try_reserve(sheets.len())
        .map_err(|source| allocation("source-backed worksheet target lookup", source))?;
    for (position, sheet) in sheets.iter().enumerate() {
        let relationship = workbook.rels().get(&sheet.relationship_id).ok_or_else(|| {
            invalid(format!(
                "sheet '{}' references missing relationship '{}'",
                sheet.name, sheet.relationship_id
            ))
        })?;
        if relationship.is_external() {
            return Err(invalid(format!(
                "sheet '{}' relationship cannot be external",
                sheet.name
            )));
        }
        let target = relationship.target_partname()?;
        let part = package.part(&target)?;
        let kind = match relationship.reltype() {
            rt::WORKSHEET | rt::STRICT_WORKSHEET => {
                require_content_type(sheet, part.content_type(), ct::SML_WORKSHEET)?;
                WorksheetKind::Worksheet
            },
            CHARTSHEET_REL | STRICT_CHARTSHEET_REL => {
                require_content_type(sheet, part.content_type(), CHARTSHEET_CONTENT_TYPE)?;
                WorksheetKind::Chart
            },
            DIALOGSHEET_REL => WorksheetKind::Dialog,
            MACROSHEET_REL | INTL_MACROSHEET_REL => WorksheetKind::Macro,
            _ => WorksheetKind::Unknown,
        };
        let uri = part.partname().clone();
        if let Some(previous) = targets.insert(uri.clone(), position) {
            return Err(invalid(format!(
                "sheet part '{uri}' is referenced by both '{}' and '{}'",
                sheets[previous].name, sheet.name
            )));
        }
        parts.push(SheetPart { kind, uri });
    }
    Ok(parts)
}

fn validate_shared_strings(
    package: &SourceBackedPackage,
    workbook: &PartView<'_>,
) -> Result<Option<PackURI>> {
    let mut found = None;
    for relationship in workbook.rels().iter().filter(|relationship| {
        matches!(
            relationship.reltype(),
            rt::SHARED_STRINGS | rt::STRICT_SHARED_STRINGS
        )
    }) {
        if found.is_some() {
            return Err(invalid("workbook has multiple shared-string relationships"));
        }
        if relationship.is_external() {
            return Err(invalid("shared-string relationship cannot be external"));
        }
        let uri = relationship.target_partname()?;
        let part = package.part(&uri)?;
        if part.content_type() != ct::SML_SHARED_STRINGS {
            return Err(invalid(format!(
                "shared-string part has content type '{}', expected '{}'",
                part.content_type(),
                ct::SML_SHARED_STRINGS
            )));
        }
        found = Some(uri);
    }
    Ok(found)
}

fn validate_styles(
    package: &SourceBackedPackage,
    workbook: &PartView<'_>,
) -> Result<Option<PackURI>> {
    let mut found = None;
    for relationship in workbook
        .rels()
        .iter()
        .filter(|relationship| matches!(relationship.reltype(), rt::STYLES | rt::STRICT_STYLES))
    {
        if found.is_some() {
            return Err(invalid("workbook has multiple styles relationships"));
        }
        if relationship.is_external() {
            return Err(invalid("styles relationship cannot be external"));
        }
        let uri = relationship.target_partname()?;
        let part = package.part(&uri)?;
        if part.content_type() != ct::SML_STYLES {
            return Err(invalid(format!(
                "styles part has content type '{}', expected '{}'",
                part.content_type(),
                ct::SML_STYLES
            )));
        }
        found = Some(uri);
    }
    Ok(found)
}

fn require_content_type(sheet: &raw::Sheet, actual: &str, expected: &str) -> Result<()> {
    if actual != expected {
        return Err(invalid(format!(
            "sheet '{}' has content type '{actual}', expected '{expected}'",
            sheet.name
        )));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use std::io;
    use std::sync::Arc;
    use std::sync::atomic::{AtomicU64, AtomicUsize, Ordering};

    use litchi_core::{ReadAt, SourceVersion};
    use litchi_opc::constants::content_type as ct;
    use soapberry_zip::office::StreamingArchiveWriter;

    use super::{SourceBackedWorkbook, SourceCellView};
    use crate::{Cell, Error, Value};

    const SECOND_MARKER: &[u8] = b"source-backed-unrequested-second-sheet";

    struct CountingSource {
        bytes: Vec<u8>,
        marker_offset: usize,
        second_body_marker_reads: AtomicUsize,
        revision: AtomicU64,
    }

    impl CountingSource {
        fn new(bytes: Vec<u8>) -> Self {
            let marker_offset = bytes
                .windows(SECOND_MARKER.len())
                .position(|window| window == SECOND_MARKER)
                .expect("second worksheet marker is stored in archive");
            Self {
                bytes,
                marker_offset,
                second_body_marker_reads: AtomicUsize::new(0),
                revision: AtomicU64::new(0),
            }
        }

        fn changed(&self) {
            self.revision.fetch_add(1, Ordering::SeqCst);
        }
    }

    impl ReadAt for CountingSource {
        fn len(&self) -> io::Result<u64> {
            Ok(self.bytes.len() as u64)
        }

        fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
            let offset = usize::try_from(offset)
                .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "offset too large"))?;
            if offset >= self.bytes.len() {
                return Ok(0);
            }
            let count = output.len().min(self.bytes.len() - offset);
            let end = offset + count;
            if offset < self.marker_offset + SECOND_MARKER.len() && self.marker_offset < end {
                self.second_body_marker_reads.fetch_add(1, Ordering::SeqCst);
            }
            output[..count].copy_from_slice(&self.bytes[offset..end]);
            Ok(count)
        }

        fn version(&self) -> io::Result<SourceVersion> {
            Ok(SourceVersion::new(91, self.revision.load(Ordering::SeqCst)))
        }
    }

    fn source_backed_xlsx() -> Vec<u8> {
        let mut writer = StreamingArchiveWriter::new();
        writer
            .write_stored(
                "[Content_Types].xml",
                format!(
                    r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="{}"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="{}"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="{}"/></Types>"#,
                    ct::SML_SHEET_MAIN,
                    ct::SML_WORKSHEET,
                    ct::SML_WORKSHEET,
                )
                .as_bytes(),
            )
            .unwrap();
        writer
            .write_stored(
                "_rels/.rels",
                br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#,
            )
            .unwrap();
        writer
            .write_stored(
                "xl/workbook.xml",
                br#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="First" sheetId="1" r:id="rId1"/><sheet name="Second" sheetId="2" r:id="rId2"/></sheets></workbook>"#,
            )
            .unwrap();
        writer
            .write_stored(
                "xl/_rels/workbook.xml.rels",
                br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>"#,
            )
            .unwrap();
        writer
            .write_stored(
                "xl/worksheets/sheet1.xml",
                br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>7</v></c></row></sheetData></worksheet>"#,
            )
            .unwrap();
        let padding = "x".repeat(128 * 1024);
        let second = format!(
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><!--{SECOND_MARKER_TEXT}{padding}--><sheetData><row r="1"><c r="A1"><v>9</v></c></row></sheetData></worksheet>"#,
            SECOND_MARKER_TEXT = std::str::from_utf8(SECOND_MARKER).unwrap(),
        );
        writer
            .write_stored("xl/worksheets/sheet2.xml", second.as_bytes())
            .unwrap();
        writer.finish_to_bytes().unwrap()
    }

    #[test]
    fn deferred_catalog_and_first_sheet_do_not_reach_second_worksheet_body_marker() {
        let source = Arc::new(CountingSource::new(source_backed_xlsx()));
        let workbook = SourceBackedWorkbook::from_read_at(source.clone()).unwrap();

        assert_eq!(workbook.len(), 2);
        assert_eq!(
            workbook
                .sheets()
                .map(|sheet| sheet.name().to_owned())
                .collect::<Vec<_>>(),
            ["First", "Second"]
        );
        // This marker is inside the second worksheet body. It demonstrates
        // that catalog listing and first-sheet semantic reads do not extract
        // or parse that body. It does not claim ZIP indexing made no physical
        // reads of every byte range belonging to the unrequested member.
        assert_eq!(source.second_body_marker_reads.load(Ordering::SeqCst), 0);

        let first = workbook.sheet("First").unwrap().unwrap();
        assert!(matches!(
            first.cell("A1").unwrap(),
            SourceCellView::Stored(Cell::Value(Value::Number(ref value))) if value.as_str() == "7"
        ));
        assert_eq!(source.second_body_marker_reads.load(Ordering::SeqCst), 0);

        let selected = first.cells("A1:B2").unwrap();
        assert_eq!(selected.len(), 1);
        assert_eq!(selected[0].address.to_string(), "A1");
        assert_eq!(source.second_body_marker_reads.load(Ordering::SeqCst), 0);

        let second = workbook.sheet("Second").unwrap().unwrap();
        assert!(matches!(
            second.cell("A1").unwrap(),
            SourceCellView::Stored(Cell::Value(Value::Number(ref value))) if value.as_str() == "9"
        ));
        assert!(source.second_body_marker_reads.load(Ordering::SeqCst) > 0);
    }

    #[test]
    fn source_changes_are_returned_as_typed_opc_errors() {
        let source = Arc::new(CountingSource::new(source_backed_xlsx()));
        let workbook = SourceBackedWorkbook::from_read_at(source.clone()).unwrap();
        let sheet = workbook.sheet("First").unwrap().unwrap();
        assert!(matches!(
            sheet.cell("A1").unwrap(),
            SourceCellView::Stored(Cell::Value(Value::Number(ref value))) if value.as_str() == "7"
        ));
        source.changed();

        assert!(matches!(
            sheet.cell("A1"),
            Err(Error::Package(litchi_opc::OpcError::SourceChanged { .. }))
        ));
    }
}
