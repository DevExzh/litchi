//! XLSX-to-umbrella spreadsheet bridge.
//!
//! The standalone XLSX crate intentionally exposes its lossless semantic
//! model instead of the umbrella's older dynamic traits. This adapter keeps
//! that ownership boundary explicit and converts only at the high-level
//! facade seam.

use crate::xlsx::{self, Address, Rect};
use litchi_core::sheet::{
    Cell as CoreCell, CellIterator, CellValue, Result as SheetResult, RowIterator, WorkbookTrait,
    Worksheet as CoreWorksheet, WorksheetIterator,
};

type BoxError = Box<dyn std::error::Error + Send + Sync>;

fn boxed_error(error: impl std::fmt::Display) -> BoxError {
    Box::new(litchi_core::Error::Other(error.to_string()))
}

fn boxed_xlsx_error(error: xlsx::Error) -> BoxError {
    match error {
        xlsx::Error::Package(litchi_opc::OpcError::SourceChanged { expected, actual }) => {
            Box::new(litchi_core::Error::SourceChanged {
                expected,
                observed: actual,
            })
        },
        xlsx::Error::Package(
            error @ (litchi_opc::OpcError::SourceBackedOverlayUnavailable { .. }
            | litchi_opc::OpcError::PreservationUnavailable { .. }
            | litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy),
        ) => Box::new(litchi_core::Error::Unsupported(error.to_string())),
        error => Box::new(error),
    }
}

fn convert_value(value: &xlsx::cell::Value) -> CellValue {
    match value {
        xlsx::cell::Value::Bool(value) => CellValue::Bool(*value),
        xlsx::cell::Value::Number(value) => match value.as_str().parse::<i64>() {
            Ok(value) => CellValue::Int(value),
            Err(_) => CellValue::Float(value.as_f64().unwrap_or_default()),
        },
        xlsx::cell::Value::Text(value) => CellValue::String(value.as_str().to_owned()),
        xlsx::cell::Value::Date(value) => CellValue::String(value.as_str().to_owned()),
        xlsx::cell::Value::Error(value) => CellValue::Error(value.as_str().to_owned()),
        _ => CellValue::Error("unknown XLSX value kind".to_owned()),
    }
}

fn convert_cell(cell: &xlsx::cell::Cell) -> CellValue {
    match cell {
        xlsx::cell::Cell::Empty => CellValue::Empty,
        xlsx::cell::Cell::Value(value) => convert_value(value),
        xlsx::cell::Cell::Formula(formula) => CellValue::Formula {
            formula: formula.text().to_owned(),
            cached_value: formula
                .cached()
                .map(|cache| Box::new(convert_value(cache.value()))),
            is_array: matches!(formula.kind(), xlsx::formula::Kind::Array { .. }),
            array_range: None,
        },
        xlsx::cell::Cell::Unknown(unknown) => {
            CellValue::Error(format!("unknown XLSX cell kind: {}", unknown.kind()))
        },
        _ => CellValue::Error("unknown XLSX cell kind".to_owned()),
    }
}

fn coordinate(row: u32, column: u32) -> Address {
    Address::at(row, column).expect("validated XLSX grid coordinate")
}

/// Internal dynamic-trait view over a standalone XLSX workbook snapshot.
///
/// Filesystem-backed opens retain the standalone source-backed owner so the
/// workbook catalog can be listed without materializing worksheet payloads.
/// Bytes-backed opens continue to use the historical owned snapshot.
pub(crate) struct Workbook {
    workbook: WorkbookModel,
    names: Box<[String]>,
}

enum WorkbookModel {
    Owned(xlsx::Workbook),
    Source(xlsx::SourceBackedWorkbook),
}

impl std::fmt::Debug for Workbook {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Workbook")
            .field("worksheet_count", &self.names.len())
            .finish()
    }
}

impl Workbook {
    pub(crate) fn new(workbook: xlsx::Workbook) -> Self {
        let names = workbook
            .sheets()
            .map(|sheet| sheet.name().to_owned())
            .collect();
        Self {
            workbook: WorkbookModel::Owned(workbook),
            names,
        }
    }

    pub(crate) fn from_source_backed(workbook: xlsx::SourceBackedWorkbook) -> Self {
        let names = workbook
            .sheets()
            .map(|sheet| sheet.name().to_owned())
            .collect();
        Self {
            workbook: WorkbookModel::Source(workbook),
            names,
        }
    }

    pub(crate) fn ensure_source_current(&self) -> SheetResult<()> {
        match &self.workbook {
            WorkbookModel::Owned(_) => Ok(()),
            WorkbookModel::Source(workbook) => workbook
                .source_version()
                .map(|_| ())
                .map_err(boxed_xlsx_error),
        }
    }

    #[cfg(test)]
    pub(crate) const fn is_source_backed(&self) -> bool {
        matches!(&self.workbook, WorkbookModel::Source(_))
    }
}

impl WorkbookTrait for Workbook {
    fn active_worksheet(&self) -> SheetResult<Box<dyn CoreWorksheet + '_>> {
        match &self.workbook {
            WorkbookModel::Owned(workbook) => {
                let worksheet = workbook
                    .active_sheet()
                    .ok_or_else(|| boxed_error("XLSX workbook has no active worksheet"))?;
                Ok(Box::new(Worksheet {
                    worksheet: WorksheetModel::Owned(worksheet),
                }))
            },
            WorkbookModel::Source(workbook) => {
                let worksheet = workbook
                    .active_sheet()
                    .ok_or_else(|| boxed_error("XLSX workbook has no active worksheet"))?;
                Ok(Box::new(Worksheet {
                    worksheet: WorksheetModel::Source(worksheet),
                }))
            },
        }
    }

    fn worksheet_names(&self) -> &[String] {
        &self.names
    }

    fn worksheet_by_name(&self, name: &str) -> SheetResult<Box<dyn CoreWorksheet + '_>> {
        match &self.workbook {
            WorkbookModel::Owned(workbook) => {
                let worksheet = workbook
                    .sheet(name)
                    .map_err(boxed_xlsx_error)?
                    .ok_or_else(|| boxed_error(format!("XLSX worksheet '{name}' was not found")))?;
                Ok(Box::new(Worksheet {
                    worksheet: WorksheetModel::Owned(worksheet),
                }))
            },
            WorkbookModel::Source(workbook) => {
                let worksheet = workbook
                    .sheet(name)
                    .map_err(boxed_xlsx_error)?
                    .ok_or_else(|| boxed_error(format!("XLSX worksheet '{name}' was not found")))?;
                Ok(Box::new(Worksheet {
                    worksheet: WorksheetModel::Source(worksheet),
                }))
            },
        }
    }

    fn worksheet_by_index(&self, index: usize) -> SheetResult<Box<dyn CoreWorksheet + '_>> {
        match &self.workbook {
            WorkbookModel::Owned(workbook) => {
                let worksheet = workbook
                    .sheet(index)
                    .map_err(boxed_xlsx_error)?
                    .ok_or_else(|| {
                        boxed_error(format!("XLSX worksheet index {index} is out of bounds"))
                    })?;
                Ok(Box::new(Worksheet {
                    worksheet: WorksheetModel::Owned(worksheet),
                }))
            },
            WorkbookModel::Source(workbook) => {
                let worksheet = workbook
                    .sheet(index)
                    .map_err(boxed_xlsx_error)?
                    .ok_or_else(|| {
                        boxed_error(format!("XLSX worksheet index {index} is out of bounds"))
                    })?;
                Ok(Box::new(Worksheet {
                    worksheet: WorksheetModel::Source(worksheet),
                }))
            },
        }
    }

    fn worksheets(&self) -> Box<dyn WorksheetIterator<'_> + '_> {
        Box::new(Worksheets {
            workbook: self,
            index: 0,
        })
    }

    fn worksheet_count(&self) -> usize {
        match &self.workbook {
            WorkbookModel::Owned(workbook) => workbook.len(),
            WorkbookModel::Source(workbook) => workbook.len(),
        }
    }

    fn active_sheet_index(&self) -> usize {
        match &self.workbook {
            WorkbookModel::Owned(workbook) => {
                workbook.active_sheet().map_or(0, |sheet| sheet.position())
            },
            WorkbookModel::Source(workbook) => {
                workbook.active_sheet().map_or(0, |sheet| sheet.position())
            },
        }
    }

    fn is_1904_date_system(&self) -> bool {
        match &self.workbook {
            WorkbookModel::Owned(workbook) => {
                matches!(workbook.date_system(), xlsx::DateSystem::Excel1904)
            },
            WorkbookModel::Source(workbook) => {
                matches!(workbook.date_system(), xlsx::DateSystem::Excel1904)
            },
        }
    }
}

struct Worksheets<'a> {
    workbook: &'a Workbook,
    index: usize,
}

impl<'a> WorksheetIterator<'a> for Worksheets<'a> {
    fn next(&mut self) -> Option<SheetResult<Box<dyn CoreWorksheet + 'a>>> {
        if self.index >= self.workbook.worksheet_count() {
            return None;
        }
        let index = self.index;
        self.index += 1;
        Some(self.workbook.worksheet_by_index(index))
    }
}

#[derive(Clone)]
struct Worksheet {
    worksheet: WorksheetModel,
}

#[derive(Clone)]
enum WorksheetModel {
    Owned(xlsx::Worksheet),
    Source(xlsx::SourceWorksheet),
}

impl Worksheet {
    fn extent(&self) -> SheetResult<Option<Rect>> {
        let result = match &self.worksheet {
            WorksheetModel::Owned(worksheet) => worksheet.stored_extent(),
            WorksheetModel::Source(worksheet) => worksheet.stored_extent(),
        };
        match result {
            Ok(extent) => Ok(extent),
            // The legacy dynamic worksheet trait models non-grid sheet kinds
            // as empty row/cell iterators. Preserve that contract while
            // allowing malformed grid payload and source errors through the
            // fallible iterators below.
            Err(xlsx::Error::NotWorksheet { .. }) => Ok(None),
            Err(error) => Err(boxed_xlsx_error(error)),
        }
    }

    fn value_at(&self, row: u32, column: u32) -> SheetResult<CellValue> {
        match &self.worksheet {
            WorksheetModel::Owned(worksheet) => {
                match worksheet.cell((row, column)).map_err(boxed_xlsx_error)? {
                    xlsx::cell::View::Stored(cell) => Ok(convert_cell(cell)),
                    xlsx::cell::View::Missing | xlsx::cell::View::Covered(_) => {
                        Ok(CellValue::Empty)
                    },
                    _ => Ok(CellValue::Empty),
                }
            },
            WorksheetModel::Source(worksheet) => {
                match worksheet.cell((row, column)).map_err(boxed_xlsx_error)? {
                    xlsx::SourceCellView::Stored(cell) => Ok(convert_cell(&cell)),
                    xlsx::SourceCellView::Missing | xlsx::SourceCellView::Covered(_) => {
                        Ok(CellValue::Empty)
                    },
                    _ => Ok(CellValue::Empty),
                }
            },
        }
    }

    fn dimensions_inner(&self) -> SheetResult<Option<(u32, u32, u32, u32)>> {
        let Some(extent) = self.extent()? else {
            return Ok(None);
        };
        let start = extent.start();
        let (end_row, end_column) = extent.end();
        Ok(Some((
            start.row().get(),
            start.column().get(),
            end_row.saturating_sub(1),
            end_column.saturating_sub(1),
        )))
    }

    fn row_values(&self, row: u32) -> SheetResult<Vec<CellValue>> {
        let Some((_, _, _, end_column)) = self.dimensions_inner()? else {
            return Ok(Vec::new());
        };
        match &self.worksheet {
            WorksheetModel::Owned(_) => (0..=end_column)
                .map(|column| self.value_at(row, column))
                .collect(),
            WorksheetModel::Source(worksheet) => {
                let end_row = row
                    .checked_add(1)
                    .ok_or_else(|| boxed_error("XLSX row range overflow"))?;
                let end_column = end_column
                    .checked_add(1)
                    .ok_or_else(|| boxed_error("XLSX column range overflow"))?;
                let width = usize::try_from(end_column)
                    .map_err(|_| boxed_error("XLSX row width does not fit usize"))?;
                let mut values = Vec::new();
                values
                    .try_reserve_exact(width)
                    .map_err(|error| boxed_error(format!("XLSX row allocation failed: {error}")))?;
                values.resize(width, CellValue::Empty);
                for entry in worksheet
                    .cells((row, 0, end_row, end_column))
                    .map_err(boxed_xlsx_error)?
                {
                    let column = usize::try_from(entry.address.column().get())
                        .map_err(|_| boxed_error("XLSX column does not fit usize"))?;
                    let value = values
                        .get_mut(column)
                        .ok_or_else(|| boxed_error("XLSX cell lies outside its row extent"))?;
                    *value = convert_cell(&entry.cell);
                }
                Ok(values)
            },
        }
    }

    fn stored_cells(&self) -> SheetResult<Cells> {
        let Some(extent) = self.extent()? else {
            return Ok(Cells::empty());
        };
        match &self.worksheet {
            WorksheetModel::Owned(worksheet) => Ok(Cells {
                inner: CellsInner::Owned(
                    worksheet
                        .cells(extent)
                        .map_err(boxed_xlsx_error)?
                        .map(|(address, cell)| {
                            XlsxCell::new(
                                address.row().get(),
                                address.column().get(),
                                convert_cell(cell),
                            )
                        })
                        .collect::<Vec<_>>()
                        .into_iter(),
                ),
            }),
            WorksheetModel::Source(worksheet) => Ok(Cells {
                inner: CellsInner::Source(
                    worksheet
                        .cells(extent)
                        .map_err(boxed_xlsx_error)?
                        .into_iter(),
                ),
            }),
        }
    }
}

impl CoreWorksheet for Worksheet {
    fn name(&self) -> &str {
        match &self.worksheet {
            WorksheetModel::Owned(worksheet) => worksheet.name(),
            WorksheetModel::Source(worksheet) => worksheet.name(),
        }
    }

    fn row_count(&self) -> usize {
        self.dimensions_inner()
            .ok()
            .flatten()
            .map_or(0, |(_, _, row, _)| row as usize + 1)
    }

    fn column_count(&self) -> usize {
        self.dimensions_inner()
            .ok()
            .flatten()
            .map_or(0, |(_, _, _, column)| column as usize + 1)
    }

    fn dimensions(&self) -> Option<(u32, u32, u32, u32)> {
        self.dimensions_inner().ok().flatten()
    }

    fn cell(&self, row: u32, column: u32) -> SheetResult<Box<dyn CoreCell + '_>> {
        Ok(Box::new(XlsxCell::new(
            row.saturating_sub(1),
            column.saturating_sub(1),
            self.value_at(row.saturating_sub(1), column.saturating_sub(1))?,
        )))
    }

    fn cell_by_coordinate(&self, coordinate: &str) -> SheetResult<Box<dyn CoreCell + '_>> {
        let address = Address::from_a1(coordinate).map_err(boxed_error)?;
        self.cell(address.row().get() + 1, address.column().get() + 1)
    }

    fn cells(&self) -> Box<dyn CellIterator<'_> + '_> {
        Box::new(match self.stored_cells() {
            Ok(cells) => cells,
            Err(error) => Cells::error(error),
        })
    }

    fn rows(&self) -> Box<dyn RowIterator<'_> + '_> {
        let dimensions = self.dimensions_inner();
        let (end, error) = match dimensions {
            Ok(Some((_, _, row, _))) => (row as usize + 1, None),
            Ok(None) => (0, None),
            Err(error) => (0, Some(error)),
        };
        Box::new(Rows {
            worksheet: self,
            index: 0,
            end,
            error,
        })
    }

    fn row(&self, row_idx: usize) -> SheetResult<std::borrow::Cow<'_, [CellValue]>> {
        Ok(std::borrow::Cow::Owned(self.row_values(row_idx as u32)?))
    }

    fn cell_value(&self, row: u32, column: u32) -> SheetResult<std::borrow::Cow<'_, CellValue>> {
        Ok(std::borrow::Cow::Owned(self.value_at(
            row.saturating_sub(1),
            column.saturating_sub(1),
        )?))
    }
}

struct Cells {
    inner: CellsInner,
}

enum CellsInner {
    Owned(std::vec::IntoIter<XlsxCell>),
    Source(std::vec::IntoIter<xlsx::SourceCell>),
    Error(Option<BoxError>),
}

impl Cells {
    fn empty() -> Self {
        Self {
            inner: CellsInner::Owned(Vec::new().into_iter()),
        }
    }

    fn error(error: BoxError) -> Self {
        Self {
            inner: CellsInner::Error(Some(error)),
        }
    }
}

impl<'a> CellIterator<'a> for Cells {
    fn next(&mut self) -> Option<SheetResult<Box<dyn CoreCell + 'a>>> {
        let cell = match &mut self.inner {
            CellsInner::Owned(cells) => cells.next(),
            CellsInner::Source(cells) => cells.next().map(|entry| {
                XlsxCell::new(
                    entry.address.row().get(),
                    entry.address.column().get(),
                    convert_cell(&entry.cell),
                )
            }),
            CellsInner::Error(error) => return error.take().map(Err),
        }?;
        Some(Ok(Box::new(cell) as Box<dyn CoreCell + 'a>))
    }
}

struct Rows<'a> {
    worksheet: &'a Worksheet,
    index: usize,
    end: usize,
    error: Option<BoxError>,
}

impl<'a> RowIterator<'a> for Rows<'a> {
    fn next(&mut self) -> Option<SheetResult<std::borrow::Cow<'a, [CellValue]>>> {
        if let Some(error) = self.error.take() {
            return Some(Err(error));
        }
        if self.index >= self.end {
            return None;
        }
        let index = self.index;
        self.index += 1;
        Some(
            self.worksheet
                .row_values(index as u32)
                .map(std::borrow::Cow::Owned),
        )
    }
}

#[derive(Debug, Clone)]
struct XlsxCell {
    row: u32,
    column: u32,
    value: CellValue,
}

impl XlsxCell {
    fn new(row: u32, column: u32, value: CellValue) -> Self {
        Self {
            row: row + 1,
            column: column + 1,
            value,
        }
    }
}

impl CoreCell for XlsxCell {
    fn row(&self) -> u32 {
        self.row
    }

    fn column(&self) -> u32 {
        self.column
    }

    fn coordinate(&self) -> String {
        coordinate(self.row - 1, self.column - 1).a1()
    }

    fn value(&self) -> &CellValue {
        &self.value
    }

    fn is_formula(&self) -> bool {
        matches!(self.value, CellValue::Formula { .. })
    }
}

#[cfg(test)]
mod tests {
    use std::io::{Cursor, Write};

    use super::{boxed_xlsx_error, xlsx};
    use crate::sheet::{CellValue, WorkbookTrait};

    #[test]
    fn package_publication_capabilities_reach_the_core_error_class() {
        for error in [
            litchi_opc::OpcError::PreservationUnavailable {
                reason: "opaque ZIP framing".to_owned(),
            },
            litchi_opc::OpcError::SourceBackedOverlayUnavailable {
                reason: "opaque member cannot be patched".to_owned(),
            },
            litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy,
        ] {
            let boxed = boxed_xlsx_error(xlsx::Error::Package(error));
            assert!(matches!(
                boxed.downcast_ref::<litchi_core::Error>(),
                Some(litchi_core::Error::Unsupported(_))
            ));
        }
    }

    const SPREADSHEETML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    const RELATIONSHIPS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const PACKAGE_RELATIONSHIPS: &str =
        "http://schemas.openxmlformats.org/package/2006/relationships";
    const OFFICE_DOCUMENT_RELATIONSHIP: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
    const WORKSHEET_RELATIONSHIP: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";

    fn fixture(case_variant_targets: bool, malformed_second: bool) -> Vec<u8> {
        let root_target = if case_variant_targets {
            "XL/WORKBOOK.XML"
        } else {
            "xl/workbook.xml"
        };
        let first_target = if case_variant_targets {
            "WORKSHEETS/SHEET1.XML"
        } else {
            "worksheets/sheet1.xml"
        };
        let second_target = if case_variant_targets {
            "WORKSHEETS/SHEET2.XML"
        } else {
            "worksheets/sheet2.xml"
        };
        let workbook_relationships = format!(
            r#"<Relationships xmlns="{PACKAGE_RELATIONSHIPS}"><Relationship Id="rId1" Type="{WORKSHEET_RELATIONSHIP}" Target="{first_target}"/><Relationship Id="rId2" Type="{WORKSHEET_RELATIONSHIP}" Target="{second_target}"/></Relationships>"#
        );
        let root_relationships = format!(
            r#"<Relationships xmlns="{PACKAGE_RELATIONSHIPS}"><Relationship Id="rId1" Type="{OFFICE_DOCUMENT_RELATIONSHIP}" Target="{root_target}"/></Relationships>"#
        );
        let content_types = r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>"#;
        let workbook = format!(
            r#"<workbook xmlns="{SPREADSHEETML}" xmlns:r="{RELATIONSHIPS}"><workbookPr date1904="1"/><bookViews><workbookView activeTab="1"/></bookViews><sheets><sheet name="First" sheetId="1" r:id="rId1"/><sheet name="Second" sheetId="2" r:id="rId2"/></sheets></workbook>"#
        );
        let first = format!(
            r#"<worksheet xmlns="{SPREADSHEETML}"><!--FIRST-WORKSHEET-PAYLOAD-MARKER--><dimension ref="A1:B2"/><sheetData><row r="1"><c r="A1"><v>7</v></c></row><row r="2"><c r="B2" t="inlineStr"><is><t>second</t></is></c></row></sheetData></worksheet>"#
        );
        let second = if malformed_second {
            format!(
                r#"<worksheet xmlns="{SPREADSHEETML}"><!--SECOND-WORKSHEET-PAYLOAD-MARKER--><sheetData><row r="1"><c r="A1"><v>9</v></c>"#
            )
        } else {
            format!(
                r#"<worksheet xmlns="{SPREADSHEETML}"><!--SECOND-WORKSHEET-PAYLOAD-MARKER--><dimension ref="A1"/><sheetData><row r="1"><c r="A1"><v>9</v></c></row></sheetData></worksheet>"#
            )
        };

        let mut output = Cursor::new(Vec::new());
        let mut writer = zip::ZipWriter::new(&mut output);
        let options = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Stored);
        let entries = [
            ("[Content_Types].xml", content_types.as_bytes()),
            ("_rels/.rels", root_relationships.as_bytes()),
            ("xl/workbook.xml", workbook.as_bytes()),
            (
                "xl/_rels/workbook.xml.rels",
                workbook_relationships.as_bytes(),
            ),
            ("xl/worksheets/sheet1.xml", first.as_bytes()),
            ("xl/worksheets/sheet2.xml", second.as_bytes()),
        ];
        for (name, bytes) in entries {
            writer
                .start_file(name, options)
                .expect("start fixture member");
            writer.write_all(bytes).expect("write fixture member");
        }
        writer.finish().expect("finish fixture archive");
        output.into_inner()
    }

    fn path_for(bytes: &[u8]) -> tempfile::NamedTempFile {
        let file = tempfile::NamedTempFile::new().expect("temporary XLSX path");
        std::fs::write(file.path(), bytes).expect("write temporary XLSX");
        file
    }

    fn corrupt_central_crc(mut bytes: Vec<u8>, member: &[u8]) -> Vec<u8> {
        let signature = b"PK\x01\x02";
        let mut offset = 0;
        while let Some(relative) = bytes[offset..]
            .windows(signature.len())
            .position(|window| window == signature)
        {
            let central = offset + relative;
            if central + 46 > bytes.len() {
                break;
            }
            let name_len = u16::from_le_bytes([bytes[central + 28], bytes[central + 29]]) as usize;
            let extra_len = u16::from_le_bytes([bytes[central + 30], bytes[central + 31]]) as usize;
            let comment_len =
                u16::from_le_bytes([bytes[central + 32], bytes[central + 33]]) as usize;
            let name_start = central + 46;
            let name_end = name_start + name_len;
            if name_end > bytes.len() {
                break;
            }
            if &bytes[name_start..name_end] == member {
                bytes[central + 16] ^= 1;
                return bytes;
            }
            offset = name_end + extra_len + comment_len;
            if offset >= bytes.len() {
                break;
            }
        }
        panic!("fixture member CRC was not found");
    }

    fn assert_workbook_trait_surface(workbook: &dyn WorkbookTrait) {
        assert_eq!(workbook.worksheet_names(), ["First", "Second"]);
        assert_eq!(workbook.worksheet_count(), 2);
        assert_eq!(workbook.active_sheet_index(), 1);
        assert!(workbook.is_1904_date_system());

        let active = workbook.active_worksheet().expect("active worksheet");
        assert_eq!(active.name(), "Second");
        let by_name = workbook
            .worksheet_by_name("First")
            .expect("worksheet by name");
        let by_index = workbook.worksheet_by_index(0).expect("worksheet by index");
        assert_eq!(by_name.name(), by_index.name());
        assert!(workbook.worksheet_by_index(2).is_err());

        let mut worksheets = workbook.worksheets();
        assert_eq!(
            worksheets.next().expect("first worksheet").unwrap().name(),
            "First"
        );
        assert_eq!(
            worksheets.next().expect("second worksheet").unwrap().name(),
            "Second"
        );
        assert!(worksheets.next().is_none());

        assert_eq!(by_name.row_count(), 2);
        assert_eq!(by_name.column_count(), 2);
        assert_eq!(by_name.dimensions(), Some((0, 0, 1, 1)));
        assert_eq!(
            by_name.cell_value(1, 1).unwrap().as_ref(),
            &CellValue::Int(7)
        );
        assert_eq!(
            by_name.cell_by_coordinate("B2").unwrap().value(),
            &CellValue::String("second".to_owned())
        );
        let cell = by_name.cell(1, 1).expect("cell");
        assert_eq!(cell.coordinate(), "A1");
        assert_eq!(cell.row(), 1);
        assert_eq!(cell.column(), 1);
        assert!(!cell.is_empty());
        assert!(!cell.is_formula());

        let mut cells = by_name.cells();
        assert_eq!(
            cells.next().expect("first cell").unwrap().coordinate(),
            "A1"
        );
        assert_eq!(
            cells.next().expect("second cell").unwrap().coordinate(),
            "B2"
        );
        assert!(cells.next().is_none());
        assert_eq!(
            by_name.row(1).unwrap().as_ref(),
            &[CellValue::Empty, CellValue::String("second".to_owned())]
        );
        let mut rows = by_name.rows();
        assert_eq!(rows.next().expect("first row").unwrap().len(), 2);
        assert_eq!(rows.next().expect("second row").unwrap().len(), 2);
        assert!(rows.next().is_none());
    }

    #[test]
    fn source_path_matches_eager_bytes_for_the_complete_trait_surface() {
        let bytes = fixture(false, false);
        let path = path_for(&bytes);
        let source = crate::sheet::open_workbook(path.path()).expect("source-backed open");
        let eager = crate::sheet::open_workbook_from_bytes(&bytes).expect("eager open");

        assert_workbook_trait_surface(source.as_ref());
        assert_workbook_trait_surface(eager.as_ref());
    }

    #[cfg(any(unix, windows))]
    #[test]
    fn source_path_catalog_and_selection_defer_malformed_payload() {
        let path = path_for(&fixture(false, true));
        let workbook = crate::sheet::open_workbook_with_limits(
            path.path(),
            crate::xlsx::ReadLimits::default(),
        )
        .expect("catalog-only source-backed open");

        // Catalog methods and selecting a handle do not extract either sheet
        // body. The malformed second member is therefore still unopened.
        assert_eq!(workbook.worksheet_names(), ["First", "Second"]);
        assert_eq!(workbook.worksheet_count(), 2);
        assert_eq!(workbook.active_sheet_index(), 1);
        assert!(workbook.is_1904_date_system());
        let second = workbook
            .worksheet_by_name("Second")
            .expect("select malformed worksheet");
        assert!(second.dimensions().is_none());
        assert!(second.cell_value(1, 1).is_err());
        assert!(second.row(0).is_err());
    }

    #[cfg(any(unix, windows))]
    #[test]
    fn source_path_defers_crc_failure_until_selected_fallible_reads() {
        let bytes = corrupt_central_crc(fixture(false, false), b"xl/worksheets/sheet2.xml");
        let path = path_for(&bytes);
        let workbook = crate::sheet::open_workbook(path.path()).expect("catalog-only CRC open");
        assert_eq!(workbook.worksheet_names(), ["First", "Second"]);
        let second = workbook
            .worksheet_by_name("Second")
            .expect("select CRC-corrupt worksheet");
        assert!(second.dimensions().is_none());
        assert!(second.cell_value(1, 1).is_err());
        assert!(second.row(0).is_err());
    }

    #[cfg(any(unix, windows))]
    #[test]
    fn source_path_resolves_case_variant_relationship_targets() {
        let bytes = fixture(true, false);
        let path = path_for(&bytes);
        let workbook = crate::sheet::open_workbook(path.path()).expect("case-variant target open");
        assert_eq!(workbook.worksheet_names(), ["First", "Second"]);
        let first = workbook
            .worksheet_by_index(0)
            .expect("case-variant first worksheet");
        assert_eq!(first.cell_value(1, 1).unwrap().as_ref(), &CellValue::Int(7));
    }
}
