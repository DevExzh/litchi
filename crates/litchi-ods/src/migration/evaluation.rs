//! ODS adapters for the shared spreadsheet and formula-evaluation APIs.
//!
//! `Spreadsheet` intentionally keeps its package-oriented API mutable because
//! materializing sheets reads `content.xml`.  The shared `WorkbookTrait`, on
//! the other hand, provides immutable, borrowing worksheet access.  This
//! module bridges those designs with an owned, immutable snapshot.  Parsing is
//! performed once when the snapshot is created; subsequent cell reads borrow
//! directly from its compact row storage.

use super::{Cell as OdsCell, CellMatrixSpan, CellValue as OdsCellValue, Sheet, Spreadsheet};
use litchi_core::{
    Error, Result as OdfResult,
    sheet::{
        Cell as SheetCell, CellIterator, CellValue, Result as SheetResult, RowIterator,
        WorkbookTrait, Worksheet, WorksheetIterator,
    },
};
use std::{borrow::Cow, path::Path};

/// Immutable ODS workbook snapshot suitable for `WorkbookTrait` consumers.
///
/// Create one with [`Spreadsheet::evaluation_workbook`] when the original
/// package must remain available, or [`Spreadsheet::into_evaluation_workbook`]
/// to avoid retaining it.  The snapshot is intentionally read-only: it is a
/// stable input to consumers such as `litchi_eval::FormulaEvaluator`, not a
/// second mutable ODS model.
#[derive(Debug)]
pub struct OdsWorkbook {
    sheets: Vec<OdsWorksheetData>,
    worksheet_names: Vec<String>,
}

#[derive(Debug)]
struct OdsWorksheetData {
    name: String,
    rows: Vec<Vec<CellValue>>,
    present_cells: Vec<(u32, u32)>,
    dimensions: Option<(u32, u32, u32, u32)>,
}

impl OdsWorkbook {
    /// Open an ODS file and create an immutable evaluation snapshot.
    pub fn open(path: impl AsRef<Path>) -> OdfResult<Self> {
        Spreadsheet::open(path)?.into_evaluation_workbook()
    }

    /// Parse ODS bytes and create an immutable evaluation snapshot.
    pub fn from_bytes(bytes: Vec<u8>) -> OdfResult<Self> {
        Spreadsheet::from_bytes(bytes)?.into_evaluation_workbook()
    }

    /// Convert an existing spreadsheet into an immutable evaluation snapshot.
    pub fn from_spreadsheet(spreadsheet: Spreadsheet) -> OdfResult<Self> {
        spreadsheet.into_evaluation_workbook()
    }

    pub(crate) fn from_sheets(sheets: Vec<Sheet>) -> OdfResult<Self> {
        let mut worksheet_names = Vec::with_capacity(sheets.len());
        let mut evaluation_sheets = Vec::with_capacity(sheets.len());

        for sheet in sheets {
            let sheet = OdsWorksheetData::from_sheet(sheet)?;
            worksheet_names.push(sheet.name.clone());
            evaluation_sheets.push(sheet);
        }

        Ok(Self {
            sheets: evaluation_sheets,
            worksheet_names,
        })
    }
}

impl OdsWorksheetData {
    fn from_sheet(sheet: Sheet) -> OdfResult<Self> {
        let Sheet { name, rows, .. } = sheet;
        let mut stored_rows = Vec::<Vec<CellValue>>::new();
        let mut present_cells = Vec::new();
        let mut max_row = 0u32;
        let mut max_col = 0u32;

        for row in rows {
            let row_index = one_based_coordinate(row.index, "row")?;
            let row_slot = row.index;
            if stored_rows.len() <= row_slot {
                stored_rows.resize_with(row_slot.saturating_add(1), Vec::new);
            }

            let values = &mut stored_rows[row_slot];
            for cell in row.cells {
                let column_index = one_based_coordinate(cell.col, "column")?;
                let column_slot = cell.col;
                if values.len() <= column_slot {
                    values.resize_with(column_slot.saturating_add(1), || CellValue::Empty);
                }
                values[column_slot] = convert_cell(cell);
                present_cells.push((row_index, column_index));
                max_row = max_row.max(row_index);
                max_col = max_col.max(column_index);
            }
        }

        let dimensions = (max_row != 0 && max_col != 0).then_some((1, 1, max_row, max_col));
        Ok(Self {
            name,
            rows: stored_rows,
            present_cells,
            dimensions,
        })
    }

    fn value_at(&self, row: u32, column: u32) -> Option<&CellValue> {
        let row = row
            .checked_sub(1)
            .and_then(|value| usize::try_from(value).ok())?;
        let column = column
            .checked_sub(1)
            .and_then(|value| usize::try_from(value).ok())?;
        self.rows.get(row)?.get(column)
    }

    fn row_values(&self, row_index: usize) -> SheetResult<Cow<'_, [CellValue]>> {
        let Some(max_col) = self.dimensions.map(|(_, _, _, max_col)| max_col) else {
            return Ok(Cow::Borrowed(&[]));
        };
        let row = self.rows.get(row_index).ok_or_else(|| {
            sheet_error(format!(
                "ODS row index {row_index} is outside the worksheet"
            ))
        })?;
        let max_col = usize::try_from(max_col)
            .map_err(|_| sheet_error("ODS column count does not fit platform memory"))?;

        if row.len() == max_col {
            return Ok(Cow::Borrowed(row));
        }

        let mut values = Vec::with_capacity(max_col);
        values.extend(row.iter().cloned());
        values.resize_with(max_col, || CellValue::Empty);
        Ok(Cow::Owned(values))
    }
}

impl WorkbookTrait for OdsWorkbook {
    fn active_worksheet(&self) -> SheetResult<Box<dyn Worksheet + '_>> {
        self.worksheet_by_index(self.active_sheet_index())
    }

    fn worksheet_names(&self) -> &[String] {
        &self.worksheet_names
    }

    fn worksheet_by_name(&self, name: &str) -> SheetResult<Box<dyn Worksheet + '_>> {
        let sheet = self
            .sheets
            .iter()
            .find(|sheet| sheet.name == name)
            .ok_or_else(|| sheet_error(format!("ODS worksheet '{name}' was not found")))?;
        Ok(Box::new(OdsWorksheet { sheet }))
    }

    fn worksheet_by_index(&self, index: usize) -> SheetResult<Box<dyn Worksheet + '_>> {
        let sheet = self.sheets.get(index).ok_or_else(|| {
            sheet_error(format!(
                "ODS worksheet index {index} is outside the workbook"
            ))
        })?;
        Ok(Box::new(OdsWorksheet { sheet }))
    }

    fn worksheets(&self) -> Box<dyn WorksheetIterator<'_> + '_> {
        Box::new(OdsWorksheetIterator {
            workbook: self,
            index: 0,
        })
    }

    fn worksheet_count(&self) -> usize {
        self.sheets.len()
    }

    fn active_sheet_index(&self) -> usize {
        // ODF stores active-table UI state in settings markup.  It is not a
        // workbook calculation setting, so use the first sheet consistently.
        0
    }
}

struct OdsWorksheet<'a> {
    sheet: &'a OdsWorksheetData,
}

impl Worksheet for OdsWorksheet<'_> {
    fn name(&self) -> &str {
        &self.sheet.name
    }

    fn row_count(&self) -> usize {
        self.sheet.rows.len()
    }

    fn column_count(&self) -> usize {
        self.sheet
            .dimensions
            .and_then(|(_, _, _, max_col)| usize::try_from(max_col).ok())
            .unwrap_or(0)
    }

    fn dimensions(&self) -> Option<(u32, u32, u32, u32)> {
        self.sheet.dimensions
    }

    fn cell(&self, row: u32, column: u32) -> SheetResult<Box<dyn SheetCell + '_>> {
        validate_one_based(row, "row")?;
        validate_one_based(column, "column")?;
        Ok(Box::new(OdsSheetCell {
            row,
            column,
            value: self.sheet.value_at(row, column).unwrap_or(CellValue::EMPTY),
        }))
    }

    fn cell_by_coordinate(&self, coordinate: &str) -> SheetResult<Box<dyn SheetCell + '_>> {
        let (row, column) = parse_a1_coordinate(coordinate)?;
        self.cell(row, column)
    }

    fn cells(&self) -> Box<dyn CellIterator<'_> + '_> {
        Box::new(OdsCellIterator {
            sheet: self.sheet,
            index: 0,
        })
    }

    fn rows(&self) -> Box<dyn RowIterator<'_> + '_> {
        Box::new(OdsRowIterator {
            sheet: self.sheet,
            index: 0,
        })
    }

    fn row(&self, row_idx: usize) -> SheetResult<Cow<'_, [CellValue]>> {
        self.sheet.row_values(row_idx)
    }

    fn cell_value(&self, row: u32, column: u32) -> SheetResult<Cow<'_, CellValue>> {
        validate_one_based(row, "row")?;
        validate_one_based(column, "column")?;
        Ok(Cow::Borrowed(
            self.sheet.value_at(row, column).unwrap_or(CellValue::EMPTY),
        ))
    }
}

struct OdsSheetCell<'a> {
    row: u32,
    column: u32,
    value: &'a CellValue,
}

impl SheetCell for OdsSheetCell<'_> {
    fn row(&self) -> u32 {
        self.row
    }

    fn column(&self) -> u32 {
        self.column
    }

    fn coordinate(&self) -> String {
        format!("{}{}", column_to_letters(self.column), self.row)
    }

    fn value(&self) -> &CellValue {
        self.value
    }

    fn is_formula(&self) -> bool {
        matches!(self.value, CellValue::Formula { .. })
    }
}

struct OdsCellIterator<'a> {
    sheet: &'a OdsWorksheetData,
    index: usize,
}

impl<'a> CellIterator<'a> for OdsCellIterator<'a> {
    fn next(&mut self) -> Option<SheetResult<Box<dyn SheetCell + 'a>>> {
        let &(row, column) = self.sheet.present_cells.get(self.index)?;
        self.index += 1;
        let value = self.sheet.value_at(row, column)?;
        Some(Ok(Box::new(OdsSheetCell { row, column, value })))
    }
}

struct OdsRowIterator<'a> {
    sheet: &'a OdsWorksheetData,
    index: usize,
}

impl<'a> RowIterator<'a> for OdsRowIterator<'a> {
    fn next(&mut self) -> Option<SheetResult<Cow<'a, [CellValue]>>> {
        if self.index >= self.sheet.rows.len() {
            return None;
        }
        let row_index = self.index;
        self.index += 1;
        Some(self.sheet.row_values(row_index))
    }
}

struct OdsWorksheetIterator<'a> {
    workbook: &'a OdsWorkbook,
    index: usize,
}

impl<'a> WorksheetIterator<'a> for OdsWorksheetIterator<'a> {
    fn next(&mut self) -> Option<SheetResult<Box<dyn Worksheet + 'a>>> {
        let sheet = self.workbook.sheets.get(self.index)?;
        self.index += 1;
        Some(Ok(Box::new(OdsWorksheet { sheet })))
    }
}

fn convert_cell(cell: OdsCell) -> CellValue {
    let OdsCell {
        value,
        formula,
        matrix_span,
        row,
        col,
        ..
    } = cell;
    let cached_value = convert_value(value);

    let Some(formula) = formula else {
        return cached_value;
    };

    CellValue::Formula {
        formula: normalize_open_formula(&formula),
        cached_value: (!matches!(cached_value, CellValue::Empty)).then(|| Box::new(cached_value)),
        is_array: matrix_span.is_some(),
        array_range: matrix_span.and_then(|span| matrix_range(row, col, span)),
    }
}

fn convert_value(value: OdsCellValue) -> CellValue {
    match value {
        OdsCellValue::Empty => CellValue::Empty,
        OdsCellValue::Text(value) => CellValue::String(value),
        OdsCellValue::Number(value)
        | OdsCellValue::Currency(value, _)
        | OdsCellValue::Percentage(value) => CellValue::Float(value),
        // The shared value model has an Excel-style serial datetime only.  An
        // ODF date can use an arbitrary null-date setting, so retaining its
        // ISO 8601 lexical representation is safer than inventing a serial.
        OdsCellValue::Date(value) | OdsCellValue::Time(value) => CellValue::String(value),
        OdsCellValue::Boolean(value) => CellValue::Bool(value),
    }
}

fn matrix_range(row: usize, col: usize, span: CellMatrixSpan) -> Option<String> {
    let start_row = one_based_u32(row)?;
    let start_col = one_based_u32(col)?;
    let end_row = start_row.checked_add(u32::try_from(span.rows()).ok()?.checked_sub(1)?)?;
    let end_col = start_col.checked_add(u32::try_from(span.columns()).ok()?.checked_sub(1)?)?;
    Some(format!(
        "{}:{}",
        a1_coordinate(start_row, start_col),
        a1_coordinate(end_row, end_col)
    ))
}

/// Normalize the common OpenFormula spelling used in ODS files for the
/// shared evaluator's Excel-like expression grammar.
///
/// The conversion intentionally only handles syntax that is semantically
/// equivalent: `of:=` prefixes, semicolon argument separators, and bracketed
/// A1 references.  Unsupported namespaces, external workbook references, and
/// unrecognized bracket contents are retained verbatim, causing the evaluator
/// to report an ordinary unsupported-formula result instead of resolving an
/// external target.
pub fn normalize_open_formula(input: &str) -> String {
    let formula = input.trim();
    let body = formula
        .get(..4)
        .filter(|prefix| prefix.eq_ignore_ascii_case("of:="))
        .map(|_| &formula[4..])
        .or_else(|| formula.strip_prefix('='))
        .unwrap_or(formula);

    let mut output = String::with_capacity(body.len());
    let bytes = body.as_bytes();
    let mut index = 0;
    let mut quoted = false;

    while index < bytes.len() {
        match bytes[index] {
            b'"' => {
                output.push('"');
                index += 1;
                if quoted && bytes.get(index) == Some(&b'"') {
                    output.push('"');
                    index += 1;
                } else {
                    quoted = !quoted;
                }
            },
            b'[' if !quoted => {
                let Some(end_offset) = bytes[index + 1..].iter().position(|byte| *byte == b']')
                else {
                    output.push('[');
                    index += 1;
                    continue;
                };
                let end = index + 1 + end_offset;
                let reference = &body[index + 1..end];
                if let Some(normalized) = normalize_bracket_reference(reference) {
                    output.push_str(&normalized);
                } else {
                    output.push('[');
                    output.push_str(reference);
                    output.push(']');
                }
                index = end + 1;
            },
            b';' if !quoted => {
                output.push(',');
                index += 1;
            },
            _ => {
                let character = body[index..]
                    .chars()
                    .next()
                    .expect("index is within a UTF-8 string");
                output.push(character);
                index += character.len_utf8();
            },
        }
    }

    output
}

#[derive(Debug)]
struct OpenFormulaReference<'a> {
    sheet: Option<&'a str>,
    cell: &'a str,
}

fn normalize_bracket_reference(value: &str) -> Option<String> {
    let (start, end) = split_reference_range(value)?;
    let start = parse_open_formula_reference(start)?;
    let end = match end {
        Some(value) => Some(parse_open_formula_reference(value)?),
        None => None,
    };

    let sheet = match (
        start.sheet,
        end.as_ref().and_then(|reference| reference.sheet),
    ) {
        (Some(left), Some(right)) if left != right => return None,
        (Some(sheet), _) | (None, Some(sheet)) => Some(sheet),
        (None, None) => None,
    };

    let mut output = String::new();
    if let Some(sheet) = sheet {
        output.push_str(&normalize_sheet_name(sheet)?);
        output.push('!');
    }
    output.push_str(start.cell);
    if let Some(end) = end {
        output.push(':');
        output.push_str(end.cell);
    }
    Some(output)
}

fn split_reference_range(value: &str) -> Option<(&str, Option<&str>)> {
    let value = value.trim();
    if value.is_empty() || value.contains('#') {
        // `#` is used by external ODF file references.  Leave it untouched:
        // the evaluator must never resolve external documents.
        return None;
    }
    let mut quoted = false;
    for (index, character) in value.char_indices() {
        if character == '\'' {
            quoted = !quoted;
        } else if character == ':' && !quoted {
            return Some((&value[..index], Some(&value[index + 1..])));
        }
    }
    Some((value, None))
}

fn parse_open_formula_reference(value: &str) -> Option<OpenFormulaReference<'_>> {
    let value = value.trim();
    if value.is_empty() {
        return None;
    }
    if let Some(cell) = value.strip_prefix('.') {
        return is_a1_reference(cell).then_some(OpenFormulaReference { sheet: None, cell });
    }

    let (sheet, cell) = value.rsplit_once('.')?;
    let sheet = sheet.strip_prefix('$').unwrap_or(sheet).trim();
    let cell = cell.trim();
    (!sheet.is_empty() && is_a1_reference(cell)).then_some(OpenFormulaReference {
        sheet: Some(sheet),
        cell,
    })
}

fn is_a1_reference(value: &str) -> bool {
    parse_a1_coordinate(value).is_ok()
}

fn normalize_sheet_name(value: &str) -> Option<String> {
    let value = value.trim();
    if value.is_empty() {
        return None;
    }
    let quoted = value.starts_with('\'') || value.ends_with('\'');
    let unquoted = if quoted {
        value.strip_prefix('\'')?.strip_suffix('\'')?
    } else {
        value
    };
    if unquoted.is_empty() || (!quoted && unquoted.contains('\'')) {
        return None;
    }

    let mut unescaped = String::with_capacity(unquoted.len());
    let mut characters = unquoted.chars();
    while let Some(character) = characters.next() {
        if character == '\'' && characters.next() != Some('\'') {
            return None;
        }
        unescaped.push(character);
    }

    if unescaped
        .chars()
        .all(|character| character.is_ascii_alphanumeric() || character == '_')
    {
        Some(unescaped)
    } else {
        Some(format!("'{}'", unescaped.replace('\'', "''")))
    }
}

fn one_based_coordinate(value: usize, dimension: &str) -> OdfResult<u32> {
    one_based_u32(value).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "ODS {dimension} index {value} exceeds the shared API limit"
        ))
    })
}

fn one_based_u32(value: usize) -> Option<u32> {
    u32::try_from(value).ok()?.checked_add(1)
}

fn validate_one_based(value: u32, dimension: &str) -> SheetResult<()> {
    if value == 0 {
        return Err(sheet_error(format!(
            "ODS {dimension} coordinates are 1-based"
        )));
    }
    Ok(())
}

fn sheet_error(message: impl Into<String>) -> Box<dyn std::error::Error + Send + Sync> {
    Box::new(Error::InvalidFormat(message.into()))
}

fn parse_a1_coordinate(value: &str) -> SheetResult<(u32, u32)> {
    let value = value.trim();
    let mut characters = value.chars().peekable();
    if characters.peek() == Some(&'$') {
        characters.next();
    }

    let mut column = 0u32;
    let mut saw_column = false;
    while let Some(character) = characters.peek().copied() {
        if !character.is_ascii_alphabetic() {
            break;
        }
        saw_column = true;
        let digit = u32::from(character.to_ascii_uppercase() as u8 - b'A' + 1);
        column = column
            .checked_mul(26)
            .and_then(|value| value.checked_add(digit))
            .ok_or_else(|| {
                sheet_error(format!(
                    "ODS cell coordinate '{value}' has an overflowing column"
                ))
            })?;
        characters.next();
    }
    if !saw_column {
        return Err(sheet_error(format!(
            "invalid ODS cell coordinate '{value}'"
        )));
    }
    if characters.peek() == Some(&'$') {
        characters.next();
    }

    let mut row = 0u32;
    let mut saw_row = false;
    for character in characters {
        if !character.is_ascii_digit() {
            return Err(sheet_error(format!(
                "invalid ODS cell coordinate '{value}'"
            )));
        }
        saw_row = true;
        let digit = character
            .to_digit(10)
            .expect("ASCII digit has a numeric value");
        row = row
            .checked_mul(10)
            .and_then(|value| value.checked_add(digit))
            .ok_or_else(|| {
                sheet_error(format!(
                    "ODS cell coordinate '{value}' has an overflowing row"
                ))
            })?;
    }
    if !saw_row || row == 0 || column == 0 {
        return Err(sheet_error(format!(
            "invalid ODS cell coordinate '{value}'"
        )));
    }
    Ok((row, column))
}

fn a1_coordinate(row: u32, column: u32) -> String {
    format!("{}{}", column_to_letters(column), row)
}

fn column_to_letters(mut column: u32) -> String {
    let mut letters = String::new();
    while column != 0 {
        column -= 1;
        letters.insert(0, char::from(b'A' + (column % 26) as u8));
        column /= 26;
    }
    letters
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::SpreadsheetBuilder;

    #[test]
    fn normalizes_common_open_formula_syntax_without_touching_strings() {
        assert_eq!(
            normalize_open_formula("of:=IF([.A1]>0;[.B1];\"[.C1];\")"),
            "IF(A1>0,B1,\"[.C1];\")"
        );
        assert_eq!(
            normalize_open_formula("of:=SUM([$Inputs.$A$1:.$B$2])"),
            "SUM(Inputs!$A$1:$B$2)"
        );
        assert_eq!(
            normalize_open_formula("of:=SUM(['My Sheet'.$A$1:.$B$2])"),
            "SUM('My Sheet'!$A$1:$B$2)"
        );
        assert_eq!(
            normalize_open_formula("of:=SUM(['file:///never.ods'#$Sheet1.$A$1])"),
            "SUM(['file:///never.ods'#$Sheet1.$A$1])"
        );
        assert_eq!(
            normalize_open_formula("of:=['Bob''s'.$A$1]"),
            "'Bob''s'!$A$1"
        );
        assert_eq!(
            normalize_open_formula("of:=SUM([.NamedRange])"),
            "SUM([.NamedRange])"
        );
    }

    #[test]
    fn exposes_ods_values_through_the_shared_workbook_traits() {
        let mut builder = SpreadsheetBuilder::new();
        builder.add_sheet("Sheet1").unwrap();
        builder.add_row_with_numbers(&[2.0, 3.0]).unwrap();
        builder
            .set_cell_formula(1, 0, "of:=SUM([.A1:.B1])")
            .unwrap();

        let mut spreadsheet = Spreadsheet::from_bytes(builder.build().unwrap()).unwrap();
        let workbook = spreadsheet.evaluation_workbook().unwrap();
        assert_eq!(workbook.worksheet_names(), &["Sheet1".to_string()]);

        let sheet = workbook.worksheet_by_name("Sheet1").unwrap();
        assert_eq!(sheet.dimensions(), Some((1, 1, 2, 2)));
        assert!(matches!(
            sheet.cell_value(1, 1).unwrap().as_ref(),
            CellValue::Float(value) if *value == 2.0
        ));
        let formula = sheet.cell_value(2, 1).unwrap();
        assert!(
            matches!(
                formula.as_ref(),
                CellValue::Formula { formula, cached_value: None, .. } if formula == "SUM(A1:B1)"
            ),
            "{formula:?}"
        );
        assert_eq!(sheet.cell_by_coordinate("B1").unwrap().coordinate(), "B1");
        assert!(sheet.cell_value(0, 1).is_err());
    }

    #[tokio::test]
    async fn evaluates_uncached_open_formula_cells() {
        let mut builder = SpreadsheetBuilder::new();
        builder.add_sheet("Sheet1").unwrap();
        builder.add_row_with_numbers(&[2.0, 3.0]).unwrap();
        builder
            .set_cell_formula(1, 0, "of:=SUM([.A1:.B1])")
            .unwrap();
        builder
            .set_cell_formula(1, 1, "of:=IF([.A1]>0;[.B1];0)")
            .unwrap();

        let workbook = OdsWorkbook::from_bytes(builder.build().unwrap()).unwrap();
        let evaluator = litchi_eval::FormulaEvaluator::new(&workbook);

        assert!(matches!(
            evaluator.evaluate_cell("Sheet1", 2, 1).await.unwrap(),
            CellValue::Float(value) if value == 5.0
        ));
        assert!(matches!(
            evaluator.evaluate_cell("Sheet1", 2, 2).await.unwrap(),
            CellValue::Float(value) if value == 3.0
        ));
    }
}
