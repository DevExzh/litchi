use std::collections::{HashMap, HashSet};

use litchi_core::sheet::{CellValue, Result};

const MAX_ROW: u32 = 1_048_576;
const MAX_COLUMN: u32 = 16_384;

#[derive(Debug, Clone)]
pub(crate) struct SharedFormulaCell {
    pub(crate) row: u32,
    pub(crate) column: u32,
    pub(crate) index: u32,
    pub(crate) reference: Option<String>,
    pub(crate) formula: String,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct CellRange {
    first_row: u32,
    first_column: u32,
    last_row: u32,
    last_column: u32,
}

impl CellRange {
    fn parse(value: &str) -> Result<Self> {
        let (first, last) = value.split_once(':').unwrap_or((value, value));
        if last.contains(':') {
            return Err(format!("invalid shared formula range '{value}'").into());
        }
        let (first_row, first_column) = parse_cell_position(first)
            .ok_or_else(|| format!("invalid shared formula range '{value}'"))?;
        let (last_row, last_column) = parse_cell_position(last)
            .ok_or_else(|| format!("invalid shared formula range '{value}'"))?;
        if first_row > last_row || first_column > last_column {
            return Err(format!("reversed shared formula range '{value}'").into());
        }
        Ok(Self {
            first_row,
            first_column,
            last_row,
            last_column,
        })
    }

    fn contains(self, row: u32, column: u32) -> bool {
        (self.first_row..=self.last_row).contains(&row)
            && (self.first_column..=self.last_column).contains(&column)
    }

}

#[derive(Debug)]
struct Master {
    row: u32,
    column: u32,
    range: CellRange,
    formula: String,
}

pub(crate) fn resolve_shared_formulas(
    cells: &mut HashMap<u32, HashMap<u32, CellValue>>,
    shared_cells: &[SharedFormulaCell],
) -> Result<()> {
    if shared_cells.is_empty() {
        return Ok(());
    }

    let mut occupied = HashSet::with_capacity(shared_cells.len());
    for cell in shared_cells {
        if !occupied.insert((cell.row, cell.column)) {
            return Err(format!(
                "duplicate or ambiguous shared formula membership at row {}, column {}",
                cell.row, cell.column
            )
            .into());
        }
    }

    let mut masters = HashMap::<u32, Master>::new();
    for cell in shared_cells {
        if cell.reference.is_none() && cell.formula.is_empty() {
            continue;
        }
        let reference = cell.reference.as_deref().ok_or_else(|| {
            format!(
                "shared formula master at row {}, column {} is missing ref",
                cell.row, cell.column
            )
        })?;
        if cell.formula.is_empty() {
            return Err(format!(
                "shared formula master at row {}, column {} has no formula",
                cell.row, cell.column
            )
            .into());
        }
        let range = CellRange::parse(reference)?;
        if (cell.row, cell.column) != (range.first_row, range.first_column) {
            return Err(format!(
                "shared formula master at row {}, column {} is not the first cell of ref '{reference}'",
                cell.row, cell.column
            )
            .into());
        }
        let master = Master {
            row: cell.row,
            column: cell.column,
            range,
            formula: cell.formula.clone(),
        };
        if masters.insert(cell.index, master).is_some() {
            return Err(format!("duplicate shared formula master for si={}", cell.index).into());
        }
    }

    for cell in shared_cells {
        if !masters.contains_key(&cell.index) {
            return Err(format!(
                "shared formula at row {}, column {} has no master for si={}",
                cell.row, cell.column, cell.index
            )
            .into());
        }
    }

    for cell in shared_cells {
        let master = &masters[&cell.index];
        if !master.range.contains(cell.row, cell.column) {
            return Err(format!(
                "shared formula at row {}, column {} lies outside master range for si={}",
                cell.row, cell.column, cell.index
            )
            .into());
        }
        let is_master = (cell.row, cell.column) == (master.row, master.column);
        if !is_master && (!cell.formula.is_empty() || cell.reference.is_some()) {
            return Err(format!(
                "shared formula follower at row {}, column {} contains master data",
                cell.row, cell.column
            )
            .into());
        }
        if is_master {
            continue;
        }
        let value = cells
            .get_mut(&cell.row)
            .and_then(|row| row.get_mut(&cell.column))
            .ok_or_else(|| {
                format!(
                    "missing materialized shared formula cell at row {}, column {}",
                    cell.row, cell.column
                )
            })?;
        let CellValue::Formula { formula, .. } = value else {
            return Err(format!(
                "shared formula metadata points to a non-formula cell at row {}, column {}",
                cell.row, cell.column
            )
            .into());
        };
        *formula = translate_formula(
            &master.formula,
            master.row,
            master.column,
            cell.row,
            cell.column,
        );
    }

    Ok(())
}

fn parse_cell_position(value: &str) -> Option<(u32, u32)> {
    let bytes = value.as_bytes();
    let mut index = usize::from(bytes.first() == Some(&b'$'));
    let column_start = index;
    while bytes.get(index).is_some_and(u8::is_ascii_alphabetic) {
        index += 1;
    }
    if index == column_start {
        return None;
    }
    let column = decode_column(&value[column_start..index])?;
    if bytes.get(index) == Some(&b'$') {
        index += 1;
    }
    let row_start = index;
    while bytes.get(index).is_some_and(u8::is_ascii_digit) {
        index += 1;
    }
    if index == row_start || index != bytes.len() {
        return None;
    }
    let row = value[row_start..index].parse::<u32>().ok()?;
    (row != 0 && row <= MAX_ROW && column <= MAX_COLUMN).then_some((row, column))
}

fn decode_column(value: &str) -> Option<u32> {
    let mut column = 0u32;
    for byte in value.bytes() {
        column = column
            .checked_mul(26)?
            .checked_add(u32::from(byte.to_ascii_uppercase() - b'A' + 1))?;
    }
    (column != 0).then_some(column)
}

fn encode_column(mut column: u32, output: &mut String) {
    let start = output.len();
    while column != 0 {
        column -= 1;
        output.push(char::from(b'A' + (column % 26) as u8));
        column /= 26;
    }
    let suffix = output.split_off(start);
    output.extend(suffix.chars().rev());
}

#[derive(Debug, Clone, Copy)]
struct Axis {
    value: u32,
    absolute: bool,
}

#[derive(Debug, Clone, Copy)]
struct CellReference {
    row: Axis,
    column: Axis,
}

#[derive(Debug, Clone, Copy)]
enum Reference {
    Cell(CellReference),
    Area(CellReference, CellReference),
    Columns(Axis, Axis),
    Rows(Axis, Axis),
}

struct ParsedReference<'a> {
    prefix: &'a str,
    reference: Reference,
    end: usize,
}

fn translate_formula(
    formula: &str,
    origin_row: u32,
    origin_column: u32,
    target_row: u32,
    target_column: u32,
) -> String {
    let row_delta = i64::from(target_row) - i64::from(origin_row);
    let column_delta = i64::from(target_column) - i64::from(origin_column);
    let bytes = formula.as_bytes();
    let mut output = String::with_capacity(formula.len());
    let mut index = 0usize;

    while index < bytes.len() {
        if bytes[index] == b'"' {
            let end = quoted_string_end(bytes, index);
            output.push_str(&formula[index..end]);
            index = end;
            continue;
        }
        if let Some(parsed) = parse_reference(formula, index) {
            output.push_str(parsed.prefix);
            render_reference(
                parsed.reference,
                row_delta,
                column_delta,
                &mut output,
            );
            index = parsed.end;
            continue;
        }
        if bytes[index] == b'[' {
            let end = bracket_end(bytes, index);
            output.push_str(&formula[index..end]);
            index = end;
            continue;
        }
        let character = formula[index..].chars().next().expect("valid UTF-8");
        output.push(character);
        index += character.len_utf8();
    }
    output
}

fn quoted_string_end(bytes: &[u8], start: usize) -> usize {
    let mut index = start + 1;
    while index < bytes.len() {
        if bytes[index] == b'"' {
            if bytes.get(index + 1) == Some(&b'"') {
                index += 2;
            } else {
                return index + 1;
            }
        } else {
            index += 1;
        }
    }
    bytes.len()
}

fn bracket_end(bytes: &[u8], start: usize) -> usize {
    let mut depth = 0usize;
    let mut index = start;
    while index < bytes.len() {
        match bytes[index] {
            b'[' => depth += 1,
            b']' => {
                depth -= 1;
                if depth == 0 {
                    return index + 1;
                }
            }
            _ => {}
        }
        index += 1;
    }
    bytes.len()
}

fn parse_reference(formula: &str, start: usize) -> Option<ParsedReference<'_>> {
    let bytes = formula.as_bytes();
    if start != 0 && is_identifier_byte(bytes[start - 1]) {
        return None;
    }

    let reference_start = parse_prefix(formula, start).unwrap_or(start);
    let prefix = &formula[start..reference_start];
    let (reference, end) = parse_reference_body(formula, reference_start)?;
    if bytes.get(end).is_some_and(|byte| is_identifier_byte(*byte) || *byte == b'(') {
        return None;
    }
    Some(ParsedReference {
        prefix,
        reference,
        end,
    })
}

fn parse_prefix(formula: &str, start: usize) -> Option<usize> {
    let bytes = formula.as_bytes();
    if bytes.get(start) == Some(&b'\'') {
        let mut index = start + 1;
        while index < bytes.len() {
            if bytes[index] == b'\'' {
                if bytes.get(index + 1) == Some(&b'\'') {
                    index += 2;
                    continue;
                }
                return (bytes.get(index + 1) == Some(&b'!')).then_some(index + 2);
            }
            index += 1;
        }
        return None;
    }

    let mut index = start;
    let mut bracket_depth = 0usize;
    while index < bytes.len() {
        match bytes[index] {
            b'[' => bracket_depth += 1,
            b']' if bracket_depth != 0 => bracket_depth -= 1,
            b'!' if bracket_depth == 0 && index != start => return Some(index + 1),
            byte if bracket_depth == 0 && !is_prefix_byte(byte) => return None,
            _ => {}
        }
        index += 1;
    }
    None
}

fn is_prefix_byte(byte: u8) -> bool {
    byte.is_ascii_alphanumeric()
        || matches!(byte, b'_' | b'.' | b':' | b'\\' | b'/' | b'[' | b']' | b'-')
}

fn is_identifier_byte(byte: u8) -> bool {
    byte.is_ascii_alphanumeric() || matches!(byte, b'_' | b'.')
}

fn parse_reference_body(formula: &str, start: usize) -> Option<(Reference, usize)> {
    if let Some((first, first_end)) = parse_cell_reference(formula, start) {
        if formula.as_bytes().get(first_end) == Some(&b':') {
            let (last, end) = parse_cell_reference(formula, first_end + 1)?;
            return Some((Reference::Area(first, last), end));
        }
        return Some((Reference::Cell(first), first_end));
    }
    if let Some((first, first_end)) = parse_column_reference(formula, start)
        && formula.as_bytes().get(first_end) == Some(&b':')
    {
        let (last, end) = parse_column_reference(formula, first_end + 1)?;
        return Some((Reference::Columns(first, last), end));
    }
    if let Some((first, first_end)) = parse_row_reference(formula, start)
        && formula.as_bytes().get(first_end) == Some(&b':')
    {
        let (last, end) = parse_row_reference(formula, first_end + 1)?;
        return Some((Reference::Rows(first, last), end));
    }
    None
}

fn parse_cell_reference(formula: &str, start: usize) -> Option<(CellReference, usize)> {
    let (column, index) = parse_column_reference(formula, start)?;
    let (row, end) = parse_row_reference(formula, index)?;
    Some((CellReference { row, column }, end))
}

fn parse_column_reference(formula: &str, start: usize) -> Option<(Axis, usize)> {
    let bytes = formula.as_bytes();
    let absolute = bytes.get(start) == Some(&b'$');
    let mut index = start + usize::from(absolute);
    let letters_start = index;
    while bytes.get(index).is_some_and(u8::is_ascii_alphabetic) && index - letters_start < 3 {
        index += 1;
    }
    if index == letters_start || bytes.get(index).is_some_and(u8::is_ascii_alphabetic) {
        return None;
    }
    let value = decode_column(&formula[letters_start..index])?;
    (value <= MAX_COLUMN).then_some((Axis { value, absolute }, index))
}

fn parse_row_reference(formula: &str, start: usize) -> Option<(Axis, usize)> {
    let bytes = formula.as_bytes();
    let absolute = bytes.get(start) == Some(&b'$');
    let mut index = start + usize::from(absolute);
    let digits_start = index;
    while bytes.get(index).is_some_and(u8::is_ascii_digit) {
        index += 1;
    }
    if index == digits_start {
        return None;
    }
    let value = formula[digits_start..index].parse::<u32>().ok()?;
    (value != 0 && value <= MAX_ROW).then_some((Axis { value, absolute }, index))
}

fn shifted(axis: Axis, delta: i64, maximum: u32) -> Option<u32> {
    if axis.absolute {
        return Some(axis.value);
    }
    let value = i64::from(axis.value).checked_add(delta)?;
    (value >= 1 && value <= i64::from(maximum)).then_some(value as u32)
}

fn render_reference(reference: Reference, row_delta: i64, column_delta: i64, output: &mut String) {
    match reference {
        Reference::Cell(cell) => render_cell(cell, row_delta, column_delta, output),
        Reference::Area(first, last) => {
            render_cell(first, row_delta, column_delta, output);
            output.push(':');
            render_cell(last, row_delta, column_delta, output);
        }
        Reference::Columns(first, last) => {
            render_column(first, column_delta, output);
            output.push(':');
            render_column(last, column_delta, output);
        }
        Reference::Rows(first, last) => {
            render_row(first, row_delta, output);
            output.push(':');
            render_row(last, row_delta, output);
        }
    }
}

fn render_cell(cell: CellReference, row_delta: i64, column_delta: i64, output: &mut String) {
    let Some(column) = shifted(cell.column, column_delta, MAX_COLUMN) else {
        output.push_str("#REF!");
        return;
    };
    let Some(row) = shifted(cell.row, row_delta, MAX_ROW) else {
        output.push_str("#REF!");
        return;
    };
    if cell.column.absolute {
        output.push('$');
    }
    encode_column(column, output);
    if cell.row.absolute {
        output.push('$');
    }
    output.push_str(&row.to_string());
}

fn render_column(column: Axis, delta: i64, output: &mut String) {
    let Some(value) = shifted(column, delta, MAX_COLUMN) else {
        output.push_str("#REF!");
        return;
    };
    if column.absolute {
        output.push('$');
    }
    encode_column(value, output);
}

fn render_row(row: Axis, delta: i64, output: &mut String) {
    let Some(value) = shifted(row, delta, MAX_ROW) else {
        output.push_str("#REF!");
        return;
    };
    if row.absolute {
        output.push('$');
    }
    output.push_str(&value.to_string());
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xlsx::parsers::worksheet_parser::parse_worksheet_data;
    use litchi_opc::{OpcPackage, PackURI};

    fn formula(value: &CellValue) -> (&str, Option<&CellValue>) {
        let CellValue::Formula {
            formula,
            cached_value,
            ..
        } = value
        else {
            panic!("expected formula cell")
        };
        (formula, cached_value.as_deref())
    }

    fn parse_fixture(bytes: &[u8], sheet: &str) -> Result<crate::xlsx::parsers::worksheet_parser::ParsedWorksheetData> {
        let package = OpcPackage::from_bytes(bytes)?;
        let part = package.get_part(&PackURI::new(sheet)?)?;
        parse_worksheet_data(std::str::from_utf8(part.blob())?)
    }

    #[test]
    fn translates_reference_forms_without_rewriting_other_tokens() {
        let input = r#"SUM(A1,$A1,A$1,$A$1,A1:B2,A:C,1:3,'My Sheet'!A1,Sheet1:Sheet3!A1,[1]Sheet1!A1,INDIRECT("A1"),LOG10(2),Table1[A1])"#;
        assert_eq!(
            translate_formula(input, 1, 1, 3, 2),
            r#"SUM(B3,$A3,B$1,$A$1,B3:C4,B:D,3:5,'My Sheet'!B3,Sheet1:Sheet3!B3,[1]Sheet1!B3,INDIRECT("A1"),LOG10(2),Table1[A1])"#
        );
    }

    #[test]
    fn checked_reference_shift_produces_ref_error() {
        assert_eq!(translate_formula("A1:$A$1", 2, 2, 1, 1), "#REF!:$A$1");
        assert_eq!(
            translate_formula("XFD1048576", 1, 1, 2, 2),
            "#REF!"
        );
    }

    #[test]
    fn validates_missing_duplicate_and_outside_groups() {
        let xml = |body: &str| {
            format!(r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>{body}</sheetData></worksheet>"#)
        };
        for body in [
            r#"<row r="1"><c r="A1"><f t="shared" si="0"/></c></row>"#,
            r#"<row r="1"><c r="A1"><f t="shared" ref="A1:A2" si="0">A1</f></c><c r="B1"><f t="shared" ref="B1:B2" si="0">B1</f></c></row>"#,
            r#"<row r="1"><c r="A1"><f t="shared" ref="A1:A1" si="0">A1</f></c><c r="B1"><f t="shared" si="0"/></c></row>"#,
        ] {
            assert!(parse_worksheet_data(&xml(body)).is_err(), "accepted {body}");
        }
    }

    #[test]
    fn intersecting_declared_ranges_use_explicit_si_membership() {
        let xml = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>
            <row r="1">
                <c r="A1"><f t="shared" ref="A1:B2" si="0">A1</f><v>1</v></c>
                <c r="B1"><f t="shared" ref="B1:C2" si="1">B1</f><v>2</v></c>
                <c r="C1"><f t="shared" si="1"/><v>3</v></c>
            </row>
            <row r="2">
                <c r="A2"><f t="shared" si="0"/><v>4</v></c>
                <c r="B2"><f>B2+1</f><v>5</v></c>
            </row>
        </sheetData></worksheet>"#;
        let data = parse_worksheet_data(xml).unwrap();
        assert_eq!(formula(&data.cells[&2][&1]).0, "A2");
        assert_eq!(formula(&data.cells[&1][&3]).0, "C1");
        assert_eq!(formula(&data.cells[&2][&2]).0, "B2+1");
    }

    #[test]
    fn rejects_duplicate_or_ambiguous_actual_membership() {
        let mut cells = HashMap::from([
            (1, HashMap::from([
                (1, CellValue::Formula { formula: "A1".into(), cached_value: None, is_array: false, array_range: None }),
                (2, CellValue::Formula { formula: "B1".into(), cached_value: None, is_array: false, array_range: None }),
            ])),
            (2, HashMap::from([
                (1, CellValue::Formula { formula: String::new(), cached_value: None, is_array: false, array_range: None }),
            ])),
        ]);
        let shared = vec![
            SharedFormulaCell { row: 1, column: 1, index: 0, reference: Some("A1:A2".into()), formula: "A1".into() },
            SharedFormulaCell { row: 1, column: 2, index: 1, reference: Some("B1:B2".into()), formula: "B1".into() },
            SharedFormulaCell { row: 2, column: 1, index: 0, reference: None, formula: String::new() },
            SharedFormulaCell { row: 2, column: 1, index: 1, reference: None, formula: String::new() },
        ];
        assert!(resolve_shared_formulas(&mut cells, &shared).is_err());
    }

    #[test]
    fn preserves_cached_values_and_explicit_formulas() {
        let xml = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>
            <row r="1"><c r="A1" t="str"><f t="shared" ref="A1:A3" si="7">B1</f><v>one</v></c></row>
            <row r="2"><c r="A2" t="str"><f>B2+1</f><v>two</v></c></row>
            <row r="3"><c r="A3" t="str"><f t="shared" si="7"/><v>three</v></c></row>
        </sheetData></worksheet>"#;
        let data = parse_worksheet_data(xml).unwrap();
        assert_eq!(formula(&data.cells[&1][&1]), ("B1", Some(&CellValue::String("one".into()))));
        assert_eq!(formula(&data.cells[&2][&1]).0, "B2+1");
        assert_eq!(formula(&data.cells[&3][&1]), ("B3", Some(&CellValue::String("three".into()))));
    }

    #[test]
    fn poi_and_libreoffice_fixture_oracles_expand() {
        let poi = parse_fixture(
            include_bytes!("../../../../3rdparty/poi/test-data/spreadsheet/shared_formulas.xlsx"),
            "/xl/worksheets/sheet1.xml",
        )
        .unwrap();
        assert_eq!(formula(&poi.cells[&3][&1]).0, "B3");
        assert_eq!(formula(&poi.cells[&41][&1]).0, "B41");

        let shifted = parse_fixture(
            include_bytes!("../../../../3rdparty/poi/test-data/spreadsheet/TestShiftRowSharedFormula.xlsx"),
            "/xl/worksheets/sheet1.xml",
        )
        .unwrap();
        assert_eq!(formula(&shifted.cells[&5][&5]).0, "SUM(E2:E4)");

        let basic = parse_fixture(
            include_bytes!("../../../../3rdparty/libreoffice-core/sc/qa/unit/data/xlsx/shared-formula/basic.xlsx"),
            "/xl/worksheets/sheet1.xml",
        )
        .unwrap();
        assert_eq!(formula(&basic.cells[&4][&2]).0, "A4*10");
        assert_eq!(formula(&basic.cells[&19][&2]).0, "A19*10");

        let updated = parse_fixture(
            include_bytes!("../../../../3rdparty/libreoffice-core/sc/qa/unit/data/xlsx/shared-formula/refupdate.xlsx"),
            "/xl/worksheets/sheet1.xml",
        )
        .unwrap();
        assert_eq!(formula(&updated.cells[&1][&3]).0, "C30+1");
        assert_eq!(formula(&updated.cells[&1][&5]).0, "E30+1");

        let text = parse_fixture(
            include_bytes!("../../../../3rdparty/libreoffice-core/sc/qa/unit/data/xlsx/shared-formula/text-results.xlsx"),
            "/xl/worksheets/sheet1.xml",
        )
        .unwrap();
        assert_eq!(formula(&text.cells[&4][&2]), ("A4", Some(&CellValue::String("C".into()))));

        parse_fixture(
            include_bytes!("../../../../3rdparty/libreoffice-core/sc/qa/unit/data/xlsx/shared-formula/3d-reference.xlsx"),
            "/xl/worksheets/sheet1.xml",
        )
        .unwrap();
    }

    #[test]
    fn poi_intersecting_range_regression_fixtures_resolve_by_si() {
        for bytes in [
            include_bytes!("../../../../3rdparty/poi/test-data/spreadsheet/testSharedFormulasSetBlank.xlsx").as_slice(),
            include_bytes!("../../../../3rdparty/poi/test-data/spreadsheet/testSharedFormulasRangeSetBlankBug.xlsx").as_slice(),
        ] {
            parse_fixture(bytes, "/xl/worksheets/sheet1.xml").unwrap();
        }
    }
}
