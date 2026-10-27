//! Differential test of Numbers row and column edits on tables with
//! formulas (record 0318): the focused `litchi-numbers` topology transaction
//! against the retired host editor's formula-coordinate shift.
//!
//! Every corpus table is edited both ways at several section positions. The
//! two outputs are compared object by object (every IWA object's messages,
//! by identifier, in every component but the package metadata). A case the
//! host editor refuses must be refused by the leaf too, with a typed reason.

use std::collections::BTreeMap;
use std::path::PathBuf;

use litchi_iwa::numbers::{self, UnverifiedAxisChange};
use litchi_iwa_archive::iwa::{Archive, SnappyStream};
use litchi_iwa_archive::package::{Catalog, EntryEdit};
use litchi_iwa_archive::{ComponentCatalog, Limits};
use litchi_numbers::formula::{CellReference, Expression};
use litchi_numbers::table::cells::Input;
use litchi_numbers::table::headers::{Count, Settings};
use litchi_numbers::table::topology::{
    ColumnDeletion, ColumnInsertion, Error as TopologyError, RowDeletion, RowInsertion,
};
use litchi_numbers::{At, Package, TableSelector};

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

fn corpus() -> TestResult<Vec<PathBuf>> {
    let mut paths = Vec::new();
    for directory in [
        "../../test-data/iwork/numbers",
        "../../test-data/iwork/synthetic/numbers",
    ] {
        let root = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join(directory);
        for entry in std::fs::read_dir(root)? {
            let path = entry?.path();
            if path
                .extension()
                .is_some_and(|extension| extension == "numbers")
                && path.is_file()
            {
                paths.push(path);
            }
        }
    }
    paths.sort();
    Ok(paths)
}

/// The calculation-engine messages and formula lists, by identifier.
type Objects = BTreeMap<u64, Vec<(u32, Vec<u8>)>>;

/// `TSCE.FormulaOwnerDependenciesArchive`, `TSCE.CellRecordTileArchive`,
/// `TSCE.RangePrecedentsTileArchive`.
const CALC_TYPES: [u32; 3] = [4_008, 4_009, 4_010];
/// `TST.TableDataList`.
const LIST_TYPE: u32 = 6_005;

/// The `TST.TableDataList.listType` of a payload (field 1).
fn list_type(data: &[u8]) -> Option<u64> {
    let (&tag, rest) = data.split_first()?;
    if tag != 0x08 {
        return None;
    }
    let mut value = 0_u64;
    for (shift, byte) in rest.iter().enumerate().take(10) {
        value |= u64::from(byte & 0x7f) << (7 * shift);
        if byte & 0x80 == 0 {
            return Some(value);
        }
    }
    None
}

fn objects(bytes: &[u8]) -> TestResult<Objects> {
    let catalog = ComponentCatalog::from_bytes(bytes, Default::default())?;
    let mut output = BTreeMap::new();
    for component in catalog.iter() {
        for object in &component.archive().objects {
            let identifier = object
                .archive_info
                .identifier
                .ok_or("object without identifier")?;
            let messages: Vec<(u32, Vec<u8>)> = object
                .messages
                .iter()
                .filter(|message| {
                    CALC_TYPES.contains(&message.type_)
                        || (message.type_ == LIST_TYPE && list_type(&message.data) == Some(3))
                })
                .map(|message| (message.type_, message.data.clone()))
                .collect();
            if !messages.is_empty() {
                output.insert(identifier, messages);
            }
        }
    }
    Ok(output)
}

/// Every formula cell of every table, rendered: `(sheet, table, row,
/// column, text)`.
fn formula_texts(bytes: &[u8]) -> TestResult<Vec<(usize, usize, u32, u32, String)>> {
    let package = Package::from_bytes(bytes)?;
    let mut output = Vec::new();
    for (sheet_index, sheet) in package.sheets().iter().enumerate() {
        for (table_index, table) in sheet.tables().enumerate() {
            for cell in table.iter_cells() {
                if let litchi_numbers::cell::Value::Formula(text) = cell.value() {
                    output.push((
                        sheet_index,
                        table_index,
                        cell.position().row(),
                        cell.position().column(),
                        text.clone(),
                    ));
                }
            }
        }
    }
    output.sort();
    Ok(output)
}

#[derive(Debug, Clone, Copy)]
enum Op {
    InsertRow(RowInsertion),
    InsertColumn(ColumnInsertion),
    RemoveRow(RowDeletion),
    RemoveColumn(ColumnDeletion),
}

/// The body row and column counts, from the leaf's range refusal.
fn body_lengths(package: &Package, sheet: usize, table: usize) -> TestResult<(usize, usize)> {
    let length = |result: Result<_, TopologyError>| match result {
        Err(TopologyError::Refused(
            litchi_numbers::table::topology::Refusal::PositionOutOfRange { length, .. },
        )) => Ok(length),
        _ => Err("no range refusal"),
    };
    let rows = length(
        package
            .edit_table_topology(sheet, table)?
            .insert_row(RowInsertion::body(usize::MAX))
            .map(|_| ()),
    )?;
    let columns = length(
        package
            .edit_table_topology(sheet, table)?
            .insert_column(ColumnInsertion::body(usize::MAX))
            .map(|_| ()),
    )?;
    Ok((rows, columns))
}

fn ops(body_rows: usize, body_columns: usize) -> Vec<Op> {
    let mut ops = Vec::new();
    for index in [0, body_rows / 2, body_rows] {
        ops.push(Op::InsertRow(RowInsertion::body(index)));
    }
    for index in [0, body_columns / 2, body_columns] {
        ops.push(Op::InsertColumn(ColumnInsertion::body(index)));
    }
    for index in [0, body_rows / 2, body_rows.saturating_sub(1)] {
        ops.push(Op::RemoveRow(RowDeletion::body(index)));
    }
    for index in [0, body_columns / 2, body_columns.saturating_sub(1)] {
        ops.push(Op::RemoveColumn(ColumnDeletion::body(index)));
    }
    ops.dedup_by_key(|op| format!("{op:?}"));
    ops
}

fn host(source: &[u8], table: &str, op: Op) -> Result<Vec<u8>, String> {
    let mut editor = numbers::Draft::from_bytes(source).map_err(|error| error.to_string())?;
    let edit = match op {
        Op::InsertRow(insertion) => UnverifiedAxisChange::InsertRow(insertion),
        Op::InsertColumn(insertion) => UnverifiedAxisChange::InsertColumn(insertion),
        Op::RemoveRow(deletion) => UnverifiedAxisChange::RemoveRow(deletion),
        Op::RemoveColumn(deletion) => UnverifiedAxisChange::RemoveColumn(deletion),
    };
    editor
        .__edit_table_topology_unverified(TableSelector::name(table), edit)
        .map_err(|error| error.to_string())?;
    editor.to_bytes().map_err(|error| error.to_string())
}

fn leaf(package: &Package, sheet: usize, table: usize, op: Op) -> Result<Vec<u8>, TopologyError> {
    let edit = package.edit_table_topology(sheet, table)?;
    let edit = match op {
        Op::InsertRow(insertion) => edit.insert_row(insertion)?,
        Op::InsertColumn(insertion) => edit.insert_column(insertion)?,
        Op::RemoveRow(deletion) => edit.remove_row(deletion)?,
        Op::RemoveColumn(deletion) => edit.remove_column(deletion)?,
    };
    let commit = edit.commit()?;
    let mut bytes = Vec::new();
    commit
        .package()
        .write_to(&mut bytes)
        .map_err(|_error| TopologyError::InvalidSource)?;
    Ok(bytes)
}

/// For a removal, the source with the removed axis's stored cells cleared
/// by the leaf cell writer (the leaf transaction clears them the same way
/// before it shifts; the host editor's own clear leaves range dependencies
/// behind and then refuses). `None` when the clear is refused.
fn cleared_source(package: &Package, sheet: usize, table: usize, op: Op) -> Option<Vec<u8>> {
    let settings = package.edit_table_headers(sheet, table).ok()?.settings();
    let (row_axis, physical) = match op {
        Op::RemoveRow(RowDeletion::Body { index }) => (true, settings.header_row_count() + index),
        Op::RemoveColumn(ColumnDeletion::Body { index }) => {
            (false, settings.header_column_count() + index)
        },
        _ => return None,
    };
    let semantic = package.sheets()[sheet].tables().nth(table)?;
    let positions: Vec<At> = semantic
        .iter_cells()
        .map(|cell| cell.position())
        .filter(|position| {
            let coordinate = if row_axis {
                position.row()
            } else {
                position.column()
            };
            coordinate as usize == physical
        })
        .collect();
    if positions.is_empty() {
        return None;
    }
    let mut edit = package.edit_table_cells(sheet, table).ok()?;
    for position in positions {
        edit = edit.clear(position).ok()?;
    }
    write(&edit.commit().ok()?.into_package()).ok()
}

/// One A1 reference of a formula text: `(column absolute, column, row
/// absolute, row)`, zero-based, and the byte after it.
fn parse_reference(text: &str, start: usize) -> Option<((bool, u32, bool, u32), usize)> {
    let bytes = text.as_bytes();
    let mut cursor = start;
    let column_absolute = bytes.get(cursor) == Some(&b'$');
    if column_absolute {
        cursor += 1;
    }
    let letters_start = cursor;
    while cursor < bytes.len() && bytes[cursor].is_ascii_uppercase() {
        cursor += 1;
    }
    let letters = &text[letters_start..cursor];
    let row_absolute = bytes.get(cursor) == Some(&b'$');
    if row_absolute {
        cursor += 1;
    }
    let digits_start = cursor;
    while cursor < bytes.len() && bytes[cursor].is_ascii_digit() {
        cursor += 1;
    }
    let digits = &text[digits_start..cursor];
    let preceded = start > 0 && bytes[start - 1].is_ascii_alphanumeric();
    let followed =
        cursor < bytes.len() && (bytes[cursor].is_ascii_alphanumeric() || bytes[cursor] == b'(');
    if letters.is_empty() || digits.is_empty() || preceded || followed {
        return None;
    }
    let mut column = 0_u32;
    for letter in letters.bytes() {
        column = column * 26 + u32::from(letter - b'A' + 1);
    }
    let row = digits.parse::<u32>().ok()? - 1;
    Some(((column_absolute, column - 1, row_absolute, row), cursor))
}

fn format_reference(
    (column_absolute, column, row_absolute, row): (bool, u32, bool, u32),
) -> String {
    let mut name = String::new();
    let mut value = column + 1;
    while value > 0 {
        let digit = (value - 1) % 26;
        name.insert(0, char::from(b'A' + digit as u8));
        value = (value - 1) / 26;
    }
    format!(
        "{}{name}{}{}",
        if column_absolute { "$" } else { "" },
        if row_absolute { "$" } else { "" },
        row + 1
    )
}

/// Record 0329: every A1 reference of a local formula text after removing
/// `physical`: a reference to it reads `#REF!`, a range that loses it
/// shrinks, a range wholly on it reads `#REF!`.
fn remove_text(text: &str, row_axis: bool, physical: u32) -> String {
    let axis = |reference: (bool, u32, bool, u32)| if row_axis { reference.3 } else { reference.1 };
    let with = |reference: (bool, u32, bool, u32), value: u32| {
        if row_axis {
            (reference.0, reference.1, reference.2, value)
        } else {
            (reference.0, value, reference.2, reference.3)
        }
    };
    let shifted = |value: u32| if value > physical { value - 1 } else { value };
    let mut output = String::new();
    let mut index = 0;
    while index < text.len() {
        if let Some((first, after_first)) = parse_reference(text, index) {
            if text.as_bytes().get(after_first) == Some(&b':')
                && let Some((second, after_second)) = parse_reference(text, after_first + 1)
            {
                let (low, high) = (axis(first).min(axis(second)), axis(first).max(axis(second)));
                if low == physical && high == physical {
                    output.push_str("#REF!");
                } else {
                    let moved = |reference| {
                        let value = axis(reference);
                        with(
                            reference,
                            if value == physical && value == high {
                                physical - 1
                            } else if value == physical {
                                physical
                            } else {
                                shifted(value)
                            },
                        )
                    };
                    output.push_str(&format_reference(moved(first)));
                    output.push(':');
                    output.push_str(&format_reference(moved(second)));
                }
                index = after_second;
                continue;
            }
            if axis(first) == physical {
                output.push_str("#REF!");
            } else {
                output.push_str(&format_reference(with(first, shifted(axis(first)))));
            }
            index = after_first;
            continue;
        }
        let character = text[index..].chars().next().unwrap_or(' ');
        output.push(character);
        index += character.len_utf8();
    }
    output
}

/// Shift every A1 reference of a local formula text (no `::`) as the edit
/// moves its target: `(row axis, physical index, insert)`.
fn shift_text(text: &str, row_axis: bool, physical: u32, insert: bool) -> Option<String> {
    if text.contains("::") {
        return None;
    }
    if !insert {
        return Some(remove_text(text, row_axis, physical));
    }
    let bytes = text.as_bytes();
    let mut output = String::new();
    let mut index = 0;
    while index < bytes.len() {
        let start = index;
        let mut cursor = index;
        let column_absolute = bytes.get(cursor) == Some(&b'$');
        if column_absolute {
            cursor += 1;
        }
        let letters_start = cursor;
        while cursor < bytes.len() && bytes[cursor].is_ascii_uppercase() {
            cursor += 1;
        }
        let letters = &text[letters_start..cursor];
        let row_absolute = bytes.get(cursor) == Some(&b'$');
        if row_absolute {
            cursor += 1;
        }
        let digits_start = cursor;
        while cursor < bytes.len() && bytes[cursor].is_ascii_digit() {
            cursor += 1;
        }
        let digits = &text[digits_start..cursor];
        let preceded = start > 0 && (bytes[start - 1].is_ascii_alphanumeric());
        let followed = cursor < bytes.len()
            && (bytes[cursor].is_ascii_alphanumeric() || bytes[cursor] == b'(');
        if !letters.is_empty() && !digits.is_empty() && !preceded && !followed {
            let mut column = 0_u32;
            for letter in letters.bytes() {
                column = column * 26 + u32::from(letter - b'A' + 1);
            }
            let column = column - 1;
            let row = digits.parse::<u32>().ok()? - 1;
            let shift = |value: u32| -> Option<u32> {
                if value < physical {
                    Some(value)
                } else if insert {
                    Some(value + 1)
                } else if value > physical {
                    Some(value - 1)
                } else {
                    None
                }
            };
            let (row, column) = if row_axis {
                (shift(row)?, column)
            } else {
                (row, shift(column)?)
            };
            let mut name = String::new();
            let mut value = column + 1;
            while value > 0 {
                let digit = (value - 1) % 26;
                name.insert(0, char::from(b'A' + digit as u8));
                value = (value - 1) / 26;
            }
            if column_absolute {
                output.push('$');
            }
            output.push_str(&name);
            if row_absolute {
                output.push('$');
            }
            output.push_str(&(row + 1).to_string());
            index = cursor;
            continue;
        }
        let character = text[index..].chars().next()?;
        output.push(character);
        index += character.len_utf8();
    }
    Some(output)
}

/// Shift the references that follow `prefix` (`Sheet::Table::`) in another
/// table's formula text: a cell or the two endpoints of a range.
fn shift_prefixed(
    text: &str,
    prefix: &str,
    row_axis: bool,
    physical: u32,
    insert: bool,
) -> Option<String> {
    let mut output = String::new();
    let mut rest = text;
    while let Some(found) = rest.find(prefix) {
        output.push_str(&rest[..found + prefix.len()]);
        let after = &rest[found + prefix.len()..];
        let end = after
            .find(|character: char| {
                !(character.is_ascii_alphanumeric() || character == '$' || character == ':')
            })
            .unwrap_or(after.len());
        output.push_str(&shift_text(&after[..end], row_axis, physical, insert)?);
        rest = &after[end..];
    }
    output.push_str(rest);
    Some(output)
}

/// The edited table's local formulas after the edit, moved and with their
/// references shifted, and every other table's formulas with their
/// references into the edited table shifted: an oracle independent of both
/// implementations.
fn oracle_texts(
    source: &[u8],
    sheet: usize,
    table: usize,
    op: Op,
) -> TestResult<Vec<(usize, usize, u32, u32, String)>> {
    let package = Package::from_bytes(source)?;
    let settings = package.edit_table_headers(sheet, table)?.settings();
    let header_rows = settings.header_row_count() as u32;
    let header_columns = settings.header_column_count() as u32;
    let rows = package.sheets()[sheet]
        .tables()
        .nth(table)
        .ok_or("table")?
        .row_count();
    let footer_rows = settings.footer_row_count() as u32;
    let (row_axis, physical, insert) = match op {
        Op::InsertRow(RowInsertion::Body { index }) => (true, header_rows + index as u32, true),
        Op::InsertColumn(ColumnInsertion::Body { index }) => {
            (false, header_columns + index as u32, true)
        },
        Op::RemoveRow(RowDeletion::Body { index }) => (true, header_rows + index as u32, false),
        Op::RemoveColumn(ColumnDeletion::Body { index }) => {
            (false, header_columns + index as u32, false)
        },
        _ => return Ok(Vec::new()),
    };
    let sheet_name = package.sheets()[sheet].name().to_owned();
    let table_name = package.sheets()[sheet]
        .tables()
        .nth(table)
        .ok_or("table")?
        .name()
        .to_owned();
    let prefix = format!("{sheet_name}::{table_name}::");
    let mut output = Vec::new();
    for (sheet_index, table_index, row, column, text) in formula_texts(source)? {
        if (sheet_index, table_index) != (sheet, table) {
            if let Some(shifted) = shift_prefixed(&text, &prefix, row_axis, physical, insert) {
                output.push((sheet_index, table_index, row, column, shifted));
            }
            continue;
        }
        // Footer ranges grow and shrink at the body's end: not modelled.
        if footer_rows > 0 && row >= rows - footer_rows {
            continue;
        }
        let coordinate = if row_axis { row } else { column };
        let moved = if coordinate < physical {
            coordinate
        } else if insert {
            coordinate + 1
        } else if coordinate > physical {
            coordinate - 1
        } else {
            continue;
        };
        // A reference into another table keeps its text; only its host
        // cell moves.
        let shifted = if text.contains("::") && !text.contains(&prefix) {
            text
        } else if let Some(shifted) = shift_text(&text, row_axis, physical, insert) {
            shifted
        } else {
            continue;
        };
        let (row, column) = if row_axis {
            (moved, column)
        } else {
            (row, moved)
        };
        output.push((sheet_index, table_index, row, column, shifted));
    }
    Ok(output)
}

fn has_formulas(package: &Package, sheet: usize, table: usize) -> bool {
    package.sheets()[sheet]
        .tables()
        .nth(table)
        .is_some_and(|table| {
            table
                .iter_cells()
                .any(|cell| matches!(cell.value(), litchi_numbers::cell::Value::Formula(_)))
        })
}

// ---------------------------------------------------------------------------
// Packed dependency edges.
//
// This crate's formula writer packs the repeated edge coordinates of a
// dependency record; native Numbers writes them unpacked, and the host
// editor's shift only rewrites unpacked edges. The synthetic workbooks are
// compared with the host editor in their unpacked form, and the leaf's
// result on the packed original is compared with its result on the unpacked
// form after unpacking it the same way.
// ---------------------------------------------------------------------------

fn read_varint(data: &[u8], offset: &mut usize) -> TestResult<u64> {
    let mut value = 0_u64;
    for shift in 0..10 {
        let byte = *data.get(*offset).ok_or("truncated varint")?;
        *offset += 1;
        value |= u64::from(byte & 0x7f) << (7 * shift);
        if byte & 0x80 == 0 {
            return Ok(value);
        }
    }
    Err("overlong varint".into())
}

fn push_varint(output: &mut Vec<u8>, mut value: u64) {
    while value >= 0x80 {
        output.push((value as u8) | 0x80);
        value >>= 7;
    }
    output.push(value as u8);
}

/// `(number, wire type, raw field, payload)` of every field.
fn wire_fields(data: &[u8]) -> TestResult<Vec<(u32, u8, &[u8], &[u8])>> {
    let mut output = Vec::new();
    let mut offset = 0;
    while offset < data.len() {
        let start = offset;
        let key = read_varint(data, &mut offset)?;
        let number = u32::try_from(key >> 3)?;
        let wire_type = (key & 7) as u8;
        let payload_start;
        match wire_type {
            0 => {
                payload_start = offset;
                read_varint(data, &mut offset)?;
            },
            1 => {
                payload_start = offset;
                offset += 8;
            },
            2 => {
                let length = usize::try_from(read_varint(data, &mut offset)?)?;
                payload_start = offset;
                offset += length;
            },
            5 => {
                payload_start = offset;
                offset += 4;
            },
            _ => return Err("wire type".into()),
        }
        let raw = data.get(start..offset).ok_or("truncated field")?;
        let payload = data.get(payload_start..offset).ok_or("truncated field")?;
        output.push((number, wire_type, raw, payload));
    }
    Ok(output)
}

fn map_fields(
    data: &[u8],
    mut map: impl FnMut(u32, u8, &[u8]) -> TestResult<Option<Vec<u8>>>,
) -> TestResult<Vec<u8>> {
    let mut output = Vec::new();
    for (number, wire_type, raw, payload) in wire_fields(data)? {
        match map(number, wire_type, payload)? {
            Some(replacement) => output.extend(replacement),
            None => output.extend_from_slice(raw),
        }
    }
    Ok(output)
}

fn length_delimited(number: u32, payload: &[u8]) -> Vec<u8> {
    let mut output = Vec::new();
    push_varint(&mut output, (u64::from(number) << 3) | 2);
    push_varint(&mut output, payload.len() as u64);
    output.extend_from_slice(payload);
    output
}

fn unpack_edges(edges: &[u8]) -> TestResult<Vec<u8>> {
    map_fields(edges, |number, wire_type, payload| {
        if wire_type != 2 || !(1..=5).contains(&number) {
            return Ok(None);
        }
        let mut output = Vec::new();
        let mut offset = 0;
        while offset < payload.len() {
            let value = read_varint(payload, &mut offset)?;
            push_varint(&mut output, u64::from(number) << 3);
            push_varint(&mut output, value);
        }
        Ok(Some(output))
    })
}

fn unpack_record(record: &[u8]) -> TestResult<Vec<u8>> {
    map_fields(record, |number, wire_type, payload| {
        if number != 6 || wire_type != 2 {
            return Ok(None);
        }
        Ok(Some(length_delimited(6, &unpack_edges(payload)?)))
    })
}

fn unpack_records(container: &[u8], field: u32) -> TestResult<Vec<u8>> {
    map_fields(container, |number, wire_type, payload| {
        if number != field || wire_type != 2 {
            return Ok(None);
        }
        Ok(Some(length_delimited(field, &unpack_record(payload)?)))
    })
}

/// Unpack every dependency record's edges.
fn unpacked(bytes: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(bytes, Default::default())?;
    let mut rewritten = Vec::new();
    for member in catalog
        .iter()
        .filter(|member| member.name().ends_with(".iwa"))
    {
        let stream = SnappyStream::decompress(member.data(), Default::default())?;
        let mut archive = Archive::parse(stream.as_bytes(), Default::default())?;
        let mut changed = false;
        for object in &mut archive.objects {
            for message in &mut object.messages {
                let data = match message.type_ {
                    4_008 => map_fields(&message.data, |number, wire_type, payload| {
                        if number != 4 || wire_type != 2 {
                            return Ok(None);
                        }
                        Ok(Some(length_delimited(4, &unpack_records(payload, 1)?)))
                    })?,
                    4_009 => unpack_records(&message.data, 4)?,
                    _ => continue,
                };
                if data != message.data {
                    message.data = data;
                    changed = true;
                }
            }
        }
        if changed {
            rewritten.push((
                member.name().to_owned(),
                SnappyStream::compress(&archive.to_bytes(Default::default())?, Default::default())?,
            ));
        }
    }
    let edits: Vec<EntryEdit<'_>> = rewritten
        .iter()
        .map(|(name, data)| EntryEdit::new(name, data))
        .collect();
    Ok(catalog.reassemble_to_bytes(&edits, Limits::default())?)
}

fn write(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

/// Fill `cells` (`(A1, value)`) and `formulas` (`(A1, text)`) of one table.
fn fill(
    package: &Package,
    sheet: usize,
    table: usize,
    values: &[(&str, f64)],
    formulas: &[(&str, &str)],
) -> TestResult<Package> {
    let mut edit = package.edit_table_cells(sheet, table)?;
    for (address, value) in values {
        edit = edit.set_a1(address, Input::number(*value)?)?;
    }
    for (address, text) in formulas {
        edit = edit.set_formula_a1(address, Expression::parse(text)?)?;
    }
    Ok(edit
        .commit()
        .map_err(|error| format!("{formulas:?}: {error:?}"))?
        .into_package())
}

const VALUES: [(&str, f64); 12] = [
    ("B2", 1.0),
    ("C2", 2.0),
    ("D2", 3.0),
    ("B3", 4.0),
    ("C3", 5.0),
    ("D3", 6.0),
    ("B4", 7.0),
    ("C4", 8.0),
    ("B5", 9.0),
    ("C5", 10.0),
    ("B6", 11.0),
    ("C6", 12.0),
];

/// Synthetic workbooks with formulas, authored by the leaf: `(name, bytes)`.
fn synthetic() -> TestResult<Vec<(String, Vec<u8>)>> {
    let mut output = Vec::new();
    let local = fill(
        &Package::new(),
        0,
        0,
        &VALUES,
        &[
            ("E2", "B2+C2"),
            ("E3", "$B$3*2"),
            ("E4", "SUM(B2:D4)"),
            ("E5", "SUM($B2:B$6)"),
            ("E6", "B6-C5"),
            ("D7", "SUM(C$2:$C6)"),
            ("C8", "B2*$C$2"),
            ("A9", "SUM(B2:C3)+D2"),
        ],
    )?;
    output.push(("synthetic-local".to_owned(), write(&local)?));
    let single = fill(&Package::new(), 0, 0, &VALUES, &[("E7", "B3")])?;
    output.push(("synthetic-single".to_owned(), write(&single)?));
    // Relative formulas the leaf interns into shared list entries.
    let shared = fill(
        &Package::new(),
        0,
        0,
        &VALUES,
        &[("E7", "B2"), ("E8", "B3"), ("E9", "B3"), ("D8", "B2")],
    )?;
    output.push(("synthetic-shared".to_owned(), write(&shared)?));
    let cross = cross_table(false)?;
    output.push(("synthetic-cross-table".to_owned(), write(&cross)?));
    let incoming = cross_table(true)?;
    output.push(("synthetic-incoming".to_owned(), write(&incoming)?));
    let footer = package_with_footer()?;
    output.push(("synthetic-footer".to_owned(), write(&footer)?));
    Ok(output)
}

/// A workbook with a second table on the first sheet and a table on a
/// second sheet. `Table 1` reads both with relative, absolute and range
/// references; with `incoming`, the other tables read `Table 1` instead.
fn cross_table(incoming: bool) -> TestResult<Package> {
    let package = Package::new()
        .edit_table_lifecycle(0)?
        .add("Other", 8, 5)?
        .commit()?
        .into_package();
    let package = package
        .edit_sheet_lifecycle()
        .add("Second")?
        .commit()?
        .into_package();
    let package = if package.sheets()[1].table_count() == 0 {
        package
            .edit_table_lifecycle(1)?
            .add("Far", 8, 5)?
            .commit()?
            .into_package()
    } else {
        package
    };
    let package = fill(&package, 0, 0, &VALUES, &[])?;
    let package = fill(&package, 0, 1, &VALUES, &[])?;
    let package = fill(&package, 1, 0, &VALUES, &[])?;
    let (reader, targets): ((usize, usize), [(usize, usize); 2]) = if incoming {
        ((0, 1), [(0, 0), (0, 0)])
    } else {
        ((0, 0), [(0, 1), (1, 0)])
    };
    let mut package = package;
    for (index, target) in targets.iter().enumerate() {
        let edit = package.edit_table_cells(reader.0, reader.1)?;
        let table = edit.formula_table(target.0, target.1)?;
        let sum = |expression| Expression::function("SUM", [expression]);
        let row = 6 + index;
        let edit = edit
            .set_formula(
                At::new(row as u32, 1),
                Expression::table_cell(&table, CellReference::relative(2, 1)),
            )?
            .set_formula(
                At::new(row as u32, 2),
                Expression::table_cell(&table, CellReference::absolute(3, 2)),
            )?
            .set_formula(
                At::new(row as u32, 3),
                sum(Expression::table_range(
                    &table,
                    CellReference::relative(1, 1),
                    CellReference::relative(4, 2),
                )?)?,
            )?;
        package = edit
            .commit()
            .map_err(|error| format!("cross-table formulas: {error:?}"))?
            .into_package();
    }
    Ok(package)
}

/// A table with one footer row summing the body above it.
fn package_with_footer() -> TestResult<Package> {
    let package = Package::new();
    let rows = package.sheets()[0]
        .tables()
        .next()
        .ok_or("table")?
        .row_count();
    let footer = format!("B{rows}");
    let last_body = rows - 1;
    let sum = format!("SUM(B2:B{last_body})");
    let values: Vec<(String, f64)> = (2..=last_body)
        .map(|row| (format!("B{row}"), f64::from(row)))
        .collect();
    let values: Vec<(&str, f64)> = values.iter().map(|(a, v)| (a.as_str(), *v)).collect();
    let package = fill(&package, 0, 0, &values, &[(footer.as_str(), sum.as_str())])?;
    let settings = package.edit_table_headers(0, 0)?.settings();
    let settings = Settings {
        footer_rows: Some(Count::ONE),
        ..settings
    };
    Ok(package
        .edit_table_headers(0, 0)?
        .set(settings)
        .commit()?
        .into_package())
}

#[test]
fn leaf_topology_matches_the_host_editor_on_tables_with_formulas() -> TestResult {
    let mut summary = BTreeMap::<String, usize>::new();
    let mut mismatches = Vec::new();
    let mut refusals = Vec::new();
    let mut host_defects = Vec::new();
    let mut inputs = Vec::new();
    for path in corpus()? {
        let name = path
            .file_name()
            .and_then(|name| name.to_str())
            .unwrap_or("")
            .to_owned();
        inputs.push((name, std::fs::read(&path)?));
    }
    let packed = synthetic()?;
    for (name, bytes) in &packed {
        inputs.push((format!("{name} (unpacked)"), unpacked(bytes)?));
    }
    for (name, source) in &inputs {
        let source = source.as_slice();
        let package = Package::from_bytes(source)?;
        let name = name.as_str();
        if numbers::Draft::from_bytes(source).is_err() {
            *summary
                .entry(format!("skipped file: host cannot open {name}"))
                .or_default() += 1;
            continue;
        }
        for sheet in 0..package.sheets().len() {
            for (table, semantic) in package.sheets()[sheet].tables().enumerate() {
                let formulas = has_formulas(&package, sheet, table);
                let matching = package
                    .sheets()
                    .iter()
                    .flat_map(|sheet| sheet.tables())
                    .filter(|candidate| candidate.name() == semantic.name())
                    .count();
                if matching != 1 {
                    *summary.entry("skipped: ambiguous name".into()).or_default() += 1;
                    continue;
                }
                let (body_rows, body_columns) = body_lengths(&package, sheet, table)?;
                for op in ops(body_rows, body_columns) {
                    let host_source = cleared_source(&package, sheet, table, op);
                    let host_result = host(
                        host_source.as_deref().unwrap_or(source),
                        semantic.name(),
                        op,
                    );
                    let leaf_result = leaf(&package, sheet, table, op);
                    let formula_tag = if formulas { "formulas" } else { "plain" };
                    let key = match (&host_result, &leaf_result) {
                        (Ok(host_bytes), Ok(leaf_bytes)) => {
                            let host_objects = objects(host_bytes)?;
                            let leaf_objects = objects(leaf_bytes)?;
                            let mut differing = Vec::new();
                            for identifier in host_objects.keys().chain(leaf_objects.keys()) {
                                let left = host_objects.get(identifier);
                                let right = leaf_objects.get(identifier);
                                if left != right {
                                    let types: Vec<u32> = left
                                        .or(right)
                                        .map(|messages| messages.iter().map(|m| m.0).collect())
                                        .unwrap_or_default();
                                    differing.push((*identifier, types));
                                }
                            }
                            differing.sort();
                            differing.dedup();
                            let host_texts = formula_texts(host_bytes)?;
                            let leaf_texts = formula_texts(leaf_bytes)?;
                            let expected = oracle_texts(source, sheet, table, op)?;
                            let oracle_holds =
                                expected.iter().all(|cell| leaf_texts.contains(cell));
                            if !oracle_holds {
                                mismatches.push(format!(
                                    "{name} table {} {op:?}: leaf formula texts {leaf_texts:?} miss the shifted local texts {expected:?}",
                                    semantic.name()
                                ));
                            }
                            if host_texts != leaf_texts {
                                // The host editor moves a cell whose shifted
                                // formula equals another list entry onto that
                                // entry and then rewrites that entry in place
                                // for its own cells, so the moved cell takes
                                // their formula. Its result is wrong; the
                                // leaf's must then match the oracle.
                                let host_holds =
                                    expected.iter().all(|cell| host_texts.contains(cell));
                                if oracle_holds && !host_holds {
                                    host_defects.push(format!(
                                        "{name} table {} {op:?}: host {host_texts:?} leaf {leaf_texts:?}",
                                        semantic.name()
                                    ));
                                    *summary
                                        .entry(format!(
                                            "{formula_tag}: both ok, host formulas wrong, leaf matches the oracle"
                                        ))
                                        .or_default() += 1;
                                    continue;
                                }
                                mismatches.push(format!(
                                    "{name} table {} {op:?}: formula texts differ: host {host_texts:?} leaf {leaf_texts:?}",
                                    semantic.name()
                                ));
                            }
                            if differing.is_empty() {
                                format!("{formula_tag}: both ok, equal")
                            } else if oracle_holds
                                && host_texts == leaf_texts
                                && differing.iter().all(|(_, types)| types == &[6_005])
                            {
                                // The host editor moves a cell whose shifted
                                // formula equals another entry onto that
                                // entry; the leaf rewrites the cell's own
                                // entry in place. Same formulas, another
                                // formula-list shape.
                                format!(
                                    "{formula_tag}: both ok, same formulas, formula list interned differently"
                                )
                            } else {
                                mismatches.push(format!(
                                    "{name} table {} {op:?}: {differing:?}",
                                    semantic.name()
                                ));
                                format!("{formula_tag}: both ok, DIFFERENT")
                            }
                        },
                        (Ok(_), Err(leaf_error)) => {
                            refusals.push(format!(
                                "{name} table {} {op:?}: leaf {leaf_error:?}",
                                semantic.name()
                            ));
                            format!("{formula_tag}: host ok, leaf {leaf_error:?}")
                        },
                        (Err(host_error), Ok(leaf_bytes))
                            if host_error.contains("referenced by formulas in another table") =>
                        {
                            // Beyond the host editor: other tables' formulas
                            // that read the edited table move with it.
                            let expected = oracle_texts(source, sheet, table, op)?;
                            let leaf_texts = formula_texts(leaf_bytes)?;
                            if expected.is_empty()
                                || !expected.iter().all(|cell| leaf_texts.contains(cell))
                            {
                                mismatches.push(format!(
                                    "{name} table {} {op:?}: incoming: leaf {leaf_texts:?} oracle {expected:?}",
                                    semantic.name()
                                ));
                            }
                            format!(
                                "{formula_tag}: host refuses incoming references, leaf matches the oracle"
                            )
                        },
                        (Err(host_error), Ok(leaf_bytes))
                            if host_error.contains("a formula still references it")
                                || host_error.contains("invalid IWA wire data") =>
                        {
                            // Record 0329: native Numbers writes `#REF!`
                            // for a deleted reference; the host editor
                            // refuses it (and, on one corpus table, fails
                            // to rewrite its own clear). The leaf matches
                            // the oracle.
                            let expected = oracle_texts(source, sheet, table, op)?;
                            let leaf_texts = formula_texts(leaf_bytes)?;
                            if !expected.iter().all(|cell| leaf_texts.contains(cell)) {
                                mismatches.push(format!(
                                    "{name} table {} {op:?}: deleted: leaf {leaf_texts:?} oracle {expected:?}",
                                    semantic.name()
                                ));
                            }
                            format!(
                                "{formula_tag}: host refuses deleted references, leaf writes #REF! and matches the oracle"
                            )
                        },
                        (Err(host_error), Ok(_)) => {
                            mismatches.push(format!(
                                "{name} table {} {op:?}: host refused ({host_error}), leaf published",
                                semantic.name()
                            ));
                            format!("{formula_tag}: host refused, leaf ok")
                        },
                        (Err(host_error), Err(leaf_error)) => {
                            refusals.push(format!(
                                "{name} table {} {op:?}: both; host {host_error}; leaf {leaf_error:?}",
                                semantic.name()
                            ));
                            format!("{formula_tag}: both refused, leaf {leaf_error:?}")
                        },
                    };
                    *summary.entry(key).or_default() += 1;
                }
            }
        }
    }
    // The leaf on the packed originals publishes, after unpacking, what it
    // publishes on the unpacked forms.
    for (name, source) in &packed {
        let package = Package::from_bytes(source)?;
        let unpacked_package = Package::from_bytes(&unpacked(source)?)?;
        for sheet in 0..package.sheets().len() {
            for table in 0..package.sheets()[sheet].table_count() {
                let (body_rows, body_columns) = body_lengths(&package, sheet, table)?;
                for op in ops(body_rows, body_columns) {
                    let packed_result = leaf(&package, sheet, table, op);
                    let unpacked_result = leaf(&unpacked_package, sheet, table, op);
                    let key = match (&packed_result, &unpacked_result) {
                        (Ok(packed_bytes), Ok(unpacked_bytes)) => {
                            let left = objects(&unpacked(packed_bytes)?)?;
                            let right = objects(unpacked_bytes)?;
                            if left != right
                                || formula_texts(packed_bytes)? != formula_texts(unpacked_bytes)?
                            {
                                mismatches.push(format!(
                                    "{name} table {table} {op:?}: packed result differs"
                                ));
                            }
                            "packed: both ok".to_owned()
                        },
                        (Err(left), Err(right)) if left == right => {
                            format!("packed: both refused, {left:?}")
                        },
                        _ => {
                            mismatches.push(format!(
                                "{name} table {table} {op:?}: packed {packed_result:?} unpacked {:?}",
                                unpacked_result.as_ref().map(|_| ())
                            ));
                            "packed: different".to_owned()
                        },
                    };
                    *summary.entry(key).or_default() += 1;
                }
            }
        }
    }
    for (key, count) in &summary {
        eprintln!("{count:5} {key}");
    }
    for defect in &host_defects {
        eprintln!("host defect: {defect}");
    }
    for refusal in &refusals {
        eprintln!("refused: {refusal}");
    }
    for mismatch in &mismatches {
        eprintln!("{mismatch}");
    }
    assert!(mismatches.is_empty(), "{} mismatches", mismatches.len());
    let count = |key: &str| summary.get(key).copied().unwrap_or(0);
    assert!(count("formulas: both ok, equal") >= 40, "{summary:?}");
    assert!(count("plain: both ok, equal") >= 300, "{summary:?}");
    assert!(count("packed: both ok") >= 40, "{summary:?}");
    assert!(
        count("plain: host refuses incoming references, leaf matches the oracle") >= 8,
        "{summary:?}"
    );
    // Record 0329: removals the host refuses for a deleted reference write
    // `#REF!`, and every removal the host publishes the leaf publishes too.
    assert!(
        count(
            "formulas: host refuses deleted references, leaf writes #REF! and matches the oracle"
        ) >= 8,
        "{summary:?}"
    );
    assert!(
        !summary.keys().any(|key| key.contains("host ok, leaf")),
        "{summary:?}"
    );
    Ok(())
}
