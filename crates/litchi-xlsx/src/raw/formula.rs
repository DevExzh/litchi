//! Shared-formula expansion without exposing storage IDs in the facade.

use std::borrow::Cow;

use litchi_sheet::{COLUMNS, ROWS};

use crate::error::{Result, invalid};

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct Range {
    pub(super) first_row: u32,
    pub(super) first_column: u32,
    pub(super) last_row: u32,
    pub(super) last_column: u32,
}

impl Range {
    pub(super) fn parse(value: &str) -> Result<Self> {
        let (first, last) = value.split_once(':').unwrap_or((value, value));
        if last.contains(':') {
            return Err(invalid(format!("invalid shared formula range '{value}'")));
        }
        let (first_row, first_column) = position(first)
            .ok_or_else(|| invalid(format!("invalid shared formula range '{value}'")))?;
        let (last_row, last_column) = position(last)
            .ok_or_else(|| invalid(format!("invalid shared formula range '{value}'")))?;
        if first_row > last_row || first_column > last_column {
            return Err(invalid(format!("reversed shared formula range '{value}'")));
        }
        Ok(Self {
            first_row,
            first_column,
            last_row,
            last_column,
        })
    }

    pub(super) fn contains(self, row: u32, column: u32) -> bool {
        (self.first_row..=self.last_row).contains(&row)
            && (self.first_column..=self.last_column).contains(&column)
    }
}

fn position(value: &str) -> Option<(u32, u32)> {
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
    let row = value[row_start..].parse::<u32>().ok()?;
    (row != 0 && row <= ROWS && column <= COLUMNS).then_some((row, column))
}

fn decode_column(value: &str) -> Option<u32> {
    let mut column = 0u32;
    for byte in value.bytes() {
        let upper = byte.to_ascii_uppercase();
        if !upper.is_ascii_uppercase() {
            return None;
        }
        column = column
            .checked_mul(26)?
            .checked_add(u32::from(upper - b'A' + 1))?;
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

/// One simultaneous local-sheet rename. External workbook indexes are never
/// matched against this mapping.
#[derive(Debug, Clone, Copy)]
pub(super) struct Rename<'a> {
    pub(super) before: &'a str,
    pub(super) after: &'a str,
}

/// Result of scanning one formula-like value. `text` is allocated only when
/// at least one byte must change; `matched` also reports case-preserving no-ops.
#[derive(Debug)]
pub(super) struct RenameResult {
    pub(super) text: Option<String>,
    pub(super) matched: bool,
}

pub(super) fn translate(
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
            render_reference(parsed.reference, row_delta, column_delta, &mut output);
            index = parsed.end;
            continue;
        }
        if bytes[index] == b'[' {
            let end = bracket_end(bytes, index);
            output.push_str(&formula[index..end]);
            index = end;
            continue;
        }
        let Some(character) = formula[index..].chars().next() else {
            break;
        };
        output.push(character);
        index += character.len_utf8();
    }
    output
}

/// Rewrite local single-sheet and 3-D prefixes in one pass.
///
/// String constants, structured-reference brackets, and nonzero/external
/// workbook prefixes remain exact. All mappings observe the source formula,
/// so swaps such as `One <-> Two` do not cascade.
pub(super) fn rename_sheets(formula: &str, renames: &[Rename<'_>]) -> RenameResult {
    if formula.is_empty() || renames.is_empty() {
        return RenameResult {
            text: None,
            matched: false,
        };
    }
    let bytes = formula.as_bytes();
    let mut output = None::<String>;
    let mut copied = 0usize;
    let mut index = 0usize;
    let mut matched = false;

    while index < bytes.len() {
        if bytes[index] == b'"' {
            index = quoted_string_end(bytes, index);
            continue;
        }

        let candidate = if bytes[index] == b'\'' {
            quoted_sheet_prefix_end(bytes, index).map(|end| (end, true))
        } else if is_prefix_start(bytes, index) {
            unquoted_sheet_prefix_end(bytes, index).map(|end| (end, false))
        } else {
            None
        };
        if let Some((end, quoted)) = candidate {
            let prefix_end = end - 1;
            let raw = &formula[index..prefix_end];
            if let Some(rewritten) = rewrite_sheet_prefix(raw, quoted, renames) {
                matched = true;
                if rewritten != raw {
                    let target = output.get_or_insert_with(|| {
                        let mut value = String::with_capacity(formula.len());
                        value.push_str(&formula[..index]);
                        value
                    });
                    if copied != 0 {
                        target.push_str(&formula[copied..index]);
                    }
                    target.push_str(&rewritten);
                    target.push('!');
                    copied = end;
                }
            }
            index = end;
            continue;
        }

        let Some(character) = formula[index..].chars().next() else {
            break;
        };
        index += character.len_utf8();
    }

    if let Some(output) = output.as_mut() {
        output.push_str(&formula[copied..]);
    }
    RenameResult {
        text: output,
        matched,
    }
}

/// Whether a formula-like value has a local sheet reference whose meaning
/// depends on one checked sheet position.
///
/// Direct prefixes and local 3-D ranges are recognized. External workbook
/// indexes and inert string or structured-reference text are excluded.
pub(super) fn depends_on_sheet(
    formula: &str,
    target: &str,
    target_position: usize,
    sheets: &[&str],
) -> bool {
    if formula.is_empty() || target_position >= sheets.len() {
        return false;
    }
    let bytes = formula.as_bytes();
    let mut index = 0usize;
    while index < bytes.len() {
        if bytes[index] == b'"' {
            index = quoted_string_end(bytes, index);
            continue;
        }
        let candidate = if bytes[index] == b'\'' {
            quoted_sheet_prefix_end(bytes, index).map(|end| (end, true))
        } else if is_prefix_start(bytes, index) {
            unquoted_sheet_prefix_end(bytes, index).map(|end| (end, false))
        } else {
            None
        };
        if let Some((end, quoted)) = candidate {
            if prefix_depends_on_sheet(
                &formula[index..end - 1],
                quoted,
                target,
                target_position,
                sheets,
            ) {
                return true;
            }
            index = end;
            continue;
        }
        if bytes[index] == b'[' {
            index = bracket_end(bytes, index);
            continue;
        }
        let Some(character) = formula[index..].chars().next() else {
            break;
        };
        index += character.len_utf8();
    }
    false
}

/// Whether formula evaluation can construct a reference from runtime text.
///
/// Static dependency analysis cannot prove which sheet these functions will
/// address, so destructive callers must treat them as an unmodeled reference.
pub(super) fn has_dynamic_reference(formula: &str) -> bool {
    let bytes = formula.as_bytes();
    let mut index = 0usize;
    while index < bytes.len() {
        if bytes[index] == b'"' {
            index = quoted_string_end(bytes, index);
            continue;
        }
        if bytes[index] == b'[' {
            index = bracket_end(bytes, index);
            continue;
        }
        if bytes[index].is_ascii_alphabetic() || matches!(bytes[index], b'_' | b'.') {
            let start = index;
            index += 1;
            while bytes
                .get(index)
                .is_some_and(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'_' | b'.'))
            {
                index += 1;
            }
            let function = formula[start..index].rsplit('.').next().unwrap_or_default();
            let mut next = index;
            while bytes.get(next).is_some_and(u8::is_ascii_whitespace) {
                next += 1;
            }
            if bytes.get(next) == Some(&b'(')
                && matches_ignore_ascii_case(function, &["INDIRECT", "EVALUATE"])
            {
                return true;
            }
            continue;
        }
        let Some(character) = formula[index..].chars().next() else {
            break;
        };
        index += character.len_utf8();
    }
    false
}

fn matches_ignore_ascii_case(value: &str, expected: &[&str]) -> bool {
    expected
        .iter()
        .any(|candidate| value.eq_ignore_ascii_case(candidate))
}

fn prefix_depends_on_sheet(
    raw: &str,
    quoted: bool,
    target: &str,
    target_position: usize,
    sheets: &[&str],
) -> bool {
    let decoded;
    let prefix = if quoted {
        let Some(inner) = raw
            .strip_prefix('\'')
            .and_then(|value| value.strip_suffix('\''))
        else {
            return false;
        };
        decoded = if inner.contains("''") {
            Cow::Owned(inner.replace("''", "'"))
        } else {
            Cow::Borrowed(inner)
        };
        decoded.as_ref()
    } else {
        raw
    };
    let Some((_, sheet_range)) = local_workbook_prefix(prefix) else {
        return false;
    };
    let (first, last) = sheet_range
        .split_once(':')
        .map_or((sheet_range, None), |(first, last)| (first, Some(last)));
    if first.is_empty()
        || last.is_some_and(str::is_empty)
        || last.is_some_and(|last| last.contains(':'))
    {
        return false;
    }
    if crate::sheet::equivalent(first, target)
        || last.is_some_and(|last| crate::sheet::equivalent(last, target))
    {
        return true;
    }
    let Some(last) = last else {
        return false;
    };
    let first_position = sheets
        .iter()
        .position(|candidate| crate::sheet::equivalent(candidate, first));
    let last_position = sheets
        .iter()
        .position(|candidate| crate::sheet::equivalent(candidate, last));
    let (Some(first_position), Some(last_position)) = (first_position, last_position) else {
        return false;
    };
    (first_position.min(last_position)..=first_position.max(last_position))
        .contains(&target_position)
}

fn quoted_sheet_prefix_end(bytes: &[u8], start: usize) -> Option<usize> {
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
    None
}

fn unquoted_sheet_prefix_end(bytes: &[u8], start: usize) -> Option<usize> {
    let mut index = start;
    while bytes.get(index).is_some_and(|byte| is_prefix_byte(*byte)) {
        index += 1;
    }
    (index != start && bytes.get(index) == Some(&b'!')).then_some(index + 1)
}

fn is_prefix_start(bytes: &[u8], start: usize) -> bool {
    is_prefix_byte(bytes[start]) && (start == 0 || !is_prefix_byte(bytes[start - 1]))
}

fn rewrite_sheet_prefix(raw: &str, quoted: bool, renames: &[Rename<'_>]) -> Option<String> {
    let decoded;
    let prefix = if quoted {
        let inner = raw.strip_prefix('\'')?.strip_suffix('\'')?;
        decoded = if inner.contains("''") {
            Cow::Owned(inner.replace("''", "'"))
        } else {
            Cow::Borrowed(inner)
        };
        decoded.as_ref()
    } else {
        raw
    };
    let (book, sheets) = local_workbook_prefix(prefix)?;
    let (first, last) = sheets
        .split_once(':')
        .map_or((sheets, None), |(first, last)| (first, Some(last)));
    if first.is_empty()
        || last.is_some_and(str::is_empty)
        || last.is_some_and(|last| last.contains(':'))
    {
        return None;
    }

    let mut matched = false;
    let first = mapped_name(first, renames, &mut matched);
    let last = last.map(|last| mapped_name(last, renames, &mut matched));
    if !matched {
        return None;
    }

    let quote = quoted
        || !formula_name_is_unquoted(first)
        || last.is_some_and(|last| !formula_name_is_unquoted(last));
    let mut output = String::with_capacity(raw.len());
    if quote {
        output.push('\'');
    }
    output.push_str(book);
    push_formula_name(&mut output, first, quote);
    if let Some(last) = last {
        output.push(':');
        push_formula_name(&mut output, last, quote);
    }
    if quote {
        output.push('\'');
    }
    Some(output)
}

fn local_workbook_prefix(prefix: &str) -> Option<(&str, &str)> {
    if let Some(rest) = prefix.strip_prefix('[') {
        let close = rest.find(']')?;
        let workbook = &rest[..close];
        if workbook.parse::<u32>().ok()? != 0 {
            return None;
        }
        let book_end = close + 2;
        return Some((&prefix[..book_end], &prefix[book_end..]));
    }
    // Local sheet names cannot contain these characters. Their presence here
    // denotes an external path/book prefix rather than a local sheet.
    if prefix.contains(['[', ']', '\\', '/']) {
        return None;
    }
    Some(("", prefix))
}

fn mapped_name<'a>(name: &'a str, renames: &[Rename<'a>], matched: &mut bool) -> &'a str {
    for rename in renames {
        if crate::sheet::equivalent(name, rename.before) {
            *matched = true;
            return rename.after;
        }
    }
    name
}

fn formula_name_is_unquoted(name: &str) -> bool {
    !name.is_empty()
        && name
            .chars()
            .all(|character| character.is_alphanumeric() || matches!(character, '_' | '.'))
}

fn push_formula_name(output: &mut String, name: &str, quoted: bool) {
    if quoted {
        for character in name.chars() {
            output.push(character);
            if character == '\'' {
                output.push('\'');
            }
        }
    } else {
        output.push_str(name);
    }
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
                depth = depth.saturating_sub(1);
                if depth == 0 {
                    return index + 1;
                }
            },
            _ => {},
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
    if bytes
        .get(end)
        .is_some_and(|byte| is_identifier_byte(*byte) || *byte == b'(')
    {
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
            _ => {},
        }
        index += 1;
    }
    None
}

fn is_prefix_byte(byte: u8) -> bool {
    !byte.is_ascii()
        || byte.is_ascii_alphanumeric()
        || matches!(
            byte,
            b'_' | b'.' | b':' | b'\\' | b'/' | b'[' | b']' | b'-' | b'$'
        )
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
    (value <= COLUMNS).then_some((Axis { value, absolute }, index))
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
    (value != 0 && value <= ROWS).then_some((Axis { value, absolute }, index))
}

fn shifted(axis: Axis, delta: i64, maximum: u32) -> Option<u32> {
    if axis.absolute {
        return Some(axis.value);
    }
    let value = i64::from(axis.value).checked_add(delta)?;
    u32::try_from(value)
        .ok()
        .filter(|value| (1..=maximum).contains(value))
}

fn render_reference(reference: Reference, row_delta: i64, column_delta: i64, output: &mut String) {
    match reference {
        Reference::Cell(cell) => render_cell(cell, row_delta, column_delta, output),
        Reference::Area(first, last) => {
            render_cell(first, row_delta, column_delta, output);
            output.push(':');
            render_cell(last, row_delta, column_delta, output);
        },
        Reference::Columns(first, last) => {
            render_column(first, column_delta, output);
            output.push(':');
            render_column(last, column_delta, output);
        },
        Reference::Rows(first, last) => {
            render_row(first, row_delta, output);
            output.push(':');
            render_row(last, row_delta, output);
        },
    }
}

fn render_cell(cell: CellReference, row_delta: i64, column_delta: i64, output: &mut String) {
    let Some(column) = shifted(cell.column, column_delta, COLUMNS) else {
        output.push_str("#REF!");
        return;
    };
    let Some(row) = shifted(cell.row, row_delta, ROWS) else {
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
    let Some(value) = shifted(column, delta, COLUMNS) else {
        output.push_str("#REF!");
        return;
    };
    if column.absolute {
        output.push('$');
    }
    encode_column(value, output);
}

fn render_row(row: Axis, delta: i64, output: &mut String) {
    let Some(value) = shifted(row, delta, ROWS) else {
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

    #[test]
    fn translates_reference_forms_without_touching_strings_or_structured_refs() {
        let input = r#"SUM(A1,$A1,A$1,$A$1,A1:B2,A:C,1:3,'My Sheet'!A1,Sheet1:Sheet3!A1,[1]Sheet1!A1,INDIRECT("A1"),LOG10(2),Table1[A1])"#;
        assert_eq!(
            translate(input, 1, 1, 3, 2),
            r#"SUM(B3,$A3,B$1,$A$1,B3:C4,B:D,3:5,'My Sheet'!B3,Sheet1:Sheet3!B3,[1]Sheet1!B3,INDIRECT("A1"),LOG10(2),Table1[A1])"#
        );
    }

    #[test]
    fn checked_reference_shift_produces_ref_error() {
        assert_eq!(translate("A1:$A$1", 2, 2, 1, 1), "#REF!:$A$1");
        assert_eq!(translate("XFD1048576", 1, 1, 2, 2), "#REF!");
    }

    #[test]
    fn renames_local_prefixes_simultaneously_and_preserves_external_references() {
        let renames = [
            Rename {
                before: "One",
                after: "Two",
            },
            Rename {
                before: "Two",
                after: "One",
            },
        ];
        let result = rename_sheets(
            r#"One!A1+Two!LocalName+'One:Two'!$A$1+[0]One!A1+[1]One!A1+'[Book.xlsx]One'!A1+INDIRECT("One!A1")"#,
            &renames,
        );
        assert!(result.matched);
        assert_eq!(
            result.text.as_deref(),
            Some(
                r#"Two!A1+One!LocalName+'Two:One'!$A$1+[0]Two!A1+[1]One!A1+'[Book.xlsx]One'!A1+INDIRECT("One!A1")"#
            )
        );
    }

    #[test]
    fn renames_quoted_apostrophes_and_unicode_caseless_names() {
        let renames = [Rename {
            before: "Straße",
            after: "O'Brien 2026",
        }];
        let result = rename_sheets("STRASSE!A1+'Straße'!Named", &renames);
        assert_eq!(
            result.text.as_deref(),
            Some("'O''Brien 2026'!A1+'O''Brien 2026'!Named")
        );
    }

    #[test]
    fn detects_direct_and_implicit_three_d_sheet_dependencies() {
        let sheets = ["One", "Middle", "Three"];
        assert!(depends_on_sheet("Middle!A1", "Middle", 1, &sheets));
        assert!(depends_on_sheet("One:Three!A1", "Middle", 1, &sheets));
        assert!(depends_on_sheet(
            "'[0]One:Three'!Named",
            "Middle",
            1,
            &sheets
        ));
        assert!(depends_on_sheet("[0]Middle!A1", "Middle", 1, &sheets));
        assert!(!depends_on_sheet("One!A1", "Middle", 1, &sheets));
        assert!(!depends_on_sheet("[1]Middle!A1", "Middle", 1, &sheets));
        assert!(!depends_on_sheet(
            r#"INDIRECT("Middle!A1")+Table1[[Middle!A1]]"#,
            "Middle",
            1,
            &sheets
        ));
        assert!(has_dynamic_reference(r#"INDIRECT("Middle!A1")"#));
        assert!(has_dynamic_reference(r#"_xlfn.EVALUATE ("Middle!A1")"#));
        assert!(!has_dynamic_reference(
            r#""INDIRECT(Middle!A1)"+Table1[INDIRECT()]"#
        ));
    }

    #[test]
    fn rename_scan_does_not_allocate_without_a_local_match() {
        let result = rename_sheets(
            r#"[2]One!A1+Table1[One]+"One!A1""#,
            &[Rename {
                before: "One",
                after: "Renamed",
            }],
        );
        assert!(!result.matched);
        assert!(result.text.is_none());
    }
}
