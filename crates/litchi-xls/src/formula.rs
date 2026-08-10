//! BIFF8 formula token rendering.
//!
//! Formula expressions are stored as reverse-Polish `Ptg` token streams.  This
//! renderer resolves workbook-dependent names and 3-D references through a
//! bounded, inert context. Callers retain the original bytes when rendering
//! returns `None`, so malformed or unsupported expressions remain lossless.

use super::external_link::{Links, NameBody};
use std::fmt::{self, Write as _};

const MAX_FORMULA_TOKEN_BYTES: usize = 1024 * 1024;
const MAX_FORMULA_STACK_ENTRIES: usize = 16 * 1024;
const MAX_RENDERED_FORMULA_BYTES: usize = 8 * 1024 * 1024;

/// Fallible, bounded text assembly for inert formula rendering.
struct BoundedText {
    value: String,
}

impl BoundedText {
    fn new() -> Self {
        Self {
            value: String::new(),
        }
    }

    fn push_str(&mut self, value: &str) -> Result<(), ()> {
        self.write_str(value).map_err(|_error| ())
    }

    fn finish(self) -> String {
        self.value
    }
}

impl fmt::Write for BoundedText {
    fn write_str(&mut self, value: &str) -> fmt::Result {
        let length = self
            .value
            .len()
            .checked_add(value.len())
            .ok_or(fmt::Error)?;
        if length > MAX_RENDERED_FORMULA_BYTES {
            return Err(fmt::Error);
        }
        self.value
            .try_reserve(value.len())
            .map_err(|_error| fmt::Error)?;
        self.value.push_str(value);
        Ok(())
    }
}

fn render_text(arguments: fmt::Arguments<'_>) -> Result<String, ()> {
    let mut text = BoundedText::new();
    fmt::write(&mut text, arguments).map_err(|_error| ())?;
    Ok(text.finish())
}

/// Workbook-global context needed to resolve BIFF `ixti` sheet references.
#[derive(Debug, Default)]
pub(crate) struct FormulaContext {
    sup_books: Vec<SupBookKind>,
    extern_sheets: Vec<ExternSheetRef>,
    sheet_names: Vec<String>,
    defined_names: Vec<Option<FormulaDefinedName>>,
    external_names: Vec<Vec<String>>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct FormulaDefinedName {
    name: String,
    sheet_index: Option<usize>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
enum SupBookKind {
    Internal,
    External(ExternalSupBook),
    Other,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct ExternalSupBook {
    workbook: String,
    sheet_names: Vec<String>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct ExternSheetRef {
    sup_book: u16,
    first_sheet: i16,
    last_sheet: i16,
}

impl FormulaContext {
    pub(crate) fn add_sup_book(&mut self, data: &[u8]) {
        let kind = if data.len() == 4 && data[2..4] == [0x01, 0x04] {
            SupBookKind::Internal
        } else if let Some(external) = parse_external_sup_book(data) {
            SupBookKind::External(external)
        } else {
            SupBookKind::Other
        };
        self.sup_books.push(kind);
    }

    pub(crate) fn add_extern_sheet(&mut self, data: &[u8]) -> Result<(), &'static str> {
        let count_bytes = data.get(..2).ok_or("EXTERNSHEET is missing cXTI")?;
        let count = usize::from(u16::from_le_bytes([count_bytes[0], count_bytes[1]]));
        let expected = count
            .checked_mul(6)
            .and_then(|size| size.checked_add(2))
            .ok_or("EXTERNSHEET count overflows")?;
        if data.len() != expected {
            return Err("EXTERNSHEET length does not match cXTI");
        }

        for entry in data[2..].chunks_exact(6) {
            self.extern_sheets.push(ExternSheetRef {
                sup_book: u16::from_le_bytes([entry[0], entry[1]]),
                first_sheet: i16::from_le_bytes([entry[2], entry[3]]),
                last_sheet: i16::from_le_bytes([entry[4], entry[5]]),
            });
        }
        Ok(())
    }

    pub(crate) fn set_sheet_names(&mut self, sheet_names: Vec<String>) {
        self.sheet_names = sheet_names;
    }

    pub(crate) fn set_scoped_defined_names(
        &mut self,
        defined_names: Vec<Option<(String, Option<usize>)>>,
    ) {
        self.defined_names = defined_names
            .into_iter()
            .map(|name| name.map(|(name, sheet_index)| FormulaDefinedName { name, sheet_index }))
            .collect();
    }

    pub(crate) fn set_external_links(&mut self, links: &Links) {
        self.external_names = vec![Vec::new(); links.supporting_books().len()];
        for external_name in links.external_names() {
            let name = match external_name.body() {
                NameBody::ExternalDefinedName { name, .. }
                | NameBody::AddInFunction { name, .. }
                | NameBody::DdeOrOle { name, .. }
                | NameBody::DdeStandardDocumentName { name } => name,
            };
            let Some(book_names) = self
                .external_names
                .get_mut(usize::from(external_name.supporting_book_index()))
            else {
                continue;
            };
            book_names.push(name.clone());
        }
    }

    fn defined_name(&self, one_based_index: u32) -> Option<&str> {
        let index = usize::try_from(one_based_index.checked_sub(1)?).ok()?;
        Some(self.defined_names.get(index)?.as_ref()?.name.as_str())
    }

    fn name_x(&self, extern_sheet: u16, one_based_index: u16) -> Option<String> {
        let reference = self.extern_sheets.get(usize::from(extern_sheet))?;
        let name_index = usize::from(one_based_index.checked_sub(1)?);
        if let Some(name) = self
            .external_names
            .get(usize::from(reference.sup_book))
            .and_then(|names| names.get(name_index))
        {
            return render_text(format_args!("{name}")).ok();
        }
        if reference.first_sheet != -2 || reference.last_sheet != -2 {
            return None;
        }
        let name = self.defined_names.get(name_index)?.as_ref()?;
        match name.sheet_index {
            Some(sheet_index) => {
                let sheet_name = escape_formula_name(self.sheet_names.get(sheet_index)?)?;
                render_text(format_args!("'{sheet_name}'!{}", name.name)).ok()
            },
            None => render_text(format_args!("{}", name.name)).ok(),
        }
    }

    fn sheet_prefix(&self, extern_sheet: u16) -> Option<String> {
        let reference = self.extern_sheets.get(usize::from(extern_sheet))?;
        let first = usize::try_from(reference.first_sheet).ok()?;
        let last = usize::try_from(reference.last_sheet).ok()?;
        match self.sup_books.get(usize::from(reference.sup_book))? {
            SupBookKind::Internal => {
                let first_name = self.sheet_names.get(first)?;
                let last_name = self.sheet_names.get(last)?;
                let first_name = escape_formula_name(first_name)?;
                let last_name = escape_formula_name(last_name)?;
                if first == last {
                    render_text(format_args!("'{first_name}'!")).ok()
                } else {
                    render_text(format_args!("'{first_name}:{last_name}'!")).ok()
                }
            },
            SupBookKind::External(external) => {
                let first_name = external.sheet_names.get(first)?;
                let last_name = external.sheet_names.get(last)?;
                let workbook = normalize_external_workbook(&external.workbook)?;
                let workbook = escape_formula_name(&workbook)?;
                let first_name = escape_formula_name(first_name)?;
                if first == last {
                    render_text(format_args!("'[{workbook}]{first_name}'!")).ok()
                } else {
                    let last_name = escape_formula_name(last_name)?;
                    render_text(format_args!("'[{workbook}]{first_name}:{last_name}'!")).ok()
                }
            },
            SupBookKind::Other => None,
        }
    }
}

fn escape_formula_name(name: &str) -> Option<String> {
    let mut value = BoundedText::new();
    for character in name.chars() {
        if character == '\'' {
            value.push_str("''").ok()?;
        } else {
            fmt::Write::write_char(&mut value, character).ok()?;
        }
    }
    Some(value.finish())
}

fn normalize_external_workbook(value: &str) -> Option<String> {
    let mut output = BoundedText::new();
    for character in value.chars() {
        match character {
            '[' => output.push_str("(").ok()?,
            ']' => output.push_str(")").ok()?,
            character => fmt::Write::write_char(&mut output, character).ok()?,
        }
    }
    Some(output.finish())
}

fn parse_external_sup_book(data: &[u8]) -> Option<ExternalSupBook> {
    if data.len() <= 4 {
        return None;
    }
    let sheet_count = usize::from(u16::from_le_bytes([*data.first()?, *data.get(1)?]));
    let (encoded_workbook, mut offset) = parse_biff_unicode_string(data, 2)?;
    let workbook = decode_sup_book_url(&encoded_workbook)?;
    let mut sheet_names = Vec::with_capacity(sheet_count);
    for _ in 0..sheet_count {
        let (sheet_name, next) = parse_biff_unicode_string(data, offset)?;
        if sheet_name.is_empty()
            || sheet_name
                .chars()
                .any(|character| matches!(character, '\0' | '\r' | '\n'))
        {
            return None;
        }
        sheet_names.push(sheet_name);
        offset = next;
    }
    if offset != data.len() {
        return None;
    }
    Some(ExternalSupBook {
        workbook,
        sheet_names,
    })
}

fn parse_biff_unicode_string(data: &[u8], offset: usize) -> Option<(String, usize)> {
    let header = data.get(offset..offset.checked_add(3)?)?;
    let count = usize::from(u16::from_le_bytes([header[0], header[1]]));
    let flags = header[2];
    if flags & !1 != 0 {
        return None;
    }
    let width = if flags == 0 { 1usize } else { 2 };
    let start = offset.checked_add(3)?;
    let end = start.checked_add(count.checked_mul(width)?)?;
    let bytes = data.get(start..end)?;
    let value = if width == 1 {
        bytes.iter().map(|byte| char::from(*byte)).collect()
    } else {
        let units = bytes
            .chunks_exact(2)
            .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
            .collect::<Vec<_>>();
        String::from_utf16(&units).ok()?
    };
    Some((value, end))
}

fn decode_sup_book_url(encoded: &str) -> Option<String> {
    const ENCODED_FILE: char = '\u{1}';
    const SELF_REFERENCE: char = '\u{2}';
    const EMPTY_WORKBOOK: char = '\0';
    const SAME_VOLUME: char = '\u{2}';
    const DOWN_DIRECTORY: char = '\u{3}';
    const UP_DIRECTORY: char = '\u{4}';
    const LONG_VOLUME: char = '\u{5}';
    const STARTUP_DIRECTORY: char = '\u{6}';
    const ALTERNATE_STARTUP_DIRECTORY: char = '\u{7}';
    const LIBRARY_DIRECTORY: char = '\u{8}';

    let mut characters = encoded.chars();
    match characters.next()? {
        EMPTY_WORKBOOK | SELF_REFERENCE => Some(characters.collect()),
        ENCODED_FILE => {
            let mut output = String::with_capacity(encoded.len());
            while let Some(character) = characters.next() {
                match character {
                    ENCODED_FILE => {
                        let volume = characters.next()?;
                        if volume == '@' {
                            output.push_str("\\\\");
                        } else {
                            output.push(volume);
                            output.push(':');
                        }
                    },
                    SAME_VOLUME | DOWN_DIRECTORY => output.push('\\'),
                    UP_DIRECTORY => output.push_str("..\\"),
                    LONG_VOLUME => {},
                    STARTUP_DIRECTORY | ALTERNATE_STARTUP_DIRECTORY | LIBRARY_DIRECTORY => {
                        output.push_str(".\\");
                    },
                    '\0' | '\r' | '\n' => return None,
                    other => output.push(other),
                }
            }
            Some(output)
        },
        _ => None,
    }
}

/// Render a BIFF8 formula token stream in A1 notation.
pub(crate) fn render_formula(tokens: &[u8], context: Option<&FormulaContext>) -> Option<String> {
    if tokens.len() > MAX_FORMULA_TOKEN_BYTES {
        return None;
    }
    let mut decoder = FormulaDecoder::new(tokens, context, None);
    decoder
        .decode()
        .ok()
        .and_then(|formula| render_text(format_args!("={formula}")).ok())
}

/// Certifies a bounded Formula token stream as independent of workbook or
/// worksheet coordinates. Unknown tokens, names, and every 2-D/3-D reference
/// form fail closed.
pub(crate) fn formula_is_reference_free(tokens: &[u8]) -> bool {
    if tokens.len() > MAX_FORMULA_TOKEN_BYTES {
        return false;
    }
    let mut decoder = FormulaDecoder::new(tokens, None, None);
    decoder.decode().is_ok() && !decoder.coordinate_dependency
}

/// Render a shared formula template at a particular formula-cell origin.
pub(crate) fn render_shared_formula(
    tokens: &[u8],
    context: Option<&FormulaContext>,
    row: u16,
    column: u16,
) -> Option<String> {
    if tokens.len() > MAX_FORMULA_TOKEN_BYTES {
        return None;
    }
    let mut decoder = FormulaDecoder::new(tokens, context, Some((row, column)));
    decoder
        .decode()
        .ok()
        .and_then(|formula| render_text(format_args!("={formula}")).ok())
}

/// Return the shared/array formula anchor encoded by a standalone `PtgExp`.
pub(crate) fn ptg_exp_anchor(tokens: &[u8]) -> Option<(u16, u16)> {
    if tokens.len() != 5 || tokens[0] != 0x01 || tokens[4] != 0 {
        return None;
    }
    Some((
        u16::from_le_bytes([tokens[1], tokens[2]]),
        u16::from_le_bytes([tokens[3], tokens[4]]),
    ))
}

struct FormulaDecoder<'a> {
    data: &'a [u8],
    pos: usize,
    stack: Vec<String>,
    stack_bytes: usize,
    name_x_operands: Vec<usize>,
    context: Option<&'a FormulaContext>,
    shared_origin: Option<(u16, u16)>,
    coordinate_dependency: bool,
}

impl<'a> FormulaDecoder<'a> {
    fn new(
        data: &'a [u8],
        context: Option<&'a FormulaContext>,
        shared_origin: Option<(u16, u16)>,
    ) -> Self {
        Self {
            data,
            pos: 0,
            stack: Vec::new(),
            stack_bytes: 0,
            name_x_operands: Vec::new(),
            context,
            shared_origin,
            coordinate_dependency: false,
        }
    }

    fn decode(&mut self) -> Result<String, ()> {
        if self.data.is_empty() {
            return Err(());
        }

        while self.pos < self.data.len() {
            let opcode = self.byte()?;
            if opcode < 0x20 {
                self.decode_base(opcode)?;
            } else {
                self.decode_classified(opcode)?;
            }
        }

        if self.stack.len() != 1 {
            return Err(());
        }
        self.pop()
    }

    fn push_text(&mut self, value: String) -> Result<(), ()> {
        if self.stack.len() >= MAX_FORMULA_STACK_ENTRIES {
            return Err(());
        }
        let stack_bytes = self.stack_bytes.checked_add(value.len()).ok_or(())?;
        if stack_bytes > MAX_RENDERED_FORMULA_BYTES {
            return Err(());
        }
        self.stack.try_reserve(1).map_err(|_error| ())?;
        self.stack.push(value);
        self.stack_bytes = stack_bytes;
        Ok(())
    }

    fn push_rendered(&mut self, arguments: fmt::Arguments<'_>) -> Result<(), ()> {
        self.push_text(render_text(arguments)?)
    }

    fn push_quoted(&mut self, value: &str) -> Result<(), ()> {
        let mut text = BoundedText::new();
        fmt::Write::write_char(&mut text, '"').map_err(|_error| ())?;
        for character in value.chars() {
            if character == '"' {
                text.push_str("\"\"")?;
            } else {
                fmt::Write::write_char(&mut text, character).map_err(|_error| ())?;
            }
        }
        fmt::Write::write_char(&mut text, '"').map_err(|_error| ())?;
        self.push_text(text.finish())
    }

    fn truncate_stack(&mut self, start: usize) -> Result<(), ()> {
        if start > self.stack.len() {
            return Err(());
        }
        let removed = self.stack[start..]
            .iter()
            .try_fold(0usize, |total, value| {
                total.checked_add(value.len()).ok_or(())
            })?;
        self.stack.truncate(start);
        self.stack_bytes = self.stack_bytes.checked_sub(removed).ok_or(())?;
        Ok(())
    }

    fn decode_base(&mut self, opcode: u8) -> Result<(), ()> {
        match opcode {
            0x03 => self.binary("+"),
            0x04 => self.binary("-"),
            0x05 => self.binary("*"),
            0x06 => self.binary("/"),
            0x07 => self.binary("^"),
            0x08 => self.binary("&"),
            0x09 => self.binary("<"),
            0x0a => self.binary("<="),
            0x0b => self.binary("="),
            0x0c => self.binary(">="),
            0x0d => self.binary(">"),
            0x0e => self.binary("<>"),
            0x0f => self.binary(" "),
            0x10 => self.binary(","),
            0x11 => self.binary(":"),
            0x12 => self.unary_prefix("+"),
            0x13 => self.unary_prefix("-"),
            0x14 => self.unary_suffix("%"),
            0x15 => {
                let value = self.pop()?;
                self.push_rendered(format_args!("({value})"))
            },
            0x16 => self.push_text(String::new()),
            0x17 => {
                let value = self.formula_string()?;
                self.push_quoted(&value)
            },
            0x19 => self.attribute(),
            0x1c => {
                let error = match self.byte()? {
                    0x00 => "#NULL!",
                    0x07 => "#DIV/0!",
                    0x0f => "#VALUE!",
                    0x17 => "#REF!",
                    0x1d => "#NAME?",
                    0x24 => "#NUM!",
                    0x2a => "#N/A",
                    _ => return Err(()),
                };
                self.push_text(error.to_owned())
            },
            0x1d => {
                let value = match self.byte()? {
                    0 => "FALSE",
                    1 => "TRUE",
                    _ => return Err(()),
                };
                self.push_text(value.to_owned())
            },
            0x1e => {
                let value = self.u16()?;
                self.push_rendered(format_args!("{value}"))
            },
            0x1f => {
                let value = self.f64()?;
                if !value.is_finite() {
                    return Err(());
                }
                self.push_rendered(format_args!("{value}"))
            },
            _ => Err(()),
        }
    }

    fn decode_classified(&mut self, opcode: u8) -> Result<(), ()> {
        let base = (opcode & 0x1f) | 0x20;
        match base {
            0x21 => {
                let index = self.u16()?;
                let metadata = function_metadata(index).ok_or(())?;
                self.function(metadata.name, metadata.fixed_arity().ok_or(())?)
            },
            0x22 => {
                let args = self.byte()? as usize;
                let raw_index = self.u16()?;
                if raw_index == 255 {
                    return self.external_function(args);
                }
                const COMMAND_EQUIVALENT_BIT: u16 = 0x8000;
                let index = raw_index & !COMMAND_EQUIVALENT_BIT;
                if raw_index & COMMAND_EQUIVALENT_BIT != 0 {
                    let name = command_function_name(index).ok_or(())?;
                    return self.function(name, args);
                }
                let metadata = function_metadata(index).ok_or(())?;
                if !metadata.accepts_arity(args) {
                    return Err(());
                }
                self.function(metadata.name, args)
            },
            0x23 => {
                self.coordinate_dependency = true;
                let index = self.u32()?;
                let name = self
                    .context
                    .and_then(|context| context.defined_name(index))
                    .ok_or(())?;
                self.push_text(name.to_owned())
            },
            0x24 => {
                self.coordinate_dependency = true;
                let row = self.u16()?;
                let col = self.u16()?;
                let (row, col) = resolve_shared_reference(row, col, self.shared_origin)?;
                self.push_text(cell_reference(row, col)?)
            },
            0x25 => {
                self.coordinate_dependency = true;
                let first_row = self.u16()?;
                let last_row = self.u16()?;
                let first_col = self.u16()?;
                let last_col = self.u16()?;
                let (first_row, first_col) =
                    resolve_shared_reference(first_row, first_col, self.shared_origin)?;
                let (last_row, last_col) =
                    resolve_shared_reference(last_row, last_col, self.shared_origin)?;
                let first = cell_reference(first_row, first_col)?;
                let last = cell_reference(last_row, last_col)?;
                self.push_rendered(format_args!("{first}:{last}"))
            },
            // Memory tokens carry evaluator bookkeeping and do not contribute
            // an expression operand.
            0x26 | 0x27 => self.skip(6),
            0x29 => self.skip(2),
            0x2c => {
                self.coordinate_dependency = true;
                let row = self.u16()?;
                let col = self.u16()?;
                let origin = self.shared_origin.ok_or(())?;
                let (row, col) = resolve_shared_reference(row, col, Some(origin))?;
                self.push_text(cell_reference(row, col)?)
            },
            0x2d => {
                self.coordinate_dependency = true;
                let first_row = self.u16()?;
                let last_row = self.u16()?;
                let first_col = self.u16()?;
                let last_col = self.u16()?;
                let origin = self.shared_origin.ok_or(())?;
                let (first_row, first_col) =
                    resolve_shared_reference(first_row, first_col, Some(origin))?;
                let (last_row, last_col) =
                    resolve_shared_reference(last_row, last_col, Some(origin))?;
                let first = cell_reference(first_row, first_col)?;
                let last = cell_reference(last_row, last_col)?;
                self.push_rendered(format_args!("{first}:{last}"))
            },
            0x39 => {
                self.coordinate_dependency = true;
                let extern_sheet = self.u16()?;
                let name_index = self.u16()?;
                if self.u16()? != 0 {
                    return Err(());
                }
                let name = self
                    .context
                    .and_then(|context| context.name_x(extern_sheet, name_index))
                    .ok_or(())?;
                self.name_x_operands.try_reserve(1).map_err(|_error| ())?;
                self.name_x_operands.push(self.stack.len());
                self.push_text(name)
            },
            0x3a => {
                self.coordinate_dependency = true;
                let extern_sheet = self.u16()?;
                let row = self.u16()?;
                let col = self.u16()?;
                let (row, col) = resolve_shared_reference(row, col, self.shared_origin)?;
                let prefix = self
                    .context
                    .and_then(|context| context.sheet_prefix(extern_sheet))
                    .ok_or(())?;
                let reference = cell_reference(row, col)?;
                self.push_rendered(format_args!("{prefix}{reference}"))
            },
            0x3b => {
                self.coordinate_dependency = true;
                let extern_sheet = self.u16()?;
                let first_row = self.u16()?;
                let last_row = self.u16()?;
                let first_col = self.u16()?;
                let last_col = self.u16()?;
                let (first_row, first_col) =
                    resolve_shared_reference(first_row, first_col, self.shared_origin)?;
                let (last_row, last_col) =
                    resolve_shared_reference(last_row, last_col, self.shared_origin)?;
                let prefix = self
                    .context
                    .and_then(|context| context.sheet_prefix(extern_sheet))
                    .ok_or(())?;
                let first = cell_reference(first_row, first_col)?;
                let last = cell_reference(last_row, last_col)?;
                self.push_rendered(format_args!("{prefix}{first}:{last}"))
            },
            0x3c => {
                self.coordinate_dependency = true;
                let extern_sheet = self.u16()?;
                self.skip(4)?;
                let prefix = self
                    .context
                    .and_then(|context| context.sheet_prefix(extern_sheet))
                    .ok_or(())?;
                self.push_rendered(format_args!("{prefix}#REF!"))
            },
            0x3d => {
                self.coordinate_dependency = true;
                let extern_sheet = self.u16()?;
                self.skip(8)?;
                let prefix = self
                    .context
                    .and_then(|context| context.sheet_prefix(extern_sheet))
                    .ok_or(())?;
                self.push_rendered(format_args!("{prefix}#REF!"))
            },
            _ => Err(()),
        }
    }

    fn attribute(&mut self) -> Result<(), ()> {
        let options = self.byte()?;
        let data = self.u16()? as usize;
        if options & 0x04 != 0 {
            self.skip((data + 1).checked_mul(2).ok_or(())?)?;
        }
        if options & 0x10 != 0 {
            return self.function("SUM", 1);
        }
        // Optimized control-flow, volatility, and whitespace attributes are
        // rendering hints rather than formula operands.
        if options & !0x4f == 0 {
            Ok(())
        } else {
            Err(())
        }
    }

    fn binary(&mut self, operator: &str) -> Result<(), ()> {
        let right = self.pop()?;
        let left = self.pop()?;
        self.push_rendered(format_args!("({left}{operator}{right})"))
    }

    fn unary_prefix(&mut self, operator: &str) -> Result<(), ()> {
        let operand = self.pop()?;
        self.push_rendered(format_args!("{operator}{operand}"))
    }

    fn unary_suffix(&mut self, operator: &str) -> Result<(), ()> {
        let operand = self.pop()?;
        self.push_rendered(format_args!("{operand}{operator}"))
    }

    fn function(&mut self, name: &str, count: usize) -> Result<(), ()> {
        if self.stack.len() < count {
            return Err(());
        }
        let start = self.stack.len() - count;
        self.name_x_operands.retain(|operand| *operand < start);
        let mut output = BoundedText::new();
        output.push_str(name)?;
        output.push_str("(")?;
        for (index, argument) in self.stack[start..].iter().enumerate() {
            if index != 0 {
                output.push_str(",")?;
            }
            output.push_str(argument)?;
        }
        output.push_str(")")?;
        self.truncate_stack(start)?;
        self.push_text(output.finish())
    }

    fn external_function(&mut self, count: usize) -> Result<(), ()> {
        if count == 0 || self.stack.len() < count {
            return Err(());
        }
        let start = self.stack.len() - count;
        if !self.name_x_operands.contains(&start) {
            return Err(());
        }
        self.name_x_operands.retain(|operand| *operand < start);
        let operands = self.stack.get(start..).ok_or(())?;
        let name = operands.first().ok_or(())?;
        let mut output = BoundedText::new();
        output.push_str(name)?;
        output.push_str("(")?;
        for (index, argument) in operands.iter().skip(1).enumerate() {
            if index != 0 {
                output.push_str(",")?;
            }
            output.push_str(argument)?;
        }
        output.push_str(")")?;
        self.truncate_stack(start)?;
        self.push_text(output.finish())
    }

    fn formula_string(&mut self) -> Result<String, ()> {
        let count = self.byte()? as usize;
        let flags = self.byte()?;
        if flags & !0x01 != 0 {
            return Err(());
        }
        let mut output = BoundedText::new();
        if flags & 0x01 == 0 {
            let bytes = self.take(count)?;
            for &byte in bytes {
                fmt::Write::write_char(&mut output, char::from(byte)).map_err(|_error| ())?;
            }
        } else {
            let byte_count = count.checked_mul(2).ok_or(())?;
            let bytes = self.take(byte_count)?;
            for result in char::decode_utf16(
                bytes
                    .chunks_exact(2)
                    .map(|pair| u16::from_le_bytes([pair[0], pair[1]])),
            ) {
                let character = result.map_err(|_error| ())?;
                fmt::Write::write_char(&mut output, character).map_err(|_error| ())?;
            }
        }
        Ok(output.finish())
    }

    fn pop(&mut self) -> Result<String, ()> {
        let index = self.stack.len().checked_sub(1).ok_or(())?;
        self.name_x_operands.retain(|operand| *operand != index);
        let value = self.stack.pop().ok_or(())?;
        self.stack_bytes = self.stack_bytes.checked_sub(value.len()).ok_or(())?;
        Ok(value)
    }

    fn byte(&mut self) -> Result<u8, ()> {
        let byte = *self.data.get(self.pos).ok_or(())?;
        self.pos += 1;
        Ok(byte)
    }

    fn u16(&mut self) -> Result<u16, ()> {
        let bytes = self.take(2)?;
        Ok(u16::from_le_bytes([bytes[0], bytes[1]]))
    }

    fn u32(&mut self) -> Result<u32, ()> {
        let bytes = self.take(4)?;
        Ok(u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]))
    }

    fn f64(&mut self) -> Result<f64, ()> {
        let bytes = self.take(8)?;
        let mut value = [0; 8];
        value.copy_from_slice(bytes);
        Ok(f64::from_le_bytes(value))
    }

    fn skip(&mut self, count: usize) -> Result<(), ()> {
        self.take(count).map(|_| ())
    }

    fn take(&mut self, count: usize) -> Result<&'a [u8], ()> {
        let end = self.pos.checked_add(count).ok_or(())?;
        let bytes = self.data.get(self.pos..end).ok_or(())?;
        self.pos = end;
        Ok(bytes)
    }
}

fn resolve_shared_reference(
    row: u16,
    column_flags: u16,
    origin: Option<(u16, u16)>,
) -> Result<(u16, u16), ()> {
    let Some((origin_row, origin_column)) = origin else {
        return Ok((row, column_flags));
    };
    let row_relative = column_flags & 0x8000 != 0;
    let column_relative = column_flags & 0x4000 != 0;
    let resolved_row = if row_relative {
        origin_row.wrapping_add_signed(crate::utils::wrap_u16_to_i16(row))
    } else {
        row
    };
    let raw_column = column_flags & 0x00ff;
    let resolved_column = if column_relative {
        let offset = crate::utils::wrap_u8_to_i8(raw_column as u8);
        origin_column.wrapping_add_signed(i16::from(offset)) & 0x00ff
    } else {
        raw_column
    };
    Ok((resolved_row, resolved_column | (column_flags & 0xc000)))
}

fn cell_reference(row: u16, column_flags: u16) -> Result<String, ()> {
    let column = usize::from(column_flags & 0x3fff);
    if column > 255 {
        return Err(());
    }
    let row_relative = column_flags & 0x8000 != 0;
    let column_relative = column_flags & 0x4000 != 0;
    let column_name = column_name(column);
    Ok(format!(
        "{}{}{}{}",
        if column_relative { "" } else { "$" },
        column_name,
        if row_relative { "" } else { "$" },
        usize::from(row) + 1
    ))
}

fn column_name(mut column: usize) -> String {
    let mut result = String::new();
    loop {
        result.insert(
            0,
            char::from(b'A' + crate::utils::truncate_usize_to_u8(column % 26)),
        );
        if column < 26 {
            return result;
        }
        column = column / 26 - 1;
    }
}

include!("formula_function_metadata.rs");

/// Return the exact fixed arity of a built-in BIFF function, or `None` for
/// variable-arity and unknown functions.
pub(crate) fn fixed_function_arity(index: u16) -> Option<usize> {
    function_metadata(index)?.fixed_arity()
}

#[cfg(test)]
mod tests {
    use super::{
        FormulaContext, MAX_FORMULA_STACK_ENTRIES, MAX_FORMULA_TOKEN_BYTES,
        formula_is_reference_free, ptg_exp_anchor, render_formula, render_shared_formula,
    };

    #[test]
    fn reference_free_certification_rejects_coordinate_tokens() {
        assert!(formula_is_reference_free(&[0x1e, 1, 0]));
        assert!(formula_is_reference_free(&[0x1e, 1, 0, 0x1e, 2, 0, 0x03,]));
        assert!(!formula_is_reference_free(&[0x24, 0, 0, 0, 0]));
        assert!(!formula_is_reference_free(&[0x25, 0, 0, 1, 0, 0, 0, 1, 0]));
        assert!(!formula_is_reference_free(&[0xff]));
    }

    #[test]
    fn renders_constants_operators_and_references() {
        let tokens = [
            0x24, 0x00, 0x00, 0x00, 0xc0, // A1
            0x24, 0x00, 0x00, 0x01, 0xc0, // B1
            0x1e, 0x02, 0x00, // 2
            0x05, // multiply
            0x03, // add
        ];
        assert_eq!(
            render_formula(&tokens, None).as_deref(),
            Some("=(A1+(B1*2))")
        );
    }

    #[test]
    fn renders_absolute_area_and_optimized_sum() {
        let tokens = [
            0x25, 0x00, 0x00, 0x09, 0x00, 0x00, 0x00, 0x02, 0x00, // $A$1:$C$10
            0x19, 0x10, 0x00, 0x00,
        ];
        assert_eq!(
            render_formula(&tokens, None).as_deref(),
            Some("=SUM($A$1:$C$10)")
        );
    }

    #[test]
    fn renders_unicode_strings_and_variable_functions() {
        let tokens = [
            0x17, 0x02, 0x01, 0x60, 0x4f, 0x7d, 0x59, // "你好"
            0x1e, 0x02, 0x00, // 2
            0x42, 0x02, 0x50, 0x01, // CONCATENATE, two arguments
        ];
        assert_eq!(
            render_formula(&tokens, None).as_deref(),
            Some("=CONCATENATE(\"你好\",2)")
        );
    }

    #[test]
    fn rejects_workbook_dependent_and_malformed_tokens() {
        assert_eq!(render_formula(&[0x3a, 0, 0, 0, 0, 0, 0], None), None);
        assert_eq!(render_formula(&[0x1f, 0, 0], None), None);
        assert_eq!(render_formula(&[0x1e, 1, 0, 0x1e, 2, 0], None), None);
    }

    #[test]
    fn rejects_oversized_token_streams_and_operand_stacks() {
        let oversized = vec![0x16; MAX_FORMULA_TOKEN_BYTES + 1];
        assert_eq!(render_formula(&oversized, None), None);

        let too_many_operands = vec![0x16; MAX_FORMULA_STACK_ENTRIES + 1];
        assert_eq!(render_formula(&too_many_operands, None), None);
    }

    #[test]
    fn decodes_surrogate_pairs_without_allocating_a_unit_buffer() {
        let tokens = [0x17, 0x02, 0x01, 0x3d, 0xd8, 0x00, 0xde];
        assert_eq!(render_formula(&tokens, None).as_deref(), Some("=\"😀\""));
    }

    #[test]
    fn renders_internal_3d_references_from_workbook_context() {
        let mut context = FormulaContext::default();
        context.add_sup_book(&[2, 0, 0x01, 0x04]);
        context
            .add_extern_sheet(&[
                2, 0, // cXTI
                0, 0, 0, 0, 0, 0, // Sheet One
                0, 0, 0, 0, 1, 0, // Sheet One:O'Brien
            ])
            .unwrap();
        context.set_sheet_names(vec!["Sheet One".to_string(), "O'Brien".to_string()]);

        let reference = [0x5a, 0, 0, 2, 0, 1, 0xc0];
        assert_eq!(
            render_formula(&reference, Some(&context)).as_deref(),
            Some("='Sheet One'!B3")
        );

        let area = [
            0x3b, 1, 0, 0, 0, 2, 0, 0, 0, 1, 0, // $A$1:$B$3
        ];
        assert_eq!(
            render_formula(&area, Some(&context)).as_deref(),
            Some("='Sheet One:O''Brien'!$A$1:$B$3")
        );
    }

    #[test]
    fn renders_deleted_internal_3d_references() {
        let mut context = FormulaContext::default();
        context.add_sup_book(&[2, 0, 0x01, 0x04]);
        context.add_extern_sheet(&[1, 0, 0, 0, 0, 0, 0, 0]).unwrap();
        context.set_sheet_names(vec!["Sheet1".to_string()]);

        assert_eq!(
            render_formula(&[0x3c, 0, 0, 0, 0, 0, 0], Some(&context)).as_deref(),
            Some("='Sheet1'!#REF!")
        );
        assert_eq!(
            render_formula(&[0x3d, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], Some(&context)).as_deref(),
            Some("='Sheet1'!#REF!")
        );
    }

    #[test]
    fn renders_one_based_internal_name_without_expanding_it() {
        let mut context = FormulaContext::default();
        context.set_scoped_defined_names(vec![None, Some(("TaxRate".to_string(), None))]);
        assert_eq!(
            render_formula(&[0x23, 2, 0, 0, 0], Some(&context)).as_deref(),
            Some("=TaxRate")
        );
        assert_eq!(render_formula(&[0x23, 1, 0, 0, 0], Some(&context)), None);
        assert_eq!(render_formula(&[0x23, 3, 0, 0, 0], Some(&context)), None);
        assert_eq!(render_formula(&[0x23, 0, 0, 0, 0], Some(&context)), None);
    }

    #[test]
    fn rejects_external_or_malformed_sheet_context() {
        let mut context = FormulaContext::default();
        context.add_sup_book(&[1, 0, 1, 0, 0]);
        context.add_extern_sheet(&[1, 0, 0, 0, 0, 0, 0, 0]).unwrap();
        context.set_sheet_names(vec!["Sheet1".to_string()]);
        assert_eq!(
            render_formula(&[0x3a, 0, 0, 0, 0, 0, 0], Some(&context)),
            None
        );
        assert!(context.add_extern_sheet(&[2, 0, 0, 0]).is_err());
    }

    #[test]
    fn renders_shared_relative_references_at_each_origin() {
        let tokens = [
            0x4c, 0xff, 0xff, 0xff, 0xc0, // previous row/column
            0x1e, 0x02, 0x00, 0x05, // * 2
        ];
        assert_eq!(
            render_shared_formula(&tokens, None, 5, 3).as_deref(),
            Some("=(C5*2)")
        );
        assert_eq!(
            render_shared_formula(&tokens, None, 6, 4).as_deref(),
            Some("=(D6*2)")
        );
        assert_eq!(ptg_exp_anchor(&[0x01, 5, 0, 3, 0]), Some((5, 3)));
        assert_eq!(ptg_exp_anchor(&[0x01, 5, 0]), None);
    }

    fn push_biff_unicode(data: &mut Vec<u8>, value: &str) {
        let units = value.encode_utf16().collect::<Vec<_>>();
        data.extend_from_slice(&(units.len() as u16).to_le_bytes());
        let compressed = units.iter().all(|unit| *unit <= 0xff);
        data.push(u8::from(!compressed));
        for unit in units {
            if compressed {
                data.push(unit as u8);
            } else {
                data.extend_from_slice(&unit.to_le_bytes());
            }
        }
    }

    #[test]
    fn renders_inert_external_workbook_cell_and_sheet_range_references() {
        let mut sup_book = 2u16.to_le_bytes().to_vec();
        push_biff_unicode(&mut sup_book, "\u{1}\u{1}C\u{3}Book.xls");
        push_biff_unicode(&mut sup_book, "Data One");
        push_biff_unicode(&mut sup_book, "Data Two");
        let mut context = FormulaContext::default();
        context.add_sup_book(&sup_book);
        context
            .add_extern_sheet(&[2, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0])
            .unwrap();

        assert_eq!(
            render_formula(&[0x3a, 0, 0, 0, 0, 0, 0], Some(&context)).as_deref(),
            Some("='[C:\\Book.xls]Data One'!$A$1")
        );
        assert_eq!(
            render_formula(&[0x3b, 1, 0, 0, 0, 1, 0, 0, 0, 1, 0], Some(&context)).as_deref(),
            Some("='[C:\\Book.xls]Data One:Data Two'!$A$1:$B$2")
        );
    }

    #[test]
    fn malformed_or_add_in_sup_books_remain_lossless_and_unrendered() {
        let mut malformed = 1u16.to_le_bytes().to_vec();
        malformed.extend_from_slice(&[1, 0, 0x80, b'X']);
        let mut context = FormulaContext::default();
        context.add_sup_book(&malformed);
        context.add_sup_book(&[1, 0, 1, 0x3a]);
        context
            .add_extern_sheet(&[2, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0xff, 0xff, 0xff, 0xff])
            .unwrap();
        assert_eq!(
            render_formula(&[0x3a, 0, 0, 0, 0, 0, 0], Some(&context)),
            None
        );
        assert_eq!(
            render_formula(&[0x3a, 1, 0, 0, 0, 0, 0], Some(&context)),
            None
        );
    }

    #[test]
    fn renders_external_and_internal_name_x_tokens_by_contextual_index() {
        let mut external = FormulaContext::default();
        external.add_sup_book(&[1, 0, 1, 0x3a]);
        external
            .add_extern_sheet(&[1, 0, 0, 0, 0xfe, 0xff, 0xfe, 0xff])
            .unwrap();
        external.external_names = vec![vec!["ISODD".to_string(), "RemoteName".to_string()]];
        assert_eq!(
            render_formula(&[0x59, 0, 0, 2, 0, 0, 0], Some(&external)).as_deref(),
            Some("=RemoteName")
        );

        let mut internal = FormulaContext::default();
        internal.add_sup_book(&[1, 0, 1, 4]);
        internal
            .add_extern_sheet(&[1, 0, 0, 0, 0xfe, 0xff, 0xfe, 0xff])
            .unwrap();
        internal.set_sheet_names(vec!["Data One".to_string()]);
        internal.set_scoped_defined_names(vec![Some(("LocalRate".to_string(), Some(0)))]);
        assert_eq!(
            render_formula(&[0x39, 0, 0, 1, 0, 0, 0], Some(&internal)).as_deref(),
            Some("='Data One'!LocalRate")
        );
    }

    #[test]
    fn rejects_invalid_name_x_indices_and_reserved_field() {
        let mut context = FormulaContext::default();
        context.add_sup_book(&[1, 0, 1, 0x3a]);
        context
            .add_extern_sheet(&[1, 0, 0, 0, 0xfe, 0xff, 0xfe, 0xff])
            .unwrap();
        context.external_names = vec![vec!["Fn".to_string()]];
        for tokens in [
            [0x39, 0, 0, 0, 0, 0, 0],
            [0x39, 0, 0, 2, 0, 0, 0],
            [0x39, 0, 0, 1, 0, 1, 0],
        ] {
            assert_eq!(render_formula(&tokens, Some(&context)), None);
        }
    }

    #[test]
    fn renders_inert_add_in_functions_from_name_x_and_func_var() {
        let mut context = FormulaContext::default();
        context.add_sup_book(&[1, 0, 1, 0x3a]);
        context
            .add_extern_sheet(&[1, 0, 0, 0, 0xfe, 0xff, 0xfe, 0xff])
            .unwrap();
        context.external_names = vec![vec!["ISODD".to_string()]];

        assert_eq!(
            render_formula(
                &[
                    0x39, 0, 0, 1, 0, 0, 0, // NameX ISODD
                    0x1e, 3, 0, // integer 3
                    0x42, 2, 0xff, 0, // external FuncVar with two operands
                ],
                Some(&context),
            )
            .as_deref(),
            Some("=ISODD(3)")
        );
        assert_eq!(
            render_formula(
                &[
                    0x39, 0, 0, 1, 0, 0, 0, // NameX ISODD
                    0x42, 1, 0xff, 0, // external FuncVar with no call arguments
                ],
                Some(&context),
            )
            .as_deref(),
            Some("=ISODD()")
        );
    }

    #[test]
    fn rejects_external_func_var_without_leading_name_x() {
        assert_eq!(render_formula(&[0x1e, 3, 0, 0x42, 1, 0xff, 0], None), None);
        assert_eq!(render_formula(&[0x42, 0, 0xff, 0], None), None);
    }

    #[test]
    fn renders_complete_biff_builtin_metadata_and_checks_arity() {
        assert_eq!(
            render_formula(&[0x1e, 0, 0, 0x41, 15, 0], None).as_deref(),
            Some("=SIN(0)")
        );
        assert_eq!(
            render_formula(&[0x1e, 2, 0, 0x1e, 8, 0, 0x41, 0x51, 0x01], None,).as_deref(),
            Some("=POWER(2,8)")
        );
        assert_eq!(
            render_formula(&[0x1e, 2, 0, 0x1e, 8, 0, 0x42, 2, 0xe3, 0], None,).as_deref(),
            Some("=MEDIAN(2,8)")
        );
        assert_eq!(
            render_formula(&[0x1e, 1, 0, 0x1e, 2, 0, 0x42, 2, 102, 0], None),
            None
        );
    }

    #[test]
    fn renders_command_equivalents_as_inert_text_only() {
        assert_eq!(
            render_formula(&[0x42, 0, 0, 0x80], None).as_deref(),
            Some("=BEEP()")
        );
        assert_eq!(render_formula(&[0x42, 0, 0x29, 0x83], None), None);
    }
}
