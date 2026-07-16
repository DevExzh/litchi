//! BIFF8 formula token rendering.
//!
//! Formula expressions are stored as reverse-Polish `Ptg` token streams.  This
//! renderer intentionally handles only tokens that can be interpreted without
//! workbook-global tables.  Callers retain the original bytes when rendering
//! returns `None`, so unsupported names and 3-D references remain lossless.

/// Workbook-global context needed to resolve BIFF `ixti` sheet references.
#[derive(Debug, Default)]
pub(crate) struct FormulaContext {
    sup_books: Vec<SupBookKind>,
    extern_sheets: Vec<ExternSheetRef>,
    sheet_names: Vec<String>,
    defined_names: Vec<Option<String>>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum SupBookKind {
    Internal,
    Other,
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

    pub(crate) fn set_defined_names(&mut self, defined_names: Vec<Option<String>>) {
        self.defined_names = defined_names;
    }

    fn defined_name(&self, one_based_index: u32) -> Option<&str> {
        let index = usize::try_from(one_based_index.checked_sub(1)?).ok()?;
        self.defined_names.get(index)?.as_deref()
    }

    fn sheet_prefix(&self, extern_sheet: u16) -> Option<String> {
        let reference = self.extern_sheets.get(usize::from(extern_sheet))?;
        if self.sup_books.get(usize::from(reference.sup_book))? != &SupBookKind::Internal {
            return None;
        }
        let first = usize::try_from(reference.first_sheet).ok()?;
        let last = usize::try_from(reference.last_sheet).ok()?;
        let first_name = self.sheet_names.get(first)?;
        let last_name = self.sheet_names.get(last)?;
        let name = if first == last {
            escape_sheet_name(first_name)
        } else {
            format!(
                "{}:{}",
                escape_sheet_name(first_name),
                escape_sheet_name(last_name)
            )
        };
        Some(format!("'{name}'!"))
    }
}

fn escape_sheet_name(name: &str) -> String {
    name.replace('\'', "''")
}

/// Render a BIFF8 formula token stream in A1 notation.
pub(crate) fn render_formula(tokens: &[u8], context: Option<&FormulaContext>) -> Option<String> {
    let mut decoder = FormulaDecoder::new(tokens, context, None);
    decoder.decode().ok().map(|formula| format!("={formula}"))
}

/// Render a shared formula template at a particular formula-cell origin.
pub(crate) fn render_shared_formula(
    tokens: &[u8],
    context: Option<&FormulaContext>,
    row: u16,
    column: u16,
) -> Option<String> {
    let mut decoder = FormulaDecoder::new(tokens, context, Some((row, column)));
    decoder.decode().ok().map(|formula| format!("={formula}"))
}

/// Return the shared/array formula anchor encoded by a standalone `PtgExp`.
pub(crate) fn ptg_exp_anchor(tokens: &[u8]) -> Option<(u16, u16)> {
    if tokens.len() != 5 || tokens[0] & 0x7f != 0x01 {
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
    context: Option<&'a FormulaContext>,
    shared_origin: Option<(u16, u16)>,
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
            context,
            shared_origin,
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
        self.stack.pop().ok_or(())
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
                self.stack.push(format!("({value})"));
                Ok(())
            },
            0x16 => {
                self.stack.push(String::new());
                Ok(())
            },
            0x17 => {
                let value = self.formula_string()?;
                self.stack
                    .push(format!("\"{}\"", value.replace('"', "\"\"")));
                Ok(())
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
                self.stack.push(error.to_string());
                Ok(())
            },
            0x1d => {
                let value = match self.byte()? {
                    0 => "FALSE",
                    1 => "TRUE",
                    _ => return Err(()),
                };
                self.stack.push(value.to_string());
                Ok(())
            },
            0x1e => {
                let value = self.u16()?;
                self.stack.push(value.to_string());
                Ok(())
            },
            0x1f => {
                let value = self.f64()?;
                if !value.is_finite() {
                    return Err(());
                }
                self.stack.push(value.to_string());
                Ok(())
            },
            _ => Err(()),
        }
    }

    fn decode_classified(&mut self, opcode: u8) -> Result<(), ()> {
        let base = (opcode & 0x1f) | 0x20;
        match base {
            0x21 => {
                let index = self.u16()?;
                let (name, fixed_args) = function_metadata(index).ok_or(())?;
                self.function(name, fixed_args.ok_or(())?)
            },
            0x22 => {
                let args = self.byte()? as usize;
                let raw_index = self.u16()?;
                if raw_index & 0xf000 != 0 {
                    return Err(());
                }
                let (name, _) = function_metadata(raw_index).ok_or(())?;
                self.function(name, args)
            },
            0x23 => {
                let index = self.u32()?;
                let name = self
                    .context
                    .and_then(|context| context.defined_name(index))
                    .ok_or(())?;
                self.stack.push(name.to_string());
                Ok(())
            },
            0x24 => {
                let row = self.u16()?;
                let col = self.u16()?;
                let (row, col) = resolve_shared_reference(row, col, self.shared_origin)?;
                self.stack.push(cell_reference(row, col)?);
                Ok(())
            },
            0x25 => {
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
                self.stack.push(format!("{first}:{last}"));
                Ok(())
            },
            // Memory tokens carry evaluator bookkeeping and do not contribute
            // an expression operand.
            0x26 | 0x27 => self.skip(6),
            0x29 => self.skip(2),
            0x2c => {
                let row = self.u16()?;
                let col = self.u16()?;
                let origin = self.shared_origin.ok_or(())?;
                let (row, col) = resolve_shared_reference(row, col, Some(origin))?;
                self.stack.push(cell_reference(row, col)?);
                Ok(())
            },
            0x2d => {
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
                self.stack.push(format!("{first}:{last}"));
                Ok(())
            },
            0x3a => {
                let extern_sheet = self.u16()?;
                let row = self.u16()?;
                let col = self.u16()?;
                let (row, col) = resolve_shared_reference(row, col, self.shared_origin)?;
                let prefix = self
                    .context
                    .and_then(|context| context.sheet_prefix(extern_sheet))
                    .ok_or(())?;
                let reference = cell_reference(row, col)?;
                self.stack.push(format!("{prefix}{reference}"));
                Ok(())
            },
            0x3b => {
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
                self.stack.push(format!("{prefix}{first}:{last}"));
                Ok(())
            },
            0x3c => {
                let extern_sheet = self.u16()?;
                self.skip(4)?;
                let prefix = self
                    .context
                    .and_then(|context| context.sheet_prefix(extern_sheet))
                    .ok_or(())?;
                self.stack.push(format!("{prefix}#REF!"));
                Ok(())
            },
            0x3d => {
                let extern_sheet = self.u16()?;
                self.skip(8)?;
                let prefix = self
                    .context
                    .and_then(|context| context.sheet_prefix(extern_sheet))
                    .ok_or(())?;
                self.stack.push(format!("{prefix}#REF!"));
                Ok(())
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
        self.stack.push(format!("({left}{operator}{right})"));
        Ok(())
    }

    fn unary_prefix(&mut self, operator: &str) -> Result<(), ()> {
        let operand = self.pop()?;
        self.stack.push(format!("{operator}{operand}"));
        Ok(())
    }

    fn unary_suffix(&mut self, operator: &str) -> Result<(), ()> {
        let operand = self.pop()?;
        self.stack.push(format!("{operand}{operator}"));
        Ok(())
    }

    fn function(&mut self, name: &str, count: usize) -> Result<(), ()> {
        if self.stack.len() < count {
            return Err(());
        }
        let start = self.stack.len() - count;
        let args = self.stack.split_off(start);
        self.stack.push(format!("{name}({})", args.join(",")));
        Ok(())
    }

    fn formula_string(&mut self) -> Result<String, ()> {
        let count = self.byte()? as usize;
        let flags = self.byte()?;
        if flags & !0x01 != 0 {
            return Err(());
        }
        if flags & 0x01 == 0 {
            let bytes = self.take(count)?;
            Ok(bytes.iter().map(|byte| char::from(*byte)).collect())
        } else {
            let byte_count = count.checked_mul(2).ok_or(())?;
            let bytes = self.take(byte_count)?;
            let units: Vec<u16> = bytes
                .chunks_exact(2)
                .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
                .collect();
            String::from_utf16(&units).map_err(|_| ())
        }
    }

    fn pop(&mut self) -> Result<String, ()> {
        self.stack.pop().ok_or(())
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
        origin_row.wrapping_add_signed(row as i16)
    } else {
        row
    };
    let raw_column = column_flags & 0x00ff;
    let resolved_column = if column_relative {
        let offset = (raw_column as u8) as i8;
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
        result.insert(0, char::from(b'A' + (column % 26) as u8));
        if column < 26 {
            return result;
        }
        column = column / 26 - 1;
    }
}

/// Names and fixed arities for the writer's currently supported built-ins.
/// A `None` arity denotes a variable-argument function.
fn function_metadata(index: u16) -> Option<(&'static str, Option<usize>)> {
    match index {
        0 => Some(("COUNT", None)),
        1 => Some(("IF", None)),
        4 => Some(("SUM", None)),
        5 => Some(("AVERAGE", None)),
        6 => Some(("MIN", None)),
        7 => Some(("MAX", None)),
        24 => Some(("ABS", Some(1))),
        27 => Some(("ROUND", Some(2))),
        31 => Some(("MID", Some(3))),
        32 => Some(("LEN", Some(1))),
        102 => Some(("VLOOKUP", None)),
        115 => Some(("LEFT", None)),
        116 => Some(("RIGHT", None)),
        336 => Some(("CONCATENATE", None)),
        _ => None,
    }
}

#[cfg(test)]
mod tests {
    use super::{FormulaContext, ptg_exp_anchor, render_formula, render_shared_formula};

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
        context
            .add_extern_sheet(&[1, 0, 0, 0, 0, 0, 0, 0])
            .unwrap();
        context.set_sheet_names(vec!["Sheet1".to_string()]);

        assert_eq!(
            render_formula(&[0x3c, 0, 0, 0, 0, 0, 0], Some(&context)).as_deref(),
            Some("='Sheet1'!#REF!")
        );
        assert_eq!(
            render_formula(
                &[0x3d, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
                Some(&context)
            )
            .as_deref(),
            Some("='Sheet1'!#REF!")
        );
    }

    #[test]
    fn renders_one_based_internal_name_without_expanding_it() {
        let mut context = FormulaContext::default();
        context.set_defined_names(vec![None, Some("TaxRate".to_string())]);
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
}
