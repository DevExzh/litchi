//! BIFF8 formula token rendering.
//!
//! Formula expressions are stored as reverse-Polish `Ptg` token streams.  This
//! renderer intentionally handles only tokens that can be interpreted without
//! workbook-global tables.  Callers retain the original bytes when rendering
//! returns `None`, so unsupported names and 3-D references remain lossless.

/// Render a context-free BIFF8 formula token stream in A1 notation.
pub(crate) fn render_formula(tokens: &[u8]) -> Option<String> {
    let mut decoder = FormulaDecoder::new(tokens);
    decoder.decode().ok().map(|formula| format!("={formula}"))
}

struct FormulaDecoder<'a> {
    data: &'a [u8],
    pos: usize,
    stack: Vec<String>,
}

impl<'a> FormulaDecoder<'a> {
    fn new(data: &'a [u8]) -> Self {
        Self {
            data,
            pos: 0,
            stack: Vec::new(),
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
            0x24 => {
                let row = self.u16()?;
                let col = self.u16()?;
                self.stack.push(cell_reference(row, col)?);
                Ok(())
            },
            0x25 => {
                let first_row = self.u16()?;
                let last_row = self.u16()?;
                let first_col = self.u16()?;
                let last_col = self.u16()?;
                let first = cell_reference(first_row, first_col)?;
                let last = cell_reference(last_row, last_col)?;
                self.stack.push(format!("{first}:{last}"));
                Ok(())
            },
            // Memory tokens carry evaluator bookkeeping and do not contribute
            // an expression operand.
            0x26 | 0x27 => self.skip(6),
            0x29 => self.skip(2),
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
    use super::render_formula;

    #[test]
    fn renders_constants_operators_and_references() {
        let tokens = [
            0x24, 0x00, 0x00, 0x00, 0xc0, // A1
            0x24, 0x00, 0x00, 0x01, 0xc0, // B1
            0x1e, 0x02, 0x00, // 2
            0x05, // multiply
            0x03, // add
        ];
        assert_eq!(render_formula(&tokens).as_deref(), Some("=(A1+(B1*2))"));
    }

    #[test]
    fn renders_absolute_area_and_optimized_sum() {
        let tokens = [
            0x25, 0x00, 0x00, 0x09, 0x00, 0x00, 0x00, 0x02, 0x00, // $A$1:$C$10
            0x19, 0x10, 0x00, 0x00,
        ];
        assert_eq!(render_formula(&tokens).as_deref(), Some("=SUM($A$1:$C$10)"));
    }

    #[test]
    fn renders_unicode_strings_and_variable_functions() {
        let tokens = [
            0x17, 0x02, 0x01, 0x60, 0x4f, 0x7d, 0x59, // "你好"
            0x1e, 0x02, 0x00, // 2
            0x42, 0x02, 0x50, 0x01, // CONCATENATE, two arguments
        ];
        assert_eq!(
            render_formula(&tokens).as_deref(),
            Some("=CONCATENATE(\"你好\",2)")
        );
    }

    #[test]
    fn rejects_workbook_dependent_and_malformed_tokens() {
        assert_eq!(render_formula(&[0x3a, 0, 0, 0, 0, 0, 0]), None);
        assert_eq!(render_formula(&[0x1f, 0, 0]), None);
        assert_eq!(render_formula(&[0x1e, 1, 0, 0x1e, 2, 0]), None);
    }
}
