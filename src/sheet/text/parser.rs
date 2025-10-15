//! Streaming parser for text-based spreadsheet formats

use std::io::{Read, Seek, SeekFrom};
use crate::sheet::{CellValue, Result as SheetResult};

/// Streaming parser for delimited text formats
pub struct TextParser<'a, R: Read + Seek> {
    reader: &'a mut R,
    config: super::workbook::TextConfig,
    buffer: Vec<u8>,
    buffer_pos: usize,
    buffer_len: usize,
    line_start_pos: u64,
}

impl<'a, R: Read + Seek> TextParser<'a, R> {
    /// Create a new text parser
    pub fn new(reader: &'a mut R, config: super::workbook::TextConfig) -> Self {
        // Seek to beginning
        let _ = reader.seek(SeekFrom::Start(0));

        let buffer_size = config.buffer_size;
        TextParser {
            reader,
            config,
            buffer: vec![0; buffer_size],
            buffer_pos: 0,
            buffer_len: 0,
            line_start_pos: 0,
        }
    }

    /// Reset the parser to the beginning
    pub fn reset(&mut self) -> SheetResult<()> {
        self.reader.seek(SeekFrom::Start(0))?;
        self.buffer_pos = 0;
        self.buffer_len = 0;
        self.line_start_pos = 0;
        Ok(())
    }

    /// Parse the next row from the input
    pub fn parse_row(&mut self) -> SheetResult<Option<SheetResult<Vec<CellValue>>>> {
        let mut fields = Vec::new();
        let mut field_start = true;
        let mut in_quotes = false;
        let mut current_field = Vec::new();

        loop {
            // Fill buffer if needed
            if self.buffer_pos >= self.buffer_len {
                self.buffer_len = self.reader.read(&mut self.buffer)?;
                self.buffer_pos = 0;

                if self.buffer_len == 0 {
                    // End of file
                    if !fields.is_empty() || !current_field.is_empty() {
                        // Finish current field if we have data
                        self.finish_field(&mut current_field, &mut fields);
                        return Ok(Some(Ok(fields)));
                    }
                    return Ok(None);
                }
            }

            let byte = self.buffer[self.buffer_pos];
            self.buffer_pos += 1;

            match byte {
                b'\n' => {
                    // End of line
                    if in_quotes {
                        // Newline inside quotes is part of the field
                        current_field.push(byte);
                    } else {
                        // Finish current field and return row
                        self.finish_field(&mut current_field, &mut fields);
                        return Ok(Some(Ok(fields)));
                    }
                }
                b'\r' => {
                    // Handle CRLF - just skip CR, let LF handle the line end
                    if !in_quotes {
                        continue;
                    } else {
                        current_field.push(byte);
                    }
                }
                quote if quote == self.config.quote => {
                    if in_quotes {
                        // This might be a closing quote or escaped quote
                        if self.buffer_pos < self.buffer_len && self.buffer[self.buffer_pos] == self.config.quote {
                            // Escaped quote (doubled quote) - include one quote and skip the next
                            current_field.push(self.config.quote);
                            self.buffer_pos += 1;
                        } else {
                            // Closing quote - don't include it in the field
                            in_quotes = false;
                        }
                    } else {
                        // Opening quote
                        in_quotes = true;
                        field_start = false; // Don't treat this as field start
                    }
                }
                delim if delim == self.config.delimiter && !in_quotes => {
                    // Field separator
                    self.finish_field(&mut current_field, &mut fields);
                    field_start = true;
                }
                b'\\' if in_quotes => {
                    // Handle escape sequences inside quotes
                    if self.buffer_pos < self.buffer_len {
                        let next_byte = self.buffer[self.buffer_pos];
                        self.buffer_pos += 1;
                        match next_byte {
                            b'n' => current_field.push(b'\n'),
                            b'r' => current_field.push(b'\r'),
                            b't' => current_field.push(b'\t'),
                            b'\\' => current_field.push(b'\\'),
                            quote if quote == self.config.quote => current_field.push(quote),
                            _ => {
                                // Unknown escape, include both characters
                                current_field.push(byte);
                                current_field.push(next_byte);
                            }
                        }
                    } else {
                        current_field.push(byte);
                    }
                }
                _ => {
                    // Regular character
                    if field_start && self.config.comment == Some(byte) && fields.is_empty() && !in_quotes {
                        // Comment line, skip to end of line
                        while self.buffer_pos < self.buffer_len {
                            let b = self.buffer[self.buffer_pos];
                            self.buffer_pos += 1;
                            if b == b'\n' {
                                break;
                            }
                        }
                        // Recursively parse next line
                        return self.parse_row();
                    }

                    current_field.push(byte);
                    field_start = false;
                }
            }
        }
    }

    /// Finish parsing a field and add it to the fields vector
    fn finish_field(&self, current_field: &mut Vec<u8>, fields: &mut Vec<CellValue>) {
        let mut field_bytes = std::mem::take(current_field);

        // Trim whitespace if configured
        if self.config.trim_whitespace {
            // Trim from both ends
            let start = field_bytes.iter().position(|&b| !b.is_ascii_whitespace()).unwrap_or(field_bytes.len());
            let end = field_bytes.iter().rposition(|&b| !b.is_ascii_whitespace()).map(|i| i + 1).unwrap_or(0);
            if start < end {
                field_bytes = field_bytes[start..end].to_vec();
            } else {
                field_bytes.clear();
            }
        }

        // Convert to string and determine cell value type
        let field_str = match String::from_utf8(field_bytes) {
            Ok(s) => s,
            Err(e) => {
                // Handle invalid UTF-8 by replacing invalid sequences
                let valid_bytes = e.into_bytes();
                String::from_utf8_lossy(&valid_bytes).to_string()
            }
        };

        let cell_value = if field_str.is_empty() {
            CellValue::Empty
        } else if let Ok(int_val) = field_str.parse::<i64>() {
            CellValue::Int(int_val)
        } else if let Ok(float_val) = fast_float2::parse(&field_str) {
            CellValue::Float(float_val)
        } else {
            // Check for boolean values (case insensitive)
            match field_str.to_lowercase().as_str() {
                "true" | "1" | "yes" | "on" => CellValue::Bool(true),
                "false" | "0" | "no" | "off" => CellValue::Bool(false),
                _ => CellValue::String(field_str),
            }
        };

        fields.push(cell_value);
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Cursor;
    use crate::sheet::CellValue;

    #[test]
    fn test_simple_csv_parsing() {
        let csv = "name,age,city\nJohn,25,New York\nJane,30,London";
        let config = super::super::workbook::TextConfig::default();
        let mut cursor = Cursor::new(csv.as_bytes());
        let mut parser = TextParser::new(&mut cursor, config);

        // First row (headers)
        let row1 = parser.parse_row().unwrap().unwrap().unwrap();
        assert_eq!(row1.len(), 3);
        assert_eq!(row1[0], CellValue::String("name".to_string()));
        assert_eq!(row1[1], CellValue::String("age".to_string()));
        assert_eq!(row1[2], CellValue::String("city".to_string()));

        // Second row
        let row2 = parser.parse_row().unwrap().unwrap().unwrap();
        assert_eq!(row2.len(), 3);
        assert_eq!(row2[0], CellValue::String("John".to_string()));
        assert_eq!(row2[1], CellValue::Int(25));
        assert_eq!(row2[2], CellValue::String("New York".to_string()));

        // Third row
        let row3 = parser.parse_row().unwrap().unwrap().unwrap();
        assert_eq!(row3.len(), 3);
        assert_eq!(row3[0], CellValue::String("Jane".to_string()));
        assert_eq!(row3[1], CellValue::Int(30));
        assert_eq!(row3[2], CellValue::String("London".to_string()));

        // End of file
        assert!(parser.parse_row().unwrap().is_none());
    }

    #[test]
    fn test_quoted_fields() {
        let csv = "\"Hello, World\",\"Value with \"\"quotes\"\"\",\"Normal\"";
        let config = super::super::workbook::TextConfig::default();
        let mut cursor = Cursor::new(csv.as_bytes());
        let mut parser = TextParser::new(&mut cursor, config);

        let row = parser.parse_row().unwrap().unwrap().unwrap();
        assert_eq!(row.len(), 3);
        assert_eq!(row[0], CellValue::String("Hello, World".to_string()));
        assert_eq!(row[1], CellValue::String("Value with \"quotes\"".to_string()));
        assert_eq!(row[2], CellValue::String("Normal".to_string()));
    }

    #[test]
    fn test_tsv_parsing() {
        let tsv = "name\tage\tcity\nJohn\t25\tNew York";
        let config = super::super::workbook::TextConfig::tsv();
        let mut cursor = Cursor::new(tsv.as_bytes());
        let mut parser = TextParser::new(&mut cursor, config);

        let row1 = parser.parse_row().unwrap().unwrap().unwrap();
        assert_eq!(row1.len(), 3);
        assert_eq!(row1[0], CellValue::String("name".to_string()));
        assert_eq!(row1[1], CellValue::String("age".to_string()));
        assert_eq!(row1[2], CellValue::String("city".to_string()));

        let row2 = parser.parse_row().unwrap().unwrap().unwrap();
        assert_eq!(row2.len(), 3);
        assert_eq!(row2[0], CellValue::String("John".to_string()));
        assert_eq!(row2[1], CellValue::Int(25));
        assert_eq!(row2[2], CellValue::String("New York".to_string()));
    }
}
