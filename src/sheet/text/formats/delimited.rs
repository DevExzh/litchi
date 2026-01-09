//! Delimited text (CSV/TSV/PRN) handler.

use crate::common::{BomKind, strip_bom, write_bom};
use crate::sheet::{CellValue, Result as SheetResult};
use std::io::{Read, Seek, Write};

#[derive(Debug, Clone)]
pub struct DelimitedConfig {
    pub delimiter: u8,
    pub quote: u8,
    pub comment: Option<u8>,
    pub trim_whitespace: bool,
    pub strip_bom: bool,
    pub write_bom: Option<BomKind>,
}

impl Default for DelimitedConfig {
    fn default() -> Self {
        Self {
            delimiter: b',',
            quote: b'"',
            comment: Some(b'#'),
            trim_whitespace: false,
            strip_bom: false,
            write_bom: Some(BomKind::Utf8),
        }
    }
}

impl DelimitedConfig {
    pub fn csv() -> Self {
        Self::default()
    }

    pub fn tsv() -> Self {
        Self {
            delimiter: b'\t',
            ..Self::default()
        }
    }

    pub fn prn() -> Self {
        Self {
            delimiter: b';',
            ..Self::default()
        }
    }

    pub fn with_write_bom(mut self, bom: Option<BomKind>) -> Self {
        self.write_bom = bom;
        self
    }
}

pub fn read_delimited<R: Read + Seek>(
    reader: &mut R,
    config: DelimitedConfig,
) -> SheetResult<Vec<Vec<CellValue>>> {
    if config.strip_bom {
        let _ = strip_bom(reader)?;
    }

    let mut parser = super::super::parser::TextParser::new(
        reader,
        super::super::workbook::TextConfig {
            delimiter: config.delimiter,
            quote: config.quote,
            comment: config.comment,
            trim_whitespace: config.trim_whitespace,
            has_headers: false,
            max_line_length: 1024 * 1024,
            buffer_size: 8192,
            strip_bom: false,
            write_bom: None,
        },
    );

    let mut data = Vec::new();
    while let Some(row_result) = parser.parse_row()? {
        data.push(row_result?);
    }

    Ok(data)
}

pub fn write_delimited<W: Write>(
    data: &[Vec<CellValue>],
    writer: &mut W,
    config: DelimitedConfig,
) -> SheetResult<()> {
    if let Some(bom) = config.write_bom {
        write_bom(writer, bom)?;
    }

    let delimiter = config.delimiter;
    let quote = config.quote;

    for (row_idx, row) in data.iter().enumerate() {
        for (col_idx, cell) in row.iter().enumerate() {
            if col_idx > 0 {
                writer.write_all(&[delimiter])?;
            }

            let mut needs_quote = false;
            let mut field = match cell {
                CellValue::Empty => String::new(),
                CellValue::Bool(b) => {
                    needs_quote = true;
                    if *b { "TRUE" } else { "FALSE" }.to_string()
                },
                CellValue::Int(i) => i.to_string(),
                CellValue::Float(f) => f.to_string(),
                CellValue::DateTime(dt) => {
                    needs_quote = true;
                    dt.to_string()
                },
                CellValue::String(s) => {
                    if s.contains(char::from(delimiter))
                        || s.contains('\n')
                        || s.contains('\r')
                        || s.contains(char::from(quote))
                    {
                        needs_quote = true;
                    }
                    s.clone()
                },
                CellValue::Error(err) => {
                    needs_quote = true;
                    err.clone()
                },
                CellValue::Formula { formula, .. } => {
                    needs_quote = true;
                    format!("={}", formula)
                },
            };

            if needs_quote {
                field = field.replace(char::from(quote), &format!("{0}{0}", char::from(quote)));
                let mut quoted = String::with_capacity(field.len() + 2);
                quoted.push(char::from(quote));
                quoted.push_str(&field);
                quoted.push(char::from(quote));
                field = quoted;
            }

            writer.write_all(field.as_bytes())?;
        }
        if row_idx + 1 < data.len() {
            writer.write_all(b"\n")?;
        }
    }

    Ok(())
}
