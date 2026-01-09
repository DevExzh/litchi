//! SYLK (SYmbolic LinK) format reader/writer.
//!
//! SYLK is a text-based spreadsheet format originally from Multiplan,
//! later supported by Excel and other spreadsheet applications.
//!
//! Format structure:
//! - ID: File identifier and version
//! - P: Page/Print settings
//! - F: Format definition
//! - B: Bounds/dimensions
//! - C: Cell data
//! - E: End of file

use crate::common::{BomKind, strip_bom, write_bom};
use crate::sheet::{CellValue, Result as SheetResult};
use std::io::{BufRead, BufReader, Read, Seek, Write};

#[derive(Debug, Clone)]
pub struct SylkConfig {
    pub strip_bom: bool,
    pub write_bom: Option<BomKind>,
}

impl Default for SylkConfig {
    fn default() -> Self {
        Self {
            strip_bom: true,
            write_bom: None,
        }
    }
}

pub struct SylkParser<R: Read> {
    reader: BufReader<R>,
    max_row: usize,
    max_col: usize,
}

impl<R: Read> SylkParser<R> {
    pub fn new(reader: R) -> Self {
        Self {
            reader: BufReader::new(reader),
            max_row: 0,
            max_col: 0,
        }
    }

    pub fn parse(&mut self) -> SheetResult<Vec<Vec<CellValue>>> {
        let mut data: Vec<Vec<CellValue>> = Vec::new();
        let mut line = String::new();

        while self.reader.read_line(&mut line)? > 0 {
            let trimmed = line.trim();

            if trimmed.is_empty() || trimmed.starts_with("ID") || trimmed.starts_with("E;") {
                line.clear();
                continue;
            }

            if trimmed.starts_with("C;") {
                self.parse_cell_record(trimmed, &mut data)?;
            } else if trimmed.starts_with("B;") {
                self.parse_bounds_record(trimmed)?;
            }

            line.clear();
        }

        Ok(data)
    }

    fn parse_cell_record(
        &mut self,
        record: &str,
        data: &mut Vec<Vec<CellValue>>,
    ) -> SheetResult<()> {
        let mut row: Option<usize> = None;
        let mut col: Option<usize> = None;
        let mut value: Option<String> = None;
        let mut is_formula = false;

        for part in record.split(';').skip(1) {
            if part.is_empty() {
                continue;
            }

            let prefix = part.chars().next().unwrap_or(' ');
            let content = &part[1..];

            match prefix {
                'Y' => row = content.parse::<usize>().ok().map(|r| r.saturating_sub(1)),
                'X' => col = content.parse::<usize>().ok().map(|c| c.saturating_sub(1)),
                'K' => {
                    value = Some(content.trim_matches('"').to_string());
                },
                'E' => {
                    is_formula = true;
                    value = Some(content.to_string());
                },
                _ => {},
            }
        }

        if let (Some(r), Some(c)) = (row, col) {
            while data.len() <= r {
                data.push(Vec::new());
            }
            while data[r].len() <= c {
                data[r].push(CellValue::Empty);
            }

            if let Some(val) = value {
                let cell_value = if is_formula {
                    CellValue::Formula {
                        formula: val,
                        cached_value: None,
                        is_array: false,
                        array_range: None,
                    }
                } else {
                    CellValue::infer_from_str(&val)
                };
                data[r][c] = cell_value;
            }

            self.max_row = self.max_row.max(r + 1);
            self.max_col = self.max_col.max(c + 1);
        }

        Ok(())
    }

    fn parse_bounds_record(&mut self, record: &str) -> SheetResult<()> {
        for part in record.split(';').skip(1) {
            if part.is_empty() {
                continue;
            }

            let prefix = part.chars().next().unwrap_or(' ');
            let content = &part[1..];

            match prefix {
                'Y' => {
                    if let Ok(r) = content.parse::<usize>() {
                        self.max_row = r;
                    }
                },
                'X' => {
                    if let Ok(c) = content.parse::<usize>() {
                        self.max_col = c;
                    }
                },
                _ => {},
            }
        }
        Ok(())
    }
}

pub fn read_sylk<R: Read + Seek>(
    reader: &mut R,
    config: SylkConfig,
) -> SheetResult<Vec<Vec<CellValue>>> {
    if config.strip_bom {
        let _ = strip_bom(reader)?;
    }

    let mut parser = SylkParser::new(reader);
    parser.parse()
}

pub fn write_sylk<W: Write>(
    data: &[Vec<CellValue>],
    writer: &mut W,
    config: SylkConfig,
) -> SheetResult<()> {
    if let Some(bom) = config.write_bom {
        write_bom(writer, bom)?;
    }

    writeln!(writer, "ID;PWXL;N;E")?;

    let max_row = data.len();
    let max_col = data.iter().map(|r| r.len()).max().unwrap_or(0);

    if max_row > 0 && max_col > 0 {
        writeln!(writer, "B;Y{};X{}", max_row, max_col)?;
    }

    for (row_idx, row) in data.iter().enumerate() {
        for (col_idx, cell) in row.iter().enumerate() {
            if matches!(cell, CellValue::Empty) {
                continue;
            }

            let row_num = row_idx + 1;
            let col_num = col_idx + 1;

            match cell {
                CellValue::Empty => {},
                CellValue::Int(i) => {
                    writeln!(writer, "C;Y{};X{};K{}", row_num, col_num, i)?;
                },
                CellValue::Float(f) => {
                    writeln!(writer, "C;Y{};X{};K{}", row_num, col_num, f)?;
                },
                CellValue::Bool(b) => {
                    writeln!(
                        writer,
                        "C;Y{};X{};K\"{}\"",
                        row_num,
                        col_num,
                        if *b { "TRUE" } else { "FALSE" }
                    )?;
                },
                CellValue::String(s) => {
                    let escaped = s.replace('\"', "\"\"");
                    writeln!(writer, "C;Y{};X{};K\"{}\"", row_num, col_num, escaped)?;
                },
                CellValue::DateTime(dt) => {
                    writeln!(writer, "C;Y{};X{};K{}", row_num, col_num, dt)?;
                },
                CellValue::Error(err) => {
                    writeln!(writer, "C;Y{};X{};K\"{}\"", row_num, col_num, err)?;
                },
                CellValue::Formula { formula, .. } => {
                    writeln!(writer, "C;Y{};X{};E{}", row_num, col_num, formula)?;
                },
            }
        }
    }

    writeln!(writer, "E")?;
    Ok(())
}
