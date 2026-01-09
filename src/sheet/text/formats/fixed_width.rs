//! Fixed-width PRN format reader/writer.
//!
//! PRN format can be either:
//! - Delimited (space/semicolon separated) - handled by delimited module
//! - Fixed-width columns - handled by this module

use crate::common::{BomKind, strip_bom, write_bom};
use crate::sheet::{CellValue, Result as SheetResult};
use std::io::{BufRead, BufReader, Read, Seek, Write};

#[derive(Debug, Clone)]
pub struct FixedWidthConfig {
    pub strip_bom: bool,
    pub write_bom: Option<BomKind>,
    pub column_widths: Vec<usize>,
    pub auto_detect_widths: bool,
    pub trim_fields: bool,
}

impl Default for FixedWidthConfig {
    fn default() -> Self {
        Self {
            strip_bom: true,
            write_bom: None,
            column_widths: Vec::new(),
            auto_detect_widths: true,
            trim_fields: true,
        }
    }
}

impl FixedWidthConfig {
    pub fn with_column_widths(mut self, widths: Vec<usize>) -> Self {
        self.column_widths = widths;
        self.auto_detect_widths = false;
        self
    }

    pub fn with_auto_detect(mut self, auto: bool) -> Self {
        self.auto_detect_widths = auto;
        self
    }
}

pub struct FixedWidthParser<R: Read> {
    reader: BufReader<R>,
    config: FixedWidthConfig,
}

impl<R: Read> FixedWidthParser<R> {
    pub fn new(reader: R, config: FixedWidthConfig) -> Self {
        Self {
            reader: BufReader::new(reader),
            config,
        }
    }

    pub fn parse(&mut self) -> SheetResult<Vec<Vec<CellValue>>> {
        let mut data = Vec::new();
        let mut lines = Vec::new();
        let mut line = String::new();

        while self.reader.read_line(&mut line)? > 0 {
            if !line.trim().is_empty() {
                lines.push(line.clone());
            }
            line.clear();
        }

        if lines.is_empty() {
            return Ok(data);
        }

        let column_widths = if self.config.auto_detect_widths {
            self.detect_column_widths(&lines)?
        } else {
            self.config.column_widths.clone()
        };

        for line in &lines {
            let row = self.parse_line(line, &column_widths)?;
            if !row.is_empty() {
                data.push(row);
            }
        }

        Ok(data)
    }

    fn detect_column_widths(&self, lines: &[String]) -> SheetResult<Vec<usize>> {
        if lines.is_empty() {
            return Ok(Vec::new());
        }

        let max_len = lines.iter().map(|l| l.len()).max().unwrap_or(0);
        let mut space_positions = vec![0usize; max_len];

        for line in lines {
            for (pos, ch) in line.chars().enumerate() {
                if ch.is_whitespace() {
                    space_positions[pos] += 1;
                }
            }
        }

        let threshold = lines.len() / 2;
        let mut column_boundaries = Vec::new();
        let mut in_space_run = false;
        let mut space_run_start = 0;

        for (pos, &count) in space_positions.iter().enumerate() {
            if count >= threshold && !in_space_run {
                in_space_run = true;
                space_run_start = pos;
            } else if in_space_run {
                let boundary = (space_run_start + pos) / 2;
                column_boundaries.push(boundary);
                in_space_run = false;
            }
        }

        if column_boundaries.is_empty() {
            return Ok(vec![max_len]);
        }

        let mut widths = Vec::new();
        let mut prev = 0;
        for &boundary in &column_boundaries {
            widths.push(boundary - prev);
            prev = boundary;
        }
        widths.push(max_len - prev);

        Ok(widths)
    }

    fn parse_line(&self, line: &str, widths: &[usize]) -> SheetResult<Vec<CellValue>> {
        let mut row = Vec::new();
        let mut pos = 0;

        for &width in widths {
            if pos >= line.len() {
                break;
            }

            let end = (pos + width).min(line.len());
            let field = &line[pos..end];

            let field_str = if self.config.trim_fields {
                field.trim()
            } else {
                field
            };

            let cell_value = CellValue::infer_from_str(field_str);
            row.push(cell_value);
            pos = end;
        }

        Ok(row)
    }
}

pub fn read_fixed_width<R: Read + Seek>(
    reader: &mut R,
    config: FixedWidthConfig,
) -> SheetResult<Vec<Vec<CellValue>>> {
    if config.strip_bom {
        let _ = strip_bom(reader)?;
    }

    let mut parser = FixedWidthParser::new(reader, config);
    parser.parse()
}

pub fn write_fixed_width<W: Write>(
    data: &[Vec<CellValue>],
    writer: &mut W,
    config: FixedWidthConfig,
) -> SheetResult<()> {
    if let Some(bom) = config.write_bom {
        write_bom(writer, bom)?;
    }

    let column_widths = if config.column_widths.is_empty() {
        calculate_column_widths(data)
    } else {
        config.column_widths.clone()
    };

    for row in data {
        let mut line = String::new();

        for (col_idx, cell) in row.iter().enumerate() {
            let width = column_widths.get(col_idx).copied().unwrap_or(10);
            let cell_str = format_cell_value(cell);

            let padded = if cell_str.len() > width {
                cell_str[..width].to_string()
            } else {
                format!("{:<width$}", cell_str, width = width)
            };

            line.push_str(&padded);
        }

        writeln!(writer, "{}", line.trim_end())?;
    }

    Ok(())
}

fn calculate_column_widths(data: &[Vec<CellValue>]) -> Vec<usize> {
    let max_cols = data.iter().map(|r| r.len()).max().unwrap_or(0);
    let mut widths = vec![0usize; max_cols];

    for row in data {
        for (col_idx, cell) in row.iter().enumerate() {
            let cell_str = format_cell_value(cell);
            widths[col_idx] = widths[col_idx].max(cell_str.len() + 2);
        }
    }

    widths
}

fn format_cell_value(cell: &CellValue) -> String {
    match cell {
        CellValue::Empty => String::new(),
        CellValue::Bool(b) => if *b { "TRUE" } else { "FALSE" }.to_string(),
        CellValue::Int(i) => i.to_string(),
        CellValue::Float(f) => f.to_string(),
        CellValue::DateTime(dt) => dt.to_string(),
        CellValue::String(s) => s.clone(),
        CellValue::Error(err) => err.clone(),
        CellValue::Formula {
            cached_value,
            formula,
            ..
        } => {
            if let Some(cached) = cached_value {
                format_cell_value(cached)
            } else {
                format!("={}", formula)
            }
        },
    }
}
