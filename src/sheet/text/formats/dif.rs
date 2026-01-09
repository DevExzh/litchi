//! DIF (Data Interchange Format) reader/writer.
//!
//! DIF is a text-based format for spreadsheet data interchange.
//! Format structure:
//! - TABLE section: Metadata (version, vectors, tuples, data)
//! - VECTORS section: Column count
//! - TUPLES section: Row count
//! - DATA section: Cell data with type indicators
//!   - -1,0 = BOT (Beginning of Tuple/Row)
//!   - 0,value = Numeric
//!   - 1,0 = String (followed by quoted string on next line)
//!   - 1,0 with V = Boolean TRUE/FALSE

use crate::common::{BomKind, strip_bom, write_bom};
use crate::sheet::{CellValue, Result as SheetResult};
use std::io::{BufRead, BufReader, Read, Seek, Write};

#[derive(Debug, Clone)]
pub struct DifConfig {
    pub strip_bom: bool,
    pub write_bom: Option<BomKind>,
}

impl Default for DifConfig {
    fn default() -> Self {
        Self {
            strip_bom: true,
            write_bom: None,
        }
    }
}

pub struct DifParser<R: Read> {
    reader: BufReader<R>,
    vectors: usize,
    tuples: usize,
}

impl<R: Read> DifParser<R> {
    pub fn new(reader: R) -> Self {
        Self {
            reader: BufReader::new(reader),
            vectors: 0,
            tuples: 0,
        }
    }

    pub fn parse(&mut self) -> SheetResult<Vec<Vec<CellValue>>> {
        self.parse_header()?;
        self.parse_data()
    }

    fn parse_header(&mut self) -> SheetResult<()> {
        let mut line = String::new();
        let mut in_data_section = false;

        while self.reader.read_line(&mut line)? > 0 {
            let trimmed = line.trim();

            if trimmed == "DATA" {
                in_data_section = true;
                line.clear();
                break;
            }

            if trimmed == "VECTORS" {
                line.clear();
                self.reader.read_line(&mut line)?;
                if let Some((type_val, _)) = Self::parse_dif_line(&line)
                    && type_val == 0
                {
                    line.clear();
                    self.reader.read_line(&mut line)?;
                    if let Ok(vec_count) = line.trim().trim_matches('"').parse::<usize>() {
                        self.vectors = vec_count;
                    }
                }
            } else if trimmed == "TUPLES" {
                line.clear();
                self.reader.read_line(&mut line)?;
                if let Some((type_val, _)) = Self::parse_dif_line(&line)
                    && type_val == 0
                {
                    line.clear();
                    self.reader.read_line(&mut line)?;
                    if let Ok(tup_count) = line.trim().trim_matches('"').parse::<usize>() {
                        self.tuples = tup_count;
                    }
                }
            }

            line.clear();
        }

        if !in_data_section {
            return Err("Missing DATA section in DIF file".into());
        }

        Ok(())
    }

    fn parse_data(&mut self) -> SheetResult<Vec<Vec<CellValue>>> {
        let mut data = Vec::new();
        let mut current_row: Vec<CellValue> = Vec::new();
        let mut line = String::new();

        while self.reader.read_line(&mut line)? > 0 {
            let trimmed = line.trim();

            if let Some((type_indicator, numeric_val)) = Self::parse_dif_line(trimmed) {
                if type_indicator == -1 && numeric_val == 0.0 {
                    if !current_row.is_empty() {
                        data.push(current_row);
                        current_row = Vec::new();
                    }
                } else if type_indicator == 0 {
                    current_row.push(CellValue::Float(numeric_val));
                } else if type_indicator == 1 {
                    line.clear();
                    self.reader.read_line(&mut line)?;
                    let str_val = line.trim().trim_matches('"').to_string();

                    if numeric_val == 0.0 {
                        match str_val.to_uppercase().as_str() {
                            "TRUE" | "V" => current_row.push(CellValue::Bool(true)),
                            "FALSE" => current_row.push(CellValue::Bool(false)),
                            "NA" | "" => current_row.push(CellValue::Empty),
                            _ => current_row.push(CellValue::String(str_val)),
                        }
                    } else {
                        current_row.push(CellValue::String(str_val));
                    }
                }
            }

            line.clear();
        }

        if !current_row.is_empty() {
            data.push(current_row);
        }

        Ok(data)
    }

    fn parse_dif_line(line: &str) -> Option<(i32, f64)> {
        let parts: Vec<&str> = line.split(',').collect();
        if parts.len() >= 2 {
            let type_indicator = parts[0].trim().parse::<i32>().ok()?;
            let numeric_val = parts[1].trim().parse::<f64>().unwrap_or(0.0);
            Some((type_indicator, numeric_val))
        } else {
            None
        }
    }
}

pub fn read_dif<R: Read + Seek>(
    reader: &mut R,
    config: DifConfig,
) -> SheetResult<Vec<Vec<CellValue>>> {
    if config.strip_bom {
        let _ = strip_bom(reader)?;
    }

    let mut parser = DifParser::new(reader);
    parser.parse()
}

pub fn write_dif<W: Write>(
    data: &[Vec<CellValue>],
    writer: &mut W,
    config: DifConfig,
) -> SheetResult<()> {
    if let Some(bom) = config.write_bom {
        write_bom(writer, bom)?;
    }

    let tuples = data.len();
    let vectors = data.iter().map(|r| r.len()).max().unwrap_or(0);

    writeln!(writer, "TABLE")?;
    writeln!(writer, "0,1")?;
    writeln!(writer, "\"LITCHI\"")?;
    writeln!(writer, "VECTORS")?;
    writeln!(writer, "0,{}", vectors)?;
    writeln!(writer, "\"\"")?;
    writeln!(writer, "TUPLES")?;
    writeln!(writer, "0,{}", tuples)?;
    writeln!(writer, "\"\"")?;
    writeln!(writer, "DATA")?;
    writeln!(writer, "0,0")?;
    writeln!(writer, "\"\"")?;

    for row in data {
        writeln!(writer, "-1,0")?;
        writeln!(writer, "BOT")?;

        for cell in row {
            match cell {
                CellValue::Empty => {
                    writeln!(writer, "1,0")?;
                    writeln!(writer, "\"NA\"")?;
                },
                CellValue::Int(i) => {
                    writeln!(writer, "0,{}", i)?;
                    writeln!(writer, "V")?;
                },
                CellValue::Float(f) => {
                    writeln!(writer, "0,{}", f)?;
                    writeln!(writer, "V")?;
                },
                CellValue::Bool(b) => {
                    writeln!(writer, "1,0")?;
                    writeln!(writer, "\"{}\"", if *b { "TRUE" } else { "FALSE" })?;
                },
                CellValue::String(s) => {
                    writeln!(writer, "1,0")?;
                    writeln!(writer, "\"{}\"", s.replace('\"', "\"\""))?;
                },
                CellValue::DateTime(dt) => {
                    writeln!(writer, "0,{}", dt)?;
                    writeln!(writer, "V")?;
                },
                CellValue::Error(err) => {
                    writeln!(writer, "1,0")?;
                    writeln!(writer, "\"{}\"", err)?;
                },
                CellValue::Formula { cached_value, .. } => {
                    if let Some(cached) = cached_value {
                        match &**cached {
                            CellValue::Int(i) => {
                                writeln!(writer, "0,{}", i)?;
                                writeln!(writer, "V")?;
                            },
                            CellValue::Float(f) => {
                                writeln!(writer, "0,{}", f)?;
                                writeln!(writer, "V")?;
                            },
                            _ => {
                                writeln!(writer, "1,0")?;
                                writeln!(writer, "\"0\"")?;
                            },
                        }
                    } else {
                        writeln!(writer, "1,0")?;
                        writeln!(writer, "\"0\"")?;
                    }
                },
            }
        }
    }

    writeln!(writer, "-1,0")?;
    writeln!(writer, "EOD")?;

    Ok(())
}
