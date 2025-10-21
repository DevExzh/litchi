//! Workbook implementation for text-based formats

use std::fs::File;
use std::io::{BufReader, Read, Seek};
use std::path::Path;

use super::iterators::TextWorksheetIterator;
use crate::sheet::{CellValue, Result as SheetResult, WorkbookTrait, Worksheet, WorksheetIterator};

/// Configuration for parsing text-based spreadsheet files
#[derive(Debug, Clone)]
pub struct TextConfig {
    /// Field delimiter character
    pub delimiter: u8,
    /// Quote character for quoted fields
    pub quote: u8,
    /// Comment character (lines starting with this are ignored)
    pub comment: Option<u8>,
    /// Whether to trim whitespace from fields
    pub trim_whitespace: bool,
    /// Whether the first row contains headers
    pub has_headers: bool,
    /// Maximum line length for memory allocation
    pub max_line_length: usize,
    /// Buffer size for reading
    pub buffer_size: usize,
}

impl Default for TextConfig {
    fn default() -> Self {
        Self {
            delimiter: b',',              // CSV default
            quote: b'"',                  // Standard CSV quoting
            comment: Some(b'#'),          // Common comment character
            trim_whitespace: false,       // Preserve whitespace by default
            has_headers: true,            // Assume first row is headers
            max_line_length: 1024 * 1024, // 1MB max line length
            buffer_size: 8192,            // 8KB buffer
        }
    }
}

impl TextConfig {
    /// Create a new default configuration
    pub fn new() -> Self {
        Self::default()
    }

    /// Set the field delimiter
    pub fn with_delimiter(mut self, delimiter: u8) -> Self {
        self.delimiter = delimiter;
        self
    }

    /// Set the quote character
    pub fn with_quote(mut self, quote: u8) -> Self {
        self.quote = quote;
        self
    }

    /// Set the comment character (None to disable comments)
    pub fn with_comment(mut self, comment: Option<u8>) -> Self {
        self.comment = comment;
        self
    }

    /// Enable/disable whitespace trimming
    pub fn with_trim_whitespace(mut self, trim: bool) -> Self {
        self.trim_whitespace = trim;
        self
    }

    /// Set whether first row contains headers
    pub fn with_headers(mut self, has_headers: bool) -> Self {
        self.has_headers = has_headers;
        self
    }

    /// Set maximum line length
    pub fn with_max_line_length(mut self, max_len: usize) -> Self {
        self.max_line_length = max_len;
        self
    }

    /// Set buffer size
    pub fn with_buffer_size(mut self, size: usize) -> Self {
        self.buffer_size = size;
        self
    }

    /// Create TSV (tab-separated) configuration
    pub fn tsv() -> Self {
        Self::new().with_delimiter(b'\t')
    }

    /// Create PRN (semicolon-separated) configuration
    pub fn prn() -> Self {
        Self::new().with_delimiter(b';')
    }

    /// Create pipe-separated configuration
    pub fn pipe() -> Self {
        Self::new().with_delimiter(b'|')
    }
}

/// Workbook implementation for text-based formats
#[derive(Debug)]
pub struct TextWorkbook {
    data: Vec<Vec<CellValue>>,
    config: TextConfig,
    worksheet_name: String,
}

impl TextWorkbook {
    /// Open a text workbook from a file path with default configuration
    pub fn open<P: AsRef<Path>>(path: P) -> SheetResult<Self> {
        Self::from_path_with_config(path, TextConfig::default())
    }

    /// Open a text workbook from a file path with custom configuration
    pub fn from_path_with_config<P: AsRef<Path>>(path: P, config: TextConfig) -> SheetResult<Self> {
        let file = File::open(path)?;
        let mut reader = BufReader::with_capacity(config.buffer_size, file);
        Self::from_reader(&mut reader, config)
    }

    /// Create a text workbook from any reader with configuration
    pub fn from_reader<R: Read + Seek>(reader: &mut R, config: TextConfig) -> SheetResult<Self> {
        let mut parser = super::parser::TextParser::new(reader, config.clone());
        let mut data = Vec::new();

        while let Some(row_result) = parser.parse_row()? {
            data.push(row_result?);
        }

        let worksheet_name = "Sheet1".to_string();

        Ok(TextWorkbook {
            data,
            config,
            worksheet_name,
        })
    }

    /// Create a text workbook from bytes with configuration
    pub fn from_bytes(bytes: &[u8], config: TextConfig) -> SheetResult<Self> {
        let mut cursor = std::io::Cursor::new(bytes);
        Self::from_reader(&mut cursor, config)
    }

    /// Get a reference to the configuration
    pub fn config(&self) -> &TextConfig {
        &self.config
    }

    /// Get the worksheet name
    pub fn worksheet_name(&self) -> &str {
        &self.worksheet_name
    }

    /// Set the worksheet name
    pub fn set_worksheet_name(&mut self, name: String) {
        self.worksheet_name = name;
    }
}

impl WorkbookTrait for TextWorkbook {
    fn active_worksheet(&self) -> SheetResult<Box<dyn Worksheet + '_>> {
        Ok(Box::new(super::worksheet::TextWorksheet::from_data(
            &self.data,
            self.worksheet_name.clone(),
        )))
    }

    fn worksheet_names(&self) -> Vec<String> {
        vec![self.worksheet_name.clone()]
    }

    fn worksheet_by_name(&self, name: &str) -> SheetResult<Box<dyn Worksheet + '_>> {
        if name == self.worksheet_name {
            self.active_worksheet()
        } else {
            Err(format!("Worksheet '{}' not found", name).into())
        }
    }

    fn worksheet_by_index(&self, index: usize) -> SheetResult<Box<dyn Worksheet + '_>> {
        match index {
            0 => self.active_worksheet(),
            _ => Err(format!("Worksheet index {} out of range", index).into()),
        }
    }

    fn worksheets(&self) -> Box<dyn WorksheetIterator<'_> + '_> {
        Box::new(TextWorksheetIterator::new(self))
    }

    fn worksheet_count(&self) -> usize {
        1
    }

    fn active_sheet_index(&self) -> usize {
        0
    }
}
