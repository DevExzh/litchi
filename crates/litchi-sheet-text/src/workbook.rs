//! The one-sheet delimited text workbook.

use std::fs::File;
use std::io::{BufReader, Read, Write};
use std::path::Path;

use litchi_core::BomKind;
use litchi_core::sheet::{CellValue, WorkbookTrait};

use crate::delimited::{self, Options};
use crate::error::{Error, Result};
use crate::worksheet::{Worksheet, Worksheets};

/// A delimited text file read as a workbook of one worksheet named
/// `Sheet1`.
///
/// The rows are read into memory under the [`Options`]' limits. The
/// workbook implements the generic sheet [`WorkbookTrait`]; its worksheet
/// rows and cells are one-based there, as in every workbook format.
///
/// ```
/// use litchi_core::sheet::{CellValue, Worksheet as _};
/// use litchi_sheet_text::Workbook;
/// use litchi_sheet_text::delimited::Options;
///
/// let workbook = Workbook::from_bytes_with_options(b"a\tb\n1\t2", Options::tsv())?;
/// assert_eq!(workbook.rows()[1], [CellValue::Int(1), CellValue::Int(2)]);
/// let sheet = workbook.worksheet();
/// assert_eq!(sheet.cell_value(2, 2)?.as_ref(), &CellValue::Int(2));
///
/// let mut csv = Vec::new();
/// workbook.write_to(&mut csv)?;
/// assert_eq!(csv, b"a\tb\n1\t2");
/// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
/// ```
#[derive(Debug, Clone, PartialEq)]
pub struct Workbook {
    rows: Vec<Vec<CellValue>>,
    options: Options,
    name: String,
    bom: Option<BomKind>,
}

/// The worksheet name of every text workbook.
const SHEET_NAME: &str = "Sheet1";

impl Workbook {
    /// A workbook of `rows` that writes with `options`.
    ///
    /// ```
    /// use litchi_core::sheet::CellValue;
    /// use litchi_sheet_text::Workbook;
    /// use litchi_sheet_text::delimited::Options;
    ///
    /// let workbook = Workbook::new(vec![vec![CellValue::Int(1)]], Options::csv()).with_name("Data");
    /// assert_eq!((workbook.name(), workbook.to_bytes()?), ("Data", b"1".to_vec()));
    /// # Ok::<(), litchi_sheet_text::Error>(())
    /// ```
    #[must_use]
    pub fn new(rows: Vec<Vec<CellValue>>, options: Options) -> Self {
        Self {
            rows,
            options,
            name: SHEET_NAME.to_owned(),
            bom: None,
        }
    }

    /// The same workbook with its worksheet named `name`.
    #[must_use]
    pub fn with_name(mut self, name: impl Into<String>) -> Self {
        self.name = name.into();
        self
    }

    /// Open a CSV file with the default [`Options`].
    ///
    /// # Errors
    ///
    /// The errors of [`Self::from_reader_with_options`].
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        Self::open_with_options(path, Options::default())
    }

    /// Open a delimited file with `options`.
    ///
    /// # Errors
    ///
    /// The errors of [`Self::from_reader_with_options`].
    pub fn open_with_options(path: impl AsRef<Path>, options: Options) -> Result<Self> {
        let file = File::open(path)?;
        Self::from_reader_with_options(BufReader::new(file), options)
    }

    /// Read CSV bytes with the default [`Options`].
    ///
    /// # Errors
    ///
    /// The errors of [`Self::from_reader_with_options`].
    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        Self::from_reader_with_options(bytes, Options::default())
    }

    /// Read delimited bytes with `options`.
    ///
    /// # Errors
    ///
    /// The errors of [`Self::from_reader_with_options`].
    pub fn from_bytes_with_options(bytes: &[u8], options: Options) -> Result<Self> {
        Self::from_reader_with_options(bytes, options)
    }

    /// Read CSV from `reader` with the default [`Options`].
    ///
    /// # Errors
    ///
    /// The errors of [`Self::from_reader_with_options`].
    pub fn from_reader(reader: impl Read) -> Result<Self> {
        Self::from_reader_with_options(reader, Options::default())
    }

    /// Read delimited text from `reader` with `options`.
    ///
    /// # Errors
    ///
    /// [`Error::Io`], [`Error::Encoding`], [`Error::InvalidOption`],
    /// [`Error::UnterminatedQuote`] and [`Error::Limit`].
    pub fn from_reader_with_options(reader: impl Read, options: Options) -> Result<Self> {
        let (rows, bom) = delimited::read_with_bom(reader, &options)?;
        Ok(Self {
            rows,
            options,
            name: SHEET_NAME.to_owned(),
            bom,
        })
    }

    /// The options the workbook was read with, and writes with.
    #[must_use]
    pub const fn options(&self) -> &Options {
        &self.options
    }

    /// The worksheet name.
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// The rows, zero-based.
    #[must_use]
    pub fn rows(&self) -> &[Vec<CellValue>] {
        &self.rows
    }

    /// The byte-order mark the source started with, if any.
    #[must_use]
    pub const fn bom(&self) -> Option<BomKind> {
        self.bom
    }

    /// The worksheet, borrowing the rows.
    #[must_use]
    pub fn worksheet(&self) -> Worksheet<'_> {
        Worksheet::new(&self.rows, &self.name)
    }

    /// Write the rows with the workbook's options ([`delimited::write`]).
    ///
    /// # Errors
    ///
    /// [`Error::Io`] and [`Error::InvalidOption`].
    pub fn write_to(&self, writer: impl Write) -> Result<()> {
        delimited::write(&self.rows, writer, &self.options)
    }

    /// The rows as delimited bytes.
    ///
    /// # Errors
    ///
    /// [`Error::InvalidOption`].
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        let mut bytes = Vec::new();
        self.write_to(&mut bytes)?;
        Ok(bytes)
    }
}

impl WorkbookTrait for Workbook {
    type Error = Error;
    type Worksheet<'a> = Worksheet<'a>;
    type Worksheets<'a> = Worksheets<'a>;

    fn active_worksheet(&self) -> Result<Worksheet<'_>> {
        Ok(self.worksheet())
    }

    fn worksheet_names(&self) -> &[String] {
        std::slice::from_ref(&self.name)
    }

    fn worksheet_by_name(&self, name: &str) -> Result<Worksheet<'_>> {
        if name == self.name {
            self.active_worksheet()
        } else {
            Err(Error::MissingWorksheet)
        }
    }

    fn worksheet_by_index(&self, index: usize) -> Result<Worksheet<'_>> {
        match index {
            0 => self.active_worksheet(),
            _ => Err(Error::MissingWorksheet),
        }
    }

    fn worksheets(&self) -> Worksheets<'_> {
        Worksheets::new(self.worksheet())
    }

    fn worksheet_count(&self) -> usize {
        1
    }

    fn active_sheet_index(&self) -> usize {
        0
    }

    fn is_1904_date_system(&self) -> bool {
        false
    }
}
