//! Main Spreadsheet structure and implementation.

use super::Sheet;
use crate::common::{Error, Metadata, Result};
use crate::odf::core::{Content, Meta, Package, Styles};
use std::io::Cursor;
use std::path::Path;

/// An OpenDocument spreadsheet (.ods).
///
/// This struct represents a complete ODS spreadsheet and provides methods to access
/// its sheets, cells, and metadata.
///
/// # Examples
///
/// ```no_run
/// use litchi::odf::Spreadsheet;
///
/// # fn main() -> litchi::Result<()> {
/// let mut spreadsheet = Spreadsheet::open("data.ods")?;
///
/// // Get sheet count
/// println!("Sheets: {}", spreadsheet.sheet_count()?);
///
/// // Access first sheet
/// if let Some(sheet) = spreadsheet.sheet_by_index(0)? {
///     println!("Sheet: {}", sheet.name()?);
///     println!("Rows: {}, Columns: {}", sheet.row_count()?, sheet.column_count()?);
/// }
///
/// // Export to CSV
/// let csv = spreadsheet.to_csv()?;
/// # Ok(())
/// # }
/// ```
pub struct Spreadsheet {
    package: Package<Cursor<Vec<u8>>>,
    #[allow(dead_code)]
    content: Content,
    #[allow(dead_code)]
    styles: Option<Styles>,
    meta: Option<Meta>,
}

impl Spreadsheet {
    /// Open an ODS spreadsheet from a file path.
    ///
    /// # Arguments
    ///
    /// * `path` - Path to the .ods file
    ///
    /// # Errors
    ///
    /// Returns an error if the file cannot be read or is not a valid ODS file.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi::odf::Spreadsheet;
    ///
    /// # fn main() -> litchi::Result<()> {
    /// let spreadsheet = Spreadsheet::open("data.ods")?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self> {
        let bytes = std::fs::read(path.as_ref())?;
        Self::from_bytes(bytes)
    }

    /// Create a Spreadsheet from a byte buffer.
    ///
    /// # Arguments
    ///
    /// * `bytes` - Complete ODS file contents as bytes
    ///
    /// # Errors
    ///
    /// Returns an error if the bytes do not represent a valid ODS file.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi::odf::Spreadsheet;
    ///
    /// # fn main() -> litchi::Result<()> {
    /// let bytes = std::fs::read("data.ods")?;
    /// let spreadsheet = Spreadsheet::from_bytes(bytes)?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        let cursor = Cursor::new(bytes);
        let package = Package::from_reader(cursor)?;

        // Verify this is a spreadsheet
        let mime_type = package.mimetype();
        if !mime_type.contains("opendocument.spreadsheet") {
            return Err(Error::InvalidFormat(format!(
                "Not an ODS file: MIME type is {}",
                mime_type
            )));
        }

        // Parse core components
        let content_bytes = package.get_file("content.xml")?;
        let content = Content::from_bytes(&content_bytes)?;

        let styles = if package.has_file("styles.xml") {
            let styles_bytes = package.get_file("styles.xml")?;
            Some(Styles::from_bytes(&styles_bytes)?)
        } else {
            None
        };

        let meta = if package.has_file("meta.xml") {
            let meta_bytes = package.get_file("meta.xml")?;
            Some(Meta::from_bytes(&meta_bytes)?)
        } else {
            None
        };

        Ok(Self {
            package,
            content,
            styles,
            meta,
        })
    }

    /// Create an ODS spreadsheet from an already-parsed ZIP archive.
    ///
    /// This is used for single-pass parsing where the ZIP archive has already
    /// been parsed during format detection. It avoids double-parsing.
    pub fn from_zip_archive(
        zip_archive: zip::ZipArchive<std::io::Cursor<Vec<u8>>>,
    ) -> Result<Self> {
        let package = Package::from_zip_archive(zip_archive)?;

        // Verify this is a spreadsheet
        let mime_type = package.mimetype();
        if !mime_type.contains("opendocument.spreadsheet") {
            return Err(Error::InvalidFormat(format!(
                "Not an ODS file: MIME type is {}",
                mime_type
            )));
        }

        // Parse core components
        let content_bytes = package.get_file("content.xml")?;
        let content = Content::from_bytes(&content_bytes)?;

        let styles = if package.has_file("styles.xml") {
            let styles_bytes = package.get_file("styles.xml")?;
            Some(Styles::from_bytes(&styles_bytes)?)
        } else {
            None
        };

        let meta = if package.has_file("meta.xml") {
            let meta_bytes = package.get_file("meta.xml")?;
            Some(Meta::from_bytes(&meta_bytes)?)
        } else {
            None
        };

        Ok(Self {
            package,
            content,
            styles,
            meta,
        })
    }

    /// Get the number of sheets in the spreadsheet.
    pub fn sheet_count(&mut self) -> Result<usize> {
        let sheets = self.sheets()?;
        Ok(sheets.len())
    }

    /// Get all sheets in the spreadsheet.
    ///
    /// Returns a vector of `Sheet` objects representing all sheets in the document.
    pub fn sheets(&mut self) -> Result<Vec<Sheet>> {
        use super::parser::OdsParser;

        let content_bytes = self.package.get_file("content.xml")?;
        let content = Content::from_bytes(&content_bytes)?;

        OdsParser::parse_sheets(content.xml_content())
    }

    /// Get a sheet by name.
    ///
    /// Returns `Some(sheet)` if a sheet with the given name exists, `None` otherwise.
    ///
    /// # Arguments
    ///
    /// * `name` - Name of the sheet to find
    pub fn sheet_by_name(&mut self, name: &str) -> Result<Option<Sheet>> {
        let sheets = self.sheets()?;
        Ok(sheets.into_iter().find(|sheet| sheet.name == name))
    }

    /// Get a sheet by index.
    ///
    /// Returns `Some(sheet)` if a sheet exists at the given index, `None` otherwise.
    ///
    /// # Arguments
    ///
    /// * `index` - 0-based index of the sheet
    pub fn sheet_by_index(&mut self, index: usize) -> Result<Option<Sheet>> {
        let sheets = self.sheets()?;
        Ok(sheets.into_iter().nth(index))
    }

    /// Extract all text content from the spreadsheet.
    ///
    /// Returns text from all cells, separated by newlines.
    pub fn text(&mut self) -> Result<String> {
        let sheets = self.sheets()?;
        let mut all_text = Vec::new();

        for sheet in sheets {
            for row in sheet.rows {
                for cell in row.cells {
                    if !cell.text.trim().is_empty() {
                        all_text.push(cell.text.trim().to_string());
                    }
                }
            }
        }

        Ok(all_text.join("\n"))
    }

    /// Export spreadsheet data as CSV.
    ///
    /// Converts all sheets to CSV format, with sheets separated by double newlines.
    /// Properly escapes CSV special characters.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi::odf::Spreadsheet;
    ///
    /// # fn main() -> litchi::Result<()> {
    /// let mut spreadsheet = Spreadsheet::open("data.ods")?;
    /// let csv = spreadsheet.to_csv()?;
    /// std::fs::write("output.csv", csv)?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn to_csv(&mut self) -> Result<String> {
        let sheets = self.sheets()?;
        let mut csv_output = String::new();

        for (sheet_index, sheet) in sheets.iter().enumerate() {
            if sheet_index > 0 {
                csv_output.push_str("\n\n"); // Separate sheets with double newline
            }

            for (row_index, row) in sheet.rows.iter().enumerate() {
                if row_index > 0 {
                    csv_output.push('\n');
                }

                for (col_index, cell) in row.cells.iter().enumerate() {
                    if col_index > 0 {
                        csv_output.push(',');
                    }

                    // Escape CSV special characters and wrap in quotes if needed
                    let cell_text = &cell.text;
                    if cell_text.contains(',')
                        || cell_text.contains('"')
                        || cell_text.contains('\n')
                    {
                        let escaped = cell_text.replace('"', "\"\"");
                        csv_output.push('"');
                        csv_output.push_str(&escaped);
                        csv_output.push('"');
                    } else {
                        csv_output.push_str(cell_text);
                    }
                }
            }
        }

        Ok(csv_output)
    }

    /// Get document metadata.
    ///
    /// Extracts metadata from the meta.xml file.
    pub fn metadata(&self) -> Result<Metadata> {
        if let Some(meta) = &self.meta {
            Ok(meta.extract_metadata())
        } else {
            Ok(Metadata::default())
        }
    }

    // Note: For spreadsheet modification operations, see `MutableSpreadsheet` which provides
    // full CRUD operations on sheets, rows, and cells including set_cell, clear_cell, add/remove
    // rows and sheets.

    /// Save the spreadsheet to a new file.
    ///
    /// This method saves the current spreadsheet state to a new file.
    ///
    /// # Arguments
    ///
    /// * `path` - Path where the ODS file should be saved
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi::odf::Spreadsheet;
    ///
    /// # fn main() -> litchi::Result<()> {
    /// let spreadsheet = Spreadsheet::open("input.ods")?;
    /// spreadsheet.save("output.ods")?;
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// # Note
    ///
    /// Full spreadsheet modification support is planned for future releases. For now,
    /// to modify a spreadsheet, use `SpreadsheetBuilder` to create a new one with
    /// the desired content.
    pub fn save<P: AsRef<std::path::Path>>(&self, path: P) -> Result<()> {
        let bytes = self.to_bytes()?;
        std::fs::write(path, bytes)?;
        Ok(())
    }

    /// Convert the spreadsheet to bytes.
    ///
    /// This method serializes the spreadsheet to an ODF-compliant ZIP archive.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi::odf::Spreadsheet;
    ///
    /// # fn main() -> litchi::Result<()> {
    /// let spreadsheet = Spreadsheet::open("data.ods")?;
    /// let bytes = spreadsheet.to_bytes()?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        use crate::odf::core::PackageWriter;

        let mut writer = PackageWriter::new();

        // Set MIME type
        writer.set_mimetype(self.package.mimetype())?;

        // Add content.xml
        let content_xml = self.content.xml_content();
        writer.add_file("content.xml", content_xml.as_bytes())?;

        // Add styles.xml if present
        if let Some(ref styles) = self.styles {
            let styles_xml = styles.xml_content();
            writer.add_file("styles.xml", styles_xml.as_bytes())?;
        }

        // Add meta.xml if present
        if let Some(ref meta) = self.meta {
            let meta_xml = meta.xml_content();
            writer.add_file("meta.xml", meta_xml.as_bytes())?;
        }

        // Copy settings.xml if present
        if self.package.has_file("settings.xml") {
            let settings_bytes = self.package.get_file("settings.xml")?;
            writer.add_file("settings.xml", &settings_bytes)?;
        }

        // Copy all media files (images, charts, etc.) from the original package
        let media_files = self.package.media_files()?;
        for media_path in media_files {
            if let Ok(media_bytes) = self.package.get_file(&media_path) {
                writer.add_file(&media_path, &media_bytes)?;
            }
        }

        // Copy other common ODF files if they exist
        let other_files = vec!["Thumbnails/thumbnail.png", "Configurations2/"];
        for file_path in other_files {
            if self.package.has_file(file_path)
                && let Ok(file_bytes) = self.package.get_file(file_path)
            {
                writer.add_file(file_path, &file_bytes)?;
            }
        }

        writer.finish_to_bytes()
    }

    // Note: DELETE operations are available via `MutableSpreadsheet`. To modify this spreadsheet:
    //   1. Convert: `let mut mutable = MutableSpreadsheet::from_spreadsheet(spreadsheet)?`
    //   2. Modify: `mutable.remove_sheet(0)?`, `mutable.set_cell(0, 0, 0, value)?`, etc.
    //   3. Save: `mutable.save("output.ods")?`
    // Available methods: remove_sheet, remove_row, set_cell, clear_cell, clear_sheet, etc.
}
