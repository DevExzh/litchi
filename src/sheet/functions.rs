//! Functions for opening workbooks.

use super::types::Result;
use super::traits::WorkbookTrait;

/// Open a workbook from a file path.
///
/// **Note**: This requires the `ooxml` feature to be enabled.
#[cfg(feature = "ooxml")]
pub fn open_workbook<P: AsRef<std::path::Path>>(path: P) -> Result<Box<dyn WorkbookTrait>> {
    let package = crate::ooxml::opc::OpcPackage::open(path)?;
    let workbook = crate::ooxml::xlsx::Workbook::new(package)?;
    Ok(Box::new(workbook))
}

/// Open a workbook from bytes.
///
/// **Note**: This requires the `ooxml` feature to be enabled.
#[cfg(feature = "ooxml")]
pub fn open_workbook_from_bytes(bytes: &[u8]) -> Result<Box<dyn WorkbookTrait>> {
    use std::io::Cursor;
    let cursor = Cursor::new(bytes);
    let package = crate::ooxml::opc::OpcPackage::from_reader(cursor)?;
    let workbook = crate::ooxml::xlsx::Workbook::new(package)?;
    Ok(Box::new(workbook))
}

/// Open an XLS workbook from a file path.
///
/// **Note**: This requires the `ole` feature to be enabled.
#[cfg(feature = "ole")]
pub fn open_xls_workbook<P: AsRef<std::path::Path>>(path: P) -> Result<crate::ole::xls::XlsWorkbook<std::fs::File>> {
    use std::fs::File;
    let file = File::open(path)?;
    let workbook = crate::ole::xls::XlsWorkbook::new(file)?;
    Ok(workbook)
}

/// Open an XLS workbook from bytes.
///
/// **Note**: This requires the `ole` feature to be enabled.
#[cfg(feature = "ole")]
pub fn open_xls_workbook_from_bytes(bytes: &[u8]) -> Result<crate::ole::xls::XlsWorkbook<std::io::Cursor<&[u8]>>> {
    use std::io::Cursor;
    let cursor = Cursor::new(bytes);
    let workbook = crate::ole::xls::XlsWorkbook::new(cursor)?;
    Ok(workbook)
}

/// Open an XLS workbook as a trait object from a file path.
///
/// **Note**: This requires the `ole` feature to be enabled.
#[cfg(feature = "ole")]
pub fn open_xls_workbook_dyn<P: AsRef<std::path::Path>>(path: P) -> Result<Box<dyn WorkbookTrait>> {
    let workbook = open_xls_workbook(path)?;
    Ok(Box::new(workbook))
}

/// Open an XLS workbook as a trait object from bytes.
///
/// **Note**: This requires the `ole` feature to be enabled.
#[cfg(feature = "ole")]
pub fn open_xls_workbook_from_bytes_dyn(bytes: &[u8]) -> Result<Box<dyn WorkbookTrait>> {
    use std::io::Cursor;
    let cursor = Cursor::new(bytes.to_vec());
    let workbook = crate::ole::xls::XlsWorkbook::new(cursor)?;
    Ok(Box::new(workbook))
}

/// Open an XLSB workbook from a file path.
///
/// **Note**: This requires the `ooxml` feature to be enabled.
#[cfg(feature = "ooxml")]
pub fn open_xlsb_workbook<P: AsRef<std::path::Path>>(path: P) -> Result<crate::ooxml::xlsb::XlsbWorkbook> {
    use std::fs::File;
    let file = File::open(path)?;
    let workbook = crate::ooxml::xlsb::XlsbWorkbook::new(file)?;
    Ok(workbook)
}

/// Open an XLSB workbook from bytes.
///
/// **Note**: This requires the `ooxml` feature to be enabled.
#[cfg(feature = "ooxml")]
pub fn open_xlsb_workbook_from_bytes(bytes: &[u8]) -> Result<crate::ooxml::xlsb::XlsbWorkbook> {
    use std::io::Cursor;
    let cursor = Cursor::new(bytes);
    let workbook = crate::ooxml::xlsb::XlsbWorkbook::new(cursor)?;
    Ok(workbook)
}

/// Open an XLSB workbook as a trait object from a file path.
///
/// **Note**: This requires the `ooxml` feature to be enabled.
#[cfg(feature = "ooxml")]
pub fn open_xlsb_workbook_dyn<P: AsRef<std::path::Path>>(path: P) -> Result<Box<dyn WorkbookTrait>> {
    let workbook = open_xlsb_workbook(path)?;
    Ok(Box::new(workbook))
}

/// Open an XLSB workbook as a trait object from bytes.
///
/// **Note**: This requires the `ooxml` feature to be enabled.
#[cfg(feature = "ooxml")]
pub fn open_xlsb_workbook_from_bytes_dyn(bytes: &[u8]) -> Result<Box<dyn WorkbookTrait>> {
    use std::io::Cursor;
    let cursor = Cursor::new(bytes.to_vec());
    let workbook = crate::ooxml::xlsb::XlsbWorkbook::new(cursor)?;
    Ok(Box::new(workbook))
}

/// Open a CSV workbook from a file path.
pub fn open_csv_workbook<P: AsRef<std::path::Path>>(path: P) -> Result<Box<dyn WorkbookTrait>> {
    let workbook = crate::sheet::text::TextWorkbook::open(path)?;
    Ok(Box::new(workbook))
}

/// Open a CSV workbook from bytes.
pub fn open_csv_workbook_from_bytes(bytes: &[u8]) -> Result<Box<dyn WorkbookTrait>> {
    let workbook = crate::sheet::text::TextWorkbook::from_bytes(bytes, crate::sheet::text::TextConfig::default())?;
    Ok(Box::new(workbook))
}

/// Open a TSV workbook from a file path.
pub fn open_tsv_workbook<P: AsRef<std::path::Path>>(path: P) -> Result<Box<dyn WorkbookTrait>> {
    let config = crate::sheet::text::TextConfig::tsv();
    let workbook = crate::sheet::text::TextWorkbook::from_path_with_config(path, config)?;
    Ok(Box::new(workbook))
}

/// Open a TSV workbook from bytes.
pub fn open_tsv_workbook_from_bytes(bytes: &[u8]) -> Result<Box<dyn WorkbookTrait>> {
    let config = crate::sheet::text::TextConfig::tsv();
    let workbook = crate::sheet::text::TextWorkbook::from_bytes(bytes, config)?;
    Ok(Box::new(workbook))
}

/// Open a PRN workbook from a file path.
pub fn open_prn_workbook<P: AsRef<std::path::Path>>(path: P) -> Result<Box<dyn WorkbookTrait>> {
    let config = crate::sheet::text::TextConfig::prn();
    let workbook = crate::sheet::text::TextWorkbook::from_path_with_config(path, config)?;
    Ok(Box::new(workbook))
}

/// Open a PRN workbook from bytes.
pub fn open_prn_workbook_from_bytes(bytes: &[u8]) -> Result<Box<dyn WorkbookTrait>> {
    let config = crate::sheet::text::TextConfig::prn();
    let workbook = crate::sheet::text::TextWorkbook::from_bytes(bytes, config)?;
    Ok(Box::new(workbook))
}

/// Open a text workbook with custom configuration from a file path.
pub fn open_text_workbook_with_config<P: AsRef<std::path::Path>>(
    path: P,
    config: crate::sheet::text::TextConfig
) -> Result<Box<dyn WorkbookTrait>> {
    let workbook = crate::sheet::text::TextWorkbook::from_path_with_config(path, config)?;
    Ok(Box::new(workbook))
}

/// Open a text workbook with custom configuration from bytes.
pub fn open_text_workbook_from_bytes_with_config(
    bytes: &[u8],
    config: crate::sheet::text::TextConfig
) -> Result<Box<dyn WorkbookTrait>> {
    let workbook = crate::sheet::text::TextWorkbook::from_bytes(bytes, config)?;
    Ok(Box::new(workbook))
}

