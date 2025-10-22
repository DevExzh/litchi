//! Unified workbook types and format detection.

use crate::common::Error;
use std::io::{Read, Seek, SeekFrom};

// ZIP is needed for refining detection of OOXML, iWork and ODF containers
#[cfg(any(feature = "iwa", feature = "ooxml", feature = "odf"))]
use zip;

type Result<T> = std::result::Result<T, Error>;

/// Internal representation of different workbook implementations.
/// This enum wraps format-specific workbook types, providing
/// a unified API. Users typically don't interact with this enum directly,
/// but instead use the methods on `UnifiedWorkbook`.
#[allow(clippy::large_enum_variant)]
pub(super) enum WorkbookImpl {
    #[cfg(feature = "iwa")]
    Numbers(crate::iwa::numbers::NumbersDocument),

    // OOXML-based formats
    #[cfg(feature = "ooxml")]
    Xlsx(crate::ooxml::xlsx::Workbook),
    #[cfg(feature = "ooxml")]
    Xlsb(crate::ooxml::xlsb::XlsbWorkbook),

    // Legacy OLE-based Excel
    #[cfg(feature = "ole")]
    XlsFile(crate::ole::xls::XlsWorkbook<std::io::BufReader<std::fs::File>>),
    #[cfg(feature = "ole")]
    XlsMem(crate::ole::xls::XlsWorkbook<std::io::Cursor<Vec<u8>>>),

    // OpenDocument Spreadsheet
    #[cfg(feature = "odf")]
    Ods(std::cell::RefCell<crate::odf::Spreadsheet>),

    // For other formats, we just indicate they're not yet fully unified
    #[cfg(any(feature = "ole", feature = "ooxml"))]
    #[allow(dead_code)]
    Other,
}

/// Format of the workbook file.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[allow(unused)] // Since the library is feature-gated, this enum may not be used
pub(super) enum WorkbookFormat {
    /// Legacy Excel Binary Format (.xls)
    Xls,
    /// Office Open XML Workbook (.xlsx)
    Xlsx,
    /// Office Open XML Binary Workbook (.xlsb)
    Xlsb,
    /// OpenDocument Spreadsheet (.ods)
    Ods,
    /// Apple Numbers (.numbers)
    Numbers,
}

/// Detect workbook format from file signature.
pub(super) fn detect_workbook_format_from_signature<R: Read + Seek>(
    reader: &mut R,
) -> Result<WorkbookFormat> {
    let mut header = [0u8; 8];
    reader.seek(SeekFrom::Start(0))?;
    reader.read_exact(&mut header)?;
    reader.seek(SeekFrom::Start(0))?;

    // Check for OLE2 format (XLS)
    if &header[0..8] == b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1" {
        return Ok(WorkbookFormat::Xls);
    }

    // Check for ZIP format (XLSX, XLSB, or Numbers)
    if &header[0..4] == b"PK\x03\x04" {
        // Default to XLSX, will be refined later
        return Ok(WorkbookFormat::Xlsx);
    }

    Err(Error::NotOfficeFile)
}

/// Refine ZIP-based workbook format detection (XLSX vs XLSB vs Numbers)
pub(super) fn refine_workbook_format<R: Read + Seek>(
    reader: &mut R,
    initial_format: WorkbookFormat,
) -> Result<WorkbookFormat> {
    use std::io::SeekFrom;

    // Only refine if it's a ZIP-based format
    if initial_format != WorkbookFormat::Xlsx {
        return Ok(initial_format);
    }

    reader.seek(SeekFrom::Start(0))?;

    // Open ZIP archive once
    let mut archive = match zip::ZipArchive::new(reader) {
        Ok(archive) => archive,
        Err(_) => return Ok(initial_format),
    };

    // Check for ODF by inspecting the mimetype file
    #[cfg(feature = "odf")]
    {
        if let Ok(mut mimetype_file) = archive.by_name("mimetype") {
            use std::io::Read as _;
            let mut mime = String::new();
            // Best-effort read; if it fails, just continue detection
            let _ = mimetype_file.read_to_string(&mut mime);
            let mime_trimmed = mime.trim();
            if mime_trimmed == "application/vnd.oasis.opendocument.spreadsheet"
                || mime_trimmed == "application/vnd.oasis.opendocument.spreadsheet-template"
            {
                return Ok(WorkbookFormat::Ods);
            }
        }
    }

    // Check for iWork Numbers format (Index/*.iwa files)
    #[cfg(feature = "iwa")]
    {
        // Check for Index.zip (older iWork format)
        if archive.by_name("Index.zip").is_ok() {
            return Ok(WorkbookFormat::Numbers);
        }

        // Check for Index/ directory with .iwa files (newer iWork format)
        for i in 0..archive.len() {
            if let Ok(file) = archive.by_index(i) {
                let name = file.name();
                if name.starts_with("Index/") && name.ends_with(".iwa") {
                    return Ok(WorkbookFormat::Numbers);
                }
            }
        }
    }

    // Check for XLSB by looking at the workbook part
    #[cfg(feature = "ooxml")]
    {
        // XLSB uses .bin extension for parts
        if archive.by_name("xl/workbook.bin").is_ok() {
            return Ok(WorkbookFormat::Xlsb);
        }
    }

    Ok(WorkbookFormat::Xlsx)
}
