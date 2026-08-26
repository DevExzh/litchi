//! Unified workbook types and format detection.

use litchi_core::Error;
#[cfg(any(feature = "ods", feature = "xlsx", feature = "xlsb"))]
use litchi_core::FileFormat;
use std::io::{Read, Seek, SeekFrom};

type Result<T> = std::result::Result<T, Error>;

/// Internal representation of different workbook implementations.
/// This enum wraps format-specific workbook types, providing
/// a unified API. Users typically don't interact with this enum directly,
/// but instead use the methods on `UnifiedWorkbook`.
#[allow(
    clippy::large_enum_variant,
    reason = "crate-internal facade enum; boxing the large variant would complicate every match for no measurable gain"
)]
pub(super) enum WorkbookImpl {
    #[cfg(feature = "numbers")]
    Numbers(litchi_numbers::Package),

    // OOXML-based formats
    #[cfg(feature = "xlsx")]
    Xlsx(super::adapters::Workbook),
    #[cfg(feature = "xlsb")]
    Xlsb(crate::xlsb::Workbook),
    #[cfg(feature = "xlsb")]
    XlsbSource(super::adapters::XlsbWorkbook),

    // Legacy OLE-based Excel
    #[cfg(feature = "xls")]
    #[allow(
        dead_code,
        reason = "the payload is only read by feature-gated match arms"
    )]
    XlsFile(crate::xls::Workbook<std::io::BufReader<std::fs::File>>),
    #[cfg(feature = "xls")]
    XlsMem(crate::xls::Workbook<std::io::Cursor<Vec<u8>>>),

    /// Parsed ODS package retained for the dedicated facade to expose richer
    /// worksheet APIs as they become available.
    #[cfg(feature = "ods")]
    Ods(std::cell::RefCell<litchi_ods::Spreadsheet>),

    /// Filesystem-backed ODS with deferred package-member reads.
    #[cfg(all(feature = "ods", any(unix, windows)))]
    OdsSource(litchi_ods::SourceBackedSpreadsheet),

    /// Filesystem-backed XLS with deferred CFB/BIFF reads.
    #[cfg(all(feature = "xls", any(unix, windows)))]
    XlsSource(super::workbook::XlsSource),

    // For other formats, we just indicate they're not yet fully unified
    #[cfg(any(feature = "xls", any(feature = "xlsx", feature = "xlsb")))]
    #[allow(dead_code, reason = "placeholder for formats not yet fully unified")]
    Other,
}

/// Format of the workbook file.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[allow(
    unused,
    reason = "the enum is only constructed when the corresponding format feature is enabled"
)]
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
#[allow(
    dead_code,
    reason = "kept as a public detection helper; callers are expected to prefer smart detection"
)]
pub fn detect_workbook_format_from_signature<R: Read + Seek>(
    reader: &mut R,
) -> Result<WorkbookFormat> {
    const OLE2_SIGNATURE: [u8; 8] = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
    const ZIP_SIGNATURE: [u8; 4] = [b'P', b'K', 0x03, 0x04];

    let original = reader.stream_position()?;
    let detected = (|| {
        reader.seek(SeekFrom::Start(0))?;

        let mut prefix = [0u8; ZIP_SIGNATURE.len()];
        reader.read_exact(&mut prefix)?;

        // Check for ZIP-based workbooks (XLSX, XLSB, or Numbers). The
        // complete four-byte local-file signature is sufficient here; format
        // refinement performs package validation when it is requested.
        if prefix == ZIP_SIGNATURE {
            return Ok(WorkbookFormat::Xlsx);
        }

        // OLE2 has an eight-byte signature. Do not classify a truncated
        // prefix as XLS, even when the first four bytes happen to match.
        if prefix == OLE2_SIGNATURE[..prefix.len()] {
            let mut suffix = [0u8; OLE2_SIGNATURE.len() - ZIP_SIGNATURE.len()];
            reader.read_exact(&mut suffix)?;
            if suffix == OLE2_SIGNATURE[prefix.len()..] {
                return Ok(WorkbookFormat::Xls);
            }
        }

        Err(Error::NotOfficeFile)
    })();

    let restored = reader.seek(SeekFrom::Start(original));
    match (detected, restored) {
        (Ok(format), Ok(_)) => Ok(format),
        (Err(error), Ok(_)) => Err(error),
        (Ok(_), Err(error)) | (Err(_), Err(error)) => Err(error.into()),
    }
}

/// Refine ZIP-based workbook format detection (XLSX vs XLSB vs Numbers)
#[allow(
    dead_code,
    reason = "kept as a detection helper; callers are expected to prefer smart detection"
)]
#[cfg(any(
    feature = "numbers",
    any(feature = "xlsx", feature = "xlsb"),
    feature = "ods"
))]
pub fn refine_workbook_format<R: Read + Seek>(
    reader: &mut R,
    initial_format: WorkbookFormat,
) -> Result<WorkbookFormat> {
    // Only refine if it's a ZIP-based format
    if initial_format != WorkbookFormat::Xlsx {
        return Ok(initial_format);
    }

    let original = reader.stream_position()?;
    let refined = (|| {
        reader.seek(SeekFrom::Start(0))?;
        let mut data = Vec::new();
        reader.read_to_end(&mut data)?;

        #[cfg(any(feature = "xlsx", feature = "xlsb"))]
        if crate::detection_smart::ooxml::detect_zip_format(&data) == Some(FileFormat::Xlsb) {
            return Ok(WorkbookFormat::Xlsb);
        }

        #[cfg(feature = "ods")]
        if litchi_odf_common::detect::bytes(&data) == Some(FileFormat::Ods) {
            return Ok(WorkbookFormat::Ods);
        }

        #[cfg(feature = "numbers")]
        if litchi_iwa_detect::bytes(&data).ok().flatten()
            == Some(litchi_iwa_detect::Format::Numbers)
        {
            return Ok(WorkbookFormat::Numbers);
        }

        Ok(initial_format)
    })();
    reader.seek(SeekFrom::Start(original))?;
    refined
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Cursor;

    #[test]
    fn test_workbook_format_variants() {
        assert_eq!(WorkbookFormat::Xls, WorkbookFormat::Xls);
        assert_eq!(WorkbookFormat::Xlsx, WorkbookFormat::Xlsx);
        assert_eq!(WorkbookFormat::Xlsb, WorkbookFormat::Xlsb);
        assert_eq!(WorkbookFormat::Ods, WorkbookFormat::Ods);
        assert_eq!(WorkbookFormat::Numbers, WorkbookFormat::Numbers);
    }

    #[test]
    fn test_workbook_format_inequality() {
        assert_ne!(WorkbookFormat::Xls, WorkbookFormat::Xlsx);
        assert_ne!(WorkbookFormat::Xlsx, WorkbookFormat::Xlsb);
        assert_ne!(WorkbookFormat::Ods, WorkbookFormat::Numbers);
    }

    #[test]
    fn test_detect_workbook_format_from_signature_xls() {
        // OLE2 signature for XLS files
        let data = b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1";
        let mut cursor = Cursor::new(data);
        let result = detect_workbook_format_from_signature(&mut cursor);
        assert!(result.is_ok());
        assert_eq!(result.unwrap(), WorkbookFormat::Xls);
    }

    #[test]
    fn test_detect_workbook_format_from_signature_xlsx() {
        // The complete ZIP local-file signature is enough for the initial
        // classification; package refinement validates the actual workbook.
        let data = b"PK\x03\x04";
        let mut cursor = Cursor::new(data);
        cursor.set_position(2);
        let result = detect_workbook_format_from_signature(&mut cursor);
        assert!(result.is_ok());
        // Returns Xlsx by default for ZIP files
        assert_eq!(result.unwrap(), WorkbookFormat::Xlsx);
        assert_eq!(cursor.position(), 2);
    }

    #[test]
    fn test_detect_workbook_format_rejects_truncated_ole2_signature() {
        let mut cursor = Cursor::new(b"\xD0\xCF\x11\xE0".to_vec());
        let result = detect_workbook_format_from_signature(&mut cursor);
        assert!(matches!(result, Err(Error::Io(_))));
        assert_eq!(cursor.position(), 0);
    }

    #[test]
    fn test_detect_workbook_format_from_signature_invalid() {
        // Invalid signature
        let data = b"NOTVALID";
        let mut cursor = Cursor::new(data);
        let result = detect_workbook_format_from_signature(&mut cursor);
        assert!(result.is_err());
    }

    #[test]
    fn test_detect_workbook_format_from_signature_empty() {
        // Empty data
        let data = b"";
        let mut cursor = Cursor::new(data);
        let result = detect_workbook_format_from_signature(&mut cursor);
        assert!(result.is_err());
    }

    #[test]
    fn test_workbook_format_debug() {
        let format = WorkbookFormat::Xlsx;
        let debug_str = format!("{:?}", format);
        assert!(debug_str.contains("Xlsx"));
    }

    #[test]
    fn test_workbook_format_clone() {
        let format = WorkbookFormat::Xls;
        let cloned = format;
        assert_eq!(format, cloned);
    }

    #[test]
    fn test_workbook_format_copy() {
        let format = WorkbookFormat::Ods;
        let copied = format;
        assert_eq!(format, copied);
    }

    #[test]
    #[cfg(any(
        feature = "numbers",
        any(feature = "xlsx", feature = "xlsb"),
        feature = "ods"
    ))]
    fn refinement_restores_a_nonzero_cursor() {
        let mut reader = Cursor::new(b"PK\x03\x04not-a-valid-package".as_slice());
        reader.set_position(5);
        assert_eq!(
            refine_workbook_format(&mut reader, WorkbookFormat::Xlsx).unwrap(),
            WorkbookFormat::Xlsx
        );
        assert_eq!(reader.position(), 5);
    }
}
