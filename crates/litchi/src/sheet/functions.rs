//! Functions for opening workbooks.

use super::traits::WorkbookTrait;
use super::types::Result;

/// Open a workbook from a file path.
///
/// On Unix and Windows, validates the XLSX OPC catalog and relationship graph
/// at open while deferring selected worksheet payload extraction and parsing
/// until worksheet reads. Other targets retain the portable eager fallback.
/// Uses the default XLSX OPC resource policy. Use
/// [`open_workbook_with_limits`] when the input is untrusted.
#[cfg(feature = "xlsx")]
pub fn open_workbook<P: AsRef<std::path::Path>>(path: P) -> Result<Box<dyn WorkbookTrait>> {
    open_workbook_with_limits(path, crate::xlsx::ReadLimits::default())
}

/// Open an XLSX workbook from a file path with an explicit OPC resource policy.
///
/// On Unix and Windows, catalog validation happens at open and selected
/// worksheet payloads are extracted and parsed on demand. Other targets use
/// the portable eager fallback.
#[cfg(feature = "xlsx")]
pub fn open_workbook_with_limits<P: AsRef<std::path::Path>>(
    path: P,
    limits: crate::xlsx::ReadLimits,
) -> Result<Box<dyn WorkbookTrait>> {
    // `SourceBackedWorkbook` is available on regular filesystem targets. Keep
    // the historical eager path as the portable fallback for targets without
    // a native filesystem-backed `FileSource`.
    #[cfg(any(unix, windows))]
    {
        let workbook = crate::xlsx::SourceBackedWorkbook::from_path_with_limits(path, limits)
            .map_err(crate::map_ooxml_error)?;
        Ok(Box::new(super::adapters::Workbook::from_source_backed(
            workbook,
        )))
    }
    #[cfg(not(any(unix, windows)))]
    {
        let workbook = crate::xlsx::Package::open_with_limits(path, limits)
            .map_err(crate::map_ooxml_error)?
            .into_workbook()
            .map_err(crate::map_ooxml_error)?;
        Ok(Box::new(super::adapters::Workbook::new(workbook)))
    }
}

/// Open a workbook from bytes.
///
/// Uses the default XLSX OPC resource policy. Use
/// [`open_workbook_from_bytes_with_limits`] when the input is untrusted.
#[cfg(feature = "xlsx")]
pub fn open_workbook_from_bytes(bytes: &[u8]) -> Result<Box<dyn WorkbookTrait>> {
    open_workbook_from_bytes_with_limits(bytes, crate::xlsx::ReadLimits::default())
}

/// Open XLSX workbook bytes with an explicit OPC resource policy.
#[cfg(feature = "xlsx")]
pub fn open_workbook_from_bytes_with_limits(
    bytes: &[u8],
    limits: crate::xlsx::ReadLimits,
) -> Result<Box<dyn WorkbookTrait>> {
    let workbook = crate::xlsx::Package::from_slice_with_limits(bytes, limits)
        .map_err(crate::map_ooxml_error)?
        .into_workbook()
        .map_err(crate::map_ooxml_error)?;
    Ok(Box::new(super::adapters::Workbook::new(workbook)))
}

/// Open an XLS workbook from a file path.
///
/// **Note**: This requires the `xls` feature to be enabled.
#[cfg(feature = "xls")]
pub fn open_xls_workbook<P: AsRef<std::path::Path>>(
    path: P,
) -> Result<crate::xls::Workbook<std::fs::File>> {
    use std::fs::File;
    let file = File::open(path)?;
    let workbook = crate::xls::Workbook::new(file)?;
    Ok(workbook)
}

/// Open an XLS workbook from bytes.
///
/// **Note**: This requires the `xls` feature to be enabled.
#[cfg(feature = "xls")]
pub fn open_xls_workbook_from_bytes(
    bytes: &[u8],
) -> Result<crate::xls::Workbook<std::io::Cursor<&[u8]>>> {
    use std::io::Cursor;
    let cursor = Cursor::new(bytes);
    let workbook = crate::xls::Workbook::new(cursor)?;
    Ok(workbook)
}

/// Open an XLS workbook as a trait object from a file path.
///
/// **Note**: This requires the `xls` feature to be enabled.
#[cfg(feature = "xls")]
pub fn open_xls_workbook_dyn<P: AsRef<std::path::Path>>(path: P) -> Result<Box<dyn WorkbookTrait>> {
    let workbook = open_xls_workbook(path)?;
    Ok(Box::new(workbook))
}

/// Open an XLS workbook as a trait object from bytes.
///
/// **Note**: This requires the `xls` feature to be enabled.
#[cfg(feature = "xls")]
pub fn open_xls_workbook_from_bytes_dyn(bytes: &[u8]) -> Result<Box<dyn WorkbookTrait>> {
    use std::io::Cursor;
    let cursor = Cursor::new(bytes.to_vec());
    let workbook = crate::xls::Workbook::new(cursor)?;
    Ok(Box::new(workbook))
}

/// Open an XLSB workbook from a file path.
///
/// Uses the default XLSB OPC resource policy. Use
/// [`open_xlsb_workbook_with_limits`] when the input is untrusted.
#[cfg(feature = "xlsb")]
pub fn open_xlsb_workbook<P: AsRef<std::path::Path>>(path: P) -> Result<crate::xlsb::Workbook> {
    open_xlsb_workbook_with_limits(path, crate::xlsb::ReadLimits::default())
}

/// Open an XLSB workbook from a file path with an explicit OPC resource policy.
#[cfg(feature = "xlsb")]
pub fn open_xlsb_workbook_with_limits<P: AsRef<std::path::Path>>(
    path: P,
    limits: crate::xlsb::ReadLimits,
) -> Result<crate::xlsb::Workbook> {
    let workbook = crate::xlsb::Package::open_with_limits(path, limits)
        .map_err(crate::map_ooxml_error)?
        .into_workbook()
        .map_err(crate::map_ooxml_error)?;
    Ok(workbook)
}

/// Open an XLSB workbook from bytes.
///
/// Uses the default XLSB OPC resource policy. Use
/// [`open_xlsb_workbook_from_bytes_with_limits`] when the input is untrusted.
#[cfg(feature = "xlsb")]
pub fn open_xlsb_workbook_from_bytes(bytes: &[u8]) -> Result<crate::xlsb::Workbook> {
    open_xlsb_workbook_from_bytes_with_limits(bytes, crate::xlsb::ReadLimits::default())
}

/// Open XLSB workbook bytes with an explicit OPC resource policy.
#[cfg(feature = "xlsb")]
pub fn open_xlsb_workbook_from_bytes_with_limits(
    bytes: &[u8],
    limits: crate::xlsb::ReadLimits,
) -> Result<crate::xlsb::Workbook> {
    let workbook = crate::xlsb::Package::from_slice_with_limits(bytes, limits)
        .map_err(crate::map_ooxml_error)?
        .into_workbook()
        .map_err(crate::map_ooxml_error)?;
    Ok(workbook)
}

/// Open an XLSB workbook as a trait object from a file path.
///
/// Uses the default XLSB OPC resource policy. Use
/// [`open_xlsb_workbook_dyn_with_limits`] when the input is untrusted.
#[cfg(feature = "xlsb")]
pub fn open_xlsb_workbook_dyn<P: AsRef<std::path::Path>>(
    path: P,
) -> Result<Box<dyn WorkbookTrait>> {
    open_xlsb_workbook_dyn_with_limits(path, crate::xlsb::ReadLimits::default())
}

/// Open an XLSB workbook as a trait object with an explicit OPC resource policy.
#[cfg(feature = "xlsb")]
pub fn open_xlsb_workbook_dyn_with_limits<P: AsRef<std::path::Path>>(
    path: P,
    limits: crate::xlsb::ReadLimits,
) -> Result<Box<dyn WorkbookTrait>> {
    #[cfg(all(
        any(feature = "xlsx", feature = "ods", feature = "xls", feature = "xlsb"),
        any(unix, windows)
    ))]
    {
        if let crate::detection_smart::detected::WorkbookSourcePathDetection::Xlsb {
            workbook,
            source,
            limits,
            ..
        } = crate::detection_smart::detected::detect_workbook_source_path_with_limits(
            path.as_ref(),
            limits,
        )? {
            let workbook =
                super::adapters::XlsbWorkbook::from_source_backed(workbook, source, limits)?;
            return Ok(Box::new(workbook));
        }
    }

    let workbook = open_xlsb_workbook_with_limits(path, limits)?;
    Ok(Box::new(workbook))
}

/// Open an XLSB workbook as a trait object from bytes.
///
/// Uses the default XLSB OPC resource policy. Use
/// [`open_xlsb_workbook_from_bytes_dyn_with_limits`] when the input is
/// untrusted.
#[cfg(feature = "xlsb")]
pub fn open_xlsb_workbook_from_bytes_dyn(bytes: &[u8]) -> Result<Box<dyn WorkbookTrait>> {
    open_xlsb_workbook_from_bytes_dyn_with_limits(bytes, crate::xlsb::ReadLimits::default())
}

/// Open XLSB workbook bytes as a trait object with an explicit OPC resource
/// policy.
#[cfg(feature = "xlsb")]
pub fn open_xlsb_workbook_from_bytes_dyn_with_limits(
    bytes: &[u8],
    limits: crate::xlsb::ReadLimits,
) -> Result<Box<dyn WorkbookTrait>> {
    let workbook = open_xlsb_workbook_from_bytes_with_limits(bytes, limits)?;
    Ok(Box::new(workbook))
}

/// Open a CSV workbook from a file path.
pub fn open_csv_workbook<P: AsRef<std::path::Path>>(path: P) -> Result<Box<dyn WorkbookTrait>> {
    let workbook = crate::sheet::text::TextWorkbook::open(path)?;
    Ok(Box::new(workbook))
}

/// Open a CSV workbook from bytes.
pub fn open_csv_workbook_from_bytes(bytes: &[u8]) -> Result<Box<dyn WorkbookTrait>> {
    let workbook = crate::sheet::text::TextWorkbook::from_bytes(
        bytes,
        crate::sheet::text::TextConfig::default(),
    )?;
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
    config: crate::sheet::text::TextConfig,
) -> Result<Box<dyn WorkbookTrait>> {
    let workbook = crate::sheet::text::TextWorkbook::from_path_with_config(path, config)?;
    Ok(Box::new(workbook))
}

/// Open a text workbook with custom configuration from bytes.
pub fn open_text_workbook_from_bytes_with_config(
    bytes: &[u8],
    config: crate::sheet::text::TextConfig,
) -> Result<Box<dyn WorkbookTrait>> {
    let workbook = crate::sheet::text::TextWorkbook::from_bytes(bytes, config)?;
    Ok(Box::new(workbook))
}

/// Open a SYLK (Symbolic Link) workbook from a file path.
pub fn open_sylk_workbook<P: AsRef<std::path::Path>>(path: P) -> Result<Box<dyn WorkbookTrait>> {
    use std::fs::File;
    let mut file = File::open(path)?;
    let data = crate::sheet::text::formats::read_sylk(&mut file, Default::default())?;
    let mut workbook = crate::sheet::text::TextWorkbook::from_bytes(
        &[],
        crate::sheet::text::TextConfig::default(),
    )?;
    workbook.set_data(data);
    Ok(Box::new(workbook))
}

/// Open a SYLK workbook from bytes.
pub fn open_sylk_workbook_from_bytes(bytes: &[u8]) -> Result<Box<dyn WorkbookTrait>> {
    let mut cursor = std::io::Cursor::new(bytes);
    let data = crate::sheet::text::formats::read_sylk(&mut cursor, Default::default())?;
    let mut workbook = crate::sheet::text::TextWorkbook::from_bytes(
        &[],
        crate::sheet::text::TextConfig::default(),
    )?;
    workbook.set_data(data);
    Ok(Box::new(workbook))
}

/// Open a DIF (Data Interchange Format) workbook from a file path.
pub fn open_dif_workbook<P: AsRef<std::path::Path>>(path: P) -> Result<Box<dyn WorkbookTrait>> {
    use std::fs::File;
    let mut file = File::open(path)?;
    let data = crate::sheet::text::formats::read_dif(&mut file, Default::default())?;
    let mut workbook = crate::sheet::text::TextWorkbook::from_bytes(
        &[],
        crate::sheet::text::TextConfig::default(),
    )?;
    workbook.set_data(data);
    Ok(Box::new(workbook))
}

/// Open a DIF workbook from bytes.
pub fn open_dif_workbook_from_bytes(bytes: &[u8]) -> Result<Box<dyn WorkbookTrait>> {
    let mut cursor = std::io::Cursor::new(bytes);
    let data = crate::sheet::text::formats::read_dif(&mut cursor, Default::default())?;
    let mut workbook = crate::sheet::text::TextWorkbook::from_bytes(
        &[],
        crate::sheet::text::TextConfig::default(),
    )?;
    workbook.set_data(data);
    Ok(Box::new(workbook))
}

/// Open a fixed-width PRN workbook from a file path.
pub fn open_fixed_width_workbook<P: AsRef<std::path::Path>>(
    path: P,
) -> Result<Box<dyn WorkbookTrait>> {
    use std::fs::File;
    let mut file = File::open(path)?;
    let data = crate::sheet::text::formats::read_fixed_width(&mut file, Default::default())?;
    let mut workbook = crate::sheet::text::TextWorkbook::from_bytes(
        &[],
        crate::sheet::text::TextConfig::default(),
    )?;
    workbook.set_data(data);
    Ok(Box::new(workbook))
}

/// Open a fixed-width PRN workbook from bytes.
pub fn open_fixed_width_workbook_from_bytes(bytes: &[u8]) -> Result<Box<dyn WorkbookTrait>> {
    let mut cursor = std::io::Cursor::new(bytes);
    let data = crate::sheet::text::formats::read_fixed_width(&mut cursor, Default::default())?;
    let mut workbook = crate::sheet::text::TextWorkbook::from_bytes(
        &[],
        crate::sheet::text::TextConfig::default(),
    )?;
    workbook.set_data(data);
    Ok(Box::new(workbook))
}
