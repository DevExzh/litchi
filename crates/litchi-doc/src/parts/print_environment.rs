//! Bounded printer metadata from the DOC `PrDrvr`, `PrEnvPort`, and
//! `PrEnvLand` table-stream records (MS-DOC 2.9.211–213).
//!
//! `PrDrvr` contains four ANSI, NUL-terminated names. The portrait and
//! landscape environment records are printer-provided binary blocks that the
//! specification marks unused and ignored, so they remain inert exact blobs.
//! No printer is contacted and no print settings are applied.

use std::ops::Range;

use super::super::package::{Error as PackageError, Result};
use super::fib::FileInformationBlock;

/// FIB index for `PrDrvr`.
pub const FIB_INDEX_PR_DRVR: usize = 27;
/// FIB index for the portrait `PrEnvPort` blob.
pub const FIB_INDEX_PR_ENV_PORT: usize = 28;
/// FIB index for the landscape `PrEnvLand` blob.
pub const FIB_INDEX_PR_ENV_LAND: usize = 29;
/// Maximum size materialized for any printer metadata table.
pub const MAX_PRINT_METADATA_BYTES: usize = 1024 * 1024;

fn corrupted(message: impl Into<String>) -> PackageError {
    PackageError::Corrupted(message.into())
}

fn copy_source(data: &[u8]) -> Result<Vec<u8>> {
    let mut source = Vec::new();
    source
        .try_reserve_exact(data.len())
        .map_err(|_| corrupted("print metadata allocation failed"))?;
    source.extend_from_slice(data);
    Ok(source)
}

/// Four ANSI printer strings from `PrDrvr`, retaining their exact bytes.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PrintDriver {
    source: Vec<u8>,
    printer: Range<usize>,
    port: Range<usize>,
    driver: Range<usize>,
    product: Range<usize>,
    trailing: Range<usize>,
}

impl PrintDriver {
    /// Parse one complete `PrDrvr` payload.
    pub fn parse_bytes(data: &[u8]) -> Result<Self> {
        if data.len() > MAX_PRINT_METADATA_BYTES {
            return Err(corrupted(format!(
                "PrDrvr exceeds the {MAX_PRINT_METADATA_BYTES}-byte limit"
            )));
        }
        let mut cursor = 0;
        let printer = read_string(data, &mut cursor, "szPrinter")?;
        let port = read_string(data, &mut cursor, "szPrPort")?;
        let driver = read_string(data, &mut cursor, "szPrDriver")?;
        let product = read_string(data, &mut cursor, "szTruePrnName")?;
        let trailing = cursor..data.len();
        Ok(Self {
            source: copy_source(data)?,
            printer,
            port,
            driver,
            product,
            trailing,
        })
    }

    /// Exact serialized bytes, including trailing producer data.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.source
    }

    /// Printer name bytes without the terminating NUL.
    #[must_use]
    pub fn printer(&self) -> &[u8] {
        &self.source[self.printer.clone()]
    }

    /// Printer port bytes without the terminating NUL.
    #[must_use]
    pub fn port(&self) -> &[u8] {
        &self.source[self.port.clone()]
    }

    /// Printer driver bytes without the terminating NUL.
    #[must_use]
    pub fn driver(&self) -> &[u8] {
        &self.source[self.driver.clone()]
    }

    /// Printer product name bytes without the terminating NUL.
    #[must_use]
    pub fn product(&self) -> &[u8] {
        &self.source[self.product.clone()]
    }

    /// Producer-specific bytes after the four required strings.
    #[must_use]
    pub fn trailing(&self) -> &[u8] {
        &self.source[self.trailing.clone()]
    }
}

/// An inert printer-provided portrait or landscape environment block.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PrintEnvironment {
    source: Vec<u8>,
}

impl PrintEnvironment {
    /// Parse one bounded printer environment blob.
    pub fn parse_bytes(data: &[u8]) -> Result<Self> {
        if data.len() > MAX_PRINT_METADATA_BYTES {
            return Err(corrupted(format!(
                "print environment exceeds the {MAX_PRINT_METADATA_BYTES}-byte limit"
            )));
        }
        Ok(Self {
            source: copy_source(data)?,
        })
    }

    /// Exact printer-provided bytes. The blob is never interpreted.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.source
    }

    /// Whether the printer supplied no bytes.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.source.is_empty()
    }
}

/// The optional printer metadata tables referenced by the main FIB.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DocumentPrintEnvironment {
    driver: Option<PrintDriver>,
    portrait: Option<PrintEnvironment>,
    landscape: Option<PrintEnvironment>,
}

impl DocumentPrintEnvironment {
    pub(crate) fn from_parts(
        driver: Option<PrintDriver>,
        portrait: Option<PrintEnvironment>,
        landscape: Option<PrintEnvironment>,
    ) -> Option<Self> {
        if driver.is_none() && portrait.is_none() && landscape.is_none() {
            return None;
        }
        Some(Self {
            driver,
            portrait,
            landscape,
        })
    }

    /// Parse all three optional printer metadata records.
    pub fn parse(fib: &FileInformationBlock, table_stream: &[u8]) -> Result<Option<Self>> {
        let driver = parse_driver(fib, table_stream)?;
        let portrait = parse_blob(fib, table_stream, FIB_INDEX_PR_ENV_PORT, "PrEnvPort")?;
        let landscape = parse_blob(fib, table_stream, FIB_INDEX_PR_ENV_LAND, "PrEnvLand")?;
        Ok(Self::from_parts(driver, portrait, landscape))
    }

    /// Optional `PrDrvr` metadata.
    #[must_use]
    pub fn driver(&self) -> Option<&PrintDriver> {
        self.driver.as_ref()
    }

    /// Optional portrait `PrEnvPort` bytes.
    #[must_use]
    pub fn portrait(&self) -> Option<&PrintEnvironment> {
        self.portrait.as_ref()
    }

    /// Optional landscape `PrEnvLand` bytes.
    #[must_use]
    pub fn landscape(&self) -> Option<&PrintEnvironment> {
        self.landscape.as_ref()
    }
}

fn read_string(data: &[u8], cursor: &mut usize, field: &str) -> Result<Range<usize>> {
    let remaining = data
        .get(*cursor..)
        .ok_or_else(|| corrupted(format!("PrDrvr {field} offset is invalid")))?;
    let Some(end) = remaining.iter().position(|byte| *byte == 0) else {
        return Err(corrupted(format!("PrDrvr {field} is not NUL-terminated")));
    };
    let end = (*cursor)
        .checked_add(end)
        .ok_or_else(|| corrupted(format!("PrDrvr {field} range overflows")))?;
    let value = *cursor..end;
    *cursor = end
        .checked_add(1)
        .ok_or_else(|| corrupted(format!("PrDrvr {field} range overflows")))?;
    Ok(value)
}

fn parse_driver(fib: &FileInformationBlock, table_stream: &[u8]) -> Result<Option<PrintDriver>> {
    let Some(data) = table_range(fib, table_stream, FIB_INDEX_PR_DRVR, "PrDrvr")? else {
        return Ok(None);
    };
    PrintDriver::parse_bytes(data).map(Some)
}

fn parse_blob(
    fib: &FileInformationBlock,
    table_stream: &[u8],
    index: usize,
    field: &str,
) -> Result<Option<PrintEnvironment>> {
    let Some(data) = table_range(fib, table_stream, index, field)? else {
        return Ok(None);
    };
    PrintEnvironment::parse_bytes(data).map(Some)
}

fn table_range<'a>(
    fib: &FileInformationBlock,
    table_stream: &'a [u8],
    index: usize,
    field: &str,
) -> Result<Option<&'a [u8]>> {
    let Some((offset, length)) = fib.get_table_pointer(index) else {
        return Ok(None);
    };
    if length == 0 {
        return Ok(None);
    }
    let length = usize::try_from(length)
        .map_err(|_| corrupted(format!("{field} length does not fit in memory")))?;
    if length > MAX_PRINT_METADATA_BYTES {
        return Err(corrupted(format!(
            "{field} exceeds the {MAX_PRINT_METADATA_BYTES}-byte limit"
        )));
    }
    let start = usize::try_from(offset)
        .map_err(|_| corrupted(format!("{field} offset does not fit in memory")))?;
    let end = start
        .checked_add(length)
        .ok_or_else(|| corrupted(format!("{field} range overflows")))?;
    table_stream
        .get(start..end)
        .ok_or_else(|| corrupted(format!("{field} extends beyond the table stream")))
        .map(Some)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parses_driver_strings_and_retains_trailing_bytes() {
        let data = b"printer\0port\0driver\0product\0\xA5";
        let driver = PrintDriver::parse_bytes(data).unwrap();
        assert_eq!(driver.printer(), b"printer");
        assert_eq!(driver.port(), b"port");
        assert_eq!(driver.driver(), b"driver");
        assert_eq!(driver.product(), b"product");
        assert_eq!(driver.trailing(), &[0xA5]);
        assert_eq!(driver.bytes(), data);
    }

    #[test]
    fn rejects_unterminated_driver_string_and_bounds_blob() {
        for data in [
            b"printer".as_slice(),
            b"printer\0port",
            b"printer\0port\0driver",
            b"printer\0port\0driver\0product",
        ] {
            assert!(PrintDriver::parse_bytes(data).is_err());
        }
        let too_large = vec![0; MAX_PRINT_METADATA_BYTES + 1];
        assert!(PrintDriver::parse_bytes(&too_large).is_err());
        assert!(PrintEnvironment::parse_bytes(&too_large).is_err());
    }

    #[test]
    fn retains_empty_driver_strings_and_binary_tail() {
        let data = b"\0\0\0\0\xff\0\x80";
        let driver = PrintDriver::parse_bytes(data).unwrap();
        assert!(driver.printer().is_empty());
        assert!(driver.port().is_empty());
        assert!(driver.driver().is_empty());
        assert!(driver.product().is_empty());
        assert_eq!(driver.trailing(), b"\xff\0\x80");
        assert_eq!(driver.bytes(), data);
    }

    #[test]
    fn parses_driver_and_ignored_blobs_from_their_fib_ranges() {
        let driver = b"printer\0port\0driver\0product\0";
        let portrait = [0x11, 0x22];
        let landscape = [0x33, 0x44, 0x55];
        let driver_offset = 4u32;
        let portrait_offset = driver_offset + driver.len() as u32;
        let landscape_offset = portrait_offset + portrait.len() as u32;
        let highest = FIB_INDEX_PR_ENV_LAND;
        let mut fib_data = vec![0; 154 + (highest + 1) * 8];
        fib_data[0..2].copy_from_slice(&0xA5ECu16.to_le_bytes());
        fib_data[2..4].copy_from_slice(&0x00C1u16.to_le_bytes());
        fib_data[152..154].copy_from_slice(&((highest + 1) as u16).to_le_bytes());
        for (index, offset, length) in [
            (FIB_INDEX_PR_DRVR, driver_offset, driver.len()),
            (FIB_INDEX_PR_ENV_PORT, portrait_offset, portrait.len()),
            (FIB_INDEX_PR_ENV_LAND, landscape_offset, landscape.len()),
        ] {
            let pointer = 154 + index * 8;
            fib_data[pointer..pointer + 4].copy_from_slice(&offset.to_le_bytes());
            fib_data[pointer + 4..pointer + 8].copy_from_slice(&(length as u32).to_le_bytes());
        }
        let fib = FileInformationBlock::parse(&fib_data).unwrap();
        let mut table = vec![0xCC; landscape_offset as usize + landscape.len()];
        table[driver_offset as usize..portrait_offset as usize].copy_from_slice(driver);
        table[portrait_offset as usize..landscape_offset as usize].copy_from_slice(&portrait);
        table[landscape_offset as usize..].copy_from_slice(&landscape);

        let metadata = DocumentPrintEnvironment::parse(&fib, &table)
            .unwrap()
            .unwrap();
        assert_eq!(metadata.driver().unwrap().printer(), b"printer");
        assert_eq!(metadata.portrait().unwrap().bytes(), portrait);
        assert_eq!(metadata.landscape().unwrap().bytes(), landscape);
    }
}
