//! Byte Order Mark (BOM) utilities shared across modules.
//!
//! Provides detection, stripping, and writing helpers for common Unicode
//! encodings used in text-based formats.

use crate::common::Result;
use std::io::{Read, Seek, SeekFrom, Write};

/// Supported BOM encodings.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BomKind {
    Utf8,
    Utf16Le,
    Utf16Be,
    Utf32Le,
    Utf32Be,
}

impl BomKind {
    /// Returns the byte representation of the BOM.
    #[inline]
    pub const fn as_bytes(&self) -> &'static [u8] {
        match self {
            BomKind::Utf8 => &UTF8_BOM,
            BomKind::Utf16Le => &UTF16_LE_BOM,
            BomKind::Utf16Be => &UTF16_BE_BOM,
            BomKind::Utf32Le => &UTF32_LE_BOM,
            BomKind::Utf32Be => &UTF32_BE_BOM,
        }
    }

    /// Returns the length in bytes of the BOM.
    #[inline]
    #[allow(clippy::len_without_is_empty)] // No need to check for empty BOMs
    pub const fn len(&self) -> usize {
        self.as_bytes().len()
    }
}

/// UTF-8 BOM bytes.
pub const UTF8_BOM: [u8; 3] = [0xEF, 0xBB, 0xBF];
/// UTF-16 little-endian BOM bytes.
pub const UTF16_LE_BOM: [u8; 2] = [0xFF, 0xFE];
/// UTF-16 big-endian BOM bytes.
pub const UTF16_BE_BOM: [u8; 2] = [0xFE, 0xFF];
/// UTF-32 little-endian BOM bytes.
pub const UTF32_LE_BOM: [u8; 4] = [0xFF, 0xFE, 0x00, 0x00];
/// UTF-32 big-endian BOM bytes.
pub const UTF32_BE_BOM: [u8; 4] = [0x00, 0x00, 0xFE, 0xFF];

/// Detects and consumes a BOM if present.
///
/// Returns the detected BOM kind and leaves the reader positioned after the
/// BOM. When no BOM is found, rewinds the reader to the original position and
/// returns `Ok(None)`.
pub fn strip_bom<R: Read + Seek>(reader: &mut R) -> Result<Option<(BomKind, usize)>> {
    let start = reader.stream_position()?;
    let mut buf = [0u8; 4];
    let mut read = 0usize;

    while read < buf.len() {
        match reader.read(&mut buf[read..])? {
            0 => break,
            n => read += n,
        }
    }

    if let Some((kind, len)) = detect_bom(&buf, read) {
        reader.seek(SeekFrom::Start(start + len as u64))?;
        return Ok(Some((kind, len)));
    }

    reader.seek(SeekFrom::Start(start))?;
    Ok(None)
}

/// Writes the requested BOM to the writer.
pub fn write_bom<W: Write>(writer: &mut W, kind: BomKind) -> Result<()> {
    writer.write_all(kind.as_bytes())?;
    Ok(())
}

fn detect_bom(buf: &[u8; 4], read: usize) -> Option<(BomKind, usize)> {
    if read >= UTF32_BE_BOM.len() {
        if buf[..UTF32_BE_BOM.len()] == UTF32_BE_BOM {
            return Some((BomKind::Utf32Be, UTF32_BE_BOM.len()));
        }
        if buf[..UTF32_LE_BOM.len()] == UTF32_LE_BOM {
            return Some((BomKind::Utf32Le, UTF32_LE_BOM.len()));
        }
    }

    if read >= UTF8_BOM.len() && buf[..UTF8_BOM.len()] == UTF8_BOM {
        return Some((BomKind::Utf8, UTF8_BOM.len()));
    }

    if read >= UTF16_BE_BOM.len() && buf[..UTF16_BE_BOM.len()] == UTF16_BE_BOM {
        return Some((BomKind::Utf16Be, UTF16_BE_BOM.len()));
    }
    if read >= UTF16_LE_BOM.len() && buf[..UTF16_LE_BOM.len()] == UTF16_LE_BOM {
        return Some((BomKind::Utf16Le, UTF16_LE_BOM.len()));
    }

    None
}
