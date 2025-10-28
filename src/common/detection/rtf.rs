//! RTF (Rich Text Format) detection.
//!
//! RTF files have a simple text-based signature that makes them easy to detect.
//! The signature is `{\rtf` or `{\rtf1` at the start of the file.

use crate::common::detection::FileFormat;
use std::io::{Read, Seek};

/// RTF signature patterns.
/// RTF files start with `{\rtf` followed optionally by version number.
const RTF_SIGNATURE_MIN: &[u8] = b"{\\rtf";
const RTF_SIGNATURE_LEN: usize = 5;

/// Detect RTF format from byte content.
///
/// RTF files are plain text files that start with `{\rtf` or `{\rtf1`.
/// This function checks the first few bytes for this signature.
///
/// # Arguments
///
/// * `bytes` - The file data as bytes
///
/// # Returns
///
/// * `Some(FileFormat::Rtf)` if the file is RTF format
/// * `None` if not RTF format
///
/// # Examples
///
/// ```rust
/// use litchi::common::detection::rtf::detect_rtf_format;
///
/// let rtf_data = b"{\\rtf1\\ansi\\deff0 Hello World}";
/// assert!(detect_rtf_format(rtf_data).is_some());
///
/// let non_rtf_data = b"Plain text file";
/// assert!(detect_rtf_format(non_rtf_data).is_none());
/// ```
#[inline]
pub fn detect_rtf_format(bytes: &[u8]) -> Option<FileFormat> {
    if bytes.len() < RTF_SIGNATURE_LEN {
        return None;
    }

    // Check for RTF signature
    if &bytes[0..RTF_SIGNATURE_LEN] == RTF_SIGNATURE_MIN {
        return Some(FileFormat::Rtf);
    }

    None
}

/// Detect RTF format from a reader.
///
/// This function reads the first few bytes from the reader to check
/// for the RTF signature, then resets the reader to the start.
///
/// # Arguments
///
/// * `reader` - A reader that implements Read + Seek
///
/// # Returns
///
/// * `Some(FileFormat::Rtf)` if the file is RTF format
/// * `None` if not RTF format or read error
pub fn detect_rtf_format_from_reader<R: Read + Seek>(reader: &mut R) -> Option<FileFormat> {
    use std::io::SeekFrom;

    // Read first bytes for signature check
    let mut buffer = [0u8; RTF_SIGNATURE_LEN];
    if reader.read_exact(&mut buffer).is_err() {
        return None;
    }

    // Reset to beginning
    let _ = reader.seek(SeekFrom::Start(0));

    detect_rtf_format(&buffer)
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Cursor;

    #[test]
    fn test_detect_rtf_valid() {
        let rtf_data = b"{\\rtf1\\ansi\\deff0 {\\fonttbl {\\f0 Times New Roman;}} Hello World}";
        assert_eq!(detect_rtf_format(rtf_data), Some(FileFormat::Rtf));
    }

    #[test]
    fn test_detect_rtf_minimal() {
        let rtf_data = b"{\\rtf}";
        assert_eq!(detect_rtf_format(rtf_data), Some(FileFormat::Rtf));
    }

    #[test]
    fn test_detect_rtf_invalid() {
        let non_rtf_data = b"Plain text file";
        assert_eq!(detect_rtf_format(non_rtf_data), None);
    }

    #[test]
    fn test_detect_rtf_too_short() {
        let short_data = b"{\\rt";
        assert_eq!(detect_rtf_format(short_data), None);
    }

    #[test]
    fn test_detect_rtf_from_reader() {
        let rtf_data = b"{\\rtf1\\ansi Hello World}";
        let mut cursor = Cursor::new(rtf_data);
        assert_eq!(
            detect_rtf_format_from_reader(&mut cursor),
            Some(FileFormat::Rtf)
        );

        // Verify reader was reset
        let mut buffer = Vec::new();
        cursor.read_to_end(&mut buffer).unwrap();
        assert_eq!(&buffer[..], rtf_data);
    }
}
