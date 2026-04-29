/// CurrentUser stream parser for PowerPoint presentations.
///
/// The CurrentUser stream contains information about the current editing session,
/// including the offset to the current user edit record. This follows Apache POI's
/// CurrentUserAtom implementation.
use super::package::{PptError, Result};
use zerocopy::FromBytes;
use zerocopy_derive::FromBytes as DeriveFromBytes;

/// Minimum size of CurrentUser stream in bytes
const CURRENT_USER_MIN_SIZE: usize = 28;

/// Current User information.
///
/// Based on Apache POI's CurrentUserAtom, this contains information about
/// the current editing session in a PowerPoint presentation.
#[derive(Debug, Clone)]
pub struct CurrentUser {
    /// Offset to the current UserEditAtom record
    current_edit_offset: u32,
    /// Release version
    release_version: u16,
    /// Username (UTF-16LE encoded)
    username: String,
    /// Relative path to the presentation
    rel_path: String,
}

/// CurrentUser header structure (first 16 bytes after size field)
#[derive(Debug, Clone, DeriveFromBytes)]
#[repr(C)]
struct CurrentUserHeader {
    /// Header token (must be 0xF3D1C4DF)
    header_token: u32,
    /// Offset to the current UserEditAtom record
    current_edit_offset: u32,
    /// Username length in characters (Unicode)
    username_len: u16,
    /// Release version
    release_version: u16,
    /// ANSI username length in characters
    ansi_username_len: u16,
    /// Padding to align to 16 bytes
    _padding: u16,
}

impl CurrentUser {
    /// Parse a CurrentUser stream from binary data.
    ///
    /// # Arguments
    ///
    /// * `data` - The CurrentUser stream data
    ///
    /// # Returns
    ///
    /// A parsed CurrentUser structure or an error if the data is invalid.
    ///
    /// # Format (based on Apache POI's CurrentUserAtom)
    ///
    /// - Bytes 0-3: Size (little-endian u32)
    /// - Bytes 4-7: Header token (0xF3D1C4DF)
    /// - Bytes 8-11: Current edit offset (u32)
    /// - Bytes 12-13: Username length (u16)
    /// - Bytes 14-15: Release version (u16)
    /// - Bytes 16-19: Major/minor version
    /// - Bytes 20+: Username (UTF-16LE) and relative path
    pub fn parse(data: &[u8]) -> Result<Self> {
        if data.len() < CURRENT_USER_MIN_SIZE {
            return Err(PptError::Corrupted(
                "CurrentUser stream too short".to_string(),
            ));
        }

        // Parse header using zerocopy (bytes 4-19)
        let header = CurrentUserHeader::read_from_bytes(&data[4..20])
            .map_err(|_| PptError::Corrupted("Invalid CurrentUser header format".to_string()))?;

        // Validate header token (magic number)
        if header.header_token != 0xF3D1C4DF {
            return Err(PptError::InvalidFormat(format!(
                "Invalid CurrentUser header token: 0x{:08X}",
                header.header_token
            )));
        }

        let current_edit_offset = header.current_edit_offset;
        let username_len = header.username_len;
        let release_version = header.release_version;

        // Parse username (UTF-16LE encoded)
        let username = if username_len > 0 && data.len() >= 20 {
            let username_byte_len = (username_len as usize) * 2; // UTF-16LE = 2 bytes per char
            let username_start = 20;
            let username_end = username_start + username_byte_len;

            if username_end <= data.len() {
                Self::parse_utf16le_string(&data[username_start..username_end])
            } else {
                String::new()
            }
        } else {
            String::new()
        };

        // Parse relative path (if present, after username)
        let rel_path_start = 20 + (username_len as usize) * 2;
        let rel_path = if rel_path_start < data.len() {
            // Relative path is typically null-terminated ASCII
            Self::parse_ascii_string(&data[rel_path_start..])
        } else {
            String::new()
        };

        Ok(Self {
            current_edit_offset,
            release_version,
            username,
            rel_path,
        })
    }

    /// Get the offset to the current UserEditAtom record.
    ///
    /// This offset points to the location in the PowerPoint Document stream
    /// where the current UserEditAtom record is located.
    #[inline]
    pub fn current_edit_offset(&self) -> u32 {
        self.current_edit_offset
    }

    /// Get the username.
    #[inline]
    pub fn username(&self) -> &str {
        &self.username
    }

    /// Get the relative path to the presentation.
    #[inline]
    pub fn relative_path(&self) -> &str {
        &self.rel_path
    }

    /// Get the release version.
    #[inline]
    pub fn release_version(&self) -> u16 {
        self.release_version
    }

    /// Parse a UTF-16LE encoded string from binary data.
    /// Optimized for performance with minimal allocations.
    fn parse_utf16le_string(data: &[u8]) -> String {
        if data.is_empty() || data.len() < 2 {
            return String::new();
        }

        // Pre-allocate capacity for the result
        let estimated_chars = data.len() / 2;
        let mut result = String::with_capacity(estimated_chars);

        // Process in chunks of 2 bytes
        for chunk in data.chunks_exact(2) {
            let code_unit = u16::from_le_bytes([chunk[0], chunk[1]]);

            // Stop at null terminator
            if code_unit == 0 {
                break;
            }

            // Convert to char and add to result
            if let Some(ch) = char::from_u32(code_unit as u32) {
                result.push(ch);
            }
        }

        result.shrink_to_fit();
        result
    }

    /// Parse a null-terminated ASCII string from binary data.
    /// Optimized for performance with minimal allocations.
    fn parse_ascii_string(data: &[u8]) -> String {
        if data.is_empty() {
            return String::new();
        }

        // Find null terminator position
        let null_pos = data.iter().position(|&b| b == 0).unwrap_or(data.len());

        // Convert bytes to string (ASCII-compatible)
        String::from_utf8_lossy(&data[..null_pos]).to_string()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_current_user_min_size() {
        let short_data = vec![0u8; 16];
        let result = CurrentUser::parse(&short_data);
        assert!(result.is_err());
    }

    #[test]
    fn test_current_user_header_validation() {
        let mut data = vec![0u8; 32];
        // Set invalid header token
        data[4] = 0xFF;
        data[5] = 0xFF;
        data[6] = 0xFF;
        data[7] = 0xFF;

        let result = CurrentUser::parse(&data);
        assert!(result.is_err());
    }

    #[test]
    fn test_current_user_valid() {
        let mut data = vec![0u8; 32];
        // Set size (first 4 bytes before header)
        data[0] = 0x1C;
        data[1] = 0x00;
        data[2] = 0x00;
        data[3] = 0x00;
        // CurrentUserHeader starts at byte 4 (16 bytes total)
        // Set valid header token (0xF3D1C4DF)
        data[4] = 0xDF;
        data[5] = 0xC4;
        data[6] = 0xD1;
        data[7] = 0xF3;
        // Set current edit offset (offset 4-7 in header, bytes 8-11 in data)
        data[8] = 0x00;
        data[9] = 0x10;
        data[10] = 0x00;
        data[11] = 0x00;
        // Set username length (offset 8-9 in header, bytes 12-13 in data)
        data[12] = 0x00;
        data[13] = 0x00;
        // Set release version (offset 10-13 in header, bytes 14-17 in data)
        data[14] = 0x03;
        data[15] = 0x00;
        data[16] = 0x00;
        data[17] = 0x00;
        // Set ANSI username length (offset 14-15 in header, bytes 18-19 in data)
        data[18] = 0x00;
        data[19] = 0x00;

        let result = CurrentUser::parse(&data);
        if let Err(ref e) = result {
            eprintln!("Parse error: {:?}", e);
        }
        assert!(result.is_ok());

        let current_user = result.unwrap();
        assert_eq!(current_user.current_edit_offset(), 0x1000);
    }

    #[test]
    fn test_utf16le_parsing() {
        let data = vec![
            0x41, 0x00, // 'A'
            0x42, 0x00, // 'B'
            0x43, 0x00, // 'C'
            0x00, 0x00, // null terminator
        ];

        let result = CurrentUser::parse_utf16le_string(&data);
        assert_eq!(result, "ABC");
    }

    #[test]
    fn test_ascii_parsing() {
        let data = b"Hello\0World";
        let result = CurrentUser::parse_ascii_string(data);
        assert_eq!(result, "Hello");
    }
}
