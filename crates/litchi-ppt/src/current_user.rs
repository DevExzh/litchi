/// CurrentUser stream parser for PowerPoint presentations.
///
/// The CurrentUser stream contains information about the current editing session,
/// including the offset to the current user edit record. This follows Apache POI's
/// CurrentUserAtom implementation.
use super::package::{PptError, Result};

/// Minimum size of CurrentUser stream in bytes
const CURRENT_USER_MIN_SIZE: usize = 28;
const CURRENT_USER_RECORD_TYPE: u16 = 0x0FF6;
const UNENCRYPTED_HEADER_TOKEN: u32 = 0xE391_C05F;
const ENCRYPTED_HEADER_TOKEN: u32 = 0xF3D1_C4DF;

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
    /// Document file version
    document_version: u16,
    /// Whether the encrypted header token is present
    encrypted: bool,
    /// Username (UTF-16LE encoded)
    username: String,
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
    /// - Bytes 0-7: PowerPoint record header
    /// - Bytes 8-27: Fixed CurrentUserAtom fields
    /// - Bytes 28+: ANSI username, release version, and optional UTF-16LE username
    pub fn parse(data: &[u8]) -> Result<Self> {
        if data.len() < CURRENT_USER_MIN_SIZE {
            return Err(PptError::Corrupted(
                "CurrentUser stream too short".to_string(),
            ));
        }

        let ver_instance = u16::from_le_bytes([data[0], data[1]]);
        let record_type = u16::from_le_bytes([data[2], data[3]]);
        if ver_instance != 0 || record_type != CURRENT_USER_RECORD_TYPE {
            return Err(PptError::InvalidFormat(format!(
                "Invalid CurrentUser record header: ver/instance=0x{ver_instance:04X}, type=0x{record_type:04X}"
            )));
        }

        let fixed_size = u32::from_le_bytes([data[8], data[9], data[10], data[11]]);
        if fixed_size != 20 {
            return Err(PptError::InvalidFormat(format!(
                "Invalid CurrentUser fixed size: {fixed_size}"
            )));
        }

        let header_token = u32::from_le_bytes([data[12], data[13], data[14], data[15]]);
        let encrypted = match header_token {
            UNENCRYPTED_HEADER_TOKEN => false,
            ENCRYPTED_HEADER_TOKEN => true,
            _ => {
                return Err(PptError::InvalidFormat(format!(
                    "Invalid CurrentUser header token: 0x{header_token:08X}"
                )));
            },
        };
        let current_edit_offset = u32::from_le_bytes([data[16], data[17], data[18], data[19]]);
        let username_len = u16::from_le_bytes([data[20], data[21]]) as usize;
        if username_len > 255 {
            return Err(PptError::InvalidFormat(format!(
                "Invalid CurrentUser username length: {username_len}"
            )));
        }
        let document_version = u16::from_le_bytes([data[22], data[23]]);

        let ansi_start = CURRENT_USER_MIN_SIZE;
        let Some(release_start) = ansi_start.checked_add(username_len) else {
            return Err(PptError::Corrupted(
                "CurrentUser username length overflow".to_string(),
            ));
        };
        let Some(unicode_start) = release_start.checked_add(4) else {
            return Err(PptError::Corrupted(
                "CurrentUser release offset overflow".to_string(),
            ));
        };
        if unicode_start > data.len() {
            return Err(PptError::Corrupted(
                "CurrentUser stream truncates the ANSI username or release version".to_string(),
            ));
        }

        let release_version_raw = u32::from_le_bytes([
            data[release_start],
            data[release_start + 1],
            data[release_start + 2],
            data[release_start + 3],
        ]);
        let release_version = u16::try_from(release_version_raw).unwrap_or(0);
        let unicode_len = username_len.saturating_mul(2);
        let username = if unicode_start
            .checked_add(unicode_len)
            .is_some_and(|end| end <= data.len())
        {
            Self::parse_utf16le_string(&data[unicode_start..unicode_start + unicode_len])
        } else {
            Self::parse_ansi_string(&data[ansi_start..release_start])
        };

        Ok(Self {
            current_edit_offset,
            release_version,
            document_version,
            encrypted,
            username,
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

    /// Legacy compatibility accessor.
    ///
    /// `CurrentUserAtom` has no relative-path field, so this always returns an empty string.
    #[inline]
    pub fn relative_path(&self) -> &str {
        ""
    }

    /// Get the release version.
    #[inline]
    pub fn release_version(&self) -> u16 {
        self.release_version
    }

    /// Return the document file version stored in the fixed atom fields.
    #[inline]
    pub fn document_version(&self) -> u16 {
        self.document_version
    }

    /// Return whether the stream identifies an encrypted presentation.
    #[inline]
    pub fn is_encrypted(&self) -> bool {
        self.encrypted
    }

    /// Parse a UTF-16LE encoded string from binary data.
    /// Optimized for performance with minimal allocations.
    fn parse_utf16le_string(data: &[u8]) -> String {
        crate::text::extractor::from_utf16le_lossy(data)
    }

    /// Parse the low-byte Unicode fallback stored in `ansiUserName`.
    fn parse_ansi_string(data: &[u8]) -> String {
        let null_pos = data.iter().position(|&b| b == 0).unwrap_or(data.len());
        crate::text::extractor::decode_text_bytes(&data[..null_pos])
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn current_user_stream(ansi_name: &[u8], unicode_name: Option<&str>, token: u32) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(&CURRENT_USER_RECORD_TYPE.to_le_bytes());
        data.extend_from_slice(&0u32.to_le_bytes());
        data.extend_from_slice(&20u32.to_le_bytes());
        data.extend_from_slice(&token.to_le_bytes());
        data.extend_from_slice(&0x1000u32.to_le_bytes());
        data.extend_from_slice(&(ansi_name.len() as u16).to_le_bytes());
        data.extend_from_slice(&0x03F4u16.to_le_bytes());
        data.extend_from_slice(&[3, 0, 0, 0]);
        data.extend_from_slice(ansi_name);
        data.extend_from_slice(&8u32.to_le_bytes());
        if let Some(name) = unicode_name {
            for code_unit in name.encode_utf16() {
                data.extend_from_slice(&code_unit.to_le_bytes());
            }
        }
        let record_len = (data.len() - 8) as u32;
        data[4..8].copy_from_slice(&record_len.to_le_bytes());
        data
    }

    #[test]
    fn test_current_user_min_size() {
        let short_data = vec![0u8; 16];
        let result = CurrentUser::parse(&short_data);
        assert!(result.is_err());
    }

    #[test]
    fn test_current_user_header_validation() {
        let data = current_user_stream(b"A", Some("A"), 0xFFFF_FFFF);

        let result = CurrentUser::parse(&data);
        assert!(result.is_err());
    }

    #[test]
    fn test_current_user_valid() {
        let data = current_user_stream(b"??", Some("😀"), UNENCRYPTED_HEADER_TOKEN);
        let current_user = CurrentUser::parse(&data).unwrap();

        assert_eq!(current_user.current_edit_offset(), 0x1000);
        assert_eq!(current_user.username(), "😀");
        assert_eq!(current_user.release_version(), 8);
        assert_eq!(current_user.document_version(), 0x03F4);
        assert!(!current_user.is_encrypted());
        assert_eq!(current_user.relative_path(), "");
    }

    #[test]
    fn test_utf16le_parsing() {
        let data = vec![
            0x3D, 0xD8, // high surrogate
            0x00, 0xDE, // low surrogate
            0x00, 0x00, // null terminator
        ];

        let result = CurrentUser::parse_utf16le_string(&data);
        assert_eq!(result, "😀");
    }

    #[test]
    fn falls_back_to_low_byte_username_when_unicode_is_omitted() {
        let data = current_user_stream(&[0x80, 0xE9], None, ENCRYPTED_HEADER_TOKEN);
        let current_user = CurrentUser::parse(&data).unwrap();

        assert_eq!(current_user.username(), "\u{80}é");
        assert!(current_user.is_encrypted());
    }
}
