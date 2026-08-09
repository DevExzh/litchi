//! BIFF8 `Theme` record (MS-XLS 2.4.326): the document theme.
//!
//! A custom theme carries its ECMA-376 theme part as an opaque byte stream
//! that may span `ContinueFrt12` records (MS-XLS 2.4.74); the contents are
//! stored verbatim and never interpreted.

use super::{Error, Result};

/// Record type of the `Theme` record.
pub(crate) const THEME_RECORD_TYPE: u16 = 0x0896;
/// Record type of the `ContinueFrt12` record.
pub(crate) const CONTINUE_FRT12_RECORD_TYPE: u16 = 0x087F;

/// Size in bytes of an `FrtHeader`/`FrtHeader12` (MS-XLS 2.5.135/2.5.137).
const FRT_HEADER_LEN: usize = 12;
/// Largest BIFF8 record payload.
const MAX_RECORD_PAYLOAD: usize = 8_224;
/// `dwThemeVersion` value selecting a custom theme with inline contents.
const THEME_VERSION_CUSTOM: u32 = 0;
/// `dwThemeVersion` value selecting the application default theme.
const THEME_VERSION_DEFAULT: u32 = 124_226;

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidRecord {
        record_type: THEME_RECORD_TYPE,
        message: message.into(),
    }
}

/// The theme in use in the document (MS-XLS 2.4.326).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Theme {
    /// Raw `dwThemeVersion` theme type.
    version: u32,
    /// Opaque ECMA-376 theme part, present for custom themes.
    contents: Option<Vec<u8>>,
}

impl Theme {
    /// A custom theme with inline ECMA-376 theme contents.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn custom(contents: Vec<u8>) -> Result<Self> {
        if contents.is_empty() {
            return Err(invalid("custom theme must carry theme contents"));
        }
        Ok(Self {
            version: THEME_VERSION_CUSTOM,
            contents: Some(contents),
        })
    }

    /// The application default theme.
    #[must_use]
    pub fn default_theme() -> Self {
        Self {
            version: THEME_VERSION_DEFAULT,
            contents: None,
        }
    }

    /// Raw `dwThemeVersion` theme type.
    #[must_use]
    pub const fn version(&self) -> u32 {
        self.version
    }

    /// Whether the theme is a custom theme with inline contents.
    #[must_use]
    pub const fn is_custom(&self) -> bool {
        self.version == THEME_VERSION_CUSTOM
    }

    /// Opaque ECMA-376 theme contents, when present.
    #[must_use]
    pub fn contents(&self) -> Option<&[u8]> {
        self.contents.as_deref()
    }

    /// Parse a `Theme` record payload plus the payloads of the
    /// `ContinueFrt12` records that follow it.
    pub(crate) fn parse(data: &[u8], continues: &[Vec<u8>]) -> Result<Self> {
        if data.len() < FRT_HEADER_LEN + 4 {
            return Err(Error::InvalidLength {
                expected: FRT_HEADER_LEN + 4,
                found: data.len(),
            });
        }
        if u16::from_le_bytes([data[0], data[1]]) != THEME_RECORD_TYPE {
            return Err(invalid("Theme FrtHeader.rt mismatch"));
        }
        let version = u32::from_le_bytes(data[12..16].try_into().expect("length checked"));
        let mut contents = data[FRT_HEADER_LEN + 4..].to_vec();
        for (index, continuation) in continues.iter().enumerate() {
            if continuation.len() < FRT_HEADER_LEN {
                return Err(Error::InvalidLength {
                    expected: FRT_HEADER_LEN,
                    found: continuation.len(),
                });
            }
            if u16::from_le_bytes([continuation[0], continuation[1]]) != CONTINUE_FRT12_RECORD_TYPE
            {
                return Err(invalid(format!(
                    "Theme continuation {index} is not a ContinueFrt12 record"
                )));
            }
            contents.extend_from_slice(&continuation[FRT_HEADER_LEN..]);
        }
        let contents = if contents.is_empty() {
            None
        } else {
            Some(contents)
        };
        if version == THEME_VERSION_CUSTOM && contents.is_none() {
            return Err(invalid("custom theme must carry theme contents"));
        }
        Ok(Self { version, contents })
    }

    /// Serialize as a sequence of complete record payloads: the `Theme`
    /// record followed by `ContinueFrt12` records when the contents exceed
    /// one record.
    pub(crate) fn to_record_payloads(&self) -> Vec<Vec<u8>> {
        let mut first = Vec::new();
        first.extend_from_slice(&THEME_RECORD_TYPE.to_le_bytes());
        first.extend_from_slice(&[0; FRT_HEADER_LEN - 2]);
        first.extend_from_slice(&self.version.to_le_bytes());
        let mut records = Vec::new();
        if let Some(contents) = &self.contents {
            let first_chunk = MAX_RECORD_PAYLOAD - (FRT_HEADER_LEN + 4);
            let mut chunks = contents.chunks(first_chunk);
            first.extend_from_slice(chunks.next().expect("nonempty contents"));
            records.push(first);
            for chunk in chunks {
                let mut continuation = Vec::with_capacity(FRT_HEADER_LEN + chunk.len());
                continuation.extend_from_slice(&CONTINUE_FRT12_RECORD_TYPE.to_le_bytes());
                continuation.extend_from_slice(&[0; FRT_HEADER_LEN - 2]);
                continuation.extend_from_slice(chunk);
                records.push(continuation);
            }
        } else {
            records.push(first);
        }
        records
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn theme_record(version: u32, contents: &[u8]) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&THEME_RECORD_TYPE.to_le_bytes());
        data.extend_from_slice(&[0; 10]);
        data.extend_from_slice(&version.to_le_bytes());
        data.extend_from_slice(contents);
        data
    }

    fn continuation(contents: &[u8]) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&CONTINUE_FRT12_RECORD_TYPE.to_le_bytes());
        data.extend_from_slice(&[0; 10]);
        data.extend_from_slice(contents);
        data
    }

    #[test]
    fn parses_custom_theme_with_continuations() {
        let parsed = Theme::parse(
            &theme_record(0, b"<a:theme>"),
            &[continuation(b"part2"), continuation(b"part3")],
        )
        .unwrap();
        assert!(parsed.is_custom());
        assert_eq!(parsed.contents(), Some(b"<a:theme>part2part3".as_slice()));
    }

    #[test]
    fn round_trips_across_record_boundaries() {
        let contents = vec![0x5Au8; 30_000];
        let theme = Theme::custom(contents.clone()).unwrap();
        let payloads = theme.to_record_payloads();
        assert!(payloads.len() > 1);
        for payload in &payloads {
            assert!(payload.len() <= MAX_RECORD_PAYLOAD);
        }
        let parsed = Theme::parse(&payloads[0], &payloads[1..]).unwrap();
        assert_eq!(parsed, theme);
        assert_eq!(parsed.contents(), Some(contents.as_slice()));
    }

    #[test]
    fn parses_default_theme_without_contents() {
        let parsed = Theme::parse(&theme_record(THEME_VERSION_DEFAULT, &[]), &[]).unwrap();
        assert!(!parsed.is_custom());
        assert_eq!(parsed.version(), THEME_VERSION_DEFAULT);
        assert_eq!(parsed.contents(), None);
        let payloads = parsed.to_record_payloads();
        assert_eq!(payloads.len(), 1);
        assert_eq!(payloads[0], theme_record(THEME_VERSION_DEFAULT, &[]));
    }

    #[test]
    fn rejects_malformed_records() {
        // Truncated.
        assert!(Theme::parse(&[0; 10], &[]).is_err());
        // Wrong FrtHeader.rt.
        let mut wrong_rt = theme_record(THEME_VERSION_DEFAULT, &[]);
        wrong_rt[0..2].copy_from_slice(&0x087Fu16.to_le_bytes());
        assert!(Theme::parse(&wrong_rt, &[]).is_err());
        // Custom theme without contents.
        assert!(Theme::parse(&theme_record(0, &[]), &[]).is_err());
        // Continuation with a wrong record type.
        let mut bad = continuation(b"x");
        bad[0..2].copy_from_slice(&0x003Cu16.to_le_bytes());
        assert!(Theme::parse(&theme_record(0, b"a"), &[bad]).is_err());
        // Empty builder contents.
        assert!(Theme::custom(Vec::new()).is_err());
    }
}
