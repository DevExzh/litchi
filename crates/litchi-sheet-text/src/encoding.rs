//! Source text decoding.

use std::borrow::Cow;

use litchi_core::BomKind;

use crate::error::{Error, Result};

/// How source bytes are decoded.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
#[non_exhaustive]
pub enum Encoding {
    /// Read the byte-order mark (UTF-8, UTF-16 LE or UTF-16 BE); without
    /// one, UTF-8.
    #[default]
    Auto,
    /// UTF-8, with or without a byte-order mark.
    Utf8,
    /// UTF-16 little endian, with or without a byte-order mark.
    Utf16Le,
    /// UTF-16 big endian, with or without a byte-order mark.
    Utf16Be,
}

impl Encoding {
    /// The canonical name: `auto`, `utf-8`, `utf-16le` or `utf-16be`.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Auto => "auto",
            Self::Utf8 => "utf-8",
            Self::Utf16Le => "utf-16le",
            Self::Utf16Be => "utf-16be",
        }
    }

    /// Parse a name, ignoring ASCII case: `auto`, `utf-8`/`utf8`,
    /// `utf-16le`/`utf16le` or `utf-16be`/`utf16be`.
    #[must_use]
    pub fn parse(name: &str) -> Option<Self> {
        let name = name.trim();
        [
            ("auto", Self::Auto),
            ("utf-8", Self::Utf8),
            ("utf8", Self::Utf8),
            ("utf-16le", Self::Utf16Le),
            ("utf16le", Self::Utf16Le),
            ("utf-16be", Self::Utf16Be),
            ("utf16be", Self::Utf16Be),
        ]
        .into_iter()
        .find_map(|(spelling, encoding)| name.eq_ignore_ascii_case(spelling).then_some(encoding))
    }

    /// Decode `source`, removing its byte-order mark, and report the mark
    /// found.
    ///
    /// UTF-8 input is borrowed; UTF-16 input is decoded into a new string.
    ///
    /// ```
    /// use litchi_core::BomKind;
    /// use litchi_sheet_text::Encoding;
    ///
    /// let (text, bom) = Encoding::Auto.decode(b"\xEF\xBB\xBFa,b")?;
    /// assert_eq!((text.as_ref(), bom), ("a,b", Some(BomKind::Utf8)));
    /// let (text, bom) = Encoding::Auto.decode(b"\xFF\xFEa\x00")?;
    /// assert_eq!((text.as_ref(), bom), ("a", Some(BomKind::Utf16Le)));
    /// # Ok::<(), litchi_sheet_text::Error>(())
    /// ```
    ///
    /// # Errors
    ///
    /// [`Error::Encoding`] at the first invalid sequence, and
    /// [`Error::Allocation`] when the decoded text cannot be allocated.
    pub fn decode(self, source: &[u8]) -> Result<(Cow<'_, str>, Option<BomKind>)> {
        const UTF8_BOM: &[u8] = &litchi_core::UTF8_BOM;
        let encoding = match self {
            Self::Auto if source.starts_with(&litchi_core::UTF16_LE_BOM) => Self::Utf16Le,
            Self::Auto if source.starts_with(&litchi_core::UTF16_BE_BOM) => Self::Utf16Be,
            Self::Auto => Self::Utf8,
            explicit => explicit,
        };
        match encoding {
            Self::Utf16Le | Self::Utf16Be => {
                let little = encoding == Self::Utf16Le;
                let (mark, kind): (&[u8], BomKind) = if little {
                    (&litchi_core::UTF16_LE_BOM, BomKind::Utf16Le)
                } else {
                    (&litchi_core::UTF16_BE_BOM, BomKind::Utf16Be)
                };
                let found = source.starts_with(mark);
                let skip = if found { 2 } else { 0 };
                let body = &source[skip..];
                if body.len() % 2 != 0 {
                    return Err(Error::Encoding {
                        offset: source.len() - 1,
                    });
                }
                let units = body.chunks_exact(2).map(|pair| {
                    let pair = [pair[0], pair[1]];
                    if little {
                        u16::from_le_bytes(pair)
                    } else {
                        u16::from_be_bytes(pair)
                    }
                });
                let mut text = String::new();
                text.try_reserve(body.len())?;
                let mut offset = skip;
                for unit in char::decode_utf16(units) {
                    match unit {
                        Ok(character) => {
                            text.push(character);
                            offset += character.len_utf16() * 2;
                        },
                        Err(_) => return Err(Error::Encoding { offset }),
                    }
                }
                Ok((Cow::Owned(text), found.then_some(kind)))
            },
            Self::Utf8 | Self::Auto => {
                let found = source.starts_with(UTF8_BOM);
                let skip = if found { 3 } else { 0 };
                std::str::from_utf8(&source[skip..])
                    .map(|text| (Cow::Borrowed(text), found.then_some(BomKind::Utf8)))
                    .map_err(|error| Error::Encoding {
                        offset: skip + error.valid_up_to(),
                    })
            },
        }
    }
}
