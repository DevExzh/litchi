//! Word embedded TrueType font table (`SttbTtmbd`, MS-DOC 2.9.296).
//!
//! The table lists the TrueType fonts embedded in the document. It is parsed
//! as inert metadata: the font data itself stays in the `WordDocument` stream
//! and is never loaded, installed, or executed.

use super::super::package::{Error as PackageError, Result};
use super::fib::FileInformationBlock;

/// Table-pointer index of `fcSttbTtmbd`/`lcbSttbTtmbd`.
const STTB_TTMBD: usize = 61;
/// Fixed `SttbW6` header size in bytes (MS-DOC 2.9.297).
const HEADER_LEN: usize = 10;
/// Size in bytes of one `Ttmbd` element (MS-DOC 2.9.331).
const TTMBD_LEN: usize = 12;
/// `ibstMac` limit: at most 64 fonts may be embedded in a document.
const MAX_EMBEDDED_FONTS: usize = 64;
/// `ibstMax`: the mandated maximum-fonts-supported value.
const IBST_MAX: u16 = 64;
/// `fcSubset` value marking a font that is embedded in its entirety.
const WHOLE_FONT: u32 = 0xFFFF_FFFF;
/// Largest possible table: a full `u16` `brgbst` plus 64 `Ttmbd` elements.
const MAX_TABLE_BYTES: usize = u16::MAX as usize + MAX_EMBEDDED_FONTS * TTMBD_LEN;

fn corrupted(message: impl Into<String>) -> PackageError {
    PackageError::Corrupted(message.into())
}

fn read_u16(data: &[u8], offset: usize, field: &str) -> Result<u16> {
    litchi_core::binary::read_u16_le(data, offset)
        .map_err(|error| corrupted(format!("invalid {field}: {error}")))
}

fn read_u32(data: &[u8], offset: usize, field: &str) -> Result<u32> {
    litchi_core::binary::read_u32_le(data, offset)
        .map_err(|error| corrupted(format!("invalid {field}: {error}")))
}

/// One embedded TrueType font description (MS-DOC 2.9.331).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct EmbeddedFont {
    /// Offset into the `WordDocument` stream where the font data is stored.
    font_data_offset: u32,
    /// Index into the `SttbfFfn` font table of the embedded font.
    font_table_index: u16,
    bold: bool,
    italic: bool,
    /// First-use order when only used characters are embedded.
    subset_order: Option<u32>,
}

impl EmbeddedFont {
    /// Offset into the `WordDocument` stream where the embedded font data
    /// begins. Inert metadata: this library never loads the font data.
    #[must_use]
    pub fn font_data_offset(&self) -> u32 {
        self.font_data_offset
    }

    /// Index into the `SttbfFfn` font-name table of the embedded font.
    #[must_use]
    pub fn font_table_index(&self) -> u16 {
        self.font_table_index
    }

    /// Whether the embedded font is bold.
    #[must_use]
    pub fn is_bold(&self) -> bool {
        self.bold
    }

    /// Whether the embedded font is italic.
    #[must_use]
    pub fn is_italic(&self) -> bool {
        self.italic
    }

    /// The font's first-use order when only the characters used by the
    /// document are embedded, or `None` when the entire font is embedded.
    #[must_use]
    pub fn subset_order(&self) -> Option<u32> {
        self.subset_order
    }
}

/// The embedded TrueType fonts listed in a document's `SttbTtmbd` table.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct DocumentEmbeddedFonts {
    fonts: Vec<EmbeddedFont>,
}

impl DocumentEmbeddedFonts {
    /// The embedded-font descriptions in table order.
    #[must_use]
    pub fn fonts(&self) -> &[EmbeddedFont] {
        &self.fonts
    }

    /// Parse the `SttbTtmbd` from the table stream, or `None` when the
    /// document embeds no TrueType fonts.
    pub fn parse(
        fib: &FileInformationBlock,
        table_stream: &[u8],
    ) -> Result<Option<DocumentEmbeddedFonts>> {
        let Some((offset, length)) = fib.get_table_pointer(STTB_TTMBD) else {
            return Ok(None);
        };
        if length == 0 {
            return Ok(None);
        }
        let length = usize::try_from(length)
            .map_err(|_| corrupted("SttbTtmbd length does not fit in memory"))?;
        if length > MAX_TABLE_BYTES {
            return Err(corrupted(
                "SttbTtmbd exceeds its specification-derived size cap",
            ));
        }
        let start = usize::try_from(offset)
            .map_err(|_| corrupted("SttbTtmbd offset does not fit in memory"))?;
        let end = start
            .checked_add(length)
            .ok_or_else(|| corrupted("SttbTtmbd range overflows"))?;
        let data = table_stream
            .get(start..end)
            .ok_or_else(|| corrupted("SttbTtmbd extends past the table stream"))?;
        Self::parse_bytes(data).map(Some)
    }

    /// Parse one complete `SttbTtmbd` payload.
    pub fn parse_bytes(data: &[u8]) -> Result<DocumentEmbeddedFonts> {
        if data.len() < HEADER_LEN {
            return Err(corrupted("SttbTtmbd header is truncated"));
        }
        // `unused1` and `unused2` are ignored per MS-DOC 2.9.297.
        let count = read_u16(data, 2, "SttbTtmbd ibstMac")?;
        if count > MAX_EMBEDDED_FONTS as u16 {
            return Err(corrupted("SttbTtmbd ibstMac exceeds 64"));
        }
        if read_u16(data, 4, "SttbTtmbd ibstMax")? != IBST_MAX {
            return Err(corrupted("SttbTtmbd ibstMax is not 64"));
        }
        let entry_offset = usize::from(read_u16(data, 8, "SttbTtmbd brgbst")?);
        let count = usize::from(count);
        if count > 0 {
            let entries_end = entry_offset
                .checked_add(
                    count
                        .checked_mul(TTMBD_LEN)
                        .ok_or_else(|| corrupted("SttbTtmbd element range overflows"))?,
                )
                .ok_or_else(|| corrupted("SttbTtmbd element range overflows"))?;
            if entry_offset < HEADER_LEN || entries_end > data.len() {
                return Err(corrupted("SttbTtmbd element array is truncated"));
            }
        }

        let mut fonts = Vec::with_capacity(count);
        for index in 0..count {
            let base = entry_offset + index * TTMBD_LEN;
            let field = |name: &str| format!("SttbTtmbd font {index} {name}");
            let font_data_offset = read_u32(data, base, &field("fc"))?;
            if font_data_offset == 0 {
                return Err(corrupted(format!("SttbTtmbd font {index} fc is zero")));
            }
            let font_table_index = read_u16(data, base + 4, &field("iiffn"))?;
            if font_table_index & 0x8000 != 0 {
                return Err(corrupted(format!(
                    "SttbTtmbd font {index} iiffn is negative"
                )));
            }
            let flags = read_u16(data, base + 6, &field("flags"))?;
            let fc_subset = read_u32(data, base + 8, &field("fcSubset"))?;
            fonts.push(EmbeddedFont {
                font_data_offset,
                font_table_index,
                bold: flags & 0x0001 != 0,
                italic: flags & 0x0002 != 0,
                subset_order: (fc_subset != WHOLE_FONT).then_some(fc_subset),
            });
        }
        Ok(DocumentEmbeddedFonts { fonts })
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn sttb_ttmbd(brgbst: u16, fonts: &[(u32, u16, u16, u32)]) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&0u16.to_le_bytes()); // unused1
        data.extend_from_slice(&(fonts.len() as u16).to_le_bytes()); // ibstMac
        data.extend_from_slice(&IBST_MAX.to_le_bytes()); // ibstMax
        data.extend_from_slice(&0u16.to_le_bytes()); // unused2
        data.extend_from_slice(&brgbst.to_le_bytes()); // brgbst
        data.resize(usize::from(brgbst).max(HEADER_LEN), 0);
        for &(fc, iiffn, flags, fc_subset) in fonts {
            data.extend_from_slice(&fc.to_le_bytes());
            data.extend_from_slice(&iiffn.to_le_bytes());
            data.extend_from_slice(&flags.to_le_bytes());
            data.extend_from_slice(&fc_subset.to_le_bytes());
        }
        data
    }

    #[test]
    fn parses_embedded_font_entries() {
        let data = sttb_ttmbd(
            HEADER_LEN as u16,
            &[(4096, 3, 0x0003, 0), (8192, 5, 0x0000, WHOLE_FONT)],
        );
        let parsed = DocumentEmbeddedFonts::parse_bytes(&data).unwrap();
        assert_eq!(parsed.fonts().len(), 2);
        let first = parsed.fonts()[0];
        assert_eq!(first.font_data_offset(), 4096);
        assert_eq!(first.font_table_index(), 3);
        assert!(first.is_bold());
        assert!(first.is_italic());
        assert_eq!(first.subset_order(), Some(0));
        let second = parsed.fonts()[1];
        assert!(!second.is_bold());
        assert_eq!(second.subset_order(), None);
    }

    #[test]
    fn honors_a_nonstandard_brgbst() {
        // Word is known to emit values other than the recommended 10.
        let data = sttb_ttmbd(26, &[]);
        let parsed = DocumentEmbeddedFonts::parse_bytes(&data).unwrap();
        assert!(parsed.fonts().is_empty());
    }

    #[test]
    fn rejects_malformed_tables() {
        // Truncated header.
        assert!(DocumentEmbeddedFonts::parse_bytes(&sttb_ttmbd(10, &[])[..8]).is_err());
        // ibstMac above the 64-font limit.
        let mut too_many = sttb_ttmbd(10, &[]);
        too_many[2] = 65;
        assert!(DocumentEmbeddedFonts::parse_bytes(&too_many).is_err());
        // ibstMax other than 64.
        let mut wrong_max = sttb_ttmbd(10, &[]);
        wrong_max[4] = 0;
        assert!(DocumentEmbeddedFonts::parse_bytes(&wrong_max).is_err());
        // Zero font-data offset.
        assert!(
            DocumentEmbeddedFonts::parse_bytes(&sttb_ttmbd(10, &[(0, 0, 0, WHOLE_FONT)])).is_err()
        );
        // Negative font-table index.
        assert!(
            DocumentEmbeddedFonts::parse_bytes(&sttb_ttmbd(10, &[(4096, 0x8000, 0, WHOLE_FONT)]))
                .is_err()
        );
        // Element array shorter than ibstMac declares.
        let mut truncated = sttb_ttmbd(10, &[(4096, 0, 0, WHOLE_FONT)]);
        truncated.truncate(HEADER_LEN + 4);
        assert!(DocumentEmbeddedFonts::parse_bytes(&truncated).is_err());
    }
}
