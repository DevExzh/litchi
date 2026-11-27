//! Byte codecs for the writer-side File Information Block.

use super::IoError;
use super::flags::{BaseFlags, FIB_BASE_VERSION, PRODUCT_VERSION};
use super::header::FibBuilder;
use super::validation::validate_layout;

/// Word 2002 FIB size: 136 fc/lcb pairs and two new shorts.
pub(super) const FIB_SIZE: usize = 32 + 2 + 28 + 2 + 88 + 2 + (136 * 8) + 2 + 2 + 2;

impl FibBuilder {
    /// Generate the complete FIB as bytes.
    ///
    /// The layout is the Word 2002 form (nFibNew = 0x0101) with 136
    /// `FibRgFcLcb` pairs. FibBase.nFib remains 0x00C1 as required by [MS-DOC].
    pub fn generate(&self) -> Result<Vec<u8>, IoError> {
        let mut fib = vec![0u8; FIB_SIZE];
        validate_layout(&fib)?;

        // Base FIB (32 bytes).
        self.write_base(&mut fib)?;

        // csw (count of shorts in FibRgW) = 14.
        fib[32..34].copy_from_slice(&0x000Eu16.to_le_bytes());

        // FibRgW (28 bytes starting at offset 34).
        self.write_fibrgw(&mut fib[34..])?;

        // cslw (count of longs in FibRgLw) = 22.
        fib[62..64].copy_from_slice(&0x0016u16.to_le_bytes());

        // FibRgLw (88 bytes starting at offset 64).
        self.write_fibrglw(&mut fib[64..])?;

        // cbRgFcLcb (count of fc/lcb pairs) = 136.
        fib[152..154].copy_from_slice(&0x0088u16.to_le_bytes());

        // FibRgFcLcb (1088 bytes starting at offset 154).
        self.offsets.write_into(&mut fib[154..]);

        // FibRgCswNew at offset 1242.
        let offset = 154 + 136 * 8;
        fib[offset..offset + 2].copy_from_slice(&0x0002u16.to_le_bytes());
        fib[offset + 2..offset + 4].copy_from_slice(&0x0101u16.to_le_bytes());
        // The second short is reserved and remains zero.

        Ok(fib)
    }

    fn write_base(&self, fib: &mut [u8]) -> Result<(), IoError> {
        fib[0..2].copy_from_slice(&0xA5ECu16.to_le_bytes());
        fib[2..4].copy_from_slice(&FIB_BASE_VERSION.to_le_bytes());
        fib[4..6].copy_from_slice(&PRODUCT_VERSION.to_le_bytes());
        fib[6..8].copy_from_slice(&0x0409u16.to_le_bytes());
        fib[8..10].copy_from_slice(&self.header.next_fib_page.to_le_bytes());

        let flags = BaseFlags {
            complex: self.header.flags.complex,
            glossary: self.header.flags.glossary,
            template: self.header.flags.template,
        }
        .encode();
        fib[10..12].copy_from_slice(&flags.to_le_bytes());

        // nFibBack, lKey, envr, and flags2.
        fib[12..14].copy_from_slice(&0x00BFu16.to_le_bytes());
        fib[14..18].copy_from_slice(&0x00000000u32.to_le_bytes());
        fib[18] = 0;
        fib[19] = 0;

        // Deprecated character-set fields.
        fib[20..22].copy_from_slice(&0x0000u16.to_le_bytes());
        fib[22..24].copy_from_slice(&0x0000u16.to_le_bytes());
        fib[24..28].copy_from_slice(&self.header.fc_min.to_le_bytes());
        fib[28..32].copy_from_slice(&self.header.fc_mac.to_le_bytes());

        Ok(())
    }

    fn write_fibrgw(&self, buf: &mut [u8]) -> Result<(), IoError> {
        // Microsoft Word magic signature "jb" in all four magic fields.
        buf[0..2].copy_from_slice(&0x6A62u16.to_le_bytes());
        buf[2..4].copy_from_slice(&0x6A62u16.to_le_bytes());
        buf[4..6].copy_from_slice(&0x6A62u16.to_le_bytes());
        buf[6..8].copy_from_slice(&0x6A62u16.to_le_bytes());
        buf[8..28].fill(0);
        Ok(())
    }

    fn write_fibrglw(&self, buf: &mut [u8]) -> Result<(), IoError> {
        let stories = &self.stories;

        // FibRgLw97 character counts and stream size.
        buf[0..4].copy_from_slice(&self.header.cb_mac.to_le_bytes());
        buf[4..8].copy_from_slice(&0u32.to_le_bytes());
        buf[8..12].copy_from_slice(&0u32.to_le_bytes());
        buf[12..16].copy_from_slice(&stories.main_text_length.to_le_bytes());
        buf[16..20].copy_from_slice(&stories.footnote_length.to_le_bytes());
        buf[20..24].copy_from_slice(&stories.header_length.to_le_bytes());
        buf[24..28].copy_from_slice(&0u32.to_le_bytes());
        buf[28..32].copy_from_slice(&stories.comment_length.to_le_bytes());
        buf[32..36].copy_from_slice(&stories.endnote_length.to_le_bytes());
        buf[36..40].copy_from_slice(&stories.textbox_length.to_le_bytes());
        buf[40..44].copy_from_slice(&stories.header_textbox_length.to_le_bytes());
        buf[44..88].fill(0);
        Ok(())
    }
}
