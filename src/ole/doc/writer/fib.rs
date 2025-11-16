//! FIB (File Information Block) generation for DOC files
//!
//! The FIB is the central structure in a Word document that contains file metadata
//! and pointers to all other structures in the document.
//!
//! Based on Microsoft's "[MS-DOC]" specification and Apache POI's FIBFieldHandler.
//!
//! # Structure
//!
//! The FIB contains:
//! - Base information (version, encryption status)
//! - csw/cslw arrays (counts and sizes of various structures)
//! - FibRgFcLcb (file character positions and byte counts)
//! - FibRgCswNew (additional data)

/// Error type for DOC operations
pub type DocError = std::io::Error;

/// FIB version for Word 2007+ (LibreOffice also uses this)
/// CRITICAL: Modern Word prefers 0x0101 over 0x00C1!
const FIB_VERSION: u16 = 0x0101;

/// Product version
const PRODUCT_VERSION: u16 = 0x0000;

/// FIB (File Information Block) builder
#[derive(Debug)]
#[allow(dead_code)] // Many fields are for future use
pub struct FibBuilder {
    /// Text stream size
    text_size: u32,
    /// Table stream size  
    table_size: u32,
    /// Main document text position
    main_text_start: u32,
    main_text_length: u32,
    /// Footnote text position
    footnote_start: u32,
    footnote_length: u32,
    /// Header text position
    header_start: u32,
    header_length: u32,
    /// Comment text position
    comment_start: u32,
    comment_length: u32,
    /// Endnote text position
    endnote_start: u32,
    endnote_length: u32,
    /// Textbox text position
    textbox_start: u32,
    textbox_length: u32,
    /// Complex (formatted) flag
    is_complex: bool,

    // File offsets and byte counts for various structures (FibRgFcLcb)
    /// StyleSheet offset and size
    fc_stshf: u32,
    lcb_stshf: u32,
    /// Document properties offset and size
    fc_dop: u32,
    lcb_dop: u32,
    /// Complex table (piece table) offset and size
    fc_clx: u32,
    lcb_clx: u32,
    /// Character bin table offset and size
    fc_plcfbte_chpx: u32,
    lcb_plcfbte_chpx: u32,
    /// Paragraph bin table offset and size
    fc_plcfbte_papx: u32,
    lcb_plcfbte_papx: u32,
    /// Section table offset and size
    fc_plcfsed: u32,
    lcb_plcfsed: u32,
    /// Font table offset and size
    fc_sttbfffn: u32,
    lcb_sttbfffn: u32,
    /// Headers/Footers PLCF (PlcfHdd) offset and size (in 1Table)
    fc_plcfhdd: u32,
    lcb_plcfhdd: u32,

    // FibBase fields that need to be set (Apache POI line 906-914)
    fc_min: u32, // Start of text in WordDocument stream
    fc_mac: u32, // End of text in WordDocument stream
    cb_mac: u32, // Total size of WordDocument stream
}

impl FibBuilder {
    /// Create a new FIB builder
    pub fn new() -> Self {
        Self {
            text_size: 0,
            table_size: 0,
            main_text_start: 0,
            main_text_length: 0,
            footnote_start: 0,
            footnote_length: 0,
            header_start: 0,
            header_length: 0,
            comment_start: 0,
            comment_length: 0,
            endnote_start: 0,
            endnote_length: 0,
            textbox_start: 0,
            textbox_length: 0,
            is_complex: true, // Use complex format by default
            fc_stshf: 0,
            lcb_stshf: 0,
            fc_dop: 0,
            lcb_dop: 0,
            fc_clx: 0,
            lcb_clx: 0,
            fc_plcfbte_chpx: 0,
            lcb_plcfbte_chpx: 0,
            fc_plcfbte_papx: 0,
            lcb_plcfbte_papx: 0,
            fc_plcfsed: 0,
            lcb_plcfsed: 0,
            fc_sttbfffn: 0,
            lcb_sttbfffn: 0,
            fc_plcfhdd: 0,
            lcb_plcfhdd: 0,
            fc_min: 0,
            fc_mac: 0,
            cb_mac: 0,
        }
    }

    /// Set main document text range
    pub fn set_main_text(&mut self, start: u32, length: u32) {
        self.main_text_start = start;
        self.main_text_length = length;
        self.text_size = start + length;
    }

    /// Set table stream size
    pub fn set_table_size(&mut self, size: u32) {
        self.table_size = size;
    }

    /// Set StyleSheet offset and size
    pub fn set_stshf(&mut self, offset: u32, size: u32) {
        self.fc_stshf = offset;
        self.lcb_stshf = size;
    }

    /// Set Document Properties offset and size
    pub fn set_dop(&mut self, offset: u32, size: u32) {
        self.fc_dop = offset;
        self.lcb_dop = size;
    }

    /// Set Complex table (piece table) offset and size
    pub fn set_clx(&mut self, offset: u32, size: u32) {
        self.fc_clx = offset;
        self.lcb_clx = size;
    }

    /// Set Character bin table offset and size
    pub fn set_plcfbte_chpx(&mut self, offset: u32, size: u32) {
        self.fc_plcfbte_chpx = offset;
        self.lcb_plcfbte_chpx = size;
    }

    /// Set Paragraph bin table offset and size
    pub fn set_plcfbte_papx(&mut self, offset: u32, size: u32) {
        self.fc_plcfbte_papx = offset;
        self.lcb_plcfbte_papx = size;
    }

    /// Set Section table offset and size
    pub fn set_plcfsed(&mut self, offset: u32, size: u32) {
        self.fc_plcfsed = offset;
        self.lcb_plcfsed = size;
    }

    /// Set Font table offset and size
    pub fn set_sttbfffn(&mut self, offset: u32, size: u32) {
        self.fc_sttbfffn = offset;
        self.lcb_sttbfffn = size;
    }

    /// Set PlcfHdd (headers/footers PLCF) offset and size
    pub fn set_plcfhdd(&mut self, offset: u32, size: u32) {
        self.fc_plcfhdd = offset;
        self.lcb_plcfhdd = size;
    }

    /// Set ccpHdd (header/footer story character count)
    pub fn set_ccp_hdd(&mut self, length: u32) {
        self.header_length = length;
    }

    /// Set FibBase fields (Apache POI line 906-914)
    pub fn set_base_fields(&mut self, fc_min: u32, fc_mac: u32, cb_mac: u32) {
        self.fc_min = fc_min;
        self.fc_mac = fc_mac;
        self.cb_mac = cb_mac;
    }

    /// Generate the FIB as bytes
    ///
    /// # Returns
    ///
    /// FIB structure as a byte vector
    ///
    /// Size depends on version:
    /// - Word 97-2003 (nFib 0x00C1): ~898 bytes (93 pairs)
    /// - Word 2007+ (nFib 0x0101): ~1242 bytes (136 pairs)
    pub fn generate(&self) -> Result<Vec<u8>, DocError> {
        // Word 2007+ format requires larger FIB
        // 32 (base) + 30 (RgW) + 90 (RgLw) + 1090 (RgFcLcb) = 1242 bytes
        let mut fib = vec![0u8; 1242];

        // Base FIB (32 bytes)
        self.write_base(&mut fib)?;

        // csw (count of shorts in FibRgW)
        fib[32] = 0x0E; // 14 shorts
        fib[33] = 0x00;

        // FibRgW (28 bytes starting at offset 34)
        self.write_fibrgw(&mut fib[34..])?;

        // cslw (count of longs in FibRgLw)
        fib[62] = 0x16; // 22 longs
        fib[63] = 0x00;

        // FibRgLw (88 bytes starting at offset 64)
        self.write_fibrglw(&mut fib[64..])?;

        // cbRgFcLcb (count of file character position and byte count pairs)
        // CRITICAL: Word 2007+ format uses 0x88 (136 pairs) instead of 0x5D (93 pairs)
        fib[152] = 0x88; // 136 pairs for Word 2007+
        fib[153] = 0x00;

        // FibRgFcLcb97 (744 bytes starting at offset 154)
        self.write_fibrgfclcb(&mut fib[154..])?;

        Ok(fib)
    }

    fn write_base(&self, fib: &mut [u8]) -> Result<(), DocError> {
        // Word document magic number
        fib[0..2].copy_from_slice(&0xA5ECu16.to_le_bytes());

        // FIB version
        fib[2..4].copy_from_slice(&FIB_VERSION.to_le_bytes());

        // Product version
        fib[4..6].copy_from_slice(&PRODUCT_VERSION.to_le_bytes());

        // Language ID (0x0409 = English US)
        fib[6..8].copy_from_slice(&0x0409u16.to_le_bytes());

        // pnNext (Next available ID for internal references)
        // MUST be 0 for simple documents (no macros/auto-save)
        fib[8..10].copy_from_slice(&0x0000u16.to_le_bytes());

        // Option flags (FibBase.flags1)
        // Set according to MS-DOC and POI defaults:
        // - fWhichTblStm = 1 (0x0200) => use 1Table
        // - fExtChar = 1 (0x1000)     => required by spec
        // - fComplex = 1 (0x0004)     => we use CLX piece table
        let mut flags: u16 = 0;
        flags |= 0x0200; // fWhichTblStm
        flags |= 0x1000; // fExtChar
        flags |= 0x0004; // fComplex
        fib[10..12].copy_from_slice(&flags.to_le_bytes());

        // Encrypted flag (nFibBack)
        fib[12..14].copy_from_slice(&0x00BFu16.to_le_bytes());

        // lKey (file encryption key - offset 0xe = 14) - 4 bytes (u32)
        fib[14..18].copy_from_slice(&0x00000000u32.to_le_bytes());

        // envr (environment flags - offset 0x12 = 18) - MUST be 0
        fib[18] = 0x00;

        // flags2 (offset 0x13 = 19) - keep 0 per spec (fMac MUST be 0, others ignored)
        fib[19] = 0x00;

        // Chs (offset 0x14 = 20) - 2 bytes (deprecated, set to 0)
        fib[20..22].copy_from_slice(&0x0000u16.to_le_bytes());

        // chsTables (offset 0x16 = 22) - 2 bytes (deprecated, set to 0)
        fib[22..24].copy_from_slice(&0x0000u16.to_le_bytes());

        // fcMin (offset 0x18 = 24) - start of text in WordDocument stream (POI line 906)
        fib[24..28].copy_from_slice(&self.fc_min.to_le_bytes());

        // fcMac (offset 0x1c = 28) - end of text in WordDocument stream (POI line 907)
        fib[28..32].copy_from_slice(&self.fc_mac.to_le_bytes());

        Ok(())
    }

    fn write_fibrgw(&self, buf: &mut [u8]) -> Result<(), DocError> {
        // FibRgW contains various configuration values

        // wMagicCreated (offset 0 in FibRgW = offset 34 in FIB)
        // CRITICAL: Microsoft Word magic signature "jb" = 0x6A62
        // Without this, Word may reject the file!
        buf[0..2].copy_from_slice(&0x6A62u16.to_le_bytes());

        // wMagicRevised (offset 2 in FibRgW = offset 36 in FIB)
        buf[2..4].copy_from_slice(&0x6A62u16.to_le_bytes());

        // wMagicCreatedPrivate (offset 4 in FibRgW = offset 38 in FIB)
        buf[4..6].copy_from_slice(&0x6A62u16.to_le_bytes());

        // wMagicRevisedPrivate (offset 6 in FibRgW = offset 40 in FIB)
        buf[6..8].copy_from_slice(&0x6A62u16.to_le_bytes());

        // Rest are 0 for basic documents
        buf[8..28].fill(0);
        Ok(())
    }

    fn write_fibrglw(&self, buf: &mut [u8]) -> Result<(), DocError> {
        // FibRgLw97 structure - character COUNTS for subdocuments
        // Based on POI's FibRgLw97AbstractType.serialize()

        // field_1_cbMac: Total byte count of the document
        buf[0..4].copy_from_slice(&self.cb_mac.to_le_bytes());

        // field_2_reserved1
        buf[4..8].copy_from_slice(&0u32.to_le_bytes());

        // field_3_reserved2
        buf[8..12].copy_from_slice(&0u32.to_le_bytes());

        // field_4_ccpText: Character count in main document
        buf[12..16].copy_from_slice(&self.main_text_length.to_le_bytes());

        // field_5_ccpFtn: Character count in footnotes
        buf[16..20].copy_from_slice(&self.footnote_length.to_le_bytes());

        // field_6_ccpHdd: Character count in headers/footers
        buf[20..24].copy_from_slice(&self.header_length.to_le_bytes());

        // field_7_reserved3
        buf[24..28].copy_from_slice(&0u32.to_le_bytes());

        // field_8_ccpAtn: Character count in annotations
        buf[28..32].copy_from_slice(&self.comment_length.to_le_bytes());

        // field_9_ccpEdn: Character count in endnotes
        buf[32..36].copy_from_slice(&self.endnote_length.to_le_bytes());

        // field_10_ccpTxbx: Character count in textboxes
        buf[36..40].copy_from_slice(&self.textbox_length.to_le_bytes());

        // field_11_ccpHdrTxbx: Character count in header textboxes
        buf[40..44].copy_from_slice(&0u32.to_le_bytes());

        // field_12-22: Reserved fields
        buf[44..88].fill(0);

        Ok(())
    }

    fn write_fibrgfclcb(&self, buf: &mut [u8]) -> Result<(), DocError> {
        // FibRgFcLcb contains file offsets and byte counts for various structures
        // Each entry is an (fc, lcb) pair - file offset and byte count
        // Based on Apache POI's FIBFieldHandler field indices

        // Field indices from Apache POI's FIBFieldHandler.java
        const STSHF: usize = 1; // StyleSheet
        const PLCFSED: usize = 6; // Section table
        const PLCFBTECHPX: usize = 12; // Character bin table
        const PLCFBTEPAPX: usize = 13; // Paragraph bin table
        const DOP: usize = 31; // Document properties
        const CLX: usize = 33; // Complex table (piece table)

        // Zero all fields first
        buf.fill(0);

        // Helper to set field offset and size
        let set_field = |buf: &mut [u8], field_index: usize, fc: u32, lcb: u32| {
            let offset = field_index * 8;
            if offset + 8 <= buf.len() {
                buf[offset..offset + 4].copy_from_slice(&fc.to_le_bytes());
                buf[offset + 4..offset + 8].copy_from_slice(&lcb.to_le_bytes());
            }
        };

        // Field indices from Apache POI's FIBFieldHandler.java
        const STTBFFFN: usize = 15; // Font table
        const PLCFHDD: usize = 11; // Headers/Footers PLCF

        // Write field offsets and sizes
        set_field(buf, STSHF, self.fc_stshf, self.lcb_stshf);
        set_field(buf, PLCFSED, self.fc_plcfsed, self.lcb_plcfsed);
        set_field(
            buf,
            PLCFBTECHPX,
            self.fc_plcfbte_chpx,
            self.lcb_plcfbte_chpx,
        );
        set_field(
            buf,
            PLCFBTEPAPX,
            self.fc_plcfbte_papx,
            self.lcb_plcfbte_papx,
        );
        set_field(buf, STTBFFFN, self.fc_sttbfffn, self.lcb_sttbfffn); // Font table (POI line 900-903)
        set_field(buf, PLCFHDD, self.fc_plcfhdd, self.lcb_plcfhdd); // Headers/Footers PLCF
        set_field(buf, DOP, self.fc_dop, self.lcb_dop);
        set_field(buf, CLX, self.fc_clx, self.lcb_clx);

        Ok(())
    }
}

impl Default for FibBuilder {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_fib_generation() {
        let mut fib = FibBuilder::new();
        fib.set_main_text(0, 1000);

        let fib_bytes = fib.generate().unwrap();
        assert_eq!(fib_bytes.len(), 1242);

        // Check magic number
        assert_eq!(u16::from_le_bytes([fib_bytes[0], fib_bytes[1]]), 0xA5EC);

        // Check FIB version
        assert_eq!(
            u16::from_le_bytes([fib_bytes[2], fib_bytes[3]]),
            FIB_VERSION
        );
    }
}
