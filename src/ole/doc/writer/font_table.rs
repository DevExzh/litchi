//! Font Table (STTBFFFN) generation for DOC files
//!
//! The font table lists all fonts used in the document.
//! Based on Microsoft's "[MS-DOC]" specification and Apache POI's FontTable.

/// Generate a minimal font table (STTBFFFN) matching Apache POI
///
/// Layout:
/// - stringCount (u16)
/// - extraDataSz (u16) -> 0
/// - repeated FFN structures (cbFfnM1 header + UTF-16LE name with null)
pub fn generate_minimal_font_table() -> Vec<u8> {
    let mut sttbfffn = Vec::new();
    sttbfffn.extend_from_slice(&1u16.to_le_bytes()); // stringCount
    sttbfffn.extend_from_slice(&0u16.to_le_bytes()); // extraDataSz

    let ffn = build_ffn("Times New Roman");
    sttbfffn.extend_from_slice(&ffn);

    sttbfffn
}

/// Builder for STTBFFFN (Font Table)
///
/// Collects fonts used in the document and generates a proper STTBFFFN block.
#[derive(Debug, Default)]
pub struct FontTableBuilder {
    fonts: Vec<String>,
}

impl FontTableBuilder {
    /// Create a new builder with default font "Times New Roman"
    pub fn new() -> Self {
        let mut b = Self { fonts: Vec::new() };
        b.get_or_add("Times New Roman");
        b
    }

    /// Get the index of a font, inserting it if not present
    pub fn get_or_add(&mut self, name: &str) -> u16 {
        if let Some(idx) = self.fonts.iter().position(|f| f.eq_ignore_ascii_case(name)) {
            return idx as u16;
        }
        self.fonts.push(name.to_string());
        (self.fonts.len() as u16) - 1
    }

    /// Generate the STTBFFFN bytes for all collected fonts (POI-compatible)
    pub fn generate(&self) -> Vec<u8> {
        let mut sttbfffn = Vec::new();
        sttbfffn.extend_from_slice(&(self.fonts.len() as u16).to_le_bytes());
        sttbfffn.extend_from_slice(&0u16.to_le_bytes()); // extraDataSz = 0

        for name in &self.fonts {
            let ffn = build_ffn(name);
            sttbfffn.extend_from_slice(&ffn);
        }

        sttbfffn
    }
}

/// Build an FFN structure for a given font name (UTF-16LE, zero-terminated)
fn build_ffn(name: &str) -> Vec<u8> {
    // Header is 1+1+2+1+1+10+24 = 40 bytes, name is UTF-16LE with null terminator
    let mut header = vec![0u8; 40];
    // prq=2, fTrueType=1, ff=1 (Roman)
    header[1] = 0x02 | 0x04 | (0x01 << 4);
    // wWeight = 400
    header[2..4].copy_from_slice(&400u16.to_le_bytes());
    // chs = 0 (ANSI)
    header[4] = 0x00;
    // ixchSzAlt = 0 (no alternate)
    header[5] = 0x00;

    // Name UTF-16LE + terminating 0x0000
    let mut name_bytes = Vec::with_capacity((name.len() + 1) * 2);
    for ch in name.encode_utf16() {
        name_bytes.extend_from_slice(&ch.to_le_bytes());
    }
    name_bytes.extend_from_slice(&0u16.to_le_bytes());

    // cbFfnM1 = total length - 1
    let total_len = header.len() + name_bytes.len();
    header[0] = (total_len as u8).wrapping_sub(1);

    let mut ffn = header;
    ffn.extend_from_slice(&name_bytes);
    ffn
}
