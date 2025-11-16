//! DocumentProperties (DOP) generation for DOC files
//!
//! The DOP contains document-level properties and settings.
//! Based on Microsoft's "[MS-DOC]" specification Section 2.7.1.

/// Generate document properties (DOP) with minimal fields required by Word
///
/// - `facing_pages`: sets DOP.fFacingPages (enables different odd/even pages in UI)
/// - `doc_grpf_ihdt`: section `grpfIhdt` bitfield to derive include header/footer flags
pub fn generate_dop(facing_pages: bool, doc_grpf_ihdt: u8) -> Vec<u8> {
    // DOP size matches Apache POI's DOPAbstractType::getSize() (~0x1F4 bytes)
    let mut dop = vec![0u8; 0x1F4];

    // formatFlags (offset 0x00, 1 byte)
    // bit0: fFacingPages, bit1: fWidowControl, bits5-6: fpc (page shadow mode), default 01
    let mut format_flags: u8 = 0;
    if facing_pages {
        format_flags |= 0x01;
    }
    format_flags |= 0x02; // fWidowControl ON by default
    format_flags |= 0x20; // fpc default 1
    dop[0x00] = format_flags;

    // docinfo5 (offset 0x19A, 2 bytes): set fIncludeHeader (0x1000) / fIncludeFooter (0x2000)
    // Derive from section-level grpfIhdt bits:
    // 0x01=HeaderEven, 0x02=HeaderOdd, 0x10=HeaderFirst; 0x04=FooterEven, 0x08=FooterOdd, 0x20=FooterFirst
    let has_header = (doc_grpf_ihdt & (0x01 | 0x02 | 0x10)) != 0;
    let has_footer = (doc_grpf_ihdt & (0x04 | 0x08 | 0x20)) != 0;
    let mut docinfo5: u16 = 0;
    if has_header {
        docinfo5 |= 0x1000;
    }
    if has_footer {
        docinfo5 |= 0x2000;
    }
    dop[0x19A..0x19C].copy_from_slice(&docinfo5.to_le_bytes());

    dop
}

/// Generate minimal document properties (no facing pages, no headers/footers)
#[inline]
pub fn generate_minimal_dop() -> Vec<u8> {
    generate_dop(false, 0)
}
