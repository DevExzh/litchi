//! StyleSheet (STSH) generation for DOC files
//!
//! The StyleSheet contains style definitions required for document formatting.
//! Based on Microsoft's "[MS-DOC]" specification Section 2.9.271 and
//! Apache POI's StyleSheet.java / StdfBaseAbstractType.java.
//!
//! # Structure
//!
//! The STSH (Style Sheet) consists of:
//! - `cbStshi` (u16): Size of the STSHI header
//! - STSHI header (Stshif): General stylesheet information
//! - Array of LPStd entries: Each is a 2-byte size followed by STD data
//!
//! The StdfBase within each STD uses **bit-packed** fields, not separate u16s.

/// Minimum number of styles required by the MS-DOC spec.
/// Per spec: "cstd MUST be equal to or greater than 0x000F"
const MIN_CSTD: u16 = 0x000F;

/// StdfBase size in bytes (5 packed shorts = 10 bytes)
const STDF_BASE_SIZE: u16 = 10;

/// Build a bit-packed StdfBase (10 bytes) for a style definition.
///
/// # Layout (per `StdfBaseAbstractType.java`):
/// - Word 0 (info1): sti[11:0], fScratch[12], fInvalHeight[13], fHasUpe[14], fMassCopy[15]
/// - Word 1 (info2): stk[3:0], istdBase[15:4]
/// - Word 2 (info3): cupx[3:0], istdNext[15:4]
/// - Word 3: bchUpe (u16)
/// - Word 4: grfstd (u16)
fn build_stdf_base(sti: u16, stk: u16, istd_base: u16, cupx: u16, istd_next: u16) -> [u8; 10] {
    let mut buf = [0u8; 10];

    // info1: sti in bits 0-11, flags in bits 12-15 (all 0)
    let info1: u16 = sti & 0x0FFF;
    buf[0..2].copy_from_slice(&info1.to_le_bytes());

    // info2: stk in bits 0-3, istdBase in bits 4-15
    let info2: u16 = (stk & 0x000F) | ((istd_base & 0x0FFF) << 4);
    buf[2..4].copy_from_slice(&info2.to_le_bytes());

    // info3: cupx in bits 0-3, istdNext in bits 4-15
    let info3: u16 = (cupx & 0x000F) | ((istd_next & 0x0FFF) << 4);
    buf[4..6].copy_from_slice(&info3.to_le_bytes());

    // bchUpe = 0, grfstd = 0
    // Already zero-initialized

    buf
}

/// Build the STD byte array for the Normal (istd=0) paragraph style.
///
/// Based on Apache POI's `StyleDescription.toByteArray()`.
/// For a paragraph style, cupx=2 means two UPXs: one for paragraph (PAPX) and
/// one for character (CHPX). Both are empty (size 0) in a minimal stylesheet.
fn build_normal_style_std() -> Vec<u8> {
    let mut std_data = Vec::new();

    // StdfBase: sti=0 (Normal), stk=1 (paragraph), istdBase=0xFFF (none),
    //           cupx=2 (paragraph+character UPX), istdNext=0 (Normal)
    let stdf_base = build_stdf_base(0, 1, 0x0FFF, 2, 0);
    std_data.extend_from_slice(&stdf_base);

    // Style name: length (u16) + UTF-16LE chars + null terminator (u16)
    let name = "Normal";
    let name_len = name.len() as u16;
    std_data.extend_from_slice(&name_len.to_le_bytes());
    for c in name.encode_utf16() {
        std_data.extend_from_slice(&c.to_le_bytes());
    }
    // Null terminator after name (UTF-16LE)
    std_data.extend_from_slice(&0u16.to_le_bytes());

    // UPX 1: Paragraph formatting (PAPX) - empty
    std_data.extend_from_slice(&0u16.to_le_bytes()); // upxSize = 0

    // UPX 2: Character formatting (CHPX) - empty
    std_data.extend_from_slice(&0u16.to_le_bytes()); // upxSize = 0

    std_data
}

/// Build the STD byte array for the Default Paragraph Font (istd=10) character style.
///
/// This is a required built-in character style (sti=10, stk=2).
fn build_default_paragraph_font_std() -> Vec<u8> {
    let mut std_data = Vec::new();

    // StdfBase: sti=10, stk=2 (character), istdBase=0xFFF (none),
    //           cupx=1 (character UPX only), istdNext=10 (self)
    let stdf_base = build_stdf_base(10, 2, 0x0FFF, 1, 10);
    std_data.extend_from_slice(&stdf_base);

    // Style name
    let name = "Default Paragraph Font";
    let name_len = name.len() as u16;
    std_data.extend_from_slice(&name_len.to_le_bytes());
    for c in name.encode_utf16() {
        std_data.extend_from_slice(&c.to_le_bytes());
    }
    // Null terminator
    std_data.extend_from_slice(&0u16.to_le_bytes());

    // UPX 1: Character formatting (CHPX) - empty
    std_data.extend_from_slice(&0u16.to_le_bytes());

    std_data
}

/// Generate a minimal but spec-compliant stylesheet.
///
/// Creates a stylesheet with:
/// - Normal style (istd=0, paragraph style)
/// - Default Paragraph Font (istd=10, character style)
/// - All other required slots (istd 1-14) as null entries
///
/// Based on Apache POI's `StyleSheet.writeTo()` and MS-DOC spec Section 2.9.271.
pub fn generate_minimal_stylesheet() -> Vec<u8> {
    let mut stsh = Vec::new();

    // cbStshi (size of STSHI = Stshif) = 18 bytes
    let cb_stshi = 18u16;
    stsh.extend_from_slice(&cb_stshi.to_le_bytes());

    // Stshif (18 bytes) - General stylesheet information
    // Per StshifAbstractType.java: 9 fields × 2 bytes = 18 bytes
    let cstd = MIN_CSTD; // Must be >= 0x000F
    stsh.extend_from_slice(&cstd.to_le_bytes()); // cstd
    stsh.extend_from_slice(&STDF_BASE_SIZE.to_le_bytes()); // cbSTDBaseInFile = 10
    stsh.extend_from_slice(&1u16.to_le_bytes()); // info3: fHasOriginalStyle=1
    stsh.extend_from_slice(&cstd.to_le_bytes()); // stiMaxWhenSaved
    stsh.extend_from_slice(&(cstd - 1).to_le_bytes()); // istdMaxFixedWhenSaved
    stsh.extend_from_slice(&0u16.to_le_bytes()); // nVerBuiltInNamesWhenSaved
    stsh.extend_from_slice(&0u16.to_le_bytes()); // ftcAsci (default font)
    stsh.extend_from_slice(&0u16.to_le_bytes()); // ftcFE (default font)
    stsh.extend_from_slice(&0u16.to_le_bytes()); // ftcOther (default font)

    // Write LPStd array (cstd entries)
    // Each entry: 2-byte cbStd + STD data (or cbStd=0 for null entry)
    for istd in 0..cstd {
        let std_data = match istd {
            0 => Some(build_normal_style_std()),
            10 => Some(build_default_paragraph_font_std()),
            _ => None, // Null entry
        };

        if let Some(data) = std_data {
            // cbStd: adjusted to word boundary per POI line 159
            let std_size = data.len() as u16;
            let adjusted_size = std_size + (std_size % 2);
            stsh.extend_from_slice(&adjusted_size.to_le_bytes());
            stsh.extend_from_slice(&data);
            // Pad to word boundary if needed
            if std_size % 2 == 1 {
                stsh.push(0);
            }
        } else {
            // Null style entry: cbStd = 0
            stsh.extend_from_slice(&0u16.to_le_bytes());
        }
    }

    stsh
}
