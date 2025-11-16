//! StyleSheet (STSH) generation for DOC files
//!
//! The StyleSheet contains style definitions required for document formatting.
//! Based on Microsoft's "[MS-DOC]" specification Section 2.9.271.

/// Generate a minimal stylesheet
///
/// Creates a basic stylesheet with Normal style only
/// Based on Apache POI's StyleSheet.writeTo() implementation
pub fn generate_minimal_stylesheet() -> Vec<u8> {
    let mut stsh = Vec::new();

    // Write cbStshi (size of STSHI) first (line 139-144)
    // STSHI is always 18 bytes for Word 97-2003
    let cb_stshi = 18u16;
    stsh.extend_from_slice(&cb_stshi.to_le_bytes());

    // STSHI (Style Sheet Information) - 18 bytes
    // cstd (count of styles) = 1 (just Normal style)
    stsh.extend_from_slice(&1u16.to_le_bytes());

    // cbSTDBaseInFile (size of STD base)
    stsh.extend_from_slice(&10u16.to_le_bytes());

    // fStdStylenamesWritten
    stsh.extend_from_slice(&1u16.to_le_bytes());

    // stiMaxWhenSaved (max style identifier)
    stsh.extend_from_slice(&1u16.to_le_bytes());

    // istdMaxFixedWhenSaved
    stsh.extend_from_slice(&1u16.to_le_bytes());

    // nVerBuiltInNamesWhenSaved
    stsh.extend_from_slice(&0u16.to_le_bytes());

    // ftcAsci, ftcFE, ftcOther, ftcBi (default fonts - all 0 = default)
    for _ in 0..4 {
        stsh.extend_from_slice(&0u16.to_le_bytes());
    }

    // Now write STD (Style Definition) for Normal style
    // POI writes: 2-byte size + STD data (line 153-171)

    // Build STD data first
    let mut std_data = Vec::new();

    // sti (style identifier) = 0 (Normal)
    std_data.extend_from_slice(&0u16.to_le_bytes());

    // sgc (style type) = 1 (paragraph style)
    std_data.extend_from_slice(&1u16.to_le_bytes());

    // istdBase (base style) = 0xFFF (no base)
    std_data.extend_from_slice(&0x0FFFu16.to_le_bytes());

    // cupx (count of UPX) = 0 (no formatting)
    std_data.extend_from_slice(&0u16.to_le_bytes());

    // bchUpe (size of UPX) = 0
    std_data.extend_from_slice(&0u16.to_le_bytes());

    // grupe (UPX array) - empty for minimal

    // Style name length
    std_data.extend_from_slice(&6u16.to_le_bytes()); // "Normal" = 6 chars

    // Style name in Unicode (UTF-16LE)
    for c in "Normal".encode_utf16() {
        std_data.extend_from_slice(&c.to_le_bytes());
    }

    // POI adjusts size to word boundary (line 159)
    let std_size = std_data.len() as u16;
    let adjusted_size = std_size + (std_size % 2);

    // Write 2-byte size (line 159-160)
    stsh.extend_from_slice(&adjusted_size.to_le_bytes());

    // Write STD data (line 161)
    stsh.extend_from_slice(&std_data);

    // Add padding byte if needed to align to word boundary (line 163-165)
    if std_size % 2 == 1 {
        stsh.push(0);
    }

    stsh
}
