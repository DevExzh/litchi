//! Section table (PLCF of SEDs) generation for DOC files
//!
//! The section table defines document sections and page layout.
//! Based on Microsoft's "[MS-DOC]" specification Section 2.9.245 and
//! Apache POI's SectionTable implementation.

/// Generate SEPX (Section Properties) structure with optional first/odd-even header/footer flags
///
/// - When `first_page_header` is true, emits `sprmSFTitlePage` to enable different first page.
/// - When `grpf_ihdt` != 0, emits `sprmSGprfIhdt` to declare which headers/footers exist in this section.
///   Bits follow LibreOffice nsHdFtFlags/Word semantics:
///   0x01=HeaderEven, 0x02=HeaderOdd, 0x04=FooterEven, 0x08=FooterOdd, 0x10=HeaderFirst, 0x20=FooterFirst
pub fn generate_sepx(first_page_header: bool, grpf_ihdt: u8) -> Vec<u8> {
    let mut grpprl: Vec<u8> = Vec::with_capacity(8);
    if first_page_header {
        // sprmSFTitlePage (u16 opcode) + 1-byte operand (1)
        grpprl.extend_from_slice(&crate::ole::sprm_operations::SPRM_S_F_TITLE_PAGE.to_le_bytes());
        grpprl.push(1u8);
    }
    if grpf_ihdt != 0 {
        // sprmSGprfIhdt (u16 opcode) + 1-byte operand (bitfield)
        grpprl.extend_from_slice(&crate::ole::sprm_operations::SPRM_S_GPRF_IHDT.to_le_bytes());
        grpprl.push(grpf_ihdt);
    }
    let size = grpprl.len() as u16;
    let mut sepx = Vec::with_capacity(2 + grpprl.len());
    sepx.extend_from_slice(&size.to_le_bytes());
    sepx.extend_from_slice(&grpprl);
    sepx
}

/// Generate minimal SEPX (Section Properties) structure (no section SPRMs)
#[inline]
pub fn generate_minimal_sepx() -> Vec<u8> {
    generate_sepx(false, 0)
}

/// Generate section table (PLCF of SEDs)
///
/// Creates a single section covering the entire document
///
/// # Arguments
///
/// * `text_length` - Total length of document text in characters  
/// * `sepx_offset` - Offset in WordDocument stream where SEPX was written
pub fn generate_section_table(text_length: u32, sepx_offset: u32) -> Vec<u8> {
    let mut plcfsed = Vec::new();

    // PLCF structure (Apache POI's PlexOfCps):
    // - Array of n+1 CPs (character positions)
    // - Array of n data elements (SEDs)

    // We have 1 section, so we need 2 CPs

    // CP[0] = 0 (start of document)
    plcfsed.extend_from_slice(&0u32.to_le_bytes());

    // CP[1] = text_length (end of document)
    plcfsed.extend_from_slice(&text_length.to_le_bytes());

    // SED (Section Descriptor) - 12 bytes (POI's SectionDescriptor.toByteArray())

    // fn (short) - used internally by Word - 0 for new documents
    plcfsed.extend_from_slice(&0u16.to_le_bytes());

    // fcSepx (int) - CRITICAL: Must point to SEPX in WordDocument stream
    // Apache POI sets this to the offset where SEPX was written (line 195)
    plcfsed.extend_from_slice(&sepx_offset.to_le_bytes());

    // fnMpr (short) - used internally - 0
    plcfsed.extend_from_slice(&0u16.to_le_bytes());

    // fcMpr (int) - Mac print record offset - 0
    plcfsed.extend_from_slice(&0u32.to_le_bytes());

    plcfsed
}
