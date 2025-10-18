// OLE and MTEF header structures
//
// Based on rtf2latex2e EQN_OLE_FILE_HDR and related structures

use zerocopy::FromBytes;

/// MTEF OLE file header (28 bytes)
#[derive(Debug, Clone, FromBytes)]
#[repr(C)]
#[allow(dead_code)]
pub struct OleFileHeader {
    pub cb_hdr: u16,        // Total header length = 28
    pub version: u32,       // Version number (0x00020000)
    pub format: u16,        // Clipboard format (0xC2D3)
    pub size: u32,          // "MTEF header + MTEF data" length
    pub reserved: [u32; 4], // Reserved fields
}

/// MTEF header (8 bytes)
/// Kept for future use when more detailed header parsing is needed
#[allow(dead_code)]
#[derive(Debug, Clone, FromBytes)]
#[repr(C)]
pub struct MtefHeader {
    pub signature: [u8; 4], // 0x28 0x04 0x6D 0x74
    pub major_ver: u8,      // 5
    pub minor_ver: u8,      // 1
    pub product_flag: u16,  // Product flags
}
