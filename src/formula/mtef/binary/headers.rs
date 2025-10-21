//! OLE and MTEF header structures
//!
//! This module defines the binary header structures for MTEF equation data.
//! Based on rtf2latex2e EQN_OLE_FILE_HDR and related structures.
//!
//! MTEF data embedded in OLE documents starts with a 28-byte OLE header,
//! followed by the MTEF header and equation data.

use zerocopy_derive::FromBytes as DeriveFromBytes;

/// MTEF OLE file header (28 bytes)
///
/// This header precedes MTEF equation data when embedded in OLE documents.
/// The structure matches the EQN_OLE_FILE_HDR from rtf2latex2e.
///
/// Note: This structure is kept for reference and future zerocopy-based parsing.
/// Current implementation uses manual parsing for better error handling.
#[allow(dead_code)] // Kept for reference and future zerocopy parsing
#[derive(Debug, Clone, DeriveFromBytes)]
#[repr(C)]
pub struct OleFileHeader {
    /// Total header length (should be 28)
    pub cb_hdr: u16,
    /// Version number (typically 0x00020000)
    pub version: u32,
    /// Clipboard format code (varies by application)
    pub format: u16,
    /// Length of MTEF header plus MTEF data
    pub size: u32,
    /// Reserved fields (should be zero)
    pub reserved: [u32; 4],
}

/// MTEF header (variable length, minimum 5 bytes)
///
/// The MTEF header starts with a signature and version information.
/// This structure is used for reference but actual parsing is done
/// manually in the parser to handle version differences.
///
/// Note: Kept for reference and documentation. Parser uses manual parsing
/// to handle different MTEF versions (1-5) with varying header layouts.
#[allow(dead_code)] // Kept for reference and documentation
#[derive(Debug, Clone, DeriveFromBytes)]
#[repr(C)]
pub struct MtefHeader {
    /// Signature bytes: "(", 0x04, "m", "t" or just version byte for headerless format
    pub signature: [u8; 4],
    /// MTEF major version (typically 5)
    pub major_ver: u8,
    /// Platform identifier (0=Mac, 1=Windows)
    pub minor_ver: u8,
    /// Product and version flags
    pub product_flag: u16,
}
