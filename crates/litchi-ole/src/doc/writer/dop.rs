//! DocumentProperties (DOP) generation for DOC files.

use crate::doc::parts::document_properties::DocumentProperties;

/// Generate document properties (DOP) with minimal fields required by Word
///
/// - `facing_pages`: sets DOP.fFacingPages (enables different odd/even pages in UI)
/// - `doc_grpf_ihdt`: section `grpfIhdt` bitfield to derive include header/footer flags
pub fn generate_dop(facing_pages: bool, doc_grpf_ihdt: u8) -> Vec<u8> {
    let has_header = (doc_grpf_ihdt & (0x01 | 0x02 | 0x10)) != 0;
    let has_footer = (doc_grpf_ihdt & (0x04 | 0x08 | 0x20)) != 0;
    DocumentProperties::word97_writer_bytes(facing_pages, has_header, has_footer)
}

/// Generate minimal document properties (no facing pages, no headers/footers)
#[inline]
pub fn generate_minimal_dop() -> Vec<u8> {
    generate_dop(false, 0)
}
