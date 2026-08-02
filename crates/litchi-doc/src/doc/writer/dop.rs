//! DocumentProperties (DOP) generation for DOC files.

use crate::doc::parts::document_properties::DocumentProperties;

/// Generate document properties (DOP) with minimal fields required by Word
///
/// - `facing_pages`: sets DOP.fFacingPages (enables different odd/even pages in UI)
/// - `doc_grpf_ihdt`: section `grpfIhdt` bitfield to derive include header/footer flags
/// - `embed_factoids`: emits the Word 2002 `fEmbedFactoids` preservation flag
pub fn generate_dop(facing_pages: bool, doc_grpf_ihdt: u8, embed_factoids: bool) -> Vec<u8> {
    let has_header = (doc_grpf_ihdt & (0x01 | 0x02 | 0x10)) != 0;
    let has_footer = (doc_grpf_ihdt & (0x04 | 0x08 | 0x20)) != 0;
    DocumentProperties::writer_bytes(facing_pages, has_header, has_footer, embed_factoids)
}

/// Generate minimal document properties (no facing pages, no headers/footers)
#[inline]
pub fn generate_minimal_dop() -> Vec<u8> {
    generate_dop(false, 0, false)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn upgrades_dop_and_sets_factoid_preservation_only_when_requested() {
        let basic = DocumentProperties::parse_bytes(&generate_dop(false, 0, false)).unwrap();
        assert_eq!(basic.embeds_factoids(), None);

        let factoids = DocumentProperties::parse_bytes(&generate_dop(false, 0, true)).unwrap();
        assert_eq!(
            factoids.version(),
            crate::doc::DocumentPropertyVersion::Word2002
        );
        assert_eq!(factoids.embeds_factoids(), Some(true));
    }
}
