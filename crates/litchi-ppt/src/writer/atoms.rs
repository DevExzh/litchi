//! PPT Atom record builders
//!
//! This module provides builders for various PPT atom records following the MS-PPT specification.

use super::records::{RecordBuilder, record_type};

/// DocumentAtom builder per MS-PPT section 2.4.2
///
/// Creates a 40-byte payload with:
/// - slideSize (8 bytes): Slide dimensions in master units
/// - notesSize (8 bytes): Notes dimensions in master units
/// - serverZoom (8 bytes): Zoom ratio
/// - notesMasterPersistIdRef (4 bytes)
/// - handoutMasterPersistIdRef (4 bytes)
/// - firstSlideNumber (2 bytes)
/// - slideSizeType (2 bytes)
/// - fSaveWithFonts, fOmitTitlePlace, fRightToLeft, fShowComments (4 bytes)
pub fn build_document_atom(slide_width_emu: u32, slide_height_emu: u32) -> Vec<u8> {
    // Convert EMU to master units (576 units per inch, 914400 EMUs per inch)
    let to_master = |emu: u32| -> u32 { ((emu as u64 * 576) / 914_400) as u32 };

    let slide_w = to_master(slide_width_emu);
    let slide_h = to_master(slide_height_emu);

    let mut data = Vec::with_capacity(40);

    // slideSize (8 bytes)
    data.extend_from_slice(&slide_w.to_le_bytes());
    data.extend_from_slice(&slide_h.to_le_bytes());

    // notesSize (8 bytes) - POI uses height first then width (portrait for notes)
    data.extend_from_slice(&slide_h.to_le_bytes());
    data.extend_from_slice(&slide_w.to_le_bytes());

    // serverZoom (8 bytes) - ratio 5:10 per POI empty.ppt
    data.extend_from_slice(&5u32.to_le_bytes());
    data.extend_from_slice(&10u32.to_le_bytes());

    // notesMasterPersistIdRef (4 bytes) - 0 = none
    data.extend_from_slice(&0u32.to_le_bytes());

    // handoutMasterPersistIdRef (4 bytes) - 0 = none
    data.extend_from_slice(&0u32.to_le_bytes());

    // firstSlideNumber (2 bytes)
    data.extend_from_slice(&1u16.to_le_bytes());

    // slideSizeType (2 bytes) - 0 = SS_OnScreen
    data.extend_from_slice(&0u16.to_le_bytes());

    // Flags (4 bytes)
    data.push(0u8); // fSaveWithFonts
    data.push(0u8); // fOmitTitlePlace
    data.push(0u8); // fRightToLeft
    data.push(1u8); // fShowComments

    // Build the record with recVer=1, recInstance=0, recType=RT_DocumentAtom
    let mut builder = RecordBuilder::new(0x01, 0, record_type::DOCUMENT_ATOM);
    builder.write_data(&data);
    builder.build().expect("DocumentAtom build")
}

/// SlideAtom builder per MS-PPT section 2.5.2
///
/// Creates a 24-byte payload with:
/// - geom (4 bytes): Slide layout geometry
/// - rgPlaceholderTypes (8 bytes): Placeholder type IDs
/// - masterIdRef (4 bytes): Master slide persist ID
/// - notesIdRef (4 bytes): Notes slide persist ID  
/// - slideFlags (2 bytes)
/// - unused (2 bytes)
pub fn build_slide_atom(master_id_ref: u32, notes_id_ref: u32) -> Vec<u8> {
    let mut data = Vec::with_capacity(24);

    // SSlideLayoutAtom (12 bytes)
    // geom = SL_Blank (0x10) - blank slide layout
    data.extend_from_slice(&0x0010u32.to_le_bytes());
    // rgPlaceholderTypes (8 bytes of placeholder type IDs)
    data.extend_from_slice(&[0u8; 8]);

    // masterIdRef (4 bytes)
    data.extend_from_slice(&master_id_ref.to_le_bytes());

    // notesIdRef (4 bytes)
    data.extend_from_slice(&notes_id_ref.to_le_bytes());

    // slideFlags (2 bytes): fMasterObjects | fMasterScheme | fMasterBackground = 0x07
    data.extend_from_slice(&0x0007u16.to_le_bytes());

    // unused (2 bytes)
    data.extend_from_slice(&0u16.to_le_bytes());

    let mut builder = RecordBuilder::new(0x02, 0, record_type::SLIDE_ATOM);
    builder.write_data(&data);
    builder.build().expect("SlideAtom build")
}

/// SlidePersistAtom builder per MS-PPT section 2.4.14.5
///
/// Creates a 20-byte payload with:
/// - persistIdRef (4 bytes)
/// - flags (4 bytes)
/// - numberTexts (4 bytes)
/// - slideId (4 bytes)
/// - reserved (4 bytes)
pub fn build_slide_persist_atom(persist_id_ref: u32, slide_id: u32, has_shapes: bool) -> Vec<u8> {
    let mut data = Vec::with_capacity(20);

    // persistIdRef
    data.extend_from_slice(&persist_id_ref.to_le_bytes());

    // flags: bit 2 = fNonOutlineData (has shapes other than placeholders)
    let flags: u32 = if has_shapes { 0x04 } else { 0x00 };
    data.extend_from_slice(&flags.to_le_bytes());

    // numberTexts
    data.extend_from_slice(&0u32.to_le_bytes());

    // slideId
    data.extend_from_slice(&slide_id.to_le_bytes());

    // reserved
    data.extend_from_slice(&0u32.to_le_bytes());

    let mut builder = RecordBuilder::new(0x00, 0, record_type::SLIDE_PERSIST_ATOM);
    builder.write_data(&data);
    builder.build().expect("SlidePersistAtom build")
}

/// ColorSchemeAtom builder per MS-PPT section 2.4.17
///
/// Creates a 32-byte payload with 8 RGBX colors
pub fn build_color_scheme_atom(instance: u16, colors: &[u32; 8]) -> Vec<u8> {
    let mut data = Vec::with_capacity(32);
    for color in colors {
        data.extend_from_slice(&color.to_le_bytes());
    }

    let mut builder = RecordBuilder::new(0x00, instance, record_type::COLOR_SCHEME_ATOM);
    builder.write_data(&data);
    builder.build().expect("ColorSchemeAtom build")
}

/// Default color scheme (white background, black text)
pub fn default_color_scheme() -> [u32; 8] {
    [
        0x00FFFFFF, // background
        0x00000000, // text and lines
        0x00808080, // shadows
        0x00000000, // title text
        0x00E3E0BB, // fills
        0x00993333, // accent
        0x00999900, // accent and hyperlink
        0x000099CC, // accent and followed hyperlink
    ]
}

/// EndDocumentAtom builder per MS-PPT section 2.4.13
pub fn build_end_document_atom() -> Vec<u8> {
    let builder = RecordBuilder::new(0x00, 0, record_type::END_DOCUMENT);
    builder.build().expect("EndDocumentAtom build")
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_document_atom_size() {
        let atom = build_document_atom(9144000, 6858000);
        // 8 bytes header + 40 bytes data = 48 bytes
        assert_eq!(atom.len(), 48);
    }

    #[test]
    fn test_document_atom_different_sizes() {
        // Standard 4:3
        let atom1 = build_document_atom(9144000, 6858000);
        assert_eq!(atom1.len(), 48);

        // Widescreen 16:9
        let atom2 = build_document_atom(12192000, 6858000);
        assert_eq!(atom2.len(), 48);

        // Verify record type in header
        let record_type = u16::from_le_bytes([atom1[2], atom1[3]]);
        assert_eq!(record_type, record_type::DOCUMENT_ATOM);
    }

    #[test]
    fn test_document_atom_content() {
        let atom = build_document_atom(9144000, 6858000);

        // Verify header version (should be 1 for DocumentAtom)
        let ver_inst = u16::from_le_bytes([atom[0], atom[1]]);
        let version = ver_inst & 0x0F;
        assert_eq!(version, 0x01);

        // Verify record type
        let rec_type = u16::from_le_bytes([atom[2], atom[3]]);
        assert_eq!(rec_type, record_type::DOCUMENT_ATOM);

        // Verify length (40 bytes)
        let length = u32::from_le_bytes([atom[4], atom[5], atom[6], atom[7]]);
        assert_eq!(length, 40);
    }

    #[test]
    fn test_slide_atom_size() {
        let atom = build_slide_atom(0x80000000, 0);
        // 8 bytes header + 24 bytes data = 32 bytes
        assert_eq!(atom.len(), 32);
    }

    #[test]
    fn test_slide_atom_variations() {
        // With master reference only
        let atom1 = build_slide_atom(0x80000000, 0);
        assert_eq!(atom1.len(), 32);

        // With notes reference
        let atom2 = build_slide_atom(0x80000000, 257);
        assert_eq!(atom2.len(), 32);

        // Verify header version (should be 2 for SlideAtom)
        let ver_inst = u16::from_le_bytes([atom1[0], atom1[1]]);
        let version = ver_inst & 0x0F;
        assert_eq!(version, 0x02);
    }

    #[test]
    fn test_slide_atom_content_structure() {
        let atom = build_slide_atom(0x80000000, 0);

        // Record type
        let rec_type = u16::from_le_bytes([atom[2], atom[3]]);
        assert_eq!(rec_type, record_type::SLIDE_ATOM);

        // Data length
        let length = u32::from_le_bytes([atom[4], atom[5], atom[6], atom[7]]);
        assert_eq!(length, 24);

        // SSlideLayoutAtom at offset 8
        let geom = u32::from_le_bytes([atom[8], atom[9], atom[10], atom[11]]);
        assert_eq!(geom, 0x0010); // SL_Blank

        // masterIdRef at offset 20
        let master_id = u32::from_le_bytes([atom[20], atom[21], atom[22], atom[23]]);
        assert_eq!(master_id, 0x80000000);

        // flags at offset 28
        let flags = u16::from_le_bytes([atom[28], atom[29]]);
        assert_eq!(flags, 0x0007); // fMasterObjects | fMasterScheme | fMasterBackground
    }

    #[test]
    fn test_slide_persist_atom_size() {
        let atom = build_slide_persist_atom(1, 256, true);
        // 8 bytes header + 20 bytes data = 28 bytes
        assert_eq!(atom.len(), 28);
    }

    #[test]
    fn test_slide_persist_atom_with_shapes() {
        let atom = build_slide_persist_atom(1, 256, true);

        // Record type
        let rec_type = u16::from_le_bytes([atom[2], atom[3]]);
        assert_eq!(rec_type, record_type::SLIDE_PERSIST_ATOM);

        // Check flags (should have fNonOutlineData set)
        let flags = u32::from_le_bytes([atom[12], atom[13], atom[14], atom[15]]);
        assert_eq!(flags, 0x04);

        // Check slideId
        let slide_id = u32::from_le_bytes([atom[20], atom[21], atom[22], atom[23]]);
        assert_eq!(slide_id, 256);
    }

    #[test]
    fn test_slide_persist_atom_without_shapes() {
        let atom = build_slide_persist_atom(2, 257, false);

        // Check flags (should be 0)
        let flags = u32::from_le_bytes([atom[12], atom[13], atom[14], atom[15]]);
        assert_eq!(flags, 0x00);
    }

    #[test]
    fn test_color_scheme_atom() {
        let colors = default_color_scheme();
        let atom = build_color_scheme_atom(1, &colors);

        // 8 bytes header + 32 bytes data (8 colors * 4 bytes)
        assert_eq!(atom.len(), 40);

        // Record type
        let rec_type = u16::from_le_bytes([atom[2], atom[3]]);
        assert_eq!(rec_type, record_type::COLOR_SCHEME_ATOM);

        // Instance should be 1
        let ver_inst = u16::from_le_bytes([atom[0], atom[1]]);
        let instance = ver_inst >> 4;
        assert_eq!(instance, 1);
    }

    #[test]
    fn test_color_scheme_atom_different_instances() {
        let colors = default_color_scheme();

        let atom0 = build_color_scheme_atom(0, &colors);
        let atom1 = build_color_scheme_atom(1, &colors);
        let atom2 = build_color_scheme_atom(2, &colors);

        // Same size, different instance values
        assert_eq!(atom0.len(), atom1.len());
        assert_eq!(atom1.len(), atom2.len());

        // Check instance values in headers
        let inst0 = u16::from_le_bytes([atom0[0], atom0[1]]) >> 4;
        let inst1 = u16::from_le_bytes([atom1[0], atom1[1]]) >> 4;
        let inst2 = u16::from_le_bytes([atom2[0], atom2[1]]) >> 4;

        assert_eq!(inst0, 0);
        assert_eq!(inst1, 1);
        assert_eq!(inst2, 2);
    }

    #[test]
    fn test_default_color_scheme_values() {
        let colors = default_color_scheme();
        assert_eq!(colors.len(), 8);

        // Background color (white)
        assert_eq!(colors[0], 0x00FFFFFF);

        // Text and lines color (black)
        assert_eq!(colors[1], 0x00000000);

        // Shadows (gray)
        assert_eq!(colors[2], 0x00808080);

        // Title text (black)
        assert_eq!(colors[3], 0x00000000);
    }

    #[test]
    fn test_end_document_atom() {
        let atom = build_end_document_atom();
        assert_eq!(atom.len(), 8); // Just header, no data

        // Record type
        let rec_type = u16::from_le_bytes([atom[2], atom[3]]);
        assert_eq!(rec_type, record_type::END_DOCUMENT);

        // Length should be 0
        let length = u32::from_le_bytes([atom[4], atom[5], atom[6], atom[7]]);
        assert_eq!(length, 0);
    }

    #[test]
    fn test_document_atom_conversion() {
        // EMU to master units conversion
        // 914400 EMUs = 1 inch = 576 master units
        let atom = build_document_atom(914400, 914400);

        // After conversion, should be 576x576 master units
        let slide_w = u32::from_le_bytes([atom[8], atom[9], atom[10], atom[11]]);
        let slide_h = u32::from_le_bytes([atom[12], atom[13], atom[14], atom[15]]);

        assert_eq!(slide_w, 576);
        assert_eq!(slide_h, 576);
    }

    #[test]
    fn test_slide_persist_atom_multiple() {
        let ids = vec![(1, 256), (2, 257), (3, 258), (4, 259)];

        for (persist_id, slide_id) in ids {
            let atom = build_slide_persist_atom(persist_id, slide_id, true);

            // Verify persistIdRef
            let stored_persist_id = u32::from_le_bytes([atom[8], atom[9], atom[10], atom[11]]);
            assert_eq!(stored_persist_id, persist_id);

            // Verify slideId
            let stored_slide_id = u32::from_le_bytes([atom[20], atom[21], atom[22], atom[23]]);
            assert_eq!(stored_slide_id, slide_id);
        }
    }
}
