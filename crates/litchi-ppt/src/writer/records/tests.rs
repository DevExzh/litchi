use super::*;

#[test]
fn test_record_header() {
    let header = RecordHeader::new(0x0F, 0, record_type::SLIDE, 100);
    assert_eq!(header.version, 0x0F);
    assert_eq!(header.instance, 0);
    assert_eq!(header.total_size(), 108); // 8 byte header + 100 data
}

#[test]
fn test_record_builder() {
    let mut builder = RecordBuilder::new(0x00, 0, record_type::TEXT_CHARS_ATOM);
    builder.write_data(b"test");

    let record = builder.build().unwrap();
    assert!(record.len() >= 12); // At least 8 bytes header + 4 bytes data
}

#[test]
fn test_create_text_atom() {
    let atom = create_text_atom("Hello").unwrap();
    assert!(!atom.is_empty());
}

#[test]
fn test_record_header_total_size() {
    let header = RecordHeader::new(0x0F, 0, record_type::SLIDE, 100);
    assert_eq!(header.total_size(), 108); // 8 + 100
}

#[test]
fn test_record_header_write() {
    let header = RecordHeader::new(0x00, 1, record_type::DOCUMENT_ATOM, 48);
    let mut buf = Vec::new();
    header.write(&mut buf).unwrap();
    assert_eq!(buf.len(), 8);

    // Verify the bytes
    let ver_inst = u16::from_le_bytes([buf[0], buf[1]]);
    assert_eq!(ver_inst & 0x0F, 0x00); // version
    assert_eq!(ver_inst >> 4, 1); // instance
    let rec_type = u16::from_le_bytes([buf[2], buf[3]]);
    assert_eq!(rec_type, record_type::DOCUMENT_ATOM);
    let length = u32::from_le_bytes([buf[4], buf[5], buf[6], buf[7]]);
    assert_eq!(length, 48);
}

#[test]
fn test_record_builder_with_children() {
    let mut parent = RecordBuilder::new(0x0F, 0, record_type::DOCUMENT);

    let mut child1 = RecordBuilder::new(0x00, 0, record_type::DOCUMENT_ATOM);
    child1.write_data(&[0u8; 48]);
    parent.write_child(&child1.build().unwrap());

    let child2 = RecordBuilder::new(0x00, 0, record_type::END_DOCUMENT);
    parent.write_child(&child2.build().unwrap());

    let record = parent.build().unwrap();
    assert!(record.len() > 16); // Header + at least some data
}

#[test]
fn test_record_builder_empty() {
    let builder = RecordBuilder::new(0x00, 0, record_type::END_DOCUMENT);
    let record = builder.build().unwrap();
    assert_eq!(record.len(), 8); // Just header, no data
}

#[test]
fn test_create_document_atom() {
    let atom = create_document_atom(9144000, 6858000, 1, 0, 0).unwrap();
    assert!(!atom.is_empty());
    assert!(atom.len() >= 48); // Header + 40 bytes data
}

#[test]
fn test_create_main_master_container() {
    let ppdrawing = vec![0u8; 100]; // Mock PPDrawing data
    let container = create_main_master_container(&ppdrawing).unwrap();
    assert!(!container.is_empty());
    assert!(container.len() > 8);
}

#[test]
fn test_create_slide_container() {
    let container = create_slide_container(0x80000000, "").unwrap();
    assert!(!container.is_empty());
    assert!(container.len() > 8);
}

#[test]
fn test_wrap_dgg_into_ppdrawing_group() {
    let dgg_data = vec![1u8, 2, 3, 4, 5];
    let wrapped = wrap_dgg_into_ppdrawing_group(&dgg_data).unwrap();
    assert!(!wrapped.is_empty());
    assert!(wrapped.len() > dgg_data.len());
}

#[test]
fn test_wrap_dg_into_ppdrawing() {
    let dg_data = vec![1u8, 2, 3, 4, 5];
    let wrapped = wrap_dg_into_ppdrawing(&dg_data).unwrap();
    assert!(!wrapped.is_empty());
    assert!(wrapped.len() > dg_data.len());
}

#[test]
fn test_create_environment_minimal() {
    let env = create_environment_minimal().unwrap();
    assert!(!env.is_empty());
    // Should contain FontCollection and other required atoms
    assert!(env.len() > 100);

    let (environment, consumed) = crate::records::Record::parse(&env, 0).unwrap();
    assert_eq!(consumed, env.len());
    let collection = environment
        .find_child(crate::consts::RecordType::FontCollection)
        .unwrap();
    let fonts = crate::FontCollection::parse(collection).unwrap();
    assert_eq!(fonts.fonts.len(), 1);
    assert_eq!(fonts.fonts[0].name, "Arial");
    assert!(fonts.fonts[0].truetype);
}

#[test]
fn test_create_slide_list_with_text_slides_empty() {
    let slwt = create_slide_list_with_text_slides(&[]).unwrap();
    assert!(!slwt.is_empty());
    assert_eq!(slwt.len(), 8); // Just container header
}

#[test]
fn test_create_slide_list_with_text_slides_single() {
    let entries = vec![(1u32, 256u32)];
    let slwt = create_slide_list_with_text_slides(&entries).unwrap();
    assert!(!slwt.is_empty());
    assert!(slwt.len() > 8);
}

#[test]
fn test_create_slide_list_with_text_slides_multiple() {
    let entries = vec![(1, 256), (2, 257), (3, 258)];
    let slwt = create_slide_list_with_text_slides(&entries).unwrap();
    assert!(!slwt.is_empty());
    assert!(slwt.len() > 24); // More than just headers
}

#[test]
fn test_create_slide_list_with_text_notes() {
    let entries = vec![(4, 256), (5, 257)];
    let slwt = create_slide_list_with_text_notes(&entries).unwrap();
    assert!(!slwt.is_empty());
    assert!(slwt.len() > 8);
}

#[test]
fn test_create_docinfo_list_container() {
    let docinfo = create_docinfo_list_container_minimal().unwrap();
    assert!(!docinfo.is_empty());
    assert!(docinfo.len() > 50);
}

#[test]
fn test_create_end_document() {
    let end_doc = create_end_document().unwrap();
    assert_eq!(end_doc.len(), 8); // Just header

    // Verify record type
    let record_type = u16::from_le_bytes([end_doc[2], end_doc[3]]);
    assert_eq!(record_type, record_type::END_DOCUMENT);
}

#[test]
fn test_create_text_atom_unicode() {
    let atom = create_text_atom("Hello 世界 🌍").unwrap();
    assert!(!atom.is_empty());
    // Unicode text should be longer than ASCII
    assert!(atom.len() > 12);
}

#[test]
fn test_create_text_atom_empty() {
    let atom = create_text_atom("").unwrap();
    assert_eq!(atom.len(), 8); // Just header
}

#[test]
fn test_record_type_constants() {
    assert_eq!(record_type::DOCUMENT, 1000);
    assert_eq!(record_type::DOCUMENT_ATOM, 1001);
    assert_eq!(record_type::SLIDE, 1006);
    assert_eq!(record_type::SLIDE_ATOM, 1007);
    assert_eq!(record_type::END_DOCUMENT, 1002);
    assert_eq!(record_type::TEXT_CHARS_ATOM, 4000);
}

#[test]
fn test_record_builder_multiple_children() {
    let mut parent = RecordBuilder::new(0x0F, 0, record_type::DOCUMENT);

    for i in 0..5 {
        let mut child = RecordBuilder::new(0x00, i as u16, record_type::DOCUMENT_ATOM);
        child.write_data(&[i as u8; 10]);
        parent.write_child(&child.build().unwrap());
    }

    let record = parent.build().unwrap();
    assert!(record.len() > 50);
}

#[test]
fn test_create_environment_consistency() {
    let env1 = create_environment_minimal().unwrap();
    let env2 = create_environment_minimal().unwrap();
    assert_eq!(env1, env2);
}

#[test]
fn test_record_builder_with_large_data() {
    let mut builder = RecordBuilder::new(0x00, 0, record_type::TEXT_CHARS_ATOM);
    let large_data = vec![0xABu8; 10000];
    builder.write_data(&large_data);

    let record = builder.build().unwrap();
    assert_eq!(record.len(), 8 + large_data.len());
}
