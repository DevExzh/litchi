use super::{FibBuilder, flags::FIB_BASE_VERSION};

#[test]
fn test_fib_generation() {
    let mut fib = FibBuilder::new();
    fib.set_main_text(0, 1000);

    let fib_bytes = fib.generate().unwrap();
    assert_eq!(fib_bytes.len(), 1248);
    assert_eq!(u16::from_le_bytes([fib_bytes[0], fib_bytes[1]]), 0xA5EC);
    assert_eq!(
        u16::from_le_bytes([fib_bytes[2], fib_bytes[3]]),
        FIB_BASE_VERSION
    );
    assert_eq!(
        u16::from_le_bytes([fib_bytes[1244], fib_bytes[1245]]),
        0x0101
    );
}

#[test]
fn test_fib_default() {
    let fib = FibBuilder::default();
    let fib_bytes = fib.generate().unwrap();
    assert_eq!(fib_bytes.len(), 1248);
}

#[test]
fn template_and_secondary_fib_page_are_serialized() {
    let mut fib = FibBuilder::new();
    fib.set_template(true);
    fib.set_next_fib_page(8);

    let bytes = fib.generate().unwrap();
    assert_eq!(u16::from_le_bytes(bytes[8..10].try_into().unwrap()), 8);
    let flags = u16::from_le_bytes(bytes[10..12].try_into().unwrap());
    assert_ne!(flags & 0x0001, 0);
    assert_eq!(flags & 0x0002, 0);
}

#[test]
fn test_fib_with_stylesheet() {
    let mut fib = FibBuilder::new();
    fib.set_main_text(0, 1000);
    fib.set_stshf(512, 256);

    let fib_bytes = fib.generate().unwrap();
    assert_eq!(fib_bytes.len(), 1248);
    assert_eq!(u16::from_le_bytes([fib_bytes[0], fib_bytes[1]]), 0xA5EC);
}

#[test]
fn test_fib_with_document_properties() {
    let mut fib = FibBuilder::new();
    fib.set_main_text(0, 1000);
    fib.set_dop(768, 128);

    let fib_bytes = fib.generate().unwrap();
    assert_eq!(fib_bytes.len(), 1248);
}

#[test]
fn test_fib_with_piece_table() {
    let mut fib = FibBuilder::new();
    fib.set_main_text(0, 1000);
    fib.set_clx(1024, 512);

    let fib_bytes = fib.generate().unwrap();
    assert_eq!(fib_bytes.len(), 1248);
}

#[test]
fn test_fib_with_font_table() {
    let mut fib = FibBuilder::new();
    fib.set_main_text(0, 1000);
    fib.set_sttbfffn(1536, 256);

    let fib_bytes = fib.generate().unwrap();
    assert_eq!(fib_bytes.len(), 1248);
}

#[test]
fn test_fib_with_section_table() {
    let mut fib = FibBuilder::new();
    fib.set_main_text(0, 1000);
    fib.set_plcfsed(2048, 128);

    let fib_bytes = fib.generate().unwrap();
    assert_eq!(fib_bytes.len(), 1248);
}

#[test]
fn test_fib_with_headers_footers() {
    let mut fib = FibBuilder::new();
    fib.set_main_text(0, 1000);
    fib.set_plcfhdd(2304, 256);

    let fib_bytes = fib.generate().unwrap();
    assert_eq!(fib_bytes.len(), 1248);
}

#[test]
fn test_fib_with_footnotes() {
    let mut fib = FibBuilder::new();
    fib.set_main_text(0, 1000);
    fib.set_plcffnd_ref(2816, 128);
    fib.set_plcffnd_txt(2944, 256);

    let fib_bytes = fib.generate().unwrap();
    assert_eq!(fib_bytes.len(), 1248);
}

#[test]
fn test_fib_with_endnotes() {
    let mut fib = FibBuilder::new();
    fib.set_main_text(0, 1000);
    fib.set_plcfend_ref(3584, 128);
    fib.set_plcfend_txt(3712, 256);

    let fib_bytes = fib.generate().unwrap();
    assert_eq!(fib_bytes.len(), 1248);
}

#[test]
fn test_fib_with_comments() {
    let mut fib = FibBuilder::new();
    fib.set_main_text(0, 1000);
    fib.set_ccp_atn(17);
    fib.set_plcfand_ref(3200, 38);
    fib.set_plcfand_txt(3238, 12);
    fib.set_grp_xst_atn_owners(3250, 14);
    fib.set_sttbf_atn_bkmk(3264, 18);
    fib.set_plcf_atn_bkf(3282, 12);
    fib.set_plcf_atn_bkl(3294, 8);
    fib.set_atrd_extra(3302, 18);

    let bytes = fib.generate().unwrap();
    assert_eq!(u32::from_le_bytes(bytes[92..96].try_into().unwrap()), 17);
    let field = |index: usize| {
        let offset = 154 + index * 8;
        (
            u32::from_le_bytes(bytes[offset..offset + 4].try_into().unwrap()),
            u32::from_le_bytes(bytes[offset + 4..offset + 8].try_into().unwrap()),
        )
    };
    assert_eq!(field(4), (3200, 38));
    assert_eq!(field(5), (3238, 12));
    assert_eq!(field(36), (3250, 14));
    assert_eq!(field(37), (3264, 18));
    assert_eq!(field(42), (3282, 12));
    assert_eq!(field(43), (3294, 8));
    assert_eq!(field(112), (3302, 18));
}

#[test]
fn test_fib_with_list_tables() {
    let mut fib = FibBuilder::new();
    fib.set_main_text(0, 1000);
    fib.set_plflst(4352, 512);
    fib.set_plflfo(4864, 256);

    let fib_bytes = fib.generate().unwrap();
    assert_eq!(fib_bytes.len(), 1248);
}

#[test]
fn test_fib_with_field_table() {
    let mut fib = FibBuilder::new();
    fib.set_main_text(0, 1000);
    fib.set_plcffld_mom(5120, 256);

    let fib_bytes = fib.generate().unwrap();
    assert_eq!(fib_bytes.len(), 1248);
}

#[test]
fn test_fib_with_bin_tables() {
    let mut fib = FibBuilder::new();
    fib.set_main_text(0, 1000);
    fib.set_plcfbte_chpx(5376, 128);
    fib.set_plcfbte_papx(5504, 128);

    let fib_bytes = fib.generate().unwrap();
    assert_eq!(fib_bytes.len(), 1248);
}
