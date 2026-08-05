use super::*;

#[test]
fn test_sprm_type_extraction() {
    // sprmCFBold = 0x0835
    // Bits: 0000 1000 0011 0101
    // Type (bits 10-12) = 010 = 2 (CHP)
    assert_eq!(get_sprm_type(SPRM_C_F_BOLD), 2);

    // sprmPJc = 0x2403
    // Bits: 0010 0100 0000 0011
    // Type (bits 10-12) = 001 = 1 (PAP)
    assert_eq!(get_sprm_type(SPRM_P_JC), 1);

    // sprmTJc = 0x5400
    // Bits: 0101 0100 0000 0000
    // Type (bits 10-12) = 101 = 5 (TAP)
    assert_eq!(get_sprm_type(SPRM_T_JC), 5);
}

#[test]
fn table_sprm_constants_use_canonical_word_2000_opcodes() {
    assert_eq!(SPRM_T_DXA_LEFT, 0x9601);
    assert_eq!(SPRM_T_DEF_TABLE, 0xD608);
    assert_eq!(SPRM_T_TABLE_BORDERS, 0xD613);
    assert_eq!(SPRM_T_WIDTH_BEFORE, 0xF617);
    assert_eq!(SPRM_T_F_CANT_SPLIT, 0x3466);
    assert_eq!(SPRM_T_ISTD, 0x563A);
    assert_eq!(SPRM_T_CNF, 0xD66A);
    assert_eq!(SPRM_T_CELL_PADDING_STYLE, 0xD63E);
    assert_eq!(SPRM_T_CELL_BRC_TOP_STYLE, 0xD47F);
    assert_eq!(SPRM_T_CELL_SHD_STYLE, 0xD687);
    assert_eq!(SPRM_T_JC, 0x548A);
}

#[test]
fn test_sprm_size_code_extraction() {
    // sprmCFBold = 0x0835
    // Bits: 0000 1000 0011 0101
    // Size code (bits 13-15) = 000 = 0 (1 byte)
    assert_eq!(get_sprm_size_code(SPRM_C_F_BOLD), 0);

    // sprmCHps = 0x4A43
    // Bits: 0100 1010 0100 0011
    // Size code (bits 13-15) = 010 = 2 (2 bytes)
    assert_eq!(get_sprm_size_code(SPRM_C_HPS), 2);

    // sprmCPicLocation = 0x6A03
    // Bits: 0110 1010 0000 0011
    // Size code (bits 13-15) = 011 = 3 (4 bytes)
    assert_eq!(get_sprm_size_code(SPRM_C_PIC_LOCATION), 3);
}

#[test]
fn test_sprm_operation_extraction() {
    // sprmCFBold = 0x0835
    // Operation (bits 0-8) = 0x35 = 53
    assert_eq!(get_sprm_operation(SPRM_C_F_BOLD), 0x35);

    // sprmCHps = 0x4A43
    // Operation (bits 0-8) = 0x43 = 67
    assert_eq!(get_sprm_operation(SPRM_C_HPS), 0x43);
}

#[test]
fn test_sprm_special_flag() {
    // sprmCPlain = 0x2A33
    // Bit 9 = 1 (special)
    assert!(is_sprm_special(SPRM_C_PLAIN));

    // sprmCFBold = 0x0835
    // Bit 9 = 0 (not special)
    assert!(!is_sprm_special(SPRM_C_F_BOLD));
}

#[test]
fn current_high_paragraph_sprms_match_ms_doc() {
    assert_eq!(SPRM_P_BRC_TOP80, 0x6424);
    assert_eq!(SPRM_P_BRC_LEFT80, 0x6425);
    assert_eq!(SPRM_P_BRC_BOTTOM80, 0x6426);
    assert_eq!(SPRM_P_BRC_RIGHT80, 0x6427);
    assert_eq!(SPRM_P_BRC_BETWEEN80, 0x6428);
    assert_eq!(SPRM_P_BRC_BAR80, 0x6629);
    assert_eq!(SPRM_P_BRC_TOP, 0xC64E);
    assert_eq!(SPRM_P_BRC_LEFT, 0xC64F);
    assert_eq!(SPRM_P_BRC_BOTTOM, 0xC650);
    assert_eq!(SPRM_P_BRC_RIGHT, 0xC651);
    assert_eq!(SPRM_P_BRC_BETWEEN, 0xC652);
    assert_eq!(SPRM_P_BRC_BAR, 0xC653);
    assert_eq!(SPRM_P_DXC_RIGHT, 0x4455);
    assert_eq!(SPRM_P_DXC_LEFT, 0x4456);
    assert_eq!(SPRM_P_DXC_LEFT1, 0x4457);
    assert_eq!(SPRM_P_DYL_BEFORE, 0x4458);
    assert_eq!(SPRM_P_DYL_AFTER, 0x4459);
    assert_eq!(SPRM_P_F_OPEN_TCH, 0x245A);
    assert_eq!(SPRM_P_F_DYA_BEFORE_AUTO, 0x245B);
    assert_eq!(SPRM_P_F_DYA_AFTER_AUTO, 0x245C);
    assert_eq!(SPRM_P_DXA_RIGHT_2000, 0x845D);
    assert_eq!(SPRM_P_DXA_LEFT_2000, 0x845E);
    assert_eq!(SPRM_P_NEST_2000, 0x465F);
    assert_eq!(SPRM_P_DXA_LEFT1_2000, 0x8460);
    assert_eq!(SPRM_P_JC_LOGICAL, 0x2461);
    assert_eq!(SPRM_P_F_NO_ALLOW_OVERLAP, 0x2462);
    assert_eq!(SPRM_P_WALL, 0x2664);
    assert_eq!(SPRM_P_IPGP, 0x6465);
    assert_eq!(SPRM_P_CNF, 0xC666);
    assert_eq!(SPRM_P_RSID, 0x6467);
    assert_eq!(SPRM_P_ISTD_LIST_PERMUTE, 0xC669);
    assert_eq!(SPRM_P_TABLE_PROPS, 0x646B);
    assert_eq!(SPRM_P_T_ISTD_INFO, 0xC66C);
    assert_eq!(SPRM_P_F_CONTEXTUAL_SPACING, 0x246D);
    assert_eq!(SPRM_P_PROP_RMARK_CURRENT, 0xC66F);
    assert_eq!(SPRM_P_F_MIRROR_INDENTS, 0x2470);
    assert_eq!(SPRM_P_TTWO, 0x2471);
}
