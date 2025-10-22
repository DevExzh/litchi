/// SPRM (Single Property Modifier) operation constants and utilities.
///
/// This module provides comprehensive SPRM definitions based on Apache POI's implementation.
/// SPRMs are used in DOC and PPT formats to modify character, paragraph, table, and section properties.
///
/// Reference: Apache POI's hwpf/sprm package and usermodel/*Properties.java
///
/// # SPRM Structure
///
/// A SPRM consists of:
/// - **Opcode** (2 bytes): Encodes the operation type and size
///   - Bits 0-8: Operation code
///   - Bits 9: Special flag
///   - Bits 10-12: Property type (CHP=2, PAP=1, TAP=5, etc.)
///   - Bits 13-15: Size code (determines operand size)
/// - **Operand** (variable): The data for the operation
///
/// # Size Codes
///
/// - 0, 1: 1-byte operand
/// - 2, 4, 5: 2-byte operand
/// - 3: 4-byte operand
/// - 6: Variable length (size in first byte or word)
/// - 7: 3-byte operand
// CHP (Character Properties) SPRM opcodes
// Based on Apache POI's CharacterProperties and CharacterSprmUncompressor

/// sprmCFRMarkDel - Mark deleted revision (operation 0x00)
pub const SPRM_C_F_RMARK_DEL: u16 = 0x0800;

/// sprmCFRMark - Mark revision (operation 0x01)
pub const SPRM_C_F_RMARK: u16 = 0x0801;

/// sprmCFFldVanish - Field vanish flag (operation 0x02)
pub const SPRM_C_F_FLD_VANISH: u16 = 0x0802;

/// sprmCPicLocation - Picture/object location in Data stream (operation 0x03)
/// A signed 32-bit integer specifying position in Data Stream
pub const SPRM_C_PIC_LOCATION: u16 = 0x6A03;

/// sprmCIbstRMark - Revision mark author index (operation 0x04)
pub const SPRM_C_IBST_RMARK: u16 = 0x4804;

/// sprmCDttmRMark - Revision mark date/time (operation 0x05)
pub const SPRM_C_DTTM_RMARK: u16 = 0x6805;

/// sprmCFData - Data flag (operation 0x06)
pub const SPRM_C_F_DATA: u16 = 0x0806;

/// sprmCIdslRMark - Revision mark ID (operation 0x07)
pub const SPRM_C_IDSL_RMARK: u16 = 0x4807;

/// sprmCChs - Complex character set info (operation 0x08)
pub const SPRM_C_CHS: u16 = 0x6A08;

/// sprmCSymbol - Symbol character (operation 0x09)
pub const SPRM_C_SYMBOL: u16 = 0x6A09;

/// sprmCFOle2 - OLE2 embedded object flag (operation 0x0A)
pub const SPRM_C_F_OLE2: u16 = 0x080A;

/// sprmCIcoHighlight - Highlight color (operation 0x0C)
pub const SPRM_C_ICO_HIGHLIGHT: u16 = 0x2A0C;

/// sprmCObjLocation - Object location (operation 0x0E)
pub const SPRM_C_OBJ_LOCATION: u16 = 0x680E;

/// sprmCFWebHidden - Web hidden flag (operation 0x11)
pub const SPRM_C_F_WEB_HIDDEN: u16 = 0x0811;

/// sprmCRsidProp - Revision save ID property (operation 0x15)
pub const SPRM_C_RSID_PROP: u16 = 0x6815;

/// sprmCRsidText - Revision save ID text (operation 0x16)
pub const SPRM_C_RSID_TEXT: u16 = 0x6816;

/// sprmCRsidRMDel - Revision save ID deletion (operation 0x17)
pub const SPRM_C_RSID_RM_DEL: u16 = 0x6817;

/// sprmCFSpecVanish - Special vanish flag (operation 0x18)
pub const SPRM_C_F_SPEC_VANISH: u16 = 0x0818;

/// sprmCFMathPr - Math properties flag (operation 0x1A)
pub const SPRM_C_F_MATH_PR: u16 = 0x081A;

/// sprmCIstd - Style index (operation 0x30)
pub const SPRM_C_ISTD: u16 = 0x4A30;

/// sprmCIstdPermute - Style permutation (operation 0x31)
pub const SPRM_C_ISTD_PERMUTE: u16 = 0xCA31;

/// sprmCDefault - Reset to default (operation 0x32)
pub const SPRM_C_DEFAULT: u16 = 0x2A32;

/// sprmCPlain - Plain text (reset all formatting) (operation 0x33)
pub const SPRM_C_PLAIN: u16 = 0x2A33;

/// sprmCKcd - Keyboard code (operation 0x34)
pub const SPRM_C_KCD: u16 = 0x4A34;

/// sprmCFBold - Bold (operation 0x35)
pub const SPRM_C_F_BOLD: u16 = 0x0835;

/// sprmCFItalic - Italic (operation 0x36)
pub const SPRM_C_F_ITALIC: u16 = 0x0836;

/// sprmCFStrike - Strikethrough (operation 0x37)
pub const SPRM_C_F_STRIKE: u16 = 0x0837;

/// sprmCFOutline - Outline (operation 0x38)
pub const SPRM_C_F_OUTLINE: u16 = 0x0838;

/// sprmCFShadow - Shadow (operation 0x39)
pub const SPRM_C_F_SHADOW: u16 = 0x0839;

/// sprmCFSmallCaps - Small caps (operation 0x3A)
pub const SPRM_C_F_SMALL_CAPS: u16 = 0x083A;

/// sprmCFCaps - All caps (operation 0x3B)
pub const SPRM_C_F_CAPS: u16 = 0x083B;

/// sprmCFVanish - Hidden text (operation 0x3C)
pub const SPRM_C_F_VANISH: u16 = 0x083C;

/// sprmCFtcDefault - Default font (operation 0x3D)
pub const SPRM_C_FTC_DEFAULT: u16 = 0x4A3D;

/// sprmCKul - Underline style (operation 0x3E)
pub const SPRM_C_KUL: u16 = 0x2A3E;

/// sprmCSizePos - Size and position (operation 0x3F)
pub const SPRM_C_SIZE_POS: u16 = 0x8840; // Note: POI shows this as 0x8840 in code

/// sprmCDxaSpace - Character spacing (operation 0x40)
pub const SPRM_C_DXA_SPACE: u16 = 0x8840;

/// sprmCLid - Language ID (operation 0x41)
pub const SPRM_C_LID: u16 = 0x4A41;

/// sprmCIco - Text color index (operation 0x42)
pub const SPRM_C_ICO: u16 = 0x2A42;

/// sprmCHps - Font size in half-points (operation 0x43)
pub const SPRM_C_HPS: u16 = 0x4A43;

/// sprmCHpsInc - Font size increment (operation 0x44)
pub const SPRM_C_HPS_INC: u16 = 0x2A44;

/// sprmCHpsPos - Superscript/subscript position (operation 0x45)
pub const SPRM_C_HPS_POS: u16 = 0x4845;

/// sprmCHpsPosAdj - Position adjustment (operation 0x46)
pub const SPRM_C_HPS_POS_ADJ: u16 = 0x2A46;

/// sprmCMajority - Majority formatting (operation 0x47)
pub const SPRM_C_MAJORITY: u16 = 0xCA47;

/// sprmCIss - Superscript/subscript (operation 0x48)
pub const SPRM_C_ISS: u16 = 0x2A48;

/// sprmCHpsNew50 - Font size (Word 6.0) (operation 0x49)
pub const SPRM_C_HPS_NEW50: u16 = 0x484B; // Note: POI code shows 0x49 in switch but constant is different

/// sprmCHpsInc1 - Font size increment v1 (operation 0x4A)
pub const SPRM_C_HPS_INC1: u16 = 0x484A;

/// sprmCHpsKern - Kerning (operation 0x4B)
pub const SPRM_C_HPS_KERN: u16 = 0x484B;

/// sprmCMajority50 - Majority formatting (Word 6.0) (operation 0x4C)
pub const SPRM_C_MAJORITY50: u16 = 0xCA4C;

/// sprmCHpsMul - Font size multiplier (operation 0x4D)
pub const SPRM_C_HPS_MUL: u16 = 0x4A4D;

/// sprmCHresi - Hyphenation (operation 0x4E)
pub const SPRM_C_HRESI: u16 = 0x484E;

/// sprmCRgFtc0 - Font for ASCII characters (operation 0x4F)
pub const SPRM_C_RG_FTC0: u16 = 0x4A4F;

/// sprmCRgFtc1 - Font for Far East characters (operation 0x50)
pub const SPRM_C_RG_FTC1: u16 = 0x4A50;

/// sprmCRgFtc2 - Font for non-Far East characters (operation 0x51)
pub const SPRM_C_RG_FTC2: u16 = 0x4A51;

/// sprmCCharScale - Character scale (operation 0x52)
pub const SPRM_C_CHAR_SCALE: u16 = 0x4852;

/// sprmCFDStrike - Double strikethrough (operation 0x53)
pub const SPRM_C_F_D_STRIKE: u16 = 0x2A53;

/// sprmCFImprint - Imprint (operation 0x54)
pub const SPRM_C_F_IMPRINT: u16 = 0x0854;

/// sprmCFSpec - Special character flag (operation 0x55)
pub const SPRM_C_F_SPEC: u16 = 0x0855;

/// sprmCFObj - Object flag (operation 0x56)
pub const SPRM_C_F_OBJ: u16 = 0x0856;

/// sprmCPropRMark - Property revision mark (operation 0x57)
pub const SPRM_C_PROP_RMARK: u16 = 0xCA57;

/// sprmCFEmboss - Emboss (operation 0x58)
pub const SPRM_C_F_EMBOSS: u16 = 0x0858;

/// sprmCSfxtText - Text animation (operation 0x59)
pub const SPRM_C_SFXT_TEXT: u16 = 0x2859;

/// sprmCFBiDi - Bi-directional text (operation 0x5A)
pub const SPRM_C_F_BI_DI: u16 = 0x085A;

/// sprmCFDiacColor - Diacritic color (operation 0x5B)
pub const SPRM_C_F_DIAC_COLOR: u16 = 0x085B;

/// sprmCFBoldBi - Bold bi-directional (operation 0x5C)
pub const SPRM_C_F_BOLD_BI: u16 = 0x085C;

/// sprmCFItalicBi - Italic bi-directional (operation 0x5D)
pub const SPRM_C_F_ITALIC_BI: u16 = 0x085D;

/// sprmCFtcBi - Bi-directional font (operation 0x5E)
pub const SPRM_C_FTC_BI: u16 = 0x4A5E;

/// sprmCLidBi - Bi-directional language ID (operation 0x5F)
pub const SPRM_C_LID_BI: u16 = 0x4A5F;

/// sprmCIcoBi - Bi-directional text color (operation 0x60)
pub const SPRM_C_ICO_BI: u16 = 0x4A60;

/// sprmCHpsBi - Bi-directional font size (operation 0x61)
pub const SPRM_C_HPS_BI: u16 = 0x4A61;

/// sprmCDispFldRMark - Display field revision mark (operation 0x62)
pub const SPRM_C_DISP_FLD_RMARK: u16 = 0xCA62;

/// sprmCIbstRMarkDel - Revision mark deletion author (operation 0x63)
pub const SPRM_C_IBST_RMARK_DEL: u16 = 0x4863;

/// sprmCDttmRMarkDel - Revision mark deletion date/time (operation 0x64)
pub const SPRM_C_DTTM_RMARK_DEL: u16 = 0x6864;

/// sprmCBrc - Border (operation 0x65)
pub const SPRM_C_BRC: u16 = 0x6865;

/// sprmCShd80 - Shading (Word 97-2000) (operation 0x66)
pub const SPRM_C_SHD80: u16 = 0x4866;

/// sprmCIdslRMarkDel - Deletion revision ID (operation 0x67)
pub const SPRM_C_IDSL_RMARK_DEL: u16 = 0x4867;

/// sprmCFUsePgsuSettings - Use page setup settings (operation 0x68)
pub const SPRM_C_F_USE_PGSU_SETTINGS: u16 = 0x0868;

/// sprmCCpg - Code page (operation 0x6B)
pub const SPRM_C_CPG: u16 = 0x486B;

/// sprmCRgLid0 - Language ID for ASCII (operation 0x6D)
pub const SPRM_C_RG_LID0: u16 = 0x486D;

/// sprmCRgLid1 - Language ID for Far East (operation 0x6E)
pub const SPRM_C_RG_LID1: u16 = 0x486E;

/// sprmCIdctHint - Font hint (operation 0x6F)
pub const SPRM_C_IDCT_HINT: u16 = 0x286F;

/// sprmCCv - Color value (RGB) (operation 0x70)
pub const SPRM_C_CV: u16 = 0x6870;

/// sprmCShd - Shading (Word 2002+) (operation 0x71)
pub const SPRM_C_SHD: u16 = 0xCA71;

/// sprmCBrc80 - Border (Word 97-2000) (operation 0x72)
pub const SPRM_C_BRC80: u16 = 0xCA72;

/// sprmCRgLid0_80 - Language ID v80 (operation 0x73)
pub const SPRM_C_RG_LID0_80: u16 = 0x4873;

/// sprmCRgLid1_80 - Language ID v80 (operation 0x74)
pub const SPRM_C_RG_LID1_80: u16 = 0x4874;

/// sprmCFNoProof - No proofing (operation 0x75)
pub const SPRM_C_F_NO_PROOF: u16 = 0x0875;

// PAP (Paragraph Properties) SPRM opcodes
// Based on Apache POI's ParagraphProperties and ParagraphSprmUncompressor

/// sprmPIstd - Paragraph style (operation 0x00)
pub const SPRM_P_ISTD: u16 = 0x4600;

/// sprmPIstdPermute - Style permutation (operation 0x01)
pub const SPRM_P_ISTD_PERMUTE: u16 = 0xC601;

/// sprmPIncLvl - Increment outline level (operation 0x02)
pub const SPRM_P_INC_LVL: u16 = 0x2602;

/// sprmPJc - Paragraph justification (operation 0x03)
pub const SPRM_P_JC: u16 = 0x2403;

/// sprmPFSideBySide - Side-by-side paragraphs (operation 0x04)
pub const SPRM_P_F_SIDE_BY_SIDE: u16 = 0x2404;

/// sprmPFKeep - Keep paragraph intact (operation 0x05)
pub const SPRM_P_F_KEEP: u16 = 0x2405;

/// sprmPFKeepFollow - Keep with next (operation 0x06)
pub const SPRM_P_F_KEEP_FOLLOW: u16 = 0x2406;

/// sprmPFPageBreakBefore - Page break before (operation 0x07)
pub const SPRM_P_F_PAGE_BREAK_BEFORE: u16 = 0x2407;

/// sprmPBrcl - Border location (operation 0x08)
pub const SPRM_P_BRCL: u16 = 0x2408;

/// sprmPBrcp - Border position (operation 0x09)
pub const SPRM_P_BRCP: u16 = 0x2409;

/// sprmPIlvl - List level (operation 0x0A)
pub const SPRM_P_ILVL: u16 = 0x260A;

/// sprmPIlfo - List format override (operation 0x0B)
pub const SPRM_P_ILFO: u16 = 0x460B;

/// sprmPFNoLineNumb - No line numbering (operation 0x0C)
pub const SPRM_P_F_NO_LINE_NUMB: u16 = 0x240C;

/// sprmPChgTabsPapx - Tab stops (operation 0x0D)
pub const SPRM_P_CHG_TABS_PAPX: u16 = 0xC60D;

/// sprmPDxaRight - Right indent (operation 0x0E)
pub const SPRM_P_DXA_RIGHT: u16 = 0x840E;

/// sprmPDxaLeft - Left indent (operation 0x0F)
pub const SPRM_P_DXA_LEFT: u16 = 0x840F;

/// sprmPNest - Nested indent (operation 0x10)
pub const SPRM_P_NEST: u16 = 0x4610;

/// sprmPDxaLeft1 - First line indent (operation 0x11)
pub const SPRM_P_DXA_LEFT1: u16 = 0x8411;

/// sprmPDyaLine - Line spacing (operation 0x12)
pub const SPRM_P_DYA_LINE: u16 = 0x6412;

/// sprmPDyaBefore - Space before (operation 0x13)
pub const SPRM_P_DYA_BEFORE: u16 = 0xA413;

/// sprmPDyaAfter - Space after (operation 0x14)
pub const SPRM_P_DYA_AFTER: u16 = 0xA414;

/// sprmPChgTabs - Change tabs (operation 0x15)
pub const SPRM_P_CHG_TABS: u16 = 0xC615;

/// sprmPFInTable - In table flag (operation 0x16)
pub const SPRM_P_F_IN_TABLE: u16 = 0x2416;

/// sprmPFTtp - Table row end (operation 0x17)
pub const SPRM_P_F_TTP: u16 = 0x2417;

/// sprmPDxaAbs - Absolute horizontal position (operation 0x18)
pub const SPRM_P_DXA_ABS: u16 = 0x8418;

/// sprmPDyaAbs - Absolute vertical position (operation 0x19)
pub const SPRM_P_DYA_ABS: u16 = 0x8419;

/// sprmPDxaWidth - Absolute width (operation 0x1A)
pub const SPRM_P_DXA_WIDTH: u16 = 0x841A;

/// sprmPPc - Positioning code (operation 0x1B)
pub const SPRM_P_PC: u16 = 0x261B;

/// sprmPBrcTop10 - Top border (Word 6.0) (operation 0x1C)
pub const SPRM_P_BRC_TOP10: u16 = 0x461C;

/// sprmPBrcLeft10 - Left border (Word 6.0) (operation 0x1D)
pub const SPRM_P_BRC_LEFT10: u16 = 0x461D;

/// sprmPBrcBottom10 - Bottom border (Word 6.0) (operation 0x1E)
pub const SPRM_P_BRC_BOTTOM10: u16 = 0x461E;

/// sprmPBrcRight10 - Right border (Word 6.0) (operation 0x1F)
pub const SPRM_P_BRC_RIGHT10: u16 = 0x461F;

/// sprmPBrcBetween10 - Between border (Word 6.0) (operation 0x20)
pub const SPRM_P_BRC_BETWEEN10: u16 = 0x4620;

/// sprmPBrcBar10 - Bar border (Word 6.0) (operation 0x21)
pub const SPRM_P_BRC_BAR10: u16 = 0x4621;

/// sprmPDxaFromText10 - Distance from text (Word 6.0) (operation 0x22)
pub const SPRM_P_DXA_FROM_TEXT10: u16 = 0x4622;

/// sprmPWr - Text wrapping (operation 0x23)
pub const SPRM_P_WR: u16 = 0x2423;

/// sprmPBrcTop - Top border (operation 0x24)
pub const SPRM_P_BRC_TOP: u16 = 0x6424;

/// sprmPBrcLeft - Left border (operation 0x25)
pub const SPRM_P_BRC_LEFT: u16 = 0x6425;

/// sprmPBrcBottom - Bottom border (operation 0x26)
pub const SPRM_P_BRC_BOTTOM: u16 = 0x6426;

/// sprmPBrcRight - Right border (operation 0x27)
pub const SPRM_P_BRC_RIGHT: u16 = 0x6427;

/// sprmPBrcBetween - Between border (operation 0x28)
pub const SPRM_P_BRC_BETWEEN: u16 = 0x6428;

/// sprmPBrcBar - Bar border (operation 0x29)
pub const SPRM_P_BRC_BAR: u16 = 0x6429;

/// sprmPFNoAutoHyph - No auto hyphenation (operation 0x2A)
pub const SPRM_P_F_NO_AUTO_HYPH: u16 = 0x242A;

/// sprmPWHeightAbs - Absolute row height (operation 0x2B)
pub const SPRM_P_W_HEIGHT_ABS: u16 = 0x442B;

/// sprmPDcs - Drop cap (operation 0x2C)
pub const SPRM_P_DCS: u16 = 0x442C;

/// sprmPShd80 - Shading (Word 97-2000) (operation 0x2D)
pub const SPRM_P_SHD80: u16 = 0x442D;

/// sprmPDyaFromText - Vertical distance from text (operation 0x2E)
pub const SPRM_P_DYA_FROM_TEXT: u16 = 0x842E;

/// sprmPDxaFromText - Horizontal distance from text (operation 0x2F)
pub const SPRM_P_DXA_FROM_TEXT: u16 = 0x842F;

/// sprmPFLocked - Locked paragraph (operation 0x30)
pub const SPRM_P_F_LOCKED: u16 = 0x2430;

/// sprmPFWidowControl - Widow/orphan control (operation 0x31)
pub const SPRM_P_F_WIDOW_CONTROL: u16 = 0x2431;

/// sprmPFKinsoku - Kinsoku (operation 0x33)
pub const SPRM_P_F_KINSOKU: u16 = 0x2433;

/// sprmPFWordWrap - Word wrap (operation 0x34)
pub const SPRM_P_F_WORD_WRAP: u16 = 0x2434;

/// sprmPFOverflowPunct - Overflow punctuation (operation 0x35)
pub const SPRM_P_F_OVERFLOW_PUNCT: u16 = 0x2435;

/// sprmPFTopLinePunct - Top line punctuation (operation 0x36)
pub const SPRM_P_F_TOP_LINE_PUNCT: u16 = 0x2436;

/// sprmPFAutoSpaceDE - Auto space DE (operation 0x37)
pub const SPRM_P_F_AUTO_SPACE_DE: u16 = 0x2437;

/// sprmPFAutoSpaceDN - Auto space DN (operation 0x38)
pub const SPRM_P_F_AUTO_SPACE_DN: u16 = 0x2438;

/// sprmPWAlignFont - Font alignment (operation 0x39)
pub const SPRM_P_W_ALIGN_FONT: u16 = 0x4439;

/// sprmPFrameTextFlow - Frame text flow (operation 0x3A)
pub const SPRM_P_FRAME_TEXT_FLOW: u16 = 0x443A;

/// sprmPISnapBaseLine - Snap to baseline (operation 0x3B)
pub const SPRM_P_I_SNAP_BASE_LINE: u16 = 0x243B;

/// sprmPAnld - Autonumber list data (operation 0x3E)
pub const SPRM_P_ANLD: u16 = 0xC63E;

/// sprmPPropRMark - Property revision mark (operation 0x3F)
pub const SPRM_P_PROP_RMARK: u16 = 0xC63F;

/// sprmPOutLvl - Outline level (operation 0x40)
pub const SPRM_P_OUT_LVL: u16 = 0x2640;

/// sprmPFBiDi - Bi-directional paragraph (operation 0x41)
pub const SPRM_P_F_BI_DI: u16 = 0x2441;

/// sprmPFNumRMIns - Numbering revision insert (operation 0x43)
pub const SPRM_P_F_NUM_RM_INS: u16 = 0x2443;

/// sprmPCrLf - CR/LF (operation 0x44)
pub const SPRM_P_CR_LF: u16 = 0x2444;

/// sprmPNumRM - Numbering revision mark (operation 0x45)
pub const SPRM_P_NUM_RM: u16 = 0xC645;

/// sprmPHugePapx - Huge PAPX (operation 0x46)
pub const SPRM_P_HUGE_PAPX: u16 = 0x6646;

/// sprmPFUsePgsuSettings - Use page setup settings (operation 0x47)
pub const SPRM_P_F_USE_PGSU_SETTINGS: u16 = 0x2447;

/// sprmPFAdjustRight - Adjust right (operation 0x48)
pub const SPRM_P_F_ADJUST_RIGHT: u16 = 0x2448;

/// sprmPItap - Table nesting level (operation 0x49)
pub const SPRM_P_ITAP: u16 = 0x6649;

/// sprmPDtap - Table nesting delta (operation 0x4A)
pub const SPRM_P_DTAP: u16 = 0x664A;

/// sprmPFInnerTableCell - Inner table cell (operation 0x4B)
pub const SPRM_P_F_INNER_TABLE_CELL: u16 = 0x244B;

/// sprmPFInnerTtp - Inner table row end (operation 0x4C)
pub const SPRM_P_F_INNER_TTP: u16 = 0x244C;

/// sprmPShd - Shading (Word 2002+) (operation 0x4D)
pub const SPRM_P_SHD: u16 = 0xC64D;

/// sprmPBrcTop80 - Top border v80 (operation 0x4E)
pub const SPRM_P_BRC_TOP80: u16 = 0x664E;

/// sprmPBrcLeft80 - Left border v80 (operation 0x4F)
pub const SPRM_P_BRC_LEFT80: u16 = 0x664F;

/// sprmPBrcBottom80 - Bottom border v80 (operation 0x50)
pub const SPRM_P_BRC_BOTTOM80: u16 = 0x6650;

/// sprmPBrcRight80 - Right border v80 (operation 0x51)
pub const SPRM_P_BRC_RIGHT80: u16 = 0x6651;

/// sprmPBrcBetween80 - Between border v80 (operation 0x52)
pub const SPRM_P_BRC_BETWEEN80: u16 = 0x6652;

/// sprmPBrcBar80 - Bar border v80 (operation 0x53)
pub const SPRM_P_BRC_BAR80: u16 = 0x6653;

/// sprmPFNoAllowOverlap - No allow overlap (operation 0x54)
pub const SPRM_P_F_NO_ALLOW_OVERLAP: u16 = 0x2454;

/// sprmPWall - Wall (operation 0x55)
pub const SPRM_P_WALL: u16 = 0x6455;

/// sprmPIpgp - Page number (operation 0x56)
pub const SPRM_P_IPGP: u16 = 0x6456;

/// sprmPCnf - Conditional formatting (operation 0x57)
pub const SPRM_P_CNF: u16 = 0x6457;

/// sprmPRsid - Revision save ID (operation 0x60)
pub const SPRM_P_RSID: u16 = 0x6460;

/// sprmPIstdListPermute - List style permutation (operation 0x61)
pub const SPRM_P_ISTD_LIST_PERMUTE: u16 = 0xC661;

/// sprmPTableProps - Table properties (operation 0x62)
pub const SPRM_P_TABLE_PROPS: u16 = 0x6462;

/// sprmPTIstdInfo - Table style info (operation 0x63)
pub const SPRM_P_T_ISTD_INFO: u16 = 0xC663;

/// sprmPFContextualSpacing - Contextual spacing (operation 0x64)
pub const SPRM_P_F_CONTEXTUAL_SPACING: u16 = 0x2464;

/// sprmPPropRMark90 - Property revision mark v90 (operation 0x65)
pub const SPRM_P_PROP_RMARK90: u16 = 0xC665;

/// sprmPFMirrorIndents - Mirror indents (operation 0x66)
pub const SPRM_P_F_MIRROR_INDENTS: u16 = 0x2466;

/// sprmPTtwo - Table two (operation 0x67)
pub const SPRM_P_TTWO: u16 = 0x2467;

// TAP (Table Properties) SPRM opcodes
// Based on Apache POI's TableProperties and TableSprmUncompressor

/// sprmTJc - Table justification (operation 0x00)
pub const SPRM_T_JC: u16 = 0x5400;

/// sprmTIstd - Table style (operation 0x01)
pub const SPRM_T_ISTD: u16 = 0x5401; // Note: Size code should be 3 (4-byte)

/// sprmTDxaLeft - Table left position (operation 0x02)
pub const SPRM_T_DXA_LEFT: u16 = 0x9602;

/// sprmTDxaGapHalf - Half gap between cells (operation 0x03)
pub const SPRM_T_DXA_GAP_HALF: u16 = 0x9603;

/// sprmTFCantSplit - Table row cannot split (operation 0x04)
pub const SPRM_T_F_CANT_SPLIT: u16 = 0x3404;

/// sprmTTableHeader - Table header row (operation 0x05)
pub const SPRM_T_TABLE_HEADER: u16 = 0x3405;

/// sprmTTableBorders - Table borders (Word 6.0) (operation 0x06)
pub const SPRM_T_TABLE_BORDERS: u16 = 0xD606;

/// sprmTDefTable10 - Table definition (Word 6.0) (operation 0x07)
pub const SPRM_T_DEF_TABLE10: u16 = 0xD607;

/// sprmTDyaRowHeight - Row height (operation 0x08)
pub const SPRM_T_DYA_ROW_HEIGHT: u16 = 0x9608;

/// sprmTDefTable - Table definition (operation 0x09) (LONG SPRM)
pub const SPRM_T_DEF_TABLE: u16 = 0xD608;

/// sprmTDefTableShd - Table shading (operation 0x0A)
pub const SPRM_T_DEF_TABLE_SHD: u16 = 0xD60A;

/// sprmTTlp - Table layout (operation 0x0B)
pub const SPRM_T_TLP: u16 = 0x740B;

/// sprmTFBiDi - Bi-directional table (operation 0x0C)
pub const SPRM_T_F_BI_DI: u16 = 0x560C;

/// sprmTHTMLProps - HTML properties (operation 0x0D)
pub const SPRM_T_HTML_PROPS: u16 = 0x740D;

/// sprmTSetBrc - Set cell borders (operation 0x0E)
pub const SPRM_T_SET_BRC: u16 = 0xD60E;

/// sprmTInsert - Insert cells (operation 0x0F)
pub const SPRM_T_INSERT: u16 = 0x760F;

/// sprmTDelete - Delete cells (operation 0x10)
pub const SPRM_T_DELETE: u16 = 0x5610;

/// sprmTDxaCol - Column width (operation 0x11)
pub const SPRM_T_DXA_COL: u16 = 0x7611;

/// sprmTMerge - Merge cells (operation 0x12)
pub const SPRM_T_MERGE: u16 = 0x5612;

/// sprmTSplit - Split cells (operation 0x13)
pub const SPRM_T_SPLIT: u16 = 0x5613;

/// sprmTSetBrc10 - Set borders (Word 6.0) (operation 0x14)
pub const SPRM_T_SET_BRC10: u16 = 0xD614;

/// sprmTSetShd - Set shading (operation 0x15) (LONG SPRM in some cases)
pub const SPRM_T_SET_SHD: u16 = 0xD615;

/// sprmTSetShdOdd - Set shading odd (operation 0x16)
pub const SPRM_T_SET_SHD_ODD: u16 = 0xD616;

/// sprmTTextFlow - Text flow direction (operation 0x17)
pub const SPRM_T_TEXT_FLOW: u16 = 0x7617;

/// sprmTDiagLine - Diagonal line (operation 0x18) - Not used
pub const SPRM_T_DIAG_LINE: u16 = 0xD618;

/// sprmTVertMerge - Vertical merge (operation 0x19)
pub const SPRM_T_VERT_MERGE: u16 = 0xD619;

/// sprmTVertAlign - Vertical alignment (operation 0x1A)
pub const SPRM_T_VERT_ALIGN: u16 = 0xD61A;

/// sprmTCellPadding - Cell padding (operation 0x1B)
pub const SPRM_T_CELL_PADDING: u16 = 0xD61B;

/// sprmTCellSpacingDefault - Default cell spacing (operation 0x1C)
pub const SPRM_T_CELL_SPACING_DEFAULT: u16 = 0xD61C;

/// sprmTCellPaddingDefault - Default cell padding (operation 0x1D)
pub const SPRM_T_CELL_PADDING_DEFAULT: u16 = 0xD61D;

/// sprmTCellWidth - Cell width (operation 0x1E)
pub const SPRM_T_CELL_WIDTH: u16 = 0xD61E;

/// sprmTFitText - Fit text (operation 0x1F)
pub const SPRM_T_FIT_TEXT: u16 = 0xF61F;

/// sprmTFCellNoWrap - Cell no wrap (operation 0x20)
pub const SPRM_T_F_CELL_NO_WRAP: u16 = 0xD620;

/// sprmTIstdPermute - Table style permutation (operation 0x21)
pub const SPRM_T_ISTD_PERMUTE: u16 = 0xD621;

/// sprmTCellPaddingStyle - Cell padding style (operation 0x22)
pub const SPRM_T_CELL_PADDING_STYLE: u16 = 0xD622;

/// sprmTCellFHideMark - Hide end of cell mark (operation 0x23)
pub const SPRM_T_CELL_F_HIDE_MARK: u16 = 0xD623;

/// sprmTSetShdTable - Set table shading (operation 0x24)
pub const SPRM_T_SET_SHD_TABLE: u16 = 0xD624;

/// sprmTWidthBefore - Width before table (operation 0x25)
pub const SPRM_T_WIDTH_BEFORE: u16 = 0xF625;

/// sprmTWidthAfter - Width after table (operation 0x26)
pub const SPRM_T_WIDTH_AFTER: u16 = 0xF626;

/// sprmTFBiDi90 - Bi-directional v90 (operation 0x27)
pub const SPRM_T_F_BI_DI90: u16 = 0x3427;

/// sprmTFNoAllowOverlap - No allow overlap (operation 0x28)
pub const SPRM_T_F_NO_ALLOW_OVERLAP: u16 = 0x3428;

/// sprmTFCantOverlap - Cannot overlap (operation 0x29)
pub const SPRM_T_F_CANT_OVERLAP: u16 = 0x3429;

/// sprmTIpgp - Page number (operation 0x2A)
pub const SPRM_T_IPGP: u16 = 0x742A;

/// sprmTCnf - Conditional formatting (operation 0x2B)
pub const SPRM_T_CNF: u16 = 0xD62B;

/// sprmTDefTableShd80 - Table shading v80 (operation 0x2C)
pub const SPRM_T_DEF_TABLE_SHD80: u16 = 0xD62C;

/// sprmTDefTableShd2nd - Table shading 2nd (operation 0x2D)
pub const SPRM_T_DEF_TABLE_SHD2ND: u16 = 0xD62D;

/// sprmTDefTableShd3rd - Table shading 3rd (operation 0x2E)
pub const SPRM_T_DEF_TABLE_SHD3RD: u16 = 0xD62E;

/// sprmTCellBrcType - Cell border type (operation 0x2F)
pub const SPRM_T_CELL_BRC_TYPE: u16 = 0xD62F;

/// sprmTFAutofit - Autofit table (operation 0x30)
pub const SPRM_T_F_AUTOFIT: u16 = 0x3430;

/// sprmTDefTableShd - Table shading (operation 0x31)
pub const SPRM_T_DEF_TABLE_SHD_RAW: u16 = 0xD631;

/// sprmTDefTableShd2ndRaw - Table shading 2nd raw (operation 0x32)
pub const SPRM_T_DEF_TABLE_SHD2ND_RAW: u16 = 0xD632;

/// sprmTDefTableShd3rdRaw - Table shading 3rd raw (operation 0x33)
pub const SPRM_T_DEF_TABLE_SHD3RD_RAW: u16 = 0xD633;

/// sprmTRsid - Revision save ID (operation 0x34)
pub const SPRM_T_RSID: u16 = 0x7434;

/// sprmTCellVertAlignStyle - Cell vertical alignment style (operation 0x35)
pub const SPRM_T_CELL_VERT_ALIGN_STYLE: u16 = 0xD635;

/// sprmTCellNoWrapStyle - Cell no wrap style (operation 0x36)
pub const SPRM_T_CELL_NO_WRAP_STYLE: u16 = 0xD636;

/// sprmTCellBrcTopStyle - Cell top border style (operation 0x37)
pub const SPRM_T_CELL_BRC_TOP_STYLE: u16 = 0xD637;

/// sprmTCellBrcBottomStyle - Cell bottom border style (operation 0x38)
pub const SPRM_T_CELL_BRC_BOTTOM_STYLE: u16 = 0xD638;

/// sprmTCellBrcLeftStyle - Cell left border style (operation 0x39)
pub const SPRM_T_CELL_BRC_LEFT_STYLE: u16 = 0xD639;

/// sprmTCellBrcRightStyle - Cell right border style (operation 0x3A)
pub const SPRM_T_CELL_BRC_RIGHT_STYLE: u16 = 0xD63A;

/// sprmTCellBrcInsideHStyle - Cell inside horizontal border style (operation 0x3B)
pub const SPRM_T_CELL_BRC_INSIDE_H_STYLE: u16 = 0xD63B;

/// sprmTCellBrcInsideVStyle - Cell inside vertical border style (operation 0x3C)
pub const SPRM_T_CELL_BRC_INSIDE_V_STYLE: u16 = 0xD63C;

/// sprmTCellBrcTL2BRStyle - Cell top-left to bottom-right border style (operation 0x3D)
pub const SPRM_T_CELL_BRC_TL2BR_STYLE: u16 = 0xD63D;

/// sprmTCellBrcTR2BLStyle - Cell top-right to bottom-left border style (operation 0x3E)
pub const SPRM_T_CELL_BRC_TR2BL_STYLE: u16 = 0xD63E;

/// sprmTCellShdStyle - Cell shading style (operation 0x3F)
pub const SPRM_T_CELL_SHD_STYLE: u16 = 0xD63F;

/// sprmTCHorzBands - Horizontal banded columns (operation 0x40)
pub const SPRM_T_C_HORZ_BANDS: u16 = 0x3440;

/// sprmTCVertBands - Vertical banded rows (operation 0x41)
pub const SPRM_T_C_VERT_BANDS: u16 = 0x3441;

/// sprmTJc90 - Table justification v90 (operation 0x42)
pub const SPRM_T_JC90: u16 = 0x5442;

// SEP (Section Properties) SPRM opcodes (partial list for completeness)
// Based on Apache POI's SectionProperties

/// sprmSBkc - Break code (operation 0x00)
pub const SPRM_S_BKC: u16 = 0x3000;

/// sprmSFTitlePage - Title page (operation 0x01)
pub const SPRM_S_F_TITLE_PAGE: u16 = 0x3001;

/// sprmSCcolumns - Number of columns (operation 0x02)
pub const SPRM_S_C_COLUMNS: u16 = 0x3002;

/// sprmSDxaColumns - Column width (operation 0x03)
pub const SPRM_S_DXA_COLUMNS: u16 = 0x9003;

/// Extract SPRM type from opcode (bits 10-12).
///
/// Returns:
/// - 1: PAP (Paragraph Properties)
/// - 2: CHP (Character Properties)
/// - 3: PIC (Picture Properties)
/// - 4: SEP (Section Properties)
/// - 5: TAP (Table Properties)
#[inline]
pub fn get_sprm_type(opcode: u16) -> u8 {
    ((opcode >> 10) & 0x07) as u8
}

/// Extract SPRM operation code from opcode (bits 0-8).
#[inline]
pub fn get_sprm_operation(opcode: u16) -> u16 {
    opcode & 0x01FF
}

/// Extract SPRM size code from opcode (bits 13-15).
///
/// Returns:
/// - 0, 1: 1-byte operand
/// - 2, 4, 5: 2-byte operand
/// - 3: 4-byte operand
/// - 6: Variable length
/// - 7: 3-byte operand
#[inline]
pub fn get_sprm_size_code(opcode: u16) -> u8 {
    ((opcode >> 13) & 0x07) as u8
}

/// Check if SPRM is a "special" operation (bit 9).
#[inline]
pub fn is_sprm_special(opcode: u16) -> bool {
    (opcode & 0x0200) != 0
}

#[cfg(test)]
mod tests {
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
}
