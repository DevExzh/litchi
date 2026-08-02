//! Typed BIFF12 record-kind constants.

use super::Kind;

// Basic cell records
pub const ROW_HDR: Kind = Kind(0x0000);
pub const CELL_BLANK: Kind = Kind(0x0001);
pub const CELL_RK: Kind = Kind(0x0002);
pub const CELL_ERROR: Kind = Kind(0x0003);
pub const CELL_BOOL: Kind = Kind(0x0004);
pub const CELL_REAL: Kind = Kind(0x0005);
pub const CELL_ST: Kind = Kind(0x0006);
pub const CELL_ISST: Kind = Kind(0x0007);

// Formula records
pub const FMLA_STRING: Kind = Kind(0x0008);
pub const FMLA_NUM: Kind = Kind(0x0009);
pub const FMLA_BOOL: Kind = Kind(0x000A);
pub const FMLA_ERROR: Kind = Kind(0x000B);

// Shared string table
pub const SST_ITEM: Kind = Kind(0x0013);

// Format and style records
pub const FONT: Kind = Kind(0x002B);
pub const FMT: Kind = Kind(0x002C);
pub const FILL: Kind = Kind(0x002D);
pub const BORDER: Kind = Kind(0x002E);
pub const XF: Kind = Kind(0x002F);
pub const STYLE: Kind = Kind(0x0030);
pub const CELL_META: Kind = Kind(0x0031);
pub const VALUE_META: Kind = Kind(0x0032);

// Column and dimension records
pub const COL_INFO: Kind = Kind(0x003C);
pub const CELL_R_STRING: Kind = Kind(0x003E);

// Extension wrappers
pub const FRT_BEGIN: Kind = Kind(0x0023);
pub const FRT_END: Kind = Kind(0x0024);
pub const AC_BEGIN: Kind = Kind(0x0025);
pub const AC_END: Kind = Kind(0x0026);
pub const UID: Kind = Kind(0x028B);
pub const BEGIN_WEB_EXTENSIONS: Kind = Kind(2068);
pub const END_WEB_EXTENSIONS: Kind = Kind(2069);
pub const WEB_EXTENSION: Kind = Kind(2070);

// Workbook structure records
pub const FILE_VERSION: Kind = Kind(0x0080);
pub const BEGIN_SHEET: Kind = Kind(0x0081);
pub const END_SHEET: Kind = Kind(0x0082);
pub const BEGIN_BOOK: Kind = Kind(0x0083);
pub const END_BOOK: Kind = Kind(0x0084);
pub const BEGIN_WS_VIEWS: Kind = Kind(0x0085);
pub const END_WS_VIEWS: Kind = Kind(0x0086);
pub const BEGIN_BOOK_VIEWS: Kind = Kind(0x0087);
pub const END_BOOK_VIEWS: Kind = Kind(0x0088);
pub const BEGIN_WS_VIEW: Kind = Kind(0x0089);
pub const END_WS_VIEW: Kind = Kind(0x008A);
pub const BEGIN_CS_VIEWS: Kind = Kind(0x008B);
pub const END_CS_VIEWS: Kind = Kind(0x008C);
pub const BEGIN_CS_VIEW: Kind = Kind(0x008D);
pub const END_CS_VIEW: Kind = Kind(0x008E);
pub const BEGIN_BUNDLE_SHS: Kind = Kind(0x008F);
pub const END_BUNDLE_SHS: Kind = Kind(0x0090);
pub const BEGIN_SHEET_DATA: Kind = Kind(0x0091);
pub const END_SHEET_DATA: Kind = Kind(0x0092);
pub const WS_PROP: Kind = Kind(0x0093);
pub const WS_DIM: Kind = Kind(0x0094);

// Worksheet view records
pub const PANE: Kind = Kind(0x0097);
pub const SEL: Kind = Kind(0x0098);
pub const WORKBOOK_PROP: Kind = Kind(0x0099);
pub const BUNDLE_SH: Kind = Kind(0x009C);
pub const CALC_PROP: Kind = Kind(0x009D);
pub const BOOK_VIEW: Kind = Kind(0x009E);
pub const BEGIN_SST: Kind = Kind(0x009F);
pub const END_SST: Kind = Kind(0x00A0);

// Filter records
pub const BEGIN_A_FILTER: Kind = Kind(0x00A1);
pub const END_A_FILTER: Kind = Kind(0x00A2);
pub const BEGIN_FILTER_COLUMN: Kind = Kind(0x00A3);
pub const END_FILTER_COLUMN: Kind = Kind(0x00A4);
pub const BEGIN_FILTERS: Kind = Kind(0x00A5);
pub const END_FILTERS: Kind = Kind(0x00A6);
pub const FILTER: Kind = Kind(0x00A7);
pub const COLOR_FILTER: Kind = Kind(0x00A8);
pub const ICON_FILTER: Kind = Kind(0x00A9);
pub const TOP10_FILTER: Kind = Kind(0x00AA);
pub const DYNAMIC_FILTER: Kind = Kind(0x00AB);
pub const BEGIN_CUSTOM_FILTERS: Kind = Kind(0x00AC);
pub const END_CUSTOM_FILTERS: Kind = Kind(0x00AD);
pub const CUSTOM_FILTER: Kind = Kind(0x00AE);
pub const A_FILTER_DATE_GROUP_ITEM: Kind = Kind(0x00AF);

// Merge cells
pub const MERGE_CELL: Kind = Kind(0x00B0);
pub const BEGIN_MERGE_CELLS: Kind = Kind(0x00B1);
pub const END_MERGE_CELLS: Kind = Kind(0x00B2);

// Named ranges
pub const NAME: Kind = Kind(0x0027);

// Formulas and tables
pub const ARR_FMLA: Kind = Kind(0x01AA);
pub const SHR_FMLA: Kind = Kind(0x01AB);
pub const TABLE: Kind = Kind(0x01AC);

// Structured tables (ListObject stream, MS-XLSB 2.1.7.51)
pub const BEGIN_LIST: Kind = Kind(343);
pub const END_LIST: Kind = Kind(344);
pub const BEGIN_LIST_COLS: Kind = Kind(345);
pub const END_LIST_COLS: Kind = Kind(346);
pub const BEGIN_LIST_COL: Kind = Kind(347);
pub const END_LIST_COL: Kind = Kind(348);
pub const BEGIN_LIST_XML_CPR: Kind = Kind(349);
pub const END_LIST_XML_CPR: Kind = Kind(350);
pub const LIST_CC_FMLA: Kind = Kind(351);
pub const LIST_TR_FMLA: Kind = Kind(352);
pub const BEGIN_LIST_PARTS: Kind = Kind(660);
pub const LIST_PART: Kind = Kind(661);
pub const END_LIST_PARTS: Kind = Kind(662);
pub const LIST14: Kind = Kind(1111);

// External connections and links
pub const BEGIN_EXT_CONNECTION: Kind = Kind(201);
pub const END_EXT_CONNECTION: Kind = Kind(202);
pub const BEGIN_EC_DB_PROPS: Kind = Kind(203);
pub const END_EC_DB_PROPS: Kind = Kind(204);
pub const BEGIN_EC_OLAP_PROPS: Kind = Kind(205);
pub const END_EC_OLAP_PROPS: Kind = Kind(206);
pub const BEGIN_EC_WEB_PROPS: Kind = Kind(261);
pub const END_EC_WEB_PROPS: Kind = Kind(262);
pub const BEGIN_EC_WP_TABLES: Kind = Kind(263);
pub const END_EC_WP_TABLES: Kind = Kind(264);
pub const BEGIN_EC_PARAMS: Kind = Kind(265);
pub const END_EC_PARAMS: Kind = Kind(266);
pub const BEGIN_EC_PARAM: Kind = Kind(267);
pub const END_EC_PARAM: Kind = Kind(268);
pub const BEGIN_EXT_CONNECTIONS: Kind = Kind(429);
pub const END_EXT_CONNECTIONS: Kind = Kind(430);
pub const BEGIN_EC_TXT_WIZ: Kind = Kind(538);
pub const END_EC_TXT_WIZ: Kind = Kind(539);
pub const BEGIN_EXT_CONN14: Kind = Kind(1068);
pub const END_EXT_CONN14: Kind = Kind(1069);
pub const BEGIN_EXT_CONN15: Kind = Kind(2109);
pub const END_EXT_CONN15: Kind = Kind(2110);
pub const BEGIN_EC_TXT_WIZ15: Kind = Kind(2129);
pub const END_EC_TXT_WIZ15: Kind = Kind(2130);
pub const BEGIN_DATA_FEED_PR15: Kind = Kind(2113);
pub const END_DATA_FEED_PR15: Kind = Kind(2114);
pub const BEGIN_DB_TABLES15: Kind = Kind(2118);
pub const END_DB_TABLES15: Kind = Kind(2119);
pub const BEGIN_EXTERNALS: Kind = Kind(0x0161);
pub const END_EXTERNALS: Kind = Kind(0x0162);
pub const SUP_BOOK_SRC: Kind = Kind(0x0163);
pub const SUP_ADDIN: Kind = Kind(0x0164);
pub const SUP_SELF: Kind = Kind(0x0165);
pub const SUP_SAME: Kind = Kind(0x0166);
pub const SUP_TABS: Kind = Kind(0x0167);
pub const BEGIN_SUP_BOOK: Kind = Kind(0x0168);
pub const PLACEHOLDER_NAME: Kind = Kind(0x0169);
pub const EXTERN_SHEET: Kind = Kind(0x016A);
pub const EXTERN_TABLE_START: Kind = Kind(0x016B);
pub const EXTERN_TABLE_END: Kind = Kind(0x016C);
pub const EXTERN_ROW_HDR: Kind = Kind(0x016E);
pub const EXTERN_CELL_BLANK: Kind = Kind(0x016F);
pub const EXTERN_CELL_REAL: Kind = Kind(0x0170);
pub const EXTERN_CELL_BOOL: Kind = Kind(0x0171);
pub const EXTERN_CELL_ERROR: Kind = Kind(0x0172);
pub const EXTERN_CELL_STRING: Kind = Kind(0x0173);
pub const SUP_NAME_START: Kind = Kind(0x0241);
pub const SUP_NAME_VALUE_START: Kind = Kind(0x0242);
pub const SUP_NAME_VALUE_END: Kind = Kind(0x0243);
pub const SUP_NAME_NUM: Kind = Kind(0x0244);
pub const SUP_NAME_ERROR: Kind = Kind(0x0245);
pub const SUP_NAME_STRING: Kind = Kind(0x0246);
pub const SUP_NAME_NIL: Kind = Kind(0x0247);
pub const SUP_NAME_BOOL: Kind = Kind(0x0248);
pub const SUP_NAME_FORMULA: Kind = Kind(0x0249);
pub const SUP_NAME_BITS: Kind = Kind(0x024A);
pub const SUP_NAME_END: Kind = Kind(0x024B);
pub const END_SUP_BOOK: Kind = Kind(0x024C);

// Style sheet records
pub const BEGIN_STYLE_SHEET: Kind = Kind(0x0116);
pub const END_STYLE_SHEET: Kind = Kind(0x0117);
pub const BEGIN_FMTS: Kind = Kind(0x0267);
pub const END_FMTS: Kind = Kind(0x0268);
pub const BEGIN_FONTS: Kind = Kind(0x0263);
pub const END_FONTS: Kind = Kind(0x0264);
pub const BEGIN_FILLS: Kind = Kind(0x025B);
pub const END_FILLS: Kind = Kind(0x025C);
pub const BEGIN_BORDERS: Kind = Kind(0x0265);
pub const END_BORDERS: Kind = Kind(0x0266);
pub const BEGIN_CELL_XFS: Kind = Kind(0x0269);
pub const END_CELL_XFS: Kind = Kind(0x026A);
pub const BEGIN_STYLES: Kind = Kind(0x026B);
pub const END_STYLES: Kind = Kind(0x026C);
pub const BEGIN_CELL_STYLE_XFS: Kind = Kind(0x0272);
pub const END_CELL_STYLE_XFS: Kind = Kind(0x0273);
pub const BEGIN_DXFS: Kind = Kind(0x01F9);
pub const END_DXFS: Kind = Kind(0x01FA);
pub const DXF: Kind = Kind(0x01FB);
pub const BEGIN_TABLE_STYLES: Kind = Kind(0x01FC);
pub const END_TABLE_STYLES: Kind = Kind(0x01FD);

// Comments
pub const BEGIN_COMMENTS: Kind = Kind(0x0274);
pub const END_COMMENTS: Kind = Kind(0x0275);
pub const BEGIN_COMMENT_AUTHORS: Kind = Kind(0x0276);
pub const END_COMMENT_AUTHORS: Kind = Kind(0x0277);
pub const COMMENT_AUTHOR: Kind = Kind(0x0278);
pub const BEGIN_COMMENT_LIST: Kind = Kind(0x0279);
pub const END_COMMENT_LIST: Kind = Kind(0x027A);
pub const BEGIN_COMMENT: Kind = Kind(0x027B);
pub const END_COMMENT: Kind = Kind(0x027C);
pub const COMMENT_TEXT: Kind = Kind(0x027D);

// Hyperlinks
pub const H_LINK: Kind = Kind(0x01EE);

// Page setup
pub const MARGINS: Kind = Kind(0x01DC);
pub const PRINT_OPTIONS: Kind = Kind(0x01DD);
pub const PAGE_SETUP: Kind = Kind(0x01DE);
pub const BEGIN_HEADER_FOOTER: Kind = Kind(0x01DF);
pub const END_HEADER_FOOTER: Kind = Kind(0x01E0);

// Column information
pub const BEGIN_COL_INFOS: Kind = Kind(0x0186);
pub const END_COL_INFOS: Kind = Kind(0x0187);

// Drawing
pub const DRAWING: Kind = Kind(0x0226);
pub const LEGACY_DRAWING: Kind = Kind(0x0227);
pub const LEGACY_DRAWING_HF: Kind = Kind(0x0228);

// Chart sheet stream (MS-XLSB 2.1.7.7). The chart sheet view records
// (BEGIN_CS_VIEWS/END_CS_VIEWS/BEGIN_CS_VIEW/END_CS_VIEW) and the drawing
// records above are shared with the worksheet stream; these are specific
// to chart sheets. Note that `UID` (used for comment ACUid records) and
// `CS_PROP` share record type 651 in their respective streams.
pub const CS_PROP: Kind = Kind(651);
pub const CS_PAGE_SETUP: Kind = Kind(652);
pub const CS_PROTECTION: Kind = Kind(669);
pub const CS_PROTECTION_ISO: Kind = Kind(679);

// Data validation
pub const BEGIN_D_VALS: Kind = Kind(0x023D);
pub const END_D_VALS: Kind = Kind(0x023E);
pub const D_VAL: Kind = Kind(0x0040);
pub const D_VAL_LIST: Kind = Kind(0x02A9);
pub const D_VAL14: Kind = Kind(0x041D);
pub const BEGIN_D_VALS14: Kind = Kind(0x041E);
pub const END_D_VALS14: Kind = Kind(0x0482);

// Conditional formatting
pub const BEGIN_COND_FORMATTING: Kind = Kind(0x01CD);
pub const END_COND_FORMATTING: Kind = Kind(0x01CE);
pub const BEGIN_CF_RULE: Kind = Kind(0x01CF);
pub const END_CF_RULE: Kind = Kind(0x01D0);
pub const BEGIN_ICON_SET: Kind = Kind(0x01D1);
pub const END_ICON_SET: Kind = Kind(0x01D2);
pub const BEGIN_DATABAR: Kind = Kind(0x01D3);
pub const END_DATABAR: Kind = Kind(0x01D4);
pub const BEGIN_COLOR_SCALE: Kind = Kind(0x01D5);
pub const END_COLOR_SCALE: Kind = Kind(0x01D6);
pub const CFVO: Kind = Kind(0x01D7);
pub const COLOR: Kind = Kind(0x0234);
pub const BEGIN_COND_FORMATTING14: Kind = Kind(0x0416);
pub const END_COND_FORMATTING14: Kind = Kind(0x0417);
pub const BEGIN_CF_RULE14: Kind = Kind(0x0418);
pub const END_CF_RULE14: Kind = Kind(0x0419);
pub const CFVO14: Kind = Kind(0x041A);
pub const BEGIN_DATABAR14: Kind = Kind(0x041B);
pub const BEGIN_ICON_SET14: Kind = Kind(0x041C);
pub const COLOR14: Kind = Kind(0x041F);
pub const CF_ICON: Kind = Kind(0x0458);
pub const CF_RULE_EXT: Kind = Kind(0x047A);
pub const END_ICON_SET14: Kind = Kind(0x0483);
pub const END_DATABAR14: Kind = Kind(0x0484);
pub const BEGIN_COLOR_SCALE14: Kind = Kind(0x0485);
pub const END_COLOR_SCALE14: Kind = Kind(0x0486);

// Protection
pub const BOOK_PROTECTION: Kind = Kind(0x0216);
pub const SHEET_PROTECTION: Kind = Kind(0x0217);
pub const SHEET_PROTECTION_ISO: Kind = Kind(0x02A6);
pub const RANGE_PROTECTION: Kind = Kind(0x0218);

// Miscellaneous
pub const WS_FMT_INFO: Kind = Kind(0x01E5);
pub const BIG_NAME: Kind = Kind(0x0271);
pub const FILE_SHARING: Kind = Kind(0x0224);
pub const OLE_SIZE: Kind = Kind(0x0225);
pub const WEB_OPT: Kind = Kind(0x0229);
pub const PHONETIC_INFO: Kind = Kind(0x0219);

// Excel 2013+ records
pub const ABS_PATH15: Kind = Kind(0x0817);
pub const BEGIN_SPARKLINE_GROUPS: Kind = Kind(0x0422);
pub const END_SPARKLINE_GROUPS: Kind = Kind(0x0423);
pub const BEGIN_SPARKLINE_GROUP: Kind = Kind(0x0411);
pub const END_SPARKLINE_GROUP: Kind = Kind(0x0412);
pub const SPARKLINE: Kind = Kind(0x0413);

// PivotCache definition stream (MS-XLSB 2.1.7.38)
pub const BEGIN_PIVOT_CACHE_DEF: Kind = Kind(179);
pub const END_PIVOT_CACHE_DEF: Kind = Kind(180);
pub const BEGIN_PCD_FIELDS: Kind = Kind(181);
pub const END_PCD_FIELDS: Kind = Kind(182);
pub const BEGIN_PCD_FIELD: Kind = Kind(183);
pub const END_PCD_FIELD: Kind = Kind(184);
pub const BEGIN_PCD_SOURCE: Kind = Kind(185);
pub const END_PCD_SOURCE: Kind = Kind(186);
pub const BEGIN_PCDS_RANGE: Kind = Kind(187);
pub const END_PCDS_RANGE: Kind = Kind(188);
pub const BEGIN_PCDF_ATBL: Kind = Kind(189);
pub const END_PCDF_ATBL: Kind = Kind(190);
pub const BEGIN_PCDI_RUN: Kind = Kind(191);
pub const END_PCDI_RUN: Kind = Kind(192);
pub const BEGIN_PCD_HIERARCHIES: Kind = Kind(195);
pub const END_PCD_HIERARCHIES: Kind = Kind(196);
pub const BEGIN_PCD_HIERARCHY: Kind = Kind(197);
pub const END_PCD_HIERARCHY: Kind = Kind(198);
pub const BEGIN_PCDH_FIELDS_USAGE: Kind = Kind(199);
pub const END_PCDH_FIELDS_USAGE: Kind = Kind(200);
pub const BEGIN_PCDS_CONSOL: Kind = Kind(207);
pub const END_PCDS_CONSOL: Kind = Kind(208);
pub const BEGIN_PCDSC_PAGES: Kind = Kind(209);
pub const END_PCDSC_PAGES: Kind = Kind(210);
pub const BEGIN_PCDSC_PAGE: Kind = Kind(211);
pub const END_PCDSC_PAGE: Kind = Kind(212);
pub const BEGIN_PCDSCP_ITEM: Kind = Kind(213);
pub const END_PCDSCP_ITEM: Kind = Kind(214);
pub const BEGIN_PCDSC_SETS: Kind = Kind(215);
pub const END_PCDSC_SETS: Kind = Kind(216);
pub const BEGIN_PCDSC_SET: Kind = Kind(217);
pub const END_PCDSC_SET: Kind = Kind(218);
pub const BEGIN_PCDF_GROUP: Kind = Kind(219);
pub const END_PCDF_GROUP: Kind = Kind(220);
pub const BEGIN_PCDFG_ITEMS: Kind = Kind(221);
pub const END_PCDFG_ITEMS: Kind = Kind(222);
pub const BEGIN_PCDFG_RANGE: Kind = Kind(223);
pub const END_PCDFG_RANGE: Kind = Kind(224);
pub const BEGIN_PCDFG_DISCRETE: Kind = Kind(225);
pub const END_PCDFG_DISCRETE: Kind = Kind(226);
pub const BEGIN_PCDSD_TUPLE_CACHE: Kind = Kind(227);
pub const END_PCDSD_TUPLE_CACHE: Kind = Kind(228);
pub const BEGIN_PCDSDTC_ENTRIES: Kind = Kind(229);
pub const END_PCDSDTC_ENTRIES: Kind = Kind(230);
pub const BEGIN_PCDSDTC_MEMBERS: Kind = Kind(231);
pub const END_PCDSDTC_MEMBERS: Kind = Kind(232);
pub const BEGIN_PCDSDTC_MEMBER: Kind = Kind(233);
pub const END_PCDSDTC_MEMBER: Kind = Kind(234);
pub const BEGIN_PCDSDTC_QUERIES: Kind = Kind(235);
pub const END_PCDSDTC_QUERIES: Kind = Kind(236);
pub const BEGIN_PCDSDTC_QUERY: Kind = Kind(237);
pub const END_PCDSDTC_QUERY: Kind = Kind(238);
pub const BEGIN_PCDSDTC_SETS: Kind = Kind(239);
pub const END_PCDSDTC_SETS: Kind = Kind(240);
pub const BEGIN_PCDSDTC_SET: Kind = Kind(241);
pub const END_PCDSDTC_SET: Kind = Kind(242);
pub const BEGIN_PCD_CALC_ITEMS: Kind = Kind(243);
pub const END_PCD_CALC_ITEMS: Kind = Kind(244);
pub const BEGIN_PCD_CALC_ITEM: Kind = Kind(245);
pub const END_PCD_CALC_ITEM: Kind = Kind(246);
pub const BEGIN_PR_FILTERS: Kind = Kind(249);
pub const END_PR_FILTERS: Kind = Kind(250);
pub const BEGIN_PR_FILTER: Kind = Kind(251);
pub const END_PR_FILTER: Kind = Kind(252);
pub const BEGIN_P_NAMES: Kind = Kind(253);
pub const END_P_NAMES: Kind = Kind(254);
pub const BEGIN_P_NAME: Kind = Kind(255);
pub const END_P_NAME: Kind = Kind(256);
pub const BEGIN_PN_PAIRS: Kind = Kind(257);
pub const END_PN_PAIRS: Kind = Kind(258);
pub const BEGIN_PN_PAIR: Kind = Kind(259);
pub const END_PN_PAIR: Kind = Kind(260);
pub const BEGIN_PRF_ITEM: Kind = Kind(382);
pub const END_PRF_ITEM: Kind = Kind(383);
pub const BEGIN_PCD_CALC_MEMS: Kind = Kind(431);
pub const END_PCD_CALC_MEMS: Kind = Kind(432);
pub const BEGIN_PCD_CALC_MEM: Kind = Kind(433);
pub const END_PCD_CALC_MEM: Kind = Kind(434);
pub const BEGIN_PCDHG_LEVELS: Kind = Kind(435);
pub const END_PCDHG_LEVELS: Kind = Kind(436);
pub const BEGIN_PCDHG_LEVEL: Kind = Kind(437);
pub const END_PCDHG_LEVEL: Kind = Kind(438);
pub const BEGIN_PCDHGL_GROUPS: Kind = Kind(439);
pub const END_PCDHGL_GROUPS: Kind = Kind(440);
pub const BEGIN_PCDHGL_GROUP: Kind = Kind(441);
pub const END_PCDHGL_GROUP: Kind = Kind(442);
pub const BEGIN_PCDHGLG_MEMBERS: Kind = Kind(443);
pub const END_PCDHGLG_MEMBERS: Kind = Kind(444);
pub const BEGIN_PCDHGLG_MEMBER: Kind = Kind(445);
pub const END_PCDHGLG_MEMBER: Kind = Kind(446);
pub const BEGIN_PCDSDTC_MEMBERS_SORT_BY: Kind = Kind(646);
pub const END_PCDSDTC_MEMBERS_SORT_BY: Kind = Kind(647);
pub const BEGIN_PCD_SFCI_ENTRIES: Kind = Kind(657);
pub const END_PCD_SFCI_ENTRIES: Kind = Kind(658);
pub const BEGIN_PCD14: Kind = Kind(1066);
pub const END_PCD14: Kind = Kind(1067);
pub const BEGIN_PCD_CALC_MEM14: Kind = Kind(1038);
pub const END_PCD_CALC_MEM14: Kind = Kind(1039);
pub const BEGIN_PCD_CALC_MEM_EXT: Kind = Kind(1137);
pub const END_PCD_CALC_MEM_EXT: Kind = Kind(1138);
pub const BEGIN_PCD_CALC_MEMS_EXT: Kind = Kind(1139);
pub const END_PCD_CALC_MEMS_EXT: Kind = Kind(1140);
pub const BEGIN_PR_FILTERS14: Kind = Kind(1163);
pub const END_PR_FILTERS14: Kind = Kind(1164);
pub const BEGIN_PR_FILTER14: Kind = Kind(1165);
pub const END_PR_FILTER14: Kind = Kind(1166);
pub const BEGIN_PRF_ITEM14: Kind = Kind(1167);
pub const END_PRF_ITEM14: Kind = Kind(1168);
pub const BEGIN_ITEM_UNIQUE_NAMES: Kind = Kind(2106);
pub const END_ITEM_UNIQUE_NAMES: Kind = Kind(2107);

// PivotCache shared/cache items (MS-XLSB 2.4.728-2.4.740)
pub const PCDI_MISSING: Kind = Kind(20);
pub const PCDI_NUMBER: Kind = Kind(21);
pub const PCDI_BOOLEAN: Kind = Kind(22);
pub const PCDI_ERROR: Kind = Kind(23);
pub const PCDI_STRING: Kind = Kind(24);
pub const PCDI_DATETIME: Kind = Kind(25);
pub const PCDI_INDEX: Kind = Kind(26);
pub const PCDIA_MISSING: Kind = Kind(27);
pub const PCDIA_NUMBER: Kind = Kind(28);
pub const PCDIA_BOOLEAN: Kind = Kind(29);
pub const PCDIA_ERROR: Kind = Kind(30);
pub const PCDIA_STRING: Kind = Kind(31);
pub const PCDIA_DATETIME: Kind = Kind(32);

// PivotCache records stream (MS-XLSB 2.1.7.39)
pub const PCR_RECORD: Kind = Kind(33);
pub const PCR_RECORD_DT: Kind = Kind(34);
pub const BEGIN_PIVOT_CACHE_RECORDS: Kind = Kind(193);
pub const END_PIVOT_CACHE_RECORDS: Kind = Kind(194);

// Workbook PivotCache identifier collection (MS-XLSB 2.4.169-2.4.170)
pub const BEGIN_PIVOT_CACHE_IDS: Kind = Kind(384);
pub const END_PIVOT_CACHE_IDS: Kind = Kind(385);
pub const BEGIN_PIVOT_CACHE_ID: Kind = Kind(386);
pub const END_PIVOT_CACHE_ID: Kind = Kind(387);

// PivotTable definition stream (MS-XLSB 2.1.7.40)
pub const END_SXVI: Kind = Kind(281);
pub const BEGIN_SXVI: Kind = Kind(282);
pub const BEGIN_SXVIS: Kind = Kind(283);
pub const END_SXVIS: Kind = Kind(284);
pub const BEGIN_SXVD: Kind = Kind(285);
pub const END_SXVD: Kind = Kind(286);
pub const BEGIN_SXVDS: Kind = Kind(287);
pub const END_SXVDS: Kind = Kind(288);
pub const BEGIN_SXPI: Kind = Kind(289);
pub const END_SXPI: Kind = Kind(290);
pub const BEGIN_SXPIS: Kind = Kind(291);
pub const END_SXPIS: Kind = Kind(292);
pub const BEGIN_SXDI: Kind = Kind(293);
pub const END_SXDI: Kind = Kind(294);
pub const BEGIN_SXDIS: Kind = Kind(295);
pub const END_SXDIS: Kind = Kind(296);
pub const BEGIN_SXLI: Kind = Kind(297);
pub const END_SXLI: Kind = Kind(298);
pub const BEGIN_SXLI_RWS: Kind = Kind(299);
pub const END_SXLI_RWS: Kind = Kind(300);
pub const BEGIN_SXLI_COLS: Kind = Kind(301);
pub const END_SXLI_COLS: Kind = Kind(302);
pub const BEGIN_SX_FORMAT: Kind = Kind(303);
pub const END_SX_FORMAT: Kind = Kind(304);
pub const BEGIN_SX_FORMATS: Kind = Kind(305);
pub const END_SX_FORMATS: Kind = Kind(306);
pub const BEGIN_ISXVD_RWS: Kind = Kind(309);
pub const END_ISXVD_RWS: Kind = Kind(310);
pub const BEGIN_ISXVD_COLS: Kind = Kind(311);
pub const END_ISXVD_COLS: Kind = Kind(312);
pub const END_SX_LOCATION: Kind = Kind(313);
pub const BEGIN_SX_LOCATION: Kind = Kind(314);
pub const BEGIN_SX_VIEW: Kind = Kind(280);
pub const END_SX_VIEW: Kind = Kind(315);
pub const BEGIN_SX_THS: Kind = Kind(316);
pub const END_SX_THS: Kind = Kind(317);
pub const BEGIN_SX_TH: Kind = Kind(318);
pub const END_SX_TH: Kind = Kind(319);
pub const BEGIN_SX_TDMPS: Kind = Kind(324);
pub const END_SX_TDMPS: Kind = Kind(325);
pub const BEGIN_SX_TDMP: Kind = Kind(326);
pub const END_SX_TDMP: Kind = Kind(327);
pub const BEGIN_SX_TH_ITEMS: Kind = Kind(328);
pub const END_SX_TH_ITEMS: Kind = Kind(329);
pub const BEGIN_SX_TH_ITEM: Kind = Kind(330);
pub const END_SX_TH_ITEM: Kind = Kind(331);
pub const BEGIN_ISXVIS: Kind = Kind(388);
pub const END_ISXVIS: Kind = Kind(389);
pub const BEGIN_SX_CRT_FORMAT: Kind = Kind(481);
pub const END_SX_CRT_FORMAT: Kind = Kind(482);
pub const BEGIN_SX_CRT_FORMATS: Kind = Kind(483);
pub const END_SX_CRT_FORMATS: Kind = Kind(484);
pub const TABLE_STYLE_CLIENT: Kind = Kind(513);
pub const BEGIN_SX_COND_FMT: Kind = Kind(558);
pub const END_SX_COND_FMT: Kind = Kind(559);
pub const BEGIN_SX_COND_FMTS: Kind = Kind(560);
pub const END_SX_COND_FMTS: Kind = Kind(561);
pub const BEGIN_SX_FILTERS: Kind = Kind(599);
pub const END_SX_FILTERS: Kind = Kind(600);
pub const BEGIN_SX_FILTER: Kind = Kind(601);
pub const END_SX_FILTER: Kind = Kind(602);
pub const BEGIN_SX_TUPLE_SET: Kind = Kind(1026);
pub const END_SX_TUPLE_SET: Kind = Kind(1027);
pub const BEGIN_SX_TUPLE_SET_HEADER: Kind = Kind(1028);
pub const END_SX_TUPLE_SET_HEADER: Kind = Kind(1029);
pub const BEGIN_SX_TUPLE_SET_DATA: Kind = Kind(1031);
pub const END_SX_TUPLE_SET_DATA: Kind = Kind(1032);
pub const BEGIN_SX_TUPLE_SET_ROW: Kind = Kind(1033);
pub const END_SX_TUPLE_SET_ROW: Kind = Kind(1034);
pub const BEGIN_SX_VIEW14: Kind = Kind(1062);
pub const END_SX_VIEW14: Kind = Kind(1063);
pub const BEGIN_SX_VIEW16: Kind = Kind(1064);
pub const END_SX_VIEW16: Kind = Kind(1065);
pub const BEGIN_SX_EDITS: Kind = Kind(1120);
pub const END_SX_EDITS: Kind = Kind(1121);
pub const BEGIN_SX_CHANGE: Kind = Kind(1122);
pub const END_SX_CHANGE: Kind = Kind(1123);
pub const BEGIN_SX_CHANGES: Kind = Kind(1124);
pub const END_SX_CHANGES: Kind = Kind(1125);
pub const BEGIN_SX_COND_FMT14: Kind = Kind(1147);
pub const END_SX_COND_FMT14: Kind = Kind(1148);
pub const BEGIN_SX_COND_FMTS14: Kind = Kind(1149);
pub const END_SX_COND_FMTS14: Kind = Kind(1150);
pub const BEGIN_SXVCELLS: Kind = Kind(2055);
pub const END_SXVCELLS: Kind = Kind(2056);
pub const BEGIN_PIVOT_VERSION_INFO: Kind = Kind(5109);
pub const END_PIVOT_VERSION_INFO: Kind = Kind(5110);
pub const BEGIN_SXVD_SUBTOTALS: Kind = Kind(5136);
pub const END_SXVD_SUBTOTALS: Kind = Kind(5137);
pub const BEGIN_PIVOT_RULE_FILTER_SUBTOTALS: Kind = Kind(5139);
pub const END_PIVOT_RULE_FILTER_SUBTOTALS: Kind = Kind(5140);
pub const BEGIN_SXVD_SUBTOTAL_LINE_ITEMS: Kind = Kind(5143);
pub const END_SXVD_SUBTOTAL_LINE_ITEMS: Kind = Kind(5144);

// Record kinds used by strict host parsers but omitted from the original table.
pub const BEGIN_PRULE: Kind = Kind(247);
pub const END_PRULE: Kind = Kind(248);
pub const BEGIN_PCD_KPIS: Kind = Kind(269);
pub const END_PCD_KPIS: Kind = Kind(270);
pub const BEGIN_PCD_KPI: Kind = Kind(271);
pub const END_PCD_KPI: Kind = Kind(272);
pub const PCD_H14: Kind = Kind(1037);
pub const PCD_FIELD14: Kind = Kind(1141);
pub const PCD_SFCI_ENTRY: Kind = Kind(659);
