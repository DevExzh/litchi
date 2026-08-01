//! XLSB record parsing for Excel 2007+ binary format
//!
//! XLSB (Excel Binary Workbook) uses a different record structure than
//! the older XLS BIFF format. Records are stored in a ZIP container
//! and use a binary record format with variable-length encoding.

use crate::xlsb::error::{XlsbError, XlsbResult};
use bytes::Bytes;
use litchi_core::binary;
use std::io::Read;

/// XLSB record header (variable length encoding)
#[derive(Debug, Clone)]
pub struct XlsbRecordHeader {
    pub record_type: u16,
    pub data_len: usize,
}

impl XlsbRecordHeader {
    /// Read record header with variable-length encoding
    #[inline]
    pub fn read<R: Read>(reader: &mut R) -> XlsbResult<Self> {
        let mut b = [0u8; 1];
        reader.read_exact(&mut b)?;
        let mut record_type = (b[0] & 0x7F) as u16;

        if (b[0] & 0x80) != 0 {
            reader.read_exact(&mut b)?;
            record_type |= ((b[0] & 0x7F) as u16) << 7;

            if (b[0] & 0x80) != 0 {
                reader.read_exact(&mut b)?;
                record_type |= ((b[0] & 0x7F) as u16) << 14;
            }
        }

        // Read variable-length data size
        let mut data_len = 0usize;
        let mut shift = 0;

        loop {
            reader.read_exact(&mut b)?;
            data_len |= ((b[0] & 0x7F) as usize) << shift;
            shift += 7;

            if (b[0] & 0x80) == 0 {
                break;
            }

            if shift > 28 {
                return Err(XlsbError::InvalidLength {
                    expected: 0,
                    found: data_len,
                });
            }
        }

        Ok(XlsbRecordHeader {
            record_type,
            data_len,
        })
    }
}

/// XLSB record with header and data
#[derive(Debug, Clone)]
pub struct XlsbRecord {
    pub header: XlsbRecordHeader,
    pub data: Bytes,
}

impl XlsbRecord {
    /// Read a complete XLSB record
    pub fn read<R: Read>(reader: &mut R) -> XlsbResult<Self> {
        let header = XlsbRecordHeader::read(reader)?;

        let mut data_buf = vec![0u8; header.data_len];
        reader.read_exact(&mut data_buf)?;
        let data = Bytes::from(data_buf);

        Ok(XlsbRecord { header, data })
    }
}

/// Iterator over XLSB records in a stream
pub struct XlsbRecordIter<R> {
    reader: R,
}

impl<R: Read> XlsbRecordIter<R> {
    pub fn new(reader: R) -> Self {
        XlsbRecordIter { reader }
    }
}

impl<R: Read> Iterator for XlsbRecordIter<R> {
    type Item = XlsbResult<XlsbRecord>;

    fn next(&mut self) -> Option<Self::Item> {
        match XlsbRecord::read(&mut self.reader) {
            Ok(record) => Some(Ok(record)),
            Err(XlsbError::Io(e)) if e.kind() == std::io::ErrorKind::UnexpectedEof => None,
            Err(e) => Some(Err(e)),
        }
    }
}

/// XLSB record types (matching MS-XLSB specification)
/// Reference: [MS-XLSB] <https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-xlsb/>
#[allow(dead_code)]
pub mod record_types {
    // Basic cell records
    pub const ROW_HDR: u16 = 0x0000;
    pub const CELL_BLANK: u16 = 0x0001;
    pub const CELL_RK: u16 = 0x0002;
    pub const CELL_ERROR: u16 = 0x0003;
    pub const CELL_BOOL: u16 = 0x0004;
    pub const CELL_REAL: u16 = 0x0005;
    pub const CELL_ST: u16 = 0x0006;
    pub const CELL_ISST: u16 = 0x0007;

    // Formula records
    pub const FMLA_STRING: u16 = 0x0008;
    pub const FMLA_NUM: u16 = 0x0009;
    pub const FMLA_BOOL: u16 = 0x000A;
    pub const FMLA_ERROR: u16 = 0x000B;

    // Shared string table
    pub const SST_ITEM: u16 = 0x0013;

    // Format and style records
    pub const FONT: u16 = 0x002B;
    pub const FMT: u16 = 0x002C;
    pub const FILL: u16 = 0x002D;
    pub const BORDER: u16 = 0x002E;
    pub const XF: u16 = 0x002F;
    pub const STYLE: u16 = 0x0030;
    pub const CELL_META: u16 = 0x0031;
    pub const VALUE_META: u16 = 0x0032;

    // Column and dimension records
    pub const COL_INFO: u16 = 0x003C;
    pub const CELL_R_STRING: u16 = 0x003E;

    // Extension wrappers
    pub const FRT_BEGIN: u16 = 0x0023;
    pub const FRT_END: u16 = 0x0024;
    pub const AC_BEGIN: u16 = 0x0025;
    pub const AC_END: u16 = 0x0026;
    pub const UID: u16 = 0x028B;
    pub const BEGIN_WEB_EXTENSIONS: u16 = 2068;
    pub const END_WEB_EXTENSIONS: u16 = 2069;
    pub const WEB_EXTENSION: u16 = 2070;

    // Workbook structure records
    pub const FILE_VERSION: u16 = 0x0080;
    pub const BEGIN_SHEET: u16 = 0x0081;
    pub const END_SHEET: u16 = 0x0082;
    pub const BEGIN_BOOK: u16 = 0x0083;
    pub const END_BOOK: u16 = 0x0084;
    pub const BEGIN_WS_VIEWS: u16 = 0x0085;
    pub const END_WS_VIEWS: u16 = 0x0086;
    pub const BEGIN_BOOK_VIEWS: u16 = 0x0087;
    pub const END_BOOK_VIEWS: u16 = 0x0088;
    pub const BEGIN_WS_VIEW: u16 = 0x0089;
    pub const END_WS_VIEW: u16 = 0x008A;
    pub const BEGIN_CS_VIEWS: u16 = 0x008B;
    pub const END_CS_VIEWS: u16 = 0x008C;
    pub const BEGIN_CS_VIEW: u16 = 0x008D;
    pub const END_CS_VIEW: u16 = 0x008E;
    pub const BEGIN_BUNDLE_SHS: u16 = 0x008F;
    pub const END_BUNDLE_SHS: u16 = 0x0090;
    pub const BEGIN_SHEET_DATA: u16 = 0x0091;
    pub const END_SHEET_DATA: u16 = 0x0092;
    pub const WS_PROP: u16 = 0x0093;
    pub const WS_DIM: u16 = 0x0094;

    // Worksheet view records
    pub const PANE: u16 = 0x0097;
    pub const SEL: u16 = 0x0098;
    pub const WORKBOOK_PROP: u16 = 0x0099;
    pub const BUNDLE_SH: u16 = 0x009C;
    pub const CALC_PROP: u16 = 0x009D;
    pub const BOOK_VIEW: u16 = 0x009E;
    pub const BEGIN_SST: u16 = 0x009F;
    pub const END_SST: u16 = 0x00A0;

    // Filter records
    pub const BEGIN_A_FILTER: u16 = 0x00A1;
    pub const END_A_FILTER: u16 = 0x00A2;
    pub const BEGIN_FILTER_COLUMN: u16 = 0x00A3;
    pub const END_FILTER_COLUMN: u16 = 0x00A4;
    pub const BEGIN_FILTERS: u16 = 0x00A5;
    pub const END_FILTERS: u16 = 0x00A6;
    pub const FILTER: u16 = 0x00A7;
    pub const COLOR_FILTER: u16 = 0x00A8;
    pub const ICON_FILTER: u16 = 0x00A9;
    pub const TOP10_FILTER: u16 = 0x00AA;
    pub const DYNAMIC_FILTER: u16 = 0x00AB;
    pub const BEGIN_CUSTOM_FILTERS: u16 = 0x00AC;
    pub const END_CUSTOM_FILTERS: u16 = 0x00AD;
    pub const CUSTOM_FILTER: u16 = 0x00AE;
    pub const A_FILTER_DATE_GROUP_ITEM: u16 = 0x00AF;

    // Merge cells
    pub const MERGE_CELL: u16 = 0x00B0;
    pub const BEGIN_MERGE_CELLS: u16 = 0x00B1;
    pub const END_MERGE_CELLS: u16 = 0x00B2;

    // Named ranges
    pub const NAME: u16 = 0x0027;

    // Formulas and tables
    pub const ARR_FMLA: u16 = 0x01AA;
    pub const SHR_FMLA: u16 = 0x01AB;
    pub const TABLE: u16 = 0x01AC;

    // Structured tables (ListObject stream, MS-XLSB 2.1.7.51)
    pub const BEGIN_LIST: u16 = 343;
    pub const END_LIST: u16 = 344;
    pub const BEGIN_LIST_COLS: u16 = 345;
    pub const END_LIST_COLS: u16 = 346;
    pub const BEGIN_LIST_COL: u16 = 347;
    pub const END_LIST_COL: u16 = 348;
    pub const BEGIN_LIST_XML_CPR: u16 = 349;
    pub const END_LIST_XML_CPR: u16 = 350;
    pub const LIST_CC_FMLA: u16 = 351;
    pub const LIST_TR_FMLA: u16 = 352;
    pub const BEGIN_LIST_PARTS: u16 = 660;
    pub const LIST_PART: u16 = 661;
    pub const END_LIST_PARTS: u16 = 662;
    pub const LIST14: u16 = 1111;

    // External connections and links
    pub const BEGIN_EXT_CONNECTION: u16 = 201;
    pub const END_EXT_CONNECTION: u16 = 202;
    pub const BEGIN_EC_DB_PROPS: u16 = 203;
    pub const END_EC_DB_PROPS: u16 = 204;
    pub const BEGIN_EC_OLAP_PROPS: u16 = 205;
    pub const END_EC_OLAP_PROPS: u16 = 206;
    pub const BEGIN_EC_WEB_PROPS: u16 = 261;
    pub const END_EC_WEB_PROPS: u16 = 262;
    pub const BEGIN_EC_WP_TABLES: u16 = 263;
    pub const END_EC_WP_TABLES: u16 = 264;
    pub const BEGIN_EC_PARAMS: u16 = 265;
    pub const END_EC_PARAMS: u16 = 266;
    pub const BEGIN_EC_PARAM: u16 = 267;
    pub const END_EC_PARAM: u16 = 268;
    pub const BEGIN_EXT_CONNECTIONS: u16 = 429;
    pub const END_EXT_CONNECTIONS: u16 = 430;
    pub const BEGIN_EC_TXT_WIZ: u16 = 538;
    pub const END_EC_TXT_WIZ: u16 = 539;
    pub const BEGIN_EXT_CONN14: u16 = 1068;
    pub const END_EXT_CONN14: u16 = 1069;
    pub const BEGIN_EXT_CONN15: u16 = 2109;
    pub const END_EXT_CONN15: u16 = 2110;
    pub const BEGIN_EC_TXT_WIZ15: u16 = 2129;
    pub const END_EC_TXT_WIZ15: u16 = 2130;
    pub const BEGIN_DATA_FEED_PR15: u16 = 2113;
    pub const END_DATA_FEED_PR15: u16 = 2114;
    pub const BEGIN_DB_TABLES15: u16 = 2118;
    pub const END_DB_TABLES15: u16 = 2119;
    pub const BEGIN_EXTERNALS: u16 = 0x0161;
    pub const END_EXTERNALS: u16 = 0x0162;
    pub const SUP_BOOK_SRC: u16 = 0x0163;
    pub const SUP_ADDIN: u16 = 0x0164;
    pub const SUP_SELF: u16 = 0x0165;
    pub const SUP_SAME: u16 = 0x0166;
    pub const SUP_TABS: u16 = 0x0167;
    pub const BEGIN_SUP_BOOK: u16 = 0x0168;
    pub const PLACEHOLDER_NAME: u16 = 0x0169;
    pub const EXTERN_SHEET: u16 = 0x016A;
    pub const EXTERN_TABLE_START: u16 = 0x016B;
    pub const EXTERN_TABLE_END: u16 = 0x016C;
    pub const EXTERN_ROW_HDR: u16 = 0x016E;
    pub const EXTERN_CELL_BLANK: u16 = 0x016F;
    pub const EXTERN_CELL_REAL: u16 = 0x0170;
    pub const EXTERN_CELL_BOOL: u16 = 0x0171;
    pub const EXTERN_CELL_ERROR: u16 = 0x0172;
    pub const EXTERN_CELL_STRING: u16 = 0x0173;
    pub const SUP_NAME_START: u16 = 0x0241;
    pub const SUP_NAME_VALUE_START: u16 = 0x0242;
    pub const SUP_NAME_VALUE_END: u16 = 0x0243;
    pub const SUP_NAME_NUM: u16 = 0x0244;
    pub const SUP_NAME_ERROR: u16 = 0x0245;
    pub const SUP_NAME_STRING: u16 = 0x0246;
    pub const SUP_NAME_NIL: u16 = 0x0247;
    pub const SUP_NAME_BOOL: u16 = 0x0248;
    pub const SUP_NAME_FORMULA: u16 = 0x0249;
    pub const SUP_NAME_BITS: u16 = 0x024A;
    pub const SUP_NAME_END: u16 = 0x024B;
    pub const END_SUP_BOOK: u16 = 0x024C;

    // Style sheet records
    pub const BEGIN_STYLE_SHEET: u16 = 0x0116;
    pub const END_STYLE_SHEET: u16 = 0x0117;
    pub const BEGIN_FMTS: u16 = 0x0267;
    pub const END_FMTS: u16 = 0x0268;
    pub const BEGIN_FONTS: u16 = 0x0263;
    pub const END_FONTS: u16 = 0x0264;
    pub const BEGIN_FILLS: u16 = 0x025B;
    pub const END_FILLS: u16 = 0x025C;
    pub const BEGIN_BORDERS: u16 = 0x0265;
    pub const END_BORDERS: u16 = 0x0266;
    pub const BEGIN_CELL_XFS: u16 = 0x0269;
    pub const END_CELL_XFS: u16 = 0x026A;
    pub const BEGIN_STYLES: u16 = 0x026B;
    pub const END_STYLES: u16 = 0x026C;
    pub const BEGIN_CELL_STYLE_XFS: u16 = 0x0272;
    pub const END_CELL_STYLE_XFS: u16 = 0x0273;
    pub const BEGIN_DXFS: u16 = 0x01F9;
    pub const END_DXFS: u16 = 0x01FA;
    pub const DXF: u16 = 0x01FB;
    pub const BEGIN_TABLE_STYLES: u16 = 0x01FC;
    pub const END_TABLE_STYLES: u16 = 0x01FD;

    // Comments
    pub const BEGIN_COMMENTS: u16 = 0x0274;
    pub const END_COMMENTS: u16 = 0x0275;
    pub const BEGIN_COMMENT_AUTHORS: u16 = 0x0276;
    pub const END_COMMENT_AUTHORS: u16 = 0x0277;
    pub const COMMENT_AUTHOR: u16 = 0x0278;
    pub const BEGIN_COMMENT_LIST: u16 = 0x0279;
    pub const END_COMMENT_LIST: u16 = 0x027A;
    pub const BEGIN_COMMENT: u16 = 0x027B;
    pub const END_COMMENT: u16 = 0x027C;
    pub const COMMENT_TEXT: u16 = 0x027D;

    // Hyperlinks
    pub const H_LINK: u16 = 0x01EE;

    // Page setup
    pub const MARGINS: u16 = 0x01DC;
    pub const PRINT_OPTIONS: u16 = 0x01DD;
    pub const PAGE_SETUP: u16 = 0x01DE;
    pub const BEGIN_HEADER_FOOTER: u16 = 0x01DF;
    pub const END_HEADER_FOOTER: u16 = 0x01E0;

    // Column information
    pub const BEGIN_COL_INFOS: u16 = 0x0186;
    pub const END_COL_INFOS: u16 = 0x0187;

    // Drawing
    pub const DRAWING: u16 = 0x0226;
    pub const LEGACY_DRAWING: u16 = 0x0227;
    pub const LEGACY_DRAWING_HF: u16 = 0x0228;

    // Chart sheet stream (MS-XLSB 2.1.7.7). The chart sheet view records
    // (BEGIN_CS_VIEWS/END_CS_VIEWS/BEGIN_CS_VIEW/END_CS_VIEW) and the drawing
    // records above are shared with the worksheet stream; these are specific
    // to chart sheets. Note that `UID` (used for comment ACUid records) and
    // `CS_PROP` share record type 651 in their respective streams.
    pub const CS_PROP: u16 = 651;
    pub const CS_PAGE_SETUP: u16 = 652;
    pub const CS_PROTECTION: u16 = 669;
    pub const CS_PROTECTION_ISO: u16 = 679;

    // Data validation
    pub const BEGIN_D_VALS: u16 = 0x023D;
    pub const END_D_VALS: u16 = 0x023E;
    pub const D_VAL: u16 = 0x0040;
    pub const D_VAL_LIST: u16 = 0x02A9;
    pub const D_VAL14: u16 = 0x041D;
    pub const BEGIN_D_VALS14: u16 = 0x041E;
    pub const END_D_VALS14: u16 = 0x0482;

    // Conditional formatting
    pub const BEGIN_COND_FORMATTING: u16 = 0x01CD;
    pub const END_COND_FORMATTING: u16 = 0x01CE;
    pub const BEGIN_CF_RULE: u16 = 0x01CF;
    pub const END_CF_RULE: u16 = 0x01D0;
    pub const BEGIN_ICON_SET: u16 = 0x01D1;
    pub const END_ICON_SET: u16 = 0x01D2;
    pub const BEGIN_DATABAR: u16 = 0x01D3;
    pub const END_DATABAR: u16 = 0x01D4;
    pub const BEGIN_COLOR_SCALE: u16 = 0x01D5;
    pub const END_COLOR_SCALE: u16 = 0x01D6;
    pub const CFVO: u16 = 0x01D7;
    pub const COLOR: u16 = 0x0234;
    pub const BEGIN_COND_FORMATTING14: u16 = 0x0416;
    pub const END_COND_FORMATTING14: u16 = 0x0417;
    pub const BEGIN_CF_RULE14: u16 = 0x0418;
    pub const END_CF_RULE14: u16 = 0x0419;
    pub const CFVO14: u16 = 0x041A;
    pub const BEGIN_DATABAR14: u16 = 0x041B;
    pub const BEGIN_ICON_SET14: u16 = 0x041C;
    pub const COLOR14: u16 = 0x041F;
    pub const CF_ICON: u16 = 0x0458;
    pub const CF_RULE_EXT: u16 = 0x047A;
    pub const END_ICON_SET14: u16 = 0x0483;
    pub const END_DATABAR14: u16 = 0x0484;
    pub const BEGIN_COLOR_SCALE14: u16 = 0x0485;
    pub const END_COLOR_SCALE14: u16 = 0x0486;

    // Protection
    pub const BOOK_PROTECTION: u16 = 0x0216;
    pub const SHEET_PROTECTION: u16 = 0x0217;
    pub const SHEET_PROTECTION_ISO: u16 = 0x02A6;
    pub const RANGE_PROTECTION: u16 = 0x0218;

    // Miscellaneous
    pub const WS_FMT_INFO: u16 = 0x01E5;
    pub const BIG_NAME: u16 = 0x0271;
    pub const FILE_SHARING: u16 = 0x0224;
    pub const OLE_SIZE: u16 = 0x0225;
    pub const WEB_OPT: u16 = 0x0229;
    pub const PHONETIC_INFO: u16 = 0x0219;

    // Excel 2013+ records
    pub const ABS_PATH15: u16 = 0x0817;
    pub const BEGIN_SPARKLINE_GROUPS: u16 = 0x0422;
    pub const END_SPARKLINE_GROUPS: u16 = 0x0423;
    pub const BEGIN_SPARKLINE_GROUP: u16 = 0x0411;
    pub const END_SPARKLINE_GROUP: u16 = 0x0412;
    pub const SPARKLINE: u16 = 0x0413;

    // PivotCache definition stream (MS-XLSB 2.1.7.38)
    pub const BEGIN_PIVOT_CACHE_DEF: u16 = 179;
    pub const END_PIVOT_CACHE_DEF: u16 = 180;
    pub const BEGIN_PCD_FIELDS: u16 = 181;
    pub const END_PCD_FIELDS: u16 = 182;
    pub const BEGIN_PCD_FIELD: u16 = 183;
    pub const END_PCD_FIELD: u16 = 184;
    pub const BEGIN_PCD_SOURCE: u16 = 185;
    pub const END_PCD_SOURCE: u16 = 186;
    pub const BEGIN_PCDS_RANGE: u16 = 187;
    pub const END_PCDS_RANGE: u16 = 188;
    pub const BEGIN_PCDF_ATBL: u16 = 189;
    pub const END_PCDF_ATBL: u16 = 190;
    pub const BEGIN_PCDI_RUN: u16 = 191;
    pub const END_PCDI_RUN: u16 = 192;
    pub const BEGIN_PCD_HIERARCHIES: u16 = 195;
    pub const END_PCD_HIERARCHIES: u16 = 196;
    pub const BEGIN_PCD_HIERARCHY: u16 = 197;
    pub const END_PCD_HIERARCHY: u16 = 198;
    pub const BEGIN_PCDH_FIELDS_USAGE: u16 = 199;
    pub const END_PCDH_FIELDS_USAGE: u16 = 200;
    pub const BEGIN_PCDS_CONSOL: u16 = 207;
    pub const END_PCDS_CONSOL: u16 = 208;
    pub const BEGIN_PCDSC_PAGES: u16 = 209;
    pub const END_PCDSC_PAGES: u16 = 210;
    pub const BEGIN_PCDSC_PAGE: u16 = 211;
    pub const END_PCDSC_PAGE: u16 = 212;
    pub const BEGIN_PCDSCP_ITEM: u16 = 213;
    pub const END_PCDSCP_ITEM: u16 = 214;
    pub const BEGIN_PCDSC_SETS: u16 = 215;
    pub const END_PCDSC_SETS: u16 = 216;
    pub const BEGIN_PCDSC_SET: u16 = 217;
    pub const END_PCDSC_SET: u16 = 218;
    pub const BEGIN_PCDF_GROUP: u16 = 219;
    pub const END_PCDF_GROUP: u16 = 220;
    pub const BEGIN_PCDFG_ITEMS: u16 = 221;
    pub const END_PCDFG_ITEMS: u16 = 222;
    pub const BEGIN_PCDFG_RANGE: u16 = 223;
    pub const END_PCDFG_RANGE: u16 = 224;
    pub const BEGIN_PCDFG_DISCRETE: u16 = 225;
    pub const END_PCDFG_DISCRETE: u16 = 226;
    pub const BEGIN_PCDSD_TUPLE_CACHE: u16 = 227;
    pub const END_PCDSD_TUPLE_CACHE: u16 = 228;
    pub const BEGIN_PCDSDTC_ENTRIES: u16 = 229;
    pub const END_PCDSDTC_ENTRIES: u16 = 230;
    pub const BEGIN_PCDSDTC_MEMBERS: u16 = 231;
    pub const END_PCDSDTC_MEMBERS: u16 = 232;
    pub const BEGIN_PCDSDTC_MEMBER: u16 = 233;
    pub const END_PCDSDTC_MEMBER: u16 = 234;
    pub const BEGIN_PCDSDTC_QUERIES: u16 = 235;
    pub const END_PCDSDTC_QUERIES: u16 = 236;
    pub const BEGIN_PCDSDTC_QUERY: u16 = 237;
    pub const END_PCDSDTC_QUERY: u16 = 238;
    pub const BEGIN_PCDSDTC_SETS: u16 = 239;
    pub const END_PCDSDTC_SETS: u16 = 240;
    pub const BEGIN_PCDSDTC_SET: u16 = 241;
    pub const END_PCDSDTC_SET: u16 = 242;
    pub const BEGIN_PCD_CALC_ITEMS: u16 = 243;
    pub const END_PCD_CALC_ITEMS: u16 = 244;
    pub const BEGIN_PCD_CALC_ITEM: u16 = 245;
    pub const END_PCD_CALC_ITEM: u16 = 246;
    pub const BEGIN_PR_FILTERS: u16 = 249;
    pub const END_PR_FILTERS: u16 = 250;
    pub const BEGIN_PR_FILTER: u16 = 251;
    pub const END_PR_FILTER: u16 = 252;
    pub const BEGIN_P_NAMES: u16 = 253;
    pub const END_P_NAMES: u16 = 254;
    pub const BEGIN_P_NAME: u16 = 255;
    pub const END_P_NAME: u16 = 256;
    pub const BEGIN_PN_PAIRS: u16 = 257;
    pub const END_PN_PAIRS: u16 = 258;
    pub const BEGIN_PN_PAIR: u16 = 259;
    pub const END_PN_PAIR: u16 = 260;
    pub const BEGIN_PRF_ITEM: u16 = 382;
    pub const END_PRF_ITEM: u16 = 383;
    pub const BEGIN_PCD_CALC_MEMS: u16 = 431;
    pub const END_PCD_CALC_MEMS: u16 = 432;
    pub const BEGIN_PCD_CALC_MEM: u16 = 433;
    pub const END_PCD_CALC_MEM: u16 = 434;
    pub const BEGIN_PCDHG_LEVELS: u16 = 435;
    pub const END_PCDHG_LEVELS: u16 = 436;
    pub const BEGIN_PCDHG_LEVEL: u16 = 437;
    pub const END_PCDHG_LEVEL: u16 = 438;
    pub const BEGIN_PCDHGL_GROUPS: u16 = 439;
    pub const END_PCDHGL_GROUPS: u16 = 440;
    pub const BEGIN_PCDHGL_GROUP: u16 = 441;
    pub const END_PCDHGL_GROUP: u16 = 442;
    pub const BEGIN_PCDHGLG_MEMBERS: u16 = 443;
    pub const END_PCDHGLG_MEMBERS: u16 = 444;
    pub const BEGIN_PCDHGLG_MEMBER: u16 = 445;
    pub const END_PCDHGLG_MEMBER: u16 = 446;
    pub const BEGIN_PCDSDTC_MEMBERS_SORT_BY: u16 = 646;
    pub const END_PCDSDTC_MEMBERS_SORT_BY: u16 = 647;
    pub const BEGIN_PCD_SFCI_ENTRIES: u16 = 657;
    pub const END_PCD_SFCI_ENTRIES: u16 = 658;
    pub const BEGIN_PCD14: u16 = 1066;
    pub const END_PCD14: u16 = 1067;
    pub const BEGIN_PCD_CALC_MEM14: u16 = 1038;
    pub const END_PCD_CALC_MEM14: u16 = 1039;
    pub const BEGIN_PCD_CALC_MEM_EXT: u16 = 1137;
    pub const END_PCD_CALC_MEM_EXT: u16 = 1138;
    pub const BEGIN_PCD_CALC_MEMS_EXT: u16 = 1139;
    pub const END_PCD_CALC_MEMS_EXT: u16 = 1140;
    pub const BEGIN_PR_FILTERS14: u16 = 1163;
    pub const END_PR_FILTERS14: u16 = 1164;
    pub const BEGIN_PR_FILTER14: u16 = 1165;
    pub const END_PR_FILTER14: u16 = 1166;
    pub const BEGIN_PRF_ITEM14: u16 = 1167;
    pub const END_PRF_ITEM14: u16 = 1168;
    pub const BEGIN_ITEM_UNIQUE_NAMES: u16 = 2106;
    pub const END_ITEM_UNIQUE_NAMES: u16 = 2107;

    // PivotCache shared/cache items (MS-XLSB 2.4.728-2.4.740)
    pub const PCDI_MISSING: u16 = 20;
    pub const PCDI_NUMBER: u16 = 21;
    pub const PCDI_BOOLEAN: u16 = 22;
    pub const PCDI_ERROR: u16 = 23;
    pub const PCDI_STRING: u16 = 24;
    pub const PCDI_DATETIME: u16 = 25;
    pub const PCDI_INDEX: u16 = 26;
    pub const PCDIA_MISSING: u16 = 27;
    pub const PCDIA_NUMBER: u16 = 28;
    pub const PCDIA_BOOLEAN: u16 = 29;
    pub const PCDIA_ERROR: u16 = 30;
    pub const PCDIA_STRING: u16 = 31;
    pub const PCDIA_DATETIME: u16 = 32;

    // PivotCache records stream (MS-XLSB 2.1.7.39)
    pub const PCR_RECORD: u16 = 33;
    pub const PCR_RECORD_DT: u16 = 34;
    pub const BEGIN_PIVOT_CACHE_RECORDS: u16 = 193;
    pub const END_PIVOT_CACHE_RECORDS: u16 = 194;

    // Workbook PivotCache identifier collection (MS-XLSB 2.4.169-2.4.170)
    pub const BEGIN_PIVOT_CACHE_IDS: u16 = 384;
    pub const END_PIVOT_CACHE_IDS: u16 = 385;
    pub const BEGIN_PIVOT_CACHE_ID: u16 = 386;
    pub const END_PIVOT_CACHE_ID: u16 = 387;

    // PivotTable definition stream (MS-XLSB 2.1.7.40)
    pub const END_SXVI: u16 = 281;
    pub const BEGIN_SXVI: u16 = 282;
    pub const BEGIN_SXVIS: u16 = 283;
    pub const END_SXVIS: u16 = 284;
    pub const BEGIN_SXVD: u16 = 285;
    pub const END_SXVD: u16 = 286;
    pub const BEGIN_SXVDS: u16 = 287;
    pub const END_SXVDS: u16 = 288;
    pub const BEGIN_SXPI: u16 = 289;
    pub const END_SXPI: u16 = 290;
    pub const BEGIN_SXPIS: u16 = 291;
    pub const END_SXPIS: u16 = 292;
    pub const BEGIN_SXDI: u16 = 293;
    pub const END_SXDI: u16 = 294;
    pub const BEGIN_SXDIS: u16 = 295;
    pub const END_SXDIS: u16 = 296;
    pub const BEGIN_SXLI: u16 = 297;
    pub const END_SXLI: u16 = 298;
    pub const BEGIN_SXLI_RWS: u16 = 299;
    pub const END_SXLI_RWS: u16 = 300;
    pub const BEGIN_SXLI_COLS: u16 = 301;
    pub const END_SXLI_COLS: u16 = 302;
    pub const BEGIN_SX_FORMAT: u16 = 303;
    pub const END_SX_FORMAT: u16 = 304;
    pub const BEGIN_SX_FORMATS: u16 = 305;
    pub const END_SX_FORMATS: u16 = 306;
    pub const BEGIN_ISXVD_RWS: u16 = 309;
    pub const END_ISXVD_RWS: u16 = 310;
    pub const BEGIN_ISXVD_COLS: u16 = 311;
    pub const END_ISXVD_COLS: u16 = 312;
    pub const END_SX_LOCATION: u16 = 313;
    pub const BEGIN_SX_LOCATION: u16 = 314;
    pub const BEGIN_SX_VIEW: u16 = 280;
    pub const END_SX_VIEW: u16 = 315;
    pub const BEGIN_SX_THS: u16 = 316;
    pub const END_SX_THS: u16 = 317;
    pub const BEGIN_SX_TH: u16 = 318;
    pub const END_SX_TH: u16 = 319;
    pub const BEGIN_SX_TDMPS: u16 = 324;
    pub const END_SX_TDMPS: u16 = 325;
    pub const BEGIN_SX_TDMP: u16 = 326;
    pub const END_SX_TDMP: u16 = 327;
    pub const BEGIN_SX_TH_ITEMS: u16 = 328;
    pub const END_SX_TH_ITEMS: u16 = 329;
    pub const BEGIN_SX_TH_ITEM: u16 = 330;
    pub const END_SX_TH_ITEM: u16 = 331;
    pub const BEGIN_ISXVIS: u16 = 388;
    pub const END_ISXVIS: u16 = 389;
    pub const BEGIN_SX_CRT_FORMAT: u16 = 481;
    pub const END_SX_CRT_FORMAT: u16 = 482;
    pub const BEGIN_SX_CRT_FORMATS: u16 = 483;
    pub const END_SX_CRT_FORMATS: u16 = 484;
    pub const TABLE_STYLE_CLIENT: u16 = 513;
    pub const BEGIN_SX_COND_FMT: u16 = 558;
    pub const END_SX_COND_FMT: u16 = 559;
    pub const BEGIN_SX_COND_FMTS: u16 = 560;
    pub const END_SX_COND_FMTS: u16 = 561;
    pub const BEGIN_SX_FILTERS: u16 = 599;
    pub const END_SX_FILTERS: u16 = 600;
    pub const BEGIN_SX_FILTER: u16 = 601;
    pub const END_SX_FILTER: u16 = 602;
    pub const BEGIN_SX_TUPLE_SET: u16 = 1026;
    pub const END_SX_TUPLE_SET: u16 = 1027;
    pub const BEGIN_SX_TUPLE_SET_HEADER: u16 = 1028;
    pub const END_SX_TUPLE_SET_HEADER: u16 = 1029;
    pub const BEGIN_SX_TUPLE_SET_DATA: u16 = 1031;
    pub const END_SX_TUPLE_SET_DATA: u16 = 1032;
    pub const BEGIN_SX_TUPLE_SET_ROW: u16 = 1033;
    pub const END_SX_TUPLE_SET_ROW: u16 = 1034;
    pub const BEGIN_SX_VIEW14: u16 = 1062;
    pub const END_SX_VIEW14: u16 = 1063;
    pub const BEGIN_SX_VIEW16: u16 = 1064;
    pub const END_SX_VIEW16: u16 = 1065;
    pub const BEGIN_SX_EDITS: u16 = 1120;
    pub const END_SX_EDITS: u16 = 1121;
    pub const BEGIN_SX_CHANGE: u16 = 1122;
    pub const END_SX_CHANGE: u16 = 1123;
    pub const BEGIN_SX_CHANGES: u16 = 1124;
    pub const END_SX_CHANGES: u16 = 1125;
    pub const BEGIN_SX_COND_FMT14: u16 = 1147;
    pub const END_SX_COND_FMT14: u16 = 1148;
    pub const BEGIN_SX_COND_FMTS14: u16 = 1149;
    pub const END_SX_COND_FMTS14: u16 = 1150;
    pub const BEGIN_SXVCELLS: u16 = 2055;
    pub const END_SXVCELLS: u16 = 2056;
    pub const BEGIN_PIVOT_VERSION_INFO: u16 = 5109;
    pub const END_PIVOT_VERSION_INFO: u16 = 5110;
    pub const BEGIN_SXVD_SUBTOTALS: u16 = 5136;
    pub const END_SXVD_SUBTOTALS: u16 = 5137;
    pub const BEGIN_PIVOT_RULE_FILTER_SUBTOTALS: u16 = 5139;
    pub const END_PIVOT_RULE_FILTER_SUBTOTALS: u16 = 5140;
    pub const BEGIN_SXVD_SUBTOTAL_LINE_ITEMS: u16 = 5143;
    pub const END_SXVD_SUBTOTAL_LINE_ITEMS: u16 = 5144;
}

/// Decode wide string (UTF-16LE) from XLSB format
pub fn wide_str(buf: &[u8], str_len: &mut usize) -> XlsbResult<String> {
    if buf.len() < 4 {
        return Err(XlsbError::InvalidLength {
            expected: 4,
            found: buf.len(),
        });
    }

    let len = binary::read_u32_le_at(buf, 0)? as usize;
    let consumed = len
        .checked_mul(2)
        .and_then(|byte_len| byte_len.checked_add(4))
        .ok_or_else(|| XlsbError::Encoding("wide string length overflow".to_string()))?;
    if buf.len() < consumed {
        return Err(XlsbError::WideStringLength {
            expected: consumed,
            actual: buf.len(),
        });
    }

    *str_len = consumed;
    let utf16_data = &buf[4..*str_len];

    // Convert UTF-16LE to UTF-8 using encoding_rs
    use encoding_rs::UTF_16LE;
    Ok(UTF_16LE.decode(utf16_data).0.into_owned())
}

/// Decode wide string (UTF-16LE) from XLSB format and return consumed bytes
pub fn wide_str_with_len(buf: &[u8]) -> XlsbResult<(String, usize)> {
    if buf.len() < 4 {
        return Err(XlsbError::InvalidLength {
            expected: 4,
            found: buf.len(),
        });
    }

    let len = binary::read_u32_le_at(buf, 0)? as usize;
    let consumed = len
        .checked_mul(2)
        .and_then(|byte_len| byte_len.checked_add(4))
        .ok_or_else(|| XlsbError::Encoding("wide string length overflow".to_string()))?;
    if buf.len() < consumed {
        return Err(XlsbError::WideStringLength {
            expected: consumed,
            actual: buf.len(),
        });
    }

    let utf16_data = &buf[4..consumed];

    // Convert UTF-16LE to UTF-8 using encoding_rs
    use encoding_rs::UTF_16LE;
    Ok((UTF_16LE.decode(utf16_data).0.into_owned(), consumed))
}

/// Workbook properties record
#[derive(Debug, Clone)]
pub struct WorkbookPropRecord {
    pub is_date1904: bool,
}

impl WorkbookPropRecord {
    pub fn parse(data: &[u8]) -> XlsbResult<Self> {
        if data.is_empty() {
            return Ok(WorkbookPropRecord { is_date1904: false });
        }

        let flags = data[0];
        let is_date1904 = (flags & 0x01) != 0;

        Ok(WorkbookPropRecord { is_date1904 })
    }
}

/// Bundle sheet record (worksheet metadata)
#[derive(Debug, Clone)]
#[allow(dead_code)]
pub struct BundleSheetRecord {
    pub id: u32,
    pub name: String,
    pub state: u32,
    pub rel_id: Option<String>,
}

impl BundleSheetRecord {
    pub fn parse(data: &[u8]) -> XlsbResult<Self> {
        if data.len() < 12 {
            return Err(XlsbError::InvalidLength {
                expected: 12,
                found: data.len(),
            });
        }

        let state = binary::read_u32_le_at(data, 0)?;
        if state > 2 {
            return Err(XlsbError::Unrecognized {
                typ: "BrtBundleSh hsState".to_string(),
                val: state.to_string(),
            });
        }
        let current_id = binary::read_u32_le_at(data, 4)?;
        let (id, strings_offset) = if (1..=0xFFFF).contains(&current_id) {
            (current_id, 8)
        } else {
            // Excel beta XLSB files have an undocumented extra four bytes
            // before iTabID. This layout is also recognized by Apache POI.
            if data.len() < 16 {
                return Err(XlsbError::Unrecognized {
                    typ: "BrtBundleSh iTabID".to_string(),
                    val: current_id.to_string(),
                });
            }
            let beta_id = binary::read_u32_le_at(data, 8)?;
            if !(1..=0xFFFF).contains(&beta_id) {
                return Err(XlsbError::Unrecognized {
                    typ: "BrtBundleSh iTabID".to_string(),
                    val: format!("current {current_id}, beta {beta_id}"),
                });
            }
            (beta_id, 12)
        };

        let (rel_id, rel_consumed) = if binary::read_u32_le_at(data, strings_offset)? == u32::MAX {
            (None, 4)
        } else {
            let (value, consumed) = wide_str_with_len(&data[strings_offset..])?;
            if value.is_empty() {
                return Err(XlsbError::Unrecognized {
                    typ: "BrtBundleSh strRelID".to_string(),
                    val: "empty relationship ID".to_string(),
                });
            }
            (Some(value), consumed)
        };
        if rel_id.is_none() && state != 2 {
            return Err(XlsbError::Unrecognized {
                typ: "BrtBundleSh strRelID".to_string(),
                val: "NULL relationship on a sheet that is not very hidden".to_string(),
            });
        }
        let name_offset = strings_offset.checked_add(rel_consumed).ok_or_else(|| {
            XlsbError::Encoding("BrtBundleSh relationship size overflow".to_string())
        })?;
        let (name, name_consumed) = wide_str_with_len(&data[name_offset..])?;
        if name_offset + name_consumed != data.len() {
            return Err(XlsbError::Unrecognized {
                typ: "BrtBundleSh".to_string(),
                val: format!(
                    "{} trailing bytes",
                    data.len() - name_offset - name_consumed
                ),
            });
        }
        let name_len = name.encode_utf16().count();
        if name_len == 0
            || name_len > 31
            || name.contains(['\0', '\u{0003}', ':', '\\', '*', '?', '/', '[', ']'])
            || name.starts_with('\'')
            || name.ends_with('\'')
        {
            return Err(XlsbError::Unrecognized {
                typ: "BrtBundleSh strName".to_string(),
                val: name,
            });
        }

        Ok(BundleSheetRecord {
            id,
            name,
            state,
            rel_id,
        })
    }
}

/// Row header record
#[derive(Debug, Clone)]
#[allow(dead_code)]
pub struct RowHeaderRecord {
    pub row: u32,
    pub first_col: u16,
    pub last_col: u16,
}

#[allow(dead_code)]
impl RowHeaderRecord {
    pub fn parse(data: &[u8]) -> XlsbResult<Self> {
        if data.len() < 8 {
            return Err(XlsbError::InvalidLength {
                expected: 8,
                found: data.len(),
            });
        }

        let row = binary::read_u32_le_at(data, 0)?;
        let first_col = binary::read_u16_le_at(data, 4)?;
        let last_col = binary::read_u16_le_at(data, 6)?;

        Ok(RowHeaderRecord {
            row,
            first_col,
            last_col,
        })
    }
}

/// Cell value types
#[derive(Debug, Clone)]
pub enum CellValue {
    Blank,
    Bool(bool),
    Error(u8),
    Real(f64),
    String(String),
    Isst(u32), // Index into shared string table
    Formula {
        value: Box<CellValue>,
        formula: Option<Vec<u8>>, // Raw formula bytes
    },
}

/// Cell record base
#[derive(Debug, Clone)]
pub struct CellRecord {
    pub row: u32,
    pub col: u16,
    pub value: CellValue,
}

impl CellRecord {
    pub fn parse(record_type: u16, data: &[u8]) -> XlsbResult<Self> {
        if data.len() < 4 {
            return Err(XlsbError::InvalidLength {
                expected: 4,
                found: data.len(),
            });
        }

        let row = binary::read_u32_le_at(data, 0)?;
        let col = binary::read_u16_le_at(data, 4)?;

        let value = match record_type {
            record_types::CELL_BLANK => CellValue::Blank,
            record_types::CELL_BOOL => {
                if data.len() < 7 {
                    return Err(XlsbError::InvalidLength {
                        expected: 7,
                        found: data.len(),
                    });
                }
                CellValue::Bool(data[6] != 0)
            },
            record_types::CELL_ERROR => {
                if data.len() < 7 {
                    return Err(XlsbError::InvalidLength {
                        expected: 7,
                        found: data.len(),
                    });
                }
                CellValue::Error(data[6])
            },
            record_types::CELL_REAL => {
                if data.len() < 14 {
                    return Err(XlsbError::InvalidLength {
                        expected: 14,
                        found: data.len(),
                    });
                }
                CellValue::Real(binary::read_f64_le_at(data, 6)?)
            },
            record_types::CELL_ST => {
                let mut str_len = 0;
                let string = wide_str(&data[6..], &mut str_len)?;
                CellValue::String(string.to_owned())
            },
            record_types::CELL_ISST => {
                if data.len() < 10 {
                    return Err(XlsbError::InvalidLength {
                        expected: 10,
                        found: data.len(),
                    });
                }
                CellValue::Isst(binary::read_u32_le_at(data, 6)?)
            },
            record_types::CELL_RK => {
                if data.len() < 10 {
                    return Err(XlsbError::InvalidLength {
                        expected: 10,
                        found: data.len(),
                    });
                }
                let rk_value = binary::read_u32_le_at(data, 6)?;
                let real_value = rk_to_f64(rk_value);
                CellValue::Real(real_value)
            },
            // Formula records - parse formula bytes and cached value
            record_types::FMLA_STRING => {
                if data.len() < 10 {
                    return Err(XlsbError::InvalidLength {
                        expected: 10,
                        found: data.len(),
                    });
                }
                // Skip style_id (4 bytes) and flags (1 byte) and formula length (4 bytes)
                let formula_len = binary::read_u32_le_at(data, 6)? as usize;
                if data.len() < 10 + formula_len {
                    return Err(XlsbError::InvalidLength {
                        expected: 10 + formula_len,
                        found: data.len(),
                    });
                }
                let formula_bytes = data[10..10 + formula_len].to_vec();

                // Read cached string value after formula
                let mut str_len = 0;
                let string = wide_str(&data[10 + formula_len..], &mut str_len)?;
                CellValue::Formula {
                    value: Box::new(CellValue::String(string)),
                    formula: Some(formula_bytes),
                }
            },
            record_types::FMLA_NUM => {
                if data.len() < 18 {
                    return Err(XlsbError::InvalidLength {
                        expected: 18,
                        found: data.len(),
                    });
                }
                let formula_len = binary::read_u32_le_at(data, 6)? as usize;
                if data.len() < 10 + formula_len + 8 {
                    return Err(XlsbError::InvalidLength {
                        expected: 10 + formula_len + 8,
                        found: data.len(),
                    });
                }
                let formula_bytes = data[10..10 + formula_len].to_vec();
                let num_value = binary::read_f64_le_at(data, 10 + formula_len)?;
                CellValue::Formula {
                    value: Box::new(CellValue::Real(num_value)),
                    formula: Some(formula_bytes),
                }
            },
            record_types::FMLA_BOOL => {
                if data.len() < 11 {
                    return Err(XlsbError::InvalidLength {
                        expected: 11,
                        found: data.len(),
                    });
                }
                let formula_len = binary::read_u32_le_at(data, 6)? as usize;
                if data.len() < 10 + formula_len + 1 {
                    return Err(XlsbError::InvalidLength {
                        expected: 10 + formula_len + 1,
                        found: data.len(),
                    });
                }
                let formula_bytes = data[10..10 + formula_len].to_vec();
                let bool_value = data[10 + formula_len] != 0;
                CellValue::Formula {
                    value: Box::new(CellValue::Bool(bool_value)),
                    formula: Some(formula_bytes),
                }
            },
            record_types::FMLA_ERROR => {
                if data.len() < 11 {
                    return Err(XlsbError::InvalidLength {
                        expected: 11,
                        found: data.len(),
                    });
                }
                let formula_len = binary::read_u32_le_at(data, 6)? as usize;
                if data.len() < 10 + formula_len + 1 {
                    return Err(XlsbError::InvalidLength {
                        expected: 10 + formula_len + 1,
                        found: data.len(),
                    });
                }
                let formula_bytes = data[10..10 + formula_len].to_vec();
                let error_code = data[10 + formula_len];
                CellValue::Formula {
                    value: Box::new(CellValue::Error(error_code)),
                    formula: Some(formula_bytes),
                }
            },
            _ => return Err(XlsbError::InvalidRecordType(record_type)),
        };

        Ok(CellRecord { row, col, value })
    }
}

/// Convert RK value to f64 (same as XLS)
pub fn rk_to_f64(rk: u32) -> f64 {
    let d100 = (rk & 0x02) != 0;
    let is_int = (rk & 0x01) != 0;

    let value = if is_int {
        let int_val = (rk >> 2) as i32;
        if d100 {
            if int_val % 100 != 0 {
                int_val as f64 / 100.0
            } else {
                (int_val / 100) as f64
            }
        } else {
            int_val as f64
        }
    } else {
        // Float value - reconstruct from 30 bits
        let mut float_bits = [0u8; 8];
        float_bits[0..4].copy_from_slice(&(rk & 0xFFFFFFFC).to_le_bytes());
        f64::from_le_bytes(float_bits)
    };

    if d100 && !is_int {
        value / 100.0
    } else {
        value
    }
}

/// Column information record
#[allow(dead_code)]
#[derive(Debug, Clone)]
pub struct ColInfoRecord {
    pub first_col: u32,
    pub last_col: u32,
    pub width: f64,
    pub style_xf: u32,
    pub custom_width: bool,
    pub hidden: bool,
    pub best_fit: bool,
}

impl ColInfoRecord {
    #[allow(dead_code)]
    pub fn parse(data: &[u8]) -> XlsbResult<Self> {
        if data.len() < 12 {
            return Err(XlsbError::InvalidLength {
                expected: 12,
                found: data.len(),
            });
        }

        let first_col = binary::read_u32_le_at(data, 0)?;
        let last_col = binary::read_u32_le_at(data, 4)?;
        // Width is stored as 256ths of a character
        let width_raw = binary::read_u32_le_at(data, 8)?;
        let width = width_raw as f64 / 256.0;

        let style_xf = if data.len() >= 16 {
            binary::read_u32_le_at(data, 12)?
        } else {
            0
        };

        let flags = if data.len() >= 18 {
            binary::read_u16_le_at(data, 16)?
        } else {
            0
        };

        let custom_width = (flags & 0x0002) != 0;
        let hidden = (flags & 0x0001) != 0;
        let best_fit = (flags & 0x0008) != 0;

        Ok(ColInfoRecord {
            first_col,
            last_col,
            width,
            style_xf,
            custom_width,
            hidden,
            best_fit,
        })
    }
}

/// Merged cell record
#[allow(dead_code)]
#[derive(Debug, Clone)]
pub struct MergeCellRecord {
    pub row_first: u32,
    pub row_last: u32,
    pub col_first: u32,
    pub col_last: u32,
}

impl MergeCellRecord {
    #[allow(dead_code)]
    pub fn parse(data: &[u8]) -> XlsbResult<Self> {
        if data.len() < 16 {
            return Err(XlsbError::InvalidLength {
                expected: 16,
                found: data.len(),
            });
        }

        Ok(MergeCellRecord {
            row_first: binary::read_u32_le_at(data, 0)?,
            row_last: binary::read_u32_le_at(data, 4)?,
            col_first: binary::read_u32_le_at(data, 8)?,
            col_last: binary::read_u32_le_at(data, 12)?,
        })
    }
}

/// Hyperlink record
#[allow(dead_code)]
#[derive(Debug, Clone)]
pub struct HyperlinkRecord {
    pub row_first: u32,
    pub row_last: u32,
    pub col_first: u32,
    pub col_last: u32,
    pub r_id: String,
    pub location: Option<String>,
    pub tooltip: Option<String>,
    pub display: Option<String>,
}

impl HyperlinkRecord {
    #[allow(dead_code)]
    pub fn parse(data: &[u8]) -> XlsbResult<Self> {
        if data.len() < 16 {
            return Err(XlsbError::InvalidLength {
                expected: 16,
                found: data.len(),
            });
        }

        let row_first = binary::read_u32_le_at(data, 0)?;
        let row_last = binary::read_u32_le_at(data, 4)?;
        let col_first = binary::read_u32_le_at(data, 8)?;
        let col_last = binary::read_u32_le_at(data, 12)?;

        let mut offset = 16;

        // Read relationship ID
        let (r_id, consumed) = wide_str_with_len(&data[offset..])?;
        offset += consumed;

        // Read location (optional)
        let (location, consumed) = if offset < data.len() {
            let (loc, c) = wide_str_with_len(&data[offset..])?;
            (if loc.is_empty() { None } else { Some(loc) }, c)
        } else {
            (None, 0)
        };
        offset += consumed;

        // Read tooltip (optional)
        let (tooltip, consumed) = if offset < data.len() {
            let (tt, c) = wide_str_with_len(&data[offset..])?;
            (if tt.is_empty() { None } else { Some(tt) }, c)
        } else {
            (None, 0)
        };
        offset += consumed;

        // Read display text (optional)
        let display = if offset < data.len() {
            let (disp, _) = wide_str_with_len(&data[offset..])?;
            if disp.is_empty() { None } else { Some(disp) }
        } else {
            None
        };

        Ok(HyperlinkRecord {
            row_first,
            row_last,
            col_first,
            col_last,
            r_id,
            location,
            tooltip,
            display,
        })
    }
}

/// Named range record
#[allow(dead_code)]
#[derive(Debug, Clone)]
pub struct NameRecord {
    pub name: String,
    pub formula: Option<Vec<u8>>,
    pub sheet_id: Option<u32>,
    pub hidden: bool,
    pub function: bool,
}

impl NameRecord {
    #[allow(dead_code)]
    pub fn parse(data: &[u8]) -> XlsbResult<Self> {
        let named_range = crate::xlsb::named_ranges::NamedRange::parse(data)?;

        Ok(NameRecord {
            name: named_range.name,
            formula: named_range.formula,
            sheet_id: named_range.sheet_id,
            hidden: named_range.hidden,
            function: named_range.function,
        })
    }
}

/// Record iterator for XLSB parsing
pub struct RecordIter<R> {
    reader: R,
    buffer: [u8; 1],
}

impl<R: Read> RecordIter<R> {
    pub fn new(reader: R) -> Self {
        RecordIter {
            reader,
            buffer: [0],
        }
    }

    pub fn from_cursor(cursor: std::io::Cursor<&[u8]>) -> RecordIter<std::io::Cursor<&[u8]>> {
        RecordIter::new(cursor)
    }

    fn read_u8(&mut self) -> Result<u8, std::io::Error> {
        self.reader.read_exact(&mut self.buffer)?;
        Ok(self.buffer[0])
    }

    /// Read next type, until we have no future record
    pub fn read_type(&mut self) -> Result<u16, std::io::Error> {
        let b = self.read_u8()?;
        let typ = if (b & 0x80) == 0x80 {
            (b & 0x7F) as u16 + (((self.read_u8()? & 0x7F) as u16) << 7)
        } else {
            b as u16
        };
        Ok(typ)
    }

    pub fn fill_buffer(&mut self, buf: &mut Vec<u8>) -> Result<usize, std::io::Error> {
        let mut b = self.read_u8()?;
        let mut len = (b & 0x7F) as usize;
        for i in 1..4 {
            if (b & 0x80) == 0 {
                break;
            }
            b = self.read_u8()?;
            len += ((b & 0x7F) as usize) << (7 * i);
        }
        buf.resize(len, 0);

        self.reader.read_exact(&mut buf[..len])?;
        Ok(len)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn wide_string(value: &str) -> Vec<u8> {
        let mut data = (value.encode_utf16().count() as u32).to_le_bytes().to_vec();
        for unit in value.encode_utf16() {
            data.extend_from_slice(&unit.to_le_bytes());
        }
        data
    }

    #[test]
    fn parses_current_and_excel_beta_bundle_sheets() {
        let mut current = 0u32.to_le_bytes().to_vec();
        current.extend_from_slice(&7u32.to_le_bytes());
        current.extend_from_slice(&wide_string("rId7"));
        current.extend_from_slice(&wide_string("Data"));
        let sheet = BundleSheetRecord::parse(&current).unwrap();
        assert_eq!(sheet.id, 7);
        assert_eq!(sheet.state, 0);
        assert_eq!(sheet.rel_id.as_deref(), Some("rId7"));
        assert_eq!(sheet.name, "Data");

        let mut beta = 0u64.to_le_bytes().to_vec();
        beta.extend_from_slice(&8u32.to_le_bytes());
        beta.extend_from_slice(&wide_string("rId8"));
        beta.extend_from_slice(&wide_string("Legacy"));
        let sheet = BundleSheetRecord::parse(&beta).unwrap();
        assert_eq!(sheet.id, 8);
        assert_eq!(sheet.rel_id.as_deref(), Some("rId8"));
        assert_eq!(sheet.name, "Legacy");
    }

    #[test]
    fn rejects_malformed_bundle_sheet_metadata() {
        let mut invalid = 0u32.to_le_bytes().to_vec();
        invalid.extend_from_slice(&0u32.to_le_bytes());
        invalid.extend_from_slice(&0u32.to_le_bytes());
        assert!(matches!(
            BundleSheetRecord::parse(&invalid),
            Err(XlsbError::Unrecognized { .. })
        ));

        let mut null_visible = 0u32.to_le_bytes().to_vec();
        null_visible.extend_from_slice(&1u32.to_le_bytes());
        null_visible.extend_from_slice(&u32::MAX.to_le_bytes());
        null_visible.extend_from_slice(&wide_string("Module"));
        assert!(matches!(
            BundleSheetRecord::parse(&null_visible),
            Err(XlsbError::Unrecognized { .. })
        ));
    }
}
