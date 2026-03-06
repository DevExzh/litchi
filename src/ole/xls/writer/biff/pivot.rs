//! Pivot table BIFF8 record writers.
//!
//! Writes the family of SX* records that define pivot table structures:
//!
//! - **SXVIEW** (0x00B0): View definition — the main pivot table header.
//! - **SXVD** (0x00B1): View field — describes a single field (dimension).
//! - **SXVI** (0x00B2): View item — a single item within a field.
//! - **SXDI** (0x00C5): Data item — describes a data field (value area).
//! - **SXVS** (0x00E3): View source — source type of the pivot cache.
//! - **SXPI** (0x00B6): Page item — page field entries.
//!
//! # References
//!
//! - MS-XLS sections 2.4.271–2.4.283
//! - Reader counterpart: `crate::ole::xls::pivot_table`

use crate::ole::xls::XlsResult;
use std::io::Write;

use super::write_record_header;

// ---------------------------------------------------------------------------
// XLUnicodeStringNoCch encoder (shared helper)
// ---------------------------------------------------------------------------

/// Encode a string as XLUnicodeStringNoCch: `[flags: u8][chars...]`.
///
/// Uses compressed Latin-1 for ASCII strings, UTF-16LE otherwise.
/// Returns an **empty** `Vec` for an empty string (no flags byte),
/// matching the reader's `read_xl_string_no_cch` which returns
/// immediately when `cch == 0`.
fn encode_xl_string_no_cch(s: &str) -> Vec<u8> {
    if s.is_empty() {
        return Vec::new();
    }
    if s.is_ascii() {
        let mut buf = Vec::with_capacity(1 + s.len());
        buf.push(0x00); // flags: compressed
        buf.extend_from_slice(s.as_bytes());
        buf
    } else {
        let utf16: Vec<u16> = s.encode_utf16().collect();
        let mut buf = Vec::with_capacity(1 + utf16.len() * 2);
        buf.push(0x01); // flags: UTF-16LE
        for ch in &utf16 {
            buf.extend_from_slice(&ch.to_le_bytes());
        }
        buf
    }
}

// ---------------------------------------------------------------------------
// SXVS — View Source (0x00E3)
// ---------------------------------------------------------------------------

/// Write an SXVS record (2 bytes: source type).
///
/// Source type values:
/// - `0x0001` — Worksheet
/// - `0x0002` — External
/// - `0x0004` — Consolidation
/// - `0x0010` — Scenario
pub fn write_sxvs<W: Write>(writer: &mut W, source_type: u16) -> XlsResult<()> {
    write_record_header(writer, 0x00E3, 2)?;
    writer.write_all(&source_type.to_le_bytes())?;
    Ok(())
}

// ---------------------------------------------------------------------------
// SXVIEW — View Definition (0x00B0)
// ---------------------------------------------------------------------------

/// Configuration for writing an SXVIEW record.
pub struct SxViewConfig<'a> {
    pub first_row: u16,
    pub last_row: u16,
    pub first_col: u16,
    pub last_col: u16,
    pub first_header_row: u16,
    pub first_data_row: u16,
    pub first_data_col: u16,
    /// Pivot cache index (0-based). Links this view to a PIVOTCACHE.
    pub cache_index: u16,
    /// Axis for the data field header (0=none, 1=row, 2=col, 4=page).
    pub data_axis: u16,
    /// Position of data field label within the axis (-1 = end).
    pub data_position: u16,
    /// Total number of fields.
    pub field_count: u16,
    pub row_field_count: u16,
    pub col_field_count: u16,
    pub page_field_count: u16,
    pub data_field_count: u16,
    /// Number of visible data rows.
    pub data_row_count: u16,
    /// Number of visible data columns.
    pub data_col_count: u16,
    /// Option flags: bit 0 = fRwGrand, bit 1 = fColGrand, bit 3 = fAutoFormat.
    pub flags: u16,
    /// Auto-format index (0 = none).
    pub auto_format_index: u16,
    /// Pivot table name.
    pub name: &'a str,
    /// Data field header name (e.g. "Values").
    pub data_field_name: &'a str,
}

/// Write an SXVIEW record.
///
/// Layout (per MS-XLS §2.4.293, verified against §3.10.7 example):
/// ```text
///  0  u16  rwFirst
///  2  u16  rwLast
///  4  u16  colFirst
///  6  u16  colLast
///  8  u16  rwFirstHead
/// 10  u16  rwFirstData
/// 12  u16  colFirstData
/// 14  u16  iCache          — pivot cache index
/// 16  u16  reserved        — 0
/// 18  u16  sxaxis4Data
/// 20  u16  ipos4Data
/// 22  u16  cDim            — total fields
/// 24  u16  cDimRw
/// 26  u16  cDimCol
/// 28  u16  cDimPg
/// 30  u16  cDimData
/// 32  u16  cRw             — data row count
/// 34  u16  cCol            — data column count
/// 36  u16  grbit           — flags
/// 38  u16  itblAutoFmt
/// 40  u16  cchTableName
/// 42  u16  cchDataName
/// 44  var  stTable         — XLUnicodeStringNoCch
///     var  stData          — XLUnicodeStringNoCch
/// ```
pub fn write_sxview<W: Write>(writer: &mut W, cfg: &SxViewConfig<'_>) -> XlsResult<()> {
    let name_bytes = encode_xl_string_no_cch(cfg.name);
    let data_name_bytes = encode_xl_string_no_cch(cfg.data_field_name);

    let cch_name = cfg.name.chars().count() as u16;
    let cch_data = cfg.data_field_name.chars().count() as u16;

    // Fixed header: 44 bytes + variable name strings
    let data_len = 44u16 + name_bytes.len() as u16 + data_name_bytes.len() as u16;

    write_record_header(writer, 0x00B0, data_len)?;

    writer.write_all(&cfg.first_row.to_le_bytes())?; //  0: rwFirst
    writer.write_all(&cfg.last_row.to_le_bytes())?; //  2: rwLast
    writer.write_all(&cfg.first_col.to_le_bytes())?; //  4: colFirst
    writer.write_all(&cfg.last_col.to_le_bytes())?; //  6: colLast
    writer.write_all(&cfg.first_header_row.to_le_bytes())?; //  8: rwFirstHead
    writer.write_all(&cfg.first_data_row.to_le_bytes())?; // 10: rwFirstData
    writer.write_all(&cfg.first_data_col.to_le_bytes())?; // 12: colFirstData
    writer.write_all(&cfg.cache_index.to_le_bytes())?; // 14: iCache
    writer.write_all(&0u16.to_le_bytes())?; // 16: reserved
    writer.write_all(&cfg.data_axis.to_le_bytes())?; // 18: sxaxis4Data
    writer.write_all(&cfg.data_position.to_le_bytes())?; // 20: ipos4Data
    writer.write_all(&cfg.field_count.to_le_bytes())?; // 22: cDim
    writer.write_all(&cfg.row_field_count.to_le_bytes())?; // 24: cDimRw
    writer.write_all(&cfg.col_field_count.to_le_bytes())?; // 26: cDimCol
    writer.write_all(&cfg.page_field_count.to_le_bytes())?; // 28: cDimPg
    writer.write_all(&cfg.data_field_count.to_le_bytes())?; // 30: cDimData
    writer.write_all(&cfg.data_row_count.to_le_bytes())?; // 32: cRw
    writer.write_all(&cfg.data_col_count.to_le_bytes())?; // 34: cCol
    writer.write_all(&cfg.flags.to_le_bytes())?; // 36: grbit
    writer.write_all(&cfg.auto_format_index.to_le_bytes())?; // 38: itblAutoFmt
    writer.write_all(&cch_name.to_le_bytes())?; // 40: cchTableName
    writer.write_all(&cch_data.to_le_bytes())?; // 42: cchDataName

    writer.write_all(&name_bytes)?;
    writer.write_all(&data_name_bytes)?;

    Ok(())
}

// ---------------------------------------------------------------------------
// SXVD — View Field (0x00B1)
// ---------------------------------------------------------------------------

/// Configuration for writing an SXVD record.
pub struct SxVdConfig<'a> {
    /// Axis: 0=none, 1=row, 2=col, 4=page, 8=data.
    pub axis: u16,
    /// Number of subtotals.
    pub subtotal_count: u16,
    /// Subtotal function bitmask.
    pub subtotal_flags: u16,
    /// Number of items in this field.
    pub item_count: u16,
    /// Optional field name override (`None` → use source name).
    pub name: Option<&'a str>,
}

/// Write an SXVD record.
pub fn write_sxvd<W: Write>(writer: &mut W, cfg: &SxVdConfig<'_>) -> XlsResult<()> {
    let (cch_name, name_bytes) = match cfg.name {
        Some(n) => {
            let bytes = encode_xl_string_no_cch(n);
            (n.chars().count() as u16, bytes)
        },
        None => (0xFFFFu16, Vec::new()),
    };

    let data_len = 10u16 + name_bytes.len() as u16;

    write_record_header(writer, 0x00B1, data_len)?;
    writer.write_all(&cfg.axis.to_le_bytes())?; // 0
    writer.write_all(&cfg.subtotal_count.to_le_bytes())?; // 2
    writer.write_all(&cfg.subtotal_flags.to_le_bytes())?; // 4
    writer.write_all(&cfg.item_count.to_le_bytes())?; // 6
    writer.write_all(&cch_name.to_le_bytes())?; // 8
    writer.write_all(&name_bytes)?;

    Ok(())
}

// ---------------------------------------------------------------------------
// SXVI — View Item (0x00B2)
// ---------------------------------------------------------------------------

/// Configuration for writing an SXVI record.
pub struct SxViConfig<'a> {
    /// Item type: 0x0000=Data, 0x0001=Default subtotal, 0x0002=Sum, etc.
    pub item_type: u16,
    /// Option flags.
    pub flags: u16,
    /// Cache index.
    pub cache_index: u16,
    /// Optional item name override.
    pub name: Option<&'a str>,
}

/// Write an SXVI record.
pub fn write_sxvi<W: Write>(writer: &mut W, cfg: &SxViConfig<'_>) -> XlsResult<()> {
    let (cch_name, name_bytes) = match cfg.name {
        Some(n) => {
            let bytes = encode_xl_string_no_cch(n);
            (n.chars().count() as u16, bytes)
        },
        None => (0xFFFFu16, Vec::new()),
    };

    let data_len = 8u16 + name_bytes.len() as u16;

    write_record_header(writer, 0x00B2, data_len)?;
    writer.write_all(&cfg.item_type.to_le_bytes())?; // 0
    writer.write_all(&cfg.flags.to_le_bytes())?; // 2
    writer.write_all(&cfg.cache_index.to_le_bytes())?; // 4
    writer.write_all(&cch_name.to_le_bytes())?; // 6
    writer.write_all(&name_bytes)?;

    Ok(())
}

// ---------------------------------------------------------------------------
// SXDI — Data Item (0x00C5)
// ---------------------------------------------------------------------------

/// Configuration for writing an SXDI record.
pub struct SxDiConfig<'a> {
    /// Index of the source field in the pivot cache.
    pub source_field_index: u16,
    /// Aggregation function: 0=Sum,1=Count,2=Average,3=Max,4=Min,5=Product,...
    pub function: u16,
    /// Display format flags.
    pub display_format: u16,
    /// Base field index (for "show values as").
    pub base_field_index: u16,
    /// Base item index.
    pub base_item_index: u16,
    /// Number format index.
    pub num_format_index: u16,
    /// Optional name override.
    pub name: &'a str,
}

/// Write an SXDI record.
///
/// When `cfg.name` is empty, `cchName` is set to `0xFFFF` (not present),
/// matching the convention used by SXVD / SXVI.
pub fn write_sxdi<W: Write>(writer: &mut W, cfg: &SxDiConfig<'_>) -> XlsResult<()> {
    let (cch_name, name_bytes) = if cfg.name.is_empty() {
        (0xFFFFu16, Vec::new())
    } else {
        let bytes = encode_xl_string_no_cch(cfg.name);
        (cfg.name.chars().count() as u16, bytes)
    };

    let data_len = 14u16 + name_bytes.len() as u16;

    write_record_header(writer, 0x00C5, data_len)?;
    writer.write_all(&cfg.source_field_index.to_le_bytes())?; // 0
    writer.write_all(&cfg.function.to_le_bytes())?; // 2
    writer.write_all(&cfg.display_format.to_le_bytes())?; // 4
    writer.write_all(&cfg.base_field_index.to_le_bytes())?; // 6
    writer.write_all(&cfg.base_item_index.to_le_bytes())?; // 8
    writer.write_all(&cfg.num_format_index.to_le_bytes())?; // 10
    writer.write_all(&cch_name.to_le_bytes())?; // 12
    writer.write_all(&name_bytes)?;

    Ok(())
}

// ---------------------------------------------------------------------------
// SXPI — Page Item (0x00B6)
// ---------------------------------------------------------------------------

/// Write an SXPI record for one or more page field entries.
///
/// Each entry is 6 bytes per LO `XclPTPageFieldInfo`:
///   `(mnField: u16, mnSelItem: u16, mnObjId: u16)`
///
/// The tuple format is `(item_index, field_index, object_id)` in the
/// public API, but the BIFF wire order is field-first.
pub fn write_sxpi<W: Write>(writer: &mut W, entries: &[(u16, u16, u16)]) -> XlsResult<()> {
    if entries.is_empty() {
        return Ok(());
    }
    let data_len = (entries.len() * 6) as u16;
    write_record_header(writer, 0x00B6, data_len)?;
    for &(item_idx, field_idx, obj_id) in entries {
        writer.write_all(&field_idx.to_le_bytes())?; // mnField
        writer.write_all(&item_idx.to_le_bytes())?; // mnSelItem
        writer.write_all(&obj_id.to_le_bytes())?; // mnObjId
    }
    Ok(())
}

// ---------------------------------------------------------------------------
// SxStreamID — Pivot Cache Stream Identifier (0x00D5)
// ---------------------------------------------------------------------------

/// Write an SxStreamID record in the **globals** substream.
///
/// This links a pivot view to its cache stream. The `id_stm` value
/// is the zero-based index of the pivot cache (typically 0 for the
/// first pivot table).
pub fn write_sx_stream_id<W: Write>(writer: &mut W, id_stm: u16) -> XlsResult<()> {
    write_record_header(writer, 0x00D5, 2)?;
    writer.write_all(&id_stm.to_le_bytes())?;
    Ok(())
}

// ---------------------------------------------------------------------------
// SxEx — Extended PivotTable View (0x00F1)
// ---------------------------------------------------------------------------

/// Configuration for writing an SXEX record.
pub struct SxExConfig {
    /// Number of SxFormat records (usually 0).
    pub sx_format_count: u16,
    /// Number of selected items (usually 0).
    pub sx_select_count: u16,
    /// Rows used for page fields (usually 0).
    pub page_rows: u16,
    /// Columns used for page fields (usually 1 when page fields exist).
    pub page_cols: u16,
    /// Flags (u32 at offset 14 per XclPTExtInfo.mnFlags).
    ///
    /// Default: `EXC_SXEX_DEFAULTFLAGS = 0x004F0200` (per LO `xlpivot.hxx`).
    /// Includes fEnableDrilldown(bit17) and other standard flags.
    pub flags: u32,
}

impl Default for SxExConfig {
    fn default() -> Self {
        Self {
            sx_format_count: 0,
            sx_select_count: 0,
            page_rows: 0,
            page_cols: 0,
            // EXC_SXEX_DEFAULTFLAGS from LO xlpivot.hxx — includes drilldown,
            // wizard, and other required bits.
            flags: 0x004F_0200,
        }
    }
}

/// Write an SxEx record (Extended PivotTable View properties).
///
/// Layout (per MS-XLS §2.4.282, 24 bytes fixed when all cch* = 0xFFFF):
/// ```text
///  0  u16  csxformat
///  2  u16  cchErrorString   — 0xFFFF = not set
///  4  u16  cchNullString    — 0xFFFF = not set
///  6  u16  cchTag           — 0xFFFF = not set
///  8  u16  csxselect
/// 10  u16  crwPage
/// 12  u16  ccolPage
/// 14  u32  mnFlags          — XclPTExtInfo flags (default 0x004F0200)
/// 18  u16  cchPageFieldStyle — 0xFFFF = not set
/// 20  u16  cchTableStyle     — 0xFFFF = not set
/// 22  u16  cchVacateStyle    — 0xFFFF = not set
/// ```
pub fn write_sxex<W: Write>(writer: &mut W, cfg: &SxExConfig) -> XlsResult<()> {
    write_record_header(writer, 0x00F1, 24)?;
    writer.write_all(&cfg.sx_format_count.to_le_bytes())?; //  0: csxformat
    writer.write_all(&0xFFFFu16.to_le_bytes())?; //  2: cchErrorString
    writer.write_all(&0xFFFFu16.to_le_bytes())?; //  4: cchNullString
    writer.write_all(&0xFFFFu16.to_le_bytes())?; //  6: cchTag
    writer.write_all(&cfg.sx_select_count.to_le_bytes())?; //  8: csxselect
    writer.write_all(&cfg.page_rows.to_le_bytes())?; // 10: crwPage
    writer.write_all(&cfg.page_cols.to_le_bytes())?; // 12: ccolPage
    writer.write_all(&cfg.flags.to_le_bytes())?; // 14: mnFlags (u32!)
    writer.write_all(&0xFFFFu16.to_le_bytes())?; // 18: cchPageFieldStyle
    writer.write_all(&0xFFFFu16.to_le_bytes())?; // 20: cchTableStyle
    writer.write_all(&0xFFFFu16.to_le_bytes())?; // 22: cchVacateStyle
    Ok(())
}

// ---------------------------------------------------------------------------
// QsiSxTag — Query/Pivot Table Tag (0x0802)
// ---------------------------------------------------------------------------

/// Write a QsiSxTag record.
///
/// This record is always written by LO after SxEx. Its layout
/// (per `XclExpPivotTable::WriteQsiSxTag` in `xepivot.cxx`):
/// ```text
///  0  u16  rt            = 0x0802
///  2  u16  dummyFlags    = 0x0000
///  4  u16  tableType     = 1 (pivot table)
///  6  u16  flags         = 0x0001 (bEnableRefresh)
///  8  u32  options       = 0x00000002
/// 12  u8   verLastRefreshed  = 2
/// 13  u8   verMinRefresh     = 0
/// 14  u8   nOffsetBytes      = 16
/// 15  u8   verFirstCreated   = 0
/// 16  var  XclExpString(tableName)
///      u16  unknown       = 0x0100
/// ```
fn write_qsi_sx_tag<W: Write>(writer: &mut W, table_name: &str) -> XlsResult<()> {
    let name_bytes = encode_xl_unicode_string(table_name);
    let data_len = 16u16 + name_bytes.len() as u16 + 2;

    write_record_header(writer, 0x0802, data_len)?;
    writer.write_all(&0x0802u16.to_le_bytes())?;
    writer.write_all(&0x0000u16.to_le_bytes())?;
    writer.write_all(&0x0001u16.to_le_bytes())?;
    writer.write_all(&0x0001u16.to_le_bytes())?;
    writer.write_all(&0x0000_0002u32.to_le_bytes())?;
    writer.write_all(&[0x02u8])?;
    writer.write_all(&[0x00u8])?;
    writer.write_all(&[16u8])?;
    writer.write_all(&[0u8])?;
    writer.write_all(&name_bytes)?;
    writer.write_all(&0x0100u16.to_le_bytes())?;
    Ok(())
}

fn write_sxaddl_record<W: Write>(
    writer: &mut W,
    sxc: u8,
    sxd: u8,
    payload: &[u8],
) -> XlsResult<()> {
    write_record_header(writer, 0x0864, (6 + payload.len()) as u16)?;
    writer.write_all(&0x0864u16.to_le_bytes())?;
    writer.write_all(&0x0000u16.to_le_bytes())?;
    writer.write_all(&[sxc, sxd])?;
    writer.write_all(payload)?;
    Ok(())
}

fn write_sxaddl_name_record<W: Write>(writer: &mut W, sxc: u8, name: &str) -> XlsResult<()> {
    let mut payload = Vec::with_capacity(6 + name.len().saturating_mul(2));
    payload.extend_from_slice(&(name.chars().count() as u32).to_le_bytes());
    payload.extend_from_slice(&0u16.to_le_bytes());
    payload.extend_from_slice(&encode_xl_unicode_string(name));
    write_sxaddl_record(writer, sxc, 0x00, &payload)
}

fn write_sxaddl_style_name<W: Write>(writer: &mut W) -> XlsResult<()> {
    static PAYLOAD: &[u8] = &[
        0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x33, 0x00, 0x11, 0x00, 0x50, 0x00, 0x69, 0x00, 0x76,
        0x00, 0x6F, 0x00, 0x74, 0x00, 0x53, 0x00, 0x74, 0x00, 0x79, 0x00, 0x6C, 0x00, 0x65, 0x00,
        0x4C, 0x00, 0x69, 0x00, 0x67, 0x00, 0x68, 0x00, 0x74, 0x00, 0x31, 0x00, 0x36, 0x00,
    ];
    write_sxaddl_record(writer, 0x00, 0x1E, PAYLOAD)
}

pub fn write_pivot_modern_extensions<W: Write>(
    writer: &mut W,
    table_name: &str,
    field_names: &[&str],
) -> XlsResult<()> {
    static VIEW_VER10_INFO: &[u8] = &[0x08, 0x41, 0x40, 0x00, 0x00, 0x00];
    static VIEW_VER12_INFO: &[u8] = &[0x9F, 0x00, 0x40, 0x00, 0x00, 0x00];
    static VIEW_FLAG_A: &[u8] = &[0x02, 0x00, 0x00, 0x00, 0x00, 0x00];
    static VIEW_FLAG_B: &[u8] = &[0xFF, 0x00, 0x00, 0x00, 0x00, 0x00];
    static FIELD_VER12_INFO: &[u8] = &[0x28, 0x00, 0x00, 0x00, 0x00, 0x00];
    static FIELD_FLAG_A: &[u8] = &[0x02, 0x00, 0x00, 0x00, 0x00, 0x00];
    static FIELD_FLAG_B: &[u8] = &[0xFF, 0x00, 0x00, 0x00, 0x00, 0x00];
    static END_PAYLOAD: &[u8] = &[0x00, 0x00, 0x00, 0x00, 0x00, 0x00];

    write_qsi_sx_tag(writer, table_name)?;
    write_sx_view_ex9(writer)?;

    write_sxaddl_name_record(writer, 0x00, table_name)?;
    write_sxaddl_record(writer, 0x00, 0x02, VIEW_VER10_INFO)?;
    write_sxaddl_record(writer, 0x00, 0x19, VIEW_VER12_INFO)?;

    for field_name in field_names {
        write_sxaddl_name_record(writer, 0x17, field_name)?;
        write_sxaddl_record(writer, 0x17, 0x19, FIELD_VER12_INFO)?;
        write_sxaddl_record(writer, 0x17, 0x01, FIELD_FLAG_A)?;
        write_sxaddl_record(writer, 0x17, 0x01, FIELD_FLAG_B)?;
        write_sxaddl_record(writer, 0x17, 0xFF, END_PAYLOAD)?;
    }

    write_sxaddl_style_name(writer)?;
    write_sxaddl_record(writer, 0x00, 0x01, VIEW_FLAG_A)?;
    write_sxaddl_record(writer, 0x00, 0x01, VIEW_FLAG_B)?;
    write_sxaddl_record(writer, 0x00, 0xFF, END_PAYLOAD)?;
    Ok(())
}

// ---------------------------------------------------------------------------
// SxViewEx9 — Extended Pivot Table View (0x0810)
// ---------------------------------------------------------------------------

/// Write an SxViewEx9 record (pivot table autoformat / grid layout).
///
/// This FRT record is written by Excel and LibreOffice (when grid layout is
/// active) immediately after `QsiSxTag`.
///
/// Layout (matching the Excel-authored reference worksheet, 17 bytes):
/// ```text
///  0  u16  rt              = 0x0810  (FrtHeader record type echo)
///  2  u16  reserved        = 0x0002
///  4  u32  mbReport        = 0
///  8  u32  flags           = 0x24
/// 12  u16  mnAutoFormat    = 1
/// 14  var  grand total name (XclExpString, empty = 3 bytes: 00 00 00)
/// ```
pub fn write_sx_view_ex9<W: Write>(writer: &mut W) -> XlsResult<()> {
    const SX_VIEW_EX9_RECORD_ID: u16 = 0x0810;
    const SX_VIEW_EX9_FRT_FLAGS: u16 = 0x0002;
    const SX_VIEW_EX9_REPORT_FLAGS: u32 = 0;
    const SX_VIEW_EX9_VIEW_FLAGS: u32 = 0x24;
    const SX_VIEW_EX9_AUTOFORMAT: u16 = 1;
    const EMPTY_GRAND_TOTAL_NAME: [u8; 3] = [0, 0, 0];

    write_record_header(writer, SX_VIEW_EX9_RECORD_ID, 17)?;
    writer.write_all(&SX_VIEW_EX9_RECORD_ID.to_le_bytes())?;
    writer.write_all(&SX_VIEW_EX9_FRT_FLAGS.to_le_bytes())?;
    writer.write_all(&SX_VIEW_EX9_REPORT_FLAGS.to_le_bytes())?;
    writer.write_all(&SX_VIEW_EX9_VIEW_FLAGS.to_le_bytes())?;
    writer.write_all(&SX_VIEW_EX9_AUTOFORMAT.to_le_bytes())?;
    writer.write_all(&EMPTY_GRAND_TOTAL_NAME)?;
    Ok(())
}

const MSO_DRAWING_GROUP_RECORD_ID: u16 = 0x00EB;
const MSO_DRAWING_RECORD_ID: u16 = 0x00EC;
const OBJ_RECORD_ID: u16 = 0x005D;

const ESCHER_DGG_CONTAINER: u16 = 0xF000;
const ESCHER_DG_CONTAINER: u16 = 0xF002;
const ESCHER_SPGR_CONTAINER: u16 = 0xF003;
const ESCHER_SP_CONTAINER: u16 = 0xF004;
const ESCHER_DGG: u16 = 0xF006;
const ESCHER_DG: u16 = 0xF008;
const ESCHER_SPGR: u16 = 0xF009;
const ESCHER_SP: u16 = 0xF00A;
const ESCHER_OPT: u16 = 0xF00B;
const ESCHER_CLIENT_ANCHOR: u16 = 0xF010;
const ESCHER_CLIENT_DATA: u16 = 0xF011;
const ESCHER_SPLIT_MENU_COLORS: u16 = 0xF11E;

const OBJ_FT_CMO: u16 = 0x0015;
const OBJ_FT_CBLS: u16 = 0x000C;
const OBJ_FT_LBS_DATA: u16 = 0x0013;

const COMBO_BOX_OBJECT_TYPE: u16 = 0x0014;
const PIVOT_PAGE_OBJECT_ID: u16 = 0x0001;
const PIVOT_PAGE_OBJECT_FLAGS: u16 = 0x2101;
const PIVOT_PAGE_CBLS_RESERVED_PREFIX: [u8; 8] = [0; 8];
const PIVOT_PAGE_CBLS_ACCELERATOR: u16 = 0x0064;
const PIVOT_PAGE_CBLS_STATE: u16 = 0x0001;
const PIVOT_PAGE_CBLS_TEXT_LEN: u16 = 0x000A;
const PIVOT_PAGE_CBLS_RESERVED_A: u16 = 0x0000;
const PIVOT_PAGE_CBLS_RESERVED_B: u16 = 0x0010;
const PIVOT_PAGE_CBLS_RESERVED_C: u16 = 0x0001;
const PIVOT_PAGE_LBS_CB_CONTINUED: u16 = 0x1FEE;
const PIVOT_PAGE_LBS_FLAGS: u16 = 0x0101;
const PIVOT_PAGE_DROPDOWN_STYLE: u16 = 0x0002;
const PIVOT_PAGE_DROPDOWN_LINES: u16 = 0x0008;
const PIVOT_PAGE_DROPDOWN_MIN_WIDTH: u16 = 0x0000;

#[derive(Clone, Copy)]
struct EscherRecordHeader {
    options: u16,
    record_id: u16,
    data_size: u32,
}

#[derive(Clone, Copy)]
struct EscherProperty {
    property_id: u16,
    value: u32,
}

#[derive(Clone, Copy)]
struct DrawingGroupCluster {
    drawing_group_id: u32,
    used_shape_ids: u32,
}

#[derive(Clone, Copy)]
struct ObjSubrecordHeader {
    sid: u16,
    data_size: u16,
}

fn write_escher_record_header<W: Write>(
    writer: &mut W,
    header: EscherRecordHeader,
) -> XlsResult<()> {
    writer.write_all(&header.options.to_le_bytes())?;
    writer.write_all(&header.record_id.to_le_bytes())?;
    writer.write_all(&header.data_size.to_le_bytes())?;
    Ok(())
}

fn write_escher_properties<W: Write>(
    writer: &mut W,
    properties: &[EscherProperty],
) -> XlsResult<()> {
    for property in properties {
        writer.write_all(&property.property_id.to_le_bytes())?;
        writer.write_all(&property.value.to_le_bytes())?;
    }
    Ok(())
}

fn write_obj_subrecord_header<W: Write>(
    writer: &mut W,
    header: ObjSubrecordHeader,
) -> XlsResult<()> {
    writer.write_all(&header.sid.to_le_bytes())?;
    writer.write_all(&header.data_size.to_le_bytes())?;
    Ok(())
}

fn write_ft_cmo_combo_box<W: Write>(writer: &mut W) -> XlsResult<()> {
    write_obj_subrecord_header(
        writer,
        ObjSubrecordHeader {
            sid: OBJ_FT_CMO,
            data_size: 18,
        },
    )?;
    writer.write_all(&COMBO_BOX_OBJECT_TYPE.to_le_bytes())?;
    writer.write_all(&PIVOT_PAGE_OBJECT_ID.to_le_bytes())?;
    writer.write_all(&PIVOT_PAGE_OBJECT_FLAGS.to_le_bytes())?;
    writer.write_all(&0u32.to_le_bytes())?;
    writer.write_all(&0u32.to_le_bytes())?;
    writer.write_all(&0u32.to_le_bytes())?;
    Ok(())
}

fn write_ft_cbls<W: Write>(writer: &mut W) -> XlsResult<()> {
    write_obj_subrecord_header(
        writer,
        ObjSubrecordHeader {
            sid: OBJ_FT_CBLS,
            data_size: 20,
        },
    )?;
    writer.write_all(&PIVOT_PAGE_CBLS_RESERVED_PREFIX)?;
    writer.write_all(&PIVOT_PAGE_CBLS_ACCELERATOR.to_le_bytes())?;
    writer.write_all(&PIVOT_PAGE_CBLS_STATE.to_le_bytes())?;
    writer.write_all(&PIVOT_PAGE_CBLS_TEXT_LEN.to_le_bytes())?;
    writer.write_all(&PIVOT_PAGE_CBLS_RESERVED_A.to_le_bytes())?;
    writer.write_all(&PIVOT_PAGE_CBLS_RESERVED_B.to_le_bytes())?;
    writer.write_all(&PIVOT_PAGE_CBLS_RESERVED_C.to_le_bytes())?;
    Ok(())
}

fn write_ft_lbs_data<W: Write>(writer: &mut W) -> XlsResult<()> {
    write_obj_subrecord_header(
        writer,
        ObjSubrecordHeader {
            sid: OBJ_FT_LBS_DATA,
            data_size: PIVOT_PAGE_LBS_CB_CONTINUED,
        },
    )?;
    writer.write_all(&0u16.to_le_bytes())?;
    writer.write_all(&0u16.to_le_bytes())?;
    writer.write_all(&0u16.to_le_bytes())?;
    writer.write_all(&PIVOT_PAGE_LBS_FLAGS.to_le_bytes())?;
    writer.write_all(&0u16.to_le_bytes())?;
    writer.write_all(&PIVOT_PAGE_DROPDOWN_STYLE.to_le_bytes())?;
    writer.write_all(&PIVOT_PAGE_DROPDOWN_LINES.to_le_bytes())?;
    writer.write_all(&PIVOT_PAGE_DROPDOWN_MIN_WIDTH.to_le_bytes())?;
    writer.write_all(&0u16.to_le_bytes())?;
    writer.write_all(&[0u8, 0u8])?;
    Ok(())
}

fn write_escher_group_shape<W: Write>(writer: &mut W, shape_id: u32) -> XlsResult<()> {
    write_escher_record_header(
        writer,
        EscherRecordHeader {
            options: 0x000F,
            record_id: ESCHER_SP_CONTAINER,
            data_size: 0x28,
        },
    )?;
    write_escher_record_header(
        writer,
        EscherRecordHeader {
            options: 0x0001,
            record_id: ESCHER_SPGR,
            data_size: 0x10,
        },
    )?;
    for _ in 0..4 {
        writer.write_all(&0u32.to_le_bytes())?;
    }
    write_escher_record_header(
        writer,
        EscherRecordHeader {
            options: 0x0002,
            record_id: ESCHER_SP,
            data_size: 0x08,
        },
    )?;
    writer.write_all(&shape_id.to_le_bytes())?;
    writer.write_all(&0x0000_0005u32.to_le_bytes())?;
    Ok(())
}

fn write_escher_client_anchor<W: Write>(
    writer: &mut W,
    options: u16,
    fields: [u16; 9],
) -> XlsResult<()> {
    write_escher_record_header(
        writer,
        EscherRecordHeader {
            options,
            record_id: ESCHER_CLIENT_ANCHOR,
            data_size: 0x12,
        },
    )?;
    for field in fields {
        writer.write_all(&field.to_le_bytes())?;
    }
    Ok(())
}

// ---------------------------------------------------------------------------
// MsoDrawing + Obj — Escher drawing container for pivot page field dropdown
// ---------------------------------------------------------------------------

/// Write a MsoDrawingGroup record (0x00EB) to the globals stream.
/// This initializes the Escher drawing container group for the workbook.
pub fn write_mso_drawing_group<W: Write>(writer: &mut W) -> XlsResult<()> {
    write_record_header(writer, MSO_DRAWING_GROUP_RECORD_ID, 98)?;
    write_escher_record_header(
        writer,
        EscherRecordHeader {
            options: 0x000F,
            record_id: ESCHER_DGG_CONTAINER,
            data_size: 0x5A,
        },
    )?;
    write_escher_record_header(
        writer,
        EscherRecordHeader {
            options: 0x0000,
            record_id: ESCHER_DGG,
            data_size: 0x20,
        },
    )?;
    writer.write_all(&0x0000_0802u32.to_le_bytes())?;
    writer.write_all(&3u32.to_le_bytes())?;
    writer.write_all(&3u32.to_le_bytes())?;
    writer.write_all(&2u32.to_le_bytes())?;
    for cluster in [
        DrawingGroupCluster {
            drawing_group_id: 1,
            used_shape_ids: 1,
        },
        DrawingGroupCluster {
            drawing_group_id: 2,
            used_shape_ids: 2,
        },
    ] {
        writer.write_all(&cluster.drawing_group_id.to_le_bytes())?;
        writer.write_all(&cluster.used_shape_ids.to_le_bytes())?;
    }
    write_escher_record_header(
        writer,
        EscherRecordHeader {
            options: 0x0033,
            record_id: ESCHER_OPT,
            data_size: 0x12,
        },
    )?;
    write_escher_properties(
        writer,
        &[
            EscherProperty {
                property_id: 0x00BF,
                value: 0x0008_0008,
            },
            EscherProperty {
                property_id: 0x0181,
                value: 0x0800_0041,
            },
            EscherProperty {
                property_id: 0x01C0,
                value: 0x0800_0040,
            },
        ],
    )?;
    write_escher_record_header(
        writer,
        EscherRecordHeader {
            options: 0x0040,
            record_id: ESCHER_SPLIT_MENU_COLORS,
            data_size: 0x10,
        },
    )?;
    for color in [
        0x0800_000Du32,
        0x0800_000Cu32,
        0x0800_0017u32,
        0x1000_00F7u32,
    ] {
        writer.write_all(&color.to_le_bytes())?;
    }
    Ok(())
}

/// Write an empty MsoDrawing record (0x00EC) to the data sheet.
/// Excel expects this when a drawing group is defined but the sheet has no shapes.
pub fn write_mso_drawing_sheet1<W: Write>(writer: &mut W) -> XlsResult<()> {
    write_record_header(writer, MSO_DRAWING_RECORD_ID, 80)?;
    write_escher_record_header(
        writer,
        EscherRecordHeader {
            options: 0x000F,
            record_id: ESCHER_DG_CONTAINER,
            data_size: 0x48,
        },
    )?;
    write_escher_record_header(
        writer,
        EscherRecordHeader {
            options: 0x0010,
            record_id: ESCHER_DG,
            data_size: 0x08,
        },
    )?;
    writer.write_all(&1u32.to_le_bytes())?;
    writer.write_all(&0x0000_0400u32.to_le_bytes())?;
    write_escher_record_header(
        writer,
        EscherRecordHeader {
            options: 0x000F,
            record_id: ESCHER_SPGR_CONTAINER,
            data_size: 0x30,
        },
    )?;
    write_escher_group_shape(writer, 0x0000_0400)?;
    Ok(())
}

pub fn write_pivot_page_mso_drawing<W: Write>(writer: &mut W) -> XlsResult<()> {
    write_record_header(writer, MSO_DRAWING_RECORD_ID, 170)?;
    write_escher_record_header(
        writer,
        EscherRecordHeader {
            options: 0x000F,
            record_id: ESCHER_DG_CONTAINER,
            data_size: 0xA2,
        },
    )?;
    write_escher_record_header(
        writer,
        EscherRecordHeader {
            options: 0x0020,
            record_id: ESCHER_DG,
            data_size: 0x08,
        },
    )?;
    writer.write_all(&2u32.to_le_bytes())?;
    writer.write_all(&0x0000_0801u32.to_le_bytes())?;
    write_escher_record_header(
        writer,
        EscherRecordHeader {
            options: 0x000F,
            record_id: ESCHER_SPGR_CONTAINER,
            data_size: 0x8A,
        },
    )?;
    write_escher_group_shape(writer, 0x0000_0800)?;
    write_escher_record_header(
        writer,
        EscherRecordHeader {
            options: 0x000F,
            record_id: ESCHER_SP_CONTAINER,
            data_size: 0x52,
        },
    )?;
    write_escher_record_header(
        writer,
        EscherRecordHeader {
            options: 0x0C92,
            record_id: ESCHER_SP,
            data_size: 0x08,
        },
    )?;
    writer.write_all(&0x0000_0801u32.to_le_bytes())?;
    writer.write_all(&0x0000_0A00u32.to_le_bytes())?;
    write_escher_record_header(
        writer,
        EscherRecordHeader {
            options: 0x0043,
            record_id: ESCHER_OPT,
            data_size: 0x18,
        },
    )?;
    write_escher_properties(
        writer,
        &[
            EscherProperty {
                property_id: 0x007F,
                value: 0x0104_0104,
            },
            EscherProperty {
                property_id: 0x00BF,
                value: 0x0008_0008,
            },
            EscherProperty {
                property_id: 0x01FF,
                value: 0x0008_0000,
            },
            EscherProperty {
                property_id: 0x03BF,
                value: 0x0002_0000,
            },
        ],
    )?;
    write_escher_client_anchor(writer, 0x0000, [1, 1, 0, 0, 0, 2, 0, 1, 0])?;
    write_escher_record_header(
        writer,
        EscherRecordHeader {
            options: 0x0000,
            record_id: ESCHER_CLIENT_DATA,
            data_size: 0,
        },
    )?;
    Ok(())
}

pub fn write_pivot_page_obj<W: Write>(writer: &mut W) -> XlsResult<()> {
    write_record_header(writer, OBJ_RECORD_ID, 70)?;
    write_ft_cmo_combo_box(writer)?;
    write_ft_cbls(writer)?;
    write_ft_lbs_data(writer)?;
    Ok(())
}

// DCONREF — Data Consolidation Reference (0x0051)
// ---------------------------------------------------------------------------

/// Write a DCONREF record that specifies the worksheet source range for a
/// pivot table.
///
/// Layout:
/// ```text
///  0  u16  rwFirst
///  2  u16  rwLast
///  4  u8   colFirst
///  5  u8   colLast
///  6  u16  cchFile  — length of the virtual path string
///  8  var  stFile   — XLUnicodeStringNoCch (virtual path)
/// ```
///
/// For an internal worksheet source the virtual path is encoded as
/// `\x02` + sheet name (the `\x02` prefix is the self-referential encoded
/// URL, matching Excel's behaviour per MS-XLS VirtPath).
pub fn write_dconref<W: Write>(
    writer: &mut W,
    first_row: u16,
    last_row: u16,
    first_col: u8,
    last_col: u8,
    sheet_name: &str,
) -> XlsResult<()> {
    // Virtual path = 0x02 + sheet_name (self-referential encoded URL)
    let vpath: String = format!("\x02{}", sheet_name);
    let vpath_bytes = encode_xl_string_no_cch(&vpath);
    let cch_file = vpath.chars().count() as u16;

    let data_len = 6u16 + 2 + vpath_bytes.len() as u16;

    write_record_header(writer, 0x0051, data_len)?;
    writer.write_all(&first_row.to_le_bytes())?; // 0
    writer.write_all(&last_row.to_le_bytes())?; // 2
    writer.write_all(&[first_col])?; // 4
    writer.write_all(&[last_col])?; // 5
    writer.write_all(&cch_file.to_le_bytes())?; // 6
    writer.write_all(&vpath_bytes)?;
    Ok(())
}

// ---------------------------------------------------------------------------
// SXDB — Pivot Cache Definition (0x00C6)
// ---------------------------------------------------------------------------

/// Configuration for writing an SXDB record.
pub struct SxDbConfig {
    /// Number of source data records (rows excluding header).
    pub record_count: u32,
    /// Stream ID (must match the SxStreamID for this cache).
    pub stream_id: u16,
    /// Number of standard fields (columns) in the source data.
    pub field_count: u16,
    /// SXDB flags. Use `0x0003` (fSaveData | fInvalid) when source data
    /// records (SXDBB/SXNUM) are included in the cache stream.
    pub flags: u16,
}

/// Write an SXDB record (pivot cache header).
///
/// Layout (verified against LibreOffice `xlpivot.cxx` `operator<<(XclPCInfo)`):
/// ```text
///  0  u32  mnSrcRecs     — number of source data records
///  4  u16  mnStrmId      — stream ID
///  6  u16  mnFlags       — cache flags
///  8  u16  mnBlockRecs   — 0x1FFF (max records per block)
/// 10  u16  mnStdFields   — standard (source) field count
/// 12  u16  mnTotalFields — total field count (= std for ungrouped)
/// 14  u16  crdbUsed      — records used to build cache (= mnSrcRecs)
/// 16  u16  reserved      — MUST be 0x0001
/// 18  var  userName      — ShortXLUnicodeString (empty = 3 bytes)
/// ```
pub fn write_sxdb<W: Write>(writer: &mut W, cfg: &SxDbConfig) -> XlsResult<()> {
    let user_name_bytes = encode_xl_unicode_string(""); // empty XLUnicodeString = 3 bytes
    let data_len = 18u16 + user_name_bytes.len() as u16;
    write_record_header(writer, 0x00C6, data_len)?;
    writer.write_all(&cfg.record_count.to_le_bytes())?; //  0: mnSrcRecs
    writer.write_all(&cfg.stream_id.to_le_bytes())?; //  4: mnStrmId
    writer.write_all(&cfg.flags.to_le_bytes())?; //  6: mnFlags
    writer.write_all(&0x07FFu16.to_le_bytes())?; //  8: mnBlockRecs
    writer.write_all(&cfg.field_count.to_le_bytes())?; // 10: mnStdFields
    writer.write_all(&cfg.field_count.to_le_bytes())?; // 12: mnTotalFields
    // crdbUsed: number of source records used to build the cache.
    // Per MS-XLS §2.4.269, setting this to 0 is inconsistent with having
    // cache items and may cause Excel to reject the file.
    writer.write_all(&(cfg.record_count as u16).to_le_bytes())?; // 14: crdbUsed
    writer.write_all(&0x0001u16.to_le_bytes())?; // 16: reserved (MUST be 1)
    writer.write_all(&user_name_bytes)?; // 18: userName (XLUnicodeString)
    Ok(())
}

// ---------------------------------------------------------------------------
// SXFDB — Pivot Cache Field (0x00C7)
// ---------------------------------------------------------------------------

/// Configuration for writing an SXFDB record.
pub struct SxFdbConfig<'a> {
    /// Number of unique items for this field.
    pub item_count: u16,
    /// Field name.
    pub name: &'a str,
    /// Whether this field has inline cache items following it (string field).
    pub has_items: bool,
    /// Whether this is a numeric (data-axis) field.
    ///
    /// When `true`, SXFDB flags are set to `0x0560`
    /// (fTextEtcField | fNumMinMaxValid | fNonDates | fCantGetUniqueItems)
    /// and `csxOrig` is set to 0 (no original string items).
    pub is_numeric: bool,
}

/// Encode a string as **XLUnicodeString**: `[u16 cch][u8 flags][chars…]`.
///
/// This is the standard Excel string format used by SXFDB field names
/// and SXDB userName (verified against LibreOffice `XclExpString` default
/// constructor which uses 16-bit character count).
fn encode_xl_unicode_string(s: &str) -> Vec<u8> {
    let cch = s.chars().count() as u16;
    if s.is_ascii() {
        let mut buf = Vec::with_capacity(3 + s.len());
        buf.extend_from_slice(&cch.to_le_bytes()); // 2-byte character count
        buf.push(0x00); // flags: compressed Latin-1
        buf.extend_from_slice(s.as_bytes());
        buf
    } else {
        let utf16: Vec<u16> = s.encode_utf16().collect();
        let mut buf = Vec::with_capacity(3 + utf16.len() * 2);
        buf.extend_from_slice(&cch.to_le_bytes());
        buf.push(0x01); // flags: UTF-16LE
        for ch in &utf16 {
            buf.extend_from_slice(&ch.to_le_bytes());
        }
        buf
    }
}

/// Write an SXFDB record (pivot cache field definition).
///
/// Layout (verified against LibreOffice `xepivot.cxx` `WriteSxfield`):
/// ```text
///  0  u16  flags          — fAllAtoms(bit0)=1, fOrigItems(bit9)=1
///  2  u16  ifdbParent     — group child field index (0 = base)
///  4  u16  ifdbBase       — group base field index (0 = ungrouped)
///  6  u16  citmUnq        — visible unique items
///  8  u16  csxGroup       — group items (0)
/// 10  u16  csxBase        — base items (0)
/// 12  u16  csxOrig        — original items (= citmUnq for base fields)
/// 14  var  ShortXLUnicodeString — field name
/// ```
pub fn write_sxfdb<W: Write>(writer: &mut W, cfg: &SxFdbConfig<'_>) -> XlsResult<()> {
    let name_bytes = encode_xl_unicode_string(cfg.name);
    let data_len = 14u16 + name_bytes.len() as u16;

    // Build flags:
    //   String fields with items: fAllAtoms(bit0) | DATA_STR(0x0480) = 0x0481
    //   Numeric fields: fTextEtcField | fNumMinMaxValid | fNonDates | fCantGetUniqueItems = 0x0560
    //   Other fields without items: 0x0000
    let flags: u16 = if cfg.has_items {
        0x0001 | 0x0480 // HASITEMS | DATA_STR = 0x0481
    } else if cfg.is_numeric {
        0x0560
    } else {
        0
    };

    write_record_header(writer, 0x00C7, data_len)?;
    writer.write_all(&flags.to_le_bytes())?; //  0: flags
    // 0xFFFF = no parent/base group field (ungrouped).
    // 0x0000 would mean "parent is field 0", implying grouping → "group edit mode" error.
    writer.write_all(&0xFFFFu16.to_le_bytes())?; //  2: ifdbParent
    writer.write_all(&0xFFFFu16.to_le_bytes())?; //  4: ifdbBase
    writer.write_all(&cfg.item_count.to_le_bytes())?; //  6: citmUnq (visible)
    writer.write_all(&0u16.to_le_bytes())?; //  8: csxGroup
    writer.write_all(&0u16.to_le_bytes())?; // 10: csxBase
    // csxOrig: for string fields = citmUnq, for numeric fields = 0
    let csx_orig = if cfg.is_numeric { 0u16 } else { cfg.item_count };
    writer.write_all(&csx_orig.to_le_bytes())?; // 12: csxOrig
    writer.write_all(&name_bytes)?;
    Ok(())
}

// ---------------------------------------------------------------------------
// SXSTRING — Pivot Cache String Item (0x00CD)
// ---------------------------------------------------------------------------

/// Write an SXSTRING record containing a single cache string item.
///
/// The data is an `XLUnicodeString` (u16 cch + u8 flags + chars).
pub fn write_sxstring<W: Write>(writer: &mut W, value: &str) -> XlsResult<()> {
    let cch = value.chars().count() as u16;
    if value.is_ascii() {
        let data_len = 3u16 + cch; // u16 cch + u8 flags(0) + cch bytes
        write_record_header(writer, 0x00CD, data_len)?;
        writer.write_all(&cch.to_le_bytes())?;
        writer.write_all(&[0x00])?; // flags: compressed
        writer.write_all(value.as_bytes())?;
    } else {
        let utf16: Vec<u16> = value.encode_utf16().collect();
        let data_len = 3u16 + (utf16.len() as u16) * 2;
        write_record_header(writer, 0x00CD, data_len)?;
        writer.write_all(&cch.to_le_bytes())?;
        writer.write_all(&[0x01])?; // flags: UTF-16LE
        for ch in &utf16 {
            writer.write_all(&ch.to_le_bytes())?;
        }
    }
    Ok(())
}

// ---------------------------------------------------------------------------
// SXFDBTYPE — Pivot Cache Field Data Type (0x01BB)
// ---------------------------------------------------------------------------

/// Write an SXFDBTYPE record (2 bytes).
///
/// Always written by LO after each SXFDB record in the cache stream
/// (see `XclExpPCField::Save` in `xepivot.cxx`).
/// Value is `EXC_SXFDBTYPE_DEFAULT = 0x0000`.
pub fn write_sxfdbtype<W: Write>(writer: &mut W) -> XlsResult<()> {
    write_record_header(writer, 0x01BB, 2)?;
    writer.write_all(&0u16.to_le_bytes())?;
    Ok(())
}

// ---------------------------------------------------------------------------
// SXDBEX — Extended Pivot Cache Info (0x0122)
// ---------------------------------------------------------------------------

/// Write an SXDBEX record (12 bytes).
///
/// Layout (per LibreOffice `WriteSxdbex`):
/// ```text
///  0  f64  fSxCreationDate — creation timestamp (51901.0296527…)
///  8  u32  cSxFormula      — number of SXFORMULA records (0)
/// ```
pub fn write_sxdbex<W: Write>(writer: &mut W) -> XlsResult<()> {
    write_record_header(writer, 0x0122, 12)?;
    writer.write_all(&super::sxdbex_creation_timestamp_bytes())?;
    writer.write_all(&0u32.to_le_bytes())?; // cSxFormula = 0
    Ok(())
}

// ---------------------------------------------------------------------------
// SXLI — Line Item Array (0x00B5)
// ---------------------------------------------------------------------------

/// Write an SXLI record (line item array).
///
/// For the simple non-OLAP pivots we generate, Excel expects sequential item
/// indices for each visible line and a trailing grand-total line.
///
/// Each line entry is `8 + 2 * index_count` bytes:
/// ```text
///  0  u16  cSic         — 0 (equal index count)
///  2  u16  itmType      — 0x0000 (DATA) or 0x000D (GRAND)
///  4  u16  isxviMac     — index_count
///  6  u16  cchS         — 0x0000 for data lines, 0x0A00 for grand total
///  8  [u16 × index_count]  — pivot item indices
/// ```
pub fn write_sxli<W: Write>(writer: &mut W, line_count: u16, index_count: u16) -> XlsResult<()> {
    if line_count == 0 {
        return Ok(());
    }
    let line_size = 8u32 + 2 * index_count as u32;
    let total = line_size * line_count as u32;
    write_record_header(writer, 0x00B5, total as u16)?;

    let last_line = line_count - 1;
    for line_idx in 0..line_count {
        writer.write_all(&0u16.to_le_bytes())?; // cSic
        let is_grand_total = line_idx == last_line;
        let itm_type = if is_grand_total { 0x000Du16 } else { 0x0000u16 };
        writer.write_all(&itm_type.to_le_bytes())?;
        writer.write_all(&index_count.to_le_bytes())?; // isxviMac
        let cch_s = if is_grand_total { 0x0A00u16 } else { 0u16 };
        writer.write_all(&cch_s.to_le_bytes())?;
        let item_index = if is_grand_total { 0u16 } else { line_idx };
        for _ in 0..index_count {
            writer.write_all(&item_index.to_le_bytes())?;
        }
    }
    Ok(())
}

// ---------------------------------------------------------------------------
// SXVDEx — Extended Pivot Field Properties (0x0100)
// ---------------------------------------------------------------------------

/// Default flags for SXVDEx per Excel reference file:
///
/// - fDragToRow(bit1)=1, fDragToColumn(bit2)=1, fDragToPage(bit3)=1,
///   fDragToHide(bit4)=1
/// - fAscendSort(bit10)=1, fTopAutoShow(bit12)=1
/// - fLayoutReport(bit21)=1, fLayoutTop(bit23)=1
/// - citmAutoShow(bits24-31)=10
const SXVDEX_DEFAULT_FLAGS: u32 = 0x0AA0_141E;

/// Write an SXVDEx record (extended pivot field properties).
///
/// Layout (20 bytes, verified against LibreOffice `xlpivot.cxx` `operator<<`):
/// ```text
///  0  u32  flags           — default 0x0A00141E
///  4  u16  isxdiAutoSort   — -1 = sort by own values
///  6  u16  isxdiAutoShow   — -1 = none
///  8  u16  ifmt            — 0 = no number format
/// 10  u16  cchSubName      — 0xFFFF = not present
/// 12  [8]  reserved        — zeros
/// ```
pub fn write_sxvdex<W: Write>(writer: &mut W) -> XlsResult<()> {
    write_record_header(writer, 0x0100, 20)?;
    writer.write_all(&SXVDEX_DEFAULT_FLAGS.to_le_bytes())?; //  0: flags
    writer.write_all(&0xFFFFu16.to_le_bytes())?; //  4: isxdiAutoSort (-1)
    writer.write_all(&0xFFFFu16.to_le_bytes())?; //  6: isxdiAutoShow (-1)
    writer.write_all(&0u16.to_le_bytes())?; //  8: ifmt
    writer.write_all(&0xFFFFu16.to_le_bytes())?; // 10: cchSubName (not present)
    writer.write_all(&[0u8; 8])?; // 12: reserved
    Ok(())
}

// ---------------------------------------------------------------------------
// SXIVD — Pivot Field Index List (0x00B4)
// ---------------------------------------------------------------------------

/// Write an SXIVD record containing a list of pivot field indices.
///
/// One SXIVD is written for row fields and another for column fields.
/// Each entry is a `u16` pivot field index. The special value `0xFFFE`
/// (`EXC_SXIVD_DATA`) represents the data-layout pseudo-field.
pub fn write_sxivd<W: Write>(writer: &mut W, field_indices: &[u16]) -> XlsResult<()> {
    if field_indices.is_empty() {
        return Ok(());
    }
    let data_len = (field_indices.len() * 2) as u16;
    write_record_header(writer, 0x00B4, data_len)?;
    for &idx in field_indices {
        writer.write_all(&idx.to_le_bytes())?;
    }
    Ok(())
}

// ---------------------------------------------------------------------------
// Pivot cache stream generator
// ---------------------------------------------------------------------------

/// Per-field info for the pivot cache stream.
pub struct PivotCacheFieldInfo<'a> {
    /// Field (column) name.
    pub name: &'a str,
    /// Cache item string values (unique values from source data).
    /// Empty for numeric (data-axis) fields.
    pub items: &'a [&'a str],
    /// Whether this is a numeric (data-axis) field.
    pub is_numeric: bool,
    /// Number of unique numeric values (only used when `is_numeric` is true).
    pub unique_numeric_count: u16,
}

/// A single source data row for the pivot cache.
///
/// Contains one byte per string field (the item index) and one f64 per
/// numeric field. The order matches the field order.
pub struct PivotCacheSourceRow<'a> {
    /// Packed string-field indices (one byte per string field, in order).
    pub string_indices: &'a [u8],
    /// Numeric values (one per numeric field, in order).
    pub numeric_values: &'a [f64],
}

/// Information needed to generate one pivot cache storage stream.
pub struct PivotCacheStreamInfo<'a> {
    /// Stream ID (matches the SxStreamID in globals).
    pub stream_id: u16,
    /// Number of source data records (rows excluding header).
    pub record_count: u32,
    /// Field definitions with their cache items.
    pub fields: &'a [PivotCacheFieldInfo<'a>],
    /// Source data rows. When non-empty, SXDBB + SXNUM records are written
    /// and SXDB `fSaveData` flag is set.
    pub source_rows: &'a [PivotCacheSourceRow<'a>],
}

/// Write an SXDBB record (0x00C8) — packed string-field indices for one row.
///
/// Each byte is the item index for one string field. Numeric fields are
/// excluded (they get separate SXNUM records).
fn write_sxdbb<W: Write>(writer: &mut W, indices: &[u8]) -> XlsResult<()> {
    write_record_header(writer, 0x00C8, indices.len() as u16)?;
    writer.write_all(indices)?;
    Ok(())
}

/// Write an SxNum record (0x00C9) — a single f64 numeric value for one row.
fn write_sxnum<W: Write>(writer: &mut W, value: f64) -> XlsResult<()> {
    write_record_header(writer, 0x00C9, 8)?;
    writer.write_all(&value.to_le_bytes())?;
    Ok(())
}

/// Generate a complete pivot cache storage stream (for `_SX_DB_CUR/nnnn`).
///
/// Layout:
/// ```text
/// SXDB + SXDBEX
///   + (SXFDB + SXFDBTYPE + SXSTRING×items)×n
///   + (SXDBB + SXNUM×numeric_fields)×rows   // when source_rows is non-empty
///   + EOF
/// ```
///
/// Note: LibreOffice does NOT write BOF for cache streams.
pub fn generate_pivot_cache_stream(info: &PivotCacheStreamInfo<'_>) -> XlsResult<Vec<u8>> {
    let mut buf = Vec::new();

    let has_source_data = !info.source_rows.is_empty();

    // SXDB — cache definition header
    // flags: fSaveData(bit0) | fInvalid(bit1) when source data is present,
    //        fInvalid(bit1) | fRefreshOnLoad(bit2) | fEnableRefresh(bit5) otherwise.
    let sxdb_flags = if has_source_data {
        0x0003u16
    } else {
        0x0026u16
    };
    write_sxdb(
        &mut buf,
        &SxDbConfig {
            record_count: info.record_count,
            stream_id: info.stream_id,
            field_count: info.fields.len() as u16,
            flags: sxdb_flags,
        },
    )?;

    // SXDBEX — extended cache info (creation date + formula count)
    write_sxdbex(&mut buf)?;

    // Per-field: SXFDB + SXFDBTYPE + SXSTRING items
    // Order per LO XclExpPCField::Save(): SXFIELD, SXFDBTYPE, items
    for field in info.fields {
        let has_items = !field.items.is_empty();
        let item_count = if field.is_numeric {
            field.unique_numeric_count
        } else {
            field.items.len() as u16
        };
        write_sxfdb(
            &mut buf,
            &SxFdbConfig {
                item_count,
                name: field.name,
                has_items,
                is_numeric: field.is_numeric,
            },
        )?;

        // SXFDBTYPE — always written after SXFDB per LO
        write_sxfdbtype(&mut buf)?;

        // Write cache items as SXSTRING records (string fields only)
        for &item_value in field.items {
            write_sxstring(&mut buf, item_value)?;
        }
    }

    // Source data records: SXDBB + SXNUM per row
    for row in info.source_rows {
        write_sxdbb(&mut buf, row.string_indices)?;
        for &val in row.numeric_values {
            write_sxnum(&mut buf, val)?;
        }
    }

    // EOF
    super::write_eof(&mut buf)?;

    Ok(buf)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_write_sxvs() {
        let mut buf = Vec::new();
        write_sxvs(&mut buf, 0x0001).unwrap();
        // Record type 0x00E3, length 2
        assert_eq!(&buf[0..2], &[0xE3, 0x00]);
        assert_eq!(&buf[2..4], &[0x02, 0x00]);
        assert_eq!(&buf[4..6], &[0x01, 0x00]); // Worksheet source
    }

    #[test]
    fn test_write_sxvd_no_name() {
        let mut buf = Vec::new();
        write_sxvd(
            &mut buf,
            &SxVdConfig {
                axis: 0x0001,
                subtotal_count: 0,
                subtotal_flags: 0,
                item_count: 5,
                name: None,
            },
        )
        .unwrap();
        // Record type 0x00B1
        assert_eq!(&buf[0..2], &[0xB1, 0x00]);
        // cchName = 0xFFFF at offset 4+8 = 12
        assert_eq!(&buf[12..14], &[0xFF, 0xFF]);
    }

    #[test]
    fn test_write_sxvi_data() {
        let mut buf = Vec::new();
        write_sxvi(
            &mut buf,
            &SxViConfig {
                item_type: 0x00FE,
                flags: 0,
                cache_index: 3,
                name: None,
            },
        )
        .unwrap();
        assert_eq!(&buf[0..2], &[0xB2, 0x00]); // SXVI record type
    }

    #[test]
    fn test_write_sxpi() {
        let mut buf = Vec::new();
        write_sxpi(&mut buf, &[(1, 0, 0), (2, 1, 0)]).unwrap();
        assert_eq!(&buf[0..2], &[0xB6, 0x00]); // SXPI record type
        assert_eq!(&buf[2..4], &[12, 0]); // data length = 2*6 = 12
    }

    #[test]
    fn test_write_sx_stream_id() {
        let mut buf = Vec::new();
        write_sx_stream_id(&mut buf, 0).unwrap();
        assert_eq!(&buf[0..2], &[0xD5, 0x00]); // record type 0x00D5
        assert_eq!(&buf[2..4], &[0x02, 0x00]); // length 2
        assert_eq!(&buf[4..6], &[0x00, 0x00]); // id_stm = 0
    }

    #[test]
    fn test_write_sxex_default() {
        let mut buf = Vec::new();
        write_sxex(&mut buf, &SxExConfig::default()).unwrap();
        assert_eq!(&buf[0..2], &[0xF1, 0x00]); // record type 0x00F1
        assert_eq!(&buf[2..4], &[24, 0]); // length 24
        // cchErrorString at offset 6 (after csxformat at 4) = 0xFFFF
        assert_eq!(&buf[6..8], &[0xFF, 0xFF]);
    }

    #[test]
    fn test_encode_xl_string_no_cch_empty() {
        // Empty string produces no bytes (no flags byte).
        let bytes = encode_xl_string_no_cch("");
        assert!(bytes.is_empty());
    }

    #[test]
    fn test_write_sxdi_empty_name() {
        let mut buf = Vec::new();
        write_sxdi(
            &mut buf,
            &SxDiConfig {
                source_field_index: 0,
                function: 0,
                display_format: 0,
                base_field_index: 0,
                base_item_index: 0,
                num_format_index: 0,
                name: "",
            },
        )
        .unwrap();
        // cchName at offset 4+12 = 16 should be 0xFFFF
        assert_eq!(&buf[16..18], &[0xFF, 0xFF]);
        // Total length should be just the 14-byte fixed header
        assert_eq!(&buf[2..4], &[14, 0]);
    }
}
