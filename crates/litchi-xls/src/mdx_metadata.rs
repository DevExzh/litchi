//! BIFF8 MDX (OLAP cube) metadata records: the `METADATA` production of the
//! workbook globals substream (MS-XLS 2.1).
//!
//! The grammar is:
//!
//! ```text
//! METADATA     = *MDTINFO *MDXSTR *(MDXTUPLESET / MDXProp / MDXKPI) *MDBLOCK
//! MDTINFO      = MDTInfo *ContinueFrt12
//! MDXSTR       = MDXStr *ContinueFrt12
//! MDXTUPLESET  = (MDXTuple / MDXSet) *ContinueFrt12
//! MDBLOCK      = MDB *ContinueFrt12
//! ```
//!
//! This module implements typed readers and writers for all seven records:
//!
//! - **MDTInfo** (0x0884): behavior flags and name of one metadata type.
//! - **MDXStr** (0x0885): a shared text string referenced by index.
//! - **MDXTuple** (0x0886): tuple metadata produced by a cube function.
//! - **MDXSet** (0x0887): set metadata with a sort order.
//! - **MDXProp** (0x0888): member property metadata.
//! - **MDXKPI** (0x0889): key performance indicator metadata.
//! - **MDB** (0x088A): a block of metadata type/value pairs referenced by
//!   cells carrying value metadata.
//!
//! Everything in this module is INERT: connection names, MDX unique names,
//! and set definitions are stored verbatim and no OLAP server is ever
//! contacted.
//!
//! # References
//!
//! - MS-XLS sections 2.4.161–2.4.167, 2.5.135 (FrtHeader), 2.5.179
//!   (LPWideString), 2.5.180–2.5.182 (MDir, MDTInfoIndex, MDXStrIndex),
//!   2.5.169 (KPIProp), 2.5.233 (SD_SetSortOrder), 2.5.267 (Tag_Fn_MDX)

use super::{Error, Result};

/// Record type of the `MDTInfo` record (MS-XLS 2.4.162).
pub(crate) const MDT_INFO_RECORD_TYPE: u16 = 0x0884;
/// Record type of the `MDXStr` record (MS-XLS 2.4.166).
pub(crate) const MDX_STR_RECORD_TYPE: u16 = 0x0885;
/// Record type of the `MDXTuple` record (MS-XLS 2.4.167).
pub(crate) const MDX_TUPLE_RECORD_TYPE: u16 = 0x0886;
/// Record type of the `MDXSet` record (MS-XLS 2.4.165).
pub(crate) const MDX_SET_RECORD_TYPE: u16 = 0x0887;
/// Record type of the `MDXProp` record (MS-XLS 2.4.164).
pub(crate) const MDX_PROP_RECORD_TYPE: u16 = 0x0888;
/// Record type of the `MDXKPI` record (MS-XLS 2.4.163).
pub(crate) const MDX_KPI_RECORD_TYPE: u16 = 0x0889;
/// Record type of the `MDB` record (MS-XLS 2.4.161).
pub(crate) const MDB_RECORD_TYPE: u16 = 0x088A;
/// Record type of the `ContinueFrt12` record (MS-XLS 2.4.61) that continues
/// a `METADATA` payload.
pub(crate) const CONTINUE_FRT12_RECORD_TYPE: u16 = 0x087F;

/// Size in bytes of an `FrtHeader` (MS-XLS 2.5.135).
const FRT_HEADER_LEN: usize = 12;
/// Size in bytes of an `MDir` structure (MS-XLS 2.5.180).
const MDIR_LEN: usize = 8;
/// Maximum character count of an `LPWideString` (MS-XLS 2.5.179): the
/// length prefix is an unsigned 16-bit character count.
const MAX_LP_WIDE_STRING_CHARS: usize = u16::MAX as usize;

fn invalid(record_type: u16, message: impl Into<String>) -> Error {
    Error::InvalidRecord {
        record_type,
        message: message.into(),
    }
}

fn read_u32(data: &[u8], offset: usize) -> u32 {
    u32::from_le_bytes([
        data[offset],
        data[offset + 1],
        data[offset + 2],
        data[offset + 3],
    ])
}

fn read_i32(data: &[u8], offset: usize) -> i32 {
    i32::from_le_bytes([
        data[offset],
        data[offset + 1],
        data[offset + 2],
        data[offset + 3],
    ])
}

/// Validate the `FrtHeader` and return the record body that follows it.
fn record_body(data: &[u8], record_type: u16) -> Result<&[u8]> {
    super::differential_format::validate_frt_header(data, record_type)?;
    Ok(&data[FRT_HEADER_LEN..])
}

/// Parse an `LPWideString` (MS-XLS 2.5.179) that must span `data` exactly:
/// a 16-bit character count followed by that many UTF-16LE code units.
fn parse_lp_wide_string(data: &[u8], record_type: u16) -> Result<String> {
    if data.len() < 2 {
        return Err(invalid(record_type, "truncated LPWideString"));
    }
    let count = usize::from(u16::from_le_bytes([data[0], data[1]]));
    if data.len() != 2 + count * 2 {
        return Err(invalid(record_type, "LPWideString length mismatch"));
    }
    let units: Vec<u16> = data[2..]
        .chunks_exact(2)
        .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
        .collect();
    String::from_utf16(&units)
        .map_err(|_error| invalid(record_type, "LPWideString is not valid UTF-16LE"))
}

/// Serialize an `LPWideString` (MS-XLS 2.5.179) into `output`.
fn append_lp_wide_string(record_type: u16, value: &str, output: &mut Vec<u8>) -> Result<()> {
    let units: Vec<u16> = value.encode_utf16().collect();
    if units.len() > MAX_LP_WIDE_STRING_CHARS {
        return Err(Error::InvalidData(format!(
            "record 0x{record_type:04X} LPWideString exceeds {MAX_LP_WIDE_STRING_CHARS} UTF-16 code units"
        )));
    }
    output.extend_from_slice(&crate::utils::truncate_usize_to_u16(units.len()).to_le_bytes());
    for unit in units {
        output.extend_from_slice(&unit.to_le_bytes());
    }
    Ok(())
}

/// Begin a record payload with the 12-byte `FrtHeader` (MS-XLS 2.5.135):
/// the record type echo followed by zero flags and reserved bytes.
fn frt_header_payload(record_type: u16) -> Vec<u8> {
    let mut payload = Vec::with_capacity(FRT_HEADER_LEN);
    payload.extend_from_slice(&record_type.to_le_bytes());
    payload.extend_from_slice(&[0; FRT_HEADER_LEN - 2]);
    payload
}

/// The cube function that generated an MDX metadata record (`Tag_Fn_MDX`,
/// MS-XLS 2.5.267).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum CubeFunction {
    /// `CUBEMEMBER` (TFNCUBEMEMBER).
    CubeMember = 0x01,
    /// `CUBEVALUE` (TFNCUBEVALUE).
    CubeValue = 0x02,
    /// `CUBESET` (TFNCUBESET).
    CubeSet = 0x03,
    /// `CUBESETCOUNT` (TFNCUBESETCOUNT).
    CubeSetCount = 0x04,
    /// `CUBERANKEDMEMBER` (TFNCUBERANKEDMEMBER).
    CubeRankedMember = 0x05,
    /// `CUBEMEMBERPROPERTY` (TFNCUBEMEMBERPROPERTY).
    CubeMemberProperty = 0x06,
    /// `CUBEKPIPROPERTY` (TFNCUBEKPIPROPERTY).
    CubeKpiProperty = 0x07,
}

impl CubeFunction {
    fn from_code(record_type: u16, code: u8) -> Result<Self> {
        Ok(match code {
            0x01 => Self::CubeMember,
            0x02 => Self::CubeValue,
            0x03 => Self::CubeSet,
            0x04 => Self::CubeSetCount,
            0x05 => Self::CubeRankedMember,
            0x06 => Self::CubeMemberProperty,
            0x07 => Self::CubeKpiProperty,
            other => {
                return Err(invalid(
                    record_type,
                    format!("unknown Tag_Fn_MDX value 0x{other:02X}"),
                ));
            },
        })
    }

    /// Raw `Tag_Fn_MDX` code.
    #[must_use]
    pub const fn code(self) -> u8 {
        self as u8
    }
}

/// The kind of KPI property an `MDXKPI` record carries (`KPIProp`,
/// MS-XLS 2.5.169).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum KpiProperty {
    /// Value (KPIPROPVALUE).
    Value = 0x01,
    /// Goal (KPIPROPGOAL).
    Goal = 0x02,
    /// Status (KPIPROPSTATUS).
    Status = 0x03,
    /// Trend (KPIPROPTREND).
    Trend = 0x04,
    /// Weight (KPIPROPWEIGHT).
    Weight = 0x05,
    /// Current time member (KPIPROPCURRENTTIMEMEMBER).
    CurrentTimeMember = 0x06,
}

impl KpiProperty {
    fn from_code(code: u8) -> Result<Self> {
        Ok(match code {
            0x01 => Self::Value,
            0x02 => Self::Goal,
            0x03 => Self::Status,
            0x04 => Self::Trend,
            0x05 => Self::Weight,
            0x06 => Self::CurrentTimeMember,
            other => {
                return Err(invalid(
                    MDX_KPI_RECORD_TYPE,
                    format!("unknown KPIProp value 0x{other:02X}"),
                ));
            },
        })
    }

    /// Raw `KPIProp` code.
    #[must_use]
    pub const fn code(self) -> u8 {
        self as u8
    }
}

/// The sort order of an MDX set (`SD_SetSortOrder`, MS-XLS 2.5.233).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum MdxSetSortOrder {
    /// No sorting order (SSONONE).
    None = 0x00,
    /// Ascending order (SSOASC).
    Ascending = 0x01,
    /// Descending order (SSODESC).
    Descending = 0x02,
    /// Ascending order by caption (SSOALPHAASC).
    AlphaAscending = 0x03,
    /// Descending order by caption (SSOALPHADESC).
    AlphaDescending = 0x04,
    /// Ascending order by the natural order of the data (SSONATURALASC).
    NaturalAscending = 0x05,
    /// Descending order by the natural order of the data (SSONATURALDESC).
    NaturalDescending = 0x06,
}

impl MdxSetSortOrder {
    fn from_code(code: u8) -> Result<Self> {
        Ok(match code {
            0x00 => Self::None,
            0x01 => Self::Ascending,
            0x02 => Self::Descending,
            0x03 => Self::AlphaAscending,
            0x04 => Self::AlphaDescending,
            0x05 => Self::NaturalAscending,
            0x06 => Self::NaturalDescending,
            other => {
                return Err(invalid(
                    MDX_SET_RECORD_TYPE,
                    format!("unknown SD_SetSortOrder value 0x{other:02X}"),
                ));
            },
        })
    }

    /// Raw `SD_SetSortOrder` code.
    #[must_use]
    pub const fn code(self) -> u8 {
        self as u8
    }
}

// `MDTInfo` flag bits (MS-XLS 2.4.162).
const F_GHOST_ROW: u32 = 0x0000_0001;
const F_GHOST_COL: u32 = 0x0000_0002;
const F_EDIT: u32 = 0x0000_0004;
const F_DELETE: u32 = 0x0000_0008;
const F_COPY: u32 = 0x0000_0010;
const F_PASTE_ALL: u32 = 0x0000_0020;
const F_PASTE_FORMULAS: u32 = 0x0000_0040;
const F_PASTE_VALUES: u32 = 0x0000_0080;
const F_PASTE_FORMATS: u32 = 0x0000_0100;
const F_PASTE_COMMENTS: u32 = 0x0000_0200;
const F_PASTE_DATA_VALIDATION: u32 = 0x0000_0400;
const F_PASTE_BORDERS: u32 = 0x0000_0800;
const F_PASTE_COL_WIDTHS: u32 = 0x0000_1000;
const F_PASTE_NUMBER_FORMATS: u32 = 0x0000_2000;
const F_MERGE: u32 = 0x0000_4000;
const F_SPLIT_FIRST: u32 = 0x0000_8000;
const F_SPLIT_ALL: u32 = 0x0001_0000;
const F_ROW_COL_SHIFT: u32 = 0x0002_0000;
const F_CLEAR_ALL: u32 = 0x0004_0000;
const F_CLEAR_FORMATS: u32 = 0x0008_0000;
const F_CLEAR_CONTENTS: u32 = 0x0010_0000;
const F_CLEAR_COMMENTS: u32 = 0x0020_0000;
const F_ASSIGN: u32 = 0x0040_0000;
const F_COERCE: u32 = 0x1000_0000;
const F_ADJUST: u32 = 0x2000_0000;
const F_CELL_META: u32 = 0x4000_0000;

/// Behavior flags of one metadata type (`MDTInfo.grbit`, MS-XLS 2.4.162).
///
/// Each flag specifies whether the metadata is preserved, copied, or
/// applied when the cell carrying it is edited, deleted, copied, pasted,
/// merged, split, shifted, cleared, or coerced.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub struct MdtInfoFlags(u32);

impl MdtInfoFlags {
    /// Create a flag set from raw bits; undefined bits are preserved so the
    /// value round-trips losslessly.
    #[must_use]
    pub const fn from_bits(bits: u32) -> Self {
        Self(bits)
    }

    /// Raw 32-bit flag value.
    #[must_use]
    pub const fn bits(self) -> u32 {
        self.0
    }

    /// `fGhostRow`: applied to all cells in newly inserted rows.
    #[must_use]
    pub const fn ghost_row(self) -> bool {
        self.0 & F_GHOST_ROW != 0
    }
    /// `fGhostCol`: applied to all cells in newly inserted columns.
    #[must_use]
    pub const fn ghost_column(self) -> bool {
        self.0 & F_GHOST_COL != 0
    }
    /// `fEdit`: preserved when the cell is edited.
    #[must_use]
    pub const fn preserved_on_edit(self) -> bool {
        self.0 & F_EDIT != 0
    }
    /// `fDelete`: preserved when the cell's value is deleted.
    #[must_use]
    pub const fn preserved_on_delete(self) -> bool {
        self.0 & F_DELETE != 0
    }
    /// `fCopy`: copied when the cell is copied.
    #[must_use]
    pub const fn copied_with_cell(self) -> bool {
        self.0 & F_COPY != 0
    }
    /// `fPasteAll`: pasted when everything is pasted from the copied cell.
    #[must_use]
    pub const fn paste_all(self) -> bool {
        self.0 & F_PASTE_ALL != 0
    }
    /// `fPasteFormulas`: pasted when only formulas are pasted.
    #[must_use]
    pub const fn paste_formulas(self) -> bool {
        self.0 & F_PASTE_FORMULAS != 0
    }
    /// `fPasteValues`: pasted when only values are pasted.
    #[must_use]
    pub const fn paste_values(self) -> bool {
        self.0 & F_PASTE_VALUES != 0
    }
    /// `fPasteFormats`: pasted when only formatting is pasted.
    #[must_use]
    pub const fn paste_formats(self) -> bool {
        self.0 & F_PASTE_FORMATS != 0
    }
    /// `fPasteComments`: pasted when only comments are pasted.
    #[must_use]
    pub const fn paste_comments(self) -> bool {
        self.0 & F_PASTE_COMMENTS != 0
    }
    /// `fPasteDataValidation`: pasted when only data validation rules are
    /// pasted.
    #[must_use]
    pub const fn paste_data_validation(self) -> bool {
        self.0 & F_PASTE_DATA_VALIDATION != 0
    }
    /// `fPasteBorders`: pasted when only borders are pasted.
    #[must_use]
    pub const fn paste_borders(self) -> bool {
        self.0 & F_PASTE_BORDERS != 0
    }
    /// `fPasteColWidths`: pasted when only column widths are pasted.
    #[must_use]
    pub const fn paste_col_widths(self) -> bool {
        self.0 & F_PASTE_COL_WIDTHS != 0
    }
    /// `fPasteNumberFormats`: pasted when only number formatting is pasted.
    #[must_use]
    pub const fn paste_number_formats(self) -> bool {
        self.0 & F_PASTE_NUMBER_FORMATS != 0
    }
    /// `fMerge`: preserved after cells are merged.
    #[must_use]
    pub const fn preserved_on_merge(self) -> bool {
        self.0 & F_MERGE != 0
    }
    /// `fSplitFirst`: copied to the top-left resulting cell when split.
    #[must_use]
    pub const fn split_first(self) -> bool {
        self.0 & F_SPLIT_FIRST != 0
    }
    /// `fSplitAll`: copied to all resulting cells when split.
    #[must_use]
    pub const fn split_all(self) -> bool {
        self.0 & F_SPLIT_ALL != 0
    }
    /// `fRowColShift`: preserved when the cell is shifted by row or column
    /// insertion or deletion.
    #[must_use]
    pub const fn preserved_on_row_col_shift(self) -> bool {
        self.0 & F_ROW_COL_SHIFT != 0
    }
    /// `fClearAll`: preserved when contents, formatting, and comments are
    /// cleared.
    #[must_use]
    pub const fn preserved_on_clear_all(self) -> bool {
        self.0 & F_CLEAR_ALL != 0
    }
    /// `fClearFormats`: preserved when the formatting is cleared.
    #[must_use]
    pub const fn preserved_on_clear_formats(self) -> bool {
        self.0 & F_CLEAR_FORMATS != 0
    }
    /// `fClearContents`: preserved when the contents are cleared.
    #[must_use]
    pub const fn preserved_on_clear_contents(self) -> bool {
        self.0 & F_CLEAR_CONTENTS != 0
    }
    /// `fClearComments`: preserved when the comments are cleared.
    #[must_use]
    pub const fn preserved_on_clear_comments(self) -> bool {
        self.0 & F_CLEAR_COMMENTS != 0
    }
    /// `fAssign`: preserved when the cell's value is changed by formula
    /// assignment.
    #[must_use]
    pub const fn preserved_on_assign(self) -> bool {
        self.0 & F_ASSIGN != 0
    }
    /// `fCoerce`: preserved when the cell's value is coerced to a different
    /// type.
    #[must_use]
    pub const fn preserved_on_coerce(self) -> bool {
        self.0 & F_COERCE != 0
    }
    /// `fAdjust`: updated when the cell's location changes.
    #[must_use]
    pub const fn adjusted_on_move(self) -> bool {
        self.0 & F_ADJUST != 0
    }
    /// `fCellMeta`: this is cell metadata rather than value metadata.
    #[must_use]
    pub const fn is_cell_metadata(self) -> bool {
        self.0 & F_CELL_META != 0
    }
}

/// An `MDTInfo` record (MS-XLS 2.4.162): behavior flags and the name of a
/// single metadata type.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MdtInfo {
    /// Behavior flags (`grbit`).
    pub flags: MdtInfoFlags,
    /// Name of the metadata type (`stName`).
    pub name: String,
}

impl MdtInfo {
    /// Parse an `MDTInfo` record payload (`FrtHeader` included).
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn parse(data: &[u8]) -> Result<Self> {
        let body = record_body(data, MDT_INFO_RECORD_TYPE)?;
        if body.len() < 4 {
            return Err(invalid(MDT_INFO_RECORD_TYPE, "truncated MDTInfo flags"));
        }
        let flags = MdtInfoFlags::from_bits(read_u32(body, 0));
        let name = parse_lp_wide_string(&body[4..], MDT_INFO_RECORD_TYPE)?;
        Ok(Self { flags, name })
    }

    /// Serialize the record payload (`FrtHeader` included).
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn to_payload(&self) -> Result<Vec<u8>> {
        let mut payload = frt_header_payload(MDT_INFO_RECORD_TYPE);
        payload.extend_from_slice(&self.flags.bits().to_le_bytes());
        append_lp_wide_string(MDT_INFO_RECORD_TYPE, &self.name, &mut payload)?;
        Ok(payload)
    }
}

/// An `MDXTuple` record (MS-XLS 2.4.167): tuple metadata generated by a
/// `CUBEMEMBER`, `CUBEVALUE`, or `CUBERANKEDMEMBER` cube function.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MdxTuple {
    /// Index of the connection name in the shared MDX string table
    /// (`istrConnName`).
    pub connection_name_index: i32,
    /// The cube function that generated the metadata (`tfnSrc`).
    pub function: CubeFunction,
    /// Indexes of the MDX unique name strings in the shared MDX string
    /// table (`rgistr`).
    pub string_indexes: Vec<i32>,
}

impl MdxTuple {
    /// Parse an `MDXTuple` record payload (`FrtHeader` included).
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn parse(data: &[u8]) -> Result<Self> {
        let body = record_body(data, MDX_TUPLE_RECORD_TYPE)?;
        if body.len() < 9 {
            return Err(invalid(MDX_TUPLE_RECORD_TYPE, "truncated MDXTuple"));
        }
        let connection_name_index = read_i32(body, 0);
        let function = CubeFunction::from_code(MDX_TUPLE_RECORD_TYPE, body[4])?;
        if !matches!(
            function,
            CubeFunction::CubeMember | CubeFunction::CubeValue | CubeFunction::CubeRankedMember
        ) {
            return Err(invalid(
                MDX_TUPLE_RECORD_TYPE,
                "MDXTuple tfnSrc must be CUBEMEMBER, CUBEVALUE, or CUBERANKEDMEMBER",
            ));
        }
        let string_indexes = parse_index_array(body, 5, MDX_TUPLE_RECORD_TYPE)?;
        Ok(Self {
            connection_name_index,
            function,
            string_indexes,
        })
    }

    /// Serialize the record payload (`FrtHeader` included).
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn to_payload(&self) -> Result<Vec<u8>> {
        let mut payload = frt_header_payload(MDX_TUPLE_RECORD_TYPE);
        payload.extend_from_slice(&self.connection_name_index.to_le_bytes());
        payload.push(self.function.code());
        append_index_array(&self.string_indexes, &mut payload)?;
        Ok(payload)
    }
}

/// An `MDXSet` record (MS-XLS 2.4.165): set metadata generated by a
/// `CUBESET` or `CUBESETCOUNT` cube function.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MdxSet {
    /// Index of the connection name in the shared MDX string table
    /// (`istrConnName`).
    pub connection_name_index: i32,
    /// The cube function that generated the metadata (`tfnSrc`).
    pub function: CubeFunction,
    /// The set sort order (`sso`).
    pub sort_order: MdxSetSortOrder,
    /// Index of the set definition string in the shared MDX string table
    /// (`istrSetDef`).
    pub set_definition_index: i32,
    /// Indexes of the MDX unique name strings in the shared MDX string
    /// table (`rgistr`).
    pub string_indexes: Vec<i32>,
}

impl MdxSet {
    /// Parse an `MDXSet` record payload (`FrtHeader` included).
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn parse(data: &[u8]) -> Result<Self> {
        let body = record_body(data, MDX_SET_RECORD_TYPE)?;
        if body.len() < 14 {
            return Err(invalid(MDX_SET_RECORD_TYPE, "truncated MDXSet"));
        }
        let connection_name_index = read_i32(body, 0);
        let function = CubeFunction::from_code(MDX_SET_RECORD_TYPE, body[4])?;
        if !matches!(function, CubeFunction::CubeSet | CubeFunction::CubeSetCount) {
            return Err(invalid(
                MDX_SET_RECORD_TYPE,
                "MDXSet tfnSrc must be CUBESET or CUBESETCOUNT",
            ));
        }
        let sort_order = MdxSetSortOrder::from_code(body[5])?;
        let set_definition_index = read_i32(body, 6);
        let string_indexes = parse_index_array(body, 10, MDX_SET_RECORD_TYPE)?;
        Ok(Self {
            connection_name_index,
            function,
            sort_order,
            set_definition_index,
            string_indexes,
        })
    }

    /// Serialize the record payload (`FrtHeader` included).
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn to_payload(&self) -> Result<Vec<u8>> {
        let mut payload = frt_header_payload(MDX_SET_RECORD_TYPE);
        payload.extend_from_slice(&self.connection_name_index.to_le_bytes());
        payload.push(self.function.code());
        payload.push(self.sort_order.code());
        payload.extend_from_slice(&self.set_definition_index.to_le_bytes());
        append_index_array(&self.string_indexes, &mut payload)?;
        Ok(payload)
    }
}

/// An `MDXProp` record (MS-XLS 2.4.164): member property metadata generated
/// by a `CUBEMEMBERPROPERTY` cube function.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MdxProp {
    /// Index of the connection name in the shared MDX string table
    /// (`istrConnName`).
    pub connection_name_index: i32,
    /// The cube function that generated the metadata (`tfnSrc`); always
    /// [`CubeFunction::CubeMemberProperty`].
    pub function: CubeFunction,
    /// Index of the MDX unique name string (`istrMbr`).
    pub member_index: i32,
    /// Index of the property name string (`istrProp`).
    pub property_index: i32,
}

impl MdxProp {
    /// Fixed body size: `istrConnName` + `tfnSrc` + `istrMbr` + `istrProp`.
    const BODY_LEN: usize = 4 + 1 + 4 + 4;

    /// Parse an `MDXProp` record payload (`FrtHeader` included).
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn parse(data: &[u8]) -> Result<Self> {
        let body = record_body(data, MDX_PROP_RECORD_TYPE)?;
        if body.len() != Self::BODY_LEN {
            return Err(invalid(MDX_PROP_RECORD_TYPE, "malformed MDXProp body"));
        }
        let function = CubeFunction::from_code(MDX_PROP_RECORD_TYPE, body[4])?;
        if function != CubeFunction::CubeMemberProperty {
            return Err(invalid(
                MDX_PROP_RECORD_TYPE,
                "MDXProp tfnSrc must be CUBEMEMBERPROPERTY",
            ));
        }
        Ok(Self {
            connection_name_index: read_i32(body, 0),
            function,
            member_index: read_i32(body, 5),
            property_index: read_i32(body, 9),
        })
    }

    /// Serialize the record payload (`FrtHeader` included).
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn to_payload(&self) -> Result<Vec<u8>> {
        let mut payload = frt_header_payload(MDX_PROP_RECORD_TYPE);
        payload.extend_from_slice(&self.connection_name_index.to_le_bytes());
        payload.push(self.function.code());
        payload.extend_from_slice(&self.member_index.to_le_bytes());
        payload.extend_from_slice(&self.property_index.to_le_bytes());
        Ok(payload)
    }
}

/// An `MDXKPI` record (MS-XLS 2.4.163): key performance indicator metadata
/// generated by a `CUBEKPIPROPERTY` cube function.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MdxKpi {
    /// Index of the connection name in the shared MDX string table
    /// (`istrConnName`).
    pub connection_name_index: i32,
    /// The cube function that generated the metadata (`tfnSrc`); always
    /// [`CubeFunction::CubeKpiProperty`].
    pub function: CubeFunction,
    /// The KPI property kind (`kpiprop`).
    pub kpi_property: KpiProperty,
    /// Index of the MDX unique name string (`istrKPIName`).
    pub kpi_name_index: i32,
    /// Index of the key performance indicator name string (`istrMbrKPI`).
    pub member_kpi_index: i32,
}

impl MdxKpi {
    /// Fixed body size: `istrConnName` + `tfnSrc` + `kpiprop` + `istrKPIName`
    /// + `istrMbrKPI`.
    const BODY_LEN: usize = 4 + 1 + 1 + 4 + 4;

    /// Parse an `MDXKPI` record payload (`FrtHeader` included).
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn parse(data: &[u8]) -> Result<Self> {
        let body = record_body(data, MDX_KPI_RECORD_TYPE)?;
        if body.len() != Self::BODY_LEN {
            return Err(invalid(MDX_KPI_RECORD_TYPE, "malformed MDXKPI body"));
        }
        let function = CubeFunction::from_code(MDX_KPI_RECORD_TYPE, body[4])?;
        if function != CubeFunction::CubeKpiProperty {
            return Err(invalid(
                MDX_KPI_RECORD_TYPE,
                "MDXKPI tfnSrc must be CUBEKPIPROPERTY",
            ));
        }
        Ok(Self {
            connection_name_index: read_i32(body, 0),
            function,
            kpi_property: KpiProperty::from_code(body[5])?,
            kpi_name_index: read_i32(body, 6),
            member_kpi_index: read_i32(body, 10),
        })
    }

    /// Serialize the record payload (`FrtHeader` included).
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn to_payload(&self) -> Result<Vec<u8>> {
        let mut payload = frt_header_payload(MDX_KPI_RECORD_TYPE);
        payload.extend_from_slice(&self.connection_name_index.to_le_bytes());
        payload.push(self.function.code());
        payload.push(self.kpi_property.code());
        payload.extend_from_slice(&self.kpi_name_index.to_le_bytes());
        payload.extend_from_slice(&self.member_kpi_index.to_le_bytes());
        Ok(payload)
    }
}

/// Parse `cistr`/`rgistr`: a signed count at `offset` followed by that many
/// signed `MDXStrIndex` entries (MS-XLS 2.5.182), spanning the body exactly.
fn parse_index_array(body: &[u8], offset: usize, record_type: u16) -> Result<Vec<i32>> {
    let count = read_i32(body, offset);
    if count < 0 {
        return Err(invalid(record_type, "negative MDX string index count"));
    }
    let count = crate::utils::sign_extend_i32_to_usize(count);
    if body.len() != offset + 4 + count * 4 {
        return Err(invalid(
            record_type,
            "MDX string index array length mismatch",
        ));
    }
    Ok((0..count)
        .map(|i| read_i32(body, offset + 4 + i * 4))
        .collect())
}

/// Serialize `cistr`/`rgistr`.
fn append_index_array(indexes: &[i32], output: &mut Vec<u8>) -> Result<()> {
    let count = i32::try_from(indexes.len())
        .map_err(|_error| Error::InvalidData("too many MDX string indexes".to_string()))?;
    output.extend_from_slice(&count.to_le_bytes());
    for index in indexes {
        output.extend_from_slice(&index.to_le_bytes());
    }
    Ok(())
}

/// One metadata type/value pair of an `MDB` block (`MDir`, MS-XLS 2.5.180).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct MdxMetadataDir {
    /// One-based index of the `MDTInfo` record identifying the metadata
    /// type (`imdt`).
    pub info_index: i32,
    /// Zero-based index of the MDX metadata record (`MDXTuple`, `MDXSet`,
    /// `MDXProp`, or `MDXKPI`) in the workbook globals (`mdd`).
    pub metadata_index: u32,
}

/// An `MDB` record (MS-XLS 2.4.161): a unique set of metadata type/value
/// pairs shared by all cells that reference MDX value metadata.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Mdb {
    /// The metadata type/value pairs of this block (`rgmdir`).
    pub entries: Vec<MdxMetadataDir>,
}

impl Mdb {
    /// Parse an `MDB` record payload (`FrtHeader` included).
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn parse(data: &[u8]) -> Result<Self> {
        let body = record_body(data, MDB_RECORD_TYPE)?;
        if body.len() % MDIR_LEN != 0 {
            return Err(invalid(
                MDB_RECORD_TYPE,
                "MDB body is not a whole MDir array",
            ));
        }
        let entries = body
            .chunks_exact(MDIR_LEN)
            .map(|chunk| MdxMetadataDir {
                info_index: read_i32(chunk, 0),
                metadata_index: read_u32(chunk, 4),
            })
            .collect();
        Ok(Self { entries })
    }

    /// Serialize the record payload (`FrtHeader` included).
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn to_payload(&self) -> Result<Vec<u8>> {
        let mut payload = frt_header_payload(MDB_RECORD_TYPE);
        for entry in &self.entries {
            payload.extend_from_slice(&entry.info_index.to_le_bytes());
            payload.extend_from_slice(&entry.metadata_index.to_le_bytes());
        }
        Ok(payload)
    }
}

/// One MDX metadata record of the `*(MDXTUPLESET / MDXProp / MDXKPI)` run
/// (MS-XLS 2.1), in workbook record order.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum MdxMetadataRecord {
    /// An `MDXTuple` record.
    Tuple(MdxTuple),
    /// An `MDXSet` record.
    Set(MdxSet),
    /// An `MDXProp` record.
    Prop(MdxProp),
    /// An `MDXKPI` record.
    Kpi(MdxKpi),
}

impl MdxMetadataRecord {
    /// Every `MDXStrIndex` this record references.
    fn referenced_string_indexes(&self) -> Vec<i32> {
        match self {
            Self::Tuple(tuple) => {
                let mut indexes = vec![tuple.connection_name_index];
                indexes.extend_from_slice(&tuple.string_indexes);
                indexes
            },
            Self::Set(set) => {
                let mut indexes = vec![set.connection_name_index, set.set_definition_index];
                indexes.extend_from_slice(&set.string_indexes);
                indexes
            },
            Self::Prop(prop) => {
                vec![
                    prop.connection_name_index,
                    prop.member_index,
                    prop.property_index,
                ]
            },
            Self::Kpi(kpi) => {
                vec![
                    kpi.connection_name_index,
                    kpi.kpi_name_index,
                    kpi.member_kpi_index,
                ]
            },
        }
    }

    /// Serialize the record payload (`FrtHeader` included).
    fn to_payload(&self) -> Result<Vec<u8>> {
        match self {
            Self::Tuple(tuple) => tuple.to_payload(),
            Self::Set(set) => set.to_payload(),
            Self::Prop(prop) => prop.to_payload(),
            Self::Kpi(kpi) => kpi.to_payload(),
        }
    }
}

/// The `METADATA` production of the workbook globals substream (MS-XLS 2.1):
/// all MDX metadata of the workbook, in record order.
///
/// Records are collected in ABNF order — metadata types (`MDTInfo`), the
/// shared string table (`MDXStr`), the metadata records (`MDXTuple`,
/// `MDXSet`, `MDXProp`, `MDXKPI`), and the metadata blocks (`MDB`) — and
/// every cross-record index is validated against the entries collected so
/// far, as MS-XLS requires.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct MdxMetadata {
    infos: Vec<MdtInfo>,
    strings: Vec<String>,
    records: Vec<MdxMetadataRecord>,
    blocks: Vec<Mdb>,
}

impl MdxMetadata {
    /// Metadata types (`MDTInfo` records), in record order. `MDTInfoIndex`
    /// values are one-based indexes into this collection.
    #[must_use]
    pub fn infos(&self) -> &[MdtInfo] {
        &self.infos
    }

    /// The shared MDX string table (`MDXStr` records), in record order.
    /// `MDXStrIndex` values are zero-based indexes into this collection.
    #[must_use]
    pub fn strings(&self) -> &[String] {
        &self.strings
    }

    /// The MDX metadata records (`MDXTuple`, `MDXSet`, `MDXProp`, `MDXKPI`),
    /// in record order. `MDir.mdd` values are zero-based indexes into this
    /// collection.
    #[must_use]
    pub fn records(&self) -> &[MdxMetadataRecord] {
        &self.records
    }

    /// The metadata blocks (`MDB` records), in record order.
    #[must_use]
    pub fn blocks(&self) -> &[Mdb] {
        &self.blocks
    }

    /// Whether no MDX metadata was collected.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.infos.is_empty()
            && self.strings.is_empty()
            && self.records.is_empty()
            && self.blocks.is_empty()
    }

    /// Append a metadata type declaration (`MDTInfo`).
    pub fn add_info(&mut self, info: MdtInfo) {
        self.infos.push(info);
    }

    /// Append a shared MDX string (`MDXStr`).
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn add_string(&mut self, value: String) -> Result<()> {
        if value.encode_utf16().count() > MAX_LP_WIDE_STRING_CHARS {
            return Err(Error::InvalidData(
                "MDXStr string exceeds the LPWideString length limit".to_string(),
            ));
        }
        self.strings.push(value);
        Ok(())
    }

    /// Append an MDX metadata record, validating its `MDXStrIndex`
    /// references against the shared strings collected so far.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn add_record(&mut self, record: MdxMetadataRecord) -> Result<()> {
        for index in record.referenced_string_indexes() {
            self.validate_string_index(index)?;
        }
        self.records.push(record);
        Ok(())
    }

    /// Append a metadata block (`MDB`), validating its `MDir` references
    /// against the types and records collected so far.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn add_block(&mut self, block: Mdb) -> Result<()> {
        for entry in &block.entries {
            if entry.info_index < 1
                || crate::utils::sign_extend_i32_to_usize(entry.info_index) > self.infos.len()
            {
                return Err(invalid(
                    MDB_RECORD_TYPE,
                    format!(
                        "MDir.imdt {} is out of range of the MDTInfo collection",
                        entry.info_index
                    ),
                ));
            }
            if entry.metadata_index as usize >= self.records.len() {
                return Err(invalid(
                    MDB_RECORD_TYPE,
                    format!(
                        "MDir.mdd {} is out of range of the MDX metadata record collection",
                        entry.metadata_index
                    ),
                ));
            }
        }
        self.blocks.push(block);
        Ok(())
    }

    fn validate_string_index(&self, index: i32) -> Result<()> {
        if index < 0 || crate::utils::sign_extend_i32_to_usize(index) >= self.strings.len() {
            return Err(invalid(
                MDX_STR_RECORD_TYPE,
                format!(
                    "MDXStrIndex {index} is out of range of the MDXStr collection ({} strings)",
                    self.strings.len()
                ),
            ));
        }
        Ok(())
    }

    /// Parse one `METADATA` record (payload including any trailing
    /// `ContinueFrt12` bodies concatenated) into the collection.
    pub(crate) fn push_record(&mut self, record_type: u16, payload: &[u8]) -> Result<()> {
        match record_type {
            MDT_INFO_RECORD_TYPE => self.add_info(MdtInfo::parse(payload)?),
            MDX_STR_RECORD_TYPE => {
                self.strings.push(parse_lp_wide_string(
                    record_body(payload, record_type)?,
                    record_type,
                )?);
            },
            MDX_TUPLE_RECORD_TYPE => {
                self.add_record(MdxMetadataRecord::Tuple(MdxTuple::parse(payload)?))?;
            },
            MDX_SET_RECORD_TYPE => {
                self.add_record(MdxMetadataRecord::Set(MdxSet::parse(payload)?))?;
            },
            MDX_PROP_RECORD_TYPE => {
                self.add_record(MdxMetadataRecord::Prop(MdxProp::parse(payload)?))?;
            },
            MDX_KPI_RECORD_TYPE => {
                self.add_record(MdxMetadataRecord::Kpi(MdxKpi::parse(payload)?))?;
            },
            MDB_RECORD_TYPE => self.add_block(Mdb::parse(payload)?)?,
            other => {
                return Err(Error::UnexpectedRecordType {
                    expected: MDT_INFO_RECORD_TYPE,
                    found: other,
                });
            },
        }
        Ok(())
    }

    /// Serialize the whole collection in ABNF order as record payloads
    /// (`FrtHeader` included, not yet chunked into BIFF records).
    pub(crate) fn to_record_payloads(&self) -> Result<Vec<(u16, Vec<u8>)>> {
        let mut payloads = Vec::new();
        for info in &self.infos {
            payloads.push((MDT_INFO_RECORD_TYPE, info.to_payload()?));
        }
        for value in &self.strings {
            let mut payload = frt_header_payload(MDX_STR_RECORD_TYPE);
            append_lp_wide_string(MDX_STR_RECORD_TYPE, value, &mut payload)?;
            payloads.push((MDX_STR_RECORD_TYPE, payload));
        }
        for record in &self.records {
            let record_type = match record {
                MdxMetadataRecord::Tuple(_) => MDX_TUPLE_RECORD_TYPE,
                MdxMetadataRecord::Set(_) => MDX_SET_RECORD_TYPE,
                MdxMetadataRecord::Prop(_) => MDX_PROP_RECORD_TYPE,
                MdxMetadataRecord::Kpi(_) => MDX_KPI_RECORD_TYPE,
            };
            payloads.push((record_type, record.to_payload()?));
        }
        for block in &self.blocks {
            payloads.push((MDB_RECORD_TYPE, block.to_payload()?));
        }
        Ok(payloads)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn payload(record_type: u16, body: &[u8]) -> Vec<u8> {
        let mut data = frt_header_payload(record_type);
        data.extend_from_slice(body);
        data
    }

    fn wide_string(value: &str) -> Vec<u8> {
        let units: Vec<u16> = value.encode_utf16().collect();
        let mut data = Vec::new();
        data.extend_from_slice(&(units.len() as u16).to_le_bytes());
        for unit in units {
            data.extend_from_slice(&unit.to_le_bytes());
        }
        data
    }

    #[test]
    fn mdt_info_parses_flags_and_name() {
        let mut body = (F_COPY | F_PASTE_VALUES | F_CELL_META)
            .to_le_bytes()
            .to_vec();
        body.extend_from_slice(&wide_string("MDXValueMetadata"));
        let info = MdtInfo::parse(&payload(MDT_INFO_RECORD_TYPE, &body)).unwrap();
        assert_eq!(info.name, "MDXValueMetadata");
        assert!(info.flags.copied_with_cell());
        assert!(info.flags.paste_values());
        assert!(info.flags.is_cell_metadata());
        assert!(!info.flags.ghost_row());
        assert!(!info.flags.paste_all());
        assert!(!info.flags.preserved_on_edit());
        assert!(!info.flags.preserved_on_delete());
        assert!(!info.flags.paste_formulas());
        assert!(!info.flags.paste_formats());
        assert!(!info.flags.paste_comments());
        assert!(!info.flags.paste_data_validation());
        assert!(!info.flags.paste_borders());
        assert!(!info.flags.paste_col_widths());
        assert!(!info.flags.paste_number_formats());
        assert!(!info.flags.preserved_on_merge());
        assert!(!info.flags.split_first());
        assert!(!info.flags.split_all());
        assert!(!info.flags.preserved_on_row_col_shift());
        assert!(!info.flags.preserved_on_clear_all());
        assert!(!info.flags.preserved_on_clear_formats());
        assert!(!info.flags.preserved_on_clear_contents());
        assert!(!info.flags.preserved_on_clear_comments());
        assert!(!info.flags.preserved_on_assign());
        assert!(!info.flags.preserved_on_coerce());
        assert!(!info.flags.adjusted_on_move());
        assert!(!info.flags.ghost_column());
    }

    #[test]
    fn mdt_info_rejects_truncated_and_mismatched_strings() {
        assert!(MdtInfo::parse(&payload(MDT_INFO_RECORD_TYPE, &[0; 3])).is_err());
        // Declared character count exceeds the body.
        let mut body = 0u32.to_le_bytes().to_vec();
        body.extend_from_slice(&5u16.to_le_bytes());
        body.extend_from_slice(&[0; 4]);
        assert!(MdtInfo::parse(&payload(MDT_INFO_RECORD_TYPE, &body)).is_err());
        // Wrong FrtHeader record type echo.
        let mut body = 0u32.to_le_bytes().to_vec();
        body.extend_from_slice(&wide_string("x"));
        let mut data = payload(MDX_STR_RECORD_TYPE, &body);
        assert!(MdtInfo::parse(&data).is_err());
        data[0] = 0x84;
        data[1] = 0x08;
        assert!(MdtInfo::parse(&data).is_ok());
    }

    #[test]
    fn mdx_str_parses_shared_string() {
        let body = wide_string("Adventure Works");
        let mut metadata = MdxMetadata::default();
        metadata
            .push_record(MDX_STR_RECORD_TYPE, &payload(MDX_STR_RECORD_TYPE, &body))
            .unwrap();
        assert_eq!(metadata.strings(), &["Adventure Works".to_string()]);
    }

    #[test]
    fn mdx_tuple_parses_and_validates_function() {
        let mut body = 0i32.to_le_bytes().to_vec();
        body.push(CubeFunction::CubeMember.code());
        body.extend_from_slice(&2i32.to_le_bytes());
        body.extend_from_slice(&1i32.to_le_bytes());
        body.extend_from_slice(&2i32.to_le_bytes());
        let tuple = MdxTuple::parse(&payload(MDX_TUPLE_RECORD_TYPE, &body)).unwrap();
        assert_eq!(tuple.connection_name_index, 0);
        assert_eq!(tuple.function, CubeFunction::CubeMember);
        assert_eq!(tuple.string_indexes, vec![1, 2]);

        // CubeSet is not a valid tuple source function.
        body[4] = CubeFunction::CubeSet.code();
        assert!(MdxTuple::parse(&payload(MDX_TUPLE_RECORD_TYPE, &body)).is_err());
        // Negative count and mismatched rgistr length.
        body[4] = CubeFunction::CubeValue.code();
        body[5..9].copy_from_slice(&(-1i32).to_le_bytes());
        assert!(MdxTuple::parse(&payload(MDX_TUPLE_RECORD_TYPE, &body)).is_err());
        body[5..9].copy_from_slice(&3i32.to_le_bytes());
        assert!(MdxTuple::parse(&payload(MDX_TUPLE_RECORD_TYPE, &body)).is_err());
        // Truncated fixed part.
        assert!(MdxTuple::parse(&payload(MDX_TUPLE_RECORD_TYPE, &[0; 8])).is_err());
    }

    #[test]
    fn mdx_set_parses_sort_order() {
        let mut body = 0i32.to_le_bytes().to_vec();
        body.push(CubeFunction::CubeSetCount.code());
        body.push(MdxSetSortOrder::NaturalDescending.code());
        body.extend_from_slice(&3i32.to_le_bytes());
        body.extend_from_slice(&0i32.to_le_bytes());
        let set = MdxSet::parse(&payload(MDX_SET_RECORD_TYPE, &body)).unwrap();
        assert_eq!(set.function, CubeFunction::CubeSetCount);
        assert_eq!(set.sort_order, MdxSetSortOrder::NaturalDescending);
        assert_eq!(set.set_definition_index, 3);
        assert!(set.string_indexes.is_empty());

        body[5] = 0x07; // invalid SD_SetSortOrder
        assert!(MdxSet::parse(&payload(MDX_SET_RECORD_TYPE, &body)).is_err());
        body[5] = MdxSetSortOrder::Ascending.code();
        body[4] = CubeFunction::CubeMember.code(); // not a set function
        assert!(MdxSet::parse(&payload(MDX_SET_RECORD_TYPE, &body)).is_err());
    }

    #[test]
    fn mdx_prop_parses_fixed_body() {
        let mut body = 0i32.to_le_bytes().to_vec();
        body.push(CubeFunction::CubeMemberProperty.code());
        body.extend_from_slice(&1i32.to_le_bytes());
        body.extend_from_slice(&2i32.to_le_bytes());
        let prop = MdxProp::parse(&payload(MDX_PROP_RECORD_TYPE, &body)).unwrap();
        assert_eq!(prop.member_index, 1);
        assert_eq!(prop.property_index, 2);

        body.push(0); // trailing byte
        assert!(MdxProp::parse(&payload(MDX_PROP_RECORD_TYPE, &body)).is_err());
        body.pop();
        body[4] = CubeFunction::CubeValue.code();
        assert!(MdxProp::parse(&payload(MDX_PROP_RECORD_TYPE, &body)).is_err());
    }

    #[test]
    fn mdx_kpi_parses_fixed_body() {
        let mut body = 0i32.to_le_bytes().to_vec();
        body.push(CubeFunction::CubeKpiProperty.code());
        body.push(KpiProperty::Status.code());
        body.extend_from_slice(&1i32.to_le_bytes());
        body.extend_from_slice(&2i32.to_le_bytes());
        let kpi = MdxKpi::parse(&payload(MDX_KPI_RECORD_TYPE, &body)).unwrap();
        assert_eq!(kpi.kpi_property, KpiProperty::Status);
        assert_eq!(kpi.kpi_name_index, 1);
        assert_eq!(kpi.member_kpi_index, 2);

        body[5] = 0x00; // invalid KPIProp
        assert!(MdxKpi::parse(&payload(MDX_KPI_RECORD_TYPE, &body)).is_err());
        body[5] = KpiProperty::CurrentTimeMember.code();
        body[4] = CubeFunction::CubeMemberProperty.code();
        assert!(MdxKpi::parse(&payload(MDX_KPI_RECORD_TYPE, &body)).is_err());
    }

    #[test]
    fn mdb_parses_dir_array() {
        let mut body = 1i32.to_le_bytes().to_vec();
        body.extend_from_slice(&0u32.to_le_bytes());
        body.extend_from_slice(&2i32.to_le_bytes());
        body.extend_from_slice(&3u32.to_le_bytes());
        let block = Mdb::parse(&payload(MDB_RECORD_TYPE, &body)).unwrap();
        assert_eq!(
            block.entries,
            vec![
                MdxMetadataDir {
                    info_index: 1,
                    metadata_index: 0
                },
                MdxMetadataDir {
                    info_index: 2,
                    metadata_index: 3
                },
            ]
        );

        body.push(0); // not a whole MDir array
        assert!(Mdb::parse(&payload(MDB_RECORD_TYPE, &body)).is_err());
    }

    #[test]
    fn metadata_collection_validates_cross_record_indexes() {
        let mut metadata = MdxMetadata::default();
        assert!(metadata.is_empty());

        metadata.add_info(MdtInfo {
            flags: MdtInfoFlags::from_bits(F_COPY),
            name: "ValueMetadata".to_string(),
        });
        metadata.add_string("connection".to_string()).unwrap();
        metadata
            .add_string("[Product].[All Products]".to_string())
            .unwrap();

        // A tuple referencing a string that does not exist yet is rejected.
        let tuple = MdxMetadataRecord::Tuple(MdxTuple {
            connection_name_index: 0,
            function: CubeFunction::CubeValue,
            string_indexes: vec![1, 2],
        });
        assert!(metadata.add_record(tuple.clone()).is_err());

        metadata
            .add_string("[Measures].[Sales]".to_string())
            .unwrap();
        metadata.add_record(tuple).unwrap();

        // MDir indexes must reference collected types and records; MDTInfo
        // indexes are one-based.
        let bad_info = Mdb {
            entries: vec![MdxMetadataDir {
                info_index: 0,
                metadata_index: 0,
            }],
        };
        assert!(metadata.add_block(bad_info).is_err());
        let bad_record = Mdb {
            entries: vec![MdxMetadataDir {
                info_index: 1,
                metadata_index: 1,
            }],
        };
        assert!(metadata.add_block(bad_record).is_err());
        let good = Mdb {
            entries: vec![MdxMetadataDir {
                info_index: 1,
                metadata_index: 0,
            }],
        };
        metadata.add_block(good).unwrap();
        assert!(!metadata.is_empty());
        assert_eq!(metadata.records().len(), 1);
        assert_eq!(metadata.blocks().len(), 1);
    }

    #[test]
    fn push_record_rejects_unknown_record_type() {
        let mut metadata = MdxMetadata::default();
        assert!(metadata.push_record(0x0801, &[]).is_err());
    }

    #[test]
    fn payloads_round_trip_through_parse() {
        let mut metadata = MdxMetadata::default();
        metadata.add_info(MdtInfo {
            flags: MdtInfoFlags::from_bits(F_COPY | F_PASTE_ALL),
            name: "ValueMetadata".to_string(),
        });
        for value in ["conn", "set-definition", "member", "property"] {
            metadata.add_string(value.to_string()).unwrap();
        }
        metadata
            .add_record(MdxMetadataRecord::Set(MdxSet {
                connection_name_index: 0,
                function: CubeFunction::CubeSet,
                sort_order: MdxSetSortOrder::AlphaAscending,
                set_definition_index: 1,
                string_indexes: vec![2],
            }))
            .unwrap();
        metadata
            .add_record(MdxMetadataRecord::Kpi(MdxKpi {
                connection_name_index: 0,
                function: CubeFunction::CubeKpiProperty,
                kpi_property: KpiProperty::Trend,
                kpi_name_index: 2,
                member_kpi_index: 3,
            }))
            .unwrap();
        metadata
            .add_block(Mdb {
                entries: vec![
                    MdxMetadataDir {
                        info_index: 1,
                        metadata_index: 0,
                    },
                    MdxMetadataDir {
                        info_index: 1,
                        metadata_index: 1,
                    },
                ],
            })
            .unwrap();

        let payloads = metadata.to_record_payloads().unwrap();
        assert_eq!(payloads.len(), 8);
        assert_eq!(
            payloads
                .iter()
                .map(|(record_type, _)| *record_type)
                .collect::<Vec<_>>(),
            vec![
                MDT_INFO_RECORD_TYPE,
                MDX_STR_RECORD_TYPE,
                MDX_STR_RECORD_TYPE,
                MDX_STR_RECORD_TYPE,
                MDX_STR_RECORD_TYPE,
                MDX_SET_RECORD_TYPE,
                MDX_KPI_RECORD_TYPE,
                MDB_RECORD_TYPE,
            ]
        );

        let mut parsed = MdxMetadata::default();
        for (record_type, data) in &payloads {
            parsed.push_record(*record_type, data).unwrap();
        }
        assert_eq!(parsed, metadata);
    }
}
