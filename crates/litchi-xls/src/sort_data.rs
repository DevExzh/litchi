//! BIFF8 extended range-sort metadata.
//!
//! `SortData` ([MS-XLS] 2.4.264) stores its fixed fields in record 0x0895 and
//! places each `SortCond12` ([MS-XLS] 2.5.242) in a separate `ContinueFrt12`
//! record. This module deliberately models that record group as one value.

use crate::{XlsError, XlsResult};
use std::io::Write;
use std::ops::RangeInclusive;

/// `SortData` record identifier.
pub const SORT_DATA_RECORD_TYPE: u16 = 0x0895;
/// `ContinueFrt12` record identifier.
pub const CONTINUE_FRT12_RECORD_TYPE: u16 = 0x087f;

const SORT_DATA_BODY_LEN: usize = 38;
const FRT_HEADER_LEN: usize = 12;
const SORT_CONDITION_FIXED_LEN: usize = 30;
const MAX_CONTINUE_RGB_LEN: usize = 8_212;
const MAX_ROW_INDEX: u32 = 0x000f_ffff;
const MAX_COLUMN_INDEX: u32 = 0x0000_3fff;
const FRT_REF_FLAG: u16 = 0x0001;
const FRT_ALERT_FLAG: u16 = 0x0002;
const FRT_KNOWN_FLAGS: u16 = FRT_REF_FLAG | FRT_ALERT_FLAG;

fn invalid(message: impl Into<String>) -> XlsError {
    XlsError::InvalidData(message.into())
}

const fn allocation(context: &'static str) -> XlsError {
    XlsError::Allocation(context)
}

fn copy_bytes(data: &[u8], context: &'static str) -> XlsResult<Vec<u8>> {
    let mut copy = Vec::new();
    copy.try_reserve_exact(data.len())
        .map_err(|_| allocation(context))?;
    copy.extend_from_slice(data);
    Ok(copy)
}

fn read_u16(data: &[u8], offset: usize) -> XlsResult<u16> {
    let bytes = data
        .get(offset..offset + 2)
        .ok_or(XlsError::InvalidLength {
            expected: offset + 2,
            found: data.len(),
        })?;
    Ok(u16::from_le_bytes([bytes[0], bytes[1]]))
}

fn read_u32(data: &[u8], offset: usize) -> XlsResult<u32> {
    let bytes = data
        .get(offset..offset + 4)
        .ok_or(XlsError::InvalidLength {
            expected: offset + 4,
            found: data.len(),
        })?;
    Ok(u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]))
}

fn read_i32(data: &[u8], offset: usize) -> XlsResult<i32> {
    Ok(read_u32(data, offset)? as i32)
}

/// Checked native value for the signed four-byte `[MS-XLS]` `Rw12` field.
///
/// The private representation is an `i32` because that is the wire scalar's
/// signed type. Its invariant is the non-negative `Rw12` domain from
/// `[MS-XLS]` 2.5.228, not the complete `i32` domain.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
struct Rw12(i32);

impl Rw12 {
    fn new(value: u32) -> XlsResult<Self> {
        if value > MAX_ROW_INDEX {
            return Err(invalid("sort row exceeds the Rw12 maximum"));
        }
        let value = i32::try_from(value)
            .map_err(|_| invalid("sort row cannot be represented by signed Rw12"))?;
        Ok(Self(value))
    }

    const fn index(self) -> u32 {
        self.0 as u32
    }

    fn from_wire(value: i32) -> XlsResult<Self> {
        if value < 0 {
            return Err(invalid("sort Rw12 is negative"));
        }
        Self::new(value as u32)
    }

    const fn wire_value(self) -> i32 {
        self.0
    }

    fn parse(data: &[u8], offset: usize) -> XlsResult<Self> {
        Self::from_wire(read_i32(data, offset)?)
    }

    fn write_to(self, output: &mut Vec<u8>) {
        output.extend_from_slice(&self.wire_value().to_le_bytes());
    }
}

/// Checked native value for the signed four-byte `[MS-XLS]` `Col12` field.
///
/// The wire field is four bytes, but its valid signed domain fits in an `i16`.
/// The retained representation therefore does not pay for impossible values.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
struct Col12(i16);

impl Col12 {
    fn new(value: u32) -> XlsResult<Self> {
        if value > MAX_COLUMN_INDEX {
            return Err(invalid("sort column exceeds the Col12 maximum"));
        }
        let value = i16::try_from(value)
            .map_err(|_| invalid("sort column cannot be represented by signed Col12"))?;
        Ok(Self(value))
    }

    const fn index(self) -> u16 {
        self.0 as u16
    }

    fn from_wire(value: i32) -> XlsResult<Self> {
        if value < 0 {
            return Err(invalid("sort Col12 is negative"));
        }
        Self::new(value as u32)
    }

    const fn wire_value(self) -> i32 {
        self.0 as i32
    }

    fn parse(data: &[u8], offset: usize) -> XlsResult<Self> {
        Self::from_wire(read_i32(data, offset)?)
    }

    fn write_to(self, output: &mut Vec<u8>) {
        output.extend_from_slice(&self.wire_value().to_le_bytes());
    }
}

/// A checked row in the extended `Rw12` domain (`0..=1_048_575`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
pub struct Row(Rw12);

impl Row {
    /// Check a zero-based row index without applying the smaller BIFF8 cell grid.
    pub fn new(index: u32) -> XlsResult<Self> {
        Rw12::new(index).map(Self)
    }

    /// Return the zero-based row index.
    pub const fn index(self) -> u32 {
        self.0.index()
    }
}

/// A checked column in the extended `Col12` domain (`0..=16_383`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
pub struct Col(Col12);

impl Col {
    /// Check a zero-based column index without applying the smaller BIFF8 cell grid.
    pub fn new(index: u32) -> XlsResult<Self> {
        Col12::new(index).map(Self)
    }

    /// Return the zero-based column index.
    pub const fn index(self) -> u16 {
        self.0.index()
    }
}

/// A validated `RFX` cell range used by extended sorting.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Range {
    first_row: Rw12,
    last_row: Rw12,
    first_col: Col12,
    last_col: Col12,
}

impl Range {
    /// Create an inclusive range with the complete `Rw12` and `Col12` domains.
    ///
    /// `Range::new(1..=20, 0..=4)` is rows 2 through 21 and columns A through E.
    pub fn new(rows: RangeInclusive<u32>, cols: RangeInclusive<u32>) -> XlsResult<Self> {
        let first_row = Rw12::new(*rows.start())?;
        let last_row = Rw12::new(*rows.end())?;
        let first_col = Col12::new(*cols.start())?;
        let last_col = Col12::new(*cols.end())?;
        if first_row > last_row {
            return Err(invalid("sort range first row exceeds last row"));
        }
        if first_col > last_col {
            return Err(invalid("sort range first column exceeds last column"));
        }
        Ok(Self {
            first_row,
            last_row,
            first_col,
            last_col,
        })
    }

    /// Return the first row.
    pub const fn first_row(self) -> Row {
        Row(self.first_row)
    }

    /// Return the last row.
    pub const fn last_row(self) -> Row {
        Row(self.last_row)
    }

    /// Return the first column.
    pub const fn first_col(self) -> Col {
        Col(self.first_col)
    }

    /// Return the last column.
    pub const fn last_col(self) -> Col {
        Col(self.last_col)
    }

    const fn contains(self, other: Self) -> bool {
        self.first_row.0 <= other.first_row.0
            && other.last_row.0 <= self.last_row.0
            && self.first_col.0 <= other.first_col.0
            && other.last_col.0 <= self.last_col.0
    }

    fn parse(data: &[u8], offset: usize) -> XlsResult<Self> {
        let first_row = Rw12::parse(data, offset)?;
        let last_row = Rw12::parse(data, offset + 4)?;
        let first_col = Col12::parse(data, offset + 8)?;
        let last_col = Col12::parse(data, offset + 12)?;
        if first_row > last_row {
            return Err(invalid("sort range first row exceeds last row"));
        }
        if first_col > last_col {
            return Err(invalid("sort range first column exceeds last column"));
        }
        Ok(Self {
            first_row,
            last_row,
            first_col,
            last_col,
        })
    }

    fn write_to(self, output: &mut Vec<u8>) {
        self.first_row.write_to(output);
        self.last_row.write_to(output);
        self.first_col.write_to(output);
        self.last_col.write_to(output);
    }

    fn bytes(self) -> [u8; 16] {
        let mut bytes = [0; 16];
        bytes[..4].copy_from_slice(&self.first_row.wire_value().to_le_bytes());
        bytes[4..8].copy_from_slice(&self.last_row.wire_value().to_le_bytes());
        bytes[8..12].copy_from_slice(&self.first_col.wire_value().to_le_bytes());
        bytes[12..].copy_from_slice(&self.last_col.wire_value().to_le_bytes());
        bytes
    }
}

/// Axis whose cells are reordered by the sort.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Axis {
    /// Top-to-bottom: reorder rows using column keys.
    Rows,
    /// Left-to-right: reorder columns using row keys.
    Cols,
}

/// Character-order versus locale-specific alternate sorting.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Method {
    CharacterOrder,
    Alternate,
}

/// Object which owns the sort field (`sfp` and `idParent`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Parent {
    Sheet,
    Table { id: u32 },
    AutoFilter,
    QueryTable { index: u32 },
}

/// A DXF table index used by color-based sorting.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Dxf(u32);

impl Dxf {
    pub const fn new(index: u32) -> Self {
        Self(index)
    }

    pub const fn index(self) -> u32 {
        self.0
    }
}

/// Icon sets allowed by the `KPISets` enumeration.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum IconSet {
    NoIcon,
    ThreeArrows,
    ThreeArrowsGray,
    ThreeFlags,
    ThreeTrafficLights1,
    ThreeTrafficLights2,
    ThreeSigns,
    ThreeSymbols,
    ThreeSymbols2,
    FourArrows,
    FourArrowsGray,
    FourRedToBlack,
    FourRating,
    FourTrafficLights,
    FiveArrows,
    FiveArrowsGray,
    FiveRating,
    FiveQuarters,
}

impl IconSet {
    fn code(self) -> u32 {
        match self {
            Self::NoIcon => u32::MAX,
            Self::ThreeArrows => 0,
            Self::ThreeArrowsGray => 1,
            Self::ThreeFlags => 2,
            Self::ThreeTrafficLights1 => 3,
            Self::ThreeTrafficLights2 => 4,
            Self::ThreeSigns => 5,
            Self::ThreeSymbols => 6,
            Self::ThreeSymbols2 => 7,
            Self::FourArrows => 8,
            Self::FourArrowsGray => 9,
            Self::FourRedToBlack => 10,
            Self::FourRating => 11,
            Self::FourTrafficLights => 12,
            Self::FiveArrows => 13,
            Self::FiveArrowsGray => 14,
            Self::FiveRating => 15,
            Self::FiveQuarters => 16,
        }
    }

    fn from_code(code: u32) -> XlsResult<Self> {
        Ok(match code {
            u32::MAX => Self::NoIcon,
            0 => Self::ThreeArrows,
            1 => Self::ThreeArrowsGray,
            2 => Self::ThreeFlags,
            3 => Self::ThreeTrafficLights1,
            4 => Self::ThreeTrafficLights2,
            5 => Self::ThreeSigns,
            6 => Self::ThreeSymbols,
            7 => Self::ThreeSymbols2,
            8 => Self::FourArrows,
            9 => Self::FourArrowsGray,
            10 => Self::FourRedToBlack,
            11 => Self::FourRating,
            12 => Self::FourTrafficLights,
            13 => Self::FiveArrows,
            14 => Self::FiveArrowsGray,
            15 => Self::FiveRating,
            16 => Self::FiveQuarters,
            _ => return Err(invalid("SortCond12 contains an unknown KPISets value")),
        })
    }
}

/// Icon ordinal used by icon-set sorting.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Icon {
    NoIcon,
    First,
    Second,
    Third,
    Fourth,
    Fifth,
}

impl Icon {
    fn code(self) -> i32 {
        match self {
            Self::NoIcon => -1,
            Self::First => 0,
            Self::Second => 1,
            Self::Third => 2,
            Self::Fourth => 3,
            Self::Fifth => 4,
        }
    }

    fn from_code(code: i32) -> XlsResult<Self> {
        Ok(match code {
            -1 => Self::NoIcon,
            0 => Self::First,
            1 => Self::Second,
            2 => Self::Third,
            3 => Self::Fourth,
            4 => Self::Fifth,
            _ => return Err(invalid("SortCond12 icon index is outside -1 through 4")),
        })
    }
}

/// The criterion used by a `SortCond12` entry.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum On {
    Values { custom_list: Option<String> },
    CellColor { differential_format: Dxf },
    FontColor { differential_format: Dxf },
    Icon { set: IconSet, icon: Icon },
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum KeyRange {
    Col {
        first_row: Rw12,
        last_row: Rw12,
        col: Col12,
    },
    Row {
        row: Rw12,
        first_col: Col12,
        last_col: Col12,
    },
}

impl KeyRange {
    const fn axis(self) -> Axis {
        match self {
            Self::Col { .. } => Axis::Rows,
            Self::Row { .. } => Axis::Cols,
        }
    }

    const fn range(self) -> Range {
        match self {
            Self::Col {
                first_row,
                last_row,
                col,
            } => Range {
                first_row,
                last_row,
                first_col: col,
                last_col: col,
            },
            Self::Row {
                row,
                first_col,
                last_col,
            } => Range {
                first_row: row,
                last_row: row,
                first_col,
                last_col,
            },
        }
    }
}

/// One checked sort key.
///
/// Use [`Key::col`] when rows are reordered and [`Key::row`] when columns are
/// reordered. A key cannot represent an ambiguous two-dimensional rectangle.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Key {
    range: KeyRange,
    descending: bool,
    on: On,
}

impl Key {
    /// Create a column key for a top-to-bottom row sort.
    pub fn col(range: Range, descending: bool, on: On) -> XlsResult<Self> {
        if range.first_col != range.last_col {
            return Err(invalid("column sort key must contain exactly one column"));
        }
        Self::new(
            KeyRange::Col {
                first_row: range.first_row,
                last_row: range.last_row,
                col: range.first_col,
            },
            descending,
            on,
        )
    }

    /// Create a row key for a left-to-right column sort.
    pub fn row(range: Range, descending: bool, on: On) -> XlsResult<Self> {
        if range.first_row != range.last_row {
            return Err(invalid("row sort key must contain exactly one row"));
        }
        Self::new(
            KeyRange::Row {
                row: range.first_row,
                first_col: range.first_col,
                last_col: range.last_col,
            },
            descending,
            on,
        )
    }

    fn new(range: KeyRange, descending: bool, on: On) -> XlsResult<Self> {
        validate_on(&on)?;
        Ok(Self {
            range,
            descending,
            on,
        })
    }

    /// Return the axis this key can sort.
    pub const fn axis(&self) -> Axis {
        self.range.axis()
    }

    /// Return the key's inclusive range.
    pub const fn range(&self) -> Range {
        self.range.range()
    }

    /// Return whether this key sorts descending.
    pub const fn descending(&self) -> bool {
        self.descending
    }

    /// Borrow the criterion applied by this key.
    pub const fn on(&self) -> &On {
        &self.on
    }
}

/// Complete extended sorting metadata represented by one BIFF record group.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Config {
    range: Range,
    axis: Axis,
    case_sensitive: bool,
    method: Method,
    parent: Parent,
    keys: Vec<Key>,
}

impl Config {
    /// Create an empty top-to-bottom sort for `range`.
    pub fn new(range: Range, parent: Parent) -> Self {
        Self {
            range,
            axis: Axis::Rows,
            case_sensitive: false,
            method: Method::CharacterOrder,
            parent,
            keys: Vec::new(),
        }
    }

    /// Replace the reordered axis after validating every retained key.
    ///
    /// On failure, the configuration is unchanged.
    pub fn put_axis(&mut self, axis: Axis) -> XlsResult<Axis> {
        if self.keys.iter().any(|key| key.axis() != axis) {
            return Err(invalid(
                "sort axis does not match its retained key direction",
            ));
        }
        Ok(std::mem::replace(&mut self.axis, axis))
    }

    /// Set case-sensitive comparison.
    pub fn set_case(&mut self, case_sensitive: bool) {
        self.case_sensitive = case_sensitive;
    }

    /// Select the character-order or alternate comparison method.
    pub fn set_method(&mut self, method: Method) {
        self.method = method;
    }

    /// Append a checked key.
    ///
    /// The key direction and containment are checked before `self` changes.
    pub fn add(&mut self, key: Key) -> XlsResult<()> {
        if key.axis() != self.axis {
            return Err(invalid("sort key direction does not match the sort axis"));
        }
        if !self.range.contains(key.range()) {
            return Err(invalid("sort key is outside the range being sorted"));
        }
        let next_len = self
            .keys
            .len()
            .checked_add(1)
            .ok_or_else(|| allocation("computing the next sort-key count"))?;
        u32::try_from(next_len).map_err(|_| invalid("SortData has more than u32::MAX keys"))?;
        self.keys
            .try_reserve(1)
            .map_err(|_| allocation("reserving storage for a sort key"))?;
        self.keys.push(key);
        Ok(())
    }

    /// Return the range being sorted.
    pub const fn range(&self) -> Range {
        self.range
    }

    /// Return the reordered axis.
    pub const fn axis(&self) -> Axis {
        self.axis
    }

    /// Return whether comparison is case-sensitive.
    pub const fn case_sensitive(&self) -> bool {
        self.case_sensitive
    }

    /// Return the comparison method.
    pub const fn method(&self) -> Method {
        self.method
    }

    /// Return the owning object.
    pub const fn parent(&self) -> Parent {
        self.parent
    }

    /// Borrow the keys in priority order.
    pub fn keys(&self) -> &[Key] {
        &self.keys
    }

    /// Write the `SortData` record followed by one `ContinueFrt12` per key.
    pub(crate) fn write_biff_records<W: Write>(&self, writer: &mut W) -> XlsResult<()> {
        let condition_count = u32::try_from(self.keys.len())
            .map_err(|_| invalid("SortData has more than u32::MAX keys"))?;
        let (parent_kind, parent_id) = match self.parent {
            Parent::Sheet => (0u16, 0),
            Parent::Table { id } => (1, id),
            Parent::AutoFilter => (2, 0),
            Parent::QueryTable { index } => (3, index),
        };
        let flags = u16::from(self.axis == Axis::Cols)
            | (u16::from(self.case_sensitive) << 1)
            | (u16::from(self.method == Method::Alternate) << 2)
            | (parent_kind << 3);

        write_record_header(writer, SORT_DATA_RECORD_TYPE, SORT_DATA_BODY_LEN)?;
        writer.write_all(&SORT_DATA_RECORD_TYPE.to_le_bytes())?;
        writer.write_all(&0u16.to_le_bytes())?;
        writer.write_all(&[0; 8])?;
        writer.write_all(&flags.to_le_bytes())?;
        writer.write_all(&self.range.bytes())?;
        writer.write_all(&condition_count.to_le_bytes())?;
        writer.write_all(&parent_id.to_le_bytes())?;

        for key in &self.keys {
            let body = encode_key(key)?;
            write_record_header(
                writer,
                CONTINUE_FRT12_RECORD_TYPE,
                FRT_HEADER_LEN + body.len(),
            )?;
            writer.write_all(&CONTINUE_FRT12_RECORD_TYPE.to_le_bytes())?;
            writer.write_all(&0u16.to_le_bytes())?;
            writer.write_all(&[0; 8])?;
            writer.write_all(&body)?;
        }
        Ok(())
    }
}

fn write_record_header<W: Write>(writer: &mut W, record_type: u16, len: usize) -> XlsResult<()> {
    let len = u16::try_from(len).map_err(|_| invalid("BIFF record body exceeds u16::MAX"))?;
    writer.write_all(&record_type.to_le_bytes())?;
    writer.write_all(&len.to_le_bytes())?;
    Ok(())
}

fn validate_sort_data_frt_header(data: &[u8]) -> XlsResult<()> {
    let echoed_type = read_u16(data, 0)?;
    if echoed_type != SORT_DATA_RECORD_TYPE {
        return Err(XlsError::UnexpectedRecordType {
            expected: SORT_DATA_RECORD_TYPE,
            found: echoed_type,
        });
    }
    let flags = read_u16(data, 2)?;
    if flags & !FRT_KNOWN_FLAGS != 0 {
        return Err(invalid(
            "SortData FrtHeader contains nonzero reserved flag bits",
        ));
    }
    if flags != 0 {
        return Err(invalid(
            "SortData FrtHeader fFrtRef and fFrtAlert must be zero",
        ));
    }
    if data[4..FRT_HEADER_LEN].iter().any(|byte| *byte != 0) {
        return Err(invalid(
            "SortData FrtHeader contains nonzero reserved bytes",
        ));
    }
    Ok(())
}

fn validate_continue_frt_header(data: &[u8]) -> XlsResult<()> {
    let echoed_type = read_u16(data, 0)?;
    if echoed_type != CONTINUE_FRT12_RECORD_TYPE {
        return Err(XlsError::UnexpectedRecordType {
            expected: CONTINUE_FRT12_RECORD_TYPE,
            found: echoed_type,
        });
    }
    let flags = read_u16(data, 2)?;
    if flags & !FRT_KNOWN_FLAGS != 0 {
        return Err(invalid(
            "ContinueFrt12 FrtRefHeader contains nonzero reserved flag bits",
        ));
    }
    if flags & FRT_ALERT_FLAG != 0 {
        return Err(invalid("ContinueFrt12 fFrtAlert must be zero"));
    }
    let reference = &data[4..FRT_HEADER_LEN];
    if flags & FRT_REF_FLAG == 0 {
        if reference.iter().any(|byte| *byte != 0) {
            return Err(invalid(
                "ContinueFrt12 has a reference while fFrtRef is zero",
            ));
        }
    } else {
        validate_ref8(reference)?;
    }
    Ok(())
}

fn validate_ref8(data: &[u8]) -> XlsResult<()> {
    let first_row = read_u16(data, 0)?;
    let last_row = read_u16(data, 2)?;
    let first_col = read_u16(data, 4)?;
    let last_col = read_u16(data, 6)?;
    if first_row > last_row {
        return Err(invalid("ContinueFrt12 Ref8 first row exceeds last row"));
    }
    if first_col > last_col {
        return Err(invalid(
            "ContinueFrt12 Ref8 first column exceeds last column",
        ));
    }
    if last_col > u16::from(u8::MAX) {
        return Err(invalid(
            "ContinueFrt12 Ref8 column exceeds the BIFF8 cell-grid maximum",
        ));
    }
    Ok(())
}

fn validate_on(on: &On) -> XlsResult<()> {
    let On::Values {
        custom_list: Some(custom_list),
    } = on
    else {
        return Ok(());
    };
    if custom_list.is_empty() {
        return Err(invalid(
            "sort custom list must be None rather than an ambiguous empty string",
        ));
    }

    let mut units = 0usize;
    let mut wide = false;
    for unit in custom_list.encode_utf16() {
        units = units
            .checked_add(1)
            .ok_or_else(|| invalid("SortCond12 custom-list length overflows usize"))?;
        wide |= unit > 0xff;
    }
    let encoded_len = units
        .checked_mul(if wide { 2 } else { 1 })
        .and_then(|len| len.checked_add(1))
        .ok_or_else(|| invalid("SortCond12 custom-list byte length overflows usize"))?;
    let body_len = SORT_CONDITION_FIXED_LEN
        .checked_add(encoded_len)
        .ok_or_else(|| invalid("SortCond12 encoded length overflows usize"))?;
    if body_len > MAX_CONTINUE_RGB_LEN {
        return Err(invalid("SortCond12 exceeds the ContinueFrt12 rgb limit"));
    }
    Ok(())
}

fn encode_key(key: &Key) -> XlsResult<Vec<u8>> {
    validate_on(&key.on)?;
    let (sort_on, cond_data, custom_list) = match &key.on {
        On::Values { custom_list } => (0u16, [0u8; 8], custom_list.as_deref()),
        On::CellColor {
            differential_format,
        } => {
            let mut data = [0u8; 8];
            data[..4].copy_from_slice(&differential_format.index().to_le_bytes());
            (1, data, None)
        },
        On::FontColor {
            differential_format,
        } => {
            let mut data = [0u8; 8];
            data[..4].copy_from_slice(&differential_format.index().to_le_bytes());
            (2, data, None)
        },
        On::Icon { set, icon } => {
            let mut data = [0u8; 8];
            data[..4].copy_from_slice(&set.code().to_le_bytes());
            data[4..].copy_from_slice(&icon.code().to_le_bytes());
            (3, data, None)
        },
    };
    let mut units = Vec::new();
    if let Some(value) = custom_list {
        let unit_count = value.encode_utf16().count();
        units
            .try_reserve_exact(unit_count)
            .map_err(|_| allocation("reserving SortCond12 custom-list storage"))?;
        units.extend(value.encode_utf16());
    }
    let char_count = u32::try_from(units.len())
        .map_err(|_| invalid("SortCond12 custom list exceeds u32::MAX UTF-16 code units"))?;
    let compressed = units.iter().all(|unit| *unit <= 0xff);
    let string_bytes = if units.is_empty() {
        0
    } else if compressed {
        1 + units.len()
    } else {
        1 + units.len() * 2
    };
    let body_len = SORT_CONDITION_FIXED_LEN
        .checked_add(string_bytes)
        .ok_or_else(|| invalid("SortCond12 encoded length overflow"))?;
    if body_len > MAX_CONTINUE_RGB_LEN {
        return Err(invalid("SortCond12 exceeds the ContinueFrt12 rgb limit"));
    }

    let mut output = Vec::new();
    output
        .try_reserve_exact(body_len)
        .map_err(|_| allocation("reserving SortCond12 record storage"))?;
    output.extend_from_slice(&((sort_on << 1) | u16::from(key.descending)).to_le_bytes());
    key.range().write_to(&mut output);
    output.extend_from_slice(&cond_data);
    output.extend_from_slice(&char_count.to_le_bytes());
    if !units.is_empty() {
        output.push(u8::from(!compressed));
        if compressed {
            output.extend(units.into_iter().map(|unit| unit as u8));
        } else {
            for unit in units {
                output.extend_from_slice(&unit.to_le_bytes());
            }
        }
    }
    Ok(output)
}

/// Parses a `SortData` payload and the following `ContinueFrt12` payloads.
///
/// Inputs exclude the standard four-byte BIFF record headers, consistent with
/// the rest of the XLS record parsers.
pub(crate) fn parse_sort_data(base: &[u8], continuations: &[&[u8]]) -> XlsResult<Config> {
    if base.len() != SORT_DATA_BODY_LEN {
        return Err(XlsError::InvalidLength {
            expected: SORT_DATA_BODY_LEN,
            found: base.len(),
        });
    }
    validate_sort_data_frt_header(base)?;
    let flags = read_u16(base, 12)?;
    let parent_kind = (flags >> 3) & 0x0007;
    let range = Range::parse(base, 14)?;
    let condition_count = usize::try_from(read_u32(base, 30)?)
        .map_err(|_| invalid("SortData condition count is not addressable"))?;
    if condition_count != continuations.len() {
        return Err(invalid(format!(
            "SortData declares {condition_count} conditions but {} continuations were supplied",
            continuations.len()
        )));
    }
    let parent_id = read_u32(base, 34)?;
    let parent = match parent_kind {
        0 => Parent::Sheet,
        1 => Parent::Table { id: parent_id },
        2 => Parent::AutoFilter,
        3 => Parent::QueryTable { index: parent_id },
        _ => return Err(invalid("SortData sfp is outside 0 through 3")),
    };
    let axis = if flags & 0x0001 != 0 {
        Axis::Cols
    } else {
        Axis::Rows
    };
    let mut keys = Vec::new();
    keys.try_reserve_exact(condition_count)
        .map_err(|_| allocation("reserving parsed SortData key storage"))?;
    let mut config = Config {
        range,
        axis,
        case_sensitive: flags & 0x0002 != 0,
        method: if flags & 0x0004 != 0 {
            Method::Alternate
        } else {
            Method::CharacterOrder
        },
        parent,
        keys,
    };
    for continuation in continuations {
        let key = parse_continuation(continuation, axis)?;
        config.add(key)?;
    }
    Ok(config)
}

#[derive(Debug)]
struct PendingSortData {
    base: Vec<u8>,
    continuations: Vec<Vec<u8>>,
    expected_conditions: usize,
}

/// Sequential record-group assembler used by the normal worksheet parser.
#[derive(Debug, Default)]
pub(crate) struct SortDataCollector {
    pending: Option<PendingSortData>,
}

impl SortDataCollector {
    pub(crate) fn feed_record(
        &mut self,
        record_type: u16,
        data: &[u8],
    ) -> XlsResult<Option<Config>> {
        if let Some(pending) = self.pending.as_mut() {
            if record_type != CONTINUE_FRT12_RECORD_TYPE {
                return Err(XlsError::InvalidRecord {
                    record_type,
                    message: format!(
                        "SortData must be followed immediately by {} ContinueFrt12 records",
                        pending.expected_conditions
                    ),
                });
            }
            if pending.continuations.len() >= pending.expected_conditions {
                return Err(invalid("SortData received too many ContinueFrt12 records"));
            }
            let continuation = copy_bytes(data, "reserving SortData continuation payload storage")?;
            pending.continuations.push(continuation);
            if pending.continuations.len() == pending.expected_conditions {
                let pending = self
                    .pending
                    .take()
                    .ok_or_else(|| invalid("SortData collector lost its pending record"))?;
                let mut continuations = Vec::new();
                continuations
                    .try_reserve_exact(pending.continuations.len())
                    .map_err(|_| allocation("reserving SortData continuation references"))?;
                continuations.extend(pending.continuations.iter().map(Vec::as_slice));
                return parse_sort_data(&pending.base, &continuations).map(Some);
            }
            return Ok(None);
        }

        if record_type != SORT_DATA_RECORD_TYPE {
            return Ok(None);
        }
        if data.len() != SORT_DATA_BODY_LEN {
            return Err(XlsError::InvalidLength {
                expected: SORT_DATA_BODY_LEN,
                found: data.len(),
            });
        }
        let expected_conditions = usize::try_from(read_u32(data, 30)?)
            .map_err(|_| invalid("SortData condition count is not addressable"))?;
        if expected_conditions == 0 {
            return parse_sort_data(data, &[]).map(Some);
        }
        let base = copy_bytes(data, "reserving SortData base-record storage")?;
        let mut continuations = Vec::new();
        continuations
            .try_reserve_exact(expected_conditions)
            .map_err(|_| allocation("reserving SortData continuation storage"))?;
        self.pending = Some(PendingSortData {
            base,
            continuations,
            expected_conditions,
        });
        Ok(None)
    }

    pub(crate) fn finish(self) -> XlsResult<()> {
        if let Some(pending) = self.pending {
            return Err(XlsError::InvalidRecord {
                record_type: SORT_DATA_RECORD_TYPE,
                message: format!(
                    "worksheet ended after {} of {} SortData conditions",
                    pending.continuations.len(),
                    pending.expected_conditions
                ),
            });
        }
        Ok(())
    }
}

fn parse_continuation(data: &[u8], axis: Axis) -> XlsResult<Key> {
    if data.len() < FRT_HEADER_LEN + SORT_CONDITION_FIXED_LEN {
        return Err(XlsError::InvalidLength {
            expected: FRT_HEADER_LEN + SORT_CONDITION_FIXED_LEN,
            found: data.len(),
        });
    }
    if data.len() - FRT_HEADER_LEN > MAX_CONTINUE_RGB_LEN {
        return Err(invalid("ContinueFrt12 rgb exceeds 8,212 bytes"));
    }
    validate_continue_frt_header(data)?;

    let body = &data[FRT_HEADER_LEN..];
    let flags = read_u16(body, 0)?;
    if flags & 0xffe0 != 0 {
        return Err(invalid("SortCond12 reserved flag bits are nonzero"));
    }
    let sort_on_code = (flags >> 1) & 0x000f;
    let range = Range::parse(body, 2)?;
    let data_value = read_u32(body, 18)?;
    let reserved_data = read_u32(body, 22)?;
    let char_count = read_i32(body, 26)?;
    if char_count < 0 {
        return Err(invalid("SortCond12 cchSt is negative"));
    }
    let char_count = char_count as usize;
    let sort_on = match sort_on_code {
        0 => {
            if data_value != 0 || reserved_data != 0 {
                return Err(invalid("value SortCond12 has nonzero CondDataValue fields"));
            }
            let custom_list = parse_custom_list(body, char_count)?;
            On::Values { custom_list }
        },
        1 | 2 => {
            if reserved_data != 0 || char_count != 0 || body.len() != SORT_CONDITION_FIXED_LEN {
                return Err(invalid("color SortCond12 has reserved or trailing data"));
            }
            let differential_format = Dxf::new(data_value);
            if sort_on_code == 1 {
                On::CellColor {
                    differential_format,
                }
            } else {
                On::FontColor {
                    differential_format,
                }
            }
        },
        3 => {
            if char_count != 0 || body.len() != SORT_CONDITION_FIXED_LEN {
                return Err(invalid(
                    "icon SortCond12 has a custom list or trailing data",
                ));
            }
            On::Icon {
                set: IconSet::from_code(data_value)?,
                icon: Icon::from_code(reserved_data as i32)?,
            }
        },
        _ => return Err(invalid("SortCond12 sortOn is outside 0 through 3")),
    };
    match axis {
        Axis::Rows => Key::col(range, flags & 0x0001 != 0, sort_on),
        Axis::Cols => Key::row(range, flags & 0x0001 != 0, sort_on),
    }
}

fn parse_custom_list(body: &[u8], char_count: usize) -> XlsResult<Option<String>> {
    if char_count == 0 {
        if body.len() != SORT_CONDITION_FIXED_LEN {
            return Err(invalid("SortCond12 has trailing bytes with zero cchSt"));
        }
        return Ok(None);
    }
    let flags = *body
        .get(SORT_CONDITION_FIXED_LEN)
        .ok_or(XlsError::InvalidLength {
            expected: SORT_CONDITION_FIXED_LEN + 1,
            found: body.len(),
        })?;
    if flags & !0x01 != 0 {
        return Err(invalid("XLUnicodeStringNoCch has unsupported flag bits"));
    }
    let wide = flags & 0x01 != 0;
    let encoded_len = char_count
        .checked_mul(if wide { 2 } else { 1 })
        .ok_or_else(|| invalid("SortCond12 string byte length overflow"))?;
    let expected = SORT_CONDITION_FIXED_LEN + 1 + encoded_len;
    if body.len() != expected {
        return Err(XlsError::InvalidLength {
            expected,
            found: body.len(),
        });
    }
    let encoded = &body[SORT_CONDITION_FIXED_LEN + 1..];
    let reserve = char_count
        .checked_mul(if wide { 3 } else { 2 })
        .ok_or_else(|| invalid("SortCond12 decoded string length overflows usize"))?;
    let mut value = String::new();
    value
        .try_reserve_exact(reserve)
        .map_err(|_| allocation("reserving decoded SortCond12 string storage"))?;
    if wide {
        let units = encoded
            .chunks_exact(2)
            .map(|bytes| u16::from_le_bytes([bytes[0], bytes[1]]));
        for character in char::decode_utf16(units) {
            value.push(character.map_err(|_| {
                XlsError::Encoding("invalid UTF-16 in SortCond12 custom list".into())
            })?);
        }
    } else {
        for byte in encoded {
            value.push(char::from(*byte));
        }
    }
    Ok(Some(value))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn split_records(bytes: &[u8]) -> Vec<(u16, Vec<u8>)> {
        let mut records = Vec::new();
        let mut offset = 0;
        while offset < bytes.len() {
            let record_type = u16::from_le_bytes([bytes[offset], bytes[offset + 1]]);
            let len = u16::from_le_bytes([bytes[offset + 2], bytes[offset + 3]]) as usize;
            offset += 4;
            records.push((record_type, bytes[offset..offset + len].to_vec()));
            offset += len;
        }
        records
    }

    #[test]
    fn round_trips_every_sort_on_kind_and_unicode() {
        let range = Range::new(2..=MAX_ROW_INDEX, 1..=MAX_COLUMN_INDEX).unwrap();
        let mut value = Config::new(range, Parent::Table { id: 41 });
        value.put_axis(Axis::Cols).unwrap();
        value.set_case(true);
        value.set_method(Method::Alternate);
        value
            .add(
                Key::row(
                    Range::new(2..=2, 1..=20).unwrap(),
                    true,
                    On::Values {
                        custom_list: Some("High,中,Low".into()),
                    },
                )
                .unwrap(),
            )
            .unwrap();
        value
            .add(
                Key::row(
                    Range::new(3..=3, 1..=20).unwrap(),
                    false,
                    On::CellColor {
                        differential_format: Dxf::new(7),
                    },
                )
                .unwrap(),
            )
            .unwrap();
        value
            .add(
                Key::row(
                    Range::new(4..=4, 1..=20).unwrap(),
                    true,
                    On::FontColor {
                        differential_format: Dxf::new(11),
                    },
                )
                .unwrap(),
            )
            .unwrap();
        value
            .add(
                Key::row(
                    Range::new(5..=5, 1..=20).unwrap(),
                    false,
                    On::Icon {
                        set: IconSet::FiveQuarters,
                        icon: Icon::Fifth,
                    },
                )
                .unwrap(),
            )
            .unwrap();

        let mut bytes = Vec::new();
        value.write_biff_records(&mut bytes).unwrap();
        let records = split_records(&bytes);
        assert_eq!(records.len(), 5);
        assert_eq!(records[0].0, SORT_DATA_RECORD_TYPE);
        assert_eq!(&records[0].1[18..22], &MAX_ROW_INDEX.to_le_bytes());
        assert_eq!(&records[0].1[26..30], &MAX_COLUMN_INDEX.to_le_bytes());
        assert!(
            records[1..]
                .iter()
                .all(|record| record.0 == CONTINUE_FRT12_RECORD_TYPE)
        );
        let continuations = records[1..]
            .iter()
            .map(|record| record.1.as_slice())
            .collect::<Vec<_>>();
        let parsed = parse_sort_data(&records[0].1, &continuations).unwrap();
        assert_eq!(parsed, value);
    }

    #[test]
    fn exact_domains_use_narrow_checked_storage() {
        assert_eq!(std::mem::size_of::<Rw12>(), 4);
        assert_eq!(std::mem::size_of::<Col12>(), 2);
        assert_eq!(std::mem::size_of::<Row>(), 4);
        assert_eq!(std::mem::size_of::<Col>(), 2);
        assert_eq!(std::mem::size_of::<Range>(), 12);

        let range = Range::new(0..=MAX_ROW_INDEX, 0..=MAX_COLUMN_INDEX).unwrap();
        assert_eq!(range.first_row().index(), 0);
        assert_eq!(range.last_row().index(), MAX_ROW_INDEX);
        assert_eq!(range.first_col().index(), 0);
        assert_eq!(u32::from(range.last_col().index()), MAX_COLUMN_INDEX);
    }

    #[test]
    fn signed_rfx_fields_have_explicit_four_byte_wire_round_trip() {
        let range = Range::new(0..=MAX_ROW_INDEX, 0..=MAX_COLUMN_INDEX).unwrap();
        let mut encoded = Vec::new();
        range.write_to(&mut encoded);

        assert_eq!(&encoded[0..4], &0i32.to_le_bytes());
        assert_eq!(
            &encoded[4..8],
            &i32::try_from(MAX_ROW_INDEX).unwrap().to_le_bytes()
        );
        assert_eq!(&encoded[8..12], &0i32.to_le_bytes());
        assert_eq!(
            &encoded[12..16],
            &i32::try_from(MAX_COLUMN_INDEX).unwrap().to_le_bytes()
        );
        assert_eq!(Range::parse(&encoded, 0).unwrap(), range);

        let max_range = Range::new(
            MAX_ROW_INDEX..=MAX_ROW_INDEX,
            MAX_COLUMN_INDEX..=MAX_COLUMN_INDEX,
        )
        .unwrap();
        let mut max_base = Config::new(max_range, Parent::Sheet);
        max_base
            .add(
                Key::col(
                    Range::new(
                        MAX_ROW_INDEX..=MAX_ROW_INDEX,
                        MAX_COLUMN_INDEX..=MAX_COLUMN_INDEX,
                    )
                    .unwrap(),
                    false,
                    On::Values { custom_list: None },
                )
                .unwrap(),
            )
            .unwrap();
        let mut bytes = Vec::new();
        max_base.write_biff_records(&mut bytes).unwrap();
        let records = split_records(&bytes);
        assert_eq!(
            &records[0].1[14..30],
            &[
                i32::try_from(MAX_ROW_INDEX).unwrap().to_le_bytes(),
                i32::try_from(MAX_ROW_INDEX).unwrap().to_le_bytes(),
                i32::try_from(MAX_COLUMN_INDEX).unwrap().to_le_bytes(),
                i32::try_from(MAX_COLUMN_INDEX).unwrap().to_le_bytes(),
            ]
            .concat()
        );
        assert_eq!(
            parse_sort_data(&records[0].1, &[&records[1].1]).unwrap(),
            max_base
        );
    }

    #[test]
    fn rejects_negative_and_out_of_domain_signed_rfx_values_without_unwind() {
        let value = Config::new(Range::new(0..=0, 0..=0).unwrap(), Parent::Sheet);
        let mut bytes = Vec::new();
        value.write_biff_records(&mut bytes).unwrap();
        let records = split_records(&bytes);
        let base = &records[0].1;

        for (offset, value) in [
            (14, -1i32),
            (18, -1i32),
            (22, -1i32),
            (26, -1i32),
            (14, i32::MIN),
            (18, i32::MIN),
            (22, i32::MIN),
            (26, i32::MIN),
            (14, i32::try_from(MAX_ROW_INDEX).unwrap() + 1),
            (18, i32::try_from(MAX_ROW_INDEX).unwrap() + 1),
            (22, i32::try_from(MAX_COLUMN_INDEX).unwrap() + 1),
            (26, i32::try_from(MAX_COLUMN_INDEX).unwrap() + 1),
        ] {
            let mut malformed = base.clone();
            malformed[offset..offset + 4].copy_from_slice(&value.to_le_bytes());
            let parsed = std::panic::catch_unwind(|| parse_sort_data(&malformed, &[]));
            assert!(
                matches!(parsed, Ok(Err(_))),
                "offset {offset}, value {value}"
            );
        }
    }

    #[test]
    fn validates_frt_reserved_bits_and_conditional_ref8_without_rejecting_ignored_fields() {
        let value = Config::new(Range::new(0..=0, 0..=0).unwrap(), Parent::Sheet);
        let mut bytes = Vec::new();
        value.write_biff_records(&mut bytes).unwrap();
        let records = split_records(&bytes);

        let mut ignored = records[0].1.clone();
        ignored[12..14].copy_from_slice(&0xff80u16.to_le_bytes());
        ignored[34..38].copy_from_slice(&u32::MAX.to_le_bytes());
        assert_eq!(parse_sort_data(&ignored, &[]).unwrap(), value);

        for (offset, flags) in [(2, 0x0001u16), (2, 0x0002), (2, 0x0004), (2, 0x8000)] {
            let mut malformed = records[0].1.clone();
            malformed[offset..offset + 2].copy_from_slice(&flags.to_le_bytes());
            let parsed = std::panic::catch_unwind(|| parse_sort_data(&malformed, &[]));
            assert!(matches!(parsed, Ok(Err(_))), "SortData flags {flags:#06x}");
        }

        let mut reserved_bytes = records[0].1.clone();
        reserved_bytes[4] = 1;
        assert!(parse_sort_data(&reserved_bytes, &[]).is_err());

        let range = Range::new(0..=5, 0..=0).unwrap();
        let mut with_key = Config::new(range, Parent::AutoFilter);
        with_key
            .add(Key::col(range, false, On::Values { custom_list: None }).unwrap())
            .unwrap();
        let mut key_bytes = Vec::new();
        with_key.write_biff_records(&mut key_bytes).unwrap();
        let key_records = split_records(&key_bytes);
        let continuation = &key_records[1].1;

        let mut valid_ref = continuation.clone();
        valid_ref[2..4].copy_from_slice(&FRT_REF_FLAG.to_le_bytes());
        valid_ref[4..12].copy_from_slice(&[0, 0, 5, 0, 0, 0, 0, 0]);
        assert_eq!(
            parse_sort_data(&key_records[0].1, &[&valid_ref]).unwrap(),
            with_key
        );

        for flags in [0x0002u16, 0x0004, 0x8000] {
            let mut malformed = continuation.clone();
            malformed[2..4].copy_from_slice(&flags.to_le_bytes());
            let parsed = std::panic::catch_unwind(|| {
                parse_sort_data(&key_records[0].1, &[malformed.as_slice()])
            });
            assert!(
                matches!(parsed, Ok(Err(_))),
                "ContinueFrt12 flags {flags:#06x}"
            );
        }

        let mut missing_ref = continuation.clone();
        missing_ref[4] = 1;
        assert!(parse_sort_data(&key_records[0].1, &[&missing_ref]).is_err());

        let mut reversed_ref = valid_ref.clone();
        reversed_ref[4..6].copy_from_slice(&6u16.to_le_bytes());
        assert!(parse_sort_data(&key_records[0].1, &[&reversed_ref]).is_err());

        let mut wide_ref = valid_ref;
        wide_ref[10..12].copy_from_slice(&0x0100u16.to_le_bytes());
        assert!(parse_sort_data(&key_records[0].1, &[&wide_ref]).is_err());
    }

    #[test]
    fn rejects_truncated_sort_data_and_continuations_without_unwind() {
        let range = Range::new(0..=0, 0..=0).unwrap();
        let mut value = Config::new(range, Parent::Sheet);
        value
            .add(Key::col(range, false, On::Values { custom_list: None }).unwrap())
            .unwrap();
        let mut bytes = Vec::new();
        value.write_biff_records(&mut bytes).unwrap();
        let records = split_records(&bytes);
        let base = &records[0].1;
        let continuation = &records[1].1;

        for length in 0..base.len() {
            let parsed = std::panic::catch_unwind(|| parse_sort_data(&base[..length], &[]));
            assert!(matches!(parsed, Ok(Err(_))), "base length {length}");
        }
        for length in 0..continuation.len() {
            let parsed =
                std::panic::catch_unwind(|| parse_sort_data(base, &[&continuation[..length]]));
            assert!(matches!(parsed, Ok(Err(_))), "continuation length {length}");
        }
    }

    #[test]
    fn rejects_invalid_ranges_and_condition_count_mismatch_without_unwind() {
        let rejected = std::panic::catch_unwind(|| {
            let first = 8;
            let last = 7;
            (
                Range::new(first..=last, 0..=0),
                Row::new(MAX_ROW_INDEX + 1),
                Col::new(MAX_COLUMN_INDEX + 1),
                Col::new(u32::MAX),
            )
        });
        let (reversed, row_overflow, col_overflow, wide_overflow) = rejected.unwrap();
        assert!(reversed.is_err());
        assert!(row_overflow.is_err());
        assert!(col_overflow.is_err());
        assert!(wide_overflow.is_err());

        let value = Config::new(Range::new(0..=0, 0..=0).unwrap(), Parent::Sheet);
        let mut bytes = Vec::new();
        value.write_biff_records(&mut bytes).unwrap();
        let mut records = split_records(&bytes);
        records[0].1[30..34].copy_from_slice(&1u32.to_le_bytes());
        assert!(parse_sort_data(&records[0].1, &[]).is_err());

        records[0].1[30..34].copy_from_slice(&0u32.to_le_bytes());
        records[0].1[14..18].copy_from_slice(&(-1i32).to_le_bytes());
        let parsed = std::panic::catch_unwind(|| parse_sort_data(&records[0].1, &[]));
        assert!(matches!(parsed, Ok(Err(_))));
    }

    #[test]
    fn key_edits_reject_ambiguity_and_are_failure_atomic() {
        let range = Range::new(0..=5, 0..=2).unwrap();
        let mut value = Config::new(range, Parent::Sheet);
        let before = value.clone();
        let row_key = Key::row(
            Range::new(0..=0, 0..=2).unwrap(),
            false,
            On::Values { custom_list: None },
        )
        .unwrap();
        assert!(value.add(row_key).is_err());
        assert_eq!(value, before);

        let outside = Key::col(
            Range::new(0..=5, 3..=3).unwrap(),
            false,
            On::Values { custom_list: None },
        )
        .unwrap();
        assert!(value.add(outside).is_err());
        assert_eq!(value, before);

        let ambiguous = Range::new(0..=5, 0..=1).unwrap();
        assert!(Key::col(ambiguous, false, On::Values { custom_list: None }).is_err());

        value
            .add(
                Key::col(
                    Range::new(0..=5, 0..=0).unwrap(),
                    false,
                    On::Values { custom_list: None },
                )
                .unwrap(),
            )
            .unwrap();
        let before_axis = value.clone();
        assert!(value.put_axis(Axis::Cols).is_err());
        assert_eq!(value, before_axis);
    }

    #[test]
    fn rejects_malformed_continuation_fields_and_unicode() {
        let range = Range::new(0..=5, 0..=0).unwrap();
        let mut value = Config::new(range, Parent::AutoFilter);
        value
            .add(
                Key::col(
                    range,
                    false,
                    On::Values {
                        custom_list: Some("中".into()),
                    },
                )
                .unwrap(),
            )
            .unwrap();
        let mut bytes = Vec::new();
        value.write_biff_records(&mut bytes).unwrap();
        let records = split_records(&bytes);
        let mut malformed = records[1].1.clone();
        malformed[FRT_HEADER_LEN + SORT_CONDITION_FIXED_LEN] = 1;
        malformed[FRT_HEADER_LEN + SORT_CONDITION_FIXED_LEN + 1..]
            .copy_from_slice(&0xd800u16.to_le_bytes());
        malformed[FRT_HEADER_LEN + 26..FRT_HEADER_LEN + 30].copy_from_slice(&1u32.to_le_bytes());
        assert!(parse_sort_data(&records[0].1, &[&malformed]).is_err());

        let mut bad_icon = records[1].1.clone();
        bad_icon[FRT_HEADER_LEN..FRT_HEADER_LEN + 2].copy_from_slice(&6u16.to_le_bytes());
        bad_icon[FRT_HEADER_LEN + 18..FRT_HEADER_LEN + 22].copy_from_slice(&99u32.to_le_bytes());
        bad_icon.truncate(FRT_HEADER_LEN + SORT_CONDITION_FIXED_LEN);
        assert!(parse_sort_data(&records[0].1, &[&bad_icon]).is_err());

        let mut reserved_bit = records[1].1.clone();
        let flags = read_u16(&reserved_bit, FRT_HEADER_LEN).unwrap() | 0x8000;
        reserved_bit[FRT_HEADER_LEN..FRT_HEADER_LEN + 2].copy_from_slice(&flags.to_le_bytes());
        assert!(parse_sort_data(&records[0].1, &[&reserved_bit]).is_err());

        let mut ambiguous = records[1].1.clone();
        ambiguous[FRT_HEADER_LEN + 14..FRT_HEADER_LEN + 18].copy_from_slice(&1u32.to_le_bytes());
        let parsed =
            std::panic::catch_unwind(|| parse_sort_data(&records[0].1, &[ambiguous.as_slice()]));
        assert!(matches!(parsed, Ok(Err(_))));
    }

    #[test]
    fn rejects_oversized_or_empty_custom_lists_before_retention() {
        let range = Range::new(0..=0, 0..=0).unwrap();
        let oversized = std::panic::catch_unwind(|| {
            Key::col(
                range,
                false,
                On::Values {
                    custom_list: Some("a".repeat(MAX_CONTINUE_RGB_LEN)),
                },
            )
        });
        assert!(matches!(oversized, Ok(Err(_))));
        assert!(
            Key::col(
                range,
                false,
                On::Values {
                    custom_list: Some(String::new()),
                },
            )
            .is_err()
        );
    }
}
