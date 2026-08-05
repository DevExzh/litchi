//! BIFF8 chart future-record (FRT) records: `ChartFrtInfo`, `CatLab`,
//! `StartBlock`, and `EndBlock` of the chart sheet substream (MS-XLS 2.1).
//!
//! These records all carry a legacy 4-byte `FrtHeaderOld` (MS-XLS 2.5.136)
//! and describe how post-BIFF8 chart features are versioned and scoped:
//!
//! - **ChartFrtInfo** (0x0850): application versions that created/last saved
//!   the file and the `CFrtId` ranges of Future Record Type identifiers used
//!   in the chart (MS-XLS 2.4.49).
//! - **CatLab** (0x0856): attributes of an axis label (MS-XLS 2.4.38).
//! - **StartBlock** (0x0852) / **EndBlock** (0x0853): bracket a collection of
//!   chart-specific future records so applications that do not support a
//!   feature can preserve it (MS-XLS 2.4.266 / 2.4.100).
//!
//! Everything in this module is INERT: fields are stored verbatim and no
//! chart is rendered.
//!
//! # References
//!
//! - MS-XLS 2.4.38 (CatLab), 2.4.49 (ChartFrtInfo), 2.4.100 (EndBlock),
//!   2.4.266 (StartBlock), 2.5.37 (CFrtId), 2.5.134 (FrtFlags),
//!   2.5.136 (FrtHeaderOld)

use crate::{Error, Result};

/// Record type of the `ChartFrtInfo` record (MS-XLS 2.4.49).
pub(crate) const CHART_FRT_INFO_RECORD_TYPE: u16 = 0x0850;
/// Record type of the `StartBlock` record (MS-XLS 2.4.266).
pub(crate) const START_BLOCK_RECORD_TYPE: u16 = 0x0852;
/// Record type of the `EndBlock` record (MS-XLS 2.4.100).
pub(crate) const END_BLOCK_RECORD_TYPE: u16 = 0x0853;
/// Record type of the `CatLab` record (MS-XLS 2.4.38).
pub(crate) const CAT_LAB_RECORD_TYPE: u16 = 0x0856;

/// Size in bytes of an `FrtHeaderOld` (MS-XLS 2.5.136): `rt` + `grbitFrt`.
const FRT_HEADER_OLD_LEN: usize = 4;
/// `fFrtRef` and `fFrtAlert` are not valid for these chart future records.
const FRT_FLAGS_FORBIDDEN: u16 = 0x0003;
/// Size in bytes of a `CFrtId` structure (MS-XLS 2.5.37).
const C_FRT_ID_LEN: usize = 4;
/// Size in bytes of the fixed part of a `ChartFrtInfo` record:
/// `FrtHeaderOld` (4) + `verOriginator` (1) + `verWriter` (1) + `cCFRTID` (2).
const CHART_FRT_INFO_BASE_LEN: usize = 8;
/// Size in bytes of a `CatLab` record payload.
const CAT_LAB_LEN: usize = 12;
/// Size in bytes of a `StartBlock` record payload.
const START_BLOCK_LEN: usize = 12;
/// Size in bytes of an `EndBlock` record payload.
const END_BLOCK_LEN: usize = 12;

/// Highest legal `wOffset` of a `CatLab` record (1000% of the default
/// distance, MS-XLS 2.4.38).
const CAT_LAB_MAX_OFFSET: u16 = 1000;
/// Flag bit of the `CatLab` bitfield: `cAutoCatLabelReal` (MS-XLS 2.4.38).
const CAT_LAB_AUTO_LABEL: u16 = 1;

/// Read a little-endian `u16` from a fixed offset.
fn read_u16(data: &[u8], offset: usize) -> Result<u16> {
    let end = offset.checked_add(2).ok_or(Error::InvalidLength {
        expected: 2,
        found: data.len(),
    })?;
    let bytes = data.get(offset..end).ok_or(Error::InvalidLength {
        expected: end,
        found: data.len(),
    })?;
    let [first, second] = bytes.try_into().map_err(|_| Error::InvalidLength {
        expected: 2,
        found: bytes.len(),
    })?;
    Ok(u16::from_le_bytes([first, second]))
}

/// Read and validate the fixed `FrtHeaderOld` prefix.
fn read_frt_header_old(data: &[u8], record_type: u16) -> Result<u16> {
    if data.len() < FRT_HEADER_OLD_LEN {
        return Err(Error::InvalidLength {
            expected: FRT_HEADER_OLD_LEN,
            found: data.len(),
        });
    }
    if read_u16(data, 0)? != record_type {
        return Err(Error::InvalidRecord {
            record_type,
            message: "FrtHeaderOld.rt mismatch".to_string(),
        });
    }
    let flags = read_u16(data, 2)?;
    if flags & FRT_FLAGS_FORBIDDEN != 0 {
        return Err(Error::InvalidRecord {
            record_type,
            message: format!("FrtHeaderOld.grbitFrt {flags:#06X} sets fFrtRef or fFrtAlert"),
        });
    }
    Ok(flags)
}

/// Validate the `FrtHeaderOld` and the fixed payload length.
fn validate_frt_header_old(data: &[u8], record_type: u16, expected_len: usize) -> Result<u16> {
    if data.len() != expected_len {
        return Err(Error::InvalidLength {
            expected: expected_len,
            found: data.len(),
        });
    }
    read_frt_header_old(data, record_type)
}

/// Application version that created or last saved the file, as declared by a
/// `ChartFrtInfo` record (MS-XLS 2.4.49). Variant names carry the numeric
/// version value (0x9, 0xA, 0xC, 0xE).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Version {
    /// Application version 0x9; the chart uses one `CFrtId` range.
    Version9,
    /// Application version 0xA; the chart uses three `CFrtId` ranges.
    Version10,
    /// Application version 0xC; the chart uses four `CFrtId` ranges.
    Version12,
    /// Application version 0xE; the chart uses four `CFrtId` ranges.
    Version14,
}

impl Version {
    /// Decode a raw version byte.
    fn from_byte(byte: u8, record_type: u16) -> Result<Self> {
        match byte {
            0x9 => Ok(Self::Version9),
            0xA => Ok(Self::Version10),
            0xC => Ok(Self::Version12),
            0xE => Ok(Self::Version14),
            _ => Err(Error::InvalidRecord {
                record_type,
                message: format!("unsupported ChartFrtInfo application version {byte:#04X}"),
            }),
        }
    }

    /// Raw version byte.
    fn to_byte(self) -> u8 {
        match self {
            Self::Version9 => 0x9,
            Self::Version10 => 0xA,
            Self::Version12 => 0xC,
            Self::Version14 => 0xE,
        }
    }

    /// Number of `CFrtId` ranges the version mandates in `cCFRTID`.
    fn expected_range_count(self) -> usize {
        match self {
            Self::Version9 => 1,
            Self::Version10 => 3,
            Self::Version12 | Self::Version14 => 4,
        }
    }
}

/// A `CFrtId` structure (MS-XLS 2.5.37): an inclusive range of Future Record
/// Type identifier values used in a chart.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RecordRange {
    /// First Future Record Type in the range (`rtFirst`).
    first: u16,
    /// Last Future Record Type in the range (`rtLast`).
    last: u16,
}

impl RecordRange {
    /// First Future Record Type in the range.
    pub fn first(&self) -> u16 {
        self.first
    }

    /// Last Future Record Type in the range.
    pub fn last(&self) -> u16 {
        self.last
    }

    /// Construct a range, rejecting an inverted interval.
    pub(crate) fn new(first: u16, last: u16) -> Result<Self> {
        if first > last {
            return Err(Error::InvalidRecord {
                record_type: CHART_FRT_INFO_RECORD_TYPE,
                message: format!("CFrtId range {first:#06X}..={last:#06X} is inverted"),
            });
        }
        Ok(Self { first, last })
    }
}

/// Typed `ChartFrtInfo` record content (MS-XLS 2.4.49).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Info {
    /// Raw `FrtHeaderOld.grbitFrt`; forbidden reference/alert bits are zero
    /// and the remaining bits are retained for lossless rewriting.
    frt_flags: u16,
    /// Application version that originally created the file (`verOriginator`).
    originator: Version,
    /// Application version that last saved the file (`verWriter`).
    writer: Version,
    /// Ranges of Future Record Type identifiers used in the chart (`rgCFRTID`).
    ranges: Vec<RecordRange>,
}

impl Info {
    /// Application version that originally created the file.
    pub fn originator(&self) -> Version {
        self.originator
    }

    /// Raw `FrtHeaderOld.grbitFrt` bits retained from the source record.
    pub fn frt_flags(&self) -> u16 {
        self.frt_flags
    }

    /// Application version that last saved the file.
    pub fn writer(&self) -> Version {
        self.writer
    }

    /// Ranges of Future Record Type identifiers used in the chart.
    pub fn ranges(&self) -> &[RecordRange] {
        &self.ranges
    }

    /// Parse a `ChartFrtInfo` record payload.
    pub fn parse(data: &[u8]) -> Result<Self> {
        let invalid = |message: String| Error::InvalidRecord {
            record_type: CHART_FRT_INFO_RECORD_TYPE,
            message,
        };
        if data.len() < CHART_FRT_INFO_BASE_LEN {
            return Err(Error::InvalidLength {
                expected: CHART_FRT_INFO_BASE_LEN,
                found: data.len(),
            });
        }
        let frt_flags = read_frt_header_old(data, CHART_FRT_INFO_RECORD_TYPE)?;
        let originator = Version::from_byte(data[4], CHART_FRT_INFO_RECORD_TYPE)?;
        let writer = Version::from_byte(data[5], CHART_FRT_INFO_RECORD_TYPE)?;
        let count = usize::from(read_u16(data, 6)?);
        if count != writer.expected_range_count() {
            return Err(invalid(format!(
                "ChartFrtInfo cCFRTID {count} disagrees with the writer version"
            )));
        }
        let expected_len = CHART_FRT_INFO_BASE_LEN + count * C_FRT_ID_LEN;
        if data.len() != expected_len {
            return Err(Error::InvalidLength {
                expected: expected_len,
                found: data.len(),
            });
        }
        let mut ranges = Vec::with_capacity(count);
        for index in 0..count {
            let offset = CHART_FRT_INFO_BASE_LEN + index * C_FRT_ID_LEN;
            ranges.push(RecordRange::new(
                read_u16(data, offset)?,
                read_u16(data, offset + 2)?,
            )?);
        }
        Ok(Self {
            frt_flags,
            originator,
            writer,
            ranges,
        })
    }

    /// Serialize back to a complete `ChartFrtInfo` record payload.
    pub fn to_payload(&self) -> Vec<u8> {
        let mut payload =
            Vec::with_capacity(CHART_FRT_INFO_BASE_LEN + self.ranges.len() * C_FRT_ID_LEN);
        // FrtHeaderOld: rt and the preserved grbitFrt bits.
        payload.extend_from_slice(&CHART_FRT_INFO_RECORD_TYPE.to_le_bytes());
        payload.extend_from_slice(&self.frt_flags.to_le_bytes());
        payload.push(self.originator.to_byte());
        payload.push(self.writer.to_byte());
        payload.extend_from_slice(&(self.ranges.len() as u16).to_le_bytes());
        for range in &self.ranges {
            payload.extend_from_slice(&range.first.to_le_bytes());
            payload.extend_from_slice(&range.last.to_le_bytes());
        }
        payload
    }
}

/// Alignment of an axis label, as declared by the `at` field of a `CatLab`
/// record (MS-XLS 2.4.38).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Alignment {
    /// Top-aligned when the axis text is rotated; left-aligned under a
    /// left-to-right reading order, right-aligned otherwise.
    TopLeft,
    /// Center-aligned.
    Center,
    /// Bottom-aligned when the axis text is rotated; right-aligned under a
    /// left-to-right reading order, left-aligned otherwise.
    BottomRight,
}

impl Alignment {
    /// Decode the raw `at` value.
    fn from_u16(value: u16) -> Result<Self> {
        match value {
            0x0001 => Ok(Self::TopLeft),
            0x0002 => Ok(Self::Center),
            0x0003 => Ok(Self::BottomRight),
            _ => Err(Error::InvalidRecord {
                record_type: CAT_LAB_RECORD_TYPE,
                message: format!("unsupported CatLab alignment {value:#06X}"),
            }),
        }
    }

    /// Raw `at` value.
    fn to_u16(self) -> u16 {
        match self {
            Self::TopLeft => 0x0001,
            Self::Center => 0x0002,
            Self::BottomRight => 0x0003,
        }
    }
}

/// Typed `CatLab` record content (MS-XLS 2.4.38): attributes of an axis label.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct CatLab {
    /// Raw `FrtHeaderOld.grbitFrt`; forbidden reference/alert bits are zero.
    frt_flags: u16,
    /// Distance between the axis and the axis label, as a percentage of the
    /// default distance (`wOffset`, 0..=1000).
    offset: u16,
    /// Alignment of the axis label (`at`).
    alignment: Alignment,
    /// Whether the number of categories between axis labels is automatically
    /// calculated (`cAutoCatLabelReal`).
    /// The complete `cAutoCatLabelReal`/unused bitfield, retained verbatim.
    category_flags: u16,
    /// Reserved bytes following the category-label bitfield.
    reserved: [u8; 2],
}

impl CatLab {
    /// Distance between the axis and the axis label (percent of the default).
    pub fn offset(&self) -> u16 {
        self.offset
    }

    /// Raw `FrtHeaderOld.grbitFrt` bits retained from the source record.
    pub fn frt_flags(&self) -> u16 {
        self.frt_flags
    }

    /// Alignment of the axis label.
    pub fn alignment(&self) -> Alignment {
        self.alignment
    }

    /// Whether the number of categories between axis labels is automatic.
    pub fn auto_category_label(&self) -> bool {
        self.category_flags & CAT_LAB_AUTO_LABEL != 0
    }

    /// Raw `CatLab` category-label bitfield, including ignored bits.
    pub fn category_flags(&self) -> u16 {
        self.category_flags
    }

    /// Reserved bytes following the category-label bitfield.
    pub fn reserved(&self) -> [u8; 2] {
        self.reserved
    }

    /// Parse a `CatLab` record payload.
    pub fn parse(data: &[u8]) -> Result<Self> {
        let frt_flags = validate_frt_header_old(data, CAT_LAB_RECORD_TYPE, CAT_LAB_LEN)?;
        let offset = read_u16(data, 4)?;
        if offset > CAT_LAB_MAX_OFFSET {
            return Err(Error::InvalidRecord {
                record_type: CAT_LAB_RECORD_TYPE,
                message: format!("CatLab wOffset {offset} exceeds {CAT_LAB_MAX_OFFSET}"),
            });
        }
        let alignment = Alignment::from_u16(read_u16(data, 6)?)?;
        let flags = read_u16(data, 8)?;
        let reserved = [data[10], data[11]];
        Ok(Self {
            frt_flags,
            offset,
            alignment,
            category_flags: flags,
            reserved,
        })
    }

    /// Serialize back to a complete `CatLab` record payload.
    pub fn to_payload(&self) -> Vec<u8> {
        let mut payload = Vec::with_capacity(CAT_LAB_LEN);
        // FrtHeaderOld: rt and the preserved grbitFrt bits.
        payload.extend_from_slice(&CAT_LAB_RECORD_TYPE.to_le_bytes());
        payload.extend_from_slice(&self.frt_flags.to_le_bytes());
        payload.extend_from_slice(&self.offset.to_le_bytes());
        payload.extend_from_slice(&self.alignment.to_u16().to_le_bytes());
        payload.extend_from_slice(&self.category_flags.to_le_bytes());
        payload.extend_from_slice(&self.reserved);
        payload
    }
}

/// Type of chart object encompassed by a `StartBlock`/`EndBlock` pair
/// (`iObjectKind`, MS-XLS 2.4.266 / 2.4.100).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BlockKind {
    /// Axis group.
    AxisGroup,
    /// Attached label record.
    AttachedLabel,
    /// Axis.
    Axis,
    /// Chart group.
    ChartGroup,
    /// Dat record.
    Dat,
    /// Frame.
    Frame,
    /// Legend.
    Legend,
    /// LegendException record.
    LegendException,
    /// Series.
    Series,
    /// Sheet.
    Sheet,
    /// DataFormat record.
    DataFormat,
    /// DropBar record.
    DropBar,
}

impl BlockKind {
    /// Decode the raw `iObjectKind` value.
    fn from_u16(value: u16, record_type: u16) -> Result<Self> {
        match value {
            0x0000 => Ok(Self::AxisGroup),
            0x0002 => Ok(Self::AttachedLabel),
            0x0004 => Ok(Self::Axis),
            0x0005 => Ok(Self::ChartGroup),
            0x0006 => Ok(Self::Dat),
            0x0007 => Ok(Self::Frame),
            0x0009 => Ok(Self::Legend),
            0x000A => Ok(Self::LegendException),
            0x000C => Ok(Self::Series),
            0x000D => Ok(Self::Sheet),
            0x000E => Ok(Self::DataFormat),
            0x000F => Ok(Self::DropBar),
            _ => Err(Error::InvalidRecord {
                record_type,
                message: format!("unsupported block object kind {value:#06X}"),
            }),
        }
    }

    /// Raw `iObjectKind` value.
    fn to_u16(self) -> u16 {
        match self {
            Self::AxisGroup => 0x0000,
            Self::AttachedLabel => 0x0002,
            Self::Axis => 0x0004,
            Self::ChartGroup => 0x0005,
            Self::Dat => 0x0006,
            Self::Frame => 0x0007,
            Self::Legend => 0x0009,
            Self::LegendException => 0x000A,
            Self::Series => 0x000C,
            Self::Sheet => 0x000D,
            Self::DataFormat => 0x000E,
            Self::DropBar => 0x000F,
        }
    }
}

/// Typed `StartBlock` record content (MS-XLS 2.4.266): the beginning of a
/// collection of chart-specific future records preserved for applications
/// that do not support the feature.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct StartBlock {
    /// Raw `FrtHeaderOld.grbitFrt`; forbidden reference/alert bits are zero.
    frt_flags: u16,
    /// Type of object encompassed by the block (`iObjectKind`).
    object_kind: BlockKind,
    /// Context of the object (`iObjectContext`); its meaning depends on
    /// `object_kind` and is preserved verbatim.
    object_context: u16,
    /// First instance qualifier (`iObjectInstance1`), preserved verbatim.
    object_instance1: u16,
    /// Second instance qualifier (`iObjectInstance2`), preserved verbatim.
    object_instance2: u16,
}

impl StartBlock {
    /// Type of object encompassed by the block.
    pub fn object_kind(&self) -> BlockKind {
        self.object_kind
    }

    /// Context of the object.
    pub fn object_context(&self) -> u16 {
        self.object_context
    }

    /// First instance qualifier.
    pub fn object_instance1(&self) -> u16 {
        self.object_instance1
    }

    /// Second instance qualifier.
    pub fn object_instance2(&self) -> u16 {
        self.object_instance2
    }

    /// Parse a `StartBlock` record payload.
    pub fn parse(data: &[u8]) -> Result<Self> {
        let frt_flags = validate_frt_header_old(data, START_BLOCK_RECORD_TYPE, START_BLOCK_LEN)?;
        Ok(Self {
            frt_flags,
            object_kind: BlockKind::from_u16(read_u16(data, 4)?, START_BLOCK_RECORD_TYPE)?,
            object_context: read_u16(data, 6)?,
            object_instance1: read_u16(data, 8)?,
            object_instance2: read_u16(data, 10)?,
        })
    }

    /// Serialize back to a complete `StartBlock` record payload.
    pub fn to_payload(&self) -> Vec<u8> {
        let mut payload = Vec::with_capacity(START_BLOCK_LEN);
        // FrtHeaderOld: rt and the preserved grbitFrt bits.
        payload.extend_from_slice(&START_BLOCK_RECORD_TYPE.to_le_bytes());
        payload.extend_from_slice(&self.frt_flags.to_le_bytes());
        payload.extend_from_slice(&self.object_kind.to_u16().to_le_bytes());
        payload.extend_from_slice(&self.object_context.to_le_bytes());
        payload.extend_from_slice(&self.object_instance1.to_le_bytes());
        payload.extend_from_slice(&self.object_instance2.to_le_bytes());
        payload
    }

    /// Raw `FrtHeaderOld.grbitFrt` bits retained from the source record.
    pub fn frt_flags(&self) -> u16 {
        self.frt_flags
    }
}

/// Typed `EndBlock` record content (MS-XLS 2.4.100): the end of the collection
/// opened by the associated `StartBlock` record.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct EndBlock {
    /// Raw `FrtHeaderOld.grbitFrt`; forbidden reference/alert bits are zero.
    frt_flags: u16,
    /// Type of object encompassed by the block (`iObjectKind`); MUST equal the
    /// `iObjectKind` of the associated `StartBlock` record.
    object_kind: BlockKind,
    /// The three undefined fields, retained verbatim.
    unused: [u16; 3],
}

impl EndBlock {
    /// Type of object encompassed by the block.
    pub fn object_kind(&self) -> BlockKind {
        self.object_kind
    }

    /// Parse an `EndBlock` record payload.
    pub fn parse(data: &[u8]) -> Result<Self> {
        let frt_flags = validate_frt_header_old(data, END_BLOCK_RECORD_TYPE, END_BLOCK_LEN)?;
        Ok(Self {
            frt_flags,
            object_kind: BlockKind::from_u16(read_u16(data, 4)?, END_BLOCK_RECORD_TYPE)?,
            unused: [read_u16(data, 6)?, read_u16(data, 8)?, read_u16(data, 10)?],
        })
    }

    /// Serialize back to a complete `EndBlock` record payload.
    pub fn to_payload(&self) -> Vec<u8> {
        let mut payload = Vec::with_capacity(END_BLOCK_LEN);
        // FrtHeaderOld: rt and the preserved grbitFrt bits.
        payload.extend_from_slice(&END_BLOCK_RECORD_TYPE.to_le_bytes());
        payload.extend_from_slice(&self.frt_flags.to_le_bytes());
        payload.extend_from_slice(&self.object_kind.to_u16().to_le_bytes());
        for value in self.unused {
            payload.extend_from_slice(&value.to_le_bytes());
        }
        payload
    }

    /// The three undefined fields retained from the source record.
    pub fn unused(&self) -> [u16; 3] {
        self.unused
    }

    /// Raw `FrtHeaderOld.grbitFrt` bits retained from the source record.
    pub fn frt_flags(&self) -> u16 {
        self.frt_flags
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn frt_header_old(record_type: u16) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&record_type.to_le_bytes());
        data.extend_from_slice(&[0; 2]);
        data
    }

    fn chart_frt_info_payload(
        originator: u8,
        writer: u8,
        count: u16,
        ranges: &[(u16, u16)],
    ) -> Vec<u8> {
        let mut data = frt_header_old(CHART_FRT_INFO_RECORD_TYPE);
        data.push(originator);
        data.push(writer);
        data.extend_from_slice(&count.to_le_bytes());
        for (first, last) in ranges {
            data.extend_from_slice(&first.to_le_bytes());
            data.extend_from_slice(&last.to_le_bytes());
        }
        data
    }

    fn cat_lab_payload(offset: u16, at: u16, flags: u16) -> Vec<u8> {
        let mut data = frt_header_old(CAT_LAB_RECORD_TYPE);
        data.extend_from_slice(&offset.to_le_bytes());
        data.extend_from_slice(&at.to_le_bytes());
        data.extend_from_slice(&flags.to_le_bytes());
        data.extend_from_slice(&[0; 2]);
        data
    }

    fn start_block_payload(kind: u16, context: u16, instance1: u16, instance2: u16) -> Vec<u8> {
        let mut data = frt_header_old(START_BLOCK_RECORD_TYPE);
        data.extend_from_slice(&kind.to_le_bytes());
        data.extend_from_slice(&context.to_le_bytes());
        data.extend_from_slice(&instance1.to_le_bytes());
        data.extend_from_slice(&instance2.to_le_bytes());
        data
    }

    fn end_block_payload(kind: u16) -> Vec<u8> {
        let mut data = frt_header_old(END_BLOCK_RECORD_TYPE);
        data.extend_from_slice(&kind.to_le_bytes());
        data.extend_from_slice(&[0; 6]);
        data
    }

    #[test]
    fn fixed_width_reads_reject_truncation_and_offset_overflow() {
        assert!(read_u16(&[0], 0).is_err());
        assert!(read_u16(&[], usize::MAX).is_err());
    }

    #[test]
    fn parses_chart_frt_info_for_each_writer_version() {
        // Version 0x9: a single CFrtId range.
        let payload = chart_frt_info_payload(0x9, 0x9, 1, &[(0x0850, 0x085A)]);
        let parsed = Info::parse(&payload).unwrap();
        assert_eq!(parsed.originator(), Version::Version9);
        assert_eq!(parsed.writer(), Version::Version9);
        assert_eq!(parsed.ranges().len(), 1);
        assert_eq!(parsed.ranges()[0].first(), 0x0850);
        assert_eq!(parsed.ranges()[0].last(), 0x085A);
        assert_eq!(parsed.to_payload(), payload);

        // Version 0xA: three CFrtId ranges.
        let payload = chart_frt_info_payload(
            0x9,
            0xA,
            3,
            &[(0x0850, 0x085A), (0x0861, 0x0861), (0x086A, 0x086B)],
        );
        let parsed = Info::parse(&payload).unwrap();
        assert_eq!(parsed.writer(), Version::Version10);
        assert_eq!(parsed.ranges().len(), 3);
        assert_eq!(parsed.to_payload(), payload);

        // Version 0xC and 0xE: four CFrtId ranges.
        for writer in [0xC, 0xE] {
            let payload = chart_frt_info_payload(
                writer,
                writer,
                4,
                &[
                    (0x0850, 0x085A),
                    (0x0861, 0x0861),
                    (0x086A, 0x086B),
                    (0x089D, 0x08A6),
                ],
            );
            let parsed = Info::parse(&payload).unwrap();
            assert_eq!(parsed.ranges().len(), 4);
            assert_eq!(parsed.to_payload(), payload);
        }
    }

    #[test]
    fn rejects_malformed_chart_frt_info() {
        let valid = chart_frt_info_payload(0x9, 0x9, 1, &[(0x0850, 0x085A)]);
        // Truncated.
        assert!(Info::parse(&valid[..7]).is_err());
        // Wrong FrtHeaderOld.rt.
        let mut wrong_rt = valid.clone();
        wrong_rt[0..2].copy_from_slice(&0x0851u16.to_le_bytes());
        assert!(Info::parse(&wrong_rt).is_err());
        // FrtFlags.fFrtRef/fFrtAlert are forbidden for ChartFrtInfo.
        for flags in [0x0001u16, 0x0002] {
            let mut bad_flags = valid.clone();
            bad_flags[2..4].copy_from_slice(&flags.to_le_bytes());
            assert!(Info::parse(&bad_flags).is_err());
        }
        // Unsupported originator/writer versions.
        assert!(Info::parse(&chart_frt_info_payload(0xB, 0x9, 1, &[(0x0850, 0x085A)])).is_err());
        assert!(Info::parse(&chart_frt_info_payload(0x9, 0xF, 1, &[(0x0850, 0x085A)])).is_err());
        // cCFRTID disagrees with the writer version.
        assert!(Info::parse(&chart_frt_info_payload(0x9, 0x9, 3, &[(0x0850, 0x085A); 3])).is_err());
        assert!(Info::parse(&chart_frt_info_payload(0x9, 0xA, 1, &[(0x0850, 0x085A)])).is_err());
        // Declared count disagreeing with the payload length.
        let mut padded = valid.clone();
        padded.extend_from_slice(&[0; 4]);
        assert!(Info::parse(&padded).is_err());
        // Inverted CFrtId range.
        assert!(Info::parse(&chart_frt_info_payload(0x9, 0x9, 1, &[(0x085B, 0x085A)])).is_err());
    }

    #[test]
    fn parses_cat_lab_round_trip() {
        for (at, expected) in [
            (0x0001, Alignment::TopLeft),
            (0x0002, Alignment::Center),
            (0x0003, Alignment::BottomRight),
        ] {
            let payload = cat_lab_payload(1000, at, CAT_LAB_AUTO_LABEL);
            let parsed = CatLab::parse(&payload).unwrap();
            assert_eq!(parsed.offset(), 1000);
            assert_eq!(parsed.alignment(), expected);
            assert!(parsed.auto_category_label());
            assert_eq!(parsed.to_payload(), payload);
        }
        let manual = CatLab::parse(&cat_lab_payload(0, 0x0002, 0)).unwrap();
        assert_eq!(manual.offset(), 0);
        assert!(!manual.auto_category_label());
    }

    #[test]
    fn rejects_malformed_cat_lab() {
        let valid = cat_lab_payload(100, 0x0002, 0);
        // Truncated and overlong payloads.
        assert!(CatLab::parse(&valid[..11]).is_err());
        let mut padded = valid.clone();
        padded.push(0);
        assert!(CatLab::parse(&padded).is_err());
        // Wrong FrtHeaderOld.rt.
        let mut wrong_rt = valid.clone();
        wrong_rt[0..2].copy_from_slice(&0x0857u16.to_le_bytes());
        assert!(CatLab::parse(&wrong_rt).is_err());
        // FrtFlags.fFrtRef/fFrtAlert are forbidden for CatLab.
        for flags in [0x0001u16, 0x0002] {
            let mut bad_flags = valid.clone();
            bad_flags[2..4].copy_from_slice(&flags.to_le_bytes());
            assert!(CatLab::parse(&bad_flags).is_err());
        }
        // wOffset above 1000%.
        assert!(CatLab::parse(&cat_lab_payload(1001, 0x0002, 0)).is_err());
        // Unsupported alignment.
        assert!(CatLab::parse(&cat_lab_payload(100, 0x0000, 0)).is_err());
        assert!(CatLab::parse(&cat_lab_payload(100, 0x0004, 0)).is_err());
    }

    #[test]
    fn parses_start_block_round_trip() {
        let payload = start_block_payload(0x0002, 0x0005, 0x0007, 0x0009);
        let parsed = StartBlock::parse(&payload).unwrap();
        assert_eq!(parsed.object_kind(), BlockKind::AttachedLabel);
        assert_eq!(parsed.object_context(), 0x0005);
        assert_eq!(parsed.object_instance1(), 0x0007);
        assert_eq!(parsed.object_instance2(), 0x0009);
        assert_eq!(parsed.to_payload(), payload);

        // Every iObjectKind value from the spec table is accepted.
        for (kind, expected) in [
            (0x0000, BlockKind::AxisGroup),
            (0x0004, BlockKind::Axis),
            (0x0005, BlockKind::ChartGroup),
            (0x0006, BlockKind::Dat),
            (0x0007, BlockKind::Frame),
            (0x0009, BlockKind::Legend),
            (0x000A, BlockKind::LegendException),
            (0x000C, BlockKind::Series),
            (0x000D, BlockKind::Sheet),
            (0x000E, BlockKind::DataFormat),
            (0x000F, BlockKind::DropBar),
        ] {
            let parsed = StartBlock::parse(&start_block_payload(kind, 0, 0, 0)).unwrap();
            assert_eq!(parsed.object_kind(), expected);
        }
    }

    #[test]
    fn rejects_malformed_start_block() {
        let valid = start_block_payload(0x000D, 0, 0, 0);
        // Truncated.
        assert!(StartBlock::parse(&valid[..11]).is_err());
        // Wrong FrtHeaderOld.rt.
        let mut wrong_rt = valid.clone();
        wrong_rt[0..2].copy_from_slice(&END_BLOCK_RECORD_TYPE.to_le_bytes());
        assert!(StartBlock::parse(&wrong_rt).is_err());
        // Unknown iObjectKind.
        assert!(StartBlock::parse(&start_block_payload(0x0001, 0, 0, 0)).is_err());
        assert!(StartBlock::parse(&start_block_payload(0x0010, 0, 0, 0)).is_err());
    }

    #[test]
    fn parses_end_block_round_trip() {
        let payload = end_block_payload(0x000D);
        let parsed = EndBlock::parse(&payload).unwrap();
        assert_eq!(parsed.object_kind(), BlockKind::Sheet);
        assert_eq!(parsed.to_payload(), payload);
    }

    #[test]
    fn preserves_ignored_header_and_record_bytes() {
        let mut info = chart_frt_info_payload(0x9, 0x9, 1, &[(0x0850, 0x085A)]);
        info[2..4].copy_from_slice(&0xFFFCu16.to_le_bytes());
        let parsed = Info::parse(&info).unwrap();
        assert_eq!(parsed.frt_flags(), 0xFFFC);
        assert_eq!(parsed.to_payload(), info);

        let mut cat_lab = cat_lab_payload(100, 0x0002, 0x8001);
        cat_lab[2..4].copy_from_slice(&0xFFFCu16.to_le_bytes());
        cat_lab[10..12].copy_from_slice(&[0xA5, 0x5A]);
        let parsed = CatLab::parse(&cat_lab).unwrap();
        assert_eq!(parsed.frt_flags(), 0xFFFC);
        assert_eq!(parsed.category_flags(), 0x8001);
        assert_eq!(parsed.reserved(), [0xA5, 0x5A]);
        assert_eq!(parsed.to_payload(), cat_lab);

        let mut start = start_block_payload(0x000D, 0x1234, 0x5678, 0x9ABC);
        start[2..4].copy_from_slice(&0xFFFCu16.to_le_bytes());
        let parsed = StartBlock::parse(&start).unwrap();
        assert_eq!(parsed.frt_flags(), 0xFFFC);
        assert_eq!(parsed.to_payload(), start);

        let mut end = end_block_payload(0x000D);
        end[2..4].copy_from_slice(&0xFFFCu16.to_le_bytes());
        end[6..12].copy_from_slice(&[1, 2, 3, 4, 5, 6]);
        let parsed = EndBlock::parse(&end).unwrap();
        assert_eq!(parsed.frt_flags(), 0xFFFC);
        assert_eq!(parsed.unused(), [0x0201, 0x0403, 0x0605]);
        assert_eq!(parsed.to_payload(), end);
    }

    #[test]
    fn rejects_malformed_end_block() {
        let valid = end_block_payload(0x000D);
        // Truncated.
        assert!(EndBlock::parse(&valid[..11]).is_err());
        // Wrong FrtHeaderOld.rt.
        let mut wrong_rt = valid.clone();
        wrong_rt[0..2].copy_from_slice(&START_BLOCK_RECORD_TYPE.to_le_bytes());
        assert!(EndBlock::parse(&wrong_rt).is_err());
        // Unknown iObjectKind.
        assert!(EndBlock::parse(&end_block_payload(0x000B)).is_err());
    }
}
