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

use super::{XlsError, XlsResult};

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
fn read_u16(data: &[u8], offset: usize) -> XlsResult<u16> {
    let end = offset.checked_add(2).ok_or(XlsError::InvalidLength {
        expected: 2,
        found: data.len(),
    })?;
    let bytes = data.get(offset..end).ok_or(XlsError::InvalidLength {
        expected: end,
        found: data.len(),
    })?;
    let [first, second] = bytes.try_into().map_err(|_| XlsError::InvalidLength {
        expected: 2,
        found: bytes.len(),
    })?;
    Ok(u16::from_le_bytes([first, second]))
}

/// Validate the `rt` field of an `FrtHeaderOld` and the payload length.
fn validate_frt_header_old(data: &[u8], record_type: u16, expected_len: usize) -> XlsResult<()> {
    if data.len() != expected_len {
        return Err(XlsError::InvalidLength {
            expected: expected_len,
            found: data.len(),
        });
    }
    if read_u16(data, 0)? != record_type {
        return Err(XlsError::InvalidRecord {
            record_type,
            message: "FrtHeaderOld.rt mismatch".to_string(),
        });
    }
    Ok(())
}

/// Application version that created or last saved the file, as declared by a
/// `ChartFrtInfo` record (MS-XLS 2.4.49). Variant names carry the numeric
/// version value (0x9, 0xA, 0xC, 0xE).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlsChartFrtVersion {
    /// Application version 0x9; the chart uses one `CFrtId` range.
    Version9,
    /// Application version 0xA; the chart uses three `CFrtId` ranges.
    Version10,
    /// Application version 0xC; the chart uses four `CFrtId` ranges.
    Version12,
    /// Application version 0xE; the chart uses four `CFrtId` ranges.
    Version14,
}

impl XlsChartFrtVersion {
    /// Decode a raw version byte.
    fn from_byte(byte: u8, record_type: u16) -> XlsResult<Self> {
        match byte {
            0x9 => Ok(Self::Version9),
            0xA => Ok(Self::Version10),
            0xC => Ok(Self::Version12),
            0xE => Ok(Self::Version14),
            _ => Err(XlsError::InvalidRecord {
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
pub struct XlsChartFutureRecordRange {
    /// First Future Record Type in the range (`rtFirst`).
    first: u16,
    /// Last Future Record Type in the range (`rtLast`).
    last: u16,
}

impl XlsChartFutureRecordRange {
    /// First Future Record Type in the range.
    pub fn first(&self) -> u16 {
        self.first
    }

    /// Last Future Record Type in the range.
    pub fn last(&self) -> u16 {
        self.last
    }

    /// Construct a range, rejecting an inverted interval.
    pub(crate) fn new(first: u16, last: u16) -> XlsResult<Self> {
        if first > last {
            return Err(XlsError::InvalidRecord {
                record_type: CHART_FRT_INFO_RECORD_TYPE,
                message: format!("CFrtId range {first:#06X}..={last:#06X} is inverted"),
            });
        }
        Ok(Self { first, last })
    }
}

/// Typed `ChartFrtInfo` record content (MS-XLS 2.4.49).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsChartFrtInfo {
    /// Application version that originally created the file (`verOriginator`).
    originator: XlsChartFrtVersion,
    /// Application version that last saved the file (`verWriter`).
    writer: XlsChartFrtVersion,
    /// Ranges of Future Record Type identifiers used in the chart (`rgCFRTID`).
    ranges: Vec<XlsChartFutureRecordRange>,
}

impl XlsChartFrtInfo {
    /// Application version that originally created the file.
    pub fn originator(&self) -> XlsChartFrtVersion {
        self.originator
    }

    /// Application version that last saved the file.
    pub fn writer(&self) -> XlsChartFrtVersion {
        self.writer
    }

    /// Ranges of Future Record Type identifiers used in the chart.
    pub fn ranges(&self) -> &[XlsChartFutureRecordRange] {
        &self.ranges
    }

    /// Parse a `ChartFrtInfo` record payload.
    pub fn parse(data: &[u8]) -> XlsResult<Self> {
        let invalid = |message: String| XlsError::InvalidRecord {
            record_type: CHART_FRT_INFO_RECORD_TYPE,
            message,
        };
        if data.len() < CHART_FRT_INFO_BASE_LEN {
            return Err(XlsError::InvalidLength {
                expected: CHART_FRT_INFO_BASE_LEN,
                found: data.len(),
            });
        }
        if read_u16(data, 0)? != CHART_FRT_INFO_RECORD_TYPE {
            return Err(invalid("ChartFrtInfo FrtHeaderOld.rt mismatch".to_string()));
        }
        let originator = XlsChartFrtVersion::from_byte(data[4], CHART_FRT_INFO_RECORD_TYPE)?;
        let writer = XlsChartFrtVersion::from_byte(data[5], CHART_FRT_INFO_RECORD_TYPE)?;
        let count = usize::from(read_u16(data, 6)?);
        if count != writer.expected_range_count() {
            return Err(invalid(format!(
                "ChartFrtInfo cCFRTID {count} disagrees with the writer version"
            )));
        }
        let expected_len = CHART_FRT_INFO_BASE_LEN + count * C_FRT_ID_LEN;
        if data.len() != expected_len {
            return Err(XlsError::InvalidLength {
                expected: expected_len,
                found: data.len(),
            });
        }
        let mut ranges = Vec::with_capacity(count);
        for index in 0..count {
            let offset = CHART_FRT_INFO_BASE_LEN + index * C_FRT_ID_LEN;
            ranges.push(XlsChartFutureRecordRange::new(
                read_u16(data, offset)?,
                read_u16(data, offset + 2)?,
            )?);
        }
        Ok(Self {
            originator,
            writer,
            ranges,
        })
    }

    /// Serialize back to a complete `ChartFrtInfo` record payload.
    pub fn to_payload(&self) -> Vec<u8> {
        let mut payload =
            Vec::with_capacity(CHART_FRT_INFO_BASE_LEN + self.ranges.len() * C_FRT_ID_LEN);
        // FrtHeaderOld: rt, grbitFrt (0).
        payload.extend_from_slice(&CHART_FRT_INFO_RECORD_TYPE.to_le_bytes());
        payload.extend_from_slice(&[0; FRT_HEADER_OLD_LEN - 2]);
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
pub enum XlsCatLabAlignment {
    /// Top-aligned when the axis text is rotated; left-aligned under a
    /// left-to-right reading order, right-aligned otherwise.
    TopLeft,
    /// Center-aligned.
    Center,
    /// Bottom-aligned when the axis text is rotated; right-aligned under a
    /// left-to-right reading order, left-aligned otherwise.
    BottomRight,
}

impl XlsCatLabAlignment {
    /// Decode the raw `at` value.
    fn from_u16(value: u16) -> XlsResult<Self> {
        match value {
            0x0001 => Ok(Self::TopLeft),
            0x0002 => Ok(Self::Center),
            0x0003 => Ok(Self::BottomRight),
            _ => Err(XlsError::InvalidRecord {
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
pub struct XlsCatLab {
    /// Distance between the axis and the axis label, as a percentage of the
    /// default distance (`wOffset`, 0..=1000).
    offset: u16,
    /// Alignment of the axis label (`at`).
    alignment: XlsCatLabAlignment,
    /// Whether the number of categories between axis labels is automatically
    /// calculated (`cAutoCatLabelReal`).
    auto_category_label: bool,
}

impl XlsCatLab {
    /// Distance between the axis and the axis label (percent of the default).
    pub fn offset(&self) -> u16 {
        self.offset
    }

    /// Alignment of the axis label.
    pub fn alignment(&self) -> XlsCatLabAlignment {
        self.alignment
    }

    /// Whether the number of categories between axis labels is automatic.
    pub fn auto_category_label(&self) -> bool {
        self.auto_category_label
    }

    /// Parse a `CatLab` record payload.
    pub fn parse(data: &[u8]) -> XlsResult<Self> {
        validate_frt_header_old(data, CAT_LAB_RECORD_TYPE, CAT_LAB_LEN)?;
        let offset = read_u16(data, 4)?;
        if offset > CAT_LAB_MAX_OFFSET {
            return Err(XlsError::InvalidRecord {
                record_type: CAT_LAB_RECORD_TYPE,
                message: format!("CatLab wOffset {offset} exceeds {CAT_LAB_MAX_OFFSET}"),
            });
        }
        let alignment = XlsCatLabAlignment::from_u16(read_u16(data, 6)?)?;
        let flags = read_u16(data, 8)?;
        // Bytes 10..12 are reserved (MUST be zero) and MUST be ignored.
        Ok(Self {
            offset,
            alignment,
            auto_category_label: flags & CAT_LAB_AUTO_LABEL != 0,
        })
    }

    /// Serialize back to a complete `CatLab` record payload.
    pub fn to_payload(&self) -> Vec<u8> {
        let mut payload = Vec::with_capacity(CAT_LAB_LEN);
        // FrtHeaderOld: rt, grbitFrt (0).
        payload.extend_from_slice(&CAT_LAB_RECORD_TYPE.to_le_bytes());
        payload.extend_from_slice(&[0; FRT_HEADER_OLD_LEN - 2]);
        payload.extend_from_slice(&self.offset.to_le_bytes());
        payload.extend_from_slice(&self.alignment.to_u16().to_le_bytes());
        let flags = if self.auto_category_label {
            CAT_LAB_AUTO_LABEL
        } else {
            0
        };
        payload.extend_from_slice(&flags.to_le_bytes());
        // reserved (0).
        payload.extend_from_slice(&[0; 2]);
        payload
    }
}

/// Type of chart object encompassed by a `StartBlock`/`EndBlock` pair
/// (`iObjectKind`, MS-XLS 2.4.266 / 2.4.100).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlsChartBlockObjectKind {
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

impl XlsChartBlockObjectKind {
    /// Decode the raw `iObjectKind` value.
    fn from_u16(value: u16, record_type: u16) -> XlsResult<Self> {
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
            _ => Err(XlsError::InvalidRecord {
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
pub struct XlsStartBlock {
    /// Type of object encompassed by the block (`iObjectKind`).
    object_kind: XlsChartBlockObjectKind,
    /// Context of the object (`iObjectContext`); its meaning depends on
    /// `object_kind` and is preserved verbatim.
    object_context: u16,
    /// First instance qualifier (`iObjectInstance1`), preserved verbatim.
    object_instance1: u16,
    /// Second instance qualifier (`iObjectInstance2`), preserved verbatim.
    object_instance2: u16,
}

impl XlsStartBlock {
    /// Type of object encompassed by the block.
    pub fn object_kind(&self) -> XlsChartBlockObjectKind {
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
    pub fn parse(data: &[u8]) -> XlsResult<Self> {
        validate_frt_header_old(data, START_BLOCK_RECORD_TYPE, START_BLOCK_LEN)?;
        Ok(Self {
            object_kind: XlsChartBlockObjectKind::from_u16(
                read_u16(data, 4)?,
                START_BLOCK_RECORD_TYPE,
            )?,
            object_context: read_u16(data, 6)?,
            object_instance1: read_u16(data, 8)?,
            object_instance2: read_u16(data, 10)?,
        })
    }

    /// Serialize back to a complete `StartBlock` record payload.
    pub fn to_payload(&self) -> Vec<u8> {
        let mut payload = Vec::with_capacity(START_BLOCK_LEN);
        // FrtHeaderOld: rt, grbitFrt (0).
        payload.extend_from_slice(&START_BLOCK_RECORD_TYPE.to_le_bytes());
        payload.extend_from_slice(&[0; FRT_HEADER_OLD_LEN - 2]);
        payload.extend_from_slice(&self.object_kind.to_u16().to_le_bytes());
        payload.extend_from_slice(&self.object_context.to_le_bytes());
        payload.extend_from_slice(&self.object_instance1.to_le_bytes());
        payload.extend_from_slice(&self.object_instance2.to_le_bytes());
        payload
    }
}

/// Typed `EndBlock` record content (MS-XLS 2.4.100): the end of the collection
/// opened by the associated `StartBlock` record.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XlsEndBlock {
    /// Type of object encompassed by the block (`iObjectKind`); MUST equal the
    /// `iObjectKind` of the associated `StartBlock` record.
    object_kind: XlsChartBlockObjectKind,
}

impl XlsEndBlock {
    /// Type of object encompassed by the block.
    pub fn object_kind(&self) -> XlsChartBlockObjectKind {
        self.object_kind
    }

    /// Parse an `EndBlock` record payload.
    pub fn parse(data: &[u8]) -> XlsResult<Self> {
        validate_frt_header_old(data, END_BLOCK_RECORD_TYPE, END_BLOCK_LEN)?;
        // Bytes 6..12 are undefined (unused1..3) and MUST be ignored.
        Ok(Self {
            object_kind: XlsChartBlockObjectKind::from_u16(
                read_u16(data, 4)?,
                END_BLOCK_RECORD_TYPE,
            )?,
        })
    }

    /// Serialize back to a complete `EndBlock` record payload.
    pub fn to_payload(&self) -> Vec<u8> {
        let mut payload = Vec::with_capacity(END_BLOCK_LEN);
        // FrtHeaderOld: rt, grbitFrt (0).
        payload.extend_from_slice(&END_BLOCK_RECORD_TYPE.to_le_bytes());
        payload.extend_from_slice(&[0; FRT_HEADER_OLD_LEN - 2]);
        payload.extend_from_slice(&self.object_kind.to_u16().to_le_bytes());
        // unused1..3 (0).
        payload.extend_from_slice(&[0; 6]);
        payload
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
        let parsed = XlsChartFrtInfo::parse(&payload).unwrap();
        assert_eq!(parsed.originator(), XlsChartFrtVersion::Version9);
        assert_eq!(parsed.writer(), XlsChartFrtVersion::Version9);
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
        let parsed = XlsChartFrtInfo::parse(&payload).unwrap();
        assert_eq!(parsed.writer(), XlsChartFrtVersion::Version10);
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
            let parsed = XlsChartFrtInfo::parse(&payload).unwrap();
            assert_eq!(parsed.ranges().len(), 4);
            assert_eq!(parsed.to_payload(), payload);
        }
    }

    #[test]
    fn rejects_malformed_chart_frt_info() {
        let valid = chart_frt_info_payload(0x9, 0x9, 1, &[(0x0850, 0x085A)]);
        // Truncated.
        assert!(XlsChartFrtInfo::parse(&valid[..7]).is_err());
        // Wrong FrtHeaderOld.rt.
        let mut wrong_rt = valid.clone();
        wrong_rt[0..2].copy_from_slice(&0x0851u16.to_le_bytes());
        assert!(XlsChartFrtInfo::parse(&wrong_rt).is_err());
        // Unsupported originator/writer versions.
        assert!(
            XlsChartFrtInfo::parse(&chart_frt_info_payload(0xB, 0x9, 1, &[(0x0850, 0x085A)]))
                .is_err()
        );
        assert!(
            XlsChartFrtInfo::parse(&chart_frt_info_payload(0x9, 0xF, 1, &[(0x0850, 0x085A)]))
                .is_err()
        );
        // cCFRTID disagrees with the writer version.
        assert!(
            XlsChartFrtInfo::parse(&chart_frt_info_payload(0x9, 0x9, 3, &[(0x0850, 0x085A); 3]))
                .is_err()
        );
        assert!(
            XlsChartFrtInfo::parse(&chart_frt_info_payload(0x9, 0xA, 1, &[(0x0850, 0x085A)]))
                .is_err()
        );
        // Declared count disagreeing with the payload length.
        let mut padded = valid.clone();
        padded.extend_from_slice(&[0; 4]);
        assert!(XlsChartFrtInfo::parse(&padded).is_err());
        // Inverted CFrtId range.
        assert!(
            XlsChartFrtInfo::parse(&chart_frt_info_payload(0x9, 0x9, 1, &[(0x085B, 0x085A)]))
                .is_err()
        );
    }

    #[test]
    fn parses_cat_lab_round_trip() {
        for (at, expected) in [
            (0x0001, XlsCatLabAlignment::TopLeft),
            (0x0002, XlsCatLabAlignment::Center),
            (0x0003, XlsCatLabAlignment::BottomRight),
        ] {
            let payload = cat_lab_payload(1000, at, CAT_LAB_AUTO_LABEL);
            let parsed = XlsCatLab::parse(&payload).unwrap();
            assert_eq!(parsed.offset(), 1000);
            assert_eq!(parsed.alignment(), expected);
            assert!(parsed.auto_category_label());
            assert_eq!(parsed.to_payload(), payload);
        }
        let manual = XlsCatLab::parse(&cat_lab_payload(0, 0x0002, 0)).unwrap();
        assert_eq!(manual.offset(), 0);
        assert!(!manual.auto_category_label());
    }

    #[test]
    fn rejects_malformed_cat_lab() {
        let valid = cat_lab_payload(100, 0x0002, 0);
        // Truncated and overlong payloads.
        assert!(XlsCatLab::parse(&valid[..11]).is_err());
        let mut padded = valid.clone();
        padded.push(0);
        assert!(XlsCatLab::parse(&padded).is_err());
        // Wrong FrtHeaderOld.rt.
        let mut wrong_rt = valid.clone();
        wrong_rt[0..2].copy_from_slice(&0x0857u16.to_le_bytes());
        assert!(XlsCatLab::parse(&wrong_rt).is_err());
        // wOffset above 1000%.
        assert!(XlsCatLab::parse(&cat_lab_payload(1001, 0x0002, 0)).is_err());
        // Unsupported alignment.
        assert!(XlsCatLab::parse(&cat_lab_payload(100, 0x0000, 0)).is_err());
        assert!(XlsCatLab::parse(&cat_lab_payload(100, 0x0004, 0)).is_err());
    }

    #[test]
    fn parses_start_block_round_trip() {
        let payload = start_block_payload(0x0002, 0x0005, 0x0007, 0x0009);
        let parsed = XlsStartBlock::parse(&payload).unwrap();
        assert_eq!(parsed.object_kind(), XlsChartBlockObjectKind::AttachedLabel);
        assert_eq!(parsed.object_context(), 0x0005);
        assert_eq!(parsed.object_instance1(), 0x0007);
        assert_eq!(parsed.object_instance2(), 0x0009);
        assert_eq!(parsed.to_payload(), payload);

        // Every iObjectKind value from the spec table is accepted.
        for (kind, expected) in [
            (0x0000, XlsChartBlockObjectKind::AxisGroup),
            (0x0004, XlsChartBlockObjectKind::Axis),
            (0x0005, XlsChartBlockObjectKind::ChartGroup),
            (0x0006, XlsChartBlockObjectKind::Dat),
            (0x0007, XlsChartBlockObjectKind::Frame),
            (0x0009, XlsChartBlockObjectKind::Legend),
            (0x000A, XlsChartBlockObjectKind::LegendException),
            (0x000C, XlsChartBlockObjectKind::Series),
            (0x000D, XlsChartBlockObjectKind::Sheet),
            (0x000E, XlsChartBlockObjectKind::DataFormat),
            (0x000F, XlsChartBlockObjectKind::DropBar),
        ] {
            let parsed = XlsStartBlock::parse(&start_block_payload(kind, 0, 0, 0)).unwrap();
            assert_eq!(parsed.object_kind(), expected);
        }
    }

    #[test]
    fn rejects_malformed_start_block() {
        let valid = start_block_payload(0x000D, 0, 0, 0);
        // Truncated.
        assert!(XlsStartBlock::parse(&valid[..11]).is_err());
        // Wrong FrtHeaderOld.rt.
        let mut wrong_rt = valid.clone();
        wrong_rt[0..2].copy_from_slice(&END_BLOCK_RECORD_TYPE.to_le_bytes());
        assert!(XlsStartBlock::parse(&wrong_rt).is_err());
        // Unknown iObjectKind.
        assert!(XlsStartBlock::parse(&start_block_payload(0x0001, 0, 0, 0)).is_err());
        assert!(XlsStartBlock::parse(&start_block_payload(0x0010, 0, 0, 0)).is_err());
    }

    #[test]
    fn parses_end_block_round_trip() {
        let payload = end_block_payload(0x000D);
        let parsed = XlsEndBlock::parse(&payload).unwrap();
        assert_eq!(parsed.object_kind(), XlsChartBlockObjectKind::Sheet);
        assert_eq!(parsed.to_payload(), payload);
    }

    #[test]
    fn rejects_malformed_end_block() {
        let valid = end_block_payload(0x000D);
        // Truncated.
        assert!(XlsEndBlock::parse(&valid[..11]).is_err());
        // Wrong FrtHeaderOld.rt.
        let mut wrong_rt = valid.clone();
        wrong_rt[0..2].copy_from_slice(&START_BLOCK_RECORD_TYPE.to_le_bytes());
        assert!(XlsEndBlock::parse(&wrong_rt).is_err());
        // Unknown iObjectKind.
        assert!(XlsEndBlock::parse(&end_block_payload(0x000B)).is_err());
    }
}
