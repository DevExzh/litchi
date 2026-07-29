//! BIFF8 `ObjectLink` record (0x1027, MS-XLS 2.4.182) of the Chart Sheet
//! substream (MS-XLS 2.1): the chart object a `Text` record is linked to.
//!
//! Everything in this module is INERT: the link values are stored verbatim
//! and no text is attached to a chart object.
//!
//! # References
//!
//! - MS-XLS 2.4.182 (ObjectLink)

use super::{XlsError, XlsResult};

/// Record type of the `ObjectLink` record (MS-XLS 2.4.182).
pub(crate) const OBJECT_LINK_RECORD_TYPE: u16 = 0x1027;

/// Byte length of an `ObjectLink` record payload.
const PAYLOAD_LEN: usize = 6;
/// `wLinkVar2` value selecting the whole series instead of one data point
/// (MS-XLS 2.4.182).
const LINK_TO_SERIES: u16 = 0xFFFF;
/// Maximum `wLinkVar1` value when linking to a series or data point
/// (MS-XLS 2.4.182).
const MAX_SERIES_INDEX: u16 = 254;
/// Maximum `wLinkVar2` value when linking to one data point
/// (MS-XLS 2.4.182).
const MAX_DATA_POINT_INDEX: u16 = 31999;

fn invalid(message: impl Into<String>) -> XlsError {
    XlsError::InvalidRecord {
        record_type: OBJECT_LINK_RECORD_TYPE,
        message: message.into(),
    }
}

/// The `wLinkObj` chart object a `Text` record is linked to (MS-XLS 2.4.182).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u16)]
pub enum XlsObjectLinkTarget {
    /// 0x0001: the entire chart.
    EntireChart = 0x0001,
    /// 0x0002: the value axis (vertical value axis on bubble/scatter groups).
    ValueAxis = 0x0002,
    /// 0x0003: the category axis (horizontal value axis on bubble/scatter
    /// groups).
    CategoryAxis = 0x0003,
    /// 0x0004: a series or data points; `wLinkVar1`/`wLinkVar2` apply.
    SeriesOrDataPoints = 0x0004,
    /// 0x0007: the series axis.
    SeriesAxis = 0x0007,
    /// 0x000C: the display units labels of an axis.
    DisplayUnitsLabels = 0x000C,
}

impl XlsObjectLinkTarget {
    fn parse(value: u16) -> XlsResult<Self> {
        match value {
            0x0001 => Ok(Self::EntireChart),
            0x0002 => Ok(Self::ValueAxis),
            0x0003 => Ok(Self::CategoryAxis),
            0x0004 => Ok(Self::SeriesOrDataPoints),
            0x0007 => Ok(Self::SeriesAxis),
            0x000C => Ok(Self::DisplayUnitsLabels),
            other => Err(invalid(format!(
                "ObjectLink wLinkObj {other:#06X} is not a defined link target"
            ))),
        }
    }
}

/// Typed `ObjectLink` record content (MS-XLS 2.4.182): the chart object a
/// `Text` record is linked to.
///
/// When the target is not `SeriesOrDataPoints`, `wLinkVar1`/`wLinkVar2` MUST
/// be zero and MUST be ignored; they are preserved verbatim so the record
/// round-trips unchanged.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XlsObjectLink {
    /// The chart object the text is linked to (`wLinkObj`).
    target: XlsObjectLinkTarget,
    /// Zero-based index into the Series records (`wLinkVar1`); meaningful
    /// only for `SeriesOrDataPoints`, when it is at most 254.
    series_index: u16,
    /// Zero-based category index within the series (`wLinkVar2`); meaningful
    /// only for `SeriesOrDataPoints`, when it is `0xFFFF` (the whole series)
    /// or at most 31999.
    category_index: u16,
}

impl XlsObjectLink {
    /// Parse an `ObjectLink` record payload.
    pub fn parse(data: &[u8]) -> XlsResult<Self> {
        if data.len() != PAYLOAD_LEN {
            return Err(XlsError::InvalidLength {
                expected: PAYLOAD_LEN,
                found: data.len(),
            });
        }
        let target = XlsObjectLinkTarget::parse(u16::from_le_bytes([data[0], data[1]]))?;
        let series_index = u16::from_le_bytes([data[2], data[3]]);
        let category_index = u16::from_le_bytes([data[4], data[5]]);
        if target == XlsObjectLinkTarget::SeriesOrDataPoints {
            if series_index > MAX_SERIES_INDEX {
                return Err(invalid(format!(
                    "ObjectLink wLinkVar1 {series_index:#06X} exceeds {MAX_SERIES_INDEX:#06X}"
                )));
            }
            if category_index != LINK_TO_SERIES && category_index > MAX_DATA_POINT_INDEX {
                return Err(invalid(format!(
                    "ObjectLink wLinkVar2 {category_index:#06X} is neither {LINK_TO_SERIES:#06X} nor a data point index"
                )));
            }
        }
        Ok(Self {
            target,
            series_index,
            category_index,
        })
    }

    /// Serialize back to a complete `ObjectLink` record payload.
    pub fn to_payload(&self) -> Vec<u8> {
        let mut payload = Vec::with_capacity(PAYLOAD_LEN);
        payload.extend_from_slice(&(self.target as u16).to_le_bytes());
        payload.extend_from_slice(&self.series_index.to_le_bytes());
        payload.extend_from_slice(&self.category_index.to_le_bytes());
        payload
    }

    /// The chart object the text is linked to (`wLinkObj`).
    pub fn target(&self) -> XlsObjectLinkTarget {
        self.target
    }

    /// Zero-based index into the Series records (`wLinkVar1`); preserved
    /// verbatim when the target is not `SeriesOrDataPoints`.
    pub fn series_index(&self) -> u16 {
        self.series_index
    }

    /// Zero-based category index within the series (`wLinkVar2`), or `0xFFFF`
    /// for the whole series; preserved verbatim when the target is not
    /// `SeriesOrDataPoints`.
    pub fn category_index(&self) -> u16 {
        self.category_index
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn record(target: u16, var1: u16, var2: u16) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&target.to_le_bytes());
        data.extend_from_slice(&var1.to_le_bytes());
        data.extend_from_slice(&var2.to_le_bytes());
        data
    }

    #[test]
    fn round_trip_all_targets() {
        for (value, expected) in [
            (0x0001, XlsObjectLinkTarget::EntireChart),
            (0x0002, XlsObjectLinkTarget::ValueAxis),
            (0x0003, XlsObjectLinkTarget::CategoryAxis),
            (0x0004, XlsObjectLinkTarget::SeriesOrDataPoints),
            (0x0007, XlsObjectLinkTarget::SeriesAxis),
            (0x000C, XlsObjectLinkTarget::DisplayUnitsLabels),
        ] {
            let bytes = record(value, 0, 0);
            let parsed = XlsObjectLink::parse(&bytes).unwrap();
            assert_eq!(parsed.target(), expected);
            assert_eq!(parsed.to_payload(), bytes);
        }
    }

    #[test]
    fn series_link_indexes() {
        // Whole-series link.
        let bytes = record(0x0004, 254, 0xFFFF);
        let parsed = XlsObjectLink::parse(&bytes).unwrap();
        assert_eq!(parsed.series_index(), 254);
        assert_eq!(parsed.category_index(), 0xFFFF);
        assert_eq!(parsed.to_payload(), bytes);
        // Single data point link.
        let bytes = record(0x0004, 0, 31999);
        assert_eq!(XlsObjectLink::parse(&bytes).unwrap().to_payload(), bytes);
    }

    #[test]
    fn preserves_ignored_var_fields_for_other_targets() {
        // wLinkVar1/wLinkVar2 MUST be zero and MUST be ignored for non-series
        // targets; they round-trip verbatim.
        let bytes = record(0x0001, 0xAAAA, 0xBBBB);
        let parsed = XlsObjectLink::parse(&bytes).unwrap();
        assert_eq!(parsed.series_index(), 0xAAAA);
        assert_eq!(parsed.to_payload(), bytes);
    }

    #[test]
    fn rejects_malformed_records() {
        let bytes = record(0x0001, 0, 0);
        // Truncated and overlong payloads.
        assert!(XlsObjectLink::parse(&bytes[..5]).is_err());
        assert!(XlsObjectLink::parse(&[bytes.as_slice(), &[0]].concat()).is_err());
        // Undefined wLinkObj.
        assert!(XlsObjectLink::parse(&record(0x0005, 0, 0)).is_err());
        // Series/data-point index bounds.
        assert!(XlsObjectLink::parse(&record(0x0004, 255, 0)).is_err());
        assert!(XlsObjectLink::parse(&record(0x0004, 0, 32000)).is_err());
        assert!(XlsObjectLink::parse(&record(0x0004, 0, 0xFFFE)).is_err());
    }
}
