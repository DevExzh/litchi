//! BIFF8 `Chart3d` record (0x103A, MS-XLS 2.4.46) of the Chart Sheet
//! substream (MS-XLS 2.1): the attributes of a 3-D plot area.
//!
//! Everything in this module is INERT: values are stored verbatim and no 3-D
//! scene is rendered. The chart-group-dependent constraints of MS-XLS 2.4.46
//! (bar/pie/surface group types, `Bar.fTranspose`, and the flag rules tied
//! to them) are cross-record constraints documented on the accessors, not
//! enforced by the record reader.
//!
//! # References
//!
//! - MS-XLS 2.4.46 (Chart3d)

use super::{XlsError, XlsResult};

/// Record type of the `Chart3d` record (MS-XLS 2.4.46).
pub(crate) const CHART_3D_RECORD_TYPE: u16 = 0x103A;

/// Byte length of a `Chart3d` record payload.
const PAYLOAD_LEN: usize = 14;
/// Flags bit: `fPerspective` (vanishing point rendering).
const FLAG_PERSPECTIVE: u16 = 0x0001;
/// Flags bit: `fCluster` (clustered data points in a bar chart group).
const FLAG_CLUSTER: u16 = 0x0002;
/// Flags bit: `f3DScaling` (automatic plot area height).
const FLAG_3D_SCALING: u16 = 0x0004;
/// Flags bit: `fNotPieChart` (chart group type is not pie).
const FLAG_NOT_PIE_CHART: u16 = 0x0010;
/// Flags bit: `fWalls2D` (walls rendered in 2-D).
const FLAG_WALLS_2D: u16 = 0x0020;
/// Maximum `anRot` value, in degrees (MS-XLS 2.4.46).
const MAX_ROTATION: i16 = 360;
/// Minimum `anElev` value, in degrees (MS-XLS 2.4.46).
const MIN_ELEVATION: i16 = -90;
/// Maximum `anElev` value, in degrees (MS-XLS 2.4.46).
const MAX_ELEVATION: i16 = 90;
/// Maximum `pcDist` value (exclusive) (MS-XLS 2.4.46).
const MAX_DISTANCE: i16 = 200;
/// Minimum `pcDepth` value (MS-XLS 2.4.46).
const MIN_DEPTH: i16 = 1;
/// Maximum `pcDepth` value (MS-XLS 2.4.46).
const MAX_DEPTH: i16 = 2000;
/// Maximum `pcGap` value (MS-XLS 2.4.46).
const MAX_GAP: u16 = 500;
/// Maximum `pcHeight` value (exclusive): 65535 (MS-XLS 2.4.46).
const MAX_HEIGHT_EXCLUSIVE: u16 = 0xFFFF;

fn invalid(message: impl Into<String>) -> XlsError {
    XlsError::InvalidRecord {
        record_type: CHART_3D_RECORD_TYPE,
        message: message.into(),
    }
}

/// Typed `Chart3d` record content (MS-XLS 2.4.46): the attributes of a 3-D
/// plot area.
///
/// The `reserved1`/`reserved2` bits (MUST be zero, and MUST be ignored) are
/// preserved verbatim so the record round-trips unchanged.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct XlsChart3d {
    /// Clockwise rotation around the vertical center line, in degrees
    /// (`anRot`), in 0..=360.
    rotation: i16,
    /// Rotation around the horizontal center line, in degrees (`anElev`),
    /// in -90..=90.
    elevation: i16,
    /// Field of view angle (`pcDist`), in 0..200.
    distance: i16,
    /// Pie thickness or 3-D plot area height percentage (`pcHeight`). Its
    /// signedness depends on `fNotPieChart` (MS-XLS 2.4.46); the raw 16-bit
    /// value is stored.
    height: u16,
    /// Depth of the 3-D plot area as a percentage of its width (`pcDepth`),
    /// in 1..=2000.
    depth: i16,
    /// Gap width between series and the plot area edges (`pcGap`), at most
    /// 500.
    gap: u16,
    /// Raw flags word: `fPerspective`, `fCluster`, `f3DScaling`,
    /// `fNotPieChart`, `fWalls2D`, and the 11 reserved bits.
    flags: u16,
}

impl XlsChart3d {
    /// Parse a `Chart3d` record payload.
    pub fn parse(data: &[u8]) -> XlsResult<Self> {
        if data.len() != PAYLOAD_LEN {
            return Err(XlsError::InvalidLength {
                expected: PAYLOAD_LEN,
                found: data.len(),
            });
        }
        let rotation = i16::from_le_bytes([data[0], data[1]]);
        if !(0..=MAX_ROTATION).contains(&rotation) {
            return Err(invalid(format!(
                "Chart3d anRot {rotation} is outside 0..={MAX_ROTATION}"
            )));
        }
        let elevation = i16::from_le_bytes([data[2], data[3]]);
        if !(MIN_ELEVATION..=MAX_ELEVATION).contains(&elevation) {
            return Err(invalid(format!(
                "Chart3d anElev {elevation} is outside {MIN_ELEVATION}..={MAX_ELEVATION}"
            )));
        }
        let distance = i16::from_le_bytes([data[4], data[5]]);
        if !(0..MAX_DISTANCE).contains(&distance) {
            return Err(invalid(format!(
                "Chart3d pcDist {distance} is outside 0..{MAX_DISTANCE}"
            )));
        }
        let height = u16::from_le_bytes([data[6], data[7]]);
        if height == MAX_HEIGHT_EXCLUSIVE {
            return Err(invalid("Chart3d pcHeight is not less than 65535"));
        }
        let depth = i16::from_le_bytes([data[8], data[9]]);
        if !(MIN_DEPTH..=MAX_DEPTH).contains(&depth) {
            return Err(invalid(format!(
                "Chart3d pcDepth {depth} is outside {MIN_DEPTH}..={MAX_DEPTH}"
            )));
        }
        let gap = u16::from_le_bytes([data[10], data[11]]);
        if gap > MAX_GAP {
            return Err(invalid(format!(
                "Chart3d pcGap {gap} exceeds {MAX_GAP}"
            )));
        }
        Ok(Self {
            rotation,
            elevation,
            distance,
            height,
            depth,
            gap,
            flags: u16::from_le_bytes([data[12], data[13]]),
        })
    }

    /// Serialize back to a complete `Chart3d` record payload.
    pub fn to_payload(&self) -> Vec<u8> {
        let mut payload = Vec::with_capacity(PAYLOAD_LEN);
        payload.extend_from_slice(&self.rotation.to_le_bytes());
        payload.extend_from_slice(&self.elevation.to_le_bytes());
        payload.extend_from_slice(&self.distance.to_le_bytes());
        payload.extend_from_slice(&self.height.to_le_bytes());
        payload.extend_from_slice(&self.depth.to_le_bytes());
        payload.extend_from_slice(&self.gap.to_le_bytes());
        payload.extend_from_slice(&self.flags.to_le_bytes());
        payload
    }

    /// Clockwise rotation around the vertical center line, in degrees
    /// (`anRot`). MUST be at most 44 for transposed bar chart groups
    /// (MS-XLS 2.4.46; cross-record constraint).
    pub fn rotation(&self) -> i16 {
        self.rotation
    }

    /// Rotation around the horizontal center line, in degrees (`anElev`).
    pub fn elevation(&self) -> i16 {
        self.elevation
    }

    /// Field of view angle (`pcDist`).
    pub fn distance(&self) -> i16 {
        self.distance
    }

    /// Pie thickness or 3-D plot area height percentage (`pcHeight`), as the
    /// raw 16-bit value; its interpretation depends on `fNotPieChart`.
    pub fn height(&self) -> u16 {
        self.height
    }

    /// Depth of the 3-D plot area as a percentage of its width (`pcDepth`).
    pub fn depth(&self) -> i16 {
        self.depth
    }

    /// Gap width between series and the plot area edges (`pcGap`).
    pub fn gap(&self) -> u16 {
        self.gap
    }

    /// Whether the plot area is rendered with a vanishing point
    /// (`fPerspective`). MUST be 0 for pie chart groups (MS-XLS 2.4.46).
    pub fn perspective(&self) -> bool {
        self.flags & FLAG_PERSPECTIVE != 0
    }

    /// Whether data points are clustered in a bar chart group (`fCluster`).
    pub fn cluster(&self) -> bool {
        self.flags & FLAG_CLUSTER != 0
    }

    /// Whether the plot area height is automatically determined
    /// (`f3DScaling`).
    pub fn auto_scaling(&self) -> bool {
        self.flags & FLAG_3D_SCALING != 0
    }

    /// Whether the chart group type is not pie (`fNotPieChart`).
    pub fn not_pie_chart(&self) -> bool {
        self.flags & FLAG_NOT_PIE_CHART != 0
    }

    /// Whether the walls are rendered in 2-D (`fWalls2D`).
    pub fn walls_2d(&self) -> bool {
        self.flags & FLAG_WALLS_2D != 0
    }

    /// Raw flags word, including the 11 reserved bits.
    pub fn flags(&self) -> u16 {
        self.flags
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn record(fields: [i16; 3], height: u16, depth: i16, gap: u16, flags: u16) -> Vec<u8> {
        let mut data = Vec::new();
        for field in fields {
            data.extend_from_slice(&field.to_le_bytes());
        }
        data.extend_from_slice(&height.to_le_bytes());
        data.extend_from_slice(&depth.to_le_bytes());
        data.extend_from_slice(&gap.to_le_bytes());
        data.extend_from_slice(&flags.to_le_bytes());
        data
    }

    #[test]
    fn round_trip() {
        let bytes = record([30, 15, 20], 100, 100, 150, 0x0037);
        let parsed = XlsChart3d::parse(&bytes).unwrap();
        assert_eq!(parsed.rotation(), 30);
        assert_eq!(parsed.elevation(), 15);
        assert_eq!(parsed.distance(), 20);
        assert_eq!(parsed.height(), 100);
        assert_eq!(parsed.depth(), 100);
        assert_eq!(parsed.gap(), 150);
        assert!(parsed.perspective());
        assert!(parsed.cluster());
        assert!(parsed.auto_scaling());
        assert!(parsed.not_pie_chart());
        assert!(parsed.walls_2d());
        assert_eq!(parsed.to_payload(), bytes);
    }

    #[test]
    fn accepts_bounds_and_preserves_reserved_bits() {
        assert!(XlsChart3d::parse(&record([0, -90, 0], 0, 1, 0, 0)).is_ok());
        assert!(XlsChart3d::parse(&record([360, 90, 199], 0xFFFE, 2000, 500, 0)).is_ok());
        // The 11 reserved bits MUST be ignored but round-trip verbatim.
        let bytes = record([0, 0, 0], 0, 1, 0, 0xFFC8);
        let parsed = XlsChart3d::parse(&bytes).unwrap();
        assert_eq!(parsed.flags(), 0xFFC8);
        assert!(!parsed.perspective());
        assert_eq!(parsed.to_payload(), bytes);
    }

    #[test]
    fn rejects_malformed_records() {
        let bytes = record([0, 0, 0], 0, 1, 0, 0);
        // Truncated and overlong payloads.
        assert!(XlsChart3d::parse(&bytes[..13]).is_err());
        assert!(XlsChart3d::parse(&[bytes.as_slice(), &[0]].concat()).is_err());
        // Field bounds.
        assert!(XlsChart3d::parse(&record([-1, 0, 0], 0, 1, 0, 0)).is_err());
        assert!(XlsChart3d::parse(&record([361, 0, 0], 0, 1, 0, 0)).is_err());
        assert!(XlsChart3d::parse(&record([0, -91, 0], 0, 1, 0, 0)).is_err());
        assert!(XlsChart3d::parse(&record([0, 91, 0], 0, 1, 0, 0)).is_err());
        assert!(XlsChart3d::parse(&record([0, 0, -1], 0, 1, 0, 0)).is_err());
        assert!(XlsChart3d::parse(&record([0, 0, 200], 0, 1, 0, 0)).is_err());
        assert!(XlsChart3d::parse(&record([0, 0, 0], 0xFFFF, 1, 0, 0)).is_err());
        assert!(XlsChart3d::parse(&record([0, 0, 0], 0, 0, 0, 0)).is_err());
        assert!(XlsChart3d::parse(&record([0, 0, 0], 0, 2001, 0, 0)).is_err());
        assert!(XlsChart3d::parse(&record([0, 0, 0], 0, 1, 501, 0)).is_err());
    }
}
