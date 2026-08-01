//! BIFF8 chart layout future records of the Chart Sheet substream
//! (MS-XLS 2.1):
//!
//! - **CrtLayout12** (0x089D): layout information for an attached label or
//!   legend (MS-XLS 2.4.66).
//! - **CrtLayout12A** (0x08A7): layout information for a plot area
//!   (MS-XLS 2.4.67).
//!
//! Everything in this module is INERT: layout values and checksums are stored
//! verbatim and no chart layout is computed or applied. The `dwCheckSum`
//! fields are preserved as read; the checksum inputs include data owned by
//! other records, so they are not recomputed here.
//!
//! # References
//!
//! - MS-XLS 2.4.66 (CrtLayout12), 2.4.67 (CrtLayout12A), 2.5.62
//!   (CrtLayout12Mode), 2.5.134 (FrtFlags), 2.5.135 (FrtHeader), 2.5.342
//!   (Xnum)

use super::{XlsError, XlsResult};

/// Record type of the `CrtLayout12` record (MS-XLS 2.4.66); also the required
/// `frtHeader.rt` value.
pub(crate) const CRT_LAYOUT_12_RECORD_TYPE: u16 = 0x089D;

/// Record type of the `CrtLayout12A` record (MS-XLS 2.4.67); also the
/// required `frtHeader.rt` value.
pub(crate) const CRT_LAYOUT_12_A_RECORD_TYPE: u16 = 0x08A7;

/// Size in bytes of an `FrtHeader` (MS-XLS 2.5.135).
const FRT_HEADER_LEN: usize = 12;
/// `FrtFlags` bits that MUST be zero in an `FrtHeader` (MS-XLS 2.5.135):
/// `fFrtRef` and `fFrtAlert`.
const FRT_FLAGS_FORBIDDEN: u16 = 0x0003;
/// Byte length of a `CrtLayout12` record payload.
const CRT_LAYOUT_12_LEN: usize = 60;
/// Byte length of a `CrtLayout12A` record payload.
const CRT_LAYOUT_12_A_LEN: usize = 68;
/// Mask of the 4-bit `autolayouttype` field of `CrtLayout12` (MS-XLS 2.4.66).
const AUTO_LAYOUT_TYPE_MASK: u16 = 0x001E;
/// `CrtLayout12A` flag: the layout target is the inner plot area
/// (MS-XLS 2.4.67).
const LAYOUT_TARGET_INNER: u16 = 0x0001;

fn invalid(record_type: u16, message: impl Into<String>) -> XlsError {
    XlsError::InvalidRecord {
        record_type,
        message: message.into(),
    }
}

/// A `CrtLayout12Mode` layout mode (MS-XLS 2.5.62): the meaning of the `x`,
/// `y`, `dx`, and `dy` fields.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u16)]
pub enum XlsCrtLayout12Mode {
    /// 0x0000: position and dimension are determined by the application.
    Auto = 0x0000,
    /// 0x0001: `x`/`y` are offsets from the default position, `dx`/`dy` are
    /// dimensions, all as fractions of the chart area.
    Factor = 0x0001,
    /// 0x0002: `x`/`y` are the upper-left corner, `dx`/`dy` the bottom-right
    /// corner, all as fractions of the chart area.
    Edge = 0x0002,
}

impl XlsCrtLayout12Mode {
    fn parse(value: u16, record_type: u16) -> XlsResult<Self> {
        match value {
            0x0000 => Ok(Self::Auto),
            0x0001 => Ok(Self::Factor),
            0x0002 => Ok(Self::Edge),
            other => Err(invalid(
                record_type,
                format!("CrtLayout12Mode {other:#06X} is not a defined layout mode"),
            )),
        }
    }
}

/// The four layout modes and values shared by `CrtLayout12` and
/// `CrtLayout12A`.
#[derive(Debug, Clone, Copy, PartialEq)]
struct LayoutModes {
    x_mode: XlsCrtLayout12Mode,
    y_mode: XlsCrtLayout12Mode,
    width_mode: XlsCrtLayout12Mode,
    height_mode: XlsCrtLayout12Mode,
    /// Raw `x`/`y`/`dx`/`dy` Xnum bit patterns (MS-XLS 2.5.342).
    x: f64,
    y: f64,
    dx: f64,
    dy: f64,
}

impl LayoutModes {
    fn parse(data: &[u8], offset: usize, record_type: u16) -> XlsResult<Self> {
        let mode = |index: usize| {
            XlsCrtLayout12Mode::parse(
                u16::from_le_bytes([data[offset + index], data[offset + index + 1]]),
                record_type,
            )
        };
        let xnum = |index: usize| {
            f64::from_le_bytes(
                data[offset + index..offset + index + 8]
                    .try_into()
                    .expect("sized"),
            )
        };
        Ok(Self {
            x_mode: mode(0)?,
            y_mode: mode(2)?,
            width_mode: mode(4)?,
            height_mode: mode(6)?,
            x: xnum(8),
            y: xnum(16),
            dx: xnum(24),
            dy: xnum(32),
        })
    }

    fn write_payload(&self, output: &mut Vec<u8>) {
        for mode in [self.x_mode, self.y_mode, self.width_mode, self.height_mode] {
            output.extend_from_slice(&(mode as u16).to_le_bytes());
        }
        for value in [self.x, self.y, self.dx, self.dy] {
            output.extend_from_slice(&value.to_le_bytes());
        }
    }
}

/// Validate an `FrtHeader` (MS-XLS 2.5.135): the `rt` field and the
/// `fFrtRef`/`fFrtAlert` bits that MUST be zero.
fn validate_frt_header(data: &[u8], record_type: u16, name: &str) -> XlsResult<u16> {
    if u16::from_le_bytes([data[0], data[1]]) != record_type {
        return Err(invalid(
            record_type,
            format!("{name} FrtHeader.rt mismatch"),
        ));
    }
    let flags = u16::from_le_bytes([data[2], data[3]]);
    if flags & FRT_FLAGS_FORBIDDEN != 0 {
        return Err(invalid(
            record_type,
            format!("{name} FrtHeader.grbitFrt {flags:#06X} sets fFrtRef or fFrtAlert"),
        ));
    }
    Ok(flags)
}

/// Typed `CrtLayout12` record content (MS-XLS 2.4.66): layout information
/// for an attached label or legend.
///
/// The unused flag bit, the 11 `reserved1` bits, and the trailing `reserved2`
/// field (MUST be ignored) are preserved verbatim so the record round-trips
/// unchanged.
#[derive(Debug, Clone, PartialEq)]
pub struct XlsCrtLayout12 {
    /// Raw `frtHeader.grbitFrt` bitfield (`fFrtRef`/`fFrtAlert` are zero).
    frt_flags: u16,
    /// `frtHeader.reserved` bytes, preserved verbatim.
    frt_reserved: [u8; 8],
    /// Raw `dwCheckSum` of the layout values, preserved verbatim.
    checksum: u32,
    /// Raw flags: unused bit, 4-bit `autolayouttype`, and 11 `reserved1` bits.
    flags: u16,
    modes: LayoutModes,
    /// Trailing `reserved2` field, preserved verbatim.
    reserved2: u16,
}

impl XlsCrtLayout12 {
    /// Parse a `CrtLayout12` record payload.
    pub fn parse(data: &[u8]) -> XlsResult<Self> {
        if data.len() != CRT_LAYOUT_12_LEN {
            return Err(XlsError::InvalidLength {
                expected: CRT_LAYOUT_12_LEN,
                found: data.len(),
            });
        }
        let frt_flags = validate_frt_header(data, CRT_LAYOUT_12_RECORD_TYPE, "CrtLayout12")?;
        Ok(Self {
            frt_flags,
            frt_reserved: data[4..FRT_HEADER_LEN].try_into().expect("length checked"),
            checksum: u32::from_le_bytes(data[12..16].try_into().expect("length checked")),
            flags: u16::from_le_bytes([data[16], data[17]]),
            modes: LayoutModes::parse(data, 18, CRT_LAYOUT_12_RECORD_TYPE)?,
            reserved2: u16::from_le_bytes([data[58], data[59]]),
        })
    }

    /// Serialize back to a complete `CrtLayout12` record payload.
    pub fn to_payload(&self) -> Vec<u8> {
        let mut payload = Vec::with_capacity(CRT_LAYOUT_12_LEN);
        payload.extend_from_slice(&CRT_LAYOUT_12_RECORD_TYPE.to_le_bytes());
        payload.extend_from_slice(&self.frt_flags.to_le_bytes());
        payload.extend_from_slice(&self.frt_reserved);
        payload.extend_from_slice(&self.checksum.to_le_bytes());
        payload.extend_from_slice(&self.flags.to_le_bytes());
        self.modes.write_payload(&mut payload);
        payload.extend_from_slice(&self.reserved2.to_le_bytes());
        payload
    }

    /// The automatic layout type of the legend (`autolayouttype`, 4 bits).
    /// MUST be ignored when the record is in an ATTACHEDLABEL rule sequence
    /// (MS-XLS 2.4.66); defined values are 0x0 through 0x4.
    pub fn auto_layout_type(&self) -> u8 {
        ((self.flags & AUTO_LAYOUT_TYPE_MASK) >> 1) as u8
    }

    /// Raw flags word, including the unused and `reserved1` bits.
    pub fn flags(&self) -> u16 {
        self.flags
    }

    /// Raw `dwCheckSum` value, preserved verbatim.
    pub fn checksum(&self) -> u32 {
        self.checksum
    }

    /// Layout mode of `x` (`wXMode`).
    pub fn x_mode(&self) -> XlsCrtLayout12Mode {
        self.modes.x_mode
    }

    /// Layout mode of `y` (`wYMode`).
    pub fn y_mode(&self) -> XlsCrtLayout12Mode {
        self.modes.y_mode
    }

    /// Layout mode of `dx` (`wWidthMode`).
    pub fn width_mode(&self) -> XlsCrtLayout12Mode {
        self.modes.width_mode
    }

    /// Layout mode of `dy` (`wHeightMode`).
    pub fn height_mode(&self) -> XlsCrtLayout12Mode {
        self.modes.height_mode
    }

    /// Horizontal offset (`x`), interpreted per [`Self::x_mode`].
    pub fn x(&self) -> f64 {
        self.modes.x
    }

    /// Vertical offset (`y`), interpreted per [`Self::y_mode`].
    pub fn y(&self) -> f64 {
        self.modes.y
    }

    /// Width or horizontal offset (`dx`), interpreted per [`Self::width_mode`].
    pub fn dx(&self) -> f64 {
        self.modes.dx
    }

    /// Height or vertical offset (`dy`), interpreted per [`Self::height_mode`].
    pub fn dy(&self) -> f64 {
        self.modes.dy
    }
}

/// Typed `CrtLayout12A` record content (MS-XLS 2.4.67): layout information
/// for a plot area.
///
/// The 15 `reserved1` bits and the trailing `reserved2` field (MUST be
/// ignored) are preserved verbatim so the record round-trips unchanged.
#[derive(Debug, Clone, PartialEq)]
pub struct XlsCrtLayout12A {
    /// Raw `frtHeader.grbitFrt` bitfield (`fFrtRef`/`fFrtAlert` are zero).
    frt_flags: u16,
    /// `frtHeader.reserved` bytes, preserved verbatim.
    frt_reserved: [u8; 8],
    /// `dwCheckSum`: 0x00000000 or 0x00000001 (MS-XLS 2.4.67).
    checksum: u32,
    /// Raw flags: `fLayoutTargetInner` and 15 `reserved1` bits.
    flags: u16,
    /// `xTL`: horizontal offset of the plot area's upper-left corner, in SPRC.
    x_top_left: i16,
    /// `yTL`: vertical offset of the plot area's upper-left corner, in SPRC.
    y_top_left: i16,
    /// `xBR`: width of the plot area, in SPRC.
    x_bottom_right: i16,
    /// `yBR`: height of the plot area, in SPRC.
    y_bottom_right: i16,
    modes: LayoutModes,
    /// Trailing `reserved2` field, preserved verbatim.
    reserved2: u16,
}

impl XlsCrtLayout12A {
    /// Parse a `CrtLayout12A` record payload.
    pub fn parse(data: &[u8]) -> XlsResult<Self> {
        if data.len() != CRT_LAYOUT_12_A_LEN {
            return Err(XlsError::InvalidLength {
                expected: CRT_LAYOUT_12_A_LEN,
                found: data.len(),
            });
        }
        let frt_flags = validate_frt_header(data, CRT_LAYOUT_12_A_RECORD_TYPE, "CrtLayout12A")?;
        let checksum = u32::from_le_bytes(data[12..16].try_into().expect("length checked"));
        // MS-XLS 2.4.67: dwCheckSum MUST be 0x00000000 or 0x00000001.
        if checksum > 1 {
            return Err(invalid(
                CRT_LAYOUT_12_A_RECORD_TYPE,
                format!("CrtLayout12A dwCheckSum {checksum:#X} is not 0x00000000 or 0x00000001"),
            ));
        }
        let i16_at = |index: usize| i16::from_le_bytes([data[index], data[index + 1]]);
        Ok(Self {
            frt_flags,
            frt_reserved: data[4..FRT_HEADER_LEN].try_into().expect("length checked"),
            checksum,
            flags: u16::from_le_bytes([data[16], data[17]]),
            x_top_left: i16_at(18),
            y_top_left: i16_at(20),
            x_bottom_right: i16_at(22),
            y_bottom_right: i16_at(24),
            modes: LayoutModes::parse(data, 26, CRT_LAYOUT_12_A_RECORD_TYPE)?,
            reserved2: u16::from_le_bytes([data[66], data[67]]),
        })
    }

    /// Serialize back to a complete `CrtLayout12A` record payload.
    pub fn to_payload(&self) -> Vec<u8> {
        let mut payload = Vec::with_capacity(CRT_LAYOUT_12_A_LEN);
        payload.extend_from_slice(&CRT_LAYOUT_12_A_RECORD_TYPE.to_le_bytes());
        payload.extend_from_slice(&self.frt_flags.to_le_bytes());
        payload.extend_from_slice(&self.frt_reserved);
        payload.extend_from_slice(&self.checksum().to_le_bytes());
        payload.extend_from_slice(&self.flags.to_le_bytes());
        for value in [
            self.x_top_left,
            self.y_top_left,
            self.x_bottom_right,
            self.y_bottom_right,
        ] {
            payload.extend_from_slice(&value.to_le_bytes());
        }
        self.modes.write_payload(&mut payload);
        payload.extend_from_slice(&self.reserved2.to_le_bytes());
        payload
    }

    /// The `dwCheckSum` value: 0x00000001 when the plot area layout is manual
    /// and not always automatically computed, 0x00000000 otherwise (derived
    /// from the `ShtProps` flags, MS-XLS 2.4.67).
    pub fn checksum(&self) -> u32 {
        self.checksum
    }

    /// Whether the layout target is the inner plot area (`fLayoutTargetInner`).
    pub fn is_layout_target_inner(&self) -> bool {
        self.flags & LAYOUT_TARGET_INNER != 0
    }

    /// Raw flags word, including the 15 `reserved1` bits.
    pub fn flags(&self) -> u16 {
        self.flags
    }

    /// Horizontal offset of the plot area's upper-left corner (`xTL`), in SPRC.
    pub fn x_top_left(&self) -> i16 {
        self.x_top_left
    }

    /// Vertical offset of the plot area's upper-left corner (`yTL`), in SPRC.
    pub fn y_top_left(&self) -> i16 {
        self.y_top_left
    }

    /// Width of the plot area (`xBR`), in SPRC.
    pub fn x_bottom_right(&self) -> i16 {
        self.x_bottom_right
    }

    /// Height of the plot area (`yBR`), in SPRC.
    pub fn y_bottom_right(&self) -> i16 {
        self.y_bottom_right
    }

    /// Layout mode of `x` (`wXMode`).
    pub fn x_mode(&self) -> XlsCrtLayout12Mode {
        self.modes.x_mode
    }

    /// Layout mode of `y` (`wYMode`).
    pub fn y_mode(&self) -> XlsCrtLayout12Mode {
        self.modes.y_mode
    }

    /// Layout mode of `dx` (`wWidthMode`).
    pub fn width_mode(&self) -> XlsCrtLayout12Mode {
        self.modes.width_mode
    }

    /// Layout mode of `dy` (`wHeightMode`).
    pub fn height_mode(&self) -> XlsCrtLayout12Mode {
        self.modes.height_mode
    }

    /// Horizontal offset (`x`), interpreted per [`Self::x_mode`].
    pub fn x(&self) -> f64 {
        self.modes.x
    }

    /// Vertical offset (`y`), interpreted per [`Self::y_mode`].
    pub fn y(&self) -> f64 {
        self.modes.y
    }

    /// Width or horizontal offset (`dx`), interpreted per [`Self::width_mode`].
    pub fn dx(&self) -> f64 {
        self.modes.dx
    }

    /// Height or vertical offset (`dy`), interpreted per [`Self::height_mode`].
    pub fn dy(&self) -> f64 {
        self.modes.dy
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn layout12_record(checksum: u32, flags: u16, modes: [u16; 4], values: [f64; 4]) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&CRT_LAYOUT_12_RECORD_TYPE.to_le_bytes());
        data.extend_from_slice(&[0; FRT_HEADER_LEN - 2]);
        data.extend_from_slice(&checksum.to_le_bytes());
        data.extend_from_slice(&flags.to_le_bytes());
        for mode in modes {
            data.extend_from_slice(&mode.to_le_bytes());
        }
        for value in values {
            data.extend_from_slice(&value.to_le_bytes());
        }
        data.extend_from_slice(&[0; 2]);
        data
    }

    fn layout12a_record(
        checksum: u32,
        flags: u16,
        corners: [i16; 4],
        modes: [u16; 4],
        values: [f64; 4],
    ) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&CRT_LAYOUT_12_A_RECORD_TYPE.to_le_bytes());
        data.extend_from_slice(&[0; FRT_HEADER_LEN - 2]);
        data.extend_from_slice(&checksum.to_le_bytes());
        data.extend_from_slice(&flags.to_le_bytes());
        for corner in corners {
            data.extend_from_slice(&corner.to_le_bytes());
        }
        for mode in modes {
            data.extend_from_slice(&mode.to_le_bytes());
        }
        for value in values {
            data.extend_from_slice(&value.to_le_bytes());
        }
        data.extend_from_slice(&[0; 2]);
        data
    }

    #[test]
    fn crt_layout12_round_trip() {
        let bytes = layout12_record(
            0x0000_4321,
            0x0008,
            [0x0001, 0x0001, 0x0002, 0x0002],
            [0.25, -0.5, 0.75, 1.0],
        );
        let parsed = XlsCrtLayout12::parse(&bytes).unwrap();
        assert_eq!(parsed.checksum(), 0x0000_4321);
        assert_eq!(parsed.auto_layout_type(), 0x4);
        assert_eq!(parsed.x_mode(), XlsCrtLayout12Mode::Factor);
        assert_eq!(parsed.width_mode(), XlsCrtLayout12Mode::Edge);
        assert_eq!(parsed.x(), 0.25);
        assert_eq!(parsed.y(), -0.5);
        assert_eq!(parsed.dx(), 0.75);
        assert_eq!(parsed.dy(), 1.0);
        assert_eq!(parsed.to_payload(), bytes);
    }

    #[test]
    fn crt_layout12_preserves_unused_and_reserved_bits() {
        // The unused bit, the 11 reserved1 bits, and reserved2 MUST be ignored
        // but round-trip verbatim.
        let mut bytes = layout12_record(0, 0xF801, [0; 4], [0.0; 4]);
        bytes[58..60].copy_from_slice(&0x7F7Fu16.to_le_bytes());
        bytes[4..FRT_HEADER_LEN].copy_from_slice(&[1, 2, 3, 4, 5, 6, 7, 8]);
        let parsed = XlsCrtLayout12::parse(&bytes).unwrap();
        assert_eq!(parsed.flags(), 0xF801);
        assert_eq!(parsed.auto_layout_type(), 0);
        assert_eq!(parsed.to_payload(), bytes);
    }

    #[test]
    fn crt_layout12_rejects_malformed_records() {
        let bytes = layout12_record(0, 0, [0; 4], [0.0; 4]);
        // Truncated and overlong payloads.
        assert!(XlsCrtLayout12::parse(&bytes[..59]).is_err());
        assert!(XlsCrtLayout12::parse(&[bytes.as_slice(), &[0]].concat()).is_err());
        // Wrong FrtHeader.rt.
        let mut wrong_rt = bytes.clone();
        wrong_rt[0..2].copy_from_slice(&0x089Eu16.to_le_bytes());
        assert!(XlsCrtLayout12::parse(&wrong_rt).is_err());
        // fFrtRef / fFrtAlert set.
        let mut bad_flags = bytes.clone();
        bad_flags[2..4].copy_from_slice(&0x0001u16.to_le_bytes());
        assert!(XlsCrtLayout12::parse(&bad_flags).is_err());
        // Undefined layout mode.
        assert!(XlsCrtLayout12::parse(&layout12_record(0, 0, [3, 0, 0, 0], [0.0; 4])).is_err());
    }

    #[test]
    fn crt_layout12a_round_trip() {
        for checksum in [0, 1] {
            let bytes = layout12a_record(
                checksum,
                0x0001,
                [100, -50, 4000, 3000],
                [0x0000, 0x0001, 0x0002, 0x0000],
                [0.1, 0.2, 0.3, 0.4],
            );
            let parsed = XlsCrtLayout12A::parse(&bytes).unwrap();
            assert_eq!(parsed.checksum(), checksum);
            assert!(parsed.is_layout_target_inner());
            assert_eq!(parsed.x_top_left(), 100);
            assert_eq!(parsed.y_top_left(), -50);
            assert_eq!(parsed.x_bottom_right(), 4000);
            assert_eq!(parsed.y_bottom_right(), 3000);
            assert_eq!(parsed.y_mode(), XlsCrtLayout12Mode::Factor);
            assert_eq!(parsed.width_mode(), XlsCrtLayout12Mode::Edge);
            assert_eq!(parsed.dx(), 0.3);
            assert_eq!(parsed.to_payload(), bytes);
        }
    }

    #[test]
    fn crt_layout12a_rejects_malformed_records() {
        let bytes = layout12a_record(0, 0, [0; 4], [0; 4], [0.0; 4]);
        // Truncated.
        assert!(XlsCrtLayout12A::parse(&bytes[..67]).is_err());
        // Wrong FrtHeader.rt.
        let mut wrong_rt = bytes.clone();
        wrong_rt[0..2].copy_from_slice(&0x089Du16.to_le_bytes());
        assert!(XlsCrtLayout12A::parse(&wrong_rt).is_err());
        // dwCheckSum outside 0x00000000..=0x00000001.
        assert!(XlsCrtLayout12A::parse(&layout12a_record(2, 0, [0; 4], [0; 4], [0.0; 4])).is_err());
        assert!(
            XlsCrtLayout12A::parse(&layout12a_record(0xFFFF_FFFF, 0, [0; 4], [0; 4], [0.0; 4]))
                .is_err()
        );
        // Undefined layout mode.
        assert!(
            XlsCrtLayout12A::parse(&layout12a_record(0, 0, [0; 4], [0, 9, 0, 0], [0.0; 4]))
                .is_err()
        );
    }
}
