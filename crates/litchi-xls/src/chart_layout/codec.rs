//! Bounded BIFF8 codecs for chart layout future records.

use crate::{Error, Result};

use super::model::{CrtLayout12, CrtLayout12A, CrtLayout12Mode, LayoutModes};

/// Record type of the `CrtLayout12` record (MS-XLS 2.4.66); also the required
/// `frtHeader.rt` value.
pub(super) const CRT_LAYOUT_12_RECORD_TYPE: u16 = 0x089D;

/// Record type of the `CrtLayout12A` record (MS-XLS 2.4.67); also the
/// required `frtHeader.rt` value.
pub(super) const CRT_LAYOUT_12_A_RECORD_TYPE: u16 = 0x08A7;

/// Size in bytes of an `FrtHeader` (MS-XLS 2.5.135).
pub(super) const FRT_HEADER_LEN: usize = 12;
/// Size in bytes of the reserved tail of an `FrtHeader`.
const FRT_HEADER_RESERVED_LEN: usize = FRT_HEADER_LEN - 4;
/// `FrtFlags` bits that MUST be zero in an `FrtHeader` (MS-XLS 2.5.135):
/// `fFrtRef` and `fFrtAlert`.
const FRT_FLAGS_FORBIDDEN: u16 = 0x0003;
/// Byte length of a `CrtLayout12` record payload.
const CRT_LAYOUT_12_LEN: usize = 60;
/// Byte length of a `CrtLayout12A` record payload.
const CRT_LAYOUT_12_A_LEN: usize = 68;
fn invalid(record_type: u16, message: impl Into<String>) -> Error {
    Error::InvalidRecord {
        record_type,
        message: message.into(),
    }
}

/// Bounded reader for the fixed-width fields in `CrtLayout12` and
/// `CrtLayout12A`.
///
/// The records have fixed layouts, but keeping the cursor checked here makes
/// the field readers safe if a caller or a future layout change ever bypasses
/// the public exact-length checks. The cursor advances only after a complete
/// field has been obtained.
pub(super) struct LayoutReader<'a> {
    pub(super) data: &'a [u8],
    pub(super) offset: usize,
    pub(super) record_type: u16,
}

impl<'a> LayoutReader<'a> {
    pub(super) fn new(data: &'a [u8], record_type: u16) -> Self {
        Self {
            data,
            offset: 0,
            record_type,
        }
    }

    fn read_bytes<const N: usize>(&mut self) -> Result<[u8; N]> {
        let end = self.offset.checked_add(N).ok_or_else(|| {
            invalid(
                self.record_type,
                "chart layout field offset overflows usize",
            )
        })?;
        let bytes = self
            .data
            .get(self.offset..end)
            .ok_or(Error::InvalidLength {
                expected: end,
                found: self.data.len(),
            })?;
        let value: [u8; N] = bytes.try_into().map_err(|_| Error::InvalidLength {
            expected: N,
            found: bytes.len(),
        })?;
        self.offset = end;
        Ok(value)
    }

    fn read_u16(&mut self) -> Result<u16> {
        Ok(u16::from_le_bytes(self.read_bytes()?))
    }

    fn read_i16(&mut self) -> Result<i16> {
        Ok(i16::from_le_bytes(self.read_bytes()?))
    }

    fn read_u32(&mut self) -> Result<u32> {
        Ok(u32::from_le_bytes(self.read_bytes()?))
    }

    fn read_f64(&mut self) -> Result<f64> {
        Ok(f64::from_le_bytes(self.read_bytes()?))
    }
}

impl CrtLayout12Mode {
    fn parse(value: u16, record_type: u16) -> Result<Self> {
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

impl LayoutModes {
    pub(super) fn parse(reader: &mut LayoutReader<'_>) -> Result<Self> {
        let mode = |reader: &mut LayoutReader<'_>| {
            CrtLayout12Mode::parse(reader.read_u16()?, reader.record_type)
        };
        Ok(Self {
            x_mode: mode(reader)?,
            y_mode: mode(reader)?,
            width_mode: mode(reader)?,
            height_mode: mode(reader)?,
            x: reader.read_f64()?,
            y: reader.read_f64()?,
            dx: reader.read_f64()?,
            dy: reader.read_f64()?,
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
fn validate_frt_header(reader: &mut LayoutReader<'_>, record_type: u16, name: &str) -> Result<u16> {
    if reader.read_u16()? != record_type {
        return Err(invalid(
            record_type,
            format!("{name} FrtHeader.rt mismatch"),
        ));
    }
    let flags = reader.read_u16()?;
    if flags & FRT_FLAGS_FORBIDDEN != 0 {
        return Err(invalid(
            record_type,
            format!("{name} FrtHeader.grbitFrt {flags:#06X} sets fFrtRef or fFrtAlert"),
        ));
    }
    Ok(flags)
}

impl CrtLayout12 {
    /// Parse a `CrtLayout12` record payload.
    pub fn parse(data: &[u8]) -> Result<Self> {
        if data.len() != CRT_LAYOUT_12_LEN {
            return Err(Error::InvalidLength {
                expected: CRT_LAYOUT_12_LEN,
                found: data.len(),
            });
        }
        let mut reader = LayoutReader::new(data, CRT_LAYOUT_12_RECORD_TYPE);
        let frt_flags = validate_frt_header(&mut reader, CRT_LAYOUT_12_RECORD_TYPE, "CrtLayout12")?;
        Ok(Self {
            frt_flags,
            frt_reserved: reader.read_bytes::<FRT_HEADER_RESERVED_LEN>()?,
            checksum: reader.read_u32()?,
            flags: reader.read_u16()?,
            modes: LayoutModes::parse(&mut reader)?,
            reserved2: reader.read_u16()?,
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
}

impl CrtLayout12A {
    /// Parse a `CrtLayout12A` record payload.
    pub fn parse(data: &[u8]) -> Result<Self> {
        if data.len() != CRT_LAYOUT_12_A_LEN {
            return Err(Error::InvalidLength {
                expected: CRT_LAYOUT_12_A_LEN,
                found: data.len(),
            });
        }
        let mut reader = LayoutReader::new(data, CRT_LAYOUT_12_A_RECORD_TYPE);
        let frt_flags =
            validate_frt_header(&mut reader, CRT_LAYOUT_12_A_RECORD_TYPE, "CrtLayout12A")?;
        let frt_reserved = reader.read_bytes::<FRT_HEADER_RESERVED_LEN>()?;
        let checksum = reader.read_u32()?;
        // MS-XLS 2.4.67: dwCheckSum MUST be 0x00000000 or 0x00000001.
        if checksum > 1 {
            return Err(invalid(
                CRT_LAYOUT_12_A_RECORD_TYPE,
                format!("CrtLayout12A dwCheckSum {checksum:#X} is not 0x00000000 or 0x00000001"),
            ));
        }
        Ok(Self {
            frt_flags,
            frt_reserved,
            checksum,
            flags: reader.read_u16()?,
            x_top_left: reader.read_i16()?,
            y_top_left: reader.read_i16()?,
            x_bottom_right: reader.read_i16()?,
            y_bottom_right: reader.read_i16()?,
            modes: LayoutModes::parse(&mut reader)?,
            reserved2: reader.read_u16()?,
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
}
