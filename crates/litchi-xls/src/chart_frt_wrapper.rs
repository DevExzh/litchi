//! BIFF8 chart future-record (FRT) wrappers of the Chart Sheet substream
//! (MS-XLS 2.1):
//!
//! - **StartObject** (0x0854): the beginning of a nested FRT object block
//!   (MS-XLS 2.4.267).
//! - **EndObject** (0x0855): the end of that block (MS-XLS 2.4.101).
//! - **FrtWrapper** (0x0851): wraps a non-FRT chart record inside an FRT
//!   record (MS-XLS 2.4.130).
//!
//! Everything in this module is INERT: wrapped and block contents are stored
//! verbatim and no chart feature is reconstructed. The
//! `StartObject`/`EndObject` pairing and nesting rules (up to 100 levels,
//! matching `iObjectKind` values) are cross-record constraints documented on
//! the types; single-record readers cannot enforce them.
//!
//! # References
//!
//! - MS-XLS 2.4.101 (EndObject), 2.4.130 (FrtWrapper), 2.4.267 (StartObject),
//!   2.5.134 (FrtFlags), 2.5.136 (FrtHeaderOld)

use super::{XlsError, XlsResult};

/// Record type of the `FrtWrapper` record (MS-XLS 2.4.130); also the required
/// `frtHeaderOld.rt` value.
pub(crate) const FRT_WRAPPER_RECORD_TYPE: u16 = 0x0851;

/// Record type of the `StartObject` record (MS-XLS 2.4.267); also the
/// required `frtHeaderOld.rt` value.
pub(crate) const START_OBJECT_RECORD_TYPE: u16 = 0x0854;

/// Record type of the `EndObject` record (MS-XLS 2.4.101); also the required
/// `frtHeaderOld.rt` value.
pub(crate) const END_OBJECT_RECORD_TYPE: u16 = 0x0855;

/// Size in bytes of an `FrtHeaderOld` (MS-XLS 2.5.136): `rt` + `grbitFrt`.
const FRT_HEADER_OLD_LEN: usize = 4;
/// `FrtFlags` bits that MUST be zero in an `FrtHeaderOld` (MS-XLS 2.5.136):
/// `fFrtRef` and `fFrtAlert`.
const FRT_FLAGS_FORBIDDEN: u16 = 0x0003;
/// Byte length of a `StartObject` record payload.
const START_OBJECT_LEN: usize = 12;
/// Byte length of an `EndObject` record payload.
const END_OBJECT_LEN: usize = 12;
/// Minimum size of the `wrappedRecord`/`frtWrapperPadding` region of an
/// `FrtWrapper` (MS-XLS 2.4.130).
const MIN_WRAPPED_LEN: usize = 8;

fn invalid(record_type: u16, message: impl Into<String>) -> XlsError {
    XlsError::InvalidRecord {
        record_type,
        message: message.into(),
    }
}

/// Validate the `rt` and `grbitFrt` fields of an `FrtHeaderOld`
/// (MS-XLS 2.5.136), returning the raw flags word.
fn validate_frt_header_old(data: &[u8], record_type: u16, name: &str) -> XlsResult<u16> {
    if u16::from_le_bytes([data[0], data[1]]) != record_type {
        return Err(invalid(
            record_type,
            format!("{name} FrtHeaderOld.rt mismatch"),
        ));
    }
    let flags = u16::from_le_bytes([data[2], data[3]]);
    if flags & FRT_FLAGS_FORBIDDEN != 0 {
        return Err(invalid(
            record_type,
            format!("{name} FrtHeaderOld.grbitFrt {flags:#06X} sets fFrtRef or fFrtAlert"),
        ));
    }
    Ok(flags)
}

/// The `iObjectKind` of a `StartObject`/`EndObject` FRT block (MS-XLS 2.4.267
/// / 2.4.101).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u16)]
pub enum XlsFrtObjectKind {
    /// 0x0010: a `YMult` (axis multiplier) block.
    YMult = 0x0010,
    /// 0x0011: an `FrtFontList` block.
    FrtFontList = 0x0011,
    /// 0x0012: a `DataLabExt` block.
    DataLabExt = 0x0012,
}

impl XlsFrtObjectKind {
    fn parse(value: u16, record_type: u16, name: &str) -> XlsResult<Self> {
        match value {
            0x0010 => Ok(Self::YMult),
            0x0011 => Ok(Self::FrtFontList),
            0x0012 => Ok(Self::DataLabExt),
            other => Err(invalid(
                record_type,
                format!("{name} iObjectKind {other:#06X} is not a defined object kind"),
            )),
        }
    }
}

/// Typed `StartObject` record content (MS-XLS 2.4.267): the beginning of a
/// nested FRT object block.
///
/// The 14 reserved `grbitFrt` bits are preserved verbatim so the record
/// round-trips unchanged.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XlsStartObject {
    /// Raw `frtHeaderOld.grbitFrt` bitfield (`fFrtRef`/`fFrtAlert` are zero).
    frt_flags: u16,
    /// The kind of object encompassed by the block (`iObjectKind`).
    kind: XlsFrtObjectKind,
    /// `iObjectInstance1`: additional context. Guaranteed zero for `YMult` and
    /// `DataLabExt`; an application version (0x0008..=0x000F except 0x000D)
    /// for `FrtFontList`.
    object_instance1: u16,
}

impl XlsStartObject {
    /// Parse a `StartObject` record payload.
    pub fn parse(data: &[u8]) -> XlsResult<Self> {
        if data.len() != START_OBJECT_LEN {
            return Err(XlsError::InvalidLength {
                expected: START_OBJECT_LEN,
                found: data.len(),
            });
        }
        let frt_flags = validate_frt_header_old(data, START_OBJECT_RECORD_TYPE, "StartObject")?;
        let kind = XlsFrtObjectKind::parse(
            u16::from_le_bytes([data[4], data[5]]),
            START_OBJECT_RECORD_TYPE,
            "StartObject",
        )?;
        // MS-XLS 2.4.267: iObjectContext MUST be 0x0000.
        let context = u16::from_le_bytes([data[6], data[7]]);
        if context != 0 {
            return Err(invalid(
                START_OBJECT_RECORD_TYPE,
                format!("StartObject iObjectContext {context:#06X} is not 0x0000"),
            ));
        }
        let object_instance1 = u16::from_le_bytes([data[8], data[9]]);
        match kind {
            // MS-XLS 2.4.267: MUST equal 0x0000 for YMult and DataLabExt.
            XlsFrtObjectKind::YMult | XlsFrtObjectKind::DataLabExt if object_instance1 != 0 => {
                return Err(invalid(
                    START_OBJECT_RECORD_TYPE,
                    format!("StartObject iObjectInstance1 {object_instance1:#06X} is not 0x0000"),
                ));
            },
            // MS-XLS 2.4.267: an application version for FrtFontList.
            XlsFrtObjectKind::FrtFontList
                if !matches!(object_instance1, 0x0008..=0x000C | 0x000E | 0x000F) =>
            {
                return Err(invalid(
                    START_OBJECT_RECORD_TYPE,
                    format!(
                        "StartObject iObjectInstance1 {object_instance1:#06X} is not an application version"
                    ),
                ));
            },
            _ => {},
        }
        // MS-XLS 2.4.267: iObjectInstance2 MUST equal 0x0000.
        let object_instance2 = u16::from_le_bytes([data[10], data[11]]);
        if object_instance2 != 0 {
            return Err(invalid(
                START_OBJECT_RECORD_TYPE,
                format!("StartObject iObjectInstance2 {object_instance2:#06X} is not 0x0000"),
            ));
        }
        Ok(Self {
            frt_flags,
            kind,
            object_instance1,
        })
    }

    /// Serialize back to a complete `StartObject` record payload.
    pub fn to_payload(&self) -> Vec<u8> {
        let mut payload = Vec::with_capacity(START_OBJECT_LEN);
        payload.extend_from_slice(&START_OBJECT_RECORD_TYPE.to_le_bytes());
        payload.extend_from_slice(&self.frt_flags.to_le_bytes());
        payload.extend_from_slice(&(self.kind as u16).to_le_bytes());
        payload.extend_from_slice(&0u16.to_le_bytes());
        payload.extend_from_slice(&self.object_instance1.to_le_bytes());
        payload.extend_from_slice(&0u16.to_le_bytes());
        payload
    }

    /// The kind of object encompassed by the block (`iObjectKind`).
    pub fn kind(&self) -> XlsFrtObjectKind {
        self.kind
    }

    /// Additional object context (`iObjectInstance1`); an application version
    /// for `FrtFontList` blocks, zero otherwise.
    pub fn object_instance1(&self) -> u16 {
        self.object_instance1
    }

    /// Raw `frtHeaderOld.grbitFrt` bitfield (`fFrtRef`/`fFrtAlert` are zero).
    pub fn frt_flags(&self) -> u16 {
        self.frt_flags
    }
}

/// Typed `EndObject` record content (MS-XLS 2.4.101): the end of a nested FRT
/// object block. Its `iObjectKind` MUST equal the associated `StartObject`
/// record's value (a cross-record constraint the caller validates).
///
/// The three `unused` fields (MUST be ignored) are preserved verbatim so the
/// record round-trips unchanged.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XlsEndObject {
    /// Raw `frtHeaderOld.grbitFrt` bitfield (`fFrtRef`/`fFrtAlert` are zero).
    frt_flags: u16,
    /// The kind of object encompassed by the block (`iObjectKind`).
    kind: XlsFrtObjectKind,
    /// The `unused1`, `unused2`, and `unused3` fields, preserved verbatim.
    unused: [u16; 3],
}

impl XlsEndObject {
    /// Parse an `EndObject` record payload.
    pub fn parse(data: &[u8]) -> XlsResult<Self> {
        if data.len() != END_OBJECT_LEN {
            return Err(XlsError::InvalidLength {
                expected: END_OBJECT_LEN,
                found: data.len(),
            });
        }
        let frt_flags = validate_frt_header_old(data, END_OBJECT_RECORD_TYPE, "EndObject")?;
        let kind = XlsFrtObjectKind::parse(
            u16::from_le_bytes([data[4], data[5]]),
            END_OBJECT_RECORD_TYPE,
            "EndObject",
        )?;
        Ok(Self {
            frt_flags,
            kind,
            unused: [
                u16::from_le_bytes([data[6], data[7]]),
                u16::from_le_bytes([data[8], data[9]]),
                u16::from_le_bytes([data[10], data[11]]),
            ],
        })
    }

    /// Serialize back to a complete `EndObject` record payload.
    pub fn to_payload(&self) -> Vec<u8> {
        let mut payload = Vec::with_capacity(END_OBJECT_LEN);
        payload.extend_from_slice(&END_OBJECT_RECORD_TYPE.to_le_bytes());
        payload.extend_from_slice(&self.frt_flags.to_le_bytes());
        payload.extend_from_slice(&(self.kind as u16).to_le_bytes());
        for value in self.unused {
            payload.extend_from_slice(&value.to_le_bytes());
        }
        payload
    }

    /// The kind of object encompassed by the block (`iObjectKind`).
    pub fn kind(&self) -> XlsFrtObjectKind {
        self.kind
    }

    /// The preserved `unused1`/`unused2`/`unused3` fields.
    pub fn unused(&self) -> [u16; 3] {
        self.unused
    }

    /// Raw `frtHeaderOld.grbitFrt` bitfield (`fFrtRef`/`fFrtAlert` are zero).
    pub fn frt_flags(&self) -> u16 {
        self.frt_flags
    }
}

/// Typed `FrtWrapper` record content (MS-XLS 2.4.130): a non-FRT chart record
/// wrapped inside a future record.
///
/// The `wrappedRecord` bytes (a complete BIFF record) and the zero
/// `frtWrapperPadding` bytes are stored verbatim; the wrapped record is never
/// interpreted.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsFrtWrapper {
    /// Raw `frtHeaderOld.grbitFrt` bitfield (`fFrtRef`/`fFrtAlert` are zero).
    frt_flags: u16,
    /// The complete wrapped BIFF record bytes (`wrappedRecord`).
    wrapped: Vec<u8>,
    /// The zero padding bytes (`frtWrapperPadding`), preserved verbatim.
    padding: Vec<u8>,
}

impl XlsFrtWrapper {
    /// Parse an `FrtWrapper` record payload.
    pub fn parse(data: &[u8]) -> XlsResult<Self> {
        // MS-XLS 2.4.130: the padded FrtWrapper is never smaller than the
        // 12-byte FrtHeader structure.
        if data.len() < FRT_HEADER_OLD_LEN + MIN_WRAPPED_LEN {
            return Err(XlsError::InvalidLength {
                expected: FRT_HEADER_OLD_LEN + MIN_WRAPPED_LEN,
                found: data.len(),
            });
        }
        let frt_flags = validate_frt_header_old(data, FRT_WRAPPER_RECORD_TYPE, "FrtWrapper")?;
        // The wrapped record is a complete BIFF record: a 4-byte header
        // (record type + payload length) plus its payload.
        let wrapped_len = FRT_HEADER_OLD_LEN + usize::from(u16::from_le_bytes([data[6], data[7]]));
        let expected = if wrapped_len < MIN_WRAPPED_LEN {
            // frtWrapperPadding pads the region to 8 bytes.
            FRT_HEADER_OLD_LEN + MIN_WRAPPED_LEN
        } else {
            FRT_HEADER_OLD_LEN + wrapped_len
        };
        if data.len() != expected {
            return Err(invalid(
                FRT_WRAPPER_RECORD_TYPE,
                format!(
                    "FrtWrapper wrappedRecord size requires {expected} bytes, found {}",
                    data.len()
                ),
            ));
        }
        let wrapped = data[FRT_HEADER_OLD_LEN..FRT_HEADER_OLD_LEN + wrapped_len].to_vec();
        let padding = data[FRT_HEADER_OLD_LEN + wrapped_len..].to_vec();
        // MS-XLS 2.4.130: padding elements MUST be zero and MUST be ignored;
        // they are preserved verbatim for the round trip.
        Ok(Self {
            frt_flags,
            wrapped,
            padding,
        })
    }

    /// Serialize back to a complete `FrtWrapper` record payload.
    pub fn to_payload(&self) -> Vec<u8> {
        let mut payload =
            Vec::with_capacity(FRT_HEADER_OLD_LEN + self.wrapped.len() + self.padding.len());
        payload.extend_from_slice(&FRT_WRAPPER_RECORD_TYPE.to_le_bytes());
        payload.extend_from_slice(&self.frt_flags.to_le_bytes());
        payload.extend_from_slice(&self.wrapped);
        payload.extend_from_slice(&self.padding);
        payload
    }

    /// The complete wrapped BIFF record bytes (`wrappedRecord`).
    pub fn wrapped_record(&self) -> &[u8] {
        &self.wrapped
    }

    /// The record type identifier of the wrapped record.
    pub fn wrapped_record_type(&self) -> u16 {
        u16::from_le_bytes([self.wrapped[0], self.wrapped[1]])
    }

    /// The preserved `frtWrapperPadding` bytes.
    pub fn padding(&self) -> &[u8] {
        &self.padding
    }

    /// Raw `frtHeaderOld.grbitFrt` bitfield (`fFrtRef`/`fFrtAlert` are zero).
    pub fn frt_flags(&self) -> u16 {
        self.frt_flags
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn start_object(kind: u16, instance1: u16) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&START_OBJECT_RECORD_TYPE.to_le_bytes());
        data.extend_from_slice(&[0; 2]);
        data.extend_from_slice(&kind.to_le_bytes());
        data.extend_from_slice(&[0; 2]);
        data.extend_from_slice(&instance1.to_le_bytes());
        data.extend_from_slice(&[0; 2]);
        data
    }

    fn end_object(kind: u16, unused: [u16; 3]) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&END_OBJECT_RECORD_TYPE.to_le_bytes());
        data.extend_from_slice(&[0; 2]);
        data.extend_from_slice(&kind.to_le_bytes());
        for value in unused {
            data.extend_from_slice(&value.to_le_bytes());
        }
        data
    }

    #[test]
    fn start_object_round_trip_all_kinds() {
        for (kind, instance1, expected) in [
            (0x0010, 0, XlsFrtObjectKind::YMult),
            (0x0011, 0x000E, XlsFrtObjectKind::FrtFontList),
            (0x0012, 0, XlsFrtObjectKind::DataLabExt),
        ] {
            let bytes = start_object(kind, instance1);
            let parsed = XlsStartObject::parse(&bytes).unwrap();
            assert_eq!(parsed.kind(), expected);
            assert_eq!(parsed.object_instance1(), instance1);
            assert_eq!(parsed.to_payload(), bytes);
        }
    }

    #[test]
    fn start_object_rejects_malformed_records() {
        let bytes = start_object(0x0010, 0);
        // Truncated.
        assert!(XlsStartObject::parse(&bytes[..11]).is_err());
        // Wrong FrtHeaderOld.rt.
        let mut wrong_rt = bytes.clone();
        wrong_rt[0..2].copy_from_slice(&0x0855u16.to_le_bytes());
        assert!(XlsStartObject::parse(&wrong_rt).is_err());
        // fFrtRef set.
        let mut bad_flags = bytes.clone();
        bad_flags[2..4].copy_from_slice(&0x0001u16.to_le_bytes());
        assert!(XlsStartObject::parse(&bad_flags).is_err());
        // Undefined iObjectKind.
        assert!(XlsStartObject::parse(&start_object(0x0013, 0)).is_err());
        // Nonzero iObjectContext.
        let mut bad_context = bytes.clone();
        bad_context[6..8].copy_from_slice(&1u16.to_le_bytes());
        assert!(XlsStartObject::parse(&bad_context).is_err());
        // iObjectInstance1 rules per kind.
        assert!(XlsStartObject::parse(&start_object(0x0010, 1)).is_err());
        assert!(XlsStartObject::parse(&start_object(0x0012, 1)).is_err());
        assert!(XlsStartObject::parse(&start_object(0x0011, 0)).is_err());
        assert!(XlsStartObject::parse(&start_object(0x0011, 0x000D)).is_err());
        assert!(XlsStartObject::parse(&start_object(0x0011, 0x0010)).is_err());
        // Nonzero iObjectInstance2.
        let mut bad_instance2 = bytes.clone();
        bad_instance2[10..12].copy_from_slice(&1u16.to_le_bytes());
        assert!(XlsStartObject::parse(&bad_instance2).is_err());
    }

    #[test]
    fn end_object_round_trip_and_rejects() {
        let bytes = end_object(0x0011, [1, 2, 3]);
        let parsed = XlsEndObject::parse(&bytes).unwrap();
        assert_eq!(parsed.kind(), XlsFrtObjectKind::FrtFontList);
        assert_eq!(parsed.unused(), [1, 2, 3]);
        assert_eq!(parsed.to_payload(), bytes);

        assert!(XlsEndObject::parse(&bytes[..9]).is_err());
        assert!(XlsEndObject::parse(&end_object(0x0000, [0; 3])).is_err());
        let mut wrong_rt = bytes.clone();
        wrong_rt[0..2].copy_from_slice(&0x0854u16.to_le_bytes());
        assert!(XlsEndObject::parse(&wrong_rt).is_err());
    }

    #[test]
    fn frt_wrapper_round_trip_with_and_without_padding() {
        // A wrapped record of 12 bytes: no padding.
        let mut wrapped = 0x101Cu16.to_le_bytes().to_vec();
        wrapped.extend_from_slice(&8u16.to_le_bytes());
        wrapped.extend_from_slice(&[0; 8]);
        let mut bytes = FRT_WRAPPER_RECORD_TYPE.to_le_bytes().to_vec();
        bytes.extend_from_slice(&[0; 2]);
        bytes.extend_from_slice(&wrapped);
        let parsed = XlsFrtWrapper::parse(&bytes).unwrap();
        assert_eq!(parsed.wrapped_record(), wrapped.as_slice());
        assert_eq!(parsed.wrapped_record_type(), 0x101C);
        assert!(parsed.padding().is_empty());
        assert_eq!(parsed.to_payload(), bytes);

        // A wrapped record of 6 bytes: padded to 8.
        let mut small = 0x0018u16.to_le_bytes().to_vec();
        small.extend_from_slice(&2u16.to_le_bytes());
        small.extend_from_slice(&[0x41, 0x00]);
        let mut padded = FRT_WRAPPER_RECORD_TYPE.to_le_bytes().to_vec();
        padded.extend_from_slice(&[0; 2]);
        padded.extend_from_slice(&small);
        padded.extend_from_slice(&[0; 2]);
        let parsed = XlsFrtWrapper::parse(&padded).unwrap();
        assert_eq!(parsed.wrapped_record(), small.as_slice());
        assert_eq!(parsed.padding(), &[0, 0]);
        assert_eq!(parsed.to_payload(), padded);
    }

    #[test]
    fn frt_wrapper_rejects_malformed_records() {
        let mut wrapped = 0x101Cu16.to_le_bytes().to_vec();
        wrapped.extend_from_slice(&8u16.to_le_bytes());
        wrapped.extend_from_slice(&[0; 8]);
        let mut bytes = FRT_WRAPPER_RECORD_TYPE.to_le_bytes().to_vec();
        bytes.extend_from_slice(&[0; 2]);
        bytes.extend_from_slice(&wrapped);
        // Below the 12-byte minimum.
        let mut tiny = FRT_WRAPPER_RECORD_TYPE.to_le_bytes().to_vec();
        tiny.extend_from_slice(&[0; 7]);
        assert!(XlsFrtWrapper::parse(&tiny).is_err());
        // Wrong FrtHeaderOld.rt.
        let mut wrong_rt = bytes.clone();
        wrong_rt[0..2].copy_from_slice(&0x0854u16.to_le_bytes());
        assert!(XlsFrtWrapper::parse(&wrong_rt).is_err());
        // Trailing bytes beyond the wrapped record.
        let mut trailing = bytes.clone();
        trailing.push(0);
        assert!(XlsFrtWrapper::parse(&trailing).is_err());
        // Declared wrapped size does not fit.
        let mut oversize = FRT_WRAPPER_RECORD_TYPE.to_le_bytes().to_vec();
        oversize.extend_from_slice(&[0; 2]);
        oversize.extend_from_slice(&0x101Cu16.to_le_bytes());
        oversize.extend_from_slice(&0xFFFFu16.to_le_bytes());
        oversize.extend_from_slice(&[0; 4]);
        assert!(XlsFrtWrapper::parse(&oversize).is_err());
    }
}
