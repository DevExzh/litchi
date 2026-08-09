//! BIFF8 shared-workbook user records and routing-slip records of the
//! Workbook stream (MS-XLS 2.1):
//!
//! - **CUsr** (0x0191): count of unique users that have the shared workbook
//!   open (MS-XLS 2.4.72).
//! - **CbUsr** (0x018A): byte count of each `UsrInfo` record stored in the
//!   user names stream of a shared workbook (MS-XLS 2.4.40).
//! - **UsrInfo** (0x0192): one user of the shared workbook, from the user
//!   names stream (MS-XLS 2.4.340).
//! - **DocRoute** (0x00B8): routing-slip delivery options and strings of an
//!   e-mail routed document (MS-XLS 2.4.91).
//! - **RecipName** (0x00B9): one recipient of a routing slip (MS-XLS 2.4.216).
//!
//! Everything in this module is INERT: fields are stored verbatim, no mail is
//! sent and no shared-workbook state is applied. The cross-record constraint
//! that `CbUsr` elements at an index greater than or equal to `CUsr.iCount`
//! MUST be zero (MS-XLS 2.4.40) is a property of the record sequence and is
//! documented on [`CbUsr`] rather than enforced by these payload readers.
//!
//! # References
//!
//! - MS-XLS 2.4.40 (CbUsr), 2.4.72 (CUsr), 2.4.91 (DocRoute),
//!   2.4.216 (RecipName), 2.4.340 (UsrInfo), 2.5.239 (ShortDTR),
//!   2.5.294 (XLUnicodeString)

use super::revision_records::ShortDtr;
use super::{Error, Result};

/// Record type of the `CUsr` record (MS-XLS 2.4.72).
pub(crate) const C_USR_RECORD_TYPE: u16 = 0x0191;
/// Record type of the `CbUsr` record (MS-XLS 2.4.40).
pub(crate) const CB_USR_RECORD_TYPE: u16 = 0x018A;
/// Record type of the `UsrInfo` record (MS-XLS 2.4.340).
pub(crate) const USR_INFO_RECORD_TYPE: u16 = 0x0192;
/// Record type of the `DocRoute` record (MS-XLS 2.4.91).
pub(crate) const DOC_ROUTE_RECORD_TYPE: u16 = 0x00B8;
/// Record type of the `RecipName` record (MS-XLS 2.4.216).
pub(crate) const RECIP_NAME_RECORD_TYPE: u16 = 0x00B9;

/// Byte length of a `CUsr` record payload (MS-XLS 2.4.72).
const C_USR_LEN: usize = 2;
/// Byte length of a `CbUsr` record payload: 256 two-byte counts
/// (MS-XLS 2.4.40).
const CB_USR_LEN: usize = 512;
/// Number of `UsrInfo` byte counts in a `CbUsr` record.
const CB_USR_COUNT: usize = CB_USR_LEN / 2;
/// Highest legal `CUsr.iCount` value (MS-XLS 2.4.72).
const C_USR_MAX_COUNT: u16 = 255;

/// Byte length of the fixed `UsrInfo` prefix: `lUsrId` (4) + `guid` (16) +
/// `shortdtr` (8) (MS-XLS 2.4.340).
const USR_INFO_PREFIX_LEN: usize = 4 + 16 + 8;
/// Byte length of a GUID as stored in `UsrInfo.guid` (MS-DTYP 2.3.4).
const GUID_LEN: usize = 16;
/// Byte length of a `ShortDTR` structure (MS-XLS 2.5.239).
const SHORT_DTR_LEN: usize = 8;
/// Header of an `XLUnicodeString`: `cch` (2) + option flags (1)
/// (MS-XLS 2.5.294).
const XL_UNICODE_STRING_HEADER_LEN: usize = 3;
/// Minimum character count of `UsrInfo.stUserName` (MS-XLS 2.4.340).
const USR_INFO_MIN_NAME_CHARS: usize = 1;
/// Maximum character count of `UsrInfo.stUserName` (MS-XLS 2.4.340).
const USR_INFO_MAX_NAME_CHARS: usize = 54;
/// `fHighByte` option bit of an `XLUnicodeString` (MS-XLS 2.5.294).
const STRING_HIGH_BYTE: u8 = 0x01;
/// Option bits of an `XLUnicodeString` that select rich-text runs or an
/// extended string (`fRichSt`/`fExtSt`); neither can appear in the
/// fixed-layout `UsrInfo` record (MS-XLS 2.5.294).
const STRING_UNSUPPORTED_FLAGS: u8 = 0x0C;

/// Byte length of the fixed `DocRoute` header: ten 2-byte fields plus the
/// 4-byte `ulEIDSize` (MS-XLS 2.4.91).
const DOC_ROUTE_HEADER_LEN: usize = 24;
/// Highest legal character count of any counted `DocRoute` string
/// (MS-XLS 2.4.91).
const DOC_ROUTE_MAX_STRING_CHARS: u16 = 256;
/// Highest legal total of the `DocRoute` character counts and `ulEIDSize`
/// (MS-XLS 2.4.91).
const DOC_ROUTE_MAX_TOTAL_CHARS: u64 = 8202;
/// `DocRoute` delivery option: one recipient at a time (MS-XLS 2.4.91).
const DELIVER_ONE_AT_A_TIME: u16 = 0x0000;
/// `DocRoute` delivery option: all recipients at once (MS-XLS 2.4.91).
const DELIVER_ALL_AT_ONCE: u16 = 0x0001;
/// `DocRoute` bitfield: `fRouted` (MS-XLS 2.4.91).
const ROUTE_FLAG_ROUTED: u16 = 0x0001;
/// `DocRoute` bitfield: `fReturnOrig` (MS-XLS 2.4.91).
const ROUTE_FLAG_RETURN_ORIG: u16 = 0x0002;
/// `DocRoute` bitfield: `fTrackStatus` (MS-XLS 2.4.91).
const ROUTE_FLAG_TRACK_STATUS: u16 = 0x0004;
/// `DocRoute` bitfield: `fCustomType` (MS-XLS 2.4.91).
const ROUTE_FLAG_CUSTOM_TYPE: u16 = 0x0008;
/// `DocRoute` bitfield: `fSaveRouteInfo`; MUST be set (MS-XLS 2.4.91).
const ROUTE_FLAG_SAVE_ROUTE_INFO: u16 = 0x0080;

/// Byte length of the fixed `RecipName` header: `cchRecip` (2) +
/// `ulEIDSize` (4) (MS-XLS 2.4.216).
const RECIP_NAME_HEADER_LEN: usize = 6;
/// Highest legal character count of `RecipName.szFriendly` (MS-XLS 2.4.216).
const RECIP_NAME_MAX_FRIENDLY_CHARS: u16 = 256;

/// Read a little-endian `u16` from a fixed offset (length checked by caller).
fn read_u16(data: &[u8], offset: usize) -> u16 {
    u16::from_le_bytes(data[offset..offset + 2].try_into().expect("length checked"))
}

/// Read a little-endian `u32` from a fixed offset (length checked by caller).
fn read_u32(data: &[u8], offset: usize) -> u32 {
    u32::from_le_bytes(data[offset..offset + 4].try_into().expect("length checked"))
}

fn invalid(record_type: u16, message: impl Into<String>) -> Error {
    Error::InvalidRecord {
        record_type,
        message: message.into(),
    }
}

/// Decode a counted, NULL-terminated ANSI string field, stripping the
/// terminator. `count` is the on-disk byte length of the field including the
/// NULL; a zero count means the field is absent (MS-XLS 2.4.91).
fn read_nul_terminated_ansi(
    record_type: u16,
    data: &[u8],
    offset: &mut usize,
    count: usize,
    field: &'static str,
) -> Result<Vec<u8>> {
    if count == 0 {
        return Ok(Vec::new());
    }
    let end = offset
        .checked_add(count)
        .ok_or_else(|| invalid(record_type, format!("{field} length overflow")))?;
    if end > data.len() {
        return Err(invalid(
            record_type,
            format!("{field} extends past the end of the record"),
        ));
    }
    let field_bytes = &data[*offset..end];
    if *field_bytes.last().expect("count is non-zero") != 0 {
        return Err(invalid(
            record_type,
            format!("{field} is not NULL terminated"),
        ));
    }
    *offset = end;
    Ok(field_bytes[..field_bytes.len() - 1].to_vec())
}

/// Append a NULL-terminated ANSI string field, or nothing when empty.
fn push_nul_terminated_ansi(payload: &mut Vec<u8>, content: &[u8]) {
    if content.is_empty() {
        return;
    }
    payload.extend_from_slice(content);
    payload.push(0);
}

/// On-disk byte count of a NULL-terminated ANSI string field (0 when absent).
fn nul_terminated_len(content: &[u8]) -> usize {
    if content.is_empty() {
        0
    } else {
        content.len() + 1
    }
}

/// Typed `CUsr` record content (MS-XLS 2.4.72): the number of unique users
/// that have the shared workbook open.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct CUsr {
    /// Number of unique users (`iCount`).
    count: u16,
}

impl CUsr {
    /// Parse a `CUsr` record payload.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn parse(data: &[u8]) -> Result<Self> {
        if data.len() != C_USR_LEN {
            return Err(Error::InvalidLength {
                expected: C_USR_LEN,
                found: data.len(),
            });
        }
        let count = read_u16(data, 0);
        if count > C_USR_MAX_COUNT {
            return Err(invalid(
                C_USR_RECORD_TYPE,
                format!("CUsr iCount {count} exceeds {C_USR_MAX_COUNT}"),
            ));
        }
        Ok(Self { count })
    }

    /// Serialize back to a complete `CUsr` record payload.
    #[must_use]
    pub fn to_payload(&self) -> Vec<u8> {
        self.count.to_le_bytes().to_vec()
    }

    /// Number of unique users that have the shared workbook open.
    #[must_use]
    pub fn count(&self) -> u16 {
        self.count
    }
}

/// Typed `CbUsr` record content (MS-XLS 2.4.40): the byte count of each
/// `UsrInfo` record stored in the user names stream of a shared workbook.
///
/// Elements whose index is greater than or equal to the `iCount` field of
/// the preceding `CUsr` record MUST be zero and MUST be ignored; that
/// constraint spans records and is not enforced by this payload reader.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CbUsr {
    /// Byte counts of the `UsrInfo` records (`rgCbUsr`).
    sizes: [u16; CB_USR_COUNT],
}

impl CbUsr {
    /// Parse a `CbUsr` record payload.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn parse(data: &[u8]) -> Result<Self> {
        if data.len() != CB_USR_LEN {
            return Err(invalid(
                CB_USR_RECORD_TYPE,
                format!("CbUsr has {} bytes; expected {CB_USR_LEN}", data.len()),
            ));
        }
        let mut sizes = [0u16; CB_USR_COUNT];
        for (index, size) in sizes.iter_mut().enumerate() {
            *size = read_u16(data, index * 2);
        }
        Ok(Self { sizes })
    }

    /// Serialize back to a complete `CbUsr` record payload.
    #[must_use]
    pub fn to_payload(&self) -> Vec<u8> {
        let mut payload = Vec::with_capacity(CB_USR_LEN);
        for size in &self.sizes {
            payload.extend_from_slice(&size.to_le_bytes());
        }
        payload
    }

    /// Byte counts of the `UsrInfo` records, in user order.
    #[must_use]
    pub fn sizes(&self) -> &[u16; CB_USR_COUNT] {
        &self.sizes
    }
}

/// Typed `UsrInfo` record content (MS-XLS 2.4.340): information about a user
/// who currently has the shared workbook open.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct UsrInfo {
    /// Unique user identifier (`lUsrId`).
    user_id: i32,
    /// Last set of revisions synced to by this user (`guid`).
    guid: [u8; GUID_LEN],
    /// Date and time the user opened the shared workbook (`shortdtr`).
    opened_at: ShortDtr,
    /// Name of this user (`stUserName`).
    user_name: String,
    /// Undefined trailing byte (`unused`), preserved verbatim.
    unused: u8,
}

impl UsrInfo {
    /// Parse a `UsrInfo` record payload.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    /// # Panics
    ///
    /// Panics only if an internal BIFF invariant has been violated.
    pub fn parse(data: &[u8]) -> Result<Self> {
        if data.len() < USR_INFO_PREFIX_LEN + XL_UNICODE_STRING_HEADER_LEN + 1 {
            return Err(Error::InvalidLength {
                expected: USR_INFO_PREFIX_LEN + XL_UNICODE_STRING_HEADER_LEN + 1,
                found: data.len(),
            });
        }
        let user_id = i32::from_le_bytes(data[0..4].try_into().expect("length checked"));
        let mut guid = [0u8; GUID_LEN];
        guid.copy_from_slice(&data[4..4 + GUID_LEN]);
        let opened_at = ShortDtr::parse(USR_INFO_RECORD_TYPE, &data[20..20 + SHORT_DTR_LEN])?;

        // stUserName: XLUnicodeString (MS-XLS 2.5.294).
        let cch = read_u16(data, USR_INFO_PREFIX_LEN) as usize;
        if !(USR_INFO_MIN_NAME_CHARS..=USR_INFO_MAX_NAME_CHARS).contains(&cch) {
            return Err(invalid(
                USR_INFO_RECORD_TYPE,
                format!(
                    "UsrInfo stUserName has {cch} characters; expected \
                     {USR_INFO_MIN_NAME_CHARS}..={USR_INFO_MAX_NAME_CHARS}"
                ),
            ));
        }
        let flags = data[USR_INFO_PREFIX_LEN + 2];
        if flags & STRING_UNSUPPORTED_FLAGS != 0 {
            return Err(invalid(
                USR_INFO_RECORD_TYPE,
                "UsrInfo stUserName cannot carry rich-text or extended-string data",
            ));
        }
        let high_byte = flags & STRING_HIGH_BYTE != 0;
        let char_bytes = cch * if high_byte { 2 } else { 1 };
        let chars_offset = USR_INFO_PREFIX_LEN + XL_UNICODE_STRING_HEADER_LEN;
        let expected_len = chars_offset + char_bytes + 1;
        if data.len() != expected_len {
            return Err(Error::InvalidLength {
                expected: expected_len,
                found: data.len(),
            });
        }
        let raw = &data[chars_offset..chars_offset + char_bytes];
        let user_name = if high_byte {
            let units: Vec<u16> = raw
                .chunks_exact(2)
                .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
                .collect();
            String::from_utf16(&units)
                .map_err(|error| Error::Encoding(format!("UsrInfo stUserName: {error}")))?
        } else {
            // Compressed Unicode supplies an implicit zero high byte.
            raw.iter().map(|&byte| byte as char).collect()
        };
        let unused = data[chars_offset + char_bytes];
        Ok(Self {
            user_id,
            guid,
            opened_at,
            user_name,
            unused,
        })
    }

    /// Serialize back to a complete `UsrInfo` record payload.
    #[must_use]
    pub fn to_payload(&self) -> Vec<u8> {
        let mut payload = Vec::with_capacity(USR_INFO_PREFIX_LEN + XL_UNICODE_STRING_HEADER_LEN);
        payload.extend_from_slice(&self.user_id.to_le_bytes());
        payload.extend_from_slice(&self.guid);
        let dtr = &self.opened_at;
        payload.extend_from_slice(&dtr.year().to_le_bytes());
        payload.push(dtr.month());
        payload.push(dtr.day());
        payload.push(dtr.hour());
        payload.push(dtr.minute());
        payload.push(dtr.second());
        payload.push(dtr.weekday());
        let high_byte = self.user_name.chars().any(|ch| u32::from(ch) > 0xFF);
        payload.extend_from_slice(
            &crate::utils::truncate_usize_to_u16(self.user_name.chars().count()).to_le_bytes(),
        );
        payload.push(if high_byte { STRING_HIGH_BYTE } else { 0 });
        if high_byte {
            for unit in self.user_name.encode_utf16() {
                payload.extend_from_slice(&unit.to_le_bytes());
            }
        } else {
            payload.extend(self.user_name.chars().map(|ch| ch as u8));
        }
        payload.push(self.unused);
        payload
    }

    /// Unique user identifier.
    #[must_use]
    pub fn user_id(&self) -> i32 {
        self.user_id
    }

    /// GUID of the last set of revisions synced to by this user.
    #[must_use]
    pub fn guid(&self) -> &[u8; GUID_LEN] {
        &self.guid
    }

    /// Date and time the user opened the shared workbook.
    #[must_use]
    pub fn opened_at(&self) -> ShortDtr {
        self.opened_at
    }

    /// Name of this user.
    #[must_use]
    pub fn user_name(&self) -> &str {
        &self.user_name
    }
}

/// Delivery option of a routing slip (`DocRoute.delOption`, MS-XLS 2.4.91).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RoutingDelivery {
    /// Deliver to one recipient at a time.
    OneAtATime,
    /// Deliver to all recipients at once.
    AllAtOnce,
}

/// Typed `DocRoute` record content (MS-XLS 2.4.91): routing information for
/// a routing slip that sends the document in an e-mail message.
///
/// The ANSI strings are preserved as raw bytes without their NULL
/// terminator; the `CODEPAGE` interpretation is left to the caller.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DocRoute {
    /// Routing stage of the slip (`iStage`).
    stage: u16,
    /// Number of `RecipName` records that follow (`cRecip`).
    recipient_count: u16,
    /// Delivery option (`delOption`).
    delivery: RoutingDelivery,
    /// Whether the document has been routed (`fRouted`).
    routed: bool,
    /// Whether the document returns to the originator after the last
    /// recipient (`fReturnOrig`).
    return_to_originator: bool,
    /// Whether a status message is sent to the originator after routing
    /// (`fTrackStatus`).
    track_status: bool,
    /// Whether `custom_type` defines a custom message type (`fCustomType`).
    custom_type_defined: bool,
    /// Subject of the routed document (`szSubject`, without the NULL).
    subject: Vec<u8>,
    /// Message of the routed document (`szMessage`, without the NULL).
    message: Vec<u8>,
    /// Name of the routing identifier (`szRouteID`, without the NULL).
    route_id: Vec<u8>,
    /// Custom message type (`szCustType`, without the NULL).
    custom_type: Vec<u8>,
    /// Workbook title (`szBookTitle`, without the NULL).
    book_title: Vec<u8>,
    /// Originator's friendly name (`szOrg`, without the NULL).
    originator_name: Vec<u8>,
    /// Originator's messaging-system address identifier (`rgchSSAddr`,
    /// without the NULL).
    originator_address: Vec<u8>,
}

impl DocRoute {
    /// Parse a `DocRoute` record payload.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn parse(data: &[u8]) -> Result<Self> {
        if data.len() < DOC_ROUTE_HEADER_LEN {
            return Err(Error::InvalidLength {
                expected: DOC_ROUTE_HEADER_LEN,
                found: data.len(),
            });
        }
        let stage = read_u16(data, 0);
        let recipient_count = read_u16(data, 2);
        if stage > recipient_count.saturating_add(1) {
            return Err(invalid(
                DOC_ROUTE_RECORD_TYPE,
                format!("DocRoute iStage {stage} exceeds cRecip + 1"),
            ));
        }
        let delivery = match read_u16(data, 4) {
            DELIVER_ONE_AT_A_TIME => RoutingDelivery::OneAtATime,
            DELIVER_ALL_AT_ONCE => RoutingDelivery::AllAtOnce,
            other => {
                return Err(invalid(
                    DOC_ROUTE_RECORD_TYPE,
                    format!("unknown DocRoute delOption {other:#06X}"),
                ));
            },
        };
        let flags = read_u16(data, 6);
        if flags & ROUTE_FLAG_SAVE_ROUTE_INFO == 0 {
            return Err(invalid(
                DOC_ROUTE_RECORD_TYPE,
                "DocRoute fSaveRouteInfo must be set",
            ));
        }
        let custom_type_defined = flags & ROUTE_FLAG_CUSTOM_TYPE != 0;
        let cch_subject = read_u16(data, 8);
        let cch_message = read_u16(data, 10);
        let cch_route_id = read_u16(data, 12);
        let cch_custom_type = read_u16(data, 14);
        let cch_book_title = read_u16(data, 16);
        let cch_originator = read_u16(data, 18);
        for (field, cch) in [
            ("cchSubject", cch_subject),
            ("cchMessage", cch_message),
            ("cchRouteID", cch_route_id),
            ("cchCustType", cch_custom_type),
            ("cchBookTitle", cch_book_title),
            ("cchOrg", cch_originator),
        ] {
            if cch > DOC_ROUTE_MAX_STRING_CHARS {
                return Err(invalid(
                    DOC_ROUTE_RECORD_TYPE,
                    format!("DocRoute {field} {cch} exceeds {DOC_ROUTE_MAX_STRING_CHARS}"),
                ));
            }
        }
        if cch_custom_type != 0 && !custom_type_defined {
            return Err(invalid(
                DOC_ROUTE_RECORD_TYPE,
                "DocRoute cchCustType must be zero when fCustomType is clear",
            ));
        }
        let address_len = read_u32(data, 20);
        let total = u64::from(cch_subject)
            + u64::from(cch_message)
            + u64::from(cch_route_id)
            + u64::from(cch_custom_type)
            + u64::from(cch_book_title)
            + u64::from(cch_originator)
            + u64::from(address_len);
        if total > DOC_ROUTE_MAX_TOTAL_CHARS {
            return Err(invalid(
                DOC_ROUTE_RECORD_TYPE,
                format!(
                    "DocRoute string lengths total {total}; expected at most {DOC_ROUTE_MAX_TOTAL_CHARS}"
                ),
            ));
        }
        let mut offset = DOC_ROUTE_HEADER_LEN;
        let subject = read_nul_terminated_ansi(
            DOC_ROUTE_RECORD_TYPE,
            data,
            &mut offset,
            usize::from(cch_subject),
            "szSubject",
        )?;
        let message = read_nul_terminated_ansi(
            DOC_ROUTE_RECORD_TYPE,
            data,
            &mut offset,
            usize::from(cch_message),
            "szMessage",
        )?;
        let route_id = read_nul_terminated_ansi(
            DOC_ROUTE_RECORD_TYPE,
            data,
            &mut offset,
            usize::from(cch_route_id),
            "szRouteID",
        )?;
        let custom_type = read_nul_terminated_ansi(
            DOC_ROUTE_RECORD_TYPE,
            data,
            &mut offset,
            usize::from(cch_custom_type),
            "szCustType",
        )?;
        let book_title = read_nul_terminated_ansi(
            DOC_ROUTE_RECORD_TYPE,
            data,
            &mut offset,
            usize::from(cch_book_title),
            "szBookTitle",
        )?;
        let originator_name = read_nul_terminated_ansi(
            DOC_ROUTE_RECORD_TYPE,
            data,
            &mut offset,
            usize::from(cch_originator),
            "szOrg",
        )?;
        let originator_address = read_nul_terminated_ansi(
            DOC_ROUTE_RECORD_TYPE,
            data,
            &mut offset,
            address_len as usize,
            "rgchSSAddr",
        )?;
        if offset != data.len() {
            return Err(Error::InvalidLength {
                expected: offset,
                found: data.len(),
            });
        }
        Ok(Self {
            stage,
            recipient_count,
            delivery,
            routed: flags & ROUTE_FLAG_ROUTED != 0,
            return_to_originator: flags & ROUTE_FLAG_RETURN_ORIG != 0,
            track_status: flags & ROUTE_FLAG_TRACK_STATUS != 0,
            custom_type_defined,
            subject,
            message,
            route_id,
            custom_type,
            book_title,
            originator_name,
            originator_address,
        })
    }

    /// Serialize back to a complete `DocRoute` record payload.
    #[must_use]
    pub fn to_payload(&self) -> Vec<u8> {
        let mut payload = Vec::with_capacity(DOC_ROUTE_HEADER_LEN);
        payload.extend_from_slice(&self.stage.to_le_bytes());
        payload.extend_from_slice(&self.recipient_count.to_le_bytes());
        let del_option = match self.delivery {
            RoutingDelivery::OneAtATime => DELIVER_ONE_AT_A_TIME,
            RoutingDelivery::AllAtOnce => DELIVER_ALL_AT_ONCE,
        };
        payload.extend_from_slice(&del_option.to_le_bytes());
        let mut flags = ROUTE_FLAG_SAVE_ROUTE_INFO;
        if self.routed {
            flags |= ROUTE_FLAG_ROUTED;
        }
        if self.return_to_originator {
            flags |= ROUTE_FLAG_RETURN_ORIG;
        }
        if self.track_status {
            flags |= ROUTE_FLAG_TRACK_STATUS;
        }
        if self.custom_type_defined {
            flags |= ROUTE_FLAG_CUSTOM_TYPE;
        }
        payload.extend_from_slice(&flags.to_le_bytes());
        for content in [
            &self.subject,
            &self.message,
            &self.route_id,
            &self.custom_type,
            &self.book_title,
            &self.originator_name,
        ] {
            payload.extend_from_slice(
                &crate::utils::truncate_usize_to_u16(nul_terminated_len(content)).to_le_bytes(),
            );
        }
        payload.extend_from_slice(
            &crate::utils::truncate_usize_to_u32(nul_terminated_len(&self.originator_address))
                .to_le_bytes(),
        );
        push_nul_terminated_ansi(&mut payload, &self.subject);
        push_nul_terminated_ansi(&mut payload, &self.message);
        push_nul_terminated_ansi(&mut payload, &self.route_id);
        push_nul_terminated_ansi(&mut payload, &self.custom_type);
        push_nul_terminated_ansi(&mut payload, &self.book_title);
        push_nul_terminated_ansi(&mut payload, &self.originator_name);
        push_nul_terminated_ansi(&mut payload, &self.originator_address);
        payload
    }

    /// Routing stage of the slip.
    #[must_use]
    pub fn stage(&self) -> u16 {
        self.stage
    }

    /// Number of `RecipName` records that follow this record.
    #[must_use]
    pub fn recipient_count(&self) -> u16 {
        self.recipient_count
    }

    /// Delivery option of the routing slip.
    #[must_use]
    pub fn delivery(&self) -> RoutingDelivery {
        self.delivery
    }

    /// Whether the document has been routed.
    #[must_use]
    pub fn routed(&self) -> bool {
        self.routed
    }

    /// Whether the document returns to the originator after the last recipient.
    #[must_use]
    pub fn return_to_originator(&self) -> bool {
        self.return_to_originator
    }

    /// Whether a status message is sent to the originator after routing.
    #[must_use]
    pub fn track_status(&self) -> bool {
        self.track_status
    }

    /// Whether a custom message type is defined by [`Self::custom_type`].
    #[must_use]
    pub fn custom_type_defined(&self) -> bool {
        self.custom_type_defined
    }

    /// Subject of the routed document (ANSI bytes without the NULL).
    #[must_use]
    pub fn subject(&self) -> &[u8] {
        &self.subject
    }

    /// Message of the routed document (ANSI bytes without the NULL).
    #[must_use]
    pub fn message(&self) -> &[u8] {
        &self.message
    }

    /// Name of the routing identifier (ANSI bytes without the NULL).
    #[must_use]
    pub fn route_id(&self) -> &[u8] {
        &self.route_id
    }

    /// Custom message type (ANSI bytes without the NULL).
    #[must_use]
    pub fn custom_type(&self) -> &[u8] {
        &self.custom_type
    }

    /// Workbook title (ANSI bytes without the NULL).
    #[must_use]
    pub fn book_title(&self) -> &[u8] {
        &self.book_title
    }

    /// Originator's friendly name (ANSI bytes without the NULL).
    #[must_use]
    pub fn originator_name(&self) -> &[u8] {
        &self.originator_name
    }

    /// Originator's messaging-system address identifier (ANSI bytes without
    /// the NULL).
    #[must_use]
    pub fn originator_address(&self) -> &[u8] {
        &self.originator_address
    }
}

/// Typed `RecipName` record content (MS-XLS 2.4.216): information about one
/// recipient of a routing slip.
///
/// The ANSI strings are preserved as raw bytes; the `CODEPAGE`
/// interpretation is left to the caller.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RecipName {
    /// Recipient's friendly name (`szFriendly`, without the NULL).
    friendly_name: Vec<u8>,
    /// Recipient's messaging-system address identifier (`rgchSSAddr`).
    address: Vec<u8>,
}

impl RecipName {
    /// Parse a `RecipName` record payload.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn parse(data: &[u8]) -> Result<Self> {
        if data.len() < RECIP_NAME_HEADER_LEN {
            return Err(Error::InvalidLength {
                expected: RECIP_NAME_HEADER_LEN,
                found: data.len(),
            });
        }
        let cch_recip = read_u16(data, 0);
        if cch_recip > RECIP_NAME_MAX_FRIENDLY_CHARS {
            return Err(invalid(
                RECIP_NAME_RECORD_TYPE,
                format!("RecipName cchRecip {cch_recip} exceeds {RECIP_NAME_MAX_FRIENDLY_CHARS}"),
            ));
        }
        let address_len = read_u32(data, 2) as usize;
        let mut offset = RECIP_NAME_HEADER_LEN;
        let friendly_name = read_nul_terminated_ansi(
            RECIP_NAME_RECORD_TYPE,
            data,
            &mut offset,
            usize::from(cch_recip),
            "szFriendly",
        )?;
        let end = offset
            .checked_add(address_len)
            .ok_or_else(|| invalid(RECIP_NAME_RECORD_TYPE, "rgchSSAddr length overflow"))?;
        if end > data.len() {
            return Err(invalid(
                RECIP_NAME_RECORD_TYPE,
                "rgchSSAddr extends past the end of the record",
            ));
        }
        let address = data[offset..end].to_vec();
        if end != data.len() {
            return Err(Error::InvalidLength {
                expected: end,
                found: data.len(),
            });
        }
        Ok(Self {
            friendly_name,
            address,
        })
    }

    /// Serialize back to a complete `RecipName` record payload.
    #[must_use]
    pub fn to_payload(&self) -> Vec<u8> {
        let mut payload = Vec::with_capacity(RECIP_NAME_HEADER_LEN);
        payload.extend_from_slice(
            &crate::utils::truncate_usize_to_u16(nul_terminated_len(&self.friendly_name))
                .to_le_bytes(),
        );
        payload.extend_from_slice(
            &crate::utils::truncate_usize_to_u32(self.address.len()).to_le_bytes(),
        );
        push_nul_terminated_ansi(&mut payload, &self.friendly_name);
        payload.extend_from_slice(&self.address);
        payload
    }

    /// Recipient's friendly name (ANSI bytes without the NULL).
    #[must_use]
    pub fn friendly_name(&self) -> &[u8] {
        &self.friendly_name
    }

    /// Recipient's messaging-system address identifier.
    #[must_use]
    pub fn address(&self) -> &[u8] {
        &self.address
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn short_dtr_bytes() -> [u8; SHORT_DTR_LEN] {
        // 2024-05-06 07:08:09, Tuesday (weekday 2).
        [0xE8, 0x07, 5, 6, 7, 8, 9, 2]
    }

    fn usr_info_payload(name_cch: u16, name_flags: u8, name_bytes: &[u8]) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&42i32.to_le_bytes());
        data.extend_from_slice(&[0xAB; GUID_LEN]);
        data.extend_from_slice(&short_dtr_bytes());
        data.extend_from_slice(&name_cch.to_le_bytes());
        data.push(name_flags);
        data.extend_from_slice(name_bytes);
        data.push(0xEE); // unused
        data
    }

    #[test]
    fn cusr_round_trip() {
        let payload = [0x03, 0x00];
        let record = CUsr::parse(&payload).unwrap();
        assert_eq!(record.count(), 3);
        assert_eq!(record.to_payload(), payload);
    }

    #[test]
    fn cusr_rejects_bad_length_and_count() {
        assert!(CUsr::parse(&[0x01]).is_err());
        assert!(CUsr::parse(&[0x01, 0x00, 0x00]).is_err());
        // iCount MUST be at most 255.
        assert!(CUsr::parse(&[0x00, 0x01]).is_err());
    }

    #[test]
    fn cbusr_round_trip() {
        let mut payload = vec![0u8; CB_USR_LEN];
        payload[0..2].copy_from_slice(&57u16.to_le_bytes());
        payload[2..4].copy_from_slice(&60u16.to_le_bytes());
        let record = CbUsr::parse(&payload).unwrap();
        assert_eq!(record.sizes()[0], 57);
        assert_eq!(record.sizes()[1], 60);
        assert_eq!(record.sizes()[2], 0);
        assert_eq!(record.to_payload(), payload);
    }

    #[test]
    fn cbusr_rejects_bad_length() {
        assert!(CbUsr::parse(&[0u8; CB_USR_LEN - 1]).is_err());
        assert!(CbUsr::parse(&[0u8; CB_USR_LEN + 1]).is_err());
    }

    #[test]
    fn usr_info_round_trip_compressed() {
        let payload = usr_info_payload(5, 0, b"Alice");
        let record = UsrInfo::parse(&payload).unwrap();
        assert_eq!(record.user_id(), 42);
        assert_eq!(record.guid(), &[0xAB; GUID_LEN]);
        assert_eq!(record.opened_at().year(), 2024);
        assert_eq!(record.opened_at().month(), 5);
        assert_eq!(record.opened_at().weekday(), 2);
        assert_eq!(record.user_name(), "Alice");
        assert_eq!(record.to_payload(), payload);
    }

    #[test]
    fn usr_info_round_trip_utf16() {
        // "Яб": two Cyrillic characters that do not fit in one byte.
        let name: Vec<u8> = "Яб".encode_utf16().flat_map(u16::to_le_bytes).collect();
        let payload = usr_info_payload(2, STRING_HIGH_BYTE, &name);
        let record = UsrInfo::parse(&payload).unwrap();
        assert_eq!(record.user_name(), "Яб");
        assert_eq!(record.to_payload(), payload);
    }

    #[test]
    fn usr_info_rejects_bad_name_counts() {
        // Zero characters is below the minimum of 1.
        let mut payload = usr_info_payload(5, 0, b"Alice");
        payload[USR_INFO_PREFIX_LEN] = 0;
        payload[USR_INFO_PREFIX_LEN + 1] = 0;
        assert!(UsrInfo::parse(&payload).is_err());
        // 55 characters exceeds the maximum of 54.
        let payload = usr_info_payload(55, 0, &[b'a'; 55]);
        assert!(UsrInfo::parse(&payload).is_err());
    }

    #[test]
    fn usr_info_rejects_rich_and_extended_strings() {
        let payload = usr_info_payload(5, 0x08, b"Alice"); // fRichSt
        assert!(UsrInfo::parse(&payload).is_err());
        let payload = usr_info_payload(5, 0x04, b"Alice"); // fExtSt
        assert!(UsrInfo::parse(&payload).is_err());
    }

    #[test]
    fn usr_info_rejects_truncation_and_trailing_garbage() {
        let payload = usr_info_payload(5, 0, b"Alice");
        assert!(UsrInfo::parse(&payload[..payload.len() - 1]).is_err());
        let mut longer = payload.clone();
        longer.push(0);
        assert!(UsrInfo::parse(&longer).is_err());
        // Invalid ShortDTR month.
        let mut bad_dtr = payload;
        bad_dtr[4 + GUID_LEN + 2] = 13;
        assert!(UsrInfo::parse(&bad_dtr).is_err());
    }

    /// Build a `DocRoute` payload with all strings populated.
    fn doc_route_payload() -> Vec<u8> {
        let strings: [&[u8]; 7] = [
            b"Subject",
            b"Message",
            b"RouteID",
            b"IPM.Route",
            b"Book1.xls",
            b"Alice",
            b"alice@example.com",
        ];
        let mut data = Vec::new();
        data.extend_from_slice(&2u16.to_le_bytes()); // iStage
        data.extend_from_slice(&3u16.to_le_bytes()); // cRecip
        data.extend_from_slice(&DELIVER_ALL_AT_ONCE.to_le_bytes());
        let flags = ROUTE_FLAG_ROUTED
            | ROUTE_FLAG_RETURN_ORIG
            | ROUTE_FLAG_TRACK_STATUS
            | ROUTE_FLAG_CUSTOM_TYPE
            | ROUTE_FLAG_SAVE_ROUTE_INFO;
        data.extend_from_slice(&flags.to_le_bytes());
        for content in &strings[..6] {
            data.extend_from_slice(&(content.len() as u16 + 1).to_le_bytes());
        }
        data.extend_from_slice(&(strings[6].len() as u32 + 1).to_le_bytes());
        for content in &strings {
            data.extend_from_slice(content);
            data.push(0);
        }
        data
    }

    #[test]
    fn doc_route_round_trip() {
        let payload = doc_route_payload();
        let record = DocRoute::parse(&payload).unwrap();
        assert_eq!(record.stage(), 2);
        assert_eq!(record.recipient_count(), 3);
        assert_eq!(record.delivery(), RoutingDelivery::AllAtOnce);
        assert!(record.routed());
        assert!(record.return_to_originator());
        assert!(record.track_status());
        assert!(record.custom_type_defined());
        assert_eq!(record.subject(), b"Subject");
        assert_eq!(record.message(), b"Message");
        assert_eq!(record.route_id(), b"RouteID");
        assert_eq!(record.custom_type(), b"IPM.Route");
        assert_eq!(record.book_title(), b"Book1.xls");
        assert_eq!(record.originator_name(), b"Alice");
        assert_eq!(record.originator_address(), b"alice@example.com");
        assert_eq!(record.to_payload(), payload);
    }

    #[test]
    fn doc_route_round_trip_empty_strings() {
        let mut data = Vec::new();
        data.extend_from_slice(&0u16.to_le_bytes()); // iStage
        data.extend_from_slice(&1u16.to_le_bytes()); // cRecip
        data.extend_from_slice(&DELIVER_ONE_AT_A_TIME.to_le_bytes());
        data.extend_from_slice(&ROUTE_FLAG_SAVE_ROUTE_INFO.to_le_bytes());
        data.extend_from_slice(&[0; 12]); // six zero cch fields
        data.extend_from_slice(&0u32.to_le_bytes()); // ulEIDSize
        let record = DocRoute::parse(&data).unwrap();
        assert!(!record.routed());
        assert!(!record.custom_type_defined());
        assert!(record.subject().is_empty());
        assert!(record.originator_address().is_empty());
        assert_eq!(record.to_payload(), data);
    }

    #[test]
    fn doc_route_rejects_stage_past_recipients() {
        let mut payload = doc_route_payload();
        payload[0..2].copy_from_slice(&5u16.to_le_bytes()); // iStage 5 > 3 + 1
        assert!(DocRoute::parse(&payload).is_err());
    }

    #[test]
    fn doc_route_rejects_unknown_delivery_option() {
        let mut payload = doc_route_payload();
        payload[4..6].copy_from_slice(&2u16.to_le_bytes());
        assert!(DocRoute::parse(&payload).is_err());
    }

    #[test]
    fn doc_route_requires_save_route_info() {
        // fSaveRouteInfo is bit 7 of the bitfield, i.e. the low flag byte.
        let mut payload = doc_route_payload();
        payload[6] &= !0x80;
        assert!(DocRoute::parse(&payload).is_err());
    }

    #[test]
    fn doc_route_rejects_oversized_and_orphaned_strings() {
        // cchSubject above the 256 maximum.
        let mut payload = doc_route_payload();
        payload[8..10].copy_from_slice(&257u16.to_le_bytes());
        assert!(DocRoute::parse(&payload).is_err());
        // cchCustType without fCustomType.
        let mut payload = doc_route_payload();
        payload[6] &= !ROUTE_FLAG_CUSTOM_TYPE as u8;
        assert!(DocRoute::parse(&payload).is_err());
    }

    #[test]
    fn doc_route_rejects_total_over_limit() {
        // ulEIDSize pushes the combined length past 8202.
        let mut payload = doc_route_payload();
        payload[20..24].copy_from_slice(&9000u32.to_le_bytes());
        assert!(DocRoute::parse(&payload).is_err());
    }

    #[test]
    fn doc_route_rejects_missing_nul_and_truncation() {
        // Overwrite the terminator of szSubject.
        let mut payload = doc_route_payload();
        payload[DOC_ROUTE_HEADER_LEN + 7] = b'!';
        assert!(DocRoute::parse(&payload).is_err());
        // Truncated record.
        let payload = doc_route_payload();
        assert!(DocRoute::parse(&payload[..payload.len() - 3]).is_err());
        assert!(DocRoute::parse(&payload[..DOC_ROUTE_HEADER_LEN - 1]).is_err());
    }

    #[test]
    fn recip_name_round_trip() {
        let mut payload = Vec::new();
        payload.extend_from_slice(&8u16.to_le_bytes()); // cchRecip
        payload.extend_from_slice(&16u32.to_le_bytes()); // ulEIDSize
        payload.extend_from_slice(b"Bob Doe\0");
        payload.extend_from_slice(b"bob@example.com\0"); // opaque bytes
        let record = RecipName::parse(&payload).unwrap();
        assert_eq!(record.friendly_name(), b"Bob Doe");
        assert_eq!(record.address(), b"bob@example.com\0");
        assert_eq!(record.to_payload(), payload);
    }

    #[test]
    fn recip_name_rejects_bad_counts_and_truncation() {
        // cchRecip above the 256 maximum.
        let mut payload = Vec::new();
        payload.extend_from_slice(&257u16.to_le_bytes());
        payload.extend_from_slice(&0u32.to_le_bytes());
        assert!(RecipName::parse(&payload).is_err());
        // Header too short.
        assert!(RecipName::parse(&[0u8; RECIP_NAME_HEADER_LEN - 1]).is_err());
        // szFriendly without its NULL terminator.
        let mut payload = Vec::new();
        payload.extend_from_slice(&4u16.to_le_bytes());
        payload.extend_from_slice(&0u32.to_le_bytes());
        payload.extend_from_slice(b"Bob!");
        assert!(RecipName::parse(&payload).is_err());
        // rgchSSAddr extends past the end.
        let mut payload = Vec::new();
        payload.extend_from_slice(&0u16.to_le_bytes());
        payload.extend_from_slice(&10u32.to_le_bytes());
        payload.extend_from_slice(b"abc");
        assert!(RecipName::parse(&payload).is_err());
        // Trailing garbage.
        let mut payload = Vec::new();
        payload.extend_from_slice(&0u16.to_le_bytes());
        payload.extend_from_slice(&0u32.to_le_bytes());
        payload.push(0);
        assert!(RecipName::parse(&payload).is_err());
    }
}
