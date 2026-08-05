//! Grammar-checker cookie tables (`Plcfcookie` and `PlcfcookieOld`).
//!
//! Cookies are opaque implementation-specific blobs stored in the `RgCdb`
//! referenced by `fcCookieData`; each `FCKS` (MS-DOC 2.9.76) or `FCKSOLD`
//! (MS-DOC 2.9.77) element describes where one cookie applies. The cookie
//! bytes themselves are never interpreted.

use super::super::package::{Error as PackageError, Result};
use super::fib::FileInformationBlock;

/// Table-pointer index of `fcPlcfcookie`/`lcbPlcfcookie` (MS-DOC 2.5.8 FibRgFcLcb2002).
const COOKIE_FIB_INDEX: usize = 116;
/// Table-pointer index of `fcPlcfCookieOld`/`lcbPlcfCookieOld` (MS-DOC 2.5.7 FibRgFcLcb2000).
const COOKIE_OLD_FIB_INDEX: usize = 101;
/// Table-pointer index of `fcCookieData`/`lcbCookieData` (MS-DOC 2.5.6 FibRgFcLcb97).
const COOKIE_DATA_FIB_INDEX: usize = 62;
const MAX_COOKIE_ENTRIES: usize = 1_000_000;
/// CPs are signed 31-bit positions in the set of all document parts (MS-DOC 2.2.1).
const MAX_CP: u32 = i32::MAX as u32;

fn corrupted(message: impl Into<String>) -> PackageError {
    PackageError::Corrupted(message.into())
}

fn read_i16(data: &[u8], offset: usize, field: &str) -> Result<i16> {
    litchi_core::binary::read_i16_le(data, offset)
        .map_err(|error| corrupted(format!("invalid {field}: {error}")))
}

fn read_u16(data: &[u8], offset: usize, field: &str) -> Result<u16> {
    litchi_core::binary::read_u16_le(data, offset)
        .map_err(|error| corrupted(format!("invalid {field}: {error}")))
}

fn read_u32(data: &[u8], offset: usize, field: &str) -> Result<u32> {
    litchi_core::binary::read_u32_le(data, offset)
        .map_err(|error| corrupted(format!("invalid {field}: {error}")))
}

/// Grammar checker error classification (`FCKS.cet`/`FCKSOLD.cet`, MS-DOC 2.9.76).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum CookieErrorType {
    /// Not a typo, homonym, or consistency error.
    Default = 0x0,
    Typo = 0x1,
    Homonym = 0x2,
    Consistency = 0x3,
}

impl CookieErrorType {
    fn from_raw(value: u8) -> Result<Self> {
        match value {
            0x0 => Ok(Self::Default),
            0x1 => Ok(Self::Typo),
            0x2 => Ok(Self::Homonym),
            0x3 => Ok(Self::Consistency),
            _ => Err(corrupted(format!("invalid cookie error type 0x{value:X}"))),
        }
    }
}

/// A grammar checker cookie descriptor (`FCKS`, MS-DOC 2.9.76; 10 bytes).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct GrammarCookie {
    dcp: i16,
    dcp_sent: i16,
    icdb: u32,
    error_type: CookieErrorType,
    is_error: bool,
    lid_sub: u8,
    lid_primary: u8,
    is_header: bool,
}

impl GrammarCookie {
    /// Serialized size of one `FCKS` (MS-DOC 2.9.76).
    pub const SIZE: usize = 10;

    const LID_SUB_MAX: u8 = 0x1F;
    const LID_PRIMARY_MAX: u8 = 0x7F;

    /// Create a cookie descriptor, validating the bit-field widths.
    #[allow(clippy::too_many_arguments)]
    pub fn try_new(
        dcp: i16,
        dcp_sent: i16,
        icdb: u32,
        error_type: CookieErrorType,
        is_error: bool,
        lid_sub: u8,
        lid_primary: u8,
        is_header: bool,
    ) -> Result<Self> {
        if lid_sub > Self::LID_SUB_MAX {
            return Err(corrupted("FCKS lidSub exceeds its 5-bit field"));
        }
        if lid_primary > Self::LID_PRIMARY_MAX {
            return Err(corrupted("FCKS lidPrimary exceeds its 7-bit field"));
        }
        Ok(Self {
            dcp,
            dcp_sent,
            icdb,
            error_type,
            is_error,
            lid_sub,
            lid_primary,
            is_header,
        })
    }

    /// Decode one 10-byte `FCKS`.
    pub fn from_bytes(data: &[u8]) -> Result<Self> {
        if data.len() != Self::SIZE {
            return Err(corrupted("FCKS must be exactly 10 bytes"));
        }
        let flags = read_u16(data, 8, "FCKS flags")?;
        Ok(Self {
            dcp: read_i16(data, 0, "FCKS dcp")?,
            dcp_sent: read_i16(data, 2, "FCKS dcpSent")?,
            icdb: read_u32(data, 4, "FCKS icdb")?,
            error_type: CookieErrorType::from_raw((flags & 0x3) as u8)?,
            is_error: flags & 0x4 != 0,
            lid_sub: ((flags >> 3) & 0x1F) as u8,
            lid_primary: ((flags >> 8) & 0x7F) as u8,
            is_header: flags & 0x8000 != 0,
        })
    }

    /// Serialize exactly as decoded.
    pub fn to_bytes(self) -> [u8; Self::SIZE] {
        let flags = self.error_type as u16
            | u16::from(self.is_error) << 2
            | u16::from(self.lid_sub) << 3
            | u16::from(self.lid_primary) << 8
            | u16::from(self.is_header) << 15;
        let mut data = [0u8; Self::SIZE];
        data[0..2].copy_from_slice(&self.dcp.to_le_bytes());
        data[2..4].copy_from_slice(&self.dcp_sent.to_le_bytes());
        data[4..8].copy_from_slice(&self.icdb.to_le_bytes());
        data[8..10].copy_from_slice(&flags.to_le_bytes());
        data
    }

    /// Characters spanned by the cookie text (ignored for header entries).
    pub fn dcp(&self) -> i16 {
        self.dcp
    }
    /// Characters from the cookie text back to its sentence start.
    pub fn dcp_sent(&self) -> i16 {
        self.dcp_sent
    }
    /// Byte offset of the cookie within the `RgCdb` at `fcCookieData`.
    pub fn icdb(&self) -> u32 {
        self.icdb
    }
    pub fn error_type(&self) -> CookieErrorType {
        self.error_type
    }
    /// Whether the cookie corresponds to an error displayed to the user.
    pub fn is_error(&self) -> bool {
        self.is_error
    }
    /// Bits 9-13 of the creating grammar checker's language ID.
    pub fn lid_sub(&self) -> u8 {
        self.lid_sub
    }
    /// The 7 least significant bits of the checker's language ID.
    pub fn lid_primary(&self) -> u8 {
        self.lid_primary
    }
    /// Whether this is the checker's single implementation-specific header entry.
    pub fn is_header(&self) -> bool {
        self.is_header
    }
}

/// A legacy grammar checker cookie descriptor (`FCKSOLD`, MS-DOC 2.9.77; 16 bytes).
///
/// The padding and spare fields are undefined in the format; they are ignored
/// when reading and written as zero.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct LegacyGrammarCookie {
    lid: u16,
    dcp: i16,
    dcp_sent: i16,
    error_type: CookieErrorType,
    is_error: bool,
    icdb: u32,
}

impl LegacyGrammarCookie {
    /// Serialized size of one `FCKSOLD` (MS-DOC 2.9.77).
    pub const SIZE: usize = 16;

    /// Create a legacy cookie descriptor, validating the signed field ranges.
    pub fn try_new(
        lid: u16,
        dcp: i16,
        dcp_sent: i16,
        error_type: CookieErrorType,
        is_error: bool,
        icdb: u32,
    ) -> Result<Self> {
        if dcp < 0 {
            return Err(corrupted("FCKSOLD dcp must be nonnegative"));
        }
        if dcp_sent > 0 {
            return Err(corrupted("FCKSOLD dcpSent must be nonpositive"));
        }
        Ok(Self {
            lid,
            dcp,
            dcp_sent,
            error_type,
            is_error,
            icdb,
        })
    }

    /// Decode one 16-byte `FCKSOLD`. Undefined padding is ignored.
    pub fn from_bytes(data: &[u8]) -> Result<Self> {
        if data.len() != Self::SIZE {
            return Err(corrupted("FCKSOLD must be exactly 16 bytes"));
        }
        let flags = read_u16(data, 8, "FCKSOLD flags")?;
        Self::try_new(
            read_u16(data, 0, "FCKSOLD lid")?,
            read_i16(data, 2, "FCKSOLD dcp")?,
            read_i16(data, 4, "FCKSOLD dcpSent")?,
            CookieErrorType::from_raw((flags & 0x3) as u8)?,
            flags & 0x8000 != 0,
            read_u32(data, 12, "FCKSOLD icdb")?,
        )
    }

    /// Serialize with zeroed padding and spare bits.
    pub fn to_bytes(self) -> [u8; Self::SIZE] {
        let flags = self.error_type as u16 | u16::from(self.is_error) << 15;
        let mut data = [0u8; Self::SIZE];
        data[0..2].copy_from_slice(&self.lid.to_le_bytes());
        data[2..4].copy_from_slice(&self.dcp.to_le_bytes());
        data[4..6].copy_from_slice(&self.dcp_sent.to_le_bytes());
        data[8..10].copy_from_slice(&flags.to_le_bytes());
        data[12..16].copy_from_slice(&self.icdb.to_le_bytes());
        data
    }

    /// Language ID of the grammar checker that created the cookie.
    pub fn lid(&self) -> u16 {
        self.lid
    }
    /// Characters spanned by the cookie text.
    pub fn dcp(&self) -> i16 {
        self.dcp
    }
    /// Characters from the cookie text back to its sentence start.
    pub fn dcp_sent(&self) -> i16 {
        self.dcp_sent
    }
    pub fn error_type(&self) -> CookieErrorType {
        self.error_type
    }
    /// Whether the cookie corresponds to an error displayed to the user.
    pub fn is_error(&self) -> bool {
        self.is_error
    }
    /// Byte offset of the cookie within the `RgCdb` at `fcCookieData`.
    pub fn icdb(&self) -> u32 {
        self.icdb
    }
}

/// Element stored in a grammar cookie PLC.
pub trait CookieElement: Copy {
    /// Serialized size of one element.
    const SIZE: usize;
    /// Structure name used in error messages.
    const KIND: &'static str;
    /// Decode one element from exactly [`CookieElement::SIZE`] bytes.
    fn parse_element(data: &[u8]) -> Result<Self>;
    /// Append the serialized element.
    fn write_element(self, out: &mut Vec<u8>);
    /// Byte offset of the cookie within the `RgCdb` cookie data.
    fn icdb(&self) -> u32;
    /// Checker identity `(lidPrimary, lidSub)` when this is a header entry.
    fn header_checker(&self) -> Option<(u8, u8)> {
        None
    }
}

impl CookieElement for GrammarCookie {
    const SIZE: usize = Self::SIZE;
    const KIND: &'static str = "FCKS";

    fn parse_element(data: &[u8]) -> Result<Self> {
        Self::from_bytes(data)
    }
    fn write_element(self, out: &mut Vec<u8>) {
        out.extend_from_slice(&self.to_bytes());
    }
    fn icdb(&self) -> u32 {
        self.icdb()
    }
    fn header_checker(&self) -> Option<(u8, u8)> {
        self.is_header()
            .then_some((self.lid_primary(), self.lid_sub()))
    }
}

impl CookieElement for LegacyGrammarCookie {
    const SIZE: usize = Self::SIZE;
    const KIND: &'static str = "FCKSOLD";

    fn parse_element(data: &[u8]) -> Result<Self> {
        Self::from_bytes(data)
    }
    fn write_element(self, out: &mut Vec<u8>) {
        out.extend_from_slice(&self.to_bytes());
    }
    fn icdb(&self) -> u32 {
        self.icdb()
    }
}

/// One cookie applying to text starting at `start_cp`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct CookieEntry<E> {
    start_cp: u32,
    cookie: E,
}

impl<E> CookieEntry<E> {
    pub const fn new(start_cp: u32, cookie: E) -> Self {
        Self { start_cp, cookie }
    }

    pub fn start_cp(&self) -> u32 {
        self.start_cp
    }
    pub fn cookie(&self) -> &E {
        &self.cookie
    }
}

/// A typed `Plcfcookie` (MS-DOC 2.8.14) or `PlcfcookieOld` (MS-DOC 2.8.15).
///
/// CPs are nondecreasing and duplicates are permitted. The final CP only
/// terminates the PLC and is ignored, as the format requires.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct GrammarCookiePlc<E> {
    entries: Vec<CookieEntry<E>>,
    terminal_cp: u32,
}

impl<E: CookieElement> GrammarCookiePlc<E> {
    pub fn try_new(entries: Vec<CookieEntry<E>>, terminal_cp: u32) -> Result<Self> {
        validate_entries(&entries, terminal_cp, None)?;
        Ok(Self {
            entries,
            terminal_cp,
        })
    }

    pub fn parse_bytes(data: &[u8]) -> Result<Self> {
        Self::parse_bytes_with_limits(data, None, None)
    }

    fn parse_bytes_with_limits(
        data: &[u8],
        maximum_cp: Option<u32>,
        cookie_data_len: Option<u32>,
    ) -> Result<Self> {
        let stride = 4 + E::SIZE;
        if data.len() < 4 || !(data.len() - 4).is_multiple_of(stride) {
            return Err(corrupted(format!(
                "{} PLC length must have form {}n + 4",
                E::KIND,
                stride
            )));
        }
        let count = (data.len() - 4) / stride;
        if count > MAX_COOKIE_ENTRIES {
            return Err(corrupted(format!(
                "{} PLC exceeds one-million-entry cap",
                E::KIND
            )));
        }
        let cp_bytes = count
            .checked_add(1)
            .and_then(|value| value.checked_mul(4))
            .ok_or_else(|| corrupted(format!("{} PLC CP array size overflows", E::KIND)))?;
        let mut positions = Vec::with_capacity(count + 1);
        for index in 0..=count {
            positions.push(read_u32(data, index * 4, "cookie PLC CP")?);
        }
        let terminal_cp = positions[count];
        let mut entries = Vec::with_capacity(count);
        for (index, &start_cp) in positions[..count].iter().enumerate() {
            let element_start = cp_bytes + index * E::SIZE;
            let cookie = E::parse_element(&data[element_start..element_start + E::SIZE])?;
            entries.push(CookieEntry::new(start_cp, cookie));
        }
        validate_entries(&entries, terminal_cp, maximum_cp)?;
        if let Some(data_len) = cookie_data_len {
            for entry in &entries {
                if entry.cookie.icdb() >= data_len {
                    return Err(corrupted(format!(
                        "{} icdb exceeds the RgCdb cookie data",
                        E::KIND
                    )));
                }
            }
        }
        Ok(Self {
            entries,
            terminal_cp,
        })
    }

    pub fn entries(&self) -> &[CookieEntry<E>] {
        &self.entries
    }
    /// Final PLC CP. The format ignores this value; it carries no range.
    pub fn terminal_cp(&self) -> u32 {
        self.terminal_cp
    }
    pub fn len(&self) -> usize {
        self.entries.len()
    }
    pub fn is_empty(&self) -> bool {
        self.entries.is_empty()
    }

    /// Serialize the complete PLC deterministically.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        validate_entries(&self.entries, self.terminal_cp, None)?;
        let size = self
            .entries
            .len()
            .checked_mul(4 + E::SIZE)
            .and_then(|value| value.checked_add(4))
            .ok_or_else(|| corrupted(format!("{} PLC serialized size overflows", E::KIND)))?;
        let mut data = Vec::with_capacity(size);
        for entry in &self.entries {
            data.extend_from_slice(&entry.start_cp.to_le_bytes());
        }
        data.extend_from_slice(&self.terminal_cp.to_le_bytes());
        for entry in &self.entries {
            entry.cookie.write_element(&mut data);
        }
        Ok(data)
    }
}

/// Typed `Plcfcookie` (Word 2002+ grammar cookies).
pub type GrammarCookieTable = GrammarCookiePlc<GrammarCookie>;
/// Typed `PlcfcookieOld` (legacy grammar cookies).
pub type LegacyGrammarCookieTable = GrammarCookiePlc<LegacyGrammarCookie>;

fn validate_entries<E: CookieElement>(
    entries: &[CookieEntry<E>],
    terminal_cp: u32,
    maximum_cp: Option<u32>,
) -> Result<()> {
    if entries.len() > MAX_COOKIE_ENTRIES {
        return Err(corrupted(format!(
            "{} PLC exceeds one-million-entry cap",
            E::KIND
        )));
    }
    let mut previous = None;
    let mut header_checkers = std::collections::HashSet::new();
    for (index, entry) in entries.iter().enumerate() {
        if entry.start_cp > MAX_CP {
            return Err(corrupted(format!(
                "{} PLC CP {index} exceeds signed CP range",
                E::KIND
            )));
        }
        if previous.is_some_and(|value| entry.start_cp < value) {
            return Err(corrupted(format!(
                "{} PLC CPs are not nondecreasing",
                E::KIND
            )));
        }
        previous = Some(entry.start_cp);
        if let Some(checker) = entry.cookie.header_checker()
            && !header_checkers.insert(checker)
        {
            return Err(corrupted(format!(
                "{} PLC has duplicate header entries for one grammar checker",
                E::KIND
            )));
        }
    }
    if terminal_cp > MAX_CP {
        return Err(corrupted(format!(
            "{} PLC terminal CP exceeds signed CP range",
            E::KIND
        )));
    }
    if previous.is_some_and(|value| terminal_cp < value) {
        return Err(corrupted(format!(
            "{} PLC terminal CP precedes the final entry",
            E::KIND
        )));
    }
    if maximum_cp.is_some_and(|maximum| previous.is_some_and(|value| value > maximum)) {
        return Err(corrupted(format!(
            "{} PLC CP exceeds the document parts",
            E::KIND
        )));
    }
    Ok(())
}

/// Optional current and legacy grammar cookie tables for a document.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct GrammarCookieTables {
    current: Option<GrammarCookieTable>,
    legacy: Option<LegacyGrammarCookieTable>,
}

impl GrammarCookieTables {
    /// Parse both cookie PLCFs from the Table Stream.
    ///
    /// When the `RgCdb` cookie data is present (`lcbCookieData` nonzero),
    /// every cookie's `icdb` must fall inside it.
    pub fn parse(fib: &FileInformationBlock, table_stream: &[u8]) -> Result<Self> {
        let maximum_cp = fib
            .get_document_parts_end()
            .and_then(|value| value.checked_add(2))
            .ok_or_else(|| corrupted("document-parts cookie CP ceiling overflows"))?;
        let cookie_data_len = fib
            .get_table_pointer(COOKIE_DATA_FIB_INDEX)
            .map(|(_, length)| length)
            .filter(|length| *length > 0);
        Ok(Self {
            current: parse_fib_table(
                fib,
                table_stream,
                COOKIE_FIB_INDEX,
                maximum_cp,
                cookie_data_len,
            )?,
            legacy: parse_fib_table(
                fib,
                table_stream,
                COOKIE_OLD_FIB_INDEX,
                maximum_cp,
                cookie_data_len,
            )?,
        })
    }

    /// Current grammar cookies (`Plcfcookie`, MS-DOC 2.8.14).
    pub fn current(&self) -> Option<&GrammarCookieTable> {
        self.current.as_ref()
    }
    /// Legacy grammar cookies (`PlcfcookieOld`, MS-DOC 2.8.15).
    pub fn legacy(&self) -> Option<&LegacyGrammarCookieTable> {
        self.legacy.as_ref()
    }
}

fn parse_fib_table<E: CookieElement>(
    fib: &FileInformationBlock,
    table_stream: &[u8],
    index: usize,
    maximum_cp: u32,
    cookie_data_len: Option<u32>,
) -> Result<Option<GrammarCookiePlc<E>>> {
    let Some((offset, length)) = fib.get_table_pointer(index) else {
        return Ok(None);
    };
    if length == 0 {
        return Ok(None);
    }
    let start = usize::try_from(offset).map_err(|_| corrupted("cookie PLC offset is too large"))?;
    let length =
        usize::try_from(length).map_err(|_| corrupted("cookie PLC length is too large"))?;
    let end = start
        .checked_add(length)
        .ok_or_else(|| corrupted("cookie PLC range overflows"))?;
    let data = table_stream
        .get(start..end)
        .ok_or_else(|| corrupted("cookie PLC extends beyond the table stream"))?;
    GrammarCookiePlc::parse_bytes_with_limits(data, Some(maximum_cp), cookie_data_len).map(Some)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn cookie(dcp: i16, icdb: u32, is_header: bool) -> GrammarCookie {
        GrammarCookie::try_new(
            dcp,
            -2,
            icdb,
            CookieErrorType::Typo,
            true,
            0x05,
            0x09,
            is_header,
        )
        .unwrap()
    }

    fn plc_bytes<E: CookieElement>(cps: &[u32], terminal: u32, cookies: &[E]) -> Vec<u8> {
        let mut data = Vec::new();
        for cp in cps {
            data.extend_from_slice(&cp.to_le_bytes());
        }
        data.extend_from_slice(&terminal.to_le_bytes());
        for cookie in cookies {
            cookie.write_element(&mut data);
        }
        data
    }

    #[test]
    fn fcks_round_trips_exactly() {
        let cookie = cookie(7, 48, false);
        let bytes = cookie.to_bytes();
        assert_eq!(bytes.len(), GrammarCookie::SIZE);
        assert_eq!(GrammarCookie::from_bytes(&bytes).unwrap(), cookie);
        // Flags: cet Typo | fError | lidSub 0x05 | lidPrimary 0x09.
        let flags = u16::from_le_bytes([bytes[8], bytes[9]]);
        assert_eq!(flags, 0x1 | 0x4 | 0x05 << 3 | 0x09 << 8);
        assert!(GrammarCookie::from_bytes(&bytes[..9]).is_err());
        assert!(
            GrammarCookie::try_new(1, 0, 0, CookieErrorType::Default, false, 0x20, 0, false)
                .is_err()
        );
        assert!(
            GrammarCookie::try_new(1, 0, 0, CookieErrorType::Default, false, 0, 0x80, false)
                .is_err()
        );
    }

    #[test]
    fn fcksold_validates_signed_ranges_and_zeroes_padding() {
        let legacy =
            LegacyGrammarCookie::try_new(0x0409, 5, -3, CookieErrorType::Homonym, true, 12)
                .unwrap();
        let bytes = legacy.to_bytes();
        assert_eq!(bytes.len(), LegacyGrammarCookie::SIZE);
        assert_eq!(&bytes[6..8], &[0, 0]);
        assert_eq!(&bytes[10..12], &[0, 0]);
        let flags = u16::from_le_bytes([bytes[8], bytes[9]]);
        assert_eq!(flags, 0x2 | 0x8000);
        assert_eq!(LegacyGrammarCookie::from_bytes(&bytes).unwrap(), legacy);
        // Undefined padding and spare bits are ignored on read.
        let mut messy = bytes;
        messy[6] = 0xFF;
        messy[9] |= 0x7C;
        assert_eq!(LegacyGrammarCookie::from_bytes(&messy).unwrap(), legacy);
        assert!(
            LegacyGrammarCookie::try_new(0, -1, 0, CookieErrorType::Default, false, 0).is_err()
        );
        assert!(LegacyGrammarCookie::try_new(0, 0, 1, CookieErrorType::Default, false, 0).is_err());
    }

    #[test]
    fn plcfcookie_parses_and_round_trips() {
        let cookies = [
            cookie(7, 0, false),
            cookie(3, 40, true),
            cookie(2, 48, false),
        ];
        let bytes = plc_bytes(&[10, 20, 20], 30, &cookies);
        let table = GrammarCookieTable::parse_bytes(&bytes).unwrap();
        assert_eq!(table.len(), 3);
        assert_eq!(table.terminal_cp(), 30);
        assert_eq!(table.entries()[2].start_cp(), 20);
        assert_eq!(table.entries()[2].cookie().icdb(), 48);
        assert!(table.entries()[1].cookie().is_header());
        assert_eq!(table.to_bytes().unwrap(), bytes);
    }

    #[test]
    fn plcfcookieold_parses_and_round_trips() {
        let cookies = [
            LegacyGrammarCookie::try_new(0x0409, 5, -1, CookieErrorType::Default, false, 0)
                .unwrap(),
            LegacyGrammarCookie::try_new(0x0809, 2, 0, CookieErrorType::Consistency, true, 8)
                .unwrap(),
        ];
        let bytes = plc_bytes(&[0, 12], 20, &cookies);
        let table = LegacyGrammarCookieTable::parse_bytes(&bytes).unwrap();
        assert_eq!(table.len(), 2);
        assert_eq!(table.entries()[1].cookie().lid(), 0x0809);
        assert_eq!(table.to_bytes().unwrap(), bytes);
    }

    #[test]
    fn rejects_malformed_plc_shapes() {
        assert!(GrammarCookieTable::parse_bytes(&[]).is_err());
        assert!(GrammarCookieTable::parse_bytes(&[0; 5]).is_err());
        let mut bytes = plc_bytes(&[10], 20, &[cookie(1, 0, false)]);
        bytes.pop();
        assert!(GrammarCookieTable::parse_bytes(&bytes).is_err());
        // Decreasing CPs.
        let bytes = plc_bytes(&[10, 5], 20, &[cookie(1, 0, false), cookie(1, 4, false)]);
        assert!(GrammarCookieTable::parse_bytes(&bytes).is_err());
        // Terminal CP before the final entry.
        let bytes = plc_bytes(&[10], 5, &[cookie(1, 0, false)]);
        assert!(GrammarCookieTable::parse_bytes(&bytes).is_err());
        // CPs beyond the signed range.
        let bytes = plc_bytes(&[0x8000_0000], 0x8000_0001, &[cookie(1, 0, false)]);
        assert!(GrammarCookieTable::parse_bytes(&bytes).is_err());
    }

    #[test]
    fn rejects_duplicate_checker_headers() {
        let bytes = plc_bytes(&[10, 20], 30, &[cookie(1, 0, true), cookie(1, 4, true)]);
        assert!(GrammarCookieTable::parse_bytes(&bytes).is_err());
        // Distinct checkers may each carry one header entry.
        let other_checker =
            GrammarCookie::try_new(1, 0, 8, CookieErrorType::Default, false, 0x05, 0x0A, true)
                .unwrap();
        let bytes = plc_bytes(&[10, 20], 30, &[cookie(1, 0, true), other_checker]);
        assert!(GrammarCookieTable::parse_bytes(&bytes).is_ok());
    }

    fn fib_with_pointers(pairs: &[(usize, u32, u32)]) -> FileInformationBlock {
        let mut data = vec![0u8; 154 + 117 * 8];
        data[0..2].copy_from_slice(&0xA5ECu16.to_le_bytes());
        data[2..4].copy_from_slice(&0x0101u16.to_le_bytes());
        data[152..154].copy_from_slice(&117u16.to_le_bytes());
        data[0x4C..0x50].copy_from_slice(&100u32.to_le_bytes());
        for (index, offset, length) in pairs {
            let pointer = 154 + index * 8;
            data[pointer..pointer + 4].copy_from_slice(&offset.to_le_bytes());
            data[pointer + 4..pointer + 8].copy_from_slice(&length.to_le_bytes());
        }
        FileInformationBlock::parse(&data).unwrap()
    }

    #[test]
    fn parses_both_tables_through_fib_with_cookie_data_bounds() {
        let current = plc_bytes(&[10], 20, &[cookie(5, 4, false)]);
        let legacy = plc_bytes(
            &[2],
            20,
            &[
                LegacyGrammarCookie::try_new(0x0409, 3, 0, CookieErrorType::Default, false, 8)
                    .unwrap(),
            ],
        );
        let mut table_stream = vec![0u8; 16];
        table_stream.extend_from_slice(&current);
        table_stream.extend_from_slice(&legacy);
        let fib = fib_with_pointers(&[
            (COOKIE_DATA_FIB_INDEX, 0x800, 16),
            (COOKIE_FIB_INDEX, 16, current.len() as u32),
            (
                COOKIE_OLD_FIB_INDEX,
                (16 + current.len()) as u32,
                legacy.len() as u32,
            ),
        ]);
        let tables = GrammarCookieTables::parse(&fib, &table_stream).unwrap();
        assert_eq!(tables.current().unwrap().len(), 1);
        assert_eq!(tables.legacy().unwrap().len(), 1);
    }

    #[test]
    fn rejects_icdb_outside_cookie_data() {
        let current = plc_bytes(&[10], 20, &[cookie(5, 16, false)]);
        let fib = fib_with_pointers(&[
            (COOKIE_DATA_FIB_INDEX, 0x800, 16),
            (COOKIE_FIB_INDEX, 0, current.len() as u32),
        ]);
        assert!(GrammarCookieTables::parse(&fib, &current).is_err());
        // Without a cookie-data region the icdb bounds check does not apply.
        let fib = fib_with_pointers(&[(COOKIE_FIB_INDEX, 0, current.len() as u32)]);
        assert!(GrammarCookieTables::parse(&fib, &current).is_ok());
    }

    #[test]
    fn rejects_cookie_cp_beyond_document_parts() {
        let current = plc_bytes(&[500], 600, &[cookie(5, 0, false)]);
        let fib = fib_with_pointers(&[(COOKIE_FIB_INDEX, 0, current.len() as u32)]);
        assert!(GrammarCookieTables::parse(&fib, &current).is_err());
    }
}
