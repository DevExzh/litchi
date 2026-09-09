//! Typed, inert `PGPArray` paragraph-group properties (MS-DOC 2.9.187–189).
//!
//! PGP entries describe paragraph borders, margins, and HTML-oriented block
//! types. They are metadata only: this module does not apply paragraph
//! properties or rewrite the document's paragraph stream. The exact bounded
//! source is retained so callers can inspect producer-specific border bytes.

use super::super::package::{Error as PackageError, Result};
use super::fib::FileInformationBlock;

/// FIB `FibRgFcLcb2000` index for `fcPlcfPgp`/`lcbPlcfPgp`.
pub const FIB_INDEX_PGP: usize = 109;
/// Maximum `PGPArray` payload materialized by this reader.
pub const MAX_PGP_BYTES: usize = 16 * 1024 * 1024;
const PGP_INFO_HEADER_SIZE: usize = 14;
const PGP_BRC_SIZE: usize = 8;

fn corrupted(message: impl Into<String>) -> PackageError {
    PackageError::Corrupted(message.into())
}

fn copy_bytes(data: &[u8], field: &str) -> Result<Vec<u8>> {
    let mut copy = Vec::new();
    copy.try_reserve_exact(data.len())
        .map_err(|_| corrupted(format!("PGP {field} allocation failed")))?;
    copy.extend_from_slice(data);
    Ok(copy)
}

fn read_u16(data: &[u8], offset: usize, field: &str) -> Result<u16> {
    let bytes = data
        .get(offset..offset + 2)
        .ok_or_else(|| corrupted(format!("PGP {field} is truncated")))?;
    Ok(u16::from_le_bytes([bytes[0], bytes[1]]))
}

fn read_u32(data: &[u8], offset: usize, field: &str) -> Result<u32> {
    let bytes = data
        .get(offset..offset + 4)
        .ok_or_else(|| corrupted(format!("PGP {field} is truncated")))?;
    Ok(u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]))
}

fn read_i32(data: &[u8], offset: usize, field: &str) -> Result<i32> {
    let bytes = data
        .get(offset..offset + 4)
        .ok_or_else(|| corrupted(format!("PGP {field} is truncated")))?;
    Ok(i32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]))
}

/// The HTML-oriented type carried by `PGPOptions.type`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u16)]
pub enum PgpType {
    /// The default `DIV` interpretation.
    Div = 0,
    /// The block is represented as a `BLOCKQUOTE` when exported as HTML.
    BlockQuote = 1,
    /// The block is represented as a `BODY` where legal in HTML.
    Body = 2,
}

impl PgpType {
    fn from_raw(value: u16) -> Result<Self> {
        match value {
            0 => Ok(Self::Div),
            1 => Ok(Self::BlockQuote),
            2 => Ok(Self::Body),
            other => Err(corrupted(format!(
                "PGPOptions type has invalid value {other:#06x}"
            ))),
        }
    }
}

/// The variable properties selected by `PgpInfo.grfElements`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PgpOptions {
    cb_option: u16,
    dxa_left: Option<i32>,
    dxa_right: Option<i32>,
    dya_before: Option<i32>,
    dya_after: Option<i32>,
    brc_left: Option<[u8; PGP_BRC_SIZE]>,
    brc_right: Option<[u8; PGP_BRC_SIZE]>,
    brc_top: Option<[u8; PGP_BRC_SIZE]>,
    brc_bottom: Option<[u8; PGP_BRC_SIZE]>,
    kind: Option<PgpType>,
    /// Future or producer-specific option bits are retained without guessing
    /// their layout.
    opaque: Vec<u8>,
}

impl PgpOptions {
    /// Number of bytes following `cbOption` in the source.
    #[must_use]
    pub const fn cb_option(&self) -> u16 {
        self.cb_option
    }

    #[must_use]
    pub const fn dxa_left(&self) -> Option<i32> {
        self.dxa_left
    }

    #[must_use]
    pub const fn dxa_right(&self) -> Option<i32> {
        self.dxa_right
    }

    #[must_use]
    pub const fn dya_before(&self) -> Option<i32> {
        self.dya_before
    }

    #[must_use]
    pub const fn dya_after(&self) -> Option<i32> {
        self.dya_after
    }

    /// Raw `Brc` bytes for the left border, if present.
    #[must_use]
    pub fn brc_left(&self) -> Option<&[u8; PGP_BRC_SIZE]> {
        self.brc_left.as_ref()
    }

    /// Raw `Brc` bytes for the right border, if present.
    #[must_use]
    pub fn brc_right(&self) -> Option<&[u8; PGP_BRC_SIZE]> {
        self.brc_right.as_ref()
    }

    /// Raw `Brc` bytes for the top border, if present.
    #[must_use]
    pub fn brc_top(&self) -> Option<&[u8; PGP_BRC_SIZE]> {
        self.brc_top.as_ref()
    }

    /// Raw `Brc` bytes for the bottom border, if present.
    #[must_use]
    pub fn brc_bottom(&self) -> Option<&[u8; PGP_BRC_SIZE]> {
        self.brc_bottom.as_ref()
    }

    /// HTML-oriented block type, if the source supplied it.
    #[must_use]
    pub const fn kind(&self) -> Option<PgpType> {
        self.kind
    }

    /// Opaque bytes selected by future or producer-specific option bits.
    #[must_use]
    pub fn opaque(&self) -> &[u8] {
        &self.opaque
    }
}

/// One fixed-header `PGPInfo` plus its bounded variable options.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PgpInfo {
    ipgp_self: u32,
    ipgp_parent: u32,
    itap: u32,
    grf_elements: u16,
    options: PgpOptions,
}

impl PgpInfo {
    /// The nonzero identifier applied by paragraph property modifiers.
    #[must_use]
    pub const fn ipgp_self(&self) -> u32 {
        self.ipgp_self
    }

    /// The immediate parent identifier, or zero for an outermost group.
    #[must_use]
    pub const fn ipgp_parent(&self) -> u32 {
        self.ipgp_parent
    }

    /// Table depth at which this group applies.
    #[must_use]
    pub const fn itap(&self) -> u32 {
        self.itap
    }

    /// Presence bits for the variable `PgpOptions` members.
    #[must_use]
    pub const fn grf_elements(&self) -> u16 {
        self.grf_elements
    }

    /// Typed variable options and retained unknown option bytes.
    #[must_use]
    pub fn options(&self) -> &PgpOptions {
        &self.options
    }
}

/// A bounded `PGPArray` from the main table stream.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PgpArray {
    source: Vec<u8>,
    entries: Vec<PgpInfo>,
}

impl PgpArray {
    /// Parse the optional `PGPArray` selected by the Word 2000 FIB extension.
    pub fn parse(fib: &FileInformationBlock, table_stream: &[u8]) -> Result<Option<Self>> {
        let Some((offset, length)) = fib.get_table_pointer(FIB_INDEX_PGP) else {
            return Ok(None);
        };
        if length == 0 {
            return Ok(None);
        }
        let length = usize::try_from(length)
            .map_err(|_| corrupted("PGPArray length does not fit in memory"))?;
        if length > MAX_PGP_BYTES {
            return Err(corrupted(format!(
                "PGPArray exceeds the {MAX_PGP_BYTES}-byte limit"
            )));
        }
        let start = usize::try_from(offset)
            .map_err(|_| corrupted("PGPArray offset does not fit in memory"))?;
        let end = start
            .checked_add(length)
            .ok_or_else(|| corrupted("PGPArray range overflows"))?;
        let data = table_stream
            .get(start..end)
            .ok_or_else(|| corrupted("PGPArray extends beyond the table stream"))?;
        Self::parse_bytes(data).map(Some)
    }

    /// Parse one complete `PGPArray` payload.
    pub fn parse_bytes(data: &[u8]) -> Result<Self> {
        if data.len() > MAX_PGP_BYTES {
            return Err(corrupted(format!(
                "PGPArray exceeds the {MAX_PGP_BYTES}-byte limit"
            )));
        }
        let count = usize::from(read_u16(data, 0, "cpgp")?);
        let minimum = count
            .checked_mul(PGP_INFO_HEADER_SIZE)
            .and_then(|size| size.checked_add(2))
            .ok_or_else(|| corrupted("PGPArray entry count overflows"))?;
        if minimum > data.len() {
            return Err(corrupted("PGPArray has truncated PGPInfo entries"));
        }

        let mut offset = 2usize;
        let mut entries = Vec::new();
        entries
            .try_reserve_exact(count)
            .map_err(|_| corrupted("PGPArray entry allocation failed"))?;
        for index in 0..count {
            let ipgp_self = read_u32(data, offset, "ipgpSelf")?;
            if ipgp_self == 0 {
                return Err(corrupted(format!(
                    "PGPInfo {index} has a zero ipgpSelf identifier"
                )));
            }
            let ipgp_parent = read_u32(data, offset + 4, "ipgpParent")?;
            let itap = read_u32(data, offset + 8, "itap")?;
            let grf_elements = read_u16(data, offset + 12, "grfElements")?;
            offset += PGP_INFO_HEADER_SIZE;
            let options = parse_options(data, &mut offset, grf_elements)?;
            entries.push(PgpInfo {
                ipgp_self,
                ipgp_parent,
                itap,
                grf_elements,
                options,
            });
        }
        if offset != data.len() {
            return Err(corrupted(format!(
                "PGPArray has {} trailing bytes after cpgp entries",
                data.len() - offset
            )));
        }
        validate_relationships(&entries)?;
        Ok(Self {
            source: copy_bytes(data, "source")?,
            entries,
        })
    }

    /// Exact bounded source bytes, including retained opaque option bytes.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.source
    }

    /// Copy the exact serialized array.
    #[must_use]
    pub fn to_bytes(&self) -> Vec<u8> {
        self.source.clone()
    }

    /// Ordered PGP entries.
    #[must_use]
    pub fn entries(&self) -> &[PgpInfo] {
        &self.entries
    }

    /// Number of PGP entries.
    #[must_use]
    pub fn len(&self) -> usize {
        self.entries.len()
    }

    /// Whether the array contains no entries.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.entries.is_empty()
    }
}

/// Validate the identity graph described by `PGPInfo.ipgpSelf` and
/// `PGPInfo.ipgpParent`.
///
/// The parent identifiers are allowed to refer to an entry that appears later
/// in the array, so this is deliberately performed after all entries have been
/// decoded.  The sorted index keeps referent lookup bounded without retaining
/// a second owned copy of any source payload.  A three-state walk then rejects
/// cycles without recursive stack growth.
fn validate_relationships(entries: &[PgpInfo]) -> Result<()> {
    let mut identifiers = Vec::new();
    identifiers
        .try_reserve_exact(entries.len())
        .map_err(|_| corrupted("PGPArray identifier index allocation failed"))?;
    identifiers.extend(
        entries
            .iter()
            .enumerate()
            .map(|(index, entry)| (entry.ipgp_self, index)),
    );
    identifiers.sort_unstable_by_key(|(identifier, _)| *identifier);

    if let Some(pair) = identifiers.windows(2).find(|pair| pair[0].0 == pair[1].0) {
        return Err(corrupted(format!(
            "PGPInfo identifier {} is not unique",
            pair[0].0
        )));
    }

    for (index, entry) in entries.iter().enumerate() {
        if entry.ipgp_parent != 0 && find_identifier(&identifiers, entry.ipgp_parent).is_none() {
            return Err(corrupted(format!(
                "PGPInfo {index} has an unknown ipgpParent {}",
                entry.ipgp_parent
            )));
        }
    }

    // 0 = not visited, 1 = on the current ancestry walk, 2 = validated.
    let mut state = Vec::new();
    state
        .try_reserve_exact(entries.len())
        .map_err(|_| corrupted("PGPArray validation state allocation failed"))?;
    state.resize(entries.len(), 0);
    let mut path = Vec::new();
    path.try_reserve_exact(entries.len())
        .map_err(|_| corrupted("PGPArray ancestry allocation failed"))?;
    for start in 0..entries.len() {
        if state[start] == 2 {
            continue;
        }
        path.clear();
        let mut current = start;
        loop {
            match state[current] {
                2 => break,
                1 => {
                    return Err(corrupted(format!(
                        "PGPInfo ancestry contains a cycle at identifier {}",
                        entries[current].ipgp_self
                    )));
                },
                _ => {
                    state[current] = 1;
                    path.push(current);
                    let parent = entries[current].ipgp_parent;
                    if parent == 0 {
                        break;
                    }
                    let Some(parent) = find_identifier(&identifiers, parent) else {
                        return Err(corrupted(format!(
                            "PGPInfo {} has an unknown ipgpParent {parent}",
                            entries[current].ipgp_self
                        )));
                    };
                    current = parent;
                },
            }
        }
        for &index in &path {
            state[index] = 2;
        }
    }
    Ok(())
}

fn find_identifier(identifiers: &[(u32, usize)], identifier: u32) -> Option<usize> {
    identifiers
        .binary_search_by(|(candidate, _)| candidate.cmp(&identifier))
        .ok()
        .map(|index| identifiers[index].1)
}

fn parse_options(data: &[u8], offset: &mut usize, grf_elements: u16) -> Result<PgpOptions> {
    if grf_elements == 0 {
        return Ok(PgpOptions {
            cb_option: 0,
            dxa_left: None,
            dxa_right: None,
            dya_before: None,
            dya_after: None,
            brc_left: None,
            brc_right: None,
            brc_top: None,
            brc_bottom: None,
            kind: None,
            opaque: Vec::new(),
        });
    }

    let cb_option = read_u16(data, *offset, "cbOption")?;
    *offset += 2;
    let option_length = usize::from(cb_option);
    let end = (*offset)
        .checked_add(option_length)
        .ok_or_else(|| corrupted("PGPOptions range overflows"))?;
    if end > data.len() {
        return Err(corrupted("PGPOptions extends beyond PGPArray"));
    }

    let mut cursor = *offset;
    let dxa_left = read_optional_i32(data, &mut cursor, end, grf_elements, 0x0001, "dxaLeft")?;
    let dxa_right = read_optional_i32(data, &mut cursor, end, grf_elements, 0x0002, "dxaRight")?;
    let dya_before = read_optional_i32(data, &mut cursor, end, grf_elements, 0x0004, "dyaBefore")?;
    let dya_after = read_optional_i32(data, &mut cursor, end, grf_elements, 0x0008, "dyaAfter")?;
    let brc_left = read_optional_brc(data, &mut cursor, end, grf_elements, 0x0010, "brcLeft")?;
    let brc_right = read_optional_brc(data, &mut cursor, end, grf_elements, 0x0020, "brcRight")?;
    let brc_top = read_optional_brc(data, &mut cursor, end, grf_elements, 0x0040, "brcTop")?;
    let brc_bottom = read_optional_brc(data, &mut cursor, end, grf_elements, 0x0080, "brcBottom")?;
    let kind = if grf_elements & 0x0100 != 0 {
        if cursor.checked_add(2).is_none_or(|next| next > end) {
            return Err(corrupted("PGPOptions type exceeds cbOption"));
        }
        let value = read_u16(data, cursor, "type")?;
        cursor += 2;
        Some(PgpType::from_raw(value)?)
    } else {
        None
    };
    let opaque_source = data
        .get(cursor..end)
        .ok_or_else(|| corrupted("PGPOptions opaque range is invalid"))?;
    let opaque = copy_bytes(opaque_source, "opaque option")?;
    *offset = end;
    Ok(PgpOptions {
        cb_option,
        dxa_left,
        dxa_right,
        dya_before,
        dya_after,
        brc_left,
        brc_right,
        brc_top,
        brc_bottom,
        kind,
        opaque,
    })
}

fn read_optional_i32(
    data: &[u8],
    cursor: &mut usize,
    end: usize,
    bits: u16,
    bit: u16,
    field: &str,
) -> Result<Option<i32>> {
    if bits & bit == 0 {
        return Ok(None);
    }
    if (*cursor).checked_add(4).is_none_or(|next| next > end) {
        return Err(corrupted(format!("PGPOptions {field} is truncated")));
    }
    let value = read_i32(data, *cursor, field)?;
    *cursor += 4;
    Ok(Some(value))
}

fn read_optional_brc(
    data: &[u8],
    cursor: &mut usize,
    end: usize,
    bits: u16,
    bit: u16,
    field: &str,
) -> Result<Option<[u8; PGP_BRC_SIZE]>> {
    if bits & bit == 0 {
        return Ok(None);
    }
    let next = (*cursor)
        .checked_add(PGP_BRC_SIZE)
        .ok_or_else(|| corrupted(format!("PGPOptions {field} range overflows")))?;
    if next > end {
        return Err(corrupted(format!("PGPOptions {field} exceeds cbOption")));
    }
    let source = data
        .get(*cursor..next)
        .ok_or_else(|| corrupted(format!("PGPOptions {field} is truncated")))?;
    let value = source
        .try_into()
        .map_err(|_| corrupted(format!("PGPOptions {field} has invalid size")))?;
    *cursor = next;
    Ok(Some(value))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parses_bounded_pgp_info_and_opaque_options() {
        let mut bytes = Vec::new();
        bytes.extend_from_slice(&1u16.to_le_bytes());
        bytes.extend_from_slice(&7u32.to_le_bytes());
        bytes.extend_from_slice(&0u32.to_le_bytes());
        bytes.extend_from_slice(&2u32.to_le_bytes());
        bytes.extend_from_slice(&0x0101u16.to_le_bytes());
        // dxaLeft and type, followed by a future option byte.
        bytes.extend_from_slice(&7u16.to_le_bytes());
        bytes.extend_from_slice(&(-120i32).to_le_bytes());
        bytes.extend_from_slice(&1u16.to_le_bytes());
        bytes.push(0xA5);

        let array = PgpArray::parse_bytes(&bytes).unwrap();
        assert_eq!(array.len(), 1);
        assert_eq!(array.entries()[0].ipgp_self(), 7);
        assert_eq!(array.entries()[0].options().dxa_left(), Some(-120));
        assert_eq!(
            array.entries()[0].options().kind(),
            Some(PgpType::BlockQuote)
        );
        assert_eq!(array.entries()[0].options().opaque(), &[0xA5]);
        assert_eq!(array.bytes(), bytes.as_slice());
    }

    #[test]
    fn rejects_invalid_identifiers_types_and_option_lengths() {
        let mut zero_id = vec![1, 0];
        zero_id.extend_from_slice(&[0; PGP_INFO_HEADER_SIZE]);
        assert!(PgpArray::parse_bytes(&zero_id).is_err());

        let mut invalid_type = vec![1, 0];
        invalid_type.extend_from_slice(&1u32.to_le_bytes());
        invalid_type.extend_from_slice(&[0; 8]);
        invalid_type.extend_from_slice(&0x0100u16.to_le_bytes());
        invalid_type.extend_from_slice(&2u16.to_le_bytes());
        invalid_type.extend_from_slice(&[0xFF, 0xFF]);
        assert!(PgpArray::parse_bytes(&invalid_type).is_err());

        let mut short = vec![1, 0];
        short.extend_from_slice(&1u32.to_le_bytes());
        short.extend_from_slice(&[0; 8]);
        short.extend_from_slice(&1u16.to_le_bytes());
        short.extend_from_slice(&4u16.to_le_bytes());
        short.extend_from_slice(&[0; 2]);
        assert!(PgpArray::parse_bytes(&short).is_err());
    }

    #[test]
    fn validates_identifier_referents_and_acyclic_ancestry() {
        fn entry(bytes: &mut Vec<u8>, identifier: u32, parent: u32) {
            bytes.extend_from_slice(&identifier.to_le_bytes());
            bytes.extend_from_slice(&parent.to_le_bytes());
            bytes.extend_from_slice(&0u32.to_le_bytes());
            bytes.extend_from_slice(&0u16.to_le_bytes());
        }

        let mut duplicate = vec![2, 0];
        entry(&mut duplicate, 1, 0);
        entry(&mut duplicate, 1, 0);
        assert!(PgpArray::parse_bytes(&duplicate).is_err());

        let mut missing_parent = vec![1, 0];
        entry(&mut missing_parent, 1, 99);
        assert!(PgpArray::parse_bytes(&missing_parent).is_err());

        let mut cycle = vec![2, 0];
        entry(&mut cycle, 1, 2);
        entry(&mut cycle, 2, 1);
        assert!(PgpArray::parse_bytes(&cycle).is_err());

        // Parents may be serialized after their children, provided every
        // referent exists and the resulting ancestry is acyclic.
        let mut out_of_order = vec![2, 0];
        entry(&mut out_of_order, 2, 0);
        entry(&mut out_of_order, 1, 2);
        let parsed = PgpArray::parse_bytes(&out_of_order).unwrap();
        assert_eq!(parsed.entries()[1].ipgp_parent(), 2);
    }

    #[test]
    fn parses_only_the_declared_fib_table_range() {
        let bytes = [0u8, 0u8];
        let offset = 5u32;
        let index = FIB_INDEX_PGP;
        let mut fib_data = vec![0; 154 + (index + 1) * 8];
        fib_data[0..2].copy_from_slice(&0xA5ECu16.to_le_bytes());
        fib_data[2..4].copy_from_slice(&0x00C1u16.to_le_bytes());
        fib_data[152..154].copy_from_slice(&((index + 1) as u16).to_le_bytes());
        let pointer = 154 + index * 8;
        fib_data[pointer..pointer + 4].copy_from_slice(&offset.to_le_bytes());
        fib_data[pointer + 4..pointer + 8].copy_from_slice(&(bytes.len() as u32).to_le_bytes());
        let fib = FileInformationBlock::parse(&fib_data).unwrap();
        let mut table = vec![0xCC; 16];
        table[offset as usize..offset as usize + bytes.len()].copy_from_slice(&bytes);
        let parsed = PgpArray::parse(&fib, &table).unwrap().unwrap();
        assert!(parsed.is_empty());

        table.truncate(offset as usize + bytes.len() - 1);
        assert!(PgpArray::parse(&fib, &table).is_err());
    }
}
