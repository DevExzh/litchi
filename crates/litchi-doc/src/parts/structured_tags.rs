//! Word 2003 structured document tag (SDT) bookmark tables.
//!
//! These are the `SttbfBkmkSdt`, `PlcfBkfSdt`, and `PlcfBklSdt` tables
//! (MS-DOC 2.9.284, 2.8.10, and 2.9.239) that record the document ranges
//! wrapped by structured document tags, together with each tag's name, kind,
//! attributes, and placeholder text.
//!
//! All structures are parsed as inert metadata: no XML schema is resolved,
//! no placeholder is rendered, and no document content is modified.

use super::fib::FileInformationBlock;
use crate::package::{Error as PackageError, Result};
use std::collections::HashSet;

/// Table-pointer index of `fcSttbfBkmkSdt`/`lcbSttbfBkmkSdt`.
const STTBF_BKMK_SDT: usize = 137;
/// Table-pointer index of `fcPlcfBkfSdt`/`lcbPlcfBkfSdt`.
const PLCF_BKF_SDT: usize = 138;
/// Table-pointer index of `fcPlcfBklSdt`/`lcbPlcfBklSdt`.
const PLCF_BKL_SDT: usize = 139;

/// `fExtend` value of an extended STTB.
const STTB_F_EXTEND: u16 = 0xFFFF;
/// Maximum number of structured document tag bookmarks in a document
/// (`SttbfBkmkSdt` `cData` limit, MS-DOC 2.9.284).
const MAX_TAGS: u32 = 0x7FFF_FFFF;
/// `cchData` of every `SttbfBkmkSdt` string (MS-DOC 2.9.284).
const SDTI_FIXED_CHARS: u16 = 0x000C;
/// Fixed-size prefix of one `SDTI`: `dwId`, `tiq`, `sdtt`, `cfsdap`, and
/// `cbPlaceholder` (MS-DOC 2.9.239).
const SDTI_FIXED_SIZE: usize = 24;
/// Minimum size of one `FSDAP`: an 8-byte `TIQ`, a 2-byte count, and a
/// null terminator (MS-DOC 2.9.96).
const MIN_FSDAP_SIZE: usize = 12;
/// Minimum size of one `SttbfBkmkSdt` entry: the 2-byte `cchData`, the
/// fixed `SDTI` prefix, and a null placeholder terminator.
const MIN_ENTRY_SIZE: usize = 2 + SDTI_FIXED_SIZE + 2;
/// Maximum `ixsdr` value of a `TIQ` (MS-DOC 2.9.325).
const MAX_SCHEMA_INDEX: u32 = 0x7FFF_FFFF;

/// `BKC` flag bit `fPub`, which must be zero.
const BKC_F_PUB: u16 = 0x0080;
/// `BKC` flag bit `fNative`.
const BKC_F_NATIVE: u16 = 0x4000;
/// `BKC` flag bit `fCol`.
const BKC_F_COL: u16 = 0x8000;
/// `BKC` mask of `itcFirst`.
const BKC_ITC_FIRST_MASK: u16 = 0x007F;
/// `BKC` shift and mask of `itcLim`.
const BKC_ITC_LIM_SHIFT: u16 = 8;
const BKC_ITC_LIM_MASK: u16 = 0x003F;

/// A structured document tag node or attribute name reference (`TIQ`,
/// MS-DOC 2.9.325).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct StructuredTagName {
    /// Zero-based index of the `XSDR` within the `Hplxsdr` that namespaces
    /// this name.
    pub schema_index: u32,
    /// Zero-based index of the name string within the namespace string table.
    pub name_index: u32,
}

/// The type of structured document tag a bookmark represents (`SDTT`,
/// MS-DOC 2.9.240).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum StructuredTagKind {
    /// The tag encloses a range of characters.
    Range,
    /// The tag encloses a range of paragraphs.
    Paragraphs,
    /// The tag encloses a range of cells in a table.
    TableCells,
    /// The tag encloses a range of rows in a table.
    TableRows,
}

/// One attribute of a structured document tag (`FSDAP`, MS-DOC 2.9.96).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct StructuredTagAttribute {
    /// The attribute name reference.
    pub name: StructuredTagName,
    /// The attribute value.
    pub value: String,
}

/// The `SDTI` stored parallel to one structured document tag bookmark
/// (MS-DOC 2.9.239).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct StructuredTagInfo {
    /// Unique nonzero identifier of the tag within the document.
    pub id: u32,
    /// The tag name reference.
    pub name: StructuredTagName,
    /// The kind of range the tag encloses.
    pub kind: StructuredTagKind,
    /// Attributes carried by the tag.
    pub attributes: Vec<StructuredTagAttribute>,
    /// Text to show when the tag is empty and XML tag characters are hidden.
    pub placeholder: String,
}

/// A validated structured document tag bookmark.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DocumentStructuredTag {
    /// Start CP of the tagged range.
    pub start: u32,
    /// End CP of the tagged range.
    pub end: u32,
    /// Overlap depth recorded at the start of the range.
    pub start_depth: u16,
    /// Overlap depth recorded at the end of the range.
    pub end_depth: u16,
    /// Whether the bookmark was created by Word rather than a producer.
    pub is_native: bool,
    /// Table-column range, when the tag is confined to columns.
    pub column_range: Option<(u8, u8)>,
    /// The tag data stored parallel to the bookmark.
    pub info: StructuredTagInfo,
}

/// The structured document tags of a document, in start-CP order.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DocumentStructuredTags {
    tags: Vec<DocumentStructuredTag>,
}

impl DocumentStructuredTags {
    /// Parse the three parallel structured document tag tables addressed by
    /// the FIB, or `None` when the document carries none of them.
    pub fn parse(
        fib: &FileInformationBlock,
        table_stream: &[u8],
    ) -> Result<Option<DocumentStructuredTags>> {
        let lengths = [STTBF_BKMK_SDT, PLCF_BKF_SDT, PLCF_BKL_SDT]
            .map(|index| fib.get_table_pointer(index).map_or(0, |(_, length)| length));
        if lengths.iter().all(|&length| length == 0) {
            return Ok(None);
        }
        if lengths.contains(&0) {
            return Err(corrupted(
                "the three parallel structured document tag tables must be present together",
            ));
        }

        let infos = parse_infos(required_slice(
            fib,
            table_stream,
            STTBF_BKMK_SDT,
            "SttbfBkmkSdt",
        )?)?;
        let starts = parse_start_plcf(required_slice(
            fib,
            table_stream,
            PLCF_BKF_SDT,
            "PlcfBkfSdt",
        )?)?;
        let ends = parse_end_plcf(required_slice(
            fib,
            table_stream,
            PLCF_BKL_SDT,
            "PlcfBklSdt",
        )?)?;
        if infos.len() != starts.len() || starts.len() != ends.len() {
            return Err(corrupted(
                "structured document tag info, start, and end table counts do not match",
            ));
        }

        let document_end = fib
            .get_document_parts_end()
            .ok_or_else(|| corrupted("document-part character counts overflow"))?;
        validate_positions(&starts, document_end, "PlcfBkfSdt")?;
        validate_positions(&ends, document_end, "PlcfBklSdt")?;

        let mut used_end_indexes = HashSet::with_capacity(starts.len());
        let mut tags = Vec::with_capacity(starts.len());
        for (start_index, ((start, start_data), info)) in starts.iter().zip(&infos).enumerate() {
            let end_index = usize::from(start_data.end_index);
            if end_index >= ends.len() || !used_end_indexes.insert(end_index) {
                return Err(corrupted(
                    "structured document tag start end-index values must be unique and in range",
                ));
            }
            let (end, end_data) = &ends[end_index];
            if usize::from(end_data.start_index) != start_index {
                return Err(corrupted(
                    "structured document tag start and end bookmark indexes are not reciprocal",
                ));
            }
            if start > end {
                return Err(corrupted(
                    "structured document tag start CP exceeds its end CP",
                ));
            }
            tags.push(DocumentStructuredTag {
                start: *start,
                end: *end,
                start_depth: start_data.depth,
                end_depth: end_data.depth,
                is_native: start_data.is_native,
                column_range: start_data.column_range,
                info: info.clone(),
            });
        }

        Ok(Some(Self { tags }))
    }

    /// All structured document tags in start-CP order.
    pub fn tags(&self) -> &[DocumentStructuredTag] {
        &self.tags
    }
}

#[derive(Debug)]
struct StartData {
    end_index: u16,
    depth: u16,
    is_native: bool,
    column_range: Option<(u8, u8)>,
}

#[derive(Debug)]
struct EndData {
    start_index: u16,
    depth: u16,
}

/// Parse `SttbfBkmkSdt` (MS-DOC 2.9.284): an extended STTB whose strings are
/// `SDTI` structures (MS-DOC 2.9.239).
fn parse_infos(data: &[u8]) -> Result<Vec<StructuredTagInfo>> {
    if data.len() < 8
        || read_u16(data, 0, "SttbfBkmkSdt fExtend")? != STTB_F_EXTEND
        || read_u16(data, 6, "SttbfBkmkSdt cbExtra")? != 0
    {
        return Err(corrupted("SttbfBkmkSdt has an invalid header"));
    }
    let count = read_u32(data, 2, "SttbfBkmkSdt cData")?;
    if count > MAX_TAGS {
        return Err(corrupted("SttbfBkmkSdt contains too many entries"));
    }
    let count =
        usize::try_from(count).map_err(|_| corrupted("SttbfBkmkSdt count exceeds usize"))?;
    if count > (data.len() - 8) / MIN_ENTRY_SIZE {
        return Err(corrupted(
            "SttbfBkmkSdt byte length does not match its count",
        ));
    }
    let mut infos = Vec::with_capacity(count);
    let mut ids = HashSet::with_capacity(count);
    let mut offset = 8usize;
    for _ in 0..count {
        if read_u16(data, offset, "SttbfBkmkSdt cchData")? != SDTI_FIXED_CHARS {
            return Err(corrupted(
                "SttbfBkmkSdt entries must declare twelve UTF-16 code units",
            ));
        }
        let (info, size) = parse_sdti(&data[offset + 2..])?;
        if !ids.insert(info.id) {
            return Err(corrupted("SDTI ids must be unique"));
        }
        infos.push(info);
        offset += 2 + size;
    }
    if offset != data.len() {
        return Err(corrupted("SttbfBkmkSdt contains trailing bytes"));
    }
    Ok(infos)
}

/// Parse one `SDTI`, returning the tag info and the consumed byte count.
fn parse_sdti(data: &[u8]) -> Result<(StructuredTagInfo, usize)> {
    if data.len() < SDTI_FIXED_SIZE {
        return Err(corrupted("SDTI is truncated"));
    }
    let id = read_u32(data, 0, "SDTI dwId")?;
    if id == 0 {
        return Err(corrupted("SDTI dwId must be nonzero"));
    }
    let name = parse_tiq(&data[4..12], "SDTI")?;
    let kind = match read_u32(data, 12, "SDTI sdtt")? {
        1 => StructuredTagKind::Range,
        2 => StructuredTagKind::Paragraphs,
        3 => StructuredTagKind::TableCells,
        4 => StructuredTagKind::TableRows,
        _ => return Err(corrupted("SDTI contains an invalid sdtt")),
    };
    let attribute_count = usize::try_from(read_u32(data, 16, "SDTI cfsdap")?)
        .map_err(|_| corrupted("SDTI cfsdap exceeds usize"))?;
    let placeholder_bytes = usize::try_from(read_u32(data, 20, "SDTI cbPlaceholder")?)
        .map_err(|_| corrupted("SDTI cbPlaceholder exceeds usize"))?;
    if placeholder_bytes < 2 || placeholder_bytes % 2 != 0 {
        return Err(corrupted(
            "SDTI cbPlaceholder must count a nonempty UTF-16 string",
        ));
    }
    if attribute_count > (data.len() - SDTI_FIXED_SIZE) / MIN_FSDAP_SIZE {
        return Err(corrupted("SDTI fsdaparray is truncated"));
    }

    let mut attributes = Vec::with_capacity(attribute_count);
    let mut offset = SDTI_FIXED_SIZE;
    for _ in 0..attribute_count {
        let name = parse_tiq(&data[offset..offset + 8], "FSDAP")?;
        let value_chars = usize::from(read_u16(data, offset + 8, "FSDAP cch")?);
        let value_start = offset + 10;
        let value_end = value_start
            .checked_add((value_chars + 1) * 2)
            .ok_or_else(|| corrupted("FSDAP value range overflows"))?;
        let value = parse_terminated_string(data, value_start, value_end, "FSDAP rgValue")?;
        attributes.push(StructuredTagAttribute { name, value });
        offset = value_end;
    }

    let placeholder_end = offset
        .checked_add(placeholder_bytes)
        .ok_or_else(|| corrupted("SDTI placeholder range overflows"))?;
    let placeholder =
        parse_terminated_string(data, offset, placeholder_end, "SDTI xszPlaceholder")?;

    Ok((
        StructuredTagInfo {
            id,
            name,
            kind,
            attributes,
            placeholder,
        },
        placeholder_end,
    ))
}

fn parse_tiq(data: &[u8], parent: &str) -> Result<StructuredTagName> {
    let schema_index = read_u32(data, 0, "TIQ ixsdr")?;
    if schema_index >= MAX_SCHEMA_INDEX {
        return Err(corrupted(format!("{parent} TIQ ixsdr is out of range")));
    }
    Ok(StructuredTagName {
        schema_index,
        name_index: read_u32(data, 4, "TIQ ixstElement")?,
    })
}

/// Decode a null-terminated UTF-16 string occupying `data[start..end]`,
/// returning the content without its terminator.
fn parse_terminated_string(data: &[u8], start: usize, end: usize, field: &str) -> Result<String> {
    let bytes = data
        .get(start..end)
        .ok_or_else(|| corrupted(format!("{field} is truncated")))?;
    let units = bytes
        .chunks_exact(2)
        .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
        .collect::<Vec<_>>();
    if units.last().copied() != Some(0) {
        return Err(corrupted(format!("{field} must be null-terminated")));
    }
    String::from_utf16(&units[..units.len() - 1])
        .map_err(|_| corrupted(format!("{field} is invalid UTF-16")))
}

/// Parse `PlcfBkfSdt`: a `Plcfbkfd` whose 6-byte data elements are `FBKFD`
/// structures (MS-DOC 2.9.71).
fn parse_start_plcf(data: &[u8]) -> Result<Vec<(u32, StartData)>> {
    let count = plcf_count(data, 6, "PlcfBkfSdt")?;
    let properties = (count + 1)
        .checked_mul(4)
        .ok_or_else(|| corrupted("PlcfBkfSdt position bytes overflow"))?;
    let mut values = Vec::with_capacity(count);
    for index in 0..count {
        let property = properties + index * 6;
        let bkc = read_u16(data, property + 2, "structured document tag BKC")?;
        if bkc & BKC_F_PUB != 0 {
            return Err(corrupted("structured document tag BKC fPub must be zero"));
        }
        let column_range = if bkc & BKC_F_COL != 0 {
            let first = (bkc & BKC_ITC_FIRST_MASK) as u8;
            let limit = ((bkc >> BKC_ITC_LIM_SHIFT) & BKC_ITC_LIM_MASK) as u8;
            if first >= limit {
                return Err(corrupted(
                    "structured document tag BKC column range is empty or reversed",
                ));
            }
            Some((first, limit))
        } else {
            None
        };
        values.push((
            read_u32(data, index * 4, "structured document tag start CP")?,
            StartData {
                end_index: read_u16(data, property, "structured document tag ibkl")?,
                depth: read_u16(data, property + 4, "structured document tag start depth")?,
                is_native: bkc & BKC_F_NATIVE != 0,
                column_range,
            },
        ));
    }
    Ok(values)
}

/// Parse `PlcfBklSdt`: a `Plcfbkld` whose 4-byte data elements are `FBKLD`
/// structures (MS-DOC 2.9.72).
fn parse_end_plcf(data: &[u8]) -> Result<Vec<(u32, EndData)>> {
    let count = plcf_count(data, 4, "PlcfBklSdt")?;
    let properties = (count + 1)
        .checked_mul(4)
        .ok_or_else(|| corrupted("PlcfBklSdt position bytes overflow"))?;
    let mut values = Vec::with_capacity(count);
    for index in 0..count {
        let property = properties + index * 4;
        values.push((
            read_u32(data, index * 4, "structured document tag end CP")?,
            EndData {
                start_index: read_u16(data, property, "structured document tag ibkf")?,
                depth: read_u16(data, property + 2, "structured document tag end depth")?,
            },
        ));
    }
    Ok(values)
}

fn plcf_count(data: &[u8], property_size: usize, name: &str) -> Result<usize> {
    if data.len() < 4 || !(data.len() - 4).is_multiple_of(4 + property_size) {
        return Err(corrupted(format!("{name} has an invalid byte length")));
    }
    Ok((data.len() - 4) / (4 + property_size))
}

fn validate_positions<T>(values: &[(u32, T)], document_end: u32, name: &str) -> Result<()> {
    if values.iter().any(|(cp, _)| *cp > document_end)
        || values.windows(2).any(|pair| pair[0].0 > pair[1].0)
    {
        return Err(corrupted(format!(
            "{name} contains out-of-range or non-monotonic CPs"
        )));
    }
    Ok(())
}

fn required_slice<'a>(
    fib: &FileInformationBlock,
    table_stream: &'a [u8],
    index: usize,
    name: &str,
) -> Result<&'a [u8]> {
    let (offset, length) = fib
        .get_table_pointer(index)
        .filter(|(_, length)| *length != 0)
        .ok_or_else(|| corrupted(format!("{name} is missing")))?;
    let start =
        usize::try_from(offset).map_err(|_| corrupted(format!("{name} offset exceeds usize")))?;
    let length =
        usize::try_from(length).map_err(|_| corrupted(format!("{name} length exceeds usize")))?;
    let end = start
        .checked_add(length)
        .ok_or_else(|| corrupted(format!("{name} range overflows")))?;
    table_stream
        .get(start..end)
        .ok_or_else(|| corrupted(format!("{name} extends beyond the table stream")))
}

fn read_u16(data: &[u8], offset: usize, field: &str) -> Result<u16> {
    litchi_core::binary::read_u16_le(data, offset)
        .map_err(|error| corrupted(format!("invalid {field}: {error}")))
}

fn read_u32(data: &[u8], offset: usize, field: &str) -> Result<u32> {
    litchi_core::binary::read_u32_le(data, offset)
        .map_err(|error| corrupted(format!("invalid {field}: {error}")))
}

fn corrupted(message: impl Into<String>) -> PackageError {
    PackageError::Corrupted(message.into())
}

#[cfg(test)]
mod tests {
    use super::*;

    const FIB_POINTERS: usize = 145;

    fn set_fib_pointer(fib: &mut [u8], index: usize, offset: u32, length: u32) {
        let declared = u16::from_le_bytes([fib[152], fib[153]]);
        let count = declared.max(u16::try_from(index + 1).unwrap());
        fib[152..154].copy_from_slice(&count.to_le_bytes());
        let start = 154 + index * 8;
        fib[start..start + 4].copy_from_slice(&offset.to_le_bytes());
        fib[start + 4..start + 8].copy_from_slice(&length.to_le_bytes());
    }

    fn sdti(
        id: u32,
        name: (u32, u32),
        kind: u32,
        attributes: &[((u32, u32), &str)],
        placeholder: &str,
    ) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&id.to_le_bytes());
        data.extend_from_slice(&name.0.to_le_bytes());
        data.extend_from_slice(&name.1.to_le_bytes());
        data.extend_from_slice(&kind.to_le_bytes());
        data.extend_from_slice(&(attributes.len() as u32).to_le_bytes());
        let placeholder_units = placeholder.encode_utf16().count() + 1;
        data.extend_from_slice(&((placeholder_units * 2) as u32).to_le_bytes());
        for (attribute_name, value) in attributes {
            data.extend_from_slice(&attribute_name.0.to_le_bytes());
            data.extend_from_slice(&attribute_name.1.to_le_bytes());
            let units = value.encode_utf16().collect::<Vec<_>>();
            data.extend_from_slice(&(units.len() as u16).to_le_bytes());
            for unit in units.iter().chain([&0]) {
                data.extend_from_slice(&unit.to_le_bytes());
            }
        }
        for unit in placeholder.encode_utf16().chain([0]) {
            data.extend_from_slice(&unit.to_le_bytes());
        }
        data
    }

    fn info_table(entries: &[Vec<u8>]) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&STTB_F_EXTEND.to_le_bytes());
        data.extend_from_slice(&(entries.len() as u32).to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes());
        for entry in entries {
            data.extend_from_slice(&SDTI_FIXED_CHARS.to_le_bytes());
            data.extend_from_slice(entry);
        }
        data
    }

    /// Build a FIB plus table stream holding two tags: a range tag at
    /// CP 1..8 and a paragraph tag at CP 2..5.
    fn fixture() -> (FileInformationBlock, Vec<u8>, usize) {
        let mut fib_data = vec![0; 154 + FIB_POINTERS * 8];
        fib_data[0..2].copy_from_slice(&0xA5ECu16.to_le_bytes());
        fib_data[2..4].copy_from_slice(&0x00C1u16.to_le_bytes());
        fib_data[76..80].copy_from_slice(&10u32.to_le_bytes());
        fib_data[152..154].copy_from_slice(&(FIB_POINTERS as u16).to_le_bytes());

        let mut table = Vec::new();
        let infos = info_table(&[
            sdti(5, (0, 2), 1, &[((1, 3), "ab")], "tip"),
            sdti(6, (0, 4), 2, &[], ""),
        ]);
        let infos_offset = table.len() as u32;
        table.extend_from_slice(&infos);
        set_fib_pointer(
            &mut fib_data,
            STTBF_BKMK_SDT,
            infos_offset,
            infos.len() as u32,
        );

        let starts_offset = table.len();
        for cp in [1u32, 2, 11] {
            table.extend_from_slice(&cp.to_le_bytes());
        }
        table.extend_from_slice(&1u16.to_le_bytes()); // ibkl
        table.extend_from_slice(&0x4000u16.to_le_bytes()); // bkc: fNative
        table.extend_from_slice(&0u16.to_le_bytes()); // cDepth
        table.extend_from_slice(&0u16.to_le_bytes()); // ibkl
        table.extend_from_slice(&0u16.to_le_bytes()); // bkc
        table.extend_from_slice(&1u16.to_le_bytes()); // cDepth
        set_fib_pointer(&mut fib_data, PLCF_BKF_SDT, starts_offset as u32, 24);

        let ends_offset = table.len() as u32;
        for cp in [5u32, 8, 11] {
            table.extend_from_slice(&cp.to_le_bytes());
        }
        table.extend_from_slice(&1u16.to_le_bytes()); // ibkf
        table.extend_from_slice(&1u16.to_le_bytes()); // cDepth
        table.extend_from_slice(&0u16.to_le_bytes()); // ibkf
        table.extend_from_slice(&0u16.to_le_bytes()); // cDepth
        set_fib_pointer(&mut fib_data, PLCF_BKL_SDT, ends_offset, 20);

        (
            FileInformationBlock::parse(&fib_data).unwrap(),
            table,
            starts_offset,
        )
    }

    #[test]
    fn parses_structured_document_tags() {
        let (fib, table, _) = fixture();
        let parsed = DocumentStructuredTags::parse(&fib, &table)
            .unwrap()
            .expect("tags present");
        assert_eq!(
            parsed.tags(),
            [
                DocumentStructuredTag {
                    start: 1,
                    end: 8,
                    start_depth: 0,
                    end_depth: 0,
                    is_native: true,
                    column_range: None,
                    info: StructuredTagInfo {
                        id: 5,
                        name: StructuredTagName {
                            schema_index: 0,
                            name_index: 2,
                        },
                        kind: StructuredTagKind::Range,
                        attributes: vec![StructuredTagAttribute {
                            name: StructuredTagName {
                                schema_index: 1,
                                name_index: 3,
                            },
                            value: "ab".to_string(),
                        }],
                        placeholder: "tip".to_string(),
                    },
                },
                DocumentStructuredTag {
                    start: 2,
                    end: 5,
                    start_depth: 1,
                    end_depth: 1,
                    is_native: false,
                    column_range: None,
                    info: StructuredTagInfo {
                        id: 6,
                        name: StructuredTagName {
                            schema_index: 0,
                            name_index: 4,
                        },
                        kind: StructuredTagKind::Paragraphs,
                        attributes: Vec::new(),
                        placeholder: String::new(),
                    },
                },
            ]
        );
    }

    #[test]
    fn absent_tables_yield_none() {
        let mut fib_data = vec![0; 154 + FIB_POINTERS * 8];
        fib_data[0..2].copy_from_slice(&0xA5ECu16.to_le_bytes());
        fib_data[2..4].copy_from_slice(&0x00C1u16.to_le_bytes());
        fib_data[76..80].copy_from_slice(&10u32.to_le_bytes());
        fib_data[152..154].copy_from_slice(&(FIB_POINTERS as u16).to_le_bytes());
        let fib = FileInformationBlock::parse(&fib_data).unwrap();
        assert!(DocumentStructuredTags::parse(&fib, &[]).unwrap().is_none());
    }

    #[test]
    fn rejects_partial_table_sets() {
        let (fib, table, _) = fixture();
        let mut fib_data = fib.raw_data().to_vec();
        set_fib_pointer(&mut fib_data, PLCF_BKL_SDT, 0, 0);
        let fib = FileInformationBlock::parse(&fib_data).unwrap();
        assert!(DocumentStructuredTags::parse(&fib, &table).is_err());
    }

    #[test]
    fn rejects_malformed_tables() {
        let (fib, table, starts_offset) = fixture();

        // Non-monotonic start CPs.
        let mut bad_starts = table.clone();
        bad_starts[starts_offset..starts_offset + 4].copy_from_slice(&11u32.to_le_bytes());
        assert!(DocumentStructuredTags::parse(&fib, &bad_starts).is_err());

        // Broken reciprocal indexes: the second end no longer points back at
        // the second start.
        let mut bad_ibkf = table.clone();
        let second_ibkf = table.len() - 4;
        bad_ibkf[second_ibkf..second_ibkf + 2].copy_from_slice(&1u16.to_le_bytes());
        assert!(DocumentStructuredTags::parse(&fib, &bad_ibkf).is_err());

        // Zero id.
        assert!(parse_infos(&info_table(&[sdti(0, (0, 2), 1, &[], "")])).is_err());

        // Duplicate ids.
        assert!(
            parse_infos(&info_table(&[
                sdti(5, (0, 2), 1, &[], ""),
                sdti(5, (0, 4), 2, &[], ""),
            ]))
            .is_err()
        );

        // sdttUnknown.
        assert!(parse_infos(&info_table(&[sdti(5, (0, 2), 0, &[], "")])).is_err());

        // Out-of-range schema index.
        assert!(parse_infos(&info_table(&[sdti(5, (0x7FFF_FFFF, 2), 1, &[], "")])).is_err());

        // Missing placeholder terminator.
        let mut unterminated = sdti(5, (0, 2), 1, &[], "");
        *unterminated.last_mut().unwrap() = 1;
        assert!(parse_infos(&info_table(&[unterminated])).is_err());

        // Trailing bytes after the last entry.
        let mut trailing = info_table(&[sdti(5, (0, 2), 1, &[], "")]);
        trailing.push(0);
        assert!(parse_infos(&trailing).is_err());

        // Declared count exceeds what the byte length can hold.
        let mut inflated = info_table(&[sdti(5, (0, 2), 1, &[], "")]);
        inflated[2..6].copy_from_slice(&2u32.to_le_bytes());
        assert!(parse_infos(&inflated).is_err());
    }
}
