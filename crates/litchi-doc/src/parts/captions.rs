//! Caption label tables (`SttbfCaption` and `SttbfAutoCaption`).
//!
//! `SttbfCaption` (MS-DOC 2.9.285) holds the caption labels a document offers;
//! each string carries a `CAPI` (MS-DOC 2.9.24) that describes how the caption
//! is inserted. `SttbfAutoCaption` (MS-DOC 2.9.278) maps OLE object ProgIDs to
//! caption labels so inserting such an object automatically adds a caption.
//! Both tables are metadata only: labels are never rendered and objects are
//! never activated.

use super::super::package::{DocError, Result};
use super::fib::FileInformationBlock;

/// Table-pointer index of `fcSttbfCaption`/`lcbSttbfCaption` (MS-DOC 2.5.6 FibRgFcLcb97).
const CAPTION_FIB_INDEX: usize = 52;
/// Table-pointer index of `fcSttbfAutoCaption`/`lcbSttbfAutoCaption` (MS-DOC 2.5.6 FibRgFcLcb97).
const AUTO_CAPTION_FIB_INDEX: usize = 53;
/// Serialized size of one `CAPI` extra-data record (MS-DOC 2.9.24).
const CAPI_SIZE: usize = 6;
/// Serialized size of one `SttbfAutoCaption` extra-data record (MS-DOC 2.9.278).
const AUTO_CAPTION_EXTRA_SIZE: u16 = 2;
/// A caption label MUST have at most 40 characters (MS-DOC 2.9.285).
const MAX_LABEL_UNITS: usize = 40;
/// Size cap mirroring the other STTB readers; far beyond any conforming table.
const MAX_TABLE_BYTES: usize = 16 * 1024 * 1024;

fn corrupted(message: impl Into<String>) -> DocError {
    DocError::Corrupted(message.into())
}

fn read_u16(data: &[u8], offset: usize, field: &str) -> Result<u16> {
    litchi_core::binary::read_u16_le(data, offset)
        .map_err(|error| corrupted(format!("invalid {field}: {error}")))
}

fn decode_utf16(data: &[u8], context: &str) -> Result<String> {
    char::decode_utf16(
        data.chunks_exact(2)
            .map(|unit| u16::from_le_bytes([unit[0], unit[1]])),
    )
    .collect::<std::result::Result<String, _>>()
    .map_err(|error| corrupted(format!("invalid {context}: {error}")))
}

/// Insert location for a caption (`CAPI.iLocation`, MS-DOC 2.9.24).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum CaptionLocation {
    /// Insert the caption below the selected item.
    Below = 0x0,
    /// Insert the caption above the selected item.
    Above = 0x1,
}

impl CaptionLocation {
    fn from_raw(value: u8) -> Result<Self> {
        match value {
            0x0 => Ok(Self::Below),
            0x1 => Ok(Self::Above),
            _ => Err(corrupted(format!("invalid caption location 0x{value:X}"))),
        }
    }
}

/// Heading style that starts a new chapter (`CAPI.iHeading`, MS-DOC 2.9.24).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum ChapterHeading {
    Heading1 = 0x1,
    Heading2 = 0x2,
    Heading3 = 0x3,
    Heading4 = 0x4,
    Heading5 = 0x5,
    Heading6 = 0x6,
    Heading7 = 0x7,
    Heading8 = 0x8,
    Heading9 = 0x9,
}

impl ChapterHeading {
    fn from_raw(value: u8) -> Result<Self> {
        match value {
            0x1 => Ok(Self::Heading1),
            0x2 => Ok(Self::Heading2),
            0x3 => Ok(Self::Heading3),
            0x4 => Ok(Self::Heading4),
            0x5 => Ok(Self::Heading5),
            0x6 => Ok(Self::Heading6),
            0x7 => Ok(Self::Heading7),
            0x8 => Ok(Self::Heading8),
            0x9 => Ok(Self::Heading9),
            _ => Err(corrupted(format!("invalid chapter heading 0x{value:X}"))),
        }
    }
}

/// Character separating the chapter and caption numbers (`CAPI.xchSeparator`, MS-DOC 2.9.24).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u16)]
pub enum ChapterSeparator {
    Hyphen = 0x001E,
    Period = 0x002E,
    Colon = 0x003A,
    EnDash = 0x2013,
    EmDash = 0x2014,
}

impl ChapterSeparator {
    fn from_raw(value: u16) -> Result<Self> {
        match value {
            0x001E => Ok(Self::Hyphen),
            0x002E => Ok(Self::Period),
            0x003A => Ok(Self::Colon),
            0x2013 => Ok(Self::EnDash),
            0x2014 => Ok(Self::EmDash),
            _ => Err(corrupted(format!(
                "invalid chapter number separator 0x{value:04X}"
            ))),
        }
    }
}

/// Chapter-numbering settings used when a caption includes a chapter number
/// (`CAPI.fChapNum`, `CAPI.iHeading`, and `CAPI.xchSeparator`, MS-DOC 2.9.24).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ChapterNumbering {
    heading: ChapterHeading,
    separator: ChapterSeparator,
}

impl ChapterNumbering {
    pub const fn new(heading: ChapterHeading, separator: ChapterSeparator) -> Self {
        Self { heading, separator }
    }

    /// Heading style that marks the beginning of a new chapter.
    pub fn heading(&self) -> ChapterHeading {
        self.heading
    }
    /// Character between the chapter number and the caption number.
    pub fn separator(&self) -> ChapterSeparator {
        self.separator
    }
}

/// Caption insertion metadata (`CAPI`, MS-DOC 2.9.24; 6 bytes).
///
/// The `unused1` bit field is undefined in the format; it is ignored when
/// reading and written as zero.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct CaptionInfo {
    location: CaptionLocation,
    chapter_numbering: Option<ChapterNumbering>,
    omit_label: bool,
    number_format: u16,
}

impl CaptionInfo {
    /// Serialized size of one `CAPI` (MS-DOC 2.9.24).
    pub const SIZE: usize = CAPI_SIZE;

    /// Create caption metadata. `number_format` is an MSONFC value as
    /// specified in MS-OSHARED 2.2.1.3 and is kept opaque.
    pub const fn new(
        location: CaptionLocation,
        chapter_numbering: Option<ChapterNumbering>,
        omit_label: bool,
        number_format: u16,
    ) -> Self {
        Self {
            location,
            chapter_numbering,
            omit_label,
            number_format,
        }
    }

    /// Decode one 6-byte `CAPI`. When `fChapNum` is zero the format ignores
    /// `iHeading` and `xchSeparator`; both are dropped on read.
    pub fn from_bytes(data: &[u8]) -> Result<Self> {
        if data.len() != Self::SIZE {
            return Err(corrupted("CAPI must be exactly 6 bytes"));
        }
        let flags = read_u16(data, 0, "CAPI flags")?;
        let chapter_numbering = if flags & 0x4 != 0 {
            Some(ChapterNumbering::new(
                ChapterHeading::from_raw(((flags >> 3) & 0xF) as u8)?,
                ChapterSeparator::from_raw(read_u16(data, 4, "CAPI xchSeparator")?)?,
            ))
        } else {
            None
        };
        Ok(Self {
            location: CaptionLocation::from_raw((flags & 0x3) as u8)?,
            chapter_numbering,
            omit_label: flags & 0x8000 != 0,
            number_format: read_u16(data, 2, "CAPI nfc")?,
        })
    }

    /// Serialize with zeroed undefined bits.
    pub fn to_bytes(self) -> [u8; Self::SIZE] {
        let flags = self.location as u16
            | u16::from(self.chapter_numbering.is_some()) << 2
            | self
                .chapter_numbering
                .map_or(0, |chapter| (chapter.heading as u16) << 3)
            | u16::from(self.omit_label) << 15;
        let separator = self
            .chapter_numbering
            .map_or(0, |chapter| chapter.separator as u16);
        let mut data = [0u8; Self::SIZE];
        data[0..2].copy_from_slice(&flags.to_le_bytes());
        data[2..4].copy_from_slice(&self.number_format.to_le_bytes());
        data[4..6].copy_from_slice(&separator.to_le_bytes());
        data
    }

    /// Where the caption is inserted relative to the selected item.
    pub fn location(&self) -> CaptionLocation {
        self.location
    }
    /// Chapter-numbering settings, present iff the caption includes a chapter number.
    pub fn chapter_numbering(&self) -> Option<ChapterNumbering> {
        self.chapter_numbering
    }
    /// Whether the label is excluded from the caption. Producers may ignore this bit.
    pub fn omit_label(&self) -> bool {
        self.omit_label
    }
    /// MSONFC formatting of the caption number (MS-OSHARED 2.2.1.3), kept opaque.
    pub fn number_format(&self) -> u16 {
        self.number_format
    }
}

/// One caption label and its insertion metadata.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CaptionDefinition {
    label: String,
    info: CaptionInfo,
}

impl CaptionDefinition {
    /// Create a caption definition, validating the 40-character label cap.
    pub fn try_new(label: String, info: CaptionInfo) -> Result<Self> {
        if label.encode_utf16().count() > MAX_LABEL_UNITS {
            return Err(corrupted("caption label exceeds 40 UTF-16 code units"));
        }
        Ok(Self { label, info })
    }

    /// The caption label text.
    pub fn label(&self) -> &str {
        &self.label
    }
    /// Insertion metadata for this caption.
    pub fn info(&self) -> &CaptionInfo {
        &self.info
    }
}

/// One `SttbfAutoCaption` entry: an OLE object ProgID and the zero-based
/// `SttbfCaption` index of the caption inserted with it (MS-DOC 2.9.278).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AutoCaptionEntry {
    prog_id: String,
    caption_index: u16,
}

impl AutoCaptionEntry {
    pub const fn new(prog_id: String, caption_index: u16) -> Self {
        Self {
            prog_id,
            caption_index,
        }
    }

    /// ProgID of the OLE object that automatically receives a caption.
    pub fn prog_id(&self) -> &str {
        &self.prog_id
    }
    /// Zero-based index into `SttbfCaption` selecting the inserted caption.
    pub fn caption_index(&self) -> u16 {
        self.caption_index
    }
}

/// Validate and parse the shared extended-STTB header (MS-DOC 2.9.271),
/// returning the declared string count.
fn parse_sttb_header(data: &[u8], extra: u16, name: &str) -> Result<usize> {
    if data.len() > MAX_TABLE_BYTES {
        return Err(corrupted(format!("{name} exceeds the table size cap")));
    }
    if data.len() < 6
        || read_u16(data, 0, &format!("{name} fExtend"))? != u16::MAX
        || read_u16(data, 4, &format!("{name} cbExtra"))? != extra
    {
        return Err(corrupted(format!("{name} has an invalid header")));
    }
    Ok(usize::from(read_u16(data, 2, &format!("{name} cData"))?))
}

/// Read one length-prefixed UTF-16 string, returning it and the next offset.
fn parse_string(
    data: &[u8],
    offset: usize,
    max_units: usize,
    table: &str,
    index: usize,
) -> Result<(String, usize)> {
    let units = usize::from(read_u16(
        data,
        offset,
        &format!("{table} string {index} length"),
    )?);
    if units > max_units {
        return Err(corrupted(format!(
            "{table} string {index} exceeds {max_units} UTF-16 code units"
        )));
    }
    let start = offset
        .checked_add(2)
        .ok_or_else(|| corrupted(format!("{table} string offset overflows")))?;
    let end = start
        .checked_add(
            units
                .checked_mul(2)
                .ok_or_else(|| corrupted(format!("{table} string size overflows")))?,
        )
        .ok_or_else(|| corrupted(format!("{table} string range overflows")))?;
    let bytes = data
        .get(start..end)
        .ok_or_else(|| corrupted(format!("{table} string {index} is truncated")))?;
    Ok((
        decode_utf16(bytes, &format!("{table} string {index}"))?,
        end,
    ))
}

fn write_string(out: &mut Vec<u8>, value: &str) {
    let units: Vec<u16> = value.encode_utf16().collect();
    out.extend_from_slice(&(units.len() as u16).to_le_bytes());
    for unit in units {
        out.extend_from_slice(&unit.to_le_bytes());
    }
}

/// A typed `SttbfCaption` (MS-DOC 2.9.285): the document's caption labels.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct CaptionLabelTable {
    definitions: Vec<CaptionDefinition>,
}

impl CaptionLabelTable {
    /// Create a label table, validating entry count and label lengths.
    pub fn try_new(definitions: Vec<CaptionDefinition>) -> Result<Self> {
        if definitions.len() > u16::MAX as usize {
            return Err(corrupted("SttbfCaption count exceeds 65535 entries"));
        }
        Ok(Self { definitions })
    }

    /// Parse a complete `SttbfCaption` from the Table Stream.
    pub fn parse_bytes(data: &[u8]) -> Result<Self> {
        let count = parse_sttb_header(data, CAPI_SIZE as u16, "SttbfCaption")?;
        let mut definitions = Vec::with_capacity(count);
        let mut offset = 6usize;
        for index in 0..count {
            let (label, next) = parse_string(data, offset, MAX_LABEL_UNITS, "SttbfCaption", index)?;
            let extra = data.get(next..next + CAPI_SIZE).ok_or_else(|| {
                corrupted(format!("SttbfCaption entry {index} CAPI is truncated"))
            })?;
            definitions.push(CaptionDefinition::try_new(
                label,
                CaptionInfo::from_bytes(extra)?,
            )?);
            offset = next + CAPI_SIZE;
        }
        if offset != data.len() {
            return Err(corrupted("SttbfCaption has trailing bytes"));
        }
        Self::try_new(definitions)
    }

    pub fn definitions(&self) -> &[CaptionDefinition] {
        &self.definitions
    }
    pub fn len(&self) -> usize {
        self.definitions.len()
    }
    pub fn is_empty(&self) -> bool {
        self.definitions.is_empty()
    }

    /// Serialize the complete STTB deterministically.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        if self.definitions.len() > u16::MAX as usize {
            return Err(corrupted("SttbfCaption count exceeds 65535 entries"));
        }
        let mut data = Vec::new();
        data.extend_from_slice(&u16::MAX.to_le_bytes());
        data.extend_from_slice(&(self.definitions.len() as u16).to_le_bytes());
        data.extend_from_slice(&(CAPI_SIZE as u16).to_le_bytes());
        for definition in &self.definitions {
            if definition.label.encode_utf16().count() > MAX_LABEL_UNITS {
                return Err(corrupted("caption label exceeds 40 UTF-16 code units"));
            }
            write_string(&mut data, &definition.label);
            data.extend_from_slice(&definition.info.to_bytes());
        }
        Ok(data)
    }
}

/// A typed `SttbfAutoCaption` (MS-DOC 2.9.278): ProgID to caption mappings.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct AutoCaptionTable {
    entries: Vec<AutoCaptionEntry>,
}

impl AutoCaptionTable {
    /// Create an AutoCaption table, validating the entry count.
    pub fn try_new(entries: Vec<AutoCaptionEntry>) -> Result<Self> {
        if entries.len() > u16::MAX as usize {
            return Err(corrupted("SttbfAutoCaption count exceeds 65535 entries"));
        }
        Ok(Self { entries })
    }

    /// Parse a complete `SttbfAutoCaption` from the Table Stream.
    pub fn parse_bytes(data: &[u8]) -> Result<Self> {
        let count = parse_sttb_header(data, AUTO_CAPTION_EXTRA_SIZE, "SttbfAutoCaption")?;
        let mut entries = Vec::with_capacity(count);
        let mut offset = 6usize;
        for index in 0..count {
            let (prog_id, next) =
                parse_string(data, offset, u16::MAX as usize, "SttbfAutoCaption", index)?;
            let caption_index = read_u16(
                data,
                next,
                &format!("SttbfAutoCaption entry {index} caption index"),
            )?;
            entries.push(AutoCaptionEntry::new(prog_id, caption_index));
            offset = next + usize::from(AUTO_CAPTION_EXTRA_SIZE);
        }
        if offset != data.len() {
            return Err(corrupted("SttbfAutoCaption has trailing bytes"));
        }
        Self::try_new(entries)
    }

    pub fn entries(&self) -> &[AutoCaptionEntry] {
        &self.entries
    }
    pub fn len(&self) -> usize {
        self.entries.len()
    }
    pub fn is_empty(&self) -> bool {
        self.entries.is_empty()
    }

    /// Serialize the complete STTB deterministically.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        if self.entries.len() > u16::MAX as usize {
            return Err(corrupted("SttbfAutoCaption count exceeds 65535 entries"));
        }
        let mut data = Vec::new();
        data.extend_from_slice(&u16::MAX.to_le_bytes());
        data.extend_from_slice(&(self.entries.len() as u16).to_le_bytes());
        data.extend_from_slice(&AUTO_CAPTION_EXTRA_SIZE.to_le_bytes());
        for entry in &self.entries {
            write_string(&mut data, &entry.prog_id);
            data.extend_from_slice(&entry.caption_index.to_le_bytes());
        }
        Ok(data)
    }
}

/// Optional caption label and AutoCaption tables for a document.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct CaptionTables {
    captions: Option<CaptionLabelTable>,
    auto_captions: Option<AutoCaptionTable>,
}

impl CaptionTables {
    /// Parse both caption STTBs from the Table Stream.
    ///
    /// Every `SttbfAutoCaption` entry references a `SttbfCaption` index, so
    /// each index must fall inside the parsed label table.
    pub fn parse(fib: &FileInformationBlock, table_stream: &[u8]) -> Result<Self> {
        let captions = parse_fib_table(fib, table_stream, CAPTION_FIB_INDEX, "SttbfCaption")?
            .map(CaptionLabelTable::parse_bytes)
            .transpose()?;
        let auto_captions = parse_fib_table(
            fib,
            table_stream,
            AUTO_CAPTION_FIB_INDEX,
            "SttbfAutoCaption",
        )?
        .map(AutoCaptionTable::parse_bytes)
        .transpose()?;
        if let Some(auto_captions) = &auto_captions {
            let caption_count = captions.as_ref().map_or(0, CaptionLabelTable::len);
            for (index, entry) in auto_captions.entries().iter().enumerate() {
                if usize::from(entry.caption_index()) >= caption_count {
                    return Err(corrupted(format!(
                        "SttbfAutoCaption entry {index} references a missing SttbfCaption label"
                    )));
                }
            }
        }
        Ok(Self {
            captions,
            auto_captions,
        })
    }

    /// Caption labels (`SttbfCaption`, MS-DOC 2.9.285).
    pub fn captions(&self) -> Option<&CaptionLabelTable> {
        self.captions.as_ref()
    }
    /// AutoCaption ProgID mappings (`SttbfAutoCaption`, MS-DOC 2.9.278).
    pub fn auto_captions(&self) -> Option<&AutoCaptionTable> {
        self.auto_captions.as_ref()
    }
}

/// Slice out one FIB-referenced table, returning `None` when absent or empty.
fn parse_fib_table<'a>(
    fib: &FileInformationBlock,
    table_stream: &'a [u8],
    index: usize,
    name: &str,
) -> Result<Option<&'a [u8]>> {
    let Some((offset, length)) = fib.get_table_pointer(index) else {
        return Ok(None);
    };
    if length == 0 {
        return Ok(None);
    }
    let start =
        usize::try_from(offset).map_err(|_| corrupted(format!("{name} offset is too large")))?;
    let length =
        usize::try_from(length).map_err(|_| corrupted(format!("{name} length is too large")))?;
    let end = start
        .checked_add(length)
        .ok_or_else(|| corrupted(format!("{name} range overflows")))?;
    table_stream
        .get(start..end)
        .map(Some)
        .ok_or_else(|| corrupted(format!("{name} extends beyond the table stream")))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn caption(label: &str, info: CaptionInfo) -> CaptionDefinition {
        CaptionDefinition::try_new(label.to_string(), info).unwrap()
    }

    fn simple_info() -> CaptionInfo {
        CaptionInfo::new(CaptionLocation::Below, None, false, 0)
    }

    fn chapter_info() -> CaptionInfo {
        CaptionInfo::new(
            CaptionLocation::Above,
            Some(ChapterNumbering::new(
                ChapterHeading::Heading2,
                ChapterSeparator::EnDash,
            )),
            true,
            3,
        )
    }

    fn caption_sttb(definitions: &[CaptionDefinition]) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&u16::MAX.to_le_bytes());
        data.extend_from_slice(&(definitions.len() as u16).to_le_bytes());
        data.extend_from_slice(&(CAPI_SIZE as u16).to_le_bytes());
        for definition in definitions {
            write_string(&mut data, &definition.label);
            data.extend_from_slice(&definition.info.to_bytes());
        }
        data
    }

    fn auto_caption_sttb(entries: &[AutoCaptionEntry]) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&u16::MAX.to_le_bytes());
        data.extend_from_slice(&(entries.len() as u16).to_le_bytes());
        data.extend_from_slice(&AUTO_CAPTION_EXTRA_SIZE.to_le_bytes());
        for entry in entries {
            write_string(&mut data, &entry.prog_id);
            data.extend_from_slice(&entry.caption_index.to_le_bytes());
        }
        data
    }

    #[test]
    fn capi_round_trips_exactly() {
        let info = chapter_info();
        let bytes = info.to_bytes();
        assert_eq!(bytes.len(), CaptionInfo::SIZE);
        // Flags: location Above | fChapNum | iHeading 2 | fNoLabel.
        let flags = u16::from_le_bytes([bytes[0], bytes[1]]);
        assert_eq!(flags, 0x1 | 0x4 | 0x2 << 3 | 0x8000);
        assert_eq!(u16::from_le_bytes([bytes[2], bytes[3]]), 3);
        assert_eq!(u16::from_le_bytes([bytes[4], bytes[5]]), 0x2013);
        assert_eq!(CaptionInfo::from_bytes(&bytes).unwrap(), info);
        assert!(CaptionInfo::from_bytes(&bytes[..5]).is_err());
    }

    #[test]
    fn capi_ignores_chapter_fields_without_fchapnum() {
        let mut bytes = simple_info().to_bytes();
        // Undefined bits plus ignored iHeading/xchSeparator values on read.
        bytes[0] |= 0x78;
        bytes[1] |= 0x7F;
        bytes[4] = 0xFF;
        bytes[5] = 0xFF;
        let info = CaptionInfo::from_bytes(&bytes).unwrap();
        assert_eq!(info, simple_info());
        assert_eq!(info.to_bytes(), simple_info().to_bytes());
    }

    #[test]
    fn capi_rejects_out_of_range_fields() {
        let mut bytes = simple_info().to_bytes();
        // iLocation 0x2 is not a valid insert location.
        bytes[0] = 0x2;
        assert!(CaptionInfo::from_bytes(&bytes).is_err());
        // fChapNum set with iHeading 0 and with iHeading 10.
        bytes = [0x4, 0, 0, 0, 0x2E, 0];
        assert!(CaptionInfo::from_bytes(&bytes).is_err());
        bytes = [0x4 | 0xA << 3, 0, 0, 0, 0x2E, 0];
        assert!(CaptionInfo::from_bytes(&bytes).is_err());
        // fChapNum set with a separator outside the enumerated set.
        bytes = [0x4 | 0x1 << 3, 0, 0, 0, 0x3B, 0];
        assert!(CaptionInfo::from_bytes(&bytes).is_err());
    }

    #[test]
    fn sttbfcaption_parses_and_round_trips() {
        let definitions = [
            caption("Equation", simple_info()),
            caption("Figure", chapter_info()),
            caption(
                "Table",
                CaptionInfo::new(CaptionLocation::Below, None, true, 1),
            ),
        ];
        let bytes = caption_sttb(&definitions);
        let table = CaptionLabelTable::parse_bytes(&bytes).unwrap();
        assert_eq!(table.len(), 3);
        assert_eq!(table.definitions()[1].label(), "Figure");
        assert_eq!(
            table.definitions()[1].info().chapter_numbering(),
            Some(ChapterNumbering::new(
                ChapterHeading::Heading2,
                ChapterSeparator::EnDash
            ))
        );
        assert!(table.definitions()[2].info().omit_label());
        assert_eq!(table.to_bytes().unwrap(), bytes);
    }

    #[test]
    fn sttbfcaption_rejects_malformed_tables() {
        assert!(CaptionLabelTable::parse_bytes(&[]).is_err());
        // Non-extended STTB header.
        let mut bytes = caption_sttb(&[caption("Figure", simple_info())]);
        bytes[0] = 0;
        bytes[1] = 0;
        assert!(CaptionLabelTable::parse_bytes(&bytes).is_err());
        // Wrong cbExtra.
        let mut bytes = caption_sttb(&[caption("Figure", simple_info())]);
        bytes[4] = 2;
        assert!(CaptionLabelTable::parse_bytes(&bytes).is_err());
        // Declared count exceeds the payload.
        let mut bytes = caption_sttb(&[caption("Figure", simple_info())]);
        bytes[2] = 2;
        assert!(CaptionLabelTable::parse_bytes(&bytes).is_err());
        // Truncated CAPI.
        let mut bytes = caption_sttb(&[caption("Figure", simple_info())]);
        bytes.truncate(bytes.len() - 2);
        assert!(CaptionLabelTable::parse_bytes(&bytes).is_err());
        // Trailing bytes after the final entry.
        let mut bytes = caption_sttb(&[caption("Figure", simple_info())]);
        bytes.push(0);
        assert!(CaptionLabelTable::parse_bytes(&bytes).is_err());
        // Label over the 40-character cap.
        let long = "a".repeat(41);
        let mut data = Vec::new();
        data.extend_from_slice(&u16::MAX.to_le_bytes());
        data.extend_from_slice(&1u16.to_le_bytes());
        data.extend_from_slice(&(CAPI_SIZE as u16).to_le_bytes());
        write_string(&mut data, &long);
        data.extend_from_slice(&simple_info().to_bytes());
        assert!(CaptionLabelTable::parse_bytes(&data).is_err());
        assert!(CaptionDefinition::try_new(long, simple_info()).is_err());
        // Unpaired surrogate in a label.
        let mut bytes = caption_sttb(&[caption("Figure", simple_info())]);
        let label_start = 8;
        bytes[label_start] = 0x00;
        bytes[label_start + 1] = 0xD8;
        assert!(CaptionLabelTable::parse_bytes(&bytes).is_err());
    }

    #[test]
    fn sttbfautocaption_parses_and_round_trips() {
        let entries = [
            AutoCaptionEntry::new("Excel.Sheet.8".to_string(), 1),
            AutoCaptionEntry::new("Equation.3".to_string(), 0),
        ];
        let bytes = auto_caption_sttb(&entries);
        let table = AutoCaptionTable::parse_bytes(&bytes).unwrap();
        assert_eq!(table.len(), 2);
        assert_eq!(table.entries()[0].prog_id(), "Excel.Sheet.8");
        assert_eq!(table.entries()[0].caption_index(), 1);
        assert_eq!(table.to_bytes().unwrap(), bytes);
    }

    #[test]
    fn sttbfautocaption_rejects_malformed_tables() {
        assert!(AutoCaptionTable::parse_bytes(&[]).is_err());
        // Wrong cbExtra.
        let mut bytes = auto_caption_sttb(&[AutoCaptionEntry::new("Equation.3".to_string(), 0)]);
        bytes[4] = 6;
        assert!(AutoCaptionTable::parse_bytes(&bytes).is_err());
        // Truncated extra data.
        let mut bytes = auto_caption_sttb(&[AutoCaptionEntry::new("Equation.3".to_string(), 0)]);
        bytes.truncate(bytes.len() - 1);
        assert!(AutoCaptionTable::parse_bytes(&bytes).is_err());
        // Trailing bytes.
        let mut bytes = auto_caption_sttb(&[AutoCaptionEntry::new("Equation.3".to_string(), 0)]);
        bytes.extend_from_slice(&[0, 0]);
        assert!(AutoCaptionTable::parse_bytes(&bytes).is_err());
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
    fn parses_both_tables_through_fib_with_index_bounds() {
        let captions = caption_sttb(&[
            caption("Equation", simple_info()),
            caption("Figure", chapter_info()),
        ]);
        let auto = auto_caption_sttb(&[AutoCaptionEntry::new("Equation.3".to_string(), 1)]);
        let mut table_stream = vec![0u8; 16];
        table_stream.extend_from_slice(&captions);
        table_stream.extend_from_slice(&auto);
        let fib = fib_with_pointers(&[
            (CAPTION_FIB_INDEX, 16, captions.len() as u32),
            (
                AUTO_CAPTION_FIB_INDEX,
                (16 + captions.len()) as u32,
                auto.len() as u32,
            ),
        ]);
        let tables = CaptionTables::parse(&fib, &table_stream).unwrap();
        assert_eq!(tables.captions().unwrap().len(), 2);
        assert_eq!(
            tables.auto_captions().unwrap().entries()[0].caption_index(),
            1
        );
    }

    #[test]
    fn rejects_auto_caption_index_outside_label_table() {
        let captions = caption_sttb(&[caption("Equation", simple_info())]);
        let auto = auto_caption_sttb(&[AutoCaptionEntry::new("Equation.3".to_string(), 1)]);
        let mut table_stream = captions.clone();
        table_stream.extend_from_slice(&auto);
        let fib = fib_with_pointers(&[
            (CAPTION_FIB_INDEX, 0, captions.len() as u32),
            (
                AUTO_CAPTION_FIB_INDEX,
                captions.len() as u32,
                auto.len() as u32,
            ),
        ]);
        assert!(CaptionTables::parse(&fib, &table_stream).is_err());
        // AutoCaption entries without any label table are also rejected.
        let fib = fib_with_pointers(&[(AUTO_CAPTION_FIB_INDEX, 0, auto.len() as u32)]);
        assert!(CaptionTables::parse(&fib, &auto).is_err());
        // An empty AutoCaption table is fine without labels.
        let empty = auto_caption_sttb(&[]);
        let fib = fib_with_pointers(&[(AUTO_CAPTION_FIB_INDEX, 0, empty.len() as u32)]);
        let tables = CaptionTables::parse(&fib, &empty).unwrap();
        assert!(tables.captions().is_none());
        assert!(tables.auto_captions().unwrap().is_empty());
    }

    #[test]
    fn rejects_tables_extending_beyond_table_stream() {
        let captions = caption_sttb(&[caption("Equation", simple_info())]);
        let fib = fib_with_pointers(&[(CAPTION_FIB_INDEX, 0, captions.len() as u32 + 1)]);
        assert!(CaptionTables::parse(&fib, &captions).is_err());
        // Zero-length pointers mean the table is absent.
        let fib = fib_with_pointers(&[(CAPTION_FIB_INDEX, 4, 0)]);
        assert!(
            CaptionTables::parse(&fib, &captions)
                .unwrap()
                .captions()
                .is_none()
        );
    }
}
