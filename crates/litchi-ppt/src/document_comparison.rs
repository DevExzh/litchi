//! Inert PowerPoint 10 document-comparison metadata.
//!
//! These records describe reviewing UI state and slide creation timestamps.
//! Parsing them never compares presentations, opens external data, or executes
//! embedded content.

use super::package::{Error, Result};
use super::records::Record;
use crate::consts::RecordType;

const RECORD_HEADER_SIZE: usize = 8;
const SIZE_ATOM_PAYLOAD_SIZE: usize = 4;
const ENTRY_PAYLOAD_SIZE: usize = 12;
const SIZE_ATOM_RECORD_SIZE: usize = RECORD_HEADER_SIZE + SIZE_ATOM_PAYLOAD_SIZE;
const ENTRY_RECORD_SIZE: usize = RECORD_HEADER_SIZE + ENTRY_PAYLOAD_SIZE;
const MAX_SLIDE_LIST_ENTRIES: usize = 1_000_000;

/// Reviewing toolbar and gallery display state from `DocToolbarStates10Atom`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ReviewingToolbarStates {
    show_reviewing_toolbar: bool,
    show_reviewing_gallery: bool,
    ignored_reserved_bits: u8,
}

/// Maximum supported nesting for PowerPoint 10 document-comparison records.
pub const POWERPOINT_DIFF_MAX_DEPTH: usize = 32;
/// Maximum number of diff records accepted in one comparison tree.
pub const POWERPOINT_DIFF_MAX_RECORDS: usize = 65_536;

const DIFF_HEADER_SIZE: usize = 28;
const DIFF_FIXED_SIZE: usize = 32;
const REVIEWER_NAME_RECORD_TYPE: u16 = 0x0FBA;
const MAX_REVIEWER_NAME_BYTES: usize = 104;

/// A value of `DiffTypeEnum` from MS-PPT section 2.13.8.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u32)]
pub enum DiffType {
    Document = 0x00,
    Slide = 0x02,
    MainMaster = 0x03,
    SlideList = 0x04,
    MasterList = 0x05,
    ShapeList = 0x06,
    Shape = 0x07,
    Text = 0x09,
    Notes = 0x0A,
    SlideShow = 0x0B,
    HeaderFooter = 0x0C,
    NamedShow = 0x0E,
    NamedShowList = 0x0F,
    RecolorInfo = 0x12,
    ExternalObject = 0x13,
    TableList = 0x15,
    Table = 0x16,
    InteractiveInfo = 0x17,
}

impl DiffType {
    pub const fn as_u32(self) -> u32 {
        self as u32
    }
}

impl TryFrom<u32> for DiffType {
    type Error = Error;

    fn try_from(value: u32) -> Result<Self> {
        Ok(match value {
            0x00 => Self::Document,
            0x02 => Self::Slide,
            0x03 => Self::MainMaster,
            0x04 => Self::SlideList,
            0x05 => Self::MasterList,
            0x06 => Self::ShapeList,
            0x07 => Self::Shape,
            0x09 => Self::Text,
            0x0A => Self::Notes,
            0x0B => Self::SlideShow,
            0x0C => Self::HeaderFooter,
            0x0E => Self::NamedShow,
            0x0F => Self::NamedShowList,
            0x12 => Self::RecolorInfo,
            0x13 => Self::ExternalObject,
            0x15 => Self::TableList,
            0x16 => Self::Table,
            0x17 => Self::InteractiveInfo,
            _ => {
                return Err(Error::Corrupted(format!(
                    "invalid DiffTypeEnum value {value:#010X}"
                )));
            },
        })
    }
}

/// A value of `ElementTypeEnum` from MS-PPT section 2.13.9.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u32)]
pub enum ElementType {
    Shape = 0x01,
    Sound = 0x02,
}

impl ElementType {
    pub const fn as_u32(self) -> u32 {
        self as u32
    }
}

impl TryFrom<u32> for ElementType {
    type Error = Error;

    fn try_from(value: u32) -> Result<Self> {
        match value {
            0x01 => Ok(Self::Shape),
            0x02 => Ok(Self::Sound),
            _ => Err(Error::Corrupted(format!(
                "invalid ElementTypeEnum value {value:#010X}"
            ))),
        }
    }
}

macro_rules! define_diff_flags {
    ($name:ident { $($field:ident: $bit:expr),+ $(,)? }) => {
        #[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
        pub struct $name {
            $(pub $field: bool,)+
        }

        impl $name {
            fn from_raw(raw: u32) -> Self {
                Self { $($field: raw & (1 << $bit) != 0,)+ }
            }

            fn to_raw(self) -> u32 {
                0 $(| (u32::from(self.$field) << $bit))+
            }

            const fn mask() -> u32 {
                0 $(| (1 << $bit))+
            }
        }
    };
}

define_diff_flags!(DocDiffFlags {
    slide_size: 2,
    omit_title_placeholder: 3,
    named_show_list: 4,
    slide_header_footer: 5,
    notes_header_footer: 6,
});

define_diff_flags!(SlideDiffFlags {
    scheme: 0,
    background: 1,
    add_slide: 4,
    delete_slide: 5,
    layout: 6,
    slide_show: 7,
    header_footer: 8,
    master: 10,
    position: 11,
    time_node: 12,
});

define_diff_flags!(MainMasterDiffFlags {
    scheme: 0,
    background: 1,
    time_node: 12,
    add_main_master: 13,
    delete_main_master: 14,
    locked: 15,
});

define_diff_flags!(ShapeDiffFlags {
    add_shape: 0,
    delete_shape: 1,
    child: 2,
    position: 3,
    recolor_info: 4,
    external_object: 5,
    interactive_info_on_over: 6,
    interactive_info_on_click: 7,
    settings_3d: 9,
    black_and_white_settings: 10,
    auto_shape: 11,
    line_style: 12,
    fill_style: 13,
    shadow_style: 14,
    word_art: 15,
    picture: 16,
    orientation: 17,
    text_settings: 18,
    size: 20,
    ruler: 22,
});

define_diff_flags!(TableDiffFlags {
    add_table: 0,
    delete_table: 1,
    modified_table: 2,
    position: 3,
});

define_diff_flags!(TextDiffFlags { word_list: 2 });

/// Typed interpretation of the 32-bit payload following `DiffRecordHeaders`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DiffFlags {
    Document(DocDiffFlags),
    Slide(SlideDiffFlags),
    MainMaster(MainMasterDiffFlags),
    Shape(ShapeDiffFlags),
    Table(TableDiffFlags),
    Text(TextDiffFlags),
    Notes(TextDiffFlags),
    None,
}

impl DiffFlags {
    pub fn for_type(diff_type: DiffType) -> Self {
        Self::from_raw(diff_type, 0).0
    }

    fn from_raw(diff_type: DiffType, raw: u32) -> (Self, u32) {
        let (flags, mask) = match diff_type {
            DiffType::Document => (
                Self::Document(DocDiffFlags::from_raw(raw)),
                DocDiffFlags::mask(),
            ),
            DiffType::Slide => (
                Self::Slide(SlideDiffFlags::from_raw(raw)),
                SlideDiffFlags::mask(),
            ),
            DiffType::MainMaster => (
                Self::MainMaster(MainMasterDiffFlags::from_raw(raw)),
                MainMasterDiffFlags::mask(),
            ),
            DiffType::Shape => (
                Self::Shape(ShapeDiffFlags::from_raw(raw)),
                ShapeDiffFlags::mask(),
            ),
            DiffType::Table => (
                Self::Table(TableDiffFlags::from_raw(raw)),
                TableDiffFlags::mask(),
            ),
            DiffType::Text => (
                Self::Text(TextDiffFlags::from_raw(raw)),
                TextDiffFlags::mask(),
            ),
            DiffType::Notes => (
                Self::Notes(TextDiffFlags::from_raw(raw)),
                TextDiffFlags::mask(),
            ),
            _ => (Self::None, 0),
        };
        (flags, raw & !mask)
    }

    fn diff_type(self) -> Option<DiffType> {
        match self {
            Self::Document(_) => Some(DiffType::Document),
            Self::Slide(_) => Some(DiffType::Slide),
            Self::MainMaster(_) => Some(DiffType::MainMaster),
            Self::Shape(_) => Some(DiffType::Shape),
            Self::Table(_) => Some(DiffType::Table),
            Self::Text(_) => Some(DiffType::Text),
            Self::Notes(_) => Some(DiffType::Notes),
            Self::None => None,
        }
    }

    fn to_raw(self) -> u32 {
        match self {
            Self::Document(value) => value.to_raw(),
            Self::Slide(value) => value.to_raw(),
            Self::MainMaster(value) => value.to_raw(),
            Self::Shape(value) => value.to_raw(),
            Self::Table(value) => value.to_raw(),
            Self::Text(value) | Self::Notes(value) => value.to_raw(),
            Self::None => 0,
        }
    }
}

/// The atom portion of a `DiffRecordHeaders` structure.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DiffRecordHeaders {
    index: bool,
    diff_type: DiffType,
    ignored_prefix: [u8; 3],
    ignored_tail: u32,
}

impl DiffRecordHeaders {
    pub const fn new(index: bool, diff_type: DiffType) -> Self {
        Self {
            index,
            diff_type,
            ignored_prefix: [0; 3],
            ignored_tail: 0,
        }
    }

    pub const fn index(&self) -> bool {
        self.index
    }

    pub const fn diff_type(&self) -> DiffType {
        self.diff_type
    }

    pub const fn ignored_prefix(&self) -> [u8; 3] {
        self.ignored_prefix
    }

    pub const fn ignored_tail(&self) -> u32 {
        self.ignored_tail
    }
}

/// One recursively parsed `RT_Diff10` container.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DiffNode {
    headers: DiffRecordHeaders,
    flags: DiffFlags,
    ignored_flag_bits: u32,
    children: Vec<Self>,
}

impl DiffNode {
    pub fn new(
        diff_type: DiffType,
        index: bool,
        flags: DiffFlags,
        children: Vec<Self>,
    ) -> Result<Self> {
        let node = Self {
            headers: DiffRecordHeaders::new(index, diff_type),
            flags,
            ignored_flag_bits: 0,
            children,
        };
        node.validate_node()?;
        Ok(node)
    }

    pub const fn headers(&self) -> &DiffRecordHeaders {
        &self.headers
    }

    pub const fn diff_type(&self) -> DiffType {
        self.headers.diff_type
    }

    pub const fn flags(&self) -> DiffFlags {
        self.flags
    }

    pub const fn ignored_flag_bits(&self) -> u32 {
        self.ignored_flag_bits
    }

    pub fn children(&self) -> &[Self] {
        &self.children
    }

    pub fn children_of_type(&self, diff_type: DiffType) -> impl Iterator<Item = &Self> {
        self.children
            .iter()
            .filter(move |child| child.diff_type() == diff_type)
    }

    pub fn to_record_bytes(&self) -> Result<Vec<u8>> {
        let mut count = 0;
        self.encode(0, &mut count)
    }

    fn parse(
        data: &[u8],
        depth: usize,
        limits: DiffLimits,
        count: &mut usize,
    ) -> Result<(Self, usize)> {
        if depth > limits.max_depth {
            return corrupted("document-comparison tree exceeds the nesting limit");
        }
        *count = count.checked_add(1).ok_or_else(|| {
            Error::Corrupted("document-comparison record count overflow".to_string())
        })?;
        if *count > limits.max_records {
            return corrupted("document-comparison tree exceeds the record-count limit");
        }
        if data.len() < DIFF_FIXED_SIZE {
            return corrupted("truncated Diff10 container");
        }
        let (version, instance) = unpack_version_instance(read_u16(data, 0)?);
        if version != 0xF || instance != 0 || read_u16(data, 2)? != RecordType::Diff10.as_u16() {
            return corrupted("invalid Diff10 record header");
        }
        let payload_len = usize::try_from(read_u32(data, 4)?)
            .map_err(|_| Error::Corrupted("Diff10 length does not fit usize".to_string()))?;
        let total_len = RECORD_HEADER_SIZE
            .checked_add(payload_len)
            .ok_or_else(|| Error::Corrupted("Diff10 length overflow".to_string()))?;
        if total_len > data.len() || payload_len < DIFF_FIXED_SIZE - RECORD_HEADER_SIZE {
            return corrupted("Diff10 record extends beyond its parent");
        }
        let (atom_version, atom_instance) = unpack_version_instance(read_u16(data, 8)?);
        if atom_version != 0
            || atom_instance != 0
            || read_u16(data, 10)? != RecordType::Diff10Atom.as_u16()
            || read_u32(data, 12)? != 12
        {
            return corrupted("invalid Diff10Atom header");
        }
        let index = match data[16] {
            0 => false,
            1 => true,
            _ => return corrupted("Diff10Atom fIndex is not a bool1"),
        };
        let diff_type = DiffType::try_from(read_u32(data, 20)?)?;
        if index
            && !matches!(
                diff_type,
                DiffType::HeaderFooter | DiffType::InteractiveInfo
            )
        {
            return corrupted("Diff10 fIndex is invalid for its diff type");
        }
        let headers = DiffRecordHeaders {
            index,
            diff_type,
            ignored_prefix: [data[17], data[18], data[19]],
            ignored_tail: read_u32(data, 24)?,
        };
        let raw_flags = read_u32(data, DIFF_HEADER_SIZE)?;
        let (flags, ignored_flag_bits) = DiffFlags::from_raw(diff_type, raw_flags);
        let mut children = Vec::new();
        let mut offset = DIFF_FIXED_SIZE;
        while offset < total_len {
            let (child, consumed) =
                Self::parse(&data[offset..total_len], depth + 1, limits, count)?;
            if consumed == 0 {
                return corrupted("zero-length Diff10 child");
            }
            children.push(child);
            offset += consumed;
        }
        let node = Self {
            headers,
            flags,
            ignored_flag_bits,
            children,
        };
        node.validate_node()?;
        Ok((node, total_len))
    }

    fn encode(&self, depth: usize, count: &mut usize) -> Result<Vec<u8>> {
        if depth > POWERPOINT_DIFF_MAX_DEPTH {
            return corrupted("document-comparison tree exceeds the nesting limit");
        }
        *count = count.checked_add(1).ok_or_else(|| {
            Error::Corrupted("document-comparison record count overflow".to_string())
        })?;
        if *count > POWERPOINT_DIFF_MAX_RECORDS {
            return corrupted("document-comparison tree exceeds the record-count limit");
        }
        self.validate_node()?;
        let mut atom = Vec::with_capacity(12);
        atom.push(u8::from(self.headers.index));
        atom.extend_from_slice(&[0; 3]);
        atom.extend_from_slice(&self.headers.diff_type.as_u32().to_le_bytes());
        atom.extend_from_slice(&0u32.to_le_bytes());
        let mut payload = encode_record(0, 0, RecordType::Diff10Atom, &atom);
        payload.extend_from_slice(&self.flags.to_raw().to_le_bytes());
        for child in &self.children {
            payload.extend_from_slice(&child.encode(depth + 1, count)?);
        }
        Ok(encode_record(0xF, 0, RecordType::Diff10, &payload))
    }

    fn validate_node(&self) -> Result<()> {
        if self.headers.index
            && !matches!(
                self.headers.diff_type,
                DiffType::HeaderFooter | DiffType::InteractiveInfo
            )
        {
            return corrupted("Diff10 fIndex is invalid for its diff type");
        }
        if let Some(flag_type) = self.flags.diff_type() {
            if flag_type != self.headers.diff_type {
                return corrupted("Diff10 flags do not match the record tag");
            }
        } else if matches!(
            self.headers.diff_type,
            DiffType::Document
                | DiffType::Slide
                | DiffType::MainMaster
                | DiffType::Shape
                | DiffType::Table
                | DiffType::Text
                | DiffType::Notes
        ) {
            return corrupted("Diff10 record is missing its typed flags");
        }
        validate_diff_children(self.headers.diff_type, &self.children)
    }
}

#[derive(Debug, Clone, Copy)]
struct DiffLimits {
    max_depth: usize,
    max_records: usize,
}

/// A complete `DiffTree10Container`, without dereferencing its reviewer document.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DiffTree10 {
    reviewer_name: String,
    document_diff: DiffNode,
}

impl DiffTree10 {
    pub fn new(reviewer_name: String, document_diff: DiffNode) -> Result<Self> {
        validate_reviewer_name(&reviewer_name)?;
        if document_diff.diff_type() != DiffType::Document {
            return corrupted("DiffTree10 root is not a DocDiff10 container");
        }
        Ok(Self {
            reviewer_name,
            document_diff,
        })
    }

    pub fn parse(record: &Record) -> Result<Self> {
        Self::parse_with_limits(
            record,
            POWERPOINT_DIFF_MAX_DEPTH,
            POWERPOINT_DIFF_MAX_RECORDS,
        )
    }

    pub fn parse_with_limits(
        record: &Record,
        max_depth: usize,
        max_records: usize,
    ) -> Result<Self> {
        if max_records == 0 {
            return corrupted("document-comparison record limit must be nonzero");
        }
        if record.record_type_raw != RecordType::DiffTree10.as_u16()
            || record.version != 0xF
            || record.instance != 0
            || usize::try_from(record.data_length).ok() != Some(record.data.len())
        {
            return corrupted("invalid DiffTree10 container header");
        }
        let limits = DiffLimits {
            max_depth: max_depth.min(POWERPOINT_DIFF_MAX_DEPTH),
            max_records: max_records.min(POWERPOINT_DIFF_MAX_RECORDS),
        };
        let (reviewer_name, reviewer_len) = parse_reviewer_name(&record.data)?;
        let mut count = 0;
        let (document_diff, diff_len) =
            DiffNode::parse(&record.data[reviewer_len..], 0, limits, &mut count)?;
        if document_diff.diff_type() != DiffType::Document {
            return corrupted("DiffTree10 root is not a DocDiff10 container");
        }
        if reviewer_len + diff_len != record.data.len() {
            return corrupted("DiffTree10 has trailing records");
        }
        Ok(Self {
            reviewer_name,
            document_diff,
        })
    }

    pub fn reviewer_name(&self) -> &str {
        &self.reviewer_name
    }

    pub const fn document_diff(&self) -> &DiffNode {
        &self.document_diff
    }

    pub fn to_record_bytes(&self) -> Result<Vec<u8>> {
        validate_reviewer_name(&self.reviewer_name)?;
        if self.document_diff.diff_type() != DiffType::Document {
            return corrupted("DiffTree10 root is not a DocDiff10 container");
        }
        let mut reviewer_payload = Vec::new();
        for unit in self.reviewer_name.encode_utf16() {
            reviewer_payload.extend_from_slice(&unit.to_le_bytes());
        }
        let mut payload = encode_record_raw(0, 0, REVIEWER_NAME_RECORD_TYPE, &reviewer_payload);
        payload.extend_from_slice(&self.document_diff.to_record_bytes()?);
        Ok(encode_record(0xF, 0, RecordType::DiffTree10, &payload))
    }
}

fn validate_diff_children(parent: DiffType, children: &[DiffNode]) -> Result<()> {
    use DiffType as T;
    match parent {
        T::NamedShowList => require_repeated(children, &[T::NamedShow]),
        T::MasterList => require_repeated(children, &[T::MainMaster, T::Slide]),
        T::SlideList => require_repeated(children, &[T::Slide]),
        T::ShapeList => require_repeated(children, &[T::Shape]),
        T::TableList => require_repeated(children, &[T::Table]),
        T::Document => require_ordered(
            children,
            &[
                (T::HeaderFooter, Some(true)),
                (T::HeaderFooter, Some(false)),
                (T::NamedShowList, None),
                (T::MasterList, None),
                (T::SlideList, None),
            ],
        ),
        T::MainMaster => require_ordered(
            children,
            &[(T::ShapeList, None), (T::TableList, None), (T::Notes, None)],
        ),
        T::Slide => require_ordered(
            children,
            &[
                (T::ShapeList, None),
                (T::TableList, None),
                (T::SlideShow, None),
                (T::HeaderFooter, Some(true)),
                (T::Notes, None),
            ],
        ),
        T::Shape => require_ordered(
            children,
            &[
                (T::Text, None),
                (T::RecolorInfo, None),
                (T::ExternalObject, None),
                (T::InteractiveInfo, Some(true)),
                (T::InteractiveInfo, Some(false)),
            ],
        ),
        _ if children.is_empty() => Ok(()),
        _ => corrupted("leaf Diff10 record contains child records"),
    }
}

fn require_repeated(children: &[DiffNode], allowed: &[DiffType]) -> Result<()> {
    if children
        .iter()
        .all(|child| allowed.contains(&child.diff_type()))
    {
        Ok(())
    } else {
        corrupted("Diff10 list contains a child of the wrong type")
    }
}

fn require_ordered(children: &[DiffNode], grammar: &[(DiffType, Option<bool>)]) -> Result<()> {
    let mut previous = None;
    for child in children {
        let rank = grammar
            .iter()
            .position(|(diff_type, index)| {
                child.diff_type() == *diff_type
                    && index.is_none_or(|value| child.headers.index == value)
            })
            .ok_or_else(|| {
                Error::Corrupted("Diff10 container contains a child of the wrong type".to_string())
            })?;
        if previous.is_some_and(|value| rank <= value) {
            return corrupted("Diff10 children are duplicated or out of order");
        }
        previous = Some(rank);
    }
    Ok(())
}

fn parse_reviewer_name(data: &[u8]) -> Result<(String, usize)> {
    if data.len() < RECORD_HEADER_SIZE {
        return corrupted("DiffTree10 is missing ReviewerNameAtom");
    }
    let (version, instance) = unpack_version_instance(read_u16(data, 0)?);
    let byte_len = usize::try_from(read_u32(data, 4)?)
        .map_err(|_| Error::Corrupted("reviewer-name length does not fit usize".to_string()))?;
    let total_len = RECORD_HEADER_SIZE
        .checked_add(byte_len)
        .ok_or_else(|| Error::Corrupted("reviewer-name length overflow".to_string()))?;
    if version != 0
        || instance != 0
        || read_u16(data, 2)? != REVIEWER_NAME_RECORD_TYPE
        || byte_len > MAX_REVIEWER_NAME_BYTES
        || byte_len % 2 != 0
        || total_len > data.len()
    {
        return corrupted("invalid ReviewerNameAtom");
    }
    let mut units = Vec::with_capacity(byte_len / 2);
    for chunk in data[RECORD_HEADER_SIZE..total_len].chunks_exact(2) {
        let unit = u16::from_le_bytes([chunk[0], chunk[1]]);
        if unit == 0 {
            break;
        }
        units.push(unit);
    }
    let name = String::from_utf16(&units)
        .map_err(|_| Error::Corrupted("ReviewerNameAtom contains invalid UTF-16".to_string()))?;
    validate_reviewer_name(&name)?;
    Ok((name, total_len))
}

fn validate_reviewer_name(name: &str) -> Result<()> {
    if name.encode_utf16().count() * 2 > MAX_REVIEWER_NAME_BYTES {
        return corrupted("reviewer name exceeds 104 bytes");
    }
    if name.chars().any(|character| {
        let value = character as u32;
        value == 0 || value <= 0x1F || (0x7F..=0x9F).contains(&value)
    }) {
        return corrupted("reviewer name is not a PrintableUnicodeString");
    }
    Ok(())
}

fn unpack_version_instance(value: u16) -> (u16, u16) {
    (value & 0x000F, value >> 4)
}

fn read_u16(data: &[u8], offset: usize) -> Result<u16> {
    let bytes = data
        .get(offset..offset + 2)
        .ok_or_else(|| Error::Corrupted("truncated document-comparison field".to_string()))?;
    Ok(u16::from_le_bytes([bytes[0], bytes[1]]))
}

fn read_u32(data: &[u8], offset: usize) -> Result<u32> {
    let bytes = data
        .get(offset..offset + 4)
        .ok_or_else(|| Error::Corrupted("truncated document-comparison field".to_string()))?;
    Ok(u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]))
}

fn encode_record_raw(version: u16, instance: u16, record_type: u16, payload: &[u8]) -> Vec<u8> {
    let mut bytes = Vec::with_capacity(RECORD_HEADER_SIZE + payload.len());
    bytes.extend_from_slice(&((instance << 4) | (version & 0xF)).to_le_bytes());
    bytes.extend_from_slice(&record_type.to_le_bytes());
    bytes.extend_from_slice(&(payload.len() as u32).to_le_bytes());
    bytes.extend_from_slice(payload);
    bytes
}

#[cfg(test)]
mod diff_tree_tests {
    use super::*;

    fn node(diff_type: DiffType, index: bool, children: Vec<DiffNode>) -> DiffNode {
        DiffNode::new(diff_type, index, DiffFlags::for_type(diff_type), children).unwrap()
    }

    fn record_from_tree(tree: &DiffTree10) -> Record {
        let bytes = tree.to_record_bytes().unwrap();
        Record::parse(&bytes, 0).unwrap().0
    }

    #[test]
    fn diff_enums_are_exhaustive_and_reject_gaps() {
        let values = [
            0, 2, 3, 4, 5, 6, 7, 9, 10, 11, 12, 14, 15, 18, 19, 21, 22, 23,
        ];
        assert!(
            values
                .into_iter()
                .all(|value| DiffType::try_from(value).is_ok())
        );
        assert!(DiffType::try_from(1).is_err());
        assert_eq!(ElementType::try_from(1).unwrap(), ElementType::Shape);
        assert!(ElementType::try_from(0).is_err());
    }

    #[test]
    fn diff_tree_round_trips_and_canonicalizes_ignored_bits() {
        let named_show = node(DiffType::NamedShow, false, vec![]);
        let named_show_list = node(DiffType::NamedShowList, false, vec![named_show]);
        let document = DiffNode::new(
            DiffType::Document,
            false,
            DiffFlags::Document(DocDiffFlags {
                slide_size: true,
                ..Default::default()
            }),
            vec![named_show_list],
        )
        .unwrap();
        let tree = DiffTree10::new("Reviewer".to_string(), document).unwrap();
        let mut bytes = tree.to_record_bytes().unwrap();
        let reviewer_len = 8 + "Reviewer".encode_utf16().count() * 2;
        let doc_flags_offset = 8 + reviewer_len + DIFF_HEADER_SIZE;
        bytes[doc_flags_offset + 3] = 0x80;
        let record = Record::parse(&bytes, 0).unwrap().0;
        let parsed = DiffTree10::parse(&record).unwrap();
        assert_eq!(parsed.reviewer_name(), "Reviewer");
        assert_eq!(parsed.document_diff().ignored_flag_bits(), 0x8000_0000);
        let canonical = parsed.to_record_bytes().unwrap();
        assert_eq!(canonical[doc_flags_offset + 3] & 0x80, 0);
    }

    #[test]
    fn malformed_tag_and_child_order_are_rejected() {
        let document = node(DiffType::Document, false, vec![]);
        let tree = DiffTree10::new("R".to_string(), document).unwrap();
        let mut bytes = tree.to_record_bytes().unwrap();
        let tag_offset = 8 + 8 + 2 + 20;
        bytes[tag_offset..tag_offset + 4].copy_from_slice(&1u32.to_le_bytes());
        let record = Record::parse(&bytes, 0).unwrap().0;
        assert!(DiffTree10::parse(&record).is_err());

        let wrong_child = node(DiffType::Text, false, vec![]);
        assert!(
            DiffNode::new(
                DiffType::Document,
                false,
                DiffFlags::for_type(DiffType::Document),
                vec![wrong_child],
            )
            .is_err()
        );
    }

    #[test]
    fn depth_and_record_count_limits_are_enforced() {
        let text = node(DiffType::Text, false, vec![]);
        let shape = node(DiffType::Shape, false, vec![text]);
        let shape_list = node(DiffType::ShapeList, false, vec![shape]);
        let slide = node(DiffType::Slide, false, vec![shape_list]);
        let slide_list = node(DiffType::SlideList, false, vec![slide]);
        let document = node(DiffType::Document, false, vec![slide_list]);
        let tree = DiffTree10::new("R".to_string(), document).unwrap();
        let record = record_from_tree(&tree);
        assert!(DiffTree10::parse_with_limits(&record, 2, 100).is_err());
        assert!(DiffTree10::parse_with_limits(&record, 32, 3).is_err());
    }
}

impl ReviewingToolbarStates {
    /// Construct reviewing UI state. Reserved bits are emitted as zero.
    pub const fn new(show_reviewing_toolbar: bool, show_reviewing_gallery: bool) -> Self {
        Self {
            show_reviewing_toolbar,
            show_reviewing_gallery,
            ignored_reserved_bits: 0,
        }
    }

    /// Parse a strict `DocToolbarStates10Atom` record.
    pub fn parse(record: &Record) -> Result<Self> {
        validate_atom(record, RecordType::DocToolbarStates10Atom, 1)?;
        let value = record.data[0];
        Ok(Self {
            show_reviewing_toolbar: value & 0x01 != 0,
            show_reviewing_gallery: value & 0x02 != 0,
            // MS-PPT requires writers to zero these bits and readers to ignore
            // them. Retaining them keeps a parsed record byte-stable.
            ignored_reserved_bits: value & 0xfc,
        })
    }

    pub const fn show_reviewing_toolbar(self) -> bool {
        self.show_reviewing_toolbar
    }

    pub const fn show_reviewing_gallery(self) -> bool {
        self.show_reviewing_gallery
    }

    /// Raw ignored bits retained from an existing record.
    pub const fn ignored_reserved_bits(self) -> u8 {
        self.ignored_reserved_bits
    }

    pub fn set_show_reviewing_toolbar(&mut self, value: bool) {
        self.show_reviewing_toolbar = value;
    }

    pub fn set_show_reviewing_gallery(&mut self, value: bool) {
        self.show_reviewing_gallery = value;
    }

    /// Serialize the exact atom, preserving ignored bits from parsed input.
    pub fn to_record_bytes(self) -> Vec<u8> {
        let value = self.ignored_reserved_bits
            | u8::from(self.show_reviewing_toolbar)
            | (u8::from(self.show_reviewing_gallery) << 1);
        encode_record(0, 0, RecordType::DocToolbarStates10Atom, &[value])
    }
}

/// One slide creation-time entry from a document-comparison slide list.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SlideCreationEntry {
    slide_id_ref: u32,
    file_time: u64,
}

impl SlideCreationEntry {
    pub const fn new(slide_id_ref: u32, file_time: u64) -> Self {
        Self {
            slide_id_ref,
            file_time,
        }
    }

    pub const fn slide_id_ref(self) -> u32 {
        self.slide_id_ref
    }

    /// Raw Windows `FILETIME` ticks retained without clock conversion.
    pub const fn file_time(self) -> u64 {
        self.file_time
    }

    fn parse_payload(payload: &[u8]) -> Self {
        let slide_id_ref = u32::from_le_bytes(payload[0..4].try_into().expect("fixed payload"));
        let high = u32::from_le_bytes(payload[4..8].try_into().expect("fixed payload"));
        let low = u32::from_le_bytes(payload[8..12].try_into().expect("fixed payload"));
        Self {
            slide_id_ref,
            file_time: (u64::from(high) << 32) | u64::from(low),
        }
    }

    fn payload(self) -> [u8; ENTRY_PAYLOAD_SIZE] {
        let mut payload = [0; ENTRY_PAYLOAD_SIZE];
        payload[..4].copy_from_slice(&self.slide_id_ref.to_le_bytes());
        payload[4..8].copy_from_slice(&(self.file_time >> 32).to_le_bytes()[..4]);
        payload[8..].copy_from_slice(&(self.file_time as u32).to_le_bytes());
        payload
    }

    /// Serialize a strict `SlideListEntry10Atom` record.
    pub fn to_record_bytes(self) -> Vec<u8> {
        encode_record(0, 0, RecordType::SlideListEntry10Atom, &self.payload())
    }
}

/// Strict `SlideListTable10Container` metadata.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct SlideListTable10 {
    entries: Vec<SlideCreationEntry>,
}

impl SlideListTable10 {
    pub fn new(entries: Vec<SlideCreationEntry>) -> Result<Self> {
        validate_count(entries.len())?;
        Ok(Self { entries })
    }

    pub fn entries(&self) -> &[SlideCreationEntry] {
        &self.entries
    }

    pub fn push(&mut self, entry: SlideCreationEntry) -> Result<()> {
        validate_count(self.entries.len().saturating_add(1))?;
        self.entries.push(entry);
        Ok(())
    }

    /// Parse a strict container without allocating intermediate child records.
    pub fn parse(record: &Record) -> Result<Self> {
        if record.record_type != RecordType::SlideListTable10
            || record.version != 0x0f
            || record.instance != 0
            || usize::try_from(record.data_length).ok() != Some(record.data.len())
        {
            return corrupted("SlideListTable10Container has an invalid record header");
        }
        let data = record.data.as_slice();
        if data.len() < SIZE_ATOM_RECORD_SIZE {
            return corrupted("SlideListTable10Container is truncated");
        }
        let size_payload = child_payload(
            data,
            0,
            RecordType::SlideListTableSize10Atom,
            SIZE_ATOM_PAYLOAD_SIZE,
        )?;
        let signed_count = i32::from_le_bytes(size_payload.try_into().expect("fixed payload"));
        let count = usize::try_from(signed_count)
            .map_err(|_| Error::Corrupted("negative slide-list table count".to_string()))?;
        validate_count(count)?;
        let expected = count
            .checked_mul(ENTRY_RECORD_SIZE)
            .and_then(|size| size.checked_add(SIZE_ATOM_RECORD_SIZE))
            .ok_or_else(|| Error::Corrupted("slide-list table size overflow".to_string()))?;
        if data.len() != expected {
            return corrupted("SlideListTable10Container count does not match its payload");
        }

        let mut entries = Vec::with_capacity(count);
        let mut offset = SIZE_ATOM_RECORD_SIZE;
        for _ in 0..count {
            let payload = child_payload(
                data,
                offset,
                RecordType::SlideListEntry10Atom,
                ENTRY_PAYLOAD_SIZE,
            )?;
            entries.push(SlideCreationEntry::parse_payload(payload));
            offset += ENTRY_RECORD_SIZE;
        }
        Ok(Self { entries })
    }

    /// Serialize the size atom followed by the exact declared entry array.
    pub fn to_record_bytes(&self) -> Result<Vec<u8>> {
        validate_count(self.entries.len())?;
        let count = i32::try_from(self.entries.len())
            .map_err(|_| Error::Corrupted("slide-list table count overflow".to_string()))?;
        let capacity = SIZE_ATOM_RECORD_SIZE
            .checked_add(self.entries.len().saturating_mul(ENTRY_RECORD_SIZE))
            .ok_or_else(|| Error::Corrupted("slide-list table size overflow".to_string()))?;
        let mut payload = Vec::with_capacity(capacity);
        payload.extend_from_slice(&encode_record(
            0,
            0,
            RecordType::SlideListTableSize10Atom,
            &count.to_le_bytes(),
        ));
        for entry in &self.entries {
            payload.extend_from_slice(&entry.to_record_bytes());
        }
        Ok(encode_record(
            0x0f,
            0,
            RecordType::SlideListTable10,
            &payload,
        ))
    }
}

fn validate_atom(record: &Record, kind: RecordType, payload_size: usize) -> Result<()> {
    if record.record_type != kind
        || record.version != 0
        || record.instance != 0
        || record.data.len() != payload_size
        || usize::try_from(record.data_length).ok() != Some(payload_size)
        || !record.children.is_empty()
    {
        return corrupted("PowerPoint document-comparison atom has an invalid header or length");
    }
    Ok(())
}

fn validate_count(count: usize) -> Result<()> {
    if count > MAX_SLIDE_LIST_ENTRIES {
        return corrupted("slide-list table count exceeds the MS-PPT limit");
    }
    Ok(())
}

fn child_payload(
    data: &[u8],
    offset: usize,
    kind: RecordType,
    payload_size: usize,
) -> Result<&[u8]> {
    let header_end = offset
        .checked_add(RECORD_HEADER_SIZE)
        .ok_or_else(|| Error::Corrupted("record offset overflow".to_string()))?;
    let header = data
        .get(offset..header_end)
        .ok_or_else(|| Error::Corrupted("truncated document-comparison child".to_string()))?;
    let version_instance = u16::from_le_bytes([header[0], header[1]]);
    let raw_type = u16::from_le_bytes([header[2], header[3]]);
    let length = u32::from_le_bytes(header[4..8].try_into().expect("fixed header"));
    if version_instance != 0 || raw_type != kind.as_u16() || length as usize != payload_size {
        return corrupted("document-comparison child has an invalid record header");
    }
    let end = header_end
        .checked_add(payload_size)
        .ok_or_else(|| Error::Corrupted("record length overflow".to_string()))?;
    data.get(header_end..end)
        .ok_or_else(|| Error::Corrupted("truncated document-comparison payload".to_string()))
}

fn encode_record(version: u16, instance: u16, kind: RecordType, payload: &[u8]) -> Vec<u8> {
    let mut output = Vec::with_capacity(RECORD_HEADER_SIZE + payload.len());
    output.extend_from_slice(&((instance << 4) | version).to_le_bytes());
    output.extend_from_slice(&kind.as_u16().to_le_bytes());
    output.extend_from_slice(&(payload.len() as u32).to_le_bytes());
    output.extend_from_slice(payload);
    output
}

fn corrupted<T>(message: &str) -> Result<T> {
    Err(Error::Corrupted(message.to_string()))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn record(version: u16, kind: RecordType, data: Vec<u8>) -> Record {
        Record {
            version,
            instance: 0,
            record_type: kind,
            record_type_raw: kind.as_u16(),
            data_length: data.len() as u32,
            data,
            children: Vec::new(),
        }
    }

    #[test]
    fn reviewing_toolbar_round_trips_ignored_bits_and_mutation() {
        let parsed = ReviewingToolbarStates::parse(&record(
            0,
            RecordType::DocToolbarStates10Atom,
            vec![0xfd],
        ))
        .unwrap();
        assert!(parsed.show_reviewing_toolbar());
        assert!(!parsed.show_reviewing_gallery());
        assert_eq!(parsed.ignored_reserved_bits(), 0xfc);
        assert_eq!(parsed.to_record_bytes()[8], 0xfd);

        let mut created = ReviewingToolbarStates::new(false, false);
        created.set_show_reviewing_gallery(true);
        assert_eq!(created.to_record_bytes()[8], 0x02);
    }

    #[test]
    fn slide_list_table_round_trips_filetime_order_and_entries() {
        let table = SlideListTable10::new(vec![
            SlideCreationEntry::new(7, 0x1122_3344_5566_7788),
            SlideCreationEntry::new(u32::MAX, u64::MAX),
        ])
        .unwrap();
        let bytes = table.to_record_bytes().unwrap();
        let parsed_record = record(0x0f, RecordType::SlideListTable10, bytes[8..].to_vec());
        let parsed = SlideListTable10::parse(&parsed_record).unwrap();
        assert_eq!(parsed, table);
        assert_eq!(bytes[32..36], 0x1122_3344u32.to_le_bytes());
        assert_eq!(bytes[36..40], 0x5566_7788u32.to_le_bytes());
    }

    #[test]
    fn rejects_bad_headers_counts_order_truncation_and_trailing_data() {
        let entry = SlideCreationEntry::new(1, 2);
        let table = SlideListTable10::new(vec![entry]).unwrap();
        let bytes = table.to_record_bytes().unwrap();
        let payload = bytes[8..].to_vec();

        let mut cases = Vec::new();
        let mut negative = payload.clone();
        negative[8..12].copy_from_slice(&(-1i32).to_le_bytes());
        cases.push(negative);
        let mut mismatch = payload.clone();
        mismatch[8..12].copy_from_slice(&2i32.to_le_bytes());
        cases.push(mismatch);
        let mut wrong_child = payload.clone();
        wrong_child[14..16]
            .copy_from_slice(&RecordType::SlideListTableSize10Atom.as_u16().to_le_bytes());
        cases.push(wrong_child);
        cases.push(payload[..payload.len() - 1].to_vec());
        let mut trailing = payload.clone();
        trailing.push(0);
        cases.push(trailing);

        for data in cases {
            assert!(
                SlideListTable10::parse(&record(0x0f, RecordType::SlideListTable10, data,))
                    .is_err()
            );
        }
        assert!(
            SlideListTable10::parse(&record(0, RecordType::SlideListTable10, payload,)).is_err()
        );
    }

    #[test]
    fn rejects_atom_children_and_oversized_builder() {
        let mut atom = record(0, RecordType::DocToolbarStates10Atom, vec![0]);
        atom.children
            .push(record(0, RecordType::Unknown, Vec::new()));
        assert!(ReviewingToolbarStates::parse(&atom).is_err());

        let oversized = vec![SlideCreationEntry::new(0, 0); MAX_SLIDE_LIST_ENTRIES + 1];
        assert!(SlideListTable10::new(oversized).is_err());
    }
}
