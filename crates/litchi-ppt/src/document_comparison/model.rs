//! Typed semantic models for `PowerPoint` 10 document-comparison records.

use crate::package::{Error, Result};
use std::sync::Arc;

use super::validation::{validate_count, validate_reviewer_name};

/// Maximum supported nesting for `PowerPoint` 10 document-comparison records.
pub const POWERPOINT_DIFF_MAX_DEPTH: usize = 32;
/// Maximum number of diff records accepted in one comparison tree.
pub const POWERPOINT_DIFF_MAX_RECORDS: usize = 65_536;

/// Reviewing toolbar and gallery display state from `DocToolbarStates10Atom`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ReviewingToolbarStates {
    pub(super) show_reviewing_toolbar: bool,
    pub(super) show_reviewing_gallery: bool,
    pub(super) ignored_reserved_bits: u8,
}

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
    #[must_use]
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
    #[must_use]
    pub const fn as_u32(self) -> u32 {
        self as u32
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
    #[must_use]
    pub fn for_type(diff_type: DiffType) -> Self {
        Self::from_raw(diff_type, 0).0
    }

    pub(super) fn from_raw(diff_type: DiffType, raw: u32) -> (Self, u32) {
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
            DiffType::SlideList
            | DiffType::MasterList
            | DiffType::ShapeList
            | DiffType::SlideShow
            | DiffType::HeaderFooter
            | DiffType::NamedShow
            | DiffType::NamedShowList
            | DiffType::RecolorInfo
            | DiffType::ExternalObject
            | DiffType::TableList
            | DiffType::InteractiveInfo => (Self::None, 0),
        };
        (flags, raw & !mask)
    }

    pub(super) fn diff_type(self) -> Option<DiffType> {
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

    pub(super) fn to_raw(self) -> u32 {
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
    pub(super) index: bool,
    pub(super) diff_type: DiffType,
    pub(super) ignored_prefix: [u8; 3],
    pub(super) ignored_tail: u32,
}

impl DiffRecordHeaders {
    #[must_use]
    pub const fn new(index: bool, diff_type: DiffType) -> Self {
        Self {
            index,
            diff_type,
            ignored_prefix: [0; 3],
            ignored_tail: 0,
        }
    }

    #[must_use]
    pub const fn index(&self) -> bool {
        self.index
    }

    #[must_use]
    pub const fn diff_type(&self) -> DiffType {
        self.diff_type
    }

    #[must_use]
    pub const fn ignored_prefix(&self) -> [u8; 3] {
        self.ignored_prefix
    }

    #[must_use]
    pub const fn ignored_tail(&self) -> u32 {
        self.ignored_tail
    }
}

/// One recursively parsed `RT_Diff10` container.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DiffNode {
    pub(super) headers: DiffRecordHeaders,
    pub(super) flags: DiffFlags,
    pub(super) ignored_flag_bits: u32,
    pub(super) children: Vec<Self>,
}

impl DiffNode {
    /// Construct one diff node with typed display flags and child nodes.
    ///
    /// # Errors
    ///
    /// Returns an error if the flags do not match `diff_type` or the children
    /// violate the structural grammar for `diff_type`.
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

    #[must_use]
    pub const fn headers(&self) -> &DiffRecordHeaders {
        &self.headers
    }

    #[must_use]
    pub const fn diff_type(&self) -> DiffType {
        self.headers.diff_type
    }

    #[must_use]
    pub const fn flags(&self) -> DiffFlags {
        self.flags
    }

    #[must_use]
    pub const fn ignored_flag_bits(&self) -> u32 {
        self.ignored_flag_bits
    }

    /// Replace the typed display flags while retaining source-reserved bits.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_flags(&mut self, flags: DiffFlags) -> Result<()> {
        self.flags = flags;
        self.validate_node()
    }

    #[must_use]
    pub fn children(&self) -> &[Self] {
        &self.children
    }

    pub fn children_of_type(&self, diff_type: DiffType) -> impl Iterator<Item = &Self> {
        self.children
            .iter()
            .filter(move |child| child.diff_type() == diff_type)
    }
}

/// A complete `DiffTree10Container`, without dereferencing its reviewer document.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DiffTree10 {
    pub(super) reviewer_name: String,
    pub(super) document_diff: DiffNode,
}

impl DiffTree10 {
    /// Construct a reviewer tree from a reviewer name and its document diff.
    ///
    /// # Errors
    ///
    /// Returns an error if the reviewer name is not a valid
    /// `PrintableUnicodeString` or the root node is not a `DocDiff10`
    /// container.
    pub fn new(reviewer_name: String, document_diff: DiffNode) -> Result<Self> {
        validate_reviewer_name(&reviewer_name)?;
        if document_diff.diff_type() != DiffType::Document {
            return super::validation::corrupted("DiffTree10 root is not a DocDiff10 container");
        }
        Ok(Self {
            reviewer_name,
            document_diff,
        })
    }

    #[must_use]
    pub fn reviewer_name(&self) -> &str {
        &self.reviewer_name
    }

    #[must_use]
    pub const fn document_diff(&self) -> &DiffNode {
        &self.document_diff
    }

    /// Return the document-level display flags for this reviewer tree.
    #[must_use]
    pub const fn document_flags(&self) -> DiffFlags {
        self.document_diff.flags()
    }
}

/// One slide creation-time entry from a document-comparison slide list.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SlideCreationEntry {
    pub(super) slide_id_ref: u32,
    pub(super) file_time: u64,
}

impl SlideCreationEntry {
    #[must_use]
    pub const fn new(slide_id_ref: u32, file_time: u64) -> Self {
        Self {
            slide_id_ref,
            file_time,
        }
    }

    #[must_use]
    pub const fn slide_id_ref(self) -> u32 {
        self.slide_id_ref
    }

    /// Raw Windows `FILETIME` ticks retained without clock conversion.
    #[must_use]
    pub const fn file_time(self) -> u64 {
        self.file_time
    }
}

/// Strict `SlideListTable10Container` metadata.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct SlideListTable10 {
    pub(super) entries: Vec<SlideCreationEntry>,
}

impl SlideListTable10 {
    /// Construct a slide-list table from creation-time entries.
    ///
    /// # Errors
    ///
    /// Returns an error if the entry count exceeds the MS-PPT limit.
    pub fn new(entries: Vec<SlideCreationEntry>) -> Result<Self> {
        validate_count(entries.len())?;
        Ok(Self { entries })
    }

    #[must_use]
    pub fn entries(&self) -> &[SlideCreationEntry] {
        &self.entries
    }

    /// Append one creation-time entry to the table.
    ///
    /// # Errors
    ///
    /// Returns an error if the resulting entry count would exceed the MS-PPT
    /// limit.
    pub fn push(&mut self, entry: SlideCreationEntry) -> Result<()> {
        validate_count(self.entries.len().saturating_add(1))?;
        self.entries.push(entry);
        Ok(())
    }
}

/// Resource limits for one document-comparison snapshot.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    /// Maximum number of records in the captured document tree.
    pub max_records: usize,
    /// Maximum serialized document size.
    pub max_bytes: usize,
    /// Maximum number of records in the inert PP10 review payload.
    pub max_review_records: usize,
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            max_records: 65_536,
            max_bytes: 16 * 1024 * 1024,
            max_review_records: 65_536,
        }
    }
}

impl Limits {
    pub(crate) fn validate(self) -> Result<Self> {
        if self.max_records == 0 || self.max_bytes < 8 || self.max_review_records == 0 {
            return super::validation::corrupted("document-comparison limits allow no records");
        }
        Ok(self)
    }
}

/// An opaque record retained from the PP10 review payload.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Unknown {
    record_type: u16,
    version: u16,
    instance: u16,
    bytes: Arc<[u8]>,
}

impl Unknown {
    pub(crate) fn new(record_type: u16, version: u16, instance: u16, bytes: Vec<u8>) -> Self {
        Self {
            record_type,
            version,
            instance,
            bytes: Arc::from(bytes.into_boxed_slice()),
        }
    }

    /// Raw record type, including vendor-defined values.
    #[must_use]
    pub const fn record_type(&self) -> u16 {
        self.record_type
    }
    /// Record version nibble as represented by the source record.
    #[must_use]
    pub const fn version(&self) -> u16 {
        self.version
    }
    /// Record instance as represented by the source record.
    #[must_use]
    pub const fn instance(&self) -> u16 {
        self.instance
    }
    /// Exact serialized record bytes.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.bytes
    }
}

/// One ordered typed or opaque record in the document review payload.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Entry {
    /// Reviewing toolbar/gallery state.
    Toolbar(ReviewingToolbarStates),
    /// Slide creation-time table.
    SlideList(SlideListTable10),
    /// One inert recursive document-difference tree.
    Diff(DiffTree10),
    /// Opaque record not owned by this focused review facade.
    Unknown(Unknown),
}

/// A lossless semantic view of the bounded document review payload.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Review {
    pub(crate) entries: Vec<Entry>,
}

impl Review {
    /// Borrow ordered typed and opaque entries.
    #[must_use]
    pub fn entries(&self) -> &[Entry] {
        &self.entries
    }

    /// Return the typed reviewing-toolbar state, if present.
    #[must_use]
    pub fn toolbar(&self) -> Option<ReviewingToolbarStates> {
        self.entries.iter().find_map(|entry| match entry {
            Entry::Toolbar(value) => Some(*value),
            Entry::SlideList(_) | Entry::Diff(_) | Entry::Unknown(_) => None,
        })
    }

    /// Return the typed slide-list table, if present.
    #[must_use]
    pub fn slide_list(&self) -> Option<&SlideListTable10> {
        self.entries.iter().find_map(|entry| match entry {
            Entry::SlideList(value) => Some(value),
            Entry::Toolbar(_) | Entry::Diff(_) | Entry::Unknown(_) => None,
        })
    }

    /// Return all typed difference trees in source order.
    pub fn diff_trees(&self) -> impl Iterator<Item = &DiffTree10> {
        self.entries.iter().filter_map(|entry| match entry {
            Entry::Diff(value) => Some(value),
            Entry::Toolbar(_) | Entry::SlideList(_) | Entry::Unknown(_) => None,
        })
    }

    /// Return the `index`th reviewer tree in source order.
    #[must_use]
    pub fn diff_tree(&self, index: usize) -> Option<&DiffTree10> {
        self.diff_trees().nth(index)
    }

    /// Return the number of reviewer trees in this payload.
    #[must_use]
    pub fn diff_tree_count(&self) -> usize {
        self.diff_trees().count()
    }

    /// Whether this PP10 payload has no records.
    #[must_use]
    pub const fn is_empty(&self) -> bool {
        self.entries.is_empty()
    }

    /// Return opaque records retained in their source order.
    pub fn unknown_records(&self) -> impl Iterator<Item = &Unknown> {
        self.entries.iter().filter_map(|entry| match entry {
            Entry::Unknown(value) => Some(value),
            Entry::Toolbar(_) | Entry::SlideList(_) | Entry::Diff(_) => None,
        })
    }
}
