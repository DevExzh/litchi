//! Strict Word 97+ field-table (`Plcfld`) parsing.
//!
//! Implements [MS-DOC] sections 2.8.25, 2.9.88 through 2.9.90, and 2.9.110
//! for all seven field-bearing document stories. Field instructions and results
//! remain inert text ranges; this module never evaluates or executes them.

use super::super::package::{DocError, Result};
use super::fib::FileInformationBlock;

const FLD_SIZE: usize = 2;
const CP_SIZE: usize = 4;
const MAX_PLCFLD_BYTES: usize = 64 * 1024 * 1024;
const MAX_FIELD_MARKERS: usize = 1_000_000;

/// A Word subdocument with its own field-character PLCF.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum FieldStory {
    /// Main document text (`fcPlcfFldMom`, FIB pointer 16).
    Main,
    /// Header and footer text (`fcPlcfFldHdr`, pointer 17).
    Header,
    /// Footnote text (`fcPlcfFldFtn`, pointer 18).
    Footnote,
    /// Comment/annotation text (`fcPlcfFldAtn`, pointer 19).
    Comment,
    /// Endnote text (`fcPlcfFldEdn`, pointer 48).
    Endnote,
    /// Main-document textbox text (`fcPlcfFldTxbx`, pointer 57).
    Textbox,
    /// Header/footer textbox text (`fcPlcffldHdrTxbx`, pointer 59).
    HeaderTextbox,
}

impl FieldStory {
    /// Stories in FIB pointer order.
    pub const ALL: [Self; 7] = [
        Self::Main,
        Self::Header,
        Self::Footnote,
        Self::Comment,
        Self::Endnote,
        Self::Textbox,
        Self::HeaderTextbox,
    ];

    const fn pointer_index(self) -> usize {
        match self {
            Self::Main => 16,
            Self::Header => 17,
            Self::Footnote => 18,
            Self::Comment => 19,
            Self::Endnote => 48,
            Self::Textbox => 57,
            Self::HeaderTextbox => 59,
        }
    }

    fn character_count(self, fib: &FileInformationBlock) -> u32 {
        let range = match self {
            Self::Main => Some(fib.get_main_doc_range()),
            Self::Header => fib.get_header_range(),
            Self::Footnote => fib.get_footnote_range(),
            Self::Comment => fib.get_comment_range(),
            Self::Endnote => fib.get_endnote_range(),
            Self::Textbox => fib.get_textbox_range(),
            Self::HeaderTextbox => fib.get_header_textbox_range(),
        };
        range.map_or(0, |(start, end)| end.saturating_sub(start))
    }
}

/// The field type stored in a begin-marker `Fld.grffld` byte.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum FieldType {
    Unparsed,
    RefWithoutKeyword,
    Ref,
    FootnoteRef,
    Set,
    If,
    Index,
    StyleRef,
    Sequence,
    TableOfContents,
    Info,
    Title,
    Subject,
    Author,
    Keywords,
    Comments,
    LastSavedBy,
    CreateDate,
    SaveDate,
    PrintDate,
    RevisionNumber,
    EditTime,
    NumberOfPages,
    NumberOfWords,
    NumberOfCharacters,
    FileName,
    Template,
    Date,
    Time,
    Page,
    Formula,
    Quote,
    Include,
    PageRef,
    Ask,
    FillIn,
    Data,
    Next,
    NextIf,
    SkipIf,
    MergeRecord,
    Dde,
    DdeAuto,
    Glossary,
    Print,
    Equation,
    GoToButton,
    MacroButton,
    AutoNumOutline,
    AutoNumLegal,
    AutoNum,
    Import,
    Link,
    Symbol,
    EmbeddedObject,
    MergeField,
    UserName,
    UserInitials,
    UserAddress,
    BarCode,
    DocumentVariable,
    Section,
    SectionPages,
    IncludePicture,
    IncludeText,
    FileSize,
    FormText,
    FormCheckbox,
    NoteRef,
    TableOfAuthorities,
    MergeSequence,
    AutoText,
    Compare,
    AddIn,
    FormDropdown,
    Advance,
    DocumentProperty,
    Control,
    Hyperlink,
    AutoTextList,
    ListNumber,
    HtmlControl,
    BidiOutline,
    AddressBlock,
    GreetingLine,
    Shape,
    /// An identifier not listed by [MS-DOC] 2.9.90.
    Unknown(u8),
}

impl FieldType {
    /// Convert the typed value to its `flt` byte.
    pub const fn as_u8(self) -> u8 {
        match self {
            Self::Unparsed => 0x01,
            Self::RefWithoutKeyword => 0x02,
            Self::Ref => 0x03,
            Self::FootnoteRef => 0x05,
            Self::Set => 0x06,
            Self::If => 0x07,
            Self::Index => 0x08,
            Self::StyleRef => 0x0A,
            Self::Sequence => 0x0C,
            Self::TableOfContents => 0x0D,
            Self::Info => 0x0E,
            Self::Title => 0x0F,
            Self::Subject => 0x10,
            Self::Author => 0x11,
            Self::Keywords => 0x12,
            Self::Comments => 0x13,
            Self::LastSavedBy => 0x14,
            Self::CreateDate => 0x15,
            Self::SaveDate => 0x16,
            Self::PrintDate => 0x17,
            Self::RevisionNumber => 0x18,
            Self::EditTime => 0x19,
            Self::NumberOfPages => 0x1A,
            Self::NumberOfWords => 0x1B,
            Self::NumberOfCharacters => 0x1C,
            Self::FileName => 0x1D,
            Self::Template => 0x1E,
            Self::Date => 0x1F,
            Self::Time => 0x20,
            Self::Page => 0x21,
            Self::Formula => 0x22,
            Self::Quote => 0x23,
            Self::Include => 0x24,
            Self::PageRef => 0x25,
            Self::Ask => 0x26,
            Self::FillIn => 0x27,
            Self::Data => 0x28,
            Self::Next => 0x29,
            Self::NextIf => 0x2A,
            Self::SkipIf => 0x2B,
            Self::MergeRecord => 0x2C,
            Self::Dde => 0x2D,
            Self::DdeAuto => 0x2E,
            Self::Glossary => 0x2F,
            Self::Print => 0x30,
            Self::Equation => 0x31,
            Self::GoToButton => 0x32,
            Self::MacroButton => 0x33,
            Self::AutoNumOutline => 0x34,
            Self::AutoNumLegal => 0x35,
            Self::AutoNum => 0x36,
            Self::Import => 0x37,
            Self::Link => 0x38,
            Self::Symbol => 0x39,
            Self::EmbeddedObject => 0x3A,
            Self::MergeField => 0x3B,
            Self::UserName => 0x3C,
            Self::UserInitials => 0x3D,
            Self::UserAddress => 0x3E,
            Self::BarCode => 0x3F,
            Self::DocumentVariable => 0x40,
            Self::Section => 0x41,
            Self::SectionPages => 0x42,
            Self::IncludePicture => 0x43,
            Self::IncludeText => 0x44,
            Self::FileSize => 0x45,
            Self::FormText => 0x46,
            Self::FormCheckbox => 0x47,
            Self::NoteRef => 0x48,
            Self::TableOfAuthorities => 0x49,
            Self::MergeSequence => 0x4B,
            Self::AutoText => 0x4F,
            Self::Compare => 0x50,
            Self::AddIn => 0x51,
            Self::FormDropdown => 0x53,
            Self::Advance => 0x54,
            Self::DocumentProperty => 0x55,
            Self::Control => 0x57,
            Self::Hyperlink => 0x58,
            Self::AutoTextList => 0x59,
            Self::ListNumber => 0x5A,
            Self::HtmlControl => 0x5B,
            Self::BidiOutline => 0x5C,
            Self::AddressBlock => 0x5D,
            Self::GreetingLine => 0x5E,
            Self::Shape => 0x5F,
            Self::Unknown(value) => value,
        }
    }

    /// Whether this identifier is listed by [MS-DOC] 2.9.90.
    pub const fn is_specified(self) -> bool {
        !matches!(self, Self::Unknown(_))
    }
}

impl From<u8> for FieldType {
    fn from(value: u8) -> Self {
        match value {
            0x01 => Self::Unparsed,
            0x02 => Self::RefWithoutKeyword,
            0x03 => Self::Ref,
            0x05 => Self::FootnoteRef,
            0x06 => Self::Set,
            0x07 => Self::If,
            0x08 => Self::Index,
            0x0A => Self::StyleRef,
            0x0C => Self::Sequence,
            0x0D => Self::TableOfContents,
            0x0E => Self::Info,
            0x0F => Self::Title,
            0x10 => Self::Subject,
            0x11 => Self::Author,
            0x12 => Self::Keywords,
            0x13 => Self::Comments,
            0x14 => Self::LastSavedBy,
            0x15 => Self::CreateDate,
            0x16 => Self::SaveDate,
            0x17 => Self::PrintDate,
            0x18 => Self::RevisionNumber,
            0x19 => Self::EditTime,
            0x1A => Self::NumberOfPages,
            0x1B => Self::NumberOfWords,
            0x1C => Self::NumberOfCharacters,
            0x1D => Self::FileName,
            0x1E => Self::Template,
            0x1F => Self::Date,
            0x20 => Self::Time,
            0x21 => Self::Page,
            0x22 => Self::Formula,
            0x23 => Self::Quote,
            0x24 => Self::Include,
            0x25 => Self::PageRef,
            0x26 => Self::Ask,
            0x27 => Self::FillIn,
            0x28 => Self::Data,
            0x29 => Self::Next,
            0x2A => Self::NextIf,
            0x2B => Self::SkipIf,
            0x2C => Self::MergeRecord,
            0x2D => Self::Dde,
            0x2E => Self::DdeAuto,
            0x2F => Self::Glossary,
            0x30 => Self::Print,
            0x31 => Self::Equation,
            0x32 => Self::GoToButton,
            0x33 => Self::MacroButton,
            0x34 => Self::AutoNumOutline,
            0x35 => Self::AutoNumLegal,
            0x36 => Self::AutoNum,
            0x37 => Self::Import,
            0x38 => Self::Link,
            0x39 => Self::Symbol,
            0x3A => Self::EmbeddedObject,
            0x3B => Self::MergeField,
            0x3C => Self::UserName,
            0x3D => Self::UserInitials,
            0x3E => Self::UserAddress,
            0x3F => Self::BarCode,
            0x40 => Self::DocumentVariable,
            0x41 => Self::Section,
            0x42 => Self::SectionPages,
            0x43 => Self::IncludePicture,
            0x44 => Self::IncludeText,
            0x45 => Self::FileSize,
            0x46 => Self::FormText,
            0x47 => Self::FormCheckbox,
            0x48 => Self::NoteRef,
            0x49 => Self::TableOfAuthorities,
            0x4B => Self::MergeSequence,
            0x4F => Self::AutoText,
            0x50 => Self::Compare,
            0x51 => Self::AddIn,
            0x53 => Self::FormDropdown,
            0x54 => Self::Advance,
            0x55 => Self::DocumentProperty,
            0x57 => Self::Control,
            0x58 => Self::Hyperlink,
            0x59 => Self::AutoTextList,
            0x5A => Self::ListNumber,
            0x5B => Self::HtmlControl,
            0x5C => Self::BidiOutline,
            0x5D => Self::AddressBlock,
            0x5E => Self::GreetingLine,
            0x5F => Self::Shape,
            other => Self::Unknown(other),
        }
    }
}

/// The low five-bit field-character tag.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum FieldBoundary {
    Begin = 0x13,
    Separator = 0x14,
    End = 0x15,
}

/// Flags stored on a field end marker (`grffldEnd`).
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Hash)]
pub struct FieldEndFlags {
    pub differ: bool,
    pub zombie_embed: bool,
    pub results_dirty: bool,
    pub results_edited: bool,
    pub locked: bool,
    pub private_result: bool,
    pub nested: bool,
    pub has_separator: bool,
}

impl FieldEndFlags {
    fn from_byte(value: u8) -> Self {
        Self {
            differ: value & 0x01 != 0,
            zombie_embed: value & 0x02 != 0,
            results_dirty: value & 0x04 != 0,
            results_edited: value & 0x08 != 0,
            locked: value & 0x10 != 0,
            private_result: value & 0x20 != 0,
            nested: value & 0x40 != 0,
            has_separator: value & 0x80 != 0,
        }
    }

    fn to_byte(self) -> u8 {
        u8::from(self.differ)
            | (u8::from(self.zombie_embed) << 1)
            | (u8::from(self.results_dirty) << 2)
            | (u8::from(self.results_edited) << 3)
            | (u8::from(self.locked) << 4)
            | (u8::from(self.private_result) << 5)
            | (u8::from(self.nested) << 6)
            | (u8::from(self.has_separator) << 7)
    }
}

/// Typed interpretation of the second byte of an `Fld`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum FieldMarkerValue {
    Begin(FieldType),
    Separator { ignored: u8 },
    End(FieldEndFlags),
}

/// A complete two-byte `Fld` descriptor.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct FieldDescriptor {
    /// Reserved upper three bits of `fldch`, preserved for byte-stable writing.
    pub reserved_bits: u8,
    pub value: FieldMarkerValue,
}

impl FieldDescriptor {
    /// Parse exactly one `Fld`.
    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        if bytes.len() != FLD_SIZE {
            return Err(corrupted("Fld must contain exactly two bytes"));
        }
        let reserved_bits = bytes[0] >> 5;
        let value = match bytes[0] & 0x1F {
            0x13 => {
                let field_type = FieldType::from(bytes[1]);
                if !field_type.is_specified() {
                    return Err(corrupted("Fld begin marker uses an unspecified field type"));
                }
                FieldMarkerValue::Begin(field_type)
            },
            0x14 => FieldMarkerValue::Separator { ignored: bytes[1] },
            0x15 => FieldMarkerValue::End(FieldEndFlags::from_byte(bytes[1])),
            _ => return Err(corrupted("Fld has an invalid field-character tag")),
        };
        Ok(Self {
            reserved_bits,
            value,
        })
    }

    /// Serialize the exact two-byte descriptor, retaining ignored bits.
    pub fn to_bytes(self) -> [u8; 2] {
        let (boundary, second) = match self.value {
            FieldMarkerValue::Begin(field_type) => (FieldBoundary::Begin, field_type.as_u8()),
            FieldMarkerValue::Separator { ignored } => (FieldBoundary::Separator, ignored),
            FieldMarkerValue::End(flags) => (FieldBoundary::End, flags.to_byte()),
        };
        [
            (boundary as u8) | ((self.reserved_bits & 0x07) << 5),
            second,
        ]
    }

    pub fn boundary(&self) -> FieldBoundary {
        match self.value {
            FieldMarkerValue::Begin(_) => FieldBoundary::Begin,
            FieldMarkerValue::Separator { .. } => FieldBoundary::Separator,
            FieldMarkerValue::End(_) => FieldBoundary::End,
        }
    }

    pub fn is_begin(&self) -> bool {
        self.boundary() == FieldBoundary::Begin
    }

    pub fn is_separator(&self) -> bool {
        self.boundary() == FieldBoundary::Separator
    }

    pub fn is_end(&self) -> bool {
        self.boundary() == FieldBoundary::End
    }
}

/// One field-character marker at a story-relative character position.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct FieldMarker {
    pub position: u32,
    pub descriptor: FieldDescriptor,
}

/// A fully paired field. Instruction and result text remain unevaluated ranges.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Field {
    pub story: FieldStory,
    pub start_cp: u32,
    pub separator_cp: Option<u32>,
    pub end_cp: u32,
    pub field_type: FieldType,
    pub end_flags: FieldEndFlags,
    pub nesting_depth: usize,
    pub has_separator: bool,
}

impl Field {
    pub fn code_range(&self) -> (u32, u32) {
        (self.start_cp + 1, self.separator_cp.unwrap_or(self.end_cp))
    }

    pub fn result_range(&self) -> Option<(u32, u32)> {
        self.separator_cp.map(|separator| (separator + 1, self.end_cp))
    }

    pub fn is_embedded_object(&self) -> bool {
        self.field_type == FieldType::EmbeddedObject
    }
}

/// One parsed and validated story-local `Plcfld`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FieldStoryTable {
    story: FieldStory,
    markers: Vec<FieldMarker>,
    terminal_cp: u32,
    fields: Vec<Field>,
}

impl FieldStoryTable {
    /// Strictly parse one complete PLCF, validating its FieldList grammar.
    pub fn parse_plcf(story: FieldStory, story_length: u32, data: &[u8]) -> Result<Self> {
        if data.len() > MAX_PLCFLD_BYTES {
            return Err(corrupted("Plcfld exceeds the allocation limit"));
        }
        if data.len() < CP_SIZE || (data.len() - CP_SIZE) % (CP_SIZE + FLD_SIZE) != 0 {
            return Err(corrupted("Plcfld has an invalid byte length"));
        }
        let marker_count = (data.len() - CP_SIZE) / (CP_SIZE + FLD_SIZE);
        if marker_count > MAX_FIELD_MARKERS {
            return Err(corrupted("Plcfld contains too many field markers"));
        }
        let cp_bytes = marker_count
            .checked_add(1)
            .and_then(|count| count.checked_mul(CP_SIZE))
            .ok_or_else(|| corrupted("Plcfld CP array size overflow"))?;
        let mut markers = Vec::with_capacity(marker_count);
        let mut previous_cp = None;
        for index in 0..=marker_count {
            let offset = index * CP_SIZE;
            let cp = u32::from_le_bytes(data[offset..offset + CP_SIZE].try_into().unwrap());
            // The terminal PLCF CP is undefined and does not locate a field
            // character (MS-DOC 2.8.25), so only marker CPs are story-bounded.
            if index < marker_count && cp > story_length {
                return Err(corrupted("Plcfld CP exceeds its story character count"));
            }
            if previous_cp.is_some_and(|previous| cp <= previous) {
                return Err(corrupted("Plcfld CPs are not strictly increasing"));
            }
            previous_cp = Some(cp);
            if index < marker_count {
                let descriptor_offset = cp_bytes + index * FLD_SIZE;
                let descriptor = FieldDescriptor::from_bytes(
                    &data[descriptor_offset..descriptor_offset + FLD_SIZE],
                )?;
                markers.push(FieldMarker {
                    position: cp,
                    descriptor,
                });
            }
        }
        let terminal_cp = previous_cp.unwrap_or(0);
        let fields = build_fields(story, &markers)?;
        Ok(Self {
            story,
            markers,
            terminal_cp,
            fields,
        })
    }

    pub fn story(&self) -> FieldStory {
        self.story
    }

    pub fn markers(&self) -> &[FieldMarker] {
        &self.markers
    }

    pub fn terminal_cp(&self) -> u32 {
        self.terminal_cp
    }

    pub fn fields(&self) -> &[Field] {
        &self.fields
    }

    /// Serialize the PLCF deterministically, retaining ignored descriptor bits
    /// and the undefined terminal CP.
    pub fn to_plcf_bytes(&self) -> Result<Vec<u8>> {
        let size = self
            .markers
            .len()
            .checked_add(1)
            .and_then(|count| count.checked_mul(CP_SIZE))
            .and_then(|cp_bytes| {
                self.markers
                    .len()
                    .checked_mul(FLD_SIZE)
                    .and_then(|fld_bytes| cp_bytes.checked_add(fld_bytes))
            })
            .ok_or_else(|| corrupted("Plcfld serialization size overflow"))?;
        if size > MAX_PLCFLD_BYTES {
            return Err(corrupted("Plcfld exceeds the serialization limit"));
        }
        let mut output = Vec::with_capacity(size);
        for marker in &self.markers {
            output.extend_from_slice(&marker.position.to_le_bytes());
        }
        output.extend_from_slice(&self.terminal_cp.to_le_bytes());
        for marker in &self.markers {
            output.extend_from_slice(&marker.descriptor.to_bytes());
        }
        Ok(output)
    }
}

#[derive(Debug)]
struct OpenField {
    start_cp: u32,
    separator_cp: Option<u32>,
    field_type: FieldType,
    nesting_depth: usize,
}

fn build_fields(story: FieldStory, markers: &[FieldMarker]) -> Result<Vec<Field>> {
    let mut stack: Vec<OpenField> = Vec::new();
    let mut fields = Vec::with_capacity(markers.len() / 2);
    for marker in markers {
        match marker.descriptor.value {
            FieldMarkerValue::Begin(field_type) => stack.push(OpenField {
                start_cp: marker.position,
                separator_cp: None,
                field_type,
                nesting_depth: stack.len(),
            }),
            FieldMarkerValue::Separator { .. } => {
                let open = stack
                    .last_mut()
                    .ok_or_else(|| corrupted("FieldList has a separator outside a field"))?;
                if open.separator_cp.replace(marker.position).is_some() {
                    return Err(corrupted("FieldList has duplicate separators in one field"));
                }
            },
            FieldMarkerValue::End(flags) => {
                let open = stack
                    .pop()
                    .ok_or_else(|| corrupted("FieldList has an unmatched end marker"))?;
                let has_separator = open.separator_cp.is_some();
                if flags.has_separator != has_separator {
                    return Err(corrupted("grffldEnd.fHasSep disagrees with the FieldList"));
                }
                if flags.nested != !stack.is_empty() {
                    return Err(corrupted("grffldEnd.fNested disagrees with field containment"));
                }
                fields.push(Field {
                    story,
                    start_cp: open.start_cp,
                    separator_cp: open.separator_cp,
                    end_cp: marker.position,
                    field_type: open.field_type,
                    end_flags: flags,
                    nesting_depth: open.nesting_depth,
                    has_separator,
                });
            },
        }
    }
    if !stack.is_empty() {
        return Err(corrupted("FieldList has an unmatched begin marker"));
    }
    fields.sort_unstable_by_key(|field| field.start_cp);
    Ok(fields)
}

/// All present Word field PLCFs.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct FieldsTable {
    stories: Vec<FieldStoryTable>,
}

impl FieldsTable {
    /// Parse all seven story field tables from checked FIB ranges.
    pub fn parse(fib: &FileInformationBlock, table_stream: &[u8]) -> Result<Self> {
        let mut stories = Vec::with_capacity(FieldStory::ALL.len());
        for story in FieldStory::ALL {
            let Some((offset, length)) = fib.get_table_pointer(story.pointer_index()) else {
                continue;
            };
            if length == 0 {
                continue;
            }
            let start = usize::try_from(offset)
                .map_err(|_| corrupted("Plcfld offset does not fit usize"))?;
            let length = usize::try_from(length)
                .map_err(|_| corrupted("Plcfld length does not fit usize"))?;
            if length > MAX_PLCFLD_BYTES {
                return Err(corrupted("Plcfld exceeds the allocation limit"));
            }
            let end = start
                .checked_add(length)
                .ok_or_else(|| corrupted("Plcfld table range overflow"))?;
            let data = table_stream
                .get(start..end)
                .ok_or_else(|| corrupted("Plcfld table range is outside the Table stream"))?;
            stories.push(FieldStoryTable::parse_plcf(
                story,
                story.character_count(fib),
                data,
            )?);
        }
        Ok(Self { stories })
    }

    pub fn story(&self, story: FieldStory) -> Option<&FieldStoryTable> {
        self.stories.iter().find(|table| table.story == story)
    }

    pub fn stories(&self) -> &[FieldStoryTable] {
        &self.stories
    }

    pub fn fields(&self, story: FieldStory) -> &[Field] {
        self.story(story).map_or(&[], FieldStoryTable::fields)
    }

    /// Compatibility accessor for existing main-document users.
    pub fn main_document_fields(&self) -> &[Field] {
        self.fields(FieldStory::Main)
    }

    pub fn find_field_at_position(&self, cp: u32) -> Option<&Field> {
        self.find_field(FieldStory::Main, cp)
    }

    pub fn find_field(&self, story: FieldStory, cp: u32) -> Option<&Field> {
        self.fields(story)
            .iter()
            .find(|field| field.start_cp <= cp && cp <= field.end_cp)
    }

    pub fn get_embedded_object_fields(&self) -> Vec<&Field> {
        self.main_document_fields()
            .iter()
            .filter(|field| field.is_embedded_object())
            .collect()
    }
}

fn corrupted(message: impl Into<String>) -> DocError {
    DocError::Corrupted(message.into())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn plcf(cps: &[u32], descriptors: &[[u8; 2]]) -> Vec<u8> {
        assert_eq!(cps.len(), descriptors.len() + 1);
        let mut data = Vec::new();
        for cp in cps {
            data.extend_from_slice(&cp.to_le_bytes());
        }
        for descriptor in descriptors {
            data.extend_from_slice(descriptor);
        }
        data
    }

    #[test]
    fn nested_fields_flags_and_roundtrip_are_exact() {
        let data = plcf(
            &[1, 3, 5, 7, 9, 11, 13],
            &[
                [0x13, 0x07],
                [0x13, 0x21],
                [0x14, 0xFF],
                [0x15, 0xC1],
                [0x14, 0xA5],
                [0x15, 0xB1],
            ],
        );
        let table = FieldStoryTable::parse_plcf(FieldStory::Main, 20, &data).unwrap();
        assert_eq!(table.to_plcf_bytes().unwrap(), data);
        assert_eq!(table.fields().len(), 2);
        assert_eq!(table.fields()[0].field_type, FieldType::If);
        assert_eq!(table.fields()[0].separator_cp, Some(9));
        assert!(table.fields()[0].end_flags.locked);
        assert_eq!(table.fields()[1].field_type, FieldType::Page);
        assert_eq!(table.fields()[1].nesting_depth, 1);
        assert!(table.fields()[1].end_flags.nested);
    }

    #[test]
    fn malformed_plcf_and_fieldlist_matrix_is_rejected() {
        let valid = plcf(&[1, 3, 5, 7], &[[0x13, 0x21], [0x14, 0], [0x15, 0x80]]);
        assert!(FieldStoryTable::parse_plcf(FieldStory::Main, 7, &valid).is_ok());
        for invalid in [
            Vec::new(),
            vec![0; 5],
            plcf(&[1, 1], &[[0x13, 0x21]]),
            plcf(&[2, 1], &[[0x13, 0x21]]),
            plcf(&[1, 9], &[[0x13, 0x21]]),
            plcf(&[1, 3], &[[0x12, 0x21]]),
            plcf(&[1, 3], &[[0x13, 0x04]]),
            plcf(&[1, 3], &[[0x14, 0]]),
            plcf(&[1, 3], &[[0x15, 0]]),
            plcf(&[1, 3], &[[0x13, 0x21]]),
            plcf(&[1, 3, 5], &[[0x13, 0x21], [0x15, 0x80]]),
            plcf(
                &[1, 2, 3, 4, 5],
                &[[0x13, 0x21], [0x14, 0], [0x14, 0], [0x15, 0x80]],
            ),
            plcf(
                &[1, 2, 3, 4, 5],
                &[[0x13, 0x07], [0x13, 0x21], [0x15, 0], [0x15, 0]],
            ),
        ] {
            assert!(FieldStoryTable::parse_plcf(FieldStory::Main, 7, &invalid).is_err());
        }
    }

    #[test]
    fn all_end_flags_and_reserved_descriptor_bits_are_preserved() {
        let descriptor = FieldDescriptor::from_bytes(&[0xF5, 0xFF]).unwrap();
        assert_eq!(descriptor.reserved_bits, 7);
        let FieldMarkerValue::End(flags) = descriptor.value else {
            panic!("end descriptor");
        };
        assert!(flags.differ && flags.zombie_embed && flags.results_dirty);
        assert!(flags.results_edited && flags.locked && flags.private_result);
        assert!(flags.nested && flags.has_separator);
        assert_eq!(descriptor.to_bytes(), [0xF5, 0xFF]);
    }

    #[test]
    fn field_type_mapping_covers_specified_and_unknown_values() {
        assert_eq!(FieldType::from(0x3A), FieldType::EmbeddedObject);
        assert_eq!(FieldType::from(0x58), FieldType::Hyperlink);
        assert_eq!(FieldType::from(0x45), FieldType::FileSize);
        assert_eq!(FieldType::from(0x04), FieldType::Unknown(0x04));
    }
}

#[cfg(test)]
mod terminal_cp_regression_tests {
    use super::{FieldStory, FieldStoryTable};

    fn field_plcf(cps: &[u32]) -> Vec<u8> {
        let mut data = Vec::new();
        for cp in cps {
            data.extend_from_slice(&cp.to_le_bytes());
        }
        data.extend_from_slice(&[0x13, 0x21]);
        data.extend_from_slice(&[0x15, 0x00]);
        data
    }

    #[test]
    fn terminal_cp_is_not_story_bounded_but_marker_cps_are() {
        let data = field_plcf(&[1, 3, u32::MAX]);
        let parsed = FieldStoryTable::parse_plcf(FieldStory::Main, 3, &data)
            .expect("undefined terminal CP may exceed the story length");
        assert_eq!(parsed.terminal_cp, u32::MAX);
        assert_eq!(parsed.markers[0].position, 1);
        assert_eq!(parsed.markers[1].position, 3);

        let marker_outside_story = field_plcf(&[1, 4, 5]);
        assert!(FieldStoryTable::parse_plcf(FieldStory::Main, 3, &marker_outside_story).is_err());
    }
}
