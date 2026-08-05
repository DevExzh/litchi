use super::codec::{
    parse_index_entry_field_parts, parse_referenced_document_field_parts,
    parse_table_of_authorities_entry_field_parts, parse_table_of_contents_entry_field_parts,
    private_field_opaque_instructions,
};
use crate::package::{DocError, Result};
use crate::parts::fib::FileInformationBlock;

pub(super) const FLD_SIZE: usize = 2;
pub(super) const CP_SIZE: usize = 4;
pub(super) const MAX_PLCFLD_BYTES: usize = 64 * 1024 * 1024;
pub(super) const MAX_FIELD_MARKERS: usize = 1_000_000;
const FIELD_BEGIN_CHARACTER: char = '\u{0013}';
const FIELD_SEPARATOR_CHARACTER: char = '\u{0014}';
const FIELD_END_CHARACTER: char = '\u{0015}';

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

    pub(super) const fn pointer_index(self) -> usize {
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

    pub(crate) fn range(self, fib: &FileInformationBlock) -> Option<(u32, u32)> {
        match self {
            Self::Main => Some(fib.get_main_doc_range()),
            Self::Header => fib.get_header_range(),
            Self::Footnote => fib.get_footnote_range(),
            Self::Comment => fib.get_comment_range(),
            Self::Endnote => fib.get_endnote_range(),
            Self::Textbox => fib.get_textbox_range(),
            Self::HeaderTextbox => fib.get_header_textbox_range(),
        }
    }

    pub(super) fn character_count(self, fib: &FileInformationBlock) -> u32 {
        self.range(fib)
            .map_or(0, |(start, end)| end.saturating_sub(start))
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
    ///
    /// The descriptor is preserved as opaque metadata so a newer producer does
    /// not make an otherwise well-formed field table unreadable.
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

    /// Map a stored field-instruction keyword to its native `flt` type.
    ///
    /// Keywords are ASCII case-insensitive. The five text-only field kinds
    /// excluded from `Plcfld` (`TC`, `TA`, `XE`, `RD`, and `PRIVATE`) return
    /// `None`, as do unrecognized keywords.
    pub fn from_keyword(keyword: &str) -> Option<Self> {
        const TYPES: &[(&str, FieldType)] = &[
            ("REF", FieldType::Ref),
            ("FTNREF", FieldType::FootnoteRef),
            ("SET", FieldType::Set),
            ("IF", FieldType::If),
            ("INDEX", FieldType::Index),
            ("STYLEREF", FieldType::StyleRef),
            ("SEQ", FieldType::Sequence),
            ("TOC", FieldType::TableOfContents),
            ("INFO", FieldType::Info),
            ("TITLE", FieldType::Title),
            ("SUBJECT", FieldType::Subject),
            ("AUTHOR", FieldType::Author),
            ("KEYWORDS", FieldType::Keywords),
            ("COMMENTS", FieldType::Comments),
            ("LASTSAVEDBY", FieldType::LastSavedBy),
            ("CREATEDATE", FieldType::CreateDate),
            ("SAVEDATE", FieldType::SaveDate),
            ("PRINTDATE", FieldType::PrintDate),
            ("REVNUM", FieldType::RevisionNumber),
            ("EDITTIME", FieldType::EditTime),
            ("NUMPAGES", FieldType::NumberOfPages),
            ("NUMWORDS", FieldType::NumberOfWords),
            ("NUMCHARS", FieldType::NumberOfCharacters),
            ("FILENAME", FieldType::FileName),
            ("TEMPLATE", FieldType::Template),
            ("DATE", FieldType::Date),
            ("TIME", FieldType::Time),
            ("PAGE", FieldType::Page),
            ("=", FieldType::Formula),
            ("QUOTE", FieldType::Quote),
            ("INCLUDE", FieldType::Include),
            ("PAGEREF", FieldType::PageRef),
            ("ASK", FieldType::Ask),
            ("FILLIN", FieldType::FillIn),
            ("DATA", FieldType::Data),
            ("NEXT", FieldType::Next),
            ("NEXTIF", FieldType::NextIf),
            ("SKIPIF", FieldType::SkipIf),
            ("MERGEREC", FieldType::MergeRecord),
            ("DDE", FieldType::Dde),
            ("DDEAUTO", FieldType::DdeAuto),
            ("GLOSSARY", FieldType::Glossary),
            ("PRINT", FieldType::Print),
            ("EQ", FieldType::Equation),
            ("GOTOBUTTON", FieldType::GoToButton),
            ("MACROBUTTON", FieldType::MacroButton),
            ("AUTONUMOUT", FieldType::AutoNumOutline),
            ("AUTONUMLGL", FieldType::AutoNumLegal),
            ("AUTONUM", FieldType::AutoNum),
            ("IMPORT", FieldType::Import),
            ("LINK", FieldType::Link),
            ("SYMBOL", FieldType::Symbol),
            ("EMBED", FieldType::EmbeddedObject),
            ("MERGEFIELD", FieldType::MergeField),
            ("USERNAME", FieldType::UserName),
            ("USERINITIALS", FieldType::UserInitials),
            ("USERADDRESS", FieldType::UserAddress),
            ("BARCODE", FieldType::BarCode),
            ("DOCVARIABLE", FieldType::DocumentVariable),
            ("SECTION", FieldType::Section),
            ("SECTIONPAGES", FieldType::SectionPages),
            ("INCLUDEPICTURE", FieldType::IncludePicture),
            ("INCLUDETEXT", FieldType::IncludeText),
            ("FILESIZE", FieldType::FileSize),
            ("FORMTEXT", FieldType::FormText),
            ("FORMCHECKBOX", FieldType::FormCheckbox),
            ("NOTEREF", FieldType::NoteRef),
            ("TOA", FieldType::TableOfAuthorities),
            ("MERGESEQ", FieldType::MergeSequence),
            ("AUTOTEXT", FieldType::AutoText),
            ("COMPARE", FieldType::Compare),
            ("ADDIN", FieldType::AddIn),
            ("FORMDROPDOWN", FieldType::FormDropdown),
            ("ADVANCE", FieldType::Advance),
            ("DOCPROPERTY", FieldType::DocumentProperty),
            ("CONTROL", FieldType::Control),
            ("HYPERLINK", FieldType::Hyperlink),
            ("AUTOTEXTLIST", FieldType::AutoTextList),
            ("LISTNUM", FieldType::ListNumber),
            ("HTMLCONTROL", FieldType::HtmlControl),
            ("BIDIOUTLINE", FieldType::BidiOutline),
            ("ADDRESSBLOCK", FieldType::AddressBlock),
            ("GREETINGLINE", FieldType::GreetingLine),
            ("SHAPE", FieldType::Shape),
        ];
        TYPES.iter().find_map(|(name, field_type)| {
            keyword.eq_ignore_ascii_case(name).then_some(*field_type)
        })
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
            0x13 => FieldMarkerValue::Begin(FieldType::from(bytes[1])),
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
        self.separator_cp
            .map(|separator| (separator + 1, self.end_cp))
    }

    pub fn is_embedded_object(&self) -> bool {
        self.field_type == FieldType::EmbeddedObject
    }
}

/// Stored text associated with a Word field.
///
/// MS-DOC section 2.8.25 defines the instruction as the text after the field
/// begin character and before its separator or end character. When a separator
/// exists, result holds the stored cached text after that separator. Neither
/// value is evaluated, refreshed, resolved, or otherwise activated.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FieldText {
    /// The paired field markers and their story-relative positions.
    pub field: Field,
    /// Stored field instruction text.
    pub instruction: String,
    /// Stored cached result text, if the field has a separator.
    pub result: Option<String>,
}

/// Stored text for a field whose characters are intentionally absent from a
/// `Plcfld` table.
///
/// MS-DOC excludes several marker-field types, including `TC`, from the
/// `aFld` array. This internal representation reconstructs their text ranges
/// directly from their stored field characters.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct NonPlcfFieldText<'a> {
    pub story: FieldStory,
    pub start_cp: u32,
    pub separator_cp: Option<u32>,
    pub end_cp: u32,
    pub instruction: &'a str,
    pub result: Option<&'a str>,
}

#[derive(Debug, Clone, Copy)]
struct OpenNonPlcf {
    pub(super) start_cp: u32,
    pub(super) instruction_start_byte: usize,
    pub(super) separator_cp: Option<u32>,
    pub(super) separator_start_byte: Option<usize>,
    pub(super) result_start_byte: Option<usize>,
}

/// Reconstruct balanced stored fields from the control characters embedded in
/// one story's text.
///
/// This is deliberately independent of `Plcfld`: callers use it only for
/// field kinds MS-DOC explicitly omits from that table. Unbalanced marker
/// characters are ignored so malformed opaque text does not prevent the
/// document from opening.
pub(crate) fn non_plcf_field_texts(story: FieldStory, text: &str) -> Vec<NonPlcfFieldText<'_>> {
    let mut fields = Vec::new();
    let mut open_fields = Vec::new();
    let mut cp = 0u32;
    let mut begin_count = 0usize;

    for (byte_index, character) in text.char_indices() {
        let next_byte = byte_index + character.len_utf8();
        match character {
            FIELD_BEGIN_CHARACTER => {
                begin_count += 1;
                if begin_count > MAX_FIELD_MARKERS {
                    return Vec::new();
                }
                open_fields.push(OpenNonPlcf {
                    start_cp: cp,
                    instruction_start_byte: next_byte,
                    separator_cp: None,
                    separator_start_byte: None,
                    result_start_byte: None,
                });
            },
            FIELD_SEPARATOR_CHARACTER => {
                if let Some(field) = open_fields.last_mut()
                    && field.separator_cp.is_none()
                {
                    field.separator_cp = Some(cp);
                    field.separator_start_byte = Some(byte_index);
                    field.result_start_byte = Some(next_byte);
                }
            },
            FIELD_END_CHARACTER => {
                if let Some(field) = open_fields.pop() {
                    let instruction_end = field.separator_start_byte.unwrap_or(byte_index);
                    let result = field
                        .result_start_byte
                        .map(|start| &text[start..byte_index]);
                    fields.push(NonPlcfFieldText {
                        story,
                        start_cp: field.start_cp,
                        separator_cp: field.separator_cp,
                        end_cp: cp,
                        instruction: &text[field.instruction_start_byte..instruction_end],
                        result,
                    });
                }
            },
            _ => {},
        }
        cp = cp.saturating_add(character.len_utf16() as u32);
    }

    fields.sort_unstable_by_key(|field| field.start_cp);
    fields
}

/// A typed, inert legacy Word `MACROBUTTON` field.
///
/// ECMA-376 Part 1 §17.16.5.34 defines two stored field arguments: a macro or
/// command name and the text or graphic used as its button.
///
/// This preserves the stored macro or command name, button text, cached
/// result, and field-marker state. It never resolves, loads, invokes, or
/// otherwise executes the named macro or command.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MacroButtonField {
    pub(super) field: Field,
    pub(super) instruction: String,
    pub(super) macro_name: String,
    pub(super) display_text: String,
    pub(super) cached_result: Option<String>,
}

impl MacroButtonField {
    /// Return the paired field markers and their story-relative positions.
    pub fn field(&self) -> &Field {
        &self.field
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the stored macro or command name without resolving or invoking it.
    pub fn macro_name(&self) -> &str {
        &self.macro_name
    }

    /// Return the stored button text.
    ///
    /// This is source metadata, not a generated result.
    pub fn display_text(&self) -> &str {
        &self.display_text
    }

    /// Return the stored cached field result, if present.
    ///
    /// This value is never regenerated from the macro or command.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a producer marked the stored result stale.
    pub fn is_dirty(&self) -> bool {
        self.field.end_flags.results_dirty
    }

    /// Whether a producer locked this field against refresh.
    pub fn is_locked(&self) -> bool {
        self.field.end_flags.locked
    }
}

/// A typed, inert legacy Word `GOTOBUTTON` field.
///
/// [MS-DOC] §2.9.90 identifies its native field-type byte, and ECMA-376 Part
/// 1 §17.16.5.23 defines two stored field arguments: a destination and the
/// text or graphic used as its button. This type exposes stored text only; it
/// never resolves a destination, changes the insertion point, or activates a
/// jump.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct GoToButtonField {
    pub(super) field: Field,
    pub(super) instruction: String,
    pub(super) target: String,
    pub(super) button_text: String,
    pub(super) cached_result: Option<String>,
}

impl GoToButtonField {
    /// Return the paired field markers and their story-relative positions.
    pub fn field(&self) -> &Field {
        &self.field
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the stored destination without resolving or navigating to it.
    ///
    /// A destination can be a bookmark, page reference, annotation, footnote,
    /// line, page, or section expression.
    pub fn target(&self) -> &str {
        &self.target
    }

    /// Return the stored text or graphic-label expression for the button.
    ///
    /// This is source metadata, not an activated control.
    pub fn button_text(&self) -> &str {
        &self.button_text
    }

    /// Return the stored cached field result, if present.
    ///
    /// This value is never regenerated from the destination.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a producer marked the stored result stale.
    pub fn is_dirty(&self) -> bool {
        self.field.end_flags.results_dirty
    }

    /// Whether a producer locked this field against refresh.
    pub fn is_locked(&self) -> bool {
        self.field.end_flags.locked
    }
}

/// The stored category of a legacy Word active-content field.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ActiveContentFieldKind {
    /// An `ADDIN` field that stores add-in-created data.
    AddIn,
    /// A `CONTROL` field that represents an OCX control.
    OcxControl,
    /// An `HTMLCONTROL` field that represents an HTML control.
    HtmlControl,
}

/// Typed, inert metadata for a legacy Word add-in or control field.
///
/// [MS-DOC] §2.9.90 identifies the native `ADDIN`, `CONTROL`, and
/// `HTMLCONTROL` field types. This type retains only the stored category,
/// instruction, cached result, and field state. It never loads an add-in,
/// instantiates an OCX or HTML control, invokes code, executes script, renders
/// a control, or accesses an external resource.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ActiveContentField {
    pub(super) field: Field,
    pub(super) instruction: String,
    pub(super) kind: ActiveContentFieldKind,
    pub(super) cached_result: Option<String>,
}

impl ActiveContentField {
    /// Return the paired field markers and their story-relative positions.
    pub fn field(&self) -> &Field {
        &self.field
    }

    /// Return the complete stored field instruction.
    ///
    /// This string remains opaque metadata and is never interpreted.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return whether this stores add-in, OCX-control, or HTML-control metadata.
    pub fn kind(&self) -> ActiveContentFieldKind {
        self.kind
    }

    /// Return the stored cached field result, if present.
    ///
    /// This value is never regenerated by loading or running content.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a producer marked the stored result stale.
    pub fn is_dirty(&self) -> bool {
        self.field.end_flags.results_dirty
    }

    /// Whether a producer locked this field against refresh.
    pub fn is_locked(&self) -> bool {
        self.field.end_flags.locked
    }
}

/// Typed, inert metadata for a legacy Word `PRINT` field.
///
/// [MS-DOC] §2.9.90 identifies native `PRINT` fields with type `0x30`.
/// This type retains opaque printer-instruction text, a cached result, and
/// field-marker state only. It never interprets printer-control codes, opens a
/// printer, sends output, changes print settings, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PrintField {
    pub(super) field: Field,
    pub(super) instruction: String,
    pub(super) printer_instructions: String,
    pub(super) cached_result: Option<String>,
}

impl PrintField {
    /// Return the paired field markers and their story-relative positions.
    pub fn field(&self) -> &Field {
        &self.field
    }

    /// Return the complete stored field instruction.
    ///
    /// This string remains opaque metadata and is never interpreted or sent to
    /// a printer.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the stored printer-instruction text after the `PRINT` keyword.
    ///
    /// This can include printer-control or PostScript text. It is never parsed,
    /// interpreted, or sent to a printer.
    pub fn printer_instructions(&self) -> &str {
        &self.printer_instructions
    }

    /// Return the stored cached field result, if present.
    ///
    /// This value is never regenerated by printing.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a producer marked the stored result stale.
    pub fn is_dirty(&self) -> bool {
        self.field.end_flags.results_dirty
    }

    /// Whether a producer locked this field against refresh.
    pub fn is_locked(&self) -> bool {
        self.field.end_flags.locked
    }
}

/// Typed, inert metadata for a legacy Word `EMBED` field.
///
/// [MS-DOC] §2.9.90 identifies native `EMBED` fields with type `0x3A`.
/// This type retains opaque object-instruction text, a cached result, and
/// field-marker state only. It never loads, inspects, deserializes, activates,
/// renders, or executes an embedded object, accesses an external resource, or
/// refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct EmbedField {
    pub(super) field: Field,
    pub(super) instruction: String,
    pub(super) object_instructions: String,
    pub(super) cached_result: Option<String>,
}

impl EmbedField {
    /// Return the paired field markers and their story-relative positions.
    pub fn field(&self) -> &Field {
        &self.field
    }

    /// Return the complete stored `EMBED` field instruction.
    ///
    /// This string remains opaque metadata and is never used to load or
    /// activate an object.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the stored opaque object-instruction text after `EMBED`.
    ///
    /// It is never parsed, used to locate an object, or used to load, inspect,
    /// deserialize, activate, render, or execute object content.
    pub fn object_instructions(&self) -> &str {
        &self.object_instructions
    }

    /// Return the stored cached field result, if present.
    ///
    /// This value is cached text only and is never regenerated from an object.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a producer marked the stored result stale.
    pub fn is_dirty(&self) -> bool {
        self.field.end_flags.results_dirty
    }

    /// Whether a producer locked this field against refresh.
    pub fn is_locked(&self) -> bool {
        self.field.end_flags.locked
    }
}

/// Typed, inert metadata for a legacy Word `BARCODE` field.
///
/// [MS-DOC] §2.9.90 identifies native `BARCODE` fields with type `0x3F`.
/// This type retains opaque barcode-instruction text, a cached result, and
/// field-marker state only. It never parses or validates barcode data or
/// symbology, generates or renders a barcode, accesses an external resource,
/// or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct BarcodeField {
    pub(super) field: Field,
    pub(super) instruction: String,
    pub(super) barcode_instructions: String,
    pub(super) cached_result: Option<String>,
}

impl BarcodeField {
    /// Return the paired field markers and their story-relative positions.
    pub fn field(&self) -> &Field {
        &self.field
    }

    /// Return the complete stored `BARCODE` field instruction.
    ///
    /// This string remains opaque metadata and is never used to generate or
    /// render a barcode.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the stored opaque barcode-instruction text after `BARCODE`.
    ///
    /// It is never parsed, validated, interpreted, or used to generate or
    /// render barcode content.
    pub fn barcode_instructions(&self) -> &str {
        &self.barcode_instructions
    }

    /// Return the stored cached field result, if present.
    ///
    /// This value is cached text only and is never regenerated from barcode
    /// data.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a producer marked the stored result stale.
    pub fn is_dirty(&self) -> bool {
        self.field.end_flags.results_dirty
    }

    /// Whether a producer locked this field against refresh.
    pub fn is_locked(&self) -> bool {
        self.field.end_flags.locked
    }
}

/// Typed, inert metadata for a legacy Word `BIDIOUTLINE` field.
///
/// [MS-DOC] §2.9.90 identifies native `BIDIOUTLINE` fields with type
/// `0x5C`. This type retains opaque instruction text, a cached result, and
/// field-marker state only. It never reads right-to-left language, paragraph
/// outline, or layout state; chooses a numbering system; calculates a result;
/// or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct BidiOutlineField {
    pub(super) field: Field,
    pub(super) instruction: String,
    pub(super) opaque_instructions: String,
    pub(super) cached_result: Option<String>,
}

impl BidiOutlineField {
    /// Return the paired field markers and their story-relative positions.
    pub fn field(&self) -> &Field {
        &self.field
    }

    /// Return the complete stored `BIDIOUTLINE` field instruction.
    ///
    /// This string remains opaque metadata and is never used to calculate an
    /// outline number.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return opaque stored instruction text after `BIDIOUTLINE`.
    ///
    /// It is never parsed, interpreted, or used to resolve language, outline,
    /// numbering, or layout state.
    pub fn opaque_instructions(&self) -> &str {
        &self.opaque_instructions
    }

    /// Return the stored cached field result, if present.
    ///
    /// This value is cached text only and is never regenerated from document
    /// state.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a producer marked the stored result stale.
    pub fn is_dirty(&self) -> bool {
        self.field.end_flags.results_dirty
    }

    /// Whether a producer locked this field against refresh.
    pub fn is_locked(&self) -> bool {
        self.field.end_flags.locked
    }
}

/// Typed, inert metadata for a legacy Word `SHAPE` field.
///
/// [MS-DOC] §2.9.90 identifies native `SHAPE` fields with type `0x5F`.
/// Word uses this legacy field as a drawing-canvas anchor. This type retains
/// opaque instruction text, a cached result, and field-marker state only. It
/// never locates, links, loads, positions, lays out, or renders a drawing or
/// canvas, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ShapeField {
    pub(super) field: Field,
    pub(super) instruction: String,
    pub(super) opaque_instructions: String,
    pub(super) cached_result: Option<String>,
}

impl ShapeField {
    /// Return the paired field markers and their story-relative positions.
    pub fn field(&self) -> &Field {
        &self.field
    }

    /// Return the complete stored `SHAPE` field instruction.
    ///
    /// This string remains opaque metadata and is never used to locate or
    /// position a drawing canvas.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return opaque stored instruction text after `SHAPE`.
    ///
    /// It is never parsed, interpreted, or used to link a field to a drawing,
    /// resolve an anchor, or calculate layout.
    pub fn opaque_instructions(&self) -> &str {
        &self.opaque_instructions
    }

    /// Return the stored cached field result, if present.
    ///
    /// This value is cached metadata only and is never regenerated from a
    /// drawing canvas.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a producer marked the stored result stale.
    pub fn is_dirty(&self) -> bool {
        self.field.end_flags.results_dirty
    }

    /// Whether a producer locked this field against refresh.
    pub fn is_locked(&self) -> bool {
        self.field.end_flags.locked
    }
}

/// The stored kind of a legacy Word form-code field.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LegacyFormFieldKind {
    /// A `FORMTEXT` text-box form field.
    Text,
    /// A `FORMCHECKBOX` checkbox form field.
    CheckBox,
    /// A `FORMDROPDOWN` drop-down-list form field.
    DropDown,
}

/// Typed, inert metadata for a legacy Word form-code field.
///
/// [MS-DOC] §2.9.90 identifies native `FORMTEXT`, `FORMCHECKBOX`, and
/// `FORMDROPDOWN` fields with types `0x46`, `0x47`, and `0x53`. This type
/// retains only the stored kind, opaque instruction text, cached result, and
/// field-marker state, plus the stored `FFData` form state when it could be
/// located and parsed. It never fills a form, changes a selection or checkbox
/// state, invokes entry or exit macros, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct LegacyFormField {
    pub(super) field: Field,
    pub(super) instruction: String,
    pub(super) kind: LegacyFormFieldKind,
    pub(super) opaque_instructions: String,
    pub(super) cached_result: Option<String>,
    pub(super) form_data: Option<crate::parts::form_fields::FormFieldData>,
}

impl LegacyFormField {
    /// Return the paired field markers and their story-relative positions.
    pub fn field(&self) -> &Field {
        &self.field
    }

    /// Return the complete stored field instruction.
    ///
    /// This string remains opaque metadata and is never used to change a form
    /// field.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return whether this is a text, checkbox, or drop-down form-code field.
    pub const fn kind(&self) -> LegacyFormFieldKind {
        self.kind
    }

    /// Return opaque stored instruction text after the form-code keyword.
    ///
    /// It is never parsed, interpreted, or used to fill a form, change a
    /// checkbox or selection, or invoke a macro.
    pub fn opaque_instructions(&self) -> &str {
        &self.opaque_instructions
    }

    /// Return the stored cached field result, if present.
    ///
    /// This value is cached metadata only and is never regenerated from form
    /// state.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Return the parsed stored form state (`FFData`, MS-DOC 2.9.78), when the
    /// field's `NilPICFAndBinData` could be located in the Data stream and was
    /// well-formed.
    ///
    /// The returned data is inert: entry and exit macro names are stored
    /// verbatim and never invoked, the form is never filled, and checkbox or
    /// selection state is never changed. Fields constructed without Data
    /// stream access (or whose stored binary data is invalid, which MS-DOC
    /// §2.9.158 says MUST be ignored) return `None`.
    pub fn form_data(&self) -> Option<&crate::parts::form_fields::FormFieldData> {
        self.form_data.as_ref()
    }

    /// Attach the parsed stored form state. Crate-internal: only the document
    /// layer can locate the `NilPICFAndBinData` in the Data stream.
    pub(crate) fn set_form_data(
        &mut self,
        form_data: Option<crate::parts::form_fields::FormFieldData>,
    ) {
        self.form_data = form_data;
    }

    /// Whether a producer marked the stored result stale.
    pub fn is_dirty(&self) -> bool {
        self.field.end_flags.results_dirty
    }

    /// Whether a producer locked this field against refresh.
    pub fn is_locked(&self) -> bool {
        self.field.end_flags.locked
    }
}

/// One recognized stored option of a legacy Word `TOC` field.
///
/// These values retain how a producer configured a table of contents. They
/// are metadata only: this crate never scans entries, paginates, generates a
/// table, follows links, or refreshes the field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum TableOfContentsOption {
    /// The `\\a` caption label whose item labels and numbers are omitted.
    CaptionWithoutLabel(String),
    /// The `\\b` bookmark that bounds included entries.
    Bookmark(String),
    /// The `\\c` sequence identifier for a table of captions.
    CaptionSequence(String),
    /// The `\\d` separator between sequence and page numbers.
    SequencePageSeparator(String),
    /// The `\\f` contents-entry identifier that selects entries.
    TableEntryIdentifier(String),
    /// The `\\h` switch requests hyperlinks for entries.
    Hyperlinks,
    /// The `\\l` range of contents-entry levels to include.
    TableEntryLevels(String),
    /// The `\\n` switch omits page numbers, optionally for an entry-level range.
    OmitPageNumbers(Option<String>),
    /// The `\\o` built-in heading-style range, or all used heading levels.
    HeadingStyleRange(Option<String>),
    /// The `\\p` separator between an entry and its page number.
    EntryPageNumberSeparator(String),
    /// The `\\s` sequence identifier whose number prefixes page numbers.
    SequenceIdentifier(String),
    /// The `\\t` custom style-name/contents-level mappings.
    StyleMappings(String),
    /// The `\\u` switch uses applied paragraph outline levels.
    OutlineLevels,
    /// The `\\w` switch preserves tab characters within entries.
    PreserveTabs,
    /// The `\\x` switch preserves newline characters within entries.
    PreserveNewlines,
    /// The `\\z` switch hides page numbers and leaders in Web Layout view.
    HidePageNumbersInWebLayout,
}

/// Typed, inert metadata for a legacy Word table-of-contents field.
///
/// [MS-DOC] §2.9.90 maps native `TOC` field markers to ECMA-376 Part 1
/// §17.16.5.68. This type exposes only stored configuration, unrecognized
/// switches, cached result, and field state. It never scans entries, reads
/// bookmarks, resolves links, paginates, regenerates a table of contents, or
/// refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TableOfContentsField {
    pub(super) field: Field,
    pub(super) instruction: String,
    pub(super) options: Vec<TableOfContentsOption>,
    pub(super) unknown_switches: Vec<MergeFieldSwitch>,
    pub(super) cached_result: Option<String>,
}

impl TableOfContentsField {
    /// Return the paired field markers and their story-relative positions.
    pub fn field(&self) -> &Field {
        &self.field
    }

    /// Return the complete stored `TOC` field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return recognized stored configuration options in source order.
    ///
    /// This metadata is never used to generate or update a table.
    pub fn options(&self) -> &[TableOfContentsOption] {
        &self.options
    }

    /// Return unrecognized stored switches in source order.
    ///
    /// They are retained without interpretation or execution.
    pub fn unknown_switches(&self) -> &[MergeFieldSwitch] {
        &self.unknown_switches
    }

    /// Return the stored cached field result, if present.
    ///
    /// This value is never regenerated through pagination or field evaluation.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a producer marked the stored result stale.
    pub fn is_dirty(&self) -> bool {
        self.field.end_flags.results_dirty
    }

    /// Whether a producer locked this field against refresh.
    pub fn is_locked(&self) -> bool {
        self.field.end_flags.locked
    }
}

/// One recognized stored option of a legacy Word `TC` field.
///
/// These values identify how the entry participates in a table of contents.
/// They are inert metadata only: this crate never changes hidden text,
/// calculates page numbers, or generates a table.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum TableOfContentsEntryOption {
    /// The `\\f` contents-list identifier.
    ListIdentifier(String),
    /// The `\\l` entry level.
    Level(String),
    /// The `\\n` switch omits the entry page number.
    OmitPageNumber,
}

/// Typed, inert metadata for a legacy Word table-of-contents entry (`TC`)
/// field.
///
/// MS-DOC excludes `TC` field characters from the `Plcfld` `aFld` array, so
/// this type retains story-relative control-character positions instead of a
/// `Field` descriptor. It exposes only the stored entry, switches, and cached
/// result. It never changes hidden text, calculates page numbers, generates a
/// table of contents, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TableOfContentsEntryField {
    pub(super) story: FieldStory,
    pub(super) start_cp: u32,
    pub(super) separator_cp: Option<u32>,
    pub(super) end_cp: u32,
    pub(super) instruction: String,
    pub(super) entry: String,
    pub(super) options: Vec<TableOfContentsEntryOption>,
    pub(super) unknown_switches: Vec<MergeFieldSwitch>,
    pub(super) cached_result: Option<String>,
}

impl TableOfContentsEntryField {
    pub(crate) fn from_non_plcf_field(field: &NonPlcfFieldText<'_>) -> Option<Self> {
        let parts = parse_table_of_contents_entry_field_parts(field.instruction)?;
        Some(Self {
            story: field.story,
            start_cp: field.start_cp,
            separator_cp: field.separator_cp,
            end_cp: field.end_cp,
            instruction: field.instruction.to_string(),
            entry: parts.entry,
            options: parts.options,
            unknown_switches: parts.unknown_switches,
            cached_result: field.result.map(str::to_string),
        })
    }

    /// Return the story that stores this field.
    pub const fn story(&self) -> FieldStory {
        self.story
    }

    /// Return the story-relative position of this field's begin character.
    pub const fn start_position(&self) -> u32 {
        self.start_cp
    }

    /// Return the story-relative position of this field's separator character.
    ///
    /// `TC` fields normally have no cached result and therefore no separator.
    pub const fn separator_position(&self) -> Option<u32> {
        self.separator_cp
    }

    /// Return the story-relative position of this field's end character.
    pub const fn end_position(&self) -> u32 {
        self.end_cp
    }

    /// Return the complete stored `TC` field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the stored text marked for a table of contents.
    ///
    /// This is metadata only and is never inserted into generated content.
    pub fn entry(&self) -> &str {
        &self.entry
    }

    /// Return recognized `TC` options in stored source order.
    ///
    /// These options are never used to calculate page numbers, change hidden
    /// text, or update a table of contents.
    pub fn options(&self) -> &[TableOfContentsEntryOption] {
        &self.options
    }

    /// Return unrecognized stored field switches in source order.
    ///
    /// They are retained without interpretation or execution.
    pub fn unknown_switches(&self) -> &[MergeFieldSwitch] {
        &self.unknown_switches
    }

    /// Return the stored field result when a producer supplied one.
    ///
    /// This value is never regenerated through pagination or field evaluation.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }
}

/// One recognized stored option of a legacy Word `TA` field.
///
/// These values describe a legal-authority entry marker. They are inert
/// metadata only: this crate never finds cited text, changes hidden text,
/// follows bookmarks, calculates page numbers, or generates a table.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum TableOfAuthoritiesEntryOption {
    /// The `\\b` switch requests bold page-number formatting.
    BoldPageNumber,
    /// The `\\c` authority category.
    Category(String),
    /// The `\\i` switch requests italic page-number formatting.
    ItalicPageNumber,
    /// The `\\l` long citation text.
    LongCitation(String),
    /// The `\\r` bookmark that marks the cited page range.
    PageRangeBookmark(String),
    /// The `\\s` short citation text.
    ShortCitation(String),
}

/// Typed, inert metadata for a legacy Word table-of-authorities entry (`TA`)
/// field.
///
/// MS-DOC excludes `TA` field characters from the `Plcfld` `aFld` array, so
/// this type retains story-relative control-character positions instead of a
/// `Field` descriptor. It exposes only the stored switches and cached result.
/// It never finds citations, changes hidden text, follows bookmarks, calculates
/// page numbers, generates a table of authorities, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TableOfAuthoritiesEntryField {
    pub(super) story: FieldStory,
    pub(super) start_cp: u32,
    pub(super) separator_cp: Option<u32>,
    pub(super) end_cp: u32,
    pub(super) instruction: String,
    pub(super) options: Vec<TableOfAuthoritiesEntryOption>,
    pub(super) unknown_switches: Vec<MergeFieldSwitch>,
    pub(super) cached_result: Option<String>,
}

impl TableOfAuthoritiesEntryField {
    pub(crate) fn from_non_plcf_field(field: &NonPlcfFieldText<'_>) -> Option<Self> {
        let parts = parse_table_of_authorities_entry_field_parts(field.instruction)?;
        Some(Self {
            story: field.story,
            start_cp: field.start_cp,
            separator_cp: field.separator_cp,
            end_cp: field.end_cp,
            instruction: field.instruction.to_string(),
            options: parts.options,
            unknown_switches: parts.unknown_switches,
            cached_result: field.result.map(str::to_string),
        })
    }

    /// Return the story that stores this field.
    pub const fn story(&self) -> FieldStory {
        self.story
    }

    /// Return the story-relative position of this field's begin character.
    pub const fn start_position(&self) -> u32 {
        self.start_cp
    }

    /// Return the story-relative position of this field's separator character.
    ///
    /// `TA` fields normally have no cached result and therefore no separator.
    pub const fn separator_position(&self) -> Option<u32> {
        self.separator_cp
    }

    /// Return the story-relative position of this field's end character.
    pub const fn end_position(&self) -> u32 {
        self.end_cp
    }

    /// Return the complete stored `TA` field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return recognized `TA` options in stored source order.
    ///
    /// These options are never used to find citations, calculate page numbers,
    /// or generate a table of authorities.
    pub fn options(&self) -> &[TableOfAuthoritiesEntryOption] {
        &self.options
    }

    /// Return unrecognized stored field switches in source order.
    ///
    /// They are retained without interpretation or execution.
    pub fn unknown_switches(&self) -> &[MergeFieldSwitch] {
        &self.unknown_switches
    }

    /// Return the stored field result when a producer supplied one.
    ///
    /// This value is never regenerated through pagination or field evaluation.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }
}

/// One recognized stored option of a legacy Word `XE` field.
///
/// These values identify how an index marker participates in an `INDEX` field.
/// They are inert metadata only: this crate never changes hidden text,
/// calculates page numbers, follows bookmarks, or generates an index.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum IndexEntryOption {
    /// The `\\b` switch requests bold page-number formatting.
    BoldPageNumber,
    /// The `\\f` entry type that selects this marker.
    EntryType(String),
    /// The `\\i` switch requests italic page-number formatting.
    ItalicPageNumber,
    /// The `\\r` bookmark that marks a page range.
    PageRangeBookmark(String),
    /// The `\\t` text that replaces a page number with a cross reference.
    CrossReference(String),
    /// The `\\y` yomi sorting text.
    Yomi(String),
}

/// Typed, inert metadata for a legacy Word index-entry (`XE`) field.
///
/// MS-DOC excludes `XE` field characters from the `Plcfld` `aFld` array, so
/// this type retains story-relative control-character positions instead of a
/// `Field` descriptor. It exposes only the stored entry, switches, and cached
/// result. It never changes hidden text, resolves a bookmark, calculates page
/// numbers, sorts entries, generates an index, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct IndexEntryField {
    pub(super) story: FieldStory,
    pub(super) start_cp: u32,
    pub(super) separator_cp: Option<u32>,
    pub(super) end_cp: u32,
    pub(super) instruction: String,
    pub(super) entry: String,
    pub(super) options: Vec<IndexEntryOption>,
    pub(super) unknown_switches: Vec<MergeFieldSwitch>,
    pub(super) cached_result: Option<String>,
}

impl IndexEntryField {
    pub(crate) fn from_non_plcf_field(field: &NonPlcfFieldText<'_>) -> Option<Self> {
        let parts = parse_index_entry_field_parts(field.instruction)?;
        Some(Self {
            story: field.story,
            start_cp: field.start_cp,
            separator_cp: field.separator_cp,
            end_cp: field.end_cp,
            instruction: field.instruction.to_string(),
            entry: parts.entry,
            options: parts.options,
            unknown_switches: parts.unknown_switches,
            cached_result: field.result.map(str::to_string),
        })
    }

    /// Return the story that stores this field.
    pub const fn story(&self) -> FieldStory {
        self.story
    }

    /// Return the story-relative position of this field's begin character.
    pub const fn start_position(&self) -> u32 {
        self.start_cp
    }

    /// Return the story-relative position of this field's separator character.
    ///
    /// `XE` fields normally have no cached result and therefore no separator.
    pub const fn separator_position(&self) -> Option<u32> {
        self.separator_cp
    }

    /// Return the story-relative position of this field's end character.
    pub const fn end_position(&self) -> u32 {
        self.end_cp
    }

    /// Return the complete stored `XE` field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the stored text marked for an index.
    ///
    /// This is metadata only and is never inserted into generated content.
    pub fn entry(&self) -> &str {
        &self.entry
    }

    /// Return recognized `XE` options in stored source order.
    ///
    /// These options are never used to change hidden text, resolve bookmarks,
    /// calculate pages, or generate an index.
    pub fn options(&self) -> &[IndexEntryOption] {
        &self.options
    }

    /// Return unrecognized stored field switches in source order.
    ///
    /// They are retained without interpretation or execution.
    pub fn unknown_switches(&self) -> &[MergeFieldSwitch] {
        &self.unknown_switches
    }

    /// Return the stored field result when a producer supplied one.
    ///
    /// This value is never regenerated through pagination or field evaluation.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }
}

/// Typed, inert metadata for a legacy Word referenced-document (`RD`) field.
///
/// MS-DOC excludes `RD` field characters from the `Plcfld` `aFld` array, so
/// this type retains story-relative control-character positions instead of a
/// `Field` descriptor. It exposes only the stored source, relative-path request,
/// switches, and cached result. It never opens, resolves, reads, imports,
/// refreshes, evaluates, or executes the referenced document.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ReferencedDocumentField {
    pub(super) story: FieldStory,
    pub(super) start_cp: u32,
    pub(super) separator_cp: Option<u32>,
    pub(super) end_cp: u32,
    pub(super) instruction: String,
    pub(super) source: String,
    pub(super) relative_path: bool,
    pub(super) switches: Vec<MergeFieldSwitch>,
    pub(super) cached_result: Option<String>,
}

impl ReferencedDocumentField {
    pub(crate) fn from_non_plcf_field(field: &NonPlcfFieldText<'_>) -> Option<Self> {
        let parts = parse_referenced_document_field_parts(field.instruction)?;
        Some(Self {
            story: field.story,
            start_cp: field.start_cp,
            separator_cp: field.separator_cp,
            end_cp: field.end_cp,
            instruction: field.instruction.to_string(),
            source: parts.source,
            relative_path: parts.relative_path,
            switches: parts.switches,
            cached_result: field.result.map(str::to_string),
        })
    }

    /// Return the story that stores this field.
    pub const fn story(&self) -> FieldStory {
        self.story
    }

    /// Return the story-relative position of this field's begin character.
    pub const fn start_position(&self) -> u32 {
        self.start_cp
    }

    /// Return the story-relative position of this field's separator character.
    ///
    /// `RD` fields normally have no cached result and therefore no separator.
    pub const fn separator_position(&self) -> Option<u32> {
        self.separator_cp
    }

    /// Return the story-relative position of this field's end character.
    pub const fn end_position(&self) -> u32 {
        self.end_cp
    }

    /// Return the complete stored `RD` field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the stored referenced-document path without opening it.
    pub fn source(&self) -> &str {
        &self.source
    }

    /// Whether the stored `RD` instruction's `\\f` switch requests a path relative
    /// to this document.
    ///
    /// This is metadata only. The API never resolves the path.
    pub fn uses_relative_path(&self) -> bool {
        self.relative_path
    }

    /// Return all stored field switches in source order.
    ///
    /// They are retained without interpretation or execution.
    pub fn switches(&self) -> &[MergeFieldSwitch] {
        &self.switches
    }

    /// Return the stored field result when a producer supplied one.
    ///
    /// This value is never regenerated by opening or updating a source.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }
}

/// Typed, inert metadata for a legacy Word `PRIVATE` conversion-data field.
///
/// MS-DOC excludes `PRIVATE` field characters from the `Plcfld` `aFld` array,
/// so this type retains story-relative control-character positions instead of a
/// `Field` descriptor. It exposes only opaque stored instructions and a cached
/// result. It never converts a document, interprets field data, reveals hidden
/// content, changes layout, or refreshes a field. Despite its name, `PRIVATE` does
/// not provide confidentiality semantics.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PrivateField {
    pub(super) story: FieldStory,
    pub(super) start_cp: u32,
    pub(super) separator_cp: Option<u32>,
    pub(super) end_cp: u32,
    pub(super) instruction: String,
    pub(super) opaque_instructions: String,
    pub(super) cached_result: Option<String>,
}

impl PrivateField {
    pub(crate) fn from_non_plcf_field(field: &NonPlcfFieldText<'_>) -> Option<Self> {
        let opaque_instructions = private_field_opaque_instructions(field.instruction)?;
        Some(Self {
            story: field.story,
            start_cp: field.start_cp,
            separator_cp: field.separator_cp,
            end_cp: field.end_cp,
            instruction: field.instruction.to_string(),
            opaque_instructions,
            cached_result: field.result.map(str::to_string),
        })
    }

    /// Return the story that stores this field.
    pub const fn story(&self) -> FieldStory {
        self.story
    }

    /// Return the story-relative position of this field's begin character.
    pub const fn start_position(&self) -> u32 {
        self.start_cp
    }

    /// Return the story-relative position of this field's separator character.
    ///
    /// `PRIVATE` fields normally have no cached result and therefore no separator.
    pub const fn separator_position(&self) -> Option<u32> {
        self.separator_cp
    }

    /// Return the story-relative position of this field's end character.
    pub const fn end_position(&self) -> u32 {
        self.end_cp
    }

    /// Return the complete stored `PRIVATE` field instruction.
    ///
    /// This string remains opaque metadata and is never used to convert a
    /// document or reveal hidden content.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return opaque stored instruction text after `PRIVATE`.
    ///
    /// It is never parsed, interpreted, or used to convert a document, reveal
    /// hidden content, or calculate layout.
    pub fn opaque_instructions(&self) -> &str {
        &self.opaque_instructions
    }

    /// Return the stored field result when a producer supplied one.
    ///
    /// This value is never regenerated by conversion or used to change
    /// hidden-text visibility.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }
}

/// Typed fields whose marker characters MS-DOC excludes from every `Plcfld`.
///
/// MS-DOC §2.8.25 lists exactly five such field types: `TC`, `TA`, `XE`, `RD`,
/// and `PRIVATE`. This collection reconstructs them from balanced field
/// characters in stored story text. All values remain inert; no generated
/// table, index, referenced document, conversion payload, or cached result is
/// resolved, opened, interpreted, refreshed, or executed.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct NonPlcfFields {
    pub(super) table_of_contents_entries: Vec<TableOfContentsEntryField>,
    pub(super) table_of_authorities_entries: Vec<TableOfAuthoritiesEntryField>,
    pub(super) index_entries: Vec<IndexEntryField>,
    pub(super) referenced_documents: Vec<ReferencedDocumentField>,
    pub(super) private_fields: Vec<PrivateField>,
}

impl NonPlcfFields {
    pub(crate) fn from_story_texts<'a>(
        stories: impl IntoIterator<Item = (FieldStory, &'a str)>,
    ) -> Self {
        let mut output = Self::default();
        for (story, text) in stories {
            for field in non_plcf_field_texts(story, text) {
                if let Some(value) = TableOfContentsEntryField::from_non_plcf_field(&field) {
                    output.table_of_contents_entries.push(value);
                } else if let Some(value) =
                    TableOfAuthoritiesEntryField::from_non_plcf_field(&field)
                {
                    output.table_of_authorities_entries.push(value);
                } else if let Some(value) = IndexEntryField::from_non_plcf_field(&field) {
                    output.index_entries.push(value);
                } else if let Some(value) = ReferencedDocumentField::from_non_plcf_field(&field) {
                    output.referenced_documents.push(value);
                } else if let Some(value) = PrivateField::from_non_plcf_field(&field) {
                    output.private_fields.push(value);
                }
            }
        }
        output
    }

    /// Return stored `TC` table-of-contents entries in story and source order.
    pub fn table_of_contents_entries(&self) -> &[TableOfContentsEntryField] {
        &self.table_of_contents_entries
    }

    /// Return stored `TA` table-of-authorities entries in story and source order.
    pub fn table_of_authorities_entries(&self) -> &[TableOfAuthoritiesEntryField] {
        &self.table_of_authorities_entries
    }

    /// Return stored `XE` index entries in story and source order.
    pub fn index_entries(&self) -> &[IndexEntryField] {
        &self.index_entries
    }

    /// Return stored `RD` referenced-document fields without opening them.
    pub fn referenced_documents(&self) -> &[ReferencedDocumentField] {
        &self.referenced_documents
    }

    /// Return stored opaque `PRIVATE` conversion-data fields.
    pub fn private_fields(&self) -> &[PrivateField] {
        &self.private_fields
    }

    /// Whether no recognized excluded field is present.
    pub fn is_empty(&self) -> bool {
        self.table_of_contents_entries.is_empty()
            && self.table_of_authorities_entries.is_empty()
            && self.index_entries.is_empty()
            && self.referenced_documents.is_empty()
            && self.private_fields.is_empty()
    }

    /// Total number of recognized excluded fields.
    pub fn len(&self) -> usize {
        [
            self.table_of_contents_entries.len(),
            self.table_of_authorities_entries.len(),
            self.index_entries.len(),
            self.referenced_documents.len(),
            self.private_fields.len(),
        ]
        .into_iter()
        .fold(0usize, usize::saturating_add)
    }
}

/// One recognized stored option of a legacy Word `TOA` field.
///
/// These values retain how a producer configured a table of authorities. They
/// are metadata only: this crate never finds citations, follows bookmarks,
/// calculates page numbers, paginates, generates a table, or refreshes a
/// field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum TableOfAuthoritiesOption {
    /// The `\\b` bookmark that bounds included entries.
    Bookmark(String),
    /// The `\\c` authority category to include.
    Category(String),
    /// The `\\d` separator between sequence and page numbers.
    SequencePageSeparator(String),
    /// The `\\e` separator between an entry and its page number.
    EntryPageNumberSeparator(String),
    /// The `\\f` entry-formatting switch.
    EntryFormatting,
    /// The `\\g` separator between page numbers in a page range.
    PageRangeSeparator(String),
    /// The `\\h` switch includes category headings.
    CategoryHeadings,
    /// The `\\l` separator between multiple page references.
    PageReferenceSeparator(String),
    /// The `\\p` switch requests passim handling.
    UsePassim,
    /// The `\\s` sequence identifier whose number prefixes page numbers.
    SequenceIdentifier(String),
}

/// Typed, inert metadata for a legacy Word table-of-authorities field.
///
/// [MS-DOC] §2.9.90 maps native `TOA` field markers to ECMA-376 Part 1
/// §17.16.5.67. This type exposes only stored configuration, unrecognized
/// switches, cached result, and field state. It never finds citations, scans
/// hidden text, reads bookmarks, follows links, calculates page numbers,
/// paginates, regenerates a table of authorities, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TableOfAuthoritiesField {
    pub(super) field: Field,
    pub(super) instruction: String,
    pub(super) options: Vec<TableOfAuthoritiesOption>,
    pub(super) unknown_switches: Vec<MergeFieldSwitch>,
    pub(super) cached_result: Option<String>,
}

impl TableOfAuthoritiesField {
    /// Return the paired field markers and their story-relative positions.
    pub fn field(&self) -> &Field {
        &self.field
    }

    /// Return the complete stored `TOA` field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return recognized stored configuration options in source order.
    ///
    /// This metadata is never used to generate or update a table.
    pub fn options(&self) -> &[TableOfAuthoritiesOption] {
        &self.options
    }

    /// Return unrecognized stored switches in source order.
    ///
    /// They are retained without interpretation or execution.
    pub fn unknown_switches(&self) -> &[MergeFieldSwitch] {
        &self.unknown_switches
    }

    /// Return the stored cached field result, if present.
    ///
    /// This value is never regenerated through pagination or field evaluation.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a producer marked the stored result stale.
    pub fn is_dirty(&self) -> bool {
        self.field.end_flags.results_dirty
    }

    /// Whether a producer locked this field against refresh.
    pub fn is_locked(&self) -> bool {
        self.field.end_flags.locked
    }
}

/// One recognized stored option of a legacy Word `INDEX` field.
///
/// These values retain how a producer configured an index. They are metadata
/// only: this crate never scans index markers, reads bookmarks, calculates
/// page numbers, sorts entries, paginates, generates an index, or refreshes a
/// field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum IndexOption {
    /// The `\\b` bookmark that bounds included entries.
    Bookmark(String),
    /// The `\\c` requested number of index columns.
    Columns(String),
    /// The `\\d` separator between sequence and page numbers.
    SequencePageSeparator(String),
    /// The `\\e` separator between an entry and its first page number.
    EntryPageNumberSeparator(String),
    /// The `\\f` entry type that selects matching index markers.
    EntryType(String),
    /// The `\\g` separator between the start and end of a page range.
    PageRangeSeparator(String),
    /// The `\\h` heading text for each index-letter set.
    Heading(String),
    /// The `\\k` separator between an entry and its cross reference.
    CrossReferenceSeparator(String),
    /// The `\\l` separator between page numbers in a page-number list.
    PageNumberSeparator(String),
    /// Word's `\\o` East Asian sort-order extension, retained verbatim.
    EastAsianSortOrder(String),
    /// The `\\p` range of entry initial letters to include.
    LetterRange(String),
    /// The `\\r` switch runs subentries into their main-entry line.
    RunIn,
    /// The `\\s` sequence identifier whose number prefixes page numbers.
    SequenceIdentifier(String),
    /// The `\\y` switch enables yomi text for index entries.
    UseYomi,
    /// The `\\z` language identifier used to generate the index.
    LanguageId(String),
}

/// Typed, inert metadata for a legacy Word `INDEX` field.
///
/// [MS-DOC] §2.9.90 maps native `INDEX` field markers to ECMA-376 Part 1
/// §17.16.5.29. This type exposes only stored configuration, unrecognized
/// switches, cached result, and field state. It never scans index markers,
/// reads bookmarks, calculates page numbers, sorts entries, paginates,
/// generates an index, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct IndexField {
    pub(super) field: Field,
    pub(super) instruction: String,
    pub(super) options: Vec<IndexOption>,
    pub(super) unknown_switches: Vec<MergeFieldSwitch>,
    pub(super) cached_result: Option<String>,
}

impl IndexField {
    /// Return the paired field markers and their story-relative positions.
    pub fn field(&self) -> &Field {
        &self.field
    }

    /// Return the complete stored `INDEX` field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return recognized stored configuration options in source order.
    ///
    /// This metadata is never used to sort, generate, or update an index.
    pub fn options(&self) -> &[IndexOption] {
        &self.options
    }

    /// Return unrecognized stored switches in source order.
    ///
    /// They are retained without interpretation or execution.
    pub fn unknown_switches(&self) -> &[MergeFieldSwitch] {
        &self.unknown_switches
    }

    /// Return the stored cached field result, if present.
    ///
    /// This value is never regenerated through pagination or field evaluation.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a producer marked the stored result stale.
    pub fn is_dirty(&self) -> bool {
        self.field.end_flags.results_dirty
    }

    /// Whether a producer locked this field against refresh.
    pub fn is_locked(&self) -> bool {
        self.field.end_flags.locked
    }
}

/// The stored category of a legacy Word bookmark-reference field.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ReferenceFieldKind {
    /// A `REF` field.
    Reference,
    /// A `REF` field whose instruction omits the `REF` keyword.
    ReferenceWithoutKeyword,
    /// A `PAGEREF` field.
    PageReference,
    /// A historical `FTNREF` field.
    FootnoteReference,
    /// A `NOTEREF` field.
    NoteReference,
}

/// One recognized stored option of a legacy Word bookmark-reference field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ReferenceFieldOption {
    /// The `\\d` `REF` separator between sequence and page numbers.
    SequencePageSeparator(String),
    /// The `\\f` `REF` request for referenced note or comment content.
    ReferencedNoteContent,
    /// The `\\h` switch requests a hyperlink to the stored bookmark.
    Hyperlink,
    /// The `\\n` `REF` request for a paragraph number without context.
    ParagraphNumberWithoutContext,
    /// The `\\p` switch requests relative-position text.
    RelativePosition,
    /// The `\\r` `REF` request for a paragraph number in relative context.
    ParagraphNumberRelativeContext,
    /// The `\\t` `REF` request to suppress non-delimiter or non-numerical text.
    SuppressNonNumberText,
    /// The `\\w` `REF` request for a paragraph number in full context.
    ParagraphNumberFullContext,
    /// The `\\f` `NOTEREF` request to format the referenced note mark.
    NoteMarkFormatting,
}

/// Typed, inert metadata for a legacy Word bookmark-reference field.
///
/// [MS-DOC] §2.9.90 maps native `REF`, `PAGEREF`, `FTNREF`, and `NOTEREF`
/// field markers to ECMA-376 Part 1 §§17.16.5.51, 17.16.5.45, and
/// 17.16.5.40. This type exposes only stored category, bookmark name, options,
/// switches, cached result, and field state. It never looks up a bookmark,
/// reads a referenced range, resolves a page or note number, creates a link,
/// calculates a relative position, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ReferenceField {
    pub(super) field: Field,
    pub(super) instruction: String,
    pub(super) kind: ReferenceFieldKind,
    pub(super) bookmark: String,
    pub(super) options: Vec<ReferenceFieldOption>,
    pub(super) unknown_switches: Vec<MergeFieldSwitch>,
    pub(super) cached_result: Option<String>,
}

impl ReferenceField {
    /// Return the paired field markers and their story-relative positions.
    pub fn field(&self) -> &Field {
        &self.field
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the stored reference-field category.
    pub fn kind(&self) -> ReferenceFieldKind {
        self.kind
    }

    /// Return the stored bookmark name without resolving it.
    pub fn bookmark(&self) -> &str {
        &self.bookmark
    }

    /// Return recognized stored options in source order.
    ///
    /// This metadata is never used to navigate, resolve, or activate a link.
    pub fn options(&self) -> &[ReferenceFieldOption] {
        &self.options
    }

    /// Return unrecognized stored switches in source order.
    ///
    /// They are retained without interpretation or execution.
    pub fn unknown_switches(&self) -> &[MergeFieldSwitch] {
        &self.unknown_switches
    }

    /// Return the stored cached field result, if present.
    ///
    /// This value is never regenerated by resolving a bookmark or page number.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a producer marked the stored result stale.
    pub fn is_dirty(&self) -> bool {
        self.field.end_flags.results_dirty
    }

    /// Whether a producer locked this field against refresh.
    pub fn is_locked(&self) -> bool {
        self.field.end_flags.locked
    }
}

/// Typed, inert metadata for a legacy Word `SET` field.
///
/// [MS-DOC] §2.9.90 maps native `SET` field markers to ECMA-376 Part 1
/// §17.16.5.57. This type exposes only the stored target name, opaque
/// expression, cached result, and field state. It never evaluates the
/// expression, looks up or changes a bookmark, changes document state, or
/// refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SetField {
    pub(super) field: Field,
    pub(super) instruction: String,
    pub(super) target_name: String,
    pub(super) expression: String,
    pub(super) cached_result: Option<String>,
}

impl SetField {
    /// Return the paired field markers and their story-relative positions.
    pub fn field(&self) -> &Field {
        &self.field
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the stored target name without looking it up or changing it.
    pub fn target_name(&self) -> &str {
        &self.target_name
    }

    /// Return the opaque stored expression text.
    ///
    /// This text is never parsed, evaluated, or used to change document state.
    pub fn expression(&self) -> &str {
        &self.expression
    }

    /// Return the stored cached field result, if present.
    ///
    /// This value is never regenerated by evaluating the expression.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a producer marked the stored result stale.
    pub fn is_dirty(&self) -> bool {
        self.field.end_flags.results_dirty
    }

    /// Whether a producer locked this field against refresh.
    pub fn is_locked(&self) -> bool {
        self.field.end_flags.locked
    }
}

/// Typed, inert metadata for a legacy Word `=` (formula) field.
///
/// [MS-DOC] §2.9.90 maps native `=` field markers to ECMA-376 Part 1
/// §17.16.3.3. This type exposes only the stored optional formula, cached
/// result, and field state. It never parses or evaluates a formula, reads table
/// cells or bookmarks, resolves field values, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FormulaField {
    pub(super) field: Field,
    pub(super) instruction: String,
    pub(super) formula: Option<String>,
    pub(super) cached_result: Option<String>,
}

impl FormulaField {
    /// Return the paired field markers and their story-relative positions.
    pub fn field(&self) -> &Field {
        &self.field
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the opaque stored formula text after the leading `=`, if present.
    ///
    /// This text is never parsed, evaluated, or used to read table cells,
    /// bookmarks, or field values.
    pub fn formula(&self) -> Option<&str> {
        self.formula.as_deref()
    }

    /// Return the stored cached field result, if present.
    ///
    /// This value is never regenerated by evaluating a formula.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a producer marked the stored result stale.
    pub fn is_dirty(&self) -> bool {
        self.field.end_flags.results_dirty
    }

    /// Whether a producer locked this field against refresh.
    pub fn is_locked(&self) -> bool {
        self.field.end_flags.locked
    }
}

/// Typed, inert metadata for a legacy Word `EQ` equation field.
///
/// [MS-DOC] §2.9.90 maps native `EQ` field markers to ECMA-376 Part 4
/// §14.10.4.6. This type exposes only the stored opaque equation expression,
/// cached result, and field state. It never parses, calculates, formats,
/// renders, or refreshes an equation.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct EquationField {
    pub(super) field: Field,
    pub(super) instruction: String,
    pub(super) expression: String,
    pub(super) cached_result: Option<String>,
}

impl EquationField {
    /// Return the paired field markers and their story-relative positions.
    pub fn field(&self) -> &Field {
        &self.field
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the opaque equation expression after the `EQ` keyword.
    ///
    /// This syntax is never parsed, calculated, formatted, or rendered.
    pub fn expression(&self) -> &str {
        &self.expression
    }

    /// Return the stored cached field result, if present.
    ///
    /// This value is never regenerated from equation syntax.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a producer marked the stored result stale.
    pub fn is_dirty(&self) -> bool {
        self.field.end_flags.results_dirty
    }

    /// Whether a producer locked this field against refresh.
    pub fn is_locked(&self) -> bool {
        self.field.end_flags.locked
    }
}

/// Typed, inert metadata for a legacy Word `HYPERLINK` field.
///
/// [MS-DOC] §2.9.90 maps native `HYPERLINK` field markers to ECMA-376
/// Part 1 §17.16.5.25. This type exposes only stored link metadata, cached
/// results, and field state. It never opens, resolves, follows, activates, or
/// refreshes a link.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct HyperlinkField {
    pub(super) field: Field,
    pub(super) instruction: String,
    pub(super) external_target: Option<String>,
    pub(super) bookmark: Option<String>,
    pub(super) screen_tip: Option<String>,
    pub(super) target_frame: Option<String>,
    pub(super) appends_image_map_coordinates: bool,
    pub(super) opens_new_window: bool,
    pub(super) unknown_switches: Vec<MergeFieldSwitch>,
    pub(super) cached_result: Option<String>,
}

impl HyperlinkField {
    /// Return the paired field markers and their story-relative positions.
    pub fn field(&self) -> &Field {
        &self.field
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the stored external target without resolving or opening it.
    pub fn external_target(&self) -> Option<&str> {
        self.external_target.as_deref()
    }

    /// Return the stored internal bookmark target without resolving it.
    pub fn bookmark(&self) -> Option<&str> {
        self.bookmark.as_deref()
    }

    /// Return the stored screen-tip text, if present.
    ///
    /// This is metadata only and is never displayed by the library.
    pub fn screen_tip(&self) -> Option<&str> {
        self.screen_tip.as_deref()
    }

    /// Return the stored target frame, if present.
    ///
    /// This is metadata only and is never used to open a window or frame.
    pub fn target_frame(&self) -> Option<&str> {
        self.target_frame.as_deref()
    }

    /// Whether the target receives click coordinates for a server-side image map.
    ///
    /// This records producer intent only; no navigation or hit testing occurs.
    pub fn appends_image_map_coordinates(&self) -> bool {
        self.appends_image_map_coordinates
    }

    /// Whether the field requests opening its target in a new window.
    ///
    /// This records producer intent only; no window is opened.
    pub fn opens_new_window(&self) -> bool {
        self.opens_new_window
    }

    /// Return unrecognized stored switches in source order.
    ///
    /// They are retained without interpretation or execution.
    pub fn unknown_switches(&self) -> &[MergeFieldSwitch] {
        &self.unknown_switches
    }

    /// Return the stored cached field result, if present.
    ///
    /// This value is never regenerated by resolving a link.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a producer marked the stored result stale.
    pub fn is_dirty(&self) -> bool {
        self.field.end_flags.results_dirty
    }

    /// Whether a producer locked this field against refresh.
    pub fn is_locked(&self) -> bool {
        self.field.end_flags.locked
    }
}

/// Typed, inert metadata for a legacy Word `QUOTE` field.
///
/// [MS-DOC] §2.9.90 maps native `QUOTE` field markers to ECMA-376 Part 1
/// §17.16.5.49. This type exposes only the stored text argument, switches,
/// cached result, and field state. It never interprets character codes, expands
/// nested fields, inserts text, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct QuoteField {
    pub(super) field: Field,
    pub(super) instruction: String,
    pub(super) text: String,
    pub(super) switches: Vec<MergeFieldSwitch>,
    pub(super) cached_result: Option<String>,
}

impl QuoteField {
    /// Return the paired field markers and their story-relative positions.
    pub fn field(&self) -> &Field {
        &self.field
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the stored text argument without inserting or transforming it.
    pub fn text(&self) -> &str {
        &self.text
    }

    /// Return preserved switches in source order without interpreting them.
    pub fn switches(&self) -> &[MergeFieldSwitch] {
        &self.switches
    }

    /// Return the stored cached field result, if present.
    ///
    /// This value is never regenerated by inserting text.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a producer marked the stored result stale.
    pub fn is_dirty(&self) -> bool {
        self.field.end_flags.results_dirty
    }

    /// Whether a producer locked this field against refresh.
    pub fn is_locked(&self) -> bool {
        self.field.end_flags.locked
    }
}

/// Typed, inert metadata for a legacy Word `SYMBOL` field.
///
/// [MS-DOC] §2.9.90 maps native `SYMBOL` field markers to ECMA-376 Part 1
/// §17.16.5.61. This type exposes only the stored character argument, switches,
/// cached result, and field state. It never converts a character code, looks up
/// a font, inserts a glyph, changes formatting or layout, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SymbolField {
    pub(super) field: Field,
    pub(super) instruction: String,
    pub(super) character_argument: String,
    pub(super) switches: Vec<MergeFieldSwitch>,
    pub(super) cached_result: Option<String>,
}

impl SymbolField {
    /// Return the paired field markers and their story-relative positions.
    pub fn field(&self) -> &Field {
        &self.field
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the stored character argument without converting it to a glyph.
    pub fn character_argument(&self) -> &str {
        &self.character_argument
    }

    /// Return preserved switches in source order without interpreting them.
    pub fn switches(&self) -> &[MergeFieldSwitch] {
        &self.switches
    }

    /// Return the stored cached field result, if present.
    ///
    /// This value is never regenerated by mapping a character code or inserting
    /// a glyph.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a producer marked the stored result stale.
    pub fn is_dirty(&self) -> bool {
        self.field.end_flags.results_dirty
    }

    /// Whether a producer locked this field against refresh.
    pub fn is_locked(&self) -> bool {
        self.field.end_flags.locked
    }
}

/// The legacy Word automatic-numbering field category.
///
/// `AUTONUM`, `AUTONUMLGL`, and `AUTONUMOUT` are retained for
/// document compatibility. This enum preserves the stored field kind only; it
/// does not calculate a number, inspect paragraphs, headings, or styles, or
/// change layout.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum AutoNumberFieldKind {
    AutoNum,
    AutoNumLegal,
    AutoNumOutline,
}

impl AutoNumberFieldKind {
    /// The uppercase field keyword stored in a Word field instruction.
    pub const fn field_keyword(self) -> &'static str {
        match self {
            Self::AutoNum => "AUTONUM",
            Self::AutoNumLegal => "AUTONUMLGL",
            Self::AutoNumOutline => "AUTONUMOUT",
        }
    }

    pub(super) fn from_field_type(field_type: FieldType) -> Option<Self> {
        match field_type {
            FieldType::AutoNum => Some(Self::AutoNum),
            FieldType::AutoNumLegal => Some(Self::AutoNumLegal),
            FieldType::AutoNumOutline => Some(Self::AutoNumOutline),
            _ => None,
        }
    }

    pub(super) fn from_keyword(keyword: &str) -> Option<Self> {
        [Self::AutoNum, Self::AutoNumLegal, Self::AutoNumOutline]
            .into_iter()
            .find(|kind| keyword.eq_ignore_ascii_case(kind.field_keyword()))
    }
}

/// Typed, inert metadata for a legacy Word automatic-numbering field.
///
/// [MS-DOC] §2.9.90 maps native `flt` values 0x34 through 0x36 to
/// `AUTONUMOUT`, `AUTONUMLGL`, and `AUTONUM`. This type exposes only the
/// stored kind, switches, cached result, and field state. It never calculates
/// paragraph numbers, reads heading or style state, changes paragraphs or
/// layout, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AutoNumberField {
    pub(super) field: Field,
    pub(super) instruction: String,
    pub(super) kind: AutoNumberFieldKind,
    pub(super) switches: Vec<MergeFieldSwitch>,
    pub(super) cached_result: Option<String>,
}

impl AutoNumberField {
    /// Return the paired field markers and their story-relative positions.
    pub fn field(&self) -> &Field {
        &self.field
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the recognized automatic-numbering category.
    pub const fn kind(&self) -> AutoNumberFieldKind {
        self.kind
    }

    /// Return preserved switches in source order without interpreting them.
    pub fn switches(&self) -> &[MergeFieldSwitch] {
        &self.switches
    }

    /// Return the stored cached field result, if present.
    ///
    /// This value is never regenerated by calculating a paragraph number.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a producer marked the stored result stale.
    pub fn is_dirty(&self) -> bool {
        self.field.end_flags.results_dirty
    }

    /// Whether a producer locked this field against refresh.
    pub fn is_locked(&self) -> bool {
        self.field.end_flags.locked
    }
}

/// Typed, inert metadata for a legacy Word `LISTNUM` field.
///
/// [MS-DOC] §2.9.90 maps native `LISTNUM` field markers to ECMA-376 Part 1
/// §17.16.5.33. This type exposes only the stored optional list name, switches,
/// cached result, and field state. It never looks up a list, determines a level
/// or start value, calculates a number, changes layout, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ListNumberField {
    pub(super) field: Field,
    pub(super) instruction: String,
    pub(super) list_name: Option<String>,
    pub(super) switches: Vec<MergeFieldSwitch>,
    pub(super) cached_result: Option<String>,
}

impl ListNumberField {
    /// Return the paired field markers and their story-relative positions.
    pub fn field(&self) -> &Field {
        &self.field
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the stored optional list name without looking it up.
    pub fn list_name(&self) -> Option<&str> {
        self.list_name.as_deref()
    }

    /// Return preserved switches in source order without interpreting them.
    pub fn switches(&self) -> &[MergeFieldSwitch] {
        &self.switches
    }

    /// Return the stored cached field result, if present.
    ///
    /// This value is never regenerated by calculating a list number.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a producer marked the stored result stale.
    pub fn is_dirty(&self) -> bool {
        self.field.end_flags.results_dirty
    }

    /// Whether a producer locked this field against refresh.
    pub fn is_locked(&self) -> bool {
        self.field.end_flags.locked
    }
}

/// Typed, inert metadata for a legacy Word `SEQ` field.
///
/// [MS-DOC] §2.9.90 maps native `SEQ` field markers to ECMA-376 Part 1
/// §17.16.5.56. This type exposes only the stored identifier, optional bookmark,
/// opaque tail, cached result, and field state. It never looks up a bookmark,
/// increments or resets a sequence, calculates a number, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SequenceField {
    pub(super) field: Field,
    pub(super) instruction: String,
    pub(super) identifier: String,
    pub(super) bookmark: Option<String>,
    pub(super) tail: String,
    pub(super) cached_result: Option<String>,
}

impl SequenceField {
    /// Return the paired field markers and their story-relative positions.
    pub fn field(&self) -> &Field {
        &self.field
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the stored sequence identifier without calculating its value.
    pub fn identifier(&self) -> &str {
        &self.identifier
    }

    /// Return the optional stored bookmark name without looking it up.
    pub fn bookmark(&self) -> Option<&str> {
        self.bookmark.as_deref()
    }

    /// Return opaque stored text after the identifier and optional bookmark.
    ///
    /// This text is never parsed to change or calculate a sequence.
    pub fn tail(&self) -> &str {
        &self.tail
    }

    /// Return the stored cached field result, if present.
    ///
    /// This value is never regenerated by calculating a sequence number.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a producer marked the stored result stale.
    pub fn is_dirty(&self) -> bool {
        self.field.end_flags.results_dirty
    }

    /// Whether a producer locked this field against refresh.
    pub fn is_locked(&self) -> bool {
        self.field.end_flags.locked
    }
}

/// One recognized stored option of a legacy Word `STYLEREF` field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum StyleReferenceFieldOption {
    /// The `\\l` request for the nearest styled text following the field.
    FollowingText,
    /// The `\\n` request for the referenced paragraph number.
    ParagraphNumber,
    /// The `\\p` request for the referenced paragraph's relative position.
    RelativePosition,
    /// The `\\r` request for the referenced paragraph number in relative context.
    ParagraphNumberRelativeContext,
    /// The `\\t` request to suppress non-delimiter or non-numerical text.
    SuppressNonNumberText,
    /// The `\\w` request for the referenced paragraph number in full context.
    ParagraphNumberFullContext,
}

/// Typed, inert metadata for a legacy Word `STYLEREF` field.
///
/// [MS-DOC] §2.9.90 maps native `STYLEREF` field markers to ECMA-376 Part
/// 1 §17.16.5.59. This type exposes only stored style name, options, switches,
/// cached result, and field state. It never looks up styled text, searches
/// document stories, calculates paragraph numbers or relative positions,
/// resolves page layout, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct StyleReferenceField {
    pub(super) field: Field,
    pub(super) instruction: String,
    pub(super) style_name: String,
    pub(super) options: Vec<StyleReferenceFieldOption>,
    pub(super) unknown_switches: Vec<MergeFieldSwitch>,
    pub(super) cached_result: Option<String>,
}

impl StyleReferenceField {
    /// Return the paired field markers and their story-relative positions.
    pub fn field(&self) -> &Field {
        &self.field
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the stored style name without looking up text that uses it.
    pub fn style_name(&self) -> &str {
        &self.style_name
    }

    /// Return recognized stored options in source order.
    ///
    /// This metadata is never used to search text, calculate a number, or
    /// resolve layout.
    pub fn options(&self) -> &[StyleReferenceFieldOption] {
        &self.options
    }

    /// Return unrecognized stored switches in source order.
    ///
    /// They are retained without interpretation or execution.
    pub fn unknown_switches(&self) -> &[MergeFieldSwitch] {
        &self.unknown_switches
    }

    /// Return the stored cached field result, if present.
    ///
    /// This value is never regenerated by searching styled text or layout.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a producer marked the stored result stale.
    pub fn is_dirty(&self) -> bool {
        self.field.end_flags.results_dirty
    }

    /// Whether a producer locked this field against refresh.
    pub fn is_locked(&self) -> bool {
        self.field.end_flags.locked
    }
}

/// The stored category of a legacy Word building-block field.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AutoTextFieldKind {
    /// A historical `GLOSSARY` field.
    Glossary,
    /// An `AUTOTEXT` field.
    AutoText,
}

/// Typed, inert metadata for a legacy Word `GLOSSARY` or `AUTOTEXT` field.
///
/// [MS-DOC] §2.9.90 identifies `GLOSSARY` as identical to `AUTOTEXT` and maps
/// both native field types to ECMA-376 Part 1 §17.16.5.5. This type exposes
/// only the stored category, entry name, switches, cached result, and field
/// state. It never looks up a building block, reads a template, inserts
/// content, changes bookmarks, refreshes a field, or opens an external
/// resource.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AutoTextField {
    pub(super) field: Field,
    pub(super) instruction: String,
    pub(super) kind: AutoTextFieldKind,
    pub(super) entry_name: String,
    pub(super) unknown_switches: Vec<MergeFieldSwitch>,
    pub(super) cached_result: Option<String>,
}

impl AutoTextField {
    /// Return the paired field markers and their story-relative positions.
    pub fn field(&self) -> &Field {
        &self.field
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return whether the native marker is `GLOSSARY` or `AUTOTEXT`.
    pub fn kind(&self) -> AutoTextFieldKind {
        self.kind
    }

    /// Return the stored building-block entry name.
    ///
    /// This name is never resolved, looked up, or inserted.
    pub fn entry_name(&self) -> &str {
        &self.entry_name
    }

    /// Return unrecognized stored switches in source order.
    ///
    /// They are retained without interpretation or execution.
    pub fn unknown_switches(&self) -> &[MergeFieldSwitch] {
        &self.unknown_switches
    }

    /// Return the stored cached field result, if present.
    ///
    /// This value is never regenerated by looking up or inserting content.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a producer marked the stored result stale.
    pub fn is_dirty(&self) -> bool {
        self.field.end_flags.results_dirty
    }

    /// Whether a producer locked this field against refresh.
    pub fn is_locked(&self) -> bool {
        self.field.end_flags.locked
    }
}

/// One recognized stored option of a legacy Word `AUTOTEXTLIST` field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum AutoTextListOption {
    /// The `\\s` style name used to limit eligible building blocks.
    Style(String),
    /// The `\\t` stored tip text.
    Tip(String),
}

/// Typed, inert metadata for a legacy Word `AUTOTEXTLIST` field.
///
/// [MS-DOC] §2.9.90 maps this native field type to ECMA-376 Part 1
/// §17.16.5.6. This type exposes only stored display text, style/tip options,
/// unrecognized switches, cached result, and field state. It never shows a
/// selection UI, looks up eligible building blocks, reads a template, inserts
/// content, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AutoTextListField {
    pub(super) field: Field,
    pub(super) instruction: String,
    pub(super) display_text: Option<String>,
    pub(super) options: Vec<AutoTextListOption>,
    pub(super) unknown_switches: Vec<MergeFieldSwitch>,
    pub(super) cached_result: Option<String>,
}

impl AutoTextListField {
    /// Return the paired field markers and their story-relative positions.
    pub fn field(&self) -> &Field {
        &self.field
    }

    /// Return the complete stored `AUTOTEXTLIST` field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the optional stored display text.
    ///
    /// This text is metadata only and never triggers a selection UI.
    pub fn display_text(&self) -> Option<&str> {
        self.display_text.as_deref()
    }

    /// Return recognized stored options in source order.
    ///
    /// This metadata is never used to query, select, or insert a building
    /// block.
    pub fn options(&self) -> &[AutoTextListOption] {
        &self.options
    }

    /// Return unrecognized stored switches in source order.
    ///
    /// They are retained without interpretation or execution.
    pub fn unknown_switches(&self) -> &[MergeFieldSwitch] {
        &self.unknown_switches
    }

    /// Return the stored cached field result, if present.
    ///
    /// This value is never regenerated by selection or content insertion.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a producer marked the stored result stale.
    pub fn is_dirty(&self) -> bool {
        self.field.end_flags.results_dirty
    }

    /// Whether a producer locked this field against refresh.
    pub fn is_locked(&self) -> bool {
        self.field.end_flags.locked
    }
}

/// One stored switch in a legacy Word `MERGEFIELD` instruction.
///
/// The name excludes its leading backslash and is normalized to ASCII
/// lowercase. The argument is stored text only and is never resolved against a
/// mail-merge data source.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MergeFieldSwitch {
    pub(super) name: char,
    pub(super) argument: Option<String>,
}

impl MergeFieldSwitch {
    /// Return the switch character, without its leading backslash.
    pub fn name(&self) -> char {
        self.name
    }

    /// Return the stored switch argument, if present.
    pub fn argument(&self) -> Option<&str> {
        self.argument.as_deref()
    }
}

/// A typed, inert legacy Word `MERGEFIELD` field.
///
/// [MS-DOC] §2.9.90 identifies its field-type byte, and ECMA-376 Part 1
/// §17.16.5.35 defines one stored data-column name followed by optional
/// switches. This type exposes that persisted metadata and cached result only.
/// It never opens a data source, resolves a record, performs a merge, or
/// refreshes the field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MergeField {
    pub(super) field: Field,
    pub(super) instruction: String,
    pub(super) field_name: String,
    pub(super) switches: Vec<MergeFieldSwitch>,
    pub(super) cached_result: Option<String>,
}

impl MergeField {
    /// Return the paired field markers and their story-relative positions.
    pub fn field(&self) -> &Field {
        &self.field
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the stored data-column name without resolving a data source.
    pub fn field_name(&self) -> &str {
        &self.field_name
    }

    /// Return stored field switches in source order without interpreting them.
    pub fn switches(&self) -> &[MergeFieldSwitch] {
        &self.switches
    }

    /// Return the stored cached field result, if present.
    ///
    /// This value is never regenerated by a merge.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a producer marked the stored result stale.
    pub fn is_dirty(&self) -> bool {
        self.field.end_flags.results_dirty
    }

    /// Whether a producer locked this field against refresh.
    pub fn is_locked(&self) -> bool {
        self.field.end_flags.locked
    }

    /// Check whether a case-insensitive ASCII switch appears in this field.
    pub fn has_switch(&self, name: char) -> bool {
        let name = name.to_ascii_lowercase();
        self.switches.iter().any(|switch| switch.name == name)
    }
}

/// A typed, inert legacy Word `DATA` mail-merge source field.
///
/// [MS-DOC] §2.9.90 specifies `DATA datafile [headerfile]` as a field that
/// redirects mail-merge data and header files. This type exposes only those
/// stored operands, switches, cached result, and field state. It never opens,
/// reads, connects to, resolves, or modifies either source; it never selects a
/// record, performs a merge, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MailMergeDataField {
    pub(super) field: Field,
    pub(super) instruction: String,
    pub(super) data_source: String,
    pub(super) header_source: Option<String>,
    pub(super) switches: Vec<MergeFieldSwitch>,
    pub(super) cached_result: Option<String>,
}

impl MailMergeDataField {
    /// Return the paired field markers and their story-relative positions.
    pub fn field(&self) -> &Field {
        &self.field
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the stored mail-merge data-source identifier without opening it.
    pub fn data_source(&self) -> &str {
        &self.data_source
    }

    /// Return the optional stored mail-merge header-source identifier.
    ///
    /// This value is never opened or resolved.
    pub fn header_source(&self) -> Option<&str> {
        self.header_source.as_deref()
    }

    /// Return stored field switches in source order without interpreting them.
    pub fn switches(&self) -> &[MergeFieldSwitch] {
        &self.switches
    }

    /// Return the stored cached field result, if present.
    ///
    /// This value is never regenerated by a merge.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a producer marked the stored result stale.
    pub fn is_dirty(&self) -> bool {
        self.field.end_flags.results_dirty
    }

    /// Whether a producer locked this field against refresh.
    pub fn is_locked(&self) -> bool {
        self.field.end_flags.locked
    }
}

/// A typed, inert legacy Word `DOCVARIABLE` field.
///
/// [MS-DOC] §2.9.90 identifies its native field-type byte, and ECMA-376 Part
/// 1 §17.16.5.15 defines `DOCVARIABLE` with one document-variable name.
/// This type exposes the stored name, any preserved switches, and cached result
/// only. It never reads document variables, resolves a value, or refreshes a
/// field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DocumentVariableField {
    pub(super) field: Field,
    pub(super) instruction: String,
    pub(super) variable_name: String,
    pub(super) unknown_switches: Vec<MergeFieldSwitch>,
    pub(super) cached_result: Option<String>,
}

impl DocumentVariableField {
    /// Return the paired field markers and their story-relative positions.
    pub fn field(&self) -> &Field {
        &self.field
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the stored document-variable name without resolving it.
    pub fn variable_name(&self) -> &str {
        &self.variable_name
    }

    /// Return preserved switches in source order without interpreting them.
    ///
    /// `DOCVARIABLE` has no field-specific switches. These values remain
    /// inert source metadata and are never applied.
    pub fn unknown_switches(&self) -> &[MergeFieldSwitch] {
        &self.unknown_switches
    }

    /// Return the stored cached field result, if present.
    ///
    /// This value is never regenerated from a document variable.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a producer marked the stored result stale.
    pub fn is_dirty(&self) -> bool {
        self.field.end_flags.results_dirty
    }

    /// Whether a producer locked this field against refresh.
    pub fn is_locked(&self) -> bool {
        self.field.end_flags.locked
    }
}

/// A typed, inert legacy Word `DOCPROPERTY` field.
///
/// [MS-DOC] §2.9.90 identifies its native field-type byte, and ECMA-376 Part
/// 1 §17.16.5.14 defines `DOCPROPERTY` with one document-property name.
/// This type exposes the stored name, preserved switches, and cached result
/// only. It never reads document properties, resolves a value, or refreshes a
/// field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DocumentPropertyField {
    pub(super) field: Field,
    pub(super) instruction: String,
    pub(super) property_name: String,
    pub(super) switches: Vec<MergeFieldSwitch>,
    pub(super) cached_result: Option<String>,
}

impl DocumentPropertyField {
    /// Return the paired field markers and their story-relative positions.
    pub fn field(&self) -> &Field {
        &self.field
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the stored document-property name without resolving it.
    pub fn property_name(&self) -> &str {
        &self.property_name
    }

    /// Return preserved switches in source order without interpreting them.
    ///
    /// These values remain inert source metadata and are never applied.
    pub fn switches(&self) -> &[MergeFieldSwitch] {
        &self.switches
    }

    /// Return the stored cached field result, if present.
    ///
    /// This value is never regenerated from a document property.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a producer marked the stored result stale.
    pub fn is_dirty(&self) -> bool {
        self.field.end_flags.results_dirty
    }

    /// Whether a producer locked this field against refresh.
    pub fn is_locked(&self) -> bool {
        self.field.end_flags.locked
    }
}

/// A typed, inert legacy Word `INFO` field.
///
/// [MS-DOC] §2.9.90 identifies native `INFO` fields with type `0x0E`.
/// Word permits the `INFO` keyword to be omitted, and the native type
/// disambiguates that stored form from standalone document-information fields.
/// This type retains the stored property selector, optional replacement value,
/// switches, cached result, and field-marker state only. It never reads,
/// resolves, modifies, or writes document or template properties, or refreshes
/// a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct InfoField {
    pub(super) field: Field,
    pub(super) instruction: String,
    pub(super) information_type: String,
    pub(super) new_value: Option<String>,
    pub(super) switches: Vec<MergeFieldSwitch>,
    pub(super) cached_result: Option<String>,
}

impl InfoField {
    /// Return the paired field markers and their story-relative positions.
    pub fn field(&self) -> &Field {
        &self.field
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the stored document or template property selector.
    ///
    /// The selector is preserved as metadata and is never looked up.
    pub fn information_type(&self) -> &str {
        &self.information_type
    }

    /// Return the stored optional replacement value.
    ///
    /// This value is never applied to a document or template property.
    pub fn new_value(&self) -> Option<&str> {
        self.new_value.as_deref()
    }

    /// Return preserved switches in source order without interpreting them.
    ///
    /// These values remain inert source metadata and are never applied.
    pub fn switches(&self) -> &[MergeFieldSwitch] {
        &self.switches
    }

    /// Return the stored cached field result, if present.
    ///
    /// This value is never regenerated from a property.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a producer marked the stored result stale.
    pub fn is_dirty(&self) -> bool {
        self.field.end_flags.results_dirty
    }

    /// Whether a producer locked this field against refresh.
    pub fn is_locked(&self) -> bool {
        self.field.end_flags.locked
    }
}

/// The built-in Word document-information field category.
///
/// [MS-DOC] §2.9.90 assigns the native `flt` values 0x0F through 0x1C to
/// these fourteen Word field types. This enum preserves the stored category
/// only; it does not resolve document metadata or calculate dates, revisions,
/// or statistics.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum DocumentInformationFieldKind {
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
}

impl DocumentInformationFieldKind {
    /// The uppercase field keyword stored in a Word field instruction.
    pub const fn field_keyword(self) -> &'static str {
        match self {
            Self::Title => "TITLE",
            Self::Subject => "SUBJECT",
            Self::Author => "AUTHOR",
            Self::Keywords => "KEYWORDS",
            Self::Comments => "COMMENTS",
            Self::LastSavedBy => "LASTSAVEDBY",
            Self::CreateDate => "CREATEDATE",
            Self::SaveDate => "SAVEDATE",
            Self::PrintDate => "PRINTDATE",
            Self::RevisionNumber => "REVNUM",
            Self::EditTime => "EDITTIME",
            Self::NumberOfPages => "NUMPAGES",
            Self::NumberOfWords => "NUMWORDS",
            Self::NumberOfCharacters => "NUMCHARS",
        }
    }

    pub(super) fn from_field_type(field_type: FieldType) -> Option<Self> {
        match field_type {
            FieldType::Title => Some(Self::Title),
            FieldType::Subject => Some(Self::Subject),
            FieldType::Author => Some(Self::Author),
            FieldType::Keywords => Some(Self::Keywords),
            FieldType::Comments => Some(Self::Comments),
            FieldType::LastSavedBy => Some(Self::LastSavedBy),
            FieldType::CreateDate => Some(Self::CreateDate),
            FieldType::SaveDate => Some(Self::SaveDate),
            FieldType::PrintDate => Some(Self::PrintDate),
            FieldType::RevisionNumber => Some(Self::RevisionNumber),
            FieldType::EditTime => Some(Self::EditTime),
            FieldType::NumberOfPages => Some(Self::NumberOfPages),
            FieldType::NumberOfWords => Some(Self::NumberOfWords),
            FieldType::NumberOfCharacters => Some(Self::NumberOfCharacters),
            _ => None,
        }
    }

    pub(super) fn from_keyword(keyword: &str) -> Option<Self> {
        [
            Self::Title,
            Self::Subject,
            Self::Author,
            Self::Keywords,
            Self::Comments,
            Self::LastSavedBy,
            Self::CreateDate,
            Self::SaveDate,
            Self::PrintDate,
            Self::RevisionNumber,
            Self::EditTime,
            Self::NumberOfPages,
            Self::NumberOfWords,
            Self::NumberOfCharacters,
        ]
        .into_iter()
        .find(|kind| keyword.eq_ignore_ascii_case(kind.field_keyword()))
    }
}

/// A typed, inert legacy Word built-in document-information field.
///
/// This type exposes the stored native category, instruction, switches, and
/// cached result only. It never reads document properties, reads or modifies
/// host identity data, calculates dates, revisions, or statistics, resolves a
/// value, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DocumentInformationField {
    pub(super) field: Field,
    pub(super) instruction: String,
    pub(super) kind: DocumentInformationFieldKind,
    pub(super) switches: Vec<MergeFieldSwitch>,
    pub(super) cached_result: Option<String>,
}

impl DocumentInformationField {
    /// Return the paired field markers and their story-relative positions.
    pub fn field(&self) -> &Field {
        &self.field
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the recognized built-in document-information category.
    pub const fn kind(&self) -> DocumentInformationFieldKind {
        self.kind
    }

    /// Return preserved switches in source order without interpreting them.
    ///
    /// These values remain inert source metadata and are never applied.
    pub fn switches(&self) -> &[MergeFieldSwitch] {
        &self.switches
    }

    /// Return the stored cached field result, if present.
    ///
    /// This value is never regenerated from document metadata or a host user
    /// profile.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a producer marked the stored result stale.
    pub fn is_dirty(&self) -> bool {
        self.field.end_flags.results_dirty
    }

    /// Whether a producer locked this field against refresh.
    pub fn is_locked(&self) -> bool {
        self.field.end_flags.locked
    }
}

/// The built-in Word document-context and runtime field category.
///
/// [MS-DOC] §2.9.90 assigns the native `flt` values 0x1D through 0x21 to
/// `FILENAME`, `TEMPLATE`, `DATE`, `TIME`, and `PAGE`, and values 0x41, 0x42, and
/// 0x45 to `SECTION`, `SECTIONPAGES`, and `FILESIZE`. This enum preserves the stored
/// category only; it does not read a document path, attached template, host
/// filesystem state or file size, current clock, or page and section layout.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum DocumentContextFieldKind {
    FileName,
    Template,
    Date,
    Time,
    Page,
    FileSize,
    Section,
    SectionPages,
}

impl DocumentContextFieldKind {
    /// The uppercase field keyword stored in a Word field instruction.
    pub const fn field_keyword(self) -> &'static str {
        match self {
            Self::FileName => "FILENAME",
            Self::Template => "TEMPLATE",
            Self::Date => "DATE",
            Self::Time => "TIME",
            Self::Page => "PAGE",
            Self::FileSize => "FILESIZE",
            Self::Section => "SECTION",
            Self::SectionPages => "SECTIONPAGES",
        }
    }

    pub(super) fn from_field_type(field_type: FieldType) -> Option<Self> {
        match field_type {
            FieldType::FileName => Some(Self::FileName),
            FieldType::Template => Some(Self::Template),
            FieldType::Date => Some(Self::Date),
            FieldType::Time => Some(Self::Time),
            FieldType::Page => Some(Self::Page),
            FieldType::FileSize => Some(Self::FileSize),
            FieldType::Section => Some(Self::Section),
            FieldType::SectionPages => Some(Self::SectionPages),
            _ => None,
        }
    }

    pub(super) fn from_keyword(keyword: &str) -> Option<Self> {
        [
            Self::FileName,
            Self::Template,
            Self::Date,
            Self::Time,
            Self::Page,
            Self::FileSize,
            Self::Section,
            Self::SectionPages,
        ]
        .into_iter()
        .find(|kind| keyword.eq_ignore_ascii_case(kind.field_keyword()))
    }
}

/// A typed, inert legacy Word built-in document-context or runtime field.
///
/// This type exposes the stored native category, instruction, switches, and
/// cached result only. It never reads a document path, attached template, host
/// filesystem state or file size, current clock, or page and section layout,
/// resolves a value, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DocumentContextField {
    pub(super) field: Field,
    pub(super) instruction: String,
    pub(super) kind: DocumentContextFieldKind,
    pub(super) switches: Vec<MergeFieldSwitch>,
    pub(super) cached_result: Option<String>,
}

impl DocumentContextField {
    /// Return the paired field markers and their story-relative positions.
    pub fn field(&self) -> &Field {
        &self.field
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the recognized built-in document-context or runtime category.
    pub const fn kind(&self) -> DocumentContextFieldKind {
        self.kind
    }

    /// Return preserved switches in source order without interpreting them.
    ///
    /// These values remain inert source metadata and are never applied.
    pub fn switches(&self) -> &[MergeFieldSwitch] {
        &self.switches
    }

    /// Return the stored cached field result, if present.
    ///
    /// This value is never regenerated from a document path, attached
    /// template, host filesystem state or file size, current clock, or page
    /// and section layout.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a producer marked the stored result stale.
    pub fn is_dirty(&self) -> bool {
        self.field.end_flags.results_dirty
    }

    /// Whether a producer locked this field against refresh.
    pub fn is_locked(&self) -> bool {
        self.field.end_flags.locked
    }
}

/// The stored kind of a legacy Word DDE field.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DdeFieldKind {
    /// A `DDE` field, which can request automatic updates with `\\a`.
    Dde,
    /// A `DDEAUTO` field, which declares automatic updates.
    DdeAuto,
}

/// One stored DDE result-representation switch.
///
/// This value describes a requested representation only. It never causes a
/// source to be contacted, converted, embedded, or displayed.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DdeRepresentation {
    /// The `\\b` switch requests a bitmap representation.
    Bitmap,
    /// The `\\h` switch requests HTML-formatted text.
    Html,
    /// The `\\p` switch requests a picture representation.
    Picture,
    /// The `\\r` switch requests rich-text format.
    RichText,
    /// The `\\t` switch requests text-only format.
    Text,
    /// The `\\u` switch requests Unicode text.
    UnicodeText,
}

/// Typed, inert metadata for a legacy Word `DDE` or `DDEAUTO` field.
///
/// [MS-DOC] §2.9.90 identifies their native field-type bytes. Application,
/// source, item, representation, and storage switches remain stored metadata
/// only. This type never launches an application, initiates a DDE conversation,
/// opens a source, requests data, refreshes content, converts content, or
/// executes code.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DdeField {
    pub(super) field: Field,
    pub(super) instruction: String,
    pub(super) kind: DdeFieldKind,
    pub(super) application: String,
    pub(super) source: String,
    pub(super) item: Option<String>,
    pub(super) automatic_updates: bool,
    pub(super) representation: Option<DdeRepresentation>,
    pub(super) omit_graphic_data: bool,
    pub(super) unknown_switches: Vec<MergeFieldSwitch>,
    pub(super) cached_result: Option<String>,
}

impl DdeField {
    /// Return the paired field markers and their story-relative positions.
    pub fn field(&self) -> &Field {
        &self.field
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return whether this is a `DDE` or `DDEAUTO` field.
    pub fn kind(&self) -> DdeFieldKind {
        self.kind
    }

    /// Return the stored DDE application name without launching it.
    pub fn application(&self) -> &str {
        &self.application
    }

    /// Return the stored source identifier without opening or resolving it.
    pub fn source(&self) -> &str {
        &self.source
    }

    /// Return the optional stored source item, such as a cell range or bookmark.
    pub fn item(&self) -> Option<&str> {
        self.item.as_deref()
    }

    /// Whether the stored instruction requests automatic DDE updates.
    ///
    /// This is metadata only. The API never performs an update.
    pub fn requests_automatic_updates(&self) -> bool {
        self.automatic_updates
    }

    /// Return the requested stored result representation, if present.
    ///
    /// This is metadata only and never triggers source access or conversion.
    pub fn representation(&self) -> Option<DdeRepresentation> {
        self.representation
    }

    /// Whether the stored `\\d` switch omits graphic data from the document.
    ///
    /// This is stored metadata only. The API never reads the source to obtain
    /// omitted data.
    pub fn omits_graphic_data(&self) -> bool {
        self.omit_graphic_data
    }

    /// Return unrecognized stored field switches in source order.
    ///
    /// These values are preserved as inert metadata and never interpreted.
    pub fn unknown_switches(&self) -> &[MergeFieldSwitch] {
        &self.unknown_switches
    }

    /// Return the stored cached field result, if present.
    ///
    /// This value is never regenerated from a DDE source.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a producer marked the stored result stale.
    pub fn is_dirty(&self) -> bool {
        self.field.end_flags.results_dirty
    }

    /// Whether a producer locked this field against refresh.
    pub fn is_locked(&self) -> bool {
        self.field.end_flags.locked
    }
}

/// One stored result or storage switch for a Word `LINK` field.
///
/// These values describe a linked-object representation or whether graphic data
/// is stored. They never cause a source to be opened, contacted, converted, or
/// displayed.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LinkResultOption {
    /// The `\\b` switch requests a bitmap representation.
    Bitmap,
    /// The `\\d` switch omits graphic data from the document.
    OmitGraphicData,
    /// The `\\h` switch requests HTML-formatted text.
    Html,
    /// The `\\p` switch requests a picture representation.
    Picture,
    /// The `\\r` switch requests rich-text format.
    RichText,
    /// The `\\t` switch requests text-only format.
    Text,
    /// The `\\u` switch requests Unicode text.
    UnicodeText,
}

/// One integral `LINK` `\\f` formatting mode.
///
/// ECMA-376 marks modes 1 and 3 unsupported. Those values, and values outside
/// its defined set, are retained as metadata without applying any formatting.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LinkFormatting {
    /// 0: preserve formatting from the source file.
    Source,
    /// 2: match formatting in the destination document.
    Destination,
    /// 4: preserve source formatting for a SpreadsheetML workbook source.
    SpreadsheetSource,
    /// 5: match destination formatting for a SpreadsheetML workbook source.
    SpreadsheetDestination,
    /// An ECMA-376-unsupported or otherwise unrecognized integral mode.
    Unsupported(i64),
}

/// Typed, inert metadata for a legacy Word `LINK` field.
///
/// [MS-DOC] §2.9.90 identifies its native field-type byte. Application type,
/// source, item, and all result/formatting switches are retained as stored
/// field data. This type never activates an OLE server, launches an
/// application, opens a source, requests data, refreshes content, converts
/// content, or executes code.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct LinkField {
    pub(super) field: Field,
    pub(super) instruction: String,
    pub(super) application_type: String,
    pub(super) source: String,
    pub(super) item: Option<String>,
    pub(super) automatic_updates: bool,
    pub(super) result_options: Vec<LinkResultOption>,
    pub(super) formatting_modes: Vec<LinkFormatting>,
    pub(super) switches: Vec<MergeFieldSwitch>,
    pub(super) cached_result: Option<String>,
}

impl LinkField {
    /// Return the paired field markers and their story-relative positions.
    pub fn field(&self) -> &Field {
        &self.field
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the stored linked-object application type.
    ///
    /// Word commonly stores an OLE Programmatic Identifier here. It is never
    /// looked up or activated by this API.
    pub fn application_type(&self) -> &str {
        &self.application_type
    }

    /// Return the stored source identifier without opening or resolving it.
    pub fn source(&self) -> &str {
        &self.source
    }

    /// Return the optional stored source item, such as a cell range or bookmark.
    pub fn item(&self) -> Option<&str> {
        self.item.as_deref()
    }

    /// Whether the stored instruction requests automatic updates.
    ///
    /// This is metadata only. The API never performs an update.
    pub fn requests_automatic_updates(&self) -> bool {
        self.automatic_updates
    }

    /// Return recognized result and storage switches in stored source order.
    ///
    /// When several are present, `Self::effective_result_option` reflects
    /// Word's documented last-switch behavior. Neither method contacts the
    /// linked source.
    pub fn result_options(&self) -> &[LinkResultOption] {
        &self.result_options
    }

    /// Return the effective result or storage option under Word's documented
    /// last-switch behavior, if one was stored.
    pub fn effective_result_option(&self) -> Option<LinkResultOption> {
        self.result_options.last().copied()
    }

    /// Return integral `\\f` formatting modes in stored source order.
    ///
    /// These are metadata only; this API never formats linked content.
    pub fn formatting_modes(&self) -> &[LinkFormatting] {
        &self.formatting_modes
    }

    /// Return all stored field switches in source order.
    pub fn switches(&self) -> &[MergeFieldSwitch] {
        &self.switches
    }

    /// Return the stored cached field result, if present.
    ///
    /// This value is never regenerated from a linked source.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a producer marked the stored result stale.
    pub fn is_dirty(&self) -> bool {
        self.field.end_flags.results_dirty
    }

    /// Whether a producer locked this field against refresh.
    pub fn is_locked(&self) -> bool {
        self.field.end_flags.locked
    }
}

/// The kind of externally sourced legacy Word include field.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum IncludeFieldKind {
    /// An `INCLUDETEXT` or historical `INCLUDE` field that stores a document or XML source.
    Text,
    /// An `INCLUDEPICTURE` or historical `IMPORT` field that stores an image source.
    Picture,
}

/// One recognized stored option of a legacy Word external-include field.
///
/// These values are configuration metadata only. This API never opens,
/// resolves, imports, transforms, or evaluates the referenced source.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ExternalIncludeOption {
    /// A document or graphics converter name from the `\\c` switch.
    Converter(String),
    /// A source encoding from the `INCLUDETEXT \\e` switch.
    Encoding(String),
    /// A source MIME type from the `INCLUDETEXT \\m` switch.
    MimeType(String),
    /// An XML namespace mapping from the `INCLUDETEXT \\n` switch.
    NamespaceMapping(String),
    /// An XSLT location from the `INCLUDETEXT \\t` switch.
    Xslt(String),
    /// An XPath expression from the `INCLUDETEXT \\x` switch.
    XPath(String),
}

/// Typed, inert metadata for a legacy Word external-include field.
///
/// ECMA-376 Part 1 §§17.16.5.27–28 defines these stored source operands and
/// switches. [MS-DOC] §2.9.90 also identifies historical `INCLUDE` and
/// `IMPORT` native aliases for `INCLUDETEXT` and `INCLUDEPICTURE`,
/// respectively. Source identifiers, bookmarks, options, and cached results
/// are retained as stored field data. This type never opens, resolves,
/// imports, fetches, refreshes, transforms, converts, evaluates, or executes
/// source content.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ExternalIncludeField {
    pub(super) field: Field,
    pub(super) instruction: String,
    pub(super) kind: IncludeFieldKind,
    pub(super) source: String,
    pub(super) bookmark: Option<String>,
    pub(super) suppress_nested_field_updates: bool,
    pub(super) omit_picture_data: bool,
    pub(super) options: Vec<ExternalIncludeOption>,
    pub(super) unknown_switches: Vec<MergeFieldSwitch>,
    pub(super) cached_result: Option<String>,
}

impl ExternalIncludeField {
    /// Return the paired field markers and their story-relative positions.
    pub fn field(&self) -> &Field {
        &self.field
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return whether this stores a text or picture external-include field.
    pub fn kind(&self) -> IncludeFieldKind {
        self.kind
    }

    /// Return the stored source identifier without opening or resolving it.
    pub fn source(&self) -> &str {
        &self.source
    }

    /// Return the optional stored text-include bookmark selector.
    ///
    /// Picture-include fields do not define a bookmark operand, so they
    /// always return `None` here.
    pub fn bookmark(&self) -> Option<&str> {
        self.bookmark.as_deref()
    }

    /// Return the optional stored `\\c` converter name.
    ///
    /// The converter is never looked up or invoked.
    pub fn converter(&self) -> Option<&str> {
        self.options.iter().find_map(|option| match option {
            ExternalIncludeOption::Converter(value) => Some(value.as_str()),
            _ => None,
        })
    }

    /// Return recognized converter and XML options in stored source order.
    ///
    /// All options are inert metadata. This method never resolves a converter,
    /// opens a source, runs XSLT, or evaluates XPath.
    pub fn options(&self) -> &[ExternalIncludeOption] {
        &self.options
    }

    /// Whether the stored text-include `\\!` switch suppresses nested field updates.
    ///
    /// This is stored metadata only. The API never performs an update.
    pub fn suppresses_nested_field_updates(&self) -> bool {
        self.suppress_nested_field_updates
    }

    /// Whether the stored picture-include `\\d` switch omits picture data.
    ///
    /// This is stored metadata only. The API never reads the source to obtain
    /// omitted picture data.
    pub fn omits_picture_data(&self) -> bool {
        self.omit_picture_data
    }

    /// Return unrecognized stored field switches in source order.
    ///
    /// These values are preserved as inert metadata and never interpreted.
    pub fn unknown_switches(&self) -> &[MergeFieldSwitch] {
        &self.unknown_switches
    }

    /// Return the stored cached field result, if present.
    ///
    /// This value is never regenerated from an external source.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a producer marked the stored result stale.
    pub fn is_dirty(&self) -> bool {
        self.field.end_flags.results_dirty
    }

    /// Whether a producer locked this field against refresh.
    pub fn is_locked(&self) -> bool {
        self.field.end_flags.locked
    }
}

/// The stored kind of a legacy Word mail-merge counter field.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MailMergeCounterKind {
    /// A `MERGEREC` field, which stores a selected-record position.
    Record,
    /// A `MERGESEQ` field, which stores a merged-record sequence position.
    Sequence,
}

/// A typed, inert legacy Word `MERGEREC` or `MERGESEQ` field.
///
/// [MS-DOC] §2.9.90 identifies their field-type bytes, and ECMA-376 Part 1
/// §§17.16.5.36–37 define them as zero-argument fields. This type exposes the
/// persisted kind and cached result only. It never selects or counts records,
/// opens a data source, performs a merge, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MailMergeCounterField {
    pub(super) field: Field,
    pub(super) instruction: String,
    pub(super) kind: MailMergeCounterKind,
    pub(super) cached_result: Option<String>,
}

impl MailMergeCounterField {
    /// Return the paired field markers and their story-relative positions.
    pub fn field(&self) -> &Field {
        &self.field
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return whether this is a `MERGEREC` or `MERGESEQ` field.
    pub fn kind(&self) -> MailMergeCounterKind {
        self.kind
    }

    /// Return the stored cached field result, if present.
    ///
    /// This value is never regenerated by a merge.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a producer marked the stored result stale.
    pub fn is_dirty(&self) -> bool {
        self.field.end_flags.results_dirty
    }

    /// Whether a producer locked this field against refresh.
    pub fn is_locked(&self) -> bool {
        self.field.end_flags.locked
    }
}

/// A typed, inert legacy Word `NEXT` mail-merge control field.
///
/// [MS-DOC] §2.9.90 identifies its field-type byte, and ECMA-376 Part 1
/// §17.16.5.38 defines `NEXT` as a zero-argument instruction. This type
/// exposes persisted cached content and state only. It never advances a record,
/// opens a data source, performs a merge, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MailMergeNextField {
    pub(super) field: Field,
    pub(super) instruction: String,
    pub(super) cached_result: Option<String>,
}

impl MailMergeNextField {
    /// Return the paired field markers and their story-relative positions.
    pub fn field(&self) -> &Field {
        &self.field
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the stored cached field result, if present.
    ///
    /// This value is never regenerated by a merge.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a producer marked the stored result stale.
    pub fn is_dirty(&self) -> bool {
        self.field.end_flags.results_dirty
    }

    /// Whether a producer locked this field against refresh.
    pub fn is_locked(&self) -> bool {
        self.field.end_flags.locked
    }
}

/// The stored kind of a conditional mail-merge control field.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MailMergeConditionalControlKind {
    /// A `NEXTIF` field, which can advance a merge record when its comparison is true.
    NextIf,
    /// A `SKIPIF` field, which can omit a merge record when its comparison is true.
    SkipIf,
}

/// A typed, inert legacy Word `NEXTIF` or `SKIPIF` mail-merge control field.
///
/// [MS-DOC] §2.9.90 identifies the native field-type bytes, and ECMA-376 Part
/// 1 §§17.16.5.39 and 17.16.5.58 define these controls. This type exposes the
/// unparsed comparison and cached result only. It never parses or evaluates a
/// comparison, advances or skips a record, opens a data source, performs a
/// merge, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MailMergeConditionalControlField {
    pub(super) field: Field,
    pub(super) instruction: String,
    pub(super) kind: MailMergeConditionalControlKind,
    pub(super) comparison: String,
    pub(super) cached_result: Option<String>,
}

impl MailMergeConditionalControlField {
    /// Return the paired field markers and their story-relative positions.
    pub fn field(&self) -> &Field {
        &self.field
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return whether this is a `NEXTIF` or `SKIPIF` control.
    pub fn kind(&self) -> MailMergeConditionalControlKind {
        self.kind
    }

    /// Return the stored comparison without parsing or evaluating it.
    pub fn comparison(&self) -> &str {
        &self.comparison
    }

    /// Return the stored cached field result, if present.
    ///
    /// This value is never regenerated by a merge.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a producer marked the stored result stale.
    pub fn is_dirty(&self) -> bool {
        self.field.end_flags.results_dirty
    }

    /// Whether a producer locked this field against refresh.
    pub fn is_locked(&self) -> bool {
        self.field.end_flags.locked
    }
}

/// A typed, inert legacy Word `IF` field.
///
/// [MS-DOC] §2.9.90 identifies its native field-type byte, and ECMA-376 Part
/// 1 §17.16.5.26 defines `IF` using a comparison and two branches. This type
/// exposes the unparsed expression and cached result only. It never parses or
/// evaluates an expression, resolves field values, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct IfField {
    pub(super) field: Field,
    pub(super) instruction: String,
    pub(super) expression: String,
    pub(super) cached_result: Option<String>,
}

impl IfField {
    /// Return the paired field markers and their story-relative positions.
    pub fn field(&self) -> &Field {
        &self.field
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the stored expression without parsing or evaluating it.
    pub fn expression(&self) -> &str {
        &self.expression
    }

    /// Return the stored cached field result, if present.
    ///
    /// This value is never regenerated by field evaluation.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a producer marked the stored result stale.
    pub fn is_dirty(&self) -> bool {
        self.field.end_flags.results_dirty
    }

    /// Whether a producer locked this field against refresh.
    pub fn is_locked(&self) -> bool {
        self.field.end_flags.locked
    }
}

/// A typed, inert legacy Word `COMPARE` field.
///
/// [MS-DOC] §2.9.90 identifies its native field-type byte, and ECMA-376 Part
/// 1 §17.16.5.10 defines `COMPARE` using a comparison with a result of 1 or
/// 0. This type exposes the unparsed comparison and cached result only. It
/// never parses or evaluates a comparison, resolves nested field values, or
/// refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CompareField {
    pub(super) field: Field,
    pub(super) instruction: String,
    pub(super) comparison: String,
    pub(super) cached_result: Option<String>,
}

impl CompareField {
    /// Return the paired field markers and their story-relative positions.
    pub fn field(&self) -> &Field {
        &self.field
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the stored comparison without parsing or evaluating it.
    pub fn comparison(&self) -> &str {
        &self.comparison
    }

    /// Return the stored cached field result, if present.
    ///
    /// This value is never regenerated by field evaluation.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a producer marked the stored result stale.
    pub fn is_dirty(&self) -> bool {
        self.field.end_flags.results_dirty
    }

    /// Whether a producer locked this field against refresh.
    pub fn is_locked(&self) -> bool {
        self.field.end_flags.locked
    }
}

/// The stored kind of a prompt field.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PromptFieldKind {
    /// An `ASK` field that associates a response with a bookmark.
    Ask,
    /// A `FILLIN` field whose cached result represents a response.
    FillIn,
}

/// A typed, inert legacy Word `ASK` or `FILLIN` field.
///
/// [MS-DOC] §2.9.90 identifies their native field-type bytes, and ECMA-376
/// Part 1 §§17.16.5.3 and 17.16.5.19 define these fields. This type exposes
/// stored prompt and default-response metadata only. It never displays a
/// prompt, captures a response, creates or updates a bookmark, performs a
/// merge, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PromptField {
    pub(super) field: Field,
    pub(super) instruction: String,
    pub(super) kind: PromptFieldKind,
    pub(super) bookmark: Option<String>,
    pub(super) prompt: Option<String>,
    pub(super) default_response: Option<String>,
    pub(super) prompts_once_per_mail_merge: bool,
    pub(super) cached_result: Option<String>,
}

impl PromptField {
    /// Return the paired field markers and their story-relative positions.
    pub fn field(&self) -> &Field {
        &self.field
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return whether this is an `ASK` or `FILLIN` field.
    pub fn kind(&self) -> PromptFieldKind {
        self.kind
    }

    /// Return the bookmark name stored by an `ASK` field, if any.
    ///
    /// This is stored metadata only. It is never resolved, created, or updated.
    pub fn bookmark(&self) -> Option<&str> {
        self.bookmark.as_deref()
    }

    /// Return the stored prompt text, if any.
    ///
    /// This method returns metadata only and never displays a prompt.
    pub fn prompt(&self) -> Option<&str> {
        self.prompt.as_deref()
    }

    /// Return the stored default response, if one was supplied.
    ///
    /// `Some("")` represents an explicitly supplied blank default response. This
    /// is metadata only and is never selected, captured, or written into the
    /// document.
    pub fn default_response(&self) -> Option<&str> {
        self.default_response.as_deref()
    }

    /// Whether the stored `\o` switch requests one prompt for a mail merge.
    ///
    /// This request is never acted on: no merge is performed and no data source
    /// is opened.
    pub fn prompts_once_per_mail_merge(&self) -> bool {
        self.prompts_once_per_mail_merge
    }

    /// Return the stored cached field result, if present.
    ///
    /// This value is never regenerated by field evaluation.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a producer marked the stored result stale.
    pub fn is_dirty(&self) -> bool {
        self.field.end_flags.results_dirty
    }

    /// Whether a producer locked this field against refresh.
    pub fn is_locked(&self) -> bool {
        self.field.end_flags.locked
    }
}

/// The stored kind of a user-identity field.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum UserIdentityFieldKind {
    /// A `USERADDRESS` field.
    Address,
    /// A `USERINITIALS` field.
    Initials,
    /// A `USERNAME` field.
    Name,
}

/// A general-formatting request stored by a user-identity field.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum UserIdentityFormatting {
    /// The `\\* Caps` formatting request.
    Caps,
    /// The `\\* FirstCap` formatting request.
    FirstCap,
    /// The `\\* Lower` formatting request.
    Lower,
    /// The `\\* Upper` formatting request.
    Upper,
}

/// A typed, inert legacy Word `USERADDRESS`, `USERINITIALS`, or `USERNAME` field.
///
/// [MS-DOC] §2.9.90 identifies their native field-type bytes, and ECMA-376
/// Part 1 §§17.16.5.69–71 define these fields. This type exposes a stored
/// override, formatting request, and cached result only. It never reads or
/// modifies a host user's identity, applies formatting, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct UserIdentityField {
    pub(super) field: Field,
    pub(super) instruction: String,
    pub(super) kind: UserIdentityFieldKind,
    pub(super) override_value: Option<String>,
    pub(super) formatting: Option<UserIdentityFormatting>,
    pub(super) cached_result: Option<String>,
}

impl UserIdentityField {
    /// Return the paired field markers and their story-relative positions.
    pub fn field(&self) -> &Field {
        &self.field
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return whether this is an address, initials, or name field.
    pub fn kind(&self) -> UserIdentityFieldKind {
        self.kind
    }

    /// Return the optional stored value that overrides the host user context.
    ///
    /// `Some("")` represents an explicitly supplied blank override. This
    /// stored text is never written to, read from, or compared with a host
    /// identity.
    pub fn override_value(&self) -> Option<&str> {
        self.override_value.as_deref()
    }

    /// Return the stored general-formatting request, if any.
    ///
    /// This request is metadata only and is never applied to an identity value.
    pub fn formatting(&self) -> Option<UserIdentityFormatting> {
        self.formatting
    }

    /// Return the stored cached field result, if present.
    ///
    /// This value is never regenerated from a host identity.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a producer marked the stored result stale.
    pub fn is_dirty(&self) -> bool {
        self.field.end_flags.results_dirty
    }

    /// Whether a producer locked this field against refresh.
    pub fn is_locked(&self) -> bool {
        self.field.end_flags.locked
    }
}

/// One stored point-based `ADVANCE` placement operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AdvanceFieldOperation {
    /// The `\\d` switch moves subsequent text down.
    Down,
    /// The `\\l` switch moves subsequent text left.
    Left,
    /// The `\\r` switch moves subsequent text right.
    Right,
    /// The `\\u` switch moves subsequent text up.
    Up,
    /// The `\\x` switch specifies a horizontal position from the left edge
    /// of the column, frame, or text box.
    HorizontalPosition,
    /// The `\\y` switch specifies a vertical position relative to the page.
    VerticalPosition,
}

/// One stored `ADVANCE` point adjustment.
///
/// This is an instruction for a word processor's layout engine only. It is
/// never applied by this library.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct AdvanceFieldAdjustment {
    pub(super) operation: AdvanceFieldOperation,
    pub(super) points: i64,
}

impl AdvanceFieldAdjustment {
    /// Return the requested placement operation.
    pub fn operation(&self) -> AdvanceFieldOperation {
        self.operation
    }

    /// Return the stored signed integral number of points.
    pub fn points(&self) -> i64 {
        self.points
    }
}

/// A typed, inert legacy Word `ADVANCE` field.
///
/// [MS-DOC] §2.9.90 identifies its native field-type byte, and ECMA-376 Part
/// 1 §17.16.5.2 defines its six point-based placement switches. This type
/// exposes stored adjustments and cached content only. It never moves text,
/// changes layout, reflows content, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AdvanceField {
    pub(super) field: Field,
    pub(super) instruction: String,
    pub(super) adjustments: Vec<AdvanceFieldAdjustment>,
    pub(super) cached_result: Option<String>,
}

impl AdvanceField {
    /// Return the paired field markers and their story-relative positions.
    pub fn field(&self) -> &Field {
        &self.field
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the stored placement adjustments in source order.
    ///
    /// Repeated operations are preserved; this library does not resolve or
    /// apply them.
    pub fn adjustments(&self) -> &[AdvanceFieldAdjustment] {
        &self.adjustments
    }

    /// Return the stored cached field result, if present.
    ///
    /// `ADVANCE` has no regenerated value here; any returned text is stored
    /// source content only.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a producer marked the stored result stale.
    pub fn is_dirty(&self) -> bool {
        self.field.end_flags.results_dirty
    }

    /// Whether a producer locked this field against refresh.
    pub fn is_locked(&self) -> bool {
        self.field.end_flags.locked
    }
}

/// The stored kind of a mail-merge recipient layout field.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MailMergeRecipientFieldKind {
    /// An `ADDRESSBLOCK` field.
    AddressBlock,
    /// A `GREETINGLINE` field.
    GreetingLine,
}

/// How an `ADDRESSBLOCK` field requests country/region text.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AddressBlockCountryInclusion {
    /// `\\c 0` omits the country/region.
    Omit,
    /// `\\c 1` includes the country/region regardless of exclusions.
    Always,
    /// `\\c 2` includes the country/region unless it matches an excluded
    /// country/region.
    UnlessExcluded,
}

/// A typed, inert legacy Word `ADDRESSBLOCK` or `GREETINGLINE` field.
///
/// [MS-DOC] §2.9.90 identifies their native field-type bytes, and ECMA-376
/// Part 1 §§17.16.5.1 and 17.16.5.24 define their stored mail-merge recipient
/// layout metadata. This type never opens a data source, selects a record,
/// performs a merge, expands placeholders, generates text, or refreshes a
/// field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MailMergeRecipientField {
    pub(super) field: Field,
    pub(super) instruction: String,
    pub(super) kind: MailMergeRecipientFieldKind,
    pub(super) country_inclusion: Option<AddressBlockCountryInclusion>,
    pub(super) formats_using_recipient_country: bool,
    pub(super) excluded_countries: Vec<String>,
    pub(super) format_template: Option<String>,
    pub(super) language: Option<String>,
    pub(super) greeting_fallback_text: Option<String>,
    pub(super) unknown_switches: Vec<MergeFieldSwitch>,
    pub(super) cached_result: Option<String>,
}

impl MailMergeRecipientField {
    /// Return the paired field markers and their story-relative positions.
    pub fn field(&self) -> &Field {
        &self.field
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return whether this is an `ADDRESSBLOCK` or `GREETINGLINE` field.
    pub fn kind(&self) -> MailMergeRecipientFieldKind {
        self.kind
    }

    /// Return how an `ADDRESSBLOCK` requests country/region text.
    ///
    /// This is `None` when the instruction has no `\\c` switch or when the
    /// field is a `GREETINGLINE`. The stored request is never used to render
    /// an address.
    pub fn country_inclusion(&self) -> Option<AddressBlockCountryInclusion> {
        self.country_inclusion
    }

    /// Whether an `ADDRESSBLOCK` stores the `\\d` request to use the
    /// recipient country's address format.
    ///
    /// This request is metadata only and never causes a record or country
    /// format to be resolved.
    pub fn formats_using_recipient_country(&self) -> bool {
        self.formats_using_recipient_country
    }

    /// Return country/region names excluded by `ADDRESSBLOCK` `\\e` switches.
    ///
    /// ECMA-376 permits repeated `\\e` switches; values are retained in source
    /// order. They are never matched against a recipient record.
    pub fn excluded_countries(&self) -> &[String] {
        &self.excluded_countries
    }

    /// Return the stored `\\f` layout template, if any.
    ///
    /// For `ADDRESSBLOCK`, this is the standard address/name placeholder
    /// template. For `GREETINGLINE`, this accepts Word's documented
    /// compatibility form. Placeholder text remains opaque metadata and is
    /// never expanded.
    pub fn format_template(&self) -> Option<&str> {
        self.format_template.as_deref()
    }

    /// Return the stored `\\l` language identifier, if any.
    ///
    /// The identifier is not used to choose locale-specific formatting.
    pub fn language(&self) -> Option<&str> {
        self.language.as_deref()
    }

    /// Return the stored `GREETINGLINE` fallback text, if any.
    ///
    /// ECMA-376 names `\\c` as this switch; Word-compatible fields can use
    /// `\\e`. Both forms are accepted as stored metadata, but neither is ever
    /// selected or displayed by this API.
    pub fn greeting_fallback_text(&self) -> Option<&str> {
        self.greeting_fallback_text.as_deref()
    }

    /// Return switches not specific to the recognized recipient-field kind.
    ///
    /// This includes formatting or producer-specific switches, retained in
    /// source order as inert metadata.
    pub fn unknown_switches(&self) -> &[MergeFieldSwitch] {
        &self.unknown_switches
    }

    /// Return the stored cached field result, if present.
    ///
    /// This value is never regenerated by a merge.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a producer marked the stored result stale.
    pub fn is_dirty(&self) -> bool {
        self.field.end_flags.results_dirty
    }

    /// Whether a producer locked this field against refresh.
    pub fn is_locked(&self) -> bool {
        self.field.end_flags.locked
    }
}

pub(super) fn corrupted(message: impl Into<String>) -> DocError {
    DocError::Corrupted(message.into())
}
