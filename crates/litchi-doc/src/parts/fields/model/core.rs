use super::corrupted;
use crate::package::Result;
use crate::parts::fib::FileInformationBlock;

pub(in crate::parts::fields) const FLD_SIZE: usize = 2;
pub(in crate::parts::fields) const CP_SIZE: usize = 4;
pub(in crate::parts::fields) const MAX_PLCFLD_BYTES: usize = 64 * 1024 * 1024;
pub(in crate::parts::fields) const MAX_FIELD_MARKERS: usize = 1_000_000;
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

    pub(in crate::parts::fields) const fn pointer_index(self) -> usize {
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

    pub(in crate::parts::fields) fn character_count(self, fib: &FileInformationBlock) -> u32 {
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
    #[must_use]
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
    #[must_use]
    pub const fn is_specified(self) -> bool {
        !matches!(self, Self::Unknown(_))
    }

    /// Map a stored field-instruction keyword to its native `flt` type.
    ///
    /// Keywords are ASCII case-insensitive. The five text-only field kinds
    /// excluded from `Plcfld` (`TC`, `TA`, `XE`, `RD`, and `PRIVATE`) return
    /// `None`, as do unrecognized keywords.
    #[must_use]
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
    #[must_use]
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

    #[must_use]
    pub fn boundary(&self) -> FieldBoundary {
        match self.value {
            FieldMarkerValue::Begin(_) => FieldBoundary::Begin,
            FieldMarkerValue::Separator { .. } => FieldBoundary::Separator,
            FieldMarkerValue::End(_) => FieldBoundary::End,
        }
    }

    #[must_use]
    pub fn is_begin(&self) -> bool {
        self.boundary() == FieldBoundary::Begin
    }

    #[must_use]
    pub fn is_separator(&self) -> bool {
        self.boundary() == FieldBoundary::Separator
    }

    #[must_use]
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
    #[must_use]
    pub fn code_range(&self) -> (u32, u32) {
        (self.start_cp + 1, self.separator_cp.unwrap_or(self.end_cp))
    }

    #[must_use]
    pub fn result_range(&self) -> Option<(u32, u32)> {
        self.separator_cp
            .map(|separator| (separator + 1, self.end_cp))
    }

    #[must_use]
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
    pub(in crate::parts::fields) start_cp: u32,
    pub(in crate::parts::fields) instruction_start_byte: usize,
    pub(in crate::parts::fields) separator_cp: Option<u32>,
    pub(in crate::parts::fields) separator_start_byte: Option<usize>,
    pub(in crate::parts::fields) result_start_byte: Option<usize>,
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
