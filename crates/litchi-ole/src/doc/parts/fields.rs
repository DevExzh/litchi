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
    field: Field,
    instruction: String,
    macro_name: String,
    display_text: String,
    cached_result: Option<String>,
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
    field: Field,
    instruction: String,
    target: String,
    button_text: String,
    cached_result: Option<String>,
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

/// One stored switch in a legacy Word `MERGEFIELD` instruction.
///
/// The name excludes its leading backslash and is normalized to ASCII
/// lowercase. The argument is stored text only and is never resolved against a
/// mail-merge data source.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MergeFieldSwitch {
    name: char,
    argument: Option<String>,
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
    field: Field,
    instruction: String,
    field_name: String,
    switches: Vec<MergeFieldSwitch>,
    cached_result: Option<String>,
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

/// A typed, inert legacy Word `DOCVARIABLE` field.
///
/// [MS-DOC] §2.9.90 identifies its native field-type byte, and ECMA-376 Part
/// 1 §17.16.5.15 defines `DOCVARIABLE` with one document-variable name.
/// This type exposes the stored name, any preserved switches, and cached result
/// only. It never reads document variables, resolves a value, or refreshes a
/// field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DocumentVariableField {
    field: Field,
    instruction: String,
    variable_name: String,
    unknown_switches: Vec<MergeFieldSwitch>,
    cached_result: Option<String>,
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
    field: Field,
    instruction: String,
    kind: DdeFieldKind,
    application: String,
    source: String,
    item: Option<String>,
    automatic_updates: bool,
    representation: Option<DdeRepresentation>,
    omit_graphic_data: bool,
    unknown_switches: Vec<MergeFieldSwitch>,
    cached_result: Option<String>,
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
    field: Field,
    instruction: String,
    application_type: String,
    source: String,
    item: Option<String>,
    automatic_updates: bool,
    result_options: Vec<LinkResultOption>,
    formatting_modes: Vec<LinkFormatting>,
    switches: Vec<MergeFieldSwitch>,
    cached_result: Option<String>,
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
    field: Field,
    instruction: String,
    kind: MailMergeCounterKind,
    cached_result: Option<String>,
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
    field: Field,
    instruction: String,
    cached_result: Option<String>,
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
    field: Field,
    instruction: String,
    kind: MailMergeConditionalControlKind,
    comparison: String,
    cached_result: Option<String>,
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
    field: Field,
    instruction: String,
    expression: String,
    cached_result: Option<String>,
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
    field: Field,
    instruction: String,
    comparison: String,
    cached_result: Option<String>,
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
    field: Field,
    instruction: String,
    kind: PromptFieldKind,
    bookmark: Option<String>,
    prompt: Option<String>,
    default_response: Option<String>,
    prompts_once_per_mail_merge: bool,
    cached_result: Option<String>,
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
    field: Field,
    instruction: String,
    kind: UserIdentityFieldKind,
    override_value: Option<String>,
    formatting: Option<UserIdentityFormatting>,
    cached_result: Option<String>,
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
    operation: AdvanceFieldOperation,
    points: i64,
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
    field: Field,
    instruction: String,
    adjustments: Vec<AdvanceFieldAdjustment>,
    cached_result: Option<String>,
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
    field: Field,
    instruction: String,
    kind: MailMergeRecipientFieldKind,
    country_inclusion: Option<AddressBlockCountryInclusion>,
    formats_using_recipient_country: bool,
    excluded_countries: Vec<String>,
    format_template: Option<String>,
    language: Option<String>,
    greeting_fallback_text: Option<String>,
    unknown_switches: Vec<MergeFieldSwitch>,
    cached_result: Option<String>,
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

impl FieldText {
    pub(crate) fn from_field<F>(field: &Field, mut text_at_range: F) -> Result<Self>
    where
        F: FnMut(u32, u32) -> Result<String>,
    {
        let instruction_start = field
            .start_cp
            .checked_add(1)
            .ok_or_else(|| corrupted("field instruction start overflows"))?;
        let instruction_end = field.separator_cp.unwrap_or(field.end_cp);
        if instruction_start > instruction_end {
            return Err(corrupted(
                "field instruction range has its start after its end",
            ));
        }
        let instruction = text_at_range(instruction_start, instruction_end)?;
        let result = match field.separator_cp {
            Some(separator) => {
                let start = separator
                    .checked_add(1)
                    .ok_or_else(|| corrupted("field result start overflows"))?;
                if start > field.end_cp {
                    return Err(corrupted("field result range has its start after its end"));
                }
                Some(text_at_range(start, field.end_cp)?)
            },
            None => None,
        };

        Ok(Self {
            field: field.clone(),
            instruction,
            result,
        })
    }

    /// Return inert typed metadata when this is a well-formed `MACROBUTTON`
    /// field.
    ///
    /// The macro or command name and button text are parsed only from stored
    /// field text. Neither is resolved, loaded, invoked, or executed.
    /// Malformed instructions remain available through this generic type and
    /// return `None` here.
    pub fn macro_button(&self) -> Option<MacroButtonField> {
        if self.field.field_type != FieldType::MacroButton {
            return None;
        }
        let (macro_name, display_text) = parse_macro_button_parts(&self.instruction)?;
        Some(MacroButtonField {
            field: self.field.clone(),
            instruction: self.instruction.clone(),
            macro_name,
            display_text,
            cached_result: self.result.clone(),
        })
    }

    /// Return inert typed metadata when this is a well-formed `GOTOBUTTON`
    /// field.
    ///
    /// The destination and button text are parsed only from stored field text.
    /// Neither is resolved, navigated to, or activated. Malformed instructions
    /// remain available through this generic type and return `None` here.
    pub fn go_to_button(&self) -> Option<GoToButtonField> {
        if self.field.field_type != FieldType::GoToButton {
            return None;
        }
        let (target, button_text) = parse_go_to_button_parts(&self.instruction)?;
        Some(GoToButtonField {
            field: self.field.clone(),
            instruction: self.instruction.clone(),
            target,
            button_text,
            cached_result: self.result.clone(),
        })
    }

    /// Return inert typed metadata when this is a well-formed `MERGEFIELD` field.
    ///
    /// The stored data-column name, switches, and cached result are never
    /// resolved against a data source, merged into the document, or refreshed.
    /// Malformed instructions remain available through this generic type and
    /// return `None` here.
    pub fn merge_field(&self) -> Option<MergeField> {
        if self.field.field_type != FieldType::MergeField {
            return None;
        }
        let (field_name, switches) = parse_merge_field_parts(&self.instruction)?;
        Some(MergeField {
            field: self.field.clone(),
            instruction: self.instruction.clone(),
            field_name,
            switches,
            cached_result: self.result.clone(),
        })
    }

    /// Return inert typed metadata when this is a well-formed `DOCVARIABLE`
    /// field.
    ///
    /// The stored variable name, switches, and cached result are never resolved
    /// against document variables or refreshed. Malformed instructions remain
    /// available through this generic type and return `None` here.
    pub fn document_variable(&self) -> Option<DocumentVariableField> {
        if self.field.field_type != FieldType::DocumentVariable {
            return None;
        }
        let (variable_name, unknown_switches) =
            parse_document_variable_field_parts(&self.instruction)?;
        Some(DocumentVariableField {
            field: self.field.clone(),
            instruction: self.instruction.clone(),
            variable_name,
            unknown_switches,
            cached_result: self.result.clone(),
        })
    }

    /// Return inert typed metadata when this is a well-formed `DDE` or
    /// `DDEAUTO` field.
    ///
    /// Stored application, source, item, and switch data are never used to
    /// launch an application, initiate a DDE conversation, open a source,
    /// request data, refresh a field, convert content, or execute code.
    /// Malformed instructions remain available through this generic type and
    /// return `None` here.
    pub fn dde_link(&self) -> Option<DdeField> {
        let parts = parse_dde_field_parts(&self.instruction)?;
        if !matches!(
            (self.field.field_type, parts.kind),
            (FieldType::Dde, DdeFieldKind::Dde) | (FieldType::DdeAuto, DdeFieldKind::DdeAuto)
        ) {
            return None;
        }
        Some(DdeField {
            field: self.field.clone(),
            instruction: self.instruction.clone(),
            kind: parts.kind,
            application: parts.application,
            source: parts.source,
            item: parts.item,
            automatic_updates: parts.automatic_updates,
            representation: parts.representation,
            omit_graphic_data: parts.omit_graphic_data,
            unknown_switches: parts.unknown_switches,
            cached_result: self.result.clone(),
        })
    }

    /// Return inert typed metadata when this is a well-formed `LINK` field.
    ///
    /// Stored application, source, item, and switch data are never used to
    /// activate an OLE server, launch an application, open a source, request
    /// data, refresh a field, convert content, or execute code. Malformed
    /// instructions remain available through this generic type and return
    /// `None` here.
    pub fn link_field(&self) -> Option<LinkField> {
        if self.field.field_type != FieldType::Link {
            return None;
        }
        let parts = parse_link_field_parts(&self.instruction)?;
        Some(LinkField {
            field: self.field.clone(),
            instruction: self.instruction.clone(),
            application_type: parts.application_type,
            source: parts.source,
            item: parts.item,
            automatic_updates: parts.automatic_updates,
            result_options: parts.result_options,
            formatting_modes: parts.formatting_modes,
            switches: parts.switches,
            cached_result: self.result.clone(),
        })
    }

    /// Return inert typed metadata when this is a well-formed mail-merge counter.
    ///
    /// The stored kind and cached result are never used to select or count
    /// records, open a data source, perform a merge, or refresh a field
    /// result. Malformed instructions remain available through this generic
    /// type and return `None` here.
    pub fn mail_merge_counter(&self) -> Option<MailMergeCounterField> {
        let kind = parse_mail_merge_counter_kind(&self.instruction)?;
        if !matches!(
            (self.field.field_type, kind),
            (FieldType::MergeRecord, MailMergeCounterKind::Record)
                | (FieldType::MergeSequence, MailMergeCounterKind::Sequence)
        ) {
            return None;
        }
        Some(MailMergeCounterField {
            field: self.field.clone(),
            instruction: self.instruction.clone(),
            kind,
            cached_result: self.result.clone(),
        })
    }

    /// Return inert typed metadata when this is a well-formed `NEXT` field.
    ///
    /// Cached text and field state are never used to advance a record, open a
    /// data source, perform a merge, or refresh a field result. Malformed
    /// instructions remain available through this generic type and return
    /// `None` here.
    pub fn mail_merge_next(&self) -> Option<MailMergeNextField> {
        if self.field.field_type != FieldType::Next
            || !is_mail_merge_next_instruction(&self.instruction)
        {
            return None;
        }
        Some(MailMergeNextField {
            field: self.field.clone(),
            instruction: self.instruction.clone(),
            cached_result: self.result.clone(),
        })
    }

    /// Return inert metadata when this is a `NEXTIF` or `SKIPIF` field with a comparison.
    ///
    /// The stored comparison and cached result are never parsed or evaluated.
    /// This method never changes record selection, opens a data source,
    /// performs a merge, or refreshes a field result. Instructions without a
    /// comparison remain available through this generic type and return `None`
    /// here.
    pub fn mail_merge_conditional_control(&self) -> Option<MailMergeConditionalControlField> {
        let (kind, comparison) = parse_mail_merge_conditional_control_parts(&self.instruction)?;
        if !matches!(
            (self.field.field_type, kind),
            (FieldType::NextIf, MailMergeConditionalControlKind::NextIf)
                | (FieldType::SkipIf, MailMergeConditionalControlKind::SkipIf)
        ) {
            return None;
        }
        Some(MailMergeConditionalControlField {
            field: self.field.clone(),
            instruction: self.instruction.clone(),
            kind,
            comparison,
            cached_result: self.result.clone(),
        })
    }

    /// Return inert metadata when this is an `IF` field with an expression.
    ///
    /// The stored expression and cached result are never parsed or evaluated.
    /// This method never resolves field values or refreshes a field result.
    /// Instructions without an expression remain available through this generic
    /// type and return `None` here.
    pub fn if_field(&self) -> Option<IfField> {
        if self.field.field_type != FieldType::If {
            return None;
        }
        let expression = parse_if_field_expression(&self.instruction)?;
        Some(IfField {
            field: self.field.clone(),
            instruction: self.instruction.clone(),
            expression,
            cached_result: self.result.clone(),
        })
    }

    /// Return inert metadata when this is a `COMPARE` field with a comparison.
    ///
    /// The stored comparison and cached result are never parsed or evaluated.
    /// This method never resolves nested field values or refreshes a field
    /// result. Instructions without a comparison remain available through this
    /// generic type and return `None` here.
    pub fn compare_field(&self) -> Option<CompareField> {
        if self.field.field_type != FieldType::Compare {
            return None;
        }
        let comparison = parse_compare_field_comparison(&self.instruction)?;
        Some(CompareField {
            field: self.field.clone(),
            instruction: self.instruction.clone(),
            comparison,
            cached_result: self.result.clone(),
        })
    }

    /// Return inert typed metadata when this is a well-formed `ASK` or `FILLIN` field.
    ///
    /// Stored prompt, bookmark, default-response, and cached-result data are
    /// never used to display a prompt, capture a response, create or update a
    /// bookmark, perform a merge, or refresh a field. Malformed instructions
    /// remain available through this generic type and return `None` here.
    pub fn prompt_field(&self) -> Option<PromptField> {
        let (kind, bookmark, prompt, default_response, prompts_once_per_mail_merge) =
            parse_prompt_field_parts(&self.instruction)?;
        if !matches!(
            (self.field.field_type, kind),
            (FieldType::Ask, PromptFieldKind::Ask) | (FieldType::FillIn, PromptFieldKind::FillIn)
        ) {
            return None;
        }
        Some(PromptField {
            field: self.field.clone(),
            instruction: self.instruction.clone(),
            kind,
            bookmark,
            prompt,
            default_response,
            prompts_once_per_mail_merge,
            cached_result: self.result.clone(),
        })
    }

    /// Return inert typed metadata when this is a well-formed user-identity
    /// field.
    ///
    /// Stored override, formatting, and cached-result data are never used to
    /// read or modify a host user's identity, apply formatting, or refresh a
    /// field. Malformed instructions remain available through this generic type
    /// and return `None` here.
    pub fn user_identity_field(&self) -> Option<UserIdentityField> {
        let (kind, override_value, formatting) =
            parse_user_identity_field_parts(&self.instruction)?;
        if !matches!(
            (self.field.field_type, kind),
            (FieldType::UserAddress, UserIdentityFieldKind::Address)
                | (FieldType::UserInitials, UserIdentityFieldKind::Initials)
                | (FieldType::UserName, UserIdentityFieldKind::Name)
        ) {
            return None;
        }
        Some(UserIdentityField {
            field: self.field.clone(),
            instruction: self.instruction.clone(),
            kind,
            override_value,
            formatting,
            cached_result: self.result.clone(),
        })
    }

    /// Return inert typed metadata when this is a well-formed `ADVANCE` field.
    ///
    /// Stored point adjustments and cached-result data are never used to move
    /// text, change layout, reflow content, or refresh a field. Malformed
    /// instructions remain available through this generic type and return
    /// `None` here.
    pub fn advance_field(&self) -> Option<AdvanceField> {
        if self.field.field_type != FieldType::Advance {
            return None;
        }
        let adjustments = parse_advance_field_adjustments(&self.instruction)?;
        Some(AdvanceField {
            field: self.field.clone(),
            instruction: self.instruction.clone(),
            adjustments,
            cached_result: self.result.clone(),
        })
    }

    /// Return inert metadata when this is a well-formed `ADDRESSBLOCK` or
    /// `GREETINGLINE` field.
    ///
    /// Stored layout, locale, country, fallback, and cached-result data are
    /// never used to open a data source, select a record, perform a merge,
    /// expand placeholders, generate text, or refresh a field. Malformed
    /// instructions remain generic fields and return `None` here.
    pub fn mail_merge_recipient_field(&self) -> Option<MailMergeRecipientField> {
        let (
            kind,
            country_inclusion,
            formats_using_recipient_country,
            excluded_countries,
            format_template,
            language,
            greeting_fallback_text,
            unknown_switches,
        ) = parse_mail_merge_recipient_field_parts(&self.instruction)?;
        if !matches!(
            (self.field.field_type, kind),
            (
                FieldType::AddressBlock,
                MailMergeRecipientFieldKind::AddressBlock
            ) | (
                FieldType::GreetingLine,
                MailMergeRecipientFieldKind::GreetingLine
            )
        ) {
            return None;
        }
        Some(MailMergeRecipientField {
            field: self.field.clone(),
            instruction: self.instruction.clone(),
            kind,
            country_inclusion,
            formats_using_recipient_country,
            excluded_countries,
            format_template,
            language,
            greeting_fallback_text,
            unknown_switches,
            cached_result: self.result.clone(),
        })
    }
}

const MAX_MACRO_BUTTON_INSTRUCTION_BYTES: usize = 64 * 1024;
const MAX_GO_TO_BUTTON_INSTRUCTION_BYTES: usize = 64 * 1024;
const MAX_MERGE_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
const MAX_MERGE_FIELD_SWITCHES: usize = 64;
const MAX_DOCUMENT_VARIABLE_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
const MAX_DOCUMENT_VARIABLE_FIELD_SWITCHES: usize = 64;
const MAX_DDE_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
const MAX_DDE_FIELD_SWITCHES: usize = 64;
const MAX_LINK_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
const MAX_LINK_FIELD_SWITCHES: usize = 64;
const MAX_MAIL_MERGE_COUNTER_INSTRUCTION_BYTES: usize = 64 * 1024;
const MAX_MAIL_MERGE_NEXT_INSTRUCTION_BYTES: usize = 64 * 1024;
const MAX_MAIL_MERGE_CONDITIONAL_CONTROL_INSTRUCTION_BYTES: usize = 64 * 1024;
const MAX_IF_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
const MAX_COMPARE_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
const MAX_PROMPT_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
const MAX_USER_IDENTITY_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
const MAX_ADVANCE_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
const MAX_ADVANCE_FIELD_ADJUSTMENTS: usize = 64;
const MAX_MAIL_MERGE_RECIPIENT_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
const MAX_MAIL_MERGE_RECIPIENT_FIELD_SWITCHES: usize = 64;

struct DdeFieldParts {
    kind: DdeFieldKind,
    application: String,
    source: String,
    item: Option<String>,
    automatic_updates: bool,
    representation: Option<DdeRepresentation>,
    omit_graphic_data: bool,
    unknown_switches: Vec<MergeFieldSwitch>,
}

struct LinkFieldParts {
    application_type: String,
    source: String,
    item: Option<String>,
    automatic_updates: bool,
    result_options: Vec<LinkResultOption>,
    formatting_modes: Vec<LinkFormatting>,
    switches: Vec<MergeFieldSwitch>,
}

fn parse_macro_button_parts(instruction: &str) -> Option<(String, String)> {
    if instruction.len() > MAX_MACRO_BUTTON_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    if !keyword.eq_ignore_ascii_case("MACROBUTTON") {
        return None;
    }

    let macro_name = next_field_argument(instruction, &mut position).ok()??;
    if macro_name.is_empty() {
        return None;
    }
    let display_text = next_field_argument(instruction, &mut position).ok()??;
    if display_text.is_empty() {
        return None;
    }
    if next_field_argument(instruction, &mut position)
        .ok()?
        .is_some()
    {
        return None;
    }

    Some((macro_name, display_text))
}

fn parse_go_to_button_parts(instruction: &str) -> Option<(String, String)> {
    if instruction.len() > MAX_GO_TO_BUTTON_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    if !keyword.eq_ignore_ascii_case("GOTOBUTTON") {
        return None;
    }

    let target = next_field_argument(instruction, &mut position).ok()??;
    if target.is_empty() {
        return None;
    }
    let button_text = next_field_argument(instruction, &mut position).ok()??;
    if button_text.is_empty() {
        return None;
    }
    if next_field_argument(instruction, &mut position)
        .ok()?
        .is_some()
    {
        return None;
    }

    Some((target, button_text))
}

fn parse_merge_field_parts(instruction: &str) -> Option<(String, Vec<MergeFieldSwitch>)> {
    if instruction.len() > MAX_MERGE_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    if !keyword.eq_ignore_ascii_case("MERGEFIELD") {
        return None;
    }

    let field_name = next_field_argument(instruction, &mut position).ok()??;
    if field_name.is_empty() {
        return None;
    }

    let mut switches = Vec::new();
    loop {
        skip_field_whitespace(instruction, &mut position);
        let Some(introducer) = next_field_character(instruction, &mut position) else {
            break;
        };
        if introducer != '\\' || switches.len() >= MAX_MERGE_FIELD_SWITCHES {
            return None;
        }

        let name = next_field_character(instruction, &mut position)?;
        if name == '\\' || name.is_whitespace() {
            return None;
        }

        skip_field_whitespace(instruction, &mut position);
        let argument = match peek_field_character(instruction, position) {
            None | Some('\\') => None,
            Some(_) => next_field_argument(instruction, &mut position).ok()?,
        };
        switches.push(MergeFieldSwitch {
            name: name.to_ascii_lowercase(),
            argument,
        });
    }

    Some((field_name, switches))
}

fn parse_document_variable_field_parts(
    instruction: &str,
) -> Option<(String, Vec<MergeFieldSwitch>)> {
    if instruction.len() > MAX_DOCUMENT_VARIABLE_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    if !keyword.eq_ignore_ascii_case("DOCVARIABLE") {
        return None;
    }

    let variable_name = next_field_argument(instruction, &mut position).ok()??;
    if variable_name.is_empty() {
        return None;
    }

    let mut unknown_switches = Vec::new();
    loop {
        skip_field_whitespace(instruction, &mut position);
        let Some(introducer) = next_field_character(instruction, &mut position) else {
            break;
        };
        if introducer != '\\' || unknown_switches.len() >= MAX_DOCUMENT_VARIABLE_FIELD_SWITCHES {
            return None;
        }

        let name = next_field_character(instruction, &mut position)?;
        if name == '\\' || name.is_whitespace() {
            return None;
        }

        skip_field_whitespace(instruction, &mut position);
        let argument = match peek_field_character(instruction, position) {
            None | Some('\\') => None,
            Some(_) => next_field_argument(instruction, &mut position).ok()?,
        };
        unknown_switches.push(MergeFieldSwitch {
            name: name.to_ascii_lowercase(),
            argument,
        });
    }

    Some((variable_name, unknown_switches))
}

fn parse_dde_field_parts(instruction: &str) -> Option<DdeFieldParts> {
    if instruction.len() > MAX_DDE_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    let kind = if keyword.eq_ignore_ascii_case("DDE") {
        DdeFieldKind::Dde
    } else if keyword.eq_ignore_ascii_case("DDEAUTO") {
        DdeFieldKind::DdeAuto
    } else {
        return None;
    };

    let application = next_field_argument(instruction, &mut position).ok()??;
    if application.is_empty() {
        return None;
    }
    let source = next_field_argument(instruction, &mut position).ok()??;
    if source.is_empty() {
        return None;
    }

    skip_field_whitespace(instruction, &mut position);
    let item = match peek_field_character(instruction, position) {
        None | Some('\\') => None,
        Some(_) => Some(next_field_argument(instruction, &mut position).ok()??),
    };

    let mut automatic_updates = kind == DdeFieldKind::DdeAuto;
    let mut saw_automatic_update = false;
    let mut representation = None;
    let mut omit_graphic_data = false;
    let mut unknown_switches = Vec::new();
    loop {
        skip_field_whitespace(instruction, &mut position);
        let Some(introducer) = next_field_character(instruction, &mut position) else {
            break;
        };
        if introducer != '\\' || unknown_switches.len() >= MAX_DDE_FIELD_SWITCHES {
            return None;
        }

        let name = next_field_character(instruction, &mut position)?;
        if name == '\\' || name.is_whitespace() {
            return None;
        }
        let name = name.to_ascii_lowercase();

        skip_field_whitespace(instruction, &mut position);
        let argument = match peek_field_character(instruction, position) {
            None | Some('\\') => None,
            Some(_) => next_field_argument(instruction, &mut position).ok()?,
        };
        match name {
            'a' if kind == DdeFieldKind::Dde => {
                if saw_automatic_update || argument.is_some() {
                    return None;
                }
                automatic_updates = true;
                saw_automatic_update = true;
            },
            'a' => return None,
            'd' => {
                if representation.is_some() || omit_graphic_data || argument.is_some() {
                    return None;
                }
                omit_graphic_data = true;
            },
            'b' | 'h' | 'p' | 'r' | 't' | 'u' => {
                if representation.is_some() || omit_graphic_data || argument.is_some() {
                    return None;
                }
                representation = Some(match name {
                    'b' => DdeRepresentation::Bitmap,
                    'h' => DdeRepresentation::Html,
                    'p' => DdeRepresentation::Picture,
                    'r' => DdeRepresentation::RichText,
                    't' => DdeRepresentation::Text,
                    'u' => DdeRepresentation::UnicodeText,
                    _ => unreachable!("DDE representation switch was matched above"),
                });
            },
            _ => unknown_switches.push(MergeFieldSwitch { name, argument }),
        }
    }

    Some(DdeFieldParts {
        kind,
        application,
        source,
        item,
        automatic_updates,
        representation,
        omit_graphic_data,
        unknown_switches,
    })
}

fn parse_link_field_parts(instruction: &str) -> Option<LinkFieldParts> {
    if instruction.len() > MAX_LINK_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    if !keyword.eq_ignore_ascii_case("LINK") {
        return None;
    }

    let application_type = next_field_argument(instruction, &mut position).ok()??;
    if application_type.is_empty() {
        return None;
    }
    let source = next_field_argument(instruction, &mut position).ok()??;
    if source.is_empty() {
        return None;
    }

    skip_field_whitespace(instruction, &mut position);
    let item = match peek_field_character(instruction, position) {
        None | Some('\\') => None,
        Some(_) => Some(next_field_argument(instruction, &mut position).ok()??),
    };

    let mut switches = Vec::new();
    loop {
        skip_field_whitespace(instruction, &mut position);
        let Some(introducer) = next_field_character(instruction, &mut position) else {
            break;
        };
        if introducer != '\\' || switches.len() >= MAX_LINK_FIELD_SWITCHES {
            return None;
        }

        let name = next_field_character(instruction, &mut position)?;
        if name == '\\' || name.is_whitespace() {
            return None;
        }
        let name = name.to_ascii_lowercase();

        skip_field_whitespace(instruction, &mut position);
        let argument = match peek_field_character(instruction, position) {
            None | Some('\\') => None,
            Some(_) => next_field_argument(instruction, &mut position).ok()?,
        };
        switches.push(MergeFieldSwitch { name, argument });
    }

    let mut automatic_updates = false;
    let mut result_options = Vec::new();
    let mut formatting_modes = Vec::new();
    for switch in &switches {
        match switch.name {
            'a' => {
                if switch.argument.is_some() {
                    return None;
                }
                automatic_updates = true;
            },
            'f' => {
                let value = switch.argument.as_deref()?.parse::<i64>().ok()?;
                formatting_modes.push(match value {
                    0 => LinkFormatting::Source,
                    2 => LinkFormatting::Destination,
                    4 => LinkFormatting::SpreadsheetSource,
                    5 => LinkFormatting::SpreadsheetDestination,
                    other => LinkFormatting::Unsupported(other),
                });
            },
            'b' | 'd' | 'h' | 'p' | 'r' | 't' | 'u' => {
                if switch.argument.is_some() {
                    return None;
                }
                result_options.push(match switch.name {
                    'b' => LinkResultOption::Bitmap,
                    'd' => LinkResultOption::OmitGraphicData,
                    'h' => LinkResultOption::Html,
                    'p' => LinkResultOption::Picture,
                    'r' => LinkResultOption::RichText,
                    't' => LinkResultOption::Text,
                    'u' => LinkResultOption::UnicodeText,
                    _ => unreachable!("LINK result switch was matched above"),
                });
            },
            _ => {},
        }
    }

    Some(LinkFieldParts {
        application_type,
        source,
        item,
        automatic_updates,
        result_options,
        formatting_modes,
        switches,
    })
}

fn parse_mail_merge_counter_kind(instruction: &str) -> Option<MailMergeCounterKind> {
    if instruction.len() > MAX_MAIL_MERGE_COUNTER_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    let kind = if keyword.eq_ignore_ascii_case("MERGEREC") {
        MailMergeCounterKind::Record
    } else if keyword.eq_ignore_ascii_case("MERGESEQ") {
        MailMergeCounterKind::Sequence
    } else {
        return None;
    };
    if next_field_argument(instruction, &mut position)
        .ok()?
        .is_some()
    {
        return None;
    }

    Some(kind)
}

fn is_mail_merge_next_instruction(instruction: &str) -> bool {
    if instruction.len() > MAX_MAIL_MERGE_NEXT_INSTRUCTION_BYTES {
        return false;
    }

    let mut position = 0;
    let Ok(Some(keyword)) = next_field_argument(instruction, &mut position) else {
        return false;
    };
    keyword.eq_ignore_ascii_case("NEXT")
        && matches!(next_field_argument(instruction, &mut position), Ok(None))
}

fn parse_mail_merge_conditional_control_parts(
    instruction: &str,
) -> Option<(MailMergeConditionalControlKind, String)> {
    if instruction.len() > MAX_MAIL_MERGE_CONDITIONAL_CONTROL_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    let kind = if keyword.eq_ignore_ascii_case("NEXTIF") {
        MailMergeConditionalControlKind::NextIf
    } else if keyword.eq_ignore_ascii_case("SKIPIF") {
        MailMergeConditionalControlKind::SkipIf
    } else {
        return None;
    };
    let comparison = instruction.get(position..)?.trim();
    (!comparison.is_empty()).then_some((kind, comparison.to_string()))
}

fn parse_if_field_expression(instruction: &str) -> Option<String> {
    if instruction.len() > MAX_IF_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    if !keyword.eq_ignore_ascii_case("IF") {
        return None;
    }
    let expression = instruction.get(position..)?.trim();
    (!expression.is_empty()).then_some(expression.to_string())
}

fn parse_compare_field_comparison(instruction: &str) -> Option<String> {
    if instruction.len() > MAX_COMPARE_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    if !keyword.eq_ignore_ascii_case("COMPARE") {
        return None;
    }
    let comparison = instruction.get(position..)?.trim();
    (!comparison.is_empty()).then_some(comparison.to_string())
}

fn parse_prompt_field_parts(
    instruction: &str,
) -> Option<(
    PromptFieldKind,
    Option<String>,
    Option<String>,
    Option<String>,
    bool,
)> {
    if instruction.len() > MAX_PROMPT_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    let kind = if keyword.eq_ignore_ascii_case("ASK") {
        PromptFieldKind::Ask
    } else if keyword.eq_ignore_ascii_case("FILLIN") {
        PromptFieldKind::FillIn
    } else {
        return None;
    };

    let (bookmark, prompt) = match kind {
        PromptFieldKind::Ask => {
            let bookmark = next_field_argument(instruction, &mut position).ok()??;
            if bookmark.is_empty() {
                return None;
            }
            let prompt = next_field_argument(instruction, &mut position).ok()??;
            (Some(bookmark), Some(prompt))
        },
        PromptFieldKind::FillIn => {
            skip_field_whitespace(instruction, &mut position);
            let prompt = match peek_field_character(instruction, position) {
                None | Some('\\') => None,
                Some(_) => next_field_argument(instruction, &mut position).ok()?,
            };
            (None, prompt)
        },
    };

    let mut default_response = None;
    let mut prompts_once_per_mail_merge = false;
    loop {
        skip_field_whitespace(instruction, &mut position);
        let Some(introducer) = next_field_character(instruction, &mut position) else {
            break;
        };
        if introducer != '\\' {
            return None;
        }

        let name = next_field_character(instruction, &mut position)?;
        if name == '\\' || name.is_whitespace() {
            return None;
        }
        match name.to_ascii_lowercase() {
            'd' => {
                if default_response.is_some() {
                    return None;
                }
                default_response = Some(next_field_argument(instruction, &mut position).ok()??);
            },
            'o' => {
                if prompts_once_per_mail_merge {
                    return None;
                }
                skip_field_whitespace(instruction, &mut position);
                if !matches!(
                    peek_field_character(instruction, position),
                    None | Some('\\')
                ) {
                    return None;
                }
                prompts_once_per_mail_merge = true;
            },
            _ => return None,
        }
    }

    Some((
        kind,
        bookmark,
        prompt,
        default_response,
        prompts_once_per_mail_merge,
    ))
}

fn parse_user_identity_field_parts(
    instruction: &str,
) -> Option<(
    UserIdentityFieldKind,
    Option<String>,
    Option<UserIdentityFormatting>,
)> {
    if instruction.len() > MAX_USER_IDENTITY_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    let kind = if keyword.eq_ignore_ascii_case("USERADDRESS") {
        UserIdentityFieldKind::Address
    } else if keyword.eq_ignore_ascii_case("USERINITIALS") {
        UserIdentityFieldKind::Initials
    } else if keyword.eq_ignore_ascii_case("USERNAME") {
        UserIdentityFieldKind::Name
    } else {
        return None;
    };

    skip_field_whitespace(instruction, &mut position);
    let override_value = match peek_field_character(instruction, position) {
        None | Some('\\') => None,
        Some(_) => Some(next_field_argument(instruction, &mut position).ok()??),
    };

    let mut formatting = None;
    loop {
        skip_field_whitespace(instruction, &mut position);
        let Some(introducer) = next_field_character(instruction, &mut position) else {
            break;
        };
        if introducer != '\\' {
            return None;
        }
        let name = next_field_character(instruction, &mut position)?;
        if name != '*' || formatting.is_some() {
            return None;
        }
        let value = next_field_argument(instruction, &mut position).ok()??;
        formatting = Some(if value.eq_ignore_ascii_case("Caps") {
            UserIdentityFormatting::Caps
        } else if value.eq_ignore_ascii_case("FirstCap") {
            UserIdentityFormatting::FirstCap
        } else if value.eq_ignore_ascii_case("Lower") {
            UserIdentityFormatting::Lower
        } else if value.eq_ignore_ascii_case("Upper") {
            UserIdentityFormatting::Upper
        } else {
            return None;
        });
    }

    Some((kind, override_value, formatting))
}

fn parse_advance_field_adjustments(instruction: &str) -> Option<Vec<AdvanceFieldAdjustment>> {
    if instruction.len() > MAX_ADVANCE_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    if !keyword.eq_ignore_ascii_case("ADVANCE") {
        return None;
    }

    let mut adjustments = Vec::new();
    loop {
        skip_field_whitespace(instruction, &mut position);
        let Some(introducer) = next_field_character(instruction, &mut position) else {
            break;
        };
        if introducer != '\\' {
            return None;
        }
        let name = next_field_character(instruction, &mut position)?;
        let operation = match name.to_ascii_lowercase() {
            'd' => AdvanceFieldOperation::Down,
            'l' => AdvanceFieldOperation::Left,
            'r' => AdvanceFieldOperation::Right,
            'u' => AdvanceFieldOperation::Up,
            'x' => AdvanceFieldOperation::HorizontalPosition,
            'y' => AdvanceFieldOperation::VerticalPosition,
            _ => return None,
        };
        if adjustments.len() >= MAX_ADVANCE_FIELD_ADJUSTMENTS {
            return None;
        }
        let points = next_field_argument(instruction, &mut position)
            .ok()??
            .parse::<i64>()
            .ok()?;
        adjustments.push(AdvanceFieldAdjustment { operation, points });
    }

    Some(adjustments)
}

#[allow(clippy::type_complexity)]
fn parse_mail_merge_recipient_field_parts(
    instruction: &str,
) -> Option<(
    MailMergeRecipientFieldKind,
    Option<AddressBlockCountryInclusion>,
    bool,
    Vec<String>,
    Option<String>,
    Option<String>,
    Option<String>,
    Vec<MergeFieldSwitch>,
)> {
    if instruction.len() > MAX_MAIL_MERGE_RECIPIENT_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    let kind = if keyword.eq_ignore_ascii_case("ADDRESSBLOCK") {
        MailMergeRecipientFieldKind::AddressBlock
    } else if keyword.eq_ignore_ascii_case("GREETINGLINE") {
        MailMergeRecipientFieldKind::GreetingLine
    } else {
        return None;
    };

    let mut country_inclusion = None;
    let mut formats_using_recipient_country = false;
    let mut excluded_countries = Vec::new();
    let mut format_template = None;
    let mut language = None;
    let mut greeting_fallback_text = None;
    let mut unknown_switches = Vec::new();
    let mut switch_count = 0;
    loop {
        skip_field_whitespace(instruction, &mut position);
        let Some(introducer) = next_field_character(instruction, &mut position) else {
            break;
        };
        if introducer != '\\' || switch_count >= MAX_MAIL_MERGE_RECIPIENT_FIELD_SWITCHES {
            return None;
        }
        switch_count += 1;

        let name = next_field_character(instruction, &mut position)?;
        if name == '\\' || name.is_whitespace() {
            return None;
        }
        match (kind, name.to_ascii_lowercase()) {
            (MailMergeRecipientFieldKind::AddressBlock, 'c') => {
                if country_inclusion.is_some() {
                    return None;
                }
                let value = next_field_argument(instruction, &mut position).ok()??;
                country_inclusion = Some(match value.as_str() {
                    "0" => AddressBlockCountryInclusion::Omit,
                    "1" => AddressBlockCountryInclusion::Always,
                    "2" => AddressBlockCountryInclusion::UnlessExcluded,
                    _ => return None,
                });
            },
            (MailMergeRecipientFieldKind::AddressBlock, 'd') => {
                if formats_using_recipient_country {
                    return None;
                }
                skip_field_whitespace(instruction, &mut position);
                if !matches!(
                    peek_field_character(instruction, position),
                    None | Some('\\')
                ) {
                    return None;
                }
                formats_using_recipient_country = true;
            },
            (MailMergeRecipientFieldKind::AddressBlock, 'e') => {
                excluded_countries.push(next_field_argument(instruction, &mut position).ok()??);
            },
            (_, 'f') => {
                if format_template.is_some() {
                    return None;
                }
                format_template = Some(next_field_argument(instruction, &mut position).ok()??);
            },
            (_, 'l') => {
                if language.is_some() {
                    return None;
                }
                language = Some(next_field_argument(instruction, &mut position).ok()??);
            },
            (MailMergeRecipientFieldKind::GreetingLine, 'c' | 'e') => {
                if greeting_fallback_text.is_some() {
                    return None;
                }
                greeting_fallback_text =
                    Some(next_field_argument(instruction, &mut position).ok()??);
            },
            _ => {
                skip_field_whitespace(instruction, &mut position);
                let argument = match peek_field_character(instruction, position) {
                    None | Some('\\') => None,
                    Some(_) => next_field_argument(instruction, &mut position).ok()?,
                };
                unknown_switches.push(MergeFieldSwitch {
                    name: name.to_ascii_lowercase(),
                    argument,
                });
            },
        }
    }

    Some((
        kind,
        country_inclusion,
        formats_using_recipient_country,
        excluded_countries,
        format_template,
        language,
        greeting_fallback_text,
        unknown_switches,
    ))
}

fn next_field_argument(
    input: &str,
    position: &mut usize,
) -> std::result::Result<Option<String>, ()> {
    skip_field_whitespace(input, position);
    let Some(first) = next_field_character(input, position) else {
        return Ok(None);
    };

    if first != '"' {
        *position -= first.len_utf8();
        let mut value = String::new();
        while let Some(character) = next_field_character(input, position) {
            if character.is_whitespace() || character == '"' {
                *position -= character.len_utf8();
                break;
            }
            if character == '\\' {
                let escaped = next_field_character(input, position).ok_or(())?;
                if !matches!(escaped, '"' | '\\') {
                    return Err(());
                }
                value.push(escaped);
            } else {
                value.push(character);
            }
        }
        return Ok(Some(value));
    }

    let mut value = String::new();
    loop {
        let character = next_field_character(input, position).ok_or(())?;
        match character {
            '"' => return Ok(Some(value)),
            '\\' => {
                let escaped = next_field_character(input, position).ok_or(())?;
                if !matches!(escaped, '"' | '\\') {
                    return Err(());
                }
                value.push(escaped);
            },
            _ => value.push(character),
        }
    }
}

fn skip_field_whitespace(input: &str, position: &mut usize) {
    while let Some(character) = input.get(*position..).and_then(|rest| rest.chars().next()) {
        if !character.is_whitespace() {
            break;
        }
        *position += character.len_utf8();
    }
}

fn next_field_character(input: &str, position: &mut usize) -> Option<char> {
    let character = input.get(*position..)?.chars().next()?;
    *position += character.len_utf8();
    Some(character)
}

fn peek_field_character(input: &str, position: usize) -> Option<char> {
    input.get(position..)?.chars().next()
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
                    return Err(corrupted(
                        "grffldEnd.fNested disagrees with field containment",
                    ));
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

    pub(crate) fn field_texts<F>(&self, mut text_at_range: F) -> Result<Vec<FieldText>>
    where
        F: FnMut(FieldStory, u32, u32) -> Result<String>,
    {
        self.stories
            .iter()
            .flat_map(|story| story.fields())
            .map(|field| {
                FieldText::from_field(field, |start, end| text_at_range(field.story, start, end))
            })
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

    #[test]
    fn field_text_keeps_instruction_and_cached_result_separate() {
        let field = Field {
            story: FieldStory::Header,
            start_cp: 2,
            separator_cp: Some(17),
            end_cp: 23,
            field_type: FieldType::IncludeText,
            end_flags: FieldEndFlags {
                has_separator: true,
                ..FieldEndFlags::default()
            },
            nesting_depth: 0,
            has_separator: true,
        };

        let text = FieldText::from_field(&field, |start, end| match (start, end) {
            (3, 17) => Ok(r#" INCLUDETEXT "file:///draft.doc" "#.to_string()),
            (18, 23) => Ok("cached".to_string()),
            _ => Err(corrupted("unexpected field range")),
        })
        .unwrap();

        assert_eq!(text.field, field);
        assert_eq!(text.instruction, r#" INCLUDETEXT "file:///draft.doc" "#);
        assert_eq!(text.result.as_deref(), Some("cached"));
    }

    #[test]
    fn field_text_reports_absent_separator_as_no_cached_result() {
        let field = Field {
            story: FieldStory::Main,
            start_cp: 4,
            separator_cp: None,
            end_cp: 12,
            field_type: FieldType::MacroButton,
            end_flags: FieldEndFlags::default(),
            nesting_depth: 0,
            has_separator: false,
        };

        let text = FieldText::from_field(&field, |start, end| {
            assert_eq!((start, end), (5, 12));
            Ok(" MACROBUTTON NeverRun Label ".to_string())
        })
        .unwrap();

        assert_eq!(text.instruction, " MACROBUTTON NeverRun Label ");
        assert_eq!(text.result, None);
        let macro_button = text.macro_button().unwrap();
        assert_eq!(macro_button.macro_name(), "NeverRun");
        assert_eq!(macro_button.display_text(), "Label");
        assert_eq!(macro_button.cached_result(), None);
        assert!(!macro_button.is_dirty());
        assert!(!macro_button.is_locked());
    }

    #[test]
    fn macro_button_field_exposes_stored_metadata_without_execution() {
        let field = Field {
            story: FieldStory::Textbox,
            start_cp: 4,
            separator_cp: Some(37),
            end_cp: 44,
            field_type: FieldType::MacroButton,
            end_flags: FieldEndFlags {
                results_dirty: true,
                locked: true,
                has_separator: true,
                ..FieldEndFlags::default()
            },
            nesting_depth: 1,
            has_separator: true,
        };
        let text = FieldText {
            field: field.clone(),
            instruction: r#" MACROBUTTON "Never Run" "Click \"here\"\\now" "#.to_string(),
            result: Some("cached button".to_string()),
        };

        let macro_button = text.macro_button().unwrap();
        assert_eq!(macro_button.field(), &field);
        assert_eq!(macro_button.macro_name(), "Never Run");
        assert_eq!(macro_button.display_text(), r#"Click "here"\now"#);
        assert_eq!(macro_button.cached_result(), Some("cached button"));
        assert!(macro_button.is_dirty());
        assert!(macro_button.is_locked());

        let compact = FieldText {
            instruction: r#"MACROBUTTON"Never Run""Click""#.to_string(),
            ..text.clone()
        };
        let compact_button = compact.macro_button().unwrap();
        assert_eq!(compact_button.macro_name(), "Never Run");
        assert_eq!(compact_button.display_text(), "Click");

        let missing_button = FieldText {
            instruction: "MACROBUTTON NeverRun".to_string(),
            ..text.clone()
        };
        assert!(missing_button.macro_button().is_none());

        let empty_button = FieldText {
            instruction: r#"MACROBUTTON NeverRun """#.to_string(),
            ..text.clone()
        };
        assert!(empty_button.macro_button().is_none());

        let extra_argument = FieldText {
            instruction: "MACROBUTTON NeverRun Button unexpected".to_string(),
            ..text.clone()
        };
        assert!(extra_argument.macro_button().is_none());

        let invalid_escape = FieldText {
            instruction: r#"MACROBUTTON NeverRun "Click \now""#.to_string(),
            ..text.clone()
        };
        assert!(invalid_escape.macro_button().is_none());

        let wrong_keyword = FieldText {
            instruction: "DOCVARIABLE Customer".to_string(),
            ..text
        };
        assert!(wrong_keyword.macro_button().is_none());
    }

    #[test]
    fn go_to_button_field_exposes_stored_metadata_without_navigation() {
        let field = Field {
            story: FieldStory::Textbox,
            start_cp: 4,
            separator_cp: Some(37),
            end_cp: 52,
            field_type: FieldType::GoToButton,
            end_flags: FieldEndFlags {
                results_dirty: true,
                locked: true,
                has_separator: true,
                ..FieldEndFlags::default()
            },
            nesting_depth: 1,
            has_separator: true,
        };
        let text = FieldText {
            field: field.clone(),
            instruction: r#" GOTOBUTTON "f 2" "Footnote" "#.to_string(),
            result: Some("cached footnote button".to_string()),
        };

        let button = text.go_to_button().unwrap();
        assert_eq!(button.field(), &field);
        assert_eq!(button.target(), "f 2");
        assert_eq!(button.button_text(), "Footnote");
        assert_eq!(button.cached_result(), Some("cached footnote button"));
        assert!(button.is_dirty());
        assert!(button.is_locked());

        for instruction in [
            "GOTOBUTTON",
            r#"GOTOBUTTON "" Button"#,
            "GOTOBUTTON Destination",
            r#"GOTOBUTTON Destination """#,
            "GOTOBUTTON Destination Button unexpected",
            r#"GOTOBUTTON Destination Button \* MERGEFORMAT"#,
            r#"GOTOBUTTON Destination "Button \now""#,
        ] {
            let malformed = FieldText {
                field: field.clone(),
                instruction: instruction.to_string(),
                result: None,
            };
            assert!(malformed.go_to_button().is_none(), "{instruction}");
        }

        let wrong_keyword = FieldText {
            field: field.clone(),
            instruction: "GOTOBUTTONS Destination Button".to_string(),
            result: None,
        };
        assert!(wrong_keyword.go_to_button().is_none());

        let wrong_type = FieldText {
            field: Field {
                field_type: FieldType::MacroButton,
                ..field
            },
            instruction: "GOTOBUTTON Destination Button".to_string(),
            result: None,
        };
        assert!(wrong_type.go_to_button().is_none());
    }

    #[test]
    fn merge_field_exposes_stored_metadata_without_merging() {
        let field = Field {
            story: FieldStory::Textbox,
            start_cp: 4,
            separator_cp: Some(37),
            end_cp: 52,
            field_type: FieldType::MergeField,
            end_flags: FieldEndFlags {
                results_dirty: true,
                locked: true,
                has_separator: true,
                ..FieldEndFlags::default()
            },
            nesting_depth: 1,
            has_separator: true,
        };
        let text = FieldText {
            field: field.clone(),
            instruction: r#" MERGEFIELD "Customer Region" \b "Dear " \f "!" \m \v \* MERGEFORMAT "#
                .to_string(),
            result: Some("cached customer".to_string()),
        };

        let merge = text.merge_field().unwrap();
        assert_eq!(merge.field(), &field);
        assert_eq!(merge.instruction(), text.instruction);
        assert_eq!(merge.field_name(), "Customer Region");
        assert_eq!(merge.cached_result(), Some("cached customer"));
        assert!(merge.is_dirty());
        assert!(merge.is_locked());
        assert_eq!(merge.switches().len(), 5);
        assert_eq!(merge.switches()[0].name(), 'b');
        assert_eq!(merge.switches()[0].argument(), Some("Dear "));
        assert_eq!(merge.switches()[1].name(), 'f');
        assert_eq!(merge.switches()[1].argument(), Some("!"));
        assert!(merge.has_switch('m'));
        assert!(merge.has_switch('v'));
        assert!(merge.has_switch('*'));
        assert_eq!(merge.switches()[4].argument(), Some("MERGEFORMAT"));

        let compact = FieldText {
            instruction: r#"MERGEFIELD"Customer Name"\f" ""#.to_string(),
            ..text.clone()
        };
        let compact_merge = compact.merge_field().unwrap();
        assert_eq!(compact_merge.field_name(), "Customer Name");
        assert_eq!(compact_merge.switches()[0].argument(), Some(" "));

        let missing_name = FieldText {
            instruction: r#"MERGEFIELD \* MERGEFORMAT"#.to_string(),
            ..text.clone()
        };
        assert!(missing_name.merge_field().is_none());

        let unexpected_operand = FieldText {
            instruction: "MERGEFIELD Customer unexpected".to_string(),
            ..text.clone()
        };
        assert!(unexpected_operand.merge_field().is_none());

        let wrong_keyword = FieldText {
            instruction: "MERGEFIELDS Customer".to_string(),
            ..text.clone()
        };
        assert!(wrong_keyword.merge_field().is_none());

        let wrong_type = FieldText {
            field: Field {
                field_type: FieldType::DocumentVariable,
                ..field
            },
            ..text
        };
        assert!(wrong_type.merge_field().is_none());
    }

    #[test]
    fn document_variable_fields_expose_cached_metadata_without_resolution() {
        let field = Field {
            story: FieldStory::Textbox,
            start_cp: 4,
            separator_cp: Some(37),
            end_cp: 52,
            field_type: FieldType::DocumentVariable,
            end_flags: FieldEndFlags {
                results_dirty: true,
                locked: true,
                has_separator: true,
                ..FieldEndFlags::default()
            },
            nesting_depth: 1,
            has_separator: true,
        };
        let text = FieldText {
            field: field.clone(),
            instruction: r#" DOCVARIABLE "Customer Region" \* MERGEFORMAT "#.to_string(),
            result: Some("cached region".to_string()),
        };

        let variable = text.document_variable().unwrap();
        assert_eq!(variable.field(), &field);
        assert_eq!(variable.instruction(), text.instruction);
        assert_eq!(variable.variable_name(), "Customer Region");
        assert_eq!(variable.cached_result(), Some("cached region"));
        assert!(variable.is_dirty());
        assert!(variable.is_locked());
        assert_eq!(variable.unknown_switches().len(), 1);
        assert_eq!(variable.unknown_switches()[0].name(), '*');
        assert_eq!(
            variable.unknown_switches()[0].argument(),
            Some("MERGEFORMAT")
        );

        let compact = FieldText {
            instruction: r#"DOCVARIABLE"Customer Name"\*MERGEFORMAT"#.to_string(),
            ..text.clone()
        };
        let compact_variable = compact.document_variable().unwrap();
        assert_eq!(compact_variable.variable_name(), "Customer Name");
        assert_eq!(
            compact_variable.unknown_switches()[0].argument(),
            Some("MERGEFORMAT")
        );

        let missing_name = FieldText {
            instruction: r#"DOCVARIABLE \* MERGEFORMAT"#.to_string(),
            ..text.clone()
        };
        assert!(missing_name.document_variable().is_none());

        let unexpected_operand = FieldText {
            instruction: "DOCVARIABLE Customer unexpected".to_string(),
            ..text.clone()
        };
        assert!(unexpected_operand.document_variable().is_none());

        let wrong_keyword = FieldText {
            instruction: "DOCVARIABLES Customer".to_string(),
            ..text.clone()
        };
        assert!(wrong_keyword.document_variable().is_none());

        let wrong_type = FieldText {
            field: Field {
                field_type: FieldType::MergeField,
                ..field
            },
            ..text
        };
        assert!(wrong_type.document_variable().is_none());
    }

    #[test]
    fn dde_links_expose_cached_metadata_without_activation() {
        let field = Field {
            story: FieldStory::Textbox,
            start_cp: 4,
            separator_cp: Some(37),
            end_cp: 52,
            field_type: FieldType::Dde,
            end_flags: FieldEndFlags {
                results_dirty: true,
                locked: true,
                has_separator: true,
                ..FieldEndFlags::default()
            },
            nesting_depth: 1,
            has_separator: true,
        };
        let text = FieldText {
            field: field.clone(),
            instruction: r#" DDE Excel "missing.xlsx" "Sheet1!A1" \a \p \x "ignored" "#.to_string(),
            result: Some("cached DDE".to_string()),
        };

        let dde = text.dde_link().unwrap();
        assert_eq!(dde.field(), &field);
        assert_eq!(dde.instruction(), text.instruction);
        assert_eq!(dde.kind(), DdeFieldKind::Dde);
        assert_eq!(dde.application(), "Excel");
        assert_eq!(dde.source(), "missing.xlsx");
        assert_eq!(dde.item(), Some("Sheet1!A1"));
        assert!(dde.requests_automatic_updates());
        assert_eq!(dde.representation(), Some(DdeRepresentation::Picture));
        assert!(!dde.omits_graphic_data());
        assert_eq!(dde.cached_result(), Some("cached DDE"));
        assert!(dde.is_dirty());
        assert!(dde.is_locked());
        assert_eq!(dde.unknown_switches().len(), 1);
        assert_eq!(dde.unknown_switches()[0].name(), 'x');
        assert_eq!(dde.unknown_switches()[0].argument(), Some("ignored"));

        let automatic = FieldText {
            field: Field {
                field_type: FieldType::DdeAuto,
                ..field.clone()
            },
            instruction: r#"DDEAUTO Excel "missing.xlsx" "Sheet1!A2" \t"#.to_string(),
            result: Some("cached auto".to_string()),
        };
        let automatic = automatic.dde_link().unwrap();
        assert_eq!(automatic.kind(), DdeFieldKind::DdeAuto);
        assert_eq!(automatic.item(), Some("Sheet1!A2"));
        assert!(automatic.requests_automatic_updates());
        assert_eq!(automatic.representation(), Some(DdeRepresentation::Text));

        for instruction in [
            r#"DDE Excel source \p \t"#,
            r#"DDEAUTO Excel source \a"#,
            r#"DDE Excel source \d \p"#,
            r#"DDE Excel source \a value"#,
        ] {
            let malformed = FieldText {
                instruction: instruction.to_string(),
                ..text.clone()
            };
            assert!(malformed.dde_link().is_none(), "{instruction}");
        }

        let wrong_keyword = FieldText {
            instruction: "DDEAUTOMATED Excel source".to_string(),
            ..text.clone()
        };
        assert!(wrong_keyword.dde_link().is_none());

        let wrong_type = FieldText {
            field: Field {
                field_type: FieldType::Link,
                ..field
            },
            ..text
        };
        assert!(wrong_type.dde_link().is_none());
    }

    #[test]
    fn link_fields_expose_cached_metadata_without_activating_sources() {
        let field = Field {
            story: FieldStory::Textbox,
            start_cp: 4,
            separator_cp: Some(37),
            end_cp: 52,
            field_type: FieldType::Link,
            end_flags: FieldEndFlags {
                results_dirty: true,
                locked: true,
                has_separator: true,
                ..FieldEndFlags::default()
            },
            nesting_depth: 1,
            has_separator: true,
        };
        let text = FieldText {
            field: field.clone(),
            instruction:
                r#" LINK Excel.Sheet.12 "missing.xlsx" "Sheet1!A1" \a \p \f 2 \f 9 \x "ignored" "#
                    .to_string(),
            result: Some("cached link".to_string()),
        };

        let link = text.link_field().unwrap();
        assert_eq!(link.field(), &field);
        assert_eq!(link.instruction(), text.instruction);
        assert_eq!(link.application_type(), "Excel.Sheet.12");
        assert_eq!(link.source(), "missing.xlsx");
        assert_eq!(link.item(), Some("Sheet1!A1"));
        assert!(link.requests_automatic_updates());
        assert_eq!(link.result_options(), &[LinkResultOption::Picture]);
        assert_eq!(
            link.effective_result_option(),
            Some(LinkResultOption::Picture)
        );
        assert_eq!(
            link.formatting_modes(),
            &[LinkFormatting::Destination, LinkFormatting::Unsupported(9)]
        );
        assert_eq!(link.cached_result(), Some("cached link"));
        assert!(link.is_dirty());
        assert!(link.is_locked());
        assert_eq!(link.switches().len(), 5);
        assert_eq!(link.switches()[4].name(), 'x');
        assert_eq!(link.switches()[4].argument(), Some("ignored"));

        let no_item = FieldText {
            instruction: r#"LINK Excel.Sheet.12 "missing.xlsx" \d \b"#.to_string(),
            ..text.clone()
        };
        let no_item = no_item.link_field().unwrap();
        assert_eq!(no_item.item(), None);
        assert_eq!(
            no_item.result_options(),
            &[LinkResultOption::OmitGraphicData, LinkResultOption::Bitmap]
        );
        assert_eq!(
            no_item.effective_result_option(),
            Some(LinkResultOption::Bitmap)
        );

        for instruction in [
            r#"LINK Excel source \a value"#,
            r#"LINK Excel source \p value"#,
            r#"LINK Excel source \f"#,
            r#"LINK Excel source \f not-an-integer"#,
        ] {
            let malformed = FieldText {
                instruction: instruction.to_string(),
                ..text.clone()
            };
            assert!(malformed.link_field().is_none(), "{instruction}");
        }

        let wrong_keyword = FieldText {
            instruction: "LINKS Excel source".to_string(),
            ..text.clone()
        };
        assert!(wrong_keyword.link_field().is_none());

        let wrong_type = FieldText {
            field: Field {
                field_type: FieldType::Dde,
                ..field
            },
            ..text
        };
        assert!(wrong_type.link_field().is_none());
    }

    #[test]
    fn mail_merge_counters_expose_cached_metadata_without_merging() {
        let record_field = Field {
            story: FieldStory::Textbox,
            start_cp: 4,
            separator_cp: Some(13),
            end_cp: 16,
            field_type: FieldType::MergeRecord,
            end_flags: FieldEndFlags {
                results_dirty: true,
                locked: true,
                has_separator: true,
                ..FieldEndFlags::default()
            },
            nesting_depth: 1,
            has_separator: true,
        };
        let record = FieldText {
            field: record_field.clone(),
            instruction: " MERGEREC ".to_string(),
            result: Some("12".to_string()),
        };

        let counter = record.mail_merge_counter().unwrap();
        assert_eq!(counter.field(), &record_field);
        assert_eq!(counter.instruction(), record.instruction);
        assert_eq!(counter.kind(), MailMergeCounterKind::Record);
        assert_eq!(counter.cached_result(), Some("12"));
        assert!(counter.is_dirty());
        assert!(counter.is_locked());

        let sequence = FieldText {
            field: Field {
                field_type: FieldType::MergeSequence,
                ..record_field.clone()
            },
            instruction: "mergeSEQ".to_string(),
            result: Some("3".to_string()),
        };
        let sequence_counter = sequence.mail_merge_counter().unwrap();
        assert_eq!(sequence_counter.kind(), MailMergeCounterKind::Sequence);
        assert_eq!(sequence_counter.cached_result(), Some("3"));

        let unexpected_operand = FieldText {
            instruction: "MERGEREC 12".to_string(),
            ..record.clone()
        };
        assert!(unexpected_operand.mail_merge_counter().is_none());

        let unexpected_switch = FieldText {
            instruction: r"MERGESEQ \* MERGEFORMAT".to_string(),
            ..sequence.clone()
        };
        assert!(unexpected_switch.mail_merge_counter().is_none());

        let wrong_keyword = FieldText {
            instruction: "MERGERECORD".to_string(),
            ..record.clone()
        };
        assert!(wrong_keyword.mail_merge_counter().is_none());

        let wrong_type = FieldText {
            field: Field {
                field_type: FieldType::MergeSequence,
                ..record_field
            },
            ..record
        };
        assert!(wrong_type.mail_merge_counter().is_none());
    }

    #[test]
    fn mail_merge_next_fields_expose_cached_metadata_without_advancing_records() {
        let field = Field {
            story: FieldStory::Textbox,
            start_cp: 4,
            separator_cp: Some(9),
            end_cp: 22,
            field_type: FieldType::Next,
            end_flags: FieldEndFlags {
                results_dirty: true,
                locked: true,
                has_separator: true,
                ..FieldEndFlags::default()
            },
            nesting_depth: 1,
            has_separator: true,
        };
        let text = FieldText {
            field: field.clone(),
            instruction: " NEXT ".to_string(),
            result: Some("cached next".to_string()),
        };

        let next = text.mail_merge_next().unwrap();
        assert_eq!(next.field(), &field);
        assert_eq!(next.instruction(), text.instruction);
        assert_eq!(next.cached_result(), Some("cached next"));
        assert!(next.is_dirty());
        assert!(next.is_locked());

        let unexpected_operand = FieldText {
            instruction: "NEXT 12".to_string(),
            ..text.clone()
        };
        assert!(unexpected_operand.mail_merge_next().is_none());

        let unexpected_switch = FieldText {
            instruction: r"NEXT \* MERGEFORMAT".to_string(),
            ..text.clone()
        };
        assert!(unexpected_switch.mail_merge_next().is_none());

        let wrong_keyword = FieldText {
            instruction: "NEXTIF Customer = Ada".to_string(),
            ..text.clone()
        };
        assert!(wrong_keyword.mail_merge_next().is_none());

        let wrong_type = FieldText {
            field: Field {
                field_type: FieldType::NextIf,
                ..field
            },
            ..text
        };
        assert!(wrong_type.mail_merge_next().is_none());
    }

    #[test]
    fn conditional_mail_merge_controls_expose_metadata_without_merging() {
        let field = Field {
            story: FieldStory::Textbox,
            start_cp: 4,
            separator_cp: Some(9),
            end_cp: 22,
            field_type: FieldType::NextIf,
            end_flags: FieldEndFlags {
                results_dirty: true,
                locked: true,
                has_separator: true,
                ..FieldEndFlags::default()
            },
            nesting_depth: 1,
            has_separator: true,
        };
        let next_if = FieldText {
            field: field.clone(),
            instruction: r#" NEXTIF Customer = "Ada" "#.to_string(),
            result: Some("cached nextif".to_string()),
        };

        let control = next_if.mail_merge_conditional_control().unwrap();
        assert_eq!(control.field(), &field);
        assert_eq!(control.instruction(), next_if.instruction);
        assert_eq!(control.kind(), MailMergeConditionalControlKind::NextIf);
        assert_eq!(control.comparison(), r#"Customer = "Ada""#);
        assert_eq!(control.cached_result(), Some("cached nextif"));
        assert!(control.is_dirty());
        assert!(control.is_locked());

        let skip_if = FieldText {
            field: Field {
                field_type: FieldType::SkipIf,
                ..field.clone()
            },
            instruction: "skipif MERGEFIELD Order < 100".to_string(),
            result: Some("cached skipif".to_string()),
        };
        let skip_control = skip_if.mail_merge_conditional_control().unwrap();
        assert_eq!(skip_control.kind(), MailMergeConditionalControlKind::SkipIf);
        assert_eq!(skip_control.comparison(), "MERGEFIELD Order < 100");
        assert_eq!(skip_control.cached_result(), Some("cached skipif"));

        let missing_comparison = FieldText {
            instruction: "NEXTIF".to_string(),
            ..next_if.clone()
        };
        assert!(
            missing_comparison
                .mail_merge_conditional_control()
                .is_none()
        );

        let wrong_keyword = FieldText {
            instruction: "NEXTIFF Customer = Ada".to_string(),
            ..next_if.clone()
        };
        assert!(wrong_keyword.mail_merge_conditional_control().is_none());

        let wrong_type = FieldText {
            field: Field {
                field_type: FieldType::SkipIf,
                ..field
            },
            ..next_if
        };
        assert!(wrong_type.mail_merge_conditional_control().is_none());
    }

    #[test]
    fn if_fields_expose_cached_metadata_without_evaluation() {
        let field = Field {
            story: FieldStory::Textbox,
            start_cp: 4,
            separator_cp: Some(9),
            end_cp: 22,
            field_type: FieldType::If,
            end_flags: FieldEndFlags {
                results_dirty: true,
                locked: true,
                has_separator: true,
                ..FieldEndFlags::default()
            },
            nesting_depth: 1,
            has_separator: true,
        };
        let text = FieldText {
            field: field.clone(),
            instruction: r#" IF "A" = "A" "yes" "no" "#.to_string(),
            result: Some("yes".to_string()),
        };

        let if_field = text.if_field().unwrap();
        assert_eq!(if_field.field(), &field);
        assert_eq!(if_field.instruction(), text.instruction);
        assert_eq!(if_field.expression(), r#""A" = "A" "yes" "no""#);
        assert_eq!(if_field.cached_result(), Some("yes"));
        assert!(if_field.is_dirty());
        assert!(if_field.is_locked());

        let missing_expression = FieldText {
            instruction: "IF".to_string(),
            ..text.clone()
        };
        assert!(missing_expression.if_field().is_none());

        let wrong_keyword = FieldText {
            instruction: r#"IFF "A" = "A" "yes" "no""#.to_string(),
            ..text.clone()
        };
        assert!(wrong_keyword.if_field().is_none());

        let wrong_type = FieldText {
            field: Field {
                field_type: FieldType::NextIf,
                ..field
            },
            ..text
        };
        assert!(wrong_type.if_field().is_none());
    }

    #[test]
    fn compare_fields_expose_cached_metadata_without_evaluation() {
        let field = Field {
            story: FieldStory::Textbox,
            start_cp: 4,
            separator_cp: Some(9),
            end_cp: 22,
            field_type: FieldType::Compare,
            end_flags: FieldEndFlags {
                results_dirty: true,
                locked: true,
                has_separator: true,
                ..FieldEndFlags::default()
            },
            nesting_depth: 1,
            has_separator: true,
        };
        let text = FieldText {
            field: field.clone(),
            instruction: r#" COMPARE "CustomerNumber" >= 4 "#.to_string(),
            result: Some("1".to_string()),
        };

        let compare_field = text.compare_field().unwrap();
        assert_eq!(compare_field.field(), &field);
        assert_eq!(compare_field.instruction(), text.instruction);
        assert_eq!(compare_field.comparison(), r#""CustomerNumber" >= 4"#);
        assert_eq!(compare_field.cached_result(), Some("1"));
        assert!(compare_field.is_dirty());
        assert!(compare_field.is_locked());

        let nested = FieldText {
            instruction: "compare MERGEFIELD CustomerRating <= 9".to_string(),
            ..text.clone()
        };
        let compare_field = nested.compare_field().unwrap();
        assert_eq!(compare_field.comparison(), "MERGEFIELD CustomerRating <= 9");

        let missing_comparison = FieldText {
            instruction: "COMPARE".to_string(),
            ..text.clone()
        };
        assert!(missing_comparison.compare_field().is_none());

        let wrong_keyword = FieldText {
            instruction: "COMPARES Customer = 1".to_string(),
            ..text.clone()
        };
        assert!(wrong_keyword.compare_field().is_none());

        let wrong_type = FieldText {
            field: Field {
                field_type: FieldType::If,
                ..field
            },
            ..text
        };
        assert!(wrong_type.compare_field().is_none());
    }

    #[test]
    fn prompt_fields_expose_cached_metadata_without_displaying_prompts() {
        let field = Field {
            story: FieldStory::Textbox,
            start_cp: 4,
            separator_cp: Some(9),
            end_cp: 22,
            field_type: FieldType::Ask,
            end_flags: FieldEndFlags {
                results_dirty: true,
                locked: true,
                has_separator: true,
                ..FieldEndFlags::default()
            },
            nesting_depth: 1,
            has_separator: true,
        };
        let ask = FieldText {
            field: field.clone(),
            instruction: r#" ASK AskResponse "What is your first name?" \d "" \o "#.to_string(),
            result: Some("cached ask response".to_string()),
        };

        let prompt = ask.prompt_field().unwrap();
        assert_eq!(prompt.field(), &field);
        assert_eq!(prompt.instruction(), ask.instruction);
        assert_eq!(prompt.kind(), PromptFieldKind::Ask);
        assert_eq!(prompt.bookmark(), Some("AskResponse"));
        assert_eq!(prompt.prompt(), Some("What is your first name?"));
        assert_eq!(prompt.default_response(), Some(""));
        assert!(prompt.prompts_once_per_mail_merge());
        assert_eq!(prompt.cached_result(), Some("cached ask response"));
        assert!(prompt.is_dirty());
        assert!(prompt.is_locked());

        let fill_in = FieldText {
            field: Field {
                field_type: FieldType::FillIn,
                ..field.clone()
            },
            instruction: r#"fillin "Enter appointment time" \d "09:00""#.to_string(),
            result: Some("10:30".to_string()),
        };
        let fill_in_prompt = fill_in.prompt_field().unwrap();
        assert_eq!(fill_in_prompt.kind(), PromptFieldKind::FillIn);
        assert_eq!(fill_in_prompt.bookmark(), None);
        assert_eq!(fill_in_prompt.prompt(), Some("Enter appointment time"));
        assert_eq!(fill_in_prompt.default_response(), Some("09:00"));
        assert!(!fill_in_prompt.prompts_once_per_mail_merge());
        assert_eq!(fill_in_prompt.cached_result(), Some("10:30"));

        let default_only = FieldText {
            instruction: r#"FILLIN \d "recent response" \o"#.to_string(),
            result: None,
            ..fill_in.clone()
        };
        let default_only_prompt = default_only.prompt_field().unwrap();
        assert_eq!(default_only_prompt.prompt(), None);
        assert_eq!(
            default_only_prompt.default_response(),
            Some("recent response")
        );
        assert!(default_only_prompt.prompts_once_per_mail_merge());

        for instruction in [
            "ASK",
            r#"ASK "" "Question""#,
            "ASK Answer",
            r#"ASK Answer "Question" \d"#,
            r#"ASK Answer "Question" \o extra"#,
            r#"FILLIN "Question" \x"#,
            r#"FILLIN "Question" \d "first" \d "second""#,
            r#"FILLIN "Question" \o \o"#,
        ] {
            let malformed = FieldText {
                instruction: instruction.to_string(),
                ..ask.clone()
            };
            assert!(malformed.prompt_field().is_none(), "{instruction}");
        }

        let wrong_keyword = FieldText {
            instruction: r#"ASKER Answer "Question""#.to_string(),
            ..ask.clone()
        };
        assert!(wrong_keyword.prompt_field().is_none());

        let wrong_type = FieldText {
            field: Field {
                field_type: FieldType::FillIn,
                ..field
            },
            ..ask
        };
        assert!(wrong_type.prompt_field().is_none());
    }

    #[test]
    fn user_identity_fields_expose_metadata_without_reading_host_identity() {
        let field = Field {
            story: FieldStory::Textbox,
            start_cp: 4,
            separator_cp: Some(9),
            end_cp: 22,
            field_type: FieldType::UserAddress,
            end_flags: FieldEndFlags {
                results_dirty: true,
                locked: true,
                has_separator: true,
                ..FieldEndFlags::default()
            },
            nesting_depth: 1,
            has_separator: true,
        };
        let address = FieldText {
            field: field.clone(),
            instruction: r#" USERADDRESS "10 Top Secret Lane" \* Upper "#.to_string(),
            result: Some("10 TOP SECRET LANE".to_string()),
        };

        let address_field = address.user_identity_field().unwrap();
        assert_eq!(address_field.field(), &field);
        assert_eq!(address_field.instruction(), address.instruction);
        assert_eq!(address_field.kind(), UserIdentityFieldKind::Address);
        assert_eq!(address_field.override_value(), Some("10 Top Secret Lane"));
        assert_eq!(
            address_field.formatting(),
            Some(UserIdentityFormatting::Upper)
        );
        assert_eq!(address_field.cached_result(), Some("10 TOP SECRET LANE"));
        assert!(address_field.is_dirty());
        assert!(address_field.is_locked());

        let initials = FieldText {
            field: Field {
                field_type: FieldType::UserInitials,
                ..field.clone()
            },
            instruction: r#"userinitials \* Lower"#.to_string(),
            result: Some("dw".to_string()),
        };
        let initials_field = initials.user_identity_field().unwrap();
        assert_eq!(initials_field.kind(), UserIdentityFieldKind::Initials);
        assert_eq!(initials_field.override_value(), None);
        assert_eq!(
            initials_field.formatting(),
            Some(UserIdentityFormatting::Lower)
        );
        assert_eq!(initials_field.cached_result(), Some("dw"));

        let name = FieldText {
            field: Field {
                field_type: FieldType::UserName,
                ..field.clone()
            },
            instruction: r#"USERNAME "Ada Lovelace" \* FirstCap"#.to_string(),
            result: Some("Ada Lovelace".to_string()),
        };
        let name_field = name.user_identity_field().unwrap();
        assert_eq!(name_field.kind(), UserIdentityFieldKind::Name);
        assert_eq!(name_field.override_value(), Some("Ada Lovelace"));
        assert_eq!(
            name_field.formatting(),
            Some(UserIdentityFormatting::FirstCap)
        );
        assert_eq!(name_field.cached_result(), Some("Ada Lovelace"));

        let blank_override = FieldText {
            field: Field {
                field_type: FieldType::UserName,
                ..field.clone()
            },
            instruction: r#"USERNAME "" \* Caps"#.to_string(),
            result: None,
        };
        let blank_override = blank_override.user_identity_field().unwrap();
        assert_eq!(blank_override.override_value(), Some(""));
        assert_eq!(
            blank_override.formatting(),
            Some(UserIdentityFormatting::Caps)
        );

        for instruction in [
            r#"USERADDRESS \*"#,
            r#"USERADDRESS \* Title"#,
            r#"USERADDRESS Ada \* Upper \* Lower"#,
            r#"USERADDRESS Ada \l 1033"#,
            "USERADDRESS Ada Lovelace",
        ] {
            let malformed = FieldText {
                instruction: instruction.to_string(),
                result: None,
                ..address.clone()
            };
            assert!(malformed.user_identity_field().is_none(), "{instruction}");
        }

        let wrong_keyword = FieldText {
            instruction: "USERADDRESSES Ada".to_string(),
            result: None,
            ..address.clone()
        };
        assert!(wrong_keyword.user_identity_field().is_none());

        let wrong_type = FieldText {
            field: Field {
                field_type: FieldType::UserInitials,
                ..field
            },
            result: None,
            ..address
        };
        assert!(wrong_type.user_identity_field().is_none());
    }

    #[test]
    fn advance_fields_expose_metadata_without_changing_layout() {
        let field = Field {
            story: FieldStory::Textbox,
            start_cp: 4,
            separator_cp: Some(9),
            end_cp: 22,
            field_type: FieldType::Advance,
            end_flags: FieldEndFlags {
                results_dirty: true,
                locked: true,
                has_separator: true,
                ..FieldEndFlags::default()
            },
            nesting_depth: 1,
            has_separator: true,
        };
        let advance = FieldText {
            field: field.clone(),
            instruction: r#" ADVANCE \u 6 \d 12 \l 20 \r -4 \x 150 \y "72" \d -3 "#.to_string(),
            result: Some("cached placement".to_string()),
        };

        let advance_field = advance.advance_field().unwrap();
        assert_eq!(advance_field.field(), &field);
        assert_eq!(advance_field.instruction(), advance.instruction);
        let adjustments = advance_field
            .adjustments()
            .iter()
            .map(|adjustment| (adjustment.operation(), adjustment.points()))
            .collect::<Vec<_>>();
        assert_eq!(
            adjustments,
            vec![
                (AdvanceFieldOperation::Up, 6),
                (AdvanceFieldOperation::Down, 12),
                (AdvanceFieldOperation::Left, 20),
                (AdvanceFieldOperation::Right, -4),
                (AdvanceFieldOperation::HorizontalPosition, 150),
                (AdvanceFieldOperation::VerticalPosition, 72),
                (AdvanceFieldOperation::Down, -3),
            ]
        );
        assert_eq!(advance_field.cached_result(), Some("cached placement"));
        assert!(advance_field.is_dirty());
        assert!(advance_field.is_locked());

        let no_adjustments = FieldText {
            instruction: "aDvAnCe".to_string(),
            result: None,
            ..advance.clone()
        };
        assert!(
            no_adjustments
                .advance_field()
                .unwrap()
                .adjustments()
                .is_empty()
        );

        for instruction in [
            r#"ADVANCE \d"#,
            r#"ADVANCE \z 10"#,
            r#"ADVANCE \x 1.5"#,
            r#"ADVANCE \u 9223372036854775808"#,
            "ADVANCE 12",
            r#"ADVANCE \d 6 trailing"#,
        ] {
            let malformed = FieldText {
                instruction: instruction.to_string(),
                result: None,
                ..advance.clone()
            };
            assert!(malformed.advance_field().is_none(), "{instruction}");
        }

        let wrong_keyword = FieldText {
            instruction: r#"ADVANCER \u 6"#.to_string(),
            result: None,
            ..advance.clone()
        };
        assert!(wrong_keyword.advance_field().is_none());

        let wrong_type = FieldText {
            field: Field {
                field_type: FieldType::UserAddress,
                ..field
            },
            result: None,
            ..advance
        };
        assert!(wrong_type.advance_field().is_none());
    }

    #[test]
    fn recipient_fields_expose_layout_metadata_without_merging() {
        let field = Field {
            story: FieldStory::Textbox,
            start_cp: 4,
            separator_cp: Some(9),
            end_cp: 22,
            field_type: FieldType::AddressBlock,
            end_flags: FieldEndFlags {
                results_dirty: true,
                locked: true,
                has_separator: true,
                ..FieldEndFlags::default()
            },
            nesting_depth: 1,
            has_separator: true,
        };
        let address = FieldText {
            field: field.clone(),
            instruction: r#" ADDRESSBLOCK \c 2 \d \e "United States" \e Canada \f "<<_FIRST0_>> <<_LAST0_>>" \l 1033 \* MERGEFORMAT "#
                .to_string(),
            result: Some("cached address".to_string()),
        };

        let address = address.mail_merge_recipient_field().unwrap();
        assert_eq!(address.field(), &field);
        assert_eq!(address.kind(), MailMergeRecipientFieldKind::AddressBlock);
        assert_eq!(
            address.country_inclusion(),
            Some(AddressBlockCountryInclusion::UnlessExcluded)
        );
        assert!(address.formats_using_recipient_country());
        let excluded = address
            .excluded_countries()
            .iter()
            .map(String::as_str)
            .collect::<Vec<_>>();
        assert_eq!(excluded, vec!["United States", "Canada"]);
        assert_eq!(address.format_template(), Some("<<_FIRST0_>> <<_LAST0_>>"));
        assert_eq!(address.language(), Some("1033"));
        assert_eq!(address.greeting_fallback_text(), None);
        assert_eq!(address.unknown_switches().len(), 1);
        assert_eq!(address.unknown_switches()[0].name(), '*');
        assert_eq!(
            address.unknown_switches()[0].argument(),
            Some("MERGEFORMAT")
        );
        assert_eq!(address.cached_result(), Some("cached address"));
        assert!(address.is_dirty());
        assert!(address.is_locked());

        let greeting = FieldText {
            field: Field {
                field_type: FieldType::GreetingLine,
                ..field.clone()
            },
            instruction:
                r#"greetingline \f "Dear <<_FIRST0_>>," \e "To Whom It May Concern" \l en-US"#
                    .to_string(),
            result: Some("Dear Ada,".to_string()),
        };
        let greeting = greeting.mail_merge_recipient_field().unwrap();
        assert_eq!(greeting.kind(), MailMergeRecipientFieldKind::GreetingLine);
        assert_eq!(greeting.country_inclusion(), None);
        assert!(!greeting.formats_using_recipient_country());
        assert!(greeting.excluded_countries().is_empty());
        assert_eq!(greeting.format_template(), Some("Dear <<_FIRST0_>>,"));
        assert_eq!(greeting.language(), Some("en-US"));
        assert_eq!(
            greeting.greeting_fallback_text(),
            Some("To Whom It May Concern")
        );
        assert_eq!(greeting.cached_result(), Some("Dear Ada,"));

        for instruction in [
            "ADDRESSBLOCK text",
            r"ADDRESSBLOCK \c",
            r"ADDRESSBLOCK \c 3",
            r"ADDRESSBLOCK \d 1",
            r"ADDRESSBLOCK \d \d",
            r"ADDRESSBLOCK \f",
            r#"GREETINGLINE \f "Dear" \f "Hello""#,
            r"GREETINGLINE \l",
            r#"GREETINGLINE \c "First" \e "Second""#,
        ] {
            let malformed = FieldText {
                field: field.clone(),
                instruction: instruction.to_string(),
                result: None,
            };
            assert!(
                malformed.mail_merge_recipient_field().is_none(),
                "{instruction}"
            );
        }

        let wrong_keyword = FieldText {
            instruction: r"ADDRESSBLOCKING \c 1".to_string(),
            field: field.clone(),
            result: None,
        };
        assert!(wrong_keyword.mail_merge_recipient_field().is_none());

        let wrong_type = FieldText {
            field: Field {
                field_type: FieldType::GreetingLine,
                ..field
            },
            instruction: r"ADDRESSBLOCK \c 1".to_string(),
            result: None,
        };
        assert!(wrong_type.mail_merge_recipient_field().is_none());
    }

    #[test]
    fn field_text_rejects_unrepresentable_or_reversed_ranges() {
        let overflow = Field {
            story: FieldStory::Main,
            start_cp: u32::MAX,
            separator_cp: None,
            end_cp: u32::MAX,
            field_type: FieldType::If,
            end_flags: FieldEndFlags::default(),
            nesting_depth: 0,
            has_separator: false,
        };
        assert!(FieldText::from_field(&overflow, |_, _| Ok(String::new())).is_err());

        let reversed = Field {
            start_cp: 8,
            end_cp: 8,
            ..overflow
        };
        assert!(FieldText::from_field(&reversed, |_, _| Ok(String::new())).is_err());
    }

    #[test]
    fn fields_table_extracts_text_from_each_field_story() {
        let main = Field {
            story: FieldStory::Main,
            start_cp: 0,
            separator_cp: None,
            end_cp: 4,
            field_type: FieldType::Date,
            end_flags: FieldEndFlags::default(),
            nesting_depth: 0,
            has_separator: false,
        };
        let header = Field {
            story: FieldStory::Header,
            start_cp: 2,
            separator_cp: Some(7),
            end_cp: 9,
            field_type: FieldType::IncludeText,
            end_flags: FieldEndFlags {
                has_separator: true,
                ..FieldEndFlags::default()
            },
            nesting_depth: 0,
            has_separator: true,
        };
        let table = FieldsTable {
            stories: vec![
                FieldStoryTable {
                    story: FieldStory::Main,
                    markers: Vec::new(),
                    terminal_cp: 4,
                    fields: vec![main],
                },
                FieldStoryTable {
                    story: FieldStory::Header,
                    markers: Vec::new(),
                    terminal_cp: 9,
                    fields: vec![header],
                },
            ],
        };

        let text = table
            .field_texts(|story, start, end| {
                Ok(match (story, start, end) {
                    (FieldStory::Main, 1, 4) => " DATE ".to_string(),
                    (FieldStory::Header, 3, 7) => r#" INCLUDETEXT "draft.doc" "#.to_string(),
                    (FieldStory::Header, 8, 9) => "cached".to_string(),
                    _ => return Err(corrupted("unexpected field story range")),
                })
            })
            .unwrap();

        assert_eq!(text.len(), 2);
        assert_eq!(text[0].instruction, " DATE ");
        assert_eq!(text[1].instruction, r#" INCLUDETEXT "draft.doc" "#);
        assert_eq!(text[1].result.as_deref(), Some("cached"));
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
