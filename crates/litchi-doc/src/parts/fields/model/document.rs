use super::core::{Field, FieldType};
use super::mail_merge::MergeFieldSwitch;

/// A typed, inert legacy Word `DOCVARIABLE` field.
///
/// [MS-DOC] §2.9.90 identifies its native field-type byte, and ECMA-376 Part
/// 1 §17.16.5.15 defines `DOCVARIABLE` with one document-variable name.
/// This type exposes the stored name, any preserved switches, and cached result
/// only. It never reads document variables, resolves a value, or refreshes a
/// field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DocumentVariableField {
    pub(in crate::parts::fields) field: Field,
    pub(in crate::parts::fields) instruction: String,
    pub(in crate::parts::fields) variable_name: String,
    pub(in crate::parts::fields) unknown_switches: Vec<MergeFieldSwitch>,
    pub(in crate::parts::fields) cached_result: Option<String>,
}

impl DocumentVariableField {
    /// Return the paired field markers and their story-relative positions.
    #[must_use]
    pub fn field(&self) -> &Field {
        &self.field
    }

    /// Return the complete stored field instruction.
    #[must_use]
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the stored document-variable name without resolving it.
    #[must_use]
    pub fn variable_name(&self) -> &str {
        &self.variable_name
    }

    /// Return preserved switches in source order without interpreting them.
    ///
    /// `DOCVARIABLE` has no field-specific switches. These values remain
    /// inert source metadata and are never applied.
    #[must_use]
    pub fn unknown_switches(&self) -> &[MergeFieldSwitch] {
        &self.unknown_switches
    }

    /// Return the stored cached field result, if present.
    ///
    /// This value is never regenerated from a document variable.
    #[must_use]
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a producer marked the stored result stale.
    #[must_use]
    pub fn is_dirty(&self) -> bool {
        self.field.end_flags.results_dirty
    }

    /// Whether a producer locked this field against refresh.
    #[must_use]
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
    pub(in crate::parts::fields) field: Field,
    pub(in crate::parts::fields) instruction: String,
    pub(in crate::parts::fields) property_name: String,
    pub(in crate::parts::fields) switches: Vec<MergeFieldSwitch>,
    pub(in crate::parts::fields) cached_result: Option<String>,
}

impl DocumentPropertyField {
    /// Return the paired field markers and their story-relative positions.
    #[must_use]
    pub fn field(&self) -> &Field {
        &self.field
    }

    /// Return the complete stored field instruction.
    #[must_use]
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the stored document-property name without resolving it.
    #[must_use]
    pub fn property_name(&self) -> &str {
        &self.property_name
    }

    /// Return preserved switches in source order without interpreting them.
    ///
    /// These values remain inert source metadata and are never applied.
    #[must_use]
    pub fn switches(&self) -> &[MergeFieldSwitch] {
        &self.switches
    }

    /// Return the stored cached field result, if present.
    ///
    /// This value is never regenerated from a document property.
    #[must_use]
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a producer marked the stored result stale.
    #[must_use]
    pub fn is_dirty(&self) -> bool {
        self.field.end_flags.results_dirty
    }

    /// Whether a producer locked this field against refresh.
    #[must_use]
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
    pub(in crate::parts::fields) field: Field,
    pub(in crate::parts::fields) instruction: String,
    pub(in crate::parts::fields) information_type: String,
    pub(in crate::parts::fields) new_value: Option<String>,
    pub(in crate::parts::fields) switches: Vec<MergeFieldSwitch>,
    pub(in crate::parts::fields) cached_result: Option<String>,
}

impl InfoField {
    /// Return the paired field markers and their story-relative positions.
    #[must_use]
    pub fn field(&self) -> &Field {
        &self.field
    }

    /// Return the complete stored field instruction.
    #[must_use]
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the stored document or template property selector.
    ///
    /// The selector is preserved as metadata and is never looked up.
    #[must_use]
    pub fn information_type(&self) -> &str {
        &self.information_type
    }

    /// Return the stored optional replacement value.
    ///
    /// This value is never applied to a document or template property.
    #[must_use]
    pub fn new_value(&self) -> Option<&str> {
        self.new_value.as_deref()
    }

    /// Return preserved switches in source order without interpreting them.
    ///
    /// These values remain inert source metadata and are never applied.
    #[must_use]
    pub fn switches(&self) -> &[MergeFieldSwitch] {
        &self.switches
    }

    /// Return the stored cached field result, if present.
    ///
    /// This value is never regenerated from a property.
    #[must_use]
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a producer marked the stored result stale.
    #[must_use]
    pub fn is_dirty(&self) -> bool {
        self.field.end_flags.results_dirty
    }

    /// Whether a producer locked this field against refresh.
    #[must_use]
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
    #[must_use]
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

    pub(in crate::parts::fields) fn from_field_type(field_type: FieldType) -> Option<Self> {
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

    pub(in crate::parts::fields) fn from_keyword(keyword: &str) -> Option<Self> {
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
    pub(in crate::parts::fields) field: Field,
    pub(in crate::parts::fields) instruction: String,
    pub(in crate::parts::fields) kind: DocumentInformationFieldKind,
    pub(in crate::parts::fields) switches: Vec<MergeFieldSwitch>,
    pub(in crate::parts::fields) cached_result: Option<String>,
}

impl DocumentInformationField {
    /// Return the paired field markers and their story-relative positions.
    #[must_use]
    pub fn field(&self) -> &Field {
        &self.field
    }

    /// Return the complete stored field instruction.
    #[must_use]
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the recognized built-in document-information category.
    #[must_use]
    pub const fn kind(&self) -> DocumentInformationFieldKind {
        self.kind
    }

    /// Return preserved switches in source order without interpreting them.
    ///
    /// These values remain inert source metadata and are never applied.
    #[must_use]
    pub fn switches(&self) -> &[MergeFieldSwitch] {
        &self.switches
    }

    /// Return the stored cached field result, if present.
    ///
    /// This value is never regenerated from document metadata or a host user
    /// profile.
    #[must_use]
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a producer marked the stored result stale.
    #[must_use]
    pub fn is_dirty(&self) -> bool {
        self.field.end_flags.results_dirty
    }

    /// Whether a producer locked this field against refresh.
    #[must_use]
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
    #[must_use]
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

    pub(in crate::parts::fields) fn from_field_type(field_type: FieldType) -> Option<Self> {
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

    pub(in crate::parts::fields) fn from_keyword(keyword: &str) -> Option<Self> {
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
    pub(in crate::parts::fields) field: Field,
    pub(in crate::parts::fields) instruction: String,
    pub(in crate::parts::fields) kind: DocumentContextFieldKind,
    pub(in crate::parts::fields) switches: Vec<MergeFieldSwitch>,
    pub(in crate::parts::fields) cached_result: Option<String>,
}

impl DocumentContextField {
    /// Return the paired field markers and their story-relative positions.
    #[must_use]
    pub fn field(&self) -> &Field {
        &self.field
    }

    /// Return the complete stored field instruction.
    #[must_use]
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the recognized built-in document-context or runtime category.
    #[must_use]
    pub const fn kind(&self) -> DocumentContextFieldKind {
        self.kind
    }

    /// Return preserved switches in source order without interpreting them.
    ///
    /// These values remain inert source metadata and are never applied.
    #[must_use]
    pub fn switches(&self) -> &[MergeFieldSwitch] {
        &self.switches
    }

    /// Return the stored cached field result, if present.
    ///
    /// This value is never regenerated from a document path, attached
    /// template, host filesystem state or file size, current clock, or page
    /// and section layout.
    #[must_use]
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a producer marked the stored result stale.
    #[must_use]
    pub fn is_dirty(&self) -> bool {
        self.field.end_flags.results_dirty
    }

    /// Whether a producer locked this field against refresh.
    #[must_use]
    pub fn is_locked(&self) -> bool {
        self.field.end_flags.locked
    }
}
