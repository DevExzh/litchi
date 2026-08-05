//! Document-variable, document-property, and document-context field models.

use super::{Field, Switch};

use crate::error::{Error, Result};

use super::super::codec::{
    field_instruction_remainder, has_field_switch, parse_field_operand_and_switches,
    parse_field_switches, parse_info_field_parts,
};

use super::super::{
    MAX_DOCUMENT_CONTEXT_FIELD_INSTRUCTION_BYTES, MAX_DOCUMENT_INFORMATION_FIELD_INSTRUCTION_BYTES,
    MAX_DOCUMENT_PROPERTY_FIELD_INSTRUCTION_BYTES,
};

/// A typed, inert Word `DOCVARIABLE` field.
///
/// This preserves a stored variable name, field switches, and cached result.
/// It never reads a document's settings XML, resolves a variable value, or
/// refreshes the field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Variable {
    instruction: String,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
    variable_name: String,
    switches: Vec<Switch>,
}

impl Variable {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        let Some((variable_name, switches)) =
            parse_field_operand_and_switches(field.instruction(), "DOCVARIABLE")?
        else {
            return Ok(None);
        };
        let variable_name = variable_name.ok_or_else(|| {
            Error::Invalid("DOCVARIABLE field is missing its variable name".to_string())
        })?;
        if variable_name.is_empty() {
            return Err(Error::Invalid(
                "DOCVARIABLE field variable name is empty".to_string(),
            ));
        }

        Ok(Some(Self {
            instruction: field.instruction.clone(),
            cached_result: field.result.clone(),
            dirty: field.dirty,
            locked: field.locked,
            variable_name,
            switches,
        }))
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the stored document-variable name without resolving it.
    pub fn variable_name(&self) -> &str {
        &self.variable_name
    }

    /// Return the cached visible field result, if present.
    ///
    /// This is stored text only and is never regenerated from a variable.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a word processor has marked the cached result stale.
    pub fn is_dirty(&self) -> bool {
        self.dirty
    }

    /// Whether a word processor has locked this field against refresh.
    pub fn is_locked(&self) -> bool {
        self.locked
    }

    /// Return the field switches in source order.
    ///
    /// DOCVARIABLE has no field-specific switches. Preserved switches are
    /// inert source metadata and are never interpreted.
    pub fn switches(&self) -> &[Switch] {
        &self.switches
    }

    /// Check whether a case-insensitive ASCII switch appears in this field.
    pub fn has_switch(&self, name: char) -> bool {
        has_field_switch(&self.switches, name)
    }
}

/// A typed, inert Word `DOCPROPERTY` field.
///
/// ECMA-376 Part 1 §17.16.5.14 defines one stored document-property name
/// followed by optional field switches. This type exposes that persisted
/// metadata and the cached result only. It never reads core, extended, or
/// custom package properties, resolves a value, or refreshes the field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Property {
    instruction: String,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
    property_name: String,
    switches: Vec<Switch>,
}

impl Property {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        if !field.is_document_property() {
            return Ok(None);
        }
        if field.instruction().len() > MAX_DOCUMENT_PROPERTY_FIELD_INSTRUCTION_BYTES {
            return Err(Error::Invalid(format!(
                "DOCPROPERTY field instruction exceeds {MAX_DOCUMENT_PROPERTY_FIELD_INSTRUCTION_BYTES} bytes"
            )));
        }
        let Some((property_name, switches)) =
            parse_field_operand_and_switches(field.instruction(), "DOCPROPERTY")?
        else {
            unreachable!("document-property recognition and parsing must agree");
        };
        let property_name = property_name.ok_or_else(|| {
            Error::Invalid("DOCPROPERTY field is missing its property name".to_string())
        })?;
        if property_name.is_empty() {
            return Err(Error::Invalid(
                "DOCPROPERTY field property name is empty".to_string(),
            ));
        }

        Ok(Some(Self {
            instruction: field.instruction.clone(),
            cached_result: field.result.clone(),
            dirty: field.dirty,
            locked: field.locked,
            property_name,
            switches,
        }))
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the stored document-property name without resolving it.
    pub fn property_name(&self) -> &str {
        &self.property_name
    }

    /// Return the cached visible field result, if present.
    ///
    /// This is stored text only and is never regenerated from a property.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a word processor has marked the cached result stale.
    pub fn is_dirty(&self) -> bool {
        self.dirty
    }

    /// Whether a word processor has locked this field against refresh.
    pub fn is_locked(&self) -> bool {
        self.locked
    }

    /// Return the field switches in source order.
    ///
    /// Preserved switches are inert source metadata and are never interpreted.
    pub fn switches(&self) -> &[Switch] {
        &self.switches
    }

    /// Check whether a case-insensitive ASCII switch appears in this field.
    pub fn has_switch(&self, name: char) -> bool {
        has_field_switch(&self.switches, name)
    }
}

/// Typed, inert metadata for an explicit Word `INFO` field.
///
/// Word permits the `INFO` keyword to be omitted, but that form overlaps
/// standalone document-information fields such as `TITLE`. This type
/// therefore recognizes the unambiguous explicit keyword only. It retains the
/// stored property selector, optional replacement value, switches, cached
/// result, and field state only. It never reads, resolves, modifies, or writes
/// document or template properties, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Info {
    instruction: String,
    information_type: String,
    new_value: Option<String>,
    switches: Vec<Switch>,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
}

impl Info {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        let Some((information_type, new_value, switches)) =
            parse_info_field_parts(field.instruction())?
        else {
            return Ok(None);
        };

        Ok(Some(Self {
            instruction: field.instruction.clone(),
            information_type,
            new_value,
            switches,
            cached_result: field.result.clone(),
            dirty: field.dirty,
            locked: field.locked,
        }))
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
    pub fn switches(&self) -> &[Switch] {
        &self.switches
    }

    /// Check whether a case-insensitive ASCII switch appears in this field.
    pub fn has_switch(&self, name: char) -> bool {
        has_field_switch(&self.switches, name)
    }

    /// Return the cached visible field result, if present.
    ///
    /// This is stored text only and is never regenerated from a property.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a word processor has marked the cached result stale.
    pub fn is_dirty(&self) -> bool {
        self.dirty
    }

    /// Whether a word processor has locked this field against refresh.
    pub fn is_locked(&self) -> bool {
        self.locked
    }
}

/// The built-in Word document-information field category.
///
/// These fields are defined in ECMA-376 Part 1 §17.16.5. This enum preserves
/// the stored field kind only; it does not resolve document metadata or
/// calculate dates, revisions, or statistics.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum InformationKind {
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

impl InformationKind {
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

    fn from_instruction(instruction: &str) -> Option<Self> {
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
        .find(|kind| field_instruction_remainder(instruction, kind.field_keyword()).is_some())
    }
}

/// Typed, inert metadata for a built-in Word document-information field.
///
/// This type retains the stored kind, field switches, cached result, and field
/// state only. It never reads package properties, reads or modifies host
/// identity data, calculates dates, revisions, or statistics, resolves a
/// value, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Information {
    instruction: String,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
    kind: InformationKind,
    switches: Vec<Switch>,
}

impl Information {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        let Some(kind) = InformationKind::from_instruction(field.instruction()) else {
            return Ok(None);
        };
        if field.instruction().len() > MAX_DOCUMENT_INFORMATION_FIELD_INSTRUCTION_BYTES {
            return Err(Error::Invalid(format!(
                "{} field instruction exceeds {MAX_DOCUMENT_INFORMATION_FIELD_INSTRUCTION_BYTES} bytes",
                kind.field_keyword()
            )));
        }
        let switches = parse_field_switches(field.instruction(), kind.field_keyword())?
            .expect("document-information recognition and parsing must agree");

        Ok(Some(Self {
            instruction: field.instruction.clone(),
            cached_result: field.result.clone(),
            dirty: field.dirty,
            locked: field.locked,
            kind,
            switches,
        }))
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the recognized built-in document-information category.
    pub const fn kind(&self) -> InformationKind {
        self.kind
    }

    /// Return the cached visible field result, if present.
    ///
    /// This is stored text only and is never regenerated from package metadata
    /// or a host user profile.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a word processor has marked the cached result stale.
    pub fn is_dirty(&self) -> bool {
        self.dirty
    }

    /// Whether a word processor has locked this field against refresh.
    pub fn is_locked(&self) -> bool {
        self.locked
    }

    /// Return the field switches in source order.
    ///
    /// Preserved switches are inert source metadata and are never interpreted.
    pub fn switches(&self) -> &[Switch] {
        &self.switches
    }

    /// Check whether a case-insensitive ASCII switch appears in this field.
    pub fn has_switch(&self, name: char) -> bool {
        has_field_switch(&self.switches, name)
    }
}

/// The built-in Word document-context and runtime field category.
///
/// `FILENAME`, `TEMPLATE`, `DATE`, `TIME`, `PAGE`, `FILESIZE`, `SECTION`, and
/// `SECTIONPAGES` are defined in ECMA-376 Part 1 §17.16.5. This enum preserves
/// the stored field kind only; it does not read a document path, attached
/// template, host filesystem state or file size, current clock, or page and
/// section layout.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ContextKind {
    FileName,
    Template,
    Date,
    Time,
    Page,
    FileSize,
    Section,
    SectionPages,
}

impl ContextKind {
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

    fn from_instruction(instruction: &str) -> Option<Self> {
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
        .find(|kind| field_instruction_remainder(instruction, kind.field_keyword()).is_some())
    }
}

/// Typed, inert metadata for a built-in Word document-context or runtime field.
///
/// This type retains the stored kind, field switches, cached result, and field
/// state only. It never reads a document path, attached template, host
/// filesystem state or file size, current clock, or page and section layout,
/// resolves a value, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Context {
    instruction: String,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
    kind: ContextKind,
    switches: Vec<Switch>,
}

impl Context {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        let Some(kind) = ContextKind::from_instruction(field.instruction()) else {
            return Ok(None);
        };
        if field.instruction().len() > MAX_DOCUMENT_CONTEXT_FIELD_INSTRUCTION_BYTES {
            return Err(Error::Invalid(format!(
                "{} field instruction exceeds {MAX_DOCUMENT_CONTEXT_FIELD_INSTRUCTION_BYTES} bytes",
                kind.field_keyword()
            )));
        }
        let switches = parse_field_switches(field.instruction(), kind.field_keyword())?
            .expect("document-context recognition and parsing must agree");

        Ok(Some(Self {
            instruction: field.instruction.clone(),
            cached_result: field.result.clone(),
            dirty: field.dirty,
            locked: field.locked,
            kind,
            switches,
        }))
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the recognized built-in document-context or runtime category.
    pub const fn kind(&self) -> ContextKind {
        self.kind
    }

    /// Return the cached visible field result, if present.
    ///
    /// This is stored text only and is never regenerated from a document path,
    /// attached template, host filesystem state or file size, current clock,
    /// or page and section layout.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a word processor has marked the cached result stale.
    pub fn is_dirty(&self) -> bool {
        self.dirty
    }

    /// Whether a word processor has locked this field against refresh.
    pub fn is_locked(&self) -> bool {
        self.locked
    }

    /// Return the field switches in source order.
    ///
    /// Preserved switches are inert source metadata and are never interpreted.
    pub fn switches(&self) -> &[Switch] {
        &self.switches
    }

    /// Check whether a case-insensitive ASCII switch appears in this field.
    pub fn has_switch(&self, name: char) -> bool {
        has_field_switch(&self.switches, name)
    }
}

impl Field {
    /// Check whether this is a `DOCVARIABLE` field.
    ///
    /// Recognition is limited to the stored field instruction. It never reads
    /// settings XML, resolves a variable value, or refreshes the field.
    pub fn is_document_variable(&self) -> bool {
        field_instruction_remainder(&self.instruction, "DOCVARIABLE").is_some()
    }

    /// Parse this field as inert typed document-variable metadata.
    ///
    /// Returns `Ok(None)` for non-`DOCVARIABLE` fields. The result exposes the
    /// stored variable name, switches, cached content, and dirty/lock state
    /// only; it never reads settings XML, resolves a value, or refreshes a
    /// field.
    pub fn document_variable(&self) -> Result<Option<Variable>> {
        Variable::from_field(self)
    }

    /// Check whether this is a `DOCPROPERTY` field.
    ///
    /// Recognition is limited to the stored field instruction. It never reads
    /// a package property, resolves a value, or refreshes the field.
    pub fn is_document_property(&self) -> bool {
        field_instruction_remainder(&self.instruction, "DOCPROPERTY").is_some()
    }

    /// Parse this field as inert typed document-property metadata.
    ///
    /// Returns `Ok(None)` for non-`DOCPROPERTY` fields. The result exposes the
    /// stored property name, switches, cached content, and dirty/lock state
    /// only; it never reads core, extended, or custom package properties,
    /// resolves a value, or refreshes a field.
    pub fn document_property(&self) -> Result<Option<Property>> {
        Property::from_field(self)
    }

    /// Check whether this is an explicit legacy `INFO` field.
    ///
    /// Recognition is limited to the stored field instruction. It never reads,
    /// resolves, modifies, or writes document or template properties, or
    /// refreshes the field.
    pub fn is_info_field(&self) -> bool {
        field_instruction_remainder(&self.instruction, "INFO").is_some()
    }

    /// Parse this field as inert typed legacy `INFO` metadata.
    ///
    /// Returns `Ok(None)` for fields other than an explicit `INFO` field. The
    /// result exposes the stored property selector, optional replacement value,
    /// switches, cached content, and dirty/lock state only; it never reads,
    /// resolves, modifies, or writes document or template properties, or
    /// refreshes a field.
    pub fn info_field(&self) -> Result<Option<Info>> {
        Info::from_field(self)
    }

    /// Check whether this is a built-in document-information field.
    ///
    /// Recognition is limited to the stored field instruction. It never reads
    /// package metadata or host identity data, resolves a value, or refreshes
    /// the field.
    pub fn is_document_information(&self) -> bool {
        InformationKind::from_instruction(&self.instruction).is_some()
    }

    /// Parse this field as inert typed document-information metadata.
    ///
    /// Returns `Ok(None)` for fields outside the built-in document-information
    /// family. The result exposes only the stored kind, switches, cached
    /// content, and dirty/lock state; it never reads core or extended package
    /// properties, reads or modifies host identity data, calculates dates,
    /// revisions, or statistics, resolves a value, or refreshes a field.
    pub fn document_information(&self) -> Result<Option<Information>> {
        Information::from_field(self)
    }

    /// Check whether this is a built-in document-context or runtime field.
    ///
    /// Recognition is limited to the stored field instruction. It never reads
    /// a document path, attached template, host filesystem state or file size,
    /// current clock, or page and section layout, resolves a value, or refreshes
    /// the field.
    pub fn is_document_context(&self) -> bool {
        ContextKind::from_instruction(&self.instruction).is_some()
    }

    /// Parse this field as inert typed document-context or runtime metadata.
    ///
    /// Returns `Ok(None)` for fields outside the `FILENAME`, `TEMPLATE`, `DATE`,
    /// `TIME`, `PAGE`, `FILESIZE`, `SECTION`, and `SECTIONPAGES` family. The
    /// result exposes only the stored kind, switches, cached content, and
    /// dirty/lock state; it never reads a document path, attached template,
    /// host filesystem state or file size, current clock, or page and section
    /// layout, resolves a value, or refreshes a field.
    pub fn document_context(&self) -> Result<Option<Context>> {
        Context::from_field(self)
    }
}
