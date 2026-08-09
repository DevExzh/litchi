#![expect(
    clippy::shadow_reuse,
    reason = "parser bindings are intentionally refined after validation"
)]
#![expect(
    clippy::struct_excessive_bools,
    reason = "the public model preserves independent OOXML flags"
)]
//! Legacy external, linked, and referenced-document field models.

use super::{Field, Switch};

use crate::error::{Error, Result};

use super::super::codec::{
    field_instruction_remainder, parse_dde_operands_and_switches,
    parse_external_include_operands_and_switches, parse_field_operand_and_switches,
    parse_link_operands_and_switches, required_external_include_option_argument,
};

/// The stored kind of a legacy DDE field.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DdeKind {
    /// A DDE field, which can request automatic updates with its a switch.
    Dde,
    /// A DDEAUTO field, which declares automatic updates.
    DdeAuto,
}

/// One stored DDE result representation switch.
///
/// This value describes a requested representation only. It never causes a
/// source to be contacted, converted, embedded, or displayed.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DdeFormat {
    /// The b switch requests a bitmap representation.
    Bitmap,
    /// The h switch requests HTML-formatted text.
    Html,
    /// The p switch requests a picture representation.
    Picture,
    /// The r switch requests rich-text format.
    RichText,
    /// The t switch requests text-only format.
    Text,
    /// The u switch requests Unicode text.
    UnicodeText,
}

/// Typed, inert metadata for a legacy DDE or DDEAUTO field.
///
/// Application, source, item, representation, and storage switches are
/// retained as stored field data. This type never launches an application,
/// initiates a DDE conversation, opens a source, requests data, refreshes
/// content, converts content, or executes code.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Dde {
    instruction: String,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
    kind: DdeKind,
    application: String,
    source: String,
    item: Option<String>,
    automatic_updates: bool,
    representation: Option<DdeFormat>,
    omit_graphic_data: bool,
    switches: Vec<Switch>,
}

impl Dde {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        let Some((kind, application, source, item, switches)) =
            parse_dde_operands_and_switches(field.instruction())?
        else {
            return Ok(None);
        };

        let mut automatic_updates = kind == DdeKind::DdeAuto;
        let mut saw_automatic_update = false;
        let mut representation = None;
        let mut omit_graphic_data = false;
        for switch in &switches {
            match switch.name {
                'a' if kind == DdeKind::Dde => {
                    if saw_automatic_update || switch.argument.is_some() {
                        return Err(Error::Invalid(
                            "DDE \\a switch cannot be repeated or take an argument".to_string(),
                        ));
                    }
                    automatic_updates = true;
                    saw_automatic_update = true;
                },
                'a' => {
                    return Err(Error::Invalid(
                        "DDEAUTO field does not allow a \\a switch".to_string(),
                    ));
                },
                'd' => {
                    if representation.is_some() || omit_graphic_data || switch.argument.is_some() {
                        return Err(Error::Invalid(
                            "DDE result and storage switches cannot be combined".to_string(),
                        ));
                    }
                    omit_graphic_data = true;
                },
                'b' | 'h' | 'p' | 'r' | 't' | 'u' => {
                    if representation.is_some() || omit_graphic_data || switch.argument.is_some() {
                        return Err(Error::Invalid(
                            "DDE result and storage switches cannot be combined".to_string(),
                        ));
                    }
                    representation = Some(match switch.name {
                        'b' => DdeFormat::Bitmap,
                        'h' => DdeFormat::Html,
                        'p' => DdeFormat::Picture,
                        'r' => DdeFormat::RichText,
                        't' => DdeFormat::Text,
                        'u' => DdeFormat::UnicodeText,
                        _ => unreachable!("DDE representation switch was matched above"),
                    });
                },
                _ => {},
            }
        }

        Ok(Some(Self {
            instruction: field.instruction.clone(),
            cached_result: field.result.clone(),
            dirty: field.dirty,
            locked: field.locked,
            kind,
            application,
            source,
            item,
            automatic_updates,
            representation,
            omit_graphic_data,
            switches,
        }))
    }

    /// Return the complete stored field instruction.
    #[must_use]
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the cached field result, if one was stored.
    #[must_use]
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a word processor marked the cached result stale.
    #[must_use]
    pub fn is_dirty(&self) -> bool {
        self.dirty
    }

    /// Whether a word processor locked this field against refresh.
    #[must_use]
    pub fn is_locked(&self) -> bool {
        self.locked
    }

    /// Return whether this is a DDE or DDEAUTO field.
    #[must_use]
    pub fn kind(&self) -> DdeKind {
        self.kind
    }

    /// Return the stored DDE application name without launching it.
    #[must_use]
    pub fn application(&self) -> &str {
        &self.application
    }

    /// Return the stored source identifier without opening or resolving it.
    #[must_use]
    pub fn source(&self) -> &str {
        &self.source
    }

    /// Return the optional stored source item, such as a cell range or bookmark.
    #[must_use]
    pub fn item(&self) -> Option<&str> {
        self.item.as_deref()
    }

    /// Whether the stored instruction requests automatic DDE updates.
    ///
    /// This is metadata only. The API never performs an update.
    #[must_use]
    pub fn requests_automatic_updates(&self) -> bool {
        self.automatic_updates
    }

    /// Return the requested stored result representation, if present.
    ///
    /// This is metadata only and never triggers source access or conversion.
    #[must_use]
    pub fn representation(&self) -> Option<DdeFormat> {
        self.representation
    }

    /// Whether the stored d switch omits graphic data from the document.
    ///
    /// This is stored metadata only. The API never reads the source to obtain
    /// omitted data.
    #[must_use]
    pub fn omits_graphic_data(&self) -> bool {
        self.omit_graphic_data
    }

    /// Return all stored field switches in source order.
    #[must_use]
    pub fn switches(&self) -> &[Switch] {
        &self.switches
    }
}

/// The kind of externally sourced Word field.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum IncludeKind {
    /// An `INCLUDETEXT` or historical `INCLUDE` field that stores a document or XML
    /// source.
    Text,
    /// An `INCLUDEPICTURE` or historical `IMPORT` field that stores an image source.
    Picture,
}

/// One recognized stored option of an external-include field.
///
/// These values are configuration metadata only. This API never opens,
/// resolves, imports, transforms, or evaluates the referenced source.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum IncludeOption {
    /// A document or graphics converter name from the c switch.
    Converter(String),
    /// A source encoding from the INCLUDETEXT e switch.
    Encoding(String),
    /// A source MIME type from the INCLUDETEXT m switch.
    MimeType(String),
    /// An XML namespace mapping from the INCLUDETEXT n switch.
    NamespaceMapping(String),
    /// An XSLT location from the INCLUDETEXT t switch.
    Xslt(String),
    /// An `XPath` expression from the INCLUDETEXT x switch.
    XPath(String),
}

/// Typed, inert metadata for an `INCLUDETEXT`/`INCLUDEPICTURE` or historical
/// `INCLUDE`/`IMPORT` field.
///
/// Source identifiers, bookmarks, options, and cached results are retained as
/// stored field data. This type never opens, resolves, imports, fetches,
/// refreshes, transforms, converts, evaluates, or executes source content.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Include {
    instruction: String,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
    kind: IncludeKind,
    source: String,
    bookmark: Option<String>,
    suppress_nested_field_updates: bool,
    omit_picture_data: bool,
    options: Vec<IncludeOption>,
    switches: Vec<Switch>,
}

impl Include {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        let Some((kind, source, bookmark, switches)) =
            parse_external_include_operands_and_switches(field.instruction())?
        else {
            return Ok(None);
        };

        let mut suppress_nested_field_updates = false;
        let mut omit_picture_data = false;
        let mut options = Vec::new();
        for switch in &switches {
            match (kind, switch.name) {
                (IncludeKind::Text, '!') => {
                    if switch.argument.is_some() {
                        return Err(Error::Invalid(
                            "INCLUDETEXT exclamation switch does not take an argument".to_string(),
                        ));
                    }
                    suppress_nested_field_updates = true;
                },
                (IncludeKind::Picture, 'd') => {
                    if switch.argument.is_some() {
                        return Err(Error::Invalid(
                            "INCLUDEPICTURE d switch does not take an argument".to_string(),
                        ));
                    }
                    omit_picture_data = true;
                },
                (_, 'c') => options.push(IncludeOption::Converter(
                    required_external_include_option_argument(switch, kind)?,
                )),
                (IncludeKind::Text, 'e') => options.push(IncludeOption::Encoding(
                    required_external_include_option_argument(switch, kind)?,
                )),
                (IncludeKind::Text, 'm') => options.push(IncludeOption::MimeType(
                    required_external_include_option_argument(switch, kind)?,
                )),
                (IncludeKind::Text, 'n') => options.push(IncludeOption::NamespaceMapping(
                    required_external_include_option_argument(switch, kind)?,
                )),
                (IncludeKind::Text, 't') => options.push(IncludeOption::Xslt(
                    required_external_include_option_argument(switch, kind)?,
                )),
                (IncludeKind::Text, 'x') => options.push(IncludeOption::XPath(
                    required_external_include_option_argument(switch, kind)?,
                )),
                _ => {},
            }
        }

        Ok(Some(Self {
            instruction: field.instruction.clone(),
            cached_result: field.result.clone(),
            dirty: field.dirty,
            locked: field.locked,
            kind,
            source,
            bookmark,
            suppress_nested_field_updates,
            omit_picture_data,
            options,
            switches,
        }))
    }

    /// Return the complete stored field instruction.
    #[must_use]
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the cached field result, if one was stored.
    #[must_use]
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a word processor marked the cached result stale.
    #[must_use]
    pub fn is_dirty(&self) -> bool {
        self.dirty
    }

    /// Whether a word processor locked this field against refresh.
    #[must_use]
    pub fn is_locked(&self) -> bool {
        self.locked
    }

    /// Return whether this includes text or a picture.
    ///
    /// Text includes use `INCLUDETEXT` or historical `INCLUDE`; picture includes use
    /// `INCLUDEPICTURE` or historical `IMPORT`.
    #[must_use]
    pub fn kind(&self) -> IncludeKind {
        self.kind
    }

    /// Return the stored source identifier without opening or resolving it.
    #[must_use]
    pub fn source(&self) -> &str {
        &self.source
    }

    /// Return the optional stored bookmark selector for a text-include field.
    ///
    /// `INCLUDEPICTURE` and `IMPORT` fields do not define a bookmark operand, so this
    /// returns None for picture includes.
    #[must_use]
    pub fn bookmark(&self) -> Option<&str> {
        self.bookmark.as_deref()
    }

    /// Whether the stored text-include instruction suppresses nested updates.
    ///
    /// This is metadata only. The API never performs an update.
    #[must_use]
    pub fn suppresses_nested_field_updates(&self) -> bool {
        self.suppress_nested_field_updates
    }

    /// Whether the stored picture-include instruction omits picture data.
    ///
    /// This is stored metadata only. The API never reads the source to obtain
    /// omitted picture data.
    #[must_use]
    pub fn omits_picture_data(&self) -> bool {
        self.omit_picture_data
    }

    /// Return recognized converter and XML options in stored source order.
    ///
    /// All options are inert metadata. This method never resolves a converter,
    /// opens a source, runs XSLT, or evaluates `XPath`.
    #[must_use]
    pub fn options(&self) -> &[IncludeOption] {
        &self.options
    }

    /// Return all stored field switches in source order.
    #[must_use]
    pub fn switches(&self) -> &[Switch] {
        &self.switches
    }
}

/// Typed, inert metadata for an RD referenced-document field.
///
/// Source identifiers, relative-path settings, switches, and cached results
/// are retained as stored field data. This type never opens, resolves, reads,
/// imports, refreshes, evaluates, or executes the referenced document.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SubDocument {
    instruction: String,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
    source: String,
    relative_path: bool,
    switches: Vec<Switch>,
}

impl SubDocument {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        let Some((source, switches)) = parse_field_operand_and_switches(field.instruction(), "RD")?
        else {
            return Ok(None);
        };
        let source = source.filter(|value| !value.is_empty()).ok_or_else(|| {
            Error::Invalid("RD field is missing its referenced document path".to_string())
        })?;

        let mut relative_path = false;
        for switch in &switches {
            if switch.name == 'f' {
                if switch.argument.is_some() {
                    return Err(Error::Invalid(
                        "RD \\\\f switch does not take an argument".to_string(),
                    ));
                }
                if relative_path {
                    return Err(Error::Invalid(
                        "RD \\\\f switch cannot be repeated".to_string(),
                    ));
                }
                relative_path = true;
            }
        }

        Ok(Some(Self {
            instruction: field.instruction.clone(),
            cached_result: field.result.clone(),
            dirty: field.dirty,
            locked: field.locked,
            source,
            relative_path,
            switches,
        }))
    }

    /// Return the complete stored field instruction.
    #[must_use]
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the cached field result, if one was stored.
    #[must_use]
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a word processor marked the cached result stale.
    #[must_use]
    pub fn is_dirty(&self) -> bool {
        self.dirty
    }

    /// Whether a word processor locked this field against refresh.
    #[must_use]
    pub fn is_locked(&self) -> bool {
        self.locked
    }

    /// Return the stored referenced-document path without opening it.
    #[must_use]
    pub fn source(&self) -> &str {
        &self.source
    }

    /// Whether the stored RD instruction's `\\f` switch requests a path relative to this
    /// document.
    ///
    /// This is metadata only. The API never resolves the path.
    #[must_use]
    pub fn uses_relative_path(&self) -> bool {
        self.relative_path
    }

    /// Return all stored field switches in source order.
    #[must_use]
    pub fn switches(&self) -> &[Switch] {
        &self.switches
    }
}

/// One stored result or storage switch for a Word `LINK` field.
///
/// These values describe a linked-object representation or whether graphic data
/// is stored. They never cause a source to be opened, contacted, converted, or
/// displayed.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LinkResult {
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
/// ECMA-376 marks modes `1` and `3` unsupported. Those values, and values
/// outside its defined set, are retained as metadata without applying any
/// formatting.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LinkFormat {
    /// `0`: preserve formatting from the source file.
    Source,
    /// `2`: match formatting in the destination document.
    Destination,
    /// `4`: preserve source formatting for a `SpreadsheetML` workbook source.
    SpreadsheetSource,
    /// `5`: match destination formatting for a `SpreadsheetML` workbook source.
    SpreadsheetDestination,
    /// An ECMA-376-unsupported or otherwise unrecognized integral mode.
    Unsupported(i64),
}

/// Typed, inert metadata for a legacy Word `LINK` field.
///
/// Application type, source, item, and all result/formatting switches are
/// retained as stored field data. This type never activates an OLE server,
/// launches an application, opens a source, requests data, refreshes content,
/// converts content, or executes code.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Link {
    instruction: String,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
    application_type: String,
    source: String,
    item: Option<String>,
    automatic_updates: bool,
    result_options: Vec<LinkResult>,
    formatting_modes: Vec<LinkFormat>,
    switches: Vec<Switch>,
}

impl Link {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        let Some((application_type, source, item, switches)) =
            parse_link_operands_and_switches(field.instruction())?
        else {
            return Ok(None);
        };

        let mut automatic_updates = false;
        let mut result_options = Vec::new();
        let mut formatting_modes = Vec::new();
        for switch in &switches {
            match switch.name {
                'a' => {
                    if switch.argument.is_some() {
                        return Err(Error::Invalid(
                            "LINK \\a switch does not take an argument".to_string(),
                        ));
                    }
                    automatic_updates = true;
                },
                'f' => {
                    let argument = switch.argument.as_deref().ok_or_else(|| {
                        Error::Invalid(
                            "LINK \\f switch requires an integral formatting mode".to_string(),
                        )
                    })?;
                    let value = argument.parse::<i64>().map_err(|_source_error| {
                        Error::Invalid("LINK \\f formatting mode must be an integer".to_string())
                    })?;
                    formatting_modes.push(match value {
                        0 => LinkFormat::Source,
                        2 => LinkFormat::Destination,
                        4 => LinkFormat::SpreadsheetSource,
                        5 => LinkFormat::SpreadsheetDestination,
                        other => LinkFormat::Unsupported(other),
                    });
                },
                'b' | 'd' | 'h' | 'p' | 'r' | 't' | 'u' => {
                    if switch.argument.is_some() {
                        return Err(Error::Invalid(format!(
                            "LINK \\{} switch does not take an argument",
                            switch.name
                        )));
                    }
                    result_options.push(match switch.name {
                        'b' => LinkResult::Bitmap,
                        'd' => LinkResult::OmitGraphicData,
                        'h' => LinkResult::Html,
                        'p' => LinkResult::Picture,
                        'r' => LinkResult::RichText,
                        't' => LinkResult::Text,
                        'u' => LinkResult::UnicodeText,
                        _ => unreachable!("LINK result switch was matched above"),
                    });
                },
                _ => {},
            }
        }

        Ok(Some(Self {
            instruction: field.instruction.clone(),
            cached_result: field.result.clone(),
            dirty: field.dirty,
            locked: field.locked,
            application_type,
            source,
            item,
            automatic_updates,
            result_options,
            formatting_modes,
            switches,
        }))
    }

    /// Return the complete stored field instruction.
    #[must_use]
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the cached field result, if one was stored.
    #[must_use]
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a word processor marked the cached result stale.
    #[must_use]
    pub fn is_dirty(&self) -> bool {
        self.dirty
    }

    /// Whether a word processor locked this field against refresh.
    #[must_use]
    pub fn is_locked(&self) -> bool {
        self.locked
    }

    /// Return the stored linked-object application type.
    ///
    /// Word commonly stores an OLE Programmatic Identifier here. It is never
    /// looked up or activated by this API.
    #[must_use]
    pub fn application_type(&self) -> &str {
        &self.application_type
    }

    /// Return the stored source identifier without opening or resolving it.
    #[must_use]
    pub fn source(&self) -> &str {
        &self.source
    }

    /// Return the optional stored source item, such as a cell range or bookmark.
    #[must_use]
    pub fn item(&self) -> Option<&str> {
        self.item.as_deref()
    }

    /// Whether the stored instruction requests automatic updates.
    ///
    /// This is metadata only. The API never performs an update.
    #[must_use]
    pub fn requests_automatic_updates(&self) -> bool {
        self.automatic_updates
    }

    /// Return recognized result and storage switches in stored source order.
    ///
    /// When several are present, [`Self::effective_result_option`] reflects
    /// Word's documented last-switch behavior. Neither method contacts the
    /// linked source.
    #[must_use]
    pub fn result_options(&self) -> &[LinkResult] {
        &self.result_options
    }

    /// Return the effective result or storage option under Word's documented
    /// last-switch behavior, if one was stored.
    #[must_use]
    pub fn effective_result_option(&self) -> Option<LinkResult> {
        self.result_options.last().copied()
    }

    /// Return integral `\\f` formatting modes in stored source order.
    ///
    /// These are metadata only; this API never formats linked content.
    #[must_use]
    pub fn formatting_modes(&self) -> &[LinkFormat] {
        &self.formatting_modes
    }

    /// Return all stored field switches in source order.
    #[must_use]
    pub fn switches(&self) -> &[Switch] {
        &self.switches
    }
}

impl Field {
    /// Check whether this is a legacy `LINK` field.
    ///
    /// Recognition is limited to the stored field instruction. It never
    /// activates an OLE server, opens a source, or refreshes the field.
    #[must_use]
    pub fn is_link(&self) -> bool {
        field_instruction_remainder(&self.instruction, "LINK").is_some()
    }

    /// Parse this field as inert typed `LINK` metadata.
    ///
    /// Returns `Ok(None)` for non-`LINK` fields. The result exposes stored
    /// application, source, item, result, formatting, and cached metadata only;
    /// it never activates, opens, contacts, converts, evaluates, or executes
    /// anything.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn link(&self) -> Result<Option<Link>> {
        Link::from_field(self)
    }

    /// Check whether this is a legacy DDE field.
    ///
    /// Recognition is limited to the stored field instruction. It never
    /// launches an application, initiates a DDE conversation, opens a source,
    /// or refreshes the field.
    #[must_use]
    pub fn is_dde(&self) -> bool {
        field_instruction_remainder(&self.instruction, "DDE").is_some()
    }

    /// Check whether this is a legacy automatically updating DDEAUTO field.
    ///
    /// Recognition is limited to the stored field instruction. It never
    /// launches an application, initiates a DDE conversation, opens a source,
    /// or refreshes the field.
    #[must_use]
    pub fn is_dde_auto(&self) -> bool {
        field_instruction_remainder(&self.instruction, "DDEAUTO").is_some()
    }

    /// Parse this field as inert typed DDE or DDEAUTO metadata.
    ///
    /// Returns Ok(None) for other fields. The result exposes stored
    /// application, source, item, representation, and cached metadata only; it
    /// never launches an application, initiates a DDE conversation, opens,
    /// contacts, refreshes, converts, evaluates, or executes anything.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn dde_link(&self) -> Result<Option<Dde>> {
        Dde::from_field(self)
    }

    /// Check whether this is an `INCLUDETEXT` or historical `INCLUDE` field.
    ///
    /// Recognition is limited to the stored field instruction. It never opens,
    /// resolves, imports, fetches, or refreshes the referenced source.
    #[must_use]
    pub fn is_include_text(&self) -> bool {
        field_instruction_remainder(&self.instruction, "INCLUDETEXT").is_some()
            || field_instruction_remainder(&self.instruction, "INCLUDE").is_some()
    }

    /// Check whether this is an `INCLUDEPICTURE` or historical `IMPORT` field.
    ///
    /// Recognition is limited to the stored field instruction. It never opens,
    /// resolves, imports, fetches, or refreshes the referenced source.
    #[must_use]
    pub fn is_include_picture(&self) -> bool {
        field_instruction_remainder(&self.instruction, "INCLUDEPICTURE").is_some()
            || field_instruction_remainder(&self.instruction, "IMPORT").is_some()
    }

    /// Parse this field as inert external-include metadata.
    ///
    /// Returns Ok(None) for fields other than `INCLUDETEXT`/`INCLUDEPICTURE` or
    /// their historical `INCLUDE`/`IMPORT` aliases. The result exposes stored
    /// source, bookmark, converter, XML, and cached metadata only; it never
    /// opens, resolves, imports, fetches, refreshes, converts, evaluates, or
    /// executes anything.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn external_include(&self) -> Result<Option<Include>> {
        Include::from_field(self)
    }

    /// Check whether this is an RD referenced-document field.
    ///
    /// Recognition is limited to the stored field instruction. It never opens,
    /// resolves, reads, imports, or refreshes the referenced document.
    #[must_use]
    pub fn is_referenced_document(&self) -> bool {
        field_instruction_remainder(&self.instruction, "RD").is_some()
    }

    /// Parse this field as inert referenced-document metadata.
    ///
    /// Returns Ok(None) for non-RD fields. The result exposes only the stored
    /// path, relative-path request, switches, cached content, and dirty/lock
    /// state; it never opens, resolves, imports, evaluates, or executes
    /// anything.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn referenced_document(&self) -> Result<Option<SubDocument>> {
        SubDocument::from_field(self)
    }
}
