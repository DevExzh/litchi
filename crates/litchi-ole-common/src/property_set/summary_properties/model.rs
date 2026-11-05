//! Semantic document properties for a compound file being created.

use super::super::codec::validate_section;
use super::super::model::{
    CodePage, DOCUMENT_SUMMARY_INFORMATION_FMTID, Guid, MAX_TYPED_TEXT_BYTES, PID_CODEPAGE,
    PID_DICTIONARY, SUMMARY_INFORMATION_FMTID, Section, Stream, USER_DEFINED_PROPERTIES_FMTID,
    Value, invalid, try_clone_property_set, validate_property_name,
};
use super::super::summary_information::FileTime;
use super::super::{document_summary, summary_information};
use super::names::{self, DocumentSummaryText, Name, SummaryText, SummaryTime};
use chrono::{DateTime, Utc};
use litchi_cfb::OleError;

/// The first identifier a named user-defined property can take.
const FIRST_USER_DEFINED_IDENTIFIER: u32 = 2;
/// The largest identifier a named user-defined property can take: [MS-OSHARED]
/// 2.3.3.2.3 reserves the high-order bits for linked properties.
const LAST_USER_DEFINED_IDENTIFIER: u32 = 0x00FF_FFFF;

/// The document properties a legacy writer serializes into the
/// `\005SummaryInformation` and `\005DocumentSummaryInformation` streams of a
/// new compound file.
///
/// A new value has no property. [`Self::is_empty`] stays true until a
/// property is set, and a writer emits no property stream for an empty value
/// (the PPT writer keeps its fixed blank streams), so a document whose
/// properties were never set is written exactly as before. See the
/// [module documentation](super) for the names [`Self::set`] accepts and the
/// code-page rules.
#[derive(Debug, Clone, PartialEq)]
pub struct SummaryProperties {
    summary: summary_information::Snapshot,
    document_summary: document_summary::Snapshot,
    /// The user-defined section, present only while it holds a property
    /// besides its code page and dictionary.
    user_defined: Option<Section>,
}

impl SummaryProperties {
    /// Creates document properties with no property set.
    ///
    /// Both standard sections start with the Windows-1252 code page.
    #[must_use]
    pub fn new() -> Self {
        Self {
            summary: summary_information::Snapshot {
                section: blank_section(SUMMARY_INFORMATION_FMTID),
            },
            document_summary: document_summary::Snapshot {
                section: blank_section(DOCUMENT_SUMMARY_INFORMATION_FMTID),
            },
            user_defined: None,
        }
    }

    /// Builds the properties of an existing compound file from its three
    /// sections, as a property-set editor reads them.
    ///
    /// A missing standard section, or a blank one (no property, not even a
    /// code page, as the PPT writer's fixed streams carry), starts as the
    /// Windows-1252 section [`Self::new`] creates. A present section must
    /// pass its typed owner's validation. A user-defined section is kept as
    /// it is until it is edited, and dropped when it holds no property.
    ///
    /// # Errors
    ///
    /// Returns an error if a present standard section breaks its typed
    /// owner's invariants or cannot be cloned.
    pub fn from_sections(
        summary_information: Option<&Section>,
        document_summary_information: Option<&Section>,
        user_defined_properties: Option<&Section>,
    ) -> Result<Self, OleError> {
        let mut properties = Self::new();
        if let Some(section) = summary_information.filter(|section| !section.is_empty()) {
            properties.summary = summary_information::Snapshot::from_section(section)?;
        }
        if let Some(section) = document_summary_information.filter(|section| !section.is_empty()) {
            properties.document_summary = document_summary::Snapshot::from_section(section)?;
        }
        if let Some(section) = user_defined_properties.filter(|section| !holds_no_property(section))
        {
            properties.user_defined = Some(try_clone_property_set(section)?);
        }
        Ok(properties)
    }

    /// Whether no property besides the sections' code pages is set.
    ///
    /// The writers write no property stream for empty properties.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        holds_no_property(self.summary.section())
            && holds_no_property(self.document_summary.section())
            && self.user_defined.is_none()
    }

    /// The typed `SummaryInformation` section.
    #[must_use]
    pub const fn summary_information(&self) -> &summary_information::Snapshot {
        &self.summary
    }

    /// The typed `DocumentSummaryInformation` section.
    #[must_use]
    pub const fn document_summary_information(&self) -> &document_summary::Snapshot {
        &self.document_summary
    }

    /// The user-defined (custom) properties section, when it holds a property.
    #[must_use]
    pub const fn user_defined_properties(&self) -> Option<&Section> {
        self.user_defined.as_ref()
    }

    /// Applies a typed `SummaryInformation` edit.
    ///
    /// The closure is the one the property-set transactions take. The edited
    /// section is validated as a whole before it replaces the current one;
    /// on any error nothing changes. Returns whether the section changed.
    ///
    /// # Errors
    ///
    /// Returns the closure's error, or an error if the edited section breaks
    /// a `SummaryInformation` invariant (for example text its code page
    /// cannot represent, or a revision number that is not a whole number).
    pub fn edit_summary_information<F>(&mut self, edit: F) -> Result<bool, OleError>
    where
        F: for<'a> FnOnce(&mut summary_information::Edit<'a>) -> Result<(), OleError>,
    {
        let section = {
            let mut draft = self.summary.edit()?;
            edit(&mut draft)?;
            let commit = draft.commit()?;
            if !commit.diagnostics().changed() {
                return Ok(false);
            }
            commit.into_package()
        };
        self.summary = summary_information::Snapshot { section };
        Ok(true)
    }

    /// Applies a typed `DocumentSummaryInformation` edit.
    ///
    /// The closure is the one the property-set transactions take. The edited
    /// section is validated as a whole before it replaces the current one,
    /// including that every text value is representable in the section's
    /// code page; on any error nothing changes. Returns whether the section
    /// changed.
    ///
    /// # Errors
    ///
    /// Returns the closure's error, or an error if the edited section breaks
    /// a `DocumentSummaryInformation` invariant.
    pub fn edit_document_summary_information<F>(&mut self, edit: F) -> Result<bool, OleError>
    where
        F: for<'a> FnOnce(&mut document_summary::Edit<'a>) -> Result<(), OleError>,
    {
        let section = {
            let mut draft = self.document_summary.edit()?;
            edit(&mut draft)?;
            let commit = draft.commit()?;
            if !commit.diagnostics().changed() {
                return Ok(false);
            }
            commit.into_package()
        };
        // The typed owner checks string lengths but leaves representability
        // to serialization; check it here so an unwritable value is refused
        // when it is set, not when the document is saved.
        validate_text_representable(&section)?;
        self.document_summary = document_summary::Snapshot { section };
        Ok(true)
    }

    /// Applies a generic edit to the user-defined (custom) properties
    /// section.
    ///
    /// The closure receives the current section, or a new one with the
    /// Windows-1252 code page when there is none; it is the closure the
    /// property-set transactions take. The edited section must keep the
    /// `UserDefinedProperties` format identifier and a code page, pass the
    /// generic property-set validation, hold only the value types
    /// [MS-OSHARED] 2.3.3.2.3.1.2 permits (text, `VT_I4`, `VT_R8`,
    /// `VT_BOOL`, `VT_FILETIME`, `VT_BLOB`) and text its code page can
    /// represent. A section left with no property is dropped. On any error
    /// nothing changes. Returns whether the section changed.
    ///
    /// # Errors
    ///
    /// Returns the closure's error, or an error if the edited section breaks
    /// one of the rules above.
    pub fn edit_user_defined_properties<F>(&mut self, edit: F) -> Result<bool, OleError>
    where
        F: FnOnce(&mut Section) -> Result<(), OleError>,
    {
        let mut draft = match &self.user_defined {
            Some(section) => try_clone_property_set(section)?,
            None => blank_section(USER_DEFINED_PROPERTIES_FMTID),
        };
        edit(&mut draft)?;
        validate_user_defined_section(&draft)?;
        let replacement = (!holds_no_property(&draft)).then_some(draft);
        if replacement == self.user_defined {
            return Ok(false);
        }
        self.user_defined = replacement;
        Ok(true)
    }

    /// Sets one property by name from text.
    ///
    /// The [module documentation](super) lists the recognized names; any
    /// other name sets a user-defined text property (see
    /// [`Self::set_user_defined`]). Timestamps are RFC 3339 text, for example
    /// `2026-10-05T09:30:00Z`. Text the section's code page cannot represent
    /// switches that section to UTF-16 first.
    ///
    /// # Errors
    ///
    /// Returns an error, leaving every property unchanged, if the name is a
    /// built-in property that is not text or begins with `_`, if a timestamp
    /// is not RFC 3339 or cannot be a FILETIME, or if the value breaks the
    /// property's rules (a NUL character, more than 1 MiB of text, a revision
    /// number that is not a whole number).
    pub fn set(&mut self, name: &str, value: &str) -> Result<(), OleError> {
        match names::lookup(name) {
            Some(Name::Summary(property)) => self.set_summary_text(property, value),
            Some(Name::SummaryTime(property)) => {
                let time = parse_timestamp(name, value)?;
                self.edit_summary_information(|edit| match property {
                    SummaryTime::Created => edit.set_create_time(time),
                    SummaryTime::LastSaved => edit.set_last_save_time(time),
                    SummaryTime::LastPrinted => edit.set_last_printed(time),
                })?;
                Ok(())
            },
            Some(Name::DocumentSummary(property)) => {
                self.set_document_summary_text(property, value)
            },
            Some(Name::NotText) => Err(invalid(format!(
                "built-in document property '{name}' is not text; set it with a typed summary edit"
            ))),
            None if name.starts_with('_') => Err(invalid(format!(
                "user-defined property name '{name}' begins with '_', which [MS-OSHARED] reserves for typed system properties; set it with set_user_defined"
            ))),
            None => self.set_user_defined(name, Value::Lpstr(value.to_owned())),
        }
    }

    /// Removes one property by name, the inverse of [`Self::set`].
    ///
    /// Names resolve exactly as [`Self::set`] resolves them: a recognized
    /// built-in name removes that standard property, and any other name
    /// removes the user-defined property of that name (compared
    /// ASCII-case-insensitively). Returns whether a property was removed.
    ///
    /// # Errors
    ///
    /// Returns an error, leaving every property unchanged, for a built-in
    /// property that is not text or a name that begins with `_`, the names
    /// [`Self::set`] refuses as well.
    pub fn remove(&mut self, name: &str) -> Result<bool, OleError> {
        match names::lookup(name) {
            Some(Name::Summary(property)) => self.remove_summary_property(property.identifier()),
            Some(Name::SummaryTime(property)) => {
                self.remove_summary_property(property.identifier())
            },
            Some(Name::DocumentSummary(property)) => {
                let identifier = property.identifier();
                if self
                    .document_summary
                    .section()
                    .property(identifier)
                    .is_none()
                {
                    return Ok(false);
                }
                self.edit_document_summary_information(|edit| {
                    edit.remove(identifier)?;
                    Ok(())
                })
            },
            Some(Name::NotText) => Err(invalid(format!(
                "built-in document property '{name}' is not text; remove it with a typed summary edit"
            ))),
            None if name.starts_with('_') => Err(invalid(format!(
                "user-defined property name '{name}' begins with '_', which [MS-OSHARED] reserves for typed system properties"
            ))),
            None => {
                let present = self
                    .user_defined
                    .as_ref()
                    .is_some_and(|section| section.find_named(name).is_some());
                if !present {
                    return Ok(false);
                }
                self.edit_user_defined_properties(|section| {
                    section.remove_named(name);
                    Ok(())
                })
            },
        }
    }

    fn remove_summary_property(&mut self, identifier: u32) -> Result<bool, OleError> {
        if self.summary.section().property(identifier).is_none() {
            return Ok(false);
        }
        self.edit_summary_information(|edit| {
            edit.remove(identifier)?;
            Ok(())
        })
    }

    /// Sets one typed user-defined (custom) property, replacing the value of
    /// an existing property with the same name (compared
    /// ASCII-case-insensitively, keeping the existing spelling and
    /// identifier).
    ///
    /// The value must be a type [MS-OSHARED] 2.3.3.2.3.1.2 permits: text
    /// ([`Value::Lpstr`] or [`Value::Lpwstr`], stored in the form the
    /// section's code page uses), [`Value::I4`], [`Value::R8`],
    /// [`Value::Bool`], [`Value::Filetime`] or [`Value::Blob`]. A name or
    /// text the section's code page cannot represent switches the section to
    /// UTF-16 first. A new property takes the lowest free identifier.
    ///
    /// # Errors
    ///
    /// Returns an error, leaving every property unchanged, if the name is
    /// empty or contains NUL, the value has another type, text contains NUL
    /// or exceeds 1 MiB, or no identifier is free.
    pub fn set_user_defined(&mut self, name: &str, value: Value) -> Result<(), OleError> {
        validate_property_name(name)?;
        validate_text_limits(name, "user-defined property name")?;
        let text = match &value {
            Value::Lpstr(text) | Value::Lpwstr(text) => {
                validate_text_limits(text, "user-defined property value")?;
                Some(text.as_str())
            },
            Value::I4(_) | Value::R8(_) | Value::Bool(_) | Value::Filetime(_) | Value::Blob(_) => {
                None
            },
            _ => return Err(unpermitted_user_defined_type()),
        };
        let current_page = self
            .user_defined
            .as_ref()
            .and_then(Section::page)
            .unwrap_or(CodePage::WINDOWS_1252);
        let unicode = !representable(current_page, name)
            || text.is_some_and(|text| !representable(current_page, text));
        self.edit_user_defined_properties(|section| {
            if unicode {
                section.set_page(CodePage::Utf16Le);
            }
            let stored = match value {
                Value::Lpstr(text) | Value::Lpwstr(text) => text_value(section.page(), text),
                other => other,
            };
            if let Some((identifier, _)) = section.find_named(name) {
                section.update(identifier, stored)?;
            } else {
                let identifier = unused_user_defined_identifier(section)?;
                section.add_named(identifier, name.to_owned(), stored)?;
            }
            Ok(())
        })?;
        Ok(())
    }

    fn set_summary_text(&mut self, property: SummaryText, value: &str) -> Result<(), OleError> {
        let unicode = !self
            .summary
            .codepage()
            .is_some_and(|page| representable(page, value));
        self.edit_summary_information(|edit| {
            if unicode {
                edit.set_codepage(CodePage::Utf16Le)?;
            }
            match property {
                SummaryText::Title => edit.set_title(value),
                SummaryText::Subject => edit.set_subject(value),
                SummaryText::Author => edit.set_author(value),
                SummaryText::Keywords => edit.set_keywords(value),
                SummaryText::Comments => edit.set_comments(value),
                SummaryText::Template => edit.set_template(value),
                SummaryText::LastAuthor => edit.set_last_author(value),
                SummaryText::RevisionNumber => edit.set_revision_number(value),
                SummaryText::ApplicationName => edit.set_app_name(value),
            }
        })?;
        Ok(())
    }

    fn set_document_summary_text(
        &mut self,
        property: DocumentSummaryText,
        value: &str,
    ) -> Result<(), OleError> {
        let unicode = !self
            .document_summary
            .codepage()
            .is_some_and(|page| representable(page, value));
        self.edit_document_summary_information(|edit| {
            if unicode {
                edit.set_codepage(CodePage::Utf16Le);
            }
            match property {
                DocumentSummaryText::Category => edit.set_category(value),
                DocumentSummaryText::PresentationFormat => edit.set_presentation_format(value),
                DocumentSummaryText::Manager => edit.set_manager(value),
                DocumentSummaryText::Company => edit.set_company(value),
                DocumentSummaryText::ContentType => edit.set_content_type(value),
                DocumentSummaryText::ContentStatus => edit.set_content_status(value),
                DocumentSummaryText::Language => edit.set_language(value),
                DocumentSummaryText::DocumentVersion => edit.set_document_version(value),
            }
        })?;
        Ok(())
    }

    /// Borrows the three sections for serialization.
    pub(super) fn sections(&self) -> (&Section, &Section, Option<&Section>) {
        (
            self.summary.section(),
            self.document_summary.section(),
            self.user_defined.as_ref(),
        )
    }
}

impl Default for SummaryProperties {
    fn default() -> Self {
        Self::new()
    }
}

/// A new section holding only the Windows-1252 code page.
fn blank_section(format_identifier: Guid) -> Section {
    let mut section = Section::new(format_identifier);
    section.set_page(CodePage::WINDOWS_1252);
    section
}

/// Whether a section holds nothing but its code page (and so no dictionary).
fn holds_no_property(section: &Section) -> bool {
    section
        .property_ids()
        .all(|identifier| identifier == PID_CODEPAGE)
}

/// Whether `page` can encode `text`.
fn representable(page: CodePage, text: &str) -> bool {
    match page {
        CodePage::Utf16Le => true,
        CodePage::Mbcs(page) => page.encode(text).is_ok(),
    }
}

/// Text stored in the form the typed owners use for `page`: `VT_LPWSTR`
/// under UTF-16, `VT_LPSTR` under a byte code page.
fn text_value(page: Option<CodePage>, text: String) -> Value {
    if page == Some(CodePage::Utf16Le) {
        Value::Lpwstr(text)
    } else {
        Value::Lpstr(text)
    }
}

fn validate_text_limits(text: &str, field: &str) -> Result<(), OleError> {
    if text.contains('\0') {
        return Err(invalid(format!("{field} must not contain NUL")));
    }
    if text.len() > MAX_TYPED_TEXT_BYTES {
        return Err(invalid(format!("{field} exceeds the typed string limit")));
    }
    Ok(())
}

/// Checks that every byte-code-page text value and dictionary name of a
/// section can be encoded, as serialization will need.
fn validate_text_representable(section: &Section) -> Result<(), OleError> {
    let Some(CodePage::Mbcs(page)) = section.page() else {
        return Ok(());
    };
    for identifier in section.property_ids() {
        if let Some(Value::Lpstr(text)) = section.property(identifier)
            && page.encode(text).is_err()
        {
            return Err(invalid(format!(
                "property {identifier} is not representable in code page {}",
                page.id()
            )));
        }
    }
    for (name, _) in section.named_properties() {
        if page.encode(name).is_err() {
            return Err(invalid(format!(
                "property name '{name}' is not representable in code page {}",
                page.id()
            )));
        }
    }
    Ok(())
}

fn validate_user_defined_section(section: &Section) -> Result<(), OleError> {
    if section.format_identifier != USER_DEFINED_PROPERTIES_FMTID {
        return Err(invalid(
            "UserDefinedProperties format identifier does not match",
        ));
    }
    if section.page().is_none() {
        return Err(invalid(
            "UserDefinedProperties must contain a CodePage property",
        ));
    }
    validate_section(section, Stream::VERSION_0)?;
    for identifier in section.property_ids() {
        if matches!(identifier, PID_DICTIONARY | PID_CODEPAGE) {
            continue;
        }
        match section.property(identifier) {
            Some(Value::Lpstr(text) | Value::Lpwstr(text)) => {
                validate_text_limits(text, "user-defined property value")?;
            },
            Some(
                Value::I4(_) | Value::R8(_) | Value::Bool(_) | Value::Filetime(_) | Value::Blob(_),
            )
            | None => {},
            Some(_) => return Err(unpermitted_user_defined_type()),
        }
    }
    validate_text_representable(section)
}

fn unpermitted_user_defined_type() -> OleError {
    invalid(
        "user-defined properties must be text, VT_I4, VT_R8, VT_BOOL, VT_FILETIME or VT_BLOB ([MS-OSHARED] 2.3.3.2.3.1.2)",
    )
}

fn unused_user_defined_identifier(section: &Section) -> Result<u32, OleError> {
    (FIRST_USER_DEFINED_IDENTIFIER..=LAST_USER_DEFINED_IDENTIFIER)
        .find(|identifier| section.property(*identifier).is_none())
        .ok_or_else(|| invalid("no unused user-defined property identifiers remain"))
}

/// Parses an RFC 3339 timestamp into an exact FILETIME.
fn parse_timestamp(name: &str, value: &str) -> Result<FileTime, OleError> {
    let time = DateTime::parse_from_rfc3339(value).map_err(|error| {
        invalid(format!(
            "document property '{name}' must be an RFC 3339 timestamp: {error}"
        ))
    })?;
    FileTime::from_date_time(time.with_timezone(&Utc))
}
