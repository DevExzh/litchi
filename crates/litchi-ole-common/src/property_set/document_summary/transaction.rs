//! Source-checked PIDDSI transactions and typed mutation facade.

use super::super::model::{
    CodePage, DocParts, HeadingPairs, Section, Value, invalid, try_clone_property_set,
    try_clone_string,
};
use super::model::{
    BYTE_COUNT, CATEGORY, CHARACTER_COUNT_WITH_SPACES, COMPANY, CONTENT_STATUS, CONTENT_TYPE,
    DOCUMENT_PARTS, DOCUMENT_VERSION, HEADING_PAIRS, HIDDEN_COUNT, HYPERLINKS_CHANGED, LANGUAGE,
    LINE_COUNT, LINKS_DIRTY, MANAGER, MAX_TEXT_BYTES, MULTIMEDIA_CLIP_COUNT, NOTE_COUNT,
    PARAGRAPH_COUNT, PRESENTATION_FORMAT, SCALE, SHARED_DOCUMENT, SLIDE_COUNT, Snapshot, VERSION,
    Version,
};
use super::validation::validate_section;
use litchi_cfb::OleError;

/// A source-bound, isolated Document Summary Information edit.
pub struct Transaction<'source> {
    source: &'source Section,
    draft: Section,
}

impl<'source> Transaction<'source> {
    pub(super) fn from_snapshot(snapshot: &'source Snapshot) -> Result<Self, OleError> {
        Ok(Self {
            source: &snapshot.section,
            draft: try_clone_property_set(&snapshot.section)?,
        })
    }

    /// Borrows the current draft through the typed mutation facade.
    #[must_use]
    pub fn edit(&mut self) -> Edit<'_> {
        Edit {
            section: &mut self.draft,
        }
    }

    /// Publishes the draft after rechecking the complete PIDDSI invariant set.
    ///
    /// # Errors
    ///
    /// Returns an error if the draft no longer satisfies the typed Document
    /// Summary Information invariants.
    pub fn commit(self) -> Result<Commit<'source>, OleError> {
        validate_section(&self.draft)?;
        let changed = self.source != &self.draft;
        Ok(Commit {
            patch: Patch {
                source: self.source,
                replacement: self.draft,
            },
            changed,
        })
    }
}

/// A typed mutation facade borrowing one transaction draft.
pub struct Edit<'transaction> {
    section: &'transaction mut Section,
}

impl Edit<'_> {
    /// Borrows the full draft for opaque or extension-specific inspection.
    #[must_use]
    pub const fn section(&self) -> &Section {
        self.section
    }

    /// Changes the section code page while retaining all properties.
    pub fn set_codepage(&mut self, page: CodePage) {
        self.section.set_page(page);
    }

    /// Sets a PIDDSI string using LPWSTR for Unicode sections and LPSTR for
    /// code-page sections.
    ///
    /// # Errors
    ///
    /// Returns an error if the text violates the typed string constraints or
    /// the section has no code page.
    pub fn set_category(&mut self, value: &str) -> Result<(), OleError> {
        self.set_string(CATEGORY, value, "category")
    }

    /// Sets the presentation-format string.
    ///
    /// # Errors
    ///
    /// Returns an error if the text violates the typed string constraints or
    /// the section has no code page.
    pub fn set_presentation_format(&mut self, value: &str) -> Result<(), OleError> {
        self.set_string(PRESENTATION_FORMAT, value, "presentation format")
    }

    /// Sets the estimated document byte count.
    ///
    /// # Errors
    ///
    /// Returns an error if storing the replacement property fails.
    pub fn set_byte_count(&mut self, value: i32) -> Result<(), OleError> {
        self.set_i4(BYTE_COUNT, value)
    }

    /// Sets the estimated text-line count.
    ///
    /// # Errors
    ///
    /// Returns an error if storing the replacement property fails.
    pub fn set_line_count(&mut self, value: i32) -> Result<(), OleError> {
        self.set_i4(LINE_COUNT, value)
    }

    /// Sets the paragraph count.
    ///
    /// # Errors
    ///
    /// Returns an error if storing the replacement property fails.
    pub fn set_paragraph_count(&mut self, value: i32) -> Result<(), OleError> {
        self.set_i4(PARAGRAPH_COUNT, value)
    }

    /// Sets the slide count.
    ///
    /// # Errors
    ///
    /// Returns an error if storing the replacement property fails.
    pub fn set_slide_count(&mut self, value: i32) -> Result<(), OleError> {
        self.set_i4(SLIDE_COUNT, value)
    }

    /// Sets the note count.
    ///
    /// # Errors
    ///
    /// Returns an error if storing the replacement property fails.
    pub fn set_note_count(&mut self, value: i32) -> Result<(), OleError> {
        self.set_i4(NOTE_COUNT, value)
    }

    /// Sets the hidden-slide count.
    ///
    /// # Errors
    ///
    /// Returns an error if storing the replacement property fails.
    pub fn set_hidden_count(&mut self, value: i32) -> Result<(), OleError> {
        self.set_i4(HIDDEN_COUNT, value)
    }

    /// Sets the multimedia-clip count.
    ///
    /// # Errors
    ///
    /// Returns an error if storing the replacement property fails.
    pub fn set_multimedia_clip_count(&mut self, value: i32) -> Result<(), OleError> {
        self.set_i4(MULTIMEDIA_CLIP_COUNT, value)
    }

    /// Sets the required false scale flag.
    ///
    /// # Errors
    ///
    /// Returns an error if `value` is true or storing the replacement property
    /// fails.
    pub fn set_scale(&mut self, value: bool) -> Result<(), OleError> {
        self.set_fixed_false(SCALE, value)
    }

    /// Sets `HeadingPairs` while retaining all unrelated properties.
    ///
    /// # Errors
    ///
    /// Returns an error if storing the replacement property fails.
    pub fn set_heading_pairs(&mut self, value: HeadingPairs) -> Result<(), OleError> {
        self.set_value(HEADING_PAIRS, Value::HeadingPairs(value))
    }

    /// Sets document parts with the section's current string representation.
    ///
    /// # Errors
    ///
    /// Returns an error if a part name is invalid, the section has no code
    /// page, or the collection fails typed validation.
    pub fn set_document_parts(&mut self, values: Vec<String>) -> Result<(), OleError> {
        for value in &values {
            validate_text(value, "document part")?;
        }
        let encoding = match self.section.page() {
            Some(CodePage::Utf16Le) => super::super::model::TextEncoding::Unicode,
            Some(CodePage::Mbcs(_)) => super::super::model::TextEncoding::Ansi,
            None => return Err(invalid("Document Summary Information has no CodePage")),
        };
        self.set_value(
            DOCUMENT_PARTS,
            Value::DocParts(DocParts::new(encoding, values)?),
        )
    }

    /// Sets the manager string.
    ///
    /// # Errors
    ///
    /// Returns an error if the text violates the typed string constraints or
    /// the section has no code page.
    pub fn set_manager(&mut self, value: &str) -> Result<(), OleError> {
        self.set_string(MANAGER, value, "manager")
    }

    /// Sets the company string.
    ///
    /// # Errors
    ///
    /// Returns an error if the text violates the typed string constraints or
    /// the section has no code page.
    pub fn set_company(&mut self, value: &str) -> Result<(), OleError> {
        self.set_string(COMPANY, value, "company")
    }

    /// Sets the `LinksDirty` flag.
    ///
    /// # Errors
    ///
    /// Returns an error if storing the replacement property fails.
    pub fn set_links_dirty(&mut self, value: bool) -> Result<(), OleError> {
        self.set_bool(LINKS_DIRTY, value)
    }

    /// Sets the estimated character count including spaces.
    ///
    /// # Errors
    ///
    /// Returns an error if storing the replacement property fails.
    pub fn set_character_count_with_spaces(&mut self, value: i32) -> Result<(), OleError> {
        self.set_i4(CHARACTER_COUNT_WITH_SPACES, value)
    }

    /// Sets the required false `SharedDocument` flag.
    ///
    /// # Errors
    ///
    /// Returns an error if `value` is true or storing the replacement property
    /// fails.
    pub fn set_shared_document(&mut self, value: bool) -> Result<(), OleError> {
        self.set_fixed_false(SHARED_DOCUMENT, value)
    }

    /// Sets the `HyperlinksChanged` flag.
    ///
    /// # Errors
    ///
    /// Returns an error if storing the replacement property fails.
    pub fn set_hyperlinks_changed(&mut self, value: bool) -> Result<(), OleError> {
        self.set_bool(HYPERLINKS_CHANGED, value)
    }

    /// Sets the MS-OSHARED application version.
    ///
    /// # Errors
    ///
    /// Returns an error if storing the replacement property fails.
    pub fn set_version(&mut self, value: Version) -> Result<(), OleError> {
        self.set_i4(VERSION, value.raw())
    }

    /// Sets the content-type string.
    ///
    /// # Errors
    ///
    /// Returns an error if the text violates the typed string constraints or
    /// the section has no code page.
    pub fn set_content_type(&mut self, value: &str) -> Result<(), OleError> {
        self.set_string(CONTENT_TYPE, value, "content type")
    }

    /// Sets the content-status string.
    ///
    /// # Errors
    ///
    /// Returns an error if the text violates the typed string constraints or
    /// the section has no code page.
    pub fn set_content_status(&mut self, value: &str) -> Result<(), OleError> {
        self.set_string(CONTENT_STATUS, value, "content status")
    }

    /// Sets the optional language string.
    ///
    /// # Errors
    ///
    /// Returns an error if the text violates the typed string constraints or
    /// the section has no code page.
    pub fn set_language(&mut self, value: &str) -> Result<(), OleError> {
        self.set_string(LANGUAGE, value, "language")
    }

    /// Sets the optional producer-specific document-version string.
    ///
    /// # Errors
    ///
    /// Returns an error if the text violates the typed string constraints or
    /// the section has no code page.
    pub fn set_document_version(&mut self, value: &str) -> Result<(), OleError> {
        self.set_string(DOCUMENT_VERSION, value, "document version")
    }

    /// Removes an optional property while protecting the required `CodePage`.
    ///
    /// # Errors
    ///
    /// Returns an error if `identifier` is the required `CodePage` property.
    pub fn remove(&mut self, identifier: u32) -> Result<Option<Value>, OleError> {
        if identifier == super::model::CODEPAGE {
            return Err(invalid("Document Summary Information CodePage is required"));
        }
        Ok(self.section.remove(identifier))
    }

    fn set_string(&mut self, identifier: u32, value: &str, field: &str) -> Result<(), OleError> {
        validate_text(value, field)?;
        let text = try_clone_string(value, "PIDDSI string")?;
        let property_value = match self.section.page() {
            Some(CodePage::Utf16Le) => Value::Lpwstr(text),
            Some(CodePage::Mbcs(_)) => Value::Lpstr(text),
            None => return Err(invalid("Document Summary Information has no CodePage")),
        };
        self.set_value(identifier, property_value)
    }

    fn set_i4(&mut self, identifier: u32, value: i32) -> Result<(), OleError> {
        self.set_value(identifier, Value::I4(value))
    }

    fn set_bool(&mut self, identifier: u32, value: bool) -> Result<(), OleError> {
        self.set_value(identifier, Value::Bool(value))
    }

    fn set_fixed_false(&mut self, identifier: u32, value: bool) -> Result<(), OleError> {
        if value {
            return Err(invalid(format!(
                "PIDDSI property {identifier} must be FALSE"
            )));
        }
        self.set_bool(identifier, false)
    }

    fn set_value(&mut self, identifier: u32, value: Value) -> Result<(), OleError> {
        if self.section.property(identifier).is_some() {
            self.section.update(identifier, value)?;
        } else {
            self.section.add(identifier, value)?;
        }
        Ok(())
    }
}

/// A source-checked replacement for one Document Summary Information section.
pub struct Patch<'source> {
    source: &'source Section,
    replacement: Section,
}

impl Patch<'_> {
    /// Applies the replacement only to the exact source section used to make it.
    ///
    /// # Errors
    ///
    /// Returns an error if `source` differs from the recorded source or the
    /// replacement cannot be cloned.
    pub fn apply(&self, source: &Section) -> Result<Section, OleError> {
        if source != self.source {
            return Err(invalid(
                "Document Summary Information patch source does not match",
            ));
        }
        try_clone_property_set(&self.replacement)
    }

    /// Returns the original source section.
    #[must_use]
    pub const fn source(&self) -> &Section {
        self.source
    }

    /// Returns the replacement section without activating any OLE behavior.
    #[must_use]
    pub const fn replacement(&self) -> &Section {
        &self.replacement
    }
}

/// The result of committing a Document Summary Information transaction.
pub struct Commit<'source> {
    patch: Patch<'source>,
    changed: bool,
}

impl<'source> Commit<'source> {
    /// Whether the typed draft differs from its source section.
    #[must_use]
    pub const fn changed(&self) -> bool {
        self.changed
    }

    /// Borrows the complete replacement section.
    #[must_use]
    pub const fn section(&self) -> &Section {
        self.patch.replacement()
    }

    /// Borrows the source-checked patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch<'_> {
        &self.patch
    }

    /// Consumes the commit into its replacement section.
    #[must_use]
    pub fn into_section(self) -> Section {
        self.patch.replacement
    }

    /// Consumes the commit into its source-checked patch.
    #[must_use]
    pub fn into_patch(self) -> Patch<'source> {
        self.patch
    }
}

fn validate_text(value: &str, field: &str) -> Result<(), OleError> {
    if value.len() > MAX_TEXT_BYTES {
        return Err(invalid(format!("{field} exceeds the typed string limit")));
    }
    if value.contains('\0') {
        return Err(invalid(format!("{field} must not contain NUL")));
    }
    Ok(())
}
