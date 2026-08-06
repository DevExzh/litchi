//! Source-checked PIDSI transactions and reversible typed edits.

use super::super::model::{
    CodePage, Section, Value, invalid, try_clone_property_set, try_clone_string,
};
use super::model::*;
use super::validation::{validate_codepage, validate_section, validate_text_for_page};
use litchi_cfb::OleError;

/// A source-bound, isolated SummaryInformation edit.
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
    pub fn edit(&mut self) -> Edit<'_> {
        Edit {
            section: &mut self.draft,
        }
    }

    /// Publishes the draft after rechecking every PIDSI invariant.
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
    pub const fn section(&self) -> &Section {
        self.section
    }

    /// Changes the section code page after checking every existing LPSTR.
    pub fn set_codepage(&mut self, page: CodePage) -> Result<(), OleError> {
        validate_codepage(self.section, page)?;
        self.section.set_page(page);
        Ok(())
    }

    /// Sets the title using the section's code-page-aware `VT_LPSTR` form.
    pub fn set_title(&mut self, value: &str) -> Result<(), OleError> {
        self.set_string(TITLE, value, "SummaryInformation title")
    }

    /// Sets the subject.
    pub fn set_subject(&mut self, value: &str) -> Result<(), OleError> {
        self.set_string(SUBJECT, value, "SummaryInformation subject")
    }

    /// Sets the author.
    pub fn set_author(&mut self, value: &str) -> Result<(), OleError> {
        self.set_string(AUTHOR, value, "SummaryInformation author")
    }

    /// Sets the keywords.
    pub fn set_keywords(&mut self, value: &str) -> Result<(), OleError> {
        self.set_string(KEYWORDS, value, "SummaryInformation keywords")
    }

    /// Sets the comments.
    pub fn set_comments(&mut self, value: &str) -> Result<(), OleError> {
        self.set_string(COMMENTS, value, "SummaryInformation comments")
    }

    /// Sets the application-specific template.
    pub fn set_template(&mut self, value: &str) -> Result<(), OleError> {
        self.set_string(TEMPLATE, value, "SummaryInformation template")
    }

    /// Sets the last author.
    pub fn set_last_author(&mut self, value: &str) -> Result<(), OleError> {
        self.set_string(LAST_AUTHOR, value, "SummaryInformation last author")
    }

    /// Sets the application-specific revision number.
    pub fn set_revision_number(&mut self, value: &str) -> Result<(), OleError> {
        self.set_string(REVISION_NUMBER, value, "SummaryInformation revision number")
    }

    /// Sets the total editing time as an exact FILETIME.
    pub fn set_edit_time(&mut self, value: FileTime) -> Result<(), OleError> {
        self.set_value(EDIT_TIME, Value::Filetime(value.raw()))
    }

    /// Sets the most-recently-printed timestamp.
    pub fn set_last_printed(&mut self, value: FileTime) -> Result<(), OleError> {
        self.set_value(LAST_PRINTED, Value::Filetime(value.raw()))
    }

    /// Sets the document-creation timestamp.
    pub fn set_create_time(&mut self, value: FileTime) -> Result<(), OleError> {
        self.set_value(CREATE_DTM, Value::Filetime(value.raw()))
    }

    /// Sets the most-recently-saved timestamp.
    pub fn set_last_save_time(&mut self, value: FileTime) -> Result<(), OleError> {
        self.set_value(LAST_SAVE_DTM, Value::Filetime(value.raw()))
    }

    /// Sets the nonnegative page count.
    pub fn set_page_count(&mut self, value: u32) -> Result<(), OleError> {
        self.set_count(PAGE_COUNT, value)
    }

    /// Sets the nonnegative word count.
    pub fn set_word_count(&mut self, value: u32) -> Result<(), OleError> {
        self.set_count(WORD_COUNT, value)
    }

    /// Sets the nonnegative character count.
    pub fn set_character_count(&mut self, value: u32) -> Result<(), OleError> {
        self.set_count(CHARACTER_COUNT, value)
    }

    /// Sets an inert, bounded thumbnail clipboard payload.
    pub fn set_thumbnail(&mut self, value: Thumbnail) -> Result<(), OleError> {
        self.set_value(THUMBNAIL, value.into_value()?)
    }

    /// Sets the creating application name.
    pub fn set_app_name(&mut self, value: &str) -> Result<(), OleError> {
        self.set_string(APP_NAME, value, "SummaryInformation application name")
    }

    /// Sets checked document-security flags.
    pub fn set_document_security(&mut self, value: DocumentSecurity) -> Result<(), OleError> {
        self.set_value(DOC_SECURITY, Value::I4(value.bits() as i32))
    }

    /// Removes an optional or unknown property while protecting CodePage.
    pub fn remove(&mut self, identifier: u32) -> Result<Option<Value>, OleError> {
        if identifier == CODEPAGE {
            return Err(invalid("SummaryInformation CodePage is required"));
        }
        Ok(self.section.remove(identifier))
    }

    fn set_string(&mut self, identifier: u32, value: &str, field: &str) -> Result<(), OleError> {
        let page = self
            .section
            .page()
            .ok_or_else(|| invalid("SummaryInformation has no CodePage"))?;
        validate_text_for_page(value, page, field)?;
        let value = try_clone_string(value, "SummaryInformation string")?;
        let value = match page {
            CodePage::Utf16Le => Value::Lpwstr(value),
            CodePage::Mbcs(_) => Value::Lpstr(value),
        };
        self.set_value(identifier, value)
    }

    fn set_count(&mut self, identifier: u32, value: u32) -> Result<(), OleError> {
        let value =
            i32::try_from(value).map_err(|_| invalid("SummaryInformation count exceeds VT_I4"))?;
        self.set_value(identifier, Value::I4(value))
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

/// A source-checked replacement for one SummaryInformation section.
pub struct Patch<'source> {
    source: &'source Section,
    replacement: Section,
}

impl Patch<'_> {
    /// Applies the replacement only to the exact source section used to make it.
    pub fn apply(&self, source: &Section) -> Result<Section, OleError> {
        if source != self.source {
            return Err(invalid("SummaryInformation patch source does not match"));
        }
        try_clone_property_set(&self.replacement)
    }

    /// Reverts the replacement only when given the exact forward result.
    pub fn revert(&self, replacement: &Section) -> Result<Section, OleError> {
        if replacement != &self.replacement {
            return Err(invalid(
                "SummaryInformation patch replacement does not match",
            ));
        }
        try_clone_property_set(self.source)
    }

    /// Returns the original source section.
    pub const fn source(&self) -> &Section {
        self.source
    }

    /// Returns the complete replacement section.
    pub const fn replacement(&self) -> &Section {
        &self.replacement
    }
}

/// The result of committing a SummaryInformation transaction.
pub struct Commit<'source> {
    patch: Patch<'source>,
    changed: bool,
}

impl<'source> Commit<'source> {
    /// Whether the typed draft differs from its source section.
    pub const fn changed(&self) -> bool {
        self.changed
    }

    /// Borrows the complete replacement section.
    pub const fn section(&self) -> &Section {
        self.patch.replacement()
    }

    /// Borrows the source-checked reversible patch.
    pub const fn patch(&self) -> &Patch<'_> {
        &self.patch
    }

    /// Consumes the commit into its replacement section.
    pub fn into_section(self) -> Section {
        self.patch.replacement
    }

    /// Consumes the commit into its source-checked patch.
    pub fn into_patch(self) -> Patch<'source> {
        self.patch
    }
}
