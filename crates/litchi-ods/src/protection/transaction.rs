//! Immutable protection snapshots and failure-atomic ODS edits.

use super::{codec, model, validation};
use crate::model::protection as wire;
use litchi_core::Result;
use litchi_odf_common::XmlStreamLimits;

/// Return the compact protection state needed by source-backed cell edits
/// without retaining a second copy of `content.xml`.
pub(crate) fn source_edit_protection(
    content_xml: &str,
    styles_xml: Option<&str>,
    limits: &XmlStreamLimits,
) -> Result<(bool, Vec<String>)> {
    let (location, document, sheets, styles) = codec::parse(content_xml, styles_xml, limits)?;
    validation::validate_snapshot(content_xml, &location, &document, &sheets, &styles)?;
    let mut protected = Vec::new();
    protected
        .try_reserve_exact(sheets.len())
        .map_err(|source| litchi_core::Error::Allocation {
            resource: "ODS protected worksheet names",
            source,
        })?;
    protected.extend(
        sheets
            .into_iter()
            .filter(|sheet| sheet.protected == Some(true))
            .map(|sheet| sheet.name),
    );
    Ok((document.structure_protected == Some(true), protected))
}

/// A complete, immutable protection view of one `content.xml` source.
#[derive(Clone, Debug)]
pub struct Snapshot {
    pub(super) source: String,
    pub(super) styles_xml: Option<String>,
    pub(super) location: codec::Location,
    pub(super) document: model::Document,
    pub(super) sheets: Vec<model::Sheet>,
    pub(super) styles: model::Styles,
    pub(super) xml_limits: XmlStreamLimits,
}

impl Snapshot {
    /// Parse document/sheet protection and automatic cell-protection styles.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn parse(
        content_xml: impl Into<String>,
        styles_xml: Option<&str>,
        limits: &XmlStreamLimits,
    ) -> Result<Self> {
        let source = content_xml.into();
        let styles_xml = styles_xml.map(str::to_owned);
        let (location, document, sheets, styles) =
            codec::parse(&source, styles_xml.as_deref(), limits)?;
        validation::validate_snapshot(&source, &location, &document, &sheets, &styles)?;
        Ok(Self {
            source,
            styles_xml,
            location,
            document,
            sheets,
            styles,
            xml_limits: *limits,
        })
    }

    pub(crate) fn from_parts(
        content_xml: &str,
        styles_xml: Option<&str>,
        limits: &XmlStreamLimits,
    ) -> Result<Self> {
        Self::parse(content_xml.to_owned(), styles_xml, limits)
    }

    /// Document-level protection, including inert password-verifier metadata.
    #[must_use]
    pub fn document(&self) -> &model::Document {
        &self.document
    }

    /// Sheet protection in source table order.
    #[must_use]
    pub fn sheets(&self) -> &[model::Sheet] {
        &self.sheets
    }

    /// Lookup a protected sheet by its exact ODF table name.
    #[must_use]
    pub fn sheet(&self, name: &str) -> Option<&model::Sheet> {
        self.sheets.iter().find(|sheet| sheet.name == name)
    }

    /// Automatic table-cell protection styles and inert conditional rules.
    #[must_use]
    pub fn styles(&self) -> &model::Styles {
        &self.styles
    }

    /// The exact source XML captured by this snapshot.
    #[must_use]
    pub fn source_xml(&self) -> &str {
        &self.source
    }

    /// Start a failure-atomic edit.
    #[must_use]
    pub(crate) fn edit(&self) -> Edit {
        Edit {
            before: self.clone(),
            draft: self.clone(),
        }
    }
}

/// A staged protection candidate derived from one immutable [`Snapshot`].
#[derive(Clone, Debug)]
pub struct Edit {
    before: Snapshot,
    draft: Snapshot,
}

impl Edit {
    /// The current immutable candidate.
    #[must_use]
    pub fn current(&self) -> &Snapshot {
        &self.draft
    }

    /// The source snapshot used for source checks and inverse edits.
    #[must_use]
    pub fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Mutably edit document-level protection metadata.
    pub fn document_mut(&mut self) -> &mut model::Document {
        &mut self.draft.document
    }

    /// Mutably edit sheet protection metadata. Sheet names and ordering remain
    /// source-owned and are checked at commit time.
    pub fn sheets_mut(&mut self) -> &mut [model::Sheet] {
        &mut self.draft.sheets
    }

    /// Mutably edit automatic protection styles and inert conditional rules.
    pub fn styles_mut(&mut self) -> &mut model::Styles {
        &mut self.draft.styles
    }

    /// Replace one complete document-level protection value.
    pub fn set_document(&mut self, document: model::Document) {
        self.draft.document = document;
    }

    /// Replace one complete sheet protection catalog.
    pub fn set_sheets(&mut self, sheets: Vec<model::Sheet>) {
        self.draft.sheets = sheets;
    }

    /// Replace the automatic protection-style catalog.
    pub fn set_styles(&mut self, styles: model::Styles) {
        self.draft.styles = styles;
    }

    /// Discard every staged change in place.
    pub fn reset(&mut self) {
        self.draft = self.before.clone();
    }

    /// Whether the candidate changes any protected owner.
    #[must_use]
    pub(crate) fn is_changed(&self) -> bool {
        self.before.document != self.draft.document
            || self.before.sheets != self.draft.sheets
            || self.before.styles != self.draft.styles
    }

    /// Validate and atomically materialize the candidate.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub(crate) fn commit(self) -> Result<Commit> {
        validation::validate_candidate(
            &self.before.location,
            &self.before.sheets,
            &self.before.styles,
            &self.draft.document,
            &self.draft.sheets,
            &self.draft.styles,
        )?;
        if !self.is_changed() {
            return Ok(Commit {
                snapshot: self.draft,
                changed: false,
            });
        }
        let styles_changed = self.before.styles != self.draft.styles;
        let content_xml = codec::replace(
            &self.before.source,
            &self.before.location,
            &self.draft.document,
            &self.draft.sheets,
            &self.draft.styles,
            styles_changed,
        )?;
        let snapshot = Snapshot::from_parts(
            &content_xml,
            self.before.styles_xml.as_deref(),
            &self.before.xml_limits,
        )?;
        Ok(Commit {
            snapshot,
            changed: true,
        })
    }
}

/// A validated protection publication.
#[derive(Debug)]
pub(crate) struct Commit {
    snapshot: Snapshot,
    changed: bool,
}

impl Commit {
    /// Whether the candidate changed the source.
    #[must_use]
    pub(crate) fn changed(&self) -> bool {
        self.changed
    }

    #[cfg(test)]
    /// The resulting immutable snapshot.
    #[must_use]
    pub(crate) fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// The resulting `content.xml`.
    #[must_use]
    pub(crate) fn content_xml(&self) -> &str {
        &self.snapshot.source
    }
}

/// A concise document-protection editor used by callers that do not need to
/// retain the transaction object.
///
/// # Errors
/// Returns an error when the operation cannot be completed.
pub(crate) fn update<F>(snapshot: &Snapshot, edit: F) -> Result<Commit>
where
    F: FnOnce(&mut Edit) -> Result<()>,
{
    let mut transaction = snapshot.edit();
    edit(&mut transaction)?;
    transaction.commit()
}

/// Set the wire-level password-verifier fields without treating them as keys.
#[must_use]
pub fn key(value: Option<String>, digest_algorithm: Option<String>) -> wire::Key {
    wire::Key {
        value,
        digest_algorithm,
        secondary_digest_algorithm: None,
    }
}
