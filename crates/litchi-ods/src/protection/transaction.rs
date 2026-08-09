//! Immutable protection snapshots and failure-atomic ODS edits.

use super::{codec, model, validation};
use crate::model::protection as wire;
use litchi_core::Result;

/// A complete, immutable protection view of one `content.xml` source.
#[derive(Clone, Debug)]
pub struct Snapshot {
    pub(super) source: String,
    pub(super) styles_xml: Option<String>,
    pub(super) location: codec::Location,
    pub(super) document: model::Document,
    pub(super) sheets: Vec<model::Sheet>,
    pub(super) styles: model::Styles,
}

impl Snapshot {
    /// Parse document/sheet protection and automatic cell-protection styles.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn parse(content_xml: impl Into<String>, styles_xml: Option<&str>) -> Result<Self> {
        let source = content_xml.into();
        let styles_xml = styles_xml.map(str::to_owned);
        let (location, document, sheets, styles) = codec::parse(&source, styles_xml.as_deref())?;
        validation::validate_snapshot(&source, &location, &document, &sheets, &styles)?;
        Ok(Self {
            source,
            styles_xml,
            location,
            document,
            sheets,
            styles,
        })
    }

    pub(crate) fn from_parts(content_xml: &str, styles_xml: Option<&str>) -> Result<Self> {
        Self::parse(content_xml.to_owned(), styles_xml)
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
    pub fn edit(&self) -> Transaction {
        Transaction {
            before: self.clone(),
            draft: self.clone(),
        }
    }
}

/// A reversible protection patch bound to the exact XML source it edits.
///
/// Full source and target XML remain private.  Applying a patch verifies the
/// complete captured source before parsing the target again, so a stale patch
/// can never publish a protection change into a different spreadsheet.
#[derive(Clone, Debug)]
pub struct Patch {
    source: String,
    target: String,
}

impl Patch {
    /// Whether this patch makes no change.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.source == self.target
    }

    /// Return the exact-source patch that restores the accepted source.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            source: self.target.clone(),
            target: self.source.clone(),
        }
    }

    /// Apply this patch only to the exact snapshot that produced it.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn apply(&self, snapshot: &Snapshot) -> Result<Commit> {
        if snapshot.source_xml() != self.source {
            return Err(litchi_core::Error::InvalidFormat(
                "ODS protection patch source snapshot does not match".to_string(),
            ));
        }
        let target = Snapshot::from_parts(&self.target, snapshot.styles_xml.as_deref())?;
        Ok(Commit {
            snapshot: target,
            changed: !self.is_empty(),
            patch: self.clone(),
        })
    }
}

/// A staged protection candidate derived from one immutable [`Snapshot`].
#[derive(Clone, Debug)]
pub struct Transaction {
    before: Snapshot,
    draft: Snapshot,
}

impl Transaction {
    /// Create a transaction from a snapshot.
    #[must_use]
    pub fn new(snapshot: Snapshot) -> Self {
        Self {
            before: snapshot.clone(),
            draft: snapshot,
        }
    }

    /// The current immutable candidate.
    #[must_use]
    pub fn snapshot(&self) -> &Snapshot {
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

    /// Restore the source candidate.
    pub fn rollback(&mut self) {
        self.draft = self.before.clone();
    }

    /// Whether the candidate changes any protected owner.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.before.document != self.draft.document
            || self.before.sheets != self.draft.sheets
            || self.before.styles != self.draft.styles
    }

    /// Validate and atomically materialize the candidate.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn commit(self) -> Result<Commit> {
        validation::validate_candidate(
            &self.before.location,
            &self.before.sheets,
            &self.before.styles,
            &self.draft.document,
            &self.draft.sheets,
            &self.draft.styles,
        )?;
        if !self.is_changed() {
            let patch = Patch {
                source: self.before.source,
                target: self.draft.source.clone(),
            };
            return Ok(Commit {
                snapshot: self.draft,
                changed: false,
                patch,
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
        let snapshot = Snapshot::from_parts(&content_xml, self.before.styles_xml.as_deref())?;
        Ok(Commit {
            snapshot,
            changed: true,
            patch: Patch {
                source: self.before.source,
                target: content_xml,
            },
        })
    }
}

/// A validated protection publication.
#[derive(Debug)]
pub struct Commit {
    snapshot: Snapshot,
    changed: bool,
    patch: Patch,
}

impl Commit {
    /// Whether the candidate changed the source.
    #[must_use]
    pub fn changed(&self) -> bool {
        self.changed
    }

    /// The resulting immutable snapshot.
    #[must_use]
    pub fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Borrow the reversible exact-source patch produced by this commit.
    #[must_use]
    pub fn patch(&self) -> &Patch {
        &self.patch
    }

    /// The resulting `content.xml`.
    #[must_use]
    pub fn content_xml(&self) -> &str {
        &self.snapshot.source
    }

    /// Consume the publication into the rebuilt snapshot.
    #[must_use]
    pub fn into_snapshot(self) -> Snapshot {
        self.snapshot
    }

    /// Consume the publication into package-owned XML.
    #[must_use]
    pub fn into_owned_xml(self) -> String {
        self.snapshot.source
    }
}

/// A concise document-protection editor used by callers that do not need to
/// retain the transaction object.
///
/// # Errors
/// Returns an error when the operation cannot be completed.
pub fn update<F>(snapshot: &Snapshot, edit: F) -> Result<Commit>
where
    F: FnOnce(&mut Transaction) -> Result<()>,
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
