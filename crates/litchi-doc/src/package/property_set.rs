//! Source-preserving property-set transactions for a complete DOC package.
//!
//! The OLE property-set grammar and typed PID owners live in
//! `litchi-ole-common`. This module only supplies the DOC package boundary:
//! host validation, owned source bytes, and source-checked whole-package
//! commits. A generic `Package<R>` cannot safely mutate in place because its
//! reader is owned by the parsed CFB file and the original artifact bytes are
//! not recoverable from that API.

use super::{Error, Result};
use crate::user_defined_hyperlinks::{Limits, MutationError, UserDefinedHyperlinks};
use litchi_cfb::{OleError, OleFile};
use litchi_ole_common::property_set::{
    self, Binding, PropertySetReader, Section, Stream, USER_DEFINED_PROPERTIES_FMTID,
};
use std::io::Cursor;
use std::sync::Arc;

/// An immutable DOC package source for property-set edits.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Snapshot {
    bytes: Arc<[u8]>,
}

impl Snapshot {
    /// Parses and validates an owned DOC compound-file artifact.
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        validate_host(&bytes)?;
        Ok(Self {
            bytes: Arc::from(bytes.into_boxed_slice()),
        })
    }

    /// Parses a borrowed DOC artifact while retaining an owned source copy.
    pub fn parse(bytes: &[u8]) -> Result<Self> {
        Self::from_bytes(bytes.to_vec())
    }

    /// Returns the exact source artifact bytes.
    pub fn bytes(&self) -> &[u8] {
        &self.bytes
    }

    /// Projects the standard SummaryInformation section, when present.
    pub fn summary_information(
        &self,
    ) -> Result<Option<property_set::summary_information::Snapshot>> {
        let Some(stream) = read_stream(&self.bytes, Binding::SummaryInformation)? else {
            return Ok(None);
        };
        property_set::summary_information::Snapshot::from_stream(&stream)
            .map(Some)
            .map_err(Into::into)
    }

    /// Projects the standard DocumentSummaryInformation section, when present.
    pub fn document_summary_information(
        &self,
    ) -> Result<Option<property_set::document_summary::Snapshot>> {
        let Some(stream) = read_stream(&self.bytes, Binding::DocumentSummaryInformation)? else {
            return Ok(None);
        };
        property_set::document_summary::Snapshot::from_stream(&stream)
            .map(Some)
            .map_err(Into::into)
    }

    /// Returns the generic user-defined section, when present.
    pub fn user_defined_properties(&self) -> Result<Option<Section>> {
        Ok(read_stream(&self.bytes, Binding::UserDefinedProperties)?
            .and_then(|stream| stream.section(USER_DEFINED_PROPERTIES_FMTID).cloned()))
    }

    /// Reads `_PID_HLINKS` using caller-supplied DOC field context.
    ///
    /// A property-set snapshot does not own WordDocument/Table streams, so it
    /// cannot infer field associations on its own. Pass the parsed field table
    /// from the same source artifact to discover
    /// [`crate::HyperlinkAssociation::FieldCandidates`]. A numeric match is
    /// never a proven field association until the caller selects it with
    /// [`crate::UserDefinedHyperlink::resolve_field`].
    pub fn user_defined_hyperlinks(
        &self,
        fields: Option<&crate::FieldsTable>,
    ) -> Result<Option<UserDefinedHyperlinks>> {
        self.user_defined_hyperlinks_with_limits(fields, Limits::default())
    }

    /// Reads `_PID_HLINKS` with explicit typed-overlay limits.
    pub fn user_defined_hyperlinks_with_limits(
        &self,
        fields: Option<&crate::FieldsTable>,
        limits: Limits,
    ) -> Result<Option<UserDefinedHyperlinks>> {
        let Some(section) = self.user_defined_properties()? else {
            return Ok(None);
        };
        crate::user_defined_hyperlinks::from_user_defined_section_with_limits(
            &section, fields, limits,
        )
    }

    /// Starts an isolated, source-checked property-set transaction.
    ///
    /// The common editor rejects signed or encrypted containers before any
    /// mutation can be staged.
    pub fn transaction(&self) -> Result<Transaction> {
        Transaction::new(self.clone())
    }

    /// Consumes the snapshot into its owned source bytes.
    pub fn into_bytes(self) -> Vec<u8> {
        self.bytes.to_vec()
    }
}

/// An isolated DOC property-set transaction over owned CFB bytes.
pub struct Transaction {
    source: Snapshot,
    editor: property_set::Editor,
    changed: bool,
}

impl Transaction {
    fn new(source: Snapshot) -> Result<Self> {
        let editor = property_set::Editor::new(source.bytes.to_vec())?;
        Ok(Self {
            source,
            editor,
            changed: false,
        })
    }

    /// Whether a successful edit has changed the transaction state.
    pub const fn is_changed(&self) -> bool {
        self.changed
    }

    /// Projects the current transaction-local SummaryInformation section.
    pub fn summary_information(
        &self,
    ) -> Result<Option<property_set::summary_information::Snapshot>> {
        let Some(section) = self.editor.property_set(Binding::SummaryInformation)? else {
            return Ok(None);
        };
        property_set::summary_information::Snapshot::from_section(&section)
            .map(Some)
            .map_err(Into::into)
    }

    /// Projects the current transaction-local DocumentSummaryInformation section.
    pub fn document_summary_information(
        &self,
    ) -> Result<Option<property_set::document_summary::Snapshot>> {
        let Some(section) = self
            .editor
            .property_set(Binding::DocumentSummaryInformation)?
        else {
            return Ok(None);
        };
        property_set::document_summary::Snapshot::from_section(&section)
            .map(Some)
            .map_err(Into::into)
    }

    /// Returns the current transaction-local user-defined section.
    pub fn user_defined_properties(&self) -> Result<Option<Section>> {
        self.editor
            .property_set(Binding::UserDefinedProperties)
            .map_err(Into::into)
    }

    /// Applies a typed SummaryInformation edit through the common PIDSI owner.
    ///
    /// The common typed transaction is committed before its replacement is
    /// handed to the common CFB editor. A no-op typed commit does not stage a
    /// stream replacement, preserving the complete source artifact byte-for-byte.
    pub fn edit_summary_information<F>(&mut self, edit: F) -> Result<bool>
    where
        F: for<'a> FnOnce(
            &mut property_set::summary_information::Edit<'a>,
        ) -> std::result::Result<(), OleError>,
    {
        let source = self
            .editor
            .property_set(Binding::SummaryInformation)?
            .ok_or(OleError::StreamNotFound)?;
        let snapshot = property_set::summary_information::Snapshot::from_section(&source)?;
        let mut transaction = snapshot.transaction()?;
        {
            let mut draft = transaction.edit();
            edit(&mut draft)?;
        }
        let commit = transaction.commit()?;
        if !commit.changed() {
            return Ok(false);
        }
        self.editor
            .replace(Binding::SummaryInformation, commit.into_section())?;
        self.changed = true;
        Ok(true)
    }

    /// Applies a typed DocumentSummaryInformation edit through the common PIDDSI owner.
    pub fn edit_document_summary_information<F>(&mut self, edit: F) -> Result<bool>
    where
        F: for<'a> FnOnce(
            &mut property_set::document_summary::Edit<'a>,
        ) -> std::result::Result<(), OleError>,
    {
        let source = self
            .editor
            .property_set(Binding::DocumentSummaryInformation)?
            .ok_or(OleError::StreamNotFound)?;
        let snapshot = property_set::document_summary::Snapshot::from_section(&source)?;
        let mut transaction = snapshot.transaction()?;
        {
            let mut draft = transaction.edit();
            edit(&mut draft)?;
        }
        let commit = transaction.commit()?;
        if !commit.changed() {
            return Ok(false);
        }
        self.editor
            .replace(Binding::DocumentSummaryInformation, commit.into_section())?;
        self.changed = true;
        Ok(true)
    }

    /// Applies a generic user-defined property-section edit.
    pub fn edit_user_defined_properties<F>(&mut self, edit: F) -> Result<bool>
    where
        F: FnOnce(&mut Section) -> std::result::Result<(), OleError>,
    {
        let source = self.editor.property_set(Binding::UserDefinedProperties)?;
        let mut candidate = source
            .clone()
            .unwrap_or_else(|| Section::new(USER_DEFINED_PROPERTIES_FMTID));
        let initial = candidate.clone();
        edit(&mut candidate)?;
        if source.as_ref().is_some_and(|value| value == &candidate)
            || source.is_none() && candidate == initial
        {
            return Ok(false);
        }
        self.editor
            .replace(Binding::UserDefinedProperties, candidate)?;
        self.changed = true;
        Ok(true)
    }

    /// Removes the user-defined section while retaining the rest of the
    /// DocumentSummaryInformation stream.
    pub fn remove_user_defined_properties(&mut self) -> Result<bool> {
        if self
            .editor
            .property_set(Binding::UserDefinedProperties)?
            .is_none()
        {
            return Ok(false);
        }
        self.editor.remove(Binding::UserDefinedProperties)?;
        self.changed = true;
        Ok(true)
    }

    /// Replaces `_PID_HLINKS` atomically through the shared typed owner.
    ///
    /// Only caller-resolved field entries are serialized in the MS-DOC section
    /// 2.4.7 story/index order. All other entries remain inert and stable
    /// relative to one another. A changed value containing unresolved field
    /// candidates is refused; an exact raw no-op preserves it. This
    /// transaction never writes PIDDSI `0x15`.
    pub fn put_user_defined_hyperlinks(
        &mut self,
        hyperlinks: &UserDefinedHyperlinks,
    ) -> std::result::Result<bool, MutationError> {
        self.put_user_defined_hyperlinks_with_limits(hyperlinks, Limits::default())
    }

    /// Replaces `_PID_HLINKS` with explicit shared typed-overlay limits.
    pub fn put_user_defined_hyperlinks_with_limits(
        &mut self,
        hyperlinks: &UserDefinedHyperlinks,
        limits: Limits,
    ) -> std::result::Result<bool, MutationError> {
        let current = self
            .user_defined_properties()
            .map_err(MutationError::from)?;
        if crate::user_defined_hyperlinks::unchanged_with_unresolved_candidates(
            current.as_ref(),
            hyperlinks,
            limits,
        )
        .map_err(Error::from)
        .map_err(MutationError::from)?
        {
            return Ok(false);
        }
        if hyperlinks.entries().iter().any(|entry| {
            matches!(
                entry.association(),
                crate::HyperlinkAssociation::FieldCandidates(_)
            )
        }) {
            return Err(MutationError::UnresolvedFieldCandidates);
        }
        self.edit_user_defined_properties(|section| {
            crate::user_defined_hyperlinks::put(section, hyperlinks, limits)
        })
        .map_err(MutationError::from)
    }

    /// Removes only the named `_PID_HLINKS` user-defined property atomically.
    ///
    /// This never writes PIDDSI `0x15` and leaves unrelated user-defined
    /// properties and dictionary entries intact.
    pub fn remove_user_defined_hyperlinks(&mut self) -> Result<bool> {
        self.edit_user_defined_properties(|section| {
            crate::user_defined_hyperlinks::remove(section, Limits::default())?;
            Ok(())
        })
    }

    /// Publishes the transaction as an owned, source-checked package commit.
    pub fn commit(self) -> Result<Commit> {
        let before = self.source.bytes.clone();
        let bytes = self.editor.finish()?;
        let snapshot = Snapshot::from_bytes(bytes)?;
        let changed = before.as_ref() != snapshot.bytes.as_ref();
        Ok(Commit {
            patch: Patch {
                source: before,
                replacement: snapshot.bytes.clone(),
            },
            snapshot,
            changed,
        })
    }

    /// Discards the transaction and recovers its exact source snapshot.
    pub fn rollback(self) -> Snapshot {
        self.source
    }
}

/// An owned result of a DOC property-set transaction.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
    changed: bool,
}

impl Commit {
    /// Whether the complete output differs from the source artifact.
    pub const fn changed(&self) -> bool {
        self.changed
    }

    /// Borrows the validated output snapshot.
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Borrows the source-checked reversible whole-package patch.
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Consumes the commit into its owned output bytes.
    pub fn into_bytes(self) -> Vec<u8> {
        self.snapshot.into_bytes()
    }

    /// Consumes the commit into its output snapshot and patch.
    pub fn into_parts(self) -> (Snapshot, Patch) {
        (self.snapshot, self.patch)
    }
}

/// A source-checked reversible whole-DOC package patch.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Patch {
    source: Arc<[u8]>,
    replacement: Arc<[u8]>,
}

impl Patch {
    /// Returns the exact source bytes bound to this patch.
    pub fn source(&self) -> &[u8] {
        &self.source
    }

    /// Returns the exact replacement bytes produced by the commit.
    pub fn replacement(&self) -> &[u8] {
        &self.replacement
    }

    /// Applies the patch only to the exact source snapshot it was created from.
    pub fn apply(&self, source: &Snapshot) -> Result<Snapshot> {
        if source.bytes.as_ref() != self.source.as_ref() {
            return Err(Error::InvalidFormat(
                "DOC property-set patch source does not match".to_string(),
            ));
        }
        Snapshot::from_bytes(self.replacement.to_vec())
    }

    /// Reverts the patch only from its exact replacement snapshot.
    pub fn revert(&self, replacement: &Snapshot) -> Result<Snapshot> {
        if replacement.bytes.as_ref() != self.replacement.as_ref() {
            return Err(Error::InvalidFormat(
                "DOC property-set patch replacement does not match".to_string(),
            ));
        }
        Snapshot::from_bytes(self.source.to_vec())
    }
}

fn validate_host(bytes: &[u8]) -> Result<()> {
    let ole = OleFile::open(Cursor::new(bytes))?;
    if !ole.exists(&["WordDocument"]) {
        return Err(Error::InvalidFormat(
            "Not a valid Word document: WordDocument stream not found".to_string(),
        ));
    }
    Ok(())
}

fn read_stream(bytes: &[u8], binding: Binding) -> Result<Option<Stream>> {
    let mut ole = OleFile::open(Cursor::new(bytes))?;
    match ole.property_set(binding) {
        Ok(stream) => Ok(Some(stream)),
        Err(OleError::StreamNotFound) => Ok(None),
        Err(error) => Err(error.into()),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_cfb::{OleFile, OleWriter};
    use litchi_ole_common::property_set::{
        CodePage, DOCUMENT_SUMMARY_INFORMATION_FMTID, SUMMARY_INFORMATION_FMTID, Section,
        Stream as PropertyStream, Value,
    };
    use std::io::Cursor;

    fn package_bytes(extra_stream: Option<&str>) -> Vec<u8> {
        let mut summary = Section::new(SUMMARY_INFORMATION_FMTID);
        summary.set_page(CodePage::Utf16Le);
        summary.add(2, Value::Lpwstr("before".to_string())).unwrap();
        let summary = PropertyStream::new(summary).to_bytes().unwrap();

        let mut document_summary = Section::new(DOCUMENT_SUMMARY_INFORMATION_FMTID);
        document_summary.set_page(CodePage::Utf16Le);
        document_summary
            .add(0x0000_000F, Value::Lpwstr("before company".to_string()))
            .unwrap();
        let document_summary = PropertyStream::new(document_summary).to_bytes().unwrap();

        let mut writer = OleWriter::new();
        writer.create_stream(&["WordDocument"], b"word").unwrap();
        writer
            .create_stream(&["\u{0005}SummaryInformation"], &summary)
            .unwrap();
        writer
            .create_stream(&["\u{0005}DocumentSummaryInformation"], &document_summary)
            .unwrap();
        writer
            .create_stream(&["Unrelated", "Nested"], b"untouched")
            .unwrap();
        if let Some(name) = extra_stream {
            writer.create_stream(&[name], b"marker").unwrap();
        }
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output).unwrap();
        output.into_inner()
    }

    #[test]
    fn typed_edits_preserve_unrelated_streams_and_publish_a_patch() {
        let source = package_bytes(None);
        let snapshot = Snapshot::from_bytes(source.clone()).unwrap();
        assert_eq!(
            snapshot.summary_information().unwrap().unwrap().title(),
            Some("before")
        );
        assert_eq!(
            snapshot
                .document_summary_information()
                .unwrap()
                .unwrap()
                .company(),
            Some("before company")
        );

        let mut transaction = snapshot.transaction().unwrap();
        assert!(
            transaction
                .edit_summary_information(|edit| edit.set_title("after"))
                .unwrap()
        );
        assert!(
            transaction
                .edit_document_summary_information(|edit| edit.set_company("after company"))
                .unwrap()
        );
        let commit = transaction.commit().unwrap();
        assert!(commit.changed());
        assert_eq!(
            commit
                .snapshot()
                .summary_information()
                .unwrap()
                .unwrap()
                .title(),
            Some("after")
        );

        let mut ole = OleFile::open(Cursor::new(commit.snapshot().bytes().to_vec())).unwrap();
        assert_eq!(
            ole.open_stream(&["Unrelated", "Nested"]).unwrap(),
            b"untouched"
        );
        let applied = commit.patch().apply(&snapshot).unwrap();
        assert_eq!(applied.bytes(), commit.snapshot().bytes());
        assert_eq!(commit.patch().revert(&applied).unwrap().bytes(), source);
    }

    #[test]
    fn no_op_edits_return_the_exact_source_bytes() {
        let source = package_bytes(None);
        let snapshot = Snapshot::from_bytes(source.clone()).unwrap();
        let mut transaction = snapshot.transaction().unwrap();
        assert!(!transaction.edit_summary_information(|_| Ok(())).unwrap());
        let commit = transaction.commit().unwrap();
        assert!(!commit.changed());
        assert_eq!(commit.into_bytes(), source);
    }

    #[test]
    fn signed_and_encrypted_sources_are_refused_before_mutation() {
        assert!(
            Snapshot::from_bytes(package_bytes(Some("DigitalSignature")))
                .unwrap()
                .transaction()
                .is_err()
        );
        assert!(
            Snapshot::from_bytes(package_bytes(Some("EncryptedPackage")))
                .unwrap()
                .transaction()
                .is_err()
        );
    }
}
