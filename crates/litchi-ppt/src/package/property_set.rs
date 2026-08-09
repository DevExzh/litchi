//! Source-preserving property-set transactions for a complete PPT package.
//!
//! OLE property-set grammar and typed PID owners are shared by
//! `litchi-ole-common`; this module only owns the `PowerPoint` package boundary
//! and the whole-artifact source check.

use super::{Error, Result};
use litchi_cfb::{OleError, OleFile};
use litchi_ole_common::property_set::{
    self, Binding, PropertySetReader, Section, Stream, USER_DEFINED_PROPERTIES_FMTID,
};
use std::io::Cursor;
use std::sync::Arc;

/// An immutable PPT package source for property-set edits.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Snapshot {
    bytes: Arc<[u8]>,
}

impl Snapshot {
    /// Parses and validates an owned PPT compound-file artifact.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        validate_host(&bytes)?;
        Ok(Self {
            bytes: Arc::from(bytes.into_boxed_slice()),
        })
    }

    /// Parses a borrowed PPT artifact while retaining an owned source copy.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn parse(bytes: &[u8]) -> Result<Self> {
        Self::from_bytes(bytes.to_vec())
    }

    /// Returns the exact source artifact bytes.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.bytes
    }

    /// Projects the standard `SummaryInformation` section, when present.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
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

    /// Projects the standard `DocumentSummaryInformation` section, when present.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
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
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn user_defined_properties(&self) -> Result<Option<Section>> {
        Ok(read_stream(&self.bytes, Binding::UserDefinedProperties)?
            .and_then(|stream| stream.section(USER_DEFINED_PROPERTIES_FMTID).cloned()))
    }

    /// Starts an isolated, source-checked property-set transaction.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn transaction(&self) -> Result<Transaction> {
        Transaction::new(self.clone())
    }

    /// Consumes the snapshot into owned source bytes.
    #[must_use]
    pub fn into_bytes(self) -> Vec<u8> {
        self.bytes.to_vec()
    }
}

/// An isolated PPT property-set transaction over owned CFB bytes.
pub struct Transaction {
    source: Snapshot,
    editor: property_set::Editor,
    changed: bool,
}

impl Transaction {
    fn new(source: Snapshot) -> Result<Self> {
        Ok(Self {
            editor: property_set::Editor::new(source.bytes.to_vec())?,
            source,
            changed: false,
        })
    }

    /// Whether a successful edit has changed the transaction state.
    #[must_use]
    pub const fn is_changed(&self) -> bool {
        self.changed
    }

    /// Projects the current transaction-local `SummaryInformation` section.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
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

    /// Projects the current transaction-local `DocumentSummaryInformation` section.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
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
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn user_defined_properties(&self) -> Result<Option<Section>> {
        self.editor
            .property_set(Binding::UserDefinedProperties)
            .map_err(Into::into)
    }

    /// Applies a typed `SummaryInformation` edit through the common PIDSI owner.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
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

    /// Applies a typed `DocumentSummaryInformation` edit through the common PIDDSI owner.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
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
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
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
    /// `DocumentSummaryInformation` stream.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
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

    /// Publishes the transaction as an owned, source-checked package commit.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
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
    #[must_use]
    pub fn rollback(self) -> Snapshot {
        self.source
    }
}

/// An owned result of a PPT property-set transaction.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
    changed: bool,
}

impl Commit {
    /// Whether the complete output differs from the source artifact.
    #[must_use]
    pub const fn changed(&self) -> bool {
        self.changed
    }

    /// Borrows the validated output snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Borrows the source-checked reversible whole-package patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Consumes the commit into its owned output bytes.
    #[must_use]
    pub fn into_bytes(self) -> Vec<u8> {
        self.snapshot.into_bytes()
    }

    /// Consumes the commit into its output snapshot and patch.
    #[must_use]
    pub fn into_parts(self) -> (Snapshot, Patch) {
        (self.snapshot, self.patch)
    }
}

/// A source-checked reversible whole-PPT package patch.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Patch {
    source: Arc<[u8]>,
    replacement: Arc<[u8]>,
}

impl Patch {
    /// Returns the exact source bytes bound to this patch.
    #[must_use]
    pub fn source(&self) -> &[u8] {
        &self.source
    }

    /// Returns the exact replacement bytes produced by the commit.
    #[must_use]
    pub fn replacement(&self) -> &[u8] {
        &self.replacement
    }

    /// Applies the patch only to the exact source snapshot it was created from.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn apply(&self, source: &Snapshot) -> Result<Snapshot> {
        if source.bytes.as_ref() != self.source.as_ref() {
            return Err(Error::InvalidFormat(
                "PPT property-set patch source does not match".to_string(),
            ));
        }
        Snapshot::from_bytes(self.replacement.to_vec())
    }

    /// Reverts the patch only from its exact replacement snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn revert(&self, replacement: &Snapshot) -> Result<Snapshot> {
        if replacement.bytes.as_ref() != self.replacement.as_ref() {
            return Err(Error::InvalidFormat(
                "PPT property-set patch replacement does not match".to_string(),
            ));
        }
        Snapshot::from_bytes(self.source.to_vec())
    }
}

fn validate_host(bytes: &[u8]) -> Result<()> {
    let ole = OleFile::open(Cursor::new(bytes))?;
    if !ole.exists(&["PowerPoint Document"]) {
        return Err(Error::InvalidFormat(
            "Not a valid PowerPoint document: PowerPoint Document stream not found".to_string(),
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
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
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
        let summary_bytes = PropertyStream::new(summary).to_bytes().unwrap();

        let mut document_summary = Section::new(DOCUMENT_SUMMARY_INFORMATION_FMTID);
        document_summary.set_page(CodePage::Utf16Le);
        document_summary
            .add(0x0000_000F, Value::Lpwstr("before company".to_string()))
            .unwrap();
        let document_summary_bytes = PropertyStream::new(document_summary).to_bytes().unwrap();

        let mut writer = OleWriter::new();
        writer
            .create_stream(&["PowerPoint Document"], b"ppt")
            .unwrap();
        writer
            .create_stream(&["\u{0005}SummaryInformation"], &summary_bytes)
            .unwrap();
        writer
            .create_stream(
                &["\u{0005}DocumentSummaryInformation"],
                &document_summary_bytes,
            )
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
                .document_summary_information()
                .unwrap()
                .unwrap()
                .company(),
            Some("after company")
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
