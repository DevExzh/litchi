//! Owned, XLS-contextual editing of OLE Summary Information property sets.
//!
//! The compound-file and property-set grammars remain owned by
//! `litchi-ole-common`. This module only validates the XLS host stream and
//! turns the common editor into a source-checked whole-package transaction.

use crate::error::Result;
use litchi_cfb::{OleError, OleFile};
use litchi_ole_common::property_set::{
    Binding, CodePage, DOCUMENT_SUMMARY_INFORMATION_FMTID, Editor, PropertySetReader, Section,
    USER_DEFINED_PROPERTIES_FMTID,
};
use std::io::{Cursor, Read, Seek};
use std::sync::Arc;

/// An immutable, owned XLS compound-file source with typed property-set views.
#[derive(Clone, Debug, PartialEq)]
pub struct Snapshot {
    bytes: Arc<[u8]>,
    summary_information: Option<litchi_ole_common::property_set::summary_information::Snapshot>,
    document_summary_information:
        Option<litchi_ole_common::property_set::document_summary::Snapshot>,
    user_defined_properties: Option<Section>,
}

impl Snapshot {
    /// Validates an XLS compound file and captures its exact source bytes.
    ///
    /// A host is recognized by a root `Workbook` or `Book` stream. Standard
    /// property-set streams are parsed through the shared typed owners;
    /// unknown user-defined properties remain attached as their generic
    /// `Section`.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        let mut ole = OleFile::open(Cursor::new(bytes.clone()))?;
        validate_host(&mut ole)?;

        let summary_information = match ole.property_set(Binding::SummaryInformation) {
            Ok(stream) => Some(
                litchi_ole_common::property_set::summary_information::Snapshot::from_stream(
                    &stream,
                )?,
            ),
            Err(OleError::StreamNotFound) => None,
            Err(error) => return Err(error.into()),
        };

        let (document_summary_information, user_defined_properties) =
            match ole.property_set(Binding::DocumentSummaryInformation) {
                Ok(stream) => {
                    let document_summary_information =
                        if stream.section(DOCUMENT_SUMMARY_INFORMATION_FMTID).is_some() {
                            Some(
                        litchi_ole_common::property_set::document_summary::Snapshot::from_stream(
                            &stream,
                        )?,
                    )
                        } else {
                            None
                        };
                    let user_defined_properties =
                        stream.section(USER_DEFINED_PROPERTIES_FMTID).cloned();
                    (document_summary_information, user_defined_properties)
                },
                Err(OleError::StreamNotFound) => (None, None),
                Err(error) => return Err(error.into()),
            };

        Ok(Self {
            bytes: Arc::from(bytes.into_boxed_slice()),
            summary_information,
            document_summary_information,
            user_defined_properties,
        })
    }

    /// Returns the exact source bytes, including all untouched CFB content.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.bytes
    }

    /// Returns the typed `SummaryInformation` view when its stream is present.
    #[must_use]
    pub const fn summary_information(
        &self,
    ) -> Option<&litchi_ole_common::property_set::summary_information::Snapshot> {
        self.summary_information.as_ref()
    }

    /// Returns the typed `DocumentSummaryInformation` view when its section is
    /// present.
    #[must_use]
    pub const fn document_summary_information(
        &self,
    ) -> Option<&litchi_ole_common::property_set::document_summary::Snapshot> {
        self.document_summary_information.as_ref()
    }

    /// Returns the generic user-defined section when it is present.
    #[must_use]
    pub const fn user_defined_properties(&self) -> Option<&Section> {
        self.user_defined_properties.as_ref()
    }

    /// Starts a transaction whose common editor owns a copy of this source.
    ///
    /// The common editor rejects signed and encrypted containers before any
    /// edit can be staged.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn transaction(&self) -> Result<Transaction> {
        Ok(Transaction {
            source: self.clone(),
            editor: Editor::new(self.bytes.to_vec())?,
        })
    }
}

/// A failure-atomic XLS property-set transaction over an owned CFB editor.
pub struct Transaction {
    source: Snapshot,
    editor: Editor,
}

impl Transaction {
    /// Returns the immutable package source used for conflict checks.
    #[must_use]
    pub const fn before(&self) -> &Snapshot {
        &self.source
    }

    /// Applies a typed PIDSI edit, staging only an actual section change.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn edit_summary_information<F>(&mut self, edit: F) -> Result<()>
    where
        F: FnOnce(
            &mut litchi_ole_common::property_set::summary_information::Edit<'_>,
        ) -> std::result::Result<(), OleError>,
    {
        let section = self.editor.property_set(Binding::SummaryInformation)?;
        let snapshot = match section {
            Some(section) => {
                litchi_ole_common::property_set::summary_information::Snapshot::from_section(
                    &section,
                )?
            },
            None => litchi_ole_common::property_set::summary_information::Snapshot::new(
                CodePage::WINDOWS_1252,
            )?,
        };
        let mut typed = snapshot.transaction()?;
        edit(&mut typed.edit())?;
        let commit = typed.commit()?;
        if commit.changed() {
            self.editor
                .replace(Binding::SummaryInformation, commit.into_section())?;
        }
        Ok(())
    }

    /// Applies a typed PIDDSI edit, staging only an actual section change.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn edit_document_summary_information<F>(&mut self, edit: F) -> Result<()>
    where
        F: FnOnce(
            &mut litchi_ole_common::property_set::document_summary::Edit<'_>,
        ) -> std::result::Result<(), OleError>,
    {
        let section = self
            .editor
            .property_set(Binding::DocumentSummaryInformation)?;
        let snapshot = match section {
            Some(section) => {
                litchi_ole_common::property_set::document_summary::Snapshot::from_section(&section)?
            },
            None => litchi_ole_common::property_set::document_summary::Snapshot::new(
                CodePage::WINDOWS_1252,
            )?,
        };
        let mut typed = snapshot.transaction()?;
        edit(&mut typed.edit())?;
        let commit = typed.commit()?;
        if commit.changed() {
            self.editor
                .replace(Binding::DocumentSummaryInformation, commit.into_section())?;
        }
        Ok(())
    }

    /// Applies a generic user-defined property-section edit, staging only an
    /// actual section change.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn edit_user_defined_properties<F>(&mut self, edit: F) -> Result<()>
    where
        F: FnOnce(&mut Section) -> std::result::Result<(), OleError>,
    {
        let current = self
            .editor
            .property_set(Binding::UserDefinedProperties)?
            .unwrap_or_else(|| {
                let mut section = Section::new(USER_DEFINED_PROPERTIES_FMTID);
                section.set_page(CodePage::WINDOWS_1252);
                section
            });
        let mut candidate = current.clone();
        edit(&mut candidate)?;
        if candidate != current {
            self.editor
                .replace(Binding::UserDefinedProperties, candidate)?;
        }
        Ok(())
    }

    /// Removes the user-defined section while retaining the standard
    /// `DocumentSummaryInformation` section and all other package streams.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn remove_user_defined_properties(&mut self) -> Result<Option<Section>> {
        Ok(self.editor.remove(Binding::UserDefinedProperties)?)
    }

    /// Discards the transaction and returns its unchanged source snapshot.
    #[must_use]
    pub fn rollback(self) -> Snapshot {
        self.source
    }

    /// Publishes the edited package and its source-checked reversible patch.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn commit(self) -> Result<Commit> {
        let source = self.source;
        let before = source.bytes().to_vec();
        let bytes = self.editor.finish()?;
        let changed = bytes != before;
        let snapshot = if changed {
            Snapshot::from_bytes(bytes)?
        } else {
            source
        };
        let patch = Patch::new(before, snapshot.bytes().to_vec());
        Ok(Commit {
            changed,
            snapshot,
            patch,
        })
    }
}

/// A committed XLS package result and its reversible whole-package patch.
pub struct Commit {
    changed: bool,
    snapshot: Snapshot,
    patch: Patch,
}

impl Commit {
    /// Whether publication changed any package byte.
    #[must_use]
    pub const fn changed(&self) -> bool {
        self.changed
    }

    /// Returns the committed package snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Returns the source-checked whole-package patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Consumes the commit into its committed snapshot.
    #[must_use]
    pub fn into_snapshot(self) -> Snapshot {
        self.snapshot
    }

    /// Consumes the commit into its source-checked patch.
    #[must_use]
    pub fn into_patch(self) -> Patch {
        self.patch
    }

    /// Consumes the commit into owned package bytes.
    #[must_use]
    pub fn into_bytes(self) -> Vec<u8> {
        self.snapshot.bytes.to_vec()
    }
}

/// A reversible, source-checked replacement of a complete XLS CFB artifact.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Patch {
    before: Arc<[u8]>,
    after: Arc<[u8]>,
}

impl Patch {
    fn new(before: Vec<u8>, after: Vec<u8>) -> Self {
        Self {
            before: Arc::from(before.into_boxed_slice()),
            after: Arc::from(after.into_boxed_slice()),
        }
    }

    /// Returns the exact source bytes required by this patch.
    #[must_use]
    pub fn before(&self) -> &[u8] {
        &self.before
    }

    /// Returns the exact output bytes produced by this patch.
    #[must_use]
    pub fn after(&self) -> &[u8] {
        &self.after
    }

    /// Whether this patch is a byte-for-byte no-op.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after
    }

    /// Applies this patch only to a snapshot with the exact source bytes.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn apply(&self, source: &Snapshot) -> Result<Snapshot> {
        if source.bytes() != self.before() {
            return Err(crate::error::Error::UnsafeEdit(
                "XLS property-set patch source does not match its base snapshot".to_string(),
            ));
        }
        if self.is_noop() {
            return Ok(source.clone());
        }
        Snapshot::from_bytes(self.after.to_vec())
    }

    /// Reverts this patch only from its exact replacement snapshot.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn revert(&self, replacement: &Snapshot) -> Result<Snapshot> {
        if replacement.bytes() != self.after() {
            return Err(crate::error::Error::UnsafeEdit(
                "XLS property-set patch replacement does not match its target snapshot".to_string(),
            ));
        }
        Snapshot::from_bytes(self.before.to_vec())
    }

    /// Returns the exact inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: Arc::clone(&self.after),
            after: Arc::clone(&self.before),
        }
    }
}

fn validate_host<R: Read + Seek>(ole: &mut OleFile<R>) -> Result<()> {
    match ole.open_stream(&["Workbook"]) {
        Ok(_) => Ok(()),
        Err(OleError::StreamNotFound) => match ole.open_stream(&["Book"]) {
            Ok(_) => Ok(()),
            Err(OleError::StreamNotFound) => Err(crate::error::Error::InvalidData(
                "XLS package has no Workbook or Book stream".to_string(),
            )),
            Err(error) => Err(error.into()),
        },
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
        writer.create_stream(&["Workbook"], b"xls").unwrap();
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
            snapshot.summary_information().unwrap().title(),
            Some("before")
        );
        assert_eq!(
            snapshot.document_summary_information().unwrap().company(),
            Some("before company")
        );

        let mut transaction = snapshot.transaction().unwrap();
        transaction
            .edit_summary_information(|edit| edit.set_title("after"))
            .unwrap();
        transaction
            .edit_document_summary_information(|edit| edit.set_company("after company"))
            .unwrap();
        let commit = transaction.commit().unwrap();
        assert!(commit.changed());
        assert_eq!(
            commit.snapshot().summary_information().unwrap().title(),
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
        transaction.edit_summary_information(|_| Ok(())).unwrap();
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
