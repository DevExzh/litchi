//! The XLS property-set edit family (record 0351 §2.3): typed edits of the
//! OLE `SummaryInformation`, `DocumentSummaryInformation` and user-defined
//! property sets of one [`Package`].
//!
//! [`Package::property_sets`] reads the sections, [`Package::edit_property_sets`]
//! stages edits and [`Edit::commit`] publishes a new [`Package`] with a
//! whole-package [`Patch`] and [`Diagnostics`]; [`Package::apply_property_sets`]
//! replays the patch. The compound-file and property-set grammars remain owned
//! by `litchi-ole-common`; this module validates the XLS host stream.

use crate::Package;
/// The family error: the package error of the source or republished
/// package.
pub use crate::package::{Error, Result};
use litchi_cfb::{OleError, OleFile, OleFileLimits};
use litchi_ole_common::object::Limits as ObjectLimits;
use litchi_ole_common::property_set::document_summary::Snapshot as DocumentSummary;
use litchi_ole_common::property_set::summary_information::Snapshot as SummaryInformation;
use litchi_ole_common::property_set::{
    Binding, CodePage, DOCUMENT_SUMMARY_INFORMATION_FMTID, Edit as PropertySetEdit,
    PropertySetReader, SUMMARY_INFORMATION_FMTID, Section, USER_DEFINED_PROPERTIES_FMTID,
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
        Self::from_bytes_with_cfb_limits(bytes, OleFileLimits::default())
    }

    /// Parse a property-set source under the package's retained CFB policy.
    ///
    /// This is crate-private because the public snapshot constructor keeps
    /// the codec's default profile; package owners pass the profile that
    /// admitted their source when an edit first materializes.
    pub(crate) fn from_bytes_with_cfb_limits(
        bytes: Vec<u8>,
        cfb_limits: OleFileLimits,
    ) -> Result<Self> {
        let mut ole = OleFile::open(Cursor::new(bytes.clone()), cfb_limits)?;
        validate_host(&mut ole)?;

        // A blank standard section (no property, not even a code page)
        // carries nothing and is read as absent, so a later edit starts it
        // afresh instead of failing on the typed owner's code-page rule.
        let summary_information = match ole.property_set(Binding::SummaryInformation) {
            Ok(stream)
                if stream
                    .section(SUMMARY_INFORMATION_FMTID)
                    .is_some_and(Section::is_empty) =>
            {
                None
            },
            Ok(stream) => Some(SummaryInformation::from_stream(&stream)?),
            Err(litchi_ole_common::Error::Cfb(OleError::StreamNotFound)) => None,
            Err(error) => return Err(error.into()),
        };

        let (document_summary_information, user_defined_properties) =
            match ole.property_set(Binding::DocumentSummaryInformation) {
                Ok(stream) => {
                    let document_summary_information = if stream
                        .section(DOCUMENT_SUMMARY_INFORMATION_FMTID)
                        .is_some_and(|section| !section.is_empty())
                    {
                        Some(DocumentSummary::from_stream(&stream)?)
                    } else {
                        None
                    };
                    let user_defined_properties =
                        stream.section(USER_DEFINED_PROPERTIES_FMTID).cloned();
                    (document_summary_information, user_defined_properties)
                },
                Err(litchi_ole_common::Error::Cfb(OleError::StreamNotFound)) => (None, None),
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
}

/// A failure-atomic XLS property-set transaction over an owned CFB editor.
/// A staged property-set edit of one package. Dropping it discards it.
pub struct Edit {
    /// The package a commit republishes; `None` only for codec-level tests.
    base: Option<Package>,
    signed: bool,
    source: Option<Snapshot>,
    editor: Option<PropertySetEdit>,
}

impl std::fmt::Debug for Edit {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.debug_struct("Edit").finish_non_exhaustive()
    }
}

impl Edit {
    pub(crate) fn new(base: &Package) -> Result<Self> {
        if base.is_encrypted() {
            return Err(Error::EncryptionPolicy);
        }
        Ok(Self {
            base: Some(base.clone()),
            signed: base.is_signed(),
            source: None,
            editor: None,
        })
    }

    /// A codec-level edit of a bare property-set snapshot (tests only).
    #[cfg(test)]
    fn detached(source: Snapshot) -> Result<Self> {
        let signed = {
            let ole = OleFile::open(Cursor::new(source.bytes.to_vec()), OleFileLimits::default())?;
            [
                "_xmlsignatures",
                "_signatures",
                "DigitalSignature",
                "\u{0005}DigitalSignature",
            ]
            .iter()
            .any(|name| ole.exists(&[*name]))
        };
        let editor = if signed {
            None
        } else {
            Some(
                PropertySetEdit::new(source.bytes.to_vec(), OleFileLimits::default())
                    .map_err(crate::Error::from)?,
            )
        };
        Ok(Self {
            base: None,
            signed,
            source: Some(source),
            editor,
        })
    }

    /// Materialize the owned property-set source and editor on first use.
    ///
    /// The locals are installed only after both parses succeed, so a failed
    /// admission can be retried without leaving half-built staging state.
    fn ensure_ready(&mut self) -> Result<()> {
        if self.source.is_some() {
            return Ok(());
        }
        let base = self.base.as_ref().ok_or(Error::InvalidData(
            "a detached property-set edit lost its source",
        ))?;
        let (_, cfb_limits) =
            base.edit_profile(ObjectLimits::default(), OleFileLimits::default())?;
        let source = Snapshot::from_bytes_with_cfb_limits(base.to_bytes()?, cfb_limits)?;
        if self.signed {
            self.source = Some(source);
            return Ok(());
        }
        let editor =
            PropertySetEdit::new(source.bytes.to_vec(), cfb_limits).map_err(crate::Error::from)?;
        self.source = Some(source);
        self.editor = Some(editor);
        Ok(())
    }

    fn editor_ref(&mut self) -> Result<&PropertySetEdit> {
        self.ensure_ready()?;
        self.check_signed_source()?;
        self.editor
            .as_ref()
            .ok_or(Error::Reader(crate::Error::Signed))
    }

    fn editor_mut_ready(&mut self) -> Result<&mut PropertySetEdit> {
        self.ensure_ready()?;
        self.check_signed_source()?;
        self.editor
            .as_mut()
            .ok_or(Error::Reader(crate::Error::Signed))
    }

    /// Recheck a dynamic original source before a signed edit reports its
    /// policy refusal. This is a version-only fence; full bytes are checked
    /// once at publication by [`Self::check_source`].
    fn check_signed_source(&self) -> Result<()> {
        if !self.signed {
            return Ok(());
        }
        let Some(base) = &self.base else {
            return Ok(());
        };
        base.ensure_source_fence()
    }

    /// The shared property-set editor, for the package metadata transaction.
    pub(crate) fn editor(&mut self) -> Result<&PropertySetEdit> {
        self.editor_ref()
    }

    /// Mutable access to the shared editor, for the package metadata
    /// transaction.
    pub(crate) fn editor_mut(&mut self) -> Result<&mut PropertySetEdit> {
        self.editor_mut_ready()
    }

    /// Returns the immutable package source used for conflict checks.
    ///
    /// # Errors
    ///
    /// Returns an error when the deferred source cannot be materialized.
    pub fn before(&mut self) -> Result<&Snapshot> {
        self.ensure_ready()?;
        self.source
            .as_ref()
            .ok_or(Error::InvalidData("a property-set edit lost its source"))
    }

    /// The materialized read snapshot, retained for signed no-op admission
    /// and package metadata projection.
    pub(crate) fn snapshot(&mut self) -> Result<&Snapshot> {
        self.before()
    }

    /// Refuse to publish a transaction whose package source changed after
    /// its first materialization.
    pub(crate) fn check_source(&self) -> Result<()> {
        let (Some(base), Some(source)) = (&self.base, &self.source) else {
            return Ok(());
        };
        if base.to_bytes()?.as_slice() != source.bytes() {
            return Err(Error::Reader(crate::Error::PatchConflict {
                operation: "edit property sets",
                reason: "the package source changed while the edit was staged",
            }));
        }
        Ok(())
    }

    /// Applies a typed PIDSI edit, staging only an actual section change.
    ///
    /// A missing or blank section (no property, not even a code page) starts
    /// from an empty one with the Windows-1252 code page.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn edit_summary_information<F>(&mut self, edit: F) -> Result<()>
    where
        F: FnOnce(
            &mut litchi_ole_common::property_set::summary_information::Edit<'_>,
        ) -> std::result::Result<(), OleError>,
    {
        let section = self
            .editor_ref()?
            .property_set(Binding::SummaryInformation)?;
        let snapshot = match section {
            Some(section) if !section.is_empty() => {
                litchi_ole_common::property_set::summary_information::Snapshot::from_section(
                    &section,
                )?
            },
            _ => litchi_ole_common::property_set::summary_information::Snapshot::new(
                CodePage::WINDOWS_1252,
            )?,
        };
        let mut typed = snapshot.edit()?;
        edit(&mut typed)?;
        let commit = typed.commit()?;
        if commit.diagnostics().changed() {
            self.editor_mut_ready()?
                .replace(Binding::SummaryInformation, commit.into_package())?;
        }
        Ok(())
    }

    /// Applies a typed PIDDSI edit, staging only an actual section change.
    ///
    /// A missing or blank section starts from an empty one with the
    /// Windows-1252 code page; a user-defined section already in the stream
    /// is kept.
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
            .editor_ref()?
            .property_set(Binding::DocumentSummaryInformation)?;
        let snapshot = match section {
            Some(section) if !section.is_empty() => {
                litchi_ole_common::property_set::document_summary::Snapshot::from_section(&section)?
            },
            _ => litchi_ole_common::property_set::document_summary::Snapshot::new(
                CodePage::WINDOWS_1252,
            )?,
        };
        let mut typed = snapshot.edit()?;
        edit(&mut typed)?;
        let commit = typed.commit()?;
        if commit.diagnostics().changed() {
            self.editor_mut_ready()?
                .replace(Binding::DocumentSummaryInformation, commit.into_package())?;
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
        F: FnOnce(&mut Section) -> std::result::Result<(), litchi_ole_common::Error>,
    {
        let current = self
            .editor_ref()?
            .property_set(Binding::UserDefinedProperties)?
            .unwrap_or_else(|| {
                let mut section = Section::new(USER_DEFINED_PROPERTIES_FMTID);
                section.set_page(CodePage::WINDOWS_1252);
                section
            });
        let mut candidate = current.clone();
        edit(&mut candidate)?;
        if candidate != current {
            self.editor_mut_ready()?
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
        Ok(self
            .editor_mut_ready()?
            .remove(Binding::UserDefinedProperties)?)
    }

    /// Validate and publish the staged edits as a new package.
    ///
    /// # Errors
    ///
    /// Returns an error when the edited compound file cannot be rendered or
    /// republished.
    pub fn commit(self) -> Result<Commit> {
        let base = self.base.clone().ok_or(Error::InvalidData(
            "a detached property-set edit has no package",
        ))?;
        let (patch, changed) = self.finish()?;
        if changed && base.is_signed() {
            return Err(Error::Reader(crate::Error::Signed));
        }
        let package = if changed {
            crate::package::republish(&base, patch.after.to_vec())?
        } else {
            base
        };
        Ok(Commit::new(package, patch, Diagnostics::new(changed)))
    }

    /// Render the staged edits into a whole-package patch without
    /// republishing it (the metadata family republishes it itself).
    pub(crate) fn finish(self) -> Result<(Patch, bool)> {
        let mut this = self;
        this.ensure_ready()?;
        this.check_source()?;
        let source = this
            .source
            .take()
            .ok_or(Error::InvalidData("a property-set edit lost its source"))?;
        let result = (|| {
            let editor = match this.editor.take() {
                Some(editor) => editor,
                None => {
                    let before = source.bytes().to_vec();
                    return Ok((Patch::new(before.clone(), before), false));
                },
            };
            let before = source.bytes().to_vec();
            let bytes = editor.finish().map_err(crate::Error::from)?;
            let changed = bytes != before;
            Ok((Patch::new(before, bytes), changed))
        })();
        // Rendering can fail after the leading source check. Prefer a
        // concurrent source change over either a render error or success.
        if let Some(base) = this.base.as_ref() {
            base.ensure_source_fence()?;
        }
        result
    }
}

/// A published property-set edit: the new package, its whole-package patch
/// and diagnostics.
pub type Commit = litchi_core::edit::Commit<Package, Patch>;

/// Content-free diagnostics of one property-set commit.
pub use litchi_core::edit::Diagnostics;

/// The property-set family has no caller-selected limit dimension: the
/// republished package is admitted under the source package's limits.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {}

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

    /// Returns the exact inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: Arc::clone(&self.after),
            after: Arc::clone(&self.before),
        }
    }
}

impl Package {
    /// The property-set sections of this package.
    ///
    /// # Errors
    ///
    /// Returns an error when the compound file or a property set is
    /// malformed.
    pub fn property_sets(&self) -> Result<Snapshot> {
        let (_, cfb_limits) =
            self.edit_profile(ObjectLimits::default(), OleFileLimits::default())?;
        Snapshot::from_bytes_with_cfb_limits(self.to_bytes()?, cfb_limits)
    }

    /// Start a property-set edit of this package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::EncryptionPolicy`] for an encrypted workbook. A
    /// compound-file error is returned by the first property-set verb or
    /// commit, after the deferred source is materialized.
    pub fn edit_property_sets(&self) -> Result<Edit> {
        Edit::new(self)
    }

    /// Replay a property-set patch onto the exact package it was made from.
    ///
    /// # Errors
    ///
    /// Returns a patch conflict when this package is not the patch's source.
    pub fn apply_property_sets(&self, patch: &Patch) -> Result<Commit> {
        if self.to_bytes()?.as_slice() != patch.before() {
            return Err(Error::Reader(crate::Error::PatchConflict {
                operation: "edit property sets",
                reason: "the property-set patch was made from another package",
            }));
        }
        let changed = !patch.is_noop();
        if changed && self.is_signed() {
            return Err(Error::Reader(crate::Error::Signed));
        }
        let package = if changed {
            crate::package::republish(self, patch.after.to_vec())?
        } else {
            self.clone()
        };
        Ok(Commit::new(
            package,
            patch.clone(),
            Diagnostics::new(changed),
        ))
    }
}

fn validate_host<R: Read + Seek>(ole: &mut OleFile<R>) -> Result<()> {
    match ole.open_stream(&["Workbook"]) {
        Ok(_) => Ok(()),
        Err(OleError::StreamNotFound) => match ole.open_stream(&["Book"]) {
            Ok(_) => Ok(()),
            Err(OleError::StreamNotFound) => Err(Error::Reader(crate::error::Error::InvalidData(
                "XLS package has no Workbook or Book stream".to_string(),
            ))),
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

        let mut transaction = Edit::detached(snapshot).unwrap();
        transaction
            .edit_summary_information(|edit| edit.set_title("after"))
            .unwrap();
        transaction
            .edit_document_summary_information(|edit| edit.set_company("after company"))
            .unwrap();
        let (patch, changed) = transaction.finish().unwrap();
        assert!(changed);
        let edited = Snapshot::from_bytes(patch.after().to_vec()).unwrap();
        assert_eq!(edited.summary_information().unwrap().title(), Some("after"));

        let mut ole = OleFile::open(
            Cursor::new(patch.after().to_vec()),
            OleFileLimits::default(),
        )
        .unwrap();
        assert_eq!(
            ole.open_stream(&["Unrelated", "Nested"]).unwrap(),
            b"untouched"
        );
        assert_eq!(patch.before(), source.as_slice());
        assert_eq!(patch.inverse().after(), source.as_slice());
    }

    #[test]
    fn no_op_edits_return_the_exact_source_bytes() {
        let source = package_bytes(None);
        let snapshot = Snapshot::from_bytes(source.clone()).unwrap();
        let mut transaction = Edit::detached(snapshot).unwrap();
        transaction.edit_summary_information(|_| Ok(())).unwrap();
        let (patch, changed) = transaction.finish().unwrap();
        assert!(!changed);
        assert_eq!(patch.after(), source.as_slice());
    }

    #[test]
    fn blank_standard_sections_read_as_absent_and_are_replaced_by_edits() {
        // The shape the shared editor used to give a `DocumentSummaryInformation`
        // stream it created for a user-defined edit: a blank standard section
        // (no code page) followed by the user-defined section.
        let mut user_defined = Section::new(USER_DEFINED_PROPERTIES_FMTID);
        user_defined.set_page(CodePage::WINDOWS_1252);
        user_defined
            .add_named(
                2,
                "Department".to_string(),
                Value::Lpstr("Research".to_string()),
            )
            .unwrap();
        let mut document_summary =
            PropertyStream::new(Section::new(DOCUMENT_SUMMARY_INFORMATION_FMTID));
        document_summary.add_section(user_defined).unwrap();
        let document_summary = document_summary.to_bytes().unwrap();
        let summary = PropertyStream::new(Section::new(SUMMARY_INFORMATION_FMTID))
            .to_bytes()
            .unwrap();

        let mut writer = OleWriter::new();
        writer.create_stream(&["Workbook"], b"xls").unwrap();
        writer
            .create_stream(&["\u{0005}SummaryInformation"], &summary)
            .unwrap();
        writer
            .create_stream(&["\u{0005}DocumentSummaryInformation"], &document_summary)
            .unwrap();
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output).unwrap();

        let snapshot = Snapshot::from_bytes(output.into_inner()).unwrap();
        assert!(snapshot.summary_information().is_none());
        assert!(snapshot.document_summary_information().is_none());
        assert!(snapshot.user_defined_properties().is_some());

        let mut transaction = Edit::detached(snapshot).unwrap();
        transaction
            .edit_summary_information(|edit| edit.set_title("title"))
            .unwrap();
        transaction
            .edit_document_summary_information(|edit| edit.set_company("company"))
            .unwrap();
        let (patch, _) = transaction.finish().unwrap();
        let edited = &Snapshot::from_bytes(patch.after().to_vec()).unwrap();
        assert_eq!(edited.summary_information().unwrap().title(), Some("title"));
        let document_summary = edited.document_summary_information().unwrap();
        assert_eq!(document_summary.company(), Some("company"));
        assert_eq!(document_summary.codepage(), Some(CodePage::WINDOWS_1252));
        assert!(
            edited
                .user_defined_properties()
                .unwrap()
                .find_named("Department")
                .is_some()
        );
    }

    #[test]
    fn signed_sources_allow_exact_noops_but_refuse_changes() {
        let source = package_bytes(Some("DigitalSignature"));
        let transaction = Edit::detached(Snapshot::from_bytes(source.clone()).unwrap()).unwrap();
        let (patch, changed) = transaction.finish().unwrap();
        assert!(!changed);
        assert_eq!(patch.before(), source.as_slice());
        assert_eq!(patch.after(), source.as_slice());

        let mut transaction = Edit::detached(Snapshot::from_bytes(source).unwrap()).unwrap();
        assert!(matches!(
            transaction.edit_summary_information(|edit| edit.set_title("changed")),
            Err(Error::Reader(crate::Error::Signed))
        ));

        assert!(
            Snapshot::from_bytes(package_bytes(Some("EncryptedPackage")))
                .map(Edit::detached)
                .unwrap()
                .is_err()
        );
    }

    #[test]
    fn package_edit_entry_keeps_source_and_editor_cold() {
        let package = crate::Builder::new().build().unwrap();
        let edit = Edit::new(&package).unwrap();
        assert!(edit.source.is_none());
        assert!(edit.editor.is_none());
    }
}
