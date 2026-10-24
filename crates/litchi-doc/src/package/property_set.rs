//! The DOC property-set edit family (record 0351 §2.3, record 0359): typed
//! edits of the OLE `SummaryInformation`, `DocumentSummaryInformation` and
//! user-defined property sets of one [`crate::Package`].
//!
//! [`crate::Package::property_sets`] reads the sections,
//! [`crate::Package::edit_property_sets`] stages edits on an [`Edit`] and
//! [`Edit::commit`] publishes a new [`crate::Package`] with an exact
//! whole-package [`Patch`] and [`Diagnostics`];
//! [`crate::Package::apply_property_sets`] replays the patch. The OLE
//! property-set grammar and typed PID owners live in `litchi-ole-common`;
//! this module supplies the DOC package boundary: host validation, owned
//! source bytes, protection and source-checked publication. A generic `raw::Reader<R>` cannot safely mutate in place because its
//! reader is owned by the parsed CFB file and the original artifact bytes are
//! not recoverable from that API.

use super::{Error as PackageError, Result as PackageResult};
use crate::parts::fib::FileInformationBlock;
use crate::parts::protection::{EditProtection, ProtectionPolicy, classify};
use crate::user_defined_hyperlinks::{Limits, MutationError, UserDefinedHyperlinks};
use litchi_cfb::{OleError, OleFile, OleFileLimits};
use litchi_ole_common::property_set::{
    self, Binding, CodePage, PropertySetReader, Section, Stream, USER_DEFINED_PROPERTIES_FMTID,
};
use litchi_ole_common::vba_signature;
use std::io::Cursor;
use std::sync::Arc;

/// An immutable DOC package source for property-set edits.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Snapshot {
    bytes: Arc<[u8]>,
    protection: EditProtection,
}

impl Snapshot {
    /// Parses and validates an owned DOC compound-file artifact.
    pub(crate) fn from_bytes(bytes: Vec<u8>) -> PackageResult<Self> {
        validate_host(&bytes)?;
        let protection = classify_host(&bytes)?;
        Ok(Self {
            bytes: Arc::from(bytes.into_boxed_slice()),
            protection,
        })
    }

    /// Parses a source and returns its observed Word protection state.
    #[must_use]
    pub const fn protection(&self) -> EditProtection {
        self.protection
    }

    /// Returns the exact source artifact bytes.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.bytes
    }

    /// Shared ownership of the exact source allocation.
    pub(crate) fn bytes_shared(&self) -> Arc<[u8]> {
        Arc::clone(&self.bytes)
    }

    /// Projects the standard `SummaryInformation` section, when present.
    pub fn summary_information(
        &self,
    ) -> PackageResult<Option<property_set::summary_information::Snapshot>> {
        let Some(stream) = read_stream(&self.bytes, Binding::SummaryInformation)? else {
            return Ok(None);
        };
        property_set::summary_information::Snapshot::from_stream(&stream)
            .map(Some)
            .map_err(Into::into)
    }

    /// Projects the standard `DocumentSummaryInformation` section, when present.
    pub fn document_summary_information(
        &self,
    ) -> PackageResult<Option<property_set::document_summary::Snapshot>> {
        let Some(stream) = read_stream(&self.bytes, Binding::DocumentSummaryInformation)? else {
            return Ok(None);
        };
        property_set::document_summary::Snapshot::from_stream(&stream)
            .map(Some)
            .map_err(Into::into)
    }

    /// Reads the inert VBA `DigSigBlob` stored in PIDDSI
    /// `DocumentSummaryInformation`.
    ///
    /// The returned owner validates the `[MS-OSHARED]` container and retains
    /// exact source bytes. The PKCS#7 `SignedData` payload and its
    /// `SpcIndirectDataContent`/`SpcIndirectDataContentV2` `contentInfo` form
    /// remain opaque. It never verifies certificate trust, opens a VBA
    /// project, or executes code. `None` means the property stream or the
    /// `DigitalSignature` property is absent.
    pub fn vba_signature(&self) -> PackageResult<Option<vba_signature::Snapshot>> {
        self.vba_signature_with(vba_signature::Limits::default())
    }

    /// Reads the inert VBA signature with explicit payload and blob limits.
    pub fn vba_signature_with(
        &self,
        limits: vba_signature::Limits,
    ) -> PackageResult<Option<vba_signature::Snapshot>> {
        let Some(summary) = self.document_summary_information()? else {
            return Ok(None);
        };
        summary.vba_signature_with(limits).map_err(Into::into)
    }

    /// Returns the generic user-defined section, when present.
    pub fn user_defined_properties(&self) -> PackageResult<Option<Section>> {
        Ok(read_stream(&self.bytes, Binding::UserDefinedProperties)?
            .and_then(|stream| stream.section(USER_DEFINED_PROPERTIES_FMTID).cloned()))
    }

    /// Reads `_PID_HLINKS` with explicit typed-overlay limits.
    pub fn user_defined_hyperlinks(
        &self,
        fields: Option<&crate::FieldsTable>,
        limits: Limits,
    ) -> PackageResult<Option<UserDefinedHyperlinks>> {
        let Some(section) = self.user_defined_properties()? else {
            return Ok(None);
        };
        crate::user_defined_hyperlinks::from_user_defined_section(&section, fields, limits)
    }

    /// Consumes the snapshot into its owned source bytes.
    #[must_use]
    pub fn into_bytes(self) -> Vec<u8> {
        self.bytes.to_vec()
    }
}

/// The crate-private property-set staging over owned CFB bytes, shared by
/// [`Edit`] and the package metadata family.
pub(crate) struct Staging {
    source: Snapshot,
    editor: property_set::Edit,
    changed: bool,
    protection_policy: ProtectionPolicy,
}

impl Staging {
    pub(crate) fn new(
        source: Snapshot,
        protection_policy: ProtectionPolicy,
    ) -> PackageResult<Self> {
        let editor = property_set::Edit::new(source.bytes.to_vec(), OleFileLimits::default())?;
        Ok(Self {
            source,
            editor,
            changed: false,
            protection_policy,
        })
    }

    /// The shared property-set editor, for the package metadata transaction.
    pub(crate) const fn editor(&self) -> &property_set::Edit {
        &self.editor
    }

    /// Mutable access to the shared editor; staging through it marks the
    /// transaction changed.
    pub(crate) fn editor_mut(&mut self) -> &mut property_set::Edit {
        self.changed = true;
        &mut self.editor
    }

    /// Projects the current transaction-local `SummaryInformation` section.
    pub(crate) fn summary_information(
        &self,
    ) -> PackageResult<Option<property_set::summary_information::Snapshot>> {
        let Some(section) = self.editor.property_set(Binding::SummaryInformation)? else {
            return Ok(None);
        };
        property_set::summary_information::Snapshot::from_section(&section)
            .map(Some)
            .map_err(Into::into)
    }

    /// Projects the current transaction-local `DocumentSummaryInformation` section.
    pub(crate) fn document_summary_information(
        &self,
    ) -> PackageResult<Option<property_set::document_summary::Snapshot>> {
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

    /// Reads the current transaction-local VBA signature with explicit limits.
    pub(crate) fn vba_signature_with(
        &self,
        limits: vba_signature::Limits,
    ) -> PackageResult<Option<vba_signature::Snapshot>> {
        let Some(summary) = self.document_summary_information()? else {
            return Ok(None);
        };
        summary.vba_signature_with(limits).map_err(Into::into)
    }

    /// Replaces PIDDSI `DigitalSignature` with an already validated inert
    /// `DigSigBlob` snapshot. Its nested PKCS#7 `SignedData` and
    /// `contentInfo` form remain opaque to this package owner.
    pub(crate) fn set_vba_signature(
        &mut self,
        signature: &vba_signature::Snapshot,
    ) -> PackageResult<bool> {
        self.edit_document_summary_information(|edit| edit.set_vba_signature(signature))
    }

    /// Removes PIDDSI `DigitalSignature`, when the property is present.
    pub(crate) fn remove_vba_signature(&mut self) -> PackageResult<bool> {
        if self
            .editor
            .property_set(Binding::DocumentSummaryInformation)?
            .is_none()
        {
            return Ok(false);
        }
        self.edit_document_summary_information(|edit| {
            edit.remove_vba_signature();
            Ok(())
        })
    }

    /// Edits the opaque VBA signature using explicit nested-blob limits.
    pub(crate) fn edit_vba_signature<F>(
        &mut self,
        limits: vba_signature::Limits,
        edit: F,
    ) -> PackageResult<bool>
    where
        F: FnOnce(&mut vba_signature::Edit) -> std::result::Result<(), vba_signature::Error>,
    {
        let signature = self.vba_signature_with(limits)?.ok_or_else(|| {
            PackageError::InvalidFormat(
                "DocumentSummaryInformation has no VBA DigitalSignature blob".to_string(),
            )
        })?;
        let mut transaction = signature.edit();
        edit(&mut transaction).map_err(|error| {
            PackageError::InvalidFormat(format!("invalid VBA signature edit: {error}"))
        })?;
        let commit = transaction.commit().map_err(|error| {
            PackageError::InvalidFormat(format!("invalid VBA signature edit: {error}"))
        })?;
        self.set_vba_signature(commit.package())
    }

    /// Returns the current transaction-local user-defined section.
    pub(crate) fn user_defined_properties(&self) -> PackageResult<Option<Section>> {
        self.editor
            .property_set(Binding::UserDefinedProperties)
            .map_err(Into::into)
    }

    /// Applies a typed `SummaryInformation` edit through the common PIDSI owner.
    ///
    /// The common typed transaction is committed before its replacement is
    /// handed to the common CFB editor. A no-op typed commit does not stage a
    /// stream replacement, preserving the complete source artifact byte-for-byte.
    ///
    /// A package without a `\x05SummaryInformation` stream (a file the
    /// [`crate::Builder`] built without properties, for one) starts from an
    /// empty section with the Windows-1252 code page, and the stream is
    /// created when the edit changes it; an existing blank section (no
    /// property, not even a code page) is treated the same way.
    pub(crate) fn edit_summary_information<F>(&mut self, edit: F) -> PackageResult<bool>
    where
        F: for<'a> FnOnce(
            &mut property_set::summary_information::Edit<'a>,
        ) -> std::result::Result<(), OleError>,
    {
        let snapshot = match self.editor.property_set(Binding::SummaryInformation)? {
            Some(source) if !source.is_empty() => {
                property_set::summary_information::Snapshot::from_section(&source)?
            },
            _ => property_set::summary_information::Snapshot::new(CodePage::WINDOWS_1252)?,
        };
        let mut draft = snapshot.edit()?;
        edit(&mut draft)?;
        let commit = draft.commit()?;
        if !commit.diagnostics().changed() {
            return Ok(false);
        }
        self.editor
            .replace(Binding::SummaryInformation, commit.into_package())?;
        self.changed = true;
        Ok(true)
    }

    /// Applies a typed `DocumentSummaryInformation` edit through the common PIDDSI owner.
    ///
    /// A missing or blank section is handled as in
    /// [`Self::edit_summary_information`]; a user-defined section already in
    /// the stream is kept.
    pub(crate) fn edit_document_summary_information<F>(&mut self, edit: F) -> PackageResult<bool>
    where
        F: for<'a> FnOnce(
            &mut property_set::document_summary::Edit<'a>,
        ) -> std::result::Result<(), OleError>,
    {
        let snapshot = match self
            .editor
            .property_set(Binding::DocumentSummaryInformation)?
        {
            Some(source) if !source.is_empty() => {
                property_set::document_summary::Snapshot::from_section(&source)?
            },
            _ => property_set::document_summary::Snapshot::new(CodePage::WINDOWS_1252)?,
        };
        let mut draft = snapshot.edit()?;
        edit(&mut draft)?;
        let commit = draft.commit()?;
        if !commit.diagnostics().changed() {
            return Ok(false);
        }
        self.editor
            .replace(Binding::DocumentSummaryInformation, commit.into_package())?;
        self.changed = true;
        Ok(true)
    }

    /// Applies a generic user-defined property-section edit.
    pub(crate) fn edit_user_defined_properties<F>(&mut self, edit: F) -> PackageResult<bool>
    where
        F: FnOnce(&mut Section) -> std::result::Result<(), litchi_ole_common::Error>,
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
    pub(crate) fn remove_user_defined_properties(&mut self) -> PackageResult<bool> {
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

    /// Replaces `_PID_HLINKS` with explicit shared typed-overlay limits.
    pub(crate) fn put_user_defined_hyperlinks(
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
        .map_err(PackageError::from)
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
    pub(crate) fn remove_user_defined_hyperlinks(&mut self) -> PackageResult<bool> {
        self.edit_user_defined_properties(|section| {
            crate::user_defined_hyperlinks::remove(section, Limits::default())?;
            Ok(())
        })
    }

    /// The staged artifact bytes, authorized by the protection policy when
    /// they differ from the source.
    pub(crate) fn finish(self) -> PackageResult<Vec<u8>> {
        let bytes = self.editor.finish()?;
        if bytes.as_slice() != self.source.bytes.as_ref() {
            self.protection_policy.authorize(self.source.protection)?;
        }
        Ok(bytes)
    }
}

crate::scoped::family!("property set");

/// A staged property-set edit of one package. Dropping it discards it.
pub struct Edit {
    base: crate::scoped::Base,
    inner: Option<Staging>,
}

impl std::fmt::Debug for Edit {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("base", &self.base)
            .finish_non_exhaustive()
    }
}

impl Edit {
    fn open(base: crate::scoped::Base) -> Result<Self> {
        Ok(Self { base, inner: None })
    }

    fn inner_mut(&mut self) -> Result<&mut Staging> {
        if self.inner.is_none() {
            let source = Snapshot::from_bytes(self.base.to_vec()?)?;
            self.inner = Some(Staging::new(
                source,
                self.base.options().protection_policy_value(),
            )?);
        }
        self.inner.as_mut().ok_or_else(|| {
            Error::Package(PackageError::InvalidFormat(
                "property-set edit was not initialized".into(),
            ))
        })
    }

    fn take_inner(&mut self) -> Result<Staging> {
        self.inner_mut()?;
        self.inner.take().ok_or_else(|| {
            Error::Package(PackageError::InvalidFormat(
                "property-set edit was not initialized".into(),
            ))
        })
    }

    /// Projects the staged `SummaryInformation` section.
    ///
    /// # Errors
    ///
    /// Returns an error when the staged stream is not a valid property set.
    pub fn summary_information(
        &mut self,
    ) -> Result<Option<property_set::summary_information::Snapshot>> {
        Ok(self.inner_mut()?.summary_information()?)
    }

    /// Projects the staged `DocumentSummaryInformation` section.
    ///
    /// # Errors
    ///
    /// Returns an error when the staged stream is not a valid property set.
    pub fn document_summary_information(
        &mut self,
    ) -> Result<Option<property_set::document_summary::Snapshot>> {
        Ok(self.inner_mut()?.document_summary_information()?)
    }

    /// Reads the staged inert VBA `DigSigBlob` with explicit limits.
    ///
    /// # Errors
    ///
    /// Returns an error when the staged blob is malformed or over `limits`.
    pub fn vba_signature(
        &mut self,
        limits: vba_signature::Limits,
    ) -> Result<Option<vba_signature::Snapshot>> {
        Ok(self.inner_mut()?.vba_signature_with(limits)?)
    }

    /// Replaces PIDDSI `DigitalSignature` with an already validated inert
    /// `DigSigBlob` snapshot. Returns whether the staged state changed.
    ///
    /// # Errors
    ///
    /// Returns an error when the property set cannot hold the blob.
    pub fn set_vba_signature(&mut self, signature: &vba_signature::Snapshot) -> Result<bool> {
        Ok(self.inner_mut()?.set_vba_signature(signature)?)
    }

    /// Removes PIDDSI `DigitalSignature`, when present. Returns whether the
    /// staged state changed.
    ///
    /// # Errors
    ///
    /// Returns an error when the staged property set is malformed.
    pub fn remove_vba_signature(&mut self) -> Result<bool> {
        Ok(self.inner_mut()?.remove_vba_signature()?)
    }

    /// Edits the opaque VBA signature using explicit nested-blob limits.
    ///
    /// # Errors
    ///
    /// Returns an error when no signature is present or the edit is invalid.
    pub fn edit_vba_signature<F>(&mut self, limits: vba_signature::Limits, edit: F) -> Result<bool>
    where
        F: FnOnce(&mut vba_signature::Edit) -> std::result::Result<(), vba_signature::Error>,
    {
        Ok(self.inner_mut()?.edit_vba_signature(limits, edit)?)
    }

    /// Returns the staged user-defined section.
    ///
    /// # Errors
    ///
    /// Returns an error when the staged stream is not a valid property set.
    pub fn user_defined_properties(&mut self) -> Result<Option<Section>> {
        Ok(self.inner_mut()?.user_defined_properties()?)
    }

    /// Applies a typed `SummaryInformation` edit through the common PIDSI
    /// owner. A missing or blank section starts from an empty Windows-1252
    /// section. Returns whether the staged state changed.
    ///
    /// # Errors
    ///
    /// Returns an error when the edit or the staged property set is invalid.
    pub fn edit_summary_information<F>(&mut self, edit: F) -> Result<bool>
    where
        F: for<'a> FnOnce(
            &mut property_set::summary_information::Edit<'a>,
        ) -> std::result::Result<(), OleError>,
    {
        Ok(self.inner_mut()?.edit_summary_information(edit)?)
    }

    /// Applies a typed `DocumentSummaryInformation` edit through the common
    /// PIDDSI owner. Returns whether the staged state changed.
    ///
    /// # Errors
    ///
    /// Returns an error when the edit or the staged property set is invalid.
    pub fn edit_document_summary_information<F>(&mut self, edit: F) -> Result<bool>
    where
        F: for<'a> FnOnce(
            &mut property_set::document_summary::Edit<'a>,
        ) -> std::result::Result<(), OleError>,
    {
        Ok(self.inner_mut()?.edit_document_summary_information(edit)?)
    }

    /// Applies a generic user-defined property-section edit. Returns whether
    /// the staged state changed.
    ///
    /// # Errors
    ///
    /// Returns an error when the edit or the staged property set is invalid.
    pub fn edit_user_defined_properties<F>(&mut self, edit: F) -> Result<bool>
    where
        F: FnOnce(&mut Section) -> std::result::Result<(), litchi_ole_common::Error>,
    {
        Ok(self.inner_mut()?.edit_user_defined_properties(edit)?)
    }

    /// Removes the user-defined section while retaining the rest of the
    /// `DocumentSummaryInformation` stream. Returns whether it was present.
    ///
    /// # Errors
    ///
    /// Returns an error when the staged property set is malformed.
    pub fn remove_user_defined_properties(&mut self) -> Result<bool> {
        Ok(self.inner_mut()?.remove_user_defined_properties()?)
    }

    /// Replaces `_PID_HLINKS` with explicit shared typed-overlay limits.
    /// Returns whether the staged state changed.
    ///
    /// # Errors
    ///
    /// Returns a [`MutationError`] for unresolved field candidates, a limit
    /// or an invalid staged property set.
    pub fn put_user_defined_hyperlinks(
        &mut self,
        hyperlinks: &UserDefinedHyperlinks,
        limits: Limits,
    ) -> std::result::Result<bool, MutationError> {
        self.inner_mut()
            .map_err(|error| match error {
                Error::Package(error) => MutationError::Package(error),
                Error::Refused(_) | Error::PatchConflict | Error::Signed => {
                    MutationError::Package(PackageError::InvalidFormat(
                        "property-set hyperlink edit could not be initialized".into(),
                    ))
                },
            })?
            .put_user_defined_hyperlinks(hyperlinks, limits)
    }

    /// Removes only the named `_PID_HLINKS` user-defined property. Returns
    /// whether the staged state changed.
    ///
    /// # Errors
    ///
    /// Returns an error when the staged property set is malformed.
    pub fn remove_user_defined_hyperlinks(&mut self) -> Result<bool> {
        Ok(self.inner_mut()?.remove_user_defined_hyperlinks()?)
    }

    /// Validate the staged edits and publish them as a new package.
    ///
    /// # Errors
    ///
    /// Returns an error when the protection policy refuses a changed write,
    /// the package cannot be republished, or [`Error::Signed`] for a changed
    /// signed package.
    pub fn commit(mut self) -> Result<Commit> {
        let after = self.take_inner()?.finish()?;
        Ok(published(self.base.publish(after)?))
    }
}

impl crate::Package {
    /// The property-set read view of this package.
    ///
    /// # Errors
    ///
    /// Returns an error when the package bytes are not a DOC host.
    pub fn property_sets(&self) -> Result<Snapshot> {
        Ok(Snapshot::from_bytes(self.to_bytes()?)?)
    }

    /// Start a property-set edit of this package under its retained edit
    /// options.
    ///
    /// # Errors
    ///
    /// Returns an error when the package's property sets cannot be read for
    /// editing (for example a signed or encrypted container).
    pub fn edit_property_sets(&self) -> Result<Edit> {
        Edit::open(crate::scoped::Base::capture(self)?)
    }

    /// Start a property-set edit of this package under explicit edit options.
    ///
    /// # Errors
    ///
    /// Returns an error when the package's property sets cannot be read for
    /// editing.
    pub fn edit_property_sets_with_options(&self, options: crate::EditOptions) -> Result<Edit> {
        Edit::open(crate::scoped::Base::capture_with_options(self, options)?)
    }

    /// Replay a property-set patch onto the exact package it was made from.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// source, or [`Error::Signed`] for a changed signed package.
    pub fn apply_property_sets(&self, patch: &Patch) -> Result<Commit> {
        apply_patch(self, patch)
    }
}

fn validate_host(bytes: &[u8]) -> PackageResult<()> {
    let ole = OleFile::open(Cursor::new(bytes), OleFileLimits::default())?;
    if !ole.exists(&["WordDocument"]) {
        return Err(PackageError::InvalidFormat(
            "Not a valid Word document: WordDocument stream not found".to_string(),
        ));
    }
    Ok(())
}

pub(crate) fn classify_host(bytes: &[u8]) -> PackageResult<EditProtection> {
    let mut ole = OleFile::open(Cursor::new(bytes), OleFileLimits::default())?;
    let word = match ole.open_stream(&["WordDocument"]) {
        Ok(word) => word,
        Err(OleError::StreamNotFound) => return Ok(EditProtection::Unknown),
        Err(error) => return Err(error.into()),
    };
    let fib = match FileInformationBlock::parse(&word) {
        Ok(fib) => fib,
        Err(_) => return Ok(EditProtection::Unknown),
    };
    let table_name = if fib.which_table_stream() {
        "1Table"
    } else {
        "0Table"
    };
    let table = match ole.open_stream(&[table_name]) {
        Ok(table) => table,
        Err(OleError::StreamNotFound) => return Ok(EditProtection::Unknown),
        Err(error) => return Err(error.into()),
    };
    // Property-set callers may still inspect a structurally valid legacy host
    // whose protection record the classifier cannot interpret. Keep that
    // source readable, but do not turn an unparseable protection record into
    // `None`: changed publication remains fail-closed as `Unrecognized`,
    // including when a caller supplies the protected-edit capability.
    match classify(&fib, &table) {
        Ok(protection) => Ok(protection),
        Err(PackageError::Corrupted(_)) => Ok(EditProtection::Unrecognized),
        Err(error) => Err(error),
    }
}

fn read_stream(bytes: &[u8], binding: Binding) -> PackageResult<Option<Stream>> {
    let mut ole = OleFile::open(Cursor::new(bytes), OleFileLimits::default())?;
    match ole.property_set(binding) {
        Ok(stream) => Ok(Some(stream)),
        Err(litchi_ole_common::Error::Cfb(OleError::StreamNotFound)) => Ok(None),
        Err(error) => Err(error.into()),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn staging(snapshot: &Snapshot, policy: ProtectionPolicy) -> Staging {
        Staging::new(snapshot.clone(), policy).unwrap()
    }
    use crate::parts::protection::{EditProtection, ProtectionAuthorization, ProtectionPolicy};
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

        const DOP_INDEX: usize = 31;
        const POINTER_COUNT: usize = 136;
        let pointer_end = 154 + POINTER_COUNT * 8;
        let mut word = vec![0u8; pointer_end + 4];
        word[0..2].copy_from_slice(&0xa5ecu16.to_le_bytes());
        // FibBase.csw and cslw are fixed MS-DOC counts.
        word[32..34].copy_from_slice(&0x000eu16.to_le_bytes());
        word[62..64].copy_from_slice(&0x0016u16.to_le_bytes());
        word[2..4].copy_from_slice(&0x00c1u16.to_le_bytes());
        word[152..154].copy_from_slice(&(POINTER_COUNT as u16).to_le_bytes());
        word[pointer_end..pointer_end + 2].copy_from_slice(&2u16.to_le_bytes());
        word[pointer_end + 2..pointer_end + 4].copy_from_slice(&0x0101u16.to_le_bytes());

        let table = crate::parts::document_properties::DocumentProperties::writer_bytes(
            false, false, false, true,
        );
        let dop_pointer = 154 + DOP_INDEX * 8;
        word[dop_pointer..dop_pointer + 4].copy_from_slice(&0u32.to_le_bytes());
        word[dop_pointer + 4..dop_pointer + 8].copy_from_slice(&594u32.to_le_bytes());

        let mut writer = OleWriter::new();
        writer.create_stream(&["WordDocument"], &word).unwrap();
        writer.create_stream(&["0Table"], &table).unwrap();
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

    fn protected_package_bytes() -> Vec<u8> {
        const DOP_INDEX: usize = 31;
        const POINTER_COUNT: usize = 136;
        let dop_offset = 16usize;
        let mut table = vec![0x5a; dop_offset];
        let mut dop = crate::parts::document_properties::DocumentProperties::writer_bytes(
            false, false, false, true,
        );
        dop[6] = 0x10;
        table.extend_from_slice(&dop);

        let pointer_end = 154 + POINTER_COUNT * 8;
        let mut word = vec![0u8; pointer_end + 4];
        word[0..2].copy_from_slice(&0xa5ecu16.to_le_bytes());
        // FibBase.csw and cslw are fixed MS-DOC counts.
        word[32..34].copy_from_slice(&0x000eu16.to_le_bytes());
        word[62..64].copy_from_slice(&0x0016u16.to_le_bytes());
        word[2..4].copy_from_slice(&0x00c1u16.to_le_bytes());
        word[152..154].copy_from_slice(&(POINTER_COUNT as u16).to_le_bytes());
        word[pointer_end..pointer_end + 2].copy_from_slice(&2u16.to_le_bytes());
        word[pointer_end + 2..pointer_end + 4].copy_from_slice(&0x0101u16.to_le_bytes());
        let pointer = 154 + DOP_INDEX * 8;
        word[pointer..pointer + 4].copy_from_slice(&(dop_offset as u32).to_le_bytes());
        word[pointer + 4..pointer + 8].copy_from_slice(&(dop.len() as u32).to_le_bytes());

        let mut summary = Section::new(SUMMARY_INFORMATION_FMTID);
        summary.set_page(CodePage::Utf16Le);
        summary.add(2, Value::Lpwstr("before".to_string())).unwrap();
        let summary = PropertyStream::new(summary).to_bytes().unwrap();

        let mut writer = OleWriter::new();
        writer.create_stream(&["WordDocument"], &word).unwrap();
        writer.create_stream(&["0Table"], &table).unwrap();
        writer
            .create_stream(&["\u{0005}SummaryInformation"], &summary)
            .unwrap();
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

        let mut transaction = staging(&snapshot, ProtectionPolicy::default());
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
        let bytes = transaction.finish().unwrap();
        assert_ne!(bytes, source);
        let edited = Snapshot::from_bytes(bytes).unwrap();
        assert_eq!(
            edited.summary_information().unwrap().unwrap().title(),
            Some("after")
        );

        let mut ole = OleFile::open(
            Cursor::new(edited.bytes().to_vec()),
            OleFileLimits::default(),
        )
        .unwrap();
        assert_eq!(
            ole.open_stream(&["Unrelated", "Nested"]).unwrap(),
            b"untouched"
        );
    }

    #[test]
    fn missing_and_blank_standard_sections_are_created_by_typed_edits() {
        // A document the writer created without properties, given blank
        // standard sections (no property, not even a code page).
        let mut builder = crate::Builder::new();
        builder.add_paragraph("Body").unwrap();
        let written = builder.build().unwrap().to_bytes().unwrap();
        let mut source =
            OleFile::open(Cursor::new(written.clone()), OleFileLimits::default()).unwrap();
        let mut rebuilt = OleWriter::new();
        for path in source.list_streams() {
            let refs: Vec<&str> = path.iter().map(String::as_str).collect();
            let data = source.open_stream(&refs).unwrap();
            rebuilt.create_stream(&refs, &data).unwrap();
        }
        for (name, format_identifier) in [
            ("\u{0005}SummaryInformation", SUMMARY_INFORMATION_FMTID),
            (
                "\u{0005}DocumentSummaryInformation",
                DOCUMENT_SUMMARY_INFORMATION_FMTID,
            ),
        ] {
            let blank = PropertyStream::new(Section::new(format_identifier))
                .to_bytes()
                .unwrap();
            rebuilt.create_stream(&[name], &blank).unwrap();
        }
        let mut output = Cursor::new(Vec::new());
        rebuilt.write_to(&mut output).unwrap();

        for bytes in [written, output.into_inner()] {
            let snapshot = Snapshot::from_bytes(bytes).unwrap();
            let mut transaction = staging(&snapshot, ProtectionPolicy::default());
            assert!(
                transaction
                    .edit_summary_information(|edit| edit.set_title("title"))
                    .unwrap()
            );
            assert!(
                transaction
                    .edit_document_summary_information(|edit| edit.set_company("company"))
                    .unwrap()
            );
            let edited = Snapshot::from_bytes(transaction.finish().unwrap()).unwrap();
            let summary = edited.summary_information().unwrap().unwrap();
            assert_eq!(summary.title(), Some("title"));
            assert_eq!(summary.codepage(), Some(CodePage::WINDOWS_1252));
            let document_summary = edited.document_summary_information().unwrap().unwrap();
            assert_eq!(document_summary.company(), Some("company"));
            assert_eq!(document_summary.codepage(), Some(CodePage::WINDOWS_1252));
        }
    }

    #[test]
    fn no_op_edits_return_the_exact_source_bytes() {
        let source = package_bytes(None);
        let snapshot = Snapshot::from_bytes(source.clone()).unwrap();
        let mut transaction = staging(&snapshot, ProtectionPolicy::default());
        assert!(!transaction.edit_summary_information(|_| Ok(())).unwrap());
        assert_eq!(transaction.finish().unwrap(), source);
    }

    #[test]
    fn signed_and_encrypted_sources_are_refused_before_mutation() {
        assert!(
            Snapshot::from_bytes(package_bytes(Some("DigitalSignature")))
                .map(|snapshot| Staging::new(snapshot, ProtectionPolicy::default()))
                .unwrap()
                .is_err()
        );
        assert!(
            Snapshot::from_bytes(package_bytes(Some("EncryptedPackage")))
                .map(|snapshot| Staging::new(snapshot, ProtectionPolicy::default()))
                .unwrap()
                .is_err()
        );
    }

    #[test]
    fn protected_property_edits_require_authorization_and_keep_patch_capability() {
        let source = protected_package_bytes();
        let snapshot = Snapshot::from_bytes(source.clone()).unwrap();
        assert_eq!(snapshot.protection(), EditProtection::Document);

        let no_op = staging(&snapshot, ProtectionPolicy::default())
            .finish()
            .unwrap();
        assert_eq!(no_op, source);

        let mut denied = staging(&snapshot, ProtectionPolicy::default());
        denied
            .edit_summary_information(|edit| edit.set_title("after"))
            .unwrap();
        assert!(matches!(
            denied.finish(),
            Err(PackageError::ProtectionDenied(EditProtection::Document))
        ));

        let authorization =
            ProtectionAuthorization::audited("test-suite", "approved metadata repair").unwrap();
        let policy = ProtectionPolicy::allow_protected(authorization);
        let mut allowed = staging(&snapshot, policy);
        allowed
            .edit_summary_information(|edit| edit.set_title("after"))
            .unwrap();
        assert_ne!(allowed.finish().unwrap(), source);
    }

    #[test]
    fn malformed_or_incomplete_word_hosts_are_unknown_and_fail_closed() {
        let mut malformed = OleWriter::new();
        malformed
            .create_stream(&["WordDocument"], b"not-a-fib")
            .unwrap();
        let mut output = Cursor::new(Vec::new());
        malformed.write_to(&mut output).unwrap();
        let snapshot = Snapshot::from_bytes(output.into_inner()).unwrap();
        assert_eq!(snapshot.protection(), EditProtection::Unknown);
        let mut transaction = staging(&snapshot, ProtectionPolicy::default());
        transaction
            .edit_user_defined_properties(|section| {
                section.add(2, Value::Lpwstr("blocked".to_string()))?;
                Ok(())
            })
            .unwrap();
        assert!(matches!(
            transaction.finish(),
            Err(PackageError::ProtectionDenied(EditProtection::Unknown))
        ));
        let authorization =
            ProtectionAuthorization::audited("test-suite", "inspect malformed host").unwrap();
        let mut transaction = staging(&snapshot, ProtectionPolicy::allow_protected(authorization));
        transaction
            .edit_user_defined_properties(|section| {
                section.add(2, Value::Lpwstr("still blocked".to_string()))?;
                Ok(())
            })
            .unwrap();
        assert!(matches!(
            transaction.finish(),
            Err(PackageError::ProtectionDenied(EditProtection::Unknown))
        ));

        let mut missing_selected_table = OleWriter::new();
        let mut word = vec![0u8; 154 + 38 * 8];
        word[0..2].copy_from_slice(&0xa5ecu16.to_le_bytes());
        // FibBase.csw and cslw are fixed MS-DOC counts.
        word[32..34].copy_from_slice(&0x000eu16.to_le_bytes());
        word[62..64].copy_from_slice(&0x0016u16.to_le_bytes());
        word[2..4].copy_from_slice(&0x00c1u16.to_le_bytes());
        word[152..154].copy_from_slice(&38u16.to_le_bytes());
        missing_selected_table
            .create_stream(&["WordDocument"], &word)
            .unwrap();
        missing_selected_table
            .create_stream(&["1Table"], &[])
            .unwrap();
        let mut output = Cursor::new(Vec::new());
        missing_selected_table.write_to(&mut output).unwrap();
        let snapshot = Snapshot::from_bytes(output.into_inner()).unwrap();
        assert_eq!(snapshot.protection(), EditProtection::Unknown);
    }
}
