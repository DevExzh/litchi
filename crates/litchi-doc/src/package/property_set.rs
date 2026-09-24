//! Source-preserving property-set transactions for a complete DOC package.
//!
//! The OLE property-set grammar and typed PID owners live in
//! `litchi-ole-common`. This module only supplies the DOC package boundary:
//! host validation, owned source bytes, and source-checked whole-package
//! commits. A generic `Package<R>` cannot safely mutate in place because its
//! reader is owned by the parsed CFB file and the original artifact bytes are
//! not recoverable from that API.

use super::{Error, Result};
use crate::parts::fib::FileInformationBlock;
use crate::parts::protection::{EditProtection, ProtectionPolicy, classify};
use crate::user_defined_hyperlinks::{Limits, MutationError, UserDefinedHyperlinks};
use litchi_cfb::{OleError, OleFile};
use litchi_ole_common::property_set::{
    self, Binding, PropertySetReader, Section, Stream, USER_DEFINED_PROPERTIES_FMTID,
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
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
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

    /// Parses a borrowed DOC artifact while retaining an owned source copy.
    pub fn parse(bytes: &[u8]) -> Result<Self> {
        Self::from_bytes(bytes.to_vec())
    }

    /// Returns the exact source artifact bytes.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.bytes
    }

    /// Projects the standard `SummaryInformation` section, when present.
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

    /// Reads the inert VBA `DigSigBlob` stored in PIDDSI
    /// `DocumentSummaryInformation`.
    ///
    /// The returned owner validates the `[MS-OSHARED]` container and retains
    /// exact source bytes. The PKCS#7 `SignedData` payload and its
    /// `SpcIndirectDataContent`/`SpcIndirectDataContentV2` `contentInfo` form
    /// remain opaque. It never verifies certificate trust, opens a VBA
    /// project, or executes code. `None` means the property stream or the
    /// `DigitalSignature` property is absent.
    pub fn vba_signature(&self) -> Result<Option<vba_signature::Snapshot>> {
        self.vba_signature_with(vba_signature::Limits::default())
    }

    /// Reads the inert VBA signature with explicit payload and blob limits.
    pub fn vba_signature_with(
        &self,
        limits: vba_signature::Limits,
    ) -> Result<Option<vba_signature::Snapshot>> {
        let Some(summary) = self.document_summary_information()? else {
            return Ok(None);
        };
        summary.vba_signature_with(limits).map_err(Into::into)
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
        self.transaction_with_policy(ProtectionPolicy::default())
    }

    /// Starts a property-set transaction with an explicit protected-edit policy.
    pub fn transaction_with_policy(&self, policy: ProtectionPolicy) -> Result<Transaction> {
        Transaction::new(self.clone(), policy)
    }

    /// Consumes the snapshot into its owned source bytes.
    #[must_use]
    pub fn into_bytes(self) -> Vec<u8> {
        self.bytes.to_vec()
    }
}

/// An isolated DOC property-set transaction over owned CFB bytes.
pub struct Transaction {
    source: Snapshot,
    editor: property_set::Editor,
    changed: bool,
    protection_policy: ProtectionPolicy,
}

impl Transaction {
    fn new(source: Snapshot, protection_policy: ProtectionPolicy) -> Result<Self> {
        let editor = property_set::Editor::new(source.bytes.to_vec())?;
        Ok(Self {
            source,
            editor,
            changed: false,
            protection_policy,
        })
    }

    /// Whether a successful edit has changed the transaction state.
    #[must_use]
    pub const fn is_changed(&self) -> bool {
        self.changed
    }

    /// Projects the current transaction-local `SummaryInformation` section.
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

    /// Reads the current transaction-local inert VBA `DigSigBlob`.
    pub fn vba_signature(&self) -> Result<Option<vba_signature::Snapshot>> {
        self.vba_signature_with(vba_signature::Limits::default())
    }

    /// Reads the current transaction-local VBA signature with explicit limits.
    pub fn vba_signature_with(
        &self,
        limits: vba_signature::Limits,
    ) -> Result<Option<vba_signature::Snapshot>> {
        let Some(summary) = self.document_summary_information()? else {
            return Ok(None);
        };
        summary.vba_signature_with(limits).map_err(Into::into)
    }

    /// Replaces PIDDSI `DigitalSignature` with an already validated inert
    /// `DigSigBlob` snapshot. Its nested PKCS#7 `SignedData` and
    /// `contentInfo` form remain opaque to this package owner.
    pub fn set_vba_signature(&mut self, signature: &vba_signature::Snapshot) -> Result<bool> {
        self.edit_document_summary_information(|edit| edit.set_vba_signature(signature))
    }

    /// Removes PIDDSI `DigitalSignature`, when the property is present.
    pub fn remove_vba_signature(&mut self) -> Result<bool> {
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

    /// Edits the opaque PKCS#7 signature and certificate-store payloads
    /// atomically. ASN.1 content form and certificate trust remain outside
    /// this inert storage owner.
    ///
    /// The nested transaction rewrites only those payloads and preserves
    /// producer-specific gaps and padding. Its closure never receives a
    /// cryptographic verifier or a VBA execution context.
    pub fn edit_vba_signature<F>(&mut self, edit: F) -> Result<bool>
    where
        F: FnOnce(&mut vba_signature::Transaction) -> std::result::Result<(), vba_signature::Error>,
    {
        self.edit_vba_signature_with_limits(vba_signature::Limits::default(), edit)
    }

    /// Edits the opaque VBA signature using explicit nested-blob limits.
    pub fn edit_vba_signature_with_limits<F>(
        &mut self,
        limits: vba_signature::Limits,
        edit: F,
    ) -> Result<bool>
    where
        F: FnOnce(&mut vba_signature::Transaction) -> std::result::Result<(), vba_signature::Error>,
    {
        let signature = self.vba_signature_with(limits)?.ok_or_else(|| {
            Error::InvalidFormat(
                "DocumentSummaryInformation has no VBA DigitalSignature blob".to_string(),
            )
        })?;
        let mut transaction = signature.edit();
        edit(&mut transaction).map_err(|error| {
            Error::InvalidFormat(format!("invalid VBA signature edit: {error}"))
        })?;
        let commit = transaction.commit().map_err(|error| {
            Error::InvalidFormat(format!("invalid VBA signature edit: {error}"))
        })?;
        self.set_vba_signature(commit.snapshot())
    }

    /// Returns the current transaction-local user-defined section.
    pub fn user_defined_properties(&self) -> Result<Option<Section>> {
        self.editor
            .property_set(Binding::UserDefinedProperties)
            .map_err(Into::into)
    }

    /// Applies a typed `SummaryInformation` edit through the common PIDSI owner.
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

    /// Applies a typed `DocumentSummaryInformation` edit through the common PIDDSI owner.
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
    /// `DocumentSummaryInformation` stream.
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
        let changed = bytes.as_slice() != before.as_ref();
        if changed {
            self.protection_policy.authorize(self.source.protection)?;
        }
        // The editor can finish without staged changes, or a sequence of
        // edits can restore the complete source artifact byte-for-byte.  The
        // source was already validated at transaction creation, so retain
        // that exact allocation instead of reparsing and allocating a new
        // snapshot for the no-op result.
        let snapshot = if bytes.as_slice() == before.as_ref() {
            Snapshot {
                bytes: Arc::clone(&before),
                protection: self.source.protection,
            }
        } else {
            Snapshot::from_bytes(bytes)?
        };
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

/// An owned result of a DOC property-set transaction.
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

/// A source-checked reversible whole-DOC package patch.
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
    pub fn apply(&self, source: &Snapshot) -> Result<Snapshot> {
        // The destination owns the default publication policy. Reusing the
        // policy captured by an authorized transaction would let a protected
        // patch cross into an enforcing snapshot without an explicit caller
        // authorization at this boundary.
        self.apply_with_policy(source, ProtectionPolicy::default())
    }

    /// Applies the patch with an explicit protected-edit policy.
    pub fn apply_with_policy(
        &self,
        source: &Snapshot,
        policy: ProtectionPolicy,
    ) -> Result<Snapshot> {
        if source.bytes.as_ref() != self.source.as_ref() {
            return Err(Error::InvalidFormat(
                "DOC property-set patch source does not match".to_string(),
            ));
        }
        if source.bytes.as_ref() != self.replacement.as_ref() {
            policy.authorize(source.protection)?;
        }
        Snapshot::from_bytes(self.replacement.to_vec())
    }

    /// Reverts the patch only from its exact replacement snapshot.
    pub fn revert(&self, replacement: &Snapshot) -> Result<Snapshot> {
        self.revert_with_policy(replacement, ProtectionPolicy::default())
    }

    /// Reverts the patch with an explicit protected-edit policy.
    pub fn revert_with_policy(
        &self,
        replacement: &Snapshot,
        policy: ProtectionPolicy,
    ) -> Result<Snapshot> {
        if replacement.bytes.as_ref() != self.replacement.as_ref() {
            return Err(Error::InvalidFormat(
                "DOC property-set patch replacement does not match".to_string(),
            ));
        }
        if replacement.bytes.as_ref() != self.source.as_ref() {
            policy.authorize(replacement.protection)?;
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

fn classify_host(bytes: &[u8]) -> Result<EditProtection> {
    let mut ole = OleFile::open(Cursor::new(bytes))?;
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
        Err(Error::Corrupted(_)) => Ok(EditProtection::Unrecognized),
        Err(error) => Err(error),
    }
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
        let source_pointer = snapshot.bytes().as_ptr();
        let mut transaction = snapshot.transaction().unwrap();
        assert!(!transaction.edit_summary_information(|_| Ok(())).unwrap());
        let commit = transaction.commit().unwrap();
        assert!(!commit.changed());
        assert_eq!(commit.snapshot().bytes().as_ptr(), source_pointer);
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

    #[test]
    fn protected_property_edits_require_authorization_and_keep_patch_capability() {
        let source = protected_package_bytes();
        let snapshot = Snapshot::from_bytes(source.clone()).unwrap();
        assert_eq!(snapshot.protection(), EditProtection::Document);

        let no_op = snapshot.transaction().unwrap().commit().unwrap();
        assert!(!no_op.changed());
        assert_eq!(no_op.snapshot().bytes(), source.as_slice());

        let mut denied = snapshot.transaction().unwrap();
        denied
            .edit_summary_information(|edit| edit.set_title("after"))
            .unwrap();
        assert!(matches!(
            denied.commit(),
            Err(Error::ProtectionDenied(EditProtection::Document))
        ));

        let authorization =
            ProtectionAuthorization::audited("test-suite", "approved metadata repair").unwrap();
        let policy = ProtectionPolicy::allow_protected(authorization);
        let mut allowed = snapshot.transaction_with_policy(policy.clone()).unwrap();
        allowed
            .edit_summary_information(|edit| edit.set_title("after"))
            .unwrap();
        let commit = allowed.commit().unwrap();
        assert!(commit.changed());
        assert!(matches!(
            commit.patch().apply(&snapshot),
            Err(Error::ProtectionDenied(EditProtection::Document))
        ));
        let applied = commit
            .patch()
            .apply_with_policy(&snapshot, policy.clone())
            .unwrap();
        assert_eq!(applied.bytes(), commit.snapshot().bytes());
        assert!(matches!(
            commit.patch().revert(&applied),
            Err(Error::ProtectionDenied(EditProtection::Document))
        ));
        assert_eq!(
            commit
                .patch()
                .revert_with_policy(&applied, policy)
                .unwrap()
                .bytes(),
            source
        );
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
        let mut transaction = snapshot.transaction().unwrap();
        transaction
            .edit_user_defined_properties(|section| {
                section.add(2, Value::Lpwstr("blocked".to_string()))?;
                Ok(())
            })
            .unwrap();
        assert!(matches!(
            transaction.commit(),
            Err(Error::ProtectionDenied(EditProtection::Unknown))
        ));
        let authorization =
            ProtectionAuthorization::audited("test-suite", "inspect malformed host").unwrap();
        let mut transaction = snapshot
            .transaction_with_policy(ProtectionPolicy::allow_protected(authorization))
            .unwrap();
        transaction
            .edit_user_defined_properties(|section| {
                section.add(2, Value::Lpwstr("still blocked".to_string()))?;
                Ok(())
            })
            .unwrap();
        assert!(matches!(
            transaction.commit(),
            Err(Error::ProtectionDenied(EditProtection::Unknown))
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
