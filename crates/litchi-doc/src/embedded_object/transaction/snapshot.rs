//! Immutable DOC/ObjectPool snapshots.

use super::super::Limits;
use super::super::model::{Editor, Inventory, Reference, WriteOptions};
use super::Transaction;
use crate::package::Result;
use crate::parts::protection::ProtectionPolicy;
use litchi_ole_common::ole_streams::{self, NativeSnapshot, PresentationSnapshot};
use std::sync::{Arc, OnceLock};

/// Shared physical DOC bytes with a lazy slice-shaped public handle.
///
/// The common CFB owner retains the admitted input as `Arc<Vec<u8>>`. Keeping
/// that owner here avoids another payload-sized allocation on every successful
/// DOC open. The public `Arc<[u8]>` view is materialized only when a caller
/// first requests either source-view accessor and is then shared by every
/// clone of this owner.
struct SourceOwner {
    bytes: Arc<Vec<u8>>,
    shared: OnceLock<Arc<[u8]>>,
}

impl SourceOwner {
    fn new(bytes: Arc<Vec<u8>>) -> Self {
        Self {
            bytes,
            shared: OnceLock::new(),
        }
    }

    fn as_bytes(&self) -> &[u8] {
        self.bytes.as_slice()
    }

    fn shared_bytes(&self) -> &Arc<[u8]> {
        self.shared
            .get_or_init(|| Arc::<[u8]>::from(self.as_bytes()))
    }

    fn bytes_shared(&self) -> Arc<[u8]> {
        Arc::clone(self.shared_bytes())
    }
}

/// An immutable, source-preserving snapshot of one bounded Word binary file.
///
/// The snapshot owns the exact input bytes for no-op publication and keeps a
/// validated DOC/ObjectPool editor for typed, inert projections. Cloning the
/// snapshot never activates an embedded object, follows a moniker, or copies
/// any payload through an execution/runtime boundary.
#[derive(Clone)]
pub struct Snapshot {
    source: Arc<SourceOwner>,
    limits: Limits,
    editor: Editor,
}

impl Snapshot {
    /// Opens and validates a bounded DOC/ObjectPool artifact.
    ///
    /// Managed field references are checked against the owning `ObjectPool`
    /// storage before the snapshot is published. Unknown streams and opaque
    /// binary descendants are retained by the underlying CFB editor.
    /// # Errors
    ///
    /// Returns an error when the DOC FIB, field tables, CFB package, resource
    /// bounds, or `ObjectPool` ownership references are invalid.
    pub fn open(input: impl Into<Vec<u8>>, limits: Limits) -> Result<Self> {
        Self::open_with_policy(input, limits, ProtectionPolicy::default())
    }

    /// Opens and validates a snapshot with an explicit protected-edit policy.
    pub fn open_with_policy(
        input: impl Into<Vec<u8>>,
        limits: Limits,
        protection_policy: ProtectionPolicy,
    ) -> Result<Self> {
        let bytes = input.into();
        let editor = Editor::open_with_policy(bytes, limits, protection_policy)?;
        editor.validate_references()?;
        // Retain the common owner's admitted source only after DOC reference
        // validation has completed. SourceOwner keeps the common Vec owner
        // and defers the public slice-shaped handle until either source-view
        // accessor is requested.
        let source = Arc::new(SourceOwner::new(editor.package.source_shared()));
        Ok(Self {
            source,
            limits,
            editor,
        })
    }

    /// Returns the exact DOC bytes captured by this snapshot.
    ///
    /// The first call to either [`Self::bytes`] or [`Self::bytes_shared`]
    /// materializes one shared public source allocation. Later calls reuse
    /// that allocation, so the borrowed bytes and shared handle have the same
    /// byte address across snapshot clones.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        self.source.shared_bytes().as_ref()
    }

    /// Returns shared ownership of the exact source allocation.
    ///
    /// The first call to either [`Self::bytes`] or [`Self::bytes_shared`]
    /// materializes one shared public source allocation. Later calls reuse
    /// that allocation, including calls made through snapshot clones.
    #[must_use]
    pub fn bytes_shared(&self) -> Arc<[u8]> {
        self.source.bytes_shared()
    }

    /// Borrows the retained common source owner for internal comparisons.
    pub(super) fn source_bytes(&self) -> &[u8] {
        self.source.as_bytes()
    }

    /// Returns a stable non-cryptographic source fingerprint.
    ///
    /// Callers that need conflict safety must use
    /// [`crate::embedded_object::Patch::apply`], which
    /// checks both this fingerprint and the complete source bytes.
    #[must_use]
    pub fn fingerprint(&self) -> u64 {
        fingerprint(self.source.as_bytes())
    }

    /// Projects the managed field/`ObjectPool` inventory without activation.
    ///
    /// # Errors
    ///
    /// Returns an error when a managed field reference has no owning storage
    /// or a recognized metadata projection cannot be read.
    pub fn inventory(&self) -> Result<Inventory> {
        self.editor.inventory()
    }

    /// Returns managed references in DOC field order.
    ///
    /// # Errors
    ///
    /// Returns an error when a managed field does not resolve to its owning
    /// `ObjectPool` storage.
    pub fn objects(&self) -> Result<Vec<Reference>> {
        self.editor.objects()
    }

    /// Returns one managed object's inert OLEDS presentation stream.
    ///
    /// The DOC field and `ObjectPool` storage remain owned by this DOC layer;
    /// the stream payload is parsed by the shared bounded OLEDS owner and is
    /// never rendered or activated.
    ///
    /// # Errors
    ///
    /// Returns an error when the storage is not referenced, the stream index
    /// or payload is malformed, or the selected OLEDS limits are exceeded.
    pub fn presentation(
        &self,
        storage_id: u32,
        index: usize,
    ) -> Result<Option<PresentationSnapshot>> {
        self.editor.presentation(storage_id, index)
    }

    /// Returns one managed presentation under explicit OLEDS limits.
    pub fn presentation_with_limits(
        &self,
        storage_id: u32,
        index: usize,
        limits: ole_streams::Limits,
    ) -> Result<Option<PresentationSnapshot>> {
        self.editor
            .presentation_with_limits(storage_id, index, limits)
    }

    /// Returns one presentation selected by a DOC object reference.
    pub fn presentation_for(
        &self,
        reference: &Reference,
        index: usize,
    ) -> Result<Option<PresentationSnapshot>> {
        self.editor.presentation_for(reference, index)
    }

    /// Returns one reference-selected presentation under explicit limits.
    pub fn presentation_for_with_limits(
        &self,
        reference: &Reference,
        index: usize,
        limits: ole_streams::Limits,
    ) -> Result<Option<PresentationSnapshot>> {
        self.editor
            .presentation_for_with_limits(reference, index, limits)
    }

    /// Returns all managed OLEDS presentations in numeric order.
    pub fn presentations(&self, storage_id: u32) -> Result<Vec<(usize, PresentationSnapshot)>> {
        self.editor.presentations(storage_id)
    }

    /// Returns all managed presentations under explicit OLEDS limits.
    pub fn presentations_with_limits(
        &self,
        storage_id: u32,
        limits: ole_streams::Limits,
    ) -> Result<Vec<(usize, PresentationSnapshot)>> {
        self.editor.presentations_with_limits(storage_id, limits)
    }

    /// Returns all presentations selected by a DOC object reference.
    pub fn presentations_for(
        &self,
        reference: &Reference,
    ) -> Result<Vec<(usize, PresentationSnapshot)>> {
        self.editor.presentations_for(reference)
    }

    /// Returns all reference-selected presentations under explicit limits.
    pub fn presentations_for_with_limits(
        &self,
        reference: &Reference,
        limits: ole_streams::Limits,
    ) -> Result<Vec<(usize, PresentationSnapshot)>> {
        self.editor.presentations_for_with_limits(reference, limits)
    }

    /// Returns one managed object's inert OLEDS native-data stream.
    pub fn native(&self, storage_id: u32) -> Result<Option<NativeSnapshot>> {
        self.editor.native(storage_id)
    }

    /// Returns one managed native-data stream under explicit OLEDS limits.
    pub fn native_with_limits(
        &self,
        storage_id: u32,
        limits: ole_streams::Limits,
    ) -> Result<Option<NativeSnapshot>> {
        self.editor.native_with_limits(storage_id, limits)
    }

    /// Returns native data selected by a DOC object reference.
    pub fn native_for(&self, reference: &Reference) -> Result<Option<NativeSnapshot>> {
        self.editor.native_for(reference)
    }

    /// Returns reference-selected native data under explicit OLEDS limits.
    pub fn native_for_with_limits(
        &self,
        reference: &Reference,
        limits: ole_streams::Limits,
    ) -> Result<Option<NativeSnapshot>> {
        self.editor.native_for_with_limits(reference, limits)
    }

    /// Exports one managed object as a bounded inert transfer closure.
    ///
    /// The closure contains the standalone object CFB and its exact preview
    /// block. Source-local CPs, FCs, field PLCFs, and storage names are not
    /// exposed; the receiving DOC owner regenerates them for
    /// `destination_storage_id`.
    pub fn export_for_transfer(
        &self,
        storage_id: u32,
        destination_storage_id: u32,
    ) -> Result<WriteOptions> {
        self.editor
            .export_for_transfer(storage_id, destination_storage_id)
    }

    /// Starts an independent clone-staged transaction.
    #[must_use]
    pub fn edit(&self) -> Transaction {
        Transaction::new(self.clone())
    }

    /// Returns the exact source bytes. A snapshot itself is always a no-op.
    #[must_use]
    pub fn finish(&self) -> Vec<u8> {
        self.source.as_bytes().to_vec()
    }

    pub(in crate::embedded_object) fn limits(&self) -> Limits {
        self.limits
    }

    pub(in crate::embedded_object) fn editor(&self) -> &Editor {
        &self.editor
    }

    pub(in crate::embedded_object) fn protection_policy(&self) -> ProtectionPolicy {
        self.editor.protection_policy.clone()
    }

    pub(in crate::embedded_object) fn authorize_changed(&self) -> Result<()> {
        self.editor
            .protection_policy
            .authorize(self.editor.protection)
    }

    #[cfg(test)]
    fn source_view_cached(&self) -> bool {
        self.source.shared.get().is_some()
    }
}

impl std::fmt::Debug for Snapshot {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Snapshot")
            .field("bytes", &self.source.as_bytes().len())
            .field("fingerprint", &self.fingerprint())
            .field("limits", &self.limits)
            .field("editor", &"validated")
            .finish()
    }
}

impl PartialEq for Snapshot {
    fn eq(&self, other: &Self) -> bool {
        self.source.as_bytes() == other.source.as_bytes()
    }
}

impl Eq for Snapshot {}

pub(super) fn fingerprint(bytes: &[u8]) -> u64 {
    let mut value = 0xcbf2_9ce4_8422_2325u64;
    for byte in bytes {
        value ^= u64::from(*byte);
        value = value.wrapping_mul(0x0000_0100_0000_01b3);
    }
    value
}

#[cfg(test)]
mod tests {
    use super::{Limits, Snapshot, SourceOwner};
    use crate::writer::Writer;
    use std::sync::Arc;

    fn base_doc() -> Vec<u8> {
        let mut writer = Writer::new();
        writer.add_paragraph("source owner regression").unwrap();
        let mut output = std::io::Cursor::new(Vec::new());
        writer.write_to(&mut output).unwrap();
        output.into_inner()
    }

    #[test]
    fn source_owner_defers_and_caches_slice_handle() {
        let bytes = Arc::new(vec![0x10, 0x20, 0x30]);
        let owner = Arc::new(SourceOwner::new(Arc::clone(&bytes)));

        assert!(owner.shared.get().is_none());
        assert_eq!(owner.as_bytes(), bytes.as_slice());
        assert!(owner.shared.get().is_none());

        let clone = Arc::clone(&owner);
        let first = owner.shared_bytes();
        let second = clone.bytes_shared();
        assert_eq!(first.as_ref(), bytes.as_slice());
        assert!(Arc::ptr_eq(first, &second));
        assert!(owner.shared.get().is_some());
    }

    #[test]
    fn internal_source_views_do_not_materialize_public_cache() {
        let source = Snapshot::open(base_doc(), Limits::default()).unwrap();
        assert!(!source.source_view_cached());

        let candidate = source.edit().snapshot().unwrap();
        assert_eq!(candidate, source);
        assert!(!source.source_view_cached());

        let transaction = source.edit();
        assert!(!transaction.is_changed().unwrap());
        assert!(!source.source_view_cached());

        let commit = source.edit().commit().unwrap();
        assert!(!commit.changed());
        assert!(!source.source_view_cached());
        assert!(!commit.snapshot().source_view_cached());

        let applied = commit.patch().apply(&source).unwrap();
        assert_eq!(applied, source);
        assert!(!source.source_view_cached());
        assert!(!applied.source_view_cached());

        let shared = source.bytes_shared();
        assert!(source.source_view_cached());
        assert_eq!(shared.as_ptr(), source.bytes().as_ptr());
    }
}
