//! Failure-atomic typed edits for opaque VBA signature payloads.

use std::sync::Arc;

use super::codec;
use super::model::Error;
use super::patch::Patch;
use super::snapshot::Snapshot;

/// A deterministic identity for one exact serialized VBA signature blob.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Revision(u64);

impl Revision {
    pub(super) fn of(bytes: &[u8]) -> Self {
        let mut value = 0xcbf2_9ce4_8422_2325u64;
        value ^= bytes.len() as u64;
        value = value.wrapping_mul(0x0000_0100_0000_01b3);
        for byte in bytes {
            value ^= u64::from(*byte);
            value = value.wrapping_mul(0x0000_0100_0000_01b3);
        }
        Self(value)
    }

    /// Returns the raw source fingerprint.
    #[must_use]
    pub const fn value(self) -> u64 {
        self.0
    }

    /// Alias for [`Self::value`].
    #[must_use]
    pub const fn fingerprint(self) -> u64 {
        self.value()
    }
}

/// An isolated typed edit over one immutable VBA signature snapshot.
#[derive(Debug, Clone)]
pub struct Transaction {
    source: Snapshot,
    signature: Vec<u8>,
    certificate_store: Vec<u8>,
}

impl Transaction {
    pub(super) fn new(source: Snapshot) -> Self {
        Self {
            signature: source.info().signature().to_vec(),
            certificate_store: source.info().certificate_store().to_vec(),
            source,
        }
    }

    /// Borrows the immutable source used to start this transaction.
    #[must_use]
    pub const fn source(&self) -> &Snapshot {
        &self.source
    }

    /// Returns the current opaque signature payload draft.
    #[must_use]
    pub fn signature(&self) -> &[u8] {
        &self.signature
    }

    /// Returns the current opaque serialized certificate-store payload draft.
    #[must_use]
    pub fn certificate_store(&self) -> &[u8] {
        &self.certificate_store
    }

    /// Whether either opaque payload differs from the source projection.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.signature != self.source.info().signature()
            || self.certificate_store != self.source.info().certificate_store()
    }

    /// Replaces the opaque signature payload after candidate validation.
    ///
    /// # Errors
    ///
    /// Returns [`Error`] when the payload exceeds its configured bound or
    /// makes the enclosing wire structure unrepresentable.
    pub fn set_signature<P>(&mut self, payload: P) -> Result<&mut Self, Error>
    where
        P: AsRef<[u8]>,
    {
        let payload_ref = payload.as_ref();
        codec::check_payload_size(
            payload_ref.len(),
            "signature",
            self.source.limits().max_signature_bytes,
        )?;
        let mut signature = payload_ref.to_vec();
        let certificate_store = self.certificate_store.clone();
        self.stage(&mut signature, certificate_store)
    }

    /// Alias for [`Self::set_signature`] using replacement terminology.
    ///
    /// # Errors
    ///
    /// Returns the same errors as [`Self::set_signature`].
    pub fn replace_signature<P>(&mut self, payload: P) -> Result<&mut Self, Error>
    where
        P: AsRef<[u8]>,
    {
        self.set_signature(payload)
    }

    /// Replaces the opaque serialized certificate-store payload after
    /// candidate validation.
    ///
    /// # Errors
    ///
    /// Returns [`Error`] when the payload exceeds its configured bound or
    /// makes the enclosing wire structure unrepresentable.
    pub fn set_certificate_store<P>(&mut self, payload: P) -> Result<&mut Self, Error>
    where
        P: AsRef<[u8]>,
    {
        let payload_ref = payload.as_ref();
        codec::check_payload_size(
            payload_ref.len(),
            "certificate-store",
            self.source.limits().max_certificate_store_bytes,
        )?;
        let mut signature = self.signature.clone();
        let certificate_store = payload_ref.to_vec();
        self.stage(&mut signature, certificate_store)
    }

    /// Alias for [`Self::set_certificate_store`] using replacement
    /// terminology.
    ///
    /// # Errors
    ///
    /// Returns the same errors as [`Self::set_certificate_store`].
    pub fn replace_certificate_store<P>(&mut self, payload: P) -> Result<&mut Self, Error>
    where
        P: AsRef<[u8]>,
    {
        self.set_certificate_store(payload)
    }

    /// Applies a custom payload edit against cloned drafts.
    ///
    /// The candidate is published only after wire-size, offset, alignment,
    /// reserved-field, and resource validation succeeds. A closure error or
    /// malformed candidate leaves the transaction unchanged.
    ///
    /// # Errors
    ///
    /// Returns the closure's error or [`Error`] when the edited payloads are
    /// outside the configured bounds or not representable on the wire.
    pub fn update<F>(&mut self, edit: F) -> Result<&mut Self, Error>
    where
        F: FnOnce(&mut Vec<u8>, &mut Vec<u8>) -> Result<(), Error>,
    {
        let mut signature = self.signature.clone();
        let mut certificate_store = self.certificate_store.clone();
        edit(&mut signature, &mut certificate_store)?;
        self.stage(&mut signature, certificate_store)
    }

    /// Projects the current draft as a validated immutable snapshot.
    ///
    /// # Errors
    ///
    /// Returns [`Error`] when the staged payloads cannot be serialized under
    /// the source limits.
    pub fn snapshot(&self) -> Result<Snapshot, Error> {
        self.materialize()
    }

    /// Discards all staged changes and recovers the exact source snapshot.
    #[must_use]
    pub fn rollback(self) -> Snapshot {
        self.source
    }

    /// Publishes the draft as a reversible source-checked edit.
    ///
    /// # Errors
    ///
    /// Returns [`Error`] when the staged payloads cannot be serialized under
    /// the source limits.
    pub fn commit(self) -> Result<Commit, Error> {
        let snapshot = self.materialize()?;
        let patch = Patch::new(self.source, snapshot.clone());
        Ok(Commit { snapshot, patch })
    }

    /// Alias for [`Self::commit`] using writer terminology.
    ///
    /// # Errors
    ///
    /// Returns the same errors as [`Self::commit`].
    pub fn finish(self) -> Result<Commit, Error> {
        self.commit()
    }

    fn stage(
        &mut self,
        signature: &mut Vec<u8>,
        certificate_store: Vec<u8>,
    ) -> Result<&mut Self, Error> {
        let candidate = codec::rewrite(
            self.source.bytes(),
            self.source.kind(),
            self.source.layout(),
            signature,
            &certificate_store,
            self.source.limits(),
        )?;
        drop(candidate);
        self.signature = std::mem::take(signature);
        self.certificate_store = certificate_store;
        Ok(self)
    }

    fn materialize(&self) -> Result<Snapshot, Error> {
        if !self.is_changed() {
            return Ok(self.source.clone());
        }
        let bytes = codec::rewrite(
            self.source.bytes(),
            self.source.kind(),
            self.source.layout(),
            &self.signature,
            &self.certificate_store,
            self.source.limits(),
        )?;
        Snapshot::parse_shared(
            Arc::<[u8]>::from(bytes.into_boxed_slice()),
            self.source.kind(),
            self.source.limits(),
        )
    }
}

/// A successful immutable VBA signature publication and its reversible patch.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
}

impl Commit {
    /// Whether publication changed any source byte.
    #[must_use]
    pub fn changed(&self) -> bool {
        !self.patch.is_noop()
    }

    /// Borrows the published immutable snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Borrows the reversible source-checked patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Consumes the publication into its target snapshot.
    #[must_use]
    pub fn into_snapshot(self) -> Snapshot {
        self.snapshot
    }

    /// Consumes the publication into its reversible patch.
    #[must_use]
    pub fn into_patch(self) -> Patch {
        self.patch
    }

    /// Splits the publication into its target snapshot and patch.
    #[must_use]
    pub fn into_parts(self) -> (Snapshot, Patch) {
        (self.snapshot, self.patch)
    }
}

/// Runs one isolated payload edit and publishes it atomically.
///
/// # Errors
///
/// Returns the closure's error or [`Error`] when the edited payloads cannot
/// be represented under the source limits.
pub fn update<F>(snapshot: &Snapshot, edit: F) -> Result<Commit, Error>
where
    F: FnOnce(&mut Transaction) -> Result<(), Error>,
{
    let mut transaction = snapshot.edit();
    edit(&mut transaction)?;
    transaction.commit()
}
