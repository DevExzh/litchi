//! Source-bound snapshots, edits, commits, and reversible patches.

use std::sync::Arc;

use litchi_opc::OpcPackage;

use super::codec::{self, Kind, Owner, Value};
use super::{PresenceInfo, ThreadingInfo};
use crate::comments::Limits;
use crate::{Error, Result};

/// Stable fingerprint used for optimistic source checks.
pub type Revision = u64;

/// One exact author-part source carrying optional `p15:presenceInfo` data.
#[derive(Debug, Clone)]
pub struct AuthorPresenceSnapshot {
    pub(crate) source_part_name: String,
    pub(crate) source_xml: Arc<Vec<u8>>,
    pub(crate) author_id: u32,
    pub(crate) located: codec::Located,
    pub(crate) revision: Revision,
    pub(crate) limits: Limits,
}

impl AuthorPresenceSnapshot {
    /// Load one author-owned presence extension from an OPC package.
    pub fn load(package: &OpcPackage, author_id: u32) -> Result<Option<Self>> {
        super::package::load_presence_snapshot(package, author_id)
    }

    /// Parse one complete comment-author part while retaining its exact bytes.
    pub fn from_xml(source: impl AsRef<[u8]>, author_id: u32, limits: Limits) -> Result<Self> {
        let source = source.as_ref();
        check_source_size(source, limits)?;
        let mut owned = Vec::new();
        owned
            .try_reserve_exact(source.len())
            .map_err(|error| Error::Allocation {
                resource: "legacy comment collaboration source XML",
                source: error,
            })?;
        owned.extend_from_slice(source);
        Self::from_xml_owned(owned, author_id, limits)
    }

    /// Parse owned source bytes without a second source copy.
    pub fn from_xml_owned(source: Vec<u8>, author_id: u32, limits: Limits) -> Result<Self> {
        check_source_size(&source, limits)?;
        let source_xml = Arc::new(source);
        let located = codec::locate(
            source_xml.as_slice(),
            Kind::Presence,
            Owner::Author(author_id),
            limits,
        )?;
        let revision = fingerprint(source_xml.as_slice());
        Ok(Self {
            source_part_name: String::new(),
            source_xml,
            author_id,
            located,
            revision,
            limits,
        })
    }

    pub(crate) fn from_part(
        source_part_name: String,
        source_xml: Arc<Vec<u8>>,
        author_id: u32,
        limits: Limits,
    ) -> Result<Self> {
        let located = codec::locate(
            source_xml.as_slice(),
            Kind::Presence,
            Owner::Author(author_id),
            limits,
        )?;
        let revision = fingerprint(source_xml.as_slice());
        Ok(Self {
            source_part_name,
            source_xml,
            author_id,
            located,
            revision,
            limits,
        })
    }

    #[inline]
    #[must_use]
    pub const fn author_id(&self) -> u32 {
        self.author_id
    }

    #[inline]
    #[must_use]
    pub fn value(&self) -> Option<&PresenceInfo> {
        match self.located.value.as_ref() {
            Some(Value::Presence(value)) => Some(value),
            _ => None,
        }
    }

    #[inline]
    #[must_use]
    pub fn source_part_name(&self) -> &str {
        &self.source_part_name
    }

    #[inline]
    #[must_use]
    pub fn part_name(&self) -> &str {
        self.source_part_name()
    }

    #[inline]
    #[must_use]
    pub fn source_xml(&self) -> &[u8] {
        self.source_xml.as_slice()
    }

    #[inline]
    #[must_use]
    pub const fn revision(&self) -> Revision {
        self.revision
    }

    #[must_use]
    pub(crate) fn edit(&self) -> PresenceDraft {
        PresenceDraft {
            original: self.clone(),
            working: self.value().cloned(),
        }
    }

    pub(crate) fn same_source(&self, other: &Self) -> bool {
        self.source_part_name == other.source_part_name
            && self.source_xml.as_slice() == other.source_xml.as_slice()
            && self.author_id == other.author_id
            && self.value() == other.value()
            && self.revision == other.revision
            && self.limits == other.limits
    }

    pub(crate) fn source_arc(&self) -> &Arc<Vec<u8>> {
        &self.source_xml
    }
}

/// Detached edit of one author presence extension.
#[derive(Debug, Clone)]
pub(crate) struct PresenceDraft {
    original: AuthorPresenceSnapshot,
    working: Option<PresenceInfo>,
}

impl PresenceDraft {
    #[inline]
    #[must_use]
    pub(crate) fn value(&self) -> Option<&PresenceInfo> {
        self.working.as_ref()
    }

    fn is_changed(&self) -> bool {
        self.original.value() != self.value()
    }

    /// The outcome of this edit or commit.
    #[must_use]
    #[cfg(test)]
    pub(crate) fn diagnostics(&self) -> crate::diagnostics::Diagnostics {
        crate::diagnostics::Diagnostics::new(self.is_changed())
    }

    /// Set or remove the optional presence extension.
    pub(crate) fn set(&mut self, value: Option<PresenceInfo>) -> Result<bool> {
        if let Some(value) = &value {
            codec::validate_value_with_limit(
                Kind::Presence,
                &Value::Presence(value.clone()),
                self.original.limits.max_string_bytes,
            )?;
        }
        let changed = self.working != value;
        self.working = value;
        Ok(changed)
    }

    /// Set a concrete inert presence value.
    pub(crate) fn set_presence(&mut self, value: PresenceInfo) -> Result<bool> {
        self.set(Some(value))
    }

    /// Remove the recognized presence extension.
    pub(crate) fn remove(&mut self) -> Result<bool> {
        self.set(None)
    }

    /// Validate and consume into a source-checked reversible patch.
    pub(crate) fn commit(self) -> Result<PresencePrepared> {
        if !self.is_changed() {
            let patch = PresenceChange::new(self.original.clone(), self.original.clone());
            return Ok(PresencePrepared {
                snapshot: self.original,
                patch,
                changed: false,
            });
        }
        let value = self.working.clone().map(Value::Presence);
        let updated = codec::rewrite(
            self.original.source_xml.as_slice(),
            &self.original.located,
            value,
            self.original.limits,
        )?;
        let source_xml = Arc::new(updated);
        let revision = fingerprint(source_xml.as_slice());
        let located = codec::locate(
            source_xml.as_slice(),
            Kind::Presence,
            Owner::Author(self.original.author_id),
            self.original.limits,
        )?;
        let snapshot = AuthorPresenceSnapshot {
            source_part_name: self.original.source_part_name.clone(),
            source_xml,
            author_id: self.original.author_id,
            located,
            revision,
            limits: self.original.limits,
        };
        if snapshot.value() != self.value() {
            return Err(invalid("presenceInfo source did not round-trip"));
        }
        let patch = PresenceChange::new(self.original, snapshot.clone());
        Ok(PresencePrepared {
            snapshot,
            patch,
            changed: true,
        })
    }
}

/// A committed source-preserving author presence edit.
#[derive(Debug, Clone)]
pub(crate) struct PresencePrepared {
    #[cfg_attr(
        not(test),
        expect(dead_code, reason = "read by the module's source-level tests")
    )]
    snapshot: AuthorPresenceSnapshot,
    patch: PresenceChange,
    #[cfg_attr(
        not(test),
        expect(dead_code, reason = "read by the module's source-level tests")
    )]
    changed: bool,
}

impl PresencePrepared {
    #[cfg(test)]
    fn is_changed(&self) -> bool {
        self.changed
    }

    /// The outcome of this edit or commit.
    #[must_use]
    #[cfg(test)]
    pub(crate) fn diagnostics(&self) -> crate::diagnostics::Diagnostics {
        crate::diagnostics::Diagnostics::new(self.is_changed())
    }

    #[inline]
    #[must_use]
    #[cfg(test)]
    pub(crate) fn snapshot(&self) -> &AuthorPresenceSnapshot {
        &self.snapshot
    }

    #[inline]
    #[must_use]
    pub(crate) fn patch(&self) -> &PresenceChange {
        &self.patch
    }
}

/// A reversible source-checked author presence replacement.
#[derive(Debug, Clone)]
pub(crate) struct PresenceChange {
    before: AuthorPresenceSnapshot,
    after: AuthorPresenceSnapshot,
}

impl PresenceChange {
    fn new(before: AuthorPresenceSnapshot, after: AuthorPresenceSnapshot) -> Self {
        Self { before, after }
    }

    #[inline]
    #[must_use]
    pub(crate) fn before(&self) -> &AuthorPresenceSnapshot {
        &self.before
    }

    #[inline]
    #[must_use]
    pub(crate) fn after(&self) -> &AuthorPresenceSnapshot {
        &self.after
    }

    #[inline]
    #[must_use]
    pub(crate) fn is_noop(&self) -> bool {
        self.before.same_source(&self.after)
    }

    #[must_use]
    #[cfg(test)]
    pub(crate) fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

    /// Apply to an exact source buffer.
    #[cfg(test)]
    pub(crate) fn apply(&self, target: &mut Vec<u8>) -> Result<AuthorPresenceSnapshot> {
        if target.as_slice() != self.before.source_xml.as_slice() {
            return Err(invalid("author presence source is stale"));
        }
        if self.is_noop() {
            return Ok(self.before.clone());
        }
        let located = codec::locate(
            self.after.source_xml.as_slice(),
            Kind::Presence,
            Owner::Author(self.after.author_id),
            self.after.limits,
        )?;
        let result = AuthorPresenceSnapshot {
            source_part_name: self.after.source_part_name.clone(),
            source_xml: Arc::clone(&self.after.source_xml),
            author_id: self.after.author_id,
            located,
            revision: fingerprint(self.after.source_xml.as_slice()),
            limits: self.after.limits,
        };
        if !result.same_source(&self.after) {
            return Err(invalid(
                "published author presence source differs from patch",
            ));
        }
        *target = self.after.source_xml.as_ref().clone();
        Ok(result)
    }
}

/// One exact comment-part source carrying optional `p15:threadingInfo` data.
#[derive(Debug, Clone)]
pub struct CommentThreadingSnapshot {
    pub(crate) source_part_name: String,
    pub(crate) source_xml: Arc<Vec<u8>>,
    pub(crate) slide_part_name: String,
    pub(crate) author_id: u32,
    pub(crate) index: u32,
    pub(crate) located: codec::Located,
    pub(crate) revision: Revision,
    pub(crate) limits: Limits,
}

impl CommentThreadingSnapshot {
    /// Load one comment-owned threading extension from an OPC package.
    pub fn load(
        package: &OpcPackage,
        slide_part_name: &str,
        author_id: u32,
        index: u32,
    ) -> Result<Option<Self>> {
        super::package::load_threading_snapshot(package, slide_part_name, author_id, index)
    }

    /// Parse one complete comment part while retaining its exact bytes.
    pub fn from_xml(
        source: impl AsRef<[u8]>,
        author_id: u32,
        index: u32,
        limits: Limits,
    ) -> Result<Self> {
        let source = source.as_ref();
        check_source_size(source, limits)?;
        let mut owned = Vec::new();
        owned
            .try_reserve_exact(source.len())
            .map_err(|error| Error::Allocation {
                resource: "legacy comment collaboration source XML",
                source: error,
            })?;
        owned.extend_from_slice(source);
        Self::from_xml_owned(owned, String::new(), author_id, index, limits)
    }

    /// Parse owned source bytes with an optional source slide name.
    pub fn from_xml_owned(
        source: Vec<u8>,
        slide_part_name: impl Into<String>,
        author_id: u32,
        index: u32,
        limits: Limits,
    ) -> Result<Self> {
        check_source_size(&source, limits)?;
        let source_xml = Arc::new(source);
        let located = codec::locate(
            source_xml.as_slice(),
            Kind::Threading,
            Owner::Comment { author_id, index },
            limits,
        )?;
        let revision = fingerprint(source_xml.as_slice());
        Ok(Self {
            source_part_name: String::new(),
            source_xml,
            slide_part_name: slide_part_name.into(),
            author_id,
            index,
            located,
            revision,
            limits,
        })
    }

    pub(crate) fn from_part(
        source_part_name: String,
        source_xml: Arc<Vec<u8>>,
        slide_part_name: String,
        author_id: u32,
        index: u32,
        limits: Limits,
    ) -> Result<Self> {
        let located = codec::locate(
            source_xml.as_slice(),
            Kind::Threading,
            Owner::Comment { author_id, index },
            limits,
        )?;
        let revision = fingerprint(source_xml.as_slice());
        Ok(Self {
            source_part_name,
            source_xml,
            slide_part_name,
            author_id,
            index,
            located,
            revision,
            limits,
        })
    }

    #[inline]
    #[must_use]
    pub fn slide_part_name(&self) -> &str {
        &self.slide_part_name
    }

    #[inline]
    #[must_use]
    pub const fn author_id(&self) -> u32 {
        self.author_id
    }

    #[inline]
    #[must_use]
    pub const fn index(&self) -> u32 {
        self.index
    }

    #[inline]
    #[must_use]
    pub fn value(&self) -> Option<&ThreadingInfo> {
        match self.located.value.as_ref() {
            Some(Value::Threading(value)) => Some(value),
            _ => None,
        }
    }

    #[inline]
    #[must_use]
    pub fn source_part_name(&self) -> &str {
        &self.source_part_name
    }

    #[inline]
    #[must_use]
    pub fn part_name(&self) -> &str {
        self.source_part_name()
    }

    #[inline]
    #[must_use]
    pub fn source_xml(&self) -> &[u8] {
        self.source_xml.as_slice()
    }

    #[inline]
    #[must_use]
    pub const fn revision(&self) -> Revision {
        self.revision
    }

    #[must_use]
    pub(crate) fn edit(&self) -> ThreadingDraft {
        ThreadingDraft {
            original: self.clone(),
            working: self.value().cloned(),
        }
    }

    pub(crate) fn same_source(&self, other: &Self) -> bool {
        self.source_part_name == other.source_part_name
            && self.source_xml.as_slice() == other.source_xml.as_slice()
            && self.slide_part_name == other.slide_part_name
            && self.author_id == other.author_id
            && self.index == other.index
            && self.value() == other.value()
            && self.revision == other.revision
            && self.limits == other.limits
    }

    pub(crate) fn source_arc(&self) -> &Arc<Vec<u8>> {
        &self.source_xml
    }
}

fn check_source_size(source: &[u8], limits: Limits) -> Result<()> {
    if source.len() > limits.max_input_bytes {
        return Err(Error::Limit {
            kind: crate::error::LimitKind::Codec(crate::error::CodecLimitKind::XmlBytes),
            observed: source.len(),
            maximum: limits.max_input_bytes,
        });
    }
    Ok(())
}

/// Detached edit of one comment threading extension.
#[derive(Debug, Clone)]
pub(crate) struct ThreadingDraft {
    original: CommentThreadingSnapshot,
    working: Option<ThreadingInfo>,
}

impl ThreadingDraft {
    #[inline]
    #[must_use]
    pub(crate) fn value(&self) -> Option<&ThreadingInfo> {
        self.working.as_ref()
    }

    fn is_changed(&self) -> bool {
        self.original.value() != self.value()
    }

    /// Set or remove the optional threading extension.
    pub(crate) fn set(&mut self, value: Option<ThreadingInfo>) -> Result<bool> {
        if let Some(value) = &value {
            codec::validate_value_with_limit(
                Kind::Threading,
                &Value::Threading(value.clone()),
                self.original.limits.max_string_bytes,
            )?;
        }
        let changed = self.working != value;
        self.working = value;
        Ok(changed)
    }

    /// Set concrete inert threading metadata.
    pub(crate) fn set_threading(&mut self, value: ThreadingInfo) -> Result<bool> {
        self.set(Some(value))
    }

    /// Remove the recognized threading extension.
    pub(crate) fn remove(&mut self) -> Result<bool> {
        self.set(None)
    }

    /// Validate and consume into a source-checked reversible patch.
    pub(crate) fn commit(self) -> Result<ThreadingPrepared> {
        if !self.is_changed() {
            let patch = ThreadingChange::new(self.original.clone(), self.original.clone());
            return Ok(ThreadingPrepared {
                snapshot: self.original,
                patch,
            });
        }
        let value = self.working.clone().map(Value::Threading);
        let updated = codec::rewrite(
            self.original.source_xml.as_slice(),
            &self.original.located,
            value,
            self.original.limits,
        )?;
        let source_xml = Arc::new(updated);
        let revision = fingerprint(source_xml.as_slice());
        let located = codec::locate(
            source_xml.as_slice(),
            Kind::Threading,
            Owner::Comment {
                author_id: self.original.author_id,
                index: self.original.index,
            },
            self.original.limits,
        )?;
        let snapshot = CommentThreadingSnapshot {
            source_part_name: self.original.source_part_name.clone(),
            source_xml,
            slide_part_name: self.original.slide_part_name.clone(),
            author_id: self.original.author_id,
            index: self.original.index,
            located,
            revision,
            limits: self.original.limits,
        };
        if snapshot.value() != self.value() {
            return Err(invalid("threadingInfo source did not round-trip"));
        }
        let patch = ThreadingChange::new(self.original, snapshot.clone());
        Ok(ThreadingPrepared { snapshot, patch })
    }
}

/// A committed source-preserving comment-threading edit.
#[derive(Debug, Clone)]
pub(crate) struct ThreadingPrepared {
    #[cfg_attr(
        not(test),
        expect(dead_code, reason = "read by the module's source-level tests")
    )]
    snapshot: CommentThreadingSnapshot,
    patch: ThreadingChange,
}

impl ThreadingPrepared {
    #[inline]
    #[must_use]
    #[cfg(test)]
    pub(crate) fn snapshot(&self) -> &CommentThreadingSnapshot {
        &self.snapshot
    }

    #[inline]
    #[must_use]
    pub(crate) fn patch(&self) -> &ThreadingChange {
        &self.patch
    }
}

/// A reversible source-checked comment-threading replacement.
#[derive(Debug, Clone)]
pub(crate) struct ThreadingChange {
    before: CommentThreadingSnapshot,
    after: CommentThreadingSnapshot,
}

impl ThreadingChange {
    fn new(before: CommentThreadingSnapshot, after: CommentThreadingSnapshot) -> Self {
        Self { before, after }
    }

    #[inline]
    #[must_use]
    pub(crate) fn before(&self) -> &CommentThreadingSnapshot {
        &self.before
    }

    #[inline]
    #[must_use]
    pub(crate) fn after(&self) -> &CommentThreadingSnapshot {
        &self.after
    }

    #[inline]
    #[must_use]
    pub(crate) fn is_noop(&self) -> bool {
        self.before.same_source(&self.after)
    }

    #[must_use]
    #[cfg(test)]
    pub(crate) fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

    /// Apply to an exact source buffer.
    #[cfg(test)]
    pub(crate) fn apply(&self, target: &mut Vec<u8>) -> Result<CommentThreadingSnapshot> {
        if target.as_slice() != self.before.source_xml.as_slice() {
            return Err(invalid("comment threading source is stale"));
        }
        if self.is_noop() {
            return Ok(self.before.clone());
        }
        let located = codec::locate(
            self.after.source_xml.as_slice(),
            Kind::Threading,
            Owner::Comment {
                author_id: self.after.author_id,
                index: self.after.index,
            },
            self.after.limits,
        )?;
        let result = CommentThreadingSnapshot {
            source_part_name: self.after.source_part_name.clone(),
            source_xml: Arc::clone(&self.after.source_xml),
            slide_part_name: self.after.slide_part_name.clone(),
            author_id: self.after.author_id,
            index: self.after.index,
            located,
            revision: fingerprint(self.after.source_xml.as_slice()),
            limits: self.after.limits,
        };
        if !result.same_source(&self.after) {
            return Err(invalid(
                "published comment threading source differs from patch",
            ));
        }
        *target = self.after.source_xml.as_ref().clone();
        Ok(result)
    }
}

fn fingerprint(bytes: &[u8]) -> Revision {
    let mut hash = 0xcbf2_9ce4_8422_2325u64;
    for byte in bytes {
        hash ^= u64::from(*byte);
        hash = hash.wrapping_mul(0x100_0000_01b3);
    }
    hash
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}
