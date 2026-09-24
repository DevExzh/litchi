//! Source-bound snapshots, edits, commits, and reversible patches.

use std::sync::Arc;

use litchi_opc::OpcPackage;

use super::codec::{self, Kind, Owner, Value};
use super::{PresenceInfo, ThreadingInfo};
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
}

impl AuthorPresenceSnapshot {
    /// Load one author-owned presence extension from an OPC package.
    pub fn load(package: &OpcPackage, author_id: u32) -> Result<Option<Self>> {
        super::package::load_presence_snapshot(package, author_id)
    }

    /// Parse one complete comment-author part while retaining its exact bytes.
    pub fn from_xml(source: impl AsRef<[u8]>, author_id: u32) -> Result<Self> {
        Self::from_xml_owned(source.as_ref().to_vec(), author_id)
    }

    /// Parse owned source bytes without a second source copy.
    pub fn from_xml_owned(source: Vec<u8>, author_id: u32) -> Result<Self> {
        let source_xml = Arc::new(source);
        let located = codec::locate(
            source_xml.as_slice(),
            Kind::Presence,
            Owner::Author(author_id),
        )?;
        let revision = fingerprint(source_xml.as_slice());
        Ok(Self {
            source_part_name: String::new(),
            source_xml,
            author_id,
            located,
            revision,
        })
    }

    pub(crate) fn from_part(
        source_part_name: String,
        source_xml: Arc<Vec<u8>>,
        author_id: u32,
    ) -> Result<Self> {
        let located = codec::locate(
            source_xml.as_slice(),
            Kind::Presence,
            Owner::Author(author_id),
        )?;
        let revision = fingerprint(source_xml.as_slice());
        Ok(Self {
            source_part_name,
            source_xml,
            author_id,
            located,
            revision,
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
    pub fn edit(&self) -> AuthorPresenceTransaction {
        AuthorPresenceTransaction {
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
    }

    pub(crate) fn source_arc(&self) -> &Arc<Vec<u8>> {
        &self.source_xml
    }
}

/// Detached edit of one author presence extension.
#[derive(Debug, Clone)]
pub struct AuthorPresenceTransaction {
    original: AuthorPresenceSnapshot,
    working: Option<PresenceInfo>,
}

impl AuthorPresenceTransaction {
    #[inline]
    #[must_use]
    pub fn value(&self) -> Option<&PresenceInfo> {
        self.working.as_ref()
    }

    #[inline]
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.original.value() != self.value()
    }

    /// Set or remove the optional presence extension.
    pub fn set(&mut self, value: Option<PresenceInfo>) -> Result<bool> {
        if let Some(value) = &value {
            codec::validate_value(Kind::Presence, &Value::Presence(value.clone()))?;
        }
        let changed = self.working != value;
        self.working = value;
        Ok(changed)
    }

    /// Set a concrete inert presence value.
    pub fn set_presence(&mut self, value: PresenceInfo) -> Result<bool> {
        self.set(Some(value))
    }

    /// Remove the recognized presence extension.
    pub fn remove(&mut self) -> Result<bool> {
        self.set(None)
    }

    /// Validate and consume into a source-checked reversible patch.
    pub fn commit(self) -> Result<AuthorPresenceCommit> {
        if !self.is_changed() {
            let patch = AuthorPresencePatch::new(self.original.clone(), self.original.clone());
            return Ok(AuthorPresenceCommit {
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
        )?;
        let source_xml = Arc::new(updated);
        let revision = fingerprint(source_xml.as_slice());
        let located = codec::locate(
            source_xml.as_slice(),
            Kind::Presence,
            Owner::Author(self.original.author_id),
        )?;
        let snapshot = AuthorPresenceSnapshot {
            source_part_name: self.original.source_part_name.clone(),
            source_xml,
            author_id: self.original.author_id,
            located,
            revision,
        };
        if snapshot.value() != self.value() {
            return Err(invalid("presenceInfo source did not round-trip"));
        }
        let patch = AuthorPresencePatch::new(self.original, snapshot.clone());
        Ok(AuthorPresenceCommit {
            snapshot,
            patch,
            changed: true,
        })
    }
}

/// A committed source-preserving author presence edit.
#[derive(Debug, Clone)]
pub struct AuthorPresenceCommit {
    snapshot: AuthorPresenceSnapshot,
    patch: AuthorPresencePatch,
    changed: bool,
}

impl AuthorPresenceCommit {
    #[inline]
    #[must_use]
    pub const fn is_changed(&self) -> bool {
        self.changed
    }

    #[inline]
    #[must_use]
    pub fn snapshot(&self) -> &AuthorPresenceSnapshot {
        &self.snapshot
    }

    #[inline]
    #[must_use]
    pub fn patch(&self) -> &AuthorPresencePatch {
        &self.patch
    }

    #[must_use]
    pub fn into_patch(self) -> AuthorPresencePatch {
        self.patch
    }
}

/// A reversible source-checked author presence replacement.
#[derive(Debug, Clone)]
pub struct AuthorPresencePatch {
    before: AuthorPresenceSnapshot,
    after: AuthorPresenceSnapshot,
}

impl AuthorPresencePatch {
    fn new(before: AuthorPresenceSnapshot, after: AuthorPresenceSnapshot) -> Self {
        Self { before, after }
    }

    #[inline]
    #[must_use]
    pub fn before(&self) -> &AuthorPresenceSnapshot {
        &self.before
    }

    #[inline]
    #[must_use]
    pub fn after(&self) -> &AuthorPresenceSnapshot {
        &self.after
    }

    #[inline]
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.before.same_source(&self.after)
    }

    #[inline]
    #[must_use]
    pub const fn expected_revision(&self) -> Revision {
        self.before.revision
    }

    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

    /// Apply to an exact source buffer.
    pub fn apply(&self, target: &mut Vec<u8>) -> Result<AuthorPresenceSnapshot> {
        if target.as_slice() != self.before.source_xml.as_slice() {
            return Err(invalid("author presence source is stale"));
        }
        if self.is_empty() {
            return Ok(self.before.clone());
        }
        let located = codec::locate(
            self.after.source_xml.as_slice(),
            Kind::Presence,
            Owner::Author(self.after.author_id),
        )?;
        let result = AuthorPresenceSnapshot {
            source_part_name: self.after.source_part_name.clone(),
            source_xml: Arc::clone(&self.after.source_xml),
            author_id: self.after.author_id,
            located,
            revision: fingerprint(self.after.source_xml.as_slice()),
        };
        if !result.same_source(&self.after) {
            return Err(invalid(
                "published author presence source differs from patch",
            ));
        }
        *target = self.after.source_xml.as_ref().clone();
        Ok(result)
    }

    /// Apply atomically to the exact OPC author part.
    pub fn apply_package(&self, target: &mut OpcPackage) -> Result<AuthorPresenceSnapshot> {
        super::package::apply_presence_patch(target, self)
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
    pub fn from_xml(source: impl AsRef<[u8]>, author_id: u32, index: u32) -> Result<Self> {
        Self::from_xml_owned(source.as_ref().to_vec(), String::new(), author_id, index)
    }

    /// Parse owned source bytes with an optional source slide name.
    pub fn from_xml_owned(
        source: Vec<u8>,
        slide_part_name: impl Into<String>,
        author_id: u32,
        index: u32,
    ) -> Result<Self> {
        let source_xml = Arc::new(source);
        let located = codec::locate(
            source_xml.as_slice(),
            Kind::Threading,
            Owner::Comment { author_id, index },
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
        })
    }

    pub(crate) fn from_part(
        source_part_name: String,
        source_xml: Arc<Vec<u8>>,
        slide_part_name: String,
        author_id: u32,
        index: u32,
    ) -> Result<Self> {
        let located = codec::locate(
            source_xml.as_slice(),
            Kind::Threading,
            Owner::Comment { author_id, index },
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
    pub fn edit(&self) -> CommentThreadingTransaction {
        CommentThreadingTransaction {
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
    }

    pub(crate) fn source_arc(&self) -> &Arc<Vec<u8>> {
        &self.source_xml
    }
}

/// Detached edit of one comment threading extension.
#[derive(Debug, Clone)]
pub struct CommentThreadingTransaction {
    original: CommentThreadingSnapshot,
    working: Option<ThreadingInfo>,
}

impl CommentThreadingTransaction {
    #[inline]
    #[must_use]
    pub fn value(&self) -> Option<&ThreadingInfo> {
        self.working.as_ref()
    }

    #[inline]
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.original.value() != self.value()
    }

    /// Set or remove the optional threading extension.
    pub fn set(&mut self, value: Option<ThreadingInfo>) -> Result<bool> {
        if let Some(value) = &value {
            codec::validate_value(Kind::Threading, &Value::Threading(value.clone()))?;
        }
        let changed = self.working != value;
        self.working = value;
        Ok(changed)
    }

    /// Set concrete inert threading metadata.
    pub fn set_threading(&mut self, value: ThreadingInfo) -> Result<bool> {
        self.set(Some(value))
    }

    /// Remove the recognized threading extension.
    pub fn remove(&mut self) -> Result<bool> {
        self.set(None)
    }

    /// Validate and consume into a source-checked reversible patch.
    pub fn commit(self) -> Result<CommentThreadingCommit> {
        if !self.is_changed() {
            let patch = CommentThreadingPatch::new(self.original.clone(), self.original.clone());
            return Ok(CommentThreadingCommit {
                snapshot: self.original,
                patch,
                changed: false,
            });
        }
        let value = self.working.clone().map(Value::Threading);
        let updated = codec::rewrite(
            self.original.source_xml.as_slice(),
            &self.original.located,
            value,
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
        )?;
        let snapshot = CommentThreadingSnapshot {
            source_part_name: self.original.source_part_name.clone(),
            source_xml,
            slide_part_name: self.original.slide_part_name.clone(),
            author_id: self.original.author_id,
            index: self.original.index,
            located,
            revision,
        };
        if snapshot.value() != self.value() {
            return Err(invalid("threadingInfo source did not round-trip"));
        }
        let patch = CommentThreadingPatch::new(self.original, snapshot.clone());
        Ok(CommentThreadingCommit {
            snapshot,
            patch,
            changed: true,
        })
    }
}

/// A committed source-preserving comment-threading edit.
#[derive(Debug, Clone)]
pub struct CommentThreadingCommit {
    snapshot: CommentThreadingSnapshot,
    patch: CommentThreadingPatch,
    changed: bool,
}

impl CommentThreadingCommit {
    #[inline]
    #[must_use]
    pub const fn is_changed(&self) -> bool {
        self.changed
    }

    #[inline]
    #[must_use]
    pub fn snapshot(&self) -> &CommentThreadingSnapshot {
        &self.snapshot
    }

    #[inline]
    #[must_use]
    pub fn patch(&self) -> &CommentThreadingPatch {
        &self.patch
    }

    #[must_use]
    pub fn into_patch(self) -> CommentThreadingPatch {
        self.patch
    }
}

/// A reversible source-checked comment-threading replacement.
#[derive(Debug, Clone)]
pub struct CommentThreadingPatch {
    before: CommentThreadingSnapshot,
    after: CommentThreadingSnapshot,
}

impl CommentThreadingPatch {
    fn new(before: CommentThreadingSnapshot, after: CommentThreadingSnapshot) -> Self {
        Self { before, after }
    }

    #[inline]
    #[must_use]
    pub fn before(&self) -> &CommentThreadingSnapshot {
        &self.before
    }

    #[inline]
    #[must_use]
    pub fn after(&self) -> &CommentThreadingSnapshot {
        &self.after
    }

    #[inline]
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.before.same_source(&self.after)
    }

    #[inline]
    #[must_use]
    pub const fn expected_revision(&self) -> Revision {
        self.before.revision
    }

    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

    /// Apply to an exact source buffer.
    pub fn apply(&self, target: &mut Vec<u8>) -> Result<CommentThreadingSnapshot> {
        if target.as_slice() != self.before.source_xml.as_slice() {
            return Err(invalid("comment threading source is stale"));
        }
        if self.is_empty() {
            return Ok(self.before.clone());
        }
        let located = codec::locate(
            self.after.source_xml.as_slice(),
            Kind::Threading,
            Owner::Comment {
                author_id: self.after.author_id,
                index: self.after.index,
            },
        )?;
        let result = CommentThreadingSnapshot {
            source_part_name: self.after.source_part_name.clone(),
            source_xml: Arc::clone(&self.after.source_xml),
            slide_part_name: self.after.slide_part_name.clone(),
            author_id: self.after.author_id,
            index: self.after.index,
            located,
            revision: fingerprint(self.after.source_xml.as_slice()),
        };
        if !result.same_source(&self.after) {
            return Err(invalid(
                "published comment threading source differs from patch",
            ));
        }
        *target = self.after.source_xml.as_ref().clone();
        Ok(result)
    }

    /// Apply atomically to the exact OPC comment part.
    pub fn apply_package(&self, target: &mut OpcPackage) -> Result<CommentThreadingSnapshot> {
        super::package::apply_threading_patch(target, self)
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
