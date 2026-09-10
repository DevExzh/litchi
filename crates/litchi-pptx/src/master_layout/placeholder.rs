//! Source checked, lossless transactions for `p232:phTypeExt`.
//!
//! The PowerPoint 2023 placeholder type extension is a scalar value, but its
//! owner is an arbitrary shape tree that may contain markup this crate does
//! not understand.  This module therefore edits only the local name of the
//! existing `p232:cameo` or `p232:unknown` empty token.  It never serializes
//! or replaces a complete shape.

use std::borrow::Cow;
use std::ops::Range;
use std::sync::Arc;

use litchi_ooxml_common::xml::unqualified_attribute_value;
use litchi_opc::constants::content_type as ct;
use litchi_opc::{OpcPackage, PackURI, Part};
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, QName, ResolveResult};
use quick_xml::reader::NsReader;

use super::codec::{
    MAX_NAME_CHARS, SPTREE_DEPTH, insert_bytes, invalid, next_shape_id, placeholder_shape_xml,
    scan_element_span, validate_xml10_chars,
};
use super::model::PlaceholderSpec;
use crate::shape::{Key, PLACEHOLDER_TYPE_EXTENSION_URI, PlaceholderTypeExtension, Scene};
use crate::{Error, Result};

const P232: &[u8] = b"http://schemas.microsoft.com/office/powerpoint/2023/02/main";
const PML: &[u8] = b"http://schemas.openxmlformats.org/presentationml/2006/main";
const STRICT_PML: &[u8] = b"http://purl.oclc.org/ooxml/presentationml/main";
const DML: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/main";
const STRICT_DML: &[u8] = b"http://purl.oclc.org/ooxml/drawingml/main";
const MAX_OWNER_BYTES: usize = 8 * 1024 * 1024;
const MAX_GROWTH_BYTES: usize = 64 * 1024;

/// Signature disposition for a changed source checked placeholder patch.
///
/// A no-op never needs a signature disposition.  A changed patch defaults to
/// [`SignaturePolicy::Reject`], preserving the package's signature graph.  A
/// caller that has deliberately accepted invalidating signatures must choose
/// [`SignaturePolicy::Invalidate`] explicitly.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SignaturePolicy {
    /// Refuse a changed patch while signature disposition is required.
    Reject,
    /// Remove the package signature infrastructure as part of publication.
    Invalidate,
}

/// Finite bounds applied before staging a source checked patch.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    /// Maximum bytes in the slide, master, or layout owner XML.
    pub owner_bytes: usize,
    /// Maximum increase in owner bytes caused by one scalar replacement.
    pub growth_bytes: usize,
}

impl Limits {
    /// Conservative bounded limits for one slide, master, or layout part.
    pub const DEFAULT: Self = Self {
        owner_bytes: MAX_OWNER_BYTES,
        growth_bytes: MAX_GROWTH_BYTES,
    };

    /// Construct finite, nonzero limits.
    #[must_use]
    pub const fn new(owner_bytes: usize, growth_bytes: usize) -> Option<Self> {
        if owner_bytes == 0 || growth_bytes == 0 {
            None
        } else {
            Some(Self {
                owner_bytes,
                growth_bytes,
            })
        }
    }

    fn check_source(self, bytes: usize) -> Result<()> {
        if bytes > self.owner_bytes {
            return Err(Error::Limit {
                resource: "p232 placeholder owner XML bytes",
                limit: self.owner_bytes,
            });
        }
        Ok(())
    }

    fn check_growth(self, source: usize, target: usize) -> Result<()> {
        self.check_source(source)?;
        self.check_source(target)?;
        let growth = target.saturating_sub(source);
        if growth > self.growth_bytes {
            return Err(Error::Limit {
                resource: "p232 placeholder patch growth bytes",
                limit: self.growth_bytes,
            });
        }
        Ok(())
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self::DEFAULT
    }
}

/// An exact source snapshot of one existing typed placeholder extension.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Snapshot {
    owner: PackURI,
    source: Arc<Vec<u8>>,
    shape_index: usize,
    shape_name: Option<String>,
    placeholder_index: u32,
    value: PlaceholderTypeExtension,
    variant: VariantSpan,
    limits: Limits,
}

impl Snapshot {
    /// Owning slide, master, or layout part.
    #[must_use]
    pub fn owner(&self) -> &PackURI {
        &self.owner
    }

    /// Exact owner XML captured by this snapshot.
    #[must_use]
    pub fn source_xml(&self) -> &[u8] {
        self.source.as_slice()
    }

    /// Checked pre-order shape index used by this source snapshot.
    #[must_use]
    pub const fn shape_index(&self) -> usize {
        self.shape_index
    }

    /// Shape name when the source supplied one.
    #[must_use]
    pub fn shape_name(&self) -> Option<&str> {
        self.shape_name.as_deref()
    }

    /// Placeholder index (`idx`, defaulting to zero).
    #[must_use]
    pub const fn placeholder_index(&self) -> u32 {
        self.placeholder_index
    }

    /// Typed p232 scalar value.
    #[must_use]
    pub const fn value(&self) -> PlaceholderTypeExtension {
        self.value
    }

    /// Start a move-only edit tied to this exact owner source.
    #[must_use]
    pub fn edit(self) -> Edit {
        Edit {
            snapshot: self,
            target: None,
        }
    }

    fn same_source(&self, other: &Self) -> bool {
        self.owner == other.owner
            && self.source.as_slice() == other.source.as_slice()
            && self.shape_index == other.shape_index
            && self.shape_name == other.shape_name
            && self.placeholder_index == other.placeholder_index
            && self.value == other.value
    }
}

/// A move-only scalar p232 edit.
#[derive(Debug)]
pub struct Edit {
    snapshot: Snapshot,
    target: Option<PlaceholderTypeExtension>,
}

impl Edit {
    /// Value projected by this edit.
    #[must_use]
    pub fn value(&self) -> PlaceholderTypeExtension {
        self.target.unwrap_or(self.snapshot.value)
    }

    /// Set the p232 token while retaining the source's spelling and layout.
    pub fn set(&mut self, value: PlaceholderTypeExtension) {
        self.target = Some(value);
    }

    /// Whether this edit changes the typed value.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.value() != self.snapshot.value
    }

    /// Consume this edit into a source checked, reversible commit.
    ///
    /// The staged bytes are bounded before the owner buffer is allocated.
    /// Only the local token name changes; all shape and extension bytes remain
    /// otherwise identical.
    pub fn commit(self) -> Result<Commit> {
        let target = self.value();
        if target == self.snapshot.value {
            return Ok(Commit {
                snapshot: self.snapshot.clone(),
                patch: Patch {
                    before: self.snapshot.clone(),
                    after: self.snapshot,
                },
            });
        }
        let staged = replace_variant(
            self.snapshot.source.as_slice(),
            &self.snapshot.variant,
            target,
            self.snapshot.limits,
        )?;
        let after = snapshot_from_source(
            self.snapshot.owner.clone(),
            Arc::new(staged),
            self.snapshot.shape_index,
            self.snapshot.limits,
        )?;
        if after.value != target {
            return Err(Error::Invalid(
                "staged p232 placeholder type did not round-trip".into(),
            ));
        }
        Ok(Commit {
            snapshot: after.clone(),
            patch: Patch {
                before: self.snapshot,
                after,
            },
        })
    }
}

/// A successful detached edit and its reversible source patch.
#[derive(Clone, Debug)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
}

impl Commit {
    /// Snapshot expected after publication.
    #[must_use]
    pub fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Reversible source patch represented by this commit.
    #[must_use]
    pub fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Whether this commit changes owner bytes.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.patch.is_changed()
    }

    /// Alias for callers that prefer a no-op predicate.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        !self.is_changed()
    }

    /// Consume this commit into its reversible patch.
    #[must_use]
    pub fn into_patch(self) -> Patch {
        self.patch
    }
}

/// A source checked, reversible p232 scalar replacement.
#[derive(Clone, Debug)]
pub struct Patch {
    before: Snapshot,
    after: Snapshot,
}

impl Patch {
    /// Snapshot required before forward application.
    #[must_use]
    pub fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Snapshot expected after forward application.
    #[must_use]
    pub fn after(&self) -> &Snapshot {
        &self.after
    }

    /// Whether this patch is an exact byte-preserving no-op.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before.same_source(&self.after)
    }

    /// Alias used by transaction callers that prefer changed terminology.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        !self.is_noop()
    }

    /// Return the exact inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

    /// Apply with the default signature policy, which rejects a changed
    /// signed source.
    pub fn apply(&self, package: &mut OpcPackage) -> Result<Snapshot> {
        apply_patch(package, self)
    }

    /// Apply with an explicit signature disposition.
    pub fn apply_with_policy(
        &self,
        package: &mut OpcPackage,
        policy: SignaturePolicy,
    ) -> Result<Snapshot> {
        apply_patch_with_policy(package, self, policy)
    }
}

/// Load one existing p232 placeholder type extension from a slide, master, or layout.
pub fn load_snapshot<'a>(
    package: &OpcPackage,
    owner: &PackURI,
    key: impl Into<Key<'a>>,
) -> Result<Snapshot> {
    load_snapshot_with_limits(package, owner, key, Limits::DEFAULT)
}

/// Load one existing p232 placeholder type extension under finite limits.
pub fn load_snapshot_with_limits<'a>(
    package: &OpcPackage,
    owner: &PackURI,
    key: impl Into<Key<'a>>,
    limits: Limits,
) -> Result<Snapshot> {
    let part = package.get_part(owner)?;
    validate_owner(part)?;
    let source = part.blob_arc();
    snapshot_from_source_for_key(part.partname().clone(), source, key.into(), limits)
}

/// Apply a patch while rejecting any changed signature disposition.
pub fn apply_patch(package: &mut OpcPackage, patch: &Patch) -> Result<Snapshot> {
    apply_patch_with_policy(package, patch, SignaturePolicy::Reject)
}

/// Apply a patch after an exact source check and an explicit signature policy.
pub fn apply_patch_with_policy(
    package: &mut OpcPackage,
    patch: &Patch,
    policy: SignaturePolicy,
) -> Result<Snapshot> {
    let part = package.get_part(&patch.before.owner)?;
    validate_owner(part)?;
    patch.before.limits.check_source(part.blob().len())?;
    if part.blob() != patch.before.source.as_slice() {
        return Err(Error::StaleSource);
    }
    if patch.is_noop() {
        return Ok(patch.before.clone());
    }

    patch
        .before
        .limits
        .check_growth(patch.before.source.len(), patch.after.source.len())?;
    if matches!(policy, SignaturePolicy::Reject)
        && (package.is_signed() || package.requires_signature_edit_policy())
    {
        return Err(Error::Opc(
            litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy,
        ));
    }

    // All source, semantic, growth, and signature checks complete before the
    // package clone or its first mutation.
    let mut candidate = package.clone();
    let staged = Arc::clone(&patch.after.source);
    candidate
        .get_part_mut(&patch.after.owner)?
        .set_blob_shared(staged);
    let published = load_snapshot_with_limits(
        &candidate,
        &patch.after.owner,
        Key::Index(patch.after.shape_index),
        patch.after.limits,
    )?;
    if !published.same_source(&patch.after) {
        return Err(Error::Invalid(
            "published p232 placeholder patch differs from the commit".into(),
        ));
    }
    if matches!(policy, SignaturePolicy::Invalidate) {
        candidate.unsign();
    }
    *package = candidate;
    Ok(published)
}

/// Apply a committed edit with the default signature policy.
pub fn apply_commit(package: &mut OpcPackage, commit: Commit) -> Result<Snapshot> {
    apply_patch(package, &commit.patch)
}

fn snapshot_from_source_for_key<'a>(
    owner: PackURI,
    source: Arc<Vec<u8>>,
    key: Key<'a>,
    limits: Limits,
) -> Result<Snapshot> {
    limits.check_source(source.len())?;
    let scene = Scene::read(source.as_slice())?;
    if scene.is_rewritten() {
        return Err(Error::UnsafeEdit {
            operation: "load_p232_placeholder_patch",
            reason: "markup compatibility rewrote the owner; no source-preserving patch policy is available",
        });
    }
    let shape = scene.shape(key)?;
    let placeholder = shape
        .placeholder()
        .ok_or_else(|| Error::Invalid("selected shape is not a placeholder".into()))?;
    let value = placeholder.type_extension().ok_or_else(|| {
        Error::Invalid("selected placeholder has no p232:phTypeExt scalar".into())
    })?;
    let span = shape.span()?.range(scene.xml().len())?;
    let (variant, variant_value) = locate_typed_variant(scene.xml(), span)?;
    if variant_value != value {
        return Err(Error::Invalid(
            "typed p232 placeholder metadata disagrees with source XML".into(),
        ));
    }
    let shape_index = shape.common().index();
    let shape_name = shape.name().map(str::to_owned);
    let placeholder_index = placeholder.index();
    Ok(Snapshot {
        owner,
        source,
        shape_index,
        shape_name,
        placeholder_index,
        value,
        variant,
        limits,
    })
}

fn snapshot_from_source(
    owner: PackURI,
    source: Arc<Vec<u8>>,
    shape_index: usize,
    limits: Limits,
) -> Result<Snapshot> {
    let scene = Scene::read(source.as_slice())?;
    if scene.is_rewritten() {
        return Err(Error::UnsafeEdit {
            operation: "stage_p232_placeholder_patch",
            reason: "markup compatibility rewrote the owner; no source-preserving patch policy is available",
        });
    }
    let shape = scene.at(shape_index)?;
    let placeholder = shape
        .placeholder()
        .ok_or_else(|| Error::Invalid("staged shape is no longer a placeholder".into()))?;
    let value = placeholder
        .type_extension()
        .ok_or_else(|| Error::Invalid("staged placeholder has no p232:phTypeExt scalar".into()))?;
    let span = shape.span()?.range(scene.xml().len())?;
    let (variant, variant_value) = locate_typed_variant(scene.xml(), span)?;
    if variant_value != value {
        return Err(Error::Invalid(
            "staged p232 placeholder metadata disagrees with source XML".into(),
        ));
    }
    let shape_name = shape.name().map(str::to_owned);
    let placeholder_index = placeholder.index();
    Ok(Snapshot {
        owner,
        source,
        shape_index,
        shape_name,
        placeholder_index,
        value,
        variant,
        limits,
    })
}

fn replace_variant(
    source: &[u8],
    variant: &VariantSpan,
    target: PlaceholderTypeExtension,
    limits: Limits,
) -> Result<Vec<u8>> {
    let replacement = target.as_str().as_bytes().to_vec();
    let mut replacements = vec![SourceReplacement {
        range: variant.open_name.clone(),
        bytes: replacement.clone(),
    }];
    if let Some(close_name) = &variant.close_name {
        replacements.push(SourceReplacement {
            range: close_name.clone(),
            bytes: replacement,
        });
    }
    apply_replacements(source, replacements, limits)
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct VariantSpan {
    open_name: Range<usize>,
    close_name: Option<Range<usize>>,
}

fn locate_typed_variant(
    source: &[u8],
    shape: Range<usize>,
) -> Result<(VariantSpan, PlaceholderTypeExtension)> {
    let (location, value) = locate_slot(source, shape)?;
    let SlotLocation::Existing { variant, .. } = location else {
        return Err(Error::Invalid(
            "selected placeholder has no p232 empty type token".into(),
        ));
    };
    Ok((
        variant,
        value.ok_or_else(|| {
            Error::Invalid("selected placeholder has no p232 empty type token".into())
        })?,
    ))
}

fn local_name_range(source: &[u8], start: usize, name: QName<'_>) -> Result<Range<usize>> {
    let raw = name.as_ref();
    let search_start = start.saturating_sub(256);
    let mut cursor = start.min(source.len());
    let (qname_start, qname_end) = loop {
        let search_end = cursor.min(source.len().saturating_sub(1));
        if search_start > search_end {
            return Err(Error::Invalid(
                "p232 token source does not start with an XML tag".into(),
            ));
        }
        let Some(tag_start) = source[search_start..=search_end]
            .iter()
            .rposition(|byte| *byte == b'<')
            .map(|offset| search_start + offset)
        else {
            return Err(Error::Invalid(
                "p232 token source does not start with an XML tag".into(),
            ));
        };
        let prefix_len = if source.get(tag_start + 1) == Some(&b'/') {
            2
        } else {
            1
        };
        let candidate_start = tag_start
            .checked_add(prefix_len)
            .ok_or_else(|| Error::Invalid("p232 token range overflows usize".into()))?;
        let candidate_end = candidate_start
            .checked_add(raw.len())
            .ok_or_else(|| Error::Invalid("p232 token range overflows usize".into()))?;
        if source.get(candidate_start..candidate_end) == Some(raw) {
            break (candidate_start, candidate_end);
        }
        if tag_start == search_start {
            return Err(Error::Invalid(
                "p232 token source spelling is not contiguous".into(),
            ));
        }
        cursor = tag_start.saturating_sub(1);
    };
    let local_start = raw
        .iter()
        .rposition(|byte| *byte == b':')
        .map_or(0, |position| position.saturating_add(1));
    Ok((qname_start + local_start)..qname_end)
}

fn validate_owner(part: &dyn Part) -> Result<()> {
    if matches!(
        part.content_type(),
        ct::PML_SLIDE | ct::PML_SLIDE_MASTER | ct::PML_SLIDE_LAYOUT
    ) {
        return Ok(());
    }
    Err(Error::ContentType {
        expected: format!(
            "{}, {}, or {}",
            ct::PML_SLIDE,
            ct::PML_SLIDE_MASTER,
            ct::PML_SLIDE_LAYOUT
        ),
        actual: part.content_type().to_owned(),
    })
}

fn is_p232(namespace: &ResolveResult<'_>, name: QName<'_>, local: &[u8]) -> bool {
    name.local_name().as_ref() == local
        && matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == P232)
}

fn is_pml(namespace: &ResolveResult<'_>, name: QName<'_>, local: &[u8]) -> bool {
    name.local_name().as_ref() == local
        && matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == PML || *value == STRICT_PML)
}

fn is_dml(namespace: &ResolveResult<'_>, name: QName<'_>, local: &[u8]) -> bool {
    name.local_name().as_ref() == local
        && matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == DML || *value == STRICT_DML)
}

impl PlaceholderTypeExtension {
    const fn as_str(self) -> &'static str {
        match self {
            Self::Cameo => "cameo",
            Self::Unknown => "unknown",
        }
    }
}

// ============================================================================
// Source-preserving add/remove transactions
// ============================================================================

/// A source snapshot of a placeholder's optional p232 extension.
///
/// Unlike [`Snapshot`], this state object also represents an absent extension.
/// It is the transaction entry point for adding, replacing, and removing the
/// extension on a placeholder in a slide, slide master, or slide layout.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SlotSnapshot {
    owner: PackURI,
    source: Arc<Vec<u8>>,
    shape_index: usize,
    shape_name: Option<String>,
    placeholder_index: u32,
    value: Option<PlaceholderTypeExtension>,
    location: SlotLocation,
    limits: Limits,
}

impl SlotSnapshot {
    /// Owning slide, master, or layout part.
    #[must_use]
    pub fn owner(&self) -> &PackURI {
        &self.owner
    }

    /// Exact owner XML captured by this snapshot.
    #[must_use]
    pub fn source_xml(&self) -> &[u8] {
        self.source.as_slice()
    }

    /// Checked pre-order shape index used by this source snapshot.
    #[must_use]
    pub const fn shape_index(&self) -> usize {
        self.shape_index
    }

    /// Shape name when the source supplied one.
    #[must_use]
    pub fn shape_name(&self) -> Option<&str> {
        self.shape_name.as_deref()
    }

    /// Placeholder index (`idx`, defaulting to zero).
    #[must_use]
    pub const fn placeholder_index(&self) -> u32 {
        self.placeholder_index
    }

    /// Current p232 scalar, or `None` when the extension is absent.
    #[must_use]
    pub const fn value(&self) -> Option<PlaceholderTypeExtension> {
        self.value
    }

    /// Whether the placeholder currently owns a typed p232 extension.
    #[must_use]
    pub const fn is_present(&self) -> bool {
        self.value.is_some()
    }

    /// Start a move-only add/remove/update edit tied to this exact source.
    #[must_use]
    pub fn edit(self) -> SlotEdit {
        SlotEdit {
            snapshot: self,
            target: None,
            target_set: false,
        }
    }

    fn same_source(&self, other: &Self) -> bool {
        self.owner == other.owner
            && self.source.as_slice() == other.source.as_slice()
            && self.shape_index == other.shape_index
            && self.shape_name == other.shape_name
            && self.placeholder_index == other.placeholder_index
            && self.value == other.value
    }
}

/// A move-only source-preserving p232 add/remove/update edit.
#[derive(Debug)]
pub struct SlotEdit {
    snapshot: SlotSnapshot,
    target: Option<PlaceholderTypeExtension>,
    target_set: bool,
}

impl SlotEdit {
    /// Value projected by this edit, or `None` for removal.
    #[must_use]
    pub const fn value(&self) -> Option<PlaceholderTypeExtension> {
        if self.target_set {
            self.target
        } else {
            self.snapshot.value
        }
    }

    /// Set or replace the p232 scalar.
    pub fn set(&mut self, value: PlaceholderTypeExtension) {
        self.target = Some(value);
        self.target_set = true;
    }

    /// Set the optional scalar. `None` removes the complete typed extension.
    pub fn set_optional(&mut self, value: Option<PlaceholderTypeExtension>) {
        self.target = value;
        self.target_set = true;
    }

    /// Add or replace the extension with `value`.
    pub fn add(&mut self, value: PlaceholderTypeExtension) {
        self.set(value);
    }

    /// Remove the extension.
    pub fn remove(&mut self) {
        self.target = None;
        self.target_set = true;
    }

    /// Whether this edit changes owner bytes.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.value() != self.snapshot.value
    }

    /// Consume this edit into a source-checked reversible commit.
    pub fn commit(self) -> Result<SlotCommit> {
        let target = self.value();
        if target == self.snapshot.value {
            return Ok(SlotCommit {
                snapshot: self.snapshot.clone(),
                patch: SlotPatch {
                    before: self.snapshot.clone(),
                    after: self.snapshot,
                },
            });
        }
        let staged = rewrite_slot(
            self.snapshot.source.as_slice(),
            &self.snapshot.location,
            target,
            self.snapshot.limits,
        )?;
        let after = slot_snapshot_from_source(
            self.snapshot.owner.clone(),
            Arc::new(staged),
            self.snapshot.shape_index,
            self.snapshot.limits,
        )?;
        if after.value != target {
            return Err(Error::Invalid(
                "staged p232 placeholder extension did not round-trip".into(),
            ));
        }
        Ok(SlotCommit {
            snapshot: after.clone(),
            patch: SlotPatch {
                before: self.snapshot,
                after,
            },
        })
    }
}

/// A successful source-preserving p232 add/remove/update commit.
#[derive(Clone, Debug)]
pub struct SlotCommit {
    snapshot: SlotSnapshot,
    patch: SlotPatch,
}

impl SlotCommit {
    /// Snapshot expected after publication.
    #[must_use]
    pub fn snapshot(&self) -> &SlotSnapshot {
        &self.snapshot
    }

    /// Reversible source patch represented by this commit.
    #[must_use]
    pub fn patch(&self) -> &SlotPatch {
        &self.patch
    }

    /// Whether this commit changes owner bytes.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.patch.is_changed()
    }

    /// Whether this commit is an exact byte-preserving no-op.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        !self.is_changed()
    }

    /// Consume this commit into its reversible patch.
    #[must_use]
    pub fn into_patch(self) -> SlotPatch {
        self.patch
    }
}

/// A source-checked, reversible p232 add/remove/update patch.
#[derive(Clone, Debug)]
pub struct SlotPatch {
    before: SlotSnapshot,
    after: SlotSnapshot,
}

impl SlotPatch {
    /// Snapshot required before forward application.
    #[must_use]
    pub fn before(&self) -> &SlotSnapshot {
        &self.before
    }

    /// Snapshot expected after forward application.
    #[must_use]
    pub fn after(&self) -> &SlotSnapshot {
        &self.after
    }

    /// Whether this patch is an exact byte-preserving no-op.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before.same_source(&self.after)
    }

    /// Whether this patch changes owner bytes.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        !self.is_noop()
    }

    /// Return the exact inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

    /// Apply with the default signature policy, rejecting changed signed input.
    pub fn apply(&self, package: &mut OpcPackage) -> Result<SlotSnapshot> {
        apply_slot_patch(package, self)
    }

    /// Apply with an explicit signature disposition.
    pub fn apply_with_policy(
        &self,
        package: &mut OpcPackage,
        policy: SignaturePolicy,
    ) -> Result<SlotSnapshot> {
        apply_slot_patch_with_policy(package, self, policy)
    }
}

/// Load an optional p232 extension from a slide, master, or layout placeholder.
pub fn load_slot_snapshot<'a>(
    package: &OpcPackage,
    owner: &PackURI,
    key: impl Into<Key<'a>>,
) -> Result<SlotSnapshot> {
    load_slot_snapshot_with_limits(package, owner, key, Limits::DEFAULT)
}

/// Load an optional p232 extension under finite source and growth limits.
pub fn load_slot_snapshot_with_limits<'a>(
    package: &OpcPackage,
    owner: &PackURI,
    key: impl Into<Key<'a>>,
    limits: Limits,
) -> Result<SlotSnapshot> {
    let part = package.get_part(owner)?;
    validate_owner(part)?;
    slot_snapshot_from_source_for_key(part.partname().clone(), part.blob_arc(), key.into(), limits)
}

/// Apply an add/remove/update patch with the default signature policy.
pub fn apply_slot_patch(package: &mut OpcPackage, patch: &SlotPatch) -> Result<SlotSnapshot> {
    apply_slot_patch_with_policy(package, patch, SignaturePolicy::Reject)
}

/// Apply an add/remove/update patch after exact source, semantic, bound, and
/// explicit signature checks.
pub fn apply_slot_patch_with_policy(
    package: &mut OpcPackage,
    patch: &SlotPatch,
    policy: SignaturePolicy,
) -> Result<SlotSnapshot> {
    let part = package.get_part(&patch.before.owner)?;
    validate_owner(part)?;
    patch.before.limits.check_source(part.blob().len())?;
    if part.blob() != patch.before.source.as_slice() {
        return Err(Error::StaleSource);
    }
    if patch.is_noop() {
        return Ok(patch.before.clone());
    }
    patch
        .before
        .limits
        .check_growth(patch.before.source.len(), patch.after.source.len())?;
    if matches!(policy, SignaturePolicy::Reject)
        && (package.is_signed() || package.requires_signature_edit_policy())
    {
        return Err(Error::Opc(
            litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy,
        ));
    }

    // All checks happen before the package clone or its first mutation.
    let mut candidate = package.clone();
    candidate
        .get_part_mut(&patch.after.owner)?
        .set_blob_shared(Arc::clone(&patch.after.source));
    let published = load_slot_snapshot_with_limits(
        &candidate,
        &patch.after.owner,
        Key::Index(patch.after.shape_index),
        patch.after.limits,
    )?;
    if !published.same_source(&patch.after) {
        return Err(Error::Invalid(
            "published p232 placeholder slot differs from the commit".into(),
        ));
    }
    if matches!(policy, SignaturePolicy::Invalidate) {
        candidate.unsign();
    }
    *package = candidate;
    Ok(published)
}

/// Apply a committed add/remove/update patch with the default signature policy.
pub fn apply_slot_commit(package: &mut OpcPackage, commit: SlotCommit) -> Result<SlotSnapshot> {
    apply_slot_patch(package, &commit.patch)
}

/// Store one placeholder shape while preserving the bytes of an existing
/// shape and every unknown child it contains.
///
/// The legacy authoring operation historically replaced an entire matching
/// shape.  That loses extension markup owned by producers this crate does not
/// understand.  This implementation changes only the requested name, text,
/// and p232 slot; a new shape is appended only when the requested placeholder
/// does not exist.  Slide, slide-master, and slide-layout owners are all
/// accepted.
pub fn store_placeholder_shape_source(
    package: &mut OpcPackage,
    part_name: &PackURI,
    spec: &PlaceholderSpec,
) -> Result<()> {
    validate_placeholder_spec(spec)?;
    let part = package.get_part(part_name)?;
    validate_owner(part)?;
    let limits = Limits::DEFAULT;
    let blob = part.blob();
    limits.check_source(blob.len())?;
    let mut source = Vec::new();
    source
        .try_reserve_exact(blob.len())
        .map_err(|source| Error::Allocation {
            resource: "p232 placeholder owner XML bytes",
            source,
        })?;
    source.extend_from_slice(blob);
    let mut patched = prepare_placeholder_shape_source(&source, spec, limits)?;

    if patched == source {
        return Ok(());
    }
    package
        .get_part_mut(part_name)?
        .set_blob(std::mem::take(&mut patched));
    // This compatibility authoring API has always invalidated signatures on a
    // changed owner.  Source checked slot transactions above expose the
    // explicit Reject/Invalidate policy for callers that need that choice.
    package.unsign();
    Ok(())
}

/// Materialize the exact owner bytes that the legacy facade would publish.
/// This is deliberately callable against a borrowed source so the package
/// facade can run every parse, escaped-output, and final-size check before its
/// generic transaction clones the OPC graph.
fn prepare_placeholder_shape_source(
    source: &[u8],
    spec: &PlaceholderSpec,
    limits: Limits,
) -> Result<Vec<u8>> {
    validate_placeholder_spec(spec)?;
    limits.check_source(source.len())?;
    let scene = Scene::read(source)?;
    if scene.is_rewritten() {
        return Err(Error::UnsafeEdit {
            operation: "store_placeholder_shape",
            reason: "markup compatibility rewrote the owner; source-preserving shape authoring is unavailable",
        });
    }

    let mut patched = if let Some((shape_index, shape)) = find_placeholder_shape(&scene, spec)? {
        let span = shape.span()?.range(scene.xml().len())?;
        let (location, source_value) = locate_slot(scene.xml(), span)?;
        let placeholder = shape
            .placeholder()
            .ok_or_else(|| invalid("selected shape is not a placeholder"))?;
        if source_value != placeholder.type_extension() {
            return Err(invalid(
                "typed p232 placeholder metadata disagrees with source XML",
            ));
        }
        let typed = if source_value == spec.type_extension {
            Cow::Borrowed(source)
        } else {
            Cow::Owned(rewrite_slot(
                source,
                &location,
                spec.type_extension,
                limits,
            )?)
        };
        patch_existing_shape_fields(typed.as_ref(), shape_index, spec, limits)?
    } else {
        let tree = scan_element_span(source, "spTree", SPTREE_DEPTH)?
            .ok_or_else(|| invalid("owner has no shape tree"))?;
        if tree.empty {
            return Err(invalid("owner has an empty shape tree"));
        }
        let shape_id = next_shape_id(source)?;
        let shape = placeholder_shape_xml(shape_id, spec, true)?;
        insert_bytes(source, tree.close_start, shape.as_bytes())?
    };

    limits.check_source(patched.len())?;
    let final_scene = Scene::read(&patched)?;
    if final_scene.is_rewritten() {
        return Err(Error::UnsafeEdit {
            operation: "store_placeholder_shape",
            reason: "staged owner requires markup compatibility rewriting",
        });
    }
    let (_, shape) = find_placeholder_shape(&final_scene, spec)?
        .ok_or_else(|| invalid("authored placeholder shape did not round-trip"))?;
    let placeholder = shape
        .placeholder()
        .ok_or_else(|| invalid("authored shape lost its placeholder metadata"))?;
    if placeholder.type_extension() != spec.type_extension {
        return Err(invalid(
            "authored placeholder p232 metadata did not round-trip",
        ));
    }
    if let Some(name) = spec.name.as_deref()
        && shape.name() != Some(name)
    {
        return Err(invalid("authored placeholder name did not round-trip"));
    }
    if let Some(text) = spec.text.as_deref()
        && shape.text() != Some(text)
    {
        return Err(invalid("authored placeholder text did not round-trip"));
    }
    Ok(std::mem::take(&mut patched))
}

/// Validate the bounded inputs that the package facade must check before its
/// generic transaction snapshots the whole OPC graph.
pub(crate) fn preflight_store_placeholder_shape(
    package: &OpcPackage,
    part_name: &PackURI,
    spec: &PlaceholderSpec,
) -> Result<bool> {
    validate_placeholder_spec(spec)?;
    let part = package.get_part(part_name)?;
    validate_owner(part)?;
    let staged = prepare_placeholder_shape_source(part.blob(), spec, Limits::DEFAULT)?;
    Ok(staged.as_slice() == part.blob())
}

fn validate_placeholder_spec(spec: &PlaceholderSpec) -> Result<()> {
    if let Some(name) = spec.name.as_deref() {
        if name.is_empty() || name.chars().count() > MAX_NAME_CHARS {
            return Err(invalid("placeholder name must contain 1..256 characters"));
        }
        validate_xml10_chars(name, "placeholder name")?;
    }
    if let Some(text) = spec.text.as_deref() {
        validate_xml10_chars(text, "placeholder text")?;
    }
    for value in [spec.name.as_deref(), spec.text.as_deref()]
        .into_iter()
        .flatten()
    {
        if escaped_xml_len(value)? > MAX_OWNER_BYTES {
            return Err(Error::Limit {
                resource: "p232 placeholder escaped XML bytes",
                limit: MAX_OWNER_BYTES,
            });
        }
    }
    Ok(())
}

fn find_placeholder_shape<'a>(
    scene: &'a Scene<'a>,
    spec: &PlaceholderSpec,
) -> Result<Option<(usize, crate::shape::Shape<'a>)>> {
    for shape in scene.iter() {
        let Some(placeholder) = shape.placeholder() else {
            continue;
        };
        if placeholder.kind().unwrap_or("obj") == spec.kind.as_str()
            && placeholder.index() == spec.effective_index()
        {
            return Ok(Some((shape.common().index(), shape)));
        }
    }
    Ok(None)
}

#[derive(Debug)]
struct SourceReplacement {
    range: Range<usize>,
    bytes: Vec<u8>,
}

fn patch_existing_shape_fields(
    source: &[u8],
    shape_index: usize,
    spec: &PlaceholderSpec,
    limits: Limits,
) -> Result<Vec<u8>> {
    let scene = Scene::read(source)?;
    if scene.is_rewritten() {
        return Err(Error::UnsafeEdit {
            operation: "store_placeholder_shape",
            reason: "staged owner requires markup compatibility rewriting",
        });
    }
    let shape = scene.at(shape_index)?;
    let span = shape.span()?.range(scene.xml().len())?;
    let mut replacements = Vec::new();

    if let Some(name) = spec.name.as_deref() {
        let tag = find_first_start_tag(scene.xml(), span.clone(), b"cNvPr")?
            .ok_or_else(|| invalid("placeholder shape has no p:cNvPr"))?;
        let encoded = escape_xml_bounded(name, limits)?;
        if let Some(value) = attribute_value_range(scene.xml(), tag.start, tag.end, b"name")? {
            if scene.xml().get(value.clone()) != Some(encoded.as_slice()) {
                replacements.push(SourceReplacement {
                    range: value,
                    bytes: encoded,
                });
            }
        } else {
            let offset = attribute_insert_offset(scene.xml(), tag.start, tag.end)?;
            let mut bytes =
                bytes_with_capacity(encoded.len() + 9, "p232 placeholder name insertion bytes")?;
            bytes.extend_from_slice(b" name=\"");
            bytes.extend_from_slice(&encoded);
            bytes.push(b'\"');
            replacements.push(SourceReplacement {
                range: offset..offset,
                bytes,
            });
        }
    }

    if let Some(text) = spec.text.as_deref() {
        let encoded = escape_xml_bounded(text, limits)?;
        match locate_text_content(scene.xml(), span)? {
            TextLocation::Runs(runs) if !runs.is_empty() => {
                for (index, run) in runs.into_iter().enumerate() {
                    if index == 0 && run.empty {
                        let mut bytes = bytes_with_capacity(
                            encoded.len() + run.prefix.len() * 2 + 7,
                            "p232 placeholder text replacement bytes",
                        )?;
                        push_open(&mut bytes, &run.prefix, b"t");
                        bytes.extend_from_slice(&encoded);
                        push_close(&mut bytes, &run.prefix, b"t");
                        replacements.push(SourceReplacement {
                            range: run.element,
                            bytes,
                        });
                    } else {
                        let bytes = if index == 0 {
                            encoded.clone()
                        } else {
                            Vec::new()
                        };
                        if scene.xml().get(run.content.clone()) != Some(bytes.as_slice()) {
                            replacements.push(SourceReplacement {
                                range: run.content,
                                bytes,
                            });
                        }
                    }
                }
            },
            TextLocation::Runs(_) => {
                return Err(invalid("placeholder text run scan returned no runs"));
            },
            TextLocation::Insert { offset, prefix } => {
                let mut bytes = bytes_with_capacity(
                    encoded.len() + prefix.len() * 4 + 15,
                    "p232 placeholder text insertion bytes",
                )?;
                push_open(&mut bytes, &prefix, b"r");
                push_open(&mut bytes, &prefix, b"t");
                bytes.extend_from_slice(&encoded);
                push_close(&mut bytes, &prefix, b"t");
                push_close(&mut bytes, &prefix, b"r");
                replacements.push(SourceReplacement {
                    range: offset..offset,
                    bytes,
                });
            },
        }
    }

    apply_replacements(source, replacements, limits)
}

fn apply_replacements(
    source: &[u8],
    mut replacements: Vec<SourceReplacement>,
    limits: Limits,
) -> Result<Vec<u8>> {
    replacements.sort_by_key(|replacement| replacement.range.start);
    let mut target_len = source.len();
    let mut previous_end = 0usize;
    for replacement in &replacements {
        if replacement.range.start > replacement.range.end
            || replacement.range.end > source.len()
            || replacement.range.start < previous_end
        {
            return Err(invalid("overlapping source-preserving shape replacements"));
        }
        target_len = target_len
            .checked_sub(replacement.range.end - replacement.range.start)
            .and_then(|value| value.checked_add(replacement.bytes.len()))
            .ok_or(Error::Limit {
                resource: "p232 placeholder shape output bytes",
                limit: limits.owner_bytes,
            })?;
        previous_end = replacement.range.end;
    }
    limits.check_growth(source.len(), target_len)?;
    if replacements.is_empty() {
        let mut output = bytes_with_capacity(source.len(), "p232 placeholder shape output bytes")?;
        output.extend_from_slice(source);
        return Ok(output);
    }
    let mut output = bytes_with_capacity(target_len, "p232 placeholder shape output bytes")?;
    let mut cursor = 0usize;
    for replacement in replacements {
        output.extend_from_slice(&source[cursor..replacement.range.start]);
        output.extend_from_slice(&replacement.bytes);
        cursor = replacement.range.end;
    }
    output.extend_from_slice(&source[cursor..]);
    Ok(output)
}

fn escaped_xml_len(value: &str) -> Result<usize> {
    value.chars().try_fold(0usize, |length, character| {
        let extra = match character {
            '&' => 5,
            '<' | '>' => 4,
            '"' | '\'' => 6,
            '\t' | '\n' | '\r' => 5,
            _ => character.len_utf8(),
        };
        length.checked_add(extra).ok_or(Error::Limit {
            resource: "p232 placeholder escaped XML bytes",
            limit: MAX_OWNER_BYTES,
        })
    })
}

fn escape_xml_bounded(value: &str, limits: Limits) -> Result<Vec<u8>> {
    let length = escaped_xml_len(value)?;
    if length > limits.owner_bytes {
        return Err(Error::Limit {
            resource: "p232 placeholder escaped XML bytes",
            limit: limits.owner_bytes,
        });
    }
    let mut escaped = Vec::new();
    escaped
        .try_reserve_exact(length)
        .map_err(|source| Error::Allocation {
            resource: "p232 placeholder escaped XML bytes",
            source,
        })?;
    for character in value.chars() {
        match character {
            '&' => escaped.extend_from_slice(b"&amp;"),
            '<' => escaped.extend_from_slice(b"&lt;"),
            '>' => escaped.extend_from_slice(b"&gt;"),
            '"' => escaped.extend_from_slice(b"&quot;"),
            '\'' => escaped.extend_from_slice(b"&apos;"),
            '\t' => escaped.extend_from_slice(b"&#x9;"),
            '\n' => escaped.extend_from_slice(b"&#xA;"),
            '\r' => escaped.extend_from_slice(b"&#xD;"),
            _ => {
                let mut buffer = [0; 4];
                escaped.extend_from_slice(character.encode_utf8(&mut buffer).as_bytes());
            },
        }
    }
    Ok(escaped)
}

fn bytes_with_capacity(capacity: usize, resource: &'static str) -> Result<Vec<u8>> {
    let mut bytes = Vec::new();
    bytes
        .try_reserve_exact(capacity)
        .map_err(|source| Error::Allocation { resource, source })?;
    Ok(bytes)
}

#[derive(Clone, Debug)]
struct StartTag {
    start: usize,
    end: usize,
}

fn find_first_start_tag(
    source: &[u8],
    shape: Range<usize>,
    local: &[u8],
) -> Result<Option<StartTag>> {
    let mut reader = NsReader::from_reader(source);
    loop {
        let start = usize::try_from(reader.buffer_position())
            .map_err(|_| invalid("shape source position exceeds usize"))?;
        let event = reader.read_event()?.into_owned();
        let end = usize::try_from(reader.buffer_position())
            .map_err(|_| invalid("shape source position exceeds usize"))?;
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) | Event::Empty(element)
                if start >= shape.start
                    && end <= shape.end
                    && is_pml(&namespace, element.name(), local) =>
            {
                return Ok(Some(StartTag { start, end }));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(None)
}

fn attribute_value_range(
    source: &[u8],
    start: usize,
    end: usize,
    wanted: &[u8],
) -> Result<Option<Range<usize>>> {
    let bytes = source
        .get(start..end)
        .ok_or_else(|| invalid("attribute start tag is outside its owner"))?;
    if bytes.first() != Some(&b'<') || bytes.last() != Some(&b'>') {
        return Err(invalid("invalid source-preserving start tag"));
    }
    let mut cursor = 1usize;
    while cursor + 1 < bytes.len()
        && !bytes[cursor].is_ascii_whitespace()
        && !matches!(bytes[cursor], b'>' | b'/')
    {
        cursor += 1;
    }
    while cursor + 1 < bytes.len() {
        while cursor + 1 < bytes.len() && bytes[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if cursor + 1 >= bytes.len() || bytes[cursor] == b'>' || bytes[cursor] == b'/' {
            break;
        }
        let key_start = cursor;
        while cursor + 1 < bytes.len()
            && !bytes[cursor].is_ascii_whitespace()
            && !matches!(bytes[cursor], b'=' | b'>' | b'/')
        {
            cursor += 1;
        }
        let key = &bytes[key_start..cursor];
        while cursor + 1 < bytes.len() && bytes[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if bytes.get(cursor) != Some(&b'=') {
            return Err(invalid("source attribute has no equals sign"));
        }
        cursor += 1;
        while cursor + 1 < bytes.len() && bytes[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        let quote = *bytes
            .get(cursor)
            .ok_or_else(|| invalid("source attribute has no quoted value"))?;
        if quote != b'\'' && quote != b'"' {
            return Err(invalid("source attribute value is not quoted"));
        }
        cursor += 1;
        let value_start = cursor;
        while cursor < bytes.len() && bytes[cursor] != quote {
            cursor += 1;
        }
        if cursor >= bytes.len() {
            return Err(invalid("unterminated source attribute value"));
        }
        if key == wanted {
            return Ok(Some((start + value_start)..(start + cursor)));
        }
        cursor += 1;
    }
    Ok(None)
}

fn attribute_insert_offset(source: &[u8], start: usize, end: usize) -> Result<usize> {
    let bytes = source
        .get(start..end)
        .ok_or_else(|| invalid("attribute start tag is outside its owner"))?;
    if bytes.len() < 2 || bytes.first() != Some(&b'<') || bytes.last() != Some(&b'>') {
        return Err(invalid("invalid source-preserving start tag"));
    }
    if bytes.get(bytes.len().saturating_sub(2)) == Some(&b'/') {
        Ok(end - 2)
    } else {
        Ok(end - 1)
    }
}

#[derive(Debug)]
enum TextLocation {
    Runs(Vec<TextRunLocation>),
    Insert { offset: usize, prefix: Vec<u8> },
}

#[derive(Debug)]
struct TextRunLocation {
    element: Range<usize>,
    content: Range<usize>,
    empty: bool,
    prefix: Vec<u8>,
}

fn locate_text_content(source: &[u8], shape: Range<usize>) -> Result<TextLocation> {
    let mut reader = NsReader::from_reader(source);
    let mut runs = Vec::new();
    let mut open_text = None;
    let mut end_para = None;
    let mut paragraph_close = None;
    let mut paragraph_prefix = Vec::new();
    let mut tx_body_close = None;
    let mut tx_body_prefix = Vec::new();

    loop {
        let start = usize::try_from(reader.buffer_position())
            .map_err(|_| invalid("shape source position exceeds usize"))?;
        let event = reader.read_event()?.into_owned();
        let end = usize::try_from(reader.buffer_position())
            .map_err(|_| invalid("shape source position exceeds usize"))?;
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) if start >= shape.start && end <= shape.end => {
                let raw = element.name().as_ref().to_vec();
                let prefix = raw_prefix(&raw);
                if is_dml(&namespace, element.name(), b"t") {
                    if open_text.is_some() {
                        return Err(invalid("nested DrawingML text elements"));
                    }
                    open_text = Some((start, end, prefix));
                } else if is_dml(&namespace, element.name(), b"endParaRPr") && end_para.is_none() {
                    end_para = Some((start, raw_prefix(&raw)));
                } else if is_dml(&namespace, element.name(), b"p") {
                    paragraph_prefix = prefix;
                } else if is_dml(&namespace, element.name(), b"txBody") {
                    tx_body_prefix = prefix;
                }
            },
            Event::Empty(element) if start >= shape.start && end <= shape.end => {
                let raw = element.name().as_ref().to_vec();
                if is_dml(&namespace, element.name(), b"t") {
                    runs.push(TextRunLocation {
                        element: start..end,
                        content: start..start,
                        empty: true,
                        prefix: raw_prefix(&raw),
                    });
                } else if is_dml(&namespace, element.name(), b"endParaRPr") && end_para.is_none() {
                    end_para = Some((start, raw_prefix(&raw)));
                }
            },
            Event::End(element) if start >= shape.start && end <= shape.end => {
                if is_dml(&namespace, element.name(), b"t") {
                    let (element_start, content_start, prefix) = open_text
                        .take()
                        .ok_or_else(|| invalid("text source range became inconsistent"))?;
                    runs.push(TextRunLocation {
                        element: element_start..end,
                        content: content_start..start,
                        empty: false,
                        prefix,
                    });
                } else if is_dml(&namespace, element.name(), b"p") && paragraph_close.is_none() {
                    paragraph_close = Some((start, paragraph_prefix.clone()));
                } else if is_dml(&namespace, element.name(), b"txBody") && tx_body_close.is_none() {
                    tx_body_close = Some((start, tx_body_prefix.clone()));
                }
                if end >= shape.end {
                    break;
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }

    if !runs.is_empty() {
        return Ok(TextLocation::Runs(runs));
    }
    if let Some((offset, prefix)) = end_para {
        return Ok(TextLocation::Insert { offset, prefix });
    }
    if let Some((offset, prefix)) = paragraph_close {
        return Ok(TextLocation::Insert { offset, prefix });
    }
    if let Some((offset, prefix)) = tx_body_close {
        return Ok(TextLocation::Insert { offset, prefix });
    }
    Err(invalid(
        "placeholder shape has no DrawingML text insertion point",
    ))
}

fn slot_snapshot_from_source_for_key<'a>(
    owner: PackURI,
    source: Arc<Vec<u8>>,
    key: Key<'a>,
    limits: Limits,
) -> Result<SlotSnapshot> {
    limits.check_source(source.len())?;
    let scene = Scene::read(source.as_slice())?;
    if scene.is_rewritten() {
        return Err(Error::UnsafeEdit {
            operation: "load_p232_placeholder_slot",
            reason: "markup compatibility rewrote the owner; no source-preserving patch policy is available",
        });
    }
    let shape = scene.shape(key)?;
    let placeholder = shape
        .placeholder()
        .ok_or_else(|| Error::Invalid("selected shape is not a placeholder".into()))?;
    let span = shape.span()?.range(scene.xml().len())?;
    let (location, source_value) = locate_slot(scene.xml(), span)?;
    if source_value != placeholder.type_extension() {
        return Err(Error::Invalid(
            "typed p232 placeholder metadata disagrees with source XML".into(),
        ));
    }
    let shape_index = shape.common().index();
    let shape_name = shape.name().map(str::to_owned);
    let placeholder_index = placeholder.index();
    Ok(SlotSnapshot {
        owner,
        source,
        shape_index,
        shape_name,
        placeholder_index,
        value: source_value,
        location,
        limits,
    })
}

fn slot_snapshot_from_source(
    owner: PackURI,
    source: Arc<Vec<u8>>,
    shape_index: usize,
    limits: Limits,
) -> Result<SlotSnapshot> {
    let scene = Scene::read(source.as_slice())?;
    if scene.is_rewritten() {
        return Err(Error::UnsafeEdit {
            operation: "stage_p232_placeholder_slot",
            reason: "markup compatibility rewrote the owner; no source-preserving patch policy is available",
        });
    }
    let shape = scene.at(shape_index)?;
    let placeholder = shape
        .placeholder()
        .ok_or_else(|| Error::Invalid("staged shape is no longer a placeholder".into()))?;
    let span = shape.span()?.range(scene.xml().len())?;
    let (location, source_value) = locate_slot(scene.xml(), span)?;
    if source_value != placeholder.type_extension() {
        return Err(Error::Invalid(
            "staged p232 placeholder metadata disagrees with source XML".into(),
        ));
    }
    let shape_name = shape.name().map(str::to_owned);
    let placeholder_index = placeholder.index();
    Ok(SlotSnapshot {
        owner,
        source,
        shape_index,
        shape_name,
        placeholder_index,
        value: source_value,
        location,
        limits,
    })
}

#[derive(Clone, Debug, PartialEq, Eq)]
enum SlotLocation {
    Existing {
        extension: ElementRange,
        variant: VariantSpan,
    },
    Missing {
        placeholder: ElementRange,
        extension_list: Option<ElementRange>,
    },
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct ElementRange {
    start: usize,
    end: usize,
    close_start: usize,
    empty: bool,
    prefix: Vec<u8>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum ScanKind {
    Other,
    Placeholder,
    ExtensionList,
    Extension,
    TypedExtension,
    Type,
    Variant,
}

#[derive(Clone, Debug)]
struct ScanFrame {
    start: usize,
    prefix: Vec<u8>,
    kind: ScanKind,
    is_nv_pr: bool,
    extension_uri: Option<String>,
    variant: Option<(VariantSpan, PlaceholderTypeExtension)>,
}

fn locate_slot(
    source: &[u8],
    shape: Range<usize>,
) -> Result<(SlotLocation, Option<PlaceholderTypeExtension>)> {
    let mut reader = NsReader::from_reader(source);
    let mut stack = Vec::<ScanFrame>::new();
    let mut placeholder = None;
    let mut extension_list = None;
    let mut extension = None;
    let mut variant = None;
    let mut variant_value = None;

    loop {
        let start = usize::try_from(reader.buffer_position())
            .map_err(|_| Error::Invalid("p232 source position exceeds usize".into()))?;
        let event = reader.read_event()?.into_owned();
        let end = usize::try_from(reader.buffer_position())
            .map_err(|_| Error::Invalid("p232 source position exceeds usize".into()))?;
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) => {
                let raw_name = element.name().as_ref().to_vec();
                let prefix = raw_prefix(&raw_name);
                let parent_is_nv_pr = stack.last().is_some_and(|frame| frame.is_nv_pr);
                let parent = stack.last();
                let extension_uri = if is_pml(&namespace, element.name(), b"ext")
                    && parent.is_some_and(|frame| frame.kind == ScanKind::ExtensionList)
                {
                    Some(validate_extension_uri(&element, reader.decoder())?)
                } else {
                    None
                };
                let kind = scan_kind(
                    &namespace,
                    element.name(),
                    parent,
                    parent_is_nv_pr,
                    placeholder.is_none(),
                );
                let variant = if kind == ScanKind::Variant {
                    if let Some(value) = variant_value_for(&namespace, element.name()) {
                        Some((
                            VariantSpan {
                                open_name: local_name_range(source, start, element.name())?,
                                close_name: None,
                            },
                            value,
                        ))
                    } else {
                        None
                    }
                } else {
                    None
                };
                if kind == ScanKind::Placeholder && start >= shape.start && end <= shape.end {
                    placeholder = Some(ElementRange {
                        start,
                        end: 0,
                        close_start: 0,
                        empty: false,
                        prefix: prefix.clone(),
                    });
                }
                if kind == ScanKind::ExtensionList && start >= shape.start && end <= shape.end {
                    extension_list = Some(ElementRange {
                        start,
                        end: 0,
                        close_start: 0,
                        empty: false,
                        prefix: prefix.clone(),
                    });
                }
                stack.push(ScanFrame {
                    start,
                    prefix,
                    kind,
                    is_nv_pr: is_pml(&namespace, element.name(), b"nvPr"),
                    extension_uri,
                    variant,
                });
            },
            Event::Empty(element) => {
                let raw_name = element.name().as_ref().to_vec();
                let prefix = raw_prefix(&raw_name);
                let parent_is_nv_pr = stack.last().is_some_and(|frame| frame.is_nv_pr);
                let parent = stack.last();
                let extension_uri = if is_pml(&namespace, element.name(), b"ext")
                    && parent.is_some_and(|frame| frame.kind == ScanKind::ExtensionList)
                {
                    Some(validate_extension_uri(&element, reader.decoder())?)
                } else {
                    None
                };
                let kind = scan_kind(
                    &namespace,
                    element.name(),
                    parent,
                    parent_is_nv_pr,
                    placeholder.is_none(),
                );
                let range = ElementRange {
                    start,
                    end,
                    close_start: start,
                    empty: true,
                    prefix: prefix.clone(),
                };
                match kind {
                    ScanKind::Placeholder
                        if start >= shape.start && end <= shape.end && placeholder.is_none() =>
                    {
                        placeholder = Some(range);
                    },
                    ScanKind::ExtensionList if start >= shape.start && end <= shape.end => {
                        extension_list = Some(range)
                    },
                    ScanKind::Variant => {
                        if start >= shape.start
                            && end <= shape.end
                            && let Some(value) = variant_value_for(&namespace, element.name())
                        {
                            if let Some(parent) = stack.last_mut() {
                                parent.variant = Some((
                                    VariantSpan {
                                        open_name: local_name_range(source, start, element.name())?,
                                        close_name: None,
                                    },
                                    value,
                                ));
                            }
                        }
                    },
                    ScanKind::Extension => {
                        let _ = extension_uri;
                    },
                    _ => {},
                }
            },
            Event::End(element) => {
                let frame = stack.pop().ok_or_else(|| {
                    Error::Invalid("p232 source stack became inconsistent".into())
                })?;
                let range = ElementRange {
                    start: frame.start,
                    end,
                    close_start: start,
                    empty: false,
                    prefix: frame.prefix,
                };
                match frame.kind {
                    ScanKind::Placeholder if frame.start >= shape.start && end <= shape.end => {
                        if let Some(found) = placeholder.as_mut() {
                            if found.start == frame.start {
                                *found = range;
                            }
                        }
                    },
                    ScanKind::ExtensionList if frame.start >= shape.start && end <= shape.end => {
                        extension_list = Some(range)
                    },
                    ScanKind::Extension => {
                        if frame.start >= shape.start
                            && end <= shape.end
                            && let Some((variant_span, value)) = frame.variant
                        {
                            extension = Some(range);
                            variant = Some(variant_span.clone());
                            variant_value = Some(value);
                            if let Some(parent) = stack.last_mut() {
                                parent.variant = Some((variant_span, value));
                            }
                        }
                    },
                    ScanKind::TypedExtension | ScanKind::Type => {
                        if let Some(variant) = frame.variant
                            && let Some(parent) = stack.last_mut()
                        {
                            parent.variant = Some(variant);
                        }
                    },
                    ScanKind::Variant => {
                        if let Some((mut variant, value)) = frame.variant {
                            variant.close_name =
                                Some(local_name_range(source, start, element.name())?);
                            if let Some(parent) = stack.last_mut() {
                                parent.variant = Some((variant, value));
                            }
                        }
                    },
                    _ => {},
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }

    let placeholder = placeholder
        .ok_or_else(|| Error::Invalid("selected shape has no direct p:ph element".into()))?;
    if let (Some(extension), Some(variant), Some(value)) = (extension, variant, variant_value) {
        return Ok((SlotLocation::Existing { extension, variant }, Some(value)));
    }
    Ok((
        SlotLocation::Missing {
            placeholder,
            extension_list,
        },
        None,
    ))
}

fn scan_kind(
    namespace: &ResolveResult<'_>,
    name: QName<'_>,
    parent: Option<&ScanFrame>,
    parent_is_nv_pr: bool,
    placeholder_absent: bool,
) -> ScanKind {
    let parent_kind = parent.map(|frame| frame.kind);
    if is_pml(namespace, name, b"ph") && parent_is_nv_pr && placeholder_absent {
        ScanKind::Placeholder
    } else if is_pml(namespace, name, b"extLst") && parent_kind == Some(ScanKind::Placeholder) {
        ScanKind::ExtensionList
    } else if is_pml(namespace, name, b"ext") && parent_kind == Some(ScanKind::ExtensionList) {
        ScanKind::Extension
    } else if is_p232(namespace, name, b"phTypeExt")
        && parent_kind == Some(ScanKind::Extension)
        && parent.is_some_and(|frame| {
            frame.extension_uri.as_deref() == Some(PLACEHOLDER_TYPE_EXTENSION_URI)
        })
    {
        ScanKind::TypedExtension
    } else if is_p232(namespace, name, b"type") && parent_kind == Some(ScanKind::TypedExtension) {
        ScanKind::Type
    } else if parent_kind == Some(ScanKind::Type) && variant_value_for(namespace, name).is_some() {
        ScanKind::Variant
    } else {
        ScanKind::Other
    }
}

fn variant_value_for(
    namespace: &ResolveResult<'_>,
    name: QName<'_>,
) -> Option<PlaceholderTypeExtension> {
    if is_p232(namespace, name, b"cameo") {
        Some(PlaceholderTypeExtension::Cameo)
    } else if is_p232(namespace, name, b"unknown") {
        Some(PlaceholderTypeExtension::Unknown)
    } else {
        None
    }
}

fn raw_prefix(name: &[u8]) -> Vec<u8> {
    name.iter()
        .rposition(|byte| *byte == b':')
        .map_or_else(Vec::new, |position| name[..position].to_vec())
}

fn validate_extension_uri(element: &BytesStart<'_>, decoder: Decoder) -> Result<String> {
    let uri = unqualified_attribute_value(element, b"uri", decoder)?;
    let mut seen_uri = false;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        let key = attribute.key.as_ref();
        if key == b"xmlns" || key.starts_with(b"xmlns:") {
            continue;
        }
        if key != b"uri" || seen_uri {
            return Err(Error::Invalid(
                "p:ext allows exactly one unqualified uri attribute".into(),
            ));
        }
        seen_uri = true;
    }
    let Some(uri) = uri else {
        return Err(Error::Invalid(
            "p:ext is missing its required uri attribute".into(),
        ));
    };
    collapse_xsd_token(&uri)
        .ok_or_else(|| Error::Invalid("p:ext uri must be a nonempty XML Schema token".into()))
}

/// Apply the XML Schema `token` whitespace facet while leaving the source
/// spelling available to the source-preserving transaction.  The normalized
/// value is used only to classify the stable p232 owner URI.
fn collapse_xsd_token(value: &str) -> Option<String> {
    let mut collapsed = String::new();
    let mut pending_space = false;
    for character in value.chars() {
        if matches!(character, ' ' | '\t' | '\r' | '\n') {
            if !collapsed.is_empty() {
                pending_space = true;
            }
            continue;
        }
        if pending_space {
            collapsed.push(' ');
            pending_space = false;
        }
        collapsed.push(character);
    }
    (!collapsed.is_empty()).then_some(collapsed)
}

fn rewrite_slot(
    source: &[u8],
    location: &SlotLocation,
    target: Option<PlaceholderTypeExtension>,
    limits: Limits,
) -> Result<Vec<u8>> {
    match (location, target) {
        (SlotLocation::Existing { variant, .. }, Some(value)) => {
            replace_variant(source, variant, value, limits)
        },
        (SlotLocation::Existing { extension, .. }, None) => {
            replace_range(source, extension.start..extension.end, &[], limits)
        },
        (SlotLocation::Missing { .. }, None) => Err(Error::Invalid(
            "p232 placeholder slot removal is already a no-op".into(),
        )),
        (
            SlotLocation::Missing {
                placeholder,
                extension_list,
            },
            Some(value),
        ) => {
            let prefix = if extension_list
                .as_ref()
                .is_some_and(|value| !value.prefix.is_empty())
            {
                extension_list
                    .as_ref()
                    .map_or_else(Vec::new, |value| value.prefix.clone())
            } else {
                placeholder.prefix.clone()
            };
            let extension_len = extension_bytes_len(&prefix, value)?;
            match extension_list {
                Some(list) if list.empty => {
                    let list_len = list.end.saturating_sub(list.start);
                    let replacement_len = list_len
                        .checked_sub(1)
                        .and_then(|length| length.checked_add(extension_len))
                        .and_then(|length| {
                            length.checked_add(close_element_len(&prefix, b"extLst"))
                        })
                        .ok_or(Error::Limit {
                            resource: "p232 placeholder patch output bytes",
                            limit: limits.owner_bytes,
                        })?;
                    let target_len = source
                        .len()
                        .checked_sub(list_len)
                        .and_then(|length| length.checked_add(replacement_len))
                        .ok_or(Error::Limit {
                            resource: "p232 placeholder patch output bytes",
                            limit: limits.owner_bytes,
                        })?;
                    limits.check_growth(source.len(), target_len)?;
                    let extension = extension_bytes(&prefix, value)?;
                    let mut replacement =
                        bytes_with_capacity(replacement_len, "p232 placeholder extension bytes")?;
                    replacement.extend_from_slice(&source[list.start..list.end - 2]);
                    replacement.push(b'>');
                    replacement.extend_from_slice(&extension);
                    push_close(&mut replacement, &prefix, b"extLst");
                    replace_range(source, list.start..list.end, &replacement, limits)
                },
                Some(list) => {
                    limits.check_growth(
                        source.len(),
                        source
                            .len()
                            .checked_add(extension_len)
                            .ok_or(Error::Limit {
                                resource: "p232 placeholder patch output bytes",
                                limit: limits.owner_bytes,
                            })?,
                    )?;
                    let extension = extension_bytes(&prefix, value)?;
                    insert_range(source, list.close_start, &extension, limits)
                },
                None if placeholder.empty => {
                    let placeholder_len = placeholder.end.saturating_sub(placeholder.start);
                    let replacement_len = placeholder_len
                        .checked_sub(1)
                        .and_then(|length| {
                            length
                                .checked_add(open_element_len(&prefix, b"extLst"))
                                .and_then(|length| length.checked_add(extension_len))
                                .and_then(|length| {
                                    length.checked_add(close_element_len(&prefix, b"extLst"))
                                })
                                .and_then(|length| {
                                    length.checked_add(close_element_len(&prefix, b"ph"))
                                })
                        })
                        .ok_or(Error::Limit {
                            resource: "p232 placeholder patch output bytes",
                            limit: limits.owner_bytes,
                        })?;
                    let target_len = source
                        .len()
                        .checked_sub(placeholder_len)
                        .and_then(|length| length.checked_add(replacement_len))
                        .ok_or(Error::Limit {
                            resource: "p232 placeholder patch output bytes",
                            limit: limits.owner_bytes,
                        })?;
                    limits.check_growth(source.len(), target_len)?;
                    let extension = extension_bytes(&prefix, value)?;
                    let mut replacement =
                        bytes_with_capacity(replacement_len, "p232 placeholder extension bytes")?;
                    replacement.extend_from_slice(&source[placeholder.start..placeholder.end - 2]);
                    replacement.push(b'>');
                    push_ext_list(&mut replacement, &prefix, &extension);
                    push_close(&mut replacement, &prefix, b"ph");
                    replace_range(
                        source,
                        placeholder.start..placeholder.end,
                        &replacement,
                        limits,
                    )
                },
                None => {
                    let insertion_len = open_element_len(&prefix, b"extLst")
                        .checked_add(extension_len)
                        .and_then(|length| {
                            length.checked_add(close_element_len(&prefix, b"extLst"))
                        })
                        .ok_or(Error::Limit {
                            resource: "p232 placeholder patch output bytes",
                            limit: limits.owner_bytes,
                        })?;
                    limits.check_growth(
                        source.len(),
                        source
                            .len()
                            .checked_add(insertion_len)
                            .ok_or(Error::Limit {
                                resource: "p232 placeholder patch output bytes",
                                limit: limits.owner_bytes,
                            })?,
                    )?;
                    let extension = extension_bytes(&prefix, value)?;
                    let mut insertion =
                        bytes_with_capacity(insertion_len, "p232 placeholder extension bytes")?;
                    push_ext_list(&mut insertion, &prefix, &extension);
                    insert_range(source, placeholder.close_start, &insertion, limits)
                },
            }
        },
    }
}

fn replace_range(
    source: &[u8],
    range: Range<usize>,
    replacement: &[u8],
    limits: Limits,
) -> Result<Vec<u8>> {
    if range.start > range.end || range.end > source.len() {
        return Err(Error::Invalid(
            "p232 replacement range is outside its owner".into(),
        ));
    }
    let old_len = range
        .end
        .checked_sub(range.start)
        .ok_or_else(|| Error::Invalid("p232 source range is reversed".into()))?;
    let target_len = source
        .len()
        .checked_sub(old_len)
        .and_then(|value| value.checked_add(replacement.len()))
        .ok_or(Error::Limit {
            resource: "p232 placeholder patch output bytes",
            limit: limits.owner_bytes,
        })?;
    limits.check_growth(source.len(), target_len)?;
    let mut output = bytes_with_capacity(target_len, "p232 placeholder patch output bytes")?;
    output.extend_from_slice(&source[..range.start]);
    output.extend_from_slice(replacement);
    output.extend_from_slice(&source[range.end..]);
    Ok(output)
}

fn insert_range(source: &[u8], offset: usize, insertion: &[u8], limits: Limits) -> Result<Vec<u8>> {
    if offset > source.len() {
        return Err(Error::Invalid(
            "p232 insertion offset is outside its owner".into(),
        ));
    }
    let target_len = source
        .len()
        .checked_add(insertion.len())
        .ok_or(Error::Limit {
            resource: "p232 placeholder patch output bytes",
            limit: limits.owner_bytes,
        })?;
    limits.check_growth(source.len(), target_len)?;
    let mut output = bytes_with_capacity(target_len, "p232 placeholder patch output bytes")?;
    output.extend_from_slice(&source[..offset]);
    output.extend_from_slice(insertion);
    output.extend_from_slice(&source[offset..]);
    Ok(output)
}

fn open_element_len(prefix: &[u8], local: &[u8]) -> usize {
    1 + prefix.len() + if prefix.is_empty() { 0 } else { 1 } + local.len() + 1
}

fn close_element_len(prefix: &[u8], local: &[u8]) -> usize {
    2 + prefix.len() + if prefix.is_empty() { 0 } else { 1 } + local.len() + 1
}

fn extension_bytes_len(prefix: &[u8], value: PlaceholderTypeExtension) -> Result<usize> {
    let token = match value {
        PlaceholderTypeExtension::Cameo => b"<p232:cameo/>".as_slice(),
        PlaceholderTypeExtension::Unknown => b"<p232:unknown/>".as_slice(),
    };
    let parts = [
        1,
        prefix.len(),
        if prefix.is_empty() { 0 } else { 1 },
        b"ext uri=\"".len(),
        PLACEHOLDER_TYPE_EXTENSION_URI.len(),
        b"\"><p232:phTypeExt xmlns:p232=\"".len(),
        P232.len(),
        b"\"><p232:type>".len(),
        token.len(),
        b"</p232:type></p232:phTypeExt>".len(),
        close_element_len(prefix, b"ext"),
    ];
    parts.into_iter().try_fold(0usize, |length, part| {
        length.checked_add(part).ok_or(Error::Limit {
            resource: "p232 placeholder extension bytes",
            limit: MAX_OWNER_BYTES,
        })
    })
}

fn extension_bytes(prefix: &[u8], value: PlaceholderTypeExtension) -> Result<Vec<u8>> {
    let mut output = bytes_with_capacity(
        extension_bytes_len(prefix, value)?,
        "p232 placeholder extension bytes",
    )?;
    output.push(b'<');
    output.extend_from_slice(prefix);
    if !prefix.is_empty() {
        output.push(b':');
    }
    output.extend_from_slice(b"ext uri=\"");
    output.extend_from_slice(PLACEHOLDER_TYPE_EXTENSION_URI.as_bytes());
    output.extend_from_slice(b"\"><p232:phTypeExt xmlns:p232=\"");
    output.extend_from_slice(P232);
    output.extend_from_slice(b"\"><p232:type>");
    output.extend_from_slice(match value {
        PlaceholderTypeExtension::Cameo => b"<p232:cameo/>",
        PlaceholderTypeExtension::Unknown => b"<p232:unknown/>",
    });
    output.extend_from_slice(b"</p232:type></p232:phTypeExt>");
    push_close(&mut output, prefix, b"ext");
    Ok(output)
}

fn push_ext_list(output: &mut Vec<u8>, prefix: &[u8], extension: &[u8]) {
    push_open(output, prefix, b"extLst");
    output.extend_from_slice(extension);
    push_close(output, prefix, b"extLst");
}

fn push_open(output: &mut Vec<u8>, prefix: &[u8], local: &[u8]) {
    output.push(b'<');
    output.extend_from_slice(prefix);
    if !prefix.is_empty() {
        output.push(b':');
    }
    output.extend_from_slice(local);
    output.push(b'>');
}

fn push_close(output: &mut Vec<u8>, prefix: &[u8], local: &[u8]) {
    output.extend_from_slice(b"</");
    output.extend_from_slice(prefix);
    if !prefix.is_empty() {
        output.push(b':');
    }
    output.extend_from_slice(local);
    output.push(b'>');
}
