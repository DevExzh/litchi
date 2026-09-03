//! Exact-source transactions for Pages body drawable stacking order.
//!
//! The public surface is deliberately selector-first.  Callers receive
//! source-bound opaque handles from [`Package::body_drawable_order`], then
//! submit those handles to a bounded immutable transaction.  Native object
//! identifiers, component names, generated protobuf values, and retained
//! package bytes never cross this module's public API.

use std::collections::{HashMap, HashSet};
use std::fmt;
use std::num::NonZeroU64;
use std::sync::Arc;

use litchi_core::Position;
use litchi_iwa_archive::{SourceCatalog, package::EntryEdit};
use litchi_iwa_common::{
    WireLimits, decode_varint_from_bytes, varint::encoded_len, wire::WireView,
};
use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
use litchi_iwa_protos::pages_drawable_order_codec::{
    self as drawable_order_codec, DecodeError, DecodeOptions, DrawableOrderWrite, WireResourceLimit,
};
use thiserror::Error;

use super::Package;
use crate::drawable_order::{BodyDrawableHandle, BodyDrawableSelector, DrawableLayerMove};

const DOCUMENT_MEMBER: &str = "Index/Document.iwa";
const ROOT_OBJECT_IDENTIFIER: u64 = 1;
const ROOT_MESSAGE_TYPE: u32 = 10_000;
const ROOT_DRAWABLES_Z_ORDER_FIELD: u32 = 20;
const DRAWABLE_ORDER_MESSAGE_TYPE: u32 = 10_015;

const DRAWABLE_ORDER_RECURSION_LIMIT: u32 = 8;
const DRAWABLE_ORDER_FIELD_MULTIPLIER: usize = 8;
const DRAWABLE_ORDER_WORK_MULTIPLIER: usize = 64;

/// Finite resources enforced while reading or publishing one drawable-order
/// transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum BodyDrawableOrderLimitKind {
    /// Complete package input bytes.
    InputBytes,
    /// Complete edited package output bytes.
    OutputBytes,
    /// ZIP members retained by the package.
    Entries,
    /// Bytes retained by one ZIP member.
    EntryBytes,
    /// Aggregate bytes retained by ZIP members.
    TotalEntryBytes,
    /// Bytes in one decoded IWA component.
    PayloadBytes,
    /// Aggregate decoded IWA component bytes.
    TotalPayloadBytes,
    /// Native payload objects inspected by the transaction.
    PayloadObjects,
    /// Native payload messages inspected by the transaction.
    PayloadMessages,
    /// Native payload framing items inspected by the transaction.
    PayloadItems,
    /// Protobuf fields inspected by the strict order codec.
    WireFields,
    /// Protobuf bytes inspected by the strict order codec.
    WireBytes,
    /// Protobuf nesting used by the strict order codec.
    WireNesting,
    /// Aggregate strict codec work.
    WireWork,
    /// Repeated references inspected by the strict order codec.
    References,
}

impl fmt::Display for BodyDrawableOrderLimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::OutputBytes => "output bytes",
            Self::Entries => "ZIP entries",
            Self::EntryBytes => "ZIP entry bytes",
            Self::TotalEntryBytes => "total ZIP entry bytes",
            Self::PayloadBytes => "payload bytes",
            Self::TotalPayloadBytes => "total payload bytes",
            Self::PayloadObjects => "payload objects",
            Self::PayloadMessages => "payload messages",
            Self::PayloadItems => "payload items",
            Self::WireFields => "wire fields",
            Self::WireBytes => "wire bytes",
            Self::WireNesting => "wire nesting",
            Self::WireWork => "wire work",
            Self::References => "references",
        })
    }
}

/// Failure from a Pages body drawable-order read or immutable transaction.
#[derive(Debug, Error, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum BodyDrawableOrderError {
    /// The rooted Pages document has no drawable-order payload.
    #[error("the Pages document has no readable body drawable order")]
    MissingOrder,
    /// More than one rooted drawable-order payload was found.
    #[error("the Pages document has an ambiguous body drawable order")]
    AmbiguousOrder,
    /// The requested checked semantic position is outside the current order.
    #[error("the Pages body drawable position {position:?} does not exist")]
    PositionNotFound { position: Position },
    /// The supplied handle was issued by another exact source snapshot.
    #[error("the Pages body drawable handle belongs to another source snapshot")]
    HandleSourceMismatch,
    /// The supplied handle is not present in this source's current order.
    #[error("the Pages body drawable handle is not present in this source order")]
    HandleNotFound,
    /// The supplied handle sequence is not an exact permutation of the source
    /// order.
    #[error("the Pages body drawable order is not an exact existing-object permutation")]
    InvalidOrder,
    /// The source is a valid semantic snapshot but cannot publish an exact
    /// physical edit.
    #[error("this Pages source does not support physical drawable-order edits")]
    UnsupportedSource,
    /// The rooted physical graph or selected payload is not unambiguous.
    #[error("the Pages drawable-order source cannot be edited safely")]
    InvalidSource,
    /// A finite transaction resource ceiling was exceeded.
    #[error(
        "Pages body drawable order {kind} limit exceeded: observed {observed}, maximum {maximum}"
    )]
    LimitExceeded {
        /// Resource category that exceeded its ceiling.
        kind: BodyDrawableOrderLimitKind,
        /// Observed or requested resource amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded transaction allocation failed.
    #[error("could not allocate {amount} units for the Pages drawable-order transaction")]
    Allocation { amount: usize },
    /// Complete candidate reopening did not reproduce the requested order.
    #[error("the edited Pages body drawable order failed semantic verification")]
    Verification,
    /// The supplied patch does not belong to this exact package artifact.
    #[error("the Pages body drawable-order patch does not match the exact source package")]
    PatchConflict,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct OrderLocation {
    component_index: usize,
    object_index: usize,
    message_index: usize,
    object_identifier: NonZeroU64,
}

/// A mutable semantic body drawable-order edit staged against one immutable
/// package snapshot.
pub struct BodyDrawableOrderEdit<'a> {
    source: &'a Package,
    before: Vec<BodyDrawableHandle>,
    after: Vec<BodyDrawableHandle>,
}

impl fmt::Debug for BodyDrawableOrderEdit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("BodyDrawableOrderEdit")
            .field("before_positions", &positions(&self.before))
            .field("after_positions", &positions(&self.after))
            .finish_non_exhaustive()
    }
}

impl BodyDrawableOrderEdit<'_> {
    /// Borrow the source order captured when this edit began.
    #[must_use]
    pub fn before(&self) -> &[BodyDrawableHandle] {
        &self.before
    }

    /// Borrow the staged back-to-front order.
    #[must_use]
    pub fn order(&self) -> &[BodyDrawableHandle] {
        &self.after
    }

    /// Replace the staged order with an exact permutation of the handles
    /// issued by this source package.
    pub fn set_order(
        &mut self,
        requested: &[BodyDrawableHandle],
    ) -> Result<&mut Self, BodyDrawableOrderError> {
        let current = handles_to_identities(&self.before, &self.source_bytes())?;
        let desired = handles_to_identities(requested, &self.source_bytes())?;
        validate_exact_permutation(&current, &desired)?;
        let mut staged = allocate_vec(requested.len())?;
        staged.extend(requested.iter().cloned());
        self.after = staged;
        Ok(self)
    }

    /// Move one selected drawable using the native Arrange semantics.
    ///
    /// Returns `false` when the selected drawable is already at the requested
    /// boundary.  The source package is not changed until [`Self::commit`].
    pub fn move_drawable(
        &mut self,
        selector: impl Into<BodyDrawableSelector>,
        movement: DrawableLayerMove,
    ) -> Result<bool, BodyDrawableOrderError> {
        let index = resolve_staged_selector(&self.after, selector.into())?;
        let final_index = self
            .after
            .len()
            .checked_sub(1)
            .ok_or(BodyDrawableOrderError::InvalidSource)?;
        let target_index = match movement {
            DrawableLayerMove::ToBack => 0,
            DrawableLayerMove::Backward => index.saturating_sub(1),
            DrawableLayerMove::Forward => index.saturating_add(1).min(final_index),
            DrawableLayerMove::ToFront => final_index,
        };
        if target_index == index {
            return Ok(false);
        }
        let drawable = self.after.remove(index);
        self.after.insert(target_index, drawable);
        Ok(true)
    }

    /// Alias for [`Self::move_drawable`].
    pub fn move_body_drawable(
        &mut self,
        selector: impl Into<BodyDrawableSelector>,
        movement: DrawableLayerMove,
    ) -> Result<bool, BodyDrawableOrderError> {
        self.move_drawable(selector, movement)
    }

    /// Validate and atomically publish the staged immutable candidate.
    pub fn commit(self) -> Result<BodyDrawableOrderCommit, BodyDrawableOrderError> {
        commit_edit(self)
    }

    fn source_bytes(&self) -> Arc<[u8]> {
        self.source.state.source.shared_source()
    }
}

/// A reversible exact-source patch for one body drawable-order edit.
#[derive(Clone, PartialEq, Eq)]
pub struct BodyDrawableOrderPatch {
    source: Arc<[u8]>,
    target: Arc<[u8]>,
    source_fingerprint: u64,
    target_fingerprint: u64,
    proof: OrderLocation,
    before: Vec<BodyDrawableHandle>,
    after: Vec<BodyDrawableHandle>,
}

impl fmt::Debug for BodyDrawableOrderPatch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("BodyDrawableOrderPatch")
            .field("before_positions", &positions(&self.before))
            .field("after_positions", &positions(&self.after))
            .finish_non_exhaustive()
    }
}

impl BodyDrawableOrderPatch {
    /// Borrow the semantic order required before this patch can apply.
    #[must_use]
    pub fn before(&self) -> &[BodyDrawableHandle] {
        &self.before
    }

    /// Borrow the semantic order produced by this patch.
    #[must_use]
    pub fn after(&self) -> &[BodyDrawableHandle] {
        &self.after
    }

    /// Return whether this patch preserves both semantic state and exact
    /// source bytes.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        identities_equal(&self.before, &self.after)
            && self.source_fingerprint == self.target_fingerprint
            && (Arc::ptr_eq(&self.source, &self.target)
                || self.source.as_ref() == self.target.as_ref())
    }

    /// Return the exact inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            source: Arc::clone(&self.target),
            target: Arc::clone(&self.source),
            source_fingerprint: self.target_fingerprint,
            target_fingerprint: self.source_fingerprint,
            proof: self.proof,
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

    /// Fallibly construct the exact inverse when the patch contains a large
    /// order.  This is the allocation-aware counterpart to [`Self::inverse`]
    /// for callers that cannot tolerate an infallible vector clone.
    pub fn try_inverse(&self) -> Result<Self, BodyDrawableOrderError> {
        let mut before = allocate_vec(self.after.len())?;
        before.extend(self.after.iter().cloned());
        let mut after = allocate_vec(self.before.len())?;
        after.extend(self.before.iter().cloned());
        Ok(Self {
            source: Arc::clone(&self.target),
            target: Arc::clone(&self.source),
            source_fingerprint: self.target_fingerprint,
            target_fingerprint: self.source_fingerprint,
            proof: self.proof,
            before,
            after,
        })
    }
}

/// Compact evidence describing one body drawable-order commit.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct BodyDrawableOrderDiagnostics {
    changed: bool,
    touched_components: usize,
    full_reparse_performed: bool,
}

impl BodyDrawableOrderDiagnostics {
    const fn unchanged() -> Self {
        Self {
            changed: false,
            touched_components: 0,
            full_reparse_performed: false,
        }
    }

    const fn published() -> Self {
        Self {
            changed: true,
            touched_components: 1,
            full_reparse_performed: true,
        }
    }

    /// Return whether the committed package differs from its source.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// Return the number of rewritten IWA components.
    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }

    /// Return whether the complete candidate was reopened before publication.
    #[must_use]
    pub const fn full_reparse_performed(self) -> bool {
        self.full_reparse_performed
    }
}

/// Fully reopened immutable result of one body drawable-order transaction.
#[must_use = "a Pages body drawable-order commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct BodyDrawableOrderCommit {
    package: Package,
    patch: BodyDrawableOrderPatch,
    diagnostics: BodyDrawableOrderDiagnostics,
}

impl BodyDrawableOrderCommit {
    /// Borrow the fully reopened package snapshot.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Consume this commit and return its fully reopened package snapshot.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Borrow the reversible exact-source patch.
    #[must_use]
    pub const fn patch(&self) -> &BodyDrawableOrderPatch {
        &self.patch
    }

    /// Borrow compact publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &BodyDrawableOrderDiagnostics {
        &self.diagnostics
    }
}

impl Package {
    /// Read the rooted Pages body drawable order from back to front.
    ///
    /// Each returned handle is source-bound to this exact immutable package
    /// snapshot.  The handle's only public fact is its checked source
    /// position; native identity remains private to the package adapter.
    pub fn body_drawable_order(&self) -> Result<Vec<BodyDrawableHandle>, BodyDrawableOrderError> {
        let (_location, identifiers) = read_order(self)?;
        handles_for_order(self, &identifiers)
    }

    /// Select one body drawable by a checked position or source-bound handle.
    pub fn body_drawable(
        &self,
        selector: impl Into<BodyDrawableSelector>,
    ) -> Result<BodyDrawableHandle, BodyDrawableOrderError> {
        let handles = self.body_drawable_order()?;
        let index = resolve_staged_selector(&handles, selector.into())?;
        handles
            .get(index)
            .cloned()
            .ok_or(BodyDrawableOrderError::InvalidSource)
    }

    /// Start a selector-first immutable edit of the complete body drawable
    /// order.
    pub fn edit_body_drawable_order(
        &self,
    ) -> Result<BodyDrawableOrderEdit<'_>, BodyDrawableOrderError> {
        let before = self.body_drawable_order()?;
        let mut after = allocate_vec(before.len())?;
        after.extend(before.iter().cloned());
        Ok(BodyDrawableOrderEdit {
            source: self,
            before,
            after,
        })
    }

    /// Apply an exact-source-checked reversible body drawable-order patch.
    pub fn apply_body_drawable_order(
        &self,
        patch: &BodyDrawableOrderPatch,
    ) -> Result<BodyDrawableOrderCommit, BodyDrawableOrderError> {
        let source = self.state.source.shared_source();
        if patch.source_fingerprint != fingerprint(source.as_ref())
            || source.as_ref() != patch.source.as_ref()
            || !handles_belong_to_source(&patch.before, &source)
        {
            return Err(BodyDrawableOrderError::PatchConflict);
        }
        let (location, current) = read_order(self)?;
        let before = handles_to_identities(&patch.before, &source)?;
        if current != before || location != patch.proof {
            return Err(BodyDrawableOrderError::PatchConflict);
        }
        if patch.is_noop() {
            return Ok(BodyDrawableOrderCommit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: BodyDrawableOrderDiagnostics::unchanged(),
            });
        }
        if !self.state.source.source_is_exact()
            || fingerprint(patch.target.as_ref()) != patch.target_fingerprint
        {
            return Err(BodyDrawableOrderError::PatchConflict);
        }
        let candidate_source = SourceCatalog::from_shared_bytes_with_limits(
            Arc::clone(&patch.target),
            self.state.source.limits(),
        )
        .map_err(map_archive_error)?;
        let candidate =
            Package::from_source_catalog(candidate_source).map_err(map_package_error)?;
        let (candidate_location, candidate_order) = read_order(&candidate)?;
        let target_source = candidate.state.source.shared_source();
        let after = handles_to_identities(&patch.after, &target_source)?;
        if candidate_location != patch.proof || candidate_order != after {
            return Err(BodyDrawableOrderError::Verification);
        }
        verify_locality(self, &candidate, patch.proof)?;
        Ok(BodyDrawableOrderCommit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: BodyDrawableOrderDiagnostics::published(),
        })
    }
}

fn commit_edit(
    edit: BodyDrawableOrderEdit<'_>,
) -> Result<BodyDrawableOrderCommit, BodyDrawableOrderError> {
    let source = edit.source.state.source.shared_source();
    let (location, current) = read_order(edit.source)?;
    let before = handles_to_identities(&edit.before, &source)?;
    let after = handles_to_identities(&edit.after, &source)?;
    if current != before {
        return Err(BodyDrawableOrderError::InvalidSource);
    }
    validate_exact_permutation(&before, &after)?;
    let source_fingerprint = fingerprint(source.as_ref());
    if before == after {
        return Ok(BodyDrawableOrderCommit {
            package: edit.source.snapshot(),
            patch: BodyDrawableOrderPatch {
                source: Arc::clone(&source),
                target: source,
                source_fingerprint,
                target_fingerprint: source_fingerprint,
                proof: location,
                before: edit.before,
                after: edit.after,
            },
            diagnostics: BodyDrawableOrderDiagnostics::unchanged(),
        });
    }
    if !edit.source.state.source.source_is_exact() {
        return Err(BodyDrawableOrderError::UnsupportedSource);
    }

    let candidate = rewrite_order(edit.source, location, &current, &after)?;
    let (candidate_location, candidate_order) = read_order(&candidate)?;
    if candidate_location != location || candidate_order != after {
        return Err(BodyDrawableOrderError::Verification);
    }
    verify_locality(edit.source, &candidate, location)?;
    let target = candidate.state.source.shared_source();
    let target_handles = handles_for_order(&candidate, &candidate_order)?;
    Ok(BodyDrawableOrderCommit {
        package: candidate,
        patch: BodyDrawableOrderPatch {
            source,
            target: Arc::clone(&target),
            source_fingerprint,
            target_fingerprint: fingerprint(target.as_ref()),
            proof: location,
            before: edit.before,
            after: target_handles,
        },
        diagnostics: BodyDrawableOrderDiagnostics::published(),
    })
}

fn read_order(
    package: &Package,
) -> Result<(OrderLocation, Vec<NonZeroU64>), BodyDrawableOrderError> {
    let source = &package.state.source;
    let components = source.components();
    let locations = object_locations(components)?;
    let component_index = components
        .iter()
        .position(|component| component.name() == DOCUMENT_MEMBER)
        .ok_or(BodyDrawableOrderError::InvalidSource)?;
    let component = components
        .get_index(component_index)
        .ok_or(BodyDrawableOrderError::InvalidSource)?;
    let root_object_index = component
        .archive()
        .objects
        .iter()
        .position(|object| object.archive_info.identifier == Some(ROOT_OBJECT_IDENTIFIER))
        .ok_or(BodyDrawableOrderError::InvalidSource)?;
    let root = component
        .archive()
        .objects
        .get(root_object_index)
        .ok_or(BodyDrawableOrderError::InvalidSource)?;
    let root_message_index = unique_message_index(root, ROOT_MESSAGE_TYPE)?;
    let root_message = root
        .messages
        .get(root_message_index)
        .ok_or(BodyDrawableOrderError::InvalidSource)?;
    validate_message_metadata(root, root_message_index)?;
    let root_view =
        WireView::parse_with_limits(root_message.data.as_slice(), wire_limits(package)?)
            .map_err(|_| BodyDrawableOrderError::InvalidSource)?;
    let mut zorder_identifier = None;
    let mut body_storage_identifier = None;
    for field in root_view.fields() {
        field
            .validate_canonical_framing()
            .map_err(|_| BodyDrawableOrderError::InvalidSource)?;
        match field.number() {
            ROOT_DRAWABLES_Z_ORDER_FIELD => {
                if zorder_identifier.is_some() || field.wire_type() != 2 {
                    return Err(BodyDrawableOrderError::AmbiguousOrder);
                }
                zorder_identifier = Some(strict_reference(field.payload(), package)?);
            },
            // A body-storage object is a container, not a body drawable.  It
            // is retained here only so a malformed z-order cannot expose it
            // as an opaque drawable handle.
            4 => {
                if body_storage_identifier.is_some() || field.wire_type() != 2 {
                    return Err(BodyDrawableOrderError::InvalidSource);
                }
                body_storage_identifier = Some(strict_reference(field.payload(), package)?);
            },
            _ => {},
        }
    }
    let zorder_identifier = zorder_identifier.ok_or(BodyDrawableOrderError::MissingOrder)?;
    let (zorder_component_index, zorder_object_index) = locations
        .get(&zorder_identifier.get())
        .copied()
        .ok_or(BodyDrawableOrderError::InvalidSource)?;
    let zorder_component = components
        .get_index(zorder_component_index)
        .ok_or(BodyDrawableOrderError::InvalidSource)?;
    let zorder_object = zorder_component
        .archive()
        .objects
        .get(zorder_object_index)
        .ok_or(BodyDrawableOrderError::InvalidSource)?;
    let message_index = unique_message_index(zorder_object, DRAWABLE_ORDER_MESSAGE_TYPE)?;
    validate_message_metadata(zorder_object, message_index)?;
    let message = zorder_object
        .messages
        .get(message_index)
        .ok_or(BodyDrawableOrderError::InvalidSource)?;
    let options = codec_options(package, message.data.as_slice(), false)?;
    let (snapshot, _report) =
        drawable_order_codec::decode_drawable_order_with_report(&message.data, options)
            .map_err(map_codec_error)?;
    reject_external_references(snapshot)?;
    let mut identifiers = allocate_vec(snapshot.len())?;
    let mut seen = HashSet::new();
    seen.try_reserve(snapshot.len())
        .map_err(|_| BodyDrawableOrderError::Allocation {
            amount: snapshot.len(),
        })?;
    for identifier in snapshot.identifiers() {
        let identifier =
            NonZeroU64::new(identifier).ok_or(BodyDrawableOrderError::InvalidSource)?;
        if !locations.contains_key(&identifier.get()) {
            return Err(BodyDrawableOrderError::InvalidSource);
        }
        if identifier.get() == ROOT_OBJECT_IDENTIFIER
            || identifier == zorder_identifier
            || body_storage_identifier == Some(identifier)
        {
            return Err(BodyDrawableOrderError::InvalidSource);
        }
        if !seen.insert(identifier) {
            return Err(BodyDrawableOrderError::InvalidSource);
        }
        identifiers.push(identifier);
    }
    let location = OrderLocation {
        component_index: zorder_component_index,
        object_index: zorder_object_index,
        message_index,
        object_identifier: zorder_identifier,
    };
    Ok((location, identifiers))
}

fn handles_for_order(
    package: &Package,
    identifiers: &[NonZeroU64],
) -> Result<Vec<BodyDrawableHandle>, BodyDrawableOrderError> {
    let source = package.state.source.shared_source();
    let mut handles = allocate_vec(identifiers.len())?;
    for (index, &identifier) in identifiers.iter().enumerate() {
        handles.push(BodyDrawableHandle::new(
            Arc::clone(&source),
            identifier,
            Position::new(index),
        ));
    }
    Ok(handles)
}

fn handles_to_identities(
    handles: &[BodyDrawableHandle],
    source: &Arc<[u8]>,
) -> Result<Vec<NonZeroU64>, BodyDrawableOrderError> {
    let mut identities = allocate_vec(handles.len())?;
    for handle in handles {
        if !source_matches(&handle.source, source) {
            return Err(BodyDrawableOrderError::HandleSourceMismatch);
        }
        identities.push(handle.identity);
    }
    Ok(identities)
}

fn allocate_vec<T>(amount: usize) -> Result<Vec<T>, BodyDrawableOrderError> {
    let mut values = Vec::new();
    values
        .try_reserve_exact(amount)
        .map_err(|_| BodyDrawableOrderError::Allocation { amount })?;
    Ok(values)
}

fn handles_belong_to_source(handles: &[BodyDrawableHandle], source: &Arc<[u8]>) -> bool {
    handles
        .iter()
        .all(|handle| source_matches(&handle.source, source))
}

fn source_matches(left: &Arc<[u8]>, right: &Arc<[u8]>) -> bool {
    Arc::ptr_eq(left, right) || left.as_ref() == right.as_ref()
}

fn resolve_staged_selector(
    handles: &[BodyDrawableHandle],
    selector: BodyDrawableSelector,
) -> Result<usize, BodyDrawableOrderError> {
    match selector {
        BodyDrawableSelector::Position(position) => handles
            .get(position.get())
            .map(|_| position.get())
            .ok_or(BodyDrawableOrderError::PositionNotFound { position }),
        BodyDrawableSelector::Handle(handle) => handles
            .iter()
            .position(|candidate| candidate == &handle)
            .ok_or_else(|| {
                if handles
                    .first()
                    .is_some_and(|candidate| !source_matches(&candidate.source, &handle.source))
                {
                    BodyDrawableOrderError::HandleSourceMismatch
                } else {
                    BodyDrawableOrderError::HandleNotFound
                }
            }),
    }
}

fn identities_equal(left: &[BodyDrawableHandle], right: &[BodyDrawableHandle]) -> bool {
    left.len() == right.len()
        && left
            .iter()
            .zip(right)
            .all(|(left, right)| left.identity == right.identity)
}

struct Positions<'a>(&'a [BodyDrawableHandle]);

impl fmt::Debug for Positions<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_list()
            .entries(self.0.iter().map(BodyDrawableHandle::position))
            .finish()
    }
}

fn positions(handles: &[BodyDrawableHandle]) -> Positions<'_> {
    Positions(handles)
}

fn validate_exact_permutation(
    current: &[NonZeroU64],
    requested: &[NonZeroU64],
) -> Result<(), BodyDrawableOrderError> {
    if current.len() != requested.len() {
        return Err(BodyDrawableOrderError::InvalidOrder);
    }
    // An empty order is a valid Pages state.  It is also the only valid
    // empty permutation, so return before constructing validation scratch.
    if current.is_empty() {
        return Ok(());
    }
    let mut current_set = HashSet::new();
    current_set
        .try_reserve(current.len())
        .map_err(|_| BodyDrawableOrderError::Allocation {
            amount: current.len(),
        })?;
    if current
        .iter()
        .any(|identifier| !current_set.insert(*identifier))
    {
        return Err(BodyDrawableOrderError::InvalidOrder);
    }
    let mut requested_set = HashSet::new();
    requested_set
        .try_reserve(requested.len())
        .map_err(|_| BodyDrawableOrderError::Allocation {
            amount: requested.len(),
        })?;
    if requested
        .iter()
        .any(|identifier| !requested_set.insert(*identifier))
        || requested_set != current_set
    {
        return Err(BodyDrawableOrderError::InvalidOrder);
    }
    Ok(())
}

fn rewrite_order(
    source: &Package,
    location: OrderLocation,
    current: &[NonZeroU64],
    requested: &[NonZeroU64],
) -> Result<Package, BodyDrawableOrderError> {
    let catalog = &source.state.source;
    let component = catalog
        .components()
        .get_index(location.component_index)
        .ok_or(BodyDrawableOrderError::InvalidSource)?;
    let component_name = component.name().to_owned();
    let entry = catalog
        .package()
        .iter()
        .find(|entry| entry.name() == component_name)
        .ok_or(BodyDrawableOrderError::InvalidSource)?;
    if entry.is_opaque() {
        return Err(BodyDrawableOrderError::UnsupportedSource);
    }
    let archive_limits = catalog
        .limits()
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let stream = SnappyStream::decompress_with_limits(
        entry.data(),
        catalog
            .limits()
            .snappy_limits()
            .map_err(map_archive_error)?,
    )
    .map_err(map_core_error)?;
    let mut archive =
        Archive::parse_with_limits(stream.as_bytes(), archive_limits).map_err(map_core_error)?;
    validate_canonical_object_length_prefixes(stream.as_bytes(), &archive)?;
    let object = archive
        .objects
        .get(location.object_index)
        .ok_or(BodyDrawableOrderError::InvalidSource)?;
    if object.archive_info.identifier != Some(location.object_identifier.get()) {
        return Err(BodyDrawableOrderError::InvalidSource);
    }
    let message = object
        .messages
        .get(location.message_index)
        .filter(|message| message.type_ == DRAWABLE_ORDER_MESSAGE_TYPE)
        .ok_or(BodyDrawableOrderError::InvalidSource)?;
    let options = codec_options(source, &message.data, true)?;
    let (decoded, _report) =
        drawable_order_codec::decode_drawable_order_with_report(&message.data, options)
            .map_err(map_codec_error)?;
    reject_external_references(decoded)?;
    let mut decoded_identifiers = allocate_vec(decoded.len())?;
    for identifier in decoded.identifiers() {
        decoded_identifiers
            .push(NonZeroU64::new(identifier).ok_or(BodyDrawableOrderError::InvalidSource)?);
    }
    if decoded_identifiers != current {
        return Err(BodyDrawableOrderError::InvalidSource);
    }
    let mut raw_requested = allocate_vec(requested.len())?;
    raw_requested.extend(requested.iter().map(|identifier| identifier.get()));
    let (rewritten, _report) = drawable_order_codec::rewrite_drawable_order_with_report(
        &message.data,
        DrawableOrderWrite::new(&raw_requested),
        options,
    )
    .map_err(map_codec_error)?;
    let (verified, _report) = drawable_order_codec::decode_drawable_order_with_report(
        &rewritten,
        codec_options(source, &rewritten, false)?,
    )
    .map_err(map_codec_error)?;
    reject_external_references(verified)?;
    if verified.identifiers().ne(raw_requested.iter().copied()) {
        return Err(BodyDrawableOrderError::Verification);
    }
    archive
        .objects
        .get_mut(location.object_index)
        .ok_or(BodyDrawableOrderError::InvalidSource)?
        .replace_message_preserving_header_with_limits(
            location.message_index,
            RawMessage {
                type_: DRAWABLE_ORDER_MESSAGE_TYPE,
                data: rewritten,
            },
            archive_limits,
        )
        .map_err(map_core_error)?;
    let archive_bytes = archive
        .to_bytes_with_limits(archive_limits)
        .map_err(map_core_error)?;
    let compressed = SnappyStream::compress(&archive_bytes).map_err(map_core_error)?;
    let output = catalog
        .package()
        .reassemble_to_bytes(
            &[EntryEdit::new(
                component_name.as_str(),
                compressed.as_slice(),
            )],
            catalog.limits(),
        )
        .map_err(map_archive_error)?;
    let candidate_source =
        SourceCatalog::from_shared_bytes_with_limits(output.into(), catalog.limits())
            .map_err(map_archive_error)?;
    Package::from_source_catalog(candidate_source).map_err(map_package_error)
}

fn verify_locality(
    source: &Package,
    candidate: &Package,
    location: OrderLocation,
) -> Result<(), BodyDrawableOrderError> {
    let left = source.state.source.package();
    let right = candidate.state.source.package();
    if left.len() != right.len()
        || source.state.source.components().len() != candidate.state.source.components().len()
    {
        return Err(BodyDrawableOrderError::Verification);
    }
    for (index, (before, after)) in left.iter().zip(right.iter()).enumerate() {
        if before.name() != after.name() || before.raw_name() != after.raw_name() {
            return Err(BodyDrawableOrderError::Verification);
        }
        if index != location.component_index {
            if before.data() != after.data()
                || before.raw_record().local_record() != after.raw_record().local_record()
                || before.raw_record().central_directory_record()
                    != after.raw_record().central_directory_record()
            {
                return Err(BodyDrawableOrderError::Verification);
            }
        }
    }
    let before_component = source
        .state
        .source
        .components()
        .get_index(location.component_index)
        .ok_or(BodyDrawableOrderError::Verification)?;
    let after_component = candidate
        .state
        .source
        .components()
        .get_index(location.component_index)
        .ok_or(BodyDrawableOrderError::Verification)?;
    if before_component.name() != after_component.name()
        || before_component.archive().objects.len() != after_component.archive().objects.len()
    {
        return Err(BodyDrawableOrderError::Verification);
    }
    for (object_index, (before, after)) in before_component
        .archive()
        .objects
        .iter()
        .zip(after_component.archive().objects.iter())
        .enumerate()
    {
        if object_index != location.object_index {
            if !before.same_content_ignoring_offsets(after) {
                return Err(BodyDrawableOrderError::Verification);
            }
            continue;
        }
        if before.archive_info != after.archive_info
            || before.header_length != after.header_length
            || before.data_length != after.data_length
            || before.messages.len() != after.messages.len()
        {
            return Err(BodyDrawableOrderError::Verification);
        }
        for (message_index, (before, after)) in
            before.messages.iter().zip(&after.messages).enumerate()
        {
            if message_index != location.message_index && before != after {
                return Err(BodyDrawableOrderError::Verification);
            }
            if message_index == location.message_index && before.type_ != after.type_ {
                return Err(BodyDrawableOrderError::Verification);
            }
        }
    }
    Ok(())
}

fn object_locations(
    components: &litchi_iwa_archive::ComponentCatalog,
) -> Result<HashMap<u64, (usize, usize)>, BodyDrawableOrderError> {
    // Build one bounded lookup for the complete component catalog.  The old
    // per-reference scan made a large order repeatedly walk every object and
    // also silently selected the first copy of a duplicated native ID.
    let object_count = components
        .iter()
        .try_fold(0usize, |count, component| {
            count.checked_add(component.archive().objects.len())
        })
        .ok_or(BodyDrawableOrderError::InvalidSource)?;
    let mut locations = HashMap::new();
    locations
        .try_reserve(object_count)
        .map_err(|_| BodyDrawableOrderError::Allocation {
            amount: object_count,
        })?;
    for (component_index, component) in components.iter().enumerate() {
        for (object_index, object) in component.archive().objects.iter().enumerate() {
            let Some(identifier) = object.archive_info.identifier else {
                continue;
            };
            if identifier == 0
                || locations
                    .insert(identifier, (component_index, object_index))
                    .is_some()
            {
                return Err(BodyDrawableOrderError::InvalidSource);
            }
        }
    }
    Ok(locations)
}

fn unique_message_index(
    object: &ArchiveObject,
    message_type: u32,
) -> Result<usize, BodyDrawableOrderError> {
    let mut indexes = object
        .messages
        .iter()
        .enumerate()
        .filter(|(_, message)| message.type_ == message_type)
        .map(|(index, _)| index);
    let index = indexes.next().ok_or(BodyDrawableOrderError::MissingOrder)?;
    if indexes.next().is_some() {
        return Err(BodyDrawableOrderError::AmbiguousOrder);
    }
    Ok(index)
}

fn validate_message_metadata(
    object: &ArchiveObject,
    message_index: usize,
) -> Result<(), BodyDrawableOrderError> {
    let message = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(BodyDrawableOrderError::InvalidSource)?;
    if object.archive_info.should_merge == Some(true)
        || message.base_message_index.is_some()
        || !message.diff_merge_version.is_empty()
        || message.diff_field_path.is_some()
        || !message.fields_to_remove.is_empty()
        || !message.diff_read_version.is_empty()
    {
        return Err(BodyDrawableOrderError::InvalidSource);
    }
    Ok(())
}

fn strict_reference(
    source: &[u8],
    package: &Package,
) -> Result<NonZeroU64, BodyDrawableOrderError> {
    strict_reference_with_limits(source, wire_limits(package)?)
}

fn strict_reference_with_limits(
    source: &[u8],
    limits: WireLimits,
) -> Result<NonZeroU64, BodyDrawableOrderError> {
    // TSP.Reference is proto2: known optional fields have presence, so a
    // duplicate is ambiguous even when both values happen to agree.  Keep
    // this parser independent from the generated value path and reject the
    // external-reference marker before it can become a local object handle.
    let view = WireView::parse_with_limits(source, limits)
        .map_err(|_| BodyDrawableOrderError::InvalidSource)?;
    let mut identifier = None;
    let mut deprecated_type_seen = false;
    let mut deprecated_external_seen = false;
    for field in view.fields() {
        field
            .validate_canonical_framing()
            .map_err(|_| BodyDrawableOrderError::InvalidSource)?;
        match field.number() {
            1 => {
                if identifier.is_some() || field.wire_type() != 0 {
                    return Err(BodyDrawableOrderError::InvalidSource);
                }
                let (value, consumed) = decode_varint_from_bytes(field.payload())
                    .map_err(|_| BodyDrawableOrderError::InvalidSource)?;
                if consumed != field.payload().len() || encoded_len(value) != consumed {
                    return Err(BodyDrawableOrderError::InvalidSource);
                }
                identifier = NonZeroU64::new(value);
            },
            2 => {
                if deprecated_type_seen || field.wire_type() != 0 {
                    return Err(BodyDrawableOrderError::InvalidSource);
                }
                deprecated_type_seen = true;
                let (value, consumed) = decode_varint_from_bytes(field.payload())
                    .map_err(|_| BodyDrawableOrderError::InvalidSource)?;
                if consumed != field.payload().len()
                    || encoded_len(value) != consumed
                    || !is_canonical_int32(value)
                {
                    return Err(BodyDrawableOrderError::InvalidSource);
                }
            },
            3 => {
                if deprecated_external_seen || field.wire_type() != 0 {
                    return Err(BodyDrawableOrderError::InvalidSource);
                }
                deprecated_external_seen = true;
                let (value, consumed) = decode_varint_from_bytes(field.payload())
                    .map_err(|_| BodyDrawableOrderError::InvalidSource)?;
                if consumed != field.payload().len()
                    || encoded_len(value) != consumed
                    || value > 1
                    || value == 1
                {
                    return Err(BodyDrawableOrderError::InvalidSource);
                }
            },
            _ => {},
        }
    }
    identifier.ok_or(BodyDrawableOrderError::InvalidSource)
}

fn is_canonical_int32(value: u64) -> bool {
    value <= 0x7fff_ffff || value >= 0xffff_ffff_8000_0000
}

fn reject_external_references(
    snapshot: drawable_order_codec::DrawableOrderSnapshot<'_>,
) -> Result<(), BodyDrawableOrderError> {
    if snapshot
        .references()
        .any(|reference| reference.deprecated_is_external() == Some(true))
    {
        return Err(BodyDrawableOrderError::InvalidSource);
    }
    Ok(())
}

fn wire_limits(package: &Package) -> Result<WireLimits, BodyDrawableOrderError> {
    let maximum = package
        .state
        .source
        .limits()
        .effective_archive_limits()
        .map_err(map_archive_error)?
        .max_message_bytes();
    let defaults = WireLimits::default();
    let maximum = maximum.min(defaults.max_input_bytes());
    WireLimits::default()
        .with_input_bytes(maximum)
        .and_then(|limits| limits.with_output_bytes(maximum.min(defaults.max_output_bytes())))
        .map_err(|_| BodyDrawableOrderError::InvalidSource)
}

fn codec_options(
    package: &Package,
    source: &[u8],
    rewrite: bool,
) -> Result<DecodeOptions, BodyDrawableOrderError> {
    let limits = wire_limits(package)?;
    let max_fields = source
        .len()
        .saturating_mul(DRAWABLE_ORDER_FIELD_MULTIPLIER)
        .clamp(1, limits.max_fields());
    let max_work = source
        .len()
        .saturating_mul(DRAWABLE_ORDER_WORK_MULTIPLIER)
        .clamp(1, limits.max_rewrite_work());
    let max_references = source.len().clamp(1, limits.max_fields());
    let recursion_limit = u32::try_from(limits.max_nesting().min(64))
        .map_err(|_| BodyDrawableOrderError::InvalidSource)?;
    Ok(DecodeOptions::new(
        source.len().max(1).min(limits.max_input_bytes()),
        if rewrite {
            limits.max_output_bytes()
        } else {
            source.len().max(1).min(limits.max_output_bytes())
        },
        max_fields,
        max_work,
        DRAWABLE_ORDER_RECURSION_LIMIT.min(recursion_limit),
        max_references,
    ))
}

fn validate_canonical_object_length_prefixes(
    source: &[u8],
    archive: &Archive,
) -> Result<(), BodyDrawableOrderError> {
    for object in &archive.objects {
        let offset = usize::try_from(object.header_offset)
            .map_err(|_| BodyDrawableOrderError::InvalidSource)?;
        let remaining = source
            .get(offset..)
            .ok_or(BodyDrawableOrderError::InvalidSource)?;
        let (header_bytes, prefix_bytes) = decode_varint_from_bytes(remaining)
            .map_err(|_| BodyDrawableOrderError::InvalidSource)?;
        if prefix_bytes != encoded_len(header_bytes) {
            return Err(BodyDrawableOrderError::InvalidSource);
        }
        let framed_header_bytes = header_bytes
            .checked_add(
                u64::try_from(prefix_bytes).map_err(|_| BodyDrawableOrderError::InvalidSource)?,
            )
            .ok_or(BodyDrawableOrderError::InvalidSource)?;
        if framed_header_bytes != object.header_length
            || object
                .header_offset
                .checked_add(object.header_length)
                .ok_or(BodyDrawableOrderError::InvalidSource)?
                != object.data_offset
        {
            return Err(BodyDrawableOrderError::InvalidSource);
        }
    }
    Ok(())
}

fn fingerprint(bytes: &[u8]) -> u64 {
    let mut value = 0xcbf2_9ce4_8422_2325_u64;
    for byte in bytes {
        value ^= u64::from(*byte);
        value = value.wrapping_mul(0x0000_0100_0000_01b3);
    }
    value
}

fn map_codec_error(error: DecodeError) -> BodyDrawableOrderError {
    if let Some(amount) = error.allocation_amount() {
        return BodyDrawableOrderError::Allocation { amount };
    }
    match error.resource_limit() {
        Some(WireResourceLimit::InputBytes { observed, maximum }) => {
            BodyDrawableOrderError::LimitExceeded {
                kind: BodyDrawableOrderLimitKind::WireBytes,
                observed: observed as u64,
                maximum: maximum as u64,
            }
        },
        Some(WireResourceLimit::OutputBytes { observed, maximum }) => {
            BodyDrawableOrderError::LimitExceeded {
                kind: BodyDrawableOrderLimitKind::OutputBytes,
                observed: observed as u64,
                maximum: maximum as u64,
            }
        },
        Some(WireResourceLimit::Fields { observed, maximum }) => {
            BodyDrawableOrderError::LimitExceeded {
                kind: BodyDrawableOrderLimitKind::WireFields,
                observed: observed as u64,
                maximum: maximum as u64,
            }
        },
        Some(WireResourceLimit::WorkBytes { observed, maximum }) => {
            BodyDrawableOrderError::LimitExceeded {
                kind: BodyDrawableOrderLimitKind::WireWork,
                observed: observed as u64,
                maximum: maximum as u64,
            }
        },
        Some(WireResourceLimit::Nesting { observed, maximum }) => {
            BodyDrawableOrderError::LimitExceeded {
                kind: BodyDrawableOrderLimitKind::WireNesting,
                observed: u64::from(observed),
                maximum: u64::from(maximum),
            }
        },
        Some(WireResourceLimit::References { observed, maximum }) => {
            BodyDrawableOrderError::LimitExceeded {
                kind: BodyDrawableOrderLimitKind::References,
                observed: observed as u64,
                maximum: maximum as u64,
            }
        },
        Some(_) => BodyDrawableOrderError::InvalidSource,
        None => BodyDrawableOrderError::InvalidSource,
    }
}

fn map_package_error(error: super::PackageError) -> BodyDrawableOrderError {
    match error {
        super::PackageError::Archive(error) => map_archive_error(error),
        super::PackageError::Detection(_) => BodyDrawableOrderError::InvalidSource,
        super::PackageError::Allocation { amount } => BodyDrawableOrderError::Allocation { amount },
        super::PackageError::ObjectLimit { observed, limit } => {
            BodyDrawableOrderError::LimitExceeded {
                kind: BodyDrawableOrderLimitKind::PayloadObjects,
                observed: observed as u64,
                maximum: limit as u64,
            }
        },
        super::PackageError::PayloadLimit { observed, limit } => {
            BodyDrawableOrderError::LimitExceeded {
                kind: BodyDrawableOrderLimitKind::PayloadBytes,
                observed: observed as u64,
                maximum: limit as u64,
            }
        },
        super::PackageError::NotPages
        | super::PackageError::Io(_)
        | super::PackageError::InvalidFormat(_)
        | super::PackageError::Semantic(_)
        | super::PackageError::SectionNamesTooLarge { .. } => BodyDrawableOrderError::InvalidSource,
    }
}

fn map_archive_error(error: litchi_iwa_archive::Error) -> BodyDrawableOrderError {
    match error {
        litchi_iwa_archive::Error::Limit {
            kind,
            observed,
            maximum,
        } => BodyDrawableOrderError::LimitExceeded {
            kind: match kind {
                litchi_iwa_archive::LimitKind::InputBytes => BodyDrawableOrderLimitKind::InputBytes,
                litchi_iwa_archive::LimitKind::OutputBytes => {
                    BodyDrawableOrderLimitKind::OutputBytes
                },
                litchi_iwa_archive::LimitKind::Entries => BodyDrawableOrderLimitKind::Entries,
                litchi_iwa_archive::LimitKind::MemberNameBytes
                | litchi_iwa_archive::LimitKind::MetadataBytes => {
                    BodyDrawableOrderLimitKind::PayloadItems
                },
                litchi_iwa_archive::LimitKind::CompressedEntryBytes
                | litchi_iwa_archive::LimitKind::EntryBytes => {
                    BodyDrawableOrderLimitKind::EntryBytes
                },
                litchi_iwa_archive::LimitKind::TotalBytes => {
                    BodyDrawableOrderLimitKind::TotalEntryBytes
                },
                litchi_iwa_archive::LimitKind::IwaStreamBytes => {
                    BodyDrawableOrderLimitKind::PayloadBytes
                },
                litchi_iwa_archive::LimitKind::IwaTotalBytes => {
                    BodyDrawableOrderLimitKind::TotalPayloadBytes
                },
            },
            observed,
            maximum,
        },
        litchi_iwa_archive::Error::Allocation { amount, .. } => {
            BodyDrawableOrderError::Allocation { amount }
        },
        litchi_iwa_archive::Error::Iwa(error) => map_core_error(error),
        litchi_iwa_archive::Error::Reassembly(_) => BodyDrawableOrderError::UnsupportedSource,
        litchi_iwa_archive::Error::Io(_)
        | litchi_iwa_archive::Error::Zip { .. }
        | litchi_iwa_archive::Error::InvalidLimits(_)
        | litchi_iwa_archive::Error::Encrypted
        | litchi_iwa_archive::Error::SourceChanged { .. }
        | litchi_iwa_archive::Error::DirectoryChanged { .. }
        | litchi_iwa_archive::Error::InvalidBundle(_) => BodyDrawableOrderError::InvalidSource,
    }
}

fn map_core_error(error: litchi_iwa_core::Error) -> BodyDrawableOrderError {
    match error {
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => BodyDrawableOrderError::LimitExceeded {
            kind: match kind {
                litchi_iwa_core::LimitKind::Objects => BodyDrawableOrderLimitKind::PayloadObjects,
                litchi_iwa_core::LimitKind::Messages
                | litchi_iwa_core::LimitKind::MessagesPerObject => {
                    BodyDrawableOrderLimitKind::PayloadMessages
                },
                litchi_iwa_core::LimitKind::HeaderFields
                | litchi_iwa_core::LimitKind::MetadataItems
                | litchi_iwa_core::LimitKind::SnappyFrames => {
                    BodyDrawableOrderLimitKind::PayloadItems
                },
                litchi_iwa_core::LimitKind::HeaderNesting => {
                    BodyDrawableOrderLimitKind::WireNesting
                },
                litchi_iwa_core::LimitKind::ArchiveBytes
                | litchi_iwa_core::LimitKind::ObjectBytes
                | litchi_iwa_core::LimitKind::MessageBytes
                | litchi_iwa_core::LimitKind::HeaderBytes
                | litchi_iwa_core::LimitKind::HeaderMemoryBytes
                | litchi_iwa_core::LimitKind::SnappyChunkBytes
                | litchi_iwa_core::LimitKind::SnappyStreamBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedChunkBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedStreamBytes => {
                    BodyDrawableOrderLimitKind::PayloadBytes
                },
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            BodyDrawableOrderError::Allocation { amount: requested }
        },
        litchi_iwa_core::Error::InvalidArchive { .. }
        | litchi_iwa_core::Error::InvalidLimits { .. }
        | litchi_iwa_core::Error::HeaderCodec { .. }
        | litchi_iwa_core::Error::Io(_)
        | litchi_iwa_core::Error::Snappy { .. } => BodyDrawableOrderError::InvalidSource,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn move_commands_use_back_to_front_positions() {
        let source: Arc<[u8]> = Arc::from([1_u8]);
        let handles = [1_u64, 2, 3, 4]
            .into_iter()
            .enumerate()
            .map(|(position, identifier)| {
                BodyDrawableHandle::new(
                    Arc::clone(&source),
                    NonZeroU64::new(identifier).unwrap(),
                    Position::new(position),
                )
            })
            .collect::<Vec<_>>();
        for (movement, expected) in [
            (DrawableLayerMove::ToBack, vec![3, 1, 2, 4]),
            (DrawableLayerMove::Backward, vec![1, 3, 2, 4]),
            (DrawableLayerMove::Forward, vec![1, 2, 4, 3]),
            (DrawableLayerMove::ToFront, vec![1, 2, 4, 3]),
        ] {
            let index = 2;
            let target = match movement {
                DrawableLayerMove::ToBack => 0,
                DrawableLayerMove::Backward => index - 1,
                DrawableLayerMove::Forward => index + 1,
                DrawableLayerMove::ToFront => 3,
            };
            let mut moved = handles.clone();
            let value = moved.remove(index);
            moved.insert(target, value);
            assert_eq!(
                moved
                    .iter()
                    .map(|handle| handle.identity.get())
                    .collect::<Vec<_>>(),
                expected
            );
        }
    }

    #[test]
    fn permutation_rejects_duplicates_and_foreign_handles() {
        let current = [NonZeroU64::new(1).unwrap(), NonZeroU64::new(2).unwrap()];
        assert_eq!(
            validate_exact_permutation(&current, &[current[0], current[0]]),
            Err(BodyDrawableOrderError::InvalidOrder)
        );
        assert_eq!(
            validate_exact_permutation(&current, &[current[0], NonZeroU64::new(3).unwrap()]),
            Err(BodyDrawableOrderError::InvalidOrder)
        );
    }

    #[test]
    fn permutation_accepts_a_valid_empty_order() {
        assert_eq!(validate_exact_permutation(&[], &[]), Ok(()));
        assert_eq!(
            validate_exact_permutation(&[], &[NonZeroU64::new(1).unwrap()]),
            Err(BodyDrawableOrderError::InvalidOrder)
        );
    }

    #[test]
    fn root_reference_rejects_duplicate_optional_fields() {
        // identifier = 42, deprecated_type = 1 twice.
        let duplicate_type = [0x08, 0x2a, 0x10, 0x01, 0x10, 0x01];
        assert_eq!(
            strict_reference_with_limits(&duplicate_type, WireLimits::default()),
            Err(BodyDrawableOrderError::InvalidSource)
        );

        // identifier = 42, deprecated_is_external = false twice.
        let duplicate_external = [0x08, 0x2a, 0x18, 0x00, 0x18, 0x00];
        assert_eq!(
            strict_reference_with_limits(&duplicate_external, WireLimits::default()),
            Err(BodyDrawableOrderError::InvalidSource)
        );
    }

    #[test]
    fn root_reference_rejects_noncanonical_optional_values() {
        // The value 1 must use one byte, not the two-byte 0x81 0x00 form.
        let noncanonical_type = [0x08, 0x2a, 0x10, 0x81, 0x00];
        assert_eq!(
            strict_reference_with_limits(&noncanonical_type, WireLimits::default()),
            Err(BodyDrawableOrderError::InvalidSource)
        );

        // Bool values are canonical only when they are exactly zero or one.
        let noncanonical_external = [0x08, 0x2a, 0x18, 0x02];
        assert_eq!(
            strict_reference_with_limits(&noncanonical_external, WireLimits::default()),
            Err(BodyDrawableOrderError::InvalidSource)
        );
    }

    #[test]
    fn external_reference_is_rejected_for_root_and_order_payloads() {
        let root_external = [0x08, 0x2a, 0x18, 0x01];
        assert_eq!(
            strict_reference_with_limits(&root_external, WireLimits::default()),
            Err(BodyDrawableOrderError::InvalidSource)
        );

        // drawables = [TSP.Reference(identifier: 42, deprecated_is_external: true)].
        let order_external = [0x0a, 0x04, 0x08, 0x2a, 0x18, 0x01];
        let snapshot = drawable_order_codec::decode_drawable_order(
            &order_external,
            DecodeOptions::for_source(&order_external),
        )
        .expect("the generic codec should preserve the external marker");
        assert_eq!(
            reject_external_references(snapshot),
            Err(BodyDrawableOrderError::InvalidSource)
        );
    }
}
