//! Exact-source transactions for Pages body-footnote text.
//!
//! This module deliberately owns only replacement of the user-visible text
//! inside an existing footnote storage and the optional custom marker string
//! on its existing reference attachment. The body anchor, body footnote table,
//! native reference/storage/marker edges, and all object ownership metadata
//! remain untouched. Insert/remove therefore stay in the migration host until
//! their graph lifecycle can be moved without weakening cleanup.

use std::collections::HashMap;
use std::fmt;
use std::num::NonZeroU64;
use std::sync::Arc;

use litchi_iwa_archive::{ComponentCatalog, package::EntryEdit};
use litchi_iwa_common::{WireLimits, encode_varint_into, varint::encoded_len, wire::WireView};
use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
use litchi_iwa_protos::pages_footnote_codec;
use litchi_iwa_text_wire::RewriteError;
use thiserror::Error as ThisError;

use super::{
    FOOTNOTE_REFERENCE_MESSAGE_TYPE, MAX_BODY_FOOTNOTES, Package, PackageError,
    STORAGE_TEXT_PREFIX, TEXTUAL_ATTACHMENT_MESSAGE_TYPE, decode_body_storage_with_wire_error,
    effective_text_limit, find_object, footnote_decode_options, footnote_marker_decode_options,
    footnote_marker_identifier, is_body_text_message_type, root_references_with_limits,
    storage_rewrite_limits, unique_message_payload, unique_text_payload,
};
use crate::footnote::body::{Footnote, Position, Selector};

const FOOTNOTE_MARK_KIND: i32 = 2;

/// A finite resource governed while one Pages footnote text is rewritten or
/// published.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Complete input package bytes.
    InputBytes,
    /// Complete rewritten package bytes.
    OutputBytes,
    /// ZIP members, IWA objects, messages, or table entries.
    Entries,
    /// Bytes in one package member, IWA object, or message.
    EntryBytes,
    /// Aggregate package or IWA bytes.
    TotalBytes,
    /// Aggregate semantic footnote text bytes.
    TextBytes,
    /// UTF-16 units in one footnote storage.
    TextUnits,
    /// Bytes in one protobuf payload.
    WireBytes,
    /// Parsed protobuf fields.
    WireFields,
    /// Protobuf nesting depth.
    WireNesting,
    /// Aggregate protobuf scan and rewrite work.
    WireWork,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::OutputBytes => "output bytes",
            Self::Entries => "entries",
            Self::EntryBytes => "entry bytes",
            Self::TotalBytes => "total bytes",
            Self::TextBytes => "text bytes",
            Self::TextUnits => "text UTF-16 units",
            Self::WireBytes => "wire bytes",
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting",
            Self::WireWork => "wire work",
        })
    }
}

/// An error raised while selecting, staging, or publishing body-footnote
/// text.
///
/// Diagnostics intentionally omit authored text, native identifiers, member
/// names, paths, raw bytes, and lower-layer error strings.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies it
    /// (record 0351 §2.4, record 0361).
    #[error(transparent)]
    Refused(Refusal),
    /// The selected native footnote cannot be rewritten without ambiguity.
    #[error("the Pages source cannot be edited safely")]
    InvalidSource,
    /// A finite transaction resource ceiling was exceeded.
    #[error("Pages footnote-text {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        /// Resource category that exceeded its limit.
        kind: LimitKind,
        /// Observed or requested amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded destination allocation failed before publication.
    #[error("could not allocate {amount} units for the Pages footnote-text transaction")]
    Allocation {
        /// Elements or bytes requested.
        amount: usize,
    },
    /// Full semantic readback did not reproduce the requested change.
    #[error("the edited Pages footnote text failed semantic verification")]
    Verification,
    /// The patch does not belong to this exact immutable package artifact.
    #[error("the Pages footnote-text patch does not match the exact source package")]
    PatchConflict,
}

/// A well-formed request this family's route refuses, classified by
/// [`Refusal::kind`] (record 0351 §2.4, record 0361).
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    /// No footnote matched the source-order selector.
    #[error("the Pages body-footnote selector did not match a footnote")]
    NotFound,
    /// The requested replacement contains a native structural marker.
    #[error("Pages footnote text cannot contain native anchor or attachment markers")]
    StructuralMarker,
    /// The replacement exceeds the semantic footnote text budget.
    #[error("Pages footnote text exceeds its semantic byte budget")]
    TextTooLarge,
    /// The replacement exceeds the semantic custom-marker budget.
    #[error("Pages footnote custom marker exceeds its semantic byte budget")]
    CustomMarkTooLarge,
    /// The snapshot has no exact physical source suitable for a changed edit.
    #[error("this Pages source does not support physical footnote-text edits")]
    UnsupportedSource,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::NotFound { .. } => RefusalKind::NotFound,
            Self::StructuralMarker { .. } | Self::UnsupportedSource { .. } => {
                RefusalKind::UnsupportedContent
            },
            Self::TextTooLarge { .. } | Self::CustomMarkTooLarge { .. } => {
                RefusalKind::InvalidValue
            },
        }
    }
}

impl From<Refusal> for Error {
    fn from(refusal: Refusal) -> Self {
        Self::Refused(refusal)
    }
}

/// One mutable body-footnote text and custom-marker value staged against an
/// immutable package.
pub struct Edit<'a> {
    source: &'a Package,
    selector: Selector,
    position: Option<Position>,
    before: Option<Footnote>,
    text: String,
    custom_mark: Option<String>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("position", &self.position)
            .finish_non_exhaustive()
    }
}

impl<'a> Edit<'a> {
    fn resolve(&mut self) -> Result<(), Error> {
        if self.position.is_none() {
            let (position, before, _) = resolve_footnote(self.source, self.selector)?;
            self.text = try_owned_text(&before.text)?;
            self.custom_mark = before
                .custom_mark
                .as_deref()
                .map(try_owned_text)
                .transpose()?;
            self.position = Some(position);
            self.before = Some(before);
        }
        Ok(())
    }

    fn new(source: &'a Package, selector: Selector) -> Result<Self, Error> {
        Ok(Self {
            source,
            selector,
            position: None,
            before: None,
            text: String::new(),
            custom_mark: None,
        })
    }

    /// Return the checked UTF-16 body-anchor position, resolving on first access.
    pub fn position(&mut self) -> Result<Position, Error> {
        self.resolve()?;
        self.position.ok_or(Error::InvalidSource)
    }

    /// Borrow the original semantic footnote value.
    pub fn before(&mut self) -> Result<&Footnote, Error> {
        self.resolve()?;
        self.before.as_ref().ok_or(Error::InvalidSource)
    }

    /// Borrow the text currently staged for publication.
    pub fn text(&mut self) -> Result<&str, Error> {
        self.resolve()?;
        Ok(&self.text)
    }

    /// Borrow the custom marker currently staged for publication.
    pub fn custom_mark(&mut self) -> Result<Option<&str>, Error> {
        self.resolve()?;
        Ok(self.custom_mark.as_deref())
    }

    /// Stage replacement of the selected footnote's user-visible text.
    ///
    /// The native body anchor and footnote graph are not changed. Structural
    /// U+000E and U+FFFC markers remain reserved for native graph edges.
    pub fn set(&mut self, text: &str) -> Result<&mut Self, Error> {
        self.resolve()?;
        validate_text(text)?;
        let owned = try_owned_text(text)?;
        self.text = owned;
        Ok(self)
    }

    /// Stage replacement or removal of the selected footnote's custom marker.
    ///
    /// `None` removes the native field while `Some("")` retains an explicitly
    /// present empty native string. The reference, storage, and marker graph
    /// identities are not changed.
    pub fn set_custom_mark(&mut self, custom_mark: Option<&str>) -> Result<&mut Self, Error> {
        self.resolve()?;
        validate_custom_mark(custom_mark)?;
        self.custom_mark = custom_mark.map(try_owned_text).transpose()?;
        Ok(self)
    }

    /// Validate and atomically publish the staged immutable candidate.
    ///
    /// Exact semantic no-ops reuse the original source allocation. Changed
    /// edits require an exact source catalog, rewrite only the selected
    /// footnote storage and/or reference message through source-preserving
    /// wire editors, then reopen and verify the complete candidate before
    /// return.
    pub fn commit(mut self) -> Result<Commit, Error> {
        self.resolve()?;
        validate_text(&self.text)?;
        validate_custom_mark(self.custom_mark.as_deref())?;
        self.source.validate().map_err(map_package_error)?;
        let position = self.position.ok_or(Error::InvalidSource)?;
        let before_source = self.before.as_ref().ok_or(Error::InvalidSource)?;
        let (current_position, current, native_footnote) =
            resolve_footnote(self.source, Selector::Anchor(position))?;
        if current_position != position || current != *before_source {
            return Err(Error::InvalidSource);
        }
        if current.text.as_ref() != before_source.text.as_ref() {
            return Err(Error::InvalidSource);
        }
        if current.custom_mark != before_source.custom_mark {
            return Err(Error::InvalidSource);
        }

        let source_bytes = self.source.state.source.shared_source();
        let source_fingerprint = super::section_transaction::fingerprint(&source_bytes);
        let before = clone_footnote(before_source)?;
        let after = Footnote::with_custom_mark(
            position,
            try_owned_text(&self.text)?.into_boxed_str(),
            self.custom_mark
                .as_deref()
                .map(try_owned_text)
                .transpose()?
                .map(String::into_boxed_str),
        )
        .map_err(map_footnote_value_error)?;

        if before == after {
            return Ok(Commit {
                package: self.source.clone(),
                patch: Patch {
                    source_bytes: Arc::clone(&source_bytes),
                    target_bytes: source_bytes,
                    source_fingerprint,
                    target_fingerprint: source_fingerprint,
                    position,
                    before,
                    after,
                },
                diagnostics: Diagnostics::unchanged(),
            });
        }
        if !self.source.state.source.source_is_exact() {
            return Err(Error::Refused(Refusal::UnsupportedSource));
        }

        let package = rewrite_package_footnote(
            self.source,
            native_footnote,
            before_source.text.as_ref(),
            self.text.as_str(),
            before_source.custom_mark.as_deref(),
            self.custom_mark.as_deref(),
            position,
        )?;
        let touched_components = changed_component_count(self.source, &package)?;
        if touched_components == 0 {
            return Err(Error::Verification);
        }
        let target_bytes = package.state.source.shared_source();
        let target_fingerprint = super::section_transaction::fingerprint(&target_bytes);
        Ok(Commit {
            package,
            patch: Patch {
                source_bytes,
                target_bytes,
                source_fingerprint,
                target_fingerprint,
                position,
                before,
                after,
            },
            diagnostics: Diagnostics::published(touched_components),
        })
    }
}

/// An exact-source-checked, reversible Pages body-footnote text patch.
///
/// Native identifiers, member names, and exact source/target bytes remain
/// private. Fingerprints are diagnostics only; exact byte comparison
/// authorizes application.
#[derive(Clone, PartialEq, Eq)]
pub struct Patch {
    source_bytes: Arc<[u8]>,
    target_bytes: Arc<[u8]>,
    source_fingerprint: u64,
    target_fingerprint: u64,
    position: Position,
    before: Footnote,
    after: Footnote,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("position", &self.position)
            .finish_non_exhaustive()
    }
}

impl Patch {
    /// Return the semantic UTF-16 body-anchor position selected by this patch.
    #[must_use]
    pub const fn position(&self) -> Position {
        self.position
    }

    /// Borrow the complete semantic footnote required from the patch source.
    #[must_use]
    pub const fn before(&self) -> &Footnote {
        &self.before
    }

    /// Borrow the complete semantic footnote produced by the patch target.
    #[must_use]
    pub const fn after(&self) -> &Footnote {
        &self.after
    }

    /// Return the base package's compact diagnostic fingerprint.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.source_fingerprint
    }

    /// Return the committed package's compact diagnostic fingerprint.
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.target_fingerprint
    }

    /// Return whether the patch preserves both semantic value and exact bytes.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after
            && self.source_fingerprint == self.target_fingerprint
            && self.source_bytes.as_ref() == self.target_bytes.as_ref()
    }

    /// Return an exact reversible patch from the target back to its source.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            source_bytes: Arc::clone(&self.target_bytes),
            target_bytes: Arc::clone(&self.source_bytes),
            source_fingerprint: self.target_fingerprint,
            target_fingerprint: self.source_fingerprint,
            position: self.position,
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }
}

/// Compact evidence describing one footnote-text commit.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Diagnostics {
    changed: bool,
    touched_components: usize,
    full_reparse_performed: bool,
}

impl Diagnostics {
    const fn unchanged() -> Self {
        Self {
            changed: false,
            touched_components: 0,
            full_reparse_performed: false,
        }
    }

    const fn published(touched_components: usize) -> Self {
        Self {
            changed: true,
            touched_components,
            full_reparse_performed: true,
        }
    }

    /// Return whether the committed package differs from its source.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// Return the number of physical IWA components rewritten.
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

/// The fully verified result of one immutable footnote-text transaction.
#[must_use = "a Pages footnote-text commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct Commit {
    package: Package,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
    /// Borrow the fully reopened immutable package snapshot.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Consume the commit and return its immutable package snapshot.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Consume the commit and return its package and patch.
    #[must_use]
    pub fn into_parts(self) -> (Package, Patch) {
        (self.package, self.patch)
    }

    /// Borrow the reversible exact-source patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Borrow compact publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }
}

impl Package {
    /// Read every semantic body footnote through the immutable package owner.
    ///
    /// This alias keeps mutation call sites close to the setter transaction;
    /// ordinary reads remain available through [`Package::body_footnotes`].
    pub fn edit_body_footnote_text(&self, selector: Selector) -> Result<Edit<'_>, Error> {
        Edit::new(self, selector)
    }

    /// Apply an exact-source-checked body-footnote text patch.
    ///
    /// The retained target is fully reopened and semantically verified under
    /// this package's original limits before it is published.
    pub fn apply_body_footnote_text(&self, patch: &Patch) -> Result<Commit, Error> {
        let source = &self.state.source;
        let source_bytes = source.shared_source();
        if super::section_transaction::fingerprint(source.source_bytes())
            != patch.source_fingerprint
            || source.source_bytes() != patch.source_bytes.as_ref()
            || source_bytes.as_ref() != patch.source_bytes.as_ref()
        {
            return Err(Error::PatchConflict);
        }
        self.validate().map_err(map_package_error)?;
        let (_, current, _) = resolve_footnote(self, Selector::Anchor(patch.position))?;
        if current != patch.before {
            return Err(Error::PatchConflict);
        }
        if patch.is_noop() {
            return Ok(Commit {
                package: self.clone(),
                patch: patch.clone(),
                diagnostics: Diagnostics::unchanged(),
            });
        }
        if !source.source_is_exact()
            || super::section_transaction::fingerprint(&patch.target_bytes)
                != patch.target_fingerprint
        {
            return Err(Error::PatchConflict);
        }
        let candidate_source = source
            .reopen(Arc::clone(&patch.target_bytes))
            .map_err(map_archive_error)?;
        let candidate =
            Package::from_source_catalog(candidate_source).map_err(map_package_error)?;
        candidate.validate().map_err(map_package_error)?;
        let touched_components = changed_component_count(self, &candidate)?;
        if touched_components == 0 {
            return Err(Error::Verification);
        }
        verify_candidate(self, &candidate, patch.position, &patch.after)?;
        Ok(Commit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: Diagnostics::published(touched_components),
        })
    }
}

#[derive(Debug, Clone, Copy)]
pub(super) struct NativeFootnote {
    pub(super) body_identifier: NonZeroU64,
    pub(super) position: Position,
    pub(super) reference_identifier: NonZeroU64,
    pub(super) storage_identifier: NonZeroU64,
    pub(super) marker_identifier: NonZeroU64,
}

/// Borrowed object locations for one immutable source snapshot.
///
/// The map owns no archive data: every value remains borrowed from the
/// source catalog retained by the package's `Arc`.  Keeping the first value
/// for an identifier preserves the historical `find_object` first-match
/// behavior even if this helper is ever used before package-wide duplicate
/// validation.
#[derive(Debug)]
struct FootnoteObjectLocations<'source> {
    objects: HashMap<u64, &'source ArchiveObject>,
}

impl<'source> FootnoteObjectLocations<'source> {
    fn new(
        components: &'source ComponentCatalog,
        budget: &mut FootnoteGraphBudget,
    ) -> Result<Self, Error> {
        let object_count = components.iter().try_fold(0usize, |count, component| {
            count
                .checked_add(component.archive().objects.len())
                .ok_or(Error::Limit {
                    kind: LimitKind::Entries,
                    observed: usize_to_u64(usize::MAX),
                    maximum: usize_to_u64(super::MAX_OBJECTS),
                })
        })?;
        budget.charge_references(object_count)?;
        budget.charge_work(object_count)?;
        if object_count > super::MAX_OBJECTS {
            return Err(Error::Limit {
                kind: LimitKind::Entries,
                observed: usize_to_u64(object_count),
                maximum: usize_to_u64(super::MAX_OBJECTS),
            });
        }

        let mut objects = HashMap::new();
        objects
            .try_reserve(object_count)
            .map_err(|_error| Error::Allocation {
                amount: object_count,
            })?;
        for component in components.iter() {
            for object in &component.archive().objects {
                if let Some(identifier) = object.archive_info.identifier {
                    objects.entry(identifier).or_insert(object);
                }
            }
        }
        Ok(Self { objects })
    }

    fn get(&self, identifier: u64) -> Option<&'source ArchiveObject> {
        self.objects.get(&identifier).copied()
    }
}

/// Aggregate source, graph-reference, and scan work charged while a footnote
/// graph is projected.  The ordinary package reader bounds each payload and
/// the semantic projection bounds retained text, but the transaction keeps
/// both its source and reopened candidate alive during verification.  Sharing
/// this coordinator across those two graph passes prevents two individually
/// valid walks from silently doubling the work and source allowance.
#[derive(Debug, Clone, Copy)]
pub(super) struct FootnoteGraphBudget {
    maximum_source_bytes: usize,
    maximum_references: usize,
    maximum_work: usize,
    source_bytes: usize,
    references: usize,
    work: usize,
}

impl FootnoteGraphBudget {
    pub(super) fn new(limits: super::Limits) -> Result<Self, Error> {
        let archive = limits
            .effective_archive_limits()
            .map_err(map_archive_error)?;
        let maximum_source_bytes = usize::try_from(limits.max_input_bytes())
            .unwrap_or(usize::MAX)
            .checked_mul(2)
            .unwrap_or(usize::MAX);
        let maximum_references = archive
            .max_metadata_items()
            .checked_mul(limits.max_entries().max(1))
            .unwrap_or(usize::MAX);
        let maximum_work = limits
            .max_iwa_stream_bytes()
            .checked_mul(16)
            .unwrap_or(usize::MAX);
        Ok(Self {
            maximum_source_bytes,
            maximum_references,
            maximum_work,
            source_bytes: 0,
            references: 0,
            work: 0,
        })
    }

    pub(super) fn charge_source(&mut self, amount: usize) -> Result<(), Error> {
        charge_graph_counter(
            &mut self.source_bytes,
            amount,
            self.maximum_source_bytes,
            LimitKind::InputBytes,
        )
    }

    pub(super) fn charge_references(&mut self, amount: usize) -> Result<(), Error> {
        charge_graph_counter(
            &mut self.references,
            amount,
            self.maximum_references,
            LimitKind::Entries,
        )
    }

    pub(super) fn charge_work(&mut self, amount: usize) -> Result<(), Error> {
        charge_graph_counter(
            &mut self.work,
            amount,
            self.maximum_work,
            LimitKind::WireWork,
        )
    }

    #[cfg(test)]
    const fn with_test_limits(
        maximum_source_bytes: usize,
        maximum_references: usize,
        maximum_work: usize,
    ) -> Self {
        Self {
            maximum_source_bytes,
            maximum_references,
            maximum_work,
            source_bytes: 0,
            references: 0,
            work: 0,
        }
    }
}

fn charge_graph_counter(
    current: &mut usize,
    amount: usize,
    maximum: usize,
    kind: LimitKind,
) -> Result<(), Error> {
    let observed = current.checked_add(amount).unwrap_or(usize::MAX);
    if observed > maximum {
        return Err(Error::Limit {
            kind,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        });
    }
    *current = observed;
    Ok(())
}

fn resolve_footnote(
    package: &Package,
    selector: Selector,
) -> Result<(Position, Footnote, NativeFootnote), Error> {
    let graphs = native_footnotes(package)?;
    let index = match selector {
        Selector::Position(position) => position.get(),
        Selector::Anchor(position) => graphs
            .iter()
            .position(|graph| graph.position == position)
            .ok_or(Error::Refused(Refusal::NotFound))?,
    };
    let graph = graphs.get(index).ok_or(Error::Refused(Refusal::NotFound))?;
    let footnotes = package.body_footnotes().map_err(map_package_error)?;
    let footnote = footnotes.get(index).ok_or(Error::InvalidSource)?;
    Ok((graph.position, clone_footnote(footnote)?, *graph))
}

pub(super) fn native_footnotes(package: &Package) -> Result<Vec<NativeFootnote>, Error> {
    let mut budget = FootnoteGraphBudget::new(package.state.source.limits())?;
    native_footnotes_with_budget(package, &mut budget)
}

pub(super) fn native_footnotes_with_budget(
    package: &Package,
    budget: &mut FootnoteGraphBudget,
) -> Result<Vec<NativeFootnote>, Error> {
    let mut semantic_budget =
        super::FootnoteSemanticBudget::new(effective_text_limit(package.state.source.limits()));
    let components = package.state.source.components();
    budget.charge_source(package.state.source.source_bytes().len())?;
    budget.charge_work(components.len())?;
    let root = root_references_with_limits(components, package.state.source.limits())
        .map_err(map_package_error)?;
    let Some(body_identifier) = root.body else {
        return Ok(Vec::new());
    };
    let body = find_object(components, body_identifier.get()).ok_or(Error::InvalidSource)?;
    let body_payload =
        unique_text_payload(&body.messages, body_identifier).map_err(map_package_error)?;
    budget.charge_work(body_payload.len())?;
    let (body_storage, _) = decode_body_storage_with_wire_error(
        &body.messages,
        body_identifier,
        super::MAX_SECTIONS,
        effective_text_limit(package.state.source.limits()),
        package.state.source.limits(),
    )
    .map_err(map_body_storage_decode_error)?;
    budget.charge_work(body_payload.len())?;
    checked_utf16_units(body_storage.text())?;
    budget.charge_work(body_payload.len())?;
    let entries =
        super::footnote_table_entries(body_payload, body_identifier, package.state.source.limits())
            .map_err(|error| map_package_error_with_kind(error, LimitKind::Entries))?;
    if entries.len() > MAX_BODY_FOOTNOTES {
        return Err(Error::Limit {
            kind: LimitKind::Entries,
            observed: usize_to_u64(entries.len()),
            maximum: usize_to_u64(MAX_BODY_FOOTNOTES),
        });
    }
    budget.charge_references(entries.len())?;
    budget.charge_work(entries.len())?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(entries.len())
        .map_err(|_error| Error::Allocation {
            amount: entries.len(),
        })?;
    let mut seen_references = Vec::new();
    let mut seen_storages = Vec::new();
    let mut seen_markers = Vec::new();
    for seen in [&mut seen_references, &mut seen_storages, &mut seen_markers] {
        seen.try_reserve_exact(entries.len())
            .map_err(|_error| Error::Allocation {
                amount: entries.len(),
            })?;
    }
    let object_locations = FootnoteObjectLocations::new(components, budget)?;
    let mut previous = None;
    for entry in entries {
        if previous.is_some_and(|position| position >= entry.character_index) {
            return Err(Error::InvalidSource);
        }
        previous = Some(entry.character_index);
        if seen_references.contains(&entry.identifier) {
            return Err(Error::InvalidSource);
        }
        seen_references.push(entry.identifier);
        super::validate_body_footnote_anchor(
            body_storage.text(),
            body_identifier,
            entry.character_index,
        )
        .map_err(map_package_error)?;
        let reference = object_locations
            .get(entry.identifier.get())
            .ok_or(Error::InvalidSource)?;
        let payload = unique_message_payload(
            &reference.messages,
            FOOTNOTE_REFERENCE_MESSAGE_TYPE,
            "Pages footnote reference",
        )
        .map_err(map_package_error)?;
        budget.charge_work(payload.len())?;
        let decoded = pages_footnote_codec::decode_footnote_reference(
            payload,
            footnote_decode_options(payload, package.state.source.limits())
                .map_err(map_package_error)?,
        )
        .map_err(|_error| Error::InvalidSource)?;
        if decoded
            .super_kind()
            .is_some_and(|kind| kind != FOOTNOTE_MARK_KIND)
        {
            return Err(Error::InvalidSource);
        }
        let storage_identifier = decoded
            .contained_storage()
            .map(|reference| reference.identifier())
            .ok_or(Error::InvalidSource)?;
        if storage_identifier == entry.identifier || seen_storages.contains(&storage_identifier) {
            return Err(Error::InvalidSource);
        }
        seen_storages.push(storage_identifier);
        let storage = object_locations
            .get(storage_identifier.get())
            .ok_or(Error::InvalidSource)?;
        let storage_payload = unique_text_payload(&storage.messages, storage_identifier)
            .map_err(map_package_error)?;
        budget.charge_work(storage_payload.len())?;
        let (storage_value, _) = decode_body_storage_with_wire_error(
            &storage.messages,
            storage_identifier,
            super::MAX_SECTIONS,
            effective_text_limit(package.state.source.limits()),
            package.state.source.limits(),
        )
        .map_err(map_body_storage_decode_error)?;
        budget.charge_work(storage_payload.len())?;
        budget.charge_references(1)?;
        let text = storage_value
            .text()
            .strip_prefix(STORAGE_TEXT_PREFIX)
            .ok_or(Error::InvalidSource)?;
        if text.len() > crate::footnote::body::MAX_TEXT_BYTES {
            return Err(Error::InvalidSource);
        }
        let custom_mark = decoded.custom_mark_string();
        if custom_mark
            .is_some_and(|value| value.len() > crate::footnote::body::MAX_CUSTOM_MARK_BYTES)
        {
            return Err(Error::InvalidSource);
        }
        super::FootnoteSemanticBudget::charge(
            &mut semantic_budget,
            text.len(),
            custom_mark.map_or(0, str::len),
        )
        .map_err(|error| map_package_error_with_kind(error, LimitKind::TextBytes))?;
        let marker_identifier = footnote_marker_identifier(
            storage_payload,
            storage_identifier,
            package.state.source.limits(),
        )
        .map_err(map_package_error)?;
        if marker_identifier == entry.identifier
            || marker_identifier == storage_identifier
            || seen_markers.contains(&marker_identifier)
        {
            return Err(Error::InvalidSource);
        }
        seen_markers.push(marker_identifier);
        let marker = object_locations
            .get(marker_identifier.get())
            .ok_or(Error::InvalidSource)?;
        let marker_payload = unique_message_payload(
            &marker.messages,
            TEXTUAL_ATTACHMENT_MESSAGE_TYPE,
            "Pages footnote marker",
        )
        .map_err(map_package_error)?;
        budget.charge_work(marker_payload.len())?;
        budget.charge_references(1)?;
        let marker_value =
            litchi_iwa_protos::pages_footnote_marker_codec::decode_textual_attachment(
                marker_payload,
                footnote_marker_decode_options(marker_payload, package.state.source.limits())
                    .map_err(map_package_error)?,
            )
            .map_err(|_error| Error::InvalidSource)?;
        if marker_value.kind() != Some(FOOTNOTE_MARK_KIND) {
            return Err(Error::InvalidSource);
        }
        let position = position_from_anchor(entry.character_index)?;
        output.push(NativeFootnote {
            body_identifier,
            position,
            reference_identifier: entry.identifier,
            storage_identifier,
            marker_identifier,
        });
    }
    Ok(output)
}

fn rewrite_package_footnote(
    source: &Package,
    native_footnote: NativeFootnote,
    before_text: &str,
    after_text: &str,
    before_custom_mark: Option<&str>,
    after_custom_mark: Option<&str>,
    position: Position,
) -> Result<Package, Error> {
    let source_catalog = &source.state.source;
    let text_component = (before_text != after_text)
        .then(|| component_name_for_object(source, native_footnote.storage_identifier))
        .transpose()?;
    let mark_component = (before_custom_mark != after_custom_mark)
        .then(|| component_name_for_object(source, native_footnote.reference_identifier))
        .transpose()?;
    let first_component = text_component.or(mark_component);
    let first_component = first_component.ok_or(Error::InvalidSource)?;
    let second_component = match (text_component, mark_component) {
        (Some(text), Some(mark)) if text != mark => Some(mark),
        _ => None,
    };
    // Resolve every dependency and physical authority before an in-memory
    // archive is edited.  A semantic footnote value alone cannot prove that
    // its storage/reference/marker objects are exclusively owned by this
    // body-footnote graph, nor that the selected ZIP members are canonical
    // exact authorities.
    let mut graph_budget = FootnoteGraphBudget::new(source_catalog.limits())?;
    graph_budget.charge_source(source_catalog.source_bytes().len())?;
    prove_footnote_ownership(source, native_footnote, &mut graph_budget)?;
    validate_mutation_component(source, first_component)?;
    if let Some(component) = second_component {
        validate_mutation_component(source, component)?;
    }
    let physical_limits = source_catalog.limits();
    let archive_limits = physical_limits
        .effective_archive_limits()
        .map_err(map_archive_error)?;

    let mut rewritten_entries = Vec::new();
    let component_count = if second_component.is_some() { 2 } else { 1 };
    rewritten_entries
        .try_reserve_exact(component_count)
        .map_err(|_error| Error::Allocation {
            amount: component_count,
        })?;
    for component_name in [Some(first_component), second_component]
        .into_iter()
        .flatten()
    {
        let rewritten = rewrite_component(
            source,
            component_name,
            native_footnote,
            text_component == Some(component_name),
            mark_component == Some(component_name),
            before_text,
            after_text,
            before_custom_mark,
            after_custom_mark,
            archive_limits,
        )?;
        rewritten_entries.push((try_owned_text(component_name)?, rewritten));
    }

    let mut edits = Vec::new();
    edits
        .try_reserve_exact(rewritten_entries.len())
        .map_err(|_error| Error::Allocation {
            amount: rewritten_entries.len(),
        })?;
    for (component_name, compressed) in &rewritten_entries {
        edits.push(EntryEdit::new(component_name, compressed));
    }
    let output = source_catalog
        .package()
        .reassemble_to_bytes(&edits, physical_limits)
        .map_err(map_archive_error)?;
    let candidate_source = source_catalog
        .reopen(output.into())
        .map_err(map_archive_error)?;
    let candidate = Package::from_source_catalog(candidate_source).map_err(map_package_error)?;
    candidate.validate().map_err(map_package_error)?;
    let _ = changed_component_count(source, &candidate)?;
    let expected = Footnote::with_custom_mark(
        position,
        try_owned_text(after_text)?.into_boxed_str(),
        after_custom_mark
            .map(try_owned_text)
            .transpose()?
            .map(String::into_boxed_str),
    )
    .map_err(map_footnote_value_error)?;
    verify_candidate(source, &candidate, position, &expected)?;
    Ok(candidate)
}

fn component_name_for_object(package: &Package, identifier: NonZeroU64) -> Result<&str, Error> {
    let mut matching = package
        .state
        .source
        .components()
        .iter()
        .filter(|component| component.archive().object(identifier.get()).is_some());
    let component = matching.next().ok_or(Error::InvalidSource)?;
    if matching.next().is_some() {
        return Err(Error::InvalidSource);
    }
    if package
        .state
        .source
        .package()
        .iter()
        .find(|entry| entry.name() == component.name())
        .is_some_and(|entry| entry.is_opaque())
    {
        return Err(Error::InvalidSource);
    }
    Ok(component.name())
}

/// Validate one selected IWA member's physical authority before mutation.
///
/// Generic package ingress intentionally retains aliases for read-only
/// consumers.  A preserve-mode mutation has a narrower contract: the member
/// selected by an object graph must be one canonical `Index/<name>.iwa`
/// authority whose local and central ZIP names agree byte-for-byte.  This
/// check is intentionally local to the capability so the generic archive
/// owner does not acquire Pages policy.
fn validate_mutation_component(package: &Package, component_name: &str) -> Result<(), Error> {
    if !is_canonical_component_name(component_name) {
        return Err(Error::InvalidSource);
    }
    let entry = package
        .state
        .source
        .package()
        .iter()
        .find(|entry| entry.name() == component_name)
        .ok_or(Error::InvalidSource)?;
    if entry.is_opaque() {
        return Err(Error::InvalidSource);
    }
    let expected = component_name.as_bytes();
    if entry.raw_name() != expected
        || entry.metadata().local().name() != expected
        || entry.metadata().central().name() != expected
        || entry.metadata().local().compression_method()
            != entry.metadata().central().compression_method()
    {
        return Err(Error::InvalidSource);
    }
    Ok(())
}

fn is_canonical_component_name(name: &str) -> bool {
    let Some(basename) = name.strip_prefix("Index/") else {
        return false;
    };
    !basename.is_empty()
        && !basename.contains('/')
        && !basename.contains(['\\', '\0', ':'])
        && !basename.chars().any(char::is_control)
        && basename.ends_with(".iwa")
}

/// Prove that the three native objects selected for a text-only edit have no
/// incoming dependency outside this one body-footnote graph.
///
/// IWA headers may carry object-reference metadata, while Pages also stores
/// the body/reference/storage edges in typed payloads.  Checking both layers
/// matters: the small synthetic sources used by callers often omit header
/// metadata, whereas native Pages writes commonly retain it.  Any alias from
/// another owner is rejected before `rewrite_component` gets a mutable
/// archive, so a successful edit cannot silently change an unrelated graph.
pub(super) fn prove_footnote_ownership(
    package: &Package,
    native_footnote: NativeFootnote,
    budget: &mut FootnoteGraphBudget,
) -> Result<(), Error> {
    let limits = package.state.source.limits();
    for component in package.state.source.components().iter() {
        budget.charge_work(component.archive().objects.len())?;
        for object in &component.archive().objects {
            budget.charge_work(object.messages.len())?;
            let owner = object
                .archive_info
                .identifier
                .and_then(NonZeroU64::new)
                .ok_or(Error::InvalidSource)?;

            for message_info in &object.archive_info.message_infos {
                budget.charge_references(message_info.object_references.len())?;
                budget.charge_references(message_info.data_references.len())?;
                budget.charge_work(message_info.field_infos.len())?;
                for referenced in &message_info.object_references {
                    validate_graph_edge(owner.get(), *referenced, native_footnote)?;
                }
                for field_info in &message_info.field_infos {
                    budget.charge_references(field_info.object_references.len())?;
                    budget.charge_references(field_info.data_references.len())?;
                    budget.charge_work(field_info.path.path.len())?;
                    for referenced in &field_info.object_references {
                        validate_graph_edge(owner.get(), *referenced, native_footnote)?;
                    }
                }
            }

            for message in &object.messages {
                budget.charge_work(message.data.len())?;
                if message.type_ == FOOTNOTE_REFERENCE_MESSAGE_TYPE {
                    let decoded = pages_footnote_codec::decode_footnote_reference(
                        &message.data,
                        footnote_decode_options(&message.data, limits)
                            .map_err(map_package_error)?,
                    );
                    match decoded {
                        Ok(reference) => {
                            if let Some(storage) = reference.contained_storage() {
                                validate_graph_edge(
                                    owner.get(),
                                    storage.identifier().get(),
                                    native_footnote,
                                )?;
                            }
                        },
                        Err(_error) if owner == native_footnote.reference_identifier => {
                            return Err(Error::InvalidSource);
                        },
                        Err(_error) => {},
                    }
                }

                if !is_body_text_message_type(message.type_) {
                    continue;
                }
                let Some(owner_identifier) = NonZeroU64::new(owner.get()) else {
                    return Err(Error::InvalidSource);
                };
                match super::footnote_table_entries(&message.data, owner_identifier, limits) {
                    Ok(entries) => {
                        budget.charge_references(entries.len())?;
                        budget.charge_work(entries.len())?;
                        for entry in entries {
                            validate_graph_edge(
                                owner.get(),
                                entry.identifier.get(),
                                native_footnote,
                            )?;
                        }
                    },
                    Err(_error) if owner == native_footnote.body_identifier => {
                        return Err(Error::InvalidSource);
                    },
                    Err(_error) => {},
                }

                // `footnote_marker_identifier` returns an error for ordinary
                // body/section storage, which is expected.  A successful
                // result identifies a native footnote-storage owner.
                if let Ok(marker) =
                    footnote_marker_identifier(&message.data, owner_identifier, limits)
                {
                    validate_graph_edge(owner.get(), marker.get(), native_footnote)?;
                }
            }
        }
    }
    Ok(())
}

fn validate_graph_edge(
    owner: u64,
    referenced: u64,
    native_footnote: NativeFootnote,
) -> Result<(), Error> {
    let expected_owner = if referenced == native_footnote.reference_identifier.get() {
        Some(native_footnote.body_identifier.get())
    } else if referenced == native_footnote.storage_identifier.get() {
        Some(native_footnote.reference_identifier.get())
    } else if referenced == native_footnote.marker_identifier.get() {
        Some(native_footnote.storage_identifier.get())
    } else {
        None
    };
    if expected_owner.is_some_and(|expected| expected != owner) {
        return Err(Error::InvalidSource);
    }
    Ok(())
}

fn rewrite_component(
    source: &Package,
    component_name: &str,
    native_footnote: NativeFootnote,
    rewrite_text: bool,
    rewrite_custom_mark: bool,
    before_text: &str,
    after_text: &str,
    before_custom_mark: Option<&str>,
    after_custom_mark: Option<&str>,
    archive_limits: litchi_iwa_core::Limits,
) -> Result<Vec<u8>, Error> {
    let source_catalog = &source.state.source;
    let entry = source_catalog
        .package()
        .iter()
        .find(|entry| entry.name() == component_name)
        .ok_or(Error::InvalidSource)?;
    if entry.is_opaque() {
        return Err(Error::InvalidSource);
    }
    let stream = SnappyStream::decompress(
        entry.data(),
        source_catalog
            .limits()
            .snappy_limits()
            .map_err(map_archive_error)?,
    )
    .map_err(map_core_error)?;
    let mut archive = Archive::parse(stream.as_bytes(), archive_limits).map_err(map_core_error)?;
    let rewrite_limits =
        storage_rewrite_limits(source_catalog.limits()).map_err(map_storage_wire_limits_error)?;

    if rewrite_text {
        let (message_index, message_type, rewritten_payload) = {
            let object = archive
                .object(native_footnote.storage_identifier.get())
                .ok_or(Error::InvalidSource)?;
            let mut messages = object
                .messages
                .iter()
                .enumerate()
                .filter(|(_index, message)| is_body_text_message_type(message.type_));
            let (message_index, message) = messages.next().ok_or(Error::InvalidSource)?;
            if messages.next().is_some() {
                return Err(Error::InvalidSource);
            }
            let rewritten =
                rewrite_storage_payload(&message.data, before_text, after_text, rewrite_limits)?;
            (message_index, message.type_, rewritten)
        };
        archive
            .object_mut(native_footnote.storage_identifier.get())
            .ok_or(Error::InvalidSource)?
            .replace_message_preserving_header(
                message_index,
                RawMessage {
                    type_: message_type,
                    data: rewritten_payload,
                },
                archive_limits,
            )
            .map_err(map_core_error)?;
    }

    if rewrite_custom_mark {
        let (message_index, message_type, rewritten_payload) = {
            let object = archive
                .object(native_footnote.reference_identifier.get())
                .ok_or(Error::InvalidSource)?;
            let mut messages = object
                .messages
                .iter()
                .enumerate()
                .filter(|(_index, message)| message.type_ == FOOTNOTE_REFERENCE_MESSAGE_TYPE);
            let (message_index, message) = messages.next().ok_or(Error::InvalidSource)?;
            if messages.next().is_some() {
                return Err(Error::InvalidSource);
            }
            let rewritten = rewrite_reference_payload(
                &message.data,
                native_footnote,
                before_custom_mark,
                after_custom_mark,
                source_catalog.limits(),
                rewrite_limits,
            )?;
            (message_index, message.type_, rewritten)
        };
        archive
            .object_mut(native_footnote.reference_identifier.get())
            .ok_or(Error::InvalidSource)?
            .replace_message_preserving_header(
                message_index,
                RawMessage {
                    type_: message_type,
                    data: rewritten_payload,
                },
                archive_limits,
            )
            .map_err(map_core_error)?;
    }

    let rewritten_archive = archive.to_bytes(archive_limits).map_err(map_core_error)?;
    let snappy_limits = source_catalog
        .limits()
        .snappy_limits()
        .map_err(map_archive_error)?;
    SnappyStream::compress(&rewritten_archive, snappy_limits).map_err(map_core_error)
}

/// Compare source and candidate logical members while allowing only decoded
/// payload changes.  ZIP reassembly necessarily changes offsets, sizes, and
/// CRC metadata for edited members, but every name, raw authority, and
/// compression policy must remain disjoint and stable.  The returned count is
/// also the truthful physical-component diagnostic for a text+custom-marker
/// edit that happens to span two members.
fn changed_component_count(source: &Package, candidate: &Package) -> Result<usize, Error> {
    let source_entry_count = source.state.source.package().iter().count();
    let candidate_entry_count = candidate.state.source.package().iter().count();
    if source_entry_count != candidate_entry_count {
        return Err(Error::Verification);
    }
    let source_entries = source.state.source.package().iter();
    let candidate_entries = candidate.state.source.package().iter();
    let mut changed = 0usize;
    for (source_entry, candidate_entry) in source_entries.zip(candidate_entries) {
        if source_entry.name() != candidate_entry.name()
            || source_entry.raw_name() != candidate_entry.raw_name()
            || source_entry.is_opaque() != candidate_entry.is_opaque()
            || source_entry.metadata().local().name() != candidate_entry.metadata().local().name()
            || source_entry.metadata().central().name()
                != candidate_entry.metadata().central().name()
            || source_entry.metadata().local().compression_method()
                != candidate_entry.metadata().local().compression_method()
            || source_entry.metadata().central().compression_method()
                != candidate_entry.metadata().central().compression_method()
        {
            return Err(Error::Verification);
        }
        if source_entry.data() != candidate_entry.data() {
            changed = changed.checked_add(1).ok_or(Error::Verification)?;
        }
    }
    Ok(changed)
}

fn rewrite_storage_payload(
    payload: &[u8],
    before: &str,
    after: &str,
    rewrite_limits: litchi_iwa_text_wire::RewriteLimits,
) -> Result<Vec<u8>, Error> {
    let storage = litchi_iwa_text_wire::decode_storage(payload, rewrite_limits)
        .map_err(map_text_rewrite_error)?
        .into_storage();
    let full_text = storage.text();
    let content = full_text
        .strip_prefix(STORAGE_TEXT_PREFIX)
        .ok_or(Error::InvalidSource)?;
    if content != before {
        return Err(Error::InvalidSource);
    }
    let prefix_units = checked_utf16_units(STORAGE_TEXT_PREFIX)?;
    let content_units = checked_utf16_units(content)?;
    let content_end = prefix_units
        .checked_add(content_units)
        .ok_or(Error::Limit {
            kind: LimitKind::TextUnits,
            observed: u64::MAX,
            maximum: u64::from(u32::MAX),
        })?;
    let rewritten = litchi_iwa_text_wire::rewrite_storage_text(
        payload,
        prefix_units..content_end,
        after,
        rewrite_limits,
    )
    .map_err(map_text_rewrite_error)?;
    if rewritten.before_utf16_len() != full_text.encode_utf16().count()
        || rewritten.after_utf16_len()
            != prefix_units
                .checked_add(checked_utf16_units(after)?)
                .ok_or(Error::Limit {
                    kind: LimitKind::TextUnits,
                    observed: u64::MAX,
                    maximum: u64::from(u32::MAX),
                })?
        || !rewritten.removed_object_references().is_empty()
        || !rewritten.removed_object_references_by_field().is_empty()
        || rewritten.object_reference_occurrences_before()
            != rewritten.object_reference_occurrences_after()
        || !rewritten.changed()
    {
        return Err(Error::Verification);
    }
    Ok(rewritten.into_bytes())
}

fn rewrite_reference_payload(
    payload: &[u8],
    native_footnote: NativeFootnote,
    before_custom_mark: Option<&str>,
    after_custom_mark: Option<&str>,
    physical_limits: super::Limits,
    rewrite_limits: litchi_iwa_text_wire::RewriteLimits,
) -> Result<Vec<u8>, Error> {
    let decoded = pages_footnote_codec::decode_footnote_reference(
        payload,
        footnote_decode_options(payload, physical_limits).map_err(map_package_error)?,
    )
    .map_err(|_error| Error::InvalidSource)?;
    if decoded.super_kind() != Some(FOOTNOTE_MARK_KIND)
        || decoded
            .contained_storage()
            .is_none_or(|reference| reference.identifier() != native_footnote.storage_identifier)
        || decoded.custom_mark_string() != before_custom_mark
    {
        return Err(Error::InvalidSource);
    }
    let wire_limits = WireLimits::default()
        .with_input_bytes(rewrite_limits.max_message_bytes())
        .and_then(|limits| limits.with_output_bytes(rewrite_limits.max_message_bytes()))
        .and_then(|limits| limits.with_fields(rewrite_limits.max_fields()))
        .and_then(|limits| limits.with_nesting(rewrite_limits.max_nesting()))
        .and_then(|limits| limits.with_rewrite_work(rewrite_limits.max_rewrite_work()))
        .map_err(map_wire_error)?;
    let rewritten =
        rewrite_custom_mark_wire(payload, before_custom_mark, after_custom_mark, wire_limits)?;
    let readback = pages_footnote_codec::decode_footnote_reference(
        &rewritten,
        footnote_decode_options(&rewritten, physical_limits).map_err(map_package_error)?,
    )
    .map_err(|_error| Error::Verification)?;
    if readback.super_kind() != Some(FOOTNOTE_MARK_KIND)
        || readback
            .contained_storage()
            .is_none_or(|reference| reference.identifier() != native_footnote.storage_identifier)
        || readback.custom_mark_string() != after_custom_mark
    {
        return Err(Error::Verification);
    }
    Ok(rewritten)
}

fn rewrite_custom_mark_wire(
    source: &[u8],
    before: Option<&str>,
    after: Option<&str>,
    limits: WireLimits,
) -> Result<Vec<u8>, Error> {
    let view = WireView::parse(source, limits).map_err(map_wire_error)?;
    let mut selected = None;
    for field in view.fields().filter(|field| field.number() == 3) {
        if selected.is_some() || field.wire_type() != 2 {
            return Err(Error::InvalidSource);
        }
        field.validate_canonical_framing().map_err(map_wire_error)?;
        let value = std::str::from_utf8(field.payload()).map_err(|_error| Error::InvalidSource)?;
        if Some(value) != before {
            return Err(Error::InvalidSource);
        }
        selected = Some(field);
    }
    if selected.is_some() != before.is_some() {
        return Err(Error::InvalidSource);
    }
    let selected_length = selected.map_or(0, |field| field.raw().len());
    let replacement_length = after
        .map(|value| encoded_bytes_field_length(3, value.len(), limits))
        .transpose()?
        .unwrap_or(0);
    let output_length = source
        .len()
        .checked_sub(selected_length)
        .and_then(|length| length.checked_add(replacement_length))
        .ok_or(Error::InvalidSource)?;
    if output_length > limits.max_output_bytes() {
        return Err(Error::Limit {
            kind: LimitKind::OutputBytes,
            observed: usize_to_u64(output_length),
            maximum: usize_to_u64(limits.max_output_bytes()),
        });
    }
    let output_fields = view
        .len()
        .checked_sub(usize::from(selected.is_some()))
        .and_then(|fields| fields.checked_add(usize::from(after.is_some())))
        .ok_or(Error::Limit {
            kind: LimitKind::WireFields,
            observed: u64::MAX,
            maximum: usize_to_u64(limits.max_fields()),
        })?;
    if output_fields > limits.max_fields() {
        return Err(Error::Limit {
            kind: LimitKind::WireFields,
            observed: usize_to_u64(output_fields),
            maximum: usize_to_u64(limits.max_fields()),
        });
    }
    let work = view
        .len()
        .checked_add(output_fields)
        .ok_or(Error::InvalidSource)?;
    if work > limits.max_rewrite_work() {
        return Err(Error::Limit {
            kind: LimitKind::WireWork,
            observed: usize_to_u64(work),
            maximum: usize_to_u64(limits.max_rewrite_work()),
        });
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_length)
        .map_err(|_error| Error::Allocation {
            amount: output_length,
        })?;
    let mut emitted = false;
    for field in view.fields() {
        if field.number() == 3 {
            if let Some(value) = after {
                append_bytes_field(&mut output, 3, value.as_bytes());
            }
            emitted = true;
        } else {
            output.extend_from_slice(field.raw());
        }
    }
    if !emitted {
        if let Some(value) = after {
            append_bytes_field(&mut output, 3, value.as_bytes());
        }
    }
    if output.len() != output_length {
        return Err(Error::Verification);
    }
    Ok(output)
}

fn encoded_bytes_field_length(
    number: u32,
    payload_length: usize,
    limits: WireLimits,
) -> Result<usize, Error> {
    encoded_len((u64::from(number) << 3) | 2)
        .checked_add(encoded_len(usize_to_u64(payload_length)))
        .and_then(|length| length.checked_add(payload_length))
        .ok_or(Error::Limit {
            kind: LimitKind::OutputBytes,
            observed: u64::MAX,
            maximum: usize_to_u64(limits.max_output_bytes()),
        })
}

fn append_bytes_field(output: &mut Vec<u8>, number: u32, payload: &[u8]) {
    encode_varint_into(output, (u64::from(number) << 3) | 2);
    encode_varint_into(output, usize_to_u64(payload.len()));
    output.extend_from_slice(payload);
}

fn verify_candidate(
    source: &Package,
    candidate: &Package,
    position: Position,
    expected: &Footnote,
) -> Result<(), Error> {
    // Run the native graph projection first.  Besides checking that the
    // candidate retained the same private identities, this is the pass that
    // charges the aggregate footnote text/custom-marker budget with its
    // dedicated TextBytes diagnostic.  `body_footnotes` is still needed for
    // semantic readback, but its generic payload-limit mapping cannot tell an
    // aggregate semantic refusal from a wire-level refusal.
    let mut graph_budget = FootnoteGraphBudget::new(source.state.source.limits())?;
    let source_graphs = native_footnotes_with_budget(source, &mut graph_budget)?;
    let candidate_graphs = native_footnotes_with_budget(candidate, &mut graph_budget)?;
    // The source and candidate are both retained until the complete semantic
    // readback below succeeds. Share one budget across their projections so
    // two individually-valid collections cannot exceed the retained-text
    // ceiling together. Each projection charges before it owns any footnote
    // strings, and a rejected charge leaves the shared allowance unchanged.
    let max_text_bytes = effective_text_limit(source.state.source.limits());
    let mut semantic_budget = super::FootnoteSemanticBudget::new(max_text_bytes);
    let before = super::project_body_footnotes_with_budget(
        source.state.source.components(),
        source.state.source.limits(),
        max_text_bytes,
        &mut semantic_budget,
    )
    .map_err(map_body_footnote_projection_error)?;
    let after = super::project_body_footnotes_with_budget(
        candidate.state.source.components(),
        candidate.state.source.limits(),
        max_text_bytes,
        &mut semantic_budget,
    )
    .map_err(map_body_footnote_projection_error)?;
    if before.len() != after.len() {
        return Err(Error::Verification);
    }
    for (old, new) in before.iter().zip(&after) {
        if old.position != new.position {
            return Err(Error::Verification);
        }
        let expected_text = if old.position == position {
            expected.text.as_ref()
        } else {
            old.text.as_ref()
        };
        let expected_mark = if old.position == position {
            expected.custom_mark.as_deref()
        } else {
            old.custom_mark.as_deref()
        };
        if new.text.as_ref() != expected_text || new.custom_mark.as_deref() != expected_mark {
            return Err(Error::Verification);
        }
    }
    if source.stats().total_objects() != candidate.stats().total_objects() {
        return Err(Error::Verification);
    }
    if source_graphs.len() != candidate_graphs.len()
        || source_graphs
            .iter()
            .zip(&candidate_graphs)
            .any(|(old, new)| {
                old.position != new.position
                    || old.reference_identifier != new.reference_identifier
                    || old.storage_identifier != new.storage_identifier
                    || old.marker_identifier != new.marker_identifier
            })
    {
        return Err(Error::Verification);
    }
    Ok(())
}

fn validate_text(text: &str) -> Result<(), Error> {
    if text.contains('\u{000e}') || text.contains('\u{fffc}') {
        return Err(Error::Refused(Refusal::StructuralMarker));
    }
    checked_utf16_units(text)?;
    if text.len() > crate::footnote::body::MAX_TEXT_BYTES {
        return Err(Error::Refused(Refusal::TextTooLarge));
    }
    Ok(())
}

fn checked_utf16_units(text: &str) -> Result<usize, Error> {
    let units = text.encode_utf16().count();
    let maximum = usize::try_from(u32::MAX).unwrap_or(usize::MAX);
    if units > maximum {
        return Err(Error::Limit {
            kind: LimitKind::TextUnits,
            observed: usize_to_u64(units),
            maximum: u64::from(u32::MAX),
        });
    }
    Ok(units)
}

fn position_from_anchor(anchor: u32) -> Result<Position, Error> {
    let index = usize::try_from(anchor).map_err(|_error| Error::Limit {
        kind: LimitKind::TextUnits,
        observed: u64::from(anchor),
        maximum: u64::from(u32::MAX),
    })?;
    Position::from_utf16_index(index).map_err(|_error| Error::InvalidSource)
}

fn validate_custom_mark(custom_mark: Option<&str>) -> Result<(), Error> {
    if custom_mark.is_some_and(|value| value.len() > crate::footnote::body::MAX_CUSTOM_MARK_BYTES) {
        return Err(Error::Refused(Refusal::CustomMarkTooLarge));
    }
    Ok(())
}

fn clone_footnote(value: &Footnote) -> Result<Footnote, Error> {
    Footnote::with_custom_mark(
        value.position,
        try_owned_text(&value.text)?.into_boxed_str(),
        value
            .custom_mark
            .as_deref()
            .map(try_owned_text)
            .transpose()?
            .map(String::into_boxed_str),
    )
    .map_err(map_footnote_value_error)
}

fn try_owned_text(text: &str) -> Result<String, Error> {
    let mut owned = String::new();
    owned
        .try_reserve_exact(text.len())
        .map_err(|_error| Error::Allocation { amount: text.len() })?;
    owned.push_str(text);
    Ok(owned)
}

fn map_footnote_value_error(error: crate::footnote::body::Error) -> Error {
    match error {
        crate::footnote::body::Error::TextTooLarge => Error::Refused(Refusal::TextTooLarge),
        crate::footnote::body::Error::CustomMarkTooLarge => {
            Error::Refused(Refusal::CustomMarkTooLarge)
        },
        crate::footnote::body::Error::PositionOutOfRange => Error::InvalidSource,
    }
}

fn map_wire_error(error: litchi_iwa_common::Error) -> Error {
    match error {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => Error::Limit {
            kind: match kind {
                litchi_iwa_common::LimitKind::InputBytes => LimitKind::InputBytes,
                litchi_iwa_common::LimitKind::Fields => LimitKind::WireFields,
                litchi_iwa_common::LimitKind::OutputBytes => LimitKind::OutputBytes,
                litchi_iwa_common::LimitKind::Nesting => LimitKind::WireNesting,
                litchi_iwa_common::LimitKind::RewriteWork => LimitKind::WireWork,
                litchi_iwa_common::LimitKind::TableRows
                | litchi_iwa_common::LimitKind::TableColumns
                | litchi_iwa_common::LimitKind::TableCells
                | litchi_iwa_common::LimitKind::MaterializedCells => LimitKind::Entries,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(limit),
        },
        litchi_iwa_common::Error::Allocation { amount, .. } => Error::Allocation { amount },
        _ => Error::InvalidSource,
    }
}

fn map_package_error(error: PackageError) -> Error {
    map_package_error_with_kind(error, LimitKind::WireBytes)
}

fn map_body_storage_decode_error(error: super::BodyStorageDecodeError) -> Error {
    match error {
        super::BodyStorageDecodeError::Package(error) => map_package_error(error),
        super::BodyStorageDecodeError::Wire(error) => map_text_rewrite_error(error),
        super::BodyStorageDecodeError::SemanticLimit { observed, limit } => Error::Limit {
            kind: LimitKind::TextBytes,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(limit),
        },
    }
}

fn map_body_footnote_projection_error(error: super::BodyStorageDecodeError) -> Error {
    match error {
        super::BodyStorageDecodeError::SemanticLimit { observed, limit } => Error::Limit {
            kind: LimitKind::TextBytes,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(limit),
        },
        super::BodyStorageDecodeError::Package(PackageError::PayloadLimit { observed, limit }) => {
            Error::Limit {
                kind: LimitKind::Entries,
                observed: usize_to_u64(observed),
                maximum: usize_to_u64(limit),
            }
        },
        super::BodyStorageDecodeError::Package(error) => map_package_error(error),
        super::BodyStorageDecodeError::Wire(error) => map_text_rewrite_error(error),
    }
}

fn map_package_error_with_kind(error: PackageError, payload_limit_kind: LimitKind) -> Error {
    match error {
        PackageError::Archive(error) => map_archive_error(error),
        PackageError::PayloadLimit { observed, limit } => Error::Limit {
            kind: payload_limit_kind,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(limit),
        },
        PackageError::ObjectLimit { observed, limit } => Error::Limit {
            kind: LimitKind::Entries,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(limit),
        },
        PackageError::Allocation { amount } => Error::Allocation { amount },
        PackageError::SectionNamesTooLarge { observed, limit } => Error::Limit {
            kind: LimitKind::TextBytes,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(limit),
        },
        PackageError::Semantic(crate::Error::TextTooLarge { observed, limit }) => Error::Limit {
            kind: LimitKind::TextBytes,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(limit),
        },
        PackageError::Semantic(
            crate::Error::TooManySections { actual, limit }
            | crate::Error::TooManyBodyStorages { actual, limit },
        ) => Error::Limit {
            kind: LimitKind::Entries,
            observed: usize_to_u64(actual),
            maximum: usize_to_u64(limit),
        },
        PackageError::Io(_)
        | PackageError::Detection(_)
        | PackageError::UnexpectedFormat { .. }
        | PackageError::NotOfficeFile
        | PackageError::MissingIdentity
        | PackageError::InvalidLanguage(_)
        | PackageError::InvalidLocale
        | PackageError::InvalidLimit { .. }
        | PackageError::Limit { .. }
        | PackageError::InvalidFormat(_)
        | PackageError::Semantic(crate::Error::InvalidSectionIndex { .. }) => Error::InvalidSource,
    }
}

fn map_storage_wire_limits_error(error: super::StorageWireLimitsError) -> Error {
    match error {
        super::StorageWireLimitsError::Physical(error) => map_archive_error(error),
        super::StorageWireLimitsError::Wire(RewriteError::InvalidLimit {
            field: "nesting",
            value,
            ..
        }) => Error::Limit {
            kind: LimitKind::WireNesting,
            observed: 4,
            maximum: value as u64,
        },
        super::StorageWireLimitsError::Wire(error) => map_text_rewrite_error(error),
    }
}

fn map_text_rewrite_error(error: RewriteError) -> Error {
    match error {
        RewriteError::LimitExceeded {
            resource,
            observed,
            limit,
        } => Error::Limit {
            kind: text_limit_kind(resource),
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(limit),
        },
        RewriteError::InvalidLimit {
            field,
            value,
            maximum,
        } => Error::Limit {
            kind: text_limit_kind(field),
            observed: usize_to_u64(value),
            maximum: usize_to_u64(maximum),
        },
        RewriteError::Allocation { amount, .. } => Error::Allocation { amount },
        _ => Error::InvalidSource,
    }
}

fn text_limit_kind(resource: &str) -> LimitKind {
    match resource {
        "text bytes" => LimitKind::TextBytes,
        "text fragments" | "table entries" | "object references" => LimitKind::Entries,
        "nesting" => LimitKind::WireNesting,
        "fields" => LimitKind::WireFields,
        "rewrite work" | "aggregate nested scan bytes" => LimitKind::WireWork,
        _ => LimitKind::WireBytes,
    }
}

fn map_archive_error(error: litchi_iwa_archive::Error) -> Error {
    match error {
        litchi_iwa_archive::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                litchi_iwa_archive::LimitKind::InputBytes => LimitKind::InputBytes,
                litchi_iwa_archive::LimitKind::OutputBytes => LimitKind::OutputBytes,
                litchi_iwa_archive::LimitKind::Entries => LimitKind::Entries,
                litchi_iwa_archive::LimitKind::MemberNameBytes
                | litchi_iwa_archive::LimitKind::MetadataBytes
                | litchi_iwa_archive::LimitKind::CompressedEntryBytes
                | litchi_iwa_archive::LimitKind::EntryBytes
                | litchi_iwa_archive::LimitKind::IwaStreamBytes => LimitKind::EntryBytes,
                litchi_iwa_archive::LimitKind::TotalBytes
                | litchi_iwa_archive::LimitKind::IwaTotalBytes => LimitKind::TotalBytes,
            },
            observed,
            maximum,
        },
        litchi_iwa_archive::Error::Allocation { amount, .. } => Error::Allocation { amount },
        litchi_iwa_archive::Error::Iwa(error) => map_core_error(error),
        _ => Error::InvalidSource,
    }
}

fn map_core_error(error: litchi_iwa_core::Error) -> Error {
    match error {
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                litchi_iwa_core::LimitKind::Objects
                | litchi_iwa_core::LimitKind::Messages
                | litchi_iwa_core::LimitKind::MessagesPerObject
                | litchi_iwa_core::LimitKind::MetadataItems => LimitKind::Entries,
                litchi_iwa_core::LimitKind::MessageBytes => LimitKind::WireBytes,
                litchi_iwa_core::LimitKind::HeaderFields => LimitKind::WireFields,
                litchi_iwa_core::LimitKind::HeaderNesting => LimitKind::WireNesting,
                litchi_iwa_core::LimitKind::ArchiveBytes
                | litchi_iwa_core::LimitKind::ObjectBytes
                | litchi_iwa_core::LimitKind::HeaderBytes
                | litchi_iwa_core::LimitKind::HeaderMemoryBytes
                | litchi_iwa_core::LimitKind::SnappyChunkBytes
                | litchi_iwa_core::LimitKind::SnappyStreamBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedChunkBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedStreamBytes
                | litchi_iwa_core::LimitKind::SnappyFrames => LimitKind::EntryBytes,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            Error::Allocation { amount: requested }
        },
        _ => Error::InvalidSource,
    }
}

fn usize_to_u64(value: usize) -> u64 {
    u64::try_from(value).unwrap_or(u64::MAX)
}

#[cfg(test)]
mod tests {
    use std::collections::HashMap;
    use std::sync::Arc;

    use super::{
        Error, FootnoteGraphBudget, FootnoteObjectLocations, LimitKind, Patch, checked_utf16_units,
        is_canonical_component_name, map_body_footnote_projection_error,
        map_body_storage_decode_error, map_package_error_with_kind, map_storage_wire_limits_error,
        position_from_anchor, rewrite_custom_mark_wire,
    };
    use crate::footnote::body::{Footnote, Position};
    use crate::package::{BodyStorageDecodeError, StorageWireLimitsError};
    use crate::{Error as SemanticError, PackageError};
    use litchi_iwa_common::WireLimits;
    use litchi_iwa_core::ArchiveObject;
    use litchi_iwa_text_wire::RewriteError;

    #[test]
    fn mutation_authority_accepts_only_canonical_index_components() {
        assert!(is_canonical_component_name("Index/Document.iwa"));
        assert!(is_canonical_component_name("Index/Footnotes.iwa"));
        for alias in [
            "Document.iwa",
            "Index//Document.iwa",
            "Index/./Document.iwa",
            "Index\\Document.iwa",
            "Index/dir/Document.iwa",
            "Index/Document.iwa/",
        ] {
            assert!(!is_canonical_component_name(alias), "unsafe alias: {alias}");
        }
    }

    #[test]
    fn anchors_and_text_are_checked_in_typed_utf16_units() {
        assert_eq!(position_from_anchor(7).unwrap().utf16_index(), 7);
        assert_eq!(checked_utf16_units("😀"), Ok(2));
        assert!(matches!(
            position_from_anchor(u32::MAX),
            Ok(position) if position.utf16_index() == u32::MAX
        ));
        assert!(!matches!(
            checked_utf16_units("native"),
            Err(Error::Limit {
                kind: LimitKind::TextUnits,
                ..
            })
        ));
    }

    #[test]
    fn custom_marker_wire_edit_round_trips_exactly() {
        let source = [0x08, 0x02, 0x1a, 0x01, b'x', 0xa0, 0x06, 0x01];
        let limits = WireLimits::default();
        let forward = rewrite_custom_mark_wire(&source, Some("x"), Some("y"), limits)
            .unwrap_or_else(|error| panic!("forward marker edit: {error}"));
        let backward = rewrite_custom_mark_wire(&forward, Some("y"), Some("x"), limits)
            .unwrap_or_else(|error| panic!("inverse marker edit: {error}"));
        assert_eq!(backward, source);
        assert_eq!(&forward[..2], &source[..2]);
        assert_eq!(&forward[5..], &source[5..]);
    }

    #[test]
    fn semantic_package_limits_keep_typed_footnote_error_categories() {
        let text_error = map_package_error_with_kind(
            PackageError::Semantic(SemanticError::TextTooLarge {
                observed: 9,
                limit: 8,
            }),
            LimitKind::WireBytes,
        );
        assert_eq!(
            text_error,
            Error::Limit {
                kind: LimitKind::TextBytes,
                observed: 9,
                maximum: 8,
            }
        );

        let entries_error = map_package_error_with_kind(
            PackageError::Semantic(SemanticError::TooManyBodyStorages {
                actual: 5,
                limit: 4,
            }),
            LimitKind::WireBytes,
        );
        assert_eq!(
            entries_error,
            Error::Limit {
                kind: LimitKind::Entries,
                observed: 5,
                maximum: 4,
            }
        );

        let aggregate_error = map_package_error_with_kind(
            PackageError::PayloadLimit {
                observed: 19,
                limit: 14,
            },
            LimitKind::TextBytes,
        );
        assert_eq!(
            aggregate_error,
            Error::Limit {
                kind: LimitKind::TextBytes,
                observed: 19,
                maximum: 14,
            }
        );
    }

    #[test]
    fn graph_budget_charges_source_references_and_work_without_partial_consumption() {
        let mut budget = FootnoteGraphBudget::with_test_limits(8, 5, 7);
        budget.charge_source(8).unwrap();
        assert!(matches!(
            budget.charge_source(1),
            Err(Error::Limit {
                kind: LimitKind::InputBytes,
                observed: 9,
                maximum: 8,
            })
        ));
        budget.charge_references(5).unwrap();
        assert!(matches!(
            budget.charge_references(1),
            Err(Error::Limit {
                kind: LimitKind::Entries,
                observed: 6,
                maximum: 5,
            })
        ));
        budget.charge_work(7).unwrap();
        assert!(matches!(
            budget.charge_work(1),
            Err(Error::Limit {
                kind: LimitKind::WireWork,
                observed: 8,
                maximum: 7,
            })
        ));
        assert_eq!(budget.source_bytes, 8);
        assert_eq!(budget.references, 5);
        assert_eq!(budget.work, 7);
    }

    #[test]
    fn body_footnote_projection_keeps_entry_and_wire_limit_categories() {
        let entries = map_body_footnote_projection_error(BodyStorageDecodeError::Package(
            PackageError::PayloadLimit {
                observed: 4097,
                limit: 4096,
            },
        ));
        assert_eq!(
            entries,
            Error::Limit {
                kind: LimitKind::Entries,
                observed: 4097,
                maximum: 4096,
            }
        );

        let aggregate = map_body_footnote_projection_error(BodyStorageDecodeError::SemanticLimit {
            observed: 19,
            limit: 14,
        });
        assert_eq!(
            aggregate,
            Error::Limit {
                kind: LimitKind::TextBytes,
                observed: 19,
                maximum: 14,
            }
        );

        let text = map_body_footnote_projection_error(BodyStorageDecodeError::Package(
            PackageError::Semantic(SemanticError::TextTooLarge {
                observed: 9,
                limit: 8,
            }),
        ));
        assert_eq!(
            text,
            Error::Limit {
                kind: LimitKind::TextBytes,
                observed: 9,
                maximum: 8,
            }
        );

        let wire = map_body_footnote_projection_error(BodyStorageDecodeError::Wire(
            RewriteError::LimitExceeded {
                resource: "rewrite work",
                observed: 33,
                limit: 32,
            },
        ));
        assert_eq!(
            wire,
            Error::Limit {
                kind: LimitKind::WireWork,
                observed: 33,
                maximum: 32,
            }
        );
    }

    #[test]
    fn native_body_storage_limits_keep_wire_categories() {
        let cases = [
            (
                RewriteError::LimitExceeded {
                    resource: "message bytes",
                    observed: 17,
                    limit: 16,
                },
                LimitKind::WireBytes,
                17,
                16,
            ),
            (
                RewriteError::LimitExceeded {
                    resource: "text fragments",
                    observed: 9,
                    limit: 8,
                },
                LimitKind::Entries,
                9,
                8,
            ),
            (
                RewriteError::LimitExceeded {
                    resource: "rewrite work",
                    observed: 33,
                    limit: 32,
                },
                LimitKind::WireWork,
                33,
                32,
            ),
        ];
        for (error, kind, observed, maximum) in cases {
            assert_eq!(
                map_body_storage_decode_error(BodyStorageDecodeError::Wire(error)),
                Error::Limit {
                    kind,
                    observed,
                    maximum,
                }
            );
        }
        assert_eq!(
            map_body_storage_decode_error(BodyStorageDecodeError::Package(
                PackageError::PayloadLimit {
                    observed: 5,
                    limit: 4,
                },
            )),
            Error::Limit {
                kind: LimitKind::WireBytes,
                observed: 5,
                maximum: 4,
            }
        );
    }

    #[test]
    fn storage_wire_limits_keep_typed_rewrite_categories() {
        let exceeded = map_storage_wire_limits_error(StorageWireLimitsError::Wire(
            RewriteError::LimitExceeded {
                resource: "rewrite work",
                observed: 9,
                limit: 8,
            },
        ));
        assert_eq!(
            exceeded,
            Error::Limit {
                kind: LimitKind::WireWork,
                observed: 9,
                maximum: 8,
            }
        );

        let invalid = map_storage_wire_limits_error(StorageWireLimitsError::Wire(
            RewriteError::InvalidLimit {
                field: "fields",
                value: 0,
                maximum: 1_000,
            },
        ));
        assert_eq!(
            invalid,
            Error::Limit {
                kind: LimitKind::WireFields,
                observed: 0,
                maximum: 1_000,
            }
        );

        let text = map_storage_wire_limits_error(StorageWireLimitsError::Wire(
            RewriteError::LimitExceeded {
                resource: "text bytes",
                observed: 17,
                limit: 16,
            },
        ));
        assert_eq!(
            text,
            Error::Limit {
                kind: LimitKind::TextBytes,
                observed: 17,
                maximum: 16,
            }
        );

        let entries = map_storage_wire_limits_error(StorageWireLimitsError::Wire(
            RewriteError::LimitExceeded {
                resource: "table entries",
                observed: 33,
                limit: 32,
            },
        ));
        assert_eq!(
            entries,
            Error::Limit {
                kind: LimitKind::Entries,
                observed: 33,
                maximum: 32,
            }
        );
    }

    #[test]
    fn borrowed_locations_keep_first_ordered_match_and_isolate_unrelated_ids() {
        let first = ArchiveObject::new(7, Vec::new(), Default::default())
            .unwrap_or_else(|error| panic!("first object: {error}"));
        let duplicate = ArchiveObject::new(7, Vec::new(), Default::default())
            .unwrap_or_else(|error| panic!("duplicate object: {error}"));
        let unrelated = ArchiveObject::new(8, Vec::new(), Default::default())
            .unwrap_or_else(|error| panic!("unrelated object: {error}"));
        let mut objects = HashMap::new();
        objects
            .try_reserve(3)
            .unwrap_or_else(|error| panic!("location reservation: {error}"));
        objects.entry(7).or_insert(&first);
        objects.entry(7).or_insert(&duplicate);
        objects.entry(8).or_insert(&unrelated);
        let locations = FootnoteObjectLocations { objects };

        assert!(std::ptr::eq(
            locations.get(7).unwrap_or_else(|| panic!("first location")),
            &first
        ));
        assert!(std::ptr::eq(
            locations
                .get(8)
                .unwrap_or_else(|| panic!("unrelated location")),
            &unrelated
        ));
        assert!(locations.get(9).is_none());
    }

    #[test]
    fn inverse_and_exact_noop_preserve_patch_bytes() {
        let before = Footnote::new(Position::ZERO, "before")
            .unwrap_or_else(|error| panic!("before footnote: {error}"));
        let after = Footnote::new(Position::ZERO, "after")
            .unwrap_or_else(|error| panic!("after footnote: {error}"));
        let source_bytes: Arc<[u8]> = Arc::from(&b"source"[..]);
        let target_bytes: Arc<[u8]> = Arc::from(&b"target"[..]);
        let patch = Patch {
            source_bytes: Arc::clone(&source_bytes),
            target_bytes: Arc::clone(&target_bytes),
            source_fingerprint: 1,
            target_fingerprint: 2,
            position: Position::ZERO,
            before: before.clone(),
            after: after.clone(),
        };
        assert_eq!(patch.inverse().inverse(), patch);
        assert!(!patch.is_noop());

        let noop = Patch {
            source_bytes: Arc::clone(&source_bytes),
            target_bytes: source_bytes,
            source_fingerprint: 1,
            target_fingerprint: 1,
            position: Position::ZERO,
            before: before.clone(),
            after: before,
        };
        assert!(noop.is_noop());
        assert_eq!(noop.inverse().inverse(), noop);
    }
}
