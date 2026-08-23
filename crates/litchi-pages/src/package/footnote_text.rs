//! Exact-source transactions for Pages body-footnote text.
//!
//! This module deliberately owns only replacement of the user-visible text
//! inside an existing footnote storage and the optional custom marker string
//! on its existing reference attachment. The body anchor, body footnote table,
//! native reference/storage/marker edges, and all object ownership metadata
//! remain untouched. Insert/remove therefore stay in the migration host until
//! their graph lifecycle can be moved without weakening cleanup.

use std::fmt;
use std::num::NonZeroU64;
use std::sync::Arc;

use litchi_iwa_archive::{SourceCatalog, package::EntryEdit};
use litchi_iwa_common::{WireLimits, encode_varint_into, varint::encoded_len, wire::WireView};
use litchi_iwa_core::{Archive, RawMessage, SnappyStream};
use litchi_iwa_protos::pages_footnote_codec;
use litchi_iwa_text_wire::RewriteError;
use thiserror::Error;

use super::{
    FOOTNOTE_REFERENCE_MESSAGE_TYPE, MAX_BODY_FOOTNOTES, Package, PackageError,
    STORAGE_TEXT_PREFIX, TEXTUAL_ATTACHMENT_MESSAGE_TYPE, decode_body_storage,
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
pub enum FootnoteTextLimitKind {
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

impl fmt::Display for FootnoteTextLimitKind {
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
#[derive(Debug, Clone, Copy, PartialEq, Eq, Error)]
#[non_exhaustive]
pub enum FootnoteTextError {
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
    /// The selected native footnote cannot be rewritten without ambiguity.
    #[error("the Pages source cannot be edited safely")]
    InvalidSource,
    /// A finite transaction resource ceiling was exceeded.
    #[error("Pages footnote-text {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    LimitExceeded {
        /// Resource category that exceeded its limit.
        kind: FootnoteTextLimitKind,
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

/// One mutable body-footnote text and custom-marker value staged against an
/// immutable package.
pub struct FootnoteTextEdit<'a> {
    source: &'a Package,
    position: Position,
    before: Footnote,
    text: String,
    custom_mark: Option<String>,
}

impl fmt::Debug for FootnoteTextEdit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("FootnoteTextEdit")
            .field("position", &self.position)
            .finish_non_exhaustive()
    }
}

impl<'a> FootnoteTextEdit<'a> {
    fn new(source: &'a Package, selector: Selector) -> Result<Self, FootnoteTextError> {
        let (position, before, _) = resolve_footnote(source, selector)?;
        let text = try_owned_text(&before.text)?;
        let custom_mark = before
            .custom_mark
            .as_deref()
            .map(try_owned_text)
            .transpose()?;
        Ok(Self {
            source,
            position,
            before,
            text,
            custom_mark,
        })
    }

    /// Return the checked UTF-16 body-anchor position resolved at edit start.
    #[must_use]
    pub const fn position(&self) -> Position {
        self.position
    }

    /// Borrow the original semantic footnote value.
    #[must_use]
    pub const fn before(&self) -> &Footnote {
        &self.before
    }

    /// Borrow the text currently staged for publication.
    #[must_use]
    pub fn text(&self) -> &str {
        &self.text
    }

    /// Borrow the custom marker currently staged for publication.
    #[must_use]
    pub fn custom_mark(&self) -> Option<&str> {
        self.custom_mark.as_deref()
    }

    /// Stage replacement of the selected footnote's user-visible text.
    ///
    /// The native body anchor and footnote graph are not changed. Structural
    /// U+000E and U+FFFC markers remain reserved for native graph edges.
    pub fn set(&mut self, text: &str) -> Result<&mut Self, FootnoteTextError> {
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
    pub fn set_custom_mark(
        &mut self,
        custom_mark: Option<&str>,
    ) -> Result<&mut Self, FootnoteTextError> {
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
    pub fn commit(self) -> Result<FootnoteTextCommit, FootnoteTextError> {
        validate_text(&self.text)?;
        validate_custom_mark(self.custom_mark.as_deref())?;
        self.source.validate().map_err(map_package_error)?;
        let (position, current, native_footnote) =
            resolve_footnote(self.source, Selector::At(self.position))?;
        if position != self.position || current != self.before {
            return Err(FootnoteTextError::InvalidSource);
        }
        if current.text.as_ref() != self.before.text.as_ref() {
            return Err(FootnoteTextError::InvalidSource);
        }
        if current.custom_mark != self.before.custom_mark {
            return Err(FootnoteTextError::InvalidSource);
        }

        let source_bytes = self.source.state.source.shared_source();
        let source_fingerprint = super::section_transaction::fingerprint(&source_bytes);
        let before = clone_footnote(&self.before)?;
        let after = Footnote::with_custom_mark(
            self.position,
            self.text.clone().into_boxed_str(),
            self.custom_mark.clone().map(String::into_boxed_str),
        )
        .map_err(map_footnote_value_error)?;

        if before == after {
            return Ok(FootnoteTextCommit {
                package: self.source.snapshot(),
                patch: FootnoteTextPatch {
                    source_bytes: Arc::clone(&source_bytes),
                    target_bytes: source_bytes,
                    source_fingerprint,
                    target_fingerprint: source_fingerprint,
                    position: self.position,
                    before,
                    after,
                },
                diagnostics: FootnoteTextDiagnostics::unchanged(),
            });
        }
        if !self.source.state.source.source_is_exact() {
            return Err(FootnoteTextError::UnsupportedSource);
        }

        let package = rewrite_package_footnote(
            self.source,
            native_footnote,
            self.before.text.as_ref(),
            self.text.as_str(),
            self.before.custom_mark.as_deref(),
            self.custom_mark.as_deref(),
            self.position,
        )?;
        let touched_components = changed_component_count(self.source, &package)?;
        if touched_components == 0 {
            return Err(FootnoteTextError::Verification);
        }
        let target_bytes = package.state.source.shared_source();
        let target_fingerprint = super::section_transaction::fingerprint(&target_bytes);
        Ok(FootnoteTextCommit {
            package,
            patch: FootnoteTextPatch {
                source_bytes,
                target_bytes,
                source_fingerprint,
                target_fingerprint,
                position: self.position,
                before,
                after,
            },
            diagnostics: FootnoteTextDiagnostics::published(touched_components),
        })
    }
}

/// An exact-source-checked, reversible Pages body-footnote text patch.
///
/// Native identifiers, member names, and exact source/target bytes remain
/// private. Fingerprints are diagnostics only; exact byte comparison
/// authorizes application.
#[derive(Clone, PartialEq, Eq)]
pub struct FootnoteTextPatch {
    source_bytes: Arc<[u8]>,
    target_bytes: Arc<[u8]>,
    source_fingerprint: u64,
    target_fingerprint: u64,
    position: Position,
    before: Footnote,
    after: Footnote,
}

impl fmt::Debug for FootnoteTextPatch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("FootnoteTextPatch")
            .field("position", &self.position)
            .finish_non_exhaustive()
    }
}

impl FootnoteTextPatch {
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
pub struct FootnoteTextDiagnostics {
    changed: bool,
    touched_components: usize,
    full_reparse_performed: bool,
}

impl FootnoteTextDiagnostics {
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
pub struct FootnoteTextCommit {
    package: Package,
    patch: FootnoteTextPatch,
    diagnostics: FootnoteTextDiagnostics,
}

impl FootnoteTextCommit {
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

    /// Borrow the reversible exact-source patch.
    #[must_use]
    pub const fn patch(&self) -> &FootnoteTextPatch {
        &self.patch
    }

    /// Borrow compact publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &FootnoteTextDiagnostics {
        &self.diagnostics
    }
}

impl Package {
    /// Read every semantic body footnote through the immutable package owner.
    ///
    /// This alias keeps mutation call sites close to the setter transaction;
    /// ordinary reads remain available through [`Package::body_footnotes`].
    pub fn edit_body_footnote_text(
        &self,
        selector: Selector,
    ) -> Result<FootnoteTextEdit<'_>, FootnoteTextError> {
        FootnoteTextEdit::new(self, selector)
    }

    /// Apply an exact-source-checked body-footnote text patch.
    ///
    /// The retained target is fully reopened and semantically verified under
    /// this package's original limits before it is published.
    pub fn apply_body_footnote_text(
        &self,
        patch: &FootnoteTextPatch,
    ) -> Result<FootnoteTextCommit, FootnoteTextError> {
        let source = &self.state.source;
        let source_bytes = source.shared_source();
        if super::section_transaction::fingerprint(source.source_bytes())
            != patch.source_fingerprint
            || source.source_bytes() != patch.source_bytes.as_ref()
            || source_bytes.as_ref() != patch.source_bytes.as_ref()
        {
            return Err(FootnoteTextError::PatchConflict);
        }
        self.validate().map_err(map_package_error)?;
        let (_, current, _) = resolve_footnote(self, Selector::At(patch.position))?;
        if current != patch.before {
            return Err(FootnoteTextError::PatchConflict);
        }
        if patch.is_noop() {
            return Ok(FootnoteTextCommit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: FootnoteTextDiagnostics::unchanged(),
            });
        }
        if !source.source_is_exact()
            || super::section_transaction::fingerprint(&patch.target_bytes)
                != patch.target_fingerprint
        {
            return Err(FootnoteTextError::PatchConflict);
        }
        let candidate_source = SourceCatalog::from_shared_bytes_with_limits(
            Arc::clone(&patch.target_bytes),
            source.limits(),
        )
        .map_err(map_archive_error)?;
        let candidate =
            Package::from_source_catalog(candidate_source).map_err(map_package_error)?;
        candidate.validate().map_err(map_package_error)?;
        let touched_components = changed_component_count(self, &candidate)?;
        if touched_components == 0 {
            return Err(FootnoteTextError::Verification);
        }
        verify_candidate(self, &candidate, patch.position, &patch.after)?;
        Ok(FootnoteTextCommit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: FootnoteTextDiagnostics::published(touched_components),
        })
    }
}

#[derive(Debug, Clone, Copy)]
struct NativeFootnote {
    body_identifier: NonZeroU64,
    position: Position,
    reference_identifier: NonZeroU64,
    storage_identifier: NonZeroU64,
    marker_identifier: NonZeroU64,
}

fn resolve_footnote(
    package: &Package,
    selector: Selector,
) -> Result<(Position, Footnote, NativeFootnote), FootnoteTextError> {
    let graphs = native_footnotes(package)?;
    let index = match selector {
        Selector::Index(index) => index,
        Selector::At(position) => graphs
            .iter()
            .position(|graph| graph.position == position)
            .ok_or(FootnoteTextError::NotFound)?,
    };
    let graph = graphs.get(index).ok_or(FootnoteTextError::NotFound)?;
    let footnotes = package.body_footnotes().map_err(map_package_error)?;
    let footnote = footnotes
        .get(index)
        .ok_or(FootnoteTextError::InvalidSource)?;
    Ok((graph.position, clone_footnote(footnote)?, *graph))
}

fn native_footnotes(package: &Package) -> Result<Vec<NativeFootnote>, FootnoteTextError> {
    let mut semantic_budget =
        super::FootnoteSemanticBudget::new(effective_text_limit(package.state.source.limits()));
    let root = root_references_with_limits(
        package.state.source.components(),
        package.state.source.limits(),
    )
    .map_err(map_package_error)?;
    let Some(body_identifier) = root.body else {
        return Ok(Vec::new());
    };
    let body = find_object(package.state.source.components(), body_identifier.get())
        .ok_or(FootnoteTextError::InvalidSource)?;
    let body_payload =
        unique_text_payload(&body.messages, body_identifier).map_err(map_package_error)?;
    let (body_storage, _) = decode_body_storage(
        &body.messages,
        body_identifier,
        super::MAX_SECTIONS,
        effective_text_limit(package.state.source.limits()),
        package.state.source.limits(),
    )
    .map_err(map_package_error)?;
    checked_utf16_units(body_storage.text())?;
    let entries =
        super::footnote_table_entries(body_payload, body_identifier, package.state.source.limits())
            .map_err(|error| map_package_error_with_kind(error, FootnoteTextLimitKind::Entries))?;
    if entries.len() > MAX_BODY_FOOTNOTES {
        return Err(FootnoteTextError::LimitExceeded {
            kind: FootnoteTextLimitKind::Entries,
            observed: usize_to_u64(entries.len()),
            maximum: usize_to_u64(MAX_BODY_FOOTNOTES),
        });
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(entries.len())
        .map_err(|_error| FootnoteTextError::Allocation {
            amount: entries.len(),
        })?;
    let mut seen_references = Vec::new();
    let mut seen_storages = Vec::new();
    let mut seen_markers = Vec::new();
    for seen in [&mut seen_references, &mut seen_storages, &mut seen_markers] {
        seen.try_reserve_exact(entries.len())
            .map_err(|_error| FootnoteTextError::Allocation {
                amount: entries.len(),
            })?;
    }
    let mut previous = None;
    for entry in entries {
        if previous.is_some_and(|position| position >= entry.character_index) {
            return Err(FootnoteTextError::InvalidSource);
        }
        previous = Some(entry.character_index);
        if seen_references.contains(&entry.identifier) {
            return Err(FootnoteTextError::InvalidSource);
        }
        seen_references.push(entry.identifier);
        super::validate_body_footnote_anchor(
            body_storage.text(),
            body_identifier,
            entry.character_index,
        )
        .map_err(map_package_error)?;
        let reference = find_object(package.state.source.components(), entry.identifier.get())
            .ok_or(FootnoteTextError::InvalidSource)?;
        let payload = unique_message_payload(
            &reference.messages,
            FOOTNOTE_REFERENCE_MESSAGE_TYPE,
            "Pages footnote reference",
        )
        .map_err(map_package_error)?;
        let decoded = pages_footnote_codec::decode_footnote_reference(
            payload,
            footnote_decode_options(payload, package.state.source.limits())
                .map_err(map_package_error)?,
        )
        .map_err(|_error| FootnoteTextError::InvalidSource)?;
        if decoded
            .super_kind()
            .is_some_and(|kind| kind != FOOTNOTE_MARK_KIND)
        {
            return Err(FootnoteTextError::InvalidSource);
        }
        let storage_identifier = decoded
            .contained_storage()
            .map(|reference| reference.identifier())
            .ok_or(FootnoteTextError::InvalidSource)?;
        if storage_identifier == entry.identifier || seen_storages.contains(&storage_identifier) {
            return Err(FootnoteTextError::InvalidSource);
        }
        seen_storages.push(storage_identifier);
        let storage = find_object(package.state.source.components(), storage_identifier.get())
            .ok_or(FootnoteTextError::InvalidSource)?;
        let storage_payload = unique_text_payload(&storage.messages, storage_identifier)
            .map_err(map_package_error)?;
        let (storage_value, _) = decode_body_storage(
            &storage.messages,
            storage_identifier,
            super::MAX_SECTIONS,
            effective_text_limit(package.state.source.limits()),
            package.state.source.limits(),
        )
        .map_err(map_package_error)?;
        let text = storage_value
            .text()
            .strip_prefix(STORAGE_TEXT_PREFIX)
            .ok_or(FootnoteTextError::InvalidSource)?;
        if text.len() > crate::footnote::body::MAX_TEXT_BYTES {
            return Err(FootnoteTextError::InvalidSource);
        }
        let custom_mark = decoded.custom_mark_string();
        if custom_mark
            .is_some_and(|value| value.len() > crate::footnote::body::MAX_CUSTOM_MARK_BYTES)
        {
            return Err(FootnoteTextError::InvalidSource);
        }
        super::FootnoteSemanticBudget::charge(
            &mut semantic_budget,
            text.len(),
            custom_mark.map_or(0, str::len),
        )
        .map_err(|error| map_package_error_with_kind(error, FootnoteTextLimitKind::TextBytes))?;
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
            return Err(FootnoteTextError::InvalidSource);
        }
        seen_markers.push(marker_identifier);
        let marker = find_object(package.state.source.components(), marker_identifier.get())
            .ok_or(FootnoteTextError::InvalidSource)?;
        let marker_payload = unique_message_payload(
            &marker.messages,
            TEXTUAL_ATTACHMENT_MESSAGE_TYPE,
            "Pages footnote marker",
        )
        .map_err(map_package_error)?;
        let marker_value =
            litchi_iwa_protos::pages_footnote_marker_codec::decode_textual_attachment(
                marker_payload,
                footnote_marker_decode_options(marker_payload, package.state.source.limits())
                    .map_err(map_package_error)?,
            )
            .map_err(|_error| FootnoteTextError::InvalidSource)?;
        if marker_value.kind() != Some(FOOTNOTE_MARK_KIND) {
            return Err(FootnoteTextError::InvalidSource);
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
) -> Result<Package, FootnoteTextError> {
    let source_catalog = &source.state.source;
    let text_component = (before_text != after_text)
        .then(|| component_name_for_object(source, native_footnote.storage_identifier))
        .transpose()?;
    let mark_component = (before_custom_mark != after_custom_mark)
        .then(|| component_name_for_object(source, native_footnote.reference_identifier))
        .transpose()?;
    let first_component = text_component.or(mark_component);
    let first_component = first_component.ok_or(FootnoteTextError::InvalidSource)?;
    let second_component = match (text_component, mark_component) {
        (Some(text), Some(mark)) if text != mark => Some(mark),
        _ => None,
    };
    // Resolve every dependency and physical authority before an in-memory
    // archive is edited.  A semantic footnote value alone cannot prove that
    // its storage/reference/marker objects are exclusively owned by this
    // body-footnote graph, nor that the selected ZIP members are canonical
    // exact authorities.
    prove_footnote_ownership(source, native_footnote)?;
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
        .map_err(|_error| FootnoteTextError::Allocation {
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
        .map_err(|_error| FootnoteTextError::Allocation {
            amount: rewritten_entries.len(),
        })?;
    for (component_name, compressed) in &rewritten_entries {
        edits.push(EntryEdit::new(component_name, compressed));
    }
    let output = source_catalog
        .package()
        .reassemble_to_bytes(&edits, physical_limits)
        .map_err(map_archive_error)?;
    let candidate_source =
        SourceCatalog::from_shared_bytes_with_limits(output.into(), physical_limits)
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

fn component_name_for_object(
    package: &Package,
    identifier: NonZeroU64,
) -> Result<&str, FootnoteTextError> {
    let mut matching = package
        .state
        .source
        .components()
        .iter()
        .filter(|component| component.archive().object(identifier.get()).is_some());
    let component = matching.next().ok_or(FootnoteTextError::InvalidSource)?;
    if matching.next().is_some() {
        return Err(FootnoteTextError::InvalidSource);
    }
    if package
        .state
        .source
        .package()
        .iter()
        .find(|entry| entry.name() == component.name())
        .is_some_and(|entry| entry.is_opaque())
    {
        return Err(FootnoteTextError::InvalidSource);
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
fn validate_mutation_component(
    package: &Package,
    component_name: &str,
) -> Result<(), FootnoteTextError> {
    if !is_canonical_component_name(component_name) {
        return Err(FootnoteTextError::InvalidSource);
    }
    let entry = package
        .state
        .source
        .package()
        .iter()
        .find(|entry| entry.name() == component_name)
        .ok_or(FootnoteTextError::InvalidSource)?;
    if entry.is_opaque() {
        return Err(FootnoteTextError::InvalidSource);
    }
    let expected = component_name.as_bytes();
    if entry.raw_name() != expected
        || entry.metadata().local().name() != expected
        || entry.metadata().central().name() != expected
        || entry.metadata().local().compression_method()
            != entry.metadata().central().compression_method()
    {
        return Err(FootnoteTextError::InvalidSource);
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
fn prove_footnote_ownership(
    package: &Package,
    native_footnote: NativeFootnote,
) -> Result<(), FootnoteTextError> {
    let limits = package.state.source.limits();
    for component in package.state.source.components().iter() {
        for object in &component.archive().objects {
            let owner = object
                .archive_info
                .identifier
                .and_then(NonZeroU64::new)
                .ok_or(FootnoteTextError::InvalidSource)?;

            for message_info in &object.archive_info.message_infos {
                for referenced in &message_info.object_references {
                    validate_graph_edge(owner.get(), *referenced, native_footnote)?;
                }
                for field_info in &message_info.field_infos {
                    for referenced in &field_info.object_references {
                        validate_graph_edge(owner.get(), *referenced, native_footnote)?;
                    }
                }
            }

            for message in &object.messages {
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
                            return Err(FootnoteTextError::InvalidSource);
                        },
                        Err(_error) => {},
                    }
                }

                if !is_body_text_message_type(message.type_) {
                    continue;
                }
                let Some(owner_identifier) = NonZeroU64::new(owner.get()) else {
                    return Err(FootnoteTextError::InvalidSource);
                };
                match super::footnote_table_entries(&message.data, owner_identifier, limits) {
                    Ok(entries) => {
                        for entry in entries {
                            validate_graph_edge(
                                owner.get(),
                                entry.identifier.get(),
                                native_footnote,
                            )?;
                        }
                    },
                    Err(_error) if owner == native_footnote.body_identifier => {
                        return Err(FootnoteTextError::InvalidSource);
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
) -> Result<(), FootnoteTextError> {
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
        return Err(FootnoteTextError::InvalidSource);
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
) -> Result<Vec<u8>, FootnoteTextError> {
    let source_catalog = &source.state.source;
    let entry = source_catalog
        .package()
        .iter()
        .find(|entry| entry.name() == component_name)
        .ok_or(FootnoteTextError::InvalidSource)?;
    if entry.is_opaque() {
        return Err(FootnoteTextError::InvalidSource);
    }
    let stream = SnappyStream::decompress_with_limits(
        entry.data(),
        source_catalog
            .limits()
            .snappy_limits()
            .map_err(map_archive_error)?,
    )
    .map_err(map_core_error)?;
    let mut archive =
        Archive::parse_with_limits(stream.as_bytes(), archive_limits).map_err(map_core_error)?;
    let rewrite_limits =
        storage_rewrite_limits(source_catalog.limits()).map_err(map_storage_wire_limits_error)?;

    if rewrite_text {
        let (message_index, message_type, rewritten_payload) = {
            let object = archive
                .object(native_footnote.storage_identifier.get())
                .ok_or(FootnoteTextError::InvalidSource)?;
            let mut messages = object
                .messages
                .iter()
                .enumerate()
                .filter(|(_index, message)| is_body_text_message_type(message.type_));
            let (message_index, message) =
                messages.next().ok_or(FootnoteTextError::InvalidSource)?;
            if messages.next().is_some() {
                return Err(FootnoteTextError::InvalidSource);
            }
            let rewritten =
                rewrite_storage_payload(&message.data, before_text, after_text, rewrite_limits)?;
            (message_index, message.type_, rewritten)
        };
        archive
            .object_mut(native_footnote.storage_identifier.get())
            .ok_or(FootnoteTextError::InvalidSource)?
            .replace_message_preserving_header_with_limits(
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
                .ok_or(FootnoteTextError::InvalidSource)?;
            let mut messages = object
                .messages
                .iter()
                .enumerate()
                .filter(|(_index, message)| message.type_ == FOOTNOTE_REFERENCE_MESSAGE_TYPE);
            let (message_index, message) =
                messages.next().ok_or(FootnoteTextError::InvalidSource)?;
            if messages.next().is_some() {
                return Err(FootnoteTextError::InvalidSource);
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
            .ok_or(FootnoteTextError::InvalidSource)?
            .replace_message_preserving_header_with_limits(
                message_index,
                RawMessage {
                    type_: message_type,
                    data: rewritten_payload,
                },
                archive_limits,
            )
            .map_err(map_core_error)?;
    }

    let rewritten_archive = archive
        .to_bytes_with_limits(archive_limits)
        .map_err(map_core_error)?;
    SnappyStream::compress(&rewritten_archive).map_err(map_core_error)
}

/// Compare source and candidate logical members while allowing only decoded
/// payload changes.  ZIP reassembly necessarily changes offsets, sizes, and
/// CRC metadata for edited members, but every name, raw authority, and
/// compression policy must remain disjoint and stable.  The returned count is
/// also the truthful physical-component diagnostic for a text+custom-marker
/// edit that happens to span two members.
fn changed_component_count(
    source: &Package,
    candidate: &Package,
) -> Result<usize, FootnoteTextError> {
    let source_entry_count = source.state.source.package().iter().count();
    let candidate_entry_count = candidate.state.source.package().iter().count();
    if source_entry_count != candidate_entry_count {
        return Err(FootnoteTextError::Verification);
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
            return Err(FootnoteTextError::Verification);
        }
        if source_entry.data() != candidate_entry.data() {
            changed = changed
                .checked_add(1)
                .ok_or(FootnoteTextError::Verification)?;
        }
    }
    Ok(changed)
}

fn rewrite_storage_payload(
    payload: &[u8],
    before: &str,
    after: &str,
    rewrite_limits: litchi_iwa_text_wire::RewriteLimits,
) -> Result<Vec<u8>, FootnoteTextError> {
    let storage = litchi_iwa_text_wire::decode_storage_with_limits(payload, rewrite_limits)
        .map_err(map_text_rewrite_error)?
        .into_storage();
    let full_text = storage.text();
    let content = full_text
        .strip_prefix(STORAGE_TEXT_PREFIX)
        .ok_or(FootnoteTextError::InvalidSource)?;
    if content != before {
        return Err(FootnoteTextError::InvalidSource);
    }
    let prefix_units = checked_utf16_units(STORAGE_TEXT_PREFIX)?;
    let content_units = checked_utf16_units(content)?;
    let content_end =
        prefix_units
            .checked_add(content_units)
            .ok_or(FootnoteTextError::LimitExceeded {
                kind: FootnoteTextLimitKind::TextUnits,
                observed: u64::MAX,
                maximum: u64::from(u32::MAX),
            })?;
    let rewritten = litchi_iwa_text_wire::rewrite_storage_text_with_limits(
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
                .ok_or(FootnoteTextError::LimitExceeded {
                    kind: FootnoteTextLimitKind::TextUnits,
                    observed: u64::MAX,
                    maximum: u64::from(u32::MAX),
                })?
        || !rewritten.removed_object_references().is_empty()
        || !rewritten.removed_object_references_by_field().is_empty()
        || rewritten.object_reference_occurrences_before()
            != rewritten.object_reference_occurrences_after()
        || !rewritten.changed()
    {
        return Err(FootnoteTextError::Verification);
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
) -> Result<Vec<u8>, FootnoteTextError> {
    let decoded = pages_footnote_codec::decode_footnote_reference(
        payload,
        footnote_decode_options(payload, physical_limits).map_err(map_package_error)?,
    )
    .map_err(|_error| FootnoteTextError::InvalidSource)?;
    if decoded.super_kind() != Some(FOOTNOTE_MARK_KIND)
        || decoded
            .contained_storage()
            .is_none_or(|reference| reference.identifier() != native_footnote.storage_identifier)
        || decoded.custom_mark_string() != before_custom_mark
    {
        return Err(FootnoteTextError::InvalidSource);
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
    .map_err(|_error| FootnoteTextError::Verification)?;
    if readback.super_kind() != Some(FOOTNOTE_MARK_KIND)
        || readback
            .contained_storage()
            .is_none_or(|reference| reference.identifier() != native_footnote.storage_identifier)
        || readback.custom_mark_string() != after_custom_mark
    {
        return Err(FootnoteTextError::Verification);
    }
    Ok(rewritten)
}

fn rewrite_custom_mark_wire(
    source: &[u8],
    before: Option<&str>,
    after: Option<&str>,
    limits: WireLimits,
) -> Result<Vec<u8>, FootnoteTextError> {
    let view = WireView::parse_with_limits(source, limits).map_err(map_wire_error)?;
    let mut selected = None;
    for field in view.fields().filter(|field| field.number() == 3) {
        if selected.is_some() || field.wire_type() != 2 {
            return Err(FootnoteTextError::InvalidSource);
        }
        field.validate_canonical_framing().map_err(map_wire_error)?;
        let value = std::str::from_utf8(field.payload())
            .map_err(|_error| FootnoteTextError::InvalidSource)?;
        if Some(value) != before {
            return Err(FootnoteTextError::InvalidSource);
        }
        selected = Some(field);
    }
    if selected.is_some() != before.is_some() {
        return Err(FootnoteTextError::InvalidSource);
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
        .ok_or(FootnoteTextError::InvalidSource)?;
    if output_length > limits.max_output_bytes() {
        return Err(FootnoteTextError::LimitExceeded {
            kind: FootnoteTextLimitKind::OutputBytes,
            observed: usize_to_u64(output_length),
            maximum: usize_to_u64(limits.max_output_bytes()),
        });
    }
    let output_fields = view
        .len()
        .saturating_sub(if selected.is_some() { 1 } else { 0 })
        .saturating_add(if after.is_some() { 1 } else { 0 });
    if output_fields > limits.max_fields() {
        return Err(FootnoteTextError::LimitExceeded {
            kind: FootnoteTextLimitKind::WireFields,
            observed: usize_to_u64(output_fields),
            maximum: usize_to_u64(limits.max_fields()),
        });
    }
    let work = view
        .len()
        .checked_add(output_fields)
        .ok_or(FootnoteTextError::InvalidSource)?;
    if work > limits.max_rewrite_work() {
        return Err(FootnoteTextError::LimitExceeded {
            kind: FootnoteTextLimitKind::WireWork,
            observed: usize_to_u64(work),
            maximum: usize_to_u64(limits.max_rewrite_work()),
        });
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_length)
        .map_err(|_error| FootnoteTextError::Allocation {
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
        return Err(FootnoteTextError::Verification);
    }
    Ok(output)
}

fn encoded_bytes_field_length(
    number: u32,
    payload_length: usize,
    limits: WireLimits,
) -> Result<usize, FootnoteTextError> {
    encoded_len((u64::from(number) << 3) | 2)
        .checked_add(encoded_len(usize_to_u64(payload_length)))
        .and_then(|length| length.checked_add(payload_length))
        .ok_or(FootnoteTextError::LimitExceeded {
            kind: FootnoteTextLimitKind::OutputBytes,
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
) -> Result<(), FootnoteTextError> {
    let before = source.body_footnotes().map_err(map_package_error)?;
    let after = candidate.body_footnotes().map_err(map_package_error)?;
    if before.len() != after.len() {
        return Err(FootnoteTextError::Verification);
    }
    for (old, new) in before.iter().zip(&after) {
        if old.position != new.position {
            return Err(FootnoteTextError::Verification);
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
            return Err(FootnoteTextError::Verification);
        }
    }
    if source.stats().total_objects() != candidate.stats().total_objects() {
        return Err(FootnoteTextError::Verification);
    }
    let source_graphs = native_footnotes(source)?;
    let candidate_graphs = native_footnotes(candidate)?;
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
        return Err(FootnoteTextError::Verification);
    }
    Ok(())
}

fn validate_text(text: &str) -> Result<(), FootnoteTextError> {
    if text.contains('\u{000e}') || text.contains('\u{fffc}') {
        return Err(FootnoteTextError::StructuralMarker);
    }
    checked_utf16_units(text)?;
    if text.len() > crate::footnote::body::MAX_TEXT_BYTES {
        return Err(FootnoteTextError::TextTooLarge);
    }
    Ok(())
}

fn checked_utf16_units(text: &str) -> Result<usize, FootnoteTextError> {
    let units = text.encode_utf16().count();
    let maximum = usize::try_from(u32::MAX).unwrap_or(usize::MAX);
    if units > maximum {
        return Err(FootnoteTextError::LimitExceeded {
            kind: FootnoteTextLimitKind::TextUnits,
            observed: usize_to_u64(units),
            maximum: u64::from(u32::MAX),
        });
    }
    Ok(units)
}

fn position_from_anchor(anchor: u32) -> Result<Position, FootnoteTextError> {
    let index = usize::try_from(anchor).map_err(|_error| FootnoteTextError::LimitExceeded {
        kind: FootnoteTextLimitKind::TextUnits,
        observed: u64::from(anchor),
        maximum: u64::from(u32::MAX),
    })?;
    Position::from_utf16_index(index).map_err(|_error| FootnoteTextError::InvalidSource)
}

fn validate_custom_mark(custom_mark: Option<&str>) -> Result<(), FootnoteTextError> {
    if custom_mark.is_some_and(|value| value.len() > crate::footnote::body::MAX_CUSTOM_MARK_BYTES) {
        return Err(FootnoteTextError::CustomMarkTooLarge);
    }
    Ok(())
}

fn clone_footnote(value: &Footnote) -> Result<Footnote, FootnoteTextError> {
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

fn try_owned_text(text: &str) -> Result<String, FootnoteTextError> {
    let mut owned = String::new();
    owned
        .try_reserve_exact(text.len())
        .map_err(|_error| FootnoteTextError::Allocation { amount: text.len() })?;
    owned.push_str(text);
    Ok(owned)
}

fn map_footnote_value_error(error: crate::footnote::body::Error) -> FootnoteTextError {
    match error {
        crate::footnote::body::Error::TextTooLarge => FootnoteTextError::TextTooLarge,
        crate::footnote::body::Error::CustomMarkTooLarge => FootnoteTextError::CustomMarkTooLarge,
        crate::footnote::body::Error::PositionOutOfRange => FootnoteTextError::InvalidSource,
    }
}

fn map_wire_error(error: litchi_iwa_common::Error) -> FootnoteTextError {
    match error {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => FootnoteTextError::LimitExceeded {
            kind: match kind {
                litchi_iwa_common::LimitKind::InputBytes => FootnoteTextLimitKind::InputBytes,
                litchi_iwa_common::LimitKind::Fields => FootnoteTextLimitKind::WireFields,
                litchi_iwa_common::LimitKind::OutputBytes => FootnoteTextLimitKind::OutputBytes,
                litchi_iwa_common::LimitKind::Nesting => FootnoteTextLimitKind::WireNesting,
                litchi_iwa_common::LimitKind::RewriteWork => FootnoteTextLimitKind::WireWork,
                litchi_iwa_common::LimitKind::TableRows
                | litchi_iwa_common::LimitKind::TableColumns
                | litchi_iwa_common::LimitKind::TableCells
                | litchi_iwa_common::LimitKind::MaterializedCells => FootnoteTextLimitKind::Entries,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(limit),
        },
        litchi_iwa_common::Error::Allocation { amount, .. } => {
            FootnoteTextError::Allocation { amount }
        },
        litchi_iwa_common::Error::InvalidFormat(_)
        | litchi_iwa_common::Error::InvalidLimit { .. } => FootnoteTextError::InvalidSource,
    }
}

fn map_package_error(error: PackageError) -> FootnoteTextError {
    map_package_error_with_kind(error, FootnoteTextLimitKind::WireBytes)
}

fn map_package_error_with_kind(
    error: PackageError,
    payload_limit_kind: FootnoteTextLimitKind,
) -> FootnoteTextError {
    match error {
        PackageError::Archive(error) => map_archive_error(error),
        PackageError::PayloadLimit { observed, limit } => FootnoteTextError::LimitExceeded {
            kind: payload_limit_kind,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(limit),
        },
        PackageError::ObjectLimit { observed, limit } => FootnoteTextError::LimitExceeded {
            kind: FootnoteTextLimitKind::Entries,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(limit),
        },
        PackageError::Allocation { amount } => FootnoteTextError::Allocation { amount },
        PackageError::SectionNamesTooLarge { observed, limit } => {
            FootnoteTextError::LimitExceeded {
                kind: FootnoteTextLimitKind::TextBytes,
                observed: usize_to_u64(observed),
                maximum: usize_to_u64(limit),
            }
        },
        PackageError::Io(_)
        | PackageError::Detection(_)
        | PackageError::NotPages
        | PackageError::InvalidFormat(_)
        | PackageError::Semantic(_) => FootnoteTextError::InvalidSource,
    }
}

fn map_storage_wire_limits_error(error: super::StorageWireLimitsError) -> FootnoteTextError {
    match error {
        super::StorageWireLimitsError::Physical(error) => map_archive_error(error),
        super::StorageWireLimitsError::Wire(_error) => FootnoteTextError::InvalidSource,
    }
}

fn map_text_rewrite_error(error: RewriteError) -> FootnoteTextError {
    match error {
        RewriteError::LimitExceeded {
            resource,
            observed,
            limit,
        } => FootnoteTextError::LimitExceeded {
            kind: text_limit_kind(resource),
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(limit),
        },
        RewriteError::Allocation { amount, .. } => FootnoteTextError::Allocation { amount },
        _ => FootnoteTextError::InvalidSource,
    }
}

fn text_limit_kind(resource: &str) -> FootnoteTextLimitKind {
    match resource {
        "text bytes" => FootnoteTextLimitKind::TextBytes,
        "text fragments" | "table entries" | "object references" => FootnoteTextLimitKind::Entries,
        "nesting" => FootnoteTextLimitKind::WireNesting,
        "fields" => FootnoteTextLimitKind::WireFields,
        "rewrite work" | "aggregate nested scan bytes" => FootnoteTextLimitKind::WireWork,
        _ => FootnoteTextLimitKind::WireBytes,
    }
}

fn map_archive_error(error: litchi_iwa_archive::Error) -> FootnoteTextError {
    match error {
        litchi_iwa_archive::Error::Limit {
            kind,
            observed,
            maximum,
        } => FootnoteTextError::LimitExceeded {
            kind: match kind {
                litchi_iwa_archive::LimitKind::InputBytes => FootnoteTextLimitKind::InputBytes,
                litchi_iwa_archive::LimitKind::OutputBytes => FootnoteTextLimitKind::OutputBytes,
                litchi_iwa_archive::LimitKind::Entries => FootnoteTextLimitKind::Entries,
                litchi_iwa_archive::LimitKind::MemberNameBytes
                | litchi_iwa_archive::LimitKind::MetadataBytes
                | litchi_iwa_archive::LimitKind::CompressedEntryBytes
                | litchi_iwa_archive::LimitKind::EntryBytes
                | litchi_iwa_archive::LimitKind::IwaStreamBytes => {
                    FootnoteTextLimitKind::EntryBytes
                },
                litchi_iwa_archive::LimitKind::TotalBytes
                | litchi_iwa_archive::LimitKind::IwaTotalBytes => FootnoteTextLimitKind::TotalBytes,
            },
            observed,
            maximum,
        },
        litchi_iwa_archive::Error::Allocation { amount, .. } => {
            FootnoteTextError::Allocation { amount }
        },
        litchi_iwa_archive::Error::Iwa(error) => map_core_error(error),
        litchi_iwa_archive::Error::Io(_)
        | litchi_iwa_archive::Error::Zip { .. }
        | litchi_iwa_archive::Error::InvalidLimits(_)
        | litchi_iwa_archive::Error::Encrypted
        | litchi_iwa_archive::Error::SourceChanged { .. }
        | litchi_iwa_archive::Error::DirectoryChanged { .. }
        | litchi_iwa_archive::Error::Reassembly(_)
        | litchi_iwa_archive::Error::InvalidBundle(_) => FootnoteTextError::InvalidSource,
    }
}

fn map_core_error(error: litchi_iwa_core::Error) -> FootnoteTextError {
    match error {
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => FootnoteTextError::LimitExceeded {
            kind: match kind {
                litchi_iwa_core::LimitKind::Objects
                | litchi_iwa_core::LimitKind::Messages
                | litchi_iwa_core::LimitKind::MessagesPerObject
                | litchi_iwa_core::LimitKind::MetadataItems => FootnoteTextLimitKind::Entries,
                litchi_iwa_core::LimitKind::MessageBytes => FootnoteTextLimitKind::WireBytes,
                litchi_iwa_core::LimitKind::HeaderFields => FootnoteTextLimitKind::WireFields,
                litchi_iwa_core::LimitKind::HeaderNesting => FootnoteTextLimitKind::WireNesting,
                litchi_iwa_core::LimitKind::ArchiveBytes
                | litchi_iwa_core::LimitKind::ObjectBytes
                | litchi_iwa_core::LimitKind::HeaderBytes
                | litchi_iwa_core::LimitKind::HeaderMemoryBytes
                | litchi_iwa_core::LimitKind::SnappyChunkBytes
                | litchi_iwa_core::LimitKind::SnappyStreamBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedChunkBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedStreamBytes
                | litchi_iwa_core::LimitKind::SnappyFrames => FootnoteTextLimitKind::EntryBytes,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            FootnoteTextError::Allocation { amount: requested }
        },
        litchi_iwa_core::Error::InvalidArchive { .. }
        | litchi_iwa_core::Error::InvalidLimits { .. }
        | litchi_iwa_core::Error::HeaderCodec { .. }
        | litchi_iwa_core::Error::Io(_)
        | litchi_iwa_core::Error::Snappy { .. } => FootnoteTextError::InvalidSource,
    }
}

fn usize_to_u64(value: usize) -> u64 {
    u64::try_from(value).unwrap_or(u64::MAX)
}

#[cfg(test)]
mod tests {
    use super::{
        FootnoteTextError, FootnoteTextLimitKind, checked_utf16_units, is_canonical_component_name,
        position_from_anchor, rewrite_custom_mark_wire,
    };
    use litchi_iwa_common::WireLimits;

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
            Err(FootnoteTextError::LimitExceeded {
                kind: FootnoteTextLimitKind::TextUnits,
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
}
