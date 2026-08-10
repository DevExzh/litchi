//! Immutable, exact-source transactions for the Pages document page layout.

use std::fmt;
use std::sync::Arc;

use litchi_iwa_archive::{SourceCatalog, package::EntryEdit};
use litchi_iwa_common::{
    WireLimits, decode_varint_from_bytes, encode_varint_into,
    varint::encoded_len,
    wire::{WireFieldView, WireView, patch_length_delimited_field},
};
use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
use litchi_iwa_protos::pages_page_layout_codec::{
    self, DecodeOptions as PageLayoutDecodeOptions, PageLayoutSnapshot, WireResourceLimit,
};
use thiserror::Error;

use super::{Package, PackageError};
use crate::page_layout::{self, Layout, Orientation};

const DOCUMENT_IDENTIFIER: u64 = 1;
const DOCUMENT_MESSAGE_TYPE: u32 = 10_000;
const SHARED_VIEW_STATE_MESSAGE_TYPE: u32 = 210;
const VIEW_STATE_ROOT_MESSAGE_TYPE: u32 = 10_147;
const DOCUMENT_SUPER_FIELD: u32 = 15;
const DOCUMENT_DEPRECATED_LAYOUT_FIELD: u32 = 11;
const DOCUMENT_DEPRECATED_VIEW_STATE_FIELD: u32 = 12;
const SHARED_DOCUMENT_VIEW_STATE_FIELD: u32 = 5;
const SHARED_VIEW_STATE_ROOT_FIELD: u32 = 1;
const VIEW_STATE_LAYOUT_FIELD: u32 = 1;
const VIEW_STATE_UI_FIELD: u32 = 2;
const REFERENCE_IDENTIFIER_FIELD: u32 = 1;
const REFERENCE_TYPE_FIELD: u32 = 2;
const REFERENCE_EXTERNAL_FIELD: u32 = 3;
const PREVIEW_ENTRY_NAMES: [&str; 3] = ["preview.jpg", "preview-micro.jpg", "preview-web.jpg"];
const FLOAT_FIELDS: [u32; 9] = [30, 31, 32, 33, 34, 35, 36, 37, 38];
const VERTICAL_BODY_FIELD: u32 = 39;
const ORIENTATION_FIELD: u32 = 42;

/// A finite resource enforced while reading or publishing a page-layout edit.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum PageLayoutLimitKind {
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
    /// ZIP names and structural metadata.
    PackageBytes,
    /// Bytes in one decoded component payload.
    PayloadBytes,
    /// Aggregate decoded component payload bytes.
    TotalPayloadBytes,
    /// Component objects inspected by the transaction.
    PayloadObjects,
    /// Component messages inspected by the transaction.
    PayloadMessages,
    /// Component framing or metadata items inspected by the transaction.
    PayloadItems,
    /// Bytes inspected by the page-layout wire projection.
    WireBytes,
    /// Fields inspected by the page-layout wire projection.
    WireFields,
    /// Nesting used by the page-layout wire projection.
    WireNesting,
    /// Aggregate page-layout wire rewrite work.
    WireWork,
}

impl fmt::Display for PageLayoutLimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::OutputBytes => "output bytes",
            Self::Entries => "ZIP entries",
            Self::EntryBytes => "ZIP entry bytes",
            Self::TotalEntryBytes => "total ZIP entry bytes",
            Self::PackageBytes => "package metadata bytes",
            Self::PayloadBytes => "component payload bytes",
            Self::TotalPayloadBytes => "total component payload bytes",
            Self::PayloadObjects => "component objects",
            Self::PayloadMessages => "component messages",
            Self::PayloadItems => "component metadata items",
            Self::WireBytes => "wire bytes",
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting",
            Self::WireWork => "wire work",
        })
    }
}

/// Failure from a semantic Pages page-layout read or immutable transaction.
///
/// Display output never exposes package bytes, member names, object
/// identifiers, or retained patch artifacts.
#[derive(Debug, Error, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum PageLayoutError {
    /// The requested value violates an archive-free page-layout invariant.
    #[error("invalid Pages page layout: {0}")]
    InvalidLayout(#[source] page_layout::Error),
    /// The package source cannot publish a preservation-safe changed edit.
    #[error("the Pages package source does not support exact page-layout editing")]
    UnsupportedSource,
    /// The selected native layout or its derived cache cannot be handled safely.
    #[error("the Pages source has no unambiguous editable page layout")]
    InvalidSource,
    /// A finite transaction resource ceiling was exceeded.
    #[error("Pages page-layout {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    LimitExceeded {
        /// Resource category that exceeded its ceiling.
        kind: PageLayoutLimitKind,
        /// Observed or requested resource amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded transaction allocation failed.
    #[error("could not allocate {amount} units for the Pages page-layout transaction")]
    Allocation {
        /// Requested bytes or elements.
        amount: usize,
    },
    /// Complete candidate reopening did not reproduce the requested state.
    #[error("the edited Pages page layout failed semantic verification")]
    Verification,
    /// The supplied patch was not created from this exact package artifact.
    #[error("the Pages page-layout patch does not match the exact source package")]
    PatchConflict,
}

/// A document-wide page layout staged against one immutable package snapshot.
pub struct PageLayoutEdit<'a> {
    source: &'a Package,
    before: Layout,
    layout: Layout,
}

impl fmt::Debug for PageLayoutEdit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("PageLayoutEdit")
            .field("layout", &self.layout)
            .finish_non_exhaustive()
    }
}

impl PageLayoutEdit<'_> {
    /// Return the lossless layout that would be published.
    #[must_use]
    pub const fn layout(&self) -> Layout {
        self.layout
    }

    /// Replace the staged layout after validating every public invariant.
    ///
    /// # Errors
    ///
    /// Returns [`PageLayoutError::InvalidLayout`] without changing the staged
    /// value when `layout` violates an archive-free semantic invariant.
    ///
    /// # Costs
    ///
    /// Validation is constant-time and does not inspect the package.
    pub fn set_layout(&mut self, layout: Layout) -> Result<&mut Self, PageLayoutError> {
        layout.validate().map_err(PageLayoutError::InvalidLayout)?;
        self.layout = layout;
        Ok(self)
    }

    /// Remove all optional native page-layout field presence.
    pub fn clear(&mut self) -> &mut Self {
        self.layout = Layout::empty();
        self
    }

    /// Validate and atomically publish the staged immutable candidate.
    ///
    /// Equal layouts reuse the exact source allocation and retain previews and
    /// view-state caches. A changed edit requires an exact flat package,
    /// invalidates the derived layout-state edge, removes root previews, and
    /// fully reopens the complete candidate before publication.
    ///
    /// # Errors
    ///
    /// Returns a typed error if the staged value is invalid, the source is not
    /// the exact immutable snapshot captured by this edit, ownership metadata
    /// is ambiguous, a finite limit is exceeded, or complete readback fails.
    ///
    /// # Costs
    ///
    /// A no-op performs a bounded semantic read and reuses the source
    /// allocation. A changed commit scans the selected ownership graph,
    /// rewrites at most two component payloads, reassembles the package, and
    /// fully reopens and verifies the candidate under the source limits.
    pub fn commit(self) -> Result<PageLayoutCommit, PageLayoutError> {
        self.layout
            .validate()
            .map_err(PageLayoutError::InvalidLayout)?;
        if page_layout(self.source)? != self.before {
            return Err(PageLayoutError::InvalidSource);
        }
        let source_catalog = &self.source.state.source;
        let source = source_catalog.shared_source();
        let source_fingerprint = fingerprint(&source);
        if self.before == self.layout {
            return Ok(PageLayoutCommit {
                package: self.source.snapshot(),
                patch: PageLayoutPatch {
                    source: Arc::clone(&source),
                    target: source,
                    source_fingerprint,
                    target_fingerprint: source_fingerprint,
                    before: self.before,
                    after: self.layout,
                    source_layout_state: None,
                    target_layout_state: None,
                    source_preview_count: 0,
                    target_preview_count: 0,
                    touched_components: 0,
                },
                diagnostics: PageLayoutDiagnostics::unchanged(),
            });
        }
        if !source_catalog.source_is_exact() {
            return Err(PageLayoutError::UnsupportedSource);
        }

        let source_layout_state = view_state_layout_identifier(self.source)?;
        let source_preview_count = preview_count(self.source);

        let (package, touched_components) =
            rewrite_page_layout(self.source, self.before, self.layout, source_layout_state)?;
        let target = package.state.source.shared_source();
        let target_fingerprint = fingerprint(&target);
        let target_layout_state = view_state_layout_identifier(&package)?;
        let target_preview_count = preview_count(&package);
        let deleted_previews = source_preview_count.saturating_sub(target_preview_count);
        Ok(PageLayoutCommit {
            package,
            patch: PageLayoutPatch {
                source,
                target,
                source_fingerprint,
                target_fingerprint,
                before: self.before,
                after: self.layout,
                source_layout_state,
                target_layout_state,
                source_preview_count,
                target_preview_count,
                touched_components,
            },
            diagnostics: PageLayoutDiagnostics::published(touched_components, deleted_previews),
        })
    }
}

/// A reversible patch bound to exact source and target package artifacts.
///
/// Exact artifacts and ownership facts are retained privately. Exposed
/// fingerprints are compact diagnostics, not a substitute for exact matching.
#[derive(Clone, PartialEq)]
pub struct PageLayoutPatch {
    source: Arc<[u8]>,
    target: Arc<[u8]>,
    source_fingerprint: u64,
    target_fingerprint: u64,
    before: Layout,
    after: Layout,
    source_layout_state: Option<u64>,
    target_layout_state: Option<u64>,
    source_preview_count: usize,
    target_preview_count: usize,
    touched_components: usize,
}

impl fmt::Debug for PageLayoutPatch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("PageLayoutPatch")
            .field("before", &self.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl PageLayoutPatch {
    /// Return the semantic layout required from the patch source.
    #[must_use]
    pub const fn before(&self) -> Layout {
        self.before
    }

    /// Return the semantic layout represented by the patch target.
    #[must_use]
    pub const fn after(&self) -> Layout {
        self.after
    }

    /// Return the source artifact's compact diagnostic fingerprint.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.source_fingerprint
    }

    /// Return the target artifact's compact diagnostic fingerprint.
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.target_fingerprint
    }

    /// Return whether this patch retains both the semantic layout and bytes.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after
            && self.source_fingerprint == self.target_fingerprint
            && (Arc::ptr_eq(&self.source, &self.target)
                || self.source.as_ref() == self.target.as_ref())
    }

    /// Return the exact patch from the target artifact back to its source.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            source: Arc::clone(&self.target),
            target: Arc::clone(&self.source),
            source_fingerprint: self.target_fingerprint,
            target_fingerprint: self.source_fingerprint,
            before: self.after,
            after: self.before,
            source_layout_state: self.target_layout_state,
            target_layout_state: self.source_layout_state,
            source_preview_count: self.target_preview_count,
            target_preview_count: self.source_preview_count,
            touched_components: self.touched_components,
        }
    }
}

/// Compact evidence describing one page-layout commit.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct PageLayoutDiagnostics {
    changed: bool,
    touched_components: usize,
    deleted_previews: usize,
    full_reparse_performed: bool,
}

impl PageLayoutDiagnostics {
    const fn unchanged() -> Self {
        Self {
            changed: false,
            touched_components: 0,
            deleted_previews: 0,
            full_reparse_performed: false,
        }
    }

    const fn published(touched_components: usize, deleted_previews: usize) -> Self {
        Self {
            changed: true,
            touched_components,
            deleted_previews,
            full_reparse_performed: true,
        }
    }

    /// Return whether the package differs from its source.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// Return the number of rewritten IWA components.
    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }

    /// Return the number of root preview entries deleted by the edit.
    #[must_use]
    pub const fn deleted_previews(self) -> usize {
        self.deleted_previews
    }

    /// Return whether the complete target package was reopened.
    #[must_use]
    pub const fn full_reparse_performed(self) -> bool {
        self.full_reparse_performed
    }
}

/// The fully verified result of one immutable page-layout transaction.
#[must_use = "a Pages page-layout commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct PageLayoutCommit {
    package: Package,
    patch: PageLayoutPatch,
    diagnostics: PageLayoutDiagnostics,
}

impl PageLayoutCommit {
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
    pub const fn patch(&self) -> &PageLayoutPatch {
        &self.patch
    }

    /// Borrow compact publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &PageLayoutDiagnostics {
        &self.diagnostics
    }
}

impl Package {
    /// Read the lossless page dimensions, margins, scale, and orientation.
    ///
    /// # Errors
    ///
    /// Returns a typed error when the selected document payload is ambiguous,
    /// malformed, semantically invalid, or exceeds a finite decode limit.
    ///
    /// # Costs
    ///
    /// Decodes one bounded document payload without modifying the package.
    pub fn page_layout(&self) -> Result<Layout, PageLayoutError> {
        page_layout(self)
    }

    /// Start a document-wide immutable page-layout edit.
    ///
    /// # Errors
    ///
    /// Returns the same typed read errors as [`Package::page_layout`].
    ///
    /// # Costs
    ///
    /// Decodes one bounded document payload and retains only semantic values
    /// plus a borrow of this immutable package.
    pub fn edit_page_layout(&self) -> Result<PageLayoutEdit<'_>, PageLayoutError> {
        let before = page_layout(self)?;
        Ok(PageLayoutEdit {
            source: self,
            before,
            layout: before,
        })
    }

    /// Apply an exact-source-checked page-layout patch.
    ///
    /// `patch` privately retains the exact source and target artifacts; its
    /// public fingerprints are diagnostic only. A no-op reuses this package's
    /// allocation. A changed target is fully reopened and verified.
    ///
    /// # Errors
    ///
    /// Returns [`PageLayoutError::PatchConflict`] unless this package exactly
    /// matches the patch source, or another typed error if the retained target
    /// cannot be safely reopened and verified under the source limits.
    ///
    /// # Costs
    ///
    /// Exact matching scans package bytes unless allocations are shared. A
    /// changed apply also reopens the complete retained target and verifies its
    /// layout, cache invalidation, previews, and preserved document semantics.
    pub fn apply_page_layout(
        &self,
        patch: &PageLayoutPatch,
    ) -> Result<PageLayoutCommit, PageLayoutError> {
        let source_catalog = &self.state.source;
        let source = source_catalog.shared_source();
        let exact_source = Arc::ptr_eq(&source, &patch.source)
            || (fingerprint(source_catalog.source_bytes()) == patch.source_fingerprint
                && source_catalog.source_bytes() == patch.source.as_ref());
        if !exact_source || page_layout(self)? != patch.before {
            return Err(PageLayoutError::PatchConflict);
        }
        if patch.is_noop() {
            return Ok(PageLayoutCommit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: PageLayoutDiagnostics::unchanged(),
            });
        }
        if view_state_layout_identifier(self)? != patch.source_layout_state
            || preview_count(self) != patch.source_preview_count
        {
            return Err(PageLayoutError::PatchConflict);
        }
        if !source_catalog.source_is_exact()
            || fingerprint(&patch.target) != patch.target_fingerprint
        {
            return Err(PageLayoutError::PatchConflict);
        }
        let candidate_source = SourceCatalog::from_shared_bytes_with_limits(
            Arc::clone(&patch.target),
            source_catalog.limits(),
        )
        .map_err(map_archive_error)?;
        let candidate =
            Package::from_source_catalog(candidate_source).map_err(map_package_error)?;
        verify_candidate(
            self,
            &candidate,
            patch.after,
            patch.target_layout_state,
            patch.target_preview_count,
        )?;
        let deleted_previews = patch
            .source_preview_count
            .saturating_sub(patch.target_preview_count);
        Ok(PageLayoutCommit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: PageLayoutDiagnostics::published(
                patch.touched_components,
                deleted_previews,
            ),
        })
    }
}

#[derive(Debug, Clone)]
struct ViewStateLocation {
    component_name: String,
    object_identifier: u64,
    layout_identifier: u64,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum LayoutScalar {
    Fixed32(u32),
    Varint(u64),
}

fn page_layout(package: &Package) -> Result<Layout, PageLayoutError> {
    let (_component_name, _object_identifier, payload) = document_payload(package)?;
    strict_layout(payload, wire_limits(package)?)
}

fn document_payload(package: &Package) -> Result<(&str, u64, &[u8]), PageLayoutError> {
    let mut found = None;
    for component in package.state.source.components().iter() {
        let Some(object) = component.archive().object(DOCUMENT_IDENTIFIER) else {
            continue;
        };
        if found.is_some() {
            return Err(PageLayoutError::InvalidSource);
        }
        let (_index, message) = unique_message(object, DOCUMENT_MESSAGE_TYPE)?;
        found = Some((
            component.name(),
            DOCUMENT_IDENTIFIER,
            message.data.as_slice(),
        ));
    }
    found.ok_or(PageLayoutError::InvalidSource)
}

fn strict_layout(payload: &[u8], limits: WireLimits) -> Result<Layout, PageLayoutError> {
    let recursion =
        u32::try_from(limits.max_nesting()).map_err(|_error| PageLayoutError::LimitExceeded {
            kind: PageLayoutLimitKind::WireNesting,
            observed: usize_to_u64(limits.max_nesting()),
            maximum: u64::from(u32::MAX),
        })?;
    let projected = pages_page_layout_codec::decode_page_layout(
        payload,
        PageLayoutDecodeOptions::new(
            limits.max_input_bytes(),
            limits.max_fields(),
            limits.max_rewrite_work(),
            recursion,
        ),
    )
    .map_err(map_layout_codec_error)?;
    projected_layout(projected)
}

fn projected_layout(snapshot: PageLayoutSnapshot) -> Result<Layout, PageLayoutError> {
    Layout::new(
        snapshot.page_width(),
        snapshot.page_height(),
        snapshot.left_margin(),
        snapshot.right_margin(),
        snapshot.top_margin(),
        snapshot.bottom_margin(),
        snapshot.header_margin(),
        snapshot.footer_margin(),
        snapshot.page_scale(),
        snapshot.orientation().map(Orientation::from_raw),
        snapshot.lays_out_body_vertically(),
    )
    .map_err(PageLayoutError::InvalidLayout)
}

fn strict_varint(field: WireFieldView<'_>) -> Result<u64, PageLayoutError> {
    field.validate_canonical_key().map_err(map_wire_error)?;
    let (value, consumed) = decode_varint_from_bytes(field.payload())
        .map_err(|_error| PageLayoutError::InvalidSource)?;
    if consumed != field.payload().len() || encoded_len(value) != consumed {
        return Err(PageLayoutError::InvalidSource);
    }
    Ok(value)
}

fn view_state_location(package: &Package) -> Result<Option<ViewStateLocation>, PageLayoutError> {
    let limits = wire_limits(package)?;
    let (_document_component, document_object) =
        object_location(package, DOCUMENT_IDENTIFIER)?.ok_or(PageLayoutError::InvalidSource)?;
    let (document_message_index, document_message) =
        unique_message(document_object, DOCUMENT_MESSAGE_TYPE)?;
    let document_view =
        WireView::parse_with_limits(&document_message.data, limits).map_err(map_wire_error)?;
    if singular_reference(&document_view, DOCUMENT_DEPRECATED_LAYOUT_FIELD, limits)?.is_some()
        || singular_reference(&document_view, DOCUMENT_DEPRECATED_VIEW_STATE_FIELD, limits)?
            .is_some()
    {
        return Err(PageLayoutError::InvalidSource);
    }
    let super_payload = singular_bytes(&document_view, DOCUMENT_SUPER_FIELD, true)?
        .ok_or(PageLayoutError::InvalidSource)?;
    let shared_document =
        WireView::parse_with_limits(super_payload, limits).map_err(map_wire_error)?;
    let Some(view_state_identifier) =
        singular_reference(&shared_document, SHARED_DOCUMENT_VIEW_STATE_FIELD, limits)?
    else {
        return Ok(None);
    };
    validate_reference_metadata(
        document_object,
        document_message_index,
        view_state_identifier,
        &[DOCUMENT_SUPER_FIELD, SHARED_DOCUMENT_VIEW_STATE_FIELD],
    )?;
    let (_component, view_state) =
        object_location(package, view_state_identifier)?.ok_or(PageLayoutError::InvalidSource)?;
    let (view_state_message_index, view_state_message) =
        unique_message(view_state, SHARED_VIEW_STATE_MESSAGE_TYPE)?;
    let view_state_payload =
        WireView::parse_with_limits(&view_state_message.data, limits).map_err(map_wire_error)?;
    let view_state_root_identifier =
        singular_reference(&view_state_payload, SHARED_VIEW_STATE_ROOT_FIELD, limits)?
            .ok_or(PageLayoutError::InvalidSource)?;
    validate_reference_metadata(
        view_state,
        view_state_message_index,
        view_state_root_identifier,
        &[SHARED_VIEW_STATE_ROOT_FIELD],
    )?;
    let (component_name, root) = object_location(package, view_state_root_identifier)?
        .ok_or(PageLayoutError::InvalidSource)?;
    let (_message_index, root_message) = unique_message(root, VIEW_STATE_ROOT_MESSAGE_TYPE)?;
    let (layout_identifier, ui_identifier) = strict_view_state_root(&root_message.data, limits)?;
    if layout_identifier.is_some() && layout_identifier == ui_identifier {
        return Err(PageLayoutError::InvalidSource);
    }
    layout_identifier
        .map(|selected_layout_identifier| {
            Ok(ViewStateLocation {
                component_name: try_owned(component_name)?,
                object_identifier: view_state_root_identifier,
                layout_identifier: selected_layout_identifier,
            })
        })
        .transpose()
}

fn object_location(
    package: &Package,
    identifier: u64,
) -> Result<Option<(&str, &ArchiveObject)>, PageLayoutError> {
    let mut found = None;
    for component in package.state.source.components().iter() {
        let Some(object) = component.archive().object(identifier) else {
            continue;
        };
        if found.is_some() {
            return Err(PageLayoutError::InvalidSource);
        }
        found = Some((component.name(), object));
    }
    Ok(found)
}

fn singular_bytes<'a>(
    view: &WireView<'a>,
    field_number: u32,
    required: bool,
) -> Result<Option<&'a [u8]>, PageLayoutError> {
    let mut selected = None;
    for field in view.fields().filter(|field| field.number() == field_number) {
        if selected.is_some() || field.wire_type() != 2 {
            return Err(PageLayoutError::InvalidSource);
        }
        field.validate_canonical_framing().map_err(map_wire_error)?;
        selected = Some(field.payload());
    }
    if required && selected.is_none() {
        return Err(PageLayoutError::InvalidSource);
    }
    Ok(selected)
}

fn singular_reference(
    view: &WireView<'_>,
    field_number: u32,
    limits: WireLimits,
) -> Result<Option<u64>, PageLayoutError> {
    singular_bytes(view, field_number, false)?
        .map(|payload| strict_reference(payload, limits))
        .transpose()
}

fn view_state_layout_identifier(package: &Package) -> Result<Option<u64>, PageLayoutError> {
    Ok(view_state_location(package)?.map(|location| location.layout_identifier))
}

fn strict_view_state_root(
    payload: &[u8],
    limits: WireLimits,
) -> Result<(Option<u64>, Option<u64>), PageLayoutError> {
    let view = WireView::parse_with_limits(payload, limits).map_err(map_wire_error)?;
    let mut layout = None;
    let mut ui = None;
    for field in view.fields() {
        let destination = match field.number() {
            VIEW_STATE_LAYOUT_FIELD => &mut layout,
            VIEW_STATE_UI_FIELD => &mut ui,
            _ => continue,
        };
        if destination.is_some() || field.wire_type() != 2 {
            return Err(PageLayoutError::InvalidSource);
        }
        field.validate_canonical_framing().map_err(map_wire_error)?;
        *destination = Some(strict_reference(field.payload(), limits)?);
    }
    Ok((layout, ui))
}

fn strict_reference(payload: &[u8], limits: WireLimits) -> Result<u64, PageLayoutError> {
    let view = WireView::parse_with_limits(payload, limits).map_err(map_wire_error)?;
    let mut identifier = None;
    let mut deprecated_type_seen = false;
    let mut external_seen = false;
    for field in view.fields() {
        match field.number() {
            REFERENCE_IDENTIFIER_FIELD => {
                if identifier.is_some() || field.wire_type() != 0 {
                    return Err(PageLayoutError::InvalidSource);
                }
                identifier = Some(strict_varint(field)?);
            },
            REFERENCE_TYPE_FIELD => {
                if std::mem::replace(&mut deprecated_type_seen, true) || field.wire_type() != 0 {
                    return Err(PageLayoutError::InvalidSource);
                }
                strict_varint(field)?;
            },
            REFERENCE_EXTERNAL_FIELD => {
                if std::mem::replace(&mut external_seen, true) || field.wire_type() != 0 {
                    return Err(PageLayoutError::InvalidSource);
                }
                if strict_varint(field)? != 0 {
                    return Err(PageLayoutError::InvalidSource);
                }
            },
            _ => {},
        }
    }
    match identifier {
        Some(0) | None => Err(PageLayoutError::InvalidSource),
        Some(selected_identifier) => Ok(selected_identifier),
    }
}

fn rewrite_page_layout(
    source: &Package,
    before: Layout,
    after: Layout,
    source_layout_state: Option<u64>,
) -> Result<(Package, usize), PageLayoutError> {
    let source_catalog = &source.state.source;
    let root_component = try_owned(document_payload(source)?.0)?;
    let view_location = view_state_location(source)?;
    if view_location.as_ref().map(|value| value.layout_identifier) != source_layout_state {
        return Err(PageLayoutError::InvalidSource);
    }
    let shared_component = view_location
        .as_ref()
        .is_some_and(|location| location.component_name == root_component);
    let root_compressed = rewrite_document_component(
        source,
        &root_component,
        before,
        after,
        shared_component.then_some(view_location.as_ref()).flatten(),
    )?;
    let view_compressed = if shared_component {
        None
    } else {
        view_location
            .as_ref()
            .map(|location| rewrite_view_state_component(source, location))
            .transpose()?
    };

    let mut compressed = Vec::new();
    compressed
        .try_reserve_exact(usize::from(view_compressed.is_some()) + 1)
        .map_err(|_allocation| PageLayoutError::Allocation { amount: 2 })?;
    compressed.push((root_component, root_compressed));
    if let (Some(location), Some(bytes)) = (view_location, view_compressed) {
        compressed.push((location.component_name, bytes));
    }
    let mut edits = Vec::new();
    edits
        .try_reserve_exact(compressed.len())
        .map_err(|_allocation| PageLayoutError::Allocation {
            amount: compressed.len(),
        })?;
    for (name, bytes) in &compressed {
        edits.push(EntryEdit::new(name, bytes));
    }
    let mut deleted_previews = Vec::new();
    deleted_previews
        .try_reserve_exact(PREVIEW_ENTRY_NAMES.len())
        .map_err(|_allocation| PageLayoutError::Allocation {
            amount: PREVIEW_ENTRY_NAMES.len(),
        })?;
    for name in PREVIEW_ENTRY_NAMES {
        if source_catalog
            .package()
            .iter()
            .any(|entry| entry.name() == name)
        {
            deleted_previews.push(name);
        }
    }
    let output = source_catalog
        .package()
        .reassemble_with_deletions_to_bytes(&edits, &deleted_previews, source_catalog.limits())
        .map_err(map_archive_error)?;
    let touched_components = compressed.len();
    drop(edits);
    drop(deleted_previews);
    drop(compressed);
    let candidate_source =
        SourceCatalog::from_shared_bytes_with_limits(output.into(), source_catalog.limits())
            .map_err(map_archive_error)?;
    let candidate = Package::from_source_catalog(candidate_source).map_err(map_package_error)?;
    verify_candidate(source, &candidate, after, None, 0)?;
    Ok((candidate, touched_components))
}

fn rewrite_document_component(
    source: &Package,
    component_name: &str,
    before: Layout,
    after: Layout,
    view_location: Option<&ViewStateLocation>,
) -> Result<Vec<u8>, PageLayoutError> {
    let (mut archive, limits) = editable_archive(source, component_name)?;
    {
        let object = archive
            .object_mut(DOCUMENT_IDENTIFIER)
            .ok_or(PageLayoutError::InvalidSource)?;
        let (message_index, message) = unique_message(object, DOCUMENT_MESSAGE_TYPE)?;
        validate_selected_metadata(object, message_index)?;
        if strict_layout(&message.data, wire_limits(source)?)? != before {
            return Err(PageLayoutError::InvalidSource);
        }
        let rewritten = rewrite_layout_payload(&message.data, before, after, wire_limits(source)?)?;
        if strict_layout(&rewritten, wire_limits(source)?)? != after {
            return Err(PageLayoutError::Verification);
        }
        object
            .replace_message_preserving_header_with_limits(
                message_index,
                RawMessage {
                    type_: DOCUMENT_MESSAGE_TYPE,
                    data: rewritten,
                },
                limits,
            )
            .map_err(map_core_error)?;
    }
    if let Some(location) = view_location {
        invalidate_view_state_in_archive(source, &mut archive, location, limits)?;
    }
    compress_archive(&archive, limits)
}

fn rewrite_view_state_component(
    source: &Package,
    location: &ViewStateLocation,
) -> Result<Vec<u8>, PageLayoutError> {
    let (mut archive, limits) = editable_archive(source, &location.component_name)?;
    invalidate_view_state_in_archive(source, &mut archive, location, limits)?;
    compress_archive(&archive, limits)
}

fn invalidate_view_state_in_archive(
    source: &Package,
    archive: &mut Archive,
    location: &ViewStateLocation,
    limits: litchi_iwa_core::Limits,
) -> Result<(), PageLayoutError> {
    let object = archive
        .object_mut(location.object_identifier)
        .ok_or(PageLayoutError::InvalidSource)?;
    let (message_index, message) = unique_message(object, VIEW_STATE_ROOT_MESSAGE_TYPE)?;
    validate_selected_metadata(object, message_index)?;
    let (layout, ui) = strict_view_state_root(&message.data, wire_limits(source)?)?;
    if layout != Some(location.layout_identifier) || ui == layout {
        return Err(PageLayoutError::InvalidSource);
    }
    validate_reference_metadata(
        object,
        message_index,
        location.layout_identifier,
        &[VIEW_STATE_LAYOUT_FIELD],
    )?;
    let rewritten =
        patch_length_delimited_field(&message.data, VIEW_STATE_LAYOUT_FIELD, true, None)
            .map_err(map_wire_error)?;
    if strict_view_state_root(&rewritten, wire_limits(source)?)?
        .0
        .is_some()
    {
        return Err(PageLayoutError::Verification);
    }
    object
        .replace_message_pruning_object_references_preserving_header_with_limits(
            message_index,
            RawMessage {
                type_: VIEW_STATE_ROOT_MESSAGE_TYPE,
                data: rewritten,
            },
            &[location.layout_identifier],
            limits,
        )
        .map_err(map_core_error)?;
    Ok(())
}

fn rewrite_layout_payload(
    source: &[u8],
    before: Layout,
    after: Layout,
    limits: WireLimits,
) -> Result<Vec<u8>, PageLayoutError> {
    let view = WireView::parse_with_limits(source, limits).map_err(map_wire_error)?;
    let mut output_length = source.len();
    let mut output_fields = view.len();
    let mut selected_lengths = [None; 11];
    for field in view.fields() {
        let Some(index) = selected_layout_fields()
            .iter()
            .position(|number| *number == field.number())
        else {
            continue;
        };
        if selected_lengths[index].replace(field.raw().len()).is_some() {
            return Err(PageLayoutError::InvalidSource);
        }
    }
    for (index, field_number) in selected_layout_fields().into_iter().enumerate() {
        let before_value = layout_scalar(before, field_number);
        let after_value = layout_scalar(after, field_number);
        if before_value == after_value {
            continue;
        }
        if let Some(field_length) = selected_lengths[index] {
            output_length = output_length
                .checked_sub(field_length)
                .ok_or(PageLayoutError::InvalidSource)?;
            output_fields = output_fields
                .checked_sub(1)
                .ok_or(PageLayoutError::InvalidSource)?;
        }
        if let Some(value) = after_value {
            output_length = output_length
                .checked_add(encoded_layout_scalar_length(field_number, value))
                .ok_or_else(|| output_limit_error(usize::MAX, limits))?;
            output_fields = output_fields
                .checked_add(1)
                .ok_or(PageLayoutError::InvalidSource)?;
        }
    }
    if output_length > limits.max_output_bytes() {
        return Err(output_limit_error(output_length, limits));
    }
    if output_fields > limits.max_fields() {
        return Err(PageLayoutError::LimitExceeded {
            kind: PageLayoutLimitKind::WireFields,
            observed: usize_to_u64(output_fields),
            maximum: usize_to_u64(limits.max_fields()),
        });
    }
    let work = source
        .len()
        .checked_add(output_length)
        .and_then(|amount| amount.checked_add(view.len()))
        .and_then(|amount| amount.checked_add(view.len()))
        .ok_or_else(|| work_limit_error(usize::MAX, limits))?;
    if work > limits.max_rewrite_work() {
        return Err(work_limit_error(work, limits));
    }

    let mut output = Vec::new();
    output
        .try_reserve_exact(output_length)
        .map_err(|_allocation| PageLayoutError::Allocation {
            amount: output_length,
        })?;
    let mut emitted = [false; 11];
    for field in view.fields() {
        let Some(index) = selected_layout_fields()
            .iter()
            .position(|number| *number == field.number())
        else {
            output.extend_from_slice(field.raw());
            continue;
        };
        emitted[index] = true;
        let before_value = layout_scalar(before, field.number());
        let after_value = layout_scalar(after, field.number());
        if before_value == after_value {
            output.extend_from_slice(field.raw());
        } else if let Some(value) = after_value {
            append_layout_scalar(&mut output, field.number(), value);
        }
    }
    for (index, field_number) in selected_layout_fields().into_iter().enumerate() {
        if !emitted[index]
            && let Some(value) = layout_scalar(after, field_number)
        {
            append_layout_scalar(&mut output, field_number, value);
        }
    }
    if output.len() != output_length {
        return Err(PageLayoutError::Verification);
    }
    Ok(output)
}

const fn selected_layout_fields() -> [u32; 11] {
    [30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 42]
}

fn layout_scalar(layout: Layout, field_number: u32) -> Option<LayoutScalar> {
    let floats = layout_float_values(layout);
    if let Some(index) = FLOAT_FIELDS.iter().position(|field| *field == field_number) {
        return floats[index].map(|value| LayoutScalar::Fixed32(value.to_bits()));
    }
    match field_number {
        VERTICAL_BODY_FIELD => layout
            .lays_out_body_vertically()
            .map(u64::from)
            .map(LayoutScalar::Varint),
        ORIENTATION_FIELD => layout
            .orientation()
            .map(Orientation::as_raw)
            .map(u64::from)
            .map(LayoutScalar::Varint),
        _ => None,
    }
}

fn encoded_layout_scalar_length(field_number: u32, scalar: LayoutScalar) -> usize {
    let wire_type = match scalar {
        LayoutScalar::Fixed32(_) => 5,
        LayoutScalar::Varint(_) => 0,
    };
    let payload = match scalar {
        LayoutScalar::Fixed32(_) => 4,
        LayoutScalar::Varint(raw) => encoded_len(raw),
    };
    encoded_len((u64::from(field_number) << 3) | wire_type).saturating_add(payload)
}

fn append_layout_scalar(output: &mut Vec<u8>, field_number: u32, scalar: LayoutScalar) {
    match scalar {
        LayoutScalar::Fixed32(raw) => {
            encode_varint_into(output, (u64::from(field_number) << 3) | 5);
            output.extend_from_slice(&raw.to_le_bytes());
        },
        LayoutScalar::Varint(raw) => {
            encode_varint_into(output, u64::from(field_number) << 3);
            encode_varint_into(output, raw);
        },
    }
}

fn output_limit_error(observed: usize, limits: WireLimits) -> PageLayoutError {
    PageLayoutError::LimitExceeded {
        kind: PageLayoutLimitKind::OutputBytes,
        observed: usize_to_u64(observed),
        maximum: usize_to_u64(limits.max_output_bytes()),
    }
}

fn work_limit_error(observed: usize, limits: WireLimits) -> PageLayoutError {
    PageLayoutError::LimitExceeded {
        kind: PageLayoutLimitKind::WireWork,
        observed: usize_to_u64(observed),
        maximum: usize_to_u64(limits.max_rewrite_work()),
    }
}

fn layout_float_values(layout: Layout) -> [Option<f32>; 9] {
    [
        layout.page_width(),
        layout.page_height(),
        layout.left_margin(),
        layout.right_margin(),
        layout.top_margin(),
        layout.bottom_margin(),
        layout.header_margin(),
        layout.footer_margin(),
        layout.page_scale(),
    ]
}

fn editable_archive(
    package: &Package,
    component_name: &str,
) -> Result<(Archive, litchi_iwa_core::Limits), PageLayoutError> {
    let source = &package.state.source;
    let entry = source
        .package()
        .iter()
        .find(|entry| entry.name() == component_name)
        .ok_or(PageLayoutError::InvalidSource)?;
    if entry.is_opaque() {
        return Err(PageLayoutError::InvalidSource);
    }
    let limits = source
        .limits()
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let stream = SnappyStream::decompress_with_limits(
        entry.data(),
        source.limits().snappy_limits().map_err(map_archive_error)?,
    )
    .map_err(map_core_error)?;
    let archive = Archive::parse_with_limits(stream.as_bytes(), limits).map_err(map_core_error)?;
    validate_canonical_object_length_prefixes(stream.as_bytes(), &archive)?;
    Ok((archive, limits))
}

fn compress_archive(
    archive: &Archive,
    limits: litchi_iwa_core::Limits,
) -> Result<Vec<u8>, PageLayoutError> {
    let bytes = archive
        .to_bytes_with_limits(limits)
        .map_err(map_core_error)?;
    SnappyStream::compress(&bytes).map_err(map_core_error)
}

fn unique_message(
    object: &ArchiveObject,
    message_type: u32,
) -> Result<(usize, &RawMessage), PageLayoutError> {
    let mut matches = object
        .messages
        .iter()
        .enumerate()
        .filter(|(_index, message)| message.type_ == message_type);
    let result = matches.next().ok_or(PageLayoutError::InvalidSource)?;
    if matches.next().is_some() {
        return Err(PageLayoutError::InvalidSource);
    }
    Ok(result)
}

fn validate_selected_metadata(
    object: &ArchiveObject,
    message_index: usize,
) -> Result<(), PageLayoutError> {
    let message = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(PageLayoutError::InvalidSource)?;
    if object.archive_info.should_merge == Some(true)
        || message.base_message_index.is_some()
        || !message.diff_merge_version.is_empty()
        || message.diff_field_path.is_some()
        || !message.fields_to_remove.is_empty()
        || !message.diff_read_version.is_empty()
    {
        return Err(PageLayoutError::InvalidSource);
    }
    Ok(())
}

fn validate_reference_metadata(
    object: &ArchiveObject,
    message_index: usize,
    identifier: u64,
    accepted_path: &[u32],
) -> Result<(), PageLayoutError> {
    validate_selected_metadata(object, message_index)?;
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(PageLayoutError::InvalidSource)?;
    let aggregate_occurrences = info
        .object_references
        .iter()
        .filter(|candidate| **candidate == identifier)
        .count();
    if aggregate_occurrences != 1 {
        return Err(PageLayoutError::InvalidSource);
    }
    let mut field_declaration_seen = false;
    for field in &info.field_infos {
        let occurrences = field
            .object_references
            .iter()
            .filter(|candidate| **candidate == identifier)
            .count();
        if occurrences == 0 {
            continue;
        }
        if occurrences != 1
            || std::mem::replace(&mut field_declaration_seen, true)
            || field.path.as_slice() != accepted_path
        {
            return Err(PageLayoutError::InvalidSource);
        }
    }
    Ok(())
}

fn validate_canonical_object_length_prefixes(
    source: &[u8],
    archive: &Archive,
) -> Result<(), PageLayoutError> {
    for object in &archive.objects {
        let offset = usize::try_from(object.header_offset)
            .map_err(|_error| PageLayoutError::InvalidSource)?;
        let remaining = source.get(offset..).ok_or(PageLayoutError::InvalidSource)?;
        let (header_bytes, prefix_bytes) =
            decode_varint_from_bytes(remaining).map_err(|_error| PageLayoutError::InvalidSource)?;
        if prefix_bytes != encoded_len(header_bytes) {
            return Err(PageLayoutError::InvalidSource);
        }
        let framed_header_bytes = header_bytes
            .checked_add(
                u64::try_from(prefix_bytes).map_err(|_error| PageLayoutError::InvalidSource)?,
            )
            .ok_or(PageLayoutError::InvalidSource)?;
        if framed_header_bytes != object.header_length
            || object
                .header_offset
                .checked_add(object.header_length)
                .ok_or(PageLayoutError::InvalidSource)?
                != object.data_offset
        {
            return Err(PageLayoutError::InvalidSource);
        }
    }
    Ok(())
}

fn try_owned(source: &str) -> Result<String, PageLayoutError> {
    let mut value = String::new();
    value
        .try_reserve_exact(source.len())
        .map_err(|_allocation| PageLayoutError::Allocation {
            amount: source.len(),
        })?;
    value.push_str(source);
    Ok(value)
}

fn preview_count(package: &Package) -> usize {
    PREVIEW_ENTRY_NAMES
        .iter()
        .filter(|name| {
            package
                .state
                .source
                .package()
                .iter()
                .any(|entry| entry.name() == **name)
        })
        .count()
}

fn verify_candidate(
    source: &Package,
    candidate: &Package,
    expected: Layout,
    expected_layout_state: Option<u64>,
    expected_preview_count: usize,
) -> Result<(), PageLayoutError> {
    if page_layout(candidate)? != expected
        || view_state_layout_identifier(candidate)? != expected_layout_state
        || preview_count(candidate) != expected_preview_count
        || source.stats() != candidate.stats()
        || source.sections().len() != candidate.sections().len()
    {
        return Err(PageLayoutError::Verification);
    }
    for (before, after) in source.sections().iter().zip(candidate.sections()) {
        if before.name() != after.name()
            || before.section_type() != after.section_type()
            || before.heading() != after.heading()
            || before.paragraphs() != after.paragraphs()
            || before.text_storages() != after.text_storages()
            || before.page_count() != after.page_count()
        {
            return Err(PageLayoutError::Verification);
        }
    }
    Ok(())
}

fn wire_limits(package: &Package) -> Result<WireLimits, PageLayoutError> {
    let maximum = package
        .state
        .source
        .limits()
        .effective_archive_limits()
        .map_err(map_archive_error)?
        .max_message_bytes();
    WireLimits::default()
        .with_input_bytes(maximum)
        .and_then(|limits| limits.with_output_bytes(maximum))
        .map_err(map_wire_error)
}

fn map_package_error(error: PackageError) -> PageLayoutError {
    match error {
        PackageError::Archive(archive_error) => map_archive_error(archive_error),
        PackageError::SectionNamesTooLarge { observed, limit } => PageLayoutError::LimitExceeded {
            kind: PageLayoutLimitKind::PayloadBytes,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(limit),
        },
        PackageError::Io(_) | PackageError::InvalidFormat(_) | PackageError::Semantic(_) => {
            PageLayoutError::InvalidSource
        },
    }
}

#[allow(
    clippy::needless_pass_by_value,
    reason = "Result::map_err supplies the owned codec error directly."
)]
fn map_layout_codec_error(error: pages_page_layout_codec::DecodeError) -> PageLayoutError {
    if let Some((observed, maximum)) = error.field_limit_values() {
        return PageLayoutError::LimitExceeded {
            kind: PageLayoutLimitKind::WireFields,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        };
    }
    if let Some((observed, maximum)) = error.work_limit_values() {
        return PageLayoutError::LimitExceeded {
            kind: PageLayoutLimitKind::WireWork,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        };
    }
    match error.wire_resource_limit() {
        Some(WireResourceLimit::Bytes { observed, maximum }) => PageLayoutError::LimitExceeded {
            kind: PageLayoutLimitKind::WireBytes,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        Some(WireResourceLimit::Nesting { observed, maximum }) => PageLayoutError::LimitExceeded {
            kind: PageLayoutLimitKind::WireNesting,
            observed: u64::from(observed),
            maximum: u64::from(maximum),
        },
        _ => PageLayoutError::InvalidSource,
    }
}

fn map_archive_error(error: litchi_iwa_archive::Error) -> PageLayoutError {
    match error {
        litchi_iwa_archive::Error::Limit {
            kind,
            observed,
            maximum,
        } => PageLayoutError::LimitExceeded {
            kind: match kind {
                litchi_iwa_archive::LimitKind::InputBytes => PageLayoutLimitKind::InputBytes,
                litchi_iwa_archive::LimitKind::OutputBytes => PageLayoutLimitKind::OutputBytes,
                litchi_iwa_archive::LimitKind::Entries => PageLayoutLimitKind::Entries,
                litchi_iwa_archive::LimitKind::MemberNameBytes
                | litchi_iwa_archive::LimitKind::MetadataBytes => PageLayoutLimitKind::PackageBytes,
                litchi_iwa_archive::LimitKind::CompressedEntryBytes
                | litchi_iwa_archive::LimitKind::EntryBytes => PageLayoutLimitKind::EntryBytes,
                litchi_iwa_archive::LimitKind::TotalBytes => PageLayoutLimitKind::TotalEntryBytes,
                litchi_iwa_archive::LimitKind::IwaStreamBytes => PageLayoutLimitKind::PayloadBytes,
                litchi_iwa_archive::LimitKind::IwaTotalBytes => {
                    PageLayoutLimitKind::TotalPayloadBytes
                },
            },
            observed,
            maximum,
        },
        litchi_iwa_archive::Error::Allocation { amount, .. } => {
            PageLayoutError::Allocation { amount }
        },
        litchi_iwa_archive::Error::Reassembly(_) => PageLayoutError::UnsupportedSource,
        litchi_iwa_archive::Error::Iwa(core_error) => map_core_error(core_error),
        litchi_iwa_archive::Error::Io(_)
        | litchi_iwa_archive::Error::Zip { .. }
        | litchi_iwa_archive::Error::InvalidLimits(_)
        | litchi_iwa_archive::Error::Encrypted
        | litchi_iwa_archive::Error::SourceChanged { .. }
        | litchi_iwa_archive::Error::DirectoryChanged { .. }
        | litchi_iwa_archive::Error::InvalidBundle(_) => PageLayoutError::InvalidSource,
    }
}

#[allow(
    clippy::needless_pass_by_value,
    reason = "Result::map_err supplies the owned component error directly."
)]
fn map_core_error(error: litchi_iwa_core::Error) -> PageLayoutError {
    match error {
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => PageLayoutError::LimitExceeded {
            kind: match kind {
                litchi_iwa_core::LimitKind::Objects => PageLayoutLimitKind::PayloadObjects,
                litchi_iwa_core::LimitKind::Messages
                | litchi_iwa_core::LimitKind::MessagesPerObject => {
                    PageLayoutLimitKind::PayloadMessages
                },
                litchi_iwa_core::LimitKind::HeaderFields
                | litchi_iwa_core::LimitKind::MetadataItems
                | litchi_iwa_core::LimitKind::SnappyFrames => PageLayoutLimitKind::PayloadItems,
                litchi_iwa_core::LimitKind::HeaderNesting => PageLayoutLimitKind::WireNesting,
                litchi_iwa_core::LimitKind::ArchiveBytes
                | litchi_iwa_core::LimitKind::ObjectBytes
                | litchi_iwa_core::LimitKind::MessageBytes
                | litchi_iwa_core::LimitKind::HeaderBytes
                | litchi_iwa_core::LimitKind::HeaderMemoryBytes
                | litchi_iwa_core::LimitKind::SnappyChunkBytes
                | litchi_iwa_core::LimitKind::SnappyStreamBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedChunkBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedStreamBytes => {
                    PageLayoutLimitKind::PayloadBytes
                },
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            PageLayoutError::Allocation { amount: requested }
        },
        litchi_iwa_core::Error::InvalidArchive { .. }
        | litchi_iwa_core::Error::InvalidLimits { .. }
        | litchi_iwa_core::Error::HeaderCodec { .. }
        | litchi_iwa_core::Error::Io(_)
        | litchi_iwa_core::Error::Snappy { .. } => PageLayoutError::InvalidSource,
    }
}

#[allow(
    clippy::needless_pass_by_value,
    reason = "Result::map_err supplies the owned wire error directly."
)]
fn map_wire_error(error: litchi_iwa_common::Error) -> PageLayoutError {
    match error {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => PageLayoutError::LimitExceeded {
            kind: match kind {
                litchi_iwa_common::LimitKind::InputBytes => PageLayoutLimitKind::WireBytes,
                litchi_iwa_common::LimitKind::OutputBytes => PageLayoutLimitKind::OutputBytes,
                litchi_iwa_common::LimitKind::Fields => PageLayoutLimitKind::WireFields,
                litchi_iwa_common::LimitKind::Nesting => PageLayoutLimitKind::WireNesting,
                litchi_iwa_common::LimitKind::RewriteWork => PageLayoutLimitKind::WireWork,
                litchi_iwa_common::LimitKind::TableRows
                | litchi_iwa_common::LimitKind::TableColumns
                | litchi_iwa_common::LimitKind::TableCells
                | litchi_iwa_common::LimitKind::MaterializedCells => {
                    PageLayoutLimitKind::PayloadItems
                },
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(limit),
        },
        litchi_iwa_common::Error::Allocation { amount, .. } => {
            PageLayoutError::Allocation { amount }
        },
        litchi_iwa_common::Error::InvalidFormat(_)
        | litchi_iwa_common::Error::InvalidLimit { .. } => PageLayoutError::InvalidSource,
    }
}

fn fingerprint(bytes: &[u8]) -> u64 {
    let mut hash = 0xcbf2_9ce4_8422_2325_u64;
    for byte in bytes {
        hash ^= u64::from(*byte);
        hash = hash.wrapping_mul(0x0000_0100_0000_01b3);
    }
    hash
}

fn usize_to_u64(value: usize) -> u64 {
    u64::try_from(value).unwrap_or(u64::MAX)
}
