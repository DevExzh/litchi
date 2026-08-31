//! Exact-source, selector-first Keynote chart-legend visibility transactions.
//!
//! The native chart graph is resolved through the title owner's checked chart
//! selection.  This module owns only the semantic visibility bit and the
//! smallest possible physical rewrite: the generated
//! `ChartNonStyleArchive.tschchartinfodefaultshowlegend` field.  Archive
//! object identifiers and protobuf values never cross the public boundary.

#![allow(
    clippy::cast_sign_loss,
    clippy::map_err_ignore,
    clippy::needless_pass_by_value,
    clippy::wildcard_enum_match_arm,
    reason = "The package boundary redacts native graph failures."
)]

use std::fmt;
use std::sync::Arc;

use litchi_core::Position;
use litchi_iwa_archive::{
    SourceCatalog,
    package::{EntryEdit, ExactArtifacts},
};
use litchi_iwa_common::{
    WireLimits, decode_varint_from_bytes, encode_varint_into,
    varint::encoded_len,
    wire::{WireField, parse_wire_fields_with_limits},
};
use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
use litchi_iwa_protos::keynote_chart_legend_codec::{
    ChartLegendVisibilityWrite, DecodeError as ChartLegendVisibilityDecodeError,
    DecodeLimit as ChartLegendVisibilityDecodeLimit, DecodeOptions,
    WireResourceLimit as ChartLegendVisibilityWireResourceLimit, decode_chart_legend,
    prepare_chart_legend_visibility_rewrite,
};
use thiserror::Error;

use super::slide_chart_title::{ChartSelection, ChartTitleError, select_chart};
use super::{Package, PhysicalSource, ReadError, SemanticLimitKind};
use crate::{ChartSelector, SlideSelector};

const CHART_NON_STYLE_MESSAGE_TYPE: u32 = 5_023;
const GENERATED_CHART_NON_STYLE_EXTENSION_FIELD: u32 = 10_000;
const MAX_LEGEND_CODEC_BYTES: usize = 64 * 1024 * 1024;

/// A finite resource governed by one chart-legend visibility operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum ChartLegendVisibilityLimitKind {
    /// Complete input package bytes.
    InputBytes,
    /// Complete rewritten package or payload bytes.
    OutputBytes,
    /// Bytes in one protobuf payload.
    WireBytes,
    /// ZIP members, IWA objects, or IWA messages.
    Entries,
    /// Logical codec or transaction allocations.
    Allocations,
    /// Bytes in one package member, IWA object, or message.
    EntryBytes,
    /// Aggregate package or IWA bytes.
    TotalBytes,
    /// Semantic slides.
    Slides,
    /// Semantic graph references.
    References,
    /// Parsed protobuf fields.
    WireFields,
    /// Protobuf nesting depth.
    WireNesting,
    /// Aggregate protobuf rewrite work.
    WireWork,
    /// Bytes retained by the selected generated extension.
    LegendBytes,
}

impl fmt::Display for ChartLegendVisibilityLimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::OutputBytes => "output bytes",
            Self::WireBytes => "wire bytes",
            Self::Entries => "entries",
            Self::Allocations => "allocations",
            Self::EntryBytes => "entry bytes",
            Self::TotalBytes => "total bytes",
            Self::Slides => "slides",
            Self::References => "references",
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting",
            Self::WireWork => "wire work",
            Self::LegendBytes => "chart legend bytes",
        })
    }
}

/// A content-redacted failure raised by a chart-legend visibility operation.
#[derive(Debug, Clone, PartialEq, Eq, Error)]
#[non_exhaustive]
pub enum ChartLegendVisibilityError {
    /// The source was not retained as an exact physical package.
    #[error("this Keynote source does not support physical chart-legend edits")]
    UnsupportedSource,
    /// An exact-name slide or chart selector was ambiguous.
    #[error("the Keynote chart-legend selector is ambiguous")]
    AmbiguousSelector,
    /// An exact-name slide selector was empty.
    #[error("the Keynote slide selector name cannot be empty")]
    EmptySlideName,
    /// An exact-name slide selector did not match.
    #[error("the Keynote show has no slide matching the requested name")]
    SlideNameNotFound,
    /// A checked semantic slide position does not exist.
    #[error("the Keynote show has no slide at position {position:?}")]
    SlidePositionNotFound { position: Position },
    /// An exact-name chart selector did not match.
    #[error("the selected Keynote slide has no chart matching the requested name")]
    ChartNameNotFound,
    /// A checked semantic chart position does not exist.
    #[error("the selected Keynote slide has no chart at position {position:?}")]
    ChartPositionNotFound { position: Position },
    /// An empty exact chart name was supplied.
    #[error("the Keynote chart selector name cannot be empty")]
    EmptyChartName,
    /// The selected chart graph or generated extension was malformed.
    #[error("the Keynote chart-legend source cannot be edited safely")]
    InvalidSource,
    /// A finite resource ceiling was exceeded.
    #[error(
        "Keynote chart-legend visibility {kind} limit exceeded: observed {observed}, maximum {maximum}"
    )]
    LimitExceeded {
        /// Resource category that exceeded its limit.
        kind: ChartLegendVisibilityLimitKind,
        /// Observed or requested amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded destination allocation failed before publication.
    #[error("could not allocate {amount} units for the Keynote chart-legend transaction")]
    Allocation { amount: usize },
    /// Full candidate reopening did not reproduce the requested state.
    #[error("the edited Keynote chart legend failed semantic verification")]
    Verification,
    /// The patch does not belong to this exact immutable package artifact.
    #[error("the Keynote chart-legend patch does not match the exact source package")]
    PatchConflict,
}

/// One mutable chart-legend visibility value staged against an immutable
/// package snapshot.
pub struct ChartLegendVisibilityEdit<'a> {
    source: &'a Package,
    slide_position: Position,
    chart_position: Position,
    chart_identifier: u64,
    non_style_identifier: u64,
    slide_identifier: u64,
    before: bool,
    after: bool,
}

impl fmt::Debug for ChartLegendVisibilityEdit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("ChartLegendVisibilityEdit")
            .field("slide_position", &self.slide_position)
            .field("chart_position", &self.chart_position)
            .field("before", &self.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl<'a> ChartLegendVisibilityEdit<'a> {
    fn new<'slide, 'chart>(
        source: &'a Package,
        slide_selector: impl Into<SlideSelector<'slide>>,
        chart_selector: impl Into<ChartSelector<'chart>>,
    ) -> Result<Self, ChartLegendVisibilityError> {
        let selection = select_chart(source, slide_selector.into(), chart_selector.into(), true)
            .map_err(map_chart_title_error)?;
        let before = read_selected_visibility(source, &selection)?;
        Ok(Self {
            source,
            slide_position: selection.slide_position,
            chart_position: selection.chart_position,
            chart_identifier: selection.chart_identifier,
            non_style_identifier: selection.non_style_identifier,
            slide_identifier: selection.slide_identifier,
            before,
            after: before,
        })
    }

    /// Return the selected semantic slide position.
    #[must_use]
    pub const fn slide_position(&self) -> Position {
        self.slide_position
    }

    /// Return the selected semantic chart position within its slide.
    #[must_use]
    pub const fn chart_position(&self) -> Position {
        self.chart_position
    }

    /// Return the visibility observed when this edit began.
    #[must_use]
    pub const fn before(&self) -> bool {
        self.before
    }

    /// Return the visibility staged for publication.
    #[must_use]
    pub const fn after(&self) -> bool {
        self.after
    }

    /// Stage the native legend visibility.
    #[must_use]
    pub const fn set(mut self, visible: bool) -> Self {
        self.after = visible;
        self
    }

    /// Validate and atomically publish the staged immutable candidate.
    pub fn commit(self) -> Result<ChartLegendVisibilityCommit, ChartLegendVisibilityError> {
        let catalog = physical_catalog(self.source)?;
        let source_bytes = catalog.shared_source();
        let source_selection = select_chart(
            self.source,
            SlideSelector::position(self.slide_position),
            ChartSelector::index(self.chart_position.get()),
            true,
        )
        .map_err(map_chart_title_error)?;
        let current = read_selected_visibility(self.source, &source_selection)?;
        if source_selection.chart_identifier != self.chart_identifier
            || source_selection.non_style_identifier != self.non_style_identifier
            || source_selection.slide_identifier != self.slide_identifier
            || current != self.before
        {
            return Err(ChartLegendVisibilityError::InvalidSource);
        }

        if self.before == self.after {
            self.source.validate().map_err(map_read_error)?;
            return Ok(ChartLegendVisibilityCommit {
                package: self.source.snapshot(),
                patch: ChartLegendVisibilityPatch {
                    artifacts: ExactArtifacts::new(Arc::clone(&source_bytes), source_bytes),
                    slide_position: self.slide_position,
                    chart_position: self.chart_position,
                    chart_identifier: self.chart_identifier,
                    non_style_identifier: self.non_style_identifier,
                    slide_identifier: self.slide_identifier,
                    before: self.before,
                    after: self.after,
                    deleted_previews: 0,
                    target_requires_invalidated_previews: false,
                },
                diagnostics: ChartLegendVisibilityDiagnostics::unchanged(),
            });
        }

        if !catalog.source_is_exact() {
            return Err(ChartLegendVisibilityError::UnsupportedSource);
        }
        self.source.validate().map_err(map_read_error)?;
        let (package, deleted_previews) =
            rewrite_chart_legend(self.source, &source_selection, self.after)?;
        package.validate().map_err(map_read_error)?;
        verify_chart_candidate(
            self.source,
            &package,
            self.slide_position,
            self.chart_position,
            self.chart_identifier,
            self.non_style_identifier,
            self.slide_identifier,
            self.after,
            true,
        )?;
        let target = physical_catalog(&package)?.shared_source();
        Ok(ChartLegendVisibilityCommit {
            package,
            patch: ChartLegendVisibilityPatch {
                artifacts: ExactArtifacts::new(source_bytes, Arc::clone(&target)),
                slide_position: self.slide_position,
                chart_position: self.chart_position,
                chart_identifier: self.chart_identifier,
                non_style_identifier: self.non_style_identifier,
                slide_identifier: self.slide_identifier,
                before: self.before,
                after: self.after,
                deleted_previews,
                target_requires_invalidated_previews: true,
            },
            diagnostics: ChartLegendVisibilityDiagnostics::published(deleted_previews),
        })
    }
}

/// An exact-source-checked reversible semantic chart-legend patch.
#[derive(Clone, PartialEq, Eq)]
pub struct ChartLegendVisibilityPatch {
    artifacts: ExactArtifacts,
    slide_position: Position,
    chart_position: Position,
    chart_identifier: u64,
    non_style_identifier: u64,
    slide_identifier: u64,
    before: bool,
    after: bool,
    deleted_previews: usize,
    target_requires_invalidated_previews: bool,
}

impl fmt::Debug for ChartLegendVisibilityPatch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("ChartLegendVisibilityPatch")
            .field("slide_position", &self.slide_position)
            .field("chart_position", &self.chart_position)
            .field("before", &self.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl ChartLegendVisibilityPatch {
    /// Return the selected semantic slide position.
    #[must_use]
    pub const fn slide_position(&self) -> Position {
        self.slide_position
    }

    /// Return the selected semantic chart position within its slide.
    #[must_use]
    pub const fn chart_position(&self) -> Position {
        self.chart_position
    }

    /// Return the visibility required from the source package.
    #[must_use]
    pub const fn before(&self) -> bool {
        self.before
    }

    /// Return the visibility produced by the target package.
    #[must_use]
    pub const fn after(&self) -> bool {
        self.after
    }

    /// Return the base package's compact diagnostic fingerprint.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.artifacts.source_fingerprint()
    }

    /// Return the committed package's compact diagnostic fingerprint.
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.artifacts.target_fingerprint()
    }

    /// Return whether this patch preserves exact source bytes and state.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after && self.artifacts.is_byte_noop()
    }

    /// Return an exact reversible patch from the target back to its source.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            slide_position: self.slide_position,
            chart_position: self.chart_position,
            chart_identifier: self.chart_identifier,
            non_style_identifier: self.non_style_identifier,
            slide_identifier: self.slide_identifier,
            before: self.after,
            after: self.before,
            deleted_previews: 0,
            target_requires_invalidated_previews: self.before != self.after
                && !self.target_requires_invalidated_previews,
        }
    }
}

/// Compact evidence describing one chart-legend visibility commit.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ChartLegendVisibilityDiagnostics {
    changed: bool,
    touched_components: usize,
    deleted_previews: usize,
    full_reparse_performed: bool,
}

impl ChartLegendVisibilityDiagnostics {
    const fn unchanged() -> Self {
        Self {
            changed: false,
            touched_components: 0,
            deleted_previews: 0,
            full_reparse_performed: false,
        }
    }

    const fn published(deleted_previews: usize) -> Self {
        Self {
            changed: true,
            touched_components: 1,
            deleted_previews,
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

    /// Return how many stale root rendering previews were deleted.
    #[must_use]
    pub const fn deleted_previews(self) -> usize {
        self.deleted_previews
    }

    /// Return whether the complete candidate was reopened before publication.
    #[must_use]
    pub const fn full_reparse_performed(self) -> bool {
        self.full_reparse_performed
    }
}

/// The fully verified result of one immutable chart-legend transaction.
#[must_use = "a Keynote chart-legend commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct ChartLegendVisibilityCommit {
    package: Package,
    patch: ChartLegendVisibilityPatch,
    diagnostics: ChartLegendVisibilityDiagnostics,
}

impl ChartLegendVisibilityCommit {
    /// Borrow the fully reopened immutable package snapshot.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Consume this commit and return its immutable package snapshot.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Borrow the reversible exact-source patch.
    #[must_use]
    pub const fn patch(&self) -> &ChartLegendVisibilityPatch {
        &self.patch
    }

    /// Borrow compact publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &ChartLegendVisibilityDiagnostics {
        &self.diagnostics
    }
}

impl Package {
    /// Read whether the selected Keynote chart displays its legend.
    ///
    /// An absent native legend field has the effective value `false`; the
    /// distinction between absent and explicit false is retained by the
    /// source bytes and therefore remains a no-op when writing `false`.
    pub fn slide_chart_legend_visible<'slide, 'chart>(
        &self,
        slide_selector: impl Into<SlideSelector<'slide>>,
        chart_selector: impl Into<ChartSelector<'chart>>,
    ) -> Result<bool, ChartLegendVisibilityError> {
        let selection = select_chart(self, slide_selector.into(), chart_selector.into(), true)
            .map_err(map_chart_title_error)?;
        read_selected_visibility(self, &selection)
    }

    /// Start an exact immutable edit of one selected chart's legend.
    pub fn edit_slide_chart_legend<'slide, 'chart>(
        &self,
        slide_selector: impl Into<SlideSelector<'slide>>,
        chart_selector: impl Into<ChartSelector<'chart>>,
    ) -> Result<ChartLegendVisibilityEdit<'_>, ChartLegendVisibilityError> {
        ChartLegendVisibilityEdit::new(self, slide_selector, chart_selector)
    }

    /// Apply an exact-source-checked chart-legend patch.
    pub fn apply_slide_chart_legend(
        &self,
        patch: &ChartLegendVisibilityPatch,
    ) -> Result<ChartLegendVisibilityCommit, ChartLegendVisibilityError> {
        let catalog = physical_catalog(self)?;
        let source = catalog.shared_source();
        if !patch.artifacts.authorizes_source(&source) {
            return Err(ChartLegendVisibilityError::PatchConflict);
        }
        let selection = select_chart(
            self,
            SlideSelector::position(patch.slide_position),
            ChartSelector::index(patch.chart_position.get()),
            true,
        )
        .map_err(map_chart_title_error)?;
        let current = read_selected_visibility(self, &selection)?;
        if selection.chart_identifier != patch.chart_identifier
            || selection.non_style_identifier != patch.non_style_identifier
            || selection.slide_identifier != patch.slide_identifier
            || current != patch.before
        {
            return Err(ChartLegendVisibilityError::PatchConflict);
        }
        if patch.is_noop() {
            self.validate().map_err(map_read_error)?;
            return Ok(ChartLegendVisibilityCommit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: ChartLegendVisibilityDiagnostics::unchanged(),
            });
        }
        if !catalog.source_is_exact() {
            return Err(ChartLegendVisibilityError::PatchConflict);
        }
        let candidate =
            Package::from_source_with_options(patch.artifacts.target(), self.state.options)
                .map_err(map_read_error)?;
        candidate.validate().map_err(map_read_error)?;
        verify_chart_candidate(
            self,
            &candidate,
            patch.slide_position,
            patch.chart_position,
            patch.chart_identifier,
            patch.non_style_identifier,
            patch.slide_identifier,
            patch.after,
            patch.target_requires_invalidated_previews,
        )?;
        Ok(ChartLegendVisibilityCommit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: ChartLegendVisibilityDiagnostics::published(patch.deleted_previews),
        })
    }
}

fn read_selected_visibility(
    package: &Package,
    selection: &ChartSelection,
) -> Result<bool, ChartLegendVisibilityError> {
    let (_, object) = package
        .object_with_component(selection.non_style_identifier)
        .ok_or(ChartLegendVisibilityError::InvalidSource)?;
    let message = exactly_one_message(object, CHART_NON_STYLE_MESSAGE_TYPE)?;
    read_chart_legend_visibility(
        message.data.as_slice(),
        package.wire_limits().map_err(map_wire_error)?,
    )
}

fn read_chart_legend_visibility(
    data: &[u8],
    limits: WireLimits,
) -> Result<bool, ChartLegendVisibilityError> {
    let fields = parse_wire_fields_with_limits(data, limits).map_err(map_wire_error)?;
    let mut extension = None;
    for field in fields
        .iter()
        .copied()
        .filter(|field| field.number() == GENERATED_CHART_NON_STYLE_EXTENSION_FIELD)
    {
        if extension.is_some() || field.wire_type() != 2 {
            return Err(ChartLegendVisibilityError::InvalidSource);
        }
        field
            .validate_canonical_framing(data)
            .map_err(map_wire_error)?;
        extension = Some(field.payload(data).map_err(map_wire_error)?);
    }
    let Some(extension) = extension else {
        return Ok(false);
    };
    let snapshot = decode_chart_legend(extension, legend_decode_options(extension, limits)?)
        .map_err(map_legend_codec_error)?;
    Ok(snapshot.is_visible())
}

fn rewrite_chart_legend(
    source: &Package,
    selection: &ChartSelection,
    visible: bool,
) -> Result<(Package, usize), ChartLegendVisibilityError> {
    let catalog = physical_catalog(source)?;
    let entry = catalog
        .package()
        .iter()
        .find(|entry| entry.name() == selection.non_style_component_name)
        .ok_or(ChartLegendVisibilityError::InvalidSource)?;
    if entry.is_opaque() {
        return Err(ChartLegendVisibilityError::InvalidSource);
    }
    let archive_limits = source
        .state
        .options
        .archive()
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let stream = SnappyStream::decompress_with_limits(
        entry.data(),
        source
            .state
            .options
            .archive()
            .snappy_limits()
            .map_err(map_archive_error)?,
    )
    .map_err(map_core_error)?;
    let mut archive =
        Archive::parse_with_limits(stream.as_bytes(), archive_limits).map_err(map_core_error)?;
    validate_canonical_object_length_prefixes(stream.as_bytes(), &archive)?;
    let (message_index, message_data) = {
        let object = archive
            .object(selection.non_style_identifier)
            .ok_or(ChartLegendVisibilityError::InvalidSource)?;
        let mut index = None;
        for (candidate, message) in object.messages.iter().enumerate() {
            if message.type_ != CHART_NON_STYLE_MESSAGE_TYPE {
                continue;
            }
            if index.replace(candidate).is_some() {
                return Err(ChartLegendVisibilityError::InvalidSource);
            }
        }
        let index = index.ok_or(ChartLegendVisibilityError::InvalidSource)?;
        (index, object.messages[index].data.as_slice())
    };
    let limits = source.wire_limits().map_err(map_wire_error)?;
    // Codec preparation is source-borrowed and performs all strict validation
    // and resource preflight before the archive ever needs a second owned
    // copy of the selected message. The returned bytes own only the rewrite
    // candidate, so the source payload remains in the archive until mutation.
    let patched = patch_chart_non_style_visibility(message_data, visible, limits)?;
    archive
        .object_mut(selection.non_style_identifier)
        .ok_or(ChartLegendVisibilityError::InvalidSource)?
        .replace_message_preserving_header_with_limits(
            message_index,
            RawMessage {
                type_: CHART_NON_STYLE_MESSAGE_TYPE,
                data: patched,
            },
            archive_limits,
        )
        .map_err(map_core_error)?;
    let bytes = archive
        .to_bytes_with_limits(archive_limits)
        .map_err(map_core_error)?;
    let compressed = SnappyStream::compress(&bytes).map_err(map_core_error)?;
    let edit = EntryEdit::new(
        selection.non_style_component_name.as_str(),
        compressed.as_slice(),
    );
    let previews = super::rendering_invalidation::root_preview_deletions(catalog.package())
        .map_err(map_rendering_error)?;
    let edits = [edit];
    let prepared = catalog
        .package()
        .prepare_reassembly_with_deletions(&edits, previews.names(), source.state.options.archive())
        .map_err(map_archive_error)?;
    let execution_limits = prepared.execution_requirements().exact_limits();
    let output = prepared
        .execute(execution_limits)
        .map_err(map_archive_error)?;
    let package = Package::from_source_with_options(output.into(), source.state.options)
        .map_err(map_read_error)?;
    Ok((package, previews.len()))
}

fn patch_chart_non_style_visibility(
    data: &[u8],
    visible: bool,
    limits: WireLimits,
) -> Result<Vec<u8>, ChartLegendVisibilityError> {
    let fields = parse_wire_fields_with_limits(data, limits).map_err(map_wire_error)?;
    let extension_field = fields
        .iter()
        .copied()
        .find(|field| field.number() == GENERATED_CHART_NON_STYLE_EXTENSION_FIELD);
    if fields
        .iter()
        .filter(|field| field.number() == GENERATED_CHART_NON_STYLE_EXTENSION_FIELD)
        .count()
        > 1
    {
        return Err(ChartLegendVisibilityError::InvalidSource);
    }
    if let Some(field) = extension_field {
        if field.wire_type() != 2 {
            return Err(ChartLegendVisibilityError::InvalidSource);
        }
        field
            .validate_canonical_framing(data)
            .map_err(map_wire_error)?;
    }
    let extension = extension_field
        .map(|field| field.payload(data).map_err(map_wire_error))
        .transpose()?;
    let extension = extension.unwrap_or_default();
    let options = legend_decode_options(extension, limits)?;
    let prepared = prepare_chart_legend_visibility_rewrite(
        extension,
        ChartLegendVisibilityWrite::new(visible),
        options,
    )
    .map_err(map_legend_codec_error)?;
    let replacement = prepared
        .execute(prepared.execution_requirements().exact())
        .map_err(map_legend_codec_error)?
        .into_output();
    rewrite_chart_non_style_extension(data, extension_field, &replacement, limits)
}

fn rewrite_chart_non_style_extension(
    data: &[u8],
    extension_field: Option<WireField>,
    replacement: &[u8],
    limits: WireLimits,
) -> Result<Vec<u8>, ChartLegendVisibilityError> {
    let replacement_length =
        u64::try_from(replacement.len()).map_err(|_| ChartLegendVisibilityError::InvalidSource)?;
    let key_length = extension_field.map_or_else(
        || encoded_len((u64::from(GENERATED_CHART_NON_STYLE_EXTENSION_FIELD) << 3) | 2),
        |field| field.key_end() - field.start(),
    );
    let replacement_field_length = key_length
        .checked_add(encoded_len(replacement_length))
        .and_then(|length| length.checked_add(replacement.len()))
        .ok_or(ChartLegendVisibilityError::InvalidSource)?;
    let output_length = extension_field
        .map_or_else(
            || data.len().checked_add(replacement_field_length),
            |field| {
                data.len()
                    .checked_sub(field.end() - field.start())
                    .and_then(|length| length.checked_add(replacement_field_length))
            },
        )
        .ok_or(ChartLegendVisibilityError::InvalidSource)?;
    if output_length > limits.max_output_bytes() {
        return Err(ChartLegendVisibilityError::LimitExceeded {
            kind: ChartLegendVisibilityLimitKind::OutputBytes,
            observed: usize_to_u64(output_length),
            maximum: usize_to_u64(limits.max_output_bytes()),
        });
    }
    let mut output = Vec::new();
    output.try_reserve_exact(output_length).map_err(|_| {
        ChartLegendVisibilityError::Allocation {
            amount: output_length,
        }
    })?;
    match extension_field {
        Some(field) => {
            output.extend_from_slice(&data[..field.start()]);
            output.extend_from_slice(&data[field.start()..field.key_end()]);
            encode_varint_into(&mut output, replacement_length);
            output.extend_from_slice(replacement);
            output.extend_from_slice(&data[field.end()..]);
        },
        None => {
            output.extend_from_slice(data);
            encode_varint_into(
                &mut output,
                (u64::from(GENERATED_CHART_NON_STYLE_EXTENSION_FIELD) << 3) | 2,
            );
            encode_varint_into(&mut output, replacement_length);
            output.extend_from_slice(replacement);
        },
    }
    debug_assert_eq!(output.len(), output_length);
    Ok(output)
}

fn verify_chart_candidate(
    source: &Package,
    candidate: &Package,
    slide_position: Position,
    chart_position: Position,
    chart_identifier: u64,
    non_style_identifier: u64,
    slide_identifier: u64,
    expected_visibility: bool,
    require_invalidated_previews: bool,
) -> Result<(), ChartLegendVisibilityError> {
    if source.state.total_objects != candidate.state.total_objects {
        return Err(ChartLegendVisibilityError::Verification);
    }
    let source_selection = select_chart(
        source,
        SlideSelector::position(slide_position),
        ChartSelector::index(chart_position.get()),
        true,
    )
    .map_err(map_chart_title_error)?;
    let candidate_selection = select_chart(
        candidate,
        SlideSelector::position(slide_position),
        ChartSelector::index(chart_position.get()),
        true,
    )
    .map_err(map_chart_title_error)?;
    if source_selection.chart_identifier != chart_identifier
        || source_selection.non_style_identifier != non_style_identifier
        || source_selection.slide_identifier != slide_identifier
        || candidate_selection.chart_identifier != chart_identifier
        || candidate_selection.non_style_identifier != non_style_identifier
        || candidate_selection.slide_identifier != slide_identifier
        || source_selection.title != candidate_selection.title
    {
        return Err(ChartLegendVisibilityError::Verification);
    }
    if read_selected_visibility(candidate, &candidate_selection)? != expected_visibility {
        return Err(ChartLegendVisibilityError::Verification);
    }
    if source.show().map_err(map_read_error)? != candidate.show().map_err(map_read_error)? {
        return Err(ChartLegendVisibilityError::Verification);
    }
    verify_chart_package_locality(
        source,
        candidate,
        &source_selection.non_style_component_name,
        non_style_identifier,
        require_invalidated_previews,
    )
}

fn verify_chart_package_locality(
    source: &Package,
    candidate: &Package,
    selected_component_name: &str,
    selected_identifier: u64,
    require_invalidated_previews: bool,
) -> Result<(), ChartLegendVisibilityError> {
    let source_catalog = physical_catalog(source)?;
    let candidate_catalog = physical_catalog(candidate)?;
    let source_previews =
        super::rendering_invalidation::root_preview_deletions(source_catalog.package())
            .map_err(map_rendering_error)?;
    let candidate_previews =
        super::rendering_invalidation::root_preview_deletions(candidate_catalog.package())
            .map_err(map_rendering_error)?;
    if require_invalidated_previews && !candidate_previews.names().is_empty() {
        return Err(ChartLegendVisibilityError::Verification);
    }

    // Reassembly can remove the three exact root preview members and relocate
    // retained central-directory records. Every other member must stay in
    // source order with identical payload, metadata, and raw records.
    let mut source_entries = source_catalog
        .package()
        .iter()
        .filter(|entry| !source_previews.names().contains(&entry.name()));
    let mut candidate_entries = candidate_catalog
        .package()
        .iter()
        .filter(|entry| !candidate_previews.names().contains(&entry.name()));
    loop {
        match (source_entries.next(), candidate_entries.next()) {
            (Some(source_entry), Some(candidate_entry)) => {
                let selected = source_entry.name() == selected_component_name;
                if source_entry.name() != candidate_entry.name()
                    || (!selected && !same_unselected_package_entry(source_entry, candidate_entry))
                    || (selected && !same_selected_package_entry(source_entry, candidate_entry))
                {
                    return Err(ChartLegendVisibilityError::Verification);
                }
            },
            (None, None) => break,
            _ => return Err(ChartLegendVisibilityError::Verification),
        }
    }

    let source_components = source_catalog.components();
    let candidate_components = candidate_catalog.components();
    if source_components.len() != candidate_components.len() {
        return Err(ChartLegendVisibilityError::Verification);
    }
    let archive_limits = source
        .state
        .options
        .archive()
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let mut selected_component_seen = false;
    for (source_component, candidate_component) in
        source_components.iter().zip(candidate_components.iter())
    {
        if source_component.name() != candidate_component.name()
            || source_component.archive().objects.len()
                != candidate_component.archive().objects.len()
        {
            return Err(ChartLegendVisibilityError::Verification);
        }
        if source_component.name() != selected_component_name {
            for (source_object, candidate_object) in source_component
                .archive()
                .objects
                .iter()
                .zip(&candidate_component.archive().objects)
            {
                if !source_object.same_content_ignoring_offsets(candidate_object) {
                    return Err(ChartLegendVisibilityError::Verification);
                }
            }
            continue;
        }
        selected_component_seen = true;
        verify_selected_chart_component(
            source_component.archive(),
            candidate_component.archive(),
            selected_identifier,
            archive_limits,
        )?;
    }
    if !selected_component_seen {
        return Err(ChartLegendVisibilityError::Verification);
    }
    Ok(())
}

fn verify_selected_chart_component(
    source: &Archive,
    candidate: &Archive,
    selected_identifier: u64,
    archive_limits: litchi_iwa_core::ArchiveLimits,
) -> Result<(), ChartLegendVisibilityError> {
    let mut selected_seen = false;
    for (source_object, candidate_object) in source.objects.iter().zip(&candidate.objects) {
        if source_object.archive_info.identifier != candidate_object.archive_info.identifier {
            return Err(ChartLegendVisibilityError::Verification);
        }
        let identifier = source_object
            .archive_info
            .identifier
            .ok_or(ChartLegendVisibilityError::Verification)?;
        if identifier != selected_identifier {
            if !source_object.same_content_ignoring_offsets(candidate_object) {
                return Err(ChartLegendVisibilityError::Verification);
            }
            continue;
        }
        if std::mem::replace(&mut selected_seen, true) {
            return Err(ChartLegendVisibilityError::Verification);
        }
        let selected_message_index =
            unique_message_index(source_object, CHART_NON_STYLE_MESSAGE_TYPE)?;
        verify_selected_chart_object(
            source_object,
            candidate_object,
            selected_message_index,
            archive_limits,
        )?;
    }
    if !selected_seen {
        return Err(ChartLegendVisibilityError::Verification);
    }
    Ok(())
}

fn verify_selected_chart_object(
    source: &ArchiveObject,
    candidate: &ArchiveObject,
    selected_message_index: usize,
    archive_limits: litchi_iwa_core::ArchiveLimits,
) -> Result<(), ChartLegendVisibilityError> {
    if source.archive_info.identifier != candidate.archive_info.identifier
        || source.archive_info.should_merge != candidate.archive_info.should_merge
        || source.messages.len() != candidate.messages.len()
        || source.archive_info.message_infos.len() != candidate.archive_info.message_infos.len()
    {
        return Err(ChartLegendVisibilityError::Verification);
    }
    for (index, ((source_message, candidate_message), (source_info, candidate_info))) in source
        .messages
        .iter()
        .zip(&candidate.messages)
        .zip(
            source
                .archive_info
                .message_infos
                .iter()
                .zip(&candidate.archive_info.message_infos),
        )
        .enumerate()
    {
        if index == selected_message_index {
            if source_message.type_ != candidate_message.type_
                || !message_info_equal_except_length(source_info, candidate_info)
            {
                return Err(ChartLegendVisibilityError::Verification);
            }
        } else if source_message != candidate_message || source_info != candidate_info {
            return Err(ChartLegendVisibilityError::Verification);
        }
    }

    // Recreate the selected object using the candidate message. This checks
    // retained unknown ArchiveInfo/header bytes in addition to the decoded
    // message metadata above, while allowing only its length/framing update.
    let candidate_message = candidate
        .messages
        .get(selected_message_index)
        .ok_or(ChartLegendVisibilityError::Verification)?
        .clone();
    let mut expected = source.clone();
    expected
        .replace_message_preserving_header_with_limits(
            selected_message_index,
            candidate_message,
            archive_limits,
        )
        .map_err(map_core_error)?;
    expected.header_length = candidate.header_length;
    expected.data_length = candidate.data_length;
    if !expected.same_content_ignoring_offsets(candidate) {
        return Err(ChartLegendVisibilityError::Verification);
    }
    Ok(())
}

fn unique_message_index(
    object: &ArchiveObject,
    message_type: u32,
) -> Result<usize, ChartLegendVisibilityError> {
    let mut selected = None;
    for (index, message) in object.messages.iter().enumerate() {
        if message.type_ != message_type {
            continue;
        }
        if selected.replace(index).is_some() {
            return Err(ChartLegendVisibilityError::Verification);
        }
    }
    selected.ok_or(ChartLegendVisibilityError::Verification)
}

fn same_unselected_package_entry(
    source: &litchi_iwa_archive::package::Entry,
    candidate: &litchi_iwa_archive::package::Entry,
) -> bool {
    source.name() == candidate.name()
        && source.raw_name() == candidate.raw_name()
        && source.is_opaque() == candidate.is_opaque()
        && source.data() == candidate.data()
        && source.metadata() == candidate.metadata()
        && source.raw_record().local_record() == candidate.raw_record().local_record()
        && source.raw_record().compressed_data() == candidate.raw_record().compressed_data()
        && same_central_record_except_offset(
            source.raw_record().central_directory_record(),
            candidate.raw_record().central_directory_record(),
        )
}

fn same_selected_package_entry(
    source: &litchi_iwa_archive::package::Entry,
    candidate: &litchi_iwa_archive::package::Entry,
) -> bool {
    if source.name() != candidate.name()
        || source.raw_name() != candidate.raw_name()
        || source.is_opaque() != candidate.is_opaque()
        || !same_header_metadata(source.metadata().local(), candidate.metadata().local())
        || !same_header_metadata(source.metadata().central(), candidate.metadata().central())
    {
        return false;
    }
    compatible_local_zip_record(source, candidate)
        && compatible_central_zip_record(source, candidate)
}

fn same_header_metadata(
    source: &litchi_iwa_archive::package::HeaderMetadata,
    candidate: &litchi_iwa_archive::package::HeaderMetadata,
) -> bool {
    source.version_needed() == candidate.version_needed()
        && source.flags() == candidate.flags()
        && source.compression_method() == candidate.compression_method()
        && source.last_modified() == candidate.last_modified()
        && source.name() == candidate.name()
        && source.extra() == candidate.extra()
        && source.comment() == candidate.comment()
}

fn compatible_local_zip_record(
    source: &litchi_iwa_archive::package::Entry,
    candidate: &litchi_iwa_archive::package::Entry,
) -> bool {
    let source_record = source.raw_record().local_record();
    let candidate_record = candidate.raw_record().local_record();
    if source_record.len() < 30 || candidate_record.len() < 30 {
        return false;
    }
    let Some(source_name_len) = read_u16(source_record, 26) else {
        return false;
    };
    let Some(source_extra_len) = read_u16(source_record, 28) else {
        return false;
    };
    let Some(candidate_name_len) = read_u16(candidate_record, 26) else {
        return false;
    };
    let Some(candidate_extra_len) = read_u16(candidate_record, 28) else {
        return false;
    };
    let source_header_len = 30usize
        .checked_add(usize::from(source_name_len))
        .and_then(|length| length.checked_add(usize::from(source_extra_len)));
    let candidate_header_len = 30usize
        .checked_add(usize::from(candidate_name_len))
        .and_then(|length| length.checked_add(usize::from(candidate_extra_len)));
    let (Some(source_header_len), Some(candidate_header_len)) =
        (source_header_len, candidate_header_len)
    else {
        return false;
    };
    if source_header_len != candidate_header_len
        || source_header_len > source_record.len()
        || candidate_header_len > candidate_record.len()
    {
        return false;
    }
    let Some(source_suffix_start) =
        source_header_len.checked_add(source.raw_record().compressed_data().len())
    else {
        return false;
    };
    let Some(candidate_suffix_start) =
        candidate_header_len.checked_add(candidate.raw_record().compressed_data().len())
    else {
        return false;
    };
    if source_suffix_start > source_record.len() || candidate_suffix_start > candidate_record.len()
    {
        return false;
    }
    // CRC and the two size words are the only local-header values that may
    // change for an edited member; names, flags, extras, and timestamps must
    // remain source-authoritative.
    if !equal_except_ranges(
        &source_record[..source_header_len],
        &candidate_record[..candidate_header_len],
        std::slice::from_ref(&(14..26)),
    ) {
        return false;
    }
    descriptor_shape_compatible(
        source.metadata().local().flags(),
        &source_record[source_suffix_start..],
        &candidate_record[candidate_suffix_start..],
    )
}

fn compatible_central_zip_record(
    source: &litchi_iwa_archive::package::Entry,
    candidate: &litchi_iwa_archive::package::Entry,
) -> bool {
    let source_record = source.raw_record().central_directory_record();
    let candidate_record = candidate.raw_record().central_directory_record();
    if source_record.len() < 46
        || candidate_record.len() != source_record.len()
        || read_u16(source_record, 28) != read_u16(candidate_record, 28)
        || read_u16(source_record, 30) != read_u16(candidate_record, 30)
        || read_u16(source_record, 32) != read_u16(candidate_record, 32)
    {
        return false;
    }
    // CRC, compressed/uncompressed sizes, and the local-record offset are
    // rewritten by ZIP reassembly. All names, extras, comments, attributes,
    // timestamps, and ordering fields stay exact.
    equal_except_ranges(source_record, candidate_record, &[(16..28), (42..46)])
}

fn descriptor_shape_compatible(flags: u16, source: &[u8], candidate: &[u8]) -> bool {
    if flags & 0x0008 == 0 {
        return source.is_empty() && candidate.is_empty();
    }
    if source.len() != candidate.len() || !matches!(source.len(), 12 | 16) {
        return false;
    }
    if source.len() == 16 {
        source[..4] == candidate[..4]
            && read_u32(source, 0) == Some(0x0807_4b50)
            && read_u32(candidate, 0) == Some(0x0807_4b50)
    } else {
        true
    }
}

fn equal_except_ranges(
    source: &[u8],
    candidate: &[u8],
    ignored: &[std::ops::Range<usize>],
) -> bool {
    source.len() == candidate.len()
        && source.iter().enumerate().all(|(index, byte)| {
            ignored.iter().any(|range| range.contains(&index)) || Some(byte) == candidate.get(index)
        })
}

fn read_u16(bytes: &[u8], offset: usize) -> Option<u16> {
    let value = bytes.get(offset..offset.checked_add(2)?)?;
    Some(u16::from_le_bytes([value[0], value[1]]))
}

fn read_u32(bytes: &[u8], offset: usize) -> Option<u32> {
    let value = bytes.get(offset..offset.checked_add(4)?)?;
    Some(u32::from_le_bytes([value[0], value[1], value[2], value[3]]))
}

fn same_central_record_except_offset(source: &[u8], candidate: &[u8]) -> bool {
    const LOCAL_HEADER_OFFSET: std::ops::Range<usize> = 42..46;
    source.len() == candidate.len()
        && source.len() >= LOCAL_HEADER_OFFSET.end
        && source[..LOCAL_HEADER_OFFSET.start] == candidate[..LOCAL_HEADER_OFFSET.start]
        && source[LOCAL_HEADER_OFFSET.end..] == candidate[LOCAL_HEADER_OFFSET.end..]
}

fn message_info_equal_except_length(
    source: &litchi_iwa_core::MessageInfo,
    candidate: &litchi_iwa_core::MessageInfo,
) -> bool {
    source.type_ == candidate.type_
        && source.versions == candidate.versions
        && source.field_infos == candidate.field_infos
        && source.object_references == candidate.object_references
        && source.data_references == candidate.data_references
        && source.base_message_index == candidate.base_message_index
        && source.diff_merge_version == candidate.diff_merge_version
        && source.diff_field_path == candidate.diff_field_path
        && source.fields_to_remove == candidate.fields_to_remove
        && source.diff_read_version == candidate.diff_read_version
}

fn exactly_one_message(
    object: &ArchiveObject,
    message_type: u32,
) -> Result<&RawMessage, ChartLegendVisibilityError> {
    let mut selected = None;
    for message in &object.messages {
        if message.type_ != message_type {
            continue;
        }
        if selected.replace(message).is_some() {
            return Err(ChartLegendVisibilityError::InvalidSource);
        }
    }
    selected.ok_or(ChartLegendVisibilityError::InvalidSource)
}

fn validate_canonical_object_length_prefixes(
    source: &[u8],
    archive: &Archive,
) -> Result<(), ChartLegendVisibilityError> {
    for object in &archive.objects {
        let offset = usize::try_from(object.header_offset)
            .map_err(|_| ChartLegendVisibilityError::InvalidSource)?;
        let remaining = source
            .get(offset..)
            .ok_or(ChartLegendVisibilityError::InvalidSource)?;
        let (header_bytes, prefix_bytes) = decode_varint_from_bytes(remaining)
            .map_err(|_| ChartLegendVisibilityError::InvalidSource)?;
        if prefix_bytes != encoded_len(header_bytes) {
            return Err(ChartLegendVisibilityError::InvalidSource);
        }
        let framed_header_bytes = header_bytes
            .checked_add(
                u64::try_from(prefix_bytes)
                    .map_err(|_| ChartLegendVisibilityError::InvalidSource)?,
            )
            .ok_or(ChartLegendVisibilityError::InvalidSource)?;
        if framed_header_bytes != object.header_length
            || object
                .header_offset
                .checked_add(object.header_length)
                .ok_or(ChartLegendVisibilityError::InvalidSource)?
                != object.data_offset
        {
            return Err(ChartLegendVisibilityError::InvalidSource);
        }
    }
    Ok(())
}

fn legend_decode_options(
    source: &[u8],
    limits: WireLimits,
) -> Result<DecodeOptions, ChartLegendVisibilityError> {
    let recursion_limit = u32::try_from(limits.max_nesting()).map_err(|_| {
        ChartLegendVisibilityError::LimitExceeded {
            kind: ChartLegendVisibilityLimitKind::WireNesting,
            observed: usize_to_u64(limits.max_nesting()),
            maximum: u64::from(u32::MAX),
        }
    })?;
    let source_limit = source.len().max(1);
    let codec_input = source_limit.min(MAX_LEGEND_CODEC_BYTES);
    let codec_output = limits
        .max_output_bytes()
        .min(MAX_LEGEND_CODEC_BYTES)
        .max(codec_input);
    let codec_fields = limits.max_fields().clamp(1, MAX_LEGEND_CODEC_BYTES);
    let codec_work = limits.max_rewrite_work().clamp(1, MAX_LEGEND_CODEC_BYTES);
    let codec_retained = codec_input
        .checked_add(codec_output)
        .unwrap_or(MAX_LEGEND_CODEC_BYTES)
        .min(MAX_LEGEND_CODEC_BYTES);
    Ok(
        DecodeOptions::new(codec_input, codec_fields, codec_work, recursion_limit)
            .with_max_output_bytes(codec_output)
            .with_max_allocations(1)
            .with_max_retained_bytes(codec_retained.max(codec_input))
            .with_max_scratch_bytes(codec_output.max(1)),
    )
}

fn physical_catalog(package: &Package) -> Result<&SourceCatalog, ChartLegendVisibilityError> {
    match &package.state.source {
        PhysicalSource::Package(source) => Ok(source),
        PhysicalSource::Semantic(_) => Err(ChartLegendVisibilityError::UnsupportedSource),
    }
}

fn map_chart_title_error(error: ChartTitleError) -> ChartLegendVisibilityError {
    match error {
        ChartTitleError::UnsupportedSource => ChartLegendVisibilityError::UnsupportedSource,
        ChartTitleError::AmbiguousSelector => ChartLegendVisibilityError::AmbiguousSelector,
        ChartTitleError::EmptySlideName => ChartLegendVisibilityError::EmptySlideName,
        ChartTitleError::SlideNameNotFound => ChartLegendVisibilityError::SlideNameNotFound,
        ChartTitleError::SlidePositionNotFound { position } => {
            ChartLegendVisibilityError::SlidePositionNotFound { position }
        },
        ChartTitleError::ChartNameNotFound => ChartLegendVisibilityError::ChartNameNotFound,
        ChartTitleError::ChartPositionNotFound { position } => {
            ChartLegendVisibilityError::ChartPositionNotFound { position }
        },
        ChartTitleError::EmptyChartName => ChartLegendVisibilityError::EmptyChartName,
        ChartTitleError::LimitExceeded {
            kind,
            observed,
            maximum,
        } => ChartLegendVisibilityError::LimitExceeded {
            kind: map_chart_title_limit(kind),
            observed,
            maximum,
        },
        ChartTitleError::Allocation { amount } => ChartLegendVisibilityError::Allocation { amount },
        _ => ChartLegendVisibilityError::InvalidSource,
    }
}

fn map_chart_title_limit(
    kind: super::slide_chart_title::ChartTitleLimitKind,
) -> ChartLegendVisibilityLimitKind {
    use super::slide_chart_title::ChartTitleLimitKind;
    match kind {
        ChartTitleLimitKind::InputBytes => ChartLegendVisibilityLimitKind::InputBytes,
        ChartTitleLimitKind::OutputBytes => ChartLegendVisibilityLimitKind::OutputBytes,
        ChartTitleLimitKind::WireBytes => ChartLegendVisibilityLimitKind::WireBytes,
        ChartTitleLimitKind::Entries => ChartLegendVisibilityLimitKind::Entries,
        ChartTitleLimitKind::EntryBytes => ChartLegendVisibilityLimitKind::EntryBytes,
        ChartTitleLimitKind::TotalBytes => ChartLegendVisibilityLimitKind::TotalBytes,
        ChartTitleLimitKind::Slides => ChartLegendVisibilityLimitKind::Slides,
        ChartTitleLimitKind::References => ChartLegendVisibilityLimitKind::References,
        ChartTitleLimitKind::WireFields => ChartLegendVisibilityLimitKind::WireFields,
        ChartTitleLimitKind::WireNesting => ChartLegendVisibilityLimitKind::WireNesting,
        ChartTitleLimitKind::WireWork => ChartLegendVisibilityLimitKind::WireWork,
        _ => ChartLegendVisibilityLimitKind::LegendBytes,
    }
}

fn map_legend_codec_error(error: ChartLegendVisibilityDecodeError) -> ChartLegendVisibilityError {
    if let Some(limit) = error.resource_limit() {
        return match limit {
            ChartLegendVisibilityDecodeLimit::Bytes { observed, maximum } => {
                ChartLegendVisibilityError::LimitExceeded {
                    kind: ChartLegendVisibilityLimitKind::WireBytes,
                    observed: usize_to_u64(observed),
                    maximum: usize_to_u64(maximum),
                }
            },
            ChartLegendVisibilityDecodeLimit::Fields { observed, maximum } => {
                ChartLegendVisibilityError::LimitExceeded {
                    kind: ChartLegendVisibilityLimitKind::WireFields,
                    observed: usize_to_u64(observed),
                    maximum: usize_to_u64(maximum),
                }
            },
            ChartLegendVisibilityDecodeLimit::Work { observed, maximum } => {
                ChartLegendVisibilityError::LimitExceeded {
                    kind: ChartLegendVisibilityLimitKind::WireWork,
                    observed: usize_to_u64(observed),
                    maximum: usize_to_u64(maximum),
                }
            },
            ChartLegendVisibilityDecodeLimit::Output { observed, maximum } => {
                ChartLegendVisibilityError::LimitExceeded {
                    kind: ChartLegendVisibilityLimitKind::OutputBytes,
                    observed: usize_to_u64(observed),
                    maximum: usize_to_u64(maximum),
                }
            },
            ChartLegendVisibilityDecodeLimit::Nesting { observed, maximum } => {
                ChartLegendVisibilityError::LimitExceeded {
                    kind: ChartLegendVisibilityLimitKind::WireNesting,
                    observed: u64::from(observed),
                    maximum: u64::from(maximum),
                }
            },
            ChartLegendVisibilityDecodeLimit::Allocations { observed, maximum } => {
                ChartLegendVisibilityError::LimitExceeded {
                    kind: ChartLegendVisibilityLimitKind::Allocations,
                    observed: usize_to_u64(observed),
                    maximum: usize_to_u64(maximum),
                }
            },
            ChartLegendVisibilityDecodeLimit::Retained { observed, maximum }
            | ChartLegendVisibilityDecodeLimit::Scratch { observed, maximum } => {
                ChartLegendVisibilityError::LimitExceeded {
                    kind: ChartLegendVisibilityLimitKind::TotalBytes,
                    observed: usize_to_u64(observed),
                    maximum: usize_to_u64(maximum),
                }
            },
            _ => ChartLegendVisibilityError::InvalidSource,
        };
    }
    if let Some(limit) = error.wire_resource_limit() {
        return match limit {
            ChartLegendVisibilityWireResourceLimit::Bytes { observed, maximum } => {
                ChartLegendVisibilityError::LimitExceeded {
                    kind: ChartLegendVisibilityLimitKind::WireBytes,
                    observed: usize_to_u64(observed),
                    maximum: usize_to_u64(maximum),
                }
            },
            ChartLegendVisibilityWireResourceLimit::Nesting { observed, maximum } => {
                ChartLegendVisibilityError::LimitExceeded {
                    kind: ChartLegendVisibilityLimitKind::WireNesting,
                    observed: u64::from(observed),
                    maximum: u64::from(maximum),
                }
            },
            _ => ChartLegendVisibilityError::InvalidSource,
        };
    }
    if let Some(amount) = error.allocation_amount() {
        return ChartLegendVisibilityError::Allocation { amount };
    }
    ChartLegendVisibilityError::InvalidSource
}

fn map_rendering_error(
    _error: super::rendering_invalidation::RenderingInvalidationError,
) -> ChartLegendVisibilityError {
    ChartLegendVisibilityError::InvalidSource
}

fn map_read_error(error: ReadError) -> ChartLegendVisibilityError {
    match error {
        ReadError::SemanticLimit {
            kind,
            observed,
            maximum,
            ..
        } => ChartLegendVisibilityError::LimitExceeded {
            kind: match kind {
                SemanticLimitKind::Objects => ChartLegendVisibilityLimitKind::Entries,
                SemanticLimitKind::Slides => ChartLegendVisibilityLimitKind::Slides,
                SemanticLimitKind::References => ChartLegendVisibilityLimitKind::References,
                SemanticLimitKind::TextStorages
                | SemanticLimitKind::TextFragments
                | SemanticLimitKind::TextBytes => ChartLegendVisibilityLimitKind::WireFields,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        ReadError::PayloadLimit {
            kind,
            observed,
            maximum,
            ..
        } => ChartLegendVisibilityError::LimitExceeded {
            kind: match kind {
                super::PayloadLimitKind::Bytes => ChartLegendVisibilityLimitKind::WireBytes,
                super::PayloadLimitKind::Fields => ChartLegendVisibilityLimitKind::WireFields,
                super::PayloadLimitKind::Nesting => ChartLegendVisibilityLimitKind::WireNesting,
                super::PayloadLimitKind::Work => ChartLegendVisibilityLimitKind::WireWork,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        ReadError::Allocation { amount, .. } => ChartLegendVisibilityError::Allocation { amount },
        _ => ChartLegendVisibilityError::InvalidSource,
    }
}

fn map_archive_error(error: litchi_iwa_archive::Error) -> ChartLegendVisibilityError {
    match error {
        litchi_iwa_archive::Error::Limit {
            kind,
            observed,
            maximum,
        } => ChartLegendVisibilityError::LimitExceeded {
            kind: match kind {
                litchi_iwa_archive::LimitKind::InputBytes => {
                    ChartLegendVisibilityLimitKind::InputBytes
                },
                litchi_iwa_archive::LimitKind::OutputBytes => {
                    ChartLegendVisibilityLimitKind::OutputBytes
                },
                litchi_iwa_archive::LimitKind::Entries => ChartLegendVisibilityLimitKind::Entries,
                litchi_iwa_archive::LimitKind::EntryBytes => {
                    ChartLegendVisibilityLimitKind::EntryBytes
                },
                litchi_iwa_archive::LimitKind::TotalBytes => {
                    ChartLegendVisibilityLimitKind::TotalBytes
                },
                _ => ChartLegendVisibilityLimitKind::WireBytes,
            },
            observed,
            maximum,
        },
        litchi_iwa_archive::Error::Allocation { amount, .. } => {
            ChartLegendVisibilityError::Allocation { amount }
        },
        _ => ChartLegendVisibilityError::InvalidSource,
    }
}

fn map_core_error(error: litchi_iwa_core::Error) -> ChartLegendVisibilityError {
    match error {
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => ChartLegendVisibilityError::LimitExceeded {
            kind: match kind {
                litchi_iwa_core::LimitKind::ArchiveBytes => {
                    ChartLegendVisibilityLimitKind::TotalBytes
                },
                litchi_iwa_core::LimitKind::Objects
                | litchi_iwa_core::LimitKind::Messages
                | litchi_iwa_core::LimitKind::MessagesPerObject
                | litchi_iwa_core::LimitKind::MetadataItems => {
                    ChartLegendVisibilityLimitKind::Entries
                },
                litchi_iwa_core::LimitKind::ObjectBytes
                | litchi_iwa_core::LimitKind::MessageBytes
                | litchi_iwa_core::LimitKind::HeaderBytes
                | litchi_iwa_core::LimitKind::HeaderMemoryBytes
                | litchi_iwa_core::LimitKind::SnappyChunkBytes
                | litchi_iwa_core::LimitKind::SnappyStreamBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedChunkBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedStreamBytes => {
                    ChartLegendVisibilityLimitKind::WireBytes
                },
                litchi_iwa_core::LimitKind::HeaderFields => {
                    ChartLegendVisibilityLimitKind::WireFields
                },
                litchi_iwa_core::LimitKind::HeaderNesting => {
                    ChartLegendVisibilityLimitKind::WireNesting
                },
                litchi_iwa_core::LimitKind::SnappyFrames => {
                    ChartLegendVisibilityLimitKind::WireWork
                },
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            ChartLegendVisibilityError::Allocation { amount: requested }
        },
        _ => ChartLegendVisibilityError::InvalidSource,
    }
}

fn map_wire_error(error: litchi_iwa_common::Error) -> ChartLegendVisibilityError {
    match error {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => ChartLegendVisibilityError::LimitExceeded {
            kind: match kind {
                litchi_iwa_common::LimitKind::InputBytes => {
                    ChartLegendVisibilityLimitKind::WireBytes
                },
                litchi_iwa_common::LimitKind::OutputBytes => {
                    ChartLegendVisibilityLimitKind::OutputBytes
                },
                litchi_iwa_common::LimitKind::Fields
                | litchi_iwa_common::LimitKind::TableRows
                | litchi_iwa_common::LimitKind::TableColumns
                | litchi_iwa_common::LimitKind::TableCells
                | litchi_iwa_common::LimitKind::MaterializedCells => {
                    ChartLegendVisibilityLimitKind::WireFields
                },
                litchi_iwa_common::LimitKind::Nesting => {
                    ChartLegendVisibilityLimitKind::WireNesting
                },
                litchi_iwa_common::LimitKind::RewriteWork => {
                    ChartLegendVisibilityLimitKind::WireWork
                },
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(limit),
        },
        litchi_iwa_common::Error::Allocation { amount, .. } => {
            ChartLegendVisibilityError::Allocation { amount }
        },
        _ => ChartLegendVisibilityError::InvalidSource,
    }
}

fn usize_to_u64(value: usize) -> u64 {
    u64::try_from(value).unwrap_or(u64::MAX)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn codec_allocation_limit_maps_to_public_allocation_kind() {
        let options = DecodeOptions::for_source(&[])
            .with_max_output_bytes(64)
            .with_max_retained_bytes(64)
            .with_max_scratch_bytes(64)
            .with_max_allocations(0);
        let error = prepare_chart_legend_visibility_rewrite(
            &[],
            ChartLegendVisibilityWrite::new(true),
            options,
        )
        .expect_err("a non-empty legend rewrite requires one allocation");
        assert_eq!(
            map_legend_codec_error(error),
            ChartLegendVisibilityError::LimitExceeded {
                kind: ChartLegendVisibilityLimitKind::Allocations,
                observed: 1,
                maximum: 0,
            }
        );
    }
}
