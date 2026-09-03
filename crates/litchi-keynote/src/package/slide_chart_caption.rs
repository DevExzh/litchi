//! Exact-source, selector-first Keynote chart-caption text transactions.
//!
//! This owner admits the canonical inline chart-caption graph transition:
//! creation from an exclusive stand-in and removal by retaining the old graph
//! while retargeting the chart to a fresh stand-in.  Existing storage text is
//! still edited through the focused storage owner.  Cross-component graphs,
//! shared ownership, and metadata shapes that cannot be proven exactly remain
//! outside this transaction.

#![allow(
    clippy::map_err_ignore,
    clippy::needless_pass_by_value,
    clippy::shadow_reuse,
    clippy::wildcard_enum_match_arm,
    reason = "The transaction redacts native graph failures at the semantic boundary."
)]

use std::collections::HashSet;
use std::fmt;
use std::sync::Arc;

use litchi_core::Position;
use litchi_iwa_archive::{
    SourceCatalog,
    package::{EntryEdit, ExactArtifacts, ReassemblyExecutionLimits},
};
use litchi_iwa_common::{
    WireLimits, decode_varint_from_bytes,
    varint::encoded_len,
    wire::{WireView, append_length_delimited_field_with_limits, append_varint_field},
};
use litchi_iwa_core::archive::{FieldObjectReferenceTransition, ObjectReferenceTransition};
use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
use litchi_iwa_protos::package_metadata_codec::{
    self as metadata_codec, AdditionSaveTokenBatch, Batch as MetadataBatch, ComponentDescriptor,
    ComponentSelector, ObjectUuidAddition, PackageMetadataVisitor,
    RewriteOptions as MetadataRewriteOptions, SaveTokenBatch, UuidBits,
    inspect_package_metadata_with_visitor, prepare_package_metadata_additions_and_save_tokens,
    prepare_package_metadata_save_tokens,
};
use litchi_iwa_protos::{
    keynote_chart_caption_codec, keynote_chart_caption_graph_codec as graph_codec,
    keynote_movie_caption_codec, pages_movie_caption_codec,
};
use thiserror::Error;

use super::{
    Package, PhysicalSource, ReadError, STORAGE_MESSAGE_TYPE, SemanticLimitKind,
    slide_chart_title::{ChartSelection, select_chart},
    unique_payload,
};
use crate::{ChartSelector, SlideSelector};

const CHART_MESSAGE_TYPE: u32 = 5_021;
const DOCUMENT_MESSAGE_TYPE: u32 = 1;
const SHOW_MESSAGE_TYPE: u32 = 2;
const THEME_MESSAGE_TYPE: u32 = 10;
const STANDIN_MESSAGE_TYPE: u32 = 3_097;
const CAPTION_INFO_MESSAGE_TYPE: u32 = 633;
const CAPTION_PLACEMENT_MESSAGE_TYPE: u32 = 634;
const SHAPE_STYLE_MESSAGE_TYPE: u32 = 2_025;
const PACKAGE_METADATA_MESSAGE_TYPE: u32 = 11_006;
const PACKAGE_METADATA_MEMBER_NAME: &str = "Index/Metadata.iwa";
const PREVIEW_ENTRY_NAMES: [&str; 3] = ["preview.jpg", "preview-micro.jpg", "preview-web.jpg"];
const MAX_CAPTION_BYTES: usize = 64 * 1024 * 1024;

/// Native drawable edge whose caption reference is being transitioned.
///
/// Charts and movies share the graph, Metadata, and preview transaction. Only
/// the drawable message type and strict edge codec differ.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum CaptionEdgeKind {
    Chart,
    Movie,
    MovieTitle,
}

/// Operation-local accounting for the chart-caption owner.
///
/// The lower-level codecs and archive reassembler each have their own
/// source/output budgets.  This small ledger accounts the work which happens
/// between those seams: the package-wide ownership census, graph encoding,
/// metadata preparation/execution, archive serialization/compression, and the
/// exact ZIP execution plan.  It deliberately does not live in a patch, so a
/// replay cannot inherit stale budget state from the original transaction.
#[derive(Debug, Clone, Copy)]
pub(super) struct CaptionBudget {
    maximum_input: usize,
    maximum_output: usize,
    maximum_fields: usize,
    maximum_work: usize,
    maximum_depth: u32,
    maximum_components: usize,
    maximum_references: usize,
    maximum_allocations: usize,
    input: usize,
    output: usize,
    fields: usize,
    max_depth: u32,
    components: usize,
    references: usize,
    work: usize,
    allocations: usize,
    retained_bytes: usize,
    scratch_bytes: usize,
    candidate_reopens: usize,
}

impl CaptionBudget {
    pub(super) fn for_package(package: &Package) -> Result<Self, ChartCaptionError> {
        let wire = package.wire_limits().map_err(map_wire_error)?;
        let semantic = package.semantic_limits();
        let physical = package.state.options.archive();
        let physical_input = usize::try_from(physical.max_input_bytes()).unwrap_or(usize::MAX);
        let aggregate_input = physical_input.checked_mul(4).unwrap_or(usize::MAX);
        Ok(Self {
            maximum_input: aggregate_input,
            maximum_output: aggregate_input,
            maximum_fields: wire.max_fields(),
            maximum_work: wire.max_rewrite_work(),
            maximum_depth: u32::try_from(wire.max_nesting()).unwrap_or(u32::MAX),
            maximum_components: semantic.max_objects(),
            maximum_references: semantic.max_references(),
            maximum_allocations: aggregate_input.saturating_add(64),
            input: 0,
            output: 0,
            fields: 0,
            max_depth: 0,
            components: 0,
            references: 0,
            work: 0,
            allocations: 0,
            retained_bytes: 0,
            scratch_bytes: 0,
            candidate_reopens: 0,
        })
    }

    fn limit(kind: ChartCaptionLimitKind, observed: usize, maximum: usize) -> ChartCaptionError {
        ChartCaptionError::LimitExceeded {
            kind,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        }
    }

    fn charge_counter(
        current: &mut usize,
        amount: usize,
        kind: ChartCaptionLimitKind,
        maximum: usize,
    ) -> Result<(), ChartCaptionError> {
        let observed = current
            .checked_add(amount)
            .ok_or(ChartCaptionError::InvalidSource)?;
        if observed > maximum {
            return Err(Self::limit(kind, observed, maximum));
        }
        *current = observed;
        Ok(())
    }

    fn charge_input(&mut self, amount: usize) -> Result<(), ChartCaptionError> {
        Self::charge_counter(
            &mut self.input,
            amount,
            ChartCaptionLimitKind::InputBytes,
            self.maximum_input,
        )
    }

    fn charge_output(&mut self, amount: usize) -> Result<(), ChartCaptionError> {
        Self::charge_counter(
            &mut self.output,
            amount,
            ChartCaptionLimitKind::OutputBytes,
            self.maximum_output,
        )
    }

    fn charge_fields(&mut self, amount: usize) -> Result<(), ChartCaptionError> {
        Self::charge_counter(
            &mut self.fields,
            amount,
            ChartCaptionLimitKind::WireFields,
            self.maximum_fields,
        )
    }

    fn charge_components(&mut self, amount: usize) -> Result<(), ChartCaptionError> {
        Self::charge_counter(
            &mut self.components,
            amount,
            ChartCaptionLimitKind::Entries,
            self.maximum_components,
        )
    }

    fn charge_references(&mut self, amount: usize) -> Result<(), ChartCaptionError> {
        Self::charge_counter(
            &mut self.references,
            amount,
            ChartCaptionLimitKind::References,
            self.maximum_references,
        )
    }

    fn charge_work(&mut self, amount: usize) -> Result<(), ChartCaptionError> {
        Self::charge_counter(
            &mut self.work,
            amount,
            ChartCaptionLimitKind::WireWork,
            self.maximum_work,
        )
    }

    fn charge_depth(&mut self, depth: u32) -> Result<(), ChartCaptionError> {
        self.max_depth = self.max_depth.max(depth);
        if self.max_depth > self.maximum_depth {
            return Err(Self::limit(
                ChartCaptionLimitKind::WireNesting,
                self.max_depth as usize,
                self.maximum_depth as usize,
            ));
        }
        Ok(())
    }

    fn charge_allocation(
        &mut self,
        count: usize,
        retained: usize,
    ) -> Result<(), ChartCaptionError> {
        let observed = self
            .allocations
            .checked_add(count)
            .ok_or(ChartCaptionError::InvalidSource)?;
        if observed > self.maximum_allocations {
            return Err(Self::limit(
                ChartCaptionLimitKind::Entries,
                observed,
                self.maximum_allocations,
            ));
        }
        self.allocations = observed;
        self.retained_bytes = self
            .retained_bytes
            .checked_add(retained)
            .ok_or(ChartCaptionError::InvalidSource)?;
        Ok(())
    }

    fn charge_scratch(&mut self, amount: usize) -> Result<(), ChartCaptionError> {
        self.scratch_bytes = self
            .scratch_bytes
            .checked_add(amount)
            .ok_or(ChartCaptionError::InvalidSource)?;
        Ok(())
    }

    fn charge_report(
        &mut self,
        report: metadata_codec::RewriteReport,
    ) -> Result<(), ChartCaptionError> {
        self.charge_input(report.input_bytes())?;
        self.charge_output(report.output_bytes())?;
        self.charge_fields(report.fields())?;
        self.charge_work(report.work_bytes())?;
        self.charge_depth(report.max_depth())?;
        self.charge_components(report.components_scanned())?;
        self.charge_references(report.references_scanned())?;
        self.charge_allocation(report.allocations(), report.retained_bytes())?;
        self.charge_scratch(report.scratch_bytes())
    }

    fn charge_execution_requirements(
        &mut self,
        requirements: metadata_codec::RewriteExecutionRequirements,
    ) -> Result<(), ChartCaptionError> {
        self.charge_output(requirements.output_bytes())?;
        self.charge_fields(requirements.fields())?;
        self.charge_work(requirements.work_bytes())?;
        self.charge_components(requirements.components())?;
        self.charge_references(requirements.references())?;
        self.charge_allocation(requirements.allocations(), requirements.retained_bytes())?;
        self.charge_scratch(requirements.scratch_bytes())
    }

    fn observe_metadata_execution(
        &mut self,
        report: metadata_codec::RewriteReport,
        requirements: metadata_codec::RewriteExecutionRequirements,
    ) -> Result<(), ChartCaptionError> {
        if report.output_bytes() > requirements.output_bytes()
            || report.fields() > requirements.fields()
            || report.work_bytes() > requirements.work_bytes()
            || report.components_scanned() > requirements.components()
            || report.references_scanned() > requirements.references()
            || report.allocations() > requirements.allocations()
            || report.retained_bytes() > requirements.retained_bytes()
            || report.scratch_bytes() > requirements.scratch_bytes()
        {
            return Err(ChartCaptionError::Verification);
        }
        self.charge_depth(report.max_depth())
    }

    fn observe_graph_report(
        &mut self,
        report: graph_codec::EncodeReport,
    ) -> Result<(), ChartCaptionError> {
        // This report spans the graph codec's sizing, emission, and four
        // output-buffer allocations. Charge every dimension to the enclosing
        // package transaction rather than treating the codec as a separate
        // unlimited phase.
        self.charge_output(report.output_bytes())?;
        self.charge_fields(report.fields())?;
        self.charge_work(report.work_bytes())?;
        self.charge_depth(report.max_depth())?;
        self.charge_allocation(report.allocations(), report.output_bytes())?;
        self.charge_scratch(report.output_bytes())
    }

    fn precharge_graph(
        &mut self,
        text: &str,
        language: Option<&str>,
    ) -> Result<(), ChartCaptionError> {
        let text_bytes = text
            .len()
            .checked_add(language.map_or(0, str::len))
            .ok_or(ChartCaptionError::InvalidSource)?;
        if text_bytes > MAX_CAPTION_BYTES {
            return Err(Self::limit(
                ChartCaptionLimitKind::CaptionBytes,
                text_bytes,
                MAX_CAPTION_BYTES,
            ));
        }
        let payload_bound = text_bytes
            .checked_mul(16)
            .and_then(|value| value.checked_add(16 * 1024))
            .ok_or(ChartCaptionError::InvalidSource)?;
        self.charge_work(payload_bound)?;
        self.charge_fields(64)?;
        self.charge_depth(16)?;
        self.charge_allocation(4, payload_bound)?;
        Ok(())
    }

    pub(super) fn charge_catalog_scan(
        &mut self,
        package: &Package,
    ) -> Result<(), ChartCaptionError> {
        let mut work = 0usize;
        for component in package.state.source.components().iter() {
            work = work
                .checked_add(component.name().len())
                .ok_or(ChartCaptionError::InvalidSource)?;
            let archive = component.archive();
            work = work
                .checked_add(archive.encoded_len().map_err(map_core_error)?)
                .ok_or(ChartCaptionError::InvalidSource)?;
            for object in &archive.objects {
                work = work
                    .checked_add(1)
                    .and_then(|value| value.checked_add(object.messages.len()))
                    .ok_or(ChartCaptionError::InvalidSource)?;
                for message in &object.messages {
                    work = work
                        .checked_add(message.data.len())
                        .ok_or(ChartCaptionError::InvalidSource)?;
                }
            }
        }
        self.charge_input(package.state.source.shared_source().len())?;
        self.charge_components(package.state.source.components().len())?;
        self.charge_work(work)
    }

    /// Charge a package-wide validation/ownership traversal.
    ///
    /// Selection and graph verification repeatedly walk the already parsed
    /// physical catalog.  The initial catalog charge above accounts for the
    /// first pass, but it must not make later ownership passes appear free.
    /// This helper charges archive/message traversal work and every
    /// aggregate/field reference inspected by `passes` additional traversals.
    /// Source bytes are charged at the physical catalog/reopen seams. It
    /// intentionally does not pretend to know the exact
    /// number of temporary allocations made by lower-level semantic readers;
    /// callers charge those at the operation boundary where they are known.
    pub(super) fn charge_validation_scan(
        &mut self,
        package: &Package,
        passes: usize,
    ) -> Result<(), ChartCaptionError> {
        if passes == 0 {
            return Ok(());
        }
        let mut work = 0usize;
        for component in package.state.source.components().iter() {
            let archive_bytes = component.archive().encoded_len().map_err(map_core_error)?;
            work = work
                .checked_add(component.name().len())
                .and_then(|value| value.checked_add(archive_bytes))
                .ok_or(ChartCaptionError::InvalidSource)?;
            for object in &component.archive().objects {
                work = work
                    .checked_add(1)
                    .and_then(|value| value.checked_add(object.messages.len()))
                    .and_then(|value| value.checked_add(object.archive_info.message_infos.len()))
                    .ok_or(ChartCaptionError::InvalidSource)?;
                for (message_index, message) in object.messages.iter().enumerate() {
                    work = work
                        .checked_add(message.data.len())
                        .ok_or(ChartCaptionError::InvalidSource)?;
                    let Some(info) = object.archive_info.message_infos.get(message_index) else {
                        continue;
                    };
                    let mut local_references = info
                        .object_references
                        .len()
                        .checked_add(info.data_references.len())
                        .and_then(|value| value.checked_add(info.field_infos.len()))
                        .ok_or(ChartCaptionError::InvalidSource)?;
                    for field in &info.field_infos {
                        local_references =
                            local_references
                                .checked_add(field.path.path.len())
                                .and_then(|value| {
                                    value.checked_add(field.object_references.len()).and_then(
                                        |value| value.checked_add(field.data_references.len()),
                                    )
                                })
                                .ok_or(ChartCaptionError::InvalidSource)?;
                    }
                    work = work
                        .checked_add(local_references)
                        .ok_or(ChartCaptionError::InvalidSource)?;
                }
            }
        }
        // The physical source bytes are charged by `charge_catalog_scan`,
        // while a reopened candidate is charged by `charge_candidate_reopen`.
        // Repeated validation traversals operate on those already materialized
        // bytes, so charging the complete ZIP again would reject exact public
        // input profiles without measuring a new ingress allocation. Their
        // reference comparisons are included in work below rather than being
        // charged against the one-package semantic reference ceiling again.
        self.charge_work(
            work.checked_mul(passes)
                .ok_or(ChartCaptionError::InvalidSource)?,
        )
    }

    /// Charge the bounded allocations made while resolving a movie/chart
    /// selector.  The exact vector capacities are format-dependent, so this
    /// records only the known selector scratch allocations; payload and
    /// candidate buffers are charged by their owning codec/reassembler.
    pub(super) fn charge_selection_scan(
        &mut self,
        package: &Package,
        passes: usize,
    ) -> Result<(), ChartCaptionError> {
        self.charge_validation_scan(package, passes)?;
        self.charge_allocation(
            passes
                .checked_mul(2)
                .ok_or(ChartCaptionError::InvalidSource)?,
            0,
        )
    }

    fn charge_metadata_report(
        &mut self,
        report: metadata_codec::RewriteReport,
    ) -> Result<(), ChartCaptionError> {
        self.charge_report(report)
    }

    fn charge_archive_snappy_plan(
        &mut self,
        archive: &Archive,
        archive_limits: litchi_iwa_core::Limits,
        snappy_limits: litchi_iwa_core::SnappyLimits,
    ) -> Result<usize, ChartCaptionError> {
        let encoded = archive
            .encoded_len_with_limits(archive_limits)
            .map_err(map_core_error)?;
        let compressed_bound =
            SnappyStream::maximum_compressed_len(encoded).map_err(map_core_error)?;
        if compressed_bound > snappy_limits.max_compressed_stream() {
            return Err(ChartCaptionError::LimitExceeded {
                kind: ChartCaptionLimitKind::EntryBytes,
                observed: usize_to_u64(compressed_bound),
                maximum: usize_to_u64(snappy_limits.max_compressed_stream()),
            });
        }
        self.charge_work(
            encoded
                .checked_add(compressed_bound)
                .ok_or(ChartCaptionError::InvalidSource)?,
        )?;
        self.charge_output(encoded)?;
        self.charge_allocation(
            2,
            encoded
                .checked_add(compressed_bound)
                .ok_or(ChartCaptionError::InvalidSource)?,
        )?;
        Ok(compressed_bound)
    }

    fn charge_reassembly(
        &mut self,
        requirements: litchi_iwa_archive::package::ReassemblyExecutionRequirements,
    ) -> Result<ReassemblyExecutionLimits, ChartCaptionError> {
        let amount = requirements
            .output_bytes()
            .checked_add(requirements.scratch_bytes())
            .and_then(|value| value.checked_add(requirements.offset_count()))
            .ok_or(ChartCaptionError::InvalidSource)?;
        self.charge_work(amount)?;
        self.charge_output(requirements.output_bytes())?;
        self.charge_allocation(requirements.allocations(), requirements.retained_bytes())?;
        self.charge_scratch(requirements.scratch_bytes())?;
        Ok(requirements.exact_limits())
    }

    fn precharge_intermediate_candidate(
        &mut self,
        package: &Package,
        text_bytes: usize,
    ) -> Result<(), ChartCaptionError> {
        let source_bytes = package.state.source.shared_source().len();
        let estimate = source_bytes
            .checked_mul(2)
            .and_then(|value| value.checked_add(text_bytes))
            .ok_or(ChartCaptionError::InvalidSource)?;
        self.charge_work(estimate)?;
        self.charge_input(source_bytes)?;
        self.charge_output(estimate)?;
        self.charge_allocation(2, estimate)?;
        Ok(())
    }

    pub(super) fn charge_candidate_reopen(
        &mut self,
        output_bytes: usize,
    ) -> Result<(), ChartCaptionError> {
        self.charge_input(output_bytes)?;
        self.charge_work(output_bytes)?;
        self.charge_allocation(1, output_bytes)?;
        self.candidate_reopens = self
            .candidate_reopens
            .checked_add(1)
            .ok_or(ChartCaptionError::InvalidSource)?;
        Ok(())
    }

    pub(super) fn charge_exact_artifacts(
        &mut self,
        source_bytes: usize,
        target_bytes: usize,
    ) -> Result<(), ChartCaptionError> {
        self.charge_work(
            source_bytes
                .checked_add(target_bytes)
                .ok_or(ChartCaptionError::InvalidSource)?,
        )
    }

    fn remaining_input(&self) -> usize {
        self.maximum_input.checked_sub(self.input).unwrap_or(0)
    }

    fn remaining_output(&self) -> usize {
        self.maximum_output.checked_sub(self.output).unwrap_or(0)
    }

    fn remaining_fields(&self) -> usize {
        self.maximum_fields.checked_sub(self.fields).unwrap_or(0)
    }

    fn remaining_work(&self) -> usize {
        self.maximum_work.checked_sub(self.work).unwrap_or(0)
    }

    fn remaining_components(&self) -> usize {
        self.maximum_components
            .checked_sub(self.components)
            .unwrap_or(0)
    }

    fn remaining_references(&self) -> usize {
        self.maximum_references
            .checked_sub(self.references)
            .unwrap_or(0)
    }

    fn remaining_depth(&self) -> u32 {
        self.maximum_depth
    }

    fn remaining_allocations(&self) -> usize {
        self.maximum_allocations
            .checked_sub(self.allocations)
            .unwrap_or(0)
    }
}

/// A finite resource governed while a chart-caption transaction is prepared
/// or published.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum ChartCaptionLimitKind {
    /// Complete input package bytes.
    InputBytes,
    /// Complete rewritten package or payload bytes.
    OutputBytes,
    /// Bytes in one protobuf payload.
    WireBytes,
    /// ZIP members, IWA objects, or IWA messages.
    Entries,
    /// Bytes in one package member, IWA object, or message.
    EntryBytes,
    /// Aggregate package or IWA bytes.
    TotalBytes,
    /// Semantic slides.
    Slides,
    /// Semantic graph references.
    References,
    /// Semantic text-storage objects.
    TextStorages,
    /// Semantic text fragments.
    TextFragments,
    /// Aggregate semantic text bytes.
    TextBytes,
    /// Parsed protobuf fields.
    WireFields,
    /// Protobuf nesting depth.
    WireNesting,
    /// Aggregate protobuf scan and rewrite work.
    WireWork,
    /// UTF-8 bytes in one caption value.
    CaptionBytes,
}

impl fmt::Display for ChartCaptionLimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::OutputBytes => "output bytes",
            Self::WireBytes => "wire bytes",
            Self::Entries => "entries",
            Self::EntryBytes => "entry bytes",
            Self::TotalBytes => "total bytes",
            Self::Slides => "slides",
            Self::References => "references",
            Self::TextStorages => "text storages",
            Self::TextFragments => "text fragments",
            Self::TextBytes => "text bytes",
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting depth",
            Self::WireWork => "wire work",
            Self::CaptionBytes => "caption bytes",
        })
    }
}

/// A content-redacted failure raised by a chart-caption transaction.
#[derive(Debug, Clone, PartialEq, Eq, Error)]
#[non_exhaustive]
pub enum ChartCaptionError {
    /// The source was prepared without an exact physical package artifact.
    #[error("this Keynote source does not support physical chart-caption edits")]
    UnsupportedSource,
    /// The requested operation needs a graph shape outside the supported
    /// canonical inline stand-in transition.
    #[error("the requested Keynote chart-caption graph operation is unsupported")]
    UnsupportedDependency,
    /// An exact-name slide or chart selector was ambiguous.
    #[error("the Keynote chart-caption selector is ambiguous")]
    AmbiguousSelector,
    /// An exact-name slide selector was empty.
    #[error("the Keynote slide selector name cannot be empty")]
    EmptySlideName,
    /// An exact-name slide selector did not match.
    #[error("the Keynote show has no slide matching the requested name")]
    SlideNameNotFound,
    /// A checked semantic slide position does not exist.
    #[error("the Keynote show has no slide at position {position:?}")]
    SlidePositionNotFound {
        /// Missing checked semantic slide position.
        position: Position,
    },
    /// An exact-name chart selector did not match.
    #[error("the selected Keynote slide has no chart matching the requested name")]
    ChartNameNotFound,
    /// A checked semantic chart position does not exist.
    #[error("the selected Keynote slide has no chart at position {position:?}")]
    ChartPositionNotFound {
        /// Missing checked semantic chart position.
        position: Position,
    },
    /// An empty exact chart name was supplied.
    #[error("the Keynote chart selector name cannot be empty")]
    EmptyChartName,
    /// The source chart, caption graph, or text storage is malformed or unsafe.
    #[error("the Keynote chart-caption source cannot be edited safely")]
    InvalidSource,
    /// A finite resource ceiling was exceeded.
    #[error("Keynote chart-caption {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    LimitExceeded {
        /// Resource category that exceeded its limit.
        kind: ChartCaptionLimitKind,
        /// Observed or requested amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded destination allocation failed before publication.
    #[error("could not allocate {amount} units for the Keynote chart-caption transaction")]
    Allocation {
        /// Elements or bytes requested.
        amount: usize,
    },
    /// Full candidate reopening did not reproduce the requested caption.
    #[error("the edited Keynote chart caption failed semantic verification")]
    Verification,
    /// The patch does not belong to this exact immutable package artifact.
    #[error("the Keynote chart-caption patch does not match the exact source package")]
    PatchConflict,
}

/// One mutable chart-caption value staged against an immutable package.
pub struct ChartCaptionEdit<'a> {
    source: &'a Package,
    selection: CaptionSelection,
    after: Option<String>,
}

impl fmt::Debug for ChartCaptionEdit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("ChartCaptionEdit")
            .field("slide_position", &self.selection.slide_position)
            .field("chart_position", &self.selection.chart_position)
            .field("has_before", &self.selection.text.is_some())
            .field("has_after", &self.after.is_some())
            .finish_non_exhaustive()
    }
}

impl<'a> ChartCaptionEdit<'a> {
    fn new<'slide, 'chart>(
        source: &'a Package,
        slide_selector: impl Into<SlideSelector<'slide>>,
        chart_selector: impl Into<ChartSelector<'chart>>,
    ) -> Result<Self, ChartCaptionError> {
        let selection =
            select_caption(source, slide_selector.into(), chart_selector.into(), false)?;
        let after = selection.text.as_deref().map(copy_caption).transpose()?;
        Ok(Self {
            source,
            selection,
            after,
        })
    }

    /// Return the selected semantic slide position.
    #[must_use]
    pub const fn slide_position(&self) -> Position {
        self.selection.slide_position
    }

    /// Return the selected semantic chart position within its slide.
    #[must_use]
    pub const fn chart_position(&self) -> Position {
        self.selection.chart_position
    }

    /// Borrow the caption text observed when this edit began.
    #[must_use]
    pub fn before(&self) -> Option<&str> {
        self.selection.text.as_deref()
    }

    /// Borrow the caption text staged for publication.
    #[must_use]
    pub fn after(&self) -> Option<&str> {
        self.after.as_deref()
    }

    /// Stage whole-text replacement or creation of a native caption.
    pub fn set(mut self, caption: impl AsRef<str>) -> Result<Self, ChartCaptionError> {
        self.after = Some(copy_caption(caption.as_ref())?);
        Ok(self)
    }

    /// Stage caption removal.
    ///
    /// Stage removal of an active caption.  Commit replaces the edge with a
    /// fresh canonical stand-in and retains the old graph for exact history
    /// and inverse semantics. Clearing an already absent caption is an exact
    /// no-op.
    pub fn clear(mut self) -> Result<Self, ChartCaptionError> {
        self.after = None;
        Ok(self)
    }

    /// Validate and atomically publish the staged immutable candidate.
    pub fn commit(self) -> Result<ChartCaptionCommit, ChartCaptionError> {
        let catalog = physical_catalog(self.source)?;
        let source_bytes = catalog.shared_source();
        let current = select_caption(
            self.source,
            SlideSelector::position(self.selection.slide_position),
            ChartSelector::index(self.selection.chart_position.get()),
            false,
        )?;
        if !current.same_identity(&self.selection) || current.text != self.selection.text {
            return Err(ChartCaptionError::InvalidSource);
        }
        if self.selection.text == self.after {
            self.source.validate().map_err(map_read_error)?;
            return Ok(ChartCaptionCommit {
                package: self.source.snapshot(),
                patch: ChartCaptionPatch {
                    artifacts: ExactArtifacts::new(Arc::clone(&source_bytes), source_bytes),
                    target_selection: self.selection.clone(),
                    selection: self.selection,
                    after: self.after,
                    touched_components: 0,
                    deleted_previews: 0,
                    target_requires_invalidated_previews: false,
                },
                diagnostics: ChartCaptionDiagnostics::unchanged(),
            });
        }
        if !catalog.source_is_exact() {
            return Err(ChartCaptionError::UnsupportedSource);
        }
        self.source.validate().map_err(map_read_error)?;
        let guarded = select_caption(
            self.source,
            SlideSelector::position(self.selection.slide_position),
            ChartSelector::index(self.selection.chart_position.get()),
            true,
        )?;
        if !guarded.same_identity(&self.selection) || guarded.text != self.selection.text {
            return Err(ChartCaptionError::InvalidSource);
        }
        let mut budget = CaptionBudget::for_package(self.source)?;
        budget.charge_catalog_scan(self.source)?;
        budget.charge_selection_scan(self.source, 1)?;
        let (package, touched_components, deleted_previews) = rewrite_chart_caption_operation(
            self.source,
            &self.selection,
            self.after.as_deref(),
            &mut budget,
        )?;
        budget.charge_selection_scan(&package, 1)?;
        let candidate = select_caption(
            &package,
            SlideSelector::position(self.selection.slide_position),
            ChartSelector::index(self.selection.chart_position.get()),
            true,
        )?;
        if !candidate.same_chart_identity(&self.selection) || candidate.text != self.after {
            return Err(ChartCaptionError::Verification);
        }
        budget.charge_validation_scan(&package, 1)?;
        verify_caption_candidate(
            self.source,
            &package,
            &self.selection,
            &candidate,
            self.after.as_deref(),
            self.selection.text != self.after,
            &mut budget,
        )?;
        let target = physical_catalog(&package)?.shared_source();
        budget.charge_exact_artifacts(source_bytes.len(), target.len())?;
        Ok(ChartCaptionCommit {
            package,
            patch: ChartCaptionPatch {
                artifacts: ExactArtifacts::new(source_bytes, target),
                selection: self.selection,
                target_selection: candidate,
                after: self.after,
                touched_components,
                deleted_previews,
                target_requires_invalidated_previews: true,
            },
            diagnostics: ChartCaptionDiagnostics::published(touched_components, deleted_previews),
        })
    }
}

/// An exact-source-checked reversible semantic chart-caption patch.
#[derive(Clone, PartialEq, Eq)]
pub struct ChartCaptionPatch {
    artifacts: ExactArtifacts,
    selection: CaptionSelection,
    target_selection: CaptionSelection,
    after: Option<String>,
    touched_components: usize,
    deleted_previews: usize,
    target_requires_invalidated_previews: bool,
}

impl fmt::Debug for ChartCaptionPatch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("ChartCaptionPatch")
            .field("slide_position", &self.selection.slide_position)
            .field("chart_position", &self.selection.chart_position)
            .field("has_before", &self.selection.text.is_some())
            .field("has_after", &self.after.is_some())
            .finish_non_exhaustive()
    }
}

impl ChartCaptionPatch {
    /// Return the selected semantic slide position.
    #[must_use]
    pub const fn slide_position(&self) -> Position {
        self.selection.slide_position
    }

    /// Return the selected semantic chart position within its slide.
    #[must_use]
    pub const fn chart_position(&self) -> Position {
        self.selection.chart_position
    }

    /// Borrow the caption required from the source package.
    #[must_use]
    pub fn before(&self) -> Option<&str> {
        self.selection.text.as_deref()
    }

    /// Borrow the caption produced by the target package.
    #[must_use]
    pub fn after(&self) -> Option<&str> {
        self.after.as_deref()
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

    /// Return whether this patch preserves semantic caption state and bytes.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.selection == self.target_selection && self.artifacts.is_byte_noop()
    }

    /// Return an exact reversible patch from target back to source.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            selection: self.target_selection.clone(),
            target_selection: self.selection.clone(),
            after: self.selection.text.clone(),
            touched_components: self.touched_components,
            deleted_previews: 0,
            target_requires_invalidated_previews: self.touched_components != 0
                && !self.target_requires_invalidated_previews,
        }
    }
}

/// Compact evidence describing one chart-caption commit.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ChartCaptionDiagnostics {
    changed: bool,
    touched_components: usize,
    deleted_previews: usize,
    full_reparse_performed: bool,
}

impl ChartCaptionDiagnostics {
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

/// The fully verified result of one immutable chart-caption transaction.
#[must_use = "a Keynote chart-caption commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct ChartCaptionCommit {
    package: Package,
    patch: ChartCaptionPatch,
    diagnostics: ChartCaptionDiagnostics,
}

impl ChartCaptionCommit {
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
    pub const fn patch(&self) -> &ChartCaptionPatch {
        &self.patch
    }

    /// Borrow compact publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &ChartCaptionDiagnostics {
        &self.diagnostics
    }
}

impl Package {
    /// Read the caption text of one selected chart.
    ///
    /// `None` means that the chart has no native caption storage; `Some("")`
    /// is an existing empty storage.
    pub fn slide_chart_caption<'slide, 'chart>(
        &self,
        slide_selector: impl Into<SlideSelector<'slide>>,
        chart_selector: impl Into<ChartSelector<'chart>>,
    ) -> Result<Option<String>, ChartCaptionError> {
        Ok(select_caption(self, slide_selector.into(), chart_selector.into(), true)?.text)
    }

    /// Start an exact immutable edit of one selected chart caption.
    pub fn edit_slide_chart_caption<'slide, 'chart>(
        &self,
        slide_selector: impl Into<SlideSelector<'slide>>,
        chart_selector: impl Into<ChartSelector<'chart>>,
    ) -> Result<ChartCaptionEdit<'_>, ChartCaptionError> {
        ChartCaptionEdit::new(self, slide_selector, chart_selector)
    }

    /// Apply an exact-source-checked chart-caption patch.
    pub fn apply_slide_chart_caption(
        &self,
        patch: &ChartCaptionPatch,
    ) -> Result<ChartCaptionCommit, ChartCaptionError> {
        let catalog = physical_catalog(self)?;
        let source = catalog.shared_source();
        if !patch.artifacts.authorizes_source(&source) {
            return Err(ChartCaptionError::PatchConflict);
        }
        let current = select_caption(
            self,
            SlideSelector::position(patch.selection.slide_position),
            ChartSelector::index(patch.selection.chart_position.get()),
            false,
        )?;
        if !current.same_identity(&patch.selection) || current.text != patch.selection.text {
            return Err(ChartCaptionError::PatchConflict);
        }
        if patch.is_noop() {
            self.validate().map_err(map_read_error)?;
            return Ok(ChartCaptionCommit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: ChartCaptionDiagnostics::unchanged(),
            });
        }
        if !catalog.source_is_exact() {
            return Err(ChartCaptionError::PatchConflict);
        }
        let guarded = select_caption(
            self,
            SlideSelector::position(patch.selection.slide_position),
            ChartSelector::index(patch.selection.chart_position.get()),
            true,
        )?;
        if !guarded.same_identity(&patch.selection) || guarded.text != patch.selection.text {
            return Err(ChartCaptionError::PatchConflict);
        }
        let mut budget = CaptionBudget::for_package(self)?;
        budget.charge_catalog_scan(self)?;
        budget.charge_selection_scan(self, 1)?;
        budget.charge_exact_artifacts(source.len(), patch.artifacts.target().len())?;
        budget.charge_candidate_reopen(patch.artifacts.target().len())?;
        let candidate =
            Package::from_source_with_options(patch.artifacts.target(), self.state.options)
                .map_err(map_read_error)?;
        candidate.validate().map_err(map_read_error)?;
        budget.charge_selection_scan(&candidate, 1)?;
        let selected = select_caption(
            &candidate,
            SlideSelector::position(patch.selection.slide_position),
            ChartSelector::index(patch.selection.chart_position.get()),
            true,
        )?;
        if !selected.same_identity(&patch.target_selection) || selected.text != patch.after {
            return Err(ChartCaptionError::Verification);
        }
        budget.charge_validation_scan(&candidate, 1)?;
        verify_caption_candidate(
            self,
            &candidate,
            &patch.selection,
            &patch.target_selection,
            patch.after.as_deref(),
            patch.target_requires_invalidated_previews,
            &mut budget,
        )?;
        Ok(ChartCaptionCommit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: ChartCaptionDiagnostics::published(
                patch.touched_components,
                patch.deleted_previews,
            ),
        })
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct CaptionSelection {
    slide_position: Position,
    chart_position: Position,
    slide_identifier: u64,
    slide_node_identifier: u64,
    chart_identifier: u64,
    non_style_identifier: u64,
    slide_component_name: String,
    reference_identifier: Option<u64>,
    caption_info_identifier: Option<u64>,
    storage_identifier: Option<u64>,
    placement_identifier: Option<u64>,
    style_identifier: Option<u64>,
    text: Option<String>,
}

impl CaptionSelection {
    fn same_chart_identity(&self, other: &Self) -> bool {
        self.slide_position == other.slide_position
            && self.chart_position == other.chart_position
            && self.slide_identifier == other.slide_identifier
            && self.slide_node_identifier == other.slide_node_identifier
            && self.chart_identifier == other.chart_identifier
            && self.non_style_identifier == other.non_style_identifier
            && self.slide_component_name == other.slide_component_name
    }

    fn same_identity(&self, other: &Self) -> bool {
        self.same_chart_identity(other)
            && self.reference_identifier == other.reference_identifier
            && self.caption_info_identifier == other.caption_info_identifier
            && self.storage_identifier == other.storage_identifier
            && self.placement_identifier == other.placement_identifier
            && self.style_identifier == other.style_identifier
    }
}

fn select_caption(
    package: &Package,
    slide_selector: SlideSelector<'_>,
    chart_selector: ChartSelector<'_>,
    mutation_guards: bool,
) -> Result<CaptionSelection, ChartCaptionError> {
    let chart = select_chart(package, slide_selector, chart_selector, mutation_guards)
        .map_err(map_chart_title_error)?;
    caption_selection(package, chart, mutation_guards)
}

fn caption_selection(
    package: &Package,
    chart: ChartSelection,
    mutation_guards: bool,
) -> Result<CaptionSelection, ChartCaptionError> {
    let record = package
        .slide_record_at(chart.slide_position.get())
        .map_err(map_read_error)?
        .ok_or(ChartCaptionError::InvalidSource)?;
    if record.slide_identifier != chart.slide_identifier {
        return Err(ChartCaptionError::InvalidSource);
    }
    let (component, object) = package
        .object_with_component(chart.chart_identifier)
        .ok_or(ChartCaptionError::InvalidSource)?;
    if component != chart.slide_component_name || object.messages.len() != 1 {
        return Err(ChartCaptionError::InvalidSource);
    }
    let payload = unique_payload(&object.messages, &[CHART_MESSAGE_TYPE], "Keynote chart")
        .map_err(map_read_error)?;
    let reference_identifier = chart_caption_reference(package, payload)?;
    let empty = |reference_identifier| CaptionSelection {
        slide_position: chart.slide_position,
        chart_position: chart.chart_position,
        slide_identifier: chart.slide_identifier,
        slide_node_identifier: record.node_identifier,
        chart_identifier: chart.chart_identifier,
        non_style_identifier: chart.non_style_identifier,
        slide_component_name: chart.slide_component_name.clone(),
        reference_identifier,
        caption_info_identifier: None,
        storage_identifier: None,
        placement_identifier: None,
        style_identifier: None,
        text: None,
    };
    let Some(reference_identifier) = reference_identifier else {
        return Ok(empty(None));
    };
    let (info_component, info_object) = package
        .object_with_component(reference_identifier)
        .ok_or(ChartCaptionError::InvalidSource)?;
    if info_component != chart.slide_component_name || info_object.messages.len() != 1 {
        return Err(ChartCaptionError::InvalidSource);
    }
    if info_object.messages[0].type_ == STANDIN_MESSAGE_TYPE {
        validate_selected_message_metadata(info_object, 0)?;
        if !info_object.messages[0].data.is_empty() {
            return Err(ChartCaptionError::InvalidSource);
        }
        if mutation_guards {
            prove_exclusive_caption_standin(package, chart.chart_identifier, reference_identifier)?;
        }
        return Ok(empty(Some(reference_identifier)));
    }
    if info_object.messages[0].type_ != CAPTION_INFO_MESSAGE_TYPE {
        return Err(ChartCaptionError::InvalidSource);
    }
    validate_selected_message_metadata(info_object, 0)?;
    let snapshot = pages_movie_caption_codec::decode_caption_info(
        &info_object.messages[0].data,
        caption_info_decode_options(package, &info_object.messages[0].data)?,
    )
    .map_err(map_caption_info_codec_error)?;
    let storage_identifier = snapshot
        .owned_storage_identifier()
        .ok_or(ChartCaptionError::InvalidSource)?;
    if storage_identifier == 0
        || snapshot.deprecated_storage_identifier() != Some(storage_identifier)
        || snapshot.parent_identifier() != chart.chart_identifier
        || snapshot.is_text_box() != Some(true)
        || snapshot.child_info_kind() != Some(1)
    {
        return Err(ChartCaptionError::InvalidSource);
    }
    let placement_identifier = snapshot
        .placement_identifier()
        .filter(|identifier| *identifier != 0)
        .ok_or(ChartCaptionError::InvalidSource)?;
    let style_identifier = snapshot
        .style_identifier()
        .filter(|identifier| *identifier != 0)
        .ok_or(ChartCaptionError::InvalidSource)?;
    let identifiers = [
        chart.chart_identifier,
        reference_identifier,
        storage_identifier,
        placement_identifier,
        style_identifier,
    ];
    for (index, identifier) in identifiers.iter().enumerate() {
        if identifiers[..index].contains(identifier) {
            return Err(ChartCaptionError::InvalidSource);
        }
    }
    require_private_object(
        package,
        storage_identifier,
        STORAGE_MESSAGE_TYPE,
        Some(chart.slide_component_name.as_str()),
    )?;
    require_private_object(
        package,
        placement_identifier,
        CAPTION_PLACEMENT_MESSAGE_TYPE,
        Some(chart.slide_component_name.as_str()),
    )?;
    require_private_object(
        package,
        style_identifier,
        SHAPE_STYLE_MESSAGE_TYPE,
        Some(chart.slide_component_name.as_str()),
    )?;
    if mutation_guards {
        prove_exclusive_caption_storage(
            package,
            chart.chart_identifier,
            reference_identifier,
            storage_identifier,
        )?;
    }
    let text = super::slide_text::read_owned_storage_text(package, storage_identifier)
        .map_err(map_slide_text_error)?;
    Ok(CaptionSelection {
        slide_position: chart.slide_position,
        chart_position: chart.chart_position,
        slide_identifier: chart.slide_identifier,
        slide_node_identifier: record.node_identifier,
        chart_identifier: chart.chart_identifier,
        non_style_identifier: chart.non_style_identifier,
        slide_component_name: chart.slide_component_name,
        reference_identifier: Some(reference_identifier),
        caption_info_identifier: Some(reference_identifier),
        storage_identifier: Some(storage_identifier),
        placement_identifier: Some(placement_identifier),
        style_identifier: Some(style_identifier),
        text: Some(text),
    })
}

pub(super) fn require_private_object(
    package: &Package,
    identifier: u64,
    message_type: u32,
    expected_component: Option<&str>,
) -> Result<(), ChartCaptionError> {
    let (component, object) = package
        .object_with_component(identifier)
        .ok_or(ChartCaptionError::InvalidSource)?;
    if expected_component.is_some_and(|expected| expected != component)
        || object.messages.len() != 1
        || object.messages[0].type_ != message_type
    {
        return Err(ChartCaptionError::InvalidSource);
    }
    validate_selected_message_metadata(object, 0)
}

pub(super) fn validate_selected_message_metadata(
    object: &ArchiveObject,
    message_index: usize,
) -> Result<(), ChartCaptionError> {
    let message = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(ChartCaptionError::InvalidSource)?;
    if object.archive_info.should_merge == Some(true)
        || message.base_message_index.is_some()
        || !message.diff_merge_version.is_empty()
        || message.diff_field_path.is_some()
        || !message.fields_to_remove.is_empty()
        || !message.diff_read_version.is_empty()
    {
        return Err(ChartCaptionError::InvalidSource);
    }
    Ok(())
}

pub(super) fn prove_exclusive_caption_storage(
    package: &Package,
    chart_identifier: u64,
    caption_info_identifier: u64,
    storage_identifier: u64,
) -> Result<(), ChartCaptionError> {
    prove_exclusive_caption_info_owner(package, chart_identifier, caption_info_identifier)?;
    let mut payload_owner_seen = false;
    let mut metadata_owner_seen = false;
    let mut chart_edge_seen = false;
    for component in package.state.source.components().iter() {
        for object in &component.archive().objects {
            let owner_identifier = object
                .archive_info
                .identifier
                .ok_or(ChartCaptionError::InvalidSource)?;
            for (message_index, message) in object.messages.iter().enumerate() {
                if message.type_ == CHART_MESSAGE_TYPE
                    && chart_caption_reference(package, &message.data)?
                        == Some(caption_info_identifier)
                {
                    if std::mem::replace(&mut chart_edge_seen, true)
                        || owner_identifier != chart_identifier
                    {
                        return Err(ChartCaptionError::UnsupportedDependency);
                    }
                }
                if message.type_ == CAPTION_INFO_MESSAGE_TYPE {
                    let snapshot = pages_movie_caption_codec::decode_caption_info(
                        &message.data,
                        caption_info_decode_options(package, &message.data)?,
                    )
                    .map_err(map_caption_info_codec_error)?;
                    let owns = snapshot.owned_storage_identifier() == Some(storage_identifier)
                        || snapshot.deprecated_storage_identifier() == Some(storage_identifier);
                    if owns
                        && (std::mem::replace(&mut payload_owner_seen, true)
                            || owner_identifier != caption_info_identifier)
                    {
                        return Err(ChartCaptionError::UnsupportedDependency);
                    }
                }
                let info = object
                    .archive_info
                    .message_infos
                    .get(message_index)
                    .ok_or(ChartCaptionError::InvalidSource)?;
                let aggregate_count = info
                    .object_references
                    .iter()
                    .filter(|identifier| **identifier == storage_identifier)
                    .count();
                let aggregate_data = info
                    .data_references
                    .iter()
                    .filter(|identifier| **identifier == storage_identifier)
                    .count();
                let mut field_count = 0usize;
                let mut field_data = 0usize;
                for field in &info.field_infos {
                    let occurrences = field
                        .object_references
                        .iter()
                        .filter(|identifier| **identifier == storage_identifier)
                        .count();
                    if occurrences != 0 && !matches!(field.path.as_slice(), [1, 2] | [1, 4]) {
                        return Err(ChartCaptionError::UnsupportedDependency);
                    }
                    field_count = field_count
                        .checked_add(occurrences)
                        .ok_or(ChartCaptionError::InvalidSource)?;
                    field_data = field_data
                        .checked_add(
                            field
                                .data_references
                                .iter()
                                .filter(|identifier| **identifier == storage_identifier)
                                .count(),
                        )
                        .ok_or(ChartCaptionError::InvalidSource)?;
                }
                if aggregate_count != 0
                    || field_count != 0
                    || aggregate_data != 0
                    || field_data != 0
                {
                    if owner_identifier != caption_info_identifier
                        || message.type_ != CAPTION_INFO_MESSAGE_TYPE
                        || std::mem::replace(&mut metadata_owner_seen, true)
                        || aggregate_count != 1
                        || field_count > 2
                        || aggregate_data != 0
                        || field_data != 0
                    {
                        return Err(ChartCaptionError::UnsupportedDependency);
                    }
                }
            }
        }
    }
    if payload_owner_seen && metadata_owner_seen && chart_edge_seen {
        Ok(())
    } else {
        Err(ChartCaptionError::InvalidSource)
    }
}

pub(super) fn prove_exclusive_caption_standin(
    package: &Package,
    chart_identifier: u64,
    standin_identifier: u64,
) -> Result<(), ChartCaptionError> {
    prove_exclusive_caption_info_owner(package, chart_identifier, standin_identifier)?;
    let mut payload_edges = 0usize;
    let mut aggregate_edges = 0usize;
    let mut field_edges = 0usize;
    for component in package.state.source.components().iter() {
        for object in &component.archive().objects {
            let owner_identifier = object
                .archive_info
                .identifier
                .ok_or(ChartCaptionError::InvalidSource)?;
            for (message_index, message) in object.messages.iter().enumerate() {
                if message.type_ == CHART_MESSAGE_TYPE
                    && chart_caption_reference(package, &message.data)? == Some(standin_identifier)
                {
                    payload_edges = payload_edges
                        .checked_add(1)
                        .ok_or(ChartCaptionError::InvalidSource)?;
                    if payload_edges > 1 || owner_identifier != chart_identifier {
                        return Err(ChartCaptionError::UnsupportedDependency);
                    }
                }
                let info = object
                    .archive_info
                    .message_infos
                    .get(message_index)
                    .ok_or(ChartCaptionError::InvalidSource)?;
                let aggregate_count = info
                    .object_references
                    .iter()
                    .filter(|identifier| **identifier == standin_identifier)
                    .count();
                let aggregate_data = info
                    .data_references
                    .iter()
                    .filter(|identifier| **identifier == standin_identifier)
                    .count();
                if aggregate_data != 0 {
                    return Err(ChartCaptionError::UnsupportedDependency);
                }
                if aggregate_count != 0 {
                    if owner_identifier != chart_identifier
                        || message.type_ != CHART_MESSAGE_TYPE
                        || aggregate_count != 1
                    {
                        return Err(ChartCaptionError::UnsupportedDependency);
                    }
                    aggregate_edges = aggregate_edges
                        .checked_add(aggregate_count)
                        .ok_or(ChartCaptionError::InvalidSource)?;
                }
                for field in &info.field_infos {
                    let field_count = field
                        .object_references
                        .iter()
                        .filter(|identifier| **identifier == standin_identifier)
                        .count();
                    let field_data = field
                        .data_references
                        .iter()
                        .filter(|identifier| **identifier == standin_identifier)
                        .count();
                    if field_data != 0 {
                        return Err(ChartCaptionError::UnsupportedDependency);
                    }
                    if field_count != 0 {
                        if owner_identifier != chart_identifier
                            || message.type_ != CHART_MESSAGE_TYPE
                            || field_count != 1
                            || !matches!(
                                field.path.path.as_slice(),
                                [11, 1] | [1, 11, 1] | [1, 1, 11, 1]
                            )
                        {
                            return Err(ChartCaptionError::UnsupportedDependency);
                        }
                        field_edges = field_edges
                            .checked_add(field_count)
                            .ok_or(ChartCaptionError::InvalidSource)?;
                    }
                }
            }
        }
    }
    if payload_edges == 1 && aggregate_edges == 1 && field_edges <= 1 {
        Ok(())
    } else {
        Err(ChartCaptionError::InvalidSource)
    }
}

/// Prove that the chart-caption edge and its physical metadata owner are
/// exclusive.  A future message or nested FieldInfo that points at CaptionInfo
/// is not safe to ignore: replacing the chart edge would otherwise orphan a
/// live owner or leave an unaccounted alias in the graph.
fn prove_exclusive_caption_info_owner(
    package: &Package,
    chart_identifier: u64,
    caption_info_identifier: u64,
) -> Result<(), ChartCaptionError> {
    let mut payload_edges = 0usize;
    let mut aggregate_edges = 0usize;
    let mut field_edges = 0usize;
    for component in package.state.source.components().iter() {
        for object in &component.archive().objects {
            let owner_identifier = object
                .archive_info
                .identifier
                .ok_or(ChartCaptionError::InvalidSource)?;
            for (message_index, message) in object.messages.iter().enumerate() {
                if message.type_ == CHART_MESSAGE_TYPE
                    && chart_caption_reference(package, &message.data)?
                        == Some(caption_info_identifier)
                {
                    payload_edges = payload_edges
                        .checked_add(1)
                        .ok_or(ChartCaptionError::InvalidSource)?;
                    if owner_identifier != chart_identifier || payload_edges > 1 {
                        return Err(ChartCaptionError::UnsupportedDependency);
                    }
                }
                let info = object
                    .archive_info
                    .message_infos
                    .get(message_index)
                    .ok_or(ChartCaptionError::InvalidSource)?;
                let aggregate_count = info
                    .object_references
                    .iter()
                    .filter(|identifier| **identifier == caption_info_identifier)
                    .count();
                let aggregate_data = info
                    .data_references
                    .iter()
                    .filter(|identifier| **identifier == caption_info_identifier)
                    .count();
                if aggregate_data != 0 {
                    return Err(ChartCaptionError::UnsupportedDependency);
                }
                let mut local_field_edges = 0usize;
                for field in &info.field_infos {
                    let field_count = field
                        .object_references
                        .iter()
                        .filter(|identifier| **identifier == caption_info_identifier)
                        .count();
                    let field_data = field
                        .data_references
                        .iter()
                        .filter(|identifier| **identifier == caption_info_identifier)
                        .count();
                    if field_data != 0 {
                        return Err(ChartCaptionError::UnsupportedDependency);
                    }
                    if field_count != 0 {
                        if owner_identifier != chart_identifier
                            || message.type_ != CHART_MESSAGE_TYPE
                            || field_count != 1
                            || !matches!(
                                field.path.path.as_slice(),
                                [11, 1] | [1, 11, 1] | [1, 1, 11, 1]
                            )
                        {
                            return Err(ChartCaptionError::UnsupportedDependency);
                        }
                        local_field_edges = local_field_edges
                            .checked_add(field_count)
                            .ok_or(ChartCaptionError::InvalidSource)?;
                    }
                }
                if aggregate_count != 0 {
                    if owner_identifier != chart_identifier
                        || message.type_ != CHART_MESSAGE_TYPE
                        || aggregate_count != 1
                    {
                        return Err(ChartCaptionError::UnsupportedDependency);
                    }
                }
                if aggregate_count != 0 || local_field_edges != 0 {
                    aggregate_edges = aggregate_edges
                        .checked_add(aggregate_count)
                        .ok_or(ChartCaptionError::InvalidSource)?;
                    field_edges = field_edges
                        .checked_add(local_field_edges)
                        .ok_or(ChartCaptionError::InvalidSource)?;
                }
            }
        }
    }
    if payload_edges == 1 && aggregate_edges == 1 && field_edges <= 1 {
        Ok(())
    } else {
        Err(ChartCaptionError::InvalidSource)
    }
}

fn chart_caption_reference(
    package: &Package,
    payload: &[u8],
) -> Result<Option<u64>, ChartCaptionError> {
    let options = chart_caption_decode_options(package, payload)?;
    let identifier = keynote_chart_caption_codec::decode_chart_caption_identifier(payload, options)
        .map_err(map_chart_caption_codec_error)?;
    let Some(identifier) = identifier else {
        return Ok(None);
    };
    if identifier == 0 {
        return Err(ChartCaptionError::InvalidSource);
    }
    let limits = package.wire_limits().map_err(map_wire_error)?;
    let outer = WireView::parse_with_limits(payload, limits).map_err(map_wire_error)?;
    let super_payload = unique_payload_field(&outer, 1)?.ok_or(ChartCaptionError::InvalidSource)?;
    let drawable = WireView::parse_with_limits(super_payload, limits).map_err(map_wire_error)?;
    let reference = unique_payload_field(&drawable, 11)?.ok_or(ChartCaptionError::InvalidSource)?;
    let projected = super::validate_reference_payload(reference, limits, "Keynote chart caption")
        .map_err(map_wire_error)?;
    let strict_identifier = reference_identifier(reference, limits)?;
    if projected != identifier || strict_identifier != identifier {
        return Err(ChartCaptionError::InvalidSource);
    }
    Ok(Some(identifier))
}

fn validate_reference_optional_fields(view: &WireView<'_>) -> Result<(), ChartCaptionError> {
    let mut deprecated_type_seen = false;
    let mut deprecated_external_seen = false;
    for field in view.fields() {
        if !matches!(field.number(), 2 | 3) {
            continue;
        }
        if field.wire_type() != 0 {
            return Err(ChartCaptionError::InvalidSource);
        }
        field.validate_canonical_key().map_err(map_wire_error)?;
        let (value, bytes) = decode_varint_from_bytes(field.payload())
            .map_err(|_error| ChartCaptionError::InvalidSource)?;
        if bytes != field.payload().len() || bytes != encoded_len(value) || value != 0 {
            return Err(ChartCaptionError::InvalidSource);
        }
        if field.number() == 2 {
            if std::mem::replace(&mut deprecated_type_seen, true) {
                return Err(ChartCaptionError::InvalidSource);
            }
        } else if std::mem::replace(&mut deprecated_external_seen, true) {
            return Err(ChartCaptionError::InvalidSource);
        }
    }
    Ok(())
}

fn unique_payload_field<'a>(
    view: &WireView<'a>,
    number: u32,
) -> Result<Option<&'a [u8]>, ChartCaptionError> {
    let mut selected = None;
    for field in view.fields().filter(|field| field.number() == number) {
        if selected.is_some() || field.wire_type() != 2 {
            return Err(ChartCaptionError::InvalidSource);
        }
        field.validate_canonical_key().map_err(map_wire_error)?;
        field.validate_canonical_framing().map_err(map_wire_error)?;
        selected = Some(field.payload());
    }
    Ok(selected)
}

/// Execute the graph transition used when a chart moves between its canonical
/// empty stand-in and an inline CaptionInfo graph.  Existing storage edits are
/// intentionally delegated to `slide_text`; this path only adds a graph or a
/// fresh stand-in and never removes the old graph.
fn rewrite_chart_caption_operation(
    source: &Package,
    selection: &CaptionSelection,
    after: Option<&str>,
    budget: &mut CaptionBudget,
) -> Result<(Package, usize, usize), ChartCaptionError> {
    let creating = selection.storage_identifier.is_none() && after.is_some();
    let removing = selection.storage_identifier.is_some() && after.is_none();
    if !creating && !removing {
        let storage = selection
            .storage_identifier
            .ok_or(ChartCaptionError::UnsupportedDependency)?;
        let desired = after.ok_or(ChartCaptionError::UnsupportedDependency)?;
        let end = selection
            .text
            .as_deref()
            .ok_or(ChartCaptionError::InvalidSource)?
            .encode_utf16()
            .count();
        if selection
            .text
            .as_deref()
            .is_some_and(contains_dependent_marker)
            || contains_dependent_marker(desired)
        {
            return Err(ChartCaptionError::UnsupportedDependency);
        }
        if metadata_member_name(physical_catalog(source)?, source).is_ok() {
            return rewrite_existing_caption_text_with_metadata_budget(
                source,
                storage,
                selection.slide_node_identifier,
                &selection.slide_component_name,
                end,
                desired,
                budget,
            );
        }
        return Err(ChartCaptionError::InvalidSource);
    }

    rewrite_caption_graph_operation_with_budget(
        source,
        &selection.slide_component_name,
        selection.chart_identifier,
        selection.reference_identifier,
        selection.storage_identifier,
        after,
        CaptionEdgeKind::Chart,
        budget,
    )
}

/// Execute the canonical graph transition for any supported drawable edge.
///
/// Creation retains the selected stand-in and appends four canonical graph
/// objects. Removal retains the old graph and appends one fresh stand-in. The
/// caller owns the operation budget so chart and movie transactions account
/// the same physical lifecycle and can reject before publication.
pub(super) fn rewrite_caption_graph_operation_with_budget(
    source: &Package,
    slide_component_name: &str,
    drawable_identifier: u64,
    reference_identifier: Option<u64>,
    storage_identifier: Option<u64>,
    after: Option<&str>,
    edge_kind: CaptionEdgeKind,
    budget: &mut CaptionBudget,
) -> Result<(Package, usize, usize), ChartCaptionError> {
    let creating = storage_identifier.is_none() && after.is_some();
    let removing = storage_identifier.is_some() && after.is_none();
    if !creating && !removing {
        return Err(ChartCaptionError::UnsupportedDependency);
    }
    if after.is_some_and(contains_dependent_marker) {
        return Err(ChartCaptionError::UnsupportedDependency);
    }

    // The graph transition performs package-wide metadata, identifier, theme,
    // and dependency scans below.  Account the physical ownership traversal
    // before any candidate archive allocation so those scans share the same
    // operation-local budget as the codec and reassembler.
    budget.charge_validation_scan(source, 1)?;

    let catalog = physical_catalog(source)?;
    let metadata_name = metadata_member_name(catalog, source)?;
    let slide_name = slide_component_name.to_owned();
    if metadata_name == slide_name {
        return Err(ChartCaptionError::InvalidSource);
    }
    let physical_limits = source.state.options.archive();
    let archive_limits = physical_limits
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let snappy_limits = physical_limits.snappy_limits().map_err(map_archive_error)?;
    budget.charge_allocation(1, 0)?;
    let mut archives = Vec::new();
    archives
        .try_reserve_exact(2)
        .map_err(|_| ChartCaptionError::Allocation { amount: 2 })?;
    for name in [&slide_name, &metadata_name] {
        let entry = catalog
            .package()
            .iter()
            .find(|entry| entry.name() == name.as_str())
            .ok_or(ChartCaptionError::InvalidSource)?;
        if entry.is_opaque() {
            return Err(ChartCaptionError::InvalidSource);
        }
        let stream = SnappyStream::decompress_with_limits(entry.data(), snappy_limits)
            .map_err(map_core_error)?;
        let archive = Archive::parse_with_limits(stream.as_bytes(), archive_limits)
            .map_err(map_core_error)?;
        validate_canonical_object_framing(&archive, stream.as_bytes())?;
        archives.push((name.clone(), archive));
    }

    let metadata_source = metadata_payload(archive_ref(&archives, &metadata_name)?)?;
    let metadata_facts = caption_metadata_facts(
        source,
        &metadata_source.2,
        metadata_locator(&slide_name),
        budget,
    )?;
    budget.charge_metadata_report(
        metadata_facts
            .inspection_report
            .ok_or(ChartCaptionError::InvalidSource)?,
    )?;
    let slide_selector = metadata_facts.selector()?;
    prove_metadata_selector_component(source, &slide_name, slide_selector)?;
    let first_identifier = next_caption_identifier(source, &metadata_facts)?;
    let (new_identifiers, replacement_identifier, graph_objects, graph_style_identifier) =
        if creating {
            let first = first_identifier;
            let ids = CaptionGraphIds::allocate(first)?;
            let theme = caption_theme(source)?;
            prove_caption_dependencies(
                source,
                &metadata_facts,
                &slide_name,
                [theme.stylesheet, theme.paragraph_style],
            )?;
            let width = caption_drawable_width(source, drawable_identifier)?;
            budget.precharge_graph(
                after.ok_or(ChartCaptionError::InvalidSource)?,
                theme.language.as_deref(),
            )?;
            let (objects, graph_report) = caption_graph_objects(
                source,
                ids,
                drawable_identifier,
                width,
                after.ok_or(ChartCaptionError::InvalidSource)?,
                theme.stylesheet,
                theme.paragraph_style,
                theme.language.as_deref(),
                edge_kind,
                budget,
            )?;
            budget.observe_graph_report(graph_report)?;
            (
                ids,
                ids.info,
                objects,
                matches!(
                    edge_kind,
                    CaptionEdgeKind::Movie | CaptionEdgeKind::MovieTitle
                )
                .then_some(ids.style),
            )
        } else {
            let standin = first_identifier;
            budget.charge_allocation(1, 0)?;
            let object = canonical_standin(standin)?;
            (
                CaptionGraphIds::standin(standin),
                standin,
                vec![object],
                None,
            )
        };

    let slide_archive = archive_mut(&mut archives, &slide_name)?;
    patch_caption_edge(
        slide_archive,
        drawable_identifier,
        reference_identifier,
        replacement_identifier,
        source,
        archive_limits,
        budget,
        edge_kind,
        graph_style_identifier,
    )?;
    for object in graph_objects {
        slide_archive
            .insert_object_with_limits(object, archive_limits)
            .map_err(map_core_error)?;
    }

    let new_last = new_identifiers.last();
    budget.charge_allocation(1, 0)?;
    let mut additions = Vec::new();
    additions
        .try_reserve_exact(new_identifiers.all().len())
        .map_err(|_| ChartCaptionError::Allocation {
            amount: new_identifiers.all().len(),
        })?;
    for identifier in new_identifiers.all() {
        if identifier == 0 {
            continue;
        }
        additions.push(ObjectUuidAddition::new(
            slide_selector,
            identifier,
            fresh_uuid(&metadata_facts.uuids, identifier),
        ));
    }
    let selectors = [slide_selector];
    let token_batch = SaveTokenBatch::new(&selectors);
    let addition_batch =
        MetadataBatch::new(metadata_facts.last_identifier, new_last, &additions, &[]);
    let prepared_metadata = prepare_package_metadata_additions_and_save_tokens(
        &metadata_source.2,
        AdditionSaveTokenBatch::new(addition_batch, token_batch),
        metadata_options(source, budget, additions.len())?,
    )
    .map_err(map_metadata_error)?;
    budget.charge_metadata_report(prepared_metadata.prepare_report())?;
    let metadata_requirements = prepared_metadata.execution_requirements();
    budget.charge_execution_requirements(metadata_requirements)?;
    let metadata_output = prepared_metadata
        .execute(metadata_requirements.exact_limits())
        .map_err(map_metadata_error)?;
    budget.observe_metadata_execution(metadata_output.report(), metadata_requirements)?;
    let metadata_output = metadata_output.into_bytes();
    replace_metadata_payload(
        archive_mut(&mut archives, &metadata_name)?,
        metadata_source.0,
        metadata_source.1,
        metadata_output,
        archive_limits,
    )?;

    for (_, archive) in &archives {
        budget.charge_archive_snappy_plan(archive, archive_limits, snappy_limits)?;
    }
    budget.charge_allocation(1, 0)?;
    let mut compressed = Vec::new();
    compressed
        .try_reserve_exact(archives.len())
        .map_err(|_error| ChartCaptionError::Allocation {
            amount: archives.len(),
        })?;
    for (name, archive) in &archives {
        let bytes = archive
            .to_bytes_with_limits(archive_limits)
            .map_err(map_core_error)?;
        let data = SnappyStream::compress(&bytes).map_err(map_core_error)?;
        compressed.push((name.clone(), data));
    }
    budget.charge_allocation(1, 0)?;
    let mut edits = Vec::new();
    edits
        .try_reserve_exact(compressed.len())
        .map_err(|_| ChartCaptionError::Allocation {
            amount: compressed.len(),
        })?;
    for (name, data) in &compressed {
        edits.push(EntryEdit::new(name.as_str(), data.as_slice()));
    }
    let previews = super::rendering_invalidation::root_preview_deletions(catalog.package())
        .map_err(map_rendering_error)?;
    let prepared_reassembly = catalog
        .prepare_reassembly_with_deletions(&edits, previews.names(), physical_limits)
        .map_err(map_archive_error)?;
    let reassembly_requirements = prepared_reassembly.execution_requirements();
    let reassembly_limits = budget.charge_reassembly(reassembly_requirements)?;
    budget.charge_candidate_reopen(reassembly_requirements.output_bytes())?;
    let output = prepared_reassembly
        .execute(reassembly_limits)
        .map_err(map_archive_error)?;
    let candidate = Package::from_source_with_options(output.into(), source.state.options)
        .map_err(map_read_error)?;
    Ok((candidate, 2, previews.len()))
}

pub(super) fn rewrite_existing_caption_text_with_metadata_budget(
    source: &Package,
    storage_identifier: u64,
    slide_node_identifier: u64,
    slide_component_name: &str,
    end: usize,
    desired: &str,
    budget: &mut CaptionBudget,
) -> Result<(Package, usize, usize), ChartCaptionError> {
    let catalog = physical_catalog(source)?;
    let metadata_name = metadata_member_name(catalog, source)?;
    let deleted = preview_count(catalog);
    // Storage replacement traverses the source graph before staging the native
    // candidate.  Keep that ownership/selector pass visible to the enclosing
    // transaction budget; the Metadata codec report accounts only its own
    // protobuf passes.
    budget.charge_validation_scan(source, 1)?;
    budget.precharge_intermediate_candidate(source, desired.len())?;
    let (native_candidate, touched) = super::slide_text::rewrite_owned_storage_text(
        source,
        storage_identifier,
        slide_node_identifier,
        0..end,
        desired,
    )
    .map_err(map_slide_text_error)?;
    let metadata_archive = archive_for_member(&native_candidate, &metadata_name)?;
    let metadata_source = metadata_payload(&metadata_archive)?;
    let metadata_facts = caption_metadata_facts(
        &native_candidate,
        &metadata_source.2,
        metadata_locator(slide_component_name),
        budget,
    )?;
    budget.charge_metadata_report(
        metadata_facts
            .inspection_report
            .ok_or(ChartCaptionError::InvalidSource)?,
    )?;
    let touched_component_names = exact_component_names_for_identifiers(
        &native_candidate,
        &[storage_identifier, slide_node_identifier],
    )?;
    if !touched_component_names
        .iter()
        .any(|name| name == slide_component_name)
    {
        return Err(ChartCaptionError::InvalidSource);
    }
    let selectors = metadata_selectors_for_component_names(
        &native_candidate,
        &metadata_facts,
        &touched_component_names,
    )?;
    let prepared_metadata = prepare_package_metadata_save_tokens(
        &metadata_source.2,
        SaveTokenBatch::new(selectors.as_slice()),
        metadata_options(&native_candidate, budget, 0)?,
    )
    .map_err(map_metadata_error)?;
    budget.charge_metadata_report(prepared_metadata.prepare_report())?;
    let metadata_requirements = prepared_metadata.execution_requirements();
    budget.charge_execution_requirements(metadata_requirements)?;
    let output = prepared_metadata
        .execute(metadata_requirements.exact_limits())
        .map_err(map_metadata_error)?;
    budget.observe_metadata_execution(output.report(), metadata_requirements)?;
    let output = output.into_bytes();
    let mut metadata_archive = metadata_archive;
    replace_metadata_payload(
        &mut metadata_archive,
        metadata_source.0,
        metadata_source.1,
        output,
        native_candidate
            .state
            .options
            .archive()
            .effective_archive_limits()
            .map_err(map_archive_error)?,
    )?;
    let archive_limits = native_candidate
        .state
        .options
        .archive()
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let physical_limits = native_candidate.state.options.archive();
    let snappy_limits = physical_limits.snappy_limits().map_err(map_archive_error)?;
    budget.charge_archive_snappy_plan(&metadata_archive, archive_limits, snappy_limits)?;
    let metadata_bytes = metadata_archive
        .to_bytes_with_limits(archive_limits)
        .map_err(map_core_error)?;
    let metadata_compressed = SnappyStream::compress(&metadata_bytes).map_err(map_core_error)?;
    let edit = EntryEdit::new(metadata_name.as_str(), metadata_compressed.as_slice());
    let native_catalog = physical_catalog(&native_candidate)?;
    let prepared_reassembly = native_catalog
        .prepare_reassembly_with_deletions(std::slice::from_ref(&edit), &[], physical_limits)
        .map_err(map_archive_error)?;
    let reassembly_requirements = prepared_reassembly.execution_requirements();
    let reassembly_limits = budget.charge_reassembly(reassembly_requirements)?;
    budget.charge_candidate_reopen(reassembly_requirements.output_bytes())?;
    let candidate_bytes = prepared_reassembly
        .execute(reassembly_limits)
        .map_err(map_archive_error)?;
    let candidate = Package::from_source_with_options(candidate_bytes.into(), source.state.options)
        .map_err(map_read_error)?;
    candidate.validate().map_err(map_read_error)?;
    // Reopen the metadata route once more before this candidate is returned;
    // this also enforces canonical object-length prefixes on the changed
    // metadata member.
    let _ = archive_for_member(&candidate, &metadata_name)?;
    Ok((
        candidate,
        touched
            .checked_add(1)
            .ok_or(ChartCaptionError::InvalidSource)?,
        deleted,
    ))
}

fn archive_ref<'a>(
    archives: &'a [(String, Archive)],
    name: &str,
) -> Result<&'a Archive, ChartCaptionError> {
    archives
        .iter()
        .find(|(candidate, _)| candidate == name)
        .map(|(_, archive)| archive)
        .ok_or(ChartCaptionError::InvalidSource)
}

fn archive_mut<'a>(
    archives: &'a mut [(String, Archive)],
    name: &str,
) -> Result<&'a mut Archive, ChartCaptionError> {
    archives
        .iter_mut()
        .find(|(candidate, _)| candidate == name)
        .map(|(_, archive)| archive)
        .ok_or(ChartCaptionError::InvalidSource)
}

fn metadata_member_name(
    catalog: &SourceCatalog,
    package: &Package,
) -> Result<String, ChartCaptionError> {
    let entry = catalog
        .package()
        .iter()
        .find(|entry| entry.name() == PACKAGE_METADATA_MEMBER_NAME)
        .ok_or(ChartCaptionError::InvalidSource)?;
    if entry.is_opaque() {
        return Err(ChartCaptionError::InvalidSource);
    }
    let mut found = 0usize;
    for component in package.state.source.components().iter() {
        for object in &component.archive().objects {
            for message in &object.messages {
                if message.type_ != PACKAGE_METADATA_MESSAGE_TYPE {
                    continue;
                }
                if component.name() != PACKAGE_METADATA_MEMBER_NAME {
                    return Err(ChartCaptionError::InvalidSource);
                }
                found = found
                    .checked_add(1)
                    .ok_or(ChartCaptionError::InvalidSource)?;
            }
        }
    }
    if found == 1 {
        Ok(PACKAGE_METADATA_MEMBER_NAME.to_owned())
    } else {
        Err(ChartCaptionError::InvalidSource)
    }
}

fn metadata_payload(archive: &Archive) -> Result<(u64, usize, Vec<u8>), ChartCaptionError> {
    let mut selected = None;
    for object in &archive.objects {
        let identifier = object
            .archive_info
            .identifier
            .ok_or(ChartCaptionError::InvalidSource)?;
        for (index, message) in object.messages.iter().enumerate() {
            if message.type_ != PACKAGE_METADATA_MESSAGE_TYPE {
                continue;
            }
            if selected
                .replace((identifier, index, message.data.clone()))
                .is_some()
            {
                return Err(ChartCaptionError::InvalidSource);
            }
        }
    }
    selected.ok_or(ChartCaptionError::InvalidSource)
}

fn replace_metadata_payload(
    archive: &mut Archive,
    object_identifier: u64,
    index: usize,
    data: Vec<u8>,
    limits: litchi_iwa_core::Limits,
) -> Result<(), ChartCaptionError> {
    let object = archive
        .object_mut(object_identifier)
        .ok_or(ChartCaptionError::InvalidSource)?;
    if object.messages.get(index).map(|message| message.type_)
        != Some(PACKAGE_METADATA_MESSAGE_TYPE)
    {
        return Err(ChartCaptionError::InvalidSource);
    }
    object
        .replace_message_preserving_header_with_limits(
            index,
            RawMessage {
                type_: PACKAGE_METADATA_MESSAGE_TYPE,
                data,
            },
            limits,
        )
        .map_err(map_core_error)?;
    Ok(())
}

fn metadata_locator(name: &str) -> &str {
    name.strip_prefix("Index/")
        .and_then(|name| name.strip_suffix(".iwa"))
        .unwrap_or(name)
}

fn prove_metadata_selector_component(
    package: &Package,
    component_name: &str,
    selector: ComponentSelector<'_>,
) -> Result<(), ChartCaptionError> {
    if selector.locator() != metadata_locator(component_name) {
        return Err(ChartCaptionError::InvalidSource);
    }
    let mut owners = 0usize;
    for component in package.state.source.components().iter() {
        if component.archive().object(selector.identifier()).is_none() {
            continue;
        }
        if component.name() != component_name {
            return Err(ChartCaptionError::InvalidSource);
        }
        owners = owners
            .checked_add(1)
            .ok_or(ChartCaptionError::InvalidSource)?;
    }
    if owners == 1 {
        Ok(())
    } else {
        Err(ChartCaptionError::InvalidSource)
    }
}

fn exact_component_names_for_identifiers(
    package: &Package,
    identifiers: &[u64],
) -> Result<Vec<String>, ChartCaptionError> {
    let mut names = Vec::new();
    names
        .try_reserve_exact(identifiers.len())
        .map_err(|_| ChartCaptionError::Allocation {
            amount: identifiers.len(),
        })?;
    for identifier in identifiers {
        let mut owners = package
            .state
            .source
            .components()
            .iter()
            .filter(|component| component.archive().object(*identifier).is_some());
        let owner = owners.next().ok_or(ChartCaptionError::InvalidSource)?;
        if owners.next().is_some() {
            return Err(ChartCaptionError::InvalidSource);
        }
        if !names.iter().any(|name| name == owner.name()) {
            names.push(owner.name().to_owned());
        }
    }
    Ok(names)
}

fn metadata_selectors_for_component_names<'facts>(
    package: &Package,
    facts: &'facts CaptionMetadataFacts,
    component_names: &[String],
) -> Result<Vec<ComponentSelector<'facts>>, ChartCaptionError> {
    let mut selectors = Vec::new();
    selectors
        .try_reserve_exact(component_names.len())
        .map_err(|_| ChartCaptionError::Allocation {
            amount: component_names.len(),
        })?;
    for component_name in component_names {
        let locator = metadata_locator(component_name);
        let mut matches = facts
            .components
            .iter()
            .filter(|component| component.current && component.locator == locator);
        let component = matches.next().ok_or(ChartCaptionError::InvalidSource)?;
        if matches.next().is_some() {
            return Err(ChartCaptionError::InvalidSource);
        }
        let selector = ComponentSelector::new(component.identifier, component.locator.as_str());
        prove_metadata_selector_component(package, component_name, selector)?;
        selectors.push(selector);
    }
    Ok(selectors)
}

fn metadata_options(
    package: &Package,
    budget: &CaptionBudget,
    max_additions: usize,
) -> Result<MetadataRewriteOptions, ChartCaptionError> {
    let recursion = budget.remaining_depth();
    let semantic = package.semantic_limits();
    let wire = package.wire_limits().map_err(map_wire_error)?;
    Ok(MetadataRewriteOptions::new(
        budget.remaining_input().min(wire.max_input_bytes()),
        budget.remaining_output().min(wire.max_output_bytes()),
        budget.remaining_fields(),
        budget.remaining_work(),
        recursion,
        budget.remaining_components().min(semantic.max_objects()),
        budget.remaining_references().min(semantic.max_references()),
        max_additions.max(1),
    ))
}

struct CaptionMetadataFacts {
    target_locator: String,
    selected_identifier: Option<u64>,
    selected_locator: Option<String>,
    selected_count: usize,
    uuids: HashSet<(u64, u64)>,
    owned_object_identifiers: HashSet<u64>,
    components: Vec<CaptionMetadataComponent>,
    external_references: Vec<CaptionMetadataExternalReference>,
    last_identifier: u64,
    inspection_report: Option<metadata_codec::RewriteReport>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct CaptionMetadataComponent {
    identifier: u64,
    locator: String,
    current: bool,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct CaptionMetadataExternalReference {
    source_identifier: u64,
    source_locator: String,
    source_current: bool,
    target_component_identifier: u64,
    object_identifier: Option<u64>,
    weak: Option<bool>,
    versioned: bool,
}

impl CaptionMetadataFacts {
    fn selector(&self) -> Result<ComponentSelector<'_>, ChartCaptionError> {
        if self.selected_count != 1 {
            return Err(ChartCaptionError::InvalidSource);
        }
        let locator = self
            .selected_locator
            .as_deref()
            .ok_or(ChartCaptionError::InvalidSource)?;
        Ok(ComponentSelector::new(
            self.selected_identifier
                .ok_or(ChartCaptionError::InvalidSource)?,
            locator,
        ))
    }

    fn selected_component(&self) -> Result<(u64, &str), ChartCaptionError> {
        if self.selected_count != 1 {
            return Err(ChartCaptionError::InvalidSource);
        }
        Ok((
            self.selected_identifier
                .ok_or(ChartCaptionError::InvalidSource)?,
            self.selected_locator
                .as_deref()
                .ok_or(ChartCaptionError::InvalidSource)?,
        ))
    }
}

impl PackageMetadataVisitor for CaptionMetadataFacts {
    fn visit_component(
        &mut self,
        component: ComponentDescriptor<'_>,
    ) -> Result<(), metadata_codec::RewriteError> {
        let locator = component.effective_locator();
        self.components.push(CaptionMetadataComponent {
            identifier: component.identifier(),
            locator: locator.to_owned(),
            current: component.is_current(),
        });
        if component.is_current() && locator == self.target_locator {
            self.selected_identifier = Some(component.identifier());
            self.selected_locator = Some(locator.to_owned());
            self.selected_count = self.selected_count.saturating_add(1);
        }
        Ok(())
    }

    fn visit_object_uuid(
        &mut self,
        binding: metadata_codec::ObjectUuidDescriptor<'_>,
    ) -> Result<(), metadata_codec::RewriteError> {
        self.owned_object_identifiers
            .insert(binding.object_identifier());
        self.uuids
            .insert((binding.uuid().lower(), binding.uuid().upper()));
        Ok(())
    }

    fn visit_external_reference(
        &mut self,
        reference: metadata_codec::ExternalReferenceDescriptor<'_>,
    ) -> Result<(), metadata_codec::RewriteError> {
        if let Some(identifier) = reference.object_identifier() {
            self.owned_object_identifiers.insert(identifier);
        }
        let source = reference.source();
        self.external_references
            .push(CaptionMetadataExternalReference {
                source_identifier: source.identifier(),
                source_locator: source.effective_locator().to_owned(),
                source_current: source.is_current(),
                target_component_identifier: reference.target_component_identifier(),
                object_identifier: reference.object_identifier(),
                weak: reference.is_weak(),
                versioned: reference.is_versioned(),
            });
        Ok(())
    }

    fn visit_data_reference_owner(
        &mut self,
        owner: metadata_codec::DataReferenceOwnerDescriptor<'_>,
    ) -> Result<(), metadata_codec::RewriteError> {
        self.owned_object_identifiers
            .insert(owner.object_identifier());
        Ok(())
    }

    fn visit_ambiguous_object_identifier(
        &mut self,
        _component: ComponentDescriptor<'_>,
        identifier: u64,
    ) -> Result<(), metadata_codec::RewriteError> {
        self.owned_object_identifiers.insert(identifier);
        Ok(())
    }

    fn visit_data_metadata_map(
        &mut self,
        object_identifier: u64,
        _has_unknown_fields: bool,
    ) -> Result<(), metadata_codec::RewriteError> {
        self.owned_object_identifiers.insert(object_identifier);
        Ok(())
    }
}

fn caption_metadata_facts(
    package: &Package,
    source: &[u8],
    target_locator: &str,
    budget: &CaptionBudget,
) -> Result<CaptionMetadataFacts, ChartCaptionError> {
    let options = metadata_options(package, budget, 4)?;
    let mut facts = CaptionMetadataFacts {
        target_locator: target_locator.to_owned(),
        selected_identifier: None,
        selected_locator: None,
        selected_count: 0,
        uuids: HashSet::new(),
        owned_object_identifiers: HashSet::new(),
        components: Vec::new(),
        external_references: Vec::new(),
        last_identifier: 0,
        inspection_report: None,
    };
    let inspection = inspect_package_metadata_with_visitor(source, options, &mut facts)
        .map_err(map_metadata_error)?;
    facts.last_identifier = inspection.last_object_identifier();
    facts.inspection_report = Some(inspection.report());
    Ok(facts)
}

/// Prove that every caption graph dependency which leaves the selected slide
/// component is already represented by one exact current Metadata external
/// reference.  Graph creation is intentionally not allowed to manufacture
/// registry edges: an unregistered foreign stylesheet/theme/paragraph object
/// would leave a package whose physical graph and metadata disagree.
fn prove_caption_dependencies(
    package: &Package,
    facts: &CaptionMetadataFacts,
    slide_component_name: &str,
    dependencies: [u64; 2],
) -> Result<(), ChartCaptionError> {
    let (selected_identifier, selected_locator) = facts.selected_component()?;
    let mut seen = HashSet::new();
    for dependency in dependencies {
        if dependency == 0 || !seen.insert(dependency) {
            continue;
        }
        let mut owner_names = Vec::new();
        for component in package.state.source.components().iter() {
            if component
                .archive()
                .objects
                .iter()
                .any(|object| object.archive_info.identifier == Some(dependency))
            {
                owner_names.push(component.name());
            }
        }
        if owner_names.len() != 1 {
            return Err(ChartCaptionError::UnsupportedDependency);
        }
        let owner_name = owner_names[0];
        if owner_name == slide_component_name {
            continue;
        }
        let (_component_name, _object) = package
            .object_with_component(dependency)
            .ok_or(ChartCaptionError::InvalidSource)?;
        let target_locator = metadata_locator(owner_name);
        let target_components = facts
            .components
            .iter()
            .filter(|component| component.current && component.locator == target_locator)
            .collect::<Vec<_>>();
        // The component identifier is obtained from metadata rather than the
        // package object index.  There must be one current component at the
        // exact effective locator; versioned records never authorize a write.
        if target_components.len() != 1 {
            return Err(ChartCaptionError::UnsupportedDependency);
        }
        let target_identifier = target_components[0].identifier;
        if facts
            .components
            .iter()
            .filter(|component| component.current && component.identifier == target_identifier)
            .count()
            != 1
        {
            return Err(ChartCaptionError::UnsupportedDependency);
        }
        let mut matching = 0usize;
        let mut prohibited = false;
        for reference in &facts.external_references {
            if reference.source_identifier != selected_identifier
                || reference.source_locator != selected_locator
                || reference.target_component_identifier != target_identifier
                || reference.object_identifier != Some(dependency)
            {
                continue;
            }
            if !reference.source_current || reference.versioned || reference.weak == Some(true) {
                prohibited = true;
            } else {
                matching = matching
                    .checked_add(1)
                    .ok_or(ChartCaptionError::InvalidSource)?;
            }
        }
        if prohibited || matching != 1 {
            return Err(ChartCaptionError::UnsupportedDependency);
        }
    }
    Ok(())
}

#[derive(Debug, Clone, Copy)]
struct CaptionGraphIds {
    style: u64,
    info: u64,
    storage: u64,
    placement: u64,
}

impl CaptionGraphIds {
    fn allocate(first: u64) -> Result<Self, ChartCaptionError> {
        Ok(Self {
            style: first,
            info: first
                .checked_add(1)
                .ok_or(ChartCaptionError::InvalidSource)?,
            storage: first
                .checked_add(2)
                .ok_or(ChartCaptionError::InvalidSource)?,
            placement: first
                .checked_add(3)
                .ok_or(ChartCaptionError::InvalidSource)?,
        })
    }

    const fn standin(identifier: u64) -> Self {
        Self {
            style: 0,
            info: 0,
            storage: 0,
            placement: identifier,
        }
    }

    const fn all(self) -> [u64; 4] {
        [self.style, self.info, self.storage, self.placement]
    }

    const fn last(self) -> u64 {
        self.placement
    }
}

fn next_caption_identifier(
    package: &Package,
    metadata: &CaptionMetadataFacts,
) -> Result<u64, ChartCaptionError> {
    let metadata_last_identifier = metadata.last_identifier;
    let mut maximum = metadata_last_identifier;
    for component in package.state.source.components().iter() {
        for object in &component.archive().objects {
            maximum = maximum.max(
                object
                    .archive_info
                    .identifier
                    .ok_or(ChartCaptionError::InvalidSource)?,
            );
        }
    }
    for identifier in &metadata.owned_object_identifiers {
        maximum = maximum.max(*identifier);
    }
    maximum
        .checked_add(1)
        .ok_or(ChartCaptionError::InvalidSource)
}

fn fresh_uuid(existing: &HashSet<(u64, u64)>, identifier: u64) -> UuidBits {
    let mut lower = identifier.max(1);
    let mut upper = identifier.rotate_left(17) ^ 0x9e37_79b9_7f4a_7c15;
    while existing.contains(&(lower, upper)) || (lower == 0 && upper == 0) {
        lower = lower.wrapping_add(1).max(1);
        upper = upper.rotate_left(7) ^ 0xd1b5_4a32_d192_ed03;
    }
    UuidBits::new(lower, upper)
}

fn canonical_standin(identifier: u64) -> Result<ArchiveObject, ChartCaptionError> {
    let mut object = ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_: STANDIN_MESSAGE_TYPE,
            data: graph_codec::canonical_standin_payload().to_vec(),
        }],
    )
    .map_err(map_core_error)?;
    object.archive_info.message_infos[0].versions = vec![10, 1, 0];
    Ok(object)
}

fn caption_graph_objects(
    package: &Package,
    ids: CaptionGraphIds,
    drawable_identifier: u64,
    drawable_width: f32,
    text: &str,
    stylesheet_identifier: u64,
    paragraph_style_identifier: u64,
    language: Option<&str>,
    edge_kind: CaptionEdgeKind,
    budget: &mut CaptionBudget,
) -> Result<(Vec<ArchiveObject>, graph_codec::EncodeReport), ChartCaptionError> {
    let limits = package.wire_limits().map_err(map_wire_error)?;
    let wire = graph_codec::EncodeOptions::for_text(text)
        .with_max_output_bytes(budget.remaining_output().min(limits.max_output_bytes()))
        .with_max_text_bytes(budget.remaining_input().min(limits.max_input_bytes()))
        .with_max_fields(budget.remaining_fields().min(limits.max_fields()))
        .with_max_work_bytes(budget.remaining_work().min(limits.max_rewrite_work()))
        .with_max_depth(
            budget
                .remaining_depth()
                .min(u32::try_from(limits.max_nesting()).unwrap_or(u32::MAX)),
        )
        .with_max_allocations(budget.remaining_allocations());
    let graph_kind = match edge_kind {
        CaptionEdgeKind::MovieTitle => graph_codec::CaptionGraphKind::Title,
        CaptionEdgeKind::Chart | CaptionEdgeKind::Movie => graph_codec::CaptionGraphKind::Caption,
    };
    let output = graph_codec::encode_caption_graph_with_kind_with_report(
        graph_codec::CaptionGraphWrite {
            drawable_identifier,
            style_identifier: ids.style,
            info_identifier: ids.info,
            storage_identifier: ids.storage,
            placement_identifier: ids.placement,
            stylesheet_identifier,
            paragraph_style_identifier,
            drawable_width,
            text,
            language,
        },
        graph_kind,
        wire,
    )
    .map_err(map_graph_error)?;
    let report = output.report();
    let payloads = output.into_payloads().into_parts();
    let types = [
        SHAPE_STYLE_MESSAGE_TYPE,
        CAPTION_INFO_MESSAGE_TYPE,
        STORAGE_MESSAGE_TYPE,
        CAPTION_PLACEMENT_MESSAGE_TYPE,
    ];
    let identifiers = [ids.style, ids.info, ids.storage, ids.placement];
    let references = [
        vec![paragraph_style_identifier],
        vec![ids.style, ids.storage, ids.placement],
        vec![paragraph_style_identifier],
        Vec::new(),
    ];
    budget.charge_allocation(1, 0)?;
    let mut objects = Vec::new();
    objects
        .try_reserve_exact(4)
        .map_err(|_error| ChartCaptionError::Allocation { amount: 4 })?;
    for (index, (payload, references)) in payloads.into_iter().zip(references).enumerate() {
        let mut object = ArchiveObject::new(
            identifiers[index],
            vec![RawMessage {
                type_: types[index],
                data: payload,
            }],
        )
        .map_err(map_core_error)?;
        object.archive_info.message_infos[0].versions = vec![1, 0, 5];
        object.archive_info.message_infos[0].object_references = references;
        objects.push(object);
    }
    Ok((objects, report))
}

#[derive(Debug, Clone)]
struct CaptionTheme {
    stylesheet: u64,
    paragraph_style: u64,
    language: Option<String>,
}

fn caption_theme(package: &Package) -> Result<CaptionTheme, ChartCaptionError> {
    let limits = package.wire_limits().map_err(map_wire_error)?;
    let document_candidates = package
        .state
        .source
        .components()
        .iter()
        .flat_map(|component| component.archive().objects.iter())
        .filter(|object| {
            object
                .messages
                .iter()
                .any(|message| message.type_ == DOCUMENT_MESSAGE_TYPE)
        })
        .count();
    if document_candidates != 1 {
        return Err(ChartCaptionError::InvalidSource);
    }
    let document = package.object(1).ok_or(ChartCaptionError::InvalidSource)?;
    if document.messages.len() != 1 || document.messages[0].type_ != DOCUMENT_MESSAGE_TYPE {
        return Err(ChartCaptionError::InvalidSource);
    }
    let document_payload = document
        .messages
        .first()
        .ok_or(ChartCaptionError::InvalidSource)?
        .data
        .as_slice();
    let document_view =
        WireView::parse_with_limits(document_payload, limits).map_err(map_wire_error)?;
    let show_payload =
        unique_payload_field(&document_view, 2)?.ok_or(ChartCaptionError::InvalidSource)?;
    let show_identifier = reference_identifier(show_payload, limits)?;
    let show = package
        .object(show_identifier)
        .ok_or(ChartCaptionError::InvalidSource)?;
    if show.messages.len() != 1 || show.messages[0].type_ != SHOW_MESSAGE_TYPE {
        return Err(ChartCaptionError::InvalidSource);
    }
    let show_payload = show
        .messages
        .first()
        .ok_or(ChartCaptionError::InvalidSource)?
        .data
        .as_slice();
    let show_view = WireView::parse_with_limits(show_payload, limits).map_err(map_wire_error)?;
    let theme_identifier = reference_identifier(
        unique_payload_field(&show_view, 2)?.ok_or(ChartCaptionError::InvalidSource)?,
        limits,
    )?;
    let stylesheet_identifier = reference_identifier(
        unique_payload_field(&show_view, 5)?.ok_or(ChartCaptionError::InvalidSource)?,
        limits,
    )?;
    let theme = package
        .object(theme_identifier)
        .ok_or(ChartCaptionError::InvalidSource)?;
    if theme.messages.len() != 1 || theme.messages[0].type_ != THEME_MESSAGE_TYPE {
        return Err(ChartCaptionError::InvalidSource);
    }
    let theme_payload = theme
        .messages
        .first()
        .ok_or(ChartCaptionError::InvalidSource)?
        .data
        .as_slice();
    let theme_view = WireView::parse_with_limits(theme_payload, limits).map_err(map_wire_error)?;
    let theme_super =
        unique_payload_field(&theme_view, 1)?.ok_or(ChartCaptionError::InvalidSource)?;
    let theme_super_view =
        WireView::parse_with_limits(theme_super, limits).map_err(map_wire_error)?;
    let presets =
        unique_payload_field(&theme_super_view, 210)?.ok_or(ChartCaptionError::InvalidSource)?;
    let presets_view = WireView::parse_with_limits(presets, limits).map_err(map_wire_error)?;
    let paragraph_style = first_reference_identifier(&presets_view, 1, limits)?;
    let document_super =
        unique_payload_field(&document_view, 3)?.ok_or(ChartCaptionError::InvalidSource)?;
    let document_super_view =
        WireView::parse_with_limits(document_super, limits).map_err(map_wire_error)?;
    let language = match unique_payload_field(&document_super_view, 3)? {
        Some(bytes) => Some(
            std::str::from_utf8(bytes)
                .map_err(|_| ChartCaptionError::InvalidSource)?
                .to_owned(),
        ),
        None => None,
    };
    Ok(CaptionTheme {
        stylesheet: stylesheet_identifier,
        paragraph_style,
        language,
    })
}

fn reference_identifier(payload: &[u8], limits: WireLimits) -> Result<u64, ChartCaptionError> {
    let view = WireView::parse_with_limits(payload, limits).map_err(map_wire_error)?;
    let mut selected = None;
    for field in view.fields().filter(|field| field.number() == 1) {
        if selected.is_some() || field.wire_type() != 0 {
            return Err(ChartCaptionError::InvalidSource);
        }
        field.validate_canonical_key().map_err(map_wire_error)?;
        selected = Some(field);
    }
    let field = selected.ok_or(ChartCaptionError::InvalidSource)?;
    let (value, bytes) =
        decode_varint_from_bytes(field.payload()).map_err(|_| ChartCaptionError::InvalidSource)?;
    if bytes != field.payload().len() || bytes != encoded_len(value) || value == 0 {
        return Err(ChartCaptionError::InvalidSource);
    }
    validate_reference_optional_fields(&view)?;
    Ok(value)
}

fn first_reference_identifier(
    view: &WireView<'_>,
    number: u32,
    limits: WireLimits,
) -> Result<u64, ChartCaptionError> {
    let mut first = None;
    for field in view.fields().filter(|field| field.number() == number) {
        if field.wire_type() != 2 {
            return Err(ChartCaptionError::InvalidSource);
        }
        field.validate_canonical_key().map_err(map_wire_error)?;
        field.validate_canonical_framing().map_err(map_wire_error)?;
        let identifier = reference_identifier(field.payload(), limits)?;
        if first.is_none() {
            first = Some(identifier);
        }
    }
    first.ok_or(ChartCaptionError::InvalidSource)
}

fn caption_drawable_width(
    package: &Package,
    drawable_identifier: u64,
) -> Result<f32, ChartCaptionError> {
    let limits = package.wire_limits().map_err(map_wire_error)?;
    let object = package
        .object(drawable_identifier)
        .ok_or(ChartCaptionError::InvalidSource)?;
    let payload = &object
        .messages
        .first()
        .ok_or(ChartCaptionError::InvalidSource)?
        .data;
    let outer = WireView::parse_with_limits(payload, limits).map_err(map_wire_error)?;
    let drawable = unique_payload_field(&outer, 1)?.ok_or(ChartCaptionError::InvalidSource)?;
    let drawable_view = WireView::parse_with_limits(drawable, limits).map_err(map_wire_error)?;
    let geometry =
        unique_payload_field(&drawable_view, 1)?.ok_or(ChartCaptionError::InvalidSource)?;
    let geometry_view = WireView::parse_with_limits(geometry, limits).map_err(map_wire_error)?;
    let size = unique_payload_field(&geometry_view, 2)?.ok_or(ChartCaptionError::InvalidSource)?;
    let size_view = WireView::parse_with_limits(size, limits).map_err(map_wire_error)?;
    let mut width = None;
    for field in size_view.fields().filter(|field| field.number() == 1) {
        if width.is_some() || field.wire_type() != 5 {
            return Err(ChartCaptionError::InvalidSource);
        }
        field.validate_canonical_key().map_err(map_wire_error)?;
        width = Some(field);
    }
    let width = width.ok_or(ChartCaptionError::InvalidSource)?;
    if width.payload().len() != 4 {
        return Err(ChartCaptionError::InvalidSource);
    }
    let value = f32::from_le_bytes(width.payload().try_into().unwrap_or([0; 4]));
    if !value.is_finite() || value <= 0.0 {
        return Err(ChartCaptionError::InvalidSource);
    }
    Ok(value)
}

fn patch_caption_edge(
    archive: &mut Archive,
    drawable_identifier: u64,
    expected_identifier: Option<u64>,
    replacement_identifier: u64,
    package: &Package,
    archive_limits: litchi_iwa_core::Limits,
    budget: &mut CaptionBudget,
    edge_kind: CaptionEdgeKind,
    graph_style_identifier: Option<u64>,
) -> Result<(), ChartCaptionError> {
    let missing_movie_edge = expected_identifier.is_none()
        && matches!(
            edge_kind,
            CaptionEdgeKind::Movie | CaptionEdgeKind::MovieTitle
        );
    let mut expected_identifier = expected_identifier;
    if replacement_identifier == 0 {
        return Err(ChartCaptionError::InvalidSource);
    }
    let source_object = archive
        .object(drawable_identifier)
        .ok_or(ChartCaptionError::InvalidSource)?;
    let message_type = match edge_kind {
        CaptionEdgeKind::Chart => CHART_MESSAGE_TYPE,
        CaptionEdgeKind::Movie | CaptionEdgeKind::MovieTitle => 3_007,
    };
    if source_object.messages.len() != 1 || source_object.messages[0].type_ != message_type {
        return Err(ChartCaptionError::InvalidSource);
    }
    validate_selected_message_metadata(source_object, 0)?;
    let info = source_object
        .archive_info
        .message_infos
        .first()
        .ok_or(ChartCaptionError::InvalidSource)?;
    if expected_identifier.is_none() {
        if !missing_movie_edge {
            return Err(ChartCaptionError::InvalidSource);
        }
        let standins = info
            .object_references
            .iter()
            .copied()
            .filter(|identifier| {
                package.object(*identifier).is_some_and(|object| {
                    object.messages.len() == 1
                        && object.messages[0].type_ == STANDIN_MESSAGE_TYPE
                        && object.messages[0].data.is_empty()
                })
            })
            .collect::<Vec<_>>();
        if standins.len() != 1 {
            return Err(ChartCaptionError::InvalidSource);
        }
        expected_identifier = standins.first().copied();
    }
    let expected_identifier = expected_identifier.ok_or(ChartCaptionError::InvalidSource)?;
    if expected_identifier == replacement_identifier {
        return Err(ChartCaptionError::InvalidSource);
    }
    if info.data_references.contains(&expected_identifier)
        || info.data_references.contains(&replacement_identifier)
        || info
            .object_references
            .iter()
            .filter(|identifier| **identifier == expected_identifier)
            .count()
            != 1
        || info.object_references.contains(&replacement_identifier)
    {
        return Err(ChartCaptionError::UnsupportedDependency);
    }
    struct OwnedFieldTransition {
        path: Vec<u32>,
        before: Vec<u64>,
        after: Vec<u64>,
    }
    let mut field_states = Vec::new();
    budget.charge_allocation(1, 0)?;
    field_states
        .try_reserve_exact(info.field_infos.len())
        .map_err(|_error| ChartCaptionError::Allocation {
            amount: info.field_infos.len(),
        })?;
    for field in &info.field_infos {
        let count = field
            .object_references
            .iter()
            .filter(|identifier| **identifier == expected_identifier)
            .count();
        if field.data_references.contains(&expected_identifier)
            || field.data_references.contains(&replacement_identifier)
            || field.object_references.contains(&replacement_identifier)
        {
            return Err(ChartCaptionError::UnsupportedDependency);
        }
        let edge_field = match edge_kind {
            CaptionEdgeKind::Chart => 11,
            CaptionEdgeKind::Movie | CaptionEdgeKind::MovieTitle => {
                if edge_kind == CaptionEdgeKind::MovieTitle {
                    10
                } else {
                    11
                }
            },
        };
        if count != 0
            && (!matches!(field.path.path.as_slice(), [number, 1] if *number == edge_field)
                && !matches!(field.path.path.as_slice(), [1, number, 1] if *number == edge_field)
                && !matches!(field.path.path.as_slice(), [1, 1, number, 1] if *number == edge_field)
                || count != 1)
        {
            return Err(ChartCaptionError::UnsupportedDependency);
        }
        let mut after = field.object_references.clone();
        after.retain(|identifier| *identifier != expected_identifier);
        if count != 0 {
            after.push(replacement_identifier);
        }
        field_states.push(OwnedFieldTransition {
            path: field.path.path.clone(),
            before: field.object_references.clone(),
            after,
        });
    }
    let payload = source_object.messages[0].data.clone();
    let (
        rewritten,
        report_input,
        report_output,
        report_fields,
        report_work,
        report_depth,
        report_allocations,
        report_retained,
        report_scratch,
    ) = match edge_kind {
        CaptionEdgeKind::Chart => {
            let options = chart_caption_rewrite_options(package, &payload, budget)?;
            let (rewritten, report) =
                keynote_chart_caption_codec::rewrite_chart_caption_with_report(
                    &payload,
                    keynote_chart_caption_codec::ChartCaptionWrite::new(replacement_identifier),
                    options,
                )
                .map_err(map_chart_caption_codec_error)?;
            (
                rewritten,
                report.input_bytes(),
                report.output_bytes(),
                report.fields(),
                report.work_bytes(),
                report.max_depth(),
                report.allocations(),
                report.retained_bytes(),
                report.scratch_bytes(),
            )
        },
        CaptionEdgeKind::Movie | CaptionEdgeKind::MovieTitle if missing_movie_edge => {
            let limits = package.wire_limits().map_err(map_wire_error)?;
            let rewritten = insert_movie_caption_edge(
                &payload,
                replacement_identifier,
                limits,
                edge_kind,
                budget,
            )?;
            let input_bytes = payload.len();
            let output_bytes = rewritten.len();
            let source_view =
                WireView::parse_with_limits(&payload, limits).map_err(map_wire_error)?;
            let output_view =
                WireView::parse_with_limits(&rewritten, limits).map_err(map_wire_error)?;
            let source_fields = source_view.fields().count()
                + WireView::parse_with_limits(
                    unique_payload_field(&source_view, 1)?
                        .ok_or(ChartCaptionError::InvalidSource)?,
                    limits,
                )
                .map_err(map_wire_error)?
                .fields()
                .count();
            let output_fields = output_view.fields().count()
                + WireView::parse_with_limits(
                    unique_payload_field(&output_view, 1)?
                        .ok_or(ChartCaptionError::InvalidSource)?,
                    limits,
                )
                .map_err(map_wire_error)?
                .fields()
                .count();
            (
                rewritten,
                input_bytes,
                output_bytes,
                source_fields
                    .saturating_add(output_fields)
                    .saturating_add(2),
                input_bytes.saturating_add(output_bytes),
                4,
                1,
                output_bytes,
                output_bytes,
            )
        },
        CaptionEdgeKind::Movie | CaptionEdgeKind::MovieTitle => {
            let options = movie_caption_rewrite_options(package, &payload, budget)?;
            let write = match edge_kind {
                CaptionEdgeKind::Movie => {
                    keynote_movie_caption_codec::MovieCaptionWrite::caption(replacement_identifier)
                },
                CaptionEdgeKind::MovieTitle => {
                    keynote_movie_caption_codec::MovieCaptionWrite::title(replacement_identifier)
                },
                CaptionEdgeKind::Chart => unreachable!(),
            };
            let (rewritten, report) =
                keynote_movie_caption_codec::rewrite_movie_caption_with_report(
                    &payload, write, options,
                )
                .map_err(map_movie_caption_codec_error)?;
            (
                rewritten,
                report.input_bytes(),
                report.output_bytes(),
                report.fields(),
                report.work_bytes(),
                report.max_depth(),
                report.allocations(),
                report.retained_bytes(),
                report.scratch_bytes(),
            )
        },
    };
    budget.charge_input(report_input)?;
    budget.charge_output(report_output)?;
    budget.charge_fields(report_fields)?;
    budget.charge_work(report_work)?;
    budget.charge_depth(report_depth)?;
    budget.charge_allocation(report_allocations, report_retained)?;
    budget.charge_scratch(report_scratch)?;
    let aggregate_before = info.object_references.clone();
    let mut aggregate_after = aggregate_before.clone();
    aggregate_after.retain(|identifier| *identifier != expected_identifier);
    aggregate_after.push(replacement_identifier);
    if let Some(style_identifier) = graph_style_identifier {
        if style_identifier == 0 || aggregate_before.contains(&style_identifier) {
            return Err(ChartCaptionError::UnsupportedDependency);
        }
        aggregate_after.push(style_identifier);
    }
    let fields = field_states
        .iter()
        .enumerate()
        .map(|(field_info_index, field)| FieldObjectReferenceTransition {
            field_info_index,
            expected_path: field.path.as_slice(),
            before: field.before.as_slice(),
            after: field.after.as_slice(),
        })
        .collect::<Vec<_>>();
    let transition = ObjectReferenceTransition {
        aggregate_before: aggregate_before.as_slice(),
        aggregate_after: aggregate_after.as_slice(),
        fields: fields.as_slice(),
    };
    let mut rewritten_object = source_object.clone();
    rewritten_object
        .replace_message_transitioning_object_references_preserving_header_with_limits(
            0,
            RawMessage {
                type_: message_type,
                data: rewritten,
            },
            transition,
            archive_limits,
        )
        .map_err(map_core_error)?;
    *archive
        .object_mut(drawable_identifier)
        .ok_or(ChartCaptionError::InvalidSource)? = rewritten_object;
    Ok(())
}

/// Add a missing MovieArchive caption edge to a canonical stand-in drawable.
///
/// The strict movie codec intentionally rewrites an existing reference only.
/// A stand-in fixture can carry its placeholder solely in ArchiveInfo's
/// aggregate references, so creation first authors the one canonical field
/// while retaining every source wire field and then strictly decodes the
/// candidate through the same movie codec.
fn insert_movie_caption_edge(
    payload: &[u8],
    replacement_identifier: u64,
    limits: WireLimits,
    edge_kind: CaptionEdgeKind,
    budget: &mut CaptionBudget,
) -> Result<Vec<u8>, ChartCaptionError> {
    if replacement_identifier == 0 {
        return Err(ChartCaptionError::InvalidSource);
    }
    let outer = WireView::parse_with_limits(payload, limits).map_err(map_wire_error)?;
    let super_payload = unique_payload_field(&outer, 1)?.ok_or(ChartCaptionError::InvalidSource)?;
    let drawable = WireView::parse_with_limits(super_payload, limits).map_err(map_wire_error)?;
    let edge_field = match edge_kind {
        CaptionEdgeKind::Movie => 11,
        CaptionEdgeKind::MovieTitle => 10,
        CaptionEdgeKind::Chart => return Err(ChartCaptionError::InvalidSource),
    };
    if drawable.fields().any(|field| field.number() == edge_field) {
        return Err(ChartCaptionError::InvalidSource);
    }

    budget.charge_allocation(1, 0)?;
    let mut reference_payload = Vec::new();
    append_varint_field(&mut reference_payload, 1, replacement_identifier)
        .map_err(map_wire_error)?;
    budget.charge_allocation(1, 0)?;
    let mut drawable_output = Vec::new();
    drawable_output
        .try_reserve_exact(
            super_payload
                .len()
                .saturating_add(reference_payload.len() + 4),
        )
        .map_err(|_| ChartCaptionError::Allocation {
            amount: super_payload
                .len()
                .saturating_add(reference_payload.len() + 4),
        })?;
    for field in drawable.fields() {
        field.validate_canonical_key().map_err(map_wire_error)?;
        field.validate_canonical_framing().map_err(map_wire_error)?;
        drawable_output.extend_from_slice(field.raw());
    }
    append_length_delimited_field_with_limits(
        &mut drawable_output,
        edge_field,
        &reference_payload,
        limits,
    )
    .map_err(map_wire_error)?;

    budget.charge_allocation(1, 0)?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(payload.len().saturating_add(drawable_output.len()))
        .map_err(|_| ChartCaptionError::Allocation {
            amount: payload.len().saturating_add(drawable_output.len()),
        })?;
    for field in outer.fields() {
        field.validate_canonical_key().map_err(map_wire_error)?;
        field.validate_canonical_framing().map_err(map_wire_error)?;
        if field.number() == 1 {
            append_length_delimited_field_with_limits(&mut output, 1, &drawable_output, limits)
                .map_err(map_wire_error)?;
        } else {
            output.extend_from_slice(field.raw());
        }
    }
    let recursion =
        u32::try_from(limits.max_nesting()).map_err(|_error| ChartCaptionError::InvalidSource)?;
    let snapshot = keynote_movie_caption_codec::decode_movie_caption(
        &output,
        keynote_movie_caption_codec::DecodeOptions::new(
            output.len().min(limits.max_input_bytes()),
            limits.max_fields(),
            limits.max_rewrite_work(),
            recursion,
        )
        .with_max_output_bytes(limits.max_output_bytes()),
    )
    .map_err(map_movie_caption_codec_error)?;
    let actual = match edge_kind {
        CaptionEdgeKind::Movie => snapshot.caption_identifier(),
        CaptionEdgeKind::MovieTitle => snapshot.title_identifier(),
        CaptionEdgeKind::Chart => None,
    };
    if actual != Some(replacement_identifier) {
        return Err(ChartCaptionError::Verification);
    }
    Ok(output)
}

fn verify_graph_transition(
    source: &Package,
    candidate: &Package,
    before: &CaptionSelection,
    target: &CaptionSelection,
) -> Result<(), ChartCaptionError> {
    verify_caption_graph_transition(
        source,
        candidate,
        &before.slide_component_name,
        before.reference_identifier,
        before.caption_info_identifier,
        before.storage_identifier,
        before.placement_identifier,
        before.style_identifier,
        target.reference_identifier,
        target.caption_info_identifier,
        target.storage_identifier,
        target.placement_identifier,
        target.style_identifier,
    )
}

/// Verify the physical locality and retained-object rules of a canonical
/// drawable graph transition. Movie and chart selectors provide the same
/// graph facts through this private primitive.
#[allow(
    clippy::too_many_arguments,
    reason = "The private cross-owner seam carries the complete before/target graph census explicitly."
)]
pub(super) fn verify_caption_graph_transition(
    source: &Package,
    candidate: &Package,
    slide_component_name: &str,
    before_reference: Option<u64>,
    before_caption_info: Option<u64>,
    before_storage: Option<u64>,
    before_placement: Option<u64>,
    before_style: Option<u64>,
    target_reference: Option<u64>,
    target_caption_info: Option<u64>,
    target_storage: Option<u64>,
    target_placement: Option<u64>,
    target_style: Option<u64>,
) -> Result<(), ChartCaptionError> {
    let source_catalog = physical_catalog(source)?;
    let candidate_catalog = physical_catalog(candidate)?;
    let metadata_name = metadata_member_name(source_catalog, source)?;
    if metadata_member_name(candidate_catalog, candidate)? != metadata_name {
        return Err(ChartCaptionError::Verification);
    }
    let slide_name = slide_component_name;
    for entry in source_catalog.package().iter() {
        if PREVIEW_ENTRY_NAMES.contains(&entry.name())
            || entry.name() == slide_name
            || entry.name() == metadata_name
        {
            continue;
        }
        let candidate_entry = candidate_catalog
            .package()
            .iter()
            .find(|candidate_entry| candidate_entry.name() == entry.name())
            .ok_or(ChartCaptionError::Verification)?;
        if entry.data() != candidate_entry.data() {
            return Err(ChartCaptionError::Verification);
        }
    }
    let source_metadata = metadata_payload(&archive_for_member(source, &metadata_name)?)?;
    let candidate_metadata = metadata_payload(&archive_for_member(candidate, &metadata_name)?)?;
    if source_metadata.2 == candidate_metadata.2 {
        return Err(ChartCaptionError::Verification);
    }
    if before_storage.is_none() {
        // A stand-in -> active transition normally retains the source
        // stand-in.  The exact inverse of an active -> stand-in transition
        // has the same semantic orientation, but its source stand-in is the
        // newly allocated object and therefore is deliberately absent from
        // the restored candidate.  In both cases the target graph must be
        // present in full.
        if target_storage.is_none()
            || target_style.is_none()
            || target_placement.is_none()
            || target_caption_info.is_none()
            || ![
                target_caption_info,
                target_storage,
                target_placement,
                target_style,
            ]
            .into_iter()
            .flatten()
            .all(|identifier| candidate.object(identifier).is_some())
        {
            return Err(ChartCaptionError::Verification);
        }
        let source_standin_present =
            before_reference.is_some_and(|identifier| candidate.object(identifier).is_some());
        if !source_standin_present
            && ![
                target_caption_info,
                target_storage,
                target_placement,
                target_style,
            ]
            .into_iter()
            .flatten()
            .all(|identifier| candidate.object(identifier).is_some())
        {
            return Err(ChartCaptionError::Verification);
        }
    } else if target_storage.is_some() {
        return Err(ChartCaptionError::Verification);
    } else {
        // Removal creates a fresh stand-in while retaining the old graph.
        // The inverse of creation restores the original stand-in and drops
        // that old graph, so either the target stand-in or the source graph
        // is the retained side of this exact transition.
        let target_standin = candidate
            .object(target_reference.ok_or(ChartCaptionError::Verification)?)
            .is_some();
        if !target_standin {
            return Err(ChartCaptionError::Verification);
        }
        let source_graph_present = [
            before_caption_info,
            before_storage,
            before_placement,
            before_style,
        ]
        .into_iter()
        .flatten()
        .all(|identifier| candidate.object(identifier).is_some());
        if !source_graph_present && target_reference.is_none() {
            return Err(ChartCaptionError::Verification);
        }
    }
    Ok(())
}

fn archive_for_member(package: &Package, name: &str) -> Result<Archive, ChartCaptionError> {
    let catalog = physical_catalog(package)?;
    let entry = catalog
        .package()
        .iter()
        .find(|entry| entry.name() == name)
        .ok_or(ChartCaptionError::Verification)?;
    let physical_limits = package.state.options.archive();
    let stream = SnappyStream::decompress_with_limits(
        entry.data(),
        physical_limits.snappy_limits().map_err(map_archive_error)?,
    )
    .map_err(map_core_error)?;
    let archive = Archive::parse_with_limits(
        stream.as_bytes(),
        physical_limits
            .effective_archive_limits()
            .map_err(map_archive_error)?,
    )
    .map_err(map_core_error)?;
    validate_canonical_object_framing(&archive, stream.as_bytes())?;
    Ok(archive)
}

fn validate_canonical_object_framing(
    archive: &Archive,
    source: &[u8],
) -> Result<(), ChartCaptionError> {
    archive
        .validate_canonical_object_framing(source)
        .map_err(map_core_error)
}

fn verify_caption_candidate(
    source: &Package,
    candidate: &Package,
    before: &CaptionSelection,
    target: &CaptionSelection,
    expected: Option<&str>,
    require_invalidated_previews: bool,
    budget: &mut CaptionBudget,
) -> Result<(), ChartCaptionError> {
    let source_count = source.state.total_objects;
    let candidate_count = candidate.state.total_objects;
    let count_matches = match (before.storage_identifier, target.storage_identifier) {
        // Forward creation adds four graph objects; an exact inverse of a
        // removal restores those four objects from a source that only has its
        // fresh stand-in.
        (None, Some(_)) => {
            candidate_count == source_count.saturating_add(4)
                || source_count == candidate_count.saturating_add(1)
        },
        // Forward removal adds one stand-in; an exact inverse of creation
        // removes the four graph objects while restoring the old stand-in.
        (Some(_), None) => {
            candidate_count == source_count.saturating_add(1)
                || source_count == candidate_count.saturating_add(4)
        },
        _ => candidate_count == source_count,
    };
    if !count_matches {
        return Err(ChartCaptionError::Verification);
    }
    budget.charge_references(1)?;
    let selected = select_caption(
        candidate,
        SlideSelector::position(before.slide_position),
        ChartSelector::index(before.chart_position.get()),
        true,
    )?;
    if !selected.same_identity(target) || selected.text.as_deref() != expected {
        return Err(ChartCaptionError::Verification);
    }
    if let (Some(storage_identifier), Some(target_storage)) =
        (before.storage_identifier, target.storage_identifier)
    {
        if storage_identifier != target_storage {
            return Err(ChartCaptionError::Verification);
        }
        if metadata_member_name(physical_catalog(source)?, source).is_ok() {
            verify_existing_text_metadata_candidate(
                source,
                candidate,
                storage_identifier,
                before.slide_node_identifier,
                require_invalidated_previews,
            )?;
        } else {
            super::slide_text::verify_owned_storage_candidate(
                source,
                candidate,
                storage_identifier,
                before.slide_node_identifier,
                require_invalidated_previews,
            )
            .map_err(map_slide_text_error)?;
        }
    } else {
        verify_graph_transition(source, candidate, before, target)?;
        if require_invalidated_previews
            && !super::rendering_invalidation::root_previews_absent(
                candidate.state.source.package(),
            )
            .map_err(map_rendering_error)?
        {
            return Err(ChartCaptionError::Verification);
        }
    }
    Ok(())
}

pub(super) fn verify_existing_text_metadata_candidate(
    source: &Package,
    candidate: &Package,
    storage_identifier: u64,
    slide_node_identifier: u64,
    require_invalidated_previews: bool,
) -> Result<(), ChartCaptionError> {
    let source_catalog = physical_catalog(source)?;
    let candidate_catalog = physical_catalog(candidate)?;
    let metadata_name = metadata_member_name(source_catalog, source)?;
    if metadata_member_name(candidate_catalog, candidate)? != metadata_name
        || source_catalog.components().len() != candidate_catalog.components().len()
    {
        return Err(ChartCaptionError::Verification);
    }
    let source_metadata = metadata_payload(&archive_for_member(source, &metadata_name)?)?;
    let candidate_metadata = metadata_payload(&archive_for_member(candidate, &metadata_name)?)?;
    if source_metadata.2 == candidate_metadata.2 {
        return Err(ChartCaptionError::Verification);
    }
    let mut storage_seen = false;
    let mut slide_node_seen = false;
    let mut metadata_seen = false;
    for (source_component, candidate_component) in source_catalog
        .components()
        .iter()
        .zip(candidate_catalog.components().iter())
    {
        if source_component.name() != candidate_component.name()
            || source_component.archive().objects.len()
                != candidate_component.archive().objects.len()
        {
            return Err(ChartCaptionError::Verification);
        }
        for (source_object, candidate_object) in source_component
            .archive()
            .objects
            .iter()
            .zip(&candidate_component.archive().objects)
        {
            let identifier = source_object
                .archive_info
                .identifier
                .ok_or(ChartCaptionError::Verification)?;
            if candidate_object.archive_info.identifier != Some(identifier) {
                return Err(ChartCaptionError::Verification);
            }
            if identifier == storage_identifier {
                if std::mem::replace(&mut storage_seen, true) {
                    return Err(ChartCaptionError::Verification);
                }
                verify_replaced_caption_object(source_object, candidate_object, true)?;
            } else if identifier == slide_node_identifier {
                if std::mem::replace(&mut slide_node_seen, true) {
                    return Err(ChartCaptionError::Verification);
                }
                if !source_object.same_content_ignoring_offsets(candidate_object) {
                    verify_replaced_caption_object(source_object, candidate_object, true)?;
                }
                if require_invalidated_previews
                    && !super::slide_preview::is_invalidated(
                        candidate_object,
                        candidate.wire_limits().map_err(map_wire_error)?,
                    )
                    .map_err(map_slide_preview_error)?
                {
                    return Err(ChartCaptionError::Verification);
                }
            } else if identifier == source_metadata.0 {
                if std::mem::replace(&mut metadata_seen, true) {
                    return Err(ChartCaptionError::Verification);
                }
                verify_replaced_caption_object(source_object, candidate_object, true)?;
            } else if !source_object.same_content_ignoring_offsets(candidate_object) {
                return Err(ChartCaptionError::Verification);
            }
        }
    }
    if !storage_seen || !slide_node_seen || !metadata_seen {
        return Err(ChartCaptionError::Verification);
    }
    if require_invalidated_previews
        && PREVIEW_ENTRY_NAMES.iter().any(|name| {
            candidate_catalog
                .package()
                .iter()
                .any(|entry| entry.name() == *name)
        })
    {
        return Err(ChartCaptionError::Verification);
    }
    Ok(())
}

fn verify_replaced_caption_object(
    source: &ArchiveObject,
    candidate: &ArchiveObject,
    require_changed: bool,
) -> Result<(), ChartCaptionError> {
    if source.archive_info.identifier != candidate.archive_info.identifier
        || source.archive_info.should_merge != candidate.archive_info.should_merge
        || source.messages.len() != candidate.messages.len()
        || source.archive_info.message_infos.len() != candidate.archive_info.message_infos.len()
    {
        return Err(ChartCaptionError::Verification);
    }
    let mut changed = 0usize;
    for ((source_message, candidate_message), (source_info, candidate_info)) in
        source.messages.iter().zip(&candidate.messages).zip(
            source
                .archive_info
                .message_infos
                .iter()
                .zip(&candidate.archive_info.message_infos),
        )
    {
        if source_message.type_ != candidate_message.type_
            || !message_info_equal_except_length(source_info, candidate_info)
        {
            return Err(ChartCaptionError::Verification);
        }
        if source_message.data != candidate_message.data {
            changed = changed
                .checked_add(1)
                .ok_or(ChartCaptionError::Verification)?;
        } else if source_info.length != candidate_info.length {
            return Err(ChartCaptionError::Verification);
        }
    }
    if require_changed && changed != 1 {
        return Err(ChartCaptionError::Verification);
    }
    Ok(())
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

fn chart_caption_decode_options(
    package: &Package,
    payload: &[u8],
) -> Result<keynote_chart_caption_codec::DecodeOptions, ChartCaptionError> {
    let limits = package.wire_limits().map_err(map_wire_error)?;
    let recursion =
        u32::try_from(limits.max_nesting()).map_err(|_error| ChartCaptionError::InvalidSource)?;
    Ok(keynote_chart_caption_codec::DecodeOptions::new(
        payload.len().min(limits.max_input_bytes()),
        limits.max_fields(),
        limits.max_rewrite_work(),
        recursion,
    )
    // The selected identifier may legitimately grow from a one-byte varint
    // to a wider varint.  A source-sized default therefore rejects valid
    // source-preserving rewrites before the codec can measure the candidate.
    // Keep the ceiling finite, but derive it from the package's output policy.
    .with_max_output_bytes(limits.max_output_bytes()))
}

fn chart_caption_rewrite_options(
    package: &Package,
    payload: &[u8],
    budget: &CaptionBudget,
) -> Result<keynote_chart_caption_codec::DecodeOptions, ChartCaptionError> {
    let limits = package.wire_limits().map_err(map_wire_error)?;
    let recursion = u32::try_from(limits.max_nesting())
        .map_err(|_error| ChartCaptionError::InvalidSource)?
        .min(budget.remaining_depth());
    Ok(keynote_chart_caption_codec::DecodeOptions::new(
        payload
            .len()
            .min(limits.max_input_bytes())
            .min(budget.remaining_input()),
        limits.max_fields().min(budget.remaining_fields()),
        limits.max_rewrite_work().min(budget.remaining_work()),
        recursion,
    )
    .with_max_output_bytes(limits.max_output_bytes().min(budget.remaining_output())))
}

fn movie_caption_rewrite_options(
    package: &Package,
    payload: &[u8],
    budget: &CaptionBudget,
) -> Result<keynote_movie_caption_codec::DecodeOptions, ChartCaptionError> {
    let limits = package.wire_limits().map_err(map_wire_error)?;
    let recursion = u32::try_from(limits.max_nesting())
        .map_err(|_error| ChartCaptionError::InvalidSource)?
        .min(budget.remaining_depth());
    Ok(keynote_movie_caption_codec::DecodeOptions::new(
        payload
            .len()
            .min(limits.max_input_bytes())
            .min(budget.remaining_input()),
        limits.max_fields().min(budget.remaining_fields()),
        limits.max_rewrite_work().min(budget.remaining_work()),
        recursion,
    )
    .with_max_output_bytes(limits.max_output_bytes().min(budget.remaining_output())))
}

pub(super) fn caption_info_decode_options(
    package: &Package,
    payload: &[u8],
) -> Result<pages_movie_caption_codec::DecodeOptions, ChartCaptionError> {
    let limits = package.wire_limits().map_err(map_wire_error)?;
    let recursion =
        u32::try_from(limits.max_nesting()).map_err(|_error| ChartCaptionError::InvalidSource)?;
    Ok(pages_movie_caption_codec::DecodeOptions::new(
        payload.len().min(limits.max_input_bytes()),
        limits.max_fields(),
        limits.max_rewrite_work(),
        recursion,
    ))
}

fn copy_caption(value: &str) -> Result<String, ChartCaptionError> {
    if value.len() > MAX_CAPTION_BYTES {
        return Err(ChartCaptionError::LimitExceeded {
            kind: ChartCaptionLimitKind::CaptionBytes,
            observed: usize_to_u64(value.len()),
            maximum: usize_to_u64(MAX_CAPTION_BYTES),
        });
    }
    let mut copy = String::new();
    copy.try_reserve_exact(value.len())
        .map_err(|_error| ChartCaptionError::Allocation {
            amount: value.len(),
        })?;
    copy.push_str(value);
    Ok(copy)
}

fn contains_dependent_marker(text: &str) -> bool {
    text.contains('\u{000e}') || text.contains('\u{fffc}')
}

fn preview_count(catalog: &SourceCatalog) -> usize {
    PREVIEW_ENTRY_NAMES
        .iter()
        .filter(|name| catalog.package().iter().any(|entry| entry.name() == **name))
        .count()
}

fn physical_catalog(package: &Package) -> Result<&SourceCatalog, ChartCaptionError> {
    match &package.state.source {
        PhysicalSource::Package(source) => Ok(source),
        PhysicalSource::Semantic(_) => Err(ChartCaptionError::UnsupportedSource),
    }
}

fn map_chart_title_error(error: super::slide_chart_title::ChartTitleError) -> ChartCaptionError {
    match error {
        super::slide_chart_title::ChartTitleError::UnsupportedSource => {
            ChartCaptionError::UnsupportedSource
        },
        super::slide_chart_title::ChartTitleError::AmbiguousSelector => {
            ChartCaptionError::AmbiguousSelector
        },
        super::slide_chart_title::ChartTitleError::EmptySlideName => {
            ChartCaptionError::EmptySlideName
        },
        super::slide_chart_title::ChartTitleError::SlideNameNotFound => {
            ChartCaptionError::SlideNameNotFound
        },
        super::slide_chart_title::ChartTitleError::SlidePositionNotFound { position } => {
            ChartCaptionError::SlidePositionNotFound { position }
        },
        super::slide_chart_title::ChartTitleError::ChartNameNotFound => {
            ChartCaptionError::ChartNameNotFound
        },
        super::slide_chart_title::ChartTitleError::ChartPositionNotFound { position } => {
            ChartCaptionError::ChartPositionNotFound { position }
        },
        super::slide_chart_title::ChartTitleError::EmptyChartName => {
            ChartCaptionError::EmptyChartName
        },
        super::slide_chart_title::ChartTitleError::LimitExceeded {
            observed, maximum, ..
        } => ChartCaptionError::LimitExceeded {
            kind: ChartCaptionLimitKind::WireWork,
            observed,
            maximum,
        },
        super::slide_chart_title::ChartTitleError::Allocation { amount } => {
            ChartCaptionError::Allocation { amount }
        },
        super::slide_chart_title::ChartTitleError::InvalidSource
        | super::slide_chart_title::ChartTitleError::Verification
        | super::slide_chart_title::ChartTitleError::PatchConflict => {
            ChartCaptionError::InvalidSource
        },
    }
}

fn map_archive_error(error: litchi_iwa_archive::Error) -> ChartCaptionError {
    match error {
        litchi_iwa_archive::Error::Limit {
            kind,
            observed,
            maximum,
        } => ChartCaptionError::LimitExceeded {
            kind: match kind {
                litchi_iwa_archive::LimitKind::InputBytes => ChartCaptionLimitKind::InputBytes,
                litchi_iwa_archive::LimitKind::OutputBytes => ChartCaptionLimitKind::OutputBytes,
                litchi_iwa_archive::LimitKind::Entries => ChartCaptionLimitKind::Entries,
                litchi_iwa_archive::LimitKind::EntryBytes
                | litchi_iwa_archive::LimitKind::CompressedEntryBytes => {
                    ChartCaptionLimitKind::EntryBytes
                },
                litchi_iwa_archive::LimitKind::TotalBytes
                | litchi_iwa_archive::LimitKind::IwaTotalBytes => ChartCaptionLimitKind::TotalBytes,
                litchi_iwa_archive::LimitKind::IwaStreamBytes => ChartCaptionLimitKind::WireBytes,
                litchi_iwa_archive::LimitKind::MemberNameBytes
                | litchi_iwa_archive::LimitKind::MetadataBytes => ChartCaptionLimitKind::EntryBytes,
            },
            observed,
            maximum,
        },
        litchi_iwa_archive::Error::Allocation { amount, .. } => {
            ChartCaptionError::Allocation { amount }
        },
        litchi_iwa_archive::Error::Iwa(error) => map_core_error(error),
        _ => ChartCaptionError::InvalidSource,
    }
}

fn map_core_error(error: litchi_iwa_core::Error) -> ChartCaptionError {
    match error {
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => ChartCaptionError::LimitExceeded {
            kind: match kind {
                litchi_iwa_core::LimitKind::ArchiveBytes
                | litchi_iwa_core::LimitKind::ObjectBytes
                | litchi_iwa_core::LimitKind::MessageBytes
                | litchi_iwa_core::LimitKind::HeaderBytes
                | litchi_iwa_core::LimitKind::HeaderMemoryBytes
                | litchi_iwa_core::LimitKind::SnappyChunkBytes
                | litchi_iwa_core::LimitKind::SnappyStreamBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedChunkBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedStreamBytes => {
                    ChartCaptionLimitKind::WireBytes
                },
                litchi_iwa_core::LimitKind::Objects
                | litchi_iwa_core::LimitKind::Messages
                | litchi_iwa_core::LimitKind::MessagesPerObject
                | litchi_iwa_core::LimitKind::MetadataItems => ChartCaptionLimitKind::Entries,
                litchi_iwa_core::LimitKind::HeaderFields => ChartCaptionLimitKind::WireFields,
                litchi_iwa_core::LimitKind::HeaderNesting => ChartCaptionLimitKind::WireNesting,
                litchi_iwa_core::LimitKind::SnappyFrames => ChartCaptionLimitKind::WireWork,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            ChartCaptionError::Allocation { amount: requested }
        },
        _ => ChartCaptionError::InvalidSource,
    }
}

fn map_metadata_error(error: metadata_codec::RewriteError) -> ChartCaptionError {
    if let Some(limit) = error.resource_limit() {
        let (kind, observed, maximum) = match limit {
            metadata_codec::RewriteLimit::InputBytes { observed, maximum } => {
                (ChartCaptionLimitKind::WireBytes, observed, maximum)
            },
            metadata_codec::RewriteLimit::OutputBytes { observed, maximum } => {
                (ChartCaptionLimitKind::OutputBytes, observed, maximum)
            },
            metadata_codec::RewriteLimit::Fields { observed, maximum } => {
                (ChartCaptionLimitKind::WireFields, observed, maximum)
            },
            metadata_codec::RewriteLimit::Work { observed, maximum } => {
                (ChartCaptionLimitKind::WireWork, observed, maximum)
            },
            metadata_codec::RewriteLimit::Nesting { observed, maximum } => (
                ChartCaptionLimitKind::WireNesting,
                observed as usize,
                maximum as usize,
            ),
            metadata_codec::RewriteLimit::Components { observed, maximum } => {
                (ChartCaptionLimitKind::Entries, observed, maximum)
            },
            metadata_codec::RewriteLimit::References { observed, maximum } => {
                (ChartCaptionLimitKind::References, observed, maximum)
            },
            metadata_codec::RewriteLimit::Additions { observed, maximum } => {
                (ChartCaptionLimitKind::Entries, observed, maximum)
            },
            _ => return ChartCaptionError::InvalidSource,
        };
        return ChartCaptionError::LimitExceeded {
            kind,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        };
    }
    if let Some(amount) = error.allocation_request() {
        return ChartCaptionError::Allocation { amount };
    }
    ChartCaptionError::InvalidSource
}

fn map_graph_error(error: graph_codec::EncodeError) -> ChartCaptionError {
    match error {
        graph_codec::EncodeError::Resource(limit) => {
            let (kind, observed, maximum) = match limit {
                graph_codec::EncodeLimit::OutputBytes { observed, maximum } => {
                    (ChartCaptionLimitKind::OutputBytes, observed, maximum)
                },
                graph_codec::EncodeLimit::TextBytes { observed, maximum } => {
                    (ChartCaptionLimitKind::CaptionBytes, observed, maximum)
                },
                graph_codec::EncodeLimit::Fields { observed, maximum } => {
                    (ChartCaptionLimitKind::WireFields, observed, maximum)
                },
                graph_codec::EncodeLimit::WorkBytes { observed, maximum } => {
                    (ChartCaptionLimitKind::WireWork, observed, maximum)
                },
                graph_codec::EncodeLimit::Nesting { observed, maximum } => (
                    ChartCaptionLimitKind::WireNesting,
                    observed as usize,
                    maximum as usize,
                ),
                graph_codec::EncodeLimit::Allocations { observed, maximum } => {
                    (ChartCaptionLimitKind::Entries, observed, maximum)
                },
                _ => return ChartCaptionError::InvalidSource,
            };
            ChartCaptionError::LimitExceeded {
                kind,
                observed: usize_to_u64(observed),
                maximum: usize_to_u64(maximum),
            }
        },
        graph_codec::EncodeError::Allocation { amount } => ChartCaptionError::Allocation { amount },
        _ => ChartCaptionError::InvalidSource,
    }
}

fn map_rendering_error(
    _error: super::rendering_invalidation::RenderingInvalidationError,
) -> ChartCaptionError {
    ChartCaptionError::InvalidSource
}

fn map_slide_preview_error(error: super::slide_preview::InvalidationError) -> ChartCaptionError {
    match error {
        super::slide_preview::InvalidationError::InvalidSource => ChartCaptionError::InvalidSource,
        super::slide_preview::InvalidationError::Wire(error) => map_wire_error(error),
        super::slide_preview::InvalidationError::Archive(error) => map_core_error(error),
    }
}

fn map_read_error(error: ReadError) -> ChartCaptionError {
    match error {
        ReadError::Archive(error) => map_archive_error(error),
        ReadError::SemanticLimit {
            kind,
            observed,
            maximum,
            ..
        } => ChartCaptionError::LimitExceeded {
            kind: match kind {
                SemanticLimitKind::Objects => ChartCaptionLimitKind::Entries,
                SemanticLimitKind::Slides => ChartCaptionLimitKind::Slides,
                SemanticLimitKind::References => ChartCaptionLimitKind::References,
                SemanticLimitKind::TextStorages => ChartCaptionLimitKind::TextStorages,
                SemanticLimitKind::TextFragments => ChartCaptionLimitKind::TextFragments,
                SemanticLimitKind::TextBytes => ChartCaptionLimitKind::TextBytes,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        ReadError::PayloadLimit {
            kind,
            observed,
            maximum,
            ..
        } => ChartCaptionError::LimitExceeded {
            kind: match kind {
                super::PayloadLimitKind::Bytes => ChartCaptionLimitKind::WireBytes,
                super::PayloadLimitKind::Fields => ChartCaptionLimitKind::WireFields,
                super::PayloadLimitKind::Nesting => ChartCaptionLimitKind::WireNesting,
                super::PayloadLimitKind::Work => ChartCaptionLimitKind::WireWork,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        ReadError::Allocation { amount, .. } => ChartCaptionError::Allocation { amount },
        _ => ChartCaptionError::InvalidSource,
    }
}

fn map_wire_error(error: litchi_iwa_common::Error) -> ChartCaptionError {
    match error {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => ChartCaptionError::LimitExceeded {
            kind: match kind {
                litchi_iwa_common::LimitKind::InputBytes => ChartCaptionLimitKind::WireBytes,
                litchi_iwa_common::LimitKind::OutputBytes => ChartCaptionLimitKind::OutputBytes,
                litchi_iwa_common::LimitKind::Fields => ChartCaptionLimitKind::WireFields,
                litchi_iwa_common::LimitKind::Nesting => ChartCaptionLimitKind::WireNesting,
                litchi_iwa_common::LimitKind::RewriteWork => ChartCaptionLimitKind::WireWork,
                _ => ChartCaptionLimitKind::WireBytes,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(limit),
        },
        _ => ChartCaptionError::InvalidSource,
    }
}

fn map_chart_caption_codec_error(
    error: keynote_chart_caption_codec::DecodeError,
) -> ChartCaptionError {
    if let Some((observed, maximum)) = error.output_limit_values() {
        return ChartCaptionError::LimitExceeded {
            kind: ChartCaptionLimitKind::OutputBytes,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        };
    }
    if let Some(amount) = error.allocation_amount() {
        return ChartCaptionError::Allocation { amount };
    }
    if let Some(limit) = error.wire_resource_limit() {
        return match limit {
            keynote_chart_caption_codec::WireResourceLimit::Bytes { observed, maximum } => {
                ChartCaptionError::LimitExceeded {
                    kind: ChartCaptionLimitKind::WireBytes,
                    observed: usize_to_u64(observed),
                    maximum: usize_to_u64(maximum),
                }
            },
            keynote_chart_caption_codec::WireResourceLimit::Nesting { observed, maximum } => {
                ChartCaptionError::LimitExceeded {
                    kind: ChartCaptionLimitKind::WireNesting,
                    observed: u64::from(observed),
                    maximum: u64::from(maximum),
                }
            },
            _ => ChartCaptionError::InvalidSource,
        };
    }
    if let Some((observed, maximum)) = error.field_limit_values() {
        return ChartCaptionError::LimitExceeded {
            kind: ChartCaptionLimitKind::WireFields,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        };
    }
    if let Some((observed, maximum)) = error.work_limit_values() {
        return ChartCaptionError::LimitExceeded {
            kind: ChartCaptionLimitKind::WireWork,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        };
    }
    ChartCaptionError::InvalidSource
}

fn map_movie_caption_codec_error(
    error: keynote_movie_caption_codec::DecodeError,
) -> ChartCaptionError {
    if let Some((observed, maximum)) = error.output_limit_values() {
        return ChartCaptionError::LimitExceeded {
            kind: ChartCaptionLimitKind::OutputBytes,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        };
    }
    if let Some(amount) = error.allocation_amount() {
        return ChartCaptionError::Allocation { amount };
    }
    if let Some(limit) = error.wire_resource_limit() {
        return match limit {
            keynote_movie_caption_codec::WireResourceLimit::Bytes { observed, maximum } => {
                ChartCaptionError::LimitExceeded {
                    kind: ChartCaptionLimitKind::WireBytes,
                    observed: usize_to_u64(observed),
                    maximum: usize_to_u64(maximum),
                }
            },
            keynote_movie_caption_codec::WireResourceLimit::Nesting { observed, maximum } => {
                ChartCaptionError::LimitExceeded {
                    kind: ChartCaptionLimitKind::WireNesting,
                    observed: u64::from(observed),
                    maximum: u64::from(maximum),
                }
            },
            _ => ChartCaptionError::InvalidSource,
        };
    }
    if let Some((observed, maximum)) = error.field_limit_values() {
        return ChartCaptionError::LimitExceeded {
            kind: ChartCaptionLimitKind::WireFields,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        };
    }
    if let Some((observed, maximum)) = error.work_limit_values() {
        return ChartCaptionError::LimitExceeded {
            kind: ChartCaptionLimitKind::WireWork,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        };
    }
    ChartCaptionError::InvalidSource
}

pub(super) fn map_caption_info_codec_error(
    error: pages_movie_caption_codec::DecodeError,
) -> ChartCaptionError {
    if let Some((observed, maximum)) = error.message_byte_limit_values() {
        return ChartCaptionError::LimitExceeded {
            kind: ChartCaptionLimitKind::WireBytes,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        };
    }
    if let Some((observed, maximum)) = error.field_limit_values() {
        return ChartCaptionError::LimitExceeded {
            kind: ChartCaptionLimitKind::WireFields,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        };
    }
    if let Some((observed, maximum)) = error.work_limit_values() {
        return ChartCaptionError::LimitExceeded {
            kind: ChartCaptionLimitKind::WireWork,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        };
    }
    ChartCaptionError::InvalidSource
}

pub(super) fn map_slide_text_error(error: super::slide_text::SlideTextError) -> ChartCaptionError {
    match error {
        super::slide_text::SlideTextError::UnsupportedSource => {
            ChartCaptionError::UnsupportedSource
        },
        super::slide_text::SlideTextError::DependentContent
        | super::slide_text::SlideTextError::ObjectMarkerReplacement => {
            ChartCaptionError::UnsupportedDependency
        },
        super::slide_text::SlideTextError::LimitExceeded {
            observed, maximum, ..
        } => ChartCaptionError::LimitExceeded {
            kind: ChartCaptionLimitKind::WireWork,
            observed,
            maximum,
        },
        super::slide_text::SlideTextError::Allocation { amount } => {
            ChartCaptionError::Allocation { amount }
        },
        super::slide_text::SlideTextError::Verification => ChartCaptionError::Verification,
        _ => ChartCaptionError::InvalidSource,
    }
}

fn usize_to_u64(value: usize) -> u64 {
    u64::try_from(value).unwrap_or(u64::MAX)
}
