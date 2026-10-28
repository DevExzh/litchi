//! Source-preserving writer for Pages body-table comment transactions.
//!
//! This module is deliberately private to the selector transaction.  It keeps
//! native object identifiers, UUIDs, archive ownership, and PackageMetadata
//! selectors out of the public Pages API.  The writer stages all physical
//! changes in copy-on-write component archives and publishes one prepared ZIP
//! reassembly only after the operation-wide read ledger has admitted the
//! candidate buffers.

#![allow(
    clippy::arbitrary_source_item_ordering,
    clippy::too_many_arguments,
    clippy::too_many_lines,
    reason = "The private writer keeps bounded graph admission beside the exact wire edits."
)]

use std::collections::{BTreeMap, HashSet};
use std::mem::size_of;
use std::num::NonZeroU64;

use litchi_iwa_archive::package::EntryEdit;
use litchi_iwa_common::wire::{self, NestedFieldEdit, NestedFieldReplacement};
use litchi_iwa_common::{WireLimits, decode_varint_from_bytes};
use litchi_iwa_core::archive::{
    FieldInfo, FieldObjectReferenceTransition, FieldType, MessageInfo, ObjectReferenceTransition,
};
use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
use litchi_iwa_protos::annotation_author_codec;
use litchi_iwa_protos::comment_storage_codec::{
    self, CommentStorageReplyRewrite, CommentStorageVisitor, DateSnapshot, UuidSnapshot,
};
use litchi_iwa_protos::numbers_table_cell_storage_codec as table_storage;
use litchi_iwa_protos::package_metadata_codec::{
    self as metadata_codec, CombinedBatch, CombinedSaveTokenBatch, ComponentSelector,
    ExternalReferenceAddition, ExternalReferenceRemoval, ObjectUuidAddition, ObjectUuidRemoval,
    PackageMetadataVisitor, RemovalBatch, RemovalSaveTokenBatch,
    RewriteError as MetadataRewriteError, RewriteOptions as MetadataRewriteOptions, SaveTokenBatch,
    UuidBits,
};

use super::Package;
use super::body_table_comments::{
    BodyTableCommentError as Error, BodyTableCommentLimitKind as LimitKind, MAX_COMMENT_TEXT_BYTES,
};

const COMMENT_STORAGE_MESSAGE_TYPE: u32 = 3_056;
const METADATA_MESSAGE_TYPE: u32 = 11_006;
const ANNOTATION_AUTHOR_MESSAGE_TYPE: u32 = 212;
const ANNOTATION_AUTHOR_STORAGE_MESSAGE_TYPE: u32 = 213;
const COMMENT_TEXT_FIELD: u32 = 1;
const COMMENT_UUID_FIELD: u32 = 5;
const UUID_LOWER_FIELD: u32 = 1;
const UUID_UPPER_FIELD: u32 = 2;
#[allow(dead_code)]
const REFERENCE_IDENTIFIER_FIELD: u32 = 1;
const TABLE_DATA_LIST_MESSAGE_TYPE: u32 = 6_005;
const TABLE_COMMENT_LIST_TYPE: u64 = 10;
const TABLE_COMMENT_LIST_ENTRIES_FIELD: u32 = 3;
const TABLE_COMMENT_LIST_NEXT_ID_FIELD: u32 = 2;
const TABLE_COMMENT_LIST_TYPE_FIELD: u32 = 1;
const TABLE_COMMENT_ENTRY_KEY_FIELD: u32 = 1;
const TABLE_COMMENT_ENTRY_REFCOUNT_FIELD: u32 = 2;
const TABLE_COMMENT_ENTRY_STORAGE_FIELD: u32 = 10;
const TABLE_COMMENT_REFERENCE_IDENTIFIER_FIELD: u32 = 1;
const TABLE_TILE_ROWS_FIELD: u32 = 5;
const TABLE_TILE_ROW_INDEX_FIELD: u32 = 1;
const TABLE_TILE_ROW_CELL_COUNT_FIELD: u32 = 2;
const TABLE_TILE_ROW_CELL_STORAGE_FIELD: u32 = 6;
const TABLE_TILE_ROW_CELL_OFFSETS_FIELD: u32 = 7;
const TABLE_TILE_ROW_WIDE_OFFSETS_FIELD: u32 = 8;
const TABLE_MODEL_COMMENT_STORAGE_FIELD_PATH: [u32; 2] = [4, 19];

const MAX_COMMENT_RECURSION: u32 = 64;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct CandidateAdmission {
    entries: usize,
    objects: usize,
    messages: usize,
    payload_bytes: usize,
    wire_bytes: usize,
    wire_fields: usize,
    wire_work: usize,
    references: usize,
    scratch_bytes: usize,
    allocations: usize,
    catalog_bytes: usize,
    semantic_bytes: usize,
    metadata_bytes: usize,
}

impl CandidateAdmission {
    pub(crate) const EMPTY: Self = Self {
        entries: 0,
        objects: 0,
        messages: 0,
        payload_bytes: 0,
        wire_bytes: 0,
        wire_fields: 0,
        wire_work: 0,
        references: 0,
        scratch_bytes: 0,
        allocations: 0,
        catalog_bytes: 0,
        semantic_bytes: 0,
        metadata_bytes: 0,
    };

    pub(crate) fn charge(
        self,
        budget: &mut ReadBudget,
        compressed_bytes: usize,
    ) -> Result<(), Error> {
        budget.wire_work(
            compressed_bytes
                .checked_add(self.wire_work)
                .ok_or(Error::InvalidSource)?,
        )?;
        budget.scratch(
            self.scratch_bytes
                .checked_add(self.catalog_bytes)
                .and_then(|bytes| bytes.checked_add(self.semantic_bytes))
                .and_then(|bytes| bytes.checked_add(self.metadata_bytes))
                .ok_or(Error::InvalidSource)?,
        )?;
        ReadBudget::charge_counter(
            LimitKind::Allocations,
            &mut budget.allocations,
            self.allocations.max(1),
            budget.max_allocations,
        )
    }
}

#[derive(Debug, Clone, Copy)]
pub(crate) struct ReadBudget {
    input: usize,
    output: usize,
    entries: usize,
    payload: usize,
    total_payload: usize,
    objects: usize,
    messages: usize,
    wire_bytes: usize,
    wire_fields: usize,
    wire_nesting: usize,
    wire_work: usize,
    references: usize,
    replies: usize,
    text_bytes: usize,
    scratch_bytes: usize,
    allocations: usize,
    max_input: usize,
    max_output: usize,
    max_entries: usize,
    max_payload: usize,
    max_total_payload: usize,
    max_objects: usize,
    max_messages: usize,
    max_wire_bytes: usize,
    max_wire_fields: usize,
    max_wire_nesting: usize,
    max_wire_work: usize,
    max_references: usize,
    max_replies: usize,
    max_text_bytes: usize,
    max_scratch_bytes: usize,
    max_allocations: usize,
}

impl ReadBudget {
    pub(crate) fn for_package(package: &Package) -> Result<Self, Error> {
        let physical = package.state.source.limits();
        let archive = physical
            .effective_archive_limits()
            .map_err(|_| Error::UnsupportedSource)?;
        let components = package.state.source.components().len().max(1);
        let source_bytes = package.source_bytes().len();
        let max_input = usize::try_from(physical.max_input_bytes())
            .unwrap_or(usize::MAX)
            .max(1);
        let max_total = usize::try_from(physical.max_total_bytes())
            .unwrap_or(usize::MAX)
            .max(1);
        let max_objects = archive.max_objects().saturating_mul(components).max(1);
        let max_messages = archive.max_messages().saturating_mul(components).max(1);
        let max_fields = archive
            .max_header_fields()
            .saturating_mul(components)
            .max(1);
        let max_refs = archive
            .max_metadata_items()
            .saturating_mul(components)
            .max(1);
        let max_work = physical.max_iwa_stream_bytes().saturating_mul(64).max(1);
        // A transaction can retain the source archive, one copy-on-write
        // working archive, the prepared ZIP payload, and the reopened
        // candidate at once.  Header memory is still the per-archive floor,
        // while two physical package ceilings bound that whole admission.
        let package_workspace = max_total / 64;
        let max_scratch = archive
            .max_header_memory_bytes()
            .saturating_mul(components)
            .saturating_add(package_workspace)
            .max(
                max_total
                    .saturating_mul(2)
                    .saturating_add(package_workspace),
            )
            .max(1);
        let max_allocations = max_objects
            .saturating_mul(32)
            .saturating_add(max_messages.saturating_mul(16))
            .max(256);
        let mut budget = Self {
            input: 0,
            output: 0,
            entries: 0,
            payload: 0,
            total_payload: 0,
            objects: 0,
            messages: 0,
            wire_bytes: 0,
            wire_fields: 0,
            wire_nesting: 0,
            wire_work: 0,
            references: 0,
            replies: 0,
            text_bytes: 0,
            scratch_bytes: 0,
            allocations: 0,
            max_input,
            max_output: max_total,
            max_entries: physical.max_entries().max(1),
            max_payload: max_total,
            max_total_payload: max_total,
            max_objects,
            max_messages,
            max_wire_bytes: max_total,
            max_wire_fields: max_fields,
            max_wire_nesting: archive.max_header_nesting().max(1),
            max_wire_work: max_work,
            max_references: max_refs,
            max_replies: max_refs,
            max_text_bytes: max_total.min(MAX_COMMENT_TEXT_BYTES),
            max_scratch_bytes: max_scratch.max(1),
            max_allocations,
        };
        Self::charge_counter(
            LimitKind::InputBytes,
            &mut budget.input,
            source_bytes,
            budget.max_input,
        )?;
        Self::charge_counter(
            LimitKind::Entries,
            &mut budget.entries,
            package.state.source.package().len(),
            budget.max_entries,
        )?;
        Self::charge_counter(
            LimitKind::PayloadObjects,
            &mut budget.objects,
            package.state.object_count,
            budget.max_objects,
        )?;
        Self::charge_counter(
            LimitKind::Allocations,
            &mut budget.allocations,
            1,
            budget.max_allocations,
        )?;
        Ok(budget)
    }

    pub(crate) fn charge_counter(
        kind: LimitKind,
        current: &mut usize,
        amount: usize,
        maximum: usize,
    ) -> Result<(), Error> {
        let observed = current.checked_add(amount).ok_or(Error::LimitExceeded {
            kind,
            observed: u64::MAX,
            maximum: maximum as u64,
        })?;
        if observed > maximum {
            return Err(Error::LimitExceeded {
                kind,
                observed: observed as u64,
                maximum: maximum as u64,
            });
        }
        *current = observed;
        Ok(())
    }

    fn residual(
        &self,
        kind: LimitKind,
        requested: usize,
        current: usize,
        maximum: usize,
    ) -> Result<usize, Error> {
        let remaining = maximum.checked_sub(current).ok_or(Error::LimitExceeded {
            kind,
            observed: u64::MAX,
            maximum: maximum as u64,
        })?;
        if remaining == 0 {
            return Err(Error::LimitExceeded {
                kind,
                observed: current.saturating_add(requested) as u64,
                maximum: maximum as u64,
            });
        }
        Ok(requested.max(1).min(remaining))
    }

    pub(crate) fn copy_text(&mut self, length: usize) -> Result<(), Error> {
        Self::charge_counter(
            LimitKind::TextBytes,
            &mut self.text_bytes,
            length,
            self.max_text_bytes,
        )?;
        Self::charge_counter(
            LimitKind::Allocations,
            &mut self.allocations,
            1,
            self.max_allocations,
        )
    }

    fn scratch(&mut self, amount: usize) -> Result<(), Error> {
        Self::charge_counter(
            LimitKind::ScratchBytes,
            &mut self.scratch_bytes,
            amount,
            self.max_scratch_bytes,
        )
    }

    pub(crate) fn wire_work(&mut self, amount: usize) -> Result<(), Error> {
        Self::charge_counter(
            LimitKind::WireWork,
            &mut self.wire_work,
            amount,
            self.max_wire_work,
        )
    }

    fn wire_fields(&mut self, amount: usize) -> Result<(), Error> {
        Self::charge_counter(
            LimitKind::WireFields,
            &mut self.wire_fields,
            amount,
            self.max_wire_fields,
        )
    }

    fn references(&mut self, amount: usize) -> Result<(), Error> {
        Self::charge_counter(
            LimitKind::References,
            &mut self.references,
            amount,
            self.max_references,
        )
    }

    fn message(&mut self, payload: usize) -> Result<(), Error> {
        Self::charge_counter(
            LimitKind::PayloadMessages,
            &mut self.messages,
            1,
            self.max_messages,
        )?;
        Self::charge_counter(
            LimitKind::PayloadBytes,
            &mut self.payload,
            payload,
            self.max_payload,
        )?;
        Self::charge_counter(
            LimitKind::TotalPayloadBytes,
            &mut self.total_payload,
            payload,
            self.max_total_payload,
        )
    }

    fn wire_input(&mut self, bytes: usize, depth: usize) -> Result<(), Error> {
        Self::charge_counter(
            LimitKind::WireBytes,
            &mut self.wire_bytes,
            bytes,
            self.max_wire_bytes,
        )?;
        self.wire_work(bytes)?;
        if depth > self.max_wire_nesting {
            return Err(Error::LimitExceeded {
                kind: LimitKind::WireNesting,
                observed: depth as u64,
                maximum: self.max_wire_nesting as u64,
            });
        }
        self.wire_nesting = self.wire_nesting.max(depth);
        Ok(())
    }

    fn metadata_scan(
        &mut self,
        references: usize,
        fields: usize,
        paths: usize,
    ) -> Result<(), Error> {
        self.references(references)?;
        self.wire_fields(fields)?;
        self.wire_work(references.saturating_add(paths))
    }

    fn codec_depth(&self) -> u32 {
        MAX_COMMENT_RECURSION.min(self.max_wire_nesting.min(u32::MAX as usize) as u32)
    }

    fn observe_codec_depth(&mut self, depth: u32) -> Result<(), Error> {
        let depth = depth as usize;
        if depth > self.max_wire_nesting {
            return Err(Error::LimitExceeded {
                kind: LimitKind::WireNesting,
                observed: depth as u64,
                maximum: self.max_wire_nesting as u64,
            });
        }
        self.wire_nesting = self.wire_nesting.max(depth);
        Ok(())
    }

    fn codec_options(
        &self,
        source_len: usize,
    ) -> Result<comment_storage_codec::DecodeOptions, Error> {
        let bytes = source_len.max(1);
        Ok(comment_storage_codec::DecodeOptions::new(
            self.residual(
                LimitKind::WireBytes,
                bytes,
                self.wire_bytes,
                self.max_wire_bytes,
            )?,
            self.residual(
                LimitKind::WireFields,
                bytes,
                self.wire_fields,
                self.max_wire_fields,
            )?,
            self.residual(
                LimitKind::WireWork,
                bytes.saturating_mul(32),
                self.wire_work,
                self.max_wire_work,
            )?,
            self.codec_depth(),
            self.residual(
                LimitKind::References,
                bytes,
                self.references,
                self.max_references,
            )?,
            self.residual(
                LimitKind::TextBytes,
                bytes,
                self.text_bytes,
                self.max_text_bytes,
            )?,
        ))
    }

    fn annotation_options(
        &self,
        source_len: usize,
    ) -> Result<annotation_author_codec::DecodeOptions, Error> {
        let bytes = source_len.max(1);
        Ok(annotation_author_codec::DecodeOptions::new(
            self.residual(
                LimitKind::WireBytes,
                bytes,
                self.wire_bytes,
                self.max_wire_bytes,
            )?,
            self.residual(
                LimitKind::WireFields,
                bytes,
                self.wire_fields,
                self.max_wire_fields,
            )?,
            self.residual(
                LimitKind::WireWork,
                bytes.saturating_mul(64),
                self.wire_work,
                self.max_wire_work,
            )?,
            self.codec_depth(),
            self.residual(
                LimitKind::References,
                bytes,
                self.references,
                self.max_references,
            )?,
            self.residual(
                LimitKind::TextBytes,
                bytes,
                self.text_bytes,
                self.max_text_bytes,
            )?,
            self.residual(
                LimitKind::Allocations,
                bytes.saturating_mul(8),
                self.allocations,
                self.max_allocations,
            )?,
        ))
    }

    fn charge_comment_report(
        &mut self,
        report: comment_storage_codec::DecodeReport,
    ) -> Result<(), Error> {
        Self::charge_counter(
            LimitKind::WireBytes,
            &mut self.wire_bytes,
            report.source_bytes(),
            self.max_wire_bytes,
        )?;
        Self::charge_counter(
            LimitKind::WireFields,
            &mut self.wire_fields,
            report.fields(),
            self.max_wire_fields,
        )?;
        Self::charge_counter(
            LimitKind::WireWork,
            &mut self.wire_work,
            report.work_bytes(),
            self.max_wire_work,
        )?;
        self.observe_codec_depth(report.max_depth())?;
        Self::charge_counter(
            LimitKind::References,
            &mut self.references,
            report.references(),
            self.max_references,
        )?;
        Self::charge_counter(
            LimitKind::Replies,
            &mut self.replies,
            report.replies(),
            self.max_replies,
        )
    }

    fn charge_author_report(
        &mut self,
        report: annotation_author_codec::DecodeReport,
    ) -> Result<(), Error> {
        Self::charge_counter(
            LimitKind::WireBytes,
            &mut self.wire_bytes,
            report.source_bytes(),
            self.max_wire_bytes,
        )?;
        Self::charge_counter(
            LimitKind::WireFields,
            &mut self.wire_fields,
            report.fields(),
            self.max_wire_fields,
        )?;
        Self::charge_counter(
            LimitKind::WireWork,
            &mut self.wire_work,
            report.work_bytes(),
            self.max_wire_work,
        )?;
        self.observe_codec_depth(report.max_depth())?;
        Self::charge_counter(
            LimitKind::References,
            &mut self.references,
            report.references(),
            self.max_references,
        )?;
        Self::charge_counter(
            LimitKind::Allocations,
            &mut self.allocations,
            report.allocations(),
            self.max_allocations,
        )
    }

    fn charge_target(&mut self, bytes: usize) -> Result<(), Error> {
        Self::charge_counter(
            LimitKind::OutputBytes,
            &mut self.output,
            bytes,
            self.max_output,
        )?;
        self.scratch(bytes)?;
        Self::charge_counter(
            LimitKind::Allocations,
            &mut self.allocations,
            1,
            self.max_allocations,
        )
    }
}

#[derive(Debug)]
struct ReplyIds<'budget> {
    ids: Vec<u64>,
    allocation_failed: Option<usize>,
    budget_error: Option<Error>,
    budget: &'budget mut ReadBudget,
}
impl<'budget> ReplyIds<'budget> {
    fn new(budget: &'budget mut ReadBudget) -> Self {
        Self {
            ids: Vec::new(),
            allocation_failed: None,
            budget_error: None,
            budget,
        }
    }
    fn into_parts(self) -> (Vec<u64>, Option<usize>, Option<Error>) {
        (self.ids, self.allocation_failed, self.budget_error)
    }
}
impl CommentStorageVisitor for ReplyIds<'_> {
    fn visit_reply(
        &mut self,
        reply: comment_storage_codec::ReferenceRecord<'_>,
    ) -> Result<(), comment_storage_codec::DecodeError> {
        if self.allocation_failed.is_some() || self.budget_error.is_some() {
            return Ok(());
        }
        if let Err(error) = ReadBudget::charge_counter(
            LimitKind::Allocations,
            &mut self.budget.allocations,
            1,
            self.budget.max_allocations,
        ) {
            self.budget_error = Some(error);
            return Ok(());
        }
        if let Err(error) = self.budget.scratch(size_of::<u64>()) {
            self.budget_error = Some(error);
            return Ok(());
        }
        if self.ids.try_reserve(1).is_err() {
            self.allocation_failed = Some(self.ids.len().saturating_add(1));
            return Ok(());
        }
        self.ids.push(reply.identifier());
        Ok(())
    }
}

fn wire_limits(source: &Package) -> Result<WireLimits, Error> {
    let bytes = source.source_bytes().len().max(1);
    WireLimits::default()
        .with_input_bytes(bytes)
        .and_then(|v| v.with_fields(bytes))
        .and_then(|v| v.with_rewrite_work(bytes.saturating_mul(32)))
        .map_err(|_| Error::InvalidSource)
}
fn map_comment_decode_error(
    error: comment_storage_codec::DecodeError,
    budget: &ReadBudget,
) -> Error {
    let Some(limit) = error.resource_limit() else {
        return Error::InvalidSource;
    };
    let (kind, observed) = match limit {
        comment_storage_codec::DecodeLimit::Bytes { observed, .. }
        | comment_storage_codec::DecodeLimit::OutputBytes { observed, .. }
        | comment_storage_codec::DecodeLimit::ReferenceBytes { observed, .. } => {
            (LimitKind::WireBytes, observed)
        },
        comment_storage_codec::DecodeLimit::Fields { observed, .. } => {
            (LimitKind::WireFields, observed)
        },
        comment_storage_codec::DecodeLimit::Work { observed, .. } => {
            (LimitKind::WireWork, observed)
        },
        comment_storage_codec::DecodeLimit::Nesting { observed, .. } => {
            (LimitKind::WireNesting, observed as usize)
        },
        comment_storage_codec::DecodeLimit::References { observed, .. } => {
            (LimitKind::References, observed)
        },
        comment_storage_codec::DecodeLimit::Replies { observed, .. } => {
            (LimitKind::Replies, observed)
        },
        comment_storage_codec::DecodeLimit::Text { observed, .. } => {
            (LimitKind::TextBytes, observed)
        },
        comment_storage_codec::DecodeLimit::Allocations { observed, .. } => {
            (LimitKind::Allocations, observed)
        },
        comment_storage_codec::DecodeLimit::Scratch { observed, .. }
        | comment_storage_codec::DecodeLimit::Retained { observed, .. } => {
            (LimitKind::ScratchBytes, observed)
        },
        _ => return Error::InvalidSource,
    };
    let (current, maximum) = match kind {
        LimitKind::WireBytes => (budget.wire_bytes, budget.max_wire_bytes),
        LimitKind::WireFields => (budget.wire_fields, budget.max_wire_fields),
        LimitKind::WireWork => (budget.wire_work, budget.max_wire_work),
        LimitKind::WireNesting => (budget.wire_nesting, budget.max_wire_nesting),
        LimitKind::ScratchBytes => (budget.scratch_bytes, budget.max_scratch_bytes),
        LimitKind::References => (budget.references, budget.max_references),
        LimitKind::Replies => (budget.replies, budget.max_replies),
        LimitKind::TextBytes => (budget.text_bytes, budget.max_text_bytes),
        LimitKind::Allocations => (budget.allocations, budget.max_allocations),
        _ => (0, usize::MAX),
    };
    Error::LimitExceeded {
        kind,
        observed: if matches!(kind, LimitKind::WireNesting) {
            observed as u64
        } else {
            current.saturating_add(observed) as u64
        },
        maximum: maximum as u64,
    }
}

fn map_author_decode_error(
    error: annotation_author_codec::DecodeError,
    budget: &ReadBudget,
) -> Error {
    let Some(limit) = error.resource_limit() else {
        return Error::InvalidSource;
    };
    let (kind, observed) = match limit {
        annotation_author_codec::DecodeLimit::Bytes { observed, .. }
        | annotation_author_codec::DecodeLimit::OutputBytes { observed, .. } => {
            (LimitKind::WireBytes, observed)
        },
        annotation_author_codec::DecodeLimit::Fields { observed, .. } => {
            (LimitKind::WireFields, observed)
        },
        annotation_author_codec::DecodeLimit::Work { observed, .. } => {
            (LimitKind::WireWork, observed)
        },
        annotation_author_codec::DecodeLimit::Nesting { observed, .. } => {
            (LimitKind::WireNesting, observed as usize)
        },
        annotation_author_codec::DecodeLimit::References { observed, .. } => {
            (LimitKind::References, observed)
        },
        annotation_author_codec::DecodeLimit::Text { observed, .. } => {
            (LimitKind::TextBytes, observed)
        },
        annotation_author_codec::DecodeLimit::Allocations { observed, .. } => {
            (LimitKind::Allocations, observed)
        },
        annotation_author_codec::DecodeLimit::Scratch { observed, .. }
        | annotation_author_codec::DecodeLimit::Retained { observed, .. } => {
            (LimitKind::ScratchBytes, observed)
        },
        _ => return Error::InvalidSource,
    };
    let (current, maximum) = match kind {
        LimitKind::WireBytes => (budget.wire_bytes, budget.max_wire_bytes),
        LimitKind::WireFields => (budget.wire_fields, budget.max_wire_fields),
        LimitKind::WireWork => (budget.wire_work, budget.max_wire_work),
        LimitKind::WireNesting => (budget.wire_nesting, budget.max_wire_nesting),
        LimitKind::ScratchBytes => (budget.scratch_bytes, budget.max_scratch_bytes),
        LimitKind::References => (budget.references, budget.max_references),
        LimitKind::TextBytes => (budget.text_bytes, budget.max_text_bytes),
        LimitKind::Allocations => (budget.allocations, budget.max_allocations),
        _ => (0, usize::MAX),
    };
    Error::LimitExceeded {
        kind,
        observed: if matches!(kind, LimitKind::WireNesting) {
            observed as u64
        } else {
            current.saturating_add(observed) as u64
        },
        maximum: maximum as u64,
    }
}

/// Native-free selection evidence passed from the focused Pages table
/// adapter into the shared comment graph writer.  The fields are crate
/// private and are populated only after the rooted table selector, model,
/// tile, list, and cell route have been validated.
#[derive(Debug, Clone)]
pub(crate) struct TableCommentSelection {
    /// Component containing the table model and its comment-list edge.
    pub(crate) component: Box<str>,
    /// Component containing the selected tile payload.
    pub(crate) tile_component: Box<str>,
    pub(crate) model_identifier: u64,
    pub(crate) model_message_index: usize,
    pub(crate) model_message_type: u32,
    pub(crate) tile_identifier: u64,
    pub(crate) tile_message_index: usize,
    pub(crate) tile_message_type: u32,
    pub(crate) tile_row: u32,
    pub(crate) column: u32,
    pub(crate) table_columns: u32,
    pub(crate) comment_key: Option<u32>,
    pub(crate) list_identifier: Option<u64>,
    pub(crate) root_identifier: Option<u64>,
}

// Native drawable routing retained only by the generic wire helper below.
// Focused table transactions use `TableCommentSelection` instead.
#[allow(dead_code)]
#[derive(Debug, Clone)]
struct SelectionFacts {
    component: Box<str>,
    identity: NonZeroU64,
    message_index: usize,
    message_type: u32,
    route: Vec<u32>,
}

/// One selector-first semantic table-cell comment operation.  Reply ordinals
/// are resolved by the focused adapter before this private writer is called.
#[derive(Debug, Clone)]
pub(crate) enum TableCommentOperation {
    Set(Box<str>),
    Clear,
    AddReply(Box<str>),
    SetReply { index: usize, text: Box<str> },
    RemoveReply { index: usize },
}

/// Physical result of a table-cell comment mutation.  Candidate admission is
/// intentionally retained so the caller can charge/reopen the exact output
/// using the same bounded ledger as drawable comments.
#[derive(Debug)]
pub(crate) struct TableCommentMutationOutput {
    pub(crate) bytes: Vec<u8>,
    pub(crate) touched_components: usize,
    pub(crate) candidate_admission: CandidateAdmission,
}
#[derive(Debug)]
struct Working<'source> {
    source: &'source Package,
    archives: BTreeMap<String, Archive>,
    reserved_ids: Vec<u64>,
    identities: Vec<NewIdentity>,
    edge_changes: Vec<EdgeChange>,
    removed_identities: Vec<RemovedIdentity>,
    // A transaction removes only a bounded comment graph.  Keep the removal
    // ledger as a fallibly grown vector so its element bytes and linear
    // membership work can be admitted explicitly without hash-table buckets.
    removed_objects: Vec<u64>,
    metadata: Option<MetadataState>,
}

#[derive(Debug, Clone, Copy)]
struct NewIdentity {
    component: usize,
    identifier: u64,
    uuid: UuidSnapshot,
}

#[derive(Debug, Clone, Copy)]
struct EdgeChange {
    source_component: usize,
    old_target_component: Option<usize>,
    old_identifier: Option<u64>,
    new_target_component: Option<usize>,
    new_identifier: Option<u64>,
    weak: Option<bool>,
    remove_old: bool,
}

#[derive(Debug, Clone, Copy)]
struct RemovedIdentity {
    component: usize,
    identifier: u64,
    uuid: UuidSnapshot,
}

#[derive(Debug, Clone)]
struct MetadataState {
    member: Box<str>,
    archive: Archive,
    object_index: usize,
    message_index: usize,
    payload: Vec<u8>,
    last_identifier: u64,
    components: Vec<MetadataComponent>,
    uuids: HashSet<(u64, u64)>,
    uuid_bindings: Vec<MetadataUuid>,
    externals: Vec<MetadataExternal>,
    changed: bool,
}

#[derive(Debug, Clone)]
struct MetadataComponent {
    identifier: u64,
    locator: Box<str>,
    current: bool,
}

#[derive(Debug, Clone, Copy)]
struct MetadataExternal {
    source: u64,
    target: u64,
    object: Option<u64>,
    weak: Option<bool>,
    versioned: bool,
    unknown: bool,
}

#[derive(Debug, Clone, Copy)]
struct StorageFacts {
    author: Option<u64>,
    timestamp: Option<DateSnapshot>,
}

#[derive(Debug, Clone)]
struct AuthorRegistryFacts {
    object_identifier: u64,
    message_index: usize,
    author_identifier: u64,
    author_ordinal: usize,
    author_count: usize,
}

#[derive(Debug, Default)]
struct AuthorOnlyVisitor;

impl CommentStorageVisitor for AuthorOnlyVisitor {
    fn visit_reply(
        &mut self,
        _reply: comment_storage_codec::ReferenceRecord<'_>,
    ) -> Result<(), comment_storage_codec::DecodeError> {
        Ok(())
    }
}

impl<'source> Working<'source> {
    fn new(source: &'source Package, budget: &mut ReadBudget) -> Result<Self, Error> {
        let metadata = read_metadata_state(source, budget)?;
        Ok(Self {
            source,
            archives: BTreeMap::new(),
            reserved_ids: Vec::new(),
            identities: Vec::new(),
            edge_changes: Vec::new(),
            removed_identities: Vec::new(),
            removed_objects: Vec::new(),
            metadata,
        })
    }

    fn load(&mut self, name: &str, budget: &mut ReadBudget) -> Result<&mut Archive, Error> {
        if !self.archives.contains_key(name) {
            let component = self
                .source
                .state
                .source
                .components()
                .get(name)
                .ok_or(Error::InvalidSource)?;
            let limits = self
                .source
                .state
                .source
                .limits()
                .effective_archive_limits()
                .map_err(|_| Error::UnsupportedSource)?;
            let bound = charge_archive_sizing_scan(component.archive(), limits, budget)?;
            budget.scratch(bound)?;
            charge_archive_clone_admission(component.archive(), budget)?;
            ReadBudget::charge_counter(
                LimitKind::Allocations,
                &mut budget.allocations,
                1,
                budget.max_allocations,
            )?;
            let cloned = component.archive().clone();
            budget.copy_text(name.len())?;
            self.archives.insert(name.to_owned(), cloned);
        }
        self.archives.get_mut(name).ok_or(Error::InvalidSource)
    }

    fn find_object(&self, identifier: u64) -> Option<(&str, &ArchiveObject)> {
        for (name, archive) in &self.archives {
            if let Some(object) = archive.object(identifier) {
                return Some((name.as_str(), object));
            }
        }
        for component in self.source.state.source.components().iter() {
            if let Some(object) = component.archive().object(identifier) {
                return Some((component.name(), object));
            }
        }
        None
    }

    fn reserve_id(&mut self, budget: &mut ReadBudget) -> Result<u64, Error> {
        let object_count = self.source.state.object_count;
        budget.wire_work(
            object_count
                .checked_mul(size_of::<u64>())
                .ok_or(Error::InvalidSource)?,
        )?;
        let mut maximum = self
            .source
            .state
            .source
            .components()
            .iter()
            .flat_map(|component| component.archive().objects.iter())
            .filter_map(|object| object.archive_info.identifier)
            .max()
            .unwrap_or(0);
        if let Some(metadata) = &self.metadata {
            maximum = maximum.max(metadata.last_identifier);
        }
        for identity in &self.identities {
            maximum = maximum.max(identity.identifier);
        }
        let mut candidate = maximum.checked_add(1).ok_or(Error::InvalidSource)?;
        loop {
            if candidate != 0
                && !self.reserved_ids.contains(&candidate)
                && self.find_object(candidate).is_none()
            {
                ReadBudget::charge_counter(
                    LimitKind::Allocations,
                    &mut budget.allocations,
                    size_of::<u64>(),
                    budget.max_allocations,
                )?;
                self.reserved_ids.push(candidate);
                return Ok(candidate);
            }
            candidate = candidate.checked_add(1).ok_or(Error::InvalidSource)?;
        }
    }

    fn fresh_uuid(&self, identifier: u64) -> Result<UuidSnapshot, Error> {
        let mut lower = identifier ^ 0x9e37_79b9_7f4a_7c15;
        let mut upper = identifier.rotate_left(29) ^ 0xd1b5_4a32_d192_ed03;
        for _ in 0..1024 {
            if lower != 0 && upper != 0 {
                let candidate = (lower, upper);
                let used = self
                    .metadata
                    .as_ref()
                    .is_some_and(|metadata| metadata.uuids.contains(&candidate))
                    || self.identities.iter().any(|identity| {
                        (identity.uuid.lower(), identity.uuid.upper()) == candidate
                    });
                if !used {
                    return Ok(UuidSnapshot::from_parts(lower, upper));
                }
            }
            lower = lower.checked_add(1).ok_or(Error::InvalidSource)?;
            upper = upper
                .checked_add(0x9e37_79b9_7f4a_7c15)
                .ok_or(Error::InvalidSource)?;
        }
        Err(Error::UnsupportedDependency)
    }

    fn register_identity(&mut self, component: usize, identifier: u64, uuid: UuidSnapshot) {
        self.identities.push(NewIdentity {
            component,
            identifier,
            uuid,
        });
    }

    #[allow(dead_code)]
    fn record_edge(
        &mut self,
        selection: &SelectionFacts,
        old_identifier: Option<u64>,
        new_identifier: Option<u64>,
        budget: &mut ReadBudget,
    ) -> Result<(), Error> {
        self.record_edge_from_component(
            selection.component.as_ref(),
            old_identifier,
            new_identifier,
            budget,
        )
    }

    fn record_edge_from_component(
        &mut self,
        source_name: &str,
        old_identifier: Option<u64>,
        new_identifier: Option<u64>,
        budget: &mut ReadBudget,
    ) -> Result<(), Error> {
        let source_component = self
            .source
            .state
            .source
            .components()
            .iter()
            .position(|component| component.name() == source_name)
            .ok_or(Error::InvalidSource)?;
        let old_target_component = old_identifier.and_then(|identifier| {
            self.find_object(identifier).and_then(|(name, _)| {
                self.source
                    .state
                    .source
                    .components()
                    .iter()
                    .position(|component| component.name() == name)
            })
        });
        let new_target_component = new_identifier.and_then(|identifier| {
            self.find_object(identifier).and_then(|(name, _)| {
                self.source
                    .state
                    .source
                    .components()
                    .iter()
                    .position(|component| component.name() == name)
            })
        });
        if old_target_component == Some(source_component)
            && new_target_component == Some(source_component)
        {
            return Ok(());
        }
        let Some(metadata) = self.metadata.as_ref() else {
            return Ok(());
        };
        let source_metadata =
            metadata_component_for_index(metadata, source_component, self.source)?;
        let old_metadata = old_target_component
            .map(|component| metadata_component_for_index(metadata, component, self.source))
            .transpose()?;
        let mut weak = new_identifier.map(|_| true);
        let mut remove_old = false;
        if let (Some(target), Some(identifier)) = (old_metadata, old_identifier) {
            budget.wire_work(metadata.externals.len())?;
            let mut matching = None;
            for edge in &metadata.externals {
                if edge.source == source_metadata.identifier
                    && edge.target == target.identifier
                    && edge.object == Some(identifier)
                {
                    if matching.replace(edge).is_some() {
                        return Err(Error::InvalidSource);
                    }
                }
            }
            if let Some(edge) = matching {
                // PackageMetadata external references are only removable when
                // they cross component boundaries.  A same-component record
                // is malformed; deleting the object here would leave that
                // record dangling, especially when no UUID binding exists.
                if target.identifier == source_metadata.identifier {
                    return Err(Error::UnsupportedDependency);
                }
                if edge.versioned || edge.unknown {
                    return Err(Error::UnsupportedDependency);
                }
                weak = edge.weak;
                remove_old = true;
            }
        }
        if let (Some(target), Some(identifier)) = (new_target_component, new_identifier) {
            let target_metadata = metadata_component_for_index(metadata, target, self.source)?;
            budget.wire_work(metadata.externals.len())?;
            let already_owned = metadata.externals.iter().any(|edge| {
                edge.source == source_metadata.identifier
                    && edge.target == target_metadata.identifier
                    && edge.object == Some(identifier)
                    && !edge.versioned
                    && !edge.unknown
            });
            if already_owned && old_identifier.is_none() {
                return Ok(());
            }
        }
        let edge = EdgeChange {
            source_component,
            old_target_component,
            old_identifier,
            new_target_component,
            new_identifier,
            weak,
            remove_old,
        };
        ReadBudget::charge_counter(
            LimitKind::Allocations,
            &mut budget.allocations,
            size_of::<EdgeChange>(),
            budget.max_allocations,
        )?;
        self.edge_changes.push(edge);
        Ok(())
    }

    /// Count independent native owners of one storage object.
    ///
    /// Archive message metadata is the primary witness.  PackageMetadata
    /// component edges are an additional witness only when their source
    /// component has no native message witness: native producers commonly
    /// record the same cross-component edge in both projections.  Counting
    /// that pair twice would clone a uniquely owned thread; ignoring a
    /// genuinely separate metadata source would mutate a shared thread.
    fn owner_count(&self, identifier: u64, budget: &mut ReadBudget) -> Result<usize, Error> {
        if identifier == 0 {
            return Err(Error::InvalidSource);
        }
        let (object_bound, message_bound) = self.source.state.source.components().iter().try_fold(
            (0usize, 0usize),
            |(objects, messages), component| {
                let archive = self
                    .archives
                    .get(component.name())
                    .unwrap_or_else(|| component.archive());
                let objects = objects
                    .checked_add(archive.objects.len())
                    .ok_or(Error::InvalidSource)?;
                let messages = archive.objects.iter().try_fold(messages, |count, object| {
                    count
                        .checked_add(object.messages.len())
                        .ok_or(Error::InvalidSource)
                })?;
                Ok::<(usize, usize), Error>((objects, messages))
            },
        )?;
        // Admit the census walk before looking at any archive metadata.  The
        // detailed field/reference counters below refine this coarse source
        // bound and the ownership HashSet is admitted separately.
        let census_work = object_bound
            .checked_mul(size_of::<ArchiveObject>())
            .and_then(|bytes| {
                bytes.checked_add(message_bound.checked_mul(size_of::<RawMessage>())?)
            })
            .ok_or(Error::InvalidSource)?;
        budget.wire_work(census_work)?;
        budget.scratch(
            object_bound
                .checked_mul(size_of::<ArchiveObject>())
                .ok_or(Error::InvalidSource)?,
        )?;
        for component in self.source.state.source.components().iter() {
            let archive = self
                .archives
                .get(component.name())
                .unwrap_or_else(|| component.archive());
            for object in &archive.objects {
                for info in &object.archive_info.message_infos {
                    let info_fields = info
                        .field_infos
                        .len()
                        .checked_add(1)
                        .ok_or(Error::InvalidSource)?;
                    budget.wire_fields(info_fields)?;
                    let info_references = info
                        .object_references
                        .len()
                        .checked_add(info.data_references.len())
                        .ok_or(Error::InvalidSource)?;
                    budget.references(info_references)?;
                    budget.wire_work(
                        info_fields
                            .checked_add(info_references)
                            .ok_or(Error::InvalidSource)?,
                    )?;
                    for field in &info.field_infos {
                        let field_references = field
                            .object_references
                            .len()
                            .checked_add(field.data_references.len())
                            .ok_or(Error::InvalidSource)?;
                        budget.references(field_references)?;
                        budget.wire_work(
                            field
                                .path
                                .path
                                .len()
                                .checked_add(field_references)
                                .ok_or(Error::InvalidSource)?,
                        )?;
                    }
                }
            }
        }
        if let Some(metadata) = &self.metadata {
            budget.references(metadata.externals.len())?;
            budget.wire_work(metadata.externals.len())?;
        }
        let component_capacity = self.source.state.source.components().len().max(1);
        budget.scratch(
            component_capacity
                .checked_mul(size_of::<usize>())
                .ok_or(Error::InvalidSource)?,
        )?;
        ReadBudget::charge_counter(
            LimitKind::Allocations,
            &mut budget.allocations,
            component_capacity,
            budget.max_allocations,
        )?;
        let mut count = 0usize;
        let mut native_components = HashSet::new();
        native_components
            .try_reserve(component_capacity)
            .map_err(|_| Error::Allocation {
                amount: component_capacity,
            })?;
        if !self.removed_objects.is_empty() {
            let scanned_objects = self.source.state.source.components().iter().try_fold(
                0usize,
                |count, component| {
                    let archive = self
                        .archives
                        .get(component.name())
                        .unwrap_or_else(|| component.archive());
                    count
                        .checked_add(archive.objects.len())
                        .ok_or(Error::InvalidSource)
                },
            )?;
            budget.wire_work(
                scanned_objects
                    .checked_mul(self.removed_objects.len())
                    .ok_or(Error::InvalidSource)?,
            )?;
        }
        let mut count_archive = |component_index: usize, archive: &Archive| -> Result<(), Error> {
            for object in &archive.objects {
                if object
                    .archive_info
                    .identifier
                    .is_some_and(|object_identifier| {
                        self.removed_objects.contains(&object_identifier)
                    })
                {
                    continue;
                }
                for info in &object.archive_info.message_infos {
                    if info.object_references.contains(&identifier)
                        || info
                            .field_infos
                            .iter()
                            .any(|field| field.object_references.contains(&identifier))
                    {
                        count = count.checked_add(1).ok_or(Error::InvalidSource)?;
                        native_components.insert(component_index);
                    }
                }
            }
            Ok(())
        };
        for (name, archive) in &self.archives {
            let component_index = self
                .source
                .state
                .source
                .components()
                .iter()
                .position(|component| component.name() == name)
                .ok_or(Error::InvalidSource)?;
            count_archive(component_index, archive)?;
        }
        for (component_index, component) in self.source.state.source.components().iter().enumerate()
        {
            if self.archives.contains_key(component.name()) {
                continue;
            }
            count_archive(component_index, component.archive())?;
        }
        let Some(metadata) = self.metadata.as_ref() else {
            return Ok(count);
        };
        budget.wire_work(metadata.externals.len())?;
        let matching_external_count = metadata
            .externals
            .iter()
            .filter(|edge| edge.object == Some(identifier))
            .count();
        if matching_external_count != 0 && !self.edge_changes.is_empty() {
            let component_scan = metadata.components.len().max(1);
            let pending_edge_work = matching_external_count
                .checked_mul(self.edge_changes.len())
                .and_then(|work| work.checked_mul(component_scan))
                .and_then(|work| work.checked_mul(2))
                .ok_or(Error::InvalidSource)?;
            budget.wire_work(pending_edge_work)?;
        }
        for edge in &metadata.externals {
            if edge.object != Some(identifier) {
                continue;
            }
            // A PackageMetadata edge whose source and target component are
            // identical is never a safe second owner.  The metadata codec
            // cannot publish a same-component removal, while leaving the
            // edge in place after deleting the object would create a dangling
            // owner.  Reject the graph before any copy-on-write decision.
            if edge.source == edge.target {
                return Err(Error::UnsupportedDependency);
            }
            let mut removed_by_pending_change = false;
            for change in &self.edge_changes {
                if !change.remove_old || change.old_identifier != edge.object {
                    continue;
                }
                let source =
                    selector_for_component(metadata, change.source_component, self.source)?;
                let Some(target_component) = change.old_target_component else {
                    continue;
                };
                let target = selector_for_component(metadata, target_component, self.source)?;
                if source.identifier() == edge.source && target.identifier() == edge.target {
                    removed_by_pending_change = true;
                    break;
                }
            }
            if removed_by_pending_change {
                continue;
            }
            let component_index = metadata
                .components
                .iter()
                .position(|component| component.identifier == edge.source)
                .and_then(|metadata_index| {
                    self.source
                        .state
                        .source
                        .components()
                        .iter()
                        .position(|candidate| {
                            component_locator(candidate.name())
                                == Some(metadata.components[metadata_index].locator.as_ref())
                        })
                });
            let Some(component_index) = component_index else {
                return Err(Error::UnsupportedDependency);
            };
            if native_components.insert(component_index) {
                count = count.checked_add(1).ok_or(Error::InvalidSource)?;
            }
        }
        Ok(count)
    }

    fn remove_object(&mut self, identifier: u64, budget: &mut ReadBudget) -> Result<(), Error> {
        if identifier == 0 {
            return Err(Error::InvalidSource);
        }
        budget.wire_work(self.removed_objects.len())?;
        if self.removed_objects.contains(&identifier) {
            return Err(Error::InvalidSource);
        }
        if self.removed_objects.len() == self.removed_objects.capacity() {
            let requested_capacity = self
                .removed_objects
                .len()
                .checked_add(1)
                .ok_or(Error::InvalidSource)?;
            budget.scratch(
                requested_capacity
                    .checked_mul(size_of::<u64>())
                    .ok_or(Error::InvalidSource)?,
            )?;
            ReadBudget::charge_counter(
                LimitKind::Allocations,
                &mut budget.allocations,
                1,
                budget.max_allocations,
            )?;
            self.removed_objects
                .try_reserve_exact(1)
                .map_err(|_| Error::Allocation { amount: 1 })?;
        }
        self.removed_objects.push(identifier);
        let component_name = self
            .find_object(identifier)
            .map(|(name, _)| name)
            .ok_or(Error::InvalidSource)?;
        let component_name = copy_component_name(component_name, budget)?;
        let component_index = self
            .source
            .state
            .source
            .components()
            .iter()
            .position(|component| component.name() == component_name)
            .ok_or(Error::InvalidSource)?;
        let archive = self.load(&component_name, budget)?;
        if archive.remove_object(identifier).is_none() {
            return Err(Error::InvalidSource);
        }
        let Some(metadata) = self.metadata.as_ref() else {
            return Ok(());
        };
        let Some(binding) = metadata
            .uuid_bindings
            .iter()
            .find(|binding| binding.identifier == identifier)
            .copied()
        else {
            return Ok(());
        };
        let metadata_component =
            metadata_component_for_index(metadata, component_index, self.source)?.identifier;
        if metadata_component != binding.component {
            return Err(Error::InvalidSource);
        }
        ReadBudget::charge_counter(
            LimitKind::Allocations,
            &mut budget.allocations,
            size_of::<RemovedIdentity>(),
            budget.max_allocations,
        )?;
        self.removed_identities.push(RemovedIdentity {
            component: component_index,
            identifier,
            uuid: UuidSnapshot::from_parts(binding.uuid.lower(), binding.uuid.upper()),
        });
        Ok(())
    }
}

/// Execute one focused body-table comment operation using the same bounded
/// working archive, UUID registry, copy-on-write, and exact reassembly path as
/// the table comment adapter.  The table adapter owns selector and semantic
/// readback; this bridge owns only the private native graph mutation.
pub(crate) fn execute_table_mutation(
    source: &Package,
    selection: &TableCommentSelection,
    operation: &TableCommentOperation,
    budget: &mut ReadBudget,
) -> Result<TableCommentMutationOutput, Error> {
    if !source.state.source.source_is_exact() {
        return Err(Error::UnsupportedSource);
    }
    // Admit the rooted native graph through the format-neutral engine before
    // the Pages copy-on-write writer allocates or rewrites any archive.  The
    // selector resolver supplies the roots; this adapter owns only the
    // Pages-specific wire mutation and publication.
    let mut roots = [0_u64; 4];
    let mut root_count = 0_usize;
    roots[root_count] = selection.model_identifier;
    root_count += 1;
    roots[root_count] = selection.tile_identifier;
    root_count += 1;
    if let Some(identifier) = selection.list_identifier {
        roots[root_count] = identifier;
        root_count += 1;
    }
    if let Some(identifier) = selection.root_identifier {
        roots[root_count] = identifier;
        root_count += 1;
    }
    let graph = litchi_iwa_table::comments::admit_graph(source, &roots[..root_count])
        .map_err(super::body_table_comments::map_shared_error)?;
    budget.references(graph.references())?;
    budget.wire_work(graph.work())?;
    ReadBudget::charge_counter(
        LimitKind::Allocations,
        &mut budget.allocations,
        graph.allocations(),
        budget.max_allocations,
    )?;
    let mut working = Working::new(source, budget)?;
    let old_root = selection.root_identifier;
    match operation {
        TableCommentOperation::Set(text) => {
            if let Some(root) = old_root {
                if !table_root_is_shared(&working, selection, root, budget)? {
                    rewrite_storage_text(&mut working, root, text, budget)?;
                } else {
                    let clone = clone_storage(&mut working, root, budget)?;
                    rewrite_storage_text(&mut working, clone, text, budget)?;
                    replace_table_root(&mut working, selection, root, clone, budget)?;
                }
            } else {
                let root = create_storage(
                    &mut working,
                    selection.component.as_ref(),
                    text,
                    None,
                    budget,
                )?;
                attach_table_root(&mut working, selection, root, budget)?;
            }
        },
        TableCommentOperation::Clear => {
            let root = old_root.ok_or(Error::CommentNotFound)?;
            let unique = !table_root_is_shared(&working, selection, root, budget)?;
            let root_author = if unique {
                storage_facts(&working, root, budget)?.author
            } else {
                None
            };
            let replies = if unique {
                reply_ids(&working, root, budget)?
            } else {
                Vec::new()
            };
            detach_table_root(&mut working, selection, root, budget)?;
            if unique {
                remove_comment_closure(&mut working, root, replies, root_author, true, budget)?;
            }
        },
        TableCommentOperation::AddReply(text) => {
            let root = old_root.ok_or(Error::CommentNotFound)?;
            let clone = clone_table_root(&mut working, selection, root, budget)?;
            let component = working
                .find_object(root)
                .map(|(name, _)| name.to_owned())
                .ok_or(Error::InvalidSource)?;
            let reply = create_storage(
                &mut working,
                &component,
                text,
                Some((clone.facts.timestamp, clone.facts.author)),
                budget,
            )?;
            rewrite_replies(
                &mut working,
                clone.new,
                CommentStorageReplyRewrite::append(reply),
                budget,
            )?;
            finish_cloned_table_root(&mut working, clone, budget)?;
        },
        TableCommentOperation::SetReply { index, text } => {
            let root = old_root.ok_or(Error::CommentNotFound)?;
            let reply = reply_ids(&working, root, budget)?
                .get(*index)
                .copied()
                .ok_or(Error::ReplyPositionNotFound {
                    position: litchi_core::Position::new(*index),
                })?;
            let clone = clone_table_root(&mut working, selection, root, budget)?;
            // The old root remains live until the transaction has finished,
            // so replacing the reply in place would mutate that old graph as
            // well.  Host Pages clones the selected reply for this path even
            // when the root itself was uniquely owned.
            let target_reply = clone_storage(&mut working, reply, budget)?;
            rewrite_storage_text(&mut working, target_reply, text, budget)?;
            rewrite_replies(
                &mut working,
                clone.new,
                CommentStorageReplyRewrite::replace(*index, reply, target_reply),
                budget,
            )?;
            finish_cloned_table_root(&mut working, clone, budget)?;
        },
        TableCommentOperation::RemoveReply { index } => {
            let root = old_root.ok_or(Error::CommentNotFound)?;
            let reply = reply_ids(&working, root, budget)?
                .get(*index)
                .copied()
                .ok_or(Error::ReplyPositionNotFound {
                    position: litchi_core::Position::new(*index),
                })?;
            let clone = clone_table_root(&mut working, selection, root, budget)?;
            if !clone.shared {
                let component = working
                    .find_object(clone.new)
                    .map(|(name, _)| name.to_owned())
                    .ok_or(Error::InvalidSource)?;
                working.record_edge_from_component(&component, Some(reply), None, budget)?;
            }
            rewrite_replies(
                &mut working,
                clone.new,
                CommentStorageReplyRewrite::remove(*index, reply),
                budget,
            )?;
            finish_cloned_table_root(&mut working, clone, budget)?;
        },
    }
    apply_metadata(&mut working, budget)?;
    let output = serialize_candidate(&mut working, budget)?;
    Ok(TableCommentMutationOutput {
        bytes: output.bytes,
        touched_components: output.touched_components,
        candidate_admission: output.candidate_admission,
    })
}

#[derive(Debug, Clone, Copy)]
enum TableListMutation {
    Append { storage: u64 },
    Remove { key: u32, storage: u64 },
    Decrement { key: u32, expected: u32 },
}

#[derive(Debug, Clone, Copy)]
struct TableListEntryFacts {
    key: u32,
    ref_count: u32,
    storage: u64,
}

type TableListRewrite = (
    Vec<u8>,
    Option<u32>,
    Option<u64>,
    Option<u64>,
    Option<(u32, Option<u64>, Option<u64>)>,
    Option<(u32, Option<u64>, Option<u64>)>,
);

fn attach_table_root(
    working: &mut Working<'_>,
    selection: &TableCommentSelection,
    root: u64,
    budget: &mut ReadBudget,
) -> Result<(), Error> {
    let (list_identifier, key) = if let Some(list_identifier) = selection.list_identifier {
        let key = mutate_table_list(
            working,
            list_identifier,
            TableListMutation::Append { storage: root },
            budget,
        )?;
        (list_identifier, key)
    } else {
        let list_identifier = create_table_comment_list(working, selection, root, budget)?;
        (list_identifier, 2)
    };
    let _ = list_identifier;
    rewrite_table_tile(working, selection, None, Some(key), budget)
}

fn replace_table_root(
    working: &mut Working<'_>,
    selection: &TableCommentSelection,
    old_root: u64,
    new_root: u64,
    budget: &mut ReadBudget,
) -> Result<(), Error> {
    let list_identifier = selection.list_identifier.ok_or(Error::InvalidSource)?;
    let old_key = selection.comment_key.ok_or(Error::InvalidSource)?;
    let entry = table_list_entry(working, list_identifier, old_key)?;
    if entry.storage != old_root {
        return Err(Error::PatchConflict);
    }
    if entry.ref_count == 1 {
        mutate_table_list(
            working,
            list_identifier,
            TableListMutation::Remove {
                key: old_key,
                storage: old_root,
            },
            budget,
        )?;
    } else {
        mutate_table_list(
            working,
            list_identifier,
            TableListMutation::Decrement {
                key: old_key,
                expected: entry.ref_count,
            },
            budget,
        )?;
    }
    let new_key = mutate_table_list(
        working,
        list_identifier,
        TableListMutation::Append { storage: new_root },
        budget,
    )?;
    rewrite_table_tile(working, selection, Some(old_key), Some(new_key), budget)
}

#[derive(Debug)]
struct ClonedTableRoot {
    old: u64,
    new: u64,
    shared: bool,
    facts: StorageFacts,
    replies: Vec<u64>,
}

fn clone_table_root(
    working: &mut Working<'_>,
    selection: &TableCommentSelection,
    root: u64,
    budget: &mut ReadBudget,
) -> Result<ClonedTableRoot, Error> {
    let shared = table_root_is_shared(working, selection, root, budget)?;
    let facts = storage_facts(working, root, budget)?;
    let replies = reply_ids(working, root, budget)?;
    let new = clone_storage(working, root, budget)?;
    replace_table_root(working, selection, root, new, budget)?;
    Ok(ClonedTableRoot {
        old: root,
        new,
        shared,
        facts,
        replies,
    })
}

fn finish_cloned_table_root(
    working: &mut Working<'_>,
    clone: ClonedTableRoot,
    budget: &mut ReadBudget,
) -> Result<(), Error> {
    if clone.shared {
        return Ok(());
    }
    // The cloned root still owns the original root author.  Replies replaced
    // or removed from the clone may lose their last owner, however, so cull
    // their author registrations while reclaiming the old closure.
    remove_comment_closure(working, clone.old, clone.replies, None, true, budget)
}

fn detach_table_root(
    working: &mut Working<'_>,
    selection: &TableCommentSelection,
    root: u64,
    budget: &mut ReadBudget,
) -> Result<(), Error> {
    let list_identifier = selection.list_identifier.ok_or(Error::InvalidSource)?;
    let key = selection.comment_key.ok_or(Error::InvalidSource)?;
    let entry = table_list_entry(working, list_identifier, key)?;
    if entry.storage != root {
        return Err(Error::PatchConflict);
    }
    rewrite_table_tile(working, selection, Some(key), None, budget)?;
    if entry.ref_count == 1 {
        mutate_table_list(
            working,
            list_identifier,
            TableListMutation::Remove { key, storage: root },
            budget,
        )?;
    } else {
        mutate_table_list(
            working,
            list_identifier,
            TableListMutation::Decrement {
                key,
                expected: entry.ref_count,
            },
            budget,
        )?;
    }
    Ok(())
}

fn remove_comment_closure(
    working: &mut Working<'_>,
    root: u64,
    mut pending: Vec<u64>,
    root_author: Option<u64>,
    cull_authors: bool,
    budget: &mut ReadBudget,
) -> Result<(), Error> {
    working.remove_object(root, budget)?;
    while let Some(reply) = pending.pop() {
        if working.owner_count(reply, budget)? != 0 {
            continue;
        }
        let nested = reply_ids(working, reply, budget)?;
        let author = storage_facts(working, reply, budget)?.author;
        working.remove_object(reply, budget)?;
        let nested_bytes = nested
            .len()
            .checked_mul(size_of::<u64>())
            .ok_or(Error::InvalidSource)?;
        budget.scratch(nested_bytes.checked_mul(2).ok_or(Error::InvalidSource)?)?;
        ReadBudget::charge_counter(
            LimitKind::Allocations,
            &mut budget.allocations,
            nested.len(),
            budget.max_allocations,
        )?;
        pending
            .try_reserve_exact(nested.len())
            .map_err(|_| Error::Allocation {
                amount: nested.len(),
            })?;
        pending.extend(nested);
        if cull_authors {
            if let Some(author) = author {
                cull_registered_author(working, author, budget)?;
            }
        }
    }
    if cull_authors {
        if let Some(author) = root_author {
            cull_registered_author(working, author, budget)?;
        }
    }
    Ok(())
}

fn table_list_entry(
    working: &Working<'_>,
    identifier: u64,
    key: u32,
) -> Result<TableListEntryFacts, Error> {
    let (_, object) = working
        .find_object(identifier)
        .ok_or(Error::InvalidSource)?;
    let mut message = None;
    for candidate in &object.messages {
        if candidate.type_ != TABLE_DATA_LIST_MESSAGE_TYPE {
            continue;
        }
        if message.replace(candidate).is_some() {
            return Err(Error::InvalidSource);
        }
    }
    let message = message.ok_or(Error::InvalidSource)?;
    let mut found = None;
    let view = wire::WireView::parse(message.data.as_slice()).map_err(|_| Error::InvalidSource)?;
    for field in view
        .fields()
        .filter(|field| field.number() == TABLE_COMMENT_LIST_ENTRIES_FIELD)
    {
        if field.wire_type() != 2 {
            return Err(Error::InvalidSource);
        }
        let entry = decode_table_list_entry(field.payload())?;
        if entry.key == key && found.replace(entry).is_some() {
            return Err(Error::InvalidSource);
        }
    }
    found.ok_or(Error::InvalidSource)
}

fn table_root_is_shared(
    working: &Working<'_>,
    selection: &TableCommentSelection,
    root: u64,
    budget: &mut ReadBudget,
) -> Result<bool, Error> {
    let owners = working.owner_count(root, budget)?;
    if owners > 1 {
        return Ok(true);
    }
    let list_shared = match (selection.list_identifier, selection.comment_key) {
        (Some(list), Some(key)) => {
            let entry = table_list_entry(working, list, key)?;
            if entry.storage != root {
                return Err(Error::PatchConflict);
            }
            entry.ref_count > 1 || table_list_storage_owner_count(working, list, root, budget)? > 1
        },
        (None, None) => false,
        _ => return Err(Error::InvalidSource),
    };
    Ok(list_shared)
}

fn table_list_storage_owner_count(
    working: &Working<'_>,
    identifier: u64,
    storage: u64,
    budget: &mut ReadBudget,
) -> Result<usize, Error> {
    let (_, object) = working
        .find_object(identifier)
        .ok_or(Error::InvalidSource)?;
    let mut message = None;
    for candidate in &object.messages {
        if candidate.type_ != TABLE_DATA_LIST_MESSAGE_TYPE {
            continue;
        }
        if message.replace(candidate).is_some() {
            return Err(Error::InvalidSource);
        }
    }
    let message = message.ok_or(Error::InvalidSource)?;
    let view = wire::WireView::parse(message.data.as_slice()).map_err(|_| Error::InvalidSource)?;
    let mut owners = 0usize;
    for field in view
        .fields()
        .filter(|field| field.number() == TABLE_COMMENT_LIST_ENTRIES_FIELD)
    {
        budget.wire_fields(1)?;
        budget.wire_work(field.payload().len())?;
        if field.wire_type() != 2 {
            return Err(Error::InvalidSource);
        }
        if decode_table_list_entry(field.payload())?.storage == storage {
            owners = owners.checked_add(1).ok_or(Error::InvalidSource)?;
            budget.references(1)?;
        }
    }
    Ok(owners)
}

fn decode_table_list_entry(data: &[u8]) -> Result<TableListEntryFacts, Error> {
    let view = wire::WireView::parse(data).map_err(|_| Error::InvalidSource)?;
    let mut key = None;
    let mut ref_count = None;
    let mut storage = None;
    for field in view.fields() {
        match field.number() {
            TABLE_COMMENT_ENTRY_KEY_FIELD => {
                if field.wire_type() != 0 || key.is_some() {
                    return Err(Error::InvalidSource);
                }
                key = Some(read_wire_u32(field.payload())?);
            },
            TABLE_COMMENT_ENTRY_REFCOUNT_FIELD => {
                if field.wire_type() != 0 || ref_count.is_some() {
                    return Err(Error::InvalidSource);
                }
                ref_count = Some(read_wire_u32(field.payload())?);
            },
            TABLE_COMMENT_ENTRY_STORAGE_FIELD => {
                if field.wire_type() != 2 || storage.is_some() {
                    return Err(Error::InvalidSource);
                }
                let reference =
                    wire::WireView::parse(field.payload()).map_err(|_| Error::InvalidSource)?;
                let mut identifier = None;
                for nested in reference.fields() {
                    if nested.number() != TABLE_COMMENT_REFERENCE_IDENTIFIER_FIELD
                        || nested.wire_type() != 0
                        || identifier.is_some()
                    {
                        return Err(Error::InvalidSource);
                    }
                    identifier = Some(read_wire_u64(nested.payload())?);
                }
                storage = identifier;
            },
            _ => return Err(Error::UnsupportedDependency),
        }
    }
    let key = key.ok_or(Error::InvalidSource)?;
    let ref_count = ref_count.ok_or(Error::InvalidSource)?;
    let storage = storage.ok_or(Error::InvalidSource)?;
    if key == 0 || ref_count == 0 || storage == 0 {
        return Err(Error::InvalidSource);
    }
    Ok(TableListEntryFacts {
        key,
        ref_count,
        storage,
    })
}

fn mutate_table_list(
    working: &mut Working<'_>,
    identifier: u64,
    mutation: TableListMutation,
    budget: &mut ReadBudget,
) -> Result<u32, Error> {
    if working.owner_count(identifier, budget)? != 1 {
        return Err(Error::UnsupportedDependency);
    }
    let component_name = working
        .find_object(identifier)
        .map(|(name, _)| name.to_owned())
        .ok_or(Error::InvalidSource)?;
    let source = working
        .find_object(identifier)
        .and_then(|(_, object)| {
            let mut message = None;
            for candidate in &object.messages {
                if candidate.type_ != TABLE_DATA_LIST_MESSAGE_TYPE {
                    continue;
                }
                if message.replace(candidate).is_some() {
                    return None;
                }
            }
            message.map(|message| message.data.clone())
        })
        .ok_or(Error::InvalidSource)?;
    let (data, changed_key, old_edge, new_edge, removed_edge, append_edge) =
        rewrite_table_list_payload(&source, mutation)?;
    let archive_limits = working
        .source
        .state
        .source
        .limits()
        .effective_archive_limits()
        .map_err(|_| Error::UnsupportedSource)?;
    let archive = working.load(&component_name, budget)?;
    let object = archive.object_mut(identifier).ok_or(Error::InvalidSource)?;
    let message_index = object
        .messages
        .iter()
        .position(|message| message.type_ == TABLE_DATA_LIST_MESSAGE_TYPE)
        .ok_or(Error::InvalidSource)?;
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(Error::InvalidSource)?
        .clone();
    let before = info.object_references.clone();
    let after = transition_ids(&before, old_edge, new_edge)?;
    let mut field_transitions = Vec::new();
    if let Some((key, old, new)) = removed_edge.or(append_edge) {
        let path = [TABLE_COMMENT_LIST_ENTRIES_FIELD, key];
        for (index, field) in info.field_infos.iter().enumerate() {
            if field.path.path.as_slice() != path {
                continue;
            }
            let field_after = transition_ids(&field.object_references, old, new)?;
            field_transitions.push((index, field_after));
        }
    }
    let transitions = field_transitions
        .iter()
        .map(|(index, after)| FieldObjectReferenceTransition {
            field_info_index: *index,
            expected_path: info.field_infos[*index].path.path.as_slice(),
            before: info.field_infos[*index].object_references.as_slice(),
            after: after.as_slice(),
        })
        .collect::<Vec<_>>();
    object
        .replace_message_transitioning_object_references_preserving_header_with_limits(
            message_index,
            RawMessage {
                type_: TABLE_DATA_LIST_MESSAGE_TYPE,
                data,
            },
            ObjectReferenceTransition {
                aggregate_before: before.as_slice(),
                aggregate_after: after.as_slice(),
                fields: transitions.as_slice(),
            },
            archive_limits,
        )
        .map_err(|_| Error::InvalidSource)?;
    let info = object
        .archive_info
        .message_infos
        .get_mut(message_index)
        .ok_or(Error::InvalidSource)?;
    if let Some((key, _old, _new)) = append_edge {
        if info
            .field_infos
            .iter()
            .any(|field| field.path.path.as_slice() == [TABLE_COMMENT_LIST_ENTRIES_FIELD, key])
        {
            return Err(Error::InvalidSource);
        }
        let mut field = FieldInfo::new(vec![TABLE_COMMENT_LIST_ENTRIES_FIELD, key]);
        field.r#type = Some(FieldType::ObjectReference);
        field.object_references = vec![new_edge.ok_or(Error::InvalidSource)?];
        info.field_infos.push(field);
    }
    if let Some((key, _old, _new)) = removed_edge {
        let path = [TABLE_COMMENT_LIST_ENTRIES_FIELD, key];
        if let Some(index) = info
            .field_infos
            .iter()
            .position(|field| field.path.path.as_slice() == path)
        {
            if !info.field_infos[index].object_references.is_empty() {
                return Err(Error::InvalidSource);
            }
            info.field_infos.remove(index);
        }
    }
    if let Some(old) = old_edge {
        if new_edge.is_none() {
            working.record_edge_from_component(&component_name, Some(old), None, budget)?;
        }
    }
    if let Some(new) = new_edge {
        if old_edge.is_none() {
            working.record_edge_from_component(&component_name, None, Some(new), budget)?;
        } else if old_edge != Some(new) {
            working.record_edge_from_component(&component_name, old_edge, Some(new), budget)?;
        }
    }
    Ok(changed_key.unwrap_or(0))
}

fn rewrite_table_list_payload(
    source: &[u8],
    mutation: TableListMutation,
) -> Result<TableListRewrite, Error> {
    let view = wire::WireView::parse(source).map_err(|_| Error::InvalidSource)?;
    let mut list_type = None;
    let mut next_id_seen = false;
    for field in view.fields() {
        match field.number() {
            TABLE_COMMENT_LIST_TYPE_FIELD => {
                if field.wire_type() != 0 || list_type.is_some() {
                    return Err(Error::InvalidSource);
                }
                list_type = Some(read_wire_u64(field.payload())?);
            },
            TABLE_COMMENT_LIST_NEXT_ID_FIELD => {
                if field.wire_type() != 0 || next_id_seen {
                    return Err(Error::InvalidSource);
                }
                read_wire_u32(field.payload())?;
                next_id_seen = true;
            },
            TABLE_COMMENT_LIST_ENTRIES_FIELD => {
                if field.wire_type() != 2 {
                    return Err(Error::InvalidSource);
                }
                decode_table_list_entry(field.payload())?;
            },
            _ => return Err(Error::UnsupportedDependency),
        }
    }
    let list_type = list_type.ok_or(Error::InvalidSource)?;
    if list_type != TABLE_COMMENT_LIST_TYPE {
        return Err(Error::UnsupportedDependency);
    }
    match mutation {
        TableListMutation::Append { storage } => {
            let next = view
                .fields()
                .filter(|field| field.number() == TABLE_COMMENT_LIST_NEXT_ID_FIELD)
                .map(|field| {
                    if field.wire_type() != 0 {
                        return Err(Error::InvalidSource);
                    }
                    read_wire_u32(field.payload())
                })
                .collect::<Result<Vec<_>, _>>()?;
            if next.len() != 1 {
                return Err(Error::InvalidSource);
            }
            let key = next[0].max(2);
            for field in view
                .fields()
                .filter(|field| field.number() == TABLE_COMMENT_LIST_ENTRIES_FIELD)
            {
                if field.wire_type() != 2 || decode_table_list_entry(field.payload())?.key == key {
                    return Err(Error::InvalidSource);
                }
            }
            let mut reference = Vec::new();
            wire::append_varint_field(
                &mut reference,
                TABLE_COMMENT_REFERENCE_IDENTIFIER_FIELD,
                storage,
            )
            .map_err(|_| Error::InvalidSource)?;
            let mut entry = Vec::new();
            wire::append_varint_field(&mut entry, TABLE_COMMENT_ENTRY_KEY_FIELD, u64::from(key))
                .map_err(|_| Error::InvalidSource)?;
            wire::append_varint_field(&mut entry, TABLE_COMMENT_ENTRY_REFCOUNT_FIELD, 1)
                .map_err(|_| Error::InvalidSource)?;
            wire::append_length_delimited_field(
                &mut entry,
                TABLE_COMMENT_ENTRY_STORAGE_FIELD,
                &reference,
            )
            .map_err(|_| Error::InvalidSource)?;
            let mut appended = source.to_vec();
            wire::append_length_delimited_field(
                &mut appended,
                TABLE_COMMENT_LIST_ENTRIES_FIELD,
                &entry,
            )
            .map_err(|_| Error::InvalidSource)?;
            let data = wire::patch_varint_field(
                &appended,
                TABLE_COMMENT_LIST_NEXT_ID_FIELD,
                true,
                Some(u64::from(key.checked_add(1).ok_or(Error::InvalidSource)?)),
            )
            .map_err(|_| Error::InvalidSource)?;
            Ok((
                data,
                Some(key),
                None,
                Some(storage),
                None,
                Some((key, None, Some(storage))),
            ))
        },
        TableListMutation::Remove { key, storage } => {
            let mut matched = false;
            let data = wire::remove_repeated_length_delimited_field_where(
                source,
                TABLE_COMMENT_LIST_ENTRIES_FIELD,
                |entry| {
                    let facts = decode_table_list_entry(entry).map_err(|_| {
                        litchi_iwa_common::Error::InvalidFormat(
                            "invalid Pages table comment list entry".to_owned(),
                        )
                    })?;
                    if facts.key != key {
                        return Ok(false);
                    }
                    if matched || facts.storage != storage || facts.ref_count != 1 {
                        return Err(litchi_iwa_common::Error::InvalidFormat(
                            "invalid Pages table comment removal".to_owned(),
                        ));
                    }
                    matched = true;
                    Ok(true)
                },
            )
            .map_err(|_| Error::InvalidSource)?;
            if !matched {
                return Err(Error::InvalidSource);
            }
            Ok((
                data,
                Some(key),
                Some(storage),
                None,
                Some((key, Some(storage), None)),
                None,
            ))
        },
        TableListMutation::Decrement { key, expected } => {
            let mut matched = false;
            let data = wire::transform_repeated_length_delimited_fields(
                source,
                TABLE_COMMENT_LIST_ENTRIES_FIELD,
                |entry| {
                    let facts = decode_table_list_entry(entry).map_err(|_| {
                        litchi_iwa_common::Error::InvalidFormat(
                            "invalid Pages table comment list entry".to_owned(),
                        )
                    })?;
                    if facts.key != key {
                        return Ok(entry.to_vec());
                    }
                    if matched || facts.ref_count != expected || facts.ref_count <= 1 {
                        return Err(litchi_iwa_common::Error::InvalidFormat(
                            "invalid Pages table comment refcount".to_owned(),
                        ));
                    }
                    matched = true;
                    wire::patch_varint_field(
                        entry,
                        TABLE_COMMENT_ENTRY_REFCOUNT_FIELD,
                        true,
                        Some(u64::from(expected - 1)),
                    )
                },
            )
            .map_err(|_| Error::InvalidSource)?;
            if !matched {
                return Err(Error::InvalidSource);
            }
            Ok((data, Some(key), None, None, None, None))
        },
    }
}

fn create_table_comment_list(
    working: &mut Working<'_>,
    selection: &TableCommentSelection,
    root: u64,
    budget: &mut ReadBudget,
) -> Result<u64, Error> {
    let identifier = working.reserve_id(budget)?;
    let uuid = working.fresh_uuid(identifier)?;
    let mut reference = Vec::new();
    wire::append_varint_field(
        &mut reference,
        TABLE_COMMENT_REFERENCE_IDENTIFIER_FIELD,
        root,
    )
    .map_err(|_| Error::InvalidSource)?;
    let mut entry = Vec::new();
    wire::append_varint_field(&mut entry, TABLE_COMMENT_ENTRY_KEY_FIELD, 2)
        .map_err(|_| Error::InvalidSource)?;
    wire::append_varint_field(&mut entry, TABLE_COMMENT_ENTRY_REFCOUNT_FIELD, 1)
        .map_err(|_| Error::InvalidSource)?;
    wire::append_length_delimited_field(&mut entry, TABLE_COMMENT_ENTRY_STORAGE_FIELD, &reference)
        .map_err(|_| Error::InvalidSource)?;
    let mut data = Vec::new();
    wire::append_varint_field(
        &mut data,
        TABLE_COMMENT_LIST_TYPE_FIELD,
        TABLE_COMMENT_LIST_TYPE,
    )
    .map_err(|_| Error::InvalidSource)?;
    wire::append_varint_field(&mut data, TABLE_COMMENT_LIST_NEXT_ID_FIELD, 3)
        .map_err(|_| Error::InvalidSource)?;
    wire::append_length_delimited_field(&mut data, TABLE_COMMENT_LIST_ENTRIES_FIELD, &entry)
        .map_err(|_| Error::InvalidSource)?;
    let limits = working
        .source
        .state
        .source
        .limits()
        .effective_archive_limits()
        .map_err(|_| Error::UnsupportedSource)?;
    let mut object = ArchiveObject::new_with_limits(
        identifier,
        vec![RawMessage {
            type_: TABLE_DATA_LIST_MESSAGE_TYPE,
            data,
        }],
        limits,
    )
    .map_err(|_| Error::InvalidSource)?;
    let info = object
        .archive_info
        .message_infos
        .get_mut(0)
        .ok_or(Error::InvalidSource)?;
    info.object_references.push(root);
    let mut field = FieldInfo::new(vec![TABLE_COMMENT_LIST_ENTRIES_FIELD, 2]);
    field.r#type = Some(FieldType::ObjectReference);
    field.object_references.push(root);
    info.field_infos.push(field);
    let source_model = working
        .find_object(selection.model_identifier)
        .ok_or(Error::InvalidSource)?
        .1
        .messages
        .get(selection.model_message_index)
        .filter(|message| message.type_ == selection.model_message_type)
        .ok_or(Error::InvalidSource)?
        .data
        .clone();
    let archive = working.load(selection.component.as_ref(), budget)?;
    attach_table_comment_list_to_model(archive, selection, &source_model, identifier, limits)?;
    archive
        .insert_object_with_limits(object, limits)
        .map_err(|_| Error::InvalidSource)?;
    let component = working
        .source
        .state
        .source
        .components()
        .iter()
        .position(|candidate| candidate.name() == selection.component.as_ref())
        .ok_or(Error::InvalidSource)?;
    working.register_identity(component, identifier, uuid);
    Ok(identifier)
}

fn attach_table_comment_list_to_model(
    archive: &mut Archive,
    selection: &TableCommentSelection,
    source_model: &[u8],
    identifier: u64,
    limits: litchi_iwa_core::Limits,
) -> Result<(), Error> {
    let max_message = source_model.len().saturating_add(256).max(1);
    let options = table_storage::DecodeOptions::new(
        max_message,
        max_message.saturating_mul(2),
        max_message.saturating_mul(32),
        64,
        limits.max_metadata_items().max(1),
        max_message,
    );
    let prepared = table_storage::prepare_table_model_comment_storage_table_rewrite(
        source_model,
        table_storage::CommentStorageTableReferenceEdit::insert(identifier),
        options,
    )
    .map_err(|_| Error::InvalidSource)?;
    let requirements = prepared.requirements();
    let data = prepared
        .execute(requirements.exact_limits())
        .map_err(|_| Error::InvalidSource)?
        .0;
    let object = archive
        .object_mut(selection.model_identifier)
        .ok_or(Error::InvalidSource)?;
    let info = object
        .archive_info
        .message_infos
        .get(selection.model_message_index)
        .ok_or(Error::InvalidSource)?
        .clone();
    if info.object_references.contains(&identifier)
        || info
            .field_infos
            .iter()
            .any(|field| field.path.path.as_slice() == TABLE_MODEL_COMMENT_STORAGE_FIELD_PATH)
    {
        return Err(Error::InvalidSource);
    }
    let before = info.object_references.clone();
    let mut after = before.clone();
    after.push(identifier);
    let unchanged = info
        .field_infos
        .iter()
        .enumerate()
        .map(|(index, field)| FieldObjectReferenceTransition {
            field_info_index: index,
            expected_path: field.path.path.as_slice(),
            before: field.object_references.as_slice(),
            after: field.object_references.as_slice(),
        })
        .collect::<Vec<_>>();
    object
        .replace_message_transitioning_object_references_preserving_header_with_limits(
            selection.model_message_index,
            RawMessage {
                type_: selection.model_message_type,
                data,
            },
            ObjectReferenceTransition {
                aggregate_before: before.as_slice(),
                aggregate_after: after.as_slice(),
                fields: unchanged.as_slice(),
            },
            limits,
        )
        .map_err(|_| Error::InvalidSource)?;
    let info = object
        .archive_info
        .message_infos
        .get_mut(selection.model_message_index)
        .ok_or(Error::InvalidSource)?;
    let mut field = FieldInfo::new(TABLE_MODEL_COMMENT_STORAGE_FIELD_PATH.to_vec());
    field.r#type = Some(FieldType::ObjectReference);
    field.object_references = vec![identifier];
    info.field_infos.push(field);
    Ok(())
}

fn rewrite_table_tile(
    working: &mut Working<'_>,
    selection: &TableCommentSelection,
    expected: Option<u32>,
    replacement: Option<u32>,
    budget: &mut ReadBudget,
) -> Result<(), Error> {
    if working.owner_count(selection.tile_identifier, budget)? != 1 {
        return Err(Error::UnsupportedDependency);
    }
    let archive_limits = working
        .source
        .state
        .source
        .limits()
        .effective_archive_limits()
        .map_err(|_| Error::UnsupportedSource)?;
    let archive = working.load(selection.tile_component.as_ref(), budget)?;
    let object = archive
        .object(selection.tile_identifier)
        .ok_or(Error::InvalidSource)?;
    let message = object
        .messages
        .get(selection.tile_message_index)
        .filter(|message| message.type_ == selection.tile_message_type)
        .ok_or(Error::InvalidSource)?
        .data
        .clone();
    let data = rewrite_tile_comment(
        &message,
        selection.tile_row,
        selection.column,
        selection.table_columns,
        expected,
        replacement,
    )?;
    archive
        .object_mut(selection.tile_identifier)
        .ok_or(Error::InvalidSource)?
        .replace_message_preserving_header_with_limits(
            selection.tile_message_index,
            RawMessage {
                type_: selection.tile_message_type,
                data,
            },
            archive_limits,
        )
        .map(|_| ())
        .map_err(|_| Error::InvalidSource)
}

fn rewrite_tile_comment(
    source: &[u8],
    row_index: u32,
    column: u32,
    table_columns: u32,
    expected: Option<u32>,
    replacement: Option<u32>,
) -> Result<Vec<u8>, Error> {
    let mut matched = false;
    let data =
        wire::transform_repeated_length_delimited_fields(source, TABLE_TILE_ROWS_FIELD, |row| {
            let view = wire::WireView::parse(row).map_err(|_| {
                litchi_iwa_common::Error::InvalidFormat("invalid Pages table tile row".to_owned())
            })?;
            let row_value = view
                .fields()
                .filter(|field| field.number() == TABLE_TILE_ROW_INDEX_FIELD)
                .map(|field| {
                    if field.wire_type() != 0 {
                        return Err(litchi_iwa_common::Error::InvalidFormat(
                            "invalid Pages table row index".to_owned(),
                        ));
                    }
                    decode_varint_from_bytes(field.payload())
                        .map(|(value, used)| {
                            if used != field.payload().len() {
                                Err(litchi_iwa_common::Error::InvalidFormat(
                                    "noncanonical Pages table row index".to_owned(),
                                ))
                            } else {
                                Ok(value)
                            }
                        })
                        .map_err(|_| {
                            litchi_iwa_common::Error::InvalidFormat(
                                "invalid Pages table row index".to_owned(),
                            )
                        })?
                })
                .collect::<Result<Vec<_>, _>>()?;
            if row_value.len() != 1 || row_value[0] != u64::from(row_index) {
                return Ok(row.to_vec());
            }
            if matched {
                return Err(litchi_iwa_common::Error::InvalidFormat(
                    "duplicate Pages table tile row".to_owned(),
                ));
            }
            matched = true;
            rewrite_table_tile_row(row, column, table_columns, expected, replacement).map_err(
                |_| litchi_iwa_common::Error::InvalidFormat("invalid Pages table cell".to_owned()),
            )
        })
        .map_err(|_| Error::InvalidSource)?;
    if matched {
        return Ok(data);
    }
    if expected.is_some() || replacement.is_none() {
        return Err(Error::PatchConflict);
    }
    let row = new_table_tile_row(row_index, column, table_columns, replacement)?;
    let mut data = wire::patch_varint_field(
        source,
        4,
        true,
        Some(u64::from(row_index.saturating_add(1))),
    )
    .map_err(|_| Error::InvalidSource)?;
    wire::append_length_delimited_field(&mut data, TABLE_TILE_ROWS_FIELD, &row)
        .map_err(|_| Error::InvalidSource)?;
    Ok(data)
}

fn new_table_tile_row(
    row_index: u32,
    column: u32,
    table_columns: u32,
    replacement: Option<u32>,
) -> Result<Vec<u8>, Error> {
    let replacement = replacement.ok_or(Error::InvalidSource)?;
    if column >= table_columns {
        return Err(Error::InvalidSource);
    }
    let columns = usize::try_from(table_columns).map_err(|_| Error::InvalidSource)?;
    let column = usize::try_from(column).map_err(|_| Error::InvalidSource)?;
    let offset_bytes = columns.checked_mul(2).ok_or(Error::InvalidSource)?;
    let mut offsets = vec![0xff_u8; offset_bytes];
    offsets[column * 2..column * 2 + 2].copy_from_slice(&0_u16.to_le_bytes());
    let mut cell = litchi_numbers_wire::BncCell::minimal();
    cell.set_comment_identifier(Some(replacement));
    let cell = cell
        .try_encode_with_limit(usize::MAX)
        .map_err(|_| Error::InvalidSource)?;
    let mut row = Vec::new();
    wire::append_varint_field(&mut row, TABLE_TILE_ROW_INDEX_FIELD, u64::from(row_index))
        .map_err(|_| Error::InvalidSource)?;
    wire::append_varint_field(&mut row, TABLE_TILE_ROW_CELL_COUNT_FIELD, 1)
        .map_err(|_| Error::InvalidSource)?;
    wire::append_length_delimited_field(&mut row, 3, &[]).map_err(|_| Error::InvalidSource)?;
    wire::append_length_delimited_field(&mut row, 4, &[]).map_err(|_| Error::InvalidSource)?;
    wire::append_varint_field(&mut row, 5, 5).map_err(|_| Error::InvalidSource)?;
    wire::append_length_delimited_field(&mut row, TABLE_TILE_ROW_CELL_STORAGE_FIELD, &cell)
        .map_err(|_| Error::InvalidSource)?;
    wire::append_length_delimited_field(&mut row, TABLE_TILE_ROW_CELL_OFFSETS_FIELD, &offsets)
        .map_err(|_| Error::InvalidSource)?;
    wire::append_varint_field(&mut row, TABLE_TILE_ROW_WIDE_OFFSETS_FIELD, 0)
        .map_err(|_| Error::InvalidSource)?;
    Ok(row)
}

fn rewrite_table_tile_row(
    source: &[u8],
    column: u32,
    table_columns: u32,
    expected: Option<u32>,
    replacement: Option<u32>,
) -> Result<Vec<u8>, Error> {
    let view = wire::WireView::parse(source).map_err(|_| Error::InvalidSource)?;
    let storage = unique_length_field(&view, TABLE_TILE_ROW_CELL_STORAGE_FIELD)?;
    let offsets = unique_length_field(&view, TABLE_TILE_ROW_CELL_OFFSETS_FIELD)?;
    let wide = view
        .fields()
        .filter(|field| field.number() == TABLE_TILE_ROW_WIDE_OFFSETS_FIELD)
        .map(|field| {
            if field.wire_type() != 0 {
                return Err(Error::InvalidSource);
            }
            Ok(read_wire_u64(field.payload())? != 0)
        })
        .collect::<Result<Vec<_>, _>>()?;
    if wide.len() > 1 {
        return Err(Error::InvalidSource);
    }
    let wide = wide.first().copied().unwrap_or(false);
    let unit = if wide { 4usize } else { 1usize };
    let raw_offsets = offsets.payload();
    if !raw_offsets.len().is_multiple_of(2) {
        return Err(Error::InvalidSource);
    }
    let column = usize::try_from(column).map_err(|_| Error::InvalidSource)?;
    let table_columns = usize::try_from(table_columns).map_err(|_| Error::InvalidSource)?;
    if column >= table_columns {
        return Err(Error::InvalidSource);
    }
    let storage_bytes = storage.payload();
    let mut cells = Vec::new();
    cells
        .try_reserve_exact(raw_offsets.len() / 2)
        .map_err(|_| Error::Allocation {
            amount: raw_offsets.len() / 2,
        })?;
    for chunk in raw_offsets.chunks_exact(2) {
        let raw = u16::from_le_bytes([chunk[0], chunk[1]]);
        if raw == u16::MAX {
            cells.push(None);
            continue;
        }
        let start = usize::from(raw)
            .checked_mul(unit)
            .ok_or(Error::InvalidSource)?;
        if start > storage_bytes.len() {
            return Err(Error::InvalidSource);
        }
        cells.push(Some(start));
    }
    let mut slices = Vec::new();
    if cells
        .iter()
        .enumerate()
        .any(|(index, start)| index >= table_columns && start.is_some())
    {
        return Err(Error::InvalidSource);
    }
    if cells.len() < table_columns {
        cells.resize(table_columns, None);
    }
    let declared_cell_count = unique_varint_field(&view, TABLE_TILE_ROW_CELL_COUNT_FIELD)?;
    let present_cell_count = cells.iter().filter(|cell| cell.is_some()).count();
    if usize::try_from(declared_cell_count).map_err(|_| Error::InvalidSource)? != present_cell_count
    {
        return Err(Error::InvalidSource);
    }
    for (index, start) in cells.iter().enumerate() {
        let Some(start) = start else {
            slices.push(None);
            continue;
        };
        let end = cells
            .iter()
            .skip(index + 1)
            .flatten()
            .next()
            .copied()
            .unwrap_or(storage_bytes.len());
        if *start >= end {
            return Err(Error::InvalidSource);
        }
        slices.push(Some(storage_bytes[*start..end].to_vec()));
    }
    let previous = slices.get(column).and_then(Option::as_deref);
    let previous_comment = previous
        .map(|bytes| {
            litchi_numbers_wire::BncCell::parse(bytes)
                .map(|cell| cell.comment_identifier())
                .map_err(|_| Error::InvalidSource)
        })
        .transpose()?
        .flatten();
    if previous_comment != expected {
        return Err(Error::PatchConflict);
    }
    match (previous, replacement) {
        (None, None) => return Err(Error::InvalidSource),
        (None, Some(identifier)) => {
            let mut cell = litchi_numbers_wire::BncCell::minimal();
            cell.set_comment_identifier(Some(identifier));
            slices[column] = Some(
                cell.try_encode_with_limit(usize::MAX)
                    .map_err(|_| Error::InvalidSource)?,
            );
        },
        (Some(bytes), replacement) => {
            let mut cell =
                litchi_numbers_wire::BncCell::parse(bytes).map_err(|_| Error::InvalidSource)?;
            cell.set_comment_identifier(replacement);
            slices[column] = Some(
                cell.try_encode_with_limit(usize::MAX)
                    .map_err(|_| Error::InvalidSource)?,
            );
        },
    }
    let mut output_storage = Vec::new();
    let mut output_offsets = Vec::with_capacity(raw_offsets.len());
    let cell_count = slices.iter().filter(|cell| cell.is_some()).count();
    for cell in slices {
        let Some(cell) = cell else {
            output_offsets.extend_from_slice(&u16::MAX.to_le_bytes());
            continue;
        };
        let offset = output_storage
            .len()
            .checked_div(unit)
            .ok_or(Error::InvalidSource)?;
        if offset >= usize::from(u16::MAX) || output_storage.len() % unit != 0 {
            return Err(Error::InvalidSource);
        }
        output_offsets.extend_from_slice(
            &u16::try_from(offset)
                .map_err(|_| Error::InvalidSource)?
                .to_le_bytes(),
        );
        output_storage.extend_from_slice(&cell);
        let remainder = output_storage.len() % unit;
        if remainder != 0 {
            output_storage.resize(
                output_storage
                    .len()
                    .checked_add(unit - remainder)
                    .ok_or(Error::InvalidSource)?,
                0,
            );
        }
    }
    let mut output = wire::patch_length_delimited_field(
        source,
        TABLE_TILE_ROW_CELL_STORAGE_FIELD,
        true,
        Some(output_storage.as_slice()),
    )
    .map_err(|_| Error::InvalidSource)?;
    output = wire::patch_length_delimited_field(
        &output,
        TABLE_TILE_ROW_CELL_OFFSETS_FIELD,
        true,
        Some(output_offsets.as_slice()),
    )
    .map_err(|_| Error::InvalidSource)?;
    wire::patch_varint_field(
        &output,
        TABLE_TILE_ROW_CELL_COUNT_FIELD,
        true,
        Some(u64::try_from(cell_count).map_err(|_| Error::InvalidSource)?),
    )
    .map_err(|_| Error::InvalidSource)
}

fn unique_length_field<'a>(
    view: &wire::WireView<'a>,
    number: u32,
) -> Result<wire::WireFieldView<'a>, Error> {
    let mut found = None;
    for field in view.fields().filter(|field| field.number() == number) {
        if field.wire_type() != 2 || found.replace(field).is_some() {
            return Err(Error::InvalidSource);
        }
    }
    found.ok_or(Error::InvalidSource)
}

fn unique_varint_field(view: &wire::WireView<'_>, number: u32) -> Result<u64, Error> {
    let mut found = None;
    for field in view.fields().filter(|field| field.number() == number) {
        if field.wire_type() != 0 || found.is_some() {
            return Err(Error::InvalidSource);
        }
        found = Some(read_wire_u64(field.payload())?);
    }
    found.ok_or(Error::InvalidSource)
}

fn read_wire_u64(data: &[u8]) -> Result<u64, Error> {
    let (value, used) = decode_varint_from_bytes(data).map_err(|_| Error::InvalidSource)?;
    if used != data.len() {
        return Err(Error::InvalidSource);
    }
    Ok(value)
}

fn read_wire_u32(data: &[u8]) -> Result<u32, Error> {
    u32::try_from(read_wire_u64(data)?).map_err(|_| Error::InvalidSource)
}

fn storage_facts(
    working: &Working<'_>,
    identifier: u64,
    budget: &mut ReadBudget,
) -> Result<StorageFacts, Error> {
    let (_component_name, object) = working
        .find_object(identifier)
        .ok_or(Error::InvalidSource)?;
    let mut message_index = None;
    for (index, message) in object.messages.iter().enumerate() {
        if message.type_ == COMMENT_STORAGE_MESSAGE_TYPE {
            if message_index.replace(index).is_some() {
                return Err(Error::InvalidSource);
            }
        }
    }
    let message_index = message_index.ok_or(Error::InvalidSource)?;
    let message = object
        .messages
        .get(message_index)
        .ok_or(Error::InvalidSource)?;
    let options = budget.codec_options(message.data.len())?;
    let mut replies = ReplyIds::new(budget);
    let decoded = comment_storage_codec::decode_comment_storage_archive_with_visitor(
        &message.data,
        options,
        &mut replies,
    );
    let (reply_ids, allocation_failed, budget_error) = replies.into_parts();
    let (snapshot, report) = decoded.map_err(|error| map_comment_decode_error(error, budget))?;
    if let Some(error) = budget_error {
        return Err(error);
    }
    if let Some(amount) = allocation_failed {
        return Err(Error::Allocation { amount });
    }
    budget.charge_comment_report(report)?;
    if report.replies() != reply_ids.len() || reply_ids.contains(&0) {
        return Err(Error::InvalidSource);
    }
    Ok(StorageFacts {
        author: snapshot.author().map(|reference| reference.identifier()),
        timestamp: snapshot.creation_date(),
    })
}

fn reply_ids(
    working: &Working<'_>,
    identifier: u64,
    budget: &mut ReadBudget,
) -> Result<Vec<u64>, Error> {
    let (component, object) = working
        .find_object(identifier)
        .ok_or(Error::InvalidSource)?;
    let message = object
        .messages
        .iter()
        .find(|message| message.type_ == COMMENT_STORAGE_MESSAGE_TYPE)
        .ok_or(Error::InvalidSource)?;
    let options = budget.codec_options(message.data.len())?;
    let mut collector = ReplyIds::new(budget);
    let decoded = comment_storage_codec::decode_comment_storage_archive_with_visitor(
        &message.data,
        options,
        &mut collector,
    );
    let (ids, allocation_failed, budget_error) = collector.into_parts();
    let (_snapshot, report) = decoded.map_err(|error| map_comment_decode_error(error, budget))?;
    if let Some(error) = budget_error {
        return Err(error);
    }
    if let Some(amount) = allocation_failed {
        return Err(Error::Allocation { amount });
    }
    budget.charge_comment_report(report)?;
    if report.replies() != ids.len() {
        return Err(Error::InvalidSource);
    }
    let _ = component;
    Ok(ids)
}

fn clone_storage(
    working: &mut Working<'_>,
    identifier: u64,
    budget: &mut ReadBudget,
) -> Result<u64, Error> {
    let (component_name, source_object) = {
        let (name, object) = working
            .find_object(identifier)
            .ok_or(Error::InvalidSource)?;
        if object.messages.len() != 1 || object.messages[0].type_ != COMMENT_STORAGE_MESSAGE_TYPE {
            return Err(Error::InvalidSource);
        }
        let component_name = copy_component_name(name, budget)?;
        charge_archive_object_clone_admission(object, budget)?;
        (component_name, object.clone())
    };
    let new_identifier = working.reserve_id(budget)?;
    let limits = working
        .source
        .state
        .source
        .limits()
        .effective_archive_limits()
        .map_err(|_| Error::UnsupportedSource)?;
    charge_archive_object_clone_admission(&source_object, budget)?;
    let clone = source_object
        .clone_with_identity_remap_with_limits(new_identifier, &[], &source_object.messages, limits)
        .map_err(|_| Error::InvalidSource)?;
    let archive = working.load(&component_name, budget)?;
    budget.scratch(size_of::<ArchiveObject>())?;
    archive
        .insert_object_with_limits(clone, limits)
        .map_err(|_| Error::InvalidSource)?;
    let uuid = working.fresh_uuid(new_identifier)?;
    rewrite_storage_uuid(working, new_identifier, uuid, budget)?;
    let component = working
        .source
        .state
        .source
        .components()
        .iter()
        .position(|candidate| candidate.name() == component_name)
        .ok_or(Error::InvalidSource)?;
    ReadBudget::charge_counter(
        LimitKind::Allocations,
        &mut budget.allocations,
        1,
        budget.max_allocations,
    )?;
    working.register_identity(component, new_identifier, uuid);
    Ok(new_identifier)
}

fn create_storage(
    working: &mut Working<'_>,
    component_name: &str,
    text: &str,
    inherited: Option<(Option<DateSnapshot>, Option<u64>)>,
    budget: &mut ReadBudget,
) -> Result<u64, Error> {
    if text.len() > MAX_COMMENT_TEXT_BYTES {
        return Err(Error::LimitExceeded {
            kind: LimitKind::TextBytes,
            observed: text.len() as u64,
            maximum: MAX_COMMENT_TEXT_BYTES as u64,
        });
    }
    let component = working
        .source
        .state
        .source
        .components()
        .iter()
        .position(|candidate| candidate.name() == component_name)
        .ok_or(Error::InvalidSource)?;
    let identifier = working.reserve_id(budget)?;
    let uuid = working.fresh_uuid(identifier)?;
    let timestamp = inherited
        .and_then(|(timestamp, _)| timestamp)
        .unwrap_or_else(|| DateSnapshot::from_bits(42.5_f64.to_bits()));
    let author = inherited
        .and_then(|(_, author)| author)
        .or(find_registered_author(working, budget)?.map(|facts| facts.author_identifier));
    let write = match author {
        Some(author) => {
            comment_storage_codec::CommentStorageLeafWrite::new(text, timestamp, author, uuid)
        },
        None => {
            comment_storage_codec::CommentStorageLeafWrite::without_author(text, timestamp, uuid)
        },
    };
    let options = budget.codec_options(text.len().saturating_add(128))?;
    let prepared = comment_storage_codec::prepare_comment_storage_leaf_write(write, options)
        .map_err(|error| map_comment_decode_error(error, budget))?;
    let requirements = prepared.execution_requirements();
    charge_comment_requirements(budget, requirements)?;
    let payload = prepared
        .execute(requirements.exact())
        .map_err(|error| map_comment_decode_error(error, budget))?
        .into_bytes();
    let limits = working
        .source
        .state
        .source
        .limits()
        .effective_archive_limits()
        .map_err(|_| Error::UnsupportedSource)?;
    let mut object = ArchiveObject::new_with_limits(
        identifier,
        vec![RawMessage {
            type_: COMMENT_STORAGE_MESSAGE_TYPE,
            data: payload,
        }],
        limits,
    )
    .map_err(|_| Error::InvalidSource)?;
    if let Some(author) = author {
        object
            .archive_info
            .message_infos
            .get_mut(0)
            .ok_or(Error::InvalidSource)?
            .object_references
            .push(author);
    }
    charge_archive_object_clone_admission(&object, budget)?;
    let archive = working.load(component_name, budget)?;
    budget.scratch(size_of::<ArchiveObject>())?;
    archive
        .insert_object_with_limits(object, limits)
        .map_err(|_| Error::InvalidSource)?;
    ReadBudget::charge_counter(
        LimitKind::Allocations,
        &mut budget.allocations,
        1,
        budget.max_allocations,
    )?;
    working.register_identity(component, identifier, uuid);
    if let Some(author) = author {
        working.record_edge_from_component(component_name, None, Some(author), budget)?;
    }
    Ok(identifier)
}

/// Resolve the source-registered annotation author used for newly-created
/// comment storage.  Pages packages may omit the registry entirely; in that
/// legacy shape creation remains authorless, while a present registry is
/// validated strictly and reused without manufacturing a native identity.
fn find_registered_author(
    working: &Working<'_>,
    budget: &mut ReadBudget,
) -> Result<Option<AuthorRegistryFacts>, Error> {
    let object_bound = working.source.state.object_count;
    let limits = working
        .source
        .state
        .source
        .limits()
        .effective_archive_limits()
        .map_err(|_| Error::UnsupportedSource)?;
    let message_bound = object_bound
        .checked_mul(limits.max_messages_per_object())
        .ok_or(Error::InvalidSource)?;
    budget.wire_work(
        object_bound
            .checked_mul(size_of::<ArchiveObject>())
            .and_then(|bytes| {
                bytes.checked_add(message_bound.checked_mul(size_of::<RawMessage>())?)
            })
            .ok_or(Error::InvalidSource)?,
    )?;
    let mut registry = None;
    for component in working.source.state.source.components().iter() {
        for object in &component.archive().objects {
            let Some(message_index) = object
                .messages
                .iter()
                .position(|message| message.type_ == ANNOTATION_AUTHOR_STORAGE_MESSAGE_TYPE)
            else {
                continue;
            };
            if object
                .messages
                .iter()
                .filter(|message| message.type_ == ANNOTATION_AUTHOR_STORAGE_MESSAGE_TYPE)
                .count()
                != 1
            {
                return Err(Error::InvalidSource);
            }
            if registry.is_some() {
                return Err(Error::InvalidSource);
            }
            let payload = object
                .messages
                .get(message_index)
                .ok_or(Error::InvalidSource)?
                .data
                .as_slice();
            budget.message(payload.len())?;
            let (snapshot, report) =
                annotation_author_codec::decode_annotation_author_storage_with_report(
                    payload,
                    budget.annotation_options(payload.len())?,
                )
                .map_err(|error| map_author_decode_error(error, budget))?;
            budget.charge_author_report(report)?;
            let mut selected = None;
            for (ordinal, reference) in snapshot.author_refs().enumerate() {
                if reference.identifier() == 0
                    || reference.deprecated_type().is_some()
                    || reference.deprecated_is_external().is_some()
                {
                    return Err(Error::InvalidSource);
                }
                if selected.is_none() {
                    selected = Some((reference.identifier(), ordinal));
                }
                let (_author_component, author_object) = working
                    .find_object(reference.identifier())
                    .ok_or(Error::InvalidSource)?;
                let author_message = author_object
                    .messages
                    .iter()
                    .filter(|message| message.type_ == ANNOTATION_AUTHOR_MESSAGE_TYPE)
                    .count();
                if author_message != 1 {
                    return Err(Error::InvalidSource);
                }
            }
            let (author_identifier, author_ordinal) = selected.unwrap_or((0, 0));
            registry = Some(AuthorRegistryFacts {
                object_identifier: object.archive_info.identifier.ok_or(Error::InvalidSource)?,
                message_index,
                author_identifier,
                author_ordinal,
                author_count: snapshot.author_ref_count(),
            });
        }
    }
    Ok(registry.filter(|facts| facts.author_identifier != 0))
}

fn storage_author_in_object(
    object: &ArchiveObject,
    budget: &mut ReadBudget,
) -> Result<Option<u64>, Error> {
    let mut message = None;
    for candidate in &object.messages {
        if candidate.type_ != COMMENT_STORAGE_MESSAGE_TYPE {
            continue;
        }
        if message.replace(candidate).is_some() {
            return Err(Error::InvalidSource);
        }
    }
    let Some(message) = message else {
        return Ok(None);
    };
    budget.message(message.data.len())?;
    let mut visitor = AuthorOnlyVisitor;
    let (snapshot, report) = comment_storage_codec::decode_comment_storage_archive_with_visitor(
        message.data.as_slice(),
        budget.codec_options(message.data.len())?,
        &mut visitor,
    )
    .map_err(|error| map_comment_decode_error(error, budget))?;
    budget.charge_comment_report(report)?;
    Ok(snapshot.author().map(|author| author.identifier()))
}

fn author_is_used(
    working: &Working<'_>,
    author_identifier: u64,
    budget: &mut ReadBudget,
) -> Result<bool, Error> {
    let object_bound = working.source.state.object_count;
    budget.wire_work(
        object_bound
            .checked_mul(size_of::<ArchiveObject>())
            .ok_or(Error::InvalidSource)?,
    )?;
    if !working.removed_objects.is_empty() {
        let scanned_objects = working
            .archives
            .values()
            .try_fold(0usize, |count, archive| {
                count
                    .checked_add(archive.objects.len())
                    .ok_or(Error::InvalidSource)
            })?
            .checked_add(
                working
                    .source
                    .state
                    .source
                    .components()
                    .iter()
                    .filter(|component| !working.archives.contains_key(component.name()))
                    .try_fold(0usize, |count, component| {
                        count
                            .checked_add(component.archive().objects.len())
                            .ok_or(Error::InvalidSource)
                    })?,
            )
            .ok_or(Error::InvalidSource)?;
        budget.wire_work(
            scanned_objects
                .checked_mul(working.removed_objects.len())
                .ok_or(Error::InvalidSource)?,
        )?;
    }
    for archive in working.archives.values() {
        for object in &archive.objects {
            if object
                .archive_info
                .identifier
                .is_some_and(|identifier| working.removed_objects.contains(&identifier))
            {
                continue;
            }
            if storage_author_in_object(object, budget)? == Some(author_identifier) {
                return Ok(true);
            }
        }
    }
    for component in working.source.state.source.components().iter() {
        if working.archives.contains_key(component.name()) {
            continue;
        }
        for object in &component.archive().objects {
            if object
                .archive_info
                .identifier
                .is_some_and(|identifier| working.removed_objects.contains(&identifier))
            {
                continue;
            }
            if storage_author_in_object(object, budget)? == Some(author_identifier) {
                return Ok(true);
            }
        }
    }
    Ok(false)
}

fn charge_author_storage_requirements(
    budget: &mut ReadBudget,
    requirements: annotation_author_codec::RewriteExecutionRequirements,
) -> Result<(), Error> {
    budget.wire_input(
        requirements.input_bytes(),
        requirements.max_depth() as usize,
    )?;
    budget.wire_fields(requirements.fields())?;
    budget.wire_work(requirements.work_bytes())?;
    budget.observe_codec_depth(requirements.max_depth())?;
    budget.references(requirements.references())?;
    budget.scratch(
        requirements
            .scratch_bytes()
            .checked_add(requirements.retained_bytes())
            .ok_or(Error::InvalidSource)?,
    )?;
    ReadBudget::charge_counter(
        LimitKind::Allocations,
        &mut budget.allocations,
        requirements.allocations(),
        budget.max_allocations,
    )
}

fn cull_registered_author(
    working: &mut Working<'_>,
    author_identifier: u64,
    budget: &mut ReadBudget,
) -> Result<(), Error> {
    let Some(registry) = find_registered_author(working, budget)? else {
        return Ok(());
    };
    if registry.author_identifier != author_identifier || registry.author_count != 1 {
        return Ok(());
    }
    if author_is_used(working, author_identifier, budget)? {
        return Ok(());
    }
    if working.metadata.as_ref().is_some_and(|metadata| {
        metadata
            .externals
            .iter()
            .any(|edge| edge.object == Some(author_identifier))
    }) {
        return Err(Error::UnsupportedDependency);
    }
    let (component_name, object) = working
        .find_object(registry.object_identifier)
        .ok_or(Error::InvalidSource)?;
    let component_name = copy_component_name(component_name, budget)?;
    let source = object
        .messages
        .get(registry.message_index)
        .ok_or(Error::InvalidSource)?
        .data
        .clone();
    budget.scratch(source.len())?;
    let prepared = annotation_author_codec::prepare_annotation_author_storage_rewrite(
        source.as_slice(),
        annotation_author_codec::AnnotationAuthorStorageRewrite::remove(
            registry.author_ordinal,
            author_identifier,
        ),
        budget.annotation_options(source.len())?,
    )
    .map_err(|_| Error::InvalidSource)?;
    let requirements = prepared.requirements();
    charge_author_storage_requirements(budget, requirements)?;
    let candidate = prepared
        .execute(requirements.exact())
        .map_err(|_| Error::InvalidSource)?
        .0;
    let archive_limits = working
        .source
        .state
        .source
        .limits()
        .effective_archive_limits()
        .map_err(|_| Error::UnsupportedSource)?;
    let archive = working.load(&component_name, budget)?;
    let object = archive
        .object_mut(registry.object_identifier)
        .ok_or(Error::InvalidSource)?;
    let (aggregate_before, field_before) = {
        let info = object
            .archive_info
            .message_infos
            .get(registry.message_index)
            .ok_or(Error::InvalidSource)?;
        let fields = info
            .field_infos
            .iter()
            .map(|field| (field.path.path.clone(), field.object_references.clone()))
            .collect::<Vec<_>>();
        (info.object_references.clone(), fields)
    };
    let mut aggregate_after = aggregate_before.clone();
    if let Some(index) = aggregate_after
        .iter()
        .position(|identifier| *identifier == author_identifier)
    {
        aggregate_after.remove(index);
    }
    let mut field_after = Vec::new();
    for (index, (path, before)) in field_before.iter().enumerate() {
        let mut after = before.clone();
        if let Some(position) = after
            .iter()
            .position(|identifier| *identifier == author_identifier)
        {
            after.remove(position);
            field_after.push((index, path.clone(), after));
        }
    }
    if aggregate_after != aggregate_before || !field_after.is_empty() {
        let field_transitions = field_after
            .iter()
            .map(|(index, path, after)| FieldObjectReferenceTransition {
                field_info_index: *index,
                expected_path: path.as_slice(),
                before: field_before[*index].1.as_slice(),
                after: after.as_slice(),
            })
            .collect::<Vec<_>>();
        let transition = ObjectReferenceTransition {
            aggregate_before: aggregate_before.as_slice(),
            aggregate_after: aggregate_after.as_slice(),
            fields: field_transitions.as_slice(),
        };
        object
            .replace_message_transitioning_object_references_preserving_header_with_limits(
                registry.message_index,
                RawMessage {
                    type_: ANNOTATION_AUTHOR_STORAGE_MESSAGE_TYPE,
                    data: candidate,
                },
                transition,
                archive_limits,
            )
            .map_err(|_| Error::InvalidSource)?;
    } else {
        object
            .replace_message_preserving_header_with_limits(
                registry.message_index,
                RawMessage {
                    type_: ANNOTATION_AUTHOR_STORAGE_MESSAGE_TYPE,
                    data: candidate,
                },
                archive_limits,
            )
            .map_err(|_| Error::InvalidSource)?;
    }
    working.remove_object(author_identifier, budget)
}

fn rewrite_storage_uuid(
    working: &mut Working<'_>,
    identifier: u64,
    uuid: UuidSnapshot,
    budget: &mut ReadBudget,
) -> Result<(), Error> {
    let component_name = working
        .find_object(identifier)
        .map(|(name, _)| name)
        .ok_or(Error::InvalidSource)?;
    let component_name = copy_component_name(component_name, budget)?;
    let limits = wire_limits(working.source)?;
    let archive_limits = working
        .source
        .state
        .source
        .limits()
        .effective_archive_limits()
        .map_err(|_| Error::UnsupportedSource)?;
    let archive = working.load(&component_name, budget)?;
    let object = archive.object_mut(identifier).ok_or(Error::InvalidSource)?;
    let message_index = object
        .messages
        .iter()
        .position(|message| message.type_ == COMMENT_STORAGE_MESSAGE_TYPE)
        .ok_or(Error::InvalidSource)?;
    let source = object
        .messages
        .get(message_index)
        .ok_or(Error::InvalidSource)?
        .data
        .as_slice();
    budget.scratch(source.len())?;
    ReadBudget::charge_counter(
        LimitKind::Allocations,
        &mut budget.allocations,
        1,
        budget.max_allocations,
    )?;
    let source = source.to_vec();
    let edits = [
        NestedFieldEdit::new(
            &[COMMENT_UUID_FIELD, UUID_LOWER_FIELD],
            true,
            NestedFieldReplacement::Varint(Some(uuid.lower())),
        ),
        NestedFieldEdit::new(
            &[COMMENT_UUID_FIELD, UUID_UPPER_FIELD],
            true,
            NestedFieldReplacement::Varint(Some(uuid.upper())),
        ),
    ];
    let data = patch_wire(&source, &edits, limits, budget)?;
    object
        .replace_message_preserving_header_with_limits(
            message_index,
            RawMessage {
                type_: COMMENT_STORAGE_MESSAGE_TYPE,
                data,
            },
            archive_limits,
        )
        .map_err(|_| Error::InvalidSource)?;
    Ok(())
}

fn rewrite_storage_text(
    working: &mut Working<'_>,
    identifier: u64,
    text: &str,
    budget: &mut ReadBudget,
) -> Result<(), Error> {
    let (component_name, object) = working
        .find_object(identifier)
        .ok_or(Error::InvalidSource)?;
    let component_name = component_name.to_owned();
    let message = object
        .messages
        .iter()
        .find(|message| message.type_ == COMMENT_STORAGE_MESSAGE_TYPE)
        .ok_or(Error::InvalidSource)?;
    budget.scratch(message.data.len())?;
    ReadBudget::charge_counter(
        LimitKind::Allocations,
        &mut budget.allocations,
        1,
        budget.max_allocations,
    )?;
    let source = message.data.to_vec();
    let archive_limits = working
        .source
        .state
        .source
        .limits()
        .effective_archive_limits()
        .map_err(|_| Error::UnsupportedSource)?;
    if text.len() > MAX_COMMENT_TEXT_BYTES {
        return Err(Error::LimitExceeded {
            kind: LimitKind::TextBytes,
            observed: text.len() as u64,
            maximum: MAX_COMMENT_TEXT_BYTES as u64,
        });
    }
    let limits = wire_limits(working.source)?;
    let data = patch_wire(
        &source,
        &[NestedFieldEdit::new(
            &[COMMENT_TEXT_FIELD],
            true,
            NestedFieldReplacement::LengthDelimited(Some(text.as_bytes())),
        )],
        limits,
        budget,
    )?;
    let archive = working.load(&component_name, budget)?;
    let index = archive
        .object(identifier)
        .ok_or(Error::InvalidSource)?
        .messages
        .iter()
        .position(|message| message.type_ == COMMENT_STORAGE_MESSAGE_TYPE)
        .ok_or(Error::InvalidSource)?;
    archive
        .object_mut(identifier)
        .ok_or(Error::InvalidSource)?
        .replace_message_preserving_header_with_limits(
            index,
            RawMessage {
                type_: COMMENT_STORAGE_MESSAGE_TYPE,
                data,
            },
            archive_limits,
        )
        .map_err(|_| Error::InvalidSource)?;
    Ok(())
}

#[derive(Debug, Clone, Copy)]
struct ReplyReferenceStagingPlan {
    field_count: usize,
    changed_fields: usize,
    aggregate_after_len: usize,
}

fn reply_rewrite_ids(operation: CommentStorageReplyRewrite) -> (Option<u64>, Option<u64>) {
    match operation {
        CommentStorageReplyRewrite::Append { identifier } => (None, Some(identifier)),
        CommentStorageReplyRewrite::Replace {
            expected_identifier,
            replacement_identifier,
            ..
        } => (Some(expected_identifier), Some(replacement_identifier)),
        CommentStorageReplyRewrite::Remove {
            expected_identifier,
            ..
        } => (Some(expected_identifier), None),
    }
}

fn reply_aggregate_after_len(
    before: &[u64],
    removed: Option<u64>,
    added: Option<u64>,
) -> Result<usize, Error> {
    let removed_index = removed.map(|identifier| {
        before
            .iter()
            .position(|value| *value == identifier)
            .ok_or(Error::InvalidSource)
    });
    let removed_index = match removed_index {
        Some(index) => Some(index?),
        None => None,
    };
    let mut after_len = before.len();
    if removed_index.is_some() {
        after_len = after_len.checked_sub(1).ok_or(Error::InvalidSource)?;
    }
    if let Some(identifier) = added {
        if identifier == 0
            || before
                .iter()
                .enumerate()
                .any(|(index, value)| Some(index) != removed_index && *value == identifier)
        {
            return Err(Error::InvalidSource);
        }
        after_len = after_len.checked_add(1).ok_or(Error::InvalidSource)?;
    }
    Ok(after_len)
}

fn reply_field_after_len(
    before: &[u64],
    removed: Option<u64>,
    added: Option<u64>,
) -> Result<(usize, bool), Error> {
    let mut after_len = 0usize;
    let mut replaced = false;
    let mut changed = false;
    for value in before {
        if removed == Some(*value) {
            if added.is_some() && !replaced {
                after_len = after_len.checked_add(1).ok_or(Error::InvalidSource)?;
                changed = added != Some(*value);
                replaced = true;
            } else {
                changed = true;
            }
        } else {
            after_len = after_len.checked_add(1).ok_or(Error::InvalidSource)?;
        }
    }
    if removed.is_none()
        && let Some(identifier) = added
        && before.iter().any(|value| *value != 0)
        && !before.contains(&identifier)
    {
        after_len = after_len.checked_add(1).ok_or(Error::InvalidSource)?;
        changed = true;
    }
    Ok((after_len, changed))
}

fn admit_reply_reference_staging(
    info: &MessageInfo,
    operation: CommentStorageReplyRewrite,
    budget: &mut ReadBudget,
) -> Result<ReplyReferenceStagingPlan, Error> {
    let (removed, added) = reply_rewrite_ids(operation);
    let aggregate_after_len =
        reply_aggregate_after_len(info.object_references.as_slice(), removed, added)?;
    let mut metadata_references = info
        .object_references
        .len()
        .checked_add(info.data_references.len())
        .ok_or(Error::InvalidSource)?;
    let mut path_components = 0usize;
    let mut field_references = 0usize;
    let mut changed_fields = 0usize;
    let mut changed_path_components = 0usize;
    let mut changed_after_references = 0usize;
    let mut field_before_allocations = 0usize;
    let mut changed_field_allocations = 0usize;
    for field in &info.field_infos {
        path_components = path_components
            .checked_add(field.path.path.len())
            .ok_or(Error::InvalidSource)?;
        field_references = field_references
            .checked_add(field.object_references.len())
            .ok_or(Error::InvalidSource)?;
        metadata_references = metadata_references
            .checked_add(
                field
                    .object_references
                    .len()
                    .checked_add(field.data_references.len())
                    .ok_or(Error::InvalidSource)?,
            )
            .ok_or(Error::InvalidSource)?;
        field_before_allocations = field_before_allocations
            .checked_add(usize::from(!field.path.path.is_empty()))
            .and_then(|count| count.checked_add(usize::from(!field.object_references.is_empty())))
            .ok_or(Error::InvalidSource)?;
        let (after_len, changed) =
            reply_field_after_len(field.object_references.as_slice(), removed, added)?;
        if !changed {
            continue;
        }
        changed_fields = changed_fields.checked_add(1).ok_or(Error::InvalidSource)?;
        changed_path_components = changed_path_components
            .checked_add(field.path.path.len())
            .ok_or(Error::InvalidSource)?;
        changed_after_references = changed_after_references
            .checked_add(after_len)
            .ok_or(Error::InvalidSource)?;
        changed_field_allocations = changed_field_allocations
            .checked_add(usize::from(!field.path.path.is_empty()))
            .and_then(|count| count.checked_add(usize::from(after_len != 0)))
            .ok_or(Error::InvalidSource)?;
    }
    let field_count = info.field_infos.len();
    let field_before_bytes = field_count
        .checked_mul(size_of::<(Vec<u32>, Vec<u64>)>())
        .and_then(|bytes| {
            path_components
                .checked_mul(size_of::<u32>())
                .and_then(|path_bytes| bytes.checked_add(path_bytes))
        })
        .and_then(|bytes| {
            field_references
                .checked_mul(size_of::<u64>())
                .and_then(|reference_bytes| bytes.checked_add(reference_bytes))
        })
        .ok_or(Error::InvalidSource)?;
    let changed_field_bytes = changed_fields
        .checked_mul(size_of::<(usize, Vec<u32>, Vec<u64>)>())
        .and_then(|bytes| {
            changed_path_components
                .checked_mul(size_of::<u32>())
                .and_then(|path_bytes| bytes.checked_add(path_bytes))
        })
        .and_then(|bytes| {
            changed_after_references
                .checked_mul(size_of::<u64>())
                .and_then(|reference_bytes| bytes.checked_add(reference_bytes))
        })
        .ok_or(Error::InvalidSource)?;
    let scratch_bytes = info
        .object_references
        .len()
        .checked_mul(size_of::<u64>())
        .and_then(|bytes| bytes.checked_add(field_before_bytes))
        .and_then(|bytes| {
            aggregate_after_len
                .checked_mul(size_of::<u64>())
                .and_then(|aggregate_bytes| bytes.checked_add(aggregate_bytes))
        })
        .and_then(|bytes| bytes.checked_add(changed_field_bytes))
        .and_then(|bytes| {
            changed_fields
                .checked_mul(size_of::<FieldObjectReferenceTransition<'_>>())
                .and_then(|transition_bytes| bytes.checked_add(transition_bytes))
        })
        .ok_or(Error::InvalidSource)?;
    let allocations = usize::from(!info.object_references.is_empty())
        .checked_add(usize::from(field_count != 0))
        .and_then(|count| count.checked_add(field_before_allocations))
        .and_then(|count| count.checked_add(usize::from(aggregate_after_len != 0)))
        .and_then(|count| count.checked_add(usize::from(changed_fields != 0)))
        .and_then(|count| count.checked_add(changed_field_allocations))
        .and_then(|count| count.checked_add(usize::from(changed_fields != 0)))
        .ok_or(Error::InvalidSource)?;
    budget.metadata_scan(
        metadata_references,
        field_count.checked_add(1).ok_or(Error::InvalidSource)?,
        path_components,
    )?;
    budget.wire_work(
        aggregate_after_len
            .checked_add(changed_fields)
            .and_then(|work| work.checked_add(changed_after_references))
            .and_then(|work| work.checked_add(changed_path_components))
            .ok_or(Error::InvalidSource)?,
    )?;
    budget.scratch(scratch_bytes)?;
    ReadBudget::charge_counter(
        LimitKind::Allocations,
        &mut budget.allocations,
        allocations,
        budget.max_allocations,
    )?;
    Ok(ReplyReferenceStagingPlan {
        field_count,
        changed_fields,
        aggregate_after_len,
    })
}

fn copy_component_name(name: &str, budget: &mut ReadBudget) -> Result<String, Error> {
    budget.copy_text(name.len())?;
    let mut copied = String::new();
    copied
        .try_reserve_exact(name.len())
        .map_err(|_| Error::Allocation { amount: name.len() })?;
    copied.push_str(name);
    Ok(copied)
}

fn clone_reply_reference_metadata(
    info: &MessageInfo,
) -> Result<(Vec<u64>, Vec<(Vec<u32>, Vec<u64>)>), Error> {
    let mut aggregate_before = Vec::new();
    aggregate_before
        .try_reserve_exact(info.object_references.len())
        .map_err(|_| Error::Allocation {
            amount: info.object_references.len(),
        })?;
    aggregate_before.extend_from_slice(&info.object_references);

    let mut field_before = Vec::new();
    field_before
        .try_reserve_exact(info.field_infos.len())
        .map_err(|_| Error::Allocation {
            amount: info.field_infos.len(),
        })?;
    for field in &info.field_infos {
        let mut path = Vec::new();
        path.try_reserve_exact(field.path.path.len())
            .map_err(|_| Error::Allocation {
                amount: field.path.path.len(),
            })?;
        path.extend_from_slice(&field.path.path);
        let mut references = Vec::new();
        references
            .try_reserve_exact(field.object_references.len())
            .map_err(|_| Error::Allocation {
                amount: field.object_references.len(),
            })?;
        references.extend_from_slice(&field.object_references);
        field_before.push((path, references));
    }
    Ok((aggregate_before, field_before))
}

fn build_reply_aggregate(
    before: &[u64],
    operation: CommentStorageReplyRewrite,
    expected_len: usize,
) -> Result<Vec<u64>, Error> {
    let (removed, added) = reply_rewrite_ids(operation);
    let mut after = Vec::new();
    after
        .try_reserve_exact(expected_len)
        .map_err(|_| Error::Allocation {
            amount: expected_len,
        })?;
    let mut removed_once = false;
    for value in before {
        if removed == Some(*value) && !removed_once {
            removed_once = true;
        } else {
            after.push(*value);
        }
    }
    if let Some(added) = added {
        after.push(added);
    }
    if after.len() != expected_len {
        return Err(Error::InvalidSource);
    }
    Ok(after)
}

fn build_reply_fields(
    field_before: &[(Vec<u32>, Vec<u64>)],
    operation: CommentStorageReplyRewrite,
    expected_count: usize,
) -> Result<Vec<(usize, Vec<u32>, Vec<u64>)>, Error> {
    let (removed, added) = reply_rewrite_ids(operation);
    let mut fields = Vec::new();
    fields
        .try_reserve_exact(expected_count)
        .map_err(|_| Error::Allocation {
            amount: expected_count,
        })?;
    for (index, (path, before)) in field_before.iter().enumerate() {
        let (after_len, changed) = reply_field_after_len(before, removed, added)?;
        if !changed {
            continue;
        }
        let mut after = Vec::new();
        after
            .try_reserve_exact(after_len)
            .map_err(|_| Error::Allocation { amount: after_len })?;
        let mut replaced = false;
        for value in before {
            if removed == Some(*value) {
                if let Some(added) = added
                    && !replaced
                {
                    // Keep a route-local witness aligned with the ordered
                    // reply occurrence, including a one-element FieldInfo.
                    after.push(added);
                    replaced = true;
                }
            } else {
                after.push(*value);
            }
        }
        if removed.is_none()
            && let Some(added) = added
            && before.iter().any(|value| *value != 0)
            && !before.contains(&added)
        {
            after.push(added);
        }
        if after.len() != after_len {
            return Err(Error::InvalidSource);
        }
        let mut path_copy = Vec::new();
        path_copy
            .try_reserve_exact(path.len())
            .map_err(|_| Error::Allocation { amount: path.len() })?;
        path_copy.extend_from_slice(path);
        fields.push((index, path_copy, after));
    }
    if fields.len() != expected_count {
        return Err(Error::InvalidSource);
    }
    Ok(fields)
}

fn rewrite_replies(
    working: &mut Working<'_>,
    identifier: u64,
    operation: CommentStorageReplyRewrite,
    budget: &mut ReadBudget,
) -> Result<(), Error> {
    let (component_name, object) = working
        .find_object(identifier)
        .ok_or(Error::InvalidSource)?;
    let component_name = copy_component_name(component_name, budget)?;
    let message_index = object
        .messages
        .iter()
        .position(|message| message.type_ == COMMENT_STORAGE_MESSAGE_TYPE)
        .ok_or(Error::InvalidSource)?;
    let source = object
        .messages
        .get(message_index)
        .ok_or(Error::InvalidSource)?
        .data
        .as_slice();
    budget.scratch(source.len())?;
    ReadBudget::charge_counter(
        LimitKind::Allocations,
        &mut budget.allocations,
        1,
        budget.max_allocations,
    )?;
    let source = source.to_vec();
    let options = budget.codec_options(source.len())?;
    let prepared =
        comment_storage_codec::prepare_comment_storage_reply_rewrite(&source, operation, options)
            .map_err(|error| map_comment_decode_error(error, budget))?;
    let requirements = prepared.execution_requirements();
    charge_comment_requirements(budget, requirements)?;
    let data = prepared
        .execute(requirements.exact())
        .map_err(|error| map_comment_decode_error(error, budget))?
        .into_bytes();
    let archive_limits = working
        .source
        .state
        .source
        .limits()
        .effective_archive_limits()
        .map_err(|_| Error::UnsupportedSource)?;
    let archive = working.load(&component_name, budget)?;
    let info = archive
        .object(identifier)
        .ok_or(Error::InvalidSource)?
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(Error::InvalidSource)?;
    let plan = admit_reply_reference_staging(info, operation, budget)?;
    let (aggregate_before, field_before) = clone_reply_reference_metadata(info)?;
    if field_before.len() != plan.field_count {
        return Err(Error::InvalidSource);
    }
    let aggregate_after =
        build_reply_aggregate(&aggregate_before, operation, plan.aggregate_after_len)?;
    let fields = build_reply_fields(&field_before, operation, plan.changed_fields)?;
    let mut field_transitions = Vec::new();
    field_transitions
        .try_reserve_exact(fields.len())
        .map_err(|_| Error::Allocation {
            amount: fields.len(),
        })?;
    for (index, path, after) in &fields {
        field_transitions.push(FieldObjectReferenceTransition {
            field_info_index: *index,
            expected_path: path.as_slice(),
            before: field_before[*index].1.as_slice(),
            after: after.as_slice(),
        });
    }
    let transition = ObjectReferenceTransition {
        aggregate_before: aggregate_before.as_slice(),
        aggregate_after: aggregate_after.as_slice(),
        fields: field_transitions.as_slice(),
    };
    let object = archive.object_mut(identifier).ok_or(Error::InvalidSource)?;
    object
        .replace_message_transitioning_object_references_preserving_header_with_limits(
            message_index,
            RawMessage {
                type_: COMMENT_STORAGE_MESSAGE_TYPE,
                data,
            },
            transition,
            archive_limits,
        )
        .map_err(|_| Error::InvalidSource)?;
    Ok(())
}

#[allow(dead_code)]
fn rewrite_drawable(
    working: &mut Working<'_>,
    selection: &SelectionFacts,
    old: Option<u64>,
    new: Option<u64>,
    budget: &mut ReadBudget,
) -> Result<(), Error> {
    working.record_edge(selection, old, new, budget)?;
    let wire_limits = working
        .source
        .state
        .source
        .limits()
        .effective_archive_limits()
        .map_err(|_| Error::UnsupportedSource)
        .and_then(|_| wire_limits(working.source))?;
    let archive_limits = working
        .source
        .state
        .source
        .limits()
        .effective_archive_limits()
        .map_err(|_| Error::UnsupportedSource)?;
    let component_name = selection.component.as_ref();
    let archive = working.load(component_name, budget)?;
    let object = archive
        .object_mut(selection.identity.get())
        .ok_or(Error::InvalidSource)?;
    let message = object
        .messages
        .get(selection.message_index)
        .ok_or(Error::InvalidSource)?;
    if message.type_ != selection.message_type {
        return Err(Error::PatchConflict);
    }
    budget.scratch(message.data.len())?;
    ReadBudget::charge_counter(
        LimitKind::Allocations,
        &mut budget.allocations,
        1,
        budget.max_allocations,
    )?;
    let source_data = message.data.to_vec();
    let leaf_path = append_path(&selection.route, REFERENCE_IDENTIFIER_FIELD)?;
    let data = match (old, new) {
        (Some(_), Some(identifier)) => patch_wire(
            &source_data,
            &[NestedFieldEdit::new(
                leaf_path.as_slice(),
                true,
                NestedFieldReplacement::Varint(Some(identifier)),
            )],
            wire_limits,
            budget,
        )?,
        (Some(_), None) => patch_wire(
            &source_data,
            &[NestedFieldEdit::new(
                &selection.route,
                true,
                NestedFieldReplacement::LengthDelimited(None),
            )],
            wire_limits,
            budget,
        )?,
        (None, Some(identifier)) => {
            let reference = patch_wire(
                &[],
                &[NestedFieldEdit::new(
                    &[REFERENCE_IDENTIFIER_FIELD],
                    false,
                    NestedFieldReplacement::Varint(Some(identifier)),
                )],
                wire_limits,
                budget,
            )?;
            insert_nested_reference(
                &source_data,
                &selection.route,
                &reference,
                wire_limits,
                budget,
            )?
        },
        (None, None) => return Err(Error::InvalidSource),
    };
    let (aggregate_before, field_before) = {
        let info = object
            .archive_info
            .message_infos
            .get(selection.message_index)
            .ok_or(Error::InvalidSource)?;
        let fields = info
            .field_infos
            .iter()
            .map(|field| (field.path.path.clone(), field.object_references.clone()))
            .collect::<Vec<_>>();
        (info.object_references.clone(), fields)
    };
    let aggregate_after = transition_ids(&aggregate_before, old, new)?;
    let mut fields_owned = Vec::new();
    for (index, (path, before)) in field_before.iter().enumerate() {
        if path.as_slice() != selection.route {
            continue;
        }
        let after = transition_ids(before, old, new)?;
        if after != *before {
            fields_owned.push((index, path.clone(), after));
        }
    }
    let field_transitions = fields_owned
        .iter()
        .map(|(index, path, after)| FieldObjectReferenceTransition {
            field_info_index: *index,
            expected_path: path.as_slice(),
            before: field_before[*index].1.as_slice(),
            after: after.as_slice(),
        })
        .collect::<Vec<_>>();
    let transition = ObjectReferenceTransition {
        aggregate_before: aggregate_before.as_slice(),
        aggregate_after: aggregate_after.as_slice(),
        fields: field_transitions.as_slice(),
    };
    object
        .replace_message_transitioning_object_references_preserving_header_with_limits(
            selection.message_index,
            RawMessage {
                type_: selection.message_type,
                data,
            },
            transition,
            archive_limits,
        )
        .map_err(|_| Error::InvalidSource)?;
    Ok(())
}

fn transition_ids(before: &[u64], old: Option<u64>, new: Option<u64>) -> Result<Vec<u64>, Error> {
    let mut after = before.to_vec();
    if let Some(old) = old {
        let Some(index) = after.iter().position(|value| *value == old) else {
            return Err(Error::InvalidSource);
        };
        after.remove(index);
    }
    if let Some(new) = new {
        if new == 0 || after.contains(&new) {
            return Err(Error::InvalidSource);
        }
        after.push(new);
    }
    Ok(after)
}

#[allow(dead_code)]
fn append_path(path: &[u32], leaf: u32) -> Result<Vec<u32>, Error> {
    let mut result = Vec::new();
    result
        .try_reserve_exact(path.len().saturating_add(1))
        .map_err(|_| Error::Allocation {
            amount: path.len() + 1,
        })?;
    result.extend_from_slice(path);
    result.push(leaf);
    Ok(result)
}

#[allow(dead_code)]
fn insert_nested_reference(
    source: &[u8],
    path: &[u32],
    reference: &[u8],
    limits: WireLimits,
    budget: &mut ReadBudget,
) -> Result<Vec<u8>, Error> {
    let (&root, descendants) = path.split_first().ok_or(Error::InvalidSource)?;
    let view =
        wire::WireView::parse_with_limits(source, limits).map_err(|_| Error::InvalidSource)?;
    let mut selected = None;
    for field in view.fields().filter(|field| field.number() == root) {
        field
            .validate_canonical_framing()
            .map_err(|_| Error::InvalidSource)?;
        if selected.replace(field).is_some() || field.wire_type() != 2 {
            return Err(Error::InvalidSource);
        }
    }
    let nested = if let Some(field) = selected {
        if descendants.is_empty() {
            return Err(Error::InvalidSource);
        }
        let child =
            insert_nested_reference(field.payload(), descendants, reference, limits, budget)?;
        return patch_wire(
            source,
            &[NestedFieldEdit::new(
                &[root],
                true,
                NestedFieldReplacement::LengthDelimited(Some(child.as_slice())),
            )],
            limits,
            budget,
        );
    } else {
        let mut nested = reference.to_vec();
        for &field in descendants.iter().rev() {
            let path = [field];
            let edit = NestedFieldEdit::new(
                &path,
                false,
                NestedFieldReplacement::LengthDelimited(Some(nested.as_slice())),
            );
            nested = patch_wire(&[], &[edit], limits, budget)?;
        }
        nested
    };
    patch_wire(
        source,
        &[NestedFieldEdit::new(
            &[root],
            false,
            NestedFieldReplacement::LengthDelimited(Some(nested.as_slice())),
        )],
        limits,
        budget,
    )
}

fn patch_wire(
    source: &[u8],
    edits: &[NestedFieldEdit<'_>],
    limits: WireLimits,
    budget: &mut ReadBudget,
) -> Result<Vec<u8>, Error> {
    let output_bound = source
        .len()
        .checked_add(edits.len().saturating_mul(256))
        .ok_or(Error::InvalidSource)?;
    budget.wire_input(source.len(), edits.len().saturating_add(1))?;
    budget.wire_fields(edits.len())?;
    budget.wire_work(source.len().saturating_add(output_bound))?;
    budget.scratch(output_bound)?;
    let output = wire::patch_nested_fields_batched_with_limits(source, edits, limits)
        .map_err(|_| Error::InvalidSource)?;
    if output.len() > output_bound {
        return Err(Error::InvalidSource);
    }
    Ok(output)
}

fn charge_comment_requirements(
    budget: &mut ReadBudget,
    requirements: comment_storage_codec::RewriteExecutionRequirements,
) -> Result<(), Error> {
    budget.wire_input(requirements.input_bytes, requirements.max_depth as usize)?;
    budget.wire_fields(requirements.fields)?;
    budget.wire_work(requirements.work_bytes)?;
    budget.observe_codec_depth(requirements.max_depth)?;
    budget.references(requirements.references)?;
    ReadBudget::charge_counter(
        LimitKind::Replies,
        &mut budget.replies,
        requirements.replies,
        budget.max_replies,
    )?;
    budget.scratch(requirements.scratch_bytes)?;
    ReadBudget::charge_counter(
        LimitKind::Allocations,
        &mut budget.allocations,
        requirements.allocations,
        budget.max_allocations,
    )
}

fn read_metadata_state(
    source: &Package,
    budget: &mut ReadBudget,
) -> Result<Option<MetadataState>, Error> {
    let Some(component) = source.state.source.components().get("Index/Metadata.iwa") else {
        return Ok(None);
    };
    let mut location = None;
    for (object_index, object) in component.archive().objects.iter().enumerate() {
        for (message_index, message) in object.messages.iter().enumerate() {
            if message.type_ == METADATA_MESSAGE_TYPE {
                if location.replace((object_index, message_index)).is_some() {
                    return Err(Error::InvalidSource);
                }
            }
        }
    }
    let (object_index, message_index) = location.ok_or(Error::InvalidSource)?;
    let payload = component
        .archive()
        .objects
        .get(object_index)
        .and_then(|object| object.messages.get(message_index))
        .ok_or(Error::InvalidSource)?
        .data
        .as_slice();
    let archive_limits = source
        .state
        .source
        .limits()
        .effective_archive_limits()
        .map_err(|_| Error::UnsupportedSource)?;
    let archive_bound = charge_archive_sizing_scan(component.archive(), archive_limits, budget)?;
    budget.scratch(archive_bound)?;
    ReadBudget::charge_counter(
        LimitKind::Allocations,
        &mut budget.allocations,
        1,
        budget.max_allocations,
    )?;
    let archive_encoded = component
        .archive()
        .encoded_len_with_limits(archive_limits)
        .map_err(|_| Error::InvalidSource)?;
    let _ = archive_encoded;
    budget.scratch(payload.len())?;
    ReadBudget::charge_counter(
        LimitKind::Allocations,
        &mut budget.allocations,
        2,
        budget.max_allocations,
    )?;
    let payload = payload.to_vec();
    let input = payload.len().max(1);
    let archive = source
        .state
        .source
        .limits()
        .effective_archive_limits()
        .map_err(|_| Error::UnsupportedSource)?;
    let metadata_items = archive
        .max_metadata_items()
        .saturating_mul(source.state.source.components().len().max(1))
        .max(1);
    let options = MetadataRewriteOptions::new(
        input,
        input.saturating_mul(2).max(input),
        input.saturating_mul(16).max(1),
        input.saturating_mul(128).max(1),
        64,
        metadata_items,
        source
            .state
            .source
            .components()
            .len()
            .saturating_mul(64)
            .max(1),
        source
            .state
            .source
            .components()
            .len()
            .saturating_mul(16)
            .max(1),
    );
    let mut visitor = MetadataVisitor::new(budget);
    let inspection =
        metadata_codec::inspect_package_metadata_with_visitor(&payload, options, &mut visitor)
            .map_err(map_metadata_error)?;
    // The visitor owns the mutable ledger only while the codec invokes its
    // callbacks.  Move its retained projections out before charging the
    // report and the UUID index below.
    let MetadataVisitor {
        components,
        uuids: metadata_uuids,
        uuid_bindings,
        externals,
        budget_error,
        ..
    } = visitor;
    if let Some(error) = budget_error {
        return Err(error);
    }
    let report = inspection.report();
    budget.wire_input(report.input_bytes(), report.max_depth() as usize)?;
    budget.wire_fields(report.fields())?;
    budget.wire_work(
        report
            .work_bytes()
            .checked_add(report.components_scanned())
            .ok_or(Error::InvalidSource)?,
    )?;
    budget.observe_codec_depth(report.max_depth())?;
    budget.references(report.references_scanned())?;
    ReadBudget::charge_counter(
        LimitKind::Allocations,
        &mut budget.allocations,
        report.allocations(),
        budget.max_allocations,
    )?;
    let mut uuids = HashSet::new();
    budget.scratch(
        metadata_uuids
            .len()
            .checked_mul(size_of::<(u64, u64)>() * 2)
            .ok_or(Error::InvalidSource)?,
    )?;
    ReadBudget::charge_counter(
        LimitKind::Allocations,
        &mut budget.allocations,
        1,
        budget.max_allocations,
    )?;
    uuids
        .try_reserve(metadata_uuids.len())
        .map_err(|_| Error::Allocation {
            amount: metadata_uuids.len(),
        })?;
    for uuid in metadata_uuids {
        uuids.insert((uuid.uuid.lower(), uuid.uuid.upper()));
    }
    budget.scratch(
        component
            .archive()
            .objects
            .len()
            .saturating_mul(size_of::<ArchiveObject>()),
    )?;
    budget.copy_text(component.name().len())?;
    charge_archive_clone_admission(component.archive(), budget)?;
    Ok(Some(MetadataState {
        member: component.name().into(),
        archive: component.archive().clone(),
        object_index,
        message_index,
        payload,
        last_identifier: inspection.last_object_identifier(),
        components,
        uuids,
        uuid_bindings,
        externals,
        changed: false,
    }))
}

#[derive(Debug)]
struct MetadataVisitor<'budget> {
    components: Vec<MetadataComponent>,
    uuids: Vec<MetadataUuid>,
    uuid_bindings: Vec<MetadataUuid>,
    externals: Vec<MetadataExternal>,
    budget: &'budget mut ReadBudget,
    budget_error: Option<Error>,
}

impl<'budget> MetadataVisitor<'budget> {
    fn new(budget: &'budget mut ReadBudget) -> Self {
        Self {
            components: Vec::new(),
            uuids: Vec::new(),
            uuid_bindings: Vec::new(),
            externals: Vec::new(),
            budget,
            budget_error: None,
        }
    }

    fn admit(&mut self, element_bytes: usize) -> bool {
        if self.budget_error.is_some() {
            return false;
        }
        if let Err(error) = self.budget.scratch(element_bytes) {
            self.budget_error = Some(error);
            return false;
        }
        if let Err(error) = ReadBudget::charge_counter(
            LimitKind::Allocations,
            &mut self.budget.allocations,
            1,
            self.budget.max_allocations,
        ) {
            self.budget_error = Some(error);
            return false;
        }
        true
    }
}

#[derive(Debug, Clone, Copy)]
struct MetadataUuid {
    component: u64,
    identifier: u64,
    uuid: UuidBits,
}

impl PackageMetadataVisitor for MetadataVisitor<'_> {
    fn visit_component(
        &mut self,
        component: metadata_codec::ComponentDescriptor<'_>,
    ) -> Result<(), MetadataRewriteError> {
        if !self.admit(size_of::<MetadataComponent>()) {
            return Ok(());
        }
        if self.components.try_reserve_exact(1).is_err() {
            self.budget_error = Some(Error::Allocation { amount: 1 });
            return Ok(());
        }
        if let Err(error) = self.budget.copy_text(component.effective_locator().len()) {
            self.budget_error = Some(error);
            return Ok(());
        }
        self.components.push(MetadataComponent {
            identifier: component.identifier(),
            locator: component.effective_locator().into(),
            current: component.is_current(),
        });
        Ok(())
    }

    fn visit_object_uuid(
        &mut self,
        binding: metadata_codec::ObjectUuidDescriptor<'_>,
    ) -> Result<(), MetadataRewriteError> {
        if !self.admit(size_of::<MetadataUuid>()) {
            return Ok(());
        }
        if self.uuids.try_reserve_exact(1).is_err() {
            self.budget_error = Some(Error::Allocation { amount: 1 });
            return Ok(());
        }
        if !self.admit(size_of::<MetadataUuid>()) {
            return Ok(());
        }
        if self.uuid_bindings.try_reserve_exact(1).is_err() {
            self.budget_error = Some(Error::Allocation { amount: 1 });
            return Ok(());
        }
        let metadata_uuid = MetadataUuid {
            component: binding.component().identifier(),
            identifier: binding.object_identifier(),
            uuid: binding.uuid(),
        };
        self.uuids.push(metadata_uuid);
        self.uuid_bindings.push(metadata_uuid);
        Ok(())
    }

    fn visit_external_reference(
        &mut self,
        reference: metadata_codec::ExternalReferenceDescriptor<'_>,
    ) -> Result<(), MetadataRewriteError> {
        if !self.admit(size_of::<MetadataExternal>()) {
            return Ok(());
        }
        if self.externals.try_reserve_exact(1).is_err() {
            self.budget_error = Some(Error::Allocation { amount: 1 });
            return Ok(());
        }
        self.externals.push(MetadataExternal {
            source: reference.source().identifier(),
            target: reference.target_component_identifier(),
            object: reference.object_identifier(),
            weak: reference.is_weak(),
            versioned: reference.is_versioned(),
            unknown: reference.has_unknown_fields(),
        });
        Ok(())
    }
}

fn push_selector<'a>(selectors: &mut Vec<ComponentSelector<'a>>, selector: ComponentSelector<'a>) {
    if !selectors.iter().any(|candidate| {
        candidate.identifier() == selector.identifier() && candidate.locator() == selector.locator()
    }) {
        selectors.push(selector);
    }
}

fn reserve_metadata_stage<T>(
    vector: &mut Vec<T>,
    capacity: usize,
    budget: &mut ReadBudget,
) -> Result<(), Error> {
    if capacity == 0 {
        return Ok(());
    }
    budget.scratch(
        capacity
            .checked_mul(size_of::<T>())
            .ok_or(Error::InvalidSource)?,
    )?;
    ReadBudget::charge_counter(
        LimitKind::Allocations,
        &mut budget.allocations,
        1,
        budget.max_allocations,
    )?;
    vector
        .try_reserve_exact(capacity)
        .map_err(|_| Error::Allocation { amount: capacity })
}

fn apply_metadata(working: &mut Working<'_>, budget: &mut ReadBudget) -> Result<(), Error> {
    let Some(metadata) = working.metadata.as_mut() else {
        return Ok(());
    };
    if working.identities.is_empty()
        && working.edge_changes.is_empty()
        && working.removed_identities.is_empty()
        && working.archives.is_empty()
    {
        return Ok(());
    }
    let identity_count = working.identities.len();
    let removed_count = working.removed_identities.len();
    let edge_count = working.edge_changes.len();
    let selector_capacity = working
        .archives
        .len()
        .checked_add(identity_count)
        .and_then(|count| count.checked_add(removed_count))
        .ok_or(Error::InvalidSource)?;
    budget.wire_work(
        identity_count
            .checked_add(removed_count)
            .and_then(|count| count.checked_add(edge_count.saturating_mul(4)))
            .and_then(|count| count.checked_add(selector_capacity))
            .ok_or(Error::InvalidSource)?,
    )?;
    let (payload, new_last) = {
        let mut uuid_additions = Vec::new();
        reserve_metadata_stage(&mut uuid_additions, identity_count, budget)?;
        for identity in &working.identities {
            let selector = selector_for_component(metadata, identity.component, working.source)?;
            uuid_additions.push(ObjectUuidAddition::new(
                selector,
                identity.identifier,
                UuidBits::new(identity.uuid.lower(), identity.uuid.upper()),
            ));
        }
        let mut selectors = Vec::new();
        reserve_metadata_stage(&mut selectors, selector_capacity, budget)?;
        for name in working.archives.keys() {
            if name == metadata.member.as_ref() {
                continue;
            }
            let selector = selector_for_name(metadata, name, working.source)?;
            push_selector(&mut selectors, selector);
        }
        for identity in &working.identities {
            let selector = selector_for_component(metadata, identity.component, working.source)?;
            push_selector(&mut selectors, selector);
        }
        for identity in &working.removed_identities {
            let selector = selector_for_component(metadata, identity.component, working.source)?;
            push_selector(&mut selectors, selector);
        }
        let mut uuid_removals = Vec::new();
        reserve_metadata_stage(&mut uuid_removals, removed_count, budget)?;
        for identity in &working.removed_identities {
            let selector = selector_for_component(metadata, identity.component, working.source)?;
            uuid_removals.push(ObjectUuidRemoval::new(
                selector,
                identity.identifier,
                UuidBits::new(identity.uuid.lower(), identity.uuid.upper()),
            ));
        }
        let mut external_additions = Vec::new();
        let mut external_removals = Vec::new();
        reserve_metadata_stage(&mut external_additions, edge_count, budget)?;
        reserve_metadata_stage(&mut external_removals, edge_count, budget)?;
        for edge in &working.edge_changes {
            let source_selector =
                selector_for_component(metadata, edge.source_component, working.source)?;
            push_selector(&mut selectors, source_selector);
            let old_selector = edge
                .old_target_component
                .map(|component| selector_for_component(metadata, component, working.source))
                .transpose()?;
            let new_selector = edge
                .new_target_component
                .map(|component| selector_for_component(metadata, component, working.source))
                .transpose()?;
            if let (Some(target), Some(identifier)) = (old_selector, edge.old_identifier) {
                if target.identifier() != source_selector.identifier() && edge.remove_old {
                    external_removals.push(ExternalReferenceRemoval::new(
                        source_selector,
                        target,
                        identifier,
                        edge.weak,
                    ));
                    push_selector(&mut selectors, target);
                }
            }
            if let (Some(target), Some(identifier)) = (new_selector, edge.new_identifier) {
                if target.identifier() != source_selector.identifier() {
                    external_additions.push(ExternalReferenceAddition::new(
                        source_selector,
                        target,
                        identifier,
                        edge.weak,
                    ));
                    push_selector(&mut selectors, target);
                }
            }
        }
        let new_last = working
            .identities
            .iter()
            .map(|identity| identity.identifier)
            .max()
            .unwrap_or(metadata.last_identifier)
            .max(metadata.last_identifier);
        let save_tokens = SaveTokenBatch::new(selectors.as_slice());
        let options = metadata_options(working.source, metadata.payload.len())?;
        let payload = if uuid_additions.is_empty()
            && uuid_removals.is_empty()
            && external_additions.is_empty()
            && external_removals.is_empty()
        {
            let request = SaveTokenBatch::new(selectors.as_slice());
            let prepared = metadata_codec::prepare_package_metadata_save_tokens(
                metadata.payload.as_slice(),
                request,
                options,
            )
            .map_err(map_metadata_error)?;
            let requirements = prepared.execution_requirements();
            charge_metadata_requirements(budget, metadata.payload.len(), requirements)?;
            prepared
                .execute(requirements.exact_limits())
                .map_err(map_metadata_error)?
                .into_bytes()
        } else if uuid_removals.is_empty() && external_removals.is_empty() {
            let additions = metadata_codec::Batch::new(
                metadata.last_identifier,
                new_last,
                uuid_additions.as_slice(),
                external_additions.as_slice(),
            );
            let request = metadata_codec::AdditionSaveTokenBatch::new(additions, save_tokens);
            let prepared = metadata_codec::prepare_package_metadata_additions_and_save_tokens(
                metadata.payload.as_slice(),
                request,
                options,
            )
            .map_err(map_metadata_error)?;
            let requirements = prepared.execution_requirements();
            charge_metadata_requirements(budget, metadata.payload.len(), requirements)?;
            prepared
                .execute(requirements.exact_limits())
                .map_err(map_metadata_error)?
                .into_bytes()
        } else if uuid_additions.is_empty() && external_additions.is_empty() {
            let removals = RemovalBatch::new(
                metadata.last_identifier,
                uuid_removals.as_slice(),
                external_removals.as_slice(),
                &[],
            );
            let request = RemovalSaveTokenBatch::new(removals, save_tokens);
            let prepared = metadata_codec::prepare_package_metadata_removals_and_save_tokens(
                metadata.payload.as_slice(),
                request,
                options,
            )
            .map_err(map_metadata_error)?;
            let requirements = prepared.execution_requirements();
            charge_metadata_requirements(budget, metadata.payload.len(), requirements)?;
            prepared
                .execute(requirements.exact_limits())
                .map_err(map_metadata_error)?
                .into_bytes()
        } else {
            let transition = CombinedBatch::new(
                metadata.last_identifier,
                new_last,
                uuid_additions.as_slice(),
                external_additions.as_slice(),
                uuid_removals.as_slice(),
                external_removals.as_slice(),
                &[],
            );
            let request = CombinedSaveTokenBatch::new(transition, save_tokens);
            let prepared = metadata_codec::prepare_package_metadata_combined_additions_and_removals_and_save_tokens(
                metadata.payload.as_slice(),
                request,
                options,
            )
            .map_err(map_metadata_error)?;
            let requirements = prepared.execution_requirements();
            charge_metadata_requirements(budget, metadata.payload.len(), requirements)?;
            prepared
                .execute(requirements.exact_limits())
                .map_err(map_metadata_error)?
                .into_bytes()
        };
        (payload, new_last)
    };
    let limits = working
        .source
        .state
        .source
        .limits()
        .effective_archive_limits()
        .map_err(|_| Error::UnsupportedSource)?;
    let object_identifier = metadata
        .archive
        .objects
        .get(metadata.object_index)
        .and_then(|object| object.archive_info.identifier)
        .ok_or(Error::InvalidSource)?;
    metadata
        .archive
        .object_mut(object_identifier)
        .ok_or(Error::InvalidSource)?
        .replace_message_preserving_header_with_limits(
            metadata.message_index,
            RawMessage {
                type_: METADATA_MESSAGE_TYPE,
                data: payload,
            },
            limits,
        )
        .map_err(|_| Error::InvalidSource)?;
    metadata.last_identifier = new_last;
    metadata.changed = true;
    Ok(())
}

fn charge_metadata_requirements(
    budget: &mut ReadBudget,
    input_bytes: usize,
    requirements: metadata_codec::RewriteExecutionRequirements,
) -> Result<(), Error> {
    budget.wire_input(input_bytes, 1)?;
    budget.wire_fields(requirements.fields())?;
    budget.wire_work(
        requirements
            .work_bytes()
            .checked_add(requirements.components())
            .ok_or(Error::InvalidSource)?,
    )?;
    budget.references(requirements.references())?;
    let scratch = requirements
        .output_bytes()
        .checked_add(requirements.retained_bytes())
        .and_then(|bytes| bytes.checked_add(requirements.scratch_bytes()))
        .ok_or(Error::InvalidSource)?;
    budget.scratch(scratch)?;
    ReadBudget::charge_counter(
        LimitKind::Allocations,
        &mut budget.allocations,
        requirements.allocations(),
        budget.max_allocations,
    )
}

fn selector_for_name<'a>(
    metadata: &'a MetadataState,
    name: &str,
    _source: &'a Package,
) -> Result<ComponentSelector<'a>, Error> {
    let locator = component_locator(name).ok_or(Error::InvalidSource)?;
    let component = metadata_component_for_locator(metadata, locator)?;
    Ok(ComponentSelector::new(
        component.identifier,
        component.locator.as_ref(),
    ))
}

fn metadata_component_for_index<'a>(
    metadata: &'a MetadataState,
    component_index: usize,
    source: &'a Package,
) -> Result<&'a MetadataComponent, Error> {
    let selector = selector_for_component(metadata, component_index, source)?;
    metadata
        .components
        .iter()
        .find(|component| component.identifier == selector.identifier())
        .ok_or(Error::InvalidSource)
}

fn selector_for_component<'a>(
    metadata: &'a MetadataState,
    component_index: usize,
    source: &'a Package,
) -> Result<ComponentSelector<'a>, Error> {
    let name = source
        .state
        .source
        .components()
        .get_index(component_index)
        .ok_or(Error::InvalidSource)?
        .name();
    let locator = component_locator(name).ok_or(Error::InvalidSource)?;
    let component = metadata_component_for_locator(metadata, locator)?;
    Ok(ComponentSelector::new(
        component.identifier,
        component.locator.as_ref(),
    ))
}

fn metadata_component_for_locator<'a>(
    metadata: &'a MetadataState,
    locator: &str,
) -> Result<&'a MetadataComponent, Error> {
    let mut selected = None;
    for component in &metadata.components {
        if !component.current || component.locator.as_ref() != locator {
            continue;
        }
        if selected.replace(component).is_some() {
            return Err(Error::InvalidSource);
        }
    }
    selected.ok_or(Error::UnsupportedDependency)
}

fn component_locator(name: &str) -> Option<&str> {
    name.strip_prefix("Index/")
        .unwrap_or(name)
        .strip_suffix(".iwa")
}

fn metadata_options(source: &Package, payload_len: usize) -> Result<MetadataRewriteOptions, Error> {
    let archive = source
        .state
        .source
        .limits()
        .effective_archive_limits()
        .map_err(|_| Error::UnsupportedSource)?;
    let input = payload_len.max(1);
    let component_bound = source.state.source.components().len().max(1);
    let metadata_items = archive
        .max_metadata_items()
        .saturating_mul(component_bound)
        .max(1);
    Ok(MetadataRewriteOptions::new(
        input,
        source
            .state
            .source
            .limits()
            .max_iwa_stream_bytes()
            .max(input),
        archive.max_header_fields().max(input.min(1_000_000)),
        input.checked_mul(128).ok_or(Error::InvalidSource)?,
        u32::try_from(archive.max_header_nesting()).unwrap_or(u32::MAX),
        component_bound.saturating_mul(64).max(1),
        metadata_items,
        metadata_items,
    ))
}

fn map_metadata_error(error: MetadataRewriteError) -> Error {
    if let Some(limit) = error.resource_limit() {
        let (kind, observed, maximum) = match limit {
            metadata_codec::RewriteLimit::InputBytes { observed, maximum } => {
                (LimitKind::WireBytes, observed, maximum)
            },
            metadata_codec::RewriteLimit::OutputBytes { observed, maximum } => {
                (LimitKind::OutputBytes, observed, maximum)
            },
            metadata_codec::RewriteLimit::Fields { observed, maximum } => {
                (LimitKind::WireFields, observed, maximum)
            },
            metadata_codec::RewriteLimit::Work { observed, maximum } => {
                (LimitKind::WireWork, observed, maximum)
            },
            metadata_codec::RewriteLimit::Nesting { observed, maximum } => {
                return Error::LimitExceeded {
                    kind: LimitKind::WireNesting,
                    observed: u64::from(observed),
                    maximum: u64::from(maximum),
                };
            },
            metadata_codec::RewriteLimit::Components { observed, maximum } => {
                (LimitKind::References, observed, maximum)
            },
            metadata_codec::RewriteLimit::References { observed, maximum }
            | metadata_codec::RewriteLimit::Additions { observed, maximum } => {
                (LimitKind::References, observed, maximum)
            },
            _ => return Error::InvalidSource,
        };
        return Error::LimitExceeded {
            kind,
            observed: observed as u64,
            maximum: maximum as u64,
        };
    }
    if let Some(amount) = error.allocation_request() {
        return Error::Allocation { amount };
    }
    Error::InvalidSource
}

fn charge_archive_sizing_scan(
    archive: &Archive,
    limits: litchi_iwa_core::Limits,
    budget: &mut ReadBudget,
) -> Result<usize, Error> {
    let object_count = archive.objects.len();
    let message_bound = object_count
        .checked_mul(limits.max_messages_per_object())
        .ok_or(Error::InvalidSource)?;
    let work = object_count
        .checked_mul(size_of::<ArchiveObject>())
        .and_then(|bytes| bytes.checked_add(message_bound.checked_mul(size_of::<RawMessage>())?))
        .ok_or(Error::InvalidSource)?;
    budget.wire_work(work)?;
    let encoded = archive
        .encoded_len_with_limits(limits)
        .map_err(|_| Error::InvalidSource)?;
    budget.wire_work(encoded)?;
    Ok(encoded)
}

/// Admit every heap collection visited by a neutral archive clone before
/// the clone primitive executes.  The shared inventory callback runs before
/// each nested collection, so a residual budget rejects the next allocation
/// without traversing its children.
fn charge_archive_clone_admission(archive: &Archive, budget: &mut ReadBudget) -> Result<(), Error> {
    let mut visitor = |chunk: litchi_iwa_core::ArchiveCloneInventoryChunk| {
        budget.scratch(chunk.retained_bytes())?;
        budget.wire_work(chunk.traversal_bytes())?;
        match chunk.kind() {
            litchi_iwa_core::ArchiveCloneInventoryKind::FieldInfos => {
                budget.wire_fields(chunk.elements())?;
            },
            litchi_iwa_core::ArchiveCloneInventoryKind::ObjectReferences
            | litchi_iwa_core::ArchiveCloneInventoryKind::DataReferences => {
                budget.references(chunk.elements())?;
            },
            _ => {},
        }
        ReadBudget::charge_counter(
            LimitKind::Allocations,
            &mut budget.allocations,
            chunk.allocations(),
            budget.max_allocations,
        )
    };
    archive
        .clone_inventory_with_admission(&mut visitor)
        .map_err(|error| match error {
            litchi_iwa_core::ArchiveCloneInventoryError::Admission(error) => error,
            litchi_iwa_core::ArchiveCloneInventoryError::ArithmeticOverflow { .. } => {
                Error::InvalidSource
            },
            _ => Error::InvalidSource,
        })?;
    Ok(())
}

fn charge_archive_object_clone_admission(
    object: &ArchiveObject,
    budget: &mut ReadBudget,
) -> Result<(), Error> {
    let mut visitor = |chunk: litchi_iwa_core::ArchiveCloneInventoryChunk| {
        budget.scratch(chunk.retained_bytes())?;
        budget.wire_work(chunk.traversal_bytes())?;
        match chunk.kind() {
            litchi_iwa_core::ArchiveCloneInventoryKind::FieldInfos => {
                budget.wire_fields(chunk.elements())?;
            },
            litchi_iwa_core::ArchiveCloneInventoryKind::ObjectReferences
            | litchi_iwa_core::ArchiveCloneInventoryKind::DataReferences => {
                budget.references(chunk.elements())?;
            },
            _ => {},
        }
        ReadBudget::charge_counter(
            LimitKind::Allocations,
            &mut budget.allocations,
            chunk.allocations(),
            budget.max_allocations,
        )
    };
    object
        .clone_inventory_with_admission(&mut visitor)
        .map_err(|error| match error {
            litchi_iwa_core::ArchiveCloneInventoryError::Admission(error) => error,
            litchi_iwa_core::ArchiveCloneInventoryError::ArithmeticOverflow { .. } => {
                Error::InvalidSource
            },
            _ => Error::InvalidSource,
        })?;
    Ok(())
}

fn candidate_admission(
    working: &Working<'_>,
    budget: &mut ReadBudget,
) -> Result<CandidateAdmission, Error> {
    let limits = working
        .source
        .state
        .source
        .limits()
        .effective_archive_limits()
        .map_err(|_| Error::UnsupportedSource)?;
    // The detailed inventory below must not be the first ledger operation.
    // Admit the bounded archive/message walk from the immutable source census
    // before visiting any object, field path, or reference.  The exact
    // counters are charged by CandidateAdmission::charge after the walk.
    let object_bound = working
        .source
        .state
        .object_count
        .checked_add(working.identities.len())
        .ok_or(Error::InvalidSource)?;
    let message_bound = object_bound
        .checked_mul(limits.max_messages_per_object())
        .ok_or(Error::InvalidSource)?;
    let scan_work = object_bound
        .checked_mul(size_of::<ArchiveObject>())
        .and_then(|bytes| bytes.checked_add(message_bound.checked_mul(size_of::<RawMessage>())?))
        .ok_or(Error::InvalidSource)?;
    budget.wire_work(scan_work)?;
    budget.scratch(
        object_bound
            .checked_mul(size_of::<ArchiveObject>())
            .and_then(|bytes| {
                bytes.checked_add(message_bound.checked_mul(size_of::<RawMessage>())?)
            })
            .ok_or(Error::InvalidSource)?,
    )?;
    let mut admission = CandidateAdmission::EMPTY;
    let package = working.source.state.source.package();
    admission.entries = package.len();
    budget.wire_work(
        package
            .len()
            .checked_mul(size_of::<EntryEdit<'_>>())
            .ok_or(Error::InvalidSource)?,
    )?;
    let mut catalog_bytes = 0usize;
    for entry in package.iter() {
        catalog_bytes = catalog_bytes
            .checked_add(entry.raw_name().len())
            .and_then(|bytes| bytes.checked_add(entry.metadata().local().name().len()))
            .and_then(|bytes| bytes.checked_add(entry.metadata().local().extra().len()))
            .and_then(|bytes| bytes.checked_add(entry.metadata().local().comment().len()))
            .and_then(|bytes| bytes.checked_add(entry.metadata().central().name().len()))
            .and_then(|bytes| bytes.checked_add(entry.metadata().central().extra().len()))
            .and_then(|bytes| bytes.checked_add(entry.metadata().central().comment().len()))
            .ok_or(Error::InvalidSource)?;
    }
    admission.catalog_bytes = catalog_bytes;
    let document = &working.source.state.document;
    let mut semantic_bytes = document.text_len();
    let mut semantic_allocations = usize::from(!document.sections().is_empty());
    for section in document.sections() {
        semantic_bytes = semantic_bytes
            .checked_add(section.name().map_or(0, str::len))
            .and_then(|bytes| bytes.checked_add(section.heading().map_or(0, str::len)))
            .ok_or(Error::InvalidSource)?;
        semantic_allocations = semantic_allocations
            .checked_add(1)
            .and_then(|count| count.checked_add(section.paragraphs().len()))
            .and_then(|count| count.checked_add(section.text_storages().len()))
            .ok_or(Error::InvalidSource)?;
        for paragraph in section.paragraphs() {
            semantic_bytes = semantic_bytes
                .checked_add(paragraph.len())
                .ok_or(Error::InvalidSource)?;
        }
        for storage in section.text_storages() {
            semantic_bytes = semantic_bytes
                .checked_add(storage.text().len())
                .ok_or(Error::InvalidSource)?;
        }
    }
    admission.semantic_bytes = semantic_bytes;
    admission.allocations = admission
        .allocations
        .checked_add(semantic_allocations)
        .ok_or(Error::InvalidSource)?;
    let mut metadata_bytes = 0usize;
    let mut metadata_allocations = 0usize;
    macro_rules! count_metadata_text {
        ($value:expr) => {
            if let Some(value) = $value.as_ref() {
                metadata_bytes = metadata_bytes
                    .checked_add(value.len())
                    .ok_or(Error::InvalidSource)?;
                metadata_allocations = metadata_allocations
                    .checked_add(1)
                    .ok_or(Error::InvalidSource)?;
            }
        };
    }
    count_metadata_text!(working.source.state.metadata.title);
    count_metadata_text!(working.source.state.metadata.author);
    count_metadata_text!(working.source.state.metadata.keywords);
    count_metadata_text!(working.source.state.metadata.description);
    count_metadata_text!(working.source.state.metadata.application);
    count_metadata_text!(working.source.state.metadata.revision);
    count_metadata_text!(working.source.state.metadata.format_version);
    count_metadata_text!(working.source.state.metadata.build_version);
    count_metadata_text!(working.source.state.metadata.identifier);
    admission.metadata_bytes = metadata_bytes;
    admission.allocations = admission
        .allocations
        .checked_add(metadata_allocations)
        .ok_or(Error::InvalidSource)?;
    let mut archive_retained_bytes = 0usize;
    for component in working.source.state.source.components().iter() {
        let archive = if working.metadata.as_ref().is_some_and(|metadata| {
            metadata.member.as_ref() == component.name() && metadata.changed
        }) {
            &working
                .metadata
                .as_ref()
                .ok_or(Error::InvalidSource)?
                .archive
        } else {
            working
                .archives
                .get(component.name())
                .unwrap_or_else(|| component.archive())
        };
        let mut admitted_inventory = |chunk: litchi_iwa_core::ArchiveCloneInventoryChunk| {
            budget.scratch(chunk.retained_bytes())?;
            budget.wire_work(chunk.traversal_bytes())?;
            match chunk.kind() {
                litchi_iwa_core::ArchiveCloneInventoryKind::FieldInfos
                | litchi_iwa_core::ArchiveCloneInventoryKind::FieldPath => {
                    budget.wire_fields(chunk.elements())?;
                },
                litchi_iwa_core::ArchiveCloneInventoryKind::ObjectReferences
                | litchi_iwa_core::ArchiveCloneInventoryKind::DataReferences => {
                    budget.references(chunk.elements())?;
                },
                _ => {},
            }
            ReadBudget::charge_counter(
                LimitKind::Allocations,
                &mut budget.allocations,
                chunk.allocations(),
                budget.max_allocations,
            )
        };
        let inventory = archive
            .clone_inventory_with_admission(&mut admitted_inventory)
            .map_err(|error| match error {
                litchi_iwa_core::ArchiveCloneInventoryError::Admission(error) => error,
                litchi_iwa_core::ArchiveCloneInventoryError::ArithmeticOverflow { .. } => {
                    Error::InvalidSource
                },
                _ => Error::InvalidSource,
            })?;
        let encoded_bound = archive
            .encoded_len_with_limits(limits)
            .map_err(|_| Error::InvalidSource)?;
        admission.wire_bytes = admission
            .wire_bytes
            .checked_add(encoded_bound)
            .ok_or(Error::InvalidSource)?;
        admission.wire_work = admission
            .wire_work
            .checked_add(encoded_bound)
            .ok_or(Error::InvalidSource)?;
        admission.objects = admission
            .objects
            .checked_add(archive.objects.len())
            .ok_or(Error::InvalidSource)?;
        admission.allocations = admission
            .allocations
            .checked_add(inventory.allocations())
            .ok_or(Error::InvalidSource)?;
        archive_retained_bytes = archive_retained_bytes
            .checked_add(inventory.retained_bytes())
            .ok_or(Error::InvalidSource)?;
        for object in &archive.objects {
            admission.wire_fields = admission
                .wire_fields
                .checked_add(
                    object
                        .archive_info
                        .message_infos
                        .iter()
                        .map(|info| {
                            1usize
                                .saturating_add(info.field_infos.len())
                                .saturating_add(
                                    info.field_infos
                                        .iter()
                                        .map(|field| field.path.path.len())
                                        .sum::<usize>(),
                                )
                        })
                        .sum::<usize>(),
                )
                .ok_or(Error::InvalidSource)?;
            admission.references = admission
                .references
                .checked_add(
                    object
                        .archive_info
                        .message_infos
                        .iter()
                        .map(|info| {
                            info.object_references.len().saturating_add(
                                info.data_references.len().saturating_add(
                                    info.field_infos
                                        .iter()
                                        .map(|field| {
                                            field
                                                .object_references
                                                .len()
                                                .saturating_add(field.data_references.len())
                                        })
                                        .sum::<usize>(),
                                ),
                            )
                        })
                        .sum::<usize>(),
                )
                .ok_or(Error::InvalidSource)?;
            for message in &object.messages {
                admission.messages = admission
                    .messages
                    .checked_add(1)
                    .ok_or(Error::InvalidSource)?;
                admission.payload_bytes = admission
                    .payload_bytes
                    .checked_add(message.data.len())
                    .ok_or(Error::InvalidSource)?;
                admission.wire_work = admission
                    .wire_work
                    .checked_add(message.data.len())
                    .ok_or(Error::InvalidSource)?;
            }
        }
    }
    admission.scratch_bytes = admission
        .wire_bytes
        .checked_add(
            admission
                .objects
                .checked_mul(size_of::<ArchiveObject>())
                .ok_or(Error::InvalidSource)?,
        )
        .and_then(|bytes| {
            bytes.checked_add(admission.messages.checked_mul(size_of::<RawMessage>())?)
        })
        .and_then(|bytes| {
            bytes.checked_add(
                admission
                    .wire_fields
                    .checked_mul(size_of::<MessageInfo>())?,
            )
        })
        .and_then(|bytes| bytes.checked_add(archive_retained_bytes))
        .ok_or(Error::InvalidSource)?;
    Ok(admission)
}

pub(crate) fn package_candidate_admission(
    source: &Package,
    budget: &mut ReadBudget,
) -> Result<CandidateAdmission, Error> {
    let working = Working {
        source,
        archives: BTreeMap::new(),
        reserved_ids: Vec::new(),
        identities: Vec::new(),
        edge_changes: Vec::new(),
        removed_identities: Vec::new(),
        removed_objects: Vec::new(),
        metadata: None,
    };
    candidate_admission(&working, budget)
}

fn serialize_candidate(
    working: &mut Working<'_>,
    budget: &mut ReadBudget,
) -> Result<TableCommentMutationOutput, Error> {
    let metadata_name = working
        .metadata
        .as_ref()
        .filter(|metadata| metadata.changed)
        .map(|metadata| metadata.member.as_ref());
    let changed_count = working
        .archives
        .keys()
        .filter(|name| Some(name.as_str()) != metadata_name)
        .count()
        .checked_add(usize::from(metadata_name.is_some()))
        .ok_or(Error::InvalidSource)?;
    let tuple_bytes = changed_count
        .checked_mul(size_of::<(String, Vec<u8>)>())
        .ok_or(Error::InvalidSource)?;
    ReadBudget::charge_counter(
        LimitKind::Allocations,
        &mut budget.allocations,
        usize::from(changed_count != 0),
        budget.max_allocations,
    )?;
    budget.scratch(tuple_bytes)?;
    let mut compressed = Vec::new();
    compressed
        .try_reserve_exact(changed_count)
        .map_err(|_| Error::Allocation {
            amount: changed_count,
        })?;
    let archive_limits = working
        .source
        .state
        .source
        .limits()
        .effective_archive_limits()
        .map_err(|_| Error::UnsupportedSource)?;
    for (name, archive) in &working.archives {
        if let Some(metadata) = &working.metadata {
            if name == metadata.member.as_ref() {
                continue;
            }
        }
        budget.copy_text(name.len())?;
        let (_encoded, payload) = serialize_archive(archive, archive_limits, budget)?;
        compressed.push((name.clone(), payload));
    }
    if let Some(metadata) = &working.metadata {
        if !metadata.changed {
            // A metadata sidecar is borrowed and retained byte-for-byte when
            // the operation only rewrites an existing unique storage node.
            return finish_reassembly(working, &compressed, budget);
        }
        let (_encoded, payload) = serialize_archive(&metadata.archive, archive_limits, budget)?;
        budget.copy_text(metadata.member.len())?;
        compressed.push((metadata.member.to_string(), payload));
    }
    finish_reassembly(working, &compressed, budget)
}

fn serialize_archive(
    archive: &Archive,
    limits: litchi_iwa_core::Limits,
    budget: &mut ReadBudget,
) -> Result<(Vec<u8>, Vec<u8>), Error> {
    let bound = charge_archive_sizing_scan(archive, limits, budget)?;
    let maximum_compressed =
        SnappyStream::maximum_compressed_len(bound).map_err(|_| Error::InvalidSource)?;
    budget.wire_work(bound)?;
    budget.scratch(
        bound
            .checked_add(maximum_compressed)
            .ok_or(Error::InvalidSource)?,
    )?;
    ReadBudget::charge_counter(
        LimitKind::Allocations,
        &mut budget.allocations,
        2,
        budget.max_allocations,
    )?;
    let encoded_len = archive
        .encoded_len_with_limits(limits)
        .map_err(|_| Error::InvalidSource)?;
    let encoded = archive
        .to_bytes_with_limits(limits)
        .map_err(|_| Error::InvalidSource)?;
    if encoded.len() != encoded_len {
        return Err(Error::Verification);
    }
    let payload = SnappyStream::compress(&encoded).map_err(|_| Error::InvalidSource)?;
    if payload.len() > maximum_compressed {
        return Err(Error::Verification);
    }
    Ok((encoded, payload))
}

fn charge_reassembly_plan(
    working: &Working<'_>,
    compressed: &[(String, Vec<u8>)],
    budget: &mut ReadBudget,
) -> Result<(), Error> {
    let entry_count = working.source.state.source.package().len();
    let compressed_bytes = compressed.iter().try_fold(0usize, |sum, (name, payload)| {
        sum.checked_add(name.len())
            .and_then(|value| value.checked_add(payload.len()))
            .ok_or(Error::InvalidSource)
    })?;
    let edit_slots = entry_count
        .checked_mul(size_of::<EntryEdit<'_>>())
        .ok_or(Error::InvalidSource)?;
    let source_bytes = working.source.source_bytes().len();
    let planning_work = source_bytes
        .checked_add(compressed_bytes)
        .and_then(|bytes| bytes.checked_add(edit_slots))
        .ok_or(Error::InvalidSource)?;
    budget.wire_work(planning_work)?;
    budget.scratch(
        edit_slots
            .checked_add(compressed_bytes)
            .ok_or(Error::InvalidSource)?,
    )?;
    ReadBudget::charge_counter(
        LimitKind::Allocations,
        &mut budget.allocations,
        entry_count
            .checked_mul(4)
            .and_then(|count| count.checked_add(compressed.len().saturating_mul(2)))
            .ok_or(Error::InvalidSource)?,
        budget.max_allocations,
    )
}

fn finish_reassembly(
    working: &Working<'_>,
    compressed: &[(String, Vec<u8>)],
    budget: &mut ReadBudget,
) -> Result<TableCommentMutationOutput, Error> {
    let limits = working
        .source
        .state
        .source
        .limits()
        .effective_archive_limits()
        .map_err(|_| Error::UnsupportedSource)?;
    let object_bound = working
        .source
        .state
        .object_count
        .checked_add(working.identities.len())
        .ok_or(Error::InvalidSource)?;
    let message_bound = object_bound
        .checked_mul(limits.max_messages_per_object())
        .ok_or(Error::InvalidSource)?;
    budget.wire_work(
        object_bound
            .checked_mul(size_of::<ArchiveObject>())
            .and_then(|bytes| {
                bytes.checked_add(message_bound.checked_mul(size_of::<RawMessage>())?)
            })
            .ok_or(Error::InvalidSource)?,
    )?;
    let candidate_admission = candidate_admission(working, budget)?;
    // Admit the writer's decoded-inventory census before walking all
    // effective archives to compute the plan.  The candidate parse below is
    // charged a second time because it repeats that inventory from the
    // reopened physical artifact.
    budget.wire_fields(candidate_admission.wire_fields)?;
    budget.wire_work(candidate_admission.wire_work)?;
    budget.references(candidate_admission.references)?;
    let mut edits = Vec::new();
    charge_reassembly_plan(working, compressed, budget)?;
    ReadBudget::charge_counter(
        LimitKind::Allocations,
        &mut budget.allocations,
        usize::from(!compressed.is_empty()),
        budget.max_allocations,
    )?;
    edits
        .try_reserve_exact(compressed.len())
        .map_err(|_| Error::Allocation {
            amount: compressed.len(),
        })?;
    for (name, payload) in compressed {
        edits.push(EntryEdit::new(name.as_str(), payload.as_slice()));
    }
    let prepared = working
        .source
        .state
        .source
        .prepare_reassembly(&edits, working.source.state.source.limits())
        .map_err(|_| Error::InvalidSource)?;
    let requirements = prepared.execution_requirements();
    budget.charge_target(requirements.output_bytes())?;
    budget.scratch(requirements.scratch_bytes())?;
    budget.wire_work(requirements.offset_count())?;
    ReadBudget::charge_counter(
        LimitKind::Allocations,
        &mut budget.allocations,
        requirements.allocations(),
        budget.max_allocations,
    )?;
    let output = prepared
        .execute(litchi_iwa_archive::package::ReassemblyExecutionLimits::new(
            requirements.output_bytes(),
            requirements.offset_count(),
            requirements.scratch_bytes(),
            requirements.retained_bytes(),
            requirements.allocations(),
        ))
        .map_err(|_| Error::InvalidSource)?;
    Ok(TableCommentMutationOutput {
        bytes: output,
        touched_components: compressed.len(),
        candidate_admission,
    })
}
