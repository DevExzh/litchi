//! Shared native root-comment graph and rewrite engine.
//!
//! This module is deliberately below the Pages, Keynote, and Numbers public
//! facades.  A leaf resolves its selector into [`Selection`](struct.Selection)
//! and chooses one [`Profile`](struct.Profile); this module validates that evidence,
//! performs the source-preserving copy-on-write rewrite, and reopens the
//! candidate through [`structure::Working`].  Native identifiers therefore
//! never cross a supported facade and the three owners share one mutation
//! implementation.

#![allow(
    clippy::arbitrary_source_item_ordering,
    clippy::too_many_arguments,
    clippy::too_many_lines,
    reason = "The bounded wire transaction keeps its validation and rewrite phases together."
)]

use std::sync::Arc;

use litchi_iwa_common::WireLimits;
use litchi_iwa_common::decode_varint_from_bytes;
use litchi_iwa_common::wire::{self, WireView};
use litchi_iwa_core::archive::{
    FieldInfo, FieldObjectReferenceTransition, FieldPath, FieldType, ObjectReferenceTransition,
};
use litchi_iwa_core::{ArchiveObject, RawMessage};
use litchi_iwa_protos::comment_storage_codec::{
    self as comment_codec, CommentStorageLeafWrite, DateSnapshot, UuidSnapshot,
};
use litchi_iwa_protos::numbers_table_cell_storage_codec as table_codec;

use crate::structure::{self, Fault as StructureFault, TableLeaf, Working, metadata};

/// The resource dimensions specific to the native comment wire codecs.
///
/// These stay at the comment-native boundary instead of widening the shared
/// structural fault taxonomy.  Other table families therefore cannot
/// accidentally coerce a comment codec's work/reference ceiling into their
/// payload-byte limit.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum LimitKind {
    /// Source bytes inspected by a wire parser.
    InputBytes,
    /// Complete source or candidate package bytes.
    PackageBytes,
    /// Decoded table/comment payload bytes.
    PayloadBytes,
    /// Native object references inspected.
    References,
    /// Native objects inspected by the archive validator.
    Objects,
    /// Native messages inspected by the archive validator.
    Messages,
    /// Native messages per object inspected by the archive validator.
    MessagesPerObject,
    /// Native metadata items inspected by the archive validator.
    MetadataItems,
    /// Snappy frames inspected by the archive validator.
    SnappyFrames,
    /// Reply records inspected.
    Replies,
    /// Comment or reply text bytes retained.
    TextBytes,
    /// Wire fields inspected.
    WireFields,
    /// Wire bytes inspected.
    WireBytes,
    /// Wire output bytes produced.
    WireOutputBytes,
    /// Wire nesting inspected.
    WireNesting,
    /// Wire work performed.
    WireWork,
    /// Addressable table rows inspected.
    TableRows,
    /// Addressable table columns inspected.
    TableColumns,
    /// Addressable table cells inspected.
    TableCells,
    /// Materialized sparse table cells inspected.
    MaterializedCells,
    /// Repeated Numbers-wire items inspected.
    Items,
    /// Bounded codec allocations requested.
    Allocations,
    /// Temporary rewrite scratch bytes retained.
    ScratchBytes,
    /// Candidate bytes retained by the codec.
    RetainedBytes,
}

/// A bounded failure from the shared native comment writer.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Fault {
    /// The selected native graph is not supported by this profile.
    Unsupported,
    /// The selected source or wire graph is malformed.
    Invalid,
    /// A required comment dependency is outside this focused route.
    Dependency,
    /// A comment-specific resource ceiling was exceeded.
    Limit {
        /// Resource category.
        kind: LimitKind,
        /// Observed or requested amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded staging allocation failed.
    Allocation {
        /// Requested allocation amount.
        amount: usize,
    },
    /// Candidate reopen/readback did not reproduce the rewrite.
    Verification,
}

type FaultResult<T> = Result<T, Fault>;

impl From<StructureFault> for Fault {
    fn from(fault: StructureFault) -> Self {
        match fault {
            StructureFault::Unsupported => Self::Unsupported,
            StructureFault::Invalid => Self::Invalid,
            StructureFault::Dependency => Self::Dependency,
            StructureFault::Limit {
                kind,
                observed,
                maximum,
            } => Self::Limit {
                kind: match kind {
                    structure::FaultLimit::PackageBytes => LimitKind::PackageBytes,
                    structure::FaultLimit::PayloadBytes => LimitKind::PayloadBytes,
                    structure::FaultLimit::Identifiers => LimitKind::References,
                },
                observed,
                maximum,
            },
            StructureFault::Allocation { amount } => Self::Allocation { amount },
            StructureFault::Verification => Self::Verification,
        }
    }
}

fn map_wire_limit(kind: litchi_iwa_common::LimitKind) -> LimitKind {
    match kind {
        litchi_iwa_common::LimitKind::InputBytes => LimitKind::InputBytes,
        litchi_iwa_common::LimitKind::Fields => LimitKind::WireFields,
        litchi_iwa_common::LimitKind::OutputBytes => LimitKind::WireOutputBytes,
        litchi_iwa_common::LimitKind::Nesting => LimitKind::WireNesting,
        litchi_iwa_common::LimitKind::RewriteWork => LimitKind::WireWork,
        litchi_iwa_common::LimitKind::TableRows => LimitKind::TableRows,
        litchi_iwa_common::LimitKind::TableColumns => LimitKind::TableColumns,
        litchi_iwa_common::LimitKind::TableCells => LimitKind::TableCells,
        litchi_iwa_common::LimitKind::MaterializedCells => LimitKind::MaterializedCells,
    }
}

fn map_wire_error(error: litchi_iwa_common::Error) -> Fault {
    match error {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => Fault::Limit {
            kind: map_wire_limit(kind),
            observed: observed as u64,
            maximum: limit as u64,
        },
        litchi_iwa_common::Error::Allocation { amount, .. } => Fault::Allocation { amount },
        _ => Fault::Invalid,
    }
}

fn map_core_limit(kind: litchi_iwa_core::LimitKind) -> LimitKind {
    match kind {
        litchi_iwa_core::LimitKind::HeaderBytes => LimitKind::InputBytes,
        litchi_iwa_core::LimitKind::HeaderFields => LimitKind::WireFields,
        litchi_iwa_core::LimitKind::HeaderNesting => LimitKind::WireNesting,
        litchi_iwa_core::LimitKind::HeaderMemoryBytes => LimitKind::ScratchBytes,
        litchi_iwa_core::LimitKind::Messages => LimitKind::Messages,
        litchi_iwa_core::LimitKind::MessagesPerObject => LimitKind::MessagesPerObject,
        litchi_iwa_core::LimitKind::Objects => LimitKind::Objects,
        litchi_iwa_core::LimitKind::MetadataItems => LimitKind::MetadataItems,
        litchi_iwa_core::LimitKind::ObjectBytes | litchi_iwa_core::LimitKind::MessageBytes => {
            LimitKind::PayloadBytes
        },
        litchi_iwa_core::LimitKind::ArchiveBytes
        | litchi_iwa_core::LimitKind::SnappyChunkBytes
        | litchi_iwa_core::LimitKind::SnappyStreamBytes
        | litchi_iwa_core::LimitKind::SnappyCompressedChunkBytes
        | litchi_iwa_core::LimitKind::SnappyCompressedStreamBytes => LimitKind::PackageBytes,
        litchi_iwa_core::LimitKind::SnappyFrames => LimitKind::SnappyFrames,
    }
}

fn map_core_error(error: litchi_iwa_core::Error) -> Fault {
    match error {
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => Fault::Limit {
            kind: map_core_limit(kind),
            observed: observed as u64,
            maximum: maximum as u64,
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            Fault::Allocation { amount: requested }
        },
        _ => Fault::Invalid,
    }
}

fn map_bnc_error(error: litchi_numbers_wire::Error) -> Fault {
    match error {
        litchi_numbers_wire::Error::LimitExceeded {
            kind,
            observed,
            maximum,
        } => Fault::Limit {
            kind: match kind {
                litchi_numbers_wire::LimitKind::InputBytes => LimitKind::InputBytes,
                litchi_numbers_wire::LimitKind::OutputBytes => LimitKind::WireOutputBytes,
                litchi_numbers_wire::LimitKind::Fields => LimitKind::WireFields,
                litchi_numbers_wire::LimitKind::TailBytes => LimitKind::PayloadBytes,
                litchi_numbers_wire::LimitKind::Items => LimitKind::Items,
                _ => return Fault::Invalid,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        litchi_numbers_wire::Error::Allocation { requested } => {
            Fault::Allocation { amount: requested }
        },
        _ => Fault::Invalid,
    }
}

#[derive(Debug)]
struct WireCallbackError(Fault);

impl From<litchi_iwa_common::Error> for WireCallbackError {
    fn from(error: litchi_iwa_common::Error) -> Self {
        Self(map_wire_error(error))
    }
}

impl From<Fault> for WireCallbackError {
    fn from(fault: Fault) -> Self {
        Self(fault)
    }
}

impl From<StructureFault> for WireCallbackError {
    fn from(fault: StructureFault) -> Self {
        Self(Fault::from(fault))
    }
}

const TABLE_TILE_ROWS_FIELD: u32 = 5;
const TABLE_TILE_ROW_INDEX_FIELD: u32 = 1;
const TABLE_TILE_ROW_CELL_COUNT_FIELD: u32 = 2;
const TABLE_TILE_ROW_CELL_STORAGE_FIELD: u32 = 6;
const TABLE_TILE_ROW_CELL_OFFSETS_FIELD: u32 = 7;
const TABLE_TILE_ROW_WIDE_OFFSETS_FIELD: u32 = 8;
const TABLE_LIST_TYPE_FIELD: u32 = 1;
const TABLE_LIST_NEXT_ID_FIELD: u32 = 2;
const TABLE_LIST_ENTRY_FIELD: u32 = 3;
const TABLE_ENTRY_KEY_FIELD: u32 = 1;
const TABLE_ENTRY_REFCOUNT_FIELD: u32 = 2;
const TABLE_ENTRY_STORAGE_FIELD: u32 = 10;
const TABLE_REFERENCE_IDENTIFIER_FIELD: u32 = 1;
const TABLE_MODEL_COMMENT_PATH: [u32; 2] = [4, 19];
const TABLE_DATA_LIST_MESSAGE_TYPE: u32 = 6_005;
const COMMENT_STORAGE_MESSAGE_TYPE: u32 = 3_056;
const ANNOTATION_AUTHOR_MESSAGE_TYPE: u32 = 212;
const ANNOTATION_AUTHOR_STORAGE_MESSAGE_TYPE: u32 = 213;
const COMMENT_LIST_TYPE: u64 = 10;
const GENERATED_AUTHOR_NAME: &str = "litchi-iwa";
const MAX_COMMENT_TEXT_BYTES: usize = 64 * 1024 * 1024;

fn map_comment_decode_error(error: comment_codec::DecodeError) -> Fault {
    let Some(limit) = error.resource_limit() else {
        return Fault::Invalid;
    };
    match limit {
        comment_codec::DecodeLimit::Bytes { observed, maximum } => Fault::Limit {
            kind: LimitKind::WireBytes,
            observed: observed as u64,
            maximum: maximum as u64,
        },
        comment_codec::DecodeLimit::OutputBytes { observed, maximum } => Fault::Limit {
            kind: LimitKind::WireOutputBytes,
            observed: observed as u64,
            maximum: maximum as u64,
        },
        comment_codec::DecodeLimit::References { observed, maximum }
        | comment_codec::DecodeLimit::ReferenceBytes { observed, maximum } => Fault::Limit {
            kind: LimitKind::References,
            observed: observed as u64,
            maximum: maximum as u64,
        },
        comment_codec::DecodeLimit::Replies { observed, maximum } => Fault::Limit {
            kind: LimitKind::Replies,
            observed: observed as u64,
            maximum: maximum as u64,
        },
        comment_codec::DecodeLimit::Text { observed, maximum } => Fault::Limit {
            kind: LimitKind::TextBytes,
            observed: observed as u64,
            maximum: maximum as u64,
        },
        comment_codec::DecodeLimit::Fields { observed, maximum } => Fault::Limit {
            kind: LimitKind::WireFields,
            observed: observed as u64,
            maximum: maximum as u64,
        },
        comment_codec::DecodeLimit::Work { observed, maximum } => Fault::Limit {
            kind: LimitKind::WireWork,
            observed: observed as u64,
            maximum: maximum as u64,
        },
        comment_codec::DecodeLimit::Nesting { observed, maximum } => Fault::Limit {
            kind: LimitKind::WireNesting,
            observed: u64::from(observed),
            maximum: u64::from(maximum),
        },
        comment_codec::DecodeLimit::Allocations { observed, maximum } => Fault::Limit {
            kind: LimitKind::Allocations,
            observed: observed as u64,
            maximum: maximum as u64,
        },
        comment_codec::DecodeLimit::Scratch { observed, maximum } => Fault::Limit {
            kind: LimitKind::ScratchBytes,
            observed: observed as u64,
            maximum: maximum as u64,
        },
        comment_codec::DecodeLimit::Retained { observed, maximum } => Fault::Limit {
            kind: LimitKind::RetainedBytes,
            observed: observed as u64,
            maximum: maximum as u64,
        },
        _ => Fault::Invalid,
    }
}

fn map_table_decode_error(error: table_codec::DecodeError) -> Fault {
    let Some(limit) = error.resource_limit() else {
        return Fault::Invalid;
    };
    match limit {
        table_codec::DecodeLimit::Bytes { observed, maximum } => Fault::Limit {
            kind: LimitKind::PayloadBytes,
            observed: observed as u64,
            maximum: maximum as u64,
        },
        table_codec::DecodeLimit::References { observed, maximum } => Fault::Limit {
            kind: LimitKind::References,
            observed: observed as u64,
            maximum: maximum as u64,
        },
        table_codec::DecodeLimit::Text { observed, maximum } => Fault::Limit {
            kind: LimitKind::TextBytes,
            observed: observed as u64,
            maximum: maximum as u64,
        },
        table_codec::DecodeLimit::Fields { observed, maximum } => Fault::Limit {
            kind: LimitKind::WireFields,
            observed: observed as u64,
            maximum: maximum as u64,
        },
        table_codec::DecodeLimit::Work { observed, maximum } => Fault::Limit {
            kind: LimitKind::WireWork,
            observed: observed as u64,
            maximum: maximum as u64,
        },
        table_codec::DecodeLimit::Nesting { observed, maximum } => Fault::Limit {
            kind: LimitKind::WireNesting,
            observed: u64::from(observed),
            maximum: u64::from(maximum),
        },
        table_codec::DecodeLimit::Allocation { requested } => {
            Fault::Allocation { amount: requested }
        },
        table_codec::DecodeLimit::Retained { observed, maximum } => Fault::Limit {
            kind: LimitKind::RetainedBytes,
            observed: observed as u64,
            maximum: maximum as u64,
        },
        _ => Fault::Invalid,
    }
}

/// How a leaf chooses keys when it creates a comment-list entry.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ListKeys {
    /// Append after the list's current `next_list_id` (Pages/Keynote).
    Append,
    /// Reuse the canonical first key when creating the Numbers list.
    Reuse,
}

/// Format-specific constants that are safe to share with the native engine.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Profile {
    /// The table-model message type selected by the leaf.
    pub model_message_type: u32,
    /// The table-tile message type selected by the leaf.
    pub tile_message_type: u32,
    /// The table-data-list message type used by this owner.
    pub list_message_type: u32,
    /// The key policy for a newly created root entry.
    pub list_keys: ListKeys,
    /// Whether this native owner allocates a new author before its root.
    /// Keynote's author registry has historically used that order.
    pub author_before_root: bool,
    /// Whether a missing author registry gets a generated `litchi-iwa`
    /// author.  Pages and Keynote preserve their legacy authorless shape;
    /// Numbers materializes the generated author used by its focused route.
    pub generate_author_if_missing: bool,
    /// Whether this owner's native save keeps the three root preview members.
    /// Numbers invalidates them; Pages keeps its stale previews on this route.
    pub keep_previews: bool,
    /// Whether a newly-created author is appended to an existing author
    /// registry.  Keynote preserves its empty registry object and stores the
    /// author reference only on the root comment; Numbers uses the registry
    /// edge for its native shape.
    pub update_author_storage: bool,
}

impl Profile {
    /// Pages' body-table profile.  Pages keeps the shared append policy.
    pub const PAGES: Self = Self {
        model_message_type: 6_001,
        tile_message_type: 6_002,
        list_message_type: TABLE_DATA_LIST_MESSAGE_TYPE,
        list_keys: ListKeys::Append,
        author_before_root: false,
        generate_author_if_missing: false,
        keep_previews: true,
        update_author_storage: false,
    };

    /// Keynote's slide-table profile.
    pub const KEYNOTE: Self = Self {
        model_message_type: 6_001,
        tile_message_type: 6_002,
        list_message_type: TABLE_DATA_LIST_MESSAGE_TYPE,
        list_keys: ListKeys::Append,
        author_before_root: true,
        generate_author_if_missing: true,
        keep_previews: true,
        update_author_storage: false,
    };

    /// Numbers' table profile.  Its native list creation reuses key `2`.
    pub const NUMBERS: Self = Self {
        model_message_type: 6_001,
        tile_message_type: 6_002,
        list_message_type: TABLE_DATA_LIST_MESSAGE_TYPE,
        list_keys: ListKeys::Reuse,
        author_before_root: false,
        generate_author_if_missing: true,
        keep_previews: false,
        update_author_storage: true,
    };
}

/// Selector evidence supplied by one leaf after semantic and native
/// admission.  These are implementation inputs, never supported API values.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Selection {
    /// Component containing the table model.
    pub model_component: Box<str>,
    /// Component containing the selected tile.
    pub tile_component: Box<str>,
    /// Table-model object identifier.
    pub model_identifier: u64,
    /// Table-model message index.
    pub model_message_index: usize,
    /// Table-model message type.
    pub model_message_type: u32,
    /// Table-tile object identifier.
    pub tile_identifier: u64,
    /// Table-tile message index.
    pub tile_message_index: usize,
    /// Table-tile message type.
    pub tile_message_type: u32,
    /// Tile row index (within the selected tile).
    pub tile_row: u32,
    /// Zero-based table column.
    pub column: u32,
    /// Total table columns.
    pub table_columns: u32,
    /// Existing table-list key, when the selected cell has one.
    pub comment_key: Option<u32>,
    /// Existing table-list object identifier.
    pub list_identifier: Option<u64>,
    /// Existing comment-storage object identifier.
    pub root_identifier: Option<u64>,
}

/// One root-only semantic native operation.  Direct replies are deliberately
/// outside this first shared lifecycle and are refused by the facades until a
/// second slice ports their graph transitions.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum RootOperation {
    /// Replace or create the selected root text.
    Set(Box<str>),
    /// Detach and remove the selected root (and only an unshared empty thread).
    Clear,
}

/// A candidate produced by the shared native rewrite.
#[must_use = "the candidate must be semantically read back by the leaf facade"]
pub struct Mutation<P> {
    package: P,
    target: Arc<[u8]>,
    touched_components: usize,
    deleted_previews: usize,
}

impl<P> Mutation<P> {
    /// Borrow the reopened candidate package.
    #[must_use]
    pub const fn package(&self) -> &P {
        &self.package
    }

    /// Consume the native result into its candidate package.
    #[must_use]
    pub fn into_package(self) -> P {
        self.package
    }

    /// Borrow exact candidate bytes for an exact-source patch.
    #[must_use]
    pub fn target(&self) -> Arc<[u8]> {
        Arc::clone(&self.target)
    }

    /// Number of rewritten IWA components.
    #[must_use]
    pub const fn touched_components(&self) -> usize {
        self.touched_components
    }

    /// Number of stale root previews removed by publication.
    #[must_use]
    pub const fn deleted_previews(&self) -> usize {
        self.deleted_previews
    }
}

/// Execute a bounded root comment rewrite through the shared structural
/// working set.
pub fn rewrite_root<P: TableLeaf>(
    source: &P,
    selection: &Selection,
    operation: &RootOperation,
    profile: Profile,
) -> FaultResult<Mutation<P>> {
    validate_selection(source, selection, profile)?;
    let mut working = Working::new(source)?;
    // Admit the census maps before the one package-wide edge walk.  The
    // admission only counts catalog object slots; the graph itself is not
    // walked twice, and every per-edge growth remains fallible.
    let census_admission = structure::collect::Census::admit(&working)?;
    let census = structure::collect::Census::from_admission(&working, census_admission)?;
    let author_storage_component = author_storage_component(&working)?;
    let registry = metadata::registry(&working)?;
    let registry = registry.ok_or(Fault::Unsupported)?;
    if matches!(operation, RootOperation::Clear) && selection.root_identifier.is_none() {
        return Err(Fault::Invalid);
    }
    if matches!(operation, RootOperation::Set(_))
        && selection.root_identifier.is_some()
        && selection.comment_key.is_none()
    {
        return Err(Fault::Invalid);
    }

    let mut transition = metadata::Transition::default();
    let mut inserted = Vec::new();
    structure::reserve(&mut inserted, 4)?;
    let mut removed = Vec::new();
    structure::reserve(&mut removed, 3)?;
    match operation {
        RootOperation::Set(text) => {
            if text.len() > MAX_COMMENT_TEXT_BYTES {
                return Err(Fault::Limit {
                    kind: LimitKind::TextBytes,
                    observed: text.len() as u64,
                    maximum: MAX_COMMENT_TEXT_BYTES as u64,
                });
            }
            if let Some(root) = selection.root_identifier {
                ensure_unique_root(&working, &census, selection, root)?;
                validate_existing_comment_graph(
                    &working,
                    root,
                    selection.model_component.as_ref(),
                    author_storage_component.as_deref(),
                )?;
                rewrite_existing_text(&mut working, root, text)?;
            } else {
                if let Some(list) = selection.list_identifier {
                    // A new root still publishes into an existing comment
                    // list.  Validate every root already owned by that list
                    // before allocating or mutating anything: otherwise a
                    // malformed author/reply graph could be hidden by an
                    // otherwise independent create.
                    validate_existing_comment_list(
                        &working,
                        list,
                        selection.model_component.as_ref(),
                        author_storage_component.as_deref(),
                    )?;
                }
                let mut author =
                    plan_author(&working, selection, profile.generate_author_if_missing)?;
                let allocation_count = 1_usize
                    .saturating_add(if author.new_author { 1 } else { 0 })
                    .saturating_add(if author.new_storage { 1 } else { 0 });
                let mut allocated = working.allocate(allocation_count)?.into_iter();
                if profile.author_before_root && author.new_author {
                    author.author_identifier = allocated.next();
                }
                let root = allocated.next().ok_or(Fault::Invalid)?;
                if author.new_author && author.author_identifier.is_none() {
                    author.author_identifier = allocated.next();
                }
                if author.new_storage {
                    author.storage_identifier = allocated.next();
                }
                let root_component = selection
                    .list_identifier
                    .map(|list| working.locate(list))
                    .transpose()?
                    .flatten()
                    .unwrap_or_else(|| selection.model_component.to_string());
                create_root(
                    &mut working,
                    &root_component,
                    root,
                    text,
                    author.author_identifier,
                    &registry,
                    &mut transition,
                )?;
                inserted.push(root);
                if author.new_author {
                    create_author(&mut working, &author, &registry, &mut transition)?;
                    inserted.push(author.author_identifier.ok_or(Fault::Invalid)?);
                    if author.new_storage {
                        create_author_storage(&mut working, &author, &registry, &mut transition)?;
                        inserted.push(author.storage_identifier.ok_or(Fault::Invalid)?);
                    } else if profile.update_author_storage {
                        append_author_storage(&mut working, &author)?;
                    } else if author.new_author {
                        rewrite_author_storage_payload(&mut working, &author)?;
                    }
                }
                if let Some(list) =
                    attach_root(&mut working, selection, root, profile, &mut transition)?
                {
                    inserted.push(list);
                }
            }
        },
        RootOperation::Clear => {
            let root = selection.root_identifier.ok_or(Fault::Invalid)?;
            ensure_unique_root(&working, &census, selection, root)?;
            let author = ensure_empty_thread(&working, root)?;
            detach_root(&mut working, selection, root)?;
            removed.push(root);
            working.remove(root)?;
            if let Some(author) = author {
                remove_generated_author_if_unshared(
                    &mut working,
                    &census,
                    author,
                    root,
                    &mut removed,
                )?;
            }
            // The native clear releases the suffix when the removed root (or
            // its generated author) was the latest issued identifier.  Keep
            // the package watermark identical to the exact pre-edit source;
            // `metadata::apply` only lowers it for an explicit release.
            transition.release_suffix(&removed)?;
        },
    }

    let mut stamped_components = vec![
        selection.model_component.to_string(),
        selection.tile_component.to_string(),
    ];
    if let Some(identifier) = selection.list_identifier {
        stamped_components.push(working.locate(identifier)?.ok_or(Fault::Invalid)?);
    }
    if !matches!(operation, RootOperation::Clear) {
        if let Some(identifier) = selection.root_identifier {
            stamped_components.push(working.locate(identifier)?.ok_or(Fault::Invalid)?);
        }
    }
    for component in stamped_components {
        let component_id = registry.component_for_entry(&component)?.identifier;
        transition.stamp(component_id)?;
    }
    if let Some(component) = author_storage_component {
        let component_id = registry.component_for_entry(&component)?.identifier;
        transition.stamp(component_id)?;
    }

    register_insertions(&working, &registry, &inserted, &mut transition)?;
    if !removed.is_empty() {
        transition.remove_objects(&registry, &removed)?;
    }
    if let Some(token) = working.save_token()? {
        working.set_save_token(token.checked_add(1).ok_or(Fault::Limit {
            kind: LimitKind::References,
            observed: u64::MAX,
            maximum: u64::MAX,
        })?);
    }
    if profile.keep_previews {
        working.keep_previews();
    }
    let published = working.publish(&transition)?;
    let target: Arc<[u8]> = Arc::from(published.package.source_bytes());
    Ok(Mutation {
        package: published.package,
        target,
        touched_components: published.touched_components,
        deleted_previews: published.deleted_previews,
    })
}

fn validate_selection<P: TableLeaf>(
    source: &P,
    selection: &Selection,
    profile: Profile,
) -> FaultResult<()> {
    if selection.model_identifier == 0
        || selection.tile_identifier == 0
        || selection.table_columns == 0
        || selection.column >= selection.table_columns
        || selection.model_component.is_empty()
        || selection.tile_component.is_empty()
        || selection.model_message_index == usize::MAX
        || selection.tile_message_index == usize::MAX
    {
        return Err(Fault::Invalid);
    }
    if selection.model_message_type != profile.model_message_type
        || selection.tile_message_type != profile.tile_message_type
    {
        return Err(Fault::Unsupported);
    }
    let table_columns = usize::try_from(selection.table_columns).map_err(|_| Fault::Invalid)?;
    if table_columns > WireLimits::MAX_FIELDS {
        return Err(Fault::Limit {
            kind: LimitKind::WireFields,
            observed: table_columns as u64,
            maximum: WireLimits::MAX_FIELDS as u64,
        });
    }
    validate_object(
        source,
        selection.model_identifier,
        &selection.model_component,
        selection.model_message_index,
        profile.model_message_type,
    )?;
    validate_object(
        source,
        selection.tile_identifier,
        &selection.tile_component,
        selection.tile_message_index,
        profile.tile_message_type,
    )?;
    if let Some(list) = selection.list_identifier {
        validate_any_message(source, list, profile.list_message_type)?;
    } else if selection.comment_key.is_some() || selection.root_identifier.is_some() {
        return Err(Fault::Invalid);
    }
    if selection.comment_key.is_some() && selection.list_identifier.is_none() {
        return Err(Fault::Invalid);
    }
    if let Some(root) = selection.root_identifier {
        validate_any_message(source, root, COMMENT_STORAGE_MESSAGE_TYPE)?;
    }
    let _ = source.archive_limits();
    Ok(())
}

fn validate_any_message<P: TableLeaf>(
    source: &P,
    identifier: u64,
    message_type: u32,
) -> FaultResult<()> {
    let resolved = source
        .resolve_object(identifier)
        .map_err(|_| Fault::Invalid)?
        .ok_or(Fault::Invalid)?;
    if !resolved
        .messages
        .iter()
        .any(|message| message.type_ == message_type)
    {
        return Err(Fault::Unsupported);
    }
    Ok(())
}

fn validate_object<P: TableLeaf>(
    source: &P,
    identifier: u64,
    expected_component: &str,
    message_index: usize,
    message_type: u32,
) -> FaultResult<()> {
    let resolved = source
        .resolve_object(identifier)
        .map_err(|_| Fault::Invalid)?
        .ok_or(Fault::Invalid)?;
    let component = source
        .catalog()
        .get_index(resolved.component_index)
        .ok_or(Fault::Invalid)?;
    if component.name() != expected_component {
        return Err(Fault::Unsupported);
    }
    let message = resolved.messages.get(message_index).ok_or(Fault::Invalid)?;
    if message.type_ != message_type {
        return Err(Fault::Unsupported);
    }
    Ok(())
}

fn ensure_unique_root<P: TableLeaf>(
    working: &Working<'_, P>,
    census: &structure::collect::Census,
    selection: &Selection,
    root: u64,
) -> FaultResult<()> {
    if census.referrers(root).len() != 1 {
        return Err(Fault::Dependency);
    }
    let list = selection.list_identifier.ok_or(Fault::Invalid)?;
    let key = selection.comment_key.ok_or(Fault::Invalid)?;
    let entry = list_entry(working, list, key)?;
    if entry.storage != root || entry.ref_count != 1 {
        return Err(Fault::Dependency);
    }
    Ok(())
}

fn ensure_empty_thread<P: TableLeaf>(
    working: &Working<'_, P>,
    root: u64,
) -> FaultResult<Option<u64>> {
    let object = working.object(root)?;
    let storage_index = comment_storage_message_index(object)?;
    let source = object
        .messages
        .get(storage_index)
        .ok_or(Fault::Invalid)?
        .data
        .as_slice();
    let options = comment_codec::DecodeOptions::new(
        source.len().max(1),
        source.len().saturating_mul(4).max(1),
        source.len().saturating_mul(64).max(1),
        64,
        working.source().max_references(),
        MAX_COMMENT_TEXT_BYTES,
    );
    let (snapshot, report) =
        comment_codec::decode_comment_storage_archive_with_report(source, options)
            .map_err(map_comment_decode_error)?;
    if report.replies() != 0 {
        return Err(Fault::Dependency);
    }
    Ok(snapshot.author().map(|author| author.identifier()))
}

fn remove_generated_author_if_unshared<P: TableLeaf>(
    working: &mut Working<'_, P>,
    census: &structure::collect::Census,
    author: u64,
    removed_root: u64,
    removed: &mut Vec<u64>,
) -> FaultResult<()> {
    let remaining_referrers = census
        .referrers(author)
        .iter()
        .filter(|identifier| **identifier != removed_root)
        .count();
    if remaining_referrers != 0 && remaining_referrers != 1 {
        return Ok(());
    }
    let author_object = working.object(author)?;
    let mut author_data = None;
    for message in &author_object.messages {
        if message.type_ == ANNOTATION_AUTHOR_MESSAGE_TYPE {
            if author_data.replace(message.data.as_slice()).is_some() {
                return Err(Fault::Dependency);
            }
        }
    }
    let Some(author_data) = author_data else {
        return Err(Fault::Dependency);
    };
    let author_view = WireView::parse(author_data, WireLimits::DEFAULT).map_err(map_wire_error)?;
    let Some(name) = unique_field(&author_view, 1)? else {
        return Ok(());
    };
    if name.wire_type() != 2 || name.payload() != GENERATED_AUTHOR_NAME.as_bytes() {
        return Ok(());
    }

    let mut storage = None;
    for component in working.component_names() {
        for object in &working.archive(&component)?.objects {
            let Some(message) = object
                .messages
                .iter()
                .find(|message| message.type_ == ANNOTATION_AUTHOR_STORAGE_MESSAGE_TYPE)
            else {
                continue;
            };
            let references =
                author_storage_references(&message.data, working.source().max_references())?;
            if references == [author] {
                if storage
                    .replace((
                        component.clone(),
                        object.archive_info.identifier.ok_or(Fault::Invalid)?,
                    ))
                    .is_some()
                {
                    return Err(Fault::Dependency);
                }
            }
        }
    }
    if remaining_referrers != 0 && storage.is_none() {
        return Ok(());
    }
    if let Some((storage_component, storage_identifier)) = storage {
        replace_message(
            working,
            storage_identifier,
            ANNOTATION_AUTHOR_STORAGE_MESSAGE_TYPE,
            Vec::new(),
            storage_component,
        )?;
    }
    working.remove(author)?;
    removed.push(author);
    Ok(())
}

#[derive(Clone, Copy, Debug)]
struct Entry {
    key: u32,
    ref_count: u32,
    storage: u64,
}

fn list_entry<P: TableLeaf>(working: &Working<'_, P>, list: u64, key: u32) -> FaultResult<Entry> {
    let object = working.object(list)?;
    let message_index = comment_list_message_index(object, TABLE_DATA_LIST_MESSAGE_TYPE)?;
    let message = object.messages.get(message_index).ok_or(Fault::Invalid)?;
    let view =
        WireView::parse(message.data.as_slice(), WireLimits::DEFAULT).map_err(map_wire_error)?;
    let mut found = None;
    for field in view
        .fields()
        .filter(|field| field.number() == TABLE_LIST_ENTRY_FIELD)
    {
        let entry = decode_entry(field.payload())?;
        if entry.key == key && found.replace(entry).is_some() {
            return Err(Fault::Invalid);
        }
    }
    found.ok_or(Fault::Invalid)
}

/// Validate the existing roots in one comment list before adding a new root.
///
/// The shared census proves ownership of the archive headers, but it does not
/// interpret comment-storage payloads.  This deliberately narrow pass closes
/// that comment-specific gap without introducing a second package-wide graph
/// admission walk.
fn validate_existing_comment_list<P: TableLeaf>(
    working: &Working<'_, P>,
    list: u64,
    model_component: &str,
    author_storage_component: Option<&str>,
) -> FaultResult<()> {
    let object = working.object(list)?;
    let message_index = comment_list_message_index(object, TABLE_DATA_LIST_MESSAGE_TYPE)?;
    let message = object.messages.get(message_index).ok_or(Fault::Invalid)?;
    let view =
        WireView::parse(message.data.as_slice(), WireLimits::DEFAULT).map_err(map_wire_error)?;
    let mut roots = Vec::new();
    for field in view
        .fields()
        .filter(|field| field.number() == TABLE_LIST_ENTRY_FIELD)
    {
        let entry = decode_entry(field.payload())?;
        if !roots.contains(&entry.storage) {
            if roots.len() >= working.source().max_references() {
                return Err(Fault::Limit {
                    kind: LimitKind::References,
                    observed: roots.len().saturating_add(1) as u64,
                    maximum: working.source().max_references() as u64,
                });
            }
            structure::reserve(&mut roots, 1)?;
            roots.push(entry.storage);
        }
    }
    for root in roots {
        let root_component = working.locate(root)?.ok_or(Fault::Dependency)?;
        if root_component != model_component {
            return Err(Fault::Unsupported);
        }
        validate_existing_comment_graph(working, root, model_component, author_storage_component)?;
    }
    Ok(())
}

fn comment_list_message_index(object: &ArchiveObject, message_type: u32) -> FaultResult<usize> {
    let mut found = None;
    for (index, message) in object.messages.iter().enumerate() {
        if message.type_ != message_type {
            continue;
        }
        let view = WireView::parse(message.data.as_slice(), WireLimits::DEFAULT)
            .map_err(map_wire_error)?;
        let mut list_type = None;
        for field in view
            .fields()
            .filter(|field| field.number() == TABLE_LIST_TYPE_FIELD)
        {
            if list_type.replace(read_varint(field.payload())?).is_some() {
                return Err(Fault::Invalid);
            }
        }
        if list_type == Some(COMMENT_LIST_TYPE) && found.replace(index).is_some() {
            return Err(Fault::Invalid);
        }
    }
    found.ok_or(Fault::Unsupported)
}

fn decode_entry(data: &[u8]) -> FaultResult<Entry> {
    let view = WireView::parse(data, WireLimits::DEFAULT).map_err(map_wire_error)?;
    let mut key = None;
    let mut ref_count = None;
    let mut storage = None;
    for field in view.fields() {
        match field.number() {
            TABLE_ENTRY_KEY_FIELD => {
                if field.wire_type() != 0 || key.is_some() {
                    return Err(Fault::Invalid);
                }
                key = Some(read_varint(field.payload())? as u32);
            },
            TABLE_ENTRY_REFCOUNT_FIELD => {
                if field.wire_type() != 0 || ref_count.is_some() {
                    return Err(Fault::Invalid);
                }
                ref_count = Some(read_varint(field.payload())? as u32);
            },
            TABLE_ENTRY_STORAGE_FIELD => {
                if field.wire_type() != 2 || storage.is_some() {
                    return Err(Fault::Invalid);
                }
                let reference = WireView::parse(field.payload(), WireLimits::DEFAULT)
                    .map_err(map_wire_error)?;
                let mut id = None;
                for nested in reference.fields() {
                    if nested.number() != TABLE_REFERENCE_IDENTIFIER_FIELD
                        || nested.wire_type() != 0
                        || id.is_some()
                    {
                        return Err(Fault::Unsupported);
                    }
                    id = Some(read_varint(nested.payload())?);
                }
                storage = id;
            },
            _ => return Err(Fault::Unsupported),
        }
    }
    let entry = Entry {
        key: key.ok_or(Fault::Invalid)?,
        ref_count: ref_count.ok_or(Fault::Invalid)?,
        storage: storage.ok_or(Fault::Invalid)?,
    };
    if entry.key == 0 || entry.ref_count == 0 || entry.storage == 0 {
        return Err(Fault::Invalid);
    }
    Ok(entry)
}

fn read_varint(data: &[u8]) -> FaultResult<u64> {
    let (value, used) = decode_varint_from_bytes(data).map_err(|_| Fault::Invalid)?;
    if used != data.len() {
        return Err(Fault::Invalid);
    }
    Ok(value)
}

/// Return one schema field without allocating a collection proportional to
/// the complete wire view.  A duplicate singular field is malformed.
fn unique_field<'a>(
    view: &WireView<'a>,
    number: u32,
) -> FaultResult<Option<wire::WireFieldView<'a>>> {
    let mut found = None;
    for field in view.fields().filter(|field| field.number() == number) {
        if found.replace(field).is_some() {
            return Err(Fault::Invalid);
        }
    }
    Ok(found)
}

/// Resolve the one comment-storage payload owned by a root object.
///
/// A root comment has one 3056 payload.  Taking the first matching message
/// would make a malformed object ambiguous and could cause a rewrite to leave
/// a second, untouched comment behind.  Keep this invariant at the shared
/// graph boundary so every format facade gets the same refusal.
fn comment_storage_message_index(object: &ArchiveObject) -> FaultResult<usize> {
    let mut found = None;
    for (index, message) in object.messages.iter().enumerate() {
        if message.type_ != COMMENT_STORAGE_MESSAGE_TYPE {
            continue;
        }
        if found.replace(index).is_some() {
            return Err(Fault::Invalid);
        }
    }
    found.ok_or(Fault::Invalid)
}

fn rewrite_existing_text<P: TableLeaf>(
    working: &mut Working<'_, P>,
    root: u64,
    text: &str,
) -> FaultResult<()> {
    let component = working.locate(root)?.ok_or(Fault::Invalid)?;
    let storage_index = comment_storage_message_index(working.object(root)?)?;
    let source = working
        .object(root)?
        .messages
        .get(storage_index)
        .ok_or(Fault::Invalid)?
        .data
        .as_slice();
    let source = structure::copy_bytes(source)?;
    if text.len() > MAX_COMMENT_TEXT_BYTES {
        return Err(Fault::Limit {
            kind: LimitKind::TextBytes,
            observed: text.len() as u64,
            maximum: MAX_COMMENT_TEXT_BYTES as u64,
        });
    }
    let options = comment_codec::DecodeOptions::new(
        source.len().max(1),
        source.len().saturating_mul(4).max(1),
        source.len().saturating_mul(64).max(1),
        64,
        working.source().max_references(),
        MAX_COMMENT_TEXT_BYTES,
    );
    comment_codec::decode_comment_storage_archive_with_report(&source, options)
        .map_err(map_comment_decode_error)?;
    let view = WireView::parse(&source, WireLimits::DEFAULT).map_err(map_wire_error)?;
    let Some(text_field) = unique_field(&view, 1)? else {
        return Err(Fault::Invalid);
    };
    if text_field.wire_type() != 2 {
        return Err(Fault::Invalid);
    }
    text_field
        .validate_canonical_framing()
        .map_err(map_wire_error)?;
    let output = wire::patch_length_delimited_field(&source, 1, true, Some(text.as_bytes()))
        .map_err(map_wire_error)?;
    let limits = working.archive_limits();
    let archive = working.archive_mut(&component)?;
    let object = archive.object_mut(root).ok_or(Fault::Invalid)?;
    object
        .replace_message_preserving_header(
            storage_index,
            RawMessage {
                type_: COMMENT_STORAGE_MESSAGE_TYPE,
                data: output,
            },
            limits,
        )
        .map_err(map_core_error)
        .map(|_| ())
}

#[derive(Clone, Debug)]
struct AuthorStorageRoute {
    identifier: u64,
    component: String,
}

#[derive(Clone, Debug)]
struct AuthorPlan {
    author_identifier: Option<u64>,
    storage: Option<AuthorStorageRoute>,
    new_author: bool,
    new_storage: bool,
    storage_identifier: Option<u64>,
    storage_component: String,
}

fn author_storage_component<P: TableLeaf>(working: &Working<'_, P>) -> FaultResult<Option<String>> {
    let mut found = None;
    for component in working.component_names() {
        for object in &working.archive(&component)?.objects {
            if object
                .messages
                .iter()
                .any(|message| message.type_ == ANNOTATION_AUTHOR_STORAGE_MESSAGE_TYPE)
                && found.replace(component.clone()).is_some()
            {
                return Err(Fault::Dependency);
            }
        }
    }
    Ok(found)
}

fn plan_author<P: TableLeaf>(
    working: &Working<'_, P>,
    selection: &Selection,
    generate_author_if_missing: bool,
) -> FaultResult<AuthorPlan> {
    let mut generated = None;
    for component in working.component_names() {
        for object in &working.archive(&component)?.objects {
            let Some(identifier) = object.archive_info.identifier else {
                return Err(Fault::Invalid);
            };
            for message in object
                .messages
                .iter()
                .filter(|message| message.type_ == ANNOTATION_AUTHOR_MESSAGE_TYPE)
            {
                let view = WireView::parse(message.data.as_slice(), WireLimits::DEFAULT)
                    .map_err(map_wire_error)?;
                let mut name = None;
                for field in view.fields().filter(|field| field.number() == 1) {
                    if name.replace(field.payload()).is_some() {
                        return Err(Fault::Invalid);
                    }
                }
                if name
                    .and_then(|bytes| std::str::from_utf8(bytes).ok())
                    .is_some_and(|name| name == GENERATED_AUTHOR_NAME)
                {
                    if generated.replace(identifier).is_some() {
                        return Err(Fault::Dependency);
                    }
                }
            }
        }
    }
    if let Some(identifier) = generated {
        return Ok(AuthorPlan {
            author_identifier: Some(identifier),
            storage: None,
            new_author: false,
            new_storage: false,
            storage_identifier: None,
            storage_component: selection.model_component.to_string(),
        });
    }

    let mut storage = None;
    for component in working.component_names() {
        for object in &working.archive(&component)?.objects {
            let Some(identifier) = object.archive_info.identifier else {
                return Err(Fault::Invalid);
            };
            if object
                .messages
                .iter()
                .any(|message| message.type_ == ANNOTATION_AUTHOR_STORAGE_MESSAGE_TYPE)
            {
                if storage
                    .replace(AuthorStorageRoute {
                        identifier,
                        component: component.clone(),
                    })
                    .is_some()
                {
                    return Err(Fault::Dependency);
                }
            }
        }
    }
    let storage_component = storage.as_ref().map_or_else(
        || selection.model_component.to_string(),
        |route| route.component.clone(),
    );
    if let Some(route) = storage.as_ref() {
        let storage_object = working.object(route.identifier)?;
        let message_index = author_storage_message_index(storage_object)?;
        let source = storage_object
            .messages
            .get(message_index)
            .ok_or(Fault::Invalid)?
            .data
            .as_slice();
        let references = author_storage_references(source, working.source().max_references())?;
        for author in &references {
            let author_object = working.object(*author)?;
            if author_object
                .messages
                .iter()
                .filter(|message| message.type_ == ANNOTATION_AUTHOR_MESSAGE_TYPE)
                .count()
                != 1
            {
                return Err(Fault::Dependency);
            }
        }
        if let Some(&author_identifier) = references.first() {
            // Native Pages/Keynote creation reuses the first registered
            // author.  Numbers fixtures with an empty registry take the
            // generated-author path below; an existing registry is reused by
            // all profiles so a create/clear cycle cannot manufacture a
            // second identity beside the host's author.
            return Ok(AuthorPlan {
                author_identifier: Some(author_identifier),
                storage: storage.clone(),
                new_author: false,
                new_storage: false,
                storage_identifier: Some(route.identifier),
                storage_component,
            });
        }
        if !generate_author_if_missing {
            return Ok(AuthorPlan {
                author_identifier: None,
                storage: storage.clone(),
                new_author: false,
                new_storage: false,
                storage_identifier: Some(route.identifier),
                storage_component,
            });
        }
    } else if !generate_author_if_missing {
        return Ok(AuthorPlan {
            author_identifier: None,
            storage: None,
            new_author: false,
            new_storage: false,
            storage_identifier: None,
            storage_component,
        });
    }
    let new_storage = storage.is_none();
    Ok(AuthorPlan {
        author_identifier: None,
        storage_identifier: storage.as_ref().map(|route| route.identifier),
        storage,
        new_author: true,
        new_storage,
        storage_component,
    })
}

fn validate_existing_comment_graph<P: TableLeaf>(
    working: &Working<'_, P>,
    root: u64,
    model_component: &str,
    author_storage_component: Option<&str>,
) -> FaultResult<()> {
    let component = working.locate(root)?.ok_or(Fault::Invalid)?;
    let object = working.object(root)?;
    let storage_count = object
        .messages
        .iter()
        .filter(|message| message.type_ == COMMENT_STORAGE_MESSAGE_TYPE)
        .count();
    if storage_count != 1 {
        return Err(Fault::Invalid);
    }
    for message in object
        .messages
        .iter()
        .filter(|message| message.type_ == COMMENT_STORAGE_MESSAGE_TYPE)
    {
        let options = comment_codec::DecodeOptions::new(
            message.data.len().max(1),
            message.data.len().saturating_mul(4).max(1),
            message.data.len().saturating_mul(64).max(1),
            64,
            working.source().max_references(),
            MAX_COMMENT_TEXT_BYTES,
        );
        let mut reply_ids = Vec::new();
        let mut allocation_failed = None;
        let mut reference_limit = false;
        let (snapshot, _) = comment_codec::visit_comment_storage_replies(
            message.data.as_slice(),
            options,
            &mut |reply| {
                if allocation_failed.is_none() && !reference_limit {
                    if reply_ids.len() >= working.source().max_references() {
                        reference_limit = true;
                        return Ok(());
                    }
                    if reply_ids.try_reserve(1).is_err() {
                        allocation_failed = Some(reply_ids.len().saturating_add(1));
                    } else {
                        reply_ids.push(reply.identifier());
                    }
                }
                Ok(())
            },
        )
        .map_err(map_comment_decode_error)?;
        if let Some(amount) = allocation_failed {
            return Err(Fault::Allocation { amount });
        }
        if reference_limit {
            return Err(Fault::Limit {
                kind: LimitKind::Replies,
                observed: reply_ids.len().saturating_add(1) as u64,
                maximum: working.source().max_references() as u64,
            });
        }
        for reply in reply_ids {
            let reply_component = working.locate(reply)?.ok_or(Fault::Dependency)?;
            if reply_component != component && component == model_component {
                return Err(Fault::Unsupported);
            }
        }
        if let Some(author) = snapshot.author() {
            let author_component = working
                .locate(author.identifier())?
                .ok_or(Fault::Dependency)?;
            if author_component != component
                && component == model_component
                && author_storage_component != Some(author_component.as_str())
            {
                return Err(Fault::Unsupported);
            }
        }
    }
    Ok(())
}

fn create_root<P: TableLeaf>(
    working: &mut Working<'_, P>,
    component: &str,
    identifier: u64,
    text: &str,
    author_identifier: Option<u64>,
    registry: &metadata::Registry,
    transition: &mut metadata::Transition,
) -> FaultResult<()> {
    let seed = metadata::identity_seed(working, identifier)?;
    let (lower, upper) = uuid_for(seed, identifier);
    let write = match author_identifier {
        Some(author) => CommentStorageLeafWrite::new(
            text,
            DateSnapshot::from_bits(0.0_f64.to_bits()),
            author,
            UuidSnapshot::from_parts(lower, upper),
        ),
        None => CommentStorageLeafWrite::without_author(
            text,
            DateSnapshot::from_bits(0.0_f64.to_bits()),
            UuidSnapshot::from_parts(lower, upper),
        ),
    };
    let prepared = comment_codec::prepare_comment_storage_leaf_write(
        write,
        comment_codec::DecodeOptions::new(
            text.len().saturating_add(256).max(1),
            text.len().saturating_add(256).max(1),
            text.len().saturating_mul(64).saturating_add(4096).max(1),
            64,
            working.source().max_references(),
            MAX_COMMENT_TEXT_BYTES,
        ),
    )
    .map_err(map_comment_decode_error)?;
    // Preparation computes the exact source/candidate fields, work,
    // references, scratch, retained bytes, and allocation count.  Replay
    // those requirements before execute so the writer cannot silently run
    // under an unbounded ceiling.
    let requirements = prepared.requirements();
    let payload = prepared
        .execute(requirements.exact_limits())
        .map_err(map_comment_decode_error)?
        .into_bytes();
    let limits = working.archive_limits();
    let mut object = ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_: COMMENT_STORAGE_MESSAGE_TYPE,
            data: payload,
        }],
        limits,
    )
    .map_err(map_core_error)?;
    let info = object
        .archive_info
        .message_infos
        .first_mut()
        .ok_or(Fault::Invalid)?;
    if let Some(author_identifier) = author_identifier {
        info.object_references.push(author_identifier);
        let mut field =
            FieldInfo::new(FieldPath { path: vec![3] }, limits).map_err(map_core_error)?;
        field.r#type = Some(FieldType::ObjectReference);
        field.object_references.push(author_identifier);
        info.field_infos.push(field);
    }
    working.insert(component, object)?;
    let metadata_component = registry.component_for_entry(component)?;
    transition.register_uuid(registry, metadata_component.identifier, identifier, seed)?;
    Ok(())
}

fn create_author<P: TableLeaf>(
    working: &mut Working<'_, P>,
    plan: &AuthorPlan,
    registry: &metadata::Registry,
    transition: &mut metadata::Transition,
) -> FaultResult<()> {
    let identifier = plan.author_identifier.ok_or(Fault::Invalid)?;
    let mut payload = Vec::new();
    wire::append_length_delimited_field(
        &mut payload,
        1,
        GENERATED_AUTHOR_NAME.as_bytes(),
        WireLimits::DEFAULT,
    )
    .map_err(map_wire_error)?;
    let mut color = Vec::new();
    wire::append_varint_field(&mut color, 1, 1).map_err(map_wire_error)?;
    append_fixed32_field(&mut color, 3, 0.368_627_46_f32.to_bits())?;
    append_fixed32_field(&mut color, 4, 0.568_627_5_f32.to_bits())?;
    append_fixed32_field(&mut color, 5, 0.937_254_9_f32.to_bits())?;
    append_fixed32_field(&mut color, 6, 1.0_f32.to_bits())?;
    wire::append_varint_field(&mut color, 12, 1).map_err(map_wire_error)?;
    wire::append_length_delimited_field(&mut payload, 2, &color, WireLimits::DEFAULT)
        .map_err(map_wire_error)?;
    wire::append_varint_field(&mut payload, 4, 0).map_err(map_wire_error)?;
    let limits = working.archive_limits();
    let mut object = ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_: ANNOTATION_AUTHOR_MESSAGE_TYPE,
            data: payload,
        }],
        limits,
    )
    .map_err(map_core_error)?;
    let info = object
        .archive_info
        .message_infos
        .first_mut()
        .ok_or(Fault::Invalid)?;
    info.field_infos
        .push(FieldInfo::new(FieldPath { path: vec![4] }, limits).map_err(map_core_error)?);
    working.insert(&plan.storage_component, object)?;
    let component = registry.component_for_entry(&plan.storage_component)?;
    let seed = metadata::identity_seed(working, identifier)?;
    transition.register_uuid(registry, component.identifier, identifier, seed)?;
    Ok(())
}

fn append_fixed32_field(output: &mut Vec<u8>, number: u32, value: u32) -> FaultResult<()> {
    if number == 0 {
        return Err(Fault::Invalid);
    }
    litchi_iwa_common::encode_varint_into(output, (u64::from(number) << 3) | 5);
    output.extend_from_slice(&value.to_le_bytes());
    Ok(())
}

fn create_author_storage<P: TableLeaf>(
    working: &mut Working<'_, P>,
    plan: &AuthorPlan,
    registry: &metadata::Registry,
    transition: &mut metadata::Transition,
) -> FaultResult<()> {
    let identifier = plan.storage_identifier.ok_or(Fault::Invalid)?;
    let author = plan.author_identifier.ok_or(Fault::Invalid)?;
    let reference = varint_field(1, author)?;
    let mut payload = Vec::new();
    wire::append_length_delimited_field(&mut payload, 1, &reference, WireLimits::DEFAULT)
        .map_err(map_wire_error)?;
    let limits = working.archive_limits();
    let mut object = ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_: ANNOTATION_AUTHOR_STORAGE_MESSAGE_TYPE,
            data: payload,
        }],
        limits,
    )
    .map_err(map_core_error)?;
    let info = object
        .archive_info
        .message_infos
        .first_mut()
        .ok_or(Fault::Invalid)?;
    info.object_references.push(author);
    let mut field = FieldInfo::new(FieldPath { path: vec![1] }, limits).map_err(map_core_error)?;
    field.r#type = Some(FieldType::ObjectReference);
    field.object_references.push(author);
    info.field_infos.push(field);
    working.insert(&plan.storage_component, object)?;
    let component = registry.component_for_entry(&plan.storage_component)?;
    let seed = metadata::identity_seed(working, identifier)?;
    transition.register_uuid(registry, component.identifier, identifier, seed)?;
    Ok(())
}

/// Re-encode an existing empty author registry without adding an archive
/// header reference.  Keynote's native save keeps the legacy registry's
/// header metadata empty while its wire payload carries the newly-created
/// author reference.
fn rewrite_author_storage_payload<P: TableLeaf>(
    working: &mut Working<'_, P>,
    plan: &AuthorPlan,
) -> FaultResult<()> {
    let identifier = plan.storage_identifier.ok_or(Fault::Invalid)?;
    let author = plan.author_identifier.ok_or(Fault::Invalid)?;
    let reference = varint_field(1, author)?;
    let mut payload = Vec::new();
    wire::append_length_delimited_field(&mut payload, 1, &reference, WireLimits::DEFAULT)
        .map_err(map_wire_error)?;
    let component = plan.storage_component.clone();
    let limits = working.archive_limits();
    let archive = working.archive_mut(&component)?;
    let object = archive.object_mut(identifier).ok_or(Fault::Invalid)?;
    let index = object
        .messages
        .iter()
        .position(|message| message.type_ == ANNOTATION_AUTHOR_STORAGE_MESSAGE_TYPE)
        .ok_or(Fault::Invalid)?;
    object
        .replace_message_preserving_header(
            index,
            RawMessage {
                type_: ANNOTATION_AUTHOR_STORAGE_MESSAGE_TYPE,
                data: payload,
            },
            limits,
        )
        .map_err(map_core_error)
        .map(|_| ())
}

fn author_storage_message_index(object: &ArchiveObject) -> FaultResult<usize> {
    Ok(structure::unique_message(
        object,
        &[ANNOTATION_AUTHOR_STORAGE_MESSAGE_TYPE],
    )?)
}

fn author_storage_references(data: &[u8], max_references: usize) -> FaultResult<Vec<u64>> {
    let view = WireView::parse(data, WireLimits::DEFAULT).map_err(map_wire_error)?;
    let mut references = Vec::new();
    for field in view.fields() {
        if field.number() != 1 || field.wire_type() != 2 {
            return Err(Fault::Unsupported);
        }
        let nested =
            WireView::parse(field.payload(), WireLimits::DEFAULT).map_err(map_wire_error)?;
        let mut identifier = None;
        for inner in nested.fields() {
            if inner.number() != 1 || inner.wire_type() != 0 || identifier.is_some() {
                return Err(Fault::Unsupported);
            }
            identifier = Some(read_varint(inner.payload())?);
        }
        let identifier = identifier.ok_or(Fault::Invalid)?;
        if identifier == 0 || references.contains(&identifier) {
            return Err(Fault::Dependency);
        }
        if references.len() >= max_references {
            return Err(Fault::Limit {
                kind: LimitKind::References,
                observed: references.len().saturating_add(1) as u64,
                maximum: max_references as u64,
            });
        }
        structure::reserve(&mut references, 1)?;
        references.push(identifier);
    }
    Ok(references)
}

fn append_author_storage<P: TableLeaf>(
    working: &mut Working<'_, P>,
    plan: &AuthorPlan,
) -> FaultResult<()> {
    let storage = plan.storage.as_ref().ok_or(Fault::Invalid)?;
    let author = plan.author_identifier.ok_or(Fault::Invalid)?;
    let object = working.object(storage.identifier)?;
    let message_index = author_storage_message_index(object)?;
    let source = object
        .messages
        .get(message_index)
        .ok_or(Fault::Invalid)?
        .data
        .as_slice();
    let source = structure::copy_bytes(source)?;
    let present = author_storage_references(&source, working.source().max_references())?;
    if present.contains(&author) {
        return Err(Fault::Dependency);
    }
    let reference = varint_field(1, author)?;
    let mut output = source;
    wire::append_length_delimited_field(&mut output, 1, &reference, WireLimits::DEFAULT)
        .map_err(map_wire_error)?;
    let limits = working.archive_limits();
    let archive = working.archive_mut(&storage.component)?;
    let object = archive
        .object_mut(storage.identifier)
        .ok_or(Fault::Invalid)?;
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(Fault::Invalid)?
        .clone();
    let mut after = Vec::new();
    structure::reserve(&mut after, present.len().saturating_add(1))?;
    after.extend_from_slice(&present);
    after.push(author);
    if present.is_empty() {
        if !info.object_references.is_empty() || !info.field_infos.is_empty() {
            return Err(Fault::Dependency);
        }
        object
            .replace_message_preserving_header(
                message_index,
                RawMessage {
                    type_: ANNOTATION_AUTHOR_STORAGE_MESSAGE_TYPE,
                    data: output,
                },
                limits,
            )
            .map_err(map_core_error)?;
        let info = object
            .archive_info
            .message_infos
            .get_mut(message_index)
            .ok_or(Fault::Invalid)?;
        info.object_references.push(author);
        let mut field =
            FieldInfo::new(FieldPath { path: vec![1] }, limits).map_err(map_core_error)?;
        field.r#type = Some(FieldType::ObjectReference);
        field.object_references.push(author);
        info.field_infos.push(field);
        return Ok(());
    }
    if info.object_references == present && info.field_infos.is_empty() {
        object
            .replace_message_preserving_header(
                message_index,
                RawMessage {
                    type_: ANNOTATION_AUTHOR_STORAGE_MESSAGE_TYPE,
                    data: output,
                },
                limits,
            )
            .map_err(map_core_error)
            .map(|_| ())?;
        return Ok(());
    }
    let [field] = info.field_infos.as_slice() else {
        return Err(Fault::Dependency);
    };
    if info.object_references != present
        || field.path.path != [1]
        || field.object_references != present
    {
        return Err(Fault::Dependency);
    }
    object
        .replace_message_transitioning_object_references_preserving_header(
            message_index,
            RawMessage {
                type_: ANNOTATION_AUTHOR_STORAGE_MESSAGE_TYPE,
                data: output,
            },
            ObjectReferenceTransition {
                aggregate_before: present.as_slice(),
                aggregate_after: after.as_slice(),
                fields: &[FieldObjectReferenceTransition {
                    field_info_index: 0,
                    expected_path: &[1],
                    before: present.as_slice(),
                    after: after.as_slice(),
                }],
            },
            limits,
        )
        .map_err(map_core_error)
        .map(|_| ())
}

fn uuid_for(seed: u64, identifier: u64) -> (u64, u64) {
    let state = seed ^ identifier.rotate_left(17);
    let lower = metadata::mix(state);
    let upper = metadata::mix(lower ^ 0x6c69_7463_6869_0301);
    (lower, upper)
}

fn attach_root<P: TableLeaf>(
    working: &mut Working<'_, P>,
    selection: &Selection,
    root: u64,
    profile: Profile,
    transition: &mut metadata::Transition,
) -> FaultResult<Option<u64>> {
    let (key, new_list) = if let Some(list) = selection.list_identifier {
        let key = append_list(working, list, root, profile.list_keys)?;
        (key, None)
    } else {
        let list = create_list(working, selection, root, transition)?;
        (
            match profile.list_keys {
                ListKeys::Append | ListKeys::Reuse => 2,
            },
            Some(list),
        )
    };
    rewrite_tile(working, selection, None, Some(key))?;
    Ok(new_list)
}

fn detach_root<P: TableLeaf>(
    working: &mut Working<'_, P>,
    selection: &Selection,
    root: u64,
) -> FaultResult<()> {
    let list = selection.list_identifier.ok_or(Fault::Invalid)?;
    let key = selection.comment_key.ok_or(Fault::Invalid)?;
    let entry = list_entry(working, list, key)?;
    if entry.storage != root || entry.ref_count != 1 {
        return Err(Fault::Dependency);
    }
    rewrite_tile(working, selection, Some(key), None)?;
    remove_list_entry(working, list, key, root)
}

fn append_list<P: TableLeaf>(
    working: &mut Working<'_, P>,
    list: u64,
    root: u64,
    policy: ListKeys,
) -> FaultResult<u32> {
    let component = working.locate(list)?.ok_or(Fault::Invalid)?;
    let object = working.object(list)?;
    let message_index = comment_list_message_index(object, TABLE_DATA_LIST_MESSAGE_TYPE)?;
    let source = object
        .messages
        .get(message_index)
        .ok_or(Fault::Invalid)?
        .data
        .as_slice();
    let source = structure::copy_bytes(source)?;
    let view = WireView::parse(&source, WireLimits::DEFAULT).map_err(map_wire_error)?;
    let next = unique_field(&view, TABLE_LIST_NEXT_ID_FIELD)?
        .ok_or(Fault::Invalid)
        .and_then(|field| read_varint(field.payload()))?;
    let key = match policy {
        ListKeys::Append => u32::try_from(next.max(2)).map_err(|_| Fault::Invalid)?,
        ListKeys::Reuse => 2,
    };
    if view
        .fields()
        .filter(|field| field.number() == TABLE_LIST_ENTRY_FIELD)
        .any(|field| decode_entry(field.payload()).is_ok_and(|entry| entry.key == key))
    {
        return Err(Fault::Dependency);
    }
    let reference = varint_field(TABLE_REFERENCE_IDENTIFIER_FIELD, root)?;
    let mut entry = Vec::new();
    wire::append_varint_field(&mut entry, TABLE_ENTRY_KEY_FIELD, u64::from(key))
        .map_err(map_wire_error)?;
    wire::append_varint_field(&mut entry, TABLE_ENTRY_REFCOUNT_FIELD, 1).map_err(map_wire_error)?;
    wire::append_length_delimited_field(
        &mut entry,
        TABLE_ENTRY_STORAGE_FIELD,
        &reference,
        WireLimits::DEFAULT,
    )
    .map_err(map_wire_error)?;
    let mut output = source;
    wire::append_length_delimited_field(
        &mut output,
        TABLE_LIST_ENTRY_FIELD,
        &entry,
        WireLimits::DEFAULT,
    )
    .map_err(map_wire_error)?;
    let next_key = u64::from(key.checked_add(1).ok_or(Fault::Invalid)?);
    output = wire::patch_varint_field(&output, TABLE_LIST_NEXT_ID_FIELD, true, Some(next_key))
        .map_err(map_wire_error)?;
    replace_message_at(working, list, message_index, output, component)?;
    update_list_metadata(working, list, key, root, true)?;
    Ok(key)
}

fn remove_list_entry<P: TableLeaf>(
    working: &mut Working<'_, P>,
    list: u64,
    key: u32,
    root: u64,
) -> FaultResult<()> {
    let component = working.locate(list)?.ok_or(Fault::Invalid)?;
    let object = working.object(list)?;
    let message_index = comment_list_message_index(object, TABLE_DATA_LIST_MESSAGE_TYPE)?;
    let source = object
        .messages
        .get(message_index)
        .ok_or(Fault::Invalid)?
        .data
        .as_slice();
    let source = structure::copy_bytes(source)?;
    let mut matched = false;
    // The generic transform callback cannot express omission, so rebuild the
    // list while preserving every non-entry field and unknown entry fields
    // after strict validation.
    let view = WireView::parse(&source, WireLimits::DEFAULT).map_err(map_wire_error)?;
    let mut rebuilt = Vec::new();
    structure::reserve(&mut rebuilt, source.len())?;
    let mut remaining = false;
    for field in view.fields() {
        if field.number() != TABLE_LIST_ENTRY_FIELD {
            rebuilt.extend_from_slice(field.raw());
            continue;
        }
        let facts = decode_entry(field.payload())?;
        if facts.key == key {
            if facts.storage != root || facts.ref_count != 1 || matched {
                return Err(Fault::Verification);
            }
            matched = true;
            continue;
        }
        remaining = true;
        rebuilt.extend_from_slice(field.raw());
    }
    if !matched {
        return Err(Fault::Verification);
    }
    if !remaining {
        rebuilt = wire::patch_varint_field(&rebuilt, TABLE_LIST_NEXT_ID_FIELD, true, Some(2))
            .map_err(map_wire_error)?;
    }
    replace_message_at(working, list, message_index, rebuilt, component)?;
    update_list_metadata(working, list, key, root, false)
}

fn update_list_metadata<P: TableLeaf>(
    working: &mut Working<'_, P>,
    list: u64,
    key: u32,
    root: u64,
    adding: bool,
) -> FaultResult<()> {
    let component = working.locate(list)?.ok_or(Fault::Invalid)?;
    let message_index =
        comment_list_message_index(working.object(list)?, TABLE_DATA_LIST_MESSAGE_TYPE)?;
    let limits = working.archive_limits();
    let archive = working.archive_mut(&component)?;
    let object = archive.object_mut(list).ok_or(Fault::Invalid)?;
    let info = object
        .archive_info
        .message_infos
        .get_mut(message_index)
        .ok_or(Fault::Invalid)?;
    let occurrences = info
        .object_references
        .iter()
        .filter(|identifier| **identifier == root)
        .count();
    let path = vec![TABLE_LIST_ENTRY_FIELD, key];
    let mut field_index = None;
    for (index, field) in info.field_infos.iter().enumerate() {
        if field.path.path == path && field_index.replace(index).is_some() {
            return Err(Fault::Dependency);
        }
    }
    // Older host-authored table-list headers carry only the aggregate object
    // references and no per-field inventory.  The wire list above is still
    // unambiguous (the selected key/storage/refcount were checked), so keep
    // that header shape losslessly instead of manufacturing a field path the
    // source never had.
    if info.field_infos.is_empty() {
        if adding {
            if occurrences != 0 {
                return Err(Fault::Dependency);
            }
            info.object_references.push(root);
        } else {
            if occurrences != 1 {
                return Err(Fault::Dependency);
            }
            info.object_references
                .retain(|identifier| *identifier != root);
        }
        return Ok(());
    }
    if adding {
        if occurrences != 0 || field_index.is_some() {
            return Err(Fault::Dependency);
        }
        info.object_references.push(root);
        let mut field = FieldInfo::new(FieldPath { path }, limits).map_err(map_core_error)?;
        field.r#type = Some(FieldType::ObjectReference);
        field.object_references.push(root);
        info.field_infos.push(field);
    } else {
        if occurrences != 1 || field_index.is_none() {
            return Err(Fault::Dependency);
        }
        info.object_references
            .retain(|identifier| *identifier != root);
        let index = field_index.ok_or(Fault::Invalid)?;
        let field = info.field_infos.get(index).ok_or(Fault::Invalid)?;
        if field.object_references.as_slice() != [root] {
            return Err(Fault::Dependency);
        }
        info.field_infos.remove(index);
    }
    Ok(())
}

fn create_list<P: TableLeaf>(
    working: &mut Working<'_, P>,
    selection: &Selection,
    root: u64,
    transition: &mut metadata::Transition,
) -> FaultResult<u64> {
    let identifier = working
        .allocate(1)?
        .into_iter()
        .next()
        .ok_or(Fault::Invalid)?;
    let reference = varint_field(TABLE_REFERENCE_IDENTIFIER_FIELD, root)?;
    let mut entry = Vec::new();
    wire::append_varint_field(&mut entry, TABLE_ENTRY_KEY_FIELD, 2).map_err(map_wire_error)?;
    wire::append_varint_field(&mut entry, TABLE_ENTRY_REFCOUNT_FIELD, 1).map_err(map_wire_error)?;
    wire::append_length_delimited_field(
        &mut entry,
        TABLE_ENTRY_STORAGE_FIELD,
        &reference,
        WireLimits::DEFAULT,
    )
    .map_err(map_wire_error)?;
    let mut payload = Vec::new();
    wire::append_varint_field(&mut payload, TABLE_LIST_TYPE_FIELD, COMMENT_LIST_TYPE)
        .map_err(map_wire_error)?;
    wire::append_varint_field(&mut payload, TABLE_LIST_NEXT_ID_FIELD, 3).map_err(map_wire_error)?;
    wire::append_length_delimited_field(
        &mut payload,
        TABLE_LIST_ENTRY_FIELD,
        &entry,
        WireLimits::DEFAULT,
    )
    .map_err(map_wire_error)?;
    let limits = working.archive_limits();
    let mut object = ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_: TABLE_DATA_LIST_MESSAGE_TYPE,
            data: payload,
        }],
        limits,
    )
    .map_err(map_core_error)?;
    let info = object
        .archive_info
        .message_infos
        .get_mut(0)
        .ok_or(Fault::Invalid)?;
    info.object_references.push(root);
    let mut field = FieldInfo::new(
        FieldPath {
            path: vec![TABLE_LIST_ENTRY_FIELD, 2],
        },
        limits,
    )
    .map_err(map_core_error)?;
    field.r#type = Some(FieldType::ObjectReference);
    field.object_references.push(root);
    info.field_infos.push(field);
    let model_source = working
        .object(selection.model_identifier)?
        .messages
        .get(selection.model_message_index)
        .ok_or(Fault::Invalid)?
        .data
        .as_slice();
    let model_source = structure::copy_bytes(model_source)?;
    let options = table_codec::DecodeOptions::new(
        model_source.len().saturating_add(64).max(1),
        model_source.len().saturating_add(64).max(1),
        model_source
            .len()
            .saturating_add(64)
            .saturating_mul(32)
            .max(1),
        64,
        working.source().max_references(),
        MAX_COMMENT_TEXT_BYTES,
    );
    let prepared = table_codec::prepare_table_model_comment_storage_table_rewrite(
        &model_source,
        table_codec::ControlCellSpecTableReferenceEdit::insert(identifier),
        options,
    )
    .map_err(map_table_decode_error)?;
    let requirements = prepared.requirements();
    let data = prepared
        .execute(requirements.exact_limits())
        .map_err(map_table_decode_error)?
        .0;
    let model_component = selection.model_component.as_ref();
    let archive = working.archive_mut(model_component)?;
    let model = archive
        .object_mut(selection.model_identifier)
        .ok_or(Fault::Invalid)?;
    let source_info = model
        .archive_info
        .message_infos
        .get(selection.model_message_index)
        .ok_or(Fault::Invalid)?;
    let mut before = Vec::new();
    structure::reserve(&mut before, source_info.object_references.len())?;
    before.extend_from_slice(&source_info.object_references);
    if before.contains(&identifier) {
        return Err(Fault::Invalid);
    }
    let mut after = Vec::new();
    structure::reserve(&mut after, before.len().saturating_add(1))?;
    after.extend_from_slice(&before);
    after.push(identifier);
    let source_fields = source_info.field_infos.as_slice();
    let mut field_data = Vec::new();
    structure::reserve(&mut field_data, source_fields.len())?;
    for field in source_fields {
        let mut path = Vec::new();
        structure::reserve(&mut path, field.path.path.len())?;
        path.extend_from_slice(&field.path.path);
        let mut references = Vec::new();
        structure::reserve(&mut references, field.object_references.len())?;
        references.extend_from_slice(&field.object_references);
        field_data.push((path, references));
    }
    let mut fields = Vec::new();
    structure::reserve(&mut fields, field_data.len())?;
    for (field_info_index, (path, references)) in field_data.iter().enumerate() {
        fields.push(FieldObjectReferenceTransition {
            field_info_index,
            expected_path: path.as_slice(),
            before: references.as_slice(),
            after: references.as_slice(),
        });
    }
    model
        .replace_message_transitioning_object_references_preserving_header(
            selection.model_message_index,
            RawMessage {
                type_: selection.model_message_type,
                data,
            },
            ObjectReferenceTransition {
                aggregate_before: before.as_slice(),
                aggregate_after: after.as_slice(),
                fields: fields.as_slice(),
            },
            limits,
        )
        .map_err(map_core_error)?;
    let info = model
        .archive_info
        .message_infos
        .get_mut(selection.model_message_index)
        .ok_or(Fault::Invalid)?;
    let mut field = FieldInfo::new(
        FieldPath {
            path: TABLE_MODEL_COMMENT_PATH.to_vec(),
        },
        limits,
    )
    .map_err(map_core_error)?;
    field.r#type = Some(FieldType::ObjectReference);
    field.object_references = vec![identifier];
    info.field_infos.push(field);
    working.insert(model_component, object)?;
    let registry = metadata::registry(working)?.ok_or(Fault::Unsupported)?;
    let component = registry.component_for_entry(model_component)?.identifier;
    let seed = metadata::identity_seed(working, identifier)?;
    transition.register_uuid(&registry, component, identifier, seed)?;
    Ok(identifier)
}

fn replace_message<P: TableLeaf>(
    working: &mut Working<'_, P>,
    identifier: u64,
    message_type: u32,
    data: Vec<u8>,
    component: String,
) -> FaultResult<()> {
    let limits = working.archive_limits();
    let archive = working.archive_mut(&component)?;
    let object = archive.object_mut(identifier).ok_or(Fault::Invalid)?;
    let index = object
        .messages
        .iter()
        .position(|message| message.type_ == message_type)
        .ok_or(Fault::Invalid)?;
    replace_message_at_inner(object, index, message_type, data, limits)
}

fn replace_message_at<P: TableLeaf>(
    working: &mut Working<'_, P>,
    identifier: u64,
    message_index: usize,
    data: Vec<u8>,
    component: String,
) -> FaultResult<()> {
    let limits = working.archive_limits();
    let archive = working.archive_mut(&component)?;
    let object = archive.object_mut(identifier).ok_or(Fault::Invalid)?;
    let message_type = object
        .messages
        .get(message_index)
        .ok_or(Fault::Invalid)?
        .type_;
    replace_message_at_inner(object, message_index, message_type, data, limits)
}

fn replace_message_at_inner(
    object: &mut ArchiveObject,
    index: usize,
    message_type: u32,
    data: Vec<u8>,
    limits: litchi_iwa_core::Limits,
) -> FaultResult<()> {
    object
        .replace_message_preserving_header(
            index,
            RawMessage {
                type_: message_type,
                data,
            },
            limits,
        )
        .map_err(map_core_error)
        .map(|_| ())
}

fn rewrite_tile<P: TableLeaf>(
    working: &mut Working<'_, P>,
    selection: &Selection,
    expected: Option<u32>,
    replacement: Option<u32>,
) -> FaultResult<()> {
    let component = selection.tile_component.to_string();
    let source = working
        .object(selection.tile_identifier)?
        .messages
        .get(selection.tile_message_index)
        .ok_or(Fault::Invalid)?
        .data
        .as_slice();
    let source = structure::copy_bytes(source)?;
    let mut matched = false;
    let mut output =
        wire::transform_repeated_length_delimited_fields(&source, TABLE_TILE_ROWS_FIELD, |row| {
            let view =
                WireView::parse(row, WireLimits::DEFAULT).map_err(WireCallbackError::from)?;
            let mut index_field = None;
            for field in view
                .fields()
                .filter(|field| field.number() == TABLE_TILE_ROW_INDEX_FIELD)
            {
                if index_field.replace(field).is_some() {
                    return Err(WireCallbackError::from(Fault::Invalid));
                }
            }
            let Some(index_field) = index_field else {
                return structure::copy_bytes(row).map_err(WireCallbackError::from);
            };
            if read_varint(index_field.payload()).map_err(WireCallbackError::from)? as u32
                != selection.tile_row
            {
                return structure::copy_bytes(row).map_err(WireCallbackError::from);
            }
            if matched {
                return Err(WireCallbackError::from(Fault::Invalid));
            }
            matched = true;
            rewrite_row(
                row,
                selection.column,
                selection.table_columns,
                expected,
                replacement,
            )
            .map_err(WireCallbackError::from)
        })
        .map_err(|error: WireCallbackError| error.0)?;
    if !matched {
        if expected.is_some() || replacement.is_none() {
            return Err(Fault::Verification);
        }
        let row = new_row(
            selection.tile_row,
            selection.column,
            selection.table_columns,
            replacement.ok_or(Fault::Invalid)?,
        )?;
        output =
            wire::patch_varint_field(&source, 4, true, Some(u64::from(selection.tile_row + 1)))
                .map_err(map_wire_error)?;
        wire::append_length_delimited_field(
            &mut output,
            TABLE_TILE_ROWS_FIELD,
            &row,
            WireLimits::DEFAULT,
        )
        .map_err(map_wire_error)?;
    }
    replace_message(
        working,
        selection.tile_identifier,
        selection.tile_message_type,
        output,
        component,
    )
}

fn new_row(row: u32, column: u32, columns: u32, replacement: u32) -> FaultResult<Vec<u8>> {
    if column >= columns {
        return Err(Fault::Invalid);
    }
    let columns = usize::try_from(columns).map_err(|_| Fault::Invalid)?;
    let column = usize::try_from(column).map_err(|_| Fault::Invalid)?;
    if columns > WireLimits::MAX_FIELDS {
        return Err(Fault::Limit {
            kind: LimitKind::WireFields,
            observed: columns as u64,
            maximum: WireLimits::MAX_FIELDS as u64,
        });
    }
    let offset_bytes = columns.checked_mul(2).ok_or(Fault::Limit {
        kind: LimitKind::WireBytes,
        observed: u64::MAX,
        maximum: WireLimits::MAX_OUTPUT_BYTES as u64,
    })?;
    let mut offsets = Vec::new();
    structure::reserve(&mut offsets, offset_bytes)?;
    offsets.resize(offset_bytes, 0xff_u8);
    offsets[column * 2..column * 2 + 2].copy_from_slice(&0_u16.to_le_bytes());
    let mut cell = litchi_numbers_wire::BncCell::minimal();
    cell.set_comment_identifier(Some(replacement));
    let cell = cell
        .encode(&litchi_numbers_wire::Limits::DEFAULT)
        .map_err(map_bnc_error)?;
    let mut out = Vec::new();
    let output_capacity = offset_bytes
        .checked_add(cell.len())
        .and_then(|size| size.checked_add(64))
        .ok_or(Fault::Limit {
            kind: LimitKind::WireOutputBytes,
            observed: u64::MAX,
            maximum: WireLimits::MAX_OUTPUT_BYTES as u64,
        })?;
    if output_capacity > WireLimits::MAX_OUTPUT_BYTES {
        return Err(Fault::Limit {
            kind: LimitKind::WireOutputBytes,
            observed: output_capacity as u64,
            maximum: WireLimits::MAX_OUTPUT_BYTES as u64,
        });
    }
    structure::reserve(&mut out, output_capacity)?;
    wire::append_varint_field(&mut out, TABLE_TILE_ROW_INDEX_FIELD, u64::from(row))
        .map_err(map_wire_error)?;
    wire::append_varint_field(&mut out, TABLE_TILE_ROW_CELL_COUNT_FIELD, 1)
        .map_err(map_wire_error)?;
    wire::append_length_delimited_field(&mut out, 3, &[], WireLimits::DEFAULT)
        .map_err(map_wire_error)?;
    wire::append_length_delimited_field(&mut out, 4, &[], WireLimits::DEFAULT)
        .map_err(map_wire_error)?;
    wire::append_varint_field(&mut out, 5, 5).map_err(map_wire_error)?;
    wire::append_length_delimited_field(
        &mut out,
        TABLE_TILE_ROW_CELL_STORAGE_FIELD,
        &cell,
        WireLimits::DEFAULT,
    )
    .map_err(map_wire_error)?;
    wire::append_length_delimited_field(
        &mut out,
        TABLE_TILE_ROW_CELL_OFFSETS_FIELD,
        &offsets,
        WireLimits::DEFAULT,
    )
    .map_err(map_wire_error)?;
    wire::append_varint_field(&mut out, TABLE_TILE_ROW_WIDE_OFFSETS_FIELD, 0)
        .map_err(map_wire_error)?;
    Ok(out)
}

fn rewrite_row(
    source: &[u8],
    column: u32,
    columns: u32,
    expected: Option<u32>,
    replacement: Option<u32>,
) -> FaultResult<Vec<u8>> {
    let view = WireView::parse(source, WireLimits::DEFAULT).map_err(map_wire_error)?;
    let storage = unique_field(&view, TABLE_TILE_ROW_CELL_STORAGE_FIELD)?.ok_or(Fault::Invalid)?;
    let offsets = unique_field(&view, TABLE_TILE_ROW_CELL_OFFSETS_FIELD)?.ok_or(Fault::Invalid)?;
    if offsets.wire_type() != 2 || storage.wire_type() != 2 {
        return Err(Fault::Invalid);
    }
    let wide = unique_field(&view, TABLE_TILE_ROW_WIDE_OFFSETS_FIELD)?
        .map(|field| read_varint(field.payload()).map(|value| value != 0))
        .transpose()?;
    let unit = if wide.unwrap_or(false) { 4 } else { 1 };
    let raw_offsets = offsets.payload();
    if !raw_offsets.len().is_multiple_of(2) {
        return Err(Fault::Invalid);
    }
    let column = usize::try_from(column).map_err(|_| Fault::Invalid)?;
    let columns = usize::try_from(columns).map_err(|_| Fault::Invalid)?;
    if columns > WireLimits::MAX_FIELDS || column >= columns {
        if columns > WireLimits::MAX_FIELDS {
            return Err(Fault::Limit {
                kind: LimitKind::WireFields,
                observed: columns as u64,
                maximum: WireLimits::MAX_FIELDS as u64,
            });
        }
        return Err(Fault::Invalid);
    }
    let bytes = storage.payload();
    let offset_count = raw_offsets.len() / 2;
    if offset_count > WireLimits::MAX_FIELDS {
        return Err(Fault::Limit {
            kind: LimitKind::WireFields,
            observed: offset_count as u64,
            maximum: WireLimits::MAX_FIELDS as u64,
        });
    }
    let mut starts = Vec::new();
    structure::reserve(&mut starts, offset_count)?;
    for chunk in raw_offsets.chunks_exact(2) {
        let value = u16::from_le_bytes([chunk[0], chunk[1]]);
        if value == u16::MAX {
            starts.push(None);
        } else {
            let start = usize::from(value).checked_mul(unit).ok_or(Fault::Invalid)?;
            if start > bytes.len() {
                return Err(Fault::Invalid);
            }
            starts.push(Some(start));
        }
    }
    if starts
        .iter()
        .enumerate()
        .any(|(index, start)| index >= columns && start.is_some())
    {
        return Err(Fault::Invalid);
    }
    // A tile row's offset buffer is capacity-sized rather than table-sized.
    // Keep its empty tail intact: truncating the 255-slot native buffer to
    // the selected table width drops unrelated cell topology and makes a
    // comment-only rewrite rewrite the whole row.
    if starts.len() < columns {
        let missing = columns.saturating_sub(starts.len());
        structure::reserve(&mut starts, missing)?;
        starts.resize(columns, None);
    }
    let declared = unique_field(&view, TABLE_TILE_ROW_CELL_COUNT_FIELD)?
        .ok_or(Fault::Invalid)
        .and_then(|field| read_varint(field.payload()))?;
    if declared as usize != starts.iter().filter(|start| start.is_some()).count() {
        return Err(Fault::Invalid);
    }
    let mut cells = Vec::new();
    structure::reserve(&mut cells, starts.len())?;
    for (index, start) in starts.iter().enumerate() {
        let Some(start) = start else {
            cells.push(None);
            continue;
        };
        let end = starts
            .iter()
            .skip(index + 1)
            .flatten()
            .next()
            .copied()
            .unwrap_or(bytes.len());
        if *start >= end {
            return Err(Fault::Invalid);
        }
        cells.push(Some(structure::copy_bytes(&bytes[*start..end])?));
    }
    let previous = match cells.get(column).and_then(Option::as_deref) {
        Some(data) => {
            litchi_numbers_wire::BncCell::parse(data, &litchi_numbers_wire::Limits::DEFAULT)
                .map_err(map_bnc_error)?
                .comment_identifier()
        },
        None => None,
    };
    if previous != expected {
        return Err(Fault::Verification);
    }
    match (cells.get(column).and_then(Option::as_deref), replacement) {
        (None, Some(id)) => {
            let mut cell = litchi_numbers_wire::BncCell::minimal();
            cell.set_comment_identifier(Some(id));
            cells[column] = Some(
                cell.encode(&litchi_numbers_wire::Limits::DEFAULT)
                    .map_err(map_bnc_error)?,
            );
        },
        (Some(_), replacement) => {
            let mut cell = litchi_numbers_wire::BncCell::parse(
                cells[column].as_ref().ok_or(Fault::Invalid)?,
                &litchi_numbers_wire::Limits::DEFAULT,
            )
            .map_err(map_bnc_error)?;
            cell.set_comment_identifier(replacement);
            cells[column] = Some(
                cell.encode(&litchi_numbers_wire::Limits::DEFAULT)
                    .map_err(map_bnc_error)?,
            );
        },
        (None, None) => return Err(Fault::Invalid),
    }
    let count = cells.iter().filter(|cell| cell.is_some()).count();
    let mut encoded = Vec::new();
    structure::reserve(&mut encoded, bytes.len())?;
    let mut encoded_offsets = Vec::new();
    structure::reserve(&mut encoded_offsets, raw_offsets.len())?;
    for cell in cells {
        let Some(cell) = cell else {
            encoded_offsets.extend_from_slice(&u16::MAX.to_le_bytes());
            continue;
        };
        let offset = encoded.len().checked_div(unit).ok_or(Fault::Invalid)?;
        if offset >= usize::from(u16::MAX) || encoded.len() % unit != 0 {
            return Err(Fault::Limit {
                kind: LimitKind::PayloadBytes,
                observed: offset as u64,
                maximum: u16::MAX as u64,
            });
        }
        encoded_offsets.extend_from_slice(
            &u16::try_from(offset)
                .map_err(|_| Fault::Invalid)?
                .to_le_bytes(),
        );
        encoded.extend_from_slice(&cell);
        let remainder = encoded.len() % unit;
        if remainder != 0 {
            encoded.resize(encoded.len() + unit - remainder, 0);
        }
    }
    let mut output = wire::patch_length_delimited_field(
        source,
        TABLE_TILE_ROW_CELL_STORAGE_FIELD,
        true,
        Some(&encoded),
    )
    .map_err(map_wire_error)?;
    output = wire::patch_length_delimited_field(
        &output,
        TABLE_TILE_ROW_CELL_OFFSETS_FIELD,
        true,
        Some(&encoded_offsets),
    )
    .map_err(map_wire_error)?;
    wire::patch_varint_field(
        &output,
        TABLE_TILE_ROW_CELL_COUNT_FIELD,
        true,
        Some(count as u64),
    )
    .map_err(map_wire_error)
}

fn varint_field(number: u32, value: u64) -> FaultResult<Vec<u8>> {
    let mut bytes = Vec::new();
    wire::append_varint_field(&mut bytes, number, value).map_err(map_wire_error)?;
    Ok(bytes)
}

fn register_insertions<P: TableLeaf>(
    working: &Working<'_, P>,
    registry: &metadata::Registry,
    inserted: &[u64],
    transition: &mut metadata::Transition,
) -> FaultResult<()> {
    for &identifier in inserted {
        let component = working.locate(identifier)?.ok_or(Fault::Invalid)?;
        let component_id = registry.component_for_entry(&component)?.identifier;
        let object = working.object(identifier)?;
        for info in &object.archive_info.message_infos {
            for &target in &info.object_references {
                let target_component = working.locate(target)?.ok_or(Fault::Invalid)?;
                let target_id = registry.component_for_entry(&target_component)?.identifier;
                transition.reference(registry, component_id, target_id, target)?;
            }
            for field in &info.field_infos {
                for &target in &field.object_references {
                    let target_component = working.locate(target)?.ok_or(Fault::Invalid)?;
                    let target_id = registry.component_for_entry(&target_component)?.identifier;
                    transition.reference(registry, component_id, target_id, target)?;
                }
            }
        }
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::{COMMENT_STORAGE_MESSAGE_TYPE, Fault, comment_storage_message_index};
    use litchi_iwa_core::{ArchiveObject, RawMessage};

    #[test]
    fn duplicate_comment_storage_payload_is_rejected() {
        let object = ArchiveObject::new(
            1,
            vec![
                RawMessage {
                    type_: COMMENT_STORAGE_MESSAGE_TYPE,
                    data: vec![1],
                },
                RawMessage {
                    type_: COMMENT_STORAGE_MESSAGE_TYPE,
                    data: vec![2],
                },
            ],
            Default::default(),
        )
        .expect("bounded test object should construct");
        assert_eq!(
            comment_storage_message_index(&object),
            Err(Fault::Invalid),
            "ambiguous comment roots must not select the first payload"
        );
    }
}
