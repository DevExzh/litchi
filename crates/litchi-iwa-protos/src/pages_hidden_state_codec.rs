//! Strict source-preserving projection of the Pages table hidden-state graph.
//!
//! Pages stores user-hidden rows and columns in a small graph rooted at
//! `TST.TableInfoArchive.hidden_states_uuid` and
//! `TST.TableModelArchive.hidden_states_owner`.  The graph is deliberately
//! handled here instead of through generated Prost values: a rewrite changes
//! only selected graph fields, while every unrelated field (including unknown
//! fields and groups) is copied byte-for-byte from the caller-owned source.
//!
//! The private Buffa sidecar contains only singular scalar/envelope fields.
//! Repeated hidden states are parsed by the bounded wire walker below, so an
//! input-sized `RepeatedView`/owned generated collection can never cross this
//! boundary.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "The strict schema-directed pass intentionally precedes the Buffa check."
)]

use core::{fmt, mem::size_of};
use std::{num::NonZeroU64, str};

use buffa::DecodeOptions as BuffaDecodeOptions;

use crate::buffa_pages_hidden_state_generated::LitchiIwaPagesHiddenStateProjection as projection;

const TABLE_INFO_SUPER_FIELD: u32 = 1;
const TABLE_INFO_MODEL_FIELD: u32 = 2;
const TABLE_INFO_VIEW_UIDS_FIELD: u32 = 6;
const TABLE_INFO_HIDDEN_STATES_UUID_FIELD: u32 = 8;

const TABLE_MODEL_TABLE_ID_FIELD: u32 = 1;
const TABLE_MODEL_TABLE_STYLE_FIELD: u32 = 3;
const TABLE_MODEL_BASE_DATA_STORE_FIELD: u32 = 4;
const TABLE_MODEL_ROWS_FIELD: u32 = 6;
const TABLE_MODEL_COLUMNS_FIELD: u32 = 7;
const TABLE_MODEL_FILTERED_ROWS_FIELD: u32 = 40;
const TABLE_MODEL_BASE_COLUMN_ROW_UIDS_FIELD: u32 = 46;
const TABLE_MODEL_TABLE_NAME_FIELD: u32 = 8;
const TABLE_MODEL_DEFAULT_ROW_HEIGHT_FIELD: u32 = 16;
const TABLE_MODEL_DEFAULT_COLUMN_WIDTH_FIELD: u32 = 17;
const TABLE_MODEL_BODY_CELL_STYLE_FIELD: u32 = 18;
const TABLE_MODEL_HEADER_ROW_STYLE_FIELD: u32 = 19;
const TABLE_MODEL_HEADER_COLUMN_STYLE_FIELD: u32 = 20;
const TABLE_MODEL_FOOTER_ROW_STYLE_FIELD: u32 = 21;
const TABLE_MODEL_BODY_TEXT_STYLE_FIELD: u32 = 24;
const TABLE_MODEL_HEADER_ROW_TEXT_STYLE_FIELD: u32 = 25;
const TABLE_MODEL_HEADER_COLUMN_TEXT_STYLE_FIELD: u32 = 26;
const TABLE_MODEL_FOOTER_ROW_TEXT_STYLE_FIELD: u32 = 27;
const TABLE_MODEL_HIDDEN_ROWS_FIELD: u32 = 14;
const TABLE_MODEL_HIDDEN_COLUMNS_FIELD: u32 = 15;
const TABLE_MODEL_COLUMN_FORMULA_OWNER_FIELD: u32 = 34;
const TABLE_MODEL_ROW_FORMULA_OWNER_FIELD: u32 = 35;
const TABLE_MODEL_USER_HIDDEN_ROWS_FIELD: u32 = 41;
const TABLE_MODEL_USER_HIDDEN_COLUMNS_FIELD: u32 = 42;
const TABLE_MODEL_HIDDEN_STATES_OWNER_FIELD: u32 = 70;

const UUID_LOWER_FIELD: u32 = 1;
const UUID_UPPER_FIELD: u32 = 2;
const REFERENCE_IDENTIFIER_FIELD: u32 = 1;
const REFERENCE_DEPRECATED_TYPE_FIELD: u32 = 2;
const REFERENCE_DEPRECATED_EXTERNAL_FIELD: u32 = 3;

const OWNER_UID_FIELD: u32 = 1;
const OWNER_STATES_FIELD: u32 = 2;
const STATE_UID_FIELD: u32 = 1;
const STATE_COLUMN_EXTENT_FIELD: u32 = 2;
const STATE_ROW_EXTENT_FIELD: u32 = 3;
const EXTENT_UID_FIELD: u32 = 1;
const EXTENT_BASE_STATES_FIELD: u32 = 2;
const EXTENT_DIRECTION_FIELD: u32 = 3;
const EXTENT_NEEDS_FILTER_UPDATE_FIELD: u32 = 6;
const EXTENT_FILTER_SET_FIELD: u32 = 8;
const ROW_STATE_UID_FIELD: u32 = 1;
const ROW_STATE_USER_HIDDEN_FIELD: u32 = 2;
const ROW_STATE_FILTERED_FIELD: u32 = 3;
const ROW_STATE_PIVOT_HIDDEN_FIELD: u32 = 4;

const FORMULA_OWNER_UID_FIELD: u32 = 1;
const FORMULA_OWNER_INTERNAL_ID_FIELD: u32 = 2;
const FORMULA_OWNER_KIND_FIELD: u32 = 3;
const FORMULA_OWNER_REFERENCE_FIELD: u32 = 11;
const FORMULA_OWNER_BASE_UID_FIELD: u32 = 12;
const HIDDEN_FORMULA_OWNER_ID_FIELD: u32 = 1;
const HIDDEN_FORMULA_OWNER_NEEDS_UPDATE_FIELD: u32 = 3;
const FILTER_SET_TYPE_FIELD: u32 = 1;
const FILTER_SET_ENABLED_FIELD: u32 = 2;
const FILTER_SET_NEEDS_FORMULA_REWRITE_FIELD: u32 = 4;
const FILTER_SET_OFFSETS_FIELD: u32 = 5;

/// Native archive message IDs used by the Pages hidden-state object graph.
/// The IDs are routing metadata and are never inferred from payload bytes.
pub const COLUMN_ROW_UID_MAP_MESSAGE_TYPE: u32 = 6_267;
/// Legacy Numbers UID-map routing ID; callers must opt in explicitly.
pub const LEGACY_COLUMN_ROW_UID_MAP_MESSAGE_TYPE: u32 = 6_200;
pub const FORMULA_OWNER_DEPENDENCIES_MESSAGE_TYPE: u32 = 4_008;
pub const HIDDEN_STATE_FORMULA_OWNER_MESSAGE_TYPE: u32 = 6_204;
pub const FILTER_SET_MESSAGE_TYPE: u32 = 6_220;

const MAX_RECURSION: u32 = 64;
const MAX_CONSTRUCTED_STATES: usize = 16_384;
const BUFFA_MAX_MESSAGE_BYTES: usize = buffa::MAX_MESSAGE_BYTES as usize;
const MAX_CFUUID_BYTES: usize = 16;

const fn clamp_buffa_message_bytes(value: usize) -> usize {
    if value > BUFFA_MAX_MESSAGE_BYTES {
        BUFFA_MAX_MESSAGE_BYTES
    } else {
        value
    }
}

/// Finite policy for strict hidden-state reads and source-preserving rewrites.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeOptions {
    max_input_bytes: usize,
    max_output_bytes: usize,
    max_fields: usize,
    max_work_bytes: usize,
    recursion_limit: u32,
    max_states: usize,
    max_allocations: usize,
    max_retained_bytes: usize,
    max_scratch_bytes: usize,
}

impl DecodeOptions {
    /// Construct a finite input/output, field/work/nesting/state policy.
    #[must_use]
    pub const fn new(
        max_input_bytes: usize,
        max_output_bytes: usize,
        max_fields: usize,
        max_work_bytes: usize,
        recursion_limit: u32,
        max_states: usize,
    ) -> Self {
        Self {
            max_input_bytes: clamp_buffa_message_bytes(max_input_bytes),
            max_output_bytes: clamp_buffa_message_bytes(max_output_bytes),
            max_fields,
            max_work_bytes,
            recursion_limit,
            max_states,
            max_allocations: 1024,
            max_retained_bytes: max_output_bytes,
            max_scratch_bytes: {
                let scratch = max_input_bytes.saturating_mul(4);
                if scratch == 0 { 1 } else { scratch }
            },
        }
    }

    /// Build a conservative source-sized policy for one trusted payload.
    #[must_use]
    pub fn for_source(source: &[u8]) -> Self {
        let bytes = source.len().max(1);
        Self::new(
            bytes,
            bytes.saturating_mul(2).max(1),
            bytes.saturating_mul(8).max(1),
            bytes.saturating_mul(64).max(1),
            16,
            bytes.clamp(1, MAX_CONSTRUCTED_STATES),
        )
    }

    #[must_use]
    pub const fn with_max_input_bytes(mut self, value: usize) -> Self {
        self.max_input_bytes = clamp_buffa_message_bytes(value);
        self
    }

    #[must_use]
    pub const fn with_max_output_bytes(mut self, value: usize) -> Self {
        self.max_output_bytes = clamp_buffa_message_bytes(value);
        self
    }

    #[must_use]
    pub const fn with_max_fields(mut self, value: usize) -> Self {
        self.max_fields = value;
        self
    }

    #[must_use]
    pub const fn with_max_work_bytes(mut self, value: usize) -> Self {
        self.max_work_bytes = value;
        self
    }

    #[must_use]
    pub const fn with_recursion_limit(mut self, value: u32) -> Self {
        self.recursion_limit = value;
        self
    }

    #[must_use]
    pub const fn with_max_states(mut self, value: usize) -> Self {
        self.max_states = value;
        self
    }

    #[must_use]
    pub const fn with_max_allocations(mut self, value: usize) -> Self {
        self.max_allocations = value;
        self
    }

    #[must_use]
    pub const fn with_max_retained_bytes(mut self, value: usize) -> Self {
        self.max_retained_bytes = value;
        self
    }

    #[must_use]
    pub const fn with_max_scratch_bytes(mut self, value: usize) -> Self {
        self.max_scratch_bytes = value;
        self
    }

    #[must_use]
    pub const fn max_input_bytes(self) -> usize {
        self.max_input_bytes
    }

    #[must_use]
    pub const fn max_output_bytes(self) -> usize {
        self.max_output_bytes
    }

    #[must_use]
    pub const fn max_fields(self) -> usize {
        self.max_fields
    }

    #[must_use]
    pub const fn max_work_bytes(self) -> usize {
        self.max_work_bytes
    }

    #[must_use]
    pub const fn recursion_limit(self) -> u32 {
        self.recursion_limit
    }

    #[must_use]
    pub const fn max_states(self) -> usize {
        self.max_states
    }

    #[must_use]
    pub const fn max_allocations(self) -> usize {
        self.max_allocations
    }

    #[must_use]
    pub const fn max_retained_bytes(self) -> usize {
        self.max_retained_bytes
    }

    #[must_use]
    pub const fn max_scratch_bytes(self) -> usize {
        self.max_scratch_bytes
    }

    fn buffa(self) -> BuffaDecodeOptions {
        BuffaDecodeOptions::new()
            .with_max_message_size(clamp_buffa_message_bytes(self.max_input_bytes))
            .with_unknown_field_limit(self.max_fields)
            .with_element_memory_limit(0)
            .with_recursion_limit(self.recursion_limit)
    }
}

/// Finite resource observations returned by strict reads and rewrites.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum DecodeLimit {
    InputBytes { observed: usize, maximum: usize },
    OutputBytes { observed: usize, maximum: usize },
    Fields { observed: usize, maximum: usize },
    WorkBytes { observed: usize, maximum: usize },
    Nesting { observed: u32, maximum: u32 },
    States { observed: usize, maximum: usize },
    Allocations { observed: usize, maximum: usize },
    RetainedBytes { observed: usize, maximum: usize },
    ScratchBytes { observed: usize, maximum: usize },
}

/// Strict hidden-state graph failure. Diagnostics never include source bytes,
/// authored table names, or native object identifiers.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DecodeError {
    kind: DecodeErrorKind,
}

#[derive(Debug, Clone, PartialEq, Eq)]
enum DecodeErrorKind {
    Wire(buffa::DecodeError),
    Invalid(&'static str),
    MissingRequired(&'static str),
    DuplicateSingular(&'static str),
    NonCanonical(&'static str),
    Limit(DecodeLimit),
    Allocation { requested: usize },
    Projection,
}

impl DecodeError {
    const fn wire(error: buffa::DecodeError) -> Self {
        Self {
            kind: DecodeErrorKind::Wire(error),
        }
    }

    const fn invalid(reason: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::Invalid(reason),
        }
    }

    const fn missing(field: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::MissingRequired(field),
        }
    }

    const fn duplicate(field: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::DuplicateSingular(field),
        }
    }

    const fn noncanonical(reason: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::NonCanonical(reason),
        }
    }

    const fn limit(limit: DecodeLimit) -> Self {
        Self {
            kind: DecodeErrorKind::Limit(limit),
        }
    }

    const fn allocation(requested: usize) -> Self {
        Self {
            kind: DecodeErrorKind::Allocation { requested },
        }
    }

    const fn projection() -> Self {
        Self {
            kind: DecodeErrorKind::Projection,
        }
    }

    /// Return the finite resource failure, when applicable.
    #[must_use]
    pub const fn resource_limit(&self) -> Option<DecodeLimit> {
        match self.kind {
            DecodeErrorKind::Limit(limit) => Some(limit),
            _ => None,
        }
    }

    /// Return a required schema field that was absent, when applicable.
    #[must_use]
    pub const fn missing_required_field(&self) -> Option<&'static str> {
        match self.kind {
            DecodeErrorKind::MissingRequired(field) => Some(field),
            _ => None,
        }
    }

    /// Return a duplicated singular schema field, when applicable.
    #[must_use]
    pub const fn duplicate_singular_field(&self) -> Option<&'static str> {
        match self.kind {
            DecodeErrorKind::DuplicateSingular(field) => Some(field),
            _ => None,
        }
    }

    /// Return a canonical-wire failure reason, when applicable.
    #[must_use]
    pub const fn noncanonical_reason(&self) -> Option<&'static str> {
        match self.kind {
            DecodeErrorKind::NonCanonical(reason) => Some(reason),
            _ => None,
        }
    }

    /// Return the requested allocation amount, when applicable.
    #[must_use]
    pub const fn allocation_amount(&self) -> Option<usize> {
        match self.kind {
            DecodeErrorKind::Allocation { requested } => Some(requested),
            _ => None,
        }
    }
}

impl fmt::Display for DecodeError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match &self.kind {
            DecodeErrorKind::Wire(error) => error.fmt(formatter),
            DecodeErrorKind::Invalid(reason) => {
                write!(formatter, "invalid Pages hidden-state graph {reason}")
            },
            DecodeErrorKind::MissingRequired(field) => {
                write!(formatter, "missing required field {field}")
            },
            DecodeErrorKind::DuplicateSingular(field) => {
                write!(formatter, "duplicate singular field {field}")
            },
            DecodeErrorKind::NonCanonical(reason) => {
                write!(formatter, "non-canonical protobuf representation: {reason}")
            },
            DecodeErrorKind::Limit(_) => {
                formatter.write_str("Pages hidden-state graph resource limit exceeded")
            },
            DecodeErrorKind::Allocation { requested } => {
                write!(
                    formatter,
                    "cannot allocate Pages hidden-state graph storage for {requested} bytes"
                )
            },
            DecodeErrorKind::Projection => formatter.write_str(
                "Pages hidden-state graph strict preflight disagrees with the Buffa projection",
            ),
        }
    }
}

impl std::error::Error for DecodeError {}

impl From<buffa::DecodeError> for DecodeError {
    fn from(error: buffa::DecodeError) -> Self {
        Self::wire(error)
    }
}

/// Exact native UUID value used by the hidden-state graph.
#[derive(Clone, Copy, PartialEq, Eq, Hash)]
pub struct UuidSnapshot {
    lower: u64,
    upper: u64,
}

impl fmt::Debug for UuidSnapshot {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.debug_struct("UuidSnapshot").finish()
    }
}

impl UuidSnapshot {
    #[must_use]
    pub const fn new(lower: u64, upper: u64) -> Self {
        Self { lower, upper }
    }

    #[must_use]
    pub const fn lower(self) -> u64 {
        self.lower
    }

    #[must_use]
    pub const fn upper(self) -> u64 {
        self.upper
    }
}

/// Compatibility spelling for callers that use the native type name.
pub type Uuid = UuidSnapshot;

/// Strict native object reference value.
#[derive(Clone, Copy, PartialEq, Eq, Hash)]
pub struct ReferenceSnapshot {
    identifier: NonZeroU64,
    deprecated_type: Option<i32>,
    deprecated_is_external: Option<bool>,
}

impl fmt::Debug for ReferenceSnapshot {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("ReferenceSnapshot")
            .field("has_deprecated_type", &self.deprecated_type.is_some())
            .field(
                "has_deprecated_is_external",
                &self.deprecated_is_external.is_some(),
            )
            .finish()
    }
}

impl ReferenceSnapshot {
    #[must_use]
    pub const fn new(identifier: NonZeroU64) -> Self {
        Self {
            identifier,
            deprecated_type: None,
            deprecated_is_external: None,
        }
    }

    #[must_use]
    pub const fn with_deprecated(mut self, value: Option<i32>) -> Self {
        self.deprecated_type = value;
        self
    }

    #[must_use]
    pub const fn with_external(mut self, value: Option<bool>) -> Self {
        self.deprecated_is_external = value;
        self
    }

    #[must_use]
    pub const fn identifier(self) -> NonZeroU64 {
        self.identifier
    }

    #[must_use]
    pub const fn deprecated_type(self) -> Option<i32> {
        self.deprecated_type
    }

    #[must_use]
    pub const fn deprecated_is_external(self) -> Option<bool> {
        self.deprecated_is_external
    }
}

/// Presence-preserving `TSP.CFUUIDArchive` identity used by the 6204
/// hidden-state formula-owner record.  The byte form is retained as owned
/// storage so the projection never borrows a generated value.
#[derive(Clone, PartialEq, Eq)]
pub struct CfuuidSnapshot {
    uuid_bytes: Option<Vec<u8>>,
    words: [Option<u32>; 4],
}

impl fmt::Debug for CfuuidSnapshot {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("CfuuidSnapshot")
            .field("has_uuid_bytes", &self.uuid_bytes.is_some())
            .field("words_present", &self.words.map(|word| word.is_some()))
            .finish()
    }
}

impl CfuuidSnapshot {
    #[must_use]
    pub fn new(uuid_bytes: Option<Vec<u8>>, words: [Option<u32>; 4]) -> Self {
        Self { uuid_bytes, words }
    }

    #[must_use]
    pub fn uuid_bytes(&self) -> Option<&[u8]> {
        self.uuid_bytes.as_deref()
    }

    #[must_use]
    pub const fn words(&self) -> [Option<u32>; 4] {
        self.words
    }

    /// Convert the native four-word representation when it is complete.
    #[must_use]
    pub const fn words_uuid(&self) -> Option<UuidSnapshot> {
        match self.words {
            [Some(w0), Some(w1), Some(w2), Some(w3)] => Some(UuidSnapshot::new(
                w0 as u64 | ((w1 as u64) << 32),
                w2 as u64 | ((w3 as u64) << 32),
            )),
            _ => None,
        }
    }
}

/// Strict identity projection of one formula-dependency object (type 4008).
/// The existing dependency codec validates the complete dependency closure;
/// this value exposes only the selected ownership edges to Pages adapters.
#[derive(Clone, PartialEq, Eq)]
pub struct FormulaOwnerDependenciesSnapshot {
    formula_owner_uid: UuidSnapshot,
    internal_formula_owner_id: u32,
    owner_kind: Option<u32>,
    formula_owner: Option<ReferenceSnapshot>,
    base_owner_uid: Option<UuidSnapshot>,
    // Keep physical presence separate from the selected identity projection.
    // An empty dependency envelope still changes the mutation contract.
    has_dependencies: bool,
}

impl fmt::Debug for FormulaOwnerDependenciesSnapshot {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("FormulaOwnerDependenciesSnapshot")
            .field("has_owner_kind", &self.owner_kind.is_some())
            .field("has_formula_owner", &self.formula_owner.is_some())
            .field("has_base_owner_uid", &self.base_owner_uid.is_some())
            .field("has_dependencies", &self.has_dependencies)
            .finish()
    }
}

impl FormulaOwnerDependenciesSnapshot {
    #[must_use]
    pub const fn formula_owner_uid(&self) -> UuidSnapshot {
        self.formula_owner_uid
    }

    #[must_use]
    pub const fn internal_formula_owner_id(&self) -> u32 {
        self.internal_formula_owner_id
    }

    #[must_use]
    pub const fn owner_kind(&self) -> Option<u32> {
        self.owner_kind
    }

    #[must_use]
    pub const fn formula_owner(&self) -> Option<ReferenceSnapshot> {
        self.formula_owner
    }

    #[must_use]
    pub const fn base_owner_uid(&self) -> Option<UuidSnapshot> {
        self.base_owner_uid
    }

    /// Whether any physical dependency envelope is present in the source.
    ///
    /// This reports wire presence for fields 4..=10 and 13..=16, including
    /// zero-length envelopes.  It intentionally does not infer presence from
    /// decoded children because an empty envelope is still meaningful to the
    /// existing-owner mutation contract.
    #[must_use]
    pub const fn has_dependencies(&self) -> bool {
        self.has_dependencies
    }
}

/// Selected scalar/collection projection of one type-6220 filter-set object.
/// Rule records remain source-owned and are only structurally scanned.
#[derive(Clone, PartialEq, Eq)]
pub struct FilterSetSnapshot {
    filter_type: Option<i32>,
    is_enabled: Option<bool>,
    needs_formula_rewrite_for_import: Option<bool>,
    filter_offsets: Vec<u32>,
}

impl fmt::Debug for FilterSetSnapshot {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("FilterSetSnapshot")
            .field("has_type", &self.filter_type.is_some())
            .field("has_is_enabled", &self.is_enabled.is_some())
            .field(
                "has_needs_formula_rewrite_for_import",
                &self.needs_formula_rewrite_for_import.is_some(),
            )
            .field("filter_offsets_len", &self.filter_offsets.len())
            .finish()
    }
}

impl FilterSetSnapshot {
    #[must_use = "the constructor may reject an oversized offset collection"]
    pub fn new(
        filter_type: Option<i32>,
        is_enabled: Option<bool>,
        needs_formula_rewrite_for_import: Option<bool>,
        filter_offsets: impl IntoIterator<Item = u32>,
    ) -> Result<Self, DecodeError> {
        let filter_offsets = collect_u32(filter_offsets, MAX_CONSTRUCTED_STATES)?;
        Ok(Self {
            filter_type,
            is_enabled,
            needs_formula_rewrite_for_import,
            filter_offsets,
        })
    }

    #[must_use]
    pub const fn filter_type(&self) -> Option<i32> {
        self.filter_type
    }

    #[must_use]
    pub const fn is_enabled(&self) -> Option<bool> {
        self.is_enabled
    }

    #[must_use]
    pub const fn needs_formula_rewrite_for_import(&self) -> Option<bool> {
        self.needs_formula_rewrite_for_import
    }

    #[must_use]
    pub fn filter_offsets(&self) -> &[u32] {
        &self.filter_offsets
    }
}

/// Selected identity/flag projection of one type-6204 formula-owner object.
#[derive(Clone, PartialEq, Eq)]
pub struct HiddenStateFormulaOwnerSnapshot {
    owner_id: Option<CfuuidSnapshot>,
    needs_to_update_filter_set_for_import: Option<bool>,
}

impl fmt::Debug for HiddenStateFormulaOwnerSnapshot {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("HiddenStateFormulaOwnerSnapshot")
            .field("has_owner_id", &self.owner_id.is_some())
            .field(
                "has_needs_to_update_filter_set_for_import",
                &self.needs_to_update_filter_set_for_import.is_some(),
            )
            .finish()
    }
}

impl HiddenStateFormulaOwnerSnapshot {
    #[must_use]
    pub const fn owner_id(&self) -> Option<&CfuuidSnapshot> {
        self.owner_id.as_ref()
    }

    #[must_use]
    pub const fn needs_to_update_filter_set_for_import(&self) -> Option<bool> {
        self.needs_to_update_filter_set_for_import
    }
}

/// Strict object-type qualification for package routers.  Payload bytes do
/// not carry this information, so callers must pass the archive message type
/// they observed and choose legacy 6200 support explicitly.
pub fn validate_column_row_uid_map_message_type(
    message_type: u32,
    allow_legacy: bool,
) -> Result<(), DecodeError> {
    if message_type == COLUMN_ROW_UID_MAP_MESSAGE_TYPE
        || (allow_legacy && message_type == LEGACY_COLUMN_ROW_UID_MAP_MESSAGE_TYPE)
    {
        Ok(())
    } else {
        Err(DecodeError::invalid(
            "unsupported column/row UID-map message type",
        ))
    }
}

/// Native extent direction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum AxisDirection {
    Column,
    Row,
}

impl AxisDirection {
    #[must_use]
    pub const fn native_value(self) -> i32 {
        match self {
            Self::Column => 0,
            Self::Row => 1,
        }
    }

    fn from_native(value: i32) -> Result<Self, DecodeError> {
        match value {
            0 => Ok(Self::Column),
            1 => Ok(Self::Row),
            _ => Err(DecodeError::invalid("hidden-state extent direction")),
        }
    }
}

/// Compatibility spelling matching the generated enum's semantic name.
pub type RowOrColumnDirection = AxisDirection;

/// One row/column hidden-state record.
#[derive(Clone, Copy, PartialEq, Eq, Hash)]
pub struct RowOrColumnStateSnapshot {
    row_or_column_uid: UuidSnapshot,
    user_hidden: Option<bool>,
    filtered: Option<bool>,
    pivot_hidden: Option<bool>,
}

impl fmt::Debug for RowOrColumnStateSnapshot {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("RowOrColumnStateSnapshot")
            .field("has_user_hidden", &self.user_hidden.is_some())
            .field("has_filtered", &self.filtered.is_some())
            .field("has_pivot_hidden", &self.pivot_hidden.is_some())
            .finish()
    }
}

impl RowOrColumnStateSnapshot {
    #[must_use]
    pub const fn new(row_or_column_uid: UuidSnapshot) -> Self {
        Self {
            row_or_column_uid,
            user_hidden: None,
            filtered: None,
            pivot_hidden: None,
        }
    }

    #[must_use]
    pub const fn with_user_hidden(mut self, value: Option<bool>) -> Self {
        self.user_hidden = value;
        self
    }

    #[must_use]
    pub const fn with_filtered(mut self, value: Option<bool>) -> Self {
        self.filtered = value;
        self
    }

    #[must_use]
    pub const fn with_pivot_hidden(mut self, value: Option<bool>) -> Self {
        self.pivot_hidden = value;
        self
    }

    #[must_use]
    pub const fn row_or_column_uid(self) -> UuidSnapshot {
        self.row_or_column_uid
    }

    #[must_use]
    pub const fn user_hidden(self) -> Option<bool> {
        self.user_hidden
    }

    #[must_use]
    pub const fn filtered(self) -> Option<bool> {
        self.filtered
    }

    #[must_use]
    pub const fn pivot_hidden(self) -> Option<bool> {
        self.pivot_hidden
    }
}

/// Borrow-free projection of one `HiddenStateExtentArchive`.
#[derive(Clone, PartialEq, Eq)]
pub struct HiddenStateExtentSnapshot {
    hidden_state_extent_uid: UuidSnapshot,
    direction: AxisDirection,
    base_hidden_states: Vec<RowOrColumnStateSnapshot>,
    needs_to_update_filter_set_for_import: Option<bool>,
    filter_set: Option<ReferenceSnapshot>,
}

impl fmt::Debug for HiddenStateExtentSnapshot {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("HiddenStateExtentSnapshot")
            .field("direction", &self.direction)
            .field("base_hidden_states_len", &self.base_hidden_states.len())
            .field(
                "has_needs_to_update_filter_set_for_import",
                &self.needs_to_update_filter_set_for_import.is_some(),
            )
            .field("has_filter_set", &self.filter_set.is_some())
            .finish()
    }
}

impl HiddenStateExtentSnapshot {
    #[must_use = "the constructor may reject an oversized state collection"]
    pub fn new(
        hidden_state_extent_uid: UuidSnapshot,
        direction: AxisDirection,
        base_hidden_states: impl IntoIterator<Item = RowOrColumnStateSnapshot>,
    ) -> Result<Self, DecodeError> {
        let base_hidden_states = collect_states(base_hidden_states, MAX_CONSTRUCTED_STATES)?;
        Ok(Self {
            hidden_state_extent_uid,
            direction,
            base_hidden_states,
            needs_to_update_filter_set_for_import: None,
            filter_set: None,
        })
    }

    #[must_use]
    pub const fn with_needs_to_update_filter_set_for_import(mut self, value: Option<bool>) -> Self {
        self.needs_to_update_filter_set_for_import = value;
        self
    }

    #[must_use]
    pub fn with_filter_set(mut self, value: Option<ReferenceSnapshot>) -> Self {
        self.filter_set = value;
        self
    }

    #[must_use]
    pub const fn hidden_state_extent_uid(&self) -> UuidSnapshot {
        self.hidden_state_extent_uid
    }

    #[must_use]
    pub const fn direction(&self) -> AxisDirection {
        self.direction
    }

    #[must_use]
    pub fn base_hidden_states(&self) -> &[RowOrColumnStateSnapshot] {
        &self.base_hidden_states
    }

    #[must_use]
    pub fn states(&self) -> &[RowOrColumnStateSnapshot] {
        self.base_hidden_states()
    }

    #[must_use]
    pub const fn needs_to_update_filter_set_for_import(&self) -> Option<bool> {
        self.needs_to_update_filter_set_for_import
    }

    #[must_use]
    pub const fn filter_set(&self) -> Option<ReferenceSnapshot> {
        self.filter_set
    }
}

/// Borrow-free projection of one `HiddenStatesArchive`.
#[derive(Clone, PartialEq, Eq)]
pub struct HiddenStatesSnapshot {
    hidden_states_uid: UuidSnapshot,
    column_hidden_state_extent: HiddenStateExtentSnapshot,
    row_hidden_state_extent: HiddenStateExtentSnapshot,
}

impl fmt::Debug for HiddenStatesSnapshot {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("HiddenStatesSnapshot")
            .field(
                "column_hidden_state_count",
                &self.column_hidden_state_extent.base_hidden_states.len(),
            )
            .field(
                "row_hidden_state_count",
                &self.row_hidden_state_extent.base_hidden_states.len(),
            )
            .finish()
    }
}

impl HiddenStatesSnapshot {
    #[must_use = "the constructor may reject an oversized hidden-state graph"]
    pub fn new(
        hidden_states_uid: UuidSnapshot,
        column_hidden_state_extent: HiddenStateExtentSnapshot,
        row_hidden_state_extent: HiddenStateExtentSnapshot,
    ) -> Self {
        Self {
            hidden_states_uid,
            column_hidden_state_extent,
            row_hidden_state_extent,
        }
    }

    #[must_use]
    pub const fn hidden_states_uid(&self) -> UuidSnapshot {
        self.hidden_states_uid
    }

    #[must_use]
    pub const fn column_hidden_state_extent(&self) -> &HiddenStateExtentSnapshot {
        &self.column_hidden_state_extent
    }

    #[must_use]
    pub const fn row_hidden_state_extent(&self) -> &HiddenStateExtentSnapshot {
        &self.row_hidden_state_extent
    }
}

/// Borrow-free projection of one `HiddenStatesOwnerArchive`.
#[derive(Clone, PartialEq, Eq)]
pub struct HiddenStatesOwnerSnapshot {
    owner_uid: UuidSnapshot,
    hidden_states: Vec<HiddenStatesSnapshot>,
}

impl fmt::Debug for HiddenStatesOwnerSnapshot {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("HiddenStatesOwnerSnapshot")
            .field("hidden_states_len", &self.hidden_states.len())
            .finish()
    }
}

impl HiddenStatesOwnerSnapshot {
    #[must_use = "the constructor may reject an oversized hidden-state graph"]
    pub fn new(
        owner_uid: UuidSnapshot,
        hidden_states: impl IntoIterator<Item = HiddenStatesSnapshot>,
    ) -> Result<Self, DecodeError> {
        let hidden_states = collect_hidden_states(hidden_states, MAX_CONSTRUCTED_STATES)?;
        let mut records = hidden_states.len();
        for hidden_state in &hidden_states {
            records = records
                .checked_add(
                    hidden_state
                        .column_hidden_state_extent
                        .base_hidden_states
                        .len(),
                )
                .and_then(|value| {
                    value.checked_add(
                        hidden_state
                            .row_hidden_state_extent
                            .base_hidden_states
                            .len(),
                    )
                })
                .ok_or_else(DecodeError::projection)?;
            if records > MAX_CONSTRUCTED_STATES {
                return Err(DecodeError::limit(DecodeLimit::States {
                    observed: records,
                    maximum: MAX_CONSTRUCTED_STATES,
                }));
            }
        }
        Ok(Self {
            owner_uid,
            hidden_states,
        })
    }

    #[must_use]
    pub const fn owner_uid(&self) -> UuidSnapshot {
        self.owner_uid
    }

    #[must_use]
    pub fn hidden_states(&self) -> &[HiddenStatesSnapshot] {
        &self.hidden_states
    }

    #[must_use]
    pub fn states(&self) -> &[HiddenStatesSnapshot] {
        self.hidden_states()
    }
}

/// Presence-preserving projection of `TST.TableInfoArchive` fields relevant to
/// hidden axes. Unselected table-info fields remain source-owned.
#[derive(Clone, PartialEq, Eq)]
pub struct TableInfoSnapshot {
    table_model: ReferenceSnapshot,
    view_column_row_uids: Option<ReferenceSnapshot>,
    hidden_states_uuid: Option<UuidSnapshot>,
}

impl fmt::Debug for TableInfoSnapshot {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("TableInfoSnapshot")
            .field(
                "has_view_column_row_uids",
                &self.view_column_row_uids.is_some(),
            )
            .field("has_hidden_states_uuid", &self.hidden_states_uuid.is_some())
            .finish()
    }
}

impl TableInfoSnapshot {
    #[must_use]
    pub const fn new(table_model: ReferenceSnapshot) -> Self {
        Self {
            table_model,
            view_column_row_uids: None,
            hidden_states_uuid: None,
        }
    }

    #[must_use]
    pub const fn with_view_column_row_uids(mut self, value: Option<ReferenceSnapshot>) -> Self {
        self.view_column_row_uids = value;
        self
    }

    #[must_use]
    pub const fn with_hidden_states_uuid(mut self, value: Option<UuidSnapshot>) -> Self {
        self.hidden_states_uuid = value;
        self
    }

    #[must_use]
    pub const fn table_model(&self) -> ReferenceSnapshot {
        self.table_model
    }

    #[must_use]
    pub const fn view_column_row_uids(&self) -> Option<ReferenceSnapshot> {
        self.view_column_row_uids
    }

    #[must_use]
    pub const fn hidden_states_uuid(&self) -> Option<UuidSnapshot> {
        self.hidden_states_uuid
    }
}

/// Presence-preserving projection of the table-model fields used by hidden
/// axes. Required model fields are validated but remain opaque.
#[derive(Clone, PartialEq, Eq)]
pub struct TableModelSnapshot {
    number_of_rows: u32,
    number_of_columns: u32,
    number_of_filtered_rows: Option<u32>,
    number_of_hidden_rows: Option<u32>,
    number_of_hidden_columns: Option<u32>,
    number_of_user_hidden_rows: Option<u32>,
    number_of_user_hidden_columns: Option<u32>,
    hidden_state_formula_owner_for_columns: Option<ReferenceSnapshot>,
    hidden_state_formula_owner_for_rows: Option<ReferenceSnapshot>,
    base_column_row_uids: Option<ReferenceSnapshot>,
    hidden_states_owner: Option<HiddenStatesOwnerSnapshot>,
}

impl fmt::Debug for TableModelSnapshot {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("TableModelSnapshot")
            .field("has_filtered_rows", &self.number_of_filtered_rows.is_some())
            .field("has_hidden_rows", &self.number_of_hidden_rows.is_some())
            .field(
                "has_hidden_columns",
                &self.number_of_hidden_columns.is_some(),
            )
            .field(
                "has_user_hidden_rows",
                &self.number_of_user_hidden_rows.is_some(),
            )
            .field(
                "has_user_hidden_columns",
                &self.number_of_user_hidden_columns.is_some(),
            )
            .field(
                "has_hidden_states_owner",
                &self.hidden_states_owner.is_some(),
            )
            .finish()
    }
}

impl TableModelSnapshot {
    #[must_use]
    pub const fn new(number_of_rows: u32, number_of_columns: u32) -> Self {
        Self {
            number_of_rows,
            number_of_columns,
            number_of_filtered_rows: None,
            number_of_hidden_rows: None,
            number_of_hidden_columns: None,
            number_of_user_hidden_rows: None,
            number_of_user_hidden_columns: None,
            hidden_state_formula_owner_for_columns: None,
            hidden_state_formula_owner_for_rows: None,
            base_column_row_uids: None,
            hidden_states_owner: None,
        }
    }

    #[must_use]
    pub const fn with_number_of_filtered_rows(mut self, value: Option<u32>) -> Self {
        self.number_of_filtered_rows = value;
        self
    }

    #[must_use]
    pub const fn with_number_of_hidden_rows(mut self, value: Option<u32>) -> Self {
        self.number_of_hidden_rows = value;
        self
    }

    #[must_use]
    pub const fn with_number_of_hidden_columns(mut self, value: Option<u32>) -> Self {
        self.number_of_hidden_columns = value;
        self
    }

    #[must_use]
    pub const fn with_number_of_user_hidden_rows(mut self, value: Option<u32>) -> Self {
        self.number_of_user_hidden_rows = value;
        self
    }

    #[must_use]
    pub const fn with_number_of_user_hidden_columns(mut self, value: Option<u32>) -> Self {
        self.number_of_user_hidden_columns = value;
        self
    }

    #[must_use]
    pub const fn with_hidden_state_formula_owner_for_columns(
        mut self,
        value: Option<ReferenceSnapshot>,
    ) -> Self {
        self.hidden_state_formula_owner_for_columns = value;
        self
    }

    #[must_use]
    pub const fn with_hidden_state_formula_owner_for_rows(
        mut self,
        value: Option<ReferenceSnapshot>,
    ) -> Self {
        self.hidden_state_formula_owner_for_rows = value;
        self
    }

    #[must_use]
    pub const fn with_base_column_row_uids(mut self, value: Option<ReferenceSnapshot>) -> Self {
        self.base_column_row_uids = value;
        self
    }

    #[must_use]
    pub fn with_hidden_states_owner(mut self, value: Option<HiddenStatesOwnerSnapshot>) -> Self {
        self.hidden_states_owner = value;
        self
    }

    #[must_use]
    pub const fn number_of_rows(&self) -> u32 {
        self.number_of_rows
    }

    #[must_use]
    pub const fn number_of_columns(&self) -> u32 {
        self.number_of_columns
    }

    #[must_use]
    pub const fn number_of_filtered_rows(&self) -> Option<u32> {
        self.number_of_filtered_rows
    }

    #[must_use]
    pub const fn number_of_hidden_rows(&self) -> Option<u32> {
        self.number_of_hidden_rows
    }

    #[must_use]
    pub const fn number_of_hidden_columns(&self) -> Option<u32> {
        self.number_of_hidden_columns
    }

    #[must_use]
    pub const fn number_of_user_hidden_rows(&self) -> Option<u32> {
        self.number_of_user_hidden_rows
    }

    #[must_use]
    pub const fn number_of_user_hidden_columns(&self) -> Option<u32> {
        self.number_of_user_hidden_columns
    }

    #[must_use]
    pub const fn hidden_state_formula_owner_for_columns(&self) -> Option<ReferenceSnapshot> {
        self.hidden_state_formula_owner_for_columns
    }

    #[must_use]
    pub const fn hidden_state_formula_owner_for_rows(&self) -> Option<ReferenceSnapshot> {
        self.hidden_state_formula_owner_for_rows
    }

    #[must_use]
    pub const fn base_column_row_uids(&self) -> Option<ReferenceSnapshot> {
        self.base_column_row_uids
    }

    #[must_use]
    pub const fn hidden_states_owner(&self) -> Option<&HiddenStatesOwnerSnapshot> {
        self.hidden_states_owner.as_ref()
    }
}

/// Exact strict traversal accounting.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeReport {
    input_bytes: usize,
    output_bytes: usize,
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    states: usize,
    allocations: usize,
    retained_bytes: usize,
    scratch_bytes: usize,
}

impl DecodeReport {
    #[must_use]
    pub const fn input_bytes(self) -> usize {
        self.input_bytes
    }

    #[must_use]
    pub const fn output_bytes(self) -> usize {
        self.output_bytes
    }

    #[must_use]
    pub const fn fields(self) -> usize {
        self.fields
    }

    #[must_use]
    pub const fn work_bytes(self) -> usize {
        self.work_bytes
    }

    #[must_use]
    pub const fn max_depth(self) -> u32 {
        self.max_depth
    }

    #[must_use]
    pub const fn states(self) -> usize {
        self.states
    }

    #[must_use]
    pub const fn allocations(self) -> usize {
        self.allocations
    }

    #[must_use]
    pub const fn retained_bytes(self) -> usize {
        self.retained_bytes
    }

    #[must_use]
    pub const fn scratch_bytes(self) -> usize {
        self.scratch_bytes
    }
}

/// Limits replayed before a prepared candidate allocates output.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RewriteExecutionLimits {
    output_bytes: usize,
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    states: usize,
    allocations: usize,
    retained_bytes: usize,
    scratch_bytes: usize,
}

impl RewriteExecutionLimits {
    #[must_use]
    pub const fn unrestricted() -> Self {
        Self {
            output_bytes: usize::MAX,
            fields: usize::MAX,
            work_bytes: usize::MAX,
            max_depth: u32::MAX,
            states: usize::MAX,
            allocations: usize::MAX,
            retained_bytes: usize::MAX,
            scratch_bytes: usize::MAX,
        }
    }

    #[must_use]
    pub const fn exact(requirements: RewriteExecutionRequirements) -> Self {
        Self {
            output_bytes: requirements.output_bytes,
            fields: requirements.fields,
            work_bytes: requirements.work_bytes,
            max_depth: requirements.max_depth,
            states: requirements.states,
            allocations: requirements.allocations,
            retained_bytes: requirements.retained_bytes,
            scratch_bytes: requirements.scratch_bytes,
        }
    }

    #[must_use]
    pub const fn with_output_bytes(mut self, value: usize) -> Self {
        self.output_bytes = value;
        self
    }

    #[must_use]
    pub const fn with_fields(mut self, value: usize) -> Self {
        self.fields = value;
        self
    }

    #[must_use]
    pub const fn with_work_bytes(mut self, value: usize) -> Self {
        self.work_bytes = value;
        self
    }

    #[must_use]
    pub const fn with_max_depth(mut self, value: u32) -> Self {
        self.max_depth = value;
        self
    }

    #[must_use]
    pub const fn with_states(mut self, value: usize) -> Self {
        self.states = value;
        self
    }

    #[must_use]
    pub const fn with_allocations(mut self, value: usize) -> Self {
        self.allocations = value;
        self
    }

    #[must_use]
    pub const fn with_retained_bytes(mut self, value: usize) -> Self {
        self.retained_bytes = value;
        self
    }

    #[must_use]
    pub const fn with_scratch_bytes(mut self, value: usize) -> Self {
        self.scratch_bytes = value;
        self
    }
}

/// Conservative prepared execution budget.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RewriteExecutionRequirements {
    output_bytes: usize,
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    states: usize,
    allocations: usize,
    retained_bytes: usize,
    scratch_bytes: usize,
}

impl RewriteExecutionRequirements {
    #[must_use]
    pub const fn output_bytes(self) -> usize {
        self.output_bytes
    }

    #[must_use]
    pub const fn fields(self) -> usize {
        self.fields
    }

    #[must_use]
    pub const fn work_bytes(self) -> usize {
        self.work_bytes
    }

    #[must_use]
    pub const fn max_depth(self) -> u32 {
        self.max_depth
    }

    #[must_use]
    pub const fn states(self) -> usize {
        self.states
    }

    #[must_use]
    pub const fn allocations(self) -> usize {
        self.allocations
    }

    #[must_use]
    pub const fn retained_bytes(self) -> usize {
        self.retained_bytes
    }

    #[must_use]
    pub const fn scratch_bytes(self) -> usize {
        self.scratch_bytes
    }
}

/// Candidate bytes and exact execution report.
#[derive(Clone, PartialEq, Eq)]
pub struct RewriteOutput {
    bytes: Vec<u8>,
    report: DecodeReport,
}

impl fmt::Debug for RewriteOutput {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("RewriteOutput")
            .field("bytes_len", &self.bytes.len())
            .field("report", &self.report)
            .finish()
    }
}

impl RewriteOutput {
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.bytes
    }

    #[must_use]
    pub fn output(&self) -> &[u8] {
        self.bytes()
    }

    #[must_use]
    pub fn into_bytes(self) -> Vec<u8> {
        self.bytes
    }

    #[must_use]
    pub const fn report(&self) -> DecodeReport {
        self.report
    }
}

/// Prepared source-preserving rewrite of one hidden-state graph message.
#[derive(Clone)]
pub struct PreparedRewrite<'source> {
    source: &'source [u8],
    kind: RewriteKind,
    desired: RewriteDesired,
    index_plan: RewriteIndexPlan,
    requirements: RewriteExecutionRequirements,
    report: DecodeReport,
}

impl fmt::Debug for PreparedRewrite<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("PreparedRewrite")
            .field("kind", &self.kind)
            .field("requirements", &self.requirements)
            .field("report", &self.report)
            .finish()
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum RewriteKind {
    TableInfo,
    TableModel,
    HiddenStatesOwner,
    HiddenStateExtent,
    RowOrColumnState,
}

#[derive(Debug, Clone)]
enum RewriteDesired {
    TableInfo(TableInfoSnapshot),
    TableModel(TableModelSnapshot),
    HiddenStatesOwner(HiddenStatesOwnerSnapshot),
    HiddenStateExtent(HiddenStateExtentSnapshot),
    RowOrColumnState(RowOrColumnStateSnapshot),
}

impl<'source> PreparedRewrite<'source> {
    /// Return the preparation report.
    #[must_use]
    pub const fn prepare_report(&self) -> DecodeReport {
        self.report
    }

    /// Return conservative ceilings that execution will consume.
    #[must_use]
    pub const fn execution_requirements(&self) -> RewriteExecutionRequirements {
        self.requirements
    }

    /// Execute after caller-supplied ceilings pass.
    pub fn execute(self, limits: RewriteExecutionLimits) -> Result<RewriteOutput, DecodeError> {
        check_execution_limits(self.requirements, limits)?;
        let mut output = Vec::new();
        output
            .try_reserve_exact(self.requirements.output_bytes)
            .map_err(|_| DecodeError::allocation(self.requirements.output_bytes))?;
        let mut matched = new_match_marks(&self.index_plan)?;
        emit_rewrite(
            self.source,
            self.kind,
            &self.desired,
            &self.index_plan,
            &mut matched,
            &mut output,
        )?;
        if output.len() != self.requirements.output_bytes {
            return Err(DecodeError::projection());
        }
        let verify_options = DecodeOptions::new(
            output.len().max(1),
            output.len().max(1),
            limits.fields,
            limits.work_bytes,
            limits.max_depth.min(MAX_RECURSION),
            limits.states,
        )
        .with_max_allocations(limits.allocations)
        .with_max_retained_bytes(limits.retained_bytes)
        .with_max_scratch_bytes(limits.scratch_bytes);
        let kind = self.kind;
        let desired = self.desired;
        let verified = match (kind, desired) {
            (RewriteKind::TableInfo, RewriteDesired::TableInfo(value)) => {
                decode_table_info(&output, verify_options)? == value
            },
            (RewriteKind::TableModel, RewriteDesired::TableModel(value)) => {
                decode_table_model(&output, verify_options)? == value
            },
            (RewriteKind::HiddenStatesOwner, RewriteDesired::HiddenStatesOwner(value)) => {
                decode_hidden_states_owner(&output, verify_options)? == value
            },
            (RewriteKind::HiddenStateExtent, RewriteDesired::HiddenStateExtent(value)) => {
                decode_hidden_state_extent(&output, verify_options)? == value
            },
            (RewriteKind::RowOrColumnState, RewriteDesired::RowOrColumnState(value)) => {
                decode_row_or_column_state(&output, verify_options)? == value
            },
            _ => return Err(DecodeError::projection()),
        };
        if !verified {
            return Err(DecodeError::projection());
        }
        Ok(RewriteOutput {
            bytes: output,
            report: self.report,
        })
    }
}

/// Decode a table-info hidden-state projection.
pub fn decode_table_info(
    source: &[u8],
    options: DecodeOptions,
) -> Result<TableInfoSnapshot, DecodeError> {
    Ok(decode_table_info_with_report(source, options)?.0)
}

/// Decode table-info and return strict resource accounting.
pub fn decode_table_info_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(TableInfoSnapshot, DecodeReport), DecodeError> {
    let mut budget = Budget::new(options);
    validate_input(source, options)?;
    let parsed = parse_table_info(source, options, &mut budget)?;
    cross_check_table_info(source, &parsed, options)?;
    Ok((parsed.snapshot, budget.report(source.len(), source.len())))
}

/// Decode a table-model hidden-state projection.
pub fn decode_table_model(
    source: &[u8],
    options: DecodeOptions,
) -> Result<TableModelSnapshot, DecodeError> {
    Ok(decode_table_model_with_report(source, options)?.0)
}

/// Decode table-model hidden-state fields and return strict resource accounting.
pub fn decode_table_model_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(TableModelSnapshot, DecodeReport), DecodeError> {
    let mut budget = Budget::new(options);
    validate_input(source, options)?;
    let parsed = parse_table_model(source, options, &mut budget)?;
    cross_check_table_model(source, &parsed, options)?;
    Ok((parsed.snapshot, budget.report(source.len(), source.len())))
}

/// Decode one hidden-state owner and all bounded nested extents/states.
pub fn decode_hidden_states_owner(
    source: &[u8],
    options: DecodeOptions,
) -> Result<HiddenStatesOwnerSnapshot, DecodeError> {
    Ok(decode_hidden_states_owner_with_report(source, options)?.0)
}

/// Decode one hidden-state owner and return strict resource accounting.
pub fn decode_hidden_states_owner_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(HiddenStatesOwnerSnapshot, DecodeReport), DecodeError> {
    let mut budget = Budget::new(options);
    validate_input(source, options)?;
    let parsed = parse_owner(source, options, &mut budget, 1)?;
    cross_check_owner(source, &parsed, options)?;
    Ok((parsed.snapshot, budget.report(source.len(), source.len())))
}

/// Decode one extent independently. Repeated state records are bounded and
/// owned by the returned semantic snapshot rather than a generated view.
pub fn decode_hidden_state_extent(
    source: &[u8],
    options: DecodeOptions,
) -> Result<HiddenStateExtentSnapshot, DecodeError> {
    Ok(decode_hidden_state_extent_with_report(source, options)?.0)
}

/// Decode one extent and return strict resource accounting.
pub fn decode_hidden_state_extent_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(HiddenStateExtentSnapshot, DecodeReport), DecodeError> {
    let mut budget = Budget::new(options);
    validate_input(source, options)?;
    let parsed = parse_extent(source, options, &mut budget, 1)?;
    cross_check_extent(source, &parsed, options)?;
    Ok((parsed.snapshot, budget.report(source.len(), source.len())))
}

/// Decode one row/column state record.
pub fn decode_row_or_column_state(
    source: &[u8],
    options: DecodeOptions,
) -> Result<RowOrColumnStateSnapshot, DecodeError> {
    Ok(decode_row_or_column_state_with_report(source, options)?.0)
}

/// Decode one row/column state and return strict resource accounting.
pub fn decode_row_or_column_state_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(RowOrColumnStateSnapshot, DecodeReport), DecodeError> {
    let mut budget = Budget::new(options);
    validate_input(source, options)?;
    let parsed = parse_row_state(source, options, &mut budget, 1)?;
    cross_check_row_state(source, &parsed, options)?;
    Ok((parsed.snapshot, budget.report(source.len(), source.len())))
}

/// Qualify the current Pages column/row UID-map object through the existing
/// strict physical-sort codec.  This adapter intentionally accepts the
/// audited codec's finite policy rather than inventing a second collection
/// implementation here.
pub fn qualify_column_row_uid_map(
    source: &[u8],
    column_count: usize,
    row_count: usize,
    options: crate::numbers_table_physical_sort_codec::DecodeOptions,
) -> Result<crate::numbers_table_physical_sort_codec::ColumnRowUidMapSnapshot, DecodeError> {
    Ok(qualify_column_row_uid_map_with_report(source, column_count, row_count, options)?.0)
}

/// Qualify a UID map and merge the child codec's bounded traversal report.
pub fn qualify_column_row_uid_map_with_report(
    source: &[u8],
    column_count: usize,
    row_count: usize,
    options: crate::numbers_table_physical_sort_codec::DecodeOptions,
) -> Result<
    (
        crate::numbers_table_physical_sort_codec::ColumnRowUidMapSnapshot,
        DecodeReport,
    ),
    DecodeError,
> {
    let local_options = DecodeOptions::new(
        options.max_message_bytes(),
        options.max_output_bytes().max(1),
        options.max_fields(),
        options.max_work_bytes(),
        options.recursion_limit(),
        options
            .max_records()
            .max(options.max_elements())
            .min(MAX_CONSTRUCTED_STATES),
    )
    .with_max_allocations(usize::MAX)
    .with_max_retained_bytes(usize::MAX)
    .with_max_scratch_bytes(options.max_scratch_bytes());
    let mut budget = Budget::new(local_options);
    validate_input(source, local_options)?;
    let (snapshot, child_report) =
        crate::numbers_table_physical_sort_codec::decode_column_row_uid_map(
            source,
            column_count,
            row_count,
            options,
        )
        .map_err(map_uid_map_error)?;
    budget.merge_uid_map_report(child_report)?;
    Ok((snapshot, budget.report(source.len(), source.len())))
}

/// Qualify a type-4008 formula-dependency object through the existing strict
/// dependency codec and return only its selected identity edges.
pub fn decode_formula_owner_dependencies(
    source: &[u8],
    options: DecodeOptions,
) -> Result<FormulaOwnerDependenciesSnapshot, DecodeError> {
    Ok(decode_formula_owner_dependencies_with_report(source, options)?.0)
}

/// Strictly qualify a formula-dependency object and return aggregate counts.
pub fn decode_formula_owner_dependencies_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(FormulaOwnerDependenciesSnapshot, DecodeReport), DecodeError> {
    let mut budget = Budget::new(options);
    validate_input(source, options)?;
    let parsed = parse_formula_owner_dependencies(source, options, &mut budget, 1)?;
    cross_check_formula_owner_dependencies(source, &parsed, options)?;
    // The dependency codec owns the repeated dependency closure. Run it after
    // this bounded identity pass so a Pages router cannot silently fall back
    // to a complete generated message for ownership discovery.
    let (_dependency_snapshot, dependency_report) =
        crate::numbers_table_cell_dependency_codec::decode_formula_owner_dependencies_with_report(
            source,
            dependency_options(options),
        )
        .map_err(map_dependency_error)?;
    budget.merge_dependency_report(dependency_report)?;
    Ok((parsed, budget.report(source.len(), source.len())))
}

/// Decode one type-6204 hidden-state formula-owner object.
pub fn decode_hidden_state_formula_owner(
    source: &[u8],
    options: DecodeOptions,
) -> Result<HiddenStateFormulaOwnerSnapshot, DecodeError> {
    Ok(decode_hidden_state_formula_owner_with_report(source, options)?.0)
}

/// Decode one type-6204 hidden-state formula-owner object with accounting.
pub fn decode_hidden_state_formula_owner_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(HiddenStateFormulaOwnerSnapshot, DecodeReport), DecodeError> {
    let mut budget = Budget::new(options);
    validate_input(source, options)?;
    let parsed = parse_hidden_state_formula_owner(source, options, &mut budget, 1)?;
    cross_check_hidden_state_formula_owner(source, &parsed, options)?;
    Ok((parsed, budget.report(source.len(), source.len())))
}

/// Decode one type-6220 filter-set object. Rule collections are structurally
/// scanned and the selected offsets are streamed into a bounded owned vector.
pub fn decode_filter_set(
    source: &[u8],
    options: DecodeOptions,
) -> Result<FilterSetSnapshot, DecodeError> {
    Ok(decode_filter_set_with_report(source, options)?.0)
}

/// Decode one type-6220 filter-set object with accounting.
pub fn decode_filter_set_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(FilterSetSnapshot, DecodeReport), DecodeError> {
    let mut budget = Budget::new(options);
    validate_input(source, options)?;
    let parsed = parse_filter_set(source, options, &mut budget, 1)?;
    cross_check_filter_set(source, &parsed, options)?;
    Ok((parsed, budget.report(source.len(), source.len())))
}

/// Prepare a source-preserving table-info rewrite.
pub fn prepare_table_info_rewrite<'source>(
    source: &'source [u8],
    desired: &TableInfoSnapshot,
    options: DecodeOptions,
) -> Result<PreparedRewrite<'source>, DecodeError> {
    validate_table_info_desired(desired, options)?;
    prepare_rewrite(
        source,
        RewriteKind::TableInfo,
        RewriteDesired::TableInfo(desired.clone()),
        options,
    )
}

/// Prepare a source-preserving table-model rewrite.
pub fn prepare_table_model_rewrite<'source>(
    source: &'source [u8],
    desired: &TableModelSnapshot,
    options: DecodeOptions,
) -> Result<PreparedRewrite<'source>, DecodeError> {
    validate_table_model_desired(desired, options)?;
    prepare_rewrite(
        source,
        RewriteKind::TableModel,
        RewriteDesired::TableModel(desired.clone()),
        options,
    )
}

/// Prepare a source-preserving hidden-state owner rewrite.
pub fn prepare_hidden_states_owner_rewrite<'source>(
    source: &'source [u8],
    desired: &HiddenStatesOwnerSnapshot,
    options: DecodeOptions,
) -> Result<PreparedRewrite<'source>, DecodeError> {
    validate_owner_desired(desired, options)?;
    prepare_rewrite(
        source,
        RewriteKind::HiddenStatesOwner,
        RewriteDesired::HiddenStatesOwner(desired.clone()),
        options,
    )
}

/// Prepare a source-preserving extent rewrite.
pub fn prepare_hidden_state_extent_rewrite<'source>(
    source: &'source [u8],
    desired: &HiddenStateExtentSnapshot,
    options: DecodeOptions,
) -> Result<PreparedRewrite<'source>, DecodeError> {
    validate_extent_desired(desired, options)?;
    prepare_rewrite(
        source,
        RewriteKind::HiddenStateExtent,
        RewriteDesired::HiddenStateExtent(desired.clone()),
        options,
    )
}

/// Prepare a source-preserving row/column state rewrite.
pub fn prepare_row_or_column_state_rewrite<'source>(
    source: &'source [u8],
    desired: RowOrColumnStateSnapshot,
    options: DecodeOptions,
) -> Result<PreparedRewrite<'source>, DecodeError> {
    validate_row_state_desired(options)?;
    prepare_rewrite(
        source,
        RewriteKind::RowOrColumnState,
        RewriteDesired::RowOrColumnState(desired),
        options,
    )
}

/// Rewrite one table-info payload and return candidate bytes plus accounting.
pub fn rewrite_table_info(
    source: &[u8],
    desired: &TableInfoSnapshot,
    options: DecodeOptions,
) -> Result<RewriteOutput, DecodeError> {
    let prepared = prepare_table_info_rewrite(source, desired, options)?;
    let limits = RewriteExecutionLimits::exact(prepared.execution_requirements());
    prepared.execute(limits)
}

/// Rewrite one table-model payload and return candidate bytes plus accounting.
pub fn rewrite_table_model(
    source: &[u8],
    desired: &TableModelSnapshot,
    options: DecodeOptions,
) -> Result<RewriteOutput, DecodeError> {
    let prepared = prepare_table_model_rewrite(source, desired, options)?;
    let limits = RewriteExecutionLimits::exact(prepared.execution_requirements());
    prepared.execute(limits)
}

/// Rewrite one hidden-state owner payload and return candidate bytes plus accounting.
pub fn rewrite_hidden_states_owner(
    source: &[u8],
    desired: &HiddenStatesOwnerSnapshot,
    options: DecodeOptions,
) -> Result<RewriteOutput, DecodeError> {
    let prepared = prepare_hidden_states_owner_rewrite(source, desired, options)?;
    let limits = RewriteExecutionLimits::exact(prepared.execution_requirements());
    prepared.execute(limits)
}

/// Rewrite one extent payload and return candidate bytes plus accounting.
pub fn rewrite_hidden_state_extent(
    source: &[u8],
    desired: &HiddenStateExtentSnapshot,
    options: DecodeOptions,
) -> Result<RewriteOutput, DecodeError> {
    let prepared = prepare_hidden_state_extent_rewrite(source, desired, options)?;
    let limits = RewriteExecutionLimits::exact(prepared.execution_requirements());
    prepared.execute(limits)
}

/// Rewrite one row/column state payload and return candidate bytes plus accounting.
pub fn rewrite_row_or_column_state(
    source: &[u8],
    desired: RowOrColumnStateSnapshot,
    options: DecodeOptions,
) -> Result<RewriteOutput, DecodeError> {
    let prepared = prepare_row_or_column_state_rewrite(source, desired, options)?;
    let limits = RewriteExecutionLimits::exact(prepared.execution_requirements());
    prepared.execute(limits)
}

// Canonical names used by package adapters that call the native archive type
// explicitly. Keep these aliases source-compatible with the focused codecs.
pub use decode_hidden_states_owner as decode_hidden_state_owner;
pub use decode_hidden_states_owner_with_report as decode_hidden_state_owner_with_report;
pub use prepare_hidden_states_owner_rewrite as prepare_hidden_state_owner_rewrite;
pub use rewrite_hidden_states_owner as rewrite_hidden_state_owner;

#[derive(Debug)]
struct Budget {
    options: DecodeOptions,
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    states: usize,
    allocations: usize,
    retained_bytes: usize,
    scratch_bytes: usize,
}

impl Budget {
    const fn new(options: DecodeOptions) -> Self {
        Self {
            options,
            fields: 0,
            work_bytes: 0,
            max_depth: 1,
            states: 0,
            allocations: 0,
            retained_bytes: 0,
            scratch_bytes: 0,
        }
    }

    fn field(&mut self) -> Result<(), DecodeError> {
        self.fields = self
            .fields
            .checked_add(1)
            .ok_or_else(DecodeError::projection)?;
        if self.fields > self.options.max_fields {
            return Err(DecodeError::limit(DecodeLimit::Fields {
                observed: self.fields,
                maximum: self.options.max_fields,
            }));
        }
        Ok(())
    }

    fn work(&mut self, bytes: usize) -> Result<(), DecodeError> {
        let bytes = bytes.checked_mul(2).ok_or_else(DecodeError::projection)?;
        self.work_bytes = self
            .work_bytes
            .checked_add(bytes)
            .ok_or_else(DecodeError::projection)?;
        if self.work_bytes > self.options.max_work_bytes {
            return Err(DecodeError::limit(DecodeLimit::WorkBytes {
                observed: self.work_bytes,
                maximum: self.options.max_work_bytes,
            }));
        }
        Ok(())
    }

    fn depth(&mut self, depth: u32) -> Result<(), DecodeError> {
        self.max_depth = self.max_depth.max(depth);
        if depth > self.options.recursion_limit || depth > MAX_RECURSION {
            return Err(DecodeError::limit(DecodeLimit::Nesting {
                observed: depth,
                maximum: self.options.recursion_limit.min(MAX_RECURSION),
            }));
        }
        Ok(())
    }

    fn state(&mut self) -> Result<(), DecodeError> {
        self.states = self
            .states
            .checked_add(1)
            .ok_or_else(DecodeError::projection)?;
        let maximum = self.options.max_states.min(MAX_CONSTRUCTED_STATES);
        if self.states > maximum {
            return Err(DecodeError::limit(DecodeLimit::States {
                observed: self.states,
                maximum,
            }));
        }
        Ok(())
    }

    fn allocation(&mut self, bytes: usize) -> Result<(), DecodeError> {
        self.allocations = self
            .allocations
            .checked_add(1)
            .ok_or_else(DecodeError::projection)?;
        if self.allocations > self.options.max_allocations {
            return Err(DecodeError::limit(DecodeLimit::Allocations {
                observed: self.allocations,
                maximum: self.options.max_allocations,
            }));
        }
        self.scratch_bytes = self
            .scratch_bytes
            .checked_add(bytes)
            .ok_or_else(DecodeError::projection)?;
        if self.scratch_bytes > self.options.max_scratch_bytes {
            return Err(DecodeError::limit(DecodeLimit::ScratchBytes {
                observed: self.scratch_bytes,
                maximum: self.options.max_scratch_bytes,
            }));
        }
        Ok(())
    }

    fn retain(&mut self, bytes: usize) -> Result<(), DecodeError> {
        self.retained_bytes = self
            .retained_bytes
            .checked_add(bytes)
            .ok_or_else(DecodeError::projection)?;
        if self.retained_bytes > self.options.max_retained_bytes {
            return Err(DecodeError::limit(DecodeLimit::RetainedBytes {
                observed: self.retained_bytes,
                maximum: self.options.max_retained_bytes,
            }));
        }
        Ok(())
    }

    fn merge_dependency_report(
        &mut self,
        report: crate::numbers_table_cell_dependency_codec::DecodeReport,
    ) -> Result<(), DecodeError> {
        self.fields = self
            .fields
            .checked_add(report.fields())
            .ok_or_else(DecodeError::projection)?;
        if self.fields > self.options.max_fields {
            return Err(DecodeError::limit(DecodeLimit::Fields {
                observed: self.fields,
                maximum: self.options.max_fields,
            }));
        }
        self.work_bytes = self
            .work_bytes
            .checked_add(report.work_bytes())
            .ok_or_else(DecodeError::projection)?;
        if self.work_bytes > self.options.max_work_bytes {
            return Err(DecodeError::limit(DecodeLimit::WorkBytes {
                observed: self.work_bytes,
                maximum: self.options.max_work_bytes,
            }));
        }
        self.max_depth = self.max_depth.max(report.max_depth());
        self.states = self
            .states
            .checked_add(report.references())
            .ok_or_else(DecodeError::projection)?;
        let maximum = self.options.max_states.min(MAX_CONSTRUCTED_STATES);
        if self.states > maximum {
            return Err(DecodeError::limit(DecodeLimit::States {
                observed: self.states,
                maximum,
            }));
        }
        self.retain(report.reference_bytes())?;
        self.retain(report.text_bytes())?;
        Ok(())
    }

    fn merge_uid_map_report(
        &mut self,
        report: crate::numbers_table_physical_sort_codec::DecodeReport,
    ) -> Result<(), DecodeError> {
        self.fields = self
            .fields
            .checked_add(report.fields())
            .ok_or_else(DecodeError::projection)?;
        if self.fields > self.options.max_fields {
            return Err(DecodeError::limit(DecodeLimit::Fields {
                observed: self.fields,
                maximum: self.options.max_fields,
            }));
        }
        self.work_bytes = self
            .work_bytes
            .checked_add(report.work_bytes())
            .ok_or_else(DecodeError::projection)?;
        if self.work_bytes > self.options.max_work_bytes {
            return Err(DecodeError::limit(DecodeLimit::WorkBytes {
                observed: self.work_bytes,
                maximum: self.options.max_work_bytes,
            }));
        }
        self.max_depth = self.max_depth.max(report.max_depth());
        self.states = self
            .states
            .checked_add(report.records())
            .and_then(|value| value.checked_add(report.elements()))
            .ok_or_else(DecodeError::projection)?;
        let maximum = self.options.max_states.min(MAX_CONSTRUCTED_STATES);
        if self.states > maximum {
            return Err(DecodeError::limit(DecodeLimit::States {
                observed: self.states,
                maximum,
            }));
        }
        Ok(())
    }

    fn report(&self, input_bytes: usize, output_bytes: usize) -> DecodeReport {
        DecodeReport {
            input_bytes,
            output_bytes,
            fields: self.fields,
            work_bytes: self.work_bytes,
            max_depth: self.max_depth,
            states: self.states,
            allocations: self.allocations,
            retained_bytes: self.retained_bytes,
            scratch_bytes: self.scratch_bytes,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct FieldSpan {
    number: u32,
    wire: u8,
    start: usize,
    end: usize,
    value_start: usize,
    value_end: usize,
    value_canonical: bool,
    length_canonical: bool,
}

#[derive(Debug, Clone, Copy)]
struct Varint {
    value: u64,
    canonical: bool,
}

#[derive(Debug)]
struct ParsedTableInfo<'source> {
    snapshot: TableInfoSnapshot,
    _source: core::marker::PhantomData<&'source [u8]>,
}

#[derive(Debug)]
struct ParsedTableModel<'source> {
    snapshot: TableModelSnapshot,
    _source: core::marker::PhantomData<&'source [u8]>,
}

#[derive(Debug)]
struct ParsedOwner<'source> {
    snapshot: HiddenStatesOwnerSnapshot,
    _source: core::marker::PhantomData<&'source [u8]>,
}

#[derive(Debug)]
struct ParsedHiddenStates<'source> {
    snapshot: HiddenStatesSnapshot,
    _source: core::marker::PhantomData<&'source [u8]>,
}

#[derive(Debug)]
struct ParsedExtent<'source> {
    snapshot: HiddenStateExtentSnapshot,
    _source: core::marker::PhantomData<&'source [u8]>,
}

#[derive(Debug)]
struct ParsedState<'source> {
    snapshot: RowOrColumnStateSnapshot,
    _source: core::marker::PhantomData<&'source [u8]>,
}

fn dependency_options(
    options: DecodeOptions,
) -> crate::numbers_table_cell_storage_codec::DecodeOptions {
    crate::numbers_table_cell_storage_codec::DecodeOptions::new(
        options.max_input_bytes,
        options.max_fields,
        options.max_work_bytes,
        options.recursion_limit,
        options.max_states,
        options.max_input_bytes,
    )
}

fn map_dependency_error(
    error: crate::numbers_table_cell_dependency_codec::DecodeError,
) -> DecodeError {
    let Some(limit) = error.resource_limit() else {
        return DecodeError::invalid("formula-dependency strict codec rejected payload");
    };
    match limit {
        crate::numbers_table_cell_dependency_codec::DecodeLimit::Bytes { observed, maximum } => {
            DecodeError::limit(DecodeLimit::InputBytes { observed, maximum })
        },
        crate::numbers_table_cell_dependency_codec::DecodeLimit::References {
            observed,
            maximum,
        } => DecodeError::limit(DecodeLimit::States { observed, maximum }),
        crate::numbers_table_cell_dependency_codec::DecodeLimit::Text { observed, maximum } => {
            DecodeError::limit(DecodeLimit::RetainedBytes { observed, maximum })
        },
        crate::numbers_table_cell_dependency_codec::DecodeLimit::Fields { observed, maximum } => {
            DecodeError::limit(DecodeLimit::Fields { observed, maximum })
        },
        crate::numbers_table_cell_dependency_codec::DecodeLimit::Work { observed, maximum } => {
            DecodeError::limit(DecodeLimit::WorkBytes { observed, maximum })
        },
        crate::numbers_table_cell_dependency_codec::DecodeLimit::Nesting { observed, maximum } => {
            DecodeError::limit(DecodeLimit::Nesting { observed, maximum })
        },
        crate::numbers_table_cell_dependency_codec::DecodeLimit::Allocation { requested } => {
            DecodeError::allocation(requested)
        },
        crate::numbers_table_cell_dependency_codec::DecodeLimit::Retained { observed, maximum } => {
            DecodeError::limit(DecodeLimit::RetainedBytes { observed, maximum })
        },
    }
}

fn map_uid_map_error(error: crate::numbers_table_physical_sort_codec::DecodeError) -> DecodeError {
    let Some(limit) = error.resource_limit() else {
        return DecodeError::invalid("column/row UID-map strict codec rejected payload");
    };
    match limit {
        crate::numbers_table_physical_sort_codec::DecodeLimit::Bytes { observed, maximum } => {
            DecodeError::limit(DecodeLimit::InputBytes { observed, maximum })
        },
        crate::numbers_table_physical_sort_codec::DecodeLimit::Fields { observed, maximum } => {
            DecodeError::limit(DecodeLimit::Fields { observed, maximum })
        },
        crate::numbers_table_physical_sort_codec::DecodeLimit::Work { observed, maximum } => {
            DecodeError::limit(DecodeLimit::WorkBytes { observed, maximum })
        },
        crate::numbers_table_physical_sort_codec::DecodeLimit::Nesting { observed, maximum } => {
            DecodeError::limit(DecodeLimit::Nesting { observed, maximum })
        },
        crate::numbers_table_physical_sort_codec::DecodeLimit::Records { observed, maximum }
        | crate::numbers_table_physical_sort_codec::DecodeLimit::Elements { observed, maximum } => {
            DecodeError::limit(DecodeLimit::States { observed, maximum })
        },
        crate::numbers_table_physical_sort_codec::DecodeLimit::OutputBytes {
            observed,
            maximum,
        } => DecodeError::limit(DecodeLimit::OutputBytes { observed, maximum }),
        crate::numbers_table_physical_sort_codec::DecodeLimit::ScratchBytes {
            observed,
            maximum,
        } => DecodeError::limit(DecodeLimit::ScratchBytes { observed, maximum }),
        crate::numbers_table_physical_sort_codec::DecodeLimit::Allocation { requested } => {
            DecodeError::allocation(requested)
        },
    }
}

fn validate_input(source: &[u8], options: DecodeOptions) -> Result<(), DecodeError> {
    if source.len() > options.max_input_bytes {
        return Err(DecodeError::limit(DecodeLimit::InputBytes {
            observed: source.len(),
            maximum: options.max_input_bytes,
        }));
    }
    if source.len() > options.max_output_bytes {
        return Err(DecodeError::limit(DecodeLimit::OutputBytes {
            observed: source.len(),
            maximum: options.max_output_bytes,
        }));
    }
    if options.recursion_limit == 0 || options.recursion_limit > MAX_RECURSION {
        return Err(DecodeError::limit(DecodeLimit::Nesting {
            observed: options.recursion_limit,
            maximum: MAX_RECURSION,
        }));
    }
    Ok(())
}

fn scan_fields(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
    depth: u32,
) -> Result<Vec<FieldSpan>, DecodeError> {
    budget.depth(depth)?;
    budget.work(source.len())?;
    let mut fields = Vec::new();
    let lower = source.len() / 2;
    let reserve = lower.min(options.max_fields).min(32);
    if reserve != 0 {
        let bytes = reserve
            .checked_mul(size_of::<FieldSpan>())
            .ok_or_else(DecodeError::projection)?;
        budget.allocation(bytes)?;
        fields
            .try_reserve_exact(reserve)
            .map_err(|_| DecodeError::allocation(bytes))?;
    }
    let mut offset = 0;
    while offset < source.len() {
        let field = parse_field(source, &mut offset, budget, depth)?;
        if fields.len() == fields.capacity() {
            let additional = fields.capacity().max(1).min(
                options
                    .max_fields
                    .checked_sub(fields.len())
                    .unwrap_or(0)
                    .max(1),
            );
            let bytes = additional
                .checked_mul(size_of::<FieldSpan>())
                .ok_or_else(DecodeError::projection)?;
            budget.allocation(bytes)?;
            fields
                .try_reserve_exact(additional)
                .map_err(|_| DecodeError::allocation(bytes))?;
        }
        fields.push(field);
    }
    Ok(fields)
}

fn parse_field(
    source: &[u8],
    offset: &mut usize,
    budget: &mut Budget,
    depth: u32,
) -> Result<FieldSpan, DecodeError> {
    let start = *offset;
    let key = read_varint(source, offset)?;
    if !key.canonical {
        return Err(DecodeError::noncanonical("protobuf field key"));
    }
    let raw = u32::try_from(key.value).map_err(|_| buffa::DecodeError::InvalidFieldNumber)?;
    let number = raw >> 3;
    let wire = u8::try_from(raw & 7).map_err(|_| buffa::DecodeError::InvalidWireType(raw & 7))?;
    if number == 0 || number > buffa::encoding::MAX_FIELD_NUMBER || wire == 4 {
        return Err(buffa::DecodeError::InvalidFieldNumber.into());
    }
    budget.field()?;
    // For length-delimited fields, the selected value starts after the
    // length varint. Scalar/fixed/group values start immediately after the
    // key, so this default remains correct for those wire forms.
    let mut value_start = *offset;
    let (value_end, value_canonical, length_canonical) = match wire {
        0 => {
            let value = read_varint(source, offset)?;
            (*offset, value.canonical, true)
        },
        1 => {
            let end = offset.checked_add(8).ok_or_else(DecodeError::projection)?;
            if end > source.len() {
                return Err(buffa::DecodeError::UnexpectedEof.into());
            }
            *offset = end;
            (end, true, true)
        },
        2 => {
            let length = read_varint(source, offset)?;
            let length_value =
                usize::try_from(length.value).map_err(|_| buffa::DecodeError::MessageTooLarge)?;
            let payload_start = *offset;
            value_start = payload_start;
            let end = payload_start
                .checked_add(length_value)
                .ok_or_else(DecodeError::projection)?;
            if end > source.len() {
                return Err(buffa::DecodeError::UnexpectedEof.into());
            }
            *offset = end;
            (end, true, length.canonical)
        },
        3 => {
            let end = skip_group(source, offset, budget, depth, number)?;
            (end, true, true)
        },
        5 => {
            let end = offset.checked_add(4).ok_or_else(DecodeError::projection)?;
            if end > source.len() {
                return Err(buffa::DecodeError::UnexpectedEof.into());
            }
            *offset = end;
            (end, true, true)
        },
        _ => return Err(buffa::DecodeError::InvalidWireType(raw & 7).into()),
    };
    Ok(FieldSpan {
        number,
        wire,
        start,
        end: *offset,
        value_start,
        value_end,
        value_canonical,
        length_canonical,
    })
}

fn child_depth(depth: u32) -> Result<u32, DecodeError> {
    depth.checked_add(1).ok_or_else(DecodeError::projection)
}

fn skip_group(
    source: &[u8],
    offset: &mut usize,
    budget: &mut Budget,
    depth: u32,
    opening_number: u32,
) -> Result<usize, DecodeError> {
    budget.depth(child_depth(depth)?)?;
    loop {
        if *offset >= source.len() {
            return Err(buffa::DecodeError::UnexpectedEof.into());
        }
        let key = read_varint(source, offset)?;
        if !key.canonical {
            return Err(DecodeError::noncanonical("protobuf group field key"));
        }
        let raw = u32::try_from(key.value).map_err(|_| buffa::DecodeError::InvalidFieldNumber)?;
        let number = raw >> 3;
        let wire = raw & 7;
        if number == 0 || number > buffa::encoding::MAX_FIELD_NUMBER {
            return Err(buffa::DecodeError::InvalidFieldNumber.into());
        }
        budget.field()?;
        match wire {
            4 if number == opening_number => return Ok(*offset),
            4 => return Err(buffa::DecodeError::InvalidEndGroup(number).into()),
            3 => {
                // `depth` identifies the containing message/group.  Pass the
                // next level into the recursive call so its own
                // `child_depth` check observes one additional nesting level
                // for every nested group.  Reusing `depth` here would let an
                // arbitrarily deep unknown-group chain bypass the strict
                // recursion ceiling.
                skip_group(source, offset, budget, child_depth(depth)?, number)?;
            },
            0 => {
                let _ = read_varint(source, offset)?;
            },
            1 => {
                let end = offset.checked_add(8).ok_or_else(DecodeError::projection)?;
                if end > source.len() {
                    return Err(buffa::DecodeError::UnexpectedEof.into());
                }
                *offset = end;
            },
            2 => {
                let length = read_varint(source, offset)?;
                if !length.canonical {
                    return Err(DecodeError::noncanonical("protobuf group length"));
                }
                let length = usize::try_from(length.value)
                    .map_err(|_| buffa::DecodeError::MessageTooLarge)?;
                let end = offset
                    .checked_add(length)
                    .ok_or_else(DecodeError::projection)?;
                if end > source.len() {
                    return Err(buffa::DecodeError::UnexpectedEof.into());
                }
                *offset = end;
            },
            5 => {
                let end = offset.checked_add(4).ok_or_else(DecodeError::projection)?;
                if end > source.len() {
                    return Err(buffa::DecodeError::UnexpectedEof.into());
                }
                *offset = end;
            },
            _ => return Err(buffa::DecodeError::InvalidWireType(wire).into()),
        }
    }
}

fn read_varint(source: &[u8], offset: &mut usize) -> Result<Varint, DecodeError> {
    let start = *offset;
    let mut value = 0u64;
    for index in 0..10usize {
        let byte = *source
            .get(*offset)
            .ok_or(buffa::DecodeError::UnexpectedEof)?;
        *offset = offset.checked_add(1).ok_or_else(DecodeError::projection)?;
        if index == 9 && byte > 1 {
            return Err(buffa::DecodeError::VarintTooLong.into());
        }
        value |= u64::from(byte & 0x7f) << (index * 7);
        if byte & 0x80 == 0 {
            let consumed = *offset - start;
            return Ok(Varint {
                value,
                canonical: consumed == varint_len(value),
            });
        }
    }
    Err(buffa::DecodeError::VarintTooLong.into())
}

const fn varint_len(mut value: u64) -> usize {
    let mut length = 1;
    while value >= 128 {
        value >>= 7;
        length += 1;
    }
    length
}

fn field_wire(field: FieldSpan, expected: u8) -> Result<(), DecodeError> {
    if field.wire != expected {
        return Err(buffa::DecodeError::WireTypeMismatch {
            field_number: field.number,
            expected,
            actual: field.wire,
        }
        .into());
    }
    if expected == 2 && !field.length_canonical {
        return Err(DecodeError::noncanonical("length-delimited size"));
    }
    if expected == 0 && !field.value_canonical {
        return Err(DecodeError::noncanonical("protobuf varint value"));
    }
    Ok(())
}

fn field_varint(source: &[u8], field: FieldSpan) -> Result<u64, DecodeError> {
    field_wire(field, 0)?;
    let mut offset = field.value_start;
    Ok(read_varint(source, &mut offset)?.value)
}

fn field_bytes(source: &[u8], field: FieldSpan) -> Result<&[u8], DecodeError> {
    field_wire(field, 2)?;
    Ok(&source[field.value_start..field.value_end])
}

fn field_f64(source: &[u8], field: FieldSpan) -> Result<(), DecodeError> {
    field_wire(field, 1)?;
    let _ = &source[field.value_start..field.value_end];
    Ok(())
}

fn field_string(source: &[u8], field: FieldSpan) -> Result<(), DecodeError> {
    let bytes = field_bytes(source, field)?;
    str::from_utf8(bytes).map_err(|_| buffa::DecodeError::InvalidUtf8)?;
    Ok(())
}

fn singular(
    source: &[u8],
    fields: &[FieldSpan],
    number: u32,
    expected_wire: u8,
    name: &'static str,
) -> Result<Option<FieldSpan>, DecodeError> {
    let mut found = None;
    for &field in fields.iter().filter(|field| field.number == number) {
        field_wire(field, expected_wire)?;
        if found.replace(field).is_some() {
            return Err(DecodeError::duplicate(name));
        }
    }
    let _ = source;
    Ok(found)
}

#[allow(
    clippy::needless_lifetimes,
    reason = "The provenance guard associates this strict UUID preflight with the native archive source."
)]
fn parse_uuid<'source>(
    source: &'source [u8],
    options: DecodeOptions,
    budget: &mut Budget,
    depth: u32,
) -> Result<UuidSnapshot, DecodeError> {
    let fields = scan_fields(source, options, budget, depth)?;
    let lower = singular(source, &fields, UUID_LOWER_FIELD, 0, "TSP.UUID.lower")?
        .ok_or_else(|| DecodeError::missing("TSP.UUID.lower"))?;
    let upper = singular(source, &fields, UUID_UPPER_FIELD, 0, "TSP.UUID.upper")?
        .ok_or_else(|| DecodeError::missing("TSP.UUID.upper"))?;
    Ok(UuidSnapshot::new(
        field_varint(source, lower)?,
        field_varint(source, upper)?,
    ))
}

fn cross_check_uuid(
    source: &[u8],
    expected: UuidSnapshot,
    options: DecodeOptions,
) -> Result<(), DecodeError> {
    let view: projection::UuidLazyView<'_> = options.buffa().decode_lazy_view(source)?;
    if !view.has_lower()
        || !view.has_upper()
        || view.lower != expected.lower
        || view.upper != expected.upper
    {
        return Err(DecodeError::projection());
    }
    Ok(())
}

fn parse_reference(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
    depth: u32,
) -> Result<ReferenceSnapshot, DecodeError> {
    let fields = scan_fields(source, options, budget, depth)?;
    let identifier = singular(
        source,
        &fields,
        REFERENCE_IDENTIFIER_FIELD,
        0,
        "TSP.Reference.identifier",
    )?
    .ok_or_else(|| DecodeError::missing("TSP.Reference.identifier"))?;
    let identifier = NonZeroU64::new(field_varint(source, identifier)?)
        .ok_or_else(|| DecodeError::invalid("reference identifier is zero"))?;
    let deprecated_type = singular(
        source,
        &fields,
        REFERENCE_DEPRECATED_TYPE_FIELD,
        0,
        "TSP.Reference.deprecated_type",
    )?
    .map(|field| decode_int32_checked(field_varint(source, field)?))
    .transpose()?;
    let deprecated_is_external = singular(
        source,
        &fields,
        REFERENCE_DEPRECATED_EXTERNAL_FIELD,
        0,
        "TSP.Reference.deprecated_is_external",
    )?
    .map(|field| require_bool(field_varint(source, field)?))
    .transpose()?;
    Ok(ReferenceSnapshot {
        identifier,
        deprecated_type,
        deprecated_is_external,
    })
}

fn parse_formula_owner_dependencies(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
    depth: u32,
) -> Result<FormulaOwnerDependenciesSnapshot, DecodeError> {
    let fields = scan_fields(source, options, budget, depth)?;
    let has_dependencies = fields
        .iter()
        .any(|field| matches!(field.number, 4..=10 | 13..=16));
    let uid_field = singular(
        source,
        &fields,
        FORMULA_OWNER_UID_FIELD,
        2,
        "TSCE.FormulaOwnerDependenciesArchive.formula_owner_uid",
    )?
    .ok_or_else(|| {
        DecodeError::missing("TSCE.FormulaOwnerDependenciesArchive.formula_owner_uid")
    })?;
    let formula_owner_uid_bytes = field_bytes(source, uid_field)?;
    let formula_owner_uid = parse_uuid(
        formula_owner_uid_bytes,
        options,
        budget,
        child_depth(depth)?,
    )?;
    budget.retain(formula_owner_uid_bytes.len())?;
    let internal_field = singular(
        source,
        &fields,
        FORMULA_OWNER_INTERNAL_ID_FIELD,
        0,
        "TSCE.FormulaOwnerDependenciesArchive.internal_formula_owner_id",
    )?
    .ok_or_else(|| {
        DecodeError::missing("TSCE.FormulaOwnerDependenciesArchive.internal_formula_owner_id")
    })?;
    let internal_formula_owner_id = u32::try_from(field_varint(source, internal_field)?)
        .map_err(|_| DecodeError::invalid("formula-owner internal ID is out of range"))?;
    let owner_kind = singular(
        source,
        &fields,
        FORMULA_OWNER_KIND_FIELD,
        0,
        "TSCE.FormulaOwnerDependenciesArchive.owner_kind",
    )?
    .map(|field| {
        u32::try_from(field_varint(source, field)?)
            .map_err(|_| DecodeError::invalid("formula-owner kind is out of range"))
    })
    .transpose()?;
    let formula_owner = singular(
        source,
        &fields,
        FORMULA_OWNER_REFERENCE_FIELD,
        2,
        "TSCE.FormulaOwnerDependenciesArchive.formula_owner",
    )?
    .map(|field| field_bytes(source, field))
    .transpose()?
    .map(|bytes| parse_reference(bytes, options, budget, child_depth(depth)?))
    .transpose()?;
    let base_owner_uid = singular(
        source,
        &fields,
        FORMULA_OWNER_BASE_UID_FIELD,
        2,
        "TSCE.FormulaOwnerDependenciesArchive.base_owner_uid",
    )?
    .map(|field| field_bytes(source, field))
    .transpose()?
    .map(|bytes| parse_uuid(bytes, options, budget, child_depth(depth)?))
    .transpose()?;
    Ok(FormulaOwnerDependenciesSnapshot {
        formula_owner_uid,
        internal_formula_owner_id,
        owner_kind,
        formula_owner,
        base_owner_uid,
        has_dependencies,
    })
}

fn cross_check_formula_owner_dependencies(
    source: &[u8],
    expected: &FormulaOwnerDependenciesSnapshot,
    options: DecodeOptions,
) -> Result<(), DecodeError> {
    let view: projection::FormulaOwnerDependenciesArchiveLazyView<'_> =
        options.buffa().decode_lazy_view(source)?;
    let uid = view.formula_owner_uid;
    cross_check_uuid(uid, expected.formula_owner_uid, options)?;
    if view.internal_formula_owner_id != expected.internal_formula_owner_id
        || view.owner_kind != expected.owner_kind
    {
        return Err(DecodeError::projection());
    }
    match (view.formula_owner, expected.formula_owner) {
        (Some(bytes), Some(reference)) => cross_check_reference(bytes, reference, options)?,
        (None, None) => {},
        _ => return Err(DecodeError::projection()),
    }
    match (view.base_owner_uid, expected.base_owner_uid) {
        (Some(bytes), Some(uuid)) => cross_check_uuid(bytes, uuid, options)?,
        (None, None) => {},
        _ => return Err(DecodeError::projection()),
    }
    Ok(())
}

fn parse_hidden_state_formula_owner(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
    depth: u32,
) -> Result<HiddenStateFormulaOwnerSnapshot, DecodeError> {
    let fields = scan_fields(source, options, budget, depth)?;
    let owner_id = singular(
        source,
        &fields,
        HIDDEN_FORMULA_OWNER_ID_FIELD,
        2,
        "TST.HiddenStateFormulaOwnerArchive.owner_id",
    )?
    .map(|field| field_bytes(source, field))
    .transpose()?
    .map(|bytes| parse_cfuuid(bytes, options, budget, child_depth(depth)?))
    .transpose()?;
    let needs_to_update_filter_set_for_import = singular(
        source,
        &fields,
        HIDDEN_FORMULA_OWNER_NEEDS_UPDATE_FIELD,
        0,
        "TST.HiddenStateFormulaOwnerArchive.needs_to_update_filter_set_for_import",
    )?
    .map(|field| require_bool(field_varint(source, field)?))
    .transpose()?;
    Ok(HiddenStateFormulaOwnerSnapshot {
        owner_id,
        needs_to_update_filter_set_for_import,
    })
}

fn cross_check_hidden_state_formula_owner(
    source: &[u8],
    expected: &HiddenStateFormulaOwnerSnapshot,
    options: DecodeOptions,
) -> Result<(), DecodeError> {
    let view: projection::HiddenStateFormulaOwnerArchiveLazyView<'_> =
        options.buffa().decode_lazy_view(source)?;
    match (view.owner_id, expected.owner_id.as_ref()) {
        (Some(bytes), Some(owner_id)) => cross_check_cfuuid(bytes, owner_id, options)?,
        (None, None) => {},
        _ => return Err(DecodeError::projection()),
    }
    if view.needs_to_update_filter_set_for_import != expected.needs_to_update_filter_set_for_import
    {
        return Err(DecodeError::projection());
    }
    Ok(())
}

fn parse_filter_set(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
    depth: u32,
) -> Result<FilterSetSnapshot, DecodeError> {
    let fields = scan_fields(source, options, budget, depth)?;
    let filter_type = singular(
        source,
        &fields,
        FILTER_SET_TYPE_FIELD,
        0,
        "TST.FilterSetArchive.type",
    )?
    .map(|field| decode_int32_checked(field_varint(source, field)?))
    .transpose()?;
    if let Some(filter_type) = filter_type
        && !matches!(filter_type, 0 | 1)
    {
        return Err(DecodeError::invalid("filter-set type is out of range"));
    }
    let is_enabled = singular(
        source,
        &fields,
        FILTER_SET_ENABLED_FIELD,
        0,
        "TST.FilterSetArchive.is_enabled",
    )?
    .map(|field| require_bool(field_varint(source, field)?))
    .transpose()?;
    let needs_formula_rewrite_for_import = singular(
        source,
        &fields,
        FILTER_SET_NEEDS_FORMULA_REWRITE_FIELD,
        0,
        "TST.FilterSetArchive.needs_formula_rewrite_for_import",
    )?
    .map(|field| require_bool(field_varint(source, field)?))
    .transpose()?;
    // `filter_offsets` is declared as an unpacked proto2 repeated scalar, but
    // protobuf writers are allowed to use the packed representation when the
    // field is decoded by a compatible implementation.  Accept both forms
    // while keeping every packed element canonical and bounded before the
    // result vector is allocated.
    let count = fields.iter().try_fold(0usize, |count, field| {
        if field.number != FILTER_SET_OFFSETS_FIELD {
            return Ok(count);
        }
        count
            .checked_add(filter_offset_count(source, *field)?)
            .ok_or_else(DecodeError::projection)
    })?;
    let mut filter_offsets = Vec::new();
    reserve_states(&mut filter_offsets, count, options, budget)?;
    for field in fields
        .iter()
        .filter(|field| field.number == FILTER_SET_OFFSETS_FIELD)
    {
        match field.wire {
            0 => {
                filter_offsets.push(
                    u32::try_from(field_varint(source, *field)?)
                        .map_err(|_| DecodeError::invalid("filter offset is out of range"))?,
                );
                budget.state()?;
            },
            2 => {
                let packed = field_bytes(source, *field)?;
                let mut offset = 0;
                while offset < packed.len() {
                    let value = read_varint(packed, &mut offset)?;
                    if !value.canonical {
                        return Err(DecodeError::noncanonical("packed filter offset varint"));
                    }
                    filter_offsets.push(
                        u32::try_from(value.value)
                            .map_err(|_| DecodeError::invalid("filter offset is out of range"))?,
                    );
                    budget.state()?;
                }
            },
            _ => {
                // Keep the selected-field error typed like every other
                // singular scalar field. `filter_offset_count` has already
                // rejected this wire form during the bounded preflight.
                field_wire(*field, 0)?;
                unreachable!("validated filter-offset wire type");
            },
        }
    }
    Ok(FilterSetSnapshot {
        filter_type,
        is_enabled,
        needs_formula_rewrite_for_import,
        filter_offsets,
    })
}

fn filter_offset_count(source: &[u8], field: FieldSpan) -> Result<usize, DecodeError> {
    match field.wire {
        0 => {
            let value = field_varint(source, field)?;
            u32::try_from(value)
                .map(|_value| 1)
                .map_err(|_| DecodeError::invalid("filter offset is out of range"))
        },
        2 => {
            let packed = field_bytes(source, field)?;
            let mut offset = 0;
            let mut count = 0usize;
            while offset < packed.len() {
                let value = read_varint(packed, &mut offset)?;
                if !value.canonical {
                    return Err(DecodeError::noncanonical("packed filter offset varint"));
                }
                let _ = u32::try_from(value.value)
                    .map_err(|_| DecodeError::invalid("filter offset is out of range"))?;
                count = count.checked_add(1).ok_or_else(DecodeError::projection)?;
            }
            Ok(count)
        },
        _ => {
            field_wire(field, 0)?;
            unreachable!("validated filter-offset wire type");
        },
    }
}

fn cross_check_filter_set(
    source: &[u8],
    expected: &FilterSetSnapshot,
    options: DecodeOptions,
) -> Result<(), DecodeError> {
    let view: projection::FilterSetArchiveLazyView<'_> =
        options.buffa().decode_lazy_view(source)?;
    if view.r#type != expected.filter_type
        || view.is_enabled != expected.is_enabled
        || view.needs_formula_rewrite_for_import != expected.needs_formula_rewrite_for_import
    {
        return Err(DecodeError::projection());
    }
    Ok(())
}

fn parse_cfuuid(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
    depth: u32,
) -> Result<CfuuidSnapshot, DecodeError> {
    let fields = scan_fields(source, options, budget, depth)?;
    let uuid_bytes = singular(source, &fields, 1, 2, "TSP.CFUUIDArchive.uuid_bytes")?
        .map(|field| field_bytes(source, field))
        .transpose()?
        .map(|bytes| {
            if bytes.len() > MAX_CFUUID_BYTES {
                return Err(DecodeError::limit(DecodeLimit::RetainedBytes {
                    observed: bytes.len(),
                    maximum: MAX_CFUUID_BYTES,
                }));
            }
            budget.retain(bytes.len())?;
            budget.allocation(bytes.len())?;
            let mut owned = Vec::new();
            owned
                .try_reserve_exact(bytes.len())
                .map_err(|_| DecodeError::allocation(bytes.len()))?;
            owned.extend_from_slice(bytes);
            Ok::<Vec<u8>, DecodeError>(owned)
        })
        .transpose()?;
    let mut words = [None; 4];
    for (index, number) in (2..=5).enumerate() {
        words[index] = singular(source, &fields, number, 0, "TSP.CFUUIDArchive.word")?
            .map(|field| {
                u32::try_from(field_varint(source, field)?)
                    .map_err(|_| DecodeError::invalid("CFUUID word is out of range"))
            })
            .transpose()?;
    }
    if let Some(bytes) = uuid_bytes.as_deref() {
        let any_word = words.iter().any(Option::is_some);
        if any_word {
            let [Some(w0), Some(w1), Some(w2), Some(w3)] = words else {
                return Err(DecodeError::invalid(
                    "CFUUID bytes and partial word representation conflict",
                ));
            };
            if bytes.len() != MAX_CFUUID_BYTES {
                return Err(DecodeError::invalid(
                    "CFUUID bytes and word representation conflict",
                ));
            }
            let encoded =
                u128::from_be_bytes(bytes.try_into().map_err(|_| DecodeError::projection())?);
            let words_value = (u128::from(w2) << 96)
                | (u128::from(w3) << 64)
                | (u128::from(w1) << 32)
                | u128::from(w0);
            if encoded != words_value {
                return Err(DecodeError::invalid(
                    "CFUUID bytes and word representation conflict",
                ));
            }
        }
    }
    Ok(CfuuidSnapshot::new(uuid_bytes, words))
}

fn cross_check_cfuuid(
    source: &[u8],
    expected: &CfuuidSnapshot,
    options: DecodeOptions,
) -> Result<(), DecodeError> {
    let view: projection::CFUUIDArchiveLazyView<'_> = options.buffa().decode_lazy_view(source)?;
    if view.uuid_bytes != expected.uuid_bytes()
        || view.uuid_w0 != expected.words()[0]
        || view.uuid_w1 != expected.words()[1]
        || view.uuid_w2 != expected.words()[2]
        || view.uuid_w3 != expected.words()[3]
    {
        return Err(DecodeError::projection());
    }
    Ok(())
}

fn cross_check_reference(
    source: &[u8],
    expected: ReferenceSnapshot,
    options: DecodeOptions,
) -> Result<(), DecodeError> {
    let view: projection::ReferenceLazyView<'_> = options.buffa().decode_lazy_view(source)?;
    if !view.has_identifier()
        || view.identifier != expected.identifier.get()
        || view.deprecated_type != expected.deprecated_type
        || view.deprecated_is_external != expected.deprecated_is_external
    {
        return Err(DecodeError::projection());
    }
    Ok(())
}

fn parse_table_info<'source>(
    source: &'source [u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<ParsedTableInfo<'source>, DecodeError> {
    let fields = scan_fields(source, options, budget, 1)?;
    let super_field = singular(
        source,
        &fields,
        TABLE_INFO_SUPER_FIELD,
        2,
        "TST.TableInfoArchive.super",
    )?
    .ok_or_else(|| DecodeError::missing("TST.TableInfoArchive.super"))?;
    let _ = field_bytes(source, super_field)?;
    let model_field = singular(
        source,
        &fields,
        TABLE_INFO_MODEL_FIELD,
        2,
        "TST.TableInfoArchive.tableModel",
    )?
    .ok_or_else(|| DecodeError::missing("TST.TableInfoArchive.tableModel"))?;
    let model_bytes = field_bytes(source, model_field)?;
    budget.retain(model_bytes.len())?;
    let table_model = parse_reference(model_bytes, options, budget, 2)?;
    let view_column_row_uids = singular(
        source,
        &fields,
        TABLE_INFO_VIEW_UIDS_FIELD,
        2,
        "TST.TableInfoArchive.view_column_row_uids",
    )?
    .map(|field| field_bytes(source, field))
    .transpose()?
    .map(|bytes| parse_reference(bytes, options, budget, 2))
    .transpose()?;
    let hidden_states_uuid = singular(
        source,
        &fields,
        TABLE_INFO_HIDDEN_STATES_UUID_FIELD,
        2,
        "TST.TableInfoArchive.hidden_states_uuid",
    )?
    .map(|field| field_bytes(source, field))
    .transpose()?
    .map(|bytes| parse_uuid(bytes, options, budget, 2))
    .transpose()?;
    Ok(ParsedTableInfo {
        snapshot: TableInfoSnapshot {
            table_model,
            view_column_row_uids,
            hidden_states_uuid,
        },
        _source: core::marker::PhantomData,
    })
}

fn cross_check_table_info(
    source: &[u8],
    parsed: &ParsedTableInfo<'_>,
    options: DecodeOptions,
) -> Result<(), DecodeError> {
    let view: projection::TableInfoArchiveLazyView<'_> =
        options.buffa().decode_lazy_view(source)?;
    let model = view.table_model.ok_or_else(DecodeError::projection)?;
    cross_check_reference(model, parsed.snapshot.table_model, options)?;
    if let Some(reference) = parsed.snapshot.view_column_row_uids {
        let bytes = view
            .view_column_row_uids
            .ok_or_else(DecodeError::projection)?;
        cross_check_reference(bytes, reference, options)?;
    }
    if let Some(uuid) = parsed.snapshot.hidden_states_uuid {
        let bytes = view
            .hidden_states_uuid
            .ok_or_else(DecodeError::projection)?;
        cross_check_uuid(bytes, uuid, options)?;
    }
    Ok(())
}

fn parse_table_model<'source>(
    source: &'source [u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<ParsedTableModel<'source>, DecodeError> {
    let fields = scan_fields(source, options, budget, 1)?;
    validate_model_required(source, &fields)?;
    let rows = required_u32(
        source,
        &fields,
        TABLE_MODEL_ROWS_FIELD,
        "TST.TableModelArchive.number_of_rows",
    )?;
    let columns = required_u32(
        source,
        &fields,
        TABLE_MODEL_COLUMNS_FIELD,
        "TST.TableModelArchive.number_of_columns",
    )?;
    let hidden_rows = optional_u32(
        source,
        &fields,
        TABLE_MODEL_HIDDEN_ROWS_FIELD,
        "TST.TableModelArchive.number_of_hidden_rows",
    )?;
    let hidden_columns = optional_u32(
        source,
        &fields,
        TABLE_MODEL_HIDDEN_COLUMNS_FIELD,
        "TST.TableModelArchive.number_of_hidden_columns",
    )?;
    let filtered_rows = optional_u32(
        source,
        &fields,
        TABLE_MODEL_FILTERED_ROWS_FIELD,
        "TST.TableModelArchive.number_of_filtered_rows",
    )?;
    let user_hidden_rows = optional_u32(
        source,
        &fields,
        TABLE_MODEL_USER_HIDDEN_ROWS_FIELD,
        "TST.TableModelArchive.number_of_user_hidden_rows",
    )?;
    let user_hidden_columns = optional_u32(
        source,
        &fields,
        TABLE_MODEL_USER_HIDDEN_COLUMNS_FIELD,
        "TST.TableModelArchive.number_of_user_hidden_columns",
    )?;
    let column_formula = optional_reference(
        source,
        &fields,
        TABLE_MODEL_COLUMN_FORMULA_OWNER_FIELD,
        options,
        budget,
    )?;
    let row_formula = optional_reference(
        source,
        &fields,
        TABLE_MODEL_ROW_FORMULA_OWNER_FIELD,
        options,
        budget,
    )?;
    let base_column_row_uids = optional_reference(
        source,
        &fields,
        TABLE_MODEL_BASE_COLUMN_ROW_UIDS_FIELD,
        options,
        budget,
    )?;
    let hidden_owner = singular(
        source,
        &fields,
        TABLE_MODEL_HIDDEN_STATES_OWNER_FIELD,
        2,
        "TST.TableModelArchive.hidden_states_owner",
    )?
    .map(|field| field_bytes(source, field))
    .transpose()?
    .map(|bytes| parse_owner(bytes, options, budget, 2))
    .transpose()?;
    let snapshot = TableModelSnapshot {
        number_of_rows: rows,
        number_of_columns: columns,
        number_of_filtered_rows: filtered_rows,
        number_of_hidden_rows: hidden_rows,
        number_of_hidden_columns: hidden_columns,
        number_of_user_hidden_rows: user_hidden_rows,
        number_of_user_hidden_columns: user_hidden_columns,
        hidden_state_formula_owner_for_columns: column_formula,
        hidden_state_formula_owner_for_rows: row_formula,
        base_column_row_uids,
        hidden_states_owner: hidden_owner.as_ref().map(|parsed| parsed.snapshot.clone()),
    };
    if let Some(owner) = snapshot.hidden_states_owner.as_ref() {
        charge_owner_clone(owner, budget)?;
    }
    Ok(ParsedTableModel {
        snapshot,
        _source: core::marker::PhantomData,
    })
}

fn cross_check_table_model(
    source: &[u8],
    parsed: &ParsedTableModel<'_>,
    options: DecodeOptions,
) -> Result<(), DecodeError> {
    let view: projection::TableModelArchiveLazyView<'_> =
        options.buffa().decode_lazy_view(source)?;
    if view.number_of_rows != Some(parsed.snapshot.number_of_rows)
        || view.number_of_columns != Some(parsed.snapshot.number_of_columns)
        || view.number_of_filtered_rows != parsed.snapshot.number_of_filtered_rows
        || view.number_of_hidden_rows != parsed.snapshot.number_of_hidden_rows
        || view.number_of_hidden_columns != parsed.snapshot.number_of_hidden_columns
        || view.number_of_user_hidden_rows != parsed.snapshot.number_of_user_hidden_rows
        || view.number_of_user_hidden_columns != parsed.snapshot.number_of_user_hidden_columns
    {
        return Err(DecodeError::projection());
    }
    for (bytes, reference) in [
        (
            view.hidden_state_formula_owner_for_columns,
            parsed.snapshot.hidden_state_formula_owner_for_columns,
        ),
        (
            view.hidden_state_formula_owner_for_rows,
            parsed.snapshot.hidden_state_formula_owner_for_rows,
        ),
        (
            view.base_column_row_uids,
            parsed.snapshot.base_column_row_uids,
        ),
    ] {
        match (bytes, reference) {
            (Some(bytes), Some(reference)) => cross_check_reference(bytes, reference, options)?,
            (None, None) => {},
            _ => return Err(DecodeError::projection()),
        }
    }
    if let Some(owner) = &parsed.snapshot.hidden_states_owner {
        let bytes = view
            .hidden_states_owner
            .ok_or_else(DecodeError::projection)?;
        let owner_view: projection::HiddenStatesOwnerArchiveLazyView<'_> =
            options.buffa().decode_lazy_view(bytes)?;
        let uid = owner_view.owner_uid;
        cross_check_uuid(uid, owner.owner_uid, options)?;
    }
    Ok(())
}

fn validate_model_required(source: &[u8], fields: &[FieldSpan]) -> Result<(), DecodeError> {
    for (number, wire, name) in [
        (
            TABLE_MODEL_TABLE_ID_FIELD,
            2,
            "TST.TableModelArchive.table_id",
        ),
        (
            TABLE_MODEL_TABLE_STYLE_FIELD,
            2,
            "TST.TableModelArchive.table_style",
        ),
        (
            TABLE_MODEL_BASE_DATA_STORE_FIELD,
            2,
            "TST.TableModelArchive.base_data_store",
        ),
        (
            TABLE_MODEL_ROWS_FIELD,
            0,
            "TST.TableModelArchive.number_of_rows",
        ),
        (
            TABLE_MODEL_COLUMNS_FIELD,
            0,
            "TST.TableModelArchive.number_of_columns",
        ),
        (
            TABLE_MODEL_TABLE_NAME_FIELD,
            2,
            "TST.TableModelArchive.table_name",
        ),
        (
            TABLE_MODEL_DEFAULT_ROW_HEIGHT_FIELD,
            1,
            "TST.TableModelArchive.default_row_height",
        ),
        (
            TABLE_MODEL_DEFAULT_COLUMN_WIDTH_FIELD,
            1,
            "TST.TableModelArchive.default_column_width",
        ),
        (
            TABLE_MODEL_BODY_CELL_STYLE_FIELD,
            2,
            "TST.TableModelArchive.body_cell_style",
        ),
        (
            TABLE_MODEL_HEADER_ROW_STYLE_FIELD,
            2,
            "TST.TableModelArchive.header_row_style",
        ),
        (
            TABLE_MODEL_HEADER_COLUMN_STYLE_FIELD,
            2,
            "TST.TableModelArchive.header_column_style",
        ),
        (
            TABLE_MODEL_FOOTER_ROW_STYLE_FIELD,
            2,
            "TST.TableModelArchive.footer_row_style",
        ),
        (
            TABLE_MODEL_BODY_TEXT_STYLE_FIELD,
            2,
            "TST.TableModelArchive.body_text_style",
        ),
        (
            TABLE_MODEL_HEADER_ROW_TEXT_STYLE_FIELD,
            2,
            "TST.TableModelArchive.header_row_text_style",
        ),
        (
            TABLE_MODEL_HEADER_COLUMN_TEXT_STYLE_FIELD,
            2,
            "TST.TableModelArchive.header_column_text_style",
        ),
        (
            TABLE_MODEL_FOOTER_ROW_TEXT_STYLE_FIELD,
            2,
            "TST.TableModelArchive.footer_row_text_style",
        ),
    ] {
        let field = singular(source, fields, number, wire, name)?
            .ok_or_else(|| DecodeError::missing(name))?;
        match wire {
            0 => {
                let _ = field_varint(source, field)?;
            },
            1 => field_f64(source, field)?,
            2 => {
                let _ = field_bytes(source, field)?;
                if number == TABLE_MODEL_TABLE_ID_FIELD || number == TABLE_MODEL_TABLE_NAME_FIELD {
                    field_string(source, field)?;
                }
            },
            _ => unreachable!("fixed schema wire kind"),
        }
    }
    Ok(())
}

fn required_u32(
    source: &[u8],
    fields: &[FieldSpan],
    number: u32,
    name: &'static str,
) -> Result<u32, DecodeError> {
    let field =
        singular(source, fields, number, 0, name)?.ok_or_else(|| DecodeError::missing(name))?;
    let value = field_varint(source, field)?;
    u32::try_from(value).map_err(|_| DecodeError::invalid("uint32 scalar is out of range"))
}

fn optional_u32(
    source: &[u8],
    fields: &[FieldSpan],
    number: u32,
    name: &'static str,
) -> Result<Option<u32>, DecodeError> {
    singular(source, fields, number, 0, name)?
        .map(|field| {
            let value = field_varint(source, field)?;
            u32::try_from(value).map_err(|_| DecodeError::invalid("uint32 scalar is out of range"))
        })
        .transpose()
}

fn optional_reference(
    source: &[u8],
    fields: &[FieldSpan],
    number: u32,
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<Option<ReferenceSnapshot>, DecodeError> {
    singular(
        source,
        fields,
        number,
        2,
        match number {
            TABLE_MODEL_COLUMN_FORMULA_OWNER_FIELD => {
                "TST.TableModelArchive.hidden_state_formula_owner_for_columns"
            },
            TABLE_MODEL_ROW_FORMULA_OWNER_FIELD => {
                "TST.TableModelArchive.hidden_state_formula_owner_for_rows"
            },
            _ => "TST.TableModelArchive.reference",
        },
    )?
    .map(|field| field_bytes(source, field))
    .transpose()?
    .map(|bytes| {
        let reference = parse_reference(bytes, options, budget, 2)?;
        budget.retain(bytes.len())?;
        Ok(reference)
    })
    .transpose()
}

fn parse_owner<'source>(
    source: &'source [u8],
    options: DecodeOptions,
    budget: &mut Budget,
    depth: u32,
) -> Result<ParsedOwner<'source>, DecodeError> {
    let fields = scan_fields(source, options, budget, depth)?;
    let owner_uid_field = singular(
        source,
        &fields,
        OWNER_UID_FIELD,
        2,
        "TST.HiddenStatesOwnerArchive.owner_uid",
    )?
    .ok_or_else(|| DecodeError::missing("TST.HiddenStatesOwnerArchive.owner_uid"))?;
    let owner_uid_bytes = field_bytes(source, owner_uid_field)?;
    let owner_uid = parse_uuid(owner_uid_bytes, options, budget, child_depth(depth)?)?;
    budget.retain(owner_uid_bytes.len())?;
    cross_check_uuid(owner_uid_bytes, owner_uid, options)?;
    let mut states = Vec::new();
    let repeated = fields
        .iter()
        .filter(|field| field.number == OWNER_STATES_FIELD)
        .count();
    reserve_states(&mut states, repeated, options, budget)?;
    for field in fields
        .iter()
        .filter(|field| field.number == OWNER_STATES_FIELD)
    {
        let bytes = field_bytes(source, *field)?;
        budget.state()?;
        let parsed = parse_hidden_states(bytes, options, budget, child_depth(depth)?)?;
        states.push(parsed);
    }
    reject_duplicate_state_uids(&states, budget)?;
    let snapshot_bytes = states
        .len()
        .checked_mul(size_of::<HiddenStatesSnapshot>())
        .ok_or_else(DecodeError::projection)?;
    budget.allocation(snapshot_bytes)?;
    let mut snapshots = Vec::new();
    snapshots
        .try_reserve_exact(states.len())
        .map_err(|_| DecodeError::allocation(snapshot_bytes))?;
    for state in &states {
        charge_hidden_states_clone(&state.snapshot, budget)?;
        snapshots.push(state.snapshot.clone());
    }
    Ok(ParsedOwner {
        snapshot: HiddenStatesOwnerSnapshot {
            owner_uid,
            hidden_states: snapshots,
        },
        _source: core::marker::PhantomData,
    })
}

fn parse_hidden_states<'source>(
    source: &'source [u8],
    options: DecodeOptions,
    budget: &mut Budget,
    depth: u32,
) -> Result<ParsedHiddenStates<'source>, DecodeError> {
    let fields = scan_fields(source, options, budget, depth)?;
    let uid_field = singular(
        source,
        &fields,
        STATE_UID_FIELD,
        2,
        "TST.HiddenStatesArchive.hidden_states_uid",
    )?
    .ok_or_else(|| DecodeError::missing("TST.HiddenStatesArchive.hidden_states_uid"))?;
    let uid_bytes = field_bytes(source, uid_field)?;
    let uid = parse_uuid(uid_bytes, options, budget, child_depth(depth)?)?;
    let column_field = singular(
        source,
        &fields,
        STATE_COLUMN_EXTENT_FIELD,
        2,
        "TST.HiddenStatesArchive.column_hidden_state_extent",
    )?
    .ok_or_else(|| DecodeError::missing("TST.HiddenStatesArchive.column_hidden_state_extent"))?;
    let row_field = singular(
        source,
        &fields,
        STATE_ROW_EXTENT_FIELD,
        2,
        "TST.HiddenStatesArchive.row_hidden_state_extent",
    )?
    .ok_or_else(|| DecodeError::missing("TST.HiddenStatesArchive.row_hidden_state_extent"))?;
    let column_bytes = field_bytes(source, column_field)?;
    let row_bytes = field_bytes(source, row_field)?;
    let column = parse_extent(column_bytes, options, budget, child_depth(depth)?)?;
    let row = parse_extent(row_bytes, options, budget, child_depth(depth)?)?;
    if column.snapshot.direction() != AxisDirection::Column
        || row.snapshot.direction() != AxisDirection::Row
    {
        return Err(DecodeError::invalid(
            "hidden-state extent direction does not match axis",
        ));
    }
    charge_extent_clone(&column.snapshot, budget)?;
    charge_extent_clone(&row.snapshot, budget)?;
    let snapshot = HiddenStatesSnapshot::new(uid, column.snapshot.clone(), row.snapshot.clone());
    Ok(ParsedHiddenStates {
        snapshot,
        _source: core::marker::PhantomData,
    })
}

fn parse_extent<'source>(
    source: &'source [u8],
    options: DecodeOptions,
    budget: &mut Budget,
    depth: u32,
) -> Result<ParsedExtent<'source>, DecodeError> {
    let fields = scan_fields(source, options, budget, depth)?;
    let uid_field = singular(
        source,
        &fields,
        EXTENT_UID_FIELD,
        2,
        "TST.HiddenStateExtentArchive.hidden_state_extent_uid",
    )?
    .ok_or_else(|| DecodeError::missing("TST.HiddenStateExtentArchive.hidden_state_extent_uid"))?;
    let uid_bytes = field_bytes(source, uid_field)?;
    let uid = parse_uuid(uid_bytes, options, budget, child_depth(depth)?)?;
    let direction_field = singular(
        source,
        &fields,
        EXTENT_DIRECTION_FIELD,
        0,
        "TST.HiddenStateExtentArchive.row_or_column_direction",
    )?
    .ok_or_else(|| DecodeError::missing("TST.HiddenStateExtentArchive.row_or_column_direction"))?;
    let direction_value = i32::try_from(field_varint(source, direction_field)?)
        .map_err(|_| DecodeError::invalid("hidden-state extent direction is out of range"))?;
    let direction = AxisDirection::from_native(direction_value)?;
    let needs_to_update_filter_set_for_import = optional_bool(
        source,
        &fields,
        EXTENT_NEEDS_FILTER_UPDATE_FIELD,
        "TST.HiddenStateExtentArchive.needs_to_update_filter_set_for_import",
    )?;
    let filter_set = singular(
        source,
        &fields,
        EXTENT_FILTER_SET_FIELD,
        2,
        "TST.HiddenStateExtentArchive.filter_set",
    )?
    .map(|field| field_bytes(source, field))
    .transpose()?
    .map(|bytes| parse_reference(bytes, options, budget, child_depth(depth)?))
    .transpose()?;
    let repeated = fields
        .iter()
        .filter(|field| field.number == EXTENT_BASE_STATES_FIELD)
        .count();
    let mut states = Vec::new();
    reserve_states(&mut states, repeated, options, budget)?;
    for field in fields
        .iter()
        .filter(|field| field.number == EXTENT_BASE_STATES_FIELD)
    {
        let bytes = field_bytes(source, *field)?;
        budget.state()?;
        states.push(parse_row_state(
            bytes,
            options,
            budget,
            child_depth(depth)?,
        )?);
    }
    reject_duplicate_row_state_uids(&states, budget)?;
    let snapshot_bytes = states
        .len()
        .checked_mul(size_of::<RowOrColumnStateSnapshot>())
        .ok_or_else(DecodeError::projection)?;
    budget.allocation(snapshot_bytes)?;
    let mut snapshots = Vec::new();
    snapshots
        .try_reserve_exact(states.len())
        .map_err(|_| DecodeError::allocation(snapshot_bytes))?;
    snapshots.extend(states.iter().map(|state| state.snapshot));
    Ok(ParsedExtent {
        snapshot: HiddenStateExtentSnapshot {
            hidden_state_extent_uid: uid,
            direction,
            base_hidden_states: snapshots,
            needs_to_update_filter_set_for_import,
            filter_set,
        },
        _source: core::marker::PhantomData,
    })
}

fn parse_row_state<'source>(
    source: &'source [u8],
    options: DecodeOptions,
    budget: &mut Budget,
    depth: u32,
) -> Result<ParsedState<'source>, DecodeError> {
    let fields = scan_fields(source, options, budget, depth)?;
    let uid_field = singular(
        source,
        &fields,
        ROW_STATE_UID_FIELD,
        2,
        "TST.HiddenStateExtentArchive.RowOrColumnState.row_or_column_uid",
    )?
    .ok_or_else(|| {
        DecodeError::missing("TST.HiddenStateExtentArchive.RowOrColumnState.row_or_column_uid")
    })?;
    let uid_bytes = field_bytes(source, uid_field)?;
    let uid = parse_uuid(uid_bytes, options, budget, child_depth(depth)?)?;
    let user_hidden = optional_bool(
        source,
        &fields,
        ROW_STATE_USER_HIDDEN_FIELD,
        "TST.HiddenStateExtentArchive.RowOrColumnState.user_hidden",
    )?;
    let filtered = optional_bool(
        source,
        &fields,
        ROW_STATE_FILTERED_FIELD,
        "TST.HiddenStateExtentArchive.RowOrColumnState.filtered",
    )?;
    let pivot_hidden = optional_bool(
        source,
        &fields,
        ROW_STATE_PIVOT_HIDDEN_FIELD,
        "TST.HiddenStateExtentArchive.RowOrColumnState.pivot_hidden",
    )?;
    Ok(ParsedState {
        snapshot: RowOrColumnStateSnapshot {
            row_or_column_uid: uid,
            user_hidden,
            filtered,
            pivot_hidden,
        },
        _source: core::marker::PhantomData,
    })
}

fn cross_check_owner(
    source: &[u8],
    parsed: &ParsedOwner<'_>,
    options: DecodeOptions,
) -> Result<(), DecodeError> {
    let view: projection::HiddenStatesOwnerArchiveLazyView<'_> =
        options.buffa().decode_lazy_view(source)?;
    let uid = view.owner_uid;
    cross_check_uuid(uid, parsed.snapshot.owner_uid, options)
}

fn cross_check_extent(
    source: &[u8],
    parsed: &ParsedExtent<'_>,
    options: DecodeOptions,
) -> Result<(), DecodeError> {
    let view: projection::HiddenStateExtentArchiveLazyView<'_> =
        options.buffa().decode_lazy_view(source)?;
    let uid = view.hidden_state_extent_uid;
    cross_check_uuid(uid, parsed.snapshot.hidden_state_extent_uid, options)?;
    if view.row_or_column_direction != parsed.snapshot.direction.native_value() {
        return Err(DecodeError::projection());
    }
    if view.needs_to_update_filter_set_for_import
        != parsed.snapshot.needs_to_update_filter_set_for_import
    {
        return Err(DecodeError::projection());
    }
    match (view.filter_set, parsed.snapshot.filter_set) {
        (Some(bytes), Some(reference)) => cross_check_reference(bytes, reference, options),
        (None, None) => Ok(()),
        _ => Err(DecodeError::projection()),
    }
}

fn cross_check_row_state(
    source: &[u8],
    parsed: &ParsedState<'_>,
    options: DecodeOptions,
) -> Result<(), DecodeError> {
    let view: projection::RowOrColumnStateLazyView<'_> =
        options.buffa().decode_lazy_view(source)?;
    let uid = view.row_or_column_uid;
    cross_check_uuid(uid, parsed.snapshot.row_or_column_uid, options)?;
    if view.user_hidden != parsed.snapshot.user_hidden
        || view.filtered != parsed.snapshot.filtered
        || view.pivot_hidden != parsed.snapshot.pivot_hidden
    {
        return Err(DecodeError::projection());
    }
    Ok(())
}

fn require_bool(value: u64) -> Result<bool, DecodeError> {
    match value {
        0 => Ok(false),
        1 => Ok(true),
        _ => Err(DecodeError::noncanonical("bool scalar is not zero or one")),
    }
}

fn optional_bool(
    source: &[u8],
    fields: &[FieldSpan],
    number: u32,
    name: &'static str,
) -> Result<Option<bool>, DecodeError> {
    singular(source, fields, number, 0, name)?
        .map(|field| require_bool(field_varint(source, field)?))
        .transpose()
}

#[allow(
    clippy::cast_possible_truncation,
    clippy::cast_possible_wrap,
    reason = "Canonical int32 framing is validated before the cast."
)]
fn decode_int32_checked(value: u64) -> Result<i32, DecodeError> {
    // Protobuf int32 values are sign-extended to the ten-byte uint64 form.
    // Positive values use the short canonical form; negative values carry
    // ones in the high 32 bits and start at -2^31.
    if value <= i32::MAX as u64 || value >= 0xffff_ffff_8000_0000 {
        Ok(value as i32)
    } else {
        Err(DecodeError::noncanonical(
            "int32 scalar is not sign-extended",
        ))
    }
}

fn reserve_states<T>(
    states: &mut Vec<T>,
    count: usize,
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<(), DecodeError> {
    let maximum = options.max_states.min(MAX_CONSTRUCTED_STATES);
    if count > maximum {
        return Err(DecodeError::limit(DecodeLimit::States {
            observed: count,
            maximum,
        }));
    }
    if count == 0 {
        return Ok(());
    }
    let bytes = count
        .checked_mul(size_of::<T>())
        .ok_or_else(DecodeError::projection)?;
    budget.allocation(bytes)?;
    states
        .try_reserve_exact(count)
        .map_err(|_| DecodeError::allocation(bytes))
}

fn charge_extent_clone(
    extent: &HiddenStateExtentSnapshot,
    budget: &mut Budget,
) -> Result<(), DecodeError> {
    let bytes = extent
        .base_hidden_states
        .len()
        .checked_mul(size_of::<RowOrColumnStateSnapshot>())
        .ok_or_else(DecodeError::projection)?;
    if bytes != 0 {
        budget.allocation(bytes)?;
        budget.retain(bytes)?;
    }
    Ok(())
}

fn charge_hidden_states_clone(
    hidden_states: &HiddenStatesSnapshot,
    budget: &mut Budget,
) -> Result<(), DecodeError> {
    charge_extent_clone(hidden_states.column_hidden_state_extent(), budget)?;
    charge_extent_clone(hidden_states.row_hidden_state_extent(), budget)
}

fn charge_owner_clone(
    owner: &HiddenStatesOwnerSnapshot,
    budget: &mut Budget,
) -> Result<(), DecodeError> {
    for hidden_states in &owner.hidden_states {
        charge_hidden_states_clone(hidden_states, budget)?;
    }
    Ok(())
}

fn collect_states(
    states: impl IntoIterator<Item = RowOrColumnStateSnapshot>,
    maximum: usize,
) -> Result<Vec<RowOrColumnStateSnapshot>, DecodeError> {
    let mut result = Vec::new();
    let iterator = states.into_iter();
    if let Some(upper) = iterator.size_hint().1 {
        let reserve = upper.min(maximum);
        let bytes = reserve
            .checked_mul(size_of::<RowOrColumnStateSnapshot>())
            .ok_or_else(DecodeError::projection)?;
        result
            .try_reserve_exact(reserve)
            .map_err(|_| DecodeError::allocation(bytes))?;
    }
    for state in iterator {
        if result.len() >= maximum {
            let observed = result
                .len()
                .checked_add(1)
                .ok_or_else(DecodeError::projection)?;
            return Err(DecodeError::limit(DecodeLimit::States {
                observed,
                maximum,
            }));
        }
        result
            .try_reserve(1)
            .map_err(|_| DecodeError::allocation(size_of::<RowOrColumnStateSnapshot>()))?;
        result.push(state);
    }
    reject_duplicate_snapshot_uids(&result)?;
    Ok(result)
}

fn collect_hidden_states(
    states: impl IntoIterator<Item = HiddenStatesSnapshot>,
    maximum: usize,
) -> Result<Vec<HiddenStatesSnapshot>, DecodeError> {
    let mut result = Vec::new();
    let iterator = states.into_iter();
    if let Some(upper) = iterator.size_hint().1 {
        let reserve = upper.min(maximum);
        let bytes = reserve
            .checked_mul(size_of::<HiddenStatesSnapshot>())
            .ok_or_else(DecodeError::projection)?;
        result
            .try_reserve_exact(reserve)
            .map_err(|_| DecodeError::allocation(bytes))?;
    }
    for state in iterator {
        if result.len() >= maximum {
            let observed = result
                .len()
                .checked_add(1)
                .ok_or_else(DecodeError::projection)?;
            return Err(DecodeError::limit(DecodeLimit::States {
                observed,
                maximum,
            }));
        }
        result
            .try_reserve(1)
            .map_err(|_| DecodeError::allocation(size_of::<HiddenStatesSnapshot>()))?;
        result.push(state);
    }
    let bytes = result
        .len()
        .checked_mul(size_of::<UuidSnapshot>())
        .ok_or_else(DecodeError::projection)?;
    let mut ids = Vec::new();
    ids.try_reserve_exact(result.len())
        .map_err(|_| DecodeError::allocation(bytes))?;
    ids.extend(result.iter().map(|state| state.hidden_states_uid));
    ids.sort_unstable_by_key(|uid| uid_key(*uid));
    if ids.windows(2).any(|window| window[0] == window[1]) {
        return Err(DecodeError::invalid("duplicate hidden-state UUID"));
    }
    Ok(result)
}

fn collect_u32(
    values: impl IntoIterator<Item = u32>,
    maximum: usize,
) -> Result<Vec<u32>, DecodeError> {
    let mut result = Vec::new();
    let iterator = values.into_iter();
    if let Some(upper) = iterator.size_hint().1 {
        let reserve = upper.min(maximum);
        let bytes = reserve
            .checked_mul(size_of::<u32>())
            .ok_or_else(DecodeError::projection)?;
        result
            .try_reserve_exact(reserve)
            .map_err(|_| DecodeError::allocation(bytes))?;
    }
    for value in iterator {
        if result.len() >= maximum {
            let observed = result
                .len()
                .checked_add(1)
                .ok_or_else(DecodeError::projection)?;
            return Err(DecodeError::limit(DecodeLimit::States {
                observed,
                maximum,
            }));
        }
        result
            .try_reserve(1)
            .map_err(|_| DecodeError::allocation(size_of::<u32>()))?;
        result.push(value);
    }
    Ok(result)
}

fn reject_duplicate_snapshot_uids(states: &[RowOrColumnStateSnapshot]) -> Result<(), DecodeError> {
    let bytes = states
        .len()
        .checked_mul(size_of::<UuidSnapshot>())
        .ok_or_else(DecodeError::projection)?;
    let mut ids = Vec::new();
    ids.try_reserve_exact(states.len())
        .map_err(|_| DecodeError::allocation(bytes))?;
    ids.extend(states.iter().map(|state| state.row_or_column_uid));
    ids.sort_unstable_by_key(|uid| uid_key(*uid));
    if ids.windows(2).any(|window| window[0] == window[1]) {
        return Err(DecodeError::invalid("duplicate row/column UUID"));
    }
    Ok(())
}

fn reject_duplicate_row_state_uids(
    states: &[ParsedState<'_>],
    budget: &mut Budget,
) -> Result<(), DecodeError> {
    let bytes = states
        .len()
        .checked_mul(size_of::<UuidSnapshot>())
        .ok_or_else(DecodeError::projection)?;
    budget.allocation(bytes)?;
    let mut ids = Vec::new();
    ids.try_reserve_exact(states.len())
        .map_err(|_| DecodeError::allocation(bytes))?;
    ids.extend(states.iter().map(|state| state.snapshot.row_or_column_uid));
    ids.sort_unstable_by_key(|uid| uid_key(*uid));
    if ids.windows(2).any(|window| window[0] == window[1]) {
        return Err(DecodeError::invalid("duplicate row/column UUID"));
    }
    Ok(())
}

fn reject_duplicate_state_uids(
    states: &[ParsedHiddenStates<'_>],
    budget: &mut Budget,
) -> Result<(), DecodeError> {
    let bytes = states
        .len()
        .checked_mul(size_of::<UuidSnapshot>())
        .ok_or_else(DecodeError::projection)?;
    budget.allocation(bytes)?;
    let mut ids = Vec::new();
    ids.try_reserve_exact(states.len())
        .map_err(|_| DecodeError::allocation(bytes))?;
    ids.extend(states.iter().map(|state| state.snapshot.hidden_states_uid));
    ids.sort_unstable_by_key(|uid| uid_key(*uid));
    if ids.windows(2).any(|window| window[0] == window[1]) {
        return Err(DecodeError::invalid("duplicate hidden-state UUID"));
    }
    Ok(())
}

#[derive(Debug, Clone, Copy, Default)]
struct DesiredAccounting {
    states: usize,
    owned_bytes: usize,
    allocations: usize,
    match_bytes: usize,
    match_allocations: usize,
    match_work_bytes: usize,
}

fn account_vec(
    length: usize,
    element_size: usize,
    allocations: &mut usize,
    owned_bytes: &mut usize,
) -> Result<(), DecodeError> {
    if length == 0 {
        return Ok(());
    }
    let bytes = length
        .checked_mul(element_size)
        .ok_or_else(DecodeError::projection)?;
    *owned_bytes = owned_bytes
        .checked_add(bytes)
        .ok_or_else(DecodeError::projection)?;
    *allocations = allocations
        .checked_add(1)
        .ok_or_else(DecodeError::projection)?;
    Ok(())
}

fn account_match_index(
    length: usize,
    match_bytes: &mut usize,
    match_allocations: &mut usize,
    match_work_bytes: &mut usize,
) -> Result<(), DecodeError> {
    if length == 0 {
        return Ok(());
    }
    let per_entry = size_of::<UidIndexEntry>()
        .checked_add(size_of::<bool>())
        .ok_or_else(DecodeError::projection)?;
    let bytes = length
        .checked_mul(per_entry)
        .ok_or_else(DecodeError::projection)?;
    // Preparation measures once and execution emits once. Emitting nested
    // envelopes also measures their child collection before writing, so four
    // bounded passes cover every index/bitmap allocation without accepting an
    // unaccounted graph-width multiplier.
    *match_bytes = match_bytes
        .checked_add(bytes.checked_mul(4).ok_or_else(DecodeError::projection)?)
        .ok_or_else(DecodeError::projection)?;
    *match_allocations = match_allocations
        .checked_add(8)
        .ok_or_else(DecodeError::projection)?;
    *match_work_bytes = match_work_bytes
        .checked_add(bytes.checked_mul(8).ok_or_else(DecodeError::projection)?)
        .ok_or_else(DecodeError::projection)?;
    Ok(())
}

fn account_row_state(
    _state: &RowOrColumnStateSnapshot,
    accounting: &mut DesiredAccounting,
) -> Result<(), DecodeError> {
    accounting.states = accounting
        .states
        .checked_add(1)
        .ok_or_else(DecodeError::projection)?;
    Ok(())
}

fn account_extent(
    extent: &HiddenStateExtentSnapshot,
    accounting: &mut DesiredAccounting,
) -> Result<(), DecodeError> {
    account_vec(
        extent.base_hidden_states.len(),
        size_of::<RowOrColumnStateSnapshot>(),
        &mut accounting.allocations,
        &mut accounting.owned_bytes,
    )?;
    account_match_index(
        extent.base_hidden_states.len(),
        &mut accounting.match_bytes,
        &mut accounting.match_allocations,
        &mut accounting.match_work_bytes,
    )?;
    // The prepared plan also retains one metadata record for every extent,
    // including an extent with no row records. Account this independently of
    // the row-index bytes so owner width cannot hide an extent-vector
    // allocation.
    let metadata_bytes = size_of::<ExtentIndexEntry>();
    accounting.match_bytes = accounting
        .match_bytes
        .checked_add(
            metadata_bytes
                .checked_mul(4)
                .ok_or_else(DecodeError::projection)?,
        )
        .ok_or_else(DecodeError::projection)?;
    accounting.match_allocations = accounting
        .match_allocations
        .checked_add(1)
        .ok_or_else(DecodeError::projection)?;
    accounting.match_work_bytes = accounting
        .match_work_bytes
        .checked_add(
            metadata_bytes
                .checked_mul(4)
                .ok_or_else(DecodeError::projection)?,
        )
        .ok_or_else(DecodeError::projection)?;
    for state in &extent.base_hidden_states {
        account_row_state(state, accounting)?;
    }
    Ok(())
}

fn account_hidden_states(
    hidden_states: &HiddenStatesSnapshot,
    accounting: &mut DesiredAccounting,
) -> Result<(), DecodeError> {
    if hidden_states.column_hidden_state_extent.direction() != AxisDirection::Column
        || hidden_states.row_hidden_state_extent.direction() != AxisDirection::Row
    {
        return Err(DecodeError::invalid(
            "hidden-state extent direction does not match axis",
        ));
    }
    accounting.states = accounting
        .states
        .checked_add(1)
        .ok_or_else(DecodeError::projection)?;
    account_extent(hidden_states.column_hidden_state_extent(), accounting)?;
    account_extent(hidden_states.row_hidden_state_extent(), accounting)?;
    Ok(())
}

fn account_owner(
    owner: &HiddenStatesOwnerSnapshot,
    accounting: &mut DesiredAccounting,
) -> Result<(), DecodeError> {
    accounting.owned_bytes = accounting
        .owned_bytes
        .checked_add(size_of::<HiddenStatesOwnerSnapshot>())
        .ok_or_else(DecodeError::projection)?;
    account_vec(
        owner.hidden_states.len(),
        size_of::<HiddenStatesSnapshot>(),
        &mut accounting.allocations,
        &mut accounting.owned_bytes,
    )?;
    account_match_index(
        owner.hidden_states.len(),
        &mut accounting.match_bytes,
        &mut accounting.match_allocations,
        &mut accounting.match_work_bytes,
    )?;
    for hidden_states in &owner.hidden_states {
        account_hidden_states(hidden_states, accounting)?;
    }
    Ok(())
}

fn account_desired(desired: &RewriteDesired) -> Result<DesiredAccounting, DecodeError> {
    let mut accounting = DesiredAccounting::default();
    match desired {
        RewriteDesired::TableInfo(value) => {
            let _ = value;
        },
        RewriteDesired::TableModel(value) => {
            if let Some(owner) = value.hidden_states_owner.as_ref() {
                account_owner(owner, &mut accounting)?;
            }
        },
        RewriteDesired::HiddenStatesOwner(value) => account_owner(value, &mut accounting)?,
        RewriteDesired::HiddenStateExtent(value) => {
            account_extent(value, &mut accounting)?;
        },
        RewriteDesired::RowOrColumnState(value) => {
            let _ = value;
        },
    }
    Ok(accounting)
}

fn validate_desired(
    desired: &RewriteDesired,
    options: DecodeOptions,
) -> Result<DesiredAccounting, DecodeError> {
    let accounting = account_desired(desired)?;
    validate_desired_accounting(accounting, options)?;
    Ok(accounting)
}

fn validate_desired_accounting(
    accounting: DesiredAccounting,
    options: DecodeOptions,
) -> Result<(), DecodeError> {
    if accounting.states > MAX_CONSTRUCTED_STATES || accounting.states > options.max_states {
        return Err(DecodeError::limit(DecodeLimit::States {
            observed: accounting.states,
            maximum: options.max_states.min(MAX_CONSTRUCTED_STATES),
        }));
    }
    if accounting.owned_bytes > options.max_retained_bytes {
        return Err(DecodeError::limit(DecodeLimit::RetainedBytes {
            observed: accounting.owned_bytes,
            maximum: options.max_retained_bytes,
        }));
    }
    if accounting.match_bytes > options.max_scratch_bytes {
        return Err(DecodeError::limit(DecodeLimit::ScratchBytes {
            observed: accounting.match_bytes,
            maximum: options.max_scratch_bytes,
        }));
    }
    if accounting.match_work_bytes > options.max_work_bytes {
        return Err(DecodeError::limit(DecodeLimit::WorkBytes {
            observed: accounting.match_work_bytes,
            maximum: options.max_work_bytes,
        }));
    }
    // `build_rewrite_index_plan` reserves these collections immediately after
    // this validation. Check the conservative accounting before entering it so
    // a rejected rewrite cannot allocate its index storage first.
    let allocations = accounting
        .allocations
        .checked_add(accounting.match_allocations)
        .and_then(|value| value.checked_add(8))
        .ok_or_else(|| {
            DecodeError::limit(DecodeLimit::Allocations {
                observed: usize::MAX,
                maximum: options.max_allocations,
            })
        })?;
    if allocations > options.max_allocations {
        return Err(DecodeError::limit(DecodeLimit::Allocations {
            observed: allocations,
            maximum: options.max_allocations,
        }));
    }
    Ok(())
}

fn validate_table_info_desired(
    _desired: &TableInfoSnapshot,
    _options: DecodeOptions,
) -> Result<(), DecodeError> {
    Ok(())
}

fn validate_table_model_desired(
    desired: &TableModelSnapshot,
    options: DecodeOptions,
) -> Result<(), DecodeError> {
    let mut accounting = DesiredAccounting::default();
    if let Some(owner) = desired.hidden_states_owner.as_ref() {
        account_owner(owner, &mut accounting)?;
    }
    validate_desired_accounting(accounting, options)
}

fn validate_owner_desired(
    desired: &HiddenStatesOwnerSnapshot,
    options: DecodeOptions,
) -> Result<(), DecodeError> {
    let mut accounting = DesiredAccounting::default();
    account_owner(desired, &mut accounting)?;
    validate_desired_accounting(accounting, options)
}

fn validate_extent_desired(
    desired: &HiddenStateExtentSnapshot,
    options: DecodeOptions,
) -> Result<(), DecodeError> {
    let mut accounting = DesiredAccounting::default();
    account_extent(desired, &mut accounting)?;
    validate_desired_accounting(accounting, options)
}

fn validate_row_state_desired(options: DecodeOptions) -> Result<(), DecodeError> {
    if options.max_states == 0 {
        return Err(DecodeError::limit(DecodeLimit::States {
            observed: 1,
            maximum: options.max_states,
        }));
    }
    Ok(())
}

fn prepare_rewrite<'source>(
    source: &'source [u8],
    kind: RewriteKind,
    desired: RewriteDesired,
    options: DecodeOptions,
) -> Result<PreparedRewrite<'source>, DecodeError> {
    validate_input(source, options)?;
    let desired_accounting = validate_desired(&desired, options)?;
    let mut budget = Budget::new(options);
    match kind {
        RewriteKind::TableInfo => {
            let parsed = parse_table_info(source, options, &mut budget)?;
            cross_check_table_info(source, &parsed, options)?;
        },
        RewriteKind::TableModel => {
            let parsed = parse_table_model(source, options, &mut budget)?;
            cross_check_table_model(source, &parsed, options)?;
        },
        RewriteKind::HiddenStatesOwner => {
            let parsed = parse_owner(source, options, &mut budget, 1)?;
            cross_check_owner(source, &parsed, options)?;
        },
        RewriteKind::HiddenStateExtent => {
            let parsed = parse_extent(source, options, &mut budget, 1)?;
            cross_check_extent(source, &parsed, options)?;
        },
        RewriteKind::RowOrColumnState => {
            let parsed = parse_row_state(source, options, &mut budget, 1)?;
            cross_check_row_state(source, &parsed, options)?;
        },
    }
    let index_plan = build_rewrite_index_plan(&desired)?;
    let output_bytes = measure_rewrite(source, kind, &desired, &index_plan)?;
    if output_bytes > options.max_output_bytes {
        return Err(DecodeError::limit(DecodeLimit::OutputBytes {
            observed: output_bytes,
            maximum: options.max_output_bytes,
        }));
    }
    let fields = budget
        .fields
        .checked_mul(3)
        .and_then(|value| value.checked_add(64))
        .ok_or_else(DecodeError::projection)?;
    if fields > options.max_fields {
        return Err(DecodeError::limit(DecodeLimit::Fields {
            observed: fields,
            maximum: options.max_fields,
        }));
    }
    let work_bytes = budget
        .work_bytes
        .checked_add(
            source
                .len()
                .checked_add(output_bytes)
                .and_then(|value| value.checked_mul(4))
                .ok_or_else(DecodeError::projection)?,
        )
        .and_then(|value| value.checked_add(desired_accounting.match_work_bytes))
        .ok_or_else(DecodeError::projection)?;
    if work_bytes > options.max_work_bytes {
        return Err(DecodeError::limit(DecodeLimit::WorkBytes {
            observed: work_bytes,
            maximum: options.max_work_bytes,
        }));
    }
    let retained_bytes = source
        .len()
        .checked_add(output_bytes)
        .and_then(|value| value.checked_add(budget.retained_bytes))
        .and_then(|value| value.checked_add(desired_accounting.owned_bytes))
        .ok_or_else(DecodeError::projection)?;
    if retained_bytes > options.max_retained_bytes {
        return Err(DecodeError::limit(DecodeLimit::RetainedBytes {
            observed: retained_bytes,
            maximum: options.max_retained_bytes,
        }));
    }
    let scratch_bytes = budget
        .scratch_bytes
        .checked_add(output_bytes)
        .and_then(|value| value.checked_add(desired_accounting.match_bytes))
        .ok_or_else(DecodeError::projection)?;
    if scratch_bytes > options.max_scratch_bytes {
        return Err(DecodeError::limit(DecodeLimit::ScratchBytes {
            observed: scratch_bytes,
            maximum: options.max_scratch_bytes,
        }));
    }
    let allocations = budget
        .allocations
        .checked_add(desired_accounting.allocations)
        .and_then(|value| value.checked_add(desired_accounting.match_allocations))
        .and_then(|value| value.checked_add(8))
        .ok_or_else(DecodeError::projection)?;
    if allocations > options.max_allocations {
        return Err(DecodeError::limit(DecodeLimit::Allocations {
            observed: allocations,
            maximum: options.max_allocations,
        }));
    }
    let report = DecodeReport {
        input_bytes: source.len(),
        output_bytes,
        fields,
        work_bytes,
        max_depth: budget.max_depth.max(2),
        states: budget
            .states
            .checked_add(desired_accounting.states)
            .ok_or_else(DecodeError::projection)?,
        allocations,
        retained_bytes,
        scratch_bytes,
    };
    let requirements = RewriteExecutionRequirements {
        output_bytes,
        fields,
        work_bytes,
        max_depth: report.max_depth,
        states: report.states,
        allocations,
        retained_bytes,
        scratch_bytes,
    };
    Ok(PreparedRewrite {
        source,
        kind,
        desired,
        index_plan,
        requirements,
        report,
    })
}

fn check_execution_limits(
    requirements: RewriteExecutionRequirements,
    limits: RewriteExecutionLimits,
) -> Result<(), DecodeError> {
    for (observed, maximum, limit) in [
        (
            requirements.output_bytes,
            limits.output_bytes,
            DecodeLimit::OutputBytes {
                observed: requirements.output_bytes,
                maximum: limits.output_bytes,
            },
        ),
        (
            requirements.fields,
            limits.fields,
            DecodeLimit::Fields {
                observed: requirements.fields,
                maximum: limits.fields,
            },
        ),
        (
            requirements.work_bytes,
            limits.work_bytes,
            DecodeLimit::WorkBytes {
                observed: requirements.work_bytes,
                maximum: limits.work_bytes,
            },
        ),
        (
            requirements.states,
            limits.states,
            DecodeLimit::States {
                observed: requirements.states,
                maximum: limits.states,
            },
        ),
        (
            requirements.allocations,
            limits.allocations,
            DecodeLimit::Allocations {
                observed: requirements.allocations,
                maximum: limits.allocations,
            },
        ),
        (
            requirements.retained_bytes,
            limits.retained_bytes,
            DecodeLimit::RetainedBytes {
                observed: requirements.retained_bytes,
                maximum: limits.retained_bytes,
            },
        ),
        (
            requirements.scratch_bytes,
            limits.scratch_bytes,
            DecodeLimit::ScratchBytes {
                observed: requirements.scratch_bytes,
                maximum: limits.scratch_bytes,
            },
        ),
    ] {
        if observed > maximum {
            return Err(DecodeError::limit(limit));
        }
    }
    if requirements.max_depth > limits.max_depth {
        return Err(DecodeError::limit(DecodeLimit::Nesting {
            observed: requirements.max_depth,
            maximum: limits.max_depth,
        }));
    }
    Ok(())
}

fn measure_rewrite(
    source: &[u8],
    kind: RewriteKind,
    desired: &RewriteDesired,
    index_plan: &RewriteIndexPlan,
) -> Result<usize, DecodeError> {
    let mut matched = new_match_marks(index_plan)?;
    match (kind, desired) {
        (RewriteKind::TableInfo, RewriteDesired::TableInfo(value)) => {
            measure_table_info(source, value)
        },
        (RewriteKind::TableModel, RewriteDesired::TableModel(value)) => {
            measure_table_model(source, value, index_plan, &mut matched)
        },
        (RewriteKind::HiddenStatesOwner, RewriteDesired::HiddenStatesOwner(value)) => {
            measure_owner(source, value, index_plan, &mut matched)
        },
        (RewriteKind::HiddenStateExtent, RewriteDesired::HiddenStateExtent(value)) => {
            measure_extent(source, value, index_plan, &mut matched)
        },
        (RewriteKind::RowOrColumnState, RewriteDesired::RowOrColumnState(value)) => {
            measure_row_state(source, value)
        },
        _ => Err(DecodeError::projection()),
    }
}

fn add_len(total: &mut usize, value: usize) -> Result<(), DecodeError> {
    *total = total
        .checked_add(value)
        .ok_or_else(DecodeError::projection)?;
    Ok(())
}

fn varint_field_len(number: u32, value: u64) -> usize {
    varint_len(u64::from(number) << 3) + varint_len(value)
}

fn length_delimited_len(number: u32, payload_len: usize) -> Result<usize, DecodeError> {
    let payload_len = u64::try_from(payload_len).map_err(|_| DecodeError::projection())?;
    Ok(varint_len((u64::from(number) << 3) | 2) + varint_len(payload_len) + payload_len as usize)
}

fn canonical_uuid_len(value: UuidSnapshot) -> usize {
    varint_field_len(UUID_LOWER_FIELD, value.lower)
        + varint_field_len(UUID_UPPER_FIELD, value.upper)
}

fn canonical_reference_len(value: ReferenceSnapshot) -> usize {
    let mut total = varint_field_len(REFERENCE_IDENTIFIER_FIELD, value.identifier.get());
    if let Some(deprecated_type) = value.deprecated_type {
        total += varint_field_len(
            REFERENCE_DEPRECATED_TYPE_FIELD,
            deprecated_type as i64 as u64,
        );
    }
    if let Some(external) = value.deprecated_is_external {
        total += varint_field_len(REFERENCE_DEPRECATED_EXTERNAL_FIELD, u64::from(external));
    }
    total
}

fn canonical_row_state_len(value: &RowOrColumnStateSnapshot) -> Result<usize, DecodeError> {
    let mut total = length_delimited_len(
        ROW_STATE_UID_FIELD,
        canonical_uuid_len(value.row_or_column_uid),
    )?;
    if let Some(user_hidden) = value.user_hidden {
        add_len(
            &mut total,
            varint_field_len(ROW_STATE_USER_HIDDEN_FIELD, u64::from(user_hidden)),
        )?;
    }
    if let Some(filtered) = value.filtered {
        add_len(
            &mut total,
            varint_field_len(ROW_STATE_FILTERED_FIELD, u64::from(filtered)),
        )?;
    }
    if let Some(pivot_hidden) = value.pivot_hidden {
        add_len(
            &mut total,
            varint_field_len(ROW_STATE_PIVOT_HIDDEN_FIELD, u64::from(pivot_hidden)),
        )?;
    }
    Ok(total)
}

fn canonical_extent_len(value: &HiddenStateExtentSnapshot) -> Result<usize, DecodeError> {
    let mut total = length_delimited_len(
        EXTENT_UID_FIELD,
        canonical_uuid_len(value.hidden_state_extent_uid),
    )?;
    for state in &value.base_hidden_states {
        add_len(
            &mut total,
            length_delimited_len(EXTENT_BASE_STATES_FIELD, canonical_row_state_len(state)?)?,
        )?;
    }
    add_len(
        &mut total,
        varint_field_len(
            EXTENT_DIRECTION_FIELD,
            value.direction.native_value() as u64,
        ),
    )?;
    if let Some(value) = value.needs_to_update_filter_set_for_import {
        add_len(
            &mut total,
            varint_field_len(EXTENT_NEEDS_FILTER_UPDATE_FIELD, u64::from(value)),
        )?;
    }
    if let Some(filter_set) = value.filter_set {
        add_len(
            &mut total,
            length_delimited_len(EXTENT_FILTER_SET_FIELD, canonical_reference_len(filter_set))?,
        )?;
    }
    Ok(total)
}

fn canonical_hidden_states_len(value: &HiddenStatesSnapshot) -> Result<usize, DecodeError> {
    let mut total =
        length_delimited_len(STATE_UID_FIELD, canonical_uuid_len(value.hidden_states_uid))?;
    add_len(
        &mut total,
        length_delimited_len(
            STATE_COLUMN_EXTENT_FIELD,
            canonical_extent_len(value.column_hidden_state_extent())?,
        )?,
    )?;
    add_len(
        &mut total,
        length_delimited_len(
            STATE_ROW_EXTENT_FIELD,
            canonical_extent_len(value.row_hidden_state_extent())?,
        )?,
    )?;
    Ok(total)
}

fn canonical_owner_len(value: &HiddenStatesOwnerSnapshot) -> Result<usize, DecodeError> {
    let mut total = length_delimited_len(OWNER_UID_FIELD, canonical_uuid_len(value.owner_uid))?;
    for hidden_states in &value.hidden_states {
        add_len(
            &mut total,
            length_delimited_len(
                OWNER_STATES_FIELD,
                canonical_hidden_states_len(hidden_states)?,
            )?,
        )?;
    }
    Ok(total)
}

fn emit_uuid_canonical(value: UuidSnapshot, output: &mut Vec<u8>) {
    emit_varint_field(UUID_LOWER_FIELD, value.lower, output);
    emit_varint_field(UUID_UPPER_FIELD, value.upper, output);
}

fn emit_reference_canonical(value: ReferenceSnapshot, output: &mut Vec<u8>) {
    emit_varint_field(REFERENCE_IDENTIFIER_FIELD, value.identifier.get(), output);
    if let Some(deprecated_type) = value.deprecated_type {
        emit_varint_field(
            REFERENCE_DEPRECATED_TYPE_FIELD,
            deprecated_type as i64 as u64,
            output,
        );
    }
    if let Some(external) = value.deprecated_is_external {
        emit_optional_bool(REFERENCE_DEPRECATED_EXTERNAL_FIELD, Some(external), output);
    }
}

fn emit_row_state_canonical(
    value: &RowOrColumnStateSnapshot,
    output: &mut Vec<u8>,
) -> Result<(), DecodeError> {
    let payload_len = canonical_uuid_len(value.row_or_column_uid);
    emit_length_delimited_header(ROW_STATE_UID_FIELD, payload_len, output);
    emit_uuid_canonical(value.row_or_column_uid, output);
    emit_optional_bool(ROW_STATE_USER_HIDDEN_FIELD, value.user_hidden, output);
    emit_optional_bool(ROW_STATE_FILTERED_FIELD, value.filtered, output);
    emit_optional_bool(ROW_STATE_PIVOT_HIDDEN_FIELD, value.pivot_hidden, output);
    Ok(())
}

fn emit_extent_canonical(
    value: &HiddenStateExtentSnapshot,
    output: &mut Vec<u8>,
) -> Result<(), DecodeError> {
    let uid_len = canonical_uuid_len(value.hidden_state_extent_uid);
    emit_length_delimited_header(EXTENT_UID_FIELD, uid_len, output);
    emit_uuid_canonical(value.hidden_state_extent_uid, output);
    for state in &value.base_hidden_states {
        let state_len = canonical_row_state_len(state)?;
        emit_length_delimited_header(EXTENT_BASE_STATES_FIELD, state_len, output);
        emit_row_state_canonical(state, output)?;
    }
    emit_varint_field(
        EXTENT_DIRECTION_FIELD,
        value.direction.native_value() as u64,
        output,
    );
    emit_optional_bool(
        EXTENT_NEEDS_FILTER_UPDATE_FIELD,
        value.needs_to_update_filter_set_for_import,
        output,
    );
    if let Some(reference) = value.filter_set {
        let reference_len = canonical_reference_len(reference);
        emit_length_delimited_header(EXTENT_FILTER_SET_FIELD, reference_len, output);
        emit_reference_canonical(reference, output);
    }
    Ok(())
}

fn emit_hidden_states_canonical(
    value: &HiddenStatesSnapshot,
    output: &mut Vec<u8>,
) -> Result<(), DecodeError> {
    let uid_len = canonical_uuid_len(value.hidden_states_uid);
    emit_length_delimited_header(STATE_UID_FIELD, uid_len, output);
    emit_uuid_canonical(value.hidden_states_uid, output);
    let column_len = canonical_extent_len(value.column_hidden_state_extent())?;
    emit_length_delimited_header(STATE_COLUMN_EXTENT_FIELD, column_len, output);
    emit_extent_canonical(value.column_hidden_state_extent(), output)?;
    let row_len = canonical_extent_len(value.row_hidden_state_extent())?;
    emit_length_delimited_header(STATE_ROW_EXTENT_FIELD, row_len, output);
    emit_extent_canonical(value.row_hidden_state_extent(), output)?;
    Ok(())
}

fn emit_owner_canonical(
    value: &HiddenStatesOwnerSnapshot,
    output: &mut Vec<u8>,
) -> Result<(), DecodeError> {
    let uid_len = canonical_uuid_len(value.owner_uid);
    emit_length_delimited_header(OWNER_UID_FIELD, uid_len, output);
    emit_uuid_canonical(value.owner_uid, output);
    for state in &value.hidden_states {
        let state_len = canonical_hidden_states_len(state)?;
        emit_length_delimited_header(OWNER_STATES_FIELD, state_len, output);
        emit_hidden_states_canonical(state, output)?;
    }
    Ok(())
}

fn measure_table_info(source: &[u8], desired: &TableInfoSnapshot) -> Result<usize, DecodeError> {
    let mut total = 0usize;
    let mut model_seen = false;
    let mut view_seen = false;
    let mut hidden_seen = false;
    for_each_field_for_emit(source, |field| {
        match field.number {
            TABLE_INFO_MODEL_FIELD => {
                if model_seen {
                    return Err(DecodeError::duplicate("TST.TableInfoArchive.tableModel"));
                }
                model_seen = true;
                if source_reference_matches(source, field, desired.table_model)? {
                    add_len(&mut total, field.end - field.start)?;
                } else {
                    add_len(
                        &mut total,
                        length_delimited_len(
                            TABLE_INFO_MODEL_FIELD,
                            measure_reference_payload_rewrite(
                                field_bytes(source, field)?,
                                desired.table_model,
                            )?,
                        )?,
                    )?;
                }
            },
            TABLE_INFO_VIEW_UIDS_FIELD => {
                if view_seen {
                    return Err(DecodeError::duplicate(
                        "TST.TableInfoArchive.view_column_row_uids",
                    ));
                }
                view_seen = true;
                if let Some(value) = desired.view_column_row_uids {
                    if source_reference_matches(source, field, value)? {
                        add_len(&mut total, field.end - field.start)?;
                    } else {
                        add_len(
                            &mut total,
                            length_delimited_len(
                                TABLE_INFO_VIEW_UIDS_FIELD,
                                measure_reference_payload_rewrite(
                                    field_bytes(source, field)?,
                                    value,
                                )?,
                            )?,
                        )?;
                    }
                }
            },
            TABLE_INFO_HIDDEN_STATES_UUID_FIELD => {
                if hidden_seen {
                    return Err(DecodeError::duplicate(
                        "TST.TableInfoArchive.hidden_states_uuid",
                    ));
                }
                hidden_seen = true;
                if let Some(value) = desired.hidden_states_uuid {
                    if source_uuid_matches(source, field, value)? {
                        add_len(&mut total, field.end - field.start)?;
                    } else {
                        add_len(
                            &mut total,
                            length_delimited_len(
                                TABLE_INFO_HIDDEN_STATES_UUID_FIELD,
                                measure_uuid_payload_rewrite(field_bytes(source, field)?, value)?,
                            )?,
                        )?;
                    }
                }
            },
            _ => add_len(&mut total, field.end - field.start)?,
        }
        Ok(())
    })?;
    if !model_seen {
        return Err(DecodeError::missing("TST.TableInfoArchive.tableModel"));
    }
    if !hidden_seen && let Some(value) = desired.hidden_states_uuid {
        add_len(
            &mut total,
            length_delimited_len(
                TABLE_INFO_HIDDEN_STATES_UUID_FIELD,
                canonical_uuid_len(value),
            )?,
        )?;
    }
    if !view_seen && let Some(value) = desired.view_column_row_uids {
        add_len(
            &mut total,
            length_delimited_len(TABLE_INFO_VIEW_UIDS_FIELD, canonical_reference_len(value))?,
        )?;
    }
    Ok(total)
}

fn measure_table_model(
    source: &[u8],
    desired: &TableModelSnapshot,
    index_plan: &RewriteIndexPlan,
    matched: &mut [u8],
) -> Result<usize, DecodeError> {
    let mut total = 0usize;
    let mut replaced = [false; 11];
    for_each_field_for_emit(source, |field| {
        let slot = match field.number {
            TABLE_MODEL_ROWS_FIELD => Some(0),
            TABLE_MODEL_COLUMNS_FIELD => Some(1),
            TABLE_MODEL_HIDDEN_ROWS_FIELD => Some(2),
            TABLE_MODEL_HIDDEN_COLUMNS_FIELD => Some(3),
            TABLE_MODEL_FILTERED_ROWS_FIELD => Some(4),
            TABLE_MODEL_USER_HIDDEN_ROWS_FIELD => Some(5),
            TABLE_MODEL_USER_HIDDEN_COLUMNS_FIELD => Some(6),
            TABLE_MODEL_COLUMN_FORMULA_OWNER_FIELD => Some(7),
            TABLE_MODEL_ROW_FORMULA_OWNER_FIELD => Some(8),
            TABLE_MODEL_BASE_COLUMN_ROW_UIDS_FIELD => Some(9),
            TABLE_MODEL_HIDDEN_STATES_OWNER_FIELD => Some(10),
            _ => None,
        };
        let Some(slot) = slot else {
            return add_len(&mut total, field.end - field.start);
        };
        if replaced[slot] {
            return Err(DecodeError::duplicate(
                "TST.TableModelArchive.hidden-state field",
            ));
        }
        replaced[slot] = true;
        match field.number {
            TABLE_MODEL_ROWS_FIELD => measure_u32_field(
                &mut total,
                source,
                field,
                TABLE_MODEL_ROWS_FIELD,
                desired.number_of_rows,
            ),
            TABLE_MODEL_COLUMNS_FIELD => measure_u32_field(
                &mut total,
                source,
                field,
                TABLE_MODEL_COLUMNS_FIELD,
                desired.number_of_columns,
            ),
            TABLE_MODEL_HIDDEN_ROWS_FIELD => measure_optional_u32_field(
                &mut total,
                source,
                field,
                TABLE_MODEL_HIDDEN_ROWS_FIELD,
                desired.number_of_hidden_rows,
            ),
            TABLE_MODEL_HIDDEN_COLUMNS_FIELD => measure_optional_u32_field(
                &mut total,
                source,
                field,
                TABLE_MODEL_HIDDEN_COLUMNS_FIELD,
                desired.number_of_hidden_columns,
            ),
            TABLE_MODEL_FILTERED_ROWS_FIELD => measure_optional_u32_field(
                &mut total,
                source,
                field,
                TABLE_MODEL_FILTERED_ROWS_FIELD,
                desired.number_of_filtered_rows,
            ),
            TABLE_MODEL_USER_HIDDEN_ROWS_FIELD => measure_optional_u32_field(
                &mut total,
                source,
                field,
                TABLE_MODEL_USER_HIDDEN_ROWS_FIELD,
                desired.number_of_user_hidden_rows,
            ),
            TABLE_MODEL_USER_HIDDEN_COLUMNS_FIELD => measure_optional_u32_field(
                &mut total,
                source,
                field,
                TABLE_MODEL_USER_HIDDEN_COLUMNS_FIELD,
                desired.number_of_user_hidden_columns,
            ),
            TABLE_MODEL_COLUMN_FORMULA_OWNER_FIELD => measure_optional_reference_field(
                &mut total,
                source,
                field,
                TABLE_MODEL_COLUMN_FORMULA_OWNER_FIELD,
                desired.hidden_state_formula_owner_for_columns,
            ),
            TABLE_MODEL_ROW_FORMULA_OWNER_FIELD => measure_optional_reference_field(
                &mut total,
                source,
                field,
                TABLE_MODEL_ROW_FORMULA_OWNER_FIELD,
                desired.hidden_state_formula_owner_for_rows,
            ),
            TABLE_MODEL_BASE_COLUMN_ROW_UIDS_FIELD => measure_optional_reference_field(
                &mut total,
                source,
                field,
                TABLE_MODEL_BASE_COLUMN_ROW_UIDS_FIELD,
                desired.base_column_row_uids,
            ),
            TABLE_MODEL_HIDDEN_STATES_OWNER_FIELD => {
                if let Some(owner) = desired.hidden_states_owner.as_ref() {
                    let payload = source_field_payload(source, field)?;
                    add_len(
                        &mut total,
                        length_delimited_len(
                            TABLE_MODEL_HIDDEN_STATES_OWNER_FIELD,
                            measure_owner(payload, owner, index_plan, matched)?,
                        )?,
                    )?;
                }
                Ok(())
            },
            _ => Err(DecodeError::projection()),
        }
    })?;
    if !replaced[2] {
        measure_optional_u32(
            &mut total,
            TABLE_MODEL_HIDDEN_ROWS_FIELD,
            desired.number_of_hidden_rows,
        )?;
    }
    if !replaced[3] {
        measure_optional_u32(
            &mut total,
            TABLE_MODEL_HIDDEN_COLUMNS_FIELD,
            desired.number_of_hidden_columns,
        )?;
    }
    if !replaced[4] {
        measure_optional_u32(
            &mut total,
            TABLE_MODEL_FILTERED_ROWS_FIELD,
            desired.number_of_filtered_rows,
        )?;
    }
    if !replaced[5] {
        measure_optional_u32(
            &mut total,
            TABLE_MODEL_USER_HIDDEN_ROWS_FIELD,
            desired.number_of_user_hidden_rows,
        )?;
    }
    if !replaced[6] {
        measure_optional_u32(
            &mut total,
            TABLE_MODEL_USER_HIDDEN_COLUMNS_FIELD,
            desired.number_of_user_hidden_columns,
        )?;
    }
    if !replaced[7] {
        measure_optional_reference(
            &mut total,
            TABLE_MODEL_COLUMN_FORMULA_OWNER_FIELD,
            desired.hidden_state_formula_owner_for_columns,
        )?;
    }
    if !replaced[8] {
        measure_optional_reference(
            &mut total,
            TABLE_MODEL_ROW_FORMULA_OWNER_FIELD,
            desired.hidden_state_formula_owner_for_rows,
        )?;
    }
    if !replaced[9] {
        measure_optional_reference(
            &mut total,
            TABLE_MODEL_BASE_COLUMN_ROW_UIDS_FIELD,
            desired.base_column_row_uids,
        )?;
    }
    if !replaced[10]
        && let Some(owner) = desired.hidden_states_owner.as_ref()
    {
        add_len(
            &mut total,
            length_delimited_len(
                TABLE_MODEL_HIDDEN_STATES_OWNER_FIELD,
                canonical_owner_len(owner)?,
            )?,
        )?;
    }
    Ok(total)
}

fn measure_optional_u32(
    total: &mut usize,
    number: u32,
    value: Option<u32>,
) -> Result<(), DecodeError> {
    if let Some(value) = value {
        add_len(total, varint_field_len(number, u64::from(value)))?;
    }
    Ok(())
}

fn measure_optional_reference(
    total: &mut usize,
    number: u32,
    value: Option<ReferenceSnapshot>,
) -> Result<(), DecodeError> {
    if let Some(value) = value {
        add_len(
            total,
            length_delimited_len(number, canonical_reference_len(value))?,
        )?;
    }
    Ok(())
}

fn measure_u32_field(
    total: &mut usize,
    source: &[u8],
    field: FieldSpan,
    number: u32,
    value: u32,
) -> Result<(), DecodeError> {
    if source_varint_matches(source, field, u64::from(value))? {
        add_len(total, field.end - field.start)
    } else {
        add_len(total, varint_field_len(number, u64::from(value)))
    }
}

fn measure_optional_u32_field(
    total: &mut usize,
    source: &[u8],
    field: FieldSpan,
    number: u32,
    value: Option<u32>,
) -> Result<(), DecodeError> {
    if let Some(value) = value {
        measure_u32_field(total, source, field, number, value)?;
    }
    Ok(())
}

fn measure_optional_bool_field(
    total: &mut usize,
    source: &[u8],
    field: FieldSpan,
    number: u32,
    value: Option<bool>,
) -> Result<(), DecodeError> {
    if let Some(value) = value {
        if source_varint_matches(source, field, u64::from(value))? {
            add_len(total, field.end - field.start)?;
        } else {
            measure_optional_bool(total, number, Some(value))?;
        }
    }
    Ok(())
}

fn measure_optional_reference_field(
    total: &mut usize,
    source: &[u8],
    field: FieldSpan,
    number: u32,
    value: Option<ReferenceSnapshot>,
) -> Result<(), DecodeError> {
    if let Some(value) = value {
        if source_reference_matches(source, field, value)? {
            add_len(total, field.end - field.start)?;
        } else {
            add_len(
                total,
                length_delimited_len(
                    number,
                    measure_reference_payload_rewrite(field_bytes(source, field)?, value)?,
                )?,
            )?;
        }
    }
    Ok(())
}

fn measure_owner(
    source: &[u8],
    desired: &HiddenStatesOwnerSnapshot,
    index_plan: &RewriteIndexPlan,
    matched: &mut [u8],
) -> Result<usize, DecodeError> {
    let mut total = 0usize;
    let mut owner_seen = false;
    let desired_range = index_plan.owner_range.ok_or_else(DecodeError::projection)?;
    for_each_field_for_emit(source, |field| {
        match field.number {
            OWNER_UID_FIELD => {
                if owner_seen {
                    return Err(DecodeError::duplicate(
                        "TST.HiddenStatesOwnerArchive.owner_uid",
                    ));
                }
                owner_seen = true;
                if source_uuid_matches(source, field, desired.owner_uid)? {
                    add_len(&mut total, field.end - field.start)?;
                } else {
                    add_len(
                        &mut total,
                        length_delimited_len(
                            OWNER_UID_FIELD,
                            canonical_uuid_len(desired.owner_uid),
                        )?,
                    )?;
                }
            },
            OWNER_STATES_FIELD => {
                let payload = source_field_payload(source, field)?;
                let Some(uid) = state_uid_from_payload(payload)? else {
                    return Err(DecodeError::projection());
                };
                if let Some(index) = find_uid_index(index_plan, desired_range, uid)? {
                    mark_match(matched, desired_range, index)?;
                    let state = &desired.hidden_states[index];
                    add_len(
                        &mut total,
                        length_delimited_len(
                            OWNER_STATES_FIELD,
                            measure_hidden_states(payload, state, index_plan, matched)?,
                        )?,
                    )?;
                }
            },
            _ => add_len(&mut total, field.end - field.start)?,
        }
        Ok(())
    })?;
    if !owner_seen {
        return Err(DecodeError::missing(
            "TST.HiddenStatesOwnerArchive.owner_uid",
        ));
    }
    for (index, state) in desired.hidden_states.iter().enumerate() {
        if !is_matched(matched, desired_range, index)? {
            add_len(
                &mut total,
                length_delimited_len(OWNER_STATES_FIELD, canonical_hidden_states_len(state)?)?,
            )?;
        }
    }
    Ok(total)
}

fn measure_hidden_states(
    source: &[u8],
    desired: &HiddenStatesSnapshot,
    index_plan: &RewriteIndexPlan,
    matched: &mut [u8],
) -> Result<usize, DecodeError> {
    let mut total = 0usize;
    let mut uid_seen = false;
    let mut column_seen = false;
    let mut row_seen = false;
    for_each_field_for_emit(source, |field| {
        match field.number {
            STATE_UID_FIELD => {
                if uid_seen {
                    return Err(DecodeError::duplicate(
                        "TST.HiddenStatesArchive.hidden_states_uid",
                    ));
                }
                uid_seen = true;
                if source_uuid_matches(source, field, desired.hidden_states_uid)? {
                    add_len(&mut total, field.end - field.start)?;
                } else {
                    add_len(
                        &mut total,
                        length_delimited_len(
                            STATE_UID_FIELD,
                            measure_uuid_payload_rewrite(
                                field_bytes(source, field)?,
                                desired.hidden_states_uid,
                            )?,
                        )?,
                    )?;
                }
            },
            STATE_COLUMN_EXTENT_FIELD => {
                if column_seen {
                    return Err(DecodeError::duplicate(
                        "TST.HiddenStatesArchive.column_hidden_state_extent",
                    ));
                }
                column_seen = true;
                let payload = source_field_payload(source, field)?;
                add_len(
                    &mut total,
                    length_delimited_len(
                        STATE_COLUMN_EXTENT_FIELD,
                        measure_extent(
                            payload,
                            desired.column_hidden_state_extent(),
                            index_plan,
                            matched,
                        )?,
                    )?,
                )?;
            },
            STATE_ROW_EXTENT_FIELD => {
                if row_seen {
                    return Err(DecodeError::duplicate(
                        "TST.HiddenStatesArchive.row_hidden_state_extent",
                    ));
                }
                row_seen = true;
                let payload = source_field_payload(source, field)?;
                add_len(
                    &mut total,
                    length_delimited_len(
                        STATE_ROW_EXTENT_FIELD,
                        measure_extent(
                            payload,
                            desired.row_hidden_state_extent(),
                            index_plan,
                            matched,
                        )?,
                    )?,
                )?;
            },
            _ => add_len(&mut total, field.end - field.start)?,
        }
        Ok(())
    })?;
    if !uid_seen {
        return Err(DecodeError::missing(
            "TST.HiddenStatesArchive.hidden_states_uid",
        ));
    }
    if !column_seen || !row_seen {
        return Err(DecodeError::missing("TST.HiddenStatesArchive.extent"));
    }
    Ok(total)
}

fn measure_extent(
    source: &[u8],
    desired: &HiddenStateExtentSnapshot,
    index_plan: &RewriteIndexPlan,
    matched: &mut [u8],
) -> Result<usize, DecodeError> {
    let mut total = 0usize;
    let mut uid_seen = false;
    let mut direction_seen = false;
    let mut needs_filter_seen = false;
    let mut filter_set_seen = false;
    let desired_range = index_plan
        .extent_range(desired.hidden_state_extent_uid)
        .ok_or_else(DecodeError::projection)?;
    for_each_field_for_emit(source, |field| {
        match field.number {
            EXTENT_UID_FIELD => {
                if uid_seen {
                    return Err(DecodeError::duplicate(
                        "TST.HiddenStateExtentArchive.hidden_state_extent_uid",
                    ));
                }
                uid_seen = true;
                if source_uuid_matches(source, field, desired.hidden_state_extent_uid)? {
                    add_len(&mut total, field.end - field.start)?;
                } else {
                    add_len(
                        &mut total,
                        length_delimited_len(
                            EXTENT_UID_FIELD,
                            measure_uuid_payload_rewrite(
                                field_bytes(source, field)?,
                                desired.hidden_state_extent_uid,
                            )?,
                        )?,
                    )?;
                }
            },
            EXTENT_DIRECTION_FIELD => {
                if direction_seen {
                    return Err(DecodeError::duplicate(
                        "TST.HiddenStateExtentArchive.row_or_column_direction",
                    ));
                }
                direction_seen = true;
                if source_varint_matches(source, field, desired.direction.native_value() as u64)? {
                    add_len(&mut total, field.end - field.start)?;
                } else {
                    add_len(
                        &mut total,
                        varint_field_len(
                            EXTENT_DIRECTION_FIELD,
                            desired.direction.native_value() as u64,
                        ),
                    )?;
                }
            },
            EXTENT_BASE_STATES_FIELD => {
                let payload = source_field_payload(source, field)?;
                let Some(uid) = state_uid_from_payload(payload)? else {
                    return Err(DecodeError::projection());
                };
                if let Some(index) = find_uid_index(index_plan, desired_range, uid)? {
                    mark_match(matched, desired_range, index)?;
                    let state = &desired.base_hidden_states[index];
                    add_len(
                        &mut total,
                        length_delimited_len(
                            EXTENT_BASE_STATES_FIELD,
                            measure_row_state(payload, state)?,
                        )?,
                    )?;
                }
            },
            EXTENT_NEEDS_FILTER_UPDATE_FIELD => {
                needs_filter_seen = true;
                if desired.needs_to_update_filter_set_for_import.is_some() {
                    measure_optional_bool_field(
                        &mut total,
                        source,
                        field,
                        EXTENT_NEEDS_FILTER_UPDATE_FIELD,
                        desired.needs_to_update_filter_set_for_import,
                    )?;
                }
            },
            EXTENT_FILTER_SET_FIELD => {
                filter_set_seen = true;
                if let Some(reference) = desired.filter_set {
                    if source_reference_matches(source, field, reference)? {
                        add_len(&mut total, field.end - field.start)?;
                    } else {
                        add_len(
                            &mut total,
                            length_delimited_len(
                                EXTENT_FILTER_SET_FIELD,
                                measure_reference_payload_rewrite(
                                    field_bytes(source, field)?,
                                    reference,
                                )?,
                            )?,
                        )?;
                    }
                }
            },
            _ => add_len(&mut total, field.end - field.start)?,
        }
        Ok(())
    })?;
    if !uid_seen {
        return Err(DecodeError::missing(
            "TST.HiddenStateExtentArchive.hidden_state_extent_uid",
        ));
    }
    if !direction_seen {
        return Err(DecodeError::missing(
            "TST.HiddenStateExtentArchive.row_or_column_direction",
        ));
    }
    for (index, state) in desired.base_hidden_states.iter().enumerate() {
        if !is_matched(matched, desired_range, index)? {
            add_len(
                &mut total,
                length_delimited_len(EXTENT_BASE_STATES_FIELD, canonical_row_state_len(state)?)?,
            )?;
        }
    }
    if !needs_filter_seen && let Some(value) = desired.needs_to_update_filter_set_for_import {
        add_len(
            &mut total,
            varint_field_len(EXTENT_NEEDS_FILTER_UPDATE_FIELD, u64::from(value)),
        )?;
    }
    if !filter_set_seen && let Some(reference) = desired.filter_set {
        add_len(
            &mut total,
            length_delimited_len(EXTENT_FILTER_SET_FIELD, canonical_reference_len(reference))?,
        )?;
    }
    Ok(total)
}

fn measure_row_state(
    source: &[u8],
    desired: &RowOrColumnStateSnapshot,
) -> Result<usize, DecodeError> {
    let mut total = 0usize;
    let mut uid_seen = false;
    let mut replaced = [false; 3];
    for_each_field_for_emit(source, |field| {
        match field.number {
            ROW_STATE_UID_FIELD => {
                if uid_seen {
                    return Err(DecodeError::duplicate(
                        "TST.HiddenStateExtentArchive.RowOrColumnState.row_or_column_uid",
                    ));
                }
                uid_seen = true;
                if source_uuid_matches(source, field, desired.row_or_column_uid)? {
                    add_len(&mut total, field.end - field.start)?;
                } else {
                    add_len(
                        &mut total,
                        length_delimited_len(
                            ROW_STATE_UID_FIELD,
                            measure_uuid_payload_rewrite(
                                field_bytes(source, field)?,
                                desired.row_or_column_uid,
                            )?,
                        )?,
                    )?;
                }
            },
            ROW_STATE_USER_HIDDEN_FIELD => {
                replaced[0] = true;
                measure_optional_bool_field(
                    &mut total,
                    source,
                    field,
                    ROW_STATE_USER_HIDDEN_FIELD,
                    desired.user_hidden,
                )?;
            },
            ROW_STATE_FILTERED_FIELD => {
                replaced[1] = true;
                measure_optional_bool_field(
                    &mut total,
                    source,
                    field,
                    ROW_STATE_FILTERED_FIELD,
                    desired.filtered,
                )?;
            },
            ROW_STATE_PIVOT_HIDDEN_FIELD => {
                replaced[2] = true;
                measure_optional_bool_field(
                    &mut total,
                    source,
                    field,
                    ROW_STATE_PIVOT_HIDDEN_FIELD,
                    desired.pivot_hidden,
                )?;
            },
            _ => add_len(&mut total, field.end - field.start)?,
        }
        Ok(())
    })?;
    if !uid_seen {
        return Err(DecodeError::missing(
            "TST.HiddenStateExtentArchive.RowOrColumnState.row_or_column_uid",
        ));
    }
    if !replaced[0] {
        measure_optional_bool(&mut total, ROW_STATE_USER_HIDDEN_FIELD, desired.user_hidden)?;
    }
    if !replaced[1] {
        measure_optional_bool(&mut total, ROW_STATE_FILTERED_FIELD, desired.filtered)?;
    }
    if !replaced[2] {
        measure_optional_bool(
            &mut total,
            ROW_STATE_PIVOT_HIDDEN_FIELD,
            desired.pivot_hidden,
        )?;
    }
    Ok(total)
}

fn measure_optional_bool(
    total: &mut usize,
    number: u32,
    value: Option<bool>,
) -> Result<(), DecodeError> {
    if let Some(value) = value {
        add_len(total, varint_field_len(number, u64::from(value)))?;
    }
    Ok(())
}

fn emit_rewrite(
    source: &[u8],
    kind: RewriteKind,
    desired: &RewriteDesired,
    index_plan: &RewriteIndexPlan,
    matched: &mut [u8],
    output: &mut Vec<u8>,
) -> Result<(), DecodeError> {
    match (kind, desired) {
        (RewriteKind::TableInfo, RewriteDesired::TableInfo(value)) => {
            emit_table_info(source, value, output)
        },
        (RewriteKind::TableModel, RewriteDesired::TableModel(value)) => {
            emit_table_model(source, value, index_plan, matched, output)
        },
        (RewriteKind::HiddenStatesOwner, RewriteDesired::HiddenStatesOwner(value)) => {
            emit_owner(source, value, index_plan, matched, output)
        },
        (RewriteKind::HiddenStateExtent, RewriteDesired::HiddenStateExtent(value)) => {
            emit_extent(source, value, index_plan, matched, output)
        },
        (RewriteKind::RowOrColumnState, RewriteDesired::RowOrColumnState(value)) => {
            emit_row_state(source, value, output)
        },
        _ => Err(DecodeError::projection()),
    }
}

fn emit_table_info(
    source: &[u8],
    desired: &TableInfoSnapshot,
    output: &mut Vec<u8>,
) -> Result<(), DecodeError> {
    let mut model_seen = false;
    let mut hidden_seen = false;
    let mut view_seen = false;
    for_each_field_for_emit(source, |field| {
        match field.number {
            TABLE_INFO_MODEL_FIELD => {
                if model_seen {
                    return Err(DecodeError::duplicate("TST.TableInfoArchive.tableModel"));
                }
                model_seen = true;
                if source_reference_matches(source, field, desired.table_model)? {
                    output.extend_from_slice(&source[field.start..field.end]);
                } else {
                    emit_reference_field_preserving(
                        source,
                        field,
                        TABLE_INFO_MODEL_FIELD,
                        desired.table_model,
                        output,
                    )?;
                }
            },
            TABLE_INFO_VIEW_UIDS_FIELD => {
                if view_seen {
                    return Err(DecodeError::duplicate(
                        "TST.TableInfoArchive.view_column_row_uids",
                    ));
                }
                view_seen = true;
                if let Some(value) = desired.view_column_row_uids {
                    if source_reference_matches(source, field, value)? {
                        output.extend_from_slice(&source[field.start..field.end]);
                    } else {
                        emit_reference_field_preserving(
                            source,
                            field,
                            TABLE_INFO_VIEW_UIDS_FIELD,
                            value,
                            output,
                        )?;
                    }
                }
            },
            TABLE_INFO_HIDDEN_STATES_UUID_FIELD => {
                if hidden_seen {
                    return Err(DecodeError::duplicate(
                        "TST.TableInfoArchive.hidden_states_uuid",
                    ));
                }
                hidden_seen = true;
                if let Some(value) = desired.hidden_states_uuid {
                    if source_uuid_matches(source, field, value)? {
                        output.extend_from_slice(&source[field.start..field.end]);
                    } else {
                        emit_uuid_field_preserving(
                            source,
                            field,
                            TABLE_INFO_HIDDEN_STATES_UUID_FIELD,
                            value,
                            output,
                        )?;
                    }
                }
            },
            _ => output.extend_from_slice(&source[field.start..field.end]),
        }
        Ok(())
    })?;
    if !model_seen {
        return Err(DecodeError::missing("TST.TableInfoArchive.tableModel"));
    }
    if !hidden_seen && desired.hidden_states_uuid.is_some() {
        let bytes = canonical_uuid(
            desired
                .hidden_states_uuid
                .ok_or_else(DecodeError::projection)?,
        )?;
        emit_length_delimited(TABLE_INFO_HIDDEN_STATES_UUID_FIELD, &bytes, output);
    }
    if !view_seen && let Some(value) = desired.view_column_row_uids {
        let bytes = canonical_reference(value)?;
        emit_length_delimited(TABLE_INFO_VIEW_UIDS_FIELD, &bytes, output);
    }
    Ok(())
}

fn emit_table_model(
    source: &[u8],
    desired: &TableModelSnapshot,
    index_plan: &RewriteIndexPlan,
    matched: &mut [u8],
    output: &mut Vec<u8>,
) -> Result<(), DecodeError> {
    let mut replaced = [false; 11];
    for_each_field_for_emit(source, |field| {
        let slot = match field.number {
            TABLE_MODEL_ROWS_FIELD => Some(0),
            TABLE_MODEL_COLUMNS_FIELD => Some(1),
            TABLE_MODEL_HIDDEN_ROWS_FIELD => Some(2),
            TABLE_MODEL_HIDDEN_COLUMNS_FIELD => Some(3),
            TABLE_MODEL_FILTERED_ROWS_FIELD => Some(4),
            TABLE_MODEL_USER_HIDDEN_ROWS_FIELD => Some(5),
            TABLE_MODEL_USER_HIDDEN_COLUMNS_FIELD => Some(6),
            TABLE_MODEL_COLUMN_FORMULA_OWNER_FIELD => Some(7),
            TABLE_MODEL_ROW_FORMULA_OWNER_FIELD => Some(8),
            TABLE_MODEL_BASE_COLUMN_ROW_UIDS_FIELD => Some(9),
            TABLE_MODEL_HIDDEN_STATES_OWNER_FIELD => Some(10),
            _ => None,
        };
        if slot.is_none() {
            output.extend_from_slice(&source[field.start..field.end]);
            return Ok(());
        }
        let Some(slot) = slot else {
            return Err(DecodeError::projection());
        };
        if replaced[slot] {
            return Err(DecodeError::duplicate(
                "TST.TableModelArchive.hidden-state field",
            ));
        }
        replaced[slot] = true;
        match field.number {
            TABLE_MODEL_ROWS_FIELD => emit_u32_field_preserving(
                source,
                field,
                TABLE_MODEL_ROWS_FIELD,
                desired.number_of_rows,
                output,
            )?,
            TABLE_MODEL_COLUMNS_FIELD => emit_u32_field_preserving(
                source,
                field,
                TABLE_MODEL_COLUMNS_FIELD,
                desired.number_of_columns,
                output,
            )?,
            TABLE_MODEL_HIDDEN_ROWS_FIELD => emit_optional_u32_field_preserving(
                source,
                field,
                TABLE_MODEL_HIDDEN_ROWS_FIELD,
                desired.number_of_hidden_rows,
                output,
            )?,
            TABLE_MODEL_HIDDEN_COLUMNS_FIELD => emit_optional_u32_field_preserving(
                source,
                field,
                TABLE_MODEL_HIDDEN_COLUMNS_FIELD,
                desired.number_of_hidden_columns,
                output,
            )?,
            TABLE_MODEL_FILTERED_ROWS_FIELD => emit_optional_u32_field_preserving(
                source,
                field,
                TABLE_MODEL_FILTERED_ROWS_FIELD,
                desired.number_of_filtered_rows,
                output,
            )?,
            TABLE_MODEL_USER_HIDDEN_ROWS_FIELD => emit_optional_u32_field_preserving(
                source,
                field,
                TABLE_MODEL_USER_HIDDEN_ROWS_FIELD,
                desired.number_of_user_hidden_rows,
                output,
            )?,
            TABLE_MODEL_USER_HIDDEN_COLUMNS_FIELD => emit_optional_u32_field_preserving(
                source,
                field,
                TABLE_MODEL_USER_HIDDEN_COLUMNS_FIELD,
                desired.number_of_user_hidden_columns,
                output,
            )?,
            TABLE_MODEL_COLUMN_FORMULA_OWNER_FIELD => emit_optional_reference_field_preserving(
                source,
                field,
                TABLE_MODEL_COLUMN_FORMULA_OWNER_FIELD,
                desired.hidden_state_formula_owner_for_columns,
                output,
            )?,
            TABLE_MODEL_ROW_FORMULA_OWNER_FIELD => emit_optional_reference_field_preserving(
                source,
                field,
                TABLE_MODEL_ROW_FORMULA_OWNER_FIELD,
                desired.hidden_state_formula_owner_for_rows,
                output,
            )?,
            TABLE_MODEL_BASE_COLUMN_ROW_UIDS_FIELD => emit_optional_reference_field_preserving(
                source,
                field,
                TABLE_MODEL_BASE_COLUMN_ROW_UIDS_FIELD,
                desired.base_column_row_uids,
                output,
            )?,
            TABLE_MODEL_HIDDEN_STATES_OWNER_FIELD => {
                if let Some(owner) = desired.hidden_states_owner.as_ref() {
                    let raw = source_field_payload(source, field)?;
                    let payload_len = measure_owner(raw, owner, index_plan, matched)?;
                    emit_length_delimited_header(
                        TABLE_MODEL_HIDDEN_STATES_OWNER_FIELD,
                        payload_len,
                        output,
                    );
                    emit_owner(raw, owner, index_plan, matched, output)?;
                }
            },
            _ => unreachable!("selected model field"),
        }
        Ok(())
    })?;
    if !replaced[2] {
        emit_optional_u32(
            TABLE_MODEL_HIDDEN_ROWS_FIELD,
            desired.number_of_hidden_rows,
            output,
        );
    }
    if !replaced[3] {
        emit_optional_u32(
            TABLE_MODEL_HIDDEN_COLUMNS_FIELD,
            desired.number_of_hidden_columns,
            output,
        );
    }
    if !replaced[4] {
        emit_optional_u32(
            TABLE_MODEL_FILTERED_ROWS_FIELD,
            desired.number_of_filtered_rows,
            output,
        );
    }
    if !replaced[5] {
        emit_optional_u32(
            TABLE_MODEL_USER_HIDDEN_ROWS_FIELD,
            desired.number_of_user_hidden_rows,
            output,
        );
    }
    if !replaced[6] {
        emit_optional_u32(
            TABLE_MODEL_USER_HIDDEN_COLUMNS_FIELD,
            desired.number_of_user_hidden_columns,
            output,
        );
    }
    if !replaced[7] {
        emit_optional_reference(
            TABLE_MODEL_COLUMN_FORMULA_OWNER_FIELD,
            desired.hidden_state_formula_owner_for_columns,
            output,
        )?;
    }
    if !replaced[8] {
        emit_optional_reference(
            TABLE_MODEL_ROW_FORMULA_OWNER_FIELD,
            desired.hidden_state_formula_owner_for_rows,
            output,
        )?;
    }
    if !replaced[9] {
        emit_optional_reference(
            TABLE_MODEL_BASE_COLUMN_ROW_UIDS_FIELD,
            desired.base_column_row_uids,
            output,
        )?;
    }
    if !replaced[10]
        && let Some(owner) = desired.hidden_states_owner.as_ref()
    {
        let payload_len = canonical_owner_len(owner)?;
        emit_length_delimited_header(TABLE_MODEL_HIDDEN_STATES_OWNER_FIELD, payload_len, output);
        emit_owner_canonical(owner, output)?;
    }
    Ok(())
}

fn emit_owner(
    source: &[u8],
    desired: &HiddenStatesOwnerSnapshot,
    index_plan: &RewriteIndexPlan,
    matched: &mut [u8],
    output: &mut Vec<u8>,
) -> Result<(), DecodeError> {
    let mut owner_seen = false;
    let desired_range = index_plan.owner_range.ok_or_else(DecodeError::projection)?;
    for_each_field_for_emit(source, |field| {
        match field.number {
            OWNER_UID_FIELD => {
                if owner_seen {
                    return Err(DecodeError::duplicate(
                        "TST.HiddenStatesOwnerArchive.owner_uid",
                    ));
                }
                owner_seen = true;
                if source_uuid_matches(source, field, desired.owner_uid)? {
                    output.extend_from_slice(&source[field.start..field.end]);
                } else {
                    emit_uuid_field_preserving(
                        source,
                        field,
                        OWNER_UID_FIELD,
                        desired.owner_uid,
                        output,
                    )?;
                }
            },
            OWNER_STATES_FIELD => {
                let payload = source_field_payload(source, field)?;
                let Some(uid) = state_uid_from_payload(payload)? else {
                    return Err(DecodeError::projection());
                };
                if let Some(index) = find_uid_index(index_plan, desired_range, uid)? {
                    mark_match(matched, desired_range, index)?;
                    let state = &desired.hidden_states[index];
                    let payload_len = measure_hidden_states(payload, state, index_plan, matched)?;
                    emit_length_delimited_header(OWNER_STATES_FIELD, payload_len, output);
                    emit_hidden_states_from_source(
                        payload,
                        state,
                        payload_len,
                        index_plan,
                        matched,
                        output,
                    )?;
                }
            },
            _ => output.extend_from_slice(&source[field.start..field.end]),
        }
        Ok(())
    })?;
    if !owner_seen {
        return Err(DecodeError::missing(
            "TST.HiddenStatesOwnerArchive.owner_uid",
        ));
    }
    for (index, state) in desired.hidden_states.iter().enumerate() {
        if !is_matched(matched, desired_range, index)? {
            let payload_len = canonical_hidden_states_len(state)?;
            emit_length_delimited_header(OWNER_STATES_FIELD, payload_len, output);
            emit_hidden_states_canonical(state, output)?;
        }
    }
    Ok(())
}

fn emit_extent(
    source: &[u8],
    desired: &HiddenStateExtentSnapshot,
    index_plan: &RewriteIndexPlan,
    matched: &mut [u8],
    output: &mut Vec<u8>,
) -> Result<(), DecodeError> {
    let mut uid_seen = false;
    let mut direction_seen = false;
    let mut needs_filter_seen = false;
    let mut filter_set_seen = false;
    let desired_range = index_plan
        .extent_range(desired.hidden_state_extent_uid)
        .ok_or_else(DecodeError::projection)?;
    for_each_field_for_emit(source, |field| {
        match field.number {
            EXTENT_UID_FIELD => {
                if uid_seen {
                    return Err(DecodeError::duplicate(
                        "TST.HiddenStateExtentArchive.hidden_state_extent_uid",
                    ));
                }
                uid_seen = true;
                if source_uuid_matches(source, field, desired.hidden_state_extent_uid)? {
                    output.extend_from_slice(&source[field.start..field.end]);
                } else {
                    emit_uuid_field_preserving(
                        source,
                        field,
                        EXTENT_UID_FIELD,
                        desired.hidden_state_extent_uid,
                        output,
                    )?;
                }
            },
            EXTENT_DIRECTION_FIELD => {
                if direction_seen {
                    return Err(DecodeError::duplicate(
                        "TST.HiddenStateExtentArchive.row_or_column_direction",
                    ));
                }
                direction_seen = true;
                if source_varint_matches(source, field, desired.direction.native_value() as u64)? {
                    output.extend_from_slice(&source[field.start..field.end]);
                } else {
                    emit_varint_field(
                        EXTENT_DIRECTION_FIELD,
                        desired.direction.native_value() as u64,
                        output,
                    );
                }
            },
            EXTENT_BASE_STATES_FIELD => {
                let payload = source_field_payload(source, field)?;
                let Some(uid) = state_uid_from_payload(payload)? else {
                    return Err(DecodeError::projection());
                };
                if let Some(index) = find_uid_index(index_plan, desired_range, uid)? {
                    mark_match(matched, desired_range, index)?;
                    let state = &desired.base_hidden_states[index];
                    let payload_len = measure_row_state(payload, state)?;
                    emit_length_delimited_header(EXTENT_BASE_STATES_FIELD, payload_len, output);
                    emit_row_state(payload, state, output)?;
                }
            },
            EXTENT_NEEDS_FILTER_UPDATE_FIELD => {
                needs_filter_seen = true;
                emit_optional_bool_field_preserving(
                    source,
                    field,
                    EXTENT_NEEDS_FILTER_UPDATE_FIELD,
                    desired.needs_to_update_filter_set_for_import,
                    output,
                )?;
            },
            EXTENT_FILTER_SET_FIELD => {
                filter_set_seen = true;
                if let Some(reference) = desired.filter_set {
                    if source_reference_matches(source, field, reference)? {
                        output.extend_from_slice(&source[field.start..field.end]);
                    } else {
                        emit_reference_field_preserving(
                            source,
                            field,
                            EXTENT_FILTER_SET_FIELD,
                            reference,
                            output,
                        )?;
                    }
                }
            },
            _ => output.extend_from_slice(&source[field.start..field.end]),
        }
        Ok(())
    })?;
    if !uid_seen {
        return Err(DecodeError::missing(
            "TST.HiddenStateExtentArchive.hidden_state_extent_uid",
        ));
    }
    if !direction_seen {
        return Err(DecodeError::missing(
            "TST.HiddenStateExtentArchive.row_or_column_direction",
        ));
    }
    for (index, state) in desired.base_hidden_states.iter().enumerate() {
        if !is_matched(matched, desired_range, index)? {
            let payload_len = canonical_row_state_len(state)?;
            emit_length_delimited_header(EXTENT_BASE_STATES_FIELD, payload_len, output);
            emit_row_state_canonical(state, output)?;
        }
    }
    if !needs_filter_seen && let Some(value) = desired.needs_to_update_filter_set_for_import {
        emit_optional_bool(EXTENT_NEEDS_FILTER_UPDATE_FIELD, Some(value), output);
    }
    if !filter_set_seen && let Some(reference) = desired.filter_set {
        let payload_len = canonical_reference_len(reference);
        emit_length_delimited_header(EXTENT_FILTER_SET_FIELD, payload_len, output);
        emit_reference_canonical(reference, output);
    }
    Ok(())
}

fn emit_row_state(
    source: &[u8],
    desired: &RowOrColumnStateSnapshot,
    output: &mut Vec<u8>,
) -> Result<(), DecodeError> {
    let mut uid_seen = false;
    let mut replaced = [false; 3];
    for_each_field_for_emit(source, |field| {
        match field.number {
            ROW_STATE_UID_FIELD => {
                if uid_seen {
                    return Err(DecodeError::duplicate(
                        "TST.HiddenStateExtentArchive.RowOrColumnState.row_or_column_uid",
                    ));
                }
                uid_seen = true;
                if source_uuid_matches(source, field, desired.row_or_column_uid)? {
                    output.extend_from_slice(&source[field.start..field.end]);
                } else {
                    emit_uuid_field_preserving(
                        source,
                        field,
                        ROW_STATE_UID_FIELD,
                        desired.row_or_column_uid,
                        output,
                    )?;
                }
            },
            ROW_STATE_USER_HIDDEN_FIELD => {
                replaced[0] = true;
                emit_optional_bool_field_preserving(
                    source,
                    field,
                    ROW_STATE_USER_HIDDEN_FIELD,
                    desired.user_hidden,
                    output,
                )?;
            },
            ROW_STATE_FILTERED_FIELD => {
                replaced[1] = true;
                emit_optional_bool_field_preserving(
                    source,
                    field,
                    ROW_STATE_FILTERED_FIELD,
                    desired.filtered,
                    output,
                )?;
            },
            ROW_STATE_PIVOT_HIDDEN_FIELD => {
                replaced[2] = true;
                emit_optional_bool_field_preserving(
                    source,
                    field,
                    ROW_STATE_PIVOT_HIDDEN_FIELD,
                    desired.pivot_hidden,
                    output,
                )?;
            },
            _ => output.extend_from_slice(&source[field.start..field.end]),
        }
        Ok(())
    })?;
    if !uid_seen {
        return Err(DecodeError::missing(
            "TST.HiddenStateExtentArchive.RowOrColumnState.row_or_column_uid",
        ));
    }
    if !replaced[0] {
        emit_optional_bool(ROW_STATE_USER_HIDDEN_FIELD, desired.user_hidden, output);
    }
    if !replaced[1] {
        emit_optional_bool(ROW_STATE_FILTERED_FIELD, desired.filtered, output);
    }
    if !replaced[2] {
        emit_optional_bool(ROW_STATE_PIVOT_HIDDEN_FIELD, desired.pivot_hidden, output);
    }
    Ok(())
}

#[derive(Debug, Clone, Copy)]
struct UidIndexEntry {
    uid: UuidSnapshot,
    index: usize,
}

fn uid_key(uid: UuidSnapshot) -> (u64, u64) {
    (uid.lower, uid.upper)
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct IndexRange {
    start: usize,
    len: usize,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct ExtentIndexEntry {
    uid: UuidSnapshot,
    range: IndexRange,
}

/// One sorted plan is prepared once and reused by both the sizing and emit
/// passes.  The flat storage keeps allocation count independent of graph
/// width: owner entries, every extent's row entries, and extent metadata are
/// each held in one bounded vector.
#[derive(Debug, Clone, Default)]
struct RewriteIndexPlan {
    entries: Vec<UidIndexEntry>,
    extents: Vec<ExtentIndexEntry>,
    owner_range: Option<IndexRange>,
}

impl RewriteIndexPlan {
    fn extent_range(&self, uid: UuidSnapshot) -> Option<IndexRange> {
        self.extents
            .binary_search_by_key(&uid_key(uid), |entry| uid_key(entry.uid))
            .ok()
            .map(|index| self.extents[index].range)
    }

    fn entries(&self, range: IndexRange) -> Result<&[UidIndexEntry], DecodeError> {
        let end = range
            .start
            .checked_add(range.len)
            .ok_or_else(DecodeError::projection)?;
        self.entries
            .get(range.start..end)
            .ok_or_else(DecodeError::projection)
    }
}

fn append_hidden_state_index(
    entries: &mut Vec<UidIndexEntry>,
    states: &[HiddenStatesSnapshot],
) -> Result<IndexRange, DecodeError> {
    let start = entries.len();
    entries.extend(
        states
            .iter()
            .enumerate()
            .map(|(index, state)| UidIndexEntry {
                uid: state.hidden_states_uid,
                index,
            }),
    );
    let range = IndexRange {
        start,
        len: states.len(),
    };
    let slice = entries
        .get_mut(
            start
                ..start
                    .checked_add(range.len)
                    .ok_or_else(DecodeError::projection)?,
        )
        .ok_or_else(DecodeError::projection)?;
    slice.sort_unstable_by_key(|entry| uid_key(entry.uid));
    if slice
        .windows(2)
        .any(|window| window[0].uid == window[1].uid)
    {
        return Err(DecodeError::invalid("duplicate hidden-state UUID"));
    }
    Ok(range)
}

fn append_row_state_index(
    entries: &mut Vec<UidIndexEntry>,
    states: &[RowOrColumnStateSnapshot],
) -> Result<IndexRange, DecodeError> {
    let start = entries.len();
    entries.extend(
        states
            .iter()
            .enumerate()
            .map(|(index, state)| UidIndexEntry {
                uid: state.row_or_column_uid,
                index,
            }),
    );
    let range = IndexRange {
        start,
        len: states.len(),
    };
    let slice = entries
        .get_mut(
            start
                ..start
                    .checked_add(range.len)
                    .ok_or_else(DecodeError::projection)?,
        )
        .ok_or_else(DecodeError::projection)?;
    slice.sort_unstable_by_key(|entry| uid_key(entry.uid));
    if slice
        .windows(2)
        .any(|window| window[0].uid == window[1].uid)
    {
        return Err(DecodeError::invalid("duplicate row/column UUID"));
    }
    Ok(range)
}

fn index_counts(desired: &RewriteDesired) -> Result<(usize, usize), DecodeError> {
    let mut entries = 0usize;
    let mut extents = 0usize;
    let mut account_owner = |owner: &HiddenStatesOwnerSnapshot| -> Result<(), DecodeError> {
        entries = entries
            .checked_add(owner.hidden_states.len())
            .ok_or_else(DecodeError::projection)?;
        extents = extents
            .checked_add(
                owner
                    .hidden_states
                    .len()
                    .checked_mul(2)
                    .ok_or_else(DecodeError::projection)?,
            )
            .ok_or_else(DecodeError::projection)?;
        for hidden_states in &owner.hidden_states {
            entries = entries
                .checked_add(
                    hidden_states
                        .column_hidden_state_extent
                        .base_hidden_states
                        .len(),
                )
                .and_then(|value| {
                    value.checked_add(
                        hidden_states
                            .row_hidden_state_extent
                            .base_hidden_states
                            .len(),
                    )
                })
                .ok_or_else(DecodeError::projection)?;
        }
        Ok(())
    };
    match desired {
        RewriteDesired::TableInfo(_) | RewriteDesired::RowOrColumnState(_) => {},
        RewriteDesired::TableModel(value) => {
            if let Some(owner) = value.hidden_states_owner.as_ref() {
                account_owner(owner)?;
            }
        },
        RewriteDesired::HiddenStatesOwner(owner) => account_owner(owner)?,
        RewriteDesired::HiddenStateExtent(extent) => {
            entries = extent.base_hidden_states.len();
            extents = 1;
        },
    }
    Ok((entries, extents))
}

fn build_rewrite_index_plan(desired: &RewriteDesired) -> Result<RewriteIndexPlan, DecodeError> {
    let (entry_count, extent_count) = index_counts(desired)?;
    let entry_bytes = entry_count
        .checked_mul(size_of::<UidIndexEntry>())
        .ok_or_else(DecodeError::projection)?;
    let extent_bytes = extent_count
        .checked_mul(size_of::<ExtentIndexEntry>())
        .ok_or_else(DecodeError::projection)?;
    let mut plan = RewriteIndexPlan::default();
    if entry_count != 0 {
        plan.entries
            .try_reserve_exact(entry_count)
            .map_err(|_| DecodeError::allocation(entry_bytes))?;
    }
    if extent_count != 0 {
        plan.extents
            .try_reserve_exact(extent_count)
            .map_err(|_| DecodeError::allocation(extent_bytes))?;
    }
    let mut add_owner = |owner: &HiddenStatesOwnerSnapshot| -> Result<(), DecodeError> {
        plan.owner_range = Some(append_hidden_state_index(
            &mut plan.entries,
            &owner.hidden_states,
        )?);
        for hidden_states in &owner.hidden_states {
            let column = append_row_state_index(
                &mut plan.entries,
                &hidden_states.column_hidden_state_extent.base_hidden_states,
            )?;
            plan.extents.push(ExtentIndexEntry {
                uid: hidden_states
                    .column_hidden_state_extent
                    .hidden_state_extent_uid,
                range: column,
            });
            let row = append_row_state_index(
                &mut plan.entries,
                &hidden_states.row_hidden_state_extent.base_hidden_states,
            )?;
            plan.extents.push(ExtentIndexEntry {
                uid: hidden_states
                    .row_hidden_state_extent
                    .hidden_state_extent_uid,
                range: row,
            });
        }
        Ok(())
    };
    match desired {
        RewriteDesired::TableModel(value) => {
            if let Some(owner) = value.hidden_states_owner.as_ref() {
                add_owner(owner)?;
            }
        },
        RewriteDesired::HiddenStatesOwner(owner) => add_owner(owner)?,
        RewriteDesired::HiddenStateExtent(extent) => {
            let range = append_row_state_index(&mut plan.entries, &extent.base_hidden_states)?;
            plan.extents.push(ExtentIndexEntry {
                uid: extent.hidden_state_extent_uid,
                range,
            });
        },
        RewriteDesired::TableInfo(_) | RewriteDesired::RowOrColumnState(_) => {},
    }
    plan.extents
        .sort_unstable_by_key(|entry| uid_key(entry.uid));
    if plan
        .extents
        .windows(2)
        .any(|window| window[0].uid == window[1].uid)
    {
        return Err(DecodeError::invalid("duplicate hidden-state extent UUID"));
    }
    Ok(plan)
}

fn new_match_marks(plan: &RewriteIndexPlan) -> Result<Vec<u8>, DecodeError> {
    if plan.entries.is_empty() {
        return Ok(Vec::new());
    }
    let mut marks = Vec::new();
    marks
        .try_reserve_exact(plan.entries.len())
        .map_err(|_| DecodeError::allocation(plan.entries.len()))?;
    marks.resize(plan.entries.len(), 0);
    Ok(marks)
}

fn find_uid_index(
    plan: &RewriteIndexPlan,
    range: IndexRange,
    expected: UuidSnapshot,
) -> Result<Option<usize>, DecodeError> {
    let entries = plan.entries(range)?;
    Ok(entries
        .binary_search_by_key(&uid_key(expected), |entry| uid_key(entry.uid))
        .ok()
        .map(|position| entries[position].index))
}

fn mark_match(matched: &mut [u8], range: IndexRange, index: usize) -> Result<(), DecodeError> {
    let position = range
        .start
        .checked_add(index)
        .ok_or_else(DecodeError::projection)?;
    let mark = matched
        .get_mut(position)
        .ok_or_else(DecodeError::projection)?;
    // The emit pass measures each nested payload immediately before writing
    // it, so a desired record can be observed once by sizing and once by
    // emission. Source preflight has already rejected duplicate UUIDs; keep
    // this phase marker idempotent while retaining a single bitmap.
    *mark = 1;
    Ok(())
}

fn is_matched(matched: &[u8], range: IndexRange, index: usize) -> Result<bool, DecodeError> {
    let position = range
        .start
        .checked_add(index)
        .ok_or_else(DecodeError::projection)?;
    matched
        .get(position)
        .map(|mark| *mark != 0)
        .ok_or_else(DecodeError::projection)
}

fn parse_uuid_for_emit(source: &[u8]) -> Result<UuidSnapshot, DecodeError> {
    let mut offset = 0;
    let mut lower = None;
    let mut upper = None;
    while offset < source.len() {
        let field = parse_field_for_emit(source, &mut offset)?;
        match field.number {
            UUID_LOWER_FIELD => {
                field_wire(field, 0)?;
                if lower.replace(field_varint(source, field)?).is_some() {
                    return Err(DecodeError::duplicate("TSP.UUID.lower"));
                }
            },
            UUID_UPPER_FIELD => {
                field_wire(field, 0)?;
                if upper.replace(field_varint(source, field)?).is_some() {
                    return Err(DecodeError::duplicate("TSP.UUID.upper"));
                }
            },
            _ => {},
        }
    }
    Ok(UuidSnapshot::new(
        lower.ok_or_else(|| DecodeError::missing("TSP.UUID.lower"))?,
        upper.ok_or_else(|| DecodeError::missing("TSP.UUID.upper"))?,
    ))
}

fn parse_reference_for_emit(source: &[u8]) -> Result<ReferenceSnapshot, DecodeError> {
    let mut offset = 0;
    let mut identifier = None;
    let mut deprecated_type = None;
    let mut deprecated_is_external = None;
    while offset < source.len() {
        let field = parse_field_for_emit(source, &mut offset)?;
        match field.number {
            REFERENCE_IDENTIFIER_FIELD => {
                field_wire(field, 0)?;
                if identifier.replace(field_varint(source, field)?).is_some() {
                    return Err(DecodeError::duplicate("TSP.Reference.identifier"));
                }
            },
            REFERENCE_DEPRECATED_TYPE_FIELD => {
                field_wire(field, 0)?;
                if deprecated_type
                    .replace(decode_int32_checked(field_varint(source, field)?)?)
                    .is_some()
                {
                    return Err(DecodeError::duplicate("TSP.Reference.deprecated_type"));
                }
            },
            REFERENCE_DEPRECATED_EXTERNAL_FIELD => {
                field_wire(field, 0)?;
                if deprecated_is_external
                    .replace(require_bool(field_varint(source, field)?)?)
                    .is_some()
                {
                    return Err(DecodeError::duplicate(
                        "TSP.Reference.deprecated_is_external",
                    ));
                }
            },
            _ => {},
        }
    }
    let identifier = NonZeroU64::new(
        identifier.ok_or_else(|| DecodeError::missing("TSP.Reference.identifier"))?,
    )
    .ok_or_else(|| DecodeError::invalid("reference identifier is zero"))?;
    Ok(ReferenceSnapshot {
        identifier,
        deprecated_type,
        deprecated_is_external,
    })
}

fn source_uuid_matches(
    source: &[u8],
    field: FieldSpan,
    expected: UuidSnapshot,
) -> Result<bool, DecodeError> {
    Ok(parse_uuid_for_emit(field_bytes(source, field)?)? == expected)
}

fn source_reference_matches(
    source: &[u8],
    field: FieldSpan,
    expected: ReferenceSnapshot,
) -> Result<bool, DecodeError> {
    Ok(parse_reference_for_emit(field_bytes(source, field)?)? == expected)
}

fn source_varint_matches(
    source: &[u8],
    field: FieldSpan,
    expected: u64,
) -> Result<bool, DecodeError> {
    Ok(field_varint(source, field)? == expected)
}

fn measure_uuid_payload_rewrite(
    source: &[u8],
    desired: UuidSnapshot,
) -> Result<usize, DecodeError> {
    let mut total = 0usize;
    let mut lower_seen = false;
    let mut upper_seen = false;
    for_each_field_for_emit(source, |field| {
        match field.number {
            UUID_LOWER_FIELD => {
                if lower_seen {
                    return Err(DecodeError::duplicate("TSP.UUID.lower"));
                }
                lower_seen = true;
                if source_varint_matches(source, field, desired.lower)? {
                    add_len(&mut total, field.end - field.start)?;
                } else {
                    add_len(
                        &mut total,
                        varint_field_len(UUID_LOWER_FIELD, desired.lower),
                    )?;
                }
            },
            UUID_UPPER_FIELD => {
                if upper_seen {
                    return Err(DecodeError::duplicate("TSP.UUID.upper"));
                }
                upper_seen = true;
                if source_varint_matches(source, field, desired.upper)? {
                    add_len(&mut total, field.end - field.start)?;
                } else {
                    add_len(
                        &mut total,
                        varint_field_len(UUID_UPPER_FIELD, desired.upper),
                    )?;
                }
            },
            _ => add_len(&mut total, field.end - field.start)?,
        }
        Ok(())
    })?;
    if !lower_seen {
        return Err(DecodeError::missing("TSP.UUID.lower"));
    }
    if !upper_seen {
        return Err(DecodeError::missing("TSP.UUID.upper"));
    }
    Ok(total)
}

fn emit_uuid_payload_preserving(
    source: &[u8],
    desired: UuidSnapshot,
    output: &mut Vec<u8>,
) -> Result<(), DecodeError> {
    let size = measure_uuid_payload_rewrite(source, desired)?;
    let start = output.len();
    let mut lower_seen = false;
    let mut upper_seen = false;
    for_each_field_for_emit(source, |field| {
        match field.number {
            UUID_LOWER_FIELD => {
                lower_seen = true;
                if source_varint_matches(source, field, desired.lower)? {
                    output.extend_from_slice(&source[field.start..field.end]);
                } else {
                    emit_varint_field(UUID_LOWER_FIELD, desired.lower, output);
                }
            },
            UUID_UPPER_FIELD => {
                upper_seen = true;
                if source_varint_matches(source, field, desired.upper)? {
                    output.extend_from_slice(&source[field.start..field.end]);
                } else {
                    emit_varint_field(UUID_UPPER_FIELD, desired.upper, output);
                }
            },
            _ => output.extend_from_slice(&source[field.start..field.end]),
        }
        Ok(())
    })?;
    let written = output
        .len()
        .checked_sub(start)
        .ok_or_else(DecodeError::projection)?;
    if !lower_seen || !upper_seen || written != size {
        return Err(DecodeError::projection());
    }
    Ok(())
}

fn measure_reference_payload_rewrite(
    source: &[u8],
    desired: ReferenceSnapshot,
) -> Result<usize, DecodeError> {
    let mut total = 0usize;
    let mut identifier_seen = false;
    let mut deprecated_type_seen = false;
    let mut external_seen = false;
    for_each_field_for_emit(source, |field| {
        match field.number {
            REFERENCE_IDENTIFIER_FIELD => {
                if identifier_seen {
                    return Err(DecodeError::duplicate("TSP.Reference.identifier"));
                }
                identifier_seen = true;
                if source_varint_matches(source, field, desired.identifier.get())? {
                    add_len(&mut total, field.end - field.start)?;
                } else {
                    add_len(
                        &mut total,
                        varint_field_len(REFERENCE_IDENTIFIER_FIELD, desired.identifier.get()),
                    )?;
                }
            },
            REFERENCE_DEPRECATED_TYPE_FIELD => {
                if deprecated_type_seen {
                    return Err(DecodeError::duplicate("TSP.Reference.deprecated_type"));
                }
                deprecated_type_seen = true;
                if let Some(value) = desired.deprecated_type {
                    let encoded = value as i64 as u64;
                    if source_varint_matches(source, field, encoded)? {
                        add_len(&mut total, field.end - field.start)?;
                    } else {
                        add_len(
                            &mut total,
                            varint_field_len(REFERENCE_DEPRECATED_TYPE_FIELD, encoded),
                        )?;
                    }
                }
            },
            REFERENCE_DEPRECATED_EXTERNAL_FIELD => {
                if external_seen {
                    return Err(DecodeError::duplicate(
                        "TSP.Reference.deprecated_is_external",
                    ));
                }
                external_seen = true;
                if let Some(value) = desired.deprecated_is_external {
                    if source_varint_matches(source, field, u64::from(value))? {
                        add_len(&mut total, field.end - field.start)?;
                    } else {
                        add_len(
                            &mut total,
                            varint_field_len(REFERENCE_DEPRECATED_EXTERNAL_FIELD, u64::from(value)),
                        )?;
                    }
                }
            },
            _ => add_len(&mut total, field.end - field.start)?,
        }
        Ok(())
    })?;
    if !identifier_seen {
        return Err(DecodeError::missing("TSP.Reference.identifier"));
    }
    if !deprecated_type_seen {
        if let Some(value) = desired.deprecated_type {
            add_len(
                &mut total,
                varint_field_len(REFERENCE_DEPRECATED_TYPE_FIELD, value as i64 as u64),
            )?;
        }
    }
    if !external_seen {
        if let Some(value) = desired.deprecated_is_external {
            add_len(
                &mut total,
                varint_field_len(REFERENCE_DEPRECATED_EXTERNAL_FIELD, u64::from(value)),
            )?;
        }
    }
    Ok(total)
}

fn emit_reference_payload_preserving(
    source: &[u8],
    desired: ReferenceSnapshot,
    output: &mut Vec<u8>,
) -> Result<(), DecodeError> {
    let size = measure_reference_payload_rewrite(source, desired)?;
    let start = output.len();
    let mut identifier_seen = false;
    let mut deprecated_type_seen = false;
    let mut external_seen = false;
    for_each_field_for_emit(source, |field| {
        match field.number {
            REFERENCE_IDENTIFIER_FIELD => {
                identifier_seen = true;
                if source_varint_matches(source, field, desired.identifier.get())? {
                    output.extend_from_slice(&source[field.start..field.end]);
                } else {
                    emit_varint_field(REFERENCE_IDENTIFIER_FIELD, desired.identifier.get(), output);
                }
            },
            REFERENCE_DEPRECATED_TYPE_FIELD => {
                deprecated_type_seen = true;
                if let Some(value) = desired.deprecated_type {
                    let encoded = value as i64 as u64;
                    if source_varint_matches(source, field, encoded)? {
                        output.extend_from_slice(&source[field.start..field.end]);
                    } else {
                        emit_varint_field(REFERENCE_DEPRECATED_TYPE_FIELD, encoded, output);
                    }
                }
            },
            REFERENCE_DEPRECATED_EXTERNAL_FIELD => {
                external_seen = true;
                if let Some(value) = desired.deprecated_is_external {
                    if source_varint_matches(source, field, u64::from(value))? {
                        output.extend_from_slice(&source[field.start..field.end]);
                    } else {
                        emit_varint_field(
                            REFERENCE_DEPRECATED_EXTERNAL_FIELD,
                            u64::from(value),
                            output,
                        );
                    }
                }
            },
            _ => output.extend_from_slice(&source[field.start..field.end]),
        }
        Ok(())
    })?;
    if !identifier_seen {
        return Err(DecodeError::missing("TSP.Reference.identifier"));
    }
    if !deprecated_type_seen {
        if let Some(value) = desired.deprecated_type {
            emit_varint_field(REFERENCE_DEPRECATED_TYPE_FIELD, value as i64 as u64, output);
        }
    }
    if !external_seen {
        if let Some(value) = desired.deprecated_is_external {
            emit_varint_field(
                REFERENCE_DEPRECATED_EXTERNAL_FIELD,
                u64::from(value),
                output,
            );
        }
    }
    let written = output
        .len()
        .checked_sub(start)
        .ok_or_else(DecodeError::projection)?;
    if written != size {
        return Err(DecodeError::projection());
    }
    Ok(())
}

fn for_each_field_for_emit(
    source: &[u8],
    mut callback: impl FnMut(FieldSpan) -> Result<(), DecodeError>,
) -> Result<(), DecodeError> {
    let mut offset = 0;
    while offset < source.len() {
        let field = parse_field_for_emit(source, &mut offset)?;
        callback(field)?;
    }
    Ok(())
}

fn parse_field_for_emit(source: &[u8], offset: &mut usize) -> Result<FieldSpan, DecodeError> {
    let start = *offset;
    let key = read_varint(source, offset)?;
    if !key.canonical {
        return Err(DecodeError::noncanonical("protobuf field key"));
    }
    let raw = u32::try_from(key.value).map_err(|_| buffa::DecodeError::InvalidFieldNumber)?;
    let number = raw >> 3;
    let wire = u8::try_from(raw & 7).map_err(|_| buffa::DecodeError::InvalidWireType(raw & 7))?;
    if number == 0 || number > buffa::encoding::MAX_FIELD_NUMBER || wire == 4 {
        return Err(buffa::DecodeError::InvalidFieldNumber.into());
    }
    // Keep the payload span separate from the length varint.  The source
    // bytes are later fed to nested strict walkers, so including the length
    // prefix here would make a valid nested message look malformed.
    let mut value_start = *offset;
    let (value_end, value_canonical, length_canonical) = match wire {
        0 => {
            let value = read_varint(source, offset)?;
            (*offset, value.canonical, true)
        },
        1 => {
            let end = offset.checked_add(8).ok_or_else(DecodeError::projection)?;
            if end > source.len() {
                return Err(buffa::DecodeError::UnexpectedEof.into());
            }
            *offset = end;
            (end, true, true)
        },
        2 => {
            let len = read_varint(source, offset)?;
            let n = usize::try_from(len.value).map_err(|_| buffa::DecodeError::MessageTooLarge)?;
            value_start = *offset;
            let end = offset.checked_add(n).ok_or_else(DecodeError::projection)?;
            if end > source.len() {
                return Err(buffa::DecodeError::UnexpectedEof.into());
            }
            *offset = end;
            (end, true, len.canonical)
        },
        3 => {
            let end = skip_group_for_emit(source, offset, number)?;
            (end, true, true)
        },
        5 => {
            let end = offset.checked_add(4).ok_or_else(DecodeError::projection)?;
            if end > source.len() {
                return Err(buffa::DecodeError::UnexpectedEof.into());
            }
            *offset = end;
            (end, true, true)
        },
        _ => return Err(buffa::DecodeError::InvalidWireType(raw & 7).into()),
    };
    Ok(FieldSpan {
        number,
        wire,
        start,
        end: *offset,
        value_start,
        value_end,
        value_canonical,
        length_canonical,
    })
}

fn skip_group_for_emit(
    source: &[u8],
    offset: &mut usize,
    opening: u32,
) -> Result<usize, DecodeError> {
    loop {
        let key = read_varint(source, offset)?;
        if !key.canonical {
            return Err(DecodeError::noncanonical("protobuf group field key"));
        }
        let raw = u32::try_from(key.value).map_err(|_| buffa::DecodeError::InvalidFieldNumber)?;
        let number = raw >> 3;
        let wire = raw & 7;
        if number == 0 || number > buffa::encoding::MAX_FIELD_NUMBER {
            return Err(buffa::DecodeError::InvalidFieldNumber.into());
        }
        match wire {
            4 if number == opening => return Ok(*offset),
            4 => return Err(buffa::DecodeError::InvalidEndGroup(number).into()),
            3 => {
                skip_group_for_emit(source, offset, number)?;
            },
            0 => {
                let _ = read_varint(source, offset)?;
            },
            1 => {
                let end = offset.checked_add(8).ok_or_else(DecodeError::projection)?;
                if end > source.len() {
                    return Err(buffa::DecodeError::UnexpectedEof.into());
                }
                *offset = end;
            },
            2 => {
                let len = read_varint(source, offset)?;
                if !len.canonical {
                    return Err(DecodeError::noncanonical("protobuf group length"));
                }
                let n =
                    usize::try_from(len.value).map_err(|_| buffa::DecodeError::MessageTooLarge)?;
                let end = offset.checked_add(n).ok_or_else(DecodeError::projection)?;
                if end > source.len() {
                    return Err(buffa::DecodeError::UnexpectedEof.into());
                }
                *offset = end;
            },
            5 => {
                let end = offset.checked_add(4).ok_or_else(DecodeError::projection)?;
                if end > source.len() {
                    return Err(buffa::DecodeError::UnexpectedEof.into());
                }
                *offset = end;
            },
            _ => return Err(buffa::DecodeError::InvalidWireType(wire).into()),
        }
    }
}

fn source_field_payload(source: &[u8], field: FieldSpan) -> Result<&[u8], DecodeError> {
    field_bytes(source, field)
}

fn state_uid_from_payload(source: &[u8]) -> Result<Option<UuidSnapshot>, DecodeError> {
    let mut offset = 0;
    let mut uid = None;
    while offset < source.len() {
        let field = parse_field_for_emit(source, &mut offset)?;
        if field.number == STATE_UID_FIELD {
            let bytes = field_bytes(source, field)?;
            let mut inner = 0;
            let mut lower = None;
            let mut upper = None;
            while inner < bytes.len() {
                let child = parse_field_for_emit(bytes, &mut inner)?;
                match child.number {
                    UUID_LOWER_FIELD => lower = Some(field_varint(bytes, child)?),
                    UUID_UPPER_FIELD => upper = Some(field_varint(bytes, child)?),
                    _ => {},
                }
            }
            if let (Some(lower), Some(upper)) = (lower, upper) {
                uid = Some(UuidSnapshot::new(lower, upper));
            }
        }
    }
    Ok(uid)
}

fn emit_hidden_states_from_source(
    source: &[u8],
    desired: &HiddenStatesSnapshot,
    expected_size: usize,
    index_plan: &RewriteIndexPlan,
    matched: &mut [u8],
    output: &mut Vec<u8>,
) -> Result<(), DecodeError> {
    let start = output.len();
    let mut uid_seen = false;
    let mut column_seen = false;
    let mut row_seen = false;
    for_each_field_for_emit(source, |field| {
        match field.number {
            STATE_UID_FIELD => {
                if uid_seen {
                    return Err(DecodeError::duplicate(
                        "TST.HiddenStatesArchive.hidden_states_uid",
                    ));
                }
                uid_seen = true;
                if source_uuid_matches(source, field, desired.hidden_states_uid)? {
                    output.extend_from_slice(&source[field.start..field.end]);
                } else {
                    emit_uuid_field_preserving(
                        source,
                        field,
                        STATE_UID_FIELD,
                        desired.hidden_states_uid,
                        output,
                    )?;
                }
            },
            STATE_COLUMN_EXTENT_FIELD => {
                if column_seen {
                    return Err(DecodeError::duplicate(
                        "TST.HiddenStatesArchive.column_hidden_state_extent",
                    ));
                }
                column_seen = true;
                let bytes = source_field_payload(source, field)?;
                let payload_len = measure_extent(
                    bytes,
                    desired.column_hidden_state_extent(),
                    index_plan,
                    matched,
                )?;
                emit_length_delimited_header(STATE_COLUMN_EXTENT_FIELD, payload_len, output);
                emit_extent(
                    bytes,
                    desired.column_hidden_state_extent(),
                    index_plan,
                    matched,
                    output,
                )?;
            },
            STATE_ROW_EXTENT_FIELD => {
                if row_seen {
                    return Err(DecodeError::duplicate(
                        "TST.HiddenStatesArchive.row_hidden_state_extent",
                    ));
                }
                row_seen = true;
                let bytes = source_field_payload(source, field)?;
                let payload_len = measure_extent(
                    bytes,
                    desired.row_hidden_state_extent(),
                    index_plan,
                    matched,
                )?;
                emit_length_delimited_header(STATE_ROW_EXTENT_FIELD, payload_len, output);
                emit_extent(
                    bytes,
                    desired.row_hidden_state_extent(),
                    index_plan,
                    matched,
                    output,
                )?;
            },
            _ => output.extend_from_slice(&source[field.start..field.end]),
        }
        Ok(())
    })?;
    if !uid_seen {
        return Err(DecodeError::missing(
            "TST.HiddenStatesArchive.hidden_states_uid",
        ));
    }
    if !column_seen || !row_seen {
        return Err(DecodeError::missing("TST.HiddenStatesArchive.extent"));
    }
    let written = output
        .len()
        .checked_sub(start)
        .ok_or_else(DecodeError::projection)?;
    if written != expected_size {
        return Err(DecodeError::projection());
    }
    Ok(())
}

fn canonical_uuid(value: UuidSnapshot) -> Result<Vec<u8>, DecodeError> {
    let mut output = Vec::new();
    let size = canonical_uuid_len(value);
    output
        .try_reserve_exact(size)
        .map_err(|_| DecodeError::allocation(size))?;
    emit_uuid_canonical(value, &mut output);
    Ok(output)
}

fn canonical_reference(value: ReferenceSnapshot) -> Result<Vec<u8>, DecodeError> {
    let mut output = Vec::new();
    let size = canonical_reference_len(value);
    output
        .try_reserve_exact(size)
        .map_err(|_| DecodeError::allocation(size))?;
    emit_reference_canonical(value, &mut output);
    Ok(output)
}

#[cfg(test)]
fn canonical_owner(value: &HiddenStatesOwnerSnapshot) -> Result<Vec<u8>, DecodeError> {
    let mut output = Vec::new();
    let size = canonical_owner_len(value)?;
    output
        .try_reserve_exact(size)
        .map_err(|_| DecodeError::allocation(size))?;
    emit_owner_canonical(value, &mut output)?;
    Ok(output)
}

fn emit_u32_field_preserving(
    source: &[u8],
    field: FieldSpan,
    number: u32,
    value: u32,
    output: &mut Vec<u8>,
) -> Result<(), DecodeError> {
    if source_varint_matches(source, field, u64::from(value))? {
        output.extend_from_slice(&source[field.start..field.end]);
    } else {
        emit_varint_field(number, u64::from(value), output);
    }
    Ok(())
}

fn emit_optional_u32_field_preserving(
    source: &[u8],
    field: FieldSpan,
    number: u32,
    value: Option<u32>,
    output: &mut Vec<u8>,
) -> Result<(), DecodeError> {
    if let Some(value) = value {
        emit_u32_field_preserving(source, field, number, value, output)?;
    }
    Ok(())
}

fn emit_optional_bool_field_preserving(
    source: &[u8],
    field: FieldSpan,
    number: u32,
    value: Option<bool>,
    output: &mut Vec<u8>,
) -> Result<(), DecodeError> {
    if let Some(value) = value {
        if source_varint_matches(source, field, u64::from(value))? {
            output.extend_from_slice(&source[field.start..field.end]);
        } else {
            emit_optional_bool(number, Some(value), output);
        }
    }
    Ok(())
}

fn emit_optional_reference_field_preserving(
    source: &[u8],
    field: FieldSpan,
    number: u32,
    value: Option<ReferenceSnapshot>,
    output: &mut Vec<u8>,
) -> Result<(), DecodeError> {
    if let Some(value) = value {
        if source_reference_matches(source, field, value)? {
            output.extend_from_slice(&source[field.start..field.end]);
        } else {
            emit_reference_field_preserving(source, field, number, value, output)?;
        }
    }
    Ok(())
}

fn emit_uuid_field_preserving(
    source: &[u8],
    field: FieldSpan,
    number: u32,
    desired: UuidSnapshot,
    output: &mut Vec<u8>,
) -> Result<(), DecodeError> {
    if source_uuid_matches(source, field, desired)? {
        output.extend_from_slice(&source[field.start..field.end]);
    } else {
        let payload = field_bytes(source, field)?;
        let payload_len = measure_uuid_payload_rewrite(payload, desired)?;
        emit_length_delimited_header(number, payload_len, output);
        emit_uuid_payload_preserving(payload, desired, output)?;
    }
    Ok(())
}

fn emit_reference_field_preserving(
    source: &[u8],
    field: FieldSpan,
    number: u32,
    desired: ReferenceSnapshot,
    output: &mut Vec<u8>,
) -> Result<(), DecodeError> {
    if source_reference_matches(source, field, desired)? {
        output.extend_from_slice(&source[field.start..field.end]);
    } else {
        let payload = field_bytes(source, field)?;
        let payload_len = measure_reference_payload_rewrite(payload, desired)?;
        emit_length_delimited_header(number, payload_len, output);
        emit_reference_payload_preserving(payload, desired, output)?;
    }
    Ok(())
}

fn emit_varint_field(number: u32, value: u64, output: &mut Vec<u8>) {
    encode_varint((u64::from(number)) << 3, output);
    encode_varint(value, output);
}
fn emit_optional_u32(number: u32, value: Option<u32>, output: &mut Vec<u8>) {
    if let Some(value) = value {
        emit_varint_field(number, u64::from(value), output);
    }
}
fn emit_optional_bool(number: u32, value: Option<bool>, output: &mut Vec<u8>) {
    if let Some(value) = value {
        emit_varint_field(number, u64::from(value), output);
    }
}
fn emit_optional_reference(
    number: u32,
    value: Option<ReferenceSnapshot>,
    output: &mut Vec<u8>,
) -> Result<(), DecodeError> {
    if let Some(value) = value {
        emit_length_delimited(number, &canonical_reference(value)?, output);
    }
    Ok(())
}
fn emit_length_delimited(number: u32, payload: &[u8], output: &mut Vec<u8>) {
    emit_length_delimited_header(number, payload.len(), output);
    output.extend_from_slice(payload);
}
fn emit_length_delimited_header(number: u32, payload_len: usize, output: &mut Vec<u8>) {
    encode_varint((u64::from(number) << 3) | 2, output);
    encode_varint(payload_len as u64, output);
}
fn encode_varint(mut value: u64, output: &mut Vec<u8>) {
    while value >= 128 {
        output.push(((value as u8) & 0x7f) | 0x80);
        value >>= 7;
    }
    output.push(value as u8);
}

#[cfg(test)]
mod tests {
    use super::*;

    fn options(source: &[u8]) -> DecodeOptions {
        DecodeOptions::new(
            source.len().max(1),
            source.len().saturating_mul(2).max(1),
            source.len().saturating_mul(32).max(64),
            source.len().saturating_mul(128).max(128),
            16,
            128,
        )
        .with_max_allocations(4_096)
        .with_max_retained_bytes(source.len().saturating_mul(8).max(1))
        .with_max_scratch_bytes(source.len().saturating_mul(128).max(1))
    }

    fn field(number: u32, payload: &[u8]) -> Vec<u8> {
        let mut out = Vec::new();
        emit_length_delimited(number, payload, &mut out);
        out
    }

    fn varint_field(number: u32, value: u64) -> Vec<u8> {
        let mut out = Vec::new();
        emit_varint_field(number, value, &mut out);
        out
    }

    fn uuid(lower: u64, upper: u64) -> Vec<u8> {
        canonical_uuid(UuidSnapshot::new(lower, upper)).expect("uuid")
    }

    #[test]
    fn unknown_groups_and_overlong_unknown_scalars_survive_owner_rewrite() {
        let state =
            RowOrColumnStateSnapshot::new(UuidSnapshot::new(10, 20)).with_user_hidden(Some(true));
        let extent =
            HiddenStateExtentSnapshot::new(UuidSnapshot::new(11, 20), AxisDirection::Row, [state])
                .expect("extent");
        let column =
            HiddenStateExtentSnapshot::new(UuidSnapshot::new(12, 20), AxisDirection::Column, [])
                .expect("extent");
        let hidden = HiddenStatesSnapshot::new(UuidSnapshot::new(13, 20), column, extent);
        let owner =
            HiddenStatesOwnerSnapshot::new(UuidSnapshot::new(14, 20), [hidden]).expect("owner");
        let mut source = canonical_owner(&owner).expect("owner");
        source.extend_from_slice(&[0x98, 0x06, 0x80, 0x00]); // unknown field 211, overlong zero
        source.extend_from_slice(&[0x9b, 0x06, 0x98, 0x06, 0x80, 0x00, 0x9c, 0x06]);
        let decoded = decode_hidden_states_owner(&source, options(&source)).expect("decode");
        let output =
            rewrite_hidden_states_owner(&source, &decoded, options(&source)).expect("rewrite");
        assert!(output.bytes().ends_with(&[
            0x98, 0x06, 0x80, 0x00, 0x9b, 0x06, 0x98, 0x06, 0x80, 0x00, 0x9c, 0x06
        ]));
        assert_eq!(
            decode_hidden_states_owner(output.bytes(), options(output.bytes())).expect("readback"),
            decoded
        );
    }

    #[test]
    fn nested_unknown_groups_consume_one_recursion_level_each() {
        // field 10 start-group, field 11 start/end-group, field 10
        // end-group.  Calling `skip_group` with the same depth for field 11
        // would incorrectly accept this chain under a two-level ceiling.
        let source = [0x53, 0x5b, 0x5c, 0x54];
        let options = DecodeOptions::new(64, 64, 64, 1_024, 2, 64);
        let mut budget = Budget::new(options);
        let error = scan_fields(&source, options, &mut budget, 1)
            .expect_err("nested groups must consume the recursion budget");
        assert_eq!(
            error.resource_limit(),
            Some(DecodeLimit::Nesting {
                observed: 3,
                maximum: 2,
            })
        );
    }

    #[test]
    fn duplicate_and_wrong_wire_selected_fields_are_rejected() {
        let uid = uuid(1, 2);
        let mut source = field(OWNER_UID_FIELD, &uid);
        source.extend_from_slice(&field(OWNER_UID_FIELD, &uid));
        let error = decode_hidden_states_owner(&source, options(&source)).expect_err("duplicate");
        assert_eq!(
            error.duplicate_singular_field(),
            Some("TST.HiddenStatesOwnerArchive.owner_uid"),
            "{error:?}"
        );
        let wrong = [0x08, 0x01];
        assert!(decode_hidden_states_owner(&wrong, options(&wrong)).is_err());
    }

    #[test]
    fn unchanged_selected_spans_and_explicit_false_are_preserved() {
        let mut source = field(EXTENT_UID_FIELD, &uuid(7, 8));
        // Unknown fields deliberately surround the explicit false flag. They
        // are not decoded into the semantic snapshot and must retain their
        // original position and bytes across an identity rewrite.
        source.extend_from_slice(&[0x90, 0x03, 0x80, 0x00]);
        source.extend_from_slice(&[0x30, 0x00]);
        source.extend_from_slice(&[0x18, 0x00]);
        source.extend_from_slice(&[0x9a, 0x03, 0x01, 0xaa]);

        let (decoded, _) = decode_hidden_state_extent_with_report(&source, options(&source))
            .expect("extent with explicit false must decode");
        assert_eq!(decoded.needs_to_update_filter_set_for_import(), Some(false));
        let output = rewrite_hidden_state_extent(&source, &decoded, options(&source))
            .expect("identity rewrite must succeed");
        assert_eq!(output.bytes(), source.as_slice());
    }

    #[test]
    fn noncanonical_selected_scalar_is_rejected() {
        let mut source = field(EXTENT_UID_FIELD, &uuid(7, 8));
        source.extend_from_slice(&[0x30, 0x80, 0x00]);
        source.extend_from_slice(&[0x18, 0x00]);
        let error = decode_hidden_state_extent(&source, options(&source))
            .expect_err("overlong selected bool must be rejected");
        assert_eq!(
            error.noncanonical_reason(),
            Some("protobuf varint value"),
            "{error:?}"
        );
    }

    #[test]
    fn changed_uuid_keeps_nested_unknown_wire_bytes() {
        let mut old_uuid = varint_field(UUID_LOWER_FIELD, 7);
        old_uuid.extend_from_slice(&[0x49, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08]); // unknown fixed64
        old_uuid.extend_from_slice(&varint_field(UUID_UPPER_FIELD, 8));
        old_uuid.extend_from_slice(&[0x52, 0x01, 0xaa]); // unknown bytes
        old_uuid.extend_from_slice(&[0x5b, 0x60, 0x01, 0x5c]); // unknown group

        let mut source = field(EXTENT_UID_FIELD, &old_uuid);
        source.extend_from_slice(&[0x30, 0x00, 0x18, 0x00]);
        let decoded = decode_hidden_state_extent(&source, options(&source)).expect("extent");
        let desired =
            HiddenStateExtentSnapshot::new(UuidSnapshot::new(17, 18), AxisDirection::Column, [])
                .expect("desired extent");
        let output = rewrite_hidden_state_extent(&source, &desired, options(&source))
            .expect("changed UUID rewrite");

        assert_ne!(output.bytes(), source.as_slice());
        assert!(
            output
                .bytes()
                .windows(9)
                .any(|window| { window == [0x49, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08] })
        );
        assert!(
            output
                .bytes()
                .windows(3)
                .any(|window| window == [0x52, 0x01, 0xaa])
        );
        assert!(
            output
                .bytes()
                .windows(4)
                .any(|window| window == [0x5b, 0x60, 0x01, 0x5c])
        );
        assert_eq!(
            decode_hidden_state_extent(output.bytes(), options(output.bytes()))
                .expect("rewritten extent"),
            desired
        );
        assert_eq!(decoded.needs_to_update_filter_set_for_import(), Some(false));
    }

    #[test]
    fn filter_set_offsets_are_strictly_projected_and_type_bounded() {
        let mut source = varint_field(FILTER_SET_TYPE_FIELD, 1);
        source.extend_from_slice(&varint_field(FILTER_SET_OFFSETS_FIELD, 7));
        source.extend_from_slice(&varint_field(FILTER_SET_OFFSETS_FIELD, 11));
        source.extend_from_slice(&varint_field(FILTER_SET_ENABLED_FIELD, 0));
        let decoded = decode_filter_set(&source, options(&source)).expect("filter-set");
        assert_eq!(decoded.filter_type(), Some(1));
        assert_eq!(decoded.filter_offsets(), &[7, 11]);
        assert_eq!(decoded.is_enabled(), Some(false));

        let invalid_type = varint_field(FILTER_SET_TYPE_FIELD, 2);
        assert!(decode_filter_set(&invalid_type, options(&invalid_type)).is_err());
    }

    #[test]
    fn packed_filter_set_offsets_match_unpacked_values() {
        let mut packed = Vec::new();
        encode_varint(7, &mut packed);
        encode_varint(300, &mut packed);
        encode_varint(u32::MAX as u64, &mut packed);
        let mut source = field(FILTER_SET_OFFSETS_FIELD, &packed);
        source.extend_from_slice(&varint_field(FILTER_SET_OFFSETS_FIELD, 11));

        let decoded = decode_filter_set(&source, options(&source)).expect("packed filter-set");
        assert_eq!(decoded.filter_offsets(), &[7, 300, u32::MAX, 11]);
    }

    #[test]
    fn packed_filter_set_offsets_require_canonical_elements() {
        let source = field(FILTER_SET_OFFSETS_FIELD, &[0x80, 0x00]);
        let error = decode_filter_set(&source, options(&source))
            .expect_err("overlong packed offset must be rejected");
        assert_eq!(
            error.noncanonical_reason(),
            Some("packed filter offset varint"),
            "{error:?}"
        );
    }

    #[test]
    fn formula_owner_dependencies_preserve_empty_envelope_presence() {
        let mut base = field(FORMULA_OWNER_UID_FIELD, &uuid(1, 2));
        base.extend_from_slice(&varint_field(FORMULA_OWNER_INTERNAL_ID_FIELD, 7));

        for field_number in [4, 5, 6, 7, 8, 9, 10, 13, 14, 15, 16] {
            let mut source = base.clone();
            source.extend_from_slice(&field(field_number, &[]));
            let decoded = decode_formula_owner_dependencies(&source, options(&source))
                .unwrap_or_else(|error| {
                    panic!("empty dependency envelope field {field_number} must decode: {error}")
                });
            assert!(
                decoded.has_dependencies(),
                "field {field_number} presence must be retained"
            );
        }

        let decoded = decode_formula_owner_dependencies(&base, options(&base))
            .expect("identity-only dependency object");
        assert!(!decoded.has_dependencies());
    }

    #[test]
    fn desired_hidden_states_reject_axis_direction_mismatch_before_execution() {
        let source = canonical_owner(
            &HiddenStatesOwnerSnapshot::new(UuidSnapshot::new(14, 20), []).expect("empty owner"),
        )
        .expect("owner");
        let state = RowOrColumnStateSnapshot::new(UuidSnapshot::new(10, 20));
        let wrong_column =
            HiddenStateExtentSnapshot::new(UuidSnapshot::new(11, 20), AxisDirection::Row, [state])
                .expect("extent");
        let row = HiddenStateExtentSnapshot::new(UuidSnapshot::new(12, 20), AxisDirection::Row, [])
            .expect("extent");
        let hidden = HiddenStatesSnapshot::new(UuidSnapshot::new(13, 20), wrong_column, row);
        let desired =
            HiddenStatesOwnerSnapshot::new(UuidSnapshot::new(14, 20), [hidden]).expect("owner");
        let error = prepare_hidden_states_owner_rewrite(&source, &desired, options(&source))
            .expect_err("mismatched axis should fail during preparation");
        assert_eq!(
            error.to_string(),
            "invalid Pages hidden-state graph hidden-state extent direction does not match axis"
        );
    }

    #[test]
    fn int32_rejects_non_sign_extended_ten_byte_values() {
        let source = varint_field(FILTER_SET_TYPE_FIELD, 0xffff_ffff_0000_0000);
        let error = decode_filter_set(&source, options(&source))
            .expect_err("non-sign-extended int32 must be rejected");
        assert_eq!(
            error.noncanonical_reason(),
            Some("int32 scalar is not sign-extended"),
            "{error:?}"
        );

        let valid_negative = varint_field(FILTER_SET_TYPE_FIELD, 0xffff_ffff_8000_0000);
        assert_eq!(decode_int32_checked(0xffff_ffff_8000_0000), Ok(i32::MIN));
        assert!(decode_filter_set(&valid_negative, options(&valid_negative)).is_err());
    }

    #[test]
    fn owner_rewrite_rejects_index_allocations_before_planning() {
        let state = RowOrColumnStateSnapshot::new(UuidSnapshot::new(10, 20));
        let extent =
            HiddenStateExtentSnapshot::new(UuidSnapshot::new(11, 20), AxisDirection::Row, [state])
                .expect("extent");
        let column =
            HiddenStateExtentSnapshot::new(UuidSnapshot::new(12, 20), AxisDirection::Column, [])
                .expect("extent");
        let hidden = HiddenStatesSnapshot::new(UuidSnapshot::new(13, 20), column, extent);
        let owner =
            HiddenStatesOwnerSnapshot::new(UuidSnapshot::new(14, 20), [hidden]).expect("owner");
        let source = canonical_owner(&owner).expect("owner");
        let accounting = account_desired(&RewriteDesired::HiddenStatesOwner(owner.clone()))
            .expect("owner accounting");
        let preflight_allocations = accounting
            .allocations
            .checked_add(accounting.match_allocations)
            .and_then(|value| value.checked_add(8))
            .expect("fixture accounting");
        validate_desired_accounting(
            accounting,
            options(&source).with_max_allocations(preflight_allocations),
        )
        .expect("the inclusive pre-plan allocation maximum must be accepted");
        let preflight_error = validate_desired_accounting(
            accounting,
            options(&source).with_max_allocations(preflight_allocations - 1),
        )
        .expect_err("one below the pre-plan allocation maximum must be rejected");
        assert!(matches!(
            preflight_error.resource_limit(),
            Some(DecodeLimit::Allocations {
                observed,
                maximum,
            }) if observed == preflight_allocations && maximum + 1 == preflight_allocations
        ));

        let error = prepare_hidden_states_owner_rewrite(
            &source,
            &owner,
            options(&source).with_max_allocations(0),
        )
        .expect_err("index storage must be budgeted before planning");

        assert!(matches!(
            error.resource_limit(),
            Some(DecodeLimit::Allocations {
                observed,
                maximum: 0,
            }) if observed > 0
        ));
    }

    #[test]
    fn public_snapshot_debug_is_content_free() {
        let uuid = UuidSnapshot::new(0xfeed_face_dead_beef, 0x0123_4567_89ab_cdef);
        let reference = ReferenceSnapshot::new(NonZeroU64::new(0x7654_3210).expect("nonzero"));
        let row = RowOrColumnStateSnapshot::new(uuid).with_user_hidden(Some(true));
        let extent = HiddenStateExtentSnapshot::new(uuid, AxisDirection::Row, [row])
            .expect("extent")
            .with_filter_set(Some(reference));
        let hidden = HiddenStatesSnapshot::new(uuid, extent.clone(), extent.clone());
        let owner = HiddenStatesOwnerSnapshot::new(uuid, [hidden]).expect("owner");
        let table_info = TableInfoSnapshot::new(reference)
            .with_hidden_states_uuid(Some(uuid))
            .with_view_column_row_uids(Some(reference));
        let table_model = TableModelSnapshot::new(123, 456).with_hidden_states_owner(Some(owner));

        for debug in [
            format!("{uuid:?}"),
            format!("{reference:?}"),
            format!("{row:?}"),
            format!("{extent:?}"),
            format!("{table_info:?}"),
            format!("{table_model:?}"),
        ] {
            assert!(!debug.contains("feed"));
            assert!(!debug.contains("dead"));
            assert!(!debug.contains("7654"));
            assert!(!debug.contains("123"));
        }
    }
}
